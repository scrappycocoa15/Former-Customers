"""
Former Customer Account Reassignment Tool
Streamlit app — fetches a Salesforce report, routes accounts to GB or National
territory reps by ZIP code, and produces a highlighted Excel output + .eml draft.
"""

import io
import re
import json
import requests
import pandas as pd
import streamlit as st
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formatdate

# ── Constants ─────────────────────────────────────────────────────────────────

INSTANCE_URL = "https://sapconcur.my.salesforce.com"
REPORT_ID    = "00O7V000006IT6Z"
API_VERSION  = "v59.0"
EMAIL_TO     = "concur_FieldServices@sap.com"

HIGHLIGHT_COLS = {"New Account Owner ID", "New Account Owner Name",
                  "Marketing Tier", "FY18 Sales Planning"}

SAP_BLUE      = "0070F2"
SAP_LIGHT     = "E1F4FF"
SAP_HEADER_BG = "00144A"

# ── Persistent territory map storage ─────────────────────────────────────────
# Files are saved alongside app.py in a data/ subfolder.
# On every run the app loads from disk — no re-upload needed.

DATA_DIR  = Path(__file__).parent / "data"
DATA_DIR.mkdir(exist_ok=True)
GB_PATH   = DATA_DIR / "gb_territory.xlsx"
NAT_PATH  = DATA_DIR / "nat_territory.xlsx"
META_PATH = DATA_DIR / "territory_meta.json"


def _load_meta() -> dict:
    """Return the saved territory metadata (filenames, save timestamps)."""
    if META_PATH.exists():
        return json.loads(META_PATH.read_text())
    return {}


def _save_meta(meta: dict):
    META_PATH.write_text(json.dumps(meta, indent=2))


def _save_territory_file(uploaded_file, dest: Path):
    dest.write_bytes(uploaded_file.getbuffer())

# ── Salesforce helpers ────────────────────────────────────────────────────────

def fetch_report(session_id: str) -> dict:
    """Call the Salesforce Analytics API and return the raw JSON."""
    url = (f"{INSTANCE_URL}/services/data/{API_VERSION}"
           f"/analytics/reports/{REPORT_ID}?includeDetails=true")
    resp = requests.get(url,
                        headers={"Authorization": f"Bearer {session_id}"},
                        timeout=30)
    if resp.status_code in (401, 403):
        raise ValueError(
            "Session ID is invalid or expired. "
            "Please paste a fresh token and try again."
        )
    resp.raise_for_status()
    return resp.json()


def parse_report(data: dict) -> pd.DataFrame:
    """
    Convert the Salesforce report factMap into a flat DataFrame.
    Uses the display label for each cell, which matches the XLS export.
    """
    col_keys   = data["reportMetadata"]["detailColumns"]
    col_info   = data["reportExtendedMetadata"]["detailColumnInfo"]
    col_labels = [col_info[k]["label"] for k in col_keys]

    rows = []
    for row in data["factMap"].get("T!T", {}).get("rows", []):
        cells = []
        for cell in row["dataCells"]:
            val   = cell.get("value")
            label = cell.get("label", "")
            if isinstance(val, dict):
                # Linked / lookup field — label is the display name
                cells.append(label or val.get("name", "") or None)
            elif val is None or str(val).strip() in ("", "-"):
                cells.append(label if label not in ("", "-") else None)
            else:
                cells.append(str(val).strip() or label or None)
        rows.append(cells)

    df = pd.DataFrame(rows, columns=col_labels)
    df.columns = [c.strip() for c in df.columns]   # strip any stray whitespace
    return df


# ── Territory helpers ─────────────────────────────────────────────────────────

def _find_zip_sheet(xl: pd.ExcelFile) -> str:
    """Return the sheet name that contains zip assignments."""
    for name in xl.sheet_names:
        if "zip" in name.lower():
            return name
    return xl.sheet_names[0]


def _find_owner_columns(df: pd.DataFrame):
    """
    Detect owner-ID and owner-name columns regardless of fiscal year prefix
    (e.g. 'FY26 Account Owner ID' → id_col, 'FY26 Account Owner' → name_col).
    """
    id_col = name_col = None
    for col in df.columns:
        cl = col.lower()
        if "owner" in cl and "id" in cl:
            id_col = col
        elif "owner" in cl and "id" not in cl:
            name_col = col
    if not id_col or not name_col:
        raise ValueError(
            f"Cannot detect owner columns in territory file. "
            f"Found: {list(df.columns)}"
        )
    return id_col, name_col


def build_lookup(file_obj) -> dict:
    """
    Read a territory Excel/CSV (uploaded file-like object or Path) and build
    a  zip5 → {'id': ..., 'name': ...}  dict.
    """
    # Determine engine from filename so Path objects work as well as uploads
    name = str(getattr(file_obj, "name", file_obj)).lower()
    if name.endswith(".csv"):
        df_all = pd.read_csv(file_obj)
        zip_sheet_df = df_all
        sheet_data = df_all
        xl = None
    else:
        engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"
        xl     = pd.ExcelFile(file_obj, engine=engine)
        sheet  = _find_zip_sheet(xl)
        sheet_data = xl.parse(sheet)
    id_col, name_col = _find_owner_columns(sheet_data)

    zips = sheet_data["Zip Code"].astype(str).str.strip()
    mask = zips.str.match(r"^\d+(\.\d+)?$")      # skip non-numeric (data noise)
    clean = sheet_data[mask].copy()
    clean["_z5"] = (clean["Zip Code"]
                    .astype(float).astype(int)
                    .astype(str).str.zfill(5))

    return {
        z: {"id": str(i).strip(), "name": str(n).strip()}
        for z, i, n in zip(clean["_z5"], clean[id_col], clean[name_col])
    }


# ── Processing helpers ────────────────────────────────────────────────────────

def _to_float(val):
    try:
        return float(str(val).replace(",", "").strip())
    except Exception:
        return None


def _normalize_zip(val) -> str:
    s = str(val).strip() if pd.notna(val) and str(val) not in ("None", "nan", "") else ""
    return s[:5] if len(s) >= 5 else s


def _clean_fy18(val):
    """Remove all 'Prev Acct Owner: <name> [YY]' fragments from a string."""
    if pd.isna(val) or str(val).strip() in ("", "nan", "None"):
        return val
    cleaned = re.sub(
        r"\s*Prev Acct Owner:\s+.+?(?:\s+\d{1,2})?\s*(?=Prev Acct Owner:|$)",
        "",
        str(val),
    ).strip()
    return cleaned if cleaned else None


def process_accounts(df: pd.DataFrame,
                     gb_lookup: dict,
                     nat_lookup: dict) -> pd.DataFrame:
    """Route each account to a rep and apply field transformations."""
    new_ids, new_names, teams = [], [], []

    for _, row in df.iterrows():
        emp    = _to_float(row.get("D&B Employees Worldwide"))
        is_nat = emp is not None and emp > 300
        team   = "National" if is_nat else "General Business"
        lookup = nat_lookup if is_nat else gb_lookup
        zip5   = _normalize_zip(row.get("Billing Zip/Postal", ""))
        match  = lookup.get(zip5)

        new_ids.append(match["id"]   if match else None)
        new_names.append(match["name"] if match else None)
        teams.append(team)

    out = df.copy()
    out["Marketing Tier"]         = "Tier 4"
    out["FY18 Sales Planning"]    = out["FY18 Sales Planning"].apply(_clean_fy18)
    out["New Account Owner ID"]   = new_ids
    out["New Account Owner Name"] = new_names
    out["_Team"]                  = teams
    return out


def split_results(df: pd.DataFrame):
    """
    Tab 1 — accounts with a real rep assigned.
    Tab 2 — unmatched, Ask RSD, or Open Territory accounts.
    """
    problem_mask = (
        df["New Account Owner ID"].isna()
        | df["New Account Owner Name"].isna()
        | df["New Account Owner ID"].astype(str).str.lower().str.contains("ask rsd",       na=False)
        | df["New Account Owner Name"].astype(str).str.lower().str.contains("ask rsd",       na=False)
        | df["New Account Owner Name"].astype(str).str.lower().str.contains("open territory", na=False)
    )
    return df[~problem_mask].copy(), df[problem_mask].copy()


# ── Excel builder ─────────────────────────────────────────────────────────────

def build_excel(ready: pd.DataFrame, rsd: pd.DataFrame) -> bytes:
    fill_highlight = PatternFill(start_color=SAP_LIGHT,     end_color=SAP_LIGHT,     fill_type="solid")
    fill_header    = PatternFill(start_color=SAP_HEADER_BG, end_color=SAP_HEADER_BG, fill_type="solid")
    font_header    = Font(color="FFFFFF", bold=True, name="Calibri")

    display_cols = [c for c in ready.columns if not c.startswith("_")]

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        ready[display_cols].to_excel(writer, sheet_name="Account Reassignments", index=False)
        rsd[display_cols].to_excel(writer,   sheet_name="Ask RSD & Unmatched",   index=False)
    buf.seek(0)
    wb = load_workbook(buf)

    for ws in wb.worksheets:
        # Header row
        for cell in ws[1]:
            cell.fill      = fill_header
            cell.font      = font_header
            cell.alignment = Alignment(horizontal="left", vertical="center")

        # Identify columns to highlight
        highlight_indices = [
            cell.column for cell in ws[1]
            if cell.value in HIGHLIGHT_COLS
        ]

        # Apply highlight to data rows
        for col_idx in highlight_indices:
            for row_idx in range(2, ws.max_row + 1):
                ws.cell(row=row_idx, column=col_idx).fill = fill_highlight

        # Auto-width
        for col in ws.columns:
            letter  = get_column_letter(col[0].column)
            max_len = max((len(str(c.value or "")) for c in col), default=8)
            ws.column_dimensions[letter].width = min(max_len + 3, 45)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ── EML builder ───────────────────────────────────────────────────────────────

def build_eml(excel_bytes: bytes, filename: str,
              n_ready: int, n_rsd: int) -> bytes:
    month_label = datetime.today().strftime("%B %Y")

    msg              = MIMEMultipart()
    msg["To"]        = EMAIL_TO
    msg["Subject"]   = f"Account Reassignment Updates - {month_label}"
    msg["Date"]      = formatdate(localtime=True)

    body = (
        f"Hello,\n\n"
        f"Please make the following updates for {n_ready} former customer "
        f"accounts (see Tab: Account Reassignments in the attached file):\n\n"
        f"  1. Update Account Owner to the New Account Owner listed\n"
        f"  2. Update the FY18 Sales Planning field "
        f"(Prev Acct Owner remarks have been removed)\n"
        f"  3. Update Marketing Tier to Tier 4\n\n"
        f"Fields highlighted in blue indicate the values to be applied.\n\n"
        f"An additional {n_rsd} account(s) in the 'Ask RSD & Unmatched' tab "
        f"require manual territory review before reassignment.\n\n"
        f"Please confirm once complete.\n\n"
        f"Thank you"
    )
    msg.attach(MIMEText(body, "plain"))

    part = MIMEBase(
        "application",
        "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    part.set_payload(excel_bytes)
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", "attachment", filename=filename)
    msg.attach(part)

    return msg.as_bytes()


# ── Page config & CSS ─────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Account Reassignment",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(f"""
<style>
  /* Sidebar */
  [data-testid="stSidebar"] {{
      background-color: #F5F6F8;
  }}
  /* Metric cards */
  .metric-card {{
      background: #E1F4FF;
      border-radius: 8px;
      padding: 18px 12px;
      text-align: center;
      border: 1px solid #C0DCF8;
  }}
  .metric-num {{
      font-size: 2rem;
      font-weight: 700;
      color: #{SAP_BLUE};
      line-height: 1.1;
  }}
  .metric-lbl {{
      font-size: 0.78rem;
      color: #555;
      margin-top: 4px;
  }}
  /* Primary button */
  div[data-testid="stButton"] > button[kind="primary"] {{
      background-color: #{SAP_BLUE};
      color: white;
      border: none;
      font-weight: 600;
  }}
  div[data-testid="stButton"] > button[kind="primary"]:hover {{
      background-color: #005AC2;
  }}
  /* Download buttons */
  div[data-testid="stDownloadButton"] > button {{
      background-color: #{SAP_BLUE};
      color: white;
      border: none;
      font-weight: 600;
      width: 100%;
  }}
  div[data-testid="stDownloadButton"] > button:hover {{
      background-color: #005AC2;
  }}
</style>
""", unsafe_allow_html=True)


# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.image(
        "https://upload.wikimedia.org/wikipedia/commons/5/59/SAP_2011_logo.svg",
        width=60,
    )
    st.markdown("## Account Reassignment")
    st.markdown("---")

    st.markdown("#### Salesforce Session ID")
    st.caption(
        "Open Salesforce → browser dev tools (F12) → Network tab → "
        "find any XHR request → copy the `Authorization: Bearer ...` value. "
        "Or use the **Salesforce Inspector** Chrome extension."
    )
    session_id = st.text_input(
        "Session ID",
        type="password",
        placeholder="00D…",
        label_visibility="collapsed",
    )

    st.markdown("---")
    st.markdown("#### Territory Maps")

    # ── Territory map status ──────────────────────────────────────────────────
    meta         = _load_meta()
    gb_on_disk   = GB_PATH.exists()
    nat_on_disk  = NAT_PATH.exists()
    maps_ready   = gb_on_disk and nat_on_disk

    if maps_ready:
        gb_label  = meta.get("gb_name",    "gb_territory.xlsx")
        nat_label = meta.get("nat_name",   "nat_territory.xlsx")
        saved_at  = meta.get("saved_at",   "unknown date")
        st.success("Territory maps loaded from disk")
        st.caption(f"**GB:** {gb_label}")
        st.caption(f"**National:** {nat_label}")
        st.caption(f"Last updated: {saved_at}")
    else:
        st.warning("No territory maps saved yet. Upload below.", icon="⚠️")

    # ── Update expander — collapsed by default once maps exist ────────────────
    with st.expander(
        "Update territory maps" if maps_ready else "Upload territory maps (required)",
        expanded=not maps_ready,
    ):
        new_gb  = st.file_uploader(
            "General Business territory (.xlsx)",
            type=["xlsx", "xls", "csv"],
            key="gb_upload",
        )
        new_nat = st.file_uploader(
            "National territory (.xlsx)",
            type=["xlsx", "xls", "csv"],
            key="nat_upload",
        )
        save_btn = st.button(
            "Save territory maps",
            disabled=not (new_gb and new_nat),
            use_container_width=True,
        )
        if save_btn and new_gb and new_nat:
            _save_territory_file(new_gb,  GB_PATH)
            _save_territory_file(new_nat, NAT_PATH)
            _save_meta({
                "gb_name":  new_gb.name,
                "nat_name": new_nat.name,
                "saved_at": datetime.today().strftime("%Y-%m-%d %H:%M"),
            })
            st.success("Territory maps saved. They will be used on all future runs.")
            st.rerun()

    st.markdown("---")
    run_btn = st.button("Fetch & Process Accounts", type="primary", use_container_width=True)


# ── Main ──────────────────────────────────────────────────────────────────────

st.markdown("# Former Customer Account Reassignment")
st.caption(f"Report: `{REPORT_ID}` · Instance: `{INSTANCE_URL}`")

if not run_btn:
    st.info(
        "Enter your Salesforce Session ID, upload both territory maps, "
        "then click **Fetch & Process Accounts**.",
        icon="ℹ️",
    )
    st.stop()

# ── Validate inputs ───────────────────────────────────────────────────────────

errors = []
if not session_id.strip():
    errors.append("A Salesforce Session ID is required.")
if not maps_ready:
    errors.append("Territory maps have not been saved yet. Upload them in the sidebar first.")
if errors:
    for e in errors:
        st.error(e)
    st.stop()

# ── Load territory maps from disk ─────────────────────────────────────────────

with st.spinner("Loading territory maps…"):
    try:
        gb_lookup  = build_lookup(GB_PATH)
        nat_lookup = build_lookup(NAT_PATH)
        st.success(
            f"Territory maps ready — "
            f"GB: {len(gb_lookup):,} ZIPs | National: {len(nat_lookup):,} ZIPs"
        )
    except Exception as exc:
        st.error(f"Territory map error: {exc}")
        st.stop()

# ── Fetch Salesforce report ───────────────────────────────────────────────────

with st.spinner("Fetching report from Salesforce…"):
    try:
        raw_data = fetch_report(session_id.strip())
        df_raw   = parse_report(raw_data)

        if not raw_data.get("allData", True):
            st.warning(
                "The report contains more than 2,000 rows. "
                "Salesforce returned the first 2,000 only. "
                "Consider splitting the report or using a filtered view.",
                icon="⚠️",
            )

        st.success(f"Report fetched — {len(df_raw):,} accounts loaded")
    except ValueError as exc:
        st.error(str(exc))
        st.stop()
    except requests.HTTPError as exc:
        st.error(f"Salesforce returned an error: {exc}")
        st.stop()
    except Exception as exc:
        st.error(f"Unexpected error fetching report: {exc}")
        st.stop()

# ── Optional: show raw columns for debugging ──────────────────────────────────

with st.expander("Raw columns returned by Salesforce report (for diagnostics)", expanded=False):
    st.write(list(df_raw.columns))

# ── Process ───────────────────────────────────────────────────────────────────

with st.spinner("Processing accounts…"):
    processed       = process_accounts(df_raw, gb_lookup, nat_lookup)
    ready_df, rsd_df = split_results(processed)

# ── Metrics ───────────────────────────────────────────────────────────────────

n_total    = len(processed)
n_ready    = len(ready_df)
n_national = int((processed["_Team"] == "National").sum())
n_gb       = int((processed["_Team"] == "General Business").sum())
n_rsd      = len(rsd_df)

c1, c2, c3, c4, c5 = st.columns(5)
for col, num, lbl in [
    (c1, n_total,    "Total Accounts"),
    (c2, n_ready,    "Ready to Reassign"),
    (c3, n_national, "Routed → National"),
    (c4, n_gb,       "Routed → GB"),
    (c5, n_rsd,      "Ask RSD / Unmatched"),
]:
    col.markdown(
        f'<div class="metric-card">'
        f'<div class="metric-num">{num}</div>'
        f'<div class="metric-lbl">{lbl}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)

# ── Results tabs ──────────────────────────────────────────────────────────────

display_cols = [c for c in processed.columns if not c.startswith("_")]

tab1, tab2 = st.tabs([
    f"Account Reassignments  ({n_ready})",
    f"Ask RSD & Unmatched  ({n_rsd})",
])

with tab1:
    st.caption("Accounts with a verified territory rep assigned, ready to be updated in Salesforce.")
    st.dataframe(ready_df[display_cols], use_container_width=True, hide_index=True)

with tab2:
    st.caption(
        "Accounts where no territory match was found, or the territory is marked "
        "'Ask RSD' / 'Open Territory'. These require manual review before reassignment."
    )
    st.dataframe(rsd_df[display_cols], use_container_width=True, hide_index=True)

# ── Outputs ───────────────────────────────────────────────────────────────────

st.markdown("---")
st.markdown("#### Export")

filename     = f"Account_Reassignments_{datetime.today().strftime('%Y_%m')}.xlsx"
excel_bytes  = build_excel(ready_df, rsd_df)
eml_bytes    = build_eml(excel_bytes, filename, n_ready, n_rsd)
eml_filename = f"Reassignment_Email_{datetime.today().strftime('%Y_%m')}.eml"

col_dl, col_em = st.columns(2)

with col_dl:
    st.download_button(
        label="Download Excel",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption(
        "Two-tab workbook: Account Reassignments + Ask RSD & Unmatched. "
        "Changed fields highlighted in blue."
    )

with col_em:
    st.download_button(
        label="Generate Email Draft (.eml)",
        data=eml_bytes,
        file_name=eml_filename,
        mime="message/rfc822",
        use_container_width=True,
    )
    st.caption(
        f"Opens in Outlook pre-addressed to {EMAIL_TO} "
        "with the Excel file attached. Review and send."
    )
