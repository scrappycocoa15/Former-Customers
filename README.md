# Account Reassignment Tool — Deployment Guide

## What the app does
- Fetches the former customer report directly from Salesforce (report ID hard-coded)
- Routes each account to a General Business or National rep based on D&B Employees Worldwide (>300 = National)
- Looks up the assigned rep by ZIP code from the territory maps you upload
- Updates Marketing Tier to Tier 4 and strips Prev Acct Owner tags from FY18 Sales Planning
- Displays results in two tabs: ready-to-reassign accounts and Ask RSD / unmatched accounts
- Exports a highlighted Excel workbook and a pre-addressed .eml email draft

---

## Files
| File | Purpose |
|---|---|
| `app.py` | Streamlit application |
| `requirements.txt` | Python dependencies |

---

## Run locally

```bash
# 1. Create and activate a virtual environment (recommended)
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate

# 2. Install dependencies
pip install -r requirements.txt

# 3. Launch the app
streamlit run app.py
```

The app opens at http://localhost:8501

---

## Deploy to Streamlit Community Cloud (free)

1. Push `app.py` and `requirements.txt` to a GitHub repository (can be private).
2. Go to https://share.streamlit.io and sign in with your GitHub account.
3. Click **New app** → select your repo, branch, and set **Main file path** to `app.py`.
4. Click **Deploy**. The app will be live at a `*.streamlit.app` URL in ~2 minutes.

No secrets or environment variables are required — the Salesforce session ID is entered at runtime.

---

## Getting your Salesforce Session ID

**Option A — Browser dev tools**
1. Log in to Salesforce.
2. Open dev tools (F12) → Network tab.
3. Reload the page or click anything that makes an API call.
4. Find any request to `*.salesforce.com` → Headers → copy the value after `Authorization: Bearer `.

**Option B — Salesforce Inspector Chrome extension**
1. Install "Salesforce Inspector Reloaded" from the Chrome Web Store.
2. Click the extension icon while logged in to Salesforce.
3. Copy the Session ID shown on the home screen.

Session IDs expire after your org's configured timeout (typically 2 hours for interactive sessions).

---

## Updating territory maps
Upload fresh GB and National territory `.xlsx` files each month via the sidebar.
The app auto-detects the correct sheet (any sheet name containing "zip") and the
owner name/ID columns (any column containing "owner" and "id").

---

## Changing the report or instance
Edit the constants near the top of `app.py`:

```python
INSTANCE_URL = "https://sapconcur.my.salesforce.com"
REPORT_ID    = "00O7V000006IT6Z"
```
