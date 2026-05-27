# NRLA Data Dashboard

A self-updating dashboard of private rented sector statistics, hosted on GitHub Pages and refreshed automatically every day via GitHub Actions.

Data is fetched from official UK government sources: ONS, Bank of England, MHCLG, MoJ/HMCTS, and the English Housing Survey. If a source hasn't published new data, the previous value is retained unchanged.

---

## Repository structure

```
nrla-data-dashboard/
├── index.html                        ← The dashboard page (open this in a browser)
├── data/
│   ├── macro.json                    ← Macro-economic indicators
│   ├── housing_tenure.json           ← FRS housing tenure data
│   ├── dwellings.json                ← Dwelling completions (MHCLG + Stats Wales)
│   ├── possession.json               ← Landlord possession statistics (MoJ)
│   └── ehs.json                      ← English Housing Survey data
├── scripts/
│   └── update_data.py                ← Python script that fetches and updates data
└── .github/
    └── workflows/
        └── update_data.yml           ← GitHub Actions workflow (runs daily at 07:00 UTC)
```

---

## Setup: step-by-step

### Step 1 — Create the GitHub repository

1. Go to [github.com](https://github.com) and sign in.
2. Click **New repository** (the green button, or the `+` icon top-right → New repository).
3. Name it something like `nrla-data-dashboard`.
4. Set visibility to **Public** (required for free GitHub Pages — see access options below).
5. Do **not** tick "Add a README" — you'll upload the files directly.
6. Click **Create repository**.

### Step 2 — Upload the files

**Option A: GitHub web interface (easiest)**
1. On your new repo page, click **uploading an existing file**.
2. Drag and drop all files and folders from this package, preserving the folder structure.
3. Click **Commit changes**.

**Option B: Git command line**
```bash
git clone https://github.com/YOUR_USERNAME/nrla-data-dashboard.git
# Copy all files from this package into the cloned folder
git add .
git commit -m "Initial dashboard setup"
git push
```

### Step 3 — Enable GitHub Pages

1. In your repo, go to **Settings** → **Pages** (in the left sidebar).
2. Under **Source**, select **Deploy from a branch**.
3. Choose **main** branch and **/ (root)** folder.
4. Click **Save**.
5. After a minute or two, your dashboard will be live at:
   `https://YOUR_USERNAME.github.io/nrla-data-dashboard/`

### Step 4 — Enable the Actions workflow

GitHub Actions runs the data updater automatically. To check it's enabled:

1. Go to the **Actions** tab in your repo.
2. If prompted, click **I understand my workflows, go ahead and enable them**.
3. To trigger a manual first run: click **Update PRS Data Dashboard** → **Run workflow** → **Run workflow**.

After the first run, the `data/` JSON files will be populated with real values and the dashboard will show live statistics.

---

## Access control

The dashboard is a simple HTML page — there's no login system built in. Here are your options:

### Option 1: Public page (recommended)
All the statistics on the dashboard are publicly available government data, so there's no risk in making the page public. Anyone with the URL can view it. Share the link with your team.

### Option 2: Keep the URL internal
Don't publicise the URL externally. It won't appear in search engines unless someone links to it. This is a reasonable "security by obscurity" approach for internal tools.

### Option 3: Password protection via Netlify (free)
[Netlify](https://netlify.com) offers free hosting with built-in password protection:
1. Connect your GitHub repo to Netlify (free account).
2. Go to **Site settings** → **Access control** → **Password protection**.
3. Set a shared password for all NRLA staff.

### Option 4: GitHub Teams (paid, ~$4/user/month)
GitHub Teams allows private repositories with private GitHub Pages. If NRLA already has a GitHub Teams account, you can make the repo private and Pages will still work.

---

## Updating the data manually

If you want to run the updater locally (e.g. to test it or force a refresh):

```bash
# Install dependencies
pip install requests beautifulsoup4 openpyxl lxml

# Run the updater
python scripts/update_data.py
```

Then commit and push the updated `data/` files.

---

## Viewing locally

Because `index.html` uses `fetch()` to load the JSON files, you need to serve it via a local web server rather than opening it directly as a file:

```bash
# From the repo root folder:
python -m http.server 8000
# Then open: http://localhost:8000
```

---

## Data sources

| Category | Source | Update frequency |
|---|---|---|
| CPI, CPIH, GDP, Earnings, PIPR | [ONS](https://www.ons.gov.uk) | Monthly / Quarterly |
| Bank base rate, Mortgage lending | [Bank of England](https://www.bankofengland.co.uk) | Monthly / Quarterly |
| Dwelling completions (England) | [MHCLG Live Tables](https://www.gov.uk/government/statistical-data-sets/live-tables-on-house-building) | Quarterly |
| Dwelling completions (Wales) | [Stats Wales](https://statswales.gov.wales/Catalogue/Housing/New-House-Building) | Quarterly |
| Landlord possession statistics | [MoJ / HMCTS](https://www.gov.uk/government/collections/mortgage-and-landlord-possession-statistics) | Quarterly |
| Housing tenure, rents, stock conditions | [English Housing Survey](https://www.gov.uk/government/collections/english-housing-survey) | Annual |
| Household tenure, residency, rents | [Family Resources Survey](https://www.gov.uk/government/collections/family-resources-survey--2) | Annual |

---

## Troubleshooting

**The Actions workflow fails** — Check the Actions tab for error details. Common causes are changes to source website layouts (gov.uk restructures its data pages periodically). The `scripts/update_data.py` file will need updating to match the new URL or file structure.

**Data isn't updating for a particular indicator** — The source may not have published new data yet. The dashboard will retain the last known value. Check the source URL directly to confirm.

**The dashboard shows "Awaiting first fetch" for all indicators** — The GitHub Actions workflow hasn't run yet. Trigger it manually from the Actions tab.
