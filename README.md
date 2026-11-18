# KPN Fresh · Cash Till Manned Dashboard

Live page: `https://<your-github-username>.github.io/Cash_till_manning/`

## Daily update (about 1 minute)
1. Download the **Cash Till - Store wise D-1** tab (File → Download → .xlsx, or save the daily Excel) and put it in the `data/` folder,
   e.g. `data/Cash Till - 18th Nov.xlsx`. If you download the whole workbook, the script picks the "Cash Till - Store wise D-1" tab by itself.
2. GitHub Desktop → Commit → Push.
3. GitHub rebuilds `data.json` (Actions tab, ~30 s). Open dashboards update themselves within a minute and show an "updated" toast.

Every file in `data/` is merged by date, so keep the old files there (each day's file is ~35 KB).
If a date is uploaded twice, the newest commit wins — re-pushing a corrected file is safe.

## Data checks (GitHub emails you)
The workflow fails — and GitHub emails the repo owner — when the newest file has a problem. Details are in the run's summary (Actions tab).
- **Critical** (dashboard NOT updated): layout changed, header/half-hour columns missing, no date, no store rows.
- **Warnings** (dashboard updated, please check): file not dated yesterday, date already loaded, active stores missing,
  new outlet IDs, store at 0 tills all day, planned missing, counts far above plan, COM missing from `config/tm_master.csv`,
  overall manned swings > 10 pp vs the 7-day average.

## TM mapping
Edit `config/tm_master.csv` (Cluster Name,TM Name) and push. COMs not listed show under TM "Unmapped".

## How the % is calculated
Manned % = Σ min(tills manned, tills planned) ÷ Σ tills planned over half-hour slots.
Overall 06:30–22:00 (31 slots) · Morning peak 08:00–12:00 · Evening peak 17:00–21:00 · Off-peak 12:00–17:00.
Store-days with zero tills in every slot (no till data) are excluded.

## One-time setup
- Repo → Settings → Pages → Deploy from a branch → `main` / root.
- Repo → Settings → Actions → General → Workflow permissions → Read and write permissions.
- Make sure GitHub notifications for failed workflow runs are on (Settings → Notifications → Actions).

## Run locally
`pip install openpyxl` · `python build_data.py` then `python -m http.server 8000` and open http://localhost:8000
