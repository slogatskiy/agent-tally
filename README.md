# Compass Agent Tally

Automated tracker of **Compass, Inc. (COMP)** self-reported real-estate agent
headcount by state and market — the alt-data KPI Emma used to update by hand.

Each run scrapes the agent count off every Compass location page, appends a dated
column to the Excel workbook, and rebuilds a public HTML widget. On the Mac Mini it
runs monthly and pushes to GitHub, so **GitHub Pages always shows the latest**.

- **Live widget:** https://slogatskiy.github.io/agent-tally/
- **Raw data:** [`data/history.json`](data/history.json) · **Excel:** `COMP_Agent)Tally.xlsx`

## Pieces

| File | Role |
|------|------|
| `tally.py` | Scrape all URLs → append a dated column to the xlsx (with backup). |
| `build_site.py` | Read the xlsx → write `data/history.json` + `docs/index.html` (the widget). |
| `run.sh` | Monthly orchestrator: `tally` → `build_site` → git commit & push. |
| `com.comptally.monthly.plist` | LaunchDaemon that runs `run.sh` on the 1st of each month. |
| `deploy/macmini.md` | Step-by-step Mac Mini setup (owner + dev split). |

The agent count comes straight from each page's HTML
(`"8417 Agents Found in California"`), so there's **no browser, API key, or login**.

## Run locally

```bash
python3 -m venv .venv && ./.venv/bin/pip install -r requirements.txt
./.venv/bin/python tally.py        # update the Excel (or --dry-run to just print)
./.venv/bin/python build_site.py   # rebuild docs/index.html + data/history.json
open docs/index.html               # preview the widget
```

`tally.py` flags dead source URLs so they never turn into silent zeros:
`<retired>` (redirects to a generic page — swap in a fresh URL), `<no-data>`,
`<error>`. See `tally.py --help` for `--date`, `--dry-run`, `--file`.

## Automation

Runs on Konstantin's Mac Mini as a system LaunchDaemon (same convention as
`vc-content-digest`), monthly, pushing to this repo. GitHub Pages serves `/docs`.
Full setup — dedicated `comptally` user, SSH-over-Tailscale, credentials — is in
[`deploy/macmini.md`](deploy/macmini.md).

## Notes

- Excel round-trips through openpyxl, which preserves data, formulas, and
  hyperlinks but drops embedded charts / some conditional formatting (none here).
- Counts are Compass's own figures and drift a few agents day to day — read
  month-over-month, not intraday.
- DC is intentionally excluded from the state total (heavy MD/VA overlap), per the
  original workbook note.
