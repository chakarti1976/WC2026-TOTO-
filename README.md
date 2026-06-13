# 🏆 WC 2026 Toto — Live Leaderboard

A fully automated website for tracking the FIFA World Cup 2026 Toto competition.

**Live site:** `https://YOUR_USERNAME.github.io/wc2026-toto/`

---

## Features

- 🏆 **Leaderboard** — Real-time rankings with prize amounts
- ⚽ **Match Results** — All 72 group stage matches with scores
- 📊 **Group Standings** — All 12 groups (A–L)
- 🎯 **Toto Teams** — Every participant's team selections with multipliers
- 📈 **Selection Overview** — Team popularity stats & charts
- 📋 **Rules** — Full rules in English

## Automation

The site updates automatically:
- **Every 2 hours** via GitHub Actions scheduled jobs
- **Instantly** when you push a new version of the Excel file to the repo
- **Manually** via the "Run workflow" button in GitHub Actions

## Setup Instructions

### 1. Create GitHub repository

```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_USERNAME/wc2026-toto.git
git push -u origin main
```

### 2. Enable GitHub Pages

1. Go to your repo → **Settings** → **Pages**
2. Under **Source**, select **GitHub Actions**
3. Save

### 3. Update scores

**Option A — Push the Excel file** (recommended):
1. Update `WC_2026_Toto__Participants_list.xlsx` with new scores
2. `git add WC_2026_Toto__Participants_list.xlsx`
3. `git commit -m "Update scores — Match 5"`
4. `git push`
5. GitHub Actions automatically regenerates `data.js` and redeploys 🚀

**Option B — Run locally then push**:
```bash
python extract_data.py   # regenerates data.js
git add data.js
git commit -m "Manual data update"
git push
```

**Option C — Manual trigger**:
Go to **Actions** → **Update WC 2026 Toto Data** → **Run workflow**

## Files

| File | Purpose |
|------|---------|
| `index.html` | The entire website (single file) |
| `data.js` | Auto-generated data from Excel |
| `extract_data.py` | Python script that reads Excel → data.js |
| `.github/workflows/update-and-deploy.yml` | GitHub Actions automation |
| `WC_2026_Toto__Participants_list.xlsx` | The master data file |

## Local Development

```bash
# Regenerate data after editing Excel
python extract_data.py

# Serve locally
python -m http.server 8080
# Then open http://localhost:8080
```

---

*Maintained by Yagub Alizada & Arūnas Čižauskas*
