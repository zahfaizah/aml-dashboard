# AML Dashboard

Live: https://zahfaizah.github.io/aml-dashboard/

Three teams feed one dashboard. **Each PIC only updates their own team.** You do
not need the other teams' Excel files.

| Team     | Your Excel                         | Double-click          | Updates              |
|----------|------------------------------------|-----------------------|----------------------|
| Sanction | `SS_Dashboard_Data_2026.xlsx`      | `update_sanction.bat` | `data/sanction.json` |
| NON TMS  | `NON_TMS_Dashboard_Data_2026.xlsx` | `update_nontms.bat`   | `data/nontms.json`   |
| TMS      | `TMS_Dashboard_Data_2026.xlsx`     | `update_tms.bat`      | `data/tms.json`      |

## Monthly update

1. Update your Excel as usual.
2. Double-click your team's `update_<team>.bat`.
3. It shows each team's months, e.g.
   ```
   Sanction  Jan 2026 - Sep 2026 from your Excel
   NON TMS   Jan 2026 - Apr 2026 as published 2026-10-01 14:15 by ...
   TMS       Jan 2026 - Aug 2026 as published 2026-10-01 14:15 by ...
   ```
   Open `index.html` in this folder if you want to check it, then press **Enter** to publish.
4. The live page updates in about a minute.

What the bat does for you:

- **Downloads** everyone's latest published data from GitHub first.
- **Reads only your Excel** and saves it to your team's `data/<team>.json`.
- **Rebuilds** `index.html` from all three `data/*.json`, so the other teams stay exactly as they published them.
- **Publishes** it. If someone else publishes at the same moment, it fetches their update, rebuilds and tries again.

It warns you, and does nothing unless you type `y`, when your Excel has **fewer
months than what's already published** for your team (e.g. you have an old copy).

## One-time setup (each PIC)

1. **Python 3**: https://www.python.org/downloads/ (tick *Add python.exe to PATH*), then in a terminal:
   ```
   pip install -r requirements.txt
   ```
2. **Git**: https://git-scm.com/download/win (default options), then:
   ```
   git config --global user.name  "Your Name"
   git config --global user.email "you@dana.id"
   ```
3. **Access**: ask the repo owner to add your GitHub account as a collaborator
   (repo Settings → Collaborators), and accept the email invite.
4. **Get the folder**:
   ```
   cd %USERPROFILE%\Documents
   git clone https://github.com/zahfaizah/aml-dashboard.git
   ```
   The first publish opens a GitHub sign-in window. Sign in once.
5. **Tell it where your Excel is**: either copy your Excel into the
   `aml-dashboard` folder, or create `excel_paths.local.json` in that folder, for example:
   ```json
   { "tms": "C:/Users/you/OneDrive - DANA INDONESIA/AML/TMS_Dashboard_Data_2026.xlsx" }
   ```
   (keys: `sanction`, `nontms`, `tms`; use forward slashes). Neither the Excel nor this
   file is ever uploaded (see `.gitignore`).

## Rules

- Don't edit `index.html` or another team's `data/*.json` by hand. They're generated.
- Design changes go in `AML_Dashboard_template.html`. After editing it, run
  `python generate_dashboard.py --rebuild`, check `index.html`, then commit and push
  both files.
- `python generate_dashboard.py --team <team>` (without `--publish`) only makes
  `preview.html` locally. Nothing is saved or published.
