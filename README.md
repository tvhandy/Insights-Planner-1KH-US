# Insights Planner | 1KH US

Apps Script code and assets for the 1000heads Insights US planner: a Google Sheet plus a web app dashboard, synced from Basecamp.

## What's here

| Path | Paste into Apps Script as |
|---|---|
| `src/Code.gs` | Code.gs (Basecamp sync, digests, team posts, tracker logging) |
| `src/Dashboard.gs` | Dashboard.gs (server functions for the dashboard) |
| `src/DashboardPage.html` | DashboardPage (the dashboard itself) |
| `assets/planner-icon-dark.png`, `assets/planner-icon-light.png` | Icons (browser tab, Basecamp link) |

## Releasing a change

1. Bump `PLANNER_VERSION` at the top of `src/Code.gs`.
2. Paste the three files into Apps Script (select all, delete, paste).
3. Deploy > Manage deployments > edit > Version: New version. Never "New deployment"; that changes the link.
4. Commit here with a short description of what changed.

To roll back, paste the files from an earlier commit and redeploy.

Do not commit secrets. Basecamp client ID and secret live in Apps Script script properties, not in this code.
