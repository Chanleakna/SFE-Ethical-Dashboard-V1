# Auto-deploy to Apps Script (stop copy-pasting)

Goal: every merge to `main` automatically updates **both** files in your Apps
Script project (`Code.gs` + the `index` HTML), publishes a new web-app version,
and clears the cache — with **zero copy-paste**.

It uses Google's `clasp` CLI from a GitHub Action (`.github/workflows/deploy-appsscript.yml`).
The workflow is **dormant** until you finish the one-time setup below.

---

## One-time setup (~10 minutes, done once)

### 1. Turn on the Apps Script API for your Google account
Open <https://script.google.com/home/usersettings> → toggle **Apps Script API = ON**.

### 2. Get a clasp credentials file (on any computer with Node)
```bash
npm install -g @google/clasp@2.4.2
clasp login
```
A browser opens — sign in with the **same Google account that owns the sheet/script**.
This creates a file `~/.clasprc.json` (Windows: `C:\Users\<you>\.clasprc.json`).
Open it and copy its **entire contents** — that's the `CLASP_CREDENTIALS` secret.

### 3. Collect three IDs from your Apps Script project
- **Script ID** — Apps Script editor → ⚙️ **Project Settings** → *IDs* → **Script ID**.
- **Deployment ID** — **Deploy → Manage deployments** → your active deployment → the **Deployment ID** (starts with `AKfyc…`).
- **Exec URL** — the web-app `…/exec` URL (same page).

### 4. Copy your REAL manifest (important — avoids breaking access)
In the Apps Script editor: ⚙️ Project Settings → tick **"Show `appsscript.json` in editor"**.
Open the `appsscript.json` file it reveals, copy its contents, and **replace the
contents of `appsscript.json` in this repo** with yours. (The one committed here
is a best-guess default; using your real one keeps your web-app access settings
exactly as they are.)

### 5. Add the secrets + enable flag on GitHub
Repo → **Settings → Secrets and variables → Actions**.

Under **Secrets**, add:
| Name | Value |
|---|---|
| `CLASP_CREDENTIALS` | full contents of `~/.clasprc.json` |
| `APPS_SCRIPT_ID` | the Script ID |
| `APPS_SCRIPT_DEPLOYMENT_ID` | the Deployment ID (`AKfyc…`) |
| `APPS_SCRIPT_EXEC_URL` | the `…/exec` URL |

Under **Variables**, add:
| Name | Value |
|---|---|
| `AUTODEPLOY_ENABLED` | `true` |

---

## After setup
- Merge a PR (or push to `main`) → the Action builds, pushes both files, publishes a
  new version, and clears the cache. Refresh the dashboard — done.
- Run it manually anytime: repo **Actions** tab → *Deploy to Apps Script* → **Run workflow**.
- To pause it: set the `AUTODEPLOY_ENABLED` variable to anything other than `true`.

## Notes
- First run does a **fresh** `clasp push`; verify the dashboard still works and the
  web-app access is unchanged (that's why step 4 matters).
- Credentials never leave GitHub Secrets; they're not committed to the repo.
