# Deploying the Quote Tool to Azure App Service

This Flask app runs on **Azure App Service (Linux, Python)** with **gunicorn**.
The repository root is **this folder** (the one containing `app.py`).

---

## 0. Important: repo root

Initialise git **inside this folder** so `app.py` sits at the repo root.
On Azure that means `gunicorn app:app` finds the app with no extra config.

```bash
cd "ONLINE QUOTE"
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<you>/<repo>.git
git push -u origin main
```

`.gitignore` already excludes `.env`, `__pycache__`, generated files, etc.
**Never commit `.env`** — it holds the SharePoint client secret.

---

## 1. Create the Azure App Service

1. Azure Portal → **Create a resource → Web App**.
2. Publish: **Code**. Runtime stack: **Python 3.11** (or 3.12). OS: **Linux**.
3. Create.

## 2. Connect GitHub (CI/CD)

1. App Service → **Deployment Center**.
2. Source: **GitHub** → authorise → pick your repo + `main` branch.
3. Save. Azure (Oryx) builds automatically: it runs `pip install -r requirements.txt`.

## 3. Startup command

App Service → **Configuration → General settings → Startup Command**:

```
gunicorn --bind=0.0.0.0:8000 --timeout 600 --workers 2 app:app
```

(Same as `startup.txt`. The long timeout covers SharePoint calls + Excel/PDF generation.)

## 4. Application settings (environment variables)

App Service → **Configuration → Application settings → New application setting**.
Add everything from `.env.example` with your real values:

| Name | Value |
|------|-------|
| `SHAREPOINT_TENANT_ID` | your tenant GUID |
| `SHAREPOINT_CLIENT_ID` | app registration client id |
| `SHAREPOINT_CLIENT_SECRET` | app registration secret |
| `SHAREPOINT_HOST` | `your-tenant.sharepoint.com` |
| `SITE_NAME` | your BMS site name |
| `SITE_NAME1` | your Operations site name |
| `LIST_ID_JMS_JOBS` … `LIST_ID_JMS_CLIENTS` | as in `.env.example` |
| `BLOB_CONNECTION_STRING` | your Storage account connection string |
| `CONTAINER` | your container name |
| `IMS_OPS_URL` | (optional) SharePoint Operations URL for the manual-browse link |
| `SCM_DO_BUILD_DURING_DEPLOYMENT` | `true` |

Click **Save** (restarts the app).

## 5. Data storage — Azure Blob

`quote_data.json` (job cache) and `quotes.json` (saved-quote index) are stored in
**Azure Blob Storage** (container = `CONTAINER`). This is the right model for App
Service — the data survives restarts, re-deploys and scales across instances. The
app filesystem (`wwwroot`) is wiped on every deploy, so nothing important is kept
there.

- Get the connection string: Storage account → **Access keys → Connection string**.
- The container is auto-created on first write if it doesn't exist.
- If `BLOB_CONNECTION_STRING` is **not** set, the app falls back to local files
  under `QUOTE_DATA_DIR` (handy for local dev).

### First run — build the job cache

The job list comes from SharePoint. Build/refresh it either way:

- **From the app:** browse to `/api/refresh-data` once (it pulls from SharePoint
  and writes `quote_data.json` to Blob), or
- **From SSH** (App Service → SSH): `cd /home/site/wwwroot && python app.py`

Both write straight to Blob. Consider a scheduled task if you want it auto-refreshed.

## 6. Azure AD app permissions

The app uses **client-credentials** Graph access (`SHAREPOINT_CLIENT_*`).
The app registration needs **application** Microsoft Graph permissions
(`Sites.ReadWrite.All`, `Files.ReadWrite.All`) with **admin consent granted** —
this is the same registration you already use locally, so no change needed.

---

## Local development

```bash
cp .env.example .env      # fill in real secrets
pip install -r requirements.txt
python app.py --serve     # http://127.0.0.1:5000
```

`python app.py` (without `--serve`) rebuilds `quote_data.json` from SharePoint.
