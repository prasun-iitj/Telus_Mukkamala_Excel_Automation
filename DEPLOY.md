# Deployment Guide

Repo: [prasun-iitj/Telus_Mukkamala_Excel_Automation](https://github.com/prasun-iitj/Telus_Mukkamala_Excel_Automation)

This Flask app can run on any platform that supports Python 3.12 + `gunicorn`, or Docker.

| Platform | Best for | Free tier | Config in repo |
|----------|----------|-----------|----------------|
| [Render](https://render.com) | Easiest GitHub deploy | Yes (sleeps) | `render.yaml` |
| [Railway](https://railway.app) | Simple, fast deploys | Trial credits | `railway.toml` + `Dockerfile` |
| [Fly.io](https://fly.io) | Global edge, Docker | Limited free | `fly.toml` + `Dockerfile` |
| [Google Cloud Run](https://cloud.google.com/run) | Pay-per-use, scales to zero | Generous free tier | `Dockerfile` |
| [Azure App Service](https://azure.microsoft.com/products/app-service) | Enterprise / Telus Azure | No | `Dockerfile` |
| [DigitalOcean App Platform](https://www.digitalocean.com/products/app-platform) | Managed PaaS | No | `Dockerfile` |
| [PythonAnywhere](https://www.pythonanywhere.com) | Pure Python hosting | Beginner free tier | Manual (below) |

**Requirements for all platforms:**
- `SECRET_KEY` env var (random string) in production
- Request timeout ≥ **120 seconds** (PPT generation can take time)
- `DSD_Mukkamala_February 2026_orig.pptx` must be in the deployment artifact

---

## Option A — Render (recommended, already configured)

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/prasun-iitj/Telus_Mukkamala_Excel_Automation)

1. Click the button above (or [Blueprint setup](https://dashboard.render.com/blueprint/new?repo=https://github.com/prasun-iitj/Telus_Mukkamala_Excel_Automation))
2. Connect GitHub → **Apply**
3. Live URL: `https://telus-mukkamala-excel-automation.onrender.com`

---

## Option B — Railway

1. Go to [railway.app](https://railway.app) → **New Project** → **Deploy from GitHub repo**
2. Select `Telus_Mukkamala_Excel_Automation`
3. Railway auto-detects `Dockerfile` / `railway.toml`
4. Add variable: `SECRET_KEY` = any long random string
5. **Deploy** → copy the generated `.railway.app` URL

CLI alternative:
```bash
npm i -g @railway/cli
railway login
railway link
railway up
```

---

## Option C — Fly.io

1. Install [flyctl](https://fly.io/docs/hands-on/install-flyctl/)
2. From the project folder:
```bash
fly auth login
fly launch    # accept defaults; uses fly.toml + Dockerfile
fly secrets set SECRET_KEY=your-random-secret
fly deploy
```
3. URL: `https://telus-mukkamala-excel-automation.fly.dev`

---

## Option D — Google Cloud Run

1. Install [Google Cloud SDK](https://cloud.google.com/sdk)
2. Build and deploy:
```bash
gcloud auth login
gcloud config set project YOUR_PROJECT_ID
gcloud run deploy telus-excel-automation \
  --source . \
  --region us-central1 \
  --allow-unauthenticated \
  --timeout 300 \
  --memory 512Mi \
  --set-env-vars SECRET_KEY=your-random-secret
```
3. Cloud Run uses `Dockerfile` automatically with `--source .`

---

## Option E — Azure App Service (Docker)

Good if your team already uses Azure.

1. Azure Portal → **Create** → **Web App**
2. Publish: **Docker Container**
3. Source: GitHub → `Telus_Mukkamala_Excel_Automation`
4. **Configuration** → Application settings:
   - `SECRET_KEY` = random string
   - `WEBSITES_PORT` = `8080`
5. **Configuration** → General settings → Startup timeout: **180** seconds

Or use Azure Container Registry + the repo `Dockerfile`.

---

## Option F — DigitalOcean App Platform

1. [cloud.digitalocean.com/apps](https://cloud.digitalocean.com/apps) → **Create App**
2. Connect GitHub repo
3. Type: **Dockerfile** (auto-detected)
4. HTTP port: **8080**
5. Add env: `SECRET_KEY`
6. Instance size: Basic (512 MB RAM minimum)

---

## Option G — PythonAnywhere

No Docker — upload code directly.

1. Create account at [pythonanywhere.com](https://www.pythonanywhere.com)
2. **Files** → upload/clone the repo (or `git clone` in Bash console)
3. Create virtualenv: `mkvirtualenv --python=/usr/bin/python3.12 telus-ppt`
4. `pip install -r requirements.txt`
5. **Web** → Add new web app → Manual configuration → Flask
6. WSGI file (`/var/www/YOURUSER_pythonanywhere_com_wsgi.py`):
```python
import sys
sys.path.insert(0, '/home/YOURUSER/Telus_Mukkamala_Excel_Automation')
from web_app import app as application
```
7. Set env var `SECRET_KEY` in the Web tab
8. Reload the web app

> Paid plan recommended — free tier has limited CPU for PPT generation.

---

## Local production test

```powershell
pip install -r requirements.txt
$env:SECRET_KEY="local-test-secret"
$env:PORT="8080"
gunicorn --bind 127.0.0.0:8080 --workers 2 --timeout 120 web_app:app
```

Or with Docker:
```powershell
docker build -t telus-excel-automation .
docker run -p 8080:8080 -e SECRET_KEY=local-test-secret telus-excel-automation
```

Open http://localhost:8080

---

## Troubleshooting

| Issue | Fix |
|-------|-----|
| Build fails | Ensure `DSD_Mukkamala_February 2026_orig.pptx` is committed to Git |
| 502 / timeout | Increase platform timeout to 120–300s |
| Cold start slow | Normal on free tiers (Render, Fly, Cloud Run) |
| Upload fails | Check max request body size on the platform (needs ~50 MB for large Excel) |
