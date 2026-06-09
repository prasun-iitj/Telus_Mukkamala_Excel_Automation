# Deployment Guide

Repo: [prasun-iitj/Telus_Mukkamala_Excel_Automation](https://github.com/prasun-iitj/Telus_Mukkamala_Excel_Automation)

This Flask app can run on any platform that supports Python 3.12 + `gunicorn`, or Docker.

| Platform | Best for | Free tier | Config in repo |
|----------|----------|-----------|----------------|
| [Render](https://render.com) | Easiest GitHub deploy | Yes (sleeps) | `render.yaml` |
| [Railway](https://railway.app) | Fast GitHub deploy, always-on | Trial credits, then ~$5+/mo | `railway.toml` + `Dockerfile` |
| [Oracle Cloud](https://www.oracle.com/cloud/free/) | Always-on, no sleep | **Always Free** VM forever | `Dockerfile` + `docker-compose.yml` |
| [Fly.io](https://fly.io) | Global edge, Docker | Trial only (new accounts) | `fly.toml` + `Dockerfile` |
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

[Railway](https://railway.app) is the closest Render alternative for **always-on** hosting: no 15-minute sleep, fast GitHub deploys, and a simple dashboard. New accounts get trial credits (~$5); after that you pay only for usage (typically **$5–15/month** for a small internal tool).

**Repo config:** `Dockerfile`, `railway.toml`

### Prerequisites

- GitHub repo: [Telus_Mukkamala_Excel_Automation](https://github.com/prasun-iitj/Telus_Mukkamala_Excel_Automation)
- Railway account ([railway.app](https://railway.app)) — sign in with GitHub

### Deploy via dashboard (recommended)

1. Go to [railway.app/new](https://railway.app/new)
2. Click **GitHub Repo** → authorize Railway → select **`Telus_Mukkamala_Excel_Automation`**
3. Railway detects `Dockerfile` and `railway.toml` automatically
4. Open the new **service** → **Variables** tab → add:

   | Variable | Value |
   |----------|-------|
   | `SECRET_KEY` | Long random string (e.g. from a password generator) |

   Railway sets `PORT` automatically — do not override it.

5. **Settings** → **Networking** → **Generate Domain** → copy your URL  
   Example: `https://telus-mukkamala-excel-automation-production.up.railway.app`
6. Wait for the build to finish (Build Logs tab). Open the URL and test an Excel upload.

### Deploy via CLI

```bash
# Install CLI (requires Node.js)
npm i -g @railway/cli

# Login and link project
railway login
cd Telus_Mukkamala_Excel_Automation
railway init          # create new project or link existing
railway variables set SECRET_KEY=your-long-random-secret
railway up            # deploy from local folder
railway domain        # generate public URL
```

### Auto-deploy on git push

Railway enables this by default when you deploy from GitHub. Each push to `main` triggers a new build. Disable under **Service → Settings → Source** if you prefer manual deploys.

### Railway vs Render (for this app)

| | Railway | Render (free) |
|---|---------|---------------|
| Cost after trial | ~$5–15/mo | $0 |
| Sleeps when idle | **No** | Yes (~15 min) |
| Cold start | Minimal | 30–60 seconds |
| Setup | GitHub connect | Blueprint / button |
| Credit card | Required after trial | Not required |

### Railway troubleshooting

| Issue | Fix |
|-------|-----|
| Build fails on `pip install` | Check Build Logs; ensure `requirements.txt` is valid |
| App crashes on start | Verify `SECRET_KEY` is set; check Deploy Logs for Python errors |
| 502 during PPT generation | **Settings → Deploy** → increase health check timeout; gunicorn timeout is already 120s in `Dockerfile` |
| Out of credits | Add payment method or switch to Render / Oracle Cloud free tier |
| Wrong port | Do not set `PORT` manually — Railway injects it; `Dockerfile` uses `${PORT}` |

---

## Option C — Oracle Cloud Always Free

[Oracle Cloud Infrastructure (OCI)](https://www.oracle.com/cloud/free/) offers an **Always Free** tier that does **not expire** — unlike trials. You get a real VM (up to **4 ARM cores + 24 GB RAM** total) where you run Docker. The app stays **always on** with no cold starts.

**Cost:** $0/month forever (credit card required for signup, not charged if you stay in Always Free resources).

**Repo config:** `Dockerfile`, `docker-compose.yml`

### What you get (Always Free)

| Resource | Limit |
|----------|-------|
| ARM VM (`VM.Standard.A1.Flex`) | Up to 4 OCPUs, 24 GB RAM (split across instances) |
| Block storage | 200 GB |
| Outbound data | 10 TB/month |
| Expiry | **None** — always free |

For this app, **2 OCPUs + 6 GB RAM** is plenty.

### Step 1 — Create an OCI account

1. Go to [oracle.com/cloud/free](https://www.oracle.com/cloud/free/)
2. Sign up (credit card required for verification; Always Free resources are not billed)
3. Choose your home region (e.g. `us-ashburn-1`, `uk-london-1`) — **cannot be changed later**

### Step 2 — Create a compute instance

1. OCI Console → **Compute** → **Instances** → **Create instance**
2. Name: `telus-excel-automation`
3. Image: **Ubuntu 22.04** (or 24.04)
4. Shape: click **Change shape** → **Ampere** → **VM.Standard.A1.Flex**
   - OCPUs: **2**
   - Memory: **6 GB**
5. Networking: use default VCN; assign a **public IPv4 address**
6. SSH keys: generate or upload your public key (save the private key)
7. Click **Create**

> If you see **Out of capacity**, try a different availability domain or region, or retry later.

### Step 3 — Open firewall ports

**A. OCI Security List (cloud firewall)**

1. **Networking** → **Virtual cloud networks** → your VCN → **Security Lists** → default
2. **Add Ingress Rules:**

   | Source CIDR | Protocol | Destination port |
   |-------------|----------|------------------|
   | `0.0.0.0/0` | TCP | 22 (SSH) |
   | `0.0.0.0/0` | TCP | 80 (HTTP) |
   | `0.0.0.0/0` | TCP | 443 (HTTPS) |
   | `0.0.0.0/0` | TCP | 8080 (app, optional for testing) |

**B. Ubuntu firewall (on the VM)**

```bash
sudo iptables -I INPUT -p tcp --dport 80 -j ACCEPT
sudo iptables -I INPUT -p tcp --dport 443 -j ACCEPT
sudo iptables -I INPUT -p tcp --dport 8080 -j ACCEPT
sudo netfilter-persistent save   # if available
```

### Step 4 — SSH into the VM and install Docker

```bash
ssh -i your-key.pem ubuntu@YOUR_PUBLIC_IP

# Install Docker
sudo apt update && sudo apt install -y docker.io docker-compose-v2 git
sudo usermod -aG docker ubuntu
# Log out and back in for group to apply
exit
ssh -i your-key.pem ubuntu@YOUR_PUBLIC_IP
```

### Step 5 — Clone repo and deploy

```bash
git clone https://github.com/prasun-iitj/Telus_Mukkamala_Excel_Automation.git
cd Telus_Mukkamala_Excel_Automation

# Set production secret
export SECRET_KEY="your-long-random-secret-here"

# Build and run (docker-compose.yml in repo)
docker compose up -d --build

# Verify
docker compose ps
docker compose logs -f web
```

Test: open `http://YOUR_PUBLIC_IP:8080` in a browser.

### Step 6 — HTTPS with Nginx (optional but recommended)

Install Nginx as a reverse proxy on port 80/443:

```bash
sudo apt install -y nginx certbot python3-certbot-nginx

sudo tee /etc/nginx/sites-available/telus-ppt << 'EOF'
server {
    listen 80;
    server_name YOUR_DOMAIN_OR_IP;

    client_max_body_size 50M;

    location / {
        proxy_pass http://127.0.0.1:8080;
        proxy_read_timeout 300s;
        proxy_connect_timeout 300s;
        proxy_send_timeout 300s;
    }
}
EOF

sudo ln -s /etc/nginx/sites-available/telus-ppt /etc/nginx/sites-enabled/
sudo nginx -t && sudo systemctl reload nginx
```

If you have a domain pointing to the VM IP:

```bash
sudo certbot --nginx -d your-domain.example.com
```

Then use `https://your-domain.example.com` (no `:8080`).

### Step 7 — Auto-restart and updates

Docker Compose is configured with `restart: unless-stopped` — the app survives reboots.

**Update after a git push:**

```bash
cd ~/Telus_Mukkamala_Excel_Automation
git pull origin main
export SECRET_KEY="your-long-random-secret-here"
docker compose up -d --build
```

Optional: add a cron job or GitHub Action to SSH and run `git pull && docker compose up -d --build`.

### Oracle Cloud troubleshooting

| Issue | Fix |
|-------|-----|
| Out of capacity | Try another availability domain or home region |
| Cannot SSH | Check security list allows port 22; verify public IP |
| Site unreachable on :8080 | Open port 8080 in OCI security list + Ubuntu iptables |
| PPT generation timeout | Nginx `proxy_read_timeout 300s` (see Step 6) |
| `docker compose` not found | Use `docker compose` (v2) or `docker-compose` (v1) |
| ARM build errors | `python:3.12-slim` image supports ARM64 — should work out of the box |
| Low memory | Use at least 4 GB RAM on the VM shape |

### Oracle vs Railway vs Render

| | Oracle Always Free | Railway | Render free |
|---|-------------------|---------|-------------|
| Monthly cost | **$0** | ~$5–15 | **$0** |
| Always on | **Yes** | Yes | No (sleeps) |
| Setup effort | High (VM + Docker) | Low | Lowest |
| Best for | Long-term free production | Fast always-on PaaS | Quick internal demo |

---

## Option D — Fly.io

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

## Option E — Google Cloud Run

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

## Option F — Azure App Service (Docker)

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

## Option G — DigitalOcean App Platform

1. [cloud.digitalocean.com/apps](https://cloud.digitalocean.com/apps) → **Create App**
2. Connect GitHub repo
3. Type: **Dockerfile** (auto-detected)
4. HTTP port: **8080**
5. Add env: `SECRET_KEY`
6. Instance size: Basic (512 MB RAM minimum)

---

## Option H — PythonAnywhere

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
