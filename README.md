# Factory Scanner

Upload up to 5 factory notebook images, extract row values with Groq vision, review and correct each table, then download separate Excel files with row totals, column totals, and a grand total.

## Local setup

Install Python 3.12+ (recommended), then:

```bash
pip install -r requirements.txt
```

Create a `.env` file (copy from `.env.example`) and set at least:

```text
GROQ_API_KEY=your_key_here
```

Optional: `GROQ_MODEL`, `PORT`, `WEB_CONCURRENCY`.

## Run locally

**Mac / Linux**

```bash
./start.sh
```

or:

```bash
python3 server.py
```

**Windows**

```bat
start.bat
```

Open the app at **http://localhost:5050** (or the `PORT` you set).

## Deploy on Vercel (free Hobby plan)

Vercel runs this app as a **Python serverless function** (Flask) and serves the UI from the **`public/`** folder. Follow these steps end to end.

### 1. Prepare the repository

1. Put the project in a Git repository (GitHub, GitLab, or Bitbucket).
2. **Do not commit `.env`** — it is listed in `.gitignore`. Secrets belong only in Vercel’s dashboard.
3. Confirm the repo contains at least: `server.py`, `requirements.txt`, `public/index.html`, `vercel.json`, `runtime.txt`.

### 2. Create a Vercel account

1. Go to [https://vercel.com](https://vercel.com) and sign up (GitHub login is fine).
2. Verify your email if asked.

### 3. Import the project

1. In the Vercel dashboard, click **Add New…** → **Project**.
2. **Import** your Git repository (`factory-scanner` or whatever you named it).
3. Vercel auto-detects Flask from **`server.py`** (root) — do **not** add a `functions` entry in `vercel.json` for `server.py`; that pattern only matches files under `api/` and will fail the build.
4. Leave **Root Directory** as the repo root unless the app lives in a subfolder.
5. **Framework Preset** can stay “Other” or whatever Vercel suggests for Python — no separate build command is required for this app.

### 4. Configure environment variables (required)

On the import screen (or later under **Project → Settings → Environment Variables**), add:

| Name | Value | Environments |
|------|--------|----------------|
| `GROQ_API_KEY` | Your Groq API key | Production, Preview, Development |

Optional:

| Name | Value |
|------|--------|
| `GROQ_MODEL` | e.g. `meta-llama/llama-4-scout-17b-16e-instruct` |

Save. Without `GROQ_API_KEY`, `/extract` will return an error.

### 5. Deploy

1. Click **Deploy**.
2. Wait for the build to finish. Open the production URL Vercel shows (e.g. `https://factory-scanner-xxx.vercel.app`).

### 6. Smoke test after deploy

1. Open the production URL — you should see the Factory Scanner UI.
2. Call **https://your-app.vercel.app/health** — expect JSON with `"ok": true` and your model name.
3. Upload a small test image and run **Read Numbers** once to confirm Groq works with the env var.

### 7. Optional: custom domain

Under **Project → Settings → Domains**, add your domain and follow Vercel’s DNS instructions.

---

### Vercel-specific notes

- **`public/index.html`** is served from the CDN; **`server.py`** handles `/extract`, `/health`, downloads, etc.
- **`vercel.json`** only includes the schema (no `functions` block). In the Vercel dashboard go to **Project → Settings → Functions** and set **Default / Function Max Duration** to **120 seconds** (or the highest your plan allows) so Groq vision requests can finish.
- **`runtime.txt`** pins **Python 3.12** for consistent installs.
- **`.vercelignore`** keeps `.env` and other junk out of uploads when using the CLI.
- **Cold starts**: the first request after idle may be slower; this is normal on serverless free tiers.
- **Alternatives**: If you prefer a traditional always-on process (e.g. long-running workers, WebSockets), consider [Railway](https://railway.app) or [Render](https://render.com) with the included `Procfile` / `gunicorn` setup instead.

## Configuration summary

| Variable | Required | Description |
|----------|----------|-------------|
| `GROQ_API_KEY` | Yes (prod) | Groq API key |
| `GROQ_MODEL` | No | Vision model ID |
| `PORT` | No | Local port (default `5050`) |
| `WEB_CONCURRENCY` | No | Gunicorn workers for `./start.sh` |

## Tips

- Good lighting and a straight-on photo improve extraction accuracy.
- Images are compressed and enhanced before sending to Groq.
- Use **Create blank table** to enter values manually.
- Downloaded Excel files use formulas for totals.
