# Psikotes Score Automation — API

Internal recap-matching service. **No UI here anymore** — the frontend is the
"Automated Recap" page in [Ordinat Dashboard](../ordinat-dashboard) (Next.js),
which proxies every request to this service. Every route requires
`Authorization: Bearer <RECAP_SERVICE_TOKEN>`; there's nothing to see by
visiting this service directly in a browser, and it should **never be exposed
to the public internet** — only the dashboard's server should be able to
reach it.

The old HTML template (`templates/index.html.bak`) is kept only as a
reference for the API's request/response shapes; it's not served.

## Quick Start
```bash
pip install -r requirements.txt
cp .env.example .env   # set RECAP_SERVICE_TOKEN to the SAME value as
                        # RECAP_SERVICE_TOKEN in ordinat-dashboard/.env
python app.py
# → http://localhost:5000 (or set PORT=5001 if 5000 is taken —
#   on macOS, AirPlay Receiver commonly squats on 5000)
```

Then set `RECAP_TOOL_URL` in `ordinat-dashboard/.env` to match whatever host/port this runs on.

## API

| Route | Method | Purpose |
|---|---|---|
| `/process` | POST | Start a recap job (multipart: `raw_file`, `rekap_file`, `threshold`, `manual_overrides`, `tgl_pemeriksaan`, `pendidikan`) → `{job_id}` |
| `/status/<job_id>` | GET | Poll job progress/result — `{status, progress, log, download_filename, error, pending?}` |
| `/review/<job_id>` | POST | Submit borderline-match decisions — `{decisions: {id: bool}}` |
| `/download/<filename>` | GET | Download the finished `.xlsx` |

## Docker
```bash
docker build -t psikotes .
docker run -p 5000:5000 -e RECAP_SERVICE_TOKEN=... psikotes
```

## Deploy (Render)

1. Push this repo to GitHub.
2. [render.com](https://render.com) → New → Web Service → connect the repo.
3. Runtime: **Docker** (the included `Dockerfile` is auto-detected — no build/start command to fill in manually).
4. Environment → add `RECAP_SERVICE_TOKEN` (same value as `ordinat-dashboard/.env`'s `RECAP_SERVICE_TOKEN`).
5. Instance type: **Free** is fine to start — note free instances sleep after 15 min idle, so the first request after a gap takes a few extra seconds to wake up.
6. Deploy. Copy the resulting `https://<name>.onrender.com` URL into `ordinat-dashboard`'s `RECAP_TOOL_URL`.

### Other platforms
Same Dockerfile works anywhere that runs containers — Railway (`railway init && railway up`), Fly.io (`fly launch && fly deploy`), DigitalOcean App Platform, Google Cloud Run, etc. Only one env var is required everywhere: `RECAP_SERVICE_TOKEN`.

Whatever platform you use, keep this service off the public internet where possible / treat its URL as internal — it has no other access control besides the shared token.
