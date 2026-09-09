# Deployment

The Flask process serves the public site, the legacy portal, the ops API and the built ops SPA. Deploying now has one extra step: **build the frontend before starting Python**.

## Build pipeline

```bash
# 1. frontend
cd apps/ops-web
npm ci
npm run build          # -> apps/ops-web/dist (git-ignored)

# 2. backend
cd ../..
pip install -r requirements.txt
python manage.py upgrade   # alembic upgrade + idempotent seeds (safe to re-run)
gunicorn "asme:create_app()" --bind 0.0.0.0:$PORT   # or: python manage.py serve (dev)
```

Flask looks for `apps/ops-web/dist/index.html`; override the location with `ASME_OPS_WEB_DIST=/abs/path/dist`. If the build is missing, `/app` shows a plain page explaining how to build it instead of a broken app.

## Render

`render.yaml` / build command:

```
cd apps/ops-web && npm ci && npm run build && cd ../.. && pip install -r requirements-render.txt
```

Start command: `python manage.py upgrade && gunicorn "asme:create_app()"`. Node 20+ is available on Render's Python runtime via `NODE_VERSION`; set it to `20`.

## Elastic Beanstalk

Add a `.platform/hooks/prebuild/01_frontend.sh` that runs the frontend build (Node must be installed in the platform hook or the `dist` folder committed to the deployment bundle by CI). The simplest reliable path is to build in CI and include `apps/ops-web/dist` in the zip that is uploaded; `.gitignore` excludes it from source control only.

## Environment

| Variable | Purpose |
| --- | --- |
| `ASME_ENV` | `development` / `production` / `testing` |
| `ASME_SECRET_KEY` | session signing (required in production) |
| `ASME_DATABASE_URL` | `postgresql://…` in production |
| `ASME_OPS_ENABLED` | default `1`; `0` disables the ops API + SPA |
| `ASME_OPS_WEB_DIST` | optional path to the built SPA |
| `ASME_FILE_MAX_MB` | attachment size limit (default 25) |
| `ASME_FILE_URL_TTL_SECONDS` | signed download link lifetime (default 900) |
| `ASME_ENABLE_LEGACY_OPS` | keep the legacy portal reachable at `/legacy/app` |
| `ASME_SESSION_COOKIE_SECURE` | `1` behind HTTPS |

Never commit `.env`. Seed credentials (`manage.py seed-demo`) are for development only and are refused nowhere by code – do not run the command against production.

## Health & smoke checks

- `GET /healthz` – legacy health.
- `GET /api/v1/ops/health` – ops API.
- `GET /app` – SPA shell (200 with HTML).
- `GET /api/v1/ops/openapi.json` – contract.

## Local development

```bash
python manage.py seed-demo        # one-off: demo chapter data, password ChangeMe123!
python manage.py serve            # Flask on :5000 (serves the built SPA if present)
cd apps/ops-web && npm run dev    # Vite on :5173 with /api proxied to :5000; open http://localhost:5173/app
```
