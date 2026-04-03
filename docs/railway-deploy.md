# Railway Deployment

This project should be deployed to Railway as two services inside one Railway project:

- `a3presentation-backend`
- `a3presentation-frontend`

This keeps the whole product in one Railway project while preserving the current architecture.

## Why two services

- backend is FastAPI and writes files to `storage/`
- frontend is a separate Vite application
- Railway services are container-based, so splitting them is the cleanest deployment model

## Railway backend service

Repository root:

```text
/
```

Suggested settings:

- Root Directory: `/`
- Watch Paths:
  - `/src/**`
  - `/tests/**`
  - `/storage/templates/**`
  - `/pyproject.toml`
- Start Command:

```bash
uvicorn a3presentation.main:app --host 0.0.0.0 --port $PORT
```

Environment variables:

```bash
STORAGE_DIR=/app/storage
```

Optional:

```bash
CORS_ORIGINS=https://your-frontend-domain.up.railway.app
```

## Railway volume

Attach a Railway Volume to the backend service and mount it to:

```text
/app/storage
```

This matches Railway's documented guidance for apps that write to a relative project path and need persistence.

The backend will then persist:

- `storage/templates/`
- `storage/outputs/`

## Railway frontend service

Repository root:

```text
/frontend
```

Suggested settings:

- Root Directory: `/frontend`
- Watch Paths:
  - `/frontend/**`
- Build Command:

```bash
yarn install --frozen-lockfile && yarn build
```

- Start Command:

```bash
yarn vite preview --host 0.0.0.0 --port $PORT
```

Environment variables:

```bash
VITE_API_BASE_URL=https://your-backend-domain.up.railway.app
```

## Order of setup

1. Create a new Railway project from this GitHub repository
2. Create the backend service from the repo root
3. Attach a volume to the backend service at `/app/storage`
4. Create the frontend service from `/frontend`
5. Set `VITE_API_BASE_URL` on the frontend to the backend public URL
6. Set `CORS_ORIGINS` on the backend to the frontend public URL
7. Deploy both services

## Notes

- Railway volumes are mounted at runtime, not build time
- if you need template files to persist after upload, they must be written under the mounted volume
- current repository templates in `storage/templates/` are part of the codebase, but newly uploaded templates should live on the mounted backend volume
