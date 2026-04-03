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
- Railway will use [Dockerfile](/C:/Project/a3presentation/Dockerfile)

Environment variables:

```bash
TEMPLATES_DIR=/data/templates
OUTPUTS_DIR=/data/outputs
```

Optional:

```bash
CORS_ORIGINS=https://your-frontend-domain.up.railway.app
```

## Railway volume

Attach a Railway Volume to the backend service and mount it to:

```text
/data
```

The application seeds bundled repository templates into `/data/templates` on startup if the mounted volume is empty.

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
- Railway will use [frontend/Dockerfile](/C:/Project/a3presentation/frontend/Dockerfile)

Environment variables:

```bash
VITE_API_BASE_URL=https://your-backend-domain.up.railway.app
```

## Order of setup

1. Create a new Railway project from this GitHub repository
2. Create the backend service from the repo root
3. Attach a volume to the backend service at `/data`
4. Set `TEMPLATES_DIR=/data/templates` and `OUTPUTS_DIR=/data/outputs` on the backend
5. Create the frontend service from `/frontend`
6. Set `VITE_API_BASE_URL` on the frontend to the backend public URL
7. Set `CORS_ORIGINS` on the backend to the frontend public URL
8. Deploy both services

## Notes

- Railway volumes are mounted at runtime, not build time
- if you need template files to persist after upload, they must be written under the mounted volume
- current repository templates in `storage/templates/` are part of the codebase, but newly uploaded templates should live on the mounted backend volume
