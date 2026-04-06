# Timeweb Cloud Server Deploy

This deployment mode uses one Linux server and Docker Compose.

## Architecture

- `backend`: FastAPI application
- `frontend`: static React build served by nginx
- `nginx`: public reverse proxy
- host `nginx` with Let's Encrypt terminates HTTPS and proxies into docker `nginx` on `127.0.0.1:8080`

Public traffic:

- `/` -> frontend
- `/api/*` -> backend

This means the browser uses one origin and does not need separate CORS setup.
For the Timeweb single-server setup, Docker publishes the internal app proxy only on `127.0.0.1:8080`, and host nginx handles public `80/443`.

## Server prerequisites

- Ubuntu 24.04 LTS
- public IPv4
- DNS A record pointed to the server IP if you want a domain

## Install Docker

Run on the server:

```bash
sudo apt update
sudo apt install -y ca-certificates curl gnupg
sudo install -m 0755 -d /etc/apt/keyrings
curl -fsSL https://download.docker.com/linux/ubuntu/gpg | sudo gpg --dearmor -o /etc/apt/keyrings/docker.gpg
sudo chmod a+r /etc/apt/keyrings/docker.gpg
echo \
  "deb [arch=$(dpkg --print-architecture) signed-by=/etc/apt/keyrings/docker.gpg] https://download.docker.com/linux/ubuntu \
  $(. /etc/os-release && echo \"$VERSION_CODENAME\") stable" | \
  sudo tee /etc/apt/sources.list.d/docker.list > /dev/null
sudo apt update
sudo apt install -y docker-ce docker-ce-cli containerd.io docker-buildx-plugin docker-compose-plugin git
sudo usermod -aG docker $USER
```

Log out and back in after adding your user to the `docker` group.

## Deploy project

```bash
git clone https://github.com/ignatovivin/a3presentation.git
cd a3presentation
mkdir -p data/outputs
bash scripts/deploy_server.sh
```

## Why templates are not mounted separately

- `storage/templates` in git is the source of truth
- backend image already contains `storage/templates`
- server deployment reads templates directly from `/app/storage/templates`
- only `data/outputs` stays on the host as persistent runtime storage

This avoids runtime drift between repo templates and production templates and removes the need to copy `template.pptx` into a separate volume on every deploy.

## Update project

```bash
cd a3presentation
git pull
bash scripts/deploy_server.sh
```

## Optional GitHub Actions auto-deploy

The repository now supports a minimal SSH deploy flow from GitHub Actions for `dev`.

Behavior:

- push to `dev`
- run backend, quality, frontend verify, and frontend smoke jobs
- if all pass, connect to Timeweb over SSH
- hard-reset server checkout to `origin/dev`
- run `bash scripts/deploy_server.sh`

Required GitHub repository secrets:

- `TIMEWEB_HOST`
- `TIMEWEB_USER`
- `TIMEWEB_SSH_KEY`
- `TIMEWEB_PORT` (optional, defaults to `22`)

Recommended server assumptions:

- project path on server: `~/a3presentation`
- deploy branch on server: `dev`
- host nginx + Let's Encrypt already configured separately

## Verify

Backend health:

```bash
curl http://127.0.0.1/api/health
```

Or from browser:

```text
http://YOUR_SERVER_IP/
```

## HTTPS

Recommended simplest next step:

- point a domain to the server
- install `certbot`
- either terminate TLS on host nginx or switch the edge proxy to Caddy/Traefik

If needed, this can be added after the first successful deploy.
