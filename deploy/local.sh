#!/bin/sh
set -eu
cd "$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)"
if [ ! -f compose.local.yaml ] && [ -f ../compose.local.yaml ]; then cd ..; fi
action=${1:-start}
case "$action" in start|stop|status|logs) ;; *) printf '%s\n' 'Usage: sh local.sh [start|stop|status|logs]'; exit 2 ;; esac
command -v docker >/dev/null 2>&1 || { printf '%s\n' 'Docker Engine 28+ with Compose v2 is required. Start Docker first.' >&2; exit 1; }
docker compose version >/dev/null
version=$(docker version --format '{{.Server.Version}}')
major=${version%%.*}
case "$major" in ''|*[!0-9]*) printf '%s\n' 'Unable to read the Docker Engine version.' >&2; exit 1 ;; esac
if [ "$major" -lt 28 ]; then printf '%s\n' 'Update to Docker Engine 28 or newer for local port isolation.' >&2; exit 1; fi
port=${JAKAL_LOCAL_PORT:-18988}
project=${JAKAL_LOCAL_PROJECT:-jakal-local}
case "$project" in ''|*[!a-z0-9_-]*|[-_]*) printf '%s\n' 'Invalid local project name.' >&2; exit 2 ;; esac
if [ "${#project}" -gt 60 ]; then printf '%s\n' 'Invalid local project name.' >&2; exit 2; fi
case "$port" in ''|*[!0-9]*) printf '%s\n' 'JAKAL_LOCAL_PORT must be an integer from 1024 to 65535.' >&2; exit 2 ;; esac
if [ "${#port}" -gt 5 ] || [ "$port" -lt 1024 ] || [ "$port" -gt 65535 ]; then printf '%s\n' 'JAKAL_LOCAL_PORT must be from 1024 to 65535.' >&2; exit 2; fi
docker compose --project-name "$project" --file compose.local.yaml config --quiet
case "$action" in
    start)
        docker compose --project-name "$project" --file compose.local.yaml up -d --build --wait --wait-timeout 240
        printf '\nOpen http://127.0.0.1:%s/ in your browser.\nStop: sh local.sh stop\n' "$port"
        ;;
    stop) docker compose --project-name "$project" --file compose.local.yaml stop ;;
    status) docker compose --project-name "$project" --file compose.local.yaml ps ;;
    logs) docker compose --project-name "$project" --file compose.local.yaml logs --tail 100 -f ;;
esac
