#!/bin/sh
set -eu
cd "$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)"
action=${1:-start}
case "$action" in start|stop|status|logs|update) ;; *) printf '%s\n' 'Usage: ./manage.sh [start|stop|status|logs|update] [domain]'; exit 2 ;; esac
command -v docker >/dev/null 2>&1 || { printf '%s\n' 'Docker with Compose v2 is required.' >&2; exit 1; }
docker compose version >/dev/null
if [ ! -f .env.production ]; then
    if [ "$action" != start ]; then printf '%s\n' 'Run ./manage.sh start first.' >&2; exit 1; fi
    domain=${2:-}
    if [ -z "$domain" ]; then printf '%s' 'Public domain connected to this server (example: editor.example.com): '; IFS= read -r domain; fi
    domain=$(printf '%s' "$domain" | tr '[:upper:]' '[:lower:]')
    if [ "${#domain}" -gt 253 ] || ! printf '%s\n' "$domain" | LC_ALL=C grep -Eq '^[a-z0-9]([a-z0-9-]{0,61}[a-z0-9])?(\.[a-z0-9]([a-z0-9-]{0,61}[a-z0-9])?)+$'; then
        printf '%s\n' 'Enter a hostname without https://, a port, or a path.' >&2; exit 2
    fi
    (umask 077; set -C; printf 'JAKAL_DOMAIN=%s\nJAKAL_MEMORY_LIMIT=3g\nJAKAL_PROXY_SUBNET=172.30.88.0/29\nJAKAL_PROXY_ADDRESS=172.30.88.2\n' "$domain" > .env.production)
fi
docker compose --env-file .env.production config --quiet
case "$action" in
    start|update)
        docker compose --env-file .env.production up -d --build --wait --wait-timeout 180
        printf '%s\n' 'Services started. The configured domain must point to this server, with ports 80 and 443 reachable.'
        ;;
    stop) docker compose --env-file .env.production stop ;;
    status) docker compose --env-file .env.production ps ;;
    logs) docker compose --env-file .env.production logs --tail 100 -f ;;
esac
