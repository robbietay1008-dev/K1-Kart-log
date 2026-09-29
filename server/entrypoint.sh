#!/bin/bash
# public app (published by Funnel) + admin app (tailnet only) from the same code and database.
# Both run as children of this script. If either one exits, the other is stopped and the container exits, so
# Docker's restart policy brings both back (before, a dead admin app went unnoticed while the container ran on).
# docker stop (SIGTERM) is passed on to both.
set -u
uvicorn app:admin --host 0.0.0.0 --port 8091 --log-level warning &
ADMIN=$!
uvicorn app:public --host 0.0.0.0 --port 8090 --proxy-headers --forwarded-allow-ips="*" &
PUBLIC=$!
stopping=0
trap 'stopping=1; kill -TERM "$ADMIN" "$PUBLIC" 2>/dev/null' TERM INT
wait -n "$ADMIN" "$PUBLIC"
rc=$?
if [ "$stopping" = 1 ]; then
  wait
  exit 0
fi
echo "entrypoint: a uvicorn process exited (rc=$rc); stopping the other so the container restarts" >&2
kill -TERM "$ADMIN" "$PUBLIC" 2>/dev/null
wait
exit $(( rc == 0 ? 1 : rc ))
