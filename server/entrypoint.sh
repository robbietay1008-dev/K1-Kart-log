#!/bin/sh
# public app (published by Funnel) + admin app (tailnet only) from the same code and database
uvicorn app:admin --host 0.0.0.0 --port 8091 --log-level warning &
exec uvicorn app:public --host 0.0.0.0 --port 8090 --proxy-headers --forwarded-allow-ips="*"
