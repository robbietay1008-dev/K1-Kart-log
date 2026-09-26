FROM python:3.12-slim
WORKDIR /srv
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app.py merge.py admin.html entrypoint.sh ./
COPY web ./web
RUN chmod +x entrypoint.sh
ENV KARTLOG_DATA=/data KARTLOG_APP=/srv/web
EXPOSE 8090 8091
CMD ["./entrypoint.sh"]
