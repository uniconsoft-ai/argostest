FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV PORT=10000
ENV HOST=0.0.0.0

EXPOSE 10000

CMD gunicorn --bind 0.0.0.0:${PORT:-10000} --workers 2 --threads 8 --timeout 120 app:app