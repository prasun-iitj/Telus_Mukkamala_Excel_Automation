FROM python:3.12-slim

WORKDIR /app

# Install dependencies first (better layer caching)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# App code + bundled PPT template
COPY . .

ENV PORT=8080
EXPOSE 8080

CMD gunicorn --bind 0.0.0.0:${PORT} --workers 2 --timeout 120 web_app:app
