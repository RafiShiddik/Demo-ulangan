FROM python:3.10-slim

WORKDIR /app

# Install dependencies for lxml and python-docx
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    libxml2-dev \
    libxslt-dev \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Set environment variables
ENV PORT=8080
EXPOSE 8080

CMD gunicorn --bind 0.0.0.0:${PORT:-8080} --workers 2 --threads 4 app:app
