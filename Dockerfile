FROM python:3.11-slim

# Install LibreOffice (for xlsx -> pdf) + fonts
RUN apt-get update && DEBIAN_FRONTEND=noninteractive apt-get install -y \
    libreoffice \
    fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy app and assets
COPY ..

# Default port for Render
ENV PORT=10000
EXPOSE 10000

# Allow all origins by default; override in Render env if needed
ENV ALLOWED_ORIGINS=*

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "10000"]
