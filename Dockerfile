FROM python:3.11-slim

WORKDIR /app

# System dependencies for PDF generation
RUN apt-get update && apt-get install -y --no-install-recommends \
    wkhtmltopdf \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Create necessary directories
RUN mkdir -p /app/uploads /app/chroma_db

VOLUME ["/app/uploads", "/app/chroma_db"]

EXPOSE 8080

CMD ["python", "app.py"]
