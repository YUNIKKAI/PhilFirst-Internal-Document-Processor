# Base image: lightweight Python
FROM python:3.11-slim

# Set working directory inside container
WORKDIR /app

# Install system dependencies (needed for some Python packages)
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Copy dependency list and install
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the whole project into the container
COPY . .

# Create logs directory (for gunicorn logs)
RUN mkdir -p logs

# Expose port (same as your .env PORT)
EXPOSE 8080

# Run Gunicorn with your app factory
CMD ["gunicorn", "-c", "gunicorn.conf.py", "app:create_app()"]
