# Base image: lightweight Python on Debian 12 (Bookworm - stable)
FROM python:3.11-slim-bookworm

# Set working directory inside container
WORKDIR /app

# Optional: Proxy configuration (uncomment if you need it inside container)
# ARG http_proxy
# ARG https_proxy
# ENV http_proxy=${http_proxy}
# ENV https_proxy=${https_proxy}

# Install system dependencies (needed for some Python packages)
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Copy dependency list and install Python packages
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
