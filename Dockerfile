# Use Python 3.12 slim image
FROM python:3.12-slim

# Install system dependencies including Pandoc
RUN apt-get update && apt-get install -y \
    pandoc \
    && rm -rf /var/lib/apt/lists/*

# Verify Pandoc installation
RUN pandoc --version

# Set working directory
WORKDIR /app

# Copy requirements and install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Create necessary directories
RUN mkdir -p uploads outputs

# Expose port (Railway will set PORT env var)
EXPOSE 8080

# Run the application
CMD ["python", "app.py"]
