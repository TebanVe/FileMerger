# Use Python 3.11 slim image as base
FROM python:3.11-slim

# Set metadata
LABEL maintainer="FileMerger Project"
LABEL description="File Merger - Merge Excel and CSV files from subdirectories"
LABEL version="1.0.0"

# Set working directory
WORKDIR /app

# Install minimal system dependencies
RUN apt-get update && apt-get install -y \
    libgomp1 \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements file first (for better Docker layer caching)
COPY requirements.txt .

# Install Python dependencies with fallback for xlwings
RUN pip install --no-cache-dir -r requirements.txt || \
    (echo "Some packages failed, trying without xlwings..." && \
     pip install --no-cache-dir pandas openpyxl xlrd numpy python-dateutil pytz six tzdata et_xmlfile lxml psutil)

# Copy source code
COPY src/ ./src/

# Create data directory for mounting
RUN mkdir -p /app/data

# Set environment variables
ENV PYTHONPATH=/app/src
ENV PYTHONUNBUFFERED=1

# Create a non-root user for security
RUN useradd --create-home --shell /bin/bash filemerger
RUN chown -R filemerger:filemerger /app
USER filemerger

# Set the default command
ENTRYPOINT ["python", "src/merge_excel_files.py"]
CMD ["/app/data"]

# Health check (optional)
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD python -c "import pandas, openpyxl, xlwings; print('Health check passed')" || exit 1
