FROM python:3.12-slim

WORKDIR /app

# Copy application files
COPY pyproject.toml .
COPY *.py ./
COPY config/ ./config/
COPY lib/ ./lib/

# Install dependencies
RUN pip install --no-cache-dir chardet openpyxl

# Create working directory for data files
WORKDIR /data

# Set the entrypoint to the main script
ENTRYPOINT ["python", "/app/main.py"]
