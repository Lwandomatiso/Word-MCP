# Use official Python runtime
FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Install build dependencies
RUN apt-get update \
    && apt-get install -y --no-install-recommends build-essential \
    && rm -rf /var/lib/apt/lists/*

# Copy project files
COPY . /app

# Install Python dependencies
RUN pip install --no-cache-dir .

# Expose the Cloud Run default port
EXPOSE 8080

# Set PORT environment variable (Cloud Run will override it)
ENV PORT 8080

# Start your existing server, using $PORT and binding to 0.0.0.0
# Assumes word_mcp_server can accept --host and --port args. If not, see note below.
ENTRYPOINT ["word_mcp_server"]
CMD ["--host", "0.0.0.0", "--port", "8080"]
