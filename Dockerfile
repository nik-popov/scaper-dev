# Use an official Python runtime as a parent image
FROM python:3.10-slim

# Set the working directory
WORKDIR /app

# Install system dependencies
RUN apt-get update --fix-missing && \
    apt-get install -y --no-install-recommends \
        apt-transport-https \
        curl \
        gnupg \
        lsb-release \
        unixodbc \
        unixodbc-dev \
        libpq-dev \
        build-essential \
        libopencv-dev \
        && apt-get clean && rm -rf /var/lib/apt/lists/*

# Add Microsoft package repository and install msodbcsql17
RUN curl -fsSL https://packages.microsoft.com/keys/microsoft.asc | gpg --dearmor -o /usr/share/keyrings/microsoft-prod.gpg && \
    echo "deb [arch=amd64 signed-by=/usr/share/keyrings/microsoft-prod.gpg] https://packages.microsoft.com/debian/12/prod bookworm main" > /etc/apt/sources.list.d/mssql-release.list && \
    apt-get update --fix-missing && \
    ACCEPT_EULA=Y apt-get install -y msodbcsql17 mssql-tools && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

# Set PATH for mssql-tools
ENV PATH="/opt/mssql-tools/bin:${PATH}"

# Verify unixODBC
RUN which odbcinst && odbcinst -j

# Install uv
RUN pip install --no-cache-dir uv

COPY pyproject.toml /app/ 
# Copy dependency files
COPY uv.lock /app/

# Copy the rest of the application
COPY app/ /app/

# Expose ports
EXPOSE 8080
EXPOSE 8265

# Generate or update uv.lock based on pyproject.toml
RUN ["uv", "lock"]

# Synchronize the virtual environment with uv.lock
RUN ["uv", "sync"]

# Set the entrypoint to use uv run for all commands
ENTRYPOINT ["uv", "run"]

# Default command to run main.py
CMD ["python", "main.py"]