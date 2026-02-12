# Use an official Python runtime as a parent image
FROM python:3.10-slim

# Set the working directory in the container
WORKDIR /app

# Copy the requirements file first to leverage Docker cache
COPY requirements.txt .

# Install dependencies
# Doing this before copying the entire application allows us to cache the installed dependencies layer
# and not reinstall them on every build unless requirements.txt changes
RUN pip install -r requirements.txt
#pip install --upgrade pip && \
    #pip install --no-cache-dir -r requirements.txt
#RUN apt-get install unixodbc

# Now copy the rest of the application into the container
COPY icon_image_lib/ icon_image_lib/
COPY main.py .
COPY app_config.py .
COPY email_utils.py .
COPY s3_utils.py .
# Clean the apt cache and update with --fix-missing
RUN apt-get clean && \
    apt-get update --fix-missing

# Install necessary packages
RUN apt-get install -y apt-transport-https curl gnupg lsb-release unixodbc unixodbc-dev
# Add Microsoft package repository and install msodbcsql17 (modern approach without apt-key)
RUN curl -fsSL https://packages.microsoft.com/keys/microsoft.asc | gpg --dearmor -o /usr/share/keyrings/microsoft-prod.gpg && \
    echo "deb [arch=amd64 signed-by=/usr/share/keyrings/microsoft-prod.gpg] https://packages.microsoft.com/debian/12/prod bookworm main" > /etc/apt/sources.list.d/mssql-release.list && \
    apt-get update --fix-missing && \
    ACCEPT_EULA=Y apt-get install -y msodbcsql17 && \
    ACCEPT_EULA=Y apt-get install -y mssql-tools

# Set PATH to include mssql-tools
ENV PATH="/opt/mssql-tools/bin:${PATH}"

# Verify installation of unixODBC
RUN which odbcinst

# Create temp directories with proper permissions
RUN mkdir -p /app/temp_files/images /app/temp_files/excel /app/jobs && \
    chmod -R 777 /app/temp_files /app/jobs

# LABEL "com.datadoghq.ad.logs"='[<LOGS_CONFIG>]'
# Make port 8000 available to the world outside this container
EXPOSE 8080

# Run with uvicorn for production with multiple workers for concurrent requests
# Workers = (2 x CPU cores) + 1 is a common formula, defaulting to 4 workers
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8080", "--workers", "4"]
