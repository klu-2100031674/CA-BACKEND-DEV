# Use Node.js as the primary base image
FROM node:20-bookworm-slim

# Install Python, Pip, and LibreOffice (headless)
# We use bookworm-slim to keep it relatively light, but we need libreoffice
RUN apt-get update && apt-get install -y \
    python3 \
    python3-pip \
    python3-venv \
    libreoffice-calc \
    libreoffice-writer \
    libreoffice-impress \
    default-jre \
    fontconfig \
    fonts-liberation \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy Node.js dependency definitions
COPY package*.json ./

# Install Node.js dependencies
RUN npm install --production

# Note: We rely on the python3-pip package installed via apt-get in the base image section.
# Debian bookworm strictly manages python packages (PEP 668).
# We should NOT upgrade system pip globally or use it without a venv.
# We will install dependencies strictly inside the virtual environment later.

# Copy Python requirements
COPY src/python-engine/requirements.txt ./src/python-engine/requirements.txt

# Create a virtual environment for Python to avoid breaking system packages
ENV VIRTUAL_ENV=/opt/venv
RUN python3 -m venv $VIRTUAL_ENV
ENV PATH="$VIRTUAL_ENV/bin:$PATH"

# Install Python dependencies
# We filter out pywin32 because it's Windows-only
RUN sed '/pywin32/d' src/python-engine/requirements.txt > requirements-linux.txt && \
    pip install --no-cache-dir -r requirements-linux.txt

# Copy the rest of the application code
COPY . .

# Environment variables
ENV NODE_ENV=production
ENV USE_COM_INTERFACE=False
# Make sure the app knows where python is (if you use 'python' command in node)
# server.js needs to invoke 'python' which is now in the venv path due to ENV PATH above.

# Expose the port the app runs on
EXPOSE 10000

# Start the application
CMD ["node", "server.js"]
