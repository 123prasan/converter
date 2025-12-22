# Use Node.js as the base image
FROM node:18-bullseye

# Install Python, LibreOffice, and PDF dependencies
RUN apt-get update && apt-get install -y \
    python3 \
    python3-pip \
    libreoffice \
    poppler-utils \
    tesseract-ocr \
    libgl1 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Node dependencies
COPY package*.json ./
RUN npm install

# Install Python dependencies
COPY requirements.txt ./
RUN pip3 install -r requirements.txt --break-system-packages

# Copy the rest of your code
COPY . .

# Set environment variables
ENV PORT=3000
ENV PYTHONUTF8=1

EXPOSE 3000

# Start the server
CMD ["node", "server.js"]