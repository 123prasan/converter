# Use Node.js 18 on Bullseye (Debian) for maximum compatibility with Python libs
FROM node:18-bullseye

# 1. Install System Dependencies
# These are required for Word/PPT conversion, PDF-to-Image, and OCR
RUN apt-get update && apt-get install -y \
    python3 \
    python3-pip \
    libreoffice \
    poppler-utils \
    tesseract-ocr \
    libgl1 \
    && rm -rf /var/lib/apt/lists/*

# 2. Set the working directory
WORKDIR /app

# 3. Handle Node.js Dependencies
# Ensure package.json exists before this step (Run 'npm init -y' locally)
COPY package*.json ./
RUN npm install

# 4. Handle Python Dependencies
# Ensure requirements.txt exists in your root folder
COPY requirements.txt ./
RUN pip3 install --no-cache-dir -r requirements.txt --break-system-packages

# 5. Copy Application Source
COPY . .

# 6. Final Setup
# Ensure the uploads folder exists for file processing
RUN mkdir -p uploads outputs

# Set environment variables for production
ENV PORT=3000
ENV PYTHONUTF8=1
ENV NODE_ENV=production

EXPOSE 3000

# Start the Node.js server
CMD ["node", "server.js"]