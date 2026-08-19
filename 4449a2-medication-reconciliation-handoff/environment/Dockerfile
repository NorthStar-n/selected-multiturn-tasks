FROM ubuntu:24.04

WORKDIR /app

# System packages
RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 python3-pip python3-venv \
    nodejs npm \
    libreoffice-calc libreoffice-writer libreoffice-impress \
    ffmpeg pandoc tesseract-ocr qpdf ghostscript imagemagick \
    libcairo2 libpango-1.0-0 libpangocairo-1.0-0 libgdk-pixbuf-2.0-0 shared-mime-info \
    curl wget git unzip \
    && rm -rf /var/lib/apt/lists/*

# Ensure 'python' is available (Ubuntu 24.04 only ships python3)
RUN ln -sf /usr/bin/python3 /usr/bin/python

# Install uv
COPY --from=ghcr.io/astral-sh/uv:latest /uv /usr/local/bin/uv

# Python packages
RUN uv pip install --system --break-system-packages \
    cairosvg==2.7.1 \
    pytest==8.3.4 \
    harbor-rewardkit==0.1.7 \
    httpx==0.28.1 \
    pydantic==2.10.4 \
    pyyaml==6.0.2 \
    openpyxl==3.1.5 \
    python-docx==1.1.2 \
    python-pptx==1.0.2 \
    pandas==2.2.3 \
    numpy==2.2.1 \
    scipy==1.15.0 \
    scikit-learn==1.6.1 \
    statsmodels==0.14.4 \
    pymupdf==1.25.1 \
    reportlab==4.2.5 \
    pdfplumber==0.11.4 \
    pypdf==5.1.0 \
    xlsxwriter==3.2.0 \
    tabulate==0.9.0 \
    matplotlib==3.10.0 \
    scienceplots==2.2.2 \
    seaborn==0.13.2 \
    pillow==12.3.0 \
    docx2python==2.10.1

# npm tools for document generation
ENV NODE_PATH=/usr/local/lib/node_modules
RUN npm install -g \
    docx@9.5.3 \
    pptxgenjs@4.0.1 \
    skia-canvas@3.0.8 \
    echarts@6.1.0 \
    chart.js@4.5.1 \
    mermaid@11.16.0 \
    playwright@1.61.1

# Install one browser for local evidence capture; generated tasks remain offline.
RUN npx playwright install --with-deps chromium \
    && rm -rf /var/lib/apt/lists/*

# Copy input data (read-only)
COPY data/ /app/data/
RUN chmod -R 444 /app/data/

# Create output directory
RUN mkdir -p /app/output

# Materialize the resumed workspace from this environment
USER root
RUN mkdir -p /app && find /app -mindepth 1 -maxdepth 1 -exec rm -rf -- {} +
COPY . /app/
WORKDIR /app
