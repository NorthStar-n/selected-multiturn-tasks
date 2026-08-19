FROM python:3.12-slim

WORKDIR /app

RUN pip install --no-cache-dir \
    docx2python==2.10.1 \
    httpx==0.28.1 \
    openpyxl==3.1.5 \
    pandas==2.2.3 \
    pillow==11.1.0 \
    pydantic==2.10.4 \
    pymupdf==1.25.1 \
    python-docx==1.1.2 \
    python-pptx==1.0.2 \
    pytest==8.3.4

COPY data/ /app/data/
RUN mkdir -p /app/output
