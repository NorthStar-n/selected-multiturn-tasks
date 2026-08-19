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

# AFO_NATIVE_RESUME_APP_TASK_V1
RUN set -eu; \
    if [ ! -d /app/task ]; then \
      mkdir -p /app/task; \
      find /app -mindepth 1 -maxdepth 1 ! -name task \
        -exec sh -c 'for entry do mv "$entry" /app/task/; done' sh {} +; \
      find /app/task -mindepth 1 -maxdepth 1 \
        -exec sh -c 'for entry do ln -s "$entry" "/app/${entry##*/}"; done' sh {} +; \
    fi
WORKDIR /app/task

# Restore the exact native-resume workspace after source image setup
USER root
WORKDIR /
RUN rm -rf /app/task && mkdir -p /app/task
COPY . /app/task/
RUN rm -f /app/task/output/.gitkeep
WORKDIR /app/task
