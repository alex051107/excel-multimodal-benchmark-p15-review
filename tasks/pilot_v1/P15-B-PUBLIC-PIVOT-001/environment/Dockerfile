FROM python:3.11-slim
RUN pip install --no-cache-dir openpyxl==3.1.5
WORKDIR /app
RUN mkdir -p /app/input /app/output /logs/verifier
COPY input/ /app/input/
