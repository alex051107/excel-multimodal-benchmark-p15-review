FROM python:3.11-slim
RUN pip install --no-cache-dir openpyxl
WORKDIR /app
RUN mkdir -p /app/input /app/output
COPY input/ /app/input/
