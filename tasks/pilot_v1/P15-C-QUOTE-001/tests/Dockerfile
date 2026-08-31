FROM python:3.11-slim
RUN pip install --no-cache-dir openpyxl==3.1.5
RUN mkdir -p /app/output /logs/verifier
COPY . /tests
RUN chmod +x /tests/test.sh
WORKDIR /app
