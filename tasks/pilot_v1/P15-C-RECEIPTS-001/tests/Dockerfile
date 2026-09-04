FROM python:3.11.16-slim-bookworm@sha256:528257d48c1da0dcecc2e725d1ae34498d60c965f1241e39cd6a85a8859bdf84
RUN pip install --no-cache-dir openpyxl==3.1.5
RUN mkdir -p /app/output /logs/verifier
COPY . /tests
RUN chmod +x /tests/test.sh
WORKDIR /app
