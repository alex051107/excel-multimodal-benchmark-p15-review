#!/usr/bin/env sh
set -eu
mkdir -p /app/output
cp "$(dirname "$0")/reference.xlsx" /app/output/answer.xlsx
