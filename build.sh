#!/usr/bin/env bash
set -o errexit

pip install -r requirements.txt

mkdir -p backend/uploads backend/converted
cp -r frontend backend/frontend
