#!/bin/sh

# Fail on errors.
set -e

echo "building docker images..."
docker build -t pptx2h5p/pyinstaller -f docker/Dockerfile .

echo "compiling requirements.text from requirements.in..."
docker run -v ".:/src/" pptx2h5p/pyinstaller wine pip-compile