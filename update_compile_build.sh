#!/bin/sh

# Fail on errors.
set -e

echo "updating docker base images..."
docker pull tobix/pywine

./compile_requirements.sh
./build.sh
