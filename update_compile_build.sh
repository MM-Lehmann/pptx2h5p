#!/bin/sh

# Fail on errors.
set -e

echo "updating docker base images..."
docker pull tobix/pywine

./compile_requirements.sh

echo "building app..."
docker run -v ".:/src/" pptx2h5p/pyinstaller
ls dist/windows/pptx2h5p.exe -sh
