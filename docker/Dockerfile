FROM tobix/pywine
LABEL author="Martin Lehmann (https://github.com/MM-Lehmann)"

RUN wine pip install pip-tools --no-warn-script-location --no-cache

COPY requirements.txt .
RUN wine pip install --no-warn-script-location --no-cache -r requirements.txt

COPY docker/entrypoint.sh /opt/
ENTRYPOINT ["/opt/entrypoint.sh"]