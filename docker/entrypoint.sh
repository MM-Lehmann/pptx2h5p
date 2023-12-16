#!/bin/sh

# Fail on errors.
set -e

# Make sure .bashrc is sourced
. /root/.bashrc

wslmounts=$(awk '$3 == "9p" && $4 ~ /aname=drvfs/ { print $2 }' /proc/mounts)
if [ -n "$wslmounts" ]
then
	cat <<-EOW
###########################################################
Some paths are using WSL2 9p mounts. Using those with wine
might lead to hard to debug errors. See
   https://github.com/webcomics/pywine/issues/16
for details. If you encounter any problems, please retry
outside of these directories:

$wslmounts

###########################################################
	EOW
	sleep 3
fi

WORKDIR=${SRCDIR:-/src}

cd $WORKDIR

if [ $# -gt 0 ]
then
	exec "$@"
else
	wine pyinstaller --clean -y --dist ./dist/windows --workpath /tmp *.spec
	chown -R --reference=. ./dist/windows
fi
