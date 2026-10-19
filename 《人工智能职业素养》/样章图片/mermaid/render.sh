#!/bin/bash
export FONTCONFIG_FILE=/tmp/mmd/fcdir/fonts.conf FONTCONFIG_PATH=/tmp/mmd/fcdir LD_LIBRARY_PATH=/tmp/al2023x/lib
cd "$(dirname "$0")"
for f in *.mmd; do
  n="${f%.mmd}"
  timeout 200 /tmp/mmd/node_modules/.bin/mmdc -p /tmp/mmd/pp.json -i "$f" -o "$n.png" -b white -s 2 >/dev/null 2>&1 && echo "OK $n $(stat -c%s $n.png)" || echo "FAIL $n"
done
