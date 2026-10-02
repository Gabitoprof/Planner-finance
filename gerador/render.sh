#!/bin/bash
# uso: render.sh "ABA1,ABA2" prefixo
cd /home/user/Planner-finance/gerador && python3 build.py --render "$1" || exit 1
cd /home/user/Planner-finance/.teste && timeout 300 soffice -env:UserInstallation=file:///tmp/lo_render --headless --convert-to pdf --outdir . render.xlsx >/dev/null 2>&1
pdftoppm -r 70 -png render.pdf "$2" && ls $2*
