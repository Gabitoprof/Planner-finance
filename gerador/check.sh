#!/bin/bash
# gera, copia para a pasta de testes e recalcula
cd /home/user/Planner-finance/gerador && python3 build.py && mkdir -p /home/user/Planner-finance/.teste && cp ../Planilha_Financeira_Avancada.xlsx ../.teste/t.xlsx
SK=$(ls -d /root/.claude/skills/synced/*/xlsx)
cd /home/user/Planner-finance/.teste && python3 $SK/scripts/recalc.py t.xlsx 580 2>/dev/null
