#!/bin/bash
# Gera as duas planilhas, recalcula no LibreOffice, injeta o cache de valores e verifica erros.
set -e
cd /home/user/Planner-finance/gerador
SK=$(ls -d /root/.claude/skills/synced/*/xlsx)
mkdir -p ../.teste
for modo in exemplo limpa; do
  if [ "$modo" = "exemplo" ]; then OUT=Planilha_Financeira_Avancada.xlsx; python3 build.py; else OUT=Planilha_Financeira_Avancada_LIMPA.xlsx; python3 build.py --limpa; fi
  cp ../$OUT ../.teste/calc.xlsx
  (cd ../.teste && python3 $SK/scripts/recalc.py calc.xlsx 580 2>/dev/null | head -8)
  python3 inject_cache.py ../$OUT ../.teste/calc.xlsx ../.teste/final_$OUT 2>/dev/null
  cp ../.teste/final_$OUT ../$OUT
done
ls -la ../*.xlsx
