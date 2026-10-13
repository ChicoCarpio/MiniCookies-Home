#!/bin/sh
# Suma de |dAncho| y |dIzq|+|dDer| de tinta-por-celda en varios casos (para elegir parametros de texto). Uso: resumen-tinta.sh
cd "$(dirname "$0")/.."
for spec in "01-estilos-fuentes 23-openpyxl-dibujo.xlsx 2 A1:B10" "03-geometria 23-openpyxl-dibujo.xlsx 1 A1:L18" "05-alineacion 23-openpyxl-dibujo.xlsx 4 A1:H33"; do
  set -- $spec
  node comparar-xlsx/tinta-por-celda.mjs comparar-xlsx/corpus/$2 $3 $4 comparar-xlsx/medidas/tmp/dibujo/salida/$1.png | grep -v "rot\|^E7\|^E10\|^E13\|^E16\|^G2\|h=distributed\|h=fill" | awk -v n="$1" '{for(i=1;i<=NF;i++){if($i=="dAncho")a+=($(i+1)<0?-$(i+1):$(i+1)); if($i=="dIzq")b+=($(i+1)<0?-$(i+1):$(i+1)); if($i=="dDer")c+=($(i+1)<0?-$(i+1):$(i+1))} k++} END {printf "%s: celdas %d  sum|dAncho| %d  sum|dIzq| %d  sum|dDer| %d\n", n,k,a,b,c}'
done
