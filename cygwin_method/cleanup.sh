#!/bin/bash
set -e
cd /cygdrive/c/Users/efairchild/Documents/Development/AMM_sgml_excel_conversion/;
echo ; 
echo "Converting AMM...";
sed -f amm1.sed AMM_AIL_A320_VRD.SGM > amm1.txt;
egrep '<TASK|<SUBTASK' amm1.txt > amm2.txt;
exit