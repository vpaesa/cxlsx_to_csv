#!/bin/bash

for i in ??_*_??.xlsx
do
  testname=${i%??.xlsx}
  sheets=${i: -7:2}
  for sheetid in $(seq  -f '%02.0f' 1 $sheets)
  do
    #echo ${testname}$sheetid
    ../cxlsx_to_csv -if $i -sh $sheetid -of validating_${testname}$sheetid.csv
    ./csvtotab expected_${testname}$sheetid.csv > expected_${testname}$sheetid.tab
    ./csvtotab validating_${testname}$sheetid.csv > validating_${testname}$sheetid.tab
    cmp expected_${testname}$sheetid.tab validating_${testname}$sheetid.tab
    if [ $? -eq 0 ]
    then echo "Passed ${testname}$sheetid"
    else echo "Failed ${testname}$sheetid"
    fi
  done  
done
