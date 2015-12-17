export LC_ALL=C.UTF-8
export LD_LIBRARY_PATH=/lib:/usr/lib:/usr/local/lib

#default XML library is Expat (most known and fastest)
cxlsx_to_csv: cxlsx_to_csv.c miniz.c
	cc -DCONFIG_EXPAT -march=native -O3 -o cxlsx_to_csv cxlsx_to_csv.c -l expat && strip cxlsx_to_csv

cxlsx_to_csv_expat: cxlsx_to_csv.c miniz.c
	cc -DCONFIG_EXPAT -march=native -O3 -o cxlsx_to_csv_expat cxlsx_to_csv.c -l expat && strip cxlsx_to_csv_expat

cxlsx_to_csv_mxml: cxlsx_to_csv.c miniz.c
	cc -DCONFIG_MXML  -march=native -O3 -D_THREAD_SAFE -D_REENTRANT -I/usr/local/include -o cxlsx_to_csv_mxml cxlsx_to_csv.c -L/usr/local/lib -l mxml -lpthread && strip cxlsx_to_csv_mxml

cxlsx_to_csv_parsifal: cxlsx_to_csv.c miniz.c
	cc -DCONFIG_PARSIFAL -march=native -O3 -o cxlsx_to_csv_parsifal -I/usr/local/include cxlsx_to_csv.c -L/usr/local/lib/ -lparsifal && strip cxlsx_to_csv_parsifal

cxlsx_to_csv_noxml: cxlsx_to_csv.c miniz.c
	cc -march=native -O3 -o cxlsx_to_csv_noxml cxlsx_to_csv.c && strip cxlsx_to_csv_noxml

alllibs: cxlsx_to_csv_expat cxlsx_to_csv_mxml cxlsx_to_csv_parsifal cxlsx_to_csv_noxml

win32: cxlsx_to_csv.c miniz.c
	i686-w64-mingw32-gcc -DCONFIG_EXPAT -O3 -o cxlsx_to_csv32.exe cxlsx_to_csv.c -l expat && i686-w64-mingw32-strip cxlsx_to_csv32.exe

win64: cxlsx_to_csv.c miniz.c
	x86_64-w64-mingw32-gcc -DCONFIG_EXPAT -O3 -o cxlsx_to_csv64.exe cxlsx_to_csv.c -l expat && x86_64-w64-mingw32-strip cxlsx_to_csv64.exe

test/csvtotab.c:
	wget 'http://dev.w3.org/cvsweb/csvtotab-vv/csvtotab.c?rev=1.1;content-type=text%2Fplain' -O test/csvtotab.c

test/csvtotab: test/csvtotab.c
	cc -o test/csvtotab test/csvtotab.c

test: cxlsx_to_csv test/csvtotab
	cd test && ./00_runtest.sh
