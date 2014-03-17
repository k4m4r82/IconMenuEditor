cls
copy SSubTmr6.dll %systemroot%\system32
copy cPopMenu6.ocx %systemroot%\system32

regsvr32 /s %systemroot%\system32\SSubTmr6.dll
regsvr32 /s %systemroot%\system32\cPopMenu6.ocx