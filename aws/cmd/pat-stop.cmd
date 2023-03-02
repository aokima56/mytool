echo PAT_START %date% %time% >> C:\log\pat_log.txt 2>>&1
rem **********************************
rem 　点灯変化OFF・共鳴OFF
rem **********************************


rem データセンター\
start /MIN C:\cygwin64\bin\rsh.exe 192.168.1.1 -l pat_user alert 099990