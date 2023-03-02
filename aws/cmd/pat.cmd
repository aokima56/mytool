echo PAT_START %date% %time% >> C:\aws\log\pat_log.txt 2>>&1
rem ******************************************
rem 　パトライト操作：点灯ON・共鳴ON
rem ******************************************

rem 開発用
rem start /MIN C:\cygwin64\bin\rsh.exe 192.168.1.1 -l pat_user alert 299993

rem rshしか受け付けてくれない
rem *** 発呼用PC（MGT-PAT-001）
start /MIN C:\cygwin64\bin\rsh.exe xxx.xxx.xxx.xxx(IPを指定）  -l pat_user alert 299993

C:\Windows\System32\timeout.exe /T 2
