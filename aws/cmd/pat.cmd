echo PAT_START %date% %time% >> C:\aws\log\pat_log.txt 2>>&1
rem ******************************************
rem �@�p�g���C�g����F�_��ON�E����ON
rem ******************************************

rem �J���p
rem start /MIN C:\cygwin64\bin\rsh.exe 192.168.1.1 -l pat_user alert 299993

rem rsh�����󂯕t���Ă���Ȃ�
rem *** ���ėpPC�iMGT-PAT-001�j
start /MIN C:\cygwin64\bin\rsh.exe xxx.xxx.xxx.xxx(IP���w��j  -l pat_user alert 299993

C:\Windows\System32\timeout.exe /T 2
