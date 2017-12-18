rem cd \d :C\Program Files (x86)\WinSCP

winscp /console /command "option batch continue" /script=.\log_file.txt
winscp /console /command "option batch continue" /script=.\psftp.bat

plink.exe -ssh -v -pw financialsys financialsys@10.21.2.10 -m .\crt.bat
putty.exe -l financialsys -pw financialsys 10.21.2.10

