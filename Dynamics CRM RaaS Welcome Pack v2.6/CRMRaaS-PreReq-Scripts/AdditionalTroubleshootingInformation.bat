echo Workstation statistics and uptime -o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o- >> AdditionalTroubleshootingInformation.txt
net statistics workstation >> AdditionalTroubleshootingInformation.txt
echo lmhosts file -o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o- >> AdditionalTroubleshootingInformation.txt
type %windir%\System32\drivers\etc\lmhosts >> AdditionalTroubleshootingInformation.txt
echo hosts file -o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o- >> AdditionalTroubleshootingInformation.txt
type %windir%\System32\drivers\etc\hosts >> AdditionalTroubleshootingInformation.txt
echo route print -o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o- >> AdditionalTroubleshootingInformation.txt
route print >> AdditionalTroubleshootingInformation.txt
echo Tcpip Parameters -o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o- >> AdditionalTroubleshootingInformation.txt
REG QUERY HKLM\System\CurrentControlSet\Services\Tcpip\Parameters >> AdditionalTroubleshootingInformation.txt
echo ipconfig /all -o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o- >> AdditionalTroubleshootingInformation.txt
ipconfig /all >> AdditionalTroubleshootingInformation.txt
echo DNS Cache -o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o- >> AdditionalTroubleshootingInformation.txt
ipconfig /displaydns >> AdditionalTroubleshootingInformation.txt
echo nbtstat -n -o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o- >> AdditionalTroubleshootingInformation.txt
netstat -n >> AdditionalTroubleshootingInformation.txt
echo nbtstat -c -o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o- >> AdditionalTroubleshootingInformation.txt
nbtstat -c >> AdditionalTroubleshootingInformation.txt
echo Malware detection Pre Win10 -o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o- >> AdditionalTroubleshootingInformation.txt
wevtutil qe System /rd:true /f:text /q:*[System[(EventID=1116)]] >> AdditionalTroubleshootingInformation.txt
wevtutil qe System /rd:true /f:text /q:*[System[(EventID=1117)]] >> AdditionalTroubleshootingInformation.txt
wevtutil qe "Microsoft-Windows-Windows Defender/Operational" /rd:true /f:text /q:*[System[(EventID=1116)]] >> AdditionalTroubleshootingInformation.txt
wevtutil qe "Microsoft-Windows-Windows Defender/Operational" /rd:true /f:text /q:*[System[(EventID=1117)]] >> AdditionalTroubleshootingInformation.txt
echo Malware detection Win10+ -o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o- >> AdditionalTroubleshootingInformation.txt
powershell -file RaaSWindowsDefender.ps1  >> AdditionalTroubleshootingInformation.txt
echo Processes and Modules -o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o- >> AdditionalTroubleshootingInformation.txt
TASKLIST /M >> AdditionalTroubleshootingInformation.txt
echo sc queryex w32time  -o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o- >> AdditionalTroubleshootingInformation.txt
sc queryex w32time >> AdditionalTroubleshootingInformation.txt
echo W32tm /query /status -o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o- >> AdditionalTroubleshootingInformation.txt
W32tm /query /status >> AdditionalTroubleshootingInformation.txt
echo W32tm /query /configuration  -o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o- >> AdditionalTroubleshootingInformation.txt
W32tm /query /configuration >> AdditionalTroubleshootingInformation.txt
echo type "%programData%\VMware\VMware Tools\tools.conf"  -o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o- >> AdditionalTroubleshootingInformation.txt
type "%programData%\VMware\VMware Tools\tools.conf"  >> AdditionalTroubleshootingInformation.txt