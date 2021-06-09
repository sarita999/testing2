import subprocess,os
sharePointUrl="https://ts.accenture.com/sites/MyPersonal422"
sharePointFolderPath="Shared Documents/General/R2/Performance Checker"

cmd = ['powershell.exe', "-ExecutionPolicy", "Unrestricted", "-File", "./downloadFilesCopy.ps1"]
ec = subprocess.call(cmd)
print("Powershell returned: {0:d}".format(ec))
psxmlgen = subprocess.Popen([r'C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe',
                             '-ExecutionPolicy',
                             'Unrestricted',
                             './downloadFilesCopy.ps1',
                             '-SharePointSiteURL:',sharePointUrl,'-SharePointFolderPath:',sharePointFolderPath], cwd=os.getcwd())
s1=subprocess.run([r'C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe', './downloadFilesCopy.ps1  https://ts.accenture.com/sites/MyPersonal422 Shared Documents/General/R2/Performance Checker'])
cmd='./downloadFilesCopy.ps1  https://ts.accenture.com/sites/MyPersonal422 Shared Documents/General/R2/Performance Checker'
s=subprocess.run([r'C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe', "-Command", cmd], capture_output=True)


print(s)