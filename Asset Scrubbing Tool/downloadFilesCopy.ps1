#PowerShell script to Download files from SharePoint Online Document Library folder to network path  
#Created By: sharmila.konreddy 
# Variable - Change the parameter as it need 
param ($SharePointSiteURL,$SharePointFolderPath)
$O365ServiceAccount = "sharmila.konreddy@accenture.com"  
$O365ServiceAccountPwd = "Lakki@426" 

#$SharePointSiteURL = "https://ts.accenture.com/sites/MyPersonal422" 
 
# Change this SharePoint Site URL  
$SharedDriveFolderPath = "C:\Users\sharmila.konreddy\Desktop\Asset-Scrubbing-Web-Page-AssetScrubbingTool\TestData"  

# Change this Network Folder path  
#$SharePointFolderPath = "Shared Documents/General/R2/Performance Checker"  

# Change the Document Library and Folder path  
#Ends[SecureString] $SecurePass = ConvertTo - SecureString $O365ServiceAccountPwd - AsPlainText - Force[System.Management.Automation.PSCredential] $PSCredentials = New - Object System.Management.Automation.PSCredential($O365ServiceAccount, $SecurePass)  

#Connecting to SharePoint site  
#Connect-PnPOnline -Url $SharePointSiteURL -Credentials $PSCredentials  
Connect-PnpOnline -Url $SharePointSiteURL -UseWebLogin
$Files = Get-PnPFolderItem -FolderSiteRelativeUrl $SharePointFolderPath -ItemType File  

foreach($File in $Files) {  
    Get-PnPFile -Url $File.ServerRelativeUrl -Path $SharedDriveFolderPath -FileName $File.Name -AsFile  
}   