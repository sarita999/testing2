#PowerShell script to Download files from SharePoint Online Document Library folder to network path  
#Created By: sharmila.konreddy 
# Variable - Change the parameter as it need 
#Note:Once the dedicated Host account is created by infra team please replace the $O365ServiceAccount ,$O365ServiceAccountPwd variable with those information
param( $SharePointSiteURL,$newSharePointFolderPath)
$O365ServiceAccount = "sharmila.konreddy@accenture.com"  
$O365ServiceAccountPwd = "Lakki@426"  

# Change this SharePoint Site URL 
#$SharePointSiteURL = "https://ts.accenture.com/sites/MyPersonal422" 
#folder name for folder creation if Scrubbed Assets folder is not available in the specified sharepoint URL
$newFolderName="Scrubbed Assets"
 
#change it with Download folder path in .exe file location
$localFolderPath = "C:\Users\sharmila.konreddy\Desktop\Docs list"   

# Change this Network Folder path  
#$newSharePointFolderPath = "Shared Documents/General/R2/Performance Checker"
$SharePointFolderPath = "Shared Documents/General/R2/Performance Checker/Scrubbed Assets" 
#$SharePointFolderPath = $newSharePointFolderPath+$newFolderName
 

# Change the Document Library and Folder path  
#Ends[SecureString] $SecurePass = ConvertTo - SecureString $O365ServiceAccountPwd - AsPlainText - Force[System.Management.Automation.PSCredential] $PSCredentials = New - Object System.Management.Automation.PSCredential($O365ServiceAccount, $SecurePass)  

#Connecting to SharePoint site
Connect-PnpOnline -Url $SharePointSiteURL -UseWebLogin

#fetching all the folder under clientname folder from sharepoint
$allItems = Get-PnPListItem -List $newSharePointFolderPath -Fields "Title", "test", "FileDirRef"
$folderExist=$false
foreach ($item in $allItems) 
{
   #checking whethere scrubbed Assets folder is exist or not if not it will lead to creation folder or else will skip the folder creation
   if (($item.FileSystemObjectType) -eq "Folder" -and $item["Title"] -eq $newFolderName)  
	{
		$folderExist=$true
		break
	}
}

if(!$folderExist)
{
	Add-PnPFolder -Name $newFolderName -Folder $newSharePointFolderPath 
}
#fetching all files from local folder
$files = Get-ChildItem $localFolderPath 

foreach ($f in $files)
{
	#adding the each file from  local folder to sharepoint specified folder
	$status=Add-PnPFile -Path $f.FullName -Folder $SharePointFolderPath -ErrorAction SilentlyContinue
	if($status)
	{ 
		Write-Host "Successful Upload"
	}
	else
	{ 
		Write-Host "unSuccessful Upload"
	}

}