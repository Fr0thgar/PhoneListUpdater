# Check current modules
Get-Module -ListAvailable

# Update PnP PowerShell module
Update-Module -Name PnP.PowerShell

# If issues persist, uninstall and reinstall
#Uninstall-Module -Name PnP.PowerShell
#Install-Module -Name PnP.PowerShell

# Check PowerShell version
$PSVersionTable.PSVersion

# Import the PnP Powershell module
Import-Module PnP.PowerShell

# Connect to Sharepoint
$siteUrl = "https://twcas.sharepoint.com/HR/Lists/Telefonliste/Allitemsg.aspx"
Connect-PnPOnline -Url $siteUrl -Interactive

# Path to the CSV File
$csvPath = "\\filsrv\it\Phonelist\upload_template.csv"

# Read the CSV File
$phoneList = Import-Csv -Path $csvPath

# Loop through each row in the csv and update the sharepoint list
foreach($row in $phoneList) {
    try {
        #Check if the item already exists in the list
        $existingItem = Get-PnPListItem -List "PhoneList" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($row.Name)</Value></Eq></Where></Query></View>"
    } catch {
        Write-Host "Error retrieving item: $_"
        Get-PnPException
        continue # Skip to the next interation if there's an error
    }
    
    if ($existingItem){
        try{
            # Update the existing item
            Set-PnPListItem -List "PhoneLIst" -Identity $existingItem.Id -Values @{
                "Direkte"   = $row.Direkte
                "Mobil"     = $row.Mobil
                "Mail"      = $row.Mail
                "Stilling"  = $row.Stilling
                "Afdeling"  = $row.Afdeling
                "Firma"     = $row.Firma
                "Lokation"  = $row.Lokation
            }
            Write-Host "Updated: $($row.Name)"
        } catch {
            Write-Host "Error updating item: $_"
            Get-PnPException
        }       
    } else {
        try {
            #  Add a new item
            Add-PnPListItem -List "PhoneList" -Values @{
                "Title"     = $row.Name
                "Direkte"   = $row.Direkte
                "Mobil"     = $row.Mobil
                "Mail"      = $row.Mail
                "Stilling"  = $row.Stilling
                "Afdeling"  = $row.Afdeling
                "Firma"     = $row.Firma
                "Lokation"  = $row.Lokation
            }
            Write-Host "Added: $($row.Name)"
        } catch {
            Write-Host "Error updating item: $_"
            Get-PnPException
        }
    }
}

Write-Host "Phone list update completed!"