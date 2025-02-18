# Check current modules
#Get-Module -ListAvailable

# Update PnP PowerShell module
#Update-Module -Name PnP.PowerShell

# If issues persist, uninstall and reinstall
#Uninstall-Module -Name PnP.PowerShell
#Install-Module -Name PnP.PowerShell

# Check PowerShell version
#$PSVersionTable.PSVersion

# Import the PnP Powershell module
Import-Module PnP.PowerShell

# Variables for app registration
$siteUrl = "https://twcas.sharepoint.com/HR"  # Ensure this is correct
$clientId = "d704eda0-54af-4fad-a62a-546401569152"
$tenantId = "38898116-fa71-4dd1-b72d-09c5fd8a7141"


# Connect to Sharepoint
Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant $tenantId -Interactive

# Test the connection
#try {
 #   $web = Get-PnPWeb
  #  Write-Host "Connected to SharePoint site: $(web.Title)"
#} catch {
 #   Write-Host "Failed to connect to SharePoint: $_"    
#}

# Check if the list exists
try {
    $list = Get-PnPList -Identity "Telefonliste"
    Write-Host "List found: $($list.Title)"
} catch {
    Write-Host "Error: The list 'Telefonliste' does not exist. Please check the name and try again."
    exit # Exit the script if the list is not found
}

# Test accessing the mobil column
try {
    $list = Get-PnPListItem -List "Telefonliste" -Id 1 
    Write-Host "Mobil: $($item['Mobil'])"
}
catch {
    Write-Host "Error accessing Mobil column: $_"
}

# Path to the CSV File
$csvPath = "\\filsrv\it\Telefonliste\Telefonliste.csv"

# Read the CSV File
$phoneList = Import-Csv -Path $csvPath

# Loop through each row in the csv and update the sharepoint list
foreach($row in $phoneList) {
    try {
        #Check if the item already exists in the list
        $existingItem = Get-PnPListItem -List "Telefonliste" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($row.Name)</Value></Eq></Where></Query></View>"
    } catch {
        Write-Host "Error retrieving item: $_"
        Get-PnPException
        continue # Skip to the next interation if there's an error
    }
    
    if ($existingItem){
        try{
            # Update the existing item
            Set-PnPListItem -List "Telefonliste" -Identity $existingItem.Id -Values @{
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
            Add-PnPListItem -List "Telefonliste" -Values @{
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