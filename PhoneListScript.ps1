Write-Host "If you get an Access Denied error and you are sure you have access on the account chosen, Close you terminal and run it fresh. This should let you re-authenticate with your account." -ForegroundColor Red

# Check Powershell Version
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host " Your PowerShell version is outdated! (Current:  $($PSVersionTable.PSVersion))" -ForegroundColor Red
    Write-Host "Upgrading to the latest PowerShell version..." -ForegroundColor Yellow

    # Downlaod & Install the latest Powershell
    $installerUrl = "https://aka.ms/install-powershell.ps1"
    $installerPath = "$env:TEMP\install-powershell.ps1"

    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath
    Write-Host "Running PowerShell installer..." -ForegroundColor Yellow
    & $installerPath -UseMSI

    Write-Host "PowerShell upgrade initiated. Please restart your terminal and run the script again." -ForegroundColor Green
    exit # Stop script execution
} else {
    Write-Host "PowerShell version is up to date: $($PSVersionTable.PSVersion)" -ForegroundColor Green
}

# Check if PnP.PowerShell is installed 
$installedModule = Get-Module -Name PnP.PowerShell -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
# Get latest stable  version from PowerShell Gallery
$latestVersion = Find-Module -Name PnP.PowerShell | Select-Object -ExpandProperty Version

if($installedModule) {
    $installedVersion = $installedModule.Version
    Write-Host "Installed PnP.PowerShell Version: $installedVersion"
    Write-Host "Latest PnP.PowerShell version: $latestVersion"

    if ($installedVersion -lt $latestVersion) {
        Write-Host "Updating PnP.PowerShell to the latest version..."
        Update-Module -Name PnP.PowerShell -Force
    } else {
        Write-Host "PnP.PowerShell is already up to date!"
    }
} else {
    Write-Host "PnP.PowerShell is not installed. Installing the latest version..."
    Install-Module -Name PnP.PowerShell -Force
}

# Import the PnP Powershell module
Import-Module PnP.PowerShell
Write-Host "PnP PowerShell module loaded successfully!" -ForegroundColor Green

# Variables for app registration
$siteUrl = "https://Yourtenant.sharepoint.com"  # Ensure this is correct
$clientId = "Your ClientID from Entra Admin Center" # Get ClientId from Entra Admin Center - App registration 
$tenantId = "Your TenantID from Entra Admin Center" # Get TenantId from Entra Admin Center - App registration

# Connect to Sharepoint
Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant $tenantId -Interactive

# Check if the list exists
try {
    $list = Get-PnPList -Identity "Telefonliste"
    Write-Host "List found: $($list.Title)"
} catch {
    Write-Host "Error: The list 'Telefonliste' does not exist. Please check the name and try again."
    exit # Exit the script if the list is not found
}

# Path to the CSV File
$csvPath = "Your Path to the CSV file"

# Read the CSV File
$telefonListe = Import-Csv -Path $csvPath

# Loop through each row in the csv and update the sharepoint list
foreach($row in $telefonListe) {
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
                # Make sure you update These to match your Column names 
                "LOKALNR"   = $row.Direkte
                "DIREKTE"     = $row.Mobil
                "E_x002d_MAIL"      = $row.Mail
                "STILLING"  = $row.Stilling
                "ADRESSE"  = $row.Afdeling
                "POSTNR"     = $row.Firma
                "BY"  = $row.Lokation
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
                # Make sure you update These to match your Column names 
                "Title" = $row.Fuldenavn
                "LOKALNR"   = $row.Direkte
                "DIREKTE"     = $row.Mobil
                "E_x002d_MAIL"      = $row.Mail
                "STILLING"  = $row.Stilling
                "ADRESSE"  = $row.Afdeling
                "POSTNR"     = $row.Firma
                "BY"  = $row.Lokation
            }
            Write-Host "Added: $($row.Name)"
        } catch {
            Write-Host "Error updating item: $_"
            Get-PnPException
        }
    }
}

Write-Host "Phone list update completed!"