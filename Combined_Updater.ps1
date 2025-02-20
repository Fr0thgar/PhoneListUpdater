# --------------------------------------
# CSV Cleanup Section
# --------------------------------------

# Define expected headers for the cleaned CSV (For documentation)
#$expectedHeaders = @("Fuldenavn", "Direkte", "Mobil", "Mail", "Stilling", "Afdeling", "Firma", "Lokation")

# Input CSV file (semicolon-delimited)
$inputCsvPath = "C:\Path\To\Your\Input.csv"

# Output folder and bas filename for cleaned CSV
$outputFolder = "C:\Path\To\Your\Outputfolder"
$baseName = "Output file name"
$extension = ".csv"

# Construct the initial output file path
$outputCsvPath = Join-Path $outputFolder ($baseName + $extension)

# If file exists, append a counter to create a unique filename
$counter = 1
while (Test-Path $outputCsvPath) {
    $outputCsvPath = Join-Path $outputFolder ("{0}_{1}{2}" -f $baseName, $counter, $extension)
    $counter++
}

# Import the CSV using semicolon as the delimiter 
$data = Import-Csv -Path $inputCsvPath -Delimiter ';'

# Process each row to create a cleaned Dataset
$cleanedData = foreach ($row in $data) {
    # Combine Fornavn and Efternavn into Fuldenavn, trimming any extra spaces
    $fuldenavn = "$($row.Fornavn) $($row.Efternavn)".Trim()

    [PSCustomObject]@{
        Fuldenavn = $fuldenavn
        Direkte   = $row.Lokalnummer      # Maps to 'Direkte'
        Mobil     = $row.Mobilnummer      # Maps to 'Mobil'
        Mail      = $row.Email            # Maps to 'Mail'
        Stilling  = $row.Stilling         # Maps to 'Stilling'
        Afdeling  = $row.Afdeling         # Maps to 'Afdeling'
        Firma     = ""                    # No corresponding input column; left blank
        Lokation  = $row.Adresse          # Maps to 'Lokation'
    }
}

# Export the cleaned data to a new CSV file with a unique filename
$cleanedData | Export-Csv -Path $outputCsvPath -NoTypeInformation -Encoding UTF8
Write-Host "CSV cleanup completed! Cleaned file save to $outputCsvPath" -ForegroundColor Green

# --------------------------------------
# PowerShell and PnP PowerShell Module checker 
# --------------------------------------

# Check PowerShell Version
if ($PSCersionTable.PSVersion.Major -lt 7) {
    Write-Host "Your PowerShell version is outdated! (Current: $($PSVersionTable.PSVersion))" -ForegroundColor Red
    Write-Host "Upgrading to the latest PowerShell version..." -ForegroundColor Yellow

    # Donload & Install the latest PowerShell
    $installerUrl = "https://aka.ms/install-powershell.ps1"
    $installerPath = "$env:TEMP\install-powershell.ps1"

    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath
    Write-Host "Running PowerShell installer..." -ForegroundColor Yellow
    & $installerPath -UseMSI

    Write-Host "PowerShell upgrade initiated. Please restart your Terminal and run the script again." -ForegroundColor Green
}

# Check if PnP.PowerShell is installed
$installedModule = Get-Module -Name PnP.PowerShell -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
# Get latest stable version from PowerShell Gallery
$latestVersion = Find-Module -Name -PnP.PowerShell | Select-Object -ExpandProperty Version

if ($installedModule) {
    $installedVersion = $installedModule.Version
    Write-Host "Installed PnP.PowerShell Version: $installedVersion"
    Write-Host "Latest PnP.PowerShell Version: $latestVersion"

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

# Import the PnP PowerShell module (ensure it is installed and updated as needed)
Import-Module PnP.PowerShell
Write-Host "PnP PowerShell module loaded successfully!" -ForegroundColor Green

# --------------------------------------
# SharePoint Phone List Updater Section
# --------------------------------------

# Variables for app registration
$siteUrl    = "https://yourTenant.sharepoint.com" # Ensure your URL is correct
$clientId   = "Your ClientID from Entra Admin Center" # Get ClientId from Entra Admin Center - App registration 
$tenantId   = "Your TenantID from Entra Admin Center" # Get TenantId from Entra Admin Center - App registration 

# Connect to SharePoint
Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant $tenantId -Interactive

# Check if the list exists
try {
    $list = Get-PnPList -Identity "Telefonliste"
    Write-Host "List found: $($list.Title)"
} catch {
    Write-Host "Error: The list 'Telefonliste' does not exist. Please Check the name and try again."
    exit # Exit the script ifi the list is not found 
}

# Path to the CSV File
 $csvPath = "Your patht o the CSV File"

 # Read tge CSV File 
 $telefonListe = Import-Csv -Path $csvPath

 # Loop through each row in the cscv and update the sharepoint list
 foreach ($row in $telefonListe) {
    try {
        # Check if the item already exists in the list
        $existingItem = Get-PnPListItem -List "Telefonliste" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($row.Name)</Value></Eq></Where></Query></View>"
    } catch {
        Write-Host "Error retrieving item: $_"
        Get-PnPException
        continue # Skip to the next interaction if there's an error
    }

    if ($existingItem) {
        try {
            # Update the existing item
            Set-PnPListItem -List "Telefonliste" -Identity $existingItem.Id -Values @{
                # Make sure you update these to match your column names
                "LOKALNR"       = $row.Direkte
                "DIREKTE"       = $row.Mobil
                "E_x002d_MAIL"  = $row.Mail
                "STILLING"      = $row.Stilling
                "ADRESSE"       = $row.Afdeling
                "POSTNR"        = $row.Firma
                "BY"            = $row.Lokation
            }
            Write-Host "Updated: $(row.Name)"
        } catch {
            Write-Host "Error updating item: $_"
            Get-PnPException
        }
    } else {
        try {
            # Add a new item
            Add-PnPListItem -List "Telefonliste" -Values @{
                # Make sure you update These to match your Column names 
                "Title"         = $row.Fuldenavn
                "LOKALNR"       = $row.Direkte
                "DIREKTE"       = $row.Mobil
                "E_x002d_MAIL"  = $row.Mail
                "STILLING"      = $row.Stilling
                "ADRESSE"       = $row.Afdeling
                "POSTNR"        = $row.Firma
                "BY"            = $row.Lokation
            }
            Write-Host "Added: $($row.Name)"
        } catch {
            Write-Host "Error updating item: $_"
            Get-PnPException
        }
    }
}

Write-Host "Phone list update completed!"
