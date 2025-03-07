# Define the correct headers (optinal, for documentation)
#$expectedHeaders = @("Fuldenavn", "Direkte", "Mobil", "Mail", "Stilling", "Afdeling", "Firma", "Lokation")

# Path to input and output CSV files (adjust path as necessary)
$inputCsvPath   = "Your input Path here"
$outputCsvPath  = "Your output Path here"
$baseName       = "Your output filename"
$extension      = ".csv" 

# Construct the initail output file path
$outputCsvPath = Join-Path $outputFolder ($baseName + $extension)

# if file exists, append a counter to create a unique  filename
$counter = 1
while (Test-Path $outputCsvPath) {
    $outputCsvPath = Join-Path $outputFolder ("{0}_{1}{2}" -f $baseName, $counter, $extension)
    $counter++
}

# Import the CSV
$data = Import-Csv -Path $inputCsvPath -Delimiter ';'

# Process each row to map the columns as needed
$cleanedData = foreach ($row in $data) {
    # Combine Fornavn and Efternavn for the full name, trimming any extra spaces
    $fuldenavn = "$($row.Fornavn) $($row.Efternavn)".Trim()

    # Create a new object with the expected columns
    [PSCustomObject]@{
        Fuldenavn = $fuldenavn
        Direkte = $row.Lokalnummer  # Maps to 'Direkte'
        Mobil = $row.Mobilnummer    # maps to 'Mobil'
        Mail =  $row.Email          # maps to 'Mail'
        Stilling = $row.Stilling    # maps to 'Stilling'
        Afdeling = $row.Afdeling    # Maps to 'afdeling'
        Firma = ""                  # no corresponding input column; left blank
        Lokation = $row.Adresse     # maps to 'Lokation'
    }
}

# Export the cleaned data to a new CSV with the correct headers
$cleanedData | Export-Csv -Path $outputCsvPath -NoTypeInformation -Encoding UTF8

Write-Host "CSV cleanup completed! Cleaned file saved to $outputCsvPath"
