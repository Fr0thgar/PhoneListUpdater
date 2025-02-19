# Define the correct headers
$expectedHeaders = @("Fuldenavn", "Direkte", "Mobil", "Mail", "Stilling", "Afdeling", "Firma", "Lokation")

# Path to input and output CSV files
$inputCsvPath = "Path here"
$outputCsvPath = "Path here"

# Import the CSV
$data = Import-Csv -Path $inputCsvPath

# Check and rename headers if needed
$headers = $data[0].PSObject.Properties.Name

# Map incorrect headers to expected ones if necessary (manually adjust if needed)
$headerMappings = @{
    ""  = "Fuldenavn"
    ""      = "Direkte"
    ""     = "Mobil"
    ""     = "Mail"
    ""  = "Stilling"
    "" = "Afdeling"
    ""    = "Firma"
    ""   = "Lokation"
}

# Create a new cleaned dataset with standardized headers
$cleanedData = @{
     $newRow = @{}

     foreach ($expectedHeader in $expectedHeaders) {
        $matchingColumn = $headers -match $expectedHeader
        if ($matchingColumn) {
            $newRow[$expectedHeader] = $row.$matchingColumn
        }
        elseif ($headerMappings.ContainsKey($expectedHeader)) {
            $mappedColumn = $headerMappings[$expectedHeader]
            if ($headers -contains $mappedColumn) {
                $newRow[$expectedHeader] = $row.$mappedColumn
            }
        }
        else {
            $newRow[$expectedHeader] = " # fill missing values with empty strings"
        }
    }

    $cleanedData += New-Object PSObject -Property $newRow
}

# Export the cleaned data
$cleanedData | Export-Csv -Path $outputCsvPath -NoTypeInformation -Encoding UTF8

Write-Host "CSV cleanup completed! Cleaned file saved to $outputCsvPath"