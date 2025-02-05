# Define file paths
$publisherFile = "C:\Users\USER\Desktop\SCCM Data\SCCM_nonprod\Cleaned\publisher.xlsx"
$sourceDirectory = "C:\Users\USER\Desktop\SCCM Data\SCCM_nonprod\Cleaned"
$outputFile = "C:\Users\USER\Desktop\SCCM Data\SCCM_nonprod\Cleaned\nonmatched_output_nonprod.xlsx"

# Import necessary module
Import-Module ImportExcel

# Read publisher names from publisher.xlsx and store in an array
$publisherNames = @()
Import-Excel -Path $publisherFile | ForEach-Object { 
    if ($_.publisher) { $publisherNames += [regex]::Escape($_.publisher.Trim()) }  # Escape special characters for regex
}

# Initialize an array to store unmatched output data
$unmatchedData = @()

# GUID pattern to exclude (matches {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx})
$guidPattern = "^{[0-9A-Fa-f\-]{36}}$"

# Process each Excel file in the source directory, excluding publisher.xlsx
Get-ChildItem -Path $sourceDirectory -Filter "*.xlsx" | Where-Object { $_.Name -ne "publisher.xlsx" } | 
ForEach-Object {
    $filePath = $_.FullName
    $fileData = Import-Excel -Path $filePath

    # Ensure fileData is not empty before proceeding
    if ($fileData) {
        # Extract OS entry (Assumes there's exactly **one** row with "Operating System")
        $osEntry = ($fileData | Where-Object { $_."Product Category" -like "*Operating System*" } | Select-Object -First 1)."Product Name"

        foreach ($row in $fileData) {
            $publisherNameInFile = $row."publisher"  # Assuming "publisher" is the column name in your files

            # Ensure publisherNameInFile is valid before lookup and exclude GUIDs
            if (![string]::IsNullOrWhiteSpace($publisherNameInFile) -and $publisherNameInFile -notmatch $guidPattern) {
                $isMatched = $false

                foreach ($publisherName in $publisherNames) {
                    if ($publisherNameInFile -match "\b$publisherName\b") {  # Match even if part of a larger name
                        $isMatched = $true
                        break  # Exit loop early since a match was found
                    }
                }

                # If no match was found, add the row to unmatched data
                if (-not $isMatched) {
                    $unmatchedData += [PSCustomObject]@{
                        "Publisher Name" = $publisherNameInFile
                        "Product Name"   = $row."Product Name"   # Keep Product Name in the output
                        "Computer Name"  = $row."Computer Name"
                        "OS"             = $osEntry  # Use the same OS for all rows in the file
                    }
                }
            }
        }
    }
}

# Export the unmatched data to an Excel file
if ($unmatchedData.Count -gt 0) {
    $unmatchedData | Export-Excel -Path $outputFile -AutoSize
    Write-Output "Script completed. Unmatched entries saved to $outputFile"
} else {
    Write-Output "All entries were matched. No unmatched data found."
}
