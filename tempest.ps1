# Define file paths
$publisherFile = "C:\Users\HP\Desktop\SCCM Data\Cleaned\publisher.xlsx"
$sourceDirectory = "C:\Users\HP\Desktop\SCCM Data\Cleaned"
$outputFile = "C:\Users\HP\Desktop\SCCM Data\Cleaned\outputfile.xlsx"

# Import necessary module
Import-Module ImportExcel

# Read publisher names from publisher.xlsx and store in an array
$publisherNames = @()
Import-Excel -Path $publisherFile | ForEach-Object { 
    if ($_.publisher) { $publisherNames += [regex]::Escape($_.publisher.Trim()) }  # Escape special characters for regex
}

# Initialize an array to store output data
$outputData = @()

# Process each Excel file in the source directory, excluding publisher.xlsx
Get-ChildItem -Path $sourceDirectory -Filter "*.xlsx" | Where-Object { $_.Name -ne "publisher.xlsx" } | 
ForEach-Object {
    $filePath = $_.FullName
    $fileData = Import-Excel -Path $filePath

    # Ensure fileData is not empty before proceeding
    if ($fileData) {
        foreach ($row in $fileData) {
            $publisherNameInFile = $row."publisher"  # Assuming "publisher" is the column name in your files

            # Ensure publisherNameInFile is valid before lookup
            if (![string]::IsNullOrWhiteSpace($publisherNameInFile)) {
                $matchFound = $false

                # Check if publisher matches any from the publisher.xlsx file
                foreach ($publisherName in $publisherNames) {
                    if ($publisherNameInFile -match "\b$publisherName\b") {
                        $matchFound = $true
                        break  # Stop checking if a match is found
                    }
                }

                # If NO match is found, include this entry
                if (-not $matchFound) {
                    $outputData += [PSCustomObject]@{
                        "Publisher Name" = $row."publisher"
                        "Product Name"   = $row."Product Name"
                        "Computer Name"  = $row."Computer Name"
                        "OS"             = $row."Product Category"
                    }
                }
            }
        }
    }
}

# Export the output data to an Excel file
if ($outputData.Count -gt 0) {
    $outputData | Export-Excel -Path $outputFile -AutoSize
    Write-Output "Script completed. Output saved to $outputFile"
} else {
    Write-Output "No unmatched data found."
}
