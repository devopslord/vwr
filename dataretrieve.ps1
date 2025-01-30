# Define file paths
$publisherFile = "C:\Users\USER\Desktop\SCCM Data\Cleaned\publisher.xlsx"
$sourceDirectory = "C:\Users\USER\Desktop\SCCM Data\Cleaned"
$outputFile = "C:\Users\USER\Desktop\SCCM Data\Cleaned\outputfile.xlsx"

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
                foreach ($publisherName in $publisherNames) {
                    if ($publisherNameInFile -match "\b$publisherName\b") {  # Match even if part of a larger name
                        # Find corresponding OS row
                        $osRow = $fileData | Where-Object { $_."Product Category" -like "*Operating System*" }

                        if ($osRow) {
                            # Add the matched data, including Product Name, to the output array
                            $outputData += [PSCustomObject]@{
                                "Publisher Name" = $publisherNameInFile
                                "Product Name"   = $row."Product Name"   # Capture the Product Name for the match
                                "Computer Name"  = $osRow."Computer Name"
                                "OS"             = $osRow."Product Name"
                            }
                        }
                        break  # No need to check further if already matched
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
    Write-Output "No matching data found."
}
