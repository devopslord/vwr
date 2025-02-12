# Ensure ImportExcel module is installed
# Install-Module -Name ImportExcel -Scope CurrentUser

# Define file paths
$book1Path = "C:\Users\USER\Desktop\New_Excel\Azuredisks_prod.xlsx"       # Source for Resource Group
$book2Path = "C:\Users\USER\Desktop\New_Excel\CombinedResult1.xlsx"       # Target for updates
$outputPath = "C:\Users\USER\Desktop\New_Excel\CombinedResult2.xlsx"     # Final output file

# Import Excel files
$book1 = Import-Excel -Path $book1Path
$book2 = Import-Excel -Path $book2Path

# Check if 'Resource Group', 'osdisk', and 'datadisk' columns exist, and add them if missing
if (-not ($book2[0].PSObject.Properties.Name -contains "Resource Group")) {
    Write-Host "'Resource Group' column not found. Adding column..."
    $book2 | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name "Resource Group" -Value "" }
}
if (-not ($book2[0].PSObject.Properties.Name -contains "osdisk")) {
    Write-Host "'osdisk' column not found. Adding column..."
    $book2 | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name "osdisk" -Value 0 }
}
if (-not ($book2[0].PSObject.Properties.Name -contains "datadisk")) {
    Write-Host "'datadisk' column not found. Adding column..."
    $book2 | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name "datadisk" -Value 0 }
}

# Resolve Value Function: Handles arrays, nulls, and unwanted spaces
function Resolve-Value {
    param ($Value)
    if ($Value -is [array]) {
        return ($Value -join ", ") # Join arrays into a single string
    } elseif ($null -eq $Value) {
        return "" # Replace nulls with empty strings
    } else {
        return $Value.ToString().Trim() # Trim strings and return as-is
    }
}

# Process each row in Book2
$book2 | ForEach-Object {
    $serverNameB2 = Resolve-Value $_.ServerName

    # Match for Resource Group (from Book1)
    $resourceGroupMatch = $book1 | Where-Object {
        $serverNameB1 = Resolve-Value $_.ServerName
        $serverNameB1 -like "${serverNameB2}*"  # Match the beginning of ServerName
    } | Select-Object -First 1 # Only take the first match

    # Update Resource Group if match found
    if ($resourceGroupMatch) {
        $_."Resource Group" = Resolve-Value $resourceGroupMatch."Resource Group"
        Write-Host "Resource Group updated for $serverNameB2"
    } else {
        Write-Host "No Resource Group match found for $serverNameB2"
    }

    # Match for Disk Sizes (from Book1)
    $diskMatches = $book1 | Where-Object { $_.ServerName -like "*$serverNameB2*" }
    if ($diskMatches) {
        $_."osdisk" = ($diskMatches | Where-Object { $_.ServerName -match "osdisk" } | Measure-Object -Property 'SIZE (GIB)' -Sum).Sum
        $_."datadisk" = ($diskMatches | Where-Object { $_.ServerName -notmatch "osdisk" } | Measure-Object -Property 'SIZE (GIB)' -Sum).Sum
        Write-Host "Disk sizes updated for $serverNameB2"
    } else {
        Write-Host "No disk size match found for $serverNameB2"
    }
}

# Export the updated data to a new Excel file
$book2 | Export-Excel -Path $outputPath -WorksheetName "UpdatedData" -AutoSize -ClearSheet

Write-Host "Book2 has been updated and saved to $outputPath"
