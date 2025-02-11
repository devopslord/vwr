# README

## Overview

## Step 1: Check ODCR Availability

ODCRs are only available in AWS Regions that support ODCR and apply to specific instance families. Ensure your target RDS instance type is eligible for ODCR.

You can check ODCR availability via AWS CLI:

aws rds describe-reserved-db-instances-offerings --query "ReservedDBInstancesOfferings[?contains(DBInstanceClass, 'db.')].{Instance:DBInstanceClass, OfferingID:ReservedDBInstancesOfferingId}" --output table


## Step 2: Request an ODCR for Your Read Replica

Use the AWS CLI or AWS Console to request an ODCR. Here’s how to do it using the AWS CLI:

aws rds purchase-reserved-db-instances-offering \
    --reserved-db-instances-offering-id <OFFERING_ID> \
    --db-instance-class db.r6g.large \
    --duration 31536000 \
    --product-description "PostgreSQL" \
    --offering-type "All Upfront" \
    --tags Key=Purpose,Value=ReadReplica


	•	<OFFERING_ID>: Retrieve from describe-reserved-db-instances-offerings
	•	db.r6g.large: Replace with your desired instance class
	•	duration 31536000: 1-year reservation (for 3 years, use 94608000)
	•	offering-type: Can be All Upfront, Partial Upfront, or No Upfront


## Step 3: Apply the ODCR to Your Read Replica

When creating a read replica, ensure it matches the instance class and region of the ODCR:

aws rds create-db-instance-read-replica \
    --db-instance-identifier my-read-replica \
    --source-db-instance-identifier my-primary-db \
    --db-instance-class db.r6g.large \
    --availability-zone us-east-1a \
    --no-publicly-accessible

## Step 4: Verify Reservation is Applied

Run this to confirm your ODCR is active:

Additional Considerations
	•	ODCRs apply automatically when a matching instance is launched.
	•	You cannot modify ODCRs after purchase.
	•	Ensure you use the reserved instance size and region, otherwise, it won’t apply.
	•	Reserved pricing does not cover storage or backups—only compute.

Let me know if you need a tailored Terraform solution! 










There are two PowerShell scripts designed to automate Excel file processing:

1. **Unmerge.ps1** - Cleans and formats Excel files by unmerging cells, removing unnecessary rows, and handling empty columns.
2. **Dataretrieve.ps1** - Retrieves specific publisher-related data from multiple Excel files and consolidates it into a single output file.

Both scripts leverage the `ImportExcel` module for Excel manipulations.

## Prerequisites
- Windows OS with PowerShell installed.
- Microsoft Excel installed.
- [`ImportExcel`](https://www.powershellgallery.com/packages/ImportExcel) PowerShell module.

### Install `ImportExcel`:
```powershell
Install-Module -Name ImportExcel -Force -Scope CurrentUser
```
75D2SXJ4KR
## Script Details

### 1. Unmerge.ps1
#### Description
`Unmerge.ps1` processes Excel files by:
- Detecting and unmerging merged cells.
- Removing unnecessary rows above a specified header row.
- Identifying and removing empty columns.
- Saving the cleaned files in a specified output directory.

#### Setup
- **Source Directory:** Define the path containing the Excel files to clean:
  ```powershell
  $sourceDir = "C:\Path\To\Your\Excel\Files\*.xlsx"
  ```
  ![image](https://github.com/user-attachments/assets/c9432a50-f590-4aef-b1d9-9efacc6baca1)

- **Output Directory:** Define where cleaned files will be saved:
  ```powershell
  $outputDir = "C:\Path\To\Save\Cleaned\Files"
  ```
  ![image](https://github.com/user-attachments/assets/77363dd5-be0d-4ede-9629-af187142878f)

#### Usage
1. Open PowerShell or Terminal in your IDE.

2. Navigate to the script’s directory:
   ```powershell
   cd "C:\Path\To\Script"
   ```
   ![image](https://github.com/user-attachments/assets/83d7c1d7-2932-4cc2-9869-92cc3318283b)

   
3. Run the script:
   ```powershell
   .\Unmerge.ps1
   ```
   ![image](https://github.com/user-attachments/assets/72d47331-002f-4ff4-b8c9-757ec94d9a13)
   

#### Output
- Cleaned Excel files saved in the specified output directory.
  ![image](https://github.com/user-attachments/assets/aba992b0-93b9-4e0e-9a47-b3f20bb0abdb)


### 2. Dataretrieve.ps1
#### Description
`dataretrieve.ps1` extracts publisher-specific data from Excel files by:
- Reading publisher names from `publisher.xlsx`.
- Searching for matches in other Excel files within a specified directory.
- Extracting related product and OS details.
- Saving results into `outputfile.xlsx`.

#### Setup
- **Publisher File:** Path to the Excel file containing a list of publishers:
  ```powershell
  $publisherFile = "C:\Path\To\publisher.xlsx"
  ```
  ![image](https://github.com/user-attachments/assets/20fb7e14-363e-4db5-b788-241b6284b776)

- **Source Directory:** Directory containing Excel files to search:
  ```powershell
  $sourceDirectory = "C:\Path\To\Source\Files"
  ```
  ![image](https://github.com/user-attachments/assets/835b3942-a053-4daf-a429-48c6622b3ff6)

- **Output File:** Path for the result file:
  ```powershell
  $outputFile = "C:\Path\To\outputfile.xlsx"
  ```
  ![image](https://github.com/user-attachments/assets/695b410f-0571-4d5e-ba7b-09fa15ea1a0e)


#### Usage
1. Ensure `publisher.xlsx` contains a `"publisher"` column with names.
2. Place Excel files in the source directory.
3. Open PowerShell as Administrator.
4. Navigate to the script’s directory:
   ```powershell
   cd "C:\Path\To\Script"
   ```
   ![image](https://github.com/user-attachments/assets/ce86de28-3e32-4d2f-9f3c-dc12d21498cb)

5. Run the script:
   ```powershell
   .\dataretrieve.ps1
   ```

#### Output
- Extracted data saved into `outputfile.xlsx`.

  ![image](https://github.com/user-attachments/assets/e5c63f82-4d44-4572-8119-e0b5db8caed6)


## Troubleshooting
- For script execution issues, temporarily bypass execution policy:
  ```powershell
  Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
  ```
- Ensure all Excel files are closed before running the scripts.
- Verify that Excel and the `ImportExcel` module are installed.

## Conclusion
These PowerShell scripts simplify the process of cleaning and extracting data from Excel files, making data preparation and analysis more efficient. For any issues or further customization, review the script comments or modify paths as needed.
