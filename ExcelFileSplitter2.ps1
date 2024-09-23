# PowerShell script to split Excel file into multiple files with max 10,000 rows and retain header and formatting

# Input and output paths
$inputFile = "C:\path\to\your\input.xlsx"
$outputDirectory = "C:\path\to\output\directory"

$maxRowsPerFile = 10000

# Create Excel Application object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # Open the workbook
    $workbook = $excel.Workbooks.Open($inputFile)
    $worksheet = $workbook.Worksheets.Item(1)

    # Get the total number of rows in the worksheet
    $totalRows = $worksheet.UsedRange.Rows.Count
    $totalColumns = $worksheet.UsedRange.Columns.Count

    # Calculate the number of files needed
    $fileCount = [math]::Ceiling(($totalRows - 1) / $maxRowsPerFile)  # Subtract 1 to account for the header row

    # Copy the header (first row) from the original worksheet
    $headerRange = $worksheet.Range("A1").Resize(1, $totalColumns)

    for ($i = 1; $i -le $fileCount; $i++) {
        # Create a new workbook for each split
        $newWorkbook = $excel.Workbooks.Add()
        $newWorksheet = $newWorkbook.Worksheets.Item(1)

        # Copy the header into the new worksheet
        $headerRange.Copy()
        $newWorksheet.Range("A1").PasteSpecial(-4163)  # Paste the header with formatting

        # Calculate the row range for the current split (excluding the header row)
        $startRow = (($i - 1) * $maxRowsPerFile) + 2  # Start from the 2nd row in original (since row 1 is the header)
        $endRow = [math]::Min($i * $maxRowsPerFile + 1, $totalRows)

        # Number of rows to copy (excluding header row)
        $rowsToCopy = $endRow - $startRow + 1

        if ($rowsToCopy -gt 0) {
            # Copy the data range from the original worksheet
            $sourceRange = $worksheet.Range("A$startRow").Resize($rowsToCopy, $totalColumns)
            $sourceRange.Copy()

            # Paste into the new worksheet, starting at row 2 (since row 1 is the header)
            $newWorksheet.Range("A2").PasteSpecial(-4163)  # xlPasteAllUsingSourceTheme
        }

        # Save the new workbook
        $outputFile = Join-Path $outputDirectory ("SplitFile_" + $i + ".xlsx")
        $newWorkbook.SaveAs($outputFile)

        # Close the new workbook
        $newWorkbook.Close()
    }

    # Close the original workbook
    $workbook.Close()
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    # Quit the Excel application
    $excel.Quit()

    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    # Force garbage collection
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "Excel file has been split into $fileCount files."
