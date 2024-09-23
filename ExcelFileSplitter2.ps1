# PowerShell script to split Excel file into multiple files with max 10,000 rows and retain formatting

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
    $fileCount = [math]::Ceiling($totalRows / $maxRowsPerFile)

    for ($i = 1; $i -le $fileCount; $i++) {
        # Create a new workbook for each split
        $newWorkbook = $excel.Workbooks.Add()
        $newWorksheet = $newWorkbook.Worksheets.Item(1)

        # Calculate the row range for the current split
        $startRow = (($i - 1) * $maxRowsPerFile) + 1
        $endRow = [math]::Min($i * $maxRowsPerFile, $totalRows)

        # Number of rows to copy
        $rowsToCopy = $endRow - $startRow + 1

        # Copy the specified range from the original worksheet
        $sourceRange = $worksheet.Range("A$startRow").Resize($rowsToCopy, $totalColumns)
        $sourceRange.Copy()

        # Paste into the new worksheet
        $newWorksheet.Range("A1").PasteSpecial(-4163)  # xlPasteAllUsingSourceTheme

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
