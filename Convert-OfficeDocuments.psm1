#Requires -Version 7.4

function Convert-OfficeDocuments {
    <#
    .SYNOPSIS
        Converts .doc and .xls to the more modern filetype formats docx and .xlsx

    .PARAMETER path
        Directory path containing the office files to convert
        Mandatory

    .PARAMETER recursive
        Recursively searches through files and subdirectories.
        Optional

    .PARAMETER logPath
        CSV log path
        Default: Current user's Desktop

    .EXAMPLE
        Convert-OfficeDocuments -path "C:\Documents" -recursive

    .NOTES
        Requires modern version of Word and/or Excel to be installed
        Requires PowerShell 7.4 LTS or later
    #>

    param(
        [Parameter(Mandatory)]
        [string]$path,
        [switch]$recursive,
        [string]$logPath = "$env:USERPROFILE\Desktop"
    )

    # Initialize COM objects
    try {
        $wordApp = New-Object -ComObject Word.Application
        $wordApp.Visible = $false
        $wordApp.ScreenUpdating = $false

        $excelApp = New-Object -ComObject Excel.Application
        $excelApp.Visible = $false
        $excelApp.ScreenUpdating = $false
    }
    catch {
        Write-Error "Office not installed: $_"
        return
    }

    # Create log file
    $logFileName = "Conversionlog_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".csv"
    $logFilePath = Join-Path -Path $logPath -ChildPath $logFileName
    "Timestamp,SourceFile,DestinationFile,Result,ErrorMessage" | Set-Content $logFilePath

    # Initialize counters for summary  
    $convertedFileCount = 0
    $skippedFileCount = 0
    $failedFileCount = 0

    # Initialize counter for progress bar
    $fileIndex = 0

    # Collect files
    $filesToConvert = @()
    $filesToConvert += Get-ChildItem -Path $path -Filter *.doc -Recurse:$recursive
    $filesToConvert += Get-ChildItem -Path $path -Filter *.xls -Recurse:$recursive

    if ($filesToConvert.Count -eq 0) {
        Write-Warning "No .doc or .xls files found in '$path'"
        return
    }

    foreach ($sourceFile in $filesToConvert) {
        $fileIndex++
        
        # Calculate progress
        $percentComplete = [int](($fileIndex / $filesToConvert.Count) * 100)
                
        # Display progress bar
        Write-Progress -Activity "Converting Office Documents" -Status "Processing: $($sourceFile.Name)" -PercentComplete $percentComplete -CurrentOperation "File $fileIndex of $($filesToConvert.Count)"

        # Reset document
        $officeDocument = $null

        try {
            # Convert word documents
            if ($sourceFile.Extension -eq '.doc') {
                # Check if file already exists
                $targetFilePath = $sourceFile.FullName -replace '\.doc$', '.docx'
                if (Test-Path $targetFilePath) {
                    Write-Warning "Skipping $($sourceFile.Name) - output already exists"
                    
                    $skippedFileCount++

                    # Add log to logfile
                    Add-Content -Path $logFilePath -Value ("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$($sourceFile.FullName),$targetFilePath,Skipped,File already exists")
                }
                else {
                    # Open, convert and save document
                    $officeDocument = $wordApp.Documents.Open($sourceFile.FullName)
                    $officeDocument.SaveAs($targetFilePath, 12)
                    $officeDocument.Close($false)

                    $convertedFileCount++

                    # Add log to logfile
                    Add-Content -Path $logFilePath -Value ("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$($sourceFile.FullName),$targetFilePath,Success,")
                }
            }
            # Convert excel documents
            else {
                # Check if file already exists
                $targetFilePath = $sourceFile.FullName -replace '\.xls$', '.xlsx'
                if (Test-Path $targetFilePath) {
                    Write-Warning "Skipping $($sourceFile.Name) - output already exists"

                    $skippedFileCount++

                    # Add log to logfile
                    Add-Content -Path $logFilePath -Value ("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$($sourceFile.FullName),$targetFilePath,Skipped,File already exists")
                }
                else {
                    # Open, convert and save document
                    $officeDocument = $excelApp.Workbooks.Open($sourceFile.FullName)
                    $officeDocument.SaveAs($targetFilePath, 51)
                    $officeDocument.Close($false)

                    $convertedFileCount++

                    # Add log to logfile
                    Add-Content -Path $logFilePath -Value ("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$($sourceFile.FullName),$targetFilePath,Success,")
                }
            }
        }
        catch {
            Write-Error "Failed to convert '$($sourceFile.FullName)': $_"
            $failedFileCount++

            # Add log to logfile
            Add-Content -Path $logFilePath -Value ("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$($sourceFile.FullName),,Failed,$($_.Exception.Message -replace '[\r\n]+', ' ')")

            continue
        }
    }

    # Complete the progress bar
    Write-Progress -Activity "Converting Office Documents" -Completed

    # Cleanup COM objects     
    if ($wordApp) {
        try { $wordApp.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordApp) | Out-Null
    }
    if ($excelApp) {
        try { $excelApp.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) | Out-Null
    }

    # Garbage collection
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    
    # Write summary
    Write-Host "Conversion Summary: Converted=$convertedFileCount, Skipped=$skippedFileCount, Failed=$failedFileCount" -ForegroundColor Green
}