#Requires -Version 7.4

function Convert-OfficeDocumentToOpenDocument {
    <#
    .SYNOPSIS
        Converts .doc, .docx, .xls, .xlsx, .ppt and .pptx to open document format using LibreOffice

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
        Convert-OfficeDocumentsToOpenDocument -path "C:\Documents" -recursive

    .NOTES
        Requires LibreOffice to be installed
        Requires PowerShell 7.4 LTS or later
    #>

    param(
        [Parameter(Mandatory)]
        [string]$path,
        [switch]$recursive,
        [string]$logPath = "$env:USERPROFILE\Desktop"
    )
    
    # Check common LibreOffice installation paths
    $defaultLibreOfficePaths = @(
        "${env:ProgramFiles}\LibreOffice\program\soffice.exe",
        "${env:ProgramFiles(x86)}\LibreOffice\program\soffice.exe"
    )

    # Check if LibreOffice is installed
    $libreOfficePath = $defaultLibreOfficePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $libreOfficePath) {
        Write-Error "LibreOffice not found. Please install LibreOffice first."
        return
    }

    # Create log file
    $logFileName = "ConversionlogLibreOffice_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".csv"
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
    $filesToConvert += Get-ChildItem -Path $path -Filter *.docx -Recurse:$recursive
    $filesToConvert += Get-ChildItem -Path $path -Filter *.xls -Recurse:$recursive
    $filesToConvert += Get-ChildItem -Path $path -Filter *.xlsx -Recurse:$recursive
    $filesToConvert += Get-ChildItem -Path $path -Filter *.ppt -Recurse:$recursive
    $filesToConvert += Get-ChildItem -Path $path -Filter *.pptx -Recurse:$recursive

    if ($filesToConvert.Count -eq 0) {
        Write-Warning "No office files found in '$path'"
        return
    }

    foreach ($sourceFile in $filesToConvert) {
        $fileIndex++
        
        # Calculate progress
        $percentComplete = [int](($fileIndex / $filesToConvert.Count) * 100)
                
        # Display progress bar
        Write-Progress -Activity "Converting Office Documents to Open Document Format" -Status "Processing: $($sourceFile.Name)" -PercentComplete $percentComplete -CurrentOperation "File $fileIndex of $($filesToConvert.Count)"

        try {
            switch ($sourceFile.Extension) {
                '.doc' {
                    $targetFilePath = $sourceFile.FullName -replace '\.doc$', '.odt'
                    $targetFileFormat = 'odt'
                }
                '.docx' {
                    $targetFilePath = $sourceFile.FullName -replace '\.docx$', '.odt'
                    $targetFileFormat = 'odt'
                }
                '.xls' {
                    $targetFilePath = $sourceFile.FullName -replace '\.xls$', '.ods'
                    $targetFileFormat = 'ods'
                }
                '.xlsx' {
                    $targetFilePath = $sourceFile.FullName -replace '\.xlsx$', '.ods'
                    $targetFileFormat = 'ods'
                }
                '.ppt' {
                    $targetFilePath = $sourceFile.FullName -replace '\.ppt$', '.odp'
                    $targetFileFormat = 'odp'
                }
                '.pptx' {
                    $targetFilePath = $sourceFile.FullName -replace '\.pptx$', '.odp'
                    $targetFileFormat = 'odp'
                }
            }

            # Check if file already exists
            if (Test-Path $targetFilePath) {
                Write-Warning "Skipping $($sourceFile.Name) - output already exists"
                
                $skippedFileCount++

                # Add log to logfile
                Add-Content -Path $logFilePath -Value ("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$($sourceFile.FullName),$targetFilePath,Skipped,File already exists")
            }
            else {
                # Convert using LibreOffice
                $arguments = @(
                    "--headless",
                    "--convert-to",
                    $targetFileFormat,
                    "--outdir",
                    (Split-Path -Parent $sourceFile.FullName),
                    $sourceFile.FullName
                )

                $process = Start-Process -FilePath $libreOfficePath.FullName -ArgumentList $arguments -NoNewWindow -PassThru -Wait

                if ($process.ExitCode -eq 0 -and (Test-Path $targetFilePath)) {
                    $convertedFileCount++

                    # Add log to logfile
                    Add-Content -Path $logFilePath -Value ("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$($sourceFile.FullName),$targetFilePath,Success,")
                }
                else {
                    throw "LibreOffice conversion process failed with exit code: $($process.ExitCode)"
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
    Write-Progress -Activity "Converting Office Documents to Open Document Format" -Completed

    # Write summary
    Write-Host "Conversion Summary: Converted=$convertedFileCount, Skipped=$skippedFileCount, Failed=$failedFileCount" -ForegroundColor Green
}
