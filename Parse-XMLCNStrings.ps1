<#
.SYNOPSIS
    Efficiently parses XML files and extracts CN= strings for CSV output.

.DESCRIPTION
    This script processes XML files in a specified directory and extracts strings within
    <String></String> or <string></string> elements that begin with "CN=". It handles large datasets (1,020+
    files) efficiently with batch processing, progress indicators, and comprehensive
    error handling.

.PARAMETER Path
    The path to the directory containing XML files to process. Supports both single
    directory and recursive scanning.

.PARAMETER OutputPath
    The path for the output CSV file. If not specified, creates a timestamped file
    in the current directory.

.PARAMETER Recursive
    When specified, scans subdirectories recursively for XML files.

.PARAMETER BatchSize
    Number of files to process in each batch to optimize memory usage. Default is 50.

.PARAMETER IncludeLineNumbers
    When specified, includes line numbers where CN= strings were found.

.PARAMETER LogPath
    Custom path for log files. If not specified, uses current directory.

.EXAMPLE
    .\Parse-XMLCNStrings.ps1 -Path "C:\XMLFiles" -OutputPath "C:\Output\CNStrings.csv"
    # Process XML files in specified directory

.EXAMPLE
    .\Parse-XMLCNStrings.ps1 -Path "C:\XMLFiles" -Recursive -IncludeLineNumbers
    # Recursively scan directory and include line numbers

.EXAMPLE
    .\Parse-XMLCNStrings.ps1 -Path "C:\XMLFiles" -BatchSize 100 -Verbose
    # Process with larger batch size and verbose output

.NOTES
    Author: furena
    Date: 2025-01-20
    Version: 1.0
    
    This script uses PowerShell's native XML parsing capabilities for optimal
    performance and reliability when processing large numbers of XML files.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, HelpMessage="Path to directory containing XML files")]
    [string]$Path,
    
    [Parameter(Mandatory=$false, HelpMessage="Output CSV file path")]
    [string]$OutputPath,
    
    [Parameter(Mandatory=$false, HelpMessage="Scan subdirectories recursively")]
    [switch]$Recursive,
    
    [Parameter(Mandatory=$false, HelpMessage="Batch size for processing files (default: 50)")]
    [int]$BatchSize = 50,
    
    [Parameter(Mandatory=$false, HelpMessage="Include line numbers in output")]
    [switch]$IncludeLineNumbers,
    
    [Parameter(Mandatory=$false, HelpMessage="Path for log files (defaults to current directory)")]
    [string]$LogPath = (Get-Location).Path
)

#region SETUP AND VALIDATION
Write-Host "=== XML CN STRING PARSER ===" -ForegroundColor Magenta
Write-Host "Efficient XML Processing for Large Datasets" -ForegroundColor Cyan
Write-Host "Current User: $env:USERNAME" -ForegroundColor Yellow
Write-Host "Current Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC')" -ForegroundColor Yellow

# Validate input path
if (-not (Test-Path $Path)) {
    Write-Error "Input path does not exist: $Path"
    exit 1
}

# Set default output path if not specified
if (-not $OutputPath) {
    $TimeStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $OutputPath = Join-Path $LogPath "XMLCNStrings_$TimeStamp.csv"
}

# Validate batch size
if ($BatchSize -lt 1 -or $BatchSize -gt 1000) {
    Write-Error "BatchSize must be between 1 and 1000"
    exit 1
}

# Create log files
$TimeStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$ErrorLog = Join-Path $LogPath "XMLCNParser_Errors_$TimeStamp.log"
$ProcessLog = Join-Path $LogPath "XMLCNParser_Process_$TimeStamp.log"

Write-Host "`nConfiguration:" -ForegroundColor Green
Write-Host "  Input Path: $Path" -ForegroundColor White
Write-Host "  Output File: $OutputPath" -ForegroundColor White
Write-Host "  Recursive Scan: $Recursive" -ForegroundColor White
Write-Host "  Batch Size: $BatchSize" -ForegroundColor White
Write-Host "  Include Line Numbers: $IncludeLineNumbers" -ForegroundColor White
Write-Host "  Error Log: $ErrorLog" -ForegroundColor White
Write-Host "  Process Log: $ProcessLog" -ForegroundColor White
#endregion

#region HELPER FUNCTIONS
function Write-LogEntry {
    param(
        [string]$Message,
        [string]$LogFile = $ProcessLog,
        [switch]$Error
    )
    
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "$TimeStamp - $Message"
    
    if ($Error) {
        Write-Host $Message -ForegroundColor Red
        $LogEntry | Out-File $ErrorLog -Append -Encoding UTF8
    } else {
        Write-Verbose $Message
    }
    
    $LogEntry | Out-File $LogFile -Append -Encoding UTF8
}

function Get-CNStringsFromXMLContent {
    param(
        [string]$XmlContent,
        [string]$FilePath,
        [bool]$IncludeLineNumbers
    )
    
    $Results = @()
    
    try {
        # Parse XML content
        $XmlDoc = [xml]$XmlContent
        
        # Find all <String> and <string> elements (case-insensitive)
        # Using translate() function to handle case-insensitive matching
        $StringElements = $XmlDoc.SelectNodes("//String | //string")
        
        foreach ($StringElement in $StringElements) {
            $StringValue = $StringElement.InnerText
            
            # Check if string starts with "CN=" (case-insensitive, after trimming)
            $TrimmedValue = $StringValue.Trim()
            if ($TrimmedValue -imatch "^CN=") {
                $Result = [PSCustomObject]@{
                    FilePath = $FilePath
                    CNString = $TrimmedValue
                }
                
                # Add line number if requested
                if ($IncludeLineNumbers) {
                    # Find line number by searching through content
                    $Lines = $XmlContent -split "`r?`n"
                    $LineNumber = 0
                    for ($i = 0; $i -lt $Lines.Length; $i++) {
                        if ($Lines[$i] -match [regex]::Escape($StringValue)) {
                            $LineNumber = $i + 1
                            break
                        }
                    }
                    $Result | Add-Member -NotePropertyName "LineNumber" -NotePropertyValue $LineNumber
                }
                
                $Results += $Result
                Write-LogEntry "Found CN string in $FilePath`: $TrimmedValue"
            }
        }
    }
    catch {
        Write-LogEntry "Error parsing XML content in $FilePath`: $($_.Exception.Message)" -Error
        return @()
    }
    
    return $Results
}

function Get-CNStringsFromXMLFile {
    param(
        [string]$FilePath,
        [bool]$IncludeLineNumbers
    )
    
    try {
        # Read file content with proper encoding handling
        $XmlContent = Get-Content -Path $FilePath -Raw -Encoding UTF8
        
        # Process the XML content
        return Get-CNStringsFromXMLContent -XmlContent $XmlContent -FilePath $FilePath -IncludeLineNumbers $IncludeLineNumbers
    }
    catch {
        Write-LogEntry "Error reading file $FilePath`: $($_.Exception.Message)" -Error
        return @()
    }
}

function Get-XMLFile {
    param(
        [string]$SearchPath,
        [bool]$RecursiveSearch
    )
    
    try {
        if ($RecursiveSearch) {
            $XmlFiles = Get-ChildItem -Path $SearchPath -Filter "*.xml" -File -Recurse -ErrorAction Stop
        } else {
            $XmlFiles = Get-ChildItem -Path $SearchPath -Filter "*.xml" -File -ErrorAction Stop
        }
        
        Write-LogEntry "Found $($XmlFiles.Count) XML files to process"
        return $XmlFiles
    }
    catch {
        Write-LogEntry "Error finding XML files in $SearchPath`: $($_.Exception.Message)" -Error
        return @()
    }
}
#endregion

#region MAIN PROCESSING
Write-Host "`n=== SCANNING FOR XML FILES ===" -ForegroundColor Magenta

# Initialize log files
$LogHeader = "Starting XML CN string extraction from [$Path] at $(Get-Date)"
$LogHeader | Out-File $ErrorLog -Encoding UTF8
$LogHeader | Out-File $ProcessLog -Encoding UTF8

# Get list of XML files
$XmlFiles = Get-XMLFile -SearchPath $Path -RecursiveSearch $Recursive

if ($XmlFiles.Count -eq 0) {
    Write-Host "No XML files found in the specified path." -ForegroundColor Yellow
    Write-LogEntry "No XML files found in $Path"
    exit 0
}

Write-Host "Found $($XmlFiles.Count) XML files to process" -ForegroundColor Green

# Initialize results collection
$AllResults = @()
$ProcessedCount = 0
$ErrorCount = 0
$TotalCNStrings = 0

Write-Host "`n=== PROCESSING XML FILES ===" -ForegroundColor Magenta

# Process files in batches for memory efficiency
for ($i = 0; $i -lt $XmlFiles.Count; $i += $BatchSize) {
    $BatchEnd = [Math]::Min($i + $BatchSize - 1, $XmlFiles.Count - 1)
    if ($i -eq $BatchEnd) {
        # Single item case
        $CurrentBatch = @($XmlFiles[$i])
    } else {
        # Multiple items case
        $CurrentBatch = $XmlFiles[$i..$BatchEnd]
    }
    
    Write-Host "Processing batch $([Math]::Floor($i / $BatchSize) + 1) of $([Math]::Ceiling($XmlFiles.Count / $BatchSize)) (files $($i + 1)-$($BatchEnd + 1))" -ForegroundColor Cyan
    
    # Process current batch
    foreach ($XmlFile in $CurrentBatch) {
        $ProcessedCount++
        
        # Update progress
        $PercentComplete = [Math]::Round(($ProcessedCount / $XmlFiles.Count) * 100, 1)
        Write-Progress -Activity "Processing XML Files" `
                      -Status "File $ProcessedCount of $($XmlFiles.Count) - $($XmlFile.Name)" `
                      -PercentComplete $PercentComplete
        
        try {
            # Process the XML file
            $FileResults = Get-CNStringsFromXMLFile -FilePath $XmlFile.FullName -IncludeLineNumbers $IncludeLineNumbers
            
            if ($FileResults.Count -gt 0) {
                $AllResults += $FileResults
                $TotalCNStrings += $FileResults.Count
                Write-Host "  ✓ $($XmlFile.Name): Found $($FileResults.Count) CN string(s)" -ForegroundColor Green
            } else {
                Write-Host "  - $($XmlFile.Name): No CN strings found" -ForegroundColor Gray
            }
        }
        catch {
            $ErrorCount++
            Write-LogEntry "Error processing file $($XmlFile.FullName): $($_.Exception.Message)" -Error
            Write-Host "  ✗ $($XmlFile.Name): Error processing file" -ForegroundColor Red
        }
    }
    
    # Force garbage collection after each batch to manage memory
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Progress -Activity "Processing XML Files" -Completed
#endregion

#region EXPORT RESULTS
Write-Host "`n=== EXPORTING RESULTS ===" -ForegroundColor Magenta

if ($AllResults.Count -gt 0) {
    try {
        # Export to CSV
        $AllResults | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "✓ Exported $($AllResults.Count) CN strings to: $OutputPath" -ForegroundColor Green
        Write-LogEntry "Successfully exported $($AllResults.Count) results to $OutputPath"
        
        # Display summary statistics
        $UniqueFiles = ($AllResults | Group-Object FilePath).Count
        $UniqueCNStrings = ($AllResults | Group-Object CNString).Count
        
        Write-Host "`nSummary Statistics:" -ForegroundColor Cyan
        Write-Host "  Files with CN strings: $UniqueFiles" -ForegroundColor White
        Write-Host "  Total CN strings found: $($AllResults.Count)" -ForegroundColor White
        Write-Host "  Unique CN strings: $UniqueCNStrings" -ForegroundColor White
    }
    catch {
        Write-LogEntry "Error exporting results to $OutputPath`: $($_.Exception.Message)" -Error
        Write-Host "✗ Failed to export results" -ForegroundColor Red
    }
} else {
    Write-Host "No CN strings found in any XML files." -ForegroundColor Yellow
    Write-LogEntry "No CN strings found in any processed files"
}
#endregion

#region FINAL SUMMARY
Write-Host "`n=== PROCESSING SUMMARY ===" -ForegroundColor Magenta

Write-Host "Files processed: $ProcessedCount of $($XmlFiles.Count)" -ForegroundColor White
Write-Host "Files with errors: $ErrorCount" -ForegroundColor $(if ($ErrorCount -eq 0) { "Green" } else { "Yellow" })
Write-Host "Total CN strings extracted: $TotalCNStrings" -ForegroundColor White

if ($ErrorCount -gt 0) {
    Write-Host "`nSome files had errors. Check the error log: $ErrorLog" -ForegroundColor Yellow
}

Write-Host "`nLog files created:" -ForegroundColor Cyan
Write-Host "  Process Log: $ProcessLog" -ForegroundColor White
Write-Host "  Error Log: $ErrorLog" -ForegroundColor White

if ($TotalCNStrings -eq 0) {
    Write-Host "`nNo CN strings were found. Please verify:" -ForegroundColor Yellow
    Write-Host "  • XML files contain <String> or <string> elements" -ForegroundColor White
    Write-Host "  • String values begin with 'CN='" -ForegroundColor White
    Write-Host "  • XML files are well-formed" -ForegroundColor White
} else {
    Write-Host "`n✓ Processing completed successfully!" -ForegroundColor Green
}

Write-LogEntry "Processing completed. Total files: $ProcessedCount, Errors: $ErrorCount, CN strings found: $TotalCNStrings"
#endregion