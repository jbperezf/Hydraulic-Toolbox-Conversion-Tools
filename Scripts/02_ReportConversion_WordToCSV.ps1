# PowerShell script to extract hydraulic analysis data from Word document to CSV
# Compatible with PowerShell 5.1

function Extract-HydraulicData {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InputFilePath,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputFilePath
    )
    
    # Display paths for user verification
    Write-Host "Input file: $InputFilePath"
    Write-Host "Output file: $OutputFilePath"

    # Check if file exists before proceeding
    if (-not (Test-Path -Path $InputFilePath)) {
        Write-Host ""
        Write-Host "TROUBLESHOOTING:" -ForegroundColor Yellow
        Write-Host "The input file could not be found. Please make sure:" -ForegroundColor Yellow
        Write-Host "1. The file exists at the specified path" -ForegroundColor Yellow
        Write-Host "2. The filename spelling matches exactly (including case)" -ForegroundColor Yellow
        return $false
    }

    # Define the CSV header
    $headers = @(
        "Channel Analysis", "Channel Type", "Side Slope 1 (Z1) ft/ft", "Side Slope 2 (Z2) ft/ft",
        "Channel Width ft", "Longitudinal Slope ft/ft", "Manning's n", "Flow cfs", "Depth ft",
        "Area of Flow ft^2", "Wetted Perimeter ft", "Hydraulic Radius ft", "Average Velocity ft/s",
        "Top Width ft", "Froude Number", "Critical Depth ft", "Critical Velocity ft/s",
        "Critical Slope ft/ft", "Critical Top Width ft", "Calculated Max Shear Stress lb/ft^2",
        "Calculated Avg Shear Stress lb/ft^2"
    )

    # Create header string for CSV file
    $csvHeader = '"' + ($headers -join '","') + '"'
    $csvHeader | Out-File -FilePath $OutputFilePath -Encoding ascii

    # Function to extract text from Word document
    function Get-WordContent {
        param (
            [string]$FilePath
        )
        
        $word = $null
        $doc = $null
        
        try {
            # Create Word COM Object with explicit ProgID
            Write-Host "Creating Word application instance..."
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $word.DisplayAlerts = 0  # Suppress alerts
            
            # Add small delay
            Start-Sleep -Milliseconds 500
            
            # Open the document with ReadOnly flag
            Write-Host "Opening document: $FilePath"
            $doc = $word.Documents.Open($FilePath, $false, $true)  # False for ConfirmConversions, True for ReadOnly
            
            # Add small delay
            Start-Sleep -Milliseconds 500
            
            # Extract text
            Write-Host "Extracting text content..."
            $content = $doc.Content.Text
            
            # Return the content
            return $content
        }
        catch {
            Write-Error "Error processing Word document: $_"
            
            # More detailed troubleshooting
            Write-Host ""
            Write-Host "TROUBLESHOOTING:" -ForegroundColor Yellow
            Write-Host "1. Close all running instances of Microsoft Word before running this script" -ForegroundColor Yellow
            Write-Host "2. Make sure you have permission to access the file and Microsoft Word" -ForegroundColor Yellow
            Write-Host "3. Try running PowerShell as administrator" -ForegroundColor Yellow
            Write-Host "4. Check if Microsoft Word is properly installed and not in need of repair" -ForegroundColor Yellow
            Write-Host ""
            
            return $null
        }
        finally {
            # Properly clean up even if there was an error
            Write-Host "Cleaning up Word resources..."
            if ($doc -ne $null) {
                try { $doc.Close($false) } catch { }
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
            }
            
            if ($word -ne $null) {
                try { $word.Quit() } catch { }
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
            }
            
            # Force garbage collection
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }

    # Read the content of the Word document
    Write-Host "Opening Word document..."
    $content = Get-WordContent -FilePath $InputFilePath

    # Check if content extraction was successful
    if ($null -eq $content) {
        Write-Host "Failed to extract content from Word document. See above errors for details." -ForegroundColor Red
        return $false
    }

    # Split by "Channel Analysis:" to get each analysis section
    $sections = $content -split "Channel Analysis:" | Where-Object { $_ -match "\S" } | Select-Object -Skip 1

    # Track the number of sections processed
    $sectionCount = 0

    # Define regex patterns for data extraction
    $patterns = @{
        "Channel Type" = "Channel Type:\s+([\w]+)"
        "Side Slope 1 (Z1) ft/ft" = "Side Slope 1 \(Z1\):\s+([\d\.]+)\s+ft/ft"
        "Side Slope 2 (Z2) ft/ft" = "Side Slope 2 \(Z2\):\s+([\d\.]+)\s+ft/ft"
        "Channel Width ft" = "Channel Width\s+([\d\.]+)\s+ft"
        "Longitudinal Slope ft/ft" = "Longitudinal Slope:\s+([\d\.]+)\s+ft/ft"
        "Manning's n" = "Manning's n:\s+([\d\.]+)"
        "Flow cfs" = "Flow\s+([\d\.]+)\s+cfs"
        "Depth ft" = "Depth\s+([\d\.]+)\s+ft"
        "Area of Flow ft^2" = "Area of Flow\s+([\d\.]+)\s+ft\^2"
        "Wetted Perimeter ft" = "Wetted Perimeter\s+([\d\.]+)\s+ft"
        "Hydraulic Radius ft" = "Hydraulic Radius\s+([\d\.]+)\s+ft"
        "Average Velocity ft/s" = "Average Velocity\s+([\d\.]+)\s+ft/s"
        "Top Width ft" = "Top Width\s+([\d\.]+)\s+ft"
        "Froude Number" = "Froude Number:\s+([\d\.]+)"
        "Critical Depth ft" = "Critical Depth\s+([\d\.]+)\s+ft"
        "Critical Velocity ft/s" = "Critical Velocity\s+([\d\.]+)\s+ft/s"
        "Critical Slope ft/ft" = "Critical Slope:\s+([\d\.]+)\s+ft/ft"
        "Critical Top Width ft" = "Critical Top Width\s+([\d\.]+)\s+ft"
        "Calculated Max Shear Stress lb/ft^2" = "Calculated Max Shear Stress\s+([\d\.]+)\s+lb/ft\^2"
        "Calculated Avg Shear Stress lb/ft^2" = "Calculated Avg Shear Stress\s+([\d\.]+)\s+lb/ft\^2"
    }

    foreach ($section in $sections) {
        # Extract only the channel analysis name (first line, before "Notes:" or any other text)
        $firstLine = ($section -split "(\r\n|\n)")[0].Trim()
        
        # If there are multiple lines in the first part, get just the first one
        if ($firstLine -match "^(.*?)(\s*Notes:|\s*Input Parameters|$)") {
            $channelAnalysis = $matches[1].Trim()
        } else {
            $channelAnalysis = $firstLine  # Fallback to just using the first line
        }
        
        # Create hashtable to store data with default values
        $data = @{ "Channel Analysis" = $channelAnalysis }
        
        # Set default for channel width (for triangular channels)
        $data["Channel Width ft"] = "0.00"
        
        # Extract all other fields using the regex patterns
        foreach ($field in $patterns.Keys) {
            if ($section -match $patterns[$field]) {
                $data[$field] = $matches[1]
            } else {
                # Only add empty value if key doesn't exist already
                if (-not $data.ContainsKey($field)) {
                    $data[$field] = ""
                }
            }
        }
        
        # Create CSV line array
        $csvValues = @()
        foreach ($header in $headers) {
            $value = if ($data.ContainsKey($header)) { $data[$header] } else { "" }
            $csvValues += $value
        }
        
        # Join values into CSV format and write to file
        $csvLine = '"' + ($csvValues -join '","') + '"'
        $csvLine | Out-File -FilePath $OutputFilePath -Encoding ascii -Append
        
        # Increment section counter
        $sectionCount++
    }

    if ($sectionCount -eq 0) {
        Write-Host "No hydraulic analysis sections were found in the document." -ForegroundColor Yellow
        Write-Host "Please check that the document contains sections starting with 'Channel Analysis:'" -ForegroundColor Yellow
        return $false
    } else {
        Write-Host "Extraction complete. Processed $sectionCount channel analyses." -ForegroundColor Green
        Write-Host "Data saved to: $OutputFilePath" -ForegroundColor Green
        return $true
    }
}

# Get script directory
$scriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

# Get parent directory
$parentPath = Split-Path -Parent -Path $scriptPath

# Define file paths
$inputPath = Join-Path -Path $parentPath -ChildPath "report.docx"
$outputPath = Join-Path -Path $parentPath -ChildPath "OutputDitchData.csv"

# Call the function with the parameters
Extract-HydraulicData -InputFilePath $inputPath -OutputFilePath $outputPath