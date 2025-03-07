function New-HydFile {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, HelpMessage="Path to the input CSV file")]
        [ValidateScript({Test-Path $_ -PathType Leaf -Include *.csv})]
        [string]$CsvFilePath,
        
        [Parameter(Mandatory=$false, HelpMessage="Project name for the report title and output filename")]
        [string]$ProjectName = ""
    )

    begin {
        # Logging and Configuration
        $Host.UI.RawUI.WindowTitle = "Hydraulic Toolbox Ditches Creator"
        $ErrorActionPreference = 'Stop'
        
        # Improved logging function
        function Write-ColorOutput {
            param(
                [Parameter(Mandatory=$true)]
                [string]$Message,
                [Parameter(Mandatory=$false)]
                [System.ConsoleColor]$ForegroundColor = 'White'
            )
            Write-Host $Message -ForegroundColor $ForegroundColor
        }
        
        # If project name is not provided, ask for it
        if ([string]::IsNullOrWhiteSpace($ProjectName)) {
            $ProjectName = Read-Host "Enter the project name"
            if ([string]::IsNullOrWhiteSpace($ProjectName)) {
                $ProjectName = "Hydraulic Toolbox"  # Default if user enters nothing
            }
        }
        
        # Determine output file path based on project name
        $OutputFileName = "$ProjectName - Ditches.hyd"
        $OutputDirectory = Split-Path -Path $CsvFilePath -Parent
        $OutputDirectory = if ([string]::IsNullOrWhiteSpace($OutputDirectory)) { 
            (Get-Location).Path 
        } else { 
            $OutputDirectory 
        }
        $OutputFilePath = Join-Path -Path $OutputDirectory -ChildPath $OutputFileName
    }

    process {
        try {
            # Resolve full path and validate input file
            $CsvFilePath = Resolve-Path $CsvFilePath

            # Create temporary file
            $TempTextFile = [System.IO.Path]::GetTempFileName()

            # Step 1: Read and process input CSV file
            Write-ColorOutput "Step 1: Processing input CSV file..." -ForegroundColor Yellow
            
            # Read CSV data with error handling
            $CsvData = Import-Csv $CsvFilePath
            Write-ColorOutput "  - Loaded $($CsvData.Count) rows from CSV" -ForegroundColor Green

            # Refactored block generation with improved readability
            function New-ChannelCalcBlock {
                param($Row)
                
                $ChannelType = if ($Row.WIDTH -eq '0') { 2 } else { 0 }
                $ChannelGuid = [System.Guid]::NewGuid()

                return @"
CHANNELCALC
CHANNELNAME          "$($Row.CHANNELNAME)"
CHANNELNOTES         "$($Row.CHANNELNOTES)"
LATITUDE             0.000000
LONGITUDE            0.000000
CHANNELTYPE          $ChannelType
ZSCALE               0
CALCTYPE             1
FLOW                 $($Row.FLOW)   
SIDESLOPE1           $($Row.SIDESLOPE1)
SIDESLOPE2           $($Row.SIDESLOPE2)
WIDTH                $($Row.WIDTH)
DEPTH                0.000000
LONGSLOPE            $($Row.LONGSLOPE)
MANNINGS             $($Row.MANNINGS)
PIPEDIAMETER         0.000000
HYDRADIUS            0.000000
PERMSHEARSTRESS      1.500000
CALCMAXSHEARSTRESS   0.000000
CALCAVGSHEARSTRESS   0.000000
AREAOFFLOW           0.000000
AVEVELOCITY          0.000000
WETTEDPERIMETER      0.000000
TOPWIDTH             0.000000
FROUDE               0.000000
CRITICALDEPTH        0.000000
CRITICALTOPWIDTH     0.000000
CRITICALVELOCITY     0.000000
CRITICALSLOPE        0.000000
STABILITYFACTOR      0.000000
RISE                 0.000000
SPAN                 0.000000
CROSSSECTIONREADONLY 0
CROSSECTIONDATA      3
STATION    0.000000 STATIONELEV 0.000000 STATIONMANNINGS 0.000000
STATION    0.000000 STATIONELEV 0.000000 STATIONMANNINGS 0.000000
STATION    0.000000 STATIONELEV 0.000000 STATIONMANNINGS 0.000000
ENDCROSSECTIONDATA
CHANNELGUID          $ChannelGuid
ENDCHANNELCALC

"@
            }

            function New-LiningCalcBlock {
                param($Row)
                
                $LiningGuid = [System.Guid]::NewGuid()

                return @"
LININGCALC
LININGNAME           "$($Row.CHANNELNAME) - Lining Calc"
LININGNOTES          ""
LATITUDE             0.000000
LONGITUDE            0.000000
LININGTYPE           2
SELMETHOD            0
SELCHANNEL           0
	RIPRAPLINING
		SAFETYFACTOR         1.000000
		GAMMAWATER           62.400000
		MANNINGSN            0.000000
		GEOMCALC             0
		CURVATURERAD         0.000000
		ENDBASECLASS
	D50                  0.000000
	GAMMASOIL            165.000000
	SHAPEFACTOR          1
	ENDRIPRAPLINING
	VEGLINING
		SAFETYFACTOR         1.000000
		GAMMAWATER           62.400000
		MANNINGSN            0.000000
		GEOMCALC             0
		CURVATURERAD         0.000000
		ENDBASECLASS
	HEIGHT               0.333000
	CN                   0.000000
	CF                   0.750000
	D75                  0.100000
	PLASTICITY           0.000000
	POROSITY             0.000000
	C1                   1.070000
	C2                   14.300000
	C3                   47.700000
	C4                   1.420000
	C5                   -0.610000
	C6                   0.000100
	COHESIVE             0
	CONDITION            2
	GROWTHFORM           2
	SOILCLASS            3
	ENDVEGLINING
	RECPLINING
		SAFETYFACTOR         1.000000
		GAMMAWATER           62.400000
		MANNINGSN            0.035000
		GEOMCALC             0
		CURVATURERAD         0.000000
		ENDBASECLASS
	SHEARRECP            2.250000
	SHEARMID             0.000000
	MANNINGLOW           0.000000
	MANNINGMID           0.000000
	MANNINGUP            0.000000
	D75                  0.100000
	PLASTICITY           0.000000
	POROSITY             0.000000
	C1                   1.070000
	C2                   14.300000
	C3                   47.700000
	C4                   1.420000
	C5                   -0.610000
	C6                   0.000100
	COHESIVE             0
	MANNINGRANGE         0
	ENDRECPLINING
	GABIONLINING
		SAFETYFACTOR         1.000000
		GAMMAWATER           62.400000
		MANNINGSN            0.000000
		GEOMCALC             0
		CURVATURERAD         0.000000
		ENDBASECLASS
	D50                  0.000000
	GAMMASOIL            165.000000
	MATTRESST            0.000000
	ENDGABIONLINING
LININGGUID           $LiningGuid
ENDLININGCALC

"@
            }

            # Process rows and generate blocks
            $ChannelCalcBlocks = $CsvData | ForEach-Object { New-ChannelCalcBlock -Row $_ }
            $LiningCalcBlocks = $CsvData | 
                Where-Object { $_.CHANNELNAME -like "*100 yr_High Slope*" } | 
                ForEach-Object { New-LiningCalcBlock -Row $_ }

            Write-ColorOutput "  - Processed $($ChannelCalcBlocks.Count) Channel Calculations" -ForegroundColor Green
            Write-ColorOutput "  - Processed $($LiningCalcBlocks.Count) Lining Calculations" -ForegroundColor Green

            # Generate calculation blocks
            $AllChannelCalcBlock = "CHANNELCALCBLOCK`n" + ($ChannelCalcBlocks -join "`n") + "`nENDCHANNELCALCBLOCK"
            $AllLiningCalcBlock = if ($LiningCalcBlocks.Count -gt 0) {
                "LININGCALCBLOCK`n" + ($LiningCalcBlocks -join "`n") + "`nENDLININGCALCBLOCK"
            } else { "" }

            # Create the base HYD file structure
            $HydFileHeader = @"
HYDRAULICTOOLBOXPROJECT53

UNITS                0
NUMCALCS             115
TITLE                "$ProjectName - Ditches"
DESIGNER             ""
DATE                 $(Get-Date -Format "M d yyyy")
NOTES                "This file was generated on $(Get-Date -Format "MM/dd/yyyy") using 'HydraulicToolbox Utilities' Script owned by Jose Perez. Contact jbperezf@gmail.com for support."





"@

            $HydFileFooter = "ENDOFFILE"

            # Combine all content
            $FullFileContent = @(
                $HydFileHeader,
                $AllChannelCalcBlock,
                $AllLiningCalcBlock,
                $HydFileFooter
            ) -join "`n"

            # Write to temporary file
            $FullFileContent | Out-File $TempTextFile -Encoding UTF8

            # Save output file
            Write-ColorOutput "Step 2: Saving output file..." -ForegroundColor Yellow
            Move-Item $TempTextFile $OutputFilePath -Force

            # Success notification
            Write-ColorOutput "`nHYD file creation completed successfully!" -ForegroundColor Green
            Write-ColorOutput "Output file: $OutputFilePath" -ForegroundColor Green

            return $OutputFilePath
        }
        catch {
            # Improved error handling
            Write-ColorOutput "An error occurred during file creation:" -ForegroundColor Red
            Write-ColorOutput $_.Exception.Message -ForegroundColor Red
            
            # Log detailed error information
            Write-Error $_.Exception.ToString()
            
            throw
        }
        finally {
            # Ensure temporary file is always cleaned up
            if (Test-Path -Path $TempTextFile -ErrorAction SilentlyContinue) {
                Remove-Item $TempTextFile -Force
            }
        }
    }

    end {
        # Optional: Add any cleanup or final processing
    }
}

# Example usage with error handling
try {
    New-HydFile -CsvFilePath "input.csv"
    # Alternatively, provide the project name directly:
    # New-HydFile -CsvFilePath "input.csv" -ProjectName "Highway 123"
}
catch {
    Write-Host "Script execution failed: $_" -ForegroundColor Red
}