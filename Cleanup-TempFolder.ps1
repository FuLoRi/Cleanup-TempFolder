# Load necessary assembly for Shell
Add-Type -AssemblyName Microsoft.VisualBasic

# Function to move items to Recycle Bin
Function MoveTo-RecycleBin {
    Param (
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    $Shell = New-Object -ComObject Shell.Application
    $Shell.NameSpace('shell:RecycleBinFolder').MoveHere($Path)
}

# Function to scan folder
Function Scan-Folder {
    Param (
        [Parameter(Mandatory = $true)]
        [string]$DirPath,

                
        [Parameter(Mandatory = $false)]
        [string[]]$OldFiles = @()
    )
    if ($VerboseOutput -eq $true) { Write-Host "VERBOSE - Scanning folder: $DirPath" }
    $Files = Get-ChildItem -Path $DirPath -File
    foreach ($File in $Files) {
        $FileAge = ($CurrentDate - $File.LastWriteTime).Days
        if ($DebugMode -eq $true) { Write-Host "  DEBUG -    Filename = $File" }
		if ($DebugMode -eq $true) { Write-Host "  DEBUG -    File age = $FileAge days" }
        if ($FileAge -gt $Age) {
            if ($VerboseOutput -eq $true) { Write-Host "VERBOSE -    $File is $FileAge days old and will be recycled." }
            $OldFiles += $File.FullName
            if ($DebugMode -eq $true) { Write-Host "  DEBUG - Old files found so far: $(($OldFiles).Count)" }
        }
    }
    $SubFolders = Get-ChildItem -Path $DirPath -Directory
    foreach ($SubFolder in $SubFolders) {
        $OldFiles = Scan-Folder -DirPath $SubFolder.FullName -OldFiles $OldFiles
    }
    # Return the updated array
    return $OldFiles
}

# Function to analyze a folder
Function Analyze-Folder {
    Param (
        [Parameter(Mandatory = $true)]
        [string]$FolderFullPath,
        
        [Parameter(Mandatory = $false)]
        [string[]]$OldFolders = @()
    )
    if ($VerboseOutput -eq $true) { Write-Host "VERBOSE - Analyzing folder: $FolderFullPath" }
    
    $AllOld = $true  # Assume all contents are old initially
    
    # Get subfolders
    $SubFolders = Get-ChildItem -Path $FolderFullPath -Directory
    
    # If there are subfolders, recurse into them first
    if ($SubFolders) {
        foreach ($SubFolder in $SubFolders) {
            $SubFolderAnalysis = Analyze-Folder -FolderFullPath $SubFolder.FullName -OldFolders $OldFolders
            $OldFolders = $SubFolderAnalysis.OldFolders
			if ($DebugMode -eq $true) { Write-Host "  DEBUG - After SubFolderAnalysis, OldFolders count = $(($OldFolders).Count)" }
			$SubFolderAllOld = $SubFolderAnalysis.AllOld
			if ($DebugMode -eq $true) { Write-Host "  DEBUG - After SubFolderAnalysis, SubFolderAllOld = $($SubFolderAllOld)" }
            if (-not $SubFolderAllOld) {
                $AllOld = $false  # If any subfolder is not old, set $AllOld to false
            }
        }
    }
    
    # Now analyze files in the current folder
    $Files = Get-ChildItem -Path $FolderFullPath -File
    foreach ($File in $Files) {
        $FileAge = ($CurrentDate - $File.LastWriteTime).Days
        if ($FileAge -le $Age) {
            $AllOld = $false  # If any file is new, set $AllOld to false
        }
    }

	# If all contents are old, mark the current folder as old
	if ($AllOld) {
		if ($VerboseOutput -eq $true) { Write-Host "VERBOSE -   Marking $FolderFullPath for recycling." }
		$OldFolders += $FolderFullPath
		if ($DebugMode -eq $true) { Write-Host "  DEBUG - After marking, OldFolders count = $(($OldFolders).Count)" }

		# Check if $FolderFullPath is in $OldFolders before proceeding
		if ($OldFolders -contains $FolderFullPath) {
			if ($VerboseOutput -eq $true) { Write-Host "VERBOSE -   Removing any previously marked subfolders of: $FolderFullPath" }
			
			# Unmark any redundant subfolders
			$OldFolders = $OldFolders | Where-Object {
				# Keep the current folder
				$_ -eq $FolderFullPath -or
				# Or any folder that is not a subfolder of the current folder
				-not ($_ -like "$FolderFullPath\*")
			}
			
			if ($DebugMode -eq $true) { Write-Host "  DEBUG - After unmarking, OldFolders count = $(($OldFolders).Count)" }
		}
	}

	# Create a result object to hold both the $AllOld flag and the updated $OldFolders array
    $result = @{
        AllOld = $AllOld
        OldFolders = $OldFolders
    }
    # Debugging output
    if ($DebugMode -eq $true) { Write-Host "  DEBUG - OldFolders array inside Analyze-Folder: $($OldFolders.Count)" }

    # Return the result object
    return $result
}

# Set defaults
$defaultVerboseOutput = $false
$defaultDebugMode = $false
$defaultTestMode = $false
$defaultPath = $env:TEMP
$defaultAge = 365  # 1 year in days

# Get and validate verbose output
do {
    $userVerboseOutput = Read-Host "Verbose output ($defaultVerboseOutput)"
    # If user input is empty, use defaults
    $VerboseOutput = if ($userVerboseOutput) { 
        $boolValid = [bool]::TryParse($userVerboseOutput, [ref]$null)
        if (-not $boolValid) {
            Write-Host "The verbose output value $userVerboseOutput is not a valid boolean (true or false). Please enter a valid value."
        }
        $boolValid
    } else { 
        $defaultVerboseOutput 
    }
} while (-not $VerboseOutput -and $userVerboseOutput)

# Get and validate debug mode
do {
    $userDebugMode = Read-Host "Debug mode ($defaultDebugMode)"
    # If user input is empty, use defaults
    $DebugMode = if ($userDebugMode) { 
        $boolValid = [bool]::TryParse($userDebugMode, [ref]$null)
        if (-not $boolValid) {
            Write-Host "The debug mode value $userDebugMode is not a valid boolean (true or false). Please enter a valid value."
        }
        $boolValid
    } else { 
        $defaultDebugMode 
    }
} while (-not $DebugMode -and $userDebugMode)

# Get and validate test mode
do {
    $userTestMode = Read-Host "Test mode ($defaultTestMode)"
    # If user input is empty, use defaults
    $TestMode = if ($userTestMode) { 
        $boolValid = [bool]::TryParse($userTestMode, [ref]$null)
        if (-not $boolValid) {
            Write-Host "The test mode value $userTestMode is not a valid boolean (true or false). Please enter a valid value."
        }
        $boolValid
    } else { 
        $defaultTestMode 
    }
} while (-not $TestMode -and $userTestMode)

# Get and validate path
do {
    $userPath = Read-Host "Enter the path ($defaultPath)"
    $Path = if ($userPath) { 
        if (-not (Test-Path $userPath)) {
            Write-Host "The path $userPath does not exist. Please enter a valid path."
            $null
        } else {
            (Get-Item -LiteralPath $userPath).FullName
        }
    } else { 
        $defaultPath 
    }
} while (-not $Path)

# Get and validate age
do {
    $userAge = Read-Host "Enter the age in days of files to be moved to Recycle Bin ($defaultAge)"
    if ([string]::IsNullOrWhiteSpace($userAge)) {
        $Age = $defaultAge
        break  # Exit the loop since a default value is used
    } else {
        $parsedAge = 0  # Initialize the variable
        $ageParsed = [int]::TryParse($userAge, [ref]$parsedAge)
        if (-not $ageParsed -or $parsedAge -le 0) {
            Write-Host "The age $userAge is not a valid positive integer. Please enter a valid age in days."
        } else {
            $Age = $parsedAge
            break  # Exit the loop since a valid value is assigned
        }
    }
} while ($true)  # Continue looping until a valid value is assigned or the default value is used

# Get the current date
$CurrentDate = Get-Date

# Arrays to hold old files and old folders
$OldFiles = @()
$OldFolders = @()

# Scan phase
Write-Host "Scan phase initiated..."
$OldFiles = Scan-Folder -DirPath $Path

Write-Host "Found $(($OldFiles).count) old file(s).`n"

# Prompt user to proceed to Analyze phase
$Proceed = Read-Host "Proceed to Analyze phase? (Y/N)"
if ($Proceed -ne "y" -or $Proceed -ne "Y") { exit 0 }

# Analyze phase to identify old folders
Write-Host "Analyze phase initiated..."
$FolderAnalysis = Analyze-Folder -FolderFullPath $Path

# Update the $OldFolders global variable with the array returned by Analyze-Folder
$global:OldFolders = $FolderAnalysis.OldFolders

# Debugging output
if ($DebugMode -eq $true) { Write-Host "  DEBUG - OldFolders array outside Analyze-Folder: $($global:OldFolders.Count)" }

# If you need to use the AllOld flag, you can access it like this:
$AllOld = $FolderAnalysis.AllOld

Write-Host "Found $(($global:OldFolders).count) old/empty folder(s).`n"

# Prompt user to proceed to Recycle phase
$Proceed = Read-Host "Proceed to Recycle phase? (Y/N)"
if ($Proceed -ne "y" -or $Proceed -ne "Y") { exit 0 }

# Recycle phase
Write-Host "Recycle phase initiated..."

# Check to see if there are any folders to be recycled
if (($global:OldFolders).Count -gt 0) {
	
	# Show the user a list of all of the old folders
	if ($VerboseOutput -eq $true) { 
		Write-Host "Old/Empty Folders to be recycled:"
		$global:OldFolders | ForEach-Object { Write-Host "   $_" }
	}

	# Prompt the user to confirm before recycling old folders
	$recycleFolders = Read-Host "Really recycle $(($global:OldFolders).count) old/empty folder(s)? (Y/N)"
	# If user confirms, proceed with recycling old folders
	if ($recycleFolders -eq "y" -or $recycleFolders -eq "Y") {
		foreach ($OldFolder in $global:OldFolders) {
			if ($VerboseOutput -eq $true) { Write-Host "VERBOSE - Recycling folder: $OldFolder" }
			if ($TestMode -ne $true) { MoveTo-RecycleBin -Path $OldFolder }
			if ($DebugMode -eq $true) { Write-Host "  DEBUG - OldFiles array before pruning: $($OldFiles.Count)" }
			# Remove files in this folder from the OldFiles array
			<# $OldFiles = $OldFiles | Where-Object { -not $_.StartsWith($OldFolder) } #>
			$OldFiles = $OldFiles | Where-Object { -not $_.StartsWith($OldFolder, [System.StringComparison]::OrdinalIgnoreCase) }
			if ($DebugMode -eq $true) { Write-Host "  DEBUG - OldFiles array after pruning: $($OldFiles.Count)" }
		}
		Write-Host "Recycled $(($global:OldFolders).count) old/empty folder(s).`n"
	}
}
elseif (($global:OldFolders).Count -eq 0) {
	Write-Host "No Old/Empty Folders to be recycled.`n"
}
else {
	Write-Host "Oopsie! The count of old/empty folders is neither greater than nor equal to zero. This shouldn't happen!`n"
	exit 1
}

# Check to see if there are any files to be recycled
if (($OldFiles).Count -gt 0) {

	# Show the user a list of all of the old files
	if ($VerboseOutput -eq $true) { 
		Write-Host "Old Files to be recycled:"
		$OldFiles | ForEach-Object { Write-Host "   $_" }
	}

	# Prompt the user to confirm before recycling old files
	$recycleFiles = Read-Host "Really recycle $(($OldFiles).count) old/empty file(s)? (Y/N)"
	# If user confirms, proceed with recycling old files
	if ($recycleFiles -eq "y" -or $recycleFiles -eq "Y") {
		foreach ($OldFile in $OldFiles) {
			if ($VerboseOutput -eq $true) { Write-Host "VERBOSE - Recycling file: $OldFile" }
			if ($TestMode -ne $true) { MoveTo-RecycleBin -Path $OldFile }
		}
		Write-Host "Recycled $(($OldFiles).count) old file(s).`n"
	}
}
elseif (($OldFiles).Count -eq 0) {
	Write-Host "No Old Files to be recycled.`n"
}
else {
	Write-Host "Oopsie! The count of old files is neither greater than nor equal to zero. This shouldn't happen!`n"
	exit 1
}

Write-Host "Done."
$Exit = Read-Host "Press Enter to finish"