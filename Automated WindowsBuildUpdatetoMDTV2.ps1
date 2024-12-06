
param (
    [switch]$Elevated
)

function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

if ((Test-Admin) -eq $false) {
    if ($Elevated) {
        Write-Host "Failed to elevate permissions. Aborting." -ForegroundColor Red
    }
    else {
        Write-Host "Elevating permissions..."
        Start-Process powershell.exe -Verb RunAs -ArgumentList ("-NoProfile", "-NoExit", "-File", "`"$($myinvocation.MyCommand.Definition)`"", "-Elevated")
    }
    exit
}

Write-Host "Running with full privileges"


# Start the timer for tracking total execution time
$global:starttime = Get-Date

#Get the hostname and serial number of the device
#$global:Compname = "$env:COMPUTERNAME"
$global:userName = "$env:USERNAME"
#$global:SN = (get-wmiobject -class Win32_BIOS).serialnumber
$global:timeStamp = get-date -format yyyy-MM-dd_HH-mm-ss
$global:module = "C:\Program Files\Microsoft Deployment Toolkit\Bin\MicrosoftDeploymentToolkit.psd1"
Import-Module $global:module

# Set the log directory and check if it exists
$global:logDir = "\\path\to \log\Folder #\${global:userName}__$global:timeStamp.log"
$global:logFilePath = Join-Path -Path $global:logDir -ChildPath "msuImport_${global:userName}_$global:timeStamp.log"

if (!(Test-Path $global:logDir)) {
    # If the directory does not exist, create it
    New-Item -ItemType Directory -Force -Path $global:logDir
}


Test-Path $global:logDir

Start-Transcript -Path $global:logFilePath   #NoClobber
Write-Host "Starting transcript to: $global:logDir" 



function Open-KBLink {
    param(
        [string]$KBNumber
    )

    $userChoice = Read-Host "Would you like to open the update catalog to search for KB $KBNumber? (y for yes/n for no)"

    if ($userChoice -eq "y") {
        $url = "https://www.catalog.update.microsoft.com/Search.aspx?q=$KBNumber"
        Start-Process $url
        Write-Host "Opening the Update Catalog for KB $KBNumber in your browser..."
    } else {
        Write-Host "Skipping catalog link."
    }
}


# Prompt for the KB number
$kbNumber = Read-Host "Enter the KB number you want to search for"

# Open the Update Catalog based on the user's input
Open-KBLink -KBNumber $kbNumber


# Function to validate path existence
function Test-PathExists {
    param (
        [string]$Path
    )

    if (-not (Test-Path -Path $Path -PathType Container)) {
        $message = "Path '$Path' does not exist. Please provide a valid path."
        Write-Host $message -ForegroundColor Red
        #Log-Error $message
        return $false
    }
    return $true
}


# Function to display progress bar
function Show-ProgressBar {
    Write-Host "Running..." -NoNewline
    for ($i = 1; $i -le 100; $i++) {
        Write-Progress -Activity "Progress" -Status "$i% Complete" -PercentComplete $i
        Start-Sleep -Milliseconds 250
    }
    Write-Host " Complete!" -ForegroundColor Green
}

# Ensure the existence of required folders
$global:UserName = ([system.security.Principal.WindowsIdentity]::GetCurrent().Name).split('\')[1]
$global:winFolderPath = "\\path to where you want to temporarily copy the wim folder"
$global:msuFolderPath = "C:\Users\$UserName\Downloads"
$global:mountFolderPath = "C:\mountfolder1"
$global:deploymentshare = Read-Host -Prompt "Enter the full deploymentSharePath "


$requiredFolders = @($winFolderPath, $msuFolderPath, $mountFolderPath)
$requiredFolders | ForEach-Object {
    if (-not (Test-Path -Path $_ -PathType Container)) {
        try {
            Write-Host "Creating $_ folder..."
            New-Item -Path $_ -ItemType Directory | Out-Null
        } catch {
            $message = "Failed to create folder: $_. Error: $_"
            Write-Host $message -ForegroundColor Red
            #Log-Error $message
            exit
        }
    }
}

#############################################

# Predefined list of locations for selection
$wimLocations = @(
    "\\Wim file loaction 1",
    "\\Wim file loaction 2",
    "\\Wim file loaction 3",
    "\\Wim file loaction 4",
    "\Wim file loaction 5"
    
)
Write-Host ""

# Present options to the user
Write-Host "Select the .wim file location:"
for ($i = 0; $i -lt $wimLocations.Count; $i++) {
    Write-Host "$($i + 1). $($wimLocations[$i])"
}
Write-Host "$($wimLocations.Count + 1). Enter a custom location"
$selectedLocationIndex = [int](Read-Host -Prompt "Enter the number corresponding to the .wim file location you want to use")

Write-Host ""

# Validate user input
if ($selectedLocationIndex -lt 1 -or $selectedLocationIndex -gt ($wimLocations.Count + 1)) {
    $message = "Invalid selection. Please enter a valid number."
    Write-Host $message -ForegroundColor Red
    #Log-Error $message
    exit
}

if ($selectedLocationIndex -eq ($wimLocations.Count + 1)) {
    $wimfilelocation = Read-Host -Prompt "Enter the custom location path"
} else {
    $wimfilelocation = $wimLocations[$selectedLocationIndex - 1]
    Write-Host "Selected path: $wimfilelocation"
    Write-Host "Path exists: $(Test-Path $wimfilelocation)"
}

# Prompt for necessary inputs
do {
    if (-not (Test-Path -Path $mountFolderPath)) {
        $mountFolderPath = Read-Host -Prompt "Please enter a valid path where you want to mount the file"
    }
} until (Test-Path -Path $mountFolderPath)

##############################################################

# Function to check if the drive has enough free space
function Check-DriveSpace {
    param (
        [string]$Path,
        [int]$RequiredSpaceGB
    )

    $drive = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -eq ([System.IO.Path]::GetPathRoot($Path)) }
    if ($drive.Free -lt ($RequiredSpaceGB * 1GB)) {
        return $false
    }
    return $true
}

# Check if the drive has enough space to mount the OS
$requiredSpaceGB = 18 # Set the required space in GB
while (-not (Check-DriveSpace -Path $mountFolderPath -RequiredSpaceGB $requiredSpaceGB)) {
    Write-Host "The drive does not have enough free space to mount the OS. Please clear some space and press Enter to continue..." -ForegroundColor Red
    Read-Host
}

Write-Host ""

###################################################

# Process to add windows package to a wim file

###################################################


#

# Get the latest .wim file in the $wimfilelocation directory
$latestWimFile = Get-ChildItem -Path $wimfilelocation -Filter "*.wim" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

$wim_Filename = $latestWimFile.BaseName

Write-Host ""

Write-Host "baseWimFileName: $wim_Filename"

Write-Host ""

$latestBuild = Read-Host -Prompt "Rename file to the latest build: Enter the last four digit of the latest build (eg. 3456)"

Write-Host ""

#$BaseName = $wim_Filename -replace '(.{4})' , ''
$BaseName = $wim_Filename.Substring(0, $wim_Filename.Length - 4)
$global:Rename = $BaseName + $latestBuild

Write-Host ""

Write-Host " The updated Win filename will be: $global:Rename"

Write-Host ""

# Function to Display progress bar
Show-ProgressBar



try {
    # Copy .wim file to c:\win folder
    Copy-Item -Path $latestWimFile.FullName -Destination $winFolderPath -Force
    Write-Host "File copied "

	Write-Host ""

    # Mount the Windows image
    mount-windowsimage -ImagePath "$winFolderPath\$wim_Filename.wim" -Index 1 -Path $mountFolderPath
    Write-Host "Image mounted successfully"
	
	Write-Host ""

    try {
        # Add Windows package
        $latestMsuFile = Get-ChildItem -Path $msuFolderPath -Filter "*.msu" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

        if ($latestMsuFile) {
            $msuFilePath = $latestMsuFile.FullName
            Write-Host "Latest .msu file found: $($latestMsuFile.Name)"
			
			Write-Host ""
			
        } else {
            $message = "No .msu files found in the directory: $msuFolderPath"
            Write-Host $message -ForegroundColor Red
            #Log-Error $message
            exit
        }

        add-windowspackage -PackagePath $msuFilePath -Path $mountFolderPath
        Write-Host "Package added successfully"
		
		Write-Host ""
		
		$imageinfo = dism /image:"$global:mountFolderPath" /get-currentEdition
		$imageversion = $imageinfo | Select-String "Image Version" | ForEach-Object {$_.line.Split(':')[1].Trim()}
		$currentEd = $imageinfo | Select-String "Current Edition" | ForEach-Object {$_.line.Split(':')[1].Trim()}
				
		$global:imageDetails = [PSCustomObject]@{
			'Image Version' = $imageversion
			'Current Edition' = $currentEd
		}
		$global:updatedWInVersion = $global:imageDetails | Format-Table -AutoSize
		$global:updatedWInVersion
        # Dismount the Windows image
        Dismount-WindowsImage -Path $mountFolderPath -Save
        Write-Host "Image unmounted successfully"
		
		Write-Host ""

        # Rename the .wim file
        Rename-Item -Path "$winFolderPath\$wim_Filename.wim" -NewName "$global:Rename.wim"
        Write-Host "File renamed to $Rename"
		
		Write-Host ""

        $backupdestinationfolder = "$wimfilelocation"

        # Copy the renamed file to the backup folder
        Copy-Item -Path $winFolderPath\$Rename.wim -Destination $backupdestinationfolder

        Write-Host "$Rename copied to the backup folder"
		
		Write-Host ""
       

    } catch {
        $message = "An error occurred after mounting. Attempting to unmount the image... Error: $_"
        Write-Host $message -ForegroundColor Yellow
        #Log-Error $message

        try {
            Dismount-WindowsImage -Path $mountFolderPath -Save
            Write-Host "Image unmounted"
        } catch {
            $message = "Failed to unmount the image. Error: $_"
            Write-Host $message -ForegroundColor Red
            #Log-Error $message
        }
    }

} catch {
    $message = "An error occurred: $_"
    Write-Host $message -ForegroundColor Red
    #Log-Error $message
    exit
}

##################################################

# Process to import OS to MDT

#################################################


try {
    #function to imporOStoMDT
function ImportOStoMDT {
        try {
         $userChoice2 = "yes" #Read-Host "Would you like to Import the OS to MDT? (yes/no)"

         if ($userChoice2 -eq "yes") {
             # Load MDT module and map Deployment Share to PSDrive
            $module = "C:\Program Files\Microsoft Deployment Toolkit\Bin\MicrosoftDeploymentToolkit.psd1"
            $global:deploymentshare

            $global:Sources_Full = "$global:winFolderPath/$global:Rename.wim"  
            Test-Path "$global:winFolderPath/$global:Rename.wim" -verbose

            $global:Destination_folder = "$global:Rename"
			Import-Module $module

            # Function to find the next available PSDrive name in the DSxxx format
            function Get-NextPSDriveName {
                # Get all existing PSDrives with names matching the pattern DSxxxx
                $existingDrives = Get-PSDrive | Where-Object { $_.Name -match "^DS\d{3}$" }

                if ($existingDrives) {
                    # Find the highest number used in existing PSDrive names
                    $maxNum = $existingDrives |
                        ForEach-Object {
                            if ($_.Name -match "^DS(\d{3})$") { 
                                [int]$matches[1]
                            } else {
                                0  # In case of unexpected names (but should not happen in this case)
                            }
                        } | Sort-Object -Descending | Select-Object -First 1

                    # Increment the highest number found and format as DSxxxx
                    return "DS" + ($maxNum + 1).ToString("D3")
                } else {
                    # If no "DSxxxx" PSDrive exists, start with DS0001
                    return "DS001"
                }
            }

            # Check if there is an existing PSDrive mapped to the deployment share root path
            $existingPSDrive = Get-PSDrive | Where-Object { $_.Root -eq $deploymentshare -and $_.Provider.Name -eq "MDTProvider" }

            # Determine the PSDrive name
            $psDriveName = if ($existingPSDrive) {
                $existingPSDrive.Name
            } else {
                Get-NextPSDriveName
            }

            # If a PSDrive is found, remove it using its name
            if ($existingPSDrive) {
                Remove-PSDrive -Name $existingPSDrive.Name -Force
                Write-Host "The existing PSDrive '$($existingPSDrive.Name)' mapped to '$deploymentshare' has been removed."
            }

            # Create a new PSDrive with the determined name (sequential if necessary)
            New-PSDrive -Name ${psDriveName} -PSProvider MDTProvider -Root $deploymentshare
            Write-Host "The PSDrive '$psDriveName' has been mapped to your Deployment Share."

            # Loop to display menu and get user input
            [int]$Menu_OS = 0
            while ($Menu_OS -lt 1 -or $Menu_OS -gt 4) {
                Write-Host ""
                Write-Host "Update your Deploymentshare Choice"
                Write-Host "===================================================================="
                Write-Host "1. Add an Operating System with full set of source files"
                Write-Host "2. Add an Operating System from a captured image (WIM file)"
                Write-Host "3. Quit and exit"
                Write-Host ""

                # Use TryParse to safely convert the user input to an integer
                $userInput = "2" #Read-Host "Please enter an option 1 to 4"
                if ([int]::TryParse($userInput, [ref]$Menu_OS) -eq $false -or $Menu_OS -lt 1 -or $Menu_OS -gt 4) {
                    Write-Host "Invalid input, please enter a valid number between 1 and 4"
                    $Menu_OS = 0
                } else {
                    Write-Host "You selected: $Menu_OS"
                }
            }

            # Prompt for destination folder and check if it exists in the mounted OS folder
            do {
                #$Destination_folder = Read-Host -Prompt '***** Type your destination folder (Ex.4412)'
                $osMountPath = "${psDriveName}:\Operating Systems\$Destination_folder"

                if (Test-Path -Path $osMountPath) {
                    Write-Host "Folder '$Destination_folder' already exists in the mounted OS folder. Please choose a different name." -ForegroundColor Red
                }
            } until (-not (Test-Path -Path $osMountPath))

            # Use the switch statement to execute the corresponding option
            Switch ($Menu_OS) {
                1 {
                    Write-Host "Import OS with full set of sources files"
                    Write-Host "===================================================================="
                    Import-MDTOperatingSystem -Path "${psDriveName}:\Operating Systems" -SourcePath $Sources_Full -DestinationFolder $Destination_folder -Verbose
                }
                2 {
                    Write-Host "Import OS from a captured image (WIM file)"
                    Write-Host "===================================================================="
                    Import-MDTOperatingSystem -Path "${psDriveName}:\Operating Systems" -SourceFile $global:Sources_Full -DestinationFolder $global:Destination_folder -Verbose
                }
                3 {
                    Write-Host "Exiting script..."
                    exit
                }
            }
        } else {
            # Delete the .wim file in the C:\win folder
            Remove-Item -Path "$winFolderPath\$Rename.wim" -Force
            Write-Host "Copy of .wim file deleted from C drive"

            # Delete the .msu file in the Downloads folder
            Remove-Item -Path $msuFilePath -Force
            Write-Host "Copy of .msu file deleted from C drive"
        }
    } catch {
        $message = "An error occurred during MDT import: $_"
        Write-Host $message -ForegroundColor Red
        #Log-Error $message
        exit
    }
}
    ImportOStoMDT
   
    
} catch {
    $message = "Failed to import OS to MDT. Error: $_"
    Write-Host $message -ForegroundColor Red
    #Log-Error $message
    exit

}finally {

Write-Host ""
#######################################################################
#

# Delete the .wim file in the C:\win folder
 Remove-Item -Path "$winFolderPath\$Rename.wim" -Force
 Write-Host "Copy of .wim file deleted from $winFolderPath"
 Write-Host ""
 # Delete the .msu file in the Downloads folder
 Remove-Item -Path $msuFilePath -Force
 Write-Host "Copy of .msu file deleted from $msuFilePath"
 Write-Host ""
 
# End of script - Calculate total plan
$endtime = Get-Date
$Totalduration = $endtime - $starttime
Write-Host "Script completed in $($Totalduration.TotalMinutes) minutes."
Write-Host ""
Write-Host "*****************************************************************"

Write-Host "$global:Rename Uploaded to $global:deploymentshare : "
Write-Host "OSBuild information: "
Write-Host ""
$global:updatedWInVersion 
Write-Host ""

Write-Host "*****************************************************************"

# Stop the transcript
Stop-Transcript
Write-Host ""
Write-Host "Press Enter to Exit"
Read-Host
}
