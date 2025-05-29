# Define the log file path and create the log directory if it doesn't exist
$logFolderPath = "C:\Windows\Logs\Scripts"
$logFilePath = "$logFolderPath\DeleteOldUserProfiles.log"

if (-not (Test-Path -Path $logFolderPath)) {
    New-Item -Path $logFolderPath -ItemType Directory -Force
}

# Set variable for current date
$currentDate = Get-Date

# Create a function to write log entries
function Write-Log {
    param (
        [string]$message        
    )
    $timestamp = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    $logMessage = "$timestamp - $message"
    Add-Content -Path $logFilePath -Value $logMessage
    Write-Output $logMessage
}

# Function to get the currently logged-in user
function Get-LoggedOnUsers {
    Get-CimInstance Win32_Process -Filter "name like 'explorer.exe'" | Invoke-CimMethod -MethodName GetOwner -ErrorAction SilentlyContinue | Select-Object -ExpandProperty User -Unique
}

# Get the currently logged-in user and store as variable to be added to exclusion list
$loggedOnUser = Get-LoggedOnUsers

# Function to get all user profiles

function Get-AllUserProfiles {
    $regProfilePath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    $exclusionList = @("Administrator*", "Public") + $loggedOnUser

    $allProfiles = Get-ChildItem "$regProfilePath" | Where-Object { $_.PSChildName -match 'S-1-5-21-' } | ForEach-Object {
        $sid = $_.PSChildName
        $profileKeyPath = "$regProfilePath\$sid"
        $profileProps = Get-ItemProperty -Path $profileKeyPath
        $profilePath = $profileProps.ProfileImagePath

        if ($null -ne $profilePath -and $profilePath -ne "") {
            $profileName = Split-Path -Leaf $profilePath
            $isExcluded = $false
            foreach ($exclusion in $exclusionList) {
                if ($profileName -like $exclusion) {
                    $isExcluded = $true
                    break
                }
            }

            if (-not $isExcluded) {
                $loadTimeHigh = $profileProps.LocalProfileLoadTimeHigh
                $loadTimeLow = $profileProps.LocalProfileLoadTimeLow

                try {
                    if ($null -ne $loadTimeHigh -and $null -ne $loadTimeLow) {
                        $loadTime = [System.DateTime]::FromFileTime(([int64]$loadTimeHigh -shl 32) -bor $loadTimeLow)
                    } else {
                        $loadTime = $null
                    }
                } catch {
                    Write-Log "Error converting load time for profile: $profilePath. Error: $_"
                    $loadTime = $null
                }

                if ($loadTime -eq [datetime]"1601-01-01T00:00:00") {
                    Write-Log "Default date found for profile: $profilePath. Using folder modification date instead."
                    $loadTime = $null
                }

                if ($loadTime) {
                    $daysSinceLastLogin = ($currentDate - $loadTime).Days
                } else {
                    $daysSinceLastLogin = "N/A"
                }

                [PSCustomObject]@{
                    UserSID = $sid
                    ProfilePath = $profilePath
                    LastLoadTime = $loadTime
                    DaysSinceLastLogin = $daysSinceLastLogin
                }
            } else {
                Write-Log "Profile $profilePath is in the exclusion list and will not be processed."
            }
        } else {
            Write-Log "Profile SID $sid has an empty or null ProfileImagePath and will be skipped."
        }
    }
    return $allProfiles
}


# Function to check folder modification date if DaysSinceLastLogin is N/A
function Get-LocalUsers {
    param (
        [PSCustomObject]$userProfile
    )
    if ($userProfile.DaysSinceLastLogin -eq "N/A") {
        $profilePath = $userProfile.ProfilePath
        if (Test-Path $profilePath) {
            $lastModified = (Get-Item $profilePath).LastWriteTime
            $daysSinceLastModified = ($currentDate - $lastModified).Days

            $userProfile.LastLoadTime = $lastModified
            $userProfile.DaysSinceLastLogin = $daysSinceLastModified

            Write-Log "Profile ${profilePath}: Last login time not found in registry. Using folder modification date: $lastModified, Days since last modification: $daysSinceLastModified"
        } else {
            Write-Log "Profile ${profilePath}: Folder not found."
        }
    }
    return $userProfile
}

# Function to take ownership of a folder
function Set-Ownership {
    param (
        [string]$path
    )
    try {
        $takeOwnCmd = "takeown /f `"$path`" /r /d y"
        $icaclsCmd = "icacls `"$path`" /grant administrators:F /t"
        cmd.exe /c $takeOwnCmd
        cmd.exe /c $icaclsCmd
        Write-Log "Took ownership of $path"
    } catch {
        Write-Log "Error taking ownership of $path. Error: $_"
    }
}

# Main function to delete profiles
function Remove-OldProfiles {
    $allProfiles = @(Get-AllUserProfiles)

    $allProfiles = $allProfiles | ForEach-Object {
        Get-LocalUsers -userProfile $_
    }

    $allProfiles | ForEach-Object {
        $profilePath = $_.ProfilePath
        $userSID = $_.UserSID
        $lastLoginTime = $_.LastLoadTime
        $daysSinceLogin = $_.DaysSinceLastLogin

        if ($daysSinceLogin -ge 10 -and $daysSinceLogin -ne "N/A") {
            if ($profilePath -match "C:\\Users\\$loggedOnUser") {
                Write-Log "Profile $profilePath is older than 10 days. Last login time was $lastLoginTime, $daysSinceLogin days since last login. Profile will not be removed as the user is currently logged in."
            } else {
                Write-Log "Profile $profilePath is older than 10 days. Last login time was $lastLoginTime, $daysSinceLogin days since last login. Attempting to remove profile."

                try {
                    Remove-Item -Path $profilePath -Recurse -Force -ErrorAction Stop
                    Write-Log "Deleted user folder: $profilePath"
                } catch {
                    Write-Log "Failed to delete profile with standard path: $profilePath. Will attempt with long path."

                    $longPath = "\\?\$profilePath"
                    try {
                        Remove-Item -Path $longPath -Recurse -Force -ErrorAction Stop
                        Write-Log "Deleted user folder using long path: $profilePath"
                    } catch {
                        Write-Log "Error deleting profile with long path: $profilePath. Error: $_.Exception.Message"

                        try {
                            Set-Ownership -path $profilePath
                            Remove-Item -Path $profilePath -Recurse -Force -ErrorAction Stop
                            Write-Log "Deleted user folder after taking ownership: $profilePath"
                        } catch {
                            Write-Log "Error deleting profile with standard path after taking ownership: $profilePath. Error: $_.Exception.Message"
                            try {
                                Remove-Item -Path $longPath -Recurse -Force -ErrorAction Stop
                                Write-Log "Deleted user folder using long path after taking ownership: $profilePath"
                            } catch {
                                Write-Log "Failed to delete profile: $profilePath after all attempts. Error: $_.Exception.Message"
                                continue
                            }
                        }
                    }
                }

                if (-not (Test-Path -Path $profilePath)) {
                    if ($userSID) {
                        $regProfilePath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$userSID"
                        try {
                            Remove-Item -Path $regProfilePath -Recurse -Force -ErrorAction Stop
                            Write-Log "Deleted registry entry: $regProfilePath"
                        } catch {
                            Write-Log "Failed to delete registry entry: $regProfilePath. Error: $_.Exception.Message"
                        }
                    }
                } else {
                    Write-Log "Failed to delete user folder: $profilePath"
                }
            }
        } else {
            Write-Log "Profile $profilePath is not older than 28 days so is not due for automatic deletion, last login time $lastLoginTime"
        }
    }
}

# Call the function to scan and remove old profiles
Remove-OldProfiles