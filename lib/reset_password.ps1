# Script to reset a user's password in Active Directory
param (
    [string]$username
)

# Import the Active Directory module
Import-Module ActiveDirectory


function generatePassword {
    param (
        [int]$length = 12
    )
    $chars = ('a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v',
            'w','x','y','z','A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S',
            'T','U','V','W','X','Y','Z','0','1','2','3','4','5','6','7','8','9','@','#','$','-')
    $password = -join ((1..$length) | ForEach-Object { $chars | Get-Random -Count 1 })
    return $password
}

function wordlistPassword {
    param (
        [string]$wordlistfile = "config\WL.txt"
    )
    
    if (-not (Test-Path -Path $wordlistFile)) {
        Write-Log "Wordlist file not found at $wordlistFile generatePassword function will be use" -ForegroundColor Red
        return $false
        }
    $wordlist = get-content -Path $wordlistFile
    $randomword = $wordlist | Get-Random -Count 1 
    $randomnumber = Get-Random -Minimum 100 -Maximum 999
    $password = "$randomword$randomnumber"
    return $password
}


function Write-Log {
    param (
        [string]$Message,
        [string]$LogFile = "logs\reset_password.log"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogFile -Value "$Timestamp - $Message"
}


try {
    $NewPassword = wordlistPassword 
    # if the wordlistPassword function returns false, generate a new password
    if (-not $NewPassword) {
        Write-Log "Wordlist password generation failed, generating a random password." -ForegroundColor Yellow
        $NewPassword = generatePassword
    }
    # Check if the user exists in Active Directory
    $user = Get-ADUser -Filter {samAccountName -eq $username } | Select -ExpandProperty samAccountName
    if ($null -eq $user) {
        Write-Log "User $username not found in Active Directory." -ForegroundColor Red
        return ("Failed", "User $username not found in Active Directory.")
        exit 1
    }

    # Reset the user's password
    Set-ADAccountPassword -Identity $user -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $NewPassword -Force)
    $displayName = Get-ADUser -Filter {samAccountName -eq $username } -Properties DisplayName | Select-Object -ExpandProperty DisplayName
    Write-Log "Password for user $username has been reset successfully." -ForegroundColor Green
    return ("success", "$displayName,$NewPassword")
} catch {
    Write-Log "An error occurred: $_" -ForegroundColor Red
    return ("Failed", "An error occurred:  $_")
}