# Script to reset a user's password in Active Directory
param (
    [string]$username
)

# Import the Active Directory module
Import-Module ActiveDirectory

function generate-password {
    param (
        [int]$length = 14
    )
    $chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789@#$-"
    $password = -join ((1..$length) | ForEach-Object { $chars | Get-Random -Count 1 })
    return $password
}

function Write-Log {
    param (
        [string]$Message,
        [string]$LogFile = "reset_password.log"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogFile -Value "$Timestamp - $Message"
}


try {
    password = generate-password -length 14
    # Check if the user exists in Active Directory
    $user = Get-ADUser -Filter {samAccountName -eq $username } | Select -ExpandProperty samAccountName
    if ($null -eq $user) {
        Write-Log "User $username not found in Active Directory." -ForegroundColor Red
        return ("Failed", "User $username not found in Active Directory.")
        exit 1
    }

    # Reset the user's password
    Set-ADAccountPassword -Identity $user -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $NewPassword -Force)
    Write-Log "Password for user $username has been reset successfully." -ForegroundColor Green
    return ("success", "Password reset successfully.")
} catch {
    Write-Log "An error occurred: $_" -ForegroundColor Red
    return ("Failed", "An error occurred:  $_")
}