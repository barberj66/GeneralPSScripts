#=============================================================================================================================
# Script Name:     Detect-HideNewOutlookToggle.ps1
# Description:     Detects if the reg key exists, if it doesn't we need to create it via remediation script
# Notes:           To be used as Intune Detection script
# Author:          James Barber
#=============================================================================================================================
$regkey = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\General"
$regname = "HideNewOutlookToggle"

try
{
    $Result = (Get-ItemProperty $regkey -ErrorAction SilentlyContinue).PSObject.Properties.Name -contains $regname

    if (-not $Result){
        Write-Output "Reg key does not exist - we need to create"
        Exit 1
    } 
    else {
        Write-Output "Value Exists - no action required"
        Exit 0
    }
}
catch{
    $errMsg = $_.Exception.Message
    Write-Error $errMsg
    exit 1
}
