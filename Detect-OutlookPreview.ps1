#=============================================================================================================================
# Script Name:     Detect-OutlookPreview.ps1
# Description:     Detects if MS Outlook Preciew edition is present on the device
# Notes:           To be used as Intune Detection script for the above purpose
# Author:          James Barber
#=============================================================================================================================


If ($null -eq (Get-AppxPackage -Name Microsoft.OutlookForWindows -AllUsers)) {
    Write-Output 'Microsoft Outlook Preview App not present'
    Exit 0
}
Else {
    Write-Output 'Microsoft Outlook Preview App present'
    Exit 1
}