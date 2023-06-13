#=============================================================================================================================
# Script Name:     Remediate-OutlookPreview.ps1
# Description:     If Outlook Preview is present this script will remove it
# Notes:           To be used as Intune Detection script for the above purpose
# Author:          James Barber
#=============================================================================================================================
If ($null -eq (Get-AppxPackage -Name Microsoft.OutlookForWindows -AllUsers)) {
    Write-Output 'Microsoft Outlook Preview App not present'
}
Else {
    Try {
        Write-Output 'Removing Microsoft Outlook Preview App'
        If (Get-Process -name olk -ErrorAction SilentlyContinue) {
            Try {
                Write-Output 'Stopping Microsoft Outlook Preview app process'
                Stop-Process -name olk -Force
                Write-Output "Stopped"
            }
            catch {
                Write-Output 'Unable to stop process, trying to remove anyway'
            }
           
        }
        Get-AppxPackage -Name Microsoft.OutlookForWindows -AllUsers | Remove-AppPackage -AllUsers
        Write-Output 'Microsoft Outlook Preview App removed successfully'
    }
    catch {
        Write-Error 'Error removing Microsoft Outlook Preview App'
    }
}