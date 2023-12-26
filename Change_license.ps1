<#
=============================================================================================
Name:           Change license Office 365 users
Author:         Pedro Rodrigues > cast.pedroh.utic@sebrae.com.br | phsr2001@gmail.com | ophsr.com.br
============================================================================================
#>


Function ConnectMgGraphModule {
    $MsGraphModule = Get-Module Microsoft.Graph -ListAvailable
    if ($MsGraphModule -eq $null) { 
        Write-host "Important: Microsoft Graph PowerShell module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        $confirm = Read-Host Are you sure you want to install Microsoft graph module? [Y] Yes [N] No  
        if ($confirm -match "[yY]") { 
            Write-host "Installing Microsoft Graph PowerShell module..."
            Install-Module Microsoft.Graph -Scope CurrentUser
            Write-host "Microsoft Graph PowerShell module is installed in the machine successfully" -ForegroundColor Magenta
        } 
        else { 
            Write-host "Exiting. `nNote: Microsoft Graph PowerShell module must be available in your system to run the script" -ForegroundColor Red
            Exit 
        } 
    }

    Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All"  -ErrorVariable ConnectionError | Out-Null
    if ($ConnectionError -ne $null) {    
        Write-Host $ConnectionError -Foregroundcolor Red
        Exit
    }
    else {
        Write-Host "Microsoft Graph PowerShell module is connected successfully" -ForegroundColor Green
    }
}

function Choice {
    $ValidInput = @('1', '2', '3')
    $response = ''
    $filePath = ''


    while ($filePath -eq '') {
        $filePath = Read-Host "Path of the file containing the users that will be modified (Ex: c:\temp\user1.csv)" 
        $testpath = Test-Path -Path $filePath -PathType Leaf
        if (-not $testpath) {
            [console]::Beep(1000, 100)
            write-host "$filePath is not valid" -ForegroundColor Red
            $filePath = ''
        }
    }
    
    Write-Host "`r`nWhich license do you want to ADD and REMOVE in the list users ($filePath)"-ForegroundColor DarkYellow
    Write-Host "## 1- ADD M365 E3 and REMOVE Office 365 E3`r`n## 2- ADD M365 F3 and REMOVE Office 365 F3`r`n## 3- ADD Office 365 E1 and REMOVE Office 365 E3`r`n" -ForegroundColor DarkYellow
    while ($response -eq '') {
        $response = Read-Host -Prompt "    Enter option (1, 2 or 3)"
        if ($response -notin $ValidInput) {
            [console]::Beep(1000, 100)
            $response = ''
        }
    }
    Switch ($response) {
        1 {
            Write-Log INFO "ADD M365 E3 and REMOVE Office 365 E3"
            ChangeLicense ChangeForM365_e3 $filePath
        }
        2 {
            Write-Log INFO "ADD M365 F3 and REMOVE Office 365 F3"
            ChangeLicense ChangeForM365_f3 $filePath
        }
        3 {
            Write-Log INFO "ADD Office 365 E1 and REMOVE Office 365 E3"
            ChangeLicense ChangeForM365_E3_E1_TO $filePath
        }
    }
}

function ChangeLicense {
    Param
    (
        [Parameter(Mandatory = $true)][ValidateSet("ChangeForM365_e3", "ChangeForM365_f3", "ChangeForM365_E3_E1_TO")]
        [string]$Type,
        [Parameter(Mandatory = $true)][string]$FilePath
    )
    $Office_f3 = Get-MgSubscribedSku -All | Where SkuPartNumber -eq 'DESKLESSPACK'
    $Office_e3 = Get-MgSubscribedSku -All | Where SkuPartNumber -eq 'ENTERPRISEPACK'
    $Office_e1 = Get-MgSubscribedSku -All | Where SkuPartNumber -eq 'STANDARDPACK'
    $Microsoft_f3 = Get-MgSubscribedSku -All | Where SkuPartNumber -eq 'SPE_F1'
    $Microsoft_e3 = Get-MgSubscribedSku -All | Where SkuPartNumber -eq 'SPE_E3'
    $UpnList = Import-Csv $FilePath ";"
    $UpnListCount = ($UpnList | Measure-Object).Count
    $counter = 0
    foreach ($upn in $UpnList) {
        $counter++
        Write-Progress -Activity 'Processing...' -CurrentOperation $upn.UPN -PercentComplete (($counter / $UpnListCount) * 100)
        if ($Type -eq "ChangeForM365_e3") {
            #Unassign Office E3
            Set-MgUserLicense -UserId $upn.UPN -AddLicenses @{} -RemoveLicenses @($Office_e3.SkuId) -ErrorVariable ConnectionError | Out-Null 
            if ($ConnectionError -ne $null) {    
                Write-Log ERROR "$upn -> $ConnectionError"
            }
            else {
                Write-Log INFO "$upn - Remove Office E3 successfully" 
            }
            #Assign M365 E3
            Set-MgUserLicense -UserId $upn.UPN -AddLicenses @{SkuId = $Microsoft_e3.SkuId }  -RemoveLicenses @() -ErrorVariable ConnectionError | Out-Null
            if ($ConnectionError -ne $null) {    
                Write-Log ERROR "$upn -> $ConnectionError"
            }
            else {
                Write-Log INFO "$upn - Add Microsoft E3 successfully" 
            }
        }
        if ($Type -eq "ChangeForM365_f3") {
            # Unassign Office f3
            Set-MgUserLicense -UserId $upn.UPN -AddLicenses @{} -RemoveLicenses @($Office_f3.SkuId) -ErrorVariable ConnectionError | Out-Null 
            if ($ConnectionError -ne $null) {    
                Write-Log ERROR "$upn -> $ConnectionError"
            }
            else {
                Write-Log INFO "$upn - Remove Office F3 successfully"
            }
            #Assign M365 f3
            Set-MgUserLicense -UserId $upn.UPN -AddLicenses @{SkuId = $Microsoft_f3.SkuId }  -RemoveLicenses @() -ErrorVariable ConnectionError | Out-Null 
            if ($ConnectionError -ne $null) {    
                Write-Log ERROR "$upn -> $ConnectionError"
            }
            else {
                Write-Log INFO "$upn - Add Microsoft F3 successfully" 
            }
        }
        if ($Type -eq "ChangeForM365_E3_E1_TO") {
            # Unassign Office e3
            Set-MgUserLicense -UserId $upn.UPN -AddLicenses @{} -RemoveLicenses @($Office_e3.SkuId) -ErrorVariable ConnectionError | Out-Null 
            if ($ConnectionError -ne $null) {    
                Write-Log ERROR "$upn -> $ConnectionError"
            }
            else {
                Write-Log INFO "$upn - Remove Office E3 successfully"
            }
            #Assign M365 e1
            Set-MgUserLicense -UserId $upn.UPN -AddLicenses @{SkuId = $Office_e1.SkuId }  -RemoveLicenses @() -ErrorVariable ConnectionError | Out-Null 
            if ($ConnectionError -ne $null) {    
                Write-Log ERROR "$upn -> $ConnectionError"
            }
            else {
                Write-Log INFO "$upn - Add Office E1 successfully" 
            }
        }
    }
}


Function Write-Log {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $False)]
        [ValidateSet("INFO", "WARN", "ERROR", "FATAL", "DEBUG")]
        [String]
        $Level = "INFO",

        [Parameter(Mandatory = $True, ValueFromPipeline)]
        [string]
        $Message,

        [Parameter(Mandatory = $False)]
        [string]
        $logfile = 'log_change_licence.log'
    )

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp $Level $Message"
    If ($logfile) {
        Add-Content $logfile -Value $Line
    }
    Else {
        Write-Output $Line
    }
}
Function main() {
    ConnectMgGraphModule
    Choice
}
. main
