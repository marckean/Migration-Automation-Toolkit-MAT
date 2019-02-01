##########################################
#      Migration Automation Script       #
#             version 1.5.2              #
#                                        #
#           ConvertVM.ps1                #
#                                        #
#         Copyright 2014 Microsoft       #
#                                        #
#      http:\\aka.ms\buildingclouds      #
##########################################

#BuildVersion = "1.5.2007"

###### Start Options ######
param (
    [Parameter(Position=0,Mandatory=$False, HelpMessage="Use .\ConvertVM.ps1 Help for script options")]
    [ValidateSet("Menu","Convert","Convert-Global","Collect","Create-List","Update","Show-VM", "Show-Status","Show-StatusLoop","Reset-List","Reset-Incomplete","Reset-All","Report-All","Report-Complete","Report-Incomplete","Report-Unsupported", "Report-Warning", "Purge","Update","Help")]
    [string]$Task = "Menu"
,   [Parameter(Position=1,Mandatory=$False)]
    [ValidateRange (0,86400)]
    [int]$DelayTimer = 0
)


###### Variables ######
$Queuelength = 2   ### Do not set Queuelength above 3 ###
$CurrentPath = "C:\MAT"
$XMLPath = "$CurrentPath\Variable.xml"
$VMList = "$CurrentPath\VMlist.txt"
$VerboseLogging = 3
$MaxCapacityLoopTimer = 60 
$Lookback = 1
$BuildVersion = "1.5.2"
$Catalog = "MAT"
$RemoteJobName = "MAT"
$TaskArgues = "$CurrentPath\ConvertVM.ps1 Convert"
$Computername = Get-Childitem env:computername 
$localhost = $Computername.Value
$Preflight = 1
$RemoveCD = 1
$RemoveFloppy = 1
$error.clear()

. $CurrentPath\ConvertVM-Functions.ps1    

##########################
###    Start Script    ###
##########################
cls
write-host "--------------------------------------------"
write-host "+        Migration Automation Script       +" 
write-host "+                version 1.5               +" 
write-host "+                                          +" 
write-host "+         Copyright 2014 Microsoft         +" 
write-host "+                                          +" 
write-host "+        http:\\aka.ms\buildingclouds      +"
write-host "--------------------------------------------"
write-host ""
sleep 1

# Elevate
Write-Host "Checking for elevation... " 
$Path = Get-Location
$CurrentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
if (($CurrentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) -eq $false)  
    {
    $ArgumentList = "-noprofile -noexit -file `"{0}`""
    If ($Task) {$ArgumentList = $ArgumentList + " $Task"}
    If ($DelayTimer) {$ArgumentList = $ArgumentList + " $DelayTimer"}
    Write-Host "Elevating..."
    Start-Process powershell.exe -Verb RunAs -ArgumentList ($ArgumentList -f ($myinvocation.MyCommand.Definition))
    Exit
    }





###### Verify XML file ######
If (Test-Path $XMLPath) 
    {
        Write-Host "Found $XMLPath"
        try {$Variable = [XML] (Get-Content $XMLPath)} 
        catch 
            {
            $Validate = $false;Write-Host "$XMLPath is invalid. Check XML syntax - Unable to proceed" -ForegroundColor Red
            exit
            }
    } 
Else 
    {
        $Validate = $false
        Write-Host "Missing $XMLPath - Unable to proceed" -ForegroundColor Red
        exit
    }


###### Load XML values as variables ######
Write-Host "Loading values from Variable.xml"
$Variable = [XML] (Get-Content "$XMLPath")
$Variable.MAT.General | ForEach-Object {$_.Variable} | Where-Object {$_.Name -ne $null} | ForEach-Object {Set-ScriptVariable -Name $_.Name -Value $_.Value}
$Variable.MAT.HyperV | ForEach-Object {$_.Variable} | Where-Object {$_.Name -ne $null} | ForEach-Object {Set-ScriptVariable -Name $_.Name -Value $_.Value}
$Variable.MAT.VMware | ForEach-Object {$_.Variable} | Where-Object {$_.Name -ne $null} | ForEach-Object {Set-ScriptVariable -Name $_.Name -Value $_.Value}
$MATLog = $logpath+"mat.log"

if ($Preflight -eq  1)
{Check-Preflight}

if ($MigrateNICs -eq 1)
{Check-MigrateNICs}

sleep 1
switch ($Task)
{
    "Menu" {Menu}
    "Convert-Global" {Start-Remotes}
    "Convert" 
        {
        Write-Log 1 $MATLog "---------------------------------------------"
        Write-Log 1 $MATLog "     Starting new batch of VMs               "
        Write-Log 1 $MATLog "---------------------------------------------"
        Write-Log 1 $MATLog "Logging events to $MATLog"
        Write-Log 1 $MATLog "Log Level = $VerboseLogging"
        Write-Log 7 $MATLog "Timer when Server is at Capacity = $MaxCapacityLoopTimer"
        Write-Log 1 $MATLog "Current Path = $CurrentPath"
        Write-Log 3 $MATLog "XML Path = $XMLPath"
        Write-Log 3 $MATLog "VMList = $VMList"
        Write-Log 3 $MATLog "Task Scheduler Job Name = $RemoteJobName"
        Write-Log 7 $MATLog "Task Arugements on Remotes = $TaskArgues"
        Write-Log 7 $MATLog "PowerShell Version = $BuildVersion"
        Write-Log 1 $MATLog "Delay Timer = $DelayTimer"
        Write-Log 3 $MATLog "Localhost = $Localhost"
        Write-Log 3 $MATLog "SQL Datasource = $Datasource"
        Write-Log 3 $MATLog "Preflight Checks = $Preflight"
        Write-Log 3 $MATLog "Migrate NICs = $MigrateNICs"
        Write-Log 3 $MATLog "Remove CD Drives = $RemoveCD"
        Write-Log 3 $MATLog "Remove Floppy drives = $RemoveFloppy"
        Write-Log 3 $MATLog "SQL Catalog = $Catalog"
        Write-Log 1 $MATLog "Max Concurrent Conversions = $Queuelength"

        ###### Use Delay Timer for executions in the future ######
        if ($DelayTimer -ge 1)
            {
            Write-Log 1 $MATLog "Delay Requested:True - Delaying execution for $DelayTimer seconds"
            Sleep -s $DelayTimer
            }
        else
            {
            Write-Log 3 $MATLog "Delay Requested:False - Starting immediate execution"
            }
        Convert
        }
    "Collect" {StartCollection}
    "Create-List" {List "USE $Catalog SELECT DisplayName FROM VMDetails_VIEW where [ReadytoConvert] = 0 AND [Completed] <> 1 AND [Supported] = 1 ORDER BY DisplayName"} 
    "Update" {UpdateDB}
    "Show-VM" {ShowVM}
    "Show-Status" {ShowStatus $Lookback}
    "Show-StatusLoop" 
        {
        do
            {
            ShowStatus $Lookback
            Sleep 5
            Write-Host ""
            Write-Host "Press 'Ctrl + C' to terminate script" -ForegroundColor DarkGray
            Sleep 5
            }
        while (1 -eq 1)
        }
    "Reset-List" {Reset List}
    "Reset-Incomplete" {Reset Incomplete}
    "Reset-All" {Reset All}
    "Report-Complete" {CSVReport "WHERE Completed = 1" $CurrentPath\Reports\Completed}
    "Report-Incomplete" {CSVReport "WHERE Completed <> 1 AND Supported = 1" $CurrentPath\Reports\Incomplete}
    "Report-Unsupported" {CSVReport "WHERE Supported = 3" $CurrentPath\Reports\Notsupported}
    "Report-Warning" {CSVReport "WHERE Warning = 1" $CurrentPath\Reports\Warning}
    "Report-All" {CSVReport $null $CurrentPath\Reports\AllData}
    "Purge" {Purge-DB}
    "Help" {Help}
  
}


