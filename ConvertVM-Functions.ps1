##########################################
#      Migration Automation Script       #
#             version 1.5.2              #
#                                        #
#       ConvertVM-Functions.ps1          #
#                                        #
#         Copyright 2014 Microsoft       #
#                                        #
#      http:\\aka.ms\buildingclouds      #
##########################################

#BuildVersion = "1.5.2007"

##################################################################
### FUNCTIONS ###
##################################################################

################################
###### General Region     ######
################################
#region General

###### Create Variables from XML ######
Function Set-ScriptVariable ($Name,$Value) 
 
{
Invoke-Expression ("`$Script:" + $Name + " = `"" + $Value + "`"")
If (($Name.Contains("Account")) -and !($Name.Contains("Password")) -and ($Value -ne "")) 
    {
    Invoke-Expression ("`$Script:" + $Name + "Domain = `"" + $Value.Split("\")[0] + "`"")
    Invoke-Expression ("`$Script:" + $Name + "Username = `"" + $Value.Split("\")[1] + "`"") 
    }
}

###### Write-Log ######
function write-log([int]$level, $logname, [string]$info, $color = "white")
{
    if ($level -le $VerboseLogging)
    {
        $time = get-date
        add-content $logname -value "$time - $info" 
        if ($color -eq "Red" -OR $color -eq "Yellow")
        {Write-Host "$time - $info" -ForegroundColor $color -BackgroundColor Black}

        else 
        {Write-Host "$time - $info" -ForegroundColor $color}
    }
} 

###### SQLWrite ######
function SQLWrite([string]$Query)
{
    $conn = New-Object System.Data.SqlClient.SqlConnection("Data Source=$DataSource; Initial Catalog=$Catalog; Integrated Security=SSPI")
	$conn.Open()
	$cmd = $conn.CreateCommand()
	$cmd.CommandText = "$Query"
	$cmd.ExecuteNonQuery() |Out-Null
	$conn.Close()

    if ($VerboseLogging -ge 7)
        {Write-Log 7 $MATLog "$Query"}
}
#endregion

#################################
###### Preflight Region    ######
#################################
#region Preflight

###### Verify Preflight checks ######
function Check-Preflight 
{
Write-Host "Running Preflight checks..."
$PSver = $PSVersionTable.PSVersion.Major
   if ($PSver -le 2)
        {
        Fail-Preflight "Powershell version 3 or higher is required."
        }

$Ver1 = $BuildVersion.CompareTo($XMLVersion)
Write-Log 7 $MATLog "Powershell version: $BuildVersion"
Write-Log 7 $MATLog "XML version: $XMLVersion" 
   if ($ver1 -ne 0)
        {
        Fail-Preflight "Powershell and XML versions DO NOT match. Exiting.", "Version incompatability, please download the MAT again."
        }

Try 
{
    $Query = "USE $Catalog SELECT TOP 1 [DBVersion] FROM [MAT].[dbo].[VERSION]"
    $conn = New-Object System.Data.SqlClient.SqlConnection("Data Source=$DataSource; Initial Catalog=$Catalog; Integrated Security=SSPI")
    $conn.Open()
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = "$Query"
    $Reader = $cmd.ExecuteReader()
    while ($Reader.Read())
	    {$DBVersion  = $Reader.GetValue(0)}
}
Catch
{
    Fail-Preflight "ERROR - DATABASE ERROR: $($_.exception.message)" 
}

$Ver2 = $BuildVersion.CompareTo($DBVersion)
Write-Log 7 $MATLog "Powershell version: $BuildVersion"
Write-Log 7 $MATLog "Database version: $DBVersion" 
   if ($ver2 -ne 0)
        {
        Fail-Preflight "Powershell and Database versions DO NOT match." "Version incompatability, please download the MAT again."
        }
}

###### Verify Startup Power States ######
function Check-MigrateNICs
{

if ([Environment]::Is64BitProcess)
    {
    Fail-Preflight "You must use 32-bit PowerShell if you want to use the MigrateNIC option." 
    }

if ($spower -ne 0)
    {
    Fail-Preflight "You have enabled the MigrateNIC feature but you have a conflicting setting." "The spower variable must = 0"
    }

}

###### Preflight Failure ######
function Fail-Preflight ($errormsg1, $errormsg2="Exiting.")

{
    Write-Host
    Write-Log 1 $MATLog "ERROR - Preflight failed. Critical Error." Red
    Write-Log 1 $MATLog "$errormsg1" Red
    Write-Log 1 $MATLog "$errormsg2" Red
    Write-Host   
    exit
}

#endregion

################################
###### Task Sched Section ######
################################
#region TaskSched
###### Connect to Task Service ######
Function Connect-TaskService ($ConvertServer, $WinAccountUsername, $WinAccountDomain, $WinPassword)
{
$Global:TaskService = New-Object -ComObject Schedule.Service
$TaskServiceConnected = $false
$TaskServiceTries = 5
$TaskServiceInterval = 60
$i = 0
While ((!($TaskServiceConnected)) -and ($TaskServiceTries -ne $i)) {
    $i++
    Write-host "Connecting to task service attempt $i"
    $Global:TaskService.Connect($ConvertServer,"$WinAccountUsername","$WinAccountDomain","$WinPassword")
    If (!(($Global:TaskService.Connected -eq "True") -and ($Global:TaskService.TargetServer -eq $ConvertServer))) {
        If ($i -eq $TaskServiceTries) {
            Fail -Server $ConvertServer
            Write-host "Failed to connect to task service"
        } Else {
            Write-host "Failed to connect to task service"
            Write-host "Waiting $TaskServiceInterval seconds" 
            Start-Sleep $TaskServiceInterval
        }
    } Else {
        Write-log 3 $MATLog "Connected to task service on $ConvertServer"
        $TaskServiceConnected = $true
    }
}
$Global:TaskFolder = $Global:TaskService.GetFolder("\")
Write-Log 3 $MATLog "Task Folder $Global:TaskFolder"
}
            
###### Connect to Task Service ######
Function Get-Task ($TaskName) 
{
    $Tasks = $Global:TaskFolder.GetTasks(0)
    Return $Tasks | Where-Object {$_.Name -eq $TaskName}
    Write-Log 3 $MATLog "Found $Task on $ConvertServer"
}

###### Create a new task ######
Function New-Task ($TaskName,$Command,$Arguments,$User,$Password) 
{
    $Task = $Global:TaskService.NewTask(0)
    $TaskPrincipal = $Task.Principal
    $TaskPrincipal.RunLevel = 1
    $TaskAction = $Task.Actions.Create(0)
    $TaskAction.Path = $Command
    $TaskAction.Arguments = $Arguments
    Write-Log 1 $MATLog "Registering task $TaskName on $ConvertServer"
    $Global:TaskFolder.RegisterTaskDefinition($TaskName,$Task,6,$User,$Password,1)
}
 
###### Start Task ######
Function Start-Task ($TaskName) {
    # Start a task
    $Task = $Global:TaskFolder.GetTask($TaskName)
    Write-Log 1 $MATLog "Starting task $TaskName on $ConvertServer"
    $Task.Run(0)
}

###### Remove Task ######
Function Remove-Task ($TaskName) {
    # Remove a task
    Write-Log 1 $MATLog "Removing task: $TaskName from $ConvertServer"
    $Global:TaskFolder.DeleteTask($TaskName,0)
}
#endregion

################################
###### Conversion Region  ######
################################
#region Conversion

######  Start remote conversions ######
function Start-Remotes
{
$Variable.MAT.SchedCreds | ForEach-Object {$_.Variable} | Where-Object {$_.Name -ne $null} | ForEach-Object {Set-ScriptVariable -Name $_.Name -Value $_.Value}
Write-Log 3 $MATLog "--------------Looking for Remote Servers to help with conversion"
$RemoteCount = $Variable.MAT.RemoteHost.ConvertServer.Count
    if ($RemoteCount -ge 1)
        {
        Write-Log 1 $MATLog "Found $Remotecount Remote Servers in $XMLPath"
        $ConvertHost = $Variable.MAT.RemoteHost.ConvertServer
        foreach ($ConvertServer in $ConvertHost) 
            {
                Write-Host $ConvertHost
                Connect-TaskService $ConvertServer $WinAccountUsername $WinAccountDomain $WinPassword
                    if (Get-Task $RemoteJobName)
                    {Remove-Task $RemoteJobName}
                New-Task $RemoteJobName %SystemRoot%\syswow64\WindowsPowerShell\v1.0\powershell.exe $TaskArgues $WinAccount $WinPassword
                Start-Task $RemoteJobName
                # Add delay to let each server have a turn at the queue
                Write-Log 1 $MATLog "Waiting 15 seconds to let server have a turn at the queue"
                Sleep 15
                ShowStatus $Lookback
                sleep 5
            }
         }
    else {Write-Log 1 $MATLog "No Remote Servers found"}
Convert
}

######  Start local conversion ######
function Convert 
{
######  Verify Capacity ######
CheckSemaphore
    while ($Semaphore -ge $Queuelength) 
    {
    Write-Log 1 $MATLog "(Capacity Check) <$localhost> Server is at capacity. Waiting for next free slot" Cyan
    Sleep 5
    ShowStatus $Lookback
    Write-Host "Waiting 30 seconds and then rechecking"
    sleep $MaxCapacityLoopTimer
    cls
    CheckSemaphore
    }
$Slots = $Queuelength - $Semaphore
Write-Log 1 $MATLog "(Capacity Check) <$localhost> $Slots slot(s) available"

######  Check Queue for VMs ######
GetUnassignedVMCount
Write-Log 3 $MATLog "Unassigned VM Count is $UnassignedVMCount"
    if ($UnassignedVMCount -ge 1)
        {
            GetNextVM
        }

    else
        {
        Write-Log 1 $MATLog "Done. There are no unassigned VMs marked as ReadyToConvert in the queue"
        Write-Log 1 $MATLog "***** End of batch conversions *****" Cyan
        CheckSemaphore
        if ($Semaphore -ge 1)
            {
            Write-Log 1 $MATLog "$Semaphore Conversions are still running but this script will exit."
            Write-Log 1 $MATLog "You can continue monitor using the Show-StatusLoop parameter."
            }
        else
            {Write-Log 1 $MATLog "No more conversions are running locally. Exiting."}
        exit
        }
}

######  Check Server capacity ######
function CheckSemaphore
{
$Query = "USE $Catalog SELECT Count(*) FROM VMQueue WHERE [InUse] = 1 AND ConvertServer = '$localhost'"

	$conn = New-Object System.Data.SqlClient.SqlConnection("Data Source=$DataSource; Initial Catalog=$Catalog; Integrated Security=SSPI")
	$conn.Open()
	$cmd = $conn.CreateCommand()
	$cmd.CommandText = "$Query"
	$Reader = $cmd.ExecuteReader()
	while ($Reader.Read())
	    {$global:Semaphore  = $Reader.GetValue(0)}
	$conn.Close()

Write-Log 3 $MATLog "(Capacity Check) <$localhost> $Semaphore of $Queuelength in use."
}

###### Get Unassigned VMs ######
function GetUnassignedVMCount
{
Write-Log 3 $MATLog  "Checking Global VM Queue for VMs that are unassigned"
$Query = "USE $Catalog SELECT Count(*) FROM VMQueue WHERE [ReadytoConvert] = 1 AND ([ConvertServer] IS NULL)"

	$conn = New-Object System.Data.SqlClient.SqlConnection("Data Source=$DataSource; Initial Catalog=$Catalog; Integrated Security=SSPI")
	$conn.Open()
	$cmd = $conn.CreateCommand()
	$cmd.CommandText = "$Query"
	$Reader = $cmd.ExecuteReader()
	while ($Reader.Read())
	    {$global:UnassignedVMCount  = $Reader.GetValue(0)}
	$conn.Close()

Write-Log 1 $MATLog "Found $UnassignedVMCount VMs ready for conversion" Cyan
}

###### Get next VM Info ######
function GetNextVM
{
Write-Log 3 $MATLog "<$localhost> Getting next VM from database"
$Query = "USE $Catalog SELECT TOP 1 JobID, VMName, DisplayName FROM VMDetails_VIEW WHERE [ReadytoConvert] = 1 AND ([ConvertServer] IS NULL) ORDER by Position"

	$conn = New-Object System.Data.SqlClient.SqlConnection("Data Source=$DataSource; Initial Catalog=$Catalog; Integrated Security=SSPI")
	$conn.Open()
	$cmd = $conn.CreateCommand()
	$cmd.CommandText = "$Query"
	$Reader = $cmd.ExecuteReader()
	while ($Reader.Read())
	    {
		$ID  = $Reader.GetValue(0)
		$VMName = $Reader.GetValue(1)
        $DisplayName = $Reader.GetValue(2)
	    }
	$conn.Close()

###### Kick off Convert.ps1 ######
SQLWrite "UPDATE VMQueue SET [ConvertServer] = '$localhost', [Inuse] = 1, [ReadytoConvert] = 0, [Position] = NULL, [Summary] = 'Submitting for Conversion' WHERE JobID = $ID" 

Write-Log 1 $MATLog "Starting BackupVM Network" 
PrepareVM $DisplayName

$StartConvert = Start-Process -FilePath powershell.exe -ArgumentList "$CurrentPath\ConvertVM-Process.ps1 $ID $Datasource $Catalog $MATLog $DisplayName $VMName $Logpath $CurrentPath $VerboseLogging $XMLPath $MigrateNICs" -PassThru -WindowStyle Minimized
Write-Log 7 $MATLog "Start-Process -FilePath powershell.exe -ArgumentList '$CurrentPath\ConvertVM-Process.ps1 $ID $Datasource $Catalog $DisplayName $VMName $Logpath $CurrentPath $VerboseLogging $XMLPath $MigrateNICs' -PassThru -WindowStyle Minimized"
Write-Log 3 $MATLog "Waiting 10 seconds to let process settle"  
Sleep 10

  # IF not the last VM loop through GetVMInfo again. (lather, rinse, repeat.) #
Write-Log 1 $MATLog "Launched $VMName conversion" Cyan
Write-Log 1 $MATLog "Checking Global list for more unassigned VMs"
Convert
}

###### Prepare for conversion #####
Function PrepareVM ($VMTarget)
{
    #Split $guserDomain and $guserName
    Invoke-Expression ("`$Script:" + "guser" + "Domain = `"" + $guser.Split("\")[0] + "`"")
    Invoke-Expression ("`$Script:" + "guser" + "Name = `"" + $guser.Split("\")[1] + "`"") 
 
    SQLWrite "UPDATE VMQueue SET [Summary] = 'Backing up VM NIC info' WHERE JobID = $ID"  
    Write-Log 1 $MATLog "Backing up VM Network function for $VMTarget" 
    Write-Log 3 $MATLog "Adding Snapin VMware.VimAutomation.Core"
    Add-PSSnapin "VMware.VimAutomation.Core"
    

    Write-Log 1 $MATLog "Connecting to VMware Server" 

    $ConnectVIServer = Connect-VIServer -Server $shost -Protocol https -User $shusername -Password $shpwd
    $VIPort = $ConnectVIServer.Port
    $VIVer = $ConnectVIServer.Version
    #Check for supported versions
    if ($VIVer -ge 5.1 -OR $VIVer -le 4.0)
        {
        Write-Log 1 $MATLog "ERROR - Your version of VMware $VIVer is not supported with the current version of MVMC/MAT" Red
        Write-Log 1 $MATLog "Unable to contiue. See the MVMC readme for more details on supported versions" Red
        exit
        }
    Write-Log 1 $MATLog "Connected to $ConnectVIServer on $VIPort version: $VIVer"
        

    #Make sure VM is running
    $VMstate = VMware.VimAutomation.Core\Get-VM -Name $VMTarget
    if ($VMState.PowerState -ne "PoweredOn")
        {
        Write-Log 1 $MATLog "ERROR - $VMTarget is not Powered On." Red
        $time = get-date
        SQLWrite "UPDATE VMQueue SET [Warning] = 1, InUse = NULL, ConvertServer = NULL, Completed = 0,  Summary = 'VM is not powered on', PID = NULL, EndTime = '$time'  WHERE JobID = $ID" 

        #Go to next VM
        Convert
        }
        
    #Disconnect any CD image or device
    if ($RemoveCD -eq 1)
        {
            Write-Log 1 $MATLog "Removing CD drives from $VMTarget" Cyan
            Try
                {
                $VMmedia = VMware.VimAutomation.Core\Get-VM -Name $VMTarget
                $VMCD = Get-CDDrive -VM $VMmedia 
                Set-CDDrive $VMCD -NoMedia -Confirm:$False | Out-Null
                }
            Catch
                {
                    Write-Log 1 $MATLog "WARNING - Unable to remove CD drives from $VMTarget $($_.exception.message)" Yellow
                    SQLWrite "UPDATE VMQueue SET [Warning] = 1, [Summary] = 'Unable to remove CDs' WHERE JobID = $ID" 
                }
        }

    #Disconnect any Floppy disk 
    if ($RemoveFloppy -eq 1)
        {
            Write-Log 1 $MATLog "Removing floppy drives from $VMTarget" Cyan
            Try
            {
                $VMmedia = VMware.VimAutomation.Core\Get-VM -Name $VMTarget
                $VMfloppy = Get-FloppyDrive -VM $VMmedia 
                Set-FloppyDrive $VMfloppy -NoMedia -Confirm:$False | Out-Null
            }
            Catch
            {
                Write-Log 1 $MATLog "WARNING - Unable to remove floppies from $VMTarget $($_.exception.message)" Yellow
                SQLWrite "UPDATE VMQueue SET [Warning] = 1, [Summary] = 'Unable to remove floppies' WHERE JobID = $ID" 
            }
        }

    #Prepare NICs for conversion
    if ($MigrateNICs -eq 1)
    {
        Write-Log 3 $MATLog "Preparing to Migrate NICs for $VMTarget"
        
        #Prepare scriptSaveCreds
        $Step1 = '{3}{0}{3},{3}{1}{3},{3}{2}{3} | out-file C:\Windows\Temp\nicCreds.txt' -f $guserName,$gPwd,$guserDomain,"'"
        $Step1 = $Step1 + "`n Write-Host 'Done'"

        #Prepare Regedit & Network Restore
        $Step2 = 
        {
            $RunOnceKey = "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce"
            set-itemproperty $RunOnceKey "ConfigureServer" "C:\Windows\Temp\nicRestore.cmd"
            $WinLogonKey ="HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Winlogon"
            $guserName,$gPwd,$guserDomain = Get-Content "C:\Windows\Temp\nicCreds.txt"
            set-itemproperty $WinLogonKey "DefaultUserName" $guserName
            set-itemproperty $WinLogonKey "DefaultPassword" $gPwd
            set-itemproperty $WinLogonKey "AutoAdminLogon" "1"
            set-itemproperty $WinLogonKey "DefaultDomainName" $guserDomain
            Remove-Item "C:\Windows\Temp\nicCreds.txt"
        $restorecode =
            {
                Start-Sleep -Seconds 30;$na = (New-Object -com shell.application).Namespace(0x31)
                Foreach ($n in (Import-Clixml -Path "C:\Windows\Temp\nicConfig.xml")){$na.Items()|?{$_.Name -eq $(gwmi win32_networkadapter -F "MACAddress='$($N.MACAddress)'").NetConnectionID}|%{ $_.Name="$($_.Name)_old" }}
                Foreach ($n in (Import-Clixml -Path "C:\Windows\Temp\nicConfig.xml")){$na.Items()|?{$_.Name -eq $(gwmi win32_networkadapter -F "MACAddress='$($N.MACAddress)'").NetConnectionID}|%{ $_.Name=$n.NetConnectionID }}
                Foreach ($n in (Import-Clixml -Path "C:\Windows\Temp\nicDNS.xml")){([wmiclass]'Win32_NetworkAdapterConfiguration').SetDNSSuffixSearchOrder($n.DNSDomainSuffixSearchOrder);gwmi win32_networkadapterconfiguration -F "MACAddress='$($N.MACAddress)'" |%{$_.SetDNSServerSearchOrder($n.DNSServerSearchOrder);$_.SetDNSDomain($n.DNSDomain)}}
                netsh -f "C:\Windows\Temp\nicConfig.txt"
                $WinLogonKey ="HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Winlogon"
                Remove-ItemProperty $WinLogonKey "AutoAdminLogon"
                Remove-itemproperty $WinLogonKey "DefaultPassword"
                Remove-itemproperty $WinLogonKey "DefaultUserName"
                C:\Windows\system32\shutdown.exe -s -t 5 -f
            }
        "C:\Windows\System32\WindowsPowerShell\v1.0\Powershell.exe -encodedcommand {0}" -f [convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($restorecode)) | Out-File -FilePath "C:\Windows\Temp\nicRestore.cmd" -Encoding ascii
        gwmi win32_NetworkAdapter | where { $_.NetConnectionID } | Select-Object -Property Name, MacAddress,NetConnectionID | Export-Clixml -Path "C:\Windows\Temp\nicConfig.xml"
        gwmi win32_networkadapterconfiguration | ? {$_.IPEnabled -eq "True"} | select MACAddress, DNSDomain, DNSDomainSuffixSearchOrder, DNSServerSearchOrder | Export-Clixml -Path "C:\Windows\Temp\nicDNS.xml"
        netsh dump |  ?{$_ -notmatch "^\s|#"} | Out-File -FilePath "C:\Windows\Temp\nicConfig.txt" -Encoding ascii
        Write-Host "Done"
        }

        #Prepare shutdown script
        $Step3 = 
        { 
            $StartDHCP = 
            {gwmi win32_NetworkAdapterConfiguration | ? {! $_.DHCPEnabled -and $_.IPAddress } | %{ $_.EnableDHCP()}
            $NetworkConnections = (New-Object -com shell.application).Namespace(0x31)
            Foreach ($NIC in (gwmi win32_NetworkAdapter | where { $_.NetConnectionID }|Select-Object -ExpandProperty NetConnectionID))
                {
                        $NetworkConnections.Items() |Where-Object {$_.Name -eq $NIC} |ForEach-Object { $_.Name="$($_.Name)_VMware"}
                }
        }
 
        # Backup existing Shutdown Scripts (they will not be run during migration)
        if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Shutdown") 
                                {
            Reg export "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Shutdown" C:\Windows\Temp\gposcriptsShutdown.reg /y
            Reg export "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\Shutdown" C:\Windows\Temp\gpostateShutdown.reg /y
            Remove-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Shutdown" -Recurse
            Remove-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\Shutdown" -Recurse
        }
 
        # Add clean-up to $scriptContents including restoration of previous Shutdown Scripts
        $scriptContents = "# You should not see this file.  It is used during virtual machine migration.`n# Consult the MAT Installation and Usage guide for more info.`n`nRemove-Item 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Shutdown' -Recurse`nRemove-Item 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\Shutdown' -Recurse`nReg Import C:\Windows\Temp\gposcriptsShutdown.reg`nReg Import C:\Windows\Temp\gpostateShutdown.reg`nRemove-Item C:\Windows\Temp\gposcriptsShutdown.reg`nRemove-Item C:\Windows\Temp\gpostateShutdown.reg`nRemove-Item C:\Windows\System32\GroupPolicy\Machine\Scripts\Shutdown\MATcleanup.ps1`n`n" + $StartDHCP
 
        # Write script to GPO 
        if (!(Test-Path "C:\Windows\System32\GroupPolicy\Machine\Scripts\Shutdown"))
        {New-Item -path "C:\Windows\System32\GroupPolicy\Machine\Scripts\Shutdown" -Type directory }
        $scriptContents | out-file C:\Windows\System32\GroupPolicy\Machine\Scripts\Shutdown\MATcleanup.ps1 -Force
        Write-Host "Done"
        }
 
        #Prepare \Group Policy\Scripts\
        $Step4 = 
        {
            if (!(Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts"))
                {new-item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts"}
            new-item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Shutdown"
            new-item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Shutdown\0"
            new-item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Shutdown\0\0"
 
            new-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Shutdown\0" -PropertyType String -Name "DisplayName" -Value "Local Group Policy"
            new-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Shutdown\0" -PropertyType String -Name "FileSysPath" -Value "C:\WINDOWS\System32\GroupPolicy\Machine"
            new-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Shutdown\0" -PropertyType String -Name "GPO-ID" -Value "LocalGPO"
            new-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Shutdown\0" -PropertyType String -Name "GPOName" -Value "Local Group Policy"
            new-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Shutdown\0" -PropertyType DWord -Name "PSScriptOrder" -Value 1
            new-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Shutdown\0" -PropertyType String -Name "SOM-ID" -Value "Local"
 
            new-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Shutdown\0\0" -PropertyType DWord -Name "ExecTime" -Value 0
            new-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Shutdown\0\0" -PropertyType DWord -Name "IsPowerShell" -Value 1
            new-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Shutdown\0\0" -PropertyType String -Name "Parameters" -Value ""
            new-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Shutdown\0\0" -PropertyType String -Name "Script" -Value "MATcleanup.ps1"
            Write-Host "Done"
        }
        
        #Prepare \Group Policy\State\Machine\Scripts\
        $Step5 = 
        {
            if (!(Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts"))
                {new-item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts"}
            new-item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\Shutdown"
            new-item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\Shutdown\0"
            new-item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\Shutdown\0\0"

            new-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\Shutdown\0" -PropertyType String -Name "DisplayName" -Value "Local Group Policy"
            new-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\Shutdown\0" -PropertyType String -Name "FileSysPath" -Value "C:\WINDOWS\System32\GroupPolicy\Machine"
            new-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\Shutdown\0" -PropertyType String -Name "GPO-ID" -Value "LocalGPO"
            new-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\Shutdown\0" -PropertyType String -Name "GPOName" -Value "Local Group Policy"
            new-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\Shutdown\0" -PropertyType DWord -Name "PSScriptOrder" -Value 1
            new-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\Shutdown\0" -PropertyType String -Name "SOM-ID" -Value "Local"

            new-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\Shutdown\0\0" -PropertyType DWord -Name "ExecTime" -Value 0
            new-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\Shutdown\0\0" -PropertyType String -Name "Parameters" -Value ""
            new-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\Shutdown\0\0" -PropertyType String -Name "Script" -Value "MATcleanup.ps1" 
            Write-Host "Done"
        }
        
   
        Try 
        {
            ### STEP 1 - ScriptSaveCreds section
            Write-Log 3 $MATLog "Step 1 - Sending credsstore script to $VMTarget"
            $CredsStore = Invoke-VMScript -ScriptText $Step1 -vm $VMTarget -GuestUser $Guser -GuestPassword $gPwd -HostUser $shusername -HostPassword $shpwd -ScriptType powershell -ErrorAction STOP 
            if ($CredsStore.ScriptOutput -match "Done")
                {write-log 3 $MATLog "Step 1 completed successfully for $VMTarget" Cyan}
            else
                {
                Write-log 1 $MATLog "WARNING - CredStore failed - aborting Regedit section" Yellow
                SQLWrite "UPDATE VMQueue SET [Warning] = 1, [Summary] = 'Invoke-Script Credstore terminated' WHERE JobID = $ID"
                return 
                }

            if ($CredsStore.State -eq "Error")
                {
                Write-Log 1 $MATLog "ERROR - Invoke-Script terminated with: $CredsStore.TerminatingError" Red
                SQLWrite "UPDATE VMQueue SET [Warning] = 1, [Summary] = 'Invoke-Script CredStore terminated' WHERE JobID = $ID"
                return 
                }


            ### STEP 2 - Regedit section
            Write-Log 3 $MATLog "Step 2 - Sending Regedit and Restore script to $VMTarget"
            $Regedit = Invoke-VMScript -ScriptText $Step2 -vm $VMTarget -GuestUser $Guser -GuestPassword $gPwd -HostUser $shusername -HostPassword $shpwd -ScriptType powershell -ErrorAction STOP
                        
            if ($Regedit.ScriptOutput -match "Done")
                {write-log 3 $MATLog "Step 2 completed successfully for $VMTarget" Cyan}
            else
                {
                Write-log 1 $MATLog "WARNING - Regeditscripts failed - aborting Regedit section" Yellow
                SQLWrite "UPDATE VMQueue SET [Warning] = 1, [Summary] = 'Invoke-Script Regedit terminated' WHERE JobID = $ID"
                return 
                }

            if ($Regedit.State -eq "Error")
                {
                Write-Log 1 $MATLog "ERROR - Invoke-Script terminated with: $Regedit.TerminatingError" Red
                SQLWrite "UPDATE VMQueue SET [Warning] = 1, [Summary] = 'Invoke-Script Regedit terminated' WHERE JobID = $ID"
                return 
                }

           ### STEP 3 - NetworkBackup Section
            Write-Log 3 $MATLog "Step 3 - Preparing shutdown script on $VMTarget"
            $NetworkBackup = Invoke-VMScript -ScriptText $Step3 -vm $VMTarget -GuestUser $Guser -GuestPassword $gPwd -HostUser $shusername -HostPassword $shpwd -ScriptType powershell -ErrorAction STOP
            if ($NetworkBackup.ScriptOutput -match "Done")
                {write-log 3 $MATLog "Step 3 completed successfully for $VMTarget" Cyan}
            else
                {
                Write-log 1 $MATLog "WARNING - NetworkBackup failed - aborting Network Rebuild section" Yellow
                SQLWrite "UPDATE VMQueue SET [Warning] = 1, [Summary] = 'Invoke-Script Networkbackup terminated' WHERE JobID = $ID"
                return 
                }

            if ($NetworkBackup.State -eq "Error")
                {
                Write-Log 1 $MATLog "ERROR - Invoke-Script terminated with: $CredsStore.TerminatingError" Red
                SQLWrite "UPDATE VMQueue SET [Warning] = 1, [Summary] = 'Invoke-Script Networkbackup terminated' WHERE JobID = $ID"
                return 
                }

           ### STEP 4 - GPO  Section 1
            Write-Log 3 $MATLog "Step 4 - Preparing GPO on $VMTarget"
            $GPO1 = Invoke-VMScript -ScriptText $Step4 -vm $VMTarget -GuestUser $Guser -GuestPassword $gPwd -HostUser $shusername -HostPassword $shpwd -ScriptType powershell -ErrorAction STOP
            if ($GPO1.ScriptOutput -match "Done")
                {write-log 3 $MATLog "Step 4 completed successfully for $VMTarget" Cyan}
            else
                {
                Write-log 1 $MATLog "WARNING - GPO1 failed - aborting Network Rebuild section" Yellow
                SQLWrite "UPDATE VMQueue SET [Warning] = 1, [Summary] = 'Invoke-Script GPO1 terminated' WHERE JobID = $ID"
                return 
                }

            if ($GPO1.State -eq "Error")
                {
                Write-Log 1 $MATLog "ERROR - Invoke-Script terminated with: $CredsStore.TerminatingError" Red
                SQLWrite "UPDATE VMQueue SET [Warning] = 1, [Summary] = 'Invoke-Script GPO1 terminated' WHERE JobID = $ID"
                return 
                }   
                  
           ### STEP 5 - GPO  Section 2
            Write-Log 3 $MATLog "Step 5 - Fininalizing GPO on $VMTarget"
            $GPO2 = Invoke-VMScript -ScriptText $Step5 -vm $VMTarget -GuestUser $Guser -GuestPassword $gPwd -HostUser $shusername -HostPassword $shpwd -ScriptType powershell -ErrorAction STOP
            if ($GPO2.ScriptOutput -match "Done")
                {write-log 3 $MATLog "Step 5 completed successfully for $VMTarget" Cyan}
            else
                {
                Write-log 1 $MATLog "WARNING - GPO2 failed - aborting Network Rebuild section" Yellow
                SQLWrite "UPDATE VMQueue SET [Warning] = 1, [Summary] = 'Invoke-Script GPO2 terminated' WHERE JobID = $ID"
                return 
                }

            if ($GPO2.State -eq "Error")
                {
                Write-Log 1 $MATLog "ERROR - Invoke-Script terminated with: $CredsStore.TerminatingError" Red
                SQLWrite "UPDATE VMQueue SET [Warning] = 1, [Summary] = 'Invoke-Script GPO2 terminated' WHERE JobID = $ID"
                return 
                }    
               

        }
        catch
        {
            Write-Log 1 $MATLog "WARNING - Unable to send scripts to $VMTarget $($_.exception.message) on Target:${VMTarget}" Yellow
            SQLWrite "UPDATE VMQueue SET [Warning] = 1, [Summary] = 'Unable to collect networking info' WHERE JobID = $ID" 
        }
    }
    
    Write-Log 3 $MATLog "Removing Snapin VMware.VimAutomation.Core"
    Remove-PSSnapin "VMware.VimAutomation.Core"
    SQLWrite "UPDATE VMQueue SET [Summary] = 'Completed VM NIC backup' WHERE JobID = $ID"  
}


#endregion

################################
###### Collection Region ######
################################
#region Collection
###### Collect valid VMs ######
Function StartCollection 
{
    ###### Get VLAN ID ######
    Function Get-VLanID ($NetworkName){
        Get-VirtualPortGroup -Name $nic.NetworkName | ForEach-Object {
            $VirtualPortGroup = $_
            Switch ($VirtualPortGroup.GetType().Name) {
                DistributedPortGroupImpl { $VirtualPortGroup.ExtensionData.Config.DefaultPortConfig.Vlan.VlanId }
                VirtualPortGroupImpl     { $VirtualPortGroup.VlanId }
            }
        }
    }
    
    cls
    $VMData = @()
    $CollectLog = $logpath+"vmdata.log"
    $error.clear()
    
    #Clearing space for progress bar
    Write-Host
    Write-Host
    Write-Host
    Write-Host
    Write-Host
    Write-Host
    Write-Host
    Write-Log 1 $CollectLog "---------------------------------------------"
    Write-Log 1 $CollectLog "        Starting new VM collection           "
    Write-Log 1 $CollectLog "---------------------------------------------"

    ###### Connect to VC ######
    Write-Log 3 $CollectLog "Adding Snapin VMware.VimAutomation.Core"
    Add-PSSnapin "VMware.VimAutomation.Core"
    
    Write-Log 1 $CollectLog "Connecting to VMware Server"
    Try
    {
        $ConnectVIServer = Connect-VIServer -Server $shost -Protocol https -User $shusername -Password $shpwd
        $VIPort = $ConnectVIServer.Port
        $VIVer = $ConnectVIServer.Version
        #Check for supported versions
        if ($VIVer -ge 5.1 -OR $VIVer -le 4.0)
        {
            Write-Log 1 $CollectLog "ERROR - Your version of VMware $VIVer is not supported with the current version of MVMC/MAT" Red
            Write-Log 1 $CollectLog "Unable to contiue. See the MVMC readme for more details on supported versions" Red
            exit
        }
        Write-Log 1 $CollectLog "Connected to $ConnectVIServer on $VIPort version: $VIVer"
    }
    Catch
    {
       Write-Log 1 $CollectLog "ERROR - Connecting to VMware Server $($_.exception.message)" Red
       exit
    }


    ###### Grab list of VMs ######
    ### Machines must be powered on! and have tools installed or we will not capture them
    Try
    {
        $VMsList = VMware.VimAutomation.Core\Get-VM | where { $_.PowerState -eq 'PoweredOn' }
        $VMCount = $VMsList.Count
        Write-Log 1 $CollectLog "Collected $VMCount VMs"
    }
    catch
    {
        Write-Log 1 $CollectLog "ERROR - The was an error in collection: $error" Red
        exit
    }
    if ($VMcount -gt 0) 
    {
        $connString = "Data Source=$Datasource;Initial Catalog=$Catalog;Integrated Security=SSPI"
        $connection = New-Object System.Data.SqlClient.SqlConnection($connString)
        $connection.Open()
            
        $sqlDelcmd = $connection.CreateCommand()
        $DelCommandText = "DELETE FROM [VMConversionData_PS]"
        $sqlDelcmd.CommandText = $DelCommandText
        $sqlDelcmd.ExecuteNonQuery()
    
        $sqlcmd = $connection.CreateCommand()
        Write-Log 1 $CollectLog "Collecting information about each VM"
        Write-host "Hang in there... this could take a while."
        $peri = 0
        $VMData = foreach ($vm in $VMsList) 
        {
            $VMobj = $null
            $Network0 = $null
            $IPv6Address0 = $null
            $IPv4Address0 = $null
            $Network1 = $null
            $IPv6Address1 = $null
            $IPv4Address1 = $null
            $Network2 = $null
            $IPv6Address2 = $null             
            $IPv4Address2 = $null             
            $DiskLocation0 = $null            
            $DiskLocation1 = $null            
            $DiskLocation2 = $null            
            $Mac0  = $null                    
            $Mac1  = $null                    
            $Mac2  = $null
            $VLAN0 = $null
            $VLAN1 = $null
            $VLAN2 = $null

            write-progress -Activity "Collecting VMs..." -status "Enumerating $($VM.Name)" -percentcomplete (($peri++/($VMsList.Count))*100)
            $VMobj = New-Object -typename System.Object
            if ($vm.Guest.State -eq "Running") 
            {
                $VMobj | Add-Member -MemberType noteProperty -name GuestVMFQDN -value $vm.ExtensionData.Guest.HostName
                $VMobj | Add-Member -MemberType noteProperty -name DisplayName -value $vm.Name
                $VMobj | Add-Member -MemberType noteProperty -name GuestOS -value $vm.Guest.OSFullName
                $VMobj | Add-Member -MemberType noteProperty -name GuestID -value $vm.ExtensionData.Guest.GuestId
            }
            else 
            {
                $VMobj | Add-Member -MemberType noteProperty -name GuestVMFQDN -value $vm.Name
                $VMobj | Add-Member -MemberType noteProperty -name DisplayName -value $vm.Name
                $VMobj | Add-Member -MemberType noteProperty -name GuestOS -value $vm.Guest
            }
            $VMobj | Add-Member -MemberType noteProperty -name VMHostName -value $vm.Host.Name
            $VMobj | Add-Member -MemberType noteProperty -name VMHostID -value $vm.HostId
            $VMobj | Add-Member -MemberType noteProperty -name GuestVMID -value $vm.Id
            $VMobj | Add-Member -MemberType noteProperty -name GuestVMMB -value $vm.MemoryMB
            $VMobj | Add-Member -MemberType noteProperty -name GuestVMCPUCount -value $vm.NumCpu
            $VMobj | Add-Member -MemberType noteProperty -name GuestVMHDProvisionedGB -value $vm.ProvisionedSpaceGB
            $VMobj | Add-Member -MemberType noteProperty -name GuestVMHDUsedGB -value $vm.UsedSpaceGB
            $VMobj | Add-Member -MemberType noteProperty -name VMwareToolsVerion -value $($vm | Get-View).config.tools.toolsVersion
            $VMobj | Add-Member -MemberType noteProperty -name VMwareHardwareVersion -value $vm.Version
            $VMobj | Add-Member -MemberType noteProperty -name FtInfo -value $vm.ExtensionData.Config.FtInfo
            $VMobj | Add-Member -MemberType noteProperty -name MoRef -value $vm.ExtensionData.MoRef.Value
      
            $di=0
            foreach ($disk in $vm.HardDisks) 
            {
                $VMobj | Add-Member -MemberType noteProperty -name DiskLocation$di -value $disk.Filename
                $di++
            }
            $VMobj | Add-Member -MemberType noteProperty -name DiskCount -value $di

            $ni=0
            foreach ($nic in $vm.Guest.Nics) 
            {
                $VMobj | Add-Member -MemberType noteProperty -name Network$ni -value $nic.NetworkName
                $VMobj | Add-Member -MemberType noteProperty -name IPAddress$ni -value $nic.IPAddress
                $VMobj | Add-Member -MemberType noteProperty -name VLAN$ni -value (Get-VLanID $nic.NetworkName)
                $VMobj | Add-Member -MemberType noteProperty -name Mac$ni -value $nic.MacAddress
                $ni++
            }
            $VMobj | Add-Member -MemberType noteProperty -name NICCount -value $ni -PassThru

            $Network0 = $null
            $IPv6Address0 = $null
            $IPv4Address0 = $null
            $Network1 = $null
            $IPv6Address1 = $null
            $IPv4Address1 = $null
            $Network2 = $null
            $IPv6Address2 = $null             
            $IPv4Address2 = $null             
            $DiskLocation0 = $null            
            $DiskLocation1 = $null            
            $DiskLocation2 = $null            
            $Mac0  = $null                    
            $Mac1  = $null                    
            $Mac2  = $null
            $VLAN0 = $null
            $VLAN1 = $null
            $VLAN2 = $null
            switch ($VMobj.DiskCount)
            {
                1 { 
                    $DiskLocation0 = $VMobj.DiskLocation0
                }
                2 {
                    $DiskLocation0 = $VMobj.DiskLocation0
                    $DiskLocation1 = $VMobj.DiskLocation1
                }
                3 { 
                    $DiskLocation0 = $VMobj.DiskLocation0
                    $DiskLocation1 = $VMobj.DiskLocation1
                    $DiskLocation2 = $VMobj.DiskLocation2
                }
            }
            switch ($VMobj.NICCount)
            {
                1 {
                    switch ($VMobj.IPAddress0.Count)
                    {
                        1   {
                                $VLAN0 = $VMobj.VLAN0
                                $Network0 = $VMobj.Network0
                                $IPv4Address0 = $VMobj.IPAddress0[0]
                        }
                        2   {
                                $VLAN0 = $VMobj.VLAN0
                                $Network0 = $VMobj.Network0
                                $IPv6Address0 = $VMobj.IPAddress0[0]
                                $IPv4Address0 = $VMobj.IPAddress0[1]
                        }
                    }
                }
                2 {
                    switch ($VMobj.IPAddress1.Count)
                    {
                        1 {
                            $VLAN0 = $VMobj.VLAN0
                            $Network0 = $VMobj.Network0
                            $IPv6Address0 = $VMobj.IPAddress0[0]
                            $IPv4Address0 = $VMobj.IPAddress0[1]
                            $VLAN1 = $VMobj.VLAN1
                            $Network1 = $VMobj.Network1
                            $IPv6Address1 = $VMobj.IPAddress1[0]
                        }
                        2 {
                            $VLAN0 = $VMobj.VLAN0
                            $Network0 = $VMobj.Network0
                            $IPv6Address0 = $VMobj.IPAddress0[0]
                            $IPv4Address0 = $VMobj.IPAddress0[1]
                            $VLAN1 = $VMobj.VLAN1
                            $Network1 = $VMobj.Network1
                            $IPv6Address1 = $VMobj.IPAddress1[0]
                            $IPv4Address1 = $VMobj.IPAddress1[1]
                        }
                    }
                }
                3 {
                    switch ($VMobj.IPAddress2.Count)
                    {
                        1 
                            {
                            $VLAN0 = $VMobj.VLAN0
                            $Network0 = $VMobj.Network0
                            $IPv6Address0 = $VMobj.IPAddress0[0]
                            $IPv4Address0 = $VMobj.IPAddress0[1]
                            $VLAN1 = $VMobj.VLAN1
                            $Network1 = $VMobj.Network1
                            $IPv6Address1 = $VMobj.IPAddress1[0]
                            $IPv4Address1 = $VMobj.IPAddress1[1]
                            $VLAN2 = $VMobj.VLAN2
                            $Network2 = $VMobj.Network2
                            $IPv6Address2 = $VMobj.IPAddress2[0]
                            }
                        2 
                            {
                            $VLAN0 = $VMobj.VLAN0
                            $Network0 = $VMobj.Network0
                            $IPv6Address0 = $VMobj.IPAddress0[0]
                            $IPv4Address0 = $VMobj.IPAddress0[1]
                            $VLAN1 = $VMobj.VLAN1
                            $Network1 = $VMobj.Network1
                            $IPv6Address1 = $VMobj.IPAddress1[0]
                            $IPv4Address1 = $VMobj.IPAddress1[1]
                            $VLAN2 = $VMobj.VLAN2
                            $Network2 = $VMobj.Network2
                            $IPv6Address2 = $VMobj.IPAddress2[0]
                            $IPv4Address2 = $VMobj.IPAddress2[1]
                            }
                    }
                }
            }

            $CommandText = "INSERT [VMConversionData_PS] VALUES ('"+$VMobj.MoRef+"','"+$VMobj.DisplayName+"','"+$VMobj.VMHostName+"','"+$VMobj.VMHostID+"','"+$VMobj.GuestVMFQDN+"','"+$VMobj.GuestVMID+"','"+$VMobj.GuestID+"','"+$VMobj.GuestOS+"','"+$VMobj.GuestVMMB+"','"+$VMobj.GuestVMCPUCount+"','"+$VMobj.GuestVMHDProvisionedGB+"','"+$VMobj.GuestVMHDUsedGB+"','"+$VMobj.VMwareToolsVerion+"','"+$VMobj.VMwareHardwareVersion+"','"+$VMobj.FtInfo+"','"+$DiskLocation0+"','"+$DiskLocation1+"','"+$DiskLocation2+"','"+$Network0+"','"+$IPv6Address0+"','"+$IPv4Address0+"','"+$Network1+"','"+$IPv6Address1+"','"+$IPv4Address1+"','"+$Network2+"','"+$IPv6Address2+"','"+$IPv4Address2+"', '"+$VMobj.MAC0+"','"+$VMobj.MAC1+"','"+$VMobj.MAC2+"','"+$null+"','"+$null+"','"+$null+"', CURRENT_TIMESTAMP)"
         
            Write-Log 3 $CollectLog "Adding $($VMObj.DisplayName) to temp table"
            Write-Log 7 $CollectLog "Issuing $CommandText"
            $sqlcmd.CommandText = $CommandText
            $sqlcmd.ExecuteNonQuery() | out-null
        }
        Write-Log 1 $CollectLog "Executing Stored Procedure to sort by OS and transfer qualifying VMs to from Temp table into VMQueue"
        $sqlSPExeccmd = $connection.CreateCommand()
        $SPExecCommandText = "EXEC [sp_TransferVMConversionData_PS]"
        $sqlSPExeccmd.CommandText = $SPExecCommandText
        $sqlSPExeccmd.ExecuteNonQuery()
        $connection.Close()
    }
    if ($error.Count -ge 1)
        {
            Write-Log 1 $CollectLog "Error thrown. $error"
            Write-Log 1 $CollectLog "Collection completed with errors"
            exit
        }
    else
        {
            Write-Log 1 $CollectLog "Collection completed."
            Sleep 2
        }
}
#endregion

################################
###### Menu Region        ######
################################
#region menu

###### Main Menu ######
function Menu 
{
cls
$title = "     Main Menu"
$message = " "

$Collect = New-Object System.Management.Automation.Host.ChoiceDescription "Co&llect VMs", `
    "Collect VMs from VMware"

$Manage = New-Object System.Management.Automation.Host.ChoiceDescription "&Manage List", `
    "Manage List Menu"

$Start = New-Object System.Management.Automation.Host.ChoiceDescription "&Convert", `
    "Start Converting VMs in list"

$bye = New-Object System.Management.Automation.Host.ChoiceDescription "E&xit", `
    "Exit the script."

$options = [System.Management.Automation.Host.ChoiceDescription[]]($collect, $Manage, $start, $bye)

$result = $host.ui.PromptForChoice($title, $message, $options,3) 

switch ($result)
    {
        0 
            {
            StartCollection
            menu
            }
        1 {SubMenu}
        2 
            {
            Write-Log 1 $MATLog "---------------------------------------------"
            Write-Log 1 $MATLog "     Starting new batch of VMs      "
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
            Write-Log 3 $MATLog "SQL Catalog = $Catalog"
            Write-Log 1 $MATLog "Max Concurrent Conversions = $Queuelength"

            ShowStatus $Lookback
            sleep 5
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
            Start-Remotes
            }
        3 {exit}
    }
}

###### Submenu ######
function submenu 
{
cls
$title = "     Manage Menu"
$message = " "

$List = New-Object System.Management.Automation.Host.ChoiceDescription "&Create VMList", `
    "Writes available VMs to a text file"

$Update = New-Object System.Management.Automation.Host.ChoiceDescription "&Update DB", `
    "Update Database with VMs to convert and order"

$Show = New-Object System.Management.Automation.Host.ChoiceDescription "&Display VMs", `
    "Displays VMs ready to be converted"

$UpLevel = New-Object System.Management.Automation.Host.ChoiceDescription "&Main Menu", `
    "Return to main menu"

$options = [System.Management.Automation.Host.ChoiceDescription[]]($List, $Update, $show, $UpLevel)

$result = $host.ui.PromptForChoice($title, $message, $options,3) 

switch ($result)
    {
 
        0 
            {
            List "USE $Catalog SELECT DisplayName FROM VMDetails_VIEW where [ReadytoConvert] = 0 AND [Completed] <> 1 AND [Supported] = 1 ORDER BY DisplayName" 
            submenu
            }
        1 
            {
            UpdateDB
            submenu
            }
        2 
           {
            ShowVM   
            submenu            
            }
        3 {menu}
    }
}

###### Show VMs ######
function ShowVM
{
#Show machines ready for conversion
    cls
    Write-Host ""
    Write-Host "The following VMs are ready to be converted:"
    Write-Host ""
    $conn = New-Object System.Data.SqlClient.SqlConnection("Data Source=$DataSource; Initial Catalog=$Catalog; Integrated Security=SSPI")
    $conn.Open()
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = "USE $Catalog SELECT DisplayName FROM VMDetails_VIEW where [ReadytoConvert] = 1 AND Completed = 0 ORDER BY Position"
    $Reader = $cmd.ExecuteReader()
    $gotrows = $Reader.HasRows
    if (!$gotrows) {Write-Host "No VMs set to convert"}
    while ($Reader.Read())
	    {
	    $VMname  = $Reader.GetValue(0)
	    Write-Host "$VMName"
        }
    $conn.Close()
    Write-Host ""
    read-host "Press Enter to continue"
}

###### List VMs ######
function List([string]$Query)
{
    cls
    if (Test-Path $VMList)
        {
        Write-Host ""
        Write-Log 3 $MATLog "VMList found"
        $ListExist = Get-Content $VMList
            if ($ListExist)
                { 
                Write-Log 1 $MATLog "VMList  not empty."
                $title = "Delete File"
                $message = "Do you want to delete $VMList ?"

                $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
                    "Deletes $VMList."

                $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
                    "Keeps existing file"

                $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

                $result = $host.ui.PromptForChoice($title, $message, $options, 1) 

                switch ($result)
                    {
                        0 {
                            Write-Log 1 $MATLog "Deleting $VMList" Cyan
                            Sleep 1
                            Remove-item $VMList
                            List $Query
                           }
                        1 {
                            Write-Log 3 $MATLog "Returning to submenu"
                            sleep 1 
                            submenu
                          }
                    }
                
                }
            else 
                {
                    Write-Log 1 $MATLog "File found but it is empty"
                    Write-Log 1 $MATLog "Deleting $VMList"
                    Sleep 1
                    Remove-item $VMList
                    List $Query
                } 
            
        }
    else
    {
    Write-Host 
	$conn = New-Object System.Data.SqlClient.SqlConnection("Data Source=$DataSource; Initial Catalog=$Catalog; Integrated Security=SSPI")
	$conn.Open()
	$cmd = $conn.CreateCommand()
	$cmd.CommandText = "$Query"
	$Reader = $cmd.ExecuteReader()
    $gotrows = $Reader.HasRows
    if (!$gotrows) {Write-Log 1 $MATLog "No supported VMs found"}
	while ($Reader.Read())
	    {
		$DisplayName = $Reader.GetValue(0)
		add-content $VMList -value "$DisplayName"
        Write-Host "Adding $DisplayName"
        }
	$conn.Close()
    
    
    ###### Set permission so anyone can use $VMList ######
    if (Test-Path $VMList)
        {
        Write-Log 3 $MATLog "Setting file permissions on $VMList to Everyone - Full Control"
        $ACL = Get-ACL $VMList
        $FullControl = New-Object  system.security.accesscontrol.filesystemaccessrule("Everyone","FullControl","Allow")
        $ACL.SetAccessRule($FullControl)
        Set-ACL $VMList $ACL
        Write-Host ""
        Write-Host "Records added to $VMList" 
        Write-Host "--------------------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host "This file will act as the list of VMs to convert in one 'batch'" -ForegroundColor Cyan
        Write-Host "Delete any VMs from this file that you don't want to convert right now" -ForegroundColor Cyan
        Write-Host "You may also reorder the list file to affect conversion order" -ForegroundColor Cyan
        Write-Host "Run the" -nonewline -ForegroundColor Cyan
        Write-Host " UPDATE" -ForegroundColor Yellow -NoNewline
        Write-Host " option when you have completed your edits" -ForegroundColor Cyan
        Write-Host "--------------------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host ""
        read-host "Press Enter to continue"
        write-host $VMList
        notepad $VMList
        }
    }
}

###### Update database ######
function UpdateDB
{
  if (Test-Path $VMList)
        {
        cls
        Write-Host ""
            $a = Get-Content $VMList
            $length = $a.count 
            foreach ($vm in $a)
                {
                    $position = $vm.ReadCount
                    SQLWrite "UPDATE VMDetails_VIEW SET ReadytoConvert = 1, Position = $position, Summary = 'Ready to Convert' WHERE DisplayName = '$vm'"
                    Write-Host "$VM added to conversion queue"
                }
        }
    else 
        {
        Write-Log 1 $MATLog "Unable to find $VMList"
        Write-Host "Run the Create List function"
        Write-Host "Aborting"
        read-host "Press Enter to continue"
        submenu
        }
    Write-Host ""
    read-host "Press Enter to continue"

}
#endregion

###################################
###### Management Region     ######
###################################
#region Management
function Help
{
cls
    Write-Host "Syntax: PS>" -NoNewline
    Write-Host " .\ConvertVM.ps1" -ForegroundColor yellow -NoNewline
    Write-Host " <action>" -ForegroundColor Green  -NoNewline
    Write-Host " (optional)" -ForegroundColor Gray -NoNewline
    Write-Host " <delay in seconds>" -ForegroundColor Cyan -NoNewline
    Write-Host " (optional)" -ForegroundColor Gray 
    Write-Host ""
    Write-Host "Valid" -NoNewline
    Write-Host " <action>" -ForegroundColor Green -NoNewline
    Write-Host " choices:"
    Write-Host ""
    Write-Host "Collect" -ForegroundColor Green -NoNewline
    Write-Host " - Starts collection cycle"
    Write-Host "Convert" -ForegroundColor Green -NoNewline
    Write-Host " - Starts Conversion locally (Used on Remote servers)"
    Write-Host "Convert-Global" -ForegroundColor Green -NoNewline
    Write-Host " - Starts Conversion on Remotes and locally"
    Write-Host "Create-List" -ForegroundColor Green -NoNewline
    Write-Host " - Creates a VMList.txt based on unconverted VMs in database"
    Write-Host "Help" -ForegroundColor Green -NoNewline
    Write-Host " - Displays this screen"
    Write-Host "Menu" -ForegroundColor Green -NoNewline
    Write-Host " - Starts Main Menu" -NoNewline
    Write-Host " <Default action>" -ForegroundColor Gray
    Write-Host "Purge" -ForegroundColor Green -NoNewline
    Write-Host " - Deletes ALL records from database" 
    Write-Host "Report-All" -ForegroundColor Green -NoNewline
    Write-Host " - Generates a CSV file with all records from the database"
    Write-Host "Report-Complete" -ForegroundColor Green -NoNewline
    Write-Host " - Generates a CSV file with all completed VMs from the database"
    Write-Host "Report-Incomplete" -ForegroundColor Green -NoNewline
    Write-Host " -  Generates a CSV file with all incomplete Vms from the database"
    Write-Host "Report-Unsupported" -ForegroundColor Green -NoNewline
    Write-Host "  - Generates a CSV file with all unsupported VMs from the database"
    Write-Host "Report-Warning" -ForegroundColor Green -NoNewline
    Write-Host "  - Generates a CSV file with all VMs that produced a warning during conversion"
    Write-Host "Reset-All" -ForegroundColor Green -NoNewline
    Write-Host " - Resets EVERY VM record in the database"
    Write-Host "Reset-Incomplete" -ForegroundColor Green -NoNewline
    Write-Host " - Resets every incompleted VM record in the database"
    Write-Host "Reset-List" -ForegroundColor Green -NoNewline
    Write-Host " - Resets only those incomplete VM records listed in $VMList"
    Write-Host "Show-Status" -ForegroundColor Green -NoNewline
    Write-Host " - Dislays the Status table once"
    Write-Host "Show-StatusLoop" -ForegroundColor Green -NoNewline
    Write-Host " - Dislays the Status table in an endless loop"
    Write-Host "Show-VM" -ForegroundColor Green -NoNewline
    Write-Host " - Dislays the VM that are Ready to Convert"
    Write-Host "Update" -ForegroundColor Green -NoNewline
    Write-Host " - Updates the records for all VMs in VMList.txt marking them as Ready to Convert"
    Write-Host ""
    Write-Host "<Delay in seconds> [int]" -ForegroundColor Cyan -NoNewline
    Write-Host " Delays the start time of conversions (default is 0)"
    Write-Host ""
}

###### Reset records ######
function Reset ($option)
{
cls
switch ($option)
    {
        "List" 
            {
            if (Test-Path $VMList)
                {
                    $a = Get-Content $VMList
                    $length = $a.count 
                    foreach ($vm in $a)
                        {
                            $position = $vm.ReadCount
                            SQLWrite "UPDATE VMDetails_VIEW SET ReadytoConvert = 0, ConvertServer = NULL, InUse = NULL, Warning = NULL, Position = NULL, Starttime = NULL, EndTime = NULL, Status = 0, Summary = NULL, Notes = 'Record reset by Manage.ps1', PID = NULL WHERE Completed <> 1 AND DisplayName = '$vm'"
                            Write-Log 1 $MATLog "Record for $VM was reset"
                        }
                }
            else 
                {
                Write-Log 1 $MATLog "Unable to find $VMList"
                Write-Host "Run the Create List function."
                Write-Host "Aborting."
                read-host "Press Enter to continue"
                submenu
                }
        
                Write-Log 1 $MATLog "All VM records from $VMList were reset"
            }
        "Incomplete"
            {
            SQLWrite "UPDATE VMQueue SET ReadytoConvert = 0, ConvertServer = NULL, InUse = NULL, Warning = NULL, Position = NULL, Starttime = NULL, EndTime = NULL, Status = 0, Summary = NULL, Notes = 'Record reset by Manage.ps1', PID = NULL WHERE Completed <> 1"
            Write-Log 1 $MATLog "All database records not maked as complete reset."
            }
        "All" 
            {
            SQLWrite "UPDATE VMQueue SET ReadytoConvert = 0, ConvertServer = NULL, InUse = NULL, Warning = NULL, Position = NULL, Starttime = NULL, EndTime = NULL, Summary = NULL, Status = 0, Notes = 'Record reset by Manage.ps1', PID = NULL"
            Write-Log 1 $MATLog "All database records reset."
            }
    }
}
 
###### Delete ALL records ######
function Purge-DB
{
 cls
    $title = "Purge Database"
    $message = "Do you want to delete EVERY record from the database?"

    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
        "Wipes all records from datase."

    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
        "Exit"

    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

    $result = $host.ui.PromptForChoice($title, $message, $options, 1) 

    switch ($result)
        {
            0 {
                SQLWrite "USE $Catalog Delete from VMQueue"
                SQLWrite "USE $Catalog Delete from VMConversionData_PS"
                Write-Log 1 $MATLog "Deleted ALL records from database" Cyan
                }
            1 {
                exit
                }
        }
}        

###### Create CSV Report ######
function CSVReport ($Clause, $Filename)
{
$timestamp = get-date -format "MM.dd.yyyy-HH.mm"
$Reportname = "$Filename.$timestamp.csv"

if (Test-path $CurrentPath\Reports)
    {Write-Log 3 $MATLog "\Reports folder exists"}
else
    {New-Item $CurrentPath\Reports -Type directory
    Write-Log 3 $MATLog "\Reports folder created"}

if (Test-path $Reportname)
    {
    write-log 3 $MATLog "$Reportname exists, deleting it"
    Remove-Item $Reportname
    }

$Query = "USE $Catalog SELECT [JobID], [VMName], [Supported], [ReadytoConvert], [ConvertServer], [StartTime], [EndTime], [Status], [Summary], [Notes], [Completed], [PID], [Warning] FROM VMQueue $Clause"
Write-Log 7 $MATLog "Query: USE $Catalog SELECT [JobID], [VMName], [Supported], [ReadytoConvert], [ConvertServer], [StartTime], [EndTime], [Status], [Summary], [Notes], [Completed], [PID], [Warning] FROM VMQueue $Clause"
	$conn = New-Object System.Data.SqlClient.SqlConnection("Data Source=$DataSource; Initial Catalog=$Catalog; Integrated Security=SSPI")
	$conn.Open()
	$cmd = $conn.CreateCommand()
	$cmd.CommandText = "$Query"
	$Reader = $cmd.ExecuteReader()
    $gotrows = $Reader.HasRows
    if (!$gotrows) 
        {
        Write-Log 1 $MATLog "WARNING - No records found" Yellow
        exit
        }
    $table = @()
    $intCounter = 0

While ($Reader.Read())
	    {
		$ID  = $Reader.GetValue(0)
		$VMGuest = $Reader.GetValue(1)
        $Supported  = $Reader.GetValue(2)
        $Ready  = $Reader.GetValue(3)
        $CServer = $Reader.GetValue(4)
        $StartTime  = $Reader.GetValue(5)
        $EndTime  = $Reader.GetValue(6)
        $Status = $Reader.GetValue(7)
        $Summary  = $Reader.GetValue(8)
        $Notes  = $Reader.GetValue(9)
        $Complete  = $Reader.GetValue(10)
        $ProcID = $Reader.GetValue(11)
        $Warn = $Reader.GetValue(12)

        $CSVrow = New-Object psObject -Property @{'ID'=$ID;'VMName'=$VMGuest; 'Supported'=$Supported; 'ReadytoConvert'=$Ready; 'ConvertServer'=$CServer; 'StartTime'=$StartTime; 'EndTime'=$EndTime; 'Status'=$Status; 'Warnings'=$Warn; 'Summary'=$Summary; 'Notes'=$Notes; 'Completed'=$Complete; 'PID'=$ProcID}
        Export-Csv -InputObject $CSVrow -Append -Encoding UTF8 -Path $Reportname
        }
	$conn.Close()

#PURGE TOP LINE
(Get-Content $Reportname | Select-Object -Skip 1) | Set-Content $Reportname
Write-Log 1 $MATLog "CSV Report was created: $Reportname"
}
#endregion

##################################
###### Status Table Region  ######
##################################
#region status table

###### Show Table ######
function ShowStatus ([int]$Lookback)
{
cls
$Query = "USE $Catalog SELECT TOP $Showrows JOBID, DisplayName, ConvertServer, PID, Warning, Summary FROM VMDetails_VIEW WHERE (ReadytoCOnvert = 1) OR (ConvertServer is not NULL OR INuse = 1) OR (DATEDIFF(HOUR, EndTime, SYSDATETIME()) <= $Lookback)  ORDER by Position"

    #Setup datatable
    $ds = new-object System.Data.DataSet
    $ds.Tables.Add("Status")
        [void]$ds.Tables["Status"].Columns.Add("JobID",[int])
        [void]$ds.Tables["Status"].Columns.Add("VM Name",[string])
        [void]$ds.Tables["Status"].Columns.Add("PID",[int])
        [void]$ds.Tables["Status"].Columns.Add("ConvertServer",[string])
        [void]$ds.Tables["Status"].Columns.Add("Warning",[bool])
        [void]$ds.Tables["Status"].Columns.Add("Summary",[string])

	$conn = New-Object System.Data.SqlClient.SqlConnection("Data Source=$DataSource; Initial Catalog=$Catalog; Integrated Security=SSPI")
	$conn.Open()
	$cmd = $conn.CreateCommand()
	$cmd.CommandText = "$Query"
	$Reader = $cmd.ExecuteReader()
	while ($Reader.Read())
	    {
		$SID  = $Reader.GetValue(0)
		$SName = $Reader.GetValue(1)
        $SCServer = $Reader.GetValue(2)
        $SPID = $Reader.GetValue(3)
        $Warn = $Reader.GetValue(4)
        $SStatus = $Reader.GetValue(5)
        
    $dr = $ds.Tables["Status"].NewRow()
    $dr["JobID"] = $SID
    $dr["VM Name"] = $SName
    $dr["ConvertServer"] = $SCServer
    $dr["PID"] = $SPID
    $dr["Summary"] = $SStatus
    $dr["Warning"] = $Warn
    $ds.Tables["Status"].Rows.Add($dr)
            
	    }
	$conn.Close()

$ds.Tables["Status"] | Format-Table -autosize
Write-Host ""
Write-Host "Displaying queued, active conversions and any finished in the last $Lookback hour(s)" -ForegroundColor DarkGray
Write-Host "Displaying a maximum of $Showrows rows" -ForegroundColor DarkGray
Write-Host ""
}
#endregion 

##################################################################
### END FUNCTIONS ###
##################################################################