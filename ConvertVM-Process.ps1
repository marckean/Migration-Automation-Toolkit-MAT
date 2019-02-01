##########################################
#      Migration Automation Script       #
#             version 1.5.2              #
#                                        #
#        ConvertVM-Process.ps1           #
#                                        #
#         Copyright 2014 Microsoft       #
#                                        #
#      http:\\aka.ms\buildingclouds      #
##########################################

#BuildVersion = "1.5.2007"


###### Collect Params ######
param(
[string]$ID,
[string]$DataSource,
[String]$Catalog,
[string]$MATLog,
[string]$DisplayName,
[string]$VMFQDN,
[string]$Logpath,
[string]$CurrentPath,
[int]$VerboseLogging,
[string]$XMLPath,
[string]$MigrateNICs
)

###### Variables ######
$error.clear()
$LogOption = 1
$LogMonitorDelay = 30
$SleepMulitplier = 1.3
$Computername = Get-Childitem env:computername 
$localhost = $Computername.Value
$fpath = "C:\Program Files (x86)\Microsoft Virtual Machine Converter Solution Accelerator\MVMC.exe"

#Option to monitor the MVMC.log file created during conversion (0 = off)
#How Verbose should logging be (0 = limted, 3 = normal, 7 = full)

##################################################################
### FUNCTIONS ###
##################################################################
#region functions

Function Set-ScriptVariable ($Name,$Value) 
    {
    Invoke-Expression ("`$Script:" + $Name + " = `"" + $Value + "`"")
    }

function write-log([int]$level, [string]$info)
{
    if ($level -le $VerboseLogging)
    {
        $time = get-date
        Write-Host "$time - [$DisplayName] $info"
        add-content $MATLog -value "$time - [$DisplayName] $info"
    }
} 

###### SQL Writer ######
function SQLWrite([string]$Query)
{
	$conn = New-Object System.Data.SqlClient.SqlConnection("Data Source=$DataSource; Initial Catalog=$Catalog; Integrated Security=SSPI")
	$conn.Open()
	$cmd = $conn.CreateCommand()
	$cmd.CommandText = "$Query"
	$cmd.ExecuteNonQuery() |Out-Null
	$conn.Close()

    if ($VerboseLogging -ge 7)
        {Write-Log 7 "$Query"}
}

###### Do Convert ######
function DoConvert
{

###### Determine MVMC.log file ######
$trylog = $Logpath+"mvmc.log"
    if (Test-Path $trylog) 
        {
        $trylog = $Logpath+"mvmc_1.log"
        if (Test-Path $trylog) 
            {
            $trylog = $Logpath+"mvmc_2.log"
            if (Test-Path $trylog) 
                {
                $trylog = $Logpath+"mvmc_3.log"
                if (Test-Path $trylog) 
                {$mvmclog = $Logpath+"mvmc_3.log"}
                    else 
                    {
                    Write-Log 1 "Please clear the log files from the $logpath"
                    Write-Log 1 "Terminating conversion of $DisplayName"
                    exit}
                        }                                            
                else {$mvmclog = $logpath+"mvmc_2.log"}
                    }
        else {$mvmclog = $logpath+"mvmc_1.log"}
        }
    else {$mvmclog = $trylog}
    Write-Log 3 "Logfile is $mvmclog"
        
    ###### Call Log monitor (if used) ######
    if ($LogOption -eq 1)
        {
        Write-Log 3 "Log monitoring Requested:True - Calling Log monitor"
        Start-Process -FilePath powershell.exe -ArgumentList "$CurrentPath\ConvertVM-Logging.ps1 $ID $Datasource $Catalog $mvmclog $MATLog $DisplayName $LogMonitorDelay $VerboseLogging" -WindowStyle Minimized
        Write-Log 7 "Start-Process -FilePath powershell.exe -ArgumentList '$CurrentPath\ConvertVM-Logging.ps1 $ID $Datasource $Catalog $mvmclog $MATLog $DisplayName $LogMonitorDelay $VerboseLogging' -WindowStyle Minimized"
        }
    else
        {
        Write-Log 3 "Log monitoring Requested:False"
        }
   
    #Create Parameters
    $params = "$dynamicdisk /SourceHost:$shost /SourceHostUser:$shusername /SourceHostPwd:$shpwd /GuestVM:$VMFQDN /GuestUser:$gUser /GuestPwd:$gPwd /TargetHost:$thost /TargetVHDPath:$tpath $sPower $tPower"
    $DisplayNamelog = $DisplayName+".log"
    $time = get-date
        
    Write-Log 1 "-------------- Starting conversion of $DisplayName ($ID) on $localhost targeting Hyper-V host:$thost via share:$tpath " 

           
    #CONVERT and collect PID
    Write-Log 3 "Command Issued: $dynamicdisk $fpath /SourceHost:$shost /SourceHostUser:$shusername /SourceHostPwd:*** /GuestVM:$VMFQDN /GuestUser:$gUser /GuestPwd:*** /TargetHost:$thost /TargetVHDPath:$tpath $sPower $tPower"


        #This is the actual conversion command"
        $doconvert = Start-Process -FilePath $fpath -ArgumentList $params -PassThru -WindowStyle Minimized
                  
        $convertPID = $doconvert.Id
        Write-Log 1 "MVMC started using PID:$convertPID"
        SQLWrite "UPDATE VMQueue SET PID = $convertPID  WHERE JobID = $ID"   

        try
        {
            while (Get-Process -id $convertPID -ErrorAction Stop) 
            {
            Write-log 3 "(Process Monitor) [$DisplayName] MVMC.exe is still running. PID $ConvertPID is still in memory."
            Sleep $SleepTime
            }
        }
        catch
        {
             write-log 1 "MVMC.exe has exited"
        }
       
        $doconvert.HasExited | Out-Null  # this will calculate the “exitCode” field
        $convertExitCode = $doconvert.GetType().GetField("exitCode", "NonPublic,Instance").GetValue($doconvert)

        #Wait for Log monitor to end
        Sleep $SleepTime 

        #Check Error Codes
        # ERROR CODE LIST
        # 0 = Success
        # -1 = HelpRequested
        # -2 = InvalidCommandLine
        # -3 = SourceHostNotFound
        # -4 = GuestVirtualMachineNotFound
        # -5 = TargetHostNotFound
        # -6 = UnexpectedException in MVMC.exe

            #Set results based on Exit codes
            $convertResults = "Unknown Error"
            $time = get-date
            if ($convertExitCode -eq 0)
                {
                $convertResults = "Successful Conversion"
                SQLWrite "UPDATE VMQueue SET Completed = 1, Status = 1, Summary = '$convertResults', PID = NULL, EndTime = '$time'  WHERE JobID = $ID" 
                }
            if ($convertExitCode -eq -1)
                {
                $convertResults = "Help Request by Application"
                SQLWrite "UPDATE VMQueue SET Completed = 0, Status = 3, Summary = '$convertResults', PID = NULL, EndTime = '$time'  WHERE JobID = $ID" 
                }
            if ($convertExitCode -eq -2)
                {
                $convertResults = "Invalid Command line"
                SQLWrite "UPDATE VMQueue SET Completed = 0, Status = 3, Summary = '$convertResults', PID = NULL, EndTime = '$time'  WHERE JobID = $ID" 
                }
            if ($convertExitCode -eq -3)
                {
                $convertResults = "Source VMware Host Not Found"
                SQLWrite "UPDATE VMQueue SET Completed = 0, Status = 3, Summary = '$convertResults', PID = NULL, EndTime = '$time'  WHERE JobID = $ID" 
                }
            if ($convertExitCode -eq -4)
                {
                $convertResults = "Guest VM Not Found"
                SQLWrite "UPDATE VMQueue SET Completed = 0, Status = 3, Summary = '$convertResults', PID = NULL, EndTime = '$time'  WHERE JobID = $ID" 
                }
            if ($convertExitCode -eq -5)
                {
                $convertResults = "Target Hyper-V Host Not Found"
                SQLWrite "UPDATE VMQueue SET Completed = 0, Status = 3, Summary = '$convertResults', PID = NULL, EndTime = '$time'  WHERE JobID = $ID" 
                }
            if ($convertExitCode -eq -6)
                {
                $convertResults = "FAILED - Unexpected Execption in MVMC.exe"
                SQLWrite "UPDATE VMQueue SET Completed = 0, Status = 3, Summary = '$convertResults', PID = NULL, EndTime = '$time'  WHERE JobID = $ID" 
                }
        
        Write-Log 1 "[$DisplayName] Exitcode = $convertExitCode Errorlevel:$convertResults"
                       
            #If we return "Unknown error" something is WRONG - bail from process
            if ($convertResults -eq "Unknown Error")
                {
                Write-Log 1 "[$DisplayName] ERROR Terminating. Unknow error. MVMC did not return an exit code."
                SQLWrite "UPDATE VMQueue SET Completed = 0, Status = 3, Summary = 'Unknown Error', PID = NULL, EndTime = '$time'  WHERE JobID = $ID"
                exit
                }
                
        
        #Handle log file renames
        if (Test-Path $logpath$DisplayNamelog)
        { 
            $datestamp = get-date -Format hhmmssMMddyy
            $oldVMlog = "$DisplayName-" + $datestamp +".log"
            Rename-Item $mvmclog $oldVMlog
            Write-Log 1 "Renaming existing $DisplayName.log to $oldVMlog"
        }


        if (Test-Path $mvmclog)
            {
            Rename-Item $mvmclog $DisplayNamelog
            write-log 1 "Renaming $mvmclog to $DisplayNamelog"
            }
        sleep 5
}
 
###### Handles rebuilding function
function RebuildNICs  
{
    Import-Module -Name Hyper-V
    Write-log 1 "Connecting to new VM for $DisplayName on Hyper-V"
    $NewVM = Hyper-V\Get-VM -Name $DisplayName -ComputerName $thost
    if (!$NewVM)
        {
        Write-log 1 "Unable to find VM: $DisplayName on Hyper-V"
        SQLWrite "UPDATE VMDetails_VIEW SET [Warning] = 1 WHERE JobID = $ID" 
        return
        }
 
    #Collect data about VM
    $conn = New-Object System.Data.SqlClient.SqlConnection("Data Source=$DataSource; Initial Catalog=$Catalog; Integrated Security=SSPI")
	$conn.Open()
	$cmd = $conn.CreateCommand()
	$cmd.CommandText = "USE $Catalog SELECT Status, Network0, MAC0, VLAN0, Network1, MAC1, VLAN1, Network2, MAC2, VLAN2 FROM VMDetails_VIEW WHERE JobID = $ID"
	$Reader = $cmd.ExecuteReader()
	while ($Reader.Read())
	{
        $status =  $Reader.GetValue(0)
        $Network0 = $Reader.GetValue(1)
        $MACAddress0 = $Reader.GetValue(2)
        $VLAN0 = $Reader.GetValue(3)
        $Network1 = $Reader.GetValue(4)
        $MACAddress1 = $Reader.GetValue(5)
        $VLAN1 = $Reader.GetValue(6)
        $Network2 = $Reader.GetValue(7)
        $MACAddress2 = $Reader.GetValue(8)
        $VLAN2 = $Reader.GetValue(9)
    }
    $conn.Close()
    Write-log 7 "Collected VM info from database"

    if ($status -ne 3)
    {
        Write-Log 7 "VM converted successfully" 
        #Rebuild Networks
        Try
        {   
            $start = get-date
            VM-Rebuild 
        }
        Catch
        {
            write-log 1 "ERROR: Network Rebuild for $DisplayName failed: $_.exception.message" Red
            SQLWrite "UPDATE VMDetails_VIEW SET [Warning] = 1, [Summary] = 'Network rebuild failed' WHERE JobID = $ID" 
        }
    }
    else
    {Write-Log 1 "ERROR Unable to Rebuild settings for $DisplayName since the conversion failed"  Yellow}
 
}

###### Recreate NICs from db settings 
Function VM-Rebuild
{
SQLWrite "UPDATE VMDetails_VIEW SET [Summary] = 'Conversion Completed restoring VM Network Settings' WHERE JobID = $ID"
Write-Log 1 "Rebuilding networks for $DisplayName" Yellow 

    Function AddVMNetAdapter
    {
        Param($VM,$NetworkName,$MAC,$VLAN)
        Try
        {
            Write-log 3 "Creating $NetworkName NetAdapter on $($VM.Name)"
            $VMNic = $VM | Hyper-V\Add-VMNetworkAdapter -SwitchName $NetworkName -StaticMacAddress $MAC
            if ($VLAN0 -ne 0)
            {
                write-log 3 "Setting NetAdapter to VLAN $VLAN"
                $VM | Hyper-V\Get-VMNetworkAdapter | Where { $_.SwitchName -eq $NetworkName} | Hyper-V\Set-VMNetworkAdapterVlan -VlanId $VLAN -Access
            }
        }
        catch
        {
            SQLWrite "UPDATE VMDetails_VIEW SET [Completed] = 0, [Status] = 3, [EndTime] = '$(Get-Date)', [Summary] = 'Error: Creating VM Network Adapter.' WHERE JobID = $ID" 
            write-log 1 "Error: Creating VM Network Adapter. $($_.exception.message)" Red
        }
    }

        
    Try
    {
        Write-log 1 "Removing default NICs from $DisplayName on Hyper-V"
        $NewVM|Hyper-V\Get-VMNetworkAdapter|Hyper-V\Remove-VMNetworkAdapter -Confirm:$false
    }
    catch
    {
        write-log 1 "WARNING - Unable to Remove Adapaters for $DisplayName $($_.exception.message)" Red
        SQLWrite "UPDATE VMDetails_VIEW SET [Warning] = 1, [Summary] = 'Error: Creating VM $($_.exception.message)' WHERE JobID = $ID" 
        return
    }
    
    # Add Network
    if ($Network0)
    {
        if (Get-VMNetworkAdapter -ComputerName $thost -SwitchName $Network0 -ManagementOS)
            {
            Write-log 1 "Adding Network0"
            AddVMNetAdapter -VM $NewVM -NetworkName $Network0 -MAC $MACAddress0 -VLAN $VLAN0
            }
        else 
            {
            Write-log 1 "WARNING - Hyper-V Switch/Adapter $Network0 does not exist"
            SQLWrite "UPDATE VMDetails_VIEW SET [Warning] = 1, [Summary] = 'Hyper-V Switch/Adapter $Network0 does not exist' WHERE JobID = $ID" 
            }

    }
    if ($Network1)
    {
        if (Get-VMNetworkAdapter -ComputerName $thost -SwitchName $Network1 -ManagementOS)
            {
            Write-log 1 "Adding Network1"
            AddVMNetAdapter -VM $NewVM -NetworkName $Network1 -MAC $MACAddress1 -VLAN $VLAN1
            }
        else 
            {
            Write-log 1 "WARNING - Hyper-V Switch/Adapter $Network1 does not exist"
            SQLWrite "UPDATE VMDetails_VIEW SET [Warning] = 1, [Summary] = 'Hyper-V Switch/Adapter $Network1 does not exist' WHERE JobID = $ID" 
            }
    }
    if ($Network2)
    {
        if (Get-VMNetworkAdapter -ComputerName $thost -SwitchName $Network2 -ManagementOS)
            {
            Write-log 1 "Adding Network2"
            AddVMNetAdapter -VM $NewVM -NetworkName $Network2 -MAC $MACAddress2 -VLAN $VLAN2
            }
        else 
            {
            Write-log 1 "WARNING - Hyper-V Switch/Adapter $Network2 does not exist"
            SQLWrite "UPDATE VMDetails_VIEW SET [Warning] = 1, [Summary] = 'Hyper-V Switch/Adapter $Network2 does not exist' WHERE JobID = $ID" 
            }
    }
       
    Write-log 3 "Starting $NewVM to rebuild network interfaces"
    $NewVM | Hyper-V\Start-VM
    
    #Call VM wait
    Write-log 7 "Calling VM-Wait function"
    VM-Wait 
}

###### Wait for the network to fix the Mac Address #####
Function VM-Wait
{
SQLWrite "UPDATE VMQueue SET [Summary] = 'Waiting for $DisplayName to finish network rebuild' WHERE JobID = $ID"
Write-Log 1 "Waiting for network rebuild and reboot" Yellow 

        Import-Module -Name Hyper-V
        # wait until the VM is powered off and then set all net adapeters 
        # to use dynamic mac addresses unless the global is set to keep.
        $watchDog = 20
        While ($NewVM.State -ne 'Off')
        {
            if ($Current -lt $watchDog)
                {
                    $Current++
                    Start-Sleep -Seconds 15
                    Write-Log 7 "Current counter at $Current, Waiting until it reaches $Watchdog"
                }
            else
                {
                    SQLWrite "UPDATE VMDetails_VIEW SET [Warning] = 1, [Summary] = 'WARNING - $DisplayName failed to Power on in the time allowed' WHERE JobID = $ID" 
                    Write-Log 3 "WARNING - $DisplayName failed to Power on in the time allowed." Yellow
                }

        }

        if ($DynMac -eq 1)
            {Get-VM $DisplayName | Get-VMNetworkAdapter | Set-VMNetworkAdapter -DynamicMacAddress}
        Get-VM $DisplayName | Start-VM
        SQLWrite "UPDATE VMQueue SET [InUse] = NULL, [ConvertServer] = NULL, [EndTime] = '$(Get-Date)', [Completed] = 1, [Status] = 1, [Summary] = 'Conversion Completed in $([math]::Round((New-TimeSpan $time).TotalMinutes)) Minutes' WHERE JobID = $ID" 
        Write-Log 1 "Completed $DisplayName conversion" Green


        #power on VM
        if ($finalpowerstate -eq 1)
            {        
                Write-log 3 "Starting $NewVM and completeing conversion"
                $NewVM | Hyper-V\Start-VM
            }
}   

 #endregion
  
##################################################################
### SCRIPT BODY ###
##################################################################
Write-Log 1 "Starting ConvertVM-Process"
Write-Log 1 "Record ID = $ID"
Write-Log 1 "Datasource = $DataSource"
Write-Log 1 "Catalog = $Catalog"
Write-Log 1 "MAT log = $MATLog"
Write-Log 1 "DisplayName = $DisplayName"
Write-Log 1 "FQDN = $VMFQDN"
Write-log 1 "Log path = $Logpath"
Write-log 1 "Current Path = $CurrentPath"
Write-log 1 "Logging level = $VerboseLogging"
Write-Log 1 "XML Path = $XMLPath"
Write-log 1 "Migrate NICs = $MigrateNICs"

$Variable = [XML] (Get-Content "$XMLPath")
$Variable.MAT.VMware | ForEach-Object {$_.Variable} | Where-Object {$_.Name -ne $null} | ForEach-Object {Set-ScriptVariable -Name $_.Name -Value $_.Value}
$Variable.MAT.HyperV | ForEach-Object {$_.Variable} | Where-Object {$_.Name -ne $null} | ForEach-Object {Set-ScriptVariable -Name $_.Name -Value $_.Value}
 
###### Get Remaining XML Settings
Write-Log 1 "Hyper-V Share = $tpath"
Write-Log 3 "Source Power = $sPower"
Write-Log 3 "Target Power = $tPower"
Write-Log 1 "VMware Host = $shost"
Write-Log 1 "Hyper-V Host = $thost"
Write-Log 1 "Dynamic Disk = $dynamicdisk"
Write-Log 3 "Log Monitor Delay = $LogMonitorDelay"
Write-Log 3 "Sleep Multiplier = $SleepMulitplier"
$SleepTime = ($LogMonitorDelay * $SleepMulitplier)
write-log 1 "Process Monitor interval = $SleepTime"

#Set Power On Options for CLI
if ($sPower -eq 1) 
    {$sPower = "/PowerOnSourceVM"}
else
    {$sPower = ""}

#Set final powerstates for CLI and Network options
if ($tPower -eq 1) 
    {
    if ($MigrateNICs = 1)
        {
        $tPower = ""
        $finalpowerstate = 1
        }
    else
        {
        $tPower = "/PowerOnDestinationVM"
        $finalpowerstate = 1
        }
    }
else
    {
    $tPower = ""
    $finalpowerstate = 0
    }

if ($dynamicdisk -eq 1) 
    {$dynamicdisk = "/Dynamic"}
else
    {$dynamicdisk = $null}

###### Start ######
write-log 1 "-------Starting Conversion"
$time = get-date
SQLWrite "UPDATE VMQueue SET [Summary] = 'Starting Conversion', [StartTime] = '$time' WHERE JobID = $ID" 
DoConvert
if ($MigrateNICs -eq 1)
    {RebuildNICs}
Write-log 1 "Done with Conversion." yellow 
SQLWrite "UPDATE VMQueue SET InUSe = NULL, ConvertServer = NULL WHERE JobID = $ID" 