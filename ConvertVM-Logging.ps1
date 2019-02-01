##########################################
#      Migration Automation Script       #
#             version 1.5.2              #
#                                        #
#              logging.ps1               #
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
[string]$mvmclog,
[string]$MATLog,
[string]$VMName,
[int]$LogMonitorDelay,
[int]$VerboseLogging
)


###### Variables ######
$Status = "zzz"
$error.clear()

##################################################################
### FUNCTIONS  ###
##################################################################
#region functions

###### Write-Log ######
function write-log([int]$level, [string]$info)
{
    if ($level -le $VerboseLogging)
    {
        $time = get-date
        Write-Host "$time - (Log Monitor) [$VMName] $info"
        add-content $MATLog -value "$time - (Log Monitor) [$VMName] $info"
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


###### Check the Log ######
function Invoke-LogCheck
{
#verify Log file exists
if (Test-Path $mvmclog)
        {
        Write-Log 1 "Found $mvmclog. Looking for updates."
        }
        else
        {
        Write-Log 1 "Terminating. Unable to find $mvmclog. Monitoring can not proceed. This is an unexpected condition."
        exit
        }


$LogCheck = Select-String -Path $mvmclog -Pattern "_4127   End - Machine conversion" -CaseSensitive -quiet
    If ($LogCheck)
        {
        $Status = "End" 
        Write-Log 1 "Conversion Ended"
        SQLWrite "UPDATE VMQueue SET Summary = 'Conversion Ended' WHERE JobID = $ID"
        }
    else {

        $LogCheck = Select-String -Path $mvmclog -Pattern "_6570 Start - Creating and configuring virtual machine on the Hyper-V host" -CaseSensitive -quiet
            If ($LogCheck)
                {
                $Status = "Stage 6 - Creating virtual machine on the Hyper-V host"
                Write-Log 3 "$Status"
                SQLWrite "UPDATE VMQueue SET Summary = '$Status' WHERE JobID = $ID" 
                }
            else {

             $LogCheck = Select-String -Path $mvmclog -Pattern "_2562 Start - Copying converted disks" -CaseSensitive -quiet
                If ($LogCheck)
                    {
                    $Status = "Stage 5 - Copying converted disks"
                    Write-Log 3 "$Status"
                    SQLWrite "UPDATE VMQueue SET Summary = '$Status' WHERE JobID = $ID" 
                    }
                else {

                 $LogCheck = Select-String -Path $mvmclog -Pattern "_5758 Start - Converting VMware virtual disks" -CaseSensitive -quiet
                    If ($LogCheck)
                        {
                        $Status = "Stage 4 - Converting Disks"
                        Write-Log 3 "$Status"
                        SQLWrite "UPDATE VMQueue SET Summary = '$Status' WHERE JobID = $ID" 
                        }
                    else {
                     $LogCheck = Select-String -Path $mvmclog -Pattern "Restoring source virtual machine to original state" -CaseSensitive -quiet
                        If ($LogCheck)
                            {
                            $Status = "Stage 3 - Restoring source VM to original state"
                            Write-Log 3 "$Status"
                            SQLWrite "UPDATE VMQueue SET Summary = '$Status' WHERE JobID = $ID" 
                            }
                        else {

                        $LogCheck = Select-String -Path $mvmclog -Pattern "Downloading VMware virtual disks from source virtual machine" -CaseSensitive -quiet
                            If ($LogCheck)
                                {
                                $Status = "Stage 2 - Downloading VM disks"
                                Write-Log 3 "$Status"
                                SQLWrite "UPDATE VMQueue SET Summary = '$Status' WHERE JobID = $ID"   
                                }
                            else {

                            $LogCheck = Select-String -Path $mvmclog -Pattern "_3609 Start - Preparing source virtual machine" -CaseSensitive -quiet
                                If ($LogCheck)
                                    {
                                    $Status = "Stage 1 - Preparing VM"
                                    Write-Log 3 "$Status"
                                    SQLWrite "UPDATE VMQueue SET Summary = '$Status' WHERE JobID = $ID"   
                                    }
                                else {

                                $LogCheck = Select-String -Path $mvmclog -Pattern "_4127 Start - Machine conversion" -CaseSensitive -quiet
                                    If ($LogCheck)
                                        {
                                        $Status = "Stage 0 - Conversion Started"
                                        Write-Log 3 "$Status"
                                        SQLWrite "UPDATE VMQueue SET Summary = '$Status' WHERE JobID = $ID"   
                                        }
                                    else
                                        {
                                        #IF we get this far without $LogCheck = True something is wrong
                                        Write-Log  "Unable to find any events in MVMC.log file. Unable to monitor log"
                                        SQLWrite "UPDATE VMQueue SET Summary = 'Unable to monitor log' WHERE JobID = $ID" 
                                        }
}}}}}}}
$Status
}


##################################################################
### END FUNCTIONS ###
##################################################################
#endregion

##################################################################
### SCRIPT BODY ###
##################################################################

Write-Log 1 "Log Monitor has started. Monitoring MVMC.log for activity about $VMName"

#Give exe a chance to settle and write first log entry
sleep 5 

Write-Log 3 "Looking for MVMC.exe for $VMName in memory"
if (Get-Process MVMC)
    {
    write-log 1 "Found MVMC.exe for $VMName. Starting Monitor"
    $Status = Invoke-LogCheck
    }
else
    {
    Write-Log 1 "Terminating. Unable to find MVMC.exe for $VMName in memory. Unable to monitor."
    exit
    }

     Write-Log 1 $Status
If ($Status -match "zzz")
    {
    Write-Log 1 "Logging monitor failed to execute properly."
    exit
    }

while ($Status -notmatch "End")
    {
    Sleep $LogMonitorDelay
    Write-Log 3 "Looking for MVMC.exe for $VMName in Memory"
        if (Get-Process MVMC)
            {
            write-log 3 "Found MVMC.exe for $VMName. Starting Monitor"
            $Status = Invoke-LogCheck
            }
        else
            {
            Write-Log 1 "Terminating. MVMC.exe for $VMName is no longer in memory. It may have been termiated by the user. Unable to monitor."
            exit
            }
    }

if ($Status -match "End")
    {
    #Log that we are done
    Write-Log 1 "Conversion of $VMName has finished. Exiting log monitor."
    exit 
    }
