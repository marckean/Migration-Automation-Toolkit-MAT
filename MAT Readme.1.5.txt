##########################################
#      Migration Automation Toolkit      #
#             version 1.5.2              #
#                                        #
#                Read Me                 #
#                                        #
#         Copyright 2014 Microsoft       #
#                                        #
#      http:\\aka.ms\buildingclouds      #
##########################################

For full details on how to use the MAT, please refer to the Installation and Usage Guide included with this package.


What's new in 1.5.2
===================
FIXED BUG: VMList.txt would have fewer VMs than expecxted - updated Setup MAT.sql to address.
FIXED BUG: Remote Helpers would start in 64-bit PowerShell instead of 32-bit required by MigrateNICs
NEW: Purge option clears temp table as well as VMQueue. 
NEW: ConvertVM-Process begins logging sooner.


What's new in 1.5.1
===================
DEPRECATED: VLAN Tagging on Hyper-V. VLAN tags caused collection problems in certain environments. Feature will be reintroduced when fixed.
FIXED BUG: Set-ACL would attempt to run if no VMList.txt was created. 

What's New in 1.5
=================
FIXED BUG: Get-Report functions fail if no data is present
FIXED BUG: Add warning for !gotrows
FIXED BUG: Reports - clear first line of erroneous XML
FIXED BUG: "Adding $($VM.GuestVMFQDN) to temp table" produces no data
FIXED BUG: Repair variable format
FIXED BUG: Fix $XMLPath
FIXED BUG: Elevation fails with parameters from right-click
FIXED BUG: Fixed/Dynamic disk switch not available

NEW: Create new function for Show-Warnings
NEW: Update Help with Show-Warnings
NEW: Add PreFlight checks
NEW: Check PowerShell version
NEW: Check for SQL connection
NEW: Check for PowerShell 32 bit if using Networking
NEW: Create general function to handle preflight errors
NEW: Check for supported versions of VMware
NEW: Verify VM is running before starting
NEW: Jump to next VM if current VM is off
NEW: Use timestamp in reports filename
NEW: Discovery and disconnect CD devices
NEW: Discovery and disconnect floppy drives
NEW: Add variables to make cd/floppy removal optional
NEW: Add network rebuild function as optional
NEW: Add scripts to save NIC info
NEW Add scripts to remove old nic/recreate nic on Hyper-V
NEW Add scripts to recreate NIC info
NEW: Add scripts to recreate VLAN info
NEW: Add option to preserve or recreate MAC
NEW: Add checks to verify power states are set properly
NEW: Add BackUPVMNetwork to ConvertVM-Functions
NEW: Add VM-Wait and Net-Rebuild to ConvertVM-Process
NEW: Check to verify Network exists
NEW: Update Management functions with DisplayName
NEW: Change user exposed VM references to refer to “Display Name” not FQDN
NEW: Add collection cycle to include all data possible
NEW: Add VMList logic for new collection data
NEW: Add logic in update db to support new collection data
NEW: Revise stored Proc
NEW: New View VMDETAILS_VIEW rolling up all VM data
NEW: New data in temp table
NEW: New field in versions table 
NEW: Add Warning field to catch any non-critical errors
NEW: Add Show-Status with new Warning field
NEW: Add reports with Warning field
NEW: Add Reset function with Warning
NEW: Add database/script version check
NEW: Tag DB with version
NEW: Tag Script with version
NEW: Add logic to verify and fail if not valid
NEW: Version Variable file
NEW: Add Check-version function
NEW: Add logic to skip function
NEW: Add dynamic to Variable.xml
NEW: Add logic to convert script to support /Dynamic
NEW: Integrate MAt4Shift Options
NEW: Add new DB collection
NEW: Add new Stored Proc
NEW: Rename scripts to mirror MAT4Shift
NEW: Add "tab complete" functions to ConvertVM.ps1 
NEW: Add collect option
NEW: Add Update option
NEW: Add Create-List option
NEW: Add region sections to ConvertVM-Functions
NEW: Add region sections for Convert-VM Logging
NEW: Add region sections for ConvertVM-Process
NEW: Alphabetize functions
NEW: Add function to open Notepad after Create-List

Syntax: PS> .\ConvertVM.ps1 <action> (optional) <delay in seconds> (optional)

Valid <action> choices:

Collect - Starts collection cycle
Convert - Starts Conversion locally (Used on Remote servers)
Convert-Global - Starts Conversion on Remotes and locally
Create-List - Creates a VMList.txt based on unconverted VMs in database
Help - Displays this screen
Menu - Starts Main Menu <Default action>
Purge - Deletes ALL records from database
Report-All - Generates a CSV file with all records from the database
Report-Complete - Generates a CSV file with all completed VMs from the database
Report-Incomplete -  Generates a CSV file with all incomplete Vms from the database
Report-Unsupported  - Generates a CSV file with all unsupported VMs from the database
Report-Warning  - Generates a CSV file with all VMs that produced a warning during conversion
Reset-All - Resets EVERY VM record in the database
Reset-Incomplete - Resets every incompleted VM record in the database
Reset-List - Resets only those incomplete VM records listed in C:\Source\MAT-Branch\VMlist.txt
Show-Status - Dislays the Status table once
Show-StatusLoop - Dislays the Status table in an endless loop
Show-VM - Dislays the VM that are Ready to Convert
Update - Updates the records for all VMs in VMList.txt marking them as Ready to Convert

<Delay in seconds> [int] Delays the start time of conversions (default is 0)

These can also be veiwed using .\ConvertVM.ps1  Help


If you have MAT questions or suggestions for improvement share them at the TechNet Forum.
If you are reporting a problem please set the value "$VerboseLogging = 7" and collect new logs. This will produce the most detailed logs possible. 

If you need technical support, the MVMC is a supported product. The command line (minus passwords) for the MVMC.exe is logged for each conversion. 
You can take this command line and issue it outside of the MAT to verify the problem is reproducible with MVMC alone. If so contact support.


