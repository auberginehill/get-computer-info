<#
Get-ComputerInfo.ps1
#>


[CmdletBinding()]
Param (
    [Parameter(ValueFromPipeline=$true,
    ValueFromPipelineByPropertyName=$true,
      HelpMessage="`r`nComputer: Which computers would you like to target? `r`n`r`nPlease enter computer names or IP addresses, one in each line. `r`n`r`nNotes:`r`n`t- To stop entering new values, please press [Enter] at an empty input row (and the script will run). `r`n`t- To exit this script, please press [Ctrl] + C`r`n")]
    [Alias("ComputerName")]
    [string[]]$Computer = "$env:COMPUTERNAME",
    [Parameter(HelpMessage="`r`nOutput: In which folder or directory would you like to find the outputted files? `r`n`r`nPlease enter a valid file system path to a directory (a full path name of a directory i.e. folder path such as C:\Windows). `r`n`r`nNotes:`r`n`t- If the path name includes space characters, please enclose the path name in quotation marks (single or double). `r`n`t- The output of GatherNetworkInfo.vbs script may be found inside the '%windir%\system32\Config' directory. `r`n")]
    [Alias("ReportPath")]
    [string]$Output = "$env:temp",
    [Parameter(HelpMessage="`r`nFile: Where is the txt file located, which contains the remote computer names? `r`n`r`nPlease enter a valid system filename ('FullPath'), which preferably includes the path to the file as well (a full path name of a file such as C:\Windows\file.txt). `r`n`r`nNotes:`r`n`t- If no path is defined, the current directory gets searched for the text file. `r`n`t- If the full filename or the directory name includes space characters, `r`n`t   please enclose the whole inputted string in quotation marks (single or double). `r`n`t- The values inside the text file could be computer names or IP addresses, one in each line. `r`n`t- If remote computers are specified, this script will use Windows Management Instrumentation (WMI) over Remote Procedure Calls (RPCs). `r`n")]
    [Alias("ListOfComputersInATxtFile","List")]
    [string]$File,
    [switch]$SystemInfo,
    [Alias("ExtractMsInfo32ToAFile","ExtractMsInfo32","MsInfo32ContentsToFile","MsInfo32Report","Expand")]
    [switch]$Extract,
    [Alias("OpenMsInfo32PopUpWindow","Window")]
    [switch]$MsInfo32,
    [Alias("Vbs")]
    [switch]$GatherNetworkInfo,
    [Alias("GetComputerInfoCmdlet","GetComputerInfo")]
    [switch]$Cmdlet
)


Begin {


    # Establish some common variables
    $ErrorActionPreference = "Stop"
    $timestamp = Get-Date -Format yyyyMMdd
    $date = Get-Date -Format g
    $time = Get-Date -Format HH.mm
    $empty_line = ""
    $computers = @()
    $osinfo = @()
    $volumes = @()
    $partition_table = @()
    $available_computers = @()
    $unavailable_computers = @()
    $host_name = $env:COMPUTERNAME
    $num_switches = 0


    # Change the following variables for the style of the report.                             # Credit: clayman2: "Disk Space"
    # Note: Using a hex format when defining the colors will probably give the best results in most browsers
    $background_color = "#FFFFFF"
    $title_font_family = "Gill Sans"
    $title_font_size = "19px"
    $title_bg_color = "#FFFFFF"
    $heading_font_family = "Arial"
    $heading_font_size = "12px"
    $heading_name_bg_color = "#FFFFFF"
    $data_font_family = "Calibri"
    $data_font_size = "11px"
    $data_alternating_row_color_odd = "#cccccc"
    $data_alternating_row_color_even = "#FFFFFF"


    # Colors for free space                                                                   # Credit: clayman2: "Disk Space"
    $very_low_space = "#b81321"                                                               # very low space:     less than 1 GB or less than 5 % free
    $low_space = "#ffca00"                                                                    # low space:          less than 5 GB or less than 10 % free
    $medium_space = "#137abb"                                                                 # medium space:       less than 10 GB or less than 15 % free


    # Function used to convert bytes to MB or GB or TB                                        # Credit: clayman2: "Disk Space"
    function ConvertBytes {
        param($size)
        If ($size -lt 1MB) {
            $drive_size = $size / 1KB
            $drive_size = [Math]::Round($drive_size, 2)
            [string]$drive_size + ' KB'
        } ElseIf ($size -lt 1GB) {
            $drive_size = $size / 1MB
            $drive_size = [Math]::Round($drive_size, 2)
            [string]$drive_size + ' MB'
        } ElseIf ($size -lt 1TB) {
            $drive_size = $size / 1GB
            $drive_size = [Math]::Round($drive_size, 2)
            [string]$drive_size + ' GB'
        } Else {
            $drive_size = $size / 1TB
            $drive_size = [Math]::Round($drive_size, 2)
            [string]$drive_size + ' TB'
        } # else
    } # function


    # Function used to convert the Time Zone Offset from minutes to hours
    function DayLight {
        param($minutes)
        If ($minutes -gt 0) {
            $hours = ($minutes / 60)
            [string]'+' + $hours + ' h'
        } ElseIf ($minutes -lt 0) {
            $hours = ($minutes / 60)
            [string]$hours + ' h'
        } ElseIf ($minutes -eq 0) {
            [string]'0 h (GMT)'
        } Else {
            [string]''
        } # else
    } # function


    # Function used to calculate the UpTime of a computer
    function UpTime {
        param ()

        $wmi_os = Get-WmiObject -class Win32_OperatingSystem -ComputerName $env:COMPUTERNAME
        $up_time = ($wmi_os.ConvertToDateTime($wmi_os.LocalDateTime)) - ($wmi_os.ConvertToDateTime($wmi_os.LastBootUpTime))

        If ($up_time.Days -ge 2) {
            $uptime_result = [string]$up_time.Days + ' days ' + $up_time.Hours + ' h ' + $up_time.Minutes + ' min'
        } ElseIf ($up_time.Days -gt 0) {
            $uptime_result = [string]$up_time.Days + ' day ' + $up_time.Hours + ' h ' + $up_time.Minutes + ' min'
        } ElseIf ($up_time.Hours -gt 0) {
            $uptime_result = [string]$up_time.Hours + ' h ' + $up_time.Minutes + ' min'
        } ElseIf ($up_time.Minutes -gt 0) {
            $uptime_result = [string]$up_time.Minutes + ' min ' + $up_time.Seconds + ' sec'
        } ElseIf ($up_time.Seconds -gt 0) {
            $uptime_result = [string]$up_time.Seconds + ' sec'
        } Else {
            $uptime_result = [string]''
        } # else (if)

            If ($uptime_result.Contains(" 0 h")) {
                $uptime_result = $uptime_result.Replace(" 0 h"," ")
                } If ($uptime_result.Contains(" 0 min")) {
                    $uptime_result = $uptime_result.Replace(" 0 min"," ")
                    } If ($uptime_result.Contains(" 0 sec")) {
                    $uptime_result = $uptime_result.Replace(" 0 sec"," ")
            } # if ($uptime_result: first)

        $uptime_result
    } # function


    # Test if the Output-path ("ReportPath") exists
    If ((Test-Path $Output) -eq $false) {

        $invalid_output_path_was_found = $true

        # Display an error message in console
        $empty_line | Out-String
        Write-Warning "'$Output' doesn't seem to be a valid path name."
        $empty_line | Out-String
        Write-Verbose "Please consider checking that the Output ('ReportPath') location '$Output', where the resulting output files are ought to be written, was typed correctly and that it is a valid file system path, which points to a directory. If the path name includes space characters, please enclose the path name in quotation marks (single or double)." -verbose
        $empty_line | Out-String
        $skip_text = "Couldn't find -Output folder '$Output'."
        Write-Output $skip_text
        $empty_line | Out-String
        Exit
        Return

    } Else {

        # Resolve the Output-path ("ReportPath") (if the Output-path is specified as relative)
        $real_output_path = Resolve-Path -Path $Output
        $csv_path = "$real_output_path\computer_info.csv"
        $html_path = "$real_output_path\computer_info.html"

        # Create a HTML-file
        # $html_file = New-Item -ItemType File -Path "$real_output_path\computer_info_$timestamp.html" -Force           # an alternative filename format
        $html_file = New-Item -ItemType File -Path $html_path -Force
        $html_file | Out-Null

    } # Else (If Test-Path $Output)


    # If an input file is specified, add the contents of the file to the list of computers to process
    If ($File) {

        If (((Test-Path $File) -eq $false) -or ((Test-Path $File -PathType Leaf) -eq $false)) {

            $invalid_txt_file_was_found = $true

            # Display an error message in console
            $empty_line | Out-String
            Write-Warning "'$File' doesn't seem to be a valid FullPath or -File parameter value."
            $empty_line | Out-String
            Write-Verbose "Please consider checking that the full filename with the path name (the '-File' variable value) '$File' was typed correctly and that it includes the path to the file as well. If the full filename or the directory name includes space characters, please enclose the whole string in quotation marks (single or double)." -verbose
            $empty_line | Out-String
            $skip_text = "Didn't open '$File'."
            Write-Output $skip_text
            Exit
            Return

        } Else {

            # Resolve path (if path is specified as relative)
            #   \S      Any nonwhitespace character (excludes space, tab and carriage return).
            #   \d      Any decimal digit.
            # Source: http://powershellcookbook.com/recipe/qAxK/appendix-b-regular-expression-reference
            $real_input_path = (Resolve-Path $File).Path
            $computer_list = (Get-Content $real_input_path) | Where { $_ -match '\S' }

                    ForEach ($item in $computer_list) {
                        $computers += $item
                    } # ForEach $item

        } # Else (If Test-Path $File)
    } Else {
        $continue = $true
    } # Else (If $File)


    # If a value for -Computer parameter is specified, add the values to the list of computers to process
    If ($Computer) {
        ForEach ($individual_computer in $Computer) {
            $computers += $individual_computer
        } # ForEach $item
    } Else {
        # Take the objects that are piped into the script
        $computers += @($input)
    } # Else (If $FilePath)


    # Count the amount of switches used
    If ($SystemInfo)            { $num_switches++ }
    If ($MsInfo32)              { $num_switches++ }
    If ($Extract)               { $num_switches++; $num_switches++ }
    If ($GatherNetworkInfo)     { $num_switches++ }
    If ($Cmdlet)                { $num_switches++ }

} # Begin




Process {

    # Try to process one available instance only once
    # Credit: Jeff Hicks: "Validating Computer Lists with PowerShell" https://www.petri.com/validating-computer-lists-with-powershell
    # $unique_computers = $computers.ToUpper() | select -Unique
    $unique_computers = $computers | select -Unique

        ForEach ($computer_candidate in $unique_computers) {

            If ($computer_candidate -match '\d' -eq $true){

                # Exclude computer candidate names that contain only numbers and return to the top of the program loop (ForEach $computer_candidate)
                #   \d      Any decimal digit.
                #   \s      Any whitespace character.
                # $env:USERNAME
                # Source: http://powershellcookbook.com/recipe/qAxK/appendix-b-regular-expression-reference
                $empty_line | Out-String
                Write-Warning "Computer '$computer_candidate': Computer name cannot contain only numbers."
                $empty_line | Out-String
                Write-Verbose "Please consider checking that the computer name '$computer_candidate' was typed correctly. Computer name cannot contain only numbers, may not be identical with the user name and cannot contain spaces." -verbose
                $empty_line | Out-String
                $skip_text = "Didn't detect '$computer_candidate'."
                Write-Output $skip_text
                Continue

            } Else {
                $connection = Test-Connection -ComputerName $computer_candidate -Count 1 -Quiet
                sleep -m 200

                If ($connection -eq $true) {
                    $available_computers += $computer_candidate
                } Else {
                    # Notify the user about the unavailable computers
                    $empty_line | Out-String
                    Write-Verbose "The computer '$computer_candidate' could not be found." -verbose
                    $unavailable_computers += $computer_candidate

                } # Else (If Test-Connection)
            } # Else
        } # ForEach

    If ($available_computers.Count -eq 0) {
        $empty_line | Out-String
        $exit_text = "Couldn't find $($Computer -join ', ')."
        Write-Output $exit_text
        $empty_line | Out-String
        Exit

    } Else {

        # Display a welcoming screen in console
        $empty_line | Out-String
        $header = "Computer Info"
        $coline = "-------------"
        Write-Output $header
        $coline | Out-String
    } # Else (If $FilePath)


    # Set the progress bar variables ($id denominates different progress bars, if more than one is being displayed)
    $activity           = "Retrieving Remote Computer Info"
    $status             = " "
    $task               = "Setting Initial Variables"
    $num_computers      = $available_computers.Count
    $threshold          = ($num_computers + $num_switches)
    $activities         = (($num_computers * 2) + $num_switches)
    $total_steps        = (($num_computers * 2) + $num_switches + 1 )
    $task_number        = 0
    $name_count         = 0
    $switch_count       = 0
    $id                 = 1

                        # Start the progress bar if there is more than one unique computer to process or any swithes were activated
                        If ($threshold -ge 2) {
                            Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $task -PercentComplete ((0.000002 / $total_steps) * 100)
                        } # If ($threshold)


    ForEach ($name in $available_computers) {

        # Increment the counters
        $task_number++
        $name_count++

                        # Update the progress bar if there is more than one unique computer to process or any swithes were activated
                        If ($threshold -ge 2) {
                            $activity = "Retrieving Remote Computer Info $task_number/$activities"
                            Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $name -PercentComplete (($task_number / $total_steps) * 100)
                        } # If ($threshold)

        # Read the registry
        $reg_key = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"

                If ( -not ( Test-Path $reg_key )) {
                    $continue = $true
                } Else {
                    $registry = Get-ItemProperty -Path $reg_key
                } # Else

        # Retrieve basic os and computer related information with WMI and display it in console
        $bios = Get-WmiObject -class Win32_BIOS -ComputerName $name
        $compsys = Get-WmiObject -class Win32_ComputerSystem -ComputerName $name
        $compsysprod = Get-WMIObject -class Win32_ComputerSystemProduct -ComputerName $name
        $enclosure = Get-WmiObject -Class Win32_SystemEnclosure -ComputerName $name
        $mobilebroadband = Get-WmiObject -Class Win32_POTSModem -ComputerName $name
        $motherboard = Get-WmiObject -class Win32_BaseBoard -ComputerName $name
        $network = Get-WmiObject -Class Win32_NetworkAdapter -ComputerName $name
        $os = Get-WmiObject -class Win32_OperatingSystem -ComputerName $name
        $processor = Get-WMIObject -class Win32_Processor -ComputerName $name
        $system = Get-WmiObject -Class MS_SystemInformation -Namespace 'root\WMI' -ComputerName $name
        $timezone = Get-WmiObject -class Win32_TimeZone -ComputerName $name
        $video = Get-WmiObject -class Win32_VideoController -ComputerName $name
        $ethernet = $network | Where-Object { $_.AdapterTypeId -ne 9 -and $_.MACAddress -ne $null -and $_.ProductName -notlike "*Virtual*" }
        $powershell = $PSVersionTable.PSVersion


                # Source: https://msdn.microsoft.com/en-us/library/aa394102(v=vs.85).aspx
                Switch ($compsys.DomainRole) {
                    { $_ -lt 0 } { $domain_role = "" }
                    { $_ -eq 0 } { $domain_role = "Standalone Workstation" }
                    { $_ -eq 1 } { $domain_role = "Member Workstation" }
                    { $_ -eq 2 } { $domain_role = "Standalone Server" }
                    { $_ -eq 3 } { $domain_role = "Member Server" }
                    { $_ -eq 4 } { $domain_role = "Backup Domain Controller" }
                    { $_ -eq 5 } { $domain_role = "Primary Domain Controller" }
                    { $_ -gt 5 } { $domain_role = "" }
                } # switch domainrole


                # Source: https://msdn.microsoft.com/en-us/library/aa394239(v=vs.85).aspx
                Switch ($os.ProductType) {
                    { $_ -lt 1 } { $product_type = "" }
                    { $_ -eq 1 } { $product_type = "Work Station" }
                    { $_ -eq 2 } { $product_type = "Domain Controller" }
                    { $_ -eq 3 } { $product_type = "Server" }
                    { $_ -gt 3 } { $product_type = "" }
                } # switch producttype


                # Source: https://msdn.microsoft.com/en-us/library/aa394474(v=vs.85).aspx
                Switch ($enclosure.ChassisTypes) {
                    { $_ -lt 1 } { $chassis = "" }
                    { $_ -eq 1 } { $chassis = "Other" }
                    { $_ -eq 2 } { $chassis = "Unknown" }
                    { $_ -eq 3 } { $chassis = "Desktop " }
                    { $_ -eq 4 } { $chassis = "Low Profile Desktop" }
                    { $_ -eq 5 } { $chassis = "Pizza Box" }
                    { $_ -eq 6 } { $chassis = "Mini Tower" }
                    { $_ -eq 7 } { $chassis = "Tower" }
                    { $_ -eq 8 } { $chassis = "Portable" }
                    { $_ -eq 9 } { $chassis = "Laptop" }
                    { $_ -eq 10 } { $chassis = "Notebook" }
                    { $_ -eq 11 } { $chassis = "Hand Held" }
                    { $_ -eq 12 } { $chassis = "Docking Station" }
                    { $_ -eq 13 } { $chassis = "All in One" }
                    { $_ -eq 14 } { $chassis = "Sub Notebook" }
                    { $_ -eq 15 } { $chassis = "Space-Saving" }
                    { $_ -eq 16 } { $chassis = "Lunch Box" }
                    { $_ -eq 17 } { $chassis = "Main System Chassis" }
                    { $_ -eq 18 } { $chassis = "Expansion Chassis" }
                    { $_ -eq 19 } { $chassis = "SubChassis" }
                    { $_ -eq 20 } { $chassis = "Bus Expansion Chassis" }
                    { $_ -eq 21 } { $chassis = "Peripheral Chassis" }
                    { $_ -eq 22 } { $chassis = "Storage Chassis" }
                    { $_ -eq 23 } { $chassis = "Rack Mount Chassis" }
                    { $_ -eq 24 } { $chassis = "Sealed-Case PC" }
                    { $_ -gt 24 } { $chassis = "" }
                } # switch chassistypes

                    $is_a_laptop = $false

                        If (($chassis -eq "Laptop") -or ($chassis -eq "Notebook") -or ($chassis -eq "Sub Notebook")) {
                            $is_a_laptop = $true
                        } Else {
                            $continue = $true
                        } # Else

                # Source: https://msdn.microsoft.com/en-us/library/aa394102(v=vs.85).aspx
                Switch ($compsys.PCSystemType) {
                    { $_ -lt 0 } { $pc_type = "" }
                    { $_ -eq 0 } { $pc_type = "Unspecified" }
                    { $_ -eq 1 } { $pc_type = "Desktop" }
                    { $_ -eq 2 } { $pc_type = "Mobile" }
                    { $_ -eq 3 } { $pc_type = "Workstation" }
                    { $_ -eq 4 } { $pc_type = "Enterprise Server" }
                    { $_ -eq 5 } { $pc_type = "Small Office and Home Office (SOHO) Server" }
                    { $_ -eq 6 } { $pc_type = "Appliance PC" }
                    { $_ -eq 7 } { $pc_type = "Performance Server" }
                    { $_ -eq 8 } { $pc_type = "Maximum" }
                    { $_ -gt 8 } { $pc_type = "" }
                } # switch pcsystemtype


                # Source: https://msdn.microsoft.com/en-us/library/aa394512(v=vs.85).aspx
                Switch ($video.CurrentScanMode) {
                    { $_ -lt 1 } { $scan_mode = "" }
                    { $_ -eq 1 } { $scan_mode = "Other" }
                    { $_ -eq 2 } { $scan_mode = "Unknown" }
                    { $_ -eq 3 } { $scan_mode = "Interlaced" }
                    { $_ -eq 4 } { $scan_mode = "Noninterlaced" }
                    { $_ -gt 4 } { $scan_mode = "" }
                } # switch CurrentScanMode


                # CPU
                $CPUArchitecture_data = $processor.Name
                If ($CPUArchitecture_data.Contains("(TM)")) {
                    $CPUArchitecture_data = $CPUArchitecture_data.Replace("(TM)","")
                    } If ($CPUArchitecture_data.Contains("(R)")) {
                            $CPUArchitecture_data = $CPUArchitecture_data.Replace("(R)","")
                } Else {
                    $continue = $true
                } # else (CPUArchitecture_data)


                # Manufacturer
                $Manufacturer_data = $compsysprod.Vendor
                If ($Manufacturer_data.Contains("HP")) {
                    $Manufacturer_data = $Manufacturer_data.Replace("HP","Hewlett-Packard")
                } Else {
                    $continue = $true
                } # else (Manufacturer_data)


                # Operating System
                $OperatingSystem_data = $os.Caption
                If ($OperatingSystem_data.Contains(",")) {
                    $OperatingSystem_data = $OperatingSystem_data.Replace(",","")
                    } If ($OperatingSystem_data.Contains("(R)")) {
                            $OperatingSystem_data = $OperatingSystem_data.Replace("(R)","")
                } Else {
                    $continue = $true
                } # else (OperatingSystem_data)


                        $osinfo += $obj_info = New-Object -TypeName PSCustomObject -Property @{
                            'Computer'                      = $name
                            'Manufacturer'                  = $Manufacturer_data
                            'Computer Model'                = $compsys.Model
                            'System Type'                   = $compsys.SystemType
                            'Domain Role'                   = $domain_role
                            'Product Type'                  = $product_type
                            'Chassis'                       = $chassis
                            'PC Type'                       = $pc_type
                            'Is a Laptop?'                  = $is_a_laptop
                            'Model Version'                 = $system.SystemSKU
                            'CPU'                           = $CPUArchitecture_data
                            'Video Card'                    = (@(ForEach ($videocard in $video) {
                                                                        If ($videocard.AdapterDACType -ne $null) {
                                                                            [string]$videocard.Name.Replace('(R)','') + ' (' + $videocard.AdapterDACType + ')'
                                                                        } Else {
                                                                            $videocard.Name.Replace('(R)','')
                                                                        } # else
                                                                }) | Out-String).Trim()
                            'Video Card_br'                 = (@(ForEach ($videocard in $video) {
                                                                        If ($videocard.AdapterDACType -ne $null) {
                                                                            [string]$videocard.Name.Replace('(R)','') + ' (' + $videocard.AdapterDACType + ')'
                                                                        } Else {
                                                                            $videocard.Name.Replace('(R)','')
                                                                        } # else
                                                                }) -join '<br />')
                            'Resolution'                    = (@(ForEach ($videocard in $video) { [string]$videocard.CurrentHorizontalResolution + ' x ' + $videocard.CurrentVerticalResolution + ' @ ' + $videocard.CurrentRefreshRate + ' MHz' + ' (' + $scan_mode + ')' }) | Out-String).Trim()
                            'Resolution_br'                 = (@(ForEach ($videocard in $video) { [string]$videocard.CurrentHorizontalResolution + ' x ' + $videocard.CurrentVerticalResolution + ' @ ' + $videocard.CurrentRefreshRate + ' MHz' + ' (' + $scan_mode + ')' }) -join '<br />')
                            'Operating System'              = $OperatingSystem_data
                            'Architecture'                  = $os.OSArchitecture
                            'Windows Edition ID'            = If ($registry.EditionID) {$registry.EditionID} Else {" "}
                            'Windows Installation Type'     = If ($registry.InstallationType) {$registry.InstallationType} Else {" "}
                            'Windows Platform'              = ([System.Environment]::OSVersion).Platform
                            'Type'                          = If ($registry.CurrentType) {$registry.CurrentType} Else {" "}
                            'SP Version'                    = $os.CSDVersion
                            'Windows BuildLab Extended'     = If ($registry.BuildLabEx) {$registry.BuildLabEx} Else {" "}
                            'Windows BuildLab'              = If ($registry.BuildLab) {$registry.BuildLab} Else {" "}
                            'Windows Build Branch'          = If ($registry.BuildBranch) {$registry.BuildBranch} Else {" "}
                            'Windows Build Number'          = $os.BuildNumber
                            'Windows Release Id'            = If ($registry.ReleaseId) {$registry.ReleaseId} Else {" "}
                            'Current Version'               = If ($registry.CurrentVersion) {$registry.CurrentVersion} Else {" "}
                            'Memory'                        = (ConvertBytes($compsys.TotalPhysicalMemory))
                            'Video Card Memory'             = (@(ForEach ($videocard in $video) { (ConvertBytes($videocard.AdapterRAM)) }) | Out-String).Trim()
                            'Video Card Memory_br'          = (@(ForEach ($videocard in $video) { (ConvertBytes($videocard.AdapterRAM)) }) -join '<br />')
                            'Logical Processors'            = $processor.NumberOfLogicalProcessors
                            'Cores'                         = $processor.NumberOfCores
                            'Physical Processors'           = $compsys.NumberOfProcessors
                            'Country Code'                  = $os.CountryCode
                            'Video Card Driver Date'        = (@(ForEach ($videocard in $video) { ($videocard.ConvertToDateTime($videocard.DriverDate)).ToShortDateString() }) | Out-String).Trim()
                            'Video Card Driver Date_br'     = (@(ForEach ($videocard in $video) { ($videocard.ConvertToDateTime($videocard.DriverDate)).ToShortDateString() }) -join '<br />')
                            'BIOS Release Date'             = (Get-Date -year ($system.BIOSReleaseDate.split("/")[-1]) -month ($system.BIOSReleaseDate.split("/")[0]) -day ($system.BIOSReleaseDate.split("/")[1])).ToShortDateString()
                            'OS Install Date'               = ($os.ConvertToDateTime($os.InstallDate)).ToShortDateString()
                            'Last BootUp'                   = (($os.ConvertToDateTime($os.LastBootUpTime)).ToShortDateString() + ' ' + ($os.ConvertToDateTime($os.LastBootUpTime)).ToShortTimeString())
                            'UpTime'                        = (Uptime)
                            'Date'                          = $date
                            'Daylight Bias'                 = ((DayLight($timezone.DaylightBias)) + ' (' + $timezone.DaylightName + ')')
                            'Time Offset (Current)'         = (DayLight($timezone.Bias))
                            'Time Offset (Normal)'          = (DayLight($os.CurrentTimeZone))
                            'Time (Current)'                = (Get-Date).ToShortTimeString()
                            'Time (Normal)'                 = If (((Get-Date).IsDaylightSavingTime()) -eq $true) {
                                                                    (((Get-Date).AddMinutes($timezone.DaylightBias)).ToShortTimeString() + ' (' + $timezone.StandardName + ')')
                                                                } ElseIf (((Get-Date).IsDaylightSavingTime()) -eq $false) {
                                                                    (Get-Date).ToShortTimeString() + ' (' + $timezone.StandardName + ')'
                                                                } Else {
                                                                    $continue = $true
                                                                } # else
                            'Daylight In Effect'            = $compsys.DaylightInEffect
                        #  'Daylight In Effect'            = (Get-Date).IsDaylightSavingTime()
                            'Time Zone'                     = $timezone.Description
                            'Connectivity'                  = (@(ForEach ($adapter in $ethernet) {
                                                                        If ($adapter.NetConnectionID -ne $null) {
                                                                            [string]$adapter.ProductName.Replace('(R)','') + ' (' + $adapter.NetConnectionID + ')'
                                                                        } Else {
                                                                            [string]$adapter.ProductName.Replace('(R)','')
                                                                        } # else
                                                                }) | Out-String).Trim()
                            'Connectivity_br'               = (@(ForEach ($adapter in $ethernet) {
                                                                        If ($adapter.NetConnectionID -ne $null) {
                                                                            [string]$adapter.ProductName.Replace('(R)','') + ' (' + $adapter.NetConnectionID + ')'
                                                                        } Else {
                                                                            [string]$adapter.ProductName.Replace('(R)','')
                                                                        } # else
                                                                }) -join '<br />')
                            'Mobile Broadband'              = (@(ForEach ($modem in $mobilebroadband) { [string]$modem.Name + ' (' + $modem.AttachedTo + ')'}) | Out-String).Trim()
                            'Mobile Broadband_br'           = (@(ForEach ($modem in $mobilebroadband) { [string]$modem.Name + ' (' + $modem.AttachedTo + ')'}) -join '<br />')
                            'OS Version'                    = $os.Version
                            'PowerShell Version'            = [string]$powershell.Major + '.' + $powershell.Minor + '.' + $powershell.Build + '.' +  $powershell.Revision
                            'BIOS Version'                  = $bios.SMBIOSBIOSVersion
                            'Mother Board Version'          = $system.BaseBoardVersion
                            'Video Card Version'            = (@(ForEach ($videocard in $video) { $videocard.DriverVersion }) | Out-String).Trim()
                            'Video Card Version_br'         = (@(ForEach ($videocard in $video) { $videocard.DriverVersion }) -join '<br />')
                            'ID'                            = $compsysprod.IdentifyingNumber
                            'Serial Number (BIOS)'          = $bios.SerialNumber
                            'Serial Number (Mother Board)'  = $motherboard.SerialNumber
                            'Serial Number (OS)'            = $os.SerialNumber
                            'UUID'                          = $compsysprod.UUID
                        } # New-Object




                    # Display OS Info in console
                    $obj_osinfo_selection = $osinfo | Select-Object 'Computer','Manufacturer','Computer Model','System Type','Domain Role','Product Type','Chassis','PC Type','Is a Laptop?','Model Version','CPU','Video Card','Resolution','Operating System','Architecture','Windows Edition ID','Windows Installation Type','Windows Platform','Type','SP Version','Windows BuildLab Extended','Windows BuildLab','Windows Build Branch','Windows Build Number','Windows Release Id','Current Version','Memory','Video Card Memory','Logical Processors','Cores','Physical Processors','Country Code','Video Card Driver Date','BIOS Release Date','OS Install Date','Last BootUp','UpTime','Date','Daylight Bias','Time Offset (Current)','Time Offset (Normal)','Time (Current)','Time (Normal)','Daylight In Effect','Time Zone','Connectivity','Mobile Broadband','OS Version','PowerShell Version','Video Card Version','BIOS Version','Mother Board Version','Serial Number (BIOS)','Serial Number (Mother Board)','Serial Number (OS)','UUID'
                    $obj_osinfo_selection.PSObject.TypeNames.Insert(0,"OSInfo")
                    Write-Output $obj_osinfo_selection
                    $empty_line | Out-String
                    $empty_line | Out-String




        # Retrieve additional disk information from volumes (Win32_Volume)
        $volumes_query = Get-WmiObject -class Win32_Volume -ComputerName $name

                ForEach ($volume in $volumes_query) {
                    $volumes += $obj_volumes = New-Object -TypeName PSCustomObject -Property @{
                            'Automount'             = $volume.Automount
                            'Block Size'            = $volume.BlockSize
                            'Boot Volume'           = $volume.BootVolume
                            'Compressed'            = $volume.Compressed
                            'Computer'              = $volume.SystemName
                            'DeviceID'              = $volume.DeviceID
                            'Drive'                 = $volume.DriveLetter
                            'DriveType'             = $volume.DriveType
                            'File System'           = $volume.FileSystem
                            'Free Space'            = (ConvertBytes($volume.FreeSpace))
                            'Free (%)'              = $free_percentage = If ($volume.Capacity -gt 0) {
                                                            $relative_free = [Math]::Round((($volume.FreeSpace / $volume.Capacity) * 100 ), 1)
                                                                [string]$relative_free + ' %'
                                                            } Else {
                                                                [string]''
                                                        } # else (if)
                            'Indexing Enabled'      = $volume.IndexingEnabled
                            'Label'                 = $volume.Label
                            'PageFile Present'      = $volume.PageFilePresent
                            'Root'                  = $volume.Name
                            'Serial Number (Volume)' = $volume.DeviceID
                            'Source'                = $volume.__CLASS
                            'System Volume'         = $volume.SystemVolume
                            'Total Size'            = (ConvertBytes($volume.Capacity))
                            'Used'                  = (ConvertBytes($volume.Capacity - $volume.FreeSpace))
                            'Used (%)'              = $used_percentage = If ($volume.Capacity -gt 0) {
                                                            $relative_size = [Math]::Round(((($volume.Capacity - $volume.FreeSpace) / $volume.Capacity) * 100 ), 1)
                                                                [string]$relative_size + ' %'
                                                            } Else {
                                                                [string]''
                                                        } # else (if)
                        } # New-Object

                } # ForEach ($volume}
    } # ForEach ($name/first)




    # Write the Computer info and a partition table to a HTML-file
    # Define the header of the HTML-file
    $html_header = '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
    <html lang="en-US" xml:lang="en-US" xmlns="http://www.w3.org/1999/xhtml">

    <head>
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <meta name="Description" content="Computer Info" />

        <title>Computer Info</title>

        <style type="text/css">
            body {
                background-color: ' + $background_color + ';
            }
            .title {
                text-align: center;
                font-family: "' + $title_font_family + '";
                font-size: ' + $title_font_size + ';
                font-weight: bold;
                background-color: ' + $title_bg_color + ';
                border: 0px solid black;
                padding: 14px;
            }

            .headings {
                text-align: center;
                font-family: "' + $heading_font_family + '";
                font-size: ' + $heading_font_size + ';
                font-weight: bold;
                background-color: ' + $heading_name_bg_color + ';
                border: 0px solid black;
                padding: 14px;
            }

            .data {
                font-family: "' + $data_font_family + '";
                font-size: ' + $data_font_size + ';
                text-align: center;
                border: 0px solid black;
                padding: 10px;
            }

            #main {
                border: 0px solid black;
                border-collapse: collapse;
                margin-left: 5em;
            }

            #main tr:nth-child(odd) {
                background-color: ' + $data_alternating_row_color_odd + ';
            }

            #main tr:nth-child(even) {
                background-color: ' + $data_alternating_row_color_even + ';
            }

            p {
                font-size: 9px;
                font-family: Calibri, "Lucida Sans", Helvetica, sans-serif;
            }

            .stats {
                font-size: 9px;
                font-family: Calibri, "Lucida Sans", Helvetica, sans-serif;
            }

            table.stats th {
                border: 0px solid black;
                text-align: left;
            }

            table.stats td {
                border: 0px solid black;
                text-align: left;
            }

            #legend {
                border: 1px solid black;
                position: absolute;
                right: 4em;
                top: 4em;
                padding: 2px;
            }
        </style>
    </head>

    <body>'




    # Write the header to the HTML-file
    Add-Content $html_file -Value $html_header




    # Write the Computer info -table and the headers of the main table
    Add-Content $html_file -Value ("
    <h3>Computer Info</h3>
    <table class='stats'>
        <tr>
            <th>Generated:</th>
            <td>" + $date + "</td>
        </tr>
        <tr>
            <th>Computer:</th>
            <td>" + $host_name  + "</td>
        </tr>
        <tr>
            <th>Manufacturer:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Manufacturer') + "</td>
        </tr>
        <tr>
            <th>Computer Model:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Computer Model') + "</td>
        </tr>
        <tr>
            <th>System Type:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'System Type') + "</td>
        </tr>
        <tr>
            <th>Domain Role:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Domain Role') + "</td>
        </tr>
        <tr>
            <th>Product Type:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Product Type') + "</td>
        </tr>
        <tr>
            <th>Chassis:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Chassis') + "</td>
        </tr>
        <tr>
            <th>PC Type:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'PC Type') + "</td>
        </tr>
        <tr>
            <th>Is a Laptop?</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Is a Laptop?') + "</td>
        </tr>
        <tr>
            <th>Model Version:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Model Version') + "</td>
        </tr>
        <tr>
            <th>CPU:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'CPU') + "</td>
        </tr>
        <tr>
            <th>Video Card:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Video Card_br') + "</td>
        </tr>
        <tr>
            <th>Resolution:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Resolution_br') + "</td>
        </tr>
        <tr>
            <th>Operating System:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Operating System') + "</td>
        </tr>
        <tr>
            <th>Architecture:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Architecture') + "</td>
        </tr>
        <tr>
            <th>Windows Edition ID:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Windows Edition ID') + "</td>
        </tr>
        <tr>
            <th>Windows Installation Type:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Windows Installation Type') + "</td>
        </tr>
        <tr>
            <th>Windows Platform:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Windows Platform') + "</td>
        </tr>
        <tr>
            <th>Type:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Type') + "</td>
        </tr>
        <tr>
            <th>SP Version:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'SP Version') + "</td>
        </tr>
        <tr>
            <th>Windows BuildLab Extended:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Windows BuildLab Extended') + "</td>
        </tr>
        <tr>
            <th>Windows BuildLab:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Windows BuildLab') + "</td>
        </tr>
        <tr>
            <th>Windows Build Branch:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Windows Build Branch') + "</td>
        </tr>
        <tr>
            <th>Windows Build Number:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Windows Build Number') + "</td>
        </tr>
        <tr>
            <th>Windows Release Id:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Windows Release Id') + "</td>
        </tr>
        <tr>
            <th>Current Version:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Current Version') + "</td>
        </tr>
        <tr>
            <th>Memory:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Memory') + "</td>
        </tr>
        <tr>
            <th>Video Card Memory:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Video Card Memory_br') + "</td>
        </tr>
        <tr>
            <th>Logical Processors:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Logical Processors') + "</td>
        </tr>
        <tr>
            <th>Cores:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Cores') + "</td>
        </tr>
        <tr>
            <th>Physical Processors:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Physical Processors') + "</td>
        </tr>
        <tr>
            <th>Country Code:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Country Code') + "</td>
        </tr>
        <tr>
            <th>Video Card Driver Date:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Video Card Driver Date_br') + "</td>
        </tr>
        <tr>
            <th>BIOS Release Date:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'BIOS Release Date') + "</td>
        </tr>
        <tr>
            <th>OS Install Date:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'OS Install Date') + "</td>
        </tr>
        <tr>
            <th>Last BootUp:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Last BootUp') + "</td>
        </tr>
        <tr>
            <th>UpTime:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'UpTime') + "</td>
        </tr>
        <tr>
            <th>Date:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Date') + "</td>
        </tr>
        <tr>
            <th>Daylight Bias:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Daylight Bias') + "</td>
        </tr>
        <tr>
            <th>Time Offset (Current):</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Time Offset (Current)') + "</td>
        </tr>
        <tr>
            <th>Time Offset (Normal):</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Time Offset (Normal)') + "</td>
        </tr>
        <tr>
            <th>Time (Current):</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Time (Current)') + "</td>
        </tr>
        <tr>
            <th>Time (Normal):</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Time (Normal)') + "</td>
        </tr>
        <tr>
            <th>Daylight In Effect:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Daylight In Effect') + "</td>
        </tr>
        <tr>
            <th>Time Zone:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Time Zone') + "</td>
        </tr>
        <tr>
            <th>Connectivity:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Connectivity_br') + "</td>
        </tr>
        <tr>
            <th>Mobile Broadband:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Mobile Broadband_br') + "</td>
        </tr>
        <tr>
            <th>OS Version:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'OS Version') + "</td>
        </tr>
        <tr>
            <th>PowerShell Version:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'PowerShell Version') + "</td>
        </tr>
        <tr>
            <th>Video Card Version:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Video Card Version_br') + "</td>
        </tr>
        <tr>
            <th>BIOS Version:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'BIOS Version') + "</td>
        </tr>
        <tr>
            <th>Mother Board Version:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Mother Board Version') + "</td>
        </tr>
        <tr>
            <th>Serial Number (BIOS):</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Serial Number (BIOS)') + "</td>
        </tr>
        <tr>
            <th>Serial Number (Mother Board):</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Serial Number (Mother Board)') + "</td>
        </tr>
        <tr>
            <th>Serial Number (OS):</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'Serial Number (OS)') + "</td>
        </tr>
        <tr>
            <th>UUID:</th>
            <td>" + ($osinfo | Select-Object -ExpandProperty 'UUID') + "</td>
        </tr>
    </table>
    <br />
    <br />


    <table id='main'>
        <tr>
            <td colspan='15' class='title'>" + "$(($available_computers -join ', ').ToUpper())" + "</td>
        </tr>
        <tr>
            <td class='headings'>Computer</td>
            <td class='headings'>Drive</td>
            <td class='headings'>Label</td>
            <td class='headings'>File System</td>
            <td class='headings'>Description</td>
            <td class='headings'>Partition</td>
            <td class='headings'>Disk</td>
            <td class='headings'>Disk Model</td>
            <td class='headings'>Compressed</td>
            <td class='headings'>Used</td>
            <td class='headings'>Used %</td>
            <td class='headings'>Status</td>
            <td class='headings'>Total Size</td>
            <td class='headings'>Free Space</td>
            <td class='headings'>Free %</td>
        </tr>")




    # Create a partition table with WMI
    ForEach ($name in $available_computers) {

        # Increment the step counter
        $task_number++

                        # Update the progress bar if there is more than one unique computer to process or any swithes were activated
                        If ($threshold -ge 2) {
                            $computer_no = ($task_number - $name_count)
                            $activity = "Retrieving a Partition Table $task_number/$activities"
                            Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $name -PercentComplete (($task_number / $total_steps) * 100)
                        } # If ($threshold)

        $disks = Get-WmiObject -class Win32_DiskDrive -ComputerName $name

        ForEach ($disk in $disks) {

                # $partitions = Get-WmiObject -class Win32_DiskPartition -ComputerName $name
                # $partitions = (Get-WmiObject -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='\\.\PHYSICALDRIVE0'} WHERE AssocClass=Win32_DiskDriveToDiskPartition")
                # $partitions = (Get-WmiObject -ComputerName $name -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='\\.\PHYSICALDRIVE0'} WHERE ResultRole=Dependent")
                $partitions = (Get-WmiObject -ComputerName $name -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='$($disk.DeviceID)'} WHERE ResultRole=Dependent")

                ForEach ($partition in ($partitions)) {

                        # $drives = Get-WmiObject -class Win32_LogicalDisk -ComputerName $name
                        # $drives = Get-WmiObject -ComputerName $name -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='Disk #0, Partition #1'} WHERE ResultRole=Dependent"
                        $drives = (Get-WmiObject -ComputerName $name -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='$($partition.DeviceID)'} WHERE AssocClass=Win32_LogicalDiskToPartition")

                        ForEach ($drive in ($drives)) {

                                $free_percentage = If ($drive.Size -gt 0) {
                                                        $relative_free = [Math]::Round((($drive.FreeSpace / $drive.Size) * 100 ), 1)
                                                            [string]$relative_free + ' %'
                                                        } Else {
                                                            [string]''
                                                        } # else (if)

                                $used_percentage = If ($drive.Size -gt 0) {
                                                        $relative_size = [Math]::Round(((($drive.Size - $drive.FreeSpace) / $drive.Size) * 100 ), 1)
                                                            [string]$relative_size + ' %'
                                                        } Else {
                                                            [string]''
                                                        } # else (if)

                                $disk_status = ""
                                    If ($used_percentage -eq 100) {
                                        $disk_status = "Full"
                                    } ElseIf ($used_percentage -ge 95) {
                                        $disk_status = "Very Low Space"
                                    } ElseIf ($used_percentage -ge 90) {
                                        $disk_status = "Low Space"
                                    } ElseIf ($used_percentage -ge 85) {
                                        $disk_status = "Medium Space"
                                    } ElseIf ($used_percentage -gt 0) {
                                        $disk_status = "OK"
                                    } Else {
                                        $disk_status = ""
                                    } # else (if)

                                $space_color = ""
                                $free_space = $drive.FreeSpace
                                If ((ConvertBytes($drive.Size)) -eq '0 KB') {
                                    $space_color = ""
                                } ElseIf ($free_space -lt 1073741824) {
                                    $space_color = $very_low_space
                                } ElseIf ($free_space -lt 5368709120) {
                                    $space_color = $low_space
                                } ElseIf ($free_space -lt 10737418240) {
                                    $space_color = $medium_space
                                } # ElseIf (previous/second/last)

                                $percentage_color = ""
                                $free_percent = If ($drive.Size -gt 0) {
                                                    [Math]::Round((($drive.FreeSpace / $drive.Size) * 100 ), 1)
                                                    } Else {
                                                        [string]''
                                                    } # else (if)
                                If ($free_percent -eq '') {
                                    $percentage_color = ""
                                } ElseIf ($free_percent -lt 5) {
                                    $percentage_color = $very_low_space
                                } ElseIf ($free_percent -lt 10) {
                                    $percentage_color = $low_space
                                } ElseIf ($free_percent -lt 15) {
                                    $percentage_color = $medium_space
                                } # ElseIf (previous/second/last)




                                    $partition_table += $obj_partition = New-Object -Type PSCustomObject -Property @{

                                        'Definition'              = $disk.Description
                                        'Disk'                    = $disk.DeviceID
                                        'Disk Capabilities'       = $disk.CapabilityDescriptions
                                        'Disk Status'             = $disk.Status
                                        'Interface'               = $disk.InterfaceType
                                        'Media Type'              = $disk.MediaType
                                        'Model'                   = $disk.Model
                                        'Serial Number (Disk)'    = $disk.SerialNumber


                                        'Boot Partition'          = $partition.BootPartition
                                        'Bootable'                = $partition.Bootable
                                        'Computer'                = $partition.SystemName
                                        'Partition'               = $partition.DeviceID
                                        'Partition Type'          = $partition.Description
                                        'Primary Partition'       = $partition.PrimaryPartition


                                    #  'Computer'                = $drive.SystemName
                                        'Compressed'              = $drive.Compressed
                                        'Description'             = $drive.Description
                                        'Drive'                   = $drive.DeviceID
                                        'File System'             = $drive.FileSystem
                                        'Free Space'              = (ConvertBytes($drive.FreeSpace))
                                        'Free (%)'                = $free_percentage
                                        'Label'                   = $drive.VolumeName
                                        'Free Space Status'       = $disk_status
                                        'Total Size'              = (ConvertBytes($drive.Size))
                                        'Used'                    = (ConvertBytes($drive.Size - $drive.FreeSpace))
                                        'Used (%)'                = $used_percentage
                                        'Serial Number (Volume)'  = $drive.VolumeSerialNumber

                                        } # New-Object


                                    # Write the bulk of the main table in the HTML-file
                                    # Please notice that the main table is not closed after this step.
                                    Add-Content $html_file -Value ('    <tr>
            <td class="data">' + $partition.SystemName + '</td>
            <td class="data">' + $drive.DeviceID + '</td>
            <td class="data">' + $drive.VolumeName + '</td>
            <td class="data">' + $drive.FileSystem + '</td>
            <td class="data">' + $drive.Description + '</td>
            <td class="data">' + $partition.DeviceID + '</td>
            <td class="data">' + $disk.DeviceID + '</td>
            <td class="data">' + $disk.Model + '</td>
            <td class="data">' + $drive.Compressed + '</td>
            <td class="data">' + (ConvertBytes($drive.Size - $drive.FreeSpace)) + '</td>
            <td class="data">' + $used_percentage + '</td>
            <td class="data">' + $disk_status + '</td>
            <td class="data">' + (ConvertBytes($drive.Size)) + '</td>
            <td class="data" bgcolor="' + $space_color + '">' + (ConvertBytes($drive.FreeSpace)) + '</td>
            <td class="data" bgcolor="' + $percentage_color + '">' + $free_percentage + '</td>
        </tr>')

                        } # ForEach ($drive)
                } # ForEach ($partition)
        } # ForEach ($disk)
    } # ForEach ($name/second)


    # Write the main table closure and the footer to the HTML-file
    Add-Content $html_file -Value ('</table>
    <div id="legend">
        <table>
            <tr>
                <td bgcolor="' + $very_low_space + '" width="10px"></td>
                <td style="font-size:12px">less than 1 GB or less than 5 % free</td>
            </tr>
            <tr>
                <td bgcolor="' + $low_space + '" width="10px"></td>
                <td style="font-size:12px">less than 5 GB or less than 10 % free</td>
            </tr>
            <tr>
                <td bgcolor="' + $medium_space + '" width="10px"></td>
                <td style="font-size:12px">less than 10 GB or less than 15 % free</td>
            </tr>
        </table>
    </div>
    <br />
    <br />
    <br />

    <p>[End of Line]</p>
    <table class="stats">
        <tr>
            <th>Generated:</th>
            <td>' + $date + '</td>
        </tr>
        <tr>
            <th>Computer:</th>
            <td>' + $host_name  + '</td>
        </tr>
    </table>

    </body>
    </html>')




} # Process




End {

    # Process the partition table
    $partition_table.PSObject.TypeNames.Insert(0,"PartitionTable")
    $partition_table_selection = $partition_table | Sort Computer,Drive | Select-Object Computer,Drive,Label,'File System','Boot Partition',Interface,'Media Type','Partition Type',Partition,Used,'Used (%)','Free Space Status','Total Size','Free Space','Free (%)'
    $partition_table_selection_screen = $partition_table | Sort Computer,Drive | Select-Object Computer,Drive,Label,Interface,'Media Type',Partition,'Used (%)','Total Size','Free Space','Free (%)'


    # Display the volumes in console
    $volumes.PSObject.TypeNames.Insert(0,"Volume")
    $volumes_selection = $volumes | Sort Computer,Drive | Select-Object 'Computer','Drive','Label','File System','System Volume','Boot Volume','Indexing Enabled','PageFile Present','Block Size','Compressed','Automount','Used','Used (%)','Total Size','Free Space','Free (%)'
    $volumes_selection_screen = $volumes | Sort Computer,Drive | Select-Object 'Computer','Drive','Label','File System','System Volume','Used','Used (%)','Total Size','Free Space','Free (%)'
    $volumes_header = "Volumes"
    $volumes_coline = "-------"
    Write-Output $volumes_header
    $volumes_coline | Out-String
    $volumes_selection_screen | Format-Table -AutoSize | Out-String


    # Write the Computer info to a landscape-oriented CSV-file (horizontal).
    Write-Verbose "Results will be written to '$csv_path' and '$html_path'."
    $obj_osinfo_selection | Export-Csv "$csv_path" -Delimiter ';' -NoTypeInformation -Encoding UTF8


    # Display the HTML-file in the default browser
    # & $real_output_path\time_zones.html
    Start-Process -FilePath "$html_path"




<#
   _____         _ _       _
  / ____|       (_) |     | |
 | (_____      ___| |_ ___| |__   ___  ___
  \___ \ \ /\ / / | __/ __| '_ \ / _ \/ __|
  ____) \ V  V /| | || (__| | | |  __/\__ \
 |_____/ \_/\_/ |_|\__\___|_| |_|\___||___/


 Switches

#>
# Source: https://technet.microsoft.com/en-us/library/ee692804.aspx
# Source: http://stackoverflow.com/questions/27175137/powershellv2-remove-last-x-characters-from-a-string#32608908
If ((($real_output_path.Path).EndsWith("\")) -eq $true) { $real_output_path = $real_output_path -replace ".{1}$"}


    # (1) SystemInfo.exe
    # Located at %windir%\system32\ directory
    # $system32 = [Environment]::GetFolderPath("System")
    # Source: https://technet.microsoft.com/en-us/library/bb491007.aspx
    # Systeminfo.exe /fo CSV | ConvertFrom-Csv | Export-Csv "system_info.csv" -Delimiter ';' -NoTypeInformation -Encoding UTF8

        If ($SystemInfo) {

            # Increment the counters
            $task_number++
            $switch_count++
            $system_info_txt = "$real_output_path\system_info.txt"

                            # Update the progress bar
                            $activity = "Processing Additional Options $task_number/$activities"
                            $task = "systeminfo.exe /fo CSV | ConvertFrom-Csv | Out-File '$system_info_txt' -Encoding UTF8"
                            Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $task -PercentComplete (($task_number / $total_steps) * 100)

                Try {
                    $system_info = systeminfo.exe /fo LIST | Out-File "$system_info_txt" -Encoding UTF8
                } Catch { Write-Debug $_.Exception }
              
        } Else {
            $continue = $true
        } # Else


    # (2) MsInfo32.exe report
    # Located at %CommonProgramFiles%\Microsoft Shared\MSInfo\ directory
    # $msinfo32_folder = [string]([Environment]::GetFolderPath("CommonProgramFiles")) + "\Microsoft Shared\MSInfo\"
    # Note: /categories +systemsummary seems to be deprecated on OSes newer than Windows XP.
    # Source: https://technet.microsoft.com/en-us/library/bb490937.aspx
    # Source: https://support.microsoft.com/en-us/help/300887/how-to-use-system-information-msinfo32-command-line-tool-switches


        If ($Extract) {

            # Increment the counters
            $task_number++
            $switch_count++
            $ms_info_txt = "$real_output_path\ms_info.txt"
            $ms_info_nfo = "$real_output_path\ms_info.nfo"

                                    # Close all msinfo32 instances
                                    # Source: http://stackoverflow.com/questions/28481811/how-to-correctly-check-if-a-process-is-running-and-stop-it
                                    $msinfo32_process = Get-Process msinfo32 -ErrorAction SilentlyContinue

                                    If ($msinfo32_process) {

                                        # Try gracefully first
                                        $msinfo32_process.CloseMainWindow()
                                        Start-Sleep -Seconds 3

                                            # Close msinfo32
                                            If ( -not $msinfo32_process.HasExited) {
                                                $msinfo32_process | Stop-Process -Force
                                            } Else {
                                                $continue = $true
                                            } # Else (If -not)
                                    } Else {
                                        $continue = $true
                                    } # Else (If $msinfo32_process)

                            # Update the progress bar and start a timer
                            # Source: https://msdn.microsoft.com/en-us/library/system.diagnostics.stopwatch(v=vs.110).aspx
                            $timer = [System.Diagnostics.Stopwatch]::StartNew()
                            $activity = "Processing Additional Options $task_number/$activities"
                            $script = "& msinfo32.exe /categories +systemsummary /report '$ms_info_txt'"
                            Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $script -PercentComplete (($task_number / $total_steps) * 100)
                            $empty_line | Out-String
                            Write-Verbose "Please hold on, the ms_info.txt file creation probably runs well over a minute..." -verbose




            Try {
                # .txt file
                If ($PSVersionTable.PSVersion -ge 5.1) {
                    & msinfo32.exe /categories +systemsummary /report "$ms_info_txt"
                } Else {
                    & msinfo32.exe /categories +systemsummary /report "$ms_info_txt" | Out-Null
                } # Else (If $PSVersionTable.PSVersion)                   
            } Catch { Write-Debug $_.Exception }

                            do {    $processes = Get-Process
                                    $time_elapsed = $timer.Elapsed

                                    # Update the progress bar                                                         # Credit: Jeff: "Powershell show elapsed time"
                                    Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation "$([string]::Format("$script | Time Elapsed: {0:d2}:{1:d2}:{2:d2}", $time_elapsed.Hours, $time_elapsed.Minutes, $time_elapsed.Seconds))" -PercentComplete (($task_number / $total_steps) * 100)
                                    Start-Sleep -Seconds 1
                            }
                            while  ( $processes.Name -contains 'msinfo32' )

            # Increment the counters
            $task_number++
            $switch_count++
                                    # Close all msinfo32 instances
                                    # Source: http://stackoverflow.com/questions/28481811/how-to-correctly-check-if-a-process-is-running-and-stop-it
                                    $msinfo32_process = Get-Process msinfo32 -ErrorAction SilentlyContinue

                                    If ($msinfo32_process) {

                                        # Try gracefully first
                                        $msinfo32_process.CloseMainWindow()
                                        Start-Sleep -Seconds 3

                                            # Close msinfo32
                                            If ( -not $msinfo32_process.HasExited) {
                                                $msinfo32_process | Stop-Process -Force
                                            } Else {
                                                $continue = $true
                                            } # Else (If -not)
                                    } Else {
                                        $continue = $true
                                    } # Else (If $msinfo32_process)

                            # Update the progress bar
                            $timer.Stop()
                            $timer.Reset()
                            $timer.Start()
                            $activity = "Processing Additional Options $task_number/$activities"
                            $script = "& msinfo32.exe /categories +systemsummary /nfo '$ms_info_nfo'"
                            Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $script -PercentComplete (($task_number / $total_steps) * 100)
                            $empty_line | Out-String
                            Write-Verbose "Please hold on, the ms_info.nfo file creation probably runs for a couple of minutes..." -verbose
            Try {
                # .nfo file
                If ($PSVersionTable.PSVersion -ge 5.1) {
                    & msinfo32.exe /categories +systemsummary /nfo "$ms_info_nfo"
                } Else {
                    & msinfo32.exe /categories +systemsummary /nfo "$ms_info_nfo" | Out-Null
                } # Else (If $PSVersionTable.PSVersion)      
            } Catch { Write-Debug $_.Exception }

                            do {    $processes = Get-Process
                                    $time_elapsed = $timer.Elapsed

                                    # Update the progress bar                                                         # Credit: Jeff: "Powershell show elapsed time"
                                    Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation "$([string]::Format("$script | Time Elapsed: {0:d2}:{1:d2}:{2:d2}", $time_elapsed.Hours, $time_elapsed.Minutes, $time_elapsed.Seconds))" -PercentComplete (($task_number / $total_steps) * 100)
                                    Start-Sleep -Seconds 1
                            }
                            while  ( $processes.Name -contains 'msinfo32' )

                    # .nfo to .xml conversion
                    If ($PSVersionTable.PSVersion -ge 5.1) {
                                    Try {                                        
                                        $ms_info_xml = [System.IO.Path]::ChangeExtension($ms_info_nfo,"xml")
                                        $source_alfa = Get-Content $ms_info_nfo
                                        $source_alfa.Replace("<>", "<type>") | Out-File "$ms_info_nfo" -Encoding UTF8
                                        $source_beta = Get-Content $ms_info_nfo
                                        $source_beta.Replace("</>", "</type>") | Out-File "$ms_info_nfo" -Encoding UTF8
                                        $source_gamma = Get-Content $ms_info_nfo
                                        $source_gamma | Out-File "$ms_info_xml" -Encoding UTF8
                                        # [xml]$msinfo = Get-Content $ms_info_xml
                                        # $msinfo.MsInfo.Category.Category.Category
                                    } Catch { Write-Debug $_.Exception }
                    } Else {
                        $continue = $true
                    } # Else (If $PSVersionTable.PSVersion)                            
            $timer.Stop()
            $timer.Reset()
        } Else {
            $continue = $true
        } # Else (If $Extract)


    # (3) MsInfo32.exe window
    # Located at %CommonProgramFiles%\Microsoft Shared\MSInfo\ directory
    # $msinfo32_folder = [string]([Environment]::GetFolderPath("CommonProgramFiles")) + "\Microsoft Shared\MSInfo\"
    # Source: https://technet.microsoft.com/en-us/library/bb490937.aspx
    # Source: https://support.microsoft.com/en-us/help/300887/how-to-use-system-information-msinfo32-command-line-tool-switches


        If ($MsInfo32) {

            # Increment the counters
            $task_number++
            $switch_count++

                            # Update the progress bar
                            $activity = "Processing Additional Options $task_number/$activities"
                            $task = "msinfo32.exe"
                            Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $task -PercentComplete (($task_number / $total_steps) * 100)
            Try {
                msinfo32.exe
            } Catch { Write-Debug $_.Exception }
        } Else {
            $continue = $true
        } # Else


    # (4) GatherNetworkInfo.vbs
    # Located at %windir%\system32\ directory
    # $system32 = [Environment]::GetFolderPath("System")
    # Adds information to the %temp%\Config\ folder when manually run.
    # $folder = [string]$env:temp + "\Config\"
    # For best results the GatherNetworkInfo.vbs script should be be run in an elevated instance (cmd-prompt or PowerShell) of an administrator account.
    #       (1)     Log into your system as the administrator.
    #       (2)     Click Start, and in the Search box at the bottom do NOT press enter after typing:       CMD
    #       (3)     When CMD.EXE is displayed in the list above, right-click it and choose RUN AS ADMINISTRATOR which will launch the command prompt with elevated permissions.
    #       (4)     In the CMD prompt, press enter after typing the following command:                      CD /D %TEMP%
    #       (5)     Next, press enter after typing the following command (which will launch the script):    Cscript c:\windows\system32\gatherNetworkInfo.vbs
    #       (6)     The script will not produce any on-screen data. It might take some time (minutes, perhaps) to complete.
    #       (7)     Launch any File browser and open the  %TEMP%\Config\  folder.
    # Probably has a scheduled task in Task Scheduler (Control Panel > Administrative Tools > Task Scheduler): Task Scheduler Library\Microsoft\Windows\NetTrace\GatherNetworkInfo
    # Source: http://www.verboon.info/2011/06/the-gathernetworkinfo-vbs-script/
    # Source: https://technet.microsoft.com/en-us/library/ff920171(v=ws.11).aspx
    # Credit: Paul-De: "Does anyone know what gatherNetworkInfo.vbs is?"  https://answers.microsoft.com/en-us/windows/forum/windows_7-security/does-anyone-know-what-gathernetworkinfovbs-is-its/63a302a6-cf69-4b9a-a3ef-4b2aff1b2514


        If ($GatherNetworkInfo) {

            # Increment the counters
            $task_number++
            $switch_count++
            $folder = [string]$env:temp + "\Config\"
            $system32 = [Environment]::GetFolderPath("System")

                            # Update the progress bar and start a timer
                            # Source: https://msdn.microsoft.com/en-us/library/system.diagnostics.stopwatch(v=vs.110).aspx
                            # Source: http://stackoverflow.com/questions/28481811/how-to-correctly-check-if-a-process-is-running-and-stop-it
                            $timer = [System.Diagnostics.Stopwatch]::StartNew()
                            $activity = "Processing Additional Options $task_number/$activities"
                            $task = "Cscript $system32\gatherNetworkInfo.vbs //Nologo"
                            Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $task -PercentComplete (($task_number / $total_steps) * 100)

            # Notify the user if the PowerShell session is not elevated (has been run as an administrator)            # Credit: alejandro5042: "How to run exe with/without elevated privileges from PowerShell"
            If (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator") -eq $true) {
                $empty_line | Out-String
                Write-Verbose "Please hold on, the GatherNetworkInfo.vbs script probably runs for a few minutes..." -verbose
                If ($PSVersionTable.PSVersion -ge 5.1) {
                    Cscript $system32\gatherNetworkInfo.vbs //Nologo

                            <#
                            # Run the GatherNetworkInfo.vbs script as an background job
                            # Source: https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.core/start-job
                            # Source: https://msdn.microsoft.com/powershell/reference/5.1/Microsoft.PowerShell.Core/about/about_Jobs
                            Remove-Job -Name VBS -ErrorAction SilentlyContinue
                            Start-Job -Name VBS -Scriptblock { Invoke-Command -ScriptBlock { $system32 = [Environment]::GetFolderPath("System"); Cscript $system32\gatherNetworkInfo.vbs //Nologo }} | Out-Null

                                        do {    $jobs = Get-Job
                                                $time_elapsed = $timer.Elapsed

                                                # Update the progress bar                                                         # Credit: Jeff: "Powershell show elapsed time"
                                                Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation "$([string]::Format("Cscript $system32\gatherNetworkInfo.vbs //Nologo | Time Elapsed: {0:d2}:{1:d2}:{2:d2}", $time_elapsed.Hours, $time_elapsed.Minutes, $time_elapsed.Seconds))" -PercentComplete (($task_number / $total_steps) * 100)
                                                Start-Sleep -Seconds 1
                                        }
                                        while  ((( $jobs | where { $_.Name -eq 'VBS' }).State) -eq 'Running' )

                            $empty_line | Out-String
                            Remove-Job -Name VBS -ErrorAction SilentlyContinue
                            $timer.Stop()
                            $timer.Reset()
                            #>
                } Else {
                    Cscript $system32\gatherNetworkInfo.vbs //Nologo                                       
                } # Else
                            If ((Test-Path $folder) -eq $true){
                                $empty_line | Out-String 
                                Write-Verbose "The output of GatherNetworkInfo.vbs script may be found inside the '$folder' directory." -verbose
                            } Else {
                                "The content creation failed."
                            } # Else
            } Else {
                $empty_line | Out-String
                Write-Warning "It seems that this script is run in a 'normal' PowerShell window."
                $empty_line | Out-String
                Write-Verbose "Please consider running this './Get-ComputerInfo' script in an elevated (administrator-level) PowerShell window, so that the called GatherNetworkInfo.vbs script can be run elevated as well." -verbose
                $empty_line | Out-String
                $admin_text = "For best results the GatherNetworkInfo.vbs script might be required to be run in an elevated window. An elevated PowerShell session can, for example, be initiated by starting PowerShell with the 'run as an administrator' option."
                Write-Output $admin_text
                $empty_line | Out-String
                $exeption_text = "Didn't run the GatherNetworkInfo.vbs script."
                Write-Output $exeption_text
            } # Else (If Security.Principal)

        } Else {
            $continue = $true
        } # Else (If $GatherNetworkInfo)


    # (5) Get-ComputerInfo cmdlet
    # Source: https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.management/get-computerinfo
    # Source: https://github.com/PowerShell/PowerShell/issues/3080
    # Source: https://blogs.technet.microsoft.com/askperf/2010/09/24/an-introduction-to-winrm-basics/
    # Source: https://blogs.technet.microsoft.com/otto/2007/02/09/a-few-good-vista-ws-man-winrm-commands/
    # Source: https://blogs.technet.microsoft.com/jonjor/2009/01/09/winrm-windows-remote-management-troubleshooting/
    # The Get-ComputerInfo cmdlet is passing localhost as the computername which (by design) goes through the winrm remoting stack. The cmdlet should be passing null instead to indicate it is a local call. If only a couple of values are displayed, winrm quickconfig (which is run from an Elevated Command prompt and which starts the Windows Remote Management service, sets the WinRM service type to auto start, creates a listener to accept requests on any IP address and enables firewall exception for WS-Management traffic (for http only)) might unveil most of the values, but could also increase the attack surface of the computer. For example, Quick Config configures a listener that accepts connections from every network interface, which is probably not ideal for edge machines that connect to unsecure networks (like the Internet). winrm invoke Restore winrm/Config @{} might close the WS-Man listener.
    # Manual commands inside the winrm quickconfig:
    #   sc config "WinRM" start= auto
    #   net start WinRM
    #   winrm create winrm/config/listener?Address=*+Transport=HTTP
    #   netsh firewall add portopening TCP 80 "Windows Remote Management"
    # It seems that Get-ComputerInfo apparently works normally (without winrm quickconfig) on Windows 10 Insider Preview Build 15007.rs_prerelease.170107-1846


        If ($Cmdlet) {

            # Increment the counters
            $task_number++
            $switch_count++
            $computer_info_original = "$real_output_path\computer_info_original.txt"
            $computer_info_txt = "$real_output_path\computer_info.txt"

                            # Update the progress bar
                            $activity = "Processing Additional Options $task_number/$activities"
                            $task = "Get-ComputerInfo"
                            Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $task -PercentComplete (($task_number / $total_steps) * 100)
            Try {
                # $win_rm = Get-Service "WinRM" -ErrorAction SilentlyContinue
                $test = Get-Command "Get-ComputerInfo" -ErrorAction SilentlyContinue
            } Catch { Write-Debug $_.Exception
            } Finally {

                If ($test -eq $null) {

                    $empty_line | Out-String
                    Write-Warning "It seems that the PowerShell version $($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor) doesn't contain the 'Get-ComputerInfo' cmdlet."
                    $empty_line | Out-String
                    $ps_text = "The 'Get-ComputerInfo' inbuilt cmdlet was first introcuded probably in PowerShell v3.1 or in PowerShell v5.1 at the latest. A command Get-Command Get-ComputerInfo might search for the cmdlet and `$PSVersionTable.PSVersion might reveal the PowerShell version."
                    Write-Output $ps_text
                    $empty_line | Out-String
                    $skip_text = "Didn't run the inbuilt 'Get-ComputerInfo' cmdlet. Please disregard any (system generated) suggestions concerning running .\Get-ComputerInfo that might occur below this line and the command prompt."
                    Write-Output $skip_text

                } Else {

                    # Run the Get-ComputerInfo as an background job with a timer
                    # Source: https://msdn.microsoft.com/en-us/library/system.diagnostics.stopwatch(v=vs.110).aspx
                    # Source: https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.core/start-job
                    # Source: https://msdn.microsoft.com/powershell/reference/5.1/Microsoft.PowerShell.Core/about/about_Jobs
                    $timer = [System.Diagnostics.Stopwatch]::StartNew()
                    Remove-Job -Command "Get-ComputerInfo" -ErrorAction SilentlyContinue
                    $job = Start-Job -ScriptBlock {Get-ComputerInfo}

                            do {    $jobs = Get-Job
                                    $time_elapsed = $timer.Elapsed

                                    # Update the progress bar                                                         # Credit: Jeff: "Powershell show elapsed time"
                                    Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation "$([string]::Format("Get-ComputerInfo | Time Elapsed: {0:d2}:{1:d2}:{2:d2}", $time_elapsed.Hours, $time_elapsed.Minutes, $time_elapsed.Seconds))" -PercentComplete (($task_number / $total_steps) * 100)
                                    Start-Sleep -Seconds 1
                            }
                            while  ((( $jobs | where { $_.Command -eq 'Get-ComputerInfo' }).State) -eq 'Running' )

                    # Output the Get-ComputerInfo background job results as a text file
                    #   \S      Any nonwhitespace character (excludes space, tab and carriage return).
                    #   \d      Any decimal digit.
                    # Source: http://powershellcookbook.com/recipe/qAxK/appendix-b-regular-expression-reference
                    $gin = Receive-Job -Job $job
                    $gin | Out-File "$computer_info_original" -Encoding UTF8
                    $gin_selection = Get-Content $computer_info_original | Where { $_ -match ': \S' }
                    $gin_sorted = $gin_selection | sort
                    $gin_sorted | Out-File "$computer_info_txt" -Encoding UTF8
                    Write-Output $gin_sorted
                    Remove-Job -Command "Get-ComputerInfo" -ErrorAction SilentlyContinue
                    $timer.Stop()
                    $timer.Reset()
                } # Else

            } # Finally (Try)
        } Else {
            $continue = $true
        } # Else

                            # Close the progress bar if it has been opened
                            If ($threshold -ge 2) {
                                $activity = "Processing Additional Options"
                                $task = "Finished processing additional options."
                                Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $task -PercentComplete (($total_steps / $total_steps) * 100) -Completed
                            } # If ($threshold)
} # End




# [End of Line]


<#

   ____        _   _
  / __ \      | | (_)
 | |  | |_ __ | |_ _  ___  _ __  ___
 | |  | | '_ \| __| |/ _ \| '_ \/ __|
 | |__| | |_) | |_| | (_) | | | \__ \
  \____/| .__/ \__|_|\___/|_| |_|___/
        | |
        |_|


# Write the partition table to a txt-file (in two steps)
$partition_table_selection | Format-Table -AutoSize | Out-File "$real_output_path\partition_table.txt" -Width 9000
$partition_table_selection | Format-List | Out-File "$real_output_path\partition_table.txt" -Append


# Write the volumes to a HTML-file and open the HTML-file in the default browser
$volumes_selection | Select-Object * | ConvertTo-Html | Out-File "$real_output_path\volumes.html"; & "$real_output_path\volumes.html"


# Display the results in two pop-up windows
$volumes_selection | Out-GridView
$obj_osinfo_selection | Out-GridView


# Open the Computer info CSV-file
Invoke-Item -Path $real_output_path\computer_info.csv


computer_info_$timestamp.csv                                                                  # an alternative filename format
computer_info_$timestamp.html                                                                 # an alternative filename format
$date = Get-Date -Format g                                                                    # a "general short" time-format (short date and short time)


   _____
  / ____|
 | (___   ___  _   _ _ __ ___ ___
  \___ \ / _ \| | | | '__/ __/ _ \
  ____) | (_) | |_| | | | (_|  __/
 |_____/ \___/ \__,_|_|  \___\___|


http://powershell.com/cs/media/p/7476.aspx                                                                          # clayman2: "Disk Space"
https://www.petri.com/validating-computer-lists-with-powershell                                                     # Jeff Hicks: "Validating Computer Lists with PowerShell"
https://answers.microsoft.com/en-us/windows/forum/windows_7-security/does-anyone-know-what-gathernetworkinfovbs-is-its/63a302a6-cf69-4b9a-a3ef-4b2aff1b2514    # Paul-De: "Does anyone know what gatherNetworkInfo.vbs is?"
http://stackoverflow.com/questions/29266622/how-to-run-exe-with-without-elevated-privileges-from-powershell?rq=1    # alejandro5042: "How to run exe with/without elevated privileges from PowerShell"
http://stackoverflow.com/questions/10941756/powershell-show-elapsed-time                                            # Jeff: "Powershell show elapsed time"


  _    _      _
 | |  | |    | |
 | |__| | ___| |_ __
 |  __  |/ _ \ | '_ \
 | |  | |  __/ | |_) |
 |_|  |_|\___|_| .__/
               | |
               |_|
#>

<#

.SYNOPSIS
Retrieves basic computer information from specified computers.

.DESCRIPTION
Get-ComputerInfo uses Windows Management Instrumentation (WMI) and reads the
"HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" registry key to retrieve
basic computer information, a list of volumes and partition tables of the
computers specified with the -Computer parameter (and/or inputted via a text
file with the -File parameter). The results are displayed on-screen and written
to a CSV- and a HTML-file. The default output destination folder $env:temp,
which points to the current temporary file location, may be changed with the
-Output parameter.

With five additional parameters (switches) the amount of gathered data may be
enlarged: -SystemInfo parameter will launch the systeminfo.exe /fo CSV Dos command,
-MsInfo32 parameter opens the System Information (msinfo32) window, -Extract
parameter will output the System Information (msinfo32.exe) data to a TXT- and
a NFO-file (and on machines running PowerShell version 5.1 or later convert the 
data to a XML-file). The -GatherNetworkInfo parameter will launch the native
GatherNetworkInfo.vbs script (which outputs to $env:temp\Config folder and doesn't
follow the -Output parameter) and -Cmdlet parameter will try to launch the native
PowerShell Get-ComputerInfo cmdlet and output its data to text files. This script
is based on clayman2's PowerShell script "Disk Space"
(http://powershell.com/cs/media/p/7476.aspx).

.PARAMETER Computer
with an alias -ComputerName. The -Computer parameter determines the objects (i.e.
the computers) for Get-ComputerInfo. To enter multiple computer names, please
separate each individual computer name with a comma. The -Computer parameter also
takes an array of strings and objects could be piped to this parameter, too.
If no value for the -Computer parameter is defined in the command launching
Get-ComputerInfo, the local machine will be defined as the -Computer parameter value.

.PARAMETER Output
with an alias -ReportPath. Specifies where most of the files are to be saved.
The default save location is $env:temp, which points to the current temporary file
location, which is set in the system. The default -Output save location is defined
at line 15 with the $Output variable. In case the path name includes space
characters, please enclose the path name in quotation marks (single or double).
For usage, please see the Examples below and for more information about $env:temp,
please see the Notes section below. Please note that the output folder for the
-GatherNetworkInfo parameter is hard coded inside the vbs script and cannot be
changed with -Output parameter.

.PARAMETER File
with aliases -ListOfComputersInATxtFile and -List. The -File parameter may be used
to define the path to a text file, which contains computer names or IP addresses
(one in each line). If the full filename or the directory name includes space
characters, please enclose the whole inputted string in quotation marks (single or
double).

.PARAMETER SystemInfo
If the -SystemInfo parameter is added to the command launching Get-ComputerInfo,
a systeminfo.exe /fo CSV Dos command is eventually launched, which outputs a
system_info.txt text file.

.PARAMETER Extract
with aliases -ExtractMsInfo32ToAFile, -ExtractMsInfo32, -MsInfo32ContentsToFile,
-MsInfo32Report, and -Expand. If the -Extract parameter is added to the command
launching Get-ComputerInfo, the data contained by the System Information
(msinfo32.exe) program is exported to ms_info.txt and ms_info.nfo files, and 
on machines running PowerShell version 5.1 or later the data is also converted 
to a XML-file. Please note that this step will have a drastical toll on the 
completion time of this script, because each of the three steps may run for minutes.

.PARAMETER MsInfo32
with aliases -OpenMsInfo32PopUpWindow and -Window. By adding the -MsInfo32 parameter
to the command launching Get-ComputerInfo, the System Information (msinfo32) window
may be opened.

.PARAMETER GatherNetworkInfo
with an alias -Vbs. If the -GatherNetworkInfo parameter is added to the command
launching Get-ComputerInfo, a native GatherNetworkInfo.vbs script (which outputs
to $env:temp\Config folder and doesn't follow the -Output parameter) is also
eventually executed when Get-ComputerInfo (this script) is run. The vbs script
resides in the %WINDOWS%\system32 directory and amasses an extensive amount of
computer related data to the %TEMP%\Config directory when run. On most Windows machines
the GatherNetworkInfo.vbs script has by default a passive scheduled task in the
Task Scheduler (i.e. Control Panel > Administrative Tools > Task Scheduler), which
for instance can be seen by opening inside the Task Scheduler a
Task Scheduler Library > Microsoft > Windows > NetTrace > GatherNetworkInfo tab.
The GatherNetworkInfo.vbs script will probably run for a few minutes. Please note
that it's mandatory to run the GatherNetworkInfo.vbs in an elevated instance 
(an elevated cmd-prompt or an elevated PowerShell window) for best results.

.PARAMETER Cmdlet
with aliases -GetComputerInfoCmdlet and -GetComputerInfo. The parameter -Cmdlet
will try to launch the native PowerShell Get-ComputerInfo cmdlet and output its
data to computer_info.txt and computer_info_original.txt text files. Please note
that the inbuilt Get-ComputerInfo cmdlet was first introcuded probably in
PowerShell v3.1 or in PowerShell v5.1 at the latest. The
Get-Command 'Get-ComputerInfo'
command may search for this cmdlet and $PSVersionTable.PSVersion may reveal
the PowerShell version.

.OUTPUTS
Displays general computer information and a volumes list in console. Opens the
generated HTML-file in the default browser. By default writes two files to
$env:temp or at the location specified with the -Output parameter.


    Default values:


$env:temp\computer_info.html            : HTML-file     : computer_info.html
$env:temp\computer_info.csv             : CSV-file      : computer_info.csv


    Optional files with the default -Output path (the files are generated, if the
    corresponding parameters (switches) are added to the command launching
    Get-ComputerInfo):


$env:temp\system_info.txt               -SystemInfo         : TXT-file  : system_info.txt
$env:temp\ms_info.txt                   -Extract            : TXT-file  : ms_info.txt
$env:temp\ms_info.nfo                   -Extract            : NFO-file  : ms_info.nfo
$env:temp\ms_info.xml                   -Extract            : XML-file  : ms_info.xml
$env:temp\computer_info.txt             -Cmdlet             : TXT-file  : computer_info.txt
$env:temp\computer_info_original.txt    -Cmdlet             : TXT-file  : computer_info_original.txt
$env:temp\Config                        -GatherNetworkInfo  : Folder    : Folder with files and a subfolder


.NOTES
Please note that all the parameters can be used in one get computer info command
and that each of the parameters can be "tab completed" before typing them fully (by
pressing the [tab] key).

Please note that the files (apart from the outputs of the -GatherNetworkInfo
parameter) are created in a directory, which is end-user settable in each get
computer info command with the -Output parameter. The default save location is
defined with the $Output variable (at line 15). The $env:temp variable points to
the current temp folder. The default value of the $env:temp variable is
C:\Users\<username>\AppData\Local\Temp (i.e. each user account has their own
separate temp folder at path %USERPROFILE%\AppData\Local\Temp). To see the current
temp path, for instance a command

    [System.IO.Path]::GetTempPath()

may be used at the PowerShell prompt window [PS>]. To change the temp folder for
instance to C:\Temp, please, for example, follow the instructions at
http://www.eightforums.com/tutorials/23500-temporary-files-folder-change-location-windows.html

    Homepage:           https://github.com/auberginehill/get-computer-info
    Short URL:          http://tinyurl.com/jxvhufb
    Version:            1.4

.EXAMPLE
./Get-ComputerInfo
Run the script. Please notice to insert ./ or .\ before the script name. Gathers
information about the local machine, displays the data in console, outputs the
default two files to the default -Output location ($env:temp) and opens the created
HTML-file in the default browser.

.EXAMPLE
help ./Get-ComputerInfo -Full
Display the help file.

.EXAMPLE
./Get-ComputerInfo -Computer dc01, dc02 -Output "E:\chiore" -SystemInfo -Extract -MsInfo32 -Vbs -Cmdlet
Run the script get all the available computer related information from the computers
dc01 and dc02. Save most of the results in the "E:\chiore" directory (the results of
the GatherNetworkInfo.vbs are saved to $env:temp\Config folder, if the command 
launching Get-ComputerInfo was run in an elevated PowerShell window). This command 
will work, because -Vbs is an alias of -GatherNetworkInfo. Since the path name 
doesn't contain any space characters, it doesn't need to be enveloped with quotation
marks, and furthermore, the word -Computer may be left out from this command, too, 
because the values dc01 and dc02 are accepted as computer names due to their position
(first).

.EXAMPLE
Set-ExecutionPolicy remotesigned
This command is altering the Windows PowerShell rights to enable script execution for
the default (LocalMachine) scope. Windows PowerShell has to be run with elevated rights
(run as an administrator) to actually be able to change the script execution properties.
The default value of the default (LocalMachine) scope is "Set-ExecutionPolicy restricted".


    Parameters:

    Restricted      Does not load configuration files or run scripts. Restricted is the default
                    execution policy.

    AllSigned       Requires that all scripts and configuration files be signed by a trusted
                    publisher, including scripts that you write on the local computer.

    RemoteSigned    Requires that all scripts and configuration files downloaded from the Internet
                    be signed by a trusted publisher.

    Unrestricted    Loads all configuration files and runs all scripts. If you run an unsigned
                    script that was downloaded from the Internet, you are prompted for permission
                    before it runs.

    Bypass          Nothing is blocked and there are no warnings or prompts.

    Undefined       Removes the currently assigned execution policy from the current scope.
                    This parameter will not remove an execution policy that is set in a Group
                    Policy scope.


For more information, please type "Get-ExecutionPolicy -List", "help Set-ExecutionPolicy -Full",
"help about_Execution_Policies" or visit https://technet.microsoft.com/en-us/library/hh849812.aspx
or http://go.microsoft.com/fwlink/?LinkID=135170.

.EXAMPLE
New-Item -ItemType File -Path C:\Temp\Get-ComputerInfo.ps1
Creates an empty ps1-file to the C:\Temp directory. The New-Item cmdlet has an inherent -NoClobber mode
built into it, so that the procedure will halt, if overwriting (replacing the contents) of an existing
file is about to happen. Overwriting a file with the New-Item cmdlet requires using the Force. If the
path name and/or the filename includes space characters, please enclose the whole -Path parameter value
in quotation marks (single or double):

    New-Item -ItemType File -Path "C:\Folder Name\Get-ComputerInfo.ps1"

For more information, please type "help New-Item -Full".

.LINK
http://powershell.com/cs/media/p/7476.aspx
https://answers.microsoft.com/en-us/windows/forum/windows_7-security/does-anyone-know-what-gathernetworkinfovbs-is-its/63a302a6-cf69-4b9a-a3ef-4b2aff1b2514
http://stackoverflow.com/questions/29266622/how-to-run-exe-with-without-elevated-privileges-from-powershell?rq=1
http://stackoverflow.com/questions/10941756/powershell-show-elapsed-time
http://learningpcs.blogspot.com/2011/10/powershell-get-wmiobject-and.html
https://4sysops.com/archives/windows-server-2012-server-core-part-5-tools/
https://social.technet.microsoft.com/Forums/windowsserver/en-US/f82e6f0b-ab97-424b-8e91-508d710e03b1/how-to-link-the-output-from-win32diskdrive-and-win32volume?forum=winserverpowershell
https://support.microsoft.com/en-us/help/300887/how-to-use-system-information-msinfo32-command-line-tool-switches
https://technet.microsoft.com/en-us/library/ff730960.aspx
https://technet.microsoft.com/en-us/library/bb491007.aspx
https://technet.microsoft.com/en-us/library/bb490937.aspx
https://msdn.microsoft.com/en-us/library/aa394102(v=vs.85).aspx
https://msdn.microsoft.com/en-us/library/aa394239(v=vs.85).aspx
https://msdn.microsoft.com/en-us/library/aa394474(v=vs.85).aspx
https://msdn.microsoft.com/en-us/library/aa394512(v=vs.85).aspx
https://msdn.microsoft.com/en-us/library/aa394360(v=vs.85).aspx
https://msdn.microsoft.com/en-us/library/aa394216(v=vs.85).aspx
https://technet.microsoft.com/en-us/library/ff920171(v=ws.11).aspx
https://msdn.microsoft.com/en-us/library/system.diagnostics.stopwatch(v=vs.110).aspx
https://msdn.microsoft.com/powershell/reference/5.1/microsoft.powershell.core/Where-Object
https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.core/start-job
https://msdn.microsoft.com/powershell/reference/5.1/Microsoft.PowerShell.Core/about/about_Jobs
https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.management/get-computerinfo
https://blogs.technet.microsoft.com/jonjor/2009/01/09/winrm-windows-remote-management-troubleshooting/
https://blogs.technet.microsoft.com/otto/2007/02/09/a-few-good-vista-ws-man-winrm-commands/
https://blogs.technet.microsoft.com/askperf/2010/09/24/an-introduction-to-winrm-basics/
http://stackoverflow.com/questions/28481811/how-to-correctly-check-if-a-process-is-running-and-stop-it
http://powershellcookbook.com/recipe/qAxK/appendix-b-regular-expression-reference
http://www.verboon.info/2011/06/the-gathernetworkinfo-vbs-script/
https://github.com/PowerShell/PowerShell/issues/3080
https://technet.microsoft.com/en-us/library/ee692804.aspx
http://stackoverflow.com/questions/27175137/powershellv2-remove-last-x-characters-from-a-string#32608908

#>
