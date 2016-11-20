<#
.\Get-ComputerInfo.ps1
#>

#########################################################################                     # Credit: clayman2: "Disk Space"

# Change the following variables based on your environment


$path = $env:temp


# Change the following variables for the style of the report.
$background_color = "#FFFFFF"                    # hex format will probably give more consistent results in different browsers
$title_font_family = "Gill Sans"
$title_font_size = "19px"
$title_bg_color = "#FFFFFF"                      # hex format will probably give more consistent results in different browsers
$heading_font_family = "Arial"
$heading_font_size = "12px"
$heading_name_bg_color = "#FFFFFF"               # hex format will probably give more consistent results in different browsers
$data_font_family = "Calibri"
$data_font_size = "11px"
$data_alternating_row_color_odd = "#cccccc"      # hex format will probably give more consistent results in different browsers
$data_alternating_row_color_even = "#FFFFFF"     # hex format will probably give more consistent results in different browsers


# Colors for free space
$very_low_space = "#b81321"                      # very low space is less than 1 GB or less than 5 % free
$low_space = "#ffca00"                           # low space is less than 5 GB or less than 10 % free
$medium_space = "#137abb"                        # medium space is less than 10 GB or less than 15 % free


#########################################################################


# Set the common parameters
$timestamp = Get-Date -UFormat "%Y%m%d"
$date = Get-Date -Format g
$time = (Get-Date -UFormat "%d. %m. %Y klo %H.%M")
$separator = '--------------------'
$path_text = "Path          : $path\"
$empty_line = ""


### If remote computers are specified, this script will use Windows Management Instrumentation (WMI) over Remote Procedure Calls (RPCs)
# $name_list = Get-Content 'C:\Temp\servers.txt'                                              # These can be computer names or IP addresses, one in each line
$name_list = $env:COMPUTERNAME




# Function used to convert bytes to MB or GB or TB                                            # Credit: clayman2: "Disk Space"
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




# Display a welcoming screen in console
$empty_line | Out-String
$title = 'Computer Info'
Write-Output $title
$separator | Out-String




# Gather basic computer information with WMI and display the computer information in console
$obj_osinfo = @()
$obj_volumes = @()


ForEach ($computer in $name_list) {

    # Retrieve basic os and computer related information and display it in console
    $bios = Get-WmiObject -class Win32_BIOS -ComputerName $computer
    $compsys = Get-WmiObject -class Win32_ComputerSystem -ComputerName $computer
    $compsysprod = Get-WMIObject -class Win32_ComputerSystemProduct -ComputerName $computer
    $enclosure = Get-WmiObject -Class Win32_SystemEnclosure -ComputerName $computer
    $motherboard = Get-WmiObject -class Win32_BaseBoard -ComputerName $computer
    $os = Get-WmiObject -class Win32_OperatingSystem -ComputerName $computer
    $processor = Get-WMIObject -class Win32_Processor -ComputerName $computer
    $timezone = Get-WmiObject -class Win32_TimeZone -ComputerName $computer


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


            Switch ($os.ProductType) {
                { $_ -lt 1 } { $product_type = "" }
                { $_ -eq 1 } { $product_type = "Work Station" }
                { $_ -eq 2 } { $product_type = "Domain Controller" }
                { $_ -eq 3 } { $product_type = "Server" }
                { $_ -gt 3 } { $product_type = "" }
            } # switch producttype


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

                    If ($enclosure | Where-Object { $_.ChassisTypes -eq 9 -or $_.ChassisTypes -eq 10 -or $_.ChassisTypes -eq 14}) {
                        $is_a_laptop = $true
                    } # if


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


            # CPU
            $CPUArchitecture_data = $processor.Name
            If ($CPUArchitecture_data.Contains("(TM)")) {
                $CPUArchitecture_data = $CPUArchitecture_data.Replace("(TM)","")
                } If ($CPUArchitecture_data.Contains("(R)")) {
                        $CPUArchitecture_data = $CPUArchitecture_data.Replace("(R)","")
            } # if (CPUArchitecture_data)


            # Manufacturer
            $Manufacturer_data = $compsysprod.Vendor
            If ($Manufacturer_data.Contains("HP")) {
                $Manufacturer_data = $Manufacturer_data.Replace("HP","Hewlett-Packard")
            } # if


            # Operating System
            $OperatingSystem_data = $os.Caption
            If ($OperatingSystem_data.Contains(",")) {
                $OperatingSystem_data = $OperatingSystem_data.Replace(",","")
                } If ($OperatingSystem_data.Contains("(R)")) {
                        $OperatingSystem_data = $OperatingSystem_data.Replace("(R)","")
            } # if (OperatingSystem_data)


            # $osa = Get-WmiObject win32_operatingsystem | Select @{Name='InstallDate';Expression={$_.ConvertToDateTime($_.InstallDate)}}
            # $InstallDate_Local = $osa.InstallDate


                    $obj_osinfo += New-Object -TypeName PSCustomObject -Property @{
                        'Computer'                      = $computer
                        'Manufacturer'                  = $Manufacturer_data
                        'Computer Model'                = $compsys.Model
                        'System Type'                   = $compsys.SystemType
                        'Domain Role'                   = $domain_role
                        'Product Type'                  = $product_type
                        'Chassis'                       = $chassis
                        'PC Type'                       = $pc_type
                        'Is a Laptop?'                  = $is_a_laptop
                        'CPU'                           = $CPUArchitecture_data
                        'Operating System'              = $OperatingSystem_data
                        'Architecture'                  = $os.OSArchitecture
                        'SP Version'                    = $os.CSDVersion
                        'Build Number'                  = $os.BuildNumber
                        'Memory'                        = (ConvertBytes($compsys.TotalPhysicalMemory))
                        'Processors'                    = $processor.NumberOfLogicalProcessors
                        'Cores'                         = $processor.NumberOfCores
                        'Country Code'                  = $os.CountryCode
                        'OS Install Date'               = ($os.ConvertToDateTime($os.InstallDate)).ToShortDateString()
                        'Last BootUp'                   = (($os.ConvertToDateTime($os.LastBootUpTime)).ToShortDateString() + ' ' + ($os.ConvertToDateTime($os.LastBootUpTime)).ToShortTimeString())
                        'UpTime'                        = (Uptime)
                        'Date'                          = $date
                        'Daylight Bias'                 = ((DayLight($timezone.DaylightBias)) + ' (' + $timezone.DaylightName + ')')
                        'Time Offset (Current)'         = (DayLight($timezone.Bias))
                        'Time Offset (Normal)'          = (DayLight($os.CurrentTimeZone))
                        'Time (Current)'                = (Get-Date).ToShortTimeString()
                        'Time (Normal)'                 = (((Get-Date).AddMinutes($timezone.DaylightBias)).ToShortTimeString() + ' (' + $timezone.StandardName + ')')
                        'Daylight In Effect'            = $compsys.DaylightInEffect
                     #  'Daylight In Effect'            = (Get-Date).IsDaylightSavingTime()
                        'Time Zone'                     = $timezone.Description
                        'OS Version'                    = $os.Version
                        'BIOS Version'                  = $bios.SMBIOSBIOSVersion
                        'ID'                            = $compsysprod.IdentifyingNumber
                        'Serial Number (BIOS)'          = $bios.SerialNumber
                        'Serial Number (Mother Board)'  = $motherboard.SerialNumber
                        'Serial Number (OS)'            = $os.SerialNumber
                        'UUID'                          = $compsysprod.UUID
                    } # New-Object


                $obj_osinfo.PSObject.TypeNames.Insert(0,"OSInfo")
                $obj_osinfo_selection = $obj_osinfo | Select-Object 'Computer','Manufacturer','Computer Model','System Type','Domain Role','Product Type','Chassis','PC Type','Is a Laptop?','CPU','Operating System','Architecture','SP Version','Build Number','Memory','Processors','Cores','Country Code','OS Install Date','Last BootUp','UpTime','Date','Daylight Bias','Time Offset (Current)','Time Offset (Normal)','Time (Current)','Time (Normal)','Daylight In Effect','Time Zone','OS Version','BIOS Version','Serial Number (BIOS)','Serial Number (Mother Board)','Serial Number (OS)','UUID'


                # Display OS Info in console
                Write-Output $obj_osinfo_selection
                $empty_line | Out-String
                $empty_line | Out-String




    # Retrieve additional disk information from volumes (Win32_Volume)

    $volumes = Get-WmiObject -class Win32_Volume -ComputerName $computer

            ForEach ($volume in $volumes) {
                $obj_volumes += New-Object -TypeName PSCustomObject -Property @{
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
                $obj_volumes.PSObject.TypeNames.Insert(0,"Volume")
            } # ForEach ($volume}
} # ForEach ($computer/first)




# Display the volumes in console
$volumes_selection = $obj_volumes | Sort Computer,Drive | Select-Object Computer,Drive,Label,'File System','System Volume','Boot Volume','Indexing Enabled','PageFile Present','Block Size','Compressed','Automount',Used,'Used (%)','Total Size','Free Space','Free (%)'
$volumes_selection_screen = $obj_volumes | Sort Computer,Drive | Select-Object Computer,Drive,Label,'File System','System Volume',Used,'Used (%)','Total Size','Free Space','Free (%)'
$volumes_header = 'Volumes'
Write-Output $volumes_header
$separator | Out-String
$volumes_selection_screen | Format-Table -AutoSize | Out-String
$empty_line | Out-String
$empty_line | Out-String




### Display the results in two pop-up windows
# $volumes_selection | Out-GridView
# $obj_osinfo_selection | Out-GridView




Write-Verbose "Results will be written to $path\computer_info.csv and $path\computer_info.html."


### Write the Computer info to a landscape-oriented CSV-file (horizontal).
$obj_osinfo_selection | Export-Csv $path\computer_info.csv -Delimiter ';' -NoTypeInformation -Encoding UTF8
# $landscape = Import-Csv -Delimiter ';' $path\computer_info.csv
# Write-Output $landscape




# Write the Computer info and a partition table to a HTML-file
# $html_file = New-Item -ItemType File -Path "$path\computer_info_$date.html" -Force                # an alternative filename format
$html_file = New-Item -ItemType File -Path "$path\computer_info.html" -Force


# Display the path and create the HTML-file
$path_text | Out-String
$html_file


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
Add-Content $html_file -Value ('
<h3>Computer Info</h3>
<table class="stats">
    <tr>
        <th>Generated:</th>
        <td>' + $time + '</td>
    </tr>
    <tr>
        <th>Computer:</th>
        <td>' + $name_list  + '</td>
    </tr>
    <tr>
        <th>Manufacturer:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Manufacturer") + '</td>
    </tr>
    <tr>
        <th>Computer Model:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Computer Model") + '</td>
    </tr>
    <tr>
        <th>System Type:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "System Type") + '</td>
    </tr>
    <tr>
        <th>Domain Role:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Domain Role") + '</td>
    </tr>
    <tr>
        <th>Product Type:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Product Type") + '</td>
    </tr>
    <tr>
        <th>Chassis:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Chassis") + '</td>
    </tr>
    <tr>
        <th>PC Type:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "PC Type") + '</td>
    </tr>
    <tr>
        <th>Is a Laptop?</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Is a Laptop?") + '</td>
    </tr>
    <tr>
        <th>CPU:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "CPU") + '</td>
    </tr>
    <tr>
        <th>Operating System:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Operating System") + '</td>
    </tr>
    <tr>
        <th>Architecture:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Architecture") + '</td>
    </tr>
    <tr>
        <th>SP Version:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "SP Version") + '</td>
    </tr>
    <tr>
        <th>Build Number:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Build Number") + '</td>
    </tr>
    <tr>
        <th>Memory:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Memory") + '</td>
    </tr>
    <tr>
        <th>Processors:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Processors") + '</td>
    </tr>
    <tr>
        <th>Cores:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Cores") + '</td>
    </tr>
    <tr>
        <th>Country Code:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Country Code") + '</td>
    </tr>
    <tr>
        <th>OS Install Date:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "OS Install Date") + '</td>
    </tr>
    <tr>
        <th>Last BootUp:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Last BootUp") + '</td>
    </tr>
    <tr>
        <th>UpTime:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "UpTime") + '</td>
    </tr>
    <tr>
        <th>Date:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Date") + '</td>
    </tr>
    <tr>
        <th>Daylight Bias:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Daylight Bias") + '</td>
    </tr>
    <tr>
        <th>Time Offset (Current):</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Time Offset (Current)") + '</td>
    </tr>
    <tr>
        <th>Time Offset (Normal):</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Time Offset (Normal)") + '</td>
    </tr>
    <tr>
        <th>Time (Current):</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Time (Current)") + '</td>
    </tr>
    <tr>
        <th>Time (Normal):</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Time (Normal)") + '</td>
    </tr>
    <tr>
        <th>Daylight In Effect:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Daylight In Effect") + '</td>
    </tr>
    <tr>
        <th>Time Zone:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Time Zone") + '</td>
    </tr>
    <tr>
        <th>OS Version:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "OS Version") + '</td>
    </tr>
    <tr>
        <th>BIOS Version:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "BIOS Version") + '</td>
    </tr>
    <tr>
        <th>Serial Number (BIOS):</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Serial Number (BIOS)") + '</td>
    </tr>
    <tr>
        <th>Serial Number (Mother Board):</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Serial Number (Mother Board)") + '</td>
    </tr>
    <tr>
        <th>Serial Number (OS):</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "Serial Number (OS)") + '</td>
    </tr>
    <tr>
        <th>UUID:</th>
        <td>' + ($obj_osinfo | Select-Object -ExpandProperty "UUID") + '</td>
    </tr>
</table>
<br />
<br />


<table id="main">
    <tr>
        <td colspan="15" class="title">' + $name_list  + '</td>
    </tr>
    <tr>
        <td class="headings">Computer</td>
        <td class="headings">Drive</td>
        <td class="headings">Label</td>
        <td class="headings">File System</td>
        <td class="headings">Description</td>
        <td class="headings">Partition</td>
        <td class="headings">Disk</td>
        <td class="headings">Disk Model</td>
        <td class="headings">Compressed</td>
        <td class="headings">Used</td>
        <td class="headings">Used %</td>
        <td class="headings">Status</td>
        <td class="headings">Total Size</td>
        <td class="headings">Free Space</td>
        <td class="headings">Free %</td>
    </tr>')




# Create a partition table with WMI
$partition_table = @()


ForEach ($computer in $name_list) {

    $disks = Get-WmiObject -class Win32_DiskDrive -ComputerName $computer

    ForEach ($disk in $disks) {

            # $partitions = Get-WmiObject -class Win32_DiskPartition -ComputerName $computer
            # $partitions = (Get-WmiObject -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='\\.\PHYSICALDRIVE0'} WHERE AssocClass=Win32_DiskDriveToDiskPartition")
            # $partitions = (Get-WmiObject -ComputerName $computer -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='\\.\PHYSICALDRIVE0'} WHERE ResultRole=Dependent")
            $partitions = (Get-WmiObject -ComputerName $computer -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='$($disk.DeviceID)'} WHERE ResultRole=Dependent")

            ForEach ($partition in ($partitions)) {

                    # $drives = Get-WmiObject -class Win32_LogicalDisk -ComputerName $computer
                    # $drives = Get-WmiObject -ComputerName $computer -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='Disk #0, Partition #1'} WHERE ResultRole=Dependent"
                    $drives = (Get-WmiObject -ComputerName $computer -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='$($partition.DeviceID)'} WHERE AssocClass=Win32_LogicalDiskToPartition")

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

                            $status = ""
                                If ($used_percentage -eq 100) {
                                    $status = "Full"
                                } ElseIf ($used_percentage -ge 95) {
                                    $status = "Very Low Space"
                                } ElseIf ($used_percentage -ge 90) {
                                    $status = "Low Space"
                                } ElseIf ($used_percentage -ge 85) {
                                    $status = "Medium Space"
                                } ElseIf ($used_percentage -gt 0) {
                                    $status = "OK"
                                } Else {
                                    $status = ""
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




                                $partition_table += New-Object -Type PSCustomObject -Property @{

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
                                    'Free Space Status'       = $status
                                    'Total Size'              = (ConvertBytes($drive.Size))
                                    'Used'                    = (ConvertBytes($drive.Size - $drive.FreeSpace))
                                    'Used (%)'                = $used_percentage
                                    'Serial Number (Volume)'  = $drive.VolumeSerialNumber

                                    } # New-Object
                                $partition_table.PSObject.TypeNames.Insert(0,"PartitionTable")
                                $partition_table_selection = $partition_table | Sort Computer,Drive | Select-Object Computer,Drive,Label,'File System','Boot Partition',Interface,'Media Type','Partition Type',Partition,Used,'Used (%)','Free Space Status','Total Size','Free Space','Free (%)'
                                $partition_table_selection_screen = $partition_table | Sort Computer,Drive | Select-Object Computer,Drive,Label,Interface,'Media Type',Partition,'Used (%)','Total Size','Free Space','Free (%)'




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
        <td class="data">' + $status + '</td>
        <td class="data">' + (ConvertBytes($drive.Size)) + '</td>
        <td class="data" bgcolor="' + $space_color + '">' + (ConvertBytes($drive.FreeSpace)) + '</td>
        <td class="data" bgcolor="' + $percentage_color + '">' + $free_percentage + '</td>
    </tr>')

                    } # ForEach ($drive)
            } # ForEach ($partition)
    } # ForEach ($disk)
} # ForEach ($computer/first)




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
        <td>' + $time + '</td>
    </tr>
    <tr>
        <th>Computer:</th>
        <td>' + $computer  + '</td>
    </tr>
</table>

</body>
</html>')




# Display the HTML-file in the default browser
# & $Path\time_zones.html
Start-Process -FilePath "$path\computer_info.html" | Out-Null




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
$partition_table_selection | Format-Table -AutoSize | Out-File $path\partition_table.txt -Width 9000
$partition_table_selection | Format-List | Out-File $path\partition_table.txt -Append


# Write the volumes to a HTML-file and open the HTML-file in the default browser
$volumes_selection | Select-Object * | ConvertTo-Html | Out-File $path\volumes.html; & "$path\volumes.html"


# Open the Computer info HTML-file in the default browser.
Invoke-Item "$path\computer_info.html"


# Open the Computer info CSV-file
Invoke-Item -Path $path\computer_info.csv


computer_info_$timestamp.csv                                                                  # an alternative filename format
computer_info_$timestamp.html                                                                 # an alternative filename format
$time = Get-Date -Format g                                                                    # a "general short" time-format (short date and short time)



   _____
  / ____|
 | (___   ___  _   _ _ __ ___ ___
  \___ \ / _ \| | | | '__/ __/ _ \
  ____) | (_) | |_| | | | (_|  __/
 |_____/ \___/ \__,_|_|  \___\___|


http://powershell.com/cs/media/p/7476.aspx                                                    # clayman2: "Disk Space"



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
Retrieves basic computer information, a list of volumes and the partition table.

.DESCRIPTION
Get-ComputerInfo uses Windows Management Instrumentation (WMI) to retrieve basic
computer information, a list of volumes and the partition table and displays
results on-screen and writes the results in a CSV- and HTML-file. This script
is based on clayman2's PowerShell script "Disk Space"
(http://powershell.com/cs/media/p/7476.aspx).

.OUTPUTS
Displays general computer information and a volumes list in console. Opens the 
generated HTML-file in the default browser. In addition
to that...


The aforementioned HTML-file and one CSV-file at $path

$env:temp\computer_info.html           : HTML-file               : computer_info.html
$env:temp\computer_info.csv            : CSV-file                : computer_info.csv


.NOTES
Please note that the two files are created in a directory, which is specified with the
$path variable (at line 10). The $env:temp variable points to the current temp folder.
The default value of the $env:temp variable is C:\Users\<username>\AppData\Local\Temp
(i.e. each user account has their own separate temp folder at path %USERPROFILE%\AppData\Local\Temp).
To see the current temp path, for instance a command

    [System.IO.Path]::GetTempPath()

may be used at the PowerShell prompt window [PS>]. To change the temp folder for instance
to C:\Temp, please, for example, follow the instructions at
http://www.eightforums.com/tutorials/23500-temporary-files-folder-change-location-windows.html

    Homepage:           https://github.com/auberginehill/get-computer-info
    Version:            1.1

.EXAMPLE
./Get-ComputerInfo
Run the script. Please notice to insert ./ or .\ before the script name.

.EXAMPLE
help ./Get-ComputerInfo -Full
Display the help file.

.EXAMPLE
Set-ExecutionPolicy remotesigned
This command is altering the Windows PowerShell rights to enable script execution. Windows PowerShell
has to be run with elevated rights (run as an administrator) to actually be able to change the script
execution properties. The default value is "Set-ExecutionPolicy restricted".


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


For more information,
type "help Set-ExecutionPolicy -Full" or visit https://technet.microsoft.com/en-us/library/hh849812.aspx.

.EXAMPLE
New-Item -ItemType File -Path C:\Temp\Get-ComputerInfo.ps1
Creates an empty ps1-file to the C:\Temp directory. The New-Item cmdlet has an inherent -NoClobber mode
built into it, so that the procedure will halt, if overwriting (replacing the contents) of an existing
file is about to happen. Overwriting a file with the New-Item cmdlet requires using the Force.
For more information, please type "help New-Item -Full".

.LINK
http://powershell.com/cs/media/p/7476.aspx
http://learningpcs.blogspot.com/2011/10/powershell-get-wmiobject-and.html
https://social.technet.microsoft.com/Forums/windowsserver/en-US/f82e6f0b-ab97-424b-8e91-508d710e03b1/how-to-link-the-output-from-win32diskdrive-and-win32volume?forum=winserverpowershell
https://technet.microsoft.com/en-us/library/ff730960.aspx
https://msdn.microsoft.com/en-us/library/aa394102(v=vs.85).aspx
https://msdn.microsoft.com/en-us/library/aa394239(v=vs.85).aspx
https://msdn.microsoft.com/en-us/library/aa394474(v=vs.85).aspx

#>
