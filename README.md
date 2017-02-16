<!-- Visual Studio Code: For a more comfortable reading experience, use the key combination Ctrl + Shift + V
     Visual Studio Code: To crop the tailing end space characters out, please use the key combination Ctrl + A Ctrl + K Ctrl + X (Formerly Ctrl + Shift + X)
     Visual Studio Code: To improve the formatting of HTML code, press Shift + Alt + F and the selected area will be reformatted in a html file.
     Visual Studio Code shortcuts: http://code.visualstudio.com/docs/customization/keybindings (or https://aka.ms/vscodekeybindings)
     Visual Studio Code shortcut PDF (Windows): https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf

   _____      _           _____                            _          _____        __
  / ____|    | |         / ____|                          | |        |_   _|      / _|
 | |  __  ___| |_ ______| |     ___  _ __ ___  _ __  _   _| |_ ___ _ __| |  _ __ | |_ ___
 | | |_ |/ _ \ __|______| |    / _ \| '_ ` _ \| '_ \| | | | __/ _ \ '__| | | '_ \|  _/ _ \
 | |__| |  __/ |_       | |___| (_) | | | | | | |_) | |_| | ||  __/ | _| |_| | | | || (_) |
  \_____|\___|\__|       \_____\___/|_| |_| |_| .__/ \__,_|\__\___|_||_____|_| |_|_| \___/
                                              | |
                                              |_|                                                   -->


## Get-ComputerInfo.ps1

<table>
   <tr>
      <td style="padding:6px"><strong>OS:</strong></td>
      <td style="padding:6px">Windows</td>
   </tr>
   <tr>
      <td style="padding:6px"><strong>Type:</strong></td>
      <td style="padding:6px">A Windows PowerShell script</td>
   </tr>
   <tr>
      <td style="padding:6px"><strong>Language:</strong></td>
      <td style="padding:6px">Windows PowerShell</td>
   </tr>
   <tr>
      <td style="padding:6px"><strong>Description:</strong></td>
      <td style="padding:6px">Get-ComputerInfo uses Windows Management Instrumentation (WMI) and reads the "<code>HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion</code>" registry key to retrieve basic computer information, a list of volumes and partition tables of the computers specified with the <code>-Computer</code> parameter (and/or inputted via a text file with the <code>-File</code> parameter). The results are displayed on-screen and written to a CSV- and a HTML-file. The default output destination folder <code>$env:temp</code>, which points to the current temporary file location, may be changed with the <code>-Output</code> parameter.
      <br />
      <br />With five additional parameters (switches) the amount of gathered data may be enlarged: <code>-SystemInfo</code> parameter will launch the <code>systeminfo.exe /fo CSV</code> Dos command, <code>-MsInfo32</code> parameter opens the System Information (<code>msinfo32</code>) window, <code>-Extract</code> parameter will output the System Information (<code>msinfo32.exe</code>) data to a TXT- and a NFO-file (and on machines running PowerShell version 5.1 or later convert the data to a XML-file). The <code>-GatherNetworkInfo</code> parameter will launch the native <code>GatherNetworkInfo.vbs</code> script (which outputs to <code>$env:temp\Config</code> folder and doesn't follow the <code>-Output</code> parameter) and <code>-Cmdlet</code> parameter will try to launch the native PowerShell <code>Get-ComputerInfo</code> cmdlet and output its data to text files. This script is based on clayman2's PowerShell script "<a href="http://powershell.com/cs/media/p/7476.aspx">Disk Space</a>" (or one of the <a href="http://web.archive.org/web/20120304222258/http://powershell.com/cs/media/p/7476.aspx">archive.org versions</a>).</td>
   </tr>
   <tr>
      <td style="padding:6px"><strong>Homepage:</strong></td>
      <td style="padding:6px"><a href="https://github.com/auberginehill/get-computer-info">https://github.com/auberginehill/get-computer-info</a>
      <br />Short URL: <a href="http://tinyurl.com/jxvhufb">http://tinyurl.com/jxvhufb</a></td>
   </tr>
   <tr>
      <td style="padding:6px"><strong>Version:</strong></td>
      <td style="padding:6px">1.4</td>
   </tr>
   <tr>
        <td style="padding:6px"><strong>Sources:</strong></td>
        <td style="padding:6px">
            <table>
                <tr>
                    <td style="padding:6px">Emojis:</td>
                    <td style="padding:6px"><a href="https://github.com/auberginehill/emoji-table">Emoji Table</a></td>
                </tr>
                <tr>
                    <td style="padding:6px">clayman2:</td>
                    <td style="padding:6px"><a href="http://powershell.com/cs/media/p/7476.aspx">Disk Space</a> (or one of the <a href="http://web.archive.org/web/20120304222258/http://powershell.com/cs/media/p/7476.aspx">archive.org versions</a>)</td>
                </tr>
                <tr>
                    <td style="padding:6px">Jeff Hicks:</td>
                    <td style="padding:6px"><a href="https://www.petri.com/validating-computer-lists-with-powershell">Validating Computer Lists with PowerShell</a></td>
                </tr>
                <tr>
                    <td style="padding:6px">Paul-De:</td>
                    <td style="padding:6px"><a href="https://answers.microsoft.com/en-us/windows/forum/windows_7-security/does-anyone-know-what-gathernetworkinfovbs-is-its/63a302a6-cf69-4b9a-a3ef-4b2aff1b2514">Does anyone know what gatherNetworkInfo.vbs is?</a></td>
                </tr>
                <tr>
                    <td style="padding:6px">alejandro5042:</td>
                    <td style="padding:6px"><a href="http://stackoverflow.com/questions/29266622/how-to-run-exe-with-without-elevated-privileges-from-powershell?rq=1">How to run exe with/without elevated privileges from PowerShell</a></td>
                </tr>
                <tr>
                    <td style="padding:6px">Jeff:</td>
                    <td style="padding:6px"><a href="http://stackoverflow.com/questions/10941756/powershell-show-elapsed-time">Powershell show elapsed time</a></td>
                </tr>
            </table>
        </td>
   </tr>
   <tr>
      <td style="padding:6px"><strong>Downloads:</strong></td>
      <td style="padding:6px">For instance <a href="https://raw.githubusercontent.com/auberginehill/get-computer-info/master/Get-ComputerInfo.ps1">Get-ComputerInfo.ps1</a>. Or <a href="https://github.com/auberginehill/get-computer-info/archive/master.zip">everything as a .zip-file</a>.</td>
   </tr>
</table>




### Screenshot

<img class="screenshot" title="screenshot" alt="screenshot" height="100%" width="100%" src="https://raw.githubusercontent.com/auberginehill/get-computer-info/master/Get-ComputerInfo.png">




### Parameters

<table>
    <tr>
        <th>:triangular_ruler:</th>
        <td style="padding:6px">
            <ul>
                <li>
                    <h5>Parameter <code>-Computer</code></h5>
                    <p>with an alias <code>-ComputerName</code>. The <code>-Computer</code> parameter determines the objects (i.e. the computers) for Get-ComputerInfo. To enter multiple computer names, please separate each individual computer name with a comma. The <code>-Computer</code> parameter also takes an array of strings and objects could be piped to this parameter, too. If no value for the <code>-Computer</code> parameter is defined in the command launching Get-ComputerInfo, the local machine will be defined as the <code>-Computer</code> parameter value.</p>
                </li>
            </ul>
        </td>
    </tr>
    <tr>
        <th></th>
        <td style="padding:6px">
            <ul>
                <p>
                    <li>
                        <h5>Parameter <code>-Output</code></h5>
                        <p>with an alias <code>-ReportPath</code>. Specifies where most of the files are to be saved. The default save location is <code>$env:temp</code>, which points to the current temporary file location, which is set in the system. The default <code>-Output</code> save location is defined at line 15 with the <code>$Output</code> variable. In case the path name includes space characters, please enclose the path name in quotation marks (single or double). For usage, please see the Examples below and for more information about <code>$env:temp</code>, please see the Notes section below. Please note that the output folder for the <code>-GatherNetworkInfo</code> parameter is hard coded inside the vbs script and cannot be changed with <code>-Output</code> parameter.</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-File</code></h5>
                        <p>with aliases <code>-ListOfComputersInATxtFile</code> and <code>-List</code>. The <code>-File</code> parameter may be used to define the path to a text file, which contains computer names or IP addresses (one in each line). If the full filename or the directory name includes space characters, please enclose the whole inputted string in quotation marks (single or double).</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-SystemInfo</code></h5>
                        <p>If the <code>-SystemInfo</code> parameter is added to the command launching Get-ComputerInfo, a <code>systeminfo.exe /fo CSV</code> Dos command is eventually launched, which outputs a <code>system_info.txt</code> text file.</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-Extract</code></h5>
                        <p>with aliases <code>-ExtractMsInfo32ToAFile</code>, <code>-ExtractMsInfo32</code>, <code>-MsInfo32ContentsToFile</code>, <code>-MsInfo32Report</code> and <code>-Expand</code>. If the <code>-Extract</code> parameter is added to the command launching Get-ComputerInfo, the data contained by the System Information (<code>msinfo32.exe</code>) program is exported to <code>ms_info.txt</code> and <code>ms_info.nfo</code> files, and on machines running PowerShell version 5.1 or later the data is also converted to a XML-file. Please note that this step will have a drastical toll on the completion time of this script, because each of the three steps may run for minutes.</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-MsInfo32</code></h5>
                        <p>with aliases <code>-OpenMsInfo32PopUpWindow</code> and <code>-Window</code>. By adding the <code>-MsInfo32</code> parameter to the command launching Get-ComputerInfo, the System Information (<code>msinfo32</code>) window may be opened.</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-GatherNetworkInfo</code></h5>
                        <p>with an alias <code>-Vbs</code>. If the <code>-GatherNetworkInfo</code> parameter is added to the command launching Get-ComputerInfo, a native <code>GatherNetworkInfo.vbs</code> script (which outputs to <code>$env:temp\Config</code> folder and doesn't follow the <code>-Output</code> parameter) is also eventually executed when Get-ComputerInfo (this script) is run. The vbs script resides in the <code>%WINDOWS%\system32</code> directory and amasses an extensive amount of computer related data to the <code>%TEMP%\Config</code> directory when run. On most Windows machines the <code>GatherNetworkInfo.vbs</code> script has by default a passive scheduled task in the Task Scheduler (i.e. Control Panel → Administrative Tools → Task Scheduler), which for instance can be seen by opening inside the Task Scheduler a Task Scheduler Library → Microsoft → Windows → NetTrace → GatherNetworkInfo tab. The <code>GatherNetworkInfo.vbs</code> script will probably run for a few minutes. Please note that it's mandatory to run the GatherNetworkInfo.vbs in an elevated instance (an elevated <code>cmd</code>-prompt or an elevated PowerShell window) for best results.</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-Cmdlet</code></h5>
                        <p>with aliases <code>-GetComputerInfoCmdlet</code> and <code>-GetComputerInfo</code>. The parameter <code>-Cmdlet</code> will try to launch the native PowerShell <code>Get-ComputerInfo</code> cmdlet and output its data to <code>computer_info.txt</code> and <code>computer_info_original.txt</code> text files. Please note that the inbuilt <code>Get-ComputerInfo</code> cmdlet was first introcuded probably in PowerShell v3.1 or in PowerShell v5.1 at the latest. The <code>Get-Command 'Get-ComputerInfo'</code> command may search for this cmdlet and <code>$PSVersionTable.PSVersion</code> may reveal the PowerShell version.</p>
                    </li>
                </p>
            </ul>
        </td>
    </tr>
</table>




### Outputs

<table>
    <tr>
        <th>:arrow_right:</th>
        <td style="padding:6px">
            <ul>
                <li>Displays general computer information (such as Computer Name, Manufacturer, Computer Model, System Type, Domain Role, Product Type, Chassis, PC Type, whether the machine is a laptop or not (based on the chassis information), Model Version, CPU, Video Card, Resolution, Operating System, Architecture, Windows Edition ID, Windows Installation Type, Windows Platform, Type, SP Version, Windows BuildLab Extended, Windows BuildLab, Windows Build Branch, Windows Build Number, Windows Release Id, Current Version, Memory, Video Card Memory, Logical Processors, Cores, Physical Processors, Country Code, Video Card Driver Date, BIOS Release Date, OS Install Date, Last BootUp, UpTime, Date, Daylight Bias, Time Offset (Current), Time Offset (Normal), Time (Current), Time (Normal), Daylight In Effect, Time Zone, Connectivity (network adapters), Mobile Broadband, OS Version, PowerShell Version, Video Card Version, BIOS Version, Mother Board Version, Serial Number (BIOS), Serial Number (Mother Board), Serial Number (OS), UUID), and a list of volumes in console. Opens the generated HTML-file in the default browser. By default writes two files to <code>$env:temp</code> or at the location specified with the <code>-Output</code> parameter.</li>
            </ul>
        </td>
    </tr>
    <tr>
        <th></th>
        <td style="padding:6px">
            <ul>
                <p>
                    <li>Default values:</li>
                </p>
                <ol>
                    <p>
                        <table>
                            <tr>
                                <td style="padding:6px"><strong>Path</strong></td>
                                <td style="padding:6px"><strong>Type</strong></td>
                                <td style="padding:6px"><strong>Name</strong></td>
                            </tr>
                            <tr>
                                <td style="padding:6px"><code>$env:temp\computer_info.html</code></td>
                                <td style="padding:6px">HTML-file</td>
                                <td style="padding:6px"><code>computer_info.html</code></td>
                            </tr>
                            <tr>
                                <td style="padding:6px"><code>$env:temp\computer_info.csv</code></td>
                                <td style="padding:6px">CSV-file</td>
                                <td style="padding:6px"><code>computer_info.csv</code></td>
                            </tr>
                        </table>
                    </p>
                </ol>
                <p>
                    <li>Optional files with the default <code>-Output</code> path (the files are generated, if the corresponding parameters (switches) are added to the command launching Get-ComputerInfo):</li>
                </p>
                <ol>
                    <p>
                        <table>
                            <tr>
                                <td style="padding:6px"><strong>Path</strong></td>
                                <td style="padding:6px"><strong>Parameter (switch)</strong></td>
                                <td style="padding:6px"><strong>Type</strong></td>
                            </tr>
                            <tr>
                                <td style="padding:6px"><code>$env:temp\system_info.txt</code></td>
                                <td style="padding:6px"><code>-SystemInfo</code></td>
                                <td style="padding:6px">TXT-file</td>
                            </tr>
                            <tr>
                                <td style="padding:6px"><code>$env:temp\ms_info.txt</code></td>
                                <td style="padding:6px"><code>-Extract</code></td>
                                <td style="padding:6px">TXT-file</td>
                            </tr>
                            <tr>
                                <td style="padding:6px"><code>$env:temp\ms_info.nfo</code></td>
                                <td style="padding:6px"><code>-Extract</code></td>
                                <td style="padding:6px">NFO-file</td>
                            </tr>
                            <tr>
                                <td style="padding:6px"><code>$env:temp\ms_info.xml</code></td>
                                <td style="padding:6px"><code>-Extract</code></td>
                                <td style="padding:6px">XML-file</td>
                            </tr>
                            <tr>
                                <td style="padding:6px"><code>$env:temp\computer_info.txt</code></td>
                                <td style="padding:6px"><code>-Cmdlet</code></td>
                                <td style="padding:6px">TXT-file</td>
                            </tr>
                            <tr>
                                <td style="padding:6px"><code>$env:temp\computer_info_original.txt</code></td>
                                <td style="padding:6px"><code>-Cmdlet</code></td>
                                <td style="padding:6px">TXT-file</td>
                            </tr>
                            <tr>
                                <td style="padding:6px"><code>$env:temp\Config</code></td>
                                <td style="padding:6px"><code>-GatherNetworkInfo</code></td>
                                <td style="padding:6px">Folder with files and a subfolder</td>
                            </tr>
                        </table>
                    </p>
                </ol>
            </ul>
        </td>
    </tr>
</table>




### Notes

<table>
    <tr>
        <th>:warning:</th>
        <td style="padding:6px">
            <ul>
                <li>Please note that all the parameters can be used in one get computer info command and that each of the parameters can be "tab completed" before typing them fully (by pressing the <code>[tab]</code> key).</li>
            </ul>
        </td>
    </tr>
    <tr>
        <th></th>
        <td style="padding:6px">
            <ul>
                <p>
                    <li>Please note that the files (apart from the outputs of the <code>-GatherNetworkInfo</code> parameter) are created in a directory, which is end-user settable in each get computer info command with the <code>-Output</code> parameter. The default save location is defined with the <code>$Output</code> variable (at line 15). The <code>$env:temp</code> variable points to the current temp folder. The default value of the <code>$env:temp</code> variable is <code>C:\Users\&lt;username&gt;\AppData\Local\Temp</code> (i.e. each user account has their own separate temp folder at path <code>%USERPROFILE%\AppData\Local\Temp</code>). To see the current temp path, for instance a command
                    <br />
                    <br /><code>[System.IO.Path]::GetTempPath()</code>
                    <br />
                    <br />may be used at the PowerShell prompt window <code>[PS&gt;]</code>. To change the temp folder for instance to <code>C:\Temp</code>, please, for example, follow the instructions at <a href="http://www.eightforums.com/tutorials/23500-temporary-files-folder-change-location-windows.html">Temporary Files Folder - Change Location in Windows</a>, which in essence are something along the lines:
                        <ol>
                           <li>Right click on Computer and click on Properties (or select Start → Control Panel → System). In the resulting window with the basic information about the computer...</li>
                           <li>Click on Advanced system settings on the left panel and select Advanced tab on the resulting pop-up window.</li>
                           <li>Click on the button near the bottom labeled Environment Variables.</li>
                           <li>In the topmost section labeled User variables both TMP and TEMP may be seen. Each different login account is assigned its own temporary locations. These values can be changed by double clicking a value or by highlighting a value and selecting Edit. The specified path will be used by Windows and many other programs for temporary files. It's advisable to set the same value (a directory path) for both TMP and TEMP.</li>
                           <li>Any running programs need to be restarted for the new values to take effect. In fact, probably also Windows itself needs to be restarted for it to begin using the new values for its own temporary files.</li>
                        </ol>
                    </li>
                </p>
            </ul>
        </td>
    </tr>
</table>




### Examples

<table>
    <tr>
        <th>:book:</th>
        <td style="padding:6px">To open this code in Windows PowerShell, for instance:</td>
   </tr>
   <tr>
        <th></th>
        <td style="padding:6px">
            <ol>
                <p>
                    <li><code>./Get-ComputerInfo</code><br />
                    Run the script. Please notice to insert <code>./</code> or <code>.\</code> before the script name. Gathers information about the local machine, displays the data in console, outputs the default two files to the default <code>-Output</code> location (<code>$env:temp</code>) and opens the created HTML-file in the default browser.</li>
                </p>
                <p>
                    <li><code>help ./Get-ComputerInfo -Full</code><br />
                    Display the help file.</li>
                </p>
                <p>
                    <li><code>./Get-ComputerInfo -Computer dc01, dc02 -Output "E:\chiore" <code>-SystemInfo</code> -Extract -MsInfo32 -Vbs -Cmdlet</code><br />
                    Run the script and get all the available computer related information from the computers <code>dc01</code> and <code>dc02</code>. Save most of the results in the "<code>E:\chiore</code>" directory (the results of the <code>GatherNetworkInfo.vbs</code> are saved to <code>$env:temp\Config</code> folder, if the command launching Get-ComputerInfo was run in an elevated PowerShell window). This command will work, because <code>-Vbs</code> is an alias of <code>-GatherNetworkInfo</code>. Since the path name doesn't contain any space characters, it doesn't need to be enveloped with quotation marks, and furthermore, the word <code>-Computer</code> may be left out from this command, too, because the values <code>dc01</code> and <code>dc02</code> are accepted as computer names due to their position (first).</li>
                </p>
                <p>
                    <li><p><code>Set-ExecutionPolicy remotesigned</code><br />
                    This command is altering the Windows PowerShell rights to enable script execution for the default (LocalMachine) scope. Windows PowerShell has to be run with elevated rights (run as an administrator) to actually be able to change the script execution properties. The default value of the default (LocalMachine) scope is "<code>Set-ExecutionPolicy restricted</code>".</p>
                        <p>Parameters:
                                <ol>
                                    <table>
                                        <tr>
                                            <td style="padding:6px"><code>Restricted</code></td>
                                            <td style="padding:6px">Does not load configuration files or run scripts. Restricted is the default execution policy.</td>
                                        </tr>
                                        <tr>
                                            <td style="padding:6px"><code>AllSigned</code></td>
                                            <td style="padding:6px">Requires that all scripts and configuration files be signed by a trusted publisher, including scripts that you write on the local computer.</td>
                                        </tr>
                                        <tr>
                                            <td style="padding:6px"><code>RemoteSigned</code></td>
                                            <td style="padding:6px">Requires that all scripts and configuration files downloaded from the Internet be signed by a trusted publisher.</td>
                                        </tr>
                                        <tr>
                                            <td style="padding:6px"><code>Unrestricted</code></td>
                                            <td style="padding:6px">Loads all configuration files and runs all scripts. If you run an unsigned script that was downloaded from the Internet, you are prompted for permission before it runs.</td>
                                        </tr>
                                        <tr>
                                            <td style="padding:6px"><code>Bypass</code></td>
                                            <td style="padding:6px">Nothing is blocked and there are no warnings or prompts.</td>
                                        </tr>
                                        <tr>
                                            <td style="padding:6px"><code>Undefined</code></td>
                                            <td style="padding:6px">Removes the currently assigned execution policy from the current scope. This parameter will not remove an execution policy that is set in a Group Policy scope.</td>
                                        </tr>
                                    </table>
                                </ol>
                        </p>
                    <p>For more information, please type "<code>Get-ExecutionPolicy -List</code>", "<code>help Set-ExecutionPolicy -Full</code>", "<code>help about_Execution_Policies</code>" or visit <a href="https://technet.microsoft.com/en-us/library/hh849812.aspx">Set-ExecutionPolicy</a> or <a href="http://go.microsoft.com/fwlink/?LinkID=135170">about_Execution_Policies</a>.</p>
                    </li>
                </p>
                <p>
                    <li><code>New-Item -ItemType File -Path C:\Temp\Get-ComputerInfo.ps1</code><br />
                    Creates an empty ps1-file to the <code>C:\Temp</code> directory. The <code>New-Item</code> cmdlet has an inherent <code>-NoClobber</code> mode built into it, so that the procedure will halt, if overwriting (replacing the contents) of an existing file is about to happen. Overwriting a file with the <code>New-Item</code> cmdlet requires using the <code>Force</code>. If the path name and/or the filename includes space characters, please enclose the whole <code>-Path</code> parameter value in quotation marks (single or double):
                        <ol>
                            <br /><code>New-Item -ItemType File -Path "C:\Folder Name\Get-ComputerInfo.ps1"</code>
                        </ol>
                    <br />For more information, please type "<code>help New-Item -Full</code>".</li>
                </p>
            </ol>
        </td>
    </tr>
</table>




### Contributing

<p>Find a bug? Have a feature request? Here is how you can contribute to this project:</p>

 <table>
   <tr>
      <th><img class="emoji" title="contributing" alt="contributing" height="28" width="28" align="absmiddle" src="https://assets-cdn.github.com/images/icons/emoji/unicode/1f33f.png"></th>
      <td style="padding:6px"><strong>Bugs:</strong></td>
      <td style="padding:6px"><a href="https://github.com/auberginehill/get-computer-info/issues">Submit bugs</a> and help us verify fixes.</td>
   </tr>
   <tr>
      <th rowspan="2"></th>
      <td style="padding:6px"><strong>Feature Requests:</strong></td>
      <td style="padding:6px">Feature request can be submitted by <a href="https://github.com/auberginehill/get-computer-info/issues">creating an Issue</a>.</td>
   </tr>
   <tr>
      <td style="padding:6px"><strong>Edit Source Files:</strong></td>
      <td style="padding:6px"><a href="https://github.com/auberginehill/get-computer-info/pulls">Submit pull requests</a> for bug fixes and features and discuss existing proposals.</td>
   </tr>
 </table>




### www

<table>
    <tr>
        <th><img class="emoji" title="www" alt="www" height="28" width="28" align="absmiddle" src="https://assets-cdn.github.com/images/icons/emoji/unicode/1f310.png"></th>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-computer-info">Script Homepage</a></td>
    </tr>
    <tr>
        <th rowspan="34"></th>
        <td style="padding:6px">clayman2: <a href="http://powershell.com/cs/media/p/7476.aspx">Disk Space</a> (or one of the <a href="http://web.archive.org/web/20120304222258/http://powershell.com/cs/media/p/7476.aspx">archive.org versions</a>)</td>
    </tr>
    <tr>
        <td style="padding:6px">Jeff Hicks: <a href="https://www.petri.com/validating-computer-lists-with-powershell">Validating Computer Lists with PowerShell</a></td>
    </tr>
    <tr>
        <td style="padding:6px">Paul-De: <a href="https://answers.microsoft.com/en-us/windows/forum/windows_7-security/does-anyone-know-what-gathernetworkinfovbs-is-its/63a302a6-cf69-4b9a-a3ef-4b2aff1b2514">Does anyone know what gatherNetworkInfo.vbs is?</a></td>
    </tr>
    <tr>
        <td style="padding:6px">alejandro5042: <a href="http://stackoverflow.com/questions/29266622/how-to-run-exe-with-without-elevated-privileges-from-powershell?rq=1">How to run exe with/without elevated privileges from PowerShell</a></td>
    </tr>
    <tr>
        <td style="padding:6px">Jeff: <a href="http://stackoverflow.com/questions/10941756/powershell-show-elapsed-time">Powershell show elapsed time</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://learningpcs.blogspot.com/2011/10/powershell-get-wmiobject-and.html">Powershell - Get-WmiObject and ASSOCIATORS OF Statement</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://4sysops.com/archives/windows-server-2012-server-core-part-5-tools/">Windows Server 2012 Server Core - Part 5: Tools</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://social.technet.microsoft.com/Forums/windowsserver/en-US/f82e6f0b-ab97-424b-8e91-508d710e03b1/how-to-link-the-output-from-win32diskdrive-and-win32volume?forum=winserverpowershell">How to link the output from win32_diskdrive and win32_volume</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://support.microsoft.com/en-us/help/300887/how-to-use-system-information-msinfo32-command-line-tool-switches">How to use System Information (msinfo32) command-line tool switches</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://technet.microsoft.com/en-us/library/ff730960.aspx">Windows PowerShell Tip of the Week: More Fun with Dates (and Times)</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://technet.microsoft.com/en-us/library/bb491007.aspx">Systeminfo</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://technet.microsoft.com/en-us/library/bb490937.aspx">Msinfo32</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/library/aa394102(v=vs.85).aspx">Win32_ComputerSystem class</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/library/aa394239(v=vs.85).aspx">Win32_OperatingSystem class</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/library/aa394474(v=vs.85).aspx">Win32_SystemEnclosure class</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/library/aa394512(v=vs.85).aspx">Win32_VideoController class</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/library/aa394360(v=vs.85).aspx">Win32_POTSModem class</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/library/aa394216(v=vs.85).aspx">Win32_NetworkAdapter class</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://technet.microsoft.com/en-us/library/ff920171(v=ws.11).aspx">Cscript</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/library/system.diagnostics.stopwatch(v=vs.110).aspx">Stopwatch Class</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/powershell/reference/5.1/microsoft.powershell.core/Where-Object">Where-Object</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.core/start-job">Start-Job</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/powershell/reference/5.1/Microsoft.PowerShell.Core/about/about_Jobs">About Jobs</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.management/get-computerinfo">Get-ComputerInfo</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://blogs.technet.microsoft.com/jonjor/2009/01/09/winrm-windows-remote-management-troubleshooting/">WinRM (Windows Remote Management) Troubleshooting</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://blogs.technet.microsoft.com/otto/2007/02/09/a-few-good-vista-ws-man-winrm-commands/">A Few Good Vista WS-Man (WinRM) Commands</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://blogs.technet.microsoft.com/askperf/2010/09/24/an-introduction-to-winrm-basics/">An Introduction to WinRM Basics</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://stackoverflow.com/questions/28481811/how-to-correctly-check-if-a-process-is-running-and-stop-it">How to Correctly Check if a Process is running and Stop it</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://powershellcookbook.com/recipe/qAxK/appendix-b-regular-expression-reference">Appendix B. Regular Expression Reference</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://www.verboon.info/2011/06/the-gathernetworkinfo-vbs-script/">The GatherNetworkInfo.vbs script</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/PowerShell/PowerShell/issues/3080">Get-ComputerInfo returns empty values on Windows 10 for most of the properties</a></td>
    </tr>   
    <tr>
        <td style="padding:6px"><a href="https://technet.microsoft.com/en-us/library/ee692804.aspx">The String’s the Thing</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://stackoverflow.com/questions/27175137/powershellv2-remove-last-x-characters-from-a-string#32608908">Powershellv2 - remove last x characters from a string</a></td>
    </tr>    
    <tr>
        <td style="padding:6px">ASCII Art: <a href="http://www.figlet.org/">http://www.figlet.org/</a> and <a href="http://www.network-science.de/ascii/">ASCII Art Text Generator</a></td>
    </tr>
</table>




### Related scripts

 <table>
    <tr>
        <th><img class="emoji" title="www" alt="www" height="28" width="28" align="absmiddle" src="https://assets-cdn.github.com/images/icons/emoji/unicode/0023-20e3.png"></th>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/aa812bfa79fa19fbd880b97bdc22e2c1">Disable-Defrag</a></td>
    </tr>
    <tr>
        <th rowspan="25"></th>
        <td style="padding:6px"><a href="https://github.com/auberginehill/firefox-customization-files">Firefox Customization Files</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-ascii-table">Get-AsciiTable</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-battery-info">Get-BatteryInfo</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-culture-tables">Get-CultureTables</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-directory-size">Get-DirectorySize</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-hash-value">Get-HashValue</a></td>
    </tr>    
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-installed-programs">Get-InstalledPrograms</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-installed-windows-updates">Get-InstalledWindowsUpdates</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-powershell-aliases-table">Get-PowerShellAliasesTable</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/9c2f26146a0c9d3d1f30ef0395b6e6f5">Get-PowerShellSpecialFolders</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-ram-info">Get-RAMInfo</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/eb07d0c781c09ea868123bf519374ee8">Get-TimeDifference</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-time-zone-table">Get-TimeZoneTable</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-unused-drive-letters">Get-UnusedDriveLetters</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/emoji-table">Emoji Table</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/java-update">Java-Update</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/remove-duplicate-files">Remove-DuplicateFiles</a></td>
    </tr>    
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/remove-empty-folders">Remove-EmptyFolders</a></td>
    </tr>    
    <tr>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/13bb9f56dc0882bf5e85a8f88ccd4610">Remove-EmptyFoldersLite</a></td>
    </tr> 
    <tr>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/176774de38ebb3234b633c5fbc6f9e41">Rename-Files</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/rock-paper-scissors">Rock-Paper-Scissors</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/toss-a-coin">Toss-a-Coin</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/unzip-silently">Unzip-Silently</a></td>
    </tr>    
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/update-adobe-flash-player">Update-AdobeFlashPlayer</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/update-mozilla-firefox">Update-MozillaFirefox</a></td>
    </tr>
</table>
