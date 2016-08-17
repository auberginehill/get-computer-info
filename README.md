<!-- Visual Studio Code: For a more comfortable reading experience, use the key combination Ctrl + Shift + V 

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
      <td style="padding:6px">Get-ComputerInfo uses Windows Management Instrumentation (WMI) to retrieve basic computer information, a list of volumes and the partition table and displays the results on-screen and writes them to two files. The HTML source code is also displayed in a pop-up window. This script is based on clayman2's PowerShell script "<a href="http://powershell.com/cs/media/p/7476.aspx">Disk Space</a>".</td>
   </tr>
   <tr>
      <td style="padding:6px"><strong>Homepage:</strong></td>
      <td style="padding:6px"><a href="https://github.com/auberginehill/get-computer-info">https://github.com/auberginehill/get-computer-info</a></td>
   </tr>
   <tr>
      <td style="padding:6px"><strong>Version:</strong></td>
      <td style="padding:6px">1.0</td>
   </tr>
   <tr>
        <td style="padding:6px"><strong>Sources:</strong></td>
        <td style="padding:6px">
            <table>
                <tr>
                    <td style="padding:6px">Emojis:</td>
                    <td style="padding:6px"><a href="https://api.github.com/emojis">https://api.github.com/emojis</a></td>
                </tr>                    
                <tr>
                    <td style="padding:6px">clayman2:</td>                
                    <td style="padding:6px"><a href="http://powershell.com/cs/media/p/7476.aspx">Disk Space</a></td>  
                </tr>
            </table>
        </td>
   </tr> 
   <tr>
      <td style="padding:6px"><strong>Downloads:</strong></td>
      <td style="padding:6px">For instance <a href="https://raw.githubusercontent.com/auberginehill/get-computer-info/master/Get-ComputerInfo.ps1">Get-ComputerInfo.ps1</a>. Or <a href="https://github.com/auberginehill/get-computer-info/archive/master.zip">everything as a .zip-file</a>.</td>
   </tr> 
</table>




#### Outputs

<table>
    <tr>
        <th>:arrow_right:</th>
        <td style="padding:6px">
            <ul>
                <li>Displays general computer information (such as Computer Name, Manufacturer, Computer Model, System Type, CPU, Operating System, Architecture, SP Version, Build Number, Memory, Processors, Cores, Country Code, Install Date, Last BootUp, UpTime, Date, Daylight Bias, Time Offset (Current), Time Offset (Normal), Time (Current), Time (Normal), Daylight In Effect, Time Zone, OS Version, BIOS Version, Serial Number (BIOS), Serial Number (Base Board), Serial Number (OS), UUID), and a list of volumes in console. In addition to that... </li>
            </ul>
        </td>
    </tr>
    <tr>
        <th></th>
        <td style="padding:6px">
            <ul>
                <p>
                    <li>One pop-up window (Out-GridView):</li>
                </p>
                <ol>
                    <p>
                        <table>
                            <tr>
                                <td style="padding:6px"><strong>Name</strong></td>
                                <td style="padding:6px"><strong>Description</strong></td>
                            </tr>
                            <tr>
                                <td style="padding:6px">$path\computer_info.html</a></td>
                                <td style="padding:6px">Displays the source code of the HTML-file</td>
                            </tr>
                        </table>
                    </p>
                </ol>
                <p>
                    <li>And also one HTML-file and one CSV-file at $path.</li>
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
            </ul>
        </td>
    </tr>
</table>




#### Notes

<table>
    <tr>
        <th>:warning:</th>
        <td style="padding:6px">
            <ul>
                <li>Please note that the two files are created in a directory, which is specified with the <code>$path</code> variable (at line 10).</li>
            </ul>
        </td>
    </tr>
    <tr>
        <th></th>
        <td style="padding:6px">
            <ul>
                <p>
                    <li>The <code>$env:temp</code> variable points to the current temp folder. The default value of the <code>$env:temp</code> variable is <code>C:\Users\&lt;username&gt;\AppData\Local\Temp</code> (i.e. each user account has their own separate temp folder at path <code>%USERPROFILE%\AppData\Local\Temp</code>). To change the temp folder for instance to <code>C:\Temp</code>, please, for example, follow the instructions at <a href="http://www.eightforums.com/tutorials/23500-temporary-files-folder-change-location-windows.html">Temporary Files Folder - Change Location in Windows</a>, which in essence are something along the lines:
                        <ol>
                           <li>Right click on Computer and click on Properties. In the resulting window with the basic information about your computer...</li>
                           <li>Click on Advanced system settings on the left panel and select Advanced tab on the resulting pop-up window.</li>
                           <li>Click on the button near the bottom labeled Environment Variables.</li>
                           <li>In the topmost section labeled User variables you may see both TMP and TEMP listed. Each different login account is assigned its own temporary locations. These values can be changed by double clicking a value or by highlighting a value and selecting Edit. The specified path will be used by Windows and many other programs for temporary files. It's advisable to set the same value (a directory path) for both TMP and TEMP.</li>
                           <li>You'll need to restart any running programs for the new value to take effect. In fact, you'll probably also need to restart Windows for it to begin using the new value for its own temporary files.</li>
                        </ol>
                    </li>
                </p>
            </ul>
        </td>
    </tr>
</table>




#### Examples

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
                    Run the script. Please notice to insert <code>./</code> or <code>.\</code> before the script name.</li>
                </p>
                <p>
                    <li><code>help ./Get-ComputerInfo -Full</code><br />
                    Display the help file.</li>
                </p>
                <p>
                    <li><p><code>Set-ExecutionPolicy remotesigned</code><br />
                    This command is altering the Windows PowerShell rights to enable script execution. Windows PowerShell has to be run with elevated rights (run as an administrator) to actually be able to change the script execution properties. The default value is "<code>Set-ExecutionPolicy restricted</code>".</p>
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
                    <p>For more information, type "<code>help Set-ExecutionPolicy -Full</code>" or visit <a href="https://technet.microsoft.com/en-us/library/hh849812.aspx">Set-ExecutionPolicy</a>.</p>
                    </li>
                </p>
                <p>
                    <li><code>New-Item -ItemType File -Path C:\Temp\Get-ComputerInfo.ps1</code><br />
                    Creates an empty ps1-file to the <code>C:\Temp</code> directory. The <code>New-Item</code> cmdlet has an inherent <code>-NoClobber</code> mode built into it, so that the procedure will halt, if overwriting (replacing the contents) of an existing file is about to happen. Overwriting a file with the <code>New-Item</code> cmdlet requires using the <code>Force</code>.<br /> 
                    For more information, please type "<code>help New-Item -Full</code>".</li>
                </p>
            </ol>
        </td>
    </tr>
</table>




#### Contributing

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




#### www

<table>
    <tr>
        <th><img class="emoji" title="www" alt="www" height="28" width="28" align="absmiddle" src="https://assets-cdn.github.com/images/icons/emoji/unicode/1f310.png"></th>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-computer-info">Script Homepage</a></td>
    </tr>
    <tr>
        <th rowspan="8"></th>
        <td style="padding:6px">clayman2: <a href="http://powershell.com/cs/media/p/7476.aspx">Disk Space</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://learningpcs.blogspot.com/2011/10/powershell-get-wmiobject-and.html">Powershell - Get-WmiObject and ASSOCIATORS OF Statement</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://social.technet.microsoft.com/Forums/windowsserver/en-US/f82e6f0b-ab97-424b-8e91-508d710e03b1/how-to-link-the-output-from-win32diskdrive-and-win32volume?forum=winserverpowershell">How to link the output from win32_diskdrive and win32_volume</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://technet.microsoft.com/en-us/library/ff730960.aspx">Windows PowerShell Tip of the Week: More Fun with Dates (and Times)</a></td>
    </tr>
    <tr>
        <td style="padding:6px">ASCII Art: <a href="http://www.figlet.org/">http://www.figlet.org/</a> and <a href="http://www.network-science.de/ascii/">ASCII Art Text Generator</a></td>
    </tr>
</table>




#### Related scripts

 <table>
    <tr>
        <th><img class="emoji" title="www" alt="www" height="28" width="28" align="absmiddle" src="https://assets-cdn.github.com/images/icons/emoji/unicode/0023-20e3.png"></th>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-ram-info">Get-RAMInfo</a></td>
    </tr>
    <tr>
        <th rowspan="5"></th>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/eb07d0c781c09ea868123bf519374ee8">Get-TimeDifference</a></td>        
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-unused-drive-letters">Get-UnusedDriveLetters</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/list-windows-updates">List-WindowsUpdates</a></td>
    </tr>
    <tr>
        <td style="padding:6px"></td>
    </tr>
    <tr>
        <td style="padding:6px"></td>
    </tr>
</table>
