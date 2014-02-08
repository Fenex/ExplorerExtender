Explorer Extender
================

Windows Explorer's Shell Extension

<h3>Features</h3>
 - Group - Move multiple files and folders into a new folder
 - Break folder - Move all folder content up one level
 - Mover - Move multiple files and folders from different folder into one folder

<h3>Installation</h3>
 - Place the DLL into a empty folder of your choice (e.g. C:\ExplorerExtender\ExplorerExtender.dll)
 - Open Command Prompt as Administrator
 - Run the following command
    - <pre>"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" "C:\ExplorerExtender\ExplorerExtender.dll" /codebase</pre>
 - Restart Windows Explorer

<h3>Uninstallation</h3>
 - Open Command Prompt as Administrator
 - Run the following command
    - <pre>"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" "C:\ExplorerExtender\ExplorerExtender.dll" /unregister</pre>
 - Restart Windows Explorer
