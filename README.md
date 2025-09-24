# printer_clean
This script is a robust, time-tested administrative utility designed to aggressively clean up all unnecessary printer configurations from a Windows workstation, including Windows 11.
Written in VBScript for maximum backward compatibility and execution stability, the script utilizes Windows Management Instrumentation (WMI) for printer removal and leverages the modern, native Windows utility pnputil.exe to ensure all orphaned printer drivers are completely uninstalled.

It is ideal for environments where computers accumulate multiple old printer installations, leading to instability or slow logon times.

#Key Features
Printer Removal: Deletes all installed printers on the system.

PDF Exclusions: Excludes specific virtual printers (e.g., CutePDF Writer, Acrobat Distiller) to preserve necessary PDF creation tools.

Port Cleanup: Identifies and deletes unused TCP/IP Printer Ports that are no longer associated with a current printer.

Modern Driver Cleanup: Uses pnputil.exe to reliably enumerate and remove all non-present (orphaned) printer driver packages, which is more effective than older methods.

Reliability: The use of pnputil.exe eliminates the need for repeated driver cleanup passes, ensuring a thorough job on the first run.

#How to Run
Save the code as a .vbs file (e.g., remove_printers_win11.vbs).

Run the script from an Administrator Command Prompt or a management system:

#Bash

cscript remove_printers_win11.vbs
