'- - - - - - - - - - - - - - - - - - -
'First delete all Printers from the computer
'(in my case i excluded two printers, used for pdf creation on
' several workstations)
'Then delete the TCPIP-printer-ports that are not used anymore by a printer
'Finaly, un-install the drivers that are not currently used by a printer
' cscript remove_printers_win11.vbs...run in cmd as admin

on Error Resume Next
Set objDictionary = CreateObject("Scripting.Dictionary")

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer")
'***delete all Printers from the computer
    For Each objPrinter in colPrinters
    If Not objPrinter.Name="CutePDF Writer" Then
    If Not objPrinter.Name="Acrobat Distiller" Then
        ObjPrinter.Delete_
    End If
    End If
Next

'***delete the TCPIP-printer-ports
For Each objPrinter in colPrinters 
    objDictionary.Add objPrinter.PortName, objPrinter.PortName
Next

Set colPorts = objWMIService.ExecQuery _
    ("Select * from Win32_TCPIPPrinterPort")
For Each objPort in colPorts
    If objDictionary.Exists(objPort.Name) Then
    Else
        ObjPort.Delete_
    End If
Next

'***un-install the printer drivers using pnputil
Set WSHShell = CreateObject("WScript.Shell")

'The following command uninstalls all non-present (orphaned) printer drivers
WSHShell.Run "pnputil.exe /enum-drivers | findstr /i 'PublishedName: oem' > temp.txt", 0, True
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("temp.txt", 1)
Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
    'Extract the driver's published name (e.g., oemXX.inf)
    arrLine = Split(strLine, ": ")
    strDriverName = Trim(arrLine(1))
    'Now, use pnputil to delete the driver package
    WSHShell.Run "pnputil.exe /delete-driver " & strDriverName & " /uninstall", 0, True
Loop
objFile.Close
objFSO.DeleteFile "temp.txt", True

'***a pause between the first run and second (this is optional but can be helpful)
WSHShell.run "ping -n 20 127.0.0.1 >NUL"

'***You can remove the second run as pnputil is more reliable at the first pass
'WSHShell.Run "pnputil.exe /delete-driver oemXX.inf /uninstall", 0, True

'- - - - - - - - - - - - - - - - - - -