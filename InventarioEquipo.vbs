' Este script alimenta un archivo de texto que almacena todos los datos importantes para el respectivo inventario
' (Nombre de Equipo, Usuario Logueado, Numero de Serie, Modelo, Procesador, Capacidad de RAM y Capacidad de Disco Duro)
' Gemis Daniel Guevara Villeda
' 18 de Mayo del 2017
' 17:15 PM
'Option Explicit
On Error Resume Next

'Preparamos el numero de serie*******************************************
'Variables donde vamos a insertar la informacion
Dim strDirectory
Dim strFile
strDirectory = "\\10.10.23.15\Inventario\"    'Dirección donde se desea guardar la información
strFile = "Equipos.txt"                       'Nombre del archivo donde se guardará la información

Dim strComputer
Dim strSerial	' Para el numero de serie
strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSMBIOS = objWMIService.ExecQuery ("Select * from Win32_SystemEnclosure")

For Each objSMBIOS in colSMBIOS
	strSerial = objSMBIOS.SerialNumber
Next

'************************************************************************

'Preparamos el nombre de equipo y el usuario logueado********************
set objFSO = CreateObject("Scripting.FileSystemObject")
Set objNetwork = WScript.CreateObject("WScript.Network")
Set objSysInfo = Wscript.createObject("ADSystemInfo")

'GET ADSPATH & FULL NAMES
strHost=objNetwork.Computername
strDNUser=objSysInfo.username
Set objUser = GetObject("LDAP://" & strDNUser)
strUser = objUser.displayname
strDepto = objUser.department
strDescription = objUser.description
strMail = objUser.mail
strLogonName = objUser.sAMAccountName

'************************************************************************

'Preparamos el modelo de la maquina y Tamaño de Memoria RAM**************

Dim strModel ' Donde guardamos el modelo de la maquina
Dim strTotalRAM ' Tamaño de RAM en MB'S

Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")

For Each objComputer in colSettings
	strModel = objComputer.Model
	strTotalRAM = objComputer.TotalPhysicalMemory / 1024 / 1024 & " MB"
Next

'Preparamos el Nombre del Procesador*************************************
Dim strProcessorName

Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")

For Each objItem in colItems
	strProcessorName = objItem.Name
Next

'Capacidad del Disco Duro************************************************

Const HARD_DISK = 3
Dim strSizeHardDisk

Set colDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk Where DriveType = " & HARD_DISK & "")
For Each objDisk in colDisks
	strSizeHardDisk = objDisk.Size / 1024 / 1024 & " MB"
Next

'Version de Windows y Office***********************************************
Dim objShell
Dim strOSVersion
Dim strOfficeVersion

Set objShell = CreateObject("WScript.Shell")

strOSVersion = objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName")' Read the registry for the operating system version

strOfficeVersion = "Office " & GetOfficeVersionNumber() ' Read the office version from the function

    Function GetOfficeVersionNumber()
        GetOfficeVersionNumber = ""  ' or you could use "Office not installed"
        Dim sTempValue
                    ' Read the Classes Root registry hive (it is a memory-only instance amalgamation of HKCU\Software\Classes and HKLM\Software\Classes registry keys) as it contains a source of information for the currently active Microsoft Office Excel application major version - it's quicker and easier to read the registry than the file version information after a location lookup). The trailing backslash on the line denotes that the @ or default registry key value is being queried.
        sTempValue = objShell.RegRead("HKCR\Excel.Application\CurVer\")
        If Len(sTempValue) > 2 Then GetOfficeVersionNumber = Replace(Right(sTempValue, 2), ".", "") ' Check the length of the value found and if greater than 2 digits then read the last two digits for the major Office version value
    End Function    ' GetOfficeVersionNumber

'************************************************************************

'INSERT DE DATOS USER & HOSTNAME 

Set objTextFile = objFSO.OpenTextFile(strDirectory & strFile, 8,True)

objTextFile.Writeline(strSerial & ";" & strUser & ";" & strHost & ";" & strModel & ";" & strProcessorName & ";" & strTotalRAM & ";" & strSizeHardDisk & ";" & strOSVersion & ";" & strOfficeVersion & ";" & strDepto & ";" & strDescription & ";" & strMail)
objTextFile.close 