'Automatic backup USB sticks and SD CARD dice their insertion.
'This program is used to automatically copy each newly inserted USB key or a SDCard.
'It is used to make incremental backups of your USB key.
'For each USB drive, it creates a folder for this form "ComputerName_VolumeUSB_SerialNumber" in the D:\USB_Backup folder and
'It is a total copy first and then incrementally, ie, it just copies the new files and changed files.
'Created on 23/09/2014 © Hackoo
'04/02/2016
'Edited and translated from french to english and changed the directory target to be copied to D:\USB_Backup
'And copies only the files with .doc .docx and .pdf extensions
Option Explicit
Do
   Call AutoSave_USB_SDCARD()
   Pause(30)
Loop
'********************************************AutoSave_USB_SDCARD()************************************************
Sub AutoSave_USB_SDCARD()
   Dim Ws,WshNetwork,ComputerName,strComputer,objWMIService,objDisk,colDisks
   Dim fso,Drive,SerialNumber,volume,Target,Amovible,Folder,Command1,Command2,Command3,Result1,Result2,Result3
   Set Ws = CreateObject("WScript.Shell")
   Set WshNetwork = CreateObject("WScript.Network")
   ComputerName = WshNetwork.ComputerName
   Target = "D:\USB_Backup"
   strComputer = "."
   Set objWMIService = GetObject("winmgmts:" _
   & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
   Set colDisks = objWMIService.ExecQuery _
   ("SELECT * FROM Win32_LogicalDisk")

   For Each objDisk in colDisks
      If objDisk.DriveType = 2 Then
         Set fso = CreateObject("Scripting.FileSystemObject")
         For Each Drive In fso.Drives
            If Drive.IsReady Then
               If Drive.DriveType = 1 Then
                  SerialNumber=fso.Drives(Drive + "\").SerialNumber
                  Amovible=fso.Drives(Drive + "\")
                  SerialNumber=ABS(INT(SerialNumber))
                  volume=fso.Drives(Drive + "\").VolumeName
                  Folder=ComputerName & "_" & volume &"_"& SerialNumber
                  Target=Target &"\"& Folder
                  Command1 = "cmd /c Xcopy.exe " & Amovible &"\*.pdf "& Target &" /I /D /Y /S /J /C"
                  Command2 = "cmd /c Xcopy.exe " & Amovible &"\*.doc "& Target &" /I /D /Y /S /J /C"
                  Command3 = "cmd /c Xcopy.exe " & Amovible &"\*.docx "& Target &" /I /D /Y /S /J /C"
                  Result1 = Ws.Run(Command1,0,True)
                  Result2 = Ws.Run(Command2,0,True)
                  Result3 = Ws.Run(Command3,0,True)
               end if
            End If   
         Next
      End If   
   Next
End Sub
'****************************************************************************************************************
Sub Pause(Sec)
   Wscript.Sleep(Sec*1000)
End Sub 
'****************************************************************************************************************
