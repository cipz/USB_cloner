'Sauvegarde automatique des clés USB et SDCARD dés leurs insertion.
'Ce Programme sert à copier automatiquement chaque clé USB nouvellement insérée ou bien une SDCard.
'Il sert à faire des Sauvegardes incrémentielles de vos clés USB.
'Pour chaque clé USB, il crée un dossier de cette forme "NomMachine_NomVolumeUSB_NumSerie" dans le dossier %AppData% et
'il fait une copie totale pour la première fois, puis incrémentielle , càd ,il copie juste les nouveaux fichiers et les fichiers modifiés.
'Crée le 23/09/2014 © Hackoo
Option Explicit
Do
   Call AutoSave_USB_SDCARD()
   Pause(30)
Loop
'********************************************AutoSave_USB_SDCARD()************************************************
Sub AutoSave_USB_SDCARD()
   Dim Ws,WshNetwork,NomMachine,AppData,strComputer,objWMIService,objDisk,colDisks
   Dim fso,Drive,NumSerie,volume,cible,Amovible,Dossier,chemin,Command,Result
   Set Ws = CreateObject("WScript.Shell")
   Set WshNetwork = CreateObject("WScript.Network")
   NomMachine = WshNetwork.ComputerName
   AppData= ws.ExpandEnvironmentStrings("%AppData%")
   cible = AppData & "\"
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
                  NumSerie=fso.Drives(Drive + "\").SerialNumber
                  Amovible=fso.Drives(Drive + "\")
                  Numserie=ABS(INT(Numserie))
                  volume=fso.Drives(Drive + "\").VolumeName
                  Dossier=NomMachine & "_" & volume &"_"& NumSerie
                  chemin=cible & Dossier
                  Command = "cmd /c Xcopy.exe " & Amovible &" "& chemin &" /I /D /Y /S /J /C"
                  Result = Ws.Run(Command,0,True)
               end if
            End If   
         Next
      End If   
   Next
End Sub
'***************************************Fin du AutoSave_USB_SDCARD()*********************************************
'****************************************************************************************************************
Sub Pause(Sec)
   Wscript.Sleep(Sec*1000)
End Sub 
'****************************************************************************************************************
