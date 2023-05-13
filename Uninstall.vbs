' Deletes settings & cache
Dim path : path = SDB.ApplicationPath&"Scripts\Auto\"
Dim i : i = InStrRev(SDB.Database.Path,"\")
Dim appPath : appPath = Left(SDB.Database.Path,i)
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")


iniSec = "ViewMode_MM"&Round(SDB.VersionHi)     'Put ini section name here
SDB.IniFile.DeleteSection(iniSec)

If fso.FileExists(path&"ViewMode.vbs") Then
	Call fso.DeleteFile(path&"ViewMode.vbs")
End If

MsgBox("I hope your experiences with View Mode were not all bad." & vbNewLine & "Please restart MediaMonkey for the uninstall to complete.")




