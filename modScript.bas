Attribute VB_Name = "modScript"
'
' Create a Basic EXE Inno Script
'
Public Sub BasicEXE_Script(ByVal Path As String, ByVal AppName As String, ByVal AppVersion As String, ByVal DirName As String, ByVal GroupName As String, ByVal SourceName As String)
  On Error Resume Next
  
  Open Path For Output As #1
  
  Dim ShortName As String, SourceDir As String
  
  'ShortName = Mid(VBInstance.ActiveVBProject.BuildFileName, InStrRev(VBInstance.ActiveVBProject.BuildFileName, "\") + 1)
  'SourceDir = Mid(VBInstance.ActiveVBProject.BuildFileName, 1, InStrRev(VBInstance.ActiveVBProject.BuildFileName, "\") - 1)
  
  ShortName = Mid(SourceName, InStrRev(SourceName, "\") + 1)
  SourceDir = Mid(SourceName, 1, InStrRev(SourceName, "\") - 1)

  ' Output Header
  Print #1, "; "
  Print #1, "; Install Script for " & AppName
  Print #1, ";  [Basic EXE Deployment Template]"
  Print #1, "; "
  Print #1, "; Generated by the 'Inno Setup Toolbar for VB6'"
  Print #1, "; Written and Programmed by Brian Haase"
  Print #1, "; "
  Print #1, " "
  
  Print #1, "[Setup]"
  Print #1, "AppName=" & AppName
  Print #1, "AppVerName=" & AppVersion
  Print #1, "AppPublisher=InnoSetupAddin"
  Print #1, "DefaultDirName={pf}\" & DirName
  Print #1, "DefaultGroupName=" & GroupName
  Print #1, "SourceDir=" & SourceDir
  Print #1, "OutputDir=" & SourceDir & "\Output"
  Print #1, "DisableStartupPrompt = yes"
  Print #1, " "
  Print #1, "[Tasks]"
  Print #1, "Name: ""desktopicon""; Description: ""Create a &desktop icon""; GroupDescription: ""Additional icons:"""
  Print #1, " "
  Print #1, "[Files]"
  Print #1, "Source: """ & SourceName & """; DestDir: ""{app}""; Flags: ignoreversion"
  Print #1, "; NOTE: Don't use ""Flags: ignoreversion"" on any shared system files"
  Print #1, " "
  Print #1, "[icons]"
  Print #1, "Name: ""{group}\" & AppName & """; Filename: ""{app}\" & ShortName & """"
  Print #1, "Name: ""{group}\Uninstall My Program""; Filename: ""{uninstallexe}"""
  Print #1, "Name: ""{userdesktop}\" & AppName & """; Filename: ""{app}\" & ShortName & """; Tasks: desktopicon"
  Print #1, " "
  Print #1, "[Run]"
  Print #1, "Filename: ""{app}\" & ShortName & """; Description: ""Launch My Program""; Flags: nowait postinstall skipifsilent"
  
  Close #1
End Sub

'
' Create an Empty Inno Script
'
Public Sub Empty_Script(ByVal Path As String)
  Open Path For Output As #1
  
  Dim ShortName As String, SourceDir As String
  
  ShortName = Mid(VBInstance.ActiveVBProject.BuildFileName, InStrRev(VBInstance.ActiveVBProject.BuildFileName, "\") + 1)
  
  SourceDir = Mid(VBInstance.ActiveVBProject.BuildFileName, 1, InStrRev(VBInstance.ActiveVBProject.BuildFileName, "\") - 1)
  
  ' Output Header
  Print #1, "; "
  Print #1, "; Empty Script Template"
  Print #1, "; "
  Print #1, "; Generated by the 'Inno Setup Toolbar for VB6'"
  Print #1, "; Written and Programmed by Brian Haase"
  Print #1, "; "
  Print #1, " "
  
  Close #1
End Sub


'
' Create a DLL Inno Script
'
Public Sub DLL_Script(ByVal Path As String, ByVal AppName As String, ByVal AppVersion As String, ByVal DirName As String, ByVal GroupName As String, ByVal SourceName As String)
  On Error Resume Next
  
  Open Path For Output As #1
  
  Dim ShortName As String, SourceDir As String
  
  'ShortName = Mid(VBInstance.ActiveVBProject.BuildFileName, InStrRev(VBInstance.ActiveVBProject.BuildFileName, "\") + 1)
  'SourceDir = Mid(VBInstance.ActiveVBProject.BuildFileName, 1, InStrRev(VBInstance.ActiveVBProject.BuildFileName, "\") - 1)
  
  ShortName = Mid(SourceName, InStrRev(SourceName, "\") + 1)
  SourceDir = Mid(SourceName, 1, InStrRev(SourceName, "\") - 1)

  ' Output Header
  Print #1, "; "
  Print #1, "; Install Script for " & AppName
  Print #1, ";  [DLL Deployment Template]"
  Print #1, "; "
  Print #1, "; Generated by the 'Inno Setup Toolbar for VB6'"
  Print #1, "; Written and Programmed by Brian Haase"
  Print #1, "; "
  Print #1, " "
  
  Print #1, "[Setup]"
  Print #1, "AppName=" & AppName
  Print #1, "AppVerName=" & AppVersion
  Print #1, "AppPublisher=InnoSetupAddin"
  Print #1, "DefaultDirName={sys}"
  Print #1, "SourceDir=" & SourceDir
  Print #1, "OutputDir=" & SourceDir & "\Output"
  Print #1, "CreateUninstallRegKey = yes"
  Print #1, "DirExistsWarning = no"
  Print #1, "DisableStartupPrompt = yes"
  Print #1, " "
  Print #1, "[Files]"
  Print #1, "Source: """ & SourceName & """; DestDir: ""{app}""; Flags: onlyifdoesntexist regserver sharedfile"
  Print #1, "; NOTE: Don't use ""Flags: ignoreversion"" on any shared system files"
  
  Close #1
End Sub

'
' Create an OCX Inno Script
'
Public Sub OCX_Script(ByVal Path As String, ByVal AppName As String, ByVal AppVersion As String, ByVal DirName As String, ByVal GroupName As String, ByVal SourceName As String)
  On Error Resume Next
  
  Open Path For Output As #1
  
  Dim ShortName As String, SourceDir As String
  
  'ShortName = Mid(VBInstance.ActiveVBProject.BuildFileName, InStrRev(VBInstance.ActiveVBProject.BuildFileName, "\") + 1)
  'SourceDir = Mid(VBInstance.ActiveVBProject.BuildFileName, 1, InStrRev(VBInstance.ActiveVBProject.BuildFileName, "\") - 1)
  
  ShortName = Mid(SourceName, InStrRev(SourceName, "\") + 1)
  SourceDir = Mid(SourceName, 1, InStrRev(SourceName, "\") - 1)

  ' Output Header
  Print #1, "; "
  Print #1, "; Install Script for " & AppName
  Print #1, ";  [OCX Deployment Template]"
  Print #1, "; "
  Print #1, "; Generated by the 'Inno Setup Toolbar for VB6'"
  Print #1, "; Written and Programmed by Brian Haase"
  Print #1, "; "
  Print #1, " "
  
  Print #1, "[Setup]"
  Print #1, "AppName=" & AppName
  Print #1, "AppVerName=" & AppVersion
  Print #1, "AppPublisher=InnoSetupAddin"
  Print #1, "DefaultDirName={sys}"
  Print #1, "SourceDir=" & SourceDir
  Print #1, "OutputDir=" & SourceDir & "\Output"
  Print #1, "CreateUninstallRegKey = yes"
  Print #1, "DirExistsWarning = no"
  Print #1, "DisableStartupPrompt = yes"
  Print #1, " "
  Print #1, "[Files]"
  Print #1, "Source: """ & SourceName & """; DestDir: ""{app}""; Flags: onlyifdoesntexist regserver sharedfile"
  Print #1, "; NOTE: Don't use ""Flags: ignoreversion"" on any shared system files"
  
  Close #1
End Sub

'
' Create an TLB Inno Script
'
Public Sub TLB_Script(ByVal Path As String, ByVal AppName As String, ByVal AppVersion As String, ByVal DirName As String, ByVal GroupName As String, ByVal SourceName As String)
  On Error Resume Next
  
  Open Path For Output As #1
  
  Dim ShortName As String, SourceDir As String
  
  'ShortName = Mid(VBInstance.ActiveVBProject.BuildFileName, InStrRev(VBInstance.ActiveVBProject.BuildFileName, "\") + 1)
  'SourceDir = Mid(VBInstance.ActiveVBProject.BuildFileName, 1, InStrRev(VBInstance.ActiveVBProject.BuildFileName, "\") - 1)
  
  ShortName = Mid(SourceName, InStrRev(SourceName, "\") + 1)
  SourceDir = Mid(SourceName, 1, InStrRev(SourceName, "\") - 1)

  ' Output Header
  Print #1, "; "
  Print #1, "; Install Script for " & AppName
  Print #1, ";  [DLL Deployment Template]"
  Print #1, "; "
  Print #1, "; Generated by the 'Inno Setup Toolbar for VB6'"
  Print #1, "; Written and Programmed by Brian Haase"
  Print #1, "; "
  Print #1, " "
  
  Print #1, "[Setup]"
  Print #1, "AppName=" & AppName
  Print #1, "AppVerName=" & AppVersion
  Print #1, "AppPublisher=InnoSetupAddin"
  Print #1, "DefaultDirName={sys}"
  Print #1, "SourceDir=" & SourceDir
  Print #1, "OutputDir=" & SourceDir & "\Output"
  Print #1, "CreateUninstallRegKey = yes"
  Print #1, "DirExistsWarning = no"
  Print #1, "DisableStartupPrompt = yes"
  Print #1, " "
  Print #1, "[Files]"
  Print #1, "Source: """ & SourceName & """; DestDir: ""{app}""; Flags: onlyifdoesntexist regtypelib sharedfile"
  Print #1, "; NOTE: Don't use ""Flags: ignoreversion"" on any shared system files"
  
  Close #1
End Sub

