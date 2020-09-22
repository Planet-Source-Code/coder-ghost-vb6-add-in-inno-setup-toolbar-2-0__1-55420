VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9960
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   13395
   _ExtentX        =   23627
   _ExtentY        =   17568
   _Version        =   393216
   Description     =   $"Connect.dsx":0000
   DisplayName     =   "Inno Setup Toolbar 2"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'This Command Bar will act as a button container.
Private cmdBar                        As CommandBar

'The Declares for the buttons themselves.
Private cmdBarBtn1                    As CommandBarButton
Private cmdBarBtn2                    As CommandBarButton
Private cmdBarBtn3                    As CommandBarButton
Private cmdBarBtn4                    As CommandBarButton

'Allow the buttons to recieve events.
Private WithEvents cmdBarBtnEvents1   As CommandBarEvents
Attribute cmdBarBtnEvents1.VB_VarHelpID = -1
Private WithEvents cmdBarBtnEvents2   As CommandBarEvents
Attribute cmdBarBtnEvents2.VB_VarHelpID = -1
Private WithEvents cmdBarBtnEvents3   As CommandBarEvents
Attribute cmdBarBtnEvents3.VB_VarHelpID = -1
Private WithEvents cmdBarBtnEvents4   As CommandBarEvents
Attribute cmdBarBtnEvents4.VB_VarHelpID = -1

Dim tmpConfig As frmConfig
Dim tmpDefault As frmDefault
Dim tmpWizard As frmWizard

'
' This method adds the Add-In to VB.
'
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    'On Error GoTo Err_Handler:
    Dim RtnA As String, RtnB As String, RtnC As String, RtnD As String
    Dim CtlType As MsoControlType
    
    'Save the VB Instance
    Set VBInstance = Application
       
    'Locate the Inno Executable, if its installed.
    InnoEXE = Find_Inno
 
    'Create the toolbar
    'Set cmdBar = VBInstance.CommandBars.Add("Inno Setup Toolbar", msoBarFloating)
    'Set cmdBar = VBInstance.CommandBars.Add("Inno Setup", msoBarTop)
    Set cmdBar = VBInstance.CommandBars.add("Inno Setup Toolbat", msoBarTop)
   
    RtnA = Get_Key("Position")
    RtnB = Get_Key("RowIndex")
    RtnC = Get_Key("Left")
    RtnD = Get_Key("Top")
    
    If RtnA <> "" Then cmdBar.Position = CLng(RtnA)
    If RtnB <> "" Then cmdBar.RowIndex = CLng(RtnB)
    If RtnC <> "" Then cmdBar.left = CLng(RtnC)
    If RtnD <> "" Then cmdBar.top = CLng(RtnD)
       
    'Make it visible
    cmdBar.Visible = True
    
    'Set the control type were adding to the command bar
    CtlType = msoControlButton
    
    'We now need to add the buttons to the toolbar
    Set cmdBarBtn1 = cmdBar.Controls.add(CtlType)
    Set cmdBarBtn2 = cmdBar.Controls.add(CtlType)
    Set cmdBarBtn3 = cmdBar.Controls.add(CtlType)
    Set cmdBarBtn4 = cmdBar.Controls.add(CtlType)
        
    'Create the properties for the toolbar.
    With cmdBarBtn1
        .Caption = "Script Editor"
        .ToolTipText = "Launch the Inno Script Editor"
        .Style = msoButtonIcon
        .FaceId = 593
    End With
    
    '625 for Wizard
    With cmdBarBtn2
        .Caption = "Script Wizard"
        .ToolTipText = "Script Wizard"
        .Style = msoButtonIcon
        .FaceId = 625
    End With
    
    With cmdBarBtn3
        .Caption = "Compile Script"
        .ToolTipText = "Compile the Script"
        .Style = msoButtonIcon
        .FaceId = 1396
    End With
    
    With cmdBarBtn4
        .Caption = "Configuration"
        .ToolTipText = "Configure the Addin"
        .Style = msoButtonIcon
        .FaceId = 2946
    End With
    
   '-------------------------------------------
   ' we now need to link the buttons to events
   '-------------------------------------------
   With VBInstance
        Set cmdBarBtnEvents1 = .Events.CommandBarEvents(cmdBarBtn1)
        Set cmdBarBtnEvents2 = .Events.CommandBarEvents(cmdBarBtn2)
        Set cmdBarBtnEvents3 = .Events.CommandBarEvents(cmdBarBtn3)
        Set cmdBarBtnEvents4 = .Events.CommandBarEvents(cmdBarBtn4)
   End With
 
 Exit Sub
Err_Handler:
     Err.Source = Err.Source & "." & VarType(Me) & ".AddinInstance_OnConnection"
     Debug.Print Err.Number & vbTab & Err.Source & Err.Description
     Err.Clear
     Resume Next
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'this is the time to destroy objects
'and references
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
 On Error GoTo Err_Handler:
   
     Call Write_Key("Position", CStr(cmdBar.Position))
     Call Write_Key("RowIndex", CStr(cmdBar.RowIndex))
     Call Write_Key("Left", CStr(cmdBar.left))
     Call Write_Key("Top", CStr(cmdBar.top))
     
     If tmpConfig Is Nothing Then
        'DO Nothing
     Else
        Unload tmpConfig
     End If
     
     If tmpDefault Is Nothing Then
        'DO Nothing
     Else
        Unload tmpDefault
     End If
     
     If tmpWizard Is Nothing Then
        'DO Nothing
     Else
        Unload tmpWizard
     End If
     
     'delete the buttons
     cmdBarBtn1.Delete
     cmdBarBtn2.Delete
     cmdBarBtn3.Delete
     cmdBarBtn4.Delete
     
     'unset buttons reference
     Set cmdBarBtn1 = Nothing
     Set cmdBarBtn2 = Nothing
     Set cmdBarBtn3 = Nothing
     Set cmdBarBtn4 = Nothing
     
     'unset events reference
     Set cmdBarBtnEvents1 = Nothing
     Set cmdBarBtnEvents2 = Nothing
     Set cmdBarBtnEvents3 = Nothing
     Set cmdBarBtnEvents4 = Nothing
     
     'destroy toolbar and its  variable
     cmdBar.Delete
     Set cmdBar = Nothing

     'kill core reference
     Set VBInstance = Nothing
     
 Exit Sub
Err_Handler:
     Err.Source = Err.Source & "." & VarType(Me) & ".AddinInstance_OnDisconnection"
     Debug.Print Err.Number & vbTab & Err.Source & Err.Description
     Err.Clear
     Resume Next
End Sub

'
' Edit Script Button Event
'
Private Sub cmdBarBtnEvents1_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  'Edit Script
  '------------------------------------------------
  Dim ScriptPath As String, Rtn As Integer
  
  If VBInstance.ActiveVBProject Is Nothing Then
     MsgBox "You must load and select a project before continuing.", vbCritical + vbOKOnly, "VB6 - Inno Setup Toolbar"
     Exit Sub
  End If
  
  ScriptPath = VBInstance.ActiveVBProject.FileName
  If Len(ScriptPath) > 3 Then ScriptPath = Mid(ScriptPath, 1, Len(ScriptPath) - 3) & "iss"
  
  'Autodetect Inno
  If InnoEXE = "" Then
     InnoEXE = Find_Inno
     If InnoEXE = "" Then
        MsgBox "Unable to locate the Inno Compiler Program. Please configure the toolbar.", vbCritical + vbOKOnly, "VB6 - Inno Setup Toolbar"
        Call ShowConfig
        Exit Sub
     End If
  End If
  
  If ScriptPath = "" Then
     MsgBox "This project has not been saved yet.", vbCritical + vbOKOnly, "VB6 - Inno Setup Toolbar"
     Exit Sub
  End If
  
  If VBInstance.ActiveVBProject.IsDirty Then
     Rtn = MsgBox("This project has been changed since your last save. Continue Anyway?", vbExclamation + vbYesNo, "VB6 - Inno Setup Toolbar")
     If Rtn <> vbYes Then Exit Sub
  End If
  
  If Not File_Exists(ScriptPath) Then
     'Create a default script
     Call ShowDefault(ScriptPath)
  Else
    'Open the Script
     Call API_WinExec(Chr(34) & InnoEXE & Chr(34) & " " & Chr(34) & ScriptPath & Chr(34), False)
  End If
End Sub

'
' Script Wizard Button Event
'
Private Sub cmdBarBtnEvents2_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  'Script Wizard
  '------------------------------------------------
  Dim ScriptPath As String, Rtn As Integer
  
  If VBInstance.ActiveVBProject Is Nothing Then
     MsgBox "You must load and select a project before continuing.", vbCritical + vbOKOnly, "VB6 - Inno Setup Toolbar"
     Exit Sub
  End If
  
  ScriptPath = VBInstance.ActiveVBProject.FileName
  If Len(ScriptPath) > 3 Then ScriptPath = Mid(ScriptPath, 1, Len(ScriptPath) - 3) & "iss"
  
  'Autodetect Inno
  If InnoEXE = "" Then
     InnoEXE = Find_Inno
     If InnoEXE = "" Then
        MsgBox "Unable to locate the Inno Compiler Program. Please configure the toolbar.", vbCritical + vbOKOnly, "VB6 - Inno Setup Toolbar"
        Call ShowConfig
        Exit Sub
     End If
  End If
  
  If ScriptPath = "" Then
     MsgBox "This project has not been saved yet.", vbCritical + vbOKOnly, "VB6 - Inno Setup Toolbar"
     Exit Sub
  End If
  
  If VBInstance.ActiveVBProject.IsDirty Then
     Rtn = MsgBox("This project has been changed since your last save. Continue Anyway?", vbExclamation + vbYesNo, "VB6 - Inno Setup Toolbar")
     If Rtn <> vbYes Then Exit Sub
  End If
  
  If File_Exists(ScriptPath) Then
     'Warn about Overwrite
     Rtn = MsgBox("There is already an Inno Setup Script for this file in your project directory." & Chr(10) & "The Script Wizard will overwrite the script. Are you sure you want to continue?", vbExclamation + vbYesNo, "VB6 - Inno Setup Toolbar")
     If Rtn = vbNo Then Exit Sub
  End If
  
  'Create a default script
  Call ShowWizard(ScriptPath)
End Sub

'
' Compile Script Button Event
'
Private Sub cmdBarBtnEvents3_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  'Compile Script
  '------------------------------------------------
  Dim ScriptPath As String, Rtn As Integer
  
  If VBInstance.ActiveVBProject Is Nothing Then
     MsgBox "You must load and select a project before continuing.", vbCritical + vbOKOnly, "VB6 - Inno Setup Toolbar"
     Exit Sub
  End If
  
  ScriptPath = VBInstance.ActiveVBProject.FileName
  If Len(ScriptPath) > 3 Then ScriptPath = Mid(ScriptPath, 1, Len(ScriptPath) - 3) & "iss"
   
  'Autodetect Inno
  If InnoEXE = "" Then
     InnoEXE = Find_Inno
     If InnoEXE = "" Then
        MsgBox "Unable to locate the Inno Compiler Program. Please configure the toolbar.", vbCritical + vbOKOnly, "VB6 - Inno Setup Toolbar"
        Call ShowConfig
        Exit Sub
     End If
  End If
  
  If ScriptPath = "" Then
     MsgBox "This project has not been saved yet.", vbCritical + vbOKOnly, "VB6 - Inno Setup Toolbar"
     Exit Sub
  End If
  
  If VBInstance.ActiveVBProject.IsDirty Then
     Rtn = MsgBox("This project has been changed since your last save. Continue Anyway?", vbExclamation + vbYesNo, "VB6 - Inno Setup Toolbar")
     If Rtn <> vbYes Then Exit Sub
  End If
  
  If Not File_Exists(ScriptPath) Then
     MsgBox "Unable to locate the script. You must generate a script first.", vbCritical + vbOKOnly, "VB6 - Inno Setup Toolbar"
     Exit Sub
  End If
  
  'Safety Check
  Rtn = MsgBox("Ready to compile the Inno Setup Script. Continue?", vbExclamation + vbYesNo, "VB6 - Inno Setup Toolbar")
  
  If Rtn <> vbYes Then Exit Sub
  
  'Recompile the Project
  If Get_Key("C0") = "1" Then
     Rtn = vbYes
  ElseIf Get_Key("C2") = "1" Then
     Rtn = vbNo
  Else
     Rtn = MsgBox("Would you like to recompile the project before continuing?", vbExclamation + vbYesNo, "VB6 - Inno Setup Toolbar")
  End If
   
  If Rtn = vbYes Then
     On Error Resume Next
     
     Err.Clear
     
     VBInstance.ActiveVBProject.MakeCompiledFile
     
     If Err.Number <> 0 Then
        Rtn = MsgBox("Failed to compile the project. It may be running. Continue anyway?", vbCritical + vbYesNo, "VB6 - Inno Setup Toolbar")
        If Rtn = vbNo Then Exit Sub
     End If
     
     On Error GoTo 0
  End If
  
  'Compile the Script
  Call API_WinExec(Chr(34) & InnoEXE & Chr(34) & " /cc " & Chr(34) & ScriptPath & Chr(34), True)
  
  'MsgBox "Script Compile Finished.", vbExclamation + vbOKOnly, "VB6 - Inno Setup Toolbar"
End Sub

'
' Configuration Button Event
'
Private Sub cmdBarBtnEvents4_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  'Configuration
  '------------------------------------------------
  
  Call ShowConfig
End Sub

'
' Show the Forms
Public Sub ShowConfig()
  If FLAG_ActiveModal Then Exit Sub
  
  If tmpConfig Is Nothing Then
     Set tmpConfig = New frmConfig
  End If
  
  Set tmpConfig.Connect = Me
  
  tmpConfig.Show
End Sub

Public Sub ShowDefault(ByVal ScriptPath As String)
  If FLAG_ActiveModal Then Exit Sub
  
  If tmpDefault Is Nothing Then
     Set tmpDefault = New frmDefault
  End If
  
  Set tmpDefault.Connect = Me
  tmpDefault.ScriptPath = ScriptPath
  
  tmpDefault.Show
End Sub

Public Sub ShowWizard(ByVal ScriptPath As String)
  If FLAG_ActiveModal Then Exit Sub
  
  If tmpWizard Is Nothing Then
     Set tmpWizard = New frmWizard
  End If
  
  Set tmpWizard.Connect = Me
  tmpWizard.ScriptPath = ScriptPath
  
  tmpWizard.Show
End Sub

'
'
' Reset the Forms
'
Public Sub ClearConfig()
  Set tmpConfig = Nothing
End Sub

Public Sub ClearDefault()
  Set tmpDefault = Nothing
End Sub

Public Sub ClearWizard()
  Set tmpWizard = Nothing
End Sub
