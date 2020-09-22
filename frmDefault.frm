VERSION 5.00
Begin VB.Form frmDefault 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inno Setup Toolbar :: Default Script"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDefault.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   5415
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btCancel 
      Caption         =   "C&ancel"
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton btCreate 
      Caption         =   "&Create Script"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   5175
      Begin VB.ComboBox lstItem 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Select Script Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   280
         Width           =   1575
      End
      Begin VB.Label lbDesc 
         Caption         =   "(Description)"
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   4935
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   2415
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   5175
      Begin VB.TextBox txtSource 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Text            =   "MyProject.exe"
         Top             =   1920
         Width           =   4935
      End
      Begin VB.TextBox txtDir 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2640
         TabIndex        =   4
         Text            =   "MyProject"
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtGroup 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Text            =   "MyProject"
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtVersion 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2640
         TabIndex        =   2
         Text            =   "1.0"
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtProject 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Text            =   "MyProject"
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label8 
         Caption         =   "Source Name"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Directory Name"
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Group Name"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Project Version"
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Project Name"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   5175
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "No settings are required."
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmDefault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Connect As Connect

Public ScriptPath As String

Private Sub btCancel_Click()
  Unload Me
End Sub

Private Sub btCreate_Click()
  'Create the Script
  Select Case lstItem.ListIndex
   Case 0: 'Empty Script
    Call Empty_Script(ScriptPath)
   Case 1: 'Simple EXE
    Call BasicEXE_Script(ScriptPath, txtProject.Text, txtVersion.Text, txtDir.Text, txtGroup.Text, txtSource.Text)
   Case 2: 'DLL
    Call DLL_Script(ScriptPath, txtProject.Text, txtVersion.Text, txtDir.Text, txtGroup.Text, txtSource.Text)
   Case 3: 'OCX
    Call OCX_Script(ScriptPath, txtProject.Text, txtVersion.Text, txtDir.Text, txtGroup.Text, txtSource.Text)
   Case 4: 'TLB
    Call TLB_Script(ScriptPath, txtProject.Text, txtVersion.Text, txtDir.Text, txtGroup.Text, txtSource.Text)
  End Select
  
  'Open the Script
  Call API_WinExec(Chr(34) & InnoEXE & Chr(34) & " " & Chr(34) & ScriptPath & Chr(34), False)
  
  Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
 lstItem.Clear
 lstItem.AddItem "Empty Script"
 lstItem.AddItem "Simple EXE Deployment Script"
 lstItem.AddItem "DLL Deployment Script"
 lstItem.AddItem "OCX Deployment Script"
 lstItem.AddItem "TLB Deployment Script"
 lstItem.ListIndex = 0
 
 txtProject.Text = VBInstance.ActiveVBProject.Name
 txtVersion.Text = txtProject.Text & " 1.0"
 txtGroup.Text = txtProject.Text
 txtDir.Text = txtProject.Text
 txtSource.Text = VBInstance.ActiveVBProject.BuildFileName
 
'  AppName = VBInstance.ActiveVBProject.Name
'  AppVersion = "1.0"
'  DirName = AppName
'  GroupName = AppName
'  SourceName = VBInstance.ActiveVBProject.BuildFileName
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Call Connect.ClearDefault
End Sub

Private Sub lstItem_Click()
 fraDetails.Visible = True
 
 Select Case lstItem.ListIndex
  Case 0: 'Empty Script
   lbDesc.Caption = "Creates an empty script, so you can build your installer from the ground up. Best for complex projects."
   fraDetails.Visible = False
  Case 1: 'Simple EXE
   lbDesc.Caption = "Creates a basic executable deployment script. Only includes the executable file, not any of the other dependicies."
  Case 2: 'DLL
   lbDesc.Caption = "Creates a ready-to-compile script to deploy and register a single DLL file. Includes uninstaller option by default."
  Case 3: 'OCX
   lbDesc.Caption = "Creates a ready-to-compile script to deploy and register a single OCX file. Includes uninstaller option by default."
  Case 4: 'TLB
   lbDesc.Caption = "Creates a ready-to-compile script to deploy and register a single TLB file. Includes uninstaller option by default."
 End Select
End Sub
