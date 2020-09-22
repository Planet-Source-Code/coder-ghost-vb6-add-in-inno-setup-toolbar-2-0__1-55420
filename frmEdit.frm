VERSION 5.00
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inno Setup Toolbar :: Script Wizard"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5160
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton btOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Source:"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.CheckBox chkRecurse 
         Caption         =   "Recurse Subdirectories"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtPath 
         BackColor       =   &H80000004&
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
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   4695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Destination:"
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   4935
      Begin VB.ComboBox lstTarget 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox txtSubdir 
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
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1560
         Width           =   4695
      End
      Begin VB.Label Label2 
         Caption         =   "Destination Subdirectory:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Destination Base Directory:"
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
         TabIndex        =   8
         Top             =   360
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Slot As Integer

Private Sub btCancel_Click()
  Unload Me
End Sub

Private Sub btOK_Click()
  If Slot >= LBound(Files) And Slot <= UBound(Files) Then
     Files(Slot).Recurse = CBool(chkRecurse.Value)
     Files(Slot).Subdir = txtSubdir.Text
     Files(Slot).Target = lstTarget.ListIndex
  End If
  
  Unload Me
End Sub

Private Sub Form_Load()
  txtPath.BackColor = RGB(200, 200, 200)
  
  lstTarget.AddItem "Application Directory", 0
  lstTarget.AddItem "Program Files Directory", 1
  lstTarget.AddItem "Common Files Directory", 2
  lstTarget.AddItem "Windows Directory", 3
  lstTarget.AddItem "Windows System Directory", 4
  lstTarget.AddItem "Setup Source Directory", 5
  lstTarget.AddItem "System Drive Root Directory", 6
  lstTarget.AddItem "Common Startup Folder", 7
  lstTarget.AddItem "User Startup Folder", 8
  
  FLAG_ActiveModal = True
End Sub

Public Sub Populate(ByVal ListIndex As Integer)
  Slot = ListIndex
  
  If Slot >= LBound(Files) And Slot <= UBound(Files) Then
     txtPath.Text = Files(Slot).Path
     chkRecurse.Value = 0 - CInt(Files(Slot).Recurse)
     txtSubdir.Text = Files(Slot).Subdir
     lstTarget.ListIndex = Files(Slot).Target
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FLAG_ActiveModal = False
End Sub
