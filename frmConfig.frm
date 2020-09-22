VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inno Setup Toolbar :: Configuration"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6135
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   5895
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Text            =   "Could Not Locate Inno Setup!"
         Top             =   600
         Width           =   5655
      End
      Begin VB.Label Label1 
         Caption         =   "Inno Install Directory:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.CommandButton btCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton btSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   5895
      Begin VB.CheckBox Options 
         Caption         =   "Disable Recompile Prompt"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   12
         ToolTipText     =   "You will never be prompted to recompile your application."
         Top             =   240
         Width           =   2895
      End
      Begin VB.CheckBox Options 
         Caption         =   "Automatically Recompile"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Script Compile will always automatically recompile the program."
         Top             =   240
         Width           =   2775
      End
      Begin VB.CheckBox Options 
         Caption         =   "Allow Unsafe Files"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "When creating a default script, it will be left empty."
         Top             =   600
         Width           =   2535
      End
      Begin VB.CheckBox Options 
         Caption         =   "Do Not Sniff Dependencies"
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   11
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.Frame Frame3 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   5895
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "You must have Inno Setup installed on your system!"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   195
         Width           =   5655
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   2505
      Width           =   4455
      Begin VB.Label Label7 
         Caption         =   "Mike Canejo - File Dependancy Sniffer"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         MouseIcon       =   "frmConfig.frx":038A
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Tag             =   "http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=31828&lngWId=1"
         ToolTipText     =   "Click to jump to Mike Canejo's File Dependancy Code on PSCODE.com"
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Label6 
         Caption         =   "Paul Mather - Common Dialog API Code"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         MouseIcon       =   "frmConfig.frx":04DC
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Tag             =   "http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=3592&lngWId=1"
         ToolTipText     =   "Click to jump to Paul Mather's Common Dialog API on PSCODE.com"
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "The Creators of Inno Setup"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         MouseIcon       =   "frmConfig.frx":062E
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Tag             =   "http://www.jrsoftware.org/isinfo.php"
         ToolTipText     =   "Click to jump to the Inno Setup Homepage"
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label lbPSC 
         Caption         =   "Planet Source Code"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         MouseIcon       =   "frmConfig.frx":0780
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Tag             =   "http://www.pscode.com/"
         ToolTipText     =   "Click to jump to Planet Source Code"
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Special Thanks To:"
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame5 
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   4080
      Width           =   5895
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Produced and Programmed by Final Stand Productions"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmConfig.frx":08D2
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Tag             =   "http://finalstand.archsysinc.com/"
         Top             =   200
         Width           =   5895
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Connect As Connect

Private Sub btCancel_Click()
  Unload Me
End Sub

Private Sub btSave_Click()
  If Find_Inno <> txtPath.Text Then
     Call Write_Key("ForcePath", txtPath.Text)
     InnoEXE = Find_Inno()
  End If
  
  For A = Options.LBound To Options.UBound
     Call Write_Key("C" & CStr(A), CStr(Options(A).Value))
  Next A
  
  Unload Me
End Sub

Private Sub Form_Load()
  Dim Rtn As String
  
  txtPath.Text = Find_Inno
  
  If txtPath.Text = "" Then txtPath.Text = "Could Not Locate Inno Setup!"
  
  For A = Options.LBound To Options.UBound
     Rtn = Get_Key("C" & CStr(A))
     If Rtn <> "" Then Options(A).Value = CInt(Rtn)
  Next A
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call Connect.ClearConfig
End Sub

Private Sub Label10_Click()

End Sub

Private Sub Label3_Click()
  Call ShellURL(Label3.Tag)
End Sub

Private Sub Label5_Click()
  Call ShellURL(Label5.Tag)
End Sub

Private Sub Label6_Click()
  Call ShellURL(Label6.Tag)
End Sub

Private Sub Label7_Click()
  Call ShellURL(Label7.Tag)
End Sub

Private Sub lbPSC_Click()
  Call ShellURL(lbPSC.Tag)
End Sub

Private Sub Options_Click(Index As Integer)
  Select Case Index
   Case 0:
     If Options(0).Value = 1 Then Options(2).Value = 0
   Case 2:
     If Options(2).Value = 1 Then Options(0).Value = 0
  End Select
End Sub
