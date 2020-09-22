VERSION 5.00
Begin VB.Form frmWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inno Setup Toolbar :: Script Wizard"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   7365
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btBack 
      Caption         =   "< &Back"
      Height          =   375
      Left            =   1620
      TabIndex        =   87
      Top             =   4860
      Width           =   1335
   End
   Begin VB.CommandButton btFinish 
      Caption         =   "Finish"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4500
      TabIndex        =   86
      Top             =   4860
      Width           =   1335
   End
   Begin VB.CommandButton btNext 
      Caption         =   "&Next >"
      Height          =   375
      Left            =   3060
      TabIndex        =   33
      Top             =   4860
      Width           =   1335
   End
   Begin VB.CommandButton btCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5940
      TabIndex        =   32
      Top             =   4860
      Width           =   1335
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   3
      Left            =   -120
      ScaleHeight     =   4665
      ScaleWidth      =   7665
      TabIndex        =   58
      Top             =   0
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CommandButton btBrowse 
         Caption         =   "B&rowse..."
         Height          =   360
         Index           =   0
         Left            =   5880
         TabIndex        =   12
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtField 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   4
         Left            =   600
         TabIndex        =   11
         Top             =   1440
         Width           =   5175
      End
      Begin VB.PictureBox picTop 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   2
         Left            =   -120
         ScaleHeight     =   915
         ScaleWidth      =   7635
         TabIndex        =   59
         Top             =   -120
         Width           =   7695
         Begin VB.PictureBox picWizard2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   750
            Index           =   2
            Left            =   6960
            Picture         =   "frmWizard.frx":038A
            ScaleHeight     =   750
            ScaleWidth      =   495
            TabIndex        =   60
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lbCaption2 
            BackStyle       =   0  'Transparent
            Caption         =   "Please specify the files that are part of your application."
            Height          =   255
            Index           =   2
            Left            =   720
            TabIndex        =   62
            Top             =   480
            Width           =   5895
         End
         Begin VB.Label lbCaption 
            BackStyle       =   0  'Transparent
            Caption         =   "Application Files"
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
            Index           =   2
            Left            =   480
            TabIndex        =   61
            Top             =   240
            Width           =   4215
         End
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Allow the user to start the application after setup has finished."
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   13
         Top             =   1920
         Value           =   1  'Checked
         Width           =   5655
      End
      Begin VB.Frame fraFiles 
         Height          =   2175
         Left            =   600
         TabIndex        =   65
         Top             =   2160
         Visible         =   0   'False
         Width           =   6735
         Begin VB.TextBox txtFile 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
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
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   96
            Top             =   1800
            Width           =   6375
         End
         Begin VB.ListBox lstFiles 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1530
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   4935
         End
         Begin VB.CommandButton btFiles 
            Caption         =   "&Remove..."
            Height          =   375
            Index           =   3
            Left            =   5160
            TabIndex        =   24
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CommandButton btFiles 
            Caption         =   "&Edit..."
            Height          =   375
            Index           =   2
            Left            =   5160
            TabIndex        =   17
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton btFiles 
            Caption         =   "Add &Directory..."
            Height          =   375
            Index           =   1
            Left            =   5160
            TabIndex        =   16
            Top             =   600
            Width           =   1455
         End
         Begin VB.CommandButton btFiles 
            Caption         =   "&Add file..."
            Height          =   375
            Index           =   0
            Left            =   5160
            TabIndex        =   15
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame fraSniff 
         Height          =   2055
         Left            =   600
         TabIndex        =   66
         Top             =   2160
         Width           =   6735
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Please Wait ... Locating Base Dependicies ..."
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   960
            Width           =   6255
         End
      End
      Begin VB.Label lbFileCount 
         Alignment       =   2  'Center
         Caption         =   "0 Files Included."
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   68
         Top             =   4380
         Width           =   4575
      End
      Begin VB.Label lbNote 
         Caption         =   "Application Main Executable or Component:"
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
         Index           =   4
         Left            =   600
         TabIndex        =   64
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label lbRequired 
         Caption         =   "* Bold Fields are required."
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   2
         Left            =   5320
         TabIndex        =   63
         Top             =   4380
         Width           =   2895
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   5
      Left            =   -120
      ScaleHeight     =   4665
      ScaleWidth      =   7665
      TabIndex        =   69
      Top             =   0
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox txtField 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   9
         Left            =   600
         TabIndex        =   20
         Top             =   3240
         Width           =   5175
      End
      Begin VB.CommandButton btBrowse 
         Caption         =   "B&rowse..."
         Height          =   360
         Index           =   3
         Left            =   5880
         TabIndex        =   23
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtField 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   8
         Left            =   600
         TabIndex        =   19
         Top             =   2280
         Width           =   5175
      End
      Begin VB.CommandButton btBrowse 
         Caption         =   "B&rowse..."
         Height          =   360
         Index           =   2
         Left            =   5880
         TabIndex        =   22
         Top             =   2280
         Width           =   1455
      End
      Begin VB.PictureBox picTop 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   4
         Left            =   -120
         ScaleHeight     =   915
         ScaleWidth      =   7635
         TabIndex        =   70
         Top             =   -120
         Width           =   7695
         Begin VB.PictureBox picWizard2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   750
            Index           =   4
            Left            =   6960
            Picture         =   "frmWizard.frx":1754
            ScaleHeight     =   750
            ScaleWidth      =   495
            TabIndex        =   71
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lbCaption 
            BackStyle       =   0  'Transparent
            Caption         =   "Application Documentation"
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
            Index           =   4
            Left            =   480
            TabIndex        =   73
            Top             =   240
            Width           =   4215
         End
         Begin VB.Label lbCaption2 
            BackStyle       =   0  'Transparent
            Caption         =   "Please specify which documentation files should be shown by the installer."
            Height          =   255
            Index           =   4
            Left            =   720
            TabIndex        =   72
            Top             =   480
            Width           =   6015
         End
      End
      Begin VB.TextBox txtField 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   7
         Left            =   600
         TabIndex        =   18
         Top             =   1320
         Width           =   5175
      End
      Begin VB.CommandButton btBrowse 
         Caption         =   "B&rowse..."
         Height          =   360
         Index           =   1
         Left            =   5880
         TabIndex        =   21
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lbNote 
         Caption         =   "Information File shown after Installation:"
         Height          =   255
         Index           =   9
         Left            =   600
         TabIndex        =   77
         Top             =   3000
         Width           =   3735
      End
      Begin VB.Label lbNote 
         Caption         =   "Information File shown before Installation:"
         Height          =   255
         Index           =   8
         Left            =   600
         TabIndex        =   76
         Top             =   2040
         Width           =   3735
      End
      Begin VB.Label lbRequired 
         Caption         =   "* Bold Fields are required."
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   4
         Left            =   5320
         TabIndex        =   75
         Top             =   4380
         Width           =   2895
      End
      Begin VB.Label lbNote 
         Caption         =   "License File:"
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   74
         Top             =   1080
         Width           =   3735
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   2
      Left            =   -120
      ScaleHeight     =   4665
      ScaleWidth      =   7665
      TabIndex        =   50
      Top             =   0
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CheckBox chkOption 
         Caption         =   "The Application does not need a directory."
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   29
         Top             =   3840
         Width           =   4695
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Allow the user to change the application directory."
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   28
         Top             =   3120
         Value           =   1  'Checked
         Width           =   4695
      End
      Begin VB.ComboBox lstOption 
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
         Index           =   0
         ItemData        =   "frmWizard.frx":2B1E
         Left            =   600
         List            =   "frmWizard.frx":2B28
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1320
         Width           =   4455
      End
      Begin VB.PictureBox picTop 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   1
         Left            =   -120
         ScaleHeight     =   915
         ScaleWidth      =   7635
         TabIndex        =   51
         Top             =   -120
         Width           =   7695
         Begin VB.PictureBox picWizard2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   750
            Index           =   1
            Left            =   6960
            Picture         =   "frmWizard.frx":2B4F
            ScaleHeight     =   750
            ScaleWidth      =   495
            TabIndex        =   52
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lbCaption 
            BackStyle       =   0  'Transparent
            Caption         =   "Application Directory"
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
            Index           =   1
            Left            =   480
            TabIndex        =   54
            Top             =   240
            Width           =   4215
         End
         Begin VB.Label lbCaption2 
            BackStyle       =   0  'Transparent
            Caption         =   "Please specify directory information for your application."
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   53
            Top             =   480
            Width           =   5895
         End
      End
      Begin VB.TextBox txtField 
         Enabled         =   0   'False
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
         Index           =   6
         Left            =   600
         TabIndex        =   26
         Top             =   1800
         Width           =   4455
      End
      Begin VB.TextBox txtField 
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
         Index           =   5
         Left            =   600
         TabIndex        =   27
         Text            =   "My Program"
         Top             =   2520
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "Other:"
         Height          =   255
         Left            =   600
         TabIndex        =   95
         Top             =   3600
         Width           =   3855
      End
      Begin VB.Label lbRequired 
         Caption         =   "* Bold Fields are required."
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   1
         Left            =   5320
         TabIndex        =   57
         Top             =   4380
         Width           =   2895
      End
      Begin VB.Label lbNote 
         Caption         =   "Application Destination Base Directory:"
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
         Index           =   7
         Left            =   600
         TabIndex        =   56
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label lbNote 
         Caption         =   "Application Directory Name:"
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
         Index           =   6
         Left            =   600
         TabIndex        =   55
         Top             =   2280
         Width           =   3855
      End
   End
   Begin VB.PictureBox picPage 
      BackColor       =   &H00FFFFFF&
      Height          =   4815
      Index           =   6
      Left            =   -120
      ScaleHeight     =   4755
      ScaleWidth      =   7635
      TabIndex        =   78
      Top             =   -120
      Width           =   7695
      Begin VB.PictureBox picEnd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4695
         Left            =   45
         ScaleHeight     =   4695
         ScaleWidth      =   7335
         TabIndex        =   79
         Top             =   70
         Width           =   7335
         Begin VB.PictureBox picWizard 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4695
            Index           =   1
            Left            =   0
            Picture         =   "frmWizard.frx":3F19
            ScaleHeight     =   4695
            ScaleWidth      =   2490
            TabIndex        =   80
            Top             =   0
            Width           =   2490
         End
         Begin VB.Label lbIntro 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Wizard Completed - Ready to Generate Script!"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   2520
            TabIndex        =   85
            Top             =   240
            Width           =   4815
         End
         Begin VB.Label lbIntro 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmWizard.frx":29FBB
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   8
            Left            =   2640
            TabIndex        =   84
            Top             =   720
            Width           =   4575
         End
         Begin VB.Label lbIntro 
            BackStyle       =   0  'Transparent
            Caption         =   "Not all features of Inno Setup are covered by this wizard. See the documentation for details on creating Inno Script setup files!"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   2640
            TabIndex        =   83
            Top             =   1800
            Width           =   4575
         End
         Begin VB.Label lbIntro 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmWizard.frx":2A057
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   735
            Index           =   6
            Left            =   2640
            TabIndex        =   82
            Top             =   2640
            Width           =   4575
         End
         Begin VB.Label lbIntro 
            BackStyle       =   0  'Transparent
            Caption         =   "Click Finish to generate the script, or Cancel to exit this wizard."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   2640
            TabIndex        =   81
            Top             =   3720
            Width           =   4575
         End
      End
   End
   Begin VB.PictureBox picPage 
      BackColor       =   &H00FFFFFF&
      Height          =   4815
      Index           =   0
      Left            =   -120
      ScaleHeight     =   4755
      ScaleWidth      =   7515
      TabIndex        =   30
      Top             =   -120
      Width           =   7575
      Begin VB.PictureBox picIntro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4695
         Left            =   45
         ScaleHeight     =   4695
         ScaleWidth      =   7335
         TabIndex        =   31
         Top             =   70
         Width           =   7335
         Begin VB.PictureBox picWizard 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4695
            Index           =   0
            Left            =   0
            Picture         =   "frmWizard.frx":2A10C
            ScaleHeight     =   4695
            ScaleWidth      =   2490
            TabIndex        =   34
            Top             =   0
            Width           =   2490
         End
         Begin VB.Label lbIntro 
            BackStyle       =   0  'Transparent
            Caption         =   "Click Next to continue, or Cancel to exit this wizard."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   2640
            TabIndex        =   39
            Top             =   3720
            Width           =   4575
         End
         Begin VB.Label lbIntro 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmWizard.frx":501AE
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   735
            Index           =   3
            Left            =   2640
            TabIndex        =   38
            Top             =   2640
            Width           =   4575
         End
         Begin VB.Label lbIntro 
            BackStyle       =   0  'Transparent
            Caption         =   "Not all features of Inno Setup are covered by this wizard. See the documentation for details on creating Inno Script setup files!"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   2640
            TabIndex        =   37
            Top             =   1920
            Width           =   4575
         End
         Begin VB.Label lbIntro 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmWizard.frx":50263
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   1
            Left            =   2640
            TabIndex        =   36
            Top             =   720
            Width           =   4575
         End
         Begin VB.Label lbIntro 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Welcome to the Visual Basic Inno Setup Wizard!"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2520
            TabIndex        =   35
            Top             =   240
            Width           =   4815
         End
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   1
      Left            =   -120
      ScaleHeight     =   4665
      ScaleWidth      =   7665
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox txtField 
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
         Index           =   3
         Left            =   600
         TabIndex        =   3
         Text            =   "http://www.mycompany.com"
         Top             =   3480
         Width           =   4455
      End
      Begin VB.TextBox txtField 
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
         Index           =   2
         Left            =   600
         TabIndex        =   2
         Text            =   "My Company, Inc."
         Top             =   2760
         Width           =   4455
      End
      Begin VB.TextBox txtField 
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
         Index           =   1
         Left            =   600
         TabIndex        =   1
         Text            =   "My Program 1.5"
         Top             =   2040
         Width           =   4455
      End
      Begin VB.TextBox txtField 
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
         Index           =   0
         Left            =   600
         TabIndex        =   0
         Text            =   "My Program"
         Top             =   1320
         Width           =   4455
      End
      Begin VB.PictureBox picTop 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   0
         Left            =   -120
         ScaleHeight     =   915
         ScaleWidth      =   7635
         TabIndex        =   41
         Top             =   -120
         Width           =   7695
         Begin VB.PictureBox picWizard2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   750
            Index           =   0
            Left            =   6960
            Picture         =   "frmWizard.frx":50336
            ScaleHeight     =   750
            ScaleWidth      =   495
            TabIndex        =   42
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lbCaption2 
            BackStyle       =   0  'Transparent
            Caption         =   "Please specify some basic information about your application."
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   44
            Top             =   480
            Width           =   5895
         End
         Begin VB.Label lbCaption 
            BackStyle       =   0  'Transparent
            Caption         =   "Application Information"
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
            Index           =   0
            Left            =   480
            TabIndex        =   43
            Top             =   240
            Width           =   4215
         End
      End
      Begin VB.Label lbNote 
         Caption         =   "Application Website:"
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   49
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Label lbNote 
         Caption         =   "Application Publisher:"
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   48
         Top             =   2520
         Width           =   3015
      End
      Begin VB.Label lbNote 
         Caption         =   "Application Name, Including Version:"
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
         Index           =   1
         Left            =   600
         TabIndex        =   47
         Top             =   1800
         Width           =   3135
      End
      Begin VB.Label lbNote 
         Caption         =   "Application Name:"
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
         Index           =   0
         Left            =   600
         TabIndex        =   46
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label lbRequired 
         Caption         =   "* Bold Fields are required."
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   0
         Left            =   5320
         TabIndex        =   45
         Top             =   4380
         Width           =   2895
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   4
      Left            =   -120
      ScaleHeight     =   4665
      ScaleWidth      =   7545
      TabIndex        =   88
      Top             =   0
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CheckBox chkOption 
         Caption         =   "Create an Uninstall Icon in the Start Menu folder."
         Height          =   240
         Index           =   6
         Left            =   600
         TabIndex        =   8
         Top             =   3000
         Value           =   1  'Checked
         Width           =   13395
      End
      Begin VB.PictureBox picTop 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   3
         Left            =   -120
         ScaleHeight     =   915
         ScaleWidth      =   7635
         TabIndex        =   89
         Top             =   -120
         Width           =   7695
         Begin VB.PictureBox picWizard2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   750
            Index           =   3
            Left            =   6960
            Picture         =   "frmWizard.frx":51700
            ScaleHeight     =   750
            ScaleWidth      =   495
            TabIndex        =   90
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lbCaption 
            BackStyle       =   0  'Transparent
            Caption         =   "Application Icons"
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
            Index           =   3
            Left            =   480
            TabIndex        =   92
            Top             =   240
            Width           =   4215
         End
         Begin VB.Label lbCaption2 
            BackStyle       =   0  'Transparent
            Caption         =   "Please specify which icons should be created for your application."
            Height          =   255
            Index           =   3
            Left            =   720
            TabIndex        =   91
            Top             =   480
            Width           =   5895
         End
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Allow user to change the Start Menu folder name."
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   5
         Top             =   1920
         Value           =   1  'Checked
         Width           =   4695
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Allow user to disable Start Menu folder creation."
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   6
         Top             =   2280
         Width           =   4695
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Create an Internet Shortcut in the Start Menu folder."
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   7
         Top             =   2640
         Width           =   4695
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Allow user to create a desktop icon."
         Height          =   255
         Index           =   7
         Left            =   600
         TabIndex        =   9
         Top             =   3480
         Value           =   1  'Checked
         Width           =   4695
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Allow user to create a Quick Launch icon."
         Height          =   240
         Index           =   8
         Left            =   600
         TabIndex        =   10
         Top             =   3855
         Width           =   4155
      End
      Begin VB.TextBox txtField 
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
         Index           =   10
         Left            =   600
         TabIndex        =   4
         Text            =   "My Project"
         Top             =   1320
         Width           =   5175
      End
      Begin VB.Label lbRequired 
         Caption         =   "* Bold Fields are required."
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   3
         Left            =   5320
         TabIndex        =   94
         Top             =   4380
         Width           =   2895
      End
      Begin VB.Label lbNote 
         Caption         =   "Application Start Menu Folder Name:"
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
         Index           =   10
         Left            =   600
         TabIndex        =   93
         Top             =   1080
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Connect As Connect

Public ScriptPath As String

Dim CurrentPage As Integer, i As Integer

Dim LastDir As String
Dim sOpen As SelectedFile

Const LAST_PAGE = 3

Private Sub btBack_Click()
  Call Show_Page(CurrentPage - 1)
End Sub

Private Sub btBrowse_Click(Index As Integer)
   FileDialog.sFilter = ""
   FileDialog.sFilter = FileDialog.sFilter & "All Files (*.*)" & Chr$(0) & "*.*" & Chr$(0)
   FileDialog.sFilter = FileDialog.sFilter & "All Files (*.*)" & Chr$(0) & "*.*"
    
   ' See Standard CommonDialog Flags for all options
   FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY
   FileDialog.sDlgTitle = "Add Files ..."
   FileDialog.sInitDir = LastDir & "\"
      
   sOpen = ShowOpen(Me.hWnd)
      
   If sOpen.bCanceled = False And UBound(sOpen.sFiles) > 0 Then
      LastDir = sOpen.sLastDirectory
      If Right(LastDir, 1) = "\" Then LastDir = Mid(LastDir, 1, Len(LastDir) - 1)
            
      MyFile = sOpen.sFiles(UBound(sOpen.sFiles))
   
      If InStr(1, MyFile, "\") = 0 Then MyFile = LastDir & "\" & MyFile
      
      Select Case Index
         Case 0: txtField(4).Text = MyFile
         Case 1: txtField(7).Text = MyFile
         Case 2: txtField(8).Text = MyFile
         Case 3: txtField(9).Text = MyFile
      End Select
   End If
End Sub

Private Sub btCancel_Click()
  Unload Me
End Sub

Private Sub btFiles_Click(Index As Integer)
  On Error Resume Next
  Dim Directory As String, Slot As Integer
  
  Slot = lstFiles.ListIndex
  
  Select Case Index
    Case 0: 'Add Files
      FileDialog.sFilter = ""
      FileDialog.sFilter = FileDialog.sFilter & "All Files (*.*)" & Chr$(0) & "*.*" & Chr$(0)
      FileDialog.sFilter = FileDialog.sFilter & "All Files (*.*)" & Chr$(0) & "*.*"
    
      ' See Standard CommonDialog Flags for all options
      FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
      FileDialog.sDlgTitle = "Add Files ..."
      FileDialog.sInitDir = LastDir & "\"
      
      sOpen = ShowOpen(Me.hWnd)
      
      Call Handle_Files
    Case 1: 'Add Folder
      Directory = BrowseForFolder(Me, "Select A Directory -", LastDir & "\")
      If Directory <> "" Then
         If Right(Directory, 1) = "\" Then Directory = Mid(Directory, 1, Len(Directory) - 1)
         LastDir = Directory
         Rtn = MsgBox("Would you like to recurse the subdirectories?", vbQuestion + vbYesNo, "VB6 - Inno Setup Toolbar")
         Files(FileCnt).Path = Directory & "\*"
         Files(FileCnt).Recurse = IIf(Rtn = vbYes, True, False)
         Files(FileCnt).Subdir = ""
         Files(FileCnt).Target = TARGET_AppDir
         FileCnt = FileCnt + 1
         Call Redraw_Files
      End If
    Case 2: 'Edit
      If Slot > -1 Then
         Dim tmpEdit As frmEdit
         
         Set tmpEdit = New frmEdit
         Call tmpEdit.Populate(Slot)
         
         tmpEdit.Show 1, Me
      End If
    Case 3: 'Remove
      If Slot > -1 Then
         For A = Slot To FileCnt - 1
           Files(A).Path = Files(A + 1).Path
           Files(A).Target = Files(A + 1).Target
           Files(A).Recurse = Files(A + 1).Recurse
           Files(A).Subdir = Files(A + 1).Subdir
         Next A
         FileCnt = FileCnt - 1
         
         Tmp = lstFiles.ListIndex
         
         Call Redraw_Files
         
         If Tmp > (lstFiles.ListCount - 1) Then Tmp = Tmp - 1
         lstFiles.ListIndex = Tmp
      End If
  End Select
End Sub

Public Sub Redraw_Files()
  lstFiles.Clear
  
  For A = 0 To FileCnt - 1
    lstFiles.AddItem Files(A).Path
    lstFiles.ListIndex = lstFiles.ListIndex + 1
  Next A
  
  txtFile.Text = ""
  lbFileCount.Caption = FileCnt & " Files Included."
End Sub

Private Sub btNext_Click()
  Dim CPage As Integer
  
  Rtn = Validate_Page(CurrentPage)
  
  If Rtn = True Then
     CPage = CurrentPage
     
     Call Show_Page(CurrentPage + 1)
   
     If CurrentPage > CPage Then Call SetupPage(CurrentPage)
  Else
     Beep
  End If
End Sub

Private Sub chkOption_Click(Index As Integer)
  Select Case Index
    Case 1:
      lstOption(0).ListIndex = 0
      chkOption(0).Enabled = Not CBool(chkOption(1).Value)
      chkOption(2).Enabled = Not CBool(chkOption(1).Value)
      lstOption(0).Enabled = Not CBool(chkOption(1).Value)
      txtField(4).Enabled = Not CBool(chkOption(1).Value)
      txtField(4).BackColor = IIf(Not CBool(chkOption(1).Value), vbWhite, RGB(200, 200, 200))
      txtField(5).Enabled = Not CBool(chkOption(1).Value)
      txtField(5).BackColor = IIf(Not CBool(chkOption(1).Value), vbWhite, RGB(200, 200, 200))
      lstOption(0).BackColor = IIf(Not CBool(chkOption(1).Value), vbWhite, RGB(200, 200, 200))
  End Select
End Sub

Private Sub Form_Load()
  Call Show_Page(0)
  
  For i = lstOption.LBound To lstOption.UBound
    lstOption(i).ListIndex = 0
  Next i
  
  'Main Executable
  txtField(4).Text = VBInstance.ActiveVBProject.BuildFileName
End Sub

Public Sub Show_Page(ByVal Page As Integer)
  If Page > picPage.UBound Then Exit Sub
  If Page < picPage.LBound Then Exit Sub
  
  If chkOption(1).Value = 1 And Page = 4 Then
     If Page < CurrentPage Then Page = Page - 1
     If Page > CurrentPage Then Page = Page + 1
  End If
  
  CurrentPage = Page
   
  For A = picPage.LBound To picPage.UBound
    If Page = A Then
       picPage(A).Visible = True
    Else
       picPage(A).Visible = False
    End If
  Next A
    
  btBack.Visible = True
  btNext.Enabled = True
  btFinish.Enabled = False
  
  If Page = 6 Then btFinish.Enabled = True
  If Page = picPage.LBound Then btBack.Visible = False
  If Page = picPage.UBound Then btNext.Enabled = False
End Sub

Public Function Validate_Page(ByVal Page As Integer) As Boolean
  Validate_Page = True
  
  Select Case Page
    Case 1:
      If Len(Trim(txtField(0).Text)) <= 0 Then Validate_Page = False
      If Len(Trim(txtField(1).Text)) <= 0 Then Validate_Page = False
    Case 2:
      If Len(Trim(txtField(5).Text)) <= 0 Then Validate_Page = False
      If Len(Trim(txtField(6).Text)) <= 0 Then
         If lstOption(0).ListIndex = 1 Then Validate_Page = False
      End If
    Case 3:
      If Len(Trim(txtField(4).Text)) <= 0 Then Validate_Page = False
    Case 4:
      If Len(Trim(txtField(10).Text)) <= 0 Then Validate_Page = False
  End Select
End Function

Private Sub lstFiles_Click()
  On Error Resume Next
  
  txtFile.Text = lstFiles.List(lstFiles.ListIndex)
  
  lbFileCount.Caption = lstFiles.ListCount & " Files Included."
End Sub

Private Sub lstOption_Click(Index As Integer)
  Select Case Index
    Case 0: 'Application Base Directory
      txtField(6).Enabled = CBool(lstOption(0).ListIndex)
      txtField(6).BackColor = IIf(CBool(lstOption(0).ListIndex), vbWhite, RGB(200, 200, 200))
  End Select
End Sub

Public Sub SetupPage(ByVal Page As Integer)
  Select Case Page
   Case 3:
    Call Search_Depends
  End Select
End Sub

Public Sub Search_Depends()
  Dim Path As String, Sniff As Boolean
  Dim A As Integer, B As Integer
  
  Sniff = True
  
  If Get_Key("C3") = "1" Then
     Sniff = False
  Else
     lstFiles.Clear
     FileCnt = 0
  End If
  
  fraFiles.Visible = False
  fraSniff.Visible = True
  
  btBack.Visible = False
  btNext.Visible = False
     
     'Scan References and add them
     'For A = 1 To VBInstance.ActiveVBProject.References.Count
     '   Path = VBInstance.ActiveVBProject.References(A).FullPath
     '
     '   If Trim(Path) <> "" Then
     '      PntA = InStrRev(Path, ".")
     '      PntB = InStr(PntA, Path, "\")
     '
     '      If PntB > 0 Then Path = Mid(Path, 1, PntB - 1)
     '
     '      For B = 1 To lstFiles.ListCount
     '        If LCase(lstFiles.List(B - 1)) = LCase(Path) Then Path = ""
     '      Next B
     '
     '      If Path <> "" Then lstFiles.AddItem Path
     '   End If
     'Next A
     
     'Scan Compiled File for Dependancies
     If Sniff Then Call Sniff_File(txtField(4).Text, lstFiles, 0)
     
     Call Redraw_Files
     
  btBack.Visible = True
  btNext.Visible = True
  
  fraSniff.Visible = False
  fraFiles.Visible = True
End Sub

Public Sub Handle_Files()
  If sOpen.bCanceled = False Then
     LastDir = sOpen.sLastDirectory
      
     If Right(LastDir, 1) = "\" Then LastDir = Mid(LastDir, 1, Len(LastDir) - 1)
     
     For A = LBound(sOpen.sFiles) To UBound(sOpen.sFiles)
        FileName = sOpen.sFiles(A)
        
        If InStr(1, FileName, "\") = 0 Then FileName = LastDir & "\" & FileName
        
        Files(FileCnt).Path = FileName
        Files(FileCnt).Recurse = False
        Files(FileCnt).Subdir = ""
        Files(FileCnt).Target = TARGET_AppDir
        FileCnt = FileCnt + 1
     Next A
     
     Call Redraw_Files
  End If
End Sub

Private Sub btFinish_Click()
  'Okay, Build the Damn Script already.
  'On Error GoTo Fail
  
  Open ScriptPath For Output As #1
  
  EXEName = Mid(txtField(4).Text, InStrRev(txtField(4).Text, "\") + 1)
  ShortName = Mid(EXEName, 1, InStr(1, EXEName, ".") - 1)
  
  '
  ' Header
  '
  Print #1, "; "
  Print #1, "; Install Script for " & txtField(0).Text
  Print #1, ";  [Inno Setup Toolbar - Script Wizard]"
  Print #1, "; "
  Print #1, "; Generated by the 'Inno Setup Toolbar for VB6'"
  Print #1, "; Written and Programmed by Brian Haase"
  Print #1, "; "
  Print #1, "; Generated for Inno Setup Compiler, Version 4+"
  Print #1, "; "
  Print #1, "  "
  
  'Section: Setup
  '---------------------------------------------------------------------------
  Print #1, "[Setup]"
  Print #1, "AppName=" & txtField(0).Text
  Print #1, "AppVerName=" & txtField(1).Text
  Print #1, "AppPublisher=" & txtField(2).Text
  Print #1, "AppPublisherURL=" & txtField(3).Text
  Print #1, "AppSupportURL=" & txtField(3).Text
  Print #1, "AppUpdatesURL=" & txtField(3).Text

  If chkOption(1).Value = 1 Then
     Print #1, "CreateAppDir = no"
  Else
     If lstOption(0).ListIndex = 1 Then
        Print #1, "DefaultDirName=" & txtField(6).Text & "\" & txtField(5).Text
     Else
        Print #1, "DefaultDirName={pf}\" & txtField(5).Text
     End If
  
     Print #1, "DefaultGroupName=" & txtField(10).Text
  
     'Print #1, "AllowNoIcons=yes"
  End If
  
  If txtField(7).Text <> "" Then Print #1, "LicenseFile=" & txtField(7).Text
  If txtField(8).Text <> "" Then Print #1, "InfoBeforeFile=" & txtField(8).Text
  If txtField(9).Text <> "" Then Print #1, "InfoAfterFile=" & txtField(9).Text

  Print #1, "Compression = lzma"
  Print #1, "SolidCompression = yes"
  Print #1, "  "

  'Section: Tasks
  '---------------------------------------------------------------------------
  If (chkOption(7).Value + chkOption(8).Value) > 0 Then Print #1, "[Tasks]"
  If chkOption(7).Value = 1 Then Print #1, "Name: ""desktopicon""; Description: ""{cm:CreateDesktopIcon}""; GroupDescription: ""{cm:AdditionalIcons}""; Flags: unchecked"
  If chkOption(8).Value = 1 Then Print #1, "Name: ""quicklaunchicon""; Description: ""{cm:CreateQuickLaunchIcon}""; GroupDescription: ""{cm:AdditionalIcons}""; Flags: unchecked"
  Print #1, "  "

  'Section: Files
  '---------------------------------------------------------------------------
  Print #1, "[Files]"
  If chkOption(1).Value = 0 Then Print #1, "Source: """ & txtField(4).Text & """; DestDir: ""{app}""; Flags: ignoreversion"
 
  For A = 0 To FileCnt - 1
    Path = Keycode(Files(A).Target)
    
    tFlags = Smartflag(Files(A).Path)
    If Files(A).Recurse Then tFlags = tFlags & " recursesubdirs"
    
    If Files(A).Subdir <> "" Then Path = Path & "\" & Files(A).Subdir
    
    Print #1, "Source: """ & Files(A).Path & """; DestDir: """ & Path & """; Flags: " & tFlags
  Next A
  
  Print #1, "; NOTE: Don't use ""Flags: ignoreversion"" on any shared system files"
  Print #1, "  "

  'Section: INI
  '---------------------------------------------------------------------------
  If chkOption(5).Value = 1 Then
     Print #1, "[INI]"
     Print #1, "Filename: ""{app}\" & ShortName & ".url""; Section: ""InternetShortcut""; Key: ""URL""; String: """ & txtField(3).Text & """"
     Print #1, "  "
  End If
  
  'Section: Icons
  '---------------------------------------------------------------------------
  If chkOption(1).Value = 0 Then
     Print #1, "[Icons]"
     Print #1, "Name: ""{group}\" & txtField(0).Text & """; Filename: ""{app}\" & EXEName & """"
     If chkOption(5).Value = 1 Then Print #1, "Name: ""{group}\{cm:ProgramOnTheWeb," & txtField(0).Text & "}""; Filename: ""{app}\" & ShortName & ".url"""
     If chkOption(6).Value = 1 Then Print #1, "Name: ""{group}\{cm:UninstallProgram," & txtField(0).Text & "}""; Filename: ""{uninstallexe}"""
     If chkOption(7).Value = 1 Then Print #1, "Name: ""{userdesktop}\" & txtField(0).Text & """; Filename: ""{app}\" & EXEName & """; Tasks: desktopicon"
     If chkOption(8).Value = 1 Then Print #1, "Name: ""{userappdata}\Microsoft\Internet Explorer\Quick Launch\" & txtField(0).Text & """; Filename: ""{app}\" & EXEName & """; Tasks: quicklaunchicon"
     Print #1, "  "
  End If
  
  'Section: Run
  '---------------------------------------------------------------------------
  If chkOption(2).Value = 1 Then
     Print #1, "[Run]"
     Print #1, "Filename: ""{app}\" & EXEName & """; Description: ""{cm:LaunchProgram," & txtField(0).Text & "}""; Flags: nowait postinstall skipifsilent"
     Print #1, "  "
  End If
  
  'Section: Uninstall
  '---------------------------------------------------------------------------
  If chkOption(6).Value = 1 Then
     Print #1, "[UninstallDelete]"
     Print #1, "Type: files; Name: ""{app}\" & ShortName & ".url"""
     Print #1, "  "
  End If
  
  Close #1
  
  'Open the Script
  Call API_WinExec(Chr(34) & InnoEXE & Chr(34) & " " & Chr(34) & ScriptPath & Chr(34), False)

  Unload Me
   
  Exit Sub
Fail:
  MsgBox "An Error Occured while trying to write the script to disk.", vbExclamation + vbOKOnly, "VB6 - Inno Setup Toolbar"
End Sub

Public Function Keycode(ByVal Target As Integer) As String
    Select Case Target
      Case 0: Rtn = "{app}"
      Case 1: Rtn = "{pf}"
      Case 2: Rtn = "{cf}"
      Case 3: Rtn = "{win}"
      Case 4: Rtn = "{sys}"
      Case 5: Rtn = "{src}"
      Case 6: Rtn = "{sd}"
      Case 7: Rtn = "{commonstartup}"
      Case 8: Rtn = "{userstartup}"
    End Select
    
    Keycode = Rtn
End Function

Public Function Smartflag(ByVal EXE As String) As String
  EXEName = Mid(EXE, InStrRev(EXE, "\") + 1)
  
  If InStr(1, EXEName, ".") <= 0 Then
     Smartflag = "ignoreversion confirmoverwrite onlyifdoesntexist"
     Exit Function
  End If
  
  EXEType = LCase(Mid(EXEName, InStr(1, EXEName, ".") + 1))
     
  Select Case EXEType
    Case "exe": Smartflag = "ignoreversion confirmoverwrite onlyifdoesntexist"
    Case "dll": Smartflag = "confirmoverwrite onlyifdoesntexist regserver sharedfile"
    Case "ocx": Smartflag = "confirmoverwrite onlyifdoesntexist regserver sharedfile"
    Case "tlb": Smartflag = "confirmoverwrite onlyifdoesntexist regtypelib sharedfile"
    Case Else:  Smartflag = "ignoreversion confirmoverwrite onlyifdoesntexist"
  End Select
    
  If Get_Key("C1") = "1" Then
    Select Case EXEType
      Case "dll": Smartflag = Smartflag & " allowunsafefiles"
      Case "ocx": Smartflag = Smartflag & " allowunsafefiles"
      Case "tlb": Smartflag = Smartflag & " allowunsafefiles"
    End Select
  End If
End Function
