Attribute VB_Name = "modData"
'Visual Basic Interface.
Public VBInstance                     As VBIDE.VBE

Public InnoEXE As String


Public Enum Target_Enum
    TARGET_AppDir = 0     'Application Directory
    TARGET_PF = 1         'Program Files Directory
    TARGET_CF = 2         'Common Files Directory
    TARGET_WIN = 3        'Windows Directory
    TARGET_SYS = 4        'Windows System Directory
    TARGET_SOURCE = 5     'Setup Source Directory
    TARGET_DRIVE = 6      'System Drive Root Directory
    TARGET_CSTART = 7     'Common Startup Folder
    TARGET_USTART = 8     'User Startup Folder
End Enum

Public Type File_Struct
    Path As String
    Subdir As String
    Recurse As Boolean
    Target As Target_Enum
End Type

Public Files(0 To 200) As File_Struct
Public FileCnt As Integer

Public FLAG_ActiveModal As Boolean
