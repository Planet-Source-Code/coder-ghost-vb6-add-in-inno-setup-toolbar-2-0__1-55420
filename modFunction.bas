Attribute VB_Name = "modFunction"
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWDEFAULT = 10

Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Const LB_FINDSTRINGEXACT = &H1A2

Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long



Public Function Find_Inno() As String
  Dim Location As String
  
  'Find_Inno = "C:\Program Files\Inno Setup 3\Compil32.exe"
  
  Find_Inno = Get_Key("ForcePath")
  
  If Find_Inno = "" Then
     Find_Inno = GetKeyValue(HKEY_LOCAL_MACHINE, "Software\CLASSES\InnoSetupScriptFile\DefaultIcon", "")
  
     Find_Inno = Mid(Find_Inno, 1, InStrRev(Find_Inno, ",") - 1)
  End If
  
  'HKEY_LOCAL_MACHINE\Software\CLASSES\InnoSetupScriptFile\DefaultIcon     :: (Default)
End Function

Public Function File_Exists(ByVal Path As String) As Boolean
  On Error GoTo Fallout
  
  File_Exists = False
  
  Open Path For Input As #1
  Close #1
    
  File_Exists = True
  
Fallout:
End Function

Public Sub API_WinExec(ByVal Command As String, ByVal Hidden As Boolean)
  Dim Mode As Integer
  
  'Debug.Print Command
  
  Mode = SW_SHOWDEFAULT
  If Hidden Then Mode = SW_SHOWMINNOACTIVE
  
  WinExec Command, Mode
End Sub

Public Function Get_Key(ByVal Name As String) As String
  Get_Key = GetKeyValue(HKEY_LOCAL_MACHINE, "Software\Final Stand\InnoToolbar2", Name)
End Function

Public Function Write_Key(ByVal Name As String, ByVal Value As String) As Long
  Dim Rtn As Long
  
  Rtn = UpdateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Final Stand\InnoToolbar2", Name, Value)
  
  Write_Key = Rtn
End Function

Public Sub ShellURL(ByVal Msg As String)
    Call ShellExecute(0&, vbNullString, Msg, vbNullString, vbNullString, vbNormalFocus)
End Sub

Public Function KillDupesAPI(lpBox As Control) As Long
On Error Resume Next
Dim nCount As Long, nPos1 As Long
Dim nPos2 As Long, nTotal As Long
    nTotal = lpBox.ListCount
    For nCount = 0 To nTotal
        Do
            nPos1 = SendMessageByString(lpBox.hWnd, LB_FINDSTRINGEXACT, nCount, lpBox.List(nCount))
            nPos2 = SendMessageByString(lpBox.hWnd, LB_FINDSTRINGEXACT, nPos1 + 1, lpBox.List(nCount))
            If Trim(lpBox.List(nCount)) = vbNullString Then lpBox.RemoveItem nCount
            If nPos2 = -1 Or nPos2 = nPos1 Then Exit Do
            lpBox.RemoveItem nCount
        Loop
    Next nCount
    KillDupesAPI = nTotal - lpBox.ListCount
End Function

Public Function GetSystemPath() As String
  Dim Data As String * 255
  
  GetSystemPath = left(Data, GetSystemDirectory(Data, 255))
End Function

Public Function IsDuplicate(ByVal Item As String) As Boolean
  IsDuplicate = True
  
  Item = LCase(Item)
  
  If FileCnt > 0 Then
    For A = 0 To FileCnt
      If LCase(Files(A).Path) = Item Then Exit Function
    Next A
  End If
  
  IsDuplicate = False
End Function
