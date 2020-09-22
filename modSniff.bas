Attribute VB_Name = "modSniff"
'#######################################################
'#                                                     #
'#   - HUGE - Thanks to Mike Canejo for this code!     #
'#                                                     #
'#######################################################
'#                                                     #
'#  --------------------------------                   #
'#    File Dependency Sniffer v3                       #
'#    Written by Mike Canejo                           #
'#  --------------------------------                   #
'#  AIM (aol): Mike3dd                                 #
'#  Email: MikeCanejo@hotmail.com                      #
'#                                                     #
'#######################################################
'
Public Sub Sniff_File(ByVal File As String, ByVal List As ListBox, ByVal Depth As Integer)
  'Keep Lowercase
  Const EXTENSION = "*.dll *.ocx *.tlb *.exe"
  
  On Error GoTo Fallout
  
  Dim i As Integer
  Dim X As Integer
  Dim Z As Integer
  Dim iFind As Long
  Dim sExt As String
  Dim iLen As Integer
  Dim sFile As String
  Dim iFree As Integer
  Dim sFound As String
  Dim sQuery As String
  Dim bValid As Boolean
  Dim iTerminator(1) As Long
  
  If Depth >= 2 Then Exit Sub
  
  iFree = FreeFile                                    'Get an unused file number
  Open File For Binary Access Read As #iFree          'Opens the file for reading
     sFile = Space(LOF(iFree))                        'Puts the null terminator at the end of string variable
     Get #iFree, , sFile                              'Can now put the file into the variable cause it has space to accommodate it
  Close #iFree                                        'Close the process
    
  sFile = LCase(sFile)                                'To prevent search ambiguity, make it non case sensitive
    
  'Searches for all files with the specified extensions.
  For i = 1 To CharsIN(EXTENSION, "*")                                  'Search for each ext in query
      sExt = Mid$(EXTENSION, CharsPOS(EXTENSION, "*", i) + 1, 4)
      
      Do
         DoEvents
         
         iFind = InStr(iFind + 1, sFile, sExt)                          'Find the file extention in the string
         
         If iFind = 0 Then Exit Do
         
         iTerminator(0) = InStrRev(sFile, Chr(0), iFind)                'Chr(0) is used to determinate the beginning of the file found and the ending
         iTerminator(1) = iFind + 4                                     'To determine the ending, ifind is the start of say for example ".dll", +4 cause theres 4 letters so it gets the ending
            
         If iTerminator(0) And Mid$(sFile, _
            iTerminator(1), 1) = Chr$(0) Then                           'Beginning point and end point of file in string
                
                If iTerminator(1) - iTerminator(0) - 1 < 20 _
                 And iTerminator(1) - iTerminator(0) - 1 > 5 Then       'Some parameters to make sure the findings are not something other than files
                                                                        'This assumes all dlls are less than a length of 20 and greater than 1 char length.
                     bValid = True
                        
                     sFound = Mid$(sFile, iTerminator(0) + 1, iTerminator(1) - iTerminator(0) - 1)
                        
                     'Filter
                     bValid = isFilename(sFound, sExt)
                     
                     If LCase(sFound) <> "vba6.dll" Then
                        sFound = GetSystemPath & "\" & sFound
                        
                        If bValid Then bValid = Not IsDuplicate(sFound)
                        
                        If bValid Then
                           'List.AddItem sFound                 'If not detected in filters then add it
                           Files(FileCnt).Path = sFound
                           Files(FileCnt).Recurse = False
                           Files(FileCnt).Subdir = ""
                           Files(FileCnt).Target = TARGET_SYS
                           FileCnt = FileCnt + 1
                        End If
                        
                        'Extended Sniffing -
                        If Right(sFound, 3) = "ocx" Or Right(sFound, 3) = "tlb" Then
                           Call Sniff_File(sFound, List, Depth + 1)
                        End If
                     End If
                    
                    'Debug.Print sFound                                 'Display in immediate window
                End If
         End If
      Loop
      
      iFind = 0
   Next i
    
   KillDupesAPI List                      'Remove doubles found from search
   
Fallout:
End Sub


Private Function isFilename(sFilename As String, sExtention As String) As Boolean
On Error Resume Next
    '                       bValid = isFilename(sFound, sExt)
    '                       sFilename points to sFound and sExtention points to sExt
                        
    'NOTE: sFilename is a pointer to a var in reference to it from the function syntax.
    'So if the contents of any pointer changes, then the var its pointing
    'to does as well.. since its really sFound with a different name in memory.
    'The search code above uses this function to check
    'a found filename using var sFound so when sFilename is changed anywhere below,
    'sFound, which its pointing to, does as well... I'm just pointing this
    'out for newcommers because this is something that took me a while
    'to figure out when i just started to write functions in vb way back when
    'and if your learning vb on your own like i did then this is something you
    'figure out by trial and error usually... c++ helped me as well..but that's
    'a different story.
    
    'okie dokie
    '-Mike Canejo  ;]
     
    Dim i As Integer, X As Integer
    Dim iLen As Integer
    
    If InStr(sFilename, "\") Then
        sFilename = Mid(sFilename, _
        InStr(sFilename, "\") + 1)      'Sometimes "\" are in the filename
                                        'cause of paths in the exe.
                                        'So this will get the right of it.
    End If
    
    
    
                                        'Two loops below to detect a funky char in the filename.
                                        'If one is found it will cut the string off
                                        'at the pos its found cause almost every time
                                        'it's a wrong beginning of the filename being found.
                                        'The ending is always correct on the Chr(0) finding.
                                        'So this so far, as I can see, takes care of it...
                                        
                                        'Please leave feedback if I am wrong!
                                        
                                        'Again-
                                        'AIM (aol): Mike3dd
                                        'E-mail: MikeCanejo@hotmail.com
                    
                    
    For i = Len(sFilename) To 1 Step -1             'Start searching from end of string to beginning
        For X = 1 To 39
            If Mid$(sFilename, i, 1) = Chr$(X) _
            Or Mid$(sFilename, i, 1) = Chr$(96) Then
                sFilename = Mid$(sFilename, i + 1)   'Funky char found, cut it at the pos
                Exit For                            'Exit the loop since it found it
            End If
        Next X
    Next i
    For i = Len(sFilename) To 1 Step -1             'Start searching from end of string to beginning
        For X = 123 To 255
            If Mid$(sFilename, i, 1) = Chr$(X) _
            Or Mid$(sFilename, i, 1) = Chr$(96) Then
                sFilename = Mid$(sFilename, i + 1)   'Funky char found, cut it at the pos
                Exit For                            'Exit the loop since it found it
            End If
        Next X
    Next i
    
    iLen = Len(Left(sFilename, InStr( _
    sFilename, sExtention) - 1))        'Length parameters to ensure the filtered filename
                                        'is considered a "correct" file name length.
                                        'You can change this to your own liking...
                                                                        
    If iLen < 20 And iLen > 1 Then
        isFilename = True
    Else
        isFilename = False
    End If

End Function

Private Function isFunky(sCheck As String) As Boolean
'On Error Resume Next
    '                       bValid = isFilename(sFound, sExt)
    '                       sFilename points to sFound and sExtention points to sExt
                        
    'NOTE: sFilename is a pointer to a var in reference to it from the function syntax.
    'So if the contents of any pointer changes, then the var its pointing
    'to does as well.. since its really sFound with a different name in memory.
    'The search code above uses this function to check
    'a found filename using var sFound so when sFilename is changed anywhere below,
    'sFound, which its pointing to, does as well... I'm just pointing this
    'out for newcommers because this is something that took me a while
    'to figure out when i just started to write functions in vb way back when
    'and if your learning vb on your own like i did then this is something you
    'figure out by trial and error usually... c++ helped me as well..but that's
    'a different story.
    
    'okie dokie
    '-Mike Canejo  ;]
     
    Dim i As Integer, X As Integer
    
    
    
    
                                        'Two loops below to detect a funky char in the filename.
                                        'If one is found it will cut the string off
                                        'at the pos its found cause almost every time
                                        'it's a wrong beginning of the filename being found.
                                        'The ending is always correct on the Chr(0) finding.
                                        'So this so far, as I can see, takes care of it...
                                        
                                        'Please leave feedback if I am wrong!
                                        
                                        'Again-
                                        'AIM (aol): Mike3dd
                                        'E-mail: MikeCanejo@hotmail.com
                    

    For i = Len(sCheck) To 1 Step -1             'Start searching from end of string to beginning
        For X = 1 To 39
            If Mid(sCheck, i, 1) = Chr(X) Then
                isFunky = True  'Funky char found, cut it at the pos
                Exit Function                         'Exit the loop since it found it
            End If
        Next X
    Next i
    
    
    For i = Len(sCheck) To 1 Step -1             'Start searching from end of string to beginning
        For X = 123 To 255
            If Mid(sCheck, i, 1) = Chr(X) Then
                isFunky = True
                Exit Function                        'Exit the loop since it found it
            End If
        Next X
    Next i
    

End Function

Private Function CharsIN(sText As String, sChar As String) As Long
    'Wrote this to find the amount
    'of extentions to query in search.
    'Rather useful function too....
    
    Dim iPos As Long, sNext As String
    sNext = sText
    Do
        iPos = InStr(sText, sChar)
        If iPos = 0 Then Exit Function
        sText = Mid(sText, iPos + 1)
        CharsIN = CharsIN + 1
    Loop
End Function

Private Function CharsPOS(sText As String, sChar As String, Optional ByVal iStart As Long = 1) As Long
    'Wrote this to get a position of a char
    'found at a certain amount of times, get the pos
    
    Dim iPos As Long, iCount As Long
    iCount = 1
    Do

        iPos = InStr(iPos + 1, sText, sChar)
        If iPos = 0 Then Exit Function
        If iCount = iStart Then
            CharsPOS = iPos
            Exit Do
        End If
        iCount = iCount + 1
    Loop
End Function

