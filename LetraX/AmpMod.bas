Attribute VB_Name = "winampApi"
Option Explicit

Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

'Public Function getWinampWindow(lngHwnd As Long) As String
''This function will get the song current ' ly playing on Winamp 3
'Dim lRet As Long
'Dim sTitle As String
'
'lRet = GetWindowText(lngHwnd, sTitle, Len(sTitle))
'MsgBox Left(sTitle, InStr(1, sTitle, vbNullChar, vbTextCompare) - 10)
'End Function






Public Function strGettext(lngHwnd As Long) As String
    Dim ilngLength As Long
    Dim strBuffer As String
    
        ilngLength = SendMessageLong(lngHwnd, &HE, 0, 0)
        strBuffer = String(ilngLength, 0)
        Call SendMessageByString(lngHwnd, &HD, ilngLength + 1, strBuffer)
        strGettext = strBuffer
        
    
End Function

Public Function lngWinamp() As Long
    'lngWinamp = FindWindow("STUDIO", vbNullString)
    'lngWinamp = FindWindow("Winamp v1.x", vbNullString)
    If FindWindow("STUDIO", vbNullString) > 0 Then
       lngWinamp = FindWindow("STUDIO", vbNullString)
       Winamp = 3
    End If
    If FindWindow("Winamp v1.x", vbNullString) > 0 Then
       lngWinamp = FindWindow("Winamp v1.x", vbNullString)
       Winamp = 2
    End If
End Function

Public Function strWinampTitle() As String
    Dim strCaption As String
    
    strCaption = strGettext(lngWinamp)
    strCaption = Mid(strCaption, InStr(strCaption, ".") + 2, InStr(strCaption, " - Winamp") - 4)
    If Right(strCaption, 1) = " " Then strCaption = Mid(strCaption, 1, Len(strCaption) - 1)
    strWinampTitle = strCaption
End Function

Public Function winampSplitTitle(Optional strArtist As String, Optional strSong As String, Optional lngTrack As Long) As String
    
    Dim strGetTitle As String
    Dim hndError
    
    On Error GoTo hndError
  
        strGetTitle = strGettext(lngWinamp)
        winampSplitTitle = Mid(strGetTitle, InStr(strGetTitle, ".") + 2, Len(strGetTitle))
        If Winamp = 2 Then
           winampSplitTitle = pReplace(winampSplitTitle, " - Winamp", "")
        Else
           winampSplitTitle = Trim(pReplace(Trim(winampSplitTitle), "(playing)", ""))
           winampSplitTitle = Trim(pReplace(Trim(winampSplitTitle), " (stopped)", ""))
           winampSplitTitle = Trim(pReplace(Trim(winampSplitTitle), " (paused)", ""))
           winampSplitTitle = Trim(Mid(winampSplitTitle, 1, Len(winampSplitTitle) - 7))
        End If
        If Right(winampSplitTitle, 1) = " " Then winampSplitTitle = Mid(winampSplitTitle, 1, Len(winampSplitTitle) - 1)
        If InStr(strGetTitle, " - ") Then
            lngTrack = Mid(strGetTitle, 1, InStr(strGetTitle, ".") - 1)
            strGetTitle = Mid(strGetTitle, InStr(strGetTitle, ".") + 2, Len(strGetTitle))
            If Winamp = 2 Then
            strGetTitle = pReplace(strGetTitle, " - Winamp", "")
            Else
           strGetTitle = Trim(pReplace(Trim(strGetTitle), "(playing)", ""))
           strGetTitle = Trim(pReplace(Trim(strGetTitle), " (stopped)", ""))
           strGetTitle = Trim(pReplace(Trim(strGetTitle), " (paused)", ""))
           strGetTitle = Trim(Mid(strGetTitle, 1, Len(strGetTitle) - 7))

           End If
            If Right(strGetTitle, 1) = " " Then strGetTitle = Mid(strGetTitle, 1, Len(strGetTitle) - 1)
            strArtist = Trim(Mid(strGetTitle, 1, InStr(strGetTitle, " -") - 1))
            strSong = Trim(Mid(strGetTitle, InStr(strGetTitle, "- ") + 2))
        End If
     



hndError:
End Function

Public Sub winampGetFileInfo(Optional strFilepath As String, Optional strTitle As String, Optional strArtist As String, Optional strAlbum As String, Optional strYear As String, Optional strComment As String)
 
    Dim lngMpegFileInfo As Long
    Dim lngEdit1 As Long
    Dim lngEdit2 As Long
    Dim lngEdit3 As Long
    Dim lngEdit4 As Long
    Dim lngEdit5 As Long
    Dim lngEdit6 As Long
    
    If lngWinamp = 0 Then Exit Sub
    PostMessage lngWinamp, 273, 40188, 0
    Do
        DoEvents
        lngMpegFileInfo = FindWindow("#32770", "MPEG file info box + ID3 tag editor")
        lngEdit1 = FindWindowEx(lngMpegFileInfo, 0, "Edit", vbNullString)
        lngEdit2 = FindWindowEx(lngMpegFileInfo, lngEdit1, "Edit", vbNullString)
        lngEdit3 = FindWindowEx(lngMpegFileInfo, lngEdit2, "Edit", vbNullString)
        lngEdit4 = FindWindowEx(lngMpegFileInfo, lngEdit3, "Edit", vbNullString)
        lngEdit5 = FindWindowEx(lngMpegFileInfo, lngEdit4, "Edit", vbNullString)
        lngEdit6 = FindWindowEx(lngMpegFileInfo, lngEdit5, "Edit", vbNullString)
    Loop Until lngEdit6 <> 0
    strFilepath = strGettext(lngEdit6)
    strTitle = strGettext(lngEdit1)
    strArtist = strGettext(lngEdit2)
    strAlbum = strGettext(lngEdit3)
    strYear = strGettext(lngEdit4)
    strComment = strGettext(lngEdit5)
    DoEvents
    PostMessage lngMpegFileInfo, &H10, 0, 0
End Sub

