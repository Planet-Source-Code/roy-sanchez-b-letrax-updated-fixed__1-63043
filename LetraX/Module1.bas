Attribute VB_Name = "Module1"

Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal iparam As Long) As Long


Public Sub formdrag(theform As Form)
    ReleaseCapture
    Call SendMessage(theform.HWnd, &HA1, 2, 0&)
End Sub

Public Function pReplace(strExpression As String, strFind As String, strReplace As String)
    Dim intX As Integer


    If (Len(strExpression) - Len(strFind)) >= 0 Then


        For intX = 1 To Len(strExpression)


            If Mid(strExpression, intX, Len(strFind)) = strFind Then
                strExpression = Left(strExpression, (intX - 1)) + strReplace + Mid(strExpression, intX + Len(strFind), Len(strExpression))
            End If
        Next
    End If
    pReplace = strExpression
End Function


Sub CreateDirPath(sPath As String)
    'This function is like MkDir, but can yo
    '     u can use it to
    'create directories inside of each other


    '     without specifying
        'each one seperately.
        '
        'by Terminal
        'http://www.darksigns.com
        'terminal@darksigns.com
        'You should leave the error on resume ne
        '     xt line in so that if a directory alread
        '     y exists it won't generate an error, it
        '     will keep creating other dirs.
        On Error Resume Next
        sPath = Trim(UCase(sPath))
        Dim tmpS As String, tmpS2 As String
        tmpS = sPath


        If InStr(1, sPath, "\") = 0 Then
            Exit Sub
        End If


        Do
            tmpS2 = tmpS2 & Mid(tmpS, 1, InStr(1, tmpS, "\"))
            tmpS = Mid(tmpS, InStr(1, tmpS, "\") + 1, Len(tmpS))


            If Len(tmpS2) > 3 Then
                MkDir Mid(tmpS2, 1, Len(tmpS2) - 1)
            End If
        Loop Until InStr(1, tmpS, "\") = 0
        tmpS = tmpS2 & tmpS
        MkDir tmpS
    End Sub
