VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{6844BC75-C63D-4869-8B19-EB2C18AF0310}#1.0#0"; "Super_Button.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Frm_LetraX 
   BorderStyle     =   0  'None
   Caption         =   "LetraX"
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Frm_LetraX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_LetraX.frx":08CA
   ScaleHeight     =   7650
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Super_Button.SuperButton SuperButton3 
      Height          =   135
      Left            =   4560
      TabIndex        =   16
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   238
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      MouseIcon       =   "Frm_LetraX.frx":4489
      MousePointer    =   99
      CaptionLeft     =   3
      CaptionTop      =   -5
      Caption         =   "=="
      EdgeType        =   0
   End
   Begin Super_Button.SuperButton SuperButton2 
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   45
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      AutoSize        =   -1  'True
      MouseIcon       =   "Frm_LetraX.frx":47A3
      MousePointer    =   99
      CaptionLeft     =   28
      CaptionTop      =   8
      EdgeType        =   0
   End
   Begin VB.Timer Timer2 
      Left            =   3360
      Top             =   1560
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   480
      Picture         =   "Frm_LetraX.frx":4ABD
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin Super_Button.SuperButton Spb_Acerca 
      Height          =   555
      Left            =   3480
      TabIndex        =   12
      ToolTipText     =   "Acerca de.."
      Top             =   360
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      AutoSize        =   -1  'True
      Object.ToolTipText     =   "Acerca de.."
      CaptionLeft     =   18
      CaptionTop      =   18
      Picture         =   "Frm_LetraX.frx":5387
      EdgeType        =   0
   End
   Begin Super_Button.SuperButton SuperButton1 
      Height          =   255
      Left            =   4080
      TabIndex        =   11
      Top             =   45
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      MouseIcon       =   "Frm_LetraX.frx":5C61
      MousePointer    =   99
      CaptionLeft     =   3
      CaptionTop      =   -2
      Caption         =   "---"
      EdgeType        =   0
   End
   Begin Super_Button.SuperButton Spb_Actualizar 
      Height          =   555
      Left            =   1320
      TabIndex        =   10
      ToolTipText     =   "Actualizar"
      Top             =   360
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      AutoSize        =   -1  'True
      Object.ToolTipText     =   "Actualizar"
      CaptionLeft     =   18
      CaptionTop      =   18
      Picture         =   "Frm_LetraX.frx":5F7B
      EdgeType        =   0
   End
   Begin Super_Button.SuperButton Spb_Config 
      Height          =   555
      Left            =   2640
      TabIndex        =   8
      ToolTipText     =   "Configurar Directorio"
      Top             =   360
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      AutoSize        =   -1  'True
      Object.ToolTipText     =   "Configurar Directorio"
      CaptionLeft     =   18
      CaptionTop      =   18
      Picture         =   "Frm_LetraX.frx":6855
      EdgeType        =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2280
      Top             =   1560
   End
   Begin MSComDlg.CommonDialog CDL1 
      Left            =   2760
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Super_Button.SuperButton Spb_texto 
      Height          =   555
      Left            =   720
      TabIndex        =   1
      ToolTipText     =   "Texto"
      Top             =   360
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      AutoSize        =   -1  'True
      Object.ToolTipText     =   "Texto"
      CaptionLeft     =   18
      CaptionTop      =   18
      Picture         =   "Frm_LetraX.frx":7187
      EdgeType        =   0
   End
   Begin Super_Button.SuperButton Spb_Guardar 
      Height          =   555
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Guardar"
      Top             =   360
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      AutoSize        =   -1  'True
      Object.ToolTipText     =   "Guardar"
      CaptionLeft     =   18
      CaptionTop      =   18
      Picture         =   "Frm_LetraX.frx":7A61
      EdgeType        =   0
   End
   Begin Super_Button.SuperButton Spb_Borrar 
      Height          =   555
      Left            =   2040
      TabIndex        =   3
      ToolTipText     =   "Borrar"
      Top             =   360
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      AutoSize        =   -1  'True
      Object.ToolTipText     =   "Borrar"
      CaptionLeft     =   18
      CaptionTop      =   18
      Picture         =   "Frm_LetraX.frx":833B
      EdgeType        =   0
   End
   Begin Super_Button.SuperButton Spb_Cerrar 
      Height          =   555
      Left            =   4440
      TabIndex        =   4
      ToolTipText     =   "Salir de Programa"
      Top             =   390
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      AutoSize        =   -1  'True
      Object.ToolTipText     =   "Salir de Programa"
      CaptionLeft     =   18
      CaptionTop      =   18
      Picture         =   "Frm_LetraX.frx":8C15
      EdgeType        =   0
   End
   Begin Super_Button.SuperButton Spb_Browser 
      Height          =   555
      Left            =   720
      TabIndex        =   7
      ToolTipText     =   "Browser"
      Top             =   360
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   1
      AutoSize        =   -1  'True
      Object.ToolTipText     =   "Browser"
      CaptionLeft     =   18
      CaptionTop      =   18
      Picture         =   "Frm_LetraX.frx":94EF
      EdgeType        =   0
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5175
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   4815
      ExtentX         =   8493
      ExtentY         =   9128
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   9128
      _Version        =   393217
      BackColor       =   796610
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Frm_LetraX.frx":9DC9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   7320
      Width           =   4815
   End
   Begin VB.Label Lbl_titulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   1200
      TabIndex        =   6
      Top             =   1530
      Width           =   45
   End
   Begin VB.Label Lbl_Artista 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   1200
      TabIndex        =   5
      Top             =   1095
      Width           =   45
   End
End
Attribute VB_Name = "Frm_LetraX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TrayIcon As New clsSysTrayIcon
Attribute TrayIcon.VB_VarHelpID = -1
Dim Artist As String, Song As String
Dim lastsong As String
Dim Fuente As String
Dim Tamaño As Integer
Dim Negrita As Boolean
Dim Italica As Boolean
Dim Trans As Boolean
Dim Subrayado As Boolean
Dim trazado As Boolean
Dim minimo As Integer
Dim Fondo_Color As OLE_COLOR


Private Sub Establecer_formato()


 Directorio = ReadINI("DIRECTORIO", "RUTA", App.Path & "\LetraX.ini")
 If Directorio = "" Then
    Directorio = App.Path & "\LETRAS\"
 End If
 Fuente = ReadINI("FUENTES", "FUENTE", App.Path & "\LetraX.ini")
 Tamaño = ReadINI("FUENTES", "TAMAÑO", App.Path & "\LetraX.ini")
 Negrita = ReadINI("FUENTES", "NEGRITA", App.Path & "\LetraX.ini")
 Italica = ReadINI("FUENTES", "ITALICA", App.Path & "\LetraX.ini")
 Fuente_Color = ReadINI("FUENTES", "FUENTE_COLOR", App.Path & "\LetraX.ini")
 Subrayado = ReadINI("FUENTES", "SUBRAYADO", App.Path & "\LetraX.ini")
 trazado = ReadINI("FUENTES", "TRAZADO", App.Path & "\LetraX.ini")
 Fondo_Color = ReadINI("FONDO", "COLOR", App.Path & "\LetraX.ini")
 
 RichTextBox1.Font.Name = Fuente
 RichTextBox1.Font.Size = Tamaño
 RichTextBox1.Font.Bold = Negrita
 RichTextBox1.Font.Italic = Italica
 RichTextBox1.Font.Underline = Subrayado
 RichTextBox1.Font.Strikethrough = trazado
 RichTextBox1.BackColor = Fondo_Color
 RichTextBox1.SelColor = Fuente_Color
 RichTextBox1.Font.Bold = Negrita
 
 
End Sub

Private Sub Form_Load()
Trans = False
minimo = 2
Establecer_formato
WebBrowser1.Visible = False
FormOnTop Me.hWnd, True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
formdrag Me
End Sub


Sub Cerrar_Programa()
Dim X As Integer
 X = WriteINI("DIRECTORIO", "RUTA", Directorio, App.Path & "\LetraX.ini")
 X = WriteINI("FUENTES", "FUENTE", RichTextBox1.Font.Name, App.Path & "\LetraX.ini")
 X = WriteINI("FUENTES", "TAMAÑO", RichTextBox1.Font.Size, App.Path & "\LetraX.ini")
 X = WriteINI("FUENTES", "NEGRITA", RichTextBox1.Font.Bold, App.Path & "\LetraX.ini")
 X = WriteINI("FUENTES", "ITALICA", RichTextBox1.Font.Italic, App.Path & "\LetraX.ini")
 X = WriteINI("FUENTES", "FUENTE_COLOR", Fuente_Color, App.Path & "\LetraX.ini")
 X = WriteINI("FUENTES", "SUBRAYADO", RichTextBox1.Font.Underline, App.Path & "\LetraX.ini")
 X = WriteINI("FUENTES", "TRAZADO", RichTextBox1.Font.Strikethrough, App.Path & "\LetraX.ini")
 
 X = WriteINI("FONDO", "COLOR", RichTextBox1.BackColor, App.Path & "\LetraX.ini")
 TrayIcon.RemoveIcon Me
End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If TrayIcon.bRunningInTray Then
    Select Case X
        Case 7755
            Call menu_emergente2
        Case 7725
            Me.WindowState = 0
            Me.Show
    End Select
End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then
        Me.Visible = False
    End If
End Sub


Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RichTextBox1.SetFocus
If Button = vbRightButton Then
   
   Call menu_emergente
End If

End Sub

Private Sub Spb_Acerca_Click()
FormOnTop Me.hWnd, False
Frm_Acerca.Show 1
End Sub

Private Sub Spb_Actualizar_Click()
Actualizar
End Sub

Private Sub Spb_Borrar_Click()
Dim DIR As String
Dim gfile
Dim Acepta As Integer
DIR = Directorio & "LETRAS\" & Left(Lbl_Artista.Caption, 1) & "\" & Lbl_Artista.Caption & " - " & Lbl_titulo.Caption & ".txt"
On Error GoTo herror
gfile = GetAttr(DIR)
Acepta = MsgBox("Seguro que desea borrar el archivo?", vbQuestion + vbYesNo, "Borrar")
If Acepta = vbYes Then
   Kill DIR
   Lbl_Artista.Caption = ""
   Lbl_titulo.Caption = ""
   RichTextBox1.Text = ""
End If
Exit Sub

herror:
MsgBox "No existe el archivo, no se puede borrar, Verifique!", vbCritical, "LetraX"
End Sub

Private Sub Spb_Browser_Click()
Spb_Browser.Visible = False
Spb_texto.Visible = True
Spb_Borrar.Visible = False
Spb_Actualizar.Visible = True
WebBrowser1.Visible = True
RichTextBox1.Visible = False
lblStatus.Visible = True

End Sub

Private Sub Spb_Cerrar_Click()
Cerrar_Programa
End Sub

Private Sub Spb_Config_Click()
FormOnTop Me.hWnd, False
Frm_Dir.Show 1
End Sub




Private Sub Spb_Texto_Click()
Spb_Browser.Visible = True
Spb_texto.Visible = False
Spb_Borrar.Visible = True
Spb_Actualizar.Visible = False
WebBrowser1.Visible = False
RichTextBox1.Visible = True
lblStatus.Visible = False

End Sub



Private Sub Spb_Guardar_Click()
Dim DIR As String
Dim a As String
    On Error Resume Next
    
If WebBrowser1.Visible = True Then
    'Select all webpage
    WebBrowser1.SetFocus
    WebBrowser1.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT
'******************************************************
    On Error Resume Next
    'Copy selected text/picture etc...
    WebBrowser1.SetFocus
    
    WebBrowser1.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT
    
    WebBrowser1.ExecWB OLECMDID_CLEARSELECTION, OLECMDEXECOPT_DODEFAULT
    Spb_Texto_Click
    RichTextBox1.Text = ""
    RichTextBox1.SetFocus
    RichTextBox1.SelColor = Fuente_Color
    Screen.ActiveControl.SelText = Clipboard.GetText()
    RichTextBox1.SelStart = 0
    
    
    Clipboard.Clear
End If
    On Error GoTo Herror_Folder
    DIR = Directorio & "LETRAS\" & UCase(Left(Lbl_Artista, 1)) & "\"
    gfolder = GetAttr(DIR)
    On Error Resume Next
    
    DIR = Directorio & "LETRAS\" & Left(Lbl_Artista, 1) & "\" & Lbl_Artista & " - " & Lbl_titulo & ".txt"
    If DIR <> "" Then
    Open DIR For Output As #1
    Print #1, RichTextBox1.Text
    Close 1
    End If
Exit Sub

Herror_Folder:
   CreateDirPath DIR
Resume Next

End Sub


Private Sub SuperButton1_Click()

TrayIcon.ChangeIcon Me, Picture1
TrayIcon.ShowIcon Me
Me.Visible = False
'Frm_LetraX.WindowState = vbMinimized
 
End Sub

Private Sub SuperButton2_Click()
If Trans = False Then
   MakeTransparent Me.hWnd, 100
   Trans = True
Else
   MakeOpaque Me.hWnd
   Trans = False
End If
End Sub

Private Sub SuperButton3_Click()
If minimo = 2 Then
    WebBrowser1.Height = 1500
    RichTextBox1.Height = 1500
    Frm_LetraX.Height = 4000
    lblStatus.Top = 3600
    minimo = 1
Else
    minimo = 2
    WebBrowser1.Height = 5175
    RichTextBox1.Height = 5175
    Frm_LetraX.Height = 7650
    lblStatus.Top = 7320
End If
End Sub

Private Sub Timer1_Timer()
    Dim X
    Dim Direct As String
    Dim ArtTemp As String, SngTemp As String
    Song = ""
    Artist = ""
    Dim Temp As String, k As Integer
    winampSplitTitle ArtTemp, SngTemp
    If lastsong = SngTemp Then Exit Sub
    For k = 1 To Len(ArtTemp): Temp = LCase(Mid$(ArtTemp, k, 1))
    If Temp = " " Or Temp = Chr$(32) Then
        Artist = Artist & "%20"
        Else
    Artist = Artist & Temp
    End If
    Next k
    
    For k = 1 To Len(SngTemp): Temp = LCase(Mid$(SngTemp, k, 1))
    If Temp = " " Or Temp = Chr$(32) Then
    Song = Song & "%20"
        Else
    Song = Song & Temp
    End If
    Next k
    Lbl_Artista = ArtTemp
    Lbl_titulo = SngTemp
    lastsong = SngTemp
    If Buscar_Cancion(ArtTemp, SngTemp) = False Then
       RichTextBox1.Text = ""
       Spb_Browser_Click
       Direct = "http://lyrc.com.ar/tema1es.php?songname=" & Trim(pReplace(Song, "%20", " ")) & "&artist=" & Trim(pReplace(Artist, "%20", " "))
       Direct = pReplace(Direct, "%20", " ")
       WebBrowser1.Navigate Direct
    Else
    Spb_Texto_Click
    End If
End Sub


Private Function Buscar_Cancion(ByVal Artista As String, ByVal titulo As String) As Boolean

On Error GoTo err
Dim DIR As String
Dim t As Long
Dim i As Long
DIR = Directorio & "LETRAS\" & Left(Artista, 1) & "\" & Artista & " - " & titulo & ".txt"
If DIR <> "" Then
i = FreeFile
Open DIR For Input As #i
Buscar_Cancion = True
RichTextBox1.Text = ""
RichTextBox1.SelColor = Fuente_Color
RichTextBox1.SelText = Input(LOF(i), i)
Close #i
RichTextBox1.SelStart = 0
End If
Exit Function



err:
Buscar_Cancion = False
Close #i

Exit Function
        
End Function

Private Sub Actualizar()
    Dim Direct As String
       Direct = "http://lyrc.com.ar/tema1es.php?songname=" & Trim(pReplace(Lbl_titulo.Caption, "%20", " ")) & "&artist=" & Trim(pReplace(Lbl_Artista.Caption, "%20", " "))
       Direct = pReplace(Direct, "%20", " ")
       WebBrowser1.Navigate Direct
    

End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
'Shows progress in status bar
lblStatus.Caption = "Reading " & Progress & "  of  " & ProgressMax
If Progress = 0 And ProgressMax = 0 Then
   lblStatus.Caption = "Done"
End If
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'shows done in the status bar

End Sub
