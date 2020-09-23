VERSION 5.00
Object = "{6844BC75-C63D-4869-8B19-EB2C18AF0310}#1.0#0"; "super_button.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Frm_Browser 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_Browser.frx":0000
   ScaleHeight     =   5970
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   5535
      ExtentX         =   9763
      ExtentY         =   8281
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
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   5280
      Top             =   5520
   End
   Begin Super_Button.SuperButton Spb_Cerrar 
      Height          =   555
      Left            =   5160
      TabIndex        =   3
      ToolTipText     =   "Salir de Programa"
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
      Object.ToolTipText     =   "Salir de Programa"
      CaptionLeft     =   18
      CaptionTop      =   18
      Picture         =   "Frm_Browser.frx":26DA
      EdgeType        =   0
   End
   Begin Super_Button.SuperButton SuperButton1 
      Height          =   555
      Left            =   120
      TabIndex        =   4
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
      Picture         =   "Frm_Browser.frx":2FB4
      EdgeType        =   0
   End
   Begin Super_Button.SuperButton Spb_Nuevo 
      Height          =   555
      Left            =   840
      TabIndex        =   5
      ToolTipText     =   "Nuevo"
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
      Object.ToolTipText     =   "Nuevo"
      CaptionLeft     =   18
      CaptionTop      =   18
      Picture         =   "Frm_Browser.frx":388E
      EdgeType        =   0
   End
   Begin VB.Label Lbl_Artista 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      Top             =   1110
      Width           =   45
   End
   Begin VB.Label Lbl_titulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   1530
      Width           =   45
   End
End
Attribute VB_Name = "Frm_Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Artist As String, Song As String
Dim lastsong As String
Private Sub Command1_Click()
     WebBrowser1.Navigate "http://www.absolutelyric.com/?mode=search&q=" & Text1 & "&type=info"

  
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

formdrag Me

End Sub


Private Sub Spb_Cerrar_Click()
Unload Me
End Sub

Private Sub Spb_Nuevo_Click()

'*******************************************************
    On Error Resume Next
    'Select all webpage
    WebBrowser1.SetFocus
    WebBrowser1.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT
'*******************************************************
'******************************************************
    On Error Resume Next
    'Copy selected text/picture etc...
    WebBrowser1.SetFocus
    WebBrowser1.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub SuperButton1_Click()

'*******************************************************
    On Error Resume Next
    'Save current page as
    WebBrowser1.SetFocus
    WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
'*******************************************************
End Sub

Private Sub Timer1_Timer()
    Dim direct As String
    Dim ArtTemp As String, SngTemp As String
    Song = ""
    Artist = ""
    Dim Temp As String, k As Integer
    winampSplitTitle ArtTemp, SngTemp
    If lastsong = SngTemp Then Exit Sub
    Artist = "&artist="
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
    direct = "http://lyrc.com.ar/tema1es.php?songname=" & Trim(pReplace(Song, "%20", " ")) & "&artist=" & Trim(pReplace(Artist, "%20", " "))
    direct = pReplace(direct, "%20", " ")
    WebBrowser1.Navigate direct
End Sub


