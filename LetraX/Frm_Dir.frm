VERSION 5.00
Object = "{6844BC75-C63D-4869-8B19-EB2C18AF0310}#1.0#0"; "Super_Button.ocx"
Begin VB.Form Frm_Dir 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_Dir.frx":0000
   ScaleHeight     =   1635
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin Super_Button.SuperButton Spb_Dir 
      Height          =   555
      Left            =   4920
      TabIndex        =   0
      ToolTipText     =   "Buscar Directorio"
      Top             =   600
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
      Object.ToolTipText     =   "Buscar Directorio"
      CaptionLeft     =   18
      CaptionTop      =   18
      Picture         =   "Frm_Dir.frx":1EAF
      EdgeType        =   0
   End
   Begin Super_Button.SuperButton Spb_Guardar 
      Height          =   555
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "Guardar"
      Top             =   1080
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
      Picture         =   "Frm_Dir.frx":2789
      EdgeType        =   0
   End
   Begin Super_Button.SuperButton Spb_Salir 
      Height          =   555
      Left            =   3120
      TabIndex        =   2
      ToolTipText     =   "Cerrar"
      Top             =   1080
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
      Object.ToolTipText     =   "Cerrar"
      CaptionLeft     =   18
      CaptionTop      =   18
      Picture         =   "Frm_Dir.frx":3063
      EdgeType        =   0
   End
   Begin VB.Label Lbl_Dir 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   4455
   End
End
Attribute VB_Name = "Frm_Dir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Some Declarations...
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
hWndOwner      As Long
pIDLRoot       As Long
pszDisplayName As Long
lpszTitle      As Long
ulFlags        As Long
lpfnCallback   As Long
lParam         As Long
iImage         As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Dim kbits As Variant

Private Sub DirBox(Msg As String, Directory As String)
On Error Resume Next
    
'Well, i will say change this when you know what do you do :).
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    'Change this to set what info is displayed.
    szTitle = Msg
    With tBrowseInfo
       .hWndOwner = Me.hWnd
       .lpszTitle = lstrcat(szTitle, "")
       .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If (lpIDList) Then
       sBuffer = Space(MAX_PATH)
       SHGetPathFromIDList lpIDList, sBuffer
       sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       Directory = sBuffer
    End If
End Sub

Private Sub Form_Load()
'FormOnTop Me.hWnd, True
Lbl_Dir.Caption = Directorio
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

formdrag Me

End Sub

Private Sub Spb_Dir_Click()
' Here is the title of the DirBox

DirBox "Seleccione el directorio", Id$

' And here is where will be the browser destination displayed

If Id$ <> "" Then
Lbl_Dir.Caption = Id$
Lbl_Dir.Caption = IIf(Right(Lbl_Dir.Caption, 1) = "\", Lbl_Dir.Caption, Lbl_Dir.Caption & "\")
End If


End Sub

Private Sub Spb_Guardar_Click()
Directorio = Lbl_Dir.Caption
X = WriteINI("DIRECTORIO", "RUTA", Directorio, App.Path & "\LetraX.ini")
End Sub

Private Sub Spb_Salir_Click()
FormOnTop Frm_LetraX.hWnd, True
Unload Me
End Sub
