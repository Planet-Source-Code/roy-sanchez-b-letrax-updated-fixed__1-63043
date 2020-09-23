VERSION 5.00
Object = "{6844BC75-C63D-4869-8B19-EB2C18AF0310}#1.0#0"; "Super_Button.ocx"
Begin VB.Form Frm_Acerca 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   Picture         =   "Frm_Acerca.frx":0000
   ScaleHeight     =   2235
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Super_Button.SuperButton Spb_Cerrar 
      Height          =   555
      Left            =   3960
      TabIndex        =   0
      Top             =   1560
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
      CaptionLeft     =   18
      CaptionTop      =   18
      Picture         =   "Frm_Acerca.frx":3032
      EdgeType        =   0
   End
End
Attribute VB_Name = "Frm_Acerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

formdrag Me

End Sub

Private Sub Spb_Cerrar_Click()
FormOnTop Frm_LetraX.hWnd, True
Unload Me
End Sub
