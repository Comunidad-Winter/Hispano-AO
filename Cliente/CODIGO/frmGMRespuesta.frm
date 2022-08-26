VERSION 5.00
Begin VB.Form frmGMRespuesta 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Respuesta GM"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   202
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtRespuestaGM 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1620
      Left            =   780
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   525
      Width           =   6090
   End
   Begin VB.Image imgCerrar 
      Height          =   375
      Left            =   3180
      MouseIcon       =   "frmGMRespuesta.frx":0000
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   1335
   End
End
Attribute VB_Name = "frmGMRespuesta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Picture = LoadPicture(DirInterfaces & "VentanaRespuestaGM.jpg")
End Sub

Private Sub imgCerrar_Click()
    frmMain.lblRespuestaGM.Caption = "0"
    Unload Me
End Sub
