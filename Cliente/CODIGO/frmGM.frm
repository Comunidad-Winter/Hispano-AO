VERSION 5.00
Begin VB.Form frmGM 
   BorderStyle     =   0  'None
   Caption         =   "RequestGM"
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
   Begin VB.TextBox txtConsulta 
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
   Begin VB.Image imgEnviar 
      Height          =   375
      Left            =   6060
      MouseIcon       =   "frmGM.frx":0000
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Image imgLimpiar 
      Height          =   375
      Left            =   4380
      MouseIcon       =   "frmGM.frx":0152
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Image imgCerrar 
      Height          =   375
      Left            =   240
      MouseIcon       =   "frmGM.frx":02A4
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   1335
   End
End
Attribute VB_Name = "frmGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Picture = LoadPicture(DirInterfaces & "VentanaGM.jpg")
End Sub

Private Sub imgCerrar_Click()
    Unload frmGM
End Sub

Private Sub imgEnviar_Click()
        
        If Not LenB(txtConsulta.Text) = 0 Then
                Call WriteGMRequest(txtConsulta.Text)
        End If
    
        Unload frmGM
End Sub

Private Sub imgLimpiar_Click()
    txtConsulta.Text = ""
End Sub
