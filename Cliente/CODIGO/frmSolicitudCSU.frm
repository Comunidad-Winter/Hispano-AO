VERSION 5.00
Begin VB.Form frmSolicitudCSU 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2805
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   187
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCSU 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Image imgCerrar 
      Height          =   375
      Left            =   240
      MouseIcon       =   "frmSolicitudCSU.frx":0000
      MousePointer    =   99  'Custom
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Image imgEnviar 
      Height          =   375
      Left            =   240
      MouseIcon       =   "frmSolicitudCSU.frx":0152
      MousePointer    =   99  'Custom
      Top             =   2160
      Width           =   2295
   End
End
Attribute VB_Name = "frmSolicitudCSU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Picture = LoadPicture(DirInterfaces & "VentanaSolicitarCSU.jpg")
End Sub

Private Sub imgCerrar_Click()
    Call WriteSendCSU(vbNullString)
    Unload Me
End Sub

Private Sub imgEnviar_Click()
    Call WriteSendCSU(txtCSU.Text)
    Unload Me
End Sub
