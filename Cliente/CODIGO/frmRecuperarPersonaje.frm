VERSION 5.00
Begin VB.Form frmRecuperarPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "frmRecuperarPersonaje"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   105
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   337
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   313
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNewPasswd 
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
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1350
      Width           =   3735
   End
   Begin VB.TextBox txtSecCode 
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
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2100
      Width           =   3735
   End
   Begin VB.TextBox txtNombre 
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
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   3735
   End
   Begin VB.Image imgSalir 
      Height          =   375
      Left            =   2400
      MouseIcon       =   "frmRecuperarPersonaje.frx":0000
      MousePointer    =   99  'Custom
      Top             =   4335
      Width           =   1890
   End
   Begin VB.Image imgRecuperar 
      Height          =   375
      Left            =   420
      MouseIcon       =   "frmRecuperarPersonaje.frx":0152
      MousePointer    =   99  'Custom
      Top             =   4335
      Width           =   1890
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   480
      TabIndex        =   2
      Top             =   2850
      Width           =   3735
   End
End
Attribute VB_Name = "frmRecuperarPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Picture = LoadPicture(DirInterfaces & "VentanaRecuperar.jpg")
End Sub

Private Sub imgRecuperar_Click()

If Not CheckData Then Exit Sub

Call WriteRecuperarPersonajes(txtNombre.Text, txtNewPasswd.Text, txtSecCode.Text)

End Sub

Private Function CheckData() As Boolean

    Select Case vbNullString
        Case Is = txtNombre.Text
            MsgBox "Debes ingresar el nombre del personaje."
            Exit Function
        
        Case Is = txtNewPasswd.Text
            MsgBox "Debes ingresar una nueva contraseña para el personaje."
            Exit Function
        
        Case Is = txtSecCode.Text
            MsgBox "Debes ingresar el código de seguridad."
            Exit Function
    End Select

CheckData = True

End Function

Private Sub imgSalir_Click()
    txtNombre.Text = vbNullString
    txtNewPasswd.Text = vbNullString
    txtSecCode.Text = vbNullString
    
    Unload Me
End Sub
