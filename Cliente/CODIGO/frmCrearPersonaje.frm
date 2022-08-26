VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   Icon            =   "frmCrearPersonaje.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSecCode 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   3540
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2700
      Width           =   5055
   End
   Begin VB.ComboBox lstAlienacion 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":1CCA
      Left            =   4680
      List            =   "frmCrearPersonaje.frx":1CD4
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   9000
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.TextBox txtMail 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3540
      TabIndex        =   3
      Top             =   2100
      Width           =   5055
   End
   Begin VB.TextBox txtConfirmPasswd 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   6180
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1500
      Width           =   2415
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   3540
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1500
      Width           =   2415
   End
   Begin VB.Timer tAnimacion 
      Left            =   9240
      Top             =   6120
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":1CE7
      Left            =   2940
      List            =   "frmCrearPersonaje.frx":1CE9
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4245
      Width           =   2625
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":1CEB
      Left            =   2940
      List            =   "frmCrearPersonaje.frx":1CF5
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4890
      Width           =   2625
   End
   Begin VB.ComboBox lstRaza 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":1D08
      Left            =   2940
      List            =   "frmCrearPersonaje.frx":1D0A
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3585
      Width           =   2625
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3540
      MaxLength       =   30
      TabIndex        =   0
      Top             =   900
      Width           =   5055
   End
   Begin VB.PictureBox picPJ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   9840
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   14
      Top             =   3960
      Width           =   615
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   1
      Left            =   9555
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   26
      Top             =   3480
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   2
      Left            =   9960
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   27
      Top             =   3480
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   3
      Left            =   10365
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   28
      Top             =   3480
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   4
      Left            =   10770
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   29
      Top             =   3480
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   0
      Left            =   9150
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   25
      Top             =   3480
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   5
      Left            =   8160
      Top             =   5010
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   4
      Left            =   7935
      Top             =   5010
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   3
      Left            =   7710
      Top             =   5010
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   2
      Left            =   7485
      Top             =   5010
      Width           =   225
   End
   Begin VB.Image imgArcoStar 
      Height          =   195
      Index           =   1
      Left            =   7260
      Top             =   5010
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   5
      Left            =   8160
      Top             =   4725
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   4
      Left            =   7935
      Top             =   4725
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   3
      Left            =   7710
      Top             =   4725
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   2
      Left            =   7485
      Top             =   4725
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   5
      Left            =   8160
      Top             =   4440
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   4
      Left            =   7935
      Top             =   4440
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   3
      Left            =   7710
      Top             =   4440
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   2
      Left            =   7485
      Top             =   4440
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   5
      Left            =   8160
      Top             =   4155
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   4
      Left            =   7935
      Top             =   4155
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   3
      Left            =   7710
      Top             =   4155
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   2
      Left            =   7485
      Top             =   4155
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   5
      Left            =   8160
      Top             =   3870
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   4
      Left            =   7935
      Top             =   3870
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   3
      Left            =   7710
      Top             =   3870
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   2
      Left            =   7485
      Top             =   3870
      Width           =   225
   End
   Begin VB.Image imgArmasStar 
      Height          =   195
      Index           =   1
      Left            =   7260
      Top             =   4725
      Width           =   225
   End
   Begin VB.Image imgEscudosStar 
      Height          =   195
      Index           =   1
      Left            =   7260
      Top             =   4440
      Width           =   225
   End
   Begin VB.Image imgVidaStar 
      Height          =   195
      Index           =   1
      Left            =   7260
      Top             =   4155
      Width           =   225
   End
   Begin VB.Image imgMagiaStar 
      Height          =   195
      Index           =   1
      Left            =   7260
      Top             =   3870
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   5
      Left            =   8160
      Top             =   3585
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   4
      Left            =   7935
      Top             =   3585
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   3
      Left            =   7710
      Top             =   3585
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   2
      Left            =   7485
      Top             =   3585
      Width           =   225
   End
   Begin VB.Image imgEvasionStar 
      Height          =   195
      Index           =   1
      Left            =   7260
      Top             =   3585
      Width           =   225
   End
   Begin VB.Label lblEspecialidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   3540
      TabIndex        =   30
      Top             =   6510
      Width           =   5055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   3
      Visible         =   0   'False
      X1              =   663
      X2              =   689
      Y1              =   256
      Y2              =   256
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   2
      Visible         =   0   'False
      X1              =   663
      X2              =   689
      Y1              =   231
      Y2              =   231
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   1
      Visible         =   0   'False
      X1              =   688
      X2              =   688
      Y1              =   232
      Y2              =   256
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   0
      Visible         =   0   'False
      X1              =   663
      X2              =   663
      Y1              =   232
      Y2              =   256
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   2460
      TabIndex        =   24
      Top             =   5310
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   2460
      TabIndex        =   23
      Top             =   4950
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   2460
      TabIndex        =   22
      Top             =   4605
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   2460
      TabIndex        =   21
      Top             =   4260
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   2460
      TabIndex        =   20
      Top             =   3930
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   2010
      TabIndex        =   19
      Top             =   5310
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   2010
      TabIndex        =   18
      Top             =   4950
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   2010
      TabIndex        =   17
      Top             =   4605
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   2010
      TabIndex        =   16
      Top             =   4260
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   2010
      TabIndex        =   15
      Top             =   3930
      Width           =   225
   End
   Begin VB.Image imgVolver 
      Height          =   330
      Left            =   480
      Top             =   8310
      Width           =   1650
   End
   Begin VB.Image imgCrear 
      Height          =   435
      Left            =   8880
      Top             =   8280
      Width           =   1650
   End
   Begin VB.Image DirPJ 
      Height          =   225
      Index           =   1
      Left            =   10215
      Picture         =   "frmCrearPersonaje.frx":1D0C
      Top             =   4920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image DirPJ 
      Height          =   225
      Index           =   0
      Left            =   9840
      Picture         =   "frmCrearPersonaje.frx":201E
      Top             =   4920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image HeadPJ 
      Height          =   225
      Index           =   1
      Left            =   11220
      Picture         =   "frmCrearPersonaje.frx":2330
      Top             =   3525
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image HeadPJ 
      Height          =   225
      Index           =   0
      Left            =   8835
      Picture         =   "frmCrearPersonaje.frx":2642
      Top             =   3525
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   3120
      Left            =   8880
      Stretch         =   -1  'True
      Top             =   9120
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Image imgHogar 
      Height          =   2850
      Left            =   5640
      Picture         =   "frmCrearPersonaje.frx":2954
      Top             =   9120
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   1560
      TabIndex        =   13
      Top             =   4950
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   1560
      TabIndex        =   12
      Top             =   4605
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   1560
      TabIndex        =   11
      Top             =   5310
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   1560
      TabIndex        =   10
      Top             =   4260
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1560
      TabIndex        =   9
      Top             =   3930
      Width           =   225
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cBotonVolver      As clsGraphicalButton

Private cBotonCrear       As clsGraphicalButton

Public LastButtonPressed  As clsGraphicalButton

Private picFullStar       As Picture

Private picHalfStar       As Picture

Private picGlowStar       As Picture

Private vEspecialidades() As String

Private Type tModRaza

        Fuerza As Single
        Agilidad As Single
        Inteligencia As Single
        Carisma As Single
        Constitucion As Single

End Type

Private Type tModClase

        Evasion As Double
        AtaqueArmas As Double
        AtaqueProyectiles As Double
        DañoArmas As Double
        DañoProyectiles As Double
        Escudo As Double
        Magia As Double
        Vida As Double
        Hit As Double

End Type

Private ModRaza()  As tModRaza

Private ModClase() As tModClase

Private NroRazas   As Integer

Private NroClases  As Integer

Private Cargando   As Boolean

Private currentGrh As Long

Private Dir        As E_Heading

Private Sub Form_Load()
        Me.Picture = LoadPicture(DirInterfaces & "VentanaCrearPersonaje.jpg")
    
        Cargando = True
        Call LoadCharInfo
        Call CargarEspecialidades
    
        Call IniciarGraficos
        Call CargarCombos
    
        Call DrawImageInPicture(picPJ, Me.Picture, 0, 0, , , picPJ.Left, picPJ.Top)
        Dir = SOUTH
    
        Cargando = False
    
        'UserClase = 0
        UserSexo = 0
        UserRaza = 0
        UserHogar = 0
        UserEmail = vbNullString
        UserHead = 0
    

End Sub

Private Sub CargarEspecialidades()

        ReDim vEspecialidades(1 To NroClases)
        
        vEspecialidades(eClass.Mage) = "Ataques a distancia con magia."
        vEspecialidades(eClass.Cleric) = "Ataques a distancia con magia y ataques melee a corta distancia."
        vEspecialidades(eClass.Warrior) = "Daño elevado a corta distancia y mucha vida."
        vEspecialidades(eClass.Assasin) = "Apuñalar."
        vEspecialidades(eClass.Thief) = "Caminar y atacar oculto. Robar."
        vEspecialidades(eClass.Bard) = "Evasión alta. Ataques a distancia con magia y posibilidad de utilizar escudos."
        vEspecialidades(eClass.Druid) = "Domar mascotas. Cascos con defensa mágica. Mimetismo."
        vEspecialidades(eClass.Bandit) = "Caminar oculto. Golpes críticos."
        vEspecialidades(eClass.Paladin) = "Daño elevado a corta distancia. Ataques a distancia con magia. Mucha vida."
        vEspecialidades(eClass.Hunter) = "Ocultarse por tiempo indefinido. Daño elevado a distancia con arcos."
        vEspecialidades(eClass.Worker) = "Extracción de recursos y construcción de objetos."
        vEspecialidades(eClass.Pirat) = "Ataques a distancia con cuchillas. Acuchillar. Menos requisitos para navegar."

End Sub

Private Sub IniciarGraficos()

        Dim GrhPath As String

        GrhPath = DirInterfaces
    
        Set cBotonVolver = New clsGraphicalButton
        Set cBotonCrear = New clsGraphicalButton
    
        Set LastButtonPressed = New clsGraphicalButton
              
        Call cBotonVolver.Initialize(imgVolver, GrhPath & "BotonVolverRollover.jpg", GrhPath & "BotonVolverRollover.jpg", _
           GrhPath & "BotonVolverClick.jpg", Me)
                                    
        Call cBotonCrear.Initialize(imgCrear, GrhPath & "BotonCrearPersonaje.jpg", GrhPath & "BotonCrearPersonaje.jpg", _
           GrhPath & "BotonCrearPersonaje.jpg", Me)

        Set picFullStar = LoadPicture(GrhPath & "EstrellaSimple.jpg")
        Set picHalfStar = LoadPicture(GrhPath & "EstrellaMitad.jpg")
        Set picGlowStar = LoadPicture(GrhPath & "EstrellaBrillante.jpg")

End Sub

Private Sub CargarCombos()

        Dim i As Integer
    
        lstProfesion.Clear
    
        For i = LBound(ListaClases) To NroClases
                lstProfesion.AddItem ListaClases(i)
        Next i
    
        lstRaza.Clear
    
        For i = LBound(ListaRazas()) To NroRazas
                lstRaza.AddItem ListaRazas(i)
        Next i
    
        lstProfesion.listIndex = 1

End Sub

Function CheckData() As Boolean

        If txtPasswd.Text <> txtConfirmPasswd.Text Then
            MsgBox "Los passwords que tipeo no coinciden, por favor vuelva a ingresarlos."
            Exit Function
        End If
    
        If Not CheckMailString(txtMail.Text) Then
            MsgBox "Direccion de mail invalida."
            Exit Function
        End If

        If UserRaza = 0 Then
            MsgBox "Seleccione la raza del personaje."
            Exit Function
        End If
    
        If UserSexo = 0 Then
            MsgBox "Seleccione el sexo del personaje."
            Exit Function
        End If
    
        If UserClase = 0 Then
            MsgBox "Seleccione la clase del personaje."
            Exit Function
        End If

        'Toqueteado x Salvito
        Dim i As Long

        For i = 1 To NUMATRIBUTOS
            If Val(lblAtributos(i).Caption) = 0 Then
                MsgBox "Los atributos del personaje son invalidos."
                Exit Function
            End If
        Next i
    
        If Len(Username) > 30 Then
            MsgBox ("El nombre debe tener menos de 30 letras.")
            Exit Function
        End If
        
        If txtSecCode.Text = txtPasswd.Text Then
            MsgBox "Tu código de seguridad no puede ser igual a tu contraseña."
            Exit Function
        End If
        
        If txtSecCode.Text = vbNullString Then
            MsgBox "Debes ingresar tu código de seguridad."
            Exit Function
        End If
        
        If Len(txtSecCode.Text) < 5 Or Len(txtSecCode.Text) > 16 Then
            MsgBox "Tu código de seguridad debe contener un mínimo de 4 caracteres y 16 como máximo."
            Exit Function
        End If
        
        If InStr(txtSecCode, Chr(32)) Then
            MsgBox "No puedes insertar espacios en el código de seguridad."
            Exit Function
        End If
        
        CheckData = True

End Function

Private Sub DirPJ_Click(Index As Integer)

        Select Case Index

                Case 0
                        Dir = CheckDir(Dir + 1)

                Case 1
                        Dir = CheckDir(Dir - 1)

        End Select
    
        Call UpdateHeadSelection

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        ClearLabel

End Sub

Private Sub HeadPJ_Click(Index As Integer)

        Select Case Index

                Case 0
                        UserHead = CheckCabeza(UserHead + 1)

                Case 1
                        UserHead = CheckCabeza(UserHead - 1)

        End Select
    
        Call UpdateHeadSelection
    
End Sub

Private Sub UpdateHeadSelection()

        Dim Head As Integer
    
        Head = UserHead
        Call DrawHead(Head, 2)
    
        Head = Head + 1
        Call DrawHead(CheckCabeza(Head), 3)
    
        Head = Head + 1
        Call DrawHead(CheckCabeza(Head), 4)
    
        Head = UserHead
    
        Head = Head - 1
        Call DrawHead(CheckCabeza(Head), 1)
    
        Head = Head - 1
        Call DrawHead(CheckCabeza(Head), 0)

End Sub

Private Sub imgCrear_Click()

        Dim i         As Integer

        Dim CharAscii As Byte
    
        Username = txtNombre.Text
            
        If Right$(Username, 1) = " " Then
                Username = RTrim$(Username)
                MsgBox "Nombre invalido, se han removido los espacios al final del nombre"

        End If
    
        UserRaza = lstRaza.listIndex + 1
        UserSexo = lstGenero.listIndex + 1
        UserClase = lstProfesion.listIndex + 1
    
        For i = 1 To NUMATRIBUTES
                UserAtributos(i) = Val(lblAtributos(i).Caption)
        Next i
         
        UserHogar = 1
    
        If Not CheckData Then Exit Sub

        UserPassword = txtPasswd.Text
    
        For i = 1 To Len(UserPassword)
                CharAscii = Asc(mid$(UserPassword, i, 1))

                If Not LegalCharacter(CharAscii) Then
                        MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
                        Exit Sub

                End If

        Next i
    
        UserEmail = txtMail.Text
        
        SecurityCode = txtSecCode.Text
        
        For i = 1 To Len(SecurityCode)
                CharAscii = Asc(mid$(SecurityCode, i, 1))

                If Not LegalCharacter(CharAscii) Then
                        MsgBox ("Código de Seguridad inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
                        Exit Sub

                End If

        Next i
       
        frmMain.Socket1.HostName = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
    
        EstadoLogin = E_MODO.CrearNuevoPj

        If Not frmMain.Socket1.Connected Then
        
                MsgBox "Error: Se ha perdido la conexion con el server."
                Unload Me
        
        Else
                Call Login

        End If
       
End Sub

Private Sub imgVolver_Click()
        Call Audio.PlayMIDI("2.mid")

        Unload Me

End Sub

Private Sub lstGenero_Click()
        UserSexo = lstGenero.listIndex + 1
        Call DarCuerpoYCabeza

End Sub

Private Sub lstProfesion_Click()

        On Error Resume Next

        UserClase = lstProfesion.listIndex + 1
    
        Call UpdateStats
        Call UpdateEspecialidad(UserClase)

End Sub

Private Sub UpdateEspecialidad(ByVal eClase As eClass)
        lblEspecialidad.Caption = vEspecialidades(eClase)

End Sub

Private Sub lstRaza_Click()
        UserRaza = lstRaza.listIndex + 1
        Call DarCuerpoYCabeza
    
        Call UpdateStats

End Sub

Private Sub picHead_Click(Index As Integer)

        ' No se mueve si clickea al medio
        If Index = 2 Then Exit Sub
    
        Dim Counter As Integer

        Dim Head    As Integer
    
        Head = UserHead
    
        If Index > 2 Then

                For Counter = Index - 2 To 1 Step -1
                        Head = CheckCabeza(Head + 1)
                Next Counter

        Else

                For Counter = 2 - Index To 1 Step -1
                        Head = CheckCabeza(Head - 1)
                Next Counter

        End If
    
        UserHead = Head
    
        Call UpdateHeadSelection
    
End Sub

Private Sub tAnimacion_Timer()

        Dim SR       As RECT

        Dim Grh      As Long

        Dim x        As Long

        Dim y        As Long

        Static Frame As Byte
    
        If currentGrh = 0 Then Exit Sub
        UserHead = CheckCabeza(UserHead)
    
        Frame = Frame + 1

        If Frame >= GrhData(currentGrh).NumFrames Then Frame = 1
        Call DrawImageInPicture(picPJ, Me.Picture, 0, 0, , , picPJ.Left, picPJ.Top)
    
        Grh = GrhData(currentGrh).Frames(Frame)
    
        With GrhData(Grh)
                SR.Left = .sX
                SR.Top = .sY
                SR.Right = SR.Left + .pixelWidth
                SR.Bottom = SR.Top + .pixelHeight
        
                x = picPJ.Width / 2 - .pixelWidth / 2
                y = (picPJ.Height - .pixelHeight) - 5
        
                Call DrawTransparentGrhtoHdc(picPJ.hdc, x, y, Grh, SR, vbBlack)
                y = y + .pixelHeight

        End With
    
        Grh = HeadData(UserHead).Head(Dir).GrhIndex
    
        With GrhData(Grh)
                SR.Left = .sX
                SR.Top = .sY
                SR.Right = SR.Left + .pixelWidth
                SR.Bottom = SR.Top + .pixelHeight
        
                x = picPJ.Width / 2 - .pixelWidth / 2
                y = y + BodyData(UserBody).HeadOffset.y - .pixelHeight
        
                Call DrawTransparentGrhtoHdc(picPJ.hdc, x, y, Grh, SR, vbBlack)

        End With

End Sub

Private Sub DrawHead(ByVal Head As Integer, ByVal PicIndex As Integer)

        Dim SR  As RECT

        Dim Grh As Long

        Dim x   As Long

        Dim y   As Long
    
        Call DrawImageInPicture(picHead(PicIndex), Me.Picture, 0, 0, , , picHead(PicIndex).Left, picHead(PicIndex).Top)
    
        Grh = HeadData(Head).Head(Dir).GrhIndex

        With GrhData(Grh)
                SR.Left = .sX
                SR.Top = .sY
                SR.Right = SR.Left + .pixelWidth
                SR.Bottom = SR.Top + .pixelHeight
        
                x = picHead(PicIndex).Width / 2 - .pixelWidth / 2
                y = 1
        
                Call DrawTransparentGrhtoHdc(picHead(PicIndex).hdc, x, y, Grh, SR, vbBlack)

        End With
    
End Sub

Private Sub txtNombre_Change()
        txtNombre.Text = LTrim$(txtNombre.Text)

End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub DarCuerpoYCabeza()

        Dim bVisible  As Boolean

        Dim PicIndex  As Integer

        Dim LineIndex As Integer
    
        Select Case UserSexo

                Case eGenero.Hombre

                        Select Case UserRaza

                                Case eRaza.Humano
                                        UserHead = HUMANO_H_PRIMER_CABEZA
                                        UserBody = HUMANO_H_CUERPO_DESNUDO
                    
                                Case eRaza.Elfo
                                        UserHead = ELFO_H_PRIMER_CABEZA
                                        UserBody = ELFO_H_CUERPO_DESNUDO
                    
                                Case eRaza.ElfoOscuro
                                        UserHead = DROW_H_PRIMER_CABEZA
                                        UserBody = DROW_H_CUERPO_DESNUDO
                    
                                Case eRaza.Enano
                                        UserHead = ENANO_H_PRIMER_CABEZA
                                        UserBody = ENANO_H_CUERPO_DESNUDO
                    
                                Case eRaza.Gnomo
                                        UserHead = GNOMO_H_PRIMER_CABEZA
                                        UserBody = GNOMO_H_CUERPO_DESNUDO
                    
                                Case Else
                                        UserHead = 0
                                        UserBody = 0

                        End Select
            
                Case eGenero.Mujer

                        Select Case UserRaza

                                Case eRaza.Humano
                                        UserHead = HUMANO_M_PRIMER_CABEZA
                                        UserBody = HUMANO_M_CUERPO_DESNUDO
                    
                                Case eRaza.Elfo
                                        UserHead = ELFO_M_PRIMER_CABEZA
                                        UserBody = ELFO_M_CUERPO_DESNUDO
                    
                                Case eRaza.ElfoOscuro
                                        UserHead = DROW_M_PRIMER_CABEZA
                                        UserBody = DROW_M_CUERPO_DESNUDO
                    
                                Case eRaza.Enano
                                        UserHead = ENANO_M_PRIMER_CABEZA
                                        UserBody = ENANO_M_CUERPO_DESNUDO
                    
                                Case eRaza.Gnomo
                                        UserHead = GNOMO_M_PRIMER_CABEZA
                                        UserBody = GNOMO_M_CUERPO_DESNUDO
                    
                                Case Else
                                        UserHead = 0
                                        UserBody = 0

                        End Select

                Case Else
                        UserHead = 0
                        UserBody = 0

        End Select
    
        bVisible = UserHead <> 0 And UserBody <> 0
    
        HeadPJ(0).Visible = bVisible
        HeadPJ(1).Visible = bVisible
        
        DirPJ(0).Visible = bVisible
        DirPJ(1).Visible = bVisible
    
        For PicIndex = 0 To 4
                picHead(PicIndex).Visible = bVisible
        Next PicIndex
    
        For LineIndex = 0 To 3
                Line1(LineIndex).Visible = bVisible
        Next LineIndex
    
        If bVisible Then Call UpdateHeadSelection
    
        currentGrh = BodyData(UserBody).Walk(Dir).GrhIndex

        If currentGrh > 0 Then _
           tAnimacion.Interval = Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames)

End Sub

Private Function CheckCabeza(ByVal Head As Integer) As Integer

        Select Case UserSexo

                Case eGenero.Hombre

                        Select Case UserRaza

                                Case eRaza.Humano

                                        If Head > HUMANO_H_ULTIMA_CABEZA Then
                                                CheckCabeza = HUMANO_H_PRIMER_CABEZA + (Head - HUMANO_H_ULTIMA_CABEZA) - 1
                                        ElseIf Head < HUMANO_H_PRIMER_CABEZA Then
                                                CheckCabeza = HUMANO_H_ULTIMA_CABEZA - (HUMANO_H_PRIMER_CABEZA - Head) + 1
                                        Else
                                                CheckCabeza = Head

                                        End If
                
                                Case eRaza.Elfo

                                        If Head > ELFO_H_ULTIMA_CABEZA Then
                                                CheckCabeza = ELFO_H_PRIMER_CABEZA + (Head - ELFO_H_ULTIMA_CABEZA) - 1
                                        ElseIf Head < ELFO_H_PRIMER_CABEZA Then
                                                CheckCabeza = ELFO_H_ULTIMA_CABEZA - (ELFO_H_PRIMER_CABEZA - Head) + 1
                                        Else
                                                CheckCabeza = Head

                                        End If
                
                                Case eRaza.ElfoOscuro

                                        If Head > DROW_H_ULTIMA_CABEZA Then
                                                CheckCabeza = DROW_H_PRIMER_CABEZA + (Head - DROW_H_ULTIMA_CABEZA) - 1
                                        ElseIf Head < DROW_H_PRIMER_CABEZA Then
                                                CheckCabeza = DROW_H_ULTIMA_CABEZA - (DROW_H_PRIMER_CABEZA - Head) + 1
                                        Else
                                                CheckCabeza = Head

                                        End If
                
                                Case eRaza.Enano

                                        If Head > ENANO_H_ULTIMA_CABEZA Then
                                                CheckCabeza = ENANO_H_PRIMER_CABEZA + (Head - ENANO_H_ULTIMA_CABEZA) - 1
                                        ElseIf Head < ENANO_H_PRIMER_CABEZA Then
                                                CheckCabeza = ENANO_H_ULTIMA_CABEZA - (ENANO_H_PRIMER_CABEZA - Head) + 1
                                        Else
                                                CheckCabeza = Head

                                        End If
                
                                Case eRaza.Gnomo

                                        If Head > GNOMO_H_ULTIMA_CABEZA Then
                                                CheckCabeza = GNOMO_H_PRIMER_CABEZA + (Head - GNOMO_H_ULTIMA_CABEZA) - 1
                                        ElseIf Head < GNOMO_H_PRIMER_CABEZA Then
                                                CheckCabeza = GNOMO_H_ULTIMA_CABEZA - (GNOMO_H_PRIMER_CABEZA - Head) + 1
                                        Else
                                                CheckCabeza = Head

                                        End If
                
                                Case Else
                                        UserRaza = lstRaza.listIndex + 1
                                        CheckCabeza = CheckCabeza(Head)

                        End Select
        
                Case eGenero.Mujer

                        Select Case UserRaza

                                Case eRaza.Humano

                                        If Head > HUMANO_M_ULTIMA_CABEZA Then
                                                CheckCabeza = HUMANO_M_PRIMER_CABEZA + (Head - HUMANO_M_ULTIMA_CABEZA) - 1
                                        ElseIf Head < HUMANO_M_PRIMER_CABEZA Then
                                                CheckCabeza = HUMANO_M_ULTIMA_CABEZA - (HUMANO_M_PRIMER_CABEZA - Head) + 1
                                        Else
                                                CheckCabeza = Head

                                        End If
                
                                Case eRaza.Elfo

                                        If Head > ELFO_M_ULTIMA_CABEZA Then
                                                CheckCabeza = ELFO_M_PRIMER_CABEZA + (Head - ELFO_M_ULTIMA_CABEZA) - 1
                                        ElseIf Head < ELFO_M_PRIMER_CABEZA Then
                                                CheckCabeza = ELFO_M_ULTIMA_CABEZA - (ELFO_M_PRIMER_CABEZA - Head) + 1
                                        Else
                                                CheckCabeza = Head

                                        End If
                
                                Case eRaza.ElfoOscuro

                                        If Head > DROW_M_ULTIMA_CABEZA Then
                                                CheckCabeza = DROW_M_PRIMER_CABEZA + (Head - DROW_M_ULTIMA_CABEZA) - 1
                                        ElseIf Head < DROW_M_PRIMER_CABEZA Then
                                                CheckCabeza = DROW_M_ULTIMA_CABEZA - (DROW_M_PRIMER_CABEZA - Head) + 1
                                        Else
                                                CheckCabeza = Head

                                        End If
                
                                Case eRaza.Enano

                                        If Head > ENANO_M_ULTIMA_CABEZA Then
                                                CheckCabeza = ENANO_M_PRIMER_CABEZA + (Head - ENANO_M_ULTIMA_CABEZA) - 1
                                        ElseIf Head < ENANO_M_PRIMER_CABEZA Then
                                                CheckCabeza = ENANO_M_ULTIMA_CABEZA - (ENANO_M_PRIMER_CABEZA - Head) + 1
                                        Else
                                                CheckCabeza = Head

                                        End If
                
                                Case eRaza.Gnomo

                                        If Head > GNOMO_M_ULTIMA_CABEZA Then
                                                CheckCabeza = GNOMO_M_PRIMER_CABEZA + (Head - GNOMO_M_ULTIMA_CABEZA) - 1
                                        ElseIf Head < GNOMO_M_PRIMER_CABEZA Then
                                                CheckCabeza = GNOMO_M_ULTIMA_CABEZA - (GNOMO_M_PRIMER_CABEZA - Head) + 1
                                        Else
                                                CheckCabeza = Head

                                        End If
                
                                Case Else
                                        UserRaza = lstRaza.listIndex + 1
                                        CheckCabeza = CheckCabeza(Head)

                        End Select

                Case Else
                        UserSexo = lstGenero.listIndex + 1
                        CheckCabeza = CheckCabeza(Head)

        End Select

End Function

Private Function CheckDir(ByRef Dir As E_Heading) As E_Heading

        If Dir > E_Heading.WEST Then Dir = E_Heading.NORTH
        If Dir < E_Heading.NORTH Then Dir = E_Heading.WEST
    
        CheckDir = Dir
    
        currentGrh = BodyData(UserBody).Walk(Dir).GrhIndex

        If currentGrh > 0 Then _
           tAnimacion.Interval = Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames)

End Function

Private Sub ClearLabel()
        LastButtonPressed.ToggleToNormal

End Sub

Public Sub UpdateStats()
        Call UpdateRazaMod
        Call UpdateStars

End Sub

Private Sub UpdateRazaMod()

        Dim SelRaza As Integer

        Dim i       As Integer
    
        If lstRaza.listIndex > -1 Then
    
                SelRaza = lstRaza.listIndex + 1
        
                With ModRaza(SelRaza)
                        lblModRaza(eAtributos.Fuerza).Caption = IIf(.Fuerza >= 0, "+", "") & .Fuerza
                        lblModRaza(eAtributos.Agilidad).Caption = IIf(.Agilidad >= 0, "+", "") & .Agilidad
                        lblModRaza(eAtributos.Inteligencia).Caption = IIf(.Inteligencia >= 0, "+", "") & .Inteligencia
                        lblModRaza(eAtributos.Carisma).Caption = IIf(.Carisma >= 0, "+", "") & .Carisma
                        lblModRaza(eAtributos.Constitucion).Caption = IIf(.Constitucion >= 0, "+", "") & .Constitucion

                End With

        End If
    
        ' Atributo total
        For i = 1 To NUMATRIBUTES
                lblAtributoFinal(i).Caption = Val(lblAtributos(i).Caption) + Val(lblModRaza(i))
        Next i
    
End Sub

Private Sub UpdateStars()

        Dim NumStars As Double
    
        If UserClase = 0 Then Exit Sub
    
        ' Estrellas de evasion
        NumStars = (2.454 + 0.073 * Val(lblAtributoFinal(eAtributos.Agilidad).Caption)) * ModClase(UserClase).Evasion
        Call SetStars(imgEvasionStar, NumStars * 2)
    
        ' Estrellas de magia
        NumStars = ModClase(UserClase).Magia * Val(lblAtributoFinal(eAtributos.Inteligencia).Caption) * 0.085
        Call SetStars(imgMagiaStar, NumStars * 2)
    
        ' Estrellas de vida
        NumStars = 0.24 + (Val(lblAtributoFinal(eAtributos.Constitucion).Caption) * 0.5 - ModClase(UserClase).Vida) * 0.475
        Call SetStars(imgVidaStar, NumStars * 2)
    
        ' Estrellas de escudo
        NumStars = 4 * ModClase(UserClase).Escudo
        Call SetStars(imgEscudosStar, NumStars * 2)
    
        ' Estrellas de armas
        NumStars = (0.509 + 0.01185 * Val(lblAtributoFinal(eAtributos.Fuerza).Caption)) * ModClase(UserClase).Hit * _
           ModClase(UserClase).DañoArmas + 0.119 * ModClase(UserClase).AtaqueArmas * _
           Val(lblAtributoFinal(eAtributos.Agilidad).Caption)
        Call SetStars(imgArmasStar, NumStars * 2)
    
        ' Estrellas de arcos
        NumStars = (0.4915 + 0.01265 * Val(lblAtributoFinal(eAtributos.Fuerza).Caption)) * _
           ModClase(UserClase).DañoProyectiles * ModClase(UserClase).Hit + 0.119 * ModClase(UserClase).AtaqueProyectiles * _
           Val(lblAtributoFinal(eAtributos.Agilidad).Caption)
        Call SetStars(imgArcoStar, NumStars * 2)

End Sub

Private Sub SetStars(ByRef ImgContainer As Object, ByVal NumStars As Integer)

        Dim FullStars   As Integer

        Dim HasHalfStar As Boolean

        Dim Index       As Integer

        Dim Counter     As Integer

        If NumStars > 0 Then
        
                If NumStars > 10 Then NumStars = 10
        
                FullStars = Int(NumStars / 2)
        
                ' Tienen brillo extra si estan todas
                If FullStars = 5 Then

                        For Index = 1 To FullStars
                                ImgContainer(Index).Picture = picGlowStar
                        Next Index

                Else

                        ' Numero impar? Entonces hay que poner "media estrella"
                        If (NumStars Mod 2) > 0 Then HasHalfStar = True
            
                        ' Muestro las estrellas enteras
                        If FullStars > 0 Then

                                For Index = 1 To FullStars
                                        ImgContainer(Index).Picture = picFullStar
                                Next Index
                
                                Counter = FullStars

                        End If
            
                        ' Muestro la mitad de la estrella (si tiene)
                        If HasHalfStar Then
                                Counter = Counter + 1
                
                                ImgContainer(Counter).Picture = picHalfStar

                        End If
            
                        ' Si estan completos los espacios, no borro nada
                        If Counter <> 5 Then

                                ' Limpio las que queden vacias
                                For Index = Counter + 1 To 5
                                        Set ImgContainer(Index).Picture = Nothing
                                Next Index

                        End If
            
                End If

        Else

                ' Limpio todo
                For Index = 1 To 5
                        Set ImgContainer(Index).Picture = Nothing
                Next Index

        End If

End Sub

Private Sub LoadCharInfo()

        Dim SearchVar As String

        Dim i         As Long
    
        NroRazas = UBound(ListaRazas())
        NroClases = UBound(ListaClases())

        ReDim ModRaza(1 To NroRazas)
        ReDim ModClase(1 To NroClases)
    
        'Modificadores de Clase
        For i = 1 To NroClases

                With ModClase(i)
                        SearchVar = ListaClases(i)
            
                        .Evasion = Val(GetVar(DirInit & "CharInfo.dat", "MODEVASION", SearchVar))
                        .AtaqueArmas = Val(GetVar(DirInit & "CharInfo.dat", "MODATAQUEARMAS", SearchVar))
                        .AtaqueProyectiles = Val(GetVar(DirInit & "CharInfo.dat", "MODATAQUEPROYECTILES", SearchVar))
                        .DañoArmas = Val(GetVar(DirInit & "CharInfo.dat", "MODDAÑOARMAS", SearchVar))
                        .DañoProyectiles = Val(GetVar(DirInit & "CharInfo.dat", "MODDAÑOPROYECTILES", SearchVar))
                        .Escudo = Val(GetVar(DirInit & "CharInfo.dat", "MODESCUDO", SearchVar))
                        .Hit = Val(GetVar(DirInit & "CharInfo.dat", "HIT", SearchVar))
                        .Magia = Val(GetVar(DirInit & "CharInfo.dat", "MODMAGIA", SearchVar))
                        .Vida = Val(GetVar(DirInit & "CharInfo.dat", "MODVIDA", SearchVar))

                End With

        Next i
    
        'Modificadores de Raza
        For i = 1 To NroRazas

                With ModRaza(i)
                        SearchVar = Replace$(ListaRazas(i), " ", vbNullString)
        
                        .Fuerza = Val(GetVar(DirInit & "CharInfo.dat", "MODRAZA", SearchVar + "Fuerza"))
                        .Agilidad = Val(GetVar(DirInit & "CharInfo.dat", "MODRAZA", SearchVar + "Agilidad"))
                        .Inteligencia = Val(GetVar(DirInit & "CharInfo.dat", "MODRAZA", SearchVar + "Inteligencia"))
                        .Carisma = Val(GetVar(DirInit & "CharInfo.dat", "MODRAZA", SearchVar + "Carisma"))
                        .Constitucion = Val(GetVar(DirInit & "CharInfo.dat", "MODRAZA", SearchVar + "Constitucion"))

                End With

        Next i

End Sub

Private Sub txtSecCode_GotFocus()
    MsgBox "ATENCIÓN: El Código de Seguridad es ÚNICO y no puede ser modificado. Nunca revele su Código de Seguridad A NADIE!, el Staff de Hispano AO NUNCA le solicitará el mismo ya que no lo necesita."
End Sub
