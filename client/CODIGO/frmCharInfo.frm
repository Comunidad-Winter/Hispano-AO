VERSION 5.00
Begin VB.Form frmCharInfo 
   BorderStyle     =   0  'None
   Caption         =   "Informaci�n del personaje"
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6390
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCharInfo.frx":0000
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   426
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPeticiones 
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
      Height          =   1080
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3210
      Width           =   5730
   End
   Begin VB.TextBox txtMiembro 
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
      Height          =   1080
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4695
      Width           =   5730
   End
   Begin VB.Image imgAceptar 
      Height          =   510
      Left            =   5160
      Tag             =   "1"
      Top             =   5955
      Width           =   1020
   End
   Begin VB.Image imgRechazar 
      Height          =   510
      Left            =   3840
      Tag             =   "1"
      Top             =   5955
      Width           =   1020
   End
   Begin VB.Image imgPeticion 
      Height          =   510
      Left            =   2640
      Tag             =   "1"
      Top             =   5955
      Width           =   1020
   End
   Begin VB.Image imgEchar 
      Height          =   510
      Left            =   1440
      Tag             =   "1"
      Top             =   5955
      Width           =   1020
   End
   Begin VB.Image imgCerrar 
      Height          =   510
      Left            =   120
      Tag             =   "1"
      Top             =   5955
      Width           =   1020
   End
   Begin VB.Label status 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label Nombre 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   700
      Width           =   1440
   End
   Begin VB.Label Nivel 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   1750
      Width           =   1185
   End
   Begin VB.Label Clase 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   1225
      Width           =   1575
   End
   Begin VB.Label Raza 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   960
      Width           =   1560
   End
   Begin VB.Label Genero 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   1500
      Width           =   1335
   End
   Begin VB.Label Oro 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   2010
      Width           =   1365
   End
   Begin VB.Label Banco 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   2250
      Width           =   1425
   End
   Begin VB.Label guildactual 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   960
      Width           =   2265
   End
   Begin VB.Label ejercito 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   1230
      Width           =   1785
   End
   Begin VB.Label Ciudadanos 
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Left            =   4800
      TabIndex        =   4
      Top             =   1500
      Width           =   1185
   End
   Begin VB.Label criminales 
      BackStyle       =   0  'Transparent
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
      Height          =   225
      Left            =   4800
      TabIndex        =   3
      Top             =   1770
      Width           =   1185
   End
   Begin VB.Label reputacion 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   1995
      Width           =   1185
   End
End
Attribute VB_Name = "frmCharInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario    As clsFormMovementManager

Private cBotonCerrar     As clsGraphicalButton

Private cBotonPeticion   As clsGraphicalButton

Private cBotonRechazar   As clsGraphicalButton

Private cBotonEchar      As clsGraphicalButton

Private cBotonAceptar    As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Public Enum CharInfoFrmType

        frmMembers
        frmMembershipRequests

End Enum

Public frmType As CharInfoFrmType

Private Sub Form_Load()
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
    
        Me.Picture = LoadPicture(DirInterfaces & "VentanaInfoPj.jpg")
    
        Call LoadButtons
    
End Sub

Private Sub LoadButtons()

        Dim GrhPath As String
    
        GrhPath = DirInterfaces

        Set cBotonCerrar = New clsGraphicalButton
        Set cBotonPeticion = New clsGraphicalButton
        Set cBotonRechazar = New clsGraphicalButton
        Set cBotonEchar = New clsGraphicalButton
        Set cBotonAceptar = New clsGraphicalButton
    
        Set LastButtonPressed = New clsGraphicalButton
    
        Call cBotonCerrar.Initialize(ImgCerrar, GrhPath & "BotonCerrarInfoChar.jpg", _
           vbNullString, _
           GrhPath & "BotonCerrarClickInfoChar.jpg", Me)

        Call cBotonPeticion.Initialize(imgPeticion, GrhPath & "BotonPeticion.jpg", _
           vbNullString, _
           GrhPath & "BotonPeticionClick.jpg", Me)

        Call cBotonRechazar.Initialize(imgRechazar, GrhPath & "BotonRechazar.jpg", _
           vbNullString, _
           GrhPath & "BotonRechazarClick.jpg", Me)

        Call cBotonEchar.Initialize(imgEchar, GrhPath & "BotonEchar.jpg", _
           vbNullString, _
           GrhPath & "BotonEcharClick.jpg", Me)
                                    
        Call cBotonAceptar.Initialize(imgAceptar, GrhPath & "BotonAceptarInfoChar.jpg", _
           vbNullString, _
           GrhPath & "BotonAceptarClickInfoChar.jpg", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        LastButtonPressed.ToggleToNormal

End Sub

Private Sub imgAceptar_Click()
        Call WriteGuildAcceptNewMember(Nombre)
        Unload frmGuildLeader
        Call WriteRequestGuildLeaderInfo
        Unload Me

End Sub

Private Sub ImgCerrar_Click()
        Unload Me

End Sub

Private Sub imgEchar_Click()
        Call WriteGuildKickMember(Nombre)
        Unload frmGuildLeader
        Call WriteRequestGuildLeaderInfo
        Unload Me

End Sub

Private Sub imgPeticion_Click()
        Call WriteGuildRequestJoinerInfo(Nombre)

End Sub

Private Sub imgRechazar_Click()
        frmCommet.T = RECHAZOPJ
        frmCommet.Nombre = Nombre.Caption
        frmCommet.Show vbModeless, frmCharInfo

End Sub

Private Sub txtMiembro_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 x As Single, _
                                 y As Single)
        LastButtonPressed.ToggleToNormal

End Sub
