VERSION 5.00
Begin VB.Form frmCanjes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Sistema de Canjes"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   105
   ClientWidth     =   6000
   Icon            =   "frmCanjes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   3360
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   1080
      Width           =   510
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3930
      Left            =   225
      TabIndex        =   0
      Top             =   1140
      Width           =   2970
   End
   Begin VB.Label Ataque 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   4875
      TabIndex        =   7
      Top             =   3405
      Width           =   1215
   End
   Begin VB.Label Dropea 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   4125
      TabIndex        =   6
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label defFisica 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   4650
      TabIndex        =   5
      Top             =   2505
      Width           =   1215
   End
   Begin VB.Label defMagica 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4815
      TabIndex        =   4
      Top             =   2925
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5070
      TabIndex        =   3
      Top             =   2025
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   270
      Left            =   5475
      Top             =   120
      Width           =   405
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   3390
      Picture         =   "frmCanjes.frx":1CCA
      Top             =   4920
      Width           =   1800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "frmCanjes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormMove     As clsFormMovementManager

Private Canjear2Changed As Boolean

Private Sub Image1_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)
    
        Image1.Picture = LoadPicture(DirInterfaces & "BotonCanjearPuntos.jpg")
    
        Canjear2Changed = True
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

        If Canjear2Changed Then
                Image1.Picture = LoadPicture(vbNullString)
       
                Canjear2Changed = False

        End If

End Sub

Private Sub Form_Load()
    
        Set clsFormMove = New clsFormMovementManager
    
        Call clsFormMove.Initialize(Me)
        
        Me.Picture = LoadPicture(DirInterfaces & "VentanaCanjes.jpg")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
        Set clsFormMove = Nothing
    
End Sub

Private Sub Image1_Click()
        Call WriteCanjear(List1.listIndex + 1)

End Sub

Private Sub Image2_Click()
        Unload Me

End Sub

Private Sub list1_Click()

        Dim Canje As Byte

        Canje = List1.listIndex + 1

        Label2.Caption = Canjes(Canje).puntos
        defFisica.Caption = Canjes(Canje).defFisicaMin & "/" & Canjes(Canje).defFisicaMax
        defMagica.Caption = Canjes(Canje).defMagicaMin & "/" & Canjes(Canje).defMagicaMax
        Ataque.Caption = Canjes(Canje).AtaqueMin & "/" & Canjes(Canje).AtaqueMax
        Dropea.Caption = IIf(Canjes(Canje).Dropea, "No", "Si")
 
        Picture1.Picture = LoadPicture(App.path & "\Graficos\" & GrhData(Canjes(Canje).Graficos).FileNum & ".bmp")

End Sub
