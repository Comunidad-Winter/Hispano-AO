VERSION 5.00
Begin VB.Form frmGuildMember 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   376
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstMiembros 
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
      Height          =   2565
      Left            =   3075
      TabIndex        =   3
      Top             =   675
      Width           =   2610
   End
   Begin VB.ListBox lstClanes 
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
      Height          =   2565
      Left            =   195
      TabIndex        =   2
      Top             =   690
      Width           =   2610
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00000000&
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
      Height          =   270
      Left            =   180
      TabIndex        =   1
      Top             =   3630
      Width           =   2640
   End
   Begin VB.Label lblCantMiembros 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
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
      Left            =   4635
      TabIndex        =   0
      Top             =   3510
      Width           =   360
   End
   Begin VB.Image imgCerrar 
      Height          =   495
      Left            =   3000
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Image imgNoticias 
      Height          =   495
      Left            =   150
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Image imgDetalles 
      Height          =   375
      Left            =   150
      Top             =   4200
      Width           =   2655
   End
End
Attribute VB_Name = "frmGuildMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario    As clsFormMovementManager

'Private cBotonNoticias   As clsGraphicalButton

Private cBotonDetalles   As clsGraphicalButton

Private cBotonCerrar     As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()

        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me

        Me.Picture = LoadPicture(DirInterfaces & "VentanaMiembroClan.jpg")
    
        Call LoadButtons
    
End Sub

Private Sub LoadButtons()

        Dim GrhPath As String
    
        GrhPath = DirInterfaces

        'Set cBotonNoticias = New clsGraphicalButton
        Set cBotonDetalles = New clsGraphicalButton
        Set cBotonCerrar = New clsGraphicalButton
    
        Set LastButtonPressed = New clsGraphicalButton
    
        Call cBotonDetalles.Initialize(imgDetalles, GrhPath & "BotonDetallesMiembroClan.jpg", _
           GrhPath & "BotonDetallesMiembroClan.jpg", _
           GrhPath & "BotonDetallesClickMiembroClan.jpg", Me)

        'Call cBotonNoticias.Initialize(imgNoticias, GrhPath & "BotonNoticiasMiembroClan.jpg", _
         GrhPath & "BotonNoticiasMiembroClan.jpg", _
         GrhPath & "BotonNoticiasClickMiembroClan.jpg", Me)

        Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCerrarMimebroClan.jpg", _
           GrhPath & "BotonCerrarMimebroClan.jpg", _
           GrhPath & "BotonCerrarClickMimebroClan.jpg", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        LastButtonPressed.ToggleToNormal

End Sub

Private Sub ImgCerrar_Click()
        Unload Me

End Sub

Private Sub imgDetalles_Click()

        If lstClanes.listIndex = -1 Then Exit Sub
    
        frmGuildBrief.EsLeader = False

        Call WriteGuildRequestDetails(lstClanes.List(lstClanes.listIndex))

End Sub

Private Sub imgNoticias_Click()
        bShowGuildNews = True
        Call WriteShowGuildNews

End Sub

Private Sub txtSearch_Change()
        Call FiltrarListaClanes(txtSearch.Text)

End Sub

Private Sub txtSearch_GotFocus()

        With txtSearch
                .SelStart = 0
                .SelLength = Len(.Text)

        End With

End Sub

Public Sub FiltrarListaClanes(ByRef sCompare As String)

        Dim lIndex As Long
    
        If UBound(GuildNames) <> 0 Then

                With lstClanes
                        'Limpio la lista
                        .Clear
            
                        .Visible = False
            
                        ' Recorro los arrays
                        For lIndex = 0 To UBound(GuildNames)

                                ' Si coincide con los patrones
                                If InStr(1, UCase$(GuildNames(lIndex)), UCase$(sCompare)) Then
                                        ' Lo agrego a la lista
                                        .AddItem GuildNames(lIndex)

                                End If

                        Next lIndex
            
                        .Visible = True

                End With

        End If

End Sub

