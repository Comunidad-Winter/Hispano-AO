VERSION 5.00
Begin VB.Form frmCommet 
   BorderStyle     =   0  'None
   Caption         =   "Oferta de paz o alianza"
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5055
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
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
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
      Height          =   1890
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   510
      Width           =   4575
   End
   Begin VB.Image imgCerrar 
      Height          =   480
      Left            =   2880
      Top             =   2520
      Width           =   960
   End
   Begin VB.Image imgEnviar 
      Height          =   480
      Left            =   1080
      Top             =   2520
      Width           =   960
   End
End
Attribute VB_Name = "frmCommet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario             As clsFormMovementManager

Private Const MAX_PROPOSAL_LENGTH As Integer = 520

Private cBotonEnviar              As clsGraphicalButton

Private cBotonCerrar              As clsGraphicalButton

Public LastButtonPressed          As clsGraphicalButton

Public Nombre                     As String

Public T                          As TIPO

Public Enum TIPO

        ALIANZA = 1
        PAZ = 2
        RECHAZOPJ = 3

End Enum

Private Sub Form_Load()
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me

        Call LoadBackGround
        Call LoadButtons

End Sub

Private Sub LoadButtons()

        Dim GrhPath As String
    
        GrhPath = DirInterfaces

        Set cBotonEnviar = New clsGraphicalButton
        Set cBotonCerrar = New clsGraphicalButton
    
        Set LastButtonPressed = New clsGraphicalButton
    
        Call cBotonEnviar.Initialize(imgEnviar, GrhPath & "BotonEnviarSolicitud.jpg", _
           vbNullString, _
           GrhPath & "BotonEnviarClickSolicitud.jpg", Me)

        Call cBotonCerrar.Initialize(ImgCerrar, GrhPath & "BotonCerrarSolicitud.jpg", _
           vbNullString, _
           GrhPath & "BotonCerrarClickSolicitud.jpg", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        LastButtonPressed.ToggleToNormal

End Sub

Private Sub ImgCerrar_Click()
        Unload Me

End Sub

Private Sub imgEnviar_Click()

        If Text1 = "" Then
                If T = PAZ Or T = ALIANZA Then
                        MsgBox "Debes redactar un mensaje solicitando la paz o alianza al l�der de " & Nombre
                Else
                        MsgBox "Debes indicar el motivo por el cual rechazas la membres�a de " & Nombre

                End If
        
                Exit Sub

        End If
    
        If T = PAZ Then
                Call WriteGuildOfferPeace(Nombre, Replace(Text1, vbCrLf, "�"))
        
        ElseIf T = ALIANZA Then
                Call WriteGuildOfferAlliance(Nombre, Replace(Text1, vbCrLf, "�"))
        
        ElseIf T = RECHAZOPJ Then
                Call WriteGuildRejectNewMember(Nombre, Replace(Replace(Text1.Text, ",", " "), vbCrLf, " "))

                'Sacamos el char de la lista de aspirantes
                Dim i As Long
        
                For i = 0 To frmGuildLeader.solicitudes.ListCount - 1

                        If frmGuildLeader.solicitudes.List(i) = Nombre Then
                                frmGuildLeader.solicitudes.RemoveItem i
                                Exit For

                        End If

                Next i
        
                Me.Hide
                Unload frmCharInfo

        End If
    
        Unload Me

End Sub

Private Sub Text1_Change()

        If Len(Text1.Text) > MAX_PROPOSAL_LENGTH Then _
           Text1.Text = Left$(Text1.Text, MAX_PROPOSAL_LENGTH)

End Sub

Private Sub LoadBackGround()

        Select Case T

                Case TIPO.ALIANZA
                        Me.Picture = LoadPicture(DirInterfaces & "VentanaPropuestaAlianza.jpg")
            
                Case TIPO.PAZ
                        Me.Picture = LoadPicture(DirInterfaces & "VentanaPropuestaPaz.jpg")
            
                Case TIPO.RECHAZOPJ
                        Me.Picture = LoadPicture(DirInterfaces & "VentanaMotivoRechazo.jpg")
            
        End Select
    
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        LastButtonPressed.ToggleToNormal

End Sub
