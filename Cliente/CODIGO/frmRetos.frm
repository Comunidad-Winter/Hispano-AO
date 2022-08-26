VERSION 5.00
Begin VB.Form frmRetos 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   2970
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
   ScaleHeight     =   4485
   ScaleWidth      =   2970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox bItems 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   1740
      MaskColor       =   &H000000FF&
      TabIndex        =   6
      Top             =   3470
      Width           =   200
   End
   Begin VB.TextBox bGold 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   260
      Left            =   400
      TabIndex        =   5
      Top             =   2980
      Width           =   2175
   End
   Begin VB.OptionButton dRetos 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Index           =   1
      Left            =   970
      TabIndex        =   4
      Top             =   3480
      Width           =   200
   End
   Begin VB.OptionButton dRetos 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   3470
      Value           =   -1  'True
      Width           =   200
   End
   Begin VB.TextBox bName 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   260
      Index           =   2
      Left            =   400
      TabIndex        =   2
      Top             =   2335
      Width           =   2175
   End
   Begin VB.TextBox bName 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   260
      Index           =   1
      Left            =   400
      TabIndex        =   1
      Top             =   1720
      Width           =   2160
   End
   Begin VB.TextBox bName 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   260
      Index           =   0
      Left            =   400
      TabIndex        =   0
      Top             =   1075
      Width           =   2175
   End
   Begin VB.Image ImgMandar 
      Height          =   375
      Left            =   840
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Image ImgCerrar 
      Height          =   375
      Left            =   2640
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmRetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function CheckDatos() As Boolean
        ' @@ Chequeamos los Datos para no precesar mierda al pedo

        If dRetos(0).value = True Then
      
                If Not Len(bName(1).Text) <> 0 Then
                        Call ShowConsoleMsg("Introduce el nombre de tu Enemigo.")
                        CheckDatos = False

                        Exit Function

                End If

                If Not IsNumeric(bGold.Text) Then
                        Call ShowConsoleMsg("Introduce el Oro en numeros.")
                        CheckDatos = False

                        Exit Function

                End If

        ElseIf dRetos(1).value = True Then

                If Not Len(bName(0).Text) <> 0 Then
                        Call ShowConsoleMsg("Introduce el nombre de tu Compañero.")
                        CheckDatos = False

                        Exit Function

                End If

                If Not Len(bName(1).Text) <> 0 Then
                        Call ShowConsoleMsg("Introduce el nombre de tu Enemigo.")
                        CheckDatos = False

                        Exit Function

                End If

                If Not Len(bName(2).Text) <> 0 Then
                        Call ShowConsoleMsg("Introduce el nombre del compañero de tu Enemigo.")
                        CheckDatos = False

                        Exit Function

                End If

                If Not IsNumeric(bGold.Text) Then
                        Call ShowConsoleMsg("Introduce el Oro en numeros.")
                        CheckDatos = False

                        Exit Function

                End If
            
        End If

        CheckDatos = True

End Function

Private Sub dRetos_Click(Index As Integer)

        Select Case Index

                Case 0
                        dRetos(1).value = False
                        dRetos(0).value = True
                        
                        bName(0).Enabled = False
                        bName(1).Enabled = True
                        bName(2).Enabled = False
                  
                        bName(0).BackColor = &H808080
                        bName(2).BackColor = &H808080
                        'dRetos(1).Enabled = False

                Case 1
                        dRetos(1).value = True
                        dRetos(0).value = False
                        bName(0).Enabled = True
                        bName(2).Enabled = True
                  
                        bName(0).BackColor = &H0&
                        bName(2).BackColor = &H0&
                        'dRetos(0).Enabled = False
                  
        End Select

End Sub

Private Sub Form_Load()

        Me.Picture = LoadPicture(DirInterfaces & "VentanaRetos.jpg")

        ' @@ Empezamos con la opcion de 1 vs 1, por default
        Call dRetos_Click(0)
      
End Sub

Private Sub imgCerrar_Click()
        Unload Me

End Sub

Private Sub ImgMandar_Click()

        If Not CheckDatos Then Exit Sub
      
        ' @@ Chequear el uso de Rtrim$(), puede llegar a ser mucho mejor usar Trim$()

        If dRetos(0).value = True Then
                Call WriteOtherSendReto(RTrim$(bName(1).Text), Val(bGold.Text), (bItems.value <> 0))
        ElseIf dRetos(1).value = True Then
                Call WriteSendReto(RTrim$(bName(0).Text), RTrim$(bName(1).Text), RTrim$(bName(2).Text), Val(bGold.Text), (bItems.value <> 0))

        End If

End Sub
