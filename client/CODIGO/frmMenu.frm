VERSION 5.00
Begin VB.Form frmMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2775
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
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image ImgCerrar 
      Height          =   255
      Left            =   2520
      Top             =   0
      Width           =   255
   End
   Begin VB.Image ImgEstadisticas 
      Height          =   255
      Left            =   525
      Top             =   3285
      Width           =   1695
   End
   Begin VB.Image ImgClanes 
      Height          =   255
      Left            =   525
      Top             =   2655
      Width           =   1695
   End
   Begin VB.Image ImgOpciones 
      Height          =   255
      Left            =   525
      Top             =   765
      Width           =   1695
   End
   Begin VB.Image ImgShops 
      Height          =   255
      Left            =   525
      Top             =   1395
      Width           =   1695
   End
   Begin VB.Image ImgRetos 
      Height          =   255
      Left            =   525
      Top             =   2025
      Width           =   1695
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario      As clsFormMovementManager

Private cBotonCanjes       As clsGraphicalButton

Private cBotonRetos        As clsGraphicalButton

Private cBotonOpciones     As clsGraphicalButton

Private cBotonEstadisticas As clsGraphicalButton

Private cBotonClanes       As clsGraphicalButton

Private Sub Form_Load()
        Me.Picture = LoadPicture(DirInterfaces & "VentanaMenu.jpg")

End Sub

Private Sub ImgCerrar_Click()
        Unload Me

End Sub

Private Sub ImgClanes_Click()

        If frmGuildLeader.Visible Then
                Unload frmGuildLeader

        End If

        Call WriteRequestGuildLeaderInfo
        
        Unload Me

End Sub

Private Sub ImgEstadisticas_Click()

        LlegaronAtrib = False
        LlegaronSkills = False
        LlegoFama = False

        Call WriteRequestAtributes
        Call WriteRequestSkills
        Call WriteRequestMiniStats
        Call WriteRequestFame
        
        Call FlushBuffer

        Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
        Loop

        frmEstadisticas.Iniciar_Labels
        frmEstadisticas.Show vbModeless, frmMain
        
        LlegaronAtrib = False
        LlegaronSkills = False
        LlegoFama = False

End Sub

Private Sub imgOpciones_Click()

        Call frmOpciones.Show(vbModeless, frmMain)
        Unload Me

End Sub

Private Sub ImgRetos_Click()

        Call frmRetos.Show(vbModeless, frmMain)
        Unload Me

End Sub

Private Sub ImgShops_Click()

        frmCanjes.Show
        frmCanjes.List1.Clear
        
        Call WriteCanje
        
        Unload Me

End Sub
