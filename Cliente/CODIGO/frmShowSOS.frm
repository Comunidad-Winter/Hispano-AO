VERSION 5.00
Begin VB.Form frmShowSOS 
   BorderStyle     =   0  'None
   ClientHeight    =   4995
   ClientLeft      =   120
   ClientTop       =   45
   ClientWidth     =   9390
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   333
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   626
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtRespuesta 
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
      Left            =   2655
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2670
      Width           =   6090
   End
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
      Left            =   2655
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   495
      Width           =   6090
   End
   Begin VB.ListBox List1 
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
      Height          =   3735
      Left            =   495
      TabIndex        =   0
      Top             =   495
      Width           =   1620
   End
   Begin VB.Image imgCerrar 
      Height          =   375
      Left            =   630
      MouseIcon       =   "frmShowSOS.frx":0000
      MousePointer    =   99  'Custom
      Top             =   4470
      Width           =   1335
   End
   Begin VB.Image imgResponder 
      Height          =   375
      Left            =   7485
      MouseIcon       =   "frmShowSOS.frx":0152
      MousePointer    =   99  'Custom
      Top             =   4470
      Width           =   1335
   End
   Begin VB.Menu menU_usuario 
      Caption         =   "Usuario"
      Visible         =   0   'False
      Begin VB.Menu mnuIR 
         Caption         =   "Ir donde esta el usuario"
      End
      Begin VB.Menu mnutraer 
         Caption         =   "Traer usuario"
      End
      Begin VB.Menu mnuBorrar 
         Caption         =   "Borrar mensaje"
      End
   End
End
Attribute VB_Name = "frmShowSOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Dim CurrentUser As String

Private Sub Form_Deactivate()
        Me.Visible = False
        List1.Clear
End Sub

Private Sub Form_Load()
    
    Me.Picture = LoadPicture(DirInterfaces & "VentanaConsultas.jpg")

    List1.Clear
End Sub

Private Sub imgCerrar_Click()
        Me.Visible = False
        List1.Clear
End Sub

Private Sub imgResponder_Click()

        If Not LenB(txtRespuesta.Text) = 0 Then

                If List1.listIndex < 0 Then Exit Sub
    
                Call WriteResponderConsulta(CurrentUser, txtRespuesta.Text)

                Dim aux As String

                aux = mid$(ReadField(1, List1.List(List1.listIndex), Asc("-")), 10, Len(ReadField(1, List1.List(List1.listIndex), Asc("-"))))
                Call WriteSOSRemove(aux)
   
                List1.RemoveItem List1.listIndex
        
                txtRespuesta.Text = vbNullString
                txtConsulta.Text = vbNullString
        End If
        
End Sub

Private Sub list1_Click()

        On Error Resume Next

        Dim ind    As Integer
        Dim Nick() As String

        ind = Val(ReadField(2, List1.List(List1.listIndex), Asc("-")))
        
        Nick() = Split(List1.List(List1.listIndex), " ", 2)
        
        'MsgBox Nick(1)
        CurrentUser = Nick(1)
        
        WriteGetConsulta CurrentUser
        
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

        If Button = vbRightButton Then
                PopupMenu menU_usuario
        End If

End Sub

Private Sub mnuBorrar_Click()

        If List1.listIndex < 0 Then Exit Sub

        Dim aux As String

        aux = mid$(ReadField(1, List1.List(List1.listIndex), Asc("-")), 10, Len(ReadField(1, List1.List(List1.listIndex), Asc("-"))))
        
        Call WriteSOSRemove(aux)
    
        List1.RemoveItem List1.listIndex
End Sub
