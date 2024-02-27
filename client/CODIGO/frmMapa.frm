VERSION 5.00
Begin VB.Form frmMapa 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgCerrar 
      Height          =   255
      Left            =   8040
      MousePointer    =   99  'Custom
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************
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
'**************************************************************************

Option Explicit

Private clsFormulario As clsFormMovementManager

Private Enum eMaps

        ieGeneral
        ieDungeon

End Enum

Private picMaps(1) As Picture

Private CurrentMap As eMaps

''
' This form is used to show the world map.
' It has two levels. The world map and the dungeons map.
' You can toggle between them pressing the arrows
'
' @file     frmMapa.frm
' @author Marco Vanotti (MarKoxX) marcovanotti15@gmail.com
' @version 1.0.0
' @date 20080724

''
' Checks what Key is down. If the key is const vbKeyDown or const vbKeyUp, it toggles the maps, else the form unloads.
'
' @param KeyCode Specifies the key pressed
' @param Shift Specifies if Shift Button is pressed
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        '*************************************************
        'Author: Marco Vanotti (MarKoxX)
        'Last modified: 24/07/08
        '
        '*************************************************

        Select Case KeyCode

                Case vbKeyDown, vbKeyUp 'Cambiamos el "nivel" del mapa, al estilo Zelda ;D
                        '           ToggleImgMaps

                Case Else
                        Unload Me

        End Select
    
End Sub

''
' Toggle which image is visible.
'
Private Sub ToggleImgMaps()
        '*************************************************
        'Author: Marco Vanotti (MarKoxX)
        'Last modified: 24/07/08
        '
        '*************************************************

        '    imgToogleMap(CurrentMap).Visible = False
    
        '   If CurrentMap = eMaps.ieGeneral Then
        '           ImgCerrar.Visible = False
        '         CurrentMap = eMaps.ieDungeon
        '          ImgCerrar.Visible = True
        '          CurrentMap = eMaps.ieGeneral

        '  End If
    
        ' imgToogleMap(CurrentMap).Visible = True
        '  Me.Picture = picMaps(CurrentMap)

End Sub

''
' Load the images. Resizes the form, adjusts image's left and top and set lblTexto's Top and Left.
'
Private Sub Form_Load()
        '*************************************************
        'Author: Marco Vanotti (MarKoxX)
        'Last modified: 24/07/08
        '
        '*************************************************

        On Error GoTo error
    
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
        
        'Cargamos las imagenes de los mapas
        'Set picMaps(eMaps.ieGeneral) =
        'Set picMaps(eMaps.ieDungeon) = LoadPicture(DirInterfaces & "mapa1.jpg")
    
        ' Imagen de fondo
        'CurrentMap = eMaps.ieGeneral
        Me.Picture = LoadPicture(DirInterfaces & "mapa1.jpg")
    
        ImgCerrar.MouseIcon = picMouseIcon
        ' imgToogleMap(0).MouseIcon = picMouseIcon
        ' imgToogleMap(1).MouseIcon = picMouseIcon
    
        Exit Sub
error:
        MsgBox Err.Description, vbInformation, "Error: " & Err.number
        Unload Me

End Sub

Private Sub ImgCerrar_Click()
        Unload Me

End Sub

Private Sub imgToogleMap_Click(Index As Integer)
        'ToggleImgMaps

End Sub
