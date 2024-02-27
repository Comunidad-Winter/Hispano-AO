Attribute VB_Name = "mod_DragAndDrop"
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
 
Public Sub DragToUser(ByVal UserIndex As Integer, _
                      ByVal tIndex As Integer, _
                      ByVal Slot As Byte, _
                      ByVal Amount As Integer)
   
        ' @@ Drag un slot a un usuario.
 
        Dim tObj       As Obj

        Dim tString    As String
        
        Dim errorFound As String

        Dim Espacio    As Boolean
        
        ' Puede dragear ?
        If Not CanDragObj(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.Navegando, errorFound) Then
                WriteConsoleMsg UserIndex, errorFound, FontTypeNames.FONTTYPE_INFOBOLD

                Exit Sub
End If
                      If UserList(tIndex).flags.Muerto = 1 Then 'cambiado userindex por targetindex mankeada de HEKAMIAH
                                Call WriteConsoleMsg(UserIndex, "¡El usuario está muerto!", FontTypeNames.FONTTYPE_INFO)
Exit Sub
        End If
        
        'Preparo el objeto.
        tObj.Amount = Amount
        tObj.objIndex = UserList(UserIndex).Invent.Object(Slot).objIndex
        
        If Not EsAdmin(UserList(UserIndex).Name) Then ' @@ Solo los admins Pueden crear los items con NoCreable = 1 -18/11/2015
                
                If Not EsAdmin(UserList(UserIndex).Name) Then
                        If ItemShop(UserList(UserIndex).Invent.Object(Slot).objIndex) = True Then
                                Call WriteConsoleMsg(UserIndex, "No puedes tirar items shop.", FontTypeNames.FONTTYPE_INFO) ' IvanLisz
                                Exit Sub

                        End If

                End If

        End If
        
         If (Amount <= 0 Or Amount > UserList(UserIndex).Invent.Object(Slot).Amount) Then
                Call WriteConsoleMsg(UserIndex, "No Tienes esa cantidad.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

        End If
        
        'TmpIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        Espacio = MeterItemEnInventario(tIndex, tObj)
 
        'No tiene espacio.

        If Not Espacio Then
        
                WriteConsoleMsg UserIndex, "El usuario no tiene espacio en su inventario.", FontTypeNames.FONTTYPE_INFOBOLD



                Exit Sub

        End If
 
        'Quito el objeto.
        QuitarUserInvItem UserIndex, Slot, Amount
 
        'Hago un update de su inventario.
        UpdateUserInv False, UserIndex, Slot
 
        'Preparo el mensaje para userINdex (quien dragea)
 
        tString = "Le has arrojado"
 
        If tObj.Amount <> 1 Then
                tString = tString & " " & tObj.Amount & " - " & ObjData(tObj.objIndex).Name
        Else
                tString = tString & " Tu " & ObjData(tObj.objIndex).Name

        End If
 
        tString = tString & " a " & UserList(tIndex).Name
 
        'Envio el mensaje
        WriteConsoleMsg UserIndex, tString, FontTypeNames.FONTTYPE_INFOBOLD
 
        'Preparo el mensaje para el otro usuario (quien recibe)
        tString = UserList(UserIndex).Name & " Te ha arrojado"
 
        If tObj.Amount <> 1 Then
                tString = tString & " " & tObj.Amount & " - " & ObjData(tObj.objIndex).Name
        Else
                tString = tString & " su " & ObjData(tObj.objIndex).Name

        End If
 
        'Envio el mensaje al otro usuario
        WriteConsoleMsg tIndex, tString & ".", FontTypeNames.FONTTYPE_INFOBOLD
 
End Sub
 
Public Sub DragToNPC(ByVal UserIndex As Integer, _
                     ByVal tNpc As Integer, _
                     ByVal Slot As Byte, _
                     ByVal Amount As Integer)

        ' @@ Drag un slot a un npc.

        On Error GoTo Errhandler
 
        Dim TeniaOro   As Long

        Dim TeniaObj   As Integer

        Dim TmpIndex   As Integer
        
        Dim errorFound As String
 
        TmpIndex = UserList(UserIndex).Invent.Object(Slot).objIndex
        TeniaOro = UserList(UserIndex).Stats.GLD
        TeniaObj = UserList(UserIndex).Invent.Object(Slot).Amount
        
        ' Puede dragear ?
        If Not CanDragObj(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.Navegando, errorFound) Then
                WriteConsoleMsg UserIndex, errorFound, FontTypeNames.FONTTYPE_INFOBOLD

                Exit Sub

        End If

        If Not EsAdmin(UserList(UserIndex).Name) Then
                If ItemShop(TmpIndex) = True Then
                        Call WriteConsoleMsg(UserIndex, "No puedes tirar items shop.", FontTypeNames.FONTTYPE_INFO) ' IvanLisz
                        Exit Sub

                End If

        End If

        'End If
              
         If (Amount <= 0 Or Amount > UserList(UserIndex).Invent.Object(Slot).Amount) Then
                Call WriteConsoleMsg(UserIndex, "No Tienes esa cantidad.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

        End If
 
        'Es un banquero?

        If Npclist(tNpc).NPCtype = eNPCType.Banquero Then
                Call UserDejaObj(UserIndex, Slot, Amount)
                
                'No tiene más el mismo amount que antes? entonces depositó.

                If TeniaObj <> UserList(UserIndex).Invent.Object(Slot).Amount Then
                        WriteConsoleMsg UserIndex, "Has depositado " & Amount & " - " & ObjData(TmpIndex).Name & ".", FontTypeNames.FONTTYPE_INFOBOLD
                        UpdateUserInv False, UserIndex, Slot

                End If

                'Es un npc comerciante?
        ElseIf Npclist(tNpc).Comercia = 1 Then
                'El npc compra cualquier tipo de items?

                If Not Npclist(tNpc).TipoItems <> eOBJType.otCualquiera Or Npclist(tNpc).TipoItems = ObjData(UserList(UserIndex).Invent.Object(Slot).objIndex).OBJType Then
                        
                        Call Comercio(eModoComercio.Venta, UserIndex, tNpc, Slot, Amount)
                        
                        'Ganó oro? si es así es porque lo vendió.

                        If TeniaOro <> UserList(UserIndex).Stats.GLD Then
                                WriteConsoleMsg UserIndex, "Le has vendido al " & Npclist(tNpc).Name & " " & Amount & " - " & ObjData(TmpIndex).Name & ".", FontTypeNames.FONTTYPE_INFOBOLD

                        End If

                Else
                        WriteConsoleMsg UserIndex, "El npc no está interesado en comprar este tipo de objetos.", FontTypeNames.FONTTYPE_INFOBOLD

                End If

        End If
 
        Exit Sub
 
Errhandler:
 
End Sub
 
Public Sub DragToPos(ByVal UserIndex As Integer, _
                     ByVal X As Byte, _
                     ByVal Y As Byte, _
                     ByVal Slot As Byte, _
                     ByVal Amount As Integer)

        '            Drag un slot a una posición.
 
        Dim errorFound As String

        Dim tObj       As Obj

        Dim tString    As String
        
        Dim TmpIndex   As Integer
        
        'No puede dragear en esa pos?

        If Not CanDragToPos(UserList(UserIndex).Pos.Map, X, Y, errorFound) Then
                WriteConsoleMsg UserIndex, errorFound, FontTypeNames.FONTTYPE_INFOBOLD

                Exit Sub

        End If
     
        ' Puede dragear ?
        If Not CanDragObj(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.Navegando, errorFound) Then
                WriteConsoleMsg UserIndex, errorFound, FontTypeNames.FONTTYPE_INFOBOLD

                Exit Sub

        End If

        'Creo el objeto.
        tObj.objIndex = UserList(UserIndex).Invent.Object(Slot).objIndex
        tObj.Amount = Amount
       
        If Not EsAdmin(UserList(UserIndex).Name) Then
                If ItemShop(tObj.objIndex) = True Then
                        Call WriteConsoleMsg(UserIndex, "No puedes tirar items shop.", FontTypeNames.FONTTYPE_INFO) ' IvanLisz
                        Exit Sub

                End If

        End If
        
         If (Amount <= 0 Or Amount > UserList(UserIndex).Invent.Object(Slot).Amount) Then
                Call WriteConsoleMsg(UserIndex, "No Tienes esa cantidad.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

        End If
 
        'Agrego el objeto a la posición.
        MakeObj tObj, UserList(UserIndex).Pos.Map, CInt(X), CInt(Y)
 
        'Quito el objeto.
        QuitarUserInvItem UserIndex, Slot, Amount
 
        'Actualizo el inventario
        UpdateUserInv False, UserIndex, Slot
 
        'Preparo el mensaje.
        tString = "Has arrojado "
 
        If tObj.Amount <> 1 Then
                tString = tString & tObj.Amount & " - " & ObjData(tObj.objIndex).Name
        Else
                tString = tString & "tu " & ObjData(tObj.objIndex).Name 'faltaba el tstring &

        End If
 
        'ENvio.
        WriteConsoleMsg UserIndex, tString & ".", FontTypeNames.FONTTYPE_INFOBOLD
 
End Sub
 
Private Function CanDragToPos(ByVal Map As Integer, _
                              ByVal X As Byte, _
                              ByVal Y As Byte, _
                              ByRef Error As String) As Boolean
 
        '            Devuelve si se puede dragear un item a x posición.
 
        CanDragToPos = False
 
        'Zona segura?

        If Not MapInfo(Map).Pk Then
                Error = "No está permitido arrojar objetos al suelo en zonas seguras."

                Exit Function

        End If
 
        'Ya hay objeto?

        If Not MapData(Map, X, Y).ObjInfo.objIndex = 0 Then
                Error = "Hay un objeto en esa posición!"

                Exit Function

        End If
 
        'Tile bloqueado?

        If Not MapData(Map, X, Y).Blocked = 0 Then
                Error = "No puedes arrojar objetos en esa posición"

                Exit Function

        End If
        
        If HayAgua(Map, X, Y) Then
                Error = "No puedes arrojar objetos al agua"
                
                Exit Function

        End If

        CanDragToPos = True
 
End Function
 
Private Function CanDragObj(ByVal objIndex As Integer, _
                            ByVal Navegando As Boolean, _
                            ByRef Error As String) As Boolean
 
        '            Devuelve si un objeto es drageable.
        
        CanDragObj = False
 
        If objIndex < 1 Or objIndex > UBound(ObjData()) Then Exit Function
 
        'Objeto newbie?

        If ObjData(objIndex).Newbie <> 0 Then
                Error = "No puedes arrojar objetos newbies!"

                Exit Function

        End If
 
        'Está navgeando?

        If Navegando Then
                Error = "No puedes arrojar un barco si estás navegando!"

                Exit Function

        End If
 
        CanDragObj = True
 
End Function

Public Sub moveItem(ByVal UserIndex As Integer, _
                    ByVal originalSlot As Integer, _
                    ByVal newSlot As Integer)

        Dim tmpObj As UserOBJ

        If (originalSlot <= 0) Or (newSlot <= 0) Then Exit Sub

        With UserList(UserIndex)

                If (originalSlot > .CurrentInventorySlots) Or (newSlot > .CurrentInventorySlots) Then Exit Sub
    
                tmpObj = .Invent.Object(originalSlot)
                .Invent.Object(originalSlot) = .Invent.Object(newSlot)
                .Invent.Object(newSlot) = tmpObj
    
                'Viva VB6 y sus putas deficiencias.
                If .Invent.AnilloEqpSlot = originalSlot Then
                        .Invent.AnilloEqpSlot = newSlot
                ElseIf .Invent.AnilloEqpSlot = newSlot Then
                        .Invent.AnilloEqpSlot = originalSlot

                End If
    
                If .Invent.ArmourEqpSlot = originalSlot Then
                        .Invent.ArmourEqpSlot = newSlot
                ElseIf .Invent.ArmourEqpSlot = newSlot Then
                        .Invent.ArmourEqpSlot = originalSlot

                End If
    
                If .Invent.BarcoSlot = originalSlot Then
                        .Invent.BarcoSlot = newSlot
                ElseIf .Invent.BarcoSlot = newSlot Then
                        .Invent.BarcoSlot = originalSlot

                End If
    
                If .Invent.CascoEqpSlot = originalSlot Then
                        .Invent.CascoEqpSlot = newSlot
                ElseIf .Invent.CascoEqpSlot = newSlot Then
                        .Invent.CascoEqpSlot = originalSlot

                End If
    
                If .Invent.EscudoEqpSlot = originalSlot Then
                        .Invent.EscudoEqpSlot = newSlot
                ElseIf .Invent.EscudoEqpSlot = newSlot Then
                        .Invent.EscudoEqpSlot = originalSlot

                End If
    
                If .Invent.MunicionEqpSlot = originalSlot Then
                        .Invent.MunicionEqpSlot = newSlot
                ElseIf .Invent.MunicionEqpSlot = newSlot Then
                        .Invent.MunicionEqpSlot = originalSlot

                End If
    
                If .Invent.WeaponEqpSlot = originalSlot Then
                        .Invent.WeaponEqpSlot = newSlot
                ElseIf .Invent.WeaponEqpSlot = newSlot Then
                        .Invent.WeaponEqpSlot = originalSlot

                End If

                Call UpdateUserInv(False, UserIndex, originalSlot)
                Call UpdateUserInv(False, UserIndex, newSlot)

        End With

End Sub

