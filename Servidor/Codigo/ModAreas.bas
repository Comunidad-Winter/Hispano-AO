Attribute VB_Name = "ModAreas"
'**************************************************************
' ModAreas.bas - Module to allow the usage of areas instead of maps.
' Saves a lot of bandwidth.
'
' Original Idea by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
' Implemented by Lucio N. Tourrilhes (DuNga)
'**************************************************************

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

' Modulo de envio por areas compatible con la versión 9.10.x ... By DuNga

Option Explicit

'>>>>>>AREAS>>>>>AREAS>>>>>>>>AREAS>>>>>>>AREAS>>>>>>>>>>
Public Type AreaInfo

        AreaPerteneceX As Integer
        AreaPerteneceY As Integer
    
        AreaReciveX As Integer
        AreaReciveY As Integer
    
        MinX As Integer '-!!!
        MinY As Integer '-!!!
    
        AreaID As Long

End Type

Public Type ConnGroup

        CountEntrys As Long
        OptValue As Long
        UserEntrys() As Long

End Type

Public Const USER_NUEVO               As Byte = 255

'Cuidado:
' ¡¡¡LAS AREAS ESTÁN HARDCODEADAS!!!
Private CurDay                        As Byte

Private CurHour                       As Byte

Private AreasInfo(1 To 100, 1 To 100) As Byte

Private PosToArea(1 To 100)           As Byte

Private AreasRecive(12)               As Integer

Public ConnGroups()                   As ConnGroup

Public Sub InitAreas()

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim Loopc As Long

        Dim loopX As Long

        ' Setup areas...
        For Loopc = 0 To 11
                AreasRecive(Loopc) = (2 ^ Loopc) Or IIf(Loopc <> 0, 2 ^ (Loopc - 1), 0) Or IIf(Loopc <> 11, 2 ^ (Loopc + 1), 0)
        Next Loopc
    
        For Loopc = 1 To 100
                PosToArea(Loopc) = Loopc \ 9
        Next Loopc
    
        For Loopc = 1 To 100
                For loopX = 1 To 100
                        'Usamos 121 IDs de area para saber si pasasamos de area "más rápido"
                        AreasInfo(Loopc, loopX) = (Loopc \ 9 + 1) * (loopX \ 9 + 1)
                Next loopX
        Next Loopc

        'Setup AutoOptimizacion de areas
        CurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
        CurHour = Fix(Hour(time) \ 3) 'A ke parte de la hora pertenece
    
        ReDim ConnGroups(1 To NumMaps) As ConnGroup
    
        For Loopc = 1 To NumMaps
                ConnGroups(Loopc).OptValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & Loopc, CurDay & "-" & CurHour))
        
                If ConnGroups(Loopc).OptValue = 0 Then ConnGroups(Loopc).OptValue = 1
                ReDim ConnGroups(Loopc).UserEntrys(1 To ConnGroups(Loopc).OptValue) As Long
        Next Loopc

End Sub

Public Sub AreasOptimizacion()

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        'Es la función de autooptimizacion.... la idea es no mandar redimensionando arrays grandes todo el tiempo
        '**************************************************************
        Dim Loopc      As Long

        Dim tCurDay    As Byte

        Dim tCurHour   As Byte

        Dim EntryValue As Long
    
        If (CurDay <> IIf(Weekday(Date) > 6, 1, 2)) Or (CurHour <> Fix(Hour(time) \ 3)) Then
        
                tCurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
                tCurHour = Fix(Hour(time) \ 3) 'A ke parte de la hora pertenece
        
                For Loopc = 1 To NumMaps
                        EntryValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & Loopc, CurDay & "-" & CurHour))
                        Call WriteVar(DatPath & "AreasStats.dat", "Mapa" & Loopc, CurDay & "-" & CurHour, CInt((EntryValue + ConnGroups(Loopc).OptValue) \ 2))
            
                        ConnGroups(Loopc).OptValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & Loopc, tCurDay & "-" & tCurHour))

                        If ConnGroups(Loopc).OptValue = 0 Then ConnGroups(Loopc).OptValue = 1
                        If ConnGroups(Loopc).OptValue >= MapInfo(Loopc).NumUsers Then ReDim Preserve ConnGroups(Loopc).UserEntrys(1 To ConnGroups(Loopc).OptValue) As Long
                Next Loopc
        
                CurDay = tCurDay
                CurHour = tCurHour

        End If

End Sub

Public Sub CheckUpdateNeededUser(ByVal UserIndex As Integer, _
                                 ByVal Head As Byte, _
                                 Optional ByVal ButIndex As Boolean = False)

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: 28/10/2010
        'Es la función clave del sistema de areas... Es llamada al mover un user
        '15/07/2009: ZaMa - Now it doesn't send an invisible admin char info
        '28/10/2010: ZaMa - Now it doesn't send a saling char invisible message.
        '**************************************************************
        If UserList(UserIndex).AreasInfo.AreaID = AreasInfo(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) Then Exit Sub
    
        Dim MinX         As Long, MaxX As Long, MinY As Long, MaxY As Long, X As Long, Y As Long

        Dim TempInt      As Long, Map As Long

        Dim isZonaOscura As Boolean
    
        With UserList(UserIndex)
                MinX = .AreasInfo.MinX
                MinY = .AreasInfo.MinY
        
                If Head = eHeading.NORTH Then
                        MaxY = MinY - 1
                        MinY = MinY - 9
                        MaxX = MinX + 26
                        .AreasInfo.MinX = CInt(MinX)
                        .AreasInfo.MinY = CInt(MinY)
        
                ElseIf Head = eHeading.SOUTH Then
                        MaxY = MinY + 35
                        MinY = MinY + 27
                        MaxX = MinX + 26
                        .AreasInfo.MinX = CInt(MinX)
                        .AreasInfo.MinY = CInt(MinY - 18)
        
                ElseIf Head = eHeading.WEST Then
                        MaxX = MinX - 1
                        MinX = MinX - 9
                        MaxY = MinY + 26
                        .AreasInfo.MinX = CInt(MinX)
                        .AreasInfo.MinY = CInt(MinY)
        
                ElseIf Head = eHeading.EAST Then
                        MaxX = MinX + 35
                        MinX = MinX + 27
                        MaxY = MinY + 26
                        .AreasInfo.MinX = CInt(MinX - 18)
                        .AreasInfo.MinY = CInt(MinY)
           
                ElseIf Head = USER_NUEVO Then
                        'Esto pasa por cuando cambiamos de mapa o logeamos...
                        MinY = ((.Pos.Y \ 9) - 1) * 9
                        MaxY = MinY + 26
            
                        MinX = ((.Pos.X \ 9) - 1) * 9
                        MaxX = MinX + 26
            
                        .AreasInfo.MinX = CInt(MinX)
                        .AreasInfo.MinY = CInt(MinY)

                End If
        
                If MinY < 1 Then MinY = 1
                If MinX < 1 Then MinX = 1
                If MaxY > 100 Then MaxY = 100
                If MaxX > 100 Then MaxX = 100
        
                Map = .Pos.Map
        
                'Esto es para ke el cliente elimine lo "fuera de area..."
                Call WriteAreaChanged(UserIndex)
        
                isZonaOscura = (MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.zonaOscura)
        
                'Actualizamos!!!
                For X = MinX To MaxX
                        For Y = MinY To MaxY
                
                                '<<< User >>>
                                If MapData(Map, X, Y).UserIndex Then
                    
                                        TempInt = MapData(Map, X, Y).UserIndex
                    
                                        If UserIndex <> TempInt Then
                        
                                                ' Solo avisa al otro cliente si no es un admin invisible
                                                If Not (UserList(TempInt).flags.AdminInvisible = 1) Then
                                                        Call MakeUserChar(False, UserIndex, TempInt, Map, X, Y)
                            
                                                        ' Si esta navegando, siempre esta visible
                                                        If UserList(TempInt).flags.Navegando = 0 Then
                                                                If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
                                                                        If MapData(Map, X, Y).trigger = eTrigger.zonaOscura Then
                                                                                Call WriteSetInvisible(UserIndex, UserList(TempInt).Char.CharIndex, True)
                                                                        Else

                                                                                'Si el user estaba invisible le avisamos al nuevo cliente de eso
                                                                                If UserList(TempInt).flags.invisible Or UserList(TempInt).flags.Oculto Then
                                                                                        Call WriteSetInvisible(UserIndex, UserList(TempInt).Char.CharIndex, True)

                                                                                End If

                                                                        End If

                                                                End If

                                                        End If

                                                End If
                        
                                                ' Solo avisa al otro cliente si no es un admin invisible
                                                If Not (.flags.AdminInvisible = 1) Then
                                                        Call MakeUserChar(False, TempInt, UserIndex, .Pos.Map, .Pos.X, .Pos.Y)
                            
                                                        ' Si esta navegando, siempre esta visible
                                                        If .flags.Navegando = 0 Then
                                                                If UserList(TempInt).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
                                                                        If isZonaOscura Then
                                                                                Call WriteSetInvisible(TempInt, .Char.CharIndex, True)
                                                                        Else

                                                                                If .flags.invisible Or .flags.Oculto Then
                                                                                        Call WriteSetInvisible(TempInt, .Char.CharIndex, True)

                                                                                End If

                                                                        End If

                                                                End If

                                                        End If

                                                End If
                        
                                                Call FlushBuffer(TempInt)
                    
                                        ElseIf Head = USER_NUEVO Then

                                                If Not ButIndex Then
                                                        Call MakeUserChar(False, UserIndex, UserIndex, Map, X, Y)

                                                End If

                                        End If

                                End If
                
                                '<<< Npc >>>
                                If MapData(Map, X, Y).NpcIndex Then
                                        Call MakeNPCChar(False, UserIndex, MapData(Map, X, Y).NpcIndex, Map, X, Y)

                                End If
                 
                                '<<< Item >>>
                                If MapData(Map, X, Y).ObjInfo.objIndex Then
                                        If MapData(Map, X, Y).trigger <> eTrigger.zonaOscura Then
                                                TempInt = MapData(Map, X, Y).ObjInfo.objIndex

                                                If Not EsObjetoFijo(ObjData(TempInt).OBJType) Then
                                                        Call WriteObjectCreate(UserIndex, ObjData(TempInt).GrhIndex, X, Y)
                            
                                                        If ObjData(TempInt).OBJType = eOBJType.otPuertas Then
                                                                Call Bloquear(False, UserIndex, X, Y, MapData(Map, X, Y).Blocked)
                                                                Call Bloquear(False, UserIndex, X - 1, Y, MapData(Map, X - 1, Y).Blocked)

                                                        End If

                                                End If

                                        End If

                                End If
            
                        Next Y
                Next X
        
                'Precalculados :P
                TempInt = .Pos.X \ 9
                .AreasInfo.AreaReciveX = AreasRecive(TempInt)
                .AreasInfo.AreaPerteneceX = 2 ^ TempInt
        
                TempInt = .Pos.Y \ 9
                .AreasInfo.AreaReciveY = AreasRecive(TempInt)
                .AreasInfo.AreaPerteneceY = 2 ^ TempInt
        
                .AreasInfo.AreaID = AreasInfo(.Pos.X, .Pos.Y)

        End With

End Sub

Public Sub CheckUpdateNeededNpc(ByVal NpcIndex As Integer, ByVal Head As Byte)

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        ' Se llama cuando se mueve un Npc
        '**************************************************************
        If Npclist(NpcIndex).AreasInfo.AreaID = AreasInfo(Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y) Then Exit Sub
    
        Dim MinX         As Long, MaxX As Long, MinY As Long, MaxY As Long, X As Long, Y As Long

        Dim TempInt      As Long

        Dim UserIndex    As Integer

        Dim isZonaOscura As Boolean
    
        With Npclist(NpcIndex)
                MinX = .AreasInfo.MinX
                MinY = .AreasInfo.MinY
        
                If Head = eHeading.NORTH Then
                        MaxY = MinY - 1
                        MinY = MinY - 9
                        MaxX = MinX + 26
                        .AreasInfo.MinX = CInt(MinX)
                        .AreasInfo.MinY = CInt(MinY)
        
                ElseIf Head = eHeading.SOUTH Then
                        MaxY = MinY + 35
                        MinY = MinY + 27
                        MaxX = MinX + 26
                        .AreasInfo.MinX = CInt(MinX)
                        .AreasInfo.MinY = CInt(MinY - 18)
        
                ElseIf Head = eHeading.WEST Then
                        MaxX = MinX - 1
                        MinX = MinX - 9
                        MaxY = MinY + 26
                        .AreasInfo.MinX = CInt(MinX)
                        .AreasInfo.MinY = CInt(MinY)
        
                ElseIf Head = eHeading.EAST Then
                        MaxX = MinX + 35
                        MinX = MinX + 27
                        MaxY = MinY + 26
                        .AreasInfo.MinX = CInt(MinX - 18)
                        .AreasInfo.MinY = CInt(MinY)
           
                ElseIf Head = USER_NUEVO Then
                        'Esto pasa por cuando cambiamos de mapa o logeamos...
                        MinY = ((.Pos.Y \ 9) - 1) * 9
                        MaxY = MinY + 26
            
                        MinX = ((.Pos.X \ 9) - 1) * 9
                        MaxX = MinX + 26
            
                        .AreasInfo.MinX = CInt(MinX)
                        .AreasInfo.MinY = CInt(MinY)

                End If
        
                If MinY < 1 Then MinY = 1
                If MinX < 1 Then MinX = 1
                If MaxY > 100 Then MaxY = 100
                If MaxX > 100 Then MaxX = 100
        
                'Actualizamos!!!
                If MapInfo(.Pos.Map).NumUsers <> 0 Then
                        isZonaOscura = (MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.zonaOscura)
            
                        For X = MinX To MaxX
                                For Y = MinY To MaxY
                                        UserIndex = MapData(.Pos.Map, X, Y).UserIndex
                    
                                        If UserIndex Then
                                                Call MakeNPCChar(False, UserIndex, NpcIndex, .Pos.Map, .Pos.X, .Pos.Y)
                        
                                                If isZonaOscura Then
                                                        Call WriteSetInvisible(UserIndex, .Char.CharIndex, True)

                                                End If

                                        End If

                                Next Y
                        Next X

                End If
        
                'Precalculados :P
                TempInt = .Pos.X \ 9
                .AreasInfo.AreaReciveX = AreasRecive(TempInt)
                .AreasInfo.AreaPerteneceX = 2 ^ TempInt
            
                TempInt = .Pos.Y \ 9
                .AreasInfo.AreaReciveY = AreasRecive(TempInt)
                .AreasInfo.AreaPerteneceY = 2 ^ TempInt
        
                .AreasInfo.AreaID = AreasInfo(.Pos.X, .Pos.Y)

        End With

End Sub

Public Sub QuitarUser(ByVal UserIndex As Integer, ByVal Map As Integer)

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        On Error GoTo ErrorHandler

        Dim TempVal As Long

        Dim Loopc   As Long
    
        'Search for the user
        For Loopc = 1 To ConnGroups(Map).CountEntrys

                If ConnGroups(Map).UserEntrys(Loopc) = UserIndex Then Exit For
        Next Loopc
    
        'Char not found
        If Loopc > ConnGroups(Map).CountEntrys Then Exit Sub
    
        'Remove from old map
        ConnGroups(Map).CountEntrys = ConnGroups(Map).CountEntrys - 1
        TempVal = ConnGroups(Map).CountEntrys
    
        'Move list back
        For Loopc = Loopc To TempVal
                ConnGroups(Map).UserEntrys(Loopc) = ConnGroups(Map).UserEntrys(Loopc + 1)
        Next Loopc
    
        If TempVal > ConnGroups(Map).OptValue Then 'Nescesito Redim?
                ReDim Preserve ConnGroups(Map).UserEntrys(1 To TempVal) As Long

        End If
    
        Exit Sub
    
ErrorHandler:
    
        Dim UserName As String

        If UserIndex > 0 Then UserName = UserList(UserIndex).Name

        Call LogError("Error en QuitarUser " & Err.Number & ": " & Err.description & _
           ". User: " & UserName & "(" & UserIndex & ")")

End Sub

Public Sub AgregarUser(ByVal UserIndex As Integer, _
                       ByVal Map As Integer, _
                       Optional ByVal ButIndex As Boolean = False)

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: 04/01/2007
        'Modified by Juan Martín Sotuyo Dodero (Maraxus)
        '   - Now the method checks for repetead users instead of trusting parameters.
        '   - If the character is new to the map, update it
        '**************************************************************
        Dim TempVal As Long

        Dim EsNuevo As Boolean

        Dim i       As Long
    
        If Not MapaValido(Map) Then Exit Sub
    
        EsNuevo = True
    
        'Prevent adding repeated users
        For i = 1 To ConnGroups(Map).CountEntrys

                If ConnGroups(Map).UserEntrys(i) = UserIndex Then
                        EsNuevo = False
                        Exit For

                End If

        Next i
    
        If EsNuevo Then
                'Update map and connection groups data
                ConnGroups(Map).CountEntrys = ConnGroups(Map).CountEntrys + 1
                TempVal = ConnGroups(Map).CountEntrys
        
                If TempVal > ConnGroups(Map).OptValue Then 'Nescesito Redim
                        ReDim Preserve ConnGroups(Map).UserEntrys(1 To TempVal) As Long

                End If
        
                ConnGroups(Map).UserEntrys(TempVal) = UserIndex

        End If
    
        With UserList(UserIndex)
                'Update user
                .AreasInfo.AreaID = 0
        
                .AreasInfo.AreaPerteneceX = 0
                .AreasInfo.AreaPerteneceY = 0
                .AreasInfo.AreaReciveX = 0
                .AreasInfo.AreaReciveY = 0

        End With
    
        Call CheckUpdateNeededUser(UserIndex, USER_NUEVO, ButIndex)

End Sub

Public Sub AgregarNpc(ByVal NpcIndex As Integer)

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        With Npclist(NpcIndex)
                .AreasInfo.AreaID = 0
        
                .AreasInfo.AreaPerteneceX = 0
                .AreasInfo.AreaPerteneceY = 0
                .AreasInfo.AreaReciveX = 0
                .AreasInfo.AreaReciveY = 0

        End With
    
        Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)

End Sub
