Attribute VB_Name = "Mod_Retos2vs2"
Option Explicit

Private Mapa_Arenas As Integer

Public Type ruleStruct

        drop_inv        As Boolean
        gold_gamble     As Long

End Type

Public Type teamStruct

        user_Index(1)   As Integer
        round_count     As Byte
        return_city     As Byte

End Type

Public Type retoStruct

        team_array(1)   As teamStruct
        general_rules   As ruleStruct
        count_Down      As Byte
        used_ring       As Boolean
        haydrop As Boolean
        nextRoundCount  As Integer

End Type

Public Type userStruct

        tempStruct      As retoStruct
        accept_count    As Byte
        reto_Index      As Integer
        nick_sender     As String
        reto_used       As Boolean
        return_city     As Byte
        tmp_Time          As Byte
        acceptedOK      As Boolean
        acceptLimit     As Integer
         
End Type

Public Type retoPosStruct

        Map As Integer
        X As Integer
        Y As Integer

End Type

Public reto_List() As retoStruct
Public retoPos()   As retoPosStruct

Public Sub loop_reto()

        '
        ' @ amishar.-
    
        Dim Loopc As Long
    
        For Loopc = LBound(reto_List()) To UBound(reto_List())

                If (reto_List(Loopc).used_ring) Then
                        Call loop_reto_index(Loopc)

                End If

        Next Loopc

End Sub

Public Function can_Attack(ByVal attackerIndex As Integer, _
                           ByVal victimIndex As Integer) As Boolean

        '
        ' @ amishar.-
    
        Dim retoIndex As Integer
        Dim teamIndex As Integer
        Dim tempIndex As Integer
        Dim teamLoop  As Long
    
        can_Attack = True
    
        retoIndex = UserList(attackerIndex).sReto.reto_Index
    
        teamIndex = -1
    
        If reto_List(retoIndex).used_ring Then

                For teamLoop = 0 To 1

                        If reto_List(retoIndex).team_array(teamLoop).user_Index(0) = attackerIndex Or reto_List(retoIndex).team_array(teamLoop).user_Index(1) = attackerIndex Then
                                teamIndex = teamLoop

                                Exit For

                        End If

                Next teamLoop
       
                If teamIndex <> -1 Then
                        tempIndex = IIf(reto_List(retoIndex).team_array(teamIndex).user_Index(0) = attackerIndex, 1, 0)

                        If reto_List(retoIndex).team_array(teamIndex).user_Index(tempIndex) = victimIndex Then
                                can_Attack = False

                        End If

                End If

        End If

End Function

Private Sub loop_reto_index(ByVal reto_Index As Integer)

        '
        ' @ amishar.-
    
        Dim i As Long
        Dim j As Long
        Dim H As Integer
        Dim m As String
    
        With reto_List(reto_Index)
        
                If (.nextRoundCount <> 0) Then
                        .nextRoundCount = .nextRoundCount - 1
            
                        If (.nextRoundCount = 0) Then
                                Call warp_Teams(reto_Index, True)
                 
                                .count_Down = 6

                        End If

                End If
        
                If (.count_Down <> 0) Then
                        .count_Down = (.count_Down - 1)
            
                        If (.count_Down > 0) Then
                                m = CStr(.count_Down) & "..."
                        Else
                                m = "¡YA!"

                        End If
            
                        For i = 0 To 1
                                For j = 0 To 1
                                        H = .team_array(i).user_Index(j)
                    
                                        If (H <> 0) Then
                                                If UserList(H).ConnID <> -1 Then
                                                        Call Protocol.WriteConsoleMsg(H, m, FontTypeNames.FONTTYPE_GUILD)
                           
                                                        If (.count_Down = 0) Then Call Protocol.WritePauseToggle(H)

                                                End If

                                        End If

                                Next j
                        Next i

                End If

        End With
    
End Sub

Public Function get_reto_index() As Integer

        '
        ' @ amishar.-
    
        Dim Loopc As Long
    
        For Loopc = LBound(reto_List()) To UBound(reto_List())

                If (reto_List(Loopc).used_ring = False) Then
                        get_reto_index = CInt(Loopc)

                        Exit Function

                End If

        Next Loopc
    
        get_reto_index = -1

End Function

Public Sub set_reto_struct(ByVal user_Index As Integer, _
                           ByVal my_team As String, _
                           ByRef enemy_name As String, _
                           ByRef team_enemy As String, _
                           ByVal invDrop As Boolean, _
                           ByVal GoldAmount As Long)

        '
        ' @ amishar.-
    
        With UserList(user_Index).sReto
                .accept_count = 0
         
                With .tempStruct
                        .count_Down = 0
                        .used_ring = False
              
                        With .team_array(0)
                                .user_Index(0) = user_Index
                                .user_Index(1) = NameIndex(my_team)

                        End With
              
                        With .team_array(1)
                                .user_Index(0) = NameIndex(enemy_name)
                                .user_Index(1) = NameIndex(team_enemy)

                        End With
              
                        With .general_rules
                                .drop_inv = invDrop
                                .gold_gamble = GoldAmount

                        End With
              
                End With
         
        End With

End Sub

Public Sub user_retoLoop(ByVal user_Index As Integer)

        '
        ' @ amishar.-
    
        With UserList(user_Index).sReto

                If (.acceptLimit <> 0) Then
                        .acceptLimit = .acceptLimit - 1
             
                        If (.acceptLimit <= 0) Then
                                Call message_reto(.tempStruct, "El reto se ha autocancelado debido a que el tiempo para aceptar ha llegado a su límite.")
                 
                                Dim j As Long
                                Dim i As Long
                                Dim n As Integer
                                Dim b As userStruct
                 
                                For j = 0 To 1
                                        For i = 0 To 1
                                                n = .tempStruct.team_array(j).user_Index(i)
                         
                                                If n > 0 Then
                                                        If UCase$(UserList(n).sReto.nick_sender) = UCase$(UserList(user_Index).Name) Then
                                                                UserList(n).sReto.nick_sender = vbNullString
                                                                UserList(n).sReto.acceptedOK = False

                                                        End If

                                                End If

                                        Next i
                                Next j
                 
                                UserList(user_Index).sReto = b

                        End If

                End If

                If (.return_city <> 0) Then
                        .return_city = .return_city - 1
             
                        If (.return_city = 0) Then

                                Dim p As WorldPos
                                Dim k As WorldPos
                 
                                p.Map = 1
                                p.X = 50
                                p.Y = 50
                 
                                Call ClosestStablePos(p, k)
                                Call WarpUserChar(user_Index, p.Map, k.X, k.Y, True)
                 
                                Call Protocol.WriteConsoleMsg(user_Index, "Regresas a la ciudad.", FontTypeNames.FONTTYPE_GUILD)
                                
                                Dim rIndex As Integer
                                rIndex = .reto_Index
                                
                                .nick_sender = vbNullString
                                .reto_Index = 0
                                
                                Call clear_data(rIndex)

                        End If
             
                End If

        End With

End Sub

Public Sub erase_userData(ByVal user_Index As Integer)

        '
        ' @ amishar.-
    
        With UserList(user_Index).sReto
    
                Dim dumpStruct As retoStruct
    
                .accept_count = 0
                .nick_sender = vbNullString
                .reto_Index = 0
                .tmp_Time = 0
                .reto_used = False
                .tempStruct = dumpStruct
                .return_city = 0
                .acceptedOK = False
            
        End With

End Sub

Public Function can_send_reto(ByVal user_Index As Integer, _
                              ByRef fError As String) As Boolean

        '
        ' @ amishar.-
    
        can_send_reto = False
    
        With UserList(user_Index)

                If (.flags.Muerto <> 0) Then
                        fError = "¡Estás muerto!"

                        Exit Function

                End If
         
                If (.Counters.Pena <> 0) Then
                        fError = "Estás en la cárcel"

                        Exit Function

                End If
         
                If (.Stats.GLD < .sReto.tempStruct.general_rules.gold_gamble) Then
                        fError = "No tienes el oro necesario"

                        Exit Function

                End If
                         
                If (.Pos.Map <> eCiudad.cUllathorpe) Then
                        fError = .Name & " está fuera de su hogar."

                        Exit Function

                End If
         
                If (.mReto.reto_Index <> 0) Or (.sReto.reto_used = True) Then
                        fError = .Name & " ya está en reto."

                        Exit Function

                End If
         
                If (.Stats.ELV < 30) Then
                        fError = "Debes ser mayor a nivel 30!"

                        Exit Function

                End If
         
                With .sReto.tempStruct
                        can_send_reto = check_User(.team_array(0).user_Index(1), fError, .general_rules.gold_gamble)
              
                        If (can_send_reto) Then
                                can_send_reto = check_User(.team_array(1).user_Index(0), fError, .general_rules.gold_gamble)
                        Else

                                Exit Function

                        End If
              
                        If (can_send_reto) Then
                                can_send_reto = check_User(.team_array(1).user_Index(1), fError, .general_rules.gold_gamble)
                        Else

                                Exit Function

                        End If
              
                End With

        End With

End Function

Private Function check_User(ByVal user_Index As Integer, _
                            ByRef fError As String, _
                            ByVal goldGamble As Long) As Boolean

        '
        ' @ amishar.-
    
        check_User = False
    
        If (user_Index = 0) Then
                fError = "Algún usuario está offline."

                Exit Function

        End If
    
        With UserList(user_Index)

                If (.flags.Muerto <> 0) Then
                        fError = .Name & " ¡Está muerto!"

                        Exit Function

                End If
         
                If (.Counters.Pena <> 0) Then
                        fError = .Name & " Está en la cárcel"

                        Exit Function

                End If
                    
                If (.Pos.Map <> eCiudad.cUllathorpe) Then
                        fError = .Name & " está fuera de su hogar."

                        Exit Function

                End If
         
                If (.mReto.reto_Index <> 0) Or (.sReto.reto_used = True) Then
                        fError = .Name & " ya está en reto."

                        Exit Function

                End If
         
                If (.Stats.GLD < goldGamble) Then
                        fError = .Name & " No tiene el oro necesario"

                        Exit Function

                End If
                  
                If (.Stats.ELV < 30) Then
                        fError = .Name & " debe ser mayor a nivel 30!"

                        Exit Function

                End If
         
                check_User = True
         
        End With

End Function

Public Sub send_Reto(ByVal user_Index As Integer)

        '
        ' @ amishar.-
    
        With UserList(user_Index).sReto

                Dim i          As Long
                Dim j          As Long
         
                Dim team_str   As String
                Dim gamble_str As String
         
                team_str = UserList(.tempStruct.team_array(0).user_Index(0)).Name & " y " & UserList(.tempStruct.team_array(0).user_Index(1)).Name & " vs " & UserList(.tempStruct.team_array(1).user_Index(0)).Name & " y " & UserList(.tempStruct.team_array(1).user_Index(1)).Name
         
                gamble_str = " apostando " & Format$(.tempStruct.general_rules.gold_gamble, "#,###") & " monedas de oro"
                  
                If (.tempStruct.general_rules.drop_inv) Then
                        gamble_str = " y los items del inventario"

                End If
         
                For i = 0 To 1
                        For j = 0 To 1
                                UserList(.tempStruct.team_array(i).user_Index(j)).sReto.nick_sender = UCase$(UserList(user_Index).Name)
                 
                                If (.tempStruct.team_array(i).user_Index(j) <> user_Index) Then
                                        Call Protocol.WriteConsoleMsg(.tempStruct.team_array(i).user_Index(j), "Solicitud de reto modalidad 2vs2 : " & team_str & " " & gamble_str & " para aceptar tipea /RETAR " & UCase$(UserList(user_Index).Name) & ".", FontTypeNames.FONTTYPE_GUILD)

                                End If

                        Next j
                Next i
         
                Call Protocol.WriteConsoleMsg(user_Index, "Se han enviado las solicitudes.", FontTypeNames.FONTTYPE_GUILD)
                .acceptLimit = 60

        End With

End Sub

Public Sub disconnect_Reto(ByVal user_Index As Integer)

        '
        ' @ amishar.-
    
        Dim team_Index  As Integer
        Dim user_slot   As Integer
        Dim team_winner As Byte
        Dim reto_Index  As Integer
    
        reto_Index = UserList(user_Index).sReto.reto_Index
    
        team_Index = find_Team(user_Index, reto_Index)

        If (team_Index <> -1) Then
                team_winner = IIf(team_Index = 1, 0, 1)
                Call finish_reto(UserList(user_Index).sReto.reto_Index, team_winner)

        End If
    
End Sub

Public Sub closeOtherReto(ByVal UserIndex As Integer)

        '
        ' @ amishar.-
    
        Dim j As Long
        Dim i As Long
        Dim n As Integer
        Dim c As Boolean
    
        n = NameIndex(UserList(UserIndex).sReto.nick_sender)
    
        If (n > 0) Then

                For i = 0 To 1
                        For j = 0 To 1

                                With UserList(n).sReto.tempStruct.team_array(i)

                                        If (.user_Index(j) = UserIndex) Then
                                                c = True

                                                Exit For

                                        End If

                                End With

                        Next j
                Next i
        
                If c = True Then

                        For i = 0 To 1
                                For j = 0 To 1

                                        With UserList(n).sReto.tempStruct.team_array(i)

                                                If (.user_Index(j) > 0) Then
                                                        If UCase$(UserList(.user_Index(j)).sReto.nick_sender) = UCase$(UserList(n).Name) Then
                                                                Call Protocol.WriteConsoleMsg(.user_Index(j), "El reto solicitado por " & UserList(n).Name & " ha sido cancelado debido a la desconexión de un participante.", FontTypeNames.FONTTYPE_GUILD)

                                                        End If

                                                End If

                                        End With

                                Next j
                        Next i

                End If

        End If
    
End Sub

Public Sub accept_Reto(ByVal user_Index As Integer, ByVal requestName As String)

        '
        ' @ amishar.-
    
        Dim sendIndex As Integer
        Dim i         As Long
        Dim j         As Long
    
        sendIndex = NameIndex(requestName)
    
        If (sendIndex = 0) Or (UCase$(requestName) <> UserList(user_Index).sReto.nick_sender) Then
                Call Protocol.WriteConsoleMsg(user_Index, requestName & " no te está retando!!", FontTypeNames.FONTTYPE_GUILD)

                Exit Sub

        End If

        If Not (UCase$(UserList(user_Index).Name) <> UserList(user_Index).sReto.nick_sender) Then
                Call Protocol.WriteConsoleMsg(user_Index, "No te puedes aceptar a ti mismo", FontTypeNames.FONTTYPE_GUILD)

                Exit Sub

        End If
      
        If (sendIndex = 0) Then Exit Sub
    
        If UserList(user_Index).sReto.acceptedOK Then
                Call Protocol.WriteConsoleMsg(user_Index, "¡Ya has aceptado!", FontTypeNames.FONTTYPE_GUILD)

                Exit Sub

        End If
    
        UserList(sendIndex).sReto.accept_count = (UserList(sendIndex).sReto.accept_count + 1)
    
        Call message_reto(UserList(sendIndex).sReto.tempStruct, UserList(user_Index).Name & " aceptó el reto.")
    
        If (UserList(sendIndex).sReto.accept_count = 3) Then
                Call init_reto(sendIndex)
                Call message_reto(UserList(sendIndex).sReto.tempStruct, "Todos los participantes han aceptado el reto.")

        End If
      
        UserList(user_Index).sReto.acceptedOK = True
    
End Sub

Private Sub init_reto(ByVal userSendIndex As Integer)

        '
        ' @ amishar.-
    
        Dim reto_Index As Integer
    
        reto_Index = get_reto_index()
    
        If (reto_Index = -1) Then
                Call message_reto(UserList(userSendIndex).sReto.tempStruct, "Reto cancelado, todas las arenas están ocupadas.")

                Exit Sub

        End If
      
        UserList(userSendIndex).sReto.acceptLimit = 0
        reto_List(reto_Index) = UserList(userSendIndex).sReto.tempStruct
        reto_List(reto_Index).used_ring = True
        reto_List(reto_Index).count_Down = 6
    
        Call warp_Teams(reto_Index)
    
End Sub

Private Sub warp_Teams(ByVal reto_Index As Integer, _
                       Optional ByVal respawnUser As Boolean = False)

        '
        ' @ amishar.-
    
        With reto_List(reto_Index)

                Dim Loopc As Long
                Dim mPosX As Byte
                Dim mPosY As Byte
                Dim nUser As Integer
         
                .count_Down = 6
         
                For Loopc = 0 To 1
                        nUser = .team_array(0).user_Index(Loopc)
              
                        If (nUser <> 0) Then
                                If (UserList(nUser).ConnID <> -1) Then
                                        mPosX = get_pos_x(reto_Index, 1, CInt(Loopc))
                                        mPosY = get_pos_y(reto_Index, 1, CInt(Loopc))
                     
                                        UserList(nUser).sReto.reto_used = True
                                        UserList(nUser).sReto.reto_Index = reto_Index
                     
                                        Call WarpUserChar(nUser, Mapa_Arenas, mPosX, mPosY, True)
                                        Call Protocol.WritePauseToggle(nUser)
                     
                                        If (respawnUser) Then
                                                If (UserList(nUser).flags.Muerto) Then
                                                        Call RevivirUsuario(nUser)

                                                End If
                         
                                                UserList(nUser).Stats.MinHp = UserList(nUser).Stats.MaxHP
                                                UserList(nUser).Stats.MinMAN = UserList(nUser).Stats.MaxMAN
                                                UserList(nUser).Stats.MinHam = 100
                                                UserList(nUser).Stats.MinAGU = 100
                                                UserList(nUser).Stats.MinSta = UserList(nUser).Stats.MaxSta
                         
                                                Call Protocol.WriteUpdateUserStats(nUser)

                                        End If
                              
                                Else
                              
                                        UserList(nUser).sReto.acceptedOK = False

                                End If

                        End If

                Next Loopc
         
                For Loopc = 0 To 1
                        nUser = .team_array(1).user_Index(Loopc)
              
                        If (nUser <> 0) Then
                                If (UserList(nUser).ConnID <> -1) Then
                                        mPosX = get_pos_x(reto_Index, 2, CInt(Loopc))
                                        mPosY = get_pos_y(reto_Index, 2, CInt(Loopc))
                    
                                        UserList(nUser).sReto.reto_used = True
                                        UserList(nUser).sReto.reto_Index = reto_Index
                                        
                                        Call WarpUserChar(nUser, Mapa_Arenas, mPosX, mPosY, True)
                                        Call Protocol.WritePauseToggle(nUser)
                     
                                        If (respawnUser) Then
                                                If (UserList(nUser).flags.Muerto) Then
                                                        Call RevivirUsuario(nUser)

                                                End If
                         
                                                UserList(nUser).Stats.MinHp = UserList(nUser).Stats.MaxHP
                                                UserList(nUser).Stats.MinMAN = UserList(nUser).Stats.MaxMAN
                                                UserList(nUser).Stats.MinHam = 100
                                                UserList(nUser).Stats.MinAGU = 100
                                                UserList(nUser).Stats.MinSta = UserList(nUser).Stats.MaxSta
                         
                                                Call Protocol.WriteUpdateUserStats(nUser)

                                        End If

                                Else
                                        UserList(nUser).sReto.acceptedOK = False

                                End If

                        End If

                Next Loopc

        End With

End Sub

Private Sub message_reto(ByRef retoStr As retoStruct, ByRef sMessage As String)

        '
        ' @ amishar.-
    
        With retoStr

                Dim i As Long
                Dim j As Long
                Dim u As Integer
         
                For i = 0 To 1
                        For j = 0 To 1
                                u = .team_array(i).user_Index(j)
                 
                                If (u <> 0) Then
                                        If (UserList(u).ConnID <> -1) Then
                                                Call Protocol.WriteConsoleMsg(u, sMessage, FontTypeNames.FONTTYPE_GUILD)
                         
                                        End If

                                End If

                        Next j
                Next i

        End With
    
End Sub

Public Sub user_die_reto(ByVal user_Index As Integer)

        '
        ' @ amishar.-
    
        Dim team_Index As Integer
        Dim user_slot  As Integer
        Dim other_user As Integer
        Dim reto_Index As Integer
    
        reto_Index = UserList(user_Index).sReto.reto_Index
    
        team_Index = find_Team(user_Index, reto_Index)

        If (team_Index <> -1) Then
                user_slot = find_user(team_Index, user_Index, reto_Index)
        Else

                Exit Sub

        End If

        If (user_slot = -1) Then Exit Sub
    
        other_user = IIf(user_slot = 0, 1, 0)
        other_user = reto_List(reto_Index).team_array(team_Index).user_Index(other_user)
    
        'is dead?

        If (other_user) Then
                If UserList(other_user).flags.Muerto Then
                        Call team_winner(reto_Index, IIf(team_Index = 0, 1, 0))

                End If

        Else
                Call team_winner(reto_Index, IIf(team_Index = 0, 1, 0))

        End If
    
End Sub

Public Function find_Team(ByVal user_Index As Integer, _
                          ByVal reto_Index As Integer) As Integer

        '
        ' @ amishar.-
    
        Dim i As Long
        Dim j As Long
    
        For i = 0 To 1
                For j = 0 To 1

                        If reto_List(reto_Index).team_array(i).user_Index(j) = user_Index Then
                                find_Team = i

                                Exit Function

                        End If

                Next j
        Next i

        find_Team = -1
    
End Function

Private Function find_user(ByVal team_Index As Integer, _
                           ByVal user_Index As Integer, _
                           ByVal reto_Index As Integer) As Integer

        '
        ' @ amishar.-
    
        Dim i As Long
    
        For i = 0 To 1

                If reto_List(reto_Index).team_array(team_Index).user_Index(i) = user_Index Then
                        find_user = i

                        Exit Function

                End If

        Next i
    
        find_user = -1

End Function

Private Sub team_winner(ByVal reto_Index As Integer, ByVal team_winner As Byte)

        '
        ' @ amishar.-
    
        With reto_List(reto_Index)
                .team_array(team_winner).round_count = (.team_array(team_winner).round_count + 1)
         
                If (.team_array(team_winner).round_count = 2) Then
                        Call finish_reto(reto_Index, team_winner)
                Else
                        Call respawn_reto(reto_Index, team_winner)

                End If
         
        End With

End Sub

Private Sub respawn_reto(ByVal reto_Index As Integer, ByVal team_winner As Integer)

        '
        ' @ amishar.-
    
        'Call warp_Teams(reto_Index, True)
    
        Dim loopX As Long
        Dim Loopc As Long
        Dim mStr  As String
        Dim Index As Integer
    
        With reto_List(reto_Index)
    
                mStr = "El equipo " & CStr(team_winner + 1) & " gana este duelo." & vbNewLine & "Resultado parcial : " & CStr(.team_array(0).round_count) & "-" & CStr(.team_array(1).round_count)
        
                For loopX = 0 To 1
                        For Loopc = 0 To 1
                                Index = .team_array(loopX).user_Index(Loopc)
                
                                If (Index <> 0) Then
                                        If UserList(Index).ConnID <> -1 Then
                                                Call Protocol.WriteConsoleMsg(Index, mStr, FontTypeNames.FONTTYPE_GUILD)
                                                Call Protocol.WriteConsoleMsg(Index, "El siguiente round iniciará en 3 segundos.", FontTypeNames.FONTTYPE_GUILD)

                                        End If

                                End If

                        Next Loopc
                Next loopX
        
                .nextRoundCount = 3
        
        End With
    
End Sub

Private Sub finish_reto(ByVal reto_Index As Integer, ByVal team_winner As Byte)

        '
        ' @ amishar.-
    
        With reto_List(reto_Index)
         
                Dim retoMessage As String
                Dim team_looser As Byte
                Dim temp_index  As Integer
         
                retoMessage = get_reto_message(reto_Index)
         
                retoMessage = retoMessage & vbNewLine & "Reto 2vs2> Ganador equipo " & CStr(team_winner + 1) & "."
         
                Call SendData(SendTarget.ToAll, 0, Protocol.PrepareMessageConsoleMsg(retoMessage, FontTypeNames.FONTTYPE_GUILD))
         
                team_looser = IIf(team_winner = 0, 1, 0)
         
                Dim Loopc  As Long
                Dim bydrop As Boolean
                Dim byGold As Long
         
                bydrop = (.general_rules.drop_inv = True)
                byGold = .general_rules.gold_gamble
         
                With .team_array(team_looser)

                        For Loopc = 0 To 1
                                temp_index = .user_Index(Loopc)
                  
                                UserList(temp_index).sReto.reto_used = False
                                UserList(temp_index).sReto.acceptedOK = False
                   
                                If (bydrop) Then
                                        Call TirarTodosLosItems(temp_index)

                                End If
                                    
                                Call WarpUserChar(temp_index, 1, 50 + Loopc, 50, True)
                                    
                                UserList(temp_index).Stats.GLD = (UserList(temp_index).Stats.GLD - byGold)
      
                                UserList(temp_index).sReto.nick_sender = vbNullString
                                UserList(temp_index).sReto.reto_Index = 0
                  
                                Call Protocol.WriteUpdateGold(temp_index)
     
                        Next Loopc

                End With
         
                With .team_array(team_winner)

                        For Loopc = 0 To 1
                                temp_index = .user_Index(Loopc)
                  
                                UserList(temp_index).sReto.reto_used = False
                                UserList(temp_index).sReto.acceptedOK = False
                   
                                If (bydrop) Then
                                        UserList(temp_index).sReto.return_city = 15
                                        reto_List(reto_Index).haydrop = True
                                        
                                        Call Protocol.WriteConsoleMsg(temp_index, "Regresarás a tu hogar en 15 segundos.", FontTypeNames.FONTTYPE_GUILD)
                                Else
                                        Call WarpUserChar(temp_index, 1, 50 + Loopc, 50, True)

                                End If
                  
                                UserList(temp_index).Stats.GLD = (UserList(temp_index).Stats.GLD + byGold)

                                If reto_List(reto_Index).haydrop Then
                                
                                        UserList(temp_index).sReto.nick_sender = vbNullString
                                        UserList(temp_index).sReto.reto_Index = 0
 
                                End If

                                Call Protocol.WriteUpdateGold(temp_index)
          
                        Next Loopc

                End With

                If .haydrop Then
                        Call clear_data(reto_Index)

                End If

        End With

End Sub

Private Sub clear_data(ByVal reto_Index As Integer)

        '
        ' @ amishar.-
    
        With reto_List(reto_Index)
        
                .haydrop = False
                .count_Down = 0
         
                With .general_rules
                        .drop_inv = False
                        .gold_gamble = 0

                End With
         
                .used_ring = False
       
                Dim i As Long
         
                For i = 0 To 1
                        .team_array(i).user_Index(0) = 0
                        .team_array(i).user_Index(1) = 0
             
                        .team_array(i).round_count = 0

                Next i
         
        End With

End Sub

Private Function get_reto_message(ByVal reto_Index As Integer) As String

        '
        ' @ amishar.-
    
        Dim TempStr  As String
        Dim tempUser As Integer
    
        With reto_List(reto_Index)
         
                TempStr = "Retos> "
         
                With .team_array(0)
                        tempUser = .user_Index(0)
              
                        If (tempUser <> 0) Then
                                If UserList(tempUser).ConnID <> -1 Then
                                        TempStr = TempStr & UserList(tempUser).Name

                                End If

                        End If
              
                        tempUser = .user_Index(1)
              
                        If (tempUser <> 0) Then
                                If UserList(tempUser).ConnID <> -1 Then
                                        TempStr = TempStr & " y " & UserList(tempUser).Name

                                End If

                        End If
              
                End With
         
                With .team_array(1)
                        tempUser = .user_Index(0)
              
                        If (tempUser <> 0) Then
                                If UserList(tempUser).ConnID <> -1 Then
                                        TempStr = TempStr & " vs " & UserList(tempUser).Name

                                End If

                        End If
              
                        tempUser = .user_Index(1)
              
                        If (tempUser <> 0) Then
                                If UserList(tempUser).ConnID <> -1 Then
                                        TempStr = TempStr & " y " & UserList(tempUser).Name

                                End If

                        End If
              
                End With
         
                With .general_rules
                     
                        TempStr = TempStr & " con apuesta de " & Format$(.gold_gamble, "#,###") & " monedas de oro"
              
                        If (.drop_inv) Then
                                TempStr = TempStr & " y los items del inventario"

                        End If

                End With
         
        End With

End Function

Private Function get_pos_x(ByVal ring_Index As Integer, _
                          ByVal team_Index As Integer, _
                          ByVal user_Index As Integer) As Integer

        '
        ' @ amishar.-
    
        Dim endPos As Integer
        
        endPos = retoPos(ring_Index, team_Index, user_Index + 1).X
    
        get_pos_x = endPos
    
End Function

Private Function get_pos_y(ByVal ring_Index As Integer, _
                          ByVal team_Index As Integer, _
                          ByVal user_Index As Integer) As Integer

        '
        ' @ amishar.-
    
        Dim endPos As Integer
    
        endPos = retoPos(ring_Index, team_Index, user_Index + 1).Y
    
        get_pos_y = endPos
    
End Function

Public Sub retos2vs2Load()

        '
        ' @ amishar.-
    
        Dim nArenas As Integer

        Dim bReader As New clsIniManager
    
        bReader.Initialize DatPath & "Retos2vs2.ini"
    
        nArenas = val(bReader.GetValue("INIT", "Arenas"))
    
        If (nArenas = 0) Then Exit Sub
    
        ReDim Mod_Retos2vs2.retoPos(1 To nArenas, 1 To 2, 1 To 2) As Mod_Retos2vs2.retoPosStruct
        ReDim Mod_Retos2vs2.reto_List(1 To nArenas) As Mod_Retos2vs2.retoStruct
    
        Dim i   As Long
        Dim j   As Long
        Dim p   As Long
        Dim s   As String
        Dim tmp As Integer
                
        Mapa_Arenas = Configuracion.Mapa2vs2
        
        tmp = Asc("-")
    
        For i = 1 To nArenas
                For j = 1 To 2
                        For p = 1 To 2
                                s = bReader.GetValue("ARENA" & CStr(i), "Equipo" & CStr(j) & "Jugador" & CStr(p))
                
                                Mod_Retos2vs2.retoPos(i, j, p).X = val(ReadField(2, s, tmp))
                                Mod_Retos2vs2.retoPos(i, j, p).Y = val(ReadField(3, s, tmp))
                
                        Next p
                Next j
        Next i

End Sub

Public Function eventAttack(ByVal attackerIndex As Integer, _
                     ByVal victimIndex As Integer) As Boolean
 
        '
        ' @ amishar
   
        If UserList(attackerIndex).sReto.reto_used = True Then
                If Mod_Retos2vs2.can_Attack(attackerIndex, victimIndex) = False Then
                        Call WriteConsoleMsg(attackerIndex, "No puedes atacar a tu compañero.", FontTypeNames.FONTTYPE_INFO)
                        eventAttack = False
 
                        Exit Function
 
                End If

        End If
   
        eventAttack = True
 
End Function

