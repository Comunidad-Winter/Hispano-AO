Attribute VB_Name = "Mod_General"
Option Explicit

Private Type tNames

        Nombre As String
        GrhIndex As Integer
        ObjType As Byte
        
        MinHit As Integer
        MaxHit As Integer
        
        Valor As Long
        
        MinDef As Integer
        MaxDef As Integer
        
        DefensaMagicaMin As Integer
        DefensaMagicaMax As Integer
        
        LingH As Integer
        LingP As Integer
        LingO As Integer
        
        Upgrade As Integer
        Madera As Integer
        MaderaElfica As Integer
            
End Type

Public Enum enum_TypeObj
        
        Nombre = 1
        GrhIndex = 2
        ObjType = 3
        
        MinHit = 4
        MaxHit = 5
        
        Valor = 6
        
        MinDef = 7
        MaxDef = 8
        
        DefensaMagicaMin = 9
        DefensaMagicaMax = 10
        
        LingH = 11
        LingP = 12
        LingO = 13
        
        Upgrade = 14
        Madera = 15
        MaderaElfica = 16
        
End Enum

Private Obj()        As tNames

Private Npcs()       As tNames

Private Spells()     As tNames

Public CurServerIp   As String

Public CurServerPort As Integer

Public bFogata       As Boolean

Private lFrameTimer  As Long

Public Function getTypeObj(ByVal ObjIndex As Integer, _
                           ByVal Tipe As enum_TypeObj) As String

        If ObjIndex < 1 Or ObjIndex > UBound(Obj) Then
               
                'Mod_General.AddtoRichTextBox frmMain.RecTxt, "Error: OBJINDEX " & ObjIndex & " no encontrado, reportar a un administrador"
                
                Exit Function

        End If

        Select Case Tipe

                Case enum_TypeObj.DefensaMagicaMax
                        getTypeObj = CStr(Obj(ObjIndex).DefensaMagicaMax)

                Case enum_TypeObj.DefensaMagicaMin
                        getTypeObj = CStr(Obj(ObjIndex).DefensaMagicaMin)

                Case enum_TypeObj.GrhIndex
                        getTypeObj = CStr(Obj(ObjIndex).GrhIndex)

                Case enum_TypeObj.MaxDef
                        getTypeObj = CStr(Obj(ObjIndex).MaxDef)

                Case enum_TypeObj.MaxHit
                        getTypeObj = CStr(Obj(ObjIndex).MaxHit)

                Case enum_TypeObj.MinDef
                        getTypeObj = CStr(Obj(ObjIndex).MinDef)

                Case enum_TypeObj.MinHit
                        getTypeObj = CStr(Obj(ObjIndex).MinHit)

                Case enum_TypeObj.Nombre
                        getTypeObj = CStr(Obj(ObjIndex).Nombre)

                Case enum_TypeObj.ObjType
                        getTypeObj = CStr(Obj(ObjIndex).ObjType)

                Case enum_TypeObj.Valor
                        getTypeObj = CStr(Obj(ObjIndex).Valor)
                        
                Case enum_TypeObj.LingH
                        getTypeObj = CStr(Obj(ObjIndex).LingH)

                Case enum_TypeObj.LingP
                        getTypeObj = CStr(Obj(ObjIndex).LingP)

                Case enum_TypeObj.LingO
                        getTypeObj = CStr(Obj(ObjIndex).LingO)

                Case enum_TypeObj.Upgrade
                        getTypeObj = CStr(Obj(ObjIndex).Upgrade)

                Case enum_TypeObj.Madera
                        getTypeObj = CStr(Obj(ObjIndex).Madera)

                Case enum_TypeObj.MaderaElfica
                        getTypeObj = CStr(Obj(ObjIndex).MaderaElfica)

        End Select

End Function

Public Function getNameNpcs(ByVal NpcIndex As Integer) As String

        If NpcIndex < 1 Or NpcIndex > UBound(Npcs) Then Exit Function
        getNameNpcs = Npcs(NpcIndex).Nombre

End Function

Public Function getNameHechizo(ByVal SpellIndex As Integer) As String

        If SpellIndex < 1 Or SpellIndex > UBound(Spells) Then Exit Function
        getNameHechizo = Spells(SpellIndex).Nombre

End Function

Private Sub LoadNameSource()

        Dim Leer    As clsIniManager
        
        Dim tmpCant As Integer

        Dim LoopC   As Long

        ' @@ Objetos
        Set Leer = New clsIniManager
        
        Leer.Initialize DirExtras & "Obj.dat"
        tmpCant = Val(Leer.GetValue("Init", "NumOBJs"))
        
        ReDim Obj(1 To tmpCant) As tNames

        For LoopC = 1 To tmpCant
                Obj(LoopC).Nombre = Leer.GetValue("OBJ" & LoopC, "Name")
                Obj(LoopC).GrhIndex = Val(Leer.GetValue("OBJ" & LoopC, "GrhIndex"))
                
                Obj(LoopC).ObjType = Val(Leer.GetValue("OBJ" & LoopC, "ObjType"))
                Obj(LoopC).Valor = Val(Leer.GetValue("OBJ" & LoopC, "Valor"))
                  
                Obj(LoopC).MinHit = Val(Leer.GetValue("OBJ" & LoopC, "MinHit"))
                Obj(LoopC).MaxHit = Val(Leer.GetValue("OBJ" & LoopC, "MaxHit"))
        
                Obj(LoopC).MinDef = Val(Leer.GetValue("OBJ" & LoopC, "MinDef"))
                Obj(LoopC).MaxDef = Val(Leer.GetValue("OBJ" & LoopC, "MaxDef"))
        
                Obj(LoopC).DefensaMagicaMin = Val(Leer.GetValue("OBJ" & LoopC, "DefensaMagicaMin"))
                Obj(LoopC).DefensaMagicaMax = Val(Leer.GetValue("OBJ" & LoopC, "DefensaMagicaMax"))
                
                Obj(LoopC).LingH = Val(Leer.GetValue("OBJ" & LoopC, "LingH"))
                Obj(LoopC).LingP = Val(Leer.GetValue("OBJ" & LoopC, "LingP"))
                Obj(LoopC).LingO = Val(Leer.GetValue("OBJ" & LoopC, "LingO"))
        
                Obj(LoopC).Upgrade = Val(Leer.GetValue("OBJ" & LoopC, "Upgrade"))
                Obj(LoopC).Madera = Val(Leer.GetValue("OBJ" & LoopC, "Madera"))
                Obj(LoopC).MaderaElfica = Val(Leer.GetValue("OBJ" & LoopC, "MaderaElfica"))
        
        Next LoopC
        
        ' @@ Hechizos
        Set Leer = New clsIniManager
        
        Leer.Initialize DirExtras & "Hechizos.dat"
        tmpCant = Val(Leer.GetValue("Init", "NumeroHechizos"))
        
        ReDim Spells(1 To tmpCant) As tNames

        For LoopC = 1 To tmpCant
                Spells(LoopC).Nombre = Leer.GetValue("HECHIZO" & LoopC, "Nombre")
        Next LoopC
        
        ' @@ Npc's
        Set Leer = New clsIniManager
        
        Leer.Initialize DirExtras & "Npcs.dat"
        tmpCant = Val(Leer.GetValue("Init", "NumNPCs"))
        
        ReDim Npcs(1 To tmpCant) As tNames

        For LoopC = 1 To tmpCant
                Npcs(LoopC).Nombre = Leer.GetValue("NPC" & LoopC, "Name")
        Next LoopC

        Set Leer = Nothing

End Sub

'>> Funciones/Subs
Sub Analizar()

        'On Error Resume Next
            
        Dim iX   As Integer

        Dim tX   As Integer

        Dim DifX As Integer
            
        'LINK1            Variable que contiene el numero de actualización correcto del servidor
        'iX = frmMain.Inet1.OpenURL("http://hispanoao.net/actualizar/VEREXE.txt")
        'iX = frmConnect.Inet1.OpenURL("http://hispano-ao.com/actualizar/VEREXE.txt")
        
        'Variable que contiene el numero de actualización del cliente
        'tX = GetVar(DirInit & "Update.ini", "INIT", "X")
        
        'Variable con la diferencia de actualizaciones servidor-cliente
        'DifX = iX - tX
            
        'If Not (DifX = 0) Then 'Si la diferencia no es nula,
                'MsgBox "Se detectaron actualizaciones pendientes. A continuación se ejecutará el autoupdate."
               ' Shell App.path & "\autoUpdate.exe", vbNormalFocus

                'End

        'End If
            End Sub

Public Function DirInterfaces() As String
        DirInterfaces = App.path & "\Interfaces\"

End Function

Public Function DirInit() As String
        DirInit = App.path & "\init\"

End Function

Public Function DirGraficos() As String
        DirGraficos = App.path & "\Graficos\"

End Function

Public Function DirSound() As String
        DirSound = App.path & "\Wav\"

End Function

Public Function DirMidi() As String
        DirMidi = App.path & "\Midi\"

End Function

Public Function DirMapas() As String
        DirMapas = App.path & "\MAPAS\"

End Function

Public Function DirExtras() As String
        DirExtras = App.path & "\EXTRAS\"

End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
        'Initialize randomizer
        Randomize Timer
    
        'Generate random number
        RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound

End Function

Public Function GetRawName(ByRef sName As String) As String
        '***************************************************
        'Author: ZaMa
        'Last Modify Date: 13/01/2010
        'Last Modified By: -
        'Returns the char name without the clan name (if it has it).
        '***************************************************

        Dim Pos As Integer
    
        Pos = InStr(1, sName, "<")
    
        If Pos > 0 Then
                GetRawName = Trim(Left(sName, Pos - 1))
        Else
                GetRawName = sName

        End If

End Function

Sub CargarAnimArmas()

        On Error Resume Next

        Dim LoopC As Long

        Dim arch  As String
    
        arch = DirInit & "armas.dat"
    
        NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
    
        ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
        For LoopC = 1 To NumWeaponAnims
                InitGrh WeaponAnimData(LoopC).WeaponWalk(1), Val(GetVar(arch, "ARMA" & LoopC, "Dir1")), 0
                InitGrh WeaponAnimData(LoopC).WeaponWalk(2), Val(GetVar(arch, "ARMA" & LoopC, "Dir2")), 0
                InitGrh WeaponAnimData(LoopC).WeaponWalk(3), Val(GetVar(arch, "ARMA" & LoopC, "Dir3")), 0
                InitGrh WeaponAnimData(LoopC).WeaponWalk(4), Val(GetVar(arch, "ARMA" & LoopC, "Dir4")), 0
        Next LoopC

End Sub

Sub CargarColores()
 
        ' @@ Miqueas : Nuevo sistema de color de nicks
        
        Dim archivoC As String

        archivoC = DirInit & "colores.dat"
 
        If Not FileExist(archivoC, vbArchive) Then
                'TODO : Si hay que reinstalar, porque no cierra???
                Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
 
                Exit Sub
 
        End If
 
        Dim i As Long
 
        For i = 0 To 48 '49 y 50 reservados para ciudadano y criminal
                ColoresPJ(i) = RGB(CByte(GetVar(archivoC, i, "R")), CByte(GetVar(archivoC, i, "G")), CByte(GetVar(archivoC, i, "B")))
 
        Next i
 
        ' Crimi
        ColoresPJ(50) = RGB(CByte(GetVar(archivoC, "CR", "R")), CByte(GetVar(archivoC, "CR", "G")), CByte(GetVar(archivoC, "CR", "B")))
 
        ' Ciuda
        ColoresPJ(49) = RGB(CByte(GetVar(archivoC, "CI", "R")), CByte(GetVar(archivoC, "CI", "G")), CByte(GetVar(archivoC, "CI", "B")))
 
        ' Atacable
        ColoresPJ(48) = RGB(CByte(GetVar(archivoC, "AT", "R")), CByte(GetVar(archivoC, "AT", "G")), CByte(GetVar(archivoC, "AT", "B")))

End Sub

Sub CargarAnimEscudos()

        On Error Resume Next

        Dim LoopC As Long

        Dim arch  As String
    
        arch = DirInit & "escudos.dat"
    
        NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
        ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
        For LoopC = 1 To NumEscudosAnims
                InitGrh ShieldAnimData(LoopC).ShieldWalk(1), Val(GetVar(arch, "ESC" & LoopC, "Dir1")), 0
                InitGrh ShieldAnimData(LoopC).ShieldWalk(2), Val(GetVar(arch, "ESC" & LoopC, "Dir2")), 0
                InitGrh ShieldAnimData(LoopC).ShieldWalk(3), Val(GetVar(arch, "ESC" & LoopC, "Dir3")), 0
                InitGrh ShieldAnimData(LoopC).ShieldWalk(4), Val(GetVar(arch, "ESC" & LoopC, "Dir4")), 0
        Next LoopC

End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, _
                     ByVal Text As String, _
                     Optional ByVal red As Integer = -1, _
                     Optional ByVal green As Integer, _
                     Optional ByVal blue As Integer, _
                     Optional ByVal bold As Boolean = False, _
                     Optional ByVal italic As Boolean = False, _
                     Optional ByVal bCrLf As Boolean = True)
                     
        '******************************************
        'Adds text to a Richtext box at the bottom.
        'Automatically scrolls to new text.
        'Text box MUST be multiline and have a 3D
        'apperance!
        'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
        'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
        '08/02/12 (D'Artagnan) - División de consolas
        '******************************************r

        With RichTextBox
     
                If Len(.Text) > 1000 Then
                        'Get rid of first line
                        .SelStart = InStr(1, .Text, vbCrLf) + 1
                        .SelLength = Len(.Text) - .SelStart + 2
                        .TextRTF = .SelRTF

                End If
                
                .SelStart = Len(.Text)
                .SelLength = 0
                .SelBold = bold
                .SelItalic = italic
                
                If Not red = -1 Then .SelColor = RGB(red, green, blue)
                
                If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
                .SelText = Text
                .Refresh

        End With

End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()

        '*****************************************************************
        'Goes through the charlist and replots all the characters on the map
        'Used to make sure everyone is visible
        '*****************************************************************
        Dim LoopC As Long
    
        For LoopC = 1 To LastChar

                If charlist(LoopC).Active = 1 Then
                        MapData(charlist(LoopC).Pos.x, charlist(LoopC).Pos.y).CharIndex = LoopC

                End If

        Next LoopC

End Sub

Function AsciiValidos(ByVal cad As String) As Boolean

        Dim car As Byte
        Dim i   As Long
    
        cad = LCase$(cad)
    
        For i = 1 To Len(cad)
                car = Asc(mid$(cad, i, 1))
        
                If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then

                        Exit Function

                End If

        Next i
    
        AsciiValidos = True

End Function

Function CheckUserData() As Boolean

        'Validamos los datos del user
        Dim LoopC     As Long

        Dim CharAscii As Integer
        
        If LenB(UserPassword) = 0 Then
                MsgBox ("Ingrese un password.")
                Exit Function

        End If
    
        For LoopC = 1 To Len(UserPassword)
                CharAscii = Asc(mid$(UserPassword, LoopC, 1))

                If Not LegalCharacter(CharAscii) Then
                        MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
                        Exit Function

                End If

        Next LoopC
    
        If LenB(Username) = 0 Then
                MsgBox ("Ingrese un nombre de personaje.")
                Exit Function

        End If
    
        If Len(Username) > 30 Then
                MsgBox ("El nombre debe tener menos de 30 letras.")
                Exit Function

        End If
    
        For LoopC = 1 To Len(Username)
                CharAscii = Asc(mid$(Username, LoopC, 1))

                If Not LegalCharacter(CharAscii) Then
                        MsgBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
                        Exit Function

                End If

        Next LoopC
    
        CheckUserData = True

End Function

Sub UnloadAllForms()

        On Error Resume Next

        Dim mifrm As Form
    
        For Each mifrm In Forms

                Unload mifrm
        Next

End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
        '*****************************************************************
        'Only allow characters that are Win 95 filename compatible
        '*****************************************************************
        'if backspace allow

        If KeyAscii = 8 Then
                LegalCharacter = True

                Exit Function

        End If
    
        'Only allow space, numbers, letters and special characters

        If KeyAscii < 32 Or KeyAscii = 44 Then

                Exit Function

        End If
    
        If KeyAscii > 126 Then

                Exit Function

        End If
    
        'Check for bad special characters in between

        If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then

                Exit Function

        End If
    
        'else everything is cool
        LegalCharacter = True

End Function

Sub SetConnected()
        '*****************************************************************
        'Sets the client to "Connect" mode
        '*****************************************************************
        
        'Set Connected
        Connected = True
    
        'Unload the connect form
        Unload frmCrearPersonaje
        Unload frmConnect
    
        frmMain.lblName.Caption = Username
        
        'Load main form
        frmMain.Visible = True
    
        Call frmMain.ControlSM(eSMType.mSpells, False)
        Call frmMain.ControlSM(eSMType.mWork, False)
    
        FPSFLAG = True

End Sub

Sub MoveTo(ByVal Direccion As E_Heading)

        '***************************************************
        'Author: Alejandro Santos (AlejoLp)
        'Last Modify Date: 06/28/2008
        'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
        ' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
        ' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
        ' 06/28/2008: NicoNZ - Saqué lo que impedía que si el usuario estaba paralizado se ejecute el sub.
        '***************************************************
        Dim LegalOk As Boolean
    
        If Cartel Then Cartel = False
    
        Select Case Direccion

                Case E_Heading.NORTH
                        LegalOk = MoveToLegalPos(UserPos.x, UserPos.y - 1)

                Case E_Heading.EAST
                        LegalOk = MoveToLegalPos(UserPos.x + 1, UserPos.y)

                Case E_Heading.SOUTH
                        LegalOk = MoveToLegalPos(UserPos.x, UserPos.y + 1)

                Case E_Heading.WEST
                        LegalOk = MoveToLegalPos(UserPos.x - 1, UserPos.y)

        End Select
    
        If LegalOk And Not UserParalizado Then
           
                Call WriteWalk(Direccion)

                If Not UserDescansar And Not UserMeditar Then
                        MoveCharbyHead UserCharIndex, Direccion
                        MoveScreen Direccion

                End If

        Else

                If charlist(UserCharIndex).Heading <> Direccion Then
                        Call WriteChangeHeading(Direccion)

                End If

        End If
    
        If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
    
        ' Update 3D sounds!
        Call Audio.MoveListener(UserPos.x, UserPos.y)

End Sub

Sub RandomMove()
        '***************************************************
        'Author: Alejandro Santos (AlejoLp)
        'Last Modify Date: 06/03/2006
        ' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
        '***************************************************
        Call MoveTo(RandomNumber(NORTH, WEST))

End Sub

Private Sub CheckKeys()

        '*****************************************************************
        'Checks keys and respond
        '*****************************************************************
        Static LastMovement As Long
    
        'No input allowed while Argentum is not the active window
        If Not Application.IsAppActive() Then Exit Sub
    
        'No walking when in commerce or banking.
        If Comerciando Then Exit Sub
    
        'No walking while writting in the forum.
        If MirandoForo Then Exit Sub
    
        'If game is paused, abort movement.
        If pausa Then Exit Sub
    
        'TODO: Debería informarle por consola?
        If Traveling Then Exit Sub

        'Control movement interval (this enforces the 1 step loss when meditating / resting client-side)
        If GetTickCount - LastMovement > 56 Then
                LastMovement = GetTickCount
        Else
                Exit Sub

        End If
    
        'Don't allow any these keys during movement..
        If UserMoving = 0 Then
    
                If Not UserEstupido Then

                        'Move Up
                        If (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Then
                             
                                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                                Call MoveTo(NORTH)
                                frmMain.Coord.Caption = "Mapa: " & UserMap & " X: " & UserPos.x & " Y: " & UserPos.y
                                Exit Sub

                        End If
            
                        'Move Right
                        If (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0) Then
                                
                                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                                Call MoveTo(EAST)
                                frmMain.Coord.Caption = "Mapa: " & UserMap & " X: " & UserPos.x & " Y: " & UserPos.y
                                Exit Sub

                        End If
        
                        'Move down
                        If (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0) Then
                               
                                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                                Call MoveTo(SOUTH)
                                frmMain.Coord.Caption = "Mapa: " & UserMap & " X: " & UserPos.x & " Y: " & UserPos.y
                                Exit Sub

                        End If
        
                        'Move left
                        If (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0) Then
                             
                                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                                Call MoveTo(WEST)
                                frmMain.Coord.Caption = "Mapa: " & UserMap & " X: " & UserPos.x & " Y: " & UserPos.y
                                Exit Sub

                        End If
            
                        ' We haven't moved - Update 3D sounds!
                        Call Audio.MoveListener(UserPos.x, UserPos.y)
                Else

                        Dim kp As Boolean

                        kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
                           GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                           GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
                           GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            
                        If kp Then
                                Call RandomMove
                        Else
                                ' We haven't moved - Update 3D sounds!
                                Call Audio.MoveListener(UserPos.x, UserPos.y)

                        End If
            
                        If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
     
                        frmMain.Coord.Caption = "X: " & UserPos.x & " Y: " & UserPos.y

                End If

        End If

End Sub

'TODO : Si bien nunca estuvo allí, el mapa es algo independiente o a lo sumo dependiente del engine, no va acá!!!
Sub SwitchMap(ByVal Map As Integer)

        '**************************************************************
        'Formato de mapas optimizado para reducir el espacio que ocupan.
        'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
        '**************************************************************
        Dim y       As Long

        Dim x       As Long

        Dim tempint As Integer

        Dim ByFlags As Byte

        Dim Handle  As Integer
    
        Handle = FreeFile()
    
        Open DirMapas & "Mapa" & Map & ".map" For Binary As Handle
        Seek Handle, 1
            
        'map Header
        Get Handle, , MapInfo.MapVersion
        Get Handle, , MiCabecera
        Get Handle, , tempint
        Get Handle, , tempint
        Get Handle, , tempint
        Get Handle, , tempint
    
        'Load arrays
        For y = YMinMapSize To YMaxMapSize
                For x = XMinMapSize To XMaxMapSize
                        Get Handle, , ByFlags
            
                        MapData(x, y).Blocked = (ByFlags And 1)
            
                        Get Handle, , MapData(x, y).Graphic(1).GrhIndex
                        InitGrh MapData(x, y).Graphic(1), MapData(x, y).Graphic(1).GrhIndex
            
                        'Layer 2 used?
                        If ByFlags And 2 Then
                                Get Handle, , MapData(x, y).Graphic(2).GrhIndex
                                InitGrh MapData(x, y).Graphic(2), MapData(x, y).Graphic(2).GrhIndex
                        Else
                                MapData(x, y).Graphic(2).GrhIndex = 0

                        End If
                
                        'Layer 3 used?
                        If ByFlags And 4 Then
                                Get Handle, , MapData(x, y).Graphic(3).GrhIndex
                                InitGrh MapData(x, y).Graphic(3), MapData(x, y).Graphic(3).GrhIndex
                        Else
                                MapData(x, y).Graphic(3).GrhIndex = 0

                        End If
                
                        'Layer 4 used?
                        If ByFlags And 8 Then
                                Get Handle, , MapData(x, y).Graphic(4).GrhIndex
                                InitGrh MapData(x, y).Graphic(4), MapData(x, y).Graphic(4).GrhIndex
                        Else
                                MapData(x, y).Graphic(4).GrhIndex = 0

                        End If
            
                        'Trigger used?
                        If ByFlags And 16 Then
                                Get Handle, , MapData(x, y).Trigger
                        Else
                                MapData(x, y).Trigger = 0

                        End If
            
                        'Erase NPCs
                        If MapData(x, y).CharIndex > 0 Then
                                Call EraseChar(MapData(x, y).CharIndex)

                        End If
            
                        'Erase OBJs
                        MapData(x, y).ObjGrh.GrhIndex = 0
                Next x
        Next y
    
        Close Handle
    
        MapInfo.Name = vbNullString
        MapInfo.Music = vbNullString
    
        CurMap = Map

End Sub

Function ReadField(ByVal Pos As Integer, _
                   ByRef Text As String, _
                   ByVal SepASCII As Byte) As String

        '*****************************************************************
        'Gets a field from a delimited string
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 11/15/2004
        '*****************************************************************
        Dim i          As Long

        Dim lastPos    As Long

        Dim CurrentPos As Long

        Dim delimiter  As String * 1
    
        delimiter = Chr$(SepASCII)
    
        For i = 1 To Pos
                lastPos = CurrentPos
                CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
        Next i
    
        If CurrentPos = 0 Then
                ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
        Else
                ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)

        End If

End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long

        '*****************************************************************
        'Gets the number of fields in a delimited string
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 07/29/2007
        '*****************************************************************
        Dim Count     As Long

        Dim curPos    As Long

        Dim delimiter As String * 1
    
        If LenB(Text) = 0 Then Exit Function
    
        delimiter = Chr$(SepASCII)
    
        curPos = 0
    
        Do
                curPos = InStr(curPos + 1, Text, delimiter)
                Count = Count + 1
        Loop While curPos <> 0
    
        FieldCount = Count

End Function

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean

        FileExist = (LenB(Dir$(File, FileType)) <> 0)

End Function

Sub WriteClientVer()

        Dim hFile As Integer
        
        hFile = FreeFile()
        Open DirInit & "Ver.bin" For Binary Access Write Lock Read As #hFile
        Put #hFile, , CLng(777)
        Put #hFile, , CLng(777)
        Put #hFile, , CLng(777)
    
        Put #hFile, , CInt(App.Major)
        Put #hFile, , CInt(App.Minor)
        Put #hFile, , CInt(App.Revision)
    
        Close #hFile

End Sub

Private Sub LoadSettings()

        CurServerIp = GetVar(DirInit & "Configuracion.ini", "INIT", "IP")
        CurServerPort = GetVar(DirInit & "Configuracion.ini", "INIT", "Puerto")

End Sub

Sub Main()

        Call WriteClientVer
        
        Call Analizar
        
        'Load ao.dat config file
        Call LoadClientSetup
    
        If ClientSetup.bDinamic Then
                Set SurfaceDB = New clsSurfaceManDyn
        Else
                Set SurfaceDB = New clsSurfaceManStatic

        End If
 
        #If Testeo = 0 Then

                If FindPreviousInstance Then
                        Call MsgBox("Hispano AO ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
                        End

                End If

        #End If
    
        ChDrive App.path
        ChDir App.path
        
        'Set resolution BEFORE the loading form is displayed, therefore it will be centered.
        Call Resolution.SetResolution
    
        ' Load constants, classes, flags, graphics..
        LoadInitialConfig

        frmMain.Socket1.Startup
      
        frmConnect.Visible = True
    
        'Inicialización de variables globales
        prgRun = True
        pausa = False
    
        ' Intervals
        LoadTimerIntervals
            
        'Set the dialog's font
        Dialogos.font = frmMain.font
        DialogosClanes.font = frmMain.font
        
        Call Mod_rDamage.Initialize
    
        lFrameTimer = GetTickCount
    
        ' Load the form for screenshots
        Call Load(frmScreenshots)
        
        Do While prgRun

                'Sólo dibujamos si la ventana no está minimizada
                If frmMain.WindowState <> 1 And frmMain.Visible Then
                        Call ShowNextFrame(frmMain.Top, frmMain.Left)
            
                        'Play ambient sounds
                        Call RenderSounds
            
                        Call CheckKeys

                End If

                'FPS Counter - mostramos las FPS
                If GetTickCount - lFrameTimer >= 1000 Then
                        If FPSFLAG Then frmMain.lblFPS.Caption = Mod_TileEngine.FPS
            
                        lFrameTimer = GetTickCount

                End If
                
                ' If there is anything to be sent, we send it
                Call FlushBuffer
        
                DoEvents
        Loop
    
        Call CloseClient

End Sub

Private Sub LoadInitialConfig()
        '***************************************************
        'Author: ZaMa
        'Last Modification: 15/03/2011
        '15/03/2011: ZaMa - Initialize classes lazy way.
        '***************************************************
        
        frmCargando.Show
        frmCargando.Refresh

        frmConnect.version = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
    
        '###########
        ' SERVIDORES
        Call AddtoRichTextBox(frmCargando.status, "Cargando Mensajes... ", 255, 255, 255, True, False, True)
        
        Call LoadSettings
        
        Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
        '###########
        ' CONSTANTES
        Call AddtoRichTextBox(frmCargando.status, "Iniciando constantes... ", 255, 255, 255, True, False, True)
        
        Call InicializarNombres
        ' Initialize FONTTYPES
        Call Protocol.InitFonts
    
        With frmConnect
                .txtNombre = ""
                .txtNombre.SelStart = 0
                .txtNombre.SelLength = Len(.txtNombre)

        End With
    
        UserMap = 1
    
        ' Mouse Pointer (Loaded before opening any form with buttons in it)
        If FileExist(DirExtras & "Hand.ico", vbArchive) Then _
           Set picMouseIcon = LoadPicture(DirExtras & "Hand.ico")
        Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
        '#######
        ' CLASES
        Call AddtoRichTextBox(frmCargando.status, "Instanciando clases... ", 255, 255, 255, True, False, True)
        
        Set Dialogos = New clsDialogs
        Set Audio = New clsAudio
        Set Inventario = New clsGrapchicalInventory
        Set CustomKeys = New clsCustomKeys
        Set CustomMessages = New clsCustomMessages
        Set incomingData = New clsByteQueue
        Set outgoingData = New clsByteQueue
        Set MainTimer = New clsTimer
        Set clsForos = New clsForum
        Set DirectX = New DirectX7
        Set Encriptacion = New clsCripto
        
        Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
        '##############
        ' MOTOR GRÁFICO
        Call AddtoRichTextBox(frmCargando.status, "Iniciando motor gráfico... ", 255, 255, 255, True, False, True)
    
        If Not InitTileEngine(frmMain.hWnd, 153, 5, 32, 32, 13, 17, 9, 8, 8, 0.018) Then
                Call CloseClient

        End If

        Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
        '###################
        ' ANIMACIONES EXTRAS
        Call AddtoRichTextBox(frmCargando.status, "Creando animaciones extra... ", 255, 255, 255, True, False, True)
        
        Call LoadNameSource
        Call CargarAnimArmas
        Call CargarAnimEscudos
        Call CargarColores
                
        Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
        '#############
        ' DIRECT SOUND
        Call AddtoRichTextBox(frmCargando.status, "Iniciando DirectSound... ", 255, 255, 255, True, False, True)
        
        'Inicializamos el sonido
        Call Audio.Initialize(DirectX, frmMain.hWnd, DirSound, DirMidi)
        
        'Enable / Disable audio
        Audio.MusicActivated = Not ClientSetup.bNoMusic
        Audio.SoundActivated = Not ClientSetup.bNoSound
        Audio.SoundEffectsActivated = Not ClientSetup.bNoSoundEffects
        
        'Inicializamos el inventario gráfico
        Call Inventario.Initialize(DirectDraw, frmMain.PicInv, MAX_INVENTORY_SLOTS)
        Call Audio.MusicMP3Play(App.path & "\MP3\" & MP3_Inicio & ".mp3")
    
        Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
        Call AddtoRichTextBox(frmCargando.status, "                    ¡Bienvenido a Hispano AO!", 255, 255, 255, True, False, True)

        'Give the user enough time to read the welcome text
        Call Sleep(500)
    
        Unload frmCargando
    
End Sub

Private Sub LoadTimerIntervals()
        '***************************************************
        'Author: ZaMa
        'Last Modification: 15/03/2011
        'Set the intervals of timers
        '***************************************************
    
        Call MainTimer.SetInterval(TimersIndex.Attack, INT_ATTACK)
        Call MainTimer.SetInterval(TimersIndex.Work, INT_WORK)
        Call MainTimer.SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU)
        Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK)
        Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
        Call MainTimer.SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
        Call MainTimer.SetInterval(TimersIndex.Arrows, INT_ARROWS)
        Call MainTimer.SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK)
        Call MainTimer.SetInterval(TimersIndex.SendDenunce, INT_SENTDENUNCE)
        Call MainTimer.SetInterval(TimersIndex.SendPhoto, INT_SENTPHOTO)
    
        frmMain.macrotrabajo.Interval = INT_MACRO_TRABAJO
        frmMain.macrotrabajo.Enabled = False
    
        'Init timers
        Call MainTimer.Start(TimersIndex.Attack)
        Call MainTimer.Start(TimersIndex.Work)
        Call MainTimer.Start(TimersIndex.UseItemWithU)
        Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
        Call MainTimer.Start(TimersIndex.SendRPU)
        Call MainTimer.Start(TimersIndex.CastSpell)
        Call MainTimer.Start(TimersIndex.Arrows)
        Call MainTimer.Start(TimersIndex.CastAttack)
        Call MainTimer.Start(TimersIndex.SendDenunce)
        Call MainTimer.Start(TimersIndex.SendPhoto)

End Sub

Sub WriteVar(ByVal File As String, _
             ByVal Main As String, _
             ByVal Var As String, _
             ByVal value As String)
             
        '*****************************************************************
        'Writes a var to a text file
        '*****************************************************************
        
        writeprivateprofilestring Main, Var, value, File

End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String) As String

        '*****************************************************************
        'Gets a Var from a text file
        '*****************************************************************
        
        Dim sSpaces As String ' This will hold the input that the program will retrieve
    
        sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
        getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    
        GetVar = RTrim$(sSpaces)
        GetVar = Left$(GetVar, Len(GetVar) - 1)

End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean

        On Error GoTo errHnd

        Dim lPos As Long

        Dim lX   As Long

        Dim iAsc As Integer
    
        '1er test: Busca un simbolo @
        lPos = InStr(sString, "@")

        If (lPos <> 0) Then

                '2do test: Busca un simbolo . después de @ + 1
                If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
                   Exit Function
        
                '3er test: Recorre todos los caracteres y los valída
                For lX = 0 To Len(sString) - 1

                        If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                                iAsc = Asc(mid$(sString, (lX + 1), 1))

                                If Not CMSValidateChar_(iAsc) Then _
                                   Exit Function

                        End If

                Next lX
        
                'Finale
                CheckMailString = True

        End If

errHnd:

End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
        CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
           (iAsc >= 65 And iAsc <= 90) Or _
           (iAsc >= 97 And iAsc <= 122) Or _
           (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)

End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer acá....
Function HayAgua(ByVal x As Integer, ByVal y As Integer) As Boolean
        HayAgua = ((MapData(x, y).Graphic(1).GrhIndex >= 1505 And MapData(x, y).Graphic(1).GrhIndex <= 1520) Or _
           (MapData(x, y).Graphic(1).GrhIndex >= 5665 And MapData(x, y).Graphic(1).GrhIndex <= 5680) Or _
           (MapData(x, y).Graphic(1).GrhIndex >= 13547 And MapData(x, y).Graphic(1).GrhIndex <= 13562)) And _
           MapData(x, y).Graphic(2).GrhIndex = 0
                
End Function

Private Sub LoadClientSetup()

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 11/19/09
        '11/19/09: Pato - Is optional show the frmGuildNews form
        '**************************************************************
        Dim fHandle As Integer
    
        If FileExist(DirInit & "ao.dat", vbArchive) Then
                fHandle = FreeFile
        
                Open DirInit & "ao.dat" For Binary Access Read Lock Write As fHandle
                Get fHandle, , ClientSetup
                Close fHandle
        Else
                'Use dynamic by default
                ClientSetup.bDinamic = True

        End If
    
        'NoRes = ClientSetup.bNoRes
        
        ClientSetup.bGuildNews = Not ClientSetup.bGuildNews
        Set DialogosClanes = New clsGuildDlg
        DialogosClanes.Activo = Not ClientSetup.bGldMsgConsole
        DialogosClanes.CantidadDialogos = ClientSetup.bCantMsgs

End Sub

Private Sub SaveClientSetup()

        '**************************************************************
        'Author: Torres Patricio (Pato)
        'Last Modify Date: 03/11/10
        '
        '**************************************************************
        Dim fHandle As Integer
    
        fHandle = FreeFile
    
        ClientSetup.bNoMusic = Not Audio.MusicActivated
        ClientSetup.bNoSound = Not Audio.SoundActivated
        ClientSetup.bNoSoundEffects = Not Audio.SoundEffectsActivated
        ClientSetup.bGuildNews = Not ClientSetup.bGuildNews
        ClientSetup.bGldMsgConsole = Not DialogosClanes.Activo
        ClientSetup.bCantMsgs = DialogosClanes.CantidadDialogos
    
        Open DirInit & "ao.dat" For Binary As fHandle
        Put fHandle, , ClientSetup
        Close fHandle

End Sub

Private Sub InicializarNombres()

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 11/27/2005
        'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
        '**************************************************************
        
        Ciudades(eCiudad.cUllathorpe) = "Ullathorpe"
        Ciudades(eCiudad.cNix) = "Nix"
        Ciudades(eCiudad.cBanderbill) = "Banderbill"
        Ciudades(eCiudad.cLindos) = "Lindos"
        Ciudades(eCiudad.cArghal) = "Arghâl"
    
        ListaRazas(eRaza.Humano) = "Humano"
        ListaRazas(eRaza.Elfo) = "Elfo"
        ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
        ListaRazas(eRaza.Gnomo) = "Gnomo"
        ListaRazas(eRaza.Enano) = "Enano"

        ListaClases(eClass.Mage) = "Mago"
        ListaClases(eClass.Cleric) = "Clerigo"
        ListaClases(eClass.Warrior) = "Guerrero"
        ListaClases(eClass.Assasin) = "Asesino"
        ListaClases(eClass.Thief) = "Ladron"
        ListaClases(eClass.Bard) = "Bardo"
        ListaClases(eClass.Druid) = "Druida"
        ListaClases(eClass.Bandit) = "Bandido"
        ListaClases(eClass.Paladin) = "Paladin"
        ListaClases(eClass.Hunter) = "Cazador"
        ListaClases(eClass.Worker) = "Trabajador"
        ListaClases(eClass.Pirat) = "Pirata"
    
        SkillsNames(eSkill.Magia) = "Magia"
        SkillsNames(eSkill.Robar) = "Robar"
        SkillsNames(eSkill.Tacticas) = "Evasión en combate"
        SkillsNames(eSkill.Armas) = "Combate cuerpo a cuerpo"
        SkillsNames(eSkill.Meditar) = "Meditar"
        SkillsNames(eSkill.Apuñalar) = "Apuñalar"
        SkillsNames(eSkill.Ocultarse) = "Ocultarse"
        SkillsNames(eSkill.Supervivencia) = "Supervivencia"
        SkillsNames(eSkill.Talar) = "Talar árboles"
        SkillsNames(eSkill.Comerciar) = "Comercio"
        SkillsNames(eSkill.Defensa) = "Defensa con escudos"
        SkillsNames(eSkill.Pesca) = "Pesca"
        SkillsNames(eSkill.Mineria) = "Mineria"
        SkillsNames(eSkill.Carpinteria) = "Carpinteria"
        SkillsNames(eSkill.Herreria) = "Herreria"
        SkillsNames(eSkill.Liderazgo) = "Liderazgo"
        SkillsNames(eSkill.Domar) = "Domar animales"
        SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
        SkillsNames(eSkill.Wrestling) = "Combate sin armas"
        SkillsNames(eSkill.Navegacion) = "Navegacion"

        AtributosNames(eAtributos.Fuerza) = "Fuerza"
        AtributosNames(eAtributos.Agilidad) = "Agilidad"
        AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
        AtributosNames(eAtributos.Carisma) = "Carisma"
        AtributosNames(eAtributos.Constitucion) = "Constitucion"

End Sub

''
' Removes all text from the console and dialogs

Public Sub CleanDialogs()
        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 11/27/2005
        'Removes all text from the console and dialogs
        '**************************************************************
        'Clean console and dialogs
        frmMain.RecTxt.Text = vbNullString
    
        Call DialogosClanes.RemoveDialogs
    
        Call Dialogos.RemoveAllDialogs

End Sub

Public Sub CloseClient()
        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 8/14/2007
        'Frees all used resources, cleans up and leaves
        '**************************************************************
        
        ' Allow new instances of the client to be opened
        Call PrevInstance.ReleaseInstance
    
        EngineRun = False
        frmCargando.Show
        Call AddtoRichTextBox(frmCargando.status, "Liberando recursos...", 0, 0, 0, 0, 0, 0)
    
        Call Resolution.ResetResolution
    
        'Stop tile engine
        Call DeinitTileEngine
    
        Call SaveClientSetup
    
        'Destruimos los objetos públicos creados
        Set CustomMessages = Nothing
        Set CustomKeys = Nothing
        Set SurfaceDB = Nothing
        Set Dialogos = Nothing
        Set DialogosClanes = Nothing
        Set Audio = Nothing
        Set Inventario = Nothing
        Set MainTimer = Nothing
        Set incomingData = Nothing
        Set outgoingData = Nothing
        Set Encriptacion = Nothing
        
        Call UnloadAllForms
    
        End

End Sub

Public Function esGM(ByVal CharIndex As Long) As Boolean
        esGM = False

        If charlist(CharIndex).priv >= 1 And charlist(CharIndex).priv <= 5 Or charlist(CharIndex).priv = 25 Then
                esGM = True
                Exit Function

        End If

End Function

Public Function getTagPosition(ByVal Nick As String) As Integer

        Dim buf As Integer

        buf = InStr(Nick, "<")

        If buf > 0 Then
                getTagPosition = buf
                Exit Function

        End If

        buf = InStr(Nick, "[")

        If buf > 0 Then
                getTagPosition = buf
                Exit Function

        End If

        getTagPosition = Len(Nick) + 2

End Function

Public Sub checkText(ByVal Text As String)

        Dim Nivel As Integer

        If Right(Text, Len(MENSAJE_FRAGSHOOTER_TE_HA_MATADO)) = MENSAJE_FRAGSHOOTER_TE_HA_MATADO Then
                Call ScreenCapture(True)
                Exit Sub

        End If

        If Left(Text, Len(MENSAJE_FRAGSHOOTER_HAS_MATADO)) = MENSAJE_FRAGSHOOTER_HAS_MATADO Then
                EsperandoLevel = True
                Exit Sub

        End If

        If EsperandoLevel Then
                If Right(Text, Len(MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA)) = MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA Then
                        If CInt(mid(Text, Len(MENSAJE_FRAGSHOOTER_HAS_GANADO), (Len(Text) - (Len(MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA) + Len(MENSAJE_FRAGSHOOTER_HAS_GANADO))))) / 2 > ClientSetup.byMurderedLevel Then
                                Call ScreenCapture(True)

                        End If

                End If

        End If

        EsperandoLevel = False

End Sub

Public Function getStrenghtColor(ByVal yFuerza As Byte) As Long

        Dim m As Long

        Dim N As Long

        m = 255 / MAXATRIBUTOS
        N = (m * yFuerza)

        If (N >= 255) Then N = 255 '// Miqueas : Parchesuli
        
        getStrenghtColor = RGB(255 - N, N, 0)

End Function

Public Function getDexterityColor(ByVal yAgilidad As Byte) As Long

        Dim m As Long

        Dim N As Long
        
        m = 255 / MAXATRIBUTOS
        N = (m * yAgilidad)

        If (N >= 255) Then N = 255 '// Miqueas : Parchesuli
         
        getDexterityColor = RGB(255, N, 0)
        
End Function

Public Function getCharIndexByName(ByVal Name As String) As Integer

        Dim i As Long

        For i = 1 To LastChar

                If charlist(i).Nombre = Name Then
                        getCharIndexByName = i
                        Exit Function

                End If

        Next i

End Function

Public Function EsAnuncio(ByVal ForumType As Byte) As Boolean

        '***************************************************
        'Author: ZaMa
        'Last Modification: 22/02/2010
        'Returns true if the post is sticky.
        '***************************************************
        Select Case ForumType

                Case eForumMsgType.ieCAOS_STICKY
                        EsAnuncio = True
            
                Case eForumMsgType.ieGENERAL_STICKY
                        EsAnuncio = True
            
                Case eForumMsgType.ieREAL_STICKY
                        EsAnuncio = True
            
        End Select
    
End Function

Public Function ForumAlignment(ByVal yForumType As Byte) As Byte

        '***************************************************
        'Author: ZaMa
        'Last Modification: 01/03/2010
        'Returns the forum alignment.
        '***************************************************
        Select Case yForumType

                Case eForumMsgType.ieCAOS, eForumMsgType.ieCAOS_STICKY
                        ForumAlignment = eForumType.ieCAOS
            
                Case eForumMsgType.ieGeneral, eForumMsgType.ieGENERAL_STICKY
                        ForumAlignment = eForumType.ieGeneral
            
                Case eForumMsgType.ieREAL, eForumMsgType.ieREAL_STICKY
                        ForumAlignment = eForumType.ieREAL
            
        End Select
    
End Function

Public Sub ResetAllInfo()
        '***************************************************
        'Author: ZaMa
        'Last Modification: 14/06/2011
        '
        '***************************************************
    
        ' Disable timers
        frmMain.Second.Enabled = False
        frmMain.macrotrabajo.Enabled = False
        frmMain.tmrBlink.Enabled = False
    
        Connected = False
    
        'Unload all forms except frmMain, frmConnect and frmCrearPersonaje
        Dim frm As Form

        For Each frm In Forms

                If frm.Name <> frmMain.Name And frm.Name <> frmConnect.Name And _
                   frm.Name <> frmCrearPersonaje.Name Then
            
                        Unload frm

                End If

        Next
    
        On Local Error GoTo 0
    
        ' Return to connection screen
        frmConnect.MousePointer = vbNormal

        If Not frmCrearPersonaje.Visible Then frmConnect.Visible = True
        frmMain.Visible = False
    
        'Stop audio
        Call Audio.StopWave
        frmMain.IsPlaying = PlayLoop.plNone
    
        ' Reset flags
        pausa = False
        UserMeditar = False
        UserEstupido = False
        UserCiego = False
        UserDescansar = False
        UserParalizado = False
        Traveling = False
        UserNavegando = False
        bRain = False
        bFogata = False
        Comerciando = False
    
        MirandoAsignarSkills = False
        MirandoCarpinteria = False
        MirandoEstadisticas = False
        MirandoForo = False
        MirandoHerreria = False
        MirandoParty = False
    
        'Delete all kind of dialogs
        Call CleanDialogs
    
        'Reset some char variables...
        Dim i As Long

        For i = 1 To LastChar
                charlist(i).invisible = False
        Next i

        ' Reset stats
        UserClase = 0
        UserSexo = 0
        UserRaza = 0
        UserHogar = 0
        UserEmail = vbNullString
        SkillPoints = 0
        Alocados = 0
    
        ' Reset skills
        For i = 1 To NUMSKILLS
                UserSkills(i) = 0
        Next i

        ' Reset attributes
        For i = 1 To NUMATRIBUTOS
                UserAtributos(i) = 0
        Next i
    
        ' Clear inventory slots
        Inventario.ClearAllSlots

        ' Connection screen midi
        Call Audio.PlayMIDI("2.mid")

End Sub


