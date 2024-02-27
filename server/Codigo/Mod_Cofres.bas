Attribute VB_Name = "Mod_Cofres"
Option Explicit

'Private Type e_Reward

'        Item() As Integer
'        Prob() As Integer

'End Type

Private Type tDrops

        objIndex As Integer
        Amount As Long
        Probability As Byte

End Type

Public Const MAX_ITEM_DROPS As Byte = 5

Public Type e_Reward

        Drop(1 To MAX_ITEM_DROPS) As tDrops

End Type

Private Type Canjeo

        Cantidad As Integer
        objIndex As Integer
        Puntos As Integer

End Type

Public Canjes() As Canjeo

Sub LoadItemDrop(ByVal objIndex As Integer, ByRef Leer As clsIniManager)

        Dim Loopc  As Long

        Dim tmpStr As String

        Dim AscII  As Integer

        AscII = Asc("-")

        With ObjData(objIndex)

                For Loopc = 1 To MAX_ITEM_DROPS
   
                        tmpStr = Leer.GetValue("OBJ" & objIndex, "Drop" & Loopc)
                        
                        .Cofres.Drop(Loopc).objIndex = val(ReadField(1, tmpStr, AscII))
                        .Cofres.Drop(Loopc).Amount = val(ReadField(2, tmpStr, AscII))
                        .Cofres.Drop(Loopc).Probability = val(ReadField(3, tmpStr, AscII))
                        
                Next Loopc

        End With

End Sub

Public Sub ItemDrop_Shop(ByVal uIndex As Integer, _
                         ByVal objIndex As Integer, _
                         ByVal Slot As Byte)
      
        Dim i     As Long

        Dim MiObj As Obj

        If objIndex > 0 Then

                For i = 1 To MAX_ITEM_DROPS
                        
                        If RandomNumber(1, 100) <= ObjData(objIndex).Cofres.Drop(i).Probability Then
                                
                                MiObj.objIndex = ObjData(objIndex).Cofres.Drop(i).objIndex

                                If MiObj.objIndex = iORO Then
                                
                                        UserList(uIndex).Stats.GLD = UserList(uIndex).Stats.GLD + ObjData(objIndex).Cofres.Drop(i).Amount

                                        If UserList(uIndex).Stats.GLD > MAXORO Then UserList(uIndex).Stats.GLD = MAXORO
                                        
                                        Call WriteUpdateGold(uIndex)
                                        
                                Else
                                        MiObj.Amount = ObjData(objIndex).Cofres.Drop(i).Amount
                                        
                                        If Not MeterItemEnInventario(uIndex, MiObj) Then
                                                Call WriteConsoleMsg(uIndex, "No tenes Espacio Suficiente en tu inventario para este item, atencion al piso", FontTypeNames.FONTTYPE_INFO)
                                                Call TirarItemAlPiso(UserList(uIndex).Pos, MiObj)

                                        End If

                                End If

                        End If
                        
                Next i
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(uIndex, Slot, 1)

                Call UpdateUserInv(False, uIndex, Slot)

        End If

        Exit Sub

End Sub

Sub LoadCanjesData()

        On Error GoTo Errhandler

        'Canjes

        If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Canjes."

        Dim Leer As New clsIniManager

        Call Leer.Initialize("ItemsShop.ini")

        Dim NumCanjes As Integer

        Dim i         As Long

        'obtiene el numero de obj
        NumCanjes = val(Leer.GetValue("INIT", "Cantidad"))

        ReDim Canjes(1 To NumCanjes) As Canjeo

        'Llena la lista

        For i = LBound(Canjes()) To UBound(Canjes())

                With Canjes(i)
                        .Cantidad = val(Leer.GetValue("SHOP" & i, "Cantidad"))
                        .objIndex = val(Leer.GetValue("SHOP" & i, "ObjIndex"))
                        .Puntos = val(Leer.GetValue("SHOP" & i, "Puntos"))

                End With

        Next i

        Set Leer = Nothing

        Exit Sub

Errhandler:
        MsgBox "error cargando canjes " & Err.Number & ": " & Err.description

        Exit Sub

End Sub

Public Function GetCanje(ByVal uIndex As Integer, _
                         ByVal Canjea As Byte, _
                         Optional ByRef StrError As String = vbNullString) As Boolean
  
        Dim MiObj As Obj

        GetCanje = False

        With UserList(uIndex)
       
                If Canjea = 0 Then Exit Function
    
                MiObj.Amount = Canjes(Canjea).Cantidad
                MiObj.objIndex = Canjes(Canjea).objIndex

                If .flags.PuntosShop < Canjes(Canjea).Puntos Then
                        StrError = "No tienes los puntos suficientes."
                        Exit Function
                ElseIf Not MeterItemEnInventario(uIndex, MiObj) Then
                        StrError = "No puedes cargar más objetos."
                        Exit Function

                End If

                .flags.PuntosShop = .flags.PuntosShop - Canjes(Canjea).Puntos
                Call WritePuntos(uIndex)

        End With

        GetCanje = True

End Function
