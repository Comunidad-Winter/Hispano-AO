Attribute VB_Name = "modPesca"
Option Explicit
 
Private Type tPescaEvent
    Activado As Byte
    Tiempo As Byte
    CantidadDeZonas As Byte
End Type
 
Private Type tPeces
    Pez As Integer
    min As Byte
    max As Byte
End Type
 
Private Type tZona
    mapa As Integer
    Cantidad As Byte
    Peces() As tPeces
End Type
 
Public PescaEvent As tPescaEvent
Public Zona() As tZona

Public MinPez As Byte
Public MaxPez As Byte
 
Public Sub LoadPesca()
    Dim Leer As clsIniManager
    Set Leer = New clsIniManager
 
    Call Leer.Initialize(App.Path & "\Dat\EventoPesca.dat")
 
    PescaEvent.Activado = Leer.GetValue("INIT", "Activado")
    PescaEvent.Tiempo = Leer.GetValue("INIT", "Tiempo")
    PescaEvent.CantidadDeZonas = Leer.GetValue("INIT", "CantidadDeZonas")
 
    Call LoadPeces
 
    Set Leer = Nothing
End Sub
 
Private Sub LoadPeces()
    Dim Leer As clsIniManager
    Set Leer = New clsIniManager
 
    Call Leer.Initialize(App.Path & "\Dat\EventoPesca.dat")
 
    Dim i As Integer
    Dim j As Integer
    
    Dim tmpStr As String
    Dim Ansci As Byte
    
    Ansci = Asc("-")
 
    With PescaEvent
        ReDim Zona(1 To .CantidadDeZonas) As tZona
 
        For i = 1 To .CantidadDeZonas
            With Zona(i)
                .mapa = Leer.GetValue("ZONA" & i, "Mapa")
                .Cantidad = Leer.GetValue("ZONA" & i, "Cantidad")
             
                ReDim .Peces(1 To .Cantidad) As tPeces
             
                For j = 1 To .Cantidad
                    With .Peces(j)
                        tmpStr = vbNullString
                        tmpStr = Leer.GetValue("ZONA" & i, "Pez" & j)
                        .Pez = General.ReadField(1, tmpStr, Ansci)
                        .min = General.ReadField(2, tmpStr, Ansci)
                        .max = General.ReadField(3, tmpStr, Ansci)
                    End With
                Next j
            End With
        Next i
    End With
 
    Set Leer = Nothing
End Sub
 
Public Function DamePez(ByVal ZonaUser As Byte) As Long
    Dim NoSe As Integer
    
    NoSe = RandomNumber(LBound(Zona(ZonaUser).Peces), UBound(Zona(ZonaUser).Peces))
    
    DamePez = Zona(ZonaUser).Peces(NoSe).Pez
    MinPez = Zona(ZonaUser).Peces(NoSe).min
    MaxPez = Zona(ZonaUser).Peces(NoSe).max
End Function
