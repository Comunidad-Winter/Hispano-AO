Attribute VB_Name = "GameIni"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Public Type tCabecera 'Cabecera de los con

        Desc As String * 255
        CRC As Long
        MagicWord As Long

End Type

Public Type tSetupMods

        bDinamic    As Boolean
        byMemory    As Byte
        bUseVideo   As Boolean
        bNoMusic    As Boolean
        bNoSound    As Boolean
        bNoRes      As Boolean ' 24/06/2006 - ^[GS]^
        bNoSoundEffects As Boolean
        sGraficos   As String * 13
        bGuildNews  As Boolean ' 11/19/09
        bDie        As Boolean ' 11/23/09 - FragShooter
        bKill       As Boolean ' 11/23/09 - FragShooter
        byMurderedLevel As Byte ' 11/23/09 - FragShooter
        bActive     As Boolean
        bGldMsgConsole As Boolean
        bCantMsgs   As Byte

End Type

Public ClientSetup As tSetupMods

Public MiCabecera  As tCabecera
