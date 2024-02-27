VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   360
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   4590
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4350
      Width           =   2820
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   195
      Left            =   4590
      TabIndex        =   0
      Top             =   3600
      Width           =   2820
   End
   Begin VB.Image imgConectarse 
      Height          =   375
      Left            =   1080
      Top             =   8160
      Width           =   2580
   End
   Begin VB.Image imgCrearPj 
      Height          =   375
      Left            =   8280
      Top             =   8160
      Width           =   2580
   End
   Begin VB.Image imgServArgentina 
      Height          =   795
      Left            =   360
      MousePointer    =   99  'Custom
      Top             =   9240
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   555
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cBotonCrearPj         As clsGraphicalButton

'Private cBotonSalir      As clsGraphicalButton

Private cBotonConectarse      As clsGraphicalButton

'Private cBotonTeclas     As clsGraphicalButton

Public LastButtonPressed      As clsGraphicalButton

Private Const WS_EX_APPWINDOW As Long = &H40000

Private Const GWL_EXSTYLE     As Long = (-20)

Private Const SW_HIDE         As Long = 0

Private Const SW_SHOW         As Long = 5
 
Private Declare Function GetWindowLong _
                Lib "user32" _
                Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Private Declare Function ShowWindow _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal nCmdShow As Long) As Long
 
Private m_bActivated As Boolean
 
Private Sub Form_Activate()

        If Not m_bActivated Then
                m_bActivated = True
                Call SetWindowLong(hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW)
                Call ShowWindow(hWnd, SW_HIDE)
                Call ShowWindow(hWnd, SW_SHOW)

        End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        If KeyCode = 27 Then
                prgRun = False

        End If

End Sub

Private Sub Form_Load()
        
        EngineRun = False

        version.Caption = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
    
        Me.Picture = LoadPicture(DirInterfaces & "VentanaConectar.jpg")
    
        Call LoadButtons

End Sub

Private Sub LoadButtons()
    
        Dim GrhPath As String
    
        GrhPath = DirInterfaces
    
        Set cBotonCrearPj = New clsGraphicalButton
        'Set cBotonSalir = New clsGraphicalButton
        Set cBotonConectarse = New clsGraphicalButton
        'Set cBotonTeclas = New clsGraphicalButton
    
        Set LastButtonPressed = New clsGraphicalButton
        
        Call cBotonCrearPj.Initialize(imgCrearPj, GrhPath & "BotonCrearPersonajeConectar.jpg", _
           GrhPath & "BotonCrearPersonajeConectar.jpg", GrhPath & "BotonCrearPersonajeClickConectar.jpg", Me)
                                    
        'Call cBotonSalir.Initialize(imgSalir, GrhPath & "BotonSalir.jpg", _
         GrhPath & "BotonSalir.jpg", GrhPath & "BotonSalirApretado.jpg", Me)
                                    
        Call cBotonConectarse.Initialize(imgConectarse, GrhPath & "BotonConectarse.jpg", _
           GrhPath & "BotonConectarse.jpg", GrhPath & "BotonConectarseClick.jpg", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        LastButtonPressed.ToggleToNormal

End Sub

Private Sub imgConectarse_Click()

        If frmMain.Socket1.Connected Then
                frmMain.Socket1.Disconnect
                frmMain.Socket1.Cleanup
                DoEvents

        End If
    
        'update user info
        UserName = txtNombre.Text
        UserPassword = txtPasswd.Text

        If CheckUserData() = True Then
                EstadoLogin = Normal
                
                frmMain.Socket1.HostName = CurServerIp
                frmMain.Socket1.RemotePort = CurServerPort
                frmMain.Socket1.Connect

        End If
    
End Sub

Private Sub imgCrearPj_Click()
    
        EstadoLogin = E_MODO.Cp
        
        If frmMain.Socket1.Connected Then
                frmMain.Socket1.Disconnect
                frmMain.Socket1.Cleanup
                DoEvents

        End If

        frmMain.Socket1.HostName = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect
        


End Sub

Private Sub txtPasswd_KeyPress(KeyAscii As Integer)

        If KeyAscii = vbKeyReturn Then imgConectarse_Click

End Sub
