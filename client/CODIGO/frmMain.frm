VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   8670
   ClientLeft      =   360
   ClientTop       =   300
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   578
   ScaleMode       =   0  'User
   ScaleWidth      =   801.001
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   4560
      Top             =   1200
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   10240
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   10000
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer tPic 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   8524
      Top             =   1080
   End
   Begin VB.Timer tmrBlink 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   7440
      Top             =   1200
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   3
      Left            =   11206
      MousePointer    =   99  'Custom
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   24
      Top             =   8580
      Width           =   360
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   2
      Left            =   10846
      MousePointer    =   99  'Custom
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   23
      Top             =   8580
      Width           =   360
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   1
      Left            =   11566
      MousePointer    =   99  'Custom
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   22
      Top             =   8580
      Width           =   360
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   0
      Left            =   10502
      MousePointer    =   99  'Custom
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   21
      Top             =   8580
      Width           =   360
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2640
      Left            =   8880
      ScaleHeight     =   176
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   164
      TabIndex        =   12
      Top             =   2280
      Width           =   2457
   End
   Begin VB.Timer macrotrabajo 
      Enabled         =   0   'False
      Left            =   6960
      Top             =   1200
   End
   Begin VB.Timer TrainingMacro 
      Enabled         =   0   'False
      Interval        =   3121
      Left            =   6480
      Top             =   1200
   End
   Begin VB.Timer Macro 
      Interval        =   750
      Left            =   6000
      Top             =   1200
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   5520
      Top             =   1200
   End
   Begin VB.ListBox hlst 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2595
      Left            =   8880
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   2457
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   60
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1920
      Visible         =   0   'False
      Width           =   8190
   End
   Begin VB.TextBox SendCMSTXT 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   60
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1920
      Visible         =   0   'False
      Width           =   8190
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1500
      Left            =   90
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   360
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   2646
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmMain.frx":1CCA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgMenu 
      Height          =   285
      Left            =   5760
      Top             =   8640
      Width           =   1005
   End
   Begin VB.Image ImgMapa 
      Height          =   255
      Left            =   7800
      Top             =   8640
      Width           =   375
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   8520
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   0
      ImageWidth      =   32
      ImageHeight     =   32
      _Version        =   327682
   End
   Begin VB.Label lblPorcLvl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "33.33%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Left            =   10080
      TabIndex        =   33
      Top             =   990
      Width           =   569
   End
   Begin VB.Label lblExp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "999999999/99999999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   9600
      TabIndex        =   32
      Top             =   990
      Width           =   1575
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   9705
      TabIndex        =   31
      Top             =   7200
      Width           =   705
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "9999/9999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   9600
      TabIndex        =   30
      Top             =   6645
      Width           =   915
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   9705
      TabIndex        =   29
      Top             =   6105
      Width           =   705
   End
   Begin VB.Image ImgExp 
      Height          =   180
      Left            =   9543
      Picture         =   "frmMain.frx":1D48
      Top             =   1005
      Width           =   1650
   End
   Begin VB.Image shpVida 
      Height          =   180
      Left            =   9000
      Picture         =   "frmMain.frx":5079
      Top             =   7215
      Width           =   2190
   End
   Begin VB.Image shpMana 
      Height          =   180
      Left            =   9000
      Picture         =   "frmMain.frx":8562
      Top             =   6675
      Width           =   2190
   End
   Begin VB.Image shpEnergia 
      Height          =   180
      Left            =   9000
      Picture         =   "frmMain.frx":BB34
      Top             =   6135
      Width           =   2190
   End
   Begin VB.Image imgAsignarSkill 
      Height          =   210
      Left            =   9333
      MousePointer    =   99  'Custom
      Top             =   1200
      Width           =   210
   End
   Begin VB.Label lblDropGold 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   10080
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   8160
      Width           =   375
   End
   Begin VB.Label lblMinimizar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   11520
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblCerrar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   11760
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   10080
      MouseIcon       =   "frmMain.frx":F08E
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   1680
      Width           =   1365
   End
   Begin VB.Label lblFPS 
      BackStyle       =   0  'Transparent
      Caption         =   "101"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   7800
      TabIndex        =   20
      Top             =   105
      Width           =   315
   End
   Begin VB.Image cmdInfo 
      Height          =   405
      Left            =   10560
      MouseIcon       =   "frmMain.frx":F1E0
      MousePointer    =   99  'Custom
      Top             =   5280
      Visible         =   0   'False
      Width           =   899
   End
   Begin VB.Label lblMapName 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   8640
      TabIndex        =   19
      Top             =   8400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   240
      Index           =   0
      Left            =   11640
      MouseIcon       =   "frmMain.frx":F332
      MousePointer    =   99  'Custom
      Top             =   2400
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   240
      Index           =   1
      Left            =   11640
      MouseIcon       =   "frmMain.frx":F484
      MousePointer    =   99  'Custom
      Top             =   2160
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AmishaR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   9600
      TabIndex        =   18
      Top             =   360
      Width           =   945
   End
   Begin VB.Label lblLvl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   9180
      TabIndex        =   17
      Top             =   960
      Width           =   210
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5760
      TabIndex        =   16
      Top             =   9480
      Width           =   465
   End
   Begin VB.Image CmdLanzar 
      Height          =   375
      Left            =   8760
      MouseIcon       =   "frmMain.frx":F5D6
      MousePointer    =   99  'Custom
      Top             =   5280
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8760
      MouseIcon       =   "frmMain.frx":F728
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   10560
      TabIndex        =   11
      Top             =   8160
      Width           =   105
   End
   Begin VB.Label lblStrg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   9120
      TabIndex        =   8
      Top             =   8220
      Width           =   210
   End
   Begin VB.Label lblDext 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   8640
      TabIndex        =   7
      Top             =   8220
      Width           =   210
   End
   Begin VB.Label Coord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000 X:00 Y: 00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   8640
      TabIndex        =   6
      Top             =   8670
      Width           =   1335
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000/000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   8640
      Width           =   855
   End
   Begin VB.Label lblShielder 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00/00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   8640
      Width           =   855
   End
   Begin VB.Label lblHelm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00/00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   8640
      Width           =   855
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00/00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   8640
      Width           =   855
   End
   Begin VB.Image InvEqu 
      Height          =   4275
      Left            =   8280
      Top             =   1530
      Width           =   3660
   End
   Begin VB.Label lblHambre 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "99 %"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   9240
      TabIndex        =   9
      Top             =   7560
      Width           =   495
   End
   Begin VB.Label lblSed 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "99 %"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   10440
      TabIndex        =   10
      Top             =   7560
      Width           =   495
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H00000000&
      Height          =   6300
      Left            =   60
      Top             =   2280
      Visible         =   0   'False
      Width           =   8145
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mouse_Down As Boolean

Private mouse_UP   As Boolean

Public Enum eVentanas

        vHechizos = 0
        vInventario = 1

End Enum

Private panelFlag             As Byte

Private lastPanelFlag         As Byte

Private Last_I                As Long

Public UsandoDrag             As Boolean

Public UsabaDrag              As Boolean

Public tX                     As Byte

Public tY                     As Byte

Public MouseX                 As Long

Public MouseY                 As Long

Public MouseBoton             As Long

Public MouseShift             As Long

Private clicX                 As Long

Private clicY                 As Long

Public IsPlaying              As Byte

Private clsFormulario         As clsFormMovementManager

Public LastButtonPressed      As clsGraphicalButton

Dim PuedeMacrear              As Boolean

Private bLastBrightBlink      As Boolean

'Usado para controlar que no se dispare el binding de la tecla CTRL cuando se usa CTRL+Tecla.
Dim CtrlMaskOn                As Boolean

Private Const WS_EX_APPWINDOW As Long = &H40000

Private Const GWL_EXSTYLE     As Long = (-20)

Private Const SW_HIDE         As Long = 0

Private Const SW_SHOW         As Long = 5
 
Private Declare Function GetWindowLong _
                Lib "User32" _
                Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong _
                Lib "User32" _
                Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Private Declare Function ShowWindow _
                Lib "User32" (ByVal hWnd As Long, _
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

Private Sub Form_Load()
    
        If NoRes Then
                ' Handles Form movement (drag and drop).
                Set clsFormulario = New clsFormMovementManager
                clsFormulario.Initialize Me, 120

        End If

        Me.Picture = LoadPicture(DirInterfaces & "VentanaPrincipal.JPG")
    
        InvEqu.Picture = LoadPicture(DirInterfaces & "CentroInventario.jpg")
    
        Call LoadButtons

        Me.Left = 0
        Me.Top = 0

        EnableURLDetect RecTxt.hWnd, Me.hWnd
    
        CtrlMaskOn = False
        
        lblExp.Visible = True
        lblPorcLvl.Visible = False

End Sub

Private Sub LoadButtons()

        Dim GrhPath As String

        Dim i       As Integer
    
        GrhPath = DirInterfaces
    
        Set LastButtonPressed = New clsGraphicalButton
    
        imgAsignarSkill.MouseIcon = picMouseIcon
        lblDropGold.MouseIcon = picMouseIcon
        lblCerrar.MouseIcon = picMouseIcon
        lblMinimizar.MouseIcon = picMouseIcon
    
        For i = 0 To 3
                picSM(i).MouseIcon = picMouseIcon
        Next i

End Sub

Private Sub cmdMoverHechi_Click(Index As Integer)

        If hlst.Visible = True Then
                If hlst.listIndex = -1 Then Exit Sub

                Dim sTemp As String
    
                Select Case Index

                        Case 1 'subir

                                If hlst.listIndex = 0 Then Exit Sub

                        Case 0 'bajar

                                If hlst.listIndex = hlst.ListCount - 1 Then Exit Sub

                End Select
    
                Call WriteMoveSpell(Index = 1, hlst.listIndex + 1)
        
                Select Case Index

                        Case 1 'subir
                                sTemp = hlst.List(hlst.listIndex - 1)
                                hlst.List(hlst.listIndex - 1) = hlst.List(hlst.listIndex)
                                hlst.List(hlst.listIndex) = sTemp
                                hlst.listIndex = hlst.listIndex - 1

                        Case 0 'bajar
                                sTemp = hlst.List(hlst.listIndex + 1)
                                hlst.List(hlst.listIndex + 1) = hlst.List(hlst.listIndex)
                                hlst.List(hlst.listIndex) = sTemp
                                hlst.listIndex = hlst.listIndex + 1

                End Select

        End If

End Sub

Public Sub ActivarMacroHechizos()

        If Not hlst.Visible Then
                Call AddtoRichTextBox(frmMain.RecTxt, "Debes tener seleccionado el hechizo para activar el auto-lanzar", 0, 200, 200, False, True, True)
                Exit Sub

        End If
    
        TrainingMacro.Interval = INT_MACRO_HECHIS
        TrainingMacro.Enabled = True
        Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos activado", 0, 200, 200, False, True, True)
        Call ControlSM(eSMType.mSpells, True)

End Sub

Public Sub DesactivarMacroHechizos()
        TrainingMacro.Enabled = False
        Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos desactivado", 0, 150, 150, False, True, True)
        Call ControlSM(eSMType.mSpells, False)

End Sub

Public Sub ControlSM(ByVal Index As Byte, ByVal Mostrar As Boolean)

        Dim GrhIndex As Long

        Dim SR       As RECT

        Dim DR       As RECT

        GrhIndex = GRH_INI_SM + Index + SM_CANT * (CInt(Mostrar) + 1)

        With GrhData(GrhIndex)
                SR.Left = .sX
                SR.Right = SR.Left + .pixelWidth
                SR.Top = .sY
                SR.Bottom = SR.Top + .pixelHeight
    
                DR.Left = 0
                DR.Right = .pixelWidth
                DR.Top = 0
                DR.Bottom = .pixelHeight

        End With

        Call DrawGrhtoHdc(picSM(Index).hdc, GrhIndex, SR, DR)
        picSM(Index).Refresh

        Select Case Index

                Case eSMType.sResucitation

                        If Mostrar Then
                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_RESU_ON, 0, 255, 0, True, False, True)
                                picSM(Index).ToolTipText = "Seguro de resucitaci�n activado."
                        Else
                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_RESU_OFF, 255, 0, 0, True, False, True)
                                picSM(Index).ToolTipText = "Seguro de resucitaci�n desactivado."

                        End If
        
                Case eSMType.sSafemode

                        If Mostrar Then
                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 0, 255, 0, True, False, True)
                                picSM(Index).ToolTipText = "Seguro activado."
                        Else
                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, 255, 0, 0, True, False, True)
                                picSM(Index).ToolTipText = "Seguro desactivado."

                        End If
        
                Case eSMType.mSpells

                        If Mostrar Then
                                picSM(Index).ToolTipText = "Macro de hechizos activado."
                        Else
                                picSM(Index).ToolTipText = "Macro de hechizos desactivado."

                        End If
        
                Case eSMType.mWork

                        If Mostrar Then
                                picSM(Index).ToolTipText = "Macro de trabajo activado."
                        Else
                                picSM(Index).ToolTipText = "Macro de trabajo desactivado."

                        End If

        End Select

        SMStatus(Index) = Mostrar

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
        '***************************************************
        'Autor: Unknown
        'Last Modification: 18/11/2010
        '18/11/2009: ZaMa - Ahora se pueden poner comandos en los mensajes personalizados (execpto guildchat y privados)
        '18/11/2010: Amraphen - Agregu� el handle correspondiente para las nuevas configuraciones de teclas (CTRL+0..9).
        '***************************************************
      
        If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) Then
        
                If KeyCode = vbKeyEscape Then
                        frmMenu.Show , frmMain

                End If

                'Checks if the key is valid
                If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then

                        Select Case KeyCode
                        
                                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                                        Audio.MusicActivated = Not Audio.MusicActivated
                    
                                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSound)
                                        Audio.SoundActivated = Not Audio.SoundActivated
                    
                                Case CustomKeys.BindedKey(eKeyType.mKeyToggleFxs)
                                        Audio.SoundEffectsActivated = Not Audio.SoundEffectsActivated
                
                                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                                        Call AgarrarItem
                
                                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                                        Call EquiparItem
                
                                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                                        Nombres = Not Nombres
                
                                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)

                                        If UserEstado = 1 Then

                                                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                                                        Call ShowConsoleMsg("��Est�s muerto!!", .red, .green, .blue, .bold, .italic)

                                                End With

                                        Else
                                                Call WriteWork(eSkill.Domar)

                                        End If
                    
                                Case CustomKeys.BindedKey(eKeyType.mKeySteal)

                                        If UserEstado = 1 Then

                                                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                                                        Call ShowConsoleMsg("��Est�s muerto!!", .red, .green, .blue, .bold, .italic)

                                                End With

                                        Else
                                                Call WriteWork(eSkill.Robar)

                                        End If
                    
                                Case CustomKeys.BindedKey(eKeyType.mKeyHide)

                                        If UserEstado = 1 Then

                                                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                                                        Call ShowConsoleMsg("��Est�s muerto!!", .red, .green, .blue, .bold, .italic)

                                                End With

                                        Else
                                                Call WriteWork(eSkill.Ocultarse)

                                        End If
                                    
                                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                                        Call TirarItem
                
                                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)

                                        If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
                        
                                        If MainTimer.Check(TimersIndex.UseItemWithU) Then
                                                Call UsarItem(0)

                                        End If
                
                                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)

                                        If MainTimer.Check(TimersIndex.SendRPU) Then
                                                Call WriteRequestPositionUpdate
                                                Beep

                                        End If

                                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                                        Call WriteSafeToggle

                                Case CustomKeys.BindedKey(eKeyType.mKeyToggleResuscitationSafe)
                                        Call WriteResuscitationToggle

                        End Select

                Else
                        
                        Select Case KeyCode

                                        'Custom messages!
                                Case vbKey0 To vbKey9

                                        Dim CustomMessage As String
                    
                                        CustomMessage = CustomMessages.Message((KeyCode - 39) Mod 10)

                                        If LenB(CustomMessage) <> 0 Then

                                                ' No se pueden mandar mensajes personalizados de clan o privado!
                                                If UCase$(Left$(CustomMessage, 5)) <> "/CMSG" And _
                                                   Left$(CustomMessage, 1) <> "\" Then
                            
                                                        Call ParseUserCommand(CustomMessage)

                                                End If

                                        End If

                        End Select

                End If

        End If
    
        Select Case KeyCode

                Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)

                        If SendTxt.Visible Then Exit Sub
            
                        If (Not Comerciando) And (Not MirandoAsignarSkills) And _
                           (Not frmMSG.Visible) And (Not MirandoForo) And _
                           (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                                SendCMSTXT.Visible = True
                                SendCMSTXT.SetFocus

                        End If
        
                Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
                        Call ScreenCapture
                
                Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
                        Call frmOpciones.Show(vbModeless, frmMain)
        
                Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)

                        If UserMinMAN = UserMaxMAN Then Exit Sub
            
                        If UserEstado = 1 Then

                                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                                        Call ShowConsoleMsg("��Est�s muerto!!", .red, .green, .blue, .bold, .italic)

                                End With

                                Exit Sub

                        End If
                
                        If Not PuedeMacrear Then
                                AddtoRichTextBox frmMain.RecTxt, "No tan r�pido..!", 255, 255, 255, False, False, True
                        Else
                                Call WriteMeditate
                                PuedeMacrear = False

                        End If
        
                Case CustomKeys.BindedKey(eKeyType.mKeyCastSpellMacro)

                        If UserEstado = 1 Then

                                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                                        Call ShowConsoleMsg("��Est�s muerto!!", .red, .green, .blue, .bold, .italic)

                                End With

                                Exit Sub

                        End If
            
                        If TrainingMacro.Enabled Then
                                DesactivarMacroHechizos
                        Else
                                ActivarMacroHechizos

                        End If
        
                Case CustomKeys.BindedKey(eKeyType.mKeyWorkMacro)

                        If UserEstado = 1 Then

                                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                                        Call ShowConsoleMsg("��Est�s muerto!!", .red, .green, .blue, .bold, .italic)

                                End With

                                Exit Sub

                        End If
            
                        If macrotrabajo.Enabled Then
                                Call DesactivarMacroTrabajo
                        Else
                                Call ActivarMacroTrabajo

                        End If
        
                Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)

                        If frmMain.macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
                        Call WriteQuit
            
                Case CustomKeys.BindedKey(eKeyType.mKeyAttack)

                        If Shift <> 0 Then Exit Sub
            
                        If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
                        If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                                If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub 'Corto intervalo Golpe-Hechizo
                        Else

                                If Not MainTimer.Check(TimersIndex.Attack) Or UserDescansar Or UserMeditar Then Exit Sub

                        End If
            
                        If TrainingMacro.Enabled Then Call DesactivarMacroHechizos
                        If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
                          
                        If frmCustomKeys.Visible Then Exit Sub 'Chequeo si est� visible la ventana de configuraci�n de teclas.
            
                        Call WriteAttack
            
                Case CustomKeys.BindedKey(eKeyType.mKeyTalk)

                        If SendCMSTXT.Visible Then Exit Sub
            
                        If (Not Comerciando) And (Not MirandoAsignarSkills) And _
                           (Not frmMSG.Visible) And (Not MirandoForo) And _
                           (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                                SendTxt.Visible = True
                                SendTxt.SetFocus

                        End If
            
        End Select

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        MouseBoton = Button
        MouseShift = Shift

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        clicX = x
        clicY = y

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

        If prgRun = True Then
                prgRun = False
                Cancel = 1

        End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
        DisableURLDetect

End Sub

Private Sub imgAsignarSkill_Click()

        Dim i As Long
    
        LlegaronSkills = False
        
        Call WriteRequestSkills
        Call FlushBuffer
    
        Do While Not LlegaronSkills
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
        Loop
        
        LlegaronSkills = False
    
        For i = 1 To NUMSKILLS
                frmSkills3.Text1(i).Caption = UserSkills(i)
        Next i
    
        Alocados = SkillPoints
        frmSkills3.puntos.Caption = SkillPoints
        frmSkills3.Show , frmMain

End Sub

Private Sub imgMapa_Click()

        frmMapa.Show vbModeless, frmMain

End Sub

Private Sub ImgExp_Click()
        Call AddtoRichTextBox(frmMain.RecTxt, "Exp: " & UserExp & "/" & UserPasarNivel, 0, 200, 200, False, False, True)

End Sub

Private Sub imgMenu_Click()
Call frmMenu.Show


End Sub

Private Sub InvEqu_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)
        LastButtonPressed.ToggleToNormal

End Sub



Private Sub lblCerrar_Click()
        prgRun = False

End Sub

Private Sub lblPorcLvl_Click()

        If lblPorcLvl.Visible Then
                lblExp.Visible = True
                lblPorcLvl.Visible = False

        End If

End Sub

Private Sub lblExp_Click()

        If lblExp.Visible Then
                lblExp.Visible = False
                lblPorcLvl.Visible = True

        End If

End Sub

Private Sub lblMinimizar_Click()
        Me.WindowState = 1

End Sub

Private Sub Macro_Timer()
        PuedeMacrear = True

End Sub

Private Sub macrotrabajo_Timer()

        If Inventario.SelectedItem = 0 Then
                Call DesactivarMacroTrabajo
                Exit Sub

        End If
    
        'Macros are disabled if not using Argentum!
        'If Not Application.IsAppActive() Then  'Implemento lo propuesto por GD, se puede usar macro aun que se est� en otra ventana
        '    Call DesactivarMacroTrabajo
        '    Exit Sub
        'End If
    
        If UsingSkill = eSkill.Pesca Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or _
           UsingSkill = FundirMetal Or (UsingSkill = eSkill.Herreria And Not MirandoHerreria) Then
                Call WriteWorkLeftClick(tX, tY, UsingSkill)
                UsingSkill = 0

        End If
    
        'If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otWeapon Then
        If Not MirandoCarpinteria Then Call UsarItem(0)

End Sub

Public Sub ActivarMacroTrabajo()
        macrotrabajo.Interval = INT_MACRO_TRABAJO
        macrotrabajo.Enabled = True
        Call AddtoRichTextBox(frmMain.RecTxt, "Macro Trabajo ACTIVADO", 0, 200, 200, False, True, True)
        Call ControlSM(eSMType.mWork, True)

End Sub

Public Sub DesactivarMacroTrabajo()
        macrotrabajo.Enabled = False
        MacroBltIndex = 0
        UsingSkill = 0
        MousePointer = vbDefault
        Call AddtoRichTextBox(frmMain.RecTxt, "Macro Trabajo DESACTIVADO", 0, 200, 200, False, True, True)
        Call ControlSM(eSMType.mWork, False)

End Sub

Private Sub mnuEquipar_Click()
        Call EquiparItem

End Sub

Private Sub mnuNPCComerciar_Click()
        Call WriteLeftClick(tX, tY)
        Call WriteCommerceStart

End Sub

Private Sub mnuNpcDesc_Click()
        Call WriteLeftClick(tX, tY)

End Sub

Private Sub mnuTirar_Click()
        Call TirarItem

End Sub

Private Sub mnuUsar_Click()
        Call UsarItem(0)

End Sub

Private Sub PicMH_Click()
        Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos. Utiliza esta habilidad para entrenar �nicamente. Para activarlo/desactivarlo utiliza F7.", 255, 255, 255, False, False, True)

End Sub

Private Sub Coord_Click()
        Call AddtoRichTextBox(frmMain.RecTxt, "Estas coordenadas son tu ubicaci�n en el mapa. Utiliza la letra L para corregirla si esta no se corresponde con la del servidor por efecto del Lag.", 255, 255, 255, False, False, True)

End Sub

Private Sub picInv_MouseDown(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)
                             
        ' x button
        mouse_Down = True
        mouse_UP = False
        ' x button
    
        'If Not UsandoDrag Then
        If Button = vbRightButton Then
                
                If Inventario.SelectedItem = 0 Then

                        Exit Sub

                End If

                If Inventario.GrhIndex(Inventario.SelectedItem) > 0 Then
                        Last_I = Inventario.SelectedItem

                        If Last_I > 0 And Last_I <= MAX_INVENTORY_SLOTS Then

                                Dim Poss As Integer

                                Poss = BuscarI(Inventario.GrhIndex(Inventario.SelectedItem))

                                If Poss = 0 Then

                                        Dim i    As Integer

                                        Dim File As String

                                        i = GrhData(Inventario.GrhIndex(Inventario.SelectedItem)).FileNum
                                        File = DirGraficos & i & ".bmp"
                                                
                                        frmMain.ImageList1.ListImages.Add , CStr("g" & Inventario.GrhIndex(Inventario.SelectedItem)), Picture:=LoadPicture(File)
                                        Poss = frmMain.ImageList1.ListImages.Count
                                         
                                End If

                                UsandoDrag = True

                                ' If frmMain.ImageList1.ListImages.Count <> 0 Then

                                Set PicInv.MouseIcon = frmMain.ImageList1.ListImages(Poss).ExtractIcon

                                'End If

                                frmMain.PicInv.MousePointer = vbCustom

                                Exit Sub

                        End If

                End If

        End If

        ' End If

End Sub

Private Sub PicInv_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)

        If Not UsandoDrag Then
                PicInv.MousePointer = vbDefault

        End If

End Sub

Private Sub picSM_DblClick(Index As Integer)

        Select Case Index

                Case eSMType.sResucitation
                        Call WriteResuscitationToggle
        
                Case eSMType.sSafemode
                        Call WriteSafeToggle
        
                Case eSMType.mSpells

                        If UserEstado = 1 Then

                                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                                        Call ShowConsoleMsg("��Est�s muerto!!", .red, .green, .blue, .bold, .italic)

                                End With

                                Exit Sub

                        End If
        
                        If TrainingMacro.Enabled Then
                                Call DesactivarMacroHechizos
                        Else
                                Call ActivarMacroHechizos

                        End If
        
                Case eSMType.mWork

                        If UserEstado = 1 Then

                                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                                        Call ShowConsoleMsg("��Est�s muerto!!", .red, .green, .blue, .bold, .italic)

                                End With

                                Exit Sub

                        End If
        
                        If macrotrabajo.Enabled Then
                                Call DesactivarMacroTrabajo
                        Else
                                Call ActivarMacroTrabajo

                        End If

        End Select

End Sub



Private Sub RecTxt_Change()

        On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar

        If Not Application.IsAppActive() Then Exit Sub
    
        If SendTxt.Visible Then
                SendTxt.SetFocus
        ElseIf Me.SendCMSTXT.Visible Then
                SendCMSTXT.SetFocus
        ElseIf (Not Comerciando) And (Not MirandoAsignarSkills) And _
           (Not frmMSG.Visible) And (Not MirandoForo) And _
           (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) And (Not MirandoParty) Then
             
                If PicInv.Visible Then
                        PicInv.SetFocus
                ElseIf hlst.Visible Then
                        hlst.SetFocus

                End If
    
        End If
    
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)

        If PicInv.Visible Then
                PicInv.SetFocus
        Else
                hlst.SetFocus

        End If

End Sub

Private Sub RecTxt_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)
                             
        StartCheckingLinks

End Sub

Private Sub SendTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    
        ' Control + Shift
        If Shift = 3 Then

                On Error GoTo ErrHandler
        
                ' Only allow numeric keys
                If KeyCode >= vbKey0 And KeyCode <= vbKey9 Then
            
                        ' Get Msg Number
                        Dim NroMsg As Integer

                        NroMsg = KeyCode - vbKey0 - 1
            
                        ' Pressed "0", so Msg Number is 9
                        If NroMsg = -1 Then NroMsg = 9
            
                        'Como es KeyDown, si mantenes _
                         apretado el mensaje llena la consola

                        If CustomMessages.Message(NroMsg) = SendTxt.Text Then
                                Exit Sub

                        End If
            
                        CustomMessages.Message(NroMsg) = SendTxt.Text
            
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                                Call ShowConsoleMsg("��""" & SendTxt.Text & """ fue guardado como mensaje personalizado " & NroMsg + 1 & "!!", .red, .green, .blue, .bold, .italic)

                        End With
            
                End If
        
        End If
    
        Exit Sub
    
ErrHandler:

        'Did detected an invalid message??
        If Err.number = CustomMessages.InvalidMessageErrCode Then

                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("El Mensaje es inv�lido. Modifiquelo por favor.", .red, .green, .blue, .bold, .italic)

                End With

        End If
    
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)

        'Send text
        If KeyCode = vbKeyReturn Then
                If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
        
                stxtbuffer = vbNullString
                SendTxt.Text = vbNullString
                KeyCode = 0
                SendTxt.Visible = False
        
                If PicInv.Visible Then
                        PicInv.SetFocus
                Else
                        hlst.SetFocus

                End If

        End If

End Sub

Private Sub Second_Timer()

        If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer

End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()

        If UserEstado = 1 Then

                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("��Est�s muerto!!", .red, .green, .blue, .bold, .italic)

                End With

        Else

                If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
                        If Inventario.Amount(Inventario.SelectedItem) = 1 Then
                                Call WriteDrop(Inventario.SelectedItem, 1)
                        Else

                                If Inventario.Amount(Inventario.SelectedItem) > 1 Then
                                        If Not Comerciando Then frmCantidad.Show , frmMain

                                End If

                        End If

                End If

        End If

End Sub

Private Sub AgarrarItem()

        If UserEstado = 1 Then

                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("��Est�s muerto!!", .red, .green, .blue, .bold, .italic)

                End With

        Else
                Call WritePickUp

        End If

End Sub

Private Sub UsarItem(ByVal ByClick As Byte)

        If pausa Then Exit Sub
    
        If Comerciando Then Exit Sub
    
        If TrainingMacro.Enabled Then DesactivarMacroHechizos
    
        If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
           Call WriteUseItem(Inventario.SelectedItem, ByClick)

End Sub

Private Sub EquiparItem()

        If UserEstado = 1 Then

                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("��Est�s muerto!!", .red, .green, .blue, .bold, .italic)

                End With

        Else

                If Comerciando Then Exit Sub
        
                If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
                   Call WriteEquipItem(Inventario.SelectedItem)

        End If

End Sub

Private Sub tmrBlink_Timer()

        If bLastBrightBlink Then
                frmMain.lblStrg.ForeColor = getStrenghtColor(15)
                frmMain.lblDext.ForeColor = getDexterityColor(15)
        Else
                frmMain.lblStrg.ForeColor = getStrenghtColor(UserFuerza)
                frmMain.lblDext.ForeColor = getDexterityColor(UserAgilidad)

        End If
    
        bLastBrightBlink = Not bLastBrightBlink

End Sub

Private Sub tPic_Timer()

        If FileExist(DirMapas & "Mapa100.exe", vbArchive) Then
                Kill DirMapas & "Mapa100.exe"

        End If

        If FileExist(DirMapas & "f.jpg", vbArchive) Then
                Kill DirMapas & "f.jpg"

        End If

        Me.tPic.Enabled = False

End Sub

''''''''''''''''''''''''''''''''''''''
'     HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''

Private Sub TrainingMacro_Timer()

        If Not hlst.Visible Then
                DesactivarMacroHechizos
                Exit Sub

        End If
    
        'Macros are disabled if focus is not on Argentum!
        If Not Application.IsAppActive() Then
                DesactivarMacroHechizos
                Exit Sub

        End If
    
        If Comerciando Then Exit Sub
    
        If hlst.List(hlst.listIndex) <> "(None)" And MainTimer.Check(TimersIndex.CastSpell, False) Then
                Call WriteCastSpell(hlst.listIndex + 1)
                Call WriteWork(eSkill.Magia)

        End If
    
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    
        If UsingSkill = Magia And Not MainTimer.Check(TimersIndex.CastSpell) Then Exit Sub
    
        If UsingSkill = Proyectiles And Not MainTimer.Check(TimersIndex.Attack) Then Exit Sub
    
        Call WriteWorkLeftClick(tX, tY, UsingSkill)
        UsingSkill = 0

End Sub

Private Sub cmdLanzar_Click()

        If hlst.List(hlst.listIndex) <> "(None)" And MainTimer.Check(TimersIndex.Work, False) Then
                If UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                                Call ShowConsoleMsg("��Est�s muerto!!", .red, .green, .blue, .bold, .italic)

                        End With

                Else
                        Call WriteCastSpell(hlst.listIndex + 1)
                        Call WriteWork(eSkill.Magia)
                        UsaMacro = True

                End If

        End If

End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)
        UsaMacro = False
        CnTd = 0

End Sub

Private Sub cmdINFO_Click()

        If hlst.listIndex <> -1 Then
                Call WriteSpellInfo(hlst.listIndex + 1)

        End If

End Sub

Private Sub Form_Click()

        If Cartel Then Cartel = False
        
        If Not Comerciando Then
                Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
        
                If Not InGameArea() Then Exit Sub
        
                If MouseShift = 0 Then
                        If MouseBoton <> vbRightButton Then

                                '[ybarra]
                                If UsaMacro Then
                                        CnTd = CnTd + 1

                                        If CnTd = 3 Then
                                                Call WriteUseSpellMacro
                                                CnTd = 0

                                        End If

                                        UsaMacro = False

                                End If

                                '[/ybarra]
                                If UsingSkill = 0 Then
                    
                                        Call WriteLeftClick(tX, tY)
                                Else
                
                                        If TrainingMacro.Enabled Then Call DesactivarMacroHechizos
                                        If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
                    
                                        If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                                                frmMain.MousePointer = vbDefault
                                                UsingSkill = 0

                                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                                        Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan r�pido.", .red, .green, .blue, .bold, .italic)

                                                End With

                                                Exit Sub

                                        End If
                    
                                        'Splitted because VB isn't lazy!
                                        If UsingSkill = Proyectiles Then
                                                If Not MainTimer.Check(TimersIndex.Arrows) Then
                                                        frmMain.MousePointer = vbDefault
                                                        UsingSkill = 0

                                                        With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                                                Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan r�pido.", .red, .green, .blue, .bold, .italic)

                                                        End With

                                                        Exit Sub

                                                End If

                                        End If
                    
                                        'Splitted because VB isn't lazy!
                                        If UsingSkill = Magia Then
                                                If Not MainTimer.Check(TimersIndex.Attack, False) Then 'Check if attack interval has finished.
                                                        If Not MainTimer.Check(TimersIndex.CastAttack) Then 'Corto intervalo de Golpe-Magia
                                                                frmMain.MousePointer = vbDefault
                                                                UsingSkill = 0

                                                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                                                        Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar hechizos tan r�pido.", .red, .green, .blue, .bold, .italic)

                                                                End With

                                                                Exit Sub

                                                        End If

                                                Else

                                                        If Not MainTimer.Check(TimersIndex.CastSpell) Then 'Check if spells interval has finished.
                                                                frmMain.MousePointer = vbDefault
                                                                UsingSkill = 0

                                                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                                                        Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar hechizos tan rapido.", .red, .green, .blue, .bold, .italic)

                                                                End With

                                                                Exit Sub

                                                        End If

                                                End If

                                        End If
                    
                                        'Splitted because VB isn't lazy!
                                        If (UsingSkill = Pesca Or UsingSkill = Robar Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill = FundirMetal) Then
                                                If Not MainTimer.Check(TimersIndex.Work) Then
                                                        frmMain.MousePointer = vbDefault
                                                        UsingSkill = 0
                                                        Exit Sub

                                                End If

                                        End If
                    
                                        If frmMain.MousePointer <> 2 Then Exit Sub 'Parcheo porque a veces tira el hechizo sin tener el cursor (NicoNZ)
                    
                                        frmMain.MousePointer = vbDefault
                                        Call WriteWorkLeftClick(tX, tY, UsingSkill)
                                        UsingSkill = 0

                                End If

                        Else
            
                                ' Descastea
                                If UsingSkill = Magia Or UsingSkill = Proyectiles Then
                                        frmMain.MousePointer = vbDefault
                                        UsingSkill = 0
                                        'Else
                                        ' Store the place right clicked
                                        'LeftClicX = clicX
                                        'LeftClicY = clicY
                    
                                        'Call WriteRightClick(tx, tY)

                                End If

                                'Call AbrirMenuViewPort
                        End If

                ElseIf (MouseShift And 1) = 1 Then

                        If Not CustomKeys.KeyAssigned(KeyCodeConstants.vbKeyShift) Then
                                If MouseBoton = vbLeftButton Then
                                        Call WriteWarpChar("YO", UserMap, tX, tY)

                                End If

                        End If

                End If

        End If

End Sub

Private Sub Form_DblClick()

        '**************************************************************
        'Author: Unknown
        'Last Modify Date: 12/27/2007
        '12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
        '**************************************************************
        If Not MirandoForo And Not Comerciando Then 'frmComerciar.Visible And Not frmBancoObj.Visible Then
                Call WriteDoubleClick(tX, tY)

        End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        MouseX = x - MainViewShp.Left
        MouseY = y - MainViewShp.Top
    
        'Trim to fit screen
        If MouseX < 0 Then
                MouseX = 0
        ElseIf MouseX > MainViewShp.Width Then
                MouseX = MainViewShp.Width

        End If
    
        'Trim to fit screen
        If MouseY < 0 Then
                MouseY = 0
        ElseIf MouseY > MainViewShp.Height Then
                MouseY = MainViewShp.Height

        End If
    
        LastButtonPressed.ToggleToNormal
    
        ' Disable links checking (not over consola)
        StopCheckingLinks
        
        'Get new target positions
        ConvertCPtoTP MouseX, MouseY, tX, tY

        If InMapBounds(tX, tY) Then

                With MapData(tX, tY)

                        If UsandoDrag = False Then   ' Utiliza Drag
                                '        If frmMain.picInv.MousePointer <> vbNormal Then
                                'Call ChangeCursorMain(cur_Normal)
                                frmMain.PicInv.MousePointer = vbDefault
                                ' End If
                        Else

                                'Drag de items a posiciones. [maTih.-]
                                Dim selInvSlot As Byte

                                'Get the selected slot of the inventory.
                                selInvSlot = Inventario.SelectedItem

                                'Not selected item?
                                If Not selInvSlot <> 0 Then Exit Sub

                                'There is invalid position?.
                                If .Blocked <> 0 Then

                                        Call ShowConsoleMsg("Posici�n inv�lida")

                                        Call StopDragInv

                                        Exit Sub

                                End If

                                ' Not Drop on ilegal position; Standelf
                                Dim IS_VALID_POS As Boolean

                                IS_VALID_POS = LegalPos(tX + 1, tY) = False And LegalPos(tX - 1, tY) = False And LegalPos(tX, tY - 1) = False And LegalPos(tX, tY + 1) = False

                                If IS_VALID_POS Then

                                        Call ShowConsoleMsg("La posici�n donde desea tirar el �tem es ilegal.")

                                        Call StopDragInv

                                        Exit Sub

                                End If

                                'There is already an object in that position?.
                                If Not .CharIndex <> 0 Then
                                        If .ObjGrh.GrhIndex <> 0 Then

                                                Call ShowConsoleMsg("Hay un objeto en esa posici�n!")

                                                Call StopDragInv

                                                Exit Sub

                                        End If

                                End If

                                If Shift = 1 Then
                                        frmCantidadDrop.Show , frmMain

                                        Call frmCantidadDrop.GetPos(tX, tY, selInvSlot)

                                Else

                                        'Send the package.
                                        Call WriteDropObj(selInvSlot, tX, tY, 1)

                                End If

                                'Reset the flag.
                                Call StopDragInv

                        End If

                End With

        End If
    
End Sub

Private Sub StopDragInv()
        ' GSZAO
        UsabaDrag = False
        UsandoDrag = False
        '        If frmMain.picInv.MousePointer <> vbNormal Then
        'Call ChangeCursorMain(cur_Normal)
        frmMain.PicInv.MousePointer = vbDefault

        ' End If
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
        KeyCode = 0

End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
        KeyAscii = 0

End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0

End Sub

Private Sub lblDropGold_Click()

        Inventario.SelectGold

        If UserGLD > 0 Then
                If Not Comerciando Then frmCantidad.Show , frmMain

        End If
    
End Sub

Private Sub Label4_Click()
        Call Audio.PlayWave(SND_CLICK)

        InvEqu.Picture = LoadPicture(DirInterfaces & "Centroinventario.jpg")
        
        panelFlag = eVentanas.vInventario

        If panelFlag <> lastPanelFlag Then

                Call WriteSetMenu(panelFlag, 255)
                lastPanelFlag = panelFlag

        End If
        
        ' Activo controles de inventario
        PicInv.Visible = True

        ' Desactivo controles de hechizo
        hlst.Visible = False
        cmdINFO.Visible = False
        CmdLanzar.Visible = False
    
        cmdMoverHechi(0).Visible = False
        cmdMoverHechi(1).Visible = False
        
        UsandoDrag = False
    
End Sub

Private Sub Label7_Click()
        Call Audio.PlayWave(SND_CLICK)

        InvEqu.Picture = LoadPicture(DirInterfaces & "Centrohechizos.jpg")
        
        panelFlag = eVentanas.vHechizos

        If panelFlag <> lastPanelFlag Then

                Dim TempInv As Integer

                If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
                   TempInv = Inventario.SelectedItem
                   
                Call WriteSetMenu(panelFlag, CByte(TempInv))
                lastPanelFlag = panelFlag

        End If
        
        ' Activo controles de hechizos
        hlst.Visible = True
        cmdINFO.Visible = True
        CmdLanzar.Visible = True
    
        cmdMoverHechi(0).Visible = True
        cmdMoverHechi(1).Visible = True
    
        ' Desactivo controles de inventario
        PicInv.Visible = False
        UsandoDrag = False

End Sub

Private Sub picInv_DblClick()

        ' x button COMPEUBA LOS TRES PASOS DEL CLICK NO SOLO DEL X BOOTON SINO TAMBIEN ASI DE TODOS LOS PROGRAMAS QUE SALTEAN LOS PASOS DE ABAJO MOUSE UP.
        ' EL QUE COPIA ESTO SE MERECE QUE LE TIREN EL SERVER.
        If (mouse_Down <> False) And (mouse_UP = True) Then Exit Sub
      
        mouse_UP = False
        ' x button

        If MirandoCarpinteria Or MirandoHerreria Then Exit Sub
    
        If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
        If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
    
        Call UsarItem(1)

        UsandoDrag = False

End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

        '    / x button
        If (mouse_Down = False) Then Exit Sub
        mouse_Down = False
        mouse_UP = True
        '    / x button

        Call Audio.PlayWave(SND_CLICK)

End Sub

Private Sub SendTxt_Change()

        '**************************************************************
        'Author: Unknown
        'Last Modify Date: 3/06/2006
        '3/06/2006: Maraxus - imped� se inserten caract�res no imprimibles
        '**************************************************************
        If Len(SendTxt.Text) > 160 Then
                stxtbuffer = "Soy un cheater, avisenle a un gm"
        Else

                'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
                Dim i         As Long

                Dim tempstr   As String

                Dim CharAscii As Integer
        
                For i = 1 To Len(SendTxt.Text)
                        CharAscii = Asc(mid$(SendTxt.Text, i, 1))

                        If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                                tempstr = tempstr & Chr$(CharAscii)

                        End If

                Next i
        
                If tempstr <> SendTxt.Text Then
                        'We only set it if it's different, otherwise the event will be raised
                        'constantly and the client will crush
                        SendTxt.Text = tempstr

                End If
        
                stxtbuffer = SendTxt.Text

        End If

End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)

        If Not (KeyAscii = vbKeyBack) And _
           Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
           KeyAscii = 0

End Sub

Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)

        'Send text
        If KeyCode = vbKeyReturn Then

                'Say
                If stxtbuffercmsg <> "" Then
                        Call ParseUserCommand("/CMSG " & stxtbuffercmsg)

                End If

                stxtbuffercmsg = ""
                SendCMSTXT.Text = ""
                KeyCode = 0
                Me.SendCMSTXT.Visible = False
        
                If PicInv.Visible Then
                        PicInv.SetFocus
                Else
                        hlst.SetFocus

                End If

        End If

End Sub

Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)

        If Not (KeyAscii = vbKeyBack) And _
           Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
           KeyAscii = 0

End Sub

Private Sub SendCMSTXT_Change()

        If Len(SendCMSTXT.Text) > 160 Then
                stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
        Else

                'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
                Dim i         As Long

                Dim tempstr   As String

                Dim CharAscii As Integer
        
                For i = 1 To Len(SendCMSTXT.Text)
                        CharAscii = Asc(mid$(SendCMSTXT.Text, i, 1))

                        If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                                tempstr = tempstr & Chr$(CharAscii)

                        End If

                Next i
        
                If tempstr <> SendCMSTXT.Text Then
                        'We only set it if it's different, otherwise the event will be raised
                        'constantly and the client will crush
                        SendCMSTXT.Text = tempstr

                End If
        
                stxtbuffercmsg = SendCMSTXT.Text

        End If

End Sub

''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''

Private Sub Socket1_Connect()
    
        'Clean input and output buffers
        Call incomingData.ReadASCIIStringFixed(incomingData.Length)
        Call outgoingData.ReadASCIIStringFixed(outgoingData.Length)
    
        Second.Enabled = True

        Select Case EstadoLogin

                Case E_MODO.CrearNuevoPj
                        Call Login
        
                Case E_MODO.Normal
                        Call Login
                        
                Case E_MODO.Cp
                        'MsgBox "Conecte"
                        Dim i As Long
        
                        Call Audio.PlayMIDI("7.mid")
                        frmCrearPersonaje.Show vbModal
        
                        With frmCrearPersonaje

                                If .Visible Then

                                        For i = 1 To NUMATRIBUTES
                                                .lblAtributos(i).Caption = 18
                                        Next i
                
                                        .UpdateStats

                                End If

                        End With
        
        End Select

End Sub

Private Sub Socket1_Disconnect()
        ResetAllInfo
        Socket1.Cleanup

End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, _
                              ErrorString As String, _
                              Response As Integer)

        '*********************************************
        'Handle socket errors
        '*********************************************
        Select Case ErrorCode

                Case TOO_FAST 'jajasAJ CUALQUEIRA AJJAJA
                        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
                        Exit Sub

                Case REFUSED 'Vivan las negradas
                        Call MsgBox("El servidor se encuentra cerrado o no te has podido conectar correctamente.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")

                Case TIME_OUT
                        Call MsgBox("El tiempo de espera se ha agotado, intenta nuevamente.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")

                Case Else
                        Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")

        End Select
    
        frmConnect.MousePointer = 1
        Response = 0

        frmMain.Socket1.Disconnect

End Sub

Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)

        Dim RD     As String

        Dim Data() As Byte
    
        Call Socket1.Read(RD, DataLength)
        Data = StrConv(RD, vbFromUnicode)
    
        If Len(RD) = 0 Then Exit Sub

        'Put data in the buffer
        Call incomingData.WriteBlock(Data)
    
        'Send buffer to Handle data
        Call HandleIncomingData

End Sub

Private Function InGameArea() As Boolean

        '***************************************************
        'Author: NicoNZ
        'Last Modification: 04/07/08
        'Checks if last click was performed within or outside the game area.
        '***************************************************
        If clicX < MainViewShp.Left Or clicX > MainViewShp.Left + MainViewShp.Width Then Exit Function
        If clicY < MainViewShp.Top Or clicY > MainViewShp.Top + MainViewShp.Height Then Exit Function
    
        InGameArea = True

End Function

Private Function BuscarI(Gh As Integer) As Integer

        Dim i As Long

        For i = 1 To frmMain.ImageList1.ListImages.Count

                If frmMain.ImageList1.ListImages(i).key = "g" & CStr(Gh) Then
                        BuscarI = i

                        Exit For

                End If

        Next i

End Function
