VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XP Container Demo"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10755
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   10755
   StartUpPosition =   2  'CenterScreen
   Begin Project1.XPContainer XPContainer12 
      Height          =   2175
      Left            =   8092
      TabIndex        =   1
      Top             =   4755
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3836
      HeaderLightColor=   15526633
      HeaderDarkColor =   14276307
      BackLightColor  =   16316407
      BackDarkColor   =   15855598
      BorderColor     =   14078671
      TextColor       =   4867908
      Caption         =   "Options"
      Theme           =   10
   End
   Begin Project1.XPContainer XPContainer3 
      Height          =   2175
      Left            =   127
      TabIndex        =   0
      Top             =   4755
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3836
      HeaderLightColor=   15852488
      HeaderDarkColor =   14993810
      BackLightColor  =   16446955
      BackDarkColor   =   16116437
      BorderColor     =   14861448
      TextColor       =   5585152
      Caption         =   "Options"
      Theme           =   4
   End
   Begin Project1.XPContainer XPContainer2 
      Height          =   2175
      Left            =   127
      TabIndex        =   2
      Top             =   2460
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3836
      Caption         =   "Options"
   End
   Begin Project1.XPContainer XPContainer6 
      Height          =   2175
      Left            =   2782
      TabIndex        =   3
      Top             =   4755
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3836
      HeaderLightColor=   14934998
      HeaderDarkColor =   13224366
      BackLightColor  =   16119280
      BackDarkColor   =   15395552
      BorderColor     =   12895398
      TextColor       =   3750173
      Caption         =   "Options"
      Theme           =   7
   End
   Begin Project1.XPContainer XPContainer9 
      Height          =   2175
      Left            =   5437
      TabIndex        =   4
      Top             =   4755
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3836
      HeaderLightColor=   13820669
      HeaderDarkColor =   10995450
      BackLightColor  =   15726078
      BackDarkColor   =   14543357
      BorderColor     =   10469626
      TextColor       =   1455725
      Caption         =   "Options"
      Theme           =   6
   End
   Begin Project1.XPContainer XPContainer4 
      Height          =   2175
      Left            =   2775
      TabIndex        =   5
      Top             =   165
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3836
      HeaderLightColor=   14214600
      HeaderDarkColor =   11651986
      BackLightColor  =   15857131
      BackDarkColor   =   14805973
      BorderColor     =   11191944
      TextColor       =   2177792
      Caption         =   "Options"
      Theme           =   2
   End
   Begin Project1.XPContainer XPContainer7 
      Height          =   2175
      Left            =   5430
      TabIndex        =   6
      Top             =   165
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3836
      HeaderLightColor=   14349027
      HeaderDarkColor =   11920840
      BackLightColor  =   15858421
      BackDarkColor   =   14939626
      BorderColor     =   11461571
      TextColor       =   2381624
      Caption         =   "Options"
      Theme           =   5
   End
   Begin Project1.XPContainer XPContainer10 
      Height          =   2175
      Left            =   8085
      TabIndex        =   7
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3836
      HeaderLightColor=   15006460
      HeaderDarkColor =   12185332
      BackLightColor  =   15662333
      BackDarkColor   =   14481402
      BorderColor     =   9822698
      TextColor       =   6739425
      Caption         =   "Options"
      Theme           =   11
   End
   Begin Project1.XPContainer XPContainer1 
      Height          =   2175
      Left            =   127
      TabIndex        =   8
      Top             =   165
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3836
      HeaderLightColor=   15523027
      HeaderDarkColor =   14334632
      BackLightColor  =   16315119
      BackDarkColor   =   15853021
      BorderColor     =   14070944
      TextColor       =   4925975
      Caption         =   "Options"
      Theme           =   1
   End
   Begin Project1.XPContainer XPContainer5 
      Height          =   2175
      Left            =   2782
      TabIndex        =   9
      Top             =   2460
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3836
      HeaderLightColor=   14740200
      HeaderDarkColor =   12768977
      BackLightColor  =   16054519
      BackDarkColor   =   15200237
      BorderColor     =   12374989
      TextColor       =   3295041
      Caption         =   "Options"
      Theme           =   3
   End
   Begin Project1.XPContainer XPContainer8 
      Height          =   2175
      Left            =   5437
      TabIndex        =   10
      Top             =   2460
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3836
      HeaderLightColor=   14078715
      HeaderDarkColor =   11446008
      BackLightColor  =   15790078
      BackDarkColor   =   14736892
      BorderColor     =   10985207
      TextColor       =   1906026
      Caption         =   "Options"
      Theme           =   9
   End
   Begin Project1.XPContainer XPContainer11 
      Height          =   2175
      Left            =   8092
      TabIndex        =   11
      Top             =   2460
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3836
      HeaderLightColor=   15390687
      HeaderDarkColor =   14004415
      BackLightColor  =   16249331
      BackDarkColor   =   15720934
      BorderColor     =   13740473
      TextColor       =   4595759
      Caption         =   "Options"
      Theme           =   8
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
