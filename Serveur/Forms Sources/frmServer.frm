VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL3N.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCN.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FRoG Server"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10245
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   683
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4800
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   8467
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   370
      TabMaxWidth     =   3175
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Discussions"
      TabPicture(0)   =   "frmServer.frx":17D2A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "mesgSaveTimer"
      Tab(0).Control(1)=   "tmrChatLogs"
      Tab(0).Control(2)=   "picCMsg"
      Tab(0).Control(3)=   "SSTab2"
      Tab(0).Control(4)=   "Say(5)"
      Tab(0).Control(5)=   "Say(4)"
      Tab(0).Control(6)=   "Say(3)"
      Tab(0).Control(7)=   "Say(2)"
      Tab(0).Control(8)=   "Say(1)"
      Tab(0).Control(9)=   "CustomMsg(5)"
      Tab(0).Control(10)=   "CustomMsg(4)"
      Tab(0).Control(11)=   "CustomMsg(3)"
      Tab(0).Control(12)=   "CustomMsg(2)"
      Tab(0).Control(13)=   "CustomMsg(1)"
      Tab(0).Control(14)=   "Frame5"
      Tab(0).Control(15)=   "Say(0)"
      Tab(0).Control(16)=   "CustomMsg(0)"
      Tab(0).Control(17)=   "mesgSave"
      Tab(0).Control(18)=   "Image1"
      Tab(0).Control(19)=   "Label6"
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Joueur"
      TabPicture(1)   =   "frmServer.frx":17D46
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TPO"
      Tab(1).Control(1)=   "lvUsers"
      Tab(1).Control(2)=   "Command66"
      Tab(1).Control(3)=   "Check1"
      Tab(1).Control(4)=   "Command14"
      Tab(1).Control(5)=   "Command51"
      Tab(1).Control(6)=   "picStats"
      Tab(1).Control(7)=   "Picskint"
      Tab(1).Control(8)=   "picReason"
      Tab(1).Control(9)=   "Command9"
      Tab(1).Control(10)=   "picWarp"
      Tab(1).Control(11)=   "picJail"
      Tab(1).Control(12)=   "picExp"
      Tab(1).Control(13)=   "Command16"
      Tab(1).Control(14)=   "Command18"
      Tab(1).Control(15)=   "Command19"
      Tab(1).Control(16)=   "Command21"
      Tab(1).Control(17)=   "Command22"
      Tab(1).Control(18)=   "Command23"
      Tab(1).Control(19)=   "Command24"
      Tab(1).Control(20)=   "Command31"
      Tab(1).Control(21)=   "Command3"
      Tab(1).Control(22)=   "Command12"
      Tab(1).Control(23)=   "Command45"
      Tab(1).Control(24)=   "Command32"
      Tab(1).Control(25)=   "Command46"
      Tab(1).Control(26)=   "Command33"
      Tab(1).Control(27)=   "Command34"
      Tab(1).ControlCount=   28
      TabCaption(2)   =   "Panneau de Contr�le"
      TabPicture(2)   =   "frmServer.frx":17D62
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblPort"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblIP"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblgamename"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lbliplocal"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame2"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame8"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Frame1"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Frame9"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Frame4"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "picWeather"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "Aide"
      TabPicture(3)   =   "frmServer.frx":17D7E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "CharInfo(21)"
      Tab(3).Control(1)=   "CharInfo(22)"
      Tab(3).Control(2)=   "CharInfo(23)"
      Tab(3).Control(3)=   "lstTopics"
      Tab(3).Control(4)=   "TopicTitle"
      Tab(3).Control(5)=   "sync"
      Tab(3).ControlCount=   6
      Begin VB.CommandButton Command34 
         Caption         =   "+1 Niveau (Tous)"
         Height          =   255
         Left            =   -66840
         TabIndex        =   144
         Top             =   4320
         Width           =   1935
      End
      Begin VB.CommandButton Command33 
         Caption         =   "Exp�rience (Tous)"
         Height          =   255
         Left            =   -66840
         TabIndex        =   143
         Top             =   4080
         Width           =   1935
      End
      Begin VB.CommandButton Command46 
         Caption         =   "Sauvegarder(Tous)"
         Height          =   255
         Left            =   -66840
         TabIndex        =   145
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton Command32 
         Caption         =   "T�l�porter (Tous)"
         Height          =   255
         Left            =   -66840
         TabIndex        =   142
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CommandButton Command45 
         Caption         =   "T�l�porter"
         Height          =   255
         Left            =   -66840
         TabIndex        =   98
         Top             =   3360
         Width           =   1935
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Gu�rir (Tous)"
         Height          =   255
         Left            =   -66840
         TabIndex        =   141
         Top             =   3120
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Gu�rir"
         Height          =   255
         Left            =   -66840
         TabIndex        =   27
         Top             =   2880
         Width           =   1935
      End
      Begin VB.CommandButton Command31 
         Caption         =   "Tuer (Tous)"
         Height          =   255
         Left            =   -66840
         TabIndex        =   140
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Tuer"
         Height          =   255
         Left            =   -66840
         TabIndex        =   15
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton Command23 
         Caption         =   "D�sactiv� mode Muet"
         Height          =   255
         Left            =   -66840
         TabIndex        =   14
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Mode Muet"
         Height          =   255
         Left            =   -66840
         TabIndex        =   13
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Message Priv�"
         Height          =   255
         Left            =   -66840
         TabIndex        =   12
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Voir informations"
         Height          =   255
         Left            =   -66840
         TabIndex        =   11
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Prison (Raison)"
         Height          =   255
         Left            =   -66840
         TabIndex        =   10
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Bannir (Raison)"
         Height          =   255
         Left            =   -66840
         TabIndex        =   9
         Top             =   960
         Width           =   1935
      End
      Begin VB.PictureBox picWeather 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   4920
         ScaleHeight     =   135
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   247
         TabIndex        =   111
         Top             =   2160
         Visible         =   0   'False
         Width           =   3735
         Begin VB.CommandButton CommandMeteoMap 
            Caption         =   "D�sactiver la m�t�o par map"
            Height          =   255
            Left            =   240
            TabIndex        =   164
            Top             =   1080
            Width           =   3255
         End
         Begin VB.CommandButton Command65 
            Caption         =   "Neige"
            Height          =   255
            Left            =   2160
            TabIndex        =   117
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton Command64 
            Caption         =   "Pluie"
            Height          =   255
            Left            =   240
            TabIndex        =   116
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton Command63 
            Caption         =   "Orage"
            Height          =   255
            Left            =   2160
            TabIndex        =   115
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton Command62 
            Caption         =   "Soleil"
            Height          =   255
            Left            =   240
            TabIndex        =   114
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton Command61 
            Caption         =   "Fermer"
            Height          =   255
            Left            =   2160
            TabIndex        =   112
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "La m�t�o par map est activ�e actuellement"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   165
            Top             =   1440
            Width           =   3240
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Temps actuel : Soleil"
            Height          =   225
            Left            =   120
            TabIndex        =   113
            Top             =   120
            Width           =   3420
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Ratio d'exp�rience"
         Height          =   1695
         Left            =   2040
         TabIndex        =   167
         Top             =   600
         Width           =   3015
         Begin VB.CommandButton Command13 
            Caption         =   "Valider"
            Enabled         =   0   'False
            Height          =   225
            Left            =   1080
            TabIndex        =   172
            Top             =   1400
            Width           =   975
         End
         Begin VB.HScrollBar HScroll2 
            Height          =   255
            Left            =   120
            Max             =   100
            Min             =   1
            TabIndex        =   171
            Top             =   1080
            Value           =   1
            Width           =   2775
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   120
            Max             =   100
            Min             =   1
            TabIndex        =   169
            Top             =   480
            Value           =   1
            Width           =   2775
         End
         Begin VB.Label Label13 
            Caption         =   "Ratio d'exp�rience qu�te : 1"
            Height          =   255
            Left            =   240
            TabIndex        =   170
            Top             =   840
            Width           =   2655
         End
         Begin VB.Label Label7 
            Caption         =   "Ratio d'exp�rience normal : 1"
            Height          =   255
            Left            =   240
            TabIndex        =   168
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Liste des Cartes"
         Height          =   1815
         Left            =   5160
         TabIndex        =   96
         Top             =   480
         Width           =   4935
         Begin VB.CommandButton Command36 
            Caption         =   "Informations sur la carte"
            Height          =   255
            Left            =   2520
            TabIndex        =   163
            Top             =   1320
            Width           =   2175
         End
         Begin VB.CommandButton Command35 
            Caption         =   "Rafra�chir"
            Height          =   255
            Left            =   120
            TabIndex        =   162
            Top             =   1320
            Width           =   2175
         End
         Begin VB.ListBox MapList 
            Height          =   960
            Left            =   120
            TabIndex        =   97
            Top             =   240
            Width           =   4575
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Scripts"
         Height          =   1575
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   1815
         Begin VB.CommandButton Command27 
            Caption         =   "Activer"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton Command26 
            Caption         =   "D�sactiver"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton Command25 
            Caption         =   "Recharger"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblscript 
            Caption         =   "Etat : d�sactiv�"
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   240
            TabIndex        =   174
            Top             =   1200
            Width           =   1335
         End
      End
      Begin VB.PictureBox picExp 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   -70200
         ScaleHeight     =   87
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   215
         TabIndex        =   148
         Top             =   2880
         Visible         =   0   'False
         Width           =   3255
         Begin VB.CommandButton Command40 
            Caption         =   "OK"
            Height          =   255
            Left            =   1560
            TabIndex        =   151
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtExp 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   150
            Top             =   360
            Width           =   2955
         End
         Begin VB.CommandButton Command39 
            Caption         =   "Annuler"
            Height          =   255
            Left            =   1560
            TabIndex        =   149
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exp�rience:"
            Height          =   195
            Left            =   120
            TabIndex        =   152
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.PictureBox picJail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   -70320
         ScaleHeight     =   167
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   80
         Top             =   1200
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton Command11 
            Caption         =   "Annuler"
            Height          =   255
            Left            =   1680
            TabIndex        =   95
            Top             =   2160
            Width           =   1575
         End
         Begin VB.HScrollBar scrlY 
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   1560
            Width           =   3135
         End
         Begin VB.HScrollBar scrlX 
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   960
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMap 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   82
            Top             =   360
            Value           =   1
            Width           =   3135
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Emprisonner"
            Height          =   255
            Left            =   1680
            TabIndex        =   81
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label txtY 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   87
            Top             =   1320
            Width           =   285
         End
         Begin VB.Label txtX 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   86
            Top             =   720
            Width           =   285
         End
         Begin VB.Label txtMap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carte : 1"
            Height          =   225
            Left            =   120
            TabIndex        =   85
            Top             =   120
            Width           =   645
         End
      End
      Begin VB.PictureBox picWarp 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   -70320
         ScaleHeight     =   167
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   215
         TabIndex        =   153
         Top             =   1200
         Visible         =   0   'False
         Width           =   3255
         Begin VB.CommandButton Command37 
            Caption         =   "T�l�porter"
            Height          =   255
            Left            =   1560
            TabIndex        =   158
            Top             =   1920
            Width           =   1575
         End
         Begin VB.HScrollBar scrlMM 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   157
            Top             =   360
            Value           =   1
            Width           =   3015
         End
         Begin VB.HScrollBar scrlMX 
            Height          =   255
            Left            =   120
            TabIndex        =   156
            Top             =   960
            Width           =   3015
         End
         Begin VB.HScrollBar scrlMY 
            Height          =   255
            Left            =   120
            TabIndex        =   155
            Top             =   1560
            Width           =   3015
         End
         Begin VB.CommandButton Command38 
            Caption         =   "Annuler"
            Height          =   255
            Left            =   1560
            TabIndex        =   154
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label lblMM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Map: 1"
            Height          =   195
            Left            =   120
            TabIndex        =   161
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lblMX 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   160
            Top             =   720
            Width           =   285
         End
         Begin VB.Label lblMY 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   159
            Top             =   1320
            Width           =   285
         End
      End
      Begin VB.CommandButton Command9 
         Caption         =   "D�connecter (Tous)"
         Height          =   255
         Left            =   -66840
         TabIndex        =   139
         Top             =   720
         Width           =   1935
      End
      Begin VB.Timer mesgSaveTimer 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   -68520
         Top             =   240
      End
      Begin VB.Frame Frame8 
         Caption         =   "Gestion du temps :"
         Height          =   2415
         Left            =   5160
         TabIndex        =   129
         Top             =   2280
         Width           =   4935
         Begin VB.CommandButton Command58 
            Caption         =   "Basculer Jour/Nuit"
            Height          =   255
            Left            =   2640
            TabIndex        =   147
            Top             =   1680
            Width           =   1935
         End
         Begin VB.CommandButton Command59 
            Caption         =   "Changer la M�t�o"
            Height          =   255
            Left            =   360
            TabIndex        =   146
            Top             =   1680
            Width           =   2055
         End
         Begin VB.CommandButton Command50 
            Caption         =   "D�sactiver"
            Height          =   255
            Left            =   3240
            TabIndex        =   135
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton Command49 
            Caption         =   "D�sactiver"
            Height          =   255
            Left            =   3360
            TabIndex        =   134
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtRandom 
            Height          =   285
            Left            =   240
            TabIndex        =   133
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Timer tmrRandom 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   8400
            Top             =   120
         End
         Begin VB.CommandButton Command48 
            Caption         =   "Activer"
            Height          =   255
            Left            =   2040
            TabIndex        =   132
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Timer tmrTemps 
            Enabled         =   0   'False
            Interval        =   60000
            Left            =   5160
            Top             =   120
         End
         Begin VB.CommandButton Command47 
            Caption         =   "D�finir"
            Height          =   255
            Left            =   2040
            TabIndex        =   131
            Top             =   480
            Width           =   1095
         End
         Begin VB.Timer tmrJournuit 
            Enabled         =   0   'False
            Interval        =   60000
            Left            =   2400
            Top             =   120
         End
         Begin VB.TextBox txtJournuit 
            Height          =   285
            Left            =   240
            TabIndex        =   130
            Top             =   480
            Width           =   1695
         End
         Begin Serveur.ctlSysTrayBalloon ctlSysTrayBalloon 
            Left            =   9000
            Top             =   120
            _ExtentX        =   2143
            _ExtentY        =   661
         End
         Begin VB.Label Label10 
            Caption         =   "Temps al�atoire ( minutes ) :"
            Height          =   255
            Left            =   120
            TabIndex        =   137
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label Label8 
            Caption         =   "Cycle jour / nuit ( minutes ) :"
            Height          =   255
            Left            =   240
            TabIndex        =   136
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Timer sync 
         Interval        =   5000
         Left            =   -68520
         Top             =   0
      End
      Begin VB.PictureBox picReason 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   -70320
         ScaleHeight     =   87
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   53
         Top             =   360
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton Command6 
            Caption         =   "Annuler"
            Height          =   255
            Left            =   1680
            TabIndex        =   94
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Caption"
            Height          =   255
            Left            =   1680
            TabIndex        =   55
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtReason 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   360
            Width           =   3075
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Raison (falcultatif) :"
            Height          =   225
            Left            =   120
            TabIndex        =   56
            Top             =   120
            Width           =   1530
         End
      End
      Begin MSWinsockLib.Winsock Socket 
         Index           =   0
         Left            =   9840
         Top             =   -14
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.PictureBox Picskint 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   -69480
         ScaleHeight     =   167
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   127
         TabIndex        =   122
         Top             =   1080
         Visible         =   0   'False
         Width           =   1935
         Begin VB.CommandButton Command52 
            Caption         =   "Annuler"
            Height          =   255
            Left            =   240
            TabIndex        =   123
            Top             =   2160
            Width           =   1575
         End
         Begin VB.OptionButton grand 
            Caption         =   "32/48 pixels"
            Height          =   255
            Left            =   120
            TabIndex        =   126
            Top             =   960
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton petit 
            Caption         =   "32/32 pixels"
            Height          =   255
            Left            =   120
            TabIndex        =   125
            Top             =   600
            Width           =   1455
         End
         Begin VB.CommandButton Command53 
            Caption         =   "Enregistrer"
            Height          =   255
            Left            =   240
            TabIndex        =   124
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Label11 
            Caption         =   "S�l�ctioner une taille (Largeur/Hauteur) :"
            Height          =   435
            Left            =   120
            TabIndex        =   127
            Top             =   120
            Width           =   1755
         End
      End
      Begin VB.PictureBox picStats 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   -74760
         ScaleHeight     =   215
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   311
         TabIndex        =   57
         Top             =   360
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CommandButton Command8 
            Caption         =   "Annuler"
            Height          =   255
            Left            =   3000
            TabIndex        =   58
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Index:"
            Height          =   195
            Index           =   20
            Left            =   2400
            TabIndex        =   79
            Top             =   1800
            Width           =   480
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Points:"
            Height          =   195
            Index           =   19
            Left            =   2400
            TabIndex        =   78
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Magie:"
            Height          =   195
            Index           =   18
            Left            =   2400
            TabIndex        =   77
            Top             =   1320
            Width           =   480
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vitesse:"
            Height          =   195
            Index           =   17
            Left            =   2400
            TabIndex        =   76
            Top             =   1080
            Width           =   570
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Def:"
            Height          =   195
            Index           =   16
            Left            =   2400
            TabIndex        =   75
            Top             =   840
            Width           =   315
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "For:"
            Height          =   195
            Index           =   15
            Left            =   2400
            TabIndex        =   74
            Top             =   600
            Width           =   300
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Acc�s Guilde :"
            Height          =   225
            Index           =   14
            Left            =   2400
            TabIndex        =   73
            Top             =   360
            Width           =   1110
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Guilde:"
            Height          =   195
            Index           =   13
            Left            =   2400
            TabIndex        =   72
            Top             =   120
            Width           =   495
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Map:"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   71
            Top             =   3000
            Width           =   360
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sexe:"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   70
            Top             =   2760
            Width           =   420
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sprite:"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   69
            Top             =   2520
            Width           =   480
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Classe:"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   68
            Top             =   2280
            Width           =   525
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PK:"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   67
            Top             =   2040
            Width           =   240
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Acc�s :"
            Height          =   225
            Index           =   7
            Left            =   120
            TabIndex        =   66
            Top             =   1800
            Width           =   555
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EXP: /"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   65
            Top             =   1560
            Width           =   435
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End: /"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   64
            Top             =   1320
            Width           =   435
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PM: /"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   63
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PV: /"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   62
            Top             =   840
            Width           =   345
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Niveau:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   61
            Top             =   600
            Width           =   555
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Personnage:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   60
            Top             =   360
            Width           =   915
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Compte:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   59
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.Timer tmrShutdown 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   8400
         Top             =   -14
      End
      Begin VB.Timer PlayerTimer 
         Interval        =   5000
         Left            =   7920
         Top             =   -14
      End
      Begin VB.Timer tmrPlayerSave 
         Interval        =   60000
         Left            =   7440
         Top             =   -14
      End
      Begin VB.CommandButton Command51 
         Caption         =   "R�gler taille des Skins"
         Height          =   255
         Left            =   -69000
         TabIndex        =   121
         Top             =   4440
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Timer tmrChatLogs 
         Interval        =   1000
         Left            =   -65160
         Top             =   360
      End
      Begin VB.PictureBox picCMsg 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   -70680
         ScaleHeight     =   1905
         ScaleWidth      =   3345
         TabIndex        =   44
         Top             =   960
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox txtMsg 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   960
            Width           =   3075
         End
         Begin VB.TextBox txtTitle 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            MaxLength       =   13
            TabIndex        =   49
            Top             =   360
            Width           =   3075
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Annuler"
            Height          =   255
            Left            =   1680
            TabIndex        =   46
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Enregistrer"
            Height          =   255
            Left            =   1680
            TabIndex        =   45
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Message:"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Titre:"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   120
            Width           =   390
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   101
         Top             =   360
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5318
         _Version        =   393216
         Style           =   1
         Tabs            =   7
         TabsPerRow      =   7
         TabHeight       =   353
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "G�n�ral"
         TabPicture(0)   =   "frmServer.frx":17D9A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "txtChat"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtText(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "�mission"
         TabPicture(1)   =   "frmServer.frx":17DB6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtText(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Globale"
         TabPicture(2)   =   "frmServer.frx":17DD2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtText(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Carte"
         TabPicture(3)   =   "frmServer.frx":17DEE
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtText(3)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Priv�"
         TabPicture(4)   =   "frmServer.frx":17E0A
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "txtText(4)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Administrateur"
         TabPicture(5)   =   "frmServer.frx":17E26
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "txtText(5)"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "Emote"
         TabPicture(6)   =   "frmServer.frx":17E42
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "txtText(6)"
         Tab(6).ControlCount=   1
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2250
            Index           =   0
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   128
            Top             =   360
            Width           =   8115
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   6
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   108
            Top             =   360
            Width           =   8115
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   5
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   107
            Top             =   360
            Width           =   8115
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   4
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   106
            Top             =   360
            Width           =   8115
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   3
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   105
            Top             =   360
            Width           =   8115
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   2
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   104
            Top             =   360
            Width           =   8115
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   1
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   103
            Top             =   360
            Width           =   8115
         End
         Begin VB.TextBox txtChat 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   102
            Top             =   2640
            Width           =   8115
         End
      End
      Begin VB.Frame TopicTitle 
         Caption         =   "Titre du Topics:"
         Height          =   4335
         Left            =   -72480
         TabIndex        =   90
         Top             =   360
         Width           =   7575
         Begin VB.TextBox txtTopic 
            Height          =   3975
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   91
            Top             =   240
            Width           =   7335
         End
      End
      Begin VB.ListBox lstTopics 
         Height          =   2760
         ItemData        =   "frmServer.frx":17E5E
         Left            =   -74760
         List            =   "frmServer.frx":17E60
         TabIndex        =   88
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton Say 
         Caption         =   "Dire"
         Height          =   255
         Index           =   5
         Left            =   -66360
         TabIndex        =   43
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton Say 
         Caption         =   "Dire"
         Height          =   255
         Index           =   4
         Left            =   -66360
         TabIndex        =   42
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton Say 
         Caption         =   "Dire"
         Height          =   255
         Index           =   3
         Left            =   -66360
         TabIndex        =   41
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton Say 
         Caption         =   "Dire"
         Height          =   255
         Index           =   2
         Left            =   -66360
         TabIndex        =   40
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton Say 
         Caption         =   "Dire"
         Height          =   255
         Index           =   1
         Left            =   -66360
         TabIndex        =   39
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton CustomMsg 
         Caption         =   "Editer msg"
         Height          =   255
         Index           =   5
         Left            =   -66360
         TabIndex        =   38
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton CustomMsg 
         Caption         =   "Editer msg"
         Height          =   255
         Index           =   4
         Left            =   -66360
         TabIndex        =   37
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CommandButton CustomMsg 
         Caption         =   "Editer msg"
         Height          =   255
         Index           =   3
         Left            =   -66360
         TabIndex        =   36
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton CustomMsg 
         Caption         =   "Editer msg"
         Height          =   255
         Index           =   2
         Left            =   -66360
         TabIndex        =   35
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton CustomMsg 
         Caption         =   "Editer msg"
         Height          =   255
         Index           =   1
         Left            =   -66360
         TabIndex        =   34
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Frame Frame5 
         Caption         =   "Configuration des discussions :"
         Height          =   615
         Left            =   -74880
         TabIndex        =   28
         Top             =   3600
         Width           =   6855
         Begin VB.CommandButton Command60 
            Caption         =   "Enregistrer"
            Height          =   255
            Left            =   5280
            TabIndex        =   109
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox chkA 
            Caption         =   "Admin"
            Height          =   375
            Left            =   4320
            TabIndex        =   52
            Top             =   170
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkG 
            Caption         =   "Globale"
            Height          =   255
            Left            =   3360
            TabIndex        =   51
            Top             =   240
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkP 
            Caption         =   "Priv�"
            Height          =   255
            Left            =   2640
            TabIndex        =   32
            Top             =   240
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkM 
            Caption         =   "Carte"
            Height          =   255
            Left            =   1920
            TabIndex        =   31
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkE 
            Caption         =   "Emote"
            Height          =   255
            Left            =   1100
            TabIndex        =   30
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkBC 
            Caption         =   "�mission"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Value           =   1  'Checked
            Width           =   1095
         End
      End
      Begin VB.Timer tmrSpawnMapItems 
         Interval        =   1000
         Left            =   9360
         Top             =   -14
      End
      Begin VB.Timer tmrGameAI 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   8880
         Top             =   -14
      End
      Begin VB.Frame Frame3 
         Caption         =   "Classes"
         Height          =   1095
         Left            =   120
         TabIndex        =   21
         Top             =   3600
         Width           =   1815
         Begin VB.CommandButton Command30 
            Caption         =   "Modifier"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   600
            Width           =   1455
         End
         Begin VB.CommandButton Command29 
            Caption         =   "Recharger"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Serveur"
         Height          =   2415
         Left            =   2040
         TabIndex        =   20
         Top             =   2280
         Width           =   3015
         Begin VB.CheckBox chkChat 
            Caption         =   "Sauvegarder les logs"
            Height          =   255
            Left            =   240
            TabIndex        =   110
            Top             =   960
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CheckBox mnuServerLog 
            Caption         =   "Logs du serveur"
            Height          =   255
            Left            =   240
            TabIndex        =   100
            Top             =   720
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox Closed 
            Caption         =   "Fermer"
            Height          =   255
            Left            =   240
            TabIndex        =   99
            Top             =   480
            Width           =   2295
         End
         Begin VB.CheckBox GMOnly 
            Caption         =   "Ma�tre de jeu seulement "
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   240
            Width           =   2415
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Fermer"
            Height          =   255
            Left            =   1680
            TabIndex        =   25
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "�teindre"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lblScriptAlert 
            Caption         =   "Erreur de script : merci de les recharger"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   470
            Left            =   120
            TabIndex        =   166
            Top             =   1920
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.Label ShutdownTime 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fermeture : d�sactiv�"
            Height          =   225
            Left            =   720
            TabIndex        =   26
            Top             =   1680
            Width           =   1680
         End
      End
      Begin VB.CommandButton Command14 
         Caption         =   "D�connecter (Raison)"
         Height          =   255
         Left            =   -66840
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Grille"
         Height          =   255
         Left            =   -69840
         TabIndex        =   4
         Top             =   4440
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CommandButton Say 
         Caption         =   "Dire"
         Height          =   255
         Index           =   0
         Left            =   -66360
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton CustomMsg 
         Caption         =   "Editer msg"
         Height          =   255
         Index           =   0
         Left            =   -66360
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command66 
         Caption         =   "Rafra�chir"
         Height          =   255
         Left            =   -71520
         TabIndex        =   119
         Top             =   4440
         Width           =   1575
      End
      Begin MSComctlLib.ListView lvUsers 
         Height          =   3855
         Left            =   -74760
         TabIndex        =   3
         Top             =   480
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   6800
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Compte"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Personnage"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Niveau"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Sprite"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Access"
            Object.Width           =   1235
         EndProperty
      End
      Begin VB.Label lbliplocal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local :"
         Height          =   225
         Left            =   120
         TabIndex        =   175
         Top             =   600
         Width           =   510
      End
      Begin VB.Label lblgamename 
         Caption         =   "Nom du jeu : "
         Height          =   255
         Left            =   120
         TabIndex        =   173
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label mesgSave 
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   -67800
         TabIndex        =   138
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   -74400
         Picture         =   "frmServer.frx":17E62
         Top             =   600
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sauvegarde dans :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   -74760
         TabIndex        =   118
         Top             =   4320
         Width           =   1170
      End
      Begin VB.Label CharInfo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.frogcreator.fr"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   23
         Left            =   -74760
         TabIndex        =   93
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label CharInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pour plus d'information :"
         Height          =   195
         Index           =   22
         Left            =   -74760
         TabIndex        =   92
         Top             =   3720
         Width           =   1740
      End
      Begin VB.Label CharInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sujets :"
         Height          =   195
         Index           =   21
         Left            =   -74760
         TabIndex        =   89
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblIP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adresse IP:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   4920
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   360
      End
      Begin VB.Label TPO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nbr Joueur en ligne:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   5
         Top             =   4440
         Width           =   1455
      End
   End
   Begin VB.Timer Bouclescript 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6960
      Top             =   0
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   495
      Left            =   4800
      TabIndex        =   120
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Menu opt_menu 
      Caption         =   "Options"
      Begin VB.Menu infbull_opt 
         Caption         =   "Infos-bulles"
      End
      Begin VB.Menu colors_opt 
         Caption         =   "Couleurs"
      End
      Begin VB.Menu ftpmaps_opt 
         Caption         =   "Cartes par FTP"
      End
   End
   Begin VB.Menu Dmenu 
      Caption         =   "dMenu"
      Visible         =   0   'False
      Begin VB.Menu ouvr 
         Caption         =   "Agrandir la fen�tre"
      End
      Begin VB.Menu rchrg 
         Caption         =   "Recharger..."
         Begin VB.Menu rchrgcls 
            Caption         =   "Recharger les classes"
         End
         Begin VB.Menu rechrgscr 
            Caption         =   "Recharger les scripts"
         End
      End
      Begin VB.Menu jn 
         Caption         =   "Jour<->Nuit"
      End
      Begin VB.Menu quit 
         Caption         =   "Fermer le serveur"
      End
   End
   Begin VB.Menu fdg 
      Caption         =   "LOGs"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu erg 
         Caption         =   "Log des joueurs (pouvoir les enregistrer)"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mhelp 
      Caption         =   "?"
      Begin VB.Menu mhelpfaq 
         Caption         =   "Aide"
      End
      Begin VB.Menu mwebsite 
         Caption         =   "Site internet de FRoG Creator"
      End
      Begin VB.Menu apopospet 
         Caption         =   "A propos..."
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Dim CM As Long
Dim Num As Long
Dim Random As Long
Dim minuteJN As Long
Dim minuteR As Long
Dim tempjn As Long
Dim tempr As Long

Private Sub apopospet_Click()
frmAbout.Show
End Sub

Private Sub Bouclescript_Timer()
If Scripting = 1 Then MyScript.ExecuteStatement "Scripts\Main.txt", "Boucle"
End Sub

Private Sub CharInfo_Click(Index As Integer)
ShellExecute Me.hwnd, "open", "http://www.frogcreator.fr", "", App.Path, 1
End Sub

Private Sub Check1_Click()
    If Check1.value = Checked Then lvUsers.GridLines = True Else lvUsers.GridLines = False
End Sub

Private Sub colors_opt_Click()
Call ChargOptCoul
End Sub

Private Sub Command1_Click()
If tmrShutdown.Enabled = False Then tmrShutdown.Enabled = True
End Sub

Private Sub Command10_Click()
Dim Index As Long
Dim i As Integer

'Index = lvUsers.ListItems(lvUsers.SelectedItem.Index).text

For i = 1 To lvUsers.ListItems.Count - 1 Step 1
    If lvUsers.ListItems.item(i).Selected = True Then
        Index = lvUsers.ListItems.item(i).text
        


        If Command10.Caption = "T�l�porter" Then
            If Index > 0 Then
                If IsPlaying(Index) Then
                    Call PlayerMsg(Index, "Tu as �t� t�l�port� par l'admin du serveur � la carte :" & scrlMap.value & " X:" & scrlX.value & " Y:" & scrlY.value, White)
                    Call PlayerWarp(Index, scrlMap.value, scrlX.value, scrlY.value)
                End If
            End If
        End If
    
        If Num = 3 Then
            If Index > 0 Then
                If IsPlaying(Index) Then Call GlobalMsg(GetPlayerName(Index) & " a �t� emprisonn� par l'admin du serveur.", White)
        
                Call PlayerWarp(Index, scrlMap.value, scrlX.value, scrlY.value)
            End If
        ElseIf Num = 4 Then
            If txtReason.text = vbNullString Then txtReason.text = "Aucune raison sp�cifi�": Exit Sub
    
            If Index > 0 Then
                If IsPlaying(Index) Then Call GlobalMsg(GetPlayerName(Index) & " a �t� emprisonn� par l'admin du serveur. Raison:(" & txtReason.text & ")", White)
            
                Call PlayerWarp(Index, scrlMap.value, scrlX.value, scrlY.value)
            End If
        End If
    End If
Next i

picReason.Visible = False
picJail.Visible = False
End Sub

Private Sub Command11_Click()
    picJail.Visible = False
    picReason.Visible = False
End Sub

Private Sub Command12_Click()
Dim Index As Long

For Index = 1 To MAX_PLAYERS
    If IsPlaying(Index) = True Then
        Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
        Call SendHP(Index)
        Call PlayerMsg(Index, "Vous avez �t� soign� par l'admin du serveur.", BrightGreen)
    End If
Next Index
End Sub

Private Sub Command13_Click()
    RATE_EXP = HScroll1.value
    RATE_QUETE = HScroll2.value
    PutVar App.Path & "\Data.ini", "RATIO", "RATE_EXP", HScroll1.value
    PutVar App.Path & "\Data.ini", "RATIO", "RATE_QUETE", HScroll2.value
    Command13.Enabled = False
End Sub

Private Sub Command14_Click()
Num = 1
Command7.Caption = "D�connexion"
Label4.Caption = "Raison :"
picReason.Height = 1335
picJail.Visible = False
picReason.Visible = True

Dim i As Long

    Call RemovePLR
    
    For i = 1 To MAX_PLAYERS
        Call ShowPLR(i)
    Next i
    
End Sub

Private Sub Command16_Click()
Num = 2
Command7.Caption = "Bannir"
Label4.Caption = "Raison :"
picReason.Height = 1335
picJail.Visible = False
picReason.Visible = True
End Sub

Private Sub Command18_Click()
Num = 4
Label4.Caption = "Raison :"
Command10.Caption = "Prison"
picReason.Height = 750
scrlMap.Max = MAX_MAPS
scrlX.Max = MAX_MAPX
scrlY.Max = MAX_MAPY
picJail.Visible = True
picReason.Visible = True
End Sub

Private Sub Command19_Click()
Dim Index As Long
If lvUsers.ListItems.Count = 0 Then Exit Sub
Index = lvUsers.ListItems(lvUsers.SelectedItem.Index).text
If Not IsPlaying(Index) Then Exit Sub

    CharInfo(0).Caption = "Compte: " & GetPlayerLogin(Index)
    CharInfo(1).Caption = "Personnage: " & GetPlayerName(Index)
    CharInfo(2).Caption = "Niveau: " & GetPlayerLevel(Index)
    CharInfo(3).Caption = "PV: " & GetPlayerHP(Index) & "/" & GetPlayerMaxHP(Index)
    CharInfo(4).Caption = "PM: " & GetPlayerMP(Index) & "/" & GetPlayerMaxMP(Index)
    CharInfo(5).Caption = "End: " & GetPlayerSP(Index) & "/" & GetPlayerMaxSP(Index)
    CharInfo(6).Caption = "Exp: " & GetPlayerExp(Index) & "/" & GetPlayerNextLevel(Index)
    CharInfo(7).Caption = "Acc�s : " & GetPlayerAccess(Index)
    CharInfo(8).Caption = "PK: " & GetPlayerPK(Index)
    CharInfo(9).Caption = "Classe: " & Classe(GetPlayerClass(Index)).Name
    CharInfo(10).Caption = "Sprite: " & GetPlayerSprite(Index)
    CharInfo(11).Caption = "Sexe: " & STR$(Player(Index).Char(Player(Index).CharNum).Sex)
    CharInfo(12).Caption = "Map: " & GetPlayerMap(Index)
    CharInfo(13).Caption = "Guilde: " & GetPlayerGuild(Index)
    CharInfo(14).Caption = "Guilde Acc�s : " & GetPlayerGuildAccess(Index)
    CharInfo(15).Caption = "For: " & GetPlayerStr(Index)
    CharInfo(16).Caption = "Def: " & GetPlayerDEF(Index)
    CharInfo(17).Caption = "Vitesse: " & GetPlayerSPEED(Index)
    CharInfo(18).Caption = "Magie: " & GetPlayerMAGI(Index)
    CharInfo(19).Caption = "Points: " & GetPlayerPOINTS(Index)
    CharInfo(20).Caption = "Index: " & Index
    picStats.Visible = True
End Sub

Private Sub Command2_Click()
    Dim z As String
    If Not InDestroy Then
    z = MsgBox("Voulez-vous vraiment fermer le serveur?", vbYesNo, "Fermeture du serveur")
    If z = vbYes Then
        Call DestroyServer
    End If
End If
End Sub

Private Sub Command20_Click()
HotelDeVente.AddAchat 1, 1, 1, 1, 1
End Sub

Private Sub Command21_Click()
Num = 5
Command7.Caption = "Envoyer"
Label4.Caption = "Message :"
picReason.Height = 1335
picJail.Visible = False
picReason.Visible = True
End Sub

Private Sub Command22_Click()
Dim Index As Long
Dim i As Integer

'Index = lvUsers.ListItems(lvUsers.SelectedItem.Index).text
For i = 1 To lvUsers.ListItems.Count - 1 Step 1
    If lvUsers.ListItems.item(i).Selected = True Then
        Index = lvUsers.ListItems.item(i).text
        Call PlayerMsg(Index, "Tu es maintenant muet.", White)
        Call TextAdd(frmServer.txtText(0), GetPlayerName(Index) & " est maintenant muet.", True)
        Player(Index).Mute = True
    End If
Next i

End Sub

Private Sub Command23_Click()
Dim Index As Long
Dim i As Integer

'Index = lvUsers.ListItems(lvUsers.SelectedItem.Index).text

For i = 1 To lvUsers.ListItems.Count - 1 Step 1
    If lvUsers.ListItems.item(i).Selected = True Then
        Index = lvUsers.ListItems.item(i).text
        Call PlayerMsg(Index, "Tu peux � nouveau parler.", White)
        Call TextAdd(frmServer.txtText(0), GetPlayerName(Index) & " peut � nouveau parler.", True)
        Player(Index).Mute = False
    End If
Next i

End Sub

Private Sub Command24_Click()
Num = 6
Command7.Caption = "Tuer"
Label4.Caption = "Dire :"
picReason.Height = 1335
picJail.Visible = False
picReason.Visible = True
End Sub

Private Sub Command25_Click()
If Scripting = 1 Then
    Set MyScript = Nothing
    Set clsScriptCommands = Nothing
    
    Set MyScript = New clsSadScript
    Set clsScriptCommands = New clsCommands
    MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
    MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
    Call TextAdd(frmServer.txtText(0), "Scripts recharg�s.", True)
    Call IBMsg("Scripts recharg�s.")
End If
    frmServer.lblScriptAlert.Visible = False
    If Scripting = 1 Then
        frmServer.lblscript.Caption = "Etat : activ�"
        frmServer.lblscript.ForeColor = RGB(51, 204, 0)
    Else
        frmServer.lblscript.Caption = "Etat : d�sactiv�"
        frmServer.lblscript.ForeColor = RGB(255, 102, 102)
    End If
End Sub

Private Sub Command26_Click()

If Scripting = 1 Then
    Scripting = 0
    PutVar App.Path & "\Data.ini", "CONFIG", "Scripting", 0
    
    If Scripting = 0 Then
        Set MyScript = Nothing
        Set clsScriptCommands = Nothing
    End If
End If
Call IBMsg("Les scripts sont d�sactiv�s.")
If Scripting = 1 Then
        frmServer.lblscript.Caption = "Etat : activ�"
        frmServer.lblscript.ForeColor = RGB(51, 204, 0)
    Else
        frmServer.lblscript.Caption = "Etat : d�sactiv�"
        frmServer.lblscript.ForeColor = RGB(255, 102, 102)
    End If
End Sub

Private Sub Command27_Click()
If Scripting = 0 Then
    Scripting = 1
    PutVar App.Path & "\Data.ini", "CONFIG", "Scripting", 1
    
    If Scripting = 1 Then
        Set MyScript = New clsSadScript
        Set clsScriptCommands = New clsCommands
        MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
        MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
    End If
End If
Call IBMsg("Les scripts sont activ�s.")
If Scripting = 1 Then
        frmServer.lblscript.Caption = "Etat : activ�"
        frmServer.lblscript.ForeColor = RGB(51, 204, 0)
    Else
        frmServer.lblscript.Caption = "Etat : d�sactiv�"
        frmServer.lblscript.ForeColor = RGB(255, 102, 102)
    End If
End Sub

Private Sub Command29_Click()
    Call LoadClasses
    Call TextAdd(frmServer.txtText(0), "Classes recharg�es.", True)
    Call IBMsg("Classes recharg�es.")
End Sub

Private Sub Command3_Click()
Num = 7
Command7.Caption = "Soin"
Label4.Caption = "Dire :"
picReason.Height = 1335
picJail.Visible = False
picReason.Visible = True
End Sub

Private Sub Command30_Click()
Dim z As String
Dim O As Long
    z = InputBox("Num�ros de la classe?", "Modifier les classes")
    If Val(z) < 0 Or Val(z) > Max_Classes Or Not IsNumeric(z) Then Exit Sub
    O = Val(z)
    frmclasseseditor.nom.text = ReadINI("CLASS", "Name", App.Path & "\Classes\Class" & O & ".ini")
    frmclasseseditor.scrlhom.value = Val(ReadINI("CLASS", "MaleSprite", App.Path & "\Classes\Class" & O & ".ini"))
    frmclasseseditor.scrlfem.value = Val(ReadINI("CLASS", "FemaleSprite", App.Path & "\Classes\Class" & O & ".ini"))
    frmclasseseditor.numsf.Caption = frmclasseseditor.scrlfem.value
    frmclasseseditor.numsh.Caption = frmclasseseditor.scrlhom.value
    frmclasseseditor.force.text = ReadINI("CLASS", "STR", App.Path & "\Classes\Class" & O & ".ini")
    frmclasseseditor.def.text = ReadINI("CLASS", "DEF", App.Path & "\Classes\Class" & O & ".ini")
    frmclasseseditor.vit.text = ReadINI("CLASS", "SPEED", App.Path & "\Classes\Class" & O & ".ini")
    frmclasseseditor.magi.text = ReadINI("CLASS", "MAGI", App.Path & "\Classes\Class" & O & ".ini")
    frmclasseseditor.carted.text = ReadINI("CLASS", "MAP", App.Path & "\Classes\Class" & O & ".ini")
    frmclasseseditor.xd.text = ReadINI("CLASS", "X", App.Path & "\Classes\Class" & O & ".ini")
    frmclasseseditor.yd.text = ReadINI("CLASS", "Y", App.Path & "\Classes\Class" & O & ".ini")
    frmclasseseditor.arme.text = ReadINI("STARTUP", "Weapon", App.Path & "\Classes\Class" & O & ".ini")
    frmclasseseditor.bouclier.text = ReadINI("STARTUP", "Shield", App.Path & "\Classes\Class" & O & ".ini")
    frmclasseseditor.armure.text = ReadINI("STARTUP", "Armor", App.Path & "\Classes\Class" & O & ".ini")
    frmclasseseditor.caske.text = ReadINI("STARTUP", "Helmet", App.Path & "\Classes\Class" & O & ".ini")
    frmclasseseditor.ajf.text = ReadINI("CLASSCHANGE", "AddStr", App.Path & "\Classes\Class" & O & ".ini")
    frmclasseseditor.ajd.text = ReadINI("CLASSCHANGE", "AddDef", App.Path & "\Classes\Class" & O & ".ini")
    frmclasseseditor.ajv.text = ReadINI("CLASSCHANGE", "AddSpeed", App.Path & "\Classes\Class" & O & ".ini")
    frmclasseseditor.ajm.text = ReadINI("CLASSCHANGE", "AddMagi", App.Path & "\Classes\Class" & O & ".ini")
    frmclasseseditor.cartem.text = ReadINI("DEATH", "Map", App.Path & "\Classes\Class" & O & ".ini")
    frmclasseseditor.xm.text = ReadINI("DEATH", "x", App.Path & "\Classes\Class" & O & ".ini")
    frmclasseseditor.ym.text = ReadINI("DEATH", "y", App.Path & "\Classes\Class" & O & ".ini")
    frmclasseseditor.lock.value = Val(ReadINI("CLASS", "Locked", App.Path & "\Classes\Class" & O & ".ini"))
    frmclasseseditor.Tag = O
    frmclasseseditor.Show
End Sub

Private Sub Command31_Click()
Dim Index As Long

For Index = 1 To MAX_PLAYERS
    If IsPlaying(Index) = True Then
        If GetPlayerAccess(Index) <= 0 Then
            Call SetPlayerHP(Index, 0)
            Call PlayerMsg(Index, "Vous avez �t� tu� par l'admin du serveur.", BrightRed)
            
            ' Warp player away
            If Scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & Index
            Else
                Call PlayerWarp(Index, START_MAP, START_X, START_Y)
            End If
            Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
            Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
            Call SetPlayerSP(Index, GetPlayerMaxSP(Index))
            Call SendHP(Index)
            Call SendMP(Index)
            Call SendSP(Index)
        End If
    End If
Next Index
End Sub

Private Sub Command32_Click()
    scrlMM.Max = MAX_MAPS
    scrlMX.Max = MAX_MAPX
    scrlMY.Max = MAX_MAPY
    picWarp.Visible = True
End Sub

Private Sub Command33_Click()
    picExp.Visible = True
End Sub

Private Sub Command34_Click()
Dim Index As Long
Dim i As Long
    
Call GlobalMsg("L'admin du serveur donne un niveau � tous.", BrightGreen)
    
For Index = 1 To MAX_PLAYERS
    If IsPlaying(Index) = True Then
        If GetPlayerLevel(Index) >= MAX_LEVEL Then
            Call SetPlayerExp(Index, experience(MAX_LEVEL))
            Call SendStats(Index)
        Else
            Call SetPlayerLevel(Index, GetPlayerLevel(Index) + 1)
                                
            i = Int(GetPlayerSPEED(Index) / 10)
            If i < 1 Then i = 1
            If i > 3 Then i = 3
                
            Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + i)
            If GetPlayerLevel(Index) >= MAX_LEVEL Then
                Call SetPlayerExp(Index, experience(MAX_LEVEL))
                Call SendStats(Index)
            End If
            Call SendStats(Index)
        End If
    End If
Next Index


Dim J As Long

    Call RemovePLR
    
    For J = 1 To MAX_PLAYERS
        Call ShowPLR(J)
    Next J
    
End Sub

Private Sub Command35_Click()
Dim i As Long
    MapList.Clear
    Call LoadMaps
    For i = 1 To MAX_MAPS
        MapList.AddItem i & ": " & Map(i).Name
    Next i
    
    frmServer.MapList.Selected(0) = True
End Sub

Private Sub Command36_Click()
Dim Index As Long
Dim i As Long

Index = MapList.ListIndex + 1

    frmInfosMap.MapInfo(0).Caption = "Carte " & Index & " - " & Map(Index).Name
    frmInfosMap.MapInfo(1).Caption = "R�vision: " & Map(Index).Revision
    frmInfosMap.MapInfo(2).Caption = "Morale: " & Map(Index).Moral
    frmInfosMap.MapInfo(3).Caption = "Haut: " & Map(Index).Up
    frmInfosMap.MapInfo(4).Caption = "Bas: " & Map(Index).Down
    frmInfosMap.MapInfo(5).Caption = "Gauche: " & Map(Index).Left
    frmInfosMap.MapInfo(6).Caption = "Droite: " & Map(Index).Right
    frmInfosMap.MapInfo(7).Caption = "Musique: " & Map(Index).Music
    frmInfosMap.MapInfo(8).Caption = "Carte de d�part: " & Map(Index).BootMap
    frmInfosMap.MapInfo(9).Caption = "D�part des X: " & Map(Index).BootX
    frmInfosMap.MapInfo(10).Caption = "D�part des Y: " & Map(Index).BootY
    frmInfosMap.MapInfo(11).Caption = "Magasin: " & Map(Index).Shop
    frmInfosMap.MapInfo(12).Caption = "Int�rieur: " & Map(Index).Indoors
    frmInfosMap.lstNPC.Clear
    For i = 1 To MAX_MAP_NPCS
        frmInfosMap.lstNPC.AddItem i & ": " & Npc(Map(Index).Npc(i)).Name
    Next i
    
    frmInfosMap.Show
End Sub

Private Sub Command37_Click()
Dim i As Long

Call GlobalMsg("T�l�portation � la carte :" & scrlMM.value & " X:" & scrlMX.value & " Y:" & scrlMY.value, Yellow)

For i = 1 To MAX_PLAYERS
    If IsPlaying(i) = True Then
        If GetPlayerAccess(i) <= 1 Then
            Call PlayerWarp(i, scrlMM.value, scrlMX.value, scrlMY.value)
        End If
    End If
Next i
    picWarp.Visible = False
End Sub

Private Sub Command38_Click()
    picWarp.Visible = False
End Sub

Private Sub Command39_Click()
    picExp.Visible = False
End Sub

Private Sub Command4_Click()
    CMessages(CM).Title = txtTitle.text
    CMessages(CM).message = txtMsg.text
    PutVar App.Path & "\CMessages.ini", "MESSAGES", "Title" & CM, CMessages(CM).Title
    PutVar App.Path & "\CMessages.ini", "MESSAGES", "Message" & CM, CMessages(CM).message
    CustomMsg(CM - 1).Caption = CMessages(CM).Title
    picCMsg.Visible = False
End Sub

Private Sub Command40_Click()
Dim Index As Long

If Not IsNumeric(txtExp.text) Then MsgBox "Entrer un chiffre s'il vous plait.": Exit Sub

If txtExp.text >= 0 And txtExp.text <= 1000000 Then
    Call GlobalMsg("L'admin du serveur donne " & txtExp.text & "pts d'exp�rience � tous.", BrightGreen)
    
    For Index = 1 To MAX_PLAYERS
        If IsPlaying(Index) = True Then
            Call SetPlayerExp(Index, GetPlayerExp(Index) + txtExp.text)
            Call CheckPlayerLevelUp(Index)
        End If
    Next Index
Else
    Call MsgBox("Merci de sp�cifier une valeur entre 1 et 1 000 000")
End If

    picExp.Visible = False
    
    
    Dim i As Long

    Call RemovePLR
    
    For i = 1 To MAX_PLAYERS
        Call ShowPLR(i)
    Next i
    
End Sub



Private Sub Command45_Click()
Command10.Caption = "T�l�porter"
picReason.Height = 750
scrlMap.Max = MAX_MAPS
scrlX.Max = MAX_MAPX
scrlY.Max = MAX_MAPY
picReason.Visible = False
picJail.Visible = True
End Sub



Private Sub Command46_Click()
Call SaveAllPlayersOnline
End Sub

Private Sub Command47_Click()
Dim value As Long
value = CInt(Val(txtJournuit.text))
If value < 1 Then
    Call MsgBox("Minimum 1 minute s'il vous plait.")
ElseIf value > 1000000 Then
    Call MsgBox("Maximum 1 000 000 minute s'il vous plait.")
ElseIf txtJournuit.text = vbNullString Then
    tmrJournuit.Enabled = False
Else
    tmrJournuit.Enabled = True
    tempjn = value
    txtJournuit.text = tempjn
End If
End Sub

Private Sub Command48_Click()
Dim value As Long
value = CInt(Val(txtRandom.text))

If value < 1 Then
    Call MsgBox("Minimum 1 minute s'il vous plait.")
ElseIf value > 1000000 Then
    Call MsgBox("Maximum 1 000 000 minute s'il vous plait.")
Else
    tempr = value
    txtRandom.text = tempr
    tmrTemps.Enabled = True
    tmrRandom.Enabled = True
End If
End Sub

Private Sub Command49_Click()
tmrTemps.Enabled = False
tmrRandom.Enabled = False
    GameWeather = WEATHER_NONE
    Call SendWeatherToAll
End Sub

Private Sub Command5_Click()
    picCMsg.Visible = False
End Sub

Private Sub Command50_Click()
tmrJournuit.Enabled = False
End Sub

Private Sub Command51_Click()
If PIC_PL = 1 And PIC_NPC1 = 1 And PIC_NPC2 = 0 Then
    frmServer.petit.value = True
    frmServer.grand.value = False
Else
    frmServer.grand.value = True
    frmServer.petit.value = False
End If

Picskint.Visible = True
End Sub

Private Sub Command52_Click()
Picskint.Visible = False
End Sub

Private Sub Command53_Click()
If petit.value = True Then
PIC_PL = 1
PIC_NPC1 = 1
PIC_NPC2 = 0
Else
PIC_PL = 64
PIC_NPC1 = 2
PIC_NPC2 = 32
End If
Call PutVar(App.Path & "\Data.ini", "MAX", "PIC_PL", STR$(PIC_PL))
Call PutVar(App.Path & "\Data.ini", "MAX", "PIC_NPC1", STR$(PIC_NPC1))
Call PutVar(App.Path & "\Data.ini", "MAX", "PIC_NPC2", STR$(PIC_NPC2))
Picskint.Visible = False
End Sub

Private Sub Command58_Click()
    If GameTime = TIME_DAY Then
        GameTime = TIME_NIGHT
    ElseIf GameTime = TIME_NIGHT Then
        GameTime = TIME_DAY
    End If
    Call SendTimeToAll
End Sub

Private Sub Command59_Click()
    Dim MeteoParMap As Byte
    
    picWeather.Visible = True
    
    MeteoParMap = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "MeteoParMap"))
    
    If MeteoParMap = 0 Then
        frmServer.Label12.Caption = "La m�t�o par map est actuellement d�sactiv�e"
        frmServer.CommandMeteoMap.Caption = "activer la m�t�o par map"
        Select Case GameWeather
            Case 0 'Soleil
                frmServer.Label5.Caption = "Temps actuel : Soleil"
            Case 1 'Pluie
                frmServer.Label5.Caption = "Temps actuel : Pluie"
            Case 2 'Neige
                frmServer.Label5.Caption = "Temps actuel : Neige"
            Case 3 'Orage
                frmServer.Label5.Caption = "Temps actuel : Orage"
        End Select
    Else
        frmServer.Label12.Caption = "La m�t�o par map est actuellement activ�e"
        frmServer.CommandMeteoMap.Caption = "d�sactiver la m�t�o par map"
        frmServer.Label5.Caption = "Temps actuel : M�t�o par map"
    End If
    
End Sub

Private Sub Command6_Click()
picReason.Visible = False
End Sub

Private Sub Command60_Click()
    Call SaveLogs
    mesgSave.Caption = "Sauvegard�"
    mesgSaveTimer.Enabled = True
End Sub

Private Sub Command61_Click()
    picWeather.Visible = False
End Sub

Private Sub Command62_Click()
    GameWeather = WEATHER_NONE
    Call SendWeatherToAll
End Sub

Private Sub Command63_Click()
    GameWeather = WEATHER_THUNDER
    Call SendWeatherToAll
End Sub

Private Sub Command64_Click()
    GameWeather = WEATHER_RAINING
    Call SendWeatherToAll
End Sub

Private Sub Command65_Click()
    GameWeather = WEATHER_SNOWING
    Call SendWeatherToAll
End Sub

Private Sub Command66_Click()
Dim i As Long

    Call RemovePLR
    
    For i = 1 To MAX_PLAYERS
        Call ShowPLR(i)
    Next i
End Sub

Private Sub Command7_Click()
Dim Index As Long
Dim i As Integer

If txtReason.text = vbNullString Then
    txtReason.text = "Aucune raison sp�cifi�e"
End If

'Index = lvUsers.ListItems(lvUsers.SelectedItem.Index).text
For i = 1 To lvUsers.ListItems.Count - 1 Step 1
    If lvUsers.ListItems.item(i).Selected = True Then
        Index = lvUsers.ListItems.item(i).text

        If Num = 1 Then
            If Index > 0 Then
                If IsPlaying(Index) Then Call GlobalMsg(GetPlayerName(Index) & " a �t� d�connect� par l'admin du serveur. Raison:(" & txtReason.text & ")", White)
            Call AlertMsg(Index, "Tu as �t� d�connect� par l'admin du serveur. Raison:(" & txtReason.text & ")")
            End If
        ElseIf Num = 2 Then
            Call BanByServer(Index, txtReason.text)
        ElseIf Num = 5 Then
            Call PlayerMsg(Index, "Message priv� de l'admin du serveur: -- " & Trim$(txtReason.text), BrightGreen)
        ElseIf Num = 6 Then
            Call SetPlayerHP(Index, 0)
            Call PlayerMsg(Index, txtReason.text, BrightRed)
    
            ' Warp player away
            If Scripting = 1 Then MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & Index Else Call PlayerWarp(Index, START_MAP, START_X, START_Y)
    
            Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
            Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
            Call SetPlayerSP(Index, GetPlayerMaxSP(Index))
            Call SendHP(Index)
            Call SendMP(Index)
            Call SendSP(Index)
        ElseIf Num = 7 Then
            Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
            Call SendHP(Index)
            Call PlayerMsg(Index, txtReason.text, BrightGreen)
        End If
    End If
Next i

picReason.Visible = False
End Sub

Private Sub Command8_Click()
    picStats.Visible = False
End Sub

Private Sub Command9_Click()
Dim Index As Long

For Index = 1 To MAX_PLAYERS
    If IsPlaying(Index) = True Then
        If GetPlayerAccess(Index) <= 0 Then
            Call GlobalMsg(GetPlayerName(Index) & " a �t� d�connect� du serveur.", White)
            Call AlertMsg(Index, "Vous avez �t� d�connect� du serveur.")
        End If
    End If
Next Index

Dim i As Long

    Call RemovePLR
    
    For i = 1 To MAX_PLAYERS
        Call ShowPLR(i)
    Next i
    
End Sub

Private Sub CommandMeteoMap_Click()
    
    Dim MeteoParMap As Byte
      
    MeteoParMap = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "MeteoParMap"))
    
    If MeteoParMap = 0 Then
        PutVar App.Path & "\Data.ini", "CONFIG", "MeteoParMap", 1
        frmServer.Label12.Caption = "La m�t�o par map est actuellement activ�e"
        frmServer.Label5.Caption = "Temps actuel : M�t�o par map"
        frmServer.CommandMeteoMap.Caption = "D�sactiver la m�t�o par map"
    Else
        PutVar App.Path & "\Data.ini", "CONFIG", "MeteoParMap", 0
        frmServer.Label12.Caption = "La m�t�o par map est actuellement d�sactiv�e"
        frmServer.CommandMeteoMap.Caption = "Activer la m�t�o par map"
        frmServer.Label5.Caption = "Temps actuel : Soleil"
    End If
End Sub

Private Sub CustomMsg_Click(Index As Integer)
    CM = Index + 1
    txtTitle.text = CMessages(CM).Title
    txtMsg.text = CMessages(CM).message
    picCMsg.Visible = True
End Sub

Private Sub Form_Load()
Random = 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lmsg As Long
    lmsg = X
    Select Case lmsg
        Case &H203
            frmServer.WindowState = vbNormal
            frmServer.Show
        Case &H204
            SetForegroundWindow Me.hwnd
            Call PopupMenu(Dmenu)
    End Select
End Sub

Private Sub Form_Resize()
    'If frmServer.WindowState = vbMinimized Then frmServer.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim z As String
If Not InDestroy Then
    z = MsgBox("Voulez-vous vraiment fermer le serveur?", vbYesNo, "Fermeture du Serveur")
    If z = vbYes Then
        Call DestroyServer
    Else
        Cancel = True
    End If
End If
End Sub



Private Sub ftpmaps_opt_Click()
frmOptFTP.Show vbModeless, frmServer
End Sub

Private Sub HScroll1_Change()
Label7.Caption = "Ratio d'exp�rience normal : " & HScroll1.value
Command13.Enabled = True
End Sub

Private Sub HScroll2_Change()
Label13.Caption = "Ratio d'exp�rience qu�te : " & HScroll2.value
Command13.Enabled = True
End Sub

Private Sub infbull_opt_Click()
Call ChargIBOpt
frmOptInfoBulle.Show vbModeless, frmServer
End Sub

Private Sub jn_Click()
Call Command58_Click
End Sub

Private Sub Label7_Click()
    Shell ("explorer http://monip.org"), vbNormalNoFocus
End Sub

Private Sub lstTopics_Click()
Dim FileName As String
Dim hFile As Long

    txtTopic.text = vbNullString
    
    TopicTitle.Caption = lstTopics.List(lstTopics.ListIndex)
    FileName = lstTopics.ListIndex + 1 & ".txt"
        
    If FileExist("Guides\" & FileName) = True And FileName <> vbNullString Then
        hFile = FreeFile
        Open App.Path & "\Guides\" & FileName For Input As #hFile
            txtTopic.text = Input$(LOF(hFile), hFile)
        Close #hFile
    End If
End Sub

Private Sub mesgSaveTimer_Timer()
mesgSave.Caption = ""
mesgSaveTimer.Enabled = False

End Sub

Private Sub mhelpfaq_Click()
ShellExecute Me.hwnd, "open", "http://www.frogcreator.fr/frogaide", "", App.Path, 1
End Sub

Private Sub mnuServerLog_Click()
    If mnuServerLog.value = Checked Then ServerLog = False Else ServerLog = True
End Sub

Private Sub optcoul_Click()

End Sub

Private Sub optftp_Click()

End Sub

Private Sub optib_Click()

End Sub

Private Sub mwebsite_Click()
ShellExecute Me.hwnd, "open", "http://www.frogcreator.fr", "", App.Path, 1
End Sub

Private Sub ouvr_Click()
frmServer.WindowState = vbNormal
frmServer.Show
End Sub

Private Sub PlayerTimer_Timer()
Dim i As Long

If PlayerI <= MAX_PLAYERS Then
    If IsPlaying(PlayerI) Then
        Call SavePlayer(PlayerI)
        Call PlayerMsg(PlayerI, GetPlayerName(PlayerI) & " est maintenant enregistr�.", Yellow)
    End If
    PlayerI = PlayerI + 1
End If
If PlayerI >= MAX_PLAYERS Then
    PlayerI = 1
    PlayerTimer.Enabled = False
    tmrPlayerSave.Enabled = True
End If

CClasses = True
End Sub

Private Sub quit_Click()
Call DestroyServer
End Sub

Private Sub rchrgcls_Click()
Call Command29_Click
End Sub

Private Sub rechrgscr_Click()
Call Command25_Click
End Sub

Private Sub Say_Click(Index As Integer)
    Call GlobalMsg(Trim$(CMessages(Index + 1).message), White)
    Call TextAdd(frmServer.txtText(0), "Msg rapide : " & Trim$(CMessages(Index + 1).message), True)
End Sub

Private Sub scrlMap_Change()
    txtMap.Caption = "Carte : " & scrlMap.value
End Sub

Private Sub scrlMM_Change()
    lblMM.Caption = "Carte : " & scrlMM.value
End Sub

Private Sub scrlMX_Change()
    lblMX.Caption = "X: " & scrlMX.value
End Sub

Private Sub scrlMY_Change()
    lblMY.Caption = "Y: " & scrlMY.value
End Sub

Private Sub scrlX_Change()
    txtX.Caption = "X: " & scrlX.value
End Sub

Private Sub scrlY_Change()
    txtY.Caption = "Y: " & scrlY.value
End Sub

Private Sub sync_Timer()
Dim i As Long
For i = 1 To MAX_PLAYERS
If IsConnected(i) And Player(i).sync = False Then
Player(i).sync = False
'Call CloseSocket(i)
End If
Player(i).sync = False
Next i
End Sub

Private Sub tmrChatLogs_Timer()
Static ChatSecs As Long
Dim SaveTime As Long

SaveTime = 3600

    If frmServer.chkChat.value = Unchecked Then
        ChatSecs = SaveTime
        Label6.Caption = "Les logs sont d�sactiv�s."
        Exit Sub
    End If
    
    If ChatSecs <= 0 Then ChatSecs = SaveTime
    If ChatSecs > 60 Then
        Label6.Caption = "Enregistrement des logs dans " & Int(ChatSecs / 60) & " Minute(s)"
    Else
        Label6.Caption = "Enregistrement des logs dans " & Int(ChatSecs) & " Seconde(s)"
    End If
    
    ChatSecs = ChatSecs - 1
    
    If ChatSecs <= 0 Then
        Call TextAdd(txtText(0), "Les logs ont �t� enregistr�s.", True)
        Call SaveLogs
        ChatSecs = 0
    End If
End Sub

Private Sub tmrGameAI_Timer()
    Call ServerLogic
End Sub

Private Sub tmrJournuit_Timer()
    minuteJN = minuteJN + 1
If tempjn = minuteJN Then
    If GameTime = TIME_DAY Then
        GameTime = TIME_NIGHT
    ElseIf GameTime = TIME_NIGHT Then
        GameTime = TIME_DAY
    End If
    Call SendTimeToAll
    minuteJN = 0
End If
End Sub

Private Sub tmrPlayerSave_Timer()
    Call PlayerSaveTimer
End Sub

Private Sub tmrRandom_Timer()
If Random < 4 Then
Random = Random + 1
ElseIf Random >= 4 Then
Random = 1
End If
End Sub

Private Sub tmrSpawnMapItems_Timer()
    Call CheckSpawnMapItems
End Sub

Private Sub tmrTemps_Timer()
minuteR = minuteR + 1
If minuteR = tempr Then
If Random = 1 Then
    GameWeather = WEATHER_NONE
    Call SendWeatherToAll
ElseIf Random = 2 Then
    GameWeather = WEATHER_RAINING
    Call SendWeatherToAll
ElseIf Random = 3 Then
    GameWeather = WEATHER_THUNDER
    Call SendWeatherToAll
ElseIf Random = 4 Then
    GameWeather = WEATHER_SNOWING
    Call SendWeatherToAll
End If
End If
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim$(txtChat.text) <> vbNullString Then
        Call GlobalMsg(txtChat.text, White)
        Call TextAdd(frmServer.txtText(0), "Serveur: " & txtChat.text, True)
        txtChat.text = vbNullString
    End If
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub tmrShutdown_Timer()
Static Secs As Long

    If Secs <= 0 Then Secs = 30
    ShutdownTime.Caption = "Fermeture: " & Secs & " Secondes"
    If Secs = 30 Then Call TextAdd(frmServer.txtText(0), "Femeture automatique dans " & Secs & " secondes.", True)
    If Secs = 30 Then Call GlobalMsg("Fermeture dans " & Secs & " secondes.", BrightBlue)
    If Secs = 25 Then Call GlobalMsg("Fermeture dans " & Secs & " secondes.", BrightBlue)
    If Secs = 20 Then Call GlobalMsg("Fermeture dans " & Secs & " secondes.", BrightBlue)
    If Secs = 15 Then Call GlobalMsg("Fermeture dans " & Secs & " secondes.", BrightBlue)
    If Secs = 10 Then Call GlobalMsg("Fermeture dans " & Secs & " secondes.", BrightBlue)
    If Secs < 6 Then Call GlobalMsg("Fermeture dans " & Secs & " secondes.", BrightBlue)
    
    Secs = Secs - 1
    If Secs <= 0 Then tmrShutdown.Enabled = False: Call DestroyServer
End Sub

Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Call AcceptConnection(Index, requestID)
End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    If IsConnected(Index) Then Call IncomingData(Index, bytesTotal)
End Sub

Private Sub Socket_Close(Index As Integer)
    Call CloseSocket(Index)
End Sub

Private Sub txtText_GotFocus(Index As Integer)
    txtChat.SetFocus
End Sub
