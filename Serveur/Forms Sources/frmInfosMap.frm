VERSION 5.00
Begin VB.Form frmInfosMap 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informations sur la map"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picMap 
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
      Height          =   3375
      Left            =   0
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   223
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton Command41 
         Caption         =   "Fermer"
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         Top             =   3000
         Width           =   1575
      End
      Begin VB.ListBox lstNPC 
         Height          =   2205
         Left            =   1680
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Map"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   300
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Révision:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   660
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Morale:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   540
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Haut:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   405
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bas:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   315
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gauche:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Droite:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Musique:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   645
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Map de départ:"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   1110
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Départ des X:"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   990
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Départ des Y:"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   6
         Top             =   2520
         Width           =   990
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Magasin:"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   645
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Intérieur:"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   4
         Top             =   3000
         Width           =   690
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PNJ :"
         Height          =   195
         Index           =   13
         Left            =   1680
         TabIndex        =   3
         Top             =   285
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmInfosMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
