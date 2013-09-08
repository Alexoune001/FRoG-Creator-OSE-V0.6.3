VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A propos du logiciel"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblthank2 
      BackStyle       =   0  'Transparent
      Caption         =   "Merci2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   7095
   End
   Begin VB.Label lblthank 
      BackStyle       =   0  'Transparent
      Caption         =   "Merci."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   4815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 0.6.3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FRoG Creator OSE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image imgLogo 
      Height          =   1965
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      Top             =   120
      Width           =   1965
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "FRoGGy Buggy Muggy Puggy Duggy Ducky "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xm As Integer
Dim ym As Integer

Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Form_Load()
lblthank.Caption = "Remerciements : Coke, GodSentdeath, Katsuo, Edouard, Dahevos et à toute la communauté de FRoG Creator." & vbCrLf & vbCrLf & "Merci à Hinomi pour sa belle bannière et à Rose pour la partie graphique de FRoG Creator"
lblthank2.Caption = "Programmation : Matsura, Rydan, GAK, Koolgraph, hugo-57, Alexoune001, Mimus, Alves57600, Mywaystar, Sarcadent, Elios." & vbCrLf & "Petite et grande aide à la programmation : Eusebe et lbalpha." & vbCrLf & "Merci à tous les autres si on en oublie."
End Sub

Private Sub imgLogo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  xm = X
  ym = Y
End Sub
  
Private Sub imgLogo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    imgLogo.Left = imgLogo.Left + X - xm
    imgLogo.Top = imgLogo.Top + Y - ym
  End If
End Sub

