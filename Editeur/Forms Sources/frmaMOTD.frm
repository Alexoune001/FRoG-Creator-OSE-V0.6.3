VERSION 5.00
Begin VB.Form frmaMOTD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editer le MOTD"
   ClientHeight    =   4335
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   3885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton enreg 
      Caption         =   "Enregistrer"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   3615
   End
   Begin VB.TextBox motd 
      Height          =   3735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmaMOTD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub enreg_Click()
Call WriteINI("INFO", "motd", motd.Text, App.Path & "\config.ini")
Call SendMOTDChange(motd.Text)
Unload Me
End Sub

Private Sub Form_Load()
motd.Text = "Bienvenue dans la version " & App.Major & "." & App.Minor & "." & App.Revision & " de FRoG Creator, si vous rencontrez un problème ou un bug veuillez le rapporter sur frogcreator.fr"
End Sub

