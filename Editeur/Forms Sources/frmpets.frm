VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmpet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crédits"
   ClientHeight    =   3960
   ClientLeft      =   165
   ClientTop       =   270
   ClientWidth     =   4905
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4905
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtpet 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5530
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      Enabled         =   0   'False
      Appearance      =   0
      TextRTF         =   $"frmpets.frx":0000
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Revenir au menu"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   3360
      Width           =   1695
   End
End
Attribute VB_Name = "frmpet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_Click()
If frmMirage.Visible Then
    frmpet.Visible = False
    frmMirage.SetFocus
Else
    frmMainMenu.Visible = True
    frmpet.Visible = False
End If
End Sub

Private Sub creditline1_Click()

End Sub

Private Sub Form_Load()
rtpet.Text = "Remerciements : Coke, GodSentdeath, Katsuo, Edouard, Dahevos et à toute la communauté de FRoG Creator." & vbCrLf & vbCrLf & "Merci à Hinomi pour sa belle bannière et à Rose pour la partie graphique de FRoG Creator" & vbCrLf & vbCrLf & "Programmation : Matsura, Rydan, GAK, Koolgraph, hugo-57, Alexoune001, Mimus, Alves57600, Mywaystar, Sarcadent, Lepetitdébutantn°2, Elios." & vbCrLf & "Petite et grande aide à la programmation : Eusebe et lbalpha." & vbCrLf & "Merci à tous les autres si on en oublie."
End Sub

