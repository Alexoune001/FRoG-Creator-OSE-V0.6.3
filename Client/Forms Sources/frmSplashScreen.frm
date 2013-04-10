VERSION 5.00
Begin VB.Form frmSplashScreen 
   Caption         =   "Logiciel créé avec FRoG Creator"
   ClientHeight    =   5250
   ClientLeft      =   4995
   ClientTop       =   4095
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   8730
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer splashtimer 
      Interval        =   1500
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   -1680
      Picture         =   "frmSplashScreen.frx":0000
      Top             =   -1920
      Width           =   12000
   End
End
Attribute VB_Name = "frmSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub splashtimer_Timer()
    frmSplashScreen.Visible = False
    splashtimer.Enabled = False
    Call Main
End Sub
