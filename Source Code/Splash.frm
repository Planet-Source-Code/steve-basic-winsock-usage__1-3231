VERSION 5.00
Begin VB.Form Splash 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   Icon            =   "Splash.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Splash.frx":030A
   ScaleHeight     =   3000
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Load frmMain
Load frmConnect
Load frmOptions
Load frmWeb
frmMain.Visible = True
Unload Me
End Sub

