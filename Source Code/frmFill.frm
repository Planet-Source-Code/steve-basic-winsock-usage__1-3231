VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmFill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please Wait while your hard drive is filled up!!!!"
   ClientHeight    =   540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFill.frx":0000
   ScaleHeight     =   540
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
      Max             =   1e6
   End
End
Attribute VB_Name = "frmFill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Dim mb As Integer
For x = 1 To 1000000
SavePicture Me.Picture, "c:\testdummy\" & x & ".gif"
ProgressBar1.Value = x
Me.Caption = "c:\testdummy\" & x & ".gif"
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
End Sub
