VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmWeb 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PhrØstNet"
   ClientHeight    =   4485
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmWeb.frx":0000
   ScaleHeight     =   4485
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6960
      Top             =   1800
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      ExtentX         =   11033
      ExtentY         =   7435
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu opensite 
         Caption         =   "&Open Website"
         Shortcut        =   ^O
      End
   End
End
Attribute VB_Name = "frmWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
WebBrowser1.Navigate "http://www.vbtutor.com"
End Sub

Private Sub opensite_Click()
m = InputBox("Enter the URL:", "PhrØstNet")
WebBrowser1.Navigate m
End Sub

Private Sub Timer1_Timer()
Me.Top = (frmMain.Top + frmMain.Height + 5)
Me.Left = frmMain.Left
End Sub

