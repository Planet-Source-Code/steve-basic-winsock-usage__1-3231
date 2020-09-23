VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Connect"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmConnect.frx":0000
   ScaleHeight     =   2985
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Add Remote IP >>>>"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Connect"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      ForeColor       =   &H0000FFFF&
      Height          =   1395
      Left            =   2520
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Recent Connections:"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remote IP:"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmWait.Show
On Error Resume Next
frmMain.win.RemoteHost = List1.Text
frmMain.win.SendData "Connect"
frmMain.rt.Text = ""
Unload Me
If Err Then MsgBox Err.Number & Chr(10) & Chr(10) & Err.Description, vbOKOnly + vbCritical, "[ERROR]"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
List1.AddItem Text1.Text
List1.ListIndex = List1.ListCount - 1
Text1.Text = ""
End Sub

Private Sub Form_Activate()
Text1.Text = frmMain.win.LocalIP
m = FileExists("c:\ppllist.lst")
If m = True Then
LoadListBox List1, "c:\ppllist.lst"
Else
SaveListBox List1, "c:\ppllist.lst"
LoadListBox List1, "c:\ppllist.lst"
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveListBox List1, "c:\ppllist.lst"
End Sub
