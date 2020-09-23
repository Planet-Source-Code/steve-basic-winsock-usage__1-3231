VERSION 5.00
Begin VB.Form frmSendKey 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sendkeys to their Computer (Cause things to be typed for them :))"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSendKey.frx":0000
   ScaleHeight     =   855
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Simply type into the textbox.  Your text will not appear, do not worry:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmSendKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
m = Chr(KeyAscii)
frmMain.win.SendData "snk " & m
If Err Then MsgBox Err.Number & Chr(10) & Chr(10) & Err.Description, vbOKOnly + vbCritical, "[ERROR]"
End Sub
