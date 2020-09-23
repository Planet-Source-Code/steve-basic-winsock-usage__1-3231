VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmClient 
   Caption         =   "Client"
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wsTCP 
      Index           =   0
      Left            =   120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   600
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   315
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtFile 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   1020
      Width           =   4695
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'Written by Igor Ostrovsky (igor@ostrosoft.com)
'Visual Basic 911 (http://www.ostrosoft.com/vb)
'==============================================
Option Explicit

Dim buffer() As Byte
Dim lBytes As Long

Private Sub cmdBrowse_Click()
  dlg.ShowOpen
  txtFile = dlg.filename
End Sub

Private Sub cmdSend_Click()
  cmdSend.Enabled = False
  lBytes = 0
  ReDim buffer(FileLen(dlg.filename) - 1)
  Open dlg.filename For Binary As 1
  Get #1, 1, buffer
  Close #1
  Load wsTCP(1)
  wsTCP(1).RemoteHost = "localhost"
  wsTCP(1).RemotePort = 1111
  wsTCP(1).Connect
  lblStatus = "Connecting..."
End Sub

Private Sub wsTCP_Close(Index As Integer)
  lblStatus = "Connection closed"
  Unload wsTCP(1)
End Sub

Private Sub wsTCP_Connect(Index As Integer)
  lblStatus = "Connected"
  wsTCP(1).SendData buffer
End Sub

Private Sub wsTCP_SendComplete(Index As Integer)
  lblStatus = "Send complete"
  Unload wsTCP(1)
  cmdSend.Enabled = True
End Sub

Private Sub wsTCP_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
  lBytes = lBytes + bytesSent
  lblStatus = lBytes & " out of " & UBound(buffer) & " bytes sent"
End Sub
