VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About PhrØstByte"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5985
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   2985
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   1320
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   240
      ScaleHeight     =   1395
      ScaleWidth      =   5475
      TabIndex        =   0
      Top             =   1440
      Width           =   5475
      Begin VB.PictureBox P1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2805
         Left            =   0
         Picture         =   "frmAbout.frx":6B1F
         ScaleHeight     =   2805
         ScaleWidth      =   5535
         TabIndex        =   1
         Top             =   720
         Width           =   5535
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   975
      Left            =   240
      MouseIcon       =   "frmAbout.frx":DE54
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "frmABout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim thetop As Long
Dim p1hgt As Long
Dim p1wid As Long
Dim theleft As Long
Dim Tempstring As String
Dim pth As String

Private Sub Command1_Click()
'ShowCursor False
  Dim MyValue As Byte, x
  Randomize
Status = "Stars"
  

  
    Do While ActionMode = "Running"
      Background = LoadPicture ' clear the background
      For i = 0 To StarCount
       
        SetPixel Background.hdc, Star(i).StarX, Star(i).StarY + Star(i).SpeedY, QBColor(Star(i).StarColor)
      Next i
      Background.Refresh
      DoEvents
    Loop
    DoEvents
  

 

Status = "Stars"
  StarSetup Picture2.Height, Picture2.Width
  

End Sub

Private Sub Form_Activate()

StarSetup P1.Height, P1.Width
  

End Sub

Private Sub Form_Click()
Unload Me
End Sub

Sub Form_Load()
pth = App.Path & "/Graphics/"
        P1.AutoRedraw = True
        P1.Visible = False
        P1.FontSize = 12
        P1.ForeColor = &H0&
        P1.BackColor = Picture2.BackColor
        P1.ScaleMode = 3
        Picture2.ScaleMode = 3
        Open (App.Path & "\about.txt") For Input As #1
        Line Input #1, Tempstring
        P1.Height = (Val(Tempstring) * P1.TextHeight("Test Height")) + 200
        Do Until EOF(1)
            Line Input #1, Tempstring
            PrintText Tempstring
        Loop
        Close #1
        theleft = 0
        thetop = Picture2.ScaleHeight
        p1hgt = P1.ScaleHeight
        p1wid = P1.ScaleWidth
        Timer1.Enabled = True
        Timer1.Interval = 10
End Sub


Private Sub PicBack_Click()

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FormDrag Me
End Sub

Private Sub Label1_Click()
frmWeb.Show
frmWeb.WebBrowser1.Navigate "http://www.vbtutor.com"

End Sub

Sub Timer1_Timer()
On Error Resume Next
Dim x%, txt$
       x% = BitBlt(Picture2.hdc, theleft, thetop, p1wid, p1hgt, P1.hdc, 0, 0, &HCC0020)
        thetop = thetop - 1
        If thetop < -p1hgt Then
        Timer1.Enabled = False
        txt$ = "Credits Completed"
        CurrentY = ScaleHeight / 2
        CurrentX = (ScaleWidth - TextWidth(txt$)) / 2
        Print txt$
        End If
If Err Then MsgBox Err.Description, vbOKOnly + vbCritical, "Error #" & Err.Number
End Sub

Sub PrintText(Text As String)
Dim x, y, i

P1.ForeColor = &H0&: x = P1.CurrentX: y = P1.CurrentY
For i = 1 To 3
    P1.Print Text
    x = x + 1: y = y + 1: P1.CurrentX = x: P1.CurrentY = y
Next i
P1.ForeColor = &HFF&


P1.Print Text
End Sub

