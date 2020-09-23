Attribute VB_Name = "ModMain"
Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Const WarpStarSpeed = 100
Type PointAPI
x As Long
y As Long
End Type
Type Stars
    SpeedY As Integer
    SpeedX As Integer
    StarX As Integer
    StarY As Integer
    StarColor As Byte
End Type
Public SpecialEffectX As Integer
Public SpecialEffectY As Integer
Public Star() As Stars 'Array of Stars Type
Public StarCount As Integer ' holds the amount of stars in array
Public Status As String 'holds Name of the current effect



Private Function ConvertIPAddressToLong(strAddress As String) As Long


    'For Ping: It changes the IP Address so it can be used to send th
    '     e ping
    On Error Resume Next
    Dim strTemp As String
    Dim lAddress As Long
    Dim iValCount As Integer
    Dim lDotValues(1 To 4) As String
    strTemp = strAddress
    iValCount = 0
    While InStr(strTemp, ".") > 0
        iValCount = iValCount + 1
        lDotValues(iValCount) = Mid(strTemp, 1, InStr(strTemp, ".") - 1)
        strTemp = Mid(strTemp, InStr(strTemp, ".") + 1)
    Wend


    iValCount = iValCount + 1
    lDotValues(iValCount) = strTemp
    If iValCount <> 4 Then
        ConvertIPAddressToLong = 0
        Exit Function
    End If


    lAddress = Val("&H" & Right("00" & Hex(lDotValues(4)), 2) & _
    Right("00" & Hex(lDotValues(3)), 2) & _
    Right("00" & Hex(lDotValues(2)), 2) & _
    Right("00" & Hex(lDotValues(1)), 2))
    ConvertIPAddressToLong = lAddress
End Function

Public Sub FormDrag(TheForm As Form)


    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub

Sub ctr(ctr As PictureBox, frm As Form)
ctr.Left = (frm.ScaleWidth - ctr.Width) / 2
End Sub

Public Sub SaveListBox(TheList As ListBox, Directory As String)


    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&


    Close #1
End Sub
Public Function FileExists(strPath As String) As Integer


    FileExists = Not (Dir(strPath) = "")
End Function


'Example: Call LoadListBox(list1, "C:\Temp\MyList.dat")
Public Sub LoadListBox(TheList As ListBox, Directory As String)


    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
            TheList.AddItem MyString$
        Wend


        Close #1
        
    End Sub

 

Sub ReDimStars(HowManyStars As Integer)
'call this to reset the amount of stars, MAX = 32,767
  StarCount = HowManyStars
  ReDim Star(0 To HowManyStars)
End Sub


Sub AddStars(NumberToAdd As Integer, WhatHeight As Integer, WhatWidth As Integer)
'call this to add more stars, MAX = 32,767
  Dim NewAmount As Integer, Starloop As Integer
  NewAmount = StarCount + NumberToAdd
  ReDim Preserve Star(0 To NewAmount)
  Select Case Status
    Case "Snow"
      For Starloop = StarCount To NewAmount
        Star(Starloop).StarX = 0
        Star(Starloop).StarX = Int(Rnd * WhatWidth)
        Star(Starloop).StarColor = 15
        Star(Starloop).SpeedY = Int(Rnd * 3) + 1
      Next Starloop
      StarCount = NewAmount
  End Select
End Sub
Sub StarSetup(WhatHeight As Integer, WhatWidth As Integer)
  Dim i As Integer, j As Integer
  If StarCount = Null Or StarCount = 0 Then Exit Sub
  Select Case Status
        
  Case "Snow"
    For i = 0 To StarCount
      Star(i).StarColor = 15
      Star(i).StarX = Int(Rnd * WhatWidth)
      Star(i).StarY = Int(Rnd * WhatHeight)
      Star(i).SpeedY = Int(Rnd * 3) + 1
    Next i
  Case "Stars"
    For i = 0 To StarCount
      Star(i).StarColor = Int(Rnd * 15) + 1
      Star(i).StarX = Int(Rnd * WhatWidth)
      Star(i).StarY = Int(Rnd * WhatHeight)
      Star(i).SpeedY = Int(Rnd * 7) + 1
    Next i
            
  Case "Black Hole"
    For i = 0 To StarCount
      Star(i).StarColor = Int(Rnd * 15) + 1
      Star(i).StarX = Int(WhatWidth / 2)
      Star(i).StarY = Int(WhatHeight / 2)
      Star(i).SpeedY = Int(Rnd * WarpStarSpeed) - (WarpStarSpeed / 2)
      Star(i).SpeedX = Int(Rnd * WarpStarSpeed) - (WarpStarSpeed / 2)
      Do While Star(i).SpeedX = 0 Or Star(i).SpeedY = 0
        Randomize
        Star(i).SpeedY = Int(Rnd * WarpStarSpeed) - (WarpStarSpeed / 2)
        Star(i).SpeedX = Int(Rnd * WarpStarSpeed) - (WarpStarSpeed / 2)
      Loop
      For j = 0 To 30
        NextStarPosition i, WhatWidth, WhatHeight
      Next j
      Next i
  End Select
End Sub


Sub NextStarPosition(StarNumber As Integer, WhatHeight As Integer, WhatWidth As Integer)

  Select Case Status

    Case "Snow"
      Star(StarNumber).StarY = Star(StarNumber).StarY + Star(StarNumber).SpeedY
      Star(StarNumber).StarX = Star(StarNumber).StarX + Int(5 * Rnd) - 2
      If Star(StarNumber).StarX > WhatWidth Then Star(StarNumber).StarX = 0
      If Star(StarNumber).StarX < 0 Then Star(StarNumber).StarX = WhatWidth
      If Star(StarNumber).StarY > WhatHeight Then
        Star(StarNumber).SpeedY = Int(2 * Rnd) + 1
        Star(StarNumber).StarY = Star(StarNumber).SpeedY
        Star(StarNumber).StarColor = 15
      End If
    
    Case "Stars"
      Star(StarNumber).StarY = Star(StarNumber).StarY + Star(StarNumber).SpeedY
      If Star(StarNumber).StarY > WhatHeight Then
        Star(StarNumber).SpeedY = Int(7 * Rnd) + 2
        Star(StarNumber).StarY = Star(StarNumber).SpeedY
        Star(StarNumber).StarColor = Int(Rnd * 15) + 1
      End If

    Case "Black Hole"
      If Star(StarNumber).StarY > WhatHeight Or Star(StarNumber).StarX > WhatWidth Or Star(StarNumber).StarY < 0 Or Star(StarNumber).StarX < 0 Then
        Star(StarNumber).StarX = SpecialEffectX 'Int(WhatWidth / 2) + SpecialEffectX
        Star(StarNumber).StarY = SpecialEffectY 'Int(WhatHeight / 2) + SpecialEffectY
        Randomize
        Star(StarNumber).SpeedX = Int(Rnd * WarpStarSpeed) - (WarpStarSpeed / 2)
        Star(StarNumber).SpeedY = Int(Rnd * WarpStarSpeed) - (WarpStarSpeed / 2)
        Do While (Star(StarNumber).SpeedX = Star(StarNumber).SpeedY Or Star(StarNumber).SpeedX = 0 Or Star(StarNumber).SpeedY = 0)
          Randomize
          Star(StarNumber).SpeedX = Int(Rnd * WarpStarSpeed) - (WarpStarSpeed / 2)
          Star(StarNumber).SpeedY = Int(Rnd * WarpStarSpeed) - (WarpStarSpeed / 2)
        Loop
      End If
      
      Star(StarNumber).StarY = Star(StarNumber).StarY + (Star(StarNumber).SpeedY)
      Star(StarNumber).StarX = Star(StarNumber).StarX + (Star(StarNumber).SpeedX)
    End Select
    
    
End Sub

Sub SendFile(file As String)

End Sub

Sub Play_Avi(VideoClipDir$) 'the VideoClipDir$ holds the VideoClips Dir data
Dim lret

    lret = mciSendString("play VideoClipDir$", 0&, 0, 0) 'This plays the File
End Sub

Sub PlayWav(Wav As String)
Dim mc

mc = sndPlaySound(Wav, 1)
End Sub
