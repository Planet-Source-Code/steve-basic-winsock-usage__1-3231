Attribute VB_Name = "srvMod"
Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Declare Function MakeSureDirectoryPathExists Lib "IMAGEHLP.DLL" (ByVal DirPath As String) As Long

#If Win16 Then


Declare Function SendMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
#Else


Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#End If
Public Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long


Declare Function SetCursorPos Lib "User32" (ByVal x As Long, ByVal y As Long) As Long

Private Declare Function WNetAddConnection Lib "mpr.dll" Alias "WNetAddConnectionA" (ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Long


Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long


Private Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnectionA" (ByVal lpszName As String, ByVal bForce As Long) As Long


Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Private Const ERROR_ACCESS_DENIED = 5&
    Private Const ERROR_ALREADY_ASSIGNED = 85&
    Private Const ERROR_BAD_DEVICE = 1200&
    Private Const ERROR_BAD_NET_NAME = 67&
    Private Const ERROR_INVALID_PASSWORD = 86&
    Private Const ERROR_INVALID_ADDRESS = 487&
    Private Const ERROR_INVALID_PARAMETER = 87
    Private Const ERROR_MORE_DATA = 234
    Private Const ERROR_UNEXP_NET_ERR = 59&
    Private Const ERROR_NOT_CONNECTED = 2250&
    Private Const ERROR_NOT_SUPPORTED = 50&
    Private Const ERROR_OPEN_FILES = 2401&
    Private Const ERROR_NOT_ENOUGH_MEMORY = 8
    Private Const NO_ERROR = 0
    
    Private Const WN_ACCESS_DENIED = ERROR_ACCESS_DENIED
    Private Const WN_ALREADY_CONNECTED = ERROR_ALREADY_ASSIGNED
    Private Const WN_BAD_LOCALNAME = ERROR_BAD_DEVICE
    Private Const WN_BAD_NETNAME = ERROR_BAD_NET_NAME
    Private Const WN_BAD_PASSWORD = ERROR_INVALID_PASSWORD
    Private Const WN_BAD_POINTER = ERROR_INVALID_ADDRESS
    Private Const WN_BAD_VALUE = ERROR_INVALID_PARAMETER
    Private Const WN_MORE_DATA = ERROR_MORE_DATA
    Private Const WN_NET_ERROR = ERROR_UNEXP_NET_ERR
    Private Const WN_NOT_CONNECTED = ERROR_NOT_CONNECTED
    Private Const WN_NOT_SUPPORTED = ERROR_NOT_SUPPORTED
    Private Const WN_OPEN_FILES = ERROR_OPEN_FILES
    Private Const WN_OUT_OF_MEMORY = ERROR_NOT_ENOUGH_MEMORY
    Private Const WN_SUCCESS = NO_ERROR


Function GetUNCPath(DriveLetter As String, DrivePath, ErrorMsg As String) As Long
    On Local Error GoTo GetUNCPath_Err
    Dim status As Long
    Dim lpszLocalName As String
    Dim lpszRemoteName As String
    Dim cbRemoteName As Long
    lpszLocalName = DriveLetter
    If Right$(lpszLocalName, 1) <> Chr$(0) Then lpszLocalName = lpszLocalName & Chr$(0)
    lpszRemoteName = String$(255, Chr$(32))
    cbRemoteName = Len(lpszRemoteName)
    status = WNetGetConnection(lpszLocalName, _
    lpszRemoteName, _
    cbRemoteName)
    
    GetUNCPath = status


    Select Case status
        Case WN_SUCCESS
        ' all is successful...
        Case WN_NOT_SUPPORTED
        ErrorMsg = "This Function is not supported"
        Case WN_OUT_OF_MEMORY
        ErrorMsg = "The System is Out of Memory."
        Case WN_NET_ERROR
        ErrorMsg = "An error occurred On the network"
        Case WN_BAD_POINTER
        ErrorMsg = "The network path is invalid"
        Case WN_BAD_VALUE
        ErrorMsg = "Invalid local device name"
        Case WN_NOT_CONNECTED
        ErrorMsg = "The drive is not connected"
        Case WN_MORE_DATA
        ErrorMsg = "The buffer was too small to return the fileservice name"
        Case Else
        ErrorMsg = "Unrecognized Error - " & Str$(status) & "."
    End Select



If Len(ErrorMsg) Then
    DrivePath = ""
Else
    ' Trim it, and remove any nulls
    DrivePath = StripNulls(lpszRemoteName)
End If

GetUNCPath_End:
Exit Function
GetUNCPath_Err:
MsgBox Err.Description, vbInformation
Resume GetUNCPath_End
End Function

'----------------------------------------------------------------
'     -----------------------------------
' GetUserName routine
'----------------------------------------------------------------
'     -----------------------------------


Function sGetUserName() As String

    Dim lpBuffer As String * 255
    Dim lRet As Long
    lRet = GetUserName(lpBuffer, 255)
    sGetUserName = StripNulls(lpBuffer)
End Function

'----------------------------------------------------------------
'     -----------------------------------
' StripNulls routine
'----------------------------------------------------------------
'     -----------------------------------


Private Function StripNulls(s As String) As String

    'Truncates string at first null character, any text after first n
    '     ull is lost
    Dim I As Integer
    StripNulls = s


    If Len(s) Then
        I = InStr(s, Chr$(0))
        If I Then StripNulls = Left$(s, I - 1)
    End If

End Function

'----------------------------------------------------------------
'     -----------------------------------
' MapNetworkDrive routine
'----------------------------------------------------------------
'     -----------------------------------


Function MapNetworkDrive(UNCname As String, _
    Password As String, _
    DriveLetter As String, _
    ErrorMsg As String) As Long
    
    Dim status As Long
    Dim tUNCname As String, tPassword As String, tDriveLetter As String
    On Local Error GoTo MapNetworkDrive_Err
    tUNCname = UNCname
    tPassword = Password
    tDriveLetter = DriveLetter
    If Right$(tUNCname, 1) <> Chr$(0) Then tUNCname = tUNCname & Chr$(0)
    If Right$(tPassword, 1) <> Chr$(0) Then tPassword = tPassword & Chr$(0)
    If Right$(tDriveLetter, 1) <> Chr$(0) Then tDriveLetter = tDriveLetter & Chr$(0)
    status = WNetAddConnection(tUNCname, tPassword, tDriveLetter)


    Select Case status
        Case WN_SUCCESS
        ErrorMsg = ""
        Case WN_NOT_SUPPORTED
        ErrorMsg = "Function is not supported."
        Case WN_OUT_OF_MEMORY:
        ErrorMsg = "The system is out of memory."
        Case WN_NET_ERROR
        ErrorMsg = "An error occurred On the network."
        Case WN_BAD_POINTER
        ErrorMsg = "The network path is invalid."
        Case WN_BAD_NETNAME
        ErrorMsg = "Invalid network resource name."
        Case WN_BAD_PASSWORD
        ErrorMsg = "The password is invalid."
        Case WN_BAD_LOCALNAME
        ErrorMsg = "The local device name is invalid."
        Case WN_ACCESS_DENIED
        ErrorMsg = "A security violation occurred."
        Case WN_ALREADY_CONNECTED
        ErrorMsg = "This drive letter is already connected to a network drive."
        Case Else
        ErrorMsg = "Unrecognized Error - " & Str$(status) & "."
    End Select

MapNetworkDrive = status
MapNetworkDrive_End:
Exit Function
MapNetworkDrive_Err:
MsgBox Err.Description, vbInformation
Resume MapNetworkDrive_End
End Function

'----------------------------------------------------------------
'     -----------------------------------
' DisconnectNetworkDrive routine
'----------------------------------------------------------------
'     -----------------------------------


Function DisconnectNetworkDrive(DriveLetter As String, ForceFileClose As Long, ErrorMsg As String) As Long
        
        Dim status As Long
        Dim tDriveLetter As String
        On Local Error GoTo DisconnectNetworkDrive_Err
        tDriveLetter = DriveLetter
        If Right$(tDriveLetter, 1) <> Chr$(0) Then tDriveLetter = tDriveLetter & Chr$(0)
        status = WNetCancelConnection(tDriveLetter, ForceFileClose)


        Select Case status
            Case WN_SUCCESS
            ErrorMsg = ""
            Case WN_BAD_POINTER:
            ErrorMsg = "The network path is invalid."
            Case WN_BAD_VALUE
            ErrorMsg = "Invalid local device name"
            Case WN_NET_ERROR:
            ErrorMsg = "An error occurred On the network."
            Case WN_NOT_CONNECTED
            ErrorMsg = "The drive is not connected"
            Case WN_NOT_SUPPORTED
            ErrorMsg = "This Function is not supported"
            Case WN_OPEN_FILES
            ErrorMsg = "Files are in use On this service. Drive was not disconnected."
            Case WN_OUT_OF_MEMORY:
            ErrorMsg = "The System is Out of Memory"
            Case Else:
            ErrorMsg = "Unrecognized Error - " & Str$(status) & "."
        End Select

    DisconnectNetworkDrive = status
DisconnectNetworkDrive_End:
    Exit Function
DisconnectNetworkDrive_Err:
    MsgBox Err.Description, vbInformation
    Resume DisconnectNetworkDrive_End
End Function

Sub TurnOnScrnSvr()


    Dim lResult As Long
    Const WM_SYSCOMMAND = &H112
    Const SC_SCREENSAVE = &HF140
    ' send this message to all the top level windows
    ' one of them will start a screen saver...
    lResult = SendMessage(-1, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
End Sub
                            


