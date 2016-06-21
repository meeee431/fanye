Attribute VB_Name = "mdAdvanceLib"
Option Explicit
'=========================================================================================
'Author:
'Detail:中间层共用模块
'=========================================================================================


'------------------------------------------------------------------------------------------
'以下常量声明
'---------------------------------------------------------
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32s = 0
Private Const MAX_PATH = 255
Const OFS_MAXPATHNAME = 128
Const OF_CREATE = &H1000
Const OF_READ = &H0
Const OF_WRITE = &H1
Const GENERIC_WRITE = &H40000000
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const OPEN_EXISTING = 3



Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6

'WinSock
'----------------------------------------


Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Const ERROR_SUCCESS       As Long = 0
Public Const WS_VERSION_REQD     As Long = &H101
Public Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD    As Long = 1
Public Const SOCKET_ERROR        As Long = -1

Public Const WM_NCACTIVATE = &H86
Public Const SC_CLOSE = &HF060&

Public Const MIIM_STATE = &H1&
Public Const MIIM_ID = &H2&

Public Const MFS_GRAYED = &H3&
Public Const MFS_CHECKED = &H8&



'------------------------------------------------------------------------------------------
'以下类型定义
'---------------------------------------------------------
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Public Enum ESystemPathType '操作系统的路径类型
    MyDocuments = 1
    Windows = 2
    WindowsSystem = 3
    CurrentPath = 4
End Enum
Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type



Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Public Type HOSTENT
   hName      As Long
   hAliases   As Long
   hAddrType  As Integer
   hLen       As Integer
   hAddrList  As Long
End Type

Public Type WSADATA
   wVersion      As Integer
   wHighVersion  As Integer
   szDescription(0 To MAX_WSADescription)   As Byte
   szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
   wMaxSockets   As Integer
   wMaxUDPDG     As Integer
   dwVendorInfo  As Long
End Type


'------------------------------------------------------------------------------------------
'API定义
'---------------------------------------------------------
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetComputerNameAPI Lib "kernel32" Alias "GetComputerNameA" _
 (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long



    '文件类
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
'Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

'Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
'   (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
'Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

   
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long

Public Declare Function WSAStartup Lib "WSOCK32.DLL" _
   (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
   
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Public Declare Function gethostname Lib "WSOCK32.DLL" _
   (ByVal szHost As String, ByVal dwHostLen As Long) As Long
   
Public Declare Function gethostbyname Lib "WSOCK32.DLL" _
   (ByVal szHost As String) As Long
   
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
      
'Public Const SWP_NOMOVE = &H2 '不移动窗体
'Public Const SWP_NOSIZE = &H1 '不改变窗体尺寸
'Public Const HWND_TOPMOST = -1 '窗体总在最前面
'Public Const HWND_NOTOPMOST = -2 '窗体不在最前面

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
Private Type POINTAPI
  x As Long
  y As Long
End Type
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPOINT As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long



Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long


'Private m_lOldProc As Long 'MDI窗体 winproc 句柄

Public Function GetComputerName() As String
    ' Set or retrieve the name of the computer.
    Dim strBuffer As String
    Dim lngLen As Long
        
    strBuffer = Space(255 + 1)
    lngLen = Len(strBuffer)
    If CBool(GetComputerNameAPI(strBuffer, lngLen)) Then
        GetComputerName = MidA(strBuffer, 1, lngLen)
    Else
        GetComputerName = ""
    End If
End Function

Public Function GetSystemPath(pePathType As ESystemPathType) As String
'得到系统路径
On Error GoTo ErrHandle
    Dim oFreeReg As New CFreeReg
    Select Case pePathType
        Case ESystemPathType.MyDocuments
            oFreeReg.Init "Explorer", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion"
            GetSystemPath = oFreeReg.GetSetting("Shell Folders", "Personal")
        Case ESystemPathType.CurrentPath
        Case ESystemPathType.Windows
        Case ESystemPathType.WindowsSystem
    End Select
    Exit Function
ErrHandle:
    GetSystemPath = ""
End Function


'获得Windows的Temp目录下的临时文件名
Public Function GetTempFile() As String
    Dim oFileSystem '文件系统对象
    Dim szTempFile As String
    Dim szTempDir As String
    
    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
    szTempDir = Environ("temp") + "\"
    szTempFile = szTempDir + oFileSystem.GetTempName
    While oFileSystem.FileExists(szTempFile)
        szTempFile = szTempDir + oFileSystem.GetTempName
    Wend
    GetTempFile = szTempFile
End Function

Public Function FileIsExist(FileName As String) As Boolean
'判断文件是否存在
    Dim WFD As WIN32_FIND_DATA
'    Dim Path As String
    Dim lSearch As Long
'    lSearch = FindFirstFile(Path & "*", WFD)
    lSearch = FindFirstFile(FileName & "*", WFD)
    If lSearch = -1 Then
        FileIsExist = False
    Else
        FindClose lSearch
        FileIsExist = True
    End If
End Function
'得到一个二进制文件的内容
Public Function GetFileData(ByVal pszFileName As String, ByRef vFileData) As Boolean
'返回是否成功
    Dim abReturn() As Byte
    Dim lRet As Long    '实际读入
    Dim lf As Long
    Dim lsize As Long
    Dim lError As Long
    Dim OF As OFSTRUCT
    lf = OpenFile(pszFileName, OF, OF_READ)
    lError = GetLastError()
    If lError <> 0 Then GoTo ErrHandle
    lsize = GetFileSize(lf, 0)
    If lsize = 0 Then GoTo ErrHandle
    ReDim abReturn(1 To lsize)
    ReadFile lf, abReturn(1), lsize, lRet, ByVal 0&
    CloseHandle lf
    
    vFileData = abReturn
    GetFileData = True
    Exit Function
ErrHandle:
    GetFileData = False
End Function
'将二进制数据存成一个文件
Public Function SaveDataToFile(ByVal pszFileName As String, pvData As Variant) As Boolean
'返回是否成功
On Error GoTo ErrHandle
    Dim abBytes() As Byte
    Dim hNewFile As Long
    Dim lRet As Long
    abBytes = pvData

    Dim hFile As Long
    hFile = FreeFile
    Open pszFileName For Binary Access Write As hFile
    Put hFile, , abBytes
    Close hFile
    
'以下API完成
'    hNewFile = CreateFile(pszFileName, GENERIC_WRITE, 0, ByVal 0&, OPEN_EXISTING, 0, 0)
'    Dim lError As Long
'    lError = GetLastError()
'    If lError <> 0 Then GoTo ErrHandle
'
'    WriteFile hNewFile, abBytes(0), UBound(abBytes) - LBound(abBytes) + 1, lRet, ByVal 0&
'    CloseHandle hNewFile
    SaveDataToFile = True
    Exit Function
ErrHandle:
    SaveDataToFile = False
End Function



Public Sub EvisibleCloseButton(sForm As Form)

    Const xMenuID = 10&
    Dim hMenu As Long, MII As MENUITEMINFO
    
    '取得系统菜单的hMnu
    hMenu = GetSystemMenu(sForm.hwnd, 0)
    '先填好MENUITEMINFO的数据结构所需之栏
    MII.cbSize = Len(MII)
    MII.dwTypeData = String(80, 0)
    MII.cch = Len(MII.dwTypeData)
    MII.fMask = MIIM_STATE
    MII.wID = SC_CLOSE
    '读取关闭命令(SC_CLOSE)的消息
    GetMenuItemInfo hMenu, SC_CLOSE, False, MII
    
    MII.wID = xMenuID        '设置新的MENU ID
    MII.fMask = MIIM_ID
    SetMenuItemInfo hMenu, SC_CLOSE, False, MII
    
    MII.fState = MII.fState Or MFS_GRAYED
    MII.fMask = MIIM_STATE
    SetMenuItemInfo hMenu, MII.wID, False, MII
    
    SendMessage sForm.hwnd, WM_NCACTIVATE, True, ByVal 0&
End Sub

'得到CD盘
Public Function GetCD() As String
    Dim rtn As String
    Dim AllDrives As String
    Dim JustOneDrive As String
    Dim DriveType As Integer
On Error GoTo ErrorHandle
    AllDrives = Space$(64)
    rtn = GetLogicalDriveStrings(Len(AllDrives), AllDrives)
    AllDrives = Left(AllDrives, rtn)
    Do
      rtn = InStr(AllDrives, Chr(0))
      If rtn Then
         JustOneDrive = Left(AllDrives, rtn)
         AllDrives = Mid(AllDrives, rtn + 1, Len(AllDrives))
         
         rtn = GetDriveType(JustOneDrive)
         If rtn = DRIVE_CDROM Then
           GetCD = UCase(JustOneDrive)
            Exit Do
         End If
      End If
    Loop Until AllDrives = "" Or DriveType = DRIVE_CDROM
Exit Function
ErrorHandle:
  MsgBox err.Description, vbInformation
End Function


'以下这些函数是得到机子名或机子IP的函数
'------------------------------------------------------

Public Function HiByte(ByVal wParam As Integer) As Byte
  
  'note: VB4-32 users should declare this function As Integer
   HiByte = (wParam And &HFF00&) \ (&H100)
 
End Function


Public Function LoByte(ByVal wParam As Integer) As Byte

  'note: VB4-32 users should declare this function As Integer
   LoByte = wParam And &HFF&

End Function


Public Sub SocketsCleanup()

    If WSACleanup() <> ERROR_SUCCESS Then
        'MsgBox "Socket error occurred in Cleanup."
    End If
    
End Sub

Public Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA
   Dim sLoByte As String
   Dim sHiByte As String
   
   If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
'      'MsgBox "The 32-bit Windows Socket is not responding."
      SocketsInitialize = False
      Exit Function
   End If
   
   
   If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
'        'MsgBox "This application requires a minimum of " & _
'                CStr(MIN_SOCKETS_REQD) & " supported sockets."
        
        SocketsInitialize = False
        Exit Function
    End If
   
   
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
     (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
      HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      
      sHiByte = CStr(HiByte(WSAD.wVersion))
      sLoByte = CStr(LoByte(WSAD.wVersion))
      
'      'MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
'             " is not supported by 32-bit Windows Sockets."
      
      SocketsInitialize = False
      Exit Function
      
   End If
    
    
  'must be OK, so lets do it
   SocketsInitialize = True
        
End Function

'得到指定机子的IP地址(缺省为本机)
Public Function GetComputerIP(Optional pszComputerName As String = "") As String
   Dim sHostName    As String '* 256
   Dim lpHost    As Long
   Dim HOST      As HOSTENT
   Dim dwIPAddr  As Long
   Dim tmpIPAddr() As Byte
   Dim i         As Integer
   Dim sIPAddr  As String
   
   On Error GoTo Error_Handle
   
   If Not SocketsInitialize() Then
      GetComputerIP = ""
      Exit Function
   End If
    
  'gethostname returns the name of the local host into
  'the buffer specified by the name parameter. The host
  'name is returned as a null-terminated string. The
  'form of the host name is dependent on the Windows
  'Sockets provider - it can be a simple host name, or
  'it can be a fully qualified domain name. However, it
  'is guaranteed that the name returned will be successfully
  'parsed by gethostbyname and WSAAsyncGetHostByName.

  'In actual application, if no local host name has been
  'configured, gethostname must succeed and return a token
  'host name that gethostbyname or WSAAsyncGetHostByName
  'can resolve.
  If pszComputerName = "" Then
    sHostName = Space(256)
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
       GetComputerIP = ""
'       'MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
'               " has occurred. Unable to successfully get Host Name."
       SocketsCleanup
       Exit Function
    End If
  Else
    sHostName = pszComputerName
  End If
   
  'gethostbyname returns a pointer to a HOSTENT structure
  '- a structure allocated by Windows Sockets. The HOSTENT
  'structure contains the results of a successful search
  'for the host specified in the name parameter.

  'The application must never attempt to modify this
  'structure or to free any of its components. Furthermore,
  'only one copy of this structure is allocated per thread,
  'so the application should copy any information it needs
  'before issuing any other Windows Sockets function calls.

  'gethostbyname function cannot resolve IP address strings
  'passed to it. Such a request is treated exactly as if an
  'unknown host name were passed. Use inet_addr to convert
  'an IP address string the string to an actual IP address,
  'then use another function, gethostbyaddr, to obtain the
  'contents of the HOSTENT structure.
   sHostName = Trim$(sHostName)
   lpHost = gethostbyname(sHostName)
    
   If lpHost = 0 Then
      GetComputerIP = ""
'      'MsgBox "Windows Sockets are not responding. " & _
'              "Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
   End If
    
  'to extract the returned IP address, we have to copy
  'the HOST structure and its members
   CopyMemory HOST, lpHost, Len(HOST)
   CopyMemory dwIPAddr, HOST.hAddrList, 4
   
  'create an array to hold the result
   ReDim tmpIPAddr(1 To HOST.hLen)
   CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen
   
  'and with the array, build the actual address,
  'appending a period between members
   For i = 1 To HOST.hLen
      sIPAddr = sIPAddr & tmpIPAddr(i) & "."
   Next i
  
  'the routine adds a period to the end of the
  'string, so remove it ErrorHandle
   GetComputerIP = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
   
   SocketsCleanup
   Exit Function
Error_Handle:
   GetComputerIP = ""
End Function






Public Function IsOsNt() As Boolean
    Dim viVerInfo As OSVERSIONINFO
    viVerInfo.dwOSVersionInfoSize = Len(viVerInfo)
    If GetVersionEx(viVerInfo) Then
        IsOsNt = IIf(viVerInfo.dwPlatformId = VER_PLATFORM_WIN32_NT, True, False)
    Else
        IsOsNt = False
    End If
End Function



Public Function FormatLen(ByVal pszStr As String, ByVal pnLen As Integer) As String
    '返回指定长度的字符串
    Dim szTemp As String
    If pnLen > 0 Then
        If LenA(pszStr) >= pnLen Then
            FormatLen = Left(pszStr, pnLen)
        Else
            FormatLen = pszStr & Space(pnLen - LenA(pszStr))
        End If
    Else
        FormatLen = ""
    End If
    
End Function



