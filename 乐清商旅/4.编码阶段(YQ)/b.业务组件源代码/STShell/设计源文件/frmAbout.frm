VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于"
   ClientHeight    =   4440
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5985
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   296
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin RTComctl3.CoolButton cmdOk 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   360
      Left            =   4605
      TabIndex        =   13
      Top             =   3660
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "确定"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAbout.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraProductInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "产品信息"
      Height          =   915
      Left            =   1215
      TabIndex        =   6
      Top             =   1740
      Width           =   4530
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "******"
         Height          =   180
         Left            =   660
         TabIndex        =   12
         Top             =   510
         Width           =   540
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "******"
         Height          =   180
         Left            =   660
         TabIndex        =   11
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用户:"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   510
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "公司:"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   270
         Width           =   450
      End
   End
   Begin VB.Label lblEnglishTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1215
      TabIndex        =   10
      Top             =   990
      Width           =   45
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "版本:"
      Height          =   255
      Left            =   3855
      TabIndex        =   9
      Top             =   720
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客运车站管理系统"
      Height          =   180
      Left            =   1215
      TabIndex        =   5
      Top             =   3015
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "版权所有:"
      Height          =   180
      Left            =   1215
      TabIndex        =   4
      Top             =   2760
      Width           =   810
   End
   Begin VB.Image imgProductLogo 
      Height          =   900
      Left            =   150
      Stretch         =   -1  'True
      Top             =   765
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   9
      X2              =   388
      Y1              =   222
      Y2              =   222
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "应用程序描述"
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   1200
      TabIndex        =   0
      Top             =   1290
      Width           =   5205
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "应用程序中文标题"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1215
      TabIndex        =   2
      Top             =   750
      Width           =   2985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   10
      X2              =   431
      Y1              =   223
      Y2              =   223
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "0.0.7"
      Height          =   225
      Left            =   4395
      TabIndex        =   3
      Top             =   750
      Width           =   1425
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0028
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   150
      TabIndex        =   1
      Top             =   3525
      Width           =   4335
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 注册表关键字安全选项...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' 注册表关键字 ROOT 类型...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' 独立的空的终结字符串
Const REG_DWORD = 4                      ' 32位数字

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"
'窗口常数
Const HWND_TOP = 0
Const HWND_TOPMOST = -1
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private m_VerMajor As String
Private m_VerMinor As String
Private m_VerBuild As String
Private m_AppName As String
Private m_AppEnglishName As String
Private m_AppDescription As String

'主版本号属性
Public Property Get VerMajor() As String
    VerMajor = m_VerMajor
End Property
Public Property Let VerMajor(val As String)
    m_VerMajor = val
End Property
'次版本号属性
Public Property Get VerMinor() As String
    VerMinor = m_VerMinor
End Property
Public Property Let VerMinor(val As String)
    m_VerMinor = val
End Property
'编译版本号属性
Public Property Get VerBuild() As String
    VerBuild = m_VerBuild
End Property
Public Property Let VerBuild(val As String)
    m_VerBuild = val
End Property
'应用程序中文名称
Public Property Get AppName() As String
    AppName = m_AppName
End Property
Public Property Let AppName(val As String)
    m_AppName = val
End Property
'应用程序英文名称
Public Property Get AppEnglishName() As String
    AppEnglishName = m_AppEnglishName
End Property
Public Property Let AppEnglishName(val As String)
    m_AppEnglishName = val
End Property
'应用程序描述
Public Property Get AppDescription() As String
    AppDescription = m_AppDescription
End Property
Public Property Let AppDescription(val As String)
    m_AppDescription = val
End Property
'应用程序图标
Public Property Set ProductLogoImage(NewProductLogoImage As StdPicture)
    Set imgProductLogo.Picture = NewProductLogoImage
End Property


Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "关于 " & m_AppName
    lblVersion.Caption = m_VerMajor & "." & m_VerMinor & "  " & m_VerBuild & "(Build)"
    lblTitle.Caption = m_AppName
    lblEnglishTitle.Caption = m_AppEnglishName & "  " & m_VerMajor & "." & m_VerMinor
    lblDescription.Caption = m_AppDescription
    fraProductInfo.BackColor = Me.BackColor
    
    '获得产品信息
    Dim oReg As New CFreeReg
    On Error GoTo PassRegInfo
    oReg.Init "Uninstall", HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion"
    lblCompany.Caption = oReg.GetSetting("{3A2F43DA-AAE9-453E-B1A1-5B7F9FF5704D}", "RegCompany")
    lblUser.Caption = oReg.GetSetting("{3A2F43DA-AAE9-453E-B1A1-5B7F9FF5704D}", "RegOwner")
'    lblSeriaNo.Caption = oReg.GetSetting("{3A2F43DA-AAE9-453E-B1A1-5B7F9FF5704D}", "ProductID")
PassRegInfo:
    err.Clear
    On Error GoTo 0
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' 循环计数器
    Dim rc As Long                                          ' 返回代码
    Dim hKey As Long                                        ' 打开的注册表关键字句柄
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' 注册表关键字数据类型
    Dim tmpVal As String                                    ' 注册表关键字值的临时存储器
    Dim KeyValSize As Long                                  ' 注册表关键自变量的尺寸
    '------------------------------------------------------------
    ' 打开 {HKEY_LOCAL_MACHINE...} 下的 RegKey
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' 打开注册表关键字
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 处理错误...
    
    tmpVal = String$(1024, 0)                             ' 分配变量空间
    KeyValSize = 1024                                       ' 标记变量尺寸
    
    '------------------------------------------------------------
    ' 检索注册表关键字的值...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' 获得/创建关键字值
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 处理错误
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 外接程序空终结字符串...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null 被找到,从字符串中分离出来
    Else                                                    ' WinNT 没有空终结字符串...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null 没有被找到, 分离字符串
    End If
    '------------------------------------------------------------
    ' 决定转换的关键字的值类型...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' 搜索数据类型...
    Case REG_SZ                                             ' 字符串注册关键字数据类型
        KeyVal = tmpVal                                     ' 复制字符串的值
    Case REG_DWORD                                          ' 四字节的注册表关键字数据类型
        For i = Len(tmpVal) To 1 Step -1                    ' 将每位进行转换
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' 生成值字符。 By Char。
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' 转换四字节的字符为字符串
    End Select
    
    GetKeyValue = True                                      ' 返回成功
    rc = RegCloseKey(hKey)                                  ' 关闭注册表关键字
    Exit Function                                           ' 退出
    
GetKeyError:      ' 错误发生后将其清除...
    KeyVal = ""                                             ' 设置返回值到空字符串
    GetKeyValue = False                                     ' 返回失败
    rc = RegCloseKey(hKey)                                  ' 关闭注册表关键字
End Function
