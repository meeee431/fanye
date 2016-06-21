VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "登录"
   ClientHeight    =   2430
   ClientLeft      =   4530
   ClientTop       =   5745
   ClientWidth     =   5775
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5775
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -30
      TabIndex        =   7
      Top             =   705
      Width           =   7695
   End
   Begin VB.TextBox txtPassWord 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1905
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1800
      Width           =   2070
   End
   Begin VB.TextBox txtUser 
      Height          =   315
      Left            =   1905
      TabIndex        =   1
      Top             =   1350
      Width           =   2070
   End
   Begin RTComctl3.CoolButton cmdOk 
      Height          =   345
      Left            =   4200
      TabIndex        =   4
      Top             =   1320
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "确定(O)"
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
      MICON           =   "frmLogin.frx":08CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   4200
      TabIndex        =   5
      Top             =   1770
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "取消(&C)"
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
      MICON           =   "frmLogin.frx":08E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   5775
      TabIndex        =   8
      Top             =   0
      Width           =   5775
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客运站务管理系统用户身份验证"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   570
         TabIndex        =   9
         Top             =   300
         Width           =   2520
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入用户及密码"
      Height          =   180
      Left            =   945
      TabIndex        =   6
      Top             =   1035
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密码(&P):"
      Height          =   180
      Left            =   945
      TabIndex        =   2
      Top             =   1867
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户名(&U):"
      Height          =   180
      Left            =   945
      TabIndex        =   0
      Top             =   1417
      Width           =   900
   End
   Begin VB.Image imgLogin 
      Height          =   480
      Left            =   240
      Picture         =   "frmLogin.frx":0902
      Top             =   1050
      Width           =   480
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'窗口常数
Const HWND_TOP = 0
Const HWND_TOPMOST = -1
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public m_bLoginOk As Boolean
'Public m_frmFlashFrom As Form
Public m_szPasword As String
Public m_szUserID As String

'用于返回的登录活动用户对象
'Public m_oActiveUser As ActiveUser

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
'此处填写用户登录代码,允许重复三次
    m_bLoginOk = True
    Unload Me
End Sub

Private Sub Form_Activate()
    txtUser.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If ActiveControl = txtPassWord And cmdOk.Enabled Then
            cmdOk_Click
        Else
            SendKeys "{TAB}"
        End If
    End If
End Sub

Private Sub Form_Load()
Dim strDate As String
Dim listWho As Byte
'获得用户上次登录的用户名
'    Load frmSplash
'    frmSplash.Show
    m_bLoginOk = False
    txtUser.Text = m_szUserID
    cmdOk.Enabled = IIf(txtUser.Text = "", False, True)
    'Set m_oActiveUser = CreateObject("SNSystem.ActiveUser")
    
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
   
'-------------------------
'-------------------------
End Sub


Private Sub Form_Unload(Cancel As Integer)
    m_szPasword = txtPassWord.Text
    m_szUserID = txtUser.Text
'    If Not m_bLoginOk Then
'        Set m_oActiveUser = Nothing
'    End If
End Sub

Private Sub txtPassWord_GotFocus()
    txtPassWord.SelStart = 0
    txtPassWord.SelLength = Len(txtPassWord.Text)
End Sub

Private Sub txtPassWord_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        cmdOk.SetFocus
'    End If
End Sub

Private Sub txtUser_Change()
    cmdOk.Enabled = IIf(txtUser.Text = "", False, True)
End Sub

Private Sub txtUser_GotFocus()
    txtUser.SelStart = 0
    txtUser.SelLength = Len(txtUser.Text)
End Sub
