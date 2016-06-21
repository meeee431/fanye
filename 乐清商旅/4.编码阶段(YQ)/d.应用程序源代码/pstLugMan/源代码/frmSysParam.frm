VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmSysParam 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置系统参数"
   ClientHeight    =   4095
   ClientLeft      =   5910
   ClientTop       =   2220
   ClientWidth     =   6075
   HelpContextID   =   7000201
   Icon            =   "frmSysParam.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6075
   Begin VB.TextBox txtCare 
      Appearance      =   0  'Flat
      Height          =   675
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   2610
      Width           =   4125
   End
   Begin VB.TextBox txtReturnRatio1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1680
      TabIndex        =   8
      Top             =   2190
      Width           =   1485
   End
   Begin RTComctl3.CoolButton cmdExit 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   4800
      TabIndex        =   10
      Top             =   3630
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "取消"
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
      MICON           =   "frmSysParam.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "行包单/签发单号长度设置"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   5805
      Begin VB.TextBox txtCheckLuggageIDLen 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   1440
         Width           =   1485
      End
      Begin VB.TextBox txtLuggageIDNumber 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4320
         TabIndex        =   4
         Top             =   255
         Width           =   1215
      End
      Begin VB.TextBox txtLuggageIDPrefix 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   255
         Width           =   1485
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   300
         X2              =   5490
         Y1              =   1305
         Y2              =   1305
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   300
         X2              =   5490
         Y1              =   1290
         Y2              =   1290
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmSysParam.frx":0028
         Height          =   615
         Left            =   750
         TabIndex        =   12
         Top             =   630
         Width           =   4830
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "说明:"
         Height          =   180
         Left            =   270
         TabIndex        =   11
         Top             =   630
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "前缀部分(P):"
         Height          =   180
         Left            =   270
         TabIndex        =   1
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "数字部分(&N):"
         Height          =   180
         Left            =   3210
         TabIndex        =   3
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "签发单长度(&C):"
         Height          =   180
         Left            =   270
         TabIndex        =   5
         Top             =   1470
         Width           =   1260
      End
   End
   Begin RTComctl3.CoolButton cmdSave 
      Default         =   -1  'True
      Height          =   345
      Left            =   3570
      TabIndex        =   9
      Top             =   3630
      Width           =   1125
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmSysParam.frx":00AF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   3120
      Left            =   -90
      TabIndex        =   14
      Top             =   3360
      Width           =   8745
      Begin RTComctl3.CoolButton cmdHelp 
         Height          =   315
         Left            =   240
         TabIndex        =   15
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "帮助(&H)"
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
         MICON           =   "frmSysParam.frx":00CB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "额外打印内容:"
      Height          =   435
      Left            =   390
      TabIndex        =   17
      Top             =   2610
      Width           =   1245
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3240
      TabIndex        =   13
      Top             =   2220
      Width           =   120
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "退运手续费(&R):"
      Height          =   180
      Left            =   390
      TabIndex        =   7
      Top             =   2250
      Width           =   1260
   End
End
Attribute VB_Name = "frmSysParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'行包参数
Dim nLuggageIDPrefixLen As Integer '行包单号前缀部分长
Dim nLuggageIDNumberLen As Integer '行包单号数字部分长
Dim nCheckLuggageIDNumberLen As Integer  '签发单数字部分长度
Dim dbReturnRatio1 As Double '行包受理单退运费率1


Private Sub cmdExit_Click()
    Unload Me

End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo here

    If nLuggageIDPrefixLen <> txtLuggageIDPrefix.Text Then m_oLugParam.LuggageIDPrefixLen = txtLuggageIDPrefix.Text
    nLuggageIDPrefixLen = txtLuggageIDPrefix.Text
    If nLuggageIDNumberLen <> txtLuggageIDNumber.Text Then m_oLugParam.LuggageIDNumberLen = txtLuggageIDNumber.Text
    nLuggageIDNumberLen = txtLuggageIDNumber.Text
    If nCheckLuggageIDNumberLen <> txtCheckLuggageIDLen.Text Then m_oLugParam.CarrySheetIDNumberLen = txtCheckLuggageIDLen.Text
    nCheckLuggageIDNumberLen = txtCheckLuggageIDLen.Text
    If dbReturnRatio1 <> txtReturnRatio1.Text Then m_oLugParam.LuggageReturnRatio1 = Val(txtReturnRatio1.Text) / 100
    dbReturnRatio1 = txtReturnRatio1.Text
    
    '行包注意事项
    Dim oFreeReg As CFreeReg
    Set oFreeReg = New CFreeReg
    oFreeReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    oFreeReg.SaveSetting cszLuggageAccount, "CareContent", Trim(txtCare.Text)
    
    
    
    Unload Me
    Exit Sub
here:

    ShowErrorMsg
End Sub

Private Sub Form_Load()
    On Error GoTo here
    AlignFormPos Me

    Dim oParam As New LuggageParam
    m_oLugParam.Init m_oAUser


    nLuggageIDPrefixLen = m_oLugParam.LuggageIDPrefixLen
    nLuggageIDNumberLen = m_oLugParam.LuggageIDNumberLen
    nCheckLuggageIDNumberLen = m_oLugParam.CarrySheetIDNumberLen
    dbReturnRatio1 = m_oLugParam.LuggageReturnRatio1 * 100


    txtLuggageIDPrefix.Text = nLuggageIDPrefixLen
    txtLuggageIDNumber.Text = nLuggageIDNumberLen
    txtCheckLuggageIDLen.Text = nCheckLuggageIDNumberLen
    txtReturnRatio1.Text = dbReturnRatio1
    
    '行包注意事项
    Dim oFreeReg As CFreeReg
    Set oFreeReg = New CFreeReg
    oFreeReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    txtCare.Text = oFreeReg.GetSetting(cszLuggageAccount, "CareContent")
    
    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveFormPos Me
End Sub

Private Sub txtCheckLuggageIDLen_Change()
    FormatTextToNumeric txtCheckLuggageIDLen, False, False

End Sub

Private Sub txtLuggageIDNumber_Change()
    FormatTextToNumeric txtLuggageIDNumber, False, False

End Sub

Private Sub txtLuggageIDPrefix_Change()
    FormatTextToNumeric txtLuggageIDPrefix, False, False
End Sub

Private Sub txtReturnRatio1_Change()
    FormatTextToNumeric txtReturnRatio1, False, True
End Sub
