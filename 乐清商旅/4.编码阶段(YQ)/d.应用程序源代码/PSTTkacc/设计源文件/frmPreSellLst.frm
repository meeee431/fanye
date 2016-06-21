VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmPreSellLst 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "预售票平衡明细表"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   6195
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   7665
      TabIndex        =   1
      Top             =   0
      Width           =   7665
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择查询条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   2
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.TextBox txtSellStation 
      Height          =   315
      Left            =   4440
      TabIndex        =   0
      Top             =   1013
      Width           =   1635
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Top             =   1620
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
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
      MICON           =   "frmPreSellLst.frx":0000
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
      Height          =   315
      Left            =   4680
      TabIndex        =   4
      Top             =   1620
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmPreSellLst.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   1013
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      Format          =   23920640
      CurrentDate     =   36572
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   1620
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "帮助"
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
      MICON           =   "frmPreSellLst.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "选择月份(&B):"
      Height          =   180
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&T):"
      Height          =   180
      Left            =   3360
      TabIndex        =   7
      Top             =   1080
      Width           =   900
   End
End
Attribute VB_Name = "frmPreSellLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Implements IConditionForm

Const cszFileName = "预售票平衡明细表.xls"

Public m_bOk As Boolean
Private m_rsData As Recordset
Private m_vaCustomData As Variant

Private Sub adUser_DataChange()
    'EnableOK
End Sub

Private Sub cboSellStation_Change()
   'FillSellerEx
End Sub

Private Sub cboSellStation_Click()
'    cboSellStation_Change
End Sub

Private Sub cmdCancel_Click()
m_bOk = False
    Unload Me
End Sub


Private Sub cmdok_Click()
    Dim oSellerStat As New TicketSellerDim
    Dim aszUserID() As String
    Dim nSelUserCount As Integer
    Dim szSellerSation As String
    Dim i As Integer
    Dim aszSellStation() As String
    On Error GoTo Error_Handle

    If Trim(txtSellStation.Text) = "" Then
        MsgBox "请输入上车站", vbInformation + vbOKOnly, Me.Caption
        Exit Sub
    End If
        
    SetMouseBusy True
    aszSellStation = Split(txtSellStation.Text, ",")
        Set m_rsData = oSellerStat.PreSellTicketlst(aszSellStation, dtpBeginDate.Value)
   ' End If
    ReDim m_vaCustomData(1 To 3, 1 To 2)
    m_vaCustomData(1, 1) = "统计开始日期"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY年MM月")
    
    m_vaCustomData(2, 1) = "上车站"
    m_vaCustomData(2, 2) = txtSellStation.Text
    
    m_vaCustomData(3, 1) = "制表人"
    m_vaCustomData(3, 2) = m_oActiveUser.UserID
    
    SetMouseBusy False
   ' SaveRecentSeller adUser.RightData
    m_bOk = True
    Unload Me
    Exit Sub
Error_Handle:
    ShowErrorMsg
End Sub

Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()

    Dim dyNow As Date
    dyNow = m_oParam.NowDate
    dtpBeginDate.Value = Format(DateAdd("m", -1, dyNow), "yyyy-mm-01")
'   FillSellStation cboSellStation

End Sub

'填充售票员
Private Sub FillSeller()
  '  Dim oSysMan As New SystemMan
    'Dim auiUserInfo() As TUserInfo
    'Dim i As Integer, nUserCount As Integer
   ' Dim aszTemp() As String, aszTemp2() As String
    'Dim nNoSelected As Integer, nSelected As Integer
   ' Dim szTemp As String
    'Dim szRecentSeller As String
    
   ' oSysMan.Init m_oActiveUser
    'auiUserInfo = oSysMan.GetAllUser()
    'nUserCount = ArrayLength(auiUserInfo)
    'If nUserCount > 0 Then
      '  szRecentSeller = GetRecentSeller()
       ' nNoSelected = 0
       ' nSelected = 0
       ' For i = 1 To nUserCount
           ' szTemp = MakeDisplayString(auiUserInfo(i).UserID, auiUserInfo(i).UserName)
           ' If InStr(1, szRecentSeller, szTemp, vbTextCompare) = 0 Then
               ' nNoSelected = nNoSelected + 1
                'ReDim Preserve aszTemp(1 To nNoSelected)
               ' aszTemp(nNoSelected) = szTemp
           ' Else
              '  nSelected = nSelected + 1
                'ReDim Preserve aszTemp2(1 To nSelected)
               ' aszTemp2(nSelected) = szTemp
           ' End If
       ' Next
   ' End If
    'adUser.LeftData = aszTemp
    'adUser.RightData = aszTemp2
End Sub

Private Sub EnableOK()
   ' Dim nCount As Integer
'nCount = ArrayLength(adUser.RightData)
  'cmdOk.Enabled = IIf(nCount > 0, True, False)
End Sub

Private Property Get IConditionForm_CustomData() As Variant
    IConditionForm_CustomData = m_vaCustomData
End Property

Private Property Get IConditionForm_FileName() As String
    IConditionForm_FileName = cszFileName
End Property

Private Property Get IConditionForm_RecordsetData() As Recordset
    Set IConditionForm_RecordsetData = m_rsData
End Property






