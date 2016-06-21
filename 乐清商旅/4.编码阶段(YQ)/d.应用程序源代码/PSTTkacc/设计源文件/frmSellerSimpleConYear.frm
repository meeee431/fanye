VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmSellerSimpleConYear 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "售票员售票年报"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   HelpContextID   =   6001401
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "报表说明"
      Height          =   555
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   6975
      Begin VB.Label Label3 
         Caption         =   "按票种指定时间段，统计票种人数、金额。用于统计售票员的售票情况。"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   6435
      End
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   315
      Left            =   4080
      TabIndex        =   3
      Top             =   180
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      Format          =   24576000
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   315
      Left            =   1140
      TabIndex        =   1
      Top             =   180
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy年"
      Format          =   24576003
      CurrentDate     =   36572
   End
   Begin RTComctl3.CoolButton cmdCancel 
      TX         =   "取消(&C)"
      Height          =   315
      Left            =   5880
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin RTComctl3.CoolButton cmdOk 
      TX         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   315
      Left            =   5880
      TabIndex        =   4
      Top             =   180
      Width           =   1215
   End
   Begin prjTKAcc.AddDel adUser 
      Height          =   2535
      Left            =   0
      TabIndex        =   8
      Top             =   800
      Width           =   5800
      _ExtentX        =   10239
      _ExtentY        =   4471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonWidth     =   1215
      ButtonHeight    =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "结束日期(&E)"
      Height          =   180
      Left            =   3060
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "年 份(&B)"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmSellerSimpleConYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Implements IConditionForm

Const cszFileName = "售票员售票年报模板.cll"

Public m_bOk As Boolean
Private m_rsData As Recordset
Private m_vaCustomData As Variant

Private Sub adUser_DataChange()
    EnableOk
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim oSellerStat As New TicketSellerDim
    Dim aszUserID() As String
    Dim nSelUserCount As Integer
    Dim i As Integer
    
    On Error GoTo Error_Handle
    '生成Recordset
    nSelUserCount = ArrayLength(adUser.RightData)
    dtpBeginDate.Value = CDate(Year(dtpBeginDate.Value) & "-01" & "-01")
    dtpEndDate.Value = DateAdd("D", -1, DateAdd("M", 12, dtpBeginDate.Value))
    
    If nSelUserCount > 0 Then
        oSellerStat.Init m_oActiveUser
        ReDim aszUserID(1 To nSelUserCount)
        For i = 1 To nSelUserCount
            aszUserID(i) = ResolveDisplay(adUser.RightData(i))
        Next
        Set m_rsData = oSellerStat.SellerDateStat(aszUserID, dtpBeginDate.Value, dtpEndDate.Value)
    End If
    ReDim m_vaCustomData(1 To 2, 1 To 2)
    m_vaCustomData(1, 1) = "统计开始日期"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY年")
    
    'm_vaCustomData(2, 1) = "统计结束日期"
    'm_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")
    SaveRecentSeller adUser.RightData
    m_bOk = True
    Unload Me
    Exit Sub
Error_Handle:
    ShowErrMsg
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    m_bOk = False
    dtpBeginDate.Value = m_oParam.NowDate
    dtpEndDate.Value = m_oParam.NowDate
    FillSeller
    EnableOk
End Sub

'填充售票员
Private Sub FillSeller()
    Dim oSysMan As New SystemMan
    Dim auiUserInfo() As TUserInfo
    Dim i As Integer, nUserCount As Integer
    Dim aszTemp() As String, aszTemp2() As String
    Dim nNoSelected As Integer, nSelected As Integer
    Dim szTemp As String
    Dim szRecentSeller As String
    
    oSysMan.Init m_oActiveUser
    auiUserInfo = oSysMan.GetAllUser()
    nUserCount = ArrayLength(auiUserInfo)
    If nUserCount > 0 Then
        szRecentSeller = GetRecentSeller()
        nNoSelected = 0
        nSelected = 0
        For i = 1 To nUserCount
            szTemp = MakeDisplayString(auiUserInfo(i).UserID, auiUserInfo(i).UserName)
            If InStr(1, szRecentSeller, szTemp, vbTextCompare) = 0 Then
                nNoSelected = nNoSelected + 1
                ReDim Preserve aszTemp(1 To nNoSelected)
                aszTemp(nNoSelected) = szTemp
            Else
                nSelected = nSelected + 1
                ReDim Preserve aszTemp2(1 To nSelected)
                aszTemp2(nSelected) = szTemp
            End If
        Next
    End If
    adUser.LeftData = aszTemp
    adUser.RightData = aszTemp2
End Sub

Private Sub EnableOk()
    Dim nCount As Integer
    nCount = ArrayLength(adUser.RightData)
    cmdOk.Enabled = IIf(nCount > 0, True, False)
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
