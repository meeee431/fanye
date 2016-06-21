VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmstationsell 
   BackColor       =   &H00E0E0E0&
   Caption         =   "上车站售票统计报表"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6405
   HelpContextID   =   60000210
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   6405
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtSellStationID 
      Height          =   300
      Left            =   1560
      TabIndex        =   11
      Top             =   1950
      Width           =   4395
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   7665
      TabIndex        =   8
      Top             =   0
      Width           =   7665
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择查询条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   9
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.ComboBox cboSellStation 
      Height          =   300
      ItemData        =   "frmstationsell.frx":0000
      Left            =   1560
      List            =   "frmstationsell.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1560
      Width           =   4395
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   315
      Left            =   3120
      TabIndex        =   1
      Top             =   2340
      Width           =   1215
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
      MICON           =   "frmstationsell.frx":0004
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
      TabIndex        =   2
      Top             =   2340
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
      MICON           =   "frmstationsell.frx":0020
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   315
      Left            =   4320
      TabIndex        =   3
      Top             =   960
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      Format          =   60096512
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   960
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      Format          =   60096512
      CurrentDate     =   36572
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   1320
      TabIndex        =   10
      Top             =   2340
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
      MICON           =   "frmstationsell.frx":003C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   3240
      TabIndex        =   7
      Top             =   1080
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&T):"
      Height          =   180
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   900
   End
End
Attribute VB_Name = "frmstationsell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Implements IConditionForm

Const cszFileName = "上车站售票统计报表.xls"

Public m_bOk As Boolean
Private m_rsData As Recordset
Private m_vaCustomData As Variant

Private Sub adUser_DataChange()
    'EnableOK
End Sub

Private Sub cboSellStation_Change()
   'FillSellerEx
End Sub
Public Sub FillSellerEx()
   ' Dim oUnit As New Unit
   ' Dim aszUser() As String
   ' Dim aszUser2() As String
   ' Dim aszUser22() As String
    '
   ' Dim oUser As New User
    'Dim i As Integer, nUserCount As Integer
    'Dim szRecentSeller As String
   ' Dim szTemp As String
    
   ' Dim nNoSelected As Integer, nSelected As Integer
    
    'oUnit.Init m_oActiveUser
   ' oUnit.Identify m_oParam.UnitID
   ' aszUser = oUnit.GetAllUserEX(, ResolveDisplay(cboSellStation))
   ' nUserCount = ArrayLength(aszUser)
    'If nUserCount > 0 Then
        
'        oUser.Init m_oActiveUser
       ' szRecentSeller = GetRecentSeller()
        
        'nNoSelected = 0
        'nSelected = 0
        
       ' For i = 1 To nUserCount
'            oUser.Identify aszUser(i)
           ' szTemp = MakeDisplayString(aszUser(i, 1), aszUser(i, 2))
            'If InStr(1, szRecentSeller, szTemp, vbTextCompare) = 0 Then
               ' nNoSelected = nNoSelected + 1
               ' ReDim Preserve aszUser2(1 To nNoSelected)
                'aszUser2(nNoSelected) = szTemp
           ' Else
               ' nSelected = nSelected + 1
                'ReDim Preserve aszUser22(1 To nSelected)
                'aszUser22(nSelected) = szTemp
           ' End If
        'Next
    'End If
    'adUser.LeftData = aszUser2
    'adUser.RightData = aszUser22
End Sub
Private Sub cboSellStation_Click()
'    cboSellStation_Change
End Sub

Private Sub cmdCancel_Click()
m_bOk = False
    Unload Me
End Sub

Private Sub cmdChart_Click()
    
   ' On Error GoTo Error_Handle
    '生成记录集
   ' Dim rsTemp As Recordset
    'Dim oSellerStat As New TicketSellerDim
    'Dim i, nSelUserCount As Integer
    'Dim aszUserID() As String
    'Dim frmTemp As frmChart
    'oSellerStat.Init m_oActiveUser
    'nSelUserCount = ArrayLength(adUser.RightData)
    'If nSelUserCount > 0 Then
       ' oSellerStat.Init m_oActiveUser
        'ReDim aszUserID(1 To nSelUserCount)
        'For i = 1 To nSelUserCount
           ' aszUserID(i) = ResolveDisplay(adUser.RightData(i))
        'Next
        'Set rsTemp = oSellerStat.SellerDateStat(aszUserID, dtpBeginDate.Value, dtpEndDate.Value)
    'End If
    
   ' Dim rsData As New Recordset
    'With rsData.Fields
    '    .Append "user_id", adBSTR
       ' .Append "total_number", adBigInt
    'End With
    'rsData.Open
    'rsTemp.MoveFirst
    'For i = 1 To rsTemp.RecordCount
        'rsData.AddNew
        'rsData!user_id = FormatDbValue(rsTemp!user_id)
        'rsData!total_number = FormatDbValue(rsTemp!total_number)
        'rsTemp.MoveNext
        'rsData.Update
    'Next i
    
   ' Dim rsdata2 As New Recordset
    'With rsdata2.Fields
       ' .Append "user_id", adBSTR
        '.Append "total_amount", adBigInt
    'End With
   ' rsdata2.Open
    'rsTemp.MoveFirst
    'For i = 1 To rsTemp.RecordCount
        'rsdata2.AddNew
        'rsdata2!user_id = FormatDbValue(rsTemp!user_id)
       ' rsdata2!total_amount = FormatDbValue(rsTemp!total_amount)
       ' rsTemp.MoveNext
        'rsdata2.Update
    'Next i
    
   ' Me.Hide
    'Set frmTemp = New frmChart
    'frmTemp.ClearChart
    'frmTemp.AddChart "人数", rsData
    'frmTemp.AddChart "金额", rsdata2
    'frmTemp.ShowChart "售票员售票简报"
    'Set frmTemp = Nothing
    'Unload Me

   ' Exit Sub
'Error_Handle:
   ' Set frmTemp = Nothing
   ' ShowErrorMsg
    
End Sub

Private Sub cmdok_Click()
    Dim oSellerStat As New TicketSellerDim
    Dim aszUserID() As String
    Dim nSelUserCount As Integer
    Dim szSellerSation As String
    Dim i As Integer
    
    On Error GoTo Error_Handle
'   szSellerSation = Trim(cboSellStation.Text)
    'If cboSellStation.ListIndex = 0 Then
   ' szSellerSation = "bz"
    'ElseIf cboSellStation.ListIndex = 1 Then
    '
    'szSellerSation = "xz"
   ' ElseIf cboSellStation.ListIndex = 2 Then
   ' szSellerSation = "bz+xz"
   ' Else
    'szSellerSation = ""
   ' End If
    
    
    '生成Recordset
   ' nSelUserCount = ArrayLength(adUser.RightData)
'    dtpBeginDate.Value = CDate(Year(dtpBeginDate.Value) & "-" & Month(dtpBeginDate.Value) & "-01")
'    dtpEndDate.Value = DateAdd("D", -1, DateAdd("M", 1, dtpBeginDate.Value))
  '  If nSelUserCount > 0 Then
       oSellerStat.Init m_oActiveUser
        'ReDim aszUserID(1 To nSelUserCount)
        'For i = 1 To nSelUserCount
         '   aszUserID(i) = ResolveDisplay(adUser.RightData(i))
        'Next
        'Set m_rsData = oSellerStat.SellerDateStat(aszUserID, dtpBeginDate.Value, dtpEndDate.Value)
       'StationDateStat
        Set m_rsData = oSellerStat.StationDateStat(IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)), dtpBeginDate.Value, dtpEndDate.Value)
   ' End If
    ReDim m_vaCustomData(1 To 3, 1 To 2)
    m_vaCustomData(1, 1) = "统计开始日期"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY年MM月DD日")
    
    m_vaCustomData(2, 1) = "统计结束日期"
    m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")
    
    m_vaCustomData(3, 1) = "制表人"
    m_vaCustomData(3, 2) = m_oActiveUser.UserID
    
    
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
    'm_bOk = False

'    dtpBeginDate.Value = DateAdd("d", -1, m_oParam.NowDate)
'    dtpEndDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    '设置为上个月的一号到31
    Dim dyNow As Date
    dyNow = m_oParam.NowDate
    dtpBeginDate.Value = Format(DateAdd("m", -1, dyNow), "yyyy-mm-01")
    dtpEndDate.Value = DateAdd("d", -1, Format(dyNow, "yyyy-mm-01"))
   FillSellStation cboSellStation
   
    If Trim(m_oActiveUser.SellStationID) <> "" Then
        txtSellStationID.Enabled = False
    End If
   ' FillSellerEx
    
   ' EnableOK
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

'Private Sub FillSellStation()
'    '判断用户属于哪个上车站,如果为空则填充一个空行,再填充所有的上车站
'
'    '否则只填充用户所属的上车站
'
 


'
'End Sub



