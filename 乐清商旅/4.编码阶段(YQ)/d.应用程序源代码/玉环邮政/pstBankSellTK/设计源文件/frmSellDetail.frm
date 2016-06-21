VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmSellDetail 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "售票明细查询"
   ClientHeight    =   5820
   ClientLeft      =   3375
   ClientTop       =   1935
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   6780
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtTicketID 
      Height          =   300
      Left            =   1320
      TabIndex        =   15
      Top             =   1320
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   -15
      ScaleHeight     =   705
      ScaleWidth      =   7665
      TabIndex        =   6
      Top             =   -45
      Width           =   7665
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择查询条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   7
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -75
      TabIndex        =   5
      Top             =   645
      Width           =   7725
   End
   Begin VB.Frame Frame1 
      Caption         =   "报表说明"
      Height          =   555
      Left            =   1185
      TabIndex        =   3
      Top             =   6165
      Width           =   6975
      Begin VB.Label Label3 
         Caption         =   "按票种指定时间段，统计票种人数、金额。用于统计售票员的售票情况。"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   6435
      End
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   2595
      TabIndex        =   0
      Top             =   5325
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
      MICON           =   "frmSellDetail.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   315
      Left            =   3975
      TabIndex        =   1
      Top             =   5325
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
      MICON           =   "frmSellDetail.frx":001C
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
      Left            =   5385
      TabIndex        =   2
      Top             =   5325
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
      MICON           =   "frmSellDetail.frx":0038
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
      Left            =   4485
      TabIndex        =   8
      Top             =   900
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      Format          =   105447424
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Top             =   900
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      Format          =   105447424
      CurrentDate     =   36572
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " RTStation"
      Enabled         =   0   'False
      Height          =   3120
      Left            =   -135
      TabIndex        =   10
      Top             =   4845
      Width           =   8745
   End
   Begin PSTBankSellTK.AddDel adUser 
      Height          =   3135
      Left            =   1440
      TabIndex        =   16
      Top             =   1680
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5530
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LeftLabel       =   "待选列表(&L)"
      RightLabel      =   "已选列表(&R)"
      ButtonWidth     =   1215
      ButtonHeight    =   315
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "票号:"
      Height          =   180
      Left            =   720
      TabIndex        =   14
      Top             =   1320
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "网点代售点代码:"
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   3375
      TabIndex        =   12
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   1080
   End
End
Attribute VB_Name = "frmSellDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Implements IConditionForm

Const cszFileName = "售票员明细查询模板.xls"

Public m_bOk As Boolean
Private m_rsData As Recordset
Private m_vaCustomData As Variant

Private Sub adUser_DataChange()
    EnableOK
End Sub

Private Sub cboSellStation_Change()
   FillSellerEx
End Sub
'Public Sub FillSellerEx()
''    Dim oUnit As New Unit
'    Dim aszUser() As String
'    Dim aszUser2() As String
'    Dim aszUser22() As String
'
''    Dim oUser As New User
'    Dim i As Integer, nUserCount As Integer
'    Dim szRecentSeller As String
'    Dim szTemp As String
'
'    Dim nNoSelected As Integer, nSelected As Integer
'
'
''    aszUser = oUnit.GetAllUserEX(, ResolveDisplay(cboSellStation))
'    nUserCount = ArrayLength(aszUser)
'    If nUserCount > 0 Then
'
''        oUser.Init m_oActiveUser
''        szRecentSeller = GetRecentSeller()
'
'        nNoSelected = 0
'        nSelected = 0
'
'        For i = 1 To nUserCount
''            oUser.Identify aszUser(i)
'            szTemp = MakeDisplayString(aszUser(i, 1), aszUser(i, 2))
'            If InStr(1, szRecentSeller, szTemp, vbTextCompare) = 0 Then
'                nNoSelected = nNoSelected + 1
'                ReDim Preserve aszUser2(1 To nNoSelected)
'                aszUser2(nNoSelected) = szTemp
'            Else
'                nSelected = nSelected + 1
'                ReDim Preserve aszUser22(1 To nSelected)
'                aszUser22(nSelected) = szTemp
'            End If
'        Next
'    End If
''    adUser.LeftData = aszUser2
''    adUser.RightData = aszUser22
'End Sub
Private Sub cboSellStation_Click()
    cboSellStation_Change
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOk_Click()
'    Dim oSellerStat As New TicketSellerDim
    Dim aszUserID() As String
    Dim nSelUserCount As Integer
    Dim i As Integer

    On Error GoTo Error_Handle
    '生成Recordset
    
    Set m_rsData = SellDetail(dtpBeginDate.Value, dtpEndDate.Value)
    
    
    
    ReDim m_vaCustomData(1 To 3, 1 To 2)
    m_vaCustomData(1, 1) = "统计开始日期"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY年MM月DD日")

    m_vaCustomData(2, 1) = "统计结束日期"
    m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")

    m_vaCustomData(3, 1) = "制表人"
    m_vaCustomData(3, 2) = m_cszOperatorID


'    SaveRecentSeller adUser.RightData
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
    m_bOk = False

'    dtpBeginDate.Value = DateAdd("d", -1, m_oParam.NowDate)
'    dtpEndDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    '设置为上个月的一号到31
    Dim dyNow As Date
    dyNow = Date
    dtpBeginDate.Value = dyNow ' DateAdd("d", -1, dyNow)
    dtpEndDate.Value = dyNow 'DateAdd("d", -1, dyNow)
    FillUser
'    FillSellStation cboSellStation
'    FillSellerEx
    
    EnableOK
End Sub

Private Sub FillUser()
FillSellerEx
End Sub

Public Sub FillSellerEx()

    Dim aszUser() As String
    Dim aszUser2() As String
    Dim aszUser22() As String
    

    Dim i As Integer, nUserCount As Integer
    Dim szRecentSeller As String
    Dim szTemp As String
    
    Dim nNoSelected As Integer, nSelected As Integer
    
    Dim odb As New ADODB.Connection
    Dim rsTemp As Recordset
    Dim szSql As String
    
    odb.ConnectionString = GetConnectionStr
    odb.CursorLocation = adUseClient
    odb.Open

    szSql = "select operatorid  as bank_id from tickets group by operatorid  "
    
    Set rsTemp = odb.Execute(szSql)

    If rsTemp.RecordCount > 0 Then
        ReDim aszUser(1 To rsTemp.RecordCount)
        For i = 1 To rsTemp.RecordCount
            aszUser(i) = FormatDbValue(rsTemp!bank_id)
            rsTemp.MoveNext
        Next
        
    End If

   
    adUser.LeftData = aszUser
End Sub
'填充售票员
Private Sub FillSeller()
'    Dim oSysMan As New SystemMan
'    Dim auiUserInfo() As TUserInfo
'    Dim i As Integer, nUserCount As Integer
'    Dim aszTemp() As String, aszTemp2() As String
'    Dim nNoSelected As Integer, nSelected As Integer
'    Dim szTemp As String
'    Dim szRecentSeller As String
'
'    oSysMan.Init m_oActiveUser
'    auiUserInfo = oSysMan.GetAllUser()
'    nUserCount = ArrayLength(auiUserInfo)
'    If nUserCount > 0 Then
'        szRecentSeller = GetRecentSeller()
'        nNoSelected = 0
'        nSelected = 0
'        For i = 1 To nUserCount
'            szTemp = MakeDisplayString(auiUserInfo(i).UserID, auiUserInfo(i).UserName)
'            If InStr(1, szRecentSeller, szTemp, vbTextCompare) = 0 Then
'                nNoSelected = nNoSelected + 1
'                ReDim Preserve aszTemp(1 To nNoSelected)
'                aszTemp(nNoSelected) = szTemp
'            Else
'                nSelected = nSelected + 1
'                ReDim Preserve aszTemp2(1 To nSelected)
'                aszTemp2(nSelected) = szTemp
'            End If
'        Next
'    End If
'    adUser.LeftData = aszTemp
'    adUser.RightData = aszTemp2
End Sub

Private Sub EnableOK()
'    Dim nCount As Integer
'    nCount = ArrayLength(adUser.RightData)
'    cmdOk.Enabled = IIf(nCount > 0, True, False)
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
'End Sub


Private Function SellDetail(pdyStartDate As Date, pdyEndDate As Date) As Recordset
    '汇总得到银行网点的代售信息
    Dim odb As New ADODB.Connection
    Dim rsTemp As Recordset
    Dim szSql As String
    odb.ConnectionString = GetConnectionStr
    odb.CursorLocation = adUseClient
    odb.Open
    
'    szSql = " SELECT u.bank_id + '[' + max(u.bank_name) + ']' bank ,  FROM tickets t , user_info u  " _
        & " WHERE t.operatorid = u.operatorid " _
        & " AND t.selldate >= " & TransFieldValueToString(pdyStartDate) _
        & " AND t.selldate < " & TransFieldValueToString(DateAdd("d", 1, pdyEndDate)) _
        & " GROUP BY u.bank_id"
    Dim szWhere As String
    Dim szOperatorID As String
    Dim aszUserID() As String
    Dim i, nSelUserCount As Integer
    Dim pszUser As String
    nSelUserCount = ArrayLength(adUser.RightData)
    
    If nSelUserCount > 0 Then
        For i = 1 To nSelUserCount
            pszUser = pszUser & "'" & adUser.RightData(i) & "',"
        Next
        pszUser = Left(pszUser, Len(pszUser) - 1)
        szWhere = " AND t.operatorid in(" & pszUser & ") "
    End If

    If txtTicketID.Text <> "" Then
        szWhere = " AND t.id='" & Trim(txtTicketID.Text) & "' "
    End If
    szSql = "select ID,scheduleid,StartTime,StartID,DestID,SeatID,StartName,DestName,SellDate,StartNo,price,t.OperatorID,TicketType,case t.Status when '2' then '作废' else '售出' end as Status from tickets t where " _
        & "  t.selldate >= " & TransFieldValueToString(pdyStartDate) _
        & " AND t.selldate < " & TransFieldValueToString(DateAdd("d", 1, pdyEndDate)) _
        & szWhere
    
    Set rsTemp = odb.Execute(szSql)
    Set SellDetail = rsTemp
    
End Function



