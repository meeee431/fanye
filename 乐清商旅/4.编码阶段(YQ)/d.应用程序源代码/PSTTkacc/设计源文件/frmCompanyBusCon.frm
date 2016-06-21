VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmCompanyBusCon 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "车次营收简报"
   ClientHeight    =   4380
   ClientLeft      =   4530
   ClientTop       =   2040
   ClientWidth     =   6540
   HelpContextID   =   6000030
   Icon            =   "frmCompanyBusCon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtSellStationID 
      Height          =   300
      Left            =   1290
      TabIndex        =   24
      Top             =   3090
      Width           =   1605
   End
   Begin RTComctl3.CoolButton cmdChart 
      Height          =   315
      Left            =   510
      TabIndex        =   23
      Top             =   3900
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "图表"
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
      MICON           =   "frmCompanyBusCon.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   345
      Left            =   2250
      TabIndex        =   21
      Top             =   3900
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
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
      MICON           =   "frmCompanyBusCon.frx":0028
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
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -60
      TabIndex        =   10
      Top             =   660
      Width           =   6885
   End
   Begin VB.ComboBox cboBusSection 
      Height          =   300
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2040
      Width           =   1905
   End
   Begin VB.OptionButton optCombine 
      BackColor       =   &H00E0E0E0&
      Caption         =   "按车次段汇总"
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   990
      Width           =   1455
   End
   Begin VB.OptionButton optCompany 
      BackColor       =   &H00E0E0E0&
      Caption         =   "按拆帐公司汇总"
      Height          =   285
      Left            =   3120
      TabIndex        =   7
      Top             =   1530
      Width           =   1605
   End
   Begin VB.ComboBox cboExtraStatus 
      Height          =   300
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2085
      Width           =   1635
   End
   Begin VB.TextBox txtLike 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4410
      TabIndex        =   5
      Top             =   2595
      Width           =   1905
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   -30
      ScaleHeight     =   735
      ScaleWidth      =   6615
      TabIndex        =   3
      Top             =   -30
      Width           =   6615
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   4
         Top             =   270
         Width           =   990
      End
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   5130
      TabIndex        =   1
      Top             =   3900
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
      MICON           =   "frmCompanyBusCon.frx":0044
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
      Height          =   345
      Left            =   3690
      TabIndex        =   0
      Top             =   3900
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
      MICON           =   "frmCompanyBusCon.frx":0060
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
      Height          =   300
      Left            =   1290
      TabIndex        =   11
      Top             =   1515
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   300
      Left            =   1290
      TabIndex        =   12
      Top             =   960
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin FText.asFlatTextBox txtTransportCompanyID 
      Height          =   300
      Left            =   4440
      TabIndex        =   13
      Top             =   2040
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      ButtonVisible   =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   3120
      Left            =   -150
      TabIndex        =   2
      Top             =   3660
      Width           =   8745
   End
   Begin FText.asFlatTextBox txtSellStation 
      Height          =   300
      Left            =   1290
      TabIndex        =   22
      ToolTipText     =   "请点...进行选择"
      Top             =   2595
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonHotBackColor=   -2147483633
      Locked          =   -1  'True
      ButtonPressedBackColor=   -2147483627
      Text            =   ""
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&T):"
      Height          =   180
      Left            =   210
      TabIndex        =   20
      Top             =   2655
      Width           =   900
   End
   Begin VB.Label lblCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参运公司(T):"
      Height          =   180
      Left            =   3120
      TabIndex        =   19
      Top             =   2130
      Width           =   1080
   End
   Begin VB.Label lblCombine 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次段序号(&R):"
      Height          =   180
      Left            =   3120
      TabIndex        =   18
      Top             =   2115
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "补票状态(&S):"
      Height          =   180
      Left            =   210
      TabIndex        =   17
      Top             =   2115
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   210
      TabIndex        =   16
      Top             =   1575
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   210
      TabIndex        =   15
      Top             =   1020
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "模糊车次(&A):"
      Height          =   180
      Left            =   3120
      TabIndex        =   14
      Top             =   2655
      Width           =   1080
   End
End
Attribute VB_Name = "frmCompanyBusCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IConditionForm
Const cszFileName = "车次售票营收简报模板.xls"


Public m_bOk As Boolean
Public m_nMode As EBusStatMode
Private m_rsData As Recordset
Private m_vaCustomData As Variant

Dim m_aszTemp() As String
Dim oDss As New TicketBusDim

Dim m_szCompanyCode As String

Dim m_szSellStationID As String


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChart_Click()
    
    On Error GoTo Error_Handle
    '生成记录集
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim szSellStation As String
    Dim frmTemp As frmChart
    '设置上车站的值
    If Trim(txtSellStation.Text) = "" Then
        txtSellStation.Text = Trim(m_oActiveUser.SellStationID)
        If txtSellStation.Text <> "" Then
            szSellStation = "'" & m_oActiveUser.SellStationID & "'"
        End If
    Else
        szSellStation = m_szSellStationID
    End If
    If txtSellStationID.Text <> "" Then
        szSellStation = txtSellStationID.Text
    End If
    If m_nMode = ST_BySalerStationAndSaleTime Then
        If optCompany.Value Then
            Set rsTemp = oDss.GetBusStatBySaleTime(m_szCompanyCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), txtLike.Text, szSellStation)
        Else
            Set rsTemp = oDss.GetCombineBusSimply(dtpBeginDate.Value, dtpEndDate.Value, Val(cboBusSection.Text), ResolveDisplay(cboExtraStatus.Text), txtLike.Text, szSellStation)
        End If
    ElseIf m_nMode = ST_ByBusStationAndBusDate Then
    
        If optCompany.Value Then
            Set rsTemp = oDss.GetBusStatByBusDate(m_szCompanyCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), txtLike.Text, szSellStation)
        Else
            Set rsTemp = oDss.GetCombineBusSimplyByBusDate(dtpBeginDate.Value, dtpEndDate.Value, Val(cboBusSection.Text), ResolveDisplay(cboExtraStatus.Text), txtLike.Text, szSellStation)
        End If
    ElseIf m_nMode = ST_BySalerStationAndBusDate Then
        If optCompany.Value Then
            Set rsTemp = oDss.GetBusStatByBusDateAndSalerStation(m_szCompanyCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), txtLike.Text, szSellStation)
        Else
            Set rsTemp = oDss.GetBusStatByBusDateAndSalerStation(m_szCompanyCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), txtLike.Text, szSellStation)
        End If
    
    End If
    
    Dim rsData As New Recordset
    With rsData.Fields
        .Append "bus_id", adBSTR
        .Append "passenger_number", adBigInt
    End With
    rsData.Open
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        rsData.AddNew
        rsData!bus_id = FormatDbValue(rsTemp!bus_id)
        rsData!passenger_number = FormatDbValue(rsTemp!passenger_number)
        rsTemp.MoveNext
        rsData.Update
    Next i
    
    Dim rsdata2 As New Recordset
    With rsdata2.Fields
        .Append "bus_id", adBSTR
        .Append "total_ticket_price", adBigInt
    End With
    rsdata2.Open
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        rsdata2.AddNew
        rsdata2!bus_id = FormatDbValue(rsTemp!bus_id)
        rsdata2!total_ticket_price = FormatDbValue(rsTemp!total_ticket_price)
        rsTemp.MoveNext
        rsdata2.Update
    Next i
    
    Me.Hide
    Set frmTemp = New frmChart
    frmTemp.ClearChart
    frmTemp.AddChart "人数", rsData
    frmTemp.AddChart "金额", rsdata2
    frmTemp.ShowChart "车次售票营收简报"
    Set frmTemp = Nothing
    Unload Me

    Exit Sub
Error_Handle:
    Set frmTemp = Nothing
    ShowErrorMsg
    
End Sub

Private Sub cmdok_Click()
    On Error GoTo Error_Handle
    '生成记录集
    Dim rsTemp As Recordset

    Dim rsData As New Recordset
    Dim i As Integer
    Dim szSellStation As String
    '设置上车站的值
    If Trim(txtSellStation.Text) = "" Then
        txtSellStation.Text = Trim(m_oActiveUser.SellStationID)
        If txtSellStation.Text <> "" Then
            szSellStation = "'" & m_oActiveUser.SellStationID & "'"
        End If
    Else
        szSellStation = m_szSellStationID
    End If
    If txtSellStationID.Text <> "" Then
        szSellStation = txtSellStationID.Text
    End If
    If m_nMode = ST_BySalerStationAndSaleTime Then
        If optCompany.Value Then
            Set rsTemp = oDss.GetBusStatBySaleTime(m_szCompanyCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), txtLike.Text, szSellStation)
        Else
            Set rsTemp = oDss.GetCombineBusSimply(dtpBeginDate.Value, dtpEndDate.Value, Val(cboBusSection.Text), ResolveDisplay(cboExtraStatus.Text), txtLike.Text, szSellStation)
        End If
    ElseIf m_nMode = ST_ByBusStationAndBusDate Then
    
        If optCompany.Value Then
            Set rsTemp = oDss.GetBusStatByBusDate(m_szCompanyCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), txtLike.Text, szSellStation)
        Else
            Set rsTemp = oDss.GetCombineBusSimplyByBusDate(dtpBeginDate.Value, dtpEndDate.Value, Val(cboBusSection.Text), ResolveDisplay(cboExtraStatus.Text), txtLike.Text, szSellStation)
        End If
    ElseIf m_nMode = ST_BySalerStationAndBusDate Then
        If optCompany.Value Then
            Set rsTemp = oDss.GetBusStatByBusDateAndSalerStation(m_szCompanyCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), txtLike.Text, szSellStation)
        Else
            Set rsTemp = oDss.GetBusStatByBusDateAndSalerStation(m_szCompanyCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), txtLike.Text, szSellStation)
        End If
    
    End If
    Set m_rsData = rsTemp
    
    ReDim m_vaCustomData(1 To 7, 1 To 2)
    
    m_vaCustomData(1, 1) = "统计开始日期"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY年MM月DD日")
    m_vaCustomData(2, 1) = "统计结束日期"
    m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")
    m_vaCustomData(3, 1) = "补票状态"
    m_vaCustomData(3, 2) = cboExtraStatus.Text
    m_vaCustomData(4, 1) = "统计方式"
    m_vaCustomData(4, 2) = GetStatName(m_nMode)
    m_vaCustomData(5, 1) = "拆帐公司"
    m_vaCustomData(5, 2) = IIf((txtTransportCompanyID.Text <> ""), txtTransportCompanyID.Text, "全部公司")
    m_vaCustomData(6, 1) = "上车站"
    m_vaCustomData(6, 2) = IIf((szSellStation <> ""), szSellStation, "全部上车站")
    
    m_vaCustomData(7, 1) = "制表人"
    m_vaCustomData(7, 2) = m_oActiveUser.UserID
    
    m_bOk = True
    Unload Me
    Exit Sub
Error_Handle:
    ShowErrorMsg
End Sub


Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_Load()
    oDss.Init m_oActiveUser
    m_szCompanyCode = ""
    m_szSellStationID = ""
    m_bOk = False
    FillCombine
'    dtpBeginDate.Value = DateAdd("d", -1, m_oParam.NowDate)
'    dtpEndDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    '设置为上个月的一号到31
    Dim dyNow As Date
    dyNow = m_oParam.NowDate
    dtpBeginDate.Value = Format(DateAdd("m", -1, dyNow), "yyyy-mm-01")
    dtpEndDate.Value = DateAdd("d", -1, Format(dyNow, "yyyy-mm-01"))
    cboExtraStatus.AddItem "1[售票]"
    cboExtraStatus.AddItem "2[补票]"
    cboExtraStatus.AddItem "3[售票+补票]"
    
    cboExtraStatus.ListIndex = 2
    
    optCompany.Value = True
    SetVisible False
'    FillSellStation cboSellStation

    Me.Caption = "车次营收简报[" & GetStatName(m_nMode) & "]"
    
    
    
    If m_nMode = ST_BySalerStationAndSaleTime Then
        lblCaption = "请输入售票的起止日期:"
    ElseIf m_nMode = ST_ByBusStationAndBusDate Then
        Me.Caption = "车次营收简报[按车次日期汇总]"
    Else
        Me.Caption = "车次营收简报[按车次日期汇总]"
    End If
    
    If Trim(m_oActiveUser.SellStationID) <> "" Then
        txtSellStationID.Enabled = False
    End If
    
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



Private Sub optCombine_Click()
    SetVisible True
End Sub

Private Sub optCompany_Click()
    SetVisible False
    
End Sub

Private Sub txtSellStation_ButtonClick()
    Dim aszTemp() As String
    Dim nCount As Integer
    On Error GoTo ErrorHandle
    aszTemp = m_oShell.SelectSellStation(m_oActiveUser.SellStationID, , True)
    txtSellStation.Text = TeamToString(aszTemp, 2, False)
    
    m_szSellStationID = TeamToString(aszTemp, 1, False)
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub
Private Sub txtTransportCompanyID_ButtonClick()
    Dim aszTransportCompanyID() As String
    aszTransportCompanyID = m_oShell.SelectCompany
    txtTransportCompanyID.Text = TeamToString(aszTransportCompanyID, 2, False)
    
    m_szCompanyCode = TeamToString(aszTransportCompanyID, 1)
    
End Sub

Private Sub FillCombine()
    '填充唯一的车次序号
    Dim aszTemp() As String
    Dim i As Integer
    Dim nCount As Integer
    Dim oCompanyDim As New TicketCompanyDim
    oCompanyDim.Init m_oActiveUser
    aszTemp = oCompanyDim.GetUniqueCombine
    nCount = ArrayLength(aszTemp)
    For i = 1 To nCount
        cboBusSection.AddItem aszTemp(i)
    Next i
    If cboBusSection.ListCount > 0 Then cboBusSection.ListIndex = 0
    Set oCompanyDim = Nothing
End Sub


Private Sub SetVisible(pbVisible As Boolean)
    lblCombine.Visible = pbVisible
    cboBusSection.Visible = pbVisible
    lblCompany.Visible = Not pbVisible
    txtTransportCompanyID.Visible = Not pbVisible
    
End Sub




'Private Sub FillSellStation()
'    '判断用户属于哪个上车站,如果为空则填充一个空行,再填充所有的上车站
'
'    '否则只填充用户所属的上车站
'End Sub


