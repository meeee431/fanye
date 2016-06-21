VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmBusStationCon 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "车次各途经站营收简报"
   ClientHeight    =   4035
   ClientLeft      =   3555
   ClientTop       =   3045
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   6570
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtSellStationID 
      Height          =   300
      Left            =   4440
      TabIndex        =   23
      Top             =   2670
      Width           =   1635
   End
   Begin FText.asFlatTextBox txtSellStation 
      Height          =   300
      Left            =   1410
      TabIndex        =   22
      ToolTipText     =   "请点...进行选择"
      Top             =   2145
      Width           =   1620
      _ExtentX        =   2858
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
   Begin VB.TextBox txtLike 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1410
      TabIndex        =   20
      Top             =   2670
      Width           =   1905
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -30
      TabIndex        =   7
      Top             =   660
      Width           =   6885
   End
   Begin VB.ComboBox cboExtraStatus 
      Height          =   300
      Left            =   1470
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2955
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.OptionButton optCompany 
      BackColor       =   &H00E0E0E0&
      Caption         =   "按拆帐公司汇总"
      Height          =   285
      Left            =   3150
      TabIndex        =   3
      Top             =   1650
      Width           =   1605
   End
   Begin VB.OptionButton optCombine 
      BackColor       =   &H00E0E0E0&
      Caption         =   "按车次段汇总"
      Height          =   255
      Left            =   3150
      TabIndex        =   2
      Top             =   1110
      Width           =   1455
   End
   Begin VB.ComboBox cboBusSection 
      Height          =   300
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2190
      Width           =   1905
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   345
      Left            =   2190
      TabIndex        =   0
      Top             =   3570
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
      MICON           =   "frmBusStationCon.frx":0000
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
      Left            =   5070
      TabIndex        =   5
      Top             =   3570
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
      MICON           =   "frmBusStationCon.frx":001C
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
      Left            =   3630
      TabIndex        =   6
      Top             =   3570
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
      MICON           =   "frmBusStationCon.frx":0038
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
      Left            =   1410
      TabIndex        =   8
      Top             =   1635
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   300
      Left            =   1410
      TabIndex        =   9
      Top             =   1080
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
      TabIndex        =   10
      Top             =   2190
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   -30
      ScaleHeight     =   735
      ScaleWidth      =   6615
      TabIndex        =   12
      Top             =   -30
      Width           =   6615
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   13
         Top             =   270
         Width           =   990
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   3120
      Left            =   -120
      TabIndex        =   11
      Top             =   3330
      Width           =   8745
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "模糊车次(&A):"
      Height          =   180
      Left            =   240
      TabIndex        =   21
      Top             =   2730
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   240
      TabIndex        =   19
      Top             =   1140
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   240
      TabIndex        =   18
      Top             =   1695
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "补票状态(&S):"
      Height          =   180
      Left            =   390
      TabIndex        =   17
      Top             =   3045
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label lblCombine 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次段序号(&R):"
      Height          =   180
      Left            =   3150
      TabIndex        =   16
      Top             =   2235
      Width           =   1260
   End
   Begin VB.Label lblCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参运公司(T):"
      Height          =   180
      Left            =   3150
      TabIndex        =   15
      Top             =   2250
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&T):"
      Height          =   180
      Left            =   240
      TabIndex        =   14
      Top             =   2220
      Width           =   900
   End
End
Attribute VB_Name = "frmBusStationCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IConditionForm
Const cszFileName = "车次途经站售票简报模板.xls"


Public m_bOk As Boolean
Public m_bBySeller As Boolean '是否按照售票员所属车站统计
Private m_rsData As Recordset
Private m_vaCustomData As Variant

Dim m_aszTemp() As String
Dim oDss As New TicketBusDim

Dim m_szCompanyCode As String
Dim m_szSellStationID As String



Private Sub cmdCancel_Click()
    Unload Me
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
    '得到接口数据
    If m_bBySeller Then
        Set rsTemp = oDss.GetBusStationStatByBusDateAndSeller(m_szCompanyCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), txtLike.Text, szSellStation)
    Else
        Set rsTemp = oDss.GetBusStationStatByBusDate(m_szCompanyCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), txtLike.Text, szSellStation)
        
    End If
    
    MakeRecordSet rsTemp
'
    ReDim m_vaCustomData(1 To 7, 1 To 2)
    
    m_vaCustomData(1, 1) = "统计开始日期"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY年MM月DD日")
    m_vaCustomData(2, 1) = "统计结束日期"
    m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")
    m_vaCustomData(3, 1) = "补票状态"
    m_vaCustomData(3, 2) = cboExtraStatus.Text
    m_vaCustomData(4, 1) = "统计方式"
    m_vaCustomData(4, 2) = IIf(m_bBySeller, "按车次日期及售票员所属车站汇总", "按车次日期及车次上车站汇总")
    m_vaCustomData(5, 1) = "拆帐公司"
    m_vaCustomData(5, 2) = IIf(txtTransportCompanyID.Text <> "", txtTransportCompanyID.Text, "所有公司")
    
    m_vaCustomData(6, 1) = "上车站"
    m_vaCustomData(6, 2) = IIf((szSellStation <> ""), szSellStation, "全部上车站")
    
    
    m_vaCustomData(7, 1) = "制表人"
    m_vaCustomData(7, 2) = m_oActiveUser.UserID
    
    If rsTemp.RecordCount = 0 Then
        m_bOk = False
    Else
        m_bOk = True
    End If
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
'    FillSellStation txtSellStation.Text

    If Trim(m_oActiveUser.SellStationID) <> "" Then
        txtSellStationID.Enabled = False
    End If

    
    lblCaption = "请输入车次的起止日期:"
    If m_bBySeller Then
        Me.Caption = "车次途经站营收简报[按车次日期及售票员所属车站汇总]"
    Else
        Me.Caption = "车次途经站营收简报[按车次日期及车次上车站汇总]"
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


Private Sub MakeRecordSet(prsData As Recordset)
    '手工生成记录集
    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer
    
    '暂放数据
    Dim szbusID As String
    Dim szStationId As String
    Dim alNumber(TP_TicketTypeCount) As Long '各种票种的张数
    Dim adbAmount(TP_TicketTypeCount) As Double  '各种票种的金额
    Dim szRouteName As String
    Dim szEndStationName As String
    Dim szBusStartTime As String
    Dim szTransportCompanyName As String
    Dim lReturnNumber As Long
    Dim dbReturnAmount As Double
    Dim dbReturnCharge As Double
    Dim lChangeNumber As Long
    Dim dbChangeAmount As Double
    Dim dbChangeCharge As Double
    Dim lCancelNumber As Long '废票张数
    Dim dbCancelAmount As Double '退票总额
    Dim dbTotalPrice As Double '总额
    Dim lTotalNumber As Long '总张数
    Dim dbTotalTicketPrice As Double '加上退票手续费的总金额
    Dim szStation As String
    
    
    Dim szStationName As String '途经站及其人数
    Dim lStationNumber As Long
    
    
    With rsTemp.Fields
        .Append "bus_id", adChar, 5
        .Append "route_name", adChar, 16
        .Append "end_station_name", adVarChar, 10 '终点站名
        .Append "bus_start_time", adChar, 5
        .Append "transport_company_name", adVarChar, 10
        For i = 1 To TP_TicketTypeCount
            .Append "number_ticket_type" & i, adInteger
            .Append "amount_ticket_type" & i, adCurrency
        Next i
        .Append "return_number", adBigInt
        .Append "return_amount", adCurrency
        .Append "return_charge", adCurrency
        .Append "change_number", adBigInt
        .Append "change_amount", adCurrency
        .Append "change_charge", adCurrency
        .Append "cancel_number", adBigInt
        .Append "cancel_amount", adCurrency
        .Append "total_number", adBigInt
        .Append "total_price", adCurrency
        .Append "total_ticket_price", adCurrency
        .Append "station", adVarChar, 255 '途经站及其张数
        
        
    End With
    rsTemp.Open
    If prsData Is Nothing Then Exit Sub
    If prsData.RecordCount = 0 Then Exit Sub
    prsData.MoveFirst
    szbusID = FormatDbValue(prsData!bus_id)
    szStationId = FormatDbValue(prsData!station_id)
    szStationName = FormatDbValue(prsData!station_name)
'    lStationNumber = FormatDbValue(prsData!passenger_number)
    
    szRouteName = FormatDbValue(prsData!route_name)
    szEndStationName = FormatDbValue(prsData!end_station_name)
    szBusStartTime = FormatDbValue(prsData!bus_start_time)
    szTransportCompanyName = FormatDbValue(prsData!transport_company_short_name)
    
    
        
    For i = 1 To prsData.RecordCount
        If szbusID <> FormatDbValue(prsData!bus_id) Then
            '赋予记录集
            
            rsTemp.AddNew
            rsTemp!bus_id = szbusID
            rsTemp!route_name = szRouteName
            rsTemp!end_station_name = szEndStationName
            rsTemp!bus_start_time = szBusStartTime
            rsTemp!transport_company_name = szTransportCompanyName
            For j = 1 To TP_TicketTypeCount
                rsTemp.Fields("number_ticket_type" & j) = alNumber(j)
                rsTemp.Fields("amount_ticket_type" & j) = adbAmount(j)
            Next j
            rsTemp!return_number = lReturnNumber
            rsTemp!return_amount = dbReturnAmount
            rsTemp!return_charge = dbReturnCharge
            rsTemp!change_number = lChangeNumber
            rsTemp!change_amount = dbChangeAmount
            rsTemp!change_charge = dbChangeCharge
            rsTemp!cancel_number = lCancelNumber
            rsTemp!cancel_amount = dbCancelAmount
            rsTemp!total_number = lTotalNumber
            rsTemp!total_price = dbTotalPrice
            rsTemp!total_ticket_price = dbTotalTicketPrice
            szStation = szStation & szStationName & lStationNumber
            
            rsTemp!Station = szStation
            rsTemp.Update
            
            '清空原值
            For j = 1 To TP_TicketTypeCount
                alNumber(j) = 0
                adbAmount(j) = 0
            Next j
            lReturnNumber = 0
            dbReturnAmount = 0
            dbReturnCharge = 0
            lChangeNumber = 0
            dbChangeAmount = 0
            dbChangeCharge = 0
            lCancelNumber = 0
            dbCancelAmount = 0
            lTotalNumber = 0
            dbTotalPrice = 0
            dbTotalTicketPrice = 0
            szStation = ""
            
            '赋该车次的初始值
            
            szbusID = FormatDbValue(prsData!bus_id)
            szRouteName = FormatDbValue(prsData!route_name)
            szEndStationName = FormatDbValue(prsData!end_station_name)
            szBusStartTime = FormatDbValue(prsData!bus_start_time)
            szTransportCompanyName = FormatDbValue(prsData!transport_company_short_name)
            
            
            szStationName = FormatDbValue(prsData!station_name)
            szStationId = FormatDbValue(prsData!station_id)
            lStationNumber = FormatDbValue(prsData!passenger_number)
            
            alNumber(prsData!ticket_type) = alNumber(prsData!ticket_type) + FormatDbValue(prsData!passenger_number2)
            adbAmount(prsData!ticket_type) = adbAmount(prsData!ticket_type) + FormatDbValue(prsData!ticket_price2)
            lReturnNumber = lReturnNumber + FormatDbValue(prsData!ticket_return_number)
            dbReturnAmount = dbReturnAmount + FormatDbValue(prsData!ticket_return_amount)
            dbReturnCharge = dbReturnCharge + FormatDbValue(prsData!ticket_return_charge)
            lChangeNumber = lChangeNumber + FormatDbValue(prsData!ticket_change_number)
            dbChangeAmount = dbChangeAmount + FormatDbValue(prsData!ticket_change_charge)
            lCancelNumber = lCancelNumber + FormatDbValue(prsData!ticket_cancel_number)
            dbCancelAmount = dbCancelAmount + FormatDbValue(prsData!ticket_cancel_amount)
            lTotalNumber = lTotalNumber + FormatDbValue(prsData!passenger_number)
            dbTotalPrice = dbTotalPrice + FormatDbValue(prsData!ticket_price)
            dbTotalTicketPrice = dbTotalTicketPrice + FormatDbValue(prsData!total_ticket_price)
            
        ElseIf szStationId <> FormatDbValue(prsData!station_id) Then
            '如果不同
            
            '记录下各站点的人数
            szStation = szStation & szStationName & lStationNumber & ","
            
            szStationName = FormatDbValue(prsData!station_name)
            szStationId = FormatDbValue(prsData!station_id)
            
            lStationNumber = FormatDbValue(prsData!passenger_number)
            alNumber(prsData!ticket_type) = alNumber(prsData!ticket_type) + FormatDbValue(prsData!passenger_number2)
            adbAmount(prsData!ticket_type) = adbAmount(prsData!ticket_type) + FormatDbValue(prsData!ticket_price2)
            lReturnNumber = lReturnNumber + FormatDbValue(prsData!ticket_return_number)
            dbReturnAmount = dbReturnAmount + FormatDbValue(prsData!ticket_return_amount)
            dbReturnCharge = dbReturnCharge + FormatDbValue(prsData!ticket_return_charge)
            lChangeNumber = lChangeNumber + FormatDbValue(prsData!ticket_change_number)
            dbChangeAmount = dbChangeAmount + FormatDbValue(prsData!ticket_change_charge)
            lCancelNumber = lCancelNumber + FormatDbValue(prsData!ticket_cancel_number)
            dbCancelAmount = dbCancelAmount + FormatDbValue(prsData!ticket_cancel_amount)
            lTotalNumber = lTotalNumber + FormatDbValue(prsData!passenger_number)
            dbTotalPrice = dbTotalPrice + FormatDbValue(prsData!ticket_price)
            dbTotalTicketPrice = dbTotalTicketPrice + FormatDbValue(prsData!total_ticket_price)
                        
        Else
            '相同则累加
            lStationNumber = lStationNumber + FormatDbValue(prsData!passenger_number)
            alNumber(prsData!ticket_type) = alNumber(prsData!ticket_type) + FormatDbValue(prsData!passenger_number2)
            adbAmount(prsData!ticket_type) = adbAmount(prsData!ticket_type) + FormatDbValue(prsData!ticket_price2)
            lReturnNumber = lReturnNumber + FormatDbValue(prsData!ticket_return_number)
            dbReturnAmount = dbReturnAmount + FormatDbValue(prsData!ticket_return_amount)
            dbReturnCharge = dbReturnCharge + FormatDbValue(prsData!ticket_return_charge)
            lChangeNumber = lChangeNumber + FormatDbValue(prsData!ticket_change_number)
            dbChangeAmount = dbChangeAmount + FormatDbValue(prsData!ticket_change_charge)
            lCancelNumber = lCancelNumber + FormatDbValue(prsData!ticket_cancel_number)
            dbCancelAmount = dbCancelAmount + FormatDbValue(prsData!ticket_cancel_amount)
            lTotalNumber = lTotalNumber + FormatDbValue(prsData!passenger_number)
            dbTotalPrice = dbTotalPrice + FormatDbValue(prsData!ticket_price)
            dbTotalTicketPrice = dbTotalTicketPrice + FormatDbValue(prsData!total_ticket_price)
            
        End If
        prsData.MoveNext
    Next i

    rsTemp.AddNew
    rsTemp!bus_id = szbusID
    rsTemp!route_name = szRouteName
    rsTemp!end_station_name = szEndStationName
    rsTemp!bus_start_time = szBusStartTime
    rsTemp!transport_company_name = szTransportCompanyName
    For j = 1 To TP_TicketTypeCount
        rsTemp.Fields("number_ticket_type" & j) = alNumber(j)
        rsTemp.Fields("amount_ticket_type" & j) = adbAmount(j)
    Next j
    rsTemp!return_number = lReturnNumber
    rsTemp!return_amount = dbReturnAmount
    rsTemp!return_charge = dbReturnCharge
    rsTemp!change_number = lChangeNumber
    rsTemp!change_amount = dbChangeAmount
    rsTemp!change_charge = dbChangeCharge
    rsTemp!cancel_number = lCancelNumber
    rsTemp!cancel_amount = dbCancelAmount
    rsTemp!total_number = lTotalNumber
    rsTemp!total_price = dbTotalPrice
    rsTemp!total_ticket_price = dbTotalTicketPrice
    
    szStation = szStation & szStationName & lStationNumber
            
    rsTemp!Station = szStation
    rsTemp.Update
    
    Set m_rsData = rsTemp
    
End Sub
