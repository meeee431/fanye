VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmBusTransStat 
   BackColor       =   &H00E0E0E0&
   Caption         =   "车次运量统计简报"
   ClientHeight    =   5145
   ClientLeft      =   2415
   ClientTop       =   3240
   ClientWidth     =   11880
   Icon            =   "frmBusTransStat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5145
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtSellStationID 
      Height          =   300
      Left            =   6990
      TabIndex        =   16
      Top             =   75
      Width           =   1245
   End
   Begin VB.ComboBox cboSellStation 
      Height          =   300
      Left            =   5250
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   75
      Width           =   1485
   End
   Begin VB.ComboBox cboStationID 
      Height          =   300
      Left            =   5250
      TabIndex        =   8
      Top             =   510
      Width           =   1485
   End
   Begin FText.asFlatTextBox txtRoute 
      Height          =   300
      Left            =   3150
      TabIndex        =   6
      Top             =   510
      Width           =   1005
      _ExtentX        =   1773
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
      Registered      =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin VB.TextBox txtBusID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   810
      TabIndex        =   4
      Top             =   510
      Width           =   1395
   End
   Begin RTComctl3.CoolButton cmdQuery 
      Height          =   345
      Left            =   6990
      TabIndex        =   9
      Top             =   510
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "查询(&Q)"
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
      MICON           =   "frmBusTransStat.frx":000C
      PICN            =   "frmBusTransStat.frx":0028
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
      Cancel          =   -1  'True
      Height          =   345
      Left            =   8430
      TabIndex        =   12
      Top             =   75
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "确定(&O)"
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
      MICON           =   "frmBusTransStat.frx":03C2
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
      Left            =   2760
      TabIndex        =   2
      Top             =   75
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   529
      _Version        =   393216
      Format          =   59965441
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   300
      Left            =   810
      TabIndex        =   1
      Top             =   75
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   529
      _Version        =   393216
      Format          =   59965441
      CurrentDate     =   36572
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6480
      Top             =   1590
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusTransStat.frx":03DE
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusTransStat.frx":053A
            Key             =   "FlowRun"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusTransStat.frx":0696
            Key             =   "FlowStop"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusTransStat.frx":07F2
            Key             =   "Merge"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusTransStat.frx":094E
            Key             =   "Replace"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusTransStat.frx":0AAA
            Key             =   "Checking"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusTransStat.frx":0C06
            Key             =   "Checked"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusTransStat.frx":0D62
            Key             =   "ExCheck"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusTransStat.frx":0EBE
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusTransStat.frx":101A
            Key             =   "SlitpBus"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilBus 
      Left            =   5310
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusTransStat.frx":1176
            Key             =   "PrjRun"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusTransStat.frx":1492
            Key             =   "PrjStop"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusTransStat.frx":15EE
            Key             =   "PrjDelete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusTransStat.frx":1A42
            Key             =   "RunBus"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusTransStat.frx":1B9E
            Key             =   "StopBus"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusTransStat.frx":1CFA
            Key             =   "Flow"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusTransStat.frx":1E56
            Key             =   "FlowStop"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusTransStat.frx":1FB2
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusTransStat.frx":288E
            Key             =   "RunStopBus"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvBus 
      Height          =   2685
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   4736
      SortKey         =   2
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ilBus"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "车次代码"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "发车时间"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "运行线路"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "检票口"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "状态"
         Object.Width           =   2541
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "当天车辆情况"
         Object.Width           =   4235
      EndProperty
   End
   Begin RTComctl3.CoolButton cmdChart 
      Height          =   345
      Left            =   8430
      TabIndex        =   15
      Top             =   510
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
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
      MICON           =   "frmBusTransStat.frx":29E8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&T):"
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "途经站(&S):"
      Height          =   180
      Left            =   4320
      TabIndex        =   7
      Top             =   555
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "线路(&R):"
      Height          =   180
      Left            =   2400
      TabIndex        =   5
      Top             =   555
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次(&C):"
      Height          =   180
      Left            =   60
      TabIndex        =   3
      Top             =   555
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "->"
      Height          =   180
      Left            =   2400
      TabIndex        =   11
      Top             =   135
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "日期(&B):"
      Height          =   180
      Left            =   60
      TabIndex        =   0
      Top             =   142
      Width           =   720
   End
End
Attribute VB_Name = "frmBusTransStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IConditionForm
Const cszFileName = "车次运量统计简报模板.xls"
'Public m_bOk As Boolean
Private m_rsData As Recordset
Private m_vaCustomData As Variant
Private m_szPlanID As String '当前执行计划代码

Private m_oBusProject As New BusProject

Private Sub cmdChart_Click()
    
    On Error GoTo Error_Handle
    '生成记录集
    Dim rsTemp As Recordset
    Dim i, nCount As Integer
    Dim aszBusID() As String
    Dim frmTemp As frmChart
    Dim oTicketBusDim As New TicketBusDim
    nCount = 0
    oTicketBusDim.Init m_oActiveUser
    Me.MousePointer = vbHourglass
    '得到所有要查询的车次的记录集
    For i = 1 To lvBus.ListItems.Count
        If lvBus.ListItems(i).Selected Then
            nCount = nCount + 1
            ReDim Preserve aszBusID(1 To nCount)
            aszBusID(nCount) = lvBus.ListItems(i).Text
        End If
    Next i
    If MDIMain.m_szMethod = False Then
        Set rsTemp = oTicketBusDim.GetBusTransStat(aszBusID, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(txtRoute.Text), IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
    Else
        Set rsTemp = oTicketBusDim.GetBusTransStatByCheck(aszBusID, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(txtRoute.Text), IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
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
    
    Dim rsdata3 As New Recordset
    With rsdata3.Fields
        .Append "bus_id", adBSTR
        .Append "fact_float", adBigInt
    End With
    rsdata3.Open
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        rsdata3.AddNew
        rsdata3!bus_id = FormatDbValue(rsTemp!bus_id)
        rsdata3!fact_float = FormatDbValue(rsTemp!fact_float)
        rsTemp.MoveNext
        rsdata3.Update
    Next i
    
    Dim rsdata4 As New Recordset
    With rsdata4.Fields
        .Append "bus_id", adBSTR
        .Append "total_float", adBigInt
    End With
    rsdata4.Open
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        rsdata4.AddNew
        rsdata4!bus_id = FormatDbValue(rsTemp!bus_id)
        rsdata4!total_float = FormatDbValue(rsTemp!total_float)
        rsTemp.MoveNext
        rsdata4.Update
    Next i
    
    Dim rsdata5 As New Recordset
    With rsdata5.Fields
        .Append "bus_id", adBSTR
        .Append "full_seat_rate", adBigInt
    End With
    rsdata5.Open
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        rsdata5.AddNew
        rsdata5!bus_id = FormatDbValue(rsTemp!bus_id)
        rsdata5!full_seat_rate = FormatDbValue(rsTemp!full_seat_rate)
        rsTemp.MoveNext
        rsdata5.Update
    Next i
   
      Dim rsdata6 As New Recordset
    With rsdata6.Fields
        .Append "bus_id", adBSTR
        .Append "fact_load_rate", adBigInt
    End With
    rsdata6.Open
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        rsdata6.AddNew
        rsdata6!bus_id = FormatDbValue(rsTemp!bus_id)
        rsdata6!fact_load_rate = FormatDbValue(rsTemp!fact_load_rate)
        rsTemp.MoveNext
        rsdata6.Update
    Next i
    
    Me.Hide
    Set frmTemp = New frmChart
    frmTemp.ClearChart
    frmTemp.AddChart "人数", rsData
    frmTemp.AddChart "金额", rsdata2
    frmTemp.AddChart "实际周转量", rsdata3
    frmTemp.AddChart "总周转量", rsdata4
    frmTemp.AddChart "上座率", rsdata5
    frmTemp.AddChart "实载率", rsdata6
    frmTemp.ShowChart "车次运量统计简报"
    Set frmTemp = Nothing
    Unload Me

    Exit Sub
Error_Handle:
    Set frmTemp = Nothing
    ShowErrorMsg
    
End Sub

'
'Private Sub CmdCancel_Click()
'    Unload Me
'End Sub



Private Sub cmdok_Click()

    Dim aszBusID() As String
    Dim nCount As Integer
    nCount = 0
    On Error GoTo Error_Handle
    '生成记录集
    Dim i As Integer
    Dim oTicketBusDim As New TicketBusDim



    oTicketBusDim.Init m_oActiveUser
    Me.MousePointer = vbHourglass
    '得到所有要查询的车次的记录集
    For i = 1 To lvBus.ListItems.Count
        If lvBus.ListItems(i).Selected Then
            nCount = nCount + 1
            ReDim Preserve aszBusID(1 To nCount)
            aszBusID(nCount) = lvBus.ListItems(i).Text
        End If
        
    Next i
    
    If MDIMain.m_szMethod = False Then
        Set m_rsData = oTicketBusDim.GetBusTransStat(aszBusID, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(txtRoute.Text), IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
    Else
        Set m_rsData = oTicketBusDim.GetBusTransStatByCheck(aszBusID, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(txtRoute.Text), IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
    End If
    
    ReDim m_vaCustomData(1 To 4, 1 To 2)

    m_vaCustomData(1, 1) = "统计开始日期"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY-MM-DD")

    m_vaCustomData(2, 1) = "统计结束日期"
    m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY-MM-DD")
    
    
    m_vaCustomData(3, 1) = "统计方式"
    If MDIMain.m_szMethod = False Then
        m_vaCustomData(3, 2) = "按售票统计"
    Else
        m_vaCustomData(3, 2) = "按检票统计"
    End If
    
    m_vaCustomData(4, 1) = "制表人"
    m_vaCustomData(4, 2) = m_oActiveUser.UserID
    
    Dim frmNewReport As New frmReport
    
    frmNewReport.m_lHelpContextID = Me.HelpContextID
    frmNewReport.ShowReport m_rsData, cszFileName, "车次运量统计简报", m_vaCustomData
    
    
    Me.MousePointer = vbDefault
    Unload Me
    Exit Sub

Error_Handle:
    Me.MousePointer = vbDefault
    ShowErrorMsg
End Sub


Private Sub cmdQuery_Click()
FindBus
End Sub



Private Sub dtpBeginDate_GotFocus()
    cmdQuery.Default = True
End Sub

Private Sub dtpBeginDate_LostFocus()
    cmdQuery.Default = False
End Sub

Private Sub dtpEndDate_GotFocus()
    cmdQuery.Default = True
End Sub

Private Sub dtpEndDate_LostFocus()
    cmdQuery.Default = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If UCase(Chr(KeyCode)) = "A" And Shift = 2 Then
        '如果按下Ctrl+A
        SelectAllBus
        
    End If
    
End Sub

Private Sub Form_Load()
    On Error GoTo Here
    '设置为上个月的一号到31
    Dim dyNow As Date
    dyNow = m_oParam.NowDate
    dtpBeginDate.Value = DateAdd("d", -1, dyNow)
    dtpEndDate.Value = DateAdd("d", -1, dyNow)
    Dim oRegularScheme As New RegularScheme
    oRegularScheme.Init m_oActiveUser
    m_szPlanID = oRegularScheme.GetExecuteBusProject(Now).szProjectID
    m_oBusProject.Init m_oActiveUser
    FillSellStation cboSellStation
    If Trim(m_oActiveUser.SellStationID) <> "" Then
        txtSellStationID.Enabled = False
    End If
    AlignHeadWidth Me.name, lvBus
    Exit Sub
Here:
    ShowMsg err.Description
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




Private Sub Form_Resize()
    Dim lTemp As Long
    On Error Resume Next
    lTemp = Me.ScaleHeight - 900
    lTemp = IIf(lTemp > 0, lTemp, 0)
    lvBus.Move 0, 900, Me.ScaleWidth - 50, lTemp
End Sub

Private Sub Form_Unload(Cancel As Integer)


    SaveHeadWidth Me.name, lvBus

End Sub

Private Sub lvBus_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Static m_nUpColumn As Integer
lvBus.SortKey = ColumnHeader.Index - 1
If m_nUpColumn = ColumnHeader.Index - 1 Then
    lvBus.SortOrder = lvwDescending
    m_nUpColumn = ColumnHeader.Index
Else
    lvBus.SortOrder = lvwAscending
    m_nUpColumn = ColumnHeader.Index - 1
End If
lvBus.Sorted = True
End Sub

Private Sub lvBus_DblClick()
    '显示某个车次的运量信息
End Sub

Private Sub lvBus_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
           Case vbKeyMenu
           lvBus_MouseDown vbRightButton, Shift, 1, 1
    End Select
End Sub

Private Sub lvBus_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyReturn
       lvBus_DblClick
End Select
End Sub



Private Sub lvBus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim bflg As Boolean
    With MDIMain
    
    If Button = 2 Then
        PopupMenu .pmnu_BusTrans
        
    End If
         
    
    End With
End Sub


Private Sub txtBusID_GotFocus()
txtBusID.SelStart = 0
txtBusID.SelLength = 100
cmdQuery.Default = True
End Sub


Private Sub txtBusID_LostFocus()
cmdQuery.Default = False
End Sub

Private Sub txtRoute_ButtonClick()
    
    Dim szaTemp() As String
    
    
    szaTemp = m_oShell.SelectRoute(False)
    'Set m_oShell = Nothing
    If ArrayLength(szaTemp) = 0 Then Exit Sub
    txtRoute.Text = szaTemp(1, 1) & "[" & szaTemp(1, 2) & "]"

End Sub

Public Sub SelectAllBus()
    Dim i As Integer
    For i = 1 To lvBus.ListItems.Count
        lvBus.ListItems(i).Selected = True
        
    Next i
End Sub




Public Function FindBus()
    Dim i As Integer, nCount As Integer
    Dim ltTemp As ListItem
    Dim szaBus() As String
    Dim szQueryRoute As String
    Dim szMsg As String
    Dim szStopDateAndStartDateMsg As String
    Dim bIsBusStop As Boolean
'    If m_szPlanID = "" Then
'        MsgBox "无当前执行计划", vbExclamation, Me.Caption
'        Exit Function
'    End If
    
On Error GoTo Here
    
    Me.MousePointer = vbHourglass
    lvBus.ListItems.Clear
    m_oBusProject.Identify
    '获得该计划的所有车次，并可按线路和车次代码模糊查询
    If txtRoute.Text <> "(全部)" Then
        szQueryRoute = GetLString(txtRoute.Text)
    End If
        
    szaBus = m_oBusProject.GetAllBus(txtBusID.Text, szQueryRoute, GetLString(Trim(cboStationID.Text)), True)
    nCount = ArrayLength(szaBus)
    If nCount = 0 Then
      Me.MousePointer = vbDefault:
      MsgBox "没有您需要的数据,请检查查询条件", vbInformation + vbOKOnly, Me.Caption
      Exit Function
    End If
    
    MyADDList szaBus, nCount  ''''
    
    Me.MousePointer = vbDefault
    
    
Exit Function

Here:
    Me.MousePointer = vbDefault
    ShowMsg err.Number

End Function


'增加列表
Private Function MyADDList(szaBus() As String, nCount As Integer)

Dim i As Integer
Dim ltTemp As ListItem
'Dim szaBus() As String
Dim szStopDateAndStartDateMsg As String
Dim bIsBusStop As Boolean

On Error GoTo Here

For i = 1 To nCount

    'szaBus(i, 6) = FormatDbValue(rsTemp!stop_start_date)
    'szaBus(i, 7) = FormatDbValue(rsTemp!stop_end_date)
    '由 函数判断车次是否停班 TestBusStatus
    szStopDateAndStartDateMsg = TestBusStatus(szaBus(i, 6), szaBus(i, 7), bIsBusStop)
     
    If Val(szaBus(i, 5)) <> TP_ScrollBus Then
        Set ltTemp = lvBus.ListItems.Add(, , Trim(szaBus(i, 1)), , "RunBus")
        If bIsBusStop = True Then
           ltTemp.SmallIcon = "StopBus"
        Else
           ltTemp.SmallIcon = "RunBus"
        End If
        
        ltTemp.ListSubItems.Add , , Format(szaBus(i, 2), "HH:MM:SS")
    Else
        Set ltTemp = lvBus.ListItems.Add(, , szaBus(i, 1), , "Flow")
         
        If bIsBusStop = True Then
           ltTemp.SmallIcon = "FlowStop"
        Else
           ltTemp.SmallIcon = "Flow"
        End If
         
        ltTemp.ListSubItems.Add , , Format(szaBus(i, 2), "HH:MM:SS")
     
    End If
     
    ltTemp.ListSubItems.Add , , Trim(szaBus(i, 4))
    ltTemp.ListSubItems.Add , , Trim(szaBus(i, 9))
    
    If bIsBusStop = True Then
        
        If szaBus(i, 11) <> "" Then
           
            ltTemp.ListSubItems.Add , , "车次停班且车辆停班" & szStopDateAndStartDateMsg
            ltTemp.ListSubItems.Item(4).ForeColor = vbRed
            ltTemp.ListSubItems.Add , , szaBus(i, 10) & "(停)"
            ltTemp.ListSubItems.Item(5).ForeColor = vbRed
        
        Else
           
            ltTemp.ListSubItems.Add , , "车次停班" & szStopDateAndStartDateMsg
            ltTemp.ListSubItems.Item(4).ForeColor = vbRed
            ltTemp.ListSubItems.Add , , szaBus(i, 10)
        
        End If
         
    Else
     
        If szaBus(i, 11) <> "" Then
           ltTemp.ListSubItems.Add , , "当天车辆停班" & szStopDateAndStartDateMsg
           ltTemp.ListSubItems.Item(4).ForeColor = vbRed
           ltTemp.ListSubItems.Add , , szaBus(i, 10) & "(停)"
           ltTemp.ListSubItems.Item(5).ForeColor = vbRed
        Else
        
            ltTemp.ListSubItems.Add , , "运行" & szStopDateAndStartDateMsg
            ltTemp.ListSubItems.Add , , szaBus(i, 10) & "(开)"
            ltTemp.ListSubItems.Item(4).ForeColor = vbDefault
            ltTemp.ListSubItems.Item(5).ForeColor = vbDefault

        End If
    End If
        szStopDateAndStartDateMsg = ""
        bIsBusStop = False
    Next
    Exit Function
Here:
     err.Raise err.Number
End Function


Private Sub cboStationID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   FindBusEx
End If
End Sub
Public Function FindBusEx()
    
    If cboStationID.Text = "" Or cboStationID.Text = "(全部)" Then
        If ResolveDisplay(cboStationID.Text) = "" Then
            AddcboStationid
            
        End If
    Else
        FindBus
        
        cboStationID.Text = "(全部)"
    
    End If
    
    
End Function



Public Function AddcboStationid() As Boolean
    Dim oBaseInfo As New BaseInfo
    Dim i As Integer
    Dim szaData() As String
    Dim nCount As Integer
    oBaseInfo.Init m_oActiveUser
    szaData = oBaseInfo.GetStation(, cboStationID.Text, cboStationID.Text, cboStationID.Text)
    Set oBaseInfo = Nothing
    nCount = ArrayLength(szaData)
    If nCount > 0 Then
        cboStationID.Clear
        For i = 1 To nCount
            cboStationID.AddItem Trim(szaData(i, 1)) & "[" & Trim(szaData(i, 2)) & "]"
        Next
        cboStationID.ListIndex = 0
    Else
        AddcboStationid = False
        Beep
    End If
    AddcboStationid = True
End Function



'函数判断车次是否停班,返回时间断停班的信息和车次状态
'Public Const cszEmptyDateStr = "1900-01-01"
 'Public Const cszForeverDateStr = "2050-01-01"
Private Function TestBusStatus(szStartdate As String, szEndDate As String, ByRef bIsBusStop As Boolean) As String
    Dim szMsg As String
    Dim szStartdateTemp As String
    Dim szEndDateTemp As String
    Dim dtEmptyDate As Date
    Dim dtForever As Date
    
    dtEmptyDate = CDate(cszEmptyDateStr)
    dtForever = CDate(cszForeverDateStr)
    szStartdateTemp = Format(szStartdate, "YYYY-MM-DD")
    szEndDateTemp = Format(szEndDate, "YYYY-MM-DD")
    bIsBusStop = False
    
    If DateDiff("d", CDate(szEndDateTemp), dtForever) = 0 Then '长停
        bIsBusStop = True
    Else
        'if szEndDate=cszEmptyDateStr and  szEndDate=cszEmptyDateStr '不停车
          '时间段停班
        If DateDiff("d", CDate(szStartdateTemp), dtEmptyDate) <> 0 And _
            DateDiff("d", CDate(szStartdateTemp), dtEmptyDate) <> 0 Then
            '结束时间应大于等于当天时间
            If DateDiff("d", Now, CDate(szEndDateTemp)) >= 0 Then
               szMsg = "在[" & szStartdateTemp & "到" & szEndDateTemp & "]时段停班"
               '开时时间应小于等于当天时间
               If DateDiff("d", Now, CDate(szStartdateTemp)) <= 0 Then
                  bIsBusStop = True
               End If
            
            End If
               
        End If
    End If
    TestBusStatus = szMsg
End Function


Private Sub txtRoute_GotFocus()
cmdQuery.Default = True
End Sub

Private Sub txtRoute_LostFocus()
cmdQuery.Default = False
End Sub
