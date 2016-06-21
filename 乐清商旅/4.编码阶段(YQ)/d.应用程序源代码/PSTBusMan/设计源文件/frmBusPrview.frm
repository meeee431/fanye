VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmBusPreview 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "生成环境预览"
   ClientHeight    =   4590
   ClientLeft      =   3390
   ClientTop       =   3270
   ClientWidth     =   8100
   HelpContextID   =   10000530
   Icon            =   "frmBusPrview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8100
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "预览方式"
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   120
      TabIndex        =   8
      Top             =   420
      Width           =   6450
      Begin VB.OptionButton optDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "预览某段时间车辆运行信息(&D)"
         Height          =   180
         Left            =   195
         TabIndex        =   10
         Top             =   585
         Width           =   2730
      End
      Begin VB.OptionButton optAll 
         BackColor       =   &H00E0E0E0&
         Caption         =   "预览某日的全部信息(&A)"
         Height          =   180
         Left            =   195
         TabIndex        =   9
         Top             =   285
         Value           =   -1  'True
         Width           =   2205
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Left            =   4695
         TabIndex        =   11
         Top             =   825
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   23789568
         CurrentDate     =   36392
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   1575
         TabIndex        =   12
         Top             =   825
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   23789568
         CurrentDate     =   36392
      End
      Begin MSComCtl2.DTPicker dtpDay 
         Height          =   300
         Left            =   4695
         TabIndex        =   16
         Top             =   285
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   529
         _Version        =   393216
         Format          =   23789568
         CurrentDate     =   36392
      End
      Begin VB.Label lblLine 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "->"
         Height          =   180
         Left            =   3330
         TabIndex        =   15
         Top             =   885
         Width           =   180
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期(&E):"
         Height          =   180
         Left            =   3600
         TabIndex        =   14
         Top             =   885
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期(&S):"
         Height          =   180
         Index           =   0
         Left            =   465
         TabIndex        =   13
         Top             =   885
         Width           =   1080
      End
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   315
      Left            =   6795
      TabIndex        =   5
      Top             =   1365
      Width           =   1185
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmBusPrview.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdBuild 
      Height          =   315
      HelpContextID   =   6000201
      Left            =   6795
      TabIndex        =   3
      Top             =   630
      Width           =   1185
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "生成(&B)..."
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
      MICON           =   "frmBusPrview.frx":0166
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdPreview 
      Default         =   -1  'True
      Height          =   315
      Left            =   6795
      TabIndex        =   2
      Top             =   270
      Width           =   1185
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "预览(P)"
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
      MICON           =   "frmBusPrview.frx":0182
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
      Left            =   6795
      TabIndex        =   4
      Top             =   1005
      Width           =   1185
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
      MICON           =   "frmBusPrview.frx":019E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgInfo 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   2100
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   3
      FixedCols       =   2
      BackColorFixed  =   14737632
      BackColorBkg    =   14737632
      AllowUserResizing=   3
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.Label lblBuild 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "生成环境预览(&V):"
      Height          =   180
      Left            =   165
      TabIndex        =   0
      Top             =   1785
      Width           =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "预览车次:"
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   135
      Width           =   810
   End
   Begin VB.Label lblBusList 
      BackStyle       =   0  'Transparent
      Caption         =   "0001,0002,0003,0005,0006,0008"
      Height          =   195
      Left            =   1035
      TabIndex        =   6
      Top             =   135
      Width           =   885
   End
End
Attribute VB_Name = "frmBusPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'某段时间车辆运行信息,网格的对应的行号
Const cnDate = 0
Const cnVehicle = 1
Const cnSerial = 2
Const cnSeatCount = 3
Const cnBookSeatCount = 4
Const cnCanSellSeatCount = 5
Const cnLicenseTag = 6
Const cnRowsDate = 7

'当天的运行信息,网格的对应位置
'Const cnStationID = 0
'Const cnLimitCount = 1
'Const cnMileage = 3
'Const cnLimitTime = 2
'Const cnFullTicket = 4
'Const cnHalfTicket = 5
'Const cnRowsCurrentDate = 6
Const cnSellStation = 0
Const cnStationID = 1
Const cnLimitCount = 2
Const cnMileage = 4
Const cnLimitTime = 3
Const cnFullTicket = 5
Const cnHalfTicket = 6
Const cnRowsCurrentDate = 7


Public m_bRealTime As Boolean
Public m_nRunCyle As Integer '运行周期
Public m_nCyleStartSerial As Integer '周期开始序号

Private m_oREScheme As New REScheme
Private m_oBus As New Bus
Private m_oRegularScheme As New RegularScheme
Private m_szBusID As String '车次代码
Private m_taBusVehicle() As TBusVehicleInfo  '车次车辆
Private m_oVehicle As New Vehicle '运行车辆
Private m_oReBus As New REBus
Private tStationInfo() As TBusStationSellInfo


Private Sub cmdBuild_Click()
    Dim nResult As VbMsgBoxResult
    Dim szPlanID As String
    On Error GoTo ErrHandle
    szPlanID = m_oRegularScheme.GetExecuteBusProject(dtpDay.Value).szProjectID
    m_oBus.Identify m_szBusID
    nResult = MsgBox("是否生成" & Format(dtpDay.Value, "YYYY年MM月DD日") & "---[车次" & m_szBusID & "]" & vbCrLf & "如要生成车次将保存车次的运行周期和起始序号", vbYesNo + vbQuestion + vbDefaultButton2, "生成车次")
    
    If nResult = vbYes Then
        
        SetBusy
        m_oBus.CycleStartSerialNo = m_nCyleStartSerial
        m_oBus.RunCycle = m_nRunCyle
        m_oBus.Update
        m_oREScheme.MakeRunEvironment dtpDay.Value, m_szBusID
        MsgBox "成功生成" & Format(dtpDay.Value, "YYYY年MM月DD日") & "---[车次" & m_szBusID & "]", vbInformation, "计划"
    End If
    SetNormal
Exit Sub
ErrHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
DisplayHelp Me
End Sub
Private Sub cmdPreview_Click()
    If m_bRealTime Then
        If optAll.Value Then
            RealTimePreviewDay
        ElseIf dtpEndDate <= DateAdd("D", cnPreViewMaxDays, dtpStartDate) Then
            RealTimePreview
        Else
            MsgBox "生成环境预览天数不能超过" & cnPreViewMaxDays & "天", vbInformation, "生成环境预览"
        End If
    End If
    
    On Error GoTo ErrHandle
    '判段环境中是否存在该车次，如存在环境座位按钮激活
    m_oReBus.Identify m_szBusID, IIf(optAll.Value, dtpDay.Value, dtpStartDate.Value)
    Exit Sub
ErrHandle:
    If err.Number = ERR_REBusNotExist Then
    
    Else
        ShowErrorMsg
    End If
End Sub


Private Sub Form_Load()
    AlignFormPos Me
    
    lblBusList.Caption = m_szBusID
    m_oVehicle.Init g_oActiveUser
    m_oREScheme.Init g_oActiveUser
    m_oRegularScheme.Init g_oActiveUser
    m_oBus.Init g_oActiveUser
    dtpDay.Value = Date
    dtpStartDate.Value = Date
    dtpEndDate.Value = DateAdd("d", Date, 7)
    m_oReBus.Init g_oActiveUser
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Sub optAll_Click()
'    If optAll.Value = True Then
    dtpStartDate.Enabled = False
    dtpEndDate.Enabled = False
    dtpDay.Enabled = True
    cmdBuild.Enabled = True
End Sub
Private Sub optDate_Click()
    dtpStartDate.Enabled = True
    dtpEndDate.Enabled = True
    dtpDay.Enabled = False
    cmdBuild.Enabled = False
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub
Public Sub RealTimeInit(szBus As String, BusVehicle() As TBusVehicleInfo, Optional RealTime As Boolean = False, Optional RunCyle As Integer = 1, Optional RunCyleStartNo As Integer = 1)
    m_szBusID = szBus
    m_taBusVehicle = BusVehicle
    m_bRealTime = RealTime
    m_nCyleStartSerial = RunCyleStartNo
    m_nRunCyle = RunCyle
End Sub
'生成多日的车次运行情况
Private Sub RealTimePreview()
    Dim nSerial As Integer
    Dim nDCount As Integer
    Dim nCount As Integer
    Dim dtPreview  As Date
    Dim szPlanID As String
    Dim i As Integer, j As Integer
    Dim nCol As Integer
    On Error GoTo ErrHandle
    hfgInfo.Clear
    nCount = ArrayLength(m_taBusVehicle)
    nDCount = DateDiff("d", dtpStartDate.Value, dtpEndDate.Value) + 1
    hfgInfo.Cols = nDCount + 1
    
    hfgInfo.Rows = cnRowsDate
    hfgInfo.TextMatrix(cnDate, nCol) = "日期"
    hfgInfo.TextMatrix(cnVehicle, nCol) = "车辆"
    hfgInfo.TextMatrix(cnSerial, nCol) = "序号"
    hfgInfo.TextMatrix(cnSeatCount, nCol) = "座位数"
    hfgInfo.TextMatrix(cnBookSeatCount, nCol) = "预留座位"
    hfgInfo.TextMatrix(cnCanSellSeatCount, nCol) = "可售座位数"
    hfgInfo.TextMatrix(cnLicenseTag, nCol) = "车牌"
    ShowSBInfo
    For i = 1 To nDCount
        hfgInfo.ColWidth(i) = 1000
        ShowSBInfo
        hfgInfo.Col = i
        hfgInfo.Row = 0
        dtPreview = DateAdd("d", i - 1, dtpStartDate.Value) '预览的日期
        hfgInfo.Text = Format(dtPreview, "YYYY-MM-DD")
        '获得运行车次车辆的序号
        nSerial = m_oREScheme.GetExecuteVehicleSerialNo(m_nRunCyle, m_nCyleStartSerial, dtPreview)
        hfgInfo.Row = 1
        '获得预览日期的执行计划
        szPlanID = m_oRegularScheme.GetExecuteBusProject(dtPreview).szProjectID
        For j = 1 To nCount '预览的车次的车次车辆情况
            If nSerial = m_taBusVehicle(j).nSerialNo Then
                If DateDiff("d", m_taBusVehicle(j).dtBeginStopDate, dtPreview) >= 0 And DateDiff("d", dtPreview, m_taBusVehicle(j).dtEndStopDate) >= 0 Then
                     m_oVehicle.Identify Trim(m_taBusVehicle(j).szVehicleID)
                     hfgInfo.CellForeColor = vbRed
                     hfgInfo.Text = "(停)" & m_taBusVehicle(j).szVehicleID
                     FullBusVehicle vbRed, j, szPlanID, m_taBusVehicle(j).szVehicleID, dtPreview
                Else
                     m_oVehicle.Identify Trim(m_taBusVehicle(j).szVehicleID)
                     If m_oVehicle.Status = ST_VehicleStop Then
                         hfgInfo.Text = "(停)" & m_taBusVehicle(j).szVehicleID
                         hfgInfo.CellForeColor = vbBlue
                         FullBusVehicle vbBlue, j, szPlanID, m_taBusVehicle(j).szVehicleID, dtPreview
                     Else
                         hfgInfo.CellForeColor = vbBlack
                         hfgInfo.Text = m_taBusVehicle(j).szVehicleID
                         FullBusVehicle vbBlack, j, szPlanID, m_taBusVehicle(j).szVehicleID, dtPreview
                     End If
                End If
                Exit For
            Else
                hfgInfo.Text = "无"
                If j = nCount Then
                     hfgInfo.CellForeColor = vbBlue
                End If
            End If
        Next
        '判断该车次是否停班
        m_oBus.Identify m_szBusID
        If DateDiff("d", m_oBus.BeginStopDate, dtPreview) >= 0 And DateDiff("d", dtPreview, m_oBus.EndStopDate) >= 0 Then
            hfgInfo.Row = 0
            hfgInfo.CellForeColor = &H80FF&
            hfgInfo.Text = "(停)" & hfgInfo.Text
        End If
    Next
    ShowSBInfo
        
    Exit Sub
ErrHandle:
    If err.Number = 14421 Then err.Description = "预览日期[" & Format(dtpDay.Value, "YYYY年MM月DD日") & "]运行的车次计划" & szPlanID & "中无该车次"
    ShowErrorMsg
End Sub

'生成某日的车次的票价情况

Private Function RealTimePreviewDay() As Boolean
    Dim dtPreview As Date
    Dim nStationCount As Integer
    Dim i As Integer
    Dim tStationInfo() As TBusStationSellInfo
    
    Dim szStationID As String
    Dim szSeatTypeID As String
    Dim szTemp As String
    Dim szPlanID As String
    Dim nCount As String
    Dim vLimitedSellTime As Variant
    Dim nSerial As Integer
    Dim j As Integer
    Dim szStationTemp As String
    Dim szSellStationTemp As String
    
    Dim szSellStation As String
    
    Dim h As Integer
    Dim nCol As Integer
    On Error GoTo ErrHandle
    With hfgInfo
        
        dtPreview = dtpDay.Value
        ShowSBInfo "获得计划信息..."
        szPlanID = m_oRegularScheme.GetExecuteBusProject(dtPreview).szProjectID
        m_oBus.Identify m_szBusID
        .Clear
        .Redraw = False
        nSerial = m_oREScheme.GetExecuteVehicleSerialNo(m_nRunCyle, m_nCyleStartSerial, DateAdd("d", i, dtpDay.Value))
        For i = 1 To ArrayLength(m_taBusVehicle)
            If nSerial = m_taBusVehicle(i).nSerialNo Then
                m_oVehicle.Identify m_taBusVehicle(i).szVehicleID
                Exit For
            End If
        Next
        ShowSBInfo "获得车次过站信息..."
        tStationInfo = m_oBus.GetAllStation(m_oVehicle.VehicleModel)
        nCount = ArrayLength(tStationInfo)
        If nCount = 0 Then
            RealTimePreviewDay = True
            Exit Function
        End If
        nStationCount = UBound(tStationInfo)
        
        .Cols = 2
        .ColWidth(0) = 800
        .ColWidth(1) = 800
        .Rows = cnRowsCurrentDate
        .FixedCols = 0
        .MergeCol(0) = True
        .MergeCol(1) = True
        
        .MergeRow(cnSellStation) = True
        .MergeRow(cnStationID) = True
        .MergeRow(cnLimitCount) = True
        .MergeRow(cnMileage) = True
        .MergeRow(cnLimitTime) = True
        .MergeRow(cnFullTicket) = True
        .MergeRow(cnHalfTicket) = True
        
        nCol = 0
        .TextMatrix(cnSellStation, nCol) = "售票站"
        .TextMatrix(cnStationID, nCol) = "站点代码"
        .TextMatrix(cnLimitCount, nCol) = "限售张数"
        .TextMatrix(cnMileage, nCol) = "里程数"
        .TextMatrix(cnLimitTime, nCol) = "限售时间"
        nCol = 1
        .TextMatrix(cnSellStation, nCol) = "售票站"
        .TextMatrix(cnStationID, nCol) = "站点代码"
        .TextMatrix(cnLimitCount, nCol) = "限售张数"
        .TextMatrix(cnMileage, nCol) = "里程数"
        .TextMatrix(cnLimitTime, nCol) = "限售时间"
        .TextMatrix(cnFullTicket, nCol) = "全票"
        .TextMatrix(cnHalfTicket, nCol) = "半票"
        
        .Redraw = True
        .MergeCells = flexMergeRestrictColumns
        ShowSBInfo
        
        szStationTemp = tStationInfo(1).szStationID
        szSellStationTemp = tStationInfo(1).szSellStationID
'        'szSeatTypeID = tStationInfo(1).szSeatTypeID
'        .MergeCol(0) = True
'        .MergeCol(1) = True
'        .Redraw = True
        
        .Cols = 2
        For i = 1 To nStationCount
        ShowSBInfo
        '对第一个站点
            If szStationID = tStationInfo(i).szStationID Then
                h = h + 1
                If szSeatTypeID = tStationInfo(i).szSeatTypeID Then
                    GoTo here2
                Else
                    .MergeCol(0) = True
                    .MergeCol(1) = True
                    If Not (szStationTemp = tStationInfo(i).szStationID And szSellStationTemp = tStationInfo(i).szSellStationID) Then GoTo here3
                    szStationID = tStationInfo(i).szStationID
                    
                    szSeatTypeID = tStationInfo(i).szSeatTypeID
                    .Rows = .Rows + 2
                    .MergeRow(.Rows - 2) = True
                    .MergeRow(.Rows - 1) = True
                    .Col = 0
                    .Row = .Rows - 2
                    .Text = tStationInfo(i).szSeatTypeName
                    .Col = 1
                    .Text = "全票"
                    .Col = 2
                    .Text = Round(tStationInfo(i).sgFullPrice, 2)
                    .Row = .Rows - 1
                    .Col = 0
                    .Text = tStationInfo(i).szSeatTypeName
                    .Row = .Rows - 1
                    .Col = 1
                    .Text = "半票"
                    
                    .Col = 2
                    .Text = Round(tStationInfo(i).sgHalfPrice, 2)
                    .Redraw = True
                    .MergeCells = flexMergeRestrictColumns
                    '           .ColAlignment(0) = 4
                    '          .Redraw = True
                    GoTo here2
                End If
            End If
            szStationID = tStationInfo(i).szStationID
            szSeatTypeID = tStationInfo(i).szSeatTypeID
            .Cols = .Cols + 1
            .ColWidth(.Cols - 1) = 800
            nCol = .Cols - 1
            
            .TextMatrix(cnSellStation, nCol) = tStationInfo(i).szSellStationName
            .TextMatrix(cnStationID, nCol) = tStationInfo(i).szStationName
            .TextMatrix(cnMileage, nCol) = tStationInfo(i).nMileage
            Select Case tStationInfo(i).nLimitedSellCount
            Case -1
                szTemp = "不限"
            Case 0
                szTemp = "不可售"
            Case Else
                szTemp = tStationInfo(i).nLimitedSellCount
            End Select
            .TextMatrix(cnLimitCount, nCol) = szTemp
            Select Case tStationInfo(i).sgLimitedSellTime
            Case -1
                szTemp = "不限"
            Case 0
                szTemp = "不限"
            Case Else
                If m_oBus.BusType <> 1 Then
                    vLimitedSellTime = GetStopTime(CStr(tStationInfo(i).sgLimitedSellTime), True)
                    szTemp = Format(DateAdd("n", -CInt(vLimitedSellTime), CDate(Format(dtpDay.Value, "YYYY-MM-DD") + Format(m_oBus.StartUpTime, " hh:mm"))), "YYYY-MM-DD hh:mm")
                Else
                    '滚动车次
                    szTemp = GetStopTime(CStr(tStationInfo(i).sgLimitedSellTime), False)
                End If
            End Select
            .TextMatrix(cnLimitTime, nCol) = szTemp
here3:
            If i <> 1 Then
                '取得填入票价位置
                For j = cnFullTicket To .Rows - 2
                    .Col = 0
                    .Row = j
                    If Trim(.Text) = Trim(tStationInfo(i).szSeatTypeName) Then
                        Exit For
                    End If
                Next
                '填入票价
                .Row = j
                .Col = i + 1 - h
                .Text = Round(tStationInfo(i).sgFullPrice, 2)
                .Row = j + 1
                .Text = Round(tStationInfo(i).sgHalfPrice, 2)
            Else
                '填充票价项名
                .MergeCol(0) = True
                .MergeCol(1) = True
                .MergeRow(5) = True
                .MergeRow(4) = True
                .Col = 0
'                .Row = cnSellStation
'                .Text = tStationInfo(i).szSellStationName
                .Row = cnFullTicket
                .Text = tStationInfo(i).szSeatTypeName
                .Row = cnHalfTicket
                .Text = tStationInfo(i).szSeatTypeName
                .Col = 2
'                .Row = cnSellStation
'                .Text = tStationInfo(i).szSellStationName
                .Row = cnFullTicket
                .Text = Round(tStationInfo(i).sgFullPrice, 2)
                .Row = cnHalfTicket
                .Text = Round(tStationInfo(i).sgHalfPrice, 2)
                .Redraw = True
                .MergeCells = flexMergeRestrictColumns
            End If
here2:
        Next
        
    End With
    ShowSBInfo
    Exit Function
ErrHandle:
    If err.Number = 14421 Then err.Description = "预览日期[" & Format(dtpDay.Value, "YYYY年MM月DD日") & "]运行的车次计划" & szPlanID & "中无该车次"
    ShowErrorMsg
End Function
Private Function FullReserveSeat(FullColor As OLE_COLOR, j As Integer, PlanID As String, dtPreview As Date) As Integer
Dim tReSeat As TReserveSeatInfo
m_oBus.Identify m_szBusID
tReSeat = m_oBus.GetReserverSeat(m_taBusVehicle(j).nSerialNo)
If tReSeat.nSeatCount <> 0 Then
    If DateDiff("d", tReSeat.dtBeginDate, dtPreview) >= 0 And DateDiff("d", dtPreview, tReSeat.dtEnddate) >= 0 Then
        hfgInfo.Row = 4
        hfgInfo.CellForeColor = FullColor
        hfgInfo.Text = tReSeat.nSeatCount
        FullReserveSeat = tReSeat.nSeatCount
        Exit Function
    End If
End If
hfgInfo.Row = 4
hfgInfo.Text = "0"
FullReserveSeat = 0
End Function
Private Sub FullBusVehicle(FullColor As OLE_COLOR, j As Integer, PlanID As String, VehicleId As String, dtPreview As Date)
    Dim nSeatCount As Integer
    hfgInfo.Row = cnSerial
    hfgInfo.CellForeColor = FullColor
    hfgInfo.Text = m_taBusVehicle(j).nSerialNo
    hfgInfo.Row = cnSeatCount
    hfgInfo.CellForeColor = FullColor
    hfgInfo.Text = m_oVehicle.SeatCount
    hfgInfo.Row = cnLicenseTag
    hfgInfo.CellForeColor = FullColor
    hfgInfo.Text = m_oVehicle.LicenseTag
    nSeatCount = FullReserveSeat(FullColor, j, PlanID, dtPreview)
    hfgInfo.Row = cnCanSellSeatCount
    hfgInfo.Text = m_oVehicle.SeatCount - nSeatCount
    hfgInfo.CellForeColor = FullColor
End Sub
