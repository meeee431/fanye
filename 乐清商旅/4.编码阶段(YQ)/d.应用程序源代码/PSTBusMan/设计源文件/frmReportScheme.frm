VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#2.1#0"; "RTReportLF.ocx"
Begin VB.Form frmReportScheme 
   BackColor       =   &H80000009&
   Caption         =   "调度系统报表"
   ClientHeight    =   6420
   ClientLeft      =   3675
   ClientTop       =   2850
   ClientWidth     =   9030
   Icon            =   "frmReportScheme.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   9030
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ptResult 
      BackColor       =   &H8000000E&
      Height          =   5835
      Left            =   3105
      ScaleHeight     =   5775
      ScaleWidth      =   5145
      TabIndex        =   10
      Top             =   -75
      Width           =   5205
      Begin RTReportLF.RTReport RTReport 
         Height          =   2880
         Left            =   90
         TabIndex        =   8
         Top             =   1335
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   5080
      End
      Begin VB.PictureBox ptQ 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   1200
         Left            =   -15
         Picture         =   "frmReportScheme.frx":014A
         ScaleHeight     =   1200
         ScaleWidth      =   5100
         TabIndex        =   11
         Top             =   45
         Width           =   5100
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "XX查询"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   780
            TabIndex        =   15
            Top             =   750
            Width           =   720
         End
         Begin VB.Image imgOpen 
            Height          =   240
            Left            =   0
            Picture         =   "frmReportScheme.frx":1040
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
      End
   End
   Begin VB.PictureBox ptQuery 
      BackColor       =   &H8000000E&
      Height          =   5805
      Left            =   0
      ScaleHeight     =   5745
      ScaleWidth      =   2775
      TabIndex        =   9
      Top             =   15
      Width           =   2835
      Begin VB.Frame fraQuery 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1935
         Left            =   105
         TabIndex        =   18
         Top             =   1050
         Width           =   2535
         Begin MSComCtl2.DTPicker dtpQueryDate 
            Height          =   300
            Left            =   0
            TabIndex        =   24
            Top             =   1560
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   529
            _Version        =   393216
            Format          =   60424193
            CurrentDate     =   37854
         End
         Begin VB.ComboBox cboCheck 
            Height          =   300
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   960
            Width           =   2505
         End
         Begin VB.ComboBox cboSellStation 
            Height          =   300
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   360
            Width           =   2505
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "时间"
            Height          =   180
            Left            =   0
            TabIndex        =   23
            Top             =   1320
            Width           =   360
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "检票口"
            Height          =   255
            Left            =   0
            TabIndex        =   20
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "上车站"
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.TextBox txtBusID 
         Height          =   300
         Left            =   105
         TabIndex        =   17
         Top             =   2460
         Visible         =   0   'False
         Width           =   2505
      End
      Begin RTComctl3.CoolButton flblClose 
         Height          =   225
         Left            =   2520
         TabIndex        =   13
         Top             =   15
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   397
         BTYPE           =   8
         TX              =   "r"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Marlett"
            Size            =   9
            Charset         =   2
            Weight          =   500
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
         MICON           =   "frmReportScheme.frx":118A
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
         Left            =   105
         TabIndex        =   5
         Top             =   1860
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   60424192
         CurrentDate     =   36523
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   105
         TabIndex        =   3
         Top             =   1305
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   60424192
         CurrentDate     =   36523
      End
      Begin RTComctl3.CoolButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   330
         Left            =   1395
         TabIndex        =   7
         Top             =   3120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "关闭(&C)"
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
         MICON           =   "frmReportScheme.frx":11A6
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
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   3120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "查询"
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
         MICON           =   "frmReportScheme.frx":11C2
         PICN            =   "frmReportScheme.frx":11DE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         ItemData        =   "frmReportScheme.frx":1578
         Left            =   105
         List            =   "frmReportScheme.frx":1597
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   690
         Width           =   2505
      End
      Begin RTComctl3.FlatLabel FlatLabel1 
         Height          =   285
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483644
         HorizontalAlignment=   1
         Caption         =   "查询条件设定"
      End
      Begin VB.Label lblBusID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次(B):"
         Height          =   180
         Left            =   105
         TabIndex        =   16
         Top             =   2205
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         Caption         =   "结束日期(&E):"
         Height          =   315
         Left            =   105
         TabIndex        =   4
         Top             =   1650
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         Caption         =   "开始日期(&S):"
         Height          =   225
         Left            =   105
         TabIndex        =   2
         Top             =   1065
         Width           =   1170
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "查询类型(&T):"
         Height          =   240
         Left            =   105
         TabIndex        =   0
         Top             =   420
         Width           =   1080
      End
   End
   Begin RTComctl3.Spliter spQuery 
      Height          =   1170
      Left            =   2910
      TabIndex        =   12
      Top             =   2445
      Width           =   165
      _ExtentX        =   291
      _ExtentY        =   2064
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmReportScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const cszPtTop = 1200

Private lMoveLeft As Long
Private m_nSeatCount As Integer
Private m_nStartSeatNo As Integer
Private F1Book As TTF160Ctl.F1Book

Private Sub cboSellStation_Click()
On Error GoTo ErrHandle
    Dim oBaseInfo As BaseInfo
    Dim aszGateInfo() As String                 '检票口信息数组
    Set oBaseInfo = New BaseInfo
    Dim i As Integer
    oBaseInfo.Init g_oActiveUser
    cboCheck.Clear
    aszGateInfo = oBaseInfo.GetAllCheckGate(, ResolveDisplay(cboSellStation))
    For i = 1 To ArrayLength(aszGateInfo)
        cboCheck.AddItem MakeDisplayString(aszGateInfo(i, 1), aszGateInfo(i, 2))
    Next i
    If cboCheck.ListCount > 0 Then cboCheck.ListIndex = 0
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub cboType_Change()
'    lblTitle.Caption = "计划代码(&P):"
    Select Case cboType.ListIndex
    Case 0 '0计划车次信息
        dtpStartDate.Enabled = False
        dtpEndDate.Enabled = False
        fraQuery.Visible = False
    Case 1 '1计划车次车辆安排
        dtpStartDate.Enabled = False
        dtpEndDate.Enabled = False
        fraQuery.Visible = False
    Case 2 '2环境车次运行情况
        dtpStartDate.Enabled = True
        dtpEndDate.Enabled = True
        fraQuery.Visible = False
    Case 5 '5计划车次停班统计
        dtpStartDate.Enabled = True
        dtpEndDate.Enabled = True
        cmdOk.Enabled = True
        fraQuery.Visible = False
    Case 6 '6环境车次停班统计
        dtpStartDate.Enabled = True
        dtpEndDate.Enabled = True
        cmdOk.Enabled = True
        fraQuery.Visible = False
        dtpQueryDate.Visible = False
    Case 7 '7环境车次加班统计
        dtpStartDate.Enabled = True
        dtpEndDate.Enabled = True
        cmdOk.Enabled = True
        fraQuery.Visible = False
    Case 8 '8环境车次拆分情况
        dtpEndDate.Enabled = False
        dtpStartDate.Enabled = True
        fraQuery.Visible = False
    Case 9 '9按地区取得车次信息
        dtpStartDate.Enabled = True
        fraQuery.Visible = False
    Case 10 '按站点
        fraQuery.Visible = False
        dtpStartDate.Enabled = False
        dtpEndDate.Enabled = False
    Case 11 '行车记录表
        dtpStartDate.Enabled = False
        dtpEndDate.Enabled = False
        fraQuery.Visible = True
        cboCheck.Enabled = True
        FillSellStation '填充上车站
        dtpQueryDate.Value = Date
        dtpQueryDate.Enabled = True
    Case 12 '安全门检
        fraQuery.Visible = True
        dtpStartDate.Enabled = False
        dtpEndDate.Enabled = False
        FillSellStation
        cboCheck.Enabled = False
        dtpQueryDate.Value = Date
        dtpQueryDate.Enabled = True
    Case 13 '日行车计划
        fraQuery.Visible = True
        dtpStartDate.Enabled = False
        dtpEndDate.Enabled = False
        FillSellStation
        cboCheck.Enabled = False
        dtpQueryDate.Value = Date
        dtpQueryDate.Enabled = True
    Case 14 '月行车计划
         fraQuery.Visible = True
        dtpStartDate.Enabled = False
        dtpEndDate.Enabled = False
        FillSellStation
        cboCheck.Enabled = False
        dtpQueryDate.Value = Date
        dtpQueryDate.Enabled = False
    Case 15 '班次资料表
        dtpStartDate.Enabled = False
        dtpEndDate.Enabled = False
        fraQuery.Visible = False
        dtpQueryDate.Value = Date
        dtpQueryDate.Enabled = True
     
    End Select
    If cboType.ListIndex = 7 Then
        lblBusID.Visible = True
        txtBusID.Visible = True
    Else
        lblBusID.Visible = False
        txtBusID.Visible = False
    End If
End Sub

Private Sub cboType_Click()
    cmdOk.Enabled = True
    cboType_Change
    lblTitle.Caption = cboType.Text
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    On Error GoTo ErrorHandle
    SetBusy
    F1Book.DeleteRange F1Book.TopRow, F1Book.LeftCol, F1Book.MaxRow, F1Book.MaxCol, F1ShiftRows
    
    Select Case cboType.ListIndex
    Case 0 '0计划车次信息
        PlanBusInfo
    Case 1 '1计划车次车辆安排
        PlanBusVehicleInfo
    Case 2 '2环境车次运行情况
        ReBusInfo
    Case 3 '3线路信息
        RouteInfo
    Case 4 '4车辆信息
        VehicleInfo
    Case 5 '5计划车次停班统计
        PlanStopInfo
    Case 6 '6环境车次停班统计
        ReBusStopInfo
    Case 7 '7环境车次加班统计
        ReBusAddInfo
    Case 8 '8环境车次拆分情况
        ReBusSiltpInfo
    Case 9 '9按地区取得车次信息
        PlanBusVehicleInfo
    Case 10 '站点查询
        StationInfo
    Case 11  '公司行车记录表
        CompanyVechileInfo
    Case 12 '公司道路营运车辆安全门检记录表
        CompanyVechileSafeInfo
    Case 13 '日服务作业计划
        CompanyDayWorkPlan
    Case 14 '总服务作业计划
        CompanyWorkPlan
    Case 15 '班次资料表
        BusInfo
    End Select
    RTReport.SetFocus
    SetNormal
    ShowSBInfo ""
'    ShowSBInfo "共有" & m_lRange & "条记录", ESB_ResultCountInfo
Exit Sub
ErrorHandle:
    ShowSBInfo ""
    SetNormal
    ShowErrorMsg
End Sub


Private Sub flblClose_Click()
    ptQuery.Visible = False
    imgOpen.Visible = True
    lMoveLeft = 240
    spQuery.LayoutIt
End Sub

Private Sub Form_Activate()
    MDIScheme.SetPrintEnabled True
End Sub
Private Sub Form_Deactivate()
    MDIScheme.SetPrintEnabled False
End Sub

Private Sub Form_Load()
    spQuery.InitSpliter ptQuery, ptResult
    lMoveLeft = 0
    FillQueryType
    dtpStartDate.Value = Date
    dtpEndDate.Value = Date
    Set F1Book = RTReport.CellObject
    F1Book.ShowColHeading = True
    F1Book.ShowRowHeading = True
End Sub

Private Sub FillQueryType()
    cboType.Clear
    cboType.AddItem "计划车次信息"
    cboType.AddItem "计划车次车辆安排"
    cboType.AddItem "环境车次运行情况"
    cboType.AddItem "线路信息"
    cboType.AddItem "车辆信息"
    cboType.AddItem "计划车次停班统计"
    cboType.AddItem "环境车次停班统计"
    cboType.AddItem "环境车次加班统计" '玉环加入
    cboType.AddItem "环境车次拆分情况"
    cboType.AddItem "按地区取得车次信息"
    cboType.AddItem "站点信息"
    
'*******************温岭加入**********************
    cboType.AddItem "公司行车记录表"
    cboType.AddItem "公司道路营运车辆安全门检记录表"
    cboType.AddItem "日服务作业计划"
    cboType.AddItem "月服务作业计划"
'**************************************************

    cboType.AddItem "班次资料表" '玉环加入
    
    cboType.ListIndex = 0
    
End Sub
Private Sub Form_Resize()
    spQuery.LayoutIt
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIScheme.SetPrintEnabled False
End Sub

Private Sub imgOpen_Click()
    ptQuery.Visible = True
    imgOpen.Visible = False
    lMoveLeft = 0
    spQuery.LayoutIt
End Sub


Private Sub ptResult_Resize()
    Dim lTemp As Long
    lTemp = IIf((ptResult.ScaleHeight - cszPtTop) <= 0, lTemp, ptResult.ScaleHeight - cszPtTop)
    RTReport.Move 0, cszPtTop, ptResult.ScaleWidth, lTemp
End Sub

Private Sub PlanBusVehicleInfo()
    Dim oPlan As New BusProject
    Dim oBaseInfo As New BaseInfo
    Dim szTemp() As TVehcileSeatType
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim vData As Variant
On Error GoTo ErrorHandle
    ShowSBInfo "获得车次车辆..."
    oPlan.Init g_oActiveUser
    oBaseInfo.Init g_oActiveUser

    If cboType.ListIndex = 8 Then
        Dim oSheme As New RegularScheme
        oPlan.Identify
        Set oSheme = Nothing
        Set rsTemp = oPlan.GetBusVehicleReport(g_szExePriceTable)
    Else
        oPlan.Identify
        Set rsTemp = oPlan.GetAllBusVehicleReport
    End If
    F1Book.MaxCol = 19
    F1Book.MaxRow = rsTemp.RecordCount + 1
    WriteProcessBar , , rsTemp.RecordCount
    '填充表头
    F1Book.TextRC(1, 1) = "车次"
    F1Book.TextRC(1, 2) = "线路"
    F1Book.TextRC(1, 3) = "发车时间"
    F1Book.TextRC(1, 4) = "检票口"
    F1Book.TextRC(1, 5) = "车次类型"
    F1Book.TextRC(1, 6) = "循环周期"
    F1Book.TextRC(1, 7) = "起始序号"
    F1Book.TextRC(1, 8) = "车辆序号"
    F1Book.TextRC(1, 9) = "车辆代码"
    F1Book.TextRC(1, 10) = "车牌"
    F1Book.TextRC(1, 11) = "车型"
    F1Book.TextRC(1, 12) = "座位数"
    F1Book.TextRC(1, 13) = "起始座号"
    F1Book.TextRC(1, 14) = "座位分配"
    F1Book.TextRC(1, 15) = "参运公司"
    F1Book.TextRC(1, 16) = "拆帐公司"
    F1Book.TextRC(1, 17) = "车主"
    F1Book.TextRC(1, 18) = "停班开始日期"
    F1Book.TextRC(1, 19) = "停班结束日期"
    For i = 2 To rsTemp.RecordCount + 1
        WriteProcessBar , i - 1, rsTemp.RecordCount
        F1Book.TextRC(i, 1) = Trim(rsTemp!bus_id)
        F1Book.TextRC(i, 2) = Trim(rsTemp!route_name)
        F1Book.TextRC(i, 3) = Format(rsTemp!bus_start_time, "HH:MM")
        F1Book.TextRC(i, 4) = Trim(rsTemp!check_gate_name)
'        If rsTemp!bus_type = TP_ScrollBus Then
'            f1book.TextRC i,5, "滚动"
'        Else
'            f1book.TextRC i,5, "固定"
'        End If
        F1Book.TextRC(i, 5) = Trim(rsTemp!bus_type_name)
        F1Book.TextRC(i, 6) = Trim(rsTemp!bus_run_cycle)
        F1Book.TextRC(i, 7) = Trim(rsTemp!run_start_serial)
        'run_start_serial
        F1Book.TextRC(i, 8) = Trim(rsTemp!vehicle_serial)
       F1Book.TextRC(i, 9) = Trim(rsTemp!vehicle_id)
        F1Book.TextRC(i, 10) = Trim(rsTemp!license_tag_no)
        F1Book.TextRC(i, 11) = Trim(rsTemp!vehicle_type_short_name)


        m_nSeatCount = Trim(rsTemp!seat_quantity)
        If rsTemp!sale_stand_ticket_quantity = 0 Then
            F1Book.TextRC(i, 12) = Trim(rsTemp!seat_quantity)
        Else
            F1Book.TextRC(i, 12) = Trim(rsTemp!seat_quantity) & "(" & Trim(rsTemp!sale_stand_ticket_quantity) & ")"
        End If
        m_nStartSeatNo = Val((rsTemp!start_seat_no))
        F1Book.TextRC(i, 13) = Trim(rsTemp!start_seat_no)
        szTemp = oBaseInfo.GetAllVehicleSeatTypeInfo(Trim(rsTemp!vehicle_id))
        F1Book.TextRC(i, 14) = FindSetSeatInfo(szTemp)
        F1Book.TextRC(i, 15) = Trim(rsTemp!transport_company_short_name)
        F1Book.TextRC(i, 16) = Trim(rsTemp!split_company_short_name)
        F1Book.TextRC(i, 17) = Trim(rsTemp!owner_name)
        If rsTemp!stop_start_date = CDate(cszEmptyDateStr) Then
            F1Book.TextRC(i, 18) = ""
            F1Book.TextRC(i, 19) = ""
        Else
            F1Book.TextRC(i, 18) = Format(rsTemp!stop_start_date, "YYYY-MM-DD")
            F1Book.TextRC(i, 19) = Format(rsTemp!stop_end_date, "YYYY-MM-DD")
        End If

        If rsTemp!Status = 1 Then
            vData = F1Book.TextRC(i, 18)
            If vData <> "" Then
                vData = vData & "且车辆停"
            Else
                vData = vData & "车辆停"
            End If
            F1Book.TextRC(i, 17) = vData
            vData = F1Book.TextRC(i, 19)
            If vData <> "" Then
                vData = vData & "且车辆停"
            Else
                vData = vData & "车辆停"
            End If
            vData = F1Book.TextRC(i, 19)
        End If

    rsTemp.MoveNext
    Next
    WriteProcessBar False
    
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub PlanBusInfo()
    Dim oPlan As New BusProject
    Dim rsTemp As Recordset
    Dim i As Integer
On Error GoTo ErrorHandle
    ShowSBInfo "获得车次信息..."
    oPlan.Init g_oActiveUser
    oPlan.Identify
    Set rsTemp = oPlan.GetAllBusReport
    F1Book.MaxCol = 9
    F1Book.MaxRow = rsTemp.RecordCount + 1
    WriteProcessBar , , rsTemp.RecordCount
    F1Book.TextRC(1, 1) = "车次"
    F1Book.TextRC(1, 2) = "线路"
    F1Book.TextRC(1, 3) = "发车时间"
    F1Book.TextRC(1, 4) = "检票口"
    F1Book.TextRC(1, 5) = "车次类型"
    F1Book.TextRC(1, 6) = "运行周期"
    F1Book.TextRC(1, 7) = "起始序号"
    F1Book.TextRC(1, 8) = "停班开始日期"
    F1Book.TextRC(1, 9) = "停班结束日期"
    F1Book.ColWidth(8) = 3000
    F1Book.ColWidth(9) = 3000
    For i = 2 To rsTemp.RecordCount + 1
        WriteProcessBar , i - 1, rsTemp.RecordCount
        F1Book.TextRC(i, 1) = Trim(rsTemp!bus_id)
        F1Book.TextRC(i, 2) = Trim(rsTemp!route_name)
        F1Book.TextRC(i, 3) = Format(rsTemp!bus_start_time, "HH:MM")
        F1Book.TextRC(i, 4) = Trim(rsTemp!check_gate_name)
        If rsTemp!bus_type = TP_ScrollBus Then
            F1Book.TextRC(i, 5) = "滚动"
        Else
            F1Book.TextRC(i, 5) = "固定"
        End If
        F1Book.TextRC(i, 6) = Trim(rsTemp!bus_run_cycle)
        F1Book.TextRC(i, 7) = Trim(rsTemp!run_start_serial)
        If rsTemp!stop_start_date = CDate(cszEmptyDateStr) Then
            F1Book.TextRC(i, 8) = ""
            F1Book.TextRC(i, 9) = ""
        Else
            F1Book.TextRC(i, 8) = Format(rsTemp!stop_start_date, "YYYY-MM-DD")
            F1Book.TextRC(i, 9) = Format(rsTemp!stop_end_date, "YYYY-MM-DD")
        End If
    rsTemp.MoveNext
    Next
    WriteProcessBar False
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

'班次资料表
Private Sub BusInfo()
On Error GoTo ErrHandle
    Dim oRScheme As New REScheme
    Dim rsTemp As Recordset
    Dim aszTmp As Variant
    ReDim aszTmp(1 To 1, 1 To 2)
    aszTmp(1, 1) = "日期"
    aszTmp(1, 2) = Format(dtpQueryDate.Value, "yyyy年MM月dd日")
    Set rsTemp = oRScheme.GetBusInfo()
    ShowReport rsTemp, "班次资料表.xls", "班次资料表", aszTmp
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub ReBusInfo()
    Dim oRScheme As New REScheme
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim szTemp As String
On Error GoTo ErrorHandle
    ShowSBInfo "获得环境车次..."
    oRScheme.Init g_oActiveUser
    Set rsTemp = oRScheme.GetREBusReport(dtpStartDate.Value, dtpEndDate.Value)
    F1Book.MaxCol = 10
    F1Book.MaxRow = rsTemp.RecordCount + 1
    WriteProcessBar , , rsTemp.RecordCount
    F1Book.TextRC(1, 1) = "车次日期"
    F1Book.TextRC(1, 2) = "车次代码"
    F1Book.TextRC(1, 3) = "线路名称"
    F1Book.TextRC(1, 4) = "发车时间"
    F1Book.TextRC(1, 5) = "检票口"
    F1Book.TextRC(1, 6) = "车次类型"
    F1Book.TextRC(1, 7) = "车次状态"
    F1Book.TextRC(1, 8) = "车牌"
    F1Book.TextRC(1, 9) = "参运公司"
    F1Book.TextRC(1, 10) = "车主"
    For i = 2 To rsTemp.RecordCount + 1
        WriteProcessBar , i - 1, rsTemp.RecordCount
        F1Book.TextRC(i, 1) = Format(rsTemp!bus_date, "YYYY-MM-DD")
        F1Book.TextRC(i, 2) = Trim(rsTemp!bus_id)
        F1Book.TextRC(i, 3) = Trim(rsTemp!route_name)
        F1Book.TextRC(i, 4) = Format(rsTemp!bus_start_time, "HH:mm")
        F1Book.TextRC(i, 5) = Trim(rsTemp!check_gate_id)
        If rsTemp!bus_type = TP_ScrollBus Then
            F1Book.TextRC(i, 6) = "滚动"
        Else
            F1Book.TextRC(i, 6) = "固定"
        End If
        Select Case rsTemp!Status
        Case ST_BusChecking
            szTemp = "正检"
        Case ST_BusMergeStopped
            szTemp = "并班"
        Case ST_BusNormal
            szTemp = "普通"
        Case ST_BusExtraChecking
            szTemp = "补检"
        Case ST_BusStopCheck
            szTemp = "停检"
        Case ST_BusStopped
            szTemp = "停班"
        End Select
        F1Book.TextRC(i, 7) = szTemp
        F1Book.TextRC(i, 8) = Trim(rsTemp!license_tag_no)
        F1Book.TextRC(i, 9) = Trim(rsTemp!transport_company_short_name)
        F1Book.TextRC(i, 10) = Trim(rsTemp!owner_name)
        rsTemp.MoveNext
    Next
    WriteProcessBar False
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub RouteInfo()
    Dim oBase As New BaseInfo
    Dim aszTemp() As String
    Dim nCount As Integer, i As Integer
On Error GoTo ErrorHandle
    SetBusy
    ShowSBInfo "获得线路信息..."
    oBase.Init g_oActiveUser
    aszTemp = oBase.GetRouteEx
    nCount = ArrayLength(aszTemp)
    If nCount = 0 Then Exit Sub
    F1Book.MaxCol = 6
    F1Book.MaxRow = nCount + 1
    WriteProcessBar , nCount, , True
    F1Book.TextRC(1, 1) = "线路代码"
    F1Book.TextRC(1, 2) = "线路名称"
    F1Book.TextRC(1, 3) = "途径站"
    F1Book.TextRC(1, 4) = "终点站"
    F1Book.TextRC(1, 5) = "状态"
    For i = 1 To nCount
        WriteProcessBar , i, nCount, "获得线路信息..."
        F1Book.TextRC(i + 1, 1) = aszTemp(i, 1)
        F1Book.TextRC(i + 1, 2) = aszTemp(i, 2)
        F1Book.TextRC(i + 1, 3) = aszTemp(i, 3)
        F1Book.TextRC(i + 1, 4) = aszTemp(i, 4)
        F1Book.TextRC(i + 1, 5) = aszTemp(i, 5)
    Next
    WriteProcessBar False
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub StationInfo()
    Dim aszStation() As String
    Dim szTemp As String
    Dim oBaseInfo As BaseInfo
    Dim i As Integer
    Dim nCount As Integer
    On Error GoTo ErrorHandle
    SetBusy
    Set oBaseInfo = New BaseInfo
    oBaseInfo.Init g_oActiveUser
    aszStation = oBaseInfo.GetStation()
    nCount = ArrayLength(aszStation)
    F1Book.MaxCol = 5
    F1Book.MaxRow = nCount + 1
    F1Book.TextRC(1, 1) = "站点代码"
    F1Book.TextRC(1, 2) = "站点名称"
    F1Book.TextRC(1, 3) = "输入码"
    F1Book.TextRC(1, 4) = "是否可售"
'    F1Book.TextRC(1, 5) = "本地码"
    F1Book.TextRC(1, 5) = "地区"
    For i = 1 To nCount
        F1Book.TextRC(i + 1, 1) = aszStation(i, 1)
        F1Book.TextRC(i + 1, 2) = aszStation(i, 2)
        F1Book.TextRC(i + 1, 3) = aszStation(i, 3)
        If Val(aszStation(i, 4)) <> TP_CanSellTicket Then
            szTemp = "不可售"
            
        Else
            szTemp = "可售"
        End If
        F1Book.TextRC(i + 1, 4) = szTemp
'        F1Book.TextRC(i + 1, 5) = aszStation(i, 5)
        F1Book.TextRC(i + 1, 5) = aszStation(i, 6)
    Next
    
    SetNormal
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg

End Sub

Private Sub VehicleInfo()
    Dim oBase As New BaseInfo
    Dim aszTemp() As String
    Dim nCount As Integer, i As Integer
On Error GoTo ErrorHandle
    ShowSBInfo "获得车辆信息..."
    oBase.Init g_oActiveUser
    aszTemp = oBase.GetVehicle
    nCount = ArrayLength(aszTemp)
    If nCount = 0 Then Exit Sub
    F1Book.MaxCol = 8
    F1Book.MaxRow = nCount + 1
    i = 1
    WriteProcessBar , , nCount
    F1Book.TextRC(i, 1) = "车辆代码"
    F1Book.TextRC(i, 2) = "车牌"
    F1Book.TextRC(i, 3) = "座位数"
    F1Book.TextRC(i, 4) = "参运公司"
    F1Book.TextRC(i, 5) = "车主"
    F1Book.TextRC(i, 6) = "车型代码"
    F1Book.TextRC(i, 7) = "车型"
    F1Book.TextRC(i, 8) = "注释"
    For i = 1 To nCount
        WriteProcessBar , i, nCount
        F1Book.TextRC(i + 1, 1) = aszTemp(i, 1)
        F1Book.TextRC(i + 1, 2) = aszTemp(i, 2)
        F1Book.TextRC(i + 1, 3) = aszTemp(i, 3)
        F1Book.TextRC(i + 1, 4) = aszTemp(i, 4)
        F1Book.TextRC(i + 1, 5) = aszTemp(i, 5)
        F1Book.TextRC(i + 1, 6) = aszTemp(i, 7)
        F1Book.TextRC(i + 1, 7) = aszTemp(i, 8)
        F1Book.TextRC(i + 1, 8) = aszTemp(i, 9)
    Next
    WriteProcessBar False
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub




Private Sub txtPlanID_Click()
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    On Error GoTo ErrorHandle
    oShell.Init g_oActiveUser
    Select Case cboType.ListIndex
    Case 7 '环境车次
        aszTemp = oShell.SelectREBus(dtpEndDate.Value)
        If ArrayLength(aszTemp) = 0 Then Exit Sub
    Case 8 '按地区取的车次
        aszTemp = oShell.SelectArea()
        If ArrayLength(aszTemp) = 0 Then Exit Sub
    Case Else
'        aszTemp = oShell.selectProject()
'        If ArrayLength(aszTemp) = 0 Then Exit Sub
'        txtPlanID.Text = aszTemp(1, 1) & "[" & Trim(aszTemp(1, 2)) & "]"
    End Select
    Set oShell = Nothing
    Exit Sub
ErrorHandle:
    Set oShell = Nothing
    SetNormal
    ShowErrorMsg
End Sub

Public Sub PrintReport(Optional pbShowDialog As Boolean = False)
    On Error GoTo ErrorHandle
    RTReport.PrintReport pbShowDialog
    Exit Sub
ErrorHandle:
End Sub

Public Sub PreView()
    RTReport.PrintView
End Sub

Public Sub PageSet()
    RTReport.OpenDialog EDialogType.PAGESET_TYPE
End Sub

Public Sub PrintSet()
    RTReport.OpenDialog EDialogType.PRINT_TYPE
End Sub
'导出文件
Public Sub ExportFile()
    Dim szFileName As String
    szFileName = RTReport.OpenDialog(EDialogType.EXPORT_FILE)
End Sub
'导出文件并打开
Public Sub ExportFileOpen()
    Dim szFileName As String
    szFileName = RTReport.OpenDialog(EDialogType.EXPORT_FILE)
    If szFileName <> "" Then
        OpenLinkedFile szFileName
    End If
End Sub

Public Sub PlanStopInfo()
    Dim oRScheme As New REScheme
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim szTemp As String
On Error GoTo ErrorHandle
    oRScheme.Init g_oActiveUser
    ShowSBInfo "获得车次停班记录..."
    Set rsTemp = oRScheme.GetPlanStopReport(dtpStartDate.Value, dtpEndDate.Value)
    F1Book.MaxCol = 6
    F1Book.MaxRow = rsTemp.RecordCount + 1
    WriteProcessBar , , rsTemp.RecordCount
    i = 1
    F1Book.TextRC(i, 1) = "车次"
    F1Book.TextRC(i, 2) = "线路代码"
    F1Book.TextRC(i, 3) = "发车时间"
    F1Book.TextRC(i, 4) = "应发班次"
    F1Book.TextRC(i, 5) = "实际班次"
    F1Book.TextRC(i, 6) = "发班率"
    F1Book.TextRC(i, 7) = "运行周期"
    For i = 2 To rsTemp.RecordCount + 1
        WriteProcessBar , i - 1, rsTemp.RecordCount
        F1Book.TextRC(i, 1) = Trim(rsTemp!bus_id)
        F1Book.TextRC(i, 2) = Trim(rsTemp!route_name)
        F1Book.TextRC(i, 3) = Format(rsTemp!bus_start_time, "HH:MM")
        F1Book.TextRC(i, 4) = Trim(rsTemp!Count)
        F1Book.TextRC(i, 5) = Trim(rsTemp!Count - rsTemp!stop_count)
        F1Book.TextRC(i, 6) = Format((rsTemp!Count - rsTemp!stop_count) / rsTemp!Count, "00%")
        F1Book.TextRC(i, 7) = Trim(rsTemp!bus_run_cycle)
        rsTemp.MoveNext
    Next
    WriteProcessBar False
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Public Sub ReBusStopInfo()
    Dim oRScheme As New REScheme
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim szTemp As String
On Error GoTo ErrorHandle
    oRScheme.Init g_oActiveUser
    ShowSBInfo "获得环境车次停班..."
    Set rsTemp = oRScheme.GetREBusStopReport(dtpStartDate.Value, dtpEndDate.Value)
    F1Book.MaxCol = 6
    If rsTemp Is Nothing Then Exit Sub
    F1Book.MaxRow = rsTemp.RecordCount + 1
    WriteProcessBar , rsTemp.RecordCount + 1, , True
    i = 1
    F1Book.TextRC(i, 1) = "车次"
    F1Book.TextRC(i, 2) = "车牌"
    F1Book.TextRC(i, 3) = "参运公司"
    F1Book.TextRC(i, 4) = "车主"
    F1Book.TextRC(i, 5) = "停班数"
    F1Book.TextRC(i, 6) = "线路名称"
    F1Book.TextRC(i, 7) = "发车时间"
    For i = 2 To rsTemp.RecordCount + 1
        WriteProcessBar , i - 1, rsTemp.RecordCount
        F1Book.TextRC(i, 1) = Trim(rsTemp!bus_id)
        F1Book.TextRC(i, 2) = Trim(rsTemp!license_tag_no)
        F1Book.TextRC(i, 3) = Trim(rsTemp!company_name)
        F1Book.TextRC(i, 4) = Trim(rsTemp!owner_name)
        F1Book.TextRC(i, 5) = Trim(rsTemp!stop_count)
        F1Book.TextRC(i, 6) = Trim(rsTemp!route_name)
        F1Book.TextRC(i, 7) = Format(rsTemp!bus_start_time, "HH:mm")
        rsTemp.MoveNext
    Next
'    F1Book.DoRedrawAll
    WriteProcessBar False
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Public Sub ReBusAddInfo()
    Dim oRScheme As New REScheme
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim szTemp As String
On Error GoTo ErrorHandle
    oRScheme.Init g_oActiveUser
    ShowSBInfo "获得环境车次加班..."
    Set rsTemp = oRScheme.GetREBusAddReport(dtpStartDate.Value, dtpEndDate.Value)
    F1Book.MaxCol = 6
    If rsTemp Is Nothing Then Exit Sub
    F1Book.MaxRow = rsTemp.RecordCount + 1
    WriteProcessBar , rsTemp.RecordCount + 1, , True
    i = 1
    F1Book.TextRC(i, 1) = "车次"
    F1Book.TextRC(i, 2) = "车牌"
    F1Book.TextRC(i, 3) = "参运公司"
    F1Book.TextRC(i, 4) = "车主"
    F1Book.TextRC(i, 5) = "加班数"
    F1Book.TextRC(i, 6) = "线路名称"
    F1Book.TextRC(i, 7) = "发车时间"
    For i = 2 To rsTemp.RecordCount + 1
        WriteProcessBar , i - 1, rsTemp.RecordCount
        F1Book.TextRC(i, 1) = Trim(rsTemp!bus_id)
        F1Book.TextRC(i, 2) = Trim(rsTemp!license_tag_no)
        F1Book.TextRC(i, 3) = Trim(rsTemp!company_name)
        F1Book.TextRC(i, 4) = Trim(rsTemp!owner_name)
        F1Book.TextRC(i, 5) = Trim(rsTemp!add_count)
        F1Book.TextRC(i, 6) = Trim(rsTemp!route_name)
        F1Book.TextRC(i, 7) = Format(rsTemp!bus_start_time, "HH:mm")
        rsTemp.MoveNext
    Next
'    F1Book.DoRedrawAll
    WriteProcessBar False
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Public Sub ReBusSiltpInfo()
Dim oRebus As New REBus
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim szTemp() As String
    Dim nCountInfo As Integer
On Error GoTo ErrorHandle
    oRebus.Init g_oActiveUser
    ShowSBInfo "获得环境拆分车次..."
    
    szTemp = oRebus.GetSlitpInfo(txtBusID.Text, dtpStartDate.Value)
    nCountInfo = ArrayLength(szTemp)
    F1Book.MaxCol = 10

    F1Book.MaxRow = nCountInfo + 1
    WriteProcessBar , , nCountInfo
    i = 1
    F1Book.TextRC(i, 1) = "目标车次"
    F1Book.TextRC(i, 2) = "拆分车次"
    F1Book.TextRC(i, 3) = "当前痤号"
    F1Book.TextRC(i, 4) = "原痤号"
    F1Book.TextRC(i, 5) = "票号"
    F1Book.TextRC(i, 6) = "检票口"
    F1Book.TextRC(i, 7) = "车型"
    F1Book.TextRC(i, 8) = "总痤位"
    F1Book.TextRC(i, 9) = "终点站"
    F1Book.TextRC(i, 10) = "发车时间"
    For i = 2 To nCountInfo + 1
        WriteProcessBar , i - 1, nCountInfo
        F1Book.TextRC(i, 1) = szTemp(i - 1, 1)
        F1Book.TextRC(i, 2) = g_szExePriceTable
        F1Book.TextRC(i, 3) = szTemp(i - 1, 2)
        F1Book.TextRC(i, 4) = szTemp(i - 1, 3)
        F1Book.TextRC(i, 5) = szTemp(i - 1, 4)
        F1Book.TextRC(i, 6) = szTemp(i - 1, 5)
        F1Book.TextRC(i, 7) = szTemp(i - 1, 6)
        F1Book.TextRC(i, 8) = szTemp(i - 1, 7)
        F1Book.TextRC(i, 9) = szTemp(i - 1, 8)
        F1Book.TextRC(i, 10) = Format(szTemp(i - 1, 9))
    Next
'    F1Book.DoRedrawAll
    WriteProcessBar False
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Function FindSetSeatInfo(szTemp() As TVehcileSeatType) As String
    Dim nCount As Integer
    Dim i As Integer
    Dim seatInfo As String
    Dim seatInfoTemp As String
    Dim sz As String
    Dim nCountSeat  As Integer
    nCount = ArrayLength(szTemp)
    sz = ","
    If nCount = 0 Then
        seatInfo = "全部普通"
    Else
        For i = 1 To nCount
            If i <> 1 Then
                seatInfo = seatInfo & sz
            End If
            If szTemp(i).szStartSeatNo <= szTemp(i).szEndSeatNo Then
                nCountSeat = nCountSeat - CInt(szTemp(i).szStartSeatNo) + CInt(szTemp(i).szEndSeatNo) + 1
            Else
                nCountSeat = nCountSeat + CInt(szTemp(i).szStartSeatNo) - CInt(szTemp(i).szEndSeatNo) + 1
            End If
            seatInfo = seatInfo & Trim(CStr(szTemp(i).szStartSeatNo)) & "～" & Trim(CStr(szTemp(i).szEndSeatNo)) & " " & Trim(szTemp(i).szSeatTypeName)
        Next
    End If
    If m_nSeatCount > nCountSeat Then
        seatInfo = seatInfo & sz & "其它普通"
    End If
    FindSetSeatInfo = seatInfo
End Function

'公司行车记录表
Private Sub CompanyVechileInfo()
    Dim g_oREScheme As New REScheme       '当前检票对象
    '得到当天所有车次信息
    Dim rsBusInfo As Recordset
    
    g_oREScheme.Init g_oActiveUser
    Set rsBusInfo = g_oREScheme.GetBusInfoRsReport(dtpQueryDate.Value, ResolveDisplay(cboCheck.Text), ResolveDisplay(cboSellStation))
    Dim aszTmp() As Variant
    ReDim aszTmp(1 To 3, 1 To 2)
    aszTmp(1, 1) = "检票口"
    aszTmp(1, 2) = ResolveDisplay(cboCheck)
    aszTmp(2, 1) = "日期"
    aszTmp(2, 2) = Format(dtpQueryDate.Value, "yyyy年MM月dd日")
    aszTmp(3, 1) = "星期"
    aszTmp(3, 2) = WeekdayName(Weekday(dtpQueryDate.Value))
    ShowReport rsBusInfo, "公司行车记录表.xls", "公司行车记录表", aszTmp
End Sub
'

'得到客务公司道路营运车辆安全门检记录表
Private Sub CompanyVechileSafeInfo()
On Error GoTo ErrHandle
Dim moScheme As New REScheme
Dim aszEnvBus() As Variant
Dim i As Integer
Dim rsTemp As New Recordset
Dim rsTmp As New Recordset
Dim szFindString As String


Dim aszTmp As Variant
ReDim aszTmp(1 To 1, 1 To 2)
aszTmp(1, 1) = "日期"
aszTmp(1, 2) = Format(dtpQueryDate.Value, "yyyy年MM月dd日")
szFindString = " AND ebi.status= 0 "
Set rsTemp = moScheme.GetRESellStationBusReport(ResolveDisplay(cboSellStation), dtpQueryDate.Value, szFindString)
ShowReport rsTemp, "公司道路营运安全门检记录表.xls", "公司道路营运安全门检记录表", aszTmp

Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'公司日行车计划
Private Sub CompanyDayWorkPlan()
On Error GoTo ErrHandle
    Dim g_oREScheme As New REScheme       '当前检票对象
    '得到当天所有车次信息
    Dim rsBusInfo As Recordset
    Dim aszTmp As Variant
    ReDim aszTmp(1 To 1, 1 To 2)
    aszTmp(1, 1) = "日期"
    aszTmp(1, 2) = Format(dtpQueryDate.Value, "yyyy年MM月dd日")
    Set rsBusInfo = g_oREScheme.GetRESellStationBusReport(ResolveDisplay(cboSellStation), dtpQueryDate.Value)
    ShowReport rsBusInfo, "日服务作业计划.xls", "日服务作业计划", aszTmp
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'公司总行车计划
Private Sub CompanyWorkPlan()
On Error GoTo ErrHandle
    Dim g_oREScheme As New REScheme       '当前检票对象
    '得到所有车次信息
    Dim rsBusInfo As Recordset

    Dim g_oTicketPriceMan As New TicketPriceMan
    Dim szFindString As String
    
    szFindString = " AND  bpl.price_table_id= '" & g_szExePriceTable & "'"
    
    
    Set rsBusInfo = g_oREScheme.GetPlanSellStationBusReport(ResolveDisplay(cboSellStation), szFindString)
    ShowReport rsBusInfo, "月服务作业计划.xls", "月服务作业计划"
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub


'填充上车站
Private Sub FillSellStation()
    Dim i As Integer
    cboSellStation.Clear
    For i = 1 To ArrayLength(g_atAllSellStation)
        cboSellStation.AddItem MakeDisplayString(g_atAllSellStation(i).szSellStationID, g_atAllSellStation(i).szSellStationFullName)
    Next i
    If cboSellStation.ListCount > 0 Then cboSellStation.ListIndex = 0
End Sub

Public Function ShowReport(prsData As Recordset, pszFileName As String, pszCaption As String, Optional pvaCustomData As Variant, Optional pnReportType As Integer = 0) As Long
    On Error GoTo Error_Handle
    Me.Caption = pszCaption
    WriteProcessBar True, , , "正在形成报表..."
    RTReport.SheetTitle = ""

    RTReport.TemplateFile = App.Path & "\" & pszFileName
    RTReport.LeftLabelVisual = True
    RTReport.TopLabelVisual = True
    RTReport.ShowReport prsData, pvaCustomData
    WriteProcessBar False, , , ""
    ShowSBInfo "共" & prsData.RecordCount & "条记录", ESB_ResultCountInfo
    Exit Function
Error_Handle:
    ShowErrorMsg
End Function
