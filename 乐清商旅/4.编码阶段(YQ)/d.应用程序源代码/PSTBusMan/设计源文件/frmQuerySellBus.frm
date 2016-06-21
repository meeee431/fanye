VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#2.1#0"; "rtreportlf.ocx"
Begin VB.Form frmQuerySellBus 
   BackColor       =   &H00E0E0E0&
   Caption         =   "统计各车次座位售票情况"
   ClientHeight    =   6645
   ClientLeft      =   1365
   ClientTop       =   3465
   ClientWidth     =   10725
   ForeColor       =   &H8000000B&
   Icon            =   "frmQuerySellBus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   10725
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   8940
      Top             =   3420
   End
   Begin VB.PictureBox ptResult 
      BackColor       =   &H80000009&
      Height          =   6030
      Left            =   3150
      ScaleHeight     =   5970
      ScaleWidth      =   5355
      TabIndex        =   10
      Top             =   210
      Width           =   5415
      Begin VB.PictureBox ptQ 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1155
         Left            =   0
         Picture         =   "frmQuerySellBus.frx":000C
         ScaleHeight     =   1155
         ScaleWidth      =   5640
         TabIndex        =   11
         Top             =   0
         Width           =   5640
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "车次座位售票情况"
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
            Left            =   870
            TabIndex        =   12
            Top             =   750
            Width           =   1920
         End
         Begin VB.Image imgOpen 
            Height          =   240
            Left            =   0
            Picture         =   "frmQuerySellBus.frx":0F02
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin RTReportLF.RTReport RTReport 
         Height          =   3225
         Left            =   240
         TabIndex        =   13
         Top             =   2205
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   5689
      End
   End
   Begin VB.PictureBox ptQuery 
      BackColor       =   &H80000009&
      Height          =   6345
      Left            =   90
      ScaleHeight     =   6285
      ScaleWidth      =   2790
      TabIndex        =   0
      Top             =   120
      Width           =   2850
      Begin VB.ComboBox cboType 
         Height          =   300
         ItemData        =   "frmQuerySellBus.frx":104C
         Left            =   90
         List            =   "frmQuerySellBus.frx":104E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1260
         Width           =   2625
      End
      Begin RTComctl3.CoolButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   330
         Left            =   1530
         TabIndex        =   1
         Top             =   2370
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
         MICON           =   "frmQuerySellBus.frx":1050
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton flbClose 
         Height          =   240
         Left            =   2490
         TabIndex        =   2
         Top             =   15
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   423
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
         MICON           =   "frmQuerySellBus.frx":106C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.FlatLabel FlatLabel1 
         Height          =   285
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   2790
         _ExtentX        =   4921
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
         BackColor       =   -2147483633
         VerticalAlignment=   1
         HorizontalAlignment=   1
         MarginLeft      =   0
         MarginTop       =   0
         NormTextColor   =   -2147483630
         Caption         =   "查询条件设定"
      End
      Begin MSComCtl2.DTPicker dtpStartTime 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-M-d"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   5
         Top             =   675
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   24969216
         CurrentDate     =   37022
      End
      Begin RTComctl3.CoolButton cmdQuery 
         Default         =   -1  'True
         Height          =   345
         Left            =   150
         TabIndex        =   6
         Top             =   2340
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "查询(&Q)"
         ENAB            =   0   'False
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
         MICON           =   "frmQuerySellBus.frx":1088
         PICN            =   "frmQuerySellBus.frx":10A4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin FText.asFlatTextBox txtObject 
         Height          =   300
         Left            =   90
         TabIndex        =   15
         Top             =   1875
         Width           =   2625
         _ExtentX        =   4630
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
         ButtonPressedBackColor=   -2147483627
         Text            =   ""
         ButtonBackColor =   -2147483633
         ButtonVisible   =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起始时间(&D):"
         Height          =   180
         Left            =   90
         TabIndex        =   9
         Top             =   450
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查询方式(&S):"
         Height          =   180
         Left            =   90
         TabIndex        =   8
         Top             =   1050
         Width           =   1080
      End
      Begin VB.Label lbltile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "站点代码(&P):"
         Height          =   180
         Left            =   90
         TabIndex        =   7
         Top             =   1665
         Width           =   1080
      End
   End
   Begin RTComctl3.Spliter spQuery 
      Height          =   1320
      Left            =   2970
      TabIndex        =   14
      Top             =   2010
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   2328
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
Attribute VB_Name = "frmQuerySellBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const cszTemplateFile = "车次座位售票情况模板.xls"
Const cnTop = 1200

Private m_lMoveLeft As Long
Private m_lRange As Long '写进度条用

Private Sub Form_Activate()
    Form_Resize
    MDIScheme.SetPrintEnabled True
End Sub

Private Sub Form_Deactivate()
    MDIScheme.SetPrintEnabled False
End Sub
Private Sub cboStation_Change()
    EnabledQuery
End Sub
Private Sub cboStation_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        FillFoundStation
    End If
End Sub
Private Sub cmdQuery_Click()
    m_lRange = 0
    Query
End Sub
Private Sub FillType()
    '填充查询类型
    cboType.AddItem "按站点查询"
    cboType.AddItem "按车次查询"
    cboType.AddItem "按线路查询"
'    cboType.AddItem "按检票口查询"
    cboType.ListIndex = 0
End Sub
Private Sub Query()
    '查询车次座位售票统计
    Dim rsTemp As Recordset
    Dim oTemp As New STDss.TicketUnitDim
    Dim i As Integer
    Dim j As Integer
    Dim bflg As Boolean
    Dim dyEndDate As Date
    dyEndDate = DateAdd("d", 1, dtpStartTime.Value)
    On Error Resume Next
    oTemp.Init g_oActiveUser
    Select Case cboType.ListIndex
    Case 0  '按站点查询
        Set rsTemp = oTemp.GetBusDateSeatSellInfo(dtpStartTime.Value, dyEndDate, ResolveDisplay(txtObject.Text))
    Case 1 '按车次查询
        Set rsTemp = oTemp.GetBusDateFromBusIDSeatSellInfo(dtpStartTime.Value, dyEndDate, ResolveDisplay(txtObject.Text))
    Case 2  '按线路查询
        Set rsTemp = oTemp.GetBusDateFromRoutIdSeatSellInfo(dtpStartTime.Value, dyEndDate, ResolveDisplay(txtObject.Text))
    End Select
    
    Dim rsReport As New Recordset
    rsReport.CursorLocation = adUseClient
    rsReport.Fields.Append "bus_start_time", adVarChar, 50
    rsReport.Fields.Append "bus_id", adVarChar, 50
    rsReport.Fields.Append "bus_type_name", adVarChar, 50
    rsReport.Fields.Append "vehicle_type_name", adVarChar, 50
    rsReport.Fields.Append "end_station_name", adVarChar, 50
    rsReport.Fields.Append "total_seat", adVarChar, 50
    rsReport.Fields.Append "sale_seat_quantity", adVarChar, 50
    rsReport.Fields.Append "have_sale_seat_quantity", adVarChar, 50
    rsReport.Fields.Append "seat_type_name", adVarChar, 50
    rsReport.Fields.Append "leave_seat", adVarChar, 200
    rsReport.Fields.Append "ticket_num1", adVarChar, 200
    rsReport.Fields.Append "ticket_num2", adVarChar, 200
    rsReport.Fields.Append "ticket_num3", adVarChar, 200
    rsReport.Fields.Append "ticket_num4", adVarChar, 200
    rsReport.Fields.Append "ticket_num5", adVarChar, 200
    rsReport.Fields.Append "ticket_id", adVarChar, 200
    rsReport.Fields.Append "status_name", adVarChar, 200
    rsReport.Fields.Append "sell_station_id", adVarChar, 200
'    rsReport.Fields.Append "ticket_num6", adVarChar, 200
    rsReport.Open
    
    For j = 1 To rsTemp.RecordCount
'        If Not rsReport.BOF Then rsReport.MovePrevious
'        If Not rsReport.BOF Then
'            If rsReport.Fields("bus_id") = FormatDbValue(rsTemp!bus_id) Then
'                rsReport.Fields("seat_type_name") = rsReport.Fields("seat_type_name") & "/" & FormatDbValue(rsTemp!seat_type_name)
'                rsReport.Fields("ticket_num1") = rsReport.Fields("ticket_num1") & "/" & FormatDbValue(rsTemp!full_price)
'                rsReport.Fields("ticket_num2") = rsReport.Fields("ticket_num2") & "/" & FormatDbValue(rsTemp!half_price)
'                rsReport.Fields("ticket_num3") = rsReport.Fields("ticket_num3") & "/" & FormatDbValue(rsTemp!preferential_ticket1)
'                rsReport.Fields("ticket_num4") = rsReport.Fields("ticket_num4") & "/" & FormatDbValue(rsTemp!preferential_ticket2)
'                rsReport.Fields("ticket_num5") = rsReport.Fields("ticket_num5") & "/" & FormatDbValue(rsTemp!preferential_ticket3)
'    '            rsReport.Fields("ticket_num6") = ""
'                rsReport.Update
'            Else
'                rsReport.MoveNext
'                rsReport.AddNew
'                rsReport.Fields("bus_start_time") = Format(rsTemp!bus_start_time, "hh:mm")
'                rsReport.Fields("bus_id") = FormatDbValue(rsTemp!bus_id)
'                rsReport.Fields("bus_type_name") = FormatDbValue(rsTemp!bus_type_name)
'                rsReport.Fields("vehicle_type_name") = FormatDbValue(rsTemp!vehicle_type_short_name)
'                rsReport.Fields("end_station_name") = FormatDbValue(rsTemp!end_station_name)
'                rsReport.Fields("total_seat") = FormatDbValue(rsTemp!TotalSeat)
'                rsReport.Fields("sale_seat_quantity") = FormatDbValue(rsTemp!sale_seat_quantity)
'                rsReport.Fields("have_sale_seat_quantity") = FormatDbValue(rsTemp!have_sale_seat_quantity)
'                rsReport.Fields("seat_type_name") = FormatDbValue(rsTemp!seat_type_name)
'                rsReport.Fields("leave_seat") = FormatDbValue(rsTemp!leave_seat_id)
'                rsReport.Fields("ticket_num1") = FormatDbValue(rsTemp!full_price)
'                rsReport.Fields("ticket_num2") = FormatDbValue(rsTemp!half_price)
'                rsReport.Fields("ticket_num3") = FormatDbValue(rsTemp!preferential_ticket1)
'                rsReport.Fields("ticket_num4") = FormatDbValue(rsTemp!preferential_ticket2)
'                rsReport.Fields("ticket_num5") = FormatDbValue(rsTemp!preferential_ticket3)
'    '            rsReport.Fields("ticket_num6") = ""
'                rsReport.Update
'            End If
'        Else
'            If Not rsReport.EOF Then rsReport.MoveNext
            rsReport.AddNew
            rsReport.Fields("bus_start_time") = Format(rsTemp!bus_start_time, "hh:mm")
            rsReport.Fields("bus_id") = FormatDbValue(rsTemp!bus_id)
            rsReport.Fields("bus_type_name") = FormatDbValue(rsTemp!bus_type_name)
            rsReport.Fields("vehicle_type_name") = FormatDbValue(rsTemp!vehicle_type_short_name)
            rsReport.Fields("end_station_name") = FormatDbValue(rsTemp!end_station_name)
            rsReport.Fields("total_seat") = FormatDbValue(rsTemp!TotalSeat)
            rsReport.Fields("sale_seat_quantity") = FormatDbValue(rsTemp!sale_seat_quantity)
            rsReport.Fields("have_sale_seat_quantity") = FormatDbValue(rsTemp!have_sale_seat_quantity)
            rsReport.Fields("seat_type_name") = FormatDbValue(rsTemp!seat_type_name)
            rsReport.Fields("leave_seat") = FormatDbValue(rsTemp!leave_seat_id)
            rsReport.Fields("ticket_num1") = FormatDbValue(rsTemp!full_price)
            rsReport.Fields("ticket_num2") = FormatDbValue(rsTemp!half_price)
            rsReport.Fields("ticket_num3") = FormatDbValue(rsTemp!preferential_ticket1)
            rsReport.Fields("ticket_num4") = FormatDbValue(rsTemp!preferential_ticket2)
            rsReport.Fields("ticket_num5") = FormatDbValue(rsTemp!preferential_ticket3)
            rsReport.Fields("ticket_id") = FormatDbValue(rsTemp!ticket_id)
            rsReport.Fields("status_name") = FormatDbValue(rsTemp!status_name)
            rsReport.Fields("sell_station_id") = FormatDbValue(rsTemp!sell_station_name)
'            rsReport.Fields("ticket_num6") = ""
            rsReport.Update
'        End If
        rsTemp.MoveNext
    Next j
    If rsReport.RecordCount > 0 Then rsReport.MoveFirst
    
    RTReport.LeftLabelVisual = True
    RTReport.TopLabelVisual = True
    RTReport.TemplateFile = App.Path & "\" & cszTemplateFile
    RTReport.ShowReport rsReport
        
    WriteProcessBar False
    ShowSBInfo
    ShowSBInfo "共有" & m_lRange & "条记录", ESB_ResultCountInfo
    
    Set rsTemp = Nothing
    Set oTemp = Nothing
    Exit Sub
ErrorHandle:
    ShowSBInfo
    WriteProcessBar False
    Set rsTemp = Nothing
    Set oTemp = Nothing
    ShowErrorMsg
End Sub

Private Sub EnabledQuery()
    '查询按钮是否可用
    If Trim(txtObject.Text) = "" Then
        cmdQuery.Enabled = False
    Else
        cmdQuery.Enabled = True
    End If
End Sub
'
Private Sub cboType_Click()
    '初始化
    txtObject.Text = ""
    SetLabelCaption
End Sub
'
Private Sub FillFoundStation()
'    '填充查找到的站点
'    AddCboStation cboStation
'    If cboStation.ListCount = 1 Then
'        Query
'    End If
End Sub


Private Sub txtObject_ButtonClick()
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init g_oActiveUser
        
    Select Case cboType.ListIndex
        Case 0  '站点
            aszTmp = oShell.SelectStation
        Case 1  '车次
            aszTmp = oShell.SelectBus
        Case 2
            aszTmp = oShell.SelectRoute
    End Select
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtObject.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))

End Sub

Private Sub txtObject_Change()
    EnabledQuery
End Sub

'Private Sub vsQuery_DblClick()
'    DisplayBusInfo
'End Sub
'
'Private Sub DisplayBusInfo()
'    '显示车次售票信息
'    On Error GoTo ErrorHandle
'
'    Dim ofrm As Object
'    Dim szBusID As String
'    Dim szRundate As String
'    szBusID = VsQuery.TextMatrix(VsQuery.Row, 2)
'    szRundate = dtpStartTime.Value
'    If szBusID = "" Then
'        MsgBox "请选择一个次,然后再试试", vbInformation + vbOKOnly, Me.Caption
'        Exit Sub
'    End If
'    frmQuerySellSeat.m_szBusID = szBusID
'    frmQuerySellSeat.m_dyBusData = szRundate
'    frmQuerySellSeat.Show vbModal
'    Exit Sub
'ErrorHandle:
'    ShowErrorMsg
'End Sub



'初始化vsQuery
Private Sub InitGrid()
    Dim oSysParam As New SystemParam
    oSysParam.Init g_oActiveUser
    '得到所有的票种
    
    Dim aszTmp(1 To 1) As Variant
    Dim arsTmp(1 To 1) As Variant
    aszTmp(1) = "票种"
    Set arsTmp(1) = oSysParam.GetAllTicketTypeRS(TP_TicketTypeValid)

    RTReport.CustomString = arsTmp
    RTReport.CustomStringCount = aszTmp

End Sub

Private Sub SetLabelCaption()
    Select Case cboType.ListIndex
    Case 0  '按站点查询
        lbltile.Caption = "站点代码(&P):"
    Case 1 '按车次查询
        lbltile.Caption = "车次代码(&P):"
    Case 2  '按线路查询
        lbltile.Caption = "线路代码(&P):"
    End Select

End Sub

Private Sub Form_Load()
    spQuery.InitSpliter ptQuery, ptResult
    m_lMoveLeft = 0
    
    dtpStartTime.Value = Date

    FillType
    InitGrid
    SetLabelCaption

End Sub
Private Sub Form_Resize()
    '重绘
    spQuery.LayoutIt
End Sub

Private Sub ptQuery_Resize()
    flbClose.Move ptQuery.ScaleWidth - 250
End Sub
Private Sub ptResult_Resize()
    Dim lTemp As Long
    lTemp = IIf((ptResult.ScaleHeight - cnTop) <= 0, lTemp, ptResult.ScaleHeight - cnTop)
    RTReport.Move 0, cnTop, ptResult.ScaleWidth, lTemp
    'lblTitle.Move 60 + m_lMoveLeft, 50, ptResult.ScaleWidth

    FlatLabel1.Width = ptQuery.ScaleWidth
    flbClose.Left = FlatLabel1.Left + FlatLabel1.Width - flbClose.Width - 30
End Sub

Private Sub cmdCancel_Click()
    '关闭
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MDIScheme.SetPrintEnabled False
End Sub
Private Sub RTReport_SetProgressRange(ByVal lRange As Variant)
    m_lRange = lRange
    
End Sub

Private Sub RTReport_SetProgressValue(ByVal lValue As Variant)
    WriteProcessBar True, lValue, m_lRange, "正在填充报表数据..."
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    
End Sub
Private Sub flbClose_Click()
    '关闭左边部分
    ptQuery.Visible = False
    imgOpen.Visible = True
    m_lMoveLeft = 240
    'lblTitle.Move 60 + m_lMoveLeft
    spQuery.LayoutIt
End Sub

Private Sub imgOpen_Click()
    '打开左边部分
    ptQuery.Visible = True
    imgOpen.Visible = False
    m_lMoveLeft = 0
    'lblTitle.Move 60 + m_lMoveLeft
    spQuery.LayoutIt
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
