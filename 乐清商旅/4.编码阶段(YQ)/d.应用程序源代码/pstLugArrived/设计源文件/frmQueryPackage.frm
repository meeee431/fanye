VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#2.1#0"; "RTReportLF.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmQueryAccept 
   BackColor       =   &H00E0E0E0&
   Caption         =   "行包到达受理单统计"
   ClientHeight    =   6465
   ClientLeft      =   2070
   ClientTop       =   2670
   ClientWidth     =   10290
   ControlBox      =   0   'False
   HelpContextID   =   7000090
   Icon            =   "frmQueryPackage.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   10290
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ptQuery 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   60
      ScaleHeight     =   6105
      ScaleWidth      =   2805
      TabIndex        =   1
      Top             =   60
      Width           =   2835
      Begin VB.OptionButton optQueryMode9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "车站行包到达月报"
         Height          =   285
         Left            =   150
         TabIndex        =   28
         Top             =   5100
         Width           =   2400
      End
      Begin VB.OptionButton optQueryMode8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "车站营收简报"
         Height          =   285
         Left            =   150
         TabIndex        =   27
         Top             =   4830
         Width           =   2400
      End
      Begin VB.OptionButton optQueryMode7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "明细查询"
         Height          =   285
         Left            =   150
         TabIndex        =   26
         Top             =   4565
         Width           =   2400
      End
      Begin VB.TextBox txtLicense 
         Height          =   300
         Left            =   150
         MaxLength       =   10
         TabIndex        =   25
         Top             =   1080
         Width           =   2490
      End
      Begin VB.OptionButton optQueryMode6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "按车号统计"
         Height          =   285
         Left            =   150
         TabIndex        =   24
         Top             =   4300
         Width           =   2400
      End
      Begin VB.ComboBox cboOperator 
         Height          =   300
         Left            =   150
         TabIndex        =   23
         Top             =   1080
         Width           =   2490
      End
      Begin VB.ComboBox cboLoader 
         Height          =   300
         Left            =   150
         TabIndex        =   22
         Top             =   1080
         Width           =   2490
      End
      Begin VB.TextBox txtStartStation 
         Height          =   300
         Left            =   150
         MaxLength       =   10
         TabIndex        =   21
         Top             =   1080
         Width           =   2490
      End
      Begin VB.ComboBox cboAreaType 
         Height          =   300
         Left            =   150
         TabIndex        =   20
         Top             =   1080
         Width           =   2490
      End
      Begin VB.ComboBox cboStatus 
         Height          =   300
         ItemData        =   "frmQueryPackage.frx":038A
         Left            =   1320
         List            =   "frmQueryPackage.frx":038C
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   450
         Width           =   1395
      End
      Begin VB.OptionButton optQueryMode5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "代收运费统计"
         Height          =   285
         Left            =   150
         TabIndex        =   14
         Top             =   5955
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.OptionButton OptQueryMode2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "按站点统计"
         Height          =   285
         Left            =   150
         TabIndex        =   12
         Top             =   3505
         Width           =   1770
      End
      Begin VB.OptionButton OptQueryMode1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "按地区统计"
         Height          =   285
         Left            =   150
         TabIndex        =   5
         Top             =   3240
         Value           =   -1  'True
         Width           =   1770
      End
      Begin VB.OptionButton OptQueryMode3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "按受理人统计"
         Height          =   285
         Left            =   150
         TabIndex        =   3
         Top             =   3770
         Width           =   2400
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   2670
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
         _Version        =   393216
         Format          =   25559040
         CurrentDate     =   37062
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   2070
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
         _Version        =   393216
         Format          =   25559040
         CurrentDate     =   36528
      End
      Begin RTComctl3.FlatLabel FlatLabel1 
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   2805
         _ExtentX        =   4948
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
         HorizontalAlignment=   1
         Caption         =   "查询条件设定"
      End
      Begin RTComctl3.CoolButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   345
         Left            =   1470
         TabIndex        =   10
         Top             =   5550
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "关闭"
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
         MICON           =   "frmQueryPackage.frx":038E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdQuery 
         Default         =   -1  'True
         Height          =   345
         Left            =   150
         TabIndex        =   11
         Top             =   5550
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "统计(&Q)"
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
         MICON           =   "frmQueryPackage.frx":03AA
         PICN            =   "frmQueryPackage.frx":03C6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.OptionButton OptQueryMode4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "按装卸工统计"
         Height          =   285
         Left            =   150
         TabIndex        =   13
         Top             =   4035
         Width           =   2400
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2640
         Y1              =   5460
         Y2              =   5460
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "提货状态(&T)"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   510
         Width           =   990
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2430
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1830
         Width           =   975
      End
      Begin VB.Label lblDateTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "受理时间(&M):"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   105
         TabIndex        =   8
         Top             =   1515
         Width           =   1080
      End
      Begin VB.Label lblQueryType 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "标识(&T):"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "查询方式;"
         Height          =   180
         Left            =   150
         TabIndex        =   6
         Top             =   3045
         Width           =   900
      End
   End
   Begin VB.PictureBox ptResult 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6060
      Left            =   2940
      ScaleHeight     =   6060
      ScaleWidth      =   7380
      TabIndex        =   0
      Top             =   0
      Width           =   7380
      Begin RTReportLF.RTReport flInfo 
         Height          =   5985
         Left            =   60
         TabIndex        =   17
         Top             =   60
         Width           =   7230
         _ExtentX        =   12753
         _ExtentY        =   10557
      End
   End
End
Attribute VB_Name = "frmQueryAccept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'Last Modify By: 陆勇庆  2005-8-16
'Last Modify In:增加两种查询模式（按运输单号和按客户）
'*******************************************************************************
Option Explicit


Private mlMaxCount As Long


Private Sub cboStatus_Click()
    If optQueryMode7.Value Then     '明细查询时都按照到达时间来
        lblDateTitle.Caption = "行包到达时间(&M):"
        Exit Sub
    End If
    
    
    If cboStatus.Text = CPick_Picked Then
        lblDateTitle.Caption = "行包提取时间(&M):"
    Else
        lblDateTitle.Caption = "行包到达时间(&M):"
    End If
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdQuery_Click()
    QueryAcceptLuggage
End Sub
Private Sub QueryAcceptLuggage()
    On Error GoTo ErrHandle
    Dim rsTemp As Recordset
    Dim szStatus As Integer
    
    Dim avCustmonData As Variant
    ReDim avCustmonData(1 To 5, 1 To 2)
    avCustmonData(1, 1) = "填表单位"
    avCustmonData(1, 2) = g_oActUser.UserUnitName
    avCustmonData(2, 1) = "填表日期"
    avCustmonData(2, 2) = Format(Date, cszLongDateFormat)
    avCustmonData(3, 1) = "填表人"
    avCustmonData(3, 2) = g_oActUser.UserID 'cboOperator.Text

    avCustmonData(4, 1) = "统计单位"
    avCustmonData(4, 2) = g_oActUser.UserUnitName
    avCustmonData(5, 1) = "统计时间"
    If dtpStart.Value = dtpEnd.Value Then
        avCustmonData(5, 2) = Format(dtpStart.Value, cszLongDateFormat)
    Else
        avCustmonData(5, 2) = Format(dtpStart.Value, cszLongDateFormat) & " 至 " & Format(dtpEnd.Value, cszLongDateFormat)
    End If
    
    Dim szTemplateFile As String, szSearch As String, szTotalBy As String
    szSearch = " 1=1 "

    Select Case cboStatus.Text
    Case CPick_Picked
        szSearch = szSearch & " AND p.status=" & EPS_Picked
    Case CPick_Normal
        szSearch = szSearch & " AND p.status=" & EPS_Normal
    Case CPick_Canceled
        szSearch = szSearch & " AND p.status=" & EPS_Cancel
    Case Else
    End Select
    If OptQueryMode1.Value = True Then
        If Trim(cboAreaType.Text) <> "" Then
            szSearch = szSearch & "  AND area_type LIKE '%" & Trim(cboAreaType.Text) & "%'"
        End If
        szTemplateFile = "到达行包统计_按地区"
        szTotalBy = "area_type"
    ElseIf OptQueryMode2.Value = True Then
        If Trim(txtStartStation.Text) <> "" Then
            szSearch = szSearch & " AND start_station_name LIKE '%" & Trim(txtStartStation.Text) & "%'"
        End If
        szTemplateFile = "到达行包统计_按站点"
        szTotalBy = "area_type,start_station_name"
    ElseIf OptQueryMode3.Value = True Then
        If Trim(cboOperator.Text) <> "" Then
            szSearch = szSearch & " AND user_name LIKE '%" & Trim(cboOperator.Text) & "%'"
        End If
        szTemplateFile = "到达行包统计_按受理人"
        szTotalBy = "t.user_name"
    ElseIf OptQueryMode4.Value Then
        If Trim(cboLoader.Text) <> "" Then
            szSearch = szSearch & " AND loader LIKE '%" & Trim(cboLoader.Text) & "%'"
        End If
        szTemplateFile = "到达行包统计_按装卸工"
        szTotalBy = "loader"
    ElseIf optQueryMode6.Value = True Then
        If Trim(txtLicense.Text) <> "" Then
            szSearch = szSearch & " AND license_tag_no LIKE '%" & Trim(txtLicense.Text) & "%'"
        End If
        szTemplateFile = "到达行包统计_按车号"
        szTotalBy = "license_tag_no"
    Else
    
    End If
    
    
    If optQueryMode7.Value Then
        '明细查询
        If Trim(cboOperator.Text) <> "" Then
            szSearch = szSearch & " AND user_name LIKE '%" & Trim(cboOperator.Text) & "%'"
        End If
        szTemplateFile = "到达行包明细查询"
        Set rsTemp = g_oPackageSvr.GetArrivedPackageRS(dtpStart, dtpEnd, "", szSearch)
    ElseIf optQueryMode8.Value Then
        szTemplateFile = "车站到达行包营收简报"
        Set rsTemp = g_oPackageSvr.StationStat(dtpStart, dtpEnd)
    ElseIf optQueryMode9.Value Then
        szTemplateFile = "车站到达行包营收月报"
        Set rsTemp = g_oPackageSvr.StationStatMonth(dtpStart, dtpEnd)
    
    ElseIf cboStatus.Text = "已提" Then
        Set rsTemp = g_oPackageSvr.StatPackageByPickedRS(dtpStart, dtpEnd, szTotalBy, szSearch)
    Else
        Set rsTemp = g_oPackageSvr.StatPackageByArrivedRS(dtpStart, dtpEnd, szTotalBy, szSearch)
    End If
    
    ShowReport rsTemp, szTemplateFile & ".xls", szTemplateFile, avCustmonData
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub FillAcceptLuggage(rsTemp As Recordset)
'    Dim Count As Integer
'    Dim i As Integer
'    Dim liTemp As ListItem
'    Dim mStatus As String
''    lvInfo.ListItems.Clear
'    '添充lvInfo
'    For i = 1 To rsTemp.RecordCount
''        Set liTemp = lvInfo.ListItems.Add(, , FormatDbValue(rsTemp!luggage_id))
''        liTemp.SubItems(1) = GetLuggageTypeString(FormatDbValue(rsTemp!accept_type))
''        liTemp.SubItems(2) = FormatDbValue(rsTemp!start_station_name)
''        liTemp.SubItems(3) = FormatDbValue(rsTemp!luggage_name)
''        liTemp.SubItems(4) = FormatDbValue(rsTemp!Mileage)
''        liTemp.SubItems(5) = FormatDbValue(rsTemp!bus_id)
''        liTemp.SubItems(6) = FormatDbValue(rsTemp!bus_date)
''        liTemp.SubItems(7) = FormatDbValue(rsTemp!cal_weight)
''        liTemp.SubItems(8) = FormatDbValue(rsTemp!fact_weight)
''        liTemp.SubItems(9) = FormatDbValue(rsTemp!start_label_id)
''        liTemp.SubItems(10) = FormatDbValue(rsTemp!baggage_number)
''        liTemp.SubItems(11) = FormatDbValue(rsTemp!over_weight_number)
''        liTemp.SubItems(12) = FormatDbValue(rsTemp!price_total) '托运费
''        Select Case FormatDbValue(rsTemp!Status)
''            Case 0
''                mStatus = "正常等待签发"
''            Case 1
''                mStatus = "废票"
''            Case 2
''                mStatus = "退票"
''            Case 3
''                mStatus = "已签发"
''        End Select
''        liTemp.SubItems(13) = mStatus
''        liTemp.SubItems(14) = FormatDbValue(rsTemp!Shipper)
''        liTemp.SubItems(15) = FormatDbValue(rsTemp!Picker)
''        liTemp.SubItems(16) = FormatDbValue(rsTemp!des_station_name)
'        rsTemp.MoveNext
'    Next

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyF1 Then
        DisplayHelp Me
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    InitDate1
    FormClear

    FillBaseInfo
    cboStatus_Click

    OptQueryMode1_Click
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub InitDate2()
    dtpStart.Value = Format(DateAdd("m", -1, Now), "yyyy-mm-01")
    dtpEnd.Value = DateAdd("d", -1, Format(Now, "yyyy-mm-01"))
End Sub

Private Sub InitDate1()
    dtpStart.Value = Date
    dtpEnd.Value = Date
End Sub

Private Sub FormClear()
'  lblAcceptID.Caption = ""
'  lblDesStationName.Caption = ""
'  lblNumber.Caption = ""
'  lblStatus.Caption = ""
'  lblTotalPrice = ""
'  lblType.Caption = ""
'  cmdQuery.Enabled = False
End Sub

Private Sub Form_Resize()
 On Error Resume Next
    Const cnMargin = 50
    ptQuery.Move cnMargin, cnMargin, ptQuery.Width, Me.ScaleHeight - 2 * cnMargin
    ptResult.Move cnMargin + ptQuery.Width + 2 * cnMargin, cnMargin, Me.ScaleWidth - ptQuery.Width - 4 * cnMargin, Me.ScaleHeight - 2 * cnMargin
End Sub

Private Sub OptQueryMode1_Click()
    lblDateTitle.Visible = True
    lblTag.Visible = True
    cboStatus.Visible = True
    cboAreaType.Visible = True
    cboLoader.Visible = False
    cboOperator.Visible = False
    txtStartStation.Visible = False
    txtLicense.Visible = False
    InitDate1
    lblQueryType.Caption = "指定地区(&T):"
End Sub

Private Sub OptQueryMode2_Click()
    lblDateTitle.Visible = True
    lblTag.Visible = True
    cboStatus.Visible = True
    cboAreaType.Visible = False
    cboLoader.Visible = False
    cboOperator.Visible = False
    txtStartStation.Visible = True
    txtLicense.Visible = False
    InitDate1
    lblQueryType.Caption = "指定站点(&T):"

End Sub

Private Sub OptQueryMode3_Click()
    lblDateTitle.Visible = True
    lblTag.Visible = True
    cboStatus.Visible = True
    cboAreaType.Visible = False
    cboLoader.Visible = False
    cboOperator.Visible = True
    txtStartStation.Visible = False
    txtLicense.Visible = False
    InitDate1
    lblQueryType.Caption = "指定受理人(&T):"
End Sub

Private Sub OptQueryMode4_Click()
    lblDateTitle.Visible = True
    lblTag.Visible = True
    cboStatus.Visible = True
    cboAreaType.Visible = False
    cboLoader.Visible = True
    cboOperator.Visible = False
    txtStartStation.Visible = False
    txtLicense.Visible = False
    InitDate1
    lblQueryType.Caption = "指定装卸工(&T):"
End Sub

Private Sub OptQueryMode6_Click()
    lblDateTitle.Visible = True
    cboAreaType.Visible = False
    cboLoader.Visible = False
    cboOperator.Visible = False
    txtStartStation.Visible = False
    txtLicense.Visible = True
    lblTag.Visible = True
    cboStatus.Visible = True
    InitDate1
    lblQueryType.Caption = "指定车号(&T):"
End Sub

Private Sub optQueryMode7_Click()
    lblDateTitle.Visible = True
    lblTag.Visible = True
    cboStatus.Visible = True
    cboAreaType.Visible = False
    cboLoader.Visible = False
    cboOperator.Visible = True
    txtStartStation.Visible = False
    txtLicense.Visible = False
    InitDate1
    lblQueryType.Caption = "指定受理人(&T):"
End Sub

Private Sub OptQueryMode8_Click()
    cboAreaType.Visible = False
    cboLoader.Visible = False
    cboOperator.Visible = False
    txtStartStation.Visible = False
    txtLicense.Visible = False
    lblDateTitle.Visible = False
    lblTag.Visible = False
    cboStatus.Visible = False
    InitDate1
    lblQueryType.Caption = "车站到达行包营收简报"
End Sub

Private Sub optQueryMode9_Click()
    cboAreaType.Visible = False
    cboLoader.Visible = False
    cboOperator.Visible = False
    txtStartStation.Visible = False
    txtLicense.Visible = False
    lblDateTitle.Visible = False
    lblTag.Visible = False
    cboStatus.Visible = False
    InitDate2
    lblQueryType.Caption = "车站到达行包营收月报"
End Sub

Private Sub Form_Activate()

    mdiMain.SetPrintEnabled True

    If OptQueryMode1.Value Then
         cboAreaType.SetFocus
    ElseIf OptQueryMode2.Value Then
        txtStartStation.SetFocus
    ElseIf OptQueryMode3.Value Then
        cboOperator.SetFocus
    ElseIf OptQueryMode4.Value Then
        cboLoader.SetFocus
    ElseIf optQueryMode6.Value Then
        txtLicense.SetFocus
    End If
End Sub

Private Sub Form_Deactivate()

    mdiMain.SetPrintEnabled False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Sub ptResult_Resize()
On Error Resume Next
    Const cnMargin = 80
'    fraInfo.Move cnMargin - 15, cnMargin
    flInfo.Move cnMargin, cnMargin, ptResult.ScaleWidth - 2 * cnMargin, ptResult.ScaleHeight - 3 * cnMargin
End Sub

Public Function ShowReport(prsData As Recordset, pszFileName As String, pszCaption As String, Optional pvaCustomData As Variant, Optional pnReportType As Integer = 0) As Long
    On Error GoTo Error_Handle
    Me.ZOrder 0
    Me.Show


'    Me.Caption = pszCaption

    WriteProcessBar True, , , "正在形成报表..."

    flInfo.SheetTitle = ""

    flInfo.TemplateFile = App.Path & "\" & pszFileName
    flInfo.LeftLabelVisual = True
    flInfo.TopLabelVisual = True
    flInfo.ShowReport prsData, pvaCustomData
    WriteProcessBar False, , , ""
    ShowSBInfo "共" & prsData.RecordCount & "条记录", ESB_ResultCountInfo
    Exit Function
Error_Handle:
    ShowErrorMsg
End Function




Public Sub ExportToFile()

    flInfo.OpenDialog EDialogType.EXPORT_FILE
End Sub

Public Sub PrintReport(Optional pbShowDialog As Boolean = False)
    On Error GoTo ErrorHandle
    flInfo.PrintReport pbShowDialog
    Exit Sub
ErrorHandle:
End Sub

Public Sub PreView()
    flInfo.PrintView
End Sub

Public Sub PageSet()
    flInfo.OpenDialog EDialogType.PAGESET_TYPE
End Sub

Public Sub PrintSet()
    flInfo.OpenDialog EDialogType.PRINT_TYPE
End Sub
'导出文件
Public Sub ExportFile()
    Dim szFileName As String
    szFileName = flInfo.OpenDialog(EDialogType.EXPORT_FILE)
End Sub
'导出文件并打开
Public Sub ExportFileOpen()
    Dim szFileName As String
    szFileName = flInfo.OpenDialog(EDialogType.EXPORT_FILE)
    If szFileName <> "" Then
        OpenLinkedFile szFileName
    End If
End Sub




Private Sub flInfo_SetProgressRange(ByVal lRange As Variant)
    mlMaxCount = lRange
End Sub

Private Sub flInfo_SetProgressValue(ByVal lValue As Variant)
    WriteProcessBar True, lValue, mlMaxCount
End Sub
'填充数据字典
Private Sub FillBaseInfo()
    cboStatus.Clear
    cboStatus.AddItem CSZNoneString
    cboStatus.AddItem CPick_Normal
    cboStatus.ItemData(1) = EPS_Normal
    cboStatus.AddItem CPick_Picked
    cboStatus.ItemData(2) = EPS_Picked
    cboStatus.AddItem CPick_Canceled
    cboStatus.ItemData(3) = EPS_Cancel
    cboStatus.ListIndex = 0
    
    Dim aszTmp() As String
    aszTmp = g_oPackageParam.ListBaseDefine(EDefineType.EDT_AreaType)
    cboAreaType.Clear
    Dim i As Integer
    For i = 1 To ArrayLength(aszTmp)
        cboAreaType.AddItem aszTmp(i, 3)
    Next i


    aszTmp = g_oPackageParam.ListBaseDefine(EDefineType.EDT_LoadWorker)
    cboLoader.Clear
    For i = 1 To ArrayLength(aszTmp)
        cboLoader.AddItem aszTmp(i, 3)
    Next i

    aszTmp = g_oPackageParam.ListBaseDefine(EDefineType.EDT_Operator)
    cboOperator.Clear
    For i = 1 To ArrayLength(aszTmp)
        cboOperator.AddItem aszTmp(i, 3)
    Next i


End Sub
