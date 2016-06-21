VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#3.1#0"; "RTReportLF.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{695ABF14-B2D8-11D2-A5ED-DE08DCF33612}#1.2#0"; "asfcombo.ocx"
Begin VB.Form frmQueryAccept 
   BackColor       =   &H00E0E0E0&
   Caption         =   "行包受理单查询"
   ClientHeight    =   6150
   ClientLeft      =   4020
   ClientTop       =   3075
   ClientWidth     =   10290
   ControlBox      =   0   'False
   HelpContextID   =   7000090
   Icon            =   "frmQueryAccept.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   10290
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ptQuery 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   60
      ScaleHeight     =   5985
      ScaleWidth      =   2805
      TabIndex        =   2
      Tag             =   $"frmQueryAccept.frx":038A
      Top             =   30
      Width           =   2835
      Begin VB.OptionButton optQueryMode5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "指定运输单号"
         Height          =   285
         Left            =   150
         TabIndex        =   18
         Top             =   4020
         Width           =   2400
      End
      Begin VB.OptionButton optQueryMode6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "指定托运人"
         Height          =   285
         Left            =   150
         TabIndex        =   19
         Top             =   4290
         Width           =   2400
      End
      Begin VB.TextBox txtSheetID 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   150
         TabIndex        =   16
         Top             =   780
         Visible         =   0   'False
         Width           =   2490
      End
      Begin FCmbo.asFlatCombo cboAcceptStatus 
         Height          =   270
         Left            =   150
         TabIndex        =   15
         Top             =   780
         Visible         =   0   'False
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   476
         ButtonDisabledForeColor=   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonHotBackColor=   8421504
         ButtonPressedBackColor=   0
         Text            =   ""
         ButtonBackColor =   8421504
         Style           =   1
         OfficeXPColors  =   -1  'True
      End
      Begin VB.OptionButton OptQueryMode2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "指定受理单号"
         Height          =   285
         Left            =   150
         TabIndex        =   14
         Top             =   3270
         Width           =   1770
      End
      Begin VB.OptionButton OptQueryMode1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "指定站点"
         Height          =   285
         Left            =   150
         TabIndex        =   7
         Top             =   3030
         Value           =   -1  'True
         Width           =   1770
      End
      Begin VB.OptionButton OptQueryMode3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "指定行包状态"
         Height          =   285
         Left            =   150
         TabIndex        =   5
         Top             =   3510
         Width           =   2400
      End
      Begin VB.TextBox txtAcceptID 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   150
         MaxLength       =   10
         TabIndex        =   3
         Top             =   780
         Visible         =   0   'False
         Width           =   2490
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   2400
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   63176707
         UpDown          =   -1  'True
         CurrentDate     =   37062
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   63176707
         UpDown          =   -1  'True
         CurrentDate     =   36528
      End
      Begin RTComctl3.FlatLabel FlatLabel1 
         Height          =   285
         Left            =   0
         TabIndex        =   11
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
         TabIndex        =   12
         Top             =   4785
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
         MICON           =   "frmQueryAccept.frx":03D1
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
         TabIndex        =   13
         Top             =   4785
         Width           =   1170
         _ExtentX        =   2064
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
         MICON           =   "frmQueryAccept.frx":03ED
         PICN            =   "frmQueryAccept.frx":0409
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
         Caption         =   "有优惠的受理单"
         Height          =   285
         Left            =   150
         TabIndex        =   17
         Top             =   3765
         Width           =   2400
      End
      Begin FText.asFlatTextBox txtEndStation 
         Height          =   315
         Left            =   150
         TabIndex        =   1
         Top             =   780
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
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
         OfficeXPColors  =   -1  'True
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2640
         Y1              =   4665
         Y2              =   4665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "受理时间(&M):"
         Height          =   180
         Left            =   105
         TabIndex        =   10
         Top             =   1245
         Width           =   1080
      End
      Begin VB.Label lblQueryType 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "到达站代码(&T):"
         Height          =   180
         Left            =   135
         TabIndex        =   9
         Top             =   510
         Width           =   1260
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "查询方式;"
         Height          =   180
         Left            =   150
         TabIndex        =   8
         Top             =   2775
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
         TabIndex        =   22
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


Private Sub cboAcceptStatus_Change()
    cmdQuery.Enabled = True
End Sub

Private Sub cboAcceptStatus_Click()
    cmdQuery.Enabled = True
End Sub

Private Sub cboAcceptStatus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdQuery.SetFocus
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
    ReDim avCustmonData(1 To 3, 1 To 2)
    avCustmonData(1, 1) = "开始日期"
    avCustmonData(1, 2) = dtpStart.Value
    avCustmonData(2, 1) = "结束日期"
    avCustmonData(2, 2) = dtpEnd.Value
    avCustmonData(3, 1) = "查询方式"
    
    
    If optQueryMode1.Value = True Then
        '        If txtEndStation.Text = "" Then
        '            MsgBox "站点不能为空", vbExclamation, "行包受理"
        '            Exit Sub
        '        End If
        Set rsTemp = moLugSvr.GetAcceptSheetRS(CDate(dtpStart), CDate(dtpEnd), , , ResolveDisplay(txtEndStation.Text))
        avCustmonData(3, 2) = optQueryMode1.Caption & "且站点为" & IIf(txtEndStation.Text = "", "空", txtEndStation.Text)
    ElseIf OptQueryMode2.Value = True Then
        If txtAcceptID.Text = "" Then
            '            MsgBox "行包单号不能为空", vbExclamation, "行包受理"
            '            Exit Sub
            Set rsTemp = moLugSvr.GetAcceptSheetRS(CDate(dtpStart), CDate(dtpEnd), , Trim(txtAcceptID.Text))
        Else
            Set rsTemp = moLugSvr.GetAcceptSheetRS(, , , Trim(txtAcceptID.Text))
        End If
        
        avCustmonData(3, 2) = OptQueryMode2.Caption & "且受理单号为" & IIf(txtAcceptID.Text = "", "空", txtAcceptID.Text)
    ElseIf OptQueryMode3.Value = True Then
        If cboAcceptStatus.Text = "" Then
            MsgBox "行包状态不能为空", vbExclamation, "行包受理"
            Exit Sub
        End If
        Select Case cboAcceptStatus.Text
        Case "正常等待签发"
            szStatus = 0
        Case "废票"
            szStatus = 1
        Case "退票"
            szStatus = 2
        Case "已签发"
            szStatus = 3
        End Select
        Set rsTemp = moLugSvr.GetAcceptSheetRS(CDate(dtpStart), CDate(dtpEnd), szStatus)

        avCustmonData(3, 2) = OptQueryMode3.Caption & "且行包状态为" & IIf(cboAcceptStatus.Text = "", "空", cboAcceptStatus.Text)
    ElseIf OptQueryMode4.Value Then
        
        Set rsTemp = moLugSvr.GetAcceptSheetRS(CDate(dtpStart), CDate(dtpEnd), , , , , True)
        avCustmonData(3, 2) = OptQueryMode4.Caption
    ElseIf OptQueryMode5.Value Then
        If Trim(txtAcceptID.Text) = "" Then
              Set rsTemp = moLugSvr.GetAcceptSheetRS(CDate(dtpStart), CDate(dtpEnd))
        Else
            Set rsTemp = moLugSvr.GetAcceptSheetRS(, , , , , , , , Trim(txtAcceptID.Text))
        End If
        
        avCustmonData(3, 2) = OptQueryMode5.Caption & "且运输单号为" & IIf(txtAcceptID.Text = "", "空", txtAcceptID.Text)
    ElseIf optQueryMode6.Value Then
        If Trim(txtAcceptID.Text) = "" Then
            Set rsTemp = moLugSvr.GetAcceptSheetRS(CDate(dtpStart), CDate(dtpEnd))
        Else
            Set rsTemp = moLugSvr.GetAcceptSheetRS(CDate(dtpStart), CDate(dtpEnd), , , , , , , , Trim(txtAcceptID.Text))
        End If
        
        avCustmonData(3, 2) = optQueryMode6.Caption & "且托运人为" & IIf(txtAcceptID.Text = "", "空", txtAcceptID.Text)
    End If
    If rsTemp.RecordCount > 0 Then
        ShowReport rsTemp, "行包受理单明细报表.xls", "行包受理单明细报表", avCustmonData
    Else
        MsgBox "符合条件的信息不存在！", vbInformation, "行包受理"
    End If
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

Private Sub flbDetailInfo_Click()
'If lvInfo.ListItems.Count > 0 Then
    '  frmLugDetail.cmdAddNew.Enabled = False
    '  frmLugDetail.cmdDelete.Enabled = False
    '  frmLugDetail.cmdEdit.Enabled = False
'      frmAcceptInfo.LuggageID = Trim(lvInfo.SelectedItem.Text)
'      frmAcceptInfo.Show vbModal

'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyF1 Then
        DisplayHelp Me
    End If
End Sub

Private Sub Form_Load()
dtpStart.Value = Date
dtpEnd.Value = DateAdd("s", -1, DateAdd("d", 1, dtpStart.Value))
FormClear
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

Private Sub lvInfo_DblClick()
flbDetailInfo_Click
End Sub

Private Sub OptQueryMode1_Click()
    txtSheetID.Visible = False
    txtEndStation.Visible = True
    cboAcceptStatus.Visible = False
    txtAcceptID.Visible = False
    lblQueryType.Caption = "到达站代码(&T):"
End Sub

Private Sub OptQueryMode2_Click()
    txtSheetID.Visible = False
    txtEndStation.Visible = False
    cboAcceptStatus.Visible = False
    txtAcceptID.Visible = True
    lblQueryType.Caption = "行包单代码(&T):"
    
End Sub

Private Sub OptQueryMode3_Click()
    txtSheetID.Visible = False
    txtEndStation.Visible = False
    cboAcceptStatus.Visible = True
    cboAcceptStatus.clear
    cboAcceptStatus.AddItem "正常等待签发"
    cboAcceptStatus.AddItem "废票"
    cboAcceptStatus.AddItem "退票"
    cboAcceptStatus.AddItem "已签发"
    cboAcceptStatus.ListIndex = 0
    txtAcceptID.Visible = False
    lblQueryType.Caption = "行包单状态(&T):"
End Sub

Private Sub OptQueryMode4_Click()
    txtSheetID.Visible = True
    txtEndStation.Visible = False
    cboAcceptStatus.Visible = False
    txtAcceptID.Visible = False
    lblQueryType.Caption = "签发单代码(&T):"
End Sub
Private Sub Form_Activate()
'    SetSheetNoLabel True, g_szAcceptSheetID
    
    mdiMain.SetPrintEnabled True
    
    If optQueryMode1.Value Then
'        txtEndStation.SetFocus
    ElseIf OptQueryMode2.Value Then
        txtAcceptID.SetFocus
    ElseIf OptQueryMode3.Value Then
        cboAcceptStatus.SetFocus
    ElseIf OptQueryMode5.Value Then
        txtAcceptID.SetFocus
    ElseIf optQueryMode6.Value Then
        txtAcceptID.SetFocus
    Else
        txtSheetID.SetFocus
    End If
End Sub

Private Sub Form_Deactivate()
'    HideSheetNoLabel
    
    mdiMain.SetPrintEnabled False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
'    HideSheetNoLabel
End Sub

Private Sub ptResult_Resize()
On Error Resume Next
    Const cnMargin = 80
'    fraInfo.Move cnMargin - 15, cnMargin
    flInfo.Move cnMargin, cnMargin, ptResult.ScaleWidth - 2 * cnMargin, ptResult.ScaleHeight - 3 * cnMargin
End Sub

Private Sub txtAcceptID_Change()
cmdQuery.Enabled = True
End Sub

Private Sub txtAcceptID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdQuery.SetFocus
    End If
End Sub

Private Sub txtEndStation_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init m_oAUser
    aszTemp = oShell.SelectStation()
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtEndStation.Text = aszTemp(1, 1) & "[" & Trim(aszTemp(1, 2)) & "]"
    cmdQuery.Enabled = True
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

Private Sub txtEndStation_Change()
cmdQuery.Enabled = True
End Sub

Private Sub txtEndStation_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdQuery.SetFocus
    End If
End Sub

Private Sub txtSheetID_Change()
cmdQuery.Enabled = True
End Sub

Private Sub txtSheetID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdQuery.SetFocus
    End If
End Sub

Public Function ShowReport(prsData As Recordset, pszFileName As String, pszCaption As String, Optional pvaCustomData As Variant, Optional pnReportType As Integer = 0) As Long
    On Error GoTo Error_Handle
    Me.ZOrder 0
    Me.Show
    
    Dim arsTemp As Variant
    Dim aszTemp As Variant
'    Dim rsTemp As Recordset
    ReDim aszTemp(1 To 1)
    ReDim arsTemp(1 To 1)
    '赋票种
    aszTemp(1) = "托运费项"
    Set arsTemp(1) = g_rsPriceItem
    
'    Me.Caption = pszCaption
    
    WriteProcessBar True, , , "正在形成报表..."
    
    flInfo.CustomStringCount = aszTemp
    flInfo.CustomString = arsTemp
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
Private Sub OptQueryMode5_Click()
    txtSheetID.Visible = False
    txtEndStation.Visible = False
    cboAcceptStatus.Visible = False
    txtAcceptID.Visible = True
    txtAcceptID.SetFocus
    lblQueryType.Caption = "运输单代码(&T):"
End Sub

Private Sub optQueryMode6_Click()
    txtSheetID.Visible = False
    txtEndStation.Visible = False
    cboAcceptStatus.Visible = False
    txtAcceptID.Visible = True
    txtAcceptID.SetFocus
    lblQueryType.Caption = "托运人名称(&T):"
End Sub
