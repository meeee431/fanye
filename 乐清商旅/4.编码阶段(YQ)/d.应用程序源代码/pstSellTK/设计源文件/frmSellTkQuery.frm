VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{BBF95DAA-F9CB-4CA9-A673-E0E9E0193752}#1.0#0"; "STSellCtl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmQuerySellTk 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   6135
   ClientLeft      =   2850
   ClientTop       =   3570
   ClientWidth     =   10425
   HelpContextID   =   4000230
   Icon            =   "frmSellTkQuery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   10425
   Begin VB.PictureBox ptResult 
      BackColor       =   &H00E0E0E0&
      Height          =   6060
      Left            =   2940
      ScaleHeight     =   6000
      ScaleWidth      =   7320
      TabIndex        =   11
      Top             =   45
      Width           =   7380
      Begin VB.Frame fraInfo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "车票信息"
         Height          =   825
         Left            =   60
         TabIndex        =   12
         Top             =   120
         Width           =   7155
         Begin RTComctl3.FloatLabel flbBusInfo 
            Height          =   315
            HelpContextID   =   3001401
            Left            =   5700
            TabIndex        =   8
            Top             =   450
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HoverBackColor  =   -2147483633
            NormTextColor   =   16711680
            Caption         =   "详细信息(D)..."
            NormUnderline   =   -1  'True
         End
         Begin VB.Label lblTkStatus 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "已退"
            Height          =   180
            Left            =   4140
            TabIndex        =   29
            Top             =   525
            Width           =   1320
         End
         Begin VB.Label label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "状态;"
            Height          =   180
            Index           =   6
            Left            =   3585
            TabIndex        =   28
            Top             =   525
            Width           =   450
         End
         Begin VB.Label label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "车次代码;"
            Height          =   180
            Index           =   4
            Left            =   1665
            TabIndex        =   24
            Top             =   255
            Width           =   810
         End
         Begin VB.Label lblBusid 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "999-2"
            Height          =   180
            Left            =   2490
            TabIndex        =   23
            Top             =   255
            Width           =   690
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "车次日期;"
            Height          =   180
            Index           =   3
            Left            =   3555
            TabIndex        =   22
            Top             =   270
            Width           =   810
         End
         Begin VB.Label lblBusDate 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "2000-11-30"
            Height          =   180
            Left            =   4395
            TabIndex        =   21
            Top             =   270
            Width           =   930
         End
         Begin VB.Label lblTkType 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "全票"
            Height          =   180
            Left            =   615
            TabIndex        =   18
            Top             =   510
            Width           =   360
         End
         Begin VB.Label lblTkPrice 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "152.03"
            Height          =   180
            Left            =   2490
            TabIndex        =   17
            Top             =   510
            Width           =   720
         End
         Begin VB.Label lblTkID 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "T000000111"
            Height          =   180
            Left            =   615
            TabIndex        =   16
            Top             =   255
            Width           =   900
         End
         Begin VB.Label label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "票种;"
            Height          =   180
            Index           =   5
            Left            =   150
            TabIndex        =   15
            Top             =   510
            Width           =   900
         End
         Begin VB.Label label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "票款;"
            Height          =   180
            Index           =   2
            Left            =   1650
            TabIndex        =   14
            Top             =   510
            Width           =   450
         End
         Begin VB.Label label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "票号;"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   13
            Top             =   255
            Width           =   450
         End
      End
      Begin MSComctlLib.ListView lvInfo 
         Height          =   4905
         Left            =   75
         TabIndex        =   27
         Top             =   1065
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   8652
         SortKey         =   1
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "TicketNum"
            Text            =   "票号"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "BusID"
            Text            =   "车次代码"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "BusDate"
            Text            =   "票款"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "TicketType"
            Text            =   "票种"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "TicketPrice"
            Text            =   "车票状态"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Key             =   "TicketStatus"
            Text            =   "车次日期"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox ptQuery 
      BackColor       =   &H00E0E0E0&
      Height          =   6045
      Left            =   30
      ScaleHeight     =   5985
      ScaleWidth      =   2775
      TabIndex        =   0
      Top             =   45
      Width           =   2835
      Begin VB.TextBox txtTicketID 
         Height          =   315
         Left            =   135
         TabIndex        =   7
         Top             =   735
         Width           =   2490
      End
      Begin VB.OptionButton OptQueryMode3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "指定票号"
         Height          =   285
         Left            =   90
         TabIndex        =   6
         Top             =   3225
         Value           =   -1  'True
         Width           =   1020
      End
      Begin MSComCtl2.DTPicker dtpSellTime 
         Height          =   315
         Left            =   150
         TabIndex        =   2
         Top             =   1425
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62324738
         CurrentDate     =   36528
      End
      Begin VB.ComboBox cboTicketType 
         Height          =   300
         ItemData        =   "frmSellTkQuery.frx":014A
         Left            =   135
         List            =   "frmSellTkQuery.frx":015A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   750
         Width           =   2505
      End
      Begin VB.OptionButton optQueryMode1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "指定时间"
         Height          =   285
         Left            =   105
         TabIndex        =   4
         Top             =   2910
         Width           =   1020
      End
      Begin VB.OptionButton optQueryMode2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "时间段"
         Height          =   240
         Left            =   1485
         TabIndex        =   5
         Top             =   2955
         Width           =   1080
      End
      Begin RTComctl3.FlatLabel FlatLabel1 
         Height          =   285
         Left            =   0
         TabIndex        =   25
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
         HorizontalAlignment=   1
         Caption         =   "查询条件设定"
      End
      Begin STSellCtl.ucUpDownText txtTimeInterval 
         Height          =   330
         Left            =   165
         TabIndex        =   3
         Top             =   2145
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
         SelectOnEntry   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Max             =   100
         Value           =   "5"
      End
      Begin RTComctl3.CoolButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   345
         Left            =   1440
         TabIndex        =   30
         Top             =   3660
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
         MICON           =   "frmSellTkQuery.frx":0182
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
         Left            =   120
         TabIndex        =   31
         Top             =   3660
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
         MICON           =   "frmSellTkQuery.frx":019E
         PICN            =   "frmSellTkQuery.frx":01BA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "分钟之前"
         Height          =   180
         Index           =   2
         Left            =   1830
         TabIndex        =   26
         Top             =   2235
         Width           =   720
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "查询方式(Q);"
         Height          =   180
         Left            =   150
         TabIndex        =   20
         Top             =   2655
         Width           =   1080
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "售票时间段(N):"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   1875
         Width           =   1260
      End
      Begin VB.Label label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "车票类型(&T):"
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   510
         Width           =   1080
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "起始售票时间(&S):"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   1185
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmQuerySellTk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const QueryMode1 = 0
Const QueryMode2 = 1
Const QueryMode3 = 2
Dim moActUser As ActiveUser



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdQuery_Click()
    '票号:1140.095 车次代码:1260.284           票款:1035.213 票种:1440     车票状态:1170.142           车次日期:1124.787
    On Error GoTo Error_Handle
    If OptQueryMode3.Value Then
        Dim oTmp As New CommDialog
        oTmp.Init moActUser
        oTmp.ShowTicketInfo txtTicketID.Text
        Set oTmp = Nothing
        Exit Sub
    End If

    MousePointer = vbHourglass
    Dim rsTemp As Recordset
    Dim vaTemp As Variant
    If optQueryMode1.Value Then
        vaTemp = SelfGetFullDateTime(m_oParam.NowDate, dtpSellTime.Value)
    Else
        vaTemp = CInt(txtTimeInterval.Value)
    End If
    Select Case cboTicketType.ListIndex
        Case 0      '正常售票
            Set rsTemp = m_oSell.GetSellTicketRs(vaTemp)
        Case 1      '改签票
            Set rsTemp = m_oSell.GetChangeTicketRs(vaTemp)
        Case 2      '已退票
            Set rsTemp = m_oSell.GetReturnTicketRs(vaTemp)
        Case 3      '已废票
            Set rsTemp = m_oSell.GetCancelTicketRs(vaTemp)
    End Select

    WriteProcessBar True
    Dim nStep As Long, nLen As Long, i As Long
    nLen = rsTemp.RecordCount
    nStep = Int(nLen / 25) + 1

    lvInfo.ListItems.Clear
    Dim liTemp As ListItem
    Dim nTemp As Integer, szTemp As String
    For i = 1 To nLen
        Set liTemp = lvInfo.ListItems.Add(, GetEncodedKey(rsTemp!ticket_id), rsTemp!ticket_id)
        liTemp.ListSubItems.Add , , rsTemp!bus_id
        liTemp.ListSubItems.Add , , Format(rsTemp!ticket_price, "#0.00")
        liTemp.ListSubItems.Add , , GetTicketTypeStr2(rsTemp!ticket_type)
        liTemp.ListSubItems.Add , , GetTicketStatusStr(rsTemp!status)
        liTemp.ListSubItems.Add , , rsTemp!bus_date
'        liTemp.ListSubItems.Add , , ToStandardDateTimeStr(rsTemp！operation_time)

'        Select Case cboTicketType.ListIndex
'            Case 0      '正常售票
'            Case 1      '改签票
'                liTemp.subitems.Add , , rsTemp！former_ticket_id
'                liTemp.subitems.Add , , rsTemp！former_ticket_price
'                liTemp.subitems.Add , , rsTemp！change_charge
'                liTemp.subitems.Add , , rsTemp！credence_sheet_id
'            Case 2      '已退票
'                liTemp.subitems.Add , , ToStandardDateTimeStr(rsTemp！return_time)
'                liTemp.subitems.Add , , rsTemp！return_charge
'                liTemp.subitems.Add , , rsTemp！credence_sheet_id
'                liTemp.subitems.Add , , GetReturnTicketTypeStr(rsTemp！return_mode)
'            Case 3      '已废票
'                liTemp.subitems.Add , , ToStandardDateTimeStr(rsTemp！cancel_time)
'                liTemp.subitems.Add , , GetCancelTicketTypeStr(rsTemp！cancel_mode)
'        End Select

         WriteProcessBar , i, nLen
        rsTemp.MoveNext
    Next i
    Set liTemp = Nothing
    If nLen > 0 Then
        lvInfo.ListItems(1).Selected = True
        '查看详细记录有效
    End If
    WriteProcessBar False
    ShowSBInfo "共" & nLen & "条记录", ESB_ResultCountInfo
    MousePointer = vbNormal

    If lvInfo.ListItems.count = 0 Then
        MsgBox "符合条件的车次信息不存在！", vbInformation, "车次查询"
    Else
        lvInfo.SetFocus
    End If
    Set rsTemp = Nothing
    Exit Sub
Error_Handle:
    WriteProcessBar False
    Me.MousePointer = vbNormal
    ShowErrorMsg
    Set rsTemp = Nothing
    If Not liTemp Is Nothing Then Set liTemp = Nothing
End Sub

Private Sub dtpSellTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdQuery.SetFocus
    End If
End Sub

Private Sub flbBusInfo_Click()
    Dim oTmp As New CommDialog
    oTmp.Init moActUser
    Dim pTicketID As String
    pTicketID = lvInfo.SelectedItem.Text
    oTmp.ShowTicketInfo pTicketID
    Set oTmp = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 And Chr(KeyCode) = "D" And flbBusInfo.Enabled Then
        flbBusInfo_Click
    End If
'    If (Chr(KeyCode) = "D" Or Chr(KeyCode) = "d") And Shift = 2 And flbBusInfo.Enabled Then
'        flbBusInfo_Click
'    End If
End Sub

Private Sub Form_Load()
    InitForm
    AlignHeadWidth Me.name, lvInfo
    cboTicketType.ListIndex = 0
End Sub

Public Property Let SelfUser(oActUser As Object)
    Set moActUser = oActUser
End Property
Private Sub InitForm()
    layFormByQueryMode QueryMode1
    dtpSellTime.Value = Time
    optQueryMode1.Value = True
    writeTkSummery
End Sub


Private Sub writeTkSummery()
    Dim pBusDate As Date            '参数
    Dim pBusId As String, pBusSerialNo As Integer
    Dim tCheckBusInfo As TBusCheckInfo  '检票车次信息
    If lvInfo.ListItems.count = 0 Then
        lblTkID.Caption = ""
        lblBusID.Caption = ""
        lblBusDate.Caption = ""
        lblTkPrice.Caption = ""
        lblTkStatus.Caption = ""
        lblTkType.Caption = ""
        flbBusInfo.Enabled = False
    Else
        If Not lvInfo.SelectedItem Is Nothing Then
            lblTkID.Caption = Trim(lvInfo.SelectedItem.Text)
            lblBusID.Caption = Trim(lvInfo.SelectedItem.SubItems(1))
            lblBusDate.Caption = lvInfo.SelectedItem.SubItems(2)
            lblTkType.Caption = lvInfo.SelectedItem.SubItems(3)
            lblTkPrice.Caption = lvInfo.SelectedItem.SubItems(4)
            lblTkStatus.Caption = lvInfo.SelectedItem.SubItems(5)
            flbBusInfo.Enabled = True
        End If
    End If
End Sub

Private Sub layFormByQueryMode(nQueryMode As Integer)
'************************************************************
'根据查询模式排列控列，模糊方式和指定方式
'************************************************************

    If nQueryMode = QueryMode1 Then
        txtTicketID.Visible = False
        cboTicketType.Visible = True
        Label1(0).Caption = "售票类型:"

        dtpSellTime.Enabled = True
        txtTimeInterval.Enabled = False
    ElseIf nQueryMode = QueryMode2 Then
        txtTicketID.Visible = False
        cboTicketType.Visible = True
        Label1(0).Caption = "售票类型:"

        dtpSellTime.Enabled = False
        txtTimeInterval.Enabled = True
    Else
        txtTicketID.Visible = True
        cboTicketType.Visible = False
        Label1(0).Caption = "车票号:"

        dtpSellTime.Enabled = False
        txtTimeInterval.Enabled = False
    End If
End Sub

Private Sub DisplayHelp(Optional HelpType As EHelpType = content)
    Dim lActiveControl As Long

    Select Case HelpType
        Case content
            lActiveControl = Me.ActiveControl.HelpContextID
            If lActiveControl = 0 Then
                TopicID = Me.HelpContextID
                CallHTMLShowTopicID
            Else
                TopicID = lActiveControl
                CallHTMLShowTopicID
            End If
        Case Index
            CallHTMLHelpIndex
        Case Support
            TopicID = clSupportID
            CallHTMLShowTopicID
    End Select

End Sub

Private Sub Form_Activate()
    If optQueryMode1.Value Then
        dtpSellTime.SetFocus
    ElseIf optQueryMode2.Value Then
        txtTimeInterval.SetFocus
    Else
        txtTicketID.SetFocus
    End If
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    Const cnMargin = 50
    ptQuery.Move cnMargin, cnMargin, ptQuery.Width, Me.ScaleHeight - 2 * cnMargin
    ptResult.Move cnMargin + ptQuery.Width + 2 * cnMargin, cnMargin, Me.ScaleWidth - ptQuery.Width - 4 * cnMargin, Me.ScaleHeight - 2 * cnMargin
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveHeadWidth Me.name, lvInfo
    ShowSBInfo "", ESB_ResultCountInfo
End Sub

Private Sub lvInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvInfo, ColumnHeader.Index
End Sub

Private Sub lvInfo_DblClick()
    flbBusInfo_Click
End Sub

Private Sub lvInfo_ItemClick(ByVal Item As MSComctlLib.ListItem)
    writeTkSummery
End Sub

Private Sub optQueryMode1_Click()
    layFormByQueryMode QueryMode1
'    dtpSellTime.SetFocus
End Sub

Private Sub optQueryMode2_Click()
    layFormByQueryMode QueryMode2
    txtTimeInterval.SetFocus
End Sub

Private Sub OptQueryMode3_Click()
    layFormByQueryMode QueryMode3
    txtTicketID.SetFocus
End Sub

Private Sub ptResult_Resize()
On Error Resume Next
    Const cnMargin = 80
    fraInfo.Move cnMargin - 15, cnMargin
    lvInfo.Move cnMargin, fraInfo.Top + fraInfo.Height + cnMargin, ptResult.ScaleWidth - 2 * cnMargin, ptResult.ScaleHeight - 3 * cnMargin - fraInfo.Height
End Sub

Private Sub txtTicketID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim szTemp As String
    Dim lTemp As Long
    On Error GoTo Error_Handel
    If Len(txtTicketID.Text) >= TicketNoNumLen() Then
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
            lTemp = Right(txtTicketID.Text, TicketNoNumLen())
            szTemp = Left(txtTicketID.Text, Len(txtTicketID.Text) - TicketNoNumLen())

            If KeyCode = vbKeyUp Then
                lTemp = lTemp + 1
            Else
                lTemp = lTemp - 1
                lTemp = IIf(lTemp > 0, lTemp, 0)
            End If
            txtTicketID.Text = MakeTicketNo(lTemp, szTemp)
            KeyCode = 0
        End If
    End If
    Exit Sub
Error_Handel:

End Sub

Private Sub txtTicketID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdQuery.SetFocus
    End If
End Sub

Private Sub txtTimeInterval_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdQuery.SetFocus
    End If
End Sub
