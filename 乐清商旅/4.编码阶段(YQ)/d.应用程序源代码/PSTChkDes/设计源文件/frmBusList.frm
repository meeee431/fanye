VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmBusList 
   Caption         =   "检票车次"
   ClientHeight    =   7530
   ClientLeft      =   1350
   ClientTop       =   1170
   ClientWidth     =   10875
   HelpContextID   =   4003001
   Icon            =   "frmBusList.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   10875
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ptLeft 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   6585
      Left            =   90
      ScaleHeight     =   6585
      ScaleWidth      =   7155
      TabIndex        =   1
      Top             =   405
      Width           =   7155
      Begin VB.PictureBox ptDown 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   690
         ScaleHeight     =   2655
         ScaleWidth      =   6210
         TabIndex        =   3
         Top             =   3630
         Width           =   6210
         Begin MSComctlLib.ListView lvCheckedBus 
            CausesValidation=   0   'False
            Height          =   2640
            Left            =   345
            TabIndex        =   5
            Top             =   -285
            Width           =   6795
            _ExtentX        =   11986
            _ExtentY        =   4657
            SortKey         =   3
            View            =   3
            LabelEdit       =   1
            SortOrder       =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "imlBusStatus"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "车次"
               Object.Width           =   1614
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "发车时间"
               Object.Width           =   1614
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "开检时间"
               Object.Width           =   1614
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "停检时间"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "路单号码"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "到站"
               Object.Width           =   1799
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "参营公司"
               Object.Width           =   1879
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "车辆"
               Object.Width           =   2090
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "车主"
               Object.Width           =   1429
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "状态"
               Object.Width           =   1764
            EndProperty
         End
      End
      Begin VB.PictureBox ptUp 
         BorderStyle     =   0  'None
         Height          =   2625
         Left            =   555
         ScaleHeight     =   2625
         ScaleWidth      =   6225
         TabIndex        =   2
         Top             =   105
         Width           =   6225
         Begin MSComctlLib.ListView lvWillCheckBus 
            CausesValidation=   0   'False
            Height          =   2910
            Left            =   360
            TabIndex        =   4
            Top             =   975
            Width           =   6795
            _ExtentX        =   11986
            _ExtentY        =   5133
            SortKey         =   1
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "imlBusStatus"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "车次"
               Object.Width           =   1614
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "发车时间"
               Object.Width           =   1614
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "到站"
               Object.Width           =   1799
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "参营公司"
               Object.Width           =   1879
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "车辆"
               Object.Width           =   2090
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "车主"
               Object.Width           =   1429
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "状态"
               Object.Width           =   1764
            EndProperty
         End
      End
      Begin RTComctl3.Spliter Spliter1 
         Height          =   165
         Left            =   960
         TabIndex        =   6
         Top             =   3855
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   291
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
         SelectColor     =   16777215
         IsVertical      =   -1  'True
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7650
      Top             =   1740
   End
   Begin MSComctlLib.ImageList imlBusStatus 
      Left            =   7530
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusList.frx":014A
            Key             =   "stop"
            Object.Tag             =   "停检"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusList.frx":02A4
            Key             =   "checked"
            Object.Tag             =   "已检"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusList.frx":03FE
            Key             =   "normal"
            Object.Tag             =   "未检"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusList.frx":0558
            Key             =   "checking"
            Object.Tag             =   "正在检票"
         EndProperty
      EndProperty
   End
   Begin ActiveBar2LibraryCtl.ActiveBar2 abAction 
      Height          =   6465
      Left            =   8280
      TabIndex        =   0
      Top             =   300
      Width           =   1440
      _LayoutVersion  =   1
      _ExtentX        =   2540
      _ExtentY        =   11404
      _DataPath       =   ""
      Bands           =   "frmBusList.frx":06B2
   End
End
Attribute VB_Name = "frmBusList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private eCurrCheckBusStatus As ECheckStatus
Private anIconIndex(1 To 10) As Integer
Private mbIsShow As Boolean '当前窗体是否正显示
'设置按钮的有效性
Private Sub EnableButton(CheckBusStatus As Integer)
    With abAction.Bands("bndActionTabs").ChildBands("actMenu")
    .Tools("mi_BusInfo").Enabled = True
    Select Case CheckBusStatus
        Case ECS_CanotCheck
            .Tools("mi_StartCheck").Enabled = False
            .Tools("mi_ExtraCheck").Enabled = False
            .Tools("mi_CheckInfo").Enabled = False
            .Tools("mi_SheetInfo").Enabled = False
        Case ECS_CanCheck
            .Tools("mi_StartCheck").Enabled = True
            .Tools("mi_ExtraCheck").Enabled = False
            .Tools("mi_CheckInfo").Enabled = False
            .Tools("mi_SheetInfo").Enabled = False
        Case ECS_BeChecking
            .Tools("mi_StartCheck").Enabled = True
            .Tools("mi_ExtraCheck").Enabled = False
            .Tools("mi_CheckInfo").Enabled = True
            .Tools("mi_SheetInfo").Enabled = False
        Case ECS_BeExtraChecking, ECS_CanExtraCheck
            .Tools("mi_StartCheck").Enabled = False
            .Tools("mi_ExtraCheck").Enabled = True
            .Tools("mi_CheckInfo").Enabled = True
            .Tools("mi_SheetInfo").Enabled = True
        Case ECS_Checked
            .Tools("mi_StartCheck").Enabled = False
            .Tools("mi_ExtraCheck").Enabled = False
            .Tools("mi_CheckInfo").Enabled = True
            .Tools("mi_SheetInfo").Enabled = True
    End Select
    End With
End Sub

Private Sub ShowBusInfo()
    On Error GoTo ErrorHandle
    If Not lvWillCheckBus.SelectedItem Is Nothing Then
        Dim oFrmBusInfo As New frmBusInfo
        oFrmBusInfo.SelfUser = g_oActiveUser
        oFrmBusInfo.BusID = lvWillCheckBus.SelectedItem.Text
        oFrmBusInfo.BusDate = Date
        oFrmBusInfo.Show vbModal
    End If
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub
'显示车次检票信息
Private Sub ShowCheckInfo()
    If Not lvCheckedBus.SelectedItem Is Nothing Then
        Dim szBusid As String, nBusSerial As Integer
        szBusid = lvCheckedBus.SelectedItem.Text        '拆分班次得到
        If lvCheckedBus.SelectedItem.SubItems(1) <> g_cszTitleScollBus Then '固定班次
            nBusSerial = 0
        Else
            nBusSerial = Val(LeftAndRight(szBusid, False, "-"))
            szBusid = LeftAndRight(szBusid, True, "-")
        End If
        Dim oFrmCheckInfo As New frmCheckBusInfo
        Set oFrmCheckInfo.g_oActiveUser = g_oActiveUser
        oFrmCheckInfo.mszBusID = szBusid
        oFrmCheckInfo.mnBusSerialNo = nBusSerial
        oFrmCheckInfo.mdtBusDate = Date
        oFrmCheckInfo.Show vbModal
        Set oFrmCheckInfo = Nothing
    End If
End Sub

Private Sub ShowCheckSheetInfo()
    If Not lvCheckedBus.SelectedItem Is Nothing Then
        Dim szSheetID As String
        szSheetID = Trim(lvCheckedBus.SelectedItem.SubItems(4))
        If szSheetID = "" Then Exit Sub
        Dim ofrmTmp As frmCheckSheet
        Set ofrmTmp = New frmCheckSheet
        Set ofrmTmp.g_oActiveUser = g_oActiveUser
        Set ofrmTmp.moChkTicket = g_oChkTicket
        ofrmTmp.mbViewMode = True
        ofrmTmp.mbNoPrintPrompt = True
        ofrmTmp.mbExitAfterPrint = False
        ofrmTmp.mszSheetID = szSheetID
        ofrmTmp.Show vbModal
    End If
End Sub


Public Sub RefreshBus()
    Me.MousePointer = vbHourglass
    ShowSBInfo "正在取得车次列表..."
    
    BuildBusCollection
    
    
    FillBusLst
    ShowSBInfo ""
    Me.MousePointer = vbDefault
End Sub
'开检当前选择的车次
Private Sub StartCheckBus()
    Dim i As Integer
    Dim nCheckLineCount As Integer
    Dim nStatus As Integer
    nStatus = Val(lvWillCheckBus.SelectedItem.Tag)
    Select Case nStatus
        Case ECS_BeChecking
            nCheckLineCount = CheckLineCount
            For i = 1 To nCheckLineCount
                If g_atCheckLine(i).BusID = lvWillCheckBus.SelectedItem.Text Then
                    Exit For
                End If
            Next
            If i <= nCheckLineCount Then
                MDIMain.tbsBusList.Tabs(i).Selected = True
            Else
                '系统出现了异常中断
                frmStartCheck.SetProperty lvWillCheckBus.SelectedItem.Text, False
                frmStartCheck.Show vbModal
            End If
        Case ECS_CanCheck
            frmStartCheck.SetProperty lvWillCheckBus.SelectedItem.Text, False
            frmStartCheck.Show vbModal
    End Select
End Sub
'补检当前选择的车次
Private Sub ExtraCheckBus()
    Dim i As Integer
    Dim nCheckLineCount As Integer
    Dim nStatus As Integer
    If lvCheckedBus.SelectedItem.SubItems(1) = g_cszTitleScollBus Then
        frmStartCheck.SetProperty LeftAndRight(lvCheckedBus.SelectedItem.Text, True, "-"), True, LeftAndRight(lvCheckedBus.SelectedItem.Text, False, "-")
        frmStartCheck.Show vbModal
        Exit Sub
    End If
    nStatus = Val(lvWillCheckBus.SelectedItem.Tag)
    Select Case nStatus
        Case ECS_BeExtraChecking
            nCheckLineCount = CheckLineCount
            For i = 1 To nCheckLineCount
                If g_atCheckLine(i).BusID = lvWillCheckBus.SelectedItem.Text Then
                    Exit For
                End If
            Next
            If i <= nCheckLineCount Then
                MDIMain.tbsBusList.Tabs(i).Selected = True
            Else        '异常中断恢复
                frmStartCheck.SetProperty lvWillCheckBus.SelectedItem.Text, True
                frmStartCheck.Show vbModal
            End If
        Case ECS_CanExtraCheck, ECS_Checked
            frmStartCheck.SetProperty lvWillCheckBus.SelectedItem.Text, True
            frmStartCheck.Show vbModal
'        Case
'            MsgboxEx "车次已过补检时间!", vbExclamation, Me.Caption
    End Select
End Sub
Private Function GetStatusName(CheckBusStatus As Integer) As String
    Select Case CheckBusStatus
        Case ECS_CanotCheck
            GetStatusName = "不能开检"
        Case ECS_CanCheck
            GetStatusName = "可以开检"
        Case ECS_BeChecking
            GetStatusName = "正在检票"
        Case ECS_Checked, ECS_CanExtraCheck
            GetStatusName = "车次停检"
        Case ECS_BeExtraChecking
            GetStatusName = "正在补检"
    End Select
End Function

Private Sub abAction_BandClose(ByVal Band As ActiveBar2LibraryCtl.Band)
    If Band.name = "bndActionTabs" Then
        abAction.Visible = False
        Call Form_Resize
    End If
End Sub

Private Sub abAction_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
On Error GoTo ErrHandle
    Select Case Tool.name
        Case "mi_StartCheck"
            StartCheckBus
        Case "mi_ExtraCheck"
            ExtraCheckBus
        Case "mi_BusInfo"
            ShowBusInfo
        Case "mi_CheckInfo"
            ShowCheckInfo
        Case "mi_SheetInfo"
            ShowCheckSheetInfo
        Case "mi_Refresh"
            RefreshBus
        Case "mi_Close"
            Unload Me
    End Select
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Activate()
    mbIsShow = True
    Spliter1.LayoutIt
    Call Form_Resize
End Sub

Private Sub Form_Load()
    AlignHeadWidth Me.name, lvWillCheckBus
    AlignHeadWidth Me.name, lvCheckedBus
    
    Spliter1.InitSpliter ptUp, ptDown

    With abAction.Bands("bndActionTabs").ChildBands("actMenu")
        .Tools("mi_StartCheck").Enabled = False
        .Tools("mi_ExtraCheck").Enabled = False
        .Tools("mi_BusInfo").Enabled = False
        .Tools("mi_CheckInfo").Enabled = False
        .Tools("mi_SheetInfo").Enabled = False
    End With
    
    '车次状态的图标在imageList中的位置
    anIconIndex(ECS_CanotCheck) = 1
    anIconIndex(ECS_BeChecking) = 4
    anIconIndex(ECS_BeExtraChecking) = 4
    anIconIndex(ECS_CanCheck) = 3
    anIconIndex(ECS_Checked) = 2
    anIconIndex(ECS_CanExtraCheck) = 2
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    Const cnMargin = 50
    
    '当操作条关闭时间处理
    ptLeft.Move cnMargin, cnMargin, Me.ScaleWidth - IIf(abAction.Visible, abAction.Width, 0) - 2 * cnMargin, Me.ScaleHeight - 2 * cnMargin
    If abAction.Visible Then
        abAction.Move ptLeft.Width + 2 * cnMargin, ptLeft.Top, abAction.Width, ptLeft.Height
    End If
    ptUp.Move 0, 0, Me.ScaleWidth
    ptDown.Move 0, 0, Me.ScaleWidth
    Spliter1.LayoutIt
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveHeadWidth Me.name, lvWillCheckBus
    SaveHeadWidth Me.name, lvCheckedBus
    
    
'    If mbIsExit Then
        If g_nCurrLineIndex <> 0 Then
'            OpenCurrCheckLine
            MDIMain.tbsBusList.Tabs(g_nCurrLineIndex).Selected = True
        End If
'    End If
    mbIsShow = False
End Sub
Private Sub lvCheckedBus_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvCheckedBus, ColumnHeader.Index
End Sub

Private Sub lvCheckedBus_DblClick()
    Dim oHit As ListItem
'    Set oHit = lvCheckedBus.HitTest(CurrentX, CurrentY)
    Set oHit = lvCheckedBus.SelectedItem
    If Not oHit Is Nothing Then
        oHit.Selected = True
        ShowCheckInfo
    End If
End Sub


Private Sub lvCheckedBus_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim nStatus As Integer
    nStatus = Val(Item.Tag)
    EnableButton Val(Item.Tag)
End Sub

Private Sub lvWillCheckBus_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvWillCheckBus, ColumnHeader.Index
End Sub

Private Sub lvWillCheckBus_DblClick()
    Dim oHit As ListItem
'    Set oHit = lvWillCheckBus.HitTest(CurrentX, CurrentY)
   Set oHit = lvWillCheckBus.SelectedItem
    If Not oHit Is Nothing Then
        oHit.Selected = True
        ShowBusInfo
    End If
End Sub

Private Sub lvWillCheckBus_ItemClick(ByVal Item As MSComctlLib.ListItem)
    EnableButton Val(Item.Tag)
End Sub
'填充车次列表
Private Sub FillBusLst()
    Dim i As Integer
    
    lvWillCheckBus.ListItems.Clear
    lvCheckedBus.ListItems.Clear
    '填充未检车次列表
    For i = 1 To g_cWillCheckBusList.Count
        UpdateWillCheckBusItem 1, g_cWillCheckBusList.Item(i)       '添加一行
    Next i
'    lvWillCheckBus.Refresh
    
    '填充已检车次
    For i = 1 To g_cCheckedBusList.Count
        UpdateCheckedBusItem 1, g_cCheckedBusList.Item(i)  '添加一行
    Next i
    ShowSBInfo g_cWillCheckBusList.Count & "个未检车次 " & g_cCheckedBusList.Count & "个已检车次", ESB_ResultCountInfo
    lvWillCheckBus.SortKey = 1
    lvWillCheckBus.Sorted = True
    lvCheckedBus.SortKey = 3
    lvCheckedBus.Sorted = True
End Sub


Public Property Get IsShow() As Boolean
    IsShow = mbIsShow
End Property

Private Function GetBusCheckStatus(ptCheckBus As tCheckBusLstInfo) As Integer
    '判断指定车次状态
    Dim dptCheckBus As Date
    Dim lHaveTime As Long
    Dim nResult As ECheckStatus
    Select Case ptCheckBus.Status
        Case EREBusStatus.ST_BusChecking
            nResult = ECS_BeChecking
        Case EREBusStatus.ST_BusExtraChecking
            nResult = ECS_BeExtraChecking
        Case EREBusStatus.ST_BusReplace, EREBusStatus.ST_BusNormal
            nResult = ECS_CanCheck
        Case EREBusStatus.ST_BusMergeStopped, EREBusStatus.ST_BusSlitpStop, EREBusStatus.ST_BusStopped
            nResult = ECS_CanotCheck
        Case EREBusStatus.ST_BusStopCheck
            nResult = ECS_CanExtraCheck
    End Select
    
    If ptCheckBus.BusMode = EBusType.TP_ScrollBus Then
        If ptCheckBus.Status <> ST_BusChecking And ptCheckBus.Status <> ST_BusCheckExChecking And ptCheckBus.Status <> ST_BusNormal Then
            nResult = ECS_CanotCheck
        End If
    Else
        dptCheckBus = Time
        lHaveTime = DateDiff("s", dptCheckBus, DateAdd("n", -g_nLatestExtraCheckTime, Format(ptCheckBus.StartUpTime, "HH:mm")))
        If lHaveTime < 0 Then   '已过了最晚检票时间
            nResult = ECS_Checked
        Else
            If nResult = ECS_CanCheck Then
                lHaveTime = DateDiff("s", dptCheckBus, Format(DateAdd("n", -g_nBeginCheckTime, ptCheckBus.StartUpTime), "HH:mm"))
                If lHaveTime > 0 Then       '还未到该检票时间
                    nResult = ECS_CanotCheck
                End If
            End If
        End If
    End If
    GetBusCheckStatus = nResult
End Function

Private Sub ptDown_Resize()
    lvCheckedBus.Move 0, 0, ptDown.ScaleWidth, ptDown.ScaleHeight
End Sub


Private Sub ptLeft_Resize()
'    Spliter1.LayoutIt

End Sub

Private Sub ptUp_Resize()
   lvWillCheckBus.Move 0, 0, ptUp.ScaleWidth, ptUp.ScaleHeight
End Sub


Private Sub Timer1_Timer()
On Error GoTo ErrHandle
    Timer1.Enabled = False
    
    '初始化窗体

    MousePointer = vbHourglass

    ShowSBInfo "正在读取车次列表..."

    If g_cWillCheckBusList.Count = 0 And g_cCheckedBusList.Count = 0 Then
        BuildBusCollection
    End If

    FillBusLst

    MousePointer = vbDefault
    ShowSBInfo ""
    Exit Sub
ErrHandle:
    MousePointer = vbDefault
    ShowSBInfo ""
    ShowErrorMsg
End Sub
'更改待检车次的某一行内容
Public Sub UpdateWillCheckBusItem(pnUpdateType As Integer, ptBusInfo As tCheckBusLstInfo)
    'pnUpdateType:更改类型  1-新增,2-更新,3-删除
    'ptBusInfo:车次检票信息，如果是删除时，则只要有BusID和BusSerial字段就可以了
    Dim oListItem As ListItem
    Dim nTmpBusStatus As Integer
    Dim i As Integer
    Select Case pnUpdateType
        Case 1, 2
            nTmpBusStatus = GetBusCheckStatus(ptBusInfo)
            If nTmpBusStatus = ECS_Checked Or nTmpBusStatus = ECS_CanExtraCheck Then Exit Sub
            If pnUpdateType = 1 Then
                Set oListItem = lvWillCheckBus.ListItems.Add(, , ptBusInfo.BusID)
            Else
                Set oListItem = lvWillCheckBus.FindItem(ptBusInfo.BusID)
            End If
            '如是滚动车次则显示"流水车次",否则显示发车时间，前面加空格是为了正确排序
            oListItem.SubItems(1) = IIf(ptBusInfo.BusMode = TP_ScrollBus, " " & g_cszTitleScollBus, Format(ptBusInfo.StartUpTime, "HH:mm"))
            oListItem.SubItems(2) = ptBusInfo.EndStationName
            If ptBusInfo.BusMode <> TP_ScrollBus Then
                oListItem.SubItems(3) = ptBusInfo.Company
                oListItem.SubItems(4) = ptBusInfo.Vehicle
                oListItem.SubItems(5) = ptBusInfo.Owner
            Else
                oListItem.SubItems(3) = ""
                oListItem.SubItems(4) = ""
                oListItem.SubItems(5) = ""
            End If
            oListItem.SubItems(6) = GetStatusString(ptBusInfo.Status)
            If oListItem.SubItems(6) = "车次停班" Then
                SetListViewLineColor lvWillCheckBus, oListItem.Index, vbRed
            End If
''            '判断单双
''            If ptBusInfo.CheckGateType Mod 2 = 1 Then
''               oListItem.SubItems(cnTypeIndex) = "单号"
''            Else
''               oListItem.SubItems(cnTypeIndex) = "双号"
''            End If
            oListItem.Tag = nTmpBusStatus
            oListItem.SmallIcon = anIconIndex(nTmpBusStatus)
        Case 3
            Set oListItem = lvWillCheckBus.FindItem(ptBusInfo.BusID)
            lvWillCheckBus.ListItems.Remove oListItem.Index
    End Select


End Sub
'更改已检车次的某一行内容
Public Sub UpdateCheckedBusItem(pnUpdateType As Integer, ptBusInfo As tCheckBusLstInfo)
    'pnUpdateType:更改类型  1-新增,2-更新,3-删除
    'ptBusInfo:车次检票信息，如果是删除时，则只要有BusID和BusSerial字段就可以了
    Dim oListItem As ListItem
    Dim nTmpBusStatus As Integer
    Dim i As Integer
    
    Select Case pnUpdateType
        Case 1, 2
            nTmpBusStatus = GetBusCheckStatus(ptBusInfo)
            If pnUpdateType = 1 Then
                Set oListItem = lvCheckedBus.ListItems.Add(, , ptBusInfo.BusID & IIf(ptBusInfo.BusMode = TP_ScrollBus, _
                                                                        "-" & ptBusInfo.BusSerial, ""))
            Else
                Set oListItem = lvCheckedBus.FindItem(ptBusInfo.BusID & IIf(ptBusInfo.BusMode = TP_ScrollBus, _
                                                                        "-" & ptBusInfo.BusSerial, ""))
            End If
            If oListItem Is Nothing Then Exit Sub
            '如是滚动车次则显示"流水车次",否则显示发车时间，前面加空格是为了正确排序
            oListItem.SubItems(1) = IIf(ptBusInfo.BusMode = TP_ScrollBus, g_cszTitleScollBus, Format(ptBusInfo.StartUpTime, "HH:mm"))
            oListItem.SubItems(2) = Format(ptBusInfo.StartChkTime, cszTimeStr)
            oListItem.SubItems(3) = Format(ptBusInfo.StopChkTime, cszTimeStr)
            oListItem.SubItems(4) = ptBusInfo.CheckSheet
            oListItem.SubItems(5) = ptBusInfo.EndStationName
            oListItem.SubItems(6) = ptBusInfo.Company
            oListItem.SubItems(7) = ptBusInfo.Vehicle
            oListItem.SubItems(8) = ptBusInfo.Owner
            oListItem.SubItems(9) = GetStatusString(ptBusInfo.Status)
''            '判断单双
''            If ptBusInfo.CheckGateType Mod 2 = 1 Then
''               oListItem.SubItems(8) = "单号"
''            Else
''               oListItem.SubItems(8) = "双号"
''            End If
            oListItem.Tag = nTmpBusStatus
            oListItem.SmallIcon = anIconIndex(nTmpBusStatus)
        Case 3
            Set oListItem = lvCheckedBus.FindItem(ptBusInfo.BusID)
            lvCheckedBus.ListItems.Remove oListItem.Index
    End Select
End Sub
