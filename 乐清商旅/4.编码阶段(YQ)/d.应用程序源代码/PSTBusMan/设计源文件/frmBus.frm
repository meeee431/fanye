VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmBus 
   BackColor       =   &H00E0E0E0&
   Caption         =   "计划管理"
   ClientHeight    =   6885
   ClientLeft      =   1140
   ClientTop       =   2595
   ClientWidth     =   11205
   HelpContextID   =   2001801
   Icon            =   "frmBus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6885
   ScaleWidth      =   11205
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ptShowInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   90
      ScaleHeight     =   990
      ScaleWidth      =   10815
      TabIndex        =   8
      Top             =   30
      Width           =   10815
      Begin VB.ComboBox cboSellStation 
         Height          =   315
         Left            =   6210
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   150
         Width           =   1920
      End
      Begin RTComctl3.CoolButton cmdFind 
         Default         =   -1  'True
         Height          =   345
         Left            =   8400
         TabIndex        =   9
         Top             =   510
         Width           =   1185
         _ExtentX        =   2090
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
         MICON           =   "frmBus.frx":014A
         PICN            =   "frmBus.frx":0166
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox cboStationID 
         Height          =   315
         Left            =   6210
         TabIndex        =   5
         Top             =   533
         Width           =   1920
      End
      Begin VB.TextBox txtBusID 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2925
         MaxLength       =   5
         TabIndex        =   1
         Top             =   150
         Width           =   1860
      End
      Begin FText.asFlatTextBox txtRoute 
         Height          =   300
         Left            =   2925
         TabIndex        =   3
         Top             =   540
         Width           =   1860
         _ExtentX        =   3281
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
         Caption         =   "上车站(&D):"
         Height          =   195
         Left            =   5100
         TabIndex        =   11
         Top             =   210
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "线路代码(&R):"
         Height          =   195
         Left            =   1800
         TabIndex        =   2
         Top             =   615
         Width           =   975
      End
      Begin VB.Label lblInputBusId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次代码(&C):"
         Height          =   195
         Left            =   1800
         TabIndex        =   0
         Top             =   210
         Width           =   960
      End
      Begin VB.Label lblBusStationID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "途经站(&S):"
         Height          =   195
         Left            =   5100
         TabIndex        =   4
         Top             =   600
         Width           =   780
      End
   End
   Begin MSComctlLib.ListView lvBus 
      Height          =   5295
      Left            =   60
      TabIndex        =   6
      Top             =   1050
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlBusIcon"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
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
         Object.Width           =   1409
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "发车时间"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "运行线路"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "检票口"
         Object.Width           =   1235
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
   Begin ActiveBar2LibraryCtl.ActiveBar2 abAction 
      Height          =   5355
      Left            =   7590
      TabIndex        =   7
      Top             =   1080
      Width           =   1500
      _LayoutVersion  =   1
      _ExtentX        =   2646
      _ExtentY        =   9446
      _DataPath       =   ""
      Bands           =   "frmBus.frx":0500
   End
   Begin MSComctlLib.ImageList imlBusIcon 
      Left            =   6960
      Top             =   2370
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBus.frx":6824
            Key             =   "StopBus"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBus.frx":697E
            Key             =   "RunBus"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBus.frx":6AD8
            Key             =   "FlowStop"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBus.frx":6E72
            Key             =   "Flow"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBus.frx":720C
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBus.frx":7AE8
            Key             =   "RunStopBus"
         EndProperty
      EndProperty
   End
   Begin VB.Menu pmnu_BusMan 
      Caption         =   "计划车次管理"
      Visible         =   0   'False
      Begin VB.Menu pmnu_BusPlanMan_Info 
         Caption         =   "车次属性"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_BusPlanMan_Allot 
         Caption         =   "车次配载"
      End
      Begin VB.Menu pmnu_BusPlanMan_SellStation 
         Caption         =   "车次售票点"
      End
      Begin VB.Menu pmnu_BusPlanMan_Price 
         Caption         =   "车次票价信息"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_BusPlanMan_Envir 
         Caption         =   "环境预览"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_Break1 
         Caption         =   "-"
      End
      Begin VB.Menu pmnu_BusPlanMan_Stop 
         Caption         =   "车次停班"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_BusPlanMan_Resume 
         Caption         =   "车次复班"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_Break2 
         Caption         =   "-"
      End
      Begin VB.Menu pmnu_BusPlanMan_Add 
         Caption         =   "新增车次"
      End
      Begin VB.Menu pmnu_BusPlanMan_Copy 
         Caption         =   "复制车次"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_BusPlanMan_Del 
         Caption         =   "删除车次"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'ListView的列位置
Const cnBusID = 0 '车次代码
Const cnOffTime = 1 '发车时间
Const cnRouteID = 2 '运行线路
Const cnCheckGate = 3 '检票口
Const cnStatus = 4 '状态
Const cnVehicleStatus = 5 '当天车辆情况

Const cszAllSellStation = "(所有上车站)"

Private moRegular As New RegularScheme
Private moBusProject As New BusProject
Private moBus As New Bus



Private Sub abAction_BandClose(ByVal Band As ActiveBar2LibraryCtl.Band)
    If Band.name = "bndActionTabs" Then
        abAction.Visible = False
        Call Form_Resize
    End If
End Sub

Private Sub abAction_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "mnu_BusPlanMan_Info", "act_BusPlanMan_Info"
            EditBus
        Case "mnu_BusPlanMan_Envir", "act_BusPlanMan_Envir"
            EnvPreview
        Case "mnu_BusPlanMan_Price", "act_BusPlanMan_Price"
            BusTicketPrice
        Case "act_BusPlanMan_Allot"
            BusAllot
        Case "mnu_BusPlanMan_Stop", "act_BusPlanMan_Stop"
            frmBusStop.Status = 1
            StopBus
        Case "mnu_BusPlanMan_Resume", "act_BusPlanMan_Resume"
'            ResumeBus
            frmBusStop.Status = 0
            StopBus
        Case "mnu_BusPlanMan_Add", "act_BusPlanMan_Add"
            AddBus
        Case "mnu_BusPlanMan_Copy", "act_BusPlanMan_Copy"
            CopyBus
        Case "mnu_BusPlanMan_Del", "act_BusPlanMan_Del"
            DeleteBus
        Case "act_BusPlanMan_Allot"
            BusAllot
        Case "act_BusPlanMan_SellStation"
            BusSellStation
    End Select
End Sub
Public Sub BusAllot()
    '车次配载
    If lvBus.SelectedItem Is Nothing Then
        MsgBox "请选择车次!", vbExclamation, "提示"
        Exit Sub
    End If
    frmBusAllot.m_bIsAllot = True
    frmBusAllot.m_szBusID = lvBus.SelectedItem.Text
    frmBusAllot.Show vbModal
End Sub
Public Sub BusSellStation()
    '车次售票点管理
    If lvBus.SelectedItem Is Nothing Then
        MsgBox "请选择车次!", vbExclamation, "提示"
        Exit Sub
    End If
    frmBusAllot.m_bIsAllot = False
    frmBusAllot.m_szBusID = lvBus.SelectedItem.Text
    frmBusAllot.Show vbModal
End Sub

Public Sub EditBus()
    
    Dim szbusID As String
    If lvBus.SelectedItem Is Nothing Then Exit Sub
    szbusID = lvBus.SelectedItem.Text
    frmArrangeBus.m_bIsParent = True
    frmArrangeBus.m_szBusID = szbusID
    frmArrangeBus.Show vbModal
End Sub


Private Sub cboStationID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        AddCboStation cboStationID
    End If
            
End Sub

Private Sub cmdFind_Click()
    QueryBus
End Sub
Private Sub Form_Activate()
    MDIScheme.ActiveToolBar "planbus", True
'    ActiveSystemToolBar True
    
    WriteTitleBar "计划车次"
    Call Form_Resize
    
End Sub

Private Sub Form_Deactivate()
    MDIScheme.ActiveToolBar "planbus", False
'    ActiveSystemToolBar False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyReturn
       SendKeys "{TAB}"
End Select
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const cnMargin = 50
    ptShowInfo.Left = 0
    ptShowInfo.Top = 0
    ptShowInfo.Width = Me.ScaleWidth
    lvBus.Left = cnMargin
    lvBus.Top = ptShowInfo.Height + cnMargin
    lvBus.Width = Me.ScaleWidth - IIf(abAction.Visible, abAction.Width, 0) - 2 * cnMargin
    lvBus.Height = Me.ScaleHeight - ptShowInfo.Height - 2 * cnMargin
    '当操作条关闭时间处理
    If abAction.Visible Then
        abAction.Move lvBus.Width + cnMargin, lvBus.Top
        abAction.Height = lvBus.Height
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIScheme.ActiveToolBar "planbus", False
'    ActiveSystemToolBar False

    '保存列头
    SaveHeadWidth Me.name, lvBus
 End Sub

Private Sub Form_Load()
    '初始化业务对象
    moRegular.Init g_oActiveUser
    moBus.Init g_oActiveUser
    moBusProject.Init g_oActiveUser
    
    FillSellStation
    
    '初始化样式
    AlignHeadWidth Me.name, lvBus
    SortListView lvBus, 2
    QueryBus
End Sub

Private Sub lvbus_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvBus, ColumnHeader.Index
End Sub

Private Sub lvBus_DblClick()
    EditBus
End Sub



Private Sub lvBus_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case cszKeyPopMenu
           lvBus_MouseUp vbRightButton, Shift, 1, 1
        Case vbKeyDelete
            DeleteBus
        Case vbKeyReturn
            EditBus
    End Select
End Sub



Private Sub lvBus_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
'        Dim oHit As ListItem
'        Set oHit = lvBus.HitTest(X, Y)
'        If Not oHit Is Nothing Then oHit.Selected = True
'        abAction.Bands("mnu_Action").PopupMenu
        PopupMenu pmnu_BusMan
    End If
End Sub

Private Sub pmnu_BusPlanMan_Add_Click()
    AddBus
End Sub

Private Sub pmnu_BusPlanMan_Allot_Click()
    BusAllot
End Sub

Private Sub pmnu_BusPlanMan_Copy_Click()
    CopyBus
End Sub

Private Sub pmnu_BusPlanMan_Del_Click()
    DeleteBus
End Sub

Private Sub pmnu_BusPlanMan_Envir_Click()
    EnvPreview
End Sub

Private Sub pmnu_BusPlanMan_Info_Click()

    EditBus
End Sub

Private Sub pmnu_BusPlanMan_Price_Click()
    BusTicketPrice
End Sub

Private Sub pmnu_BusPlanMan_Resume_Click()
'    ResumeBus
    frmBusStop.Status = 0
    StopBus
End Sub

Private Sub pmnu_BusPlanMan_SellStation_Click()
    BusSellStation
End Sub

Private Sub pmnu_BusPlanMan_Stop_Click()
    frmBusStop.Status = 1
    StopBus
End Sub

Private Sub txtBusID_GotFocus()
    txtBusID.SelStart = 0
    txtBusID.SelLength = Len(txtBusID.Text)
End Sub

Private Sub txtRoute_ButtonClick()
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init g_oActiveUser
    aszTmp = oShell.SelectRoute
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtRoute.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))
End Sub

Private Sub txtRoute_GotFocus()
    txtRoute.SelStart = 0
    txtRoute.SelLength = Len(txtRoute.Text)
End Sub


'删除车次
Public Sub DeleteBus()
    On Error GoTo ErrHandle
    If lvBus.SelectedItem Is Nothing Then
        MsgBox "请选择车次!", vbExclamation, "提示"
        Exit Sub
    End If
    Dim szbusID As String
    szbusID = lvBus.SelectedItem.Text
    
    Dim nResult As VbMsgBoxResult
    nResult = MsgBox("是否真的删除车次[" & szbusID & "]?", vbQuestion + vbYesNo + vbDefaultButton2, "删除车次")
    If nResult = vbYes Then
        SetBusy

        moBus.Identify szbusID
        moBus.Delete
        lvBus.ListItems.Remove lvBus.SelectedItem.Index
        SetNormal
    End If
    SetMenuEnabled
    Exit Sub
ErrHandle:
    SetNormal
    ShowErrorMsg
End Sub

'预览车次
Public Sub EnvPreview()
    On Error GoTo ErrHandle
    If lvBus.SelectedItem Is Nothing Then
        MsgBox "请选择车次!", vbExclamation, "提示"
        Exit Sub
    End If
    Dim szbusID As String
    szbusID = lvBus.SelectedItem.Text
    Dim tBusVehicle() As TBusVehicleInfo
    moBus.Identify szbusID
    tBusVehicle = moBus.GetAllVehicleEx
    
    frmBusPreview.RealTimeInit szbusID, tBusVehicle, True, moBus.RunCycle, moBus.CycleStartSerialNo
    frmBusPreview.Show vbModal
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Public Sub StopBus()
    If lvBus.SelectedItem Is Nothing Then
        MsgBox "请选择车次!", vbExclamation, "提示"
        Exit Sub
    End If
    Dim szbusID As String
    szbusID = lvBus.SelectedItem.Text
    frmBusStop.Init , szbusID, g_szExePriceTable
    frmBusStop.Show vbModal
End Sub
'车次复班处理
Public Sub ResumeBus()
    On Error GoTo ErrHandle
    If lvBus.SelectedItem Is Nothing Then
        MsgBox "请选择车次!", vbExclamation, "提示"
        Exit Sub
    End If
    Dim szbusID As String
    szbusID = lvBus.SelectedItem.Text
    
    Dim nResult As VbMsgBoxResult
    
    nResult = MsgBox("是否复班车次[" & szbusID & "]?", vbQuestion + vbYesNo + vbDefaultButton2, "计划")
    
    If nResult = vbNo Then Exit Sub
    moBus.Identify szbusID
    moBus.BeginStopDate = CDate(cszEmptyDateStr)
    moBus.EndStopDate = CDate(cszEmptyDateStr)
    moBus.Update
    If moBus.BusType = TP_ScrollBus Then
    lvBus.SelectedItem.SmallIcon = "Flow"
    lvBus.SelectedItem.ListSubItems(cnStatus).ForeColor = vbBlack
    lvBus.SelectedItem.ListSubItems(cnStatus).Text = "运行"
    Else
    lvBus.SelectedItem.SmallIcon = "RunBus"
    lvBus.SelectedItem.ListSubItems(cnStatus).ForeColor = vbBlack
    lvBus.SelectedItem.ListSubItems(cnStatus).Text = "运行"
    End If
    
'    nResult = MsgBox("是否复班环境内车次[" & szbusID & "]", vbQuestion + vbYesNo + vbDefaultButton2, "计划")
'    If nResult = vbNo Then Exit Sub
'    frmEnvBusStop.i , szbusID, g_szExePlanID
'    frmEnvBusStop.Show vbModal
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
Public Sub AddBus()
    frmWizardAddBus.m_nWizardType = 1 '计划车次新增
    frmWizardAddBus.m_bIsParent = True
    frmWizardAddBus.Show vbModal
End Sub

Public Sub BusTicketPrice()
    If lvBus.SelectedItem Is Nothing Then
        MsgBox "请选择车次!", vbExclamation, "提示"
        Exit Sub
    End If
    frmBusPrice.m_szBusID = lvBus.SelectedItem.Text
    frmBusPrice.Show vbModal
End Sub
'复制车次
Public Sub CopyBus()
    Dim szOldBusID As String
    Dim szNewBusID As String
    Dim oShell As New CommShell
'    Dim aszBus() As String
    Dim i As Integer
    
    On Error GoTo ErrHandle
    
    If lvBus.SelectedItem Is Nothing Then
        MsgBox "请选择待复制车次!", vbExclamation, "提示"
        Exit Sub
    End If
    szOldBusID = lvBus.SelectedItem.Text
    szNewBusID = Trim(oShell.ShowInput("复制至", "请输入目标车次代码(&I):", False, "复制" & szOldBusID))
    If szNewBusID = "" Then Exit Sub
    
    SetBusy
    ShowSBInfo "正在复制车次,请等待..."
    
    moBus.Identify szOldBusID
    moBus.CloneBus g_szExePriceTable, szNewBusID, True, True, True
    Dim oListItem As ListItem
    Set oListItem = lvBus.ListItems.Add(, , szNewBusID, , lvBus.SelectedItem.SmallIcon)
    oListItem.ListSubItems.Add(, , lvBus.SelectedItem.SubItems(1)).ForeColor = lvBus.SelectedItem.ListSubItems(1).ForeColor
    oListItem.ListSubItems.Add(, , lvBus.SelectedItem.SubItems(2)).ForeColor = lvBus.SelectedItem.ListSubItems(2).ForeColor
    oListItem.ListSubItems.Add(, , lvBus.SelectedItem.SubItems(3)).ForeColor = lvBus.SelectedItem.ListSubItems(3).ForeColor
    oListItem.ListSubItems.Add(, , lvBus.SelectedItem.SubItems(4)).ForeColor = lvBus.SelectedItem.ListSubItems(4).ForeColor
    oListItem.ListSubItems.Add(, , lvBus.SelectedItem.SubItems(5)).ForeColor = lvBus.SelectedItem.ListSubItems(5).ForeColor
    For i = 1 To lvBus.ListItems.Count
        lvBus.ListItems(i).Selected = False
    Next i
    oListItem.Selected = True
'    oListItem.EnsureVisible
    
'    aszBus = moBusProject.GetAllBus(szNewBusID)
'    FillBusItem aszBus, True

    ShowSBInfo ""
    SetNormal
    Exit Sub
ErrHandle:
    SetNormal
    ShowErrorMsg
End Sub

'填充车次列表
Public Sub AddList(pszBusID As String)
    On Error GoTo ErrHandle
    Dim i As Long, j As Integer
    Dim aszBus() As String
    Dim nCount As Integer
    SetBusy
    moBusProject.Identify
    '获得该计划的所有车次，并可按线路和车次代码模糊查询
    aszBus = moBusProject.GetAllBus(pszBusID, , , , , pszBusID)
    FillBusItem aszBus
    SetNormal
    Exit Sub
ErrHandle:
    WriteProcessBar False
    ShowSBInfo ""
    ShowErrorMsg
End Sub

Public Sub UpdateList(pszBusID As String)
    '刷新修改的信息
    On Error GoTo ErrHandle
    Dim i As Long, j As Integer
    Dim aszBus() As String
    Dim nCount As Integer
    SetBusy
    moBusProject.Identify
    '获得该计划的所有车次，并可按线路和车次代码模糊查询
    aszBus = moBusProject.GetAllBus(pszBusID, , , , , pszBusID)
    FillBusItem aszBus, True
    SetNormal
    Exit Sub
ErrHandle:
    WriteProcessBar False
    ShowSBInfo ""
    ShowErrorMsg
End Sub


Public Sub QueryBus()
On Error GoTo ErrHandle
    Dim i As Integer, nCount As Integer
    Dim oListItem As ListItem
    Dim aszBus() As String
    Dim szQueryRoute As String
    Dim aszTmp() As String
    Dim nItemNums As Integer
    Dim j As Integer
    
    SetBusy
    ShowSBInfo "填充车次信息"
    lvBus.ListItems.Clear
    moBusProject.Identify
    '获得该计划的所有车次，并可按线路和车次代码模糊查询
    If Trim(txtRoute.Text) <> "" Then
        szQueryRoute = ResolveDisplay(txtRoute.Text)
    End If
    
    aszBus = moBusProject.GetAllBus(txtBusID.Text, szQueryRoute, ResolveDisplay(Trim(cboStationID.Text)), True, IIf(ResolveDisplay(cboSellStation.Text) = cszAllSellStation, "", ResolveDisplay(cboSellStation.Text)))
    nCount = ArrayLength(aszBus)
    FillBusItem aszBus
    If nCount > 0 Then ShowSBInfo "共" & nCount & "个车次信息", ESB_ResultCountInfo
    
    SetNormal
    Exit Sub
ErrHandle:
    SetNormal
    ShowErrorMsg
End Sub
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
        If DateDiff("d", CDate(szStartdateTemp), dtEmptyDate) <> 0 And DateDiff("d", CDate(szStartdateTemp), dtEmptyDate) <> 0 Then
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


'增加一个车次
Private Sub FillBusItem(paszBus() As String, Optional pbIsUpdate As Boolean = False)
    Dim nCount As Integer
'返回列表项的索引
    Dim i As Integer
    Dim oListItem As ListItem
    Dim szStopDateAndStartDateMsg As String
    Dim bIsBusStop As Boolean
    nCount = ArrayLength(paszBus)
    If nCount = 0 Then
        SetMenuEnabled
        Exit Sub
    End If
    WriteProcessBar True, , nCount
    For i = 1 To nCount
        WriteProcessBar , i, nCount, "得到车次" & paszBus(i, 1)
        '由 函数判断车次是否停班 TestBusStatus
        szStopDateAndStartDateMsg = TestBusStatus(paszBus(i, 6), paszBus(i, 7), bIsBusStop)
        
        If Not pbIsUpdate Then
            Set oListItem = lvBus.ListItems.Add(, , Trim(paszBus(i, 1)))
        Else
            Set oListItem = lvBus.SelectedItem
        End If
        If Val(paszBus(i, 5)) <> TP_ScrollBus Then
            If bIsBusStop = True Then
                oListItem.SmallIcon = "StopBus"
            Else
                oListItem.SmallIcon = "RunBus"
            End If
        Else
            If bIsBusStop = True Then
                oListItem.SmallIcon = "FlowStop" '流水班次停班
            Else
                oListItem.SmallIcon = "Flow" '流水车次
            End If
        End If
        
        oListItem.SubItems(1) = Format(paszBus(i, 2), "HH:mm")
        oListItem.SubItems(2) = Trim(paszBus(i, 4))
        oListItem.SubItems(3) = Trim(paszBus(i, 9))
        If bIsBusStop = True Then
            If paszBus(i, 11) <> "" Then
                oListItem.SubItems(4) = "车次停班且车辆停班" & szStopDateAndStartDateMsg
                oListItem.ListSubItems.Item(cnStatus).ForeColor = vbRed
                oListItem.SubItems(5) = paszBus(i, 10) & "(停)"
                oListItem.ListSubItems.Item(cnVehicleStatus).ForeColor = vbRed
            Else
                oListItem.SubItems(4) = "车次停班" & szStopDateAndStartDateMsg
                oListItem.ListSubItems.Item(cnStatus).ForeColor = vbRed
                oListItem.SubItems(5) = paszBus(i, 10)
            End If
        Else
            If paszBus(i, 11) <> "" Then
                oListItem.SubItems(4) = "当天车辆停班" & szStopDateAndStartDateMsg
                oListItem.ListSubItems.Item(cnStatus).ForeColor = vbRed
                oListItem.SubItems(5) = paszBus(i, 10) & "(停)"
                oListItem.ListSubItems.Item(cnVehicleStatus).ForeColor = vbRed
            Else
                oListItem.SubItems(4) = "运行" & szStopDateAndStartDateMsg
                oListItem.SubItems(5) = paszBus(i, 10) & "(开)"
                oListItem.ListSubItems.Item(cnStatus).ForeColor = vbDefault
                oListItem.ListSubItems.Item(cnVehicleStatus).ForeColor = vbDefault
            End If
        End If
    Next i
    If nCount > 1 Then
        For i = 1 To lvBus.ListItems.Count
            lvBus.ListItems(i).Selected = False
        Next i
        lvBus.ListItems(1).Selected = True
'        lvBus.ListItems(1).EnsureVisible
    Else
        For i = 1 To lvBus.ListItems.Count
            lvBus.ListItems(i).Selected = False
        Next i
        
        oListItem.Selected = True
        'oListItem.EnsureVisible
    End If
    ShowSBInfo ""
    WriteProcessBar False
    SetMenuEnabled
    
End Sub

Private Sub SetMenuEnabled()
    Dim bEnabled As Boolean
    If lvBus.ListItems.Count > 0 Then
        bEnabled = True
    Else
        bEnabled = False
    End If
    MDIScheme.abMenuTool.Bands("mnu_BusPlanMan").Tools("mnu_BusPlanMan_Info").Enabled = bEnabled
    MDIScheme.abMenuTool.Bands("mnu_BusPlanMan").Tools("mnu_BusPlanMan_Price").Enabled = bEnabled
    MDIScheme.abMenuTool.Bands("mnu_BusPlanMan").Tools("mnu_BusPlanMan_Envir").Enabled = bEnabled
    MDIScheme.abMenuTool.Bands("mnu_BusPlanMan").Tools("mnu_BusPlanMan_Stop").Enabled = bEnabled
    MDIScheme.abMenuTool.Bands("mnu_BusPlanMan").Tools("mnu_BusPlanMan_Resume").Enabled = bEnabled
    MDIScheme.abMenuTool.Bands("mnu_BusPlanMan").Tools("mnu_BusPlanMan_Copy").Enabled = bEnabled
    MDIScheme.abMenuTool.Bands("mnu_BusPlanMan").Tools("mnu_BusPlanMan_Del").Enabled = bEnabled
    
    
    abAction.Bands("bndActionTabs").ChildBands("actBusPlanMan").Tools("act_BusPlanMan_Stop").Enabled = bEnabled
    abAction.Bands("bndActionTabs").ChildBands("actBusPlanMan").Tools("act_BusPlanMan_Resume").Enabled = bEnabled
    abAction.Bands("bndActionTabs").ChildBands("actBusPlanMan").Tools("act_BusPlanMan_Copy").Enabled = bEnabled
    abAction.Bands("bndActionTabs").ChildBands("actBusPlanMan").Tools("act_BusPlanMan_Del").Enabled = bEnabled
    abAction.Bands("bndActionTabs").ChildBands("actBusPlanMan").Tools("act_BusPlanMan_Info").Enabled = bEnabled
    abAction.Bands("bndActionTabs").ChildBands("actBusPlanMan").Tools("act_BusPlanMan_Price").Enabled = bEnabled
    abAction.Bands("bndActionTabs").ChildBands("actBusPlanMan").Tools("act_BusPlanMan_Envir").Enabled = bEnabled
    abAction.Bands("bndActionTabs").ChildBands("actBusPlanMan").Tools("act_BusPlanMan_Allot").Enabled = bEnabled
    abAction.Bands("bndActionTabs").ChildBands("actBusPlanMan").Tools("act_BusPlanMan_SellStation").Enabled = bEnabled
    
    
    pmnu_BusPlanMan_Info.Enabled = bEnabled
    pmnu_BusPlanMan_Price.Enabled = bEnabled
    pmnu_BusPlanMan_Envir.Enabled = bEnabled
    pmnu_BusPlanMan_Stop.Enabled = bEnabled
    pmnu_BusPlanMan_Resume.Enabled = bEnabled
    pmnu_BusPlanMan_Copy.Enabled = bEnabled
    pmnu_BusPlanMan_Del.Enabled = bEnabled
    pmnu_BusPlanMan_Allot.Enabled = bEnabled
    pmnu_BusPlanMan_SellStation.Enabled = bEnabled
    
End Sub

'===================================================
'Modify Date：2002-11-19
'Author:fl
'Reamrk:填充所有的上车站。
'===================================================b

Private Sub FillSellStation()

    '填充上车站
    Dim nCount As Integer
    Dim i As Integer
    cboSellStation.Clear
    nCount = ArrayLength(g_atAllSellStation)
    cboSellStation.AddItem cszAllSellStation
    For i = 1 To nCount
        cboSellStation.AddItem MakeDisplayString(g_atAllSellStation(i).szSellStationID, g_atAllSellStation(i).szSellStationName)
        
        'cszAllSellStation
    Next i
    
    '填充所有的上车站
    If nCount > 0 Then cboSellStation.ListIndex = 0
End Sub





