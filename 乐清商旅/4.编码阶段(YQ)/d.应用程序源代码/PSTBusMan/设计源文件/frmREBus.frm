VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmEnvBus 
   BackColor       =   &H00E0E0E0&
   Caption         =   "环境管理"
   ClientHeight    =   7650
   ClientLeft      =   1260
   ClientTop       =   2400
   ClientWidth     =   10815
   HelpContextID   =   2005001
   Icon            =   "frmREBus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   10815
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ptShowInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   30
      ScaleHeight     =   990
      ScaleWidth      =   10815
      TabIndex        =   11
      Top             =   30
      Width           =   10815
      Begin VB.ComboBox cboSellStation 
         Height          =   315
         Left            =   8220
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   180
         Width           =   1545
      End
      Begin VB.TextBox txtBusID 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2910
         MaxLength       =   5
         TabIndex        =   1
         Top             =   540
         Width           =   930
      End
      Begin VB.ComboBox cboStationID 
         Height          =   315
         Left            =   8220
         TabIndex        =   5
         Top             =   533
         Width           =   1545
      End
      Begin FText.asFlatTextBox txtRoute 
         Height          =   300
         Left            =   5070
         TabIndex        =   3
         Top             =   540
         Width           =   2280
         _ExtentX        =   4022
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
         Registered      =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   2910
         TabIndex        =   7
         Top             =   180
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62324736
         CurrentDate     =   36396
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Left            =   5070
         TabIndex        =   8
         Top             =   180
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62324736
         CurrentDate     =   36396
      End
      Begin RTComctl3.CoolButton cmdFind 
         Default         =   -1  'True
         Height          =   345
         Left            =   9825
         TabIndex        =   9
         Top             =   495
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
         MICON           =   "frmREBus.frx":014A
         PICN            =   "frmREBus.frx":0166
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
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上车站(&D):"
         Height          =   195
         Left            =   7275
         TabIndex        =   15
         Top             =   233
         Width           =   795
      End
      Begin VB.Label lblInputBusId 
         BackStyle       =   0  'Transparent
         Caption         =   "车次代码(&C):"
         Height          =   180
         Left            =   1800
         TabIndex        =   0
         Top             =   615
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次日期(&D):"
         Height          =   180
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "->"
         Height          =   225
         Left            =   4665
         TabIndex        =   13
         Top             =   225
         Width           =   195
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "线路代码(&R):"
         Height          =   180
         Left            =   3960
         TabIndex        =   2
         Top             =   615
         Width           =   1080
      End
      Begin VB.Label lblBusStationID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "站点(&S):"
         Height          =   195
         Left            =   7470
         TabIndex        =   4
         Top             =   608
         Width           =   600
      End
   End
   Begin MSComctlLib.ImageList imlBusIcon 
      Left            =   7680
      Top             =   1080
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
            Picture         =   "frmREBus.frx":0500
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmREBus.frx":065A
            Key             =   "FlowStop"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmREBus.frx":09F4
            Key             =   "FlowRun"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmREBus.frx":0D8E
            Key             =   "Merge"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmREBus.frx":1128
            Key             =   "Replace"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmREBus.frx":14C2
            Key             =   "Checking"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmREBus.frx":185C
            Key             =   "Checked"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmREBus.frx":19B6
            Key             =   "ExCheck"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmREBus.frx":1B10
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmREBus.frx":1C6A
            Key             =   "SlitpBus"
         EndProperty
      EndProperty
   End
   Begin ActiveBar2LibraryCtl.ActiveBar2 abAction 
      Height          =   6465
      Left            =   9390
      TabIndex        =   12
      Top             =   1020
      Width           =   1440
      _LayoutVersion  =   1
      _ExtentX        =   2540
      _ExtentY        =   11404
      _DataPath       =   ""
      Bands           =   "frmREBus.frx":1DC6
   End
   Begin MSComctlLib.ListView lvBus 
      Height          =   4890
      Left            =   0
      TabIndex        =   10
      Top             =   1110
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   8625
      SortKey         =   3
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlBusIcon"
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "车次"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "车次种类"
         Object.Width           =   141
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "日期"
         Object.Width           =   1889
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "发车时间"
         Text            =   "发车时间"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "运行线路"
         Object.Width           =   3281
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "检票口"
         Object.Width           =   865
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "车牌"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "车型"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "全额退票"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "状态"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "终到站"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "总座位"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "余座"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "隶属公司"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "已售数"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "预定数"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu pmnu_Action 
      Caption         =   "计划车次管理"
      Visible         =   0   'False
      Begin VB.Menu pmnu_BusEnvMan_Info 
         Caption         =   "车次属性(&I)"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_BusEnvMan_Allot 
         Caption         =   "车次配载信息(&L)"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_BusEnvMan_SellStation 
         Caption         =   "车次售票点信息(&W)"
      End
      Begin VB.Menu pmnu_BusEnvMan_Price 
         Caption         =   "车次票价信息(&P)"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_BusEnvMan_Check 
         Caption         =   "车次检票信息(&H)"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_BusEnvMan_Seat 
         Caption         =   "车次座位(&E)"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_Break1 
         Caption         =   "-"
      End
      Begin VB.Menu pmnu_BusEnvMan_Stop 
         Caption         =   "停班(&S)"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_BusEnvMan_Resume 
         Caption         =   "复班(&R)"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_BusEnvMan_Replace 
         Caption         =   "顶班(&T)"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_BusEnvMan_Merge 
         Caption         =   "并班(&M)"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_Break2 
         Caption         =   "-"
      End
      Begin VB.Menu pmnu_BusEnvMan_Add 
         Caption         =   "新增车次(&A)"
      End
      Begin VB.Menu pmnu_BusEnvMan_Copy 
         Caption         =   "复制车次(&C)"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_BusEnvMan_Del 
         Caption         =   "删除车次(&D)"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmEnvBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'车次代码      车次种类      日期          发车时间      运行线路      检票口        车牌          车型 _
全额退票      状态          终到站        总座位        隶属公司
Const cnBusID = 0
Const cnBusType = 1
Const cnDate = 2
Const cnOffTime = 3
Const cnRoute = 4
Const cnCheckGate = 5
Const cnLicenseTag = 6
Const cnVehicleType = 7
Const cnAllRefundment = 8
Const cnStatus = 9
Const cnEndStation = 10
Const cnTotalSeats = 11
Const cnSaleSeatQuantity = 12 '余座
Const cnCompany = 13
Const cnSelledNums = 14    '已售数
Const cnBookedNums = 15    '预定数

Const cszAllSellStation = "(所有上车站)"

Public m_BusID As String
Private m_oREBus As New REBus
Private moScheme As New REScheme
Dim WithEvents moMsg As MsgNotify
Attribute moMsg.VB_VarHelpID = -1

Private Sub abAction_BandClose(ByVal Band As ActiveBar2LibraryCtl.Band)
    If Band.name = "bndActionTabs" Then
        abAction.Visible = False
        Call Form_Resize
    End If
End Sub
Private Sub abAction_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "mnu_BusEnvMan_Info", "act_BusEnvMan_Info"
            EditBus
        Case "mnu_BusEnvMan_Price", "act_BusEnvMan_Price"
            BusTicketPrice
        Case "mnu_BusEnvMan_Check", "act_BusEnvMan_Check"
            BusCheckInfo
        Case "mnu_BusEnvMan_Check", "act_BusEnvMan_Check"
            BusCheckInfo
        Case "act_BusEnvMan_Allot"
            BusAllot
        Case "act_BusEnvMan_SellStation"
            BusSellStation
        Case "mnu_BusEnvMan_Stop"
            StopBus False
        Case "act_BusEnvMan_Stop"
            StopBus True
        Case "mnu_BusEnvMan_Replace", "act_BusEnvMan_Replace"
            ReplaceBus
        Case "mnu_BusEnvMan_Merge", "act_BusEnvMan_Merge"
            MergeBus
        Case "mnu_BusEnvMan_Resume", "act_BusEnvMan_Resume"
            ResumeBus
        Case "mnu_BusEnvMan_Add", "act_BusEnvMan_Add"
            AddBus
        Case "mnu_BusEnvMan_Copy", "act_BusEnvMan_Copy"
            CopyBus
        Case "mnu_BusEnvMan_Del", "act_BusEnvMan_Del"
            DeleteBus
        Case "act_BusEnvMan_Seat"
            EnvSeat
            
    End Select
End Sub

Public Sub BusAllot()
    '车次配载
    If lvBus.SelectedItem Is Nothing Then
        MsgBox "请选择车次!", vbExclamation, "提示"
        Exit Sub
    End If
    frmEnvBusAllot.m_bIsAllot = True
    frmEnvBusAllot.m_szBusID = lvBus.SelectedItem.Text
    frmEnvBusAllot.m_dtEnvDate = CDate(lvBus.SelectedItem.SubItems(cnDate))
    frmEnvBusAllot.Show vbModal
End Sub
Public Sub BusSellStation()
    '车次配载
    If lvBus.SelectedItem Is Nothing Then
        MsgBox "请选择车次!", vbExclamation, "提示"
        Exit Sub
    End If
    frmEnvBusAllot.m_bIsAllot = False
    frmEnvBusAllot.m_szBusID = lvBus.SelectedItem.Text
    frmEnvBusAllot.m_dtEnvDate = CDate(lvBus.SelectedItem.SubItems(cnDate))
    frmEnvBusAllot.Show vbModal
End Sub
Public Sub BusCheckInfo()
    On Error GoTo ErrHandle
    If lvBus.SelectedItem Is Nothing Then
        MsgBox "请选择车次!", vbExclamation, "提示"
        Exit Sub
    End If
    Dim szbusID As String, dtEnvDate As Date
    szbusID = lvBus.SelectedItem.Text
    dtEnvDate = CDate(lvBus.SelectedItem.SubItems(cnDate))

    Dim oCheckSheet As New STShell.CommDialog
    oCheckSheet.Init g_oActiveUser
    If lvBus.SelectedItem.SmallIcon = "FlowRun" Or lvBus.SelectedItem.SmallIcon = "FlowStop" Then
        
        oCheckSheet.ShowEnvScrollBusList dtEnvDate, szbusID
    Else
    
        oCheckSheet.ShowCheckInfo dtEnvDate, szbusID
    End If

    Set oCheckSheet = Nothing
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Public Sub EditBus()
    Dim szbusID As String
    If lvBus.SelectedItem Is Nothing Then Exit Sub
    szbusID = lvBus.SelectedItem.Text
    frmArrangeEnvBus.Status = EFS_Modify
    frmArrangeEnvBus.m_szBusID = szbusID
    frmArrangeEnvBus.m_dtBusDate = CDate(lvBus.SelectedItem.SubItems(2))
    frmArrangeEnvBus.Show vbModal
End Sub
Private Sub cboStationID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        AddCboStation cboStationID
    End If
            
End Sub
Private Sub cmdFind_Click()
   QueryBus
End Sub

Public Sub BusTicketPrice()
On Error GoTo ErrHandle
    If lvBus.SelectedItem Is Nothing Then
        MsgBox "请选择车次!", vbExclamation, "提示"
        Exit Sub
    End If
    frmEnvBusPrice.m_szBusID = lvBus.SelectedItem.Text
    frmEnvBusPrice.m_dtEnvDate = CDate(lvBus.SelectedItem.SubItems(cnDate))
    frmEnvBusPrice.Show vbModal
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Public Sub StopBus(pbAtOnce As Boolean)
    'pbAtOnce 是否立即停班
    If pbAtOnce Then
        SelectedStopBus , , False
    Else
        frmEnvBusStop.m_szBusID = lvBus.SelectedItem.Text
        frmEnvBusStop.m_dtBusDate = CDate(lvBus.SelectedItem.SubItems(cnDate))
        frmEnvBusStop.Show vbModal
    End If
End Sub
Public Sub ReplaceBus()
    
    frmEnvBusReplace.m_szBusID = lvBus.SelectedItem.Text
    frmEnvBusReplace.m_dtBusDate = CDate(lvBus.SelectedItem.SubItems(cnDate))
    frmEnvBusReplace.m_bIsParent = True
    frmEnvBusReplace.Show vbModal
    
End Sub
Public Sub MergeBus()
    '并班
    frmEnvBusMerge.m_szBusID = lvBus.SelectedItem.Text
    frmEnvBusMerge.m_dtBusDate = lvBus.SelectedItem.SubItems(cnDate)
    frmEnvBusMerge.Show vbModal
End Sub


Public Sub AddBus()
    frmArrangeEnvBus.Status = EFS_AddNew
    frmArrangeEnvBus.Show vbModal
'    frmWizardAddBus.m_bIsParent = True
'    frmWizardAddBus.m_nWizardType = 2
'    frmWizardAddBus.Show vbModal
End Sub




Private Sub Form_Activate()
    MDIScheme.ActiveToolBar "envbus", True
'    ActiveSystemToolBar True
    WriteTitleBar "环境车次"
    Call Form_Resize
    SetMenuEnabled
    
End Sub

Private Sub Form_Deactivate()
    MDIScheme.ActiveToolBar "envbus", False
'    ActiveSystemToolBar False
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
    MDIScheme.ActiveToolBar "envbus", False
'    ActiveSystemToolBar False
    
    '保存列头
    SaveHeadWidth Me.name, lvBus
End Sub
Private Sub Form_Load()
    '初始化业务对象
    moScheme.Init g_oActiveUser
    m_oREBus.Init g_oActiveUser
    Set moMsg = New MsgNotify
    moMsg.Unit = g_szLocalUnit
    dtpStartDate.Value = Date
    dtpEndDate.Value = Date
    '初始化样式
    FillSellStation
    
    AlignHeadWidth Me.name, lvBus
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
        PopupMenu pmnu_Action
    End If
End Sub

Private Sub pmnu_BusEnvMan_Add_Click()
    AddBus
End Sub

Private Sub pmnu_BusEnvMan_Allot_Click()
    BusAllot
End Sub

Private Sub pmnu_BusEnvMan_Check_Click()
    BusCheckInfo
End Sub

Private Sub pmnu_BusEnvMan_Copy_Click()
    CopyBus
End Sub

Private Sub pmnu_BusEnvMan_Del_Click()
    DeleteBus
End Sub

Private Sub pmnu_BusEnvMan_Info_Click()
    EditBus
End Sub

Private Sub pmnu_BusEnvMan_Merge_Click()
    '并班
    MergeBus
End Sub

Private Sub pmnu_BusEnvMan_Price_Click()
    BusTicketPrice
End Sub

Private Sub pmnu_BusEnvMan_Replace_Click()
    ReplaceBus
End Sub

Private Sub pmnu_BusEnvMan_Resume_Click()
    ResumeBus
End Sub

Private Sub pmnu_BusEnvMan_Seat_Click()
    EnvSeat
End Sub

Private Sub pmnu_BusEnvMan_SellStation_Click()
    BusSellStation
    
End Sub

Private Sub pmnu_BusEnvMan_Stop_Click()
    StopBus False
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

Private Sub moMsg_ExStartCheckBus(ByVal szbusID As String, ByVal dtBusDate As Date, ByVal nBusSerialNo As Integer)
    UpdateList szbusID, dtBusDate
End Sub

Private Sub moMsg_StartCheckBus(ByVal szbusID As String, ByVal dtBusDate As Date, ByVal nBusSerialNo As Integer)
    UpdateList szbusID, dtBusDate
End Sub

Private Sub moMsg_StopCheckBus(ByVal szbusID As String, ByVal dtBusDate As Date, ByVal nBusSerialNo As Integer)
    UpdateList szbusID, dtBusDate
End Sub

Private Sub txtBusID_GotFocus()
txtBusID.SelStart = 0
txtBusID.SelLength = 100
End Sub

Private Sub txtBusID_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyReturn
       cmdFind_Click
End Select
End Sub

 '车次路单
Public Sub BusCheckSheet()

End Sub

Public Sub SelectedStopBus(Optional dtStartStopDate As Date, Optional dtEndStopDate As Date, Optional bflg As Boolean)
    Dim i As Integer, nBusCount As Integer, nSelBus As Integer, nStopBus As Integer
    Dim dtStopDate As Date
    Dim szErrString As String, szbusID As String
    Dim lErrNumber As Long
    Dim szMsg As String
On Error GoTo ErrHandle
   
    nBusCount = lvBus.ListItems.Count
    
    For i = 1 To nBusCount
        If lvBus.ListItems(i).Selected = True Then nSelBus = nSelBus + 1
    Next
    
    If nSelBus = 1 Then szErrString = "停班选择的车次[" & Trim(lvBus.SelectedItem.Text) & "]？"
    If nSelBus > 1 Then szErrString = "停班选择的" & nSelBus & "班车次？"
    
    If MsgBox(szErrString, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
    szErrString = ""
    WriteProcessBar True, , nSelBus
    For i = 1 To nBusCount
        If lvBus.ListItems(i).Selected = True Then
            If bflg = True Then
                dtStopDate = dtStartStopDate
            Else
                dtStopDate = CDate(lvBus.ListItems(i).SubItems(cnDate))
            End If
            szbusID = Trim(lvBus.ListItems(i).Text)
            nStopBus = nStopBus + 1
            WriteProcessBar , nStopBus, nSelBus, "停班车次" & EncodeString(szbusID)
            
            m_oREBus.Identify szbusID, dtStopDate
            
            If m_oREBus.SaledSeatCount > 0 Then
                If MsgBox("车次[" & szbusID & "]已有[" & m_oREBus.SaledSeatCount & "张]售票，是否停班？", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                  'szMsg = szMsg & "车次[" & szbusID & "]停班失败!" & vbCrLf
                  GoTo NextBus
                End If
            End If
            
            If m_oREBus.HaveLugge = True Then
                If MsgBox("车次[" & szbusID & "]有行包，是否停班？", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                  szMsg = szMsg & "车次[" & szbusID & "]停班失败!" & vbCrLf
                  GoTo NextBus
                End If
            End If
            
            If IdentifyBusStatus(m_oREBus.busStatus) = False Then szMsg = szMsg & "车次[" & szbusID & "]停班失败!" & vbCrLf: GoTo NextBus
            
            If bflg = True Then
                m_oREBus.StopBus dtStartStopDate, dtEndStopDate, g_bStopAllRefundment
            Else
                m_oREBus.StopBus dtStopDate, dtStopDate, g_bStopAllRefundment
            End If
              szMsg = szMsg & "车次[" & szbusID & "]停班成功" & Chr(10)
            lvBus.ListItems(i).SmallIcon = "Stop"
            lvBus.ListItems(i).Tag = "STOP"
            lvBus.ListItems(i).SubItems(cnStatus) = "停班"
            lvBus.ListItems(i).ListSubItems(cnStatus).ForeColor = vbRed
            lvBus.ListItems(i).ListSubItems(cnStatus).ReportIcon = vbEmpty
                
            If g_bStopAllRefundment Then
                    lvBus.ListItems(i).SubItems(cnAllRefundment) = "是"
                    lvBus.ListItems(i).ListSubItems(cnAllRefundment).ForeColor = vbRed
            Else
                    lvBus.ListItems(i).SubItems(cnAllRefundment) = "否"
            End If
            
            If m_oREBus.BusType = TP_ScrollBus Then
                lvBus.ListItems(i).SmallIcon = "FlowStop"
            End If
        End If
NextBus:
    Next
    WriteProcessBar False, , , ""
    If szErrString <> "" Then MsgBox szErrString, vbExclamation, Me.Caption
    If szMsg <> "" Then MsgBox szMsg, vbInformation, Me.Caption
    Exit Sub
ErrHandle:
    szErrString = szErrString & vbCrLf & "车次[" & szbusID & "]" & err.Description
    lErrNumber = err.Number
    ShowSBInfo szErrString
    Resume NextBus
End Sub
'复班
Public Sub ResumeBus()
    Dim i As Integer, nBusCount As Integer, nSelBus As Integer, nResBus As Integer
    Dim dtStopDate As Date
    Dim szErrString As String, szbusID As String
'    Dim lErrNumber As Long
On Error GoTo ErrHandle
    nBusCount = lvBus.ListItems.Count
    If nBusCount = 0 Then Exit Sub
    For i = 1 To nBusCount
        If lvBus.ListItems(i).Selected = True Then nSelBus = nSelBus + 1
    Next
    If nSelBus = 1 Then szErrString = "复班选择的车次[" & Trim(lvBus.SelectedItem.Text) & "]"
    If nSelBus > 1 Then szErrString = "复班选择的" & nSelBus & "班车次..."
    
    If MsgBox(szErrString, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
    szErrString = ""
    WriteProcessBar True
    nResBus = 0
    For i = 1 To nBusCount
        If lvBus.ListItems(i).Selected = True Then
            dtStopDate = CDate(lvBus.ListItems(i).SubItems(cnDate))
            szbusID = Trim(lvBus.ListItems(i).Text)
            m_oREBus.Identify szbusID, dtStopDate
            m_oREBus.ResumeBus dtStopDate, dtStopDate, False
            nResBus = nResBus + 1
            WriteProcessBar , nResBus, nSelBus, "停班车次" & EncodeString(szbusID)
            lvBus.ListItems(i).SmallIcon = "Run"
            lvBus.ListItems(i).SubItems(cnStatus) = "运行"
            lvBus.ListItems(i).ListSubItems(cnStatus).ForeColor = vbBlack
            lvBus.ListItems(i).SubItems(cnAllRefundment) = "否"
            lvBus.ListItems(i).Tag = ""
            lvBus.ListItems(i).ListSubItems(cnAllRefundment).ForeColor = vbBlack
            lvBus.ListItems(i).ListSubItems(cnStatus).ReportIcon = vbEmpty
            If m_oREBus.BusType = TP_ScrollBus Then
                lvBus.ListItems(i).SmallIcon = "FlowRun"
            End If
        End If
NextBus:
    Next
    WriteProcessBar False
    If szErrString <> "" Then MsgBox szErrString, vbExclamation, Me.Caption
    Exit Sub
ErrHandle:
    WriteProcessBar False
    szErrString = szErrString & vbCrLf & "车次" & EncodeString(szbusID) & err.Description
    Resume NextBus
End Sub

Public Sub EnvSeat()
    '环境座位
    With lvBus.SelectedItem
    
        frmEnvReserveSeat.m_szBusID = .Text
        frmEnvReserveSeat.m_dtEnvDate = .SubItems(cnDate)
        frmEnvReserveSeat.Show vbModal
        
    End With
End Sub

Private Sub txtRoute_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyReturn
       cmdFind_Click
End Select
End Sub

Public Sub CopyBus()
On Error GoTo ErrHandle
    Dim szOldBusID As String
    Dim dyOldDate As Date
    Dim szNewBusID As String
    Dim oShell As New CommShell
    Dim oListItem As ListItem
    Dim dyStartDate As Date
    Dim dyEndDate As Date
    Dim i As Integer
    Dim dyTemp As Date

    dyOldDate = lvBus.SelectedItem.SubItems(cnDate)
    frmCopyEvnBus.Show vbModal
    If frmCopyEvnBus.m_bOk Then
        
        ShowSBInfo "正在复制车次,请等待..."
        
        dyStartDate = frmCopyEvnBus.m_dtStartDate
        szOldBusID = frmCopyEvnBus.m_szOldBusID
        szNewBusID = frmCopyEvnBus.m_szNewBusID
        dyEndDate = frmCopyEvnBus.m_dtEndDate
        
        SetBusy
        m_oREBus.Identify szOldBusID, dyOldDate
        m_oREBus.CloneBus dyStartDate, szNewBusID, , dyEndDate
        For i = 0 To DateDiff("d", dyStartDate, dyEndDate)
            dyTemp = DateAdd("d", i, dyStartDate)
            AddList szNewBusID, dyTemp
        Next i
        
        MsgBox "新增车次" & szNewBusID & "发车日期" & dyStartDate & "到" & dyEndDate & ":" & Chr(10) _
            & "它由车次" & szOldBusID & ",发车日期为" & dyOldDate & ",终点站为" & lvBus.SelectedItem.SubItems(cnEndStation) & "拷贝产生"
        
    End If
    ShowSBInfo ""
    SetNormal
    
    Exit Sub
ErrHandle:
    SetNormal
    ShowErrorMsg


End Sub

Public Sub PasteBus()
'    Dim liTemp As ListItem
'    Dim szBusID As String
'On Error GoTo ErrHandle
'    If m_tCopyBus.Tag = "CUT" Then
'    Else
'    szBusID = InputBox("输入粘贴车次代码", "环境--粘贴车次", "复制" & m_tCopyBus.pszBusID)
'    If Left(szBusID, 2) = "复制" Or szBusID = "" Then
'        If szBusID <> "" Then MsgBox "车次代码不正确,不能复制车次", vbExclamation, "环境"
'        Exit Sub
'    Else
'        SetBusy
'        m_oREBus.Identify m_tCopyBus.pszBusID, m_tCopyBus.BusDate
'        m_oREBus.CloneBus m_tCopyBus.BusDate, szBusID
'        SetNormal
'    End If
'    Set liTemp = lvBus.ListItems.Add(, , szBusID, , m_tCopyBus.BusSmallIcon)
'    liTemp.subitems()= m_tCopyBus.BusTypeName
'    liTemp.subitems()= Format(m_tCopyBus.BusDate, "YYYY-MM-DD")
'    liTemp.subitems()= m_tCopyBus.StartupTime
'    liTemp.subitems()= m_tCopyBus.RouteName
'    liTemp.subitems()= m_tCopyBus.CheckGate
'    liTemp.subitems()= m_tCopyBus.VehicleTag
'    liTemp.subitems()= m_tCopyBus.VehicleType
'    liTemp.subitems()= "否"
'    liTemp.Tag = "STOP"
'    liTemp.SmallIcon = "Stop"
'    liTemp.subitems()= "停班"
'    liTemp.subitems(cnStatus).ForeColor = vbRed
'    liTemp.subitems(cnStatus).ReportIcon = vbEmpty
'    liTemp.subitems()= m_tCopyBus.DestStation
'    liTemp.subitems()= m_tCopyBus.SeatTotal
'    liTemp.subitems()= m_tCopyBus.CompanyName
'    SetNormal
'    MsgBox "车次[" & szBusID & "]粘贴成功", vbInformation, "环境"
'    End If
'Exit Sub
'ErrHandle:
'    SetNormal
'    ShowErrorMsg
End Sub

Public Sub CutBus()
'    lvBus.SelectedItem.Ghosted = True
'    m_tCopyBus.pszBusID = Trim(lvBus.SelectedItem.Text)
'    m_tCopyBus.BusDate = CDate(lvBus.SelectedItem.subitems(cnBusType).Text)
'    m_tCopyBus.BusTypeName = lvBus.SelectedItem.subitems(cnDate).Text
'    m_tCopyBus.StartupTime = lvBus.SelectedItem.subitems(cnRoute).Text
'    m_tCopyBus.RouteName = lvBus.SelectedItem.subitems(cnCheckGate).Text
'    m_tCopyBus.CheckGate = lvBus.SelectedItem.subitems(cnLicenseTag).Text
'    m_tCopyBus.VehicleTag = lvBus.SelectedItem.subitems(cnVehicleType).Text
'    m_tCopyBus.VehicleType = lvBus.SelectedItem.subitems(cnAllRefundment).Text
'
'    m_tCopyBus.AllReturn = lvBus.SelectedItem.subitems(cnStatus).Text
'    m_tCopyBus.BusStatus = lvBus.SelectedItem.subitems(cnEndStation).Text
'    m_tCopyBus.DestStation = lvBus.SelectedItem.subitems(cnTotalSeats).Text
'
'    m_tCopyBus.SeatTotal = lvBus.SelectedItem.subitems(cnCompany).Text
'    m_tCopyBus.CompanyName = lvBus.SelectedItem.subitems(13).Text
'    m_tCopyBus.BusSmallIcon = lvBus.SelectedItem.SmallIcon
'    lvBus.SelectedItem.Ghosted = False
'    m_tCopyBus.Tag = ""
'    m_nIndex = lvBus.SelectedItem.Index
End Sub


Public Sub DeleteBus()
    Dim dtRunDate As Date
    Dim szbusID As String
    Dim nResult As VbMsgBoxResult
On Error GoTo ErrHandle
    szbusID = Trim(lvBus.SelectedItem.Text)
    dtRunDate = CDate(lvBus.SelectedItem.SubItems(cnDate))
    nResult = MsgBox("是否删除车次[" & szbusID & "]?", vbQuestion + vbYesNo + vbDefaultButton2, "环境")
    If nResult = vbNo Then Exit Sub
    SetBusy
    m_oREBus.Identify szbusID, dtRunDate
    m_oREBus.Delete
    lvBus.ListItems.Remove lvBus.SelectedItem.Index
    SumBusNum
    SetNormal
    SetMenuEnabled
    
Exit Sub
ErrHandle:
    SetNormal
    ShowErrorMsg
End Sub
'添加一个车次信息到Listview中
Public Sub AddList(pszBusID As String, pdyRunDate As Date)
    On Error GoTo ErrHandle
    Dim aszBus() As String
    aszBus = moScheme.GetBus(pdyRunDate, , , , , , pszBusID)
    FillBusItem aszBus, pdyRunDate, False
    SetMenuEnabled
    SumBusNum
    Exit Sub
ErrHandle:
    WriteProcessBar False
    ShowSBInfo ""
    ShowErrorMsg
End Sub

Public Sub UpdateList(pszBusID As String, pdyBusDate As Date)
    On Error GoTo ErrHandle
    '更新一个车次信息在Listview中
    Dim aszBus() As String
    aszBus = moScheme.GetBus(pdyBusDate, , , , , , pszBusID)
    If ArrayLength(aszBus) = 0 Then Exit Sub
    FillBusItem aszBus, pdyBusDate, True
    SumBusNum
    Exit Sub
ErrHandle:
    WriteProcessBar False
    ShowSBInfo ""
    ShowErrorMsg
End Sub

Private Sub QueryBus()
On Error GoTo ErrHandle
    Dim i As Integer
    Dim nCount As Integer
    Dim oListItem As ListItem
    Dim aszBus() As String
    Dim nDay As Integer
    Dim szQueryRoute As String
    Dim szDestStation As String
    Dim dtDay As Date
        
    SetBusy
    ShowSBInfo "正在查询指定的车次信息，请等待..."
    lvBus.ListItems.Clear
    
    '获得该计划的所有车次，并可按线路和车次代码模糊查询
    nDay = DateDiff("d", dtpStartDate.Value, dtpEndDate.Value)
    szDestStation = ResolveDisplay(cboStationID.Text)
    szQueryRoute = ResolveDisplay(txtRoute.Text)
    nCount = 0
    WriteProcessBar True, , nCount
    For i = 0 To nDay
        '得到指定日期的环境车次
        dtDay = DateAdd("d", i, dtpStartDate.Value)
        aszBus = moScheme.GetBus(dtDay, Trim(txtBusID.Text), szQueryRoute, szDestStation, True, IIf(ResolveDisplay(cboSellStation.Text) = "(所有上车站)", "", ResolveDisplay(cboSellStation.Text)))
        nCount = nCount + ArrayLength(aszBus)
                

        FillBusItem aszBus, dtDay '填充
    Next i
    
    SetMenuEnabled
    
    If nCount = 0 Then
      SetNormal
      MsgBox "没有您需要的数据,请检查查询条件!", vbInformation + vbOKOnly, Me.Caption
      Exit Sub
    End If
    
    SumBusNum
    
'    ShowSBInfo "共" & nCount & "个车次信息", ESB_ResultCountInfo
    
    SetNormal
    Exit Sub
ErrHandle:
    WriteProcessBar False
    ShowSBInfo ""
    SetNormal
    ShowErrorMsg
    
End Sub

'增加一个车次数组信息至Listview中
Private Sub FillBusItem(aszBus() As String, pdtDate As Date, Optional pbIsUpdate As Boolean = False)  ', nCount As Integer, Optional ByVal plIndex As Long) As Long
    'pbIsUpdate  是否是更新,默认是新增
    
    '返回列表项的索引
    Dim i As Integer
    Dim oListItem As ListItem
    Dim szStopDateAndStartDateMsg As String
    Dim eStatus As EREBusStatus
    Dim nCount As Integer
    nCount = ArrayLength(aszBus)
    If nCount = 0 Then Exit Sub
    For i = 1 To nCount
        WriteProcessBar , i, nCount, "得到车次" & aszBus(i, 1)
        If Not pbIsUpdate Then
            '如果是新增
            Set oListItem = lvBus.ListItems.Add(, , aszBus(i, 1))
        Else
            '如果是修改
            Set oListItem = lvBus.SelectedItem
        End If
        If Val(aszBus(i, 8)) <> TP_ScrollBus Then
            oListItem.SmallIcon = "Run"
            oListItem.SubItems(cnOffTime) = Format(aszBus(i, 2), "hh:mm")
        Else
            oListItem.SmallIcon = "FlowRun"
            If aszBus(i, 2) = "" Then
                oListItem.SubItems(cnOffTime) = "流水车次"
            Else
                oListItem.SubItems(cnOffTime) = Format(aszBus(i, 2), "hh:mm")
            End If
        End If
        oListItem.SubItems(cnDate) = Format(pdtDate, cszDateStr)
        oListItem.SubItems(cnBusType) = aszBus(i, 14)
        oListItem.SubItems(cnRoute) = Trim(aszBus(i, 3))
        oListItem.SubItems(cnCheckGate) = Trim(aszBus(i, 4))
        oListItem.SubItems(cnLicenseTag) = Trim(aszBus(i, 5))
        oListItem.SubItems(cnVehicleType) = Trim(aszBus(i, 6))
        If Val(aszBus(i, 9)) = 0 Then
            oListItem.SubItems(cnAllRefundment) = "否"
        Else
            oListItem.SubItems(cnAllRefundment) = "是"
            oListItem.ListSubItems(cnAllRefundment).ForeColor = vbRed
        End If
        eStatus = Val(aszBus(i, 7))
        If eStatus = ST_BusStopped Or eStatus = ST_BusMergeStopped Or eStatus = ST_BusSlitpStop Then
            oListItem.Tag = "STOP"
            '        oListItem.subitems()= "停班"
            oListItem.SubItems(cnStatus) = "停班"
            oListItem.ListSubItems(cnStatus).ForeColor = vbRed
            If Val(aszBus(i, 8)) = TP_ScrollBus Then
                oListItem.SmallIcon = "FlowStop"
            End If
        Else
            oListItem.SubItems(cnStatus) = "运行"
        End If
        Select Case eStatus
        Case ST_BusStopped
            oListItem.SmallIcon = "Stop"
        Case ST_BusChecking
            oListItem.SmallIcon = "Checking"
        Case ST_BusExtraChecking
            oListItem.SmallIcon = "ExCheck"
        Case ST_BusStopCheck
            oListItem.SmallIcon = "Checked"
        Case ST_BusReplace
            oListItem.SmallIcon = "Replace"
        Case ST_BusSlitpStop
            oListItem.SmallIcon = "Merge"
        End Select
        oListItem.SubItems(cnEndStation) = aszBus(i, 10)
        oListItem.SubItems(cnTotalSeats) = aszBus(i, 11)
        oListItem.SubItems(cnCompany) = aszBus(i, 12)
        oListItem.SubItems(cnSaleSeatQuantity) = aszBus(i, 15)
        oListItem.SubItems(cnSelledNums) = aszBus(i, 16)
        oListItem.SubItems(cnBookedNums) = aszBus(i, 17)
    Next i
    If nCount > 1 Then
        lvBus.ListItems(1).Selected = True
        lvBus.ListItems(1).EnsureVisible
    Else
        For i = 1 To lvBus.ListItems.Count
            lvBus.ListItems(i).Selected = False
        Next i
        oListItem.Selected = True
'        oListItem.EnsureVisible
    End If
    WriteProcessBar False
    ShowSBInfo ""
End Sub



Private Sub SetMenuEnabled()
    Dim bEnabled As Boolean
    
    If lvBus.SelectedItem Is Nothing Then
        bEnabled = False
    Else
        bEnabled = True
    End If
    With MDIScheme.abMenuTool
        .Bands("mnu_BusEnvMan").Tools("mnu_BusEnvMan_Info").Enabled = bEnabled
        .Bands("mnu_BusEnvMan").Tools("mnu_BusEnvMan_Price").Enabled = bEnabled
        .Bands("mnu_BusEnvMan").Tools("mnu_BusEnvMan_Check").Enabled = bEnabled
        .Bands("mnu_BusEnvMan").Tools("mnu_BusEnvMan_Stop").Enabled = bEnabled
        .Bands("mnu_BusEnvMan").Tools("mnu_BusEnvMan_Resume").Enabled = bEnabled
        .Bands("mnu_BusEnvMan").Tools("mnu_BusEnvMan_Replace").Enabled = bEnabled
        .Bands("mnu_BusEnvMan").Tools("mnu_BusEnvMan_Merge").Enabled = bEnabled
        .Bands("mnu_BusEnvMan").Tools("mnu_BusEnvMan_Copy").Enabled = bEnabled
        .Bands("mnu_BusEnvMan").Tools("mnu_BusEnvMan_Del").Enabled = bEnabled
        .Bands("mnu_BusEnvMan").Tools("mnu_BusEnvMan_Seat").Enabled = bEnabled
    End With
    With abAction.Bands("bndActionTabs").ChildBands("actBusScheme")
        .Tools("act_BusEnvMan_Stop").Enabled = bEnabled
        .Tools("act_BusEnvMan_Resume").Enabled = bEnabled
        .Tools("act_BusEnvMan_Allot").Enabled = bEnabled
        .Tools("act_BusEnvMan_SellStation").Enabled = bEnabled
        .Tools("act_BusEnvMan_Replace").Enabled = bEnabled
        .Tools("act_BusEnvMan_Merge").Enabled = bEnabled
        .Tools("act_BusEnvMan_Info").Enabled = bEnabled
        .Tools("act_BusEnvMan_Price").Enabled = bEnabled
        .Tools("act_BusEnvMan_Check").Enabled = bEnabled
        .Tools("act_BusEnvMan_Copy").Enabled = bEnabled
        .Tools("act_BusEnvMan_Del").Enabled = bEnabled
        .Tools("act_BusEnvMan_Seat").Enabled = bEnabled
    End With
    pmnu_BusEnvMan_Info.Enabled = bEnabled
    pmnu_BusEnvMan_Price.Enabled = bEnabled
    pmnu_BusEnvMan_Check.Enabled = bEnabled
    pmnu_BusEnvMan_Allot.Enabled = bEnabled
    pmnu_BusEnvMan_SellStation.Enabled = bEnabled
    pmnu_BusEnvMan_Stop.Enabled = bEnabled
    pmnu_BusEnvMan_Resume.Enabled = bEnabled
    pmnu_BusEnvMan_Replace.Enabled = bEnabled
    pmnu_BusEnvMan_Merge.Enabled = bEnabled
    pmnu_BusEnvMan_Copy.Enabled = bEnabled
    pmnu_BusEnvMan_Del.Enabled = bEnabled
    pmnu_BusEnvMan_Seat.Enabled = bEnabled
    
End Sub

Private Sub SumBusNum()
    '汇总共有的车次数及停班车次数
    Dim i As Integer
    Dim nStop As Integer
    
    For i = 1 To lvBus.ListItems.Count
        If lvBus.ListItems(i).SubItems(cnStatus) = "停班" Then
            nStop = nStop + 1
        End If
    Next i
    ShowSBInfo "共" & lvBus.ListItems.Count & "个车次", ESB_ResultCountInfo
    ShowSBInfo "其中" & nStop & "个停班", ESB_WorkingInfo
    
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
