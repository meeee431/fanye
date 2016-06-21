VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmAllRoute 
   Caption         =   "线路管理"
   ClientHeight    =   6195
   ClientLeft      =   930
   ClientTop       =   2580
   ClientWidth     =   11115
   HelpContextID   =   2000401
   Icon            =   "frmAdllRoute.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   11115
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ptShowInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   -90
      ScaleHeight     =   990
      ScaleWidth      =   10815
      TabIndex        =   8
      Top             =   540
      Width           =   10815
      Begin VB.ComboBox cboSellStation 
         Height          =   300
         ItemData        =   "frmAdllRoute.frx":014A
         Left            =   3015
         List            =   "frmAdllRoute.frx":0151
         TabIndex        =   10
         Text            =   "(全部)"
         Top             =   180
         Width           =   1545
      End
      Begin VB.ComboBox cboStation 
         Height          =   300
         ItemData        =   "frmAdllRoute.frx":015D
         Left            =   5895
         List            =   "frmAdllRoute.frx":0164
         TabIndex        =   5
         Text            =   "(全部)"
         Top             =   150
         Width           =   1515
      End
      Begin VB.TextBox txtRouteName 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5895
         MaxLength       =   20
         TabIndex        =   3
         Top             =   540
         Width           =   1515
      End
      Begin RTComctl3.CoolButton cmdFind 
         Default         =   -1  'True
         Height          =   345
         Left            =   7680
         TabIndex        =   6
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
         MICON           =   "frmAdllRoute.frx":0170
         PICN            =   "frmAdllRoute.frx":018C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin FText.asFlatTextBox txtRouteId 
         Height          =   300
         Left            =   3015
         TabIndex        =   1
         Top             =   540
         Width           =   1515
         _ExtentX        =   2672
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
         Registered      =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起始站(&S):"
         Height          =   180
         Left            =   1890
         TabIndex        =   11
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "途经站(&D):"
         Height          =   180
         Left            =   4740
         TabIndex        =   4
         Top             =   210
         Width           =   900
      End
      Begin VB.Label lblInputRouteId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "线路代码(&C):"
         Height          =   180
         Left            =   1890
         TabIndex        =   0
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "线路名称(&N):"
         Height          =   180
         Left            =   4740
         TabIndex        =   2
         Top             =   600
         Width           =   1080
      End
      Begin VB.Image Image1 
         Height          =   1275
         Left            =   0
         Picture         =   "frmAdllRoute.frx":0526
         Top             =   30
         Width           =   2010
      End
   End
   Begin MSComctlLib.ImageList imgRoute 
      Left            =   765
      Top             =   3675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdllRoute.frx":19F9
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdllRoute.frx":1B55
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvRoute 
      Height          =   4635
      Left            =   30
      TabIndex        =   7
      Top             =   1770
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   8176
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgRoute"
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "线路代码"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "线路名称"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "里程数"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "起始站"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "终点站"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "状态"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "途经站点"
         Object.Width           =   21167
      EndProperty
   End
   Begin ActiveBar2LibraryCtl.ActiveBar2 abAction 
      Height          =   5355
      Left            =   9540
      TabIndex        =   9
      Top             =   2430
      Width           =   1500
      _LayoutVersion  =   1
      _ExtentX        =   2646
      _ExtentY        =   9446
      _DataPath       =   ""
      Bands           =   "frmAdllRoute.frx":1CB1
   End
   Begin VB.Menu pmnu_Route 
      Caption         =   "线路"
      Visible         =   0   'False
      Begin VB.Menu pmnu_Add 
         Caption         =   "新增(&A)"
      End
      Begin VB.Menu pmnu_Edit 
         Caption         =   "编辑(&E)"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_Delete 
         Caption         =   "删除(&D)"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_Break1 
         Caption         =   "-"
      End
      Begin VB.Menu pmnu_Section 
         Caption         =   "路段(&S)"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_Copy 
         Caption         =   "复制(&C)"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmAllRoute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************
'* Source File Name:frmAllRoute.frm
'* Project Name:RTBusMan
'* Engineer:
'* Data Generated:2002/08/27
'* Last Revision Date:2002/08/31
'* Brief Description:线路控制台
'* Relational Document:UI_BS_SM_002.DOC
'**********************************************************
Private m_oBaseInfo As New STSettle.BackBaseInfo
Private m_oRoute As New STSettle.BackRoute


Private Sub abAction_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
    Case "act_AddRoute"
        AddRoute
    Case "act_EditRoute"
        EditRoute
    Case "act_DeleteRoute"
        DeleteRoute
    Case "act_RouteSection"
        '线路路段
        RouteSection
    Case "act_CopyRoute"
        '复制线路
        CopyRoute
    End Select
End Sub

Private Sub cmdFind_Click()
    QueryRoute
End Sub

Private Sub Form_Activate()
    'MDIScheme.ActiveToolBar "route", True
    Form_Resize
    
End Sub

Private Sub Form_Deactivate()
    'MDIScheme.ActiveToolBar "route", False
End Sub

Private Sub Form_Load()
    m_oBaseInfo.Init g_oActiveUser
    FillSellStation
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    Const cnMargin = 50
    ptShowInfo.Left = 0
    ptShowInfo.Top = 0
    ptShowInfo.Width = Me.ScaleWidth
    lvRoute.Left = cnMargin
    lvRoute.Top = ptShowInfo.Height + cnMargin
    lvRoute.Width = Me.ScaleWidth - IIf(abAction.Visible, abAction.Width, 0) - 2 * cnMargin
    lvRoute.Height = Me.ScaleHeight - ptShowInfo.Height - 2 * cnMargin
    '当操作条关闭时间处理
    If abAction.Visible Then
        abAction.Move lvRoute.Width + cnMargin, lvRoute.Top
        abAction.Height = lvRoute.Height
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'MDIScheme.ActiveToolBar "route", False
End Sub

Private Sub lvRoute_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvRoute, ColumnHeader.Index
End Sub


Private Sub lvRoute_DblClick()
    EditRoute
End Sub

Private Sub lvRoute_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = cszKeyPopMenu Then
'        lvRoute_MouseUp vbRightButton, Shift, 1, 1
'    End If
End Sub

Private Sub lvRoute_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu pmnu_Route
    End If
End Sub

Private Sub pmnu_Add_Click()
    AddRoute
End Sub

Private Sub pmnu_Copy_Click()
    CopyRoute
End Sub

Private Sub pmnu_delete_Click()
    DeleteRoute
End Sub

Private Sub pmnu_edit_Click()
    EditRoute
    
End Sub

Private Sub pmnu_Section_Click()
    RouteSection
End Sub

'Private Sub txtRouteID_ButtonClick()
'    '选择线路
'    Dim oShell As New STShell.CommDialog
'    Dim aszTemp() As String
'    oShell.Init g_oActiveUser
'    aszTemp = oShell.SelectRoute(False)
'    Set oShell = Nothing
'    If ArrayLength(aszTemp) = 0 Then Exit Sub
'    txtRouteId.Text = MakeDisplayString(aszTemp(1, 1), Trim(aszTemp(1, 2)))
'End Sub

Private Sub txtRouteId_GotFocus()
    txtRouteID.SelStart = 0
    txtRouteID.SelLength = 100
End Sub
Private Sub txtRouteName_GotFocus()
    txtRouteName.SelLength = 100
    txtRouteName.SelStart = 0
End Sub

Public Sub AddRoute()
    '新增线路
    frmRoute.m_bIsParent = True
    frmRoute.Status = AddStatus
    frmRoute.Show vbModal
    
End Sub

Public Sub EditRoute()
    '修改线路
    If lvRoute.SelectedItem Is Nothing Then Exit Sub
    
    frmRoute.m_szRouteID = lvRoute.SelectedItem.Text
    frmRoute.m_bIsParent = True
    frmRoute.Status = ModifyStatus
    frmRoute.Show vbModal
End Sub


Public Sub DeleteRoute()
    '删除线路
    On Error GoTo ErrorHandle
    Dim nResult As VbMsgBoxResult
    nResult = MsgBox("是否要删除线路[" & Trim(lvRoute.SelectedItem.ListSubItems(1).Text) & "]", vbQuestion + vbYesNo + vbDefaultButton2, "线路管理")
    If nResult = vbYes Then
    m_oRoute.Init g_oActiveUser
    m_oRoute.Identify lvRoute.SelectedItem.Text
    m_oRoute.Delete
    lvRoute.ListItems.Remove lvRoute.SelectedItem.Index
    SetMenuEnabled
    End If
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Public Sub CopyRoute()
    '复制线路
'    frmCopyRoute.m_bIsParent = True
'    frmCopyRoute.m_szOldRouteID = lvRoute.SelectedItem.Text
'    frmCopyRoute.Show vbModal
End Sub


Public Sub UpdateList(pszRouteId As String)
    '刷新修改的东东
    On Error GoTo ErrorHandle
    Dim aszRoute() As String
    aszRoute = m_oBaseInfo.GetRouteEx(pszRouteId)
    FillItem aszRoute, True
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Public Sub AddList(pszRouteId As String)
    '将新增的东东刷新出来
    Dim aszRoute() As String
    On Error GoTo ErrorHandle
    aszRoute = m_oBaseInfo.GetRouteEx(pszRouteId)
    FillItem aszRoute
    SetMenuEnabled
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Public Sub RouteSection()
    '线路路段
    
'    If g_szStationID = "" Then
'        MsgBox "用户还未在系统管理中设定本站站点码,无法安排线路路段!", vbExclamation, "线路管理"
'        Exit Sub
'    End If
    frmArrangeSection.m_szRouteID = lvRoute.SelectedItem.Text
    frmArrangeSection.Show vbModal
End Sub

'Public Sub CopyRoute()
'    '复制线路
'End Sub


Private Sub SetMenuEnabled()
    '设置菜单的可用性
    Dim bEnabled  As Boolean
    If lvRoute.ListItems.Count > 0 Then
        bEnabled = True
    Else
        bEnabled = False
    End If
    pmnu_Edit.Enabled = bEnabled
    pmnu_Delete.Enabled = bEnabled
    pmnu_Section.Enabled = bEnabled
    pmnu_Copy.Enabled = bEnabled
    With abAction.Bands("bndActionTabs").ChildBands("actRoute")
        .Tools("act_EditRoute").Enabled = bEnabled
        .Tools("act_DeleteRoute").Enabled = bEnabled
        .Tools("act_RouteSection").Enabled = bEnabled
        .Tools("act_CopyRoute").Enabled = bEnabled
    End With
'    With MDIScheme.abMenuTool.Bands("mnu_RouteMan")
'        .Tools("mnu_RouteMan_Info").Enabled = bEnabled
'        .Tools("mnu_RouteMan_Section").Enabled = bEnabled
'        .Tools("mnu_RouteMan_Copy").Enabled = bEnabled
'        .Tools("mnu_RouteMan_Del").Enabled = bEnabled
'    End With
End Sub

Private Sub QueryRoute()
    '查询
    
    Dim aszRoute() As String
    Dim i As Integer
    On Error GoTo ErrorHandle
    SetBusy
    lvRoute.ListItems.Clear '先清空lvRoute
    ShowSBInfo "获得线路..."
    aszRoute = m_oBaseInfo.GetRouteEx(ResolveDisplay(txtRouteID.Text), txtRouteName.Text, IIf(cboStation.Text = "(全部)", "", ResolveDisplay(cboStation.Text)), IIf(cboSellStation.Text = "(全部)", "", ResolveDisplay(cboSellStation.Text)))
    i = ArrayLength(aszRoute)
    FillItem aszRoute
    SetMenuEnabled
    SetNormal
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg

End Sub

Private Sub FillItem(paszRoute() As String, Optional pbIsUpdate As Boolean = False)
    '填充信息
    '线路代码、线路名称、终点站、终点站名称、里程数、状态
    Dim liTemp As ListItem
    Dim i As Integer, nCount As Integer
    
    nCount = ArrayLength(paszRoute)
    If nCount = 0 Then Exit Sub
    For i = 1 To nCount
        If Not pbIsUpdate Then
            Set liTemp = lvRoute.ListItems.Add(, , paszRoute(i, 1))
        Else
            Set liTemp = lvRoute.SelectedItem
        End If
        If Val(paszRoute(i, 6)) = ST_RouteAvailable Then
            liTemp.SmallIcon = "Run"
        Else
            liTemp.SmallIcon = "Stop"
        End If
        liTemp.SubItems(1) = paszRoute(i, 2)
        liTemp.SubItems(2) = paszRoute(i, 5)
        liTemp.SubItems(3) = paszRoute(i, 7)
        liTemp.SubItems(4) = paszRoute(i, 4)
        If Val(paszRoute(i, 5)) <> 0 Then
            liTemp.SubItems(5) = "正常"
            liTemp.SubItems(6) = paszRoute(i, 3)
        Else
            liTemp.SubItems(5) = "维修"
            liTemp.SubItems(6) = paszRoute(i, 3)
            liTemp.SmallIcon = "Stop"
            SetListViewLineColor lvRoute, liTemp.Index, vbRed
        End If
    Next
    If nCount > 1 Then
        lvRoute.ListItems(1).Selected = True
        lvRoute.ListItems(1).EnsureVisible
    Else
        liTemp.Selected = True
        liTemp.EnsureVisible
    End If
End Sub


    
Private Sub FillSellStation()
    Dim nCount As Integer
    Dim i As Integer
    cboSellStation.Clear
    nCount = ArrayLength(g_atAllSellStation)
    For i = 1 To nCount
        cboSellStation.AddItem MakeDisplayString(g_atAllSellStation(i).szSellStationID, g_atAllSellStation(i).szSellStationName)
    Next i
End Sub
