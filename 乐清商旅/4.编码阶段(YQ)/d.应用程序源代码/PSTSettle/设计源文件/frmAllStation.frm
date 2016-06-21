VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmAllStation 
   Caption         =   "站点管理"
   ClientHeight    =   7125
   ClientLeft      =   1035
   ClientTop       =   2160
   ClientWidth     =   12435
   HelpContextID   =   2000601
   Icon            =   "frmAllStation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   12435
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ptShowInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   210
      ScaleHeight     =   990
      ScaleWidth      =   10815
      TabIndex        =   2
      Top             =   90
      Width           =   10815
      Begin VB.TextBox txtInputCode 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   4770
         TabIndex        =   1
         Top             =   428
         Width           =   615
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2850
         TabIndex        =   11
         Top             =   428
         Width           =   885
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   6420
         TabIndex        =   9
         Top             =   428
         Width           =   915
      End
      Begin RTComctl3.CoolButton cmdFind 
         Default         =   -1  'True
         Height          =   345
         Left            =   9540
         TabIndex        =   3
         Top             =   420
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
         MICON           =   "frmAllStation.frx":014A
         PICN            =   "frmAllStation.frx":0166
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin FText.asFlatTextBox txtArea 
         Height          =   300
         Left            =   8190
         TabIndex        =   10
         Top             =   428
         Width           =   1275
         _ExtentX        =   2249
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "站点代码(&I):"
         Height          =   180
         Left            =   1740
         TabIndex        =   8
         Top             =   495
         Width           =   1080
      End
      Begin VB.Label lblInputRouteId 
         BackStyle       =   0  'Transparent
         Caption         =   "输入码(&P):"
         Height          =   180
         Left            =   3870
         TabIndex        =   7
         Top             =   495
         Width           =   900
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "地区(&A):"
         Height          =   180
         Left            =   7440
         TabIndex        =   6
         Top             =   495
         Width           =   720
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "站点名(&N):"
         Height          =   180
         Index           =   0
         Left            =   5490
         TabIndex        =   5
         Top             =   495
         Width           =   900
      End
      Begin VB.Image Image1 
         Height          =   1275
         Left            =   0
         Picture         =   "frmAllStation.frx":0500
         Top             =   30
         Width           =   2010
      End
   End
   Begin MSComctlLib.ImageList imgRoute 
      Left            =   0
      Top             =   1000
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
            Picture         =   "frmAllStation.frx":19D3
            Key             =   "Station"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllStation.frx":1B2D
            Key             =   "NoSell"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvStation 
      Height          =   4635
      Left            =   270
      TabIndex        =   0
      Top             =   1200
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "站点代码"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "站点名称"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "输入码"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "售票属性"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "本地码"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "地区"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ImageList imlBusIcon 
      Left            =   0
      Top             =   4680
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
            Picture         =   "frmAllStation.frx":20C7
            Key             =   "StopBus"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllStation.frx":2221
            Key             =   "RunBus"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllStation.frx":237B
            Key             =   "FlowStop"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllStation.frx":2715
            Key             =   "Flow"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllStation.frx":2AAF
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllStation.frx":338B
            Key             =   "RunStopBus"
         EndProperty
      EndProperty
   End
   Begin ActiveBar2LibraryCtl.ActiveBar2 abAction 
      Height          =   5355
      Left            =   9090
      TabIndex        =   4
      Top             =   1140
      Width           =   1500
      _LayoutVersion  =   1
      _ExtentX        =   2646
      _ExtentY        =   9446
      _DataPath       =   ""
      Bands           =   "frmAllStation.frx":3725
   End
   Begin VB.Menu pmnu_Station 
      Caption         =   "站点"
      Visible         =   0   'False
      Begin VB.Menu pmnu_AddStation 
         Caption         =   "新增(&A)"
      End
      Begin VB.Menu pmnu_EditStation 
         Caption         =   "编辑(&E)"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_DeleteStation 
         Caption         =   "删除(&D)"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmAllStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************
'* Source File Name:frmAllStation.frm
'* Project Name:RTBusMan
'* Engineer:
'* Data Generated:2002/08/30
'* Last Revision Date:2002/08/30
'* Brief Description:所有站点
'* Relational Document:
'**********************************************************
Public m_szStationID As String
Private m_oBaseInfo As New BaseInfo '基本信息对象 BaseInfo
Private m_oStation As New Station


Public Sub AddStation()
    '新增站点
    frmStation.m_bIsParent = True
    frmStation.Status = AddStatus
    frmStation.Show vbModal
End Sub

Public Sub EditStation()
    '修改站点
    If lvStation.SelectedItem Is Nothing Then Exit Sub
    frmStation.m_bIsParent = True
    frmStation.szStationID = lvStation.SelectedItem.Text
    frmStation.Status = ModifyStatus
    frmStation.Show vbModal
    
End Sub

Public Sub DeleteStation()
    '删除站点
    On Error GoTo ErrorHandle
    Dim nResult As VbMsgBoxResult
    nResult = MsgBox("是否要删除站点[" & Trim(lvStation.SelectedItem.ListSubItems(1).Text) & "]", vbQuestion + vbYesNo + vbDefaultButton2, "站点管理")
    If nResult = vbYes Then
    m_oStation.Identify Trim(lvStation.SelectedItem.Text)
    m_oStation.Delete
    lvStation.ListItems.Remove lvStation.SelectedItem.Index
    SetMenuEnabled
    End If
Exit Sub
ErrorHandle:
'    If err.Number = 91 Then Exit Sub
    ShowErrorMsg
End Sub

Public Sub AddList(pszID As String)
    '将新增站点刷新出来
    On Error GoTo ErrorHandle
    Dim paszStation() As String
    paszStation = m_oBaseInfo.GetStation(, , pszID)
    
    FillItem paszStation
    SetMenuEnabled
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Public Sub UpdateList(pszID As String)
    '将修改站点刷新出来
    On Error GoTo ErrorHandle
    Dim paszStation() As String
    Dim nCount As Integer
    paszStation = m_oBaseInfo.GetStation(, frmStation.txtStationName.Text, pszID)
    
    FillItem paszStation, True
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub abAction_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
    Case "act_EditStation"
        EditStation
    Case "act_DeleteStation"
        DeleteStation
    Case "act_AddStation"
        AddStation
    End Select
End Sub

Private Sub cmdFind_Click()
    QueryStation
End Sub

Private Sub QueryStation()
    '填充站点
    Dim paszStation() As String
    On Error GoTo ErrorHandle
    SetBusy
    lvStation.ListItems.Clear
    paszStation = m_oBaseInfo.GetStation(Trim(ResolveDisplay(txtArea.Text)), Trim(txtName.Text), Trim(txtID.Text), Trim(txtInputCode.Text))
    FillItem paszStation
    SetMenuEnabled
    SetNormal
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub FillItem(paszStation() As String, Optional pbIsUpdate As Boolean = False)
    Dim nCount As Integer
    Dim i As Integer
    Dim liTemp As ListItem
    nCount = ArrayLength(paszStation)
    If nCount = 0 Then Exit Sub
    
    For i = 1 To nCount
        
        
        
        If Not pbIsUpdate Then
            Set liTemp = lvStation.ListItems.Add(, , paszStation(i, 1), , "Station")
        Else
            
            Set liTemp = lvStation.SelectedItem
        End If
     
        liTemp.SubItems(1) = paszStation(i, 2)
   
        liTemp.SubItems(2) = paszStation(i, 3)
        If Val(paszStation(i, 4)) <> TP_CanSellTicket Then
            liTemp.SubItems(3) = "不可售"
            liTemp.SmallIcon = "NoSell"
        Else
            liTemp.SubItems(3) = "可售"
        End If
        liTemp.SubItems(4) = paszStation(i, 5)
        liTemp.SubItems(5) = paszStation(i, 6)
    Next
    If nCount > 1 Then
        lvStation.ListItems(1).Selected = True
        lvStation.ListItems(1).EnsureVisible
    Else
        liTemp.Selected = True
        liTemp.EnsureVisible
    End If
    
End Sub


Private Sub Form_Activate()
'    'MDIScheme.ActiveToolBar "station", True
    Form_Resize
End Sub

Private Sub Form_Deactivate()
    'MDIScheme.ActiveToolBar "station", False
End Sub

Private Sub Form_Load()
    m_oBaseInfo.Init g_oActiveUser
    m_oStation.Init g_oActiveUser
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Const cnMargin = 50
    ptShowInfo.Left = 0
    ptShowInfo.Top = 0
    ptShowInfo.Width = Me.ScaleWidth
    lvStation.Left = cnMargin
    lvStation.Top = ptShowInfo.Height + cnMargin
    lvStation.Width = Me.ScaleWidth - IIf(abAction.Visible, abAction.Width, 0) - 2 * cnMargin
    lvStation.Height = Me.ScaleHeight - ptShowInfo.Height - 2 * cnMargin
    '当操作条关闭时间处理
    If abAction.Visible Then
        abAction.Move lvStation.Width + cnMargin, lvStation.Top
        abAction.Height = lvStation.Height
    End If
End Sub

Private Sub SetMenuEnabled()
    '设置菜单的可用性
    Dim bEnabled As Boolean
    If lvStation.ListItems.Count > 0 Then
        bEnabled = True
    Else
        bEnabled = False
    End If
    pmnu_EditStation.Enabled = bEnabled
    pmnu_DeleteStation.Enabled = bEnabled
    With abAction.Bands("bndActionTabs").ChildBands("actStation")
        .Tools("act_EditStation").Enabled = bEnabled
        .Tools("act_DeleteStation").Enabled = bEnabled
    End With
    'With MDIScheme.abMenuTool.Bands("mnu_StationMan")
'        .Tools("mnu_StationMan_Info").Enabled = bEnabled
'        .Tools("mnu_StationMan_Del").Enabled = bEnabled
'    End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'MDIScheme.ActiveToolBar "station", False
End Sub

'
Private Sub lvStation_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvStation, ColumnHeader.Index
End Sub

Private Sub lvStation_DblClick()

    EditStation
End Sub

Private Sub lvStation_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu pmnu_Station
    End If
End Sub

Private Sub pmnu_AddStation_Click()
    AddStation
End Sub

Private Sub pmnu_DeleteStation_Click()
    DeleteStation
End Sub

Private Sub pmnu_EditStation_Click()
    EditStation
End Sub

Private Sub txtArea_ButtonClick()
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectArea(False)
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtArea.Text = aszTemp(1, 1) & "[" & aszTemp(1, 2) & "]"
End Sub

Private Sub txtArea_GotFocus()
    txtArea.SelStart = 0
    txtArea.SelLength = 100
End Sub

Private Sub txtID_GotFocus()
    txtID.SelStart = 0
    txtID.SelLength = 100
End Sub


Private Sub txtInputCode_GotFocus()
txtInputCode.SelStart = 0
txtInputCode.SelLength = 100
End Sub

Private Sub txtName_GotFocus()
    txtName.SelStart = 0
    txtName.SelLength = 100
End Sub

