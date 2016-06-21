VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmAllSection 
   Caption         =   "路段管理"
   ClientHeight    =   4620
   ClientLeft      =   990
   ClientTop       =   2430
   ClientWidth     =   11595
   HelpContextID   =   2000201
   Icon            =   "frmAllSection.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4620
   ScaleWidth      =   11595
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ptShowInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   11325
      TabIndex        =   1
      Top             =   0
      Width           =   11325
      Begin VB.ComboBox cboStation 
         Height          =   300
         ItemData        =   "frmAllSection.frx":014A
         Left            =   8955
         List            =   "frmAllSection.frx":0151
         TabIndex        =   6
         Text            =   "(全部)"
         Top             =   465
         Width           =   1050
      End
      Begin VB.TextBox txtSectionId 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5070
         MaxLength       =   4
         TabIndex        =   5
         Top             =   465
         Width           =   870
      End
      Begin VB.TextBox txtStartStationID 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   7065
         TabIndex        =   4
         Top             =   465
         Width           =   885
      End
      Begin RTComctl3.CoolButton cmdFind 
         Default         =   -1  'True
         Height          =   345
         Left            =   10080
         TabIndex        =   2
         Top             =   443
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
         MICON           =   "frmAllSection.frx":015D
         PICN            =   "frmAllSection.frx":0179
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
         Left            =   2865
         TabIndex        =   7
         Top             =   465
         Width           =   1095
         _ExtentX        =   1931
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
      Begin VB.Label lblInputRouteId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路段代码(&C):"
         Height          =   180
         Left            =   3975
         TabIndex        =   11
         Top             =   525
         Width           =   1080
      End
      Begin VB.Label lblStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "终点站(&D):"
         Height          =   180
         Left            =   7995
         TabIndex        =   10
         Top             =   525
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "线路代码(&R):"
         Height          =   180
         Left            =   1770
         TabIndex        =   9
         Top             =   525
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起站代码(&S):"
         Height          =   180
         Left            =   5985
         TabIndex        =   8
         Top             =   525
         Width           =   1080
      End
      Begin VB.Image Image1 
         Height          =   1275
         Left            =   0
         Top             =   30
         Width           =   2010
      End
   End
   Begin MSComctlLib.ImageList imgRoute 
      Left            =   -210
      Top             =   915
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
            Picture         =   "frmAllSection.frx":0513
            Key             =   "Station"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllSection.frx":066D
            Key             =   "NoSell"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlBusIcon 
      Left            =   -210
      Top             =   4590
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
            Picture         =   "frmAllSection.frx":0C07
            Key             =   "StopBus"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllSection.frx":0D61
            Key             =   "RunBus"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllSection.frx":0EBB
            Key             =   "FlowStop"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllSection.frx":1255
            Key             =   "Flow"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllSection.frx":15EF
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllSection.frx":1ECB
            Key             =   "RunStopBus"
         EndProperty
      EndProperty
   End
   Begin ActiveBar2LibraryCtl.ActiveBar2 abAction 
      Height          =   5355
      Left            =   8880
      TabIndex        =   3
      Top             =   1050
      Width           =   1500
      _LayoutVersion  =   1
      _ExtentX        =   2646
      _ExtentY        =   9446
      _DataPath       =   ""
      Bands           =   "frmAllSection.frx":2265
   End
   Begin MSComctlLib.ListView lvSection 
      Height          =   4635
      Left            =   15
      TabIndex        =   0
      Top             =   1080
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
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "路段代码"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "路段名称"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "起点站"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "终点站"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "里程数"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "公路等级"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "地区"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "路径号"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu pmnu_PopupMenu 
      Caption         =   "弹出菜单"
      Visible         =   0   'False
      Begin VB.Menu pmnu_AddSection 
         Caption         =   "新增(&A)"
      End
      Begin VB.Menu pmnu_EditSection 
         Caption         =   "编辑(&E)"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_DeleteSection 
         Caption         =   "删除(&D)"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmAllSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**********************************************************
'* Source File Name:frmAllSection.frm
'* Project Name:RTBusMan
'* Engineer:
'* Data Generated:2002/08/30
'* Last Revision Date:2002/08/30
'* Brief Description:路段控制台
'* Relational Document:
'**********************************************************
'Public bIsShow As Boolean
Private m_oBaseInfo As New BaseInfo '基本信息对象 BaseInfo
Private m_oSection As New Section '路段对象 Section
Private m_oStation As New Station
Public m_szSectionID As String

Private Sub abAction_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
    Case "act_AddSection"
        AddSection
    Case "act_EditSection"
        EditSection
    Case "act_DeleteSection"
        DeleteSection
    End Select
End Sub

Private Sub cboStation_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Static nIndex As Integer
    On Error GoTo ErrorHandle
    If KeyCode = vbKeyReturn Then
        If Val(ResolveDisplay(cboStation.Text)) > 0 Then
            m_oStation.Identify ResolveDisplay(cboStation.Text)
        Else
            m_oStation.Identify , ResolveDisplay(cboStation.Text)
        End If
        cboStation.Text = m_oStation.StationID & "[" & m_oStation.StationName & "]"
        If Trim(cboStation.Text) <> "" Then
            For i = 1 To cboStation.ListCount
                If Trim(cboStation.Text) = ResolveDisplay(cboStation.List(i)) Then
                    cmdFind_Click
                    Exit Sub
                End If
            Next
        If cboStation.ListCount >= 10 Then
            cboStation.List(nIndex + 1) = m_oStation.StationID & "[" & m_oStation.StationName & "]"
            cboStation.Text = cboStation.List(nIndex + 1)
            nIndex = nIndex + 1
            If nIndex >= 9 Then
                nIndex = 0
            End If
                cboStation.Text = m_oStation.StationID & "[" & m_oStation.StationName & "]"
                cmdFind_Click
            Else
                cboStation.AddItem m_oStation.StationID & "[" & m_oStation.StationName & "]"
                cboStation.Text = m_oStation.StationID & "[" & m_oStation.StationName & "]"
                cmdFind_Click
            End If
        End If
    End If
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

'===================================================
'Modify Date：2002-11-19
'Author:fl
'Reamrk:输入途经站
'===================================================b

Private Sub cboStation_KeyPress(KeyAscii As Integer)
      
End Sub

'**************************************************
'Member Code:S1
'Brief Description:按下查询按钮，显示该范围内的线路
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/25
'**************************************************
Public Sub cmdFind_Click()
    Dim szaSection() As String
    Dim liTemp As ListItem
    Dim i As Integer, nCount As Integer
    On Error GoTo ErrorHandle
    SetBusy
    lvSection.ListItems.Clear '先清空lvSection
    '--------------------------------------------------
    '如果是(全部)则在传入查询线路是不传入站点代码则表示查询所有，站点代码为
    '线路的终点站代码
    ShowSBInfo "获得路段信息..."
    If cboStation.Text = "(全部)" Then
        szaSection = m_oBaseInfo.GetSection(ResolveDisplay(txtRouteID.Text), txtSectionID.Text)
    Else
        szaSection = m_oBaseInfo.GetSection(ResolveDisplay(txtRouteID.Text), txtSectionID.Text, ResolveDisplay(cboStation.Text), ResolveDisplay(txtStartStationID.Text))
    End If
    '--------------------------------------------------
    '在lvSection中填充线路信息返回值的顺序是线路代码、线路名称、终点站、终点站名称、里程数、状态
    nCount = ArrayLength(szaSection)
    If nCount = 0 Then
       SetNormal
       WriteProcessBar
       MsgBox "没有您需要的数据,请检查查询条件", vbInformation + vbOKOnly, Me.Caption
       Exit Sub
    End If
    WriteProcessBar , nCount, , True
    For i = 1 To nCount
        ShowSBInfo "获得路段"  '& szaSection(i, 2) & "信息", , i
        Set liTemp = lvSection.ListItems.Add(, , szaSection(i, 1)) ', , "Section")
        liTemp.SubItems(1) = szaSection(i, 2)
        liTemp.SubItems(2) = szaSection(i, 3)
        liTemp.SubItems(3) = szaSection(i, 4)
        liTemp.SubItems(4) = szaSection(i, 5)
        liTemp.SubItems(5) = szaSection(i, 6)
        liTemp.SubItems(6) = szaSection(i, 7)
        liTemp.SubItems(7) = szaSection(i, 8)
    Next
    SetNormal
    WriteProcessBar False
    SetMenuEnabled
    Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub Form_Activate()
'    SetInfoBar SNSection
'    SetInfoBar SNReport

End Sub

Private Sub Form_Deactivate()
'    SetInfoBar NOStatus
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Const cnMargin = 50
    ptShowInfo.Left = 0
    ptShowInfo.Top = 0
    ptShowInfo.Width = Me.ScaleWidth
    lvSection.Left = cnMargin
    lvSection.Top = ptShowInfo.Height + cnMargin
    lvSection.Width = Me.ScaleWidth - IIf(abAction.Visible, abAction.Width, 0) - 2 * cnMargin
    lvSection.Height = Me.ScaleHeight - ptShowInfo.Height - 2 * cnMargin
    '当操作条关闭时间处理
    If abAction.Visible Then
        abAction.Move lvSection.Width + cnMargin, lvSection.Top
        abAction.Height = lvSection.Height
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
'    SetInfoBar NOStatus
'    bIsShow = False
'    SaveDescStation
'    SaveListViewWidth lvSection
End Sub

Private Sub Form_Load()
'    bIsShow = True

    m_oBaseInfo.Init g_oActiveUser
    m_oSection.Init g_oActiveUser
    m_oStation.Init g_oActiveUser
'    LoadDescStation
'    SetListViewWidth lvSection
End Sub

Private Sub lvSection_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvSection, ColumnHeader.Index
End Sub

Private Sub lvSection_DblClick()
    EditSection
End Sub

'Private Sub lvSection_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'    Case cszKeyPopMenu
'        lvSection_MouseDown vbRightButton, Shift, 1, 1
'    End Select
'End Sub

Private Sub lvSection_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyReturn
       lvSection_DblClick
End Select
End Sub

Private Sub lvSection_LostFocus()
'    MDIScheme.mnu_Section.Enabled = False
'    MDIScheme.mnu_DeleteSection.Enabled = False
End Sub

Private Sub lvSection_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu pmnu_PopupMenu
    End If
End Sub

Private Sub pmnu_AddSection_Click()
    AddSection
End Sub

Private Sub pmnu_DeleteSection_Click()
    DeleteSection
End Sub

Private Sub pmnu_EditSection_Click()
    EditSection
End Sub

Private Sub txtRouteID_Click()
Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectRoute(False)
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtRouteID.Text = aszTemp(1, 1) & "[" & Trim(aszTemp(1, 2)) & "]"
End Sub

Private Sub txtRouteId_GotFocus()
    txtRouteID.SelStart = 0
    txtRouteID.SelLength = 100
End Sub
Public Sub AddSection()
    frmSection.m_eStatus = EFS_AddNew
    frmSection.m_bIsParent = True
    frmSection.Show vbModal
End Sub


'**************************************************
'Member Code:S3
'Brief Description:编辑路段
'Engineer:
'Date Generated:2002/11/16
'Last Revision Date:2002/11/16
'**************************************************
Public Sub EditSection()
    If lvSection.SelectedItem Is Nothing Then Exit Sub
    frmSection.m_szSectionID = lvSection.SelectedItem.Text
    frmSection.m_eStatus = EFS_Modify
    frmSection.m_bIsParent = True
    frmSection.Show vbModal
End Sub

Public Sub DeleteSection()
    On Error GoTo ErrorHandle
    Dim nResult As VbMsgBoxResult
    nResult = MsgBox("是否要删除路段[" & Trim(lvSection.SelectedItem.ListSubItems(1).Text) & "]", vbQuestion + vbYesNo + vbDefaultButton2, "路段管理")
    If nResult = vbYes Then
    m_oSection.Identify (Trim(lvSection.SelectedItem.Text))
    m_oSection.Delete
    lvSection.ListItems.Remove lvSection.SelectedItem.Index
    End If
    SetMenuEnabled
Exit Sub
ErrorHandle:
    If err.Number = 91 Then Exit Sub
    ShowErrorMsg
End Sub

Private Sub txtStation_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyReturn
       cmdFind_Click
End Select
End Sub


Private Sub txtSectionId_GotFocus()
    txtSectionID.SelStart = 0
    txtSectionID.SelLength = 100
End Sub

Public Sub UpdateList(SectionID As String)
    Dim szaSection() As String
    Dim liTemp As ListItem
    Set liTemp = lvSection.FindItem(SectionID, , , lvwPartial)
    If liTemp Is Nothing Then Exit Sub
    szaSection = m_oBaseInfo.GetSection(, SectionID)
    liTemp.ListSubItems(1).Text = szaSection(1, 2)
    liTemp.ListSubItems(2).Text = szaSection(1, 3)
    liTemp.ListSubItems(3).Text = szaSection(1, 4)
    liTemp.ListSubItems(4).Text = szaSection(1, 5)
    liTemp.ListSubItems(5).Text = szaSection(1, 6)
    liTemp.ListSubItems(6).Text = szaSection(1, 7)
    liTemp.ListSubItems(7).Text = szaSection(1, 8)
End Sub

Public Sub AddList(SectionID As String)
    Dim szaSection() As String
    Dim liTemp As ListItem
    szaSection = m_oBaseInfo.GetSection(, SectionID)
    Set liTemp = lvSection.ListItems.Add(, , szaSection(1, 1)) ', , "Section")
    liTemp.SubItems(1) = szaSection(1, 2)
    liTemp.SubItems(2) = szaSection(1, 3)
    liTemp.SubItems(3) = szaSection(1, 4)
    liTemp.SubItems(4) = szaSection(1, 5)
    liTemp.SubItems(5) = szaSection(1, 6)
    liTemp.SubItems(6) = szaSection(1, 7)
    liTemp.SubItems(7) = szaSection(1, 8)
    SetMenuEnabled
End Sub



Private Sub SetMenuEnabled()
    '设置菜单的可用性
    Dim bEnabled As Boolean
    If lvSection.ListItems.Count > 0 Then
        bEnabled = True
    Else
        bEnabled = False
    End If
    pmnu_EditSection.Enabled = bEnabled
    pmnu_DeleteSection.Enabled = bEnabled
    With abAction.Bands("bndActionTabs").ChildBands("actSection")
        .Tools("act_EditSection").Enabled = bEnabled
        .Tools("act_DeleteSection").Enabled = bEnabled
    End With
'    With MDIScheme.abMenuTool.Bands("mnu_StationMan")
'        .Tools("mnu_StationMan_Info").Enabled = bEnabled
'        .Tools("mnu_StationMan_Del").Enabled = bEnabled
'    End With
End Sub




