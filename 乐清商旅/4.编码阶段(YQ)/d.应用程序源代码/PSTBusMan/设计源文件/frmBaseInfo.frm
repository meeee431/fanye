VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmBaseInfo 
   BackColor       =   &H00E0E0E0&
   Caption         =   "基本信息管理"
   ClientHeight    =   6420
   ClientLeft      =   1710
   ClientTop       =   2565
   ClientWidth     =   9675
   HelpContextID   =   2001601
   Icon            =   "frmBaseInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11475
   ScaleWidth      =   23880
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ptLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5985
      Left            =   120
      ScaleHeight     =   5985
      ScaleWidth      =   2265
      TabIndex        =   5
      Top             =   30
      Width           =   2265
      Begin MSComctlLib.TreeView tvBaseItem 
         Height          =   3915
         Left            =   0
         TabIndex        =   6
         Top             =   1980
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   6906
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "bigImgLists"
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image imgTreeTitle 
         Height          =   1800
         Left            =   150
         Picture         =   "frmBaseInfo.frx":08CA
         Top             =   0
         Width           =   2250
      End
   End
   Begin RTComctl3.Spliter spMove 
      Height          =   915
      Left            =   2460
      TabIndex        =   3
      Top             =   3300
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   1614
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
   Begin VB.PictureBox ptRight 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5925
      Left            =   2730
      ScaleHeight     =   5925
      ScaleWidth      =   6765
      TabIndex        =   0
      Top             =   60
      Width           =   6765
      Begin MSComctlLib.ImageList smallImgLists 
         Left            =   2340
         Top             =   1200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":29B3
               Key             =   "seattype"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":3805
               Key             =   "owner"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":3B9F
               Key             =   "company"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":49F1
               Key             =   "area"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":5843
               Key             =   "tickettype"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":6895
               Key             =   "checkgate"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":6C2F
               Key             =   "bustype"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":7A81
               Key             =   "vehicletype"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":7E1B
               Key             =   "vehiclestop"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":81B5
               Key             =   "vehiclerun"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":8A8F
               Key             =   "roadlevel"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList bigImgLists 
         Left            =   3390
         Top             =   1200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":8BE9
               Key             =   "owner"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":98C3
               Key             =   "seattype"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":A19D
               Key             =   "vehicle"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":AA77
               Key             =   "roadlevel"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":B351
               Key             =   "company"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":BC2B
               Key             =   "area"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":C505
               Key             =   "vehicletype"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":D1DF
               Key             =   "bustype"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":DAB9
               Key             =   "checkgate"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":E16E
               Key             =   "open"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":EE48
               Key             =   "tickettype"
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox ptShowInfo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   990
         Left            =   -60
         ScaleHeight     =   990
         ScaleWidth      =   6615
         TabIndex        =   1
         Top             =   0
         Width           =   6615
         Begin VB.Image imgObject 
            Height          =   480
            Left            =   1800
            Top             =   300
            Width           =   480
         End
         Begin VB.Label lblTitlePrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "检票口"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   2430
            TabIndex        =   2
            Top             =   510
            Width           =   765
         End
         Begin VB.Image Image1 
            Height          =   1275
            Left            =   60
            Picture         =   "frmBaseInfo.frx":FB22
            Top             =   30
            Width           =   2010
         End
      End
      Begin ActiveBar2LibraryCtl.ActiveBar2 abAction 
         Height          =   4875
         Left            =   5190
         TabIndex        =   4
         Top             =   1020
         Width           =   1485
         _LayoutVersion  =   1
         _ExtentX        =   2619
         _ExtentY        =   8599
         _DataPath       =   ""
         Bands           =   "frmBaseInfo.frx":10FF5
      End
      Begin MSComctlLib.ListView lvObject 
         Height          =   4515
         Left            =   240
         TabIndex        =   7
         Top             =   1650
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   7964
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "bigImgLists"
         SmallIcons      =   "smallImgLists"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "代码"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "名称"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "备注"
            Object.Width           =   4939
         EndProperty
      End
   End
   Begin VB.Menu pmnu_Action 
      Caption         =   "操作"
      Visible         =   0   'False
      Begin VB.Menu pmnu_Add 
         Caption         =   "新增(&A)"
      End
      Begin VB.Menu pmnu_BaseInfo 
         Caption         =   "属性(&R)"
      End
      Begin VB.Menu pmnu_Del 
         Caption         =   "删除(&D)"
      End
   End
End
Attribute VB_Name = "frmBaseInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===================================================
'Modify Date：2002-11-19
'Author:陆勇庆
'Reamrk:修改了对检票口的处理，加入所属车站属性
'===================================================

'以下变量定义
Dim m_oBaseInfo As New BaseInfo
Private Sub abAction_BandClose(ByVal Band As ActiveBar2LibraryCtl.Band)
    If Band.name = "bndActionTabs" Then
        abAction.Visible = False
        Call ptRight_Resize
    End If
End Sub

'Private Sub abAction_BandOpen(ByVal Band As ActiveBar2LibraryCtl.Band, ByVal Cancel As ActiveBar2LibraryCtl.ReturnBool)
''    abAction.Visible = True
'    If Band.name = "bndActionTabs" Then
'        Call ptRight_Resize
'    End If
'End Sub

Private Sub abAction_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "mnu_Add", "act_BaseMan_Add"
            AddObject
        Case "mnu_BaseInfo", "act_BaseMan_BaseInfo"
            EditObject
        Case "mnu_Del", "act_BaseMan_Del"
            DeleteObject
    End Select
End Sub





Private Sub Form_Activate()
    MDIScheme.ActiveToolBar "baseinfo", True
    SetMenuEnabled
'    ActiveSystemToolBar True
    spMove.LayoutIt
    
    WriteTitleBar "基本信息"
End Sub

Private Sub Form_Deactivate()
    MDIScheme.ActiveToolBar "baseinfo", False
'    ActiveSystemToolBar False
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandle
    m_oBaseInfo.Init g_oActiveUser
    
    spMove.InitSpliter ptLeft, ptRight
    FillBaseItemTree
    FillItemLists
    
    AlignHeadWidth Me.name, lvObject
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'设置基本信息项
Private Sub FillBaseItemTree()
    With tvBaseItem.Nodes
        .Add , , "KArea", "地区", "area", "open"
        .Add , , "KRoadLevel", "公路等级", "roadlevel", "open"
        .Add , , "KCompany", "参运公司", "company", "open"
        .Add , , "KVehicleType", "车型", "vehicletype", "open"
        .Add , , "KVehicle", "参运车辆", "vehicle", "open"
        .Add , , "KOwner", "车主", "owner", "open"
        .Add , , "KCheckGate", "检票口", "checkgate", "open"
        .Add , , "KBusType", "车次种类", "bustype", "open"
        .Add , , "KSeatType", "座位类型", "seattype", "open"
        tvBaseItem.Nodes(1).Selected = True
    End With
End Sub

Private Sub Form_Resize()
    spMove.LayoutIt
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIScheme.ActiveToolBar "baseinfo", False
'    ActiveSystemToolBar False
    '保存列头
    SaveHeadWidth Me.name, lvObject
End Sub

Private Sub lvObject_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvObject, ColumnHeader.Index
End Sub

Private Sub lvObject_DblClick()
    EditObject
End Sub

Private Sub lvObject_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            If Not lvObject.SelectedItem Is Nothing Then
                DeleteObject
            End If
    End Select
End Sub

Private Sub lvObject_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'        Case vbKeyReturn
'            lvObject_DblClick
'    End Select
End Sub

Private Sub lvObject_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
'        Dim oHit As ListItem
'        Set oHit = lvObject.HitTest(X, Y)
'        If Not oHit Is Nothing Then oHit.Selected = True
'        abAction.Bands("mnu_Action").PopupMenu
        PopupMenu pmnu_Action
    End If
End Sub

Private Sub pmnu_Add_Click()
    AddObject
End Sub

Private Sub pmnu_BaseInfo_Click()
    EditObject
End Sub

Private Sub pmnu_Del_Click()
    DeleteObject
End Sub

Private Sub ptLeft_Resize()
On Error Resume Next
    Const cnMargin = 50
    imgTreeTitle.Left = 0
    imgTreeTitle.Top = 0
    tvBaseItem.Left = imgTreeTitle.Left + cnMargin
    tvBaseItem.Top = imgTreeTitle.Top + imgTreeTitle.Height
    tvBaseItem.Width = ptLeft.ScaleWidth - 2 * cnMargin
    tvBaseItem.Height = ptLeft.ScaleHeight - imgTreeTitle.Height - cnMargin
End Sub

Private Sub ptRight_Resize()
On Error Resume Next
    Const cnMargin = 50
    ptShowInfo.Left = 0
    ptShowInfo.Top = 0
    ptShowInfo.Width = ptRight.ScaleWidth
    lvObject.Left = cnMargin
    lvObject.Top = ptShowInfo.Height + cnMargin
    lvObject.Width = ptRight.ScaleWidth - IIf(abAction.Visible, abAction.Width, 0) - 2 * cnMargin
    lvObject.Height = ptRight.ScaleHeight - ptShowInfo.Height - 2 * cnMargin
    '当操作条关闭时间处理
    If abAction.Visible Then
        abAction.Move lvObject.Width + cnMargin, lvObject.Top
        abAction.Height = lvObject.Height
    End If
End Sub

Private Sub tvBaseItem_NodeClick(ByVal Node As MSComctlLib.Node)
    FillItemLists
    SetMenuEnabled
End Sub

Public Sub AddObject()
On Error GoTo ErrHandle
    Dim szSelectKey As String
    szSelectKey = tvBaseItem.SelectedItem.Key
    Select Case szSelectKey
         Case "KArea"
          frmArea.Status = EFS_AddNew
          frmArea.Show vbModal
         Case "KRoadLevel"
          frmRoadLevel.Status = EFS_AddNew
          frmRoadLevel.Show vbModal
         Case "KCompany"
          frmCompany.Status = EFS_AddNew
          frmCompany.Show vbModal
         Case "KVehicleType"
          frmVehicleType.Status = EFS_AddNew
          frmVehicleType.Show vbModal
          Case "KVehicle"
            frmVehicle.m_bIsParent = True
            frmVehicle.Status = EFS_AddNew
            frmVehicle.Show vbModal
         Case "KOwner"
          frmBusOwner.Status = EFS_AddNew
          frmBusOwner.Show vbModal
         Case "KCheckGate"
          frmCheckDoor.Status = EFS_AddNew
          frmCheckDoor.Show vbModal
        Case "KBusType"
          frmBusType.Status = EFS_AddNew
          frmBusType.Show vbModal
        Case "KSeatType"
          frmSeatType.Status = EFS_AddNew
          frmSeatType.Show vbModal
    End Select
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'更改当前选中的项目
Public Sub EditObject()
    On Error GoTo ErrHandle
'    If lvObject.SelectedItem Is Nothing Then
'        MsgBox "请先选择需要修改的项目!", vbInformation, "基本信息"
'        Exit Sub
'    End If
    If lvObject.SelectedItem Is Nothing Then Exit Sub
    
    
    Select Case tvBaseItem.SelectedItem.Key
        Case "KArea"
            frmArea.Status = EFS_Modify
            frmArea.mszAreaID = lvObject.SelectedItem.Text
            frmArea.Show vbModal
        Case "KRoadLevel"
            frmRoadLevel.Status = EFS_Modify
            frmRoadLevel.mszRoadLevel = lvObject.SelectedItem.Text
            frmRoadLevel.Show vbModal
        Case "KCompany"
            frmCompany.Status = EFS_Modify
            frmCompany.mszCompanyID = lvObject.SelectedItem.Text
            frmCompany.Show vbModal
        Case "KVehicleType"
            frmVehicleType.Status = EFS_Modify
            frmVehicleType.mszVehicleType = lvObject.SelectedItem.Text
            frmVehicleType.Show vbModal
        Case "KVehicle"
            frmVehicle.m_bIsParent = True
            frmVehicle.Status = EFS_Modify
            frmVehicle.mszVehicleId = lvObject.SelectedItem.Text
            frmVehicle.Show vbModal
        Case "KOwner"
            frmBusOwner.Status = EFS_Modify
            frmBusOwner.m_szOwnerID = lvObject.SelectedItem
            frmBusOwner.Show vbModal
        Case "KCheckGate"
            frmCheckDoor.Status = EFS_Modify
            frmCheckDoor.mszCheckID = lvObject.SelectedItem
            frmCheckDoor.Show vbModal
        Case "KBusType"
            frmBusType.Status = EFS_Modify
            frmBusType.mszBusTypeID = lvObject.SelectedItem
            frmBusType.Show vbModal
        Case "KSeatType"
            frmSeatType.Status = EFS_Modify
            frmSeatType.mszSeatTypeID = lvObject.SelectedItem
            frmSeatType.Show vbModal
    End Select

    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'列出基本信息
Private Sub FillItemLists()
On Error GoTo ErrHandle
    Dim aszItems() As String
    Dim oListItem As ListItem
    
    lvObject.ListItems.Clear
    ShowSBInfo "正在查询，请稍等..."
    MousePointer = vbHourglass
    
    lblTitlePrompt.Caption = tvBaseItem.SelectedItem.Text
    Dim szSelectKey As String
    szSelectKey = tvBaseItem.SelectedItem.Key
    '得到基本信息
    Select Case szSelectKey
        Case "KArea"            '地区
            aszItems = m_oBaseInfo.GetAllArea
        Case "KRoadLevel"       '公路等级
            aszItems = m_oBaseInfo.GetAllRoadLevel
        Case "KVehicleType"     '车型
            aszItems = m_oBaseInfo.GetAllVehicleModel
        Case "KVehicle"         '车辆
            frmQueryVehicle.Show vbModal
            If Not frmQueryVehicle.IsCancel Then
                aszItems = GetVehicleItems
            End If
            Unload frmQueryVehicle
        Case "KOwner"           '车主
            aszItems = m_oBaseInfo.GetOwner
        Case "KCompany"         '参运公司
            aszItems = m_oBaseInfo.GetCompany
        Case "KCheckGate"       '检票口
            aszItems = m_oBaseInfo.GetAllCheckGate
        Case "KBusType"         '车次类别
            aszItems = m_oBaseInfo.GetAllBusType
        Case "KSeatType"        '座位类别
            aszItems = m_oBaseInfo.GetAllSeatType
    End Select
    
    '填充列表
    Dim nCount As Integer, i As Integer
    nCount = ArrayLength(aszItems)
    WriteProcessBar , , nCount
    Set imgObject.Picture = bigImgLists.ListImages(LCase(Mid(szSelectKey, 2))).Picture
    Dim aszTmpItem(0 To 3) As String
    For i = 1 To nCount
        WriteProcessBar , i, nCount, "得到对象[" & aszItems(i, 2) & "]"
        If szSelectKey = "KVehicle" Then
            aszTmpItem(0) = aszItems(i, 0)
        End If
        aszTmpItem(1) = aszItems(i, 1)
        aszTmpItem(2) = aszItems(i, 2)
        If szSelectKey = "KCheckGate" Then
            aszTmpItem(3) = EncodeString(aszItems(i, 6)) & aszItems(i, 3)
        ElseIf szSelectKey = "KOwner" Then   '车主
        
            aszTmpItem(3) = aszItems(i, 4)
        Else
            aszTmpItem(3) = aszItems(i, 3)
        End If
        
        
        AddList aszTmpItem
    Next
    lvObject.Refresh
    If lvObject.ListItems.Count > 0 Then lvObject.ListItems(1).Selected = True
    WriteProcessBar False
    ShowSBInfo "共" & nCount & "个对象", ESB_ResultCountInfo
    ShowSBInfo ""
    
    SetNormal
    Exit Sub
ErrHandle:
    SetNormal
    ShowErrorMsg
End Sub
'查询得到车辆数组
Private Function GetVehicleItems() As String()
    Dim aszVehicles() As String
    Dim aszReturn() As String
    Dim szCompany As String
    Dim szOwner As String
    Dim szBusType As String
    Dim szLicense As String
    Dim szVehicle As String
    Dim i As Integer, nCount As Integer
    With frmQueryVehicle
    szVehicle = Trim(.txtVehicle.Text)
    szCompany = IIf(Trim(.txtCompany.Text) = "", "", ResolveDisplay(.txtCompany.Text))
    szOwner = IIf(Trim(.txtBusOwner.Text) = "", "", ResolveDisplay(.txtBusOwner.Text))
    szLicense = IIf(Trim(.txtLicense.Text) = "", "", .txtLicense.Text)
    szBusType = IIf(Trim(.txtVehicleType.Text) = "", "", ResolveDisplay(.txtVehicleType.Text))
    End With
    
    Dim oVehicle As New BaseInfo
    oVehicle.Init g_oActiveUser
    aszVehicles = oVehicle.GetVehicle(szVehicle, szCompany, szOwner, szBusType, szLicense, True)
    nCount = ArrayLength(aszVehicles)
    If nCount > 0 Then ReDim aszReturn(1 To nCount, 0 To 3)
    For i = 1 To nCount
        aszReturn(i, 1) = Trim(aszVehicles(i, 1))
        aszReturn(i, 2) = Trim(aszVehicles(i, 2))
        aszReturn(i, 3) = EncodeString("所属公司:" & Trim(aszVehicles(i, 4))) & _
                        EncodeString("车主:" & Trim(aszVehicles(i, 5))) & _
                        EncodeString("车型:" & Trim(aszVehicles(i, 8))) & _
                        EncodeString("座位数:" & Trim(aszVehicles(i, 3)))
        If Val(aszVehicles(i, 6)) <> ST_VehicleRun Then
            aszReturn(i, 0) = "STOP"    '停班车辆
        End If
    Next
    GetVehicleItems = aszReturn
End Function

'删除对象
Public Sub DeleteObject()
    On Error GoTo ErrHandle
    Dim oBus As Object
    Dim szTmp As String
'    If lvObject.SelectedItem Is Nothing Then
'        MsgBox "请先选择需要删除的项目!", vbInformation, "基本信息"
'        Exit Sub
'    End If
    Select Case tvBaseItem.SelectedItem.Key
        Case "KArea"
            Set oBus = CreateObject("STBase.Area")
            szTmp = "地区"
        Case "KRoadLevel"
            Set oBus = CreateObject("STBase.RoadLevel")
            szTmp = "公路等级"
        Case "KCompany"
            Set oBus = CreateObject("STBase.Company")
            szTmp = "参运公司"
        Case "KVehicleType"
            Set oBus = CreateObject("STBase.VehicleModel")
            szTmp = "车型"
        Case "KVehicle"
            Set oBus = CreateObject("STBase.Vehicle")
            szTmp = "车辆"
        Case "KOwner"
            Set oBus = CreateObject("STBase.Owner")
            szTmp = "车主"
        Case "KCheckGate"
            Set oBus = CreateObject("STBase.CheckGate")
            szTmp = "检票口"
        Case "KBusType"
            Set oBus = CreateObject("STBase.BusType")
            szTmp = "车次类型"
        Case "KSeatType"
            Set oBus = CreateObject("STBase.SeatType")
            szTmp = "座位类型"
    End Select
    Dim vbYesOrNo As Integer
    vbYesOrNo = MsgBox("是否真的删除" & szTmp & "[" & lvObject.SelectedItem & "]", vbQuestion + vbYesNo + vbDefaultButton2, "删除基本信息")
    If vbYesOrNo = vbYes Then
          oBus.Init g_oActiveUser
          oBus.Identify lvObject.SelectedItem.Text
          oBus.Delete
          lvObject.ListItems.Remove lvObject.SelectedItem.Index
    End If
    SetMenuEnabled
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'添加项目到listview
Public Sub AddList(paszItems As Variant, Optional pbEnsure As Boolean)
    'pbEnsure 是否滚到此行
    Dim oListItem As ListItem
    Dim szSelectKey As String
    szSelectKey = tvBaseItem.SelectedItem.Key
    Set oListItem = lvObject.ListItems.Add(, , Trim(paszItems(1)))
    oListItem.SubItems(1) = paszItems(2)
    If szSelectKey = "KArea" Then  '地区的特殊地方
        oListItem.SmallIcon = LCase(Mid(szSelectKey, 2))
        Select Case Val(paszItems(3))
               Case EA_nInCity
                    paszItems(3) = "市内"
               Case EA_nOutCity
                    paszItems(3) = "市外"
               Case EA_nOutProvince
                    paszItems(3) = "省外"
         End Select
    End If
'    If szSelectKey = "KCheckGate" Then
'        paszItems(3) = EncodeString(paszItems(5)) & paszItems(3)
'    End If
    oListItem.SubItems(2) = paszItems(3)
    If szSelectKey = "KVehicle" Then
        If paszItems(0) = "STOP" Then
            oListItem.SmallIcon = "vehiclestop"
            SetListViewLineColor lvObject, oListItem.Index, vbRed
        Else
            oListItem.SmallIcon = "vehiclerun"
        End If
    Else
        oListItem.SmallIcon = LCase(Mid(szSelectKey, 2))
    End If
    oListItem.Selected = True
    If pbEnsure Then oListItem.EnsureVisible
    SetMenuEnabled
End Sub
'更改项目到listview
Public Sub UpdateList(paszItems As Variant)
    Dim oListItem As ListItem
    Dim szSelectKey As String
    szSelectKey = tvBaseItem.SelectedItem.Key
    Set oListItem = lvObject.SelectedItem
    If oListItem Is Nothing Then Exit Sub
    oListItem.SubItems(1) = paszItems(2)
    If szSelectKey = "KArea" Then
       Select Case Val(paszItems(3))
              Case EA_nInCity
                   oListItem.SubItems(2) = "市内"
              Case EA_nOutCity
                   oListItem.SubItems(2) = "市外"
              Case EA_nOutProvince
                   oListItem.SubItems(2) = "省外"
        End Select
    Else
        oListItem.SubItems(2) = paszItems(3)
    End If
    If szSelectKey = "KVehicle" Then
        If paszItems(0) = "STOP" Then
            oListItem.SmallIcon = "vehiclestop"
            SetListViewLineColor lvObject, oListItem.Index, vbRed
        Else
            oListItem.SmallIcon = "vehiclerun"
            SetListViewLineColor lvObject, oListItem.Index, vbBlack
        End If
        lvObject.Refresh
    End If
End Sub

Private Sub SetMenuEnabled()
    Dim bEnabled As Boolean
    If lvObject.ListItems.Count > 0 Then
        bEnabled = True
    Else
        bEnabled = False
    End If
    MDIScheme.abMenuTool.Bands("mnu_BaseMan").Tools("mnu_BaseMan_BaseInfo").Enabled = bEnabled
    MDIScheme.abMenuTool.Bands("mnu_BaseMan").Tools("mnu_BaseMan_Del").Enabled = bEnabled
    
    abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_BaseInfo").Enabled = bEnabled
    abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Del").Enabled = bEnabled
    
    pmnu_BaseInfo.Enabled = bEnabled
    pmnu_Del.Enabled = bEnabled
    
End Sub
