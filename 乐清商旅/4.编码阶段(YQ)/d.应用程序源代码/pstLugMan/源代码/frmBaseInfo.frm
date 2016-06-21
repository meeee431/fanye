VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmBaseInfo 
   BackColor       =   &H00E0E0E0&
   Caption         =   "基本信息管理"
   ClientHeight    =   8520
   ClientLeft      =   1665
   ClientTop       =   3075
   ClientWidth     =   11940
   HelpContextID   =   7000220
   Icon            =   "frmBaseInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   11940
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList bigImgLists 
      Left            =   10020
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfo.frx":038A
            Key             =   "luggagetype"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfo.frx":1064
            Key             =   "priceformula"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfo.frx":193E
            Key             =   "protocol"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfo.frx":2618
            Key             =   "vehicle"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfo.frx":2EF2
            Key             =   "formula"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfo.frx":304C
            Key             =   "priceitem"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox ptLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5985
      Left            =   120
      ScaleHeight     =   5985
      ScaleWidth      =   2445
      TabIndex        =   5
      Top             =   30
      Width           =   2445
      Begin MSComctlLib.TreeView tvBaseItem 
         Height          =   3915
         Left            =   0
         TabIndex        =   6
         Top             =   2040
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   6906
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   0
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
         Height          =   2700
         Left            =   -480
         Picture         =   "frmBaseInfo.frx":33E6
         Top             =   -480
         Width           =   3300
      End
   End
   Begin RTComctl3.Spliter spMove 
      Height          =   915
      Left            =   2550
      TabIndex        =   3
      Top             =   3285
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   1614
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
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":80C9
               Key             =   "luggagetype"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":8463
               Key             =   "protocol"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":87FD
               Key             =   "vehicle"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":8B97
               Key             =   "priceformula"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":9471
               Key             =   "formula"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":95CB
               Key             =   "priceitem"
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
            Picture         =   "frmBaseInfo.frx":9965
            Top             =   30
            Width           =   2010
         End
      End
      Begin ActiveBar2LibraryCtl.ActiveBar2 abAction 
         Height          =   4875
         Left            =   5280
         TabIndex        =   4
         Top             =   1035
         Width           =   1485
         _LayoutVersion  =   1
         _ExtentX        =   2619
         _ExtentY        =   8599
         _DataPath       =   ""
         Bands           =   "frmBaseInfo.frx":AE38
      End
      Begin MSComctlLib.ListView lvObject 
         Height          =   4515
         Left            =   600
         TabIndex        =   7
         Top             =   1440
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   7964
         View            =   3
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "代码"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "名称"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "备注"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "默认设置"
            Object.Width           =   2540
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
Public szTmp As String

Enum EOperStatus '操作状态
    EOS_Add = 0
    EOS_Modify = 1
    EOS_Delete = 2
End Enum

'以下变量定义
'Dim m_obase As New BaseInfo
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
        Case "act_BaseMan_Edit"
            frmEditProtocol.txtLicense.Text = lvObject.ListItems(1).Text
            frmEditProtocol.txtAcceptType.Text = lvObject.ListItems(lvObject.SelectedItem.Index).ListSubItems(2).Text
'            frmEditProtocol.asFlatTextBox1.Text = lvObject.ListItems(lvObject.SelectedItem.Index).ListSubItems(1).Text
            frmEditProtocol.txtVehiclID.Text = lvObject.ListItems(lvObject.SelectedItem.Index).ListSubItems(3).Text
            frmEditProtocol.Show
    End Select
End Sub
Private Sub Form_Activate()
'    mdiMain.ActiveToolBar "baseinfo", True
    SetMenuEnabled
'    ActiveSystemToolBar True
    spMove.LayoutIt

    WriteTitleBar "基本信息"
End Sub

Private Sub Form_Deactivate()
'    mdiMain.ActiveToolBar "baseinfo", False
'    ActiveSystemToolBar False

End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandle
'    m_obase.Init m_oAUser
    mdiMain.ActiveToolBar True
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
'        .Add , , "KProtocol", "设置行包拆算协议", "protocol"
'        .Add , , "KVehicle", "车辆协议情况查询", "vehicle"
'        .Add , , "KFormula", "拆算协议计算公式", "formula"
        .Add , , "KLuggageType", "行包种类", "luggagetype"
        .Add , , "KPriceItem", "行包收费项", "priceitem"
        .Add , , "KPriceFormula", "行包计算公式", "priceformula"
        tvBaseItem.Nodes(1).Selected = True
    End With
End Sub

Private Sub Form_Resize()
    spMove.LayoutIt
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    mdiMain.ActiveToolBar "baseinfo", False
'    ActiveSystemToolBar False
    '保存列头
    SaveHeadWidth Me.name, lvObject
End Sub

Private Sub lvObject_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvObject, ColumnHeader.Index
End Sub

Private Sub lvObject_DblClick()
    Dim i As Integer
    Dim szselectKey As Variant
    szselectKey = tvBaseItem.SelectedItem.Key
    If szselectKey <> "KVehicle" Then
        EditObject
    Else
        frmEditProtocol.txtLicense.Text = lvObject.SelectedItem.Text
        frmEditProtocol.txtAcceptType.Text = lvObject.ListItems(lvObject.SelectedItem.Index).ListSubItems(2).Text
        frmEditProtocol.cboProtocol.Text = lvObject.ListItems(lvObject.SelectedItem.Index).ListSubItems(1).Text
        frmEditProtocol.txtVehiclID.Text = lvObject.ListItems(lvObject.SelectedItem.Index).ListSubItems(3).Text
        
        frmEditProtocol.Show
    End If
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
    Select Case KeyAscii
        Case vbKeyReturn
            lvObject_DblClick
    End Select
End Sub

Private Sub lvObject_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    Dim szselectKey As String
    szselectKey = tvBaseItem.SelectedItem.Key
    Select Case szselectKey
    Case "KLuggageType"        '行包类型
        DoLuggageType EOS_Add
    Case "KProtocol"
        DoProtocol EOS_Add
    Case "KVehicle"
        DoKVehicle EOS_Add
    Case "KFormula"
        DoKFormula EOS_Add
    Case "KPriceFormula"
    '行包票价计算公式"
        DoKPriceFormula EOS_Add
    End Select
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'处理行包类型
Private Sub DoLuggageType(pnOper As EOperStatus)


    Select Case pnOper
        Case EOS_Add
            frmLugKinds.Status = ST_AddObj
            frmLugKinds.Show vbModal
        Case EOS_Modify
            frmLugKinds.Status = ST_EditObj
            frmLugKinds.mszLugID = lvObject.SelectedItem.Text
            frmLugKinds.Show vbModal
        Case EOS_Delete
            Dim vbYesOrNo As Integer
            vbYesOrNo = MsgBox("是否真的删除" & szTmp & "[" & lvObject.SelectedItem & "]", vbQuestion + vbYesNo + vbDefaultButton2, "删除基本信息")
            If vbYesOrNo = vbYes Then
                m_oLuggageKinds.Init m_oAUser
                m_oLuggageKinds.Identify lvObject.SelectedItem.Text
                m_oLuggageKinds.Delete
                lvObject.ListItems.Remove lvObject.SelectedItem.Index
            End If
    End Select

End Sub
'处理协议计算公式
Private Sub DoKFormula(pnOper As EOperStatus)

    Select Case pnOper
    Case EOS_Add
'        FrmSaveFormular.m_eStatus = ST_AddObj
        frmFormula.m_eStatus = ST_AddObj
        frmFormula.Show vbModal
    Case EOS_Modify
'        FrmSaveFormular.m_eStatus = ST_EditObj
        frmFormula.m_eStatus = ST_EditObj
        frmFormula.Show vbModal
    Case EOS_Delete
        Dim vbYesOrNo As Integer
        vbYesOrNo = MsgBox("是否真的删除" & szTmp & "[" & lvObject.SelectedItem & "]", vbQuestion + vbYesNo + vbDefaultButton2, "删除基本信息")
        If vbYesOrNo = vbYes Then
            m_oLugFormula.Init m_oAUser
            m_oLugFormula.Identify lvObject.SelectedItem.Text
            m_oLugFormula.Delete
            lvObject.ListItems.Remove lvObject.SelectedItem.Index
        End If
    End Select

End Sub

Private Sub DoKPriceFormula(pnOper As EOperStatus)
    
    Select Case pnOper
    Case EOS_Add
        frmLugFormula.m_bIsParent = True
        frmLugFormula.m_eStatus = eFormStatus.AddStatus
        frmLugFormula.Show vbModal
    Case EOS_Modify
        frmLugFormula.m_bIsParent = True
        frmLugFormula.m_eStatus = eFormStatus.ModifyStatus
        frmLugFormula.m_szFormulaId = lvObject.SelectedItem.Text
        frmLugFormula.Show vbModal
    Case EOS_Delete
        Dim vbYesOrNo As Integer
        vbYesOrNo = MsgBox("是否真的删除" & szTmp & "[" & lvObject.SelectedItem & "]", vbQuestion + vbYesNo + vbDefaultButton2, "删除基本信息")
        If vbYesOrNo = vbYes Then
            m_oluggageSvr.DelLugFormulaInfo lvObject.SelectedItem.Text
            
            lvObject.ListItems.Remove lvObject.SelectedItem.Index
        End If
    End Select

End Sub

'处理结算协议
Private Sub DoProtocol(pnOper As EOperStatus)


    Select Case pnOper
        Case EOS_Add
            frmAddProtocol.Status = ST_AddObj
            frmAddProtocol.Show vbModal
        Case EOS_Modify
             frmAddProtocol.mszProtocolID = lvObject.SelectedItem.Text
             frmAddProtocol.Status = ST_EditObj
             frmAddProtocol.Show vbModal
        Case EOS_Delete
           ' 删除结算协议
            Dim vbYesOrNo As Integer

            vbYesOrNo = MsgBox("是否真的删除" & szTmp & "[" & lvObject.SelectedItem & "]", vbQuestion + vbYesNo + vbDefaultButton2, "删除基本信息")
            If vbYesOrNo = vbYes Then
                m_oProtocol.Init m_oAUser
                m_oProtocol.Identify lvObject.SelectedItem.Text
                m_oProtocol.Delete
                lvObject.ListItems.Remove lvObject.SelectedItem.Index
            End If
        Case 3
            frmProtocol.Show
    End Select
End Sub
'处理车辆结算协议
Private Sub DoKVehicle(pnOper As EOperStatus)
    'pnOper=0       '新增
    'pnOper=1       '修改
    'pnOper=2       '删除
    Dim i As Integer
    Dim mCount As Integer
    Select Case pnOper
        Case EOS_Add
             
             For i = 1 To lvObject.ListItems.Count
                   lvObject.ListItems(i).Checked = False
             Next i
        Case EOS_Modify
             For i = 1 To lvObject.ListItems.Count
                   lvObject.ListItems(i).Checked = True
             Next i
             
        Case EOS_Delete
            Dim vbYesOrNo As Integer
            Dim szTemp() As String
      
            Dim nCount As Integer
            Dim k As Integer
            k = 1
            nCount = 0
            Dim j As Integer
            j = 1
            If lvObject.ListItems.Count = 0 Then Exit Sub
            
            For i = 1 To lvObject.ListItems.Count
                 If lvObject.ListItems(i).Checked = True Then
                     nCount = nCount + 1
                 End If
            Next i
       
            
            If nCount = 0 Then
                MsgBox "对不起你还没有选择要取消协议的车辆！", vbInformation, "警告"
                Exit Sub
            End If
            
            vbYesOrNo = MsgBox("是否真要取消这" & "[" & nCount & "]" & szTmp, vbQuestion + vbYesNo + vbDefaultButton2, "基本信息")
        If vbYesOrNo = vbYes Then
            ReDim szTemp(1 To nCount, 1 To 2)
            j = 1
            For i = 1 To lvObject.ListItems.Count
                If lvObject.ListItems(j).Checked = True Then
                    szTemp(j, 1) = Trim(lvObject.ListItems.Item(j).SubItems(3))
                    szTemp(j, 2) = GetLuggageTypeInt(Trim(lvObject.ListItems.Item(j).SubItems(2)))
                    j = j + 1
                End If
            Next i
            m_oProtocol.Init m_oAUser
            m_oProtocol.DelVehicleProtocol szTemp
        End If
        '刷新列表
        RefreshBusProtocol
    End Select
End Sub

Private Sub RefreshBusProtocol()
    On Error GoTo err
    Dim i As Integer
    For i = 1 To lvObject.ListItems.Count
        If lvObject.ListItems(i).Checked = True Then
            lvObject.ListItems.Item(i).SubItems(1) = "[]"
            lvObject.ListItems(i).Checked = False
        End If
    Next i
    
    Exit Sub
err:
    ShowErrorMsg
End Sub

'处理行包收费项
Private Sub DoKPriceItem(pnOper As EOperStatus)
     Dim oListItem As ListItem
    'pnOper=1       '修改

    Select Case pnOper
    Case EOS_Modify
    
    
        If Not m_bIsOneFormulaEachStation Then
    
    '
            frmPriceItemAdvance.m_szPriceItemId = lvObject.SelectedItem.Text
            frmPriceItemAdvance.m_szAcceptType = GetLuggageTypeInt(lvObject.SelectedItem.SubItems(2))
            frmPriceItemAdvance.Show vbModal
        Else
            frmPriceItem.m_szPriceItemId = lvObject.SelectedItem.Text
            frmPriceItem.m_szAcceptType = GetLuggageTypeInt(lvObject.SelectedItem.SubItems(2))
            frmPriceItem.Show vbModal
        End If
    End Select
End Sub

'更改当前选中的项目
Public Sub EditObject()
    On Error GoTo ErrHandle
    If lvObject.SelectedItem Is Nothing Then Exit Sub
    Select Case tvBaseItem.SelectedItem.Key
    Case "KLuggageType"        '行包类型
        DoLuggageType EOS_Modify
    Case "KProtocol"
        DoProtocol EOS_Modify
    Case "KVehicle"
        DoKVehicle EOS_Modify
    Case "KFormula"
        DoKFormula EOS_Modify
    Case "KPriceItem"
        DoKPriceItem EOS_Modify
    Case "KPriceFormula"
        DoKPriceFormula EOS_Modify
        
    End Select
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'列出基本信息
Private Sub FillItemLists()
    On Error GoTo ErrHandle
    Dim sLugItem() As TLugProtocol
    Dim sKVehicle() As TVehicleProtocol
    Dim nCount As Integer
    Dim aszItems() As String
    Dim rsItems As Recordset
    Dim oListItem As ListItem
    Dim szitem() As String
    '    Dim nlen As Integer
    Dim i As Integer, j As Integer
    lvObject.ListItems.clear
    ShowSBInfo "正在查询，请稍等..."
    SetBusy
    
    lblTitlePrompt.Caption = tvBaseItem.SelectedItem.Text
    Dim szselectKey As String
    szselectKey = tvBaseItem.SelectedItem.Key
    '得到基本信息
    Select Case szselectKey
    Case "KProtocol" '行包拆算协议
        FillItemHead
        m_oProtocol.Init m_oAUser
        nCount = ArrayLength(m_oProtocol.GetProtocol)
        If nCount > 0 Then
            ReDim sLugItem(1 To nCount)
            sLugItem = m_oProtocol.GetProtocol
            For i = 1 To nCount
                WriteProcessBar , i, nCount, "得到对象[" & sLugItem(i).ProtocolName & "]"
                If lvObject.ListItems.Count > 0 Then
                    For j = 1 To lvObject.ListItems.Count
                        If Trim(lvObject.ListItems(j).Text) = sLugItem(i).ProtocolID Then GoTo Nexthere
                    Next j
                End If
                Set oListItem = lvObject.ListItems.Add(, , sLugItem(i).ProtocolID)
                oListItem.SubItems(1) = sLugItem(i).ProtocolName
                oListItem.SubItems(2) = sLugItem(i).Annotation
                If sLugItem(i).IsDefault = True Then
                    oListItem.SubItems(3) = "是"
                    SetListViewLineColor lvObject, oListItem.Index, vbRed
                Else
                    oListItem.SubItems(3) = "否"
                    SetListViewLineColor lvObject, oListItem.Index, vbBlack
                End If
Nexthere:
            Next i
        Else
            SetNormal
            Exit Sub
        End If

    '计算公式
    Case "KFormula"
        FillItemHead
        m_oLugFormula.Init m_oAUser
        nCount = ArrayLength(m_oLugFormula.GetAllFormulas)
        '车辆指定协议
    Case "KVehicle"
        SetNormal
        FillItemHead
        lvObject.Refresh
        
        frmQueryVehicle.Show vbModal
        If frmQueryVehicle.IsCancel = True Then
            WriteProcessBar False
            ShowSBInfo ""
            
            Exit Sub
        End If

        Dim aszVehicles() As String
        Dim aszReturn() As String
        Dim szCompany As String
        Dim szOwner As String
        Dim szBusType As String
        Dim szLicense As String
        Dim szVehicle As String
        Dim szSplitCompany As String
        Dim szAcceptType As Integer
        
        With frmQueryVehicle
        
            szVehicle = Trim(.txtVehicle.Text)
            szCompany = IIf(Trim(.txtCompany.Text) = "", "", ResolveDisplay(.txtCompany.Text))
            szOwner = IIf(Trim(.txtBusOwner.Text) = "", "", ResolveDisplay(.txtBusOwner.Text))
            szLicense = IIf(Trim(.txtLicense.Text) = "", "", .txtLicense.Text)
            szSplitCompany = IIf(Trim(.txtSplitCompany.Text) = "", "", ResolveDisplay(.txtSplitCompany.Text))
            If .cboAcceptType.Text = "" Then
                szAcceptType = -1
            Else
    
                szAcceptType = IIf(Trim(.cboAcceptType.Text) = Trim(szAcceptTypeGeneral), 0, 1)
            End If
        End With
        m_oProtocol.Init m_oAUser
        nCount = ArrayLength(m_oProtocol.GetVehicleProtocol(szVehicle, szAcceptType, , szLicense, szCompany, szSplitCompany, szOwner))
        If nCount > 0 Then
            ReDim sKVehicle(1 To nCount)
            sKVehicle = m_oProtocol.GetVehicleProtocol(szVehicle, szAcceptType, , szLicense, szCompany, szSplitCompany, szOwner)
            For i = 1 To nCount
                WriteProcessBar , i, nCount, "得到对象[" & sKVehicle(i).ProtocolName & "]"
                SetBusy
                If lvObject.ListItems.Count > 0 Then
                    For j = 1 To lvObject.ListItems.Count
                        If Trim(lvObject.ListItems(j).Text) = sKVehicle(i).VehicleLicense And oListItem.SubItems(3) = sKVehicle(i).ProtocolID Then GoTo Nexthere1
                    Next j
                End If
                
                Set oListItem = lvObject.ListItems.Add(, , sKVehicle(i).VehicleLicense)
                oListItem.SmallIcon = "vehicle"
                
                oListItem.SubItems(1) = MakeDisplayString(sKVehicle(i).ProtocolID, sKVehicle(i).ProtocolName)
                oListItem.SubItems(2) = sKVehicle(i).AcceptType
                oListItem.SubItems(3) = sKVehicle(i).VehicleID
                
Nexthere1:
            Next i
        
        Else
            SetNormal
            Exit Sub
        End If

    '行包类型
    Case "KLuggageType"
        FillItemHead
        m_oLugParam.Init m_oAUser
        Set rsItems = m_oLugParam.GetLuggageKinds
        '票价项
    Case "KPriceItem"
        FillItemHead
        m_oLugParam.Init m_oAUser
        Set rsItems = m_oLugParam.GetPriceItemRS
        
    End Select

    
    Dim aszTmpItem(0 To 3) As String
    
    '    '填充列表
    Select Case szselectKey
    Case "KLuggageType"
        nCount = rsItems.RecordCount  '返回记录集
        If nCount > 0 Then
            For i = 1 To nCount
                WriteProcessBar , i, nCount, "得到对象[" & rsItems!kinds_name & "]"
                aszTmpItem(0) = rsItems!kinds_code
                aszTmpItem(1) = rsItems!kinds_name
                aszTmpItem(2) = rsItems!Annotation
                AddList aszTmpItem
                rsItems.MoveNext
            Next
        Else
            SetNormal
            Exit Sub
        End If
    Case "KPriceItem"
        nCount = rsItems.RecordCount  '返回记录集
        If nCount > 0 Then
            For i = 1 To nCount
                WriteProcessBar , i, nCount, "得到对象[" & rsItems!chinese_name & "]"
                aszTmpItem(0) = rsItems!charge_item
                aszTmpItem(1) = rsItems!chinese_name
                If rsItems!accept_type = 0 Then
                    aszTmpItem(2) = szAcceptTypeGeneral
                Else
                    aszTmpItem(2) = szAcceptTypeMan
                End If
                If rsItems!use_mark = 1 Then
                    aszTmpItem(3) = "否"
                Else
                    aszTmpItem(3) = "是"
                
                End If
                
                AddList aszTmpItem
                rsItems.MoveNext
            Next
        Else
            SetNormal
            Exit Sub
        End If

    Case "KFormula"
'                        nCount = ArrayLength(aszItems) '返回数组
        If nCount > 0 Then
            ReDim aszItems(1 To nCount, 1 To 3)
            aszItems = m_oLugFormula.GetAllFormulas
            For i = 1 To nCount
                WriteProcessBar , i, nCount, "得到对象[" & aszItems(i, 2) & "]"
                aszTmpItem(0) = aszItems(i, 1)
                aszTmpItem(1) = aszItems(i, 2)
                aszTmpItem(2) = aszItems(i, 3)
                AddList aszTmpItem
            Next
        Else
            SetNormal
            Exit Sub
        End If
    Case "KPriceFormula"
        FillItemHead
        Dim atFormulaInfo() As TLuggageFormulaInfo
        atFormulaInfo = m_oluggageSvr.GetLugFormulaInfo()
        nCount = ArrayLength(atFormulaInfo)
        If nCount > 0 Then
'            ReDim aszItems(1 To nCount, 1 To 2)
            
            For i = 1 To nCount
                WriteProcessBar , i, nCount, "得到对象[" & atFormulaInfo(i).FormulaName & "]"
                aszTmpItem(0) = atFormulaInfo(i).FormulaID
                aszTmpItem(1) = atFormulaInfo(i).FormulaName
'                aszTmpItem(2) = aszItems(i, 3)
                AddList aszTmpItem
            Next
        Else
            SetNormal
            Exit Sub
        End If
    End Select
    Set imgObject.Picture = bigImgLists.ListImages(LCase(Mid(szselectKey, 2))).Picture
    SetNormal
    lvObject.Refresh
    If lvObject.ListItems.Count > 0 Then lvObject.ListItems(1).Selected = True
    WriteProcessBar False
    ShowSBInfo "共" & nCount & "个对象", ESB_ResultCountInfo
    ShowSBInfo ""

  
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
'    szBusType = IIf(Trim(.txtVehicleType.Text) = "", "", ResolveDisplay(.txtVehicleType.Text))
    End With

    Dim oVehicle As New BaseInfo
    oVehicle.Init m_oAUser
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
    Dim szselectKey As String
    szselectKey = tvBaseItem.SelectedItem.Key
    Select Case szselectKey
         Case "KLuggageType"        '行包类型
            szTmp = "行包类型"
            DoLuggageType EOS_Delete
         Case "KProtocol"
            szTmp = "拆算协议"
            DoProtocol EOS_Delete
         Case "KVehicle"
            szTmp = "辆的协议"
            DoKVehicle EOS_Delete
         Case "KFormula"
            szTmp = "协议计算公式"
            DoKFormula EOS_Delete
         Case "KPriceItem"
            MsgBox "行包票价项不能删除！", vbInformation, "警告"
            Exit Sub
        Case "KPriceFormula"
            szTmp = "行包价格计算公式"
            DoKPriceFormula EOS_Delete
    End Select
    SetMenuEnabled
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'添加项目到listview
Public Sub AddList(paszItems As Variant, Optional pbEnsure As Boolean)
    '    'pbEnsure 是否滚到此行
    Dim oListItem As ListItem
    Dim szselectKey As String
    szselectKey = tvBaseItem.SelectedItem.Key
    Set oListItem = lvObject.ListItems.Add(, , Trim(paszItems(0)))
    oListItem.SubItems(1) = paszItems(1)
    oListItem.SubItems(2) = paszItems(2)
    oListItem.SubItems(3) = paszItems(3)
    If szselectKey = "KPriceItem" Or szselectKey = "KProtocol" Then
        If paszItems(3) = "是" Then
            SetListViewLineColor lvObject, oListItem.Index, vbRed
        Else
            SetListViewLineColor lvObject, oListItem.Index, vbBlack
        End If
    End If
    oListItem.Selected = True
    If pbEnsure Then oListItem.EnsureVisible
    SetMenuEnabled
    
    If szselectKey = "KProtocol" Then
        FillItemLists
    End If
End Sub
'更改项目到listview
Public Sub UpdateList(paszItems As Variant)
    Dim oListItem As ListItem
    Dim szselectKey As String
    szselectKey = tvBaseItem.SelectedItem.Key
    Set oListItem = lvObject.SelectedItem
    If oListItem Is Nothing Then Exit Sub
    oListItem.SubItems(1) = paszItems(1)
    oListItem.SubItems(2) = paszItems(2)
    oListItem.SubItems(3) = paszItems(3)
    If szselectKey = "KPriceItem" Or szselectKey = "KProtocol" Then
        If paszItems(3) = "是" Then
'            oListItem.SmallIcon = "vehiclestop"
            SetListViewLineColor lvObject, oListItem.Index, vbRed
        Else
'            oListItem.SmallIcon = "vehiclerun"
            SetListViewLineColor lvObject, oListItem.Index, vbBlack
        End If
        lvObject.Refresh
  End If
  
    If szselectKey = "KProtocol" Then
        FillItemLists
    End If

End Sub
'设置activebar是否有用
Private Sub SetMenuEnabled()
    Dim szselectKey As String
    Dim bEnabled As Boolean
    szselectKey = tvBaseItem.SelectedItem.Key
    
    If lvObject.ListItems.Count > 0 Then
        bEnabled = True
    Else
        bEnabled = False
    End If
    Select Case szselectKey
    Case "KLuggageType", "KProtocol", "KFormula", "KPriceItem", "KPriceFormula"      '行包类型
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Add").Caption = "新增"
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_BaseInfo").Caption = "属性"
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Del").Caption = "删除"
        pmnu_Add.Caption = "新增"
        pmnu_Del.Caption = "删除"
        pmnu_BaseInfo.Caption = "属性"
        If szselectKey = "KPriceItem" Then
            abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Add").Enabled = False
            abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Del").Enabled = False
            abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_BaseInfo").Enabled = True
            abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Edit").Enabled = False
            pmnu_Del.Enabled = False
            pmnu_Add.Enabled = False
            pmnu_BaseInfo.Enabled = True
            Exit Sub
        End If
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Add").Enabled = True
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_BaseInfo").Enabled = bEnabled
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Del").Enabled = bEnabled
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Edit").Enabled = False
        pmnu_Del.Enabled = bEnabled
        pmnu_Add.Enabled = True
        pmnu_BaseInfo.Enabled = bEnabled
    
    Case "KVehicle"
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Add").Caption = "取消全选"
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_BaseInfo").Caption = "全选"
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Del").Caption = "取消协议"
        pmnu_Add.Caption = "取消全选"
        pmnu_Del.Caption = "取消协议"
        pmnu_BaseInfo.Caption = "全选"
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Add").Enabled = bEnabled
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_BaseInfo").Enabled = bEnabled
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Del").Enabled = bEnabled
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Edit").Enabled = bEnabled
        pmnu_Del.Enabled = bEnabled
        pmnu_Add.Enabled = bEnabled
        pmnu_BaseInfo.Enabled = bEnabled
    End Select


End Sub
'改变LVOBJECT相应的列头。
Public Sub FillItemHead()
   On Error GoTo ErrorHandle
    Dim oListItem As ListItem
    Dim szselectKey As String
    szselectKey = tvBaseItem.SelectedItem.Key
   
    Select Case szselectKey
        Case "KProtocol"
            lvObject.Checkboxes = False
            lvObject.MultiSelect = False
            lvObject.ColumnHeaders(1).Text = "协议代号"
            lvObject.ColumnHeaders(2).Text = "折算协议"
            lvObject.ColumnHeaders(3).Text = "备注"
            lvObject.ColumnHeaders(4).Text = "默认协议"
            lvObject.ColumnHeaders(4).Width = 1440
        Case "KVehicle"
            lvObject.Checkboxes = True
            lvObject.MultiSelect = True  '车辆允许多行选择
            lvObject.ColumnHeaders(1).Text = "车牌"
            lvObject.ColumnHeaders(2).Text = "拆算协议"
            lvObject.ColumnHeaders(3).Text = "托运方式"
            lvObject.ColumnHeaders(4).Text = ""
            lvObject.ColumnHeaders(4).Width = 0
        
        
        Case "KFormula"
            lvObject.Checkboxes = False
            lvObject.MultiSelect = False
            lvObject.ColumnHeaders(1).Text = "公式代号"
            lvObject.ColumnHeaders(2).Text = "公式名称"
            lvObject.ColumnHeaders(3).Text = "公式内容"
            lvObject.ColumnHeaders(4).Text = ""
            lvObject.ColumnHeaders(4).Width = 0
        Case "KLuggageType"
            lvObject.Checkboxes = False
            lvObject.MultiSelect = False
            lvObject.ColumnHeaders(1).Text = "类型代号"
            lvObject.ColumnHeaders(2).Text = "行包类型"
            lvObject.ColumnHeaders(3).Text = "备注"
            lvObject.ColumnHeaders(4).Text = ""
            lvObject.ColumnHeaders(4).Width = 0
        Case "KPriceItem"
            lvObject.Checkboxes = False
            lvObject.MultiSelect = False
            lvObject.ColumnHeaders(1).Text = "票价项代号"
            lvObject.ColumnHeaders(2).Text = "票价项"
            lvObject.ColumnHeaders(3).Text = "托行方式"
            lvObject.ColumnHeaders(4).Text = "是否已启用"
            lvObject.ColumnHeaders(4).Width = 1440
        Case "KPriceFormula"
            lvObject.Checkboxes = False
            lvObject.MultiSelect = False
            
            lvObject.ColumnHeaders(1).Text = "公式代号"
            lvObject.ColumnHeaders(2).Text = "公式名称"
            lvObject.ColumnHeaders(3).Text = ""
            lvObject.ColumnHeaders(3).Width = 0
            lvObject.ColumnHeaders(4).Text = ""
            lvObject.ColumnHeaders(4).Width = 0
    End Select
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub


