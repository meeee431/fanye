VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSelectItem 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择项目"
   ClientHeight    =   4005
   ClientLeft      =   3285
   ClientTop       =   2415
   ClientWidth     =   6630
   Icon            =   "frmSelectItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   3000
      Top             =   1470
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
            Picture         =   "frmSelectItem.frx":27A2
            Key             =   "View Large Icons"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectItem.frx":28B4
            Key             =   "View Small Icons"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectItem.frx":29C6
            Key             =   "View List"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectItem.frx":2AD8
            Key             =   "View Details"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectItem.frx":2BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectItem.frx":2CDB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4110
      Top             =   1650
   End
   Begin VB.TextBox txtMatchString 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      TabIndex        =   2
      Top             =   3240
      Width           =   4095
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   " 选择(&O)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   315
      Left            =   5340
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   " 取消(&C)"
      Height          =   315
      Left            =   5340
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   2805
      Left            =   90
      TabIndex        =   1
      Top             =   330
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   4948
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   360
      Left            =   4680
      TabIndex        =   5
      Top             =   0
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kbigicon"
            Description     =   "大图标"
            Object.ToolTipText     =   "大图标"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ksmallicon"
            Description     =   "小图标"
            Object.ToolTipText     =   "小图标"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "klist"
            Description     =   "列表"
            Object.ToolTipText     =   "列表"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kdetail"
            Description     =   "详细资料"
            Object.ToolTipText     =   "详细资料"
            ImageIndex      =   4
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kaddnew"
            Description     =   "新增项目"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   285
      Left            =   1530
      TabIndex        =   7
      Top             =   30
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   503
      _Version        =   393216
      Format          =   22872064
      CurrentDate     =   37478
   End
   Begin VB.Label lblSelected 
      BackStyle       =   0  'Transparent
      Caption         =   "选中的项目:"
      Height          =   180
      Left            =   90
      TabIndex        =   6
      Top             =   3330
      Width           =   990
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "学生列表:"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   810
   End
End
Attribute VB_Name = "frmSelectItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim maszListData() As String
Dim maszListColumns() As String
Dim mszTitle As String
Dim mszSelectedText As String
Dim maszSelectedItems As Variant
Dim mbMultiSelect As Boolean
Dim mnMatchIndex As Integer
Dim mnReturnIndex As Integer
Dim mszDialogCaption As String
Dim mnItemSmallIcon As Variant
Dim mnItemIcon As Variant
Dim moSmallIcons As ImageList
Dim moIcons As ImageList
Dim mszMultiColumn As String
Dim mnItemViewType As Integer
Dim mbSelected As Boolean


Public IsHidding As Boolean
Public m_bDateChanged As Boolean '是否改变了日期
Private m_bDateVisible As Boolean
Private m_dyDate As Date


Private Sub dtpDate_Change()
    m_bDateChanged = True
    m_dyDate = dtpDate.Value
    Me.Hide
End Sub

Private Sub Form_Initialize()
'设置默认属性
    
    
    mnItemViewType = lvwReport
    mnMatchIndex = 1
    mnReturnIndex = 1
End Sub

Private Sub Form_Load()
    
    
    
'初始化处理
    cmdSelect.Enabled = False
   
    If mszDialogCaption = "" Then
        mszDialogCaption = Caption
    End If
    If mbMultiSelect Then
        lvList.MultiSelect = True
    Else
        lvList.MultiSelect = False
    End If
    If moIcons Is Nothing And (Not moSmallIcons Is Nothing) Then
        Set moIcons = moSmallIcons
    End If
    If moSmallIcons Is Nothing And (Not moIcons Is Nothing) Then
        Set moSmallIcons = moIcons
    End If
    
    If moIcons Is Nothing Then
        mnItemIcon = 0
    Else
        Set lvList.Icons = moIcons
    End If
    If moSmallIcons Is Nothing Then
        mnItemSmallIcon = 0
    Else
        Set lvList.SmallIcons = moSmallIcons
    End If
    
    If mnMatchIndex <= 0 Then mnMatchIndex = 1
    If mnReturnIndex <= 0 Then mnReturnIndex = 1
    
    
    If mszTitle = "" Then mszTitle = "项目列表"
    
    Caption = mszDialogCaption
    lblTitle.Caption = mszTitle & ":"
    lvList.View = mnItemViewType
    Toolbar1.Buttons(mnItemViewType + 1).Value = tbrPressed
    
    Dim aszTmp() As String
    mszSelectedText = ""
    maszSelectedItems = aszTmp
    
    m_bDateChanged = False '设置初始值
    dtpDate.Visible = m_bDateVisible '设置可见性
    dtpDate.Value = m_dyDate
End Sub

Private Sub Form_Paint()
    IsHidding = False
End Sub

Private Sub Form_Resize()
    Const cnOffset = 300
    If Me.Width < 3015 Then Me.Width = 3015
    If Me.Height < 2745 Then Me.Height = 2745
'顶部定位
    lblTitle.Top = 120
    lblTitle.Left = 90
    
'中部定位
    lvList.Top = 360
    lvList.Left = 90
    lvList.Width = Me.ScaleWidth - 180
    lvList.Height = Me.ScaleHeight - 1605 + cnOffset

    Toolbar1.Top = 15
    Toolbar1.Left = lvList.Left + lvList.Width - Toolbar1.Width

'底部定位
    txtMatchString.Top = Me.ScaleHeight - 1170 + cnOffset
    txtMatchString.Left = 1140
    txtMatchString.Width = Me.ScaleWidth - 2670
    lblSelected.Left = 90
    lblSelected.Top = txtMatchString.Top + 60
    cmdSelect.Top = txtMatchString.Top
    cmdSelect.Left = txtMatchString.Left + txtMatchString.Width + 105
    cmdCancel.Top = cmdSelect.Top + cmdSelect.Height + 45
    cmdCancel.Left = cmdSelect.Left

End Sub

Private Sub lvList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Static m_nUpColumn As Integer
On Error Resume Next
lvList.SortKey = ColumnHeader.Index - 1
If m_nUpColumn = ColumnHeader.Index - 1 Then
    lvList.SortOrder = lvwDescending
    m_nUpColumn = ColumnHeader.Index
Else
    lvList.SortOrder = lvwAscending
    m_nUpColumn = ColumnHeader.Index - 1
End If
lvList.Sorted = True
End Sub
Private Sub lvList_DblClick()
    cmdSelect_Click
End Sub

Private Sub cmdCancel_Click()
    mszSelectedText = ""
'    m_bDateChanged = False
    Unload Me
End Sub
Private Sub cmdSelect_Click()
    If lvList.ColumnHeaders.Count = 0 Then
        MsgBox "未选择任何项目!", vbExclamation, "警告"
        Exit Sub
    End If
    
    Dim oListItem As ListItem
    Set oListItem = lvList.SelectedItem
    If oListItem Is Nothing Then
        MsgBox "未选择任何项目!", vbExclamation, "警告"
        Exit Sub
    End If
    
    '返回SelectedText
    If mnReturnIndex = 1 Then
        mszSelectedText = oListItem.Text
    Else
        mszSelectedText = oListItem.SubItems(mnReturnIndex - 1)
    End If
    
    '如果允许MultiSelect，则返回SelectedItems,否则返回空串
    Dim aszReturn() As String
    If mbMultiSelect Then
        Dim i As Long, j As Integer, nTmp As Integer
        Dim nCount As Long, nColcount As Long
        For i = 1 To lvList.ListItems.Count
            If lvList.ListItems(i).Selected Then
                nCount = nCount + 1
            End If
        Next i
        aszReturn = SplitEncodeStringArray(mszMultiColumn)
        nColcount = ArrayLength(aszReturn)
        For i = 1 To nColcount
            aszReturn(i) = UnEncodeString(aszReturn(i))
        Next i

    '如果没有MulitColumn,则返回ReturnIndex指定的元素(一维数组),否则返回multiColumn指定的元素(构建二维数组)
        If nColcount > 0 Then
            ReDim maszSelectedItems(1 To nCount, 1 To nColcount)
        Else
            ReDim maszSelectedItems(1 To nCount)
        End If
            
        nCount = 0
        For i = 1 To lvList.ListItems.Count
            If lvList.ListItems(i).Selected Then
                nCount = nCount + 1
                If nColcount > 0 Then       'MultiColumn返回二维数组
                    For j = 1 To nColcount
                        nTmp = Val(aszReturn(j))
                        If nTmp = 1 Then
                            maszSelectedItems(nCount, j) = lvList.ListItems(i).Text
                        Else
                            maszSelectedItems(nCount, j) = lvList.ListItems(i).SubItems(nTmp - 1)
                        End If
                    Next j
                Else
                    If mnReturnIndex = 1 Then
                        maszSelectedItems(nCount) = lvList.ListItems(i).Text
                    Else
                        maszSelectedItems(nCount) = lvList.ListItems(i).SubItems(mnReturnIndex - 1)
                    End If
                End If

            End If
        Next i
    Else        '返回只有一项元素的二维数组maszSelectedItems
        aszReturn = SplitEncodeStringArray(mszMultiColumn)
        nColcount = ArrayLength(aszReturn)
        If nColcount > 0 Then
            For i = 1 To nColcount
                aszReturn(i) = UnEncodeString(aszReturn(i))
            Next i
            ReDim maszSelectedItems(1 To 1, 1 To nColcount)
            
            For j = 1 To nColcount
                nTmp = Val(aszReturn(j))
                If nTmp = 1 Then
                    maszSelectedItems(1, j) = lvList.SelectedItem.Text
                Else
                    maszSelectedItems(1, j) = lvList.SelectedItem.SubItems(nTmp - 1)
                End If
            Next j
        End If
    End If
'    m_bDateChanged = False
    Unload Me
End Sub


Public Property Let ListData(paszListData As Variant)
'列出的数据串(字符串数组)
    maszListData = paszListData
End Property

Public Property Get ListColumns() As Variant
    ListColumns = maszListColumns
End Property

Public Property Let ListColumns(paszListColumns As Variant)
'列出的数据标题(字符串数组)
    maszListColumns = paszListColumns
End Property

Public Property Get Title() As String
    Title = mszTitle
End Property

Public Property Let Title(ByVal pszTitle As String)
    mszTitle = pszTitle
End Property

Public Property Get SelectedText() As String
'选择项
    SelectedText = mszSelectedText
End Property

Public Property Get SelectedItems() As Variant
'返回多项选择项
    SelectedItems = maszSelectedItems
End Property


Public Property Get MultiSelect() As Boolean
    MultiSelect = mbMultiSelect
End Property

Public Property Let MultiSelect(ByVal pbMulti As Boolean)
    mbMultiSelect = pbMulti
End Property


Private Sub lvList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mbSelected = True
    If mnMatchIndex = 1 Then
        txtMatchString.Text = Item.Text
    Else
        txtMatchString.Text = Item.SubItems(mnMatchIndex - 1)
    End If
    txtMatchString.SelStart = 0
    txtMatchString.SelLength = Len(txtMatchString.Text)
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    FillListView
    txtMatchString.SetFocus
End Sub
Private Sub FillListView()
    Dim i As Long, j As Integer
    Dim nCountHeads As Integer
    Dim alLen() As Long
    
    nCountHeads = ArrayLength(maszListColumns)
    If nCountHeads = 0 Then Exit Sub
    ReDim alLen(1 To nCountHeads)
    If mnReturnIndex > nCountHeads Then mnReturnIndex = 1
    If mnMatchIndex > nCountHeads Then mnMatchIndex = 1
    '添加列头
    lvList.ColumnHeaders.Clear
    For i = 1 To nCountHeads
        alLen(i) = LenA(maszListColumns(i))
        lvList.ColumnHeaders.Add , , maszListColumns(i)
    Next i
    
    Dim oListItem As ListItem
    '得到正确的需追加的列
    lvList.ListItems.Clear
    nCountHeads = ArrayLength(maszListData, 2)
    nCountHeads = IIf(nCountHeads > lvList.ColumnHeaders.Count, lvList.ColumnHeaders.Count, nCountHeads)
    
    Dim lTmpLen As Long
    For i = 1 To ArrayLength(maszListData)
        lTmpLen = LenA(maszListData(i, 1))
        If lTmpLen > alLen(1) Then alLen(1) = lTmpLen
        Set oListItem = lvList.ListItems.Add(, , maszListData(i, 1))
        For j = 1 To nCountHeads - 1
            lTmpLen = LenA(maszListData(i, j + 1))
            If lTmpLen > alLen(j + 1) Then alLen(j + 1) = lTmpLen
            
            oListItem.SubItems(j) = maszListData(i, j + 1)
                        
            If mnItemSmallIcon <> 0 Then oListItem.SmallIcon = mnItemSmallIcon
            If mnItemIcon <> 0 Then oListItem.Icon = mnItemIcon
        Next j
    Next i
    For i = 1 To ArrayLength(alLen)
        lvList.ColumnHeaders(i).Width = (alLen(i) + 2) * 90
    Next i
End Sub

Public Property Get MatchIndex() As Integer
    MatchIndex = mnMatchIndex
End Property

Public Property Let MatchIndex(ByVal pnMatchIndex As Integer)
    mnMatchIndex = pnMatchIndex
End Property

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "kbigicon"
            '应做:添加 '大图标' 按钮代码。
            lvList.View = lvwIcon
        Case "ksmallicon"
            '应做:添加 '小图标' 按钮代码。
            lvList.View = lvwSmallIcon
        Case "klist"
            '应做:添加 '列表' 按钮代码。
            lvList.View = lvwList
        Case "kdetail"
            '应做:添加 '详细资料' 按钮代码。
            lvList.View = lvwReport
        Case "kaddnew"
            IsHidding = True
            Me.Hide
    End Select
    
End Sub

Private Sub txtMatchString_Change()
    If txtMatchString.Text <> "" Then cmdSelect.Enabled = True
    If mbSelected Then
        mbSelected = False
        Exit Sub
    End If
    If txtMatchString.Text = "" Then
        cmdSelect.Enabled = False
        Exit Sub
    End If
    RealLocateLVW txtMatchString, lvList, mnMatchIndex
End Sub


Public Property Get ReturnIndex() As Integer
'返回值的列序号
    ReturnIndex = mnReturnIndex
End Property

Public Property Let ReturnIndex(ByVal pnReturnIndex As Integer)
    mnReturnIndex = pnReturnIndex
End Property

Public Property Get DialogCaption() As String
    DialogCaption = mszDialogCaption
End Property

Public Property Let DialogCaption(ByVal pszCaption As String)
    mszDialogCaption = pszCaption
End Property

Public Property Get ItemSmallIconIndex() As Variant
    ItemSmallIconIndex = mnItemSmallIcon
End Property
Public Property Set ItemIcons(ByVal poIcons As ImageList)
    Set moIcons = poIcons
End Property

Public Property Let ItemIconIndex(ByVal pnItemIcon As Variant)
    mnItemIcon = pnItemIcon
End Property
Public Property Get ItemIconIndex() As Variant
    ItemIconIndex = mnItemIcon
End Property

Public Property Let ItemSmallIconIndex(ByVal pnSmallIcon As Variant)
    mnItemSmallIcon = pnSmallIcon
End Property
Public Property Set ItemSmallIcons(ByVal poSmallIcons As ImageList)
    Set moSmallIcons = poSmallIcons
End Property

Public Property Get MultiColumn() As String
    MultiColumn = mszMultiColumn
End Property

Public Property Let MultiColumn(ByVal pszMultiColumn As String)
    '返回多列结果的匹配串,例[1][2][4]即返回第1,2,4列
    '如果MulitColumn无效或空,则返回ReturnIndex指定的元素(一维数组),否则返回multiColumn指定的元素(构建二维数组)
    mszMultiColumn = pszMultiColumn
End Property


Public Property Get ItemViewType() As Integer
    ItemViewType = mnItemViewType
End Property

Public Property Let ItemViewType(ByVal pnItemViewType As Integer)
    mnItemViewType = pnItemViewType
End Property


'在打开时添加一新行
Public Sub InsertItemLine(paszListData As Variant, Optional ByVal pbEnsure As Boolean = False)
    'pbEnsure 是否跳到该行
On Error Resume Next
    Dim i As Integer
    Dim oListItem As ListItem
    Dim nCountHeads As Integer
    nCountHeads = ArrayLength(maszListColumns)
    If nCountHeads > ArrayLength(paszListData) Then
        nCountHeads = ArrayLength(paszListData)
    End If
    Set oListItem = lvList.ListItems.Add(, , paszListData(1))
    For j = 1 To nCountHeads - 1
        oListItem.SubItems(j) = paszListData(j + 1)
        If mnItemSmallIcon > 0 Then oListItem.SmallIcon = mnItemSmallIcon
        If mnItemIcon > 0 Then oListItem.Icon = mnItemIcon
    Next j
    If pbEnsure Then
        oListItem.EnsureVisible
        oListItem.Selected = True
    End If
End Sub



Private Sub txtMatchString_GotFocus()
    txtMatchString.SelStart = 0
    txtMatchString.SelLength = Len(txtMatchString.Text)
End Sub

Private Sub txtMatchString_KeyDown(KeyCode As Integer, Shift As Integer)
    '使用上下键进行记录移动
    Dim lIndex As Long
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        If lvList.ListItems.Count = 0 Then Exit Sub
        If lvList.SelectedItem Is Nothing Then
            lIndex = 0
        Else
            lIndex = lvList.SelectedItem.Index
        End If
        lIndex = lIndex + IIf(KeyCode = vbKeyUp, -1, 1)
        If lIndex > lvList.ListItems.Count Then
            lIndex = lvList.ListItems.Count
        End If
        If lIndex < 1 Then
            lIndex = 1
        End If
        lvList.ListItems(lIndex).EnsureVisible
        lvList.ListItems(lIndex).Selected = True
        Call lvList_ItemClick(lvList.ListItems(lIndex))
        txtMatchString.SetFocus
'        Call txtMatchString_GotFocus
    End If

End Sub

Public Property Get DateVisibled() As Boolean
    DateVisibled = m_bDateVisible
End Property

Public Property Let DateVisibled(ByVal bNewValue As Boolean)
    m_bDateVisible = bNewValue
End Property


Public Property Get SelectDate() As Variant
    SelectDate = m_dyDate
End Property

Public Property Let SelectDate(ByVal vNewValue As Variant)
    m_dyDate = vNewValue
End Property
