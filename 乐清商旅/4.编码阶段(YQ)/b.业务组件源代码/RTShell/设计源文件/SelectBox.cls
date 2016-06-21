VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelectBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event AddNewItem(ByVal FormCaption As String)
Public Event DateChange(ByVal FormCaption As String)

'Private m_bDateVisible As Boolean


Public Sub ShowBox()
    frmSelectItem.Show vbModal
    If frmSelectItem.IsHidding Then
        RaiseEvent AddNewItem(DialogCaption)
        ShowBox
    End If
    If frmSelectItem.m_bDateChanged Then
        '日期修改了
        RaiseEvent DateChange(DialogCaption)
        Unload frmSelectItem
        ShowBox
    End If
End Sub

Public Property Get ItemViewType() As Integer
    ItemViewType = frmSelectItem.ItemViewType
End Property

Public Property Let ItemViewType(ByVal pnItemViewType As Integer)
    frmSelectItem.ItemViewType = pnItemViewType
End Property
Public Property Get MultiColumn() As String
    MultiColumn = frmSelectItem.MultiColumn
End Property

Public Property Let MultiColumn(ByVal pszMultiColumn As String)
    '返回多列结果的匹配串,例[1][2][4]即返回第1,2,4列
    '如果MulitColumn无效或空,则返回ReturnIndex指定的元素(一维数组),否则返回multiColumn指定的元素(构建二维数组)
    frmSelectItem.MultiColumn = pszMultiColumn
End Property
Public Property Get ReturnIndex() As Integer
'返回值的列序号
    ReturnIndex = frmSelectItem.ReturnIndex
End Property

Public Property Let ReturnIndex(ByVal pnReturnIndex As Integer)
    frmSelectItem.ReturnIndex = pnReturnIndex
End Property

Public Property Get DialogCaption() As String
    DialogCaption = frmSelectItem.DialogCaption
End Property

Public Property Let DialogCaption(ByVal pszCaption As String)
    frmSelectItem.DialogCaption = pszCaption
End Property
Public Property Get MatchIndex() As Integer
    MatchIndex = frmSelectItem.MatchIndex
End Property

Public Property Let MatchIndex(ByVal pnMatchIndex As Integer)
    frmSelectItem.MatchIndex = pnMatchIndex
End Property
Public Property Let ListData(paszListData As Variant)
'列出的数据串(字符串数组)
    frmSelectItem.ListData = paszListData
End Property

Public Property Get ListColumns() As Variant
    ListColumns = frmSelectItem.ListColumns
End Property

Public Property Let ListColumns(paszListColumns As Variant)
'列出的数据标题(字符串数组)
    frmSelectItem.ListColumns = paszListColumns
End Property

Public Property Get Title() As String
    Title = frmSelectItem.Title
End Property

Public Property Let Title(ByVal pszTitle As String)
    frmSelectItem.Title = pszTitle
End Property

Public Property Get SelectedText() As String
'选择项
    SelectedText = frmSelectItem.SelectedText
End Property

Public Property Get SelectedItems() As Variant
'返回多项选择项
    SelectedItems = frmSelectItem.SelectedItems
End Property


Public Property Get MultiSelect() As Boolean
    MultiSelect = frmSelectItem.MultiSelect
End Property

Public Property Let MultiSelect(ByVal pbMulti As Boolean)
    frmSelectItem.MultiSelect = pbMulti
End Property
'在打开时添加一新行
Public Sub InsertItemLine(paszListData As Variant, Optional ByVal pbEnsure As Boolean = False)
    'pbEnsure 是否跳到该行
    frmSelectItem.InsertItemLine paszListData, pbEnsure
    
End Sub

Public Property Get Width() As Long
    Width = frmSelectItem.Width
End Property

Public Property Let Width(ByVal vNewValue As Long)
    frmSelectItem.Width = vNewValue
End Property
Public Property Get Height() As Long
    Height = frmSelectItem.Height
End Property

Public Property Let Height(ByVal vNewValue As Long)
    frmSelectItem.Height = vNewValue
End Property

Public Property Set ItemIcons(poIcons As Object)
    Set frmSelectItem.ItemIcons = poIcons
End Property
Public Property Let ItemIconIndex(ByVal pnItemIconIndex As Variant)
    frmSelectItem.ItemIconIndex = pnItemIconIndex
End Property
Public Property Set ItemSmallIcons(poSmallIcons As Object)
    Set frmSelectItem.ItemSmallIcons = poSmallIcons
End Property

Public Property Let ItemSmallIconIndex(ByVal pnSmallIconIndex As Variant)
    frmSelectItem.ItemSmallIconIndex = pnSmallIconIndex
End Property

Public Property Get EnableAddTool() As Boolean
    EnableAddTool = frmSelectItem.Toolbar1.Buttons("kaddnew").Enabled
End Property

Public Property Let EnableAddTool(ByVal vNewValue As Boolean)
    frmSelectItem.Toolbar1.Buttons("kaddnew").Enabled = vNewValue
End Property

Public Property Get AllowDateChange() As Boolean
    AllowDateChange = frmSelectItem.DateVisibled
End Property

Public Property Let AllowDateChange(ByVal bNewValue As Boolean)
    frmSelectItem.DateVisibled = bNewValue
End Property

Public Property Get SelectDate() As Date
    SelectDate = frmSelectItem.SelectDate
End Property

Public Property Let SelectDate(ByVal dyNewValue As Date)
    frmSelectItem.SelectDate = dyNewValue
End Property


