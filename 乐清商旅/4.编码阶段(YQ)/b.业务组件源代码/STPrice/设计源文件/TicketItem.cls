VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TicketItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**********************************************************
'* Source File Name:TicketItem.cls
'* Project Name:StPrice.vbp
'* Engineer:
'* Data Generated:
'* Last Revision Date:
'* Brief Description:票价项对象
'* Relational Document:
'**********************************************************

Option Explicit

'----------------------
'各类特有的变量
Private m_szItemID As String
Private m_szItemName As String
Private m_bItemUseMark As Boolean
Private m_szPriceTableID As String
'---------------------


'所有的实体类都有的私有类变量
Private m_nObjectStatus As EObjectStatus
Private m_oActiveUser As ActiveUser

Public Enum EErrTicketItem
   ERR_TicketItemNotExist = ERR_TicketItem + 21 '对象不存在错误
End Enum

Public Enum ERightTicketItem
    'RIGHT_TicketItemIdentify = ERR_TicketItem + cnMidRightBegin + cnMidRightStep * 0 '实例化权限  *
    RIGHT_TicketItemUpdate = ERR_TicketItem + cnMidRightBegin + cnMidRightStep * 1 '修改票价项权限
End Enum


Public Property Get SelfUser() As ActiveUser
   Set SelfUser = m_oActiveUser
   m_oActiveUser.IAmLive
End Property

Public Property Set SelfUser(vNewValue As ActiveUser)
    If m_nObjectStatus = ST_AddObj Then
        ShowError ERR_AddObj
    ElseIf m_nObjectStatus = ST_EditObj Then
        ShowError ERR_EditObj
    ElseIf m_nObjectStatus = ERR_NormalObj Then
        ShowError ERR_NormalObj
    Else
        Set m_oActiveUser = vNewValue
    End If
    m_oActiveUser.IAmLive
End Property
Public Sub Init(poAUser As ActiveUser)
    Set SelfUser = poAUser
End Sub
Public Property Get ItemID() As String
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_QueryPrice
   ItemID = m_szItemID
    
End Property
Public Property Let ItemID(vNewValue As String)
    AssertObjIsValid
    If m_nObjectStatus <> ST_AddObj Then ShowError ERR_NotAddObj
    m_szItemID = vNewValue
    
End Property
Private Sub RefreshInfo() '按照对象现在的ID将数据库中相应的信息读至对象内
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim odb As New RTConnection
    szSql = "SELECT * FROM price_item_info WHERE price_item='" & m_szItemID & "'"
    odb.ConnectionString = GetConnectionStr(cszPriceMan)
    Set rsTemp = odb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_TicketItemNotExist
    While Not rsTemp.EOF
       m_szItemName = FormatDbValue(rsTemp!chinese_name)
       m_bItemUseMark = IIf(FormatDbValue(rsTemp!use_mark) = 0, False, True)
       rsTemp.MoveNext
    Wend
    Set rsTemp = Nothing
    Set odb = Nothing
End Sub

'PriceTableID 参数不起作用，可随便使用
Public Sub Identify(PriceTableID As String, ItemID As String)
    AssertActiveUserValid m_oActiveUser, ERR_TicketItem
    'AssertHaveRight m_oActiveUser, RIGHT_TicketItemIdentify
    m_szItemID = ItemID
    RefreshInfo
    m_nObjectStatus = ST_NormalObj
End Sub

Private Sub Class_Initialize()
    m_nObjectStatus = ST_NotAvailable
End Sub

Public Sub Update()
   AssertObjIsValid
   If m_nObjectStatus = ST_EditObj Then
      AssertHaveRight m_oActiveUser, RIGHT_TicketItemUpdate
      UpdateToDB
      m_nObjectStatus = ST_NormalObj
   End If
End Sub


Private Sub AssertObjIsValid() '测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
    AssertActiveUserValid m_oActiveUser, ERR_TicketItem
    If m_nObjectStatus = ST_NotAvailable Then ShowError ERR_TicketItemNotExist
End Sub


Public Property Get ItemName() As String
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_QueryPrice
    ItemName = m_szItemName
End Property

Public Property Let ItemName(ByVal szNewValue As String)
    AssertObjIsValid
    m_szItemName = szNewValue
    If m_nObjectStatus = ST_NormalObj Then
        m_nObjectStatus = ST_EditObj
    End If
End Property

Public Property Get ItemUseMark() As Boolean
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_QueryPrice
    ItemUseMark = m_bItemUseMark
End Property

Public Property Let ItemUseMark(ByVal bNewValue As Boolean)
    AssertObjIsValid
    m_bItemUseMark = bNewValue
    If m_nObjectStatus = ST_NormalObj Then
        m_nObjectStatus = ST_EditObj
    End If
End Property
Private Sub UpdateToDB()
On Error GoTo Here
    Dim szSql As String, odb As New RTConnection
    szSql = "UPDATE price_item_info SET " _
    & "chinese_name='" & m_szItemName & "'," _
    & "use_mark=" & IIf(m_bItemUseMark, 1, 0) & " " _
    & "WHERE price_item ='" & m_szItemID & "'"
    
    odb.ConnectionString = GetConnectionStr(cszSNPrice)
    odb.Execute szSql
    WriteOperateLog m_oActiveUser, RIGHT_TicketItemUpdate, "修改票价项" & m_szItemID
    
    Exit Sub
Here:
    AssertUpdateObjectError ERR_TicketItem, odb
End Sub
