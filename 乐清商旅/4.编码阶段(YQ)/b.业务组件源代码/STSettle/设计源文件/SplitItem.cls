VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SplitItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E80EA430196"
'设置拆算费用项表(只可修改)
Option Base 0
Option Explicit

'折算项代码
'##ModelId=3E80EA5D007D
Private m_szSplitItemID As String

'折算项名称
'##ModelId=3E80EA630138
Private m_szSplitItemName As String

'使用状态
'0-使用,1-未用
'##ModelId=3E80EA6702DE
Private m_ESplitStatus As ESplitItemUseStatus

'折算类型
'0-拆给对方公司,1-拆给站方,2-留给本公司
'##ModelId=3E80EA6C03B9
Private m_ESplitType As ESplitItemType

'是否允许修改
'0-不允许修改,1-允许修改
'如果允许修改,则在结算过程中,可以修改该结算项的值
'##ModelId=3E80EA7100DA
Private m_EAllowModify As ESplitItemAllowModify

'local variable to hold the serialized class ID that was created in  Class_Initi-
'alize
'##ModelId=3E8219C00271
Private mmlClassDebugID As Long
Private m_oActiveUser As ActiveUser
Private m_nObjectStatus As EObjectStatus '对象状态


'##ModelId=3E80EA770167
Public Sub Init(opAUser As ActiveUser)
    On Error GoTo InitErr

    Set m_oActiveUser = opAUser

    Exit Sub
InitErr:
    Call RaiseError(MyUnhandledError, "SplitItem:Init Method")
End Sub

'##ModelId=3E80EA7C02AF
Public Sub Identify(pszSplitItemID As String)
    On Error GoTo IdentifyErr
    m_szSplitItemID = pszSplitItemID
    RefreshObject
    m_nObjectStatus = ST_NormalObj
    Exit Sub
IdentifyErr:
    Call RaiseError(MyUnhandledError, "SplitItem:Identify Method")
End Sub
Private Sub RefreshObject()
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "SELECT * FROM split_item_info WHERE split_item_id =" & TransFieldValueToString(m_szSplitItemID)
    Set rsTemp = oDb.Execute(szSql)
'    if rstemp.RecordCount =0 then RaiseError
    m_szSplitItemID = FormatDbValue(rsTemp!split_item_id)
    m_szSplitItemName = FormatDbValue(rsTemp!split_item_name)
    m_ESplitStatus = FormatDbValue(rsTemp!split_status)
    m_ESplitType = FormatDbValue(rsTemp!split_type)
    m_EAllowModify = FormatDbValue(rsTemp!allow_modify)
End Sub
'##ModelId=3E80EA8100CB
Public Sub UpDate()
    On Error GoTo UpDateErr
    AssertHaveRight m_oActiveUser, RIGHT_SplitItem
    UpDateToDB

    Exit Sub
UpDateErr:
    Call RaiseError(MyUnhandledError, "SplitItem:UpDate Method")
End Sub

'##ModelId=3E80EA850271
Private Sub UpDateToDB()
    On Error GoTo UpDateDBErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "UPDATE split_item_info SET " _
        & " split_item_name = " & TransFieldValueToString(m_szSplitItemName) & "," _
        & " split_status =  " & TransFieldValueToString(m_ESplitStatus) & "," _
        & " split_type = " & TransFieldValueToString(m_ESplitType) & "," _
        & " allow_modify = " & TransFieldValueToString(m_EAllowModify) _
        & " WHERE split_item_id = " & TransFieldValueToString(m_szSplitItemID)
    oDb.Execute (szSql)
    Exit Sub
UpDateDBErr:
    Call RaiseError(MyUnhandledError, "SplitItem:UpDateDB Method")
End Sub

'##ModelId=3E8219C10290
Private Sub Class_Initialize()
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
'        mmlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'##ModelId=3E8219C1030D
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

#If DebugMode Then
    '##ModelId=3E8225D8033C
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

'##ModelId=3E82291601E4
Public Property Get mlClassDebugID() As Long
    mlClassDebugID = mmlClassDebugID
End Property

'##ModelId=3E8229160138
Public Property Get AllowModify() As ESplitItemAllowModify
    AllowModify = m_EAllowModify
End Property

Public Property Let AllowModify(ByVal Values As ESplitItemAllowModify)
    m_EAllowModify = Values
End Property

'##ModelId=3E822916009C
Public Property Get SplitType() As ESplitItemType
    SplitType = m_ESplitType
End Property

Public Property Let SplitType(ByVal Values As ESplitItemType)
    m_ESplitType = Values
End Property


'##ModelId=3E8229160000
Public Property Get SplitStatus() As ESplitItemUseStatus
    SplitStatus = m_ESplitStatus
End Property

Public Property Let SplitStatus(ByVal Values As ESplitItemUseStatus)
    m_ESplitStatus = Values
End Property
'##ModelId=3E822915034B
Public Property Get SplitItemName() As String
    SplitItemName = m_szSplitItemName
End Property
Public Property Let SplitItemName(ByVal Values As String)
    m_szSplitItemName = Values
End Property

'##ModelId=3E82291502AF
Public Property Get SplitItemID() As String
    SplitItemID = m_szSplitItemID
End Property
Public Property Let SplitItemID(ByVal Values As String)
    m_szSplitItemID = Values
End Property






