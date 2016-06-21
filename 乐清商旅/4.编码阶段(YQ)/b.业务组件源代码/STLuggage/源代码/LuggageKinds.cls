VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LuggageKinds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E039B550223"
'行包类型
Option Base 0

Option Explicit

'行包种类代码
'##ModelId=3E040DA60026
Private mszKindsCode As String

'行包种类名称
'##ModelId=3E040DC00236
Private mszKindsName As String
Private mszSheetID As String
'备注
'##ModelId=3E040DC00358
Private mszAnnotation As String
Private m_oActiveUser As ActiveUser
Private m_nObjectStatus As EObjectStatus

'本模块私有定义部分
'==========================================================================
'错误号
Private Enum EErrLuggageKinds
    ERR_LunggageKindsNotInfo = ERR_LuggageKinds + 21  '没有行包种类信息 '30821开始
    ERR_LunggageKindsExist = ERR_LuggageKinds + 22   '行包种类已经存在
End Enum

'权限号
Private Enum ERightLuggageKinds
    RIGHT_LuggageKinds = ERR_LuggageKinds + cnMidRightBegin               '30900开始
    RIGHT_LuggageKindsEdit = RIGHT_LuggageKinds + cnMidRightBegin + cnMidRightStep * 1  '更改行包种类信息
'    RIGHT_LuggageKindsADD = RIGHT_LuggageKinds + cnMidRightBegin + cnMidRightStep * 2
End Enum


'初始化活动对象
'##ModelId=3E040D9C00E0
Public Sub Init(poAUser As ActiveUser)
    On Error GoTo InitErr
    Set m_oActiveUser = poAUser
    'your code goes here...

    Exit Sub
InitErr:
    Call RaiseError(err.Number, "LuggageKinds:Init Method", err.Description)
End Sub

'指定受理单
'##ModelId=3E040D9C00FE
Public Sub Identify(SheetID As String)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageKinds
    On Error GoTo IdentifyErr
    mszKindsCode = Trim(SheetID)
    RefreshMemoryInfo
    m_nObjectStatus = ST_NormalObj
    Exit Sub
IdentifyErr:
    Call RaiseError(err.Number, "LuggageKinds:Identify Method", err.Description)
End Sub

'获得行包种类的属性
Private Sub RefreshMemoryInfo()
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT * FROM luggage_kinds_code WHERE  kinds_code ='" & mszKindsCode & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_LunggageKindsNotInfo
    mszKindsCode = FormatDbValue(rsTemp!Kinds_Code)
    mszKindsName = FormatDbValue(rsTemp!Kinds_Name)
    mszAnnotation = FormatDbValue(rsTemp!Annotation)
    Set oDb = Nothing
End Sub

'新增受理单
'##ModelId=3E040D9C011C
Public Sub AddNew()
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageKinds
    On Error GoTo AddNewErr
    m_nObjectStatus = ST_AddObj
    mszAnnotation = ""
    mszKindsCode = ""
    mszKindsName = ""
    Exit Sub
AddNewErr:
    Call RaiseError(err.Number, "LuggageKinds:AddNew Method", err.Description)
End Sub

'更新受理单
'##ModelId=3E040D9C013A
Public Sub Update()
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageKinds
'    AssertHaveRight m_oActiveUser, RIGHT_LuggageKindsEdit

    On Error GoTo UpdateErr
    Select Case m_nObjectStatus
        Case ST_AddObj
            AddObject
        Case ST_EditObj
            UpdateToDB
        Case ST_NormalObj
        Case Else
            ShowError ERR_NotAvailable
    End Select
    m_nObjectStatus = ST_NormalObj

    Exit Sub
UpdateErr:
    Call RaiseError(err.Number, "LuggageKinds:Update Method", err.Description)
End Sub
'新增行包种类
Private Sub AddObject()
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    On Error GoTo Here
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT * FROM luggage_kinds_code WHERE kinds_code ='" & mszKindsCode & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 1 Then ShowError ERR_LunggageKindsExist
    szSql = "INSERT luggage_kinds_code(kinds_code,kinds_name,annotation)" _
        & " VALUES( '" _
        & mszKindsCode & "','" _
        & mszKindsName & "','" _
        & mszAnnotation & "')"
    oDb.Execute szSql
    Set oDb = Nothing
    Exit Sub
Here:
    RaiseError ERR_LunggageKindsExist
End Sub

'修改行包种类
Private Sub UpdateToDB()
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT * FROM luggage_kinds_code "
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_LunggageKindsNotInfo
    szSql = "UPDATE luggage_kinds_code " _
            & " SET kinds_name ='" & mszKindsName & "', " _
            & " annotation ='" & mszAnnotation & "' " _
            & " WHERE kinds_code='" & mszKindsCode & "'"
    oDb.Execute szSql
    Set oDb = Nothing
    Exit Sub
End Sub

'删除行包种类
Private Sub DeleteObject()
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    On Error GoTo Here
    oDb.ConnectionString = GetConnectionStr(cszLuggage)

    
    szSql = "DELETE luggage_kinds_code " _
            & " WHERE  kinds_code ='" & mszKindsCode & "'"
    oDb.Execute szSql
    Set oDb = Nothing
    Exit Sub
Here:
End Sub
'删除受理单
'##ModelId=3E040D9C0158
Public Sub Delete()
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageKinds
'    AssertHaveRight m_oActiveUser, RIGHT_LuggageKindsEdit

    On Error GoTo DeleteErr


    DeleteObject
    m_nObjectStatus = ST_NotAvailable
    

    Exit Sub
DeleteErr:
    Call RaiseError(err.Number, "LuggageKinds:Delete Method", err.Description)
End Sub

'##ModelId=3E04A7A501D2
Public Property Get Annotation() As String
    Let Annotation = mszAnnotation
End Property

'##ModelId=3E04A7A50087
Public Property Let Annotation(ByVal Value As String)
    Let mszAnnotation = Value
    If m_nObjectStatus = ST_NormalObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E04A7A403BB
Public Property Get KindsName() As String
   Let KindsName = mszKindsName
End Property

'##ModelId=3E04A7A4027B
Public Property Let KindsName(ByVal Value As String)
    Let mszKindsName = Value
    If m_nObjectStatus = ST_NormalObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E04A7A401C6
Public Property Get KindsCode() As String
   Let KindsCode = mszKindsCode
End Property


'##ModelId=3E04A7A40090
Public Property Let KindsCode(ByVal Value As String)

    Let mszKindsCode = Value
    If m_nObjectStatus = ST_NormalObj Then m_nObjectStatus = ST_EditObj
End Property
