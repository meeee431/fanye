VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BusType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**********************************************************
'* Source File Name:BusType.cls
'* Project Name:BusType.vbp
'* Engineer:
'* Data Generated:
'* Last Revision Date:
'* Brief Description:车次类型对象
'* Relational Document:
'**********************************************************

Option Explicit
'出错
Public Enum EErrBusType
    ERR_BusTypeNotDeleteBus = ERR_BusType + ERR_DeleteChildExist '因为车次类型信息有车次所以不能删除该车次类型
    ERR_BusTypeNotAddStatus = ERR_BusType + ERR_NotAddObj '车次类型信息不在新增状态
    ERR_BusTypeExist = ERR_BusType + ERR_AddDuplicate '车次类型信息已存在
    
    ERR_BusTypeNotAvailable = ERR_BusType + ERR_NotAvailable ' 33212 车次类型信息对象处于无状态
    ERR_BusTypeIDNotNull = ERR_BusType + 21 '车次类型ID不能为空
    ERR_BusTypeInfoNotExist = ERR_BusType + 22 '车次类型信息不存在
    ERR_FalseDelete = ERR_BusType + 23 ''0-固定车次 01--滚动车次 不能删除
End Enum

'权限
Public Enum ERightBusType
    RIGHT_BusTypeManagement = ERR_BusType + cnMidRightBegin + cnMidRightStep * 1 '54906校对用户是否有权限创建对象
End Enum
Private m_szBusTypeID As String '
Private m_szBusTypeName As String '
Private m_szAnnotation As String '
Private m_oActiveUser As ActiveUser
Private m_nObjectStatus As EObjectStatus
Private Sub AddObject()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim szTempSql As String
    Dim szSqlField As String
    Dim szSqlContext As String
    If Trim(m_szBusTypeID) = "" Or Trim(m_szBusTypeName) = "" Then ShowError ERR_BusTypeIDNotNull
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szTempSql = "INSERT bus_type_code ("
    szSqlField = szSqlField & "Bus_type,"
    szSqlField = szSqlField & "Bus_type_name,"
    szSqlField = szSqlField & "Annotation)"
    szSqlContext = " VALUES ('" & m_szBusTypeID & "',"
    szSqlContext = szSqlContext & " '" & m_szBusTypeName & "',"
    szSqlContext = szSqlContext & " '" & m_szAnnotation & "')"
    szTempSql = szTempSql & szSqlField & szSqlContext
    oDb.Execute (szTempSql)
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertAddObjectError ERR_BusType, oDb
End Sub
Private Sub UpdateToDB()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "UPDATE bus_type_code SET Bus_type_name='" & m_szBusTypeName & "'," _
    & "annotation='" & m_szAnnotation & "' WHERE bus_type='" & m_szBusTypeID & "'"
    oDb.Execute szSql
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertUpdateObjectError ERR_BusType, oDb
End Sub

Public Property Get ObjStatus() As EObjectStatus
    ObjStatus = m_nObjectStatus
End Property
'P2
Public Property Let ObjStatus(ByVal vData As EObjectStatus)
    m_nObjectStatus = vData
End Property
Public Function GetAllBusType(Optional ByVal szBusTypeID As String) As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    Dim szSql As String
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    
    If szBusTypeID = "" Then
    szSql = "Select * from bus_type_code "
    Else
    szSql = "SELECT *  FROM bus_type_code WHERE Bus_type='" & szBusTypeID & "' "
    End If
    
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    ReDim szaTemp(1 To nCount, 1 To 3)
    For i = 1 To nCount
        szaTemp(i, 1) = FormatDbValue(rsTemp!Bus_type)
        szaTemp(i, 2) = FormatDbValue(rsTemp!Bus_type_name)
        szaTemp(i, 3) = FormatDbValue(rsTemp!Annotation)
        rsTemp.MoveNext
    Next
    GetAllBusType = szaTemp
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function
Public Property Let Annotation(ByVal vData As String)

    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szAnnotation = vData
End Property
Public Property Get Annotation() As String
    Annotation = m_szAnnotation
End Property



Public Property Let BusTypeName(ByVal vData As String)

    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szBusTypeName = vData
End Property


Public Property Get BusTypeName() As String
  BusTypeName = m_szBusTypeName
End Property
Public Property Let BusTypeID(ByVal vData As String)
  Dim oDb As New RTConnection
  Dim rsTemp As New Recordset
  Dim szSql As String
   m_szBusTypeID = vData
   AssertActiveUserValid m_oActiveUser, ERR_SeatType
    '查询是否是新增状态
    If m_nObjectStatus <> ST_AddObj Then ShowError ERR_BusTypeNotAddStatus
    '输入的主键是否为空
    If vData = "" Then ShowError ERR_BusTypeIDNotNull
    '查询该对象是否已有记录
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT * FROM bus_type_code WHERE bus_type='" & Trim(vData) & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 1 Then ShowError ERR_BusTypeExist
    m_szBusTypeID = RTrim(vData)
    Set rsTemp = Nothing
    Set oDb = Nothing
End Property
Public Property Get BusTypeID() As String
    BusTypeID = m_szBusTypeID
End Property
Public Sub Update()
   AssertHaveRight m_oActiveUser, RIGHT_BusTypeManagement
    AssertObjIsValid
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

End Sub
Public Sub Identify(ByVal szBusTypeID As String)
    AssertActiveUserValid m_oActiveUser, ERR_BusType
    m_szBusTypeID = RTrim(szBusTypeID)
    RefreshMemoryInfo
    m_nObjectStatus = ST_NormalObj
End Sub
Public Sub Delete()
    AssertHaveRight m_oActiveUser, RIGHT_BusTypeManagement
    AssertObjIsValid
    DeleteObject
    m_nObjectStatus = ST_NotAvailable
End Sub
Public Sub Init(Optional ByVal poAUser As Variant)
Set m_oActiveUser = poAUser
End Sub
Public Sub AddNew()
    AssertHaveRight m_oActiveUser, RIGHT_BusTypeManagement
    AssertActiveUserValid m_oActiveUser, ERR_BusType
    m_nObjectStatus = ST_AddObj
End Sub
Private Sub DeleteObject()
    Dim oDb As New RTConnection
    On Error GoTo Here
    '0-固定车次 1--滚动车次 不能删除
    If Trim(m_szBusTypeID) = "0" Or Trim(m_szBusTypeID) = "1" Then ShowError ERR_FalseDelete
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    oDb.Execute ("DELETE bus_type_code WHERE Bus_type='" & m_szBusTypeID & "'")
    Set oDb = Nothing
    Exit Sub
Here:
    AssertDeleteObjectError ERR_BusType, oDb
End Sub
Private Sub RefreshMemoryInfo()
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute("SELECT * FROM bus_type_code WHERE bus_type='" & m_szBusTypeID & "'")
    If rsTemp.RecordCount = 0 Then ShowError ERR_BusTypeInfoNotExist
    m_szBusTypeName = FormatDbValue(rsTemp!Bus_type_name)
    m_szAnnotation = FormatDbValue(rsTemp!Annotation)
    Set rsTemp = Nothing
    Set oDb = Nothing
End Sub
Private Sub AssertObjIsValid() '测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
    AssertActiveUserValid m_oActiveUser, ERR_BusType
    AssertStatusAvailable
End Sub
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then ShowError ERR_BusTypeNotAvailable
End Sub

