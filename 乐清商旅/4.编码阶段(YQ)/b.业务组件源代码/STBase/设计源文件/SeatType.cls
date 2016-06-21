VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeatType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'**********************************************************
'* Source File Name:SeatType.cls
'* Project Name:STBase.vbp
'* Engineer:
'* Data Generated:
'* Last Revision Date:
'* Brief Description:座位类型对象
'* Relational Document:
'**********************************************************
Option Explicit
'出错
Public Enum EErrSeatType
    ERR_SeatTypeNotAvailable = ERR_SeatType + ERR_NotAvailable '座位类型信息对象处于无状态
    ERR_SeatTypeNotDeleteBus = ERR_SeatType + ERR_DeleteChildExist '因为有车次所以不能删除该检票口
    ERR_SeatTypeNotAddStatus = ERR_SeatType + ERR_NotAddObj '检票口对象不在新增状态
    ERR_SeatTypeExist = ERR_SeatType + ERR_AddDuplicate '座位类型信息已存在
    ERR_SeatTypeIDNotNull = ERR_SeatType + 20 '座位类型ID不能为空
    ERR_SeatTypeInfoNotExist = ERR_SeatType + 21 '座位类型信息不存在
    ERR_SeatFalseDelet = ERR_SeatType + 22 '不能删除普通座位类型
    ERR_NotDeleForBus = ERR_SeatType + 23 '座位类型已按排车辆
End Enum

'权限
Public Enum ERightSeatType
    RIGHT_SeatTypeManagement = ERR_SeatType + cnMidRightBegin + cnMidRightStep * 1 '校对用户是否有权限创建对象
End Enum

Private m_oActiveUser As ActiveUser 'P1所有的类都有的私有类变量
Private m_szSeatTypeID As String 'p1座位类型ID
Private m_szSeatTypeName As String '
Private m_szAnnotation As String '


Private m_oldszSeatTypeID As String 'p1座位类型ID
Private m_oldszSeatTypeName As String '
Private m_oldszAnnotation As String '


Private m_nObjectStatus As EObjectStatus

Private Sub AddObject()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim szTempSql As String
    Dim szSqlField As String
    Dim szSqlContext As String
    
    AssertHaveRight m_oActiveUser, RIGHT_SeatTypeManagement
    If Trim(m_szSeatTypeID) = "" Or Trim(m_szSeatTypeName) = "" Then ShowError ERR_SeatTypeIDNotNull
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szTempSql = "INSERT Seat_type_code ("
    szSqlField = szSqlField & "seat_type_id,"
    szSqlField = szSqlField & "seat_type_name,"
    szSqlField = szSqlField & "Annotation)"
    szSqlContext = " VALUES ('" & m_szSeatTypeID & "',"
    szSqlContext = szSqlContext & " '" & m_szSeatTypeName & "',"
    szSqlContext = szSqlContext & " '" & m_szAnnotation & "')"
    szTempSql = szTempSql & szSqlField & szSqlContext
    oDb.Execute (szTempSql)
    
    WriteOperateLog m_oActiveUser, RIGHT_SeatTypeManagement, "新增座位类型[" & m_szSeatTypeID & "]" & m_szSeatTypeName
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertAddObjectError ERR_SeatType, oDb
End Sub
Private Sub UpdateToDB()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szlog As String
    AssertHaveRight m_oActiveUser, RIGHT_SeatTypeManagement
    
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "UPDATE Seat_type_code SET seat_type_name='" & m_szSeatTypeName & "'," _
    & "annotation='" & m_szAnnotation & "' WHERE seat_type_id='" & m_szSeatTypeID & "'"
    oDb.Execute szSql
    Set oDb = Nothing
    
    
    
    szlog = "座位ID[" & m_szSeatTypeID & "]"
    If m_szSeatTypeName <> m_oldszSeatTypeName Then
    
       szlog = szlog & "名称由[" & m_oldszSeatTypeName & "]改为[" & m_szSeatTypeName & "]"
    End If
    
    
    If m_oldszAnnotation <> m_szAnnotation Then
      szlog = szlog & "注释由[" & m_oldszAnnotation & "]改为[" & m_szAnnotation & "]"
    End If

   ' AssertHaveRight m_oActiveUser, RIGHT_SeatTypeManagement
    WriteOperateLog m_oActiveUser, RIGHT_SeatTypeManagement, szlog
    Exit Sub
ErrHandle:
    AssertUpdateObjectError ERR_SeatType, oDb
    
End Sub

'Public Property Get ObjStatus() As EObjectStatus
'    ObjStatus = m_nObjectStatus
'End Property
''P2
'Public Property Let ObjStatus(ByVal vData As EObjectStatus)
'    m_nObjectStatus = vData
'End Property




Public Sub Update()
    AssertHaveRight m_oActiveUser, RIGHT_SeatTypeManagement
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

Public Sub Identify(szSeatTypeID As String)
    AssertActiveUserValid m_oActiveUser, ERR_SeatType
    m_szSeatTypeID = RTrim(szSeatTypeID)
    RefreshMemoryInfo
    m_nObjectStatus = ST_NormalObj
End Sub

Public Sub Init(poAUser As ActiveUser)
 Set m_oActiveUser = poAUser
End Sub

Public Sub Delete()
    AssertHaveRight m_oActiveUser, RIGHT_SeatTypeManagement
    AssertObjIsValid
    DeleteObject
    m_nObjectStatus = ST_NotAvailable
End Sub

Public Sub AddNew()
    AssertHaveRight m_oActiveUser, RIGHT_SeatTypeManagement
    AssertActiveUserValid m_oActiveUser, ERR_SeatType
    m_nObjectStatus = ST_AddObj
End Sub
Public Property Let Annotation(ByVal vData As String)

    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szAnnotation = vData
End Property

Public Property Get Annotation() As String
    Annotation = m_szAnnotation
End Property
Public Property Let SeatTypeName(ByVal vData As String)

    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szSeatTypeName = vData
End Property
Public Property Get SeatTypeName() As String
    SeatTypeName = m_szSeatTypeName
End Property
Public Property Let SeatTypeID(ByVal vData As String)
    Dim oDb As New RTConnection
    Dim rsTemp As New Recordset
    Dim szSql As String
    m_szSeatTypeID = vData
    AssertActiveUserValid m_oActiveUser, ERR_SeatType
    '查询是否是新增状态
    If m_nObjectStatus <> ST_AddObj Then ShowError ERR_SeatTypeNotAddStatus
    '输入的主键是否为空
    If vData = "" Then ShowError ERR_SeatTypeIDNotNull
    '查询该对象是否已有记录
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT * FROM Seat_type_code WHERE seat_type_id='" & Trim(vData) & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 1 Then ShowError ERR_SeatTypeExist
    m_szSeatTypeID = RTrim(vData)
    Set rsTemp = Nothing
    Set oDb = Nothing
End Property
Public Property Get SeatTypeID() As String
    SeatTypeID = m_szSeatTypeID
End Property
Private Sub DeleteObject()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As New Recordset
    
    AssertHaveRight m_oActiveUser, RIGHT_SeatTypeManagement
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    If Trim(m_szSeatTypeID) = "01" Or Trim(m_szSeatTypeID) = "02" Or Trim(m_szSeatTypeID) = "03" Then ShowError ERR_FalseDelete
    szSql = "SELECT * from vehicle_seat_type_info where seat_type_id='" & m_szSeatTypeID & "' "
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then ShowError ERR_NotDeleForBus
    oDb.Execute ("DELETE Seat_type_code WHERE seat_type_id='" & m_szSeatTypeID & "'")
    Set oDb = Nothing
    
    WriteOperateLog m_oActiveUser, RIGHT_SeatTypeManagement, "座位[" & m_szSeatTypeID & "]被删除"
    Exit Sub
ErrHandle:
    AssertDeleteObjectError ERR_SeatType, oDb

End Sub
Private Sub RefreshMemoryInfo()
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute("SELECT * FROM Seat_type_code WHERE seat_type_id='" & m_szSeatTypeID & "'")
    If rsTemp.RecordCount = 0 Then ShowError ERR_SeatTypeInfoNotExist
    m_szSeatTypeName = FormatDbValue(rsTemp!seat_type_name)
    m_szAnnotation = FormatDbValue(rsTemp!Annotation)
    
    
    m_oldszSeatTypeID = m_szSeatTypeID
    m_oldszSeatTypeName = m_szSeatTypeName
    m_oldszAnnotation = m_szAnnotation
    
    
    Set rsTemp = Nothing
    Set oDb = Nothing
End Sub
Private Sub AssertObjIsValid() '测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
    AssertActiveUserValid m_oActiveUser, ERR_SeatType
    AssertStatusAvailable
End Sub
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then ShowError ERR_SeatTypeNotAvailable
End Sub

