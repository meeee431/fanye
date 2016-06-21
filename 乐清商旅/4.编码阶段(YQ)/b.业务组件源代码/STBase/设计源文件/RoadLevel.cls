VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RoadLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'**********************************************************
'* Source File Name:RoadLevel.cls
'* Project Name:STBase.vbp
'* Engineer:
'* Data Generated:
'* Last Revision Date:
'* Brief Description:公路等级对象
'* Relational Document:
'**********************************************************
Option Explicit

'出错
Public Enum EErrRoadLevel
    ERR_RoadLevelExistNotAddNew = ERR_RoadLevel + ERR_AddDuplicate '57201已存在该公路等级不能新增
    ERR_RoadLevelNotDelete = ERR_RoadLevel + ERR_DeleteChildExist '57207该公路等级已安排路段或费率，不能删除
    ERR_RoadLevelNotAvailable = ERR_RoadLevel + ERR_NotAvailable '57212公路等级对象无状态
    ERR_RoadLevelNotAddStatus = ERR_RoadLevel + ERR_NotAddObj '57216公路等级对象处于无状态
    ERR_RoadLevelIDNotNull = ERR_RoadLevel + 21 '57221公路等级代码不能为空
    ERR_RoadLevelNotExist = ERR_RoadLevel + 22 ' 57222公路不存在
End Enum

'权限
Public Enum ERightRoadLevel
    RIGHT_RoadLevelManagement = ERR_RoadLevel + cnMidRightBegin + cnMidRightStep * 1 '57306公路等级管理
End Enum

Private m_oActiveUser As ActiveUser 'P1所有的类都有的私有类变量
Private m_nObjectStatus As EObjectStatus 'P2所有的实体类都有的私有类变量
Private m_szRoadLevelCode As String 'P3公路等级代码
Private m_szRoadLeveName As String 'P4公路等级名称
Private m_szRoadLevelShortName As String 'P5公路等级简称

'F1获得某种公路等级的所有路段
Public Function GetAllSection() As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute("SELECT section_id,section_name FROM section_info tbo WHERE tbo.road_rank_code='" & m_szRoadLevelCode & "'")
    If nCount <> 0 Then
    ReDim szaTemp(1 To nCount, 1 To 2) As String
    For i = 1 To nCount
        szaTemp(i, 1) = FormatDbValue(rsTemp!section_id)
        szaTemp(i, 2) = FormatDbValue(rsTemp!section_name)
        rsTemp.MoveNext
    Next
    GetAllSection = szaTemp
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function
'P1
Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
End Property

'P1
Public Property Set SelfUser(vNewValue As ActiveUser)
    If m_nObjectStatus = ST_AddObj Then
        RaiseError ERR_AddObj
    ElseIf m_nObjectStatus = ST_EditObj Then
        RaiseError ERR_EditObj
    ElseIf m_nObjectStatus = ST_NormalObj Then
        RaiseError ERR_NormalObj
    Else
        Set m_oActiveUser = vNewValue
    End If
End Property
'P2
Public Property Get ObjStatus() As EObjectStatus
    ObjStatus = m_nObjectStatus
End Property
'P3
Public Property Let RoadLevelCode(ByVal vData As String)
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    AssertActiveUserValid m_oActiveUser, ERR_RoadLevel
    '查询是否是新增状态
    If m_nObjectStatus <> ST_AddObj Then RaiseError ERR_RoadLevelNotAddStatus
    '输入的主键是否为空
    If vData = "" Then RaiseError ERR_RoadLevelIDNotNull
    '查询该对象是否已有记录
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT * FROM road_level_code WHERE road_rank_code='" & RTrim(vData) & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 1 Then RaiseError ERR_RoadLevelExistNotAddNew
    m_szRoadLevelCode = RTrim(vData)
    Set rsTemp = Nothing
    Set oDb = Nothing
End Property
'P3
Public Property Get RoadLevelCode() As String
    RoadLevelCode = m_szRoadLevelCode
End Property
'P4
Public Property Let RoadLeveName(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szRoadLeveName = vData
End Property
'P4
Public Property Get RoadLeveName() As String
    RoadLeveName = m_szRoadLeveName
End Property
'P5
Public Property Let RoadLevelShortName(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szRoadLevelShortName = vData
End Property
'P5
Public Property Get RoadLevelShortName() As String
    RoadLevelShortName = m_szRoadLevelShortName
End Property

'S1
Public Sub Init(poAUser As ActiveUser)
    Set m_oActiveUser = poAUser
    'AssertHaveRight m_oActiveUser, RIGHT_RoadLevelManagement
End Sub

'S2类刚建立时处于不可用状态,此时类不能使用
Private Sub Class_Initialize()
    m_nObjectStatus = ST_NotAvailable
End Sub

'S3测试对象的状态是否有效（无效状态[ST_NotAvailable]和新增状态[ST_AddObj]则出错）
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then RaiseError ERR_RoadLevelNotAvailable
End Sub
'S4
Private Sub AssertObjIsValid() '测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
    AssertActiveUserValid m_oActiveUser, ERR_RoadLevel
    AssertStatusAvailable
End Sub
'S5
Public Sub Identify(RoadLevel As String)
    AssertActiveUserValid m_oActiveUser, ERR_RoadLevel
    m_szRoadLevelCode = RTrim(RoadLevel)
    RefreshMemoryInfo
    m_nObjectStatus = ST_NormalObj
End Sub

'S6调用新增方法后，对象处于新增状态。
Public Sub AddNew()
    AssertHaveRight m_oActiveUser, RIGHT_RoadLevelManagement
    AssertActiveUserValid m_oActiveUser, ERR_RoadLevel
    m_nObjectStatus = ST_AddObj
End Sub
'S7
Public Sub Delete() '调用删除方法后，对象处于不可用状态
    AssertHaveRight m_oActiveUser, RIGHT_RoadLevelManagement
    AssertObjIsValid
    DeleteObject
    m_nObjectStatus = ST_NotAvailable
End Sub
'S8
Public Sub Update()
    AssertHaveRight m_oActiveUser, RIGHT_RoadLevelManagement
    AssertObjIsValid
    Select Case m_nObjectStatus
           Case ST_AddObj
           AddObject
           Case ST_EditObj
           UpdateToDB
           Case ST_NormalObj
           Case Else
           RaiseError ERR_NotAvailable
    End Select
    m_nObjectStatus = ST_NormalObj
End Sub

'S9
Private Sub RefreshMemoryInfo() '按照对象现在的ID将数据库中相应的信息读至对象内
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute("SELECT * FROM road_level_code WHERE road_rank_code='" & m_szRoadLevelCode & "'")
    If rsTemp.RecordCount = 0 Then RaiseError ERR_RoadLevelNotExist
    m_szRoadLevelShortName = FormatDbValue(rsTemp!short_name)
    m_szRoadLeveName = FormatDbValue(rsTemp!road_rank_name)
    Set rsTemp = Nothing
    Set oDb = Nothing
End Sub

'S10
Private Sub AddObject()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim szTempSql As String
    Dim szSqlField As String
    Dim szSqlContext As String
    If Trim(m_szRoadLevelCode) = "" Then RaiseError ERR_RoadLevelIDNotNull
    
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szTempSql = "INSERT road_level_code ("
    szSqlField = szSqlField & "road_rank_code,"
    szSqlField = szSqlField & "road_rank_name,"
    szSqlField = szSqlField & "short_name)"
    szSqlContext = " VALUES ('" & m_szRoadLevelCode & "',"
    szSqlContext = szSqlContext & " '" & m_szRoadLeveName & "',"
    szSqlContext = szSqlContext & " '" & m_szRoadLevelShortName & "')"
    szTempSql = szTempSql & szSqlField & szSqlContext
    oDb.Execute (szTempSql)
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertAddObjectError ERR_RoadLevel, oDb
End Sub
'S11
Private Sub DeleteObject()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    oDb.Execute ("DELETE road_level_code WHERE road_rank_code='" & m_szRoadLevelCode & "'")
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertDeleteObjectError ERR_RoadLevel, oDb
End Sub

'S12
Private Sub UpdateToDB()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim szTempSql As String
    Dim szSqlField As String
    Dim szSqlWhere As String
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szTempSql = "UPDATE road_level_code SET "
    szSqlField = szSqlField & "road_rank_name='" & m_szRoadLeveName & "',"
    szSqlField = szSqlField & "short_name='" & m_szRoadLevelShortName & "'"
    szSqlWhere = " WHERE road_rank_code='" & m_szRoadLevelCode & "'"
    szTempSql = szTempSql & szSqlField & szSqlWhere
    oDb.Execute (szTempSql)
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertUpdateObjectError ERR_RoadLevel, oDb
End Sub

