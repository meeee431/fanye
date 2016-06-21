VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CheckGate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'**********************************************************
'* Source File Name:CheckGate.cls
'* Project Name:STBase.vbp
'* Engineer:
'* Data Generated:
'* Last Revision Date:
'* Brief Description:检票口对象
'* Relational Document:
'**********************************************************
Option Explicit

'出错
Public Enum EErrCheckGate
    ERR_CheckGateExistNotAdd = ERR_CheckGate + ERR_AddDuplicate ' 57601该检票口已存在
    ERR_CheckGateNotDeleteBus = ERR_CheckGate + ERR_DeleteChildExist '57607因为有车次所以不能删除该检票口
    ERR_CheckGateNotAddStatus = ERR_CheckGate + ERR_NotAddObj '57616检票口对象不在新增状态
    ERR_CheckGateNotStatus = ERR_CheckGate + ERR_NotAvailable '57612检票口处于无状态
    ERR_CheckGateNotExist = ERR_CheckGate + 21 '57621检票口不存在
    ERR_CheckGateIDNotNull = ERR_CheckGate + 22 '57622检票口代码不能为空
    ERR_SellStationIDNotExist = ERR_Company + 23 '52823售票站点不存在
End Enum

'权限
Public Enum ERightCheckGate
    RIGHT_CheckGateManagement = ERR_CheckGate + cnMidRightBegin + cnMidRightStep * 1 '57706检票口管理
End Enum

Private m_nObjectStatus As EObjectStatus 'P2所有的实体类都有的私有类变量
Private m_oActiveUser As ActiveUser 'P1所有的类都有的私有类变量
Private m_szCheckGateCode As String 'P3检票口代码
Private m_szCheckGateName As String 'P4检票口名称
Private m_szAnnotation As String 'P5注释
Private m_SellStationID As String 'p6售票口代码
Private m_SellStationName As String 'p7售票口名称
'**************************************************
'Member Code:F1
'Brief Description:获得所以得车次
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/14
'**************************************************
Public Function GetAllBus(PlanID As String) As String()
'**************************************************
'PlanID (计划代码)
'**************************************************
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute("SELECT tbr.bus_id FROM bus_info tbr WHERE check_gate_id='" & m_szCheckGateCode & "'")
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    ReDim szaTemp(1 To nCount) As String
    For i = 1 To nCount
        szaTemp(i) = FormatDbValue(rsTemp!Bus_ID)
        rsTemp.MoveNext
    Next
    GetAllBus = szaTemp
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

'**************************************************
'Member Code:F2
'Brief Description:获得所有的检票员
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/25
'**************************************************
Public Function GetAllChecker() As String()
'**************************************************
'PlanID (计划代码)
'**************************************************
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute("SELECT user_id FROM Checker_gate_info WHERE check_gate_id='" & m_szCheckGateCode & "'")
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    ReDim szaTemp(1 To nCount) As String
    For i = 1 To nCount
        szaTemp(i) = FormatDbValue(rsTemp!user_id)
        rsTemp.MoveNext
    Next
    GetAllChecker = szaTemp
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

'**************************************************
'Member Code:P1
'Brief Description:活动用户
'**************************************************
Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
End Property

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
'**************************************************
'Member Code:P2
'Brief Description:对象状态
'**************************************************
Public Property Get ObjStatus() As EObjectStatus
    ObjStatus = m_nObjectStatus
End Property
'**************************************************
'Member Code:P3
'Brief Description:检票口代码
'**************************************************
Public Property Let CheckGateCode(ByVal vData As String)
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    AssertActiveUserValid m_oActiveUser, ERR_CheckGate
    '查询是否是新增状态
    If m_nObjectStatus <> ST_AddObj Then RaiseError ERR_CheckGateNotAddStatus
    '输入的主键是否为空
    If vData = "" Then RaiseError ERR_CheckGateIDNotNull
    '查询该对象是否已有记录
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT * FROM checkgate_info WHERE check_gate_id='" & RTrim(vData) & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 1 Then RaiseError ERR_CheckGateExistNotAdd
    m_szCheckGateCode = RTrim(vData)
    Set rsTemp = Nothing
    Set oDb = Nothing
End Property
Public Property Get CheckGateCode() As String
    CheckGateCode = m_szCheckGateCode
End Property
'**************************************************
'Member Code:P4
'Brief Description:检票口名称
'**************************************************
Public Property Let CheckGateName(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szCheckGateName = vData
End Property
Public Property Get CheckGateName() As String
    CheckGateName = m_szCheckGateName
End Property
'**************************************************
'Member Code:P5
'Brief Description:注释
'**************************************************
Public Property Let Annotation(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szAnnotation = vData
End Property

Public Property Get Annotation() As String
    Annotation = m_szAnnotation
End Property

'**************************************************
'Member Code:S1
'Brief Description:初始化对象
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/14
'**************************************************
Public Sub Init(poAUser As ActiveUser)
    Set m_oActiveUser = poAUser
    'AssertHaveRight m_oActiveUser, RIGHT_CheckGateManagement
End Sub

'**************************************************
'Member Code:S2
'Brief Description:类初始化
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub Class_Initialize()
    m_nObjectStatus = ST_NotAvailable
End Sub

'**************************************************
'Member Code:S3
'Brief Description:测试对象是否有效
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then RaiseError ERR_CheckGateNotStatus
End Sub
'**************************************************
'Member Code:S4
'Brief Description:测试对象是否有效
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub AssertObjIsValid()
    AssertActiveUserValid m_oActiveUser, ERR_CheckGate
    AssertStatusAvailable
End Sub
'**************************************************
'Member Code:S5
'Brief Description:标示检票口对象
'Engineer:wjb
'Date Generated:1999/10/11
'Last Revision Date:2002/11/9
'**************************************************
Public Sub Identify(CheckGateCode As String)
    AssertActiveUserValid m_oActiveUser, ERR_CheckGate
    m_szCheckGateCode = RTrim(CheckGateCode)
    RefreshMemoryInfo
    m_nObjectStatus = ST_NormalObj
End Sub

'**************************************************
'Member Code:S6
'Brief Description:新增检票口
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Public Sub AddNew()
    AssertHaveRight m_oActiveUser, RIGHT_CheckGateManagement
    AssertActiveUserValid m_oActiveUser, ERR_CheckGate
    m_nObjectStatus = ST_AddObj
End Sub
'**************************************************
'Member Code:S7
'Brief Description:删除检票口
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Public Sub Delete()
    AssertHaveRight m_oActiveUser, RIGHT_CheckGateManagement
    AssertObjIsValid
    DeleteObject
    m_nObjectStatus = ST_NotAvailable
End Sub
'**************************************************
'Member Code:S8
'Brief Description:修改检票口对象
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Public Sub Update()
    AssertHaveRight m_oActiveUser, RIGHT_CheckGateManagement
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

'**************************************************
'Member Code:S9
'Brief Description:从数据库中将地区属性读入内存
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub RefreshMemoryInfo()
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute("SELECT c.*,s.sell_station_id,s.sell_station_name " _
        & " FROM checkgate_info c,sell_station_info s" _
        & " WHERE c.sell_station_id=s.sell_station_id AND c.check_gate_id='" & m_szCheckGateCode & "'")
      
    If rsTemp.RecordCount = 0 Then RaiseError ERR_CheckGateNotExist
    m_szCheckGateName = FormatDbValue(rsTemp!check_gate_name)
    m_szAnnotation = FormatDbValue(rsTemp!Annotation)
    m_SellStationName = FormatDbValue(rsTemp!sell_station_name)
    m_SellStationID = FormatDbValue(rsTemp!sell_station_id)
    Set rsTemp = Nothing
    Set oDb = Nothing
End Sub

'**************************************************
'Member Code:S10
'Brief Description:向数据库中新增检票口
'Engineer:wjb
'Date Generated:1999/10/11
'Last Revision Date:2002/11/9
'**************************************************
Private Sub AddObject()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim szTempSql As String
    Dim szSqlField As String
    Dim szSqlContext As String
    If Trim(m_szCheckGateCode) = "" Then RaiseError ERR_CheckGateIDNotNull
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szTempSql = "INSERT checkgate_info (" _
          & "check_gate_id,check_gate_name,Annotation,sell_station_id)" _
          & "VALUES ('" & m_szCheckGateCode & "','" & m_szCheckGateName & "'," _
          & "'" & m_szAnnotation & "', '" & m_SellStationID & "')"
    
    oDb.Execute (szTempSql)
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertAddObjectError ERR_CheckGate, oDb
End Sub
'**************************************************
'Member Code:S11
'Brief Description:丛数据库中删除该检票口
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub DeleteObject()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    oDb.Execute ("DELETE checkgate_info WHERE check_gate_id='" & m_szCheckGateCode & "'")
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertDeleteObjectError ERR_CheckGate, oDb
End Sub

'**************************************************
'Member Code:S11
'Brief Description:丛数据库中修改检票口
'Engineer:wjb
'Date Generated:1999/10/11
'Last Revision Date:2002/11/9
'**************************************************
Private Sub UpdateToDB()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim szTempSql As String
    Dim szSqlField As String
    Dim szSqlWhere As String
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szTempSql = "UPDATE checkgate_info SET " _
        & "check_gate_name='" & m_szCheckGateName & "'," _
        & "Annotation='" & m_szAnnotation & "'," _
        & "sell_station_id='" & m_SellStationID & "'" _
        & "WHERE check_gate_id='" & m_szCheckGateCode & "'"
     
    oDb.Execute (szTempSql)
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertUpdateObjectError ERR_CheckGate, oDb
End Sub


'**************************************************
'Member Code:S12
'Brief Description:检票口新增检票员
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/25
'**************************************************
Public Sub AddChecker(Checker As String)
'**************************************************
'Checker(检票员代码)
'**************************************************
    Dim szSqlField As String
    Dim szSqlContext As String
    Dim szTempSql As String
    Dim oDb As New RTConnection
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szTempSql = "INSERT Checker_gate_info ("
    szSqlField = szSqlField & "check_gate_id,"
    szSqlField = szSqlField & "user_id)"
    szSqlContext = " VALUES ('" & m_szCheckGateCode & "',"
    szSqlContext = szSqlContext & " '" & Checker & "')"
    szTempSql = szTempSql & szSqlField & szSqlContext
    oDb.Execute (szTempSql)
    Set oDb = Nothing
End Sub

'**************************************************
'Member Code:S13
'Brief Description:删除检票员
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/25
'**************************************************
Public Sub DeleteChecker(Checker As String)
'**************************************************
'Checker(检票员代码)
'**************************************************
    Dim szTempSql As String
    Dim oDb As New RTConnection
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szTempSql = "DELETE Checker_gate_info WHERE check_gate_id='" & m_szCheckGateCode & "' AND user_id='" & Checker & "'"
    oDb.Execute szTempSql
    Set oDb = Nothing
End Sub
'**************************************************
'Member Code:
'Brief Description:返回（/设置）所属站的代码
'Engineer:wjb
'Date Generated:1999/10/11
'Last Revision Date:2002/11/9
'**************************************************
Public Property Let SellStationID(ByVal vData As String)
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    AssertActiveUserValid m_oActiveUser, ERR_CheckGate
    '查询该对象是否已有记录
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT * FROM sell_station_info WHERE sell_station_id='" & RTrim(vData) & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then RaiseError ERR_SellStationIDNotExist
    m_SellStationID = RTrim(vData)
    Set rsTemp = Nothing
    Set oDb = Nothing
End Property
Public Property Get SellStationID() As String
    SellStationID = m_SellStationID
End Property
'**************************************************
'Member Code:
'Brief Description:返回（/设置）所属站的名称
'Engineer:wjb
'Date Generated:1999/10/11
'Last Revision Date:2002/11/9
'**************************************************
Public Property Get SellStationName() As String
    AssertActiveUserValid m_oActiveUser, ERR_CheckGate
    SellStationName = m_SellStationName
End Property


