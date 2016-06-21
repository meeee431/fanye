VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Station"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'**********************************************************
'* Source File Name:Station.cls
'* Project Name:STBase.vbp
'* Engineer:
'* Data Generated:
'* Last Revision Date:
'* Brief Description:站点对象
'* Relational Document:
'**********************************************************
Option Explicit
'出错
Public Enum EErrStation
    ERR_StationExist = ERR_Station + ERR_AddDuplicate '56001站点已存在不能新增
    ERR_StationNoDelete = ERR_Station + ERR_DeleteChildExist '56007站点已安排路段不能删除
    ERR_StationNotEdit = ERR_Station + ERR_EditChildExist '56005站点无该地区
    ERR_StationNotAvailable = ERR_Station + ERR_NotAvailable '56012站点对象处于无状态
    ERR_StationNotAddStatus = ERR_Station + ERR_NotAddObj '56016站点对象不在新增状态
    ERR_StationNotExist = ERR_Station + 21 '56021站点无存在
    ERR_StationIDNotNull = ERR_Station + 22 '56022站点代码不能为空
End Enum

'权限
Public Enum ERightStation
    RIGHT_StationManagement = ERR_Station + cnMidRightBegin + cnMidRightStep * 1 '56106站点管理
End Enum


Private m_oActiveUser As ActiveUser 'P1所有的类都有的私有类变量
Private m_nObjectStatus As EObjectStatus 'P2所有的实体类都有的私有类变量
Private m_szStationID As String 'P3站点代码
Private m_szStationName As String 'P4站点名称
Private m_szStationInputCode As String 'P5站点输入码
Private m_szLocalCode As String 'P6站点本地码
Private m_nStationLevel As Integer 'P7站点等级
Private m_szAreaCode As String 'P8地区代码
Private m_szAreaName As String 'P9地区名称

'F1获得通过该站点的所有线路
Public Function GetAllRoute() As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT tbi.route_id,tbi.route_name FROM Route_section_lst tbo,route_info tbi WHERE tbo.end_station_id='" & m_szStationID & "' AND tbi.route_id=tbo.route_id "
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
        ReDim szaTemp(1 To nCount, 1 To 2) As String
        For i = 1 To nCount
            szaTemp(i, 1) = FormatDbValue(rsTemp!route_id)
            szaTemp(i, 2) = FormatDbValue(rsTemp!route_name)
            rsTemp.MoveNext
        Next
        GetAllRoute = szaTemp
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

'F2获得以该站点为终点站的路段
Public Function GetEndSection() As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute("SELECT tbs.*,tbi.station_name FROM section_info tbs,station_info tbi WHERE end_station='" & m_szStationID & "' AND tbs.end_station=tbi.station_id")
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
        ReDim szaTemp(1 To nCount, 1 To 8) As String
        For i = 1 To nCount
            szaTemp(i, 1) = FormatDbValue(rsTemp!section_id)
            szaTemp(i, 2) = FormatDbValue(rsTemp!section_name)
            szaTemp(i, 3) = FormatDbValue(rsTemp!start_station)
            szaTemp(i, 4) = FormatDbValue(rsTemp!end_station)
            szaTemp(i, 5) = FormatDbValue(rsTemp!Mileage)
            szaTemp(i, 6) = FormatDbValue(rsTemp!road_rank_code)
            szaTemp(i, 7) = FormatDbValue(oDb.Execute("SELECT station_name FROM station_info WHERE station_id='" & FormatDbValue(rsTemp!start_station) & "'")!station_name)
            szaTemp(i, 8) = FormatDbValue(oDb.Execute("SELECT station_name FROM station_info WHERE station_id='" & FormatDbValue(rsTemp!end_station) & "'")!station_name)
            rsTemp.MoveNext
        Next
        GetEndSection = szaTemp
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

'P3获得以该站点为起点站的路段
Public Function GetStartSection() As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute("SELECT * FROM section_info WHERE start_station='" & m_szStationID & "'")
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    ReDim szaTemp(1 To nCount, 1 To 8) As String
    For i = 1 To nCount
        szaTemp(i, 1) = FormatDbValue(rsTemp!section_id)
        szaTemp(i, 2) = FormatDbValue(rsTemp!section_name)
        szaTemp(i, 3) = FormatDbValue(rsTemp!start_station)
        szaTemp(i, 4) = FormatDbValue(rsTemp!end_station)
        szaTemp(i, 5) = FormatDbValue(rsTemp!Mileage)
        szaTemp(i, 6) = FormatDbValue(rsTemp!road_rank_code)
        szaTemp(i, 7) = FormatDbValue(oDb.Execute("SELECT station_name FROM station_info WHERE station_id='" & FormatDbValue(rsTemp!start_station) & "'")!station_name)
        szaTemp(i, 8) = FormatDbValue(oDb.Execute("SELECT station_name FROM station_info WHERE station_id='" & FormatDbValue(rsTemp!end_station) & "'")!station_name)
        rsTemp.MoveNext
    Next
    GetStartSection = szaTemp
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
        ShowError ERR_AddObj
    ElseIf m_nObjectStatus = ST_EditObj Then
        ShowError ERR_EditObj
    ElseIf m_nObjectStatus = ST_NormalObj Then
        ShowError ERR_NormalObj
    Else
        Set m_oActiveUser = vNewValue
    End If
End Property
'P2
Public Property Get ObjStatus() As EObjectStatus
    ObjStatus = m_nObjectStatus
End Property
'P3
Public Property Let StationID(ByVal vData As String)
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    AssertActiveUserValid m_oActiveUser, ERR_Station
    '查询是否是新增状态
    If m_nObjectStatus <> ST_AddObj Then ShowError ERR_StationNotAddStatus
    '输入的主键是否为空
    If vData = "" Then ShowError ERR_StationIDNotNull
    '查询该对象是否已有记录
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT station_id FROM station_info WHERE station_id='" & RTrim(vData) & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 1 Then ShowError ERR_StationExist
    m_szStationID = RTrim(vData)
    Set rsTemp = Nothing
    Set oDb = Nothing
End Property
'P3
Public Property Get StationID() As String
    StationID = m_szStationID
End Property
'P4
Public Property Let StationName(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szStationName = vData
End Property
'P4
Public Property Get StationName() As String
    StationName = m_szStationName
End Property
'P5
Public Property Let StationInputCode(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szStationInputCode = vData
End Property
'P5
Public Property Get StationInputCode() As String
    StationInputCode = m_szStationInputCode
End Property

'P6
Public Property Let LocalCode(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szLocalCode = vData
End Property
'P6
Public Property Get LocalCode() As String
    LocalCode = m_szLocalCode
End Property
'P7
Public Property Let StaionLevel(ByVal vData As Integer)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_nStationLevel = vData
    If m_nStationLevel <> 2 Then m_nStationLevel = 1
End Property
'P7
Public Property Get StaionLevel() As Integer
    StaionLevel = m_nStationLevel
End Property
'P8
Public Property Get AreaCode() As String
    AreaCode = m_szAreaCode
End Property
'P8
Public Property Let AreaCode(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szAreaCode = vData
End Property
'P9
Public Property Get AreaName() As String
    AreaName = m_szAreaName
End Property


'S1
Public Sub Init(poAUser As ActiveUser)
    Set m_oActiveUser = poAUser
    'AssertHaveRight m_oActiveUser, RIGHT_StationManagement
End Sub

'S2类刚建立时处于不可用状态,此时类不能使用
Private Sub Class_Initialize()
    m_nObjectStatus = ST_NotAvailable
End Sub

'S3测试对象的状态是否有效（无效状态[ST_NotAvailable]和新增状态[ST_AddObj]则出错）
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then ShowError ERR_StationNotAvailable
End Sub
'S4
Private Sub AssertObjIsValid() '测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
    AssertActiveUserValid m_oActiveUser, ERR_Station
    AssertStatusAvailable
End Sub
'S5
Public Sub Identify(Optional StationID As String = "", Optional StationInput As String = "")
    AssertActiveUserValid m_oActiveUser, ERR_Station
    m_szStationID = RTrim(StationID)
    m_szStationInputCode = RTrim(StationInput)
    RefreshMemoryInfo
    m_nObjectStatus = ST_NormalObj
End Sub

'S6调用新增方法后，对象处于新增状态。
Public Sub AddNew()
    AssertHaveRight m_oActiveUser, RIGHT_StationManagement
    AssertActiveUserValid m_oActiveUser, ERR_Station
    m_nObjectStatus = ST_AddObj
End Sub
'S7
Public Sub Delete() '调用删除方法后，对象处于不可用状态
    AssertHaveRight m_oActiveUser, RIGHT_StationManagement
    AssertObjIsValid
    DeleteObject
    m_nObjectStatus = ST_NotAvailable
End Sub
'S8
Public Sub Update()
    AssertHaveRight m_oActiveUser, RIGHT_StationManagement
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

'S9
Private Sub RefreshMemoryInfo() '按照对象现在的ID将数据库中相应的信息读至对象内
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    If m_szStationID = "" And m_szStationInputCode <> "" Then
        Set rsTemp = oDb.Execute("SELECT tbs.*,tba.area_name FROM station_info tbs,area_code tba WHERE station_input_code='" & m_szStationInputCode & "' AND tbs.area_code=tba.area_code")
    Else
        Set rsTemp = oDb.Execute("SELECT tbs.*,tba.area_name FROM station_info tbs,area_code tba WHERE station_id='" & m_szStationID & "' AND tbs.area_code=tba.area_code")
    End If
    If rsTemp.RecordCount = 0 Then ShowError ERR_StationNotExist
    m_szStationName = FormatDbValue(rsTemp!station_name)
    m_szStationID = FormatDbValue(rsTemp!station_id)
    m_szLocalCode = FormatDbValue(rsTemp!local_station_code)
    m_szStationInputCode = FormatDbValue(rsTemp!station_input_code)
    m_nStationLevel = FormatDbValue(rsTemp!station_rank)
    m_szAreaName = FormatDbValue(rsTemp!area_name)
    m_szAreaCode = FormatDbValue(rsTemp!area_code)
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
    
    If Trim(m_szStationID) = "" Then ShowError ERR_StationIDNotNull
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szTempSql = "INSERT station_info ("
    szSqlField = szSqlField & "station_id,"
    szSqlField = szSqlField & "station_name,"
    szSqlField = szSqlField & "local_station_code,"
    szSqlField = szSqlField & "station_input_code,"
    szSqlField = szSqlField & "station_rank,area_code)"
    szSqlContext = " VALUES ('" & m_szStationID & "',"
    szSqlContext = szSqlContext & " '" & m_szStationName & "',"
    szSqlContext = szSqlContext & " '" & m_szLocalCode & "',"
    szSqlContext = szSqlContext & " '" & m_szStationInputCode & "',"
    szSqlContext = szSqlContext & " " & m_nStationLevel & ",'" & m_szAreaCode & "')"
    szTempSql = szTempSql & szSqlField & szSqlContext
    oDb.Execute (szTempSql)
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertAddObjectError ERR_Station, oDb
End Sub
'S11
Private Sub DeleteObject()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    oDb.Execute ("DELETE station_info WHERE station_id='" & m_szStationID & "'")
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertDeleteObjectError ERR_Station, oDb
End Sub

'S12
Private Sub UpdateToDB()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim szTempSql As String
    Dim szSqlField As String
    Dim szSqlWhere As String
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    If m_nStationLevel <> TP_CanNotSellTicket Then m_nStationLevel = TP_CanSellTicket
    szTempSql = "UPDATE station_info SET "
    szSqlField = szSqlField & "station_name='" & m_szStationName & "',"
    szSqlField = szSqlField & "local_station_code='" & m_szLocalCode & "',"
    szSqlField = szSqlField & "station_input_code='" & m_szStationInputCode & "',"
    szSqlField = szSqlField & "station_rank=" & m_nStationLevel & " , "
    szSqlField = szSqlField & "area_code='" & m_szAreaCode & "'  "
    
    
    szSqlWhere = " WHERE station_id='" & m_szStationID & "'"
    szTempSql = szTempSql & szSqlField & szSqlWhere
    oDb.Execute (szTempSql)
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertUpdateObjectError ERR_Station, oDb
End Sub

