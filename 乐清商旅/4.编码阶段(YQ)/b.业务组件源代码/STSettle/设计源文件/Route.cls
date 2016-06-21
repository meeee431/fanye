VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BackRoute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'**********************************************************
'* Source File Name:BackRoute.cls
'* Project Name:STSettle.vbp
'* Engineer:
'* Data Generated:2004/03/12
'* Last Revision Date:2004/03/12
'* Brief Description:回程线路对象
'* Relational Document:
'**********************************************************
Option Explicit
'出错
Public Enum EErrRoute
    ERR_RouteExist = ERR_Route + ERR_AddDuplicate '16801线路已存在不能新增
    ERR_RouteNoDelete = ERR_Route + ERR_DeleteChildExist '16807线路已按排车次不能删除
    ERR_RouteNotAvailable = ERR_Route + ERR_NotAvailable '16812线路对象处于无状态
    ERR_RouteNotAddStatus = ERR_Route + ERR_NotAddObj '16816线路对象不在新增状态
    ERR_RouteNotExist = ERR_Route + 21 '16821线路无存在
    ERR_RouteAddSectionStartStationNotEndStation = ERR_Route + 22 '16822新增路段的起点站不是线路的终点站
    ERR_RouteIDNotNull = ERR_Route + 23 '16823线路代码不能为空
    ERR_RouteSectionNotExist = ERR_Route + 24 '16824线路的终点站无对应路段，不能删除，可能是线路数据不一致。
    ERR_RouteNotCanDulStation = ERR_Route + 25 '16825线路中不能又重复的站点
    ERR_NotMileageSetion = ERR_Route + 26 '路段没有里程或其它信息
End Enum

'权限
Public Enum ERightRoute
    RIGHT_RouteManagement = ERR_Route + cnMidRightBegin + cnMidRightStep * 1 '16906管理线路的权限
'    RIGHT_RouteInserStation = ERR_Route + cnMidRightBegin + cnMidRightStep * 2 '16912线路插入站点
End Enum


Private m_oActiveUser As ActiveUser 'P1所有的类都有的私有类变量
Private m_nObjectStatus As EObjectStatus 'P2所有的实体类都有的私有类变量
Private m_szRouteID As String 'P3线路代码
Private m_szRouteName As String 'P4线路名称
Private m_szStartStation As String 'P5线路起点站
Private m_szEndStation As String 'P6线路终点站
Private m_stMileage As Double 'P7线路里程数
Private m_eStatus As ERouteStatus 'P8线路状态
'Private m_szTicketPriceFormula As String 'P9线路票价计算公式
Private m_szAnnotation As String 'P10线路注释
Private m_szStartStationName As String 'P11
Private m_szEndStationName As String 'P12



'获得该线路第一站点
Public Function GetFirstSellStation() As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
On Error GoTo ErrHandle
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute("SELECT r.sell_station_id , s.station_name FROM back_route_section_lst r, station_info s " _
        & " WHERE r.sell_station_id = s.station_id AND route_id= '" & m_szRouteID & "'" _
        & " AND end_station_mileage= ( SELECT MAX(end_station_mileage) " _
        & " FROM  back_route_section_lst  WHERE route_id= '" & m_szRouteID & "')")
    GetFirstSellStation = FormatDbValue(rsTemp!sell_station_id)
    Set rsTemp = Nothing
    Set oDb = Nothing
    Exit Function
ErrHandle:
    err.Raise err.Number
End Function

''获得该线路第一站点
Public Function GetStartStation() As String()
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
On Error GoTo ErrHandle
    AssertActiveUserValid m_oActiveUser, ERR_Route
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Dim szSql As String
    szSql = "SELECT b.*,c.area_name FROM sell_station_info a,station_info b,area_code c WHERE a.station_id=b.station_id AND b.area_code=c.area_code"
    Set rsTemp = oDb.Execute(szSql)

    Dim szaTemp() As String
    If rsTemp.RecordCount > 0 Then
        ReDim szaTemp(1 To rsTemp.RecordCount, 1 To 6) As String
        Dim i As Integer
        For i = 1 To rsTemp.RecordCount
            szaTemp(i, 1) = FormatDbValue(rsTemp!station_id)
            szaTemp(i, 2) = FormatDbValue(rsTemp!station_name)
            szaTemp(i, 3) = FormatDbValue(rsTemp!station_input_code)
            szaTemp(i, 4) = FormatDbValue(rsTemp!station_rank)
            szaTemp(i, 5) = FormatDbValue(rsTemp!local_station_code)
            szaTemp(i, 6) = FormatDbValue(rsTemp!area_name)
            rsTemp.MoveNext
        Next
    End If
    GetStartStation = szaTemp
    Set rsTemp = Nothing
    Set oDb = Nothing
    Exit Function
ErrHandle:
    err.Raise err.Number
End Function

'F2获得路段信息
Public Function GetSectionInfo() As TRouteSectionInfo()
    Dim oDb As New RTConnection
    Dim RSInfo() As TRouteSectionInfo
    Dim i As Integer
    Dim szSql As String
    Dim rsTemp As Recordset
On Error GoTo ErrHandle '出错
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT * FROM back_route_section_lst WHERE route_id= '" & m_szRouteID & "'" _
          & " AND sell_station_id in (SELECT sell_station_id FROM back_route_section_lst " _
          & " WHERE end_station_mileage= ( SELECT MAX(end_station_mileage) " _
          & " FROM back_route_section_lst WHERE  route_id= '" & m_szRouteID & "'))" _
          & " ORDER BY back_route_section_lst.section_serial"
    Set rsTemp = oDb.Execute(szSql)
       
           
    If rsTemp.RecordCount <> 0 Then
    ReDim RSInfo(1 To rsTemp.RecordCount) As TRouteSectionInfo
    For i = 1 To rsTemp.RecordCount
        RSInfo(i).szSectionID = FormatDbValue(rsTemp!section_id)
        RSInfo(i).nEndStationType = FormatDbValue(rsTemp!station_type)
        RSInfo(i).sgEndStationMileage = FormatDbValue(rsTemp!end_station_mileage)
        rsTemp.MoveNext
    Next
    GetSectionInfo = RSInfo
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
    Exit Function
ErrHandle:
    err.Raise err.Number
End Function

'F3获得在该线路上运行的车次
Public Function GetAllBus(PlanID As String) As String()
'    Dim szaTemp() As String
'    Dim oDb As New RTConnection
'    Dim i As Integer
'    Dim rsTemp As Recordset
'    AssertObjIsValid
'    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
'    Set rsTemp = oDb.Execute("SELECT bus_id FROM bus_info WHERE route_id='" & m_szRouteID & "'")
'    If rsTemp.RecordCount <> 0 Then
'    ReDim szaTemp(1 To rsTemp.RecordCount) As String
'    For i = 1 To rsTemp.RecordCount
'        szaTemp(i) = FormatDbValue(rsTemp!Bus_ID)
'        rsTemp.MoveNext
'    Next
'    GetAllBus = szaTemp
'    End If
'    Set rsTemp = Nothing
'    Set oDb = Nothing
End Function

'F4扩展的获得路段信息
Public Function GetSectionInfoEx(Optional pszSectionID As String = "") As TRouteSectionInfoEx()
    'pszSectionID路段代码
    Dim oDb As New RTConnection
    Dim RSInfo() As TRouteSectionInfoEx
    Dim i As Integer
    Dim rsTemp As Recordset
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Dim szSql As String
    
    szSql = " SELECT tbr.*,tbs.section_name,tbs.start_station,tbb.station_name AS bstation_name,tbe.station_name AS estation_name " _
        & " FROM back_route_section_lst tbr,section_info tbs,station_info tbb,station_info tbe " _
        & " WHERE route_id='" & m_szRouteID & "' AND tbr.section_id=tbs.section_id AND tbs.start_station=tbb.station_id " _
        & " AND tbs.end_station=tbe.station_id " _
        & " AND tbr.sell_station_id in (SELECT r.sell_station_id FROM back_route_section_lst r  WHERE r.end_station_mileage" _
        & " =( SELECT MAX(r.end_station_mileage)  FROM  back_route_section_lst r WHERE r.route_id= '" & m_szRouteID & "')AND route_id = '" & m_szRouteID & "')"
        
    If pszSectionID <> "" Then
        szSql = szSql & " AND tbr.section_id = '" & pszSectionID & "' "
    End If
        szSql = szSql & " ORDER BY section_serial"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 0 Then
        ReDim RSInfo(1 To rsTemp.RecordCount) As TRouteSectionInfoEx
    End If
    For i = 1 To rsTemp.RecordCount
        RSInfo(i).szSectionID = FormatDbValue(rsTemp!section_id)
'        RSInfo(i).nEndStationType = FormatDbValue(rsTemp!station_type)
        RSInfo(i).sgEndStationMileage = FormatDbValue(rsTemp!end_station_mileage)
        RSInfo(i).szEndStationID = FormatDbValue(rsTemp!end_station_id)
        RSInfo(i).nSectionSerial = FormatDbValue(rsTemp!section_serial)
        RSInfo(i).szStartStationID = FormatDbValue(rsTemp!start_station)
        RSInfo(i).szSectionName = FormatDbValue(rsTemp!section_name)
        RSInfo(i).szEndStationName = FormatDbValue(rsTemp!estation_name)
        RSInfo(i).szStartStationName = FormatDbValue(rsTemp!bstation_name)
        rsTemp.MoveNext
    Next
    GetSectionInfoEx = RSInfo

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
Public Property Let RouteID(ByVal vData As String)
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    AssertActiveUserValid m_oActiveUser, ERR_Route
    '查询是否是新增状态
    If m_nObjectStatus <> ST_AddObj Then RaiseError ERR_RouteNotAddStatus
    '输入的主键是否为空
    If vData = "" Then RaiseError ERR_RouteIDNotNull
    '查询该对象是否已有记录
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT route_id FROM back_route_info WHERE route_id='" & RTrim(vData) & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 1 Then RaiseError ERR_RouteExist
    m_szRouteID = RTrim(vData)
    Set rsTemp = Nothing
    Set oDb = Nothing
End Property
'P3
Public Property Get RouteID() As String
    RouteID = m_szRouteID
End Property
'P4
Public Property Let RouteName(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szRouteName = vData
End Property
'P4
Public Property Get RouteName() As String
    RouteName = m_szRouteName
End Property
'P5线路的起点站不能修改
Public Property Get StartStation() As String
    StartStation = m_szStartStation
End Property
'P6线路的终点站不能修改
Public Property Get EndStation() As String
    EndStation = m_szEndStation
End Property
'P7线路的里程数不能修改
Public Property Get Mileage() As Double
    Mileage = m_stMileage
End Property
'P8
Public Property Let Status(ByVal vData As ERouteStatus)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_eStatus = vData
End Property
'P8
Public Property Get Status() As ERouteStatus
    Status = m_eStatus
End Property

'P10
Public Property Let Annotation(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szAnnotation = vData
End Property
'P10
Public Property Get Annotation() As String
    Annotation = m_szAnnotation
End Property
'P11
Public Property Get StartStationName() As String
    StartStationName = m_szStartStationName
End Property
'P12
Public Property Get EndStationName() As String
    EndStationName = m_szEndStationName
End Property
'S1
Public Sub Init(poAUser As ActiveUser)
    Set m_oActiveUser = poAUser
    'AssertHaveRight m_oActiveUser, RIGHT_RouteManagement
End Sub

'S2类刚建立时处于不可用状态,此时类不能使用
Private Sub Class_Initialize()
    m_nObjectStatus = ST_NotAvailable
End Sub

'S3测试对象的状态是否有效（无效状态[ST_NotAvailable]和新增状态[ST_AddObj]则出错）
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then RaiseError ERR_RouteNotAvailable
End Sub
'S4
Private Sub AssertObjIsValid() '测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
    AssertActiveUserValid m_oActiveUser, ERR_Route
    AssertStatusAvailable
End Sub
'S5
Public Sub Identify(Route As String)
    AssertActiveUserValid m_oActiveUser, ERR_Route
    m_szRouteID = RTrim(Route)
    RefreshMemoryInfo
    m_nObjectStatus = ST_NormalObj
End Sub

'S6调用新增方法后，对象处于新增状态。
Public Sub AddNew()
    AssertHaveRight m_oActiveUser, RIGHT_RouteManagement
    AssertActiveUserValid m_oActiveUser, ERR_Route
    m_nObjectStatus = ST_AddObj
End Sub
'S7
Public Sub Delete() '调用删除方法后，对象处于不可用状态
    AssertHaveRight m_oActiveUser, RIGHT_RouteManagement
    AssertObjIsValid
    DeleteObject
    m_nObjectStatus = ST_NotAvailable
End Sub
'S8
Public Sub Update()
    AssertHaveRight m_oActiveUser, RIGHT_RouteManagement
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
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT tbr.*,tbs.station_name as startname,tbe.station_name AS endname FROM back_route_info tbr,station_info " _
            & " tbs,station_info tbe WHERE tbr.route_id='" & m_szRouteID & "'" _
            & " AND tbr.start_station_id=tbs.station_id AND tbr.end_station_id=tbe.station_id"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then
        szSql = "SELECT *,'' AS startname,'' AS endname FROM back_route_info WHERE route_id='" & m_szRouteID & "'"
    Set rsTemp = oDb.Execute(szSql)
    End If
    If rsTemp.RecordCount = 0 Then RaiseError ERR_RouteNotExist
    m_szAnnotation = FormatDbValue(rsTemp!Annotation)
    m_szEndStation = FormatDbValue(rsTemp!end_station_id)
    m_szRouteName = FormatDbValue(rsTemp!route_name)
    m_szStartStation = FormatDbValue(rsTemp!start_station_id)
    
    m_stMileage = FormatDbValue(rsTemp!Mileage)
    m_eStatus = FormatDbValue(rsTemp!Status)
    m_szStartStationName = FormatDbValue(rsTemp!startname)
    m_szEndStationName = FormatDbValue(rsTemp!endname)
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
    If Trim(m_szRouteID) = "" Then RaiseError ERR_RouteIDNotNull
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    m_szEndStation = ""
    m_szStartStation = ""
    
    m_stMileage = 0
    m_eStatus = ST_RouteNotAvailable
    szTempSql = "INSERT back_route_info ("
    szSqlField = szSqlField & "route_id,"
    szSqlField = szSqlField & "Annotation,"
    szSqlField = szSqlField & "end_station_id,"
    szSqlField = szSqlField & "route_name,"
    szSqlField = szSqlField & "start_station_id,"
    szSqlField = szSqlField & "Mileage,"
    szSqlField = szSqlField & "Status)"
    szSqlContext = " VALUES ('" & m_szRouteID & "',"
    szSqlContext = szSqlContext & " '" & m_szAnnotation & "',"
    szSqlContext = szSqlContext & " '" & m_szEndStation & "',"
    szSqlContext = szSqlContext & " '" & m_szRouteName & "',"
    szSqlContext = szSqlContext & " '" & m_szStartStation & "',"
    szSqlContext = szSqlContext & " " & m_stMileage & ","
    szSqlContext = szSqlContext & " " & m_eStatus & ")"
    szTempSql = szTempSql & szSqlField & szSqlContext
    oDb.Execute (szTempSql)
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertAddObjectError ERR_Route, oDb
End Sub
'S11
Private Sub DeleteObject()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    oDb.Execute ("DELETE back_route_info WHERE route_id='" & m_szRouteID & "'")
    
    WriteOperateLog m_oActiveUser, RIGHT_RouteManagement, "线路ID[" & m_szRouteID & "]被删除"
    
    
    Set oDb = Nothing
    
    Exit Sub
ErrHandle:
    AssertDeleteObjectError ERR_Route, oDb
   
End Sub

'S12
Private Sub UpdateToDB()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim szTempSql As String
    Dim szSqlField As String
    Dim szSqlWhere As String
    If Trim(m_szRouteID) = "" Then RaiseError ERR_RouteIDNotNull
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szTempSql = "UPDATE back_route_info SET "
    szSqlField = szSqlField & "Annotation='" & m_szAnnotation & "',"
    szSqlField = szSqlField & "end_station_id='" & m_szEndStation & "',"
    szSqlField = szSqlField & "route_name='" & m_szRouteName & "',"
    szSqlField = szSqlField & "start_station_id='" & m_szStartStation & "',"
    szSqlField = szSqlField & "Mileage=" & m_stMileage & ","
    szSqlField = szSqlField & "Status=" & m_eStatus & ""
    szSqlWhere = " WHERE route_id='" & m_szRouteID & "'"
    szTempSql = szTempSql & szSqlField & szSqlWhere
    oDb.Execute (szTempSql)
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertUpdateObjectError ERR_Route, oDb
End Sub



'S14删除线路的最后的路段,对线路路段的修改可能影响到票价管理和计划中的车次站点表，该处未考虑
Public Sub DeleteLastSection()
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szEndSection As String
    Dim sgMileage As Double
    Dim bLastOneSection As Boolean
    Dim szTempSql As String
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    '--------获得终点站的代码
    szTempSql = "SELECT section_id FROM back_route_section_lst WHERE route_id='" & m_szRouteID & "'" _
          & "AND sell_station_id in (SELECT sell_station_id FROM back_route_section_lst " _
          & "WHERE end_station_mileage= ( SELECT MAX(end_station_mileage) " _
          & "FROM  back_route_section_lst WHERE route_id= '" & m_szRouteID & "'))" _
          & "ORDER BY end_station_mileage DESC"
    Set rsTemp = oDb.Execute(szTempSql)
    If rsTemp.RecordCount = 0 Then RaiseError ERR_SectionNotExist
    If rsTemp.RecordCount = 1 Then bLastOneSection = True
    szEndSection = FormatDbValue(rsTemp!section_id)
    szTempSql = "SELECT mileage,start_station FROM section_info WHERE section_id='" & szEndSection & "'"
    Set rsTemp = oDb.Execute(szTempSql)
    '--------------------
    '获得里程数和终点站
    sgMileage = FormatDbValue(rsTemp!Mileage)
    m_szEndStation = FormatDbValue(rsTemp!start_station) '修改终点站
    szTempSql = "DELETE back_route_section_lst WHERE route_id='" & m_szRouteID & "' AND section_id='" & szEndSection & "'"
    oDb.Execute (szTempSql)
    m_stMileage = m_stMileage - sgMileage '修改线路里程数
    If bLastOneSection Then
        m_szEndStation = ""
        m_szStartStation = ""
        m_stMileage = 0
    End If
    UpdateToDB
    Set rsTemp = Nothing
    Set oDb = Nothing
End Sub

'S15线路新增最后的路段，对线路的修改可能影响到票价管理和计划中的车次站点表，该处未考虑
Public Sub AddLastSection(SectionID As String)
    On Error GoTo ErrHandle
    
    Dim oDb As New RTConnection
    Dim nMaxSerial As Integer
    Dim szSellStationID As String
    Dim stSectionMileage As Double
    Dim szUpSectionEndStation As String
    Dim stEndStationMileage As Double
    Dim szEndStationID As String
    Dim szFSectionID As String
    Dim nCanSell As Byte
    Dim rsTemp As Recordset
'    Dim rsTemp2 As Recordset
    Dim szTempSql As String
    Dim szSqlField As String
    Dim szSqlWhere As String
    Dim szRouteStartStationID As String
    Dim szSqlContext As String
    Dim szStartStationID As String
    
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute("SELECT * FROM section_info WHERE section_id='" & Trim(SectionID) & "'")
         
    If rsTemp.RecordCount = 0 Then RaiseError ERR_SectionNotExist
    stSectionMileage = FormatDbValue(rsTemp!Mileage)
    szEndStationID = FormatDbValue(rsTemp!end_station)
    szStartStationID = FormatDbValue(rsTemp!start_station)
    
    nCanSell = FormatDbValue(oDb.Execute("SELECT station_rank FROM station_info WHERE station_id='" & szEndStationID & "'")!station_rank)
    
    '1查询线路的最终路段的路段代码和终点站的里程数和最大序号，并获得该路段的终点站，如果无路段则新增路段不限
    szTempSql = "SELECT sell_station_id , MAX(section_serial) section_serial,end_station_mileage,section_id FROM back_route_section_lst " _
              & "WHERE route_id='" & m_szRouteID & "' GROUP BY sell_station_id,section_id,end_station_mileage " _
              & " ORDER BY MAX(section_serial) DESC"
    Set rsTemp = oDb.Execute(szTempSql)
    If rsTemp.RecordCount <> 0 Then
        nMaxSerial = FormatDbValue(rsTemp!section_serial) '获得最终站点站点序号
        szSellStationID = FormatDbValue(rsTemp!sell_station_id) '获得起点站代码
    End If
    
    If nMaxSerial <> 0 Then '如果该路段不是该线路的第一个路段则
        stEndStationMileage = FormatDbValue(rsTemp!end_station_mileage) '获得最终站点的里程数
        '--------------------
        '获得上一路段的终点站
        szFSectionID = FormatDbValue(rsTemp!section_id)
        szTempSql = "SELECT end_station FROM section_info WHERE section_id='" & szFSectionID & "'"
        Set rsTemp = oDb.Execute(szTempSql)
        szUpSectionEndStation = FormatDbValue(rsTemp!end_station)
        '------------------------
        '查询该终点站是否重复
        szTempSql = "SELECT COUNT(*) AS station_count FROM back_route_section_lst WHERE end_station_id='" & szEndStationID _
              & "' AND route_id='" & Trim(m_szRouteID) & "' AND sell_station_id='" & Trim(szSellStationID) & "'"
        If oDb.Execute(szTempSql)!station_count >= 1 Then RaiseError ERR_RouteNotCanDulStation
    
        If Not (szUpSectionEndStation = szStartStationID Or szUpSectionEndStation = "") Then RaiseError ERR_RouteAddSectionStartStationNotEndStation
    End If
    
    '------------
    szTempSql = "SELECT start_station FROM section_info c WHERE section_id='" & SectionID & "'"
    Set rsTemp = oDb.Execute(szTempSql)
    If rsTemp.RecordCount > 0 Then
        szSellStationID = FormatDbValue(rsTemp!start_station)
    Else
        szSellStationID = ""
    End If
    If nMaxSerial = 0 Then
        '第一次登记时，更改当前线路的起点站属性
        m_szStartStation = szStartStationID
    End If
    
    
    oDb.BeginTrans
        '得到上车站数,如果上车站数大于1,则,批量插入,否则,只插入一个一个站点
        
    '    szTempSql = "SELECT * FROM sell_station_info "
    '    Set rsTemp2 = oDb.Execute(szTempSql)
            
        '增加起点站后的上车站对应的路段信息表 ,已经有路段加入了,继续插入其他发车站的路段
        szTempSql = "INSERT back_route_section_lst (sell_station_id,route_id,section_id,section_serial,end_station_id,end_station_mileage)" _
                    & " SELECT sell_station_id,route_id," & TransFieldValueToString(SectionID) _
                    & "     , MAX(section_serial)+1," & TransFieldValueToString(szEndStationID) _
                    & "     , MAX(end_station_mileage)+" & TransFieldValueToString(stSectionMileage) _
                    & " FROM back_route_section_lst WHERE route_id=" & TransFieldValueToString(m_szRouteID) _
                    & "     AND end_station_id = " & TransFieldValueToString(szUpSectionEndStation) _
                    & " GROUP BY sell_station_id,route_id"
                    
        oDb.Execute szTempSql
        '还没有该起点站的路段记录,插入该上车站的第一个路段
        If szSellStationID <> "" Then
            szTempSql = "INSERT back_route_section_lst (sell_station_id,route_id,section_id,section_serial,end_station_id,end_station_mileage)" & _
                        " VALUES(" & TransFieldValueToString(szSellStationID) & "," & _
                                    TransFieldValueToString(m_szRouteID) & "," & _
                                    TransFieldValueToString(SectionID) & "," & _
                                    TransFieldValueToString(1) & "," & _
                                    TransFieldValueToString(szEndStationID) & "," & _
                                    TransFieldValueToString(stSectionMileage) & ")"
            oDb.Execute szTempSql
        End If
    oDb.CommitTrans
    
    '修改线路的终点站和线路的里程数
    m_szEndStation = szEndStationID
    m_stMileage = stSectionMileage + stEndStationMileage
    m_eStatus = ST_RouteAvailable
    UpdateToDB
    Set rsTemp = Nothing
    Set oDb = Nothing
    
    Exit Sub
ErrHandle:
    oDb.RollbackTrans
    RaiseError err.Number, , err.Description
End Sub

'返回线路上的所有站点
Public Function RouteStation() As String()
    Dim oDb As New RTConnection
    Dim szTemp() As String
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute("SELECT end_station_id FROM back_route_section_lst WHERE route_id='" & m_szRouteID & "' ORDER BY section_serial")
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
        ReDim szTemp(1 To nCount)
        For i = 1 To nCount
            szTemp(i) = FormatDbValue(rsTemp!end_station_id)
            rsTemp.MoveNext
        Next
    End If
    RouteStation = szTemp
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

'返回线路上的所有站点、里程数、站点名
Public Function RouteStationEx() As String()
    Dim oDb As New RTConnection
    Dim szTemp() As String
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    
    Dim szSql As String
    
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT a.end_station_id, a.end_station_mileage,  b.station_name " _
        & " FROM back_route_section_lst a, station_info b ,back_route_info r " _
        & " WHERE a.route_id='" & m_szRouteID & "' and a.end_station_id = b.station_id " _
        & " AND a.route_id = r.route_id AND r.start_station_id = s.station_id " _
        & " ORDER BY section_serial"
        
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
        ReDim szTemp(1 To nCount, 1 To 3)
        For i = 1 To nCount
            szTemp(i, 1) = FormatDbValue(rsTemp!end_station_id)
            szTemp(i, 2) = FormatDbValue(rsTemp!end_station_mileage)
            szTemp(i, 3) = FormatDbValue(rsTemp!station_name)
            rsTemp.MoveNext
        Next
    End If
    RouteStationEx = szTemp
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function


Public Sub CloneRoute(szNewRouteID As String, Optional szNewRouteName As String, Optional szStationID As String = "")
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim nSerial As String
    Dim nMileage As Double
    Dim szStationName As String
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_RouteManagement
    On Error GoTo Here
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT * From back_route_info WHERE route_id='" & Trim(szNewRouteID) & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 0 Then RaiseError ERR_RouteExist
    oDb.BeginTrans
    '插入线路
    If szNewRouteName = "" Then
        szSql = "INSERT INTO  back_route_info (" _
            & " route_id, route_name,start_station_id,end_station_id,Mileage, Status, formula_name ,annotation ) " _
            & "  SELECT   '" & szNewRouteID & "'as route_id,   route_name, start_station_id, end_station_id, Mileage, Status, formula_name,annotation" _
            & " FROM  back_route_info WHERE  route_id='" & m_szRouteID & "'"
    Else
        szSql = "INSERT INTO  back_route_info (" _
            & " route_id, route_name,start_station_id,end_station_id,Mileage, Status, formula_name ,annotation ) " _
            & "  SELECT   '" & szNewRouteID & "'as route_id,  '" & szNewRouteName & "' as route_name, start_station_id, end_station_id, Mileage, Status, formula_name,annotation" _
            & " FROM  back_route_info WHERE  route_id='" & m_szRouteID & "'"
    End If
    oDb.Execute szSql
    '插入线路路段
    szSql = "INSERT INTO back_route_section_lst ( route_id, sell_station_id , section_id ,section_serial ,end_station_id ," _
            & " end_station_mileage , station_type ) " _
            & " SELECT  '" & szNewRouteID & "' as route_id, sell_station_id , section_id ,section_serial ," _
            & " end_station_id , end_station_mileage, station_type " _
            & " FROM  back_route_section_lst WHERE route_id='" & m_szRouteID & "' "
    oDb.Execute szSql
    If szStationID <> "" Then
        '取新终点站信息-----如无不处理
        szSql = "SELECT rs.end_station_mileage,rs.section_serial,st.station_name FROM back_route_section_lst  rs," _
               & " station_info st WHERE route_id='" & szNewRouteID & "'and  end_station_id='" & szStationID & "'"
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount <> 0 Then
            nMileage = FormatDbValue(rsTemp!end_station_mileage)
            nSerial = FormatDbValue(rsTemp!section_serial)
            szStationName = FormatDbValue(rsTemp!station_name)
            '更新线路表
            szSql = "UPDATE back_route_info SET end_station_id='" & szStationID & "' ,Mileage=" & nMileage & " WHERE route_id='" & szNewRouteID & "'"
            oDb.Execute szSql
            '删除新终点站以后路段
            szSql = "DELETE back_route_section_lst WHERE route_id='" & szNewRouteID & "' and section_serial> '" & nSerial & "'"
            oDb.Execute szSql
        End If
    End If
    oDb.CommitTrans
    Set oDb = Nothing
    Set rsTemp = Nothing
    WriteOperateLog m_oActiveUser, RIGHT_RouteManagement, "复制线路" & m_szRouteID & "为线路[" & szNewRouteID & "]"
Exit Sub
Here:
    oDb.RollbackTrans
    Set oDb = Nothing
    Set rsTemp = Nothing
    err.Raise err.Number
End Sub

Public Function GetAllSectionInfo(Optional pszSellStationID As String = "", Optional pszSectionID As String = "") As String()
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset, i As Integer
    Dim szTemp As String
    Dim aszTemp() As String
    oDb.ConnectionString = GetConnectionStr("")
    If Trim(pszSellStationID) <> "" Then
        szTemp = " AND r.Sell_station_id = " & TransFieldValueToString(pszSectionID)
    End If
    If Trim(pszSectionID) <> "" Then
        szTemp = szTemp & " AND r.section_id = " & TransFieldValueToString(pszSectionID)
    End If
    szSql = "SELECT r.route_id,ri.route_name,r.sell_station_id ,r.end_station_id, s.station_name sell_station_name,si.station_name,r.section_id,ss.section_name,r.end_station_mileage" _
        & " FROM back_route_section_lst r,back_route_info ri,station_info s,station_info si,section_info ss " _
        & " WHERE  r.route_id= " & TransFieldValueToString(m_szRouteID) _
        & " AND r.route_id = ri.route_id AND r.end_station_id =si.station_id AND r.sell_station_id = s.station_id AND r.section_id = ss.section_id " _
        & "  AND sell_station_id in (SELECT sell_station_id FROM back_route_section_lst " _
          & " WHERE  end_station_mileage= ( SELECT MAX(end_station_mileage) " _
          & " FROM back_route_section_lst WHERE  route_id= '" & m_szRouteID & "')) " _
        & szTemp _
        & " ORDER BY r.sell_station_id,r.section_serial"
    Set rsTemp = oDb.Execute(szSql)
    
'
'
'    szSql = "SELECT * FROM back_route_section_lst WHERE route_id= '" & m_szRouteID & "'" _
'          & " AND sell_station_id in (SELECT sell_station_id FROM back_route_section_lst " _
'          & " WHERE end_station_mileage= ( SELECT MAX(end_station_mileage) " _
'          & " FROM back_route_section_lst WHERE  route_id= '" & m_szRouteID & "'))" _
'          & " ORDER BY back_route_section_lst.section_serial"
'    Set rsTemp = oDb.Execute(szSql)
       
    If rsTemp.RecordCount = 0 Then Exit Function
    ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 10)
    For i = 1 To rsTemp.RecordCount
        aszTemp(i, 1) = FormatDbValue(rsTemp!route_id)
        aszTemp(i, 2) = FormatDbValue(rsTemp!route_name)
        aszTemp(i, 3) = FormatDbValue(rsTemp!sell_station_id)
        aszTemp(i, 4) = FormatDbValue(rsTemp!sell_station_name)
        aszTemp(i, 5) = FormatDbValue(rsTemp!end_station_id)
        aszTemp(i, 6) = FormatDbValue(rsTemp!station_name)
        aszTemp(i, 7) = FormatDbValue(rsTemp!section_id)
        aszTemp(i, 8) = FormatDbValue(rsTemp!section_name)
        aszTemp(i, 9) = FormatDbValue(rsTemp!end_station_mileage)
        rsTemp.MoveNext
    Next i
    GetAllSectionInfo = aszTemp
End Function



