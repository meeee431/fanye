VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Section"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**********************************************************
'* Source File Name:Section.cls
'* Project Name:STBase.vbp
'* Engineer:
'* Data Generated:
'* Last Revision Date:
'* Brief Description:路段对象
'* Relational Document:
'**********************************************************
Option Explicit
'出错
Const cszSectionCordInPut = "KK"
Public Enum EErrSection
    ERR_SectionExist = ERR_Section + ERR_AddDuplicate '56401路段已存在不能新增
    ERR_SectionNoDelete = ERR_Section + ERR_DeleteChildExist '56407路段已安排线路不能删除
    ERR_SectionNotRoadLevelOrStation = ERR_Section + ERR_EditChildExist '56405路段无该公路等级或站点
    ERR_SectionNotAvailable = ERR_Section + ERR_NotAvailable '56412路段对象处于无状态
    ERR_SectionNotAddStatus = ERR_Section + ERR_NotAddObj '56416路段对象不在新增状态
    ERR_SectionNotExist = ERR_Section + 21 '56421路段无存在
    ERR_SectionIDNotNull = ERR_Section + 22 '56422路段代码不能为空
    ERR_SectionStartNoEnd = ERR_Section + 23 '56423路段起点站不能同终点站
End Enum

'权限
Public Enum ERightSection
    RIGHT_SectionManagement = ERR_Section + cnMidRightBegin + cnMidRightStep * 1 '56506路段管理
End Enum

Private m_oActiveUser As ActiveUser 'P1所有的类都有的私有类变量
Private m_nObjectStatus As EObjectStatus 'P2所有的实体类都有的私有类变量
Private m_szSectionID As String 'P3路段代码
Private m_szSectionName As String 'P4路段名称
Private m_szAreaCode As String 'P5地区代码
Private m_szRoadLevelCode As String 'P6公路等级
Private m_stMileage As Double 'P7里程数
Private m_szBeginStationCode As String 'P8起点站代码
Private m_szEndStationCode As String 'P9终点站代码
Private m_szBeginStationName As String 'P12起点站代码
Private m_szEndStationName As String 'P13终点站代码
Private m_nSectionSerialNo As Integer 'P10路段序号
Private m_sgOldSectionMileage As Double '路段旧的里程数
Private m_szRoadLevelName As String 'P11公路等级名称
Private m_szAreaName As String 'P12地区名称





Private m_OldszSectionID As String 'P3路段代码
Private m_OldszSectionName As String 'P4路段名称
Private m_OldszAreaCode As String 'P5地区代码
Private m_oldszRoadLevelCode As String 'P6公路等级
Private m_oldstMileage As Double 'P7里程数
Private m_oldszBeginStationCode As String 'P8起点站代码
Private m_oldszEndStationCode As String 'P9终点站代码
Private m_oldszBeginStationName As String 'P12起点站代码
Private m_oldszEndStationName As String 'P13终点站代码
Private m_oldnSectionSerialNo As Integer 'P10路段序号
Private m_oldsgOldSectionMileage As Double '路段旧的里程数
Private m_oldszRoadLevelName As String 'P11公路等级名称
Private m_oldszAreaName As String 'P12地区名称

'F1
Public Function GetAllRoute() As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    Dim szSql As String
    AssertObjIsValid
    szSql = "SELECT DISTINCT route_id FROM Route_section_lst  WHERE section_id='" & m_szSectionID & "'"
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    ReDim szaTemp(1 To nCount) As String
    For i = 1 To nCount
        szaTemp(i) = FormatDbValue(rsTemp!route_id)
        rsTemp.MoveNext
    Next
    GetAllRoute = szaTemp
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

'F2
Public Function GetPassCharge(VehicleModel As String) As Double
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT pass_charge FROM Bus_passcharge_info WHERE section_id='" & m_szSectionID & "'" & IIf(VehicleModel = "", "", " and vehicle_type_code='" & VehicleModel & "'")
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 0 Then
        GetPassCharge = FormatDbValue(rsTemp!pass_charge)
    Else
        GetPassCharge = 0
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function
'F3
Public Function GetAllRouteName() As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    Dim szSql As String
    AssertObjIsValid
    szSql = "SELECT  route_name FROM Route_section_lst tbc ,route_info tbr  WHERE section_id='" & m_szSectionID & "' AND tbr.route_id =tbc.route_id"
    
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    ReDim szaTemp(1 To nCount) As String
    For i = 1 To nCount
        szaTemp(i) = FormatDbValue(rsTemp!route_name)
        rsTemp.MoveNext
    Next
    GetAllRouteName = szaTemp
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
Public Property Let SectionID(ByVal vData As String)
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    AssertActiveUserValid m_oActiveUser, ERR_Section
    '查询是否是新增状态
    If m_nObjectStatus <> ST_AddObj Then RaiseError ERR_SectionNotAddStatus
    '输入的主键是否为空
    If vData = "" Then RaiseError ERR_SectionIDNotNull
    '查询该对象是否已有记录
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT section_id FROM section_info WHERE section_id='" & RTrim(vData) & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 1 Then RaiseError ERR_SectionExist
    m_szSectionID = RTrim(vData)
    Set rsTemp = Nothing
    Set oDb = Nothing
End Property
'P3
Public Property Get SectionID() As String
    SectionID = m_szSectionID
End Property
'P4
Public Property Let SectionName(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szSectionName = vData
End Property
'P4
Public Property Get SectionName() As String
    SectionName = m_szSectionName
End Property
'P5
Public Property Let AreaCode(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szAreaCode = vData
End Property
'P5
Public Property Get AreaCode() As String
    AreaCode = m_szAreaCode
End Property

'P6
Public Property Let RoadLevelCode(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szRoadLevelCode = vData
End Property
'P6
Public Property Get RoadLevelCode() As String
    RoadLevelCode = m_szRoadLevelCode
End Property
'P7
Public Property Let Mileage(ByVal vData As Double)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_stMileage = vData
End Property
'P7
Public Property Get Mileage() As Double
    Mileage = m_stMileage
End Property
'P8
Public Property Let BeginStationCode(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szBeginStationCode = vData
End Property
'P8
Public Property Get BeginStationCode() As String
    BeginStationCode = m_szBeginStationCode
End Property

'P9
Public Property Let EndStationCode(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szEndStationCode = Trim(vData)
End Property
'P9
Public Property Get EndStationCode() As String
    EndStationCode = m_szEndStationCode
End Property

'P10
Public Property Let SectionSerialNo(ByVal vData As Integer)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_nSectionSerialNo = Trim(vData)
End Property
'P10
Public Property Get SectionSerialNo() As Integer
    SectionSerialNo = m_nSectionSerialNo
End Property
'P12
Public Property Get BeginStationName() As String
    BeginStationName = m_szBeginStationName
End Property
'P13
Public Property Get EndStationName() As String
    EndStationName = m_szEndStationName
End Property

Public Property Get RoadLevelName() As String
    RoadLevelName = m_szRoadLevelName
End Property

Public Property Get AreaName() As String
    AreaName = m_szAreaName
End Property
'S1
Public Sub Init(poAUser As ActiveUser)
    Set m_oActiveUser = poAUser
    'AssertHaveRight m_oActiveUser, RIGHT_SectionManagement
End Sub

'S2类刚建立时处于不可用状态,此时类不能使用
Private Sub Class_Initialize()
    m_nObjectStatus = ST_NotAvailable
End Sub

'S3测试对象的状态是否有效（无效状态[ST_NotAvailable]和新增状态[ST_AddObj]则出错）
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then RaiseError ERR_SectionNotAvailable
End Sub
'S4
Private Sub AssertObjIsValid() '测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
    AssertActiveUserValid m_oActiveUser, ERR_Section
    AssertStatusAvailable
End Sub
'S5
Public Sub Identify(SectionID As String)
    AssertActiveUserValid m_oActiveUser, ERR_Section
    m_szSectionID = RTrim(SectionID)
    RefreshMemoryInfo
    m_nObjectStatus = ST_NormalObj
End Sub
'S6调用新增方法后，对象处于新增状态。
Public Sub AddNew()
    AssertHaveRight m_oActiveUser, RIGHT_SectionManagement
    AssertActiveUserValid m_oActiveUser, ERR_Section
    m_nObjectStatus = ST_AddObj
End Sub
'S7
Public Sub Delete() '调用删除方法后，对象处于不可用状态
    AssertHaveRight m_oActiveUser, RIGHT_SectionManagement
    AssertObjIsValid
    DeleteObject
    m_nObjectStatus = ST_NotAvailable
End Sub
'S8
Public Sub Update()
    AssertHaveRight m_oActiveUser, RIGHT_SectionManagement
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
    Set rsTemp = oDb.Execute("SELECT tbs.*,tbr.road_rank_name,tba.area_name FROM section_info tbs,road_level_code tbr,area_code tba WHERE section_id='" & m_szSectionID & "' AND tbs.road_rank_code=tbr.road_rank_code AND tba.area_code=tbs.area_code")
    If rsTemp.RecordCount = 0 Then RaiseError ERR_SectionNotExist
    m_szAreaCode = FormatDbValue(rsTemp!area_code)
    m_szAreaName = FormatDbValue(rsTemp!area_name)
    m_szRoadLevelName = FormatDbValue(rsTemp!road_rank_name)
    m_szBeginStationCode = FormatDbValue(rsTemp!start_station)
    m_szEndStationCode = FormatDbValue(rsTemp!end_station)
    m_szRoadLevelCode = FormatDbValue(rsTemp!road_rank_code)
    m_szSectionName = FormatDbValue(rsTemp!section_name)
    m_nSectionSerialNo = FormatDbValue(rsTemp!path_no)
    m_stMileage = FormatDbValue(rsTemp!Mileage)
    m_sgOldSectionMileage = FormatDbValue(rsTemp!Mileage)
    m_szBeginStationName = Trim(oDb.Execute("SELECT station_name FROM station_info WHERE station_id='" & m_szBeginStationCode & "'")!station_name)
    m_szEndStationName = Trim(oDb.Execute("SELECT station_name FROM station_info WHERE station_id='" & m_szEndStationCode & "'")!station_name)
    
    
'''''''''''''''''''''''''''''''
m_OldszSectionID = m_szSectionID
m_OldszSectionName = m_szSectionName
m_OldszAreaCode = m_szAreaCode
m_oldszRoadLevelCode = m_szRoadLevelCode
m_oldstMileage = m_stMileage
m_oldszBeginStationCode = m_szBeginStationCode
m_oldszEndStationCode = m_szEndStationCode
m_oldszBeginStationName = m_szBeginStationName
m_oldszEndStationName = m_szEndStationName
m_oldnSectionSerialNo = m_nSectionSerialNo
m_oldsgOldSectionMileage = m_sgOldSectionMileage
m_oldszRoadLevelName = m_oldszRoadLevelName
m_oldszAreaName = m_szAreaName
''''''''''''''''''''''''''''''''
    
    
    
    Set rsTemp = Nothing
    Set oDb = Nothing
End Sub

'S10
Private Sub AddObject()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim szTempSql As String
    Dim szSqlField As String
    Dim rsTemp As Recordset
    Dim szSqlContext As String
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    If Trim(m_szSectionID) = "" Then RaiseError ERR_SectionIDNotNull
    If m_szBeginStationCode = m_szEndStationCode Then RaiseError ERR_SectionStartNoEnd
    Set rsTemp = oDb.Execute("SELECT start_station,end_station,path_no FROM section_info WHERE end_station='" & m_szEndStationCode & "' AND start_station='" & m_szBeginStationCode & "' ORDER BY path_no DESC")
    
    If rsTemp.RecordCount <> 0 Then
    If FormatDbValue(rsTemp!start_station) = Trim(m_szBeginStationCode) And FormatDbValue(rsTemp!end_station) = Trim(m_szEndStationCode) Then
        m_nSectionSerialNo = FormatDbValue(rsTemp!path_no) + 1
    End If
    End If
    szTempSql = "INSERT section_info ("
    szSqlField = szSqlField & "section_id,"
    szSqlField = szSqlField & "area_code,"
    szSqlField = szSqlField & "start_station,"
    szSqlField = szSqlField & "end_station,"
    szSqlField = szSqlField & "road_rank_code,"
    szSqlField = szSqlField & "section_name,"
    szSqlField = szSqlField & "path_no,"
    szSqlField = szSqlField & "Mileage)"
    szSqlContext = " VALUES ('" & m_szSectionID & "',"
    szSqlContext = szSqlContext & " '" & m_szAreaCode & "',"
    szSqlContext = szSqlContext & " '" & m_szBeginStationCode & "',"
    szSqlContext = szSqlContext & " '" & m_szEndStationCode & "',"
    szSqlContext = szSqlContext & " '" & m_szRoadLevelCode & "',"
    szSqlContext = szSqlContext & " '" & m_szSectionName & "',"
    szSqlContext = szSqlContext & " " & m_nSectionSerialNo & ","
    szSqlContext = szSqlContext & " " & m_stMileage & ")"
    szTempSql = szTempSql & szSqlField & szSqlContext
    
    oDb.Execute (szTempSql)
    Set rsTemp = Nothing
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertAddObjectError ERR_Section, oDb
End Sub
'S11
Private Sub DeleteObject()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    
    
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    oDb.Execute ("DELETE section_info WHERE section_id='" & m_szSectionID & "'")
    Set oDb = Nothing
    
    WriteOperateLog m_oActiveUser, RIGHT_SectionManagement, "路段ID为[" & m_szSectionID & "]被删除"
    Exit Sub
ErrHandle:
    AssertDeleteObjectError ERR_Section, oDb
    
End Sub

'S12
Private Sub UpdateToDB()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim szTempSql As String
    Dim szSqlField As String
    Dim szSqlWhere As String
'    Dim szaRouteID() As String
    Dim nCount As Integer
    Dim i As Integer
    Dim rsTemp As Recordset
    Dim sgChangeMileage As Double
    ''''''
    Dim szlog As String
    ''''''
    
    
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szTempSql = "UPDATE section_info SET "
    szSqlField = szSqlField & "area_code='" & m_szAreaCode & "',"
    szSqlField = szSqlField & "start_station='" & m_szBeginStationCode & "',"
    szSqlField = szSqlField & "end_station='" & m_szEndStationCode & "',"
    szSqlField = szSqlField & "road_rank_code='" & m_szRoadLevelCode & "',"
    szSqlField = szSqlField & "path_no=" & m_nSectionSerialNo & ","
    szSqlField = szSqlField & "Mileage=" & m_stMileage & ","
    szSqlField = szSqlField & "section_name='" & m_szSectionName & "'"
    szSqlWhere = " WHERE section_id='" & m_szSectionID & "'"
    szTempSql = szTempSql & szSqlField & szSqlWhere
    oDb.Execute (szTempSql)
  szlog = "路段" & m_szSectionID
'  If m_OldszSectionID <> m_szSectionID Then
'     szLog = szLog & "路段ID由"
'  End If
  If m_OldszSectionName <> m_szSectionName Then
     szlog = szlog & "路段名由[" & m_OldszSectionName & "]改为" & m_szSectionName
  End If
  If m_OldszAreaCode <> m_szAreaCode Then
    szlog = szlog & "地区ID由[" & m_OldszAreaCode & "]改为" & m_szAreaCode
  End If
  If m_oldszAreaName <> m_szAreaName Then
   szlog = szlog & "地区名由[" & m_oldszAreaName & "]改为" & m_szAreaName
 End If
 
 If m_oldszRoadLevelCode <> m_szRoadLevelCode Then
    szlog = szlog & "公路等级ID由[" & m_oldszRoadLevelCode & "]改为" & m_szRoadLevelCode
 End If
 If m_oldszRoadLevelName <> m_szRoadLevelName Then
   szlog = szlog & "公路等级名由[" & m_oldszRoadLevelName & "]改为" & m_szRoadLevelName
 End If
 If m_oldstMileage <> m_stMileage Then
   szlog = szlog & "里程由[" & m_oldstMileage & "]改为" & m_stMileage
 End If
 If m_oldszBeginStationCode <> m_szBeginStationCode Then
   szlog = szlog & "起点站ID由[" & m_oldszBeginStationCode & "]改为" & m_szBeginStationCode
 End If
 If m_oldszEndStationCode <> m_szEndStationCode Then
     szlog = szlog & "终站地区ID由[" & m_oldszEndStationCode & "]改为" & m_szEndStationCode
 End If
 
 If m_oldszBeginStationName <> m_szBeginStationName Then
   szlog = szlog & "起站名由[" & m_oldszBeginStationName & "]改为" & m_szBeginStationName
 End If
 If m_oldszEndStationName <> m_szEndStationName Then
   szlog = szlog & "终站名由[" & m_oldszEndStationName & "]改为" & m_szEndStationName
 End If

szlog = GetUnicodeBySize(szlog, 255)
 
WriteOperateLog m_oActiveUser, RIGHT_SectionManagement, szlog
    
    
    '因为里程数的修改将修改使用该路段的所有线路和票价表中的站点里程数
    sgChangeMileage = m_stMileage - m_sgOldSectionMileage
    If sgChangeMileage = 0 Then Exit Sub
'    szaRouteID = Me.GetAllRoute
    szTempSql = "SELECT DISTINCT sell_station_id , route_id FROM Route_section_lst  WHERE section_id='" & m_szSectionID & "'"
    Set rsTemp = oDb.Execute(szTempSql)
    nCount = rsTemp.RecordCount
    For i = 1 To nCount
        '修改线路路段中的里程数Route_section_lst
        szTempSql = "UPDATE Route_section_lst " _
            & " SET end_station_mileage=end_station_mileage+ " & Str(sgChangeMileage) _
            & " WHERE route_id='" & FormatDbValue(rsTemp!route_id) & "' " _
            & " AND section_serial>=( " _
            & " SELECT min(section_serial) FROM Route_section_lst " _
            & " WHERE section_id='" & m_szSectionID & "' AND route_id='" & FormatDbValue(rsTemp!route_id) & "' AND sell_station_id = '" & FormatDbValue(rsTemp!sell_station_id) & "'" _
            & " ) AND sell_station_id = '" & FormatDbValue(rsTemp!sell_station_id) & "'"
        oDb.Execute szTempSql
        '修改线路的里程数route_info
        szTempSql = "UPDATE route_info SET mileage=mileage+" & Str(sgChangeMileage) & " WHERE route_id='" & FormatDbValue(rsTemp!route_id) & "'"
        oDb.Execute szTempSql
        '修改线路票价中的里程数price_table_lst
        'szTempSql = "UPDATE price_table_lst SET mileage=mileage+" & Str(sgChangeMileage) & " WHERE route_id='" & szaRouteID(i) & "' AND station_serial_no>=(SELECT DISTINCT station_serial_no FROM price_table_lst WHERE station_id='" & m_szEndStationCode & "' AND route_id='" & szaRouteID(i) & "')"
        'oDb.Execute szTempSql
        rsTemp.MoveNext
    Next
    Set rsTemp = Nothing
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertUpdateObjectError ERR_Section, oDb

End Sub

'S13编辑路段通行费
Public Sub ModifyPassCharge(tTransitCharge As TTransitChargeInfo)
    Dim oDb As New RTConnection
    Dim szTempSql As String
    Dim szSqlField As String
    Dim szSqlWhere As String
    Dim nCount As String
    Dim i As Integer
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
   On Error GoTo Here
   oDb.BeginTrans
   
        szTempSql = "SELECT * FROM Bus_passcharge_info WHERE section_id='" & tTransitCharge.szSection & "' AND vehicle_type_code='" & tTransitCharge.szVehicleType & "'and seat_type_id='" & tTransitCharge.szSeatType & "' "
        If oDb.Execute(szTempSql).RecordCount = 0 Then
            AddPassCharge tTransitCharge
        Else
            szSqlField = ""
            szSqlWhere = ""
            szTempSql = "UPDATE Bus_passcharge_info SET "
            szSqlField = szSqlField & "pass_charge=" & Str(tTransitCharge.szPassCharge) & ","
            szSqlField = szSqlField & "seat_type_id='" & tTransitCharge.szSeatType & "',"
            szSqlField = szSqlField & "annotation='" & tTransitCharge.szAnnotation & "'"
            szSqlWhere = " WHERE section_id='" & tTransitCharge.szSection & "' AND "
            szSqlWhere = szSqlWhere & " vehicle_type_code='" & tTransitCharge.szVehicleType & "' and seat_type_id='" & tTransitCharge.szSeatType & "'"
            szTempSql = szTempSql & szSqlField & szSqlWhere
            oDb.Execute (szTempSql)
        End If
   oDb.CommitTrans
    Set oDb = Nothing
    Exit Sub
Here:
   oDb.RollbackTrans
   Set oDb = Nothing
   RaiseError err.Number, , "通行费输入有误"
End Sub
'S14删除路段通行费
Public Sub DeletePassCharge(VehicleModel As String, szSeatType As String)
    Dim oDb As New RTConnection
    AssertObjIsValid
    AssertActiveUserValid m_oActiveUser, RIGHT_SectionManagement
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    oDb.Execute ("DELETE Bus_passcharge_info WHERE section_id='" & m_szSectionID & "' AND vehicle_type_code='" & VehicleModel & "' and seat_type_id ='" & szSeatType & "'")
    WriteOperateLog m_oActiveUser, RIGHT_SectionManagement, Now & "路段[" & " & m_szSectionID & " & "车型[" & " & VehicleModel & " & "]座位类型" & " & szSeatType & " & "路段费被删除"
    Set oDb = Nothing
End Sub

'S15新增路段通行费
Public Sub AddPassCharge(tTransitCharge As TTransitChargeInfo)
    
    Dim oDb As New RTConnection
    Dim szTempSql As String
    Dim szSqlField As String
    Dim szSqlWhere As String
    Dim rsTemp As Recordset
    Dim szSqlContext As String
    Dim i As Integer
    Dim nCount As String
    Dim tTransitChargeTemp(1 To 1) As TTransitChargeInfo
    tTransitChargeTemp(1) = tTransitCharge
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
   On Error GoTo Here
    oDb.BeginTrans
        szTempSql = "SELECT * FROM Bus_passcharge_info"
        szSqlWhere = " WHERE section_id='" & tTransitCharge.szSection & "' AND "
        szSqlWhere = szSqlWhere & " vehicle_type_code='" & tTransitCharge.szVehicleType & "' and seat_type_id='" & tTransitCharge.szSeatType & "'"
        szTempSql = szTempSql & szSqlWhere
        
        Set rsTemp = oDb.Execute(szTempSql)
        '-----------------------------
        '如果无该通行费则新增通行费
        If rsTemp.RecordCount <> 0 Then ModifyPassCharge tTransitCharge: Exit Sub
        szSqlContext = ""
        szSqlField = ""
        szTempSql = "INSERT Bus_passcharge_info ("
        szSqlField = szSqlField & "section_id,"
        szSqlField = szSqlField & "vehicle_type_code,"
        szSqlField = szSqlField & "seat_type_id,"
        szSqlField = szSqlField & "pass_charge,"
        szSqlField = szSqlField & "annotation)"
        szSqlContext = " VALUES ('" & tTransitCharge.szSection & "',"
        szSqlContext = szSqlContext & " '" & tTransitCharge.szVehicleType & "',"
        szSqlContext = szSqlContext & " '" & tTransitCharge.szSeatType & "',"
        szSqlContext = szSqlContext & " " & Str(tTransitCharge.szPassCharge) & ","
        szSqlContext = szSqlContext & " '" & tTransitCharge.szAnnotation & "')"
        szTempSql = szTempSql & szSqlField & szSqlContext
        oDb.Execute (szTempSql)
    
    oDb.CommitTrans
    Set rsTemp = Nothing
    Set oDb = Nothing
    Exit Sub
Here:
    oDb.RollbackTrans
    Set rsTemp = Nothing
    Set oDb = Nothing
    RaiseError err.Number, , "通行费输入有误"
  End Sub
'获得已某站点为起点站和已某站点为终点站的路段
Public Function GetSESection(StartStation As String, EndStation As String) As String()
    Dim oDb As New RTConnection
    Dim szaTemp() As String
    Dim nCount As Integer, i As Integer
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute("SELECT section_id FROM section_info WHERE start_station='" & StartStation & "' AND end_station='" & EndStation & "'")
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    ReDim szaTemp(1 To nCount) As String
    For i = 1 To nCount
        szaTemp(i) = FormatDbValue(rsTemp!section_id)
        rsTemp.MoveNext
    Next
    GetSESection = szaTemp
    End If
'    Set rsTemp = Nothing
    Set rsTemp = Nothing
End Function
'由路段号得到路段起点站,终点站
'auther=ljw
Public Function GetSectionStatAndEnd(ByVal szSectionID As String) As String()
  Dim oDb As New RTConnection
    Dim szaTemp() As String
    Dim nCount As Integer, i As Integer
    Dim rsTemp As Recordset
    Dim rsTemp2 As Recordset
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute("SELECT start_station as start_station1, end_station  as start_station2 FROM section_info WHERE section_id='" & szSectionID & "'")
    
    
    nCount = rsTemp.RecordCount
    
    ReDim szaTemp(1 To 2, 1 To 2) As String
    Dim szStation As String
          
     If nCount <> 0 Then
        For i = 1 To 2
        szaTemp(i, 1) = Trim(rsTemp.Fields("start_station" & i))
        Set rsTemp2 = oDb.Execute("select station_input_code from  station_info WHERE station_id='" & Trim(szaTemp(i, 1)) & "'")
          
          If rsTemp2.RecordCount <> 0 Then
             szaTemp(i, 2) = FormatDbValue(rsTemp2!station_input_code)
          End If
        
       ' szaTemp(i, 1) = formatdbvalue(rstemp!end_station)
        
        Next
        
        GetSectionStatAndEnd = szaTemp
        
    End If
    Set rsTemp = Nothing
End Function
Public Function GetPassChargeEX(tTTransitChargeInfo As TTransitChargeInfo) As TTransitChargeInfo()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim tTTransitChargeInfoTemp() As TTransitChargeInfo
    Dim i As Integer
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT * FROM Bus_passcharge_info WHERE section_id='" & tTTransitChargeInfo.szSection & "'" & IIf(tTTransitChargeInfo.szVehicleType = "", "", " and vehicle_type_code='" & tTTransitChargeInfo.szVehicleType & "'") & IIf(tTTransitChargeInfo.szSeatType = "", "", " and seat_type_id='" & tTTransitChargeInfo.szSeatType & "'")
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 0 Then
       ReDim tTTransitChargeInfoTemp(1 To rsTemp.RecordCount)
       For i = 1 To rsTemp.RecordCount
         tTTransitChargeInfoTemp(i).szSection = FormatDbValue(rsTemp!section_id)
         tTTransitChargeInfoTemp(i).szSeatType = FormatDbValue(rsTemp!seat_type_id)
         tTTransitChargeInfoTemp(i).szPassCharge = FormatDbValue(rsTemp!pass_charge)
         tTTransitChargeInfoTemp(i).szVehicleType = FormatDbValue(rsTemp!vehicle_type_code)
         tTTransitChargeInfoTemp(i).szAnnotation = FormatDbValue(rsTemp!Annotation)
         rsTemp.MoveNext
       Next
    End If
    GetPassChargeEX = tTTransitChargeInfoTemp
    Set rsTemp = Nothing
    Set oDb = Nothing

End Function
Private Function CreateSectionID()

End Function
Public Function FindSectionIDORCreateSectionID() As Boolean
Dim i As Integer
Dim j As Integer
Dim szSql As String
Dim oDb As New RTConnection
Dim rsTemp As Recordset
Dim szSectiomID As String
Dim szSectionName As String
Dim szBeginStationCode As String
Dim szBeginStationName As String
Dim szEndStationCode As String
Dim szEndStationName As String
oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
szSql = "SELECT section_id FROM section_info where start_station='" & m_szBeginStationCode & "'  and end_station='" & m_szEndStationCode & "'"
Set rsTemp = oDb.Execute(szSql)

If rsTemp.RecordCount <> 0 Then
    m_szSectionID = FormatDbValue(rsTemp!section_id)
    m_szSectionName = FormatDbValue(rsTemp!section_name)
    FindSectionIDORCreateSectionID = True
    Exit Function
End If

szSql = "SELECT station_input_code,station_name,station_id,area_code FROM Station_info where station_id='" & m_szBeginStationCode & "' or station_id='" & m_szEndStationCode & "'"
Set rsTemp = oDb.Execute(szSql)
If rsTemp.RecordCount < 2 Then RaiseError err.Number

For i = 1 To 2
If FormatDbValue(rsTemp!station_id) = Trim(m_szBeginStationCode) Then
   szBeginStationCode = IIf(FormatDbValue(rsTemp!station_input_code) = "", cszSectionCordInPut, Left(FormatDbValue(rsTemp!station_input_code), 2))
   szBeginStationName = Left(FormatDbValue(rsTemp!station_name), 2)
Else
   szEndStationCode = IIf(FormatDbValue(rsTemp!station_input_code) = "", cszSectionCordInPut, Left(FormatDbValue(rsTemp!station_input_code), 2))
   szEndStationName = Left(FormatDbValue(rsTemp!station_name), 2)
   m_szAreaCode = FormatDbValue(rsTemp!area_code)
End If
rsTemp.MoveNext
Next
i = 0
   szSectiomID = szBeginStationCode & szEndStationCode
   szSectionName = szBeginStationName & szEndStationName
here2:
   szSql = "SELECT * FROM section_info  WHERE section_id='" & szSectiomID & "'"
   Set rsTemp = oDb.Execute(szSql)
   If rsTemp.RecordCount = 0 Then
   m_szSectionName = szSectionName
   m_szSectionID = szSectiomID
   Exit Function
   Else
   i = i + 1
   szSectiomID = szSectiomID & i
   szSectiomID = Left(szSectiomID, 1) & Trim(Right(szSectiomID, 3))
   GoTo here2
   End If
End Function

