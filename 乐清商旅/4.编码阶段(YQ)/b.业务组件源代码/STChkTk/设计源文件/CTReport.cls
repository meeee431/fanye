VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'检票报表


Option Explicit

'----------------------
'各类特有的变量
Private m_szFunctionID As String
Private m_szComponentID As String
Private m_szFunctionGroup As String
Private m_szFunctionName As String
Private m_eIfWriteLog As EIsWritelog
Private m_szAnnotation As String
'---------------------
'所有的实体类都有的私有类变量<<
Private m_nObjectStatus As EObjectStatus

'所有的类都有的私有类变量
Private m_oActiveUser As ActiveUser

Public Sub Init(poAUser As ActiveUser)
    Set SelfUser = poAUser
End Sub
Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
    m_oActiveUser.IAmLive
End Property

Public Property Set SelfUser(NewValue As ActiveUser)
    If m_nObjectStatus = ST_AddObj Then
        ShowError ERR_AddObj
    ElseIf m_nObjectStatus = ST_EditObj Then
        ShowError ERR_EditObj
    ElseIf m_nObjectStatus = ST_NormalObj Then
        ShowError ERR_NormalObj
    Else
        Set m_oActiveUser = NewValue
    End If
    m_oActiveUser.IAmLive
End Property

Public Function GetCheckTicketReport(BusDate As Date) As Recordset
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    
    Dim oDb As New RTConnection
    Dim rsTemp As New Recordset
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "Exec CheckQuery '" & Format(BusDate, "YYYY-MM-DD") & "'"
    Set rsTemp = oDb.Execute(szSql)
    Set GetCheckTicketReport = rsTemp
    
''清临时表
'    szSql = ""
'    szSql = "IF EXISTS(SELECT * FROM tempdb.dbo.sysobjects WHERE name LIKE '#SNTemp2%')" & vbCrLf
'    szSql = szSql & "DROP TABLE #SNTemp2" & vbCrLf
'    szSql = szSql & "IF EXISTS(SELECT * FROM tempdb.dbo.sysobjects WHERE name LIKE '#SNTemp%')" & vbCrLf
'    szSql = szSql & "DROP TABLE #SNTemp" & vbCrLf
'    oDb.Execute szSql
'
'    szSql = "SELECT b.bus_date,b.bus_id,b.bus_serial_no," & vbCrLf
'    szSql = szSql & "COUNT(s.ticket_id) AS 'sale_quantity'," & vbCrLf
'    szSql = szSql & "0 AS 'change_check_quantity'," & vbCrLf
'    szSql = szSql & "COUNT(c.ticket_id) AS 'check_quantity'," & vbCrLf
'    szSql = szSql & "SUM(c.ticket_price)AS 'check_price'" & vbCrLf
'    szSql = szSql & "INTO #SNTemp" & vbCrLf
'    szSql = szSql & "FROM check_ticket_lst c,sale_ticket_lst s,Bus_check_lst b" & vbCrLf
'    szSql = szSql & "WHERE s.ticket_id*=c.ticket_id AND" & vbCrLf
'    szSql = szSql & "  s.bus_date=b.bus_date AND s.bus_id=b.bus_id AND" & vbCrLf
'    szSql = szSql & "  b.bus_date*=c.bus_date AND" & vbCrLf
'    szSql = szSql & "  b.bus_id*=c.bus_id AND" & vbCrLf
'    szSql = szSql & "  b.bus_serial_no*=c.bus_serial_no AND" & vbCrLf
'    szSql = szSql & "  s.bus_date='" & Format(BusDate, "YYYY-MM-DD") & "'" & vbCrLf
'    szSql = szSql & "GROUP BY b.bus_date,b.bus_id,b.bus_serial_no" & vbCrLf
'    szSql = szSql & "ORDER BY b.bus_date,b.bus_id,b.bus_serial_no" & vbCrLf
''改乘
'    szSql = szSql & "SELECT bus_date,bus_id,bus_serial_no," & vbCrLf
'    szSql = szSql & "   COUNT(ticket_id) AS 'change_quantity'," & vbCrLf
'    szSql = szSql & "   SUM(ticket_price) AS 'change_price'" & vbCrLf
'    szSql = szSql & "INTO #SNTemp2" & vbCrLf
'    szSql = szSql & "FROM check_ticket_lst" & vbCrLf
'    szSql = szSql & "WHERE bus_date='" & Format(BusDate, "YYYY-MM-DD") & "' AND status=2" & vbCrLf
'    szSql = szSql & "GROUP BY bus_date,bus_id,bus_serial_no" & vbCrLf
'    szSql = szSql & "ORDER BY bus_date,bus_id,bus_serial_no" & vbCrLf
''将改乘数加到检票数
'    szSql = szSql & "UPDATE #SNTemp SET change_check_quantity=change_quantity,check_quantity=check_quantity+change_quantity,check_price=check_price+change_price" & vbCrLf
'    szSql = szSql & "FROM #SNTemp t1,#SNTemp2 t2" & vbCrLf
'    szSql = szSql & "WHERE t1.bus_date=t2.bus_date AND" & vbCrLf
'    szSql = szSql & "  t1.bus_id=t2.bus_id AND" & vbCrLf
'    szSql = szSql & "  t1.bus_serial_no = t2.bus_serial_no" & vbCrLf
'    oDb.Execute szSql
'
'    szSql = "SELECT bus_id,bus_serial_no,sale_quantity,change_check_quantity,check_quantity,check_price FROM #SNTemp"
'    Set rsTemp = oDb.Execute(szSql, , -1)
'    Set GetCheckTicketReport = rsTemp
End Function

Public Function GetScrollCheckTicketReport(BusDate As Date, Optional BusID As String = "") As Recordset
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    
    Dim oDb As New RTConnection
    Dim rsTemp As New Recordset
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "Exec ScrollCheckQuery '" & Format(BusDate, "YYYY-MM-DD") & "','" & BusID & "'"
    Set rsTemp = oDb.Execute(szSql)
    Set GetScrollCheckTicketReport = rsTemp
''清临时表
'    szSql = ""
'    szSql = "IF EXISTS(SELECT * FROM tempdb.dbo.sysobjects WHERE name LIKE '#SNTemp2%')" & vbCrLf
'    szSql = szSql & "DROP TABLE #SNTemp2" & vbCrLf
'    szSql = szSql & "IF EXISTS(SELECT * FROM tempdb.dbo.sysobjects WHERE name LIKE '#SNTemp%')" & vbCrLf
'    szSql = szSql & "DROP TABLE #SNTemp" & vbCrLf
'    oDb.Execute szSql
'
'    szSql = "SELECT b.bus_date,b.bus_id,b.bus_serial_no," & vbCrLf
'    szSql = szSql & "v.license_tag_no,t.vehicle_type_short_name,o.owner_name,p.transport_company_short_name," & vbCrLf
'    szSql = szSql & "COUNT(s.ticket_id) AS 'sale_quantity'," & vbCrLf
'    szSql = szSql & "SUM(s.ticket_price) AS 'sale_price'," & vbCrLf
'    szSql = szSql & "COUNT(c.ticket_id) AS 'check_quantity'," & vbCrLf
'    szSql = szSql & "SUM(c.ticket_price)AS 'check_price'" & vbCrLf
'    szSql = szSql & "INTO #SNTemp " & vbCrLf
'    szSql = szSql & "FROM check_ticket_lst c,sale_ticket_lst s," & vbCrLf
'    szSql = szSql & "     Bus_check_lst b,transport_vehicle_info v,transport_company_info p,owner_info o,vehicle_type_code t" & vbCrLf
'    szSql = szSql & "WHERE s.ticket_id*=c.ticket_id AND" & vbCrLf
'    szSql = szSql & "  s.bus_date=b.bus_date AND s.bus_id=b.bus_id AND" & vbCrLf
'    szSql = szSql & "  b.bus_date*=c.bus_date AND" & vbCrLf
'    szSql = szSql & "  b.bus_id*=c.bus_id AND" & vbCrLf
'    szSql = szSql & "  b.bus_serial_no*=c.bus_serial_no AND" & vbCrLf
'    szSql = szSql & "  s.bus_date='" & Format(BusDate, "YYYY-MM-DD") & "' AND" & vbCrLf
'    If Len(BusID) > 0 Then
'        szSql = szSql & "  s.bus_id='" & BusID & "' AND" & vbCrLf
'    End If
'    szSql = szSql & "  b.bus_serial_no>0 AND" & vbCrLf
'    szSql = szSql & "  b.vehicle_id=v.vehicle_id AND" & vbCrLf
'    szSql = szSql & "  v.vehicle_type_code=t.vehicle_type_code AND" & vbCrLf
'    szSql = szSql & "  v.owner_id=o.owner_id AND" & vbCrLf
'    szSql = szSql & "  v.transport_company_id = p.transport_company_id" & vbCrLf
'    szSql = szSql & "GROUP BY b.bus_date,b.bus_id,b.bus_serial_no,v.license_tag_no,t.vehicle_type_short_name,o.owner_name,p.transport_company_short_name" & vbCrLf
'    szSql = szSql & "ORDER BY b.bus_date,b.bus_id,b.bus_serial_no"
'    oDb.Execute szSql
'
'    szSql = "SELECT bus_id,bus_serial_no,sale_quantity,sale_price,check_quantity,check_price,license_tag_no,vehicle_type_short_name,owner_name,transport_company_short_name FROM #SNTemp"
'    Set rsTemp = oDb.Execute(szSql, , -1)
'    Set GetScrollCheckTicketReport = rsTemp
End Function

Public Function GetRegularCheckTicketReport(BusDate As Date, Optional BusID As String = "")
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    
    Dim oDb As New RTConnection
    Dim rsTemp As New Recordset
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "Exec RegularCheckQuery '" & Format(BusDate, "YYYY-MM-DD") & "','" & BusID & "'"
    Set rsTemp = oDb.Execute(szSql)
    Set GetRegularCheckTicketReport = rsTemp

''清临时表
'    szSql = ""
'    szSql = "IF EXISTS(SELECT * FROM tempdb.dbo.sysobjects WHERE name LIKE '#SNTemp2%')" & vbCrLf
'    szSql = szSql & "DROP TABLE #SNTemp2" & vbCrLf
'    szSql = szSql & "IF EXISTS(SELECT * FROM tempdb.dbo.sysobjects WHERE name LIKE '#SNTemp%')" & vbCrLf
'    szSql = szSql & "DROP TABLE #SNTemp" & vbCrLf
'    oDb.Execute szSql
'
'    szSql = "SELECT b.bus_date,b.bus_id,b.bus_serial_no," & vbCrLf
'    szSql = szSql & "v.license_tag_no,t.vehicle_type_short_name,o.owner_name,p.transport_company_short_name," & vbCrLf
'    szSql = szSql & "COUNT(s.ticket_id) AS 'sale_quantity'," & vbCrLf
'    szSql = szSql & "SUM(s.ticket_price) AS 'sale_price'," & vbCrLf
'    szSql = szSql & "COUNT(c.ticket_id) AS 'check_quantity'," & vbCrLf
'    szSql = szSql & "SUM(c.ticket_price) AS 'check_price'" & vbCrLf
'    szSql = szSql & "INTO #SNTemp" & vbCrLf
'    szSql = szSql & "FROM check_ticket_lst c,sale_ticket_lst s," & vbCrLf
'    szSql = szSql & "     Bus_check_lst b,transport_vehicle_info v,transport_company_info p,owner_info o,vehicle_type_code t" & vbCrLf
'    szSql = szSql & "WHERE s.ticket_id*=c.ticket_id AND" & vbCrLf
'    szSql = szSql & "  s.bus_date=b.bus_date AND s.bus_id=b.bus_id AND" & vbCrLf
'    szSql = szSql & "  b.bus_date*=c.bus_date AND" & vbCrLf
'    szSql = szSql & "  b.bus_id*=c.bus_id AND" & vbCrLf
'    szSql = szSql & "  b.bus_serial_no*=c.bus_serial_no AND" & vbCrLf
'    szSql = szSql & "  s.bus_date='" & Format(BusDate, "YYYY-MM-DD") & "' AND"
'    If Len(BusID) > 0 Then
'        szSql = szSql & "  s.bus_id='" & BusID & "' AND" & vbCrLf
'    End If
'    szSql = szSql & "  b.bus_serial_no=0 AND" & vbCrLf
'    szSql = szSql & "  b.vehicle_id=v.vehicle_id AND" & vbCrLf
'    szSql = szSql & "  v.vehicle_type_code=t.vehicle_type_code AND" & vbCrLf
'    szSql = szSql & "  v.owner_id=o.owner_id AND" & vbCrLf
'    szSql = szSql & "  v.transport_company_id = p.transport_company_id" & vbCrLf
'    szSql = szSql & "GROUP BY b.bus_date,b.bus_id,b.bus_serial_no,v.license_tag_no,t.vehicle_type_short_name,o.owner_name,p.transport_company_short_name" & vbCrLf
'    szSql = szSql & "ORDER BY b.bus_date,b.bus_id,b.bus_serial_no" & vbCrLf
''改乘
'    szSql = szSql & "SELECT bus_date,bus_id,bus_serial_no," & vbCrLf
'    szSql = szSql & "   COUNT(ticket_id) AS 'change_quantity'," & vbCrLf
'    szSql = szSql & "   SUM(ticket_price) AS 'change_price'" & vbCrLf
'    szSql = szSql & "INTO #SNTemp2" & vbCrLf
'    szSql = szSql & "FROM check_ticket_lst" & vbCrLf
'    szSql = szSql & "WHERE bus_date='" & Format(BusDate, "YYYY-MM-DD") & "' AND status=2" & vbCrLf
'    szSql = szSql & "GROUP BY bus_date,bus_id,bus_serial_no" & vbCrLf
'    szSql = szSql & "ORDER BY bus_date,bus_id,bus_serial_no" & vbCrLf
''将改乘数加到检票数
'    szSql = szSql & "UPDATE #SNTemp SET check_quantity=check_quantity+change_quantity,check_price=check_price+change_price" & vbCrLf
'    szSql = szSql & "FROM #SNTemp t1,#SNTemp2 t2" & vbCrLf
'    szSql = szSql & "WHERE t1.bus_date=t2.bus_date AND" & vbCrLf
'    szSql = szSql & "  t1.bus_id=t2.bus_id AND" & vbCrLf
'    szSql = szSql & "  t1.bus_serial_no = t2.bus_serial_no" & vbCrLf
'    oDb.Execute szSql
'
'    szSql = "SELECT bus_id,sale_quantity,sale_price,check_quantity,check_price,license_tag_no,vehicle_type_short_name,owner_name,transport_company_short_name FROM #SNTemp"
'    Set rsTemp = oDb.Execute(szSql, , -1)
'    Set GetRegularCheckTicketReport = rsTemp
End Function


''根据车次代码查询车次的所有检票人数
'Public Function GetBusTicketReport(dtStart As Date, dtEnd As Date, Optional BusID As String = "")
'    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
'    Dim odb As New RTConnection
'    Dim rsTemp As New Recordset
'    Dim szSql As String, szWhere As String
'    odb.ConnectionString = GetConnectionStr(cszSystemMan)
'    If BusID <> "" Then szWhere = "AND bus_id in (" & BusID & ")"
'    szSql = "select max(bus_id) as object_id,max(bus_serial_no) as object_name,count(ticket_id) as check_number,sum(ticket_price) as total_price ,max(check_time) as check_time FROM check_ticket_lst " _
'    & " WHERE  bus_date between '" & ToDBDate(dtStart) & " ' and  '" & ToDBDate(dtEnd) & " '"
'    szSql = szSql & szWhere & " group by bus_id order by bus_id"
'    Set rsTemp = odb.Execute(szSql)
'    Set GetBusTicketReport = rsTemp
'End Function

'根据车次代码查询车次的所有检票人数
Public Function GetBusTicketReport(dtStart As Date, dtEnd As Date, Optional BusID As String = "", Optional SellStationID As String = "")
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    Dim oDb As New RTConnection
    Dim szSql As String, szWhere As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    
    
    If BusID <> "" Then szWhere = "AND c.bus_id in (" & BusID & ")"
    
    szSql = "SELECT g.sell_station_id , MAX(s.sell_station_name) sell_station_name , MAX(c.bus_id) AS object_id,MAX(c.bus_serial_no) AS object_name," _
        & " COUNT(ticket_id) AS check_number,sum(ticket_price) AS total_price " _
        & " ,MAX(check_time) AS check_time " _
        & " ,MAX(end_station_name) AS end_station_name " _
        & " ,MAX(license_tag_no) AS license_tag_no " _
        & " FROM check_ticket_lst c , checkgate_info g , sell_station_info s, env_bus_info e  , vehicle_info v " _
        & " WHERE c.bus_date BETWEEN '" & ToDBDate(dtStart) & " ' AND  '" & ToDBDate(dtEnd) & " '" _
        & " AND c.check_gate_id = g.check_gate_id AND g.sell_station_id = s.sell_station_id AND c.bus_date = e.bus_date AND c.bus_id = e.bus_id " _
        & " AND e.bus_id = c.bus_id AND e.bus_date = c.bus_date" _
        & " AND e.vehicle_id = v.vehicle_id " _


'    szSql = "select g.sell_station_id , max(s.sell_station_name) sell_station_name , max(c.bus_id) as object_id,max(bus_serial_no) as object_name,count(ticket_id) as check_number,sum(ticket_price) as total_price ,max(check_time) as check_time ,MAX(end_station_name) AS end_station_name from check_ticket_lst c , checkgate_info g , sell_station_info s, env_bus_info e  " _
'    & " WHERE  c.bus_date between '" & ToDBDate(dtStart) & " ' and  '" & ToDBDate(dtEnd) & " ' and c.check_gate_id = g.check_gate_id and g.sell_station_id = s.sell_station_id and c.bus_date = e.bus_date and c.bus_id = e.bus_id and c.check_gate_id = e.check_gate_id  "
'& " group by c.bus_id , g.sell_station_id order by c.bus_id , g.sell_station_id "
    If SellStationID <> "" Then szSql = szSql & " AND g.sell_station_id = '" & SellStationID & "'"
    
    
    szSql = szSql & szWhere & " GROUP BY c.bus_id , g.sell_station_id " _
        & " ORDER BY c.bus_id , g.sell_station_id"
    
    
    
    Set GetBusTicketReport = oDb.Execute(szSql)
End Function

''根据线路代码查询线路的所有检票人数
'Public Function GetRouteTicketReport(dtStart As Date, dtEnd As Date, Optional RouteID As String = "")
'    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
'    Dim odb As New RTConnection
'    Dim rsTemp As New Recordset
'    Dim szSql As String, szWhere As String
'    odb.ConnectionString = GetConnectionStr(cszSystemMan)
'    If RouteID <> "" Then szWhere = "AND cbi.route_id in (" & RouteID & ")"
'    szSql = "SELECT max(cbi.route_id) as object_id,max(ri.route_name) as object_name,count(cti.ticket_id) as check_number,sum(cti.ticket_price)as total_price,max(cti.check_time) as check_time FROM check_ticket_lst cti,Bus_check_lst cbi,Route_info ri " _
'    & " WHERE  cti.bus_id=cbi.bus_id AND cti.bus_serial_no=cbi.bus_serial_no AND cti.bus_date=cbi.bus_date  AND cbi.route_id=ri.route_id AND  cti.bus_date between '" & ToDBDate(dtStart) & " ' and  '" & ToDBDate(dtEnd) & " '"
'    szSql = szSql & szWhere & "group by cbi.route_id"
'    Set rsTemp = odb.Execute(szSql)
'    Set GetRouteTicketReport = rsTemp
'End Function
'
'
''根据站点代码查询站点的所有检票人数
'Public Function GetStationTicketReport(dtStart As Date, dtEnd As Date, Optional StationID As String = "")
'    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
'    Dim odb As New RTConnection
'    Dim rsTemp As New Recordset
'    Dim szSql As String, szWhere As String
'    odb.ConnectionString = GetConnectionStr(cszSystemMan)
'    If StationID <> "" Then szWhere = "AND cti.end_station_id in (" & StationID & ")"
'    szSql = "SELECT max(cti.end_station_id) as object_id,max(si.station_name) as object_name,count(cti.ticket_id) as check_number,sum(cti.ticket_price) as total_price ,max(cti.check_time) as check_time FROM check_ticket_lst cti,station_info si " _
'    & " WHERE   cti.end_station_id=si.station_id AND  cti.bus_date between '" & ToDBDate(dtStart) & " ' and  '" & ToDBDate(dtEnd) & " '"
'    szSql = szSql & szWhere & "group by cti.end_station_id"
'    Set rsTemp = odb.Execute(szSql)
'    Set GetStationTicketReport = rsTemp
'End Function
''根据检票口代码查询检票口的所有检票人数
'Public Function GetCheckGateTicketReport(dtStart As Date, dtEnd As Date, Optional CheckGateID As String = "")
'    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
'    Dim odb As New RTConnection
'    Dim rsTemp As New Recordset
'    Dim szSql As String, szWhere As String
'    odb.ConnectionString = GetConnectionStr(cszSystemMan)
'    If CheckGateID <> "" Then szWhere = "AND cgi.check_gate_id in (" & CheckGateID & ")"
'    szSql = "SELECT max(ebi.check_gate_id) as object_id,max(cgi.check_gate_name) as object_name,count(cti.ticket_id) as check_number,sum(cti.ticket_price)as total_price,max(cti.check_time) as check_time FROM check_ticket_lst cti,Checkgate_info cgi,Env_bus_info ebi " _
'    & " WHERE  cti.bus_id =ebi.bus_id AND cti.bus_date=ebi.bus_date AND cgi.check_gate_id=ebi.check_gate_id AND  cti.bus_date between '" & ToDBDate(dtStart) & " ' and  '" & ToDBDate(dtEnd) & " '"
'    szSql = szSql & szWhere & "group by cgi.check_gate_id"
'    Set rsTemp = odb.Execute(szSql)
'    Set GetCheckGateTicketReport = rsTemp
'End Function
'根据线路代码查询线路的所有检票人数
Public Function GetRouteTicketReport(dtStart As Date, dtEnd As Date, Optional RouteID As String = "", Optional SellStationID As String = "")
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    Dim oDb As New RTConnection
    Dim rsTemp As New Recordset
    Dim szSql As String, szWhere As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    If RouteID <> "" Then szWhere = "AND cbi.route_id in (" & RouteID & ")"
    szSql = "select g.sell_station_id , max(s.sell_station_name) sell_station_name , max(cbi.route_id) as object_id,max(ri.route_name) as object_name,count(cti.ticket_id) as check_number,sum(cti.ticket_price)as total_price,max(cti.check_time) as check_time from check_ticket_lst cti,bus_check_lst cbi,Route_info ri , checkgate_info g , sell_station_info s " _
    & " WHERE  cti.bus_id=cbi.bus_id AND cti.bus_serial_no=cbi.bus_serial_no AND cti.bus_date=cbi.bus_date  AND cbi.route_id=ri.route_id AND  cti.bus_date between '" & ToDBDate(dtStart) & "' and  '" & ToDBDate(dtEnd) & "' AND cti.check_gate_id = g.check_gate_id AND g.sell_station_id = s.sell_station_id "
    If SellStationID <> "" Then szSql = szSql & " and g.sell_station_id = '" & SellStationID & "'"
    szSql = szSql & szWhere & " group by cbi.route_id  , g.sell_station_id  ORDER BY cbi.route_id , g.sell_station_id "
    
    Set rsTemp = oDb.Execute(szSql)
    Set GetRouteTicketReport = rsTemp
End Function
'根据站点代码查询站点的所有检票人数
Public Function GetStationTicketReport(dtStart As Date, dtEnd As Date, Optional StationID As String = "", Optional SellStationID As String = "")
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    Dim oDb As New RTConnection
    Dim rsTemp As New Recordset
    Dim szSql As String, szWhere As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    If StationID <> "" Then szWhere = "AND cti.end_station_id in (" & StationID & ")"
    szSql = "select g.sell_station_id , max(s.sell_station_name) sell_station_name , max(cti.end_station_id) as object_id,max(si.station_name) as object_name,count(cti.ticket_id) as check_number,sum(cti.ticket_price) as total_price ,max(cti.check_time) as check_time from check_ticket_lst cti,station_info si , checkgate_info g , sell_station_info s " _
    & " WHERE   cti.end_station_id=si.station_id AND  cti.bus_date between '" & ToDBDate(dtStart) & "' and  '" & ToDBDate(dtEnd) & "' AND cti.check_gate_id = g.check_gate_id AND g.sell_station_id = s.sell_station_id "
    If SellStationID <> "" Then szSql = szSql & " and g.sell_station_id = '" & SellStationID & "'"
    szSql = szSql & szWhere & "group by cti.end_station_id , g.sell_station_id "
    Set rsTemp = oDb.Execute(szSql)
    Set GetStationTicketReport = rsTemp
End Function
'根据检票口代码查询检票口的所有检票人数
Public Function GetCheckGateTicketReport(dtStart As Date, dtEnd As Date, Optional CheckGateID As String = "", Optional SellStationID As String = "")
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    Dim oDb As New RTConnection
    Dim rsTemp As New Recordset
    Dim szSql As String, szWhere As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    If CheckGateID <> "" Then szWhere = "AND cgi.check_gate_id in (" & CheckGateID & ")"
    szSql = "select g.sell_station_id , max(s.sell_station_name) sell_station_name , max(ebi.check_gate_id) as object_id,max(cgi.check_gate_name) as object_name,count(cti.ticket_id) as check_number,sum(cti.ticket_price)as total_price,max(cti.check_time) as check_time from check_ticket_lst cti,checkgate_info cgi,env_bus_info ebi , checkgate_info g , sell_station_info s " _
    & " WHERE  cti.bus_id =ebi.bus_id AND cti.bus_date=ebi.bus_date AND cgi.check_gate_id=ebi.check_gate_id AND  cti.bus_date between '" & ToDBDate(dtStart) & "' and  '" & ToDBDate(dtEnd) & "' AND cti.check_gate_id = g.check_gate_id AND g.sell_station_id = s.sell_station_id  "
    If SellStationID <> "" Then szSql = szSql & " and g.sell_station_id = '" & SellStationID & "'"
    szSql = szSql & szWhere & "group by cgi.check_gate_id ,g.sell_station_id  "
    Set rsTemp = oDb.Execute(szSql)
    Set GetCheckGateTicketReport = rsTemp
End Function




'检票员路单明细
Public Function GetCheckerSheetDetail(UserID As String, StartDate As Date, EndDate As Date) As Recordset
'    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
'    AssertHaveRight m_oActiveUser, RIGHT_SellerEveryDaySellDetail
    
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr()
    
    szSql = "SELECT * FROM sheet_lst WHERE " _
    & " checker ='" & UserID & "' AND " _
    & " sheet_make_date>='" & ToDBDateTime(StartDate) & "' AND " _
    & " sheet_make_date<'" & ToDBDateTime(EndDate) & "'" _
    & " ORDER BY check_sheet_id "
    
    Set rsTemp = oDb.Execute(szSql)
    Set GetCheckerSheetDetail = rsTemp
    
End Function

Public Function GetCheckerSheetAnotherThing(UserID As String, StartDate As Date, EndDate As Date) As Double()
    
'    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
'    AssertHaveRight m_oActiveUser, RIGHT_SellerEveryDayAnotherThing
    
    Dim adbResult(1 To 2) As Double
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    
    oDb.ConnectionString = GetConnectionStr()
    szSql = "SELECT COUNT(*) AS number  FROM sheet_lst " _
    & "  WHERE " _
    & " checker ='" & UserID & "' AND " _
    & " sheet_make_date>='" & ToDBDateTime(StartDate) & "' AND " _
    & " sheet_make_date<'" & ToDBDateTime(EndDate) & "'" _
    & " AND valid_mark = 0 "
    
    
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        adbResult(1) = FormatDbValue(rsTemp!Number)
    End If
    Set rsTemp = Nothing
    GetCheckerSheetAnotherThing = adbResult
End Function

'检票员检票统计月报
Public Function GetCheckerEveryMonth(UserID() As String, StartDate As Date, EndDate As Date) As Recordset
    
    Dim adbResult(1 To 2) As Double
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim szUser As String
    oDb.ConnectionString = GetConnectionStr()
    szUser = TeamToString(UserID, , True)
    szSql = "SELECT RTRIM(checker) +'['+RTRIM(max(user_name))+']' AS checker,RTRIM(s.check_gate_id) +'['+RTRIM(max(c.check_gate_name))+']' AS check_gate " _
        & " ,bus_date,COUNT(*) AS check_count," _
        & " SUM(quantity) AS check_number, " _
        & " SUM(ticket_price) AS check_price " _
        & " FROM sheet_lst s,user_info u,sheet_station_lst l,checkgate_info c WHERE " _
        & " checker in (" & szUser & ")  " _
        & " AND s.checker=u.user_id AND " _
        & " valid_mark <> 0 " _
        & " AND s.check_sheet_id=l.check_sheet_id " _
        & " AND c.check_gate_id=s.check_gate_id " _
        & " AND bus_date>='" & ToDBDate(StartDate) & "' AND " _
        & " bus_date<='" & ToDBDate(EndDate) & "'" _
        & " GROUP BY checker,bus_date,s.check_gate_id ORDER BY checker,bus_date,s.check_gate_id"
    
    Set rsTemp = oDb.Execute(szSql)

    Set GetCheckerEveryMonth = rsTemp
End Function


'车次途经站统计(按车次日期进行汇总)从路单站点表中取
Public Function GetBusStationStatByBusDate(pszCompanyID As String, pdtFromDate As Date, pdtToDate As Date, Optional pszLike As String = "", Optional pszSellStationID As String = "") As Recordset
'pnExtraStatus =1 表示售票，2表示补票，3表示全部



    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
'    AssertHaveRight m_oActiveUser, RIGHT_GetBusStatByBusDate
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    
    oDb.ConnectionString = GetConnectionStr("")
    If pszCompanyID <> "" Then
        szWhere = " AND g.split_company_id IN ( " & pszCompanyID & ")"
    End If
    '如果要查询的售票站不为空
    If pszSellStationID <> "" Then
        szWhere = szWhere & " AND b.sell_station_id='" & pszSellStationID & "'"
    End If
'    If pnExtraStatus <> TP_AllSold Then
'    '如果不是所有的状态
'        szWhere = szWhere & " AND extra_status = " & pnExtraStatus
'
'    End If
    If pszLike <> "" Then
        szWhere = szWhere & " AND b.bus_id like '" & pszLike & "%' "
    End If

    szSql = "   SELECT b.bus_id, convert(char(5), MAX(g.bus_start_time),108) bus_start_time , MAX(g.route_id) route_id , MAX(route_name) route_name " _
            & "    , b.station_id , MAX(s.station_name) station_name , ticket_type , MAX(e.station_name) end_station_name " _
            & " ,MAX(transport_company_short_name) transport_company_short_name  " _
            & "    , SUM(ticket_price*quantity) ticket_price " _
            & " , SUM(quantity) passenger_number"
        szSql = szSql _
            & " , SUM(b.mileage) mileage  " _
            & " FROM fin_sheet_station_lst b , route_info r , station_info s , station_info e,company_info c , env_bus_info g , sheet_lst t " _
            & " WHERE b.check_sheet_id = t.check_sheet_id AND r.route_id = g.route_id AND b.station_id = s.station_id AND r.end_station_id = e.station_id " _
            & " AND g.transport_company_id = c.transport_company_id AND b.bus_id = g.bus_id AND b.bus_date = g.bus_date " _
            & " AND  b.bus_date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND b.bus_date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "'" & szWhere & ExceptZeroBus _
            & " AND t.valid_mark = " & TransFieldValueToString(1) _
            & " GROUP BY b.bus_id , b.station_id , ticket_type ORDER BY b.bus_id , b.station_id , ticket_type "
    Set rsTemp = oDb.Execute(szSql)
    Set GetBusStationStatByBusDate = rsTemp
    
End Function



'车次途经站统计(按车次日期进行汇总)从路单站点表中取
Public Function GetVehicleStationStatByBusDate(pszCompanyID As String, pdtFromDate As Date, pdtToDate As Date, Optional pszLike As String = "", Optional pszSellStationID As String = "") As Recordset
'pnExtraStatus =1 表示售票，2表示补票，3表示全部



    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
'    AssertHaveRight m_oActiveUser, RIGHT_GetBusStatByBusDate
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    
    oDb.ConnectionString = GetConnectionStr("")
    If pszCompanyID <> "" Then
        szWhere = " AND g.split_company_id IN ( " & pszCompanyID & ")"
    End If
    '如果要查询的售票站不为空
    If pszSellStationID <> "" Then
        szWhere = szWhere & " AND b.sell_station_id='" & pszSellStationID & "'"
    End If
'    If pnExtraStatus <> TP_AllSold Then
'    '如果不是所有的状态
'        szWhere = szWhere & " AND extra_status = " & pnExtraStatus
'
'    End If
    If pszLike <> "" Then
        szWhere = szWhere & " AND b.bus_id like '" & pszLike & "%' "
    End If

    szSql = "   SELECT g.route_id , g.vehicle_id , MAX(r.route_name) route_name , MAX(v.license_tag_no) license_tag_no " _
            & "    , b.station_id , MAX(s.station_name) station_name , ticket_type " _
            & " ,MAX(transport_company_short_name) transport_company_short_name  " _
            & "    , SUM(ticket_price*quantity) ticket_price " _
            & " , SUM(quantity) passenger_number "
        szSql = szSql _
            & " , SUM(b.mileage) mileage  " _
            & " FROM  fin_sheet_station_lst b , station_info s ,company_info c , env_bus_info g , vehicle_info v ,route_info r, sheet_lst t " _
            & " WHERE b.check_sheet_id = t.check_sheet_id AND b.station_id = s.station_id AND r.route_id = g.route_id " _
            & " AND g.transport_company_id = c.transport_company_id AND b.bus_id = g.bus_id AND b.bus_date = g.bus_date " _
            & " AND g.vehicle_id = v.vehicle_id " _
            & " AND b.bus_date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND b.bus_date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "'" & szWhere & ExceptZeroBus _
            & " AND t.valid_mark = " & TransFieldValueToString(1) _
            & " GROUP BY g.route_id ,g.vehicle_id , b.station_id , ticket_type ORDER BY g.route_id ,g.vehicle_id ,b.station_id , ticket_type "
    Set rsTemp = oDb.Execute(szSql)
    Set GetVehicleStationStatByBusDate = rsTemp
    
End Function



'得到不是零的条件
Private Function ExceptZeroBus() As String

'
'
'    ExceptZeroBus = "  AND NOT ( passenger_number  = 0  AND  ticket_price = 0" _
'    & " AND ticket_cancel_number = 0 AND ticket_cancel_amount = 0 AND ticket_return_number = 0 AND ticket_return_amount = 0 " _
'    & " AND  ticket_return_charge = 0 AND ticket_change_number = 0 AND  ticket_change_amount  = 0 AND ticket_change_charge  = 0) "
''    & " AND total_seat =0 AND fact_float =0 AND total_float =0 )  "

    ExceptZeroBus = ""

End Function

