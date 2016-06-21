VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BaseInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**********************************************************
'* Source File Name:BaseInfo.cls
'* Project Name:STBase.vbp
'* Engineer:
'* Data Generated:
'* Last Revision Date:
'* Brief Description:基本对象
'* Relational Document:
'**********************************************************

Option Explicit
'出错
Public Enum EErrBaseInfo
    ERR_BaseInfoNotAvailable = ERR_BaseInfo + ERR_NotAvailable '基本信息对象处于无状态
    ERR_VechileSeatTypeINFOExistNotAddNew = ERR_BaseInfo + 22 ' 信息已存在不能重复增加
    ERR_NoVehicleSeatTypeInfo = ERR_BaseInfo + 23 '数据库车辆座位信息不存在
End Enum

'权限
Public Enum ERightBaseInfo
    RIGHT_BaseInfoManagement = ERR_BaseInfo + cnMidRightBegin + cnMidRightStep * 1 '54906校对用户是否有权限创建对象
End Enum

Private m_oActiveUser As ActiveUser 'P1所有的类都有的私有类变量

'**************************************************
'Member Code:F1
'Brief Description:获得所有的公路等级
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/14
'**************************************************
Public Function GetAllRoadLevel(Optional RoadLevel As String = "") As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    If RoadLevel = "" Then
        Set rsTemp = oDb.Execute("SELECT * FROM road_level_code")
    Else
        Set rsTemp = oDb.Execute("SELECT * FROM road_level_code WHERE road_rank_code='" & RoadLevel & "'")
    End If
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    ReDim szaTemp(1 To nCount, 1 To 3) As String
    For i = 1 To nCount
        szaTemp(i, 1) = FormatDbValue(rsTemp!road_rank_code)
        szaTemp(i, 2) = FormatDbValue(rsTemp!short_name)
        szaTemp(i, 3) = FormatDbValue(rsTemp!road_rank_name)
        rsTemp.MoveNext
    Next
    GetAllRoadLevel = szaTemp
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

'**************************************************
'Member Code:F2
'Brief Description:获得所有的地区
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/14
'**************************************************
Public Function GetAllArea(Optional Area As String = "") As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    If Area = "" Then
        Set rsTemp = oDb.Execute("SELECT * FROM area_code")
    Else
        Set rsTemp = oDb.Execute("SELECT * FROM area_code WHERE area_code='" & Area & "'")
    End If
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    ReDim szaTemp(1 To nCount, 1 To 4) As String
    For i = 1 To nCount
        szaTemp(i, 1) = FormatDbValue(rsTemp!area_code)
        szaTemp(i, 2) = FormatDbValue(rsTemp!area_name)
        szaTemp(i, 3) = CStr(FormatDbValue(rsTemp!province_in_out))    '省内或省外
        szaTemp(i, 4) = FormatDbValue(rsTemp!Annotation)
        rsTemp.MoveNext
    Next
    GetAllArea = szaTemp
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

'**************************************************
'Member Code:F3
'Brief Description:获所有的车型
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/14
'**************************************************
Public Function GetAllVehicleModel(Optional VehicleModel As String = "") As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    If VehicleModel = "" Then
        Set rsTemp = oDb.Execute("SELECT * FROM vehicle_type_code")
    Else
        Set rsTemp = oDb.Execute("SELECT * FROM vehicle_type_code WHERE vehicle_type_code='" & VehicleModel & "'")
    End If
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    ReDim szaTemp(1 To nCount, 1 To 4) As String
    For i = 1 To nCount
        szaTemp(i, 1) = FormatDbValue(rsTemp!vehicle_type_code)
        szaTemp(i, 2) = FormatDbValue(rsTemp!vehicle_type_short_name)
        szaTemp(i, 3) = FormatDbValue(rsTemp!seat_quantity)
        szaTemp(i, 4) = FormatDbValue(rsTemp!vehicle_type_name)
        rsTemp.MoveNext
    Next
    GetAllVehicleModel = szaTemp
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

'**************************************************
'Member Code:F4
'Brief Description:获得所有的检票口
'Engineer:wjb
'Date Generated:1999/10/11
'Last Revision Date:2002/11/9
'**************************************************
Public Function GetAllCheckGate(Optional CheckGate As String = "", Optional pszSellStationID As String = "") As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    Dim rzSql As String
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    If CheckGate = "" Then
        rzSql = rzSql + "SELECT c.*, s.sell_station_name,s.sell_station_full_name FROM checkgate_info c,sell_station_info s WHERE c.sell_station_id=s.sell_station_id"
    Else
        rzSql = rzSql + "SELECT c.*, s.sell_station_name,s.sell_station_full_name FROM checkgate_info c,sell_station_info s WHERE c.sell_station_id=s.sell_station_id and c.check_gate_id='" & CheckGate & "'"
    End If
    If pszSellStationID <> "" Then
        rzSql = rzSql + " AND c.sell_station_id='" & pszSellStationID & "'"
    End If
    Set rsTemp = oDb.Execute(rzSql)
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    ReDim szaTemp(1 To nCount, 1 To 6) As String
    For i = 1 To nCount
        szaTemp(i, 1) = FormatDbValue(rsTemp!check_gate_id)
        szaTemp(i, 2) = FormatDbValue(rsTemp!check_gate_name)
        szaTemp(i, 3) = FormatDbValue(rsTemp!Annotation)
        szaTemp(i, 4) = FormatDbValue(rsTemp!sell_station_id)
        szaTemp(i, 5) = FormatDbValue(rsTemp!sell_station_name)
        szaTemp(i, 6) = FormatDbValue(rsTemp!sell_station_full_name)
        rsTemp.MoveNext
    Next
    GetAllCheckGate = szaTemp
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

'**************************************************
'Member Code:F5
'Brief Description:获得所有的车辆
'Engineer:wjb
'Date Generated:1999/10/11
'Last Revision Date:2002/11/9
'**************************************************
Public Function GetVehicle(Optional VehicleID As String = "", Optional Company As String = "", Optional Owner As String = "", Optional BusType As String = "", Optional License As String = "", Optional OnlyNormal As Boolean = True) As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim i As Integer, nCount As Integer
    Dim szWhere As String
    Dim szSql As String
    
    AssertObjIsValid
    If VehicleID <> "" Then
        szWhere = " AND vehicle_id LIKE '%" & VehicleID & "%'"
    End If
    
    If Company <> "" Then
        szWhere = szWhere & " AND tbt.transport_company_id='" & Company & "'"
    End If
    
    If Owner <> "" Then
        szWhere = szWhere & " AND tbt.owner_id='" & Owner & "'"
    End If
    
    If BusType <> "" Then
        szWhere = szWhere & " AND tbt.vehicle_type_code='" & BusType & "'"
    End If
    
    If License <> "" Then
        szWhere = szWhere & " AND tbt.license_tag_no LIKE '%" & License & "%'"
    End If
    
    If OnlyNormal Then
        szWhere = szWhere & " AND tbt.delete_status =1"
'    Else
'        szWhere = szWhere & " AND delete_status =1"
    End If
    
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = " SELECT tbt.*,tbc.transport_company_short_name,tbo.owner_name,tbm.vehicle_type_short_name,tbc2.transport_company_short_name split_company_short_name " _
            & " FROM Vehicle_info tbt, Company_info tbc ,Company_info tbc2,owner_info tbo,vehicle_type_code tbm " _
            & " WHERE tbt.vehicle_type_code=tbm.vehicle_type_code AND " _
            & " tbo.owner_id=tbt.owner_id AND tbc.transport_company_id=tbt.transport_company_id" _
            & " AND tbc2.transport_company_id=tbt.split_company_id "
            
    szSql = szSql & szWhere
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    ReDim szaTemp(1 To nCount, 1 To 10) As String
    For i = 1 To rsTemp.RecordCount
        szaTemp(i, 1) = FormatDbValue(rsTemp!vehicle_id)
        szaTemp(i, 2) = FormatDbValue(rsTemp!license_tag_no)
        szaTemp(i, 3) = FormatDbValue(rsTemp!seat_quantity)
        szaTemp(i, 4) = FormatDbValue(rsTemp!transport_company_short_name)
        szaTemp(i, 5) = FormatDbValue(rsTemp!owner_name)
        szaTemp(i, 6) = FormatDbValue(rsTemp!Status)
        szaTemp(i, 7) = FormatDbValue(rsTemp!vehicle_type_code)
        szaTemp(i, 8) = FormatDbValue(rsTemp!vehicle_type_short_name)
        szaTemp(i, 9) = FormatDbValue(rsTemp!Annotation)
        szaTemp(i, 10) = FormatDbValue(rsTemp!split_company_short_name)
        rsTemp.MoveNext
    Next
    GetVehicle = szaTemp
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

'**************************************************
'Member Code:F6
'Brief Description:获得所有的车主
'Engineer:wjb
'Date Generated:1999/10/11
'Last Revision Date:2002/11/9
'**************************************************
Public Function GetOwner(Optional Company As String = "", Optional Owner As String = "", Optional OnlyNormal As Boolean = True) As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    Dim szSql As String, szWhere As String
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    If Company <> "" Then
        szWhere = " AND transport_company_id='" & Company & "'"
    End If
    If Company <> "" And Owner <> "" Then
        szWhere = " AND owner_id='" & Owner & "'"
    End If
    '此段与上相同,是否是单独判断 Owner <>""
    'If Company = "" And Owner <> "" Then
    '    szWhere = " AND owner_id='" & Owner & "'"
    'End If
    If OnlyNormal Then
        szWhere = szWhere & " AND delete_status =1"
'    Else
'        szSql = szSql & " AND delete_status =1"
    End If
    szSql = "SELECT owner_id,owner_name,annotation ,account_id FROM owner_info WHERE 1=1  " & szWhere
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    ReDim szaTemp(1 To nCount, 1 To 4) As String
    For i = 1 To nCount
        szaTemp(i, 1) = FormatDbValue(rsTemp!owner_id)
        szaTemp(i, 2) = FormatDbValue(rsTemp!owner_name)
        szaTemp(i, 3) = FormatDbValue(rsTemp!Annotation)
        szaTemp(i, 4) = FormatDbValue(rsTemp!account_id)
        rsTemp.MoveNext
    Next
    GetOwner = szaTemp
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function
'**************************************************
'Member Code:F7
'Brief Description:获得所有的参运公司
'Engineer:wjb
'Date Generated:1999/10/11
'Last Revision Date:2002/11/9
'**************************************************
Public Function GetCompany(Optional Company As String = "", Optional OnlyNormal As Boolean = True, Optional SplitCompany As Boolean = False) As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    If Company = "" Then
        szSql = "SELECT transport_company_id,transport_company_short_name,Annotation ,account_id FROM Company_info WHERE 1=1 "
    Else
        szSql = "SELECT transport_company_id,transport_company_short_name,Annotation ,account_id  FROM Company_info WHERE transport_company_id='" & Company & "'"
    End If
    If OnlyNormal Then
        szSql = szSql & " AND delete_status =1"
'    Else
'        szSql = szSql & " AND delete_status =1"
    End If
    If SplitCompany Then
        szSql = szSql & " AND transport_company_name LIKE ('%本公司%') OR transport_company_name LIKE ('%外公司%')"
    End If
    
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
        ReDim szaTemp(1 To nCount, 1 To 4) As String
        For i = 1 To nCount
            szaTemp(i, 1) = FormatDbValue(rsTemp!transport_company_id)
            szaTemp(i, 2) = FormatDbValue(rsTemp!transport_company_short_name)
            szaTemp(i, 3) = FormatDbValue(rsTemp!Annotation)
            szaTemp(i, 4) = FormatDbValue(rsTemp!account_id)
            rsTemp.MoveNext
        Next
        GetCompany = szaTemp
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

'**************************************************
'Member Code:F8
'Brief Description:获得所有的站点
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/14
'**************************************************
Public Function GetStation(Optional Area As String = "", Optional StationName As String = "", Optional StationID As String = "", Optional StationInput As String = "") As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim szSql As String, szWhere As String
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    If Area <> "" Then
        szWhere = " AND station_info.area_code='" & Area & "'"
    End If
    
    If StationName <> "" And StationInput <> "" And StationID <> "" _
           And StationName = StationInput And StationInput = StationID And StationName = StationID Then
                   szWhere = " AND ( ltrim(station_name) LIKE '" & Trim(StationName) & "%' or" _
                             & " ltrim(station_input_code) LIKE '" & Trim(StationInput) & "%' or" _
                             & "  ltrim(station_id) LIKE '" & Trim(StationID) & "%')"
    Else
    
        If StationName <> "" Then
            szWhere = szWhere & " AND ltrim(station_name) LIKE '" & Trim(StationName) & "%'"
        End If
        
        If StationInput <> "" Then
            szWhere = szWhere & " AND ltrim(station_input_code) LIKE '" & Trim(StationInput) & "%'"
        End If
        
        If StationID <> "" Then
            szWhere = szWhere & " AND ltrim(station_id) LIKE '" & Trim(StationID) & "%'"
        End If
        
    End If
    szSql = "SELECT * FROM station_info,area_code WHERE station_info.area_code=area_code.area_code " & szWhere
    
    
    Set rsTemp = oDb.Execute(szSql)
    
    nCount = rsTemp.RecordCount
    
    If nCount <> 0 Then
    
    ReDim szaTemp(1 To nCount, 1 To 6) As String
    
    For i = 1 To nCount
        szaTemp(i, 1) = FormatDbValue(rsTemp!station_id)
        szaTemp(i, 2) = FormatDbValue(rsTemp!station_name)
        szaTemp(i, 3) = FormatDbValue(rsTemp!station_input_code)
        szaTemp(i, 4) = FormatDbValue(rsTemp!station_rank)
        szaTemp(i, 5) = FormatDbValue(rsTemp!local_station_code)
        szaTemp(i, 6) = FormatDbValue(rsTemp!area_name)
        rsTemp.MoveNext
    Next
    
    GetStation = szaTemp
    End If
    
    Set rsTemp = Nothing
    Set oDb = Nothing
    
End Function
'**************************************************
'Member Code:F9
'Brief Description:获得所有的线路
'Engineer:
'Date Generated:1999/10/20
'Last Revision Date:1999/10/20
'**************************************************
Public Function GetRoute(Optional RouteID As String = "", Optional RouteName As String = "", Optional Station As String = "") As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    Dim szWhere As String
    Dim szSql As String
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    If RouteID <> "" Then
        szWhere = " AND tbr.route_id LIKE '" & RouteID & "%'"
    End If
    If RouteName <> "" Then
        szWhere = szWhere & " AND tbr.route_name LIKE '" & RouteName & "%'"
    End If
    If Station <> "" Then
        szWhere = szWhere & " AND tba.end_station_id='" & Station & "'"
    End If
    szSql = "SELECT DISTINCT tbr.*,tbs.station_input_code,tbs.station_name FROM route_info tbr,station_info tbs,Route_section_lst tba WHERE tbr.end_station_id=tbs.station_id AND tbr.route_id=tba.route_id "
    szSql = szSql & szWhere & " order by tbr.route_id,tbs.station_input_code,route_name"
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    
    '徐建华改的加了一度,线路票价公式,因是数组,不影响其它程序,故不需要特别通知
    'ReDim szaTemp(1 To nCount, 1 To 6) As String
    ReDim szaTemp(1 To nCount, 1 To 8) As String
    For i = 1 To nCount
        szaTemp(i, 1) = FormatDbValue(rsTemp!route_id)
        szaTemp(i, 2) = FormatDbValue(rsTemp!route_name)
        szaTemp(i, 3) = FormatDbValue(rsTemp!Annotation)
        szaTemp(i, 4) = FormatDbValue(rsTemp!station_name)
        szaTemp(i, 5) = FormatDbValue(rsTemp!Mileage)
        szaTemp(i, 6) = FormatDbValue(rsTemp!Status)
        szaTemp(i, 7) = FormatDbValue(rsTemp!formula_name)
        szaTemp(i, 8) = FormatDbValue(rsTemp!station_input_code)
        rsTemp.MoveNext
    Next
    GetRoute = szaTemp
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
    Set m_oActiveUser = vNewValue
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
    'AssertHaveRight m_oActiveUser, RIGHT_BaseInfoManagement  '校对用户是否有查询基本信息的权限
End Sub

'**************************************************
'Member Code:S2
'Brief Description:类初始化
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub AssertObjIsValid()
    AssertActiveUserValid m_oActiveUser, ERR_BaseInfo
End Sub

'**************************************************
'Member Code:F10
'Brief Description:获得所有的线路
'Engineer:
'Date Generated:2002/11/18
'Last Revision Date:2002/11/18
'**************************************************

Public Function GetRouteEx(Optional RouteID As String = "", Optional RouteName As String = "", Optional Station As String = "", Optional SellStation As String = "") As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer, nRoute As Integer, nNoCount
    Dim rsTemp As Recordset
    Dim szTemp As String, szUpTemp As String, szWhere As String, szSql As String, szStation As String
    AssertObjIsValid
    If RouteID <> "" Then
        szWhere = " AND r.route_id = '" & Trim(RouteID) & "'"
    End If
    If RouteName <> "" Then
        szWhere = szWhere & " AND r.route_name LIKE '" & Trim(RouteName) & "%'"
    End If
    Dim szSellStationWhere As String
    If SellStation <> "" Then
        szSellStationWhere = szSellStationWhere & " AND d.sell_station_id = " & TransFieldValueToString(SellStation)
    End If
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    '或得无路段的线路
    nNoCount = oDb.Execute("SELECT r.* FROM route_info r " _
        & " WHERE r.mileage=0 " & szWhere).RecordCount
    If Station <> "" Then
    
        szWhere = szWhere & "  AND( d.route_id IN ( SELECT route_id FROM Route_section_lst d " _
            & "  WHERE d.end_station_id='" & Station & "')    " _
            & "  or d.route_id IN (SELECT route_id FROM Route_section_lst tba WHERE tba.end_station_id IN  " _
            & "  ( select station_id from station_info where station_name LIKE '" & Trim(Station) & "%')))   "
    End If
    szSql = "SELECT r.route_id,r.route_name,r.mileage,r.status,s.station_name,  r.status,d.end_station_id,d.end_station_mileage " _
        & " , r.start_station_id , v.station_name start_station_name ,r.id_card " _
        & " FROM route_info r , Route_section_lst d ,station_info s  , station_info v, " _
        & "(SELECT a.sell_station_id,a.route_id FROM  Route_section_lst a," & _
                                                      "(SELECT route_id,MAX(end_station_mileage) end_station_mileage" & _
                                                      " FROM Route_section_lst GROUP BY route_id) b" & _
                " WHERE a.route_id=b.route_id AND a.end_station_mileage=b.end_station_mileage) e" _
        & " WHERE r.route_id = d.route_id" _
        & " AND r.start_station_id = v.station_id AND d.end_station_id = s.station_id " _
        & " AND e.sell_station_id = d.sell_station_id And e.route_id = d.route_id " _
        & szWhere & szSellStationWhere _
        & " ORDER BY d.route_id,d.section_serial "


    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nNoCount <> 0 Then
        ReDim szaTemp(1 To nNoCount, 1 To 9) As String
    End If
    If nCount <> 0 Then
    Do While rsTemp.EOF = False
        szUpTemp = FormatDbValue(rsTemp!route_id)
        If szTemp = szUpTemp Or szTemp = "" Then
            If szTemp = "" Then szTemp = szUpTemp
        Else
            nRoute = nRoute + 1
            szTemp = szUpTemp
        End If
        rsTemp.MoveNext
    Loop
    nRoute = nRoute + 1
    rsTemp.MoveFirst
    szTemp = "": szUpTemp = szTemp
    If nRoute + nNoCount = 0 Then Exit Function
    ReDim szaTemp(1 To nRoute + nNoCount, 1 To 9) As String
    nRoute = 1
    Do While rsTemp.EOF = False
        szUpTemp = FormatDbValue(rsTemp!route_id)
        If szTemp = szUpTemp Or szTemp = "" Then
            '途经站及里程合并
            szStation = szStation & FormatDbValue(rsTemp!station_name) & "[" & FormatDbValue(rsTemp!end_station_mileage) & "]"
            szTemp = szUpTemp
        Else
            '
            rsTemp.MovePrevious
            szaTemp(nRoute, 1) = FormatDbValue(rsTemp!route_id)  '线路代码
            szaTemp(nRoute, 2) = FormatDbValue(rsTemp!route_name)  '线路名称
            szaTemp(nRoute, 3) = szStation
            szaTemp(nRoute, 4) = FormatDbValue(rsTemp!station_name)
            szaTemp(nRoute, 5) = FormatDbValue(rsTemp!Mileage)
            szaTemp(nRoute, 6) = FormatDbValue(rsTemp!Status)  '线路状态
            szaTemp(nRoute, 7) = FormatDbValue(rsTemp!start_station_name)
            szaTemp(nRoute, 9) = FormatDbValue(rsTemp!id_card)
            szTemp = szUpTemp
            nRoute = nRoute + 1
            szStation = ""
        End If
    rsTemp.MoveNext
    Loop
    rsTemp.MovePrevious
    szaTemp(nRoute, 1) = FormatDbValue(rsTemp!route_id)  '线路代码
    szaTemp(nRoute, 2) = FormatDbValue(rsTemp!route_name)  '线路名称
    szaTemp(nRoute, 3) = szStation
    szaTemp(nRoute, 4) = FormatDbValue(rsTemp!station_name)
    szaTemp(nRoute, 5) = FormatDbValue(rsTemp!Mileage)
    szaTemp(nRoute, 6) = FormatDbValue(rsTemp!Status)  '线路状态
    szaTemp(nRoute, 7) = FormatDbValue(rsTemp!start_station_name)
    szaTemp(nRoute, 9) = FormatDbValue(rsTemp!id_card)
    End If
    If nNoCount <> 0 Then
    
    Set rsTemp = oDb.Execute("SELECT r.*,v.station_name FROM route_info r,station_info v" _
        & " WHERE r.start_station_id *= v.station_id AND r.mileage=0" & IIf(Station = "", szWhere, ""))
    For i = nRoute + 1 To nRoute + nNoCount
        szaTemp(i, 1) = FormatDbValue(rsTemp!route_id)  '线路代码
        szaTemp(i, 2) = FormatDbValue(rsTemp!route_name)  '线路名称
        szaTemp(i, 3) = "(无)"
        szaTemp(i, 4) = "(无)"
        szaTemp(i, 5) = FormatDbValue(rsTemp!Mileage)
        szaTemp(i, 6) = FormatDbValue(rsTemp!Status)  '线路状态
        szaTemp(i, 7) = FormatDbValue(rsTemp!station_name)
        szaTemp(i, 9) = FormatDbValue(rsTemp!id_card)
        rsTemp.MoveNext
    Next
    End If
    GetRouteEx = szaTemp
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

'**************************************************
'Member Code:F11
'Brief Description:获得所有的路段
'Engineer:
'Date Generated:1999/12/27
'Last Revision Date:1999/12/27
'**************************************************
'Public Function GetSection(Optional RouteID As String = "", Optional SectionID As String = "", Optional Station As String = "") As String()
'    Dim szaTemp() As String
'    Dim oDb As New rtConnection
'    Dim i As Integer, nCount As Integer, nRoute As Integer
'    Dim rsTemp As Recordset
'    Dim szSectionID() As String
'    Dim szTemp As String, szUpTemp As String, szWhere As String, szSql As String, szStation As String
'    AssertObjIsValid
'    If SectionID <> "" Then
'        szWhere = " AND tbd.section_id LIKE '" & SectionID & "%'"
'    End If
'    If Station <> "" Then
'        szWhere = szWhere & " AND tbd.end_station='" & Station & "'"
'    End If
'    If RouteID <> "" Then
'        szWhere = szWhere & " AND tbd.section_id IN(SELECT section_id FROM Route_section_lst WHERE route_id='" & RouteID & "')"
'    End If
'    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
'    szSql = "SELECT tbd.section_id,path_no,mileage,section_name,area_name,road_rank_name," _
'            & " tbs.station_name AS start_name,tbe.station_name AS end_name FROM " _
'            & " section_info tbd,station_info tbs,station_info tbe," _
'            & " area_code tba,road_rank_code tbr " _
'            & " WHERE tbd.start_station = tbs.station_id " _
'            & " AND tbd.end_station=tbe.station_id " _
'            & " AND tbd.area_code=tba.area_code " _
'            & " AND tbd.road_rank_code=tbr.road_rank_code " _
'            & szWhere
'    Set rsTemp = oDb.Execute(szSql)
'    nCount = rsTemp.RecordCount
'    If nCount <> 0 Then
'    ReDim szaTemp(1 To nCount, 1 To 8) As String
'    For i = 1 To nCount
'        szaTemp(i, 1) = formatdbvalue(rstemp!section_id
'        szaTemp(i, 2) = formatdbvalue(rstemp!section_name
'        szaTemp(i, 3) = formatdbvalue(rstemp!start_name
'        szaTemp(i, 4) = formatdbvalue(rstemp!end_name
'        szaTemp(i, 5) = formatdbvalue(rstemp!Mileage
'        szaTemp(i, 6) = formatdbvalue(rstemp!road_rank_name
'        szaTemp(i, 7) = formatdbvalue(rstemp!area_name
'        szaTemp(i, 8) = formatdbvalue(rstemp!path_no
'        rsTemp.MoveNext
'    Next
'    GetSection = szaTemp
'    End If
'    Set rsTemp = Nothing
'    Set oDb = Nothing
'End Function


Public Function GetSection(Optional RouteID As String = "", Optional SectionID As String = "", Optional Station As String = "", Optional szStarStationID As String = "") As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer, nRoute As Integer
    Dim rsTemp As Recordset
    Dim szTemp As String, szUpTemp As String, szWhere As String, szSql As String, szStation As String
    AssertObjIsValid
    
    If SectionID <> "" Then
        szWhere = " AND tbd.section_id LIKE '" & SectionID & "%'"
    End If
    
    If Station <> "" Then
        szWhere = szWhere & " AND tbd.end_station='" & Station & "'"
    End If
    
    If szStarStationID <> "" Then
       
       szWhere = szWhere & " and tbd.start_station='" & szStarStationID & "'"
    
    End If
    
    If RouteID <> "" Then
        szWhere = szWhere & " AND tbd.section_id IN(SELECT section_id FROM Route_section_lst WHERE route_id='" & RouteID & "')"
    End If
    
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT tbd.section_id,path_no,mileage,section_name,area_name,road_rank_name," _
            & " tbs.station_name AS start_name,tbe.station_name AS end_name FROM " _
            & " section_info tbd,station_info tbs,station_info tbe," _
            & " area_code tba,road_level_code tbr " _
            & " WHERE tbd.start_station = tbs.station_id " _
            & " AND tbd.end_station=tbe.station_id " _
            & " AND tbd.area_code=tba.area_code " _
            & " AND tbd.road_rank_code=tbr.road_rank_code " _
            & szWhere
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    ReDim szaTemp(1 To nCount, 1 To 8) As String
    For i = 1 To nCount
        szaTemp(i, 1) = FormatDbValue(rsTemp!section_id)
        szaTemp(i, 2) = FormatDbValue(rsTemp!section_name)
        szaTemp(i, 3) = FormatDbValue(rsTemp!start_name)
        szaTemp(i, 4) = FormatDbValue(rsTemp!end_name)
        szaTemp(i, 5) = FormatDbValue(rsTemp!Mileage)
        szaTemp(i, 6) = FormatDbValue(rsTemp!road_rank_name)
        szaTemp(i, 7) = FormatDbValue(rsTemp!area_name)
        szaTemp(i, 8) = FormatDbValue(rsTemp!path_no)
        rsTemp.MoveNext
    Next
    GetSection = szaTemp
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function
Public Function GetAllBusType(Optional szBusType As String) As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    Dim szSql As String
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    If szBusType = "" Then
    szSql = "SELECT * FROM bus_type_code"
    Else
    szSql = "SELECT * FROM bus_type_code Where Bus_type ='" & szBusType & "'"
    End If
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
'    if ncount=0 then showerror ERR_NotBusType
    If nCount <> 0 Then
    ReDim szaTemp(1 To nCount, 1 To 3) As String
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
Public Function GetAllSeatType(Optional szSeatType As String) As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    Dim szSql As String
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    If szSeatType = "" Then
    szSql = "SELECT * FROM Seat_type_code"
    Else
    szSql = "SELECT * FROM Seat_type_code where seat_type_id='" & szSeatType & "'"
    End If
        Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
  
    If nCount <> 0 Then
    ReDim szaTemp(1 To nCount, 1 To 3) As String
    For i = 1 To nCount
        szaTemp(i, 1) = FormatDbValue(rsTemp!seat_type_id)
        szaTemp(i, 2) = FormatDbValue(rsTemp!seat_type_name)
        szaTemp(i, 3) = FormatDbValue(rsTemp!Annotation)
        rsTemp.MoveNext
    Next
    GetAllSeatType = szaTemp
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function



Public Function GetAllVehicleSeatTypeInfo(Optional szVehicleID As String) As TVehcileSeatType()
    Dim oDb As New RTConnection
    Dim szTempSql As String
    Dim szSql As String
    Dim szSqlContext As String
    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim nCount As Integer
    Dim szTemp() As TVehcileSeatType
    
    
    
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    If szVehicleID <> "" Then
        szSql = "SELECT distinct tve.* ,sti.seat_type_name,vtc.vehicle_type_code,vtc.vehicle_type_short_name FROM vehicle_seat_type_info tve ,Seat_type_code sti ,Vehicle_info trv , vehicle_type_code vtc where tve.vehicle_id='" & szVehicleID & " 'and " _
            & " sti.seat_type_id=tve.Seat_type_id and " _
            & " trv.vehicle_id=tve.vehicle_id  and vtc.vehicle_type_code=trv.vehicle_type_code "
    Else
        szSql = "SELECT distinct tve.* ,sti.seat_type_name,vtc.vehicle_type_code,vtc.vehicle_type_short_name FROM vehicle_seat_type_info tve ,Seat_type_code sti ,Vehicle_info trv , vehicle_type_code vtc where    " _
            & " sti.seat_type_id=tve.Seat_type_id and " _
            & " trv.vehicle_id=tve.vehicle_id  and vtc.vehicle_type_code=trv.vehicle_type_code "
    End If
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
        ReDim szTemp(1 To nCount)
        For i = 1 To nCount
            szTemp(i).szSerial = FormatDbValue(rsTemp!serial_no)
            szTemp(i).szVehcileTypeName = MakeDisplayString(FormatDbValue(rsTemp!vehicle_type_code), FormatDbValue(rsTemp!vehicle_type_short_name))
            szTemp(i).szVehcileID = FormatDbValue(rsTemp!vehicle_id)
            szTemp(i).szSeatTypeID = FormatDbValue(rsTemp!seat_type_id)
            szTemp(i).szSeatTypeName = FormatDbValue(rsTemp!seat_type_name)
            szTemp(i).szStartSeatNo = FormatDbValue(rsTemp!start_seat_no)
            szTemp(i).szEndSeatNo = FormatDbValue(rsTemp!End_seat_no)
            rsTemp.MoveNext
        Next
    End If
    GetAllVehicleSeatTypeInfo = szTemp
    Set rsTemp = Nothing
    Set oDb = Nothing
    
    
End Function
'
'
''///////////////////////////////////////////////
''得到环境中指定日期间的所有车次
'Public Function GetEnviroAllBus(pdtStartTime As Date, pdtEndTime As Date, Optional pbAll As Boolean = False) As String()
'    Dim i As Integer
'    Dim szSql As String
'    Dim rsTemp As Recordset
'    Dim oDb As New RTConnection
'    Dim aszTemp() As String
'
'    AssertObjIsValid
'    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
''    szSql = "SELECT DISTINCT s.bus_id,e.vehicle_type_name ,e.bus_start_time,e.end_station_name FROM sale_ticket_lst s,enviroment_bus_info e " _
''            & "WHERE (s.operation_time BETWEEN '" & ToDBDateTime(pdtStartTime) & "' AND '" & ToDBDateTime(pdtEndTime) & "') AND " _
''            & "s.bus_id=e.bus_id AND s.bus_date=e.bus_date ORDER BY s.bus_id"
''    Set rsTemp = odb.Execute(szSql)
'
'    szSql = "SELECT DISTINCT bus_id , substring(convert(nchar,bus_start_time,20),12,9) as StartupTime ," _
'                & " end_station_name ,vehicle_type_name FROM enviroment_bus_info " _
'                & " WHERE bus_date BETWEEN '" & ToDBDate(pdtStartTime) & "' AND '" & ToDBDate(pdtEndTime) & "' ORDER BY bus_id "
'    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount <> 0 Then
'        ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 5)
'        For i = 1 To rsTemp.RecordCount
'            aszTemp(i, 1) = FormatDbValue(rsTemp!Bus_ID)
'            aszTemp(i, 2) = FormatDbValue(rsTemp!StartupTime)
'            aszTemp(i, 3) = FormatDbValue(rsTemp!end_station_name)
'            aszTemp(i, 4) = FormatDbValue(rsTemp!vehicle_type_name)
'            aszTemp(i, 5) = ""
'            rsTemp.MoveNext
'        Next i
'    End If
'    GetEnviroAllBus = aszTemp
'End Function
'
