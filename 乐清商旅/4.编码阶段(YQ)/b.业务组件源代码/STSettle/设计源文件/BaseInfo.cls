VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BackBaseInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**********************************************************
'* Source File Name:BackBaseInfo.cls
'* Project Name:STSettle.vbp
'* Engineer:陈峰
'* Data Generated:2004/03/12
'* Last Revision Date:2004/03/12
'* Brief Description:回程基本信息对象
'* Relational Document:
'**********************************************************
Private m_oActiveUser As ActiveUser


'**************************************************
'Member Code:F9
'Brief Description:获得所有的线路
'Engineer:
'Date Generated:2004/03/12
'Last Revision Date:2004/03/12
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
    szSql = "SELECT DISTINCT tbr.*,tbs.station_input_code,tbs.station_name " _
        & " FROM back_route_info tbr,station_info tbs,back_route_section_lst tba " _
        & " WHERE tbr.end_station_id=tbs.station_id AND tbr.route_id=tba.route_id "
        
    szSql = szSql & szWhere & " order by tbr.route_id,tbs.station_input_code,route_name"
    
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    
    ReDim szaTemp(1 To nCount, 1 To 7) As String
    For i = 1 To nCount
        szaTemp(i, 1) = FormatDbValue(rsTemp!route_id)
        szaTemp(i, 2) = FormatDbValue(rsTemp!route_name)
        szaTemp(i, 3) = FormatDbValue(rsTemp!Annotation)
        szaTemp(i, 4) = FormatDbValue(rsTemp!station_name)
        szaTemp(i, 5) = FormatDbValue(rsTemp!Mileage)
        szaTemp(i, 6) = FormatDbValue(rsTemp!Status)
'        szaTemp(i, 7) = FormatDbValue(rsTemp!formula_name)
        szaTemp(i, 7) = FormatDbValue(rsTemp!station_input_code)
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
'Date Generated:2003/10/11
'Last Revision Date:2003/10/14
'**************************************************
Public Sub Init(poAUser As ActiveUser)
    Set m_oActiveUser = poAUser
    'AssertHaveRight m_oActiveUser, RIGHT_BaseInfoManagement  '校对用户是否有查询基本信息的权限
End Sub

'**************************************************
'Member Code:S2
'Brief Description:类初始化
'Engineer:
'Date Generated:2003/10/11
'Last Revision Date:2003/10/12
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
    '获得无路段的线路
    nNoCount = oDb.Execute("SELECT r.* FROM back_route_info r " _
        & " WHERE r.mileage=0 " & szWhere).RecordCount
    If Station <> "" Then
    
        szWhere = szWhere & "  AND( d.route_id IN ( SELECT route_id FROM back_route_section_lst d " _
            & "  WHERE d.end_station_id='" & Station & "')    " _
            & "  or d.route_id IN (SELECT route_id FROM back_route_section_lst tba WHERE tba.end_station_id IN  " _
            & "  ( select station_id from station_info where station_name LIKE '" & Trim(Station) & "%')))   "
    End If
    szSql = "SELECT r.route_id,r.route_name,r.mileage,r.status,s.station_name,  r.status,d.end_station_id,d.end_station_mileage " _
        & " , r.start_station_id , v.station_name start_station_name " _
        & " FROM back_route_info r , back_route_section_lst d ,station_info s  , station_info v, " _
        & "(SELECT a.sell_station_id,a.route_id FROM  back_route_section_lst a," & _
                                                      "(SELECT route_id,MAX(end_station_mileage) end_station_mileage" & _
                                                      " FROM back_route_section_lst GROUP BY route_id) b" & _
                " WHERE a.route_id=b.route_id AND a.end_station_mileage=b.end_station_mileage) e" _
        & " WHERE r.route_id = d.route_id" _
        & " AND r.start_station_id = v.station_id AND d.end_station_id = s.station_id " _
        & " AND e.sell_station_id = d.sell_station_id And e.route_id = d.route_id " _
        & szWhere & szSellStationWhere _
        & " ORDER BY d.route_id,d.section_serial "
    
    
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nNoCount <> 0 Then
        ReDim szaTemp(1 To nNoCount, 1 To 8) As String
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
    ReDim szaTemp(1 To nRoute + nNoCount, 1 To 7) As String
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
    End If
    If nNoCount <> 0 Then
    
    Set rsTemp = oDb.Execute("SELECT r.*,v.station_name FROM back_route_info r,station_info v" _
        & " WHERE r.start_station_id *= v.station_id AND r.mileage=0" & IIf(Station = "", szWhere, ""))
    For i = nRoute + 1 To nRoute + nNoCount
        szaTemp(i, 1) = FormatDbValue(rsTemp!route_id)  '线路代码
        szaTemp(i, 2) = FormatDbValue(rsTemp!route_name)  '线路名称
        szaTemp(i, 3) = "(无)"
        szaTemp(i, 4) = "(无)"
        szaTemp(i, 5) = FormatDbValue(rsTemp!Mileage)
        szaTemp(i, 6) = FormatDbValue(rsTemp!Status)  '线路状态
        szaTemp(i, 7) = FormatDbValue(rsTemp!station_name)
        rsTemp.MoveNext
    Next
    End If
    GetRouteEx = szaTemp
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

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
        szWhere = szWhere & " AND tbd.section_id IN(SELECT section_id FROM back_route_section_lst WHERE route_id='" & RouteID & "')"
    End If
    
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT tbd.section_id,path_no,mileage,section_name,area_name,road_rank_name," _
            & " tbs.station_name AS start_name,tbe.station_name AS end_name " _
            & " FROM section_info tbd,station_info tbs,station_info tbe,area_code tba,road_level_code tbr " _
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


