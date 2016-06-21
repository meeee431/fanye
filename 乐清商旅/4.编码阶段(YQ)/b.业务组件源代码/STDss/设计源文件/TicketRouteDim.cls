VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TicketRouteDim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_oActiveUser As ActiveUser


'常规操作
Public Sub Init(poActiveUser As ActiveUser)
    Set SelfUser = poActiveUser
End Sub

Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
End Property

Public Property Set SelfUser(ByVal vNewValue As ActiveUser)
    Set m_oActiveUser = vNewValue
End Property


'统计车次类型为正常和流水车次的市内、市外、省外的线路
'得到按地区统计的人数及金额，并列出要求的线路的人数金额。
Public Function GetRouteByAreaBusTypeSimply(pdtFromDate As Date, pdtToDate As Date, Optional pszSellStationID As String = "") As Recordset
    Dim oDB As New RTConnection
    Dim szSql  As String
    Dim rsTemp As Recordset
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetRouteByAreaBusTypeSimply
    
    
    
    oDB.ConnectionString = GetConnectionStr
    
    szSql = "SELECT CASE a.province_in_out WHEN " & EA_nInCity & " THEN '市内' WHEN " & EA_nOutCity & " THEN '市外' WHEN " & EA_nOutProvince & "  THEN '省外' END   AS province_in_out " _
        & " , r.end_station_id , MAX(s.station_name) station_name " _
        & " , (SUM(passenger_number) - SUM(ticket_cancel_number) - SUM(ticket_return_number)- SUM(ticket_change_number)) AS passenger_number " _
        & " , COUNT(bus_id) AS bus_count , SUM( fact_float ) AS fact_float , SUM( total_float ) AS total_float  , SUM(total_seat) AS total_number " _
        & " , CASE WHEN SUM(total_seat)<>0 THEN (SUM(passenger_number) - SUM(ticket_cancel_number) - SUM(ticket_return_number)- SUM(ticket_change_number)) * 100 / CONVERT(numeric(9,1), SUM(total_seat)) ELSE 0 END  AS full_seat_rate " _
        & " , CASE WHEN SUM(total_float)<>0 THEN  SUM( fact_float ) * 100 / SUM( total_float ) ELSE 0 END AS fact_load_rate " _
        & " , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
        & "     - SUM(ticket_change_amount) + SUM(ticket_return_charge) +SUM(ticket_change_charge) ) AS total_ticket_price " _
        & " FROM  stat_bus_startoff_lst b , route_info r , station_info s , area_code a " _
        & " WHERE date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "' " _
        & " AND r.route_id = b.route_id AND  ( b.bus_type = " & TP_ScrollBus & " OR b.bus_type = " & TP_RegularBus & " ) " _
        & " AND r.end_station_id = s.station_id and s.area_code = a.area_code "
        If pszSellStationID <> "" Then
            Dim szSellStationID() As String
            Dim i As Integer
            szSellStationID = Split(pszSellStationID, ",")
            szSql = szSql & " AND (b.sell_station_id='" & szSellStationID(0) & "' "
            For i = 1 To ArrayLength(szSellStationID) - 1
                szSql = szSql & " OR b.sell_station_id='" & szSellStationID(i) & "' "
            Next
            szSql = szSql & ")"
        End If
           szSql = szSql & " GROUP BY a.province_in_out , r.end_station_id " _
                         & " ORDER BY a.province_in_out , r.end_station_id "
    
    
    Set rsTemp = oDB.Execute(szSql)
    Set GetRouteByAreaBusTypeSimply = rsTemp
    
End Function



'按照车次类型统计线路人数、金额
'得到除正常班次与滚动车次以外的车次的线路统计，按车次类型与线路进行分类列出
Public Function GetRouteByBusTypeSimply(pdtFromDate As Date, pdtToDate As Date, Optional pszSellStationID As String = "") As Recordset
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetRouteByBusTypeSimply
    
    
    oDB.ConnectionString = GetConnectionStr
    
    szSql = " SELECT b.route_id , b.bus_type , MAX(t.bus_type_name) bus_type_name " _
        & " , MAX(r.route_name) route_name , (SUM(passenger_number) - SUM(ticket_cancel_number) - SUM(ticket_return_number)- SUM(ticket_change_number)) AS passenger_number " _
        & " , COUNT(bus_id) AS bus_count , SUM( fact_float ) AS fact_float , SUM( total_float ) AS total_float  , SUM(total_seat) AS total_number " _
        & " , CASE WHEN SUM(total_seat)<>0 THEN (SUM(passenger_number) - SUM(ticket_cancel_number) - SUM(ticket_return_number)- SUM(ticket_change_number)) * 100 / CONVERT(numeric(9,1) " _
        & " , SUM(total_seat)) ELSE 0 END  AS full_seat_rate " _
        & " , CASE WHEN SUM(total_float)<>0 THEN  SUM( fact_float ) * 100 / SUM( total_float ) ELSE 0 END AS fact_load_rate " _
        & " , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
        & "     - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price " _
        & " FROM  stat_bus_startoff_lst b , route_info r , Bus_type_code  t " _
        & " WHERE date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "' " _
        & " AND r.route_id = b.route_id AND NOT ( b.bus_type = " & TP_ScrollBus & " OR b.bus_type = " & TP_RegularBus & " ) " _
        & " AND t.bus_type = b.bus_type "
     If pszSellStationID <> "" Then
        Dim szSellStationID() As String
        Dim i As Integer
        szSellStationID = Split(pszSellStationID, ",")
        szSql = szSql & " AND (b.sell_station_id='" & szSellStationID(0) & "' "
        For i = 1 To ArrayLength(szSellStationID) - 1
            szSql = szSql & " OR b.sell_station_id='" & szSellStationID(i) & "' "
        Next
        szSql = szSql & ")"
     End If
        szSql = szSql & " GROUP BY b.route_id , b.bus_type " _
                      & " ORDER BY b.bus_type , r.route_name "

    Set rsTemp = oDB.Execute(szSql)
    Set GetRouteByBusTypeSimply = rsTemp
    
End Function


Public Function GetRouteTransport(pdtFromDate As Date, pdtToDate As Date, Optional pszCondiction As String = "", Optional pszSellStationID As String = "", Optional pszStatus As Boolean) As Recordset
    'pszCondiction 统计条件, pnStatMode 统计方式暂时不用
    Dim oDB As New RTConnection
    Dim szSql  As String
    Dim rsTemp As Recordset
    Dim szWhere As String
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetRouteTransport
    
    
    
    If pszCondiction <> "" Then
        szWhere = " AND b.route_id IN (" & pszCondiction & ")"
    End If
      '如果要查询的售票站不为空
    If pszSellStationID <> "" Then
        Dim szSellStationID() As String
        Dim i As Integer
        szSellStationID = Split(pszSellStationID, ",")
        szWhere = szWhere & " AND (b.sell_station_id='" & szSellStationID(0) & "' "
        For i = 1 To ArrayLength(szSellStationID) - 1
            szWhere = szWhere & " OR b.sell_station_id='" & szSellStationID(i) & "' "
        Next
        szWhere = szWhere & ")"
    End If
    
    oDB.ConnectionString = GetConnectionStr
    If pszStatus = False Then
        szSql = "SELECT b.route_id , MAX(r.route_name) ROUTE_NAME " _
            & " , (SUM(passenger_number) - SUM(ticket_cancel_number) - SUM(ticket_return_number)- SUM(ticket_change_number)) AS passenger_number " _
            & " , COUNT(bus_id) AS bus_count , SUM( fact_float ) AS fact_float , SUM( total_float ) AS total_float  , SUM(total_seat) AS total_number " _
            & " , CASE WHEN SUM(total_seat)<>0 THEN (SUM(passenger_number) - SUM(ticket_cancel_number) - SUM(ticket_return_number)- SUM(ticket_change_number)) * 100 / CONVERT(numeric(9,1), SUM(total_seat)) ELSE 0 END  AS full_seat_rate " _
            & " , CASE WHEN SUM(total_float)<>0 THEN  SUM( fact_float ) * 100 / SUM( total_float ) ELSE 0 END AS fact_load_rate " _
            & " , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "     - SUM(ticket_change_amount) + SUM(ticket_return_charge) +SUM(ticket_change_charge) ) AS total_ticket_price " _
            & " FROM  stat_bus_startoff_lst b , route_info r " _
            & " WHERE date >= '" & ToDBDateTime(pdtFromDate) & "'" _
                & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "' " _
            & " AND r.route_id = b.route_id " _
            & szWhere _
            & ExceptZeroBus _
            & " GROUP BY b.route_id" _
            & " ORDER BY b.route_id "
    Else
        szSql = "SELECT b.route_id , MAX(transport_company_short_name) transport_company_name , MAX(r.route_name) ROUTE_NAME " _
            & " , (SUM(passenger_number) - SUM(ticket_cancel_number) - SUM(ticket_return_number)- SUM(ticket_change_number)) AS passenger_number " _
            & " , COUNT(bus_id) AS bus_count , SUM( fact_float ) AS fact_float , SUM( total_float ) AS total_float  , SUM(total_seat) AS total_number " _
            & " , CASE WHEN SUM(total_seat)<>0 THEN (SUM(passenger_number) - SUM(ticket_cancel_number) - SUM(ticket_return_number)- SUM(ticket_change_number)) * 100 / CONVERT(numeric(9,1), SUM(total_seat)) ELSE 0 END  AS full_seat_rate " _
            & " , CASE WHEN SUM(total_float)<>0 THEN  SUM( fact_float ) * 100 / SUM( total_float ) ELSE 0 END AS fact_load_rate " _
            & " , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "     - SUM(ticket_change_amount) + SUM(ticket_return_charge) +SUM(ticket_change_charge) ) AS total_ticket_price " _
            & " FROM  stat_bus_startoff_lst b , route_info r , company_info c " _
            & " WHERE date >= '" & ToDBDateTime(pdtFromDate) & "'" _
                & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "' " _
            & " AND r.route_id = b.route_id AND b.transport_company_id=c.transport_company_id " _
            & szWhere _
            & ExceptZeroBus _
            & " GROUP BY b.route_id,b.transport_company_id" _
            & " ORDER BY b.route_id,b.transport_company_id "
    End If
    
    
    Set rsTemp = oDB.Execute(szSql)
    Set GetRouteTransport = rsTemp
    
End Function

Public Function GetRouteTransportByCheck(pdtFromDate As Date, pdtToDate As Date, Optional pszCondiction As String = "", Optional pszSellStationID As String = "", Optional pszStatus As Boolean) As Recordset
    'pszCondiction 统计条件, pnStatMode 统计方式暂时不用
    Dim oDB As New RTConnection
    Dim szSql  As String
    Dim rsTemp As Recordset
    Dim szWhere As String
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetRouteTransport
    
    
    
    If pszCondiction <> "" Then
        szWhere = " AND b.route_id IN (" & pszCondiction & ")"
    End If
      '如果要查询的售票站不为空
    If pszSellStationID <> "" Then
        Dim szSellStationID() As String
        Dim i As Integer
        szSellStationID = Split(pszSellStationID, ",")
        szWhere = szWhere & " AND (b.sell_station_id='" & szSellStationID(0) & "' "
        For i = 1 To ArrayLength(szSellStationID) - 1
            szWhere = szWhere & " OR b.sell_station_id='" & szSellStationID(i) & "' "
        Next
        szWhere = szWhere & ")"
    End If
    
    oDB.ConnectionString = GetConnectionStr
    If pszStatus = False Then
        szSql = "SELECT b.route_id , MAX(r.route_name) ROUTE_NAME " _
            & " , SUM(passenger_number) AS passenger_number " _
            & " , COUNT(bus_id) AS bus_count , SUM( fact_float ) AS fact_float , SUM( total_float ) AS total_float  , SUM(total_seat) AS total_number " _
            & " , CASE WHEN SUM(total_seat)<>0 THEN SUM(passenger_number) * 100 / CONVERT(numeric(9,1), SUM(total_seat)) ELSE 0 END  AS full_seat_rate " _
            & " , CASE WHEN SUM(total_float)<>0 THEN  SUM( fact_float ) * 100 / SUM( total_float ) ELSE 0 END AS fact_load_rate " _
            & " , SUM(ticket_price) AS total_ticket_price " _
            & " FROM  stat_bus_startoff_lst2 b , route_info r " _
            & " WHERE date >= '" & ToDBDateTime(pdtFromDate) & "'" _
                & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "' " _
            & " AND r.route_id = b.route_id " _
            & szWhere _
            & " GROUP BY b.route_id" _
            & " ORDER BY b.route_id "
    Else
        szSql = "SELECT b.route_id ,  MAX(transport_company_short_name) transport_company_name , MAX(r.route_name) ROUTE_NAME " _
            & " , SUM(passenger_number) AS passenger_number " _
            & " , COUNT(bus_id) AS bus_count , SUM( fact_float ) AS fact_float , SUM( total_float ) AS total_float  , SUM(total_seat) AS total_number " _
            & " , CASE WHEN SUM(total_seat)<>0 THEN SUM(passenger_number) * 100 / CONVERT(numeric(9,1), SUM(total_seat)) ELSE 0 END  AS full_seat_rate " _
            & " , CASE WHEN SUM(total_float)<>0 THEN  SUM( fact_float ) * 100 / SUM( total_float ) ELSE 0 END AS fact_load_rate " _
            & " , SUM(ticket_price) AS total_ticket_price " _
            & " FROM  stat_bus_startoff_lst2 b , route_info r, company_info c " _
            & " WHERE date >= '" & ToDBDateTime(pdtFromDate) & "'" _
                & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "' " _
            & " AND r.route_id = b.route_id AND b.transport_company_id=c.transport_company_id " _
            & szWhere _
            & " GROUP BY b.route_id,b.transport_company_id" _
            & " ORDER BY b.route_id,b.transport_company_id "
    End If
    
    
    Set rsTemp = oDB.Execute(szSql)
    Set GetRouteTransportByCheck = rsTemp
    
End Function

'线路营收,分为总额,应拆额及应收额
'按票价项明细列出
Public Function GetRouteTurnOver(pdyStartDate As Date, pdyEndDate As Date, Optional pszCompanyID As String = "", Optional pszAreaID As String = "", Optional pszSellStationID As String = "") As Recordset
    Dim oDB As New RTConnection
    Dim szSql  As String
    Dim rsTemp As Recordset
    Dim szWhere As String
    Dim nRoadBuildChargeItem As Integer '公建金
    Dim nFarDistanceAddChargeItem As Integer '加成费
    Dim nSpringItem As Integer '春运费
    Dim oSystemParam As New SystemParam
    
    Dim szSubTotalItem As String
    Dim aszSubTotalItem() As String
    Dim nSubTotalItem As Integer
    Dim nValue As Integer
    
    
    Dim szSum As String
    Dim szSum2 As String
    Dim i As Integer
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetRouteTurnOver
    
    
    oDB.ConnectionString = GetConnectionStr
    oSystemParam.Init m_oActiveUser
    '得到公建金为第几个票价项
    nRoadBuildChargeItem = oSystemParam.RoadBuildChargeItem
    nFarDistanceAddChargeItem = oSystemParam.FarDistanceAddChargeItem
    nSpringItem = oSystemParam.SpringItem
    
    If pszCompanyID <> "" Then
        szWhere = " AND b.split_company_id IN (" & pszCompanyID & ")"
    End If
    If pszAreaID <> "" Then
        szWhere = " AND a.Area_code IN (" & pszAreaID & ")"
    End If
      '如果要查询的售票站不为空
    If pszSellStationID <> "" Then
        Dim szSellStationID() As String
        szSellStationID = Split(pszSellStationID, ",")
        szWhere = szWhere & " AND (b.sell_station_id='" & szSellStationID(0) & "' "
        For i = 1 To ArrayLength(szSellStationID) - 1
            szWhere = szWhere & " OR b.sell_station_id='" & szSellStationID(i) & "' "
        Next
        szWhere = szWhere & ")"
    End If
    
    '取得小计项参数
    szSubTotalItem = oSystemParam.SubTotalItem1
    aszSubTotalItem = StringToTeam(szSubTotalItem)
    nSubTotalItem = ArrayLength(aszSubTotalItem)
    '得到 小计项 字符串
    szSum = ""
    szSum2 = ""
    For i = 1 To nSubTotalItem
        nValue = CInt(aszSubTotalItem(i))
        
        If nValue <> 0 And nValue <> nRoadBuildChargeItem And nValue <> nFarDistanceAddChargeItem And nValue <> nSpringItem Then
            szSum = szSum & IIf(i <> 1, "+", "") & "SUM(price_item_" & nValue & ")"
            szSum2 = szSum2 & " , SUM(price_item_" & nValue & ") out_price_item_" & nValue
        ElseIf nValue = 0 Then '如果为基本运价则乘以0.9
            szSum = szSum & IIf(i <> 1, "+", "") & "(SUM(base_price) * 0.9 )"
            szSum2 = szSum2 & " , (SUM(base_price) * 0.9) out_base_price "
        ElseIf nValue = nRoadBuildChargeItem Then '如果为公建金,则乘以0.9
            szSum = szSum & IIf(i <> 1, "+", "") & "(SUM(price_item_" & nValue & ") * 0.9)"
            szSum2 = szSum2 & " , (SUM(price_item_" & nValue & ") * 0.9) out_price_item_" & nValue
        ElseIf nValue = nFarDistanceAddChargeItem Then '加成
            szSum = szSum & IIf(i <> 1, "+", "") & "(SUM(price_item_" & nValue & ") * 0.9)"
            szSum2 = szSum2 & " , (SUM(price_item_" & nValue & ") * 0.9) out_price_item_" & nValue
        ElseIf nValue = nSpringItem Then '春运
            szSum = szSum & IIf(i <> 1, "+", "") & "(SUM(price_item_" & nValue & ") * 0.9)"
            szSum2 = szSum2 & " , (SUM(price_item_" & nValue & ") * 0.9) out_price_item_" & nValue
        End If
    Next i
    '得到应收合计 , 应拆合计
    szSum = " , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
        & "   - SUM(ticket_change_amount) + SUM(ticket_return_charge) +SUM(ticket_change_charge) ) " _
        & "   - (" & szSum & ") AS some_sum2 , " & szSum & "  AS Some_Sum  "
    
    szSum = szSum & szSum2
    

    szSql = "SELECT CASE a.province_in_out WHEN " & EA_nInCity & " THEN '市内' WHEN " & EA_nOutCity & " THEN '市外' WHEN " & EA_nOutProvince & "  THEN '省外' END   AS province_in_out  " _
        & " , a.area_code end_station_id , MAX(a.area_name) station_name  , (SUM(passenger_number) - SUM(ticket_cancel_number)" _
        & "  - SUM(ticket_return_number)) AS passenger_number " _
        & " , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
        & "   - SUM(ticket_change_amount) + SUM(ticket_return_charge) +SUM(ticket_change_charge) ) AS total_ticket_price " _
        & " , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_charge) ticket_change_charge " _
        & " , SUM(base_price) base_price " _
        & " , SUM(price_item_1) price_item_1 , SUM(price_item_2) price_item_2 , SUM(price_item_3) price_item_3 " _
        & " , SUM(price_item_4) price_item_4 , SUM(price_item_5) price_item_5 , SUM(price_item_6) price_item_6 " _
        & " , SUM(price_item_7) price_item_7 , SUM(price_item_8) price_item_8 , SUM(price_item_9) price_item_9 " _
        & " , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 , SUM(price_item_12) price_item_12 " _
        & " , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 , SUM(price_item_15) price_item_15 " _
        & szSum _
        & " FROM  Stat_bus_sell_lst b , route_info r , station_info s , area_code a " _
        & " WHERE date >= '" & ToDBDateTime(pdyStartDate) & "'" _
        & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdyEndDate)) & "' " _
        & " AND r.route_id = b.route_id " _
        & " AND r.end_station_id = s.station_id and s.area_code = a.area_code" _
        & szWhere _
        & " GROUP BY a.province_in_out , a.area_code  ORDER BY a.province_in_out , a.area_code"
        
    Set rsTemp = oDB.Execute(szSql)
    Set GetRouteTurnOver = rsTemp
    

End Function


'得到不是零的条件
Private Function ExceptZeroBus() As String

    

    ExceptZeroBus = "  AND NOT ( passenger_number  = 0  AND  ticket_price = 0" _
    & " AND ticket_cancel_number = 0 AND ticket_cancel_amount = 0 AND ticket_return_number = 0 AND ticket_return_amount = 0 " _
    & " AND  ticket_return_charge = 0 AND ticket_change_number = 0 AND  ticket_change_amount  = 0 AND ticket_change_charge  = 0) "
'    & " AND total_seat =0 AND fact_float =0 AND total_float =0 )  "
End Function

