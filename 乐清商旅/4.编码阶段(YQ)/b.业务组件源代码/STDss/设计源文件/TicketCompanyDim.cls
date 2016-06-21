VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TicketCompanyDim"
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


'拆帐公司营收简报统计
Public Function GetSplitCompanySimpleCon(pszCompanyID As String, pdtFromDate As Date, pdtToDate As Date, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszSellStationID As String = "") As Recordset

    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetSplitCompanySimpleCon
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    If pszCompanyID <> "" Then
        szWhere = " AND b.split_company_id IN (" & pszCompanyID & ")"
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
    If pnExtraStatus <> TP_AllSold Then
    '如果不是所有的状态
        szWhere = szWhere & " AND extra_status = " & pnExtraStatus
        
    End If
        szSql = "   SELECT b.split_company_id , b.transport_company_id  , MAX(t.transport_company_short_name) split_company_name , MAX(t2.transport_company_short_name) transport_company_name" _
            & "    , SUM(base_price) base_price , SUM (price_item_1) price_item_1, SUM(price_item_2) price_item_2   " _
            & "    , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 , SUM(price_item_5) price_item_5  " _
            & "    , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 ,SUM(price_item_8) price_item_8  " _
            & "    , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 " _
            & "    , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 " _
            & "    , SUM(price_item_15) price_item_15 " _
            & GetTotalItem
            '& "    , SUM(base_price) + SUM(price_item_1) + SUM(price_item_2)  + SUM(price_item_3)  + SUM(price_item_7)  + SUM(price_item_12) + SUM(price_item_13) Some_Sum " _
            & "    , SUM(base_price) + SUM(price_item_1) + SUM(price_item_12) + SUM(price_item_13) Base_Price2 " _
            & "    , SUM(price_item_4) + SUM(price_item_5)  + SUM(price_item_6)  + SUM(price_item_8)  + SUM(price_item_9) + SUM(price_item_10) + SUM(price_item_11) + SUM(price_item_14)+ SUM(price_item_15)   Some_Sum2 "
            
        szSql = szSql _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number) - SUM(ticket_return_number) ) AS passenger_number " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) ) AS ticket_price " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price  " _
            & "    , SUM(ticket_cancel_number) ticket_cancel_number , SUM(ticket_cancel_amount) ticket_cancel_amount " _
            & "    , SUM(ticket_return_number) ticket_return_number , SUM(ticket_return_amount) ticket_return_amount " _
            & "    , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_number) ticket_change_number " _
            & "    , SUM(ticket_change_amount) ticket_change_amount , SUM(ticket_change_charge) ticket_change_charge " _
            & " FROM  Stat_bus_sell_lst b , Company_info t, Company_info t2  " _
            & " WHERE date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "'" _
            & " AND b.split_company_id = t.transport_company_id  and b.transport_company_id = t2.transport_company_id " & szWhere _
            & " GROUP BY b.split_company_id , b.transport_company_id  ORDER BY b.split_company_id , b.transport_company_id "
            
            

    Set rsTemp = oDB.Execute(szSql)
    Set GetSplitCompanySimpleCon = rsTemp
    
End Function



'添加车次公司组合
Public Sub AddBusCompanyCombine(pnCombineSerial As Integer, pszCombineName As String, pszStartBusID As String, pszEndBusID As String, pszTransportCompanyID As String)
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_AddBusCompanyCombine

    oDB.ConnectionString = GetConnectionStr(cszDss)
    
    szSql = " SELECT * FROM Company_info  WHERE transport_company_id = '" & pszTransportCompanyID & "'"
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount <> 1 Then ShowError ERR_CompanyNotExist
    
    szSql = " INSERT bus_assemble_info ( " _
        & " combine_serial , " _
        & " combine_name," _
        & " start_bus_id , " _
        & " end_bus_id , " _
        & " transport_company_id " _
        & " ) VALUES ( " _
        & pnCombineSerial & ", '" _
        & pszCombineName & "','" _
        & pszStartBusID & "' , '" _
        & pszEndBusID & "' , '" _
        & pszTransportCompanyID & "' ) "
    
    oDB.Execute szSql
    
    Set oDB = Nothing
    
End Sub

'删除组合中的车次公司
Public Sub DeleteCombileBusCompany(pnCombineSerial As Integer, pszStartBusID As String, pszEndBusID As String)
    Dim szSql As String
    Dim oDB As New RTConnection
    
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_DeleteCombileBusCompany
    oDB.ConnectionString = GetConnectionStr(cszDss)
    
    szSql = " DELETE bus_assemble_info " _
        & " WHERE combine_serial = " & pnCombineSerial _
        & " AND start_bus_id = '" & pszStartBusID & "'" _
        & " AND end_bus_id = '" & pszEndBusID & "'"
        
    oDB.Execute szSql
    Set oDB = Nothing
    
End Sub

'删除指定的组合
Public Sub DeleteCombile(nCombineSerial As Integer)
    Dim szSql As String
    Dim oDB As New RTConnection
    
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_DeleteCombile
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    
    szSql = " DELETE bus_assemble_info " _
        & " WHERE combine_serial = " & nCombineSerial
    
    oDB.Execute szSql
    Set oDB = Nothing
    
End Sub

'查询指定条件的车次组合
Public Function QueryCombile(Optional pnCombineSerial As Integer = -1, Optional pszCombineName As String = "", Optional pszStartBusID As String = "", Optional pszEndBusID As String = "", Optional pszCompanyID As String = "", Optional pszCompanyName As String = "") As tBusCompanyCombineInfo()
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim nCount As Integer
    Dim atTemp() As tBusCompanyCombineInfo
    Dim szWhere As String
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
     
    AssertHaveRight m_oActiveUser, RIGHT_QueryCombile
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    If pnCombineSerial <> -1 Then szWhere = szWhere & " AND combine_serial = " & pnCombineSerial
    If pszStartBusID <> "" Then szWhere = szWhere & " AND start_bus_id = '" & pszStartBusID & "'"
    If pszEndBusID <> "" Then szWhere = szWhere & " AND end_bus_id ='" & pszEndBusID & "'"
    If pszCompanyID <> "" Then szWhere = szWhere & " AND b.transport_company_id = '" & pszCompanyID & "'"
    If pszCompanyName <> "" Then szWhere = szWhere & " AND t.transport_company_short_name = '" & pszCompanyName & "'"
    
    szSql = "SELECT B.* , T.transport_company_short_name , T.transport_company_name " _
        & " FROM bus_assemble_info B , Company_info t " _
        & " WHERE b.transport_company_id = t.transport_company_id   " _
        & szWhere
        
    Set rsTemp = oDB.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then ReDim atTemp(1 To nCount)
    For i = 1 To nCount
        atTemp(i).CombineSerial = FormatDbValue(rsTemp!combine_serial)
        atTemp(i).StartBusId = FormatDbValue(rsTemp!start_bus_id)
        atTemp(i).EndBusID = FormatDbValue(rsTemp!end_bus_id)
        atTemp(i).TransportCompanyID = FormatDbValue(rsTemp!transport_company_id)
        atTemp(i).TransportCompanyName = FormatDbValue(rsTemp!transport_company_short_name)
        atTemp(i).CombineName = FormatDbValue(rsTemp!combine_name)
        rsTemp.MoveNext
    Next i
    
    QueryCombile = atTemp
    
End Function

'得到唯一的组合
Public Function GetUniqueCombine() As String()
    Dim aszTemp() As String
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim nCount As Integer
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetUniqueCombine
    oDB.ConnectionString = GetConnectionStr(cszDss)
    
    szSql = " SELECT distinct combine_serial FROM bus_assemble_info "
    Set rsTemp = oDB.Execute(szSql)
    
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then ReDim aszTemp(1 To nCount)
    For i = 1 To nCount
        aszTemp(i) = FormatDbValue(rsTemp!combine_serial)
        rsTemp.MoveNext
    Next i
    
    GetUniqueCombine = aszTemp
    
End Function


Public Function GetCombineSimply(pdyStartDate As Date, pdyEndDate As Date, pnSerial As Integer, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszSellStationID As String = "") As Recordset
    '按照车次组合查询公司简报
    Dim szSql As String
    Dim szWhere As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    
    If pnExtraStatus <> TP_AllSold Then
    '如果不是所有的状态
        szWhere = szWhere & " AND extra_status = " & pnExtraStatus
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
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetCombineSimply
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
         
    
        szSql = "   SELECT MAX(c.transport_company_id ) split_company_id,b.transport_company_id , MAX(t.transport_company_short_name) split_company_name , MAX(t2.transport_company_short_name) transport_company_name " _
            & "    , SUM(base_price) base_price , SUM (price_item_1) price_item_1, SUM(price_item_2) price_item_2   " _
            & "    , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 , SUM(price_item_5) price_item_5  " _
            & "    , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 ,SUM(price_item_8) price_item_8  " _
            & "    , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 " _
            & "    , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 " _
            & "    , SUM(price_item_15) price_item_15 " _
            & GetTotalItem
            '& "    , SUM(base_price) + SUM(price_item_1) + SUM(price_item_2)  + SUM(price_item_3)  + SUM(price_item_7)  + SUM(price_item_12) + SUM(price_item_13) Some_Sum " _
            & "    , SUM(base_price) + SUM(price_item_1) + SUM(price_item_12) + SUM(price_item_13) Base_Price2 " _
            & "    , SUM(price_item_4) + SUM(price_item_5)  + SUM(price_item_6)  + SUM(price_item_8)  + SUM(price_item_9) + SUM(price_item_10) + SUM(price_item_11) + SUM(price_item_14)+ SUM(price_item_15)   Some_Sum2 "
            
        szSql = szSql _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number) - SUM(ticket_return_number) ) AS passenger_number " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) ) AS ticket_price " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price  " _
            & "    , SUM(ticket_cancel_number) ticket_cancel_number , SUM(ticket_cancel_amount) ticket_cancel_amount " _
            & "    , SUM(ticket_return_number) ticket_return_number , SUM(ticket_return_amount) ticket_return_amount " _
            & "    , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_number) ticket_change_number " _
            & "    , SUM(ticket_change_amount) ticket_change_amount , SUM(ticket_change_charge) ticket_change_charge " _
            & " FROM  Stat_bus_sell_lst b , Company_info t , bus_assemble_info c , Company_info t2 " _
            & " WHERE date >= '" & ToDBDateTime(pdyStartDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdyEndDate)) & "'" _
            & " AND combine_serial =  " & pnSerial _
            & " AND c.transport_company_id = t.transport_company_id AND b.transport_company_id = t2.transport_company_id " _
            & " AND b.bus_id >= c.start_bus_id  AND b.bus_id <=  c.end_bus_id  " _
            & szWhere _
            & " GROUP BY c.transport_company_id , b.transport_company_id order by c.transport_company_id, b.transport_company_id  "

    
    Set rsTemp = oDB.Execute(szSql)
    
    Set GetCombineSimply = rsTemp
    
End Function




'公司流量简报统计(按售票)
Public Function GetCompanyFloatSimpleCon(pszCompanyID As String, pdyStartDate As Date, pdyEndDate As Date, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszSellStationID As String = "") As Recordset

    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetCompanyFloatSimpleCon
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    Dim szAddPreFix As String
    Dim oSystemParam As New SystemParam
    
    
    oSystemParam.Init m_oActiveUser
    szAddPreFix = oSystemParam.AdditionBusPreFix
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    If pszCompanyID <> "" Then
        szWhere = " AND b.transport_company_id IN (" & pszCompanyID & ")"
    End If
    If pnExtraStatus <> TP_AllSold Then
    '如果不是所有的状态
        szWhere = szWhere & " AND extra_status = " & pnExtraStatus
        
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
        szSql = "   SELECT a.* , ISNULL(b.Add_bus_count,0) Add_bus_count  , ISNULL(add_passenger_number,0) add_passenger_number" _
            & " , ISNULL(add_fact_float,0) add_fact_float , ISNULL(add_total_float,0) add_total_float " _
            & " , ISNULL(add_total_number,0) add_total_number , ISNULL(add_full_seat_rate,0) add_full_seat_rate " _
            & " , ISNULL(add_fact_load_rate,0) add_fact_load_rate  " _
            & " FROM (  " _
            & "   SELECT MAX(b.transport_company_id) split_company_id , MAX(transport_company_short_name) split_company_name " _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number)-SUM(ticket_return_number)- SUM(ticket_change_number)) AS passenger_number " _
            & "    , COUNT(bus_id) AS bus_count , SUM( fact_float ) AS fact_float , SUM( total_float ) AS total_float " _
            & "    , SUM(total_seat) AS total_number   " _
            & "    , CASE WHEN SUM(total_seat)<>0 THEN (SUM(passenger_number) -SUM(ticket_cancel_number)-SUM(ticket_return_number)-SUM(ticket_change_number)) * 100 / CONVERT(numeric(9,1),SUM(total_seat)) ELSE 0 END  AS full_seat_rate  " _
            & "    , CASE WHEN SUM(total_float)<>0 THEN  SUM( fact_float ) * 100 / SUM( total_float ) ELSE 0 END AS fact_load_rate " _
            & " FROM  stat_bus_startoff_lst b , Company_info t " _
            & " WHERE date >= '" & ToDBDateTime(pdyStartDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdyEndDate)) & "'" _
            & " AND b.transport_company_id *= t.transport_company_id " & szWhere _
            & " GROUP BY b.transport_company_id " _
            & " ) a ,  "
            
        szSql = szSql & " ( " _
            & " SELECT MAX(transport_company_id ) split_company_id , COUNT(bus_id) AS Add_bus_count    " _
            & " , (SUM(passenger_number) -SUM(ticket_cancel_number)-SUM(ticket_return_number)- SUM(ticket_change_number)) AS add_passenger_number " _
            & " , SUM( fact_float ) AS add_fact_float , SUM( total_float ) AS add_total_float " _
            & " , SUM(total_seat) AS add_total_number " _
            & " , CASE WHEN SUM(total_seat)<>0 THEN (SUM(passenger_number) -SUM(ticket_cancel_number)-SUM(ticket_return_number)- SUM(ticket_change_number)) * 100 / CONVERT(numeric(9,1),SUM(total_seat)) ELSE 0 END  AS add_full_seat_rate " _
            & " , CASE WHEN SUM(total_float)<>0 THEN  SUM( fact_float ) * 100 / SUM( total_float ) ELSE 0 END AS add_fact_load_rate " _
            & " FROM stat_bus_startoff_lst "
            
    szSql = szSql & " WHERE date >= '" & ToDBDateTime(pdyStartDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdyEndDate)) & "'" _
            & " AND bus_id LIKE '" & szAddPreFix & "' " _
            & " GROUP BY transport_company_id " _
            & " ) b " _
            & " WHERE a.split_company_id *= b.split_company_id "
            
    Set rsTemp = oDB.Execute(szSql)
    Set GetCompanyFloatSimpleCon = rsTemp
    
End Function

'公司流量简报统计(按检票)
Public Function GetCompanyFloatSimpleConCheck(pszCompanyID As String, pdyStartDate As Date, pdyEndDate As Date, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszSellStationID As String = "") As Recordset

    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetCompanyFloatSimpleCon
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    Dim szAddPreFix As String
    Dim oSystemParam As New SystemParam
    
    
    oSystemParam.Init m_oActiveUser
    szAddPreFix = oSystemParam.AdditionBusPreFix
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    If pszCompanyID <> "" Then
        szWhere = " AND b.transport_company_id IN (" & pszCompanyID & ")"
    End If
    If pnExtraStatus <> TP_AllSold Then
    '如果不是所有的状态
        szWhere = szWhere & " AND extra_status = " & pnExtraStatus
        
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
        szSql = "   SELECT a.* , ISNULL(b.Add_bus_count,0) Add_bus_count  , ISNULL(add_passenger_number,0) add_passenger_number" _
            & " , ISNULL(add_fact_float,0) add_fact_float , ISNULL(add_total_float,0) add_total_float " _
            & " , ISNULL(add_total_number,0) add_total_number , ISNULL(add_full_seat_rate,0) add_full_seat_rate " _
            & " , ISNULL(add_fact_load_rate,0) add_fact_load_rate  " _
            & " FROM (  " _
            & "   SELECT MAX(b.transport_company_id) split_company_id , MAX(transport_company_short_name) split_company_name " _
            & "    , SUM(passenger_number) AS passenger_number " _
            & "    , COUNT(bus_id) AS bus_count , SUM( fact_float ) AS fact_float , SUM( total_float ) AS total_float " _
            & "    , SUM(total_seat) AS total_number   " _
            & "    , CASE WHEN SUM(total_seat)<>0 THEN SUM(passenger_number) * 100 / CONVERT(numeric(9,1),SUM(total_seat)) ELSE 0 END  AS full_seat_rate  " _
            & "    , CASE WHEN SUM(total_float)<>0 THEN  SUM( fact_float ) * 100 / SUM( total_float ) ELSE 0 END AS fact_load_rate " _
            & " FROM  stat_bus_startoff_lst2 b , Company_info t " _
            & " WHERE date >= '" & ToDBDateTime(pdyStartDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdyEndDate)) & "'" _
            & " AND b.transport_company_id *= t.transport_company_id " & szWhere _
            & " GROUP BY b.transport_company_id " _
            & " ) a ,  "
            
        szSql = szSql & " ( " _
            & " SELECT MAX(transport_company_id ) split_company_id , COUNT(bus_id) AS Add_bus_count    " _
            & " , SUM(passenger_number) AS add_passenger_number " _
            & " , SUM( fact_float ) AS add_fact_float , SUM( total_float ) AS add_total_float " _
            & " , SUM(total_seat) AS add_total_number " _
            & " , CASE WHEN SUM(total_seat)<>0 THEN SUM(passenger_number) * 100 / CONVERT(numeric(9,1),SUM(total_seat)) ELSE 0 END  AS add_full_seat_rate " _
            & " , CASE WHEN SUM(total_float)<>0 THEN  SUM( fact_float ) * 100 / SUM( total_float ) ELSE 0 END AS add_fact_load_rate " _
            & " FROM stat_bus_startoff_lst2 "
            
    szSql = szSql & " WHERE date >= '" & ToDBDateTime(pdyStartDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdyEndDate)) & "'" _
            & " AND bus_id LIKE '" & szAddPreFix & "' " _
            & " GROUP BY transport_company_id " _
            & " ) b " _
            & " WHERE a.split_company_id *= b.split_company_id "
            
    Set rsTemp = oDB.Execute(szSql)
    Set GetCompanyFloatSimpleConCheck = rsTemp
    
End Function

'按照车次组合统计流量简报(按售票)
Public Function GetCombineFloatSimpleCon(pdyStartDate As Date, pdyEndDate As Date, pnSerial As Integer, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszSellStationID As String = "") As Recordset

    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetCombineFloatSimpleCon
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    Dim szAddPreFix As String
    Dim oSystemParam As New SystemParam
    
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    
    oSystemParam.Init m_oActiveUser
    szAddPreFix = oSystemParam.AdditionBusPreFix
    
    If pnExtraStatus <> TP_AllSold Then
    '如果不是所有的状态
        szWhere = szWhere & " AND extra_status = " & pnExtraStatus
        
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
szSql = "   SELECT a.* , ISNULL(b.Add_bus_count,0) Add_bus_count  , ISNULL(add_passenger_number,0) add_passenger_number" _
            & " , ISNULL(add_fact_float,0) add_fact_float , ISNULL(add_total_float,0) add_total_float " _
            & " , ISNULL(add_total_number,0) add_total_number , ISNULL(add_full_seat_rate,0) add_full_seat_rate " _
            & " , ISNULL(add_fact_load_rate,0) add_fact_load_rate  " _
            & " FROM (  " _
            & " SELECT MAX(c.transport_company_id ) split_company_id, MAX(transport_company_short_name) split_company_name " _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number)-SUM(ticket_return_number)- SUM(ticket_change_number)) AS passenger_number " _
            & "    , COUNT(bus_id) AS bus_count , SUM( fact_float ) AS fact_float , SUM( total_float ) AS total_float " _
            & "    , SUM(total_seat) AS total_number   " _
            & "    , CASE WHEN SUM(total_seat)<>0 THEN (SUM(passenger_number) -SUM(ticket_cancel_number)-SUM(ticket_return_number)- SUM(ticket_change_number)) * 100 / CONVERT(numeric(9,1),SUM(total_seat)) ELSE 0 END  AS full_seat_rate      " _
            & "    , CASE WHEN SUM(total_float)<>0 THEN  SUM( fact_float ) * 100 / SUM( total_float ) ELSE 0 END AS fact_load_rate " _
            & " FROM  stat_bus_startoff_lst b , Company_info t , bus_assemble_info c " _
            & " WHERE date >= '" & ToDBDateTime(pdyStartDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdyEndDate)) & "'" _
            & " AND combine_serial =  " & pnSerial _
            & " AND c.transport_company_id = t.transport_company_id " _
            & " AND b.bus_id >= c.start_bus_id  AND b.bus_id <=  c.end_bus_id  " _
            & " GROUP BY c.transport_company_id  " _
            & " ) a ,  "
        szSql = szSql & " ( " _
            & " SELECT MAX(c.transport_company_id ) split_company_id , COUNT(bus_id) AS Add_bus_count    " _
            & " , (SUM(passenger_number) -SUM(ticket_cancel_number)-SUM(ticket_return_number)- SUM(ticket_change_number)) AS add_passenger_number " _
            & " , SUM( fact_float ) AS add_fact_float , SUM( total_float ) AS add_total_float " _
            & " , SUM(total_seat) AS add_total_number " _
            & " , CASE WHEN SUM(total_seat)<>0 THEN (SUM(passenger_number) -SUM(ticket_cancel_number)-SUM(ticket_return_number)- SUM(ticket_change_number)) * 100 / CONVERT(numeric(9,1),SUM(total_seat)) ELSE 0 END  AS add_full_seat_rate " _
            & " , CASE WHEN SUM(total_float)<>0 THEN  SUM( fact_float ) * 100 / SUM( total_float ) ELSE 0 END AS add_fact_load_rate " _
            & " FROM stat_bus_startoff_lst  b , bus_assemble_info c "
            
            
    szSql = szSql & " WHERE date >= '" & ToDBDateTime(pdyStartDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdyEndDate)) & "'" _
            & " AND combine_serial =  " & pnSerial _
            & " AND b.bus_id >= c.start_bus_id  AND b.bus_id <=  c.end_bus_id " _
            & " AND bus_id LIKE '" & szAddPreFix & "' " _
            & " GROUP BY c.transport_company_id " _
            & " ) b " _
            & " WHERE a.split_company_id *= b.split_company_id "
            
            
    Set rsTemp = oDB.Execute(szSql)
    Set GetCombineFloatSimpleCon = rsTemp
    
End Function

'按照车次组合统计流量简报(按检票)
Public Function GetCombineFloatSimpleConCheck(pdyStartDate As Date, pdyEndDate As Date, pnSerial As Integer, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszSellStationID As String = "") As Recordset

    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetCombineFloatSimpleCon
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    Dim szAddPreFix As String
    Dim oSystemParam As New SystemParam
    
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    
    oSystemParam.Init m_oActiveUser
    szAddPreFix = oSystemParam.AdditionBusPreFix
    
    If pnExtraStatus <> TP_AllSold Then
    '如果不是所有的状态
        szWhere = szWhere & " AND extra_status = " & pnExtraStatus
        
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
szSql = "   SELECT a.* , ISNULL(b.Add_bus_count,0) Add_bus_count  , ISNULL(add_passenger_number,0) add_passenger_number" _
            & " , ISNULL(add_fact_float,0) add_fact_float , ISNULL(add_total_float,0) add_total_float " _
            & " , ISNULL(add_total_number,0) add_total_number , ISNULL(add_full_seat_rate,0) add_full_seat_rate " _
            & " , ISNULL(add_fact_load_rate,0) add_fact_load_rate  " _
            & " FROM (  " _
            & " SELECT MAX(c.transport_company_id ) split_company_id, MAX(transport_company_short_name) split_company_name " _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number)-SUM(ticket_return_number)- SUM(ticket_change_number)) AS passenger_number " _
            & "    , COUNT(bus_id) AS bus_count , SUM( fact_float ) AS fact_float , SUM( total_float ) AS total_float " _
            & "    , SUM(total_seat) AS total_number   " _
            & "    , CASE WHEN SUM(total_seat)<>0 THEN (SUM(passenger_number) -SUM(ticket_cancel_number)-SUM(ticket_return_number)- SUM(ticket_change_number)) * 100 / CONVERT(numeric(9,1),SUM(total_seat)) ELSE 0 END  AS full_seat_rate      " _
            & "    , CASE WHEN SUM(total_float)<>0 THEN  SUM( fact_float ) * 100 / SUM( total_float ) ELSE 0 END AS fact_load_rate " _
            & " FROM  stat_bus_startoff_lst b , Company_info t , bus_assemble_info c " _
            & " WHERE date >= '" & ToDBDateTime(pdyStartDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdyEndDate)) & "'" _
            & " AND combine_serial =  " & pnSerial _
            & " AND c.transport_company_id = t.transport_company_id " _
            & " AND b.bus_id >= c.start_bus_id  AND b.bus_id <=  c.end_bus_id  " _
            & " GROUP BY c.transport_company_id  " _
            & " ) a ,  "
        szSql = szSql & " ( " _
            & " SELECT MAX(c.transport_company_id ) split_company_id , COUNT(bus_id) AS Add_bus_count    " _
            & " , (SUM(passenger_number) -SUM(ticket_cancel_number)-SUM(ticket_return_number)- SUM(ticket_change_number)) AS add_passenger_number " _
            & " , SUM( fact_float ) AS add_fact_float , SUM( total_float ) AS add_total_float " _
            & " , SUM(total_seat) AS add_total_number " _
            & " , CASE WHEN SUM(total_seat)<>0 THEN (SUM(passenger_number) -SUM(ticket_cancel_number)-SUM(ticket_return_number)- SUM(ticket_change_number)) * 100 / CONVERT(numeric(9,1),SUM(total_seat)) ELSE 0 END  AS add_full_seat_rate " _
            & " , CASE WHEN SUM(total_float)<>0 THEN  SUM( fact_float ) * 100 / SUM( total_float ) ELSE 0 END AS add_fact_load_rate " _
            & " FROM stat_bus_startoff_lst  b , bus_assemble_info c "
            
            
    szSql = szSql & " WHERE date >= '" & ToDBDateTime(pdyStartDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdyEndDate)) & "'" _
            & " AND combine_serial =  " & pnSerial _
            & " AND b.bus_id >= c.start_bus_id  AND b.bus_id <=  c.end_bus_id " _
            & " AND bus_id LIKE '" & szAddPreFix & "' " _
            & " GROUP BY c.transport_company_id " _
            & " ) b " _
            & " WHERE a.split_company_id *= b.split_company_id "
            
            
    Set rsTemp = oDB.Execute(szSql)
    Set GetCombineFloatSimpleConCheck = rsTemp
    
End Function

'拆帐公司营收简报统计(按车次日期进行统计)
Public Function SplitCompanySimpleConByBusDate(pszCompanyID As String, pdtFromDate As Date, pdtToDate As Date, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszSellStationID As String = "") As Recordset

    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_SplitCompanySimpleConByBusDate
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    If pszCompanyID <> "" Then
        szWhere = " AND b.split_company_id IN (" & pszCompanyID & ")"
    End If
    If pnExtraStatus <> TP_AllSold Then
    '如果不是所有的状态
        szWhere = szWhere & " AND extra_status = " & pnExtraStatus
        
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
    szSql = "   SELECT b.split_company_id , b.transport_company_id  , MAX(t.transport_company_short_name) split_company_name , MAX(t2.transport_company_short_name) transport_company_name" _
            & "    , SUM(base_price) base_price , SUM (price_item_1) price_item_1, SUM(price_item_2) price_item_2   " _
            & "    , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 , SUM(price_item_5) price_item_5  " _
            & "    , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 ,SUM(price_item_8) price_item_8  " _
            & "    , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 " _
            & "    , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 " _
            & "    , SUM(price_item_15) price_item_15 " _
            & GetTotalItem
            '& "    , SUM(base_price) + SUM(price_item_1) + SUM(price_item_2)  + SUM(price_item_3)  + SUM(price_item_7)  + SUM(price_item_12) + SUM(price_item_13) Some_Sum " _
            & "    , SUM(base_price) + SUM(price_item_1) + SUM(price_item_12) + SUM(price_item_13) Base_Price2 " _
            & "    , SUM(price_item_4) + SUM(price_item_5)  + SUM(price_item_6)  + SUM(price_item_8)  + SUM(price_item_9) + SUM(price_item_10) + SUM(price_item_11) + SUM(price_item_14)+ SUM(price_item_15)   Some_Sum2 "
            
        szSql = szSql _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number) - SUM(ticket_return_number)- SUM(ticket_change_number) ) AS passenger_number " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) ) AS ticket_price " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price  " _
            & "    , SUM(ticket_cancel_number) ticket_cancel_number , SUM(ticket_cancel_amount) ticket_cancel_amount " _
            & "    , SUM(ticket_return_number) ticket_return_number , SUM(ticket_return_amount) ticket_return_amount " _
            & "    , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_number) ticket_change_number " _
            & "    , SUM(ticket_change_amount) ticket_change_amount , SUM(ticket_change_charge) ticket_change_charge " _
            & " FROM  stat_bus_startoff_lst b , Company_info t , Company_info t2  " _
            & " WHERE date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "'" _
            & " AND b.split_company_id = t.transport_company_id and b.transport_company_id = t2.transport_company_id " & szWhere _
            & " GROUP BY b.split_company_id , b.transport_company_id  ORDER BY b.split_company_id , b.transport_company_id "
    Set rsTemp = oDB.Execute(szSql)
    Set SplitCompanySimpleConByBusDate = rsTemp
    
End Function





Public Function GetCombineSimplyByBusDate(pdyStartDate As Date, pdyEndDate As Date, pnSerial As Integer, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszSellStationID As String = "") As Recordset
    '按照车次组合查询公司简报(按车次日期进行统计)
    Dim szSql As String
    Dim szWhere As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    
    If pnExtraStatus <> TP_AllSold Then
    '如果不是所有的状态
        szWhere = szWhere & " AND extra_status = " & pnExtraStatus
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
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetCombineSimplyByBusDate

    oDB.ConnectionString = GetConnectionStr(cszDss)
    
    
        szSql = "   SELECT MAX(c.transport_company_id ) split_company_id , b.transport_company_id , MAX(t.transport_company_short_name) split_company_name, MAX(t2.transport_company_short_name) transport_company_name " _
            & "    , SUM(base_price) base_price , SUM (price_item_1) price_item_1, SUM(price_item_2) price_item_2   " _
            & "    , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 , SUM(price_item_5) price_item_5  " _
            & "    , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 ,SUM(price_item_8) price_item_8  " _
            & "    , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 " _
            & "    , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 " _
            & "    , SUM(price_item_15) price_item_15 " _
            & GetTotalItem
            '& "    , SUM(base_price) + SUM(price_item_1) + SUM(price_item_2)  + SUM(price_item_3)  + SUM(price_item_7)  + SUM(price_item_12) + SUM(price_item_13) Some_Sum " _
            & "    , SUM(base_price) + SUM(price_item_1) + SUM(price_item_12) + SUM(price_item_13) Base_Price2 " _
            & "    , SUM(price_item_4) + SUM(price_item_5)  + SUM(price_item_6)  + SUM(price_item_8)  + SUM(price_item_9) + SUM(price_item_10) + SUM(price_item_11) + SUM(price_item_14)+ SUM(price_item_15)   Some_Sum2 "
            
        szSql = szSql _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number) - SUM(ticket_return_number) - SUM(ticket_change_number)) AS passenger_number " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) ) AS ticket_price " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price  " _
            & "    , SUM(ticket_cancel_number) ticket_cancel_number , SUM(ticket_cancel_amount) ticket_cancel_amount " _
            & "    , SUM(ticket_return_number) ticket_return_number , SUM(ticket_return_amount) ticket_return_amount " _
            & "    , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_number) ticket_change_number " _
            & "    , SUM(ticket_change_amount) ticket_change_amount , SUM(ticket_change_charge) ticket_change_charge " _
            & " FROM  stat_bus_startoff_lst b , Company_info t , bus_assemble_info c  , Company_info t2 " _
            & " WHERE date >= '" & ToDBDateTime(pdyStartDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdyEndDate)) & "'" _
            & " AND combine_serial =  " & pnSerial _
            & " AND c.transport_company_id = t.transport_company_id AND b.transport_company_id = t2.transport_company_id " _
            & " AND b.bus_id >= c.start_bus_id  AND b.bus_id <=  c.end_bus_id  " _
            & szWhere _
            & " GROUP BY c.transport_company_id, b.transport_company_id order by c.transport_company_id , b.transport_company_id "

    
    Set rsTemp = oDB.Execute(szSql)
    
    Set GetCombineSimplyByBusDate = rsTemp
    
End Function




'得到小计项1,小计项2及运价小计项
Private Function GetTotalItem() As String
    Dim oParam As New SystemParam
    Dim szSubTotalItem As String
    Dim aszSubTotalItem() As String
    Dim nSubTotalItem As Integer
    
    Dim szBasePriceTotalItem As String
    Dim aszBasePriceTotalItem() As String
    Dim nBasePriceTotalItem As Integer
    
    Dim szSum As String
    Dim nValue As Integer
    
    Dim i As Integer
    Dim j As Integer
   
    
    oParam.Init m_oActiveUser
    '取得小计项参数
    szSubTotalItem = oParam.SubTotalItem1
    aszSubTotalItem = StringToTeam(szSubTotalItem)
    nSubTotalItem = ArrayLength(aszSubTotalItem)
    
    '取得基本运价合计项
    szBasePriceTotalItem = oParam.BasePriceTotalItem
    aszBasePriceTotalItem = StringToTeam(szBasePriceTotalItem)
    nBasePriceTotalItem = ArrayLength(aszBasePriceTotalItem)
    
    '得到 小计项 字符串
    szSum = " , "
    For i = 1 To nSubTotalItem
        nValue = CInt(aszSubTotalItem(i))
        
        If nValue <> 0 Then
            szSum = szSum & IIf(i <> 1, "+", "") & "SUM(price_item_" & nValue & ")"
        Else
            szSum = szSum & IIf(i <> 1, "+", "") & "SUM(base_price)"
        End If
    Next i
    '
    szSum = szSum & " AS Some_Sum  , "
    '得到基本运价合计项字符串
    For i = 1 To nBasePriceTotalItem
        nValue = CInt(aszBasePriceTotalItem(i))
        
        If nValue <> 0 Then
            szSum = szSum & IIf(szSum <> "", "+", "") & "SUM(price_item_" & nValue & ")"
        Else
            szSum = szSum & IIf(i <> 1, "+", "") & "SUM(base_price)"
        End If
    Next i
    
    szSum = szSum & " AS Base_Price2  , "
    
    For i = 0 To 15
        For j = 1 To nSubTotalItem
            If i = CInt(aszSubTotalItem(j)) Then
                Exit For
            End If
        Next j
        If j = nSubTotalItem + 1 Then
            If i <> 0 Then
                szSum = szSum & IIf(Right(szSum, 2) <> ", ", "+", "") & "SUM(price_item_" & i & ")"
            Else
                szSum = szSum & IIf(Right(szSum, 2) <> ", ", "+", "") & "SUM(base_price)"
            End If
        
        End If
    Next i
    
    szSum = szSum & " AS Some_Sum2 "
    
    GetTotalItem = szSum
    
    
End Function





'修改统计表中的参运公司及拆帐公司
Public Function ModifyCompanyName(dtStartDate As Date, szCompanyID As String, szCompanyName As String, szStationID As String, Optional szBusid As String, Optional bflg As Boolean, Optional dtEndDate As Date)
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_ModifyCompanyName
    
    Dim rsTemp As Recordset
    Dim rsMyTemp As New Recordset
    Dim szSql As String
    Dim szSqlText As String
    Dim szDateFilds As String
    Dim dtDate As String
    Dim szMyDate As String
    Dim szMyBusID As String
    Dim oDB As New RTConnection
    Dim i As Integer
    oDB.ConnectionString = GetConnectionStr(cszDss)
    If DateDiff("d", dtStartDate, dtEndDate) > 0 Then
     szDateFilds = " bus_date > = '" & ToDBDate(dtStartDate) & "' AND bus_date <= '" & ToDBDate(dtEndDate) & "'"
     dtDate = "将在时间" & dtStartDate & "和" & dtEndDate & "之间"
     szMyDate = "date > = '" & ToDBDate(dtStartDate) & "' AND date <= '" & ToDBDate(dtEndDate) & "'"
    Else
     szDateFilds = " bus_date > = '" & ToDBDate(dtStartDate) & "'"
     dtDate = "将时间不小于" & dtStartDate
     szMyDate = "date > = '" & ToDBDate(dtStartDate) & "'"
    End If
    
    szSql = " update Stat_bus_sell_lst  " _
            & "set Stat_bus_sell_lst.transport_company_id = '" & szCompanyID & "' ,transport_company_name='" & szCompanyName & "'" _
            & " FROM Stat_bus_sell_lst   " _
            & "WHERE BUS_ID IN   " _
            & "( "
     szSqlText = "SELECT distinct bus_id FROM Env_bus_info WHERE end_station_id = '" & szStationID & "' AND  "
   If szBusid <> "" Then
    If bflg = True Then
      szBusid = szBusid & "%"
      szSqlText = szSqlText & " bus_id LIKE '" & szBusid & "' AND "
      szMyBusID = "且与" & szBusid & "代号相似的车次所有车次"
    Else
      szSqlText = szSqlText & " bus_id ='" & szBusid & "' AND   "
      szMyBusID = "且与车次代号为" & szBusid & "所有车次"
      
    End If
   End If
   szSqlText = szSqlText & szDateFilds
   
   Set rsTemp = oDB.Execute(szSqlText)
   If rsTemp.RecordCount = 0 Then
    Set oDB = Nothing
    Set rsTemp = Nothing
    ShowError RIGHT_ModifyCompanyName
   End If
   
   szSql = szSql & szSqlText & ") AND " & szMyDate
   
   On Error GoTo here
   oDB.BeginTrans
   oDB.Execute szSql
   oDB.CommitTrans
   WriteOperateLog m_oActiveUser, RIGHT_ModifyCompanyName, dtDate & "终点站为" & szStationID & szMyBusID & "参运公司代号修改为" & szCompanyID
   Exit Function
here:
   oDB.RollbackTrans
   WriteOperateLog m_oActiveUser, RIGHT_ModifyCompanyName, dtDate & "终点站为" & szStationID & szMyBusID & "参运公司代号修改为" & szCompanyID & "操作失败"
End Function


