VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TicketBusDim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private m_oActiveUser As ActiveUser

Public Enum eExtraStatus
    TP_ISExtra = 2
    TP_AllSold = 3
    TP_ISNormal = 1
End Enum

Public Type tBusCompanyCombineInfo
    CombineSerial As Integer '组合序号
    CombineName As String '组合名称
    StartBusId As String '起始车次代码
    EndBusID As String '结束车次代码
    TransportCompanyID As String '参运公司代码
    TransportCompanyName As String '名称
End Type

Public Enum EErrSellerFinance
    ERR_SellerFinanceNoActiveUser = ERR_SellerFinance + ERR_NoActiveUser
    ERR_CompanyNotExist = ERR_SellerFinance + 12 '29212
End Enum
'权限
Public Enum ERightSellerFinance
    RIGHT_QuerySellerEveryDaySellDetail = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 1 '查询售票员每日结算
    RIGHT_ModifyCompanyName = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 2 '修改参运公司
    RIGHT_GetBusStatByBusDate = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 3 '车次统计(按车次日期进行汇总)
    RIGHT_GetBusStatBySaleTime = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 4 '按售票日期汇总,得到车次统计
    RIGHT_GetBusStationTickets = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 5 '按售票车站,得到车次统计
    RIGHT_GetBusStationTicketsCount = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 6 '按售票车站汇总,得到车次统计
    RIGHT_GetBusTransStat = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 7 '得到指定条件的车次的运量
    RIGHT_GetCombineBusSimply = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 8 '按照车次组合查询公司简报(按车次列出)
    RIGHT_GetCombineBusSimplyByBusDate = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 9 '根据车次日期统计,指定车次组合内的车次的统计信息
    RIGHT_GetStationTickets = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 10 '按照车次站点票统计信息
    RIGHT_GetStationTicketsCount = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 11 '按照车次站点售票汇兑统计信息
    'RIGHT_GetTotalItem = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 12
    
    
    
End Enum


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

'按售票车站,得到车次统计
Public Function GetBusStationTickets(BusDate() As Date, BusId() As String, nORderCol As Integer) As Recordset
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusStationTickets
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim nCount As Integer
    Dim szWhere As String
    Dim szOrder As String
    Dim i As Integer
    
    nCount = ArrayLength(BusId)
    oDB.ConnectionString = GetConnectionStr(cszDss)
    szSql = "SELECT sst.*,ebd.mileage,sfo.station_name,tfo.ticket_type_name " & _
          " FROM Ticket_sell_lst sst,Env_bus_station_lst ebd,station_info sfo,Ticket_type_code tfo  "
          
    szWhere = " WHERE ebd.station_id=sst.des_station_id AND ebd.bus_date=sst.bus_date AND (( " & ST_TicketNormal & "|" & ST_TicketChecked & ") & status)=status  " & _
                 " AND ebd.bus_id=sst.bus_id AND sfo.station_id=sst.des_station_id AND tfo.ticket_type_id=sst.ticket_type AND sst.sell_station_id = EBD.sell_station_id "

    If nCount > 0 Then szWhere = szWhere & "AND ("
    For i = 1 To nCount
        If i = 1 Then
           szWhere = szWhere & " (sst.bus_date='" & Format(BusDate(i), "YYYY-MM-DD") & "' AND sst.bus_id='" & BusId(i) & "') "
        Else
           szWhere = szWhere & " OR (sst.bus_date='" & Format(BusDate(i), "YYYY-MM-dd") & "' AND sst.bus_id='" & BusId(i) & "') "
        End If
    Next i
    If nCount > 0 Then szWhere = szWhere & ") "
    
    szOrder = " ORDER BY sst.user_id ,sst.bus_date ,sst.bus_id,"
    If nORderCol = 0 Then
       szOrder = szOrder & "sst.ticket_id,ebd.mileage,sst.seat_no"
    ElseIf nORderCol = 1 Then
       szOrder = szOrder & "ebd.mileage,sst.ticket_id,sst.seat_no"
    ElseIf nORderCol = 2 Then
       szOrder = szOrder & "sst.seat_no,sst.ticket_id,ebd.mileage"
    End If
    szSql = szSql & szWhere & szOrder
    Set rsTemp = oDB.Execute(szSql)
    Set GetBusStationTickets = rsTemp
    Set rsTemp = Nothing
    Set oDB = Nothing
End Function

Public Function GetBusStationTicketsCount(BusDate() As Date, BusId() As String, nORderCol As Integer) As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusStationTicketsCount
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim nCount As Integer
    Dim szWhere As String
    Dim szOrder As String
    Dim szSqlOther As String
    Dim i As Integer
    
    nCount = ArrayLength(BusId)
    oDB.ConnectionString = GetConnectionStr(cszDss)
    szSql = "SELECT sst.*,tfo.ticket_type_name,sfo.station_name " & _
            " FROM  Ticket_type_code tfo,station_info sfo," & _
            " (SELECT bus_date,bus_id,des_station_id,ticket_type,COUNT(ticket_id) as TicketNumber," & _
            " SUM(ticket_price) As TicketMoneyTotal " & _
            " FROM Ticket_sell_lst " & _
            " WHERE ((" & ST_TicketNormal & "|" & ST_TicketChecked & ") & status)=status "
    szSqlOther = " GROUP BY bus_date,bus_id,des_station_id,ticket_type) sst "
              
    If nCount > 0 Then szWhere = " AND ("
    For i = 1 To nCount
        If i = 1 Then
           szWhere = szWhere & " (bus_date='" & Format(BusDate(i), "YYYY-MM-DD") & "' AND bus_id='" & BusId(i) & "') "
        Else
           szWhere = szWhere & " OR (bus_date='" & Format(BusDate(i), "YYYY-MM-dd") & "' AND bus_id='" & BusId(i) & "') "
        End If
    Next i
    If nCount > 0 Then szWhere = szWhere & ") "
    
    szSql = szSql & szWhere & szSqlOther

    If nORderCol = 0 Then
       szSql = "SELECT bss.bus_id,bss.ticket_type_name,bss.station_name," & _
               " SUM(bss.ticketNumber) as ticketNumber,SUM(bss.ticketMoneyTotal) as ticketMoneyTotal " & _
               " FROM (" & szSql & "" & _
               " WHERE tfo.ticket_type_id=sst.ticket_type AND sfo.station_id=sst.des_station_id) bss " & _
               " GROUP BY bss.bus_id,bss.station_name,bss.ticket_type_name " & _
               " ORDER BY bus_id "
    ElseIf nORderCol = 1 Then
        szSql = "SELECT bss.bus_date,bss.ticket_type_name,bss.station_name, " & _
                " SUM(bss.ticketNumber) as ticketNumber,SUM(bss.ticketMoneyTotal) as ticketMoneyTotal " & _
                " FROM (" & szSql & "" & _
                " WHERE tfo.ticket_type_id=sst.ticket_type AND sfo.station_id=sst.des_station_id) bss " & _
                " GROUP BY bss.bus_date,bss.station_name,bss.ticket_type_name " & _
                " ORDER BY bus_date "
    ElseIf nORderCol = 2 Then
       szSql = szSql & " WHERE tfo.ticket_type_id=sst.ticket_type AND sfo.station_id=sst.des_station_id  ORDER BY bus_date,bus_id "
    End If
    Set rsTemp = oDB.Execute(szSql)
    Set GetBusStationTicketsCount = rsTemp
    Set rsTemp = Nothing
    Set oDB = Nothing
End Function
'按照车次站点票统计信息
Public Function GetStationTickets(StartBusDate As Date, EndBusDate As Date, StationId() As String, nORderCol As Integer) As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetStationTickets
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim nCount As Integer
    Dim szWhere As String
    Dim szOrder As String
    Dim i As Integer
    
    nCount = ArrayLength(StationId)
    oDB.ConnectionString = GetConnectionStr(cszDss)
    szSql = "SELECT sst.*,ebd.mileage,sfo.station_name,tfo.ticket_type_name " & _
          " FROM Ticket_sell_lst sst,Env_bus_station_lst ebd,station_info sfo,Ticket_type_code tfo  "
          
    szWhere = " WHERE ((" & ST_TicketNormal & "|" & ST_TicketChecked & ") & sst.status)=sst.status AND " & " ebd.station_id=sst.des_station_id AND ebd.bus_date=sst.bus_date " & _
                 " AND ebd.bus_id=sst.bus_id AND sfo.station_id=sst.des_station_id AND tfo.ticket_type_id=sst.ticket_type AND sst.sell_station_id=ebd.sell_station_id " & _
                 " AND sst.bus_date>='" & Format(StartBusDate, "YYYY-MM-DD") & "' AND sst.bus_date<='" & Format(EndBusDate, "YYYY-MM-DD") & "'"
                 
    If nCount > 0 Then szWhere = szWhere & " AND ("
    For i = 1 To nCount
        If i = 1 Then
           szWhere = szWhere & "sst.des_station_id='" & StationId(i) & "' "
        Else
           szWhere = szWhere & " OR sst.des_station_id='" & StationId(i) & "' "
        End If
    Next i
    If nCount > 0 Then szWhere = szWhere & ") "
    
    szOrder = " ORDER BY sst.bus_date,sst.bus_id,"
    If nORderCol = 0 Then
       szOrder = szOrder & "sst.ticket_id,ebd.mileage,sst.seat_no"
    ElseIf nORderCol = 1 Then
       szOrder = szOrder & "ebd.mileage,sst.ticket_id,sst.seat_no"
    ElseIf nORderCol = 2 Then
       szOrder = szOrder & "sst.seat_no,sst.ticket_id,ebd.mileage"
    End If
    szSql = szSql & szWhere & szOrder
    Set rsTemp = oDB.Execute(szSql)
    Set GetStationTickets = rsTemp
    Set rsTemp = Nothing
    Set oDB = Nothing
End Function

Public Function GetStationTicketsCount(StartBusDate As Date, EndBusDate As Date, StationId() As String, nORderCol As Integer) As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetStationTicketsCount
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim nCount As Integer
    Dim szWhere As String
    Dim szOrder As String
    Dim szSqlOther As String
    Dim i As Integer
    nCount = ArrayLength(StationId)
    oDB.ConnectionString = GetConnectionStr(cszDss)
    szSql = "SELECT sst.*,tfo.ticket_type_name,sfo.station_name " & _
            " FROM  Ticket_type_code tfo,station_info sfo," & _
            " (SELECT bus_date,bus_id,des_station_id,ticket_type,COUNT(ticket_id) as TicketNumber," & _
            " SUM(ticket_price) As TicketMoneyTotal " & _
            " FROM Ticket_sell_lst " & _
            " WHERE ((" & ST_TicketNormal & "|" & ST_TicketChecked & ") & Status)=status  AND " & _
            " bus_date>='" & Format(StartBusDate, "YYYY-MM-DD") & "' AND  bus_date<='" & Format(EndBusDate, "YYYY-MM-DD") & "' "

    szSqlOther = " GROUP BY bus_date,bus_id,des_station_id,ticket_type) sst "
              
    If nCount > 0 Then szWhere = " AND  ("
    For i = 1 To nCount
        If i = 1 Then
           szWhere = szWhere & " des_station_id='" & StationId(i) & "' "
        Else
           szWhere = szWhere & " OR des_station_id='" & StationId(i) & "' "
        End If
    Next i
    If nCount > 0 Then szWhere = szWhere & ") "
    
    szSql = szSql & szWhere & szSqlOther

    If nORderCol = 0 Then
       szSql = "SELECT bss.bus_id,bss.ticket_type_name,bss.station_name," & _
               " SUM(bss.ticketNumber) as ticketNumber,SUM(bss.ticketMoneyTotal) as ticketMoneyTotal " & _
               " FROM (" & szSql & "" & _
               " WHERE tfo.ticket_type_id=sst.ticket_type AND sfo.station_id=sst.des_station_id) bss " & _
               " GROUP BY bss.bus_id,bss.station_name,bss.ticket_type_name " & _
               " ORDER BY bus_id "
    ElseIf nORderCol = 1 Then
        szSql = "SELECT bss.bus_date,bss.ticket_type_name,bss.station_name, " & _
                " SUM(bss.ticketNumber) as ticketNumber,SUM(bss.ticketMoneyTotal) as ticketMoneyTotal " & _
                " FROM (" & szSql & "" & _
                " WHERE tfo.ticket_type_id=sst.ticket_type AND sfo.station_id=sst.des_station_id) bss " & _
                " GROUP BY bss.bus_date,bss.station_name,bss.ticket_type_name " & _
                " ORDER BY bus_date "
    ElseIf nORderCol = 2 Then
       szSql = szSql & " WHERE tfo.ticket_type_id=sst.ticket_type AND sfo.station_id=sst.des_station_id  ORDER BY bus_date,bus_id "
    End If
    Set rsTemp = oDB.Execute(szSql)
    Set GetStationTicketsCount = rsTemp
    Set rsTemp = Nothing
    Set oDB = Nothing
End Function


'按售票日期汇总,得到车次统计
Public Function GetBusStatBySaleTime(pszCompanyID As String, pdtFromDate As Date, pdtToDate As Date, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszLike As String = "", Optional pszSellStationID As String = "") As Recordset
'pnExtraStatus =1 表示售票，2表示补票，3表示全部



    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusStatBySaleTime
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    If pszCompanyID <> "" Then 'FGAG
    If InStr(pszCompanyID, "'") <> 1 Then pszCompanyID = "'" + pszCompanyID + "'"
        szWhere = " AND split_company_id IN ( " & pszCompanyID & ")"
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
    If pszLike <> "" Then
        szWhere = szWhere & " AND bus_id like '" & pszLike & "%' "
    End If
    szSql = "   SELECT bus_id,b.transport_company_id,  MAX(bus_start_time) bus_start_time ,MAX(c.transport_company_short_name) transport_company_short_name " _
            & "    , MAX(B.route_id) route_id , MAX(route_name) route_name " _
            & "    , SUM(base_price) base_price , SUM (price_item_1) price_item_1, SUM(price_item_2) price_item_2   " _
            & "    , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 , SUM(price_item_5) price_item_5  " _
            & "    , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 ,SUM(price_item_8) price_item_8  " _
            & "    , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 " _
            & "    , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 " _
            & "    , SUM(price_item_15) price_item_15 " _
            & GetTotalItem
    
        szSql = szSql _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number)- SUM(ticket_return_number) - sum(ticket_change_number) ) AS passenger_number " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) ) AS ticket_price " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price  " _
            & "    , SUM(ticket_cancel_number) ticket_cancel_number , SUM(ticket_cancel_amount) ticket_cancel_amount " _
            & "    , SUM(ticket_return_number) ticket_return_number , SUM(ticket_return_amount) ticket_return_amount " _
            & "    , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_number) ticket_change_number " _
            & "    , SUM(ticket_change_amount) ticket_change_amount , SUM(ticket_change_charge) ticket_change_charge " _
            & " FROM  Stat_bus_sell_lst b , route_info r , company_info c " _
            & " WHERE r.route_id = b.route_id AND b.transport_company_id = c.transport_company_id AND date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "'" & szWhere & ExceptZeroBus _
            & " GROUP BY bus_id , b.transport_company_id ORDER BY bus_id , b.transport_company_id "
    
    Set rsTemp = oDB.Execute(szSql)
    Set GetBusStatBySaleTime = rsTemp
    
End Function




'车次统计(按车次日期进行汇总)
Public Function GetBusStatByBusDate(pszCompanyID As String, pdtFromDate As Date, pdtToDate As Date, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszLike As String = "", Optional pszSellStationID As String = "") As Recordset
'pnExtraStatus =1 表示售票，2表示补票，3表示全部



    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusStatByBusDate
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    If pszCompanyID <> "" Then
        szWhere = " AND split_company_id IN ( " & pszCompanyID & ")"
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
    If pszLike <> "" Then
        szWhere = szWhere & " AND bus_id like '" & pszLike & "%' "
    End If

    szSql = "   SELECT bus_id,b.transport_company_id ,MAX(c.transport_company_short_name) transport_company_short_name , MAX(bus_start_time) bus_start_time , B.route_id , MAX(route_name) route_name " _
            & "    , SUM(base_price) base_price , SUM (price_item_1) price_item_1, SUM(price_item_2) price_item_2   " _
            & "    , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 , SUM(price_item_5) price_item_5  " _
            & "    , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 ,SUM(price_item_8) price_item_8  " _
            & "    , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 " _
            & "    , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 " _
            & "    , SUM(price_item_15) price_item_15 " _
            & GetTotalItem
            
        szSql = szSql _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number) - SUM(ticket_return_number) - sum(ticket_change_number)  ) AS passenger_number " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) ) AS ticket_price " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price  " _
            & "    , SUM(ticket_cancel_number) ticket_cancel_number , SUM(ticket_cancel_amount) ticket_cancel_amount " _
            & "    , SUM(ticket_return_number) ticket_return_number , SUM(ticket_return_amount) ticket_return_amount " _
            & "    , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_number) ticket_change_number " _
            & "    , SUM(ticket_change_amount) ticket_change_amount , SUM(ticket_change_charge) ticket_change_charge " _
            & " FROM  stat_bus_startoff_lst b , route_info r , company_info c " _
            & " WHERE r.route_id = b.route_id AND b.transport_company_id = c.transport_company_id AND date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "'" & szWhere & ExceptZeroBus _
            & " GROUP BY bus_id,b.route_id,b.transport_company_id ORDER BY bus_id,b.route_id,b.transport_company_id "
    Set rsTemp = oDB.Execute(szSql)
    Set GetBusStatByBusDate = rsTemp

End Function



'车次统计(按车次日期及售票员所属车站进行汇总)
Public Function GetBusStatByBusDateAndSalerStation(pszCompanyID As String, pdtFromDate As Date, pdtToDate As Date, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszLike As String = "", Optional pszSellStationID As String = "") As Recordset
'pnExtraStatus =1 表示售票，2表示补票，3表示全部



    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusStatByBusDate
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    If pszCompanyID <> "" Then
        szWhere = " AND split_company_id IN ( " & pszCompanyID & ")"
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
'    If pnExtraStatus <> TP_AllSold Then
'    '如果不是所有的状态
'        szWhere = szWhere & " AND extra_status = " & pnExtraStatus
'
'    End If
    If pszLike <> "" Then
        szWhere = szWhere & " AND bus_id like '" & pszLike & "%' "
    End If

    szSql = "   SELECT bus_id, max(c.transport_company_short_name) transport_company_short_name,MAX(bus_start_time) bus_start_time , MAX(B.route_id) route_id , MAX(route_name) route_name " _
            & "    , SUM(base_price) base_price , SUM (price_item_1) price_item_1, SUM(price_item_2) price_item_2   " _
            & "    , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 , SUM(price_item_5) price_item_5  " _
            & "    , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 ,SUM(price_item_8) price_item_8  " _
            & "    , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 " _
            & "    , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 " _
            & "    , SUM(price_item_15) price_item_15 " _
            & GetTotalItem
            
        szSql = szSql _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number) - SUM(ticket_return_number) - sum(ticket_change_number)  ) AS passenger_number " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) ) AS ticket_price " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price  " _
            & "    , SUM(ticket_cancel_number) ticket_cancel_number , SUM(ticket_cancel_amount) ticket_cancel_amount " _
            & "    , SUM(ticket_return_number) ticket_return_number , SUM(ticket_return_amount) ticket_return_amount " _
            & "    , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_number) ticket_change_number " _
            & "    , SUM(ticket_change_amount) ticket_change_amount , SUM(ticket_change_charge) ticket_change_charge " _
            & " FROM  stat_bus_station_sell_lst b , route_info r , company_info c " _
            & " WHERE r.route_id = b.route_id AND c.transport_company_id=b.transport_company_id AND  date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "'" & szWhere & ExceptZeroBus _
            & " GROUP BY bus_id ORDER BY bus_id "
    Set rsTemp = oDB.Execute(szSql)
    Set GetBusStatByBusDateAndSalerStation = rsTemp
    
End Function




'根据车次日期统计,指定车次组合内的车次的统计信息
Public Function GetCombineBusSimplyByBusDate(pdtFromDate As Date, pdtToDate As Date, pnSerial As Integer, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszLike As String = "", Optional pszSellStationID As String = "") As Recordset
'pnExtraStatus =1 表示售票，2表示补票，3表示全部



    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetCombineBusSimplyByBusDate
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
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
    If pszLike <> "" Then
        szWhere = szWhere & " AND bus_id like '" & pszLike & "%' "
    End If

    szSql = "   SELECT bus_id, MAX(bus_start_time) bus_start_time  , MAX(transport_company_short_name) transport_company_short_name  " _
            & "    , MAX(B.route_id) route_id , MAX(route_name) route_name " _
            & "    , SUM(base_price) base_price , SUM (price_item_1) price_item_1, SUM(price_item_2) price_item_2   " _
            & "    , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 , SUM(price_item_5) price_item_5  " _
            & "    , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 ,SUM(price_item_8) price_item_8  " _
            & "    , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 " _
            & "    , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 " _
            & "    , SUM(price_item_15) price_item_15 " _
            & GetTotalItem
            
        szSql = szSql _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number) - SUM(ticket_return_number) - sum(ticket_change_number)  ) AS passenger_number " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) ) AS ticket_price " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price  " _
            & "    , SUM(ticket_cancel_number) ticket_cancel_number , SUM(ticket_cancel_amount) ticket_cancel_amount " _
            & "    , SUM(ticket_return_number) ticket_return_number , SUM(ticket_return_amount) ticket_return_amount " _
            & "    , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_number) ticket_change_number " _
            & "    , SUM(ticket_change_amount) ticket_change_amount , SUM(ticket_change_charge) ticket_change_charge " _
            & " FROM  stat_bus_startoff_lst b , company_info t , bus_assemble_info c , route_info r " _
            & " WHERE r.route_id = b.route_id AND date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "'" _
            & " AND combine_serial =  " & pnSerial _
            & " AND c.transport_company_id *= t.transport_company_id " _
            & " AND b.bus_id >= c.start_bus_id  AND b.bus_id <=  c.end_bus_id  " _
            & szWhere _
            & " GROUP BY B.bus_id order by B.BUS_ID "

    
    Set rsTemp = oDB.Execute(szSql)
    Set GetCombineBusSimplyByBusDate = rsTemp
    
End Function






'按照车次组合查询公司简报(按车次列出)
Public Function GetCombineBusSimply(pdtFromDate As Date, pdtToDate As Date, pnSerial As Integer, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszLike As String = "", Optional pszSellStationID As String = "") As Recordset
'pnExtraStatus =1 表示售票，2表示补票，3表示全部



    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetCombineBusSimply
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
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
    If pszLike <> "" Then
        szWhere = szWhere & " AND bus_id like '" & pszLike & "%' "
    End If
    szSql = "   SELECT bus_id, MAX(bus_start_time) bus_start_time  , MAX(transport_company_short_name) transport_company_short_name  " _
            & "    , MAX(B.route_id) route_id , MAX(route_name) route_name " _
            & "    , SUM(base_price) base_price , SUM (price_item_1) price_item_1, SUM(price_item_2) price_item_2   " _
            & "    , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 , SUM(price_item_5) price_item_5  " _
            & "    , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 ,SUM(price_item_8) price_item_8  " _
            & "    , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 " _
            & "    , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 " _
            & "    , SUM(price_item_15) price_item_15 " _
            & GetTotalItem

        szSql = szSql _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number) - SUM(ticket_return_number) - sum(ticket_change_number) ) AS passenger_number " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) ) AS ticket_price " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price  " _
            & "    , SUM(ticket_cancel_number) ticket_cancel_number , SUM(ticket_cancel_amount) ticket_cancel_amount " _
            & "    , SUM(ticket_return_number) ticket_return_number , SUM(ticket_return_amount) ticket_return_amount " _
            & "    , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_number) ticket_change_number " _
            & "    , SUM(ticket_change_amount) ticket_change_amount , SUM(ticket_change_charge) ticket_change_charge " _
            & " FROM  Stat_bus_sell_lst b , company_info t , bus_assemble_info c , route_info r " _
            & " WHERE r.route_id = b.route_id and date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "'" _
            & " AND combine_serial =  " & pnSerial _
            & " AND c.transport_company_id *= t.transport_company_id " _
            & " AND b.bus_id >= c.start_bus_id  AND b.bus_id <=  c.end_bus_id  " _
            & szWhere _
            & " GROUP BY B.bus_id order by B.BUS_ID "

    
    Set rsTemp = oDB.Execute(szSql)
    Set GetCombineBusSimply = rsTemp
    
End Function



'得到不是零的条件
Private Function ExceptZeroBus() As String

    

    ExceptZeroBus = "  AND NOT ( passenger_number  = 0  AND  ticket_price = 0" _
    & " AND ticket_cancel_number = 0 AND ticket_cancel_amount = 0 AND ticket_return_number = 0 AND ticket_return_amount = 0 " _
    & " AND  ticket_return_charge = 0 AND ticket_change_number = 0 AND  ticket_change_amount  = 0 AND ticket_change_charge  = 0) "
'    & " AND total_seat =0 AND fact_float =0 AND total_float =0 )  "
End Function


'得到指定条件的车次的运量(按售票)
Public Function GetBusTransStat(paszBusID() As String, Optional pdyStartDate As Date, Optional pdyEndDate As Date, Optional pszRouteID As String = "", Optional pszSellStationID As String = "") As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusTransStat
    
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szWhere As String
    Dim nCount As Integer
    Dim i As Integer
    Dim szCountStr As String '车次统计
    
    oDB.ConnectionString = GetConnectionStr
    nCount = ArrayLength(paszBusID)
    If nCount <> 0 Then
        szWhere = " AND bus_id IN ("
        For i = 1 To nCount - 1
            szWhere = szWhere & "'" & paszBusID(i) & "',"
        Next i
        szWhere = szWhere & "'" & paszBusID(i) & "')"
    End If
    If pszRouteID <> "" Then
        szWhere = szWhere & " AND b.route_id = '" & pszRouteID & "'"
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
    Dim nSellCount As Integer
    nSellCount = (DateDiff("d", pdyStartDate, DateAdd("d", 1, pdyEndDate))) * 2
    
    szCountStr = " count(distinct cast(date as char)+bus_id) bus_count "
    
    szSql = "SELECT bus_id , MAX(b.route_id) route_id , min(convert(char(5),bus_start_time ,108)) bus_start_time, " _
        & szCountStr _
        & " , MAX(r.route_name) route_name , (SUM(passenger_number) - SUM(ticket_cancel_number) - SUM(ticket_return_number)- SUM(ticket_change_number)) AS passenger_number " _
        & " , SUM( fact_float ) AS fact_float , SUM( total_float )/count(distinct sell_station_id) total_float " _
        & " , SUM(total_seat)/count(distinct sell_station_id) total_number " _
        & " , CASE WHEN SUM(total_seat)<>0 THEN (SUM(passenger_number)-SUM(ticket_cancel_number)-SUM(ticket_return_number)-SUM(ticket_change_number)) * 100 / CONVERT(numeric(9,1) " _
        & " , SUM(total_seat)/count(distinct sell_station_id)) " _
        & " ELSE 0 END  AS full_seat_rate " _
        & " , CASE WHEN SUM(total_float)<>0 then SUM( fact_float ) * 100 / (SUM( total_float )/count(distinct sell_station_id)) ELSE 0 END AS fact_load_rate " _
        & " , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
        & "     - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price " _
        & " FROM stat_bus_startoff_lst b , route_info r " _
        & " WHERE  b.route_id = r.route_id AND b.date >= '" & ToDBDate(pdyStartDate) & "' AND b.date < '" & ToDBDate(DateAdd("d", 1, pdyEndDate)) & "' " & szWhere _
        & " GROUP BY bus_id " _
        & " ORDER BY bus_id "
    
    
    
    Set rsTemp = oDB.Execute(szSql)
    Set GetBusTransStat = rsTemp
    
End Function

'得到指定条件的车次的运量(按检票)
Public Function GetBusTransStatByCheck(paszBusID() As String, Optional pdyStartDate As Date, Optional pdyEndDate As Date, Optional pszRouteID As String = "", Optional pszSellStationID As String = "") As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusTransStat
    
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szWhere As String
    Dim nCount As Integer
    Dim i As Integer
    
    oDB.ConnectionString = GetConnectionStr
    nCount = ArrayLength(paszBusID)
    If nCount <> 0 Then
        szWhere = " AND bus_id IN ("
        For i = 1 To nCount - 1
            szWhere = szWhere & "'" & paszBusID(i) & "',"
        Next i
        szWhere = szWhere & "'" & paszBusID(i) & "')"
    End If
    If pszRouteID <> "" Then
        szWhere = szWhere & " AND b.route_id = '" & pszRouteID & "'"
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
    
    szSql = "SELECT bus_id , MAX(b.route_id) route_id , MAX(convert(char(5),bus_start_time ,108)) bus_start_time , count(bus_serial_no) bus_count  " _
        & " , MAX(r.route_name) route_name , SUM(passenger_number) AS passenger_number " _
        & " , SUM( fact_float ) AS fact_float , SUM( total_float ) AS total_float  , SUM(total_seat) AS total_number " _
        & " , CASE WHEN SUM(total_seat)<>0 THEN SUM(passenger_number) * 100 / CONVERT(numeric(9,1) " _
        & " , SUM(total_seat)) ELSE 0 END  AS full_seat_rate " _
        & " , CASE WHEN SUM(total_float)<>0 THEN  SUM( fact_float ) * 100 / SUM( total_float ) ELSE 0 END AS fact_load_rate " _
        & " , SUM(ticket_price) AS total_ticket_price " _
        & " FROM stat_bus_startoff_lst2 b , route_info r " _
        & " WHERE  b.route_id = r.route_id AND b.date >= '" & ToDBDate(pdyStartDate) & "' AND b.date < '" & ToDBDate(DateAdd("d", 1, pdyEndDate)) & "' " & szWhere _
        & " GROUP BY bus_id " _
        & " ORDER BY bus_id "
    
    
    
    Set rsTemp = oDB.Execute(szSql)
    Set GetBusTransStatByCheck = rsTemp
    
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

'车次各上车站售票简报(按售票时间汇总)
Public Function GetBusSellStationSellInfoBySaleTime(pszCompanyID As String, pdtFromDate As Date, pdtToDate As Date, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszLike As String = "", Optional pszSellStationID As String = "") As Recordset
'pnExtraStatus =1 表示售票，2表示补票，3表示全部
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusStatBySaleTime
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    If pszCompanyID <> "" Then
        szWhere = " AND split_company_id IN ( " & pszCompanyID & ")"
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
    If pszLike <> "" Then
        szWhere = szWhere & " AND bus_id like '" & pszLike & "%' "
    End If
    szSql = "   SELECT sell_station_id , bus_id, MAX(bus_start_time) bus_start_time , MAX(B.route_id) route_id , MAX(route_name) route_name " _
            & "    , SUM(base_price) base_price , SUM (price_item_1) price_item_1, SUM(price_item_2) price_item_2   " _
            & "    , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 , SUM(price_item_5) price_item_5  " _
            & "    , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 ,SUM(price_item_8) price_item_8  " _
            & "    , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 " _
            & "    , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 " _
            & "    , SUM(price_item_15) price_item_15 " _
            & GetTotalItem
            
        szSql = szSql _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number)- SUM(ticket_return_number) - sum(ticket_change_number) ) AS passenger_number " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) ) AS ticket_price " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price  " _
            & "    , SUM(ticket_cancel_number) ticket_cancel_number , SUM(ticket_cancel_amount) ticket_cancel_amount " _
            & "    , SUM(ticket_return_number) ticket_return_number , SUM(ticket_return_amount) ticket_return_amount " _
            & "    , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_number) ticket_change_number " _
            & "    , SUM(ticket_change_amount) ticket_change_amount , SUM(ticket_change_charge) ticket_change_charge " _
            & " FROM  Stat_bus_sell_lst b , route_info r " _
            & " WHERE r.route_id = b.route_id AND  date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "'" & szWhere & ExceptZeroBus _
            & " GROUP BY bus_id,sell_station_id  ORDER BY bus_id,sell_station_id  "
            
    Set rsTemp = oDB.Execute(szSql)
    Set GetBusSellStationSellInfoBySaleTime = rsTemp
    
End Function


'车次各上车站售票简报(按发车时间汇总)
Public Function GetBusSellStationSellInfoByBusDate(pszCompanyID As String, pdtFromDate As Date, pdtToDate As Date, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszLike As String = "", Optional pszSellStationID As String = "") As Recordset
'pnExtraStatus =1 表示售票，2表示补票，3表示全部
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusStatByBusDate
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    If pszCompanyID <> "" Then
        szWhere = " AND split_company_id IN ( " & pszCompanyID & ")"
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
    If pszLike <> "" Then
        szWhere = szWhere & " AND bus_id like '" & pszLike & "%' "
    End If

    szSql = "   SELECT bus_id,sell_station_id , MAX(bus_start_time) bus_start_time , MAX(B.route_id) route_id , MAX(route_name) route_name " _
            & "    , SUM(base_price) base_price , SUM (price_item_1) price_item_1, SUM(price_item_2) price_item_2   " _
            & "    , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 , SUM(price_item_5) price_item_5  " _
            & "    , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 ,SUM(price_item_8) price_item_8  " _
            & "    , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 " _
            & "    , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 " _
            & "    , SUM(price_item_15) price_item_15 " _
            & GetTotalItem
            
        szSql = szSql _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number) - SUM(ticket_return_number) - sum(ticket_change_number)  ) AS passenger_number " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) ) AS ticket_price " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price  " _
            & "    , SUM(ticket_cancel_number) ticket_cancel_number , SUM(ticket_cancel_amount) ticket_cancel_amount " _
            & "    , SUM(ticket_return_number) ticket_return_number , SUM(ticket_return_amount) ticket_return_amount " _
            & "    , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_number) ticket_change_number " _
            & "    , SUM(ticket_change_amount) ticket_change_amount , SUM(ticket_change_charge) ticket_change_charge " _
            & " FROM  stat_bus_startoff_lst b , route_info r " _
            & " WHERE r.route_id = b.route_id AND  date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "'" & szWhere & ExceptZeroBus _
            & " GROUP BY bus_id ,sell_station_id ORDER BY bus_id ,sell_station_id "
    Set rsTemp = oDB.Execute(szSql)
    Set GetBusSellStationSellInfoByBusDate = rsTemp
    
End Function


'车次各上车站售票简报(按照车次组合查询,按售票时间汇总)
Public Function GetCombineBusSellStationBySaleTime(pdtFromDate As Date, pdtToDate As Date, pnSerial As Integer, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszLike As String = "", Optional pszSellStationID As String = "") As Recordset
'pnExtraStatus =1 表示售票，2表示补票，3表示全部



    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetCombineBusSimply
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
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
    If pszLike <> "" Then
        szWhere = szWhere & " AND b.bus_id like '" & pszLike & "%' "
    End If
    szSql = "   SELECT b.bus_id, b.sell_station_id, MAX(bus_start_time) bus_start_time  , MAX(transport_company_short_name) split_company_name  " _
            & "    , MAX(B.route_id) route_id , MAX(route_name) route_name " _
            & "    , SUM(base_price) base_price , SUM (price_item_1) price_item_1, SUM(price_item_2) price_item_2   " _
            & "    , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 , SUM(price_item_5) price_item_5  " _
            & "    , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 ,SUM(price_item_8) price_item_8  " _
            & "    , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 " _
            & "    , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 " _
            & "    , SUM(price_item_15) price_item_15 " _
            & GetTotalItem

        szSql = szSql _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number) - SUM(ticket_return_number) - sum(ticket_change_number) ) AS passenger_number " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) ) AS ticket_price " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price  " _
            & "    , SUM(ticket_cancel_number) ticket_cancel_number , SUM(ticket_cancel_amount) ticket_cancel_amount " _
            & "    , SUM(ticket_return_number) ticket_return_number , SUM(ticket_return_amount) ticket_return_amount " _
            & "    , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_number) ticket_change_number " _
            & "    , SUM(ticket_change_amount) ticket_change_amount , SUM(ticket_change_charge) ticket_change_charge " _
            & " FROM  Stat_bus_sell_lst b , company_info t , bus_assemble_info c , route_info r " _
            & " WHERE r.route_id = b.route_id and date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "'" _
            & " AND combine_serial =  " & pnSerial _
            & " AND c.transport_company_id *= t.transport_company_id " _
            & " AND b.bus_id >= c.start_bus_id  AND b.bus_id <=  c.end_bus_id  " & ExceptZeroBus _
            & szWhere _
            & " GROUP BY B.bus_id,b.sell_station_id order by B.BUS_ID , b.sell_station_id "

    
    Set rsTemp = oDB.Execute(szSql)
    Set GetCombineBusSellStationBySaleTime = rsTemp
    
End Function


'根据车次日期统计,指定车次组合内的车次的统计信息
Public Function GetCombineBusSellStationByBusDate(pdtFromDate As Date, pdtToDate As Date, pnSerial As Integer, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszLike As String = "", Optional pszSellStationID As String = "") As Recordset
'pnExtraStatus =1 表示售票，2表示补票，3表示全部
    
    
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetCombineBusSimplyByBusDate
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
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
    If pszLike <> "" Then
        szWhere = szWhere & " AND b.bus_id like '" & pszLike & "%' "
    End If

    szSql = "   SELECT b.bus_id,sell_station_id MAX(bus_start_time) bus_start_time  , MAX(transport_company_short_name) split_company_name  " _
            & "    , MAX(B.route_id) route_id , MAX(route_name) route_name " _
            & "    , SUM(base_price) base_price , SUM (price_item_1) price_item_1, SUM(price_item_2) price_item_2   " _
            & "    , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 , SUM(price_item_5) price_item_5  " _
            & "    , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 ,SUM(price_item_8) price_item_8  " _
            & "    , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 " _
            & "    , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 " _
            & "    , SUM(price_item_15) price_item_15 " _
            & GetTotalItem
            
        szSql = szSql _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number) - SUM(ticket_return_number) - sum(ticket_change_number)  ) AS passenger_number " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) ) AS ticket_price " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price  " _
            & "    , SUM(ticket_cancel_number) ticket_cancel_number , SUM(ticket_cancel_amount) ticket_cancel_amount " _
            & "    , SUM(ticket_return_number) ticket_return_number , SUM(ticket_return_amount) ticket_return_amount " _
            & "    , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_number) ticket_change_number " _
            & "    , SUM(ticket_change_amount) ticket_change_amount , SUM(ticket_change_charge) ticket_change_charge " _
            & " FROM  stat_bus_startoff_lst b , company_info t , bus_assemble_info c , route_info r " _
            & " WHERE r.route_id = b.route_id AND date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "'" _
            & " AND combine_serial =  " & pnSerial _
            & " AND c.transport_company_id *= t.transport_company_id " _
            & " AND b.bus_id >= c.start_bus_id  AND b.bus_id <=  c.end_bus_id  " & ExceptZeroBus _
            & szWhere _
            & " GROUP BY B.bus_id,sell_station_id order by B.BUS_ID,sell_station_id "
    
    
    Set rsTemp = oDB.Execute(szSql)
    Set GetCombineBusSellStationByBusDate = rsTemp
    
End Function



'车次途经站统计(按车次日期进行汇总)
Public Function GetBusStationStatByBusDate(pszCompanyID As String, pdtFromDate As Date, pdtToDate As Date, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszLike As String = "", Optional pszSellStationID As String = "") As Recordset
'pnExtraStatus =1 表示售票，2表示补票，3表示全部
'统计方式为按照车次日期,并按照车次的上车站进行统计


    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusStatByBusDate
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    If pszCompanyID <> "" Then
        szWhere = szWhere & " AND split_company_id IN ( " & pszCompanyID & ")"
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
    If pszLike <> "" Then
        szWhere = szWhere & " AND bus_id like '" & pszLike & "%' "
    End If

    szSql = "   SELECT bus_id, convert(char(5), MAX(bus_start_time),108) bus_start_time , MAX(B.route_id) route_id , MAX(route_name) route_name " _
            & "    , b.station_id , MAX(s.station_name) station_name , ticket_type , MAX(e.station_name) end_station_name,MAX(transport_company_short_name) transport_company_short_name  "
'
'            & "    , SUM(base_price) base_price , SUM (price_item_1) price_item_1, SUM(price_item_2) price_item_2   " _
'            & "    , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 , SUM(price_item_5) price_item_5  " _
'            & "    , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 ,SUM(price_item_8) price_item_8  " _
'            & "    , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 " _
'            & "    , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 " _
'            & "    , SUM(price_item_15) price_item_15 " _
'            & GetTotalItem
            
        szSql = szSql _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number) - SUM(ticket_return_number) - sum(ticket_change_number)  ) AS passenger_number " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) ) AS ticket_price " _
            & "    , SUM(passenger_number) passenger_number2 , SUM(ticket_price) ticket_price2 " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price  " _
            & "    , SUM(ticket_cancel_number) ticket_cancel_number , SUM(ticket_cancel_amount) ticket_cancel_amount " _
            & "    , SUM(ticket_return_number) ticket_return_number , SUM(ticket_return_amount) ticket_return_amount " _
            & "    , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_number) ticket_change_number " _
            & "    , SUM(ticket_change_amount) ticket_change_amount , SUM(ticket_change_charge) ticket_change_charge " _
            & " FROM  stat_bus_station_startoff_lst b , route_info r , station_info s , station_info e,company_info c" _
            & " WHERE r.route_id = b.route_id AND b.station_id = s.station_id AND r.end_station_id = e.station_id AND b.transport_company_id = c.transport_company_id " _
            & " AND  date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "'" & szWhere & ExceptZeroBus _
            & " GROUP BY bus_id , b.station_id , ticket_type ORDER BY bus_id , b.station_id , ticket_type "
    Set rsTemp = oDB.Execute(szSql)
    Set GetBusStationStatByBusDate = rsTemp
    
End Function





'车次途经站统计(按车次日期及售票员所属车站进行汇总)
Public Function GetBusStationStatByBusDateAndSeller(pszCompanyID As String, pdtFromDate As Date, pdtToDate As Date, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszLike As String = "", Optional pszSellStationID As String = "") As Recordset
'pnExtraStatus =1 表示售票，2表示补票，3表示全部
'统计方式为按照车次日期,并按照售票员的所属车站进行统计



    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusStatByBusDate
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    If pszCompanyID <> "" Then
        szWhere = szWhere & " AND split_company_id IN ( " & pszCompanyID & ")"
    End If
    '如果要查询的售票站不为空
    If Trim(pszSellStationID) <> "" Then
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
    If pszLike <> "" Then
        szWhere = szWhere & " AND bus_id like '" & pszLike & "%' "
    End If

    szSql = "   SELECT bus_id, convert(char(5), MAX(bus_start_time),108) bus_start_time , MAX(B.route_id) route_id , MAX(route_name) route_name " _
            & "    , b.station_id , MAX(s.station_name) station_name , ticket_type , MAX(e.station_name) end_station_name,MAX(transport_company_short_name) transport_company_short_name  "
'
'            & "    , SUM(base_price) base_price , SUM (price_item_1) price_item_1, SUM(price_item_2) price_item_2   " _
'            & "    , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 , SUM(price_item_5) price_item_5  " _
'            & "    , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 ,SUM(price_item_8) price_item_8  " _
'            & "    , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 " _
'            & "    , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 " _
'            & "    , SUM(price_item_15) price_item_15 " _
'            & GetTotalItem
            
        szSql = szSql _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number) - SUM(ticket_return_number) - sum(ticket_change_number)  ) AS passenger_number " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) ) AS ticket_price " _
            & "    , SUM(passenger_number) passenger_number2 , SUM(ticket_price) ticket_price2 " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price  " _
            & "    , SUM(ticket_cancel_number) ticket_cancel_number , SUM(ticket_cancel_amount) ticket_cancel_amount " _
            & "    , SUM(ticket_return_number) ticket_return_number , SUM(ticket_return_amount) ticket_return_amount " _
            & "    , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_number) ticket_change_number " _
            & "    , SUM(ticket_change_amount) ticket_change_amount , SUM(ticket_change_charge) ticket_change_charge " _
            & " FROM  stat_bus_station_sell_lst b , route_info r , station_info s , station_info e,company_info c" _
            & " WHERE r.route_id = b.route_id AND b.station_id = s.station_id AND r.end_station_id = e.station_id AND b.transport_company_id = c.transport_company_id " _
            & " AND  date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "'" & szWhere & ExceptZeroBus _
            & " GROUP BY bus_id , b.station_id , ticket_type ORDER BY bus_id , b.station_id , ticket_type "
    Set rsTemp = oDB.Execute(szSql)
    Set GetBusStationStatByBusDateAndSeller = rsTemp
    
End Function






'公司分类营收报表(按车次日期进行汇总)
Public Function GetRouteCompanyByBusDate(pdtFromDate As Date, pdtToDate As Date, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszSellStationID As String = "", Optional pszCompanyID As String = "", Optional pszRouteID As String = "") As Recordset
'pnExtraStatus =1 表示售票，2表示补票，3表示全部

    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusStatByBusDate
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    If pszCompanyID <> "" Then
        szWhere = " AND b.transport_company_id IN ( '" & pszCompanyID & "')"
    End If
    '如果要查询的售票站不为空
    If pszSellStationID <> "" Then
        szWhere = szWhere & " AND b.sell_station_id='" & pszSellStationID & "'"
    End If
    
    If pszRouteID <> "" Then
        szWhere = szWhere & " AND d.route_name= " & pszRouteID
    End If
    
'    If pnExtraStatus <> TP_AllSold Then
'    '如果不是所有的状态
'        szWhere = szWhere & " AND extra_status = " & pnExtraStatus
'
'    End If
'    If pszLike <> "" Then
'        szWhere = szWhere & " AND bus_id like '" & pszLike & "%' "
'    End If



    szSql = "   SELECT   Max(route_name) route_name ,MAX(transport_company_short_name) transport_company_short_name ,ticket_type  "
        szSql = szSql _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number) - SUM(ticket_return_number) - sum(ticket_change_number)  ) AS passenger_number " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) ) AS ticket_price " _
            & "    , SUM(passenger_number) passenger_number2 , SUM(ticket_price) ticket_price2 " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price  " _
            & "    , SUM(ticket_cancel_number) ticket_cancel_number , SUM(ticket_cancel_amount) ticket_cancel_amount " _
            & "    , SUM(ticket_return_number) ticket_return_number , SUM(ticket_return_amount) ticket_return_amount " _
            & "    , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_number) ticket_change_number " _
            & "    , SUM(ticket_change_amount) ticket_change_amount , SUM(ticket_change_charge) ticket_change_charge " _
            & " FROM  stat_bus_station_startoff_lst b ,company_info c,route_info d" _
            & " WHERE b.transport_company_id = c.transport_company_id  AND b.route_id=d.route_id " & szWhere _
            & " AND  date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "'" & szWhere & ExceptZeroBus _
            & " GROUP BY route_name,transport_company_short_name ,ticket_type ORDER BY route_name,transport_company_short_name,ticket_type  "
    Set rsTemp = oDB.Execute(szSql)
    Set GetRouteCompanyByBusDate = rsTemp
    
End Function




'车次各上车站售票简报(按售票时间汇总各类款额)//王候记在衢州加
Public Function GetBusSellStationSellInfoBySaleTimeSimply(pszCompanyID As String, pdtFromDate As Date, pdtToDate As Date, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszLike As String = "", Optional pszSellStationID As String = "") As Recordset
'pnExtraStatus =1 表示售票，2表示补票，3表示全部
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusStatBySaleTime
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    If pszCompanyID <> "" Then
        szWhere = " AND split_company_id IN ( " & pszCompanyID & ")"
    End If
    '如果要查询的售票站不为空
    If pszSellStationID <> "" Then
        szWhere = szWhere & " AND b.sell_station_id='" & pszSellStationID & "'"
    End If
    If pnExtraStatus <> TP_AllSold Then
    '如果不是所有的状态
        szWhere = szWhere & " AND extra_status = " & pnExtraStatus
    End If
    If pszLike <> "" Then
        szWhere = szWhere & " AND bus_id like '" & pszLike & "%' "
    End If
    szSql = "   SELECT  sell_station_id, Max(s.station_name) station_name  " _
            & "    , SUM(base_price) base_price , SUM (price_item_1) price_item_1, SUM(price_item_2) price_item_2   " _
            & "    , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 , SUM(price_item_5) price_item_5  " _
            & "    , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 ,SUM(price_item_8) price_item_8  " _
            & "    , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 " _
            & "    , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 " _
            & "    , SUM(price_item_15) price_item_15 " _
            & GetTotalItem
            
        szSql = szSql _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number)- SUM(ticket_return_number) - sum(ticket_change_number) ) AS passenger_number " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) ) AS ticket_price " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price  " _
            & "    , SUM(ticket_cancel_number) ticket_cancel_number , SUM(ticket_cancel_amount) ticket_cancel_amount " _
            & "    , SUM(ticket_return_number) ticket_return_number , SUM(ticket_return_amount) ticket_return_amount " _
            & "    , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_number) ticket_change_number " _
            & "    , SUM(ticket_change_amount) ticket_change_amount , SUM(ticket_change_charge) ticket_change_charge " _
            & " FROM  Stat_bus_sell_lst b,station_info s " _
            & " WHERE  b.sell_station_id=s.station_input_code AND  date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "'" & szWhere & ExceptZeroBus _
            & " GROUP BY sell_station_id  ORDER BY sell_station_id  "
            
    Set rsTemp = oDB.Execute(szSql)
    Set GetBusSellStationSellInfoBySaleTimeSimply = rsTemp
    
End Function


'车次各上车站售票简报(按发车时间汇总各类款额)
Public Function GetBusSellStationSellInfoByBusDatesimply(pszCompanyID As String, pdtFromDate As Date, pdtToDate As Date, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszLike As String = "", Optional pszSellStationID As String = "") As Recordset
'pnExtraStatus =1 表示售票，2表示补票，3表示全部
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusStatByBusDate
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    If pszCompanyID <> "" Then
        szWhere = " AND split_company_id IN ( " & pszCompanyID & ")"
    End If
    '如果要查询的售票站不为空
    If pszSellStationID <> "" Then
        szWhere = szWhere & " AND b.sell_station_id='" & pszSellStationID & "'"
    End If
    If pnExtraStatus <> TP_AllSold Then
    '如果不是所有的状态
        szWhere = szWhere & " AND extra_status = " & pnExtraStatus
        
    End If
    If pszLike <> "" Then
        szWhere = szWhere & " AND bus_id like '" & pszLike & "%' "
    End If

    szSql = "   SELECT sell_station_id, Max(s.station_name) station_name " _
            & "    , SUM(base_price) base_price , SUM (price_item_1) price_item_1, SUM(price_item_2) price_item_2   " _
            & "    , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 , SUM(price_item_5) price_item_5  " _
            & "    , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 ,SUM(price_item_8) price_item_8  " _
            & "    , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 " _
            & "    , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 " _
            & "    , SUM(price_item_15) price_item_15 " _
            & GetTotalItem
            
        szSql = szSql _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number) - SUM(ticket_return_number) - sum(ticket_change_number)  ) AS passenger_number " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) ) AS ticket_price " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price  " _
            & "    , SUM(ticket_cancel_number) ticket_cancel_number , SUM(ticket_cancel_amount) ticket_cancel_amount " _
            & "    , SUM(ticket_return_number) ticket_return_number , SUM(ticket_return_amount) ticket_return_amount " _
            & "    , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_number) ticket_change_number " _
            & "    , SUM(ticket_change_amount) ticket_change_amount , SUM(ticket_change_charge) ticket_change_charge " _
            & " FROM  stat_bus_startoff_lst b ,station_info s " _
            & " WHERE  b.sell_station_id=s.station_input_code AND  date >= " & TransFieldValueToString(ToDBDateTime(pdtFromDate)) _
            & " AND date <" & TransFieldValueToString(ToDBDateTime(DateAdd("d", 1, pdtToDate))) & szWhere _
            & " GROUP BY sell_station_id ORDER BY sell_station_id "
    Set rsTemp = oDB.Execute(szSql)
    Set GetBusSellStationSellInfoByBusDatesimply = rsTemp
    
End Function




'========================================
'此接口为查询车辆结算平衡表中的售票部分用
'========================================

'车辆统计(按车次日期进行汇总)
Public Function GetVehicleStatByBusDate(pszCompanyID As String, pdtFromDate As Date, pdtToDate As Date, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszVehicleID As String = "", Optional pszSellStationID As String = "") As Recordset
'pnExtraStatus =1 表示售票，2表示补票，3表示全部



    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusStatByBusDate
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    If pszCompanyID <> "" Then
        szWhere = " AND b.split_company_id IN ( " & pszCompanyID & ")"
    End If
    '如果要查询的售票站不为空
    If pszSellStationID <> "" Then
'        szWhere = szWhere & " AND b.sell_station_id='" & pszSellStationID & "'"
        szWhere = szWhere & " AND b.sell_station_id IN ( " & pszSellStationID & ")"
    End If
    If pnExtraStatus <> TP_AllSold Then
    '如果不是所有的状态
        szWhere = szWhere & " AND extra_status = " & pnExtraStatus
        
    End If
    If pszVehicleID <> "" Then
        szWhere = szWhere & " AND b.vehicle_id= '" & pszVehicleID & "'"
    End If

    szSql = "   SELECT b.vehicle_id, MAX(b.bus_start_time) bus_start_time , MAX(v.license_tag_no) license_tag_no " _
            & " , MAX(B.route_id) route_id , MAX(route_name) route_name " _
            & "    , SUM(base_price) base_price , SUM (price_item_1) price_item_1, SUM(price_item_2) price_item_2   " _
            & "    , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 , SUM(price_item_5) price_item_5  " _
            & "    , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 ,SUM(price_item_8) price_item_8  " _
            & "    , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 " _
            & "    , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 " _
            & "    , SUM(price_item_15) price_item_15 " _
            & GetTotalItem
            
        szSql = szSql _
            & "    , (SUM(passenger_number) -SUM(ticket_cancel_number) - SUM(ticket_return_number) - sum(ticket_change_number)  ) AS passenger_number " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) ) AS ticket_price " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price  " _
            & "    , SUM(ticket_cancel_number) ticket_cancel_number , SUM(ticket_cancel_amount) ticket_cancel_amount " _
            & "    , SUM(ticket_return_number) ticket_return_number , SUM(ticket_return_amount) ticket_return_amount " _
            & "    , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_number) ticket_change_number " _
            & "    , SUM(ticket_change_amount) ticket_change_amount , SUM(ticket_change_charge) ticket_change_charge " _
            & " FROM  stat_bus_startoff_lst b , route_info r , vehicle_info v  " _
            & " WHERE b.vehicle_id = v.vehicle_id AND r.route_id = b.route_id AND  date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "'" & szWhere & ExceptZeroBus _
            & " GROUP BY b.vehicle_id ORDER BY b.vehicle_id "
    Set rsTemp = oDB.Execute(szSql)
    Set GetVehicleStatByBusDate = rsTemp
    
End Function



