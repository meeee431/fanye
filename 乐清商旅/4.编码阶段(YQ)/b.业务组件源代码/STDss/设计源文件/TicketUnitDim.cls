VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TicketUnitDim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrTicketStationDim
    ERR_TicketStationDimNoAcitveUser = ERR_TicketStationDim + ERR_NoActiveUser
End Enum

'此处定义权限号：
Public Enum ERightTicketStationDim
    RIGHT_QueryStationDateStat = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 1  '车站售票营收日报查询
    RIGHT_QueryStationMonthStat = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 2  '车站售票营收月报查询
    RIGHT_QueryDateUnitStat = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 3  '车站互联售票统计报表查询
    RIGHT_QueryUnitDateStat = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 4  '车站互联单位售票简报查询
    RIGHT_QueryExUnitDateStat = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 5  '互联单位互售票统计报表查询
    RIGHT_AddBusCompanyCombine = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 6 '添加车次公司组合
    RIGHT_BusSellScheme = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 7
    RIGHT_BusTicketSellScheme = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 8
    RIGHT_DateExUnitStat = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 9
    RIGHT_DateUnitStat = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 10
    RIGHT_ExUnitDateStat = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 11
    RIGHT_GetBusDateBusIDQuery = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 12
    RIGHT_GetBusDateCheckGateQuery = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 13
    RIGHT_GetBusDateFromBusIDSeatSellInfo = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 14
    RIGHT_GetBusDateFromRoutIdSeatSellInfo = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 15
    RIGHT_GetBusDateRouteQuery = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 16
    RIGHT_GetBusDateSeatSellInfo = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 17
    RIGHT_GetBusDateStation = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 18
    RIGHT_GetCheckGateStatDetail = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 19
    RIGHT_GetCheckGateStatDetailSumByDate = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 20
    RIGHT_GetCheckGateStatSum = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 21
    RIGHT_GetSellBusID = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 22
    RIGHT_GetSellCheckGateID = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 23
    RIGHT_GetSellRouteID = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 24
    RIGHT_GetSellStationID = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 25
    RIGHT_GetStationDaily = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 26
    RIGHT_GetStopBus = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 27
    RIGHT_ProvinceSum = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 28
    RIGHT_StationDateStat = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 29
    RIGHT_StationMonthStat = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 30
    RIGHT_UnitDateStat = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 31
    RIGHT_UnitInterSettle = ERR_TicketStationDim + cnMidRightBegin + cnMidRightStep * 32

End Enum

Private m_oActiveUser As ActiveUser

'车站售票营收日报
Public Function StationDateStat(FromDate As Date, ToDate As Date, Optional pszSellStationID As String = "") As Recordset
'    AssertActiveUserValid m_oActiveUser, ERR_TicketStationDim
'    AssertHaveRight m_oActiveUser, RIGHT_StationDateStat
'
'    Dim rsTemp As Recordset
'    Dim szSql As String
'    Dim oDB As New RTConnection
'    Dim nTemp As Integer
'    Dim i As Integer
'
'    oDB.ConnectionString = GetConnectionStr(cszDss)
''
''    szSql = "SELECT * FROM price_item_info WHERE use_mark=1 AND price_item<>'0000'"
''    Set rsTemp = oDb.Execute(szSql)
''
''
'    szSql = "SELECT bus_date,u.ticket_type,MAX(t.ticket_type_name) ticket_type_name ," _
'    & " SUM(ticket_price) AS ticket_price ," _
'    & " SUM(passenger_number) AS passenger_number," _
'    & " SUM(base_price) AS base_price "
'
''    Do While Not rsTemp.EOF
''        nTemp = CInt(rsTemp!price_item)
'    For i = 1 To 15
'        szSql = szSql & ",SUM(price_item_" & i & ") AS price_item_" & i
'    Next i
''        rsTemp.MoveNext
''    Loop
'
'    szSql = szSql _
'    & " FROM unit_day_ticket_item_stat u , Ticket_type_code t " _
'    & " WHERE u.ticket_type = t.ticket_type_id " _
'    & " AND u.bus_date>='" & ToDBDate(FromDate) & "' " _
'    & " AND u.bus_date<='" & ToDBDate(ToDate) & "' "
'    If pszSellStationID <> "" Then
'       szSql = szSql & " AND u.sell_station_id='" & pszSellStationID & "'"
'    End If
'    szSql = szSql & " GROUP BY u.bus_date,u.ticket_type ORDER BY u.bus_date,u.ticket_type"
'    Set rsTemp = oDB.Execute(szSql)
'    Set StationDateStat = rsTemp
End Function


'车站售票营收月报
Public Function StationMonthStat(FromMonth As Date, ToMonth As Date, Optional pszSellStationID As String = "") As Recordset
'    AssertActiveUserValid m_oActiveUser, ERR_TicketStationDim
'    AssertHaveRight m_oActiveUser, RIGHT_StationMonthStat
'
'    Dim rsTemp As Recordset
'    Dim szSql As String
'    Dim oDB As New RTConnection
''    Dim dtFromMonth As Date, dtToMonth As Date
'    Dim nTemp As Integer
'    Dim i As Integer
'
'    oDB.ConnectionString = GetConnectionStr(cszDss)
''
''    szSql = "SELECT * FROM price_item_info WHERE use_mark=1 AND price_item<>'0000'"
''    Set rsTemp = oDb.Execute(szSql)
''
'
'
'
'    szSql = "SELECT CONVERT(CHAR(7),bus_date,102) AS bus_date,u.ticket_type, MAX(t.ticket_type_name) ticket_type_name ," _
'    & " SUM(ticket_price) AS ticket_price ," _
'    & " SUM(passenger_number) AS passenger_number," _
'    & " SUM(base_price) AS base_price "
'
''    Do While Not rsTemp.EOF
''        nTemp = CInt(rsTemp!price_item)
'    For i = 1 To 15
'        szSql = szSql & ",SUM(price_item_" & i & ") AS price_item_" & i
'
'    Next i
''        rsTemp.MoveNext
''    Loop
''
'    szSql = szSql _
'    & " FROM unit_day_ticket_item_stat  u , Ticket_type_code t " _
'    & " WHERE u.ticket_type = t.ticket_type_id AND u.bus_date>='" & ToDBDate(FromMonth) & "' AND " _
'    & " u.bus_date<'" & ToDBDate(ToMonth) & "' "
'    If pszSellStationID <> "" Then
'       szSql = szSql & " AND u.sell_station_id='" & pszSellStationID & "'"
'    End If
'
'    szSql = szSql & " GROUP BY CONVERT(CHAR(7),u.bus_date,102),u.ticket_type ORDER BY CONVERT(CHAR(7),u.bus_date,102),u.ticket_type"
'
'    Set rsTemp = oDB.Execute(szSql)
'
'    Set StationMonthStat = rsTemp
End Function

'车站互连售票统计报表
Public Function DateUnitStat(UnitID() As String, FromDate As Date, ToDate As Date) As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_TicketStationDim
    AssertHaveRight m_oActiveUser, RIGHT_DateUnitStat
    
    Dim nUnitCount As Integer, i As Integer
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim rsData As New Recordset
    Dim j As Integer
    Dim dbAmount As Double
    Dim lNumber As Long
    
    szSql = "SELECT RTRIM(connect_unit_info.unit_id) +'['+RTRIM(unit_short_name)+']' AS unit_id1 ,ticket_type," _
    & " SUM(ticket_sale_number) AS ticket_sale_number1, " _
    & " SUM(ticket_sale_amount) AS ticket_sale_amount1," _
    & " SUM(ticket_cancel_number) AS ticket_cancel_number1," _
    & " SUM(ticket_cancel_amount) AS ticket_cancel_amount1," _
    & " SUM(ticket_return_number) AS ticket_return_number1," _
    & " SUM(ticket_return_amount) AS ticket_return_amount1," _
    & " SUM(ticket_return_charge) AS ticket_return_charge1," _
    & " SUM(ticket_change_number) AS ticket_change_number1," _
    & " SUM(ticket_change_amount) AS ticket_change_amount1," _
    & " SUM(ticket_change_charge) AS ticket_change_charge1 " _
    & " FROM Stat_unit_sell_lst,connect_unit_info WHERE " _
    & " Stat_unit_sell_lst.unit_id=connect_unit_info.unit_id AND " _
    & " bus_date>='" & ToDBDate(FromDate) & "' AND " _
    & " bus_date<='" & ToDBDate(ToDate) & "' "
    
    nUnitCount = ArrayLength(UnitID)
    If nUnitCount > 0 Then
        szSql = szSql & " AND (Stat_unit_sell_lst.unit_id='" & UnitID(1) & "' "
        For i = 2 To nUnitCount
            szSql = szSql & " OR Stat_unit_sell_lst.unit_id='" & UnitID(i) & "' "
        Next
        szSql = szSql & ")"
    End If
    szSql = szSql & " GROUP BY connect_unit_info.unit_id,unit_short_name,ticket_type ORDER BY unit_id1,ticket_type "
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    Set rsTemp = oDB.Execute(szSql)
    
    
'做出Recordset
    '添加字段
    rsData.Fields.Append "unit_id", rsTemp("unit_id1").Type, rsTemp("unit_id1").DefinedSize
    
    '添加各种票种的数量与金额
    For j = 1 To TP_TicketTypeCount
        rsData.Fields.Append "number_ticket_type" & j, adInteger
        rsData.Fields.Append "amount_ticket_type" & j, adCurrency
    Next j

    rsData.Fields.Append "cancel_number", adInteger
    rsData.Fields.Append "cancel_amount", adCurrency
    
    rsData.Fields.Append "return_number", adInteger
    rsData.Fields.Append "return_amount", adCurrency
    rsData.Fields.Append "return_charge", adCurrency
    
    rsData.Fields.Append "change_number", adInteger
    rsData.Fields.Append "change_amount", adCurrency
    rsData.Fields.Append "change_charge", adCurrency
    
    rsData.Fields.Append "total_number", adInteger
    rsData.Fields.Append "total_amount", adCurrency
    
    rsData.Open
    Dim szLastUnitID As String
    
    Do While Not rsTemp.EOF
        If szLastUnitID <> RTrim(rsTemp!unit_id1) Or rsTemp!ticket_type = TP_FullPrice Then
            If rsData.RecordCount > 0 Then
                '求总张数及总票款
                lNumber = 0
                dbAmount = 0
                For j = 1 To TP_TicketTypeCount
                    If j <> TP_FreeTicket Then
                        dbAmount = dbAmount + rsData("amount_ticket_type" & j)
                    End If
                
                    lNumber = lNumber + rsData("number_ticket_type" & j)
                Next j
                rsData!total_number = lNumber
                rsData!total_amount = dbAmount - rsData!return_amount + rsData!return_charge - rsData!change_amount + rsData!change_charge
                rsData.Update

            End If
            rsData.AddNew
            
            szLastUnitID = RTrim(rsTemp!unit_id1)
            rsData!unit_id = szLastUnitID
        End If
        
        rsData("number_ticket_type" & rsTemp!ticket_type) = rsTemp!ticket_sale_number1
        rsData("amount_ticket_type" & rsTemp!ticket_type) = rsTemp!ticket_sale_amount1

        rsData!cancel_number = rsData!cancel_number + rsTemp!ticket_cancel_number1
        rsData!cancel_amount = rsData!cancel_amount + rsTemp!ticket_cancel_amount1
        
        rsData!return_number = rsData!return_number + rsTemp!ticket_return_number1
        rsData!return_amount = rsData!return_amount + rsTemp!ticket_return_amount1
        rsData!return_charge = rsData!return_charge + rsTemp!ticket_return_charge1
        
        rsData!change_number = rsData!change_number + rsTemp!ticket_change_number1
        rsData!change_amount = rsData!change_amount + rsTemp!ticket_change_amount1
        rsData!change_charge = rsData!change_charge + rsTemp!ticket_change_charge1
        
        rsTemp.MoveNext
    Loop
    If rsData.RecordCount > 0 Then
        '求总张数及总票款
        lNumber = 0
        dbAmount = 0
        For j = 1 To TP_TicketTypeCount
            If j <> TP_FreeTicket Then
                dbAmount = dbAmount + rsData("amount_ticket_type" & j)
            End If
            lNumber = lNumber + rsData("number_ticket_type" & j)
        Next j
        rsData!total_number = lNumber
        rsData!total_amount = dbAmount - rsData!return_amount + rsData!return_charge - rsData!change_amount + rsData!change_charge
        
        rsData.Update
    End If
    
    Set DateUnitStat = rsData
End Function

'车站互连单位售票简报
Public Function UnitDateStat(UnitID() As String, FromDate As Date, ToDate As Date, Optional pszSellStationID As String = "") As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_TicketStationDim
    AssertHaveRight m_oActiveUser, RIGHT_UnitDateStat
    Dim nUnitCount As Integer, i As Integer
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim rsData As New Recordset
    Dim j As Integer
    Dim dbAmount As Double
    Dim lNumber As Long
    
    
    szSql = "SELECT bus_date ,ticket_type," _
    & " SUM(ticket_sale_number) AS ticket_sale_number1, " _
    & " SUM(ticket_sale_amount) AS ticket_sale_amount1," _
    & " SUM(ticket_cancel_number) AS ticket_cancel_number1," _
    & " SUM(ticket_cancel_amount) AS ticket_cancel_amount1," _
    & " SUM(ticket_return_number) AS ticket_return_number1," _
    & " SUM(ticket_return_amount) AS ticket_return_amount1," _
    & " SUM(ticket_return_charge) AS ticket_return_charge1," _
    & " SUM(ticket_change_number) AS ticket_change_number1," _
    & " SUM(ticket_change_amount) AS ticket_change_amount1," _
    & " SUM(ticket_change_charge) AS ticket_change_charge1 " _
    & " FROM Stat_unit_sell_lst WHERE " _
    & " bus_date>='" & ToDBDate(FromDate) & "' AND " _
    & " bus_date<='" & ToDBDate(ToDate) & "' "
    
    nUnitCount = ArrayLength(UnitID)
    If nUnitCount > 0 Then
        szSql = szSql & " AND (unit_id='" & UnitID(1) & "' "
        For i = 2 To nUnitCount
            szSql = szSql & " OR unit_id='" & UnitID(i) & "' "
        Next
        szSql = szSql & ")"
    End If
    If pszSellStationID <> "" Then
        szSql = szSql & " AND sell_station_id='" & pszSellStationID & "'"
    End If
    szSql = szSql & " GROUP BY bus_date,ticket_type ORDER BY bus_date,ticket_type "
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    Set rsTemp = oDB.Execute(szSql)
    
    
'做出Recordset

    '添加字段
    rsData.Fields.Append "bus_date", rsTemp("bus_date").Type, rsTemp("bus_date").DefinedSize
    
    '添加各种票种的数量与金额
    For j = 1 To TP_TicketTypeCount
        rsData.Fields.Append "number_ticket_type" & j, adInteger
        rsData.Fields.Append "amount_ticket_type" & j, adCurrency
    Next j
    
    rsData.Fields.Append "cancel_number", adInteger
    rsData.Fields.Append "cancel_amount", adCurrency
    
    rsData.Fields.Append "return_number", adInteger
    rsData.Fields.Append "return_amount", adCurrency
    rsData.Fields.Append "return_charge", adCurrency
    
    rsData.Fields.Append "change_number", adInteger
    rsData.Fields.Append "change_amount", adCurrency
    rsData.Fields.Append "change_charge", adCurrency
    
    rsData.Fields.Append "total_number", adInteger
    rsData.Fields.Append "total_amount", adCurrency
    
    rsData.Open
    Dim dtLastDate As Date
    
    Do While Not rsTemp.EOF
        If dtLastDate <> rsTemp!bus_date Or rsTemp!ticket_type = TP_FullPrice Then
            If rsData.RecordCount > 0 Then
                '求总张数及总票款
                lNumber = 0
                dbAmount = 0
                For j = 1 To TP_TicketTypeCount
                    If j <> TP_FreeTicket Then
                        dbAmount = dbAmount + rsData("amount_ticket_type" & j)
                    End If
                
                    lNumber = lNumber + rsData("number_ticket_type" & j)
                Next j
                rsData!total_number = lNumber - rsData!return_number - rsData!change_number
                rsData!total_amount = dbAmount - rsData!return_amount + rsData!return_charge - rsData!change_amount + rsData!change_charge
                rsData.Update
            End If
            rsData.AddNew
            
            dtLastDate = rsTemp!bus_date
            rsData!bus_date = dtLastDate
        End If
        
        rsData("number_ticket_type" & rsTemp!ticket_type) = rsTemp!ticket_sale_number1
        rsData("amount_ticket_type" & rsTemp!ticket_type) = rsTemp!ticket_sale_amount1
        
        rsData!cancel_number = rsData!cancel_number + rsTemp!ticket_cancel_number1
        rsData!cancel_amount = rsData!cancel_amount + rsTemp!ticket_cancel_amount1
        
        rsData!return_number = rsData!return_number + rsTemp!ticket_return_number1
        rsData!return_amount = rsData!return_amount + rsTemp!ticket_return_amount1
        rsData!return_charge = rsData!return_charge + rsTemp!ticket_return_charge1
        
        rsData!change_number = rsData!change_number + rsTemp!ticket_change_number1
        rsData!change_amount = rsData!change_amount + rsTemp!ticket_change_amount1
        rsData!change_charge = rsData!change_charge + rsTemp!ticket_change_charge1
        
        rsTemp.MoveNext
    Loop
    If rsData.RecordCount > 0 Then
        '求总张数及总票款
        lNumber = 0
        dbAmount = 0
        For j = 1 To TP_TicketTypeCount
            If j <> TP_FreeTicket Then
                dbAmount = dbAmount + rsData("amount_ticket_type" & j)
            End If
        
            lNumber = lNumber + rsData("number_ticket_type" & j)
        Next j
        rsData!total_number = lNumber - rsData!return_number - rsData!change_number
        rsData!total_amount = dbAmount - rsData!return_amount + rsData!return_charge - rsData!change_amount + rsData!change_charge
        rsData.Update
    End If
    
    Set UnitDateStat = rsData
End Function

'车站售票营收简报
Public Function UnitSellSimple(SellStationID() As String, FromDate As Date, ToDate As Date) As Recordset
'    AssertActiveUserValid m_oActiveUser, ERR_TicketStationDim
'    AssertHaveRight m_oActiveUser, RIGHT_DateUnitStat
    
    Dim nUnitCount As Integer, i As Integer
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp1 As Recordset
    Dim rsTemp2 As Recordset
    Dim rsTemp3 As Recordset
    Dim rsData As New Recordset
    Dim j As Integer
    Dim dbReturnChargeAmount, dbSellNumber, dbSellAmount, dbReturnNumber, dbReturnAmount, dbCancelNumber, dbCancelAmount As Double
    Dim lNumber As Long
    
    szSql = "select ticket_type,count(*) as renshu,sum(ticket_price) as price from Agent_ticket_sale_lst WHERE " _
    & " operation_time>='" & ToDBDateTime(FromDate) & "' AND " _
    & " operation_time<='" & ToDBDateTime(ToDate) & "' "
    If SellStationID(1) <> "" Then
        Dim szSellStationID() As String
        szSellStationID = Split(SellStationID(1), ",")
        szSql = szSql & " AND (Agent_ticket_sale_lst.sell_station_id='" & szSellStationID(0) & "' "
        For i = 1 To ArrayLength(szSellStationID) - 1
            szSql = szSql & " OR Agent_ticket_sale_lst.sell_station_id='" & szSellStationID(i) & "' "
        Next
        szSql = szSql & ")"
    End If
    szSql = szSql & " GROUP BY ticket_type "
    oDB.ConnectionString = GetConnectionStr(cszDss)
    Set rsTemp1 = oDB.Execute(szSql)
    
    szSql = "select ticket_type,count(*) as renshu,sum(ticket_price) as price,sum(return_charge) as return_charge from Agent_ticket_return_lst WHERE " _
    & " operation_time>='" & ToDBDateTime(FromDate) & "' AND " _
    & " operation_time<='" & ToDBDateTime(ToDate) & "' "
    If SellStationID(1) <> "" Then
        szSql = szSql & " AND (Agent_ticket_return_lst.sell_station_id='" & szSellStationID(0) & "' "
        For i = 1 To ArrayLength(szSellStationID) - 1
            szSql = szSql & " OR Agent_ticket_return_lst.sell_station_id='" & szSellStationID(i) & "' "
        Next
        szSql = szSql & ")"
    End If
    szSql = szSql & " GROUP BY ticket_type "
    oDB.ConnectionString = GetConnectionStr(cszDss)
    Set rsTemp2 = oDB.Execute(szSql)
    
    szSql = "select ticket_type,count(*) as renshu,sum(ticket_price) as price from Agent_ticket_cancel_lst c,Agent_ticket_sale_lst s WHERE " _
    & " c.ticket_id = s.ticket_id  AND " _
    & " cancel_time>='" & ToDBDateTime(FromDate) & "' AND " _
    & " cancel_time<='" & ToDBDateTime(ToDate) & "' "
    If SellStationID(1) <> "" Then
        szSql = szSql & " AND (s.sell_station_id='" & szSellStationID(0) & "' "
        For i = 1 To ArrayLength(szSellStationID) - 1
            szSql = szSql & " OR s.sell_station_id='" & szSellStationID(i) & "' "
        Next
        szSql = szSql & ")"
    End If
    nUnitCount = ArrayLength(SellStationID)
    szSql = szSql & " GROUP BY ticket_type "
    oDB.ConnectionString = GetConnectionStr(cszDss)
    Set rsTemp3 = oDB.Execute(szSql)
    
    
'做出Recordset
    '添加字段
    
    '添加各种票种的数量与金额
    For j = 1 To TP_TicketTypeCount
        rsData.Fields.Append "sell_number_ticket_type" & j, adInteger
        rsData.Fields.Append "sell_amount_ticket_type" & j, adCurrency
    Next j

    For j = 1 To TP_TicketTypeCount
        rsData.Fields.Append "return_number_ticket_type" & j, adInteger
        rsData.Fields.Append "return_amount_ticket_type" & j, adCurrency
    Next j
    
    rsData.Fields.Append "return_charge", adCurrency
    
    For j = 1 To TP_TicketTypeCount
        rsData.Fields.Append "cancel_number_ticket_type" & j, adInteger
        rsData.Fields.Append "cancel_amount_ticket_type" & j, adCurrency
    Next j
    
    rsData.Fields.Append "total_number", adInteger
    rsData.Fields.Append "total_amount", adCurrency
    
    rsData.Open

     rsData.AddNew
    Do While Not rsTemp1.EOF
        rsData("sell_number_ticket_type" & rsTemp1!ticket_type) = rsTemp1!renshu
        rsData("sell_amount_ticket_type" & rsTemp1!ticket_type) = rsTemp1!price
        rsTemp1.MoveNext
    Loop
    
    Do While Not rsTemp2.EOF
        rsData("return_number_ticket_type" & rsTemp2!ticket_type) = rsTemp2!renshu
        rsData("return_amount_ticket_type" & rsTemp2!ticket_type) = rsTemp2!price
        rsTemp2.MoveNext
    Loop
    
    If rsTemp2.RecordCount <> 0 Then
'        dbReturnChargeAmount = 0
        rsTemp2.MoveFirst
            Do While Not rsTemp2.EOF
            dbReturnChargeAmount = dbReturnChargeAmount + rsTemp2!return_charge
            dbReturnNumber = dbReturnNumber + rsTemp2!renshu
            dbReturnAmount = dbReturnAmount + rsTemp2!price
            rsTemp2.MoveNext
        Loop
    End If
    rsData("return_charge") = dbReturnChargeAmount
    
    Do While Not rsTemp3.EOF
        rsData("cancel_number_ticket_type" & rsTemp3!ticket_type) = rsTemp3!renshu
        rsData("cancel_amount_ticket_type" & rsTemp3!ticket_type) = rsTemp3!price
        rsTemp3.MoveNext
    Loop
    
    If rsTemp1.RecordCount <> 0 Then
'        dbReturnChargeAmount = 0
        rsTemp1.MoveFirst
            Do While Not rsTemp1.EOF
            dbSellNumber = dbSellNumber + rsTemp1!renshu
            dbSellAmount = dbSellAmount + rsTemp1!price
            rsTemp1.MoveNext
        Loop
    End If
    
    If rsTemp3.RecordCount <> 0 Then
'        dbReturnChargeAmount = 0
        rsTemp3.MoveFirst
            Do While Not rsTemp3.EOF
            dbCancelNumber = dbCancelNumber + rsTemp3!renshu
            dbCancelAmount = dbCancelAmount + rsTemp3!price
            rsTemp3.MoveNext
        Loop
    End If
    
    rsData("total_number") = dbSellNumber - dbCancelNumber - dbReturnNumber
    rsData("total_amount") = dbSellAmount - dbCancelAmount - dbReturnAmount + dbReturnChargeAmount

        rsData.Update
    
    Set UnitSellSimple = rsData
End Function

'互连互售票统计报表
Public Function ExUnitDateStat(UnitID() As String, FromDate As Date, ToDate As Date) As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_TicketStationDim
    AssertHaveRight m_oActiveUser, RIGHT_ExUnitDateStat
    
    Dim oDB As New RTConnection
    Dim rsSeller As Recordset, rsUnit As Recordset, rsData As New Recordset
    Dim szSql As String
    Dim nUnitCount As Integer, i As Integer
    Dim rsUnit2 As Recordset
    Dim j As Integer
    Dim lNumber As Long
    Dim dbAmount As Double
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    nUnitCount = ArrayLength(UnitID)
    
    szSql = "SELECT * FROM connect_unit_info "
    If nUnitCount > 0 Then
        szSql = szSql & " WHERE unit_id='" & UnitID(1) & "' "
        For i = 2 To nUnitCount
            szSql = szSql & " OR unit_id='" & UnitID(i) & "' "
        Next
        szSql = szSql & " ORDER BY unit_id "
    End If
    
    Set rsUnit2 = oDB.Execute(szSql)
    '取:售票员每日票务统计表
    
    szSql = "SELECT connect_unit_info.unit_id,unit_short_name ,ticket_type," _
    & " SUM(ticket_sale_number) AS ticket_sale_number1, " _
    & " SUM(ticket_sale_amount) AS ticket_sale_amount1," _
    & " SUM(ticket_cancel_number) AS ticket_cancel_number1," _
    & " SUM(ticket_cancel_amount) AS ticket_cancel_amount1," _
    & " SUM(ticket_return_number) AS ticket_return_number1," _
    & " SUM(ticket_return_amount) AS ticket_return_amount1," _
    & " SUM(ticket_return_charge) AS ticket_return_charge1," _
    & " SUM(ticket_change_number) AS ticket_change_number1," _
    & " SUM(ticket_change_amount) AS ticket_change_amount1," _
    & " SUM(ticket_change_charge) AS ticket_change_charge1 " _
    & " FROM Stat_saler_ticket_lst,connect_unit_info WHERE " _
    & " Stat_saler_ticket_lst.unit_id=connect_unit_info.unit_id AND " _
    & " bus_date>='" & ToDBDate(FromDate) & "' AND " _
    & " bus_date<='" & ToDBDate(ToDate) & "' "
    
    
    If nUnitCount > 0 Then
        szSql = szSql & " AND (Stat_saler_ticket_lst.unit_id='" & UnitID(1) & "' "
        For i = 2 To nUnitCount
            szSql = szSql & " OR Stat_saler_ticket_lst.unit_id='" & UnitID(i) & "' "
        Next
        szSql = szSql & ")"
    End If
    szSql = szSql & " GROUP BY connect_unit_info.unit_id,unit_short_name,ticket_type ORDER BY connect_unit_info.unit_id,ticket_type "
    
    Set rsSeller = oDB.Execute(szSql)
    
    '取:车站每日票务统计表
    szSql = "SELECT unit_id ,ticket_type," _
    & " SUM(ticket_sale_number) AS ticket_sale_number1, " _
    & " SUM(ticket_sale_amount) AS ticket_sale_amount1," _
    & " SUM(ticket_cancel_number) AS ticket_cancel_number1," _
    & " SUM(ticket_cancel_amount) AS ticket_cancel_amount1," _
    & " SUM(ticket_return_number) AS ticket_return_number1," _
    & " SUM(ticket_return_amount) AS ticket_return_amount1," _
    & " SUM(ticket_return_charge) AS ticket_return_charge1," _
    & " SUM(ticket_change_number) AS ticket_change_number1," _
    & " SUM(ticket_change_amount) AS ticket_change_amount1," _
    & " SUM(ticket_change_charge) AS ticket_change_charge1 " _
    & " FROM Stat_unit_sell_lst WHERE " _
    & " bus_date>='" & ToDBDate(FromDate) & "' AND " _
    & " bus_date<='" & ToDBDate(ToDate) & "' "
    
    If nUnitCount > 0 Then
        szSql = szSql & " AND (Stat_unit_sell_lst.unit_id='" & UnitID(1) & "' "
        For i = 2 To nUnitCount
            szSql = szSql & " OR Stat_unit_sell_lst.unit_id='" & UnitID(i) & "' "
        Next
        szSql = szSql & ")"
    End If
    szSql = szSql & " GROUP BY unit_id,ticket_type ORDER BY unit_id,ticket_type "
    
    Set rsUnit = oDB.Execute(szSql)
    
    '生成记录集
    With rsData.Fields
        .Append "unit_id", rsSeller("unit_id").Type, 20
        For i = 1 To 2
            '添加各种票种的数量与金额
            For j = 1 To TP_TicketTypeCount
                rsData.Fields.Append "number_ticket_type" & i & j, adInteger
                rsData.Fields.Append "amount_ticket_type" & i & j, adCurrency
            Next j
            
            .Append "cancel_number" & i, adInteger
            .Append "cancel_amount" & i, adCurrency
            
            .Append "return_number" & i, adInteger
            .Append "return_amount" & i, adCurrency
            .Append "return_charge" & i, adCurrency
            
            .Append "change_number" & i, adInteger
            .Append "change_amount" & i, adCurrency
            .Append "change_charge" & i, adCurrency
            
            .Append "total_number" & i, adInteger
            .Append "total_amount" & i, adCurrency
            
        Next i
    End With
    rsData.Open

    Dim szUnitTemp As String
    Dim szFieldPrefix As String
    
    For i = 1 To nUnitCount
        rsData.AddNew
        rsData!unit_id = RTrim(UnitID(i)) & "[" & GetUnitNameByID(rsUnit2, UnitID(i)) & "]"
        
        If FindUnit(rsSeller, UnitID(i)) Then
            Do While Not rsSeller.EOF
                If RTrim(rsSeller!unit_id) <> RTrim(UnitID(i)) Then Exit Do
                
'                Select Case rsSeller!ticket_type
'                    Case TP_FullPrice
'                    szFieldPrefix = "full_"
'                    Case TP_HalfPrice
'                    szFieldPrefix = "half_"
'                    Case TP_FreeTicket
'                    szFieldPrefix = "free_"
'                End Select
'                rsData(szFieldPrefix & "number2") = rsSeller!ticket_sale_number1
'                rsData(szFieldPrefix & "amount2") = rsSeller!ticket_sale_amount1
                
                rsData("number_ticket_type2" & rsSeller!ticket_type) = rsSeller!ticket_sale_number1
                rsData("amount_ticket_type2" & rsSeller!ticket_type) = rsSeller!ticket_sale_amount1

                
                rsData!cancel_number2 = rsData!cancel_number2 + rsSeller!ticket_cancel_number1
                rsData!cancel_amount2 = rsData!cancel_amount2 + rsSeller!ticket_cancel_amount1
                
                rsData!return_number2 = rsData!return_number2 + rsSeller!ticket_return_number1
                rsData!return_amount2 = rsData!return_amount2 + rsSeller!ticket_return_amount1
                rsData!return_charge2 = rsData!return_charge2 + rsSeller!ticket_return_charge1
                
                rsData!change_number2 = rsData!change_number2 + rsSeller!ticket_change_number1
                rsData!change_amount2 = rsData!change_amount2 + rsSeller!ticket_change_amount1
                rsData!change_charge2 = rsData!change_charge2 + rsSeller!ticket_change_charge1
                
                rsSeller.MoveNext
            Loop
        End If
    
        If FindUnit(rsUnit, UnitID(i)) Then
            Do While Not rsUnit.EOF
                If RTrim(rsUnit!unit_id) <> RTrim(UnitID(i)) Then Exit Do
'                Select Case rsUnit!ticket_type
'                    Case TP_FullPrice
'                    szFieldPrefix = "full_"
'                    Case TP_HalfPrice
'                    szFieldPrefix = "half_"
'                    Case TP_FreeTicket
'                    szFieldPrefix = "free_"
'                End Select
'                rsData(szFieldPrefix & "number1") = rsUnit!ticket_sale_number1
'                rsData(szFieldPrefix & "amount1") = rsUnit!ticket_sale_amount1
'
                
                rsData("number_ticket_type1" & rsUnit!ticket_type) = rsUnit!ticket_sale_number1
                rsData("amount_ticket_type1" & rsUnit!ticket_type) = rsUnit!ticket_sale_amount1

                rsData!cancel_number1 = rsData!cancel_number1 + rsUnit!ticket_cancel_number1
                rsData!cancel_amount1 = rsData!cancel_amount1 + rsUnit!ticket_cancel_amount1
                
                rsData!return_number1 = rsData!return_number1 + rsUnit!ticket_return_number1
                rsData!return_amount1 = rsData!return_amount1 + rsUnit!ticket_return_amount1
                rsData!return_charge1 = rsData!return_charge1 + rsUnit!ticket_return_charge1
                
                rsData!change_number1 = rsData!change_number1 + rsUnit!ticket_change_number1
                rsData!change_amount1 = rsData!change_amount1 + rsUnit!ticket_change_amount1
                rsData!change_charge1 = rsData!change_charge1 + rsUnit!ticket_change_charge1
                
                rsUnit.MoveNext
                
            Loop
                    
        End If
        
        '求总张数及总票款
        lNumber = 0
        dbAmount = 0
        For j = 1 To TP_TicketTypeCount
            dbAmount = dbAmount + rsData("amount_ticket_type1" & j)
            lNumber = lNumber + rsData("number_ticket_type1" & j)
        Next j
        rsData!total_number1 = lNumber - rsData!return_number1 - rsData!change_number1
        rsData!total_amount1 = dbAmount - rsData!return_amount1 + rsData!return_charge1 - rsData!change_amount1 + rsData!change_charge1
        
        
        '求总张数及总票款
        lNumber = 0
        dbAmount = 0
        For j = 1 To TP_TicketTypeCount
            dbAmount = dbAmount + rsData("amount_ticket_type2" & j)
            lNumber = lNumber + rsData("number_ticket_type2" & j)
        Next j
        rsData!total_number2 = lNumber - rsData!return_number2 - rsData!change_number2
        rsData!total_amount2 = dbAmount - rsData!return_amount2 + rsData!return_charge2 - rsData!change_amount2 + rsData!change_charge2
        
'        rsData!total_number1 = rsData!full_number1 + rsData!half_number1 + rsData!free_number1
'        rsData!total_amount1 = rsData!full_amount1 + rsData!half_amount1 + rsData!free_amount1 - rsData!return_amount1 + rsData!return_charge1 - rsData!change_amount1 + rsData!change_charge1
'
'        rsData!total_number2 = rsData!full_number2 + rsData!half_number2 + rsData!free_number2
'        rsData!total_amount2 = rsData!full_amount2 + rsData!half_amount2 + rsData!free_amount2 - rsData!return_amount2 + rsData!return_charge2 - rsData!change_amount2 + rsData!change_charge2
        
        rsData.Update
    Next
    
    Set ExUnitDateStat = rsData
End Function

'先不管
Public Function DateExUnitStat(UnidID() As String, FromDate As Date, ToDate As Date) As Recordset

    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_DateExUnitStat
    
    
    
End Function



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


Private Function FindUnit(prsData As Recordset, pszUnitID As String) As Boolean
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetRouteTurnOver
    


    If prsData.RecordCount > 0 Then
        pszUnitID = RTrim(pszUnitID)
        prsData.MoveFirst
        Do While Not prsData.EOF
            If RTrim(prsData!unit_id) = pszUnitID Then
                FindUnit = True
                Exit Do
            End If
            prsData.MoveNext
        Loop
    End If
End Function

Private Function GetUnitNameByID(prsData As Recordset, pszUnitID As String) As String


    If prsData.RecordCount > 0 Then
        prsData.MoveFirst
        Do While Not prsData.EOF
            If RTrim(prsData!unit_id) = pszUnitID Then
                GetUnitNameByID = RTrim(prsData!unit_short_name)
            End If
            prsData.MoveNext
        Loop
    End If
End Function

Private Function GetUnitSellCharge(prsData As Recordset, pszUnitID As String) As Double

    If prsData.RecordCount > 0 Then
        prsData.MoveFirst
        Do While Not prsData.EOF
            If RTrim(prsData!unit_id) = pszUnitID Then
                GetUnitSellCharge = FormatDbValue(prsData!sell_charge)
            End If
            prsData.MoveNext
        Loop
    End If
End Function



'互连互售统计结算报表
Public Function UnitInterSettle(UnitID() As String, FromDate As Date, ToDate As Date) As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_TicketStationDim
    AssertHaveRight m_oActiveUser, RIGHT_UnitInterSettle
    
    Dim oDB As New RTConnection
    Dim rsSeller As Recordset, rsUnit As Recordset, rsData As New Recordset
    Dim szSql As String
    Dim nUnitCount As Integer, i As Integer
    Dim rsUnit2 As Recordset
    Dim j As Integer
    Dim lNumber As Long
    Dim dbAmount As Double
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    nUnitCount = ArrayLength(UnitID)
    
    szSql = "SELECT * FROM connect_unit_info "
    If nUnitCount > 0 Then
        szSql = szSql & " WHERE unit_id='" & UnitID(1) & "' "
        For i = 2 To nUnitCount
            szSql = szSql & " OR unit_id='" & UnitID(i) & "' "
        Next
        szSql = szSql & " ORDER BY unit_id "
    End If
    
    Set rsUnit2 = oDB.Execute(szSql)
    '取:售票员每日票务统计表
    
    szSql = "SELECT connect_unit_info.unit_id,unit_short_name , ticket_type," _
    & " SUM(ticket_sale_number) AS ticket_sale_number1, " _
    & " SUM(ticket_sale_amount) AS ticket_sale_amount1," _
    & " SUM(ticket_cancel_number) AS ticket_cancel_number1," _
    & " SUM(ticket_cancel_amount) AS ticket_cancel_amount1," _
    & " SUM(ticket_return_number) AS ticket_return_number1," _
    & " SUM(ticket_return_amount) AS ticket_return_amount1," _
    & " SUM(ticket_return_charge) AS ticket_return_charge1," _
    & " SUM(ticket_change_number) AS ticket_change_number1," _
    & " SUM(ticket_change_amount) AS ticket_change_amount1," _
    & " SUM(ticket_change_charge) AS ticket_change_charge1 " _
    & " FROM Stat_saler_ticket_lst,connect_unit_info WHERE " _
    & " Stat_saler_ticket_lst.unit_id=connect_unit_info.unit_id AND " _
    & " bus_date>='" & ToDBDate(FromDate) & "' AND " _
    & " bus_date<='" & ToDBDate(ToDate) & "' "
    
    
    If nUnitCount > 0 Then
        szSql = szSql & " AND (Stat_saler_ticket_lst.unit_id='" & UnitID(1) & "' "
        For i = 2 To nUnitCount
            szSql = szSql & " OR Stat_saler_ticket_lst.unit_id='" & UnitID(i) & "' "
        Next
        szSql = szSql & ")"
    End If
    szSql = szSql & " GROUP BY connect_unit_info.unit_id,unit_short_name,ticket_type ORDER BY connect_unit_info.unit_id,ticket_type "
    
    Set rsSeller = oDB.Execute(szSql)
    
    '取:车站每日票务统计表
    szSql = "SELECT unit_id ,ticket_type," _
    & " SUM(ticket_sale_number) AS ticket_sale_number1, " _
    & " SUM(ticket_sale_amount) AS ticket_sale_amount1," _
    & " SUM(ticket_cancel_number) AS ticket_cancel_number1," _
    & " SUM(ticket_cancel_amount) AS ticket_cancel_amount1," _
    & " SUM(ticket_return_number) AS ticket_return_number1," _
    & " SUM(ticket_return_amount) AS ticket_return_amount1," _
    & " SUM(ticket_return_charge) AS ticket_return_charge1," _
    & " SUM(ticket_change_number) AS ticket_change_number1," _
    & " SUM(ticket_change_amount) AS ticket_change_amount1," _
    & " SUM(ticket_change_charge) AS ticket_change_charge1 " _
    & " FROM Stat_unit_sell_lst WHERE " _
    & " bus_date>='" & ToDBDate(FromDate) & "' AND " _
    & " bus_date<='" & ToDBDate(ToDate) & "' "
    
    If nUnitCount > 0 Then
        szSql = szSql & " AND (Stat_unit_sell_lst.unit_id='" & UnitID(1) & "' "
        For i = 2 To nUnitCount
            szSql = szSql & " OR Stat_unit_sell_lst.unit_id='" & UnitID(i) & "' "
        Next
        szSql = szSql & ")"
    End If
    szSql = szSql & " GROUP BY unit_id,ticket_type ORDER BY unit_id,ticket_type "
    
    Set rsUnit = oDB.Execute(szSql)
    
    '生成记录集
    With rsData.Fields
        .Append "unit_id", rsSeller("unit_id").Type, 20
        For i = 1 To 2
            '添加各种票种的数量与金额
            For j = 1 To TP_TicketTypeCount
                rsData.Fields.Append "number_ticket_type" & i & j, adInteger
                rsData.Fields.Append "amount_ticket_type" & i & j, adCurrency
            Next j
            
            .Append "Sell_number" & i, adInteger
            .Append "sell_amount" & i, adCurrency
            .Append "sell_charge" & i, adCurrency
            
            .Append "cancel_number" & i, adInteger
            .Append "cancel_amount" & i, adCurrency
            
            .Append "return_number" & i, adInteger
            .Append "return_amount" & i, adCurrency
            .Append "return_charge" & i, adCurrency
            
            .Append "change_number" & i, adInteger
            .Append "change_amount" & i, adCurrency
            .Append "change_charge" & i, adCurrency
            
            .Append "total_number" & i, adInteger
            .Append "total_amount" & i, adCurrency
            
        Next i
        
        .Append "total_number", adInteger
        .Append "total_amount", adCurrency
        
    End With
    rsData.Open

    Dim szUnitTemp As String
    Dim szFieldPrefix As String
    
    For i = 1 To nUnitCount
        rsData.AddNew
        rsData!unit_id = RTrim(UnitID(i)) & "[" & GetUnitNameByID(rsUnit2, UnitID(i)) & "]"
        rsData!sell_charge1 = GetUnitSellCharge(rsUnit2, UnitID(i))
        rsData!sell_charge2 = rsData!sell_charge1
        If FindUnit(rsSeller, UnitID(i)) Then
            Do While Not rsSeller.EOF
                If RTrim(rsSeller!unit_id) <> RTrim(UnitID(i)) Then Exit Do
                

                rsData("number_ticket_type2" & rsSeller!ticket_type) = rsSeller!ticket_sale_number1
                rsData("amount_ticket_type2" & rsSeller!ticket_type) = rsSeller!ticket_sale_amount1

                
                rsData!cancel_number2 = rsData!cancel_number2 + rsSeller!ticket_cancel_number1
                rsData!cancel_amount2 = rsData!cancel_amount2 + rsSeller!ticket_cancel_amount1
                
                rsData!return_number2 = rsData!return_number2 + rsSeller!ticket_return_number1
                rsData!return_amount2 = rsData!return_amount2 + rsSeller!ticket_return_amount1
                rsData!return_charge2 = rsData!return_charge2 + rsSeller!ticket_return_charge1
                
                rsData!change_number2 = rsData!change_number2 + rsSeller!ticket_change_number1
                rsData!change_amount2 = rsData!change_amount2 + rsSeller!ticket_change_amount1
                rsData!change_charge2 = rsData!change_charge2 + rsSeller!ticket_change_charge1
                
                rsSeller.MoveNext
            Loop
        End If
    
        If FindUnit(rsUnit, UnitID(i)) Then
            Do While Not rsUnit.EOF
                If RTrim(rsUnit!unit_id) <> RTrim(UnitID(i)) Then Exit Do

                
                rsData("number_ticket_type1" & rsUnit!ticket_type) = rsUnit!ticket_sale_number1
                rsData("amount_ticket_type1" & rsUnit!ticket_type) = rsUnit!ticket_sale_amount1

                rsData!cancel_number1 = rsData!cancel_number1 + rsUnit!ticket_cancel_number1
                rsData!cancel_amount1 = rsData!cancel_amount1 + rsUnit!ticket_cancel_amount1
                
                rsData!return_number1 = rsData!return_number1 + rsUnit!ticket_return_number1
                rsData!return_amount1 = rsData!return_amount1 + rsUnit!ticket_return_amount1
                rsData!return_charge1 = rsData!return_charge1 + rsUnit!ticket_return_charge1
                
                rsData!change_number1 = rsData!change_number1 + rsUnit!ticket_change_number1
                rsData!change_amount1 = rsData!change_amount1 + rsUnit!ticket_change_amount1
                rsData!change_charge1 = rsData!change_charge1 + rsUnit!ticket_change_charge1
                
                
                rsUnit.MoveNext
                
            Loop
        
        End If
        
        '求总张数及总票款
        lNumber = 0
        dbAmount = 0
        For j = 1 To TP_TicketTypeCount
            dbAmount = dbAmount + rsData("amount_ticket_type1" & j)
            lNumber = lNumber + rsData("number_ticket_type1" & j)
        Next j
        rsData!total_number1 = lNumber - rsData!return_number1 - rsData!change_number1
        rsData!sell_number1 = lNumber
        rsData!sell_amount1 = dbAmount
        
        rsData!sell_charge1 = lNumber * rsData!sell_charge1
        rsData!total_amount1 = dbAmount - rsData!return_amount1 - rsData!change_amount1 - rsData!sell_charge1 '+ rsData!return_charge1 + rsData!change_charge1
        
        
        '求总张数及总票款
        lNumber = 0
        dbAmount = 0
        For j = 1 To TP_TicketTypeCount
            dbAmount = dbAmount + rsData("amount_ticket_type2" & j)
            lNumber = lNumber + rsData("number_ticket_type2" & j)
        Next j
        
        
        rsData!total_number2 = lNumber - rsData!return_number2 - rsData!change_number2
        rsData!sell_number2 = lNumber
        rsData!sell_amount2 = dbAmount
        rsData!sell_charge2 = lNumber * rsData!sell_charge2
        rsData!total_amount2 = dbAmount - rsData!return_amount2 - rsData!change_amount2 - rsData!sell_charge2 '+ rsData!return_charge2 + rsData!change_charge2
        
        rsData!total_number = rsData!total_number2 - rsData!total_number1
        rsData!total_amount = rsData!total_amount2 - rsData!total_amount1
        
        rsData.Update
        
    Next i
    
    Set UnitInterSettle = rsData
    
End Function

'得到日常站点信息
Public Function GetStationDaily(BusDate As Date) As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetStationDaily
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim nCount As Integer
    Dim i As Integer
    oDB.ConnectionString = GetConnectionStr(cszDss)
    '--得到营收情况和各票价项总值
    szSql = "SELECT date,sum(base_price) as base_price_total,sum(price_item_1) as price_item_total1," & _
            " sum(price_item_2) as price_item_total2,sum(price_item_3) as price_item_total3," & _
            " sum(price_item_4) as price_item_total4,sum(price_item_5) as price_item_total5," & _
            " sum(price_item_6) as price_item_total6,sum(price_item_7) as price_item_total7," & _
            " sum(price_item_8) as price_item_total8,sum(price_item_9) as price_item_total9," & _
            " sum(price_item_10) as price_item_total10,sum(price_item_11) as price_item_total11," & _
            " sum(price_item_12) as price_item_total12,sum(price_item_13) as price_item_total13," & _
            " sum(price_item_14) as price_item_total14,sum(price_item_15) as price_item_total15," & _
            " sum(passenger_number) as passenger_number_total,sum(ticket_price) as ticket_price_total," & _
            " sum(ticket_cancel_number) as ticket_cancel_number_total," & _
            " sum(ticket_cancel_amount) as ticket_calcel_amount_total," & _
            " sum(ticket_return_number) as ticket_return_number_total," & _
            " sum(ticket_return_amount) as ticket_return_amount_total," & _
            " sum(ticket_return_charge) as ticket_return_charge_total," & _
            " sum(ticket_change_number) as ticket_change_number_total," & _
            " sum(ticket_change_amount) as ticket_change_amount_total," & _
            " sum(ticket_change_charge) As ticket_change_charge_total " & _
            " FROM Stat_bus_sell_lst " & _
            " WHERE date='" & Format(BusDate, "YYYY-MM-dd") & "'group by date"
Set rsTemp = oDB.Execute(szSql)
Set GetStationDaily = rsTemp
Set oDB = Nothing
Set rsTemp = Nothing
End Function

'得到按检票口统计明细 (计划班次、实开班次、停开班次、始发人数、营收)

Public Function GetCheckGateStatDetail(dyStartDate As Date, dyEndDate As Date, Optional pszSellStationID As String = "") As Recordset
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetCheckGateStatDetail
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim nCount As Integer
    Dim oDB As New RTConnection
    Dim i As Integer
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    szSql = " SELECT (check_gate_name) check_gate_name , SUM(project_bus_number) project_bus_number , SUM( fact_bus_number ) fact_bus_number " _
        & " ,SUM( stop_bus_number )  stop_bus_number  , SUM ( addition_bus_number ) addition_bus_number  , SUM( passenger_number ) passenger_number " _
        & " ,SUM( total_seat_number ) total_seat_number  , SUM( mileage )  mileage , SUM( fact_float_number )  fact_float_number " _
        & " ,SUM( total_float_number)  total_float_number , SUM( addition_passenger_number)  addition_passenger_number " _
        & " ,SUM( total_price )  total_price , SUM(base_price ) base_price , SUM(price_item_1) price_item_1 " _
        & " ,SUM(price_item_2) price_item_2 , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 " _
        & " , SUM(price_item_5) price_item_5 , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 " _
        & " , SUM(price_item_8) price_item_8 , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 " _
        & " , SUM(price_item_11) price_item_11 , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 " _
        & " , SUM(price_item_14) price_item_14 , SUM(price_item_15) price_item_15 " _
        & " , SUM(return_charge) return_charge " _
        & " FROM Stat_checkgate_lst S , Checkgate_info C " _
        & " WHERE BUS_DATE >= '" & ToDBDateTime(dyStartDate) & "' AND BUS_DATE < '" & ToDBDateTime(dyEndDate) & "'" _
        & " And s.check_gate_id = c.check_gate_id " _
        & " AND  NOT ( project_bus_number = 0   AND fact_bus_number = 0    and project_bus_number =0 and fact_bus_number=0 " _
        & " And stop_bus_number = 0 And addition_bus_number = 0 And passenger_number = 0 And total_seat_number = 0 And mileage = 0 " _
        & " and  fact_float_number =0 and total_float_number =0 and addition_passenger_number =0 and total_price =0  ) "
        If pszSellStationID <> "" Then
            Dim szSellStationID() As String
            szSellStationID = Split(pszSellStationID, ",")
            szSql = szSql & " AND (c.sell_station_id='" & szSellStationID(0) & "' "
            For i = 1 To ArrayLength(szSellStationID) - 1
                szSql = szSql & " OR c.sell_station_id='" & szSellStationID(i) & "' "
            Next
            szSql = szSql & ")"
        End If
           szSql = szSql & " GROUP BY c.check_gate_name "

    Set rsTemp = oDB.Execute(szSql)
    Set GetCheckGateStatDetail = rsTemp
    Set rsTemp = Nothing
    Set oDB = Nothing
End Function


'累加统计

Public Function GetCheckGateStatSum(dyStartDate As Date, dyEndDate As Date, Optional pszSellStationID As String = "") As Recordset

    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetCheckGateStatSum
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim nCount As Integer
    Dim oDB As New RTConnection
    Dim i As Integer
    
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    '按售票
    szSql = "SELECT SUM(project_bus_number) project_bus_number  , SUM(fact_bus_number) fact_bus_number " _
        & " , SUM(stop_bus_number) stop_bus_number , SUM(addition_bus_number) addition_bus_number , SUM(passenger_number) passenger_number " _
        & " , SUM(total_seat_number) total_seat_number , SUM(mileage) mileage , SUM(fact_float_number) fact_float_number " _
        & " , SUM(total_float_number) total_float_number , SUM(addition_passenger_number) addition_passenger_number " _
        & " , SUM(total_price) total_price , SUM(base_price) base_price , SUM(price_item_1) price_item_1 " _
        & " , SUM(price_item_2) price_item_2 , SUM(price_item_3) price_item_3 , SUM( price_item_4 ) price_item_4 " _
        & " , SUM(price_item_5) price_item_5 , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 " _
        & " , SUM(price_item_8) price_item_8 , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 " _
        & " , SUM(price_item_11) price_item_11 , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 " _
        & " , SUM(price_item_14) price_item_14 , SUM(price_item_15) price_item_15 , SUM(return_charge) return_charge " _
        & " FROM Stat_checkgate_lst s , Checkgate_info c  " _
        & " WHERE  s.check_gate_id = c.check_gate_id AND BUS_DATE >= '" & ToDBDateTime(dyStartDate) & "' AND BUS_DATE < '" & ToDBDateTime(dyEndDate) & "'"
        
'    '按检票
'    szSql = "SELECT COUNT(bus_id) project_bus_number   " _
'        & " , SUM(passenger_number) passenger_number , 0 mileage " _
'        & " , SUM(total_seat) total_seat_number , SUM(fact_float) fact_float_number " _
'        & " , SUM(total_float) total_float_number " _
'        & " , SUM(ticket_price) total_price , SUM(base_price) base_price , SUM(price_item_1) price_item_1 " _
'        & " , SUM(price_item_2) price_item_2 , SUM(price_item_3) price_item_3 , SUM( price_item_4 ) price_item_4 " _
'        & " , SUM(price_item_5) price_item_5 , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 " _
'        & " , SUM(price_item_8) price_item_8 , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 " _
'        & " , SUM(price_item_11) price_item_11 , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 " _
'        & " , SUM(price_item_14) price_item_14 , SUM(price_item_15) price_item_15 , 0 return_charge " _
'        & " FROM stat_bus_startoff_lst2 s " _
'        & " WHERE DATE >= '" & ToDBDateTime(dyStartDate) & "' AND DATE < '" & ToDBDateTime(dyEndDate) & "'"
'
        
    If pszSellStationID <> "" Then
        Dim szSellStationID() As String
        szSellStationID = Split(pszSellStationID, ",")
        szSql = szSql & " AND (sell_station_id='" & szSellStationID(0) & "' "
        For i = 1 To ArrayLength(szSellStationID) - 1
            szSql = szSql & " OR sell_station_id='" & szSellStationID(i) & "' "
        Next
        szSql = szSql & ")"
    End If
    
    Set rsTemp = oDB.Execute(szSql)
    
    Set GetCheckGateStatSum = rsTemp
    Set rsTemp = Nothing
    Set oDB = Nothing
End Function

'按日期汇总
Public Function GetCheckGateStatDetailSumByDate(dyStartDate As Date, dyEndDate As Date, Optional pszSellStationID As String = "") As Recordset

    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetCheckGateStatDetailSumByDate
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim nCount As Integer
    Dim oDB As New RTConnection
    Dim i As Integer
    
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    szSql = "SELECT bus_date , SUM(project_bus_number) project_bus_number , SUM(project_bus_number) project_bus_number , SUM(fact_bus_number) fact_bus_number " _
        & " , SUM(stop_bus_number) stop_bus_number , SUM(addition_bus_number) addition_bus_number , SUM(passenger_number) passenger_number " _
        & " , SUM(total_seat_number) total_seat_number , SUM(mileage) mileage , SUM(fact_float_number) fact_float_number " _
        & " , SUM(total_float_number) total_float_number , SUM(addition_passenger_number) addition_passenger_number " _
        & " , SUM(total_price) total_price , SUM(base_price) base_price , SUM(price_item_1) price_item_1 " _
        & " , SUM(price_item_2) price_item_2 , SUM(price_item_3) price_item_3 , SUM( price_item_4 ) price_item_4 " _
        & " , SUM(price_item_5) price_item_5 , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 " _
        & " , SUM(price_item_8) price_item_8 , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 " _
        & " , SUM(price_item_11) price_item_11 , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 " _
        & " , SUM(price_item_14) price_item_14 , SUM(price_item_15) price_item_15 , SUM(return_charge) return_charge " _
        & " FROM Stat_checkgate_lst s, Checkgate_info c" _
        & " WHERE s.check_gate_id = c.check_gate_id AND bus_date >= '" & ToDBDateTime(dyStartDate) & "' AND bus_date <='" & ToDBDateTime(dyEndDate) & "'"
        
    If pszSellStationID <> "" Then
        Dim szSellStationID() As String
        szSellStationID = Split(pszSellStationID, ",")
        szSql = szSql & " AND (c.sell_station_id='" & szSellStationID(0) & "' "
        For i = 1 To ArrayLength(szSellStationID) - 1
            szSql = szSql & " OR c.sell_station_id='" & szSellStationID(i) & "' "
        Next
        szSql = szSql & ")"
    End If
        
    szSql = szSql & " GROUP BY bus_date ORDER BY bus_date "
    '这里结束日期用等于是由于未改界面
    
    Set rsTemp = oDB.Execute(szSql)

    Set GetCheckGateStatDetailSumByDate = rsTemp
    Set rsTemp = Nothing
    Set oDB = Nothing
End Function

'得到停班车次
Public Function GetStopBus(BusDate As Date, dyEndDate As Date) As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetStopBus
    
    oDB.ConnectionString = GetConnectionStr("")
    szSql = "SELECT *" _
        & " FROM Env_bus_info e WHERE  bus_date=  '" & ToDBDateTime(BusDate) & "'" _
        & "     AND ( status= 1  OR status= 2 OR status= 8 ) " _
        & "     AND e.bus_id NOT IN " _
        & " ( " _
        & " SELECT E.bus_id FROM Env_bus_info e , bus_info b " _
        & " WHERE e.bus_id = b.bus_id AND stop_start_date <=  '" & ToDBDateTime(BusDate) & "'  AND stop_end_date >= '" & ToDBDate(dyEndDate) & "' " _
        & " ) "
        
        
    Set rsTemp = oDB.Execute(szSql)
    Set GetStopBus = rsTemp
    Set rsTemp = Nothing
    Set oDB = Nothing
        
End Function

'////////////////////////////////
'售票查询--按售票时间线路查询
Public Function GetSellRouteID(pdyStartDate As Date, pdyEndDate As Date, pszRouteID As String, Optional pszSellStationID As String = "") As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim i As Integer
    Dim nLen As Integer
    Dim aszRouteID() As String
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetSellRouteID
    
    nLen = 0
    oDB.ConnectionString = GetConnectionStr(cszDss)
    szSql = "SELECT e.route_id ,r.route_name,COUNT(ticket_id) AS TicketCount,SUM(ticket_price) AS TotalPrice " _
          & "FROM Ticket_sell_lst s,Work_env_bus_info e ,route_info r " _
          & "WHERE (s.operation_time BETWEEN '" & ToDBDateTime(pdyStartDate) & "' AND '" & ToDBDateTime(pdyEndDate) & "') AND " _
          & " (s.status & 28) = 0 AND s.bus_id=e.bus_id AND s.bus_date=e.bus_date AND e.route_id=r.route_id "
    If pszRouteID <> "" Then
        szSql = szSql & " AND e.route_id IN (" & pszRouteID & ")"
    End If
    If pszSellStationID <> "" Then
        szSql = szSql & " AND s.sell_station_id='" & pszSellStationID & "'"
    End If
'    aszRouteID = StringToTeam(pszRouteID)
'    nLen = ArrayLength(aszRouteID)
'    For i = 1 To nLen
'        If i = nLen Then
'            szSql = szSql & " e.route_id='" & aszRouteID(i) & "')"
'        Else
'            szSql = szSql & " e.route_id='" & aszRouteID(i) & "' OR "
'        End If
'
'    Next i
    szSql = szSql & " GROUP BY e.route_id, r.route_name ORDER BY e.route_id"
    
    Set GetSellRouteID = oDB.Execute(szSql)
   
End Function

'//////////////////////////////////
'售票查询--按售票时间车次查询
Public Function GetSellBusID(pdyStartDate As Date, pdyEndDate As Date, pszBusID As String, Optional pszSellStationID As String = "") As Recordset
    Dim szSql As String
    Dim i As Integer
    Dim nLen As Integer
    Dim oDB As New RTConnection
    Dim aszBusID() As String
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetSellBusID
    
    
    nLen = 0
    oDB.ConnectionString = GetConnectionStr(cszDss)
    
    szSql = "SELECT s.bus_id ,e.vehicle_type_name,e.end_station_name station_name ,convert(char(16),e.bus_start_time,120) AS bus_date,COUNT(ticket_id) AS TicketCount,SUM(ticket_price) AS TotalPrice " _
            & "FROM Ticket_sell_lst s,Work_env_bus_info e  " _
            & "WHERE (s.operation_time BETWEEN '" & ToDBDateTime(pdyStartDate) & "' AND '" & ToDBDateTime(pdyEndDate) & "') AND " _
            & " (s.status & 28) = 0 AND s.bus_id=e.bus_id AND s.bus_date=e.bus_date "
    If pszBusID <> "" Then
        szSql = szSql & " AND s.bus_id IN (" & pszBusID & ")"
    End If
    If pszSellStationID <> "" Then
        szSql = szSql & " AND s.sell_station_id='" & pszSellStationID & "'"
    End If
'    aszBusID = StringToTeam(pszBusID)
'    nLen = ArrayLength(aszBusID)
'    For i = 1 To nLen
'        If i = nLen Then
'            szSql = szSql & " s.bus_id='" & aszBusID(i) & "')"
'        Else
'            szSql = szSql & " s.bus_id='" & aszBusID(i) & "' OR "
'        End If
'    Next i
    szSql = szSql & " GROUP BY s.bus_id,e.vehicle_type_name,e.end_station_name,e.bus_start_time ORDER BY s.bus_id"
    
    Set GetSellBusID = oDB.Execute(szSql)
End Function

'///////////////////////////
'售票查询--按售票时间检票口查询
Public Function GetSellCheckGateID(pdyStartDate As Date, pdyEndDate As Date, pszCheckGateID As String, Optional pszSellStationID As String = "") As Recordset
    Dim szSql As String
    Dim i As Integer
    Dim nLen As Integer
    Dim oDB  As New RTConnection
    Dim aszCheckGateID() As String
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetSellCheckGateID
    
    
    
    
    nLen = 0
    oDB.ConnectionString = GetConnectionStr(cszDss)
    
    szSql = "SELECT c.check_gate_id,c.check_gate_name,COUNT(ticket_id) AS TicketCount,SUM(ticket_price) AS TotalPrice " _
            & "FROM Ticket_sell_lst s ,Env_bus_info e,Checkgate_info c " _
            & "WHERE (s.operation_time BETWEEN '" & ToDBDateTime(pdyStartDate) & "' AND '" & ToDBDateTime(pdyEndDate) & "') AND " _
            & " (s.status & 28) = 0 AND  s.bus_id=e.bus_id AND s.bus_date=e.bus_date AND e.check_gate_id=c.check_gate_id "
    If pszCheckGateID <> "" Then
        szSql = szSql & " AND c.check_gate_id IN (" & pszCheckGateID & ")"
    End If
    If pszSellStationID <> "" Then
        szSql = szSql & " AND s.sell_station_id='" & pszSellStationID & "'"
    End If
    
'    aszCheckGateID = StringToTeam(pszCheckGateID)
'    nLen = ArrayLength(aszCheckGateID)
'    For i = 1 To nLen
'        If i = nLen Then
'            szSql = szSql & " e.check_gate_id='" & aszCheckGateID(i) & "' )"
'        Else
'            szSql = szSql & " e.check_gate_id='" & aszCheckGateID(i) & "' OR "
'        End If
'
'    Next
    szSql = szSql & " GROUP BY c.check_gate_id,c.check_gate_name ORDER BY c.check_gate_id"
    
    Set GetSellCheckGateID = oDB.Execute(szSql)
End Function

'///////////////////////////
'售票查询---按售票时间按站点查询
Public Function GetSellStationID(pdyStartDate As Date, pdyEndDate As Date, pszStationID As String, Optional pszSellStationID As String = "") As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim i As Integer
    Dim nLen As Integer
    Dim aszStationID() As String
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetSellStationID
    
    
    
    nLen = 0
    oDB.ConnectionString = GetConnectionStr(cszDss)
    
    szSql = "SELECT s.des_station_id,si.station_name,COUNT(ticket_id) AS TicketCount,SUM(ticket_price) AS TotalPrice " _
            & "FROM Ticket_sell_lst s,station_info si " _
            & "WHERE (s.operation_time BETWEEN '" & ToDBDateTime(pdyStartDate) & "' AND '" & ToDBDateTime(pdyEndDate) & "') AND " _
            & "s.des_station_id=si.station_id AND (s.status & 28) = 0 "
    If pszStationID <> "" Then
        szSql = szSql & " AND s.des_station_id IN (" & pszStationID & ")"
    End If
    If pszSellStationID <> "" Then
        szSql = szSql & " AND s.sell_station_id='" & pszSellStationID & "'"
    End If
'    aszStationID = StringToTeam(pszStationID)
'    nLen = ArrayLength(aszStationID)
'    For i = 1 To nLen
'        If i = nLen Then
'            szSql = szSql & " s.des_station_id='" & aszStationID(i) & "')"
'        Else
'            szSql = szSql & " s.des_station_id='" & aszStationID(i) & "' OR "
'        End If
'    Next
    szSql = szSql & " GROUP BY s.des_station_id,si.station_name ORDER BY des_station_id "
    
    Set GetSellStationID = oDB.Execute(szSql)
End Function



'//////////////////////////////////
'售票即时查询----按发车时间线路查询
Public Function GetBusDateRouteQuery(pdtStartDate As Date, pdtEndDate As Date, pszRouteID As String, Optional pszSellStationID As String = "") As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim i As Integer
    Dim nLen As Integer
    Dim aszRouteID() As String
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusDateRouteQuery
    
    
 On Error GoTo here
    nLen = 0
    oDB.ConnectionString = GetConnectionStr(cszDss)
    szSql = "SELECT e.route_id ,r.route_name,COUNT(ticket_id) AS TicketCount,SUM(ticket_price) AS TotalPrice " _
          & " FROM Ticket_sell_lst s,Work_env_bus_info e ,route_info r " _
          & " WHERE (s.bus_date BETWEEN '" & ToDBDate(pdtStartDate) & "' AND '" & ToDBDate(pdtEndDate) & "') AND " _
          & " (s.status & 28) = 0 AND s.bus_id=e.bus_id AND s.bus_date=e.bus_date AND e.route_id=r.route_id  "
    If pszRouteID <> "" Then
        szSql = szSql & " AND e.route_id in (" & pszRouteID & ")"
    End If
    If pszSellStationID <> "" Then
       szSql = szSql & " AND s.sell_station_id='" & pszSellStationID & "'"
    End If
    
'    aszRouteID = StringToTeam(pszRouteID)
'    nLen = ArrayLength(aszRouteID)
'    For i = 1 To nLen
'        If i = nLen Then
'            szSql = szSql & " e.route_id='" & aszRouteID(i) & "')"
'        Else
'            szSql = szSql & " e.route_id='" & aszRouteID(i) & "' OR "
'        End If
'
'    Next i
    szSql = szSql & " GROUP BY e.route_id, r.route_name ORDER BY e.route_id"
    
    Set GetBusDateRouteQuery = oDB.Execute(szSql)
Exit Function
here:
    err.Raise err.Number
End Function

'////////////////////////////
'售票即时查询----按发车时间检票口查询
Public Function GetBusDateCheckGateQuery(pdtStartDate As Date, pdtEndDate As Date, pszCheckGateID As String, Optional pszSellStationID As String = "") As Recordset
    Dim szSql As String
    Dim i As Integer
    Dim nLen As Integer
    Dim oDB  As New RTConnection
    Dim aszCheckGateID() As String
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusDateCheckGateQuery
    
    nLen = 0
    oDB.ConnectionString = GetConnectionStr(cszDss)
    
    szSql = "SELECT c.check_gate_id,c.check_gate_name,COUNT(ticket_id) AS TicketCount,SUM(ticket_price) AS TotalPrice " _
            & "FROM Ticket_sell_lst s ,Env_bus_info e,Checkgate_info c " _
            & "WHERE (s.bus_date BETWEEN '" & ToDBDate(pdtStartDate) & "' AND '" & ToDBDate(pdtEndDate) & "') AND " _
            & " (s.status & 28) = 0 AND  s.bus_id=e.bus_id AND s.bus_date=e.bus_date AND e.check_gate_id=c.check_gate_id "
    If pszCheckGateID <> "" Then
        szSql = szSql & " AND c.check_gate_id IN (" & pszCheckGateID & ")"
    End If
    If pszSellStationID <> "" Then
        szSql = szSql & " AND s.sell_station_id='" & pszSellStationID & "'"
    End If
'    aszCheckGateID = StringToTeam(pszCheckGateID)
'    nLen = ArrayLength(aszCheckGateID)
'    For i = 1 To nLen
'        If i = nLen Then
'            szSql = szSql & " e.check_gate_id='" & aszCheckGateID(i) & "' )"
'        Else
'            szSql = szSql & " e.check_gate_id='" & aszCheckGateID(i) & "' OR "
'        End If
'
'    Next
    szSql = szSql & " GROUP BY c.check_gate_id,c.check_gate_name ORDER BY c.check_gate_id"
    
    Set GetBusDateCheckGateQuery = oDB.Execute(szSql)
    
End Function


'////////////////////////////
'售票即时查询----按发车时间站点查询
Public Function GetBusDateStation(pdtStartDate As Date, pdtEndDate As Date, pszStationID As String, Optional pszSellStationID As String = "") As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim i As Integer
    Dim nLen As Integer
    Dim aszStationID() As String
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusDateStation
    
    
    nLen = 0
    oDB.ConnectionString = GetConnectionStr(cszDss)
    szSql = "SELECT s.des_station_id,si.station_name,COUNT(ticket_id) AS TicketCount,SUM(ticket_price) AS TotalPrice " _
            & "FROM Ticket_sell_lst s,station_info si " _
            & "WHERE (s.bus_date BETWEEN '" & ToDBDate(pdtStartDate) & "' AND '" & ToDBDate(pdtEndDate) & "') AND " _
            & "s.des_station_id=si.station_id AND (s.status & 28) = 0 "
    If pszStationID <> "" Then
        szSql = szSql & " AND s.des_station_id IN (" & pszStationID & ")"
    End If
    If pszSellStationID <> "" Then
        szSql = szSql & " AND s.sell_station_id='" & pszSellStationID & "'"
    End If
'    aszStationID = StringToTeam(pszStationID)
'    nLen = ArrayLength(aszStationID)
'    For i = 1 To nLen
'        If i = nLen Then
'            szSql = szSql & " s.des_station_id='" & aszStationID(i) & "')"
'        Else
'            szSql = szSql & " s.des_station_id='" & aszStationID(i) & "' OR "
'        End If
'    Next
    szSql = szSql & " GROUP BY s.des_station_id,si.station_name ORDER BY des_station_id "
    Set GetBusDateStation = oDB.Execute(szSql)
End Function
'/////////////////////////////
'售票即时查询---按发车日期车次查询
Public Function GetBusDateBusIDQuery(pdtStartDate As Date, pdtEndDate As Date, pszBusID As String, Optional pszSellStationID As String = "") As Recordset
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    Dim i As Integer
    Dim nLen As Integer
    Dim aszBusID() As String
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusDateBusIDQuery
    
On Error GoTo here
    nLen = 0
    oDB.ConnectionString = GetConnectionStr(cszDss)
    
    szSql = "SELECT s.bus_id ,s.bus_date,si.station_name ,COUNT(ticket_id) as TicketCount,SUM(ticket_price) AS TotalPrice " _
                & " FROM Ticket_sell_lst s,station_info si " _
                & " WHERE (s.status & 28) = 0 AND (s.bus_date BETWEEN '" & ToDBDate(pdtStartDate) & "' AND '" & ToDBDate(pdtEndDate) & "') AND s.des_station_id=si.station_id "
    If pszBusID <> "" Then
        szSql = szSql & " AND s.bus_id IN (" & pszBusID & ")"
    End If
    If pszSellStationID <> "" Then
        szSql = szSql & " AND s.sell_station_id='" & pszSellStationID & "'"
    End If
'    aszBusID = StringToTeam(pszBusID)
'    nLen = ArrayLength(aszBusID)
'
'    For i = 1 To nLen
'        If i = nLen Then
'            szSql = szSql & "s.bus_id='" & aszBusID(i) & "') "
'        Else
'            szSql = szSql & "s.bus_id='" & aszBusID(i) & "' OR "
'        End If
'    Next i
    szSql = szSql & " GROUP BY s.bus_id,s.bus_date,si.station_name ORDER BY s.bus_id"
    Set rsTemp = oDB.Execute(szSql)
    Set GetBusDateBusIDQuery = rsTemp
    Set rsTemp = Nothing
    Exit Function
here:
  Set rsTemp = Nothing
  err.Raise err.Number
End Function

'省内省外人数、营收、周转量汇总
Public Function ProvinceSum(pdtFromDate As Date, pdtToDate As Date, pszAreaCode As String, Optional pszSellStationID As String = "") As Double()
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_ProvinceSum
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    Dim adbTemp() As Double
    Dim i As Integer
    oDB.ConnectionString = GetConnectionStr(cszDss)
        If pszAreaCode <> "" Then
            szWhere = " AND si.area_code IN ( " & pszAreaCode & ") "
        End If
        If pszSellStationID <> "" Then
            Dim szSellStationID() As String
            szSellStationID = Split(pszSellStationID, ",")
            szWhere = szWhere & " AND (s.sell_station_id='" & szSellStationID(0) & "' "
            For i = 1 To ArrayLength(szSellStationID) - 1
                szWhere = szWhere & " OR s.sell_station_id='" & szSellStationID(i) & "' "
            Next
            szWhere = szWhere & ")"
        End If
        szSql = "   SELECT  a.province_in_out " _
            & "    , (SUM(pASsenger_number) -SUM(ticket_cancel_number)-SUM(ticket_return_number)-SUM(ticket_change_number)) AS pASsenger_number  " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) + SUM(ticket_return_charge)  ) AS total_ticket_price  " _
            & "    , SUM(ticket_cancel_number) ticket_cancel_number , SUM(ticket_cancel_amount) ticket_cancel_amount " _
            & "    , SUM(ticket_return_number) ticket_return_number , SUM(ticket_return_amount) ticket_return_amount " _
            & "    , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_number) ticket_change_number " _
            & "    , SUM(ticket_change_amount) ticket_change_amount , SUM(ticket_change_charge) ticket_change_charge " _
            & "    , SUM(fact_float_number) fact_float_number " _
            & " FROM  Stat_station_sell_lst s , station_info si , area_code a " _
            & " WHERE date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "' " _
            & " AND si.station_id = s.station_id " _
            & " AND a.area_code = si.area_code " _
            & szWhere _
            & " GROUP BY province_in_out "
        
    Set rsTemp = oDB.Execute(szSql)
    
    If rsTemp.RecordCount > 0 Then
        ReDim adbTemp(1 To rsTemp.RecordCount, 1 To 4)
        For i = 1 To rsTemp.RecordCount
            adbTemp(i, 1) = rsTemp!province_in_out
            adbTemp(i, 2) = rsTemp!passenger_number
            adbTemp(i, 3) = rsTemp!total_ticket_price
            adbTemp(i, 4) = rsTemp!fact_float_number
            rsTemp.MoveNext
        Next i
        
    End If
    
    ProvinceSum = adbTemp
    
End Function



'/////////////////////
'车票发售计划
Public Function BusSellScheme(pdtStartDate As Date, pdtEndDate As Date) As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nPreDay As Integer
    
    Dim oSystemParam As New SystemParam
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_BusSellScheme
    
    
    
    oSystemParam.Init m_oActiveUser
    nPreDay = oSystemParam.PreSellDate
    '取出每天的除去预售票外的可售票数及总座位.
    oDB.ConnectionString = GetConnectionStr(cszDss)
       
    szSql = "SELECT a.bus_date AS BusDate,a.TotalBusCount,a.TotalSeat ,b.TotalSellSeat FROM  " _
        & " (SELECT en.bus_date,COUNT(bus_id) AS TotalBusCount,SUM(total_seat) AS TotalSeat FROM " & IIf(DateDiff("d", pdtStartDate, Date) < nPreDay, "Work_env_bus_info", "Env_bus_info") & " en" _
        & " WHERE bus_date BETWEEN '" & ToDBDate(pdtStartDate) & "' AND '" & ToDBDate(pdtEndDate) & "' GROUP BY bus_date) a, " _
        & " (SELECT a1.bus_date,a1.totalseatcount-ISNULL(a2.totalpreseat,0) AS TotalSellSeat FROM " _
            & " (SELECT bus_date ,SUM(total_seat) AS totalseatcount FROM " & IIf(DateDiff("d", pdtStartDate, Date) < nPreDay, "Work_env_bus_info", "Env_bus_info") _
            & " WHERE bus_date BETWEEN '" & ToDBDate(pdtStartDate) & "' AND '" & ToDBDate(pdtEndDate) & "' GROUP BY bus_date) a1, " _
            & " (SELECT bus_date, COUNT(ticket_id) AS totalpreseat FROM Ticket_sell_lst " _
            & " WHERE operation_time <'" & ToDBDate(pdtStartDate) & "' " _
            & " AND bus_date BETWEEN '" & ToDBDate(pdtStartDate) & "' AND '" & ToDBDate(pdtEndDate) & "'  GROUP BY bus_date) a2 " _
        & " Where a1.bus_date *= a2.bus_date ) b " _
        & " Where a.bus_date = b.bus_date " _

    Set BusSellScheme = oDB.Execute(szSql)
    
    Set oDB = Nothing
End Function

'///////////////////////////
'车票发售计划
Public Function BusTicketSellScheme(pdtStartDate As Date, pdtEndDate As Date) As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_BusTicketSellScheme
    
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    
    szSql = "SELECT a.bus_date AS BusDate,a.TotalBusCount,a.TotalSeat ,b.TotalSellSeat FROM " _
            & "(SELECT en.bus_date,COUNT(bus_id) AS TotalBusCount,SUM(total_seat) AS TotalSeat FROM Work_env_bus_info en " _
            & " WHERE bus_date BETWEEN '" & ToDBDate(pdtStartDate) & "' AND '" & ToDBDate(pdtEndDate) & "'" _
            & "GROUP BY bus_date) a, " _
            & " (SELECT a1.bus_date,a1.totalseatcount-a2.totalpreseat AS TotalSellSeat FROM " _
            & "(SELECT bus_date ,SUM(total_seat) AS totalseatcount FROM Work_env_bus_info GROUP BY bus_date) a1," _
            & "(SELECT bus_date, COUNT(ticket_id) AS totalpreseat FROM Ticket_sell_lst WHERE bus_date<>SUBSTRING(CONVERT(CHAR,operation_time,20),1,10) GROUP BY bus_date ) a2 " _
            & " Where a1.bus_date = a2.bus_date AND (a.bus_date BETWEEN '" & ToDBDate(pdtStartDate) & "' AND '" & ToDBDate(pdtEndDate) & "' )" _
            & " ) b " _
            & " Where a.bus_date = b.bus_date "
            
    Set BusTicketSellScheme = oDB.Execute(szSql)
    
    Set oDB = Nothing
    
End Function




' 调度即时查询 ----按发车时间和站点查询座位、票价情况决定是否需要加班
Public Function GetBusDateSeatSellInfo(pdtStartDate As Date, pdtEndDate As Date, pszStationID As String, Optional nSel As Integer) As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim szTemp As String
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusDateSeatSellInfo
    
    
 oDB.ConnectionString = GetConnectionStr(cszDss)
 szTemp = "if exists("
 szSql = " SELECT ew.bus_id,ss.sell_station_name , ew.status ,ew.bus_start_time,bt.bus_type_name, ew.end_station_name," _
       & " ew.total_seat as totalSeat, " _
       & " (RTRIM(ew.total_seat-ew.sale_seat_quantity)+'\'+RTRIM(ew.sale_seat_quantity) )as leave_seat_id,ew.total_seat, ew.sale_seat_quantity ,ew.total_seat-ew.sale_seat_quantity have_sale_seat_quantity," _
       & "  st.seat_type_id, st.seat_type_name," _
       & "  vs.vehicle_type_short_name," _
       & "  RTRIM(es.full_price) as full_price," _
       & "  RTRIM(es.half_price) as half_price, " _
       & "  RTRIM(es.preferential_ticket1) as preferential_ticket1," _
       & "  RTRIM(es.preferential_ticket2) as preferential_ticket2," _
       & "  RTRIM(es.preferential_ticket3) As preferential_ticket3, " _
       & "  ts.ticket_id,CASE ts.status WHEN 1 THEN '正常售出' WHEN 2 THEN '改签售出' WHEN 5" _
       & "  THEN '废票' WHEN 9 THEN '被改签' WHEN 17 THEN '退票'WHEN 33 THEN '已检' ELSE convert(char(2),ts.status) END status_name"

szSql = szSql & " FROM Work_env_bus_info ew,Work_env_bus_station_lst es ," _
      & "   Bus_type_code bt,Seat_type_code st,vehicle_type_code vs,sell_station_info ss,ticket_sell_lst ts " _
      & "   WHERE es.max_sale_quantity <> 0 " _
      & "   and ew.bus_date>='" & ToDBDate(pdtStartDate) & "'" _
      & "   and ew.bus_date<'" & ToDBDate(DateAdd("d", 1, pdtStartDate)) & "'" _
      & "   And ew.bus_type = bt.bus_type" _
      & "   And ew.bus_id = es.bus_id " _
      & "   and ew.bus_date=es.bus_date " _
      & "   and es.seat_type_id=st.seat_type_id  " _
      & "   and ss.sell_station_id=es.sell_station_id" _
      & "   and vs.vehicle_type_code=ew.vehicle_type_code" _
      & "   and es.station_id='" & pszStationID & "' " _
      & "   and ts.sell_station_id=es.sell_station_id and ts.bus_id=ew.bus_id and ew.bus_date=ts.bus_date"
'      & " order by ew.bus_start_time,ew.bus_id,st.seat_type_id"
      '& " and (convert( char(8),ew.bus_start_time,108) Between '" & ToDBTime(pdtStartDate) & "' and  '" & ToDBTime(pdtEndDate) & "') "
szTemp = szTemp & szSql & ") begin" & szSql & " order by es.sell_station_id end else"

szSql = " SELECT ew.bus_id,ss.sell_station_name , ew.status ,ew.bus_start_time,bt.bus_type_name, ew.end_station_name," _
       & " ew.total_seat as totalSeat, " _
       & " (RTRIM(ew.total_seat-ew.sale_seat_quantity)+'\'+RTRIM(ew.sale_seat_quantity) )as leave_seat_id,ew.total_seat, ew.sale_seat_quantity ,ew.total_seat-ew.sale_seat_quantity have_sale_seat_quantity," _
       & "  st.seat_type_id, st.seat_type_name," _
       & "  vs.vehicle_type_short_name," _
       & "  RTRIM(es.full_price) as full_price," _
       & "  RTRIM(es.half_price) as half_price, " _
       & "  RTRIM(es.preferential_ticket1) as preferential_ticket1," _
       & "  RTRIM(es.preferential_ticket2) as preferential_ticket2," _
       & "  RTRIM(es.preferential_ticket3) As preferential_ticket3 "
szSql = szSql & " FROM Work_env_bus_info ew,Work_env_bus_station_lst es ," _
      & "   Bus_type_code bt,Seat_type_code st,vehicle_type_code vs,sell_station_info ss " _
      & "   WHERE es.max_sale_quantity <> 0 " _
      & "   and ew.bus_date>='" & ToDBDate(pdtStartDate) & "'" _
      & "   and ew.bus_date<'" & ToDBDate(DateAdd("d", 1, pdtStartDate)) & "'" _
      & "   And ew.bus_type = bt.bus_type" _
      & "   And ew.bus_id = es.bus_id " _
      & "   and ew.bus_date=es.bus_date " _
      & "   and es.seat_type_id=st.seat_type_id  " _
      & "   and ss.sell_station_id=es.sell_station_id" _
      & "   and vs.vehicle_type_code=ew.vehicle_type_code" _
      & "   and es.station_id='" & pszStationID & "' " _
      & " order by ew.bus_start_time,ew.bus_id,st.seat_type_id"
szTemp = szTemp & szSql
Set GetBusDateSeatSellInfo = oDB.Execute(szTemp)
'ew.bus_start_time Between  '" & ToDBTime(pdtStartDate) & "' and   '" & ToDBTime(pdtEndDate) & "'
Exit Function

End Function
' 调度即时查询 ----按发车时间和车次查询座位、票价情况决定是否需要加班
Public Function GetBusDateFromBusIDSeatSellInfo(pdtStartDate As Date, pdtEndDate As Date, pBusID As String, Optional nSel As Integer) As Recordset
    Dim szSql As String
    Dim szTemp As String
    Dim oDB As New RTConnection
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusDateFromBusIDSeatSellInfo
    
    
    
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
 szTemp = "if exists("
 szSql = " select ew.bus_id, ss.sell_station_name ,ew.status ,ew.bus_start_time,bt.bus_type_name, ew.end_station_name," _
       & " ew.total_seat as totalSeat, " _
       & " (RTRIM(ew.total_seat-ew.sale_seat_quantity)+'\'+RTRIM(ew.sale_seat_quantity) )as leave_seat_id, ew.total_seat, ew.sale_seat_quantity ,ew.total_seat-ew.sale_seat_quantity have_sale_seat_quantity, " _
       & "  st.seat_type_id, st.seat_type_name," _
       & "  vs.vehicle_type_short_name," _
       & "  RTRIM(es.full_price) as full_price," _
       & "  RTRIM(es.half_price) as half_price, " _
       & "  RTRIM(es.preferential_ticket1) as preferential_ticket1," _
       & "  RTRIM(es.preferential_ticket2) as preferential_ticket2," _
       & "  RTRIM(es.preferential_ticket3) As preferential_ticket3," _
       & "  ts.ticket_id,CASE ts.status WHEN 1 THEN '正常售出' WHEN 2 THEN '改签售出' WHEN 5" _
       & "  THEN '废票' WHEN 9 THEN '被改签' WHEN 17 THEN '退票'WHEN 33 THEN '已检' ELSE convert(char(2),ts.status) END status_name"
szSql = szSql & " from Work_env_bus_info ew,Work_env_bus_station_lst es ," _
      & "   Bus_type_code bt,Seat_type_code st,vehicle_type_code vs ,ticket_sell_lst ts,sell_station_info ss " _
      & "   Where  " _
      & "   ew.bus_date>='" & ToDBDate(pdtStartDate) & "'" _
      & "   and ew.bus_date<'" & ToDBDate(DateAdd("d", 1, pdtStartDate)) & "'" _
      & "   and ew.bus_type = bt.bus_type" _
      & "   And ew.bus_id = es.bus_id " _
      & "   and ew.bus_date=es.bus_date " _
      & "   and es.seat_type_id=st.seat_type_id  " _
      & "   and vs.vehicle_type_code=ew.vehicle_type_code" _
      & "   and es.station_id=ew.end_station_id " _
      & "   and es.bus_id='" & pBusID & "' " _
      & "   and ss.sell_station_id=es.sell_station_id" _
      & "   and ts.sell_station_id=es.sell_station_id and ts.bus_id=ew.bus_id and ew.bus_date=ts.bus_date"
    '      & " order by ew.bus_start_time,ew.bus_id,st.seat_type_id"
      '& " and (convert( char(8),ew.bus_start_time,108) Between '" & ToDBTime(pdtStartDate) & "' and  '" & ToDBTime(pdtEndDate) & "') "
szTemp = szTemp & szSql & ") begin" & szSql & " order by es.sell_station_id end else"
 szSql = " select ew.bus_id, ss.sell_station_name ,ew.status ,ew.bus_start_time,bt.bus_type_name, ew.end_station_name," _
       & " ew.total_seat as totalSeat, " _
       & " (RTRIM(ew.total_seat-ew.sale_seat_quantity)+'\'+RTRIM(ew.sale_seat_quantity) )as leave_seat_id, ew.total_seat, ew.sale_seat_quantity ,ew.total_seat-ew.sale_seat_quantity have_sale_seat_quantity, " _
       & "  st.seat_type_id, st.seat_type_name," _
       & "  vs.vehicle_type_short_name," _
       & "  RTRIM(es.full_price) as full_price," _
       & "  RTRIM(es.half_price) as half_price, " _
       & "  RTRIM(es.preferential_ticket1) as preferential_ticket1," _
       & "  RTRIM(es.preferential_ticket2) as preferential_ticket2," _
       & "  RTRIM(es.preferential_ticket3) As preferential_ticket3 "
szSql = szSql & " from Work_env_bus_info ew,Work_env_bus_station_lst es ," _
      & "   Bus_type_code bt,Seat_type_code st,vehicle_type_code vs ,sell_station_info ss " _
      & "   Where  " _
      & "   ew.bus_date>='" & ToDBDate(pdtStartDate) & "'" _
      & "   and ew.bus_date<'" & ToDBDate(DateAdd("d", 1, pdtStartDate)) & "'" _
      & "   and ew.bus_type = bt.bus_type" _
      & "   And ew.bus_id = es.bus_id " _
      & "   and ew.bus_date=es.bus_date " _
      & "   and es.seat_type_id=st.seat_type_id  " _
      & "   and vs.vehicle_type_code=ew.vehicle_type_code" _
      & "   and es.station_id=ew.end_station_id " _
      & "   and es.bus_id='" & pBusID & "' " _
      & "   and es.sell_station_id=ss.sell_station_id" _
      & " order by ew.bus_start_time,ew.bus_id,st.seat_type_id"
szTemp = szTemp & szSql
      Set GetBusDateFromBusIDSeatSellInfo = oDB.Execute(szTemp)
End Function


'调度即时查询 ----按发车时间和线路查询座位、票价情况决定是否需要加班
Public Function GetBusDateFromRoutIdSeatSellInfo(pdtStartDate As Date, pdtEndDate As Date, szRouteID As String, Optional nSel As Integer) As Recordset
    
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim szTemp As String
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
'    AssertHaveRight m_oActiveUser, RIGHT_GetBusDateFromRoutIdSeatSellInfo
    
    
     oDB.ConnectionString = GetConnectionStr(cszDss)
 szTemp = "if exists("
 szSql = " select ew.bus_id, ss.sell_station_name ,ew.status ,ew.bus_start_time,bt.bus_type_name, ew.end_station_name," _
       & " ew.total_seat as totalSeat, " _
       & " (RTRIM(ew.total_seat-ew.sale_seat_quantity)+'\'+RTRIM(ew.sale_seat_quantity) )as leave_seat_id,ew.total_seat, ew.sale_seat_quantity ,ew.total_seat-ew.sale_seat_quantity have_sale_seat_quantity," _
       & "  st.seat_type_id, st.seat_type_name," _
       & "  vs.vehicle_type_short_name," _
       & "  RTRIM(es.full_price) as full_price," _
       & "  RTRIM(es.half_price) as half_price, " _
       & "  RTRIM(es.preferential_ticket1) as preferential_ticket1," _
       & "  RTRIM(es.preferential_ticket2) as preferential_ticket2," _
       & "  RTRIM(es.preferential_ticket3) As preferential_ticket3 " _
       & "  ts.ticket_id,CASE ts.status WHEN 1 THEN '正常售出' WHEN 2 THEN '改签售出' WHEN 5" _
       & "  THEN '废票' WHEN 9 THEN '被改签' WHEN 17 THEN '退票'WHEN 33 THEN '已检' ELSE convert(char(2),ts.status) END status_name"

szSql = szSql & " from Work_env_bus_info ew,Work_env_bus_station_lst es ," _
      & "   Bus_type_code bt,Seat_type_code st,vehicle_type_code vs,sell_station_info ss,ticket_sell_lst ts  " _
      & "   Where " _
      & "    ew.bus_date>='" & ToDBDate(pdtStartDate) & "'" _
      & "   and ew.bus_date<'" & ToDBDate(DateAdd("d", 1, pdtStartDate)) & "'" _
      & "   And ew.bus_type = bt.bus_type" _
      & "   And ew.bus_id = es.bus_id " _
      & "   and ew.bus_date=es.bus_date " _
      & "   and es.seat_type_id=st.seat_type_id  " _
      & "   and vs.vehicle_type_code=ew.vehicle_type_code" _
      & "   and es.station_id=ew.end_station_id " _
      & "   and ss.sell_station_id=es.sell_station_id" _
      & "   and ew.Route_id='" & szRouteID & "' " _
      & "   and ts.sell_station_id=es.sell_station_id and ts.bus_id=ew.bus_id and ew.bus_date=ts.bus_date"
'      & " order by ew.bus_start_time,ew.bus_id,st.seat_type_id"
      '& " and (convert( char(8),ew.bus_start_time,108) Between '" & ToDBTime(pdtStartDate) & "' and  '" & ToDBTime(pdtEndDate) & "') "

szTemp = szTemp & szSql & ") begin" & szSql & " order by es.sell_station_id end else"

 szSql = " select ew.bus_id, ss.sell_station_name ,ew.status ,ew.bus_start_time,bt.bus_type_name, ew.end_station_name," _
       & " ew.total_seat as totalSeat, " _
       & " (RTRIM(ew.total_seat-ew.sale_seat_quantity)+'\'+RTRIM(ew.sale_seat_quantity) )as leave_seat_id,ew.total_seat, ew.sale_seat_quantity ,ew.total_seat-ew.sale_seat_quantity have_sale_seat_quantity," _
       & "  st.seat_type_id, st.seat_type_name," _
       & "  vs.vehicle_type_short_name," _
       & "  RTRIM(es.full_price) as full_price," _
       & "  RTRIM(es.half_price) as half_price, " _
       & "  RTRIM(es.preferential_ticket1) as preferential_ticket1," _
       & "  RTRIM(es.preferential_ticket2) as preferential_ticket2," _
       & "  RTRIM(es.preferential_ticket3) As preferential_ticket3, "
szSql = szSql & " from Work_env_bus_info ew,Work_env_bus_station_lst es ," _
      & "   Bus_type_code bt,Seat_type_code st,vehicle_type_code vs,sell_station_info ss " _
      & "   Where " _
      & "    ew.bus_date>='" & ToDBDate(pdtStartDate) & "'" _
      & "   and ew.bus_date<'" & ToDBDate(DateAdd("d", 1, pdtStartDate)) & "'" _
      & "   And ew.bus_type = bt.bus_type" _
      & "   And ew.bus_id = es.bus_id " _
      & "   and ew.bus_date=es.bus_date " _
      & "   and es.seat_type_id=st.seat_type_id  " _
      & "   and vs.vehicle_type_code=ew.vehicle_type_code" _
      & "   and es.station_id=ew.end_station_id " _
      & "   and ss.sell_station_id=es.sell_station_id" _
      & "   and ew.Route_id='" & szRouteID & "' " _
      & " order by ew.bus_start_time,ew.bus_id,st.seat_type_id"
      '& " and (convert( char(8),ew.bus_start_time,108) Between '" & ToDBTime(pdtStartDate) & "' and  '" & ToDBTime(pdtEndDate) & "') "
szTemp = szTemp & szSql
Set GetBusDateFromRoutIdSeatSellInfo = oDB.Execute(szTemp)
'ew.bus_start_time Between  '" & ToDBTime(pdtStartDate) & "' and   '" & ToDBTime(pdtEndDate) & "'
Exit Function  '" & ToDBTime(pdtStartDate) & "' and   '" & ToDBTime(pdtEndDate) & "'
End Function


'网上售票统计
Public Function InternetTkCount(UnitID() As String, FromDate As Date, ToDate As Date, bSellCount As Boolean, Optional nSellStationId As String = "") As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_TicketStationDim
    AssertHaveRight m_oActiveUser, RIGHT_ExUnitDateStat
    'bSellCount售票时间
    Dim oDB As New RTConnection
    Dim rsSell As Recordset, rsCancel As Recordset, rsData As New Recordset
    Dim szSql As String
    Dim nUnitCount As Integer, i As Integer
    Dim rsUnit2 As Recordset
    Dim j As Integer
    Dim lNumber As Long
    Dim dbAmount As Double
    Dim dbHandleAmount As Double
    oDB.ConnectionString = GetConnectionStr(cszDss)
    nUnitCount = ArrayLength(UnitID)
    
    szSql = "SELECT * FROM user_info "
    If nUnitCount > 0 Then
        szSql = szSql & " WHERE user_id='" & UnitID(1) & "' "
        For i = 2 To nUnitCount
            szSql = szSql & " OR user_id='" & UnitID(i) & "' "
        Next
        szSql = szSql & " ORDER BY user_id "
    End If
    
    Set rsUnit2 = oDB.Execute(szSql)

    szSql = "Select s.user_id,sum(cast(pay_count as money)) pay,count(*) num,ticket_type_name,ticket_type,sum(i.handle_charge) handle_charge,has_child " _
    & " from internet_sell_ticket_info i,user_info s,ticket_type_code tc" _
    & " where  i.bank_id=s.user_id and tc.ticket_type_id=i.ticket_type and (if_print=0 or if_print=1 or if_print=4 ) "
    If bSellCount = True Then
        szSql = szSql & " and busy_date>='" & ToDBDate(FromDate) & "' and busy_date<=dateadd(d,1,'" & ToDBDate(ToDate) & "') "
    Else
        szSql = szSql & " and print_time>='" & ToDBDate(FromDate) & "' and print_time<=dateadd(d,1,'" & ToDBDate(ToDate) & "') "
    End If
    If nUnitCount > 0 Then
        szSql = szSql & " AND (s.user_id='" & UnitID(1) & "' "
        For i = 2 To nUnitCount
            szSql = szSql & " OR s.user_id='" & UnitID(i) & "' "
        Next
        szSql = szSql & ")"
    End If
    
    If nSellStationId <> "" Then
        szSql = szSql & " AND i.sell_station_id= " & TransFieldValueToString(nSellStationId)
    End If
    szSql = szSql & " group by s.user_id,ticket_type,ticket_type_name,has_child"
    
    Set rsSell = oDB.Execute(szSql)
    
    
    
    
    szSql = "Select s.user_id,sum(cast(pay_count as money)) pay,count(*) num,ticket_type_name,ticket_type,sum(i.handle_charge) handle_charge ,sum(r.handle_charge) return_handle_charge,has_child " _
    & " from internet_sell_ticket_info i,user_info s,ticket_type_code tc,internet_return_ticket_info r" _
    & " where   i.bank_id=s.user_id and tc.ticket_type_id=i.ticket_type and (if_print=4) and i.id=r.id and i.getticket_id=r.getticket_id and i.validate_id=r.validate_id "
'    If bSellCount = True Then
'        szSql = szSql & " and i.busy_date>='" & ToDBDate(FromDate) & "' and i.busy_date<=dateadd(d,1,'" & ToDBDate(ToDate) & "') "
'    Else
'        szSql = szSql & " and print_time>='" & ToDBDate(FromDate) & "' and print_time<=dateadd(d,1,'" & ToDBDate(ToDate) & "') "
'    End If
    If nUnitCount > 0 Then
        szSql = szSql & " AND (s.user_id='" & UnitID(1) & "' "
        For i = 2 To nUnitCount
            szSql = szSql & " OR s.user_id='" & UnitID(i) & "' "
        Next
        szSql = szSql & ")"
    End If
    szSql = szSql & " and r.operation_time>='" & ToDBDate(FromDate) & "' and r.operation_time<=dateadd(d,1,'" & ToDBDate(ToDate) & "') "
    szSql = szSql & " group by s.user_id,ticket_type,ticket_type_name,has_child"
    
    Set rsCancel = oDB.Execute(szSql)
    '生成记录集
    With rsData.Fields
        .Append "user_id", adChar, 20
        For i = 1 To 1
            '添加各种票种的数量与金额
            For j = 1 To TP_TicketTypeCount
                rsData.Fields.Append "number_ticket_type" & i & j, adInteger
                rsData.Fields.Append "amount_ticket_type" & i & j, adCurrency
            Next j
            
            .Append "cancel_number" & i, adInteger
            .Append "cancel_amount" & i, adCurrency
            
            .Append "return_number" & i, adInteger
            .Append "return_amount" & i, adCurrency
            .Append "return_charge" & i, adCurrency
            
            .Append "change_number" & i, adInteger
            .Append "change_amount" & i, adCurrency
            .Append "change_charge" & i, adCurrency
            
            .Append "handle_charge" & i, adCurrency
            .Append "total_number" & i, adInteger
            .Append "total_amount" & i, adCurrency
            
            .Append "children_number" & i, adCurrency
            .Append "children_amount" & i, adCurrency
        Next i
    End With
        rsData.Open
    For i = 1 To nUnitCount
        rsData.AddNew
        rsData!user_id = RTrim(UnitID(i)) & "[" & GetSellerNameByID(rsUnit2, UnitID(i)) & "]"
    
        If FindSeller(rsSell, UnitID(i)) Then
            Do While Not rsSell.EOF
                If RTrim(rsSell!user_id) <> RTrim(UnitID(i)) Then Exit Do

                

                If FormatDbValue(rsSell!has_child) = True And FormatDbValue(rsSell!ticket_type) = 1 Then
                    rsData!children_number1 = rsSell!num
                    rsData!children_amount1 = rsSell!pay
                    rsData("handle_charge1") = FormatDbValue(rsData("handle_charge1")) + rsSell!handle_charge
                    GoTo there
                End If

                rsData("number_ticket_type1" & rsSell!ticket_type) = rsData("number_ticket_type1" & rsSell!ticket_type) + rsSell!num
                rsData("amount_ticket_type1" & rsSell!ticket_type) = rsData("amount_ticket_type1" & rsSell!ticket_type) + rsSell!pay
                rsData("handle_charge1") = FormatDbValue(rsData("handle_charge1")) + FormatDbValue(rsSell!handle_charge)
'                rsData!cancel_number1 = rsData!cancel_number1 + rsUnit!ticket_cancel_number1
'                rsData!cancel_amount1 = rsData!cancel_amount1 + rsUnit!ticket_cancel_amount1
'
'                rsData!return_number1 = rsData!return_number1 + rsUnit!ticket_return_number1
'                rsData!return_amount1 = rsData!return_amount1 + rsUnit!ticket_return_amount1
'                rsData!return_charge1 = rsData!return_charge1 + rsUnit!ticket_return_charge1
'
'                rsData!change_number1 = rsData!change_number1 + rsUnit!ticket_change_number1
'                rsData!change_amount1 = rsData!change_amount1 + rsUnit!ticket_change_amount1
'                rsData!change_charge1 = rsData!change_charge1 + rsUnit!ticket_change_charge1
'
there:
                rsSell.MoveNext
                
            Loop


        
        End If

        If FindSeller(rsCancel, UnitID(i)) Then
            Do While Not rsCancel.EOF
                If RTrim(rsCancel!user_id) <> RTrim(UnitID(i)) Then Exit Do

                rsData("return_number1") = rsData("return_number1") + rsCancel!num
                rsData("return_amount1") = rsData("return_amount1") + rsCancel!pay
                rsData("return_charge1") = rsData("return_charge1") + rsCancel!return_handle_charge
              
                rsCancel.MoveNext
                
            Loop


        
        End If

        '求总张数及总票款
        lNumber = 0
        dbAmount = 0
        For j = 1 To TP_TicketTypeCount
            dbAmount = dbAmount + rsData("amount_ticket_type1" & j)
            lNumber = lNumber + rsData("number_ticket_type1" & j)
        Next j
        dbAmount = dbAmount + rsData("children_amount1")
        lNumber = lNumber + rsData("children_number1")
        rsData!total_number1 = lNumber - rsData("return_number1")

        rsData!total_amount1 = dbAmount - rsData("handle_charge1") + rsData("return_charge1") - rsData("return_amount1") '+ rsData!return_charge1 + rsData!change_charge1
        
'
'        '求总张数及总票款
'        lNumber = 0
'        dbAmount = 0
'        For j = 1 To TP_TicketTypeCount
'            dbAmount = dbAmount + rsData("amount_ticket_type2" & j)
'            lNumber = lNumber + rsData("number_ticket_type2" & j)
'        Next j
'
'
'        rsData!total_number2 = lNumber - rsData!return_number2 - rsData!change_number2
'        rsData!sell_number2 = lNumber
'        rsData!sell_amount2 = dbAmount
'        rsData!sell_charge2 = lNumber * rsData!sell_charge2
'        rsData!total_amount2 = dbAmount - rsData!return_amount2 - rsData!change_amount2 - rsData!sell_charge2 '+ rsData!return_charge2 + rsData!change_charge2
'
'        rsData!total_number = rsData!total_number2 - rsData!total_number1
'        rsData!total_amount = rsData!total_amount2 - rsData!total_amount1
'
        rsData.Update
        
    Next i
        

    Set InternetTkCount = rsData
End Function


Private Function FindSeller(prsData As Recordset, pszUserID As String) As Boolean
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetRouteTurnOver
    


    If prsData.RecordCount > 0 Then
        pszUserID = RTrim(pszUserID)
        prsData.MoveFirst
        Do While Not prsData.EOF
            If RTrim(prsData!user_id) = pszUserID Then
                FindSeller = True
                Exit Do
            End If
            prsData.MoveNext
        Loop
    End If
End Function
Private Function GetSellerNameByID(prsData As Recordset, pszUserID As String) As String


    If prsData.RecordCount > 0 Then
        prsData.MoveFirst
        Do While Not prsData.EOF
            If RTrim(prsData!user_id) = pszUserID Then
                GetSellerNameByID = RTrim(prsData!user_name)
            End If
            prsData.MoveNext
        Loop
    End If
End Function
'网上售票明细
Public Function InternetTkDetail(UnitID() As String, FromDate As Date, ToDate As Date, nStatus As Integer, bSellCount As Boolean, Optional nSellStationId As String = "") As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_TicketStationDim
    AssertHaveRight m_oActiveUser, RIGHT_ExUnitDateStat
    'bSellCount按售票时间来统计
    Dim oDB As New RTConnection
    Dim rs As Recordset
    Dim szSql As String
    Dim nUnitCount As Integer, i As Integer
    Dim rsUnit2 As Recordset
    Dim j As Integer
    Dim lNumber As Long
    Dim dbAmount As Double
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    nUnitCount = ArrayLength(UnitID)
    
    szSql = "Select i.print_time,cast(r.handle_charge as money)  return_handle_charge,cast(i.handle_charge as money) handle_charge,c.unit_short_name,i.getticket_id,i.seat_no,busy_date,station_name des_station_name,user_name,sell_station_name,ticket_type,i.bus_id,i.bus_date,cast(pay_count as money) pay_count,ticket_type_name,ticket_type,ticket_id,case if_print when '0' then '网购'  when '1' then '网购' when '2' then '已取消' when '3' then '已作废' when '4' then '已退' end  if_print" _
    & " from internet_sell_ticket_info i,user_info s,ticket_type_code tc,station_info si,sell_station_info ss,internet_return_ticket_info r,connect_unit_info c" _
    & " where  i.bank_id=s.user_id and tc.ticket_type_id=i.ticket_type and si.station_id=i.des_station_id " _
    & " and ss.sell_station_id=i.sell_station_id and i.id*=r.id and i.getticket_id*=r.getticket_id and i.validate_id*=r.validate_id  and s.unit_id=c.unit_id "
    If bSellCount = True Then
        szSql = szSql & " and busy_date>='" & ToDBDate(FromDate) & "' and busy_date<=dateadd(d,1,'" & ToDBDate(ToDate) & "') "
    Else
        szSql = szSql & " and print_time>='" & ToDBDate(FromDate) & "' and print_time<=dateadd(d,1,'" & ToDBDate(ToDate) & "') "
    End If
    If nUnitCount > 0 Then
        szSql = szSql & " AND (s.user_id='" & UnitID(1) & "' "
        For i = 2 To nUnitCount
            szSql = szSql & " OR s.user_id='" & UnitID(i) & "' "
        Next
        szSql = szSql & ")"
    End If
    If nStatus = 0 Then
        szSql = szSql & " AND i.if_print=0"
    ElseIf nStatus = 1 Then
        szSql = szSql & " AND i.if_print=1"
    ElseIf nStatus = 2 Then
        szSql = szSql & " AND i.if_print=2"
    ElseIf nStatus = 3 Then
        szSql = szSql & " AND i.if_print=3"
    ElseIf nStatus = 4 Then
        szSql = szSql & " AND (i.if_print=1 or i.if_print=0)"
    ElseIf nStatus = 5 Then
        szSql = szSql & " AND i.if_print=4"
    End If
    
    If nSellStationId <> "" Then
        szSql = szSql & " AND i.sell_station_id= " & TransFieldValueToString(nSellStationId)
    End If
    
    szSql = szSql & " order by  if_print,i.getticket_id,i.bus_date"
    
    Set rs = oDB.Execute(szSql)
    
    Set InternetTkDetail = rs
    


End Function

