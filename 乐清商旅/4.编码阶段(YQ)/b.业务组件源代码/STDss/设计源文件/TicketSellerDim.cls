VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TicketSellerDim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Public Enum EErrTicketSellerDim
    ERR_TicketSellerDimNoActiveUser = ERR_TicketSellerDim + ERR_NoActiveUser
End Enum

'此处定义权限号：
Public Enum ERightTicketSellerDim
    RIGHT_QuerySellerDateStat = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 1 '查询售票员售票情况
    RIGHT_QuerySellerUnitStat = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 2  '查询售票员互联售票情况
    RIGHT_QuerySellerProxySellStat = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 3  '查询车站售票（代理售票）情况
    RIGHT_DeleteCombile = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 4
    RIGHT_DeleteCombileBusCompany = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 5
    RIGHT_GetCombineSimply = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 6
    RIGHT_GetCombineSimplyByBusDate = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 7 '按照车次组合查询公司简报(按车次日期进行统计)
    RIGHT_GetCompanyFloatSimpleCon = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 8
    RIGHT_GetSplitCompanySimpleCon = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 9
    RIGHT_GetTotalItem = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 10
    RIGHT_GetUniqueCombine = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 11
    RIGHT_QueryCombile = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 12
    RIGHT_SplitCompanySimpleConByBusDate = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 13
    RIGHT_GetRouteByAreaBusTypeSimply = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 14
    RIGHT_GetRouteByBusTypeSimply = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 15
    RIGHT_GetRouteTransport = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 16
    RIGHT_GetRouteTurnOver = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 17
    RIGHT_GetOperator = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 18 ''此权限已不用
    RIGHT_SellerDateStat = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 19
    RIGHT_SellerEveryDayAnotherThing = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 20 '此权限已不用
    RIGHT_SellerEveryDaySellDetail = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 21
    RIGHT_SellerProxySellStat = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 22
    RIGHT_SellerTicketDetail = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 23
    RIGHT_SellerTicketDetailFromAgent = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 24
    RIGHT_SellerUnitStat = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 25
    RIGHT_GetStationByParam = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 26
    RIGHT_StationCount = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 27
    RIGHT_GetCombineFloatSimpleCon = ERR_TicketSellerDim + cnMidRightBegin + cnMidRightStep * 28
    

    
    
End Enum

Private m_oActiveUser As ActiveUser

'站点售票简报
Public Function StationDateStat(ByVal asSellerStation As String, FromDate As Date, ToDate As Date) As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_TicketSellerDim
    AssertHaveRight m_oActiveUser, RIGHT_SellerDateStat
    
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim rsData As New Recordset
    Dim nUserCount As Integer, i As Integer
    
    Dim j As Integer
    Dim lNumber As Long
    Dim dbAmount As Double
    Dim aswhere As String

'If InStr(asSellerStation, "+") = 0 Then
'aswhere = " and sell_station_info.sell_station_id='" & Left(asSellerStation, 2) & "'"
'    Else
'    aswhere = ""
'    End If
    
    If asSellerStation <> "" Then
        Dim szSellStationID() As String
        szSellStationID = Split(asSellerStation, ",")
        aswhere = aswhere & " AND (sell_station_info.sell_station_id='" & szSellStationID(0) & "' "
        For i = 1 To ArrayLength(szSellStationID) - 1
            aswhere = aswhere & " OR sell_station_info.sell_station_id='" & szSellStationID(i) & "' "
        Next
        aswhere = aswhere & ")"
    End If
  
  szSql = "select  rtrim(sell_station_info.sell_station_id)+'['+rtrim(sell_station_name)+']' AS user_id1,ticket_type," _
    & " SUM(ticket_sale_number) AS ticket_sale_number1," _
     & "  SUM(ticket_sale_amount) AS ticket_sale_amount1," _
 & " SUM(ticket_cancel_number) AS ticket_cancel_number1," _
 & "   SUM(ticket_cancel_amount) AS ticket_cancel_amount1," _
 & "  SUM(ticket_return_number) AS ticket_return_number1," _
  & "  SUM(ticket_return_amount) AS ticket_return_amount1," _
  & " SUM(ticket_return_charge) AS ticket_return_charge1," _
  & " SUM(ticket_change_number) AS ticket_change_number1," _
& " SUM(ticket_change_amount) AS ticket_change_amount1," _
  & " SUM(ticket_change_charge) As ticket_change_charge1" _
 & " FROM Stat_saler_ticket_lst,sell_station_info WHERE" _
  & " Stat_saler_ticket_lst.sell_station_id=sell_station_info.sell_station_id " & aswhere & " and" _
& " bus_date>='" & ToDBDate(FromDate) & "' AND bus_date<='" & ToDBDate(ToDate) & " '  group by  sell_station_info.sell_station_id,sell_station_info.sell_station_name,ticket_type ORDER BY user_id1,ticket_type"

  
  
   ' szSql = "SELECT RTRIM(sell_station_info.sell_station_id) +'['+RTRIM(sell_station_name)+']' AS user_id1 ,ticket_type," _
      ' & " SUM(ticket_sale_number) AS ticket_sale_number1, " _
      ' & " SUM(ticket_sale_amount) AS ticket_sale_amount1," _
     ' & " SUM(ticket_cancel_number) AS ticket_cancel_number1," _
      '& " SUM(ticket_cancel_amount) AS ticket_cancel_amount1," _
      '& " SUM(ticket_return_number) AS ticket_return_number1," _
      '& " SUM(ticket_return_amount) AS ticket_return_amount1," _
     '& " SUM(ticket_return_charge) AS ticket_return_charge1," _
   ' & " SUM(ticket_change_number) AS ticket_change_number1," _
    '& " SUM(ticket_change_amount) AS ticket_change_amount1," _
    '& " SUM(ticket_change_charge) AS ticket_change_charge1 " _
    '& " FROM Stat_saler_ticket_lst,sell_station_info WHERE " _
   ' & " Stat_saler_ticket_lst.sell_station_id=sell_station_info.sell_station_id" & aswhere + "AND" _
   ' & " bus_date>='" & ToDBDate(FromDate) & "' AND " _
   ' & " bus_date<='" & ToDBDate(ToDate) & "' group by  sell_station_info.sell_station_id,sell_station_name,ticket_type ORDER BY user_id1,ticket_type"

   ' nUserCount = ArrayLength(UserID)
   ' If nUserCount > 0 Then
        'szSql = szSql & " AND (Stat_saler_ticket_lst.user_id='" & UserID(1) & "' "
        'For i = 2 To nUserCount
         '   szSql = szSql & " OR Stat_saler_ticket_lst.user_id='" & UserID(i) & "' "
        'Next
        'szSql = szSql & ")"
   ' End If
    'szSql = szSql & " GROUP BY user_info.user_id,user_name,ticket_type ORDER BY user_id1,ticket_type "
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    Set rsTemp = oDB.Execute(szSql)

'做出Recordset
    '添加字段
    With rsData.Fields
        .Append "user_id", rsTemp("user_id1").Type, rsTemp("user_id1").DefinedSize
        
        '添加各种票种的数量与金额
        For j = 1 To TP_TicketTypeCount
            .Append "number_ticket_type" & j, adInteger
            .Append "amount_ticket_type" & j, adCurrency
        Next j
        
        .Append "cancel_number", adInteger
        .Append "cancel_amount", adCurrency
        
        .Append "return_number", adInteger
        .Append "return_amount", adCurrency
        .Append "return_charge", adCurrency
        
        .Append "change_number", adInteger
        .Append "change_amount", adCurrency
        .Append "change_charge", adCurrency
        
        .Append "total_number", adInteger
        .Append "total_amount", adCurrency
    End With
    rsData.Open
    Dim szLastUserID As String
    Dim szFieldPrefix As String
    
    Do While Not rsTemp.EOF
        If szLastUserID <> RTrim(rsTemp!user_id1) Or rsTemp!ticket_type = TP_FullPrice Then
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
            
            szLastUserID = RTrim(rsTemp!user_id1)
            rsData!user_id = szLastUserID
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
    
    Set StationDateStat = rsData
End Function
'售票员售票简报
Public Function SellerDateStat(UserID() As String, FromDate As Date, ToDate As Date) As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_TicketSellerDim
    AssertHaveRight m_oActiveUser, RIGHT_SellerDateStat
    
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim rsData As New Recordset
    Dim nUserCount As Integer, i As Integer
    
    Dim j As Integer
    Dim lNumber As Long
    Dim dbAmount As Double
    
    szSql = "SELECT RTRIM(user_info.user_id) +'['+RTRIM(user_name)+']' AS user_id1 ,ticket_type," _
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
    & " FROM Stat_saler_ticket_lst,user_info WHERE " _
    & " Stat_saler_ticket_lst.user_id=user_info.user_id AND " _
    & " bus_date>='" & ToDBDate(FromDate) & "' AND " _
    & " bus_date<='" & ToDBDate(ToDate) & "' "
    
    nUserCount = ArrayLength(UserID)
    If nUserCount > 0 Then
        szSql = szSql & " AND (Stat_saler_ticket_lst.user_id='" & UserID(1) & "' "
        For i = 2 To nUserCount
            szSql = szSql & " OR Stat_saler_ticket_lst.user_id='" & UserID(i) & "' "
        Next
        szSql = szSql & ")"
    End If
    szSql = szSql & " GROUP BY user_info.user_id,user_name,ticket_type ORDER BY user_id1,ticket_type "
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    Set rsTemp = oDB.Execute(szSql)

'做出Recordset
    '添加字段
    With rsData.Fields
        .Append "user_id", rsTemp("user_id1").Type, rsTemp("user_id1").DefinedSize
        
        '添加各种票种的数量与金额
        For j = 1 To TP_TicketTypeCount
            .Append "number_ticket_type" & j, adInteger
            .Append "amount_ticket_type" & j, adCurrency
        Next j
        
        .Append "cancel_number", adInteger
        .Append "cancel_amount", adCurrency
        
        .Append "return_number", adInteger
        .Append "return_amount", adCurrency
        .Append "return_charge", adCurrency
        
        .Append "change_number", adInteger
        .Append "change_amount", adCurrency
        .Append "change_charge", adCurrency
        
        .Append "total_number", adInteger
        .Append "total_amount", adCurrency
    End With
    rsData.Open
    Dim szLastUserID As String
    Dim szFieldPrefix As String
    
    Do While Not rsTemp.EOF
        If szLastUserID <> RTrim(rsTemp!user_id1) Or rsTemp!ticket_type = TP_FullPrice Then
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
            
            szLastUserID = RTrim(rsTemp!user_id1)
            rsData!user_id = szLastUserID
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
    
    Set SellerDateStat = rsData
End Function



'售票员售票简报(有起始票号)
Public Function SellerDateStatHaveTicket(UserID() As String, FromDate As Date, ToDate As Date) As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_TicketSellerDim
    AssertHaveRight m_oActiveUser, RIGHT_SellerDateStat
    
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim rsData As New Recordset
    Dim nUserCount As Integer, i As Integer
    Dim nCredenceCount As Integer '计算凭证张数
    
    Dim j As Integer
    Dim lNumber As Long
    Dim dbAmount As Double
    
    
    Dim szWhere As String
    
    szSql = "SELECT RTRIM(user_info.user_id) +'['+RTRIM(user_name)+']' AS user_id1 ,ticket_type," _
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
    & " FROM Stat_saler_ticket_lst s ,user_info WHERE " _
    & " s.user_id=user_info.user_id AND " _
    & " bus_date>='" & ToDBDate(FromDate) & "' AND " _
    & " bus_date<='" & ToDBDate(ToDate) & "' "
    
    nUserCount = ArrayLength(UserID)
    If nUserCount > 0 Then
        szWhere = szWhere & " AND (s.user_id='" & UserID(1) & "' "
        For i = 2 To nUserCount
            szWhere = szWhere & " OR s.user_id='" & UserID(i) & "' "
        Next
        szWhere = szWhere & ")"
    End If
    szSql = szSql & szWhere & " GROUP BY user_info.user_id,user_name,ticket_type ORDER BY user_id1,ticket_type "
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    Set rsTemp = oDB.Execute(szSql)

'做出Recordset
    '添加字段
    With rsData.Fields
        .Append "user_id", rsTemp("user_id1").Type, rsTemp("user_id1").DefinedSize
        .Append "start_ticket", adChar, 10
        .Append "end_ticket", adChar, 10
        
        '添加各种票种的数量与金额
        For j = 1 To TP_TicketTypeCount
            .Append "number_ticket_type" & j, adInteger
            .Append "amount_ticket_type" & j, adCurrency
        Next j
        
        .Append "cancel_number", adInteger
        .Append "cancel_amount", adCurrency
        
        .Append "return_number", adInteger
        .Append "return_amount", adCurrency
        .Append "return_charge", adCurrency
        .Append "CredenceCount", adInteger '凭证张数
        
        
        .Append "change_number", adInteger
        .Append "change_amount", adCurrency
        .Append "change_charge", adCurrency
        
        .Append "total_number", adInteger
        .Append "total_amount", adCurrency
        
        .Append "Act_total_number", adInteger '总用票张数
        
    End With
    rsData.Open
    Dim szLastUserID As String
    Dim szFieldPrefix As String
    Dim rsTicket As Recordset
    Dim szLastTicket As String
    Dim rsReturnTicket As Recordset
    
    szSql = " SELECT * FROM agent_ticket_sale_lst s WHERE operation_time>='" & ToDBDate(FromDate) & "' AND " _
        & " operation_time<='" & ToDBDate(ToDate) & " 23:59:59" & "' " _
        & szWhere _
        & " ORDER BY user_id , ticket_id"
    
   Set rsTicket = oDB.Execute(szSql)
    
    
     szSql = " SELECT * FROM agent_ticket_return_lst s WHERE operation_time>='" & ToDBDate(FromDate) & "' AND " _
        & " operation_time<='" & ToDBDate(ToDate) & " 23:59:59" & "' " _
        & szWhere _
        & " ORDER BY user_id , operation_time"
 
    Set rsReturnTicket = oDB.Execute(szSql)
 
 
    
    
    Do While Not rsTemp.EOF
        If szLastUserID <> RTrim(rsTemp!user_id1) Or rsTemp!ticket_type = TP_FullPrice Then
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
                rsData!Act_total_number = lNumber
                '查找各票号
                If rsTicket.RecordCount > 0 Then
                    rsTicket.MoveFirst
                    rsTicket.Find " user_id = '" & ResolveDisplay(szLastUserID) & "'"
                    
                    If Not rsTicket.EOF Then
                        rsData!start_ticket = rsTicket!ticket_id
                        szLastTicket = rsTicket!ticket_id
                        Do While Not rsTicket.EOF
                            If Trim(UCase(rsTicket!user_id)) = Trim(UCase(ResolveDisplay(szLastUserID))) Then
                                szLastTicket = rsTicket!ticket_id
                                rsTicket.MoveNext
                            Else
                                Exit Do
                            End If
                        Loop
                        rsData!end_ticket = szLastTicket
                    End If
                End If
                
                
                Dim szRFticketID As String
                Dim szRLticketID As String
'                 '查找各票号
                If rsReturnTicket.RecordCount > 0 Then
                    rsReturnTicket.MoveFirst
                    rsReturnTicket.Find " user_id = '" & ResolveDisplay(szLastUserID) & "'"

                    If Not rsReturnTicket.EOF Then
     
                        szRLticketID = rsReturnTicket!credence_sheet_id
                        Do While Not rsReturnTicket.EOF
                            If Trim(UCase(rsReturnTicket!user_id)) = Trim(UCase(ResolveDisplay(szLastUserID))) Then
                                If Trim(rsReturnTicket!credence_sheet_id) <> "" Then
                                    If szRFticketID > rsReturnTicket!credence_sheet_id Or szRFticketID = "" Then
                                        szRFticketID = rsReturnTicket!credence_sheet_id
                                    End If
                                End If
                                If szRLticketID < FormatDbValue(rsReturnTicket!credence_sheet_id) Then
                                    szRLticketID = FormatDbValue(rsReturnTicket!credence_sheet_id)
                                End If
                                '如果凭证张数不为空,则认为是全额退票,不计入凭证张数中.
                                If FormatDbValue(rsReturnTicket!credence_sheet_id) <> "" Then
                                    nCredenceCount = nCredenceCount + 1
                                End If
                                
                                rsReturnTicket.MoveNext
                            Else
                                Exit Do
                            End If
                        Loop
'                        rsData!end_ticket = szRLticketID
                    End If
                End If

               '如果凭证号比起始票号要小的话，起始票号则为凭证号
              If szRFticketID <> "" Then
                    If rsData!start_ticket > szRFticketID Then
                         rsData!start_ticket = szRFticketID
                    End If
              End If

               '如果凭证号比结束票号要大的话，结束票号则为凭证号
              If szRLticketID <> "" Then
                    If rsData!end_ticket < szRLticketID Then
                         rsData!end_ticket = szRLticketID
                    End If
             End If
             rsData!CredenceCount = nCredenceCount
                rsData.Update
            End If
            rsData.AddNew
            
            szLastUserID = RTrim(rsTemp!user_id1)
            rsData!user_id = szLastUserID
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
'        rsData!Act_total_number = rsData("number_ticket_type" & rsTemp!ticket_type) + rsData!cancel_number + rsData!return_number + rsData!change_number
        
        rsTemp.MoveNext
        '清空凭证张数
        nCredenceCount = 0
        szRFticketID = ""
        szRLticketID = ""
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

'        rsTemp.MovePrevious
        '查找各票号
        If rsTicket.RecordCount > 0 Then
            rsTicket.MoveFirst
            rsTicket.Find " user_id = '" & ResolveDisplay(szLastUserID) & "'"
            
            If Not rsTicket.EOF Then
                rsData!start_ticket = rsTicket!ticket_id
                szLastTicket = rsTicket!ticket_id
                Do While Not rsTicket.EOF
                    If Trim(UCase(rsTicket!user_id)) = Trim(UCase(ResolveDisplay(szLastUserID))) Then
                        szLastTicket = rsTicket!ticket_id
                        rsTicket.MoveNext
                    Else
                        Exit Do
                    End If
                Loop
                rsData!end_ticket = szLastTicket
            End If
        End If
        
        
'                Dim szRFticketID As String
'                Dim szRLticketID As String
                 '查找各票号
                If rsReturnTicket.RecordCount > 0 Then
                    rsReturnTicket.MoveFirst
                    rsReturnTicket.Find " user_id = '" & ResolveDisplay(szLastUserID) & "'"
                    
                    If Not rsReturnTicket.EOF Then
                        If Trim(rsReturnTicket!credence_sheet_id) <> "" Then
                                szRFticketID = rsReturnTicket!credence_sheet_id
                        End If
                        szRLticketID = rsReturnTicket!credence_sheet_id
                        Do While Not rsReturnTicket.EOF
                        
                        
                            If Trim(UCase(rsReturnTicket!user_id)) = Trim(UCase(ResolveDisplay(szLastUserID))) Then
                                If Trim(rsReturnTicket!credence_sheet_id) <> "" Then
                                    If szRFticketID > rsReturnTicket!credence_sheet_id Or szRFticketID = "" Then
                                        szRFticketID = rsReturnTicket!credence_sheet_id
                                    End If
                                End If
                                If szRLticketID < FormatDbValue(rsReturnTicket!credence_sheet_id) Then
                                    szRLticketID = FormatDbValue(rsReturnTicket!credence_sheet_id)
                                End If
                                '如果凭证张数不为空,则认为是全额退票,不计入凭证张数中.
                                If FormatDbValue(rsReturnTicket!credence_sheet_id) <> "" Then
                                    nCredenceCount = nCredenceCount + 1
                                End If
                                
                                rsReturnTicket.MoveNext
                            Else
                                Exit Do
                            End If

'
'
'
'
'
'                            If Trim(UCase(rsReturnTicket!user_id)) = Trim(UCase(ResolveDisplay(szLastUserID))) Then
'                                If szRLticketID < FormatDbValue(rsReturnTicket!credence_sheet_id) Then
'                                    szRLticketID = FormatDbValue(rsReturnTicket!credence_sheet_id)
'                                End If
'                                '如果凭证张数不为空,则认为是全额退票,不计入凭证张数中.
'                                If Trim(szRLticketID) <> "" Then
'                                    nCredenceCount = nCredenceCount + 1
'                                End If
'                                rsReturnTicket.MoveNext
'                            Else
'                                Exit Do
'                            End If
                        Loop
'                        rsData!end_ticket = szRLticketID
                    End If
                End If
                
               '如果凭证号比起始票号要小的话，起始票号则为凭证号
               If szRFticketID <> "" Then
                    If rsData!start_ticket > szRFticketID Then
                         rsData!start_ticket = szRFticketID
                    End If
               End If
                    
               '如果凭证号比结束票号要大的话，结束票号则为凭证号
               If szRLticketID <> "" Then
                    If rsData!end_ticket < szRLticketID Then
                         rsData!end_ticket = szRLticketID
                    End If
               End If
        
        rsData!CredenceCount = nCredenceCount
        rsData!Act_total_number = lNumber
        rsData.Update
        
    End If
    
    Set SellerDateStatHaveTicket = rsData
End Function


'售票员互连售票简报
Public Function SellerUnitStat(UserID As String, UnitID() As String, FromDate As Date, ToDate As Date) As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_TicketSellerDim
    AssertHaveRight m_oActiveUser, RIGHT_SellerUnitStat
    
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim rsData As New Recordset
    Dim nUnitCount As Integer, i As Integer
    Dim j As Integer
    Dim lNumber As Long
    Dim dbAmount As Double
    
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
    & " FROM Stat_saler_ticket_lst,connect_unit_info WHERE " _
    & " Stat_saler_ticket_lst.unit_id=connect_unit_info.unit_id AND " _
    & " user_id='" & UserID & "' AND " _
    & " bus_date>='" & ToDBDate(FromDate) & "' AND " _
    & " bus_date<='" & ToDBDate(ToDate) & "' "
    
    nUnitCount = ArrayLength(UnitID)
    If nUnitCount > 0 Then
        szSql = szSql & " AND (Stat_saler_ticket_lst.unit_id='" & UnitID(1) & "' "
        For i = 2 To nUnitCount
            szSql = szSql & " OR Stat_saler_ticket_lst.unit_id='" & UnitID(i) & "' "
        Next
        szSql = szSql & ")"
    End If
    szSql = szSql & " GROUP BY connect_unit_info.unit_id,unit_short_name,ticket_type ORDER BY unit_id1,ticket_type "
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    Set rsTemp = oDB.Execute(szSql)
    
    
'做出Recordset
    '添加字段
    With rsData.Fields
        .Append "unit_id", rsTemp("unit_id1").Type, rsTemp("unit_id1").DefinedSize
        
        '添加各种票种的数量与金额
        For j = 1 To TP_TicketTypeCount
            .Append "number_ticket_type" & j, adInteger
            .Append "amount_ticket_type" & j, adCurrency
        Next j
        
        .Append "cancel_number", adInteger
        .Append "cancel_amount", adCurrency
        
        .Append "return_number", adInteger
        .Append "return_amount", adCurrency
        .Append "return_charge", adCurrency
        
        .Append "change_number", adInteger
        .Append "change_amount", adCurrency
        .Append "change_charge", adCurrency
        
        .Append "total_number", adInteger
        .Append "total_amount", adCurrency
    End With
    rsData.Open
    Dim szLastUnitID As String
    Dim szFieldPrefix As String
    
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
    
    Set SellerUnitStat = rsData
End Function

'车站售票（代理售票）简报
Public Function SellerProxySellStat(FromDate As Date, ToDate As Date, Optional pszSellStationID As String = "") As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_TicketSellerDim
    AssertHaveRight m_oActiveUser, RIGHT_SellerProxySellStat
    
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim rsData As New Recordset
    Dim j As Integer
    Dim lNumber As Long
    Dim dbAmount As Double
    
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
    & " FROM Stat_saler_ticket_lst WHERE " _
    & " bus_date>='" & ToDBDate(FromDate) & "' AND " _
    & " bus_date<='" & ToDBDate(ToDate) & "' "
    
    If pszSellStationID <> "" Then
        Dim szSellStationID() As String
        Dim i As Integer
        szSellStationID = Split(pszSellStationID, ",")
        szSql = szSql & " AND (sell_station_id='" & szSellStationID(0) & "' "
        For i = 1 To ArrayLength(szSellStationID) - 1
            szSql = szSql & " OR sell_station_id='" & szSellStationID(i) & "' "
        Next
        szSql = szSql & ")"
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
    Dim szFieldPrefix As String
    
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
        
    Set SellerProxySellStat = rsData
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

'售票员环境时间售票详情
Public Function SellerEveryDaySellDetail(UserID As String, StartDate As Date, EndDate As Date) As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_QuerySellerEveryDaySellDetail
    
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszDss)
    
    szSql = "SELECT * FROM Agent_ticket_sale_lst WHERE " _
    & " user_id='" & UserID & "' AND " _
    & " operation_time>='" & ToDBDateTime(StartDate) & "' AND " _
    & " operation_time<'" & ToDBDateTime(EndDate) & "'" _
    & " ORDER BY ticket_id"
      
    
    Set rsTemp = oDB.Execute(szSql)
    Set SellerEveryDaySellDetail = rsTemp
End Function

'数组相应为废、退、改
Public Function SellerEveryDayAnotherThing(UserID As String, StartDate As Date, EndDate As Date) As Double()
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_QuerySellerEveryDaySellDetail
    
    Dim adbResult(1 To 4, 1 To 3) As Double
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    szSql = "SELECT COUNT(*) AS countx ,SUM(ticket_price) AS total_ticket_price FROM " _
    & " Agent_ticket_sale_lst,Agent_ticket_cancel_lst WHERE " _
    & " Agent_ticket_sale_lst.ticket_id=Agent_ticket_cancel_lst.ticket_id AND " _
    & " Agent_ticket_sale_lst.user_id ='" & UserID & "' AND " _
    & " Agent_ticket_cancel_lst.cancel_time>='" & ToDBDateTime(StartDate) & "' AND " _
    & " Agent_ticket_cancel_lst.cancel_time<'" & ToDBDateTime(EndDate) & "'"
    
    
    Set rsTemp = oDB.Execute(szSql)
    adbResult(1, 1) = rsTemp!countx
    adbResult(1, 2) = FormatDbValue(rsTemp!total_ticket_price)
    '退票统计
    '有手续费的退票
    szSql = "SELECT COUNT(*) AS countx ,SUM(ticket_price) AS total_ticket_price, " _
    & " SUM(return_charge) AS total_return_chare FROM " _
    & " Agent_ticket_return_lst WHERE " _
    & " user_id='" & UserID & "' AND " _
    & " operation_time>='" & ToDBDateTime(StartDate) & "' AND " _
    & " operation_time<'" & ToDBDateTime(EndDate) & "'" _
    & " AND return_charge > 0 "
    
    Set rsTemp = oDB.Execute(szSql)
    adbResult(2, 1) = rsTemp!countx
    adbResult(2, 2) = FormatDbValue(rsTemp!total_ticket_price)
    adbResult(2, 3) = FormatDbValue(rsTemp!total_return_chare)
    
    
    '以下是改签票处理
    szSql = "SELECT COUNT(*) AS countx ,SUM(fORmer_ticket_price) AS total_ticket_price, " _
    & " SUM(change_charge) AS total_change_chare FROM " _
    & " Agent_ticket_change_lst WHERE " _
    & " user_id='" & UserID & "' AND " _
    & " change_time>='" & ToDBDateTime(StartDate) & "' AND " _
    & " change_time<'" & ToDBDateTime(EndDate) & "'"
    
    Set rsTemp = oDB.Execute(szSql)
    adbResult(3, 1) = rsTemp!countx
    adbResult(3, 2) = FormatDbValue(rsTemp!total_ticket_price)
    adbResult(3, 3) = FormatDbValue(rsTemp!total_change_chare)
    '改签票处理到这里完
    
    '退票统计
    '无手续费的退票,即全额退票
    szSql = "SELECT COUNT(*) AS countx ,SUM(ticket_price) AS total_ticket_price, " _
    & " SUM(return_charge) AS total_return_chare FROM " _
    & " Agent_ticket_return_lst WHERE " _
    & " user_id='" & UserID & "' AND " _
    & " operation_time>='" & ToDBDateTime(StartDate) & "' AND " _
    & " operation_time<'" & ToDBDateTime(EndDate) & "'" _
    & " AND return_charge = 0 "
    
    Set rsTemp = oDB.Execute(szSql)
    adbResult(4, 1) = rsTemp!countx
    adbResult(4, 2) = FormatDbValue(rsTemp!total_ticket_price)
    adbResult(4, 3) = FormatDbValue(rsTemp!total_return_chare)
    
    SellerEveryDayAnotherThing = adbResult
End Function

Public Function SellerTicketDetail(paszUserID As Variant, pdyStart As Date, pdyEnd As Date, Optional szStartTicketNo As String, Optional szEndTicketNo As String, Optional pnStatus As Integer = 0, Optional szBusid As String, Optional szIDCardNo As String, Optional szPersonName As String) As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_SellerTicketDetail
    '先不判断权限
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    
    Dim szTemp As String
    Dim nCount As Integer
    Dim i As Integer
    Dim szTempTime As String
    Dim szTempTicket As String
    Dim szTempBusId As String
    Dim szOrder As String '排序串
    Dim szNullSql As String
    Dim szTempRealName As String
    
    oDB.ConnectionString = GetConnectionStr(cszDss)

    If pnStatus = ST_TicketReturned Then
        nCount = ArrayLength(paszUserID)
        If nCount > 0 Then
            szTemp = " AND rst.user_id IN ('"
            For i = 1 To nCount
                szTemp = szTemp & ResolveDisplay(paszUserID(i)) & "','"
            Next i
            szTemp = Left(szTemp, Len(szTemp) - 2) & ")"
        End If
        If ToDBDateTime(pdyStart) = ToDBDateTime("0:00:00") Or ToDBDateTime(pdyEnd) = ToDBDateTime("0:00:00") Then
            szTempTime = ""
        Else
            szTempTime = " AND rst.operation_time >=" & "'" & ToDBDateTime(pdyStart) & "'" & " AND rst.operation_time <" & " '" & ToDBDateTime(pdyEnd) & "' "
        End If
        If szStartTicketNo <> "" Then
            szTempTicket = " AND sst.ticket_id>='" & szStartTicketNo & "'"
        Else
            szStartTicketNo = ""
        End If
        If szEndTicketNo <> "" Then
            szTempTicket = szTempTicket & " AND sst.ticket_id<='" & szEndTicketNo & "'"
        Else
            szStartTicketNo = szTempTicket & ""
        End If

        If szBusid <> "" Then
            szTempBusId = " AND sst.bus_id='" & szBusid & "'"
        Else
            szTempBusId = ""
        End If
        
        If szIDCardNo <> "" Then
           szTempRealName = szTempRealName & " AND sst.id_card_no='" & szIDCardNo & "'"
        End If
        If szPersonName <> "" Then
           szTempRealName = szTempRealName & " AND sst.person_name='" & szPersonName & "'"
        End If
        
        szSql = "SELECT sst.ticket_id,tfo.ticket_type_name,sst.bus_id,sst.bus_date,sst.ticket_type,sst.user_id as sale_user_id,sst.seat_no," & _
                " sst.ticket_price, sst.operation_time ,rst.operation_time as return_ticket_time,sst.card_type ,sst.id_card_no ,sst.person_name ," & _
                " rst.credence_sheet_id , rst.return_charge, rst.user_id , sfo.station_name des_station_name,sfo.station_name station_name,ssi.sell_station_name " & _
                " FROM Ticket_return_lst rst,Ticket_sell_lst sst,Ticket_type_code tfo , station_info sfo ,sell_station_info ssi" & _
                " WHERE rst.ticket_id=sst.ticket_id AND sfo.station_id=sst.des_station_id  AND tfo.ticket_type_id=sst.ticket_type AND sst.sell_station_id=ssi.sell_station_id"
        szOrder = " ORDER BY rst.user_id  , sst.ticket_id,rst.operation_time "
    ElseIf pnStatus = ST_TicketSellChange Then
        nCount = ArrayLength(paszUserID)
        If nCount > 0 Then
            szTemp = " AND ftk.change_user_id IN ('"
            For i = 1 To nCount
                szTemp = szTemp & ResolveDisplay(paszUserID(i)) & "','"
            Next i
            szTemp = Left(szTemp, Len(szTemp) - 2) & ")"
        End If
        If ToDBDateTime(pdyStart) = ToDBDateTime("0:00:00") Or ToDBDateTime(pdyEnd) = ToDBDateTime("0:00:00") Then
            szTempTime = ""
        Else
            szTempTime = " AND ftk.change_time >=" & "'" & ToDBDateTime(pdyStart) & "'" & " AND ftk.change_time <" & " '" & ToDBDateTime(pdyEnd) & "' "
        End If
        If szStartTicketNo <> "" Then
            szTempTicket = " AND ftk.new_ticket_id>='" & szStartTicketNo & "'"
        Else
            szStartTicketNo = ""
        End If
        If szEndTicketNo <> "" Then
            szTempTicket = szTempTicket & " AND ftk.new_ticket_id<='" & szEndTicketNo & "'"
        Else
            szStartTicketNo = szTempTicket & ""
        End If
        
        If szBusid <> "" Then
            szTempBusId = " AND (ctk.new_bus_id='" & szBusid & "' OR ftk.bus_id='" & szBusid & "')"
        Else
            szTempBusId = ""
        End If
        
        If szIDCardNo <> "" Then
           szTempRealName = szTempRealName & " AND ftk.id_card_no='" & szIDCardNo & "'"
        End If
        If szPersonName <> "" Then
           szTempRealName = szTempRealName & " AND ftk.person_name='" & szPersonName & "'"
        End If
        
        szSql = "SELECT ftk.*, ctk.new_bus_id,ctk.new_bus_date,ctk.new_seat_no,ctk.new_ticket_price ticket_price,ctk.new_station_name station_name ,tfo.ticket_type_name,ssi.sell_station_name FROM " & _
              " (SELECT cst.new_ticket_id,cst.change_time,cst.change_charge,cst.fORmer_ticket_id," & _
              " cst.fORmer_ticket_price,cst.user_id as change_user_id,sst.bus_date,sst.bus_id,sst.card_type ,sst.id_card_no ,sst.person_name , " & _
              " sst.des_station_id,sst.ticket_type,sst.user_id as sale_user_id, " & _
              " sst.seat_no as fORmer_seat_no,sfo.station_name as fORmer_station_name, " & _
              " sst.operation_time as sale_ticket_time  , sst.user_id as user_id ,sst.sell_station_id" & _
              " FROM ticket_change_lst cst,Ticket_sell_lst sst,station_info sfo " & _
              " WHERE cst.fORmer_ticket_id=sst.ticket_id AND sfo.station_id=sst.des_station_id ) ftk, " & _
              " (SELECT cst.new_ticket_id,sst.seat_no  as new_seat_no,sst.bus_id as new_bus_id, " & _
              " sst.bus_date as new_bus_date,sst.des_station_id as new_des_station_id, " & _
              " sfo.station_name as new_station_name,sst.ticket_price as new_ticket_price" & _
              " FROM ticket_change_lst cst,Ticket_sell_lst sst,station_info sfo " & _
              " WHERE cst.new_ticket_id=sst.ticket_id AND sfo.station_id=sst.des_station_id) ctk ," & _
              " Ticket_type_code tfo ,sell_station_info ssi " & _
              " WHERE ftk.new_ticket_id = ctk.new_ticket_id AND tfo.ticket_type_id = ftk.ticket_type  AND ftk.sell_station_id=ssi.sell_station_id"
              
        szOrder = " ORDER BY ftk.user_id "
    ElseIf pnStatus = ST_TicketChecked Then
        szSql = "SELECT cst.ticket_id,cst.bus_date,cst.bus_id,cst.bus_serial_no,cst.check_time,cst.ticket_type," & _
                " cst.ticket_price,cst.end_station_id ,sst.ticket_type,sst.user_id,sst.seat_no,sst.ticket_price,sst.card_type ,sst.id_card_no ,sst.person_name ," & _
                " tfo.ticket_type_name , sfo.station_name,sfo.station_name as des_station_name,ssi.sell_station_name   , s.check_sheet_id " & _
                " FROM check_ticket_lst cst,Ticket_sell_lst sst,Ticket_type_code tfo,station_info sfo  ,sell_station_info ssi , sheet_lst s " & _
                " WHERE cst.ticket_id = sst.ticket_id AND tfo.ticket_type_id = sst.ticket_type AND sfo.station_id = cst.end_station_id  " _
                & " AND sst.sell_station_id=ssi.sell_station_id " _
                & " and s.bus_date = cst.bus_date and s.bus_id = cst.bus_id and s.bus_serial_no = cst.bus_serial_no and s.check_gate_id = cst.check_gate_id "
                
        If ToDBDateTime(pdyStart) = ToDBDateTime("0:00:00") Or ToDBDateTime(pdyEnd) = ToDBDateTime("0:00:00") Then
            szTempTime = ""
        Else
            szTempTime = " AND cst.check_time >=" & "'" & ToDBDateTime(pdyStart) & "'" & " AND cst.check_time <" & " '" & ToDBDateTime(pdyEnd) & "' "
        End If
        
        If szBusid <> "" Then
            szTempBusId = " AND cst.bus_id='" & szBusid & "'"
        Else
            szTempBusId = ""
        End If
        
        If szStartTicketNo <> "" Then
            szTempTicket = " AND cst.ticket_id>='" & szStartTicketNo & "'"
        Else
            szStartTicketNo = ""
        End If
        If szEndTicketNo <> "" Then
            szTempTicket = szTempTicket & " AND cst.ticket_id<='" & szEndTicketNo & "'"
        Else
            szStartTicketNo = szTempTicket & ""
        End If
        If szIDCardNo <> "" Then
           szTempRealName = szTempRealName & " AND sst.id_card_no='" & szIDCardNo & "'"
        End If
        If szPersonName <> "" Then
           szTempRealName = szTempRealName & " AND sst.person_name='" & szPersonName & "'"
        End If

    Else
        nCount = ArrayLength(paszUserID)
        If nCount > 0 Then
            szTemp = " AND user_id IN ('"
            For i = 1 To nCount
                szTemp = szTemp & ResolveDisplay(paszUserID(i)) & "','"
            Next i
            szTemp = Left(szTemp, Len(szTemp) - 2) & ")"
        End If
        If pnStatus = 0 Then
            szSql = "   SELECT s.*, station_name des_station_name ,t.ticket_type_name,ssi.sell_station_name " _
                & " , CASE s.insurance WHEN 0 THEN '无' ELSE '有' END insurance_name " _
                & " , CASE s.status WHEN 1 THEN '正常售出' WHEN 2 THEN '改签售出' WHEN 5 THEN '废票' WHEN 9 THEN '被改签' WHEN 17 THEN '退票' WHEN 33 THEN '已检' ELSE convert(char(2),s.status) END status_name " _
                & " FROM Ticket_sell_lst  s, station_info i,Ticket_type_code t ,sell_station_info ssi  " _
                & " WHERE s.des_station_id = i.station_id AND s.ticket_type=t.ticket_type_id  AND s.sell_station_id=ssi.sell_station_id"
                
            If szIDCardNo <> "" Then
               szTempRealName = szTempRealName & " AND s.id_card_no='" & szIDCardNo & "'"
            End If
            If szPersonName <> "" Then
               szTempRealName = szTempRealName & " AND s.person_name='" & szPersonName & "'"
            End If
                
            szOrder = " ORDER BY s.ticket_id "
        ElseIf pnStatus = ST_TicketNormal Then
            szSql = "   SELECT s.*, station_name des_station_name ,t.ticket_type_name ,ssi.sell_station_name " _
                & " , CASE s.insurance WHEN 0 THEN '无' ELSE '有' END insurance_name " _
                & " , CASE s.status WHEN 1 THEN '正常售出' WHEN 2 THEN '改签售出' WHEN 5 THEN '废票' WHEN 9 THEN '被改签' WHEN 17 THEN '退票' WHEN 33 THEN '已检' ELSE convert(char(2),s.status) END status_name " _
                & " FROM Ticket_sell_lst  s, station_info i,Ticket_type_code t  ,sell_station_info ssi" _
                & " WHERE s.des_station_id = i.station_id AND s.status='1' AND s.ticket_type=t.ticket_type_id  AND s.sell_station_id=ssi.sell_station_id AND bus_date >='" & ToDBDate(pdyStart) & "' AND bus_date <= '" & ToDBDate(pdyEnd) & "'"
            
            If szIDCardNo <> "" Then
               szTempRealName = szTempRealName & " AND s.id_card_no='" & szIDCardNo & "'"
            End If
            If szPersonName <> "" Then
               szTempRealName = szTempRealName & " AND s.person_name='" & szPersonName & "'"
            End If
            
            szOrder = " ORDER BY s.ticket_id "
        ElseIf pnStatus = ST_TicketCanceled Then
            If nCount > 0 Then
                szTemp = " AND s.user_id IN ('"
                For i = 1 To nCount
                    szTemp = szTemp & ResolveDisplay(paszUserID(i)) & "','"
                Next i
                szTemp = Left(szTemp, Len(szTemp) - 2) & ")"
            End If
        '            szsql = "SELECT zst.*,sst.bus_date,sst.bus_id,sst.ticket_type,sst.ticket_price,sst.seat_no,sfo.station_name,tfo.ticket_type_name " & _
        '                   " FROM cancel_ticket_lst zst,Ticket_sell_lst sst,station_info sfo,Ticket_type_code tfo " & _
        '                   " WHERE zst.ticket_id = sst.ticket_id AND sst.des_station_id = sfo.station_id AND sst.ticket_type = tfo.ticket_type_id "
        '
            szSql = "SELECT s.*,sst.bus_date,sst.bus_id,sst.ticket_type,sst.ticket_price,sst.seat_no,sfo.station_name des_station_name ," _
                & "     tfo.ticket_type_name , sst.user_id AS sale_user_id ,sst.operation_time ,ssi.sell_station_name ,sst.card_type ,sst.id_card_no ,sst.person_name " _
                & " , CASE sst.insurance WHEN 0 THEN '无' ELSE '有' END insurance_name " _
                & " , CASE sst.status WHEN 1 THEN '正常售出' WHEN 2 THEN '改签售出' WHEN 5 THEN '废票' WHEN 9 THEN '被改签' WHEN 17 THEN '退票' WHEN 33 THEN '已检' ELSE convert(char(2),sst.status) END status_name " _
                & " FROM ticket_cancel_lst s,Ticket_sell_lst sst,station_info sfo,Ticket_type_code tfo ,sell_station_info ssi " _
                & " WHERE s.ticket_id = sst.ticket_id AND sst.des_station_id = sfo.station_id AND sst.ticket_type = tfo.ticket_type_id  AND sst.sell_station_id=ssi.sell_station_id"
            
            If szBusid <> "" Then
                szTempBusId = " AND sst.bus_id='" & szBusid & "'"
            Else
                szTempBusId = ""
            End If
            If szIDCardNo <> "" Then
               szTempRealName = szTempRealName & " AND sst.id_card_no='" & szIDCardNo & "'"
            End If
            If szPersonName <> "" Then
               szTempRealName = szTempRealName & " AND sst.person_name='" & szPersonName & "'"
            End If
            szOrder = " ORDER BY s.ticket_id " 's.user_id,sst.bus_date,sst.bus_id ,sst.operation_time "
        End If
    
        If ToDBDateTime(pdyStart) = ToDBDateTime("0:00:00") Or ToDBDateTime(pdyEnd) = ToDBDateTime("0:00:00") Then
            szTempTime = ""
        ElseIf pnStatus = ST_TicketCanceled Then
            szTempTime = " AND cancel_ticket_time >=" & "'" & ToDBDateTime(pdyStart) & "'" & " AND cancel_ticket_time <" & " '" & ToDBDateTime(pdyEnd) & "' "
        ElseIf pnStatus = ST_TicketNormal Then
            szTempTime = " AND bus_date >=" & "'" & ToDBDate(pdyStart) & "'" & " AND bus_date <=" & " '" & ToDBDate(pdyEnd) & "' "
        Else
            szTempTime = " AND operation_time >=" & "'" & ToDBDateTime(pdyStart) & "'" & " AND operation_time <" & " '" & ToDBDateTime(pdyEnd) & "' "
        End If
        If szStartTicketNo <> "" Then
            szTempTicket = " AND s.ticket_id>='" & szStartTicketNo & "'"
        Else
            szStartTicketNo = ""
        End If
        If szEndTicketNo <> "" Then
            szTempTicket = szTempTicket & " AND s.ticket_id<='" & szEndTicketNo & "'"
        Else
            szStartTicketNo = szTempTicket & ""
        End If
              
        If pnStatus <> ST_TicketCanceled Then
            If szBusid <> "" Then
                szTempBusId = " AND bus_id='" & szBusid & "'"
            Else
                szTempBusId = ""
            End If
        End If
    End If
    
    '查询空白票
    If pnStatus = 0 Then
        szNullSql = "SELECT RTRIM(firstnullticket_no) + ' - ' + RTRIM(lastnullticket_no) ticket_id ,'" & cszEmptyDateStr & "' bus_date ,'空白票' bus_id ," _
            & "'空白票' sell_station_id ,'空白票' des_station_id ," _
            & "'0' status ,'0' sale_ticket_type ,'0' ticket_type ,'' seat_type_id ,user_id ,'空白票' seat_no ,'0' ticket_price ,getnullticket_date operation_time ,'0' insurance , '' card_type ,'' id_card_no ,'' person_name ,'' sex ,'' person_picture ,'空白票' des_station_name ," _
            & "'空白票' ticket_type_name ,'空白票' sell_station_name ,'空白票' insurance_name ," _
            & "'未使用' status_name FROM getnullticket"
        szNullSql = szNullSql & " WHERE getnullticket_date >=" & "'" & ToDBDateTime(pdyStart) & "'" & " AND getnullticket_date <" & " '" & ToDBDateTime(pdyEnd) & "'"
        szNullSql = szNullSql & szTemp
    End If
    
    If pnStatus = 0 Then
        szSql = szSql & szTempTime & szTemp & szTempTicket & szTempBusId & szTempRealName & " UNION " & szNullSql & szOrder
    Else
        szSql = szSql & szTempTime & szTemp & szTempTicket & szTempBusId & szTempRealName & szOrder
    End If
    
    Set rsTemp = oDB.Execute(szSql)
    Set SellerTicketDetail = rsTemp
End Function




Public Function SellerTicketDetailFromAgent(paszUserID As Variant, pdyStart As Date, pdyEnd As Date, Optional szStartTicketNo As String, Optional szEndTicketNo As String, Optional pnStatus As Integer = 0, Optional szBusid As String, Optional szIDCardNo As String, Optional szPersonName As String) As Recordset
'从代理表中查询出明细信息


    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_SellerTicketDetailFromAgent
    '先不判断权限
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    
    Dim szTemp As String
    Dim nCount As Integer
    Dim i As Integer
    Dim szTempTime As String
    Dim szTempTicket As String
    Dim szTempTicketNull As String
    Dim szTempTicketReturn As String
    Dim szTempBusId As String
    Dim szOrder As String '排序串
    Dim szNullSql As String
    Dim szNullTemp As String
    Dim szReturnVouchers As String '退票凭证sql
    Dim szTempRealName As String
    
    oDB.ConnectionString = GetConnectionStr(cszDss)

    If pnStatus = ST_TicketReturned Then
        nCount = ArrayLength(paszUserID)
        If nCount > 0 Then
            szTemp = " AND rst.user_id IN ('"
            For i = 1 To nCount
                szTemp = szTemp & ResolveDisplay(paszUserID(i)) & "','"
            Next i
            szTemp = Left(szTemp, Len(szTemp) - 2) & ")"
        End If
        If ToDBDateTime(pdyStart) = ToDBDateTime("0:00:00") Or ToDBDateTime(pdyEnd) = ToDBDateTime("0:00:00") Then
           szTempTime = ""
        Else
           szTempTime = " AND rst.operation_time >=" & "'" & ToDBDateTime(pdyStart) & "'" & " AND rst.operation_time <" & " '" & ToDBDateTime(pdyEnd) & "' "
        End If
        If szStartTicketNo <> "" Then
           szTempTicket = " AND sst.ticket_id>='" & szStartTicketNo & "'"
        Else
           szStartTicketNo = ""
        End If
        If szEndTicketNo <> "" Then
           szTempTicket = szTempTicket & " AND sst.ticket_id<='" & szEndTicketNo & "'"
        Else
           szStartTicketNo = szTempTicket & ""
        End If

        If szBusid <> "" Then
            szTempBusId = " AND sst.bus_id='" & szBusid & "'"
        Else
           szTempBusId = ""
        End If
        If szIDCardNo <> "" Then
           szTempRealName = szTempRealName & " AND sst.id_card_no='" & szIDCardNo & "'"
        End If
        If szPersonName <> "" Then
           szTempRealName = szTempRealName & " AND sst.person_name='" & szPersonName & "'"
        End If
        
        szSql = "SELECT sst.ticket_id,tfo.ticket_type_name,sst.bus_id,sst.bus_date,sst.ticket_type,sst.user_id as sale_user_id,sst.seat_no," & _
                " sst.ticket_price, sst.operation_time ,rst.operation_time as return_ticket_time,sst.card_type ,sst.id_card_no ,sst.person_name ," & _
                " rst.credence_sheet_id , rst.return_charge, rst.user_id , sst.des_station_name,ssi.sell_station_name " & _
                " FROM Agent_ticket_return_lst rst,Agent_ticket_sale_lst sst,Ticket_type_code tfo,sell_station_info ssi " & _
                " WHERE rst.ticket_id=sst.ticket_id AND tfo.ticket_type_id=sst.ticket_type AND sst.sell_station_id=ssi.sell_station_id "
        szOrder = " ORDER BY rst.user_id  , sst.ticket_id,rst.operation_time "
    ElseIf pnStatus = ST_TicketSellChange Then
        nCount = ArrayLength(paszUserID)
        If nCount > 0 Then
            szTemp = " AND ftk.change_user_id IN ('"
            For i = 1 To nCount
                szTemp = szTemp & ResolveDisplay(paszUserID(i)) & "','"
            Next i
            szTemp = Left(szTemp, Len(szTemp) - 2) & ")"
        End If
        If ToDBDateTime(pdyStart) = ToDBDateTime("0:00:00") Or ToDBDateTime(pdyEnd) = ToDBDateTime("0:00:00") Then
           szTempTime = ""
        Else
           szTempTime = " AND ftk.change_time >=" & "'" & ToDBDateTime(pdyStart) & "'" & " AND ftk.change_time <" & " '" & ToDBDateTime(pdyEnd) & "' "
        End If
        If szStartTicketNo <> "" Then
           szTempTicket = " AND ftk.new_ticket_id>='" & szStartTicketNo & "'"
        Else
           szStartTicketNo = ""
        End If
        If szEndTicketNo <> "" Then
           szTempTicket = szTempTicket & " AND ftk.new_ticket_id<='" & szEndTicketNo & "'"
        Else
           szStartTicketNo = szTempTicket & ""
        End If
        
         If szBusid <> "" Then
            szTempBusId = " AND (ctk.new_bus_id='" & szBusid & "' OR ftk.bus_id='" & szBusid & "')"
        Else
           szTempBusId = ""
        End If
        If szIDCardNo <> "" Then
           szTempRealName = szTempRealName & " AND ftk.id_card_no='" & szIDCardNo & "'"
        End If
        If szPersonName <> "" Then
           szTempRealName = szTempRealName & " AND ftk.person_name='" & szPersonName & "'"
        End If
        
        szSql = "SELECT ftk.*, ctk.new_bus_id,ctk.new_bus_date,ctk.new_seat_no,ctk.new_ticket_price,ctk.new_station_name,tfo.ticket_type_name,ssi.sell_station_name FROM " & _
              " (SELECT cst.new_ticket_id,cst.change_time,cst.change_charge,cst.fORmer_ticket_id,sst.card_type ,sst.id_card_no ,sst.person_name ," & _
              " cst.fORmer_ticket_price,cst.user_id as change_user_id,sst.bus_date,sst.bus_id, " & _
              " sst.des_station_id,sst.ticket_type,sst.user_id as sale_user_id, " & _
              " sst.seat_no as fORmer_seat_no,sfo.station_name as fORmer_station_name, " & _
              " sst.operation_time as sale_ticket_time  , sst.user_id as user_id ,sst.sell_station_id" & _
              " FROM Agent_ticket_change_lst cst,Agent_ticket_sale_lst sst,station_info sfo " & _
              " WHERE cst.fORmer_ticket_id=sst.ticket_id AND sfo.station_id=sst.des_station_id ) ftk, sell_station_info ssi," & _
              " (SELECT cst.new_ticket_id,sst.seat_no  as new_seat_no,sst.bus_id as new_bus_id, " & _
              " sst.bus_date as new_bus_date,sst.des_station_id as new_des_station_id, " & _
              " sfo.station_name as new_station_name,sst.ticket_price as new_ticket_price" & _
              " FROM ticket_change_lst cst,Ticket_sell_lst sst,station_info sfo " & _
              " WHERE cst.new_ticket_id=sst.ticket_id AND sfo.station_id=sst.des_station_id) ctk ," & _
              " Ticket_type_code tfo " & _
              " WHERE ftk.new_ticket_id = ctk.new_ticket_id AND tfo.ticket_type_id = ftk.ticket_type AND ftk.sell_station_id=ssi.sell_station_id"
              
        szOrder = " ORDER BY ftk.user_id "
    ElseIf pnStatus = ST_TicketChecked Then
        
        szSql = "SELECT cst.ticket_id,cst.bus_date,cst.bus_id,cst.bus_serial_no,cst.check_time,cst.ticket_type," & _
                " cst.ticket_price,cst.end_station_id ,sst.ticket_type,sst.user_id,sst.seat_no,sst.ticket_price,sst.card_type ,sst.id_card_no ,sst.person_name ," & _
                " tfo.ticket_type_name , sfo.station_name,sfo.station_name as des_station_name,ssi.sell_station_name  , s.check_sheet_id " & _
                " FROM check_ticket_lst cst,Ticket_sell_lst sst,Ticket_type_code tfo,station_info sfo,sell_station_info ssi  , sheet_lst s " & _
                " WHERE cst.ticket_id = sst.ticket_id AND tfo.ticket_type_id = sst.ticket_type AND sfo.station_id = cst.end_station_id " & _
                " AND sst.sell_station_id=ssi.sell_station_id" & _
                " and s.bus_date = cst.bus_date and s.bus_id = cst.bus_id and s.bus_serial_no = cst.bus_serial_no and s.check_gate_id = cst.check_gate_id "
        If ToDBDateTime(pdyStart) = ToDBDateTime("0:00:00") Or ToDBDateTime(pdyEnd) = ToDBDateTime("0:00:00") Then
           szTempTime = ""
        Else
           szTempTime = " AND cst.check_time >=" & "'" & ToDBDateTime(pdyStart) & "'" & " AND cst.check_time <" & " '" & ToDBDateTime(pdyEnd) & "' "
        End If
        
        If szBusid <> "" Then
            szTempBusId = " AND cst.bus_id='" & szBusid & "'"
        Else
           szTempBusId = ""
        End If
        
        If szStartTicketNo <> "" Then
            szTempTicket = " AND cst.ticket_id>='" & szStartTicketNo & "'"
        Else
            szStartTicketNo = ""
        End If
        If szEndTicketNo <> "" Then
            szTempTicket = szTempTicket & " AND cst.ticket_id<='" & szEndTicketNo & "'"
        Else
            szStartTicketNo = szTempTicket & ""
        End If
        If szIDCardNo <> "" Then
           szTempRealName = szTempRealName & " AND sst.id_card_no='" & szIDCardNo & "'"
        End If
        If szPersonName <> "" Then
           szTempRealName = szTempRealName & " AND sst.person_name='" & szPersonName & "'"
        End If
    Else
        nCount = ArrayLength(paszUserID)
        If nCount > 0 Then
            szTemp = " AND user_id IN ('"
            For i = 1 To nCount
                szTemp = szTemp & ResolveDisplay(paszUserID(i)) & "','"
            Next i
            szTemp = Left(szTemp, Len(szTemp) - 2) & ")"
        End If
        If pnStatus = 0 Then
            szSql = "   SELECT s.*, t.ticket_type_name,ssi.sell_station_name  , CASE s.insurance WHEN 0 THEN '无' ELSE '有' END insurance_name " _
                & " , CASE s.status WHEN 1 THEN '正常售出' WHEN 2 THEN '改签售出' WHEN 5 THEN '废票' WHEN 9 THEN '被改签' WHEN 17 THEN '退票' WHEN 33 THEN '已检' ELSE convert(char(2),s.status) END status_name " _
                & " FROM Agent_ticket_sale_lst  s,Ticket_type_code t,sell_station_info ssi  " _
                & " WHERE s.ticket_type=t.ticket_type_id AND ssi.sell_station_id=s.sell_station_id"
            
            If szIDCardNo <> "" Then
               szTempRealName = szTempRealName & " AND s.id_card_no='" & szIDCardNo & "'"
            End If
            If szPersonName <> "" Then
               szTempRealName = szTempRealName & " AND s.person_name='" & szPersonName & "'"
            End If
            
            szOrder = " ORDER BY s.ticket_id "
        ElseIf pnStatus = ST_TicketNormal Then
            szSql = "   SELECT s.*, t.ticket_type_name,ssi.sell_station_name  , CASE s.insurance WHEN 0 THEN '无' ELSE '有' END insurance_name " _
                & " , CASE s.status WHEN 1 THEN '正常售出' WHEN 2 THEN '改签售出' WHEN 5 THEN '废票' WHEN 9 THEN '被改签' WHEN 17 THEN '退票' WHEN 33 THEN '已检' ELSE convert(char(2),s.status) END status_name " _
                & " FROM Agent_ticket_sale_lst  s,Ticket_type_code t,sell_station_info ssi,env_bus_info ebi " _
                & " WHERE s.status='1' AND s.ticket_type=t.ticket_type_id AND ssi.sell_station_id=s.sell_station_id AND s.bus_date=ebi.bus_date AND s.bus_id=ebi.bus_id"
                
            If szIDCardNo <> "" Then
               szTempRealName = szTempRealName & " AND s.id_card_no='" & szIDCardNo & "'"
            End If
            If szPersonName <> "" Then
               szTempRealName = szTempRealName & " AND s.person_name='" & szPersonName & "'"
            End If
            
            szOrder = " ORDER BY s.ticket_id "
        ElseIf pnStatus = ST_TicketCanceled Then
            If nCount > 0 Then
                 szTemp = " AND s.user_id IN ('"
                 For i = 1 To nCount
                     szTemp = szTemp & ResolveDisplay(paszUserID(i)) & "','"
                 Next i
                 szTemp = Left(szTemp, Len(szTemp) - 2) & ")"
               
            End If
        '            szsql = "SELECT zst.*,sst.bus_date,sst.bus_id,sst.ticket_type,sst.ticket_price,sst.seat_no,sfo.station_name,tfo.ticket_type_name " & _
        '                   " FROM cancel_ticket_lst zst,Ticket_sell_lst sst,station_info sfo,Ticket_type_code tfo " & _
        '                   " WHERE zst.ticket_id = sst.ticket_id AND sst.des_station_id = sfo.station_id AND sst.ticket_type = tfo.ticket_type_id "
        '
            szSql = "SELECT s.*,s.cancel_time cancel_ticket_time ,sst.bus_date,sst.bus_id,sst.ticket_type,sst.ticket_price,sst.seat_no,sst.card_type ,sst.id_card_no ,sst.person_name ," _
                & "     tfo.ticket_type_name , sst.user_id AS sale_user_id ,sst.operation_time , sst.des_station_name,ssi.sell_station_name  , CASE sst.insurance WHEN 0 THEN '无' ELSE '有' END insurance_name " _
                & " " _
                & " FROM Agent_ticket_cancel_lst s,Agent_ticket_sale_lst sst,Ticket_type_code tfo,sell_station_info ssi " _
                & " WHERE s.ticket_id = sst.ticket_id AND sst.ticket_type = tfo.ticket_type_id AND ssi.sell_station_id=sst.sell_station_id"
            
            If szBusid <> "" Then
                szTempBusId = " AND sst.bus_id='" & szBusid & "'"
            Else
               szTempBusId = ""
            End If
            
            If szIDCardNo <> "" Then
               szTempRealName = szTempRealName & " AND sst.id_card_no='" & szIDCardNo & "'"
            End If
            If szPersonName <> "" Then
               szTempRealName = szTempRealName & " AND sst.person_name='" & szPersonName & "'"
            End If
            
            szOrder = " ORDER BY s.ticket_id " 's.user_id,s.ticket_id,s.cancel_time  "
        End If
    
        If ToDBDateTime(pdyStart) = ToDBDateTime("0:00:00") Or ToDBDateTime(pdyEnd) = ToDBDateTime("0:00:00") Then
            szTempTime = ""
        ElseIf pnStatus = ST_TicketCanceled Then
            szTempTime = " AND cancel_time >=" & "'" & ToDBDateTime(pdyStart) & "'" & " AND cancel_time <" & " '" & ToDBDateTime(pdyEnd) & "' "
        ElseIf pnStatus = ST_TicketNormal Then
            szTempTime = " AND ebi.bus_start_time >=" & "'" & ToDBDateTime(pdyStart) & "'" & " AND ebi.bus_start_time <" & " '" & ToDBDateTime(pdyEnd) & "' "
        Else
            szTempTime = " AND operation_time >=" & "'" & ToDBDateTime(pdyStart) & "'" & " AND operation_time <" & " '" & ToDBDateTime(pdyEnd) & "' "
        End If
        If szStartTicketNo <> "" Then
            szTempTicket = " AND s.ticket_id>='" & szStartTicketNo & "'"
        Else
            szStartTicketNo = ""
        End If
        If szEndTicketNo <> "" Then
            szTempTicket = szTempTicket & " AND s.ticket_id<='" & szEndTicketNo & "'"
        Else
            szStartTicketNo = szTempTicket & ""
        End If
              
        If pnStatus <> ST_TicketCanceled Then
            If szBusid <> "" And pnStatus <> ST_TicketNormal Then
                szTempBusId = " AND bus_id='" & szBusid & "'"
            ElseIf pnStatus = ST_TicketNormal Then
                szTempBusId = " AND s.bus_id='" & szBusid & "'"
            Else
               szTempBusId = ""
            End If
        End If
    End If
    
    '查询空白票，退票凭证用票zyw
    If pnStatus = 0 Then
        If szStartTicketNo <> "" Then
           szTempTicketNull = " AND firstnullticket_no>='" & szStartTicketNo & "'"
           szTempTicketReturn = " AND credence_sheet_id>='" & szStartTicketNo & "'"
        End If
        If szEndTicketNo <> "" Then
           szTempTicketNull = szTempTicketNull & " AND lastnullticket_no<='" & szEndTicketNo & "'"
           szTempTicketReturn = szTempTicketReturn & " AND credence_sheet_id<='" & szEndTicketNo & "'"
        End If
        
        szNullSql = "SELECT RTRIM(firstnullticket_no) + ' - ' + RTRIM(lastnullticket_no) ticket_id ,'" & cszEmptyDateStr & "' bus_date ,'空白票' bus_id ," _
            & "'空白票' service_unit_id ,'空白票' sell_station_id ,'空白票' start_station_id ,'空白票' start_station_name ,'空白票' des_station_id ," _
            & "'0' sale_ticket_type ,'0' ticket_type ,'' seat_type_id ,'空白票' des_station_name ,user_id ,'空白票' seat_no ,'0' ticket_price ,'0' status ," _
            & " getnullticket_date operation_time ,'0' insurance ,'' card_type ,'' id_card_no ,'' person_name ,'' sex ,'' person_picture ,'空白票' ticket_type_name ,'空白票' sell_station_name ,'空白票' insurance_name ," _
            & "'未使用' status_name FROM getnullticket"
        szNullSql = szNullSql & " WHERE getnullticket_date >=" & "'" & ToDBDateTime(pdyStart) & "'" & " AND getnullticket_date <" & " '" & ToDBDateTime(pdyEnd) & "'"
        szNullSql = szNullSql & szTemp & szTempTicketNull
        
        
        szReturnVouchers = "SELECT RTRIM(credence_sheet_id) ticket_id ,bus_date ,bus_id ," _
            & "service_unit_id ,sell_station_id, start_station_id ,start_station_name ,des_station_id ," _
            & "'0' sale_ticket_type ,'0' ticket_type ,'' seat_type_id ,des_station_name ,user_id ,seat_number seat_no ,return_charge ticket_price ,'0' status ," _
            & " operation_time ,'0' insurance ,'' card_type ,'' id_card_no ,'' person_name ,'' sex ,'' person_picture ,'退票凭证' ticket_type_name ,'退票凭证' sell_station_name ,'退票凭证' insurance_name ," _
            & "'退票凭证' status_name FROM agent_ticket_return_lst"
        szReturnVouchers = szReturnVouchers & " WHERE operation_time >=" & "'" & ToDBDateTime(pdyStart) & "'" & " AND operation_time <" & " '" & ToDBDateTime(pdyEnd) & "'"
        szReturnVouchers = szReturnVouchers & szTemp & szTempTicketReturn
    End If
    
    If pnStatus = 0 Then
        szSql = szSql & szTempTime & szTemp & szTempTicket & szTempBusId & szTempRealName & " UNION " & szNullSql & " UNION " & szReturnVouchers & szOrder
    Else
        szSql = szSql & szTempTime & szTemp & szTempTicket & szTempBusId & szTempRealName & szOrder
    End If
    
    Set rsTemp = oDB.Execute(szSql)
    Set SellerTicketDetailFromAgent = rsTemp
    
    
End Function

'车次实名制查询
Public Function QueryBusRealNameInfo(Optional StartBusDate As Date = cszEmptyDateStr, Optional EndBusDate As Date = cszEmptyDateStr, Optional SellStationID As String = "", Optional BusID As String = "", Optional QueryTypeCheck As Boolean = False, Optional StopCheck As Boolean = False, Optional BusSerialNo As Integer = -1, Optional SheetNo As String = "") As Recordset
    'QueryTypeCheck True:按检票查询  False:按售票查询
    'StopCheck True:只查停检  False:
    
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim szWhere As String
    oDB.ConnectionString = GetConnectionStr()
    
    If QueryTypeCheck = False Then
        szSql = "SELECT a.* ,RTRIM(vi.license_tag_no) license_tag_no ,RTRIM(si.station_name) station_name ,RTRIM(ttc.ticket_type_name) ticket_type_name ,RTRIM(s.ticket_id) ticket_id ,s.seat_no ,s.ticket_type ,s.card_type ,s.id_card_no ,s.person_name ,s.sex ,s.person_picture" _
            & " FROM ticket_sell_lst s ,env_bus_allot_lst a ,station_info si ,ticket_type_code ttc ,vehicle_info vi ,env_bus_info e" _
            & " WHERE s.bus_date =a.bus_date AND s.bus_id =a.bus_id AND s.sell_station_id =a.sell_station_id AND e.vehicle_id =vi.vehicle_id" _
            & " AND s.des_station_id =si.station_id AND s.ticket_type =ttc.ticket_type_id AND e.bus_date =a.bus_date AND e.bus_id =a.bus_id" _
            & " AND s.status IN (" & ST_TicketNormal & "," & ST_TicketNormal & "|" & ST_TicketChecked & "," & ST_TicketSellChange & "," & ST_TicketSellChange & "|" & ST_TicketChecked & ")"
        If StartBusDate <> cszEmptyDateStr Then
            szWhere = szWhere & " AND s.bus_date >=" & TransFieldValueToString(StartBusDate)
        End If
        If EndBusDate <> cszEmptyDateStr Then
            szWhere = szWhere & " AND s.bus_date <=" & TransFieldValueToString(EndBusDate)
        End If
        If SellStationID <> "" Then
            szWhere = szWhere & " AND s.sell_station_id =" & TransFieldValueToString(SellStationID)
        End If
        If BusID <> "" Then
            szWhere = szWhere & " AND s.bus_id =" & TransFieldValueToString(BusID)
        End If
        szSql = szSql & szWhere & " ORDER BY s.bus_date ,s.bus_id ,s.seat_no"
    Else
        If StopCheck = False Then
            szSql = "SELECT a.* ,RTRIM(vi.license_tag_no) license_tag_no ,RTRIM(si.station_name) station_name ,RTRIM(ttc.ticket_type_name) ticket_type_name ,RTRIM(s.ticket_id) ticket_id ,s.seat_no ,s.ticket_type ,s.card_type ,s.id_card_no ,s.person_name ,s.sex ,s.person_picture" _
                & " FROM ticket_sell_lst s ,bus_check_lst a ,station_info si ,ticket_type_code ttc ,check_ticket_lst c ,checkgate_info ci ,vehicle_info vi" _
                & " WHERE c.bus_date =a.bus_date AND c.bus_id =a.bus_id AND c.check_gate_id =a.check_gate_id AND c.bus_serial_no =a.bus_serial_no AND a.vehicle_id =vi.vehicle_id" _
                & " AND s.des_station_id =si.station_id AND s.ticket_type =ttc.ticket_type_id AND s.ticket_id =c.ticket_id AND a.check_gate_id =ci.check_gate_id"
            If StartBusDate <> cszEmptyDateStr Then
                szWhere = szWhere & " AND c.bus_date >=" & TransFieldValueToString(StartBusDate)
            End If
            If EndBusDate <> cszEmptyDateStr Then
                szWhere = szWhere & " AND c.bus_date <=" & TransFieldValueToString(EndBusDate)
            End If
            If SellStationID <> "" Then
                szWhere = szWhere & " AND ci.sell_station_id =" & TransFieldValueToString(SellStationID)
            End If
            If BusID <> "" Then
                szWhere = szWhere & " AND c.bus_id =" & TransFieldValueToString(BusID)
            End If
            If BusSerialNo <> -1 Then
                szWhere = szWhere & " AND c.bus_serial_no =" & TransFieldValueToString(BusSerialNo)
            End If
            szSql = szSql & szWhere & " ORDER BY c.bus_date ,c.bus_id ,s.seat_no"
        Else
            szSql = "SELECT RTRIM(vi.license_tag_no) license_tag_no ,a.bus_id ,a.bus_date ,ci.sell_station_id ,CASE a.bus_serial_no WHEN 0 THEN a.bus_start_time ELSE a.sheet_make_date END bus_start_time ,RTRIM(si.station_name) station_name ,RTRIM(ttc.ticket_type_name) ticket_type_name ,RTRIM(s.ticket_id) ticket_id ,s.seat_no ,s.ticket_type ,s.card_type ,s.id_card_no ,s.person_name ,s.sex ,s.person_picture" _
                & " FROM ticket_sell_lst s ,sheet_lst a ,station_info si ,ticket_type_code ttc ,check_ticket_lst c ,checkgate_info ci ,vehicle_info vi" _
                & " WHERE c.bus_date =a.bus_date AND c.bus_id =a.bus_id AND c.check_gate_id =a.check_gate_id AND c.bus_serial_no =a.bus_serial_no AND a.valid_mark =1" _
                & " AND s.des_station_id =si.station_id AND s.ticket_type =ttc.ticket_type_id AND s.ticket_id =c.ticket_id AND a.check_gate_id =ci.check_gate_id AND a.vehicle_id =vi.vehicle_id"
            If StartBusDate <> cszEmptyDateStr Then
                szWhere = szWhere & " AND a.bus_date >=" & TransFieldValueToString(StartBusDate)
            End If
            If EndBusDate <> cszEmptyDateStr Then
                szWhere = szWhere & " AND a.bus_date <=" & TransFieldValueToString(EndBusDate)
            End If
            If SellStationID <> "" Then
                szWhere = szWhere & " AND ci.sell_station_id =" & TransFieldValueToString(SellStationID)
            End If
            If BusID <> "" Then
                szWhere = szWhere & " AND a.bus_id =" & TransFieldValueToString(BusID)
            End If
            If BusSerialNo <> -1 Then
                szWhere = szWhere & " AND c.bus_serial_no =" & TransFieldValueToString(BusSerialNo)
            End If
            If SheetNo <> "" Then
                szWhere = szWhere & " AND a.check_sheet_id =" & TransFieldValueToString(SheetNo)
            End If
            szSql = szSql & szWhere & " ORDER BY a.bus_date ,a.bus_id ,s.seat_no"
        End If
    End If
    
    Set rsTemp = oDB.Execute(szSql)
    Set QueryBusRealNameInfo = rsTemp
End Function

Public Function GetOperator(pdyFromDate As Date, pdyToDate As Date, Optional pszSellStationID As String = "") As String()
'得到指定时间内，进行过售票、废票、退票及改签的操作员
'该接口为了每日结算用
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szWhere As String
    Dim aszTemp() As String
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_QuerySellerEveryDaySellDetail
    
    
    
    oDB.ConnectionString = GetConnectionStr()
    If pszSellStationID <> "" Then
        szWhere = "" ' " AND sell_station_id='" & pszSellStationID & "'"
    End If
    szSql = " SELECT DISTINCT user_id FROM Agent_ticket_sale_lst " _
        & " WHERE operation_time >= '" & ToDBDateTime(pdyFromDate) & "' AND operation_time < '" & ToDBDateTime(pdyToDate) & "'" & szWhere
    Set rsTemp = oDB.Execute(szSql)
    aszTemp = GetUniqueTeam(rsTemp, aszTemp)
    szSql = " SELECT DISTINCT user_id FROM Agent_ticket_cancel_lst WHERE cancel_time >= '" & ToDBDateTime(pdyFromDate) & "' AND cancel_time < '" & ToDBDateTime(pdyToDate) & "'"
    Set rsTemp = oDB.Execute(szSql)
    aszTemp = GetUniqueTeam(rsTemp, aszTemp)
    szSql = " SELECT DISTINCT user_id FROM Agent_ticket_return_lst WHERE operation_time >= '" & ToDBDateTime(pdyFromDate) & "' AND operation_time < '" & ToDBDateTime(pdyToDate) & "'" & szWhere
    Set rsTemp = oDB.Execute(szSql)
    aszTemp = GetUniqueTeam(rsTemp, aszTemp)
    szSql = " SELECT DISTINCT user_id FROM Agent_ticket_change_lst WHERE change_time >= '" & ToDBDateTime(pdyFromDate) & "' AND change_time < '" & ToDBDateTime(pdyToDate) & "'"
    Set rsTemp = oDB.Execute(szSql)
    aszTemp = GetUniqueTeam(rsTemp, aszTemp)
    
    GetOperator = aszTemp
    
End Function


'汇总售票员统计,同时列出各票价项明细
Public Function SellerPriceItemDateStat(UserID() As String, FromDate As Date, ToDate As Date) As Recordset
'pnExtraStatus =1 表示售票，2表示补票，3表示全部



    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetBusStatBySaleTime
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    Dim nUserCount As Integer
    Dim i As Integer
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    
    szSql = "SELECT RTRIM(user_info.user_id) +'['+RTRIM(user_name)+']' AS user_id ," _
        & " SUM(ticket_sale_number) AS ticket_sale_number, " _
        & " SUM(ticket_sale_amount) AS ticket_sale_amount," _
        & " SUM(ticket_cancel_number) AS ticket_cancel_number," _
        & " SUM(ticket_cancel_amount) AS ticket_cancel_amount," _
        & " SUM(ticket_return_number) AS ticket_return_number," _
        & " SUM(ticket_return_amount) AS ticket_return_amount," _
        & " SUM(ticket_return_charge) AS ticket_return_charge," _
        & " SUM(ticket_change_number) AS ticket_change_number," _
        & " SUM(ticket_change_amount) AS ticket_change_amount," _
        & " SUM(ticket_change_charge) AS ticket_change_charge " _
        & " , SUM(base_price) base_price , SUM (price_item_1) price_item_1, SUM(price_item_2) price_item_2   " _
        & " , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 , SUM(price_item_5) price_item_5  " _
        & " , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 ,SUM(price_item_8) price_item_8  " _
        & " , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 " _
        & " , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 " _
        & " , SUM(price_item_15) price_item_15 " _
        & " , (SUM(ticket_sale_number) - SUM(ticket_return_number) - sum(ticket_change_number) ) AS passenger_number " _
        & " , (SUM(ticket_sale_amount)  - SUM(ticket_return_amount) " _
        & "     - SUM(ticket_change_amount) ) AS ticket_price " _
        & " , (SUM(ticket_sale_amount) - SUM(ticket_return_amount) " _
        & "     - SUM(ticket_change_amount) + SUM(ticket_return_charge) +Sum(ticket_change_charge) ) AS total_ticket_price  " _
    
        
szSql = szSql & " FROM Stat_saler_ticket_lst,user_info WHERE " _
    & " Stat_saler_ticket_lst.user_id=user_info.user_id AND " _
    & " bus_date>='" & ToDBDate(FromDate) & "' AND " _
    & " bus_date<='" & ToDBDate(ToDate) & "' "
    
    nUserCount = ArrayLength(UserID)
    If nUserCount > 0 Then
        szSql = szSql & " AND (Stat_saler_ticket_lst.user_id='" & UserID(1) & "' "
        For i = 2 To nUserCount
            szSql = szSql & " OR Stat_saler_ticket_lst.user_id='" & UserID(i) & "' "
        Next
        szSql = szSql & ")"
    End If

    szSql = szSql & " GROUP BY user_info.user_id,user_name ORDER BY user_id "
    Set rsTemp = oDB.Execute(szSql)
    
    Set SellerPriceItemDateStat = rsTemp
    
End Function

'*****************************************
'预售票平衡统计表
Public Function PreSellTicketCount(asSellerStation() As String, FromDate As Date) As Recordset
'    AssertActiveUserValid m_oActiveUser, ERR_TicketSellerDim
'    AssertHaveRight m_oActiveUser, RIGHT_SellerDateStat
    
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
'    Dim rsData As New Recordset
'    Dim rsData2 As New Recordset
'    Dim rsdata3 As New Recordset
'    Dim rs As New Recordset
    
    Dim nUserCount As Integer
    Dim i As Integer
    Dim nCount As Integer
'    Dim j As Integer
'    Dim lNumber As Long
'    Dim lNumber2 As Long
'    Dim aswhere As String
    Dim szTemp As String
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
'    sztemp = "("
'    For j = 0 To ArrayLength(asSellerStation) - 1
'        sztemp = sztemp & "u.sell_station_id='" & asSellerStation(j) & "' or "
'    Next j
'    If Len(sztemp) - 5 >= 0 Then
'        sztemp = Left(sztemp, Len(sztemp) - 5)
'        sztemp = sztemp & "')"
'    End If
    
'    szTemp = TeamToString2(asSellerStation)
    
'        '上月售本月
'         szSql = " Select t.bus_id ,sum(t.ticket_price) pre into #a from ticket_sell_lst t,user_info u " _
'            & " where  t.operation_time>='" & GetFirstMonthDay(DateAdd("m", -1, FromDate)) & "' and t.operation_time<'" & GetLastMonthDay(DateAdd("m", -1, FromDate)) & " " & cszENdTimeStr & "'and t.bus_date>='" & GetFirstMonthDay(FromDate) _
'            & "'  AND t.status not in('5','9','17','18') and t.user_id=u.user_id AND " & sztemp & " group by t.bus_id"
'
'        oDB.Execute szSql
'
'        szSql = " SELECT t.bus_id ,SUM(tr.return_charge) AS return_charge into #c from ticket_sell_lst t ,ticket_return_lst tr,user_info u  " _
'           & "where  t.operation_time>='" & GetFirstMonthDay(DateAdd("m", -1, FromDate)) & "' and t.operation_time<'" & GetLastMonthDay(DateAdd("m", -1, FromDate)) & " " & cszENdTimeStr & "'and t.bus_date>='" & GetFirstMonthDay(FromDate) _
'           & "' AND " & sztemp & " and tr.ticket_id=t.ticket_id and tr.user_id=u.user_id  group by t.bus_id"
'        oDB.Execute szSql
'
'        szSql = " SELECT  #a.bus_id,(#a.pre+isnull(#c.return_charge,0)) as pre ,isnull(#c.return_charge,0) as return_charge into #s from #a,#c where #a.bus_id*=#c.bus_id"
'        oDB.Execute szSql
'        '本月售本月
'        szSql = " Select t.bus_id ,sum(t.ticket_price) cur into #e from ticket_sell_lst t,user_info u " _
'            & " where t.operation_time>='" & GetFirstMonthDay(FromDate) & "' and t.operation_time<'" & GetLastMonthDay(FromDate) & " " & cszENdTimeStr & "'and t.bus_date<='" & GetLastMonthDay(FromDate) & " " & cszENdTimeStr _
'            & "'  AND  t.status not in('5','9','17','18') and t.user_id=u.user_id AND " & sztemp & " group by t.bus_id"
'
'        oDB.Execute szSql
'
'        szSql = " SELECT t.bus_id ,SUM(tr.return_charge) AS return_charge into #d from ticket_sell_lst t ,ticket_return_lst tr,user_info u  " _
'           & "where  t.operation_time>='" & GetFirstMonthDay(FromDate) & "' and t.operation_time<'" & GetLastMonthDay(FromDate) & " " & cszENdTimeStr & "'and t.bus_date<='" & GetLastMonthDay(FromDate) & " " & cszENdTimeStr _
'           & "' AND " & sztemp & " and tr.ticket_id=t.ticket_id and tr.user_id=u.user_id  group by t.bus_id"
'        oDB.Execute szSql
'
'        szSql = " SELECT  #e.bus_id,(#e.cur+isnull(#d.return_charge,0)) as cur,isnull(#d.return_charge,0) as return_charge into #b from #e,#d where #e.bus_id*=#d.bus_id"
'        oDB.Execute szSql
'        '本月售下月
'        szSql = " Select t.bus_id ,sum(t.ticket_price) next into #f from ticket_sell_lst t,user_info u " _
'            & " where   t.operation_time>='" & GetFirstMonthDay(FromDate) & "' and t.operation_time<'" & GetLastMonthDay(FromDate) & " " & cszENdTimeStr & "'and t.bus_date>='" & GetFirstMonthDay(DateAdd("m", 1, FromDate)) & " 'and t.bus_date<='" & GetLastMonthDay(DateAdd("m", 1, FromDate)) & " " & cszENdTimeStr _
'            & "'  AND  t.status not in('5','9','17','18') and t.user_id=u.user_id AND " & sztemp & " group by t.bus_id"
'        oDB.Execute szSql, lNumber2
'
'         szSql = " SELECT t.bus_id ,SUM(tr.return_charge) AS return_charge into #g from ticket_sell_lst t ,ticket_return_lst tr,user_info u  " _
'           & "where  t.operation_time>='" & GetFirstMonthDay(FromDate) & "' and t.operation_time<'" & GetLastMonthDay(FromDate) & " " & cszENdTimeStr & "'and t.bus_date>='" & GetFirstMonthDay(DateAdd("m", 1, FromDate)) & " 'and t.bus_date<='" & GetLastMonthDay(DateAdd("m", 1, FromDate)) & " " & cszENdTimeStr _
'           & "' AND " & sztemp & " and tr.ticket_id=t.ticket_id and tr.user_id=u.user_id group by t.bus_id"
'        oDB.Execute szSql
'
'        szSql = " SELECT  #f.bus_id,(#f.next+isnull(#g.return_charge,0)) as next,isnull(#g.return_charge,0) as return_charge into #x from #f,#g where #f.bus_id*=#g.bus_id"
'        oDB.Execute szSql
'
'    szSql = "select distinct t1.bus_id,r.route_name,#s.pre as pre_count,#b.cur as cur_count,#x.next as next_count,isnull(#s.pre,0)+isnull(#b.cur,0) bus_money,isnull(#x.next,0)+ isnull(#b.cur,0) ticket_money "
'
'    szSql = szSql & " from(select bus_id from #s Union select bus_id from #b Union select bus_id from #x)t1 " _
'            & " left join #s on t1.bus_id=#s.bus_id" _
'            & " left join #b on t1.bus_id=#b.bus_id" _
'            & " left join #x on t1.bus_id=#x.bus_id ,env_bus_info e,route_info r where e.bus_id=t1.bus_id and r.route_id=e.route_id"


    szSql = " EXEC QueryPreSellTicketCount " & TransFieldValueToString(GetFirstMonthDay(FromDate)) & " , " & TransFieldValueToString(GetLastMonthDay(FromDate))
    
    nCount = ArrayLength(asSellerStation) - 1
    
    For i = 0 To 9
        If i > nCount Then
            szTemp = szTemp & "," & TransFieldValueToString("")
        Else
            szTemp = szTemp & "," & TransFieldValueToString(asSellerStation(i))
        End If
        'sztemp = sztemp & "u.sell_station_id='" & asSellerStation(j) & "' or "
    Next i
    szSql = szSql & szTemp
    Set rsTemp = oDB.Execute(szSql)
    
    Set PreSellTicketCount = rsTemp
End Function
  
'*****************************************************
'预售票平衡统计明细表
'预售票平衡统计表
Public Function PreSellTicketLst(asSellerStation() As String, FromDate As Date) As Recordset
'    AssertActiveUserValid m_oActiveUser, ERR_TicketSellerDim
'    AssertHaveRight m_oActiveUser, RIGHT_SellerDateStat
    
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
'    Dim rsData As New Recordset
'    Dim rsData2 As New Recordset
'    Dim rsdata3 As New Recordset
'    Dim rs As New Recordset
    
'    Dim nUserCount As Integer
    Dim i As Integer
    
'    Dim j As Integer
'    Dim lNumber As Long
'    Dim lNumber2 As Long
'    Dim aswhere As String
    Dim szTemp As String
    Dim nCount As Integer
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
'
'    szTemp = "("
'    For j = 0 To ArrayLength(asSellerStation) - 1
'        szTemp = szTemp & "t.sell_station_id='" & asSellerStation(j) & "' or "
'    Next j
'    If Len(szTemp) - 5 >= 0 Then
'    szTemp = Left(szTemp, Len(szTemp) - 5)
'    szTemp = szTemp & "')"
'    End If
'
'
''    szSql = " Select t.bus_id ,sum(t.ticket_price) pre,count(t.ticket_id) as ticket1_number,SUM(tl.base_price) AS base_price1" _
''        & ",SUM(tl.price_item_1) AS price_item1_1 ,SUM(tl.price_item_2) AS price_item1_2,SUM(tl.price_item_3) AS price_item1_3" _
''        & ",SUM(tl.price_item_4) AS price_item1_4 ,SUM(tl.price_item_5) AS price_item1_5,SUM(tl.price_item_6) AS price_item1_6" _
''        & ",SUM(tl.price_item_7) AS price_item1_7 ,SUM(tl.price_item_8) AS price_item1_8,SUM(tl.price_item_9) AS price_item1_9" _
''        & ",SUM(tl.price_item_10) AS price_item1_10 , SUM(tl.price_item_11) AS price_item1_11" _
''        & ",SUM(tl.price_item_12) AS price_item1_12 , SUM(tl.price_item_13) AS price_item1_13 , SUM(tl.price_item_14) AS price_item1_14" _
''        & ",SUM(tl.price_item_15) AS price_item1_15 into #s from ticket_sell_lst t ,ticket_price_lst tl " _
''        & " where t.operation_time<t.bus_date and t.operation_time>='" & GetFirstMonthDay(DateAdd("m", -1, FromDate)) & "' and t.operation_time<'" & GetLastMonthDay(DateAdd("m", -1, FromDate)) & " " & cszENdTimeStr & "'and t.bus_date>='" & GetFirstMonthDay(FromDate) _
''        & "'  AND " & sztemp & "  and tl.ticket_id=t.ticket_id group by t.bus_id"
''
''    oDB.Execute szSql
''
''    szSql = " Select t.bus_id ,sum(t.ticket_price) cur,count(t.ticket_id) as ticket2_number,SUM(tl.base_price) AS base_price2" _
''        & ",SUM(tl.price_item_1) AS price_item2_1 ,SUM(tl.price_item_2) AS price_item2_2,SUM(tl.price_item_3) AS price_item2_3" _
''        & ",SUM(tl.price_item_4) AS price_item2_4 ,SUM(tl.price_item_5) AS price_item2_5,SUM(tl.price_item_6) AS price_item2_6" _
''        & ",SUM(tl.price_item_7) AS price_item2_7 ,SUM(tl.price_item_8) AS price_item2_8,SUM(tl.price_item_9) AS price_item2_9" _
''        & ",SUM(tl.price_item_10) AS price_item2_10 , SUM(tl.price_item_11) AS price_item2_11" _
''        & ",SUM(tl.price_item_12) AS price_item2_12 , SUM(tl.price_item_13) AS price_item2_13 , SUM(tl.price_item_14) AS price_item2_14" _
''        & ",SUM(tl.price_item_15) AS price_item2_15  into #b from ticket_sell_lst t,ticket_price_lst tl " _
''        & " where t.operation_time<t.bus_date and t.operation_time>='" & GetFirstMonthDay(FromDate) & "' and t.operation_time<'" & GetLastMonthDay(FromDate) & " " & cszENdTimeStr & "'and t.bus_date<='" & GetLastMonthDay(FromDate) & " " & cszENdTimeStr _
''        & "' AND " & sztemp & " and tl.ticket_id=t.ticket_id group by t.bus_id"
''
''    oDB.Execute szSql
'
'    szSql = " Select t.bus_id ,sum(t.ticket_price) next_count,count(t.ticket_id) as ticket3_number, SUM(tl.base_price) AS base_price3" _
'        & ",SUM(tl.price_item_1) AS price_item3_1 ,SUM(tl.price_item_2) AS price_item3_2,SUM(tl.price_item_3) AS price_item3_3" _
'        & ",SUM(tl.price_item_4) AS price_item3_4 ,SUM(tl.price_item_5) AS price_item3_5,SUM(tl.price_item_6) AS price_item3_6" _
'        & ",SUM(tl.price_item_7) AS price_item3_7 ,SUM(tl.price_item_8) AS price_item3_8,SUM(tl.price_item_9) AS price_item3_9" _
'        & ",SUM(tl.price_item_10) AS price_item3_10 , SUM(tl.price_item_11) AS price_item3_11" _
'        & ",SUM(tl.price_item_12) AS price_item3_12 , SUM(tl.price_item_13) AS price_item3_13 , SUM(tl.price_item_14) AS price_item3_14" _
'        & ",SUM(tl.price_item_15) AS price_item3_15 into #a from ticket_sell_lst t ,ticket_price_lst tl" _
'        & " where t.operation_time<t.bus_date and t.operation_time>='" & GetFirstMonthDay(FromDate) & "' and t.operation_time<'" & GetLastMonthDay(FromDate) & " " & cszENdTimeStr & "'and t.bus_date>='" & GetFirstMonthDay(DateAdd("m", 1, FromDate)) & " 'and t.bus_date<='" & GetLastMonthDay(DateAdd("m", 1, FromDate)) & " " & cszENdTimeStr _
'        & "' AND status not in('5','9','17','18') AND" & szTemp & " and tl.ticket_id=t.ticket_id  group by t.bus_id"
'        oDB.Execute szSql
'
'
'    szSql = " SELECT t.bus_id ,SUM(tr.return_charge) AS return_charge into #b from ticket_sell_lst t ,ticket_return_lst tr  " _
'        & "where t.operation_time<t.bus_date and t.operation_time>='" & GetFirstMonthDay(FromDate) & "' and t.operation_time<'" & GetLastMonthDay(FromDate) & " " & cszENdTimeStr & "'and t.bus_date>='" & GetFirstMonthDay(DateAdd("m", 1, FromDate)) & " 'and t.bus_date<='" & GetLastMonthDay(DateAdd("m", 1, FromDate)) & " " & cszENdTimeStr _
'        & "' AND " & szTemp & " and tr.ticket_id=t.ticket_id  group by t.bus_id"
'        oDB.Execute szSql
'    szSql = " SELECT  #a.bus_id,(#a.next_count+isnull(#b.return_charge,0)) as next_count,#a.ticket3_number,#a.base_price3" _
'            & ",#a.price_item3_1,#a.price_item3_2,#a.price_item3_3,#a.price_item3_4,#a.price_item3_5,#a.price_item3_6,#a.price_item3_7" _
'            & ",#a.price_item3_8,#a.price_item3_9,#a.price_item3_10,#a.price_item3_11,#a.price_item3_12,#a.price_item3_13,#a.price_item3_14" _
'            & ",#a.price_item3_15,isnull(#b.return_charge,0) as return_charge into #c from #a,#b where #a.bus_id*=#b.bus_id"
'        oDB.Execute szSql
'    szSql = "select distinct t1.bus_id,r.route_name,t1.* " _
'            & "from env_bus_info e,route_info r,#c t1 where e.bus_id=t1.bus_id and r.route_id=e.route_id"
'    Set rs = oDB.Execute(szSql)

    szSql = " EXEC QueryPreSellTicketLst " & TransFieldValueToString(GetFirstMonthDay(FromDate)) & " , " & TransFieldValueToString(GetLastMonthDay(FromDate))
    
    nCount = ArrayLength(asSellerStation) - 1
    
    For i = 0 To 9
        If i > nCount Then
            szTemp = szTemp & "," & TransFieldValueToString("")
        Else
            szTemp = szTemp & "," & TransFieldValueToString(asSellerStation(i))
        End If
        
    Next i
    szSql = szSql & szTemp
    Set rsTemp = oDB.Execute(szSql)
    Set PreSellTicketLst = rsTemp
End Function
  

