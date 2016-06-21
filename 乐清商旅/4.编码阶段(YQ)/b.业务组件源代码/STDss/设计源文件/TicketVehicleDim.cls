VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TicketVehicleDim"
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


Public Function GetVehicleStat(pszCompanyID As String, pdtFromDate As Date, pdtToDate As Date, Optional pnExtraStatus As eExtraStatus = TP_AllSold, Optional pszVehicleID As String = "", Optional pszSellStationID As String = "", Optional pblnCheck As Boolean = False) As Recordset
'pnExtraStatus =1 表示售票，2表示补票，3表示全部

    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    
'    If pblnCheck = False Then '售票
'        AssertHaveRight m_oActiveUser, RIGHT_GetVehicleStatBySaleTime
'    Else   '检票
'        AssertHaveRight m_oActiveUser, RIGHT_GetVehicleStatByBusDate
'    End If
                        
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    Dim szGroupBy As String
    Dim szRealTicketCount As String '实际售票张数
    Dim szRealTicketPrice As String '实际售票金额
    Dim szReturnCharge As String    '退票手续费
    Dim szTotalIncom As String    '营收合计
    Dim szTable As String '查找表名
    
    If pblnCheck = False Then '按售票
       szRealTicketCount = " isnull((SUM(passenger_number) - SUM(ticket_cancel_number)- SUM(ticket_change_number) - SUM(ticket_return_number)),0) "
       szRealTicketPrice = ", isnull((SUM(ticket_price) - SUM(ticket_cancel_amount)- SUM(ticket_return_amount)  - SUM(ticket_change_amount) ),0)"
       szTable = "stat_bus_startoff_lst "
       szReturnCharge = ", isnull(SUM(ticket_return_charge), 0) AS ticket_return_charge"
       szTotalIncom = ", isnull((SUM(ticket_price) - SUM(ticket_cancel_amount)- SUM(ticket_return_amount)  - SUM(ticket_change_amount) + SUM(ticket_return_charge) + Sum(ticket_change_charge)), 0) AS total_ticket_price  "
    Else '按检票
       szRealTicketCount = " isnull(SUM(passenger_number),0)"
       szRealTicketPrice = ", isnull(SUM(ticket_price),0)"
       szTable = "stat_bus_startoff_lst2 "
       szReturnCharge = ""
       szTotalIncom = ""
    End If
    
    
    oDB.ConnectionString = GetConnectionStr(cszDss)
    
    szGroupBy = " (CASE WHEN substring(bus_id, 1, 1) = '+' THEN 1 ELSE 0 END) "
    
    If pszCompanyID <> "" Then 'FGAG
        If InStr(pszCompanyID, "'") <> 1 Then pszCompanyID = "'" + pszCompanyID + "'"
        szWhere = " AND b.split_company_id IN ( " & pszCompanyID & ")"
    End If
    
    '如果要查询的售票站不为空
    If pszSellStationID <> "" Then
        szWhere = szWhere & " AND b.sell_station_id IN (" & pszSellStationID & ")"
    End If
    
    '如果不是所有的状态
    If pnExtraStatus <> TP_AllSold Then
        szWhere = szWhere & " AND extra_status = " & pnExtraStatus
    End If
    
    '车次代码 多车次以逗号隔开
    If pszVehicleID <> "" Then
        szWhere = szWhere & " AND b.vehicle_id  in (" & pszVehicleID & ") "
    End If
    
    szWhere = " AND date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "'" & szWhere
    
    szSql = "SELECT v.vehicle_id,rtrim(v.license_tag_no) license_tag_no," & "rtrim(o.owner_name) owner_name" _
            & " , MAX(B.route_id) route_id , rtrim(MAX(route_name)) route_name " _
            & " , MAX(bus_id) bus_id, MAX(transport_company_short_name) transport_company_short_name " _
            & GetSumItem _
            & szRealTicketPrice & " AS ticket_price, " _
            & szRealTicketCount & " AS passenger_number " _
            & szReturnCharge _
            & szTotalIncom _
            & " FROM " & szTable

    szSql = szSql & " b , route_info r , company_info c , vehicle_info v , owner_info o " _
            & " WHERE b.vehicle_id = v.vehicle_id AND r.route_id = b.route_id " _
            & " AND b.transport_company_id = c.transport_company_id AND v.owner_id = o.owner_id " _
            & szWhere

            
    szSql = szSql & " GROUP BY v.license_tag_no,v.vehicle_id,o.owner_name," & szGroupBy & " ORDER BY v.license_tag_no,v.vehicle_id,o.owner_name, " & szGroupBy
    
    Set rsTemp = oDB.Execute(szSql)
    

    Set GetVehicleStat = rsTemp
  
    
End Function

'分别得到多条记录的各个票价项之和
Public Function GetSumItem() As String
    Dim i As Integer
    Dim szTemp As String
        
    '得到字符串
    szTemp = ""
        
    For i = 0 To 15
        If i = 0 Then
            szTemp = szTemp & ", isnull(sum(base_price),0) as base_price"
        Else
            szTemp = szTemp & ",isnull(sum(price_item_" & i & "),0) as price_item_" & i
        End If
    Next i
    
    GetSumItem = szTemp
End Function
