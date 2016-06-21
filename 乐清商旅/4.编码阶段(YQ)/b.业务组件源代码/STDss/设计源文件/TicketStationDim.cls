VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TicketStationDim"
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


'站点汇总
Public Function StationCount(pdtFromDate As Date, pdtToDate As Date, pszAreaCode As String, Optional pszSellStationID As String = "") As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_StationCount
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    oDB.ConnectionString = GetConnectionStr(cszDss)
        If pszAreaCode <> "" Then
            szWhere = " AND si.area_code IN ( " & pszAreaCode & ") "
        End If
           '如果要查询的售票站不为空
    If pszSellStationID <> "" Then
        Dim szSellStationID() As String
        Dim i As Integer
        szSellStationID = Split(pszSellStationID, ",")
        szWhere = szWhere & " AND (s.sell_station_id='" & szSellStationID(0) & "' "
        For i = 1 To ArrayLength(szSellStationID) - 1
            szWhere = szWhere & " OR s.sell_station_id='" & szSellStationID(i) & "' "
        Next
        szWhere = szWhere & ")"
    End If
    
        szSql = "   SELECT  s.station_id, MAX(s.station_name) station_name , MAX(station_input_code) station_input_code " _
            & "    , SUM(bASe_price) bASe_price , SUM (price_item_1) price_item_1, SUM(price_item_2) price_item_2   " _
            & "    , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 , SUM(price_item_5) price_item_5  " _
            & "    , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 ,SUM(price_item_8) price_item_8  " _
            & "    , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 " _
            & "    , SUM(price_item_12) price_item_12 , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 " _
            & "    , SUM(price_item_15) price_item_15 " _
            & "    , (SUM(pASsenger_number) -SUM(ticket_cancel_number)-SUM(ticket_return_number)-SUM(ticket_change_number)) AS pASsenger_number " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) ) AS ticket_price " _
            & "    , (SUM(ticket_price) - SUM(ticket_cancel_amount) - SUM(ticket_return_amount) " _
            & "         - SUM(ticket_change_amount) + SUM(ticket_return_charge)  ) AS total_ticket_price  " _
            & "    , SUM(ticket_cancel_number) ticket_cancel_number , SUM(ticket_cancel_amount) ticket_cancel_amount " _
            & "    , SUM(ticket_return_number) ticket_return_number , SUM(ticket_return_amount) ticket_return_amount " _
            & "    , SUM(ticket_return_charge) ticket_return_charge , SUM(ticket_change_number) ticket_change_number " _
            & "    , SUM(ticket_change_amount) ticket_change_amount , SUM(ticket_change_charge) ticket_change_charge " _
            & "    , SUM(fact_float_number) fact_float_number " _
            & " FROM  Stat_station_sell_lst s , station_info si" _
            & " WHERE date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "' " _
            & " AND si.station_id = s.station_id " _
            & szWhere _
            & " GROUP BY s.station_id ORDER BY station_input_code "
        
    Set rsTemp = oDB.Execute(szSql)
    Set StationCount = rsTemp
    
End Function




'得到参数指定的站点的人数统计
Public Function GetStationByParam(pdtFromDate As Date, pdtToDate As Date) As String()
    Dim oParam As New SystemParam
    Dim szStationDetail As String '需统计的站点明细
    Dim aszStationDetail() As String
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim aszTemp() As String
    Dim i As Integer
    Dim nCount As Integer
    Dim szWhere As String
    
    AssertActiveUserValid m_oActiveUser, ERR_SellerFinance
    AssertHaveRight m_oActiveUser, RIGHT_GetStationByParam
    
    
    oDB.ConnectionString = GetConnectionStr()
    oParam.Init m_oActiveUser
    szStationDetail = oParam.TranNeedStationDetail
    aszStationDetail = StringToTeam(szStationDetail)
    nCount = ArrayLength(aszStationDetail)
    
    For i = 1 To nCount - 1
        szWhere = szWhere & "'" & aszStationDetail(i) & "',"
    Next i
    If i = nCount Then
        szWhere = " AND station_id IN (" & szWhere & "'" & aszStationDetail(i) & "')"
    End If
    szSql = " SELECT MAX(station_id) station_id , MAX(station_name) station_name  , (SUM(passenger_number) -SUM(ticket_cancel_number) - SUM(ticket_return_number)) AS passenger_number " _
            & " FROM  Stat_station_sell_lst s " _
            & " WHERE date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "' " _
            & szWhere _
            & " GROUP BY s.station_id "
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 3)
        
        For i = 1 To rsTemp.RecordCount
            aszTemp(i, 1) = FormatDbValue(rsTemp!station_id)
            aszTemp(i, 2) = FormatDbValue(rsTemp!station_name)
            aszTemp(i, 3) = FormatDbValue(rsTemp!passenger_number)
            rsTemp.MoveNext
        Next i
    End If
    GetStationByParam = aszTemp
    
    
    
End Function


