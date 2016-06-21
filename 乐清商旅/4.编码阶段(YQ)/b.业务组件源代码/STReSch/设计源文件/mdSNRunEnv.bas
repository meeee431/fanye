Attribute VB_Name = "mdSNRunenv"
Option Explicit


Public Const cszRunEnv = ""
Public Const cbAllRefundment = "1"
Public Const cbNotAllRefundment = "0"



'断言指定的车次无售票
Public Function BusHaveNotSellTicket(poDb As RTConnection, pdtBusDate As Date, pszBusID As String) As Boolean
    Dim szSql As String
    Dim rsTemp As Recordset
    szSql = "SELECT COUNT(*) AS countx FROM Ticket_sell_lst WHERE " _
    & " bus_id='" & pszBusID & "' AND " _
    & " bus_date='" & ToDBDate(pdtBusDate) & "'"
    Set rsTemp = poDb.Execute(szSql)
   BusHaveNotSellTicket = False
   If rsTemp.RecordCount > 0 Then
     If FormatDbValue(rsTemp!countx) > 0 Then
        BusHaveNotSellTicket = True
     End If
   End If
End Function

'内部用得到总票价
Public Function SelfGetTotalPrice(prsPriceInfo As Recordset) As Double
    Dim sgTemp As Double
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!base_carriage)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_1)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_2)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_3)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_4)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_5)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_6)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_7)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_8)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_9)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_10)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_11)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_12)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_13)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_14)
    sgTemp = sgTemp + FormatDbValue(prsPriceInfo!price_item_15)
    SelfGetTotalPrice = sgTemp
End Function

'内部用根据座位号（整型）生成真正的座位号（字符串型）
Public Function BuildSeatNo(pnSeat As Integer) As String
    BuildSeatNo = Format(pnSeat, "00")
End Function


'内部用根据座位号（字符串型）解析成座位号（整型）
Public Function ResolveSeatNo(pszSeat As String) As Integer
    ResolveSeatNo = CInt(pszSeat)
End Function

Public Function FindRightStation(prsStation As Recordset, ByVal pszStationID As String) As Boolean
  prsStation.MoveFirst
  On Error GoTo here
  Do While Not prsStation.EOF
        If Trim(FormatDbValue(prsStation!station_id)) = Trim(pszStationID) Then
            FindRightStation = True
            Exit Function
        End If
        prsStation.MoveNext
    Loop
here:
    FindRightStation = False
    Exit Function
End Function
'/////////////////////////////////////
'得到车次车辆座位类型信息
Public Function GetBusVehicleSeatType(pszVehicleID As String) As TVehcileSeatType()
Dim oDb As New RTConnection
Dim szSql As String
Dim rsTemp As Recordset
Dim i As Integer
Dim tvTemp() As TVehcileSeatType
oDb.ConnectionString = GetConnectionStr(cszRunEnv)
szSql = "SELECT * FROM vehicle_seat_type_info WHERE vehicle_id='" & pszVehicleID & "' ORDER BY seat_type_id,start_seat_no"
Set rsTemp = oDb.Execute(szSql)
If rsTemp.RecordCount <> 0 Then
    ReDim tvTemp(1 To rsTemp.RecordCount)
    For i = 1 To rsTemp.RecordCount
        tvTemp(i).szSeatTypeID = rsTemp!seat_type_id
        tvTemp(i).szStartSeatNo = rsTemp!start_seat_no
        tvTemp(i).szEndSeatNo = rsTemp!end_seat_no
        rsTemp.MoveNext
    Next i
    GetBusVehicleSeatType = tvTemp
    Set rsTemp = Nothing
End If
End Function

'//////////////////////////////////
'取座位类型
Public Function GetSeatTypeID(pszSeatNo As Integer, patvVehicleType() As TVehcileSeatType) As String
    Dim i As Integer
    Dim nLen As Integer
    nLen = ArrayLength(patvVehicleType)
    For i = 1 To nLen
        If pszSeatNo >= patvVehicleType(i).szStartSeatNo And pszSeatNo <= patvVehicleType(i).szEndSeatNo Then
            GetSeatTypeID = patvVehicleType(i).szSeatTypeID
            Exit Function
        End If
    Next i
End Function
'************************
' adde
'***********************
Public Function FormatReTicketToStationTable(tTReTicketPrice() As TRETicketPriceEx) As TBusStationSellInfo()
Dim i As Integer
Dim nCount As Integer
Dim tTReStationTicket() As TBusStationSellInfo
Dim nCountTemp As Integer
Dim j As Integer
nCount = ArrayLength(tTReTicketPrice)
ReDim tTReStationTicket(0 To 0)
For i = 1 To nCount
    'j = 0
    nCountTemp = ArrayLength(tTReStationTicket)
    Do While tTReStationTicket(j).szSeatTypeID <> tTReTicketPrice(i).szSeatType Or tTReStationTicket(j).szSellStationID <> tTReTicketPrice(i).szSellStationID
        j = j + 1
         If j >= nCountTemp Then
           ReDim Preserve tTReStationTicket(0 To nCountTemp)
           Exit Do
         End If
   Loop
   Select Case tTReTicketPrice(i).nTicketType
                   Case TP_FullPrice
                       tTReStationTicket(j).sgFullPrice = tTReTicketPrice(i).sgTotal
                   Case TP_HalfPrice
                       tTReStationTicket(j).sgHalfPrice = tTReTicketPrice(i).sgTotal
                   Case TP_PreferentialTicket1
                       tTReStationTicket(j).sgPreferentialPrice1 = tTReTicketPrice(i).sgTotal
                   Case TP_PreferentialTicket2
                       tTReStationTicket(j).sgPreferentialPrice2 = tTReTicketPrice(i).sgTotal
                   Case TP_PreferentialTicket3
                       tTReStationTicket(j).sgPreferentialPrice3 = tTReTicketPrice(i).sgTotal
    End Select
    tTReStationTicket(j).szSeatTypeID = tTReTicketPrice(i).szSeatType
    tTReStationTicket(j).nMileage = tTReTicketPrice(i).sgMileage
    tTReStationTicket(j).sgBasePrice = tTReTicketPrice(i).sgBase
    tTReStationTicket(j).szStationID = tTReTicketPrice(i).szStationID
    tTReStationTicket(j).szSellStationID = tTReTicketPrice(i).szSellStationID
Next
FormatReTicketToStationTable = tTReStationTicket
End Function
'*******************************************************************************
'***
'得到车次车辆座位类型信息,假如车辆座位表总座位小于车辆座位则加入《普通座位》类型
'并且《普通座位》类型 : 起始座位=起始车辆座位 , 结束座位=车辆总座位
Public Function GetBusVehicleSeatTypeEx(pszVehicleID As String) As TVehcileSeatType()
Dim oDb As New RTConnection
Dim szSql As String
Dim rsTemp As Recordset
Dim bflg As Boolean
Dim nCount As Integer
Dim i As Integer
Dim tvTemp() As TVehcileSeatType
oDb.ConnectionString = GetConnectionStr(cszRunEnv)
szSql = "SELECT v.*,t.seat_quantity ,t.start_seat_no AS startseat,t.vehicle_type_code FROM vehicle_seat_type_info v,Vehicle_info  t WHERE v.vehicle_id='" & pszVehicleID & "' and v.vehicle_id=t.vehicle_id ORDER BY v.seat_type_id,v.start_seat_no"
Set rsTemp = oDb.Execute(szSql)

If rsTemp.RecordCount <> 0 Then
    ReDim tvTemp(1 To rsTemp.RecordCount)
    For i = 1 To rsTemp.RecordCount
        If Trim(rsTemp!seat_type_id) = cszSeatTypeIsNormal And bflg = False Then bflg = True
        tvTemp(i).szSeatTypeID = rsTemp!seat_type_id
        tvTemp(i).szStartSeatNo = rsTemp!start_seat_no
        tvTemp(i).szEndSeatNo = rsTemp!end_seat_no
        tvTemp(i).szVehcileID = pszVehicleID
         tvTemp(i).szVehcileTypeName = rsTemp!vehicle_type_code
        If rsTemp!start_seat_no <= rsTemp!end_seat_no Then
           nCount = nCount + CInt(rsTemp!end_seat_no) - CInt(rsTemp!start_seat_no) + 1
        Else
           nCount = nCount + CInt(rsTemp!start_seat_no) - CInt(rsTemp!end_seat_no) + 1
        End If
          rsTemp.MoveNext
    Next i
    rsTemp.MoveFirst
    If nCount < rsTemp!seat_quantity And bflg = False Then
     ReDim Preserve tvTemp(1 To rsTemp.RecordCount + 1)
        tvTemp(i).szVehcileID = pszVehicleID
        tvTemp(i).szSeatTypeID = cszSeatTypeIsNormal
        tvTemp(i).szStartSeatNo = rsTemp!startseat
        tvTemp(i).szEndSeatNo = rsTemp!seat_quantity
        tvTemp(i).szVehcileTypeName = rsTemp!vehicle_type_code
    End If
    GetBusVehicleSeatTypeEx = tvTemp
    Set rsTemp = Nothing
Else
      szSql = "SELECT t.seat_quantity ,t.start_seat_no ,t.vehicle_type_code FROM Vehicle_info  t WHERE t.vehicle_id='" & pszVehicleID & "'  "
      Set rsTemp = oDb.Execute(szSql)
        ReDim tvTemp(1 To 1)
        tvTemp(1).szSeatTypeID = cszSeatTypeIsNormal
        tvTemp(1).szStartSeatNo = rsTemp!start_seat_no
        tvTemp(1).szEndSeatNo = rsTemp!seat_quantity
        tvTemp(1).szVehcileTypeName = rsTemp!vehicle_type_code
        tvTemp(1).szVehcileID = pszVehicleID
        GetBusVehicleSeatTypeEx = tvTemp
    Set rsTemp = Nothing
End If
End Function
'*************************************************
'确认座位类型有效
'在取的票价时确认
Public Function AssertSeatTypeIsValidSeatType(tvTemp() As TVehcileSeatType, szAssertSeatType As String) As Boolean
Dim nCount As Integer, i As Integer
Dim bflg As Boolean
nCount = ArrayLength(tvTemp)
For i = 1 To nCount
   If Trim(tvTemp(i).szSeatTypeID) = szAssertSeatType Then
      bflg = True
      Exit For
   End If
Next
    If bflg = False Then
       AssertSeatTypeIsValidSeatType = False
    Else
      AssertSeatTypeIsValidSeatType = True
    End If
End Function
Public Function UpdateEnviromentSeatCountEx(szBusID As String, dtDate As Date, Optional poDb As RTConnection = Nothing) As Integer

Dim oDb As New RTConnection
Dim szSql As String
Dim rsTemp As New Recordset
Dim tTSeatInfo(1 To 5) As TSeatInfoCount
Dim i As Integer
Dim nSaleSeatQuantity As Integer
Dim szSqlSet As String
Set oDb = poDb

szSql = " select seat_type_id ,count(*)as CountSeat from Env_bus_seat_lst " _
        & " where  bus_id='" & szBusID & "'and  status='" & ST_SeatCanSell & "' and bus_date='" & ToDBDate(dtDate) & "' group by seat_type_id "
Set rsTemp = oDb.Execute(szSql)

For i = 1 To rsTemp.RecordCount
    tTSeatInfo(CInt(rsTemp!seat_type_id)).seatCount = FormatDbValue(rsTemp!CountSeat)
    rsTemp.MoveNext
Next

szSqlSet = szSqlSet & "seat_remain ='" & tTSeatInfo(1).seatCount & "',"
szSqlSet = szSqlSet & "bed_remain ='" & tTSeatInfo(2).seatCount & "',"
szSqlSet = szSqlSet & "additional_remain ='" & tTSeatInfo(3).seatCount & "',"
szSqlSet = szSqlSet & "other_remain_1 ='" & tTSeatInfo(4).seatCount & "',"
szSqlSet = szSqlSet & "other_remain_2 ='" & tTSeatInfo(5).seatCount & "',"

For i = 1 To 5
  nSaleSeatQuantity = nSaleSeatQuantity + tTSeatInfo(i).seatCount
Next

szSqlSet = szSqlSet & "sale_seat_quantity='" & nSaleSeatQuantity & "'"
szSql = "UPDATE Env_bus_info SET  " & szSqlSet _
        & "  WHERE bus_id='" & szBusID & "' AND  bus_date='" & ToDBDate(dtDate) & "'"
oDb.Execute (szSql)
UpdateEnviromentSeatCountEx = nSaleSeatQuantity
End Function

Public Function UpdateEnvSeatInfo(szVehicleID As String, _
                                  nNewToalSeatCount As Integer, nNewStartSeatNo As Integer, _
                                  szBusID As String, busdate As Date, _
                                  poDb As RTConnection, _
                                  Optional bflg As Boolean = True)
Dim i As Integer
Dim nCount As Integer
Dim nCountTemp As Integer
Dim szSql As String
Dim rsTempStart As New Recordset
Dim rsTempEnd As New Recordset
Dim tvTemp() As TVehcileSeatType '新车辆座位设置情况
Dim nNewEndSeatCount As Integer '新车辆结束座位
Dim nOldEndSeatCount As Integer
Dim bSaleStart As Boolean
Dim bSaleEnd As Boolean
Dim nOldToalSeatCount As Integer, nOldStartSeatNo As Integer

On Error GoTo here
poDb.BeginTrans
'取得起始座号
szSql = "Select min(seat_no) as SeatNo from Env_bus_seat_lst where bus_id='" & szBusID & "' " _
              & " and bus_date='" & ToDBDate(busdate) & "' "
Set rsTempStart = poDb.Execute(szSql)
nOldStartSeatNo = CInt(rsTempStart!SeatNo)
'取得总座位数
szSql = "Select Count(*) as seatCount from Env_bus_seat_lst where bus_id='" & szBusID & "' " _
              & " and bus_date='" & ToDBDate(busdate) & "' "
Set rsTempStart = poDb.Execute(szSql)
nOldToalSeatCount = CInt(rsTempStart!seatCount)

nNewEndSeatCount = CInt(nNewStartSeatNo) + nNewToalSeatCount - 1
nOldEndSeatCount = nOldToalSeatCount + CInt(nOldStartSeatNo) - 1

'起始座号不一样处理
If nOldStartSeatNo > nNewStartSeatNo Then
    nCountTemp = nOldStartSeatNo - nNewStartSeatNo
   
    '插入座位test
    For i = nNewStartSeatNo To nCountTemp
          szSql = "INSERT Env_bus_seat_lst( " _
                                & "bus_date," _
                                & "bus_id," _
                                & "seat_no," _
                                & "status," _
                                & "ticket_no," _
                                & "seat_type_id) "
                 szSql = szSql & " VALUES('" _
                                & ToDBDate(busdate) & "','" _
                                & szBusID & "','" _
                                & BuildSeatNo(i) & "'," _
                                & ST_SeatCanSell & ",'','" _
                                & cszSeatTypeIsNormal & "')"
                 poDb.Execute szSql
    Next
 End If
 If nOldStartSeatNo < nNewStartSeatNo Then
     '保存售票信息  test
     szSql = " select * from Env_bus_seat_lst where bus_id='" & szBusID & "' " _
              & " and bus_date='" & ToDBDate(busdate) & "' " _
              & " and ticket_no<>'' and seat_no<'" & BuildSeatNo(nNewStartSeatNo) & "' order by Seat_no"
     'poDb.CommitTrans
     Set rsTempStart = poDb.Execute(szSql)
     bSaleStart = True
     '删除座位
     szSql = " delete Env_bus_seat_lst where bus_id='" & szBusID & "' " _
              & " and bus_date='" & ToDBDate(busdate) & "' " _
              & " and seat_no<'" & BuildSeatNo(nNewStartSeatNo) & "'"
     poDb.Execute szSql
End If
'结束座位不一样处理
If nNewEndSeatCount > nOldEndSeatCount Then
   nCountTemp = nNewEndSeatCount - nOldEndSeatCount
   'test
   For i = 1 To nCountTemp
          szSql = "INSERT Env_bus_seat_lst( " _
                                & "bus_date," _
                                & "bus_id," _
                                & "seat_no," _
                                & "status," _
                                & "ticket_no," _
                                & "seat_type_id) "
                 szSql = szSql & " VALUES('" _
                                & ToDBDate(busdate) & "','" _
                                & szBusID & "','" _
                                & BuildSeatNo(i + nOldEndSeatCount) & "'," _
                                & ST_SeatCanSell & ",'','" _
                                & cszSeatTypeIsNormal & "')"
                 poDb.Execute szSql
    Next
End If
If nNewEndSeatCount < nOldEndSeatCount Then
     '保存售票信息 test
     szSql = " select * from Env_bus_seat_lst where bus_id='" & szBusID & "' " _
              & " and bus_date='" & ToDBDate(busdate) & "' " _
              & " and ticket_no<>'' and seat_no>'" & BuildSeatNo(nNewEndSeatCount) & "'order by Seat_no "
     Set rsTempEnd = poDb.Execute(szSql)
     bSaleEnd = True
     '删除座位
     szSql = " delete Env_bus_seat_lst where bus_id='" & szBusID & "' " _
              & " and bus_date='" & ToDBDate(busdate) & "' " _
              & " and seat_no>'" & BuildSeatNo(nNewEndSeatCount) & "'"
     poDb.Execute szSql
End If
'将新车辆座位类型更新到数据库
'（1）更新座位为普通座位

szSql = "   update Env_bus_seat_lst Set  seat_type_id='" & cszSeatTypeIsNormal & "'" _
            & " where bus_id='" & szBusID & "' and bus_date='" & ToDBDate(busdate) & "'"
poDb.Execute szSql
'(2) 新车辆座位类型更新
'取得新车辆座位设置情况
tvTemp = GetBusVehicleSeatType(szVehicleID)
nCount = ArrayLength(tvTemp)
For i = 1 To nCount
    szSql = "   update Env_bus_seat_lst Set seat_type_id='" & tvTemp(i).szSeatTypeID & "'" _
            & " where ( seat_no between " _
            & " '" & Format(tvTemp(i).szStartSeatNo, "00") & "'" _
            & " and  '" & Format(tvTemp(i).szEndSeatNo, "00") & " '" _
            & " ) " _
            & " and bus_id='" & szBusID & "'  " _
            & " and bus_date='" & ToDBDate(busdate) & " '"
    poDb.Execute szSql
Next
'已售座位恢复
Dim lRow As Long
If bSaleEnd = True Then
    If rsTempEnd.RecordCount <> 0 Then
      For i = 1 To rsTempEnd.RecordCount
        
        szSql = " update Env_bus_seat_lst Set ticket_no='" & rsTempEnd!ticket_no & "',status='" & rsTempEnd!Status & "'" _
                  & " where   bus_date='" & ToDBDate(busdate) & "'and bus_id='" & szBusID & "' and seat_no =(" _
                  & " select min(seat_no) from Env_bus_seat_lst where bus_id='" & szBusID & "' " _
                  & " and  bus_date='" & ToDBDate(busdate) & "'and status='" & ST_SeatCanSell & "' " _
                  & " )"
        poDb.Execute szSql, lRow
        
        If lRow <= 0 Then
          '可能覆盖预留，预定座位
          szSql = "   update enviroment_seat_type Set ticket_no='" & rsTempEnd!ticket_no & "',status='" & rsTempEnd!Status & "'" _
                & " where bus_date='" & ToDBDate(busdate) & "'and bus_id='" & szBusID & "' " _
                & " and seat_no =(" _
                & " select max(seat_no) from Env_bus_seat_lst where bus_id='" & szBusID & "' " _
                & " and  bus_date='" & ToDBDate(busdate) & "' " _
                & " and ticket_no=''" _
                & ") "
         poDb.Execute szSql, lRow
        End If
        rsTempEnd.MoveNext
      Next
    End If
End If

If bSaleStart = True Then
    If rsTempStart.RecordCount <> 0 Then
      For i = 1 To rsTempStart.RecordCount
        szSql = "   update Env_bus_seat_lst Set ticket_no='" & rsTempStart!ticket_no & "',status='" & rsTempStart!Status & "'" _
                & " where bus_date='" & ToDBDate(busdate) & "'and bus_id='" & szBusID & "'" _
                & " and seat_no =(" _
                & " select min(seat_no) from Env_bus_seat_lst where bus_id='" & szBusID & "' " _
                & " and  bus_date='" & ToDBDate(busdate) & "'and status='" & ST_SeatCanSell & "' " _
                & ")"
                 
        poDb.Execute szSql, lRow
        If lRow <= 0 Then
          '可能覆盖预留，预定座位
          szSql = "   update Env_bus_seat_lst Set ticket_no='" & rsTempStart!ticket_no & "',status='" & rsTempStart!Status & "'" _
                & " where bus_date='" & ToDBDate(busdate) & "'and bus_id='" & szBusID & "' " _
                & " and seat_no =(" _
                & " select max(seat_no) from Env_bus_seat_lst where bus_id='" & szBusID & "' " _
                & " and  bus_date='" & ToDBDate(busdate) & "' " _
                & " and ticket_no='' " _
                & ")"
         poDb.Execute szSql, lRow
        End If
        rsTempStart.MoveNext
      Next
    End If
End If


'状态更新---将已售售位 更新为 此座位已被售出，顶班得到
'拆分得到座位状态不变
If bflg = True Then
    
    szSql = "   update Env_bus_seat_lst Set status ='" & ST_SeatReplace & "'" _
                & " where bus_id='" & szBusID & "' and bus_date='" & ToDBDate(busdate) & "'  and ticket_no<>''" _
                & " and status<> '" & ST_SeatSlitp & "'"
    poDb.Execute szSql

End If

poDb.CommitTrans
Set rsTempStart = Nothing
Set rsTempEnd = Nothing

Exit Function
here:
    poDb.CommitTrans
    err.Raise err.numer
End Function

