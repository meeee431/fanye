VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PriceSheetReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 Option Explicit


'各类特有的变量

'所有的实体类都有的私有类变量<<
Private m_nObjectStatus As EObjectStatus

'所有的类都有的私有类变量
Private m_oActiveUser As ActiveUser
'--------------------------->>
Public Sub Init(poAUser As ActiveUser)
    Set SelfUser = poAUser

End Sub

Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
    m_oActiveUser.IAmLive
End Property

Public Property Set SelfUser(vNewValue As ActiveUser)
    Set m_oActiveUser = vNewValue
    m_oActiveUser.IAmLive
End Property

Public Sub AssertObjIsValid() '测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
    AssertActiveUserValid m_oActiveUser, ERR_RouteTicketPrice
    AssertStatusAvailable
End Sub
Private Sub AssertStatusAvailable() '测试对象的状态是否有效（不为无效ECNotAvailable状态）
'    If m_nObjectStatus = ST_NotAvailable Then ShowError ERR_NotAvailable
End Sub


Public Function GetEnviromentPriceItemByBus(QueryDate As Date, BusID() As String, Optional TicketType As Integer = 0) As Recordset
    AssertObjIsValid
    Dim szSql As String
    Dim odb As New RTConnection
    Dim lRowCount As Long
    Dim rsTmp As Recordset
    Dim szCondition As String, szFields As String
    Dim nArrayLength As Integer, nTmp As Integer
    Dim i As Integer, j As Integer
    Dim szWhere As String
    Dim szFrom As String
    Dim szSelect As String
    Dim rsTemp As Recordset
    Dim szTable As String
    
    Dim szTemp As String
    
    '形成返回的线路票价项字段
    szSql = "SELECT * FROM price_item_info  "
    odb.ConnectionString = GetConnectionStr(cszPriceMan)
    Set rsTmp = odb.Execute(szSql)
    If rsTmp.RecordCount = 0 Then
        GoTo ExitPos
    End If
    
    For i = 1 To rsTmp.RecordCount
        If rsTmp!use_mark <> 0 Then
            nTmp = Val(rsTmp.Fields!price_item)
            If nTmp = 0 Then        '基本运价
                szFields = szFields & ",ept.base_carriage "
            Else
                szFields = szFields & ",ept.price_item_" & Trim(Str(nTmp))
            End If
        End If
        rsTmp.MoveNext
    Next i
        
    Set rsTemp = odb.Execute("SELECT * FROM Ticket_type_code WHERE ticket_type_valid=1 AND ticket_type_id<>'" & TP_FreeTicket & "' AND ticket_type_id<>'" & TP_FullPrice & "'")
    
    Do While Not rsTemp.EOF
        szTable = rsTemp!ticket_type_id
        szWhere = szWhere & " AND ept" & szTable & ".seat_type_id=ept.seat_type_id  AND ept" & szTable & ".bus_date=ept.bus_date  AND ept" & szTable & ".ticket_type='" & szTable & "'" & _
            " AND ept" & szTable & ".bus_id=ept.bus_id  AND ept" & szTable & ".station_serial_no=ept.station_serial_no AND ept" & szTable & ".station_id=ept.station_id " & _
            " AND ept" & szTable & ".sell_station_id=ept.sell_station_id "
        szFrom = szFrom & "Env_bus_price_lst ept" & szTable & ","
        
        If szTable = 2 Then
          szTemp = ",ept" & szTable & ".ticket_price_total as total_half_amount "
        ElseIf szTable = 4 Then
          szTemp = ",ept" & szTable & ".ticket_price_total as total_preferential1_amount "
        ElseIf szTable = 5 Then
          szTemp = ",ept" & szTable & ".ticket_price_total as total_preferential2_amount "
        ElseIf szTable = 6 Then
          szTemp = ",ept" & szTable & ".ticket_price_total as total_preferential3_amount "
        End If

        szSelect = szSelect & szTemp
        rsTemp.MoveNext
    Loop
        
    szFields = "ebi.total_seat, ebi.end_station_name end_station_name,ebi.check_gate_id check_gate,ebi.sale_seat_quantity sale_seat_quantity,case ebi.status  when 1 then " & TransFieldValueToString("停开") & " else " & TransFieldValueToString("开") & " end as status,ebi.check_gate_id check_gate,convert(char(5),ebi.bus_start_time,108) bus_start_time,ept.sell_station_id,ssi.sell_station_name,ri.route_name route,ept.bus_id bus_id,si.station_name station,ebi.vehicle_type_name vehicle_type, ebs.mileage mileage,sfi.seat_type_name " & szFields & ",ept.ticket_price_total as total_full_amount" & szSelect
       
    szCondition = " ebi.bus_date='" & Format(QueryDate, "YYYY-MM-DD") & "'"
    '组成所有车次集合条件
    nArrayLength = ArrayLength(BusID)
    If nArrayLength > 0 Then
        szCondition = szCondition & " AND ebi.bus_id IN ("
    End If
    For i = 1 To nArrayLength - 1
        szCondition = szCondition & "'" & BusID(i) & "',"
    Next i
    If i = nArrayLength Then
        szCondition = szCondition & "'" & BusID(nArrayLength) & "')  "
    End If
    
    Dim szWorkTable As String
    If DateDiff("d", QueryDate, Now) <= 0 Then
          szWorkTable = "Work_env_bus_price_lst ept,Work_env_bus_station_lst ebs,Work_env_bus_info ebi,sell_station_info ssi,"
    Else
          szWorkTable = "Env_bus_price_lst ept,Env_bus_station_lst ebs,Env_bus_info ebi,sell_station_info ssi,"
    End If
     '形成完整SQL串
    szSql = " SELECT " & szFields & _
            " FROM " & szWorkTable & szFrom & " route_info ri,station_info si,Seat_type_code sfi" & _
            " WHERE " & szCondition & " AND ept.ticket_type=1" & _
            " AND ebs.bus_id=ebi.bus_id AND ebs.bus_date=ebi.bus_date AND ebs.sell_station_id = ept.sell_station_id  " & szWhere & _
            " AND ssi.sell_station_id=ept.sell_station_id" & _
            " AND ept.bus_id=ebs.bus_id AND ept.bus_date=ebs.bus_date AND ept.station_id=ebs.station_id" & _
            " AND ri.route_id=ebi.route_id AND ept.station_id=si.station_id" & _
            " AND ept.seat_type_id=ebs.seat_type_id AND sfi.seat_type_id=ept.seat_type_id" & _
            " ORDER BY ept.sell_station_id,ept.bus_id,ebi.vehicle_type_code,ept.station_serial_no,ept.station_id,ept.seat_type_id"

    Set rsTmp = odb.Execute(szSql)
ExitPos:
    Set GetEnviromentPriceItemByBus = rsTmp
End Function

Public Function GetEnviromentPriceItemByRoute(QueryDate As Date, RouteId() As String, Optional TicketType As Integer = 0) As Recordset
    AssertObjIsValid
    Dim szSql As String, odb As New RTConnection
    Dim rsTmp As Recordset
    Dim szCondition As String, szFields As String
    Dim nArrayLength As Integer, nTmp As Integer
    Dim i As Integer
    
    '形成返回的线路票价项字段
    szSql = "SELECT price_item FROM price_item_info WHERE use_mark=1"
    odb.ConnectionString = GetConnectionStr(cszPriceMan)
    Set rsTmp = odb.Execute(szSql)
    If rsTmp.RecordCount = 0 Then
        GoTo ExitPos
    End If
    For i = 1 To rsTmp.RecordCount
        nTmp = Val(rsTmp.Fields!price_item)
        If nTmp = 0 Then        '基本运价
            szFields = szFields & ",ept.base_carriage pi_" & Trim(Str(i))
        Else
            szFields = szFields & ",ept.price_item_" & Trim(Str(nTmp)) & " pi_" & Trim(Str(i))
        End If
        rsTmp.MoveNext
    Next i
    '形成返回的完整字段列
    szFields = "ept.sell_station_id,ssi.sell_station_name,ri.route_name route,ept.bus_id bus_id,si.station_name station,ebi.vehicle_type_name vehicle_type, ept.ticket_type ticket_type,ebs.mileage mileage" & szFields
    
    szCondition = " ebi.bus_date ='" & Format(QueryDate, "YYYY-MM-DD") & "'"
    '组成所有车次集合条件
    nArrayLength = ArrayLength(RouteId)
    If nArrayLength > 0 Then
        szCondition = szCondition & " AND ebi.route_id IN ("
    End If
    For i = 1 To nArrayLength - 1
        szCondition = szCondition & "'" & RouteId(i) & "',"
    Next i
    If i = nArrayLength Then
        szCondition = szCondition & "'" & RouteId(nArrayLength) & "')  "
    End If
    
    Select Case TicketType
        Case TP_FullPrice      '全票价
            szCondition = szCondition & " AND ept.ticket_type=1"
        Case TP_HalfPrice      '半票价
            szCondition = szCondition & " AND ept.ticket_type=2"
        Case TP_PreferentialTicket1        '优惠票价1
            szCondition = szCondition & " AND ept.ticket_type=4"
        Case TP_PreferentialTicket2       '优惠票价2
            szCondition = szCondition & " AND ept.ticket_type=5"
        Case TP_PreferentialTicket3       '优惠票价3
            szCondition = szCondition & " AND ept.ticket_type=6"
    End Select

    Dim szWorkTable As String
    If DateDiff("d", QueryDate, Now) <= 0 Then
          szWorkTable = "Work_env_bus_price_lst ept,Work_env_bus_station_lst ebs,Work_env_bus_info ebi,sell_station_info ssi,"
    Else
          szWorkTable = "Env_bus_price_lst ept,Env_bus_station_lst ebs,Env_bus_info ebi,sell_station_info ssi,"
    End If
    '形成完整SQL串
    szSql = " SELECT " & szFields & _
            " FROM " & szWorkTable & "route_info ri,station_info si" & _
            " WHERE " & szCondition & _
            " AND ebs.bus_id=ebi.bus_id AND ebs.bus_date=ebi.bus_date" & _
            " AND ssi.sell_station_id=ept.sell_station_id" & _
            " AND ept.bus_id=ebs.bus_id AND ept.bus_date=ebs.bus_date AND ept.station_id=ebs.station_id" & _
            " AND ri.route_id=ebi.route_id AND ept.station_id=si.station_id" & _
            " ORDER BY ba.sell_station_id,ept.bus_id,ebi.vehicle_type_code,ept.station_serial_no,ept.station_id,ept.seat_type_id"

    Set rsTmp = odb.Execute(szSql)
ExitPos:
    Set GetEnviromentPriceItemByRoute = rsTmp
End Function

'车次总票价报表
Public Function GetBusRoutePriceItem(PriceTable As String, Project As String, BusID() As String, SeatType() As String _
            , SellStation() As String, Optional RunStation As Boolean = True) As Recordset
    
    'RunStation=True 停靠站点 RunStation=False 所有站点
    AssertObjIsValid
    Dim szSql As String, odb As New RTConnection, lRowCount As Long
    Dim rsTmp As Recordset
    Dim szCondition As String, szFields As String
    Dim nArrayLength As Integer, nTmp As Integer
    Dim i As Integer, j As Integer
    Dim szWhere As String
    Dim szFrom As String
    Dim szSelect As String
    Dim rsTemp As Recordset
    Dim szTable As String
    Dim szTotalFullPriceField As String, szTotalHalfPriceField As String
       
    '形成返回的票价项字段
    szSql = "SELECT * FROM price_item_info  "
    odb.ConnectionString = GetConnectionStr(cszPriceMan)
    Set rsTmp = odb.Execute(szSql)
    If rsTmp.RecordCount = 0 Then
        GoTo ExitPos
    End If

    For i = 1 To rsTmp.RecordCount
        nTmp = Val(rsTmp.Fields!price_item)
        If rsTmp!use_mark <> 0 Then
            If nTmp = 0 Then        '基本运价
                szFields = szFields & ",bpl.base_carriage "
                szTotalFullPriceField = szTotalFullPriceField & " + bpl.base_carriage"
            Else
                szFields = szFields & ",bpl.price_item_" & Trim(Str(nTmp))
                szTotalFullPriceField = szTotalFullPriceField & " + bpl.price_item_" & Trim(Str(nTmp))
            End If
        End If
        rsTmp.MoveNext
    Next i
    If Len(szTotalFullPriceField) > 0 Then
       szTotalFullPriceField = ",(" & Right(szTotalFullPriceField, Len(szTotalFullPriceField) - 3) & ") total_full_amount "
    End If
    
    '得到可使用票种
    Set rsTemp = odb.Execute("SELECT * FROM Ticket_type_code WHERE ticket_type_valid=1 AND ticket_type_id<>'" & TP_FreeTicket & "' AND ticket_type_id<>'" & TP_FullPrice & "'")
    
    Do While Not rsTemp.EOF
        szTotalHalfPriceField = ""
        szTable = rsTemp!ticket_type_id
        szWhere = szWhere & "AND bpl" & szTable & ".bus_id=bpl.bus_id  AND bpl" & szTable & ".seat_type_id=bpl.seat_type_id " & " AND bpl" & szTable & ".price_table_id=bpl.price_table_id  AND bpl" & szTable & ".ticket_type='" & szTable & "'" & _
            "  AND bpl" & szTable & ".vehicle_type_code=bpl.vehicle_type_code AND bpl" & szTable & ".station_id=bpl.station_id  AND bpl" & szTable & ".sell_station_id = bpl.sell_station_id  "
            
        szFrom = szFrom & "bus_price_lst bpl" & szTable & ", "
        
        rsTmp.MoveFirst
        For i = 1 To rsTmp.RecordCount
            nTmp = Val(rsTmp.Fields!price_item)
            If rsTmp!use_mark <> 0 Then
                If nTmp = 0 Then        '基本运价
                    szTotalHalfPriceField = szTotalHalfPriceField & " + bpl" & szTable & ".base_carriage"
                Else '其他票价项
                   szTotalHalfPriceField = szTotalHalfPriceField & " + bpl" & szTable & ".price_item_" & Trim(Str(nTmp))
                End If
            End If
            rsTmp.MoveNext
        Next i
         
        If Len(szTotalFullPriceField) > 0 Then
           szTotalHalfPriceField = ",(" & Right(szTotalHalfPriceField, Len(szTotalHalfPriceField) - 3) & ") "
            If szTable = 2 Then
              szTotalHalfPriceField = szTotalHalfPriceField & " total_half_amount "
            ElseIf szTable = 4 Then
              szTotalHalfPriceField = szTotalHalfPriceField & " total_preferential1_amount "
            ElseIf szTable = 5 Then
              szTotalHalfPriceField = szTotalHalfPriceField & " total_preferential2_amount "
            ElseIf szTable = 6 Then
              szTotalHalfPriceField = szTotalHalfPriceField & " total_preferential3_amount "
            End If
        End If
        szSelect = szSelect & szTotalHalfPriceField
        rsTemp.MoveNext
    Loop
        
    szFields = " convert(char(5),bfo.bus_start_time,108) bus_start_time,bpl.sell_station_id,ssi.sell_station_name,bpl.bus_id,bpl.station_id,bpl.vehicle_type_code,bpl.mileage,vtc.vehicle_type_short_name vehicle_type,sti.seat_type_name,sfo.station_name  " & szFields & szTotalFullPriceField & szSelect & " ,bpl.mileage mileage "
    
    szCondition = " AND bpl.price_table_id='" & PriceTable & "' "

    nArrayLength = ArrayLength(BusID)       '添加车次查询条件
    If nArrayLength > 0 Then
        szCondition = szCondition & " AND bpl.bus_id IN ("
    End If
    For i = 1 To nArrayLength - 1
        szCondition = szCondition & "'" & BusID(i) & "',"
    Next i
    If i = nArrayLength Then
        szCondition = szCondition & "'" & BusID(i) & "') "
    End If

    nArrayLength = ArrayLength(SeatType)       '添加座位类型查询条件
    If nArrayLength > 0 Then
        szCondition = szCondition & " AND bpl.seat_type_id IN ("
    End If
    For i = 1 To nArrayLength - 1
        szCondition = szCondition & "'" & SeatType(i) & "',"
    Next i
    If i = nArrayLength Then
        szCondition = szCondition & "'" & SeatType(nArrayLength) & "') "
    End If
    
    nArrayLength = ArrayLength(SellStation)        '添加座位类型查询条件
    If nArrayLength > 0 Then
        szCondition = szCondition & " AND bpl.sell_station_id IN ("
    End If
    For i = 1 To nArrayLength - 1
        szCondition = szCondition & "'" & SellStation(i) & "',"
    Next i
    If i = nArrayLength Then
        szCondition = szCondition & "'" & SellStation(nArrayLength) & "') "
    End If
    
    
    
    If RunStation = False Then
        szSql = " SELECT " & szFields & _
                " FROM bus_price_lst bpl,bus_info bfo,sell_station_info ssi," & szFrom & _
                " station_info sfo,vehicle_type_code vtc,Seat_type_code sti " & _
                " WHERE bfo.bus_id=bpl.bus_id " & szCondition & "AND bpl.ticket_type=1  " & szWhere & _
                " AND ssi.sell_station_id=bpl.sell_station_id" & _
                " AND vtc.vehicle_type_code=bpl.vehicle_type_code AND sfo.station_id=bpl.station_id " & _
                " AND sti.seat_type_id=bpl.seat_type_id " & _
                " ORDER BY bfo.bus_id,bpl.sell_station_id,bpl.vehicle_type_code,bpl.station_serial_no,bpl.mileage,bpl.station_id ,bpl.seat_type_id "
    Else
        szSql = " SELECT " & szFields & _
                " FROM bus_price_lst bpl,bus_info bfo,sell_station_info ssi," & szFrom & _
                " station_info sfo,vehicle_type_code vtc,Seat_type_code sti,Bus_station_info gbs " & _
                " WHERE bfo.bus_id=bpl.bus_id " & szCondition & "AND bpl.ticket_type=1  " & szWhere & _
                " AND ssi.sell_station_id=bpl.sell_station_id" & _
                " AND vtc.vehicle_type_code=bpl.vehicle_type_code AND sfo.station_id=bpl.station_id " & _
                " AND sti.seat_type_id=bpl.seat_type_id and gbs.bus_id=bpl.bus_id and gbs.station_id=bpl.station_id " & _
                " and  max_sale_quantity<>0 " & _
                " ORDER BY bfo.bus_id,bpl.sell_station_id,bpl.vehicle_type_code,bpl.station_serial_no,bpl.mileage,bpl.station_id ,bpl.seat_type_id "
    End If
    Set rsTmp = odb.Execute(szSql)
ExitPos:
    Set GetBusRoutePriceItem = rsTmp
End Function

'线路车型站点票价报表
Public Function GetRouteStationVehiclePriceRpt(PriceTable As String, Station() As String, VehicleType() As String, SeatType() As String) As Recordset
    AssertObjIsValid
    Dim szSql As String, odb As New RTConnection, lRowCount As Long
    Dim rsTmp As Recordset
    Dim szCondition As String, szFields As String
    Dim nArrayLength As Integer, nTmp As Integer
    Dim i As Integer, j As Integer
    Dim szWhere As String
    Dim szFrom As String
    Dim szSelect As String
    Dim rsTemp As Recordset
    Dim szTable As String
    
    '形成返回的线路票价项字段
    szSql = "SELECT * FROM price_item_info  "
    odb.ConnectionString = GetConnectionStr(cszPriceMan)
    Set rsTmp = odb.Execute(szSql)
    If rsTmp.RecordCount = 0 Then
        GoTo ExitPos
    End If
    
    Dim szTotalFullPriceField As String, szTotalHalfPriceField As String
    
    For i = 1 To rsTmp.RecordCount
        nTmp = Val(rsTmp.Fields!price_item)
        If rsTmp!use_mark <> 0 Then
            If nTmp = 0 Then        '基本运价
                szFields = szFields & ",pt.base_carriage "
                szTotalFullPriceField = szTotalFullPriceField & " + pt.base_carriage"
            Else
                If nTmp < 13 Then      '线路票价项
                    szFields = szFields & ",pt.price_item_" & Trim(Str(nTmp))
                    szTotalFullPriceField = szTotalFullPriceField & " + pt.price_item_" & Trim(Str(nTmp))
                End If
            End If
        End If
        rsTmp.MoveNext
    Next i
    If Len(szTotalFullPriceField) > 0 Then
       szTotalFullPriceField = ",(" & Right(szTotalFullPriceField, Len(szTotalFullPriceField) - 3) & ") total_full_amount "
    End If
    
    '得到可使用票种
    Set rsTemp = odb.Execute("SELECT * FROM Ticket_type_code WHERE ticket_type_valid=1 AND ticket_type_id<>'" & TP_FreeTicket & "' AND ticket_type_id<>'" & TP_FullPrice & "'")
    
    Do While Not rsTemp.EOF
        szTotalHalfPriceField = ""
        szTable = rsTemp!ticket_type_id
        szWhere = szWhere & " AND pt" & szTable & ".seat_type_id=pt.seat_type_id " & " AND pt" & szTable & ".price_table_id=pt.price_table_id  AND pt" & szTable & ".ticket_type='" & szTable & "'" & _
            " AND pt" & szTable & ".route_id=pt.route_id  AND pt" & szTable & ".vehicle_type_code=pt.vehicle_type_code AND pt" & szTable & ".station_id=pt.station_id "
            
        szFrom = szFrom & "price_table_lst pt" & szTable & ","
        
        rsTmp.MoveFirst
        For i = 1 To rsTmp.RecordCount
            nTmp = Val(rsTmp.Fields!price_item)
            If rsTmp!use_mark <> 0 Then
                If nTmp = 0 Then        '基本运价
                    szTotalHalfPriceField = szTotalHalfPriceField & " + pt" & szTable & ".base_carriage"
                Else
                    If nTmp < 13 Then      '线路票价项
                        szTotalHalfPriceField = szTotalHalfPriceField & " + pt" & szTable & ".price_item_" & Trim(Str(nTmp))
                    End If
                End If
            End If
            rsTmp.MoveNext
        Next i
         
        If Len(szTotalFullPriceField) > 0 Then
           szTotalHalfPriceField = ",(" & Right(szTotalHalfPriceField, Len(szTotalHalfPriceField) - 3) & ") "
            If szTable = 2 Then
              szTotalHalfPriceField = szTotalHalfPriceField & " total_half_amount "
            ElseIf szTable = 4 Then
              szTotalHalfPriceField = szTotalHalfPriceField & " total_preferential1_amount "
            ElseIf szTable = 5 Then
              szTotalHalfPriceField = szTotalHalfPriceField & " total_preferential2_amount "
            ElseIf szTable = 6 Then
              szTotalHalfPriceField = szTotalHalfPriceField & " total_preferential3_amount "
            End If
        End If
        szSelect = szSelect & szTotalHalfPriceField
        rsTemp.MoveNext
    Loop
        
    szFields = "pt.sell_station_id,ssi.sell_station_name,si.station_name station,vt.vehicle_type_short_name vehicle_type,ri.route_name route,pt.seat_type_id,sti.Seat_type_name  " & szFields & szTotalFullPriceField & szSelect & " ,pt.mileage mileage"
    
    szCondition = " AND pt.price_table_id='" & PriceTable & "'"
    nArrayLength = ArrayLength(Station)       '添加站点查询条件
    If nArrayLength > 0 Then
        szCondition = szCondition & " AND pt.station_id IN ("
    End If
    For i = 1 To nArrayLength - 1
        szCondition = szCondition & "'" & Station(i) & "',"
    Next i
    If i = nArrayLength Then
        szCondition = szCondition & "'" & Station(nArrayLength) & "') "
    End If

    nArrayLength = ArrayLength(VehicleType)       '添加车型查询条件
    If nArrayLength > 0 Then
        szCondition = szCondition & " AND pt.vehicle_type_code IN ("
    End If
    For i = 1 To nArrayLength - 1
        szCondition = szCondition & "'" & VehicleType(i) & "',"
    Next i
    If i = nArrayLength Then
        szCondition = szCondition & "'" & VehicleType(nArrayLength) & "') "
    End If
    
    nArrayLength = ArrayLength(SeatType)       '添加座位类型查询条件
    If nArrayLength > 0 Then
        szCondition = szCondition & " AND pt.seat_type_id IN ("
    End If
    For i = 1 To nArrayLength - 1
        szCondition = szCondition & "'" & SeatType(i) & "',"
    Next i
    If i = nArrayLength Then
        szCondition = szCondition & "'" & SeatType(nArrayLength) & "') "
    End If
    
    szSql = " SELECT " & szFields & _
            " FROM price_table_lst pt,sell_station_info ssi," & szFrom & _
            " vehicle_type_code vt,station_info si,route_info ri,Seat_type_code sti" & _
            " WHERE 1=1" & szCondition & " AND pt.ticket_type=1 " & szWhere & _
            " AND pt.sell_station_id=ssi.sell_station_id" & _
            " AND si.station_id=pt.station_id AND vt.vehicle_type_code=pt.vehicle_type_code AND ri.route_id=pt.route_id AND pt.seat_type_id=sti.seat_type_id" & _
            " ORDER BY pt.sell_station_id,pt.mileage ,pt.station_id,pt.vehicle_type_code,pt.route_id,pt.seat_type_id"

    Set rsTmp = odb.Execute(szSql)
ExitPos:
    Set GetRouteStationVehiclePriceRpt = rsTmp
End Function

'站点车型票价报表
Public Function GetStationVehiclePriceRpt(PriceTable As String, Station() As String, VehicleType() As String, SeatType() As String, SellStation() As String) As Recordset
    AssertObjIsValid
    Dim szSql As String, odb As New RTConnection, lRowCount As Long
    Dim rsTmp As Recordset
    Dim szCondition As String, szFields As String
    Dim nArrayLength As Integer, nTmp As Integer
    Dim i As Integer, j As Integer
    Dim szWhere As String
    Dim szFrom As String
    Dim szSelect As String
    Dim rsTemp As Recordset
    Dim szTable As String
    
    '形成返回的线路票价项字段
    szSql = "SELECT * FROM price_item_info  "
    odb.ConnectionString = GetConnectionStr(cszPriceMan)
    Set rsTmp = odb.Execute(szSql)
    If rsTmp.RecordCount = 0 Then
        GoTo ExitPos
    End If
    
    Dim szTotalFullPriceField As String, szTotalHalfPriceField As String
    
    For i = 1 To rsTmp.RecordCount
        nTmp = Val(rsTmp.Fields!price_item)
        If rsTmp!use_mark <> 0 Then
            If nTmp = 0 Then        '基本运价
                szFields = szFields & ",bpl.base_carriage "
                szTotalFullPriceField = szTotalFullPriceField & " + bpl.base_carriage"
            Else
                '线路票价项
                szFields = szFields & ",bpl.price_item_" & Trim(Str(nTmp))
                szTotalFullPriceField = szTotalFullPriceField & " + bpl.price_item_" & Trim(Str(nTmp))
            End If
        End If
        rsTmp.MoveNext
    Next i
    If Len(szTotalFullPriceField) > 0 Then
       szTotalFullPriceField = ",(" & Right(szTotalFullPriceField, Len(szTotalFullPriceField) - 3) & ") total_full_amount "
    End If
    
    '得到可使用票种
    Set rsTemp = odb.Execute("SELECT * FROM Ticket_type_code WHERE ticket_type_valid=1 AND ticket_type_id<>'" & TP_FreeTicket & "' AND ticket_type_id<>'" & TP_FullPrice & "'")
    
    Do While Not rsTemp.EOF
        szTotalHalfPriceField = ""
        szTable = rsTemp!ticket_type_id
        szWhere = szWhere & " AND bpl" & szTable & ".seat_type_id=bpl.seat_type_id " & " AND bpl" & szTable & ".price_table_id=bpl.price_table_id  AND bpl" & szTable & ".ticket_type='" & szTable & "'" & _
            " AND bpl" & szTable & ".bus_id=bpl.bus_id  AND bpl" & szTable & ".vehicle_type_code=bpl.vehicle_type_code AND bpl" & szTable & ".station_id=bpl.station_id " & " AND bpl" & szTable & ".sell_station_id=bpl.sell_station_id "
            
        szFrom = szFrom & "bus_price_lst bpl" & szTable & ","
        
        rsTmp.MoveFirst
        For i = 1 To rsTmp.RecordCount
            nTmp = Val(rsTmp.Fields!price_item)
            If rsTmp!use_mark <> 0 Then
                If nTmp = 0 Then        '基本运价
                    szTotalHalfPriceField = szTotalHalfPriceField & " + bpl" & szTable & ".base_carriage"
                Else
                    '线路票价项
                    szTotalHalfPriceField = szTotalHalfPriceField & " + bpl" & szTable & ".price_item_" & Trim(Str(nTmp))
                End If
            End If
            rsTmp.MoveNext
        Next i
         
        If Len(szTotalFullPriceField) > 0 Then
           szTotalHalfPriceField = ",(" & Right(szTotalHalfPriceField, Len(szTotalHalfPriceField) - 3) & ") "
            If szTable = 2 Then
              szTotalHalfPriceField = szTotalHalfPriceField & " total_half_amount "
            ElseIf szTable = 4 Then
              szTotalHalfPriceField = szTotalHalfPriceField & " total_preferential1_amount "
            ElseIf szTable = 5 Then
              szTotalHalfPriceField = szTotalHalfPriceField & " total_preferential2_amount "
            ElseIf szTable = 6 Then
              szTotalHalfPriceField = szTotalHalfPriceField & " total_preferential3_amount "
            End If
        End If
        szSelect = szSelect & szTotalHalfPriceField
        rsTemp.MoveNext
    Loop
        
    szFields = "DISTINCT bpl.sell_station_id,ssi.sell_station_name,bpl.station_id,si.station_name station,bpl.vehicle_type_code,vt.vehicle_type_short_name vehicle_type," & _
         " bpl.seat_type_id,sti.Seat_type_name  " & szFields & szTotalFullPriceField & szSelect & " ,bpl.mileage mileage"
    
    szCondition = " AND bpl.price_table_id='" & PriceTable & "'"
    nArrayLength = ArrayLength(Station)       '添加站点查询条件
    If nArrayLength > 0 Then
        szCondition = szCondition & " AND bpl.station_id IN ("
    End If
    For i = 1 To nArrayLength - 1
        szCondition = szCondition & "'" & Station(i) & "',"
    Next i
    If i = nArrayLength Then
        szCondition = szCondition & "'" & Station(nArrayLength) & "') "
    End If

    nArrayLength = ArrayLength(VehicleType)       '添加车型查询条件
    If nArrayLength > 0 Then
        szCondition = szCondition & " AND bpl.vehicle_type_code IN ("
    End If
    For i = 1 To nArrayLength - 1
        szCondition = szCondition & "'" & VehicleType(i) & "',"
    Next i
    If i = nArrayLength Then
        szCondition = szCondition & "'" & VehicleType(nArrayLength) & "') "
    End If
    
    nArrayLength = ArrayLength(SeatType)       '添加座位类型查询条件
    If nArrayLength > 0 Then
        szCondition = szCondition & " AND bpl.seat_type_id IN ("
    End If
    For i = 1 To nArrayLength - 1
        szCondition = szCondition & "'" & SeatType(i) & "',"
    Next i
    If i = nArrayLength Then
        szCondition = szCondition & "'" & SeatType(nArrayLength) & "') "
    End If
    
    nArrayLength = ArrayLength(SellStation)        '添加上车站查询条件
    If nArrayLength > 0 Then
        szCondition = szCondition & " AND bpl.sell_station_id IN ("
    End If
    For i = 1 To nArrayLength - 1
        szCondition = szCondition & "'" & SellStation(i) & "',"
    Next i
    If i = nArrayLength Then
        szCondition = szCondition & "'" & SellStation(nArrayLength) & "') "
    End If
    szSql = " SELECT " & szFields & _
            " FROM bus_price_lst bpl,sell_station_info ssi," & szFrom & _
            " vehicle_type_code vt,station_info si,Seat_type_code sti" & _
            " WHERE 1=1" & szCondition & " AND bpl.ticket_type=1 " & szWhere & _
            " AND bpl.sell_station_id=ssi.sell_station_id" & _
            " AND si.station_id=bpl.station_id AND vt.vehicle_type_code=bpl.vehicle_type_code AND bpl.seat_type_id=sti.seat_type_id" & _
            " ORDER BY bpl.station_id,bpl.sell_station_id,bpl.mileage ,si.station_name,bpl.vehicle_type_code,bpl.seat_type_id"

    Set rsTmp = odb.Execute(szSql)
ExitPos:
    Set GetStationVehiclePriceRpt = rsTmp
End Function
