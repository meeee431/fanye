VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**********************************************************
'* Source File Name:Bus.cls
'* Project Name:STRglSch.vbp
'* Engineer:
'* Data Generated:
'* Last Revision Date:
'* Brief Description:车次
'* Relational Document:
'**********************************************************

Option Explicit

'出错
Public Enum EErrBus
    ERR_BusAdd = ERR_Bus + ERR_AddDuplicate '14401该计划中已存在该车次，不能新增
    ERR_BusAddNotParent = ERR_Bus + ERR_AddNoParent '14402车次的其他属性不完整
    ERR_BusNotAddStatus = ERR_Bus + ERR_NotAddObj '14416车次对象不处于新增状态
    ERR_BusAddStatus = ERR_Bus + ERR_AddObj '内部错误，车次处于新增状态 ！
    
    ERR_BusNotExist = ERR_Bus + 21 '14421车次不存在
    ERR_BusNotPassStation = ERR_Bus + 22 '14422车次无过站信息
    ERR_BusNotRunVehicle = ERR_Bus + 23 '14423车次无运行车辆
    ERR_VehicleNotReserverSeat = ERR_Bus + 24 '14424车次车辆无预留
    ERR_BusNotBusPrice = ERR_Bus + 25 '14425车次没有车次票价表
    ERR_BusSerialNotVehicle = ERR_Bus + 26 '14426车辆序号无对应车辆
    ERR_ProjectIDNotNull = ERR_Bus + 27 '14427计划代码不能为空
    ERR_BusIDNotNull = ERR_Bus + 28 '14428车次代码不能为空
    ERR_BusExistNotAddNew = ERR_Bus + 29 '14429车次不存在，不能新增
    ERR_RouteNotExist = ERR_Bus + 30 '14430线路不存在
    ERR_BusVehicleSerialIsExist = ERR_Bus + 31 '14431车辆序号已存在，不能新增
    ERR_BusVehicleSerialNotZero = ERR_Bus + 32 '14432车辆序号不能为零
    ERR_VehicleNotExist = ERR_Bus + 33 '14433车辆不存在
    ERR_BusNotExistVehicleSerial = ERR_Bus + 34 '14434该车次无该车辆序号
    ERR_RouteNoCanSellStation = ERR_Bus + 35 '14435线路无可售票站点
    ERR_BusSerialExist = ERR_Bus + 36 '14436新的车次序号已存在不能修改
    ERR_StationNotInBusRoute = ERR_Bus + 37 '14437 '指定的站点不在车次的线路上
    ERR_NoPriceExcuseTable = ERR_Bus + 38 '无执行票价表,请添加票价表
    ERR_NoStationInfo = ERR_Bus + 39 '无对应站点信息
    ERR_NoSpecifyRoute = ERR_Bus + 40 '无线路信息
    ERR_NoBusPriceTable = ERR_Bus + 41 '无线路票价表
    ERR_NoCheckGate = ERR_Bus + 42 '无检票口信息
    ERR_NoBusTypeInfo = ERR_Bus + 43 '无车次种类信息
    ERR_NoSellStation = ERR_Bus + 44 '该售票站不存在
    ERR_CheckGateNotMaskStation = ERR_Bus + 45 '检票口不存在，或与所属车站不符合
    
    ERR_BusAllotAreadyExist = ERR_Bus + 45 '该配载信息已存在
    ERR_BusAllotNotExist = ERR_Bus + 46 '该配载信息不存在
End Enum

'权限
Public Enum ERightBus
    RIGHT_BusManagement = ERR_Bus + cnMidRightBegin + cnMidRightStep * 1 '14507车次管理的权限
    RIGHT_BusDelete = ERR_Bus + cnMidRightBegin + cnMidRightStep * 2 '14513删除车次
    RIGHT_BusAdd = ERR_Bus + cnMidRightBegin + cnMidRightStep * 3 '14519新增车次
    RIGHT_BusUpdate = ERR_Bus + cnMidRightBegin + cnMidRightStep * 4 '14525修改车次基本属性
    RIGHT_BusVehicleManagement = ERR_Bus + cnMidRightBegin + cnMidRightStep * 7 '14543编辑车次车辆信息
    RIGHT_BusStationManagement = ERR_Bus + cnMidRightBegin + cnMidRightStep * 9 '14555编辑站点信息
End Enum

Private m_oActiveUser As ActiveUser 'P1所有的类都有的私有类变量
Private m_nObjectStatus As EObjectStatus 'P2所有的实体类都有的私有类变量
'Private m_szProjectID As String 'P3车次计划代码
Private m_szBusID As String 'P4车次代码
Private m_szRouteID As String 'P5运行线路
Private m_szCheckGate As String 'P6车次检票口
Private m_dtStartupTime As Date 'P7车次发车时间
Private m_nRunCycle As Integer 'P8车次得运行周期
Private m_nCycleStartSerialNo As Integer 'P9开始运行序号
Private m_dtBeginStopDate As Date 'P10车次停班开始日期
Private m_dtEndStopDate As Date 'P11车次结束停班日期
Private m_nBusType As Integer 'P12车次类型固定车次、流水车次
Private m_szRouteName As String   'P16线路名称


Private m_szEndStationID As String    '
Private m_szEndStationName As String  '
Private m_nInternetStatus As Integer '互联
Private m_nScrollBusCheckTime As Integer '滚动车检票时间

Private m_nOldInternetStatus As Integer
Private m_szOldRoute As String '老的线路代码
'日志
Private m_szOldCheckGate As String 'P6车次检票口
Private m_dtOldStartupTime As Date 'P7车次发车时间
Private m_nOldRunCycle As Integer 'P8车次得运行周期
Private m_nOldCycleStartSerialNo As Integer 'P9开始运行序号
Private m_nOldBusType As Integer 'P12车次类型固定车次、流水车次
Private m_nOldScrollBusCheckTime As Integer '滚动车检票时间

Private m_dtOldBeginStopDate As Date 'P10车次停班开始日期
Private m_dtOldEndStopDate As Date 'P11车次结束停班日期



Private m_szStartStationID As String '起点站代码
Private m_szStartStationName As String '起点站名称

'Public Type TBusAddTicketPriceInfo
'    szStationID As String '站点代码
'    nTicketType As ETicketType '票种
'    sgPrice1 As Double '车次附加票价项1
'    sgPrice2 As Double '车次附加票价项2
'    sgPrice3 As Double '车次附加票价项3
'End Type



'**************************************************
'Member Code:F1
'Brief Description:获得车次的过站信息
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Public Function GetPassStationSellInfo(StationSerial As Integer) As TBusStationSellInfo
'**************************************************
'StationSerial (站点序号)
'**************************************************
Dim oDb As New RTConnection
Dim rsTemp As Recordset
Dim BSInfo As TBusStationSellInfo
Dim i As Integer
Dim szTempSql As String
AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szTempSql = "SELECT * " _
                & " FROM Bus_station_info " _
                & " WHERE bus_id='" & m_szBusID & "' AND station_serial_no=" & Str(StationSerial) & " ORDER BY mileage"
    Set rsTemp = oDb.Execute(szTempSql)
    '-------------------
    '车次无过站
    If rsTemp.RecordCount <> 0 Then
        BSInfo.szStationID = FormatDbValue(rsTemp!station_id)
        BSInfo.nMileage = FormatDbValue(rsTemp!mileage)
        BSInfo.sgLimitedSellTime = FormatDbValue(rsTemp!stop_sale_time)
        BSInfo.nLimitedSellCount = FormatDbValue(rsTemp!max_sale_quantity)
        GetPassStationSellInfo = BSInfo
    End If
End Function

''**************************************************
''Member Code:F2
''Brief Description:获得车次的全部运行车辆
''Engineer:
''Date Generated:
''Last Revision Date:
''**************************************************
Public Function GetAllVehicleEx() As TBusVehicleInfo()
Dim oDb As New RTConnection
Dim rsTemp As Recordset
Dim BVInfo() As TBusVehicleInfo
Dim i As Integer
Dim szTempSql As String

AssertObjIsValid

    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szTempSql = "SELECT * " _
                & " FROM bus_vehicle_code " _
                & " WHERE bus_id='" & m_szBusID & "'"
    Set rsTemp = oDb.Execute(szTempSql)
    '车次无运行车辆
    If rsTemp.RecordCount <> 0 Then

    ReDim BVInfo(1 To rsTemp.RecordCount) As TBusVehicleInfo
    For i = 1 To rsTemp.RecordCount
        BVInfo(i).szVehicleID = FormatDbValue(rsTemp!vehicle_id)
        BVInfo(i).nSerialNo = FormatDbValue(rsTemp!vehicle_serial)
        BVInfo(i).nStandTicketCount = FormatDbValue(rsTemp!sale_stand_ticket_quantity)
        BVInfo(i).dtBeginStopDate = FormatDbValue(rsTemp!stop_start_date)
        BVInfo(i).dtEndStopDate = FormatDbValue(rsTemp!stop_end_date)
        rsTemp.MoveNext
    Next
    GetAllVehicleEx = BVInfo
    End If
End Function

'**************************************************
'Member Code:F3
'Brief Description:获得车次的所有过站信息
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
'在新增车次时调用
Public Function GetAllPassStation(Optional VehicleModel As String) As TBusStationSellInfo()
'    Dim oDb As New RTConnection
'    Dim szTempSql As String
'    Dim atTicketType() As TTicketType
'    Dim rsTempStation As Recordset
'    Dim rsBusTP As Recordset
'    Dim szSql As String
'    Dim rsTicketType As Recordset
'    Dim szSeatType() As String
'    Dim BSInfo() As TBusStationSellInfo
'    Dim nCount As Integer
'    Dim sgPrice As Single
'    Dim i As Integer, j As Integer, h As Integer
'    Dim nSeatTypeCount As String
'    Dim rsTemp As Recordset
'    Dim nCountTemp As Integer
'    Dim lErrCode As Long
'
'    Dim szPriceExcuteTable As String
'    AssertObjIsValid
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'    '得到站点
'    szTempSql = "SELECT distinct gbs.*,sif.station_name FROM Bus_station_info gbs ,station_info sif WHERE gbs.bus_id='" & m_szBusID & "' and gbs.station_id=sif.station_id AND project_id='" & m_szProjectID & "' ORDER BY  station_serial_no"
'    Set rsTempStation = oDb.Execute(szTempSql)
'     '得到执行票价表
'    szPriceExcuteTable = BusProjectExecutePrice(m_szProjectID, Now, lErrCode)
'     '得到票价表
'     If szPriceExcuteTable = "" Then ShowError ERR_NoPriceExcuseTable
'     If VehicleModel <> "" Then
'        szTempSql = "SELECT * FROM bus_price_lst WHERE bus_id='" & m_szBusID & "'  AND vehicle_type_code='" & VehicleModel & "' AND price_table_id='" & szPriceExcuteTable & "' ORDER BY station_serial_no"
'     Else
'        szTempSql = "SELECT * FROM bus_price_lst WHERE bus_id='" & m_szBusID & "'  AND price_table_id='" & szPriceExcuteTable & "' ORDER BY station_serial_no"
'     End If
'     Set rsBusTP = oDb.Execute(szTempSql)
'     If rsBusTP.RecordCount = 0 Then ShowError ERR_BusNotBusPrice
'     '得到票种
'     szTempSql = " SELECT DISTINCT ticket_type_id FROM Ticket_type_code where ticket_type_valid='" & TP_TicketTypeValid & "' and ticket_type_name<>'免票'"
'     Set rsTicketType = oDb.Execute(szTempSql)
'     If rsTicketType.RecordCount = 0 Then Exit Function
'     On Error GoTo Here
'     If VehicleModel <> "" Then
'     szTempSql = "Select DISTINCT bpl.station_serial_no,bpl.bus_id,bpl.seat_type_id,bpl.station_id ,sti.seat_type_name FROM bus_price_lst bpl ,Seat_type_code sti WHERE sti.seat_type_id=bpl.seat_type_id and bpl.bus_id='" & m_szBusID & "'  AND bpl.price_table_id='" & szPriceExcuteTable & "'AND bpl.vehicle_type_code='" & VehicleModel & "' ORDER BY  bpl.station_serial_no,bpl.seat_type_id"
'     Else
'     szTempSql = "Select DISTINCT bpl.station_serial_no,bpl.bus_id,bpl.seat_type_id,bpl.station_id ,sti.seat_type_name FROM bus_price_lst bpl,Seat_type_code sti WHERE  sti.seat_type_id=bpl.seat_type_id,bpl.bus_id='" & m_szBusID & "'  AND bpl.price_table_id='" & szPriceExcuteTable & "' ORDER BY bpl.station_serial_no,bpl.seat_type_id"
'     End If
'     Set rsTemp = oDb.Execute(szTempSql)
'     If rsTemp.RecordCount = 0 Then Exit Function
'     ReDim BSInfo(1 To rsTemp.RecordCount) As TBusStationSellInfo
'        For i = 1 To rsTemp.RecordCount
'             FindRightStatonItem rsTempStation, rsTemp!station_id
'
'
'             BSInfo(i).szStationID = FormatDbValue(rsTempStation!station_id)
'             BSInfo(i).nMileage = FormatDbValue(rsTempStation!mileage)
'             BSInfo(i).sgLimitedSellTime = FormatDbValue(rsTempStation!stop_sale_time)
'             BSInfo(i).nLimitedSellCount = FormatDbValue(rsTempStation!max_sale_quantity)
'             BSInfo(i).szSeatTypeName = rsTemp!seat_type_Name
'             BSInfo(i).szStationName = rsTempStation!station_name
'             BSInfo(i).szSeatTypeID = rsTemp!seat_type_id
'             rsTicketType.MoveFirst
'             For j = 1 To rsTicketType.RecordCount
'
'                 If VehicleModel = "" Then
'                      FindRightPriceItem rsBusTP, BSInfo(i).szStationID, rsTicketType!ticket_type_id, rsTemp!seat_type_id, , True
'                 Else
'                      FindRightPriceItem rsBusTP, BSInfo(i).szStationID, rsTicketType!ticket_type_id, rsTemp!seat_type_id, VehicleModel, True
'                 End If
'
'                 sgPrice = SelfGetTotalPrice(rsBusTP)
'
'                 Select Case rsTicketType!ticket_type_id
'                           Case TP_FullPrice
'                               BSInfo(i).sgFullPrice = sgPrice
'                           Case TP_HalfPrice
'                               BSInfo(i).sgHalfPrice = sgPrice
'                           Case TP_PreferentialTicket1
'                               BSInfo(i).sgPreferentialPrice1 = sgPrice
'                           Case TP_PreferentialTicket2
'                               BSInfo(i).sgPreferentialPrice2 = sgPrice
'                           Case TP_PreferentialTicket3
'                               BSInfo(i).sgPreferentialPrice3 = sgPrice
'                  End Select
'
'                 rsTicketType.MoveNext
'           Next
'         rsTemp.MoveNext
'      Next
'
'
''代票价的结构
'GetAllPassStation = BSInfo
'Exit Function
'Here:
'    If err.Number = 3021 Then
'    ShowError ERR_StationNotInBusRoute
'
'    End If
End Function

Public Function GetAllStation(Optional VehicleModel As String) As TBusStationSellInfo()
Dim oDb As New RTConnection
Dim szTempSql As String
Dim atTicketType() As TTicketType
Dim rsTempStation As Recordset
Dim rsBusTP As Recordset
Dim szSql As String
Dim rsTicketType As Recordset
Dim szSeatType() As String
Dim BSInfo() As TBusStationSellInfo
Dim nCount As Integer
Dim sgPrice As Single
Dim i As Integer, j As Integer, h As Integer
Dim nSeatTypeCount As String
Dim rsTemp As Recordset
Dim nCountTemp As Integer
Dim lErrCode As Long
Dim szPriceExcuteTable As String
AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    '得到站点
    szTempSql = "SELECT distinct gbs.*,sif.station_name " _
                & " FROM Bus_station_info gbs ,station_info sif " _
                & " WHERE gbs.bus_id='" & m_szBusID & "' and gbs.station_id=sif.station_id  " _
                & " ORDER BY  station_serial_no"
    Set rsTempStation = oDb.Execute(szTempSql)
     '得到执行票价表
    szPriceExcuteTable = BusProjectExecutePrice(Now, lErrCode)  'FL m_szProjectID
     '得到票价表
     If szPriceExcuteTable = "" Then ShowError ERR_NoPriceExcuseTable
     If VehicleModel <> "" Then
        szTempSql = "SELECT * FROM bus_price_lst " _
            & " WHERE bus_id='" & m_szBusID & "'  AND vehicle_type_code='" & VehicleModel & "' AND price_table_id='" & szPriceExcuteTable & "' " _
            & " ORDER BY sell_station_id , station_serial_no"
     Else
        szTempSql = "SELECT * FROM bus_price_lst WHERE bus_id='" & m_szBusID & "'  AND price_table_id='" & szPriceExcuteTable & "' ORDER BY sell_station_id , station_serial_no"
     End If
     Set rsBusTP = oDb.Execute(szTempSql)
     If rsBusTP.RecordCount = 0 Then ShowError ERR_BusNotBusPrice
     '得到票种
     szTempSql = " SELECT DISTINCT ticket_type_id FROM Ticket_type_code WHERE ticket_type_valid='" & TP_TicketTypeValid & "' and ticket_type_name<>'免票'"
     Set rsTicketType = oDb.Execute(szTempSql)
     If rsTicketType.RecordCount = 0 Then Exit Function
     On Error GoTo Here
     If VehicleModel <> "" Then '增加了sell_station_id,sell_station_name
     szTempSql = "Select DISTINCT bpl.sell_station_id,s.sell_station_name,bpl.station_serial_no,bpl.bus_id,bpl.seat_type_id,bpl.station_id ,sti.seat_type_name " _
        & "FROM bus_price_lst bpl ,Seat_type_code sti,sell_station_info s " _
        & "WHERE bpl.sell_station_id=s.sell_station_id AND sti.seat_type_id=bpl.seat_type_id AND bpl.bus_id='" & m_szBusID & "'  AND bpl.price_table_id='" & szPriceExcuteTable & "'AND bpl.vehicle_type_code='" & VehicleModel & "'" _
        & "ORDER BY bpl.sell_station_id, bpl.station_serial_no,bpl.seat_type_id"
     Else
     szTempSql = "Select DISTINCT bpl.sell_station_id,s.sell_station_name,bpl.station_serial_no,bpl.bus_id,bpl.seat_type_id,bpl.station_id ,sti.seat_type_name " _
        & " FROM bus_price_lst bpl,Seat_type_code sti,sell_station_info s" _
        & " WHERE bpl.sell_station_id=s.sell_station_id AND sti.seat_type_id = bpl.seat_type_id AND bpl.bus_id='" & m_szBusID & "'  AND bpl.price_table_id='" & szPriceExcuteTable & "'" _
        & " ORDER BY bpl.sell_station_id, bpl.station_serial_no , bpl.seat_type_id"
     End If
     Set rsTemp = oDb.Execute(szTempSql)
     rsBusTP.MoveFirst
     If rsTemp.RecordCount = 0 Then Exit Function
     ReDim BSInfo(1 To rsTemp.RecordCount) As TBusStationSellInfo
        For i = 1 To rsTemp.RecordCount
             FindRightStatonItem rsTempStation, rsTemp!station_id
             
             BSInfo(i).szStationID = FormatDbValue(rsTempStation!station_id)
'             BSInfo(i).nMileage = FormatDbValue(rsTempStation!mileage)
             BSInfo(i).sgLimitedSellTime = FormatDbValue(rsTempStation!stop_sale_time)
             BSInfo(i).nLimitedSellCount = FormatDbValue(rsTempStation!max_sale_quantity)
             BSInfo(i).szSeatTypeName = rsTemp!seat_type_Name
             BSInfo(i).szStationName = rsTempStation!station_name
             BSInfo(i).szSeatTypeID = rsTemp!seat_type_id
             BSInfo(i).szSellStationID = rsTemp!sell_station_id
             BSInfo(i).szSellStationName = rsTemp!sell_station_name
             
             rsTicketType.MoveFirst
             For j = 1 To rsTicketType.RecordCount
                 
                 If VehicleModel = "" Then
                      FindRightPriceItem rsBusTP, BSInfo(i).szSellStationID, BSInfo(i).szStationID, rsTicketType!ticket_type_id, rsTemp!seat_type_id, , True
                 Else
                      FindRightPriceItem rsBusTP, BSInfo(i).szSellStationID, BSInfo(i).szStationID, rsTicketType!ticket_type_id, rsTemp!seat_type_id, VehicleModel, True
                 End If
'                 rsBusTP.MoveNext
                
                 'rsBusTP.MoveFirst
                 sgPrice = SelfGetTotalPrice(rsBusTP)
                 BSInfo(i).nMileage = FormatDbValue(rsBusTP!mileage)
                 Select Case rsTicketType!ticket_type_id
                           Case TP_FullPrice
                               BSInfo(i).sgFullPrice = sgPrice
                           Case TP_HalfPrice
                               BSInfo(i).sgHalfPrice = sgPrice
                           Case TP_PreferentialTicket1
                               BSInfo(i).sgPreferentialPrice1 = sgPrice
                           Case TP_PreferentialTicket2
                               BSInfo(i).sgPreferentialPrice2 = sgPrice
                           Case TP_PreferentialTicket3
                               BSInfo(i).sgPreferentialPrice3 = sgPrice
                  End Select
                 
                 rsTicketType.MoveNext
           Next
         rsTemp.MoveNext
      Next
    
 
'代票价的结构
GetAllStation = BSInfo
Exit Function
Here:
    If err.Number = 3021 Then
    ShowError ERR_StationNotInBusRoute
   
    End If
End Function


'**************************************************
'Member Code:F4
'Brief Description:获得车次的预留座位信息
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Public Function GetReserverSeat(VehicleSerial As Integer) As TReserveSeatInfo
'**************************************************
'VehicleSerial(车辆序号)
'**************************************************
Dim oDb As New RTConnection
Dim rsTemp As Recordset
Dim ReSeatInfo As TReserveSeatInfo
Dim i As Integer
Dim szTempSql As String
AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'    szTempSql = "SELECT * FROM Bus_reserve_seat_lst WHERE bus_id='" & m_szBusID & "' AND project_id='" & m_szProjectID & "' AND vehicle_serial=" & Str(VehicleSerial) & ""
    szTempSql = "SELECT * " _
            & " FROM Bus_reserve_seat_lst " _
            & " WHERE bus_id='" & m_szBusID & "' AND vehicle_serial=" & Str(VehicleSerial) & ""
    
    
    
    
    
    
    Set rsTemp = oDb.Execute(szTempSql)
    '-----------------------
    '车次无预留座位
    If rsTemp.RecordCount <> 0 Then
    ReSeatInfo.nStartSeatNo = FormatDbValue(rsTemp!start_seat_no)
    ReSeatInfo.dtBeginDate = FormatDbValue(rsTemp!start_time)
    ReSeatInfo.dtEndDate = FormatDbValue(rsTemp!end_time)
    ReSeatInfo.nSerialNo = VehicleSerial
    ReSeatInfo.nSeatCount = FormatDbValue(rsTemp!seat_quantity)
    Else
    ReSeatInfo.dtBeginDate = CDate(cszEmptyDateStr)
    ReSeatInfo.dtEndDate = CDate(cszEmptyDateStr)
    End If
    GetReserverSeat = ReSeatInfo
End Function

'**************************************************
'Member Code:F5
'Brief Description:获得车次的过站站点数
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Public Function GetPassStationCount() As Integer
Dim oDb As New RTConnection
Dim rsTemp As Recordset
Dim szTempSql As String
AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'    szTempSql = "SELECT COUNT(*) AS StationCount FROM Bus_station_info WHERE bus_id='" & m_szBusID & "' AND project_id='" & m_szProjectID & "'"
    szTempSql = "SELECT COUNT(*) AS StationCount " _
                & " FROM Bus_station_info  " _
                & " WHERE bus_id='" & m_szBusID & "'"
    
    Set rsTemp = oDb.Execute(szTempSql)
    GetPassStationCount = FormatDbValue(rsTemp!StationCount)
End Function

'**************************************************
'Member Code:F6
'Brief Description:获得该车次的所有运行车辆车型
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Public Function GetBusAllVehicleModel() As String()
    AssertObjIsValid
    
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim aszTemp() As String
    Dim rsTemp As Recordset
    Dim i As Integer
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
'    szSql = "SELECT DISTINCT vehicle_type_code FROM bus_vehicle_code,Vehicle_info WHERE " _
'    & " bus_id='" & m_szBusID & "' AND " _
'    & " project_id='" & m_szProjectID & "' AND " _
'    & " bus_vehicle_code.vehicle_id=Vehicle_info.vehicle_id "
    
    szSql = "SELECT DISTINCT vehicle_type_code " _
    & " FROM bus_vehicle_code,Vehicle_info WHERE " _
    & " bus_id='" & m_szBusID & "' AND " _
    & " bus_vehicle_code.vehicle_id=Vehicle_info.vehicle_id "
    
    
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        ReDim aszTemp(1 To rsTemp.RecordCount)
        For i = 1 To rsTemp.RecordCount
            aszTemp(i) = FormatDbValue(rsTemp!vehicle_type_code)
            rsTemp.MoveNext
        Next
    End If
    GetBusAllVehicleModel = aszTemp
    
End Function

'**************************************************
'Member Code:F7
'Brief Description:获得车次的运行车辆
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Public Function GetRunVehicle(VehicleSerial As Integer) As TBusVehicleInfo
'**************************************************
'VehicleSerial(车辆序号)
'**************************************************
Dim oDb As New RTConnection
Dim BVInfo As TBusVehicleInfo
Dim rsTemp As Recordset
Dim i As Integer
Dim szTempSql As String
AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szTempSql = "SELECT * FROM bus_vehicle_code WHERE bus_id='" & m_szBusID & "'  AND vehicle_serial=" & Str(VehicleSerial)
    '---------------------------
    '车次无运行车辆
    Set rsTemp = oDb.Execute(szTempSql)
    If rsTemp.RecordCount <> 0 Then
    BVInfo.szVehicleID = FormatDbValue(rsTemp!vehicle_id)
    BVInfo.nSerialNo = FormatDbValue(rsTemp!vehicle_serial)
    BVInfo.nStandTicketCount = FormatDbValue(rsTemp!sale_stand_ticket_quantity)
    BVInfo.dtBeginStopDate = FormatDbValue(rsTemp!stop_start_date)
    BVInfo.dtEndStopDate = FormatDbValue(rsTemp!stop_end_date)
    GetRunVehicle = BVInfo
    End If
End Function
'**************************************************
'Member Code:F8
'Brief Description:获得车次的全部运行车辆
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Public Function GetAllVehicle() As String()
Dim oDb As New RTConnection
Dim rsTemp As Recordset
Dim szaInfo() As String
Dim i As Integer
Dim szTempSql As String
AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szTempSql = "SELECT tbb.*,tbt.license_tag_no FROM bus_vehicle_code tbb,Vehicle_info tbt WHERE bus_id='" & m_szBusID & "' AND tbb.vehicle_id=tbt.vehicle_id"
    Set rsTemp = oDb.Execute(szTempSql)
    '-------------
    '车次无运行车辆
    If rsTemp.RecordCount <> 0 Then
    ReDim szaInfo(1 To rsTemp.RecordCount, 1 To 6) As String
    For i = 1 To rsTemp.RecordCount
        szaInfo(i, 1) = FormatDbValue(rsTemp!vehicle_id)
        szaInfo(i, 2) = FormatDbValue(rsTemp!vehicle_serial)
        szaInfo(i, 3) = FormatDbValue(rsTemp!sale_stand_ticket_quantity)
        szaInfo(i, 4) = FormatDbValue(rsTemp!stop_start_date)
        szaInfo(i, 5) = FormatDbValue(rsTemp!stop_end_date)
        szaInfo(i, 6) = FormatDbValue(rsTemp!license_tag_no)
        rsTemp.MoveNext
    Next
    GetAllVehicle = szaInfo()
    End If
End Function

'**************************************************
'Member Code:F9
'Brief Description:获得车次的所有过站信息和站点名称
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Public Function GetAllPassStationEx(Optional VehicleModel As String) As TBusStationSellInfo()
'Dim oDb As New RTConnection
'Dim oSystemParam As New SystemParam
'Dim rsTemp As Recordset
'Dim BSInfo() As TBusStationSellInfo
'Dim szPriceTable As String
'Dim i As Integer, j As Integer
'Dim vbBookMark As Variant
'Dim bBusPrice As Boolean
'Dim rsBusTP As Recordset
'Dim szTempSql As String
'Dim atTicketType() As TTicketType
'Dim nCount As Integer
'Dim sgPrice As Double
'Dim szExcuTablename As String
'Dim lErrCode As Long
'AssertObjIsValid
'On Error GoTo Here
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'    szTempSql = "SELECT tbg.*,tbs.station_name FROM Bus_station_info tbg,station_info tbs WHERE bus_id='" & m_szBusID & "' AND project_id='" & m_szProjectID & "' AND tbg.station_id=tbs.station_id ORDER BY tbg.station_serial_no "
'    Set rsTemp = oDb.Execute(szTempSql)
'    '***********  Modified **************
'    szExcuTablename = BusProjectExecutePrice(m_szProjectID, Now, lErrCode)
'    If lErrCode = 1 Then ShowError ERR_NoPriceExcuseTable
''    If szExcuTablename = "" Then
'    Set rsBusTP = oDb.Execute("SELECT * FROM bus_price_lst WHERE bus_id='" & m_szBusID & "' AND price_table_id='" & szExcuTablename & "' ORDER BY station_serial_no,ticket_type")
'    If rsBusTP.RecordCount = 0 Then ShowError ERR_NoBusPriceTable
'    '车次无过站
'    If rsTemp.RecordCount <> 0 Then
'    ReDim BSInfo(1 To rsTemp.RecordCount) As TBusStationSellInfo
'    For i = 1 To rsTemp.RecordCount
'        BSInfo(i).szStationID = FormatDbValue(rsTemp!station_id)
'        BSInfo(i).nMileage = FormatDbValue(rsTemp!mileage)
'        BSInfo(i).sgLimitedSellTime = FormatDbValue(rsTemp!stop_sale_time)
'        BSInfo(i).nLimitedSellCount = FormatDbValue(rsTemp!max_sale_quantity)
'        BSInfo(i).szStationName = rsTemp!station_name
'        rsTemp.MoveNext
'    Next
'    GetAllPassStationEx = BSInfo
'    End If
'Exit Function
'Here:
'    If err.Number = 3021 Then ShowError ERR_StationNotInBusRoute
End Function

'**************************************************
'Member Code:P1
'Brief Description:活动用户
'**************************************************
Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
End Property

Public Property Set SelfUser(vNewValue As ActiveUser)
    If m_nObjectStatus = ST_AddObj Then
        ShowError ERR_AddObj
    ElseIf m_nObjectStatus = ST_EditObj Then
        ShowError ERR_EditObj
    ElseIf m_nObjectStatus = ST_NormalObj Then
        ShowError ERR_NormalObj
    Else
        Set m_oActiveUser = vNewValue
    End If
End Property
'**************************************************
'Member Code:P2
'Brief Description:对象状态
'**************************************************
Public Property Get ObjStatus() As EObjectStatus
    ObjStatus = m_nObjectStatus
End Property
'**************************************************
'Member Code:P3
'Brief Description:计划代码
'**************************************************
'Public Property Let ProjectID(ByVal vData As String)
'    Dim oDb As New RTConnection
'    Dim rsTemp As Recordset
'    Dim szSql As String
'    AssertActiveUserValid m_oActiveUser, ERR_Bus
'    '查询是否是新增状态
'    If m_nObjectStatus <> ST_AddObj Then ShowError ERR_BusNotAddStatus
'    '输入的主键是否为空
'    If vData = "" Then ShowError ERR_ProjectIDNotNull
'    '是否存在该车次计划
'    szSql = "SELECT * FROM bus_project_info  WHERE project_id='" & vData & "'"
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount = 0 Then ShowError ERR_BusProjectNotExist
'    m_szProjectID = Trim(vData)
'End Property
'Public Property Get ProjectID() As String
'    ProjectID = m_szProjectID
'End Property

'**************************************************
'Member Code:P4
'Brief Description:车次代码
'**************************************************
Public Property Let BusID(ByVal vData As String)
    AssertActiveUserValid m_oActiveUser, ERR_Bus
    '查询是否是新增状态
    If m_nObjectStatus <> ST_AddObj Then ShowError ERR_BusNotAddStatus
    '输入的主键是否为空
    If vData = "" Then ShowError ERR_BusIDNotNull
    m_szBusID = Trim(vData)
End Property
Public Property Get BusID() As String
    BusID = m_szBusID
End Property
'**************************************************
'Member Code:P5
'Brief Description:线路代码
'**************************************************
Public Property Let Route(ByVal vData As String)
    Dim oDb As New RTConnection
    Dim szSql As String
    szSql = "SELECT * FROM route_info WHERE route_id='" & RTrim(vData) & "'"
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    If oDb.Execute(szSql).RecordCount = 0 Then ShowError ERR_RouteNotExist
    m_szRouteID = vData
    If m_nObjectStatus = ST_NormalObj Then m_nObjectStatus = ST_EditObj
End Property

Public Property Get Route() As String
    Route = m_szRouteID
End Property
'**************************************************
'Member Code:P6
'Brief Description:检票口
'**************************************************
Public Property Let CheckGate(ByVal vData As String)
    Dim oDb As New RTConnection
    Dim szSql As String
    szSql = "SELECT * FROM Checkgate_info WHERE check_gate_id='" & Trim(vData) & "'"
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    If oDb.Execute(szSql).RecordCount = 0 Then ShowError ERR_NoCheckGate
    m_szCheckGate = vData
    If m_nObjectStatus = ST_NormalObj Then m_nObjectStatus = ST_EditObj
End Property

Public Property Get CheckGate() As String
    CheckGate = m_szCheckGate
End Property
'**************************************************
'Member Code:P7
'Brief Description:发车时间
'**************************************************
Public Property Let StartUpTime(ByVal vData As Date)
    m_dtStartupTime = vData
    If m_nObjectStatus = ST_NormalObj Then m_nObjectStatus = ST_EditObj
End Property
Public Property Get StartUpTime() As Date
    StartUpTime = m_dtStartupTime
End Property
'**************************************************
'Member Code:P8
'Brief Description:循环周期
'**************************************************
Public Property Let RunCycle(ByVal vData As Integer)
    m_nRunCycle = vData
    If m_nObjectStatus = ST_NormalObj Then m_nObjectStatus = ST_EditObj
End Property

Public Property Get RunCycle() As Integer
    RunCycle = m_nRunCycle
End Property
'**************************************************
'Member Code:P9
'Brief Description:周期起始序号
'**************************************************
Public Property Let CycleStartSerialNo(ByVal vData As Integer)
    m_nCycleStartSerialNo = vData
    If m_nObjectStatus = ST_NormalObj Then m_nObjectStatus = ST_EditObj
End Property

Public Property Get CycleStartSerialNo() As Integer
    CycleStartSerialNo = m_nCycleStartSerialNo
End Property

'**************************************************
'Member Code:P10
'Brief Description:车次停班开始日期
'**************************************************
Public Property Let BeginStopDate(ByVal vData As Date)
    m_dtBeginStopDate = vData
    If m_nObjectStatus = ST_NormalObj Then m_nObjectStatus = ST_EditObj
End Property

Public Property Get BeginStopDate() As Date
    BeginStopDate = m_dtBeginStopDate
End Property
'**************************************************
'Member Code:P11
'Brief Description:车次停班结束日期
'**************************************************
Public Property Let EndStopDate(ByVal vData As Date)
    m_dtEndStopDate = vData
    If m_nObjectStatus = ST_NormalObj Then m_nObjectStatus = ST_EditObj
End Property
Public Property Get EndStopDate() As Date
    EndStopDate = m_dtEndStopDate
End Property
'**************************************************
'Member Code:P12
'Brief Description:车次类型
'**************************************************
Public Property Let BusType(ByVal vData As Integer)
    Dim oDb As New RTConnection
    Dim szSql As String
    szSql = "SELECT * FROM Bus_type_code WHERE bus_type='" & CInt(vData) & "'"
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    If oDb.Execute(szSql).RecordCount = 0 Then ShowError ERR_NoBusTypeInfo
   
    m_nBusType = vData
    If m_nObjectStatus = ST_NormalObj Then m_nObjectStatus = ST_EditObj
End Property
Public Property Get BusType() As Integer
    BusType = m_nBusType
End Property

'**************************************************
'Member Code:P16
'Brief Description:线路名称
'**************************************************
Public Property Get RouteName() As String
    RouteName = m_szRouteName
End Property

Public Property Get StartStationID() As String
    StartStationID = m_szStartStationID
End Property

Public Property Get StartStationName() As String
    StartStationName = m_szStartStationName
End Property


Public Property Let InternetStatus(ByVal vData As Integer)
   
    m_nOldInternetStatus = m_nInternetStatus
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_nInternetStatus = vData
'    #If SUPPORT_MSG = 1 Then
'        If m_nObjectStatus <> ST_AddObj And dtOldStartupTime <> m_dtStartupTime Then
'            m_bOffTimeChanged = True
'        End If
'    #End If
End Property

'P19
Public Property Get InternetStatus() As Integer
    InternetStatus = m_nInternetStatus
End Property

Public Property Let ScrollBusCheckTime(vData As Integer)
   If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
   m_nScrollBusCheckTime = vData
End Property

'P35滚动车次循环剪票时间
Public Property Get ScrollBusCheckTime() As Integer
    ScrollBusCheckTime = m_nScrollBusCheckTime
End Property

'**************************************************
'Member Code:S1
'Brief Description:初始化对象
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Public Sub Init(poAUser As ActiveUser)
    Set m_oActiveUser = poAUser
    'AssertHaveRight m_oActiveUser, RIGHT_BusManagement
End Sub

'**************************************************
'Member Code:S2
'Brief Description:类初始化
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Private Sub Class_Initialize()
    m_nObjectStatus = ST_NotAvailable
End Sub

'**************************************************
'Member Code:S3
'Brief Description:测试对象是否有效
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then ShowError ERR_NotAvailable
End Sub
'**************************************************
'Member Code:S4
'Brief Description:测试对象是否有效
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Private Sub AssertObjIsValid()
    AssertActiveUserValid m_oActiveUser, ERR_Bus
    AssertStatusAvailable
End Sub
'**************************************************
'Member Code:S5
'Brief Description:获得该车次得属性
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
'ProjectID As String,
Public Sub Identify(BusID As String)
    AssertActiveUserValid m_oActiveUser, ERR_Bus
    m_szBusID = Trim(BusID)
'    m_szProjectID = Trim(ProjectID)
    RefreshMemoryInfo
    m_nObjectStatus = ST_NormalObj
End Sub

'**************************************************
'Member Code:S7
'Brief Description:新增车次
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Public Sub AddNew()
    AssertActiveUserValid m_oActiveUser, ERR_Bus
    m_nObjectStatus = ST_AddObj
End Sub

'**************************************************
'Member Code:S6
'Brief Description:删除对象
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Public Sub Delete()
    AssertObjIsValid
    DeleteObject
    m_nObjectStatus = ST_NotAvailable
End Sub

'**************************************************
'Member Code:S7
'Brief Description:修改车次对象属性获新增对象
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Public Sub Update()
    AssertObjIsValid
    Select Case m_nObjectStatus
           Case ST_AddObj
           AddObject
           Case ST_EditObj
           UpdateToDB
           Case ST_NormalObj
           Case Else
           ShowError ERR_NotAvailable
    End Select
    m_nObjectStatus = ST_NormalObj
End Sub
'保存车次配载信息


'**************************************************
'Member Code:S8
'Brief Description:获得该车次得属性
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Private Sub RefreshMemoryInfo()
    Dim oDb As New RTConnection
    Dim i As Integer
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim eType As ETicketType
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    '1获得车次的基本信息
    szSql = " SELECT tbe.*,tbr.route_name, tbr.start_station_id , tbr.end_station_id , s.station_name start_station_name , e.station_name end_station_name " _
        & " FROM bus_info tbe,route_info tbr , station_info s ,station_info e " _
        & " WHERE tbr.start_station_id = s.station_id AND tbr.end_station_id = e.station_id AND tbe.bus_id='" & m_szBusID & "'" _
        & " AND tbe.route_id=tbr.route_id"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_BusNotExist
    m_szCheckGate = FormatDbValue(rsTemp!check_gate_id)
    m_szRouteID = FormatDbValue(rsTemp!route_id)
    m_dtStartupTime = FormatDbValue(rsTemp!bus_start_time)
    m_nRunCycle = FormatDbValue(rsTemp!bus_run_cycle)
    m_nCycleStartSerialNo = FormatDbValue(rsTemp!run_start_serial)
    m_dtBeginStopDate = FormatDbValue(rsTemp!stop_start_date)
    m_dtEndStopDate = FormatDbValue(rsTemp!stop_end_date)
    m_nBusType = FormatDbValue(rsTemp!bus_type)
    m_szRouteName = FormatDbValue(rsTemp!route_name)
    m_nInternetStatus = FormatDbValue(rsTemp!Internet_Status)
    m_nOldInternetStatus = FormatDbValue(rsTemp!Internet_Status)
    m_nScrollBusCheckTime = FormatDbValue(rsTemp!ScrollBus_check_time)
    
    m_szEndStationID = FormatDbValue(rsTemp!end_station_id)
    m_szEndStationName = FormatDbValue(rsTemp!end_station_name)
    m_szStartStationID = FormatDbValue(rsTemp!start_station_id)
    m_szStartStationName = FormatDbValue(rsTemp!start_station_name)
    
    m_szOldCheckGate = m_szCheckGate
    m_szOldRoute = m_szRouteID
    m_dtOldStartupTime = m_dtStartupTime
    m_nOldRunCycle = m_nRunCycle
    m_nOldCycleStartSerialNo = m_nCycleStartSerialNo
    m_nOldBusType = m_nBusType
    m_nOldScrollBusCheckTime = m_nScrollBusCheckTime
    m_dtOldBeginStopDate = m_dtBeginStopDate
    m_dtOldEndStopDate = m_dtEndStopDate
 '取得终点站
' szSql = "select s.station_id,s.station_name from  " _
'       & " Route_section_lst r ,station_info s  where section_serial=" _
'       & " (select top 1 Max(section_serial) as section_serial" _
'       & " from Route_section_lst where route_id ='" & m_szRouteID & "') " _
'       & " and r.end_station_id=s.station_id and r.route_id ='" & m_szRouteID & "'"
'Set rsTemp = oDb.Execute(szSql)
'
'If rsTemp.RecordCount <> 0 Then

  
'End If
Set oDb = Nothing
End Sub

'**************************************************
'Member Code:S9
'Brief Description:新增车次
'Engineer:wjb
'Date Generated:
'Last Revision Date:2002/11/11
'**************************************************
Private Sub AddObject()
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szTempSql As String
    Dim i As Integer
    Dim szSqlField As String
    Dim szSqlContext As String
    Dim szPriceTable As String
    Dim lErrCode As Long
    AssertHaveRight m_oActiveUser, RIGHT_BusAdd
    szPriceTable = BusProjectExecutePrice(Date, lErrCode)
    If lErrCode = 1 Then ShowError ERR_NoPriceExcuseTable
    
'    If Trim(m_szBusID) = "" Or Trim(m_szProjectID) = "" Then ShowError ERR_BusIDNotNull
    
    If Trim(m_szBusID) = "" Then ShowError ERR_BusIDNotNull
    
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    '1查询车次是否存在
'    szTempSql = "SELECT * FROM bus_info WHERE bus_id='" & m_szBusID & "'AND project_id='" & m_szProjectID & "'"
    szTempSql = "SELECT * FROM bus_info WHERE bus_id='" & m_szBusID & "'"
    
    
    
    Set rsTemp = oDb.Execute(szTempSql)
    If rsTemp.RecordCount = 1 Then ShowError ERR_BusExistNotAddNew
    
    '3.查寻检票口
    szTempSql = "SELECT * FROM Checkgate_info WHERE check_gate_id='" & m_szCheckGate & "'"
    Set rsTemp = oDb.Execute(szTempSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_NoCheckGate
    '4查寻线路
    szTempSql = "SELECT * FROM route_info WHERE mileage=(SELECT MAX(mileage) FROM route_info WHERE route_id = '" & m_szRouteID & "') AND route_id='" & m_szRouteID & "'"
    Set rsTemp = oDb.Execute(szTempSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_NoSpecifyRoute
    
    On Error GoTo Here
'    Dim m_szProjectID As String
'    m_szProjectID = "0001"
    oDb.BeginTrans
    '2新增车次
    szTempSql = "INSERT bus_info ("
    szSqlField = "bus_id,"
    szSqlField = szSqlField & "check_gate_id,"
    szSqlField = szSqlField & "route_id,"
    szSqlField = szSqlField & "bus_start_time,"
    szSqlField = szSqlField & "bus_run_cycle,"
    szSqlField = szSqlField & "run_start_serial,"
    szSqlField = szSqlField & "bus_type,"
    
    szSqlField = szSqlField & "internet_status,"
    szSqlField = szSqlField & "ScrollBus_Check_Time,"
    
    szSqlField = szSqlField & "stop_end_date,"
    szSqlField = szSqlField & "stop_start_date)"
    szSqlContext = " VALUES ('" & m_szBusID & "',"
    szSqlContext = szSqlContext & " '" & m_szCheckGate & "',"
    szSqlContext = szSqlContext & " '" & m_szRouteID & "',"
    szSqlContext = szSqlContext & "'" & ToDBDateTime(m_dtStartupTime) & "',"
    szSqlContext = szSqlContext & m_nRunCycle & ","
    szSqlContext = szSqlContext & m_nCycleStartSerialNo & ","
    szSqlContext = szSqlContext & m_nBusType & ","
    
    szSqlContext = szSqlContext & m_nInternetStatus & ","
    szSqlContext = szSqlContext & m_nScrollBusCheckTime & ","
    
    szSqlContext = szSqlContext & "'" & ToDBDate(m_dtEndStopDate) & "',"
    szSqlContext = szSqlContext & "'" & ToDBDate(m_dtBeginStopDate) & "')"
    szTempSql = szTempSql & szSqlField & szSqlContext
    
    oDb.Execute (szTempSql)
    '新增车次线路表
    '获得线路的所有站点、限售张数、里程数
    '以下为函数 RefreshPassStation
     Dim rsBusStation As New Recordset
     Dim rsNewRoute As New Recordset
     Dim bUpdate As Boolean
     
'    szTempSql = "SELECT * FROM Bus_station_info WHERE bus_id='" & m_szBusID & "'AND project_id='" & m_szProjectID & "' ORDER BY station_serial_no"
    szTempSql = "SELECT * " _
                & " FROM Bus_station_info " _
                & " WHERE bus_id='" & m_szBusID & "'" _
                & " ORDER BY station_serial_no"
    
    Set rsBusStation = oDb.Execute(szTempSql)
    
'    szTempSql = "DELETE Bus_station_info  WHERE bus_id='" & m_szBusID & "'AND project_id='" & m_szProjectID & "'"
    szTempSql = "DELETE Bus_station_info  " _
                & " WHERE bus_id='" & m_szBusID & "'"
    
    oDb.Execute szTempSql
    
    szTempSql = "SELECT * FROM Route_section_lst a," & _
                            " (SELECT a.sell_station_id,a.route_id FROM  Route_section_lst a," & _
                                                                " (SELECT route_id,MAX(section_serial) section_serial FROM Route_section_lst" & _
                                                                "   WHERE route_id=" & TransFieldValueToString(m_szRouteID) & _
                                                                "   GROUP BY route_id) b" & _
                            " WHERE a.route_id=b.route_id AND a.section_serial=b.section_serial) b" & _
                " WHERE station_type=" & TP_CanSellTicket & _
                " AND a.route_id=b.route_id AND a.sell_station_id=b.sell_station_id" & _
                " ORDER BY section_serial "
    Set rsNewRoute = oDb.Execute(szTempSql)
    '新增车次站点
    For i = 1 To rsNewRoute.RecordCount
        If rsBusStation.RecordCount <> 0 Then  '  add
        Do While Not rsBusStation.EOF
             If rsBusStation!station_id = FormatDbValue(rsNewRoute!end_station_id) Then bUpdate = True: Exit Do
             rsBusStation.MoveNext
        Loop
        End If
        szTempSql = "INSERT Bus_station_info ("
        szSqlField = "bus_id,"
        szSqlField = szSqlField & "station_id,"
        szSqlField = szSqlField & "station_serial_no,"
        szSqlField = szSqlField & "max_sale_quantity,"
        szSqlField = szSqlField & "stop_sale_time,"
        szSqlField = szSqlField & "mileage)"
        szSqlContext = " VALUES ('" & m_szBusID & "',"
        szSqlContext = szSqlContext & "'" & rsNewRoute!end_station_id & "',"
        szSqlContext = szSqlContext & i & ","
        '------------------------
        'bUpdate为真则将旧的站点信息保存
        If bUpdate Then
             szSqlContext = szSqlContext & Str(rsBusStation!max_sale_quantity) & ","
             szSqlContext = szSqlContext & Str(rsBusStation!stop_sale_time) & ","
        Else
            szSqlContext = szSqlContext & "-1,"
            szSqlContext = szSqlContext & "-1,"
        End If
        szSqlContext = szSqlContext & Str(rsNewRoute!end_station_mileage) & ")"
        szTempSql = szTempSql & szSqlField & szSqlContext
        oDb.Execute szTempSql
        bUpdate = False
        rsNewRoute.MoveNext
    Next
    
    '新增车次配信息
        '增加起点站
'    szTempSql = "INSERT bus_allot_info (bus_id,project_id,sell_station_id,check_gate_id,bus_start_time,can_sale_quantity)" & _
'                " SELECT b.bus_id,b.project_id,s.sell_station_id,b.check_gate_id,b.bus_start_time,-1" & _
'                " FROM bus_info b,route_info g,sell_station_info s" & _
'                " WHERE  b.route_id = g.route_id And s.station_id = g.start_station_id" & _
'                " AND b.bus_id=" & TransFieldValueToString(m_szBusID) & " AND b.project_id=" & TransFieldValueToString(m_szProjectID)
    
    szTempSql = "INSERT bus_allot_info (bus_id,sell_station_id,check_gate_id,bus_start_time,can_sale_quantity)" & _
                " SELECT b.bus_id,s.sell_station_id,b.check_gate_id,b.bus_start_time,-1" & _
                " FROM bus_info b,route_info g,sell_station_info s" & _
                " WHERE  b.route_id = g.route_id And s.station_id = g.start_station_id" & _
                " AND b.bus_id=" & TransFieldValueToString(m_szBusID)
    
    oDb.Execute szTempSql
    '新增车次售票点信息表
    szTempSql = "INSERT bus_sell_station_lst (bus_id,sell_station_id,can_sale_quantity)" & _
                " SELECT " & TransFieldValueToString(m_szBusID) & " ,s.sell_station_id,-1" & _
                " FROM sell_station_info s"
    
    oDb.Execute szTempSql
    Dim szSellStationID As String
''    szSellStationID = Trim(oDb.Execute("SELECT sell_station_id FROM bus_allot_info WHERE bus_id=" & TransFieldValueToString(m_szBusID) & " AND project_id=" & TransFieldValueToString(m_szProjectID)).Fields(0))
'    szTempSql = "SELECT sell_station_id " _
'                & " FROM bus_allot_info " _
'                & " WHERE bus_id='" & m_szBusID & "'"
'    szSellStationID = Trim(oDb.Execute(szTempSql)!sell_station_id)
    
    szSellStationID = Trim(oDb.Execute("SELECT sell_station_id FROM bus_allot_info WHERE bus_id=" & TransFieldValueToString(m_szBusID)).Fields(0))
    
'    szTempSql = "INSERT bus_allot_info (bus_id,project_id,sell_station_id,check_gate_id,bus_start_time,can_sale_quantity)" & _
'                " SELECT b.bus_id,b.project_id,s.sell_station_id,b.check_gate_id,b.bus_start_time,-1" & _
'                " FROM bus_info b,Route_section_lst g,sell_station_info s" & _
'                " WHERE  b.route_id = g.route_id And s.station_id = g.end_station_id AND g.sell_station_id=" & TransFieldValueToString(szSellStationID) & _
'                " AND b.bus_id=" & TransFieldValueToString(m_szBusID) & " AND b.project_id=" & TransFieldValueToString(m_szProjectID)
   
    szTempSql = "INSERT bus_allot_info (bus_id,sell_station_id,check_gate_id,bus_start_time,can_sale_quantity)" & _
                " SELECT b.bus_id,s.sell_station_id,b.check_gate_id,b.bus_start_time,-1" & _
                " FROM bus_info b,Route_section_lst g,sell_station_info s" & _
                " WHERE  b.route_id = g.route_id And s.station_id = g.end_station_id AND g.sell_station_id=" & TransFieldValueToString(szSellStationID) & _
                " AND b.bus_id=" & TransFieldValueToString(m_szBusID)
                
    
   
   oDb.Execute szTempSql
    oDb.CommitTrans
    
    Set rsBusStation = Nothing
    Set rsNewRoute = Nothing
    Set oDb = Nothing
    Set rsTemp = Nothing
    WriteOperateLog m_oActiveUser, RIGHT_BusAdd, "新增车次" & m_szBusID
    WriteOperateLog m_oActiveUser, RIGHT_BusAdd, "刷新车次" & m_szBusID & "的线路站点信息"
Exit Sub
Here:
    Set rsBusStation = Nothing
    Set rsNewRoute = Nothing
    Set rsTemp = Nothing
    oDb.RollbackTrans
    AssertAddObjectError ERR_Bus, oDb
End Sub

'**************************************************
'Member Code:S10
'Brief Description:删除该车次
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Private Sub DeleteObject()
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szExePricertable As String
    Dim lErrCode As Long
    On Error GoTo Here
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_BusDelete
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szExePricertable = BusProjectExecutePrice(Now, lErrCode)
    oDb.BeginTrans
'    szSql = "DELETE bus_info WHERE bus_id='" & m_szBusID & "' AND project_id='" & m_szProjectID & "'"
    
    szSql = "DELETE bus_info WHERE bus_id='" & m_szBusID & "'"
    oDb.Execute (szSql)
    '以下只是空操作，因为已定义了触发器
    'szSql = "DELETE Bus_station_info WHERE bus_id='" & m_szBusID & "' AND project_id='" & m_szProjectID & "'"
    If lErrCode <> 1 Then
       szSql = "DELETE bus_price_lst WHERE bus_id='" & m_szBusID & "' AND price_table_id='" & szExePricertable & "'"
       oDb.Execute (szSql)
    End If
    oDb.CommitTrans
    WriteOperateLog m_oActiveUser, RIGHT_BusDelete, "删除计划车次" & m_szBusID
    Exit Sub
Here:
    ShowError err.Number
End Sub

'**************************************************
'Member Code:S11
'Brief Description:修改车次属性
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Private Sub UpdateToDB()
    Dim oDb As New RTConnection
    Dim szTempSql As String
    Dim szSqlContext As String
    Dim szSqlField As String
    Dim eType As ETicketType
    Dim szSqlWhere As String
    Dim strlog As String
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_BusUpdate
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    '1修改车次基本信息
    On Error GoTo Here
    oDb.BeginTrans
    szTempSql = "UPDATE bus_info SET "
    szSqlField = "check_gate_id='" & m_szCheckGate & "',"
    szSqlField = szSqlField & "route_id='" & m_szRouteID & "',"
    szSqlField = szSqlField & "bus_start_time='" & ToDBDateTime(m_dtStartupTime) & "',"
    szSqlField = szSqlField & "bus_run_cycle=" & Str(m_nRunCycle) & ","
    szSqlField = szSqlField & "run_start_serial=" & Str(m_nCycleStartSerialNo) & ","
    szSqlField = szSqlField & "bus_type=" & Str(m_nBusType) & ","
    szSqlField = szSqlField & "stop_end_date='" & ToDBDate(m_dtEndStopDate) & "',"
    szSqlField = szSqlField & "stop_start_date='" & ToDBDate(m_dtBeginStopDate) & "',"
    szSqlField = szSqlField & "ScrollBus_Check_Time=" & m_nScrollBusCheckTime & ","
    szSqlField = szSqlField & "internet_status=" & m_nInternetStatus & ""
    szSqlWhere = " WHERE bus_id='" & m_szBusID & "'"
'    szSqlWhere = " WHERE bus_id='" & m_szBusID & "'AND project_id='" & m_szProjectID & "'"
    szTempSql = szTempSql & szSqlField & szSqlWhere
    oDb.Execute (szTempSql)
  
    If m_dtStartupTime <> m_dtOldStartupTime Or m_szCheckGate <> m_szOldCheckGate Then  '对应调整配载信息
'        szTempSql = "UPDATE bus_allot_info SET bus_start_time=" & TransFieldValueToString(m_dtStartupTime) & _
'                                                    ",check_gate_id=" & TransFieldValueToString(m_szCheckGate) & _
'                " WHERE bus_id=" & TransFieldValueToString(m_szBusID) & " AND project_id=" & TransFieldValueToString(m_szProjectID) & _
'                " AND bus_start_time=(SELECT MIN(bus_start_time) FROM bus_allot_info" &

        szTempSql = "UPDATE bus_allot_info SET bus_start_time=" & TransFieldValueToString(m_dtStartupTime) & _
                ",check_gate_id=" & TransFieldValueToString(m_szCheckGate) & _
                " WHERE bus_id=" & TransFieldValueToString(m_szBusID) & _
                " AND check_gate_id=" & TransFieldValueToString(m_szOldCheckGate) & ""
        oDb.Execute szTempSql
    End If
        
    oDb.CommitTrans
 
 
If Trim(m_szOldRoute) <> Trim(m_szRouteID) Then
    'RefreshPass
    StationRefresh
    strlog = strlog & "线路由[" & m_szOldRoute & "]改为" & m_szRouteID & ","
    m_szOldRoute = m_szRouteID
End If
 
If Trim(m_szOldCheckGate) <> Trim(m_szCheckGate) And m_szOldCheckGate <> "" Then
    strlog = strlog & "检票口由[" & m_szOldCheckGate & "]改为" & m_szCheckGate & ","
    m_szOldCheckGate = m_szCheckGate
End If

If Format(m_dtOldStartupTime, cszTimeStr) <> Format(m_dtStartupTime, cszTimeStr) Then
    strlog = strlog & "发车时间由[" & Format(m_dtOldStartupTime, cszTimeStr) & "]改为" & Format(m_dtStartupTime, cszTimeStr) & ","
    m_dtOldStartupTime = m_dtStartupTime
End If

If Trim(m_nOldRunCycle) <> Trim(m_nRunCycle) Then
    strlog = strlog & "运行周期由[" & m_nOldRunCycle & "]改为" & m_nRunCycle & ","
    m_nOldRunCycle = m_nRunCycle
End If

If Trim(m_nOldCycleStartSerialNo) <> Trim(m_nCycleStartSerialNo) Then
    strlog = strlog & "起始序号由[" & m_nOldCycleStartSerialNo & "]改为" & m_nCycleStartSerialNo & ","
    m_nOldCycleStartSerialNo = m_nCycleStartSerialNo
End If

If Trim(m_nOldBusType) <> Trim(m_nBusType) Then
   strlog = strlog & "车次种类由[" & m_nOldBusType & "]改为" & m_nBusType & ","
   m_nOldBusType = m_nBusType
End If

If Trim(m_nOldScrollBusCheckTime) <> Trim(m_nScrollBusCheckTime) Then
    strlog = strlog & "滚动车次开检时间[" & m_nOldScrollBusCheckTime & "]改为" & m_nScrollBusCheckTime & ","
    m_nOldScrollBusCheckTime = m_nScrollBusCheckTime
End If

If strlog <> "" Then
    strlog = Left(strlog, Len(strlog) - 1)
    Dim szLogTemp As String
    szLogTemp = "发车时间[" & Format(m_dtOldStartupTime, cszTimeStr) & "终点站[" & m_szEndStationName & "]"
End If


If m_dtOldEndStopDate <> m_dtEndStopDate Or m_dtOldBeginStopDate <> m_dtBeginStopDate Then
    If m_dtEndStopDate = cszForeverDateStr And m_dtBeginStopDate = cszEmptyDateStr Then
        strlog = strlog & "车次计划长停,"
    ElseIf m_dtEndStopDate = cszEmptyDateStr And m_dtBeginStopDate = cszEmptyDateStr Then
        strlog = strlog & "车次计划复班,"
    End If
    strlog = strlog & "结束停班日期[" & ToDBDate(m_dtOldEndStopDate) & "]改为" & ToDBDate(m_dtEndStopDate) & ","
    strlog = strlog & "开始停班日期[" & ToDBDate(m_dtOldBeginStopDate) & "]改为" & ToDBDate(m_dtBeginStopDate) & ","

    m_dtOldBeginStopDate = m_dtBeginStopDate
    m_dtOldEndStopDate = m_dtEndStopDate
End If
strlog = "计划车次[" & m_szBusID & "]" & szLogTemp & "]属性修改:" & strlog
WriteOperateLog m_oActiveUser, RIGHT_BusUpdate, strlog
 
Exit Sub
Here:
    oDb.RollbackTrans
    AssertUpdateObjectError ERR_Bus, oDb
End Sub

'**************************************************
'Member Code:S12
'Brief Description:计算半票价
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Public Sub MakeHalfTicketPrice()
       AssertObjIsValid
End Sub
'**************************************************
'Member Code:S13
'Brief Description:车次车辆预留座位
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Public Sub ReserveSeat(VehicleSerial As Integer, RSinfo As TReserveSeatInfo)
'**************************************************
'VehicleSerial(车辆序号)
'RSinfo(车次车辆预留座位结构)
'**************************************************
Dim szSql As String
Dim oDb As New RTConnection
Dim szTempSql As String
Dim szSqlField As String
Dim szSqlWhere As String
Dim szVehicleID As String
Dim szSqlContext As String
Dim rsTemp As Recordset
AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_BusManagement
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    '1查询该车次车辆表中查询车辆序号得对象车辆,车次序号无对应车辆
    
'    szSql = "SELECT * FROM bus_vehicle_code WHERE vehicle_serial=" & Str(VehicleSerial) & " AND bus_id='" & m_szBusID & "'AND project_id='" & m_szProjectID & "'"
    
    szSql = "SELECT * " _
        & " FROM bus_vehicle_code " _
        & " WHERE vehicle_serial=" & Str(VehicleSerial) & " AND bus_id='" & m_szBusID & "'"
    
    Set rsTemp = oDb.Execute(szSql)
    
    If rsTemp.RecordCount = 0 Then ShowError ERR_BusSerialNotVehicle
    
    szVehicleID = rsTemp!vehicle_id
    '2判断该车辆是否有预留，如果有则以新座位状态修改，否则新增座位状态
    szVehicleID = rsTemp!vehicle_id
    
'    szSql = "SELECT * FROM Bus_reserve_seat_lst WHERE vehicle_serial=" & Str(VehicleSerial) & " AND bus_id='" & m_szBusID & "'AND project_id='" & m_szProjectID & "'"
    szSql = "SELECT * " _
             & " FROM Bus_reserve_seat_lst " _
             & " WHERE vehicle_serial=" & Str(VehicleSerial) & " AND bus_id='" & m_szBusID & "'"
    Set rsTemp = oDb.Execute(szSql)
    '2.1新增车辆座位预留
    If rsTemp.RecordCount = 0 Then
        szTempSql = "INSERT Bus_reserve_seat_lst ("
        szSqlField = "bus_id,"
'        szSqlField = szSqlField & "project_id,"
        szSqlField = szSqlField & "vehicle_serial,"
        szSqlField = szSqlField & "start_time,"
        szSqlField = szSqlField & "end_time,"
        szSqlField = szSqlField & "start_seat_no,"
        szSqlField = szSqlField & "seat_quantity)"
        szSqlContext = " VALUES ('" & m_szBusID & "',"
'        szSqlContext = szSqlContext & " '" & m_szProjectID & "',"
        szSqlContext = szSqlContext & Str(VehicleSerial) & ","
        szSqlContext = szSqlContext & "'" & Trim(ToDBDate(RSinfo.dtBeginDate)) & "',"
        szSqlContext = szSqlContext & "'" & ToDBDate(RSinfo.dtEndDate) & "',"
        szSqlContext = szSqlContext & Trim(Str(RSinfo.nStartSeatNo)) & ","
        szSqlContext = szSqlContext & Trim(Str(RSinfo.nSeatCount)) & ")"
        szTempSql = szTempSql & szSqlField & szSqlContext
        oDb.Execute (szTempSql)
    Else '2.2修改座位预留
        szTempSql = "UPDATE Bus_reserve_seat_lst SET "
        szSqlField = "start_time='" & ToDBDate(RSinfo.dtBeginDate) & "',"
        szSqlField = szSqlField & "end_time='" & ToDBDate(RSinfo.dtEndDate) & "',"
        szSqlField = szSqlField & "start_seat_no='" & Trim(Str(RSinfo.nStartSeatNo)) & "',"
        szSqlField = szSqlField & "seat_quantity=" & Trim(Str(RSinfo.nSeatCount)) & ""
        szSqlWhere = " WHERE bus_id='" & m_szBusID & "' AND vehicle_serial=" & Str(VehicleSerial)
        szTempSql = szTempSql & szSqlField & szSqlWhere
        oDb.Execute (szTempSql)
    End If
    Dim szLogTemp As String
    szLogTemp = "发车时间[" & Format(m_dtOldStartupTime, cszTimeStr) & "终点站[" & m_szEndStationName & "]"
    WriteOperateLog m_oActiveUser, RIGHT_BusManagement, "预留车次[" & m_szBusID & "]" & szLogTemp & "座位共" & RSinfo.nSeatCount & "个起始号" & RSinfo.nStartSeatNo
End Sub

'**************************************************
'Member Code:S14
'Brief Description:修改该车次运行车辆信息
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Public Sub ModifyRunVehicle(BusVehicle As TBusVehicleInfo)
'**************************************************
'VehicleSerial(车辆序号)
'BusVehicle(车次车辆结构)
'**************************************************
Dim oDb As New RTConnection
Dim szTempSql As String
Dim rsTemp As Recordset
Dim szSqlWhere As String
Dim szSqlField As String

AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_BusVehicleManagement
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    '判断原来是否有该序号的车辆
'    szTempSql = "SELECT bus_id FROM bus_vehicle_code WHERE bus_id='" & m_szBusID & "' AND project_id='" & m_szProjectID & "' AND vehicle_serial=" & Trim(BusVehicle.nSerialNo)
    szTempSql = "SELECT bus_id " _
                & " FROM bus_vehicle_code " _
                & " WHERE bus_id='" & m_szBusID & "' AND vehicle_serial=" & Trim(BusVehicle.nSerialNo)
    Set rsTemp = oDb.Execute(szTempSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_BusNotExistVehicleSerial
    szTempSql = "UPDATE bus_vehicle_code SET "
    szSqlField = "stop_start_date='" & ToDBDate(BusVehicle.dtBeginStopDate) & "',"
    szSqlField = szSqlField & "stop_end_date='" & ToDBDate(BusVehicle.dtEndStopDate) & "',"
    szSqlField = szSqlField & "sale_stand_ticket_quantity=" & BusVehicle.nStandTicketCount
'    szSqlWhere = " WHERE bus_id='" & m_szBusID & "' AND project_id='" & m_szProjectID & "' AND vehicle_serial=" & Trim(BusVehicle.nSerialNo)
    szSqlWhere = " WHERE bus_id='" & m_szBusID & "' AND vehicle_serial=" & Trim(BusVehicle.nSerialNo)
    szTempSql = szTempSql & szSqlField & szSqlWhere
    oDb.Execute (szTempSql)
    WriteOperateLog m_oActiveUser, RIGHT_BusVehicleManagement, "修改车次" & m_szBusID & "运行车辆" & BusVehicle.szVehicleID
End Sub

'**************************************************
'Member Code:S15
'Brief Description:该车次删除一运营车辆
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Public Sub DeleteRunVehicle(VehicleSerial As Integer)
'**************************************************
'VehicleSerial(车辆序号)
'**************************************************
Dim szSql As String
Dim oDb As New RTConnection
AssertObjIsValid
     AssertHaveRight m_oActiveUser, RIGHT_BusVehicleManagement
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    '查询该车次中是否有该车次车辆
    If oDb.Execute("SELECT * FROM bus_vehicle_code WHERE bus_id='" & m_szBusID & "'  AND vehicle_serial=" & Str(VehicleSerial)).RecordCount = 0 Then ShowError ERR_BusNotExistVehicleSerial
    szSql = "DELETE bus_vehicle_code WHERE bus_id='" & m_szBusID & "' AND vehicle_serial=" & Str(VehicleSerial)
    oDb.Execute szSql
    
    Dim szLogTemp As String
    szLogTemp = "发车时间[" & Format(m_dtOldStartupTime, cszTimeStr) & "终点站[" & m_szEndStationName & "]"
     
    WriteOperateLog m_oActiveUser, RIGHT_BusVehicleManagement, "删除车次[" & m_szBusID & "]" & szLogTemp & "的运行车辆序号" & VehicleSerial
End Sub

'**************************************************
'Member Code:S16
'Brief Description:车次新增一运营车辆
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Public Sub AddRunVehicle(BusVehicle As TBusVehicleInfo)
'**************************************************
'BusVehicle(车次车辆结构)
'**************************************************
Dim szTempSql As String
Dim szSqlField As String
Dim szSqlContext As String
Dim oDb As New RTConnection
AssertObjIsValid
   AssertHaveRight m_oActiveUser, RIGHT_BusStationManagement
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    '-------------------------
    '车次车辆序号不能为0
    If BusVehicle.nSerialNo = 0 Then ShowError ERR_BusVehicleSerialNotZero
    '车次车辆的序号已存在不能新增
    szTempSql = "SELECT bus_id FROM bus_vehicle_code WHERE bus_id='" & m_szBusID & "'  AND vehicle_serial=" & BusVehicle.nSerialNo
    If oDb.Execute(szTempSql).RecordCount <> 0 Then ShowError ERR_BusVehicleSerialIsExist
    '车辆不存在
    If oDb.Execute("SELECT vehicle_id FROM Vehicle_info WHERE vehicle_id='" & BusVehicle.szVehicleID & "'").RecordCount = 0 Then ShowError ERR_VehicleNotExist
'    Dim m_szProjectID As String
   
    szTempSql = "INSERT bus_vehicle_code ("
    szSqlField = szSqlField & "bus_id,"
    szSqlField = szSqlField & "vehicle_serial,"
    szSqlField = szSqlField & "vehicle_id,"
    szSqlField = szSqlField & "stop_start_date,"
    szSqlField = szSqlField & "stop_end_date,"
    szSqlField = szSqlField & "sale_stand_ticket_quantity)"
    szSqlContext = " VALUES ('" & m_szBusID & "',"
    szSqlContext = szSqlContext & Str(BusVehicle.nSerialNo) & ","
    szSqlContext = szSqlContext & "'" & BusVehicle.szVehicleID & "',"
    szSqlContext = szSqlContext & "'" & ToDBDate(BusVehicle.dtBeginStopDate) & "',"
    szSqlContext = szSqlContext & "'" & ToDBDate(BusVehicle.dtEndStopDate) & "',"
    szSqlContext = szSqlContext & Str(BusVehicle.nStandTicketCount) & ")"
    szTempSql = szTempSql & szSqlField & szSqlContext
    oDb.Execute (szTempSql)
      Dim szLogTemp As String
   szLogTemp = "发车时间[" & Format(m_dtOldStartupTime, cszTimeStr) & "终点站[" & m_szEndStationName & "]"

    WriteOperateLog m_oActiveUser, RIGHT_BusStationManagement, "车次[" & m_szBusID & "]" & szLogTemp & "新增运行车辆:" & BusVehicle.szVehicleID
End Sub

Public Sub ModifyPassStationSellInfo(BusPassStationInfo As TBusStationSellInfo)
'**************************************************
'BusPassStationInfoEx1(车次站点结构)
'**************************************************
Dim oDb As New RTConnection
Dim szTempSql As String
Dim szSqlWhere As String
Dim szSqlField As String
Dim rsTemp As New Recordset
Dim bLimitedSellCount As Boolean
Dim bLimitedSellTime As Boolean
Dim szStationName As String
Dim szlog As String
AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szTempSql = "SELECT g.max_sale_quantity,g.stop_sale_time,s.station_name " _
                & " FROM Bus_station_info  g,station_info s " _
                & " WHERE g.bus_id='" & m_szBusID & "'" _
                & " AND g.station_id='" & BusPassStationInfo.szStationID & "'" _
                & " and s.station_id=g.station_id"
    Set rsTemp = oDb.Execute(szTempSql)
    If rsTemp.RecordCount <> 0 Then
        szStationName = rsTemp!station_name
        If Trim(rsTemp!max_sale_quantity) <> Trim(Str(BusPassStationInfo.nLimitedSellCount)) Then
             szSqlField = "max_sale_quantity=" & Str(BusPassStationInfo.nLimitedSellCount) & ","
            szlog = "限售张数由[" & Trim(rsTemp!max_sale_quantity) & "]改为[" & Trim(Str(BusPassStationInfo.nLimitedSellCount))
        End If
        
        If Trim(rsTemp!stop_sale_time) <> Trim(Str(BusPassStationInfo.sgLimitedSellTime)) Then
            szSqlField = szSqlField & "stop_sale_time=" & Str(BusPassStationInfo.sgLimitedSellTime) & ","
            szlog = szlog & "限售时间由[" & Trim(rsTemp!stop_sale_time) & "]改为[" & Trim(Str(BusPassStationInfo.sgLimitedSellTime)) & "]"
        End If
    
        If szSqlField <> "" Then
        szTempSql = "UPDATE Bus_station_info SET "
        szSqlField = szSqlField & "mileage=" & Str(BusPassStationInfo.nMileage)
        szSqlWhere = " WHERE bus_id='" & m_szBusID & "'  AND station_id='" & BusPassStationInfo.szStationID & "'"
        szTempSql = szTempSql & szSqlField & szSqlWhere
        oDb.Execute (szTempSql)
        Dim szLogTemp As String
         szLogTemp = "发车时间[" & Format(m_dtOldStartupTime, cszTimeStr) & "终点站[" & m_szEndStationName & "]"
        szlog = szlog & "编辑车次[" & m_szBusID & "]" & szLogTemp & "站点" & szStationName & "[" & BusPassStationInfo.szStationID & "]:" & szlog
        If Len(szlog) > 255 Then
          szlog = Left(szlog, 255)
        End If
        WriteOperateLog m_oActiveUser, RIGHT_BusStationManagement, szlog
        End If
    End If
End Sub
'**************************************************
'Member Code:S18
'Brief Description:以车次运行的线路为基准，对车次站点表进行修改，并保留旧的车次站点属性
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Public Sub RefreshPassStation()
Dim oDb As New RTConnection
'Dim tBusStation() As TBusStationSellInfo
Dim rsNewRoute As Recordset
Dim rsBusStation As Recordset
Dim j As Integer
Dim i As Integer
Dim bUpdate As Boolean
Dim nStationCount As Integer
Dim szTempSql As String
Dim szSqlContext As String
Dim szSqlField As String
Dim oSystemParam As New SystemParam
Dim nCount As Integer
Dim atTicketType() As TTicketType


On Error GoTo Here
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_BusStationManagement
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    oDb.BeginTrans
        szTempSql = "SELECT * " _
                   & " FROM Bus_station_info " _
                   & " WHERE bus_id='" & m_szBusID & "'" _
                   & " ORDER BY station_serial_no"
        Set rsBusStation = oDb.Execute(szTempSql)
        '删除原先的车次站点信息
        szTempSql = "DELETE Bus_station_info  " _
                    & " WHERE bus_id='" & m_szBusID & "'"
        oDb.Execute szTempSql
        
        
        szTempSql = "SELECT * FROM Route_section_lst a," & _
                        " (SELECT a.sell_station_id,a.route_id FROM  Route_section_lst a," & _
                                                            " (SELECT route_id,MAX(section_serial) section_serial FROM Route_section_lst" & _
                                                            "   WHERE route_id=" & TransFieldValueToString(m_szRouteID) & _
                                                            "   GROUP BY route_id) b" & _
                        " WHERE a.route_id=b.route_id AND a.section_serial=b.section_serial) b" & _
                " WHERE station_type=" & TP_CanSellTicket & _
                " AND a.route_id=b.route_id AND a.sell_station_id=b.sell_station_id" & _
                " ORDER BY section_serial "
        Set rsNewRoute = oDb.Execute(szTempSql)
        
        '新增车次站点
        For i = 1 To rsNewRoute.RecordCount
            '查找此站点原先是否有.如有,则保留原先的站点的设置信息
            If rsBusStation.RecordCount <> 0 Then
                bUpdate = False
                rsBusStation.MoveFirst
                Do While Not rsBusStation.EOF
                     If FormatDbValue(rsBusStation!station_id) = FormatDbValue(rsNewRoute!end_station_id) Then
                        bUpdate = True
                        Exit Do
                     End If
                     rsBusStation.MoveNext
                Loop
            End If
            szTempSql = "INSERT Bus_station_info ("
            szSqlField = "bus_id,"
            szSqlField = szSqlField & "station_id,"
            szSqlField = szSqlField & "station_serial_no,"
            szSqlField = szSqlField & "max_sale_quantity,"
            szSqlField = szSqlField & "stop_sale_time,"
            szSqlField = szSqlField & "mileage)"
            szSqlContext = " VALUES ('" & m_szBusID & "',"
            szSqlContext = szSqlContext & "'" & rsNewRoute!end_station_id & "',"
            szSqlContext = szSqlContext & i & ","
            '------------------------
            'bUpdate为真则将旧的站点信息保存
            If bUpdate Then
                 szSqlContext = szSqlContext & Str(rsBusStation!max_sale_quantity) & ","
                 szSqlContext = szSqlContext & Str(rsBusStation!stop_sale_time) & ","
            Else
                szSqlContext = szSqlContext & "-1,"
                szSqlContext = szSqlContext & "-1,"
            End If
            szSqlContext = szSqlContext & Str(rsNewRoute!end_station_mileage) & ")"
            szTempSql = szTempSql & szSqlField & szSqlContext
            oDb.Execute szTempSql
            bUpdate = False
            rsNewRoute.MoveNext
        Next i
    oDb.CommitTrans
    
    Dim szLogTemp As String
    szLogTemp = "发车时间[" & Format(m_dtOldStartupTime, cszTimeStr) & "终点站[" & m_szEndStationName & "]"
    WriteOperateLog m_oActiveUser, RIGHT_BusStationManagement, "同步车次[" & m_szBusID & "]" & szLogTemp & "的线路站点信息"
    Exit Sub
Here:
   oDb.RollbackTrans
   err.Raise err.Number
End Sub
'**************************************************
'Member Code:S19
'Brief Description:取消预留
'Engineer:
'Date Generated:
'Last Revision Date:1999/10/25
'**************************************************
Public Sub UnReserveSeat(VehicleSerial As Integer)
Dim szSql As String
Dim oDb As New RTConnection
AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_BusManagement
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "DELETE FROM Bus_reserve_seat_lst " _
            & " WHERE vehicle_serial=" & Str(VehicleSerial) & " AND bus_id='" & m_szBusID & "'"
    oDb.Execute szSql
    Dim szLogTemp As String
    szLogTemp = "发车时间[" & Format(m_dtOldStartupTime, cszTimeStr) & "终点站[" & m_szEndStationName & "]"
    WriteOperateLog m_oActiveUser, RIGHT_BusManagement, "取消预留车次[" & m_szBusID & "]" & szLogTemp & "的座位"
End Sub

'内部用将票价信息的记录集移合适的位置
Private Function FindRightPriceItem(prsPriceInfo As Recordset, ByVal pszSellStationID As String, ByVal pszStationID As String, ByVal pnTicketType As ETicketType, Optional szSeatType As String, Optional szVehicleModalID As String = "", Optional pbFromFirst As Boolean = False) As Integer
    Dim szMsg As String
    
    If pbFromFirst Then
        prsPriceInfo.MoveFirst
    Else
        If prsPriceInfo.Bookmark <> 1 Then prsPriceInfo.MovePrevious
    End If
    
    Do While Not prsPriceInfo.EOF
        If szVehicleModalID = "" Then
            If FormatDbValue(prsPriceInfo!sell_station_id) = Trim(pszSellStationID) And FormatDbValue(prsPriceInfo!station_id) = pszStationID And Trim(prsPriceInfo!ticket_type) = Trim(pnTicketType) And Trim(prsPriceInfo!seat_type_id) = szSeatType Then
                Exit Function
            End If
        Else
            If FormatDbValue(prsPriceInfo!sell_station_id) = Trim(pszSellStationID) And FormatDbValue(prsPriceInfo!station_id) = Trim(pszStationID) And Trim(prsPriceInfo!ticket_type) = Trim(pnTicketType) And Trim(prsPriceInfo!seat_type_id) = Trim(szSeatType) And Trim(prsPriceInfo!vehicle_type_code) = Trim(szVehicleModalID) Then
                Exit Function
            End If
        End If
        prsPriceInfo.MoveNext
    Loop
    
Error_Handle:
    
    
'    Select Case pnTicketType
'    Case TP_FullPrice
'        szMsg = szMsg & "全票"
'    Case TP_HalfPrice
'        szMsg = szMsg & "半票"
'    Case TP_PreferentialTicket1
'        szMsg = szMsg & "半票"
    

End Function

'Public Function PreviewBus(BusID As String, RunDate As Date) As String()
'
'End Function

'**************************************************
'Member Code:S20
'Brief Description:删除车次的所有运行车辆
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Public Sub DeleteAllRunVehicle()
Dim szSql As String
Dim oDb As New RTConnection
AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_BusManagement
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "DELETE bus_vehicle_code WHERE bus_id='" & m_szBusID & "'"
    oDb.Execute szSql
     Dim szLogTemp As String
 szLogTemp = "发车时间[" & Format(m_dtOldStartupTime, cszTimeStr) & "终点站[" & m_szEndStationName & "]"
  
    WriteOperateLog m_oActiveUser, RIGHT_BusManagement, "删除车次[" & m_szBusID & "]" & szLogTemp & "所有运行车辆"
End Sub

Public Sub CloneBus(CloneBusProject As String, CloneBusID As String, Optional BusVehicle As Boolean = False, Optional BusStop As Boolean = False, Optional BusReserve As Boolean = False)
    Dim ExeSQL As String
    Dim szSql As String
    Dim szset As String
    Dim oDb As New RTConnection
    Dim szPriceTableName As String
    Dim lErrCode As Long
    Dim rsTemp As New Recordset
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_BusManagement
    szPriceTableName = BusProjectExecutePrice(Now, lErrCode)
    If lErrCode = 1 Then ShowError ERR_NoPriceExcuseTable
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    ExeSQL = "SELECT bus_id FROM bus_info WHERE bus_id='" & CloneBusID & "'"
    If oDb.Execute(ExeSQL).RecordCount = 1 Then ShowError ERR_BusExistNotAddNew
    On Error GoTo Here
    oDb.BeginTrans
        '1在计划中添加车次
        szSql = ""
        szset = ""
        szSql = "INSERT INTO bus_info ("
        '添加车次代码
        szSql = szSql & "bus_id,"
        szset = szset & "'" & Trim(CloneBusID) & "' AS bus_id,"
        '添加计划代码
        
        
        '添加线路
        szSql = szSql & "route_id,check_gate_id,bus_start_time,bus_run_cycle,run_start_serial,internet_status, ScrollBus_check_time,"
        szset = szset & "route_id,check_gate_id,bus_start_time,bus_run_cycle,run_start_serial,internet_status, ScrollBus_check_time,"
        If BusStop Then '添加停班开始日期/结束日期
            szSql = szSql & "stop_start_date,stop_end_date,"
            szset = szset & "stop_start_date,stop_end_date,"
        Else
            szSql = szSql & "stop_start_date,stop_end_date,"
            szset = szset & "'" & ToDBDate(CDate(cszEmptyDateStr)) & "','" & ToDBDate(CDate(cszEmptyDateStr)) & "',"
        End If
        '添加车次类型
        
        
        szSql = szSql & "bus_type"
        szset = szset & "bus_type"
        ExeSQL = szSql & ") SELECT " & szset _
        & " FROM bus_info WHERE bus_id='" & m_szBusID & "'"
        
        oDb.Execute ExeSQL
        
        '2复制车次站点表
        szSql = ""
        szset = ""
        szSql = "INSERT INTO Bus_station_info ("
        '添加车次代码
        szSql = szSql & "bus_id,"
        szset = szset & "'" & Trim(CloneBusID) & "' AS bus_id,"
        '添加计划代码
'        szSql = szSql & "project_id,"
'        szset = szset & "'" & Trim(CloneBusProject) & "' AS project_id,"
        '添加线路
        szSql = szSql & "station_id,station_serial_no,max_sale_quantity,stop_sale_time,mileage"
        szset = szset & "station_id,station_serial_no,max_sale_quantity,stop_sale_time,mileage"
        ExeSQL = szSql & ") SELECT " & szset _
        & " FROM Bus_station_info WHERE bus_id='" & m_szBusID & "'"
        
        oDb.Execute ExeSQL
        
        If BusVehicle Then
            '2在计划中添加车次车辆
            szSql = ""
            szset = ""
            szSql = "INSERT INTO bus_vehicle_code ("
            '添加车次代码
            szSql = szSql & "bus_id,"
            szset = szset & "'" & Trim(CloneBusID) & "' AS bus_id,"
            '添加计划代码
'            szSql = szSql & "project_id,"
'            szset = szset & "'" & Trim(CloneBusProject) & "' AS project_id,"
            '添加车辆序号
            szSql = szSql & "vehicle_serial,vehicle_id,"
            szset = szset & "vehicle_serial,vehicle_id,"
            
            If BusStop Then
                '添加停班开始日期/结束日期
                szSql = szSql & "stop_start_date,"
                szset = szset & "stop_start_date,"
                szSql = szSql & "stop_end_date,"
                szset = szset & "stop_end_date,"
            Else
                '添加停班开始日期/结束日期
                szSql = szSql & "stop_start_date,"
                szset = szset & "'" & ToDBDate(CDate(cszEmptyDateStr)) & "',"
                szSql = szSql & "stop_end_date,"
                szset = szset & "'" & ToDBDate(CDate(cszEmptyDateStr)) & "',"
            End If
            
            '添加车次类型
            szSql = szSql & "sale_stand_ticket_quantity"
            szset = szset & "sale_stand_ticket_quantity"
            ExeSQL = szSql & ") SELECT " & szset _
            & " FROM bus_vehicle_code WHERE bus_id='" & m_szBusID & "'"
            oDb.Execute ExeSQL
        End If
        
        '    '3复制车次票价表
        
'        Dim szExePrice As String
'        szExePrice = BusProjectExecutePrice(Now, lErrCode)
'        If szExePrice <> "" Then
            szSql = "SELECT Bus_id FROM  bus_price_lst WHERE bus_id='" & m_szBusID & "'" ' " and price_table_id='" & Trim(szExePrice) & "'"
            Set rsTemp = oDb.Execute(szSql)
            If rsTemp.RecordCount <> 0 And BusVehicle Then
                Set rsTemp = Nothing
                szSql = ""
                szset = ""
                szSql = "INSERT INTO bus_price_lst ("
                '添加起点站
                szSql = szSql & "sell_station_id,"
                szset = szset & "sell_station_id,"
                '添加车次代码
                szSql = szSql & "bus_id,"
                szset = szset & "'" & Trim(CloneBusID) & "' AS bus_id,"
                szSql = szSql & "price_table_id,"
                szset = szset & "price_table_id,"
                '添加线路
                '************** 加*******************
                '*************   复制所有信息,不关联执行票价表 ,关联票价表可能出错  *************
                szSql = szSql & "station_id,ticket_type,seat_type_id,station_serial_no,mileage,base_carriage,vehicle_type_code,"
                szset = szset & "station_id,ticket_type,seat_type_id,station_serial_no,mileage,base_carriage,vehicle_type_code,"
                '*************   复制所有信息,不关联执行票价表 ,关联票价表可能出错  *************
                '************** 加*******************
                Dim i As Integer
                For i = 1 To 14
                    szSql = szSql & "price_item_" & i & ","
                    szset = szset & "price_item_" & i & ","
                Next
                szSql = szSql & "price_item_15"
                szset = szset & "price_item_15"
                ExeSQL = szSql & ") SELECT " & szset _
                & " FROM bus_price_lst  WHERE bus_id='" & m_szBusID & "'" ' " and  price_table_id='" & szExePrice & "'"   ' AND project_id='" & m_szProjectID & "'"
                oDb.Execute ExeSQL
            End If
'        End If
        '复制该车次的配载信息
        szSql = " INSERT bus_allot_info (bus_id ,sell_station_id , check_gate_id, bus_start_time , can_sale_quantity ) " _
            & " SELECT '" & Trim(CloneBusID) & "' , " & " sell_station_id , check_gate_id, bus_start_time , can_sale_quantity " _
            & " FROM bus_allot_info " _
            & " WHERE bus_id='" & m_szBusID & "'"
        oDb.Execute szSql
        
        '复制该车次的售票点信息
        szSql = " INSERT bus_sell_station_lst (bus_id ,sell_station_id, can_sale_quantity ) " _
            & " SELECT '" & Trim(CloneBusID) & "' , " & " sell_station_id , can_sale_quantity " _
            & " FROM bus_sell_station_lst " _
            & " WHERE bus_id='" & m_szBusID & "'"
        oDb.Execute szSql
        
        '3复制车次车辆预留座位
        If BusReserve And BusVehicle Then
            szSql = ""
            szset = ""
            szSql = "INSERT INTO Bus_reserve_seat_lst ("
            '添加车次代码
            szSql = szSql & "bus_id,"
            szset = szset & "'" & Trim(CloneBusID) & "' AS bus_id,"
            '添加计划代码
'            szSql = szSql & "project_id,"
'            szset = szset & "'" & Trim(CloneBusProject) & "' AS project_id,"
            '添加车辆序号
            szSql = szSql & "vehicle_serial,start_time,end_time,start_seat_no,seat_quantity"
            szset = szset & "vehicle_serial,start_time,end_time,start_seat_no,seat_quantity"
            ExeSQL = szSql & ") SELECT " & szset _
            & " FROM Bus_reserve_seat_lst WHERE bus_id='" & m_szBusID & "'"
            oDb.Execute ExeSQL
        End If
    oDb.CommitTrans
    
    Dim szLogTemp As String
    szLogTemp = "发车时间[" & Format(m_dtOldStartupTime, cszTimeStr) & "终点站[" & m_szEndStationName & "]"
    
    WriteOperateLog m_oActiveUser, RIGHT_BusManagement, " 复制车次[" & m_szBusID & "]" & szLogTemp & "为" & CloneBusID
    Exit Sub
Here:
    oDb.RollbackTrans
    AssertAddObjectError ERR_Bus, oDb
End Sub
Public Function GetSeatTypeFromBusPriceLst(szPriceTable As String, Optional szVehicleModal As String = "") As String()
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim szSeatType() As String
    Dim i As Long
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    If szVehicleModal <> "" Then
    szSql = "SELECT DISTINCT bpi.seat_type_id,sti.seat_type_name FROM bus_price_lst bpi,Seat_type_code sti where bpi.bus_id='" & m_szBusID & "' and bpi.seat_type_id=sti.seat_type_id and bpi.price_table_id='" & szPriceTable & "'and pbi.vehicle_type_code='" & szVehicleModal & "'"
    Else
    szSql = "SELECT DISTINCT bpi.seat_type_id,sti.seat_type_name FROM bus_price_lst bpi,Seat_type_code sti where bpi.bus_id='" & m_szBusID & "' and bpi.seat_type_id=sti.seat_type_id and bpi.price_table_id='" & szPriceTable & "'"
    End If
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 0 Then
       ReDim szSeatType(1 To rsTemp.RecordCount, 1 To 2)
       For i = 1 To rsTemp.RecordCount
           szSeatType(i, 1) = rsTemp!seat_type_id
           szSeatType(i, 2) = rsTemp!seat_type_Name
       Next
       GetSeatTypeFromBusPriceLst = szSeatType
    End If
End Function
Private Function FindRightStatonItem(prsPriceInfo As Recordset, ByVal pszStationID As String) As Integer
    Dim szMsg As String
      prsPriceInfo.MoveFirst
    
    Do While Not prsPriceInfo.EOF
           If Trim(FormatDbValue(prsPriceInfo!station_id)) = Trim(pszStationID) Then
                Exit Function
           End If
        prsPriceInfo.MoveNext
    Loop
    
Error_Handle:
End Function


'生成空票价表---新增车次调用
'
'Public Function MakeBusPrice(tBusVehicle As TBusVehicleInfo)
'    Dim szSql As String
'    Dim rsTemp As Recordset
'    Dim oDb As New RTConnection
'    Dim i As Integer
'    Dim j As Integer
'    Dim rsStation As Recordset
'    Dim rstPrice As Recordset
'    Dim nTemp As Integer
'    Dim szRoute As String
'    Dim szsqlFile As String
'    Dim szPriceTable As String
'    Dim szset As String
'    Dim ExeSQL As String
'    Dim nStationCount As String
'    Dim szBusVehicleModalId As String
'    Dim lErrCode As Long
'    AssertObjIsValid
'    On Error GoTo Error_Handle
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'     oDb.BeginTrans
'     szPriceTable = BusProjectExecutePrice(m_szProjectID, Date, lErrCode)
'     If lErrCode = 1 Then ShowError ERR_NoPriceExcuseTable
'
'     '取得车型
'     If szPriceTable = "" Then ShowError ERR_NoPriceExcuseTable
'     szSql = "select vehicle_type_code from Vehicle_info where vehicle_id='" & tBusVehicle.szVehicleID & "'  "
'     Set rsTemp = oDb.Execute(szSql)
'     szBusVehicleModalId = rsTemp!vehicle_type_code
'
''     '查看车次票价表是否有该车次信息如有退出,
'     '如无: 1--将线路票价表中取得复制到车次票价表中,2-生成空票价表
'     szSql = "SELECT * from  bus_price_lst  where  " _
'              & "bus_id='" & m_szBusID & "' AND price_table_id='" & Trim(szPriceTable) & "'and vehicle_type_code='" & szBusVehicleModalId & "'"
'     Set rstPrice = oDb.Execute(szSql)
'
'     If rstPrice.RecordCount <> 0 Then Exit Function
'
'     '查看线路票价表是否有该线路信息
'     szSql = "SELECT * from  price_table_lst  where  " _
'              & "route_id='" & m_szRouteID & "' AND price_table_id='" & Trim(szPriceTable) & "'and vehicle_type_code='" & szBusVehicleModalId & "'"
'     Set rstPrice = oDb.Execute(szSql)
'
'     If rstPrice.RecordCount <> 0 Then
'         '1--将线路票价表中取得复制到车次票价表中
'         szset = ""
'         szSql = ""
'         szSql = "INSERT INTO bus_price_lst ("
'         szSql = szSql & "bus_id,"
'         szset = szset & "'" & Trim(m_szBusID) & "' AS bus_id,"
'         szSql = szSql & "station_id,price_table_id,ticket_type,vehicle_type_code,seat_type_id,station_serial_no,mileage,base_carriage"
'         szset = szset & "station_id,price_table_id,ticket_type,vehicle_type_code,seat_type_id,station_serial_no,mileage,base_carriage"
'         For i = 1 To 15
'              If i <= 12 Then
'              szSql = szSql & ", price_item_" & i
'              szset = szset & ", price_item_" & i
'              Else
'               szSql = szSql & ", price_item_" & i
'               szset = szset & "," & 0 & " AS price_item_" & i
'              End If
'         Next
'         ExeSQL = szSql & ") SELECT " & szset _
'         & " FROM price_table_lst  WHERE vehicle_type_code='" & szBusVehicleModalId & "'and route_id='" & m_szRouteID & "' AND price_table_id='" & szPriceTable & "'"
'         oDb.Execute ExeSQL
'
'   Else
'    '2-生成空票价表
''    szSql = "SELECT route_id,formula_name FROM route_info  WHERE "
''    szRoute = szRoute & " route_id ='" & m_szRouteID & "'"
''    szSql = szSql & szRoute
''    Set rstemp = odb.Execute(szSql)
''    If rstemp.RecordCount <> 1 Then ShowError ERR_NoSpecifyRoute
''    '取得车次座位类型
''    Dim tSeatType() As TVehcileSeatType
''    Dim nSeatQuantity As Integer
''    Dim nSeatQuantityTemp As Integer
''    Dim nCountTemp As Integer
''    Dim szSqlSeatTypeWhere As String
''    Dim bHaveNormalSeatType As Boolean '有正常的座位类型
''    Dim baddNormalSeatType As Boolean
''    tSeatType = GetBusVehicleSeatType(tBusVehicle.szVehicleID)
''    '计算总座位
''    nCountTemp = ArrayLength(tSeatType)
''    For i = 1 To ArrayLength(tSeatType)
''    If tSeatType(i).szSeatTypeID = cszSeatTypeIsNormal Then bHaveNormalSeatType = True
''    If CInt(tSeatType(i).szEndSeatNo) > CInt(tSeatType(i).szStartSeatNo) Then
''       nSeatQuantityTemp = nSeatQuantityTemp + CInt(tSeatType(i).szEndSeatNo) - CInt(tSeatType(i).szStartSeatNo) + 1
''    Else
''        nSeatQuantityTemp = nSeatQuantityTemp + CInt(tSeatType(i).szStartSeatNo) - CInt(tSeatType(i).szEndSeatNo) + 1
''    End If
''    Next
''    szSql = "SELECT seat_quantity FROM Vehicle_info WHERE vehicle_id='" & tBusVehicle.szVehicleID & "'"
''    Set rstemp = odb.Execute(szSql)
''    If bHaveNormalSeatType = False And CInt(rstemp!seat_quantity) > nSeatQuantityTemp Then
''    baddNormalSeatType = True
''    If nCountTemp = 0 Then
''    ReDim tSeatType(1 To nCountTemp + 1)
''    Else
''    ReDim Preserve tSeatType(1 To nCountTemp + 1)
''    End If
''    nCountTemp = nCountTemp + 1
''    tSeatType(nCountTemp).szSeatTypeID = cszSeatTypeIsNormal
''    End If
''    '取得车次基本信息
''    szSql = "SELECT rsa.*,station_name,sfo.area_code,rfo.route_name, " _
''         & " tti.ticket_type_id , tti.ticket_type_name, vtc.vehicle_type_code, vtc.vehicle_type_short_name " _
''         & " FROM Route_section_lst rsa,station_info sfo,route_info rfo,  " _
''         & " Ticket_type_code tti,vehicle_type_code vtc " _
''         & " WHERE rfo.route_id='" & m_szRouteID & "' AND rsa.end_station_id=sfo.station_id AND rsa.station_type=1 " _
''         & " AND rfo.route_id=rsa.route_id AND ticket_type_valid=" & TP_TicketTypeValid & " AND " _
''         & " tti.ticket_type_id<>'" & TP_FreeTicket & "' and vtc.vehicle_type_code='" & szBusVehicleModalId & "' " _
''         & " ORDER BY rsa.route_id,vtc.vehicle_type_code,rsa.end_station_mileage, rsa.end_station_id," _
''         & " tti.ticket_type_id"
''    Set rsStation = odb.Execute(szSql)
''    nStationCount = rsStation.RecordCount
''    If rsStation.RecordCount = 0 Then ShowError ERR_NoStationInfo
''    '插入空车次票价
''    Dim szSqlFild1 As String
''    For i = 1 To 15
''      szsqlFile = szsqlFile & ", price_item_" & i
''      szSqlFild1 = szSqlFild1 & ",0"
''    Next
''    For i = 1 To nCountTemp
''        rsStation.MoveFirst
''        For j = 1 To rsStation.RecordCount
''          szSql = " Insert Into bus_price_lst ( " _
''                  & "bus_id ," _
''                  & "station_id, " _
''                  & "price_table_id," _
''                  & "ticket_type," _
''                  & "vehicle_type_code, " _
''                  & "seat_type_id," _
''                  & "station_serial_no," _
''                  & "mileage," _
''                  & "base_carriage"
''           szSql = szSql & szsqlFile & ")"
''           szSql = szSql & "Values (" _
''                  & "'" & Trim(m_szBusID) & "' ," _
''                  & "'" & Trim(rsStation!end_station_id) & "', " _
''                  & "'" & Trim(szPriceTable) & "'," _
''                  & "'" & rsStation!ticket_type_id & "'," _
''                  & "'" & szBusVehicleModalId & "', " _
''                  & "'" & Trim(tSeatType(i).szSeatTypeID) & "'," _
''                  & "'" & rsStation!section_serial & "'," _
''                  & "'" & rsStation!end_station_mileage & "'," _
''                  & "0"
''             szSql = szSql & szSqlFild1 & ")"
''            odb.Execute szSql
''            rsStation.MoveNext
''         Next
''    Next
'End If
'    Set rsStation = Nothing
'    oDb.CommitTrans
'    Set oDb = Nothing
'    Set rsTemp = Nothing
'    Exit Function
'Error_Handle:
'        oDb.RollbackTrans
'        If Not rsTemp Is Nothing Then Set rsTemp = Nothing
'        If Not rsStation Is Nothing Then Set rsStation = Nothing
'        If Not oDb Is Nothing Then Set oDb = Nothing
'        err.Raise err.Number
'End Function


Private Function GetBusVehicleSeatType(pszVehicleID As String) As TVehcileSeatType()
Dim oDb As New RTConnection
Dim szSql As String
Dim rsTemp As Recordset
Dim i As Integer
Dim tvTemp() As TVehcileSeatType
oDb.ConnectionString = GetConnectionStr(cszSystemMan)
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
'
''计划站点调用 请用 GetPassStationEx
'Public Function GetPassStation(Optional VehicleModel As String) As TBusStationSellInfo()
'Dim odb As New RTConnection
'Dim oSystemParam As New SystemParam
'Dim rsTemp As Recordset
'Dim BSInfo() As TBusStationSellInfo
'Dim i As Integer, j As Integer
'Dim szTempSql As String
'AssertObjIsValid
'On Error GoTo Here
'    odb.ConnectionString = GetConnectionStr(cszSystemMan)
'    szTempSql = "SELECT tbg.*,tbs.station_name FROM Bus_station_info tbg,station_info tbs WHERE bus_id='" & m_szBusID & "' AND project_id='" & m_szProjectID & "' AND tbg.station_id=tbs.station_id ORDER BY tbg.station_serial_no "
'    Set rsTemp = odb.Execute(szTempSql)
'    If rsTemp.RecordCount <> 0 Then
'    ReDim BSInfo(1 To rsTemp.RecordCount) As TBusStationSellInfo
'    For i = 1 To rsTemp.RecordCount
'        BSInfo(i).szStationID = FormatDbValue(rsTemp!station_id)
'        BSInfo(i).nMileage = FormatDbValue(rsTemp!mileage)
'        BSInfo(i).sgLimitedSellTime = FormatDbValue(rsTemp!stop_sale_time)
'        BSInfo(i).nLimitedSellCount = FormatDbValue(rsTemp!max_sale_quantity)
'        BSInfo(i).szStationName = rsTemp!station_name
'        rsTemp.MoveNext
'    Next
'    GetPassStation = BSInfo
'    End If
'Exit Function
'Here:
'    If err.Number = 3021 Then ShowError ERR_StationNotInBusRoute
'End Function
Public Function GetPassStation(Optional VehicleModel As String) As TBusStationSellInfo()
Dim oDb As New RTConnection
Dim oSystemParam As New SystemParam
Dim rsTemp As Recordset
Dim BSInfo() As TBusStationSellInfo
Dim i As Integer, j As Integer
Dim szTempSql As String
AssertObjIsValid
On Error GoTo Here
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szTempSql = "SELECT tbg.*,tbs.station_name FROM Bus_station_info tbg,station_info tbs WHERE bus_id='" & m_szBusID & "' AND tbg.station_id=tbs.station_id ORDER BY tbg.station_serial_no "
    Set rsTemp = oDb.Execute(szTempSql)
    If rsTemp.RecordCount <> 0 Then
    ReDim BSInfo(1 To rsTemp.RecordCount) As TBusStationSellInfo
    For i = 1 To rsTemp.RecordCount
        BSInfo(i).szStationID = FormatDbValue(rsTemp!station_id)
        BSInfo(i).nMileage = FormatDbValue(rsTemp!mileage)
        BSInfo(i).sgLimitedSellTime = FormatDbValue(rsTemp!stop_sale_time)
        BSInfo(i).nLimitedSellCount = FormatDbValue(rsTemp!max_sale_quantity)
        BSInfo(i).szStationName = rsTemp!station_name
        rsTemp.MoveNext
    Next
    GetPassStation = BSInfo
    End If
Exit Function
Here:
    If err.Number = 3021 Then ShowError ERR_StationNotInBusRoute
End Function

Public Sub StationRefresh()
Dim oDb As New RTConnection
Dim rsNewRoute As Recordset
Dim i As Integer
Dim szTempSql As String
Dim szSqlContext As String
Dim szSqlField As String
On Error GoTo Here
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_BusStationManagement
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    oDb.BeginTrans
    
    szTempSql = "DELETE bus_allot_info " _
                & " WHERE bus_id=" & TransFieldValueToString(m_szBusID)
    oDb.Execute szTempSql
    
    szTempSql = "INSERT bus_allot_info (bus_id,sell_station_id,check_gate_id,bus_start_time,can_sale_quantity)" & _
                " SELECT b.bus_id,s.sell_station_id,b.check_gate_id,b.bus_start_time,-1" & _
                " FROM bus_info b,route_info g,sell_station_info s" & _
                " WHERE  b.route_id = g.route_id And s.station_id = g.start_station_id" & _
                " AND bus_id=" & TransFieldValueToString(m_szBusID)
    oDb.Execute szTempSql
    Dim szSellStationID As String
'    szSellStationID = Trim(oDb.Execute("SELECT sell_station_id  " _
'                                       & " FROM bus_allot_info " _
'                                       & " Where bus_id = " & TransFieldValueToString(m_szBusID)))
'
    szSellStationID = Trim(oDb.Execute("SELECT sell_station_id FROM bus_allot_info WHERE bus_id=" & TransFieldValueToString(m_szBusID)).Fields(0))
                                           
    szTempSql = "INSERT bus_allot_info (bus_id,sell_station_id,check_gate_id,bus_start_time,can_sale_quantity)" & _
                " SELECT b.bus_id,s.sell_station_id,b.check_gate_id,b.bus_start_time,-1" & _
                " FROM bus_info b,Route_section_lst g,sell_station_info s" & _
                " WHERE  b.route_id = g.route_id And s.station_id = g.end_station_id AND g.sell_station_id=" & TransFieldValueToString(szSellStationID) & _
                " AND b.bus_id=" & TransFieldValueToString(m_szBusID)
    oDb.Execute szTempSql
    
    szTempSql = "DELETE Bus_station_info  WHERE bus_id='" & m_szBusID & "'"
    oDb.Execute szTempSql
    '将新的线路的信息读入
    szTempSql = "SELECT * FROM Route_section_lst a," & _
                        " (SELECT a.sell_station_id,a.route_id FROM  Route_section_lst a," & _
                                                            " (SELECT route_id,MAX(section_serial) section_serial FROM Route_section_lst" & _
                                                            "   WHERE route_id=" & TransFieldValueToString(m_szRouteID) & _
                                                            "   GROUP BY route_id) b" & _
                        " WHERE a.route_id=b.route_id AND a.section_serial=b.section_serial) b" & _
            " WHERE station_type=" & TP_CanSellTicket & _
            " AND a.route_id=b.route_id AND a.sell_station_id=b.sell_station_id" & _
            " ORDER BY section_serial "
    Set rsNewRoute = oDb.Execute(szTempSql)
    '新增车次站点
    For i = 1 To rsNewRoute.RecordCount
        szTempSql = "INSERT Bus_station_info ("
        szSqlField = "bus_id,"
'        szSqlField = szSqlField & "project_id,"
        szSqlField = szSqlField & "station_id,"
        szSqlField = szSqlField & "station_serial_no,"
        szSqlField = szSqlField & "max_sale_quantity,"
        szSqlField = szSqlField & "stop_sale_time,"
        szSqlField = szSqlField & "mileage)"
        szSqlContext = " VALUES ('" & m_szBusID & "',"
'        szSqlContext = szSqlContext & "'" & m_szProjectID & "',"
        szSqlContext = szSqlContext & "'" & rsNewRoute!end_station_id & "',"
        szSqlContext = szSqlContext & i & ","
        szSqlContext = szSqlContext & "-1,"
        szSqlContext = szSqlContext & "-1,"
        szSqlContext = szSqlContext & Str(rsNewRoute!end_station_mileage) & ")"
        szTempSql = szTempSql & szSqlField & szSqlContext
        oDb.Execute szTempSql
        rsNewRoute.MoveNext
    Next
'' Dim tTBusvehicleInfo As TBusVehicleInfo
'' MakeBusPrice tTBusvehicleInfo
'    MakeBusPrice tTBusvehicleInfo
    oDb.CommitTrans
    WriteOperateLog m_oActiveUser, RIGHT_BusStationManagement, "刷新车次" & m_szBusID & "的线路站点信息"
    Exit Sub
Here:
   oDb.RollbackTrans
   err.Raise err.Number
End Sub
'Public Function DeleteBusPrice()
'Dim oDb As New RTConnection
'Dim i As Integer
'Dim szSql As String
'Dim szExePrice As String
'Dim lErrCode As Long
'AssertObjIsValid
'AssertHaveRight m_oActiveUser, RIGHT_BusStationManagement
' oDb.ConnectionString = GetConnectionStr(cszSystemMan)
' szExePrice = BusProjectExecutePrice(m_szProjectID, Now, lErrCode)
' If lErrCode = 1 Then Exit Function
' szSql = "Delete From bus_price_lst where bus_id='" & m_szBusID & "'and  price_table_id='" & szExePrice & "'"
' oDb.Execute szSql
' Dim szLogTemp As String
' szLogTemp = "发车时间[" & Format(m_dtOldStartupTime, cszTimeStr) & "终点站[" & m_szEndStationName & "]"
'
' WriteOperateLog m_oActiveUser, RIGHT_BusStationManagement, "由于线路改变原车次[" & m_szBusID & "]" & szLogTemp & "的线路票价信息被删除"
'Set oDb = Nothing
'End Function
Public Function BusVehicleStop(szVehicleID As String, dtEndStop As Date, dtStartStop As Date)
 Dim oDb As New RTConnection
 Dim szSql As String
 AssertObjIsValid
 AssertHaveRight m_oActiveUser, RIGHT_BusManagement
  oDb.ConnectionString = GetConnectionStr(cszSystemMan)
 szSql = "UPDATE bus_vehicle_code SET stop_start_date='" & ToDBDate(dtStartStop) & "' ,stop_end_date='" & ToDBDate(dtEndStop) & "'" _
         & " WHERE bus_id='" & m_szBusID & "' AND vehicle_id='" & szVehicleID & "'"
 oDb.Execute szSql
    Dim szLogTemp As String
   szLogTemp = "发车时间[" & Format(m_dtOldStartupTime, cszTimeStr) & "终点站[" & m_szEndStationName & "]"
   
 WriteOperateLog m_oActiveUser, RIGHT_BusManagement, "车次[" & m_szBusID & "]" & szLogTemp & "车辆[" & szVehicleID & "]在[" & ToDBDate(dtStartStop) & "到" & ToDBDate(dtEndStop) & "停班"

End Function


Public Function BusVehicleRun(szVehicleID As String, dtEndStop As Date, dtStartStop As Date)
 Dim oDb As New RTConnection
 Dim szSql As String
 AssertObjIsValid
 AssertHaveRight m_oActiveUser, RIGHT_BusManagement
  oDb.ConnectionString = GetConnectionStr(cszSystemMan)
 szSql = "UPDATE bus_vehicle_code SET stop_start_date='" & ToDBDate(dtStartStop) & "' ,stop_end_date='" & ToDBDate(dtEndStop) & "'" _
         & " WHERE bus_id='" & m_szBusID & "' AND vehicle_id='" & szVehicleID & "'"
 oDb.Execute szSql
     Dim szLogTemp As String
   szLogTemp = "发车时间[" & Format(m_dtOldStartupTime, cszTimeStr) & "终点站[" & m_szEndStationName & "]"
   
 WriteOperateLog m_oActiveUser, RIGHT_BusManagement, "车次[" & m_szBusID & "]" & szLogTemp & "车辆[" & szVehicleID & "]运行"

End Function
''**********************************************************
''* Project Name:STRglSch.vbp
''* Engineer:wjb
''* Data Generated:
''* Last Revision Date:2002/11/11
''* Relational Document:
''**********************************************************
''新增配载信息
'Public Sub AddAllotInfo(SellStationID As String, CheckGateID As String, StartUpTime As Date, CanSaleQuantity As Integer)
' Dim oDb As New RTConnection
' Dim szSql As String
' Dim rzTemp As Recordset
' On Error GoTo ErrHandle
'   AssertObjIsValid
'   AssertHaveRight m_oActiveUser, RIGHT_BusManagement
'   oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'   szSql = "SELECT * FROM sell_station_info WHERE sell_station_id='" & Trim(SellStationID) & "'"
'   Set rzTemp = oDb.Execute(szSql)
'   If rzTemp.RecordCount = 0 Then RaiseError ERR_NoSellStation
'   szSql = "SELECT * FROM bus_allot_info WHERE bus_id='" & Trim(BusID) & "'" _
'     & " AND project_id='" & Trim(ProjectID) & "'" & " AND sell_station_id='" & Trim(SellStationID) & "'"
'   Set rzTemp = oDb.Execute(szSql)
'   If rzTemp.RecordCount >= 1 Then RaiseError ERR_BusAllotAreadyExist
'   szSql = "INSERT bus_allot_info(" _
'       & "bus_id,project_id,sell_station_id,check_gate_id,bus_start_time,can_sale_quantity)" _
'       & "VALUES(" _
'       & "'" & BusID & "'," _
'       & "'" & ProjectID & "'," _
'       & "'" & SellStationID & "'," _
'       & "'" & CheckGateID & "'," _
'       & "'" & StartUpTime & "'," _
'       & " " & CanSaleQuantity & ")"
'   oDb.Execute (szSql)
'
'    Set oDb = Nothing
'
'ErrHandle:
'   AssertUpdateObjectError ERR_Bus, oDb
'End Sub
''修改配载信息
'Public Sub EditAllotInfo(SellStationID As String, CheckGateID As String, StartUpTime As Date, CanSaleQuantity As Integer)
'  Dim oDb As New RTConnection
'  Dim szSql As String
'  Dim rzTemp As Recordset
'  On Error GoTo ErrHandle
'   AssertObjIsValid
'   AssertHaveRight m_oActiveUser, RIGHT_BusManagement
'   oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'   szSql = "SELECT * FROM sell_station_info WHERE sell_station_id='" & Trim(SellStationID) & "' "
'   Set rzTemp = oDb.Execute(szSql)
'   If rzTemp.RecordCount = 0 Then RaiseError ERR_NoSellStation
'   szSql = "SELECT * FROM bus_allot_info WHERE bus_id='" & Trim(BusID) & "'" _
'     & " AND project_id='" & Trim(ProjectID) & "'" & " AND sell_station_id='" & Trim(SellStationID) & "'"
'   Set rzTemp = oDb.Execute(szSql)
'   If rzTemp.RecordCount >= 1 Then RaiseError ERR_BusAllotAreadyExist
'    szSql = "UPDATE bus_allot_info SET " _
'       & "bus_id='" & BusID & "'," _
'       & "project_id='" & ProjectID & "'," _
'       & "sell_station_id='" & SellStationID & "'," _
'       & "check_gate_id='" & CheckGateID & "'," _
'       & "bus_start_time='" & StartUpTime & "'," _
'       & "can_sale_quantity=" & Str(CanSaleQuantity)
'    oDb.Execute (szSql)
'
'    Set oDb = Nothing
'
'ErrHandle:
'    AssertUpdateObjectError ERR_Bus, oDb
'End Sub
''删除配载信息
'Public Sub DeleteAllotInfo(Optional SellStationID As String)
'  Dim oDb As New RTConnection
'  Dim szSql As String
'  Dim rzTemp As Recordset
'  On Error GoTo ErrHandle
'   AssertObjIsValid
'   AssertHaveRight m_oActiveUser, RIGHT_BusManagement
'   oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'   szSql = "SELECT * FROM sell_station_info WHERE sell_station_id='" & Trim(SellStationID) & "' "
'   Set rzTemp = oDb.Execute(szSql)
'   If rzTemp.RecordCount = 0 Then RaiseError ERR_NoSellStation
'   szSql = "SELECT * FROM bus_allot_info WHERE bus_id='" & Trim(BusID) & "'" _
'     & " AND project_id='" & Trim(ProjectID) & "'" & " AND sell_station_id='" & Trim(SellStationID) & "'"
'   Set rzTemp = oDb.Execute(szSql)
'   If rzTemp.RecordCount >= 1 Then RaiseError ERR_BusAllotAreadyExist
'   If SellStationID = "" Then
'     oDb.Execute ("DELETE bus_allot_info")
'   Else
'     szSql = "DELETE bus_allot_info WHERE SellStationID='" & SellStationID & "'"
'     oDb.Execute (szSql)
'   End If
'    Set oDb = Nothing
'
'ErrHandle:
'    AssertDeleteObjectError ERR_Bus, oDb
'End Sub



'查询车次配载信息
Public Function GetAllotInfo(Optional SellStationID As String) As TBusAllotInfo()
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rzTemp As Recordset
    Dim atResult() As TBusAllotInfo
      
    AssertObjIsValid
'    AssertHaveRight m_oActiveUser, RIGHT_BusManagement
    
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "SELECT a.*,b.sell_station_name,b.sell_station_full_name,c.check_gate_name" & _
            " FROM bus_allot_info a,sell_station_info b,Checkgate_info c" & _
            " WHERE a.bus_id=" & TransFieldValueToString(m_szBusID) & _
            " AND a.check_gate_id=c.check_gate_id AND a.sell_station_id=b.sell_station_id" & _
            " ORDER BY a.bus_start_time"
            
    If SellStationID <> "" Then
      szSql = szSql & " AND a.sell_station_id='" & SellStationID & "'"
    End If
    Set rzTemp = oDb.Execute(szSql)
    If rzTemp.RecordCount = 0 Then Exit Function
    ReDim atResult(1 To rzTemp.RecordCount)
    Dim i As Integer
    For i = 1 To rzTemp.RecordCount
        atResult(i).dtRunTime = FormatDbValue(rzTemp!bus_start_time)
        atResult(i).nCanSellQuantity = FormatDbValue(rzTemp!can_sale_quantity)
        atResult(i).szBusID = FormatDbValue(rzTemp!BUS_ID)
        atResult(i).szCheckGateID = FormatDbValue(rzTemp!check_gate_id)
        atResult(i).szCheckGateName = FormatDbValue(rzTemp!check_gate_name)
'        atResult(i).szProjectID = FormatDbValue(rzTemp!project_id)
        atResult(i).szSellStationID = FormatDbValue(rzTemp!sell_station_id)
        atResult(i).szSellStationName = FormatDbValue(rzTemp!sell_station_name)
        rzTemp.MoveNext
    Next i
    
    GetAllotInfo = atResult
    Set oDb = Nothing
End Function

Public Sub SaveAllot(patAllotInfo() As TBusAllotInfo)
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_BusManagement
    
    Dim oDb As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr("")
    Dim i As Integer, nCount As Integer
    Dim nCanSell As Integer     '
    
    nCount = ArrayLength(patAllotInfo)
    For i = 1 To nCount
        nCanSell = nCanSell + patAllotInfo(i).nCanSellQuantity
        '验证检票口是否和售票站一致
        szSql = "SELECT 1 FROM Checkgate_info " & _
                " WHERE sell_station_id=" & TransFieldValueToString(patAllotInfo(i).szSellStationID) & _
                " AND check_gate_id=" & TransFieldValueToString(patAllotInfo(i).szCheckGateID)
        If oDb.Execute(szSql).RecordCount = 0 Then RaiseError ERR_CheckGateNotMaskStation
    Next i

On Error GoTo ErrHandle

    oDb.BeginTrans
        szSql = "DELETE bus_allot_info WHERE bus_id=" & TransFieldValueToString(m_szBusID) 'FL & " AND project_id=" & TransFieldValueToString(m_szProjectID)
        oDb.Execute szSql
        For i = 1 To nCount
            szSql = "INSERT bus_allot_info (bus_id,sell_station_id,check_gate_id,bus_start_time,can_sale_quantity)" & _
                    " VALUES(" & TransFieldValueToString(m_szBusID) & "," & _
                    TransFieldValueToString(patAllotInfo(i).szSellStationID) & "," & TransFieldValueToString(patAllotInfo(i).szCheckGateID) & "," & _
                    TransFieldValueToString(patAllotInfo(i).dtRunTime) & "," & TransFieldValueToString(patAllotInfo(i).nCanSellQuantity) & ")"
            oDb.Execute szSql
        Next i
        
        '如果只修改配载站的检票口,不修改车次信息的检票口。
        '这样的话,如果修改的检票口正好当前车次信息表的检票口的话，会导致车次关联不到。（新的配载站权限中用到关联）
        '所以，需要加上以下的语句，将车次信息表中的相应信息也进行修改。
        
        '将车次表中的发车时间与配载的发车时间进行对应
        szSql = " UPDATE bus_info " _
            & " SET bus_start_time = a.bus_start_time , check_gate_id = a.check_gate_id  " _
            & " FROM bus_info b , checkgate_info c , bus_allot_info a " _
            & " WHERE  b.check_gate_id = c.check_gate_id " _
            & " AND b.bus_id = a.bus_id " _
            & " AND a.sell_station_id  = c.sell_station_id " _
            & " AND b.bus_id = " & TransFieldValueToString(m_szBusID)
            
        oDb.Execute szSql
    
    oDb.CommitTrans
    Exit Sub
ErrHandle:
    oDb.RollbackTrans
    RaiseError err.Number, , err.Description
End Sub


'查询车次售票点信息
Public Function GetSellStationInfo(Optional SellStationID As String) As TBusAllotInfo()
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rzTemp As Recordset
    Dim atResult() As TBusAllotInfo
      
    AssertObjIsValid
'    AssertHaveRight m_oActiveUser, RIGHT_BusManagement
    
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "SELECT a.*,b.sell_station_name,b.sell_station_full_name " & _
            " FROM bus_sell_station_lst a , sell_station_info b " & _
            " WHERE a.bus_id=" & TransFieldValueToString(m_szBusID) & _
            " AND a.sell_station_id=b.sell_station_id " & _
            " ORDER BY a.sell_station_id "
    
    If SellStationID <> "" Then
      szSql = szSql & " AND a.sell_station_id='" & SellStationID & "'"
    End If
    Set rzTemp = oDb.Execute(szSql)
    If rzTemp.RecordCount = 0 Then Exit Function
    ReDim atResult(1 To rzTemp.RecordCount)
    Dim i As Integer
    For i = 1 To rzTemp.RecordCount
    
        atResult(i).nCanSellQuantity = FormatDbValue(rzTemp!can_sale_quantity)
        atResult(i).szBusID = FormatDbValue(rzTemp!BUS_ID)
        atResult(i).szSellStationID = FormatDbValue(rzTemp!sell_station_id)
        atResult(i).szSellStationName = FormatDbValue(rzTemp!sell_station_name)
                
        rzTemp.MoveNext
    Next i
    
    GetSellStationInfo = atResult
    Set oDb = Nothing
    
    
End Function



Public Sub SaveSellStationInfo(patAllotInfo() As TBusAllotInfo)
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_BusManagement
    
    Dim oDb As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr("")
    Dim i As Integer, nCount As Integer

    nCount = ArrayLength(patAllotInfo)
    On Error GoTo ErrHandle
    oDb.BeginTrans
        szSql = "DELETE bus_sell_station_lst WHERE bus_id=" & TransFieldValueToString(m_szBusID)
        oDb.Execute szSql
        For i = 1 To nCount
            szSql = "INSERT bus_sell_station_lst (bus_id,sell_station_id,can_sale_quantity) VALUES( " _
                    & TransFieldValueToString(m_szBusID) & "," _
                    & TransFieldValueToString(patAllotInfo(i).szSellStationID) & "," _
                    & TransFieldValueToString(patAllotInfo(i).nCanSellQuantity) & ")"
            oDb.Execute szSql
        Next i
        
    oDb.CommitTrans
    Exit Sub
ErrHandle:
    oDb.RollbackTrans
    RaiseError err.Number, , err.Description
End Sub



'**************************************************
'Member Code:
'Brief Description:获得车次的全部运行公司
'Engineer:
'Date Generated:
'Last Revision Date:2005-11-01
'**************************************************
Public Function GetAllCompany() As String()
Dim oDb As New RTConnection
Dim rsTemp As Recordset
Dim aszInfo() As String
Dim i As Integer
Dim szSql As String
AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "SELECT distinct v.transport_company_id , c.transport_company_short_name  " _
        & " FROM bus_vehicle_code b,Vehicle_info v , company_info c  " _
        & " WHERE bus_id='" & m_szBusID & "' AND  b.vehicle_id= v.vehicle_id and v.transport_company_id = c.transport_company_id "
    Set rsTemp = oDb.Execute(szSql)
    
    '车次无运行车辆
    If rsTemp.RecordCount <> 0 Then
    ReDim aszInfo(1 To rsTemp.RecordCount, 1 To 2) As String
    For i = 1 To rsTemp.RecordCount
        aszInfo(i, 1) = FormatDbValue(rsTemp!transport_company_id)
        aszInfo(i, 2) = FormatDbValue(rsTemp!transport_company_short_name)
        rsTemp.MoveNext
    Next
    GetAllCompany = aszInfo()
    End If
End Function
