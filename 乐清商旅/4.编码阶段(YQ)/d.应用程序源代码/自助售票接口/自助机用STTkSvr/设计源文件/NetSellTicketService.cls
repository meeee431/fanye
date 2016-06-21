VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetSellTicketService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'SellTicketService类模块文件:
Option Explicit

Const cszStandTicketSeatNo = "ST"
Const cszAutoSeatNo = ""


Private Enum ETicketServiceType
    TP_CancelTicket = 1 '废票
    TP_ReturnTicket = 2 '退票
    TP_ChangeTicket = 3 '改签
End Enum
Const cszNOPrint = 0
Const cszEmptyDate = "0:00:00"

Private bIsUsingRelationBus As Boolean
'Const cszErr_NoNetBook = 2001 '无网上售票记录!
'Const cszErr_NoGetNet = 2002 '网上订票还未取!不能作废!
'Const cszErr_NetBookCancel = 2000 '已作废!
'Const cszErr_OverNetBook = 2003 '取消的订票数超出已预订的!
'此处定义错误号：
Public Enum EErrNetSellTicketService
'    ERR_SellTicketServiceNoActiveUser = ERR_SellTicketService + ERR_NoActiveUser
'
'    ERR_BusStopped = ERR_SellTicketService + 21 '车次已经停班
'    ERR_BusMerged = ERR_SellTicketService + 22 '车次已经被并班停班
'    ERR_BusStopCheck = ERR_SellTicketService + 23 '车次已经停检
'    ERR_BusStopSellTicket = ERR_SellTicketService + 24 '车次已经停售
'    ERR_BusStaionStopSellTicket = ERR_SellTicketService + 25 '车次还未到指定到站的售票时间
'    ERR_BusReachMaxStationTicketCount = ERR_SellTicketService + 26 '超过指定到站的限售票数
'    ERR_BusReachMaxTicketCount = ERR_SellTicketService + 27 '指定车次的座位数已经不够
'    ERR_SeatNoCanNotSell = ERR_SellTicketService + 28 '指定的座位号不能出售（已被售出或预留）
'    ERR_CanNotSellStandTicketWhenHaveSeat = ERR_SellTicketService + 29 '还有座位时不能售站票
'    ERR_NoStandTicket = ERR_SellTicketService + 30 '已无站票
'    ERR_NoThisBus = ERR_SellTicketService + 31 '指定的车次不存在或指定的车次不过指定的车站
'    ERR_BusNotInIlegalDate = ERR_SellTicketService + 32 '指定的车次的日期不在法范围内(当天到预售天之内)
'    ERR_NoTicketInfoSpecified = ERR_SellTicketService + 33 '没有指定何购票信息
'
'    ERR_TicketIsCanceled = ERR_SellTicketService + 34 '此票已经被废
'    ERR_TicketIsReturned = ERR_SellTicketService + 35 '此票已经被退
'    ERR_TicketIsChecked = ERR_SellTicketService + 36 '此票已经被检
'    ERR_TicketIsChangeTicket = ERR_SellTicketService + 37 '此票是改签售出的票,不允许作废,只能作退票处理
'
'    ERR_UserIsNotTicketSeller = ERR_SellTicketService + 38 '操作员不是指定票的售票员,而正常废票需由本人进行
'    ERR_CancelTicketTimeOut = ERR_SellTicketService + 39 '已过正常废票时间
'    ERR_TicketNotExist = ERR_SellTicketService + 40 '指定的票不存在
'    ERR_TicketIsChanged = ERR_SellTicketService + 41 '此票已被改签
'    ERR_ReturnTicketTimeOut = ERR_SellTicketService + 42 '已过正常退票时间
'    ERR_FreeTicketCanNotReturnChange = ERR_SellTicketService + 43 '免票不能退,改签
'    ERR_OrgReturnChargeOutDated = ERR_SellTicketService + 44 '原先得到的退票费已过时
'    ERR_InputParamTypeErr = ERR_SellTicketService + 45 '输入的参数类型不合要求
'    ERR_ScrollBusCanNotExtraSell = ERR_SellTicketService + 46 '流水车次不能补票
'    ERR_TicketNoDup = ERR_SellTicketService + 47 '票号重
'    ERR_ScrollBusTicketCanNotReturn = ERR_SellTicketService + 48 '流水车次不能正常退票
'    ERR_ScroolBusReturnTicketDateOut = ERR_SellTicketService + 49 '流水车次的退票只能在当天进行
'    ERR_WriteOffTiecketFail = ERR_SellTicketService + 50 '注销票失败
'    ERR_NoExistWriteOffTiecket = ERR_SellTicketService + 51 '无此注销票
'    ERR_ForceCancelTicketRemoteNotAllowed = ERR_SellTicketService + 52 '强行废票不允许远程
'    ERR_BusNotStopSellTicket = ERR_SellTicketService + 53 '还未到补票时间,不能补票
'    ERR_TicketTypeNotDefined = ERR_SellTicketService + 54
'
'    ERR_TicketPriceOutOfLimit = ERR_SellTicketService + 55 '所售票的票价超出折扣票允许的范围
'
'    ERR_BusSlitpStop = ERR_SellTicketService + 56 '拆分停班车次不能售票
'
'    ERR_ReturnTicketNotResume = ERR_SellTicketService + 57 '车票已售完不能取消退票
'    ERR_CancelTicketNotResume = ERR_SellTicketService + 58 '车票已售完不能取消废票
'    ERR_BusIsSplit = ERR_SellTicketService + 59 '车次正在拆分不能售票
'
'    '定义增加错误号留接口
'    ERR_SellQuantityISFull = ERR_SellTicketService + 60 '12460 售票站超过分配的售票张数
'    ERR_ErrOther2 = ERR_SellTicketService + 61
'    ERR_ErrOther3 = ERR_SellTicketService + 63
'
'
'    ERR_UserCannotSell = ERR_SellTicketService + 64 '12464 用户不被允许售当天的票
'    ERR_BusAllotInfoNotExist = ERR_SellTicketService + 65 '12465 售的上车站在该车次的配载信息中不存在
'
'
'    ERR_BusSellStationInfoNotExist = ERR_SellTicketService + 66 '12466 售票员所属车站在该车次的售票点信息中不存在
'    ERR_CannotSellZeroPrice = ERR_SellTicketService + 67 '12467 票价为0的票不允许出售
    
    ERR_NoNetBook = ERR_SellTicketService + 80 '无网上售票记录!
    ERR_NoGetNet = ERR_SellTicketService + 81 '网上订票还未取!不能作废!
    ERR_NetBookCancel = ERR_SellTicketService + 82 '已作废!
    ERR_OverNetBookNum = ERR_SellTicketService + 83 '取消的订票数超出已预订的!
    ERR_UnReturnNumAndSeatNum = ERR_SellTicketService + 84 '所要废的张数与座位数不相等!
End Enum

''此处定义权限号：
'Public Enum ERightNetSellTicketService
'    RIGHT_SellTicket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 1 '售票
'    RIGHT_ExtraSellTicket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 2 '补票
'    RIGHT_CancelTicket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 3 '废票
'    RIGHT_ForceCancelTicket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 4 '强行废票
'    RIGHT_ReturnTicket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 5 '退票
'    RIGHT_ForceReturnTicket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 6 '强行退票
'    RIGHT_ChangeTicket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 7 '改签车票
'    RIGHT_QuerySelfThings = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 8 '查询自己的票处理明细
'    RIGHT_QueryOtherThings = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 9 '查询他人的票处理明细
'    RIGHT_WriteOffTiecket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 10 '注销票权限
'
'    RIGHT_SellDiscountTicket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 11 '售折扣票
'    RIGHT_SellOrderSeat = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 12     '售定座位号
'    RIGHT_ResumeReturnTicket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 13 '取消退票
'    RIGHT_ResumeCancelTicket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 14 '取消废票
'    RIGHT_SelectReturnRatio = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 15 '选择退票废率
'
'    定义增加权限接口
'    RIGHT_RightOther1 = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 16
'    RIGHT_RightOther2 = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 17  '售计划预定票
'    RIGHT_RightOther3 = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 18
'    RIGHT_AllMoneyReturn = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 19 '全额退票
'End Enum

'所有的类都有的私有类变量
Private m_oActiveUser As ActiveUser
Private m_oParam As New SystemParam
Private m_bSellStationCanSellEachOther As Boolean
Private m_bAllowScrollBusSaleForever As Boolean
Const SeatTypeID = "01"
'--------------------------->>
Public Sub Init(poAUser As ActiveUser)
    Set SelfUser = poAUser
    m_oParam.Init SelfUser
    m_bSellStationCanSellEachOther = m_oParam.SellStationCanSellEachOther
    m_bAllowScrollBusSaleForever = m_oParam.AllowScrollBusSaleForever


End Sub

Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
    m_oActiveUser.IAmLive
End Property

Public Property Set SelfUser(vNewValue As ActiveUser)
    Set m_oActiveUser = vNewValue
    m_oActiveUser.IAmLive
End Property

Public Property Get StationID() As String
    Dim oParam As SystemParam
    Set oParam = New SystemParam
    oParam.Init m_oActiveUser
    StationID = oParam.StationID
End Property

Public Property Get StationName() As String
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    
    szSql = "SELECT station_name FROM station_info WHERE " _
    & " station_id='" & StationID & "'"
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 1 Then
        StationName = FormatDbValue(rsTemp!station_name)
    Else
        szSql = "SELECT unit_short_name FROM connect_unit_info WHERE " _
        & " unit_id='" & m_oActiveUser.ServerUnit & "'"
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount = 1 Then
            StationName = FormatDbValue(rsTemp!unit_short_name)
        End If
    End If
End Property

'得到当前系统中的所有站点信息
Public Function GetAllStationRs() As Recordset
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
'    szsql = "SELECT station_id,station_name ,station_input_code FROM station_info ORDER BY station_id"
    szSql = "select a.station_id,s.station_name,s.station_input_code from (select e.station_id  from work_env_bus_station_lst e  group by e.station_id) a,station_info s where a.station_id=s.station_id and a.station_id not in (select station_id from sell_station_info) "
    Set rsTemp = oDb.Execute(szSql)
    Set GetAllStationRs = rsTemp
End Function
'售票
Public Function SellTicket(BusStartupDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String, aspSellTicketParam() As TSellTicketParam, anInsurance() As Integer, Optional oDb As RTConnection) As TSellTicketResult()

    Dim tsTemp() As TSellTicketResult
    Dim iLen As Integer
    Dim nCount As Integer
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_SellTicket
    iLen = ArrayLength(BusID)
    oDb.BeginTrans
        On Error GoTo here
        ReDim tsTemp(1 To iLen)
        For nCount = 1 To iLen
            tsTemp(nCount) = SelfSellTicket(BusStartupDate(nCount), BusID(nCount), SellStationID(nCount), DesStationID(nCount), aspSellTicketParam(nCount).BuyTicketInfo, oDb, aspSellTicketParam(nCount).pasgSellTicketPrice, anInsurance(nCount))
        Next nCount
        SellTicket = tsTemp
    oDb.CommitTrans
    
'    On Error GoTo 0
'    WriteOperateLog m_oActiveUser, RIGHT_SellTicket
    
    Exit Function
here:
    oDb.RollbackTrans
    err.Raise err.Number
End Function

'补票
Public Function ExtraSellTicket(BusStartupDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String, aspSellTicketParam() As TSellTicketParam) As TSellTicketResult()
    Dim oDb As New RTConnection
    Dim iLen As Integer
    Dim nCount As Integer
    Dim rsTemp() As TSellTicketResult
    Dim i As Integer
   
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_ExtraSellTicket
    
'    Dim asiTemp(1 To 1) As TBuyTicketInfo
'    asiTemp(1) = SellTicketInfo
    iLen = 0
    iLen = ArrayLength(BusID)
    ReDim rsTemp(1 To iLen)
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    oDb.BeginTrans
    On Error GoTo here
    For nCount = 1 To iLen
    
        rsTemp(nCount) = SelfSellTicket(BusStartupDate(nCount), BusID(nCount), SellStationID(nCount), DesStationID(nCount), aspSellTicketParam(nCount).BuyTicketInfo, oDb, aspSellTicketParam(nCount).pasgSellTicketPrice, True, ST_TicketReturned)
    
    Next nCount
    ExtraSellTicket = rsTemp
    For nCount = 1 To iLen
        For i = 1 To ArrayLength(aspSellTicketParam(nCount).BuyTicketInfo)
           WriteOperateLog m_oActiveUser, RIGHT_ExtraSellTicket, "补票票号：" & aspSellTicketParam(nCount).BuyTicketInfo(i).szTicketNo
        Next i
    Next nCount
    oDb.CommitTrans
    
    On Error GoTo 0
    
    Exit Function
here:
    oDb.RollbackTrans
    err.Raise err.Number
End Function


Private Function UpdateSeatStatusStr(pszBusID As String, pszDate As Date, pszSeatNo As String, pszTicketID As String) As String
    Dim szSql As String
    
    szSql = "UPDATE Env_bus_seat_lst SET status=" & ST_SeatSold & "," _
    & " ticket_no='" & pszTicketID & "' " _
    & " WHERE bus_id='" & pszBusID & "' AND " _
    & " bus_date='" & ToDBDate(pszDate) & "' AND " _
    & " seat_no='" & pszSeatNo & "' AND " _
    & " status=" & ST_SeatCanSell
    UpdateSeatStatusStr = szSql
End Function

'内部售票函数  票的折扣已删除 2002/11/13 wjb
Private Function SelfSellTicket(BusStartupDate As Date, BusID As String, SellStationID As String, DesStationID As String, _
                                SellTicketInfo() As TBuyTicketInfo, poDb As RTConnection, _
                                pasgSellTicketPrice() As Single, pnInsurance As Integer, _
                                Optional pbExtraSell As Boolean = False, _
                                Optional pnTicketStatus As ETicketStatus = ST_TicketNormal) As TSellTicketResult
    
    'psgDiscount票的折扣
    
    Dim oParam As New SystemParam
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim dtBusFullDateTime As Date
    Dim nTemp As Double
    Dim dtNow As Date
    Dim i As Integer
    Dim nBusStatus As EREBusStatus
    Dim nBuyTicketCount As Integer
    Dim sgFullPrice As Double, sgHalfPrice As Double
    Dim lRowAffect As Long, nStationLimitCount As Integer
    Dim szSeatNo As String, sgTicketPrice As Double
    Dim nBusType As EBusType
    
    Dim atiTicketPrice(TP_TicketTypeCount) As TRETicketPrice
    
    Dim siSellInfo As TSellTicketResult
    Dim lErrorCode As Long
    Dim rsRange As Recordset
    
    Dim sgPreferentialTicket1 As Single '折扣票1
    Dim sgPreferentialTicket2 As Single '折扣票2
    Dim sgPreferentialTicket3 As Single '折扣票3
        
    Dim bSellDiscountTicket As Boolean '所卖的是否是折扣票
    
    Dim bNeedBook As Boolean '是否上了预定模块
    Dim bBookValid As Boolean '是否还需判断预定
    Dim bSeatBooked As Boolean
    Dim szBookStationID As String '预定的到站
    Dim rsTemp2 As Recordset
    
    Dim szSqlSeat As String '得到相应票价
    Dim szSeatType As String '座位类型
    Dim nSeatStatus As Integer
    Dim nSellTicketType As Integer
    Dim nSellTicketStatus As ETicketStatus
    
    Dim szStopTime As String
    
    Dim bSellStationCanSellEachOther As Boolean '是否允许互联售票
    
    nSellTicketStatus = pnTicketStatus
    
    Select Case pnTicketStatus
        Case ST_TicketReturned
            nSellTicketType = TP_ExtraSold
            nSellTicketStatus = ST_TicketNormal
        Case ST_TicketSellChange
            nSellTicketType = TP_ChangeSold
        Case ST_TicketNormal
            nSellTicketType = TP_NormalSold
        Case Else
            nSellTicketType = 0
    End Select
    
    bNeedBook = False
    
    oParam.Init m_oActiveUser
    '判断车次的日期是否超过预售期限
    
    dtNow = oParam.NowDateTime
    bSellStationCanSellEachOther = oParam.SellStationCanSellEachOther
    
    nTemp = DateDiff("d", dtNow, BusStartupDate)
    
    '如果要购票的车次不在当天和预售天之内则出错
    If nTemp > oParam.PreSellDate Or nTemp < 0 Then ShowError ERR_BusNotInIlegalDate
    
    If nTemp > oParam.UserCanSellDate Then ShowError ERR_UserCannotSell
    'poDb.ConnectionString = GetConnectionStr(cszSellTicket)
    '得到环境车次及车次站点信息时
'    szSql = "SELECT * FROM Work_env_bus_info,Work_env_bus_station_lst " _
'    & " WHERE Work_env_bus_info.bus_id='" & BusID & "' AND " _
'    & " Work_env_bus_info.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
'    & " Work_env_bus_info.bus_id=Work_env_bus_station_lst.bus_id AND " _
'    & " Work_env_bus_info.bus_date=Work_env_bus_station_lst.bus_date AND " _
'    & " station_id='" & DesStationID & "'" _
'    & " AND Work_env_bus_station_lst.sell_station_id='" & SellStationID & "'"   '增加上车代码
    
     szSql = "SELECT e.bus_date , e.bus_id , e.vehicle_type_name , e.vehicle_type_code , e.total_seat " _
        & " , e.sale_seat_quantity , e.total_stand_quantity , e.sale_stand_seat_quantity , e.bus_start_time , e.bus_kind" _
        & " ,  e.is_all_refundment , e.route_id , e.end_station_id , e.end_station_name" _
        & " , e.owner_id , e.transport_company_id , e.vehicle_id , e.split_company_id , e.baggage_number , e.fact_weight" _
        & " , e.calculate_weight , e.over_weight_number , e.luggage_total_price  , e.scrollbus_check_time" _
        & " , e.seat_remain , e.bed_remain  , e.additional_remain , e.other_remain_1 , e.other_remain_2" _
        & " , s.sell_station_id , s.station_id , s.seat_type_id , s.station_serial_no , s.max_sale_quantity , s.sale_ticket_quantity" _
        & " , s.book_quantity , s.stop_sale_time , s.mileage  , s.full_price , s.half_price , s.preferential_ticket1 , s.preferential_ticket2" _
        & " , s.preferential_ticket3" _
        & " , a.check_gate_id , a.bus_start_time" _
        & " , a.status , a.can_sale_quantity , a.sell_quantity" _
        & " FROM Work_env_bus_info e ,Work_env_bus_station_lst s , Work_env_bus_allot_lst a" _
        & " WHERE e.bus_id='" & BusID & "' AND  e.bus_date='" & ToDBDate(BusStartupDate) & "'" _
        & " AND  e.bus_id=s.bus_id AND  e.bus_date=s.bus_date" _
        & " AND  s.station_id='" & DesStationID & "' AND s.sell_station_id='" & SellStationID & "'" _
        & " AND a.bus_id = s.bus_id AND a.bus_date = s.bus_date AND a.sell_station_id = '" & SellStationID & "'" _

    
    Set rsTemp = poDb.Execute(szSql)
    
    
    If rsTemp.RecordCount = 0 Then ShowError ERR_NoThisBus '没有指定的车次或车次不过指定的到站

    nBusType = FormatDbValue(rsTemp!bus_kind)
    
    '流水车次不能补票
    If nBusType = TP_ScrollBus And pbExtraSell Then ShowError ERR_ScrollBusCanNotExtraSell
    nBusStatus = FormatDbValue(rsTemp!Status)
    
    '并班车次不能售票
    If nBusStatus = ST_BusMergeStopped Then ShowError ERR_BusMerged
    
     '拆分停班车次不能售票
    If nBusStatus = ST_BusSlitpStop Then ShowError ERR_BusSlitpStop
    
    '车次正在拆分不能售票
    If nBusStatus >= 32 Then ShowError ERR_BusIsSplit
'
'    '停检之后不能售票（实际不太可能出现，除非停售时间设得太小了）
'    If pbExtraSell And nBusStatus = ST_BusStopCheck Then ShowError ERR_BusStopCheck
    If nBusStatus = ST_BusStopCheck Then ShowError ERR_BusStopCheck
    
    If nBusStatus = ST_BusStopped Then ShowError ERR_BusStopped
    
    dtBusFullDateTime = CDate(Format(FormatDbValue(rsTemp!bus_date), cszDateStr) & " " & Format(FormatDbValue(rsTemp!bus_start_time), cszTimeStr))
    
    '判断是否到了停售时间(如果是补票则要过了正常售票时间才能补)
    
    If nBusType <> TP_ScrollBus Then '如果是固定车次才限制
        If pbExtraSell Then
            If DateDiff("s", dtNow, dtBusFullDateTime) >= oParam.StopSellTime * 60 Then ShowError ERR_BusNotStopSellTicket
        Else
'            If DateDiff("s", dtNow, dtBusFullDateTime) < oParam.StopSellTime * 60 Then ShowError ERR_BusStopSellTicket
        End If
    End If
    
    szStopTime = FormatDbValue(rsTemp!stop_sale_time)
    
   If nBusType <> TP_ScrollBus Then '如果是固定车次才限制
        If szStopTime <> "-1" Then '如果该车次到指定的站限制售票时间则要进行判断
            If DateDiff("n", dtNow, dtBusFullDateTime) > GetStopTime(szStopTime) Then ShowError ERR_BusStaionStopSellTicket, "SellTicketService.SelfSellTicket"
        End If
        bNeedBook = oParam.IsBookValid
        If bNeedBook Then
            bBookValid = IIf(DateAdd("n", oParam.BookTime, dtNow) >= rsTemp!bus_start_time, False, True)
        End If
   Else
        If szStopTime <> "-1" Then
           dtBusFullDateTime = CDate(Format(FormatDbValue(rsTemp!bus_date), cszDateStr) & " " & GetStopTime(szStopTime, False))
           If DateDiff("n", dtNow, dtBusFullDateTime) < 0 Then ShowError ERR_BusStaionStopSellTicket, "SellTicketService.SelfSellTicket"
        End If
   End If
    
    nBuyTicketCount = ArrayLength(SellTicketInfo)
    
    If nBuyTicketCount > 0 Then '要售票的张数大于0
        
        If nBuyTicketCount > FormatDbValue(rsTemp!sale_seat_quantity) Then ShowError ERR_BusReachMaxStationTicketCount
        
        nStationLimitCount = FormatDbValue(rsTemp!max_sale_quantity)
        If nStationLimitCount > 0 Then
            If nBuyTicketCount > nStationLimitCount Then ShowError ERR_BusReachMaxStationTicketCount
        End If
        
        sgFullPrice = FormatDbValue(rsTemp!full_price)
        sgHalfPrice = FormatDbValue(rsTemp!half_price)
        sgPreferentialTicket1 = FormatDbValue(rsTemp!preferential_ticket1)
        sgPreferentialTicket2 = FormatDbValue(rsTemp!preferential_ticket2)
        sgPreferentialTicket3 = FormatDbValue(rsTemp!preferential_ticket3)
        
        ReDim siSellInfo.asgTicketPrice(1 To nBuyTicketCount)
        ReDim siSellInfo.aszSeatNO(1 To nBuyTicketCount)
        ReDim siSellInfo.aszTicketType(1 To nBuyTicketCount)
        ReDim siSellInfo.aszSeatType(1 To nBuyTicketCount)
        
        
        For i = 1 To nBuyTicketCount
            
            szSeatNo = ""
            Select Case SellTicketInfo(i).nTicketType
            Case TP_FullPrice, TP_FreeTicket
                sgTicketPrice = sgFullPrice
            Case TP_HalfPrice
                sgTicketPrice = sgHalfPrice
            Case TP_PreferentialTicket1
                sgTicketPrice = sgPreferentialTicket1
            Case TP_PreferentialTicket2
                sgTicketPrice = sgPreferentialTicket2
            Case TP_PreferentialTicket3
                sgTicketPrice = sgPreferentialTicket3
            End Select
             
            If nBusType <> TP_ScrollBus Then '如果是固定车次
                
                If SellTicketInfo(i).szSeatNo <> cszStandTicketSeatNo Then
                
                    If SellTicketInfo(i).szSeatNo = cszAutoSeatNo Then
                        '系统自动给出座位号
RetrySeatNo:
                        If bNeedBook And Not bBookValid Then        '此处加了座位类型,加上上车站代码条件
                            szSql = "SELECT s.seat_no AS SeatNo,s.status ,s.seat_type_id AS SeatType,p.ticket_price_total AS TicketPrice  FROM Env_bus_seat_lst s ,Work_env_bus_price_lst p" _
                            & " WHERE " _
                            & " s.bus_id='" & BusID & "' AND " _
                            & " s.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & " p.sell_station_id='" & SellStationID & "' AND " _
                            & " s.seat_no=" _
                            & "(SELECT MIN(seat_no) FROM Env_bus_seat_lst WHERE " _
                            & " bus_id='" & BusID & "' AND " _
                            & " bus_date='" & ToDBDate(BusStartupDate) & "' AND (" _
                            & " status=" & ST_SeatCanSell & " OR status=" & ST_SeatBooked & " or status=" & ST_SeatProjectBooked & ") AND seat_type_id='" & SellTicketInfo(i).szSeatTypeID & "' ) AND " _
                            & "s.bus_id=p.bus_id AND " _
                            & "s.bus_date=p.bus_date AND "
                            szSql = szSql & "s.seat_type_id=p.seat_type_id AND " _
                                        & "p.station_id='" & DesStationID & "' AND " _
                                        & "p.ticket_type='" & SellTicketInfo(i).nTicketType & "'"
                            
                        Else
                            szSql = "SELECT s.seat_no AS SeatNo,s.seat_type_id AS SeatType,s.status,p.ticket_price_total AS TicketPrice  FROM Env_bus_seat_lst s, Work_env_bus_price_lst p WHERE " _
                            & " s.bus_id='" & BusID & "' AND " _
                            & " s.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & " p.sell_station_id='" & SellStationID & "' AND " _
                            & "s.seat_no=(SELECT MIN(seat_no) FROM Env_bus_seat_lst WHERE bus_id='" & BusID & "' AND " _
                            & "bus_date='" & ToDBDate(BusStartupDate) & "' AND status=" & ST_SeatCanSell & " AND seat_type_id='" & SellTicketInfo(i).szSeatTypeID & "') AND " _
                            & "s.bus_id=p.bus_id AND " _
                            & "s.bus_date=p.bus_date AND "
                            szSql = szSql & "s.seat_type_id=p.seat_type_id AND " _
                                        & "p.station_id='" & DesStationID & "' AND " _
                                        & "p.ticket_type='" & SellTicketInfo(i).nTicketType & "'"

                        End If
                        
                        Set rsTemp = poDb.Execute(szSql)
                        If rsTemp.RecordCount = 0 Then ShowError ERR_BusReachMaxTicketCount, "SellTicket", "或" & SellTicketInfo(i).szSeatTypeName & "座位已售完。"
                        
                        szSeatNo = FormatDbValue(rsTemp!SeatNo) '赋座位号
                        sgTicketPrice = rsTemp!TicketPrice  '赋该座位票价
                        szSeatType = FormatDbValue(rsTemp!SeatType)
                        nSeatStatus = rsTemp!Status
                        
                        If bNeedBook And Not bBookValid Then
                            bSeatBooked = IIf(rsTemp!Status = ST_SeatBooked, True, False)
'                            If rsTemp!Status = ST_SeatProjectBooked Then
'                                AssertHaveRight m_oActiveUser, RIGHT_RightOther1  '判断是否有计划预定权限
'                            End If
                            
                        Else
                            bSeatBooked = False
                        End If
                        If bSeatBooked Then
                            szSql = UpdateSeatStatusStr2(BusID, BusStartupDate, szSeatNo, SellTicketInfo(i).szTicketNo, bBookValid)
                        Else
                            szSql = UpdateSeatStatusStr(BusID, BusStartupDate, szSeatNo, SellTicketInfo(i).szTicketNo)
                        End If
                        If rsTemp!Status = ST_SeatProjectBooked Then szSql = UpdateSeatStatusStr3(BusID, BusStartupDate, szSeatNo, SellTicketInfo(i).szTicketNo)
                        poDb.Execute szSql, lRowAffect
                        
                        If lRowAffect = 0 Then GoTo RetrySeatNo
                        
                    Else
                        '否则
                        szSeatNo = SellTicketInfo(i).szSeatNo
                        '取座位该票价和类型
                        szSqlSeat = "SELECT p.ticket_price_total AS TicketPrice, p.seat_type_id AS SeatType ,s.status AS SeatStatus FROM Work_env_bus_price_lst p,Env_bus_seat_lst s WHERE " _
                                        & "p.bus_id='" & BusID & "' AND p.station_id='" & DesStationID & "' AND  " _
                                        & "p.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                                        & "p.ticket_type=" & SellTicketInfo(i).nTicketType & " AND " _
                                        & "p.seat_type_id =(SELECT seat_type_id FROM Env_bus_seat_lst WHERE " _
                                        & "bus_id='" & BusID & "' AND bus_date='" & ToDBDate(BusStartupDate) & "' AND seat_no='" & szSeatNo & "')" _
                                        & " AND p.bus_date = s.bus_date AND p.bus_id = s.bus_id AND s.seat_no='" & szSeatNo & "'" _
                                        & " AND p.sell_station_id ='" & SellStationID & "'" _
                                        & " AND p.bus_id = s.bus_id AND p.bus_date = s.bus_date AND s.seat_no = '" & szSeatNo & "'"
                                        
                        Set rsTemp = poDb.Execute(szSqlSeat)
                        
                        If rsTemp.RecordCount = 0 Then ShowError ERR_BusReachMaxTicketCount
                        
                        sgTicketPrice = rsTemp!TicketPrice
                        szSeatType = FormatDbValue(rsTemp!SeatType)
                        nSeatStatus = rsTemp!seatstatus
                        szSql = UpdateSeatStatusStr(BusID, BusStartupDate, SellTicketInfo(i).szSeatNo, SellTicketInfo(i).szTicketNo)
                        poDb.Execute szSql, lRowAffect
                        If lRowAffect = 0 Then
                            If bNeedBook Then
                                If nSeatStatus = ST_SeatProjectBooked Then
                                    AssertHaveRight m_oActiveUser, RIGHT_RightOther1  '判断计划预定权限
                                    szSql = UpdateSeatStatusStr3(BusID, BusStartupDate, SellTicketInfo(i).szSeatNo, SellTicketInfo(i).szTicketNo)
                                    poDb.Execute szSql, lRowAffect
                                    If lRowAffect = 0 Then ShowError ERR_SeatNoCanNotSell, "SellTicketService.SelfSellTicket"
                                    bSeatBooked = False
                                Else
                                    szSql = UpdateSeatStatusStr2(BusID, BusStartupDate, SellTicketInfo(i).szSeatNo, SellTicketInfo(i).szTicketNo, bBookValid, SellTicketInfo(i).szReserved)
                                
                                    poDb.Execute szSql, lRowAffect
                                    If lRowAffect = 0 Then ShowError ERR_SeatNoCanNotSell, "SellTicketService.SelfSellTicket"
                                    bSeatBooked = True
                                End If
                            Else
                                ShowError ERR_SeatNoCanNotSell, "SellTicketService.SelfSellTicket"
                            End If
                        Else
                            bSeatBooked = False
                        End If
                        
                    End If
                    
                    If bSeatBooked Then
                        szSql = "SELECT * FROM ticket_book_lst WHERE " _
                        & " bus_id='" & BusID & "' AND " _
                        & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                        & " seat_no='" & szSeatNo & "' AND " _
                        & " status=0"
                        Set rsTemp2 = poDb.Execute(szSql)
                        szBookStationID = Trim(rsTemp2!station_id)
                        Set rsTemp2 = Nothing
                        
                       
                        
                        '更新站点的预定张数  加上了上车站代码条件
                        'If nStationLimitCount > 0 Then
                            szSql = "UPDATE Env_bus_station_lst SET " _
                            & " book_quantity=book_quantity-1 WHERE " _
                            & " bus_id='" & BusID & "' AND " _
                            & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & " station_id='" & szBookStationID & "' AND " _
                            & " sell_station_id='" & SellStationID & "' AND " _
                            & " book_quantity>0"
                            
                            poDb.Execute szSql
                        'End If
                        
                        If bBookValid Then
                            szSql = "UPDATE ticket_book_lst SET status=2 WHERE " _
                            & " bus_id='" & BusID & "' AND " _
                            & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & " seat_no='" & szSeatNo & "' AND " _
                            & " status=0 " '0状态为预定
                        Else
                            szSql = "UPDATE ticket_book_lst SET status=3 WHERE " _
                            & " bus_id='" & BusID & "' AND " _
                            & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & " seat_no='" & szSeatNo & "' AND " _
                            & " status=0 " '0状态为预定
                        End If
                        poDb.Execute szSql
                    End If
                    
                    '得到环境车次配载信息
                    Set rsTemp2 = poDb.Execute("SELECT * from Work_env_bus_allot_lst WHERE bus_date='" _
                       & ToDBDate(BusStartupDate) & "' AND " _
                       & " bus_id='" & BusID & "' AND " _
                       & " sell_station_id='" & SellStationID & "'")
                    '如果配载信息不存在,则出错
                    If rsTemp2.RecordCount = 0 Then ShowError ERR_BusAllotInfoNotExist
                    '总共允许售票张数是否>="已售张数"+1
                    
                    If rsTemp2!can_sale_quantity >= rsTemp2!sell_quantity + 1 Or rsTemp2!can_sale_quantity = -1 Then
                        poDb.Execute ("UPDATE Env_bus_allot_lst SET sell_quantity=sell_quantity + 1 WHERE bus_date='" _
                            & ToDBDate(BusStartupDate) & "' AND " _
                            & " bus_id='" & BusID & "' AND " _
                            & " sell_station_id='" & SellStationID & "'")
                    Else
                        ShowError ERR_SellQuantityISFull '超过允许售票张数
                    End If
                    
                    
                    
                    If bSellStationCanSellEachOther Then
                        '得到环境车次售票点信息
                        Set rsTemp2 = poDb.Execute("SELECT * from work_env_bus_sell_station_lst WHERE bus_date='" _
                        & ToDBDate(BusStartupDate) & "' AND " _
                        & " bus_id='" & BusID & "' AND " _
                        & " sell_station_id='" & m_oActiveUser.SellStationID & "'")
                        If rsTemp2.RecordCount = 0 Then ShowError ERR_BusSellStationInfoNotExist
                        
                        '总共允许售票张数是否大于"已售张数"+1
                        
                        If rsTemp2!can_sale_quantity >= rsTemp2!sell_quantity + 1 Or rsTemp2!can_sale_quantity = -1 Then
                            poDb.Execute ("UPDATE env_bus_sell_station_lst SET sell_quantity=sell_quantity + 1 WHERE bus_date='" _
                                & ToDBDate(BusStartupDate) & "' AND " _
                                & " bus_id='" & BusID & "' AND " _
                                & " sell_station_id='" & m_oActiveUser.SellStationID & "'")
                        Else
                            ShowError ERR_SellQuantityISFull '超过允许售票张数
                        End If
                    End If

                    Set rsTemp2 = Nothing
                    
                    If nSeatStatus = ST_SeatProjectBooked Or nSeatStatus = ST_SeatBooked Then
                        '如果座位为计划预定,环境预定
                    Else
                        '判断到站的限售张数够用否
                        If nStationLimitCount >= 0 Then
                            If nStationLimitCount = 0 Then
                                ShowError ERR_BusReachMaxStationTicketCount, "SellTicketService.SelfSellTicket"
                            Else
                                '更新可售张数 加上了上车站代码条件
                                szSql = "UPDATE Env_bus_station_lst SET " _
                                & " sale_ticket_quantity=sale_ticket_quantity-1 WHERE " _
                                & " bus_id='" & BusID & "' AND " _
                                & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                                & " sale_ticket_quantity>book_quantity AND " _
                                & " sell_station_id='" & SellStationID & "' AND " _
                                & " station_id='" & DesStationID & "'"
                                poDb.Execute szSql, lRowAffect
                                
                                If lRowAffect = 0 Then ShowError ERR_BusReachMaxStationTicketCount, "SellTicketService.SelfSellTicket"
                            End If
                            
                        End If
                        
                        szSql = "UPDATE Env_bus_info SET sale_seat_quantity=sale_seat_quantity-1 , "
                        Select Case szSeatType
                            Case cszSeatType
                                szSql = szSql & "seat_remain=seat_remain - 1 "
                            Case cszBedType
                                szSql = szSql & "bed_remain=bed_remain - 1 "
                            Case cszAdditionalType
                                szSql = szSql & "additional_remain=additional - 1 "
                            Case cszOtherType1
                                szSql = szSql & "other_remain_1 =other_remain_1 - 1"
                            Case cszOtherType2
                                szSql = szSql & "other_remain_2=other_remain_2 - 1"
                        End Select
                        szSql = szSql & " WHERE bus_id='" & BusID & "' AND bus_date='" & ToDBDate(BusStartupDate) & "'"
                        
                        poDb.Execute szSql
                    End If
                Else
                    szSql = "UPDATE Env_bus_info SET sale_stand_seat_quantity=sale_stand_seat_quantity-1 " _
                    & " WHERE bus_id='" & BusID & "' AND " _
                    & "bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                    & " sale_stand_seat_quantity>0 "
                    
                    poDb.Execute szSql, lRowAffect
                    If lRowAffect = 0 Then ShowError ERR_NoStandTicket, "SellTicketService.SelfSellTicket"
                    szSeatNo = cszStandTicketSeatNo
                End If
            Else
                szSeatType = cszSeatType
            End If
            
            '如果票价为0且不为免票,则不允许出售
            If sgTicketPrice = 0 And SellTicketInfo(i).nTicketType <> TP_FreeTicket Then
                RaiseError ERR_CannotSellZeroPrice
            End If
            
            '判断是否是折扣票,如果是折扣票则判断是否有权限.
'            szsql = " SELECT * FROM bus_price_range WHERE project_id  ='" & SelfGetExecuteBusProject(BusStartupDate, lErrorCode).szProjectID & "' AND bus_id ='" & BusID & "'"
'            Set rsRange = poDb.Execute(szsql)
'            If rsRange.RecordCount <> 0 Then
'                If pasgSellTicketPrice(i) >= rsRange!ticket_price_up_limit And pasgSellTicketPrice(i) <= rsRange!ticket_price_lower_limit Then
'                    AssertHaveRight m_oActiveUser, RIGHT_SellDiscountTicket
'                Else
'                    ShowError ERR_TicketPriceOutOfLimit
'                End If
'            ElseIf sgTicketPrice <> pasgSellTicketPrice(i) Then
'                    ShowError ERR_TicketPriceOutOfLimit
'            End If
            '写售出的票的信息 插入了上车服务站代码
            szSql = "INSERT INTO Ticket_sell_lst(" _
            & " sell_station_id," _
            & " ticket_id," _
            & " bus_date," _
            & " bus_id," _
            & " des_station_id," _
            & " status," _
            & " ticket_type," _
            & " user_id," _
            & " seat_no," _
            & " ticket_price," _
            & " sale_ticket_type," _
            & " seat_type_id," _
            & " operation_time, " _
            & " insurance ) "
            
            
            szSql = szSql & " VALUES ('" _
            & SellStationID & "','" _
            & SellTicketInfo(i).szTicketNo & "','" _
            & ToDBDate(BusStartupDate) & "','" _
            & BusID & "','" _
            & DesStationID & "'," _
            & nSellTicketStatus & "," _
            & SellTicketInfo(i).nTicketType & ",'" _
            & m_oActiveUser.UserID & "','" _
            & szSeatNo & "'," _
            & sgTicketPrice & "," _
            & nSellTicketType & ",'" _
            & szSeatType & "','" _
            & ToDBDateTime(dtNow) & "', " _
            & TransFieldValueToString(pnInsurance) & " )"
            
            On Error GoTo MayBeTicketIDDup
            poDb.Execute szSql
            
            '写详细票价信息
            '???
'            If atiTicketPrice(SellTicketInfo(i).nTicketType).sgBase = 0 Then '此处需处理
                szSql = GetTicketPriceDetailStr(BusID, BusStartupDate, SellStationID, DesStationID, szSeatType, IIf(SellTicketInfo(i).nTicketType = TP_FreeTicket, TP_FullPrice, SellTicketInfo(i).nTicketType))
                Set rsTemp = poDb.Execute(szSql)
                LetPriceDetailFromRsToType rsTemp, atiTicketPrice(SellTicketInfo(i).nTicketType), 0
'            End If
            szSql = GetInsertPriceDetailStr(atiTicketPrice(SellTicketInfo(i).nTicketType), SellTicketInfo(i).szTicketNo)
            poDb.Execute szSql
            siSellInfo.aszSeatType(i) = szSeatType
            siSellInfo.aszTicketType(i) = SellTicketInfo(i).nTicketType
            siSellInfo.asgTicketPrice(i) = sgTicketPrice
            siSellInfo.aszSeatNO(i) = szSeatNo
        Next
        szSql = "SELECT station_name FROM station_info WHERE station_id='" & DesStationID & "'"
        Set rsTemp = poDb.Execute(szSql)
        
        siSellInfo.szDesStationName = FormatDbValue(rsTemp!station_name)
        
    Else
        ShowError ERR_NoTicketInfoSpecified
    End If
    
    SelfSellTicket = siSellInfo
    Set rsTemp = Nothing
    Set rsRange = Nothing
     Set rsTemp2 = Nothing
     Set oParam = Nothing
    Exit Function
MayBeTicketIDDup:
    If poDb.HaveThisNativeErr(ERR_DBAddDuplicate) Then

        ShowError ERR_TicketNoDup
    Else
        err.Raise err.Number
    End If
    Set rsTemp = Nothing
    Set rsRange = Nothing
    Set rsTemp2 = Nothing
    Set oParam = Nothing
End Function

Private Function GetTicketPriceDetailStr(pszBusID As String, pdtDate As Date, pszSellStationID As String, pszStaionID As String, pszSeatType As String, pnTicketType As ETicketType) As String
    Dim szSql As String
    Dim szTemp As String
    Dim oParam As New SystemParam
    oParam.Init m_oActiveUser
    
    szSql = "SELECT * FROM Work_env_bus_price_lst WHERE " _
    & " bus_id='" & pszBusID & "' AND " _
    & " bus_date='" & ToDBDate(pdtDate) & "' AND " _
    & " station_id='" & pszStaionID & "' AND " _
    & " sell_station_id='" & pszSellStationID & "' AND " _
    & "seat_type_id='" & pszSeatType & "' AND " _
    & " ticket_type=" & pnTicketType

    GetTicketPriceDetailStr = szSql
End Function


Private Function GetInsertPriceDetailStr(ptiPriceDetailInfo As TRETicketPrice, pszTicketID As String) As String
    Dim szSql As String
    Dim i As Integer
    
    szSql = " INSERT INTO ticket_price_lst(" _
    & "ticket_id," _
    & " ticket_price_total," _
    & " base_price," _
    & " price_item_1," _
    & " price_item_2," _
    & " price_item_3," _
    & " price_item_4," _
    & " price_item_5," _
    & " price_item_6," _
    & " price_item_7," _
    & " price_item_8," _
    & " price_item_9," _
    & " price_item_10," _
    & " price_item_11," _
    & " price_item_12," _
    & " price_item_13," _
    & " price_item_14," _
    & " price_item_15) VALUES('" _
    & pszTicketID & "'," _
    & ptiPriceDetailInfo.sgTotal & "," _
    & ptiPriceDetailInfo.sgBase
    For i = 1 To 15
        szSql = szSql & "," & ptiPriceDetailInfo.asgPrice(i)
    Next
    szSql = szSql & ")"
    
    GetInsertPriceDetailStr = szSql
End Function

Private Sub LetPriceDetailFromRsToType(prsIn As Recordset, ptiInOut As TRETicketPrice, sgDiscountPrice As Double)
    'sgDiscountPrice  所售的价格与原票价格的差额
    Dim i As Integer
    ptiInOut.sgBase = FormatDbValue(prsIn!base_carriage) + sgDiscountPrice
    ptiInOut.sgTotal = FormatDbValue(prsIn!ticket_price_total) + sgDiscountPrice
    For i = 1 To 15
        ptiInOut.asgPrice(i) = prsIn("price_item_" & i)
    Next
    
End Sub




'回收座位或站票的资源信息
Private Sub RecycleSeat(pszTicketID As String, poDb As RTConnection)
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim szBusID As String
    Dim dtDate As Date
    Dim szDesStationID As String
    Dim szSeatNo As String
    Dim szSeatTypeID As String
    Dim szSellStationID As String
    
    szSql = "SELECT * FROM Ticket_sell_lst WHERE ticket_id='" & pszTicketID & "'"
    Set rsTemp = poDb.Execute(szSql)
    
    szBusID = FormatDbValue(rsTemp!bus_id)
    dtDate = FormatDbValue(rsTemp!bus_date)
    szDesStationID = FormatDbValue(rsTemp!des_station_id)
    szSeatNo = FormatDbValue(rsTemp!seat_no)
    szSellStationID = FormatDbValue(rsTemp!sell_station_id)
    
    If szSeatNo = cszStandTicketSeatNo Then '是站票
        '将该环境车次的可售站票数加一
        szSql = "UPDATE Env_bus_info SET " _
        & " sale_stand_seat_quantity=sale_stand_seat_quantity+1 WHERE " _
        & " bus_id='" & szBusID & "' AND " _
        & " bus_date='" & ToDBDate(dtDate) & "'"
        poDb.Execute szSql
    Else '不是站票
        szSeatTypeID = FormatDbValue(rsTemp!seat_type_id)
        
        szSeatTypeID = FormatDbValue(rsTemp!seat_type_id)
        
        '将该环境车次的可售座位数减一和可售座位类型余留减一
        szSql = "UPDATE Env_bus_info SET sale_seat_quantity=sale_seat_quantity+1 , "
        Select Case szSeatTypeID
            Case cszSeatType
                szSql = szSql & " seat_remain=seat_remain + 1  "
            Case cszBedType
                szSql = szSql & " bed_remain=bed_remain + 1 "
            Case cszAdditionalType
                szSql = szSql & " additional_remain=additional_remain + 1 "
            Case cszOtherType1
                szSql = szSql & " other_remain_1 =other_remain_1 + 1 "
            Case cszOtherType2
                szSql = szSql & " other_remain_2 =other_remain_2 + 1 "
        End Select
        szSql = szSql & " WHERE bus_id='" & szBusID & "' AND bus_date='" & ToDBDate(dtDate) & "'"
        poDb.Execute szSql
        
        '改变相应座位号的状态
        szSql = "UPDATE Env_bus_seat_lst SET " _
        & " status=" & ST_SeatCanSell & "," _
        & " ticket_no='' WHERE " _
        & " bus_id='" & szBusID & "' AND " _
        & " bus_date='" & ToDBDate(dtDate) & "' AND " _
        & " seat_no='" & szSeatNo & "'"
        
        poDb.Execute szSql
    End If
    
    '改变该车次相应到站的可售票数,加上了上车站条件
    szSql = "UPDATE Env_bus_station_lst SET " _
    & " sale_ticket_quantity=sale_ticket_quantity+1 WHERE " _
    & " bus_id='" & szBusID & "' AND " _
    & " bus_date='" & ToDBDate(dtDate) & "' AND " _
    & " station_id='" & szDesStationID & "' AND " _
    & " sell_station_id='" & szSellStationID & "' AND " _
    & " max_sale_quantity>0 "
    
    poDb.Execute szSql
    
End Sub


'得到指定日期能到指定车站的所有车次及其相关信息
'日期比较
Public Function GetBusRs(BusDate As Date, DesStationID As String, Optional IgnoreOffTime As Boolean = False, Optional BusID As String = "", Optional IsDiscount As Boolean = False) As Recordset
    Dim oDb As New RTConnection
    Dim oParam As New SystemParam
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim nStopSellTime As Integer
    Dim dtNowDateTime As Date
    Dim dbExtraBeforeTime As Double
    Dim dbExtraAfterTime As Double
    Dim dbInternetBusShow As Double
    
    Dim szSelect As String
    Dim szTable As String
    Dim szWhere As String
    
    oParam.Init m_oActiveUser
    nStopSellTime = oParam.StopSellTime
    dtNowDateTime = oParam.NowDateTime
    
    dbInternetBusShow = oParam.InternetBusShow
    
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    szSelect = "SELECT Work_env_bus_info.* ,e.check_gate_id sell_check_gate_id ,seat_type_id,route_name ,Work_env_bus_station_lst.mileage end_station_mileage" _
        & " ,Work_env_bus_station_lst.sell_station_id,sell_station_name,e.bus_start_time as busstarttime, " _
        & "         full_price, half_price,preferential_ticket1,preferential_ticket2,preferential_ticket3," _
        & "         sale_ticket_quantity,stop_sale_time, Work_env_bus_station_lst.book_quantity "
        
    szTable = "  FROM Work_env_bus_info,Work_env_bus_station_lst ,route_info,sell_station_info s,Work_env_bus_allot_lst e  "
    
    
    szWhere = "  WHERE s.sell_station_id=Work_env_bus_station_lst.sell_station_id AND e.sell_station_id=s.sell_station_id AND " _
        & "  e.bus_date=Work_env_bus_info.bus_date AND e.bus_id=Work_env_bus_info.bus_id AND "
    
    If Not IgnoreOffTime Then
        szWhere = szWhere & "(('" & ToDBDateTime(DateAdd("s", nStopSellTime * 60, dtNowDateTime)) & "'<=e.bus_start_time " _
        & IIf(Not m_bAllowScrollBusSaleForever, "", " AND bus_kind<>" & TP_ScrollBus) _
        & ")  " _
        & " OR ( " _
        & IIf(Not m_bAllowScrollBusSaleForever, "", "bus_kind=" & TP_ScrollBus & " OR  ") _
        & " e.bus_start_time>='" & ToDBDateTime(DateAdd("s", nStopSellTime * 60, dtNowDateTime)) & "'))   "
        
        szWhere = szWhere _
        & " AND Work_env_bus_info.bus_date='" & ToDBDate(BusDate) & "' AND " _
        & " Work_env_bus_info.bus_id=Work_env_bus_station_lst.bus_id AND " _
        & " Work_env_bus_info.bus_date=Work_env_bus_station_lst.bus_date AND " _
        & " Work_env_bus_station_lst.station_id='" & DesStationID & "' AND " _
        & " ( sale_seat_quantity>=0 OR sale_stand_seat_quantity>0 " _
        & IIf(Not m_bAllowScrollBusSaleForever, "", " OR bus_kind=" & TP_ScrollBus) _
        & ") AND " _
        & " ( ( max_sale_quantity>0 AND  sale_ticket_quantity>=0) OR  " _
        & " ( max_sale_quantity<0) ) AND Work_env_bus_info.interlink_quantity<>0 " _
        & " AND e.bus_start_time >= '" & DateAdd("n", dbInternetBusShow, dtNowDateTime) & "' " _
        & " AND Work_env_bus_info.route_id=route_info.route_id  AND (e.can_sale_quantity <>0 )" _
        'ORDER BY Work_env_bus_info.bus_start_time,Work_env_bus_info.bus_id  "
    
    Else
        dbExtraBeforeTime = oParam.ExtraBusShowBefore
        dbExtraAfterTime = oParam.ExtraBusShowAfter
        
        szWhere = szWhere & "'" & ToDBDateTime(DateAdd("n", dbExtraAfterTime, dtNowDateTime)) & "'>=e.bus_start_time AND " _
        & "'" & ToDBDateTime(DateAdd("n", -dbExtraBeforeTime, dtNowDateTime)) & "'<=e.bus_start_time " _
        & IIf(Not m_bAllowScrollBusSaleForever, "", " AND bus_kind<>" & TP_ScrollBus) _
        & " AND  " _
        & " Work_env_bus_info.bus_date='" & ToDBDate(BusDate) & "' AND " _
        & " Work_env_bus_info.bus_id=Work_env_bus_station_lst.bus_id AND " _
        & " Work_env_bus_info.bus_date=Work_env_bus_station_lst.bus_date AND " _
        & " Work_env_bus_station_lst.station_id='" & DesStationID & "' AND " _
        & " ( sale_seat_quantity>=0 OR sale_stand_seat_quantity>0 " _
        & IIf(Not m_bAllowScrollBusSaleForever, "", " OR bus_kind=" & TP_ScrollBus) _
        & ") AND " _
        & " ( ( max_sale_quantity>0 AND  sale_ticket_quantity>0) OR  " _
        & " ( max_sale_quantity<0) )  AND Work_env_bus_info.interlink_quantity<>0 " _
        & "  AND Work_env_bus_info.route_id=route_info.route_id " _
        'ORDER BY Work_env_bus_info.bus_start_time,Work_env_bus_info.bus_id  "
        
    
    End If
    
    If BusID <> "" Then
        szWhere = szWhere & " AND Work_env_bus_info.bus_id='" & Trim(BusID) & "'"
    End If
    If Not m_bSellStationCanSellEachOther Then
        If m_oActiveUser.SellStationID <> "" Then
             szWhere = szWhere & " AND Work_env_bus_station_lst.sell_station_id='" & Trim(m_oActiveUser.SellStationID) & "'"
        End If
    Else
        
        szTable = szTable & " , work_env_bus_sell_station_lst t "
        szWhere = szWhere & " AND t.bus_id = Work_env_bus_info.bus_id " _
            & " AND t.bus_date = Work_env_bus_info.bus_date " _
            & " AND t.sell_station_id = '" & Trim(m_oActiveUser.SellStationID) & "'  AND (t.can_sale_quantity >=t.sell_quantity or t.can_sale_quantity <0)"
        
    End If
    
    szWhere = szWhere & " ORDER BY e.bus_start_time,Work_env_bus_info.bus_id "
    Set rsTemp = oDb.Execute(szSelect & szTable & szWhere)
    Set GetBusRs = rsTemp
    Set rsTemp = Nothing
    Set oParam = Nothing
    Set oDb = Nothing
End Function



'得到指定日期的所有可能到站及其相关信息。
'日期比较
Public Function GetStationRs(BusDate As Date, Optional IgnoreOffTime As Boolean = False) As Recordset
    Dim oDb As New RTConnection
    Dim oParam As New SystemParam
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim nStopSellTime As Integer
    Dim dtNowDateTime As Date

    oParam.Init m_oActiveUser
    nStopSellTime = oParam.StopSellTime
    dtNowDateTime = oParam.NowDateTime

    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    szSql = "SELECT DISTINCT station_info.station_id,station_name ,station_input_code FROM " _
    & " Work_env_bus_info,Work_env_bus_station_lst,station_info ,bus_kind_code WHERE "

    If Not IgnoreOffTime Then
        szSql = szSql & "( '" & ToDBDateTime(DateAdd("n", nStopSellTime, dtNowDateTime)) & "'<=Work_env_bus_info.bus_start_time OR " _
        & " Work_env_bus_info.bus_kind=" & TP_ScrollBus & ") AND "

        szSql = szSql & " Work_env_bus_info.bus_date=Work_env_bus_station_lst.bus_date AND " _
        & " Work_env_bus_info.bus_id=Work_env_bus_station_lst.bus_id AND " _
        & " Work_env_bus_info.bus_date='" & ToDBDate(BusDate) & "' AND " _
        & " (sale_seat_quantity>0 OR sale_stand_seat_quantity>0 OR Work_env_bus_info.bus_kind =" & TP_ScrollBus & ") AND " _
        & " station_info.station_id=Work_env_bus_station_lst.station_id AND " _
        & " ( ( max_sale_quantity>0 AND  sale_ticket_quantity>0) OR  " _
        & " ( max_sale_quantity<0) ) AND " _
        & " Work_env_bus_info.bus_kind=bus_kind_code.bus_kind " _
        & " ORDER BY station_info.station_id"
    Else
        szSql = szSql _
        & " (Work_env_bus_info.bus_kind <>" & TP_ScrollBus & ") AND "

        szSql = szSql & " Work_env_bus_info.bus_date=Work_env_bus_station_lst.bus_date AND " _
        & " Work_env_bus_info.bus_id=Work_env_bus_station_lst.bus_id AND " _
        & " Work_env_bus_info.bus_date='" & ToDBDate(BusDate) & "' AND " _
        & " (sale_seat_quantity>0 OR sale_stand_seat_quantity>0 OR Work_env_bus_info.bus_kind =" & TP_ScrollBus & ") AND " _
        & " station_info.station_id=Work_env_bus_station_lst.station_id AND " _
        & " ( ( max_sale_quantity>0 AND  sale_ticket_quantity>0) OR  " _
        & " ( max_sale_quantity<0) ) " _
        & " Work_env_bus_info.bus_kind=bus_kind_code.bus_kind " _
        & " ORDER BY station_info.station_id"
    End If


    Set rsTemp = oDb.Execute(szSql)
    Set GetStationRs = rsTemp

End Function


Private Function UpdateSeatStatusStr2(pszBusID As String, pszDate As Date, pszSeatNo As String, pszTicketID As String, pbBookValid As Boolean, Optional pszBookNumber As String = "") As String
    Dim szSql As String
    szSql = "UPDATE Env_bus_seat_lst SET status=" & ST_SeatSold & "," _
    & " ticket_no='" & pszTicketID & "' FROM ticket_book_lst " _
    & " WHERE Env_bus_seat_lst.bus_id='" & pszBusID & "' AND " _
    & " Env_bus_seat_lst.bus_date='" & ToDBDate(pszDate) & "' AND " _
    & " Env_bus_seat_lst.seat_no='" & pszSeatNo & "' AND " _
    & " ticket_book_lst.status=" & ST_SeatCanSell & " AND " _
    & " Env_bus_seat_lst.bus_id=ticket_book_lst.bus_id AND " _
    & " Env_bus_seat_lst.bus_date=ticket_book_lst.bus_date AND " _
    & " Env_bus_seat_lst.seat_no=ticket_book_lst.seat_no AND " _
    & " Env_bus_seat_lst.status=" & ST_SeatBooked
    If pbBookValid Then
'        szSql = szSql & " AND book_number='" & pszBookNumber & "' "
    End If
    UpdateSeatStatusStr2 = szSql
End Function
Private Function UpdateSeatStatusStr3(pszBusID As String, pszBusDate As Date, pszSeatNo As String, pszTicketID As String) As String
    Dim szSql As String
    szSql = "UPDATE Env_bus_seat_lst SET status=" & ST_SeatSold & "," _
            & " ticket_no='" & pszTicketID & "' " _
            & " WHERE bus_id='" & pszBusID & "' AND " _
            & " bus_date='" & ToDBDate(pszBusDate) & "' AND " _
            & " seat_no='" & pszSeatNo & "' AND " _
            & " status=" & ST_SeatProjectBooked & " "
    UpdateSeatStatusStr3 = szSql
End Function

'//////////////////////////
'释放座位
Private Sub ResumeSeat(pszTicket As String, poDb As RTConnection)
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim rsTemp1 As Recordset
    Dim szBusID As String
    Dim dtDate As Date
    Dim szDesStationID As String
    Dim szSeatNo As String
    Dim szSeatTypeID As String
    
    On Error GoTo here
    szSql = "SELECT * FROM Ticket_sell_lst WHERE ticket_id='" & pszTicket & "'"
    Set rsTemp = poDb.Execute(szSql)
    
    szBusID = FormatDbValue(rsTemp!bus_id)
    dtDate = FormatDbValue(rsTemp!bus_date)
    szDesStationID = FormatDbValue(rsTemp!des_station_id)
    szSeatNo = FormatDbValue(rsTemp!seat_no)
     
     '////////////////////////////
    '判断是否还有可售票数
    szSql = "SELECT * FROM Env_bus_info WHERE " _
    & " bus_id='" & szBusID & "'AND " _
    & " bus_date='" & ToDBDate(dtDate) & "'"
    Set rsTemp1 = poDb.Execute(szSql)
   
    
    If szSeatNo = cszStandTicketSeatNo Then '是站票
       If rsTemp1!sale_stand_seat_quantity >= 1 Then
            '将该环境车次的可售站票数减一
            szSql = "UPDATE Env_bus_info SET " _
            & " sale_stand_seat_quantity=sale_stand_seat_quantity-1 WHERE " _
            & " bus_id='" & szBusID & "' AND " _
            & " bus_date='" & ToDBDate(dtDate) & "'"
            poDb.Execute szSql
       Else
            GoTo here
       End If
    Else '不是站票
        
        
        If rsTemp1!sale_seat_quantity >= 1 Then
        
            szSeatTypeID = FormatDbValue(rsTemp!seat_type_id)
        
            '将该环境车次的可售座位数减一和可售座位类型余留减一
            szSql = "UPDATE Env_bus_info SET sale_seat_quantity=sale_seat_quantity-1 , "
            Select Case szSeatTypeID
                Case cszSeatType
                    szSql = szSql & " seat_remain=seat_remain - 1  "
                Case cszBedType
                    szSql = szSql & " bed_remain=bed_remain - 1 "
                Case cszAdditionalType
                    szSql = szSql & " additional_remain=additional_remain - 1 "
                Case cszOtherType1
                    szSql = szSql & " other_remain_1 =other_remain_1 - 1 "
                Case cszOtherType2
                    szSql = szSql & " other_remain_2 =other_remain_2 - 1 "
            End Select
            szSql = szSql & " WHERE bus_id='" & szBusID & "' AND bus_date='" & ToDBDate(dtDate) & "'"

            poDb.Execute szSql
            
            '改变相应座位号的状态
            szSql = "UPDATE Env_bus_seat_lst SET " _
            & " status=" & ST_SeatSold & "," _
            & " ticket_no='" & pszTicket & "' WHERE " _
            & " bus_id='" & szBusID & "' AND " _
            & " bus_date='" & ToDBDate(dtDate) & "' AND " _
            & " seat_no='" & szSeatNo & "'"
            poDb.Execute szSql
        Else
            GoTo here
        End If
    End If
    
    '改变该车次相应到站的可售票数
    szSql = "UPDATE Env_bus_station_lst SET " _
    & " sale_ticket_quantity=sale_ticket_quantity-1 WHERE " _
    & " bus_id='" & szBusID & "' AND " _
    & " bus_date='" & ToDBDate(dtDate) & "' AND " _
    & " station_id='" & szDesStationID & "' AND " _
    & " max_sale_quantity>0 "
    
    poDb.Execute szSql
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    On Error GoTo 0
    Exit Sub
here:
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
'    MsgBox "已无可售票数" & vbCrLf & "取消退票或取消废票可能产生超载", vbInformation, "售票台:"
End Sub

'////////////////////////////////////////////////////////
'得到所有票种
Public Function GetAllTicketType(Optional nValid As Integer = TP_TicketTypeAll) As TTicketType()
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim szSql  As String
    Dim atTemp(1 To 6) As TTicketType
    On Error GoTo here
    
    szSql = " SELECT * FROM Ticket_type_code "
    '默认为所有的票种
    If nValid <> TP_TicketTypeAll Then
        szSql = szSql & " WHERE ticket_type_valid = " & nValid
    End If
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    Set rsTemp = oDb.Execute(szSql)
     
    For i = 1 To rsTemp.RecordCount
        atTemp(i).nTicketTypeID = rsTemp!ticket_type_id
        atTemp(i).szTicketTypeName = rsTemp!Ticket_Type_name
        atTemp(i).nTicketTypeValid = rsTemp!ticket_type_valid
        atTemp(i).szAnnotation = rsTemp!Annotation
        rsTemp.MoveNext
    Next i
    
    GetAllTicketType = atTemp
    Set rsTemp = Nothing
    Set oDb = Nothing
    On Error GoTo 0
    Exit Function
here:
    Set rsTemp = Nothing
    Set oDb = Nothing
    err.Raise err.Number
End Function
'////////////////////////////////////////////////////////
'得到线路名称和发车时间
Public Function GetRouteAndTime(pdtBusDate As Date, pszBusID As String) As String()
Dim aszTemp() As String
Dim szSql As String
Dim oDb As New RTConnection
Dim rsTemp As Recordset
ReDim aszTemp(1 To 4)
 On Error GoTo here
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    szSql = "SELECT r.route_name,e.bus_start_time,e.split_company_id,e.transport_company_id , c.transport_company_short_name " _
            & " FROM Work_env_bus_info e,route_info r,company_info c " _
            & " WHERE e.split_company_id = c.transport_company_id AND e.bus_date='" & ToDBDate(pdtBusDate) & "' AND " _
            & "e.bus_id='" & pszBusID & "' AND " _
            & "r.route_id=e.route_id"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount >= 1 Then
        aszTemp(1) = FormatDbValue(rsTemp!route_name)
        aszTemp(2) = Trim(CStr(rsTemp!bus_start_time))
        aszTemp(3) = FormatDbValue(rsTemp!split_company_id)
        aszTemp(4) = FormatDbValue(rsTemp!transport_company_short_name)
        
    End If
    GetRouteAndTime = aszTemp
On Error GoTo 0
    Set rsTemp = Nothing
    Set oDb = Nothing
Exit Function
here:
    Set rsTemp = Nothing
    Set oDb = Nothing
    err.Raise err.Number
End Function


'///////////////////////////
Public Function GetAllCheckGate(Optional CheckGate As String = "") As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    If CheckGate = "" Then
        Set rsTemp = oDb.Execute("SELECT * FROM Checkgate_info")
    Else
        Set rsTemp = oDb.Execute("SELECT * FROM Checkgate_info WHERE check_gate_id='" & CheckGate & "'")
    End If
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    ReDim szaTemp(1 To nCount, 1 To 3) As String
    For i = 1 To nCount
        szaTemp(i, 1) = rsTemp!check_gate_id
        szaTemp(i, 2) = FormatDbValue(rsTemp!check_gate_name)
        szaTemp(i, 3) = FormatDbValue(rsTemp!Annotation)
        rsTemp.MoveNext
    Next
    GetAllCheckGate = szaTemp
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

'////////////////////////
'得到座位类型代码和名称
Public Function GetAllSeatType() As String()
Dim szSql As String
Dim i As Integer
Dim rsTemp As Recordset
Dim oDb As New RTConnection
Dim aszSeatType() As String
Dim nCount As Integer
oDb.ConnectionString = GetConnectionStr(cszSellTicket)
szSql = "SELECT * FROM Seat_type_code "
Set rsTemp = oDb.Execute(szSql)
nCount = rsTemp.RecordCount
If nCount <> 0 Then
    ReDim aszSeatType(1 To nCount, 1 To 3)
    For i = 1 To nCount
        aszSeatType(i, 1) = rsTemp!seat_type_id
        aszSeatType(i, 2) = rsTemp!seat_type_name
        aszSeatType(i, 3) = rsTemp!Annotation
        rsTemp.MoveNext
    Next
    GetAllSeatType = aszSeatType
End If
Set rsTemp = Nothing
Set oDb = Nothing
End Function

'**********************************************自助售票用
'XXD 06-10-19
'取热门站点
Public Function HotStation() As Recordset
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim i As Integer
    Dim nLen As Integer
    Dim szWhere As String
    Dim rs As Recordset
    

    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    szSql = "SELECT * FROM hot_station_info,station_info where hot_station_info.station_id=station_info.station_id order by id "
    Set rs = oDb.Execute(szSql)
    Set HotStation = rs
End Function





' *******************************************************************
' *   Member Name: InternetSellTicket                               *
' *   Brief Description: 网上售票                                   *
' *   Engineer: xxd                                                 *
' *   Date Generated: 2006-11-16          *
' *   Last Revision Date: 最后一次修改日期(日期格式：YYYY/MM/DD)    *
' *******************************************************************


'网上售票
Public Function InterNetSellLocalTicket(BusStartupDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String _
             , aspSellTicketParam() As TSellTicketParam, panInsurance() As Integer, aszBankParam() As String) As TSellTicketResult()
    'panInsurance 保险费

    Dim oDb As New RTConnection
    Dim rsTemp() As TSellTicketResult
    Dim nCount As Integer
    Dim iLen As Integer
    iLen = ArrayLength(BusID)

    
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_SellTicket
    ReDim rsTemp(1 To iLen)

    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    oDb.BeginTrans
    On Error GoTo here
    For nCount = 1 To iLen
        rsTemp(nCount) = InterNetSellTicketLocalTrans(BusStartupDate(nCount), BusID(nCount), SellStationID(nCount), DesStationID(nCount), aspSellTicketParam(nCount).BuyTicketInfo, aspSellTicketParam(nCount).pasgSellTicketPrice, aszBankParam, panInsurance(nCount), False, ST_TicketNormal, oDb)
    Next nCount
    InterNetSellLocalTicket = rsTemp
    oDb.CommitTrans
    Set oDb = Nothing
    Exit Function
here:
    oDb.RollbackTrans
    Set oDb = Nothing
    
        err.Raise err.Number
End Function


'内部网上售票事务函数(其实为先做预留,等人到后取票时再售)
Private Function InterNetSellTicketLocalTrans(BusStartupDate As Date, BusID As String, SellStationID As String, DesStationID As String, SellTicketInfo() As TBuyTicketInfo _
    , pasgSellTicketPrice() As Single, paszBankParam() As String, pnInsurance As Integer, Optional pbExtraSell As Boolean = False _
    , Optional pnTicketStatus As ETicketStatus = ST_TicketNormal, Optional poDb As RTConnection) As TSellTicketResult
    
 
    Dim nBookSeatCount As Integer, i As Integer '预留座位数
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szBookID As String
    Dim oParam As New SystemParam
    Dim rsTemp As Recordset, rsTemp2 As Recordset
    Dim lRowAffect As Long
    Dim szTemp As String
    Dim nStationLimitedCount As Integer
    
    Dim BookID() As TBuyTicketInfo '预留票号
    Dim BookSeatID() As String '预留座位号
    
    Dim szSellStationID As String
    Dim szSeatID As String '座位号
    
    
    
    
    Dim dtBusFullDateTime As Date
    Dim nTemp As Double
    Dim dtNow As Date
    Dim nBusStatus As EREBusStatus
    Dim nBuyTicketCount As Integer
    Dim sgFullPrice As Double, sgHalfPrice As Double
    Dim nStationLimitCount As Integer
    Dim szSeatNo As String, sgTicketPrice As Double
    Dim nBusType As EBusType
    
    Dim atiTicketPrice() As TRETicketPrice
    Dim siSellInfo As TSellTicketResult
    
    Dim lErrorCode As Long
    Dim rsRange As Recordset
    
    Dim nStationRemainCount As Integer
    Dim nCanSellTicket As Integer '可售座位数
    Dim bNeedBook As Boolean '是否上了预定模块
    Dim bBookValid As Boolean '是否还需判断预定
    Dim bSeatBooked As Boolean
    Dim szBookStationID As String '预定的到站
    
    Dim szSeatType As String '座位类型
    Dim szSqlSeat As String  '座位票价
    Dim nSeatStatus As Integer '座位状态
    
    Dim szStopTime As String
    
    Dim nSellTicketType As Integer
    Dim nSellTicketStatus As ETicketStatus
    Dim tTSeatInfo(1 To 5) As TSeatInfoCount
    Dim j As Integer
    Dim szSqlSet As String
    Dim nSaleSeatQuantity As Integer
    Dim rs As Recordset
    oParam.Init m_oActiveUser

    dtNow = oParam.NowDateTime
    
    nTemp = DateDiff("d", dtNow, BusStartupDate)
    
    '如果要购票的车次不在当天和预售天之内则出错
    If nTemp > oParam.PreSellDate Or nTemp < 0 Then ShowError ERR_BusNotInIlegalDate
    If nTemp > oParam.UserCanSellDate Then ShowError ERR_UserCannotSell


   
'    '如果票价为0且不为免票,则不允许出售
'    For i = 1 To ArrayLength(pszTicketID)
'    If pszTPrice(1).pasgSellTicketPrice(i) = 0 And pszTicketID(i).nTicketType <> TP_FreeTicket Then
'        RaiseError ERR_CannotSellZeroPrice
'    End If
'    Next

    
    

'    poDb.ConnectionString = GetConnectionStr("")
    szSql = "SELECT * FROM internet_sell_ticket_info WHERE getticket_id='" & paszBankParam(1, 4) & "' AND validate_id='" & paszBankParam(1, 5) & "'"
    Set rs = poDb.Execute(szSql)
    If rs.RecordCount > 0 Then err.Raise 2008, err.Source, "此取票号与订单号已存在!"
    
    szSql = "SELECT * FROM work_env_bus_info e,work_env_bus_station_lst s " _
    & " WHERE e.bus_id='" & BusID & "' AND " _
    & " e.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
    & " e.bus_id=s.bus_id AND " _
    & " e.bus_date=s.bus_date AND " _
    & " station_id='" & DesStationID & "'"
    
    Set rsTemp = poDb.Execute(szSql)
    If rsTemp.RecordCount < 1 Then
        ShowError ERR_BusNotExist, "CBook.Book", "BusID=" & BusID & ",BusDate=" & ToDBDate(BusStartupDate)
    Else
        szSellStationID = FormatDbValue(rsTemp!sell_station_id)
        nStationLimitedCount = rsTemp!max_sale_quantity
        If nStationLimitedCount = 0 Then ShowError ERR_BusStationMatchLimitedCount, "CBook.Book"
    End If
    
'*********************************************************
    nBusType = FormatDbValue(rsTemp!bus_kind)
    '流水车次不能补票
'    If nBusType = TP_ScrollBus And pbExtraSell Then ShowError ERR_ScrollBusCanNotExtraSell
    nBusStatus = FormatDbValue(rsTemp!Status)
    '并班车次不能售票
    If nBusStatus = ST_BusMergeStopped Then ShowError ERR_BusMerged
    '拆分停班车次不能售票
    If nBusStatus = ST_BusSlitpStop Then ShowError ERR_BusSlitpStop
    '车次正在拆分不能售票
    If nBusStatus >= 32 Then ShowError ERR_BusIsSplit
    '停检之后不能售票（实际不太可能出现，除非停售时间设得太小了）如果是补票则不限制
'    If pbExtraSell And nBusStatus = ST_BusStopCheck Then ShowError ERR_BusStopCheck
    If nBusStatus = ST_BusStopCheck Then ShowError ERR_BusStopCheck
    If nBusStatus = ST_BusStopped Then ShowError ERR_BusStopped
    dtBusFullDateTime = CDate(Format(FormatDbValue(rsTemp!bus_date), cszDateStr) & " " & Format(FormatDbValue(rsTemp!bus_start_time), cszTimeStr))
    '判断是否到了停售时间(如果是补票则要过了正常售票时间才能补)
'    If nBusType <> TP_ScrollBus Then '如果是固定车次才限制
'        If pbExtraSell Then
'            If DateDiff("s", dtNow, dtBusFullDateTime) >= oParam.StopSellTime * 60 Then ShowError ERR_BusNotStopSellTicket
'        Else
''            If DateDiff("s", dtNow, dtBusFullDateTime) < oParam.StopSellTime * 60 Then ShowError ERR_BusStopSellTicket
'        End If
'    End If



    szStopTime = FormatDbValue(rsTemp!stop_sale_time)
    If nBusType <> TP_ScrollBus Then '如果是固定车次才限制
        If szStopTime <> "-1" Then '如果该车次到指定的站限制售票时间则要进行判断
            If DateDiff("n", dtNow, dtBusFullDateTime) > GetStopTime(szStopTime) Then ShowError ERR_BusStaionStopSellTicket, "SellTicketClient.SelfSellTicketLocalTrans"
        End If
        bNeedBook = oParam.IsBookValid
        If bNeedBook Then
            bBookValid = IIf(DateAdd("n", oParam.BookTime, dtNow) >= rsTemp!bus_start_time, False, True)
        End If
    Else
        If szStopTime <> "-1" Then
           dtBusFullDateTime = CDate(Format(FormatDbValue(rsTemp!bus_date), cszDateStr) & " " & GetStopTime(szStopTime, False))
           If DateDiff("n", dtNow, dtBusFullDateTime) < 0 Then ShowError ERR_BusStaionStopSellTicket, "SellTicketService.SelfSellTicket"
        End If
    End If
    nBuyTicketCount = ArrayLength(SellTicketInfo)
    nStationLimitCount = FormatDbValue(rsTemp!max_sale_quantity)
    nStationRemainCount = FormatDbValue(rsTemp!sale_ticket_quantity)
    nCanSellTicket = FormatDbValue(rsTemp!sale_seat_quantity)
    If nStationRemainCount > 0 Then  '此处为了折扣票由>= 改为了>
            If nBuyTicketCount > nStationRemainCount Then ShowError ERR_BusReachMaxStationTicketCount
    End If
    
    '判断可售座位数
    nBookSeatCount = ArrayLength(SellTicketInfo)
    
    If nCanSellTicket < nBookSeatCount Then
        ShowError ERR_BusReachMaxTicketCount
    End If
    
    
    '得到环境车次配载信息
    Set rsTemp2 = poDb.Execute("SELECT * from work_env_bus_allot_lst WHERE bus_date='" _
                    & ToDBDate(BusStartupDate) & "' AND " _
                    & " bus_id='" & BusID & "' AND " _
                    & " sell_station_id='" & SellStationID & "'")
    If rsTemp2.RecordCount = 0 Then ShowError ERR_BusAllotInfoNotExist
                    
    '总共允许售票张数是否大于"已售张数"+1
                    
    If rsTemp2!can_sale_quantity >= rsTemp2!sell_quantity + 1 Or rsTemp2!can_sale_quantity = -1 Then

    Else
        ShowError ERR_SellQuantityISFull '超过允许售票张数
    End If
    
            '判断到达站限售
            Set rsTemp2 = poDb.Execute("SELECT * from work_env_bus_station_lst WHERE bus_date='" _
                & ToDBDate(BusStartupDate) & "' AND " _
                & " bus_id='" & BusID & "' AND " _
                & " sell_station_id='" & SellStationID & "' AND " _
                & " station_id = '" & DesStationID & "'")
                
'            If rsTemp2.RecordCount = 0 Then ShowError ERR_BusStationInfoNotExist

            If rsTemp2!sale_ticket_quantity - nBuyTicketCount >= 0 Or rsTemp2!sale_ticket_quantity = -1 Then
            

            Else
                err.Raise 8899, , "售票站超过分配的售票张数，" & "还剩余" & rsTemp2!sale_ticket_quantity & "张"
            End If
    ReDim siSellInfo.aszSeatNO(1 To nBookSeatCount)
    ReDim siSellInfo.asgTicketPrice(1 To nBookSeatCount)
    ReDim siSellInfo.aszSeatType(1 To nBookSeatCount)
    ReDim siSellInfo.aszTicketType(1 To nBookSeatCount)
'*********************************************************
    
    poDb.BeginTrans
    On Error GoTo Error_Handle
    
    For i = 1 To nBookSeatCount
        szSqlSet = ""
        If Trim(SellTicketInfo(i).szSeatNo) = "" Then
            szSql = "SELECT MIN(seat_no) seat_no FROM Env_bus_seat_lst WHERE " _
            & " bus_id='" & BusID & "' AND " _
            & " is_special_price = 0 AND " _
            & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
            & " status=" & ST_SeatCanSell & " AND seat_type_id ='" & SellTicketInfo(i).szSeatTypeID & "'"
            Set rsTemp = poDb.Execute(szSql)
            If rsTemp.RecordCount > 0 Then
                szSeatID = rsTemp!seat_no
            Else
                ShowError ERR_BOOKSeatBooked, "CBook.Book", "BusID=" & BusID & ",BusDate=" & ToDBDate(BusStartupDate) & ",SeatNo=" & szSeatID
            End If
        Else
            szSeatID = SellTicketInfo(i).szSeatNo
        End If
        
        szSql = "SELECT is_special_price FROM Env_bus_seat_lst WHERE " _
            & " bus_id='" & BusID & "' AND " _
            & " seat_no='" & szSeatID & "' AND " _
            & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
            & " status=" & ST_SeatCanSell & " AND seat_type_id ='" & SellTicketInfo(i).szSeatTypeID & "'"
        Set rsTemp = poDb.Execute(szSql)
        '特价座位不能预订
        If rsTemp.RecordCount > 0 Then
'            If FormatDbValue(rsTemp!is_special_price) = 1 Then ShowError ERR_SpecialPriceSeatCanNotSell, "CBook.Book"
        End If
        
            szSql = "UPDATE Env_bus_seat_lst SET status=" & Str(ST_SeatReserved) & "   ," _
                  & " remark='网购' " _
                  & "  WHERE bus_id='" & BusID & "' " _
                  & " AND  bus_date='" & ToDBDate(BusStartupDate) & "' " _
                  & " AND seat_no ='" & szSeatID & "' " _
                  & " AND  Status=" & ST_SeatCanSell & ""
            poDb.Execute szSql, lRowAffect
            If lRowAffect <> 1 Then ShowError ERR_BusStationMatchLimitedCount, "CBook.Book"
        
        
            szSql = " select seat_type_id ,count(*)as CountSeat from Env_bus_seat_lst " _
                    & " where  bus_id='" & BusID & "'and  status='" & ST_SeatCanSell & "' and bus_date='" & ToDBDate(BusStartupDate) & "' group by seat_type_id "
            Set rsTemp = poDb.Execute(szSql)
            
            For j = 1 To rsTemp.RecordCount
                tTSeatInfo(CInt(rsTemp!seat_type_id)).SeatCount = FormatDbValue(rsTemp!CountSeat)
                rsTemp.MoveNext
            Next
            
            szSqlSet = szSqlSet & "seat_remain ='" & tTSeatInfo(1).SeatCount & "',"
            szSqlSet = szSqlSet & "bed_remain ='" & tTSeatInfo(2).SeatCount & "',"
            szSqlSet = szSqlSet & "additional_remain ='" & tTSeatInfo(3).SeatCount & "',"
            szSqlSet = szSqlSet & "other_remain_1 ='" & tTSeatInfo(4).SeatCount & "',"
            szSqlSet = szSqlSet & "other_remain_2 ='" & tTSeatInfo(5).SeatCount & "',"
            nSaleSeatQuantity = 0
            For j = 1 To 5
              nSaleSeatQuantity = nSaleSeatQuantity + tTSeatInfo(j).SeatCount
            Next
            
            szSqlSet = szSqlSet & "sale_seat_quantity='" & nSaleSeatQuantity & "'"
            szSql = "UPDATE Env_bus_info SET  " & szSqlSet _
                    & "  WHERE bus_id='" & BusID & "' AND  bus_date='" & ToDBDate(BusStartupDate) & "'"
            poDb.Execute (szSql)
            
            
            
            szSqlSeat = "SELECT p.ticket_price_total AS TicketPrice, p.seat_type_id AS SeatType,s.status AS SeatStatus " _
                            & " FROM work_env_bus_price_lst p,Env_bus_seat_lst s WHERE " _
                            & " p.bus_id='" & BusID & "' AND p.station_id='" & DesStationID & "' AND " _
                            & " p.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & " p.ticket_type=" & SellTicketInfo(i).nTicketType & " AND " _
                            & " p.seat_type_id =(SELECT seat_type_id FROM Env_bus_seat_lst WHERE " _
                            & " bus_id='" & BusID & "' AND bus_date='" & ToDBDate(BusStartupDate) & "' AND seat_no='" & szSeatID & "') AND " _
                            & " s.status=(SELECT status FROM Env_bus_seat_lst WHERE " _
                            & " bus_id='" & BusID & "' AND bus_date='" & ToDBDate(BusStartupDate) & "' AND seat_no='" & szSeatID & "')" _
                            & " AND p.sell_station_id ='" & SellStationID & "'" _
                            & " AND p.bus_id = s.bus_id AND p.bus_date = s.bus_date AND s.seat_no = '" & szSeatID & "'" '
                                        
            Set rsTemp = poDb.Execute(szSqlSeat)
            sgTicketPrice = rsTemp!TicketPrice
            szSeatType = FormatDbValue(rsTemp!SeatType)
            
            siSellInfo.aszSeatNO(i) = szSeatID
            siSellInfo.asgTicketPrice(i) = sgTicketPrice
            siSellInfo.aszSeatType(i) = szSeatType
            siSellInfo.aszTicketType(i) = SellTicketInfo(i).nTicketType
            '是否远程数据库需要预留
            '判断一下，系统参数
            '-----------------------------------------------------------------------------------
'            If bIsUsingRelationBus Then
'
'                oReBusService.ReserveSeat BusID, ToDBDate(BusStartupDate), szSeatID, "", ST_LockReserve
'                '-----------------------------------------------------------------------------------
'            End If
       szSql = "INSERT INTO internet_sell_ticket_info(" _
        & " bank_card_id," _
        & " busy_date," _
        & " bus_date," _
        & " bus_id," _
        & " seat_no," _
        & " sell_station_id," _
        & " des_station_id," _
        & " ticket_type," _
        & " bank_id," _
        & " center_scroll_id," _
        & " getticket_id," _
        & " validate_id," _
        & " card_id," _
        & " pay_count, " _
        & " if_print )  " _

        
        szSql = szSql & " VALUES('" _
        & paszBankParam(i, 2) & "','" _
        & paszBankParam(i, 1) & "','" _
        & ToDBDate(BusStartupDate) & "','" _
        & BusID & "','" _
        & szSeatID & "','" _
        & SellStationID & "','" _
        & DesStationID & "','" _
        & SellTicketInfo(i).nTicketType & "','" _
        & paszBankParam(i, 8) & "','" _
        & paszBankParam(i, 3) & "','" _
        & paszBankParam(i, 4) & "','" _
        & paszBankParam(i, 5) & "','" _
        & paszBankParam(i, 6) & "','" _
        & paszBankParam(i, 7) & "','" _
        & cszNOPrint & "' ) "
        poDb.Execute szSql
    Next i
    poDb.CommitTrans
    InterNetSellTicketLocalTrans = siSellInfo
    Exit Function
Error_Handle:
    poDb.RollbackTrans
    err.Raise err.Number
End Function

Public Function UnInetNetTicket(ByVal pszGetTicketID As String, ByVal pszValiDate As String, ByVal pszUnTkNum As Integer, paszSeatID() As String) As String
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim rsTemp2 As Recordset
    Dim rsTemp3 As Recordset
    Dim i As Integer
    Dim tTSeatInfo() As TSeatInfoCount
    Dim nSaleSeatQuantity As Integer
    Dim szSqlSet As String
    Dim szSeatNoTemplog As String
    Dim j As Integer
    Dim rs As Recordset
    Dim adyDate_new() As Date
    Dim aszBusID_new()  As String
    Dim aszEndStationID_new() As String
    Dim atSellTKParam() As TSellTicketParam
    Dim aszSeatNO() As String
    Dim atBuyTicketInfo() As TBuyTicketInfo
    Dim atSellTicketResultInfo() As TSellTicketResult
    Dim szRs As String
    Dim x As Integer
    oDb.ConnectionString = GetConnectionStr("")

    On Error GoTo Error_Handle
    oDb.BeginTrans
    For x = 1 To ArrayLength(paszSeatID)
        szRs = "SELECT * FROM internet_sell_ticket_info WHERE getticket_id='" & pszGetTicketID & "' AND validate_id='" & pszValiDate & "'" _
                & " AND if_print=2 AND seat_no='" & Trim(paszSeatID(x)) & "'"
        Set rs = oDb.Execute(szRs)
        If rs.RecordCount > 0 Then ShowError ERR_NetBookCancel
    Next x
    szSql = "SELECT * FROM internet_sell_ticket_info WHERE getticket_id='" & pszGetTicketID & "' AND validate_id='" & pszValiDate & "'" _
            & " AND if_print=0 "

    Set rsTemp = oDb.Execute(szSql)
    If pszUnTkNum <> ArrayLength(paszSeatID) Then ShowError ERR_UnReturnNumAndSeatNum
    If rsTemp.RecordCount = 0 Then ShowError ERR_NoNetBook
    If rsTemp.RecordCount < pszUnTkNum Then ShowError ERR_OverNetBookNum
    For i = 1 To pszUnTkNum
        ReDim tTSeatInfo(1 To 5)
        szSql = " UPDATE internet_sell_ticket_info set if_print=2 WHERE bus_id='" & Trim(rsTemp!bus_id) & "' AND bus_date='" & ToDBDate(Trim(rsTemp!bus_date)) & "'" _
            & " AND if_print=0 AND getticket_id='" & pszGetTicketID & "' AND validate_id='" & pszValiDate & "'  AND seat_no='" & Trim(paszSeatID(i)) & "'"
        oDb.Execute szSql
        szSql = "SELECT seat_no FROM Env_bus_seat_lst WHERE bus_id='" & Trim(rsTemp!bus_id) & "'" _
               & " AND  bus_date='" & ToDBDate(Trim(rsTemp!bus_date)) & "'" _
               & "AND seat_no ='" & Trim(paszSeatID(i)) & "'" _
               & " AND (Status <> " & ST_SeatReserved & " AND Status <> " & ST_SeatProjectBooked & " )"
        Set rsTemp2 = oDb.Execute(szSql)
        
        If rsTemp2.RecordCount > 0 Then
         szSeatNoTemplog = ""
         
          For j = 0 To rsTemp2.RecordCount - 1
           szSeatNoTemplog = szSeatNoTemplog & "[" & paszSeatID(i) & "]"
           rsTemp2.MoveNext
          Next
    
        ShowError ERR_SeatNotBooked, "座位:" & szSeatNoTemplog
        
        End If
        szSql = "UPDATE Env_bus_seat_lst SET status=" & Str(ST_SeatCanSell) & " ," _
                & " remark='' " _
                & "    WHERE bus_id='" & Trim(rsTemp!bus_id) & "' AND  bus_date='" & ToDBDate(Trim(rsTemp!bus_date)) & "' " _
                & " AND seat_no  ='" & Trim(paszSeatID(i)) & "'" _
                & " AND (Status = " & ST_SeatReserved & " or Status = " & ST_SeatProjectBooked & " )"
        oDb.Execute szSql
    

        
        szSql = " select seat_type_id ,count(*)as CountSeat from Env_bus_seat_lst " _
                & " where  bus_id='" & Trim(rsTemp!bus_id) & "'and  status='" & ST_SeatCanSell & "' and bus_date='" & ToDBDate(Trim(rsTemp!bus_date)) & "' group by seat_type_id "
        Set rsTemp2 = oDb.Execute(szSql)
        
        For j = 1 To rsTemp2.RecordCount
            tTSeatInfo(CInt(rsTemp2!seat_type_id)).SeatCount = FormatDbValue(rsTemp2!CountSeat)
            rsTemp2.MoveNext
        Next
        szSqlSet = ""
        nSaleSeatQuantity = 0
        szSqlSet = szSqlSet & "seat_remain ='" & tTSeatInfo(1).SeatCount & "',"
        szSqlSet = szSqlSet & "bed_remain ='" & tTSeatInfo(2).SeatCount & "',"
        szSqlSet = szSqlSet & "additional_remain ='" & tTSeatInfo(3).SeatCount & "',"
        szSqlSet = szSqlSet & "other_remain_1 ='" & tTSeatInfo(4).SeatCount & "',"
        szSqlSet = szSqlSet & "other_remain_2 ='" & tTSeatInfo(5).SeatCount & "',"
        
        For j = 1 To 5
          nSaleSeatQuantity = nSaleSeatQuantity + tTSeatInfo(j).SeatCount
        Next
        
        szSqlSet = szSqlSet & "sale_seat_quantity='" & nSaleSeatQuantity & "'"
        szSql = "UPDATE Env_bus_info SET  " & szSqlSet _
                & "  WHERE bus_id='" & Trim(rsTemp!bus_id) & "' AND  bus_date='" & ToDBDate(Trim(rsTemp!bus_date)) & "'"
        oDb.Execute (szSql)

        '是否远程数据库需要预留
        '判断一下，系统参数
        '-----------------------------------------------------------------------------------
'        If bIsUsingRelationBus Then
''            Dim oReBusService As New REBusService
'
'
'            oReBusService.UnReserveSeat Trim(rsTemp!bus_id), ToDBDate(Trim(rsTemp!bus_date)), Trim(paszSeatID(i))
'            '-----------------------------------------------------------------------------------
'        End If
        rsTemp.MoveNext
    Next i
    oDb.CommitTrans
    UnInetNetTicket = "0000"
    Exit Function
Error_Handle:
    oDb.RollbackTrans
    err.Raise err.Number
End Function

'验证网上售票
Public Function InterNetValiDate(Optional pszGetTicketID As String, Optional pszValiDate As String, Optional pszCardID As String, Optional pdtBusDate As Date) As Recordset
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim i As Integer
On Error GoTo Error_Handle
    oDb.ConnectionString = GetConnectionStr("")

    szSql = "SELECT * FROM internet_sell_ticket_info WHERE  if_print=0"
 
    If Trim(pszGetTicketID) <> "" Then
        szSql = szSql & " AND getticket_id='" & pszGetTicketID & "' "
    End If
    
    If Trim(pszValiDate) <> "" Then
        szSql = szSql & " AND validate_id='" & pszValiDate & "'"
    End If
 
    If Trim(pszCardID) <> "" Then
        szSql = szSql & " AND card_id='" & pszCardID & "'"
    End If
    
    If pdtBusDate <> cszEmptyDate Then
        szSql = szSql & " AND bus_date='" & ToDBDate(pdtBusDate) & "'"
    End If
    
    Set rsTemp = oDb.Execute(szSql)
    Set InterNetValiDate = rsTemp
    Exit Function
Error_Handle:
    err.Raise err.Number
End Function
'取网上售票
Public Function SellInetNetTicket(ByVal rsTemp As Recordset, ByVal TicketID As String, Optional pszGetNum As Integer) As TSellTicketResult()
    'pszGetNum 取票数量
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp2 As Recordset
    Dim rsTemp3 As Recordset
    Dim i As Integer
    Dim tTSeatInfo(1 To 5) As TSeatInfoCount
    Dim nSaleSeatQuantity As Integer
    Dim szSqlSet As String
    Dim szSeatNoTemplog As String
    Dim j As Integer
    
    Dim adyDate_new() As Date
    Dim aszBusID_new()  As String
    Dim aszEndStationID_new() As String
    Dim atSellTKParam() As TSellTicketParam
    Dim aszSeatNO() As String
    Dim atBuyTicketInfo() As TBuyTicketInfo
    Dim atSellTicketResultInfo() As TSellTicketResult
    Dim nCount As Integer
    Dim oParam As New SystemParam
    oParam.Init m_oActiveUser
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    If pszGetNum = 0 Then
        nCount = rsTemp.RecordCount
    ElseIf pszGetNum <= rsTemp.RecordCount Then
        nCount = pszGetNum
    ElseIf pszGetNum > rsTemp.RecordCount Then
        err.Raise 2005, err.Source, "网上订票数少于所要取的票数,请注意核对!"
    End If
    
    On Error GoTo Error_Handle
    oDb.BeginTrans
    For i = 1 To nCount
        szSql = "SELECT seat_no FROM Env_bus_seat_lst WHERE bus_id='" & Trim(rsTemp!bus_id) & "'" _
               & " AND  bus_date='" & ToDBDate(Trim(rsTemp!bus_date)) & "'" _
               & "AND seat_no ='" & Trim(rsTemp!seat_no) & "'" _
               & " AND (Status <> " & ST_SeatReserved & " AND Status <> " & ST_SeatProjectBooked & " )"
        Set rsTemp2 = oDb.Execute(szSql)
        
        If rsTemp2.RecordCount > 0 Then
         szSeatNoTemplog = ""
         
          For j = 0 To rsTemp2.RecordCount - 1
           szSeatNoTemplog = szSeatNoTemplog & "[" & FormatDbValue(rsTemp2!seat_no) & "]"
           rsTemp2.MoveNext
          Next
    
        ShowError ERR_SeatNotBooked, "座位:" & szSeatNoTemplog
        
        End If
        szSql = "UPDATE Env_bus_seat_lst SET status=" & Str(ST_SeatCanSell) & " ," _
                & " remark='' " _
                & "    WHERE bus_id='" & Trim(rsTemp!bus_id) & "' AND  bus_date='" & ToDBDate(Trim(rsTemp!bus_date)) & "' " _
                & " AND seat_no  ='" & Trim(rsTemp!seat_no) & "'" _
                & " AND (Status = " & ST_SeatReserved & " or Status = " & ST_SeatProjectBooked & " )"
        oDb.Execute szSql
    

        
        szSql = " select seat_type_id ,count(*)as CountSeat from Env_bus_seat_lst " _
                & " where  bus_id='" & Trim(rsTemp!bus_id) & "'and  status='" & ST_SeatCanSell & "' and bus_date='" & ToDBDate(Trim(rsTemp!bus_date)) & "' group by seat_type_id "
        Set rsTemp2 = oDb.Execute(szSql)
        
        For j = 1 To rsTemp2.RecordCount
            tTSeatInfo(CInt(rsTemp2!seat_type_id)).SeatCount = FormatDbValue(rsTemp2!CountSeat)
            rsTemp2.MoveNext
        Next
        szSqlSet = ""
        nSaleSeatQuantity = 0
        szSqlSet = szSqlSet & "seat_remain ='" & tTSeatInfo(1).SeatCount & "',"
        szSqlSet = szSqlSet & "bed_remain ='" & tTSeatInfo(2).SeatCount & "',"
        szSqlSet = szSqlSet & "additional_remain ='" & tTSeatInfo(3).SeatCount & "',"
        szSqlSet = szSqlSet & "other_remain_1 ='" & tTSeatInfo(4).SeatCount & "',"
        szSqlSet = szSqlSet & "other_remain_2 ='" & tTSeatInfo(5).SeatCount & "',"
        
        For j = 1 To 5
          nSaleSeatQuantity = nSaleSeatQuantity + tTSeatInfo(j).SeatCount
        Next
        
        szSqlSet = szSqlSet & "sale_seat_quantity='" & nSaleSeatQuantity & "'"
        szSql = "UPDATE Env_bus_info SET  " & szSqlSet _
                & "  WHERE bus_id='" & Trim(rsTemp!bus_id) & "' AND  bus_date='" & ToDBDate(Trim(rsTemp!bus_date)) & "'"
        oDb.Execute (szSql)
        ReDim Preserve aszSeatNO(1 To i + 1)
        aszSeatNO(i) = Trim(rsTemp!seat_no)
        rsTemp.MoveNext
    Next i
        Dim SellStationID(1 To 1) As String
        ReDim atBuyTicketInfo(1 To nCount)
        ReDim aszBusID_new(1 To 1)
        ReDim adyDate_new(1 To 1)
        ReDim aszEndStationID_new(1 To 1)
        ReDim atSellTKParam(1 To 1)
        Dim anInsurrance(1 To 1) As Integer
        Dim aszSellStationID() As String
        Dim aszUserID(1 To 1) As String
        Dim aszBankTime(1 To 1) As Date
        rsTemp.MoveFirst
        
        SellStationID(1) = Trim(rsTemp!sell_station_id)

        aszBusID_new(1) = Trim(rsTemp!bus_id)
        adyDate_new(1) = Trim(rsTemp!bus_date)
        aszEndStationID_new(1) = Trim(rsTemp!des_station_id)
        aszUserID(1) = Trim(rsTemp!bank_id)
        aszBankTime(1) = FormatDbValue(rsTemp!busy_date)
        ReDim aszResultSeatID(1 To nCount)
        Dim asgTicketPrices() As Single
        ReDim asgTicketPrices(1 To nCount)
    
        For i = 1 To nCount
        
            atBuyTicketInfo(i).nTicketType = Trim(rsTemp!ticket_type)
            atBuyTicketInfo(i).szReserved = ""
            If i > ArrayLength(aszSeatNO) Then
                '如果传过来的座位号清单的数目小于当前的数字的长数,则出错
                ShowError ERR_BookSeatNotExist
            Else
                atBuyTicketInfo(i).szSeatNo = aszSeatNO(i)
            End If
            Dim szTmp As String
            szTmp = TicketID + i
            atBuyTicketInfo(i).szTicketNo = Left(TicketID, Len(TicketID) - Len(szTmp)) & TicketID + i - 1
'            atBuyTicketInfo(i).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketOffset)
            atBuyTicketInfo(i).szSeatTypeID = "01" '座位类型固定为01
            asgTicketPrices(i) = 0
            rsTemp.MoveNext
        Next i
        atSellTKParam(1).BuyTicketInfo = atBuyTicketInfo
        atSellTKParam(1).pasgSellTicketPrice = asgTicketPrices
        anInsurrance(1) = 0
        atSellTicketResultInfo = SelfSellLocalInterNetTicket(adyDate_new, aszBusID_new, SellStationID, aszEndStationID_new, atSellTKParam, anInsurrance, aszUserID, aszBankTime, oDb)
        SellInetNetTicket = atSellTicketResultInfo
    Dim ssql As String
    rsTemp.MoveFirst
    For i = 1 To nCount '把票号返写回网上售票表.
        ssql = "UPDATE internet_sell_ticket_info SET if_print=1 , print_user_id ='" & m_oActiveUser.UserID & "', ticket_id='" & atBuyTicketInfo(i).szTicketNo & "',print_time='" & oParam.NowDateTime & "' WHERE bus_id='" & Trim(rsTemp!bus_id) & "' AND bus_date='" & ToDBDate(Trim(rsTemp!bus_date)) & "'" _
            & " AND bank_card_id='" & Trim(rsTemp!bank_card_id) & "' AND busy_date='" & Trim(rsTemp!busy_date) & "'  AND seat_no='" & Trim(rsTemp!seat_no) & "' AND if_print=" & cszNOPrint
        oDb.Execute ssql
        rsTemp.MoveNext
    Next i
    oDb.CommitTrans
    Exit Function
Error_Handle:
    oDb.RollbackTrans
    err.Raise err.Number
End Function
'网上售票
Public Function SellInterNetTicket(BusDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String, aspSellTicketParam() As TSellTicketParam, anInsurance() As Integer, aszUserID() As String, Optional oDb As RTConnection) As TSellTicketResult()
    'anInsurance 是 保险费
    
'    Dim srTemp() As TSellTicketResult
'    Dim iLen As Integer
'    Dim iLen1 As Integer
'    Dim nCount As Integer
'    Dim aszDesStationName() As String
'
'    iLen = 0
'    iLen1 = 0
'    iLen = ArrayLength(BusID)
'    ReDim srTemp(1 To iLen)
''    If Trim(m_oCurrentSellService.SelfUser.ServerUnit) <> Trim(m_oCurrentSellService.SelfUser.UserUnitID) Then
''        '如果用户单位和服务单位不一样则是售远程票否则售的是本地票
'        On Error GoTo ErrorHandle
''
''        srTemp = m_oCurrentSellService.SelfSellLocalInterNetTicket(BusDate, BusID, SellStationID, DesStationID, aspSellTicketParam, anInsurance, oDb)
''
''        iLen1 = ArrayLength(srTemp)
''        ReDim aszDesStationName(1 To iLen1)
''        For nCount = 1 To iLen1
''            aszDesStationName(nCount) = srTemp(nCount).szDesStationName
''        Next nCount
''        On Error GoTo 0
''        On Error GoTo ErrorHandle2
''        SelfSellLocalInterNetTicket BusDate, BusID, SellStationID, DesStationID, aspSellTicketParam, anInsurance, oDb
''        On Error GoTo 0
''    Else
'        srTemp = SelfSellLocalInterNetTicket(BusDate, BusID, SellStationID, DesStationID, aspSellTicketParam, anInsurance, oDb)
''    End If
'    SellInterNetTicket = srTemp
'    Exit Function
'ErrorHandle:
'    err.Raise err.Number
'    Exit Function
'ErrorHandle2:
'    '加入提示"对方车站已写入,请通知管理员在对方车站将此票删除"
'    err.Raise err.Number, err.Source, cszServiceFinish & err.Description
'    Exit Function
End Function

'网上内部本地售票
Public Function SelfSellLocalInterNetTicket(BusStartupDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String _
             , aspSellTicketParam() As TSellTicketParam, panInsurance() As Integer, paszUserID() As String, paszBankTime() As Date, Optional oDb As RTConnection) As TSellTicketResult()
    'panInsurance 保险费
    
'    Dim oDb As New RTConnection
    Dim rsTemp() As TSellTicketResult
    Dim nCount As Integer
    Dim iLen As Integer
    iLen = ArrayLength(BusID)

    
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_SellTicket
    ReDim rsTemp(1 To iLen)

'    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    oDb.BeginTrans
    On Error GoTo here
    For nCount = 1 To iLen
        rsTemp(nCount) = SelfSellInterNetTicketLocalTrans(BusStartupDate(nCount), BusID(nCount), SellStationID(nCount), DesStationID(nCount), aspSellTicketParam(nCount).BuyTicketInfo, aspSellTicketParam(nCount).pasgSellTicketPrice, panInsurance(nCount), paszUserID(nCount), paszBankTime(nCount), False, ST_TicketNormal, oDb)
    Next nCount
    SelfSellLocalInterNetTicket = rsTemp
    oDb.CommitTrans
    Exit Function
here:
    oDb.RollbackTrans
    Set oDb = Nothing
    
        err.Raise err.Number
End Function


'网上内部售票事务函数
Private Function SelfSellInterNetTicketLocalTrans(BusStartupDate As Date, BusID As String, SellStationID As String, DesStationID As String, SellTicketInfo() As TBuyTicketInfo _
    , pasgSellTicketPrice() As Single, pnInsurance As Integer, pszUserID As String, paszBankTime As Date, Optional pbExtraSell As Boolean = False _
    , Optional pnTicketStatus As ETicketStatus = ST_TicketNormal, Optional poDb As RTConnection) As TSellTicketResult
    
    Dim oParam As New SystemParam
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim dtBusFullDateTime As Date
    Dim nTemp As Double
    Dim dtNow As Date
    Dim i As Integer
    Dim nBusStatus As EREBusStatus
    Dim nBuyTicketCount As Integer
    Dim sgFullPrice As Double, sgHalfPrice As Double
    Dim lRowAffect As Long, nStationLimitCount As Integer
    Dim szSeatNo As String, sgTicketPrice As Double
    Dim nBusType As EBusType
    Dim atiTicketPrice() As TRETicketPrice
    Dim siSellInfo As TSellTicketResult
    
    Dim lErrorCode As Long
    Dim rsRange As Recordset
    
    Dim nStationRemainCount As Integer
    
    Dim sgPreferentialTicket1 As Single '折扣票1
    Dim sgPreferentialTicket2 As Single '折扣票2
    Dim sgPreferentialTicket3 As Single '折扣票3
    
    
    Dim bNeedBook As Boolean '是否上了预定模块
    Dim bBookValid As Boolean '是否还需判断预定
    Dim bSeatBooked As Boolean
    Dim szBookStationID As String '预定的到站
    Dim rsTemp2 As Recordset
    
    Dim szSeatType As String '座位类型
    Dim szSqlSeat As String  '座位票价
    Dim nSeatStatus As Integer '座位状态
    
    Dim szStopTime As String
    
    Dim nSellTicketType As Integer
    Dim nSellTicketStatus As ETicketStatus
    
    
    Dim bSellStationCanSellEachOther As Boolean
    
    
    ReDim atiTicketPrice(1 To TP_TicketTypeCount)
    
    nSellTicketStatus = pnTicketStatus
    
    Select Case pnTicketStatus
        Case ST_TicketReturned
            nSellTicketType = TP_ExtraSold
            nSellTicketStatus = ST_TicketNormal
            
        Case ST_TicketSellChange
            nSellTicketType = TP_ChangeSold
        Case ST_TicketNormal
            nSellTicketType = TP_NormalSold
        Case Else
            nSellTicketType = 0
        
    End Select
    
    oParam.Init m_oActiveUser
    '判断车次的日期是否超过预售期限
    
    dtNow = oParam.NowDateTime
    
    nTemp = DateDiff("d", dtNow, BusStartupDate)
    
    '如果要购票的车次不在当天和预售天之内则出错
    If nTemp > oParam.PreSellDate Or nTemp < 0 Then ShowError ERR_BusNotInIlegalDate
    If nTemp > oParam.UserCanSellDate Then ShowError ERR_UserCannotSell
    
    bSellStationCanSellEachOther = m_oParam.SellStationCanSellEachOther
    'poDb.ConnectionString = GetConnectionStr(cszSellTicket)
    '得到环境车次及车次站点信息
'    szSql = "SELECT * FROM work_env_bus_info,work_env_bus_station_lst " _
'    & " WHERE work_env_bus_info.bus_id='" & BusID & "' AND " _
'    & " work_env_bus_info.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
'    & " work_env_bus_info.bus_id=work_env_bus_station_lst.bus_id AND " _
'    & " work_env_bus_info.bus_date=work_env_bus_station_lst.bus_date AND " _
'    & " station_id='" & DesStationID & "'" _
'    & " AND work_env_bus_station_lst.sell_station_id='" & SellStationID & "'"
'    Set rsTemp = poDb.Execute(szSql)


    szSql = "SELECT e.bus_date , e.bus_id , e.vehicle_type_name , e.vehicle_type_code , e.total_seat " _
        & " , e.sale_seat_quantity , e.total_stand_quantity , e.sale_stand_seat_quantity , e.bus_start_time , e.bus_kind" _
        & " , e.is_all_refundment , e.route_id , e.end_station_id , e.end_station_name" _
        & " , e.owner_id , e.transport_company_id , e.vehicle_id , e.split_company_id , e.baggage_number , e.fact_weight" _
        & " , e.calculate_weight , e.over_weight_number , e.luggage_total_price  , e.scrollbus_check_time" _
        & " , e.seat_remain , e.bed_remain  , e.additional_remain , e.other_remain_1 , e.other_remain_2" _
        & " , s.sell_station_id , s.station_id , s.seat_type_id , s.station_serial_no , s.max_sale_quantity , s.sale_ticket_quantity" _
        & " , s.book_quantity , s.stop_sale_time , s.mileage  , s.full_price , s.half_price , s.preferential_ticket1 , s.preferential_ticket2" _
        & " , s.preferential_ticket3" _
        & " , a.check_gate_id , a.bus_start_time" _
        & " , a.status , a.can_sale_quantity , a.sell_quantity" _
        & " FROM work_env_bus_info e ,work_env_bus_station_lst s , work_env_bus_allot_lst a" _
        & " WHERE e.bus_id='" & BusID & "' AND  e.bus_date='" & ToDBDate(BusStartupDate) & "'" _
        & " AND  e.bus_id=s.bus_id AND  e.bus_date=s.bus_date" _
        & " AND  s.station_id='" & DesStationID & "' AND s.sell_station_id='" & SellStationID & "'" _
        & " AND a.bus_id = s.bus_id AND a.bus_date = s.bus_date AND a.sell_station_id = '" & SellStationID & "'" _
        
        
    Set rsTemp = poDb.Execute(szSql)
    
    
    
    
    If rsTemp.RecordCount = 0 Then ShowError ERR_NoThisBus '没有指定的车次或车次不过指定的到站
    nBusType = FormatDbValue(rsTemp!bus_kind)
    '流水车次不能补票
'    If nBusType = TP_ScrollBus And pbExtraSell Then ShowError ERR_ScrollBusCanNotExtraSell
    nBusStatus = FormatDbValue(rsTemp!Status)
    '并班车次不能售票
'    If nBusStatus = ST_BusMergeStopped Then ShowError ERR_BusMerged
'    '拆分停班车次不能售票
'    If nBusStatus = ST_BusSlitpStop Then ShowError ERR_BusSlitpStop
'    '车次正在拆分不能售票
'    If nBusStatus >= 32 Then ShowError ERR_BusIsSplit
'    '停检之后不能售票（实际不太可能出现，除非停售时间设得太小了）如果是补票则不限制
''    If pbExtraSell And nBusStatus = ST_BusStopCheck Then ShowError ERR_BusStopCheck
'    If nBusStatus = ST_BusStopCheck Then ShowError ERR_BusStopCheck
'    If nBusStatus = ST_BusStopped Then ShowError ERR_BusStopped
    dtBusFullDateTime = CDate(Format(FormatDbValue(rsTemp!bus_date), cszDateStr) & " " & Format(FormatDbValue(rsTemp!bus_start_time), cszTimeStr))
    '判断是否到了停售时间(如果是补票则要过了正常售票时间才能补)
    If nBusType <> TP_ScrollBus Then '如果是固定车次才限制
        If pbExtraSell Then
'            If DateDiff("s", dtNow, dtBusFullDateTime) >= oParam.StopSellTime * 60 Then ShowError ERR_BusNotStopSellTicket
        Else
'            If DateDiff("s", dtNow, dtBusFullDateTime) < oParam.StopSellTime * 60 Then ShowError ERR_BusStopSellTicket
        End If
    End If
    szStopTime = FormatDbValue(rsTemp!stop_sale_time)
    If nBusType <> TP_ScrollBus Then '如果是固定车次才限制
        If szStopTime <> "-1" Then '如果该车次到指定的站限制售票时间则要进行判断
'            If DateDiff("n", dtNow, dtBusFullDateTime) > GetStopTime(szStopTime) Then ShowError ERR_BusStaionStopSellTicket, "SellTicketClient.SelfSellTicketLocalTrans"
        End If
        bNeedBook = oParam.IsBookValid
        If bNeedBook Then
            bBookValid = IIf(DateAdd("n", oParam.BookTime, dtNow) >= rsTemp!bus_start_time, False, True)
        End If
    Else
        If szStopTime <> "-1" Then
           dtBusFullDateTime = CDate(Format(FormatDbValue(rsTemp!bus_date), cszDateStr) & " " & GetStopTime(szStopTime, False))
'           If DateDiff("n", dtNow, dtBusFullDateTime) < 0 Then ShowError ERR_BusStaionStopSellTicket, "SellTicketService.SelfSellTicket"
        End If
    End If
    nBuyTicketCount = ArrayLength(SellTicketInfo)
    If nBuyTicketCount > 0 Then '要售票的张数大于0
        nStationLimitCount = FormatDbValue(rsTemp!max_sale_quantity)
        nStationRemainCount = FormatDbValue(rsTemp!sale_ticket_quantity)
        If nStationRemainCount > 0 Then  '此处为了折扣票由>= 改为了>
            If nBuyTicketCount > nStationRemainCount Then ShowError ERR_BusReachMaxStationTicketCount
        End If
        
        sgFullPrice = FormatDbValue(rsTemp!full_price)
        sgHalfPrice = FormatDbValue(rsTemp!half_price)
        sgPreferentialTicket1 = FormatDbValue(rsTemp!preferential_ticket1)
        sgPreferentialTicket2 = FormatDbValue(rsTemp!preferential_ticket2)
        sgPreferentialTicket3 = FormatDbValue(rsTemp!preferential_ticket3)

        ReDim siSellInfo.asgTicketPrice(1 To nBuyTicketCount)
        ReDim siSellInfo.aszSeatNO(1 To nBuyTicketCount)
        ReDim siSellInfo.aszTicketType(1 To nBuyTicketCount)
        ReDim siSellInfo.aszSeatType(1 To nBuyTicketCount)
        For i = 1 To nBuyTicketCount
            szSeatNo = ""
            Select Case SellTicketInfo(i).nTicketType
            Case TP_FullPrice, TP_FreeTicket
                sgTicketPrice = sgFullPrice
            Case TP_HalfPrice
                sgTicketPrice = sgHalfPrice
            Case TP_PreferentialTicket1
                sgTicketPrice = sgPreferentialTicket1
            Case TP_PreferentialTicket2
                sgTicketPrice = sgPreferentialTicket2
            Case TP_PreferentialTicket3
                sgTicketPrice = sgPreferentialTicket3
            End Select
            
            
            
            
            If nBusType <> TP_ScrollBus Then '如果是固定车次才限制
                '判断到站的限售张数够用否
                If nStationLimitCount >= 0 Then
                    If nStationLimitCount = 0 Then
'                        ShowError ERR_BusReachMaxStationTicketCount
                    End If
                End If
            End If
            If nBusType <> TP_ScrollBus Then '如果是固定车次才进行座位处理
                If SellTicketInfo(i).szSeatNo <> cszStandTicketSeatNo Then
                    If SellTicketInfo(i).szSeatNo = cszAutoSeatNo Then  '系统自动给出座位号
RetrySeatNo:
                        '取座位该票价和类型
                        If bNeedBook And Not bBookValid Then
                            '如果可以预定
                            szSql = "SELECT p.ticket_price_total AS TicketPrice, s.seat_no AS SeatNo,s.status,p.seat_type_id AS SeatType FROM Env_bus_seat_lst s,work_env_bus_price_lst p" _
                            & " WHERE " _
                            & " s.bus_id='" & BusID & "' AND " _
                            & " s.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & " p.sell_station_id='" & SellStationID & "' AND" _
                            & " s.seat_no=" _
                            & "(SELECT MIN(seat_no) FROM Env_bus_seat_lst WHERE " _
                            & " bus_id='" & BusID & "' AND " _
                            & " bus_date='" & ToDBDate(BusStartupDate) & "' AND (" _
                            & " status=" & ST_SeatCanSell & " OR status=" & ST_SeatBooked & " or status=" & ST_SeatProjectBooked & ") AND seat_type_id ='" & SellTicketInfo(i).szSeatTypeID & "') AND " _
                            & "s.bus_id=p.bus_id AND " _
                            & "s.bus_date=p.bus_date AND "
                            szSql = szSql & "s.seat_type_id=p.seat_type_id AND " _
                                        & "p.station_id='" & DesStationID & "' AND " _
                                        & "p.ticket_type='" & IIf(SellTicketInfo(i).nTicketType = TP_FreeTicket, TP_FullPrice, SellTicketInfo(i).nTicketType) & "'"
                        Else
                            szSql = "SELECT p.ticket_price_total AS TicketPrice,s.seat_no AS SeatNo,s.seat_type_id AS SeatType,p.ticket_price_total AS TicketPrice,s.status  FROM Env_bus_seat_lst s, work_env_bus_price_lst p WHERE " _
                            & " s.bus_id='" & BusID & "' AND " _
                            & " p.sell_station_id='" & SellStationID & "' AND " _
                            & " s.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & "s.seat_no=(SELECT MIN(seat_no) FROM Env_bus_seat_lst WHERE bus_id='" & BusID & "' AND " _
                            & "bus_date='" & ToDBDate(BusStartupDate) & "' AND status=" & ST_SeatCanSell & " AND seat_type_id ='" & SellTicketInfo(i).szSeatTypeID & "') AND " _
                            & "s.bus_id=p.bus_id AND " _
                            & "s.bus_date=p.bus_date AND "
                            szSql = szSql & "s.seat_type_id=p.seat_type_id AND " _
                                        & "p.station_id='" & DesStationID & "' AND " _
                                        & "p.ticket_type='" & IIf(SellTicketInfo(i).nTicketType = TP_FreeTicket, TP_FullPrice, SellTicketInfo(i).nTicketType) & "'"
                        End If
                                              
                        Set rsTemp = poDb.Execute(szSql)
                        If rsTemp.RecordCount = 0 Then ShowError ERR_BusReachMaxTicketCount, "SellTicketClient.SelfSellTicketLocalTrans", "或" & SellTicketInfo(i).szSeatTypeName & "座位已售完。"
                        szSeatNo = FormatDbValue(rsTemp!SeatNo) '得到座位号
                        szSeatType = FormatDbValue(rsTemp!SeatType) '得到座位类型
                        sgTicketPrice = rsTemp!TicketPrice   '得到票价
                        nSeatStatus = rsTemp!Status
                        
                        If bNeedBook And Not bBookValid Then
                            bSeatBooked = IIf(rsTemp!Status = ST_SeatBooked, True, False)
'                            If rsTemp!Status = ST_SeatProjectBooked Then
'                                AssertHaveRight m_oActiveUser, RIGHT_RightOther1 '判断是否有计划预定权限
'                            End If
                        Else
                            bSeatBooked = False
                        End If
                        If bSeatBooked Then
                            szSql = UpdateSeatStatusStr2(BusID, BusStartupDate, szSeatNo, SellTicketInfo(i).szTicketNo, bBookValid)
                        Else
                            szSql = UpdateSeatStatusStr(BusID, BusStartupDate, szSeatNo, SellTicketInfo(i).szTicketNo)
                        End If
                        If rsTemp!Status = ST_SeatProjectBooked Then szSql = UpdateSeatStatusStr3(BusID, BusStartupDate, szSeatNo, SellTicketInfo(i).szTicketNo)
                        poDb.Execute szSql, lRowAffect
                        
                        If lRowAffect = 0 Then GoTo RetrySeatNo
                    Else
                        szSeatNo = SellTicketInfo(i).szSeatNo
                        
                         '取座位该票价和类型
                        szSqlSeat = "SELECT p.ticket_price_total AS TicketPrice, p.seat_type_id AS SeatType,s.status AS SeatStatus " _
                                        & " FROM work_env_bus_price_lst p,Env_bus_seat_lst s WHERE " _
                                        & " p.bus_id='" & BusID & "' AND p.station_id='" & DesStationID & "' AND " _
                                        & " p.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                                        & " p.ticket_type=" & SellTicketInfo(i).nTicketType & " AND " _
                                        & " p.seat_type_id =(SELECT seat_type_id FROM Env_bus_seat_lst WHERE " _
                                        & " bus_id='" & BusID & "' AND bus_date='" & ToDBDate(BusStartupDate) & "' AND seat_no='" & szSeatNo & "') AND " _
                                        & " s.status=(SELECT status FROM Env_bus_seat_lst WHERE " _
                                        & " bus_id='" & BusID & "' AND bus_date='" & ToDBDate(BusStartupDate) & "' AND seat_no='" & szSeatNo & "')" _
                                        & " AND p.sell_station_id ='" & SellStationID & "'" _
                                        & " AND p.bus_id = s.bus_id AND p.bus_date = s.bus_date AND s.seat_no = '" & szSeatNo & "'" '
                                        
                        Set rsTemp = poDb.Execute(szSqlSeat)
                        
                        If rsTemp.RecordCount = 0 Then ShowError ERR_BusReachMaxTicketCount
                        sgTicketPrice = rsTemp!TicketPrice
                        szSeatType = FormatDbValue(rsTemp!SeatType)
                        nSeatStatus = rsTemp!seatstatus
                        
                        szSql = UpdateSeatStatusStr(BusID, BusStartupDate, SellTicketInfo(i).szSeatNo, SellTicketInfo(i).szTicketNo)
                        poDb.Execute szSql, lRowAffect
                        If lRowAffect = 0 Then
                            If bNeedBook Then
                                If nSeatStatus = ST_SeatProjectBooked Then
                                
                                    AssertHaveRight m_oActiveUser, RIGHT_RightOther1
                                    szSql = UpdateSeatStatusStr3(BusID, BusStartupDate, SellTicketInfo(i).szSeatNo, SellTicketInfo(i).szTicketNo)
                                    poDb.Execute szSql, lRowAffect
                                    If lRowAffect = 0 Then ShowError ERR_SeatNoCanNotSell, "SellTicketService.SelfSellTicket"
                                    bSeatBooked = False
                                Else
                                    szSql = UpdateSeatStatusStr2(BusID, BusStartupDate, SellTicketInfo(i).szSeatNo, SellTicketInfo(i).szTicketNo, bBookValid, SellTicketInfo(i).szReserved)
                                    poDb.Execute szSql, lRowAffect
                                    If lRowAffect = 0 Then ShowError ERR_SeatNoCanNotSell, "SellTicketClient.SelfSellTicketLocalTrans"
                                
                                    bSeatBooked = True
                                End If
                            Else
                                ShowError ERR_SeatNoCanNotSell, "SellTicketClient.SelfSellTicketLocalTrans"
                            End If
                        Else
                            bSeatBooked = False
                        End If
                        
                    End If
                    '如果座位预定
                    If bSeatBooked Then
                        szSql = "SELECT * FROM ticket_book_lst WHERE " _
                        & " bus_id='" & BusID & "' AND " _
                        & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                        & " seat_no='" & szSeatNo & "' AND " _
                        & " status=0"
                        Set rsTemp2 = poDb.Execute(szSql)
                        szBookStationID = Trim(rsTemp2!station_id)
                        Set rsTemp2 = Nothing

                        
                        'If nStationLimitCount > 0 Then
                            '更新站点的预定张数
                        szSql = "UPDATE Env_bus_station_lst SET " _
                        & " book_quantity=book_quantity-1 WHERE " _
                        & " bus_id='" & BusID & "' AND " _
                        & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                        & " station_id='" & szBookStationID & "' AND " _
                        & " sell_station_id='" & SellStationID & "' AND " _
                        & " book_quantity>0"
                        
                        poDb.Execute szSql
                        'End If
                        
                        If bBookValid Then
                            szSql = "UPDATE ticket_book_lst SET status=2 WHERE " _
                            & " bus_id='" & BusID & "' AND " _
                            & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & " seat_no='" & szSeatNo & "' AND " _
                            & " status=0 " '0状态为预定
                        Else
                            szSql = "UPDATE ticket_book_lst SET status=3 WHERE " _
                            & " bus_id='" & BusID & "' AND " _
                            & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & " seat_no='" & szSeatNo & "' AND " _
                            & " status=0 " '0状态为预定
                        End If
                        poDb.Execute szSql
                    End If
                    
                    '得到环境车次配载信息
                    Set rsTemp2 = poDb.Execute("SELECT * from work_env_bus_allot_lst WHERE bus_date='" _
                    & ToDBDate(BusStartupDate) & "' AND " _
                    & " bus_id='" & BusID & "' AND " _
                    & " sell_station_id='" & SellStationID & "'")
                    If rsTemp2.RecordCount = 0 Then ShowError ERR_BusAllotInfoNotExist
                    
                    '总共允许售票张数是否大于"已售张数"+1
                    
                    If rsTemp2!can_sale_quantity >= rsTemp2!sell_quantity + 1 Or rsTemp2!can_sale_quantity = -1 Then
                        poDb.Execute ("UPDATE Env_bus_allot_lst SET sell_quantity=sell_quantity + 1 WHERE bus_date='" _
                            & ToDBDate(BusStartupDate) & "' AND " _
                            & " bus_id='" & BusID & "' AND " _
                            & " sell_station_id='" & SellStationID & "'")
                    Else
                        ShowError ERR_SellQuantityISFull '超过允许售票张数
                    End If
                    
                    If bSellStationCanSellEachOther Then
                        '得到环境车次售票点信息
                        Set rsTemp2 = poDb.Execute("SELECT * from work_env_bus_sell_station_lst WHERE bus_date='" _
                        & ToDBDate(BusStartupDate) & "' AND " _
                        & " bus_id='" & BusID & "' AND " _
                        & " sell_station_id='" & m_oActiveUser.SellStationID & "'")
                        If rsTemp2.RecordCount = 0 Then ShowError ERR_BusSellStationInfoNotExist
                        
                        '总共允许售票张数是否大于"已售张数"+1
                        
                        If rsTemp2!can_sale_quantity >= rsTemp2!sell_quantity + 1 Or rsTemp2!can_sale_quantity = -1 Then
                            poDb.Execute ("UPDATE env_bus_sell_station_lst SET sell_quantity=sell_quantity + 1 WHERE bus_date='" _
                                & ToDBDate(BusStartupDate) & "' AND " _
                                & " bus_id='" & BusID & "' AND " _
                                & " sell_station_id='" & m_oActiveUser.SellStationID & "'")
                        Else
                            ShowError ERR_SellQuantityISFull '超过允许售票张数
                        End If
                    End If

                    Set rsTemp2 = Nothing
                    If nSeatStatus = ST_SeatProjectBooked Or nSeatStatus = ST_SeatBooked Then
                        '如果为售预定票,则不减剩余座位数
                    Else
                        '判断到站的限售张数够用否
                        If nStationRemainCount >= 0 Then
                            If nStationRemainCount = 0 Then
'                                ShowError ERR_BusReachMaxStationTicketCount, "SellTicketClient.SelfSellTicketLocalTrans"
                            Else
                                '更新可售张数
                                szSql = "UPDATE Env_bus_station_lst SET " _
                                & " sale_ticket_quantity=sale_ticket_quantity-1 WHERE " _
                                & " bus_id='" & BusID & "' AND " _
                                & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                                & " sale_ticket_quantity>book_quantity AND " _
                                & " station_id='" & DesStationID & "'"
'                                & " sell_station_id='" & SellStationID & "' AND "
                                poDb.Execute szSql, lRowAffect
                                
                                If lRowAffect = 0 Then ShowError ERR_BusReachMaxStationTicketCount, "SellTicketClient.SelfSellTicketLocalTrans"
                            End If
                            
                        End If
                        szSql = "UPDATE Env_bus_info SET sale_seat_quantity=sale_seat_quantity-1 , "
                        Select Case szSeatType
                            Case cszSeatType
                                szSql = szSql & "seat_remain=seat_remain - 1 "
                            Case cszBedType
                                szSql = szSql & "bed_remain=bed_remain - 1 "
                            Case cszAdditionalType
                                szSql = szSql & "additional_remain=additional_remain - 1 "
                            Case cszOtherType1
                                szSql = szSql & "other_remain_1 =other_remain_1 - 1"
                            Case cszOtherType2
                                szSql = szSql & "other_remain_2=other_remain_2 - 1"
                        End Select
                        szSql = szSql & " WHERE bus_id='" & BusID & "' AND bus_date='" & ToDBDate(BusStartupDate) & "'"
                        
                        poDb.Execute szSql
                    End If
                    
                Else
                    szSql = "UPDATE Env_bus_info SET sale_stand_seat_quantity=sale_stand_seat_quantity-1 " _
                    & " WHERE bus_id='" & BusID & "' AND " _
                    & "bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                    & " sale_stand_seat_quantity>0 "
                    
                    poDb.Execute szSql, lRowAffect
                    If lRowAffect = 0 Then ShowError ERR_NoStandTicket, "SellTicketClient.SelfSellTicketLocalTrans"
                    szSeatNo = cszStandTicketSeatNo
                End If
            Else
                szSeatType = cszSeatType
            End If
            '如果票价为0且不为免票,则不允许出售
            If sgTicketPrice = 0 And SellTicketInfo(i).nTicketType <> TP_FreeTicket Then
                RaiseError ERR_CannotSellZeroPrice
            End If
            
            
            '写售票明细表
            szSql = "INSERT INTO Ticket_sell_lst(" _
            & " sell_station_id," _
            & " ticket_id," _
            & " bus_date," _
            & " bus_id," _
            & " des_station_id," _
            & " status," _
            & " ticket_type," _
            & " user_id," _
            & " seat_no," _
            & " ticket_price," _
            & " sale_ticket_type," _
            & " seat_type_id ," _
            & " operation_time, " _
            & " insurance ) "
            
            szSql = szSql & " VALUES('" _
            & SellStationID & "','" _
            & SellTicketInfo(i).szTicketNo & "','" _
            & ToDBDate(BusStartupDate) & "','" _
            & BusID & "','" _
            & DesStationID & "'," _
            & nSellTicketStatus & "," _
            & SellTicketInfo(i).nTicketType & ",'" _
            & pszUserID & "','" _
            & szSeatNo & "'," _
            & sgTicketPrice & "," _
            & nSellTicketType & ",'" _
            & szSeatType & "','" _
            & ToDBDateTime(paszBankTime) & "', " _
            & TransFieldValueToString(pnInsurance) & ")"
            
            On Error GoTo MayBeTicketIDDup
            poDb.Execute szSql
            
            '写详细票价信息
'            If atiTicketPrice(SellTicketInfo(i).nTicketType).sgBase = 0 Then
                szSql = GetTicketPriceDetailStr(BusID, BusStartupDate, SellStationID, DesStationID, szSeatType, IIf(SellTicketInfo(i).nTicketType = TP_FreeTicket, TP_FullPrice, SellTicketInfo(i).nTicketType))
                Set rsTemp = poDb.Execute(szSql)
                LetPriceDetailFromRsToType rsTemp, atiTicketPrice(SellTicketInfo(i).nTicketType), 0
'            End If
            szSql = GetInsertPriceDetailStr(atiTicketPrice(SellTicketInfo(i).nTicketType), SellTicketInfo(i).szTicketNo)
            poDb.Execute szSql
            
            siSellInfo.aszSeatType(i) = szSeatType
            siSellInfo.aszTicketType(i) = SellTicketInfo(i).nTicketType
            siSellInfo.asgTicketPrice(i) = sgTicketPrice
            siSellInfo.aszSeatNO(i) = szSeatNo
        Next
        szSql = "SELECT station_name FROM station_info WHERE station_id='" & DesStationID & "'"
        Set rsTemp = poDb.Execute(szSql)
        
        siSellInfo.szDesStationName = FormatDbValue(rsTemp!station_name)
        
    Else
        ShowError ERR_NoTicketInfoSpecified
    End If
    
    SelfSellInterNetTicketLocalTrans = siSellInfo
    
    
'    Dim nIndex As Integer
'    Dim szStartStationID As String, szStartStationName As String
'    Dim szUnitID As String
'    Dim szDesStationName As String
'    szDesStationName = siSellInfo.szDesStationName
'
'    nBuyTicketCount = ArrayLength(SellTicketInfo)
'    '此处需改为从参数上车站中得到的站点及站点名称
'    szSql = " SELECT s.* , t.station_name FROM sell_station_info s, station_info t WHERE s.station_id = t.station_id "
'    Set rsTemp = poDb.Execute(szSql)
'    szStartStationID = FormatDbValue(rsTemp!station_id)
'    szStartStationName = FormatDbValue(rsTemp!station_name)
'    '不需要从系统参数中得到,改为从传入的参数中得到
'
''    szStartStationID = m_oCurrentSellService.StationID
''    szStartStationName = m_oCurrentSellService.StationName
'
'    szUnitID = m_oActiveUser.UserID
'
'    For nIndex = 1 To nBuyTicketCount
'        '写代理售票明细表
'        szSql = "INSERT INTO Agent_ticket_sale_lst(" _
'        & " sell_station_id," _
'        & " ticket_id," _
'        & " bus_date," _
'        & " bus_id," _
'        & "service_unit_id," _
'        & "start_station_id," _
'        & "start_station_name," _
'        & " des_station_id," _
'        & " des_station_name," _
'        & " ticket_type," _
'        & " user_id," _
'        & "seat_no," _
'        & " ticket_price," _
'        & "status," _
'        & " sale_ticket_type, " _
'        & " seat_type_id," _
'        & " operation_time, " _
'        & " insurance )  " _
'
'
'        szSql = szSql & " VALUES('" _
'        & SellStationID & "','" _
'        & SellTicketInfo(nIndex).szTicketNo & "','" _
'        & ToDBDate(BusStartupDate) & "','" _
'        & BusID & "','" _
'        & szUnitID & "','" _
'        & szStartStationID & "','" _
'        & szStartStationName & "','" _
'        & DesStationID & "','" _
'        & szDesStationName & "'," _
'        & SellTicketInfo(nIndex).nTicketType & ",'" _
'        & m_oActiveUser.UserID & "','" _
'        & siSellInfo.aszSeatNO(nIndex) & "'," _
'        & siSellInfo.asgTicketPrice(nIndex) & "," _
'        & nSellTicketStatus & "," _
'        & nSellTicketType & ",'" _
'        & siSellInfo.aszSeatType(nIndex) & "','" _
'        & ToDBDateTime(dtNow) & "'," _
'        & TransFieldValueToString(pnInsurance) & " ) "
'
'        poDb.Execute szSql
'    Next nIndex
    Set rsTemp = Nothing
    Set rsRange = Nothing
    Set rsTemp2 = Nothing
    Set oParam = Nothing
Exit Function
MayBeTicketIDDup:
    
    If poDb.HaveThisNativeErr(ERR_DBAddDuplicate) Then

        ShowError ERR_TicketNoDup
    Else
        
        err.Raise err.Number
    End If
     Set rsTemp = Nothing
     Set rsRange = Nothing
     Set rsTemp2 = Nothing
     Set oParam = Nothing
End Function
Public Function GetBusExRs(BusDate As Date, BusID As String, SellStationID As String, EndStationID As String) As Recordset

    Dim oDb As New RTConnection
    Dim oParam As New SystemParam
    Dim szSql As String
    Dim rsTemp As Recordset

    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    szSql = "select Work_env_bus_info.* ,e.check_gate_id sell_check_gate_id , sell_station_name,st.station_name  " _
        & " from Work_env_bus_info,Work_env_bus_allot_lst e ,sell_station_info s,station_info st" _
        & " WHERE  e.sell_station_id=s.sell_station_id AND  e.bus_date=Work_env_bus_info.bus_date AND e.bus_id=Work_env_bus_info.bus_id "
        
    szSql = szSql & " AND s.sell_station_id='" & SellStationID & "'and  Work_env_bus_info.bus_date='" & ToDBDate(BusDate) & "' and Work_env_bus_info.bus_id='" & BusID & "'" _
        & " and  st.station_id='" & EndStationID & "'"
    Set rsTemp = oDb.Execute(szSql)
    Set GetBusExRs = rsTemp
    Set rsTemp = Nothing
    Set oDb = Nothing
    
End Function

Public Function FinPrint(BusDate As Date, BusID As String, SeatNo As String, BankID As String) As Boolean

    Dim oDb As New RTConnection
    Dim szSql As String
    Dim iAffic As Long
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    szSql = " UPDATE internet_sell_ticket_info set if_print=1 WHERE bus_id='" & Trim(BusID) & "' AND bus_date='" & ToDBDate(BusDate) & "'" _
            & " AND if_print<>1 AND bank_card_id='" & Trim(BankID) & "'  AND seat_no='" & Trim(SeatNo) & "'"
    oDb.Execute szSql, iAffic
    If iAffic = 1 Then
        FinPrint = True
    Else
        FinPrint = False
    End If
    Set oDb = Nothing
    
End Function

Public Function UnRecordset(ByVal pszGetTicketID As String, ByVal pszValiDate As String) As Recordset
    '查询已取的票
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr("")
    szSql = "SELECT * FROM internet_sell_ticket_info WHERE getticket_id='" & pszGetTicketID & "' AND validate_id='" & pszValiDate & "'" _
            & " AND if_print=1 "

    Set rsTemp = oDb.Execute(szSql)
    Set UnRecordset = rsTemp
    Set oDb = Nothing
    Set rsTemp = Nothing
End Function


Public Function SelfSellTK(BusStartupDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String _
             , DesStationName() As String, aspSellTicketParam() As TSellTicketParam, panInsurance() As Integer, bScroll As Boolean) As TSellTicketResult()
    '自助售票
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp2 As Recordset
    Dim rsTemp3 As Recordset
    Dim i As Integer
    Dim tTSeatInfo(1 To 5) As TSeatInfoCount
    Dim nSaleSeatQuantity As Integer
    Dim szSqlSet As String
    Dim szSeatNoTemplog As String
    Dim aszSeatNO() As String
    Dim j As Integer
    Dim atSellTicketResultInfo() As TSellTicketResult

    Dim TicketID As String
    Dim szEndTicket As String
    Dim rs As Recordset

    oDb.ConnectionString = GetConnectionStr(cszSellTicket)

    On Error GoTo Error_Handle
    oDb.BeginTrans
    
    szSql = "select parameter_value from system_param_info  where parameter_name='StrTicketID'"
    Set rs = oDb.Execute(szSql)
    TicketID = Trim(rs!parameter_value)
    m_lTicketNo = ResolveTicketNo(TicketID, m_szTicketPrefix)


    If bScroll = False Then
    For i = 1 To ArrayLength(aspSellTicketParam(1).BuyTicketInfo)
        szSql = "SELECT seat_no FROM Env_bus_seat_lst WHERE bus_id='" & Trim(BusID(1)) & "'" _
               & " AND  bus_date='" & ToDBDate(Trim(BusStartupDate(1))) & "'" _
               & "AND seat_no ='" & Trim(aspSellTicketParam(1).BuyTicketInfo(i).szSeatNo) & "'" _
               & " AND (Status <> " & ST_SeatReserved & " AND Status <> " & ST_SeatProjectBooked & " )"
        Set rsTemp2 = oDb.Execute(szSql)
        
        If rsTemp2.RecordCount > 0 Then
         szSeatNoTemplog = ""
         
          For j = 0 To rsTemp2.RecordCount - 1
           szSeatNoTemplog = szSeatNoTemplog & "[" & FormatDbValue(rsTemp2!seat_no) & "]"
           rsTemp2.MoveNext
          Next
    
        ShowError ERR_SeatNotBooked, "座位:" & szSeatNoTemplog
        
        End If
        szSql = "UPDATE Env_bus_seat_lst SET status=" & Str(ST_SeatCanSell) & " ," _
                & " remark='' " _
                & "    WHERE bus_id='" & Trim(BusID(1)) & "' AND  bus_date='" & ToDBDate(Trim(BusStartupDate(1))) & "' " _
                & " AND seat_no  ='" & Trim(aspSellTicketParam(1).BuyTicketInfo(i).szSeatNo) & "'" _
                & " AND (Status = " & ST_SeatReserved & " or Status = " & ST_SeatProjectBooked & " )"
        oDb.Execute szSql
    

        
        szSql = " select seat_type_id ,count(*)as CountSeat from Env_bus_seat_lst " _
                & " where  bus_id='" & Trim(BusID(1)) & "'and  status='" & ST_SeatCanSell & "' and bus_date='" & ToDBDate(Trim(BusStartupDate(1))) & "' group by seat_type_id "
        Set rsTemp2 = oDb.Execute(szSql)
        
        For j = 1 To rsTemp2.RecordCount
            tTSeatInfo(CInt(rsTemp2!seat_type_id)).SeatCount = FormatDbValue(rsTemp2!CountSeat)
            rsTemp2.MoveNext
        Next
        szSqlSet = ""
        nSaleSeatQuantity = 0
        szSqlSet = szSqlSet & "seat_remain ='" & tTSeatInfo(1).SeatCount & "',"
        szSqlSet = szSqlSet & "bed_remain ='" & tTSeatInfo(2).SeatCount & "',"
        szSqlSet = szSqlSet & "additional_remain ='" & tTSeatInfo(3).SeatCount & "',"
        szSqlSet = szSqlSet & "other_remain_1 ='" & tTSeatInfo(4).SeatCount & "',"
        szSqlSet = szSqlSet & "other_remain_2 ='" & tTSeatInfo(5).SeatCount & "',"
        
        For j = 1 To 5
          nSaleSeatQuantity = nSaleSeatQuantity + tTSeatInfo(j).SeatCount
        Next
        
        szSqlSet = szSqlSet & "sale_seat_quantity='" & nSaleSeatQuantity & "'"
        szSql = "UPDATE Env_bus_info SET  " & szSqlSet _
                & "  WHERE bus_id='" & Trim(BusID(1)) & "' AND  bus_date='" & ToDBDate(Trim(BusStartupDate(1))) & "'"
        oDb.Execute (szSql)
        ReDim Preserve aszSeatNO(1 To i + 1)
        aszSeatNO(i) = Trim(aspSellTicketParam(1).BuyTicketInfo(i).szSeatNo)

    Next i
    End If
    For i = 1 To ArrayLength(aspSellTicketParam(1).BuyTicketInfo)
        aspSellTicketParam(1).BuyTicketInfo(i).szTicketNo = GetTicketNo(1)
        szEndTicket = Trim(aspSellTicketParam(1).BuyTicketInfo(i).szTicketNo)
        m_lTicketNo = m_lTicketNo + 1
    Next i
        atSellTicketResultInfo = SellTicket(BusStartupDate, BusID, SellStationID, DesStationID, aspSellTicketParam, panInsurance, oDb)
        SelfSellTK = atSellTicketResultInfo
        szSql = "update system_param_info set parameter_value='" & szEndTicket & "' where parameter_name='StrTicketID'"
        oDb.Execute szSql
    oDb.CommitTrans
    Exit Function
Error_Handle:
    oDb.RollbackTrans
    err.Raise err.Number
End Function


Public Function BookTK(BusStartupDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String _
             , aspSellTicketParam() As TSellTicketParam, panInsurance() As Integer) As TSellTicketResult() '自助售票时先做预留
    'panInsurance 保险费

    Dim oDb As New RTConnection
    Dim rsTemp() As TSellTicketResult
    Dim nCount As Integer
    Dim iLen As Integer
    iLen = ArrayLength(BusID)


    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_SellTicket
    ReDim rsTemp(1 To iLen)

    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    oDb.BeginTrans
    On Error GoTo here
    For nCount = 1 To iLen
        rsTemp(nCount) = BookLocalTrans(BusStartupDate(nCount), BusID(nCount), SellStationID(nCount), DesStationID(nCount), aspSellTicketParam(nCount).BuyTicketInfo, aspSellTicketParam(nCount).pasgSellTicketPrice, panInsurance(nCount), False, ST_TicketNormal, oDb)
    Next nCount
    BookTK = rsTemp
    oDb.CommitTrans
    Set oDb = Nothing
    Exit Function
here:
    oDb.RollbackTrans
    Set oDb = Nothing
    
        err.Raise err.Number
End Function



Private Function BookLocalTrans(BusStartupDate As Date, BusID As String, SellStationID As String, DesStationID As String, SellTicketInfo() As TBuyTicketInfo _
    , pasgSellTicketPrice() As Single, pnInsurance As Integer, Optional pbExtraSell As Boolean = False _
    , Optional pnTicketStatus As ETicketStatus = ST_TicketNormal, Optional poDb As RTConnection) As TSellTicketResult '自助售票时先做预留
    
 
    Dim nBookSeatCount As Integer, i As Integer '预留座位数
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szBookID As String
    Dim oParam As New SystemParam
    Dim rsTemp As Recordset, rsTemp2 As Recordset
    Dim lRowAffect As Long
    Dim szTemp As String
    Dim nStationLimitedCount As Integer
    
    Dim BookID() As TBuyTicketInfo '预留票号
    Dim BookSeatID() As String '预留座位号
    
    Dim szSellStationID As String
    Dim szSeatID As String '座位号
    
    
    
    
    Dim dtBusFullDateTime As Date
    Dim nTemp As Double
    Dim dtNow As Date
    Dim nBusStatus As EREBusStatus
    Dim nBuyTicketCount As Integer
    Dim sgFullPrice As Double, sgHalfPrice As Double
    Dim nStationLimitCount As Integer
    Dim szSeatNo As String, sgTicketPrice As Double
    Dim nBusType As EBusType
    Dim atiTicketPrice() As TRETicketPrice
    Dim siSellInfo As TSellTicketResult
    
    Dim lErrorCode As Long
    Dim rsRange As Recordset
    
    Dim nStationRemainCount As Integer
    Dim nCanSellTicket As Integer '可售座位数
    Dim bNeedBook As Boolean '是否上了预定模块
    Dim bBookValid As Boolean '是否还需判断预定
    Dim bSeatBooked As Boolean
    Dim szBookStationID As String '预定的到站
    
    Dim szSeatType As String '座位类型
    Dim szSqlSeat As String  '座位票价
    Dim nSeatStatus As Integer '座位状态
    
    Dim szStopTime As String
    
    Dim nSellTicketType As Integer
    Dim nSellTicketStatus As ETicketStatus
    Dim tTSeatInfo() As TSeatInfoCount
    Dim j As Integer
    Dim szSqlSet As String
    Dim nSaleSeatQuantity As Integer
    
    oParam.Init m_oActiveUser

    dtNow = oParam.NowDateTime
    
    nTemp = DateDiff("d", dtNow, BusStartupDate)
    
    '如果要购票的车次不在当天和预售天之内则出错
    If nTemp > oParam.PreSellDate Or nTemp < 0 Then ShowError ERR_BusNotInIlegalDate
    If nTemp > oParam.UserCanSellDate Then ShowError ERR_UserCannotSell


   
'    '如果票价为0且不为免票,则不允许出售
'    For i = 1 To ArrayLength(pszTicketID)
'    If pszTPrice(1).pasgSellTicketPrice(i) = 0 And pszTicketID(i).nTicketType <> TP_FreeTicket Then
'        RaiseError ERR_CannotSellZeroPrice
'    End If
'    Next

    
    

    oDb.ConnectionString = GetConnectionStr("")
    
    szSql = "SELECT * FROM work_env_bus_info e,work_env_bus_station_lst s " _
    & " WHERE e.bus_id='" & BusID & "' AND " _
    & " e.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
    & " e.bus_id=s.bus_id AND " _
    & " e.bus_date=s.bus_date AND " _
    & " station_id='" & DesStationID & "'"
    
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount < 1 Then
        ShowError ERR_BusNotExist, "CBook.Book", "BusID=" & BusID & ",BusDate=" & ToDBDate(BusStartupDate)
    Else
        szSellStationID = FormatDbValue(rsTemp!sell_station_id)
        nStationLimitedCount = rsTemp!max_sale_quantity
        If nStationLimitedCount = 0 Then ShowError ERR_BusStationMatchLimitedCount, "CBook.Book"
    End If
    
'*********************************************************
    nBusType = FormatDbValue(rsTemp!bus_kind)
    '流水车次不能补票
'    If nBusType = TP_ScrollBus And pbExtraSell Then ShowError ERR_ScrollBusCanNotExtraSell
    nBusStatus = FormatDbValue(rsTemp!Status)
    '并班车次不能售票
    If nBusStatus = ST_BusMergeStopped Then ShowError ERR_BusMerged
    '拆分停班车次不能售票
    If nBusStatus = ST_BusSlitpStop Then ShowError ERR_BusSlitpStop
    '车次正在拆分不能售票
    If nBusStatus >= 32 Then ShowError ERR_BusIsSplit
    '停检之后不能售票（实际不太可能出现，除非停售时间设得太小了）如果是补票则不限制
'    If pbExtraSell And nBusStatus = ST_BusStopCheck Then ShowError ERR_BusStopCheck
    If nBusStatus = ST_BusStopCheck Then ShowError ERR_BusStopCheck
    If nBusStatus = ST_BusStopped Then ShowError ERR_BusStopped
    dtBusFullDateTime = CDate(Format(FormatDbValue(rsTemp!bus_date), cszDateStr) & " " & Format(FormatDbValue(rsTemp!bus_start_time), cszTimeStr))
    '判断是否到了停售时间(如果是补票则要过了正常售票时间才能补)
'    If nBusType <> TP_ScrollBus Then '如果是固定车次才限制
'        If pbExtraSell Then
'            If DateDiff("s", dtNow, dtBusFullDateTime) >= oParam.StopSellTime * 60 Then ShowError ERR_BusNotStopSellTicket
'        Else
''            If DateDiff("s", dtNow, dtBusFullDateTime) < oParam.StopSellTime * 60 Then ShowError ERR_BusStopSellTicket
'        End If
'    End If



    szStopTime = FormatDbValue(rsTemp!stop_sale_time)
    If nBusType <> TP_ScrollBus Then '如果是固定车次才限制
        If szStopTime <> "-1" Then '如果该车次到指定的站限制售票时间则要进行判断
            If DateDiff("n", dtNow, dtBusFullDateTime) > GetStopTime(szStopTime) Then ShowError ERR_BusStaionStopSellTicket, "SellTicketClient.SelfSellTicketLocalTrans"
        End If
        bNeedBook = oParam.IsBookValid
        If bNeedBook Then
            bBookValid = IIf(DateAdd("n", oParam.BookTime, dtNow) >= rsTemp!bus_start_time, False, True)
        End If
    Else
        If szStopTime <> "-1" Then
           dtBusFullDateTime = CDate(Format(FormatDbValue(rsTemp!bus_date), cszDateStr) & " " & GetStopTime(szStopTime, False))
           If DateDiff("n", dtNow, dtBusFullDateTime) < 0 Then ShowError ERR_BusStaionStopSellTicket, "SellTicketService.SelfSellTicket"
        End If
    End If
    nBuyTicketCount = ArrayLength(SellTicketInfo)
    nStationLimitCount = FormatDbValue(rsTemp!max_sale_quantity)
    nStationRemainCount = FormatDbValue(rsTemp!sale_ticket_quantity)
    nCanSellTicket = FormatDbValue(rsTemp!sale_seat_quantity)
    If nStationRemainCount > 0 Then  '此处为了折扣票由>= 改为了>
            If nBuyTicketCount > nStationRemainCount Then ShowError ERR_BusReachMaxStationTicketCount
    End If
    
    '判断可售座位数
    nBookSeatCount = ArrayLength(SellTicketInfo)
    
    If nCanSellTicket < nBookSeatCount Then
        ShowError ERR_BusReachMaxTicketCount
    End If
    
    
    '得到环境车次配载信息
    Set rsTemp2 = oDb.Execute("SELECT * from work_env_bus_allot_lst WHERE bus_date='" _
                    & ToDBDate(BusStartupDate) & "' AND " _
                    & " bus_id='" & BusID & "' AND " _
                    & " sell_station_id='" & SellStationID & "'")
    If rsTemp2.RecordCount = 0 Then ShowError ERR_BusAllotInfoNotExist
                    
    '总共允许售票张数是否大于"已售张数"+1
                    
    If rsTemp2!can_sale_quantity >= rsTemp2!sell_quantity + 1 Or rsTemp2!can_sale_quantity = -1 Then

    Else
        ShowError ERR_SellQuantityISFull '超过允许售票张数
    End If
    ReDim siSellInfo.aszSeatNO(1 To nBookSeatCount)
    ReDim siSellInfo.asgTicketPrice(1 To nBookSeatCount)
    ReDim siSellInfo.aszSeatType(1 To nBookSeatCount)
    ReDim siSellInfo.aszTicketType(1 To nBookSeatCount)
'*********************************************************
    
    oDb.BeginTrans
    On Error GoTo Error_Handle
    
    For i = 1 To nBookSeatCount
        ReDim tTSeatInfo(1 To 5)
        szSqlSet = ""
        szSql = "SELECT MIN(seat_no) seat_no FROM Env_bus_seat_lst WHERE " _
        & " bus_id='" & BusID & "' AND " _
        & " is_special_price = 0 AND " _
        & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
        & " status=" & ST_SeatCanSell & " AND seat_type_id ='" & SellTicketInfo(i).szSeatTypeID & "'"
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount > 0 Then
            szSeatID = rsTemp!seat_no
        Else
            ShowError ERR_BOOKSeatBooked, "CBook.Book", "BusID=" & BusID & ",BusDate=" & ToDBDate(BusStartupDate) & ",SeatNo=" & szSeatID
        End If
        
        szSql = "SELECT is_special_price FROM Env_bus_seat_lst WHERE " _
            & " bus_id='" & BusID & "' AND " _
            & " seat_no='" & szSeatID & "' AND " _
            & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
            & " status=" & ST_SeatCanSell & " AND seat_type_id ='" & SellTicketInfo(i).szSeatTypeID & "'"
        Set rsTemp = poDb.Execute(szSql)
        '特价座位不能预订
        If rsTemp.RecordCount > 0 Then
'            If FormatDbValue(rsTemp!is_special_price) = 1 Then ShowError ERR_SpecialPriceSeatCanNotSell, "CBook.Book"
        End If
        
            szSql = "UPDATE Env_bus_seat_lst SET status=" & Str(ST_SeatReserved) & "   ," _
                  & " remark='self于" & Date & "留' " _
                  & "  WHERE bus_id='" & BusID & "' " _
                  & " AND  bus_date='" & ToDBDate(BusStartupDate) & "' " _
                  & " AND seat_no ='" & szSeatID & "' " _
                  & " AND  Status=" & ST_SeatCanSell & ""
            oDb.Execute szSql, lRowAffect
            If lRowAffect <> 1 Then ShowError ERR_BusStationMatchLimitedCount, "CBook.Book"
        
        
            szSql = " select seat_type_id ,count(*)as CountSeat from Env_bus_seat_lst " _
                    & " where  bus_id='" & BusID & "'and is_special_price = 0 and  status='" & ST_SeatCanSell & "' and bus_date='" & ToDBDate(BusStartupDate) & "' group by seat_type_id "
            Set rsTemp = oDb.Execute(szSql)
            
            For j = 1 To rsTemp.RecordCount
                tTSeatInfo(CInt(rsTemp!seat_type_id)).SeatCount = FormatDbValue(rsTemp!CountSeat)
                rsTemp.MoveNext
            Next
            
            szSqlSet = szSqlSet & "seat_remain ='" & tTSeatInfo(1).SeatCount & "',"
            szSqlSet = szSqlSet & "bed_remain ='" & tTSeatInfo(2).SeatCount & "',"
            szSqlSet = szSqlSet & "additional_remain ='" & tTSeatInfo(3).SeatCount & "',"
            szSqlSet = szSqlSet & "other_remain_1 ='" & tTSeatInfo(4).SeatCount & "',"
            szSqlSet = szSqlSet & "other_remain_2 ='" & tTSeatInfo(5).SeatCount & "',"
            nSaleSeatQuantity = 0
            For j = 1 To 5
              nSaleSeatQuantity = nSaleSeatQuantity + tTSeatInfo(j).SeatCount
            Next
            
            szSqlSet = szSqlSet & "sale_seat_quantity='" & nSaleSeatQuantity & "'"
            szSql = "UPDATE Env_bus_info SET  " & szSqlSet _
                    & "  WHERE bus_id='" & BusID & "' AND  bus_date='" & ToDBDate(BusStartupDate) & "'"
            oDb.Execute (szSql)
            
            
            
            szSqlSeat = "SELECT p.ticket_price_total AS TicketPrice, p.seat_type_id AS SeatType,s.status AS SeatStatus " _
                            & " FROM work_env_bus_price_lst p,Env_bus_seat_lst s WHERE " _
                            & " p.bus_id='" & BusID & "' AND p.station_id='" & DesStationID & "' AND " _
                            & " p.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & " p.ticket_type=" & SellTicketInfo(i).nTicketType & " AND " _
                            & " p.seat_type_id =(SELECT seat_type_id FROM Env_bus_seat_lst WHERE " _
                            & " bus_id='" & BusID & "' AND bus_date='" & ToDBDate(BusStartupDate) & "' AND seat_no='" & szSeatID & "') AND " _
                            & " s.status=(SELECT status FROM Env_bus_seat_lst WHERE " _
                            & " bus_id='" & BusID & "' AND bus_date='" & ToDBDate(BusStartupDate) & "' AND seat_no='" & szSeatID & "')" _
                            & " AND p.sell_station_id ='" & SellStationID & "'" _
                            & " AND p.bus_id = s.bus_id AND p.bus_date = s.bus_date AND s.seat_no = '" & szSeatID & "'" '
                                        
            Set rsTemp = oDb.Execute(szSqlSeat)
            sgTicketPrice = rsTemp!TicketPrice
            szSeatType = FormatDbValue(rsTemp!SeatType)
            
            siSellInfo.aszSeatNO(i) = szSeatID
            siSellInfo.asgTicketPrice(i) = sgTicketPrice
            siSellInfo.aszSeatType(i) = szSeatType
            siSellInfo.aszTicketType(i) = SellTicketInfo(i).nTicketType
            '是否远程数据库需要预留
            '判断一下，系统参数
'            '-----------------------------------------------------------------------------------
'            If bIsUsingRelationBus Then
'
'                oReBusService.ReserveSeat BusID, ToDBDate(BusStartupDate), szSeatID, "", ST_LockReserve
'                '-----------------------------------------------------------------------------------
'            End If
            
    Next i
    oDb.CommitTrans
    BookLocalTrans = siSellInfo
    Exit Function
Error_Handle:
    oDb.RollbackTrans
    err.Raise err.Number
End Function
Public Function UnBookTK(paszBusID() As String, paszBusDate() As Date, paszSellTicket() As TSellTicketParam) As String '取消预留
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim rsTemp2 As Recordset
    Dim rsTemp3 As Recordset
    Dim i As Integer
    Dim tTSeatInfo(1 To 5) As TSeatInfoCount
    Dim nSaleSeatQuantity As Integer
    Dim szSqlSet As String
    Dim szSeatNoTemplog As String
    Dim j As Integer
    Dim rs As Recordset
    Dim adyDate_new() As Date
    Dim aszBusID_new()  As String
    Dim aszEndStationID_new() As String
    Dim atSellTKParam() As TSellTicketParam
    Dim aszSeatNO() As String
    Dim atBuyTicketInfo() As TBuyTicketInfo
    Dim atSellTicketResultInfo() As TSellTicketResult
    Dim szRs As String
    oDb.ConnectionString = GetConnectionStr("")

    On Error GoTo Error_Handle
    oDb.BeginTrans

    For i = 1 To ArrayLength(paszSellTicket(1).BuyTicketInfo)
 
        szSql = "SELECT seat_no FROM Env_bus_seat_lst WHERE bus_id='" & Trim(paszBusID(1)) & "'" _
               & " AND  bus_date='" & ToDBDate(paszBusDate(1)) & "'" _
               & "AND seat_no ='" & Trim(paszSellTicket(1).BuyTicketInfo(i).szSeatNo) & "'" _
               & " AND (Status <> " & ST_SeatReserved & " AND Status <> " & ST_SeatProjectBooked & " )"
        Set rsTemp2 = oDb.Execute(szSql)
        
        If rsTemp2.RecordCount > 0 Then
         szSeatNoTemplog = ""
         
          For j = 0 To rsTemp2.RecordCount - 1
           szSeatNoTemplog = szSeatNoTemplog & "[" & FormatDbValue(rsTemp2!seat_no) & "]"
           rsTemp2.MoveNext
          Next
    
        ShowError ERR_SeatNotBooked, "座位:" & szSeatNoTemplog
        
        End If
        szSql = "UPDATE Env_bus_seat_lst SET status=" & Str(ST_SeatCanSell) & " ," _
                & " remark='' " _
                & "    WHERE bus_id='" & Trim(paszBusID(1)) & "' AND  bus_date='" & ToDBDate(Trim(paszBusDate(1))) & "' " _
                & " AND seat_no  ='" & Trim(paszSellTicket(1).BuyTicketInfo(i).szSeatNo) & "'" _
                & " AND (Status = " & ST_SeatReserved & " or Status = " & ST_SeatProjectBooked & " )"
        oDb.Execute szSql
    

        
        szSql = " select seat_type_id ,count(*)as CountSeat from Env_bus_seat_lst " _
                & " where  bus_id='" & Trim(paszBusID(1)) & "'and  status='" & ST_SeatCanSell & "' and bus_date='" & ToDBDate(Trim(paszBusDate(1))) & "' group by seat_type_id "
        Set rsTemp2 = oDb.Execute(szSql)
        
        For j = 1 To rsTemp2.RecordCount
            tTSeatInfo(CInt(rsTemp2!seat_type_id)).SeatCount = FormatDbValue(rsTemp2!CountSeat)
            rsTemp2.MoveNext
        Next
        szSqlSet = ""
        nSaleSeatQuantity = 0
        szSqlSet = szSqlSet & "seat_remain ='" & tTSeatInfo(1).SeatCount & "',"
        szSqlSet = szSqlSet & "bed_remain ='" & tTSeatInfo(2).SeatCount & "',"
        szSqlSet = szSqlSet & "additional_remain ='" & tTSeatInfo(3).SeatCount & "',"
        szSqlSet = szSqlSet & "other_remain_1 ='" & tTSeatInfo(4).SeatCount & "',"
        szSqlSet = szSqlSet & "other_remain_2 ='" & tTSeatInfo(5).SeatCount & "',"
        
        For j = 1 To 5
          nSaleSeatQuantity = nSaleSeatQuantity + tTSeatInfo(j).SeatCount
        Next
        
        szSqlSet = szSqlSet & "sale_seat_quantity='" & nSaleSeatQuantity & "'"
        szSql = "UPDATE Env_bus_info SET  " & szSqlSet _
                & "  WHERE bus_id='" & Trim(paszBusID(1)) & "' AND  bus_date='" & ToDBDate(Trim(paszBusDate(1))) & "'"
        oDb.Execute (szSql)
        '是否远程数据库需要预留
        '判断一下，系统参数
        '-----------------------------------------------------------------------------------
'        If bIsUsingRelationBus Then
''            Dim oReBusService As New REBusService
'
'
'            oReBusService.UnReserveSeat Trim(paszBusID(1)), ToDBDate(Trim(paszBusDate(1))), Trim(paszSellTicket(1).BuyTicketInfo(i).szSeatNo)
'            '-----------------------------------------------------------------------------------
'        End If

    Next i
    oDb.CommitTrans

    Exit Function
Error_Handle:
    oDb.RollbackTrans
    err.Raise err.Number
End Function

Public Function CancelPreNetTK(BusStartupDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String _
             , aspSellTicketParam() As TSellTicketParam, panInsurance() As Integer, aszBankParam() As String, paszTicketID() As String, _
             pszGetTicketID As String, pszValiDate As String, pszNewGetTkID As String, pszNewValiDate As String, TicketID As String) As TSellTicketResult()
    '重打网上售票
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim i As Integer
    Dim j As Integer
    Dim nLen As Integer
    Dim szRs As String
    Dim rs As Recordset
    Dim atResult() As TSellTicketResult
    Dim rsTemp As Recordset
    Dim nCount As Integer
    Dim iLen As Integer
    Dim rs2 As Recordset
    Dim szTemp As String
    Dim oParam As New SystemParam
    oParam.Init m_oActiveUser
    nLen = 0
    nLen = ArrayLength(paszTicketID)
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_CancelTicket
    iLen = ArrayLength(BusID)
    szTemp = ""
    ReDim atResult(1 To iLen)
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    oDb.BeginTrans
'****************************************改网上售票表状态,3为废网上票
        szRs = "SELECT * FROM internet_sell_ticket_info WHERE getticket_id='" & pszGetTicketID & "' AND validate_id='" & pszValiDate & "'"
        Set rs = oDb.Execute(szRs)

        szRs = "SELECT * FROM internet_sell_ticket_info WHERE getticket_id='" & pszGetTicketID & "' AND validate_id='" & pszValiDate & "'" _
                & " AND if_print=2 "
        Set rs2 = oDb.Execute(szRs)
        
        If rs.RecordCount = rs2.RecordCount Then ShowError ERR_NetBookCancel  '座位已退
        For i = 1 To ArrayLength(paszTicketID)
            szTemp = szTemp & "," & TransFieldValueToString(paszTicketID(i))
        Next i
        szTemp = " ticket_id in (" & Mid(szTemp, 2, Len(szTemp) - 1) & ")"
        szSql = "SELECT * FROM internet_sell_ticket_info WHERE getticket_id='" & pszGetTicketID & "' AND validate_id='" & pszValiDate & "'" _
                & " AND if_print=1  AND " & szTemp
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount = 0 Then ShowError ERR_NoGetNet
        For j = 1 To rsTemp.RecordCount
            szSql = " UPDATE internet_sell_ticket_info set if_print=3 WHERE bus_id='" & Trim(rsTemp!bus_id) & "' AND bus_date='" & ToDBDate(Trim(rsTemp!bus_date)) & "'" _
                & " AND if_print=1  AND getticket_id='" & pszGetTicketID & "' AND validate_id='" & pszValiDate & "'  AND seat_no='" & Trim(rsTemp!seat_no) & "'"
            oDb.Execute szSql
            rsTemp.MoveNext
            
        Next j
'*********************************************
        On Error GoTo here
        SelfCancelTicket paszTicketID, oDb, TP_NormalCancel '真正的废票
        ReInsertNetTK rsTemp, oDb '插入新的一条记录
        
    Dim nGetNum As Integer
    Dim szSqlSet As String
    Dim szSeatNoTemplog As String
    Dim adyDate_new() As Date
    Dim aszBusID_new()  As String
    Dim aszEndStationID_new() As String
    Dim atSellTKParam() As TSellTicketParam
    Dim aszSeatNO() As String
    Dim atBuyTicketInfo() As TBuyTicketInfo
    Dim atSellTicketResultInfo() As TSellTicketResult
    Dim SellStationIDRX(1 To 1) As String
    ReDim atBuyTicketInfo(1 To nLen)
    ReDim aszBusID_new(1 To 1)
    ReDim adyDate_new(1 To 1)
    ReDim aszEndStationID_new(1 To 1)
    ReDim atSellTKParam(1 To 1)
    Dim anInsurrance(1 To 1) As Integer
    Dim aszSellStationID() As String
    Dim aszUserID(1 To 1) As String
    Dim aszBankTime(1 To 1) As Date
                nGetNum = rsTemp.RecordCount


                
                rsTemp.MoveFirst
                
                SellStationIDRX(1) = Trim(rsTemp!sell_station_id)
        
                aszBusID_new(1) = Trim(rsTemp!bus_id)
                adyDate_new(1) = Trim(rsTemp!bus_date)
                aszEndStationID_new(1) = Trim(rsTemp!des_station_id)
                aszUserID(1) = Trim(rsTemp!bank_id)
                aszBankTime(1) = FormatDbValue(rsTemp!busy_date)
                ReDim aszResultSeatID(1 To nGetNum)
                Dim asgTicketPrices() As Single
                ReDim asgTicketPrices(1 To nGetNum)
            
                For i = 1 To nGetNum
                
                    atBuyTicketInfo(i).nTicketType = Trim(rsTemp!ticket_type)
                    atBuyTicketInfo(i).szReserved = ""
                    atBuyTicketInfo(i).szSeatNo = Trim(rsTemp!seat_no)

                    Dim szTmp As String
                    szTmp = TicketID + i
                    atBuyTicketInfo(i).szTicketNo = Left(TicketID, Len(TicketID) - Len(szTmp)) & TicketID + i - 1
                    atBuyTicketInfo(i).szSeatTypeID = "01" '座位类型固定为01
                    asgTicketPrices(i) = 0
                    rsTemp.MoveNext
                Next i
                atSellTKParam(1).BuyTicketInfo = atBuyTicketInfo
                atSellTKParam(1).pasgSellTicketPrice = asgTicketPrices
                anInsurrance(1) = 0
                atSellTicketResultInfo = SelfSellLocalInterNetTicket(adyDate_new, aszBusID_new, SellStationIDRX, aszEndStationID_new, atSellTKParam, anInsurrance, aszUserID, aszBankTime, oDb)
                CancelPreNetTK = atSellTicketResultInfo
            Dim ssql As String
            rsTemp.MoveFirst
            For i = 1 To nGetNum '把票号返写回网上售票表.
                ssql = "UPDATE internet_sell_ticket_info SET if_print=1  , print_user_id ='" & m_oActiveUser.UserID & "', ticket_id='" & atBuyTicketInfo(i).szTicketNo & "',print_time='" & oParam.NowDateTime & "' WHERE bus_id='" & Trim(rsTemp!bus_id) & "'" _
                    & " AND bank_card_id='" & Trim(rsTemp!bank_card_id) & "' AND bus_date='" & Trim(rsTemp!bus_date) & "'  AND seat_no='" & Trim(rsTemp!seat_no) & "' AND getticket_id='" & Trim(rsTemp!getticket_id) & "'" _
                    & " AND validate_id='" & Trim(rsTemp!validate_id) & "' AND if_print=" & cszNOPrint
                oDb.Execute ssql
                rsTemp.MoveNext
            Next i
    oDb.CommitTrans
 
    
    On Error GoTo 0
    For i = 1 To nLen
    WriteOperateLog m_oActiveUser, RIGHT_CancelTicket, "废票票号：" & paszTicketID(i)
    Next i
'    CancelPreNetTK = atResult

    Exit Function
here:
    oDb.RollbackTrans
    err.Raise err.Number
End Function

Public Function RePrintNetTKTrans(paszTicketID() As String, pszGetTicketID As String, pszValiDate As String, TicketID As String) As TSellTicketResult()
    '站方重打网上售票
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim i As Integer
    Dim j As Integer
    Dim nLen As Integer
    Dim szRs As String
    Dim rs As Recordset
    Dim atResult() As TSellTicketResult
    Dim rsTemp As Recordset
    Dim nCount As Integer
    Dim iLen As Integer
    Dim rs2 As Recordset
    Dim szTemp As String
    Dim oParam As New SystemParam
    oParam.Init m_oActiveUser
    nLen = 0
    nLen = ArrayLength(paszTicketID)
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_CancelTicket

    szTemp = ""
    ReDim atResult(1 To nLen)
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    oDb.BeginTrans
'****************************************改网上售票表状态,3为废网上票
        szRs = "SELECT * FROM internet_sell_ticket_info WHERE getticket_id='" & pszGetTicketID & "' AND validate_id='" & pszValiDate & "'"
        Set rs = oDb.Execute(szRs)

        szRs = "SELECT * FROM internet_sell_ticket_info WHERE getticket_id='" & pszGetTicketID & "' AND validate_id='" & pszValiDate & "'" _
                & " AND if_print=2 "
        Set rs2 = oDb.Execute(szRs)
        
        If rs.RecordCount = rs2.RecordCount Then ShowError ERR_NetBookCancel  '座位已退
        For i = 1 To ArrayLength(paszTicketID)
            szTemp = szTemp & "," & TransFieldValueToString(paszTicketID(i))
        Next i
        szTemp = " ticket_id in (" & Mid(szTemp, 2, Len(szTemp) - 1) & ")"
        szSql = "SELECT * FROM internet_sell_ticket_info WHERE getticket_id='" & pszGetTicketID & "' AND validate_id='" & pszValiDate & "'" _
                & " AND if_print=1  AND " & szTemp
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount = 0 Then ShowError ERR_NoGetNet
        For j = 1 To rsTemp.RecordCount
            szSql = " UPDATE internet_sell_ticket_info set if_print=3 WHERE bus_id='" & Trim(rsTemp!bus_id) & "' AND bus_date='" & ToDBDate(Trim(rsTemp!bus_date)) & "'" _
                & " AND if_print=1  AND getticket_id='" & pszGetTicketID & "' AND validate_id='" & pszValiDate & "'  AND seat_no='" & Trim(rsTemp!seat_no) & "'"
            oDb.Execute szSql
            rsTemp.MoveNext
            
        Next j
'*********************************************
        On Error GoTo here
        SelfCancelTicket paszTicketID, oDb, TP_NormalCancel '真正的废票
        ReInsertNetTK rsTemp, oDb '插入新的一条记录
        
    Dim nGetNum As Integer
    Dim szSqlSet As String
    Dim szSeatNoTemplog As String
    Dim adyDate_new() As Date
    Dim aszBusID_new()  As String
    Dim aszEndStationID_new() As String
    Dim atSellTKParam() As TSellTicketParam
    Dim aszSeatNO() As String
    Dim atBuyTicketInfo() As TBuyTicketInfo
    Dim atSellTicketResultInfo() As TSellTicketResult
    Dim SellStationIDRX(1 To 1) As String
    ReDim atBuyTicketInfo(1 To nLen)
    ReDim aszBusID_new(1 To 1)
    ReDim adyDate_new(1 To 1)
    ReDim aszEndStationID_new(1 To 1)
    ReDim atSellTKParam(1 To 1)
    Dim anInsurrance(1 To 1) As Integer
    Dim aszSellStationID() As String
    Dim aszUserID(1 To 1) As String
    Dim aszBankTime(1 To 1) As Date
                nGetNum = rsTemp.RecordCount


                
                rsTemp.MoveFirst
                
                SellStationIDRX(1) = Trim(rsTemp!sell_station_id)
        
                aszBusID_new(1) = Trim(rsTemp!bus_id)
                adyDate_new(1) = Trim(rsTemp!bus_date)
                aszEndStationID_new(1) = Trim(rsTemp!des_station_id)
                aszUserID(1) = Trim(rsTemp!bank_id)
                aszBankTime(1) = FormatDbValue(rsTemp!busy_date)
                ReDim aszResultSeatID(1 To nGetNum)
                Dim asgTicketPrices() As Single
                ReDim asgTicketPrices(1 To nGetNum)
            
                For i = 1 To nGetNum
                
                    atBuyTicketInfo(i).nTicketType = Trim(rsTemp!ticket_type)
                    atBuyTicketInfo(i).szReserved = ""
                    atBuyTicketInfo(i).szSeatNo = Trim(rsTemp!seat_no)

                    Dim szTmp As String
                    szTmp = TicketID + i
                    atBuyTicketInfo(i).szTicketNo = Left(TicketID, Len(TicketID) - Len(szTmp)) & TicketID + i - 1
                    atBuyTicketInfo(i).szSeatTypeID = "01" '座位类型固定为01
                    asgTicketPrices(i) = 0
                    rsTemp.MoveNext
                Next i
                atSellTKParam(1).BuyTicketInfo = atBuyTicketInfo
                atSellTKParam(1).pasgSellTicketPrice = asgTicketPrices
                anInsurrance(1) = 0
                atSellTicketResultInfo = SelfSellLocalInterNetTicket(adyDate_new, aszBusID_new, SellStationIDRX, aszEndStationID_new, atSellTKParam, anInsurrance, aszUserID, aszBankTime, oDb)
                RePrintNetTKTrans = atSellTicketResultInfo
            Dim ssql As String
            rsTemp.MoveFirst
            For i = 1 To nGetNum '把票号返写回网上售票表.
                ssql = "UPDATE internet_sell_ticket_info SET if_print=1  , print_user_id ='" & m_oActiveUser.UserID & "', ticket_id='" & atBuyTicketInfo(i).szTicketNo & "',print_time='" & oParam.NowDateTime & "' WHERE bus_id='" & Trim(rsTemp!bus_id) & "'" _
                    & " AND bank_card_id='" & Trim(rsTemp!bank_card_id) & "' AND bus_date='" & Trim(rsTemp!bus_date) & "'  AND seat_no='" & Trim(rsTemp!seat_no) & "' AND getticket_id='" & Trim(rsTemp!getticket_id) & "'" _
                    & " AND validate_id='" & Trim(rsTemp!validate_id) & "' AND if_print=" & cszNOPrint
                oDb.Execute ssql
                rsTemp.MoveNext
            Next i
    oDb.CommitTrans
    
    On Error GoTo 0
    For i = 1 To nLen
    WriteOperateLog m_oActiveUser, RIGHT_CancelTicket, "废票票号：" & paszTicketID(i)
    Next i


    Exit Function
here:
    oDb.RollbackTrans
    err.Raise err.Number
End Function

Private Sub ReInsertNetTK(rs As Recordset, oDb As RTConnection)
    Dim szSql As String
    Dim i As Integer
    rs.MoveFirst
    For i = 1 To rs.RecordCount
        szSql = "INSERT INTO internet_sell_ticket_info(" _
        & " bank_card_id," _
        & " busy_date," _
        & " bus_date," _
        & " bus_id," _
        & " seat_no," _
        & " sell_station_id," _
        & " des_station_id," _
        & " ticket_type," _
        & " bank_id," _
        & " center_scroll_id," _
        & " getticket_id," _
        & " validate_id," _
        & " card_id," _
        & " pay_count, " _
        & " if_print )  " _

        
        szSql = szSql & " VALUES('" _
        & Trim(rs!bank_card_id) & "','" _
        & Trim(rs!busy_date) & "','" _
        & Trim(rs!bus_date) & "','" _
        & Trim(rs!bus_id) & "','" _
        & Trim(rs!seat_no) & "','" _
        & Trim(rs!sell_station_id) & "','" _
        & Trim(rs!des_station_id) & "','" _
        & Trim(rs!ticket_type) & "','" _
        & Trim(rs!bank_id) & "','" _
        & Trim(rs!center_scroll_id) & "','" _
        & Trim(rs!getticket_id) & "','" _
        & Trim(rs!validate_id) & "','" _
        & Trim(rs!card_id) & "','" _
        & Trim(rs!pay_count) & "','" _
        & cszNOPrint & "' ) "
    oDb.Execute (szSql)
        rs.MoveNext
    Next i
End Sub


'验证网上售票
Public Function NetTKTrans(pszGetTicketID As String, pszValiDate As String, paszTicketID() As String) As Recordset

    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim szTemp As String
On Error GoTo Error_Handle
    oDb.ConnectionString = GetConnectionStr("")
    For i = 1 To ArrayLength(paszTicketID)
        szTemp = szTemp & "," & TransFieldValueToString(paszTicketID(i))
    Next i
    szTemp = " ticket_id in (" & Mid(szTemp, 2, Len(szTemp) - 1) & ")"
    szSql = "SELECT * FROM internet_sell_ticket_info WHERE (if_print=1 or if_print=4)"
            
    If Trim(pszGetTicketID) <> "" Then
        szSql = szSql & " AND getticket_id='" & pszGetTicketID & "' "
    End If
    
    If Trim(pszValiDate) <> "" Then
        szSql = szSql & " AND validate_id='" & pszValiDate & "'"
    End If
    szSql = szSql & " AND " & szTemp

    Set rsTemp = oDb.Execute(szSql)
    Set NetTKTrans = rsTemp
    Exit Function
Error_Handle:
    err.Raise err.Number
End Function

'更新银行卡号(嘉兴用)

Public Function UpdBankCardID(pszGetTicketID As String, pszValiDate As String, pszBankCardID As String) As String

    Dim szSql As String
    Dim oDb As New RTConnection
    Dim i As Integer
    Dim szTemp As String
    Dim nAffect As Long
On Error GoTo Error_Handle
    oDb.ConnectionString = GetConnectionStr("")

    szSql = "UPDATE internet_sell_ticket_info SET bank_card_id='" & Trim(pszBankCardID) & "' WHERE " _
             & " getticket_id='" & pszGetTicketID & "' " _
             & " AND validate_id='" & pszValiDate & "'"

    oDb.Execute szSql, nAffect
    If nAffect > 0 Then
        UpdBankCardID = "0000"
    Else
        UpdBankCardID = "1111"
    End If
    Exit Function
Error_Handle:
    err.Raise err.Number
End Function

Public Function ResolveTicketNo(pszFullTicketNo, ByRef pszTicketPrefix As String) As Long
    Dim i As Integer, j As Integer
    Dim nCount As Integer, nTemp As Integer, nTicketPrefixLen As Integer
    'On Error Resume Next
    pszFullTicketNo = Trim(pszFullTicketNo)
    nCount = Len(pszFullTicketNo)
    m_lTicketNoNumLen = m_oParam.TicketNumberLen
    For i = 1 To nCount
        nTemp = Asc(Mid(pszFullTicketNo, nCount - i + 1, 1))
        If nTemp < vbKey0 Or nTemp > vbKey9 Then
            Exit For
        End If
    Next
    i = i - 1
    If i > 0 Then
        nTemp = m_lTicketNoNumLen
        nTemp = IIf(nTemp > i, i, nTemp)
        ResolveTicketNo = CLng(Right(pszFullTicketNo, nTemp))
        
        nTicketPrefixLen = m_oParam.TicketPrefixLen
        If nTicketPrefixLen <= Len(pszFullTicketNo) Then
            pszTicketPrefix = Left(pszFullTicketNo, nTicketPrefixLen)
        Else
            pszTicketPrefix = pszFullTicketNo
        End If
        
    Else
        pszTicketPrefix = ""
        ResolveTicketNo = 0
    End If
    
    
    
End Function



'内部用废票
Private Sub SelfCancelTicket(paszTicketID() As String, poDb As RTConnection, pnCancelType As ECancelTicketType)
    Dim szSql As String
    Dim oParam As New SystemParam
    Dim szUserID As String
    Dim nLen As Integer
    Dim i As Integer
    Dim rsTemp1 As Recordset
    nLen = 0
    nLen = ArrayLength(paszTicketID)
    On Error GoTo here
    For i = 1 To nLen
        szSql = "SELECT s.*, u.sell_station_id seller_sell_station_id FROM Ticket_sell_lst s , user_info u where s.user_id = u.user_id AND ticket_id='" & paszTicketID(i) & "'"
        Set rsTemp1 = poDb.Execute(szSql)
        If rsTemp1.RecordCount <> 1 Then ShowError ERR_TicketNotExist, , "票号为：" & paszTicketID(i)
        
        If rsTemp1!Status = ST_TicketSellChange Then ShowError ERR_TicketIsChangeTicket, , "票号为：" & paszTicketID(i)
        
        szSql = "UPDATE Ticket_sell_lst SET " _
        & " status=status | " & ST_TicketCanceled & " WHERE " _
        & " ticket_id='" & paszTicketID(i) & "'"
        poDb.Execute szSql
        
        szSql = "UPDATE Env_bus_allot_lst SET sell_quantity=sell_quantity-1 WHERE bus_id='" & rsTemp1!bus_id & "'" _
              & " AND sell_station_id='" & rsTemp1!sell_station_id & "' AND bus_date='" & rsTemp1!bus_date & "'"
        poDb.Execute szSql
        
        If m_bSellStationCanSellEachOther Then
            szSql = "UPDATE Env_bus_sell_station_lst SET sell_quantity=sell_quantity-1 WHERE bus_id='" & rsTemp1!bus_id & "'" _
                  & " AND sell_station_id='" & rsTemp1!seller_sell_station_id & "' AND bus_date='" & rsTemp1!bus_date & "'"
            poDb.Execute szSql
        End If
        
        oParam.Init m_oActiveUser
        If pnCancelType = TP_ForceCancel Then
            Dim rsTemp As Recordset
            szSql = "SELECT user_id FROM Ticket_sell_lst WHERE ticket_id='" & paszTicketID(i) & "'"
            Set rsTemp = poDb.Execute(szSql)
           ' szUserID = FormatDbValue(rsTemp!user_id)
        Else
           
        End If
        
        szUserID = m_oActiveUser.UserID
        
        szSql = "INSERT INTO ticket_cancel_lst( " _
        & " ticket_id," _
        & " cancel_ticket_time," _
        & " user_id," _
        & " cancel_mode ) VALUES('" _
        & paszTicketID(i) & "','" _
        & ToDBDateTime(oParam.NowDateTime) & "','" _
        & szUserID & "'," _
        & pnCancelType & ")"
        
        poDb.Execute szSql
        

        '回收座位
        RecycleSeat paszTicketID(i), poDb
    Next i
    On Error GoTo 0
    Exit Sub
here:
    err.Raise err.Number
End Sub


