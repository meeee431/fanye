VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SellTicketService"
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

'此处定义错误号：
Public Enum EErrSellTicketService
    ERR_SellTicketServiceNoActiveUser = ERR_SellTicketService + ERR_NoActiveUser
    
    ERR_BusStopped = ERR_SellTicketService + 21 '车次已经停班
    ERR_BusMerged = ERR_SellTicketService + 22 '车次已经被并班停班
    ERR_BusStopCheck = ERR_SellTicketService + 23 '车次已经停检
    ERR_BusStopSellTicket = ERR_SellTicketService + 24 '车次已经停售
    ERR_BusStaionStopSellTicket = ERR_SellTicketService + 25 '车次还未到指定到站的售票时间
    ERR_BusReachMaxStationTicketCount = ERR_SellTicketService + 26 '超过指定到站的限售票数
    ERR_BusReachMaxTicketCount = ERR_SellTicketService + 27 '指定车次的座位数已经不够
    ERR_SeatNoCanNotSell = ERR_SellTicketService + 28 '指定的座位号不能出售（已被售出或预留）
    ERR_CanNotSellStandTicketWhenHaveSeat = ERR_SellTicketService + 29 '还有座位时不能售站票
    ERR_NoStandTicket = ERR_SellTicketService + 30 '已无站票
    ERR_NoThisBus = ERR_SellTicketService + 31 '指定的车次不存在或指定的车次不过指定的车站
    ERR_BusNotInIlegalDate = ERR_SellTicketService + 32 '指定的车次的日期不在法范围内(当天到预售天之内)
    ERR_NoTicketInfoSpecified = ERR_SellTicketService + 33 '没有指定何购票信息

    ERR_TicketIsCanceled = ERR_SellTicketService + 34 '此票已经被废
    ERR_TicketIsReturned = ERR_SellTicketService + 35 '此票已经被退
    ERR_TicketIsChecked = ERR_SellTicketService + 36 '此票已经被检
    ERR_TicketIsChangeTicket = ERR_SellTicketService + 37 '此票是改签售出的票,不允许作废,只能作退票处理
    
    ERR_UserIsNotTicketSeller = ERR_SellTicketService + 38 '操作员不是指定票的售票员,而正常废票需由本人进行
    ERR_CancelTicketTimeOut = ERR_SellTicketService + 39 '已过正常废票时间
    ERR_TicketNotExist = ERR_SellTicketService + 40 '指定的票不存在
    ERR_TicketIsChanged = ERR_SellTicketService + 41 '此票已被改签
    ERR_ReturnTicketTimeOut = ERR_SellTicketService + 42 '已过正常退票时间
    ERR_FreeTicketCanNotReturnChange = ERR_SellTicketService + 43 '免票不能退,改签
    ERR_OrgReturnChargeOutDated = ERR_SellTicketService + 44 '原先得到的退票费已过时
    ERR_InputParamTypeErr = ERR_SellTicketService + 45 '输入的参数类型不合要求
    ERR_ScrollBusCanNotExtraSell = ERR_SellTicketService + 46 '流水车次不能补票
    ERR_TicketNoDup = ERR_SellTicketService + 47 '票号重
    ERR_ScrollBusTicketCanNotReturn = ERR_SellTicketService + 48 '流水车次不能正常退票
    ERR_ScroolBusReturnTicketDateOut = ERR_SellTicketService + 49 '流水车次的退票只能在当天进行
    ERR_WriteOffTiecketFail = ERR_SellTicketService + 50 '注销票失败
    ERR_NoExistWriteOffTiecket = ERR_SellTicketService + 51 '无此注销票
    ERR_ForceCancelTicketRemoteNotAllowed = ERR_SellTicketService + 52 '强行废票不允许远程
    ERR_BusNotStopSellTicket = ERR_SellTicketService + 53 '还未到补票时间,不能补票
    ERR_TicketTypeNotDefined = ERR_SellTicketService + 54
    
    ERR_TicketPriceOutOfLimit = ERR_SellTicketService + 55 '所售票的票价超出折扣票允许的范围
    
    ERR_BusSlitpStop = ERR_SellTicketService + 56 '拆分停班车次不能售票
    
    ERR_ReturnTicketNotResume = ERR_SellTicketService + 57 '车票已售完不能取消退票
    ERR_CancelTicketNotResume = ERR_SellTicketService + 58 '车票已售完不能取消废票
    ERR_BusIsSplit = ERR_SellTicketService + 59 '车次正在拆分不能售票
    
    '定义增加错误号留接口
    ERR_SellQuantityISFull = ERR_SellTicketService + 60 '12460 售票站超过分配的售票张数
    ERR_ErrOther2 = ERR_SellTicketService + 61
    ERR_ErrOther3 = ERR_SellTicketService + 63
    
    
    ERR_UserCannotSell = ERR_SellTicketService + 64 '12464 用户不被允许售当天的票
    ERR_BusAllotInfoNotExist = ERR_SellTicketService + 65 '12465 售的上车站在该车次的配载信息中不存在
    
    
    ERR_BusSellStationInfoNotExist = ERR_SellTicketService + 66 '12466 售票员所属车站在该车次的售票点信息中不存在
    ERR_CannotSellZeroPrice = ERR_SellTicketService + 67 '12467 票价为0的票不允许出售
    
    ERR_CanNotSellSpecialTicket = ERR_SellTicketService + 68 '12468 特种票已售完
    
    ERR_SeatRepeat = ERR_SellTicketService + 69 '座位重号
    
    ERR_ChangeTicketTimeOut = ERR_SellTicketService + 70 '已过正常改签时间
    
    ERR_CanNotSellHalfTicket = ERR_SellTicketService + 71 '半票已售完
    
    ERR_TicketNoDraw = ERR_SellTicketService + 72 '当前票号没有领用记录或未启用
    
    ERR_SeatHasBooked = ERR_SellTicketService + 73 '座位已被预定，不能出售
    
    ERR_IDCardNoCanNotNull = ERR_SellTicketService + 74 '12474 实名制验证失败，证件号不能为空
    ERR_IDCardNoAndPersonNameCanNotNullOne = ERR_SellTicketService + 75 '12475 实名制验证失败，证件号、姓名不允许同时为空
    ERR_IDCardNoAndPersonNameCanNotNullAll = ERR_SellTicketService + 76 '12476 实名制验证失败，证件号、姓名两者都不能为空
    
End Enum

'此处定义权限号：
Public Enum ERightSellTicketService
    RIGHT_SellTicket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 1 '售票
    RIGHT_ExtraSellTicket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 2 '补票
    RIGHT_CancelTicket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 3 '废票
    RIGHT_ForceCancelTicket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 4 '强行废票
    RIGHT_ReturnTicket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 5 '退票
    RIGHT_ForceReturnTicket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 6 '强行退票
    RIGHT_ChangeTicket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 7 '改签车票
    RIGHT_QuerySelfThings = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 8 '查询自己的票处理明细
    RIGHT_QueryOtherThings = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 9 '查询他人的票处理明细
    RIGHT_WriteOffTiecket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 10 '注销票权限
    
    RIGHT_SellDiscountTicket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 11 '售折扣票
    RIGHT_SellOrderSeat = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 12     '售定座位号
    RIGHT_ResumeReturnTicket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 13 '取消退票
    RIGHT_ResumeCancelTicket = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 14 '取消废票
    RIGHT_SelectReturnRatio = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 15 '选择退票废率
     
    定义增加权限接口
    RIGHT_RightOther1 = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 16
    RIGHT_RightOther2 = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 17  '售计划预定票
    RIGHT_RightOther3 = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 18
    RIGHT_AllMoneyReturn = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 19 '全额退票
    RIGHT_ForceSell = ERR_SellTicketService + cnMidRightBegin + cnMidRightStep * 20 '强行出售
End Enum

'所有的类都有的私有类变量
Private m_oActiveUser As ActiveUser
Private m_oParam As New SystemParam
Private m_bSellStationCanSellEachOther As Boolean
Private m_bAllowScrollBusSaleForever As Boolean

'--------------------------->>
Public Sub Init(poAUser As ActiveUser)
    Set SelfUser = poAUser
    m_oParam.Init SelfUser
'    m_bSellStationCanSellEachOther = m_oParam.SellStationCanSellEachOther
    If m_oActiveUser.SellStationID = "km" Then
        m_bSellStationCanSellEachOther = True
    Else
        m_bSellStationCanSellEachOther = False
    End If
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

'售票
Public Function SellTicket(BusStartupDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String, aspSellTicketParam() As TSellTicketParam, anInsurance() As Integer, Optional aszRealNameInfo As Variant) As TSellTicketResult()
    Dim oDb As New RTConnection
    Dim tsTemp() As TSellTicketResult
    Dim iLen As Integer
    Dim nCount As Integer
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_SellTicket
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    iLen = ArrayLength(BusID)
    oDb.BeginTrans
        On Error GoTo Here
        ReDim tsTemp(1 To iLen)
        
        Dim nStart As Integer
        For nCount = 1 To iLen
            tsTemp(nCount) = SelfSellTicket(BusStartupDate(nCount), BusID(nCount), SellStationID(nCount), DesStationID(nCount), aspSellTicketParam(nCount).BuyTicketInfo, oDb, aspSellTicketParam(nCount).pasgSellTicketPrice, anInsurance(nCount), False, ST_TicketNormal, nStart, aszRealNameInfo)
            nStart = nStart + ArrayLength(aspSellTicketParam(nCount).BuyTicketInfo)
        Next nCount
        SellTicket = tsTemp
    oDb.CommitTrans
    
'    On Error GoTo 0
'    WriteOperateLog m_oActiveUser, RIGHT_SellTicket
    
    Exit Function
    Set oDb = Nothing
Here:
    oDb.RollbackTrans
    Set oDb = Nothing
    err.Raise err.Number
End Function

'补票
Public Function ExtraSellTicket(BusStartupDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String, aspSellTicketParam() As TSellTicketParam, Optional aszRealNameInfo As Variant) As TSellTicketResult()
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
    On Error GoTo Here

    
    Dim nStart As Integer
    For nCount = 1 To iLen
        rsTemp(nCount) = SelfSellTicket(BusStartupDate(nCount), BusID(nCount), SellStationID(nCount), DesStationID(nCount), aspSellTicketParam(nCount).BuyTicketInfo, oDb, aspSellTicketParam(nCount).pasgSellTicketPrice, 0, True, ST_TicketNormal, nStart, aszRealNameInfo)
        nStart = nStart + ArrayLength(aspSellTicketParam(nCount).BuyTicketInfo)
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
Here:
    oDb.RollbackTrans
    err.Raise err.Number
End Function



Private Function UpdateSeatStatusStr(pszBusID As String, pszDate As Date, pszSeatNo As String, pszTicketID As String) As String
    Dim szsql As String
    
    szsql = "UPDATE Env_bus_seat_lst SET status=" & ST_SeatSold & "," _
    & " ticket_no='" & pszTicketID & "' " _
    & " WHERE bus_id='" & pszBusID & "' AND " _
    & " bus_date='" & ToDBDate(pszDate) & "' AND " _
    & " seat_no='" & pszSeatNo & "' AND " _
    & " status=" & ST_SeatCanSell
    UpdateSeatStatusStr = szsql
End Function

'内部售票函数  票的折扣已删除 2002/11/13 wjb
Private Function SelfSellTicket(BusStartupDate As Date, BusID As String, SellStationID As String, DesStationID As String, _
                                SellTicketInfo() As TBuyTicketInfo, poDb As RTConnection, _
                                pasgSellTicketPrice() As Single, pnInsurance As Integer, _
                                Optional pbExtraSell As Boolean = False, _
                                Optional pnTicketStatus As ETicketStatus = ST_TicketNormal, Optional nRealNameStart As Integer, Optional aszRealNameInfo As Variant) As TSellTicketResult
    
    'psgDiscount票的折扣
    
    Dim oParam As New SystemParam
    Dim szsql As String
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
    
    Dim bIsRealName As Boolean  '是否需要实名制
    Dim rsSellCardInfo As Recordset
    Dim nCardBuyTicketInterval As Integer '同一证件购票时间间隔（单位分钟）
    Dim dtBusStartTime As Date
    Dim szWhereCard As String
    Dim szCardTemp As String
    
    Dim m_bSpecialTicketTypePosition As String
    Dim m_bSpecialTicketTypeRatio As Double
    Dim m_bHalfTicketTypeRatio As Double
    m_bSpecialTicketTypePosition = m_oParam.SpecialTicketTypePosition
    m_bSpecialTicketTypeRatio = m_oParam.SpecialTicketTypeRatio
    m_bHalfTicketTypeRatio = m_oParam.HalfTicketTypeRatio
    
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
    
    nCardBuyTicketInterval = Val(oParam.CardBuyTicketInterval)
    
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
    
     szsql = "SELECT e.bus_date , e.bus_id , e.vehicle_type_name , e.vehicle_type_code , e.total_seat " _
        & " , e.sale_seat_quantity , e.total_stand_quantity , e.sale_stand_seat_quantity , e.bus_start_time , e.bus_type" _
        & " , e.register_status , e.is_all_refundment , e.route_id , e.end_station_id , e.end_station_name" _
        & " , e.owner_id , e.transport_company_id , e.vehicle_id , e.split_company_id , e.baggage_number , e.fact_weight" _
        & " , e.calculate_weight , e.over_weight_number , e.luggage_total_price , e.internet_status , e.scrollbus_check_time" _
        & " , e.seat_remain , e.bed_remain  , e.additional_remain , e.other_remain_1 , e.other_remain_2" _
        & " , s.sell_station_id , s.station_id , s.seat_type_id , s.station_serial_no , s.max_sale_quantity , s.sale_ticket_quantity" _
        & " , s.book_count , s.stop_sale_time , s.mileage  , s.full_price , s.half_price , s.preferential_ticket1 , s.preferential_ticket2" _
        & " , s.preferential_ticket3" _
        & " , a.check_gate_id , a.bus_start_time" _
        & " , a.status , a.can_sale_quantity , a.sell_quantity ,r.id_card" _
        & " FROM Work_env_bus_info e ,Work_env_bus_station_lst s , Work_env_bus_allot_lst a ,route_info r" _
        & " WHERE e.bus_id='" & BusID & "' AND  e.bus_date='" & ToDBDate(BusStartupDate) & "'" _
        & " AND  e.bus_id=s.bus_id AND  e.bus_date=s.bus_date AND e.route_id =r.route_id" _
        & " AND  s.station_id='" & DesStationID & "' AND s.sell_station_id='" & SellStationID & "'" _
        & " AND a.bus_id = s.bus_id AND a.bus_date = s.bus_date AND a.sell_station_id = '" & SellStationID & "'" _

    
    Set rsTemp = poDb.Execute(szsql)
    
    
    If rsTemp.RecordCount = 0 Then ShowError ERR_NoThisBus '没有指定的车次或车次不过指定的到站
    
    bIsRealName = FormatDbValue(rsTemp!id_card)
    dtBusStartTime = FormatDbValue(rsTemp!bus_start_time)
    
'******************************************************************************
'判断特种票是否售完fpd
    If nBusType <> TP_ScrollBus Then
    
        Dim nSpecialTicketTypeCount As Integer
        Dim rsSelledSpecialTicketCount As Recordset
        nSpecialTicketTypeCount = 0
        
        For i = 1 To ArrayLength(SellTicketInfo)
            If SellTicketInfo(i).nTicketType = m_bSpecialTicketTypePosition Then
                nSpecialTicketTypeCount = nSpecialTicketTypeCount + 1
            End If
        Next i
        
        If nSpecialTicketTypeCount > 0 Then
            szsql = "SELECT count(*) as selled_special_ticket_count From ticket_sell_lst " _
              & " WHERE bus_id='" & BusID & "' AND  bus_date='" & ToDBDate(BusStartupDate) & "'" _
              & " AND  status in ('" & ST_TicketNormal & "','" & ST_TicketNormal + ST_TicketChecked & "','" & ST_TicketSellChange & "','" & ST_TicketSellChange + ST_TicketChecked & "' ) AND ticket_type='" & m_bSpecialTicketTypePosition & "' "
            Set rsSelledSpecialTicketCount = poDb.Execute(szsql)
            
            If (FormatDbValue(rsSelledSpecialTicketCount!selled_special_ticket_count) + nSpecialTicketTypeCount) > Int(m_bSpecialTicketTypeRatio * FormatDbValue(rsTemp!total_seat) / 100) Then
                ShowError ERR_CanNotSellSpecialTicket
            End If
        End If
        
    End If
'******************************************************************************

'******************************************************************************
'判断半票是否售完fpd
    If nBusType <> TP_ScrollBus Then
    
        Dim nHalfTicketTypeCount As Integer
        Dim rsSelledHalfTicketCount As Recordset
        nHalfTicketTypeCount = 0
        
        For i = 1 To ArrayLength(SellTicketInfo)
            If SellTicketInfo(i).nTicketType = TP_HalfPrice Then
                nHalfTicketTypeCount = nHalfTicketTypeCount + 1
            End If
        Next i
        
        If nHalfTicketTypeCount > 0 Then
            szsql = "SELECT count(*) as selled_half_ticket_count From ticket_sell_lst " _
              & " WHERE bus_id='" & BusID & "' AND  bus_date='" & ToDBDate(BusStartupDate) & "'" _
              & " AND  status in ('" & ST_TicketNormal & "','" & ST_TicketNormal + ST_TicketChecked & "','" & ST_TicketSellChange & "','" & ST_TicketSellChange + ST_TicketChecked & "' ) AND ticket_type='" & TP_HalfPrice & "' "
            Set rsSelledHalfTicketCount = poDb.Execute(szsql)
            
            If (FormatDbValue(rsSelledHalfTicketCount!selled_half_ticket_count) + nHalfTicketTypeCount) > Int(m_bHalfTicketTypeRatio * FormatDbValue(rsTemp!total_seat) / 100) Then
                ShowError ERR_CanNotSellHalfTicket
            End If
        End If
        
    End If
'******************************************************************************

    nBusType = FormatDbValue(rsTemp!bus_type)
    
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
    
    '玉环限制每个售票窗口的停售时间，fpd
    Dim oReg As New CFreeReg
    Dim nStopSellTime As Integer
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    nStopSellTime = oReg.GetSetting("SellTk", "StopSellTime", oParam.StopSellTime)
    
    '判断是否到了停售时间(如果是补票则要过了正常售票时间才能补)
    If nBusType <> TP_ScrollBus Then '如果是固定车次才限制
        If pbExtraSell Then
            If DateDiff("s", dtNow, dtBusFullDateTime) >= nStopSellTime * 60 Then ShowError ERR_BusNotStopSellTicket
        Else
            If DateDiff("s", dtNow, dtBusFullDateTime) < nStopSellTime * 60 Then ShowError ERR_BusStopSellTicket
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
                            szsql = "SELECT s.seat_no AS SeatNo,s.status ,s.seat_type_id AS SeatType,p.ticket_price_total AS TicketPrice  FROM Env_bus_seat_lst s ,Work_env_bus_price_lst p" _
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
                            szsql = szsql & "s.seat_type_id=p.seat_type_id AND " _
                                        & "p.station_id='" & DesStationID & "' AND " _
                                        & "p.ticket_type='" & SellTicketInfo(i).nTicketType & "'"
                            
                        Else
                            szsql = "SELECT s.seat_no AS SeatNo,s.seat_type_id AS SeatType,s.status,p.ticket_price_total AS TicketPrice  FROM Env_bus_seat_lst s, Work_env_bus_price_lst p WHERE " _
                            & " s.bus_id='" & BusID & "' AND " _
                            & " s.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & " p.sell_station_id='" & SellStationID & "' AND " _
                            & "s.seat_no=(SELECT MIN(seat_no) FROM Env_bus_seat_lst WHERE bus_id='" & BusID & "' AND " _
                            & "bus_date='" & ToDBDate(BusStartupDate) & "' AND status=" & ST_SeatCanSell & " AND seat_type_id='" & SellTicketInfo(i).szSeatTypeID & "') AND " _
                            & "s.bus_id=p.bus_id AND " _
                            & "s.bus_date=p.bus_date AND "
                            szsql = szsql & "s.seat_type_id=p.seat_type_id AND " _
                                        & "p.station_id='" & DesStationID & "' AND " _
                                        & "p.ticket_type='" & SellTicketInfo(i).nTicketType & "'"

                        End If
                        
                        Set rsTemp = poDb.Execute(szsql)
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
                            szsql = UpdateSeatStatusStr2(BusID, BusStartupDate, szSeatNo, SellTicketInfo(i).szTicketNo, bBookValid)
                        Else
                            szsql = UpdateSeatStatusStr(BusID, BusStartupDate, szSeatNo, SellTicketInfo(i).szTicketNo)
                        End If
                        If rsTemp!Status = ST_SeatProjectBooked Then szsql = UpdateSeatStatusStr3(BusID, BusStartupDate, szSeatNo, SellTicketInfo(i).szTicketNo)
                        poDb.Execute szsql, lRowAffect
                        
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
                        szsql = UpdateSeatStatusStr(BusID, BusStartupDate, SellTicketInfo(i).szSeatNo, SellTicketInfo(i).szTicketNo)
                        poDb.Execute szsql, lRowAffect
                        If lRowAffect = 0 Then
                            If bNeedBook Then
                                If nSeatStatus = ST_SeatProjectBooked Then
                                    AssertHaveRight m_oActiveUser, RIGHT_RightOther1  '判断计划预定权限
                                    szsql = UpdateSeatStatusStr3(BusID, BusStartupDate, SellTicketInfo(i).szSeatNo, SellTicketInfo(i).szTicketNo)
                                    poDb.Execute szsql, lRowAffect
                                    If lRowAffect = 0 Then ShowError ERR_SeatNoCanNotSell, "SellTicketService.SelfSellTicket"
                                    bSeatBooked = False
                                Else
                                    szsql = UpdateSeatStatusStr2(BusID, BusStartupDate, SellTicketInfo(i).szSeatNo, SellTicketInfo(i).szTicketNo, bBookValid, SellTicketInfo(i).szReserved)
                                
                                    poDb.Execute szsql, lRowAffect
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
                        szsql = "SELECT * FROM ticket_book_lst WHERE " _
                        & " bus_id='" & BusID & "' AND " _
                        & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                        & " seat_no='" & szSeatNo & "' AND " _
                        & " status=0"
                        Set rsTemp2 = poDb.Execute(szsql)
                        szBookStationID = Trim(rsTemp2!station_id)
                        Set rsTemp2 = Nothing
                        
                       
                        
                        '更新站点的预定张数  加上了上车站代码条件
                        'If nStationLimitCount > 0 Then
                            szsql = "UPDATE Env_bus_station_lst SET " _
                            & " book_count=book_count-1 WHERE " _
                            & " bus_id='" & BusID & "' AND " _
                            & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & " station_id='" & szBookStationID & "' AND " _
                            & " sell_station_id='" & SellStationID & "' AND " _
                            & " book_count>0"
                            
                            poDb.Execute szsql
                        'End If
                        
                        If bBookValid Then
                            szsql = "UPDATE ticket_book_lst SET status=2 WHERE " _
                            & " bus_id='" & BusID & "' AND " _
                            & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & " seat_no='" & szSeatNo & "' AND " _
                            & " status=0 " '0状态为预定
                        Else
                            szsql = "UPDATE ticket_book_lst SET status=3 WHERE " _
                            & " bus_id='" & BusID & "' AND " _
                            & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & " seat_no='" & szSeatNo & "' AND " _
                            & " status=0 " '0状态为预定
                        End If
                        poDb.Execute szsql
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
                    
                    
                    
'                    If bSellStationCanSellEachOther Then
'                        '得到环境车次售票点信息
'                        Set rsTemp2 = poDb.Execute("SELECT * from work_env_bus_sell_station_lst WHERE bus_date='" _
'                        & ToDBDate(BusStartupDate) & "' AND " _
'                        & " bus_id='" & BusID & "' AND " _
'                        & " sell_station_id='" & m_oActiveUser.SellStationID & "'")
'                        If rsTemp2.RecordCount = 0 Then ShowError ERR_BusSellStationInfoNotExist
'
'                        '总共允许售票张数是否大于"已售张数"+1
'
'                        If rsTemp2!can_sale_quantity >= rsTemp2!sell_quantity + 1 Or rsTemp2!can_sale_quantity = -1 Then
'                            poDb.Execute ("UPDATE env_bus_sell_station_lst SET sell_quantity=sell_quantity + 1 WHERE bus_date='" _
'                                & ToDBDate(BusStartupDate) & "' AND " _
'                                & " bus_id='" & BusID & "' AND " _
'                                & " sell_station_id='" & m_oActiveUser.SellStationID & "'")
'                        Else
'                            ShowError ERR_SellQuantityISFull '超过允许售票张数
'                        End If
'                    End If

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
                                szsql = "UPDATE Env_bus_station_lst SET " _
                                & " sale_ticket_quantity=sale_ticket_quantity-1 WHERE " _
                                & " bus_id='" & BusID & "' AND " _
                                & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                                & " sale_ticket_quantity>book_count AND " _
                                & " sell_station_id='" & SellStationID & "' AND " _
                                & " station_id='" & DesStationID & "'"
                                poDb.Execute szsql, lRowAffect
                                
                                If lRowAffect = 0 Then ShowError ERR_BusReachMaxStationTicketCount, "SellTicketService.SelfSellTicket"
                            End If
                            
                        End If
                        
                        szsql = "UPDATE Env_bus_info SET sale_seat_quantity=sale_seat_quantity-1 , "
                        Select Case szSeatType
                            Case cszSeatType
                                szsql = szsql & "seat_remain=seat_remain - 1 "
                            Case cszBedType
                                szsql = szsql & "bed_remain=bed_remain - 1 "
                            Case cszAdditionalType
                                szsql = szsql & "additional_remain=additional - 1 "
                            Case cszOtherType1
                                szsql = szsql & "other_remain_1 =other_remain_1 - 1"
                            Case cszOtherType2
                                szsql = szsql & "other_remain_2=other_remain_2 - 1"
                        End Select
                        szsql = szsql & " WHERE bus_id='" & BusID & "' AND bus_date='" & ToDBDate(BusStartupDate) & "'"
                        
                        poDb.Execute szsql
                    End If
                Else
                    szsql = "UPDATE Env_bus_info SET sale_stand_seat_quantity=sale_stand_seat_quantity-1 " _
                    & " WHERE bus_id='" & BusID & "' AND " _
                    & "bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                    & " sale_stand_seat_quantity>0 "
                    
                    poDb.Execute szsql, lRowAffect
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
            
            If bIsRealName = True Then
                '如果需要实名制
                
                If aszRealNameInfo(nRealNameStart + i).szCardType = "身份证" Then
                '如果证件为身份证
                    If SellTicketInfo(i).nTicketType = TP_HalfPrice Then
                    '如果是半票
                        '证件号、姓名不允许同时为空
                        If aszRealNameInfo(nRealNameStart + i).szIDCardNo = "" And aszRealNameInfo(nRealNameStart + i).szPersonName = "" Then ShowError ERR_IDCardNoAndPersonNameCanNotNullOne
                    Else
                    '不是半票
                        '证件号、姓名两者都不能为空
                        If aszRealNameInfo(nRealNameStart + i).szIDCardNo = "" Or aszRealNameInfo(nRealNameStart + i).szPersonName = "" Then ShowError ERR_IDCardNoAndPersonNameCanNotNullAll
                    End If
                Else
                '如果证件不是身份证
                    '证件号不允许为空
                    If aszRealNameInfo(nRealNameStart + i).szIDCardNo = "" Then ShowError ERR_IDCardNoCanNotNull
                End If
                
                '实名制，判断同一证件购票间隔是否超过规定时间间隔
                If aszRealNameInfo(nRealNameStart + i).szIDCardNo <> "" Then
                    szCardTemp = aszRealNameInfo(nRealNameStart + i).szIDCardNo
                    szWhereCard = " AND s.id_card_no =" & TransFieldValueToString(szCardTemp)
                Else
                    szCardTemp = aszRealNameInfo(nRealNameStart + i).szPersonName
                    szWhereCard = " AND s.person_name =" & TransFieldValueToString(szCardTemp)
                End If
                szsql = "SELECT s.bus_id ,a.bus_start_time FROM ticket_sell_lst s ,env_bus_allot_lst a WHERE" _
                    & " s.bus_date = a.bus_date And s.bus_id = a.bus_id And s.sell_station_id = a.sell_station_id" _
                    & " AND s.status IN (" & ST_TicketNormal & "," & ST_TicketNormal & "|" & ST_TicketChecked & "," & ST_TicketSellChange & "," & ST_TicketSellChange & "|" & ST_TicketChecked & ")" _
                    & " AND s.id_card_no =" & TransFieldValueToString(aszRealNameInfo(nRealNameStart + i).szIDCardNo) _
                    & " AND s.bus_date >=" & TransFieldValueToString(DateAdd("D", -1, dtNow)) _
                    & " AND ABS(DATEDIFF(s ,'" & dtBusStartTime & "',a.bus_start_time)) <=" & nCardBuyTicketInterval * 60
                Set rsSellCardInfo = poDb.Execute(szsql)
                If rsSellCardInfo.RecordCount > 0 Then
                    err.Raise 7777, , "证件[" & szCardTemp & "]" & "未超过系统规定的时间间隔[" & nCardBuyTicketInterval & "]分钟" & vbCrLf & "已购买车次[" & Trim(FormatDbValue(rsSellCardInfo!bus_id)) & "]" & ",发车时间[" & ToDBDateTime(FormatDbValue(rsSellCardInfo!bus_start_time)) & "]"
                End If
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
            szsql = "INSERT INTO Ticket_sell_lst(" _
                & " sell_station_id," _
                & " ticket_id," _
                & " bus_date," _
                & " bus_id," _
                & " card_type," _
                & " id_card_no," _
                & " person_name," _
                & " sex," _
                & " person_picture," _
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
            
            szsql = szsql & " VALUES ('" _
                & SellStationID & "','" _
                & SellTicketInfo(i).szTicketNo & "','" _
                & ToDBDate(BusStartupDate) & "','" _
                & BusID & "','" _
                & aszRealNameInfo(nRealNameStart + i).szCardType & "','" _
                & aszRealNameInfo(nRealNameStart + i).szIDCardNo & "','" _
                & aszRealNameInfo(nRealNameStart + i).szPersonName & "','" _
                & aszRealNameInfo(nRealNameStart + i).szSex & "','" _
                & aszRealNameInfo(nRealNameStart + i).szPersonPicture & "','" _
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
            poDb.Execute szsql
            
            '写详细票价信息
            '???
'            If atiTicketPrice(SellTicketInfo(i).nTicketType).sgBase = 0 Then '此处需处理
                szsql = GetTicketPriceDetailStr(BusID, BusStartupDate, SellStationID, DesStationID, szSeatType, IIf(SellTicketInfo(i).nTicketType = TP_FreeTicket, TP_FullPrice, SellTicketInfo(i).nTicketType))
                Set rsTemp = poDb.Execute(szsql)
                LetPriceDetailFromRsToType rsTemp, atiTicketPrice(SellTicketInfo(i).nTicketType), 0
'            End If
            szsql = GetInsertPriceDetailStr(atiTicketPrice(SellTicketInfo(i).nTicketType), SellTicketInfo(i).szTicketNo)
            poDb.Execute szsql
            siSellInfo.aszSeatType(i) = szSeatType
            siSellInfo.aszTicketType(i) = SellTicketInfo(i).nTicketType
            siSellInfo.asgTicketPrice(i) = sgTicketPrice
            siSellInfo.aszSeatNO(i) = szSeatNo
        Next
        szsql = "SELECT station_name FROM station_info WHERE station_id='" & DesStationID & "'"
        Set rsTemp = poDb.Execute(szsql)
        
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
    Dim szsql As String
    Dim szTemp As String
    Dim oParam As New SystemParam
    oParam.Init m_oActiveUser
    
    szsql = "SELECT * FROM Work_env_bus_price_lst WHERE " _
    & " bus_id='" & pszBusID & "' AND " _
    & " bus_date='" & ToDBDate(pdtDate) & "' AND " _
    & " station_id='" & pszStaionID & "' AND " _
    & " sell_station_id='" & pszSellStationID & "' AND " _
    & "seat_type_id='" & pszSeatType & "' AND " _
    & " ticket_type=" & pnTicketType

    GetTicketPriceDetailStr = szsql
End Function


Private Function GetInsertPriceDetailStr(ptiPriceDetailInfo As TRETicketPrice, pszTicketID As String) As String
    Dim szsql As String
    Dim i As Integer
    
    szsql = " INSERT INTO ticket_price_lst(" _
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
        szsql = szsql & "," & ptiPriceDetailInfo.asgPrice(i)
    Next
    szsql = szsql & ")"
    
    GetInsertPriceDetailStr = szsql
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

'正常废票
Public Sub CancelTicket(paszTicketID() As String)
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim i As Integer
    Dim nLen As Integer
    nLen = 0
    nLen = ArrayLength(paszTicketID)
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_CancelTicket
    
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    oDb.BeginTrans
        For i = 1 To nLen
            AssertTicketCanCancelReturn paszTicketID(i), oDb, TP_CancelTicket, False
        Next i
    
        On Error GoTo Here
        SelfCancelTicket paszTicketID, oDb, TP_NormalCancel
    oDb.CommitTrans
    
    On Error GoTo 0
    For i = 1 To nLen
    WriteOperateLog m_oActiveUser, RIGHT_CancelTicket, "废票票号：" & paszTicketID(i)
    Next i
    Exit Sub
Here:
    oDb.RollbackTrans
    err.Raise err.Number
End Sub


'强行废票
Public Sub ForceCancelTicket(paszTicketID() As String)
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim i As Integer
    Dim nLen As Integer
    Dim iLen As Integer
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_ForceCancelTicket
    
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    nLen = ArrayLength(paszTicketID)
    For i = 1 To nLen
        AssertTicketCanCancelReturn paszTicketID(i), oDb, TP_CancelTicket, True
    Next i
    oDb.BeginTrans
    On Error GoTo Here
    SelfCancelTicket paszTicketID, oDb, TP_ForceCancel
    oDb.CommitTrans
    On Error GoTo 0
    For i = 1 To iLen
     WriteOperateLog m_oActiveUser, RIGHT_ForceCancelTicket, "强行废票票号:" & paszTicketID(i)
    Next
    Exit Sub
Here:
    oDb.RollbackTrans
    err.Raise err.Number
End Sub


'回收座位或站票的资源信息
Private Sub RecycleSeat(pszTicketID As String, poDb As RTConnection)
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim szBusID As String
    Dim dtDate As Date
    Dim szDesStationID As String
    Dim szSeatNo As String
    Dim szSeatTypeID As String
    Dim szSellStationID As String
    
    szsql = "SELECT * FROM Ticket_sell_lst WHERE ticket_id='" & pszTicketID & "'"
    Set rsTemp = poDb.Execute(szsql)
    
    szBusID = FormatDbValue(rsTemp!bus_id)
    dtDate = FormatDbValue(rsTemp!bus_date)
    szDesStationID = FormatDbValue(rsTemp!des_station_id)
    szSeatNo = FormatDbValue(rsTemp!seat_no)
    szSellStationID = FormatDbValue(rsTemp!sell_station_id)
    
    If szSeatNo = cszStandTicketSeatNo Then '是站票
        '将该环境车次的可售站票数加一
        szsql = "UPDATE Env_bus_info SET " _
        & " sale_stand_seat_quantity=sale_stand_seat_quantity+1 WHERE " _
        & " bus_id='" & szBusID & "' AND " _
        & " bus_date='" & ToDBDate(dtDate) & "'"
        poDb.Execute szsql
    Else '不是站票
        szSeatTypeID = FormatDbValue(rsTemp!seat_type_id)
        
        szSeatTypeID = FormatDbValue(rsTemp!seat_type_id)
        
        '将该环境车次的可售座位数减一和可售座位类型余留减一
        szsql = "UPDATE Env_bus_info SET sale_seat_quantity=sale_seat_quantity+1 , "
        Select Case szSeatTypeID
            Case cszSeatType
                szsql = szsql & " seat_remain=seat_remain + 1  "
            Case cszBedType
                szsql = szsql & " bed_remain=bed_remain + 1 "
            Case cszAdditionalType
                szsql = szsql & " additional_remain=additional_remain + 1 "
            Case cszOtherType1
                szsql = szsql & " other_remain_1 =other_remain_1 + 1 "
            Case cszOtherType2
                szsql = szsql & " other_remain_2 =other_remain_2 + 1 "
        End Select
        szsql = szsql & " WHERE bus_id='" & szBusID & "' AND bus_date='" & ToDBDate(dtDate) & "'"
        poDb.Execute szsql
        
        '改变相应座位号的状态
        szsql = "UPDATE Env_bus_seat_lst SET " _
        & " status=" & ST_SeatCanSell & "," _
        & " ticket_no='' WHERE " _
        & " bus_id='" & szBusID & "' AND " _
        & " bus_date='" & ToDBDate(dtDate) & "' AND " _
        & " seat_no='" & szSeatNo & "'"
        
        poDb.Execute szsql
    End If
    
    '改变该车次相应到站的可售票数,加上了上车站条件
    szsql = "UPDATE Env_bus_station_lst SET " _
    & " sale_ticket_quantity=sale_ticket_quantity+1 WHERE " _
    & " bus_id='" & szBusID & "' AND " _
    & " bus_date='" & ToDBDate(dtDate) & "' AND " _
    & " station_id='" & szDesStationID & "' AND " _
    & " sell_station_id='" & szSellStationID & "' AND " _
    & " max_sale_quantity>0 "
    
    poDb.Execute szsql
    
End Sub

'检查了票状态，是否是同一人，是否已过正常退（废）时间,免票不能改签、退
Private Sub AssertTicketCanCancelReturn(pszTicketID As String, poDb As RTConnection, pnDealType As ETicketServiceType, pbForce As Boolean)
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim nTicketStatus As ETicketStatus
    Dim nTicketType As ETicketType
    Dim szSeller As String
    Dim dtSellDateTime As Date
    Dim oParam As New SystemParam
    Dim dtNowDateTime As Date
    Dim dtStartupDateTime As Date
    Dim lBusType As EBusType
    Dim dtBusDate As Date
    
    szsql = "SELECT Ticket_sell_lst.*,bus_start_time,bus_type FROM Ticket_sell_lst,Work_env_bus_info WHERE " _
    & "Ticket_sell_lst.bus_id=Work_env_bus_info.bus_id AND " _
    & "Ticket_sell_lst.bus_date=Work_env_bus_info.bus_date AND " _
    & "ticket_id='" & pszTicketID & "'"
    Set rsTemp = poDb.Execute(szsql)
    If rsTemp.RecordCount <> 1 Then ShowError ERR_TicketNotExist, , "票号为：" & pszTicketID
    
    nTicketStatus = FormatDbValue(rsTemp!Status)
    nTicketType = FormatDbValue(rsTemp!ticket_type)
    lBusType = FormatDbValue(rsTemp!bus_type)
    dtBusDate = FormatDbValue(rsTemp!bus_date)
    
    If (nTicketStatus And ST_TicketCanceled) <> 0 Then ShowError ERR_TicketIsCanceled, , "票号为：" & pszTicketID
    If (nTicketStatus And ST_TicketChanged) <> 0 Then ShowError ERR_TicketIsChanged, , "票号为：" & pszTicketID
    If (nTicketStatus And ST_TicketChecked) <> 0 Then ShowError ERR_TicketIsChecked, , "票号为：" & pszTicketID
    If (nTicketStatus And ST_TicketReturned) <> 0 Then ShowError ERR_TicketIsReturned, , "票号为：" & pszTicketID
    
    oParam.Init m_oActiveUser
    
    If pnDealType = TP_ReturnTicket Then '进行退票(免票不能退)
        If nTicketType = TP_FreeTicket Then ShowError ERR_FreeTicketCanNotReturnChange, , "票号为：" & pszTicketID
    Else '进行废票或改签
        If (nTicketStatus And ST_TicketSellChange) <> 0 Then ShowError ERR_TicketIsChangeTicket, , "票号为：" & pszTicketID
        
        If pnDealType = TP_CancelTicket Then '废票
            If Not pbForce Then
                szSeller = FormatDbValue(rsTemp!user_id)
                If Trim(szSeller) <> Trim(m_oActiveUser.UserID) Then ShowError ERR_UserIsNotTicketSeller, , "票号为：" & pszTicketID
            Else ' 强行废票
                'If m_oActiveUser.UserUnitID <> oParam.UnitID Then ShowError ERR_ForceCancelTicketRemoteNotAllowed, , "票号为：" & pszTicketID
            End If
        Else '改签
            If nTicketType = TP_FreeTicket Then ShowError ERR_FreeTicketCanNotReturnChange, , "票号为：" & pszTicketID
        End If
    End If

    
    If Not pbForce Then
        dtSellDateTime = FormatDbValue(rsTemp!operation_time)
        dtNowDateTime = oParam.NowDateTime
        
        If pnDealType = TP_ReturnTicket Then '进行退票
            If lBusType <> TP_ScrollBus Then
                dtStartupDateTime = SelfGetFullDateTime(FormatDbValue(rsTemp!bus_date), FormatDbValue(rsTemp!bus_start_time))
                If DateDiff("n", dtNowDateTime, dtStartupDateTime) < oParam.ReturnTicketTime3 Then ShowError ERR_ReturnTicketTimeOut, , "票号为：" & pszTicketID
            Else
                If oParam.NowDate > dtBusDate Then ShowError ERR_ScroolBusReturnTicketDateOut, , "票号为：" & pszTicketID
                If Not (oParam.ScrollBusCanReturnTicket Or pbForce) Then ShowError ERR_ScrollBusTicketCanNotReturn, , "票号为：" & pszTicketID
            End If
            
        ElseIf pnDealType = TP_ChangeTicket Then '进行改签
            If lBusType <> TP_ScrollBus Then
                dtStartupDateTime = SelfGetFullDateTime(FormatDbValue(rsTemp!bus_date), FormatDbValue(rsTemp!bus_start_time))
                If DateDiff("n", dtNowDateTime, dtStartupDateTime) < 0 Then ShowError ERR_ReturnTicketTimeOut, , "票号为：" & pszTicketID 'oParam.ReturnTicketTime3
            Else
                If oParam.NowDate > dtBusDate Then ShowError ERR_ScroolBusReturnTicketDateOut, , "票号为：" & pszTicketID
            End If
        Else '进行废票
            If DateDiff("n", dtSellDateTime, dtNowDateTime) > oParam.CancelTicketTime Then ShowError ERR_CancelTicketTimeOut, , "票号为：" & pszTicketID
        End If
    End If

'            '已过发车时间不允许改签
'            If dtBusStartupDateTime < oParam.NowDateTime Then
'                ShowError ERR_ChangeTicketTimeOut
'
'            End If
End Sub

'内部用废票
Private Sub SelfCancelTicket(paszTicketID() As String, poDb As RTConnection, pnCancelType As ECancelTicketType)
    Dim szsql As String
    Dim oParam As New SystemParam
    Dim szUserID As String
    Dim nLen As Integer
    Dim i As Integer
    Dim rsTemp1 As Recordset
    nLen = 0
    nLen = ArrayLength(paszTicketID)
    On Error GoTo Here
    For i = 1 To nLen
        szsql = "SELECT status FROM Ticket_sell_lst where ticket_id='" & paszTicketID(i) & "'"
        Set rsTemp1 = poDb.Execute(szsql)
        If rsTemp1.RecordCount <> 1 Then ShowError ERR_TicketNotExist, , "票号为：" & paszTicketID(i)
        
        If rsTemp1!Status = ST_TicketSellChange Then ShowError ERR_TicketIsChangeTicket, , "票号为：" & paszTicketID(i)
        
        szsql = "UPDATE Ticket_sell_lst SET " _
        & " status=status | " & ST_TicketCanceled & " WHERE " _
        & " ticket_id='" & paszTicketID(i) & "'"
        poDb.Execute szsql
        
        oParam.Init m_oActiveUser
        If pnCancelType = TP_ForceCancel Then
            Dim rsTemp As Recordset
            szsql = "SELECT user_id FROM Ticket_sell_lst WHERE ticket_id='" & paszTicketID(i) & "'"
            Set rsTemp = poDb.Execute(szsql)
           ' szUserID = FormatDbValue(rsTemp!user_id)
        Else
           
        End If
        
        szUserID = m_oActiveUser.UserID
        
        szsql = "INSERT INTO ticket_cancel_lst( " _
        & " ticket_id," _
        & " cancel_ticket_time," _
        & " user_id," _
        & " cancel_mode ) VALUES('" _
        & paszTicketID(i) & "','" _
        & ToDBDateTime(oParam.NowDateTime) & "','" _
        & szUserID & "'," _
        & pnCancelType & ")"
        
        poDb.Execute szsql
        

        '回收座位
        RecycleSeat paszTicketID(i), poDb
    Next i
    On Error GoTo 0
    Exit Sub
Here:
    err.Raise err.Number
End Sub

'内部用退票方法
Private Sub SelfReturnlTicket(paszTicketID() As String, pszSheetID As String, pasgOrgReturnCharge() As Double, poDb As RTConnection, pnReturnType As EReturnTicketType, pReturnRatio() As Double)
    Dim szsql As String
    Dim oParam As New SystemParam
    Dim sgReturnCharge As Double, sgTicketPrice As Double
    Dim dtBusStartupDateTime As Date
    Dim rsTemp As Recordset
    Dim lBusType As EBusType
    
    Dim i As Integer
    Dim nLen As Integer
    nLen = 0
    nLen = ArrayLength(paszTicketID)
    oParam.Init m_oActiveUser
    For i = 1 To nLen
        
            
            szsql = "SELECT Work_env_bus_info.* , ticket_price FROM Ticket_sell_lst,Work_env_bus_info WHERE " _
            & " ticket_id='" & paszTicketID(i) & "' AND " _
            & " Ticket_sell_lst.bus_id= Work_env_bus_info.bus_id AND " _
            & " Ticket_sell_lst.bus_date=Work_env_bus_info.bus_date "
            
            Set rsTemp = poDb.Execute(szsql)
            sgTicketPrice = FormatDbValue(rsTemp!ticket_price)
            dtBusStartupDateTime = SelfGetFullDateTime(FormatDbValue(rsTemp!bus_date), FormatDbValue(rsTemp!bus_start_time))
            lBusType = FormatDbValue(rsTemp!bus_type)
            'If FormatDbValue(rsTemp!is_all_refundment) = 1 Then '是全额退票
               ' sgReturnCharge = 0
           ' Else
             sgReturnCharge = pasgOrgReturnCharge(i)
    '            If lBusType = TP_RegularBus Then
    '                'sgReturnCharge = Round(sgTicketPrice * oParam.GetReturnChargeRatio(dtBusStartupDateTime), 2)
    '                sgReturnCharge = oParam.GetReturnCharge(dtBusStartupDateTime, sgTicketPrice)
    '            Else
    '                sgReturnCharge = Round(oParam.ScrollBusReturnRatio * sgTicketPrice + 0.01, 1)
    '            End If
           ' End If
     
    '    If sgReturnCharge <> psgOrgReturnCharge Then ShowError ERR_OrgReturnChargeOutDated
        If sgTicketPrice = sgReturnCharge And sgTicketPrice <> 0 Then
            err.Raise ERR_SellTicketClient, , "手续费为100%,不允许退票"
        End If
        szsql = "UPDATE Ticket_sell_lst SET " _
        & " status=status | " & ST_TicketReturned & " WHERE " _
        & " ticket_id='" & paszTicketID(i) & "'"
        poDb.Execute szsql
        
        '退票费为0也应该写退票表
        'If sgReturnCharge > 0 Then
            szsql = "INSERT INTO Ticket_return_lst( " _
            & " ticket_id," _
            & " operation_time," _
            & " credence_sheet_id," _
            & " return_charge," _
            & " user_id," _
            & " return_ticket_mode,return_ratio ) VALUES('" _
            & paszTicketID(i) & "','" _
            & ToDBDateTime(oParam.NowDateTime) & "','" _
            & pszSheetID & "'," _
            & sgReturnCharge & ",'" _
            & m_oActiveUser.UserID & "'," _
            & pnReturnType & "," _
            & pReturnRatio(i) & ")"
            
            poDb.Execute szsql
            

        'End If
        
        '回收座位
        RecycleSeat paszTicketID(i), poDb
    Next i
    Set rsTemp = Nothing
    
End Sub

'//////////////////////////
'取消退票
Public Function ResumeReturnTicket(paszTicketID() As String) As String()
    Dim oDb As New RTConnection
    Dim i As Integer
    Dim nLen As Integer
    
    
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_ResumeReturnTicket
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    oDb.BeginTrans
        On Error GoTo Here
        
        ResumeReturnTicket = SelfResumeReturnlTicket(paszTicketID, oDb)
        
    oDb.CommitTrans
    On Error GoTo 0
    For i = 1 To nLen
    WriteOperateLog m_oActiveUser, RIGHT_ResumeReturnTicket, "取消退票票号:" & paszTicketID(i)
    Next i
    Exit Function
Here:
    oDb.RollbackTrans
    err.Raise err.Number
End Function
'强行退票
Public Sub ForceReturnTicket(paszTicketID() As String, pszReturnSheet As String, pOrgReturnCharge() As Double, pasgReturnRatio() As Double)
    Dim oDb As New RTConnection
    Dim i As Integer
    Dim nLen As Integer
    Dim asgReturnCharge() As Double
    Dim asgReturnRatio() As Double
    nLen = 0
    nLen = ArrayLength(paszTicketID)
    ReDim asgReturnCharge(1 To nLen)
    ReDim asgReturnRatio(1 To nLen)
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_ForceReturnTicket
    
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    oDb.BeginTrans
    
    For i = 1 To nLen
        AssertTicketCanCancelReturn paszTicketID(i), oDb, TP_ReturnTicket, True
        asgReturnCharge(i) = pOrgReturnCharge(i)
        asgReturnRatio(i) = pasgReturnRatio(i)
    Next i
        
    On Error GoTo Here
        SelfReturnlTicket paszTicketID, "", asgReturnCharge, oDb, TP_ForceReturn, asgReturnRatio

    oDb.CommitTrans
    
    
'    WriteOperateLog m_oActiveUser, RIGHT_ForceReturnTicket, "强行退票票号:" & TicketID
    On Error GoTo 0
    Exit Sub
Here:
    oDb.RollbackTrans
    err.Raise err.Number
End Sub

'删除了折扣参数
Public Function ChangeTicket(BusStartupDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String, SellTicketInfo() As TSellTicketParam, panInsurance() As Integer, Optional aszRealNameInfo As Variant) As TSellTicketResult()
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim dtBusStartupDateTime As Date
    Dim sgChangeCharge As Double
    Dim sgOrgTicketPrice As Double
    Dim srTemp() As TSellTicketResult
    Dim iLen As Integer
    Dim nCount As Integer
    
    Dim i As Integer
    Dim nLen As Integer
    
    Dim aBuyTicketInfo(1 To 1) As TBuyTicketInfo
    iLen = 0
    
    iLen = ArrayLength(BusID)
    ReDim srTemp(1 To iLen)
    
    
    Dim oParam As New SystemParam
    
    oParam.Init m_oActiveUser
    
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_ChangeTicket
    
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    oDb.BeginTrans
    
    Dim nStart As Integer
    For nCount = 1 To iLen
        nLen = ArrayLength(SellTicketInfo(nCount).aszOrgTicket)
        For i = 1 To nLen
            AssertTicketCanCancelReturn SellTicketInfo(nCount).aszOrgTicket(i), oDb, TP_ChangeTicket, False
            
            szsql = "SELECT Work_env_bus_info.* , ticket_price, ticket_sell_lst.sell_station_id, user_info.sell_station_id seller_sell_station_id  FROM Ticket_sell_lst,Work_env_bus_info, user_info WHERE " _
            & " ticket_id='" & SellTicketInfo(nCount).aszOrgTicket(i) & "' AND " _
            & " Ticket_sell_lst.bus_id= Work_env_bus_info.bus_id AND " _
            & " Ticket_sell_lst.bus_date=Work_env_bus_info.bus_date AND " _
            & " Ticket_sell_lst.user_id = user_info.user_id "
            Set rsTemp = oDb.Execute(szsql)
            sgOrgTicketPrice = FormatDbValue(rsTemp!ticket_price)
            
            dtBusStartupDateTime = SelfGetFullDateTime(FormatDbValue(rsTemp!bus_date), FormatDbValue(rsTemp!bus_start_time))
            
'            '已过发车时间不允许改签
'            If dtBusStartupDateTime < oParam.NowDateTime Then
'                ShowError ERR_ChangeTicketTimeOut
'
'            End If
            
            '这里要对sgChangeCharge赋值
            If FormatDbValue(rsTemp!is_all_refundment) = 1 Then '是全额退票
                sgChangeCharge = 0
            Else
                sgChangeCharge = oParam.ChangeCharge
            End If
            
            
            On Error GoTo Here
            
            
            If rsTemp.RecordCount = 1 Then
                szsql = "UPDATE Env_bus_allot_lst SET sell_quantity=sell_quantity-1 WHERE bus_id='" & rsTemp!bus_id & "'" _
                       & " AND sell_station_id='" & rsTemp!sell_station_id & "' AND bus_date='" & rsTemp!bus_date & "'"
                oDb.Execute szsql
                
                szsql = "UPDATE env_bus_sell_station_lst SET sell_quantity=sell_quantity-1 WHERE bus_id='" & rsTemp!bus_id & "'" _
                       & " AND sell_station_id='" & rsTemp!seller_sell_station_id & "' AND bus_date='" & rsTemp!bus_date & "'"
                oDb.Execute szsql
                
                '处理并班车次的座位
                szsql = " UPDATE env_bus_seat_lst SET status = " & ST_SeatCanSell & " , ticket_no = ''" _
                        & " FROM env_bus_seat_lst a, " _
                        & " (SELECT ebsl.*, msl.merge_bus_id FROM mergebus_split_lst msl, env_bus_seat_lst ebsl " _
                        & " WHERE msl.bus_id = ebsl.bus_id And msl.bus_date = ebsl.bus_date" _
                        & " AND msl.merge_bus_id = " & TransFieldValueToString(rsTemp!bus_id) _
                        & " AND msl.bus_date = " & TransFieldValueToString(rsTemp!bus_date) & " AND ebsl.status = " & TransFieldValueToString(ST_SeatSlitp) _
                        & " AND ebsl.ticket_no = " & TransFieldValueToString(SellTicketInfo(nCount).aszOrgTicket(i)) & " ) b " _
                        & " WHERE a.bus_id = b.bus_id And a.bus_date = b.bus_date And a.seat_no = b.seat_no"
                oDb.Execute szsql
                
            End If
            
            
            '修改票的状态为已被改签
            
            szsql = "UPDATE Ticket_sell_lst SET " _
            & " status=status | " & ST_TicketChanged & " WHERE " _
            & " ticket_id='" & SellTicketInfo(nCount).aszOrgTicket(i) & "'"
            oDb.Execute szsql
            
            aBuyTicketInfo(1) = SellTicketInfo(nCount).BuyTicketInfo(i)
            nStart = i - 1
            srTemp(nCount) = SelfSellTicket(BusStartupDate(nCount), BusID(nCount), SellStationID(nCount), DesStationID(nCount), aBuyTicketInfo, oDb, SellTicketInfo(nCount).pasgSellTicketPrice, panInsurance(nCount), False, ST_TicketSellChange, nStart, aszRealNameInfo)
            
            oParam.Init m_oActiveUser
            
            '改签费为0也应写表
            'If sgChangeCharge > 0 Then
            
                '写改签明细
                szsql = " INSERT INTO ticket_change_lst(" _
                & " new_ticket_id," _
                & " change_time," _
                & " change_charge," _
                & " former_ticket_id," _
                & " former_ticket_price," _
                & " credence_sheet_id," _
                & " user_id ) VALUES('" _
                & SellTicketInfo(nCount).BuyTicketInfo(i).szTicketNo & "','" _
                & ToDBDateTime(oParam.NowDateTime) & "'," _
                & sgChangeCharge & ",'" _
                & SellTicketInfo(nCount).aszOrgTicket(i) & "'," _
                & sgOrgTicketPrice & ",'" _
                & SellTicketInfo(nCount).aszChangeSheetID(i) & "','" _
                & m_oActiveUser.UserID & "')"
                
                oDb.Execute szsql
            'End If
            '回收座位
            RecycleSeat SellTicketInfo(nCount).aszOrgTicket(i), oDb
            srTemp(nCount).szReserved = sgChangeCharge
            WriteOperateLog m_oActiveUser, RIGHT_ChangeTicket, "改签票号:" & SellTicketInfo(nCount).BuyTicketInfo(i).szTicketNo _
                                               & "原票号：" & SellTicketInfo(nCount).aszOrgTicket(i)
        Next i
        nStart = nStart + nLen
    Next nCount
    oDb.CommitTrans
    
    ChangeTicket = srTemp

'    On Error GoTo 0
   
    
    Exit Function
Here:
    oDb.RollbackTrans
    err.Raise err.Number
End Function

'得到指定日期的所有可能到站及其相关信息。
'日期比较
Public Function GetStationRs(BusDate As Date, Optional IgnoreOffTime As Boolean = False) As Recordset
    Dim oDb As New RTConnection
    Dim oParam As New SystemParam
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim nStopSellTime As Integer
    Dim dtNowDateTime As Date

    oParam.Init m_oActiveUser
    '玉环限制每个售票窗口的停售时间，fpd
    Dim oReg As New CFreeReg
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    nStopSellTime = oReg.GetSetting("SellTk", "StopSellTime", oParam.StopSellTime)
'    nStopSellTime = oParam.StopSellTime
    dtNowDateTime = oParam.NowDateTime

    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    szsql = "SELECT DISTINCT station_info.station_id,station_name ,station_input_code FROM " _
    & " Work_env_bus_info,Work_env_bus_station_lst,station_info ,bus_type_code WHERE "

    If Not IgnoreOffTime Then
        szsql = szsql & "( '" & ToDBDateTime(DateAdd("n", nStopSellTime, dtNowDateTime)) & "'<=Work_env_bus_info.bus_start_time OR " _
        & " Work_env_bus_info.bus_type=" & TP_ScrollBus & ") AND "

        szsql = szsql & " Work_env_bus_info.bus_date=Work_env_bus_station_lst.bus_date AND " _
        & " Work_env_bus_info.bus_id=Work_env_bus_station_lst.bus_id AND " _
        & " Work_env_bus_info.bus_date='" & ToDBDate(BusDate) & "' AND " _
        & " (sale_seat_quantity>0 OR sale_stand_seat_quantity>0 OR Work_env_bus_info.bus_type=" & TP_ScrollBus & ") AND " _
        & " station_info.station_id=Work_env_bus_station_lst.station_id AND " _
        & " ( ( max_sale_quantity>0 AND  sale_ticket_quantity>0) OR  " _
        & " ( max_sale_quantity<0) ) AND " _
        & " Work_env_bus_info.bus_type=bus_type_code.bus_type " _
        & " ORDER BY station_info.station_id"
    Else
        szsql = szsql _
        & " (Work_env_bus_info.bus_type<>" & TP_ScrollBus & ") AND "

        szsql = szsql & " Work_env_bus_info.bus_date=Work_env_bus_station_lst.bus_date AND " _
        & " Work_env_bus_info.bus_id=Work_env_bus_station_lst.bus_id AND " _
        & " Work_env_bus_info.bus_date='" & ToDBDate(BusDate) & "' AND " _
        & " (sale_seat_quantity>0 OR sale_stand_seat_quantity>0 OR Work_env_bus_info.bus_type=" & TP_ScrollBus & ") AND " _
        & " station_info.station_id=Work_env_bus_station_lst.station_id AND " _
        & " ( ( max_sale_quantity>0 AND  sale_ticket_quantity>0) OR  " _
        & " ( max_sale_quantity<0) ) " _
        & " Work_env_bus_info.bus_type=bus_type_code.bus_type " _
        & " ORDER BY station_info.station_id"
    End If


    Set rsTemp = oDb.Execute(szsql)
    Set GetStationRs = rsTemp

End Function


'得到指定日期能到指定车站的所有车次及其相关信息
'日期比较
Public Function GetBusRs(BusDate As Date, DesStationID As String, Optional IgnoreOffTime As Boolean = False, Optional BusID As String = "", Optional IsDiscount As Boolean = False) As Recordset
    Dim oDb As New RTConnection
    Dim oParam As New SystemParam
    Dim szsql As String
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
    '玉环限制每个售票窗口的停售时间，fpd
    Dim oReg As New CFreeReg
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    nStopSellTime = oReg.GetSetting("SellTk", "StopSellTime", oParam.StopSellTime)
'    nStopSellTime = oParam.StopSellTime
    dtNowDateTime = oParam.NowDateTime
    
    dbInternetBusShow = oParam.InternetBusShow
    
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    szSelect = "SELECT Work_env_bus_info.* ,e.check_gate_id sell_check_gate_id ,seat_type_id,route_name ,Work_env_bus_station_lst.mileage end_station_mileage" _
        & " ,Work_env_bus_station_lst.sell_station_id,sell_station_name,e.bus_start_time as busstarttime, " _
        & "         full_price, half_price,preferential_ticket1,preferential_ticket2,preferential_ticket3," _
        & "         sale_ticket_quantity,stop_sale_time, Work_env_bus_station_lst.book_count ,route_info.id_card "
        
    szTable = "  FROM Work_env_bus_info,Work_env_bus_station_lst ,route_info,sell_station_info s,Work_env_bus_allot_lst e  "
    
    
    szWhere = "  WHERE s.sell_station_id=Work_env_bus_station_lst.sell_station_id AND e.sell_station_id=s.sell_station_id AND " _
        & "  e.bus_date=Work_env_bus_info.bus_date AND e.bus_id=Work_env_bus_info.bus_id AND "
    
    If Not IgnoreOffTime Then
        szWhere = szWhere & "(('" & ToDBDateTime(DateAdd("s", nStopSellTime * 60, dtNowDateTime)) & "'<=e.bus_start_time " _
        & IIf(Not m_bAllowScrollBusSaleForever, "", " AND bus_type<>" & TP_ScrollBus) _
        & ")  " _
        & " OR ( " _
        & IIf(Not m_bAllowScrollBusSaleForever, "", "bus_type=" & TP_ScrollBus & " OR  ") _
        & " e.bus_start_time>='" & ToDBDateTime(DateAdd("s", nStopSellTime * 60, dtNowDateTime)) & "'))   "
        
        szWhere = szWhere _
        & " AND Work_env_bus_info.bus_date='" & ToDBDate(BusDate) & "' AND " _
        & " Work_env_bus_info.bus_id=Work_env_bus_station_lst.bus_id AND " _
        & " Work_env_bus_info.bus_date=Work_env_bus_station_lst.bus_date AND " _
        & " Work_env_bus_station_lst.station_id='" & DesStationID & "' AND " _
        & " ( sale_seat_quantity>=0 OR sale_stand_seat_quantity>0 " _
        & IIf(Not m_bAllowScrollBusSaleForever, "", " OR bus_type=" & TP_ScrollBus) _
        & ") AND " _
        & " ( ( max_sale_quantity>0 AND  sale_ticket_quantity>=0) OR  " _
        & " ( max_sale_quantity<0) ) AND Work_env_bus_info.internet_status=" & CnInternetCanSell & " " _
        & " AND e.bus_start_time >= '" & DateAdd("n", dbInternetBusShow, dtNowDateTime) & "' " _
        & " AND Work_env_bus_info.route_id=route_info.route_id  AND (e.can_sale_quantity <>0 )" _
        'ORDER BY Work_env_bus_info.bus_start_time,Work_env_bus_info.bus_id  "
    
    Else
        dbExtraBeforeTime = oParam.ExtraBusShowBefore
        dbExtraAfterTime = oParam.ExtraBusShowAfter
        
        szWhere = szWhere & "'" & ToDBDateTime(DateAdd("n", dbExtraAfterTime, dtNowDateTime)) & "'>=e.bus_start_time AND " _
        & "'" & ToDBDateTime(DateAdd("n", -dbExtraBeforeTime, dtNowDateTime)) & "'<=e.bus_start_time " _
        & IIf(Not m_bAllowScrollBusSaleForever, "", " AND bus_type<>" & TP_ScrollBus) _
        & " AND  " _
        & " Work_env_bus_info.bus_date='" & ToDBDate(BusDate) & "' AND " _
        & " Work_env_bus_info.bus_id=Work_env_bus_station_lst.bus_id AND " _
        & " Work_env_bus_info.bus_date=Work_env_bus_station_lst.bus_date AND " _
        & " Work_env_bus_station_lst.station_id='" & DesStationID & "' AND " _
        & " ( sale_seat_quantity>=0 OR sale_stand_seat_quantity>0 " _
        & IIf(Not m_bAllowScrollBusSaleForever, "", " OR bus_type=" & TP_ScrollBus) _
        & ") AND " _
        & " ( ( max_sale_quantity>0 AND  sale_ticket_quantity>0) OR  " _
        & " ( max_sale_quantity<0) )  AND Work_env_bus_info.internet_status=" & CnInternetCanSell _
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
'得到指定车次日期的座位详细情况
Public Function GetSeatRs(ByVal BusDate As Date, ByVal BusID As String) As Recordset
    Dim oDb As New RTConnection
    Dim szsql As String
'    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    szsql = "SELECT e.seat_no,e.status,e.ticket_no,s.ticket_type FROM Env_bus_seat_lst e,ticket_sell_lst s WHERE " _
    & " e.bus_id='" & BusID & "' AND " _
    & " e.bus_date='" & ToDBDate(BusDate) & "'" _
    & " and e.ticket_no*=s.ticket_id "
    Set GetSeatRs = oDb.Execute(szsql)
End Function


'得到指定车次日期的已被预定但未被购详细情况
Public Function GetBookRs(ByVal BusDate As Date, ByVal BusID As String) As Recordset
    Dim oDb As New RTConnection
    Dim szsql As String
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    '得到指定车次所有的已被预定但未被购的详细情况
    szsql = "SELECT * FROM ticket_book_lst  " _
    & " WHERE bus_id='" & BusID & "'" _
    & " AND bus_date='" & ToDBDate(BusDate) & "'" _
    & " AND status = 0 "
    '0,已被预定,1已取消预定,2已被预人购买
    Set GetBookRs = oDb.Execute(szsql)
End Function



'得到指定售票员，指定时间内的售票情况明细,不包括改签售出的。
Public Function GetSellTicketRs(QueryDate As Variant, UserID As String) As Recordset
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim dtSellDateTime As Date
    Dim oParam As New SystemParam
    
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    If UserID = m_oActiveUser.UserID Then
        AssertHaveRight m_oActiveUser, RIGHT_QuerySelfThings
    Else
        AssertHaveRight m_oActiveUser, RIGHT_QueryOtherThings
    End If
    
    oParam.Init m_oActiveUser
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    If IsDate(QueryDate) Then
        dtSellDateTime = QueryDate
    ElseIf IsNumeric(QueryDate) Then
        dtSellDateTime = DateAdd("n", -QueryDate, oParam.NowDateTime)
    Else
        ShowError ERR_InputParamTypeErr
    End If
    
    szsql = "SELECT st.* ,si.station_name AS des_station_name,ssi.sell_station_name FROM Ticket_sell_lst st,station_info si,sell_station_info ssi WHERE " _
    & " st.user_id='" & UserID & "' AND " _
    & " st.operation_time>'" & ToDBDateTime(dtSellDateTime) & "' AND " _
    & " st.status & " & ST_TicketNormal & "<>0 AND " _
    & " st.des_station_id=si.station_id AND st.sell_station_id=ssi.sell_station_id"
    Set GetSellTicketRs = oDb.Execute(szsql)

End Function

'得到指定售票员，指定时间内的废票情况明细
Public Function GetCancelTicketRs(QueryDate As Variant, UserID As String) As Recordset
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim dtCancelDateTime As Date
    Dim oParam As New SystemParam
    
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    If UserID = m_oActiveUser.UserID Then
        AssertHaveRight m_oActiveUser, RIGHT_QuerySelfThings
    Else
        AssertHaveRight m_oActiveUser, RIGHT_QueryOtherThings
    End If
    
    oParam.Init m_oActiveUser
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    If IsDate(QueryDate) Then
        dtCancelDateTime = QueryDate
    ElseIf IsNumeric(QueryDate) Then
        dtCancelDateTime = DateAdd("n", -QueryDate, oParam.NowDateTime)
    Else
        ShowError ERR_InputParamTypeErr
    End If

    szsql = "SELECT sell_station_info.*, cancel_ticket_time ,cancel_mode,station_name AS des_station_name,sell_station_info.sell_station_name FROM " _
    & " ticket_cancel_lst,Ticket_sell_lst,station_info,sell_station_info WHERE  " _
    & " ticket_cancel_lst.user_id='" & UserID & "' AND " _
    & " cancel_ticket_time>'" & ToDBDateTime(dtCancelDateTime) & "' AND " _
    & " ticket_cancel_lst.ticket_id=Ticket_sell_lst.ticket_id AND " _
    & " des_station_id=station_info.station_id AND sell_station_info.sell_station_id=sell_station_info.sell_station_id"
    
    Set GetCancelTicketRs = oDb.Execute(szsql)
End Function



'得到指定售票员，指定时间内的改签售票情况明细
Public Function GetChangeTicketRs(QueryDate As Variant, UserID As String) As Recordset
    
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim dtChangeDateTime As Date
    Dim oParam As New SystemParam
    
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    If UserID = m_oActiveUser.UserID Then
        AssertHaveRight m_oActiveUser, RIGHT_QuerySelfThings
    Else
        AssertHaveRight m_oActiveUser, RIGHT_QueryOtherThings
    End If
    
    oParam.Init m_oActiveUser
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    If IsDate(QueryDate) Then
        dtChangeDateTime = QueryDate
    ElseIf IsNumeric(QueryDate) Then
        dtChangeDateTime = DateAdd("n", -QueryDate, oParam.NowDateTime)
    Else
        ShowError ERR_InputParamTypeErr
    End If
    
    szsql = "SELECT Ticket_sell_lst.* ,new_ticket_id, change_time,former_ticket_id," _
    & " former_ticket_price,change_charge , credence_sheet_id, station_name AS des_station_name,sell_station_info.sell_station_name " _
    & " FROM ticket_change_lst,Ticket_sell_lst,station_info,sell_station_info WHERE " _
    & " ticket_change_lst.user_id='" & UserID & "' AND " _
    & " change_time>'" & ToDBDateTime(dtChangeDateTime) & "' AND " _
    & " ticket_change_lst.new_ticket_id=Ticket_sell_lst.ticket_id AND " _
    & " des_station_id=station_info.station_id AND sell_station_info.sell_station_id=Ticket_sell_lst.sell_station_id"
    
    Set GetChangeTicketRs = oDb.Execute(szsql)
    
End Function

'得到指定售票员，指定时间内的退票情况明细
Public Function GetReturnTicketRs(QueryDate As Variant, UserID As String) As Recordset
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim dtReturnDateTime As Date
    Dim oParam As New SystemParam
    
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    If UserID = m_oActiveUser.UserID Then
        AssertHaveRight m_oActiveUser, RIGHT_QuerySelfThings
    Else
        AssertHaveRight m_oActiveUser, RIGHT_QueryOtherThings
    End If
    
    oParam.Init m_oActiveUser
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    If IsDate(QueryDate) Then
        dtReturnDateTime = QueryDate
    ElseIf IsNumeric(QueryDate) Then
        dtReturnDateTime = DateAdd("n", -QueryDate, oParam.NowDateTime)
    Else
        ShowError ERR_InputParamTypeErr
    End If

    szsql = "SELECT Ticket_sell_lst.*, Ticket_return_lst.operation_time AS return_time ,credence_sheet_id,return_charge," _
    & " return_ticket_mode,station_name AS des_station_name FROM " _
    & " Ticket_return_lst,Ticket_sell_lst,station_info WHERE  " _
    & " Ticket_return_lst.user_id='" & UserID & "' AND " _
    & " Ticket_return_lst.operation_time>'" & ToDBDateTime(dtReturnDateTime) & "' AND " _
    & " Ticket_return_lst.ticket_id=Ticket_sell_lst.ticket_id AND " _
    & " des_station_id=station_id"

    Set GetReturnTicketRs = oDb.Execute(szsql)
    Set oDb = Nothing
End Function


'给定车次,得到站点
Public Function GetStationFromBusRs(BusDate As Date, BusID As String) As Recordset
    Dim szsql As String
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    szsql = "SELECT Work_env_bus_station_lst.sell_station_id,s.sell_station_name,station_info.station_id,station_name ,station_input_code,seat_type_id,end_station_name," _
    & "             sale_ticket_quantity,stop_sale_time," _
    & "             bus_type,full_price,half_price,preferential_ticket1,preferential_ticket2,preferential_ticket3 " _
    & "      FROM " _
    & " Work_env_bus_station_lst,station_info,Work_env_bus_info,sell_station_info s WHERE " _
    & " Work_env_bus_info.bus_id=Work_env_bus_station_lst.bus_id AND " _
    & " Work_env_bus_info.bus_date=Work_env_bus_station_lst.bus_date AND " _
    & " Work_env_bus_station_lst.sell_station_id=s.sell_station_id AND " _
    
    szsql = szsql _
    & " Work_env_bus_info.bus_id='" & BusID & "'  AND " _
    & " Work_env_bus_info.bus_date='" & ToDBDate(BusDate) & "' AND " _
    & " station_info.station_id=Work_env_bus_station_lst.station_id AND " _
    & " ( ( max_sale_quantity>0 AND  sale_ticket_quantity>0) OR  " _
    & " ( max_sale_quantity<0) ) AND " _
    & " ( sale_seat_quantity>0 OR sale_stand_seat_quantity>0 OR bus_type=1) " _
    
    If Not m_bSellStationCanSellEachOther Then
      If m_oActiveUser.SellStationID <> "" Then
           szsql = szsql & " AND Work_env_bus_station_lst.sell_station_id='" & Trim(m_oActiveUser.SellStationID) & "'"
      End If
    End If
      szsql = szsql & " ORDER BY station_info.station_id,Work_env_bus_station_lst.seat_type_id"
    
    Set GetStationFromBusRs = oDb.Execute(szsql)
    
    Set oDb = Nothing
End Function

'得到指定售票员最后的一张票号
Public Function GetLastTicketNo(Seller As String) As String
    Dim szsql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    szsql = "SELECT MAX(ticket_id) AS ticket_no FROM Ticket_sell_lst WHERE " _
    & " operation_time=(" _
    & " SELECT MAX(operation_time) FROM Ticket_sell_lst WHERE " _
    & " user_id='" & Seller & "') "
    Set rsTemp = oDb.Execute(szsql)
    If rsTemp.RecordCount = 1 Then
        GetLastTicketNo = FormatDbValue(rsTemp!ticket_no)
    Else
        Dim oParam As New SystemParam
        oParam.Init m_oActiveUser
        GetLastTicketNo = String(oParam.TicketNumberLen, "0")
    End If
End Function

Public Property Get StationID() As String
    Dim oParam As SystemParam
    Set oParam = New SystemParam
    oParam.Init m_oActiveUser
    StationID = oParam.StationID
End Property

Public Property Get StationName() As String
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    
    szsql = "SELECT station_name FROM station_info WHERE " _
    & " station_id='" & StationID & "'"
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    Set rsTemp = oDb.Execute(szsql)
    If rsTemp.RecordCount = 1 Then
        StationName = FormatDbValue(rsTemp!station_name)
    Else
        szsql = "SELECT unit_short_name FROM connect_unit_info WHERE " _
        & " unit_id='" & m_oActiveUser.ServerUnit & "'"
        Set rsTemp = oDb.Execute(szsql)
        If rsTemp.RecordCount = 1 Then
            StationName = FormatDbValue(rsTemp!unit_short_name)
        End If
    End If
End Property

'得到当前系统中的所有站点信息
Public Function GetAllStationRs() As Recordset
    Dim szsql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    szsql = "SELECT station_id,station_name ,station_input_code FROM station_info ORDER BY station_id"
    Set rsTemp = oDb.Execute(szsql)
    Set GetAllStationRs = rsTemp
End Function

Public Sub WriteOffTicket(TicketID As String)
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim eStatus As ETicketStatus
    
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_WriteOffTiecket

    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    On Error GoTo there
    szsql = "SELECT * FROM Ticket_sell_lst WHERE ticket_id='" & TicketID & "'"
    Set rsTemp = oDb.Execute(szsql)
    If rsTemp.RecordCount < 1 Then ShowError ERR_NoExistWriteOffTiecket
    On Error GoTo 0
    
    On Error GoTo Here
    oDb.BeginTrans
    eStatus = FormatDbValue(rsTemp!Status)
    
    If Not ((eStatus And ST_TicketCanceled) And (eStatus And ST_TicketChanged) And (eStatus And ST_TicketReturned) = 0) Then
        RecycleSeat TicketID, oDb
    End If
    DeleteServerSellTicket TicketID, oDb
        
    
    
'    Select Case eStatus
'        Case ST_TicketNormal
'            '正常
'            RecycleSeat TicketID, oDb
'            DeleteServerSellTicket TicketID, oDb
'        Case ST_TicketChanged
'            '被改签
'            DeleteServerSellTicket TicketID, oDb
'        Case ST_TicketCanceled
'            '废票
'            DeleteServerSellTicket TicketID, oDb
'        Case ST_TicketReturned
'            '退票
'            DeleteServerSellTicket TicketID, oDb
'        Case ST_TicketSellChange
'            '改签售出
'            RecycleSeat TicketID, oDb
'            DeleteServerSellTicket TicketID, oDb
'        Case ST_TicketChecked
'            '已检
'            RecycleSeat TicketID, oDb
'            DeleteServerSellTicket TicketID, oDb
'        Case Else
'            '''''
'    End Select
    
    oDb.CommitTrans
    WriteOperateLog m_oActiveUser, RIGHT_WriteOffTiecket, "注销服务端车票[:" & TicketID & "]"
    
'    On Error GoTo 0
'    WriteOperateLog m_oActiveUser, RIGHT_SellTicket
    
Exit Sub
there:
    ShowError ERR_WriteOffTiecketFail
Exit Sub
Here:
    oDb.RollbackTrans
    ShowError ERR_WriteOffTiecketFail

End Sub
'内部删除票信息
'包括表:Ticket_sell_lst,ticket_price_lst,Ticket_return_lst,ticket_change_lst,ticket_cancel_lst,check_ticket_lst中主键为此票号的纪录
Private Sub DeleteServerSellTicket(TicketID As String, poDb As RTConnection)
    Dim szsql As String
    
    szsql = "DELETE ticket_price_lst WHERE ticket_id='" & TicketID & "'"
    poDb.Execute szsql
    szsql = "DELETE Ticket_return_lst WHERE ticket_id='" & TicketID & "'"
    poDb.Execute szsql
    szsql = "DELETE ticket_change_lst WHERE new_ticket_id='" & TicketID & "'"
    poDb.Execute szsql
    szsql = "DELETE ticket_cancel_lst WHERE ticket_id='" & TicketID & "'"
    poDb.Execute szsql
    szsql = "DELETE check_ticket_lst WHERE ticket_id='" & TicketID & "'"
    poDb.Execute szsql
    szsql = "DELETE Ticket_sell_lst WHERE ticket_id='" & TicketID & "'"
    poDb.Execute szsql
End Sub


Private Function UpdateSeatStatusStr2(pszBusID As String, pszDate As Date, pszSeatNo As String, pszTicketID As String, pbBookValid As Boolean, Optional pszBookNumber As String = "") As String
    Dim szsql As String
    szsql = "UPDATE Env_bus_seat_lst SET status=" & ST_SeatSold & "," _
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
    UpdateSeatStatusStr2 = szsql
End Function
Private Function UpdateSeatStatusStr3(pszBusID As String, pszBusDate As Date, pszSeatNo As String, pszTicketID As String) As String
    Dim szsql As String
    szsql = "UPDATE Env_bus_seat_lst SET status=" & ST_SeatSold & "," _
            & " ticket_no='" & pszTicketID & "' " _
            & " WHERE bus_id='" & pszBusID & "' AND " _
            & " bus_date='" & ToDBDate(pszBusDate) & "' AND " _
            & " seat_no='" & pszSeatNo & "' AND " _
            & " status=" & ST_SeatProjectBooked & " "
    UpdateSeatStatusStr3 = szsql
End Function

'/////////////////////////////
'取消废票
Public Function ResumeCancelTicket(paszTicketID() As String) As String()
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim i As Integer
    Dim nLen As Integer
    
    
    nLen = 0
    nLen = ArrayLength(paszTicketID)
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_ResumeCancelTicket
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    oDb.BeginTrans
        On Error GoTo Here
        ResumeCancelTicket = SelfResumeCancelTicket(paszTicketID, oDb)
    oDb.CommitTrans
    For i = 1 To nLen
        WriteOperateLog m_oActiveUser, RIGHT_ResumeCancelTicket, "取消废票票号:" & paszTicketID(i)
    Next i
    Exit Function
Here:
    oDb.RollbackTrans
    err.Raise err.Number
End Function
'正常退票
Public Sub ReturnTicket(paszTicketID() As String, pszSheetID As String, pOrgReturnCharge() As Double, pReturnRatio() As Double)
    Dim oDb As New RTConnection
    Dim nLen As Integer
    Dim i As Integer
    nLen = 0
    nLen = ArrayLength(paszTicketID)
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_ReturnTicket
    
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    On Error GoTo Here
    oDb.BeginTrans
        For i = 1 To nLen
            AssertTicketCanCancelReturn paszTicketID(i), oDb, TP_ReturnTicket, False
        Next i
        
        SelfReturnlTicket paszTicketID, pszSheetID, pOrgReturnCharge, oDb, TP_NormalReturn, pReturnRatio
        'RecycleSeat TicketID, oDb
   oDb.CommitTrans
        
    '    On Error GoTo 0
    For i = 1 To nLen
        WriteOperateLog m_oActiveUser, RIGHT_ReturnTicket, "退票票号：" & paszTicketID(i)
    Next i
    Set oDb = Nothing
    Exit Sub
Here:
    oDb.RollbackTrans
    Set oDb = Nothing
    err.Raise err.Number
End Sub

'/////////////////////////////
'内部用取消退票方法
Private Function SelfResumeReturnlTicket(paszTicketID() As String, poDb As RTConnection) As String()
    Dim szsql As String
    Dim oParam As New SystemParam
    Dim sgReturnCharge As Double, sgTicketPrice As Double
    Dim dtBusStartupDateTime As Date
    Dim rsTemp As Recordset
    Dim rsTemp1 As Recordset
    Dim rsTemp2 As Recordset
    Dim rsTemp3 As Recordset
    Dim lBusType As EBusType
    
    Dim i As Integer
    Dim nLen As Integer
    
    Dim szSeatNo As String
    Dim aszReturnSeatNo() As String
    
    nLen = 0
    nLen = ArrayLength(paszTicketID)
    
    If nLen <> 0 Then ReDim aszReturnSeatNo(1 To nLen)
    
    oParam.Init m_oActiveUser
    On Error GoTo ErrorHandle
    For i = 1 To nLen
         szsql = "SELECT * FROM Ticket_return_lst WHERE ticket_id ='" & paszTicketID(i) & "'"
         Set rsTemp = poDb.Execute(szsql)
        
         szsql = "SELECT status,bus_id,bus_date,seat_no FROM Ticket_sell_lst WHERE ticket_id='" & paszTicketID(i) & "'"
         Set rsTemp2 = poDb.Execute(szsql)
         If rsTemp2.RecordCount = 1 Then
              If rsTemp2!Status <> ST_TicketReturned + ST_TicketNormal Then
                  ShowError ERR_ServiceTicketNotReturn, , "票号为：" & paszTicketID(i)
              End If
         Else
              ShowError ERR_ServiceTicketNotExist, , "票号为：" & paszTicketID(i)
         End If
         If rsTemp.RecordCount = 1 Then
              '取得该票的座位类型和状态
              szsql = "SELECT seat_type_id ,status FROM Env_bus_seat_lst WHERE " _
                    & "bus_id ='" & rsTemp2!bus_id & "' AND " _
                    & "bus_date='" & ToDBDate(rsTemp2!bus_date) & "' AND " _
                    & "seat_no='" & rsTemp2!seat_no & "' "
                    
              Set rsTemp3 = poDb.Execute(szsql)
              szSeatNo = rsTemp2!seat_no
              
              If rsTemp2!Status <> ST_SeatCanSell Then
                szsql = "SELECT TOP 1 * FROM Env_bus_seat_lst WHERE " _
                    & "bus_id='" & rsTemp2!bus_id & "' AND " _
                    & "bus_date='" & ToDBDate(rsTemp2!bus_date) & "' AND " _
                    & "seat_type_id='" & rsTemp3!seat_type_id & "' AND " _
                    & "status=" & ST_SeatCanSell & ""
                    
                Set rsTemp3 = poDb.Execute(szsql)
                If rsTemp3.RecordCount = 0 Then ShowError ERR_ReturnTicketNotResume
                szSeatNo = rsTemp3!seat_no
              End If
              
              szsql = "UPDATE  Ticket_sell_lst SET status='" & ST_TicketNormal & "',seat_no='" & szSeatNo & "' WHERE ticket_id='" & paszTicketID(i) & "'"
              poDb.Execute szsql
              
              aszReturnSeatNo(i) = szSeatNo
              
'              szSql = "INSERT INTO resume_returnticket_lst(" _
'              & "ticket_id," _
'              & "operation_time," _
'              & "user_id," _
'              & "return_time," _
'              & "credence_sheet_id," _
'              & "return_charge," _
'              & "return_user_id," _
'              & "return_mode,"
'              szSql = szSql & "return_ratio)  VALUES('" _
'              & paszTicketID(i) & "','" _
'              & ToDBDateTime(oParam.NowDateTime) & "','" _
'              & m_oActiveUser.UserID & "','" _
'              & ToDBDateTime(rsTemp!operation_time) & "','" _
'              & rsTemp!credence_sheet_id & "'," _
'              & rsTemp!return_charge & ",'" _
'              & rsTemp!user_id & "'," _
'              & rsTemp!return_ticket_mode & "," _
'              & rsTemp!return_ratio & ")"
'              poDb.Execute szSql
              
              szsql = "DELETE Ticket_return_lst WHERE ticket_id='" & paszTicketID(i) & "'"
              poDb.Execute szsql
          End If
          
          '释放座位
          ResumeSeat paszTicketID(i), poDb
         
    Next i
              
   SelfResumeReturnlTicket = aszReturnSeatNo

   Set rsTemp = Nothing
   Set rsTemp1 = Nothing
   Set rsTemp2 = Nothing
   Set rsTemp3 = Nothing
   On Error GoTo 0
   Exit Function
ErrorHandle:
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    Set rsTemp2 = Nothing
    Set rsTemp3 = Nothing
    err.Raise err.Number
End Function
'//////////////////////////
'释放座位
Private Sub ResumeSeat(pszTicket As String, poDb As RTConnection)
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim rsTemp1 As Recordset
    Dim szBusID As String
    Dim dtDate As Date
    Dim szDesStationID As String
    Dim szSeatNo As String
    Dim szSeatTypeID As String
    
    On Error GoTo Here
    szsql = "SELECT * FROM Ticket_sell_lst WHERE ticket_id='" & pszTicket & "'"
    Set rsTemp = poDb.Execute(szsql)
    
    szBusID = FormatDbValue(rsTemp!bus_id)
    dtDate = FormatDbValue(rsTemp!bus_date)
    szDesStationID = FormatDbValue(rsTemp!des_station_id)
    szSeatNo = FormatDbValue(rsTemp!seat_no)
     
     '////////////////////////////
    '判断是否还有可售票数
    szsql = "SELECT * FROM Env_bus_info WHERE " _
    & " bus_id='" & szBusID & "'AND " _
    & " bus_date='" & ToDBDate(dtDate) & "'"
    Set rsTemp1 = poDb.Execute(szsql)
   
    
    If szSeatNo = cszStandTicketSeatNo Then '是站票
       If rsTemp1!sale_stand_seat_quantity >= 1 Then
            '将该环境车次的可售站票数减一
            szsql = "UPDATE Env_bus_info SET " _
            & " sale_stand_seat_quantity=sale_stand_seat_quantity-1 WHERE " _
            & " bus_id='" & szBusID & "' AND " _
            & " bus_date='" & ToDBDate(dtDate) & "'"
            poDb.Execute szsql
       Else
            GoTo Here
       End If
    Else '不是站票
        
        
        If rsTemp1!sale_seat_quantity >= 1 Then
        
            szSeatTypeID = FormatDbValue(rsTemp!seat_type_id)
        
            '将该环境车次的可售座位数减一和可售座位类型余留减一
            szsql = "UPDATE Env_bus_info SET sale_seat_quantity=sale_seat_quantity-1 , "
            Select Case szSeatTypeID
                Case cszSeatType
                    szsql = szsql & " seat_remain=seat_remain - 1  "
                Case cszBedType
                    szsql = szsql & " bed_remain=bed_remain - 1 "
                Case cszAdditionalType
                    szsql = szsql & " additional_remain=additional_remain - 1 "
                Case cszOtherType1
                    szsql = szsql & " other_remain_1 =other_remain_1 - 1 "
                Case cszOtherType2
                    szsql = szsql & " other_remain_2 =other_remain_2 - 1 "
            End Select
            szsql = szsql & " WHERE bus_id='" & szBusID & "' AND bus_date='" & ToDBDate(dtDate) & "'"

            poDb.Execute szsql
            
            '改变相应座位号的状态
            szsql = "UPDATE Env_bus_seat_lst SET " _
            & " status=" & ST_SeatSold & "," _
            & " ticket_no='" & pszTicket & "' WHERE " _
            & " bus_id='" & szBusID & "' AND " _
            & " bus_date='" & ToDBDate(dtDate) & "' AND " _
            & " seat_no='" & szSeatNo & "'"
            poDb.Execute szsql
        Else
            GoTo Here
        End If
    End If
    
    '改变该车次相应到站的可售票数
    szsql = "UPDATE Env_bus_station_lst SET " _
    & " sale_ticket_quantity=sale_ticket_quantity-1 WHERE " _
    & " bus_id='" & szBusID & "' AND " _
    & " bus_date='" & ToDBDate(dtDate) & "' AND " _
    & " station_id='" & szDesStationID & "' AND " _
    & " max_sale_quantity>0 "
    
    poDb.Execute szsql
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    On Error GoTo 0
    Exit Sub
Here:
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
'    MsgBox "已无可售票数" & vbCrLf & "取消退票或取消废票可能产生超载", vbInformation, "售票台:"
End Sub
'///////////////////////////////////
'内部用取消废票
Private Function SelfResumeCancelTicket(paszTicketID() As String, poDb As RTConnection) As String()
    Dim szsql As String
    Dim oParam As New SystemParam
    Dim rsTemp As Recordset
    Dim rsTemp1 As Recordset
    Dim rsTemp2 As Recordset
    Dim i As Integer
    Dim nLen As Integer
    Dim szSeatNo As String
    Dim aszCancelSeatNo() As String
    nLen = 0
    nLen = ArrayLength(paszTicketID)
    
    If nLen <> 0 Then ReDim aszCancelSeatNo(1 To nLen)
    oParam.Init m_oActiveUser
    
    On Error GoTo ErrorHandle
    For i = 1 To nLen
        szsql = "SELECT * FROM ticket_cancel_lst WHERE ticket_id='" & paszTicketID(i) & "'"
        Set rsTemp = poDb.Execute(szsql)
        szsql = "SELECT status,bus_date,bus_id,seat_no FROM Ticket_sell_lst WHERE ticket_id='" & paszTicketID(i) & "'"
        Set rsTemp1 = poDb.Execute(szsql)
        
        If rsTemp1.RecordCount = 1 Then
            If rsTemp1!Status <> ST_TicketCanceled + ST_TicketNormal Then
                ShowError ERR_ServiceTicketNotCancel, , "票号为：" & paszTicketID(i)
            End If
        Else
            If rsTemp1.RecordCount = 0 Then
                ShowError ERR_ServiceTicketNotExist, , "票号为：" & paszTicketID(i)
            End If
        End If
        If rsTemp.RecordCount = 1 Then
            szsql = "SELECT seat_type_id ,status FROM Env_bus_seat_lst WHERE " _
                    & "bus_id='" & rsTemp1!bus_id & "' AND " _
                    & "bus_date='" & ToDBDate(rsTemp1!bus_date) & "' AND " _
                    & "seat_no='" & rsTemp1!seat_no & "'"
            Set rsTemp2 = poDb.Execute(szsql)
            szSeatNo = rsTemp1!seat_no
            If rsTemp2!Status <> ST_SeatCanSell Then
                szsql = "SELECT TOP 1 seat_no FROM Env_bus_seat_lst WHERE  " _
                        & "bus_id='" & rsTemp1!bus_id & "' AND " _
                        & "bus_date='" & ToDBDate(rsTemp1!bus_date) & "' AND " _
                        & "seat_type_id='" & rsTemp2!seat_type_id & "' AND " _
                        & "status=" & ST_SeatCanSell & ""
                Set rsTemp2 = poDb.Execute(szsql)
                If rsTemp2.RecordCount = 0 Then ShowError ERR_CancelTicketNotResume
                szSeatNo = rsTemp2!seat_no
            End If
            szsql = "UPDATE Ticket_sell_lst SET " _
            & " status= " & ST_TicketNormal & ",seat_no='" & szSeatNo & "' WHERE " _
            & " ticket_id='" & paszTicketID(i) & "'"
            poDb.Execute szsql
            
            aszCancelSeatNo(i) = szSeatNo
            
'            szSql = "INSERT INTO resume_cancelticket_lst( " _
'            & " ticket_id," _
'            & " operation_time," _
'            & " user_id," _
'            & " cancel_time," _
'            & " cancel_user_id," _
'            & " cancel_mode ) VALUES('" _
'            & paszTicketID(i) & "','" _
'            & ToDBDateTime(oParam.NowDateTime) & "','" _
'            & m_oActiveUser.UserID & "','" _
'            & rsTemp!cancel_ticket_time & "','" _
'            & rsTemp!user_id & "'," _
'            & rsTemp!cancel_mode & ")"
'            poDb.Execute szSql
            
            szsql = "DELETE ticket_cancel_lst WHERE ticket_id='" & paszTicketID(i) & "'"
            poDb.Execute szsql
            
         End If
            '恢复座位
            ResumeSeat paszTicketID(i), poDb
       
    Next i
    
    SelfResumeCancelTicket = aszCancelSeatNo
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    Set rsTemp2 = Nothing
    On Error GoTo 0
    Exit Function
ErrorHandle:
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    Set rsTemp2 = Nothing
    err.Raise err.Number
End Function

'////////////////////////////////////////////////////////
'得到所有票种
Public Function GetAllTicketType(Optional nValid As Integer = TP_TicketTypeAll) As TTicketType()
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim szsql  As String
    Dim atTemp(1 To 6) As TTicketType
    On Error GoTo Here
    
    szsql = " SELECT * FROM Ticket_type_code "
    '默认为所有的票种
    If nValid <> TP_TicketTypeAll Then
        szsql = szsql & " WHERE ticket_type_valid = " & nValid
    End If
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    Set rsTemp = oDb.Execute(szsql)
     
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
Here:
    Set rsTemp = Nothing
    Set oDb = Nothing
    err.Raise err.Number
End Function
'//////////////////////////////////////
'得到退票废率
Public Function GetReturnRatioValue() As RETURNFEES()
    Dim rsTemp As Recordset
    Dim oParam As New SystemParam
    Dim adbValue() As Double
    Dim i As Integer
On Error GoTo Here
    oParam.Init m_oActiveUser
    
    GetReturnRatioValue = oParam.GetReturnFees
    Set oParam = Nothing
On Error GoTo 0
Exit Function
Here:
    Set oParam = Nothing
    err.Raise err.Number
    
End Function



'得到当前售票明细票价

Public Function GetSaleTicketPrice(BusID As String, BusStartDate As Date, EndStationID As String, _
                                TicketType() As ETicketType, SeatNo() As String, _
                                Optional nOrderSeatNo As Integer = 0 _
                               ) As Single()
Dim i As Integer
Dim nlen1 As Integer
Dim nLen2 As Integer    '当前票数
Dim nRealLen As Integer  '实际可售票数
Dim nLenTemp As Integer '临时票数
Dim nTicketType As ETicketType

Dim rsTemp1 As Recordset
Dim rsTemp2 As Recordset
Dim rsTemp3 As Recordset


Dim szSql1 As String
Dim szSql2 As String

Dim SeatPrice As TBusSeatPriceInfo
Dim sgTicketPrice() As Single
Dim szSeatType() As String

Dim oDb As New RTConnection

nlen1 = 0
nLen2 = 0
nRealLen = 0

On Error GoTo Here

oDb.ConnectionString = GetConnectionStr(cszSellTicket)


nLen2 = ArrayLength(TicketType)

ReDim sgTicketPrice(1 To nLen2)
ReDim szSeatType(1 To nLen2)


    '判断当前是否有订票
    If nOrderSeatNo = 0 Then
        szSql1 = "SELECT TOP " & nLen2 & " * FROM Env_bus_seat_lst e,Seat_type_code s WHERE " _
                     & "e.bus_date='" & ToDBDate(BusStartDate) & "'AND " _
                     & "e.bus_id ='" & BusID & "' AND" _
                     & "(e.status=" & ST_SeatCanSell & " OR e.status=" & ST_SeatBooked & ") AND " _
                     & " e.seat_type_id=s.seat_type_id ORDER BY seat_no Asc "
         Set rsTemp1 = oDb.Execute(szSql1)
         If rsTemp1.RecordCount = 0 Then ShowError ERR_BusReachMaxTicketCount, "SellTicketService.SelfSellTicket"
         On Error Resume Next
         For i = 1 To rsTemp1.RecordCount

            nTicketType = IIf(TicketType(i) = TP_FreeTicket, TP_FullPrice, TicketType(i))

            szSql1 = "SELECT ticket_price_total FROM Work_env_bus_price_lst WHERE " _
                        & "bus_date='" & ToDBDate(BusStartDate) & "' AND " _
                        & "station_id='" & EndStationID & "' AND " _
                        & "bus_id='" & BusID & "'AND " _
                        & "seat_type_id='" & rsTemp1!seat_type_id & "' AND " _
                        & "ticket_type=" & nTicketType & ""
            Set rsTemp2 = oDb.Execute(szSql1)
            If rsTemp2.RecordCount = 0 Then Exit Function
            sgTicketPrice(i) = rsTemp2!ticket_price_total
            szSeatType(i) = Trim(rsTemp1!seat_type_name)
            rsTemp1.MoveNext
         Next i
         nRealLen = rsTemp1.RecordCount
    Else
        szSql1 = "SELECT * FROM Env_bus_seat_lst e,Seat_type_code s WHERE " _
                    & "e.bus_date='" & ToDBDate(BusStartDate) & "' AND " _
                    & "e.bus_id='" & BusID & "' AND " _
                    & " e.seat_type_id=s.seat_type_id AND " _
                    & " ( "
        nLenTemp = nLen2 - nOrderSeatNo

        szSql2 = "SELECT TOP " & nLenTemp & " * FROM Env_bus_seat_lst e,Seat_type_code s WHERE " _
                        & "e.bus_date='" & ToDBDate(BusStartDate) & "' AND " _
                        & "e.bus_id='" & BusID & "' AND " _
                        & "e.seat_type_id=s.seat_type_id AND " _
                        & "(e.status=" & ST_SeatCanSell & " OR e.status=" & ST_SeatBooked & ") AND  " _
                        & "("
        For i = 1 To nOrderSeatNo - 1
            szSql1 = szSql1 & "e.seat_no='" & SeatNo(i) & "' OR "
            szSql2 = szSql2 & "e.seat_no<> '" & SeatNo(i) & "' AND "
        Next i
        szSql1 = szSql1 & "e.seat_no='" & SeatNo(nOrderSeatNo) & "' " & ") " & " ORDER BY seat_no ASC"
        szSql2 = szSql2 & "e.seat_no<> '" & SeatNo(nOrderSeatNo) & "' " & ") " & " ORDER BY seat_no ASC"

        Set rsTemp1 = oDb.Execute(szSql1)
        Set rsTemp2 = oDb.Execute(szSql2)

        If rsTemp1.RecordCount + rsTemp2.RecordCount = 0 Then ShowError ERR_BusReachMaxTicketCount, "SellTicketService.SelfSellTicket"

        If rsTemp1.RecordCount <> 0 Then
            For i = 1 To rsTemp1.RecordCount
                nTicketType = IIf(TicketType(i) = TP_FreeTicket, TP_FullPrice, TicketType(i))

                szSql1 = "SELECT ticket_price_total FROM Work_env_bus_price_lst WHERE " _
                                & "bus_id ='" & BusID & "' AND " _
                                & "bus_date='" & ToDBDate(BusStartDate) & "' AND " _
                                & "station_id='" & EndStationID & "' AND " _
                                & "seat_type_id='" & rsTemp1!seat_type_id & "' AND " _
                                & "ticket_type=" & nTicketType & ""
                Set rsTemp3 = oDb.Execute(szSql1)
                If rsTemp3.RecordCount = 0 Then Exit Function
                sgTicketPrice(i) = rsTemp3!ticket_price_total
                szSeatType(i) = Trim(rsTemp1!seat_type_name)
                rsTemp1.MoveNext
            Next i
        End If
        nRealLen = rsTemp1.RecordCount
        If rsTemp2.RecordCount <> 0 Then
            For i = 1 To rsTemp2.RecordCount

                nTicketType = IIf(TicketType(i + nRealLen) = TP_FreeTicket, TP_FullPrice, TicketType(i))

                szSql1 = "SELECT ticket_price_total FROM Work_env_bus_price_lst WHERE " _
                                & "bus_id ='" & BusID & "' AND " _
                                & "bus_date='" & ToDBDate(BusStartDate) & "' AND " _
                                & "station_id='" & EndStationID & "' AND " _
                                & "seat_type_id='" & rsTemp2!seat_type_id & "' AND " _
                                & "ticket_type=" & nTicketType & ""
                Set rsTemp3 = oDb.Execute(szSql1)
                sgTicketPrice(nRealLen + i) = rsTemp3!ticket_price_total
                szSeatType(nRealLen + i) = Trim(rsTemp2!seat_type_name)
                rsTemp2.MoveNext
            Next i

        End If
        nRealLen = nRealLen + rsTemp2.RecordCount
        If nRealLen < nLen2 Then
           For i = nRealLen + 1 To nLen2

                nTicketType = IIf(TicketType(i) = TP_FreeTicket, TP_FullPrice, TicketType(i))
                szSql1 = "SELECT TOP 1 * FROM Work_env_bus_price_lst e,Seat_type_code s WHERE " _
                                & "e.bus_date='" & ToDBDate(BusStartDate) & "' AND " _
                                & "e.station_id='" & EndStationID & "' AND " _
                                & "e.bus_id='" & BusID & "' AND " _
                                & "e.ticket_type=" & nTicketType & " AND " _
                                & "e.seat_type_id=s.seat_type_id"
                Set rsTemp3 = oDb.Execute(szSql1)

                sgTicketPrice(i) = rsTemp3!ticket_price_total
                szSeatType(i) = rsTemp3!seat_type_name
            Next i
        End If
    End If

 On Error GoTo 0
 On Error GoTo Here
'------------------------------------------------->>
'------------------------------------------------>>

GetSaleTicketPrice = sgTicketPrice

Set rsTemp1 = Nothing
Set rsTemp2 = Nothing
Set rsTemp3 = Nothing

Set oDb = Nothing
On Error GoTo 0
Exit Function
Here:
    Set rsTemp1 = Nothing
    Set rsTemp2 = Nothing
    Set rsTemp3 = Nothing


    Set oDb = Nothing
    err.Raise err.Number
End Function

' ////////////////////////////////
'得到有不同座位类型的车次
Public Function GetMultiSeatTypeBus() As TMultiSeatTypeBus
    Dim szsql As String
    Dim i As Integer
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    Dim tbTemp As TMultiSeatTypeBus
    Dim dtNowDate As Date
    
On Error GoTo Here
oDb.ConnectionString = GetConnectionStr(cszSellTicket)

dtNowDate = Date
szsql = "SELECT bus_id,bus_date,BusCount FROM (SELECT bus_id,bus_date,COUNT(bus_id) AS BusCount FROM " _
            & "( SELECT DISTINCT bus_id,bus_date,seat_type_id FROM Work_env_bus_price_lst WHERE bus_date>='" & ToDBDate(dtNowDate) & "') c GROUP BY bus_id,bus_date ) t " _
            & "WHERE BusCount>1 ORDER BY bus_id,bus_date "

'szSql = "SELECT DISTINCT bus_id,bus_date,seat_type_id FROM Work_env_bus_price_lst"
Set rsTemp = oDb.Execute(szsql)
If rsTemp.RecordCount = 0 Then Exit Function
ReDim tbTemp.adBusDate(1 To rsTemp.RecordCount)
ReDim tbTemp.aszBusID(1 To rsTemp.RecordCount)
For i = 1 To rsTemp.RecordCount
    tbTemp.adBusDate(i) = rsTemp!bus_date
    tbTemp.aszBusID(i) = rsTemp!bus_id
    rsTemp.MoveNext
Next i

GetMultiSeatTypeBus = tbTemp

Set rsTemp = Nothing
Set oDb = Nothing

On Error GoTo 0
Exit Function
Here:
    Set rsTemp = Nothing
Set oDb = Nothing

    err.Raise err.Number
End Function

'////////////////////////////////////////////////////////
'得到线路名称和发车时间
Public Function GetRouteAndTime(pdtBusDate As Date, pszBusID As String) As String()
Dim aszTemp() As String
Dim szsql As String
Dim oDb As New RTConnection
Dim rsTemp As Recordset
ReDim aszTemp(1 To 5)
 On Error GoTo Here
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    szsql = "SELECT r.route_name,e.bus_start_time,e.split_company_id,e.transport_company_id , c.transport_company_name,c.transport_company_short_name " _
            & " FROM Work_env_bus_info e,route_info r,company_info c " _
            & " WHERE e.transport_company_id = c.transport_company_id AND e.bus_date='" & ToDBDate(pdtBusDate) & "' AND " _
            & "e.bus_id='" & pszBusID & "' AND " _
            & "r.route_id=e.route_id"
    Set rsTemp = oDb.Execute(szsql)
    If rsTemp.RecordCount >= 1 Then
        aszTemp(1) = FormatDbValue(rsTemp!route_name)
        aszTemp(2) = Trim(CStr(rsTemp!bus_start_time))
        aszTemp(3) = FormatDbValue(rsTemp!split_company_id)
        aszTemp(5) = FormatDbValue(rsTemp!transport_company_name) '玉环打印参运公司全称
        aszTemp(4) = FormatDbValue(rsTemp!transport_company_short_name)
        
    End If
    GetRouteAndTime = aszTemp
On Error GoTo 0
    Set rsTemp = Nothing
    Set oDb = Nothing
Exit Function
Here:
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
Dim szsql As String
Dim i As Integer
Dim rsTemp As Recordset
Dim oDb As New RTConnection
Dim aszSeatType() As String
Dim nCount As Integer
oDb.ConnectionString = GetConnectionStr(cszSellTicket)
szsql = "SELECT * FROM Seat_type_code "
Set rsTemp = oDb.Execute(szsql)
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


'///////////////////////
'内部废票
'Public Sub SelfLocalCancel(TicketID() As String, Optional bFlag As Boolean = False)
'    Dim szSql As String
'    Dim oDb As New RTConnection
'    Dim oParam As New SystemParam
'    Dim i As Integer
'    Dim rsTemp As Recordset
'    oParam.Init m_oActiveUser
'    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
'On Error GoTo Here
'    oDb.BeginTrans
'    For i = 1 To ArrayLength(TicketID)
'         szSql = "SELECT service_unit_id,status FROM Agent_ticket_sale_lst WHERE ticket_id='" & TicketID(i) & "'"
'
'         Set rsTemp = oDb.Execute(szSql)
'         If rsTemp.RecordCount = 1 Then
'            If rsTemp!Status = ST_TicketSellChange Then ShowError ERR_TicketIsChangeTicket, , "票号为：" & TicketID(i)
'
'            szSql = "UPDATE Agent_ticket_sale_lst SET status=status| " & ST_TicketCanceled _
'                  & " WHERE ticket_id= '" & TicketID(i) & "'"
'            oDb.Execute szSql
'
'            szSql = "INSERT INTO Agent_ticket_cancel_lst (" _
'            & "ticket_id," _
'            & "cancel_time," _
'            & "user_id," _
'            & "cancel_mode) VALUES ('" _
'            & TicketID(i) & "','" _
'            & ToDBDateTime(oParam.NowDateTime) & "','" _
'            & m_oActiveUser.UserID & "',"
'            If bFlag Then
'                szSql = szSql & "" & TP_ForceCancel & ")"
'            Else
'                szSql = szSql & "" & TP_NormalCancel & ")"
'            End If
'
'            oDb.Execute (szSql)
'        Else
'            ShowError ERR_TicketNotExist, , "票号为：" & TicketID(i)
'
'        End If
'   Next i
'   oDb.CommitTrans
'   Set rsTemp = Nothing
'   Set oParam = Nothing
'   Set oDb = Nothing
'   Exit Sub
'Here:
'  oDb.RollbackTrans
'  err.Raise err.Number
'  Set rsTemp = Nothing
'  Set oParam = Nothing
'  Set oDb = Nothing
'End Sub


'内部强行退票
'Public Sub SelfLocalForceReturn(TicketID() As String, SheetID As String, OrgReturnCharge() As Double, pasgReturnRatio() As Double)
'  Dim szSql As String
'  Dim poDb As New RTConnection
'  Dim rsTemp As Recordset
'  Dim oParam As New SystemParam
'  Dim i As Integer
'
'  poDb.ConnectionString = GetConnectionStr(cszSellTicket)
'  oParam.Init m_oActiveUser
'
'On Error GoTo Here
'  poDb.BeginTrans
'    For i = 1 To ArrayLength(TicketID)
'
'        szSql = "SELECT * FROM Agent_ticket_sale_lst WHERE ticket_id='" & TicketID(i) & "'"
'            Set rsTemp = poDb.Execute(szSql)
'
'            szSql = "UPDATE Agent_ticket_sale_lst SET status=status|" _
'                & ST_TicketReturned & " WHERE " _
'                & " ticket_id='" & TicketID(i) & "'"
'            poDb.Execute (szSql)
'                szSql = "INSERT INTO Agent_ticket_return_lst(" _
'                & " ticket_id," & "bus_date," & "bus_id," _
'                & " service_unit_id," _
'                & "start_station_id," _
'                & "start_station_name," _
'                & " des_station_id," _
'                & " des_station_name," _
'                & " ticket_type," _
'                & " seat_number," _
'                & " ticket_price," _
'                & " operation_time," _
'                & " credence_sheet_id," _
'                & " return_charge," _
'                & " user_id,"
'                szSql = szSql & "return_mode,return_ratio)  VALUES ('" _
'                & TicketID(i) & "','" & ToDBDate(rsTemp!bus_date) & "','" & rsTemp!bus_id & "','" _
'                & rsTemp!service_unit_id & "','" _
'                & rsTemp!start_station_id & "','" _
'                & rsTemp!start_station_name & "','" _
'                & rsTemp!des_station_id & "','" _
'                & rsTemp!des_station_name & "'," _
'                & rsTemp!ticket_type & ",'" _
'                & rsTemp!seat_no & "'," _
'                & rsTemp!ticket_price & ",'" _
'                & ToDBDateTime(oParam.NowDateTime) & "','" _
'                & SheetID & "'," _
'                & OrgReturnCharge(i) & ",'" _
'                & m_oActiveUser.UserID & "'," _
'                & TP_NormalReturn & "," _
'                & pasgReturnRatio(i) & ")"
'
'                poDb.Execute (szSql)
'    Next i
'
'    poDb.CommitTrans
'    Set rsTemp = Nothing
'    Set poDb = Nothing
'    Set oParam = Nothing
'    Exit Sub
'Here:
'    poDb.RollbackTrans
'
'    Set rsTemp = Nothing
'    Set oParam = Nothing
'    Set poDb = Nothing
'
'
'End Sub

'Public Function ResumeReturnClient(TicketID() As String, SeatNo() As String) As String()
'    Dim oDb As New RTConnection
'    Dim szSql As String
'    Dim rsTemp As Recordset
'    Dim oParam As New SystemParam
'    Dim i As Integer
'    Dim szSeatNo As String
'
'    oParam.Init m_oActiveUser
'    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
'    On Error GoTo ErrHandle
'    oDb.BeginTrans
'      For i = 1 To ArrayLength(TicketID)
'        szSql = "SELECT * FROM Agent_ticket_return_lst WHERE ticket_id='" & TicketID(i) & "'"
'            Set rsTemp = oDb.Execute(szSql)
'
'            szSeatNo = SeatNo(i)
'            If rsTemp.RecordCount = 1 Then
'
'                szSql = "UPDATE Agent_ticket_sale_lst SET status=" _
'                & ST_TicketNormal & ",seat_no='" & szSeatNo & "' WHERE " _
'                & " ticket_id='" & TicketID(i) & "'"
'                oDb.Execute (szSql)
'                szSql = "INSERT INTO agent_resumereturn_lst(ticket_id,operation_time," _
'                & "user_id," _
'                & "bus_date," _
'                & "bus_id," _
'                & "service_unit_id," _
'                & "start_station_id," _
'                & "start_station_name," _
'                & "des_station_id," _
'                & "des_station_name," _
'                & "ticket_type," _
'                & "seat_number," _
'                & "ticket_price," _
'                & "return_time," _
'                & "credence_sheet_id," _
'                & "return_charge," _
'                & "return_mode ," _
'                & "return_ratio)"
'                szSql = szSql & "  VALUES('" & TicketID(i) & "','" _
'                & ToDBDateTime(oParam.NowDateTime) & "','" _
'                & m_oActiveUser.UserID & "','" _
'                & rsTemp!bus_date & "','" _
'                & rsTemp!bus_id & "','" _
'                & rsTemp!service_unit_id & "','" _
'                & rsTemp!start_station_id & "','" _
'                & rsTemp!start_station_name & "','" _
'                & rsTemp!des_station_id & "','" _
'                & rsTemp!des_station_name & "'," _
'                & rsTemp!ticket_type & ",'" _
'                & szSeatNo & "'," _
'                & rsTemp!ticket_price & ",'" _
'                & rsTemp!operation_time & "','" _
'                & rsTemp!credence_sheet_id & "'," _
'                & rsTemp!return_charge & "," _
'                & rsTemp!return_mode & "," _
'                & rsTemp!return_ratio & ")"
'                oDb.Execute (szSql)
'
'                szSql = "DELETE  FROM  Agent_ticket_return_lst WHERE ticket_id='" & TicketID(i) & "'"
'                oDb.Execute (szSql)
'          End If
'    Next i
'    oDb.CommitTrans
'    Set oParam = Nothing
'    Set rsTemp = Nothing
'    Set oDb = Nothing
'Exit Function
'ErrHandle:
'   oDb.RollbackTrans
'   Set oParam = Nothing
'   Set oDb = Nothing
'   Set rsTemp = Nothing
'   err.Raise err.Number
'End Function

'Public Function ResumeCancelClient(TicketID() As String, SeatNo() As String) As String
'Dim oDb As New RTConnection
'Dim oParam As New SystemParam
'Dim rsTemp As Recordset
'Dim szSql As String
'Dim szSeatNo As String
'Dim i As Integer
'
'oParam.Init m_oActiveUser
'oDb.ConnectionString = GetConnectionStr(cszSellTicket)
'On Error GoTo ErrHandle
'    oDb.BeginTrans
'    For i = 1 To ArrayLength(TicketID)
'         szSql = "SELECT * FROM Agent_ticket_cancel_lst WHERE ticket_id='" & TicketID(i) & "'"
'         Set rsTemp = oDb.Execute(szSql)
'
'         If rsTemp.RecordCount = 1 Then
'                szSeatNo = SeatNo(i)
'                szSql = "UPDATE Agent_ticket_sale_lst SET status=" & ST_TicketNormal & ",seat_no='" & szSeatNo & "'" _
'                      & " WHERE ticket_id= '" & TicketID(i) & "'"
'                oDb.Execute szSql
'
'                szSql = "INSERT INTO agent_resumecancel_lst (" _
'                & "ticket_id," _
'                & "operation_time," _
'                & "user_id," _
'                & "cancel_time," _
'                & "cancel_mode," _
'                & "cancel_user_id) VALUES ('" _
'                & TicketID(i) & "','" _
'                & ToDBDateTime(oParam.NowDateTime) & "','" _
'                & m_oActiveUser.UserID & "','" _
'                & rsTemp!cancel_time & "'," _
'                & rsTemp!cancel_mode & ",'" _
'                & rsTemp!user_id & "')"
'                oDb.Execute szSql
'
'                szSql = "DELETE Agent_ticket_cancel_lst WHERE ticket_id='" & TicketID(i) & "'"
'                oDb.Execute szSql
'        End If
'
'    Next i
'    oDb.CommitTrans
'    oDb.RollbackTrans
'    Set oParam = Nothing
'    Set rsTemp = Nothing
'    Set oDb = Nothing
'Exit Function
'ErrHandle:
'    oDb.RollbackTrans
'    Set oParam = Nothing
'    Set rsTemp = Nothing
'    Set oDb = Nothing
'    err.Raise err.Number
'End Function

'取消保险费
Public Sub CancelInsurance(paszTicketID() As String)
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim i As Integer
    Dim nLen As Integer
    Dim szWhere As String
    nLen = ArrayLength(paszTicketID)
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_CancelTicket
    If nLen = 0 Then Exit Sub
    
    szWhere = " "
    For i = 1 To nLen - 1
        szWhere = szWhere & TransFieldValueToString(paszTicketID(i)) & ","
    Next i
    szWhere = szWhere & TransFieldValueToString(paszTicketID(i))
    
    
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    szsql = " UPDATE ticket_sell_lst SET insurance = 0 WHERE  ticket_id IN ( " & szWhere & " ) "
    oDb.Execute szsql
    
    
    On Error GoTo 0
    '写日志
    '不过权限未定义 ,有空需要定义一下
    WriteOperateLog m_oActiveUser, RIGHT_CancelTicket, "取消保险费：" & szWhere
    
    Exit Sub
Here:
    err.Raise err.Number
End Sub


