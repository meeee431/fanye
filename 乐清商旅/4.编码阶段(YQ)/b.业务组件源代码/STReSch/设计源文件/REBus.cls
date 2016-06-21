VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**********************************************************
'* Source File Name:REBus.cls
'* Project Name:STReSch.vbp
'* Engineer:
'* Data Generated:
'* Last Revision Date:
'* Brief Description:临时调度车次对象
'* Relational Document:
'**********************************************************

Option Explicit
Public m_bStateOpen As Boolean
Public m_bState As Boolean

'出错
Public Enum EErrREBus
    ERR_REBusNotAddStatus = ERR_REBus + ERR_NotAddObj '18816环境车次不在新增状态
    ERR_REBusAdd = ERR_REBus + ERR_AddDuplicate '18801车次已存在，不能新增
    ERR_REBusNotExist = ERR_REBus + 21 '18821车次不存在
    ERR_CheckGateNotExist = ERR_REBus + ERR_EditChildExist '18805车次已有售票、检票或调度，不能修改
    ERR_REBusHaveSellTicketNotDelete = ERR_REBus + 22 '18822车次有售出票不能删除
    ERR_REBusNotStationPrice = ERR_REBus + 23 '18823车次无该站点票价
    ERR_REBusNoBeMergedBus = ERR_REBus + 24 '18824我并未并1 班其他车次，我是好的
    ERR_REBusMergedMe = ERR_REBus + 25 '18825我未被其他车次并班，我任是好的
    ERR_REBusNoSellTicket = ERR_REBus + 26 '18826本车次还未售票
    ERR_REBusNoStation = ERR_REBus + 27 '18827车次无运行站点
    ERR_REBusNotSeat = ERR_REBus + 28 '18828车次无座位
    ERR_REBusNoStartCheck = ERR_REBus + 29 '18829车次还无检票
    ERR_REBusSeatNotReserve = ERR_REBus + 30 '18830车次座位不能预留
    ERR_REBusSeatNoReserved = ERR_REBus + 31 '18831车次座位不能取消预留，因为未预留
    ERR_REBusSeatNotDelete = ERR_REBus + 32 '18832车次座位不能删除，因为座位预留或已售出
    ERR_REBusSeatExistNotAddNew = ERR_REBus + 33 '18833车次座位已存在，不能新增
    ERR_REBusVehicelNotExist = ERR_REBus + 34 '18834车辆不存在
    ERR_REBusStopBusDateNoOk = ERR_REBus + 35 '18835开始停班或停班结束日期不在预售日期之内
    ERR_REBusScrollBusNotReplace = ERR_REBus + 36 '18836流水车次不能顶班
    ERR_REBusReplaceVehicleSeatNotOk = ERR_REBus + 37 '18837顶班车辆的座位数不足
    ERR_REBusResumeBusDateNoOk = ERR_REBus + 38 '18838车次复班日期不在预售日期之内
    ERR_REBusScrollBusNotMerge = ERR_REBus + 39 '18839流水车次不能并班
    ERR_REBusNotNormalNoMerge = ERR_REBus + 40 '18840车次不是正常车次或以检票车次不能并班
    ERR_REBusMerged = ERR_REBus + 41 '18841车次已并班不能再进行并班操作
    ERR_REBusCheckBusNotMerge = ERR_REBus + 42 '18842已检票车次不能并班
    ERR_REBusNOTickePrice = ERR_REBus + 43 '18843车次无该车型的票价表
    ERR_REBusRegularBusNotSerial = ERR_REBus + 44 '18844流水车次没有车次序号
    ERR_REBusPropertyNotEdit = ERR_REBus + 45 '18845车次的该属性不能修改
    ERR_REBusAddNewNotOk = ERR_REBus + 46 '18846新增车次的代码和日期不合理
    
    '-----------------------
    ERR_REBusMergedOtherBus = ERR_REBus + 47 '18847车次并有其它车次,不能顶班
    ERR_REBusBeMerged = ERR_REBus + 48 '18848车次被并班停班,不能顶班
    ERR_REBusIsStopped = ERR_REBus + 49 '18849车次已被停班,,不能顶班
    ERR_REBusNotExecuteProject = ERR_REBus + 50 '18850当前无执行计划,不能顶班
    ERR_REBusNewVehicleNowRoutePrice = ERR_REBus + 51 '18851新车型无线路票价表或线路票价表不全,不能顶班
    ERR_REBusHaveCheckManner = ERR_REBus + 52 '18852车次已有检票行为,不能顶班
    ERR_MergeBusNoStop = ERR_REBus + 53 '并班车次不能停班
    ERR_BusStopNoStop = ERR_REBus + 54 '车次已停班不能再停
    ERR_BusCheckNoStop = ERR_REBus + 55 '车次已检票不能停班
    ERR_REBusNotStopNotResume = ERR_REBus + 56 '车次不是停班不能复班
    ERR_MergeNoStop = ERR_REBus + 57 '并班车次不能停班
    ERR_REBusScrollBusNotStop = ERR_REBus + 58 '流水车次不能停班
    ERR_REBusMergedDate = ERR_REBus + 59 '并班车次的日期不正确
    ERR_REBusSlitpNotInfo = ERR_REBus + 60 '拆分车次无信息
    ERR_BusBESlitpNotSeatNo = ERR_REBus + 61 '目标车次无座位
    ERR_NoSlitpREBus = ERR_REBus + 62   '车次没被拆分
    ERR_REYesSlitpREBus = ERR_REBus + 63 '车次已被拆分
    ERR_NOReBus = ERR_REBus + 64 '车次不在不存在
    ERR_ReNoNormal = ERR_REBus + 65 '车次可能已售票、停班、....
    ERR_REEXitStation = ERR_REBus + 66 '18866站点已存在,不能新增
   ' ERR_REBusRepalace = ERR_REBus + 68 '18868车次已顶班停班
    ERR_NOTICKET = ERR_REBus + 69  '拆分车次无售票
    ERR_SeSeatErr = ERR_REBus + 70 '拆分车次座位分有误
    ERR_NotPriceSeatType = ERR_REBus + 71 '无对应座位类型票价! 请检查该车次车辆是否有该座位类型,然后到<票价管理>新增座位类型票价.
    ERR_NotPriceItemParam = ERR_REBus + 72 '无票价参数项
    ERR_NotInsertSationID = ERR_REBus + 73 '该站点不是您在环境中插入的站点
    ERR_REBusHaveStop = ERR_REBus + 74 '原车次或目标车次已停班或已检票
    ERR_NoExePriceTable = ERR_REBus + 75 '无执行票价表
    ERR_NoRoute = ERR_REBus + 76 '无可用线路
    ERR_NoBusTypeInfo = ERR_REBus + 77 '无车次类型信息
    ERR_NoOwnerInfo = ERR_REBus + 78 '无车主信息
    ERR_NoSeatType = ERR_REBus + 79  ' 座位类型匹配有误原因：新车辆同旧车辆座位类型分配不一致而您要却要座位类型匹配
    ERR_NoCompany = ERR_REBus + 80 '没有参运公司。
    ERR_NoSlitpCompany = ERR_REBus + 81 '没有拆帐参运公司。
    ERR_NoCheckgate = ERR_REBus + 82 ' 18882检票口不存在
    
    ERR_NoSellStation = ERR_REBus + 83 '该售票站不存在****无用
    ERR_SellTicketsOut = ERR_REBus + 83 '已售票数超过允许售票数量
    ERR_Allotexistent = ERR_REBus + 84 '配载信息存在
    ERR_CheckGateNotMaskStation = ERR_REBus + 85 '检票口不存在，或与所属车站不符合
    ERR_BatchCloneEnvBusTooLong = ERR_REBus + 86 '环境车次批量复制天数超过限制
    ERR_UnSlipCoutNotEqualSlipCount = ERR_REBus + 87 '取消拆分的票数与被拆分的票数不符
 End Enum

'权限
Public Enum ERightREBus
    
    RIGHT_REBusManagement = ERR_REBus + cnMidRightBegin + cnMidRightStep * 1 '18907临时调度
    RIGHT_REBusDelete = ERR_REBus + cnMidRightBegin + cnMidRightStep * 2 ' 18913删除车次
    RIGHT_REBusAdd = ERR_REBus + cnMidRightBegin + cnMidRightStep * 3 '18919新增车次
    RIGHT_REBusUpdate = ERR_REBus + cnMidRightBegin + cnMidRightStep * 4 '18925修改车次
    RIGHT_ReBusSeatManagement = ERR_REBus + cnMidRightBegin + cnMidRightStep * 6 '19931座位管理
    RIGHT_REBusStopBus = ERR_REBus + cnMidRightBegin + cnMidRightStep * 11 '18937车次停班
    RIGHT_REBusResumeStopBus = ERR_REBus + cnMidRightBegin + cnMidRightStep * 12 '18943车次复班
    RIGHT_REBusMergeBus = ERR_REBus + cnMidRightBegin + cnMidRightStep * 13 '18949车次并班
    RIGHT_REBusReplaceBus = ERR_REBus + cnMidRightBegin + cnMidRightStep * 14 '18955车次顶班
    RIGHT_REBusModifyBusStation = ERR_REBus + cnMidRightBegin + cnMidRightStep * 15 '18961编辑车次站点
    RIGHT_REBusModifyBusTicketPrice = ERR_REBus + cnMidRightBegin + cnMidRightStep * 16 '18967编辑车次站点票价
    RIGHT_REBusUnMergeBus = ERR_REBus + cnMidRightBegin + cnMidRightStep * 17 '18973取消车次并班
    RIGHT_ReBusSlitp = ERR_REBus + cnMidRightBegin + cnMidRightStep * 18 '车次拆分
    RIGHT_REBusUnSlitp = ERR_REBus + cnMidRightBegin + cnMidRightStep * 19 '取消车次拆分
    RIGHT_REBusAddStation = ERR_REBus + cnMidRightBegin + cnMidRightStep * 20 '环境插入站点
    RIGHT_REBusAllot = ERR_REBus + cnMidRightBegin + cnMidRightStep * 21 '环境车次配载
    
    
    
    RIGHT_REBusSellStation = ERR_REBus + cnMidRightBegin + cnMidRightStep * 22 '环境车次售票点管理
    
    
    
End Enum

Private m_oActiveUser As ActiveUser 'P1所有的类都有的私有类变量
Private m_nObjectStatus As EObjectStatus 'P2所有的实体类都有的私有类变量
Private m_szBusID As String 'P3车次代码
Private m_dtRunDate As Date 'P4车次运行日期
Private m_szCheckSheet As String 'P5车次路单代码G
Private m_szEndStationName As String 'P6车次的终点站G
Private m_eCheckStatus As EREBusCheckStatus  'P7车次检票状态G
Private m_szBusMergedMe As String 'P8我被谁并了，我坏了G
Private m_tBeMergedBus As TMergeInfo  'P9被我并的车次，它坏了G
Private m_szVehicle As String 'P11车次的运行车辆G
Private m_szOwner As String 'P12车次运行车辆的车主G
Private m_szEndStation As String 'P13车次的终点站代码G
Private m_szRoute As String 'P14车次的运行线路G
Private m_eBusStatus As EREBusStatus  'P15车次的运行状态G
Private m_bAllRefundment As Boolean 'P16是否全额退票
Private m_bDelayStatus As Boolean '是否晚点
Private m_eBeingStatus As EREBusBeingStatus 'P17车次报班状态
'Private m_eBusType As EBusType  'P18车次运行类型G
Private m_nBusType As Integer
Private m_szBusTypeName As String
Private m_dtStartupTime As Date 'P19车次发车时间
Private m_szCheckGate As String 'P22车次检票口
Private m_szVehicleModelID As String 'P23车型代码G
Private m_szVehicleModelName As String 'P24车型名称G
Private m_bRefreshBusMerge As Boolean '是否刷新车次
Private m_bRefreshBusCheck As Boolean '是否刷新车次
Private m_tMergeBus As TMergeInfo '车次并班
Private m_nTotalSeat As Integer 'P25总座位数G
Private m_nSaleSeat As Integer 'P26可售座位数G
Private m_nTotalStandSeat As Integer 'P27总的站票数G
Private m_nSaleStandSeat  As Integer 'P28可售站票数G
Private m_szSellStationID As String
Private m_szCheckGateID As String
Private m_szCanSaleQuantity As String

 Private m_nAdditionalSeatCount As Integer '普通
 Private m_nSeatRemain As Integer '卧铺
 Private m_nBedSeatCount As Integer '加座



Private m_dtStartCheckTime As Date  'P29开检时间
Private m_dtStopCheckTime As Date   'P30停检时间
Private m_szVehicleTag As String 'P33车辆牌照
Private m_szRouteName As String 'P34线路名称
Private m_szOwnerName As String 'P32车主名称
Private m_szCompany As String 'P10车次车辆的参运公司G
Private m_szCompanyName As String 'P31参运公司名称
Private m_szSplitCompanyID As String 'P33拆帐公司代码
Private m_szSplitCompanyName As String 'P34拆帐公司名称
Private m_nStartSeatNo As Integer  'P34拆帐公司名称

Private m_szBusCopy As String '用来得到票价的车次（顶班用）



Private m_nBaggageNumber As Integer  '行包件数
Private m_dbFactWeight As Double '行包实重
Private m_dbCalculateWeight As Double '行包计重
Private m_nOverWeightNumber As Integer '超重件数
Private m_dbLuggageTotalPrice As Double '行包总额

'互联售票
Private m_nInternetStatus As Integer '
'是否有行包
Private m_bHaveLugge As Boolean
'滚动车次循环检票时间
Private m_nScrollBusCheckTime As Integer
Private m_noldScrollBusCheckTime As Integer

'为了写日志用
Private m_szOldBusID As String 'P3车次代码
Private m_szOldCompanyName As String 'P31参运公司名称
Private m_szOldOwnerName As String 'P32参运公司名称
Private m_szOldVehicleTag As String 'P33车辆牌照
Private m_szOldRoute As String 'P14车次的运行线路G
Private m_szOldRouteName As String 'P34线路名称
Private m_eOldBusType As Integer  'P18车次运行类型G
Private m_szOldBusTypeName As String
Private m_dtOldStartupTime As Date 'P19车次发车时间
Private m_szOldCheckGate As String 'P22车次检票口
Private m_szOldVehicleModelID As String 'P23车型代码G
Private m_szOldVehicleModelName As String 'P24车型名称G
Private m_nOldTotalSeat As Integer 'P25总座位数G
Private m_nOldSaleSeat As Integer 'P26可售座位数G
Private m_nOldTotalStandSeat As Integer 'P27总的站票数G
Private m_nOldSaleStandSeat  As Integer 'P28可售站票数G
Private m_szOldEndStation As String 'P13车次的终点站代码G
Private m_szOldEndStationName As String 'P6车次的终点站G
Private m_eOldBeingStatus As EREBusBeingStatus 'P17车次报班状态
Private m_szOldCompany As String 'P10车次车辆的参运公司G
Private m_szOldVehicle As String 'P11车次的运行车辆G
Private m_szOldOwner As String 'P12车次运行车辆的车主G
Private m_eOldBusStatus As EREBusStatus  'P15车次的运行状态G
Private m_bOldAllRefundment As Boolean 'P16是否全额退票
Private m_bOldDelayStatus As Boolean '是否晚点

Private m_szStatus As String
Private m_szSellQuantity As String



Private m_dtOldRunDate As Date 'P4车次运行日期
Private m_szOldCheckSheet As String 'P5车次路单代码G
Private m_eOldCheckStatus As EREBusCheckStatus  'P7车次检票状态G
Private m_szOldBusMergedMe As String 'P8我被谁并了，我坏了G
Private m_dtOldStartCheckTime As Date  'P29开检时间
Private m_dtOldStopCheckTime As Date   'P30停检时间
Private m_nOldBusType As Integer

Private m_szOldSplitCompanyID As String 'P33拆帐公司代码
Private m_szOldSplitCompanyName As String 'P34拆帐公司名称


Private m_nOldInternetStatus As Integer


'消息支持
'---------------------------
#If SUPPORT_MSG = 1 Then
Private m_oSndMsg As New CSendMsg
Private m_bOffTimeChanged As Boolean
Private m_bCheckGateChanged As Boolean
#End If

'F1获得车次的所有售出票
Public Function GetAllTicket() As String()
Dim oDb As New RTConnection
Dim szSql As String
Dim i As Integer
Dim szaTicket() As String
Dim rsTemp As Recordset
Dim nTemp As Integer
nTemp = ST_TicketCanceled Or ST_TicketChanged Or ST_TicketReturned
AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    szSql = "SELECT * FROM Ticket_sell_lst WHERE bus_id='" & m_szBusID & "' AND " _
    & " bus_date='" & ToDBDate(m_dtRunDate) & "' AND NOT ((status & " & Str(nTemp) & ")<>0)"
    Set rsTemp = oDb.Execute(szSql)
    '查询该车次售出的有效票
    If rsTemp.RecordCount = 0 Then ShowError ERR_REBusNoSellTicket
    ReDim szaTicket(1 To rsTemp.RecordCount) As String
    For i = 1 To rsTemp.RecordCount
        szaTicket(i) = FormatDbValue(rsTemp!ticket_id)
        rsTemp.MoveNext
    Next
    GetAllTicket = szaTicket
    Set oDb = Nothing
End Function
''F2查询车次的所有过站和站点的属性，不包括站点的里程数
'Public Function GetBusStationInfo() As TBusStationSellInfo()
'Dim odb As New rtConnection
'Dim szsql As String
'Dim i As Integer
'Dim BSInfo() As TBusStationSellInfo
'Dim rsTemp As Recordset
'AssertObjIsValid
'    odb.ConnectionString = GetConnectionStr(cszRunEnv)
'    If DateDiff("d", m_dtRunDate, Now) <= 0 Then '如果查询的是环境中的车次则使用工作表
'       szsql = "SELECT * FROM Work_env_bus_station_lst  WHERE bus_id='" & m_szBusID & "'AND " _
'       & " bus_date='" & ToDBDate(m_dtRunDate) & "' ORDER BY station_serial_no"
'    Else
'        szsql = "SELECT * FROM Env_bus_station_lst WHERE bus_id='" & m_szBusID & "' AND " _
'       & " bus_date='" & ToDBDate(m_dtRunDate) & "' ORDER BY station_serial_no"
'    End If
'    Set rsTemp = odb.Execute(szsql)
'    '-------------------------------
'    '车次无运行站点
'    If rsTemp.RecordCount = 0 Then ShowError ERR_BusNoStation
'    ReDim BSInfo(1 To rsTemp.RecordCount) As TBusStationSellInfo
'    For i = 1 To rsTemp.RecordCount
'        BSInfo(i).szStationID =  FormatDbValue(rstemp!station_id
'        BSInfo(i).sgHalfPrice =  FormatDbValue(rstemp!half_price
'        BSInfo(i).sgFullPrice =  FormatDbValue(rstemp!full_price
'        BSInfo(i).nMileage =  FormatDbValue(rstemp!mileage
'        BSInfo(i).nLimitedSellTime =  FormatDbValue(rstemp!stop_sale_time
'        BSInfo(i).nLimitedSellCount =  FormatDbValue(rstemp!max_sale_quantity
'        BSInfo(i).sgPreferentialPrice1 =  FormatDbValue(rstemp!preferential_ticket1
'        BSInfo(i).sgPreferentialPrice2 =  FormatDbValue(rstemp!preferential_ticket2
'        BSInfo(i).sgPreferentialPrice3 =  FormatDbValue(rstemp!preferential_ticket3
'        rsTemp.MoveNext
'    Next
'    GetBusStationInfo = BSInfo
'    Set odb = Nothing
'End Function
'**********************************
' 查询车次的所有过站和站点的属性
'*******************************
'Public Function GetBusStationInfo() As TBusStationSellInfo()
'Dim oDb As New RTConnection
'Dim szSql As String
'Dim i As Integer
'Dim BSInfo() As TBusStationSellInfo
'Dim rsTemp As Recordset
'AssertObjIsValid
'    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
'    If DateDiff("d", m_dtRunDate, Now) <= 0 Then '如果查询的是环境中的车次则使用工作表
'        szSql = "SELECT tbe.*,tbs.station_name,sti.seat_type_name FROM " _
'                & "Work_env_bus_station_lst tbe,station_info tbs ,Seat_type_code sti WHERE " _
'                & "tbe.bus_id='" & m_szBusID & "' AND " _
'                & "tbe.bus_date='" & ToDBDate(m_dtRunDate) & "' AND " _
'                & "tbe.station_id=tbs.station_id  AND tbe.seat_type_id =sti.seat_type_id ORDER BY tbe.station_serial_no,sti.seat_type_id"
'    Else
'        szSql = "SELECT tbe.*,tbs.station_name,sti.seat_type_name FROM " _
'                & "Env_bus_station_lst tbe,station_info tbs ,Seat_type_code sti WHERE " _
'                & "tbe.bus_id='" & m_szBusID & "' AND " _
'                & "tbe.bus_date='" & ToDBDate(m_dtRunDate) & "' AND " _
'                & "tbe.station_id=tbs.station_id  AND tbe.seat_type_id =sti.seat_type_id ORDER BY tbe.station_serial_no,sti.seat_type_id"
'    End If
'    Set rsTemp = oDb.Execute(szSql)
'    '-------------------------------
'    '车次无运行站点
'    If rsTemp.RecordCount = 0 Then ShowError ERR_BusNoStation
'    ReDim BSInfo(1 To rsTemp.RecordCount) As TBusStationSellInfo
'    For i = 1 To rsTemp.RecordCount
'        BSInfo(i).szStationID = FormatDbValue(rsTemp!station_id)
'        BSInfo(i).sgHalfPrice = FormatDbValue(rsTemp!half_price)
'        BSInfo(i).sgFullPrice = FormatDbValue(rsTemp!full_price)
'        BSInfo(i).nMileage = FormatDbValue(rsTemp!mileage)
'        BSInfo(i).sgLimitedSellTime = FormatDbValue(rsTemp!stop_sale_time)
'        BSInfo(i).nLimitedSellCount = FormatDbValue(rsTemp!max_sale_quantity)
'        BSInfo(i).szStationName = FormatDbValue(rsTemp!station_name)
'        BSInfo(i).sgPreferentialPrice1 = FormatDbValue(rsTemp!preferential_ticket1)
'        BSInfo(i).sgPreferentialPrice2 = FormatDbValue(rsTemp!preferential_ticket2)
'        BSInfo(i).sgPreferentialPrice3 = FormatDbValue(rsTemp!preferential_ticket3)
'        BSInfo(i).szSeatTypeID = FormatDbValue(rsTemp!seat_type_id)
'        BSInfo(i).szSeatTypeName = FormatDbValue(rsTemp!seat_type_name)
'        rsTemp.MoveNext
'    Next
'    GetBusStationInfo = BSInfo
'    Set rsTemp = Nothing
'    Set oDb = Nothing
'End Function

' *******************************************************************
' *   Brief Description:  查询车次的所有过站和站点，主要是价格      *
' *   Engineer: fl                                                  *
' *   Date Generated: 2002/11/12                                    *
' *   Last Revision Date:                                           *
' *******************************************************************
Public Function GetBusStationPriceInfo() As TREBusStationInfo()
Dim oDb As New RTConnection
Dim szSql As String
Dim i As Integer
Dim BSInfo() As TREBusStationInfo
               
Dim rsTemp As Recordset
AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    If DateDiff("d", m_dtRunDate, Now) <= 0 Then '如果查询的是环境中的车次则使用工作表
        szSql = "SELECT tbe.*, tbs.sell_station_name, tbt.seat_type_name FROM " _
                & " Work_env_bus_station_lst tbe,sell_station_info tbs ,Seat_type_code tbt WHERE " _
                & " tbs.sell_station_id= tbe.sell_station_id AND " _
                & " tbe.bus_id = '" & m_szBusID & "' AND " _
                & " tbt.seat_type_id = tbe.seat_type_id " _
                & " tbe.bus_date = '" & ToDBDate(m_dtRunDate) & "'" _
                & " ORDER BY tbe.station_serial_no"
    Else
        szSql = "SELECT tbe.* ,tbs.sell_station_name,tbt.seat_type_name FROM " _
                & " Env_bus_station_lst tbe , sell_station_info tbs , Seat_type_code tbt WHERE " _
                & " tbs.sell_station_id = tbe.sell_station_id AND " _
                & " tbe.bus_id = '" & m_szBusID & "' AND " _
                & " tbt.seat_type_id = tbe.seat_type_id AND " _
                & " tbe.bus_date = '" & ToDBDate(m_dtRunDate) & "'" _
                & " ORDER BY tbe.station_serial_no "
'        szSql = "SELECT tbe.*,tbs.sell_station_name, FROM " _
'                & "Env_bus_station_lst tbe,station_info tbs ,Seat_type_code sti WHERE " _
'                & "tbe.bus_id='" & m_szBusID & "' AND " _
'                & "tbe.bus_date='" & ToDBDate(m_dtRunDate) & "' AND " _
'                & "tbe.station_id=tbs.station_id  AND tbe.seat_type_id =sti.seat_type_id ORDER BY tbe.station_serial_no,sti.seat_type_id"
    End If
    Set rsTemp = oDb.Execute(szSql)

    If rsTemp.RecordCount = 0 Then ShowError ERR_BusNoStation
    ReDim BSInfo(1 To rsTemp.RecordCount) 'As TBusStationSellInfo
    For i = 1 To rsTemp.RecordCount
        BSInfo(i).szSellStationID = FormatDbValue(rsTemp!sell_station_id)
        BSInfo(i).szSellStationName = FormatDbValue(rsTemp!sell_station_name)
        
        BSInfo(i).szStationID = FormatDbValue(rsTemp!station_id)
        BSInfo(i).sgHalfPrice = FormatDbValue(rsTemp!half_price)
        BSInfo(i).sgFullPrice = FormatDbValue(rsTemp!full_price)
        BSInfo(i).nMileage = FormatDbValue(rsTemp!mileage)
'        BSInfo(i).sgLimitedSellTime = FormatDbValue(rsTemp!stop_sale_time)
'        BSInfo(i).nLimitedSellCount = FormatDbValue(rsTemp!max_sale_quantity)
'        BSInfo(i).szStationName = FormatDbValue(rsTemp!station_name)
        BSInfo(i).sgPreferentialPrice1 = FormatDbValue(rsTemp!preferential_ticket1)
        BSInfo(i).sgPreferentialPrice2 = FormatDbValue(rsTemp!preferential_ticket2)
        BSInfo(i).sgPreferentialPrice3 = FormatDbValue(rsTemp!preferential_ticket3)
        BSInfo(i).szSeatTypeID = FormatDbValue(rsTemp!seat_type_id)
        BSInfo(i).szSeatTypeName = FormatDbValue(rsTemp!seat_type_name)
        rsTemp.MoveNext
    Next
    GetBusStationPriceInfo = BSInfo
    Set rsTemp = Nothing
    Set oDb = Nothing


End Function

' *******************************************************************
' *   Brief Description:                                            *
' *   Engineer: fl                                                  *
' *   Date Generated: 2002/11/12                                    *
' *   Last Revision Date:                                           *
' *******************************************************************


Public Function GetBusStation() As TREBusStationInfo()  'TBusStationSellInfo()
Dim oDb As New RTConnection
Dim szSql As String
Dim i As Integer
Dim BSInfo() As TREBusStationInfo
Dim szTableName As String '表名

Dim rsTemp As Recordset
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    
    
    
    If DateDiff("d", m_dtRunDate, Now) <= 0 Then '如果查询的是环境中的车次则使用工作表
        szTableName = "Work_env_bus_station_lst "
    Else
        szTableName = "Env_bus_station_lst  "
    End If
'        szSql = "SELECT DISTINCT e.sell_station_id , e.station_id , station_name , mileage , station_serial_no , max_sale_quantity " _
'            & " , sale_ticket_quantity , stop_sale_time, s.sell_station_name , t.station_name   " _
'                & " FROM " & szTableName & " , sell_station_info s , station_info t " _
'                & " WHERE e.station_id = t.station_id AND e.sell_station_id = s.sell_station_id AND s.sell_station_id IN (SELECT sell_station_id FROM " & szTableName & " WHERE  " _
'                & " mileage = (SELECT max(mileage) from " & szTableName & " WHERE bus_id = '" & m_szBusID & "'" _
'                & " AND bus_date = '" & ToDBDate(m_dtRunDate) & "' ) AND bus_id = '" & m_szBusID & "'" _
'                & " AND bus_date = '" & ToDBDate(m_dtRunDate) & "' ) " _
'                & " AND e.bus_id = '" & m_szBusID & "' AND e.bus_date = '" & ToDBDate(m_dtRunDate) & "' " _
'                & " ORDER BY  e.sell_station_id , e.station_serial_no "
    szSql = "SELECT e.*, s.sell_station_name , t.station_name , a.seat_type_name   " _
        & " FROM " & szTableName & " e , sell_station_info s , station_info t, seat_type_code a " _
        & " WHERE e.station_id = t.station_id AND e.sell_station_id = s.sell_station_id AND e.seat_type_id = a.seat_type_id " _
        & " AND e.bus_id = '" & m_szBusID & "' AND e.bus_date = '" & ToDBDate(m_dtRunDate) & "' " _
        & " ORDER BY  e.sell_station_id , e.station_serial_no "

    Set rsTemp = oDb.Execute(szSql)
    
    '车次无运行站点
    If rsTemp.RecordCount = 0 Then ShowError ERR_BusNoStation
    ReDim BSInfo(1 To rsTemp.RecordCount) 'As TBusStationSellInfo
    For i = 1 To rsTemp.RecordCount
        BSInfo(i).szSellStationID = FormatDbValue(rsTemp!sell_station_id)
        BSInfo(i).szSellStationName = FormatDbValue(rsTemp!sell_station_name)
        
        BSInfo(i).szStationID = FormatDbValue(rsTemp!station_id)
        BSInfo(i).sgHalfPrice = FormatDbValue(rsTemp!half_price)
        BSInfo(i).sgFullPrice = FormatDbValue(rsTemp!full_price)
        BSInfo(i).nMileage = FormatDbValue(rsTemp!mileage)
        BSInfo(i).sgLimitedSellTime = FormatDbValue(rsTemp!stop_sale_time)
        BSInfo(i).nLimitedSellCount = FormatDbValue(rsTemp!max_sale_quantity)
        BSInfo(i).szStationName = FormatDbValue(rsTemp!station_name)
        BSInfo(i).sgPreferentialPrice1 = FormatDbValue(rsTemp!preferential_ticket1)
        BSInfo(i).sgPreferentialPrice2 = FormatDbValue(rsTemp!preferential_ticket2)
        BSInfo(i).sgPreferentialPrice3 = FormatDbValue(rsTemp!preferential_ticket3)
        BSInfo(i).szSeatTypeID = FormatDbValue(rsTemp!seat_type_id)
        BSInfo(i).szSeatTypeName = FormatDbValue(rsTemp!seat_type_name)
        rsTemp.MoveNext
    Next
    GetBusStation = BSInfo
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function
''得到某上车站的
'Public Function GetBusStation2(Optional pszCheckGateID As String) As TREBusStationInfo()   'TBusStationSellInfo()
'Dim oDb As New RTConnection
'Dim szSql As String
'Dim i As Integer
'Dim BSInfo() As TREBusStationInfo
'Dim szTableName As String '表名
'
'Dim rsTemp As Recordset
'    AssertObjIsValid
'    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
'
'
'
'    If DateDiff("d", m_dtRunDate, Now) <= 0 Then '如果查询的是环境中的车次则使用工作表
'        szTableName = "Work_env_bus_station_lst e"
'    Else
'        szTableName = "Env_bus_station_lst e"
'    End If
'    If pszCheckGateID <> "" Then
'        szTableName = szTableName & ",Checkgate_info c"
'    End If
'    szSql = "SELECT DISTINCT e.*, s.sell_station_name , t.station_name,y.seat_type_name " _
'            & " FROM " & szTableName & " , sell_station_info s , station_info t,Seat_type_code y " _
'            & " WHERE e.station_id = t.station_id AND e.bus_date = '" & ToDBDate(m_dtRunDate) & "' AND bus_id = '" & m_szBusID & "'"
'            If pszCheckGateID <> "" Then
'                szSql = szSql & " AND e.sell_station_id=c.sell_station_id AND c.check_gate_id=" & TransFieldValueToString(pszCheckGateID)
'            End If
'
'            szSql = szSql & " AND e.seat_type_id=y.seat_type_id ORDER BY  e.sell_station_id , e.station_serial_no,e.seat_type_id "
'    Set rsTemp = oDb.Execute(szSql)
'
'
'    '车次无运行站点
'    If rsTemp.RecordCount = 0 Then ShowError ERR_BusNoStation
'    ReDim BSInfo(1 To rsTemp.RecordCount) 'As TBusStationSellInfo
'    For i = 1 To rsTemp.RecordCount
'        BSInfo(i).szSellStationID = FormatDbValue(rsTemp!sell_station_id)
'        BSInfo(i).szSellStationName = FormatDbValue(rsTemp!sell_station_name)
'        BSInfo(i).szStationID = FormatDbValue(rsTemp!station_id)
'        BSInfo(i).sgHalfPrice = FormatDbValue(rsTemp!half_price)
'        BSInfo(i).sgFullPrice = FormatDbValue(rsTemp!full_price)
'        BSInfo(i).nMileage = FormatDbValue(rsTemp!mileage)
'        BSInfo(i).sgLimitedSellTime = FormatDbValue(rsTemp!stop_sale_time)
'        BSInfo(i).nLimitedSellCount = FormatDbValue(rsTemp!max_sale_quantity)
'        BSInfo(i).szStationName = FormatDbValue(rsTemp!station_name)
'        BSInfo(i).sgPreferentialPrice1 = FormatDbValue(rsTemp!preferential_ticket1)
'        BSInfo(i).sgPreferentialPrice2 = FormatDbValue(rsTemp!preferential_ticket2)
'        BSInfo(i).sgPreferentialPrice3 = FormatDbValue(rsTemp!preferential_ticket3)
'        BSInfo(i).szSeatTypeID = FormatDbValue(rsTemp!seat_type_id)
'''        BSInfo(i).szSeatTypeName = FormatDbValue(rsTemp!seat_type_name)
'        rsTemp.MoveNext
'    Next
'    GetBusStation2 = BSInfo
'    Set rsTemp = Nothing
'    Set oDb = Nothing
'End Function
'F2查询车次的所有过站和站点的属性
'Public Function GetBusStationInfoEx() As TREBusStationSellInfoEx()
'Dim oDb As New RTConnection
'Dim szSql As String
'Dim i As Integer
'Dim BSInfo() As TREBusStationSellInfoEx
'Dim rsTemp As Recordset
'AssertObjIsValid
'    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
'    If DateDiff("d", m_dtRunDate, Now) <= 0 Then '如果查询的是环境中的车次则使用工作表
'        szSql = "SELECT tbe.*,tbs.station_name FROM " _
'                & "Work_env_bus_station_lst tbe,station_info tbs   WHERE " _
'                & "tbe.bus_id='" & m_szBusID & "' AND " _
'                & "tbe.bus_date='" & ToDBDate(m_dtRunDate) & "' AND " _
'                & "tbe.station_id=tbs.station_id ORDER BY tbe.station_serial_no"
'    Else
'        szSql = "SELECT tbe.*,tbs.station_name FROM " _
'                & "Env_bus_station_lst tbe,station_info tbs  WHERE " _
'                & "tbe.bus_id='" & m_szBusID & "' AND " _
'                & "tbe.bus_date='" & ToDBDate(m_dtRunDate) & "' AND " _
'                 & "tbe.station_id=tbs.station_id ORDER BY tbe.station_serial_no"
'    End If
'    Set rsTemp = oDb.Execute(szSql)
'    '-------------------------------
'    '车次无运行站点
'    If rsTemp.RecordCount = 0 Then ShowError ERR_BusNoStation
'    ReDim BSInfo(1 To rsTemp.RecordCount) As TREBusStationSellInfoEx
'    For i = 1 To rsTemp.RecordCount
'        BSInfo(i).szStationID = FormatDbValue(rsTemp!station_id)
'        BSInfo(i).szSeatTypeID = FormatDbValue(rsTemp!seat_type_id)
'        BSInfo(i).sgHalfPrice = FormatDbValue(rsTemp!half_price)
'        BSInfo(i).sgFullPrice = FormatDbValue(rsTemp!full_price)
'        BSInfo(i).nMileage = FormatDbValue(rsTemp!mileage)
'        BSInfo(i).nLimitedSellTime = FormatDbValue(rsTemp!stop_sale_time)
'        BSInfo(i).nLimitedSellCount = FormatDbValue(rsTemp!max_sale_quantity)
'        BSInfo(i).szStationName = FormatDbValue(rsTemp!station_name)
'        BSInfo(i).sgPreferentialPrice1 = FormatDbValue(rsTemp!preferential_ticket1)
'        BSInfo(i).sgPreferentialPrice2 = FormatDbValue(rsTemp!preferential_ticket2)
'        BSInfo(i).sgPreferentialPrice3 = FormatDbValue(rsTemp!preferential_ticket3)
'        rsTemp.MoveNext
'    Next
'    GetBusStationInfoEx = BSInfo
'    Set rsTemp = Nothing
'    Set oDb = Nothing
'End Function
'不使用GetBusStationInfoEx因为使用GetBusStationInfoEx时会出现原代码能正常运行而可执行程序不能正常执行,而使用GetBusStationInfoRec
'F2查询车次的所有过站和站点的属性
'Public Function GetBusStationInfoRec() As Recordset
'Dim oDb As New RTConnection
'Dim szSql As String
'Dim i As Integer
'Dim BSInfo() As TBusStationSellInfo
'Dim rsTemp As Recordset
'AssertObjIsValid
'    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
'    If DateDiff("d", m_dtRunDate, Now) <= 0 Then '如果查询的是环境中的车次则使用工作表
'        szSql = "SELECT tbe.*,tbs.station_name FROM " _
'                & "Work_env_bus_station_lst tbe,station_info tbs   WHERE " _
'                & "tbe.bus_id='" & m_szBusID & "' AND " _
'                & "tbe.bus_date='" & ToDBDate(m_dtRunDate) & "' AND " _
'                & "tbe.station_id=tbs.station_id ORDER BY tbe.station_serial_no"
'    Else
'        szSql = "SELECT tbe.*,tbs.station_name FROM " _
'                & "Env_bus_station_lst tbe,station_info tbs  WHERE " _
'                & "tbe.bus_id='" & m_szBusID & "' AND " _
'                & "tbe.bus_date='" & ToDBDate(m_dtRunDate) & "' AND " _
'                 & "tbe.station_id=tbs.station_id ORDER BY tbe.station_serial_no"
'    End If
'    Set rsTemp = oDb.Execute(szSql)
'    '-------------------------------
'    '车次无运行站点
'    If rsTemp.RecordCount = 0 Then ShowError ERR_BusNoStation
'    Set GetBusStationInfoRec = rsTemp
'    Set rsTemp = Nothing
'    Set oDb = Nothing
'End Function

'F3获得车次过站数
'Public Function GetBusStationCount() As Integer
'Dim oDb As New RTConnection
'Dim szSql As String
'Dim rsTemp As Recordset
'AssertObjIsValid
'    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
'    If DateDiff("d", m_dtRunDate, Now) <= 0 Then '如果查询的是环境中的车次则使用工作表
'        szSql = "SELECT COUNT(*) AS StationCount FROM Work_env_bus_station_lst WHERE bus_id='" & m_szBusID & "' AND " _
'        & " bus_date='" & ToDBDate(m_dtRunDate) & "'"
'    Else
'        szSql = "SELECT COUNT(*) AS StationCount FROM Env_bus_station_lst WHERE bus_id='" & m_szBusID & "' AND " _
'        & " bus_date='" & ToDBDate(m_dtRunDate) & "'"
Public Function GetEnBusStationInfo() As String()
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As New Recordset
    Dim szDate() As String
    Dim i As Integer
    Dim nCount As Integer
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    On Error GoTo here
   szSql = "SELECT w.bus_id,w.station_id,s.station_input_code,s.station_name ,w.max_sale_quantity FROM Work_env_bus_station_lst w," _
            & " station_info s WHERE w.bus_id='" & m_szBusID & "' AND w.bus_date='" & ToDBDate(m_dtRunDate) & "' " _
            & " AND w.station_id =s.station_id order BY w.station_serial_no "
   Set rsTemp = oDb.Execute(szSql)
   nCount = rsTemp.RecordCount

   If nCount <> 0 Then
       ReDim szDate(1 To nCount, 1 To 5)

        For i = 1 To nCount

             szDate(i, 1) = FormatDbValue(rsTemp!bus_id)
             szDate(i, 2) = FormatDbValue(rsTemp!station_id)
             szDate(i, 3) = FormatDbValue(rsTemp!station_name)
             szDate(i, 4) = FormatDbValue(rsTemp!max_sale_quantity)
             szDate(i, 5) = FormatDbValue(rsTemp!station_input_code)
             rsTemp.MoveNext
        Next
        GetEnBusStationInfo = szDate
    End If
    Set oDb = Nothing
    Exit Function
here:
    err.Raise err.Number
    Set oDb = Nothing
End Function

Public Function GetRemainSeatInfo(szStationID As String) As String()
 Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As New Recordset
    Dim szDate() As String
    Dim i As Integer
    Dim nCount As Integer
    Dim szWhere As String
    Dim bflg As Boolean
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    On Error GoTo here
    If bflg = True Then
        szWhere = " bw.status='" & ST_BusNormal & "'ORDER BY esw.bus_id"
    Else
        szWhere = " ORDER BY esw.bus_id"
    End If
    szSql = " SELECT DISTINCT esw.bus_id,bw.vehicle_type_name,bw.route_id,r.route_name,bw.vehicle_type_code,bw.total_seat," _
        & " bw.sale_seat_quantity,bw.seat_remain,bw.bed_remain , " _
        & " bw.additional_remain , bw.other_remain_1, bw.other_remain_2 " _
        & " FROM Work_env_bus_station_lst  esw,Work_env_bus_info bw , route_info r " _
        & " WHERE r.route_id=bw.route_id AND  esw.bus_id=bw.bus_id " _
        & " AND esw.bus_date='" & ToDBDate(m_dtRunDate) & "' " _
        & " AND esw.station_id='" & szStationID & "'  AND esw.bus_id= bw.bus_id AND esw.bus_date = bw.bus_date  "
        ' AND esw.bus_date=bw.bus_date不需要关联环境车次站点表
    szSql = szSql & szWhere
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
        ReDim szDate(1 To nCount, 1 To 12)
        For i = 1 To nCount
            szDate(i, 1) = FormatDbValue(rsTemp!bus_id)
            szDate(i, 2) = FormatDbValue(rsTemp!vehicle_type_name)
            szDate(i, 3) = FormatDbValue(rsTemp!route_id)
            szDate(i, 4) = FormatDbValue(rsTemp!route_name)
            szDate(i, 5) = FormatDbValue(rsTemp!vehicle_type_code)
            szDate(i, 6) = FormatDbValue(rsTemp!total_seat)
            szDate(i, 7) = FormatDbValue(rsTemp!sale_seat_quantity)
            szDate(i, 8) = FormatDbValue(rsTemp!seat_remain)
            szDate(i, 9) = FormatDbValue(rsTemp!bed_remain)
            szDate(i, 10) = FormatDbValue(rsTemp!additional_remain)
            szDate(i, 11) = FormatDbValue(rsTemp!other_remain_1)
            szDate(i, 12) = FormatDbValue(rsTemp!other_remain_2)
            rsTemp.MoveNext
        Next
        GetRemainSeatInfo = szDate
    End If

    Set oDb = Nothing
    Exit Function
here:

    err.Raise err.Number
    Set oDb = Nothing
End Function
'F4获得车次座位信息
Public Function GetOtherSeatInfo(pbIsSelfStation As Boolean, Optional pszSellStationID As String = "") As TSeatInfoEx()

Dim oDb As New RTConnection
Dim szSql As String
Dim tSeat() As TSeatInfoEx
Dim i As Integer
Dim rsTemp As Recordset
Dim rsMegerTemp As Recordset
Dim rsTicketTemp As Recordset
Dim bflg As Boolean
Dim szWhere As String
AssertObjIsValid
    bflg = False
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)

    If pbIsSelfStation Then
        If pszSellStationID <> "" Then szWhere = " AND tbs.sell_station_id = '" & pszSellStationID & "'"
    Else
        If pszSellStationID <> "" Then szWhere = " AND tbs.sell_station_id <> '" & pszSellStationID & "'"
    End If
    
    szSql = "SELECT ticket_id,tbe.seat_no,ticket_price,tbs.status  AS checked_status " _
            & " FROM  (select seat_no, ticket_no, bus_date from Env_bus_seat_lst where bus_date='" & ToDBDate(m_dtRunDate) & "' AND bus_id='" & m_szBusID & "') tbe ," _
            & "       (select bus_date, sell_station_id,ticket_id,ticket_price, status from Ticket_sell_lst where bus_date='" & ToDBDate(m_dtRunDate) & "' AND bus_id='" & m_szBusID & "') tbs  " _
            & " WHERE tbe.bus_date*=tbs.bus_date " _
            & " AND tbe.ticket_no*=tbs.ticket_id " & szWhere _
            & " order BY tbe.seat_no"
    
    Set rsTicketTemp = oDb.Execute(szSql)
    'szsql = "SELECT *  FROM Mergebus_split_lst WHERE  bus_id='" & m_szBusID & "' AND " _
  '         & " bus_date='" & ToDBDate(m_dtRunDate) & "' AND status=4 "
  '  Set rsMegerTemp = odb.Execute(szsql)
   ' If rsMegerTemp.RecordCount <> 0 Then bflg = True
    '车次无座位
    If rsTicketTemp.RecordCount <> 0 Then
    ReDim tSeat(1 To rsTicketTemp.RecordCount) As TSeatInfoEx
    
    For i = 1 To rsTicketTemp.RecordCount
        tSeat(i).szSeatNo = FormatDbValue(rsTicketTemp!seat_no)
        tSeat(i).szSeatStatus = -1
'        tSeat(i).szTicketNo = FormatDbValue(rsTicketTemp!ticket_no)
        tSeat(i).szSeatType = ""
        tSeat(i).szSeatTypeName = ""
        tSeat(i).szRemark = ""
        tSeat(i).szTicketNo = FormatDbValue(rsTicketTemp!ticket_id)
        
'        If rsTicketTemp!checked_status = ST_TicketChecked + ST_TicketNormal Or rsTicketTemp!checked_status = ST_TicketChecked + ST_TicketSellChange Then
'              tSeat(i).szTicketNo = tSeat(i).szTicketNo & "[已检]"
'        End If
        
        rsTicketTemp.MoveNext
    Next
    GetOtherSeatInfo = tSeat
    End If
    Set oDb = Nothing
End Function

Public Function GetSeatInfo(Optional pszSellStationID As String = "") As TSeatInfoEx()

Dim oDb As New RTConnection
Dim szSql As String
Dim tSeat() As TSeatInfoEx
Dim i As Integer
Dim rsTemp As Recordset
Dim rsMegerTemp As Recordset
Dim rsTicketTemp As Recordset
Dim bflg As Boolean
Dim szWhere As String
AssertObjIsValid
    bflg = False
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    szSql = "SELECT eni.* ,sti.seat_type_name FROM  Env_bus_seat_lst eni ,Seat_type_code sti  WHERE bus_id='" & m_szBusID & "' AND " _
    & " bus_date='" & ToDBDate(m_dtRunDate) & "' AND sti.seat_type_id=eni.seat_type_id order BY eni.seat_no"
    Set rsTemp = oDb.Execute(szSql)
    
    
    If pszSellStationID <> "" Then szWhere = " AND tbs.sell_station_id = '" & pszSellStationID & "'"
    szSql = "SELECT ticket_no,station_name,ticket_price,tbs.status  AS checked_status FROM  Env_bus_seat_lst tbe,Ticket_sell_lst  tbs " _
            & ",station_info tbz WHERE tbe.bus_date='" & ToDBDate(m_dtRunDate) & "' AND tbe.bus_id='" & m_szBusID & "'" _
            & " AND tbe.bus_date=tbs.bus_date " _
            & " AND tbe.ticket_no=tbs.ticket_id AND tbs.des_station_id=tbz.station_id " & szWhere _
            & " order BY tbe.seat_no"
    
    Set rsTicketTemp = oDb.Execute(szSql)
    'szsql = "SELECT *  FROM Mergebus_split_lst WHERE  bus_id='" & m_szBusID & "' AND " _
  '         & " bus_date='" & ToDBDate(m_dtRunDate) & "' AND status=4 "
  '  Set rsMegerTemp = odb.Execute(szsql)
   ' If rsMegerTemp.RecordCount <> 0 Then bflg = True
    '车次无座位
    If rsTemp.RecordCount <> 0 Then
    ReDim tSeat(1 To rsTemp.RecordCount) As TSeatInfoEx
    
    For i = 1 To rsTemp.RecordCount
        tSeat(i).szSeatNo = FormatDbValue(rsTemp!seat_no)
        tSeat(i).szSeatStatus = FormatDbValue(rsTemp!Status)
        tSeat(i).szTicketNo = FormatDbValue(rsTemp!ticket_no)
        tSeat(i).szSeatType = FormatDbValue(rsTemp!seat_type_id)
        tSeat(i).szSeatTypeName = FormatDbValue(rsTemp!seat_type_name)
        tSeat(i).szRemark = FormatDbValue(rsTemp!remark)
        
        If tSeat(i).szSeatStatus = ST_SeatSold Or tSeat(i).szSeatStatus = 5 Or tSeat(i).szSeatStatus = 4 Then
           'If rsTicketTemp.RecordCount > 0 Then
            
            On Error Resume Next
                rsTicketTemp.MoveFirst
                rsTicketTemp.Find ("ticket_no='" & tSeat(i).szTicketNo & "'")
                If rsTicketTemp.EOF Then
                    '如果未找到,则该票不属于本车站.
                    tSeat(i).szSeatStatus = ST_SeatCanSell
                Else
                    tSeat(i).szDestName = FormatDbValue(rsTicketTemp!station_name)
                    tSeat(i).szTicketPrice = FormatDbValue(rsTicketTemp!ticket_price)
                    If rsTicketTemp!checked_status = ST_TicketChecked + ST_TicketNormal Or rsTicketTemp!checked_status = ST_TicketChecked + ST_TicketSellChange Then
                        
                       ' If bflg <> True Then '拆分得到
                          tSeat(i).szTicketNo = tSeat(i).szTicketNo & "[已检]"
                       ' Else
                       '   tSeat(i).szTicketNo = tSeat(i).szTicketNo
                       ' End If
                    End If
                End If
           'End If
        End If
        rsTemp.MoveNext
    Next
    GetSeatInfo = tSeat
    End If
    Set oDb = Nothing
End Function

Public Function GetSeatInfoEx(Optional pszSellStationID As String = "") As TSeatInfoEx2()

Dim oDb As New RTConnection
Dim szSql As String
Dim tSeat() As TSeatInfoEx2
Dim i As Integer
Dim rsTemp As Recordset
Dim rsMegerTemp As Recordset
Dim rsTicketTemp As Recordset
Dim bflg As Boolean
Dim szWhere As String
AssertObjIsValid
    bflg = False
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    szSql = "SELECT eni.* ,sti.seat_type_name FROM  Env_bus_seat_lst eni ,Seat_type_code sti  WHERE bus_id='" & m_szBusID & "' AND " _
    & " bus_date='" & ToDBDate(m_dtRunDate) & "' AND sti.seat_type_id=eni.seat_type_id order BY eni.seat_no"
    Set rsTemp = oDb.Execute(szSql)
    
    
    If pszSellStationID <> "" Then szWhere = " AND tbs.sell_station_id = '" & pszSellStationID & "'"
    szSql = "SELECT ticket_no,station_name,ticket_price,tbs.status  AS checked_status ,tbs.card_type ,tbs.id_card_no ,tbs.person_name ,tbs.sex ,tbs.person_picture FROM  Env_bus_seat_lst tbe,Ticket_sell_lst  tbs " _
            & ",station_info tbz WHERE tbe.bus_date='" & ToDBDate(m_dtRunDate) & "' AND tbe.bus_id='" & m_szBusID & "'" _
            & " AND tbe.bus_date=tbs.bus_date " _
            & " AND tbe.ticket_no=tbs.ticket_id AND tbs.des_station_id=tbz.station_id " & szWhere _
            & " order BY tbe.seat_no"
    
    Set rsTicketTemp = oDb.Execute(szSql)
    'szsql = "SELECT *  FROM Mergebus_split_lst WHERE  bus_id='" & m_szBusID & "' AND " _
  '         & " bus_date='" & ToDBDate(m_dtRunDate) & "' AND status=4 "
  '  Set rsMegerTemp = odb.Execute(szsql)
   ' If rsMegerTemp.RecordCount <> 0 Then bflg = True
    '车次无座位
    If rsTemp.RecordCount <> 0 Then
    ReDim tSeat(1 To rsTemp.RecordCount) As TSeatInfoEx2
    
    For i = 1 To rsTemp.RecordCount
        tSeat(i).szSeatNo = FormatDbValue(rsTemp!seat_no)
        tSeat(i).szSeatStatus = FormatDbValue(rsTemp!Status)
        tSeat(i).szTicketNo = FormatDbValue(rsTemp!ticket_no)
        tSeat(i).szSeatType = FormatDbValue(rsTemp!seat_type_id)
        tSeat(i).szSeatTypeName = FormatDbValue(rsTemp!seat_type_name)
        tSeat(i).szRemark = FormatDbValue(rsTemp!remark)
        
        If tSeat(i).szSeatStatus = ST_SeatSold Or tSeat(i).szSeatStatus = 5 Or tSeat(i).szSeatStatus = 4 Then
           'If rsTicketTemp.RecordCount > 0 Then
            
            On Error Resume Next
                rsTicketTemp.MoveFirst
                rsTicketTemp.Find ("ticket_no='" & tSeat(i).szTicketNo & "'")
                If rsTicketTemp.EOF Then
                    '如果未找到,则该票不属于本车站.
                    tSeat(i).szSeatStatus = ST_SeatCanSell
                Else
                    tSeat(i).szDestName = FormatDbValue(rsTicketTemp!station_name)
                    tSeat(i).szTicketPrice = FormatDbValue(rsTicketTemp!ticket_price)
                    
                    tSeat(i).szCardType = FormatDbValue(rsTicketTemp!card_type)
                    tSeat(i).szPersonName = FormatDbValue(rsTicketTemp!person_name)
                    tSeat(i).szIDCardNo = FormatDbValue(rsTicketTemp!id_card_no)
                    tSeat(i).szNational = ""
                    tSeat(i).szSex = FormatDbValue(rsTicketTemp!Sex)
                    tSeat(i).szAddress = ""
                    tSeat(i).szPersonPicture = FormatDbValue(rsTicketTemp!person_picture)
                    
                    If rsTicketTemp!checked_status = ST_TicketChecked + ST_TicketNormal Or rsTicketTemp!checked_status = ST_TicketChecked + ST_TicketSellChange Then
                        
                       ' If bflg <> True Then '拆分得到
                          tSeat(i).szTicketNo = tSeat(i).szTicketNo & "[已检]"
                       ' Else
                       '   tSeat(i).szTicketNo = tSeat(i).szTicketNo
                       ' End If
                    End If
                End If
           'End If
        End If
        rsTemp.MoveNext
    Next
    GetSeatInfoEx = tSeat
    End If
    Set oDb = Nothing
End Function

Public Function GetScrollBusSerial() As String()
Dim oDb As New RTConnection
Dim szSql As String
Dim i As Integer
Dim szaBusSerial() As String
Dim rsTemp As Recordset
AssertObjIsValid
    If m_nBusType = TP_RegularBus Then ShowError ERR_REBusRegularBusNotSerial
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    szSql = "SELECT * FROM Bus_check_lst WHERE bus_id='" & m_szBusID & "' AND " _
    & " bus_date='" & ToDBDate(m_dtRunDate) & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_REBusNoStartCheck
    ReDim szaBusSerial(1 To rsTemp.RecordCount) As String
    For i = 1 To rsTemp.RecordCount
        szaBusSerial(i) = FormatDbValue(rsTemp!bus_serial_no)
        rsTemp.MoveNext
    Next
    GetScrollBusSerial = szaBusSerial
    Set oDb = Nothing
End Function

' *******************************************************************
' *   Brief Description:                                            *
' *   Engineer: fl                                                  *
' *   Date Generated: 2002/11/13                                    *
' *   Last Revision Date:                                           *
' *******************************************************************

'F6获得某站点，某票中的票价信息,财务拆帐用，请不要删除
Public Function GetBusTicketInfoEx(StationID As String, TicketType As Integer, SeatTypeID As String, SellStationID As String) As TRETicketPrice
    Dim szSql As String
    Dim tTicketPrice As TRETicketPrice
    Dim rsTemp As New Recordset
    Dim oDb As New RTConnection

    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    If DateDiff("d", m_dtRunDate, Now) <= 0 Then '如果查询的是环境中的车次则使用工作表
        szSql = "SELECT tbe.*, tbs.sell_station_name " _
        & "FROM Work_env_bus_price_lst tbe ,sell_station_info tbs " _
        & "WHERE tbe.station_id = '" & StationID & "' AND " _
        & "tbe.ticket_type = '" & Trim(TicketType) & "' AND " _
        & "tbe.seat_type_id = '" & Trim(SeatTypeID) & "'AND " _
        & "tbs.sell_station_id *= tbe.sell_station_id AND " _
        & "tbs.sell_station_id = '" & Trim(SellStationID) & "' AND " _
        & "tbe.bus_date ='" & ToDBDate(m_dtRunDate) & "'"

    Else
        szSql = "SELECT tbe.* ,tbs.sell_station_name " _
        & "FROM Env_bus_price_lst tbe ,sell_station_info tbs " _
        & "WHERE tbe.station_id = '" & StationID & "' AND " _
        & "tbe.ticket_type = '" & TicketType & "' AND " _
        & "tbe.seat_type_id = '" & SeatTypeID & "'AND " _
        & "tbe.sell_station_id *= tbs.sell_station_id  AND " _
        & "tbe.sell_station_id = '" & SellStationID & "' AND " _
        & "tbe.bus_date ='" & ToDBDate(m_dtRunDate) & "'"
        
    End If
    Set rsTemp = oDb.Execute(szSql)

    If rsTemp.RecordCount = 0 Then
        ShowError ERR_REBusNotStationPrice
        Exit Function
    End If
    rsTemp.MoveFirst
        Dim k As Integer
        tTicketPrice.szSellStationID = FormatDbValue(rsTemp!sell_station_id)
        tTicketPrice.szSellStationName = FormatDbValue(rsTemp!sell_station_name)
        tTicketPrice.sgBase = FormatDbValue(rsTemp!base_carriage)
        For k = 1 To 15
            tTicketPrice.asgPrice(k) = rsTemp.Fields("price_item_" & k)
        Next

        tTicketPrice.nTicketType = TicketType
        tTicketPrice.szStationID = StationID
        GetBusTicketInfoEx = tTicketPrice

    Set rsTemp = Nothing
    Set oDb = Nothing
End Function



'Public Function GetQueryTicketInfo(TicketType As ETicketType) As TQueryTicketPrice()
'    Dim szSql As String
'    Dim tTicketPrice() As TQueryTicketPrice
'    Dim rsTemp As Recordset
'    Dim oDb As New RTConnection
'    Dim nArrayCount As Long
'    Dim nIndex As Long
'    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
'    If DateDiff("d", m_dtRunDate, Now) <= 0 Then '如果查询的是环境中的车次则使用工作表
'    szSql = "SELECT tbp.*,station_name," _
'         & " tbs.mileage,vehicle_type_name " _
'         & " FROM enviroment_price_work tbp,station_info,Work_env_bus_station_lst tbs,Work_env_bus_info tbb" _
'         & " WHERE tbp.bus_id='" & m_szBusID _
'         & "' AND tbp.bus_date='" & ToDBDate(m_dtRunDate) _
'         & "' AND tbp.station_id=station_info.station_id" _
'         & " AND tbp.bus_id=tbs.bus_id" _
'         & " AND tbp.bus_date=tbs.bus_date" _
'         & " AND tbp.station_id=tbs.station_id" _
'         & " AND tbp.bus_id=tbb.bus_id" _
'         & " AND tbp.bus_date=tbb.bus_date"
'    Else
'    szSql = "SELECT tbp.*,station_name," _
'    & " tbs.mileage,vehicle_type_name " _
'    & " FROM Env_bus_price_lst tbp,station_info,Env_bus_station_lst tbs,Env_bus_info tbb" _
'    & " WHERE tbp.bus_id='" & m_szBusID _
'    & "' AND tbp.bus_date='" & ToDBDate(m_dtRunDate) _
'    & "' AND tbp.station_id=station_info.station_id" _
'    & " AND tbp.bus_id=tbs.bus_id" _
'    & " AND tbp.bus_date=tbs.bus_date" _
'    & " AND tbp.station_id=tbs.station_id" _
'    & " AND tbp.bus_id=tbb.bus_id" _
'    & " AND tbp.bus_date=tbb.bus_date"
'    End If
'
'    If TicketType <> TP_FreeTicket Then
'        szSql = szSql & " AND Env_bus_price_lst.ticket_type='" & TicketType & "'"
'    End If
'
'
'
'    Set rsTemp = oDb.Execute(szSql)
'    nArrayCount = rsTemp.RecordCount
'    If nArrayCount > 1 Then ReDim tTicketPrice(1 To nArrayCount)
'    For nIndex = 1 To nArrayCount
'
'
'        tTicketPrice(nIndex).szStationName = FormatDbValue(rsTemp!station_name)
'        tTicketPrice(nIndex).szVicleModel = FormatDbValue(rsTemp!vehicle_type_name)
'        tTicketPrice(nIndex).sgMileage = FormatDbValue(rsTemp!mileage)
''        tTicketPrice(nIndex).szStationID = StationID
'        tTicketPrice(nIndex).sgBase = FormatDbValue(rsTemp!base_carriage)
'        tTicketPrice(nIndex).sgPrice1 = FormatDbValue(rsTemp!price_item_1)
'        tTicketPrice(nIndex).sgPrice2 = FormatDbValue(rsTemp!price_item_2)
'        tTicketPrice(nIndex).sgPrice3 = FormatDbValue(rsTemp!price_item_3)
'        tTicketPrice(nIndex).sgPrice4 = FormatDbValue(rsTemp!price_item_4)
'        tTicketPrice(nIndex).sgPrice5 = FormatDbValue(rsTemp!price_item_5)
'        tTicketPrice(nIndex).sgPrice6 = FormatDbValue(rsTemp!price_item_6)
'        tTicketPrice(nIndex).sgPrice7 = FormatDbValue(rsTemp!price_item_7)
'        tTicketPrice(nIndex).sgPrice8 = FormatDbValue(rsTemp!price_item_8)
'        tTicketPrice(nIndex).sgPrice9 = FormatDbValue(rsTemp!price_item_9)
'        tTicketPrice(nIndex).sgPrice10 = FormatDbValue(rsTemp!price_item_10)
'        tTicketPrice(nIndex).sgPrice11 = FormatDbValue(rsTemp!price_item_11)
'        tTicketPrice(nIndex).sgPrice12 = FormatDbValue(rsTemp!price_item_12)
'        tTicketPrice(nIndex).sgPrice13 = FormatDbValue(rsTemp!price_item_13)
'        tTicketPrice(nIndex).sgPrice14 = FormatDbValue(rsTemp!price_item_14)
'        tTicketPrice(nIndex).sgPrice15 = FormatDbValue(rsTemp!price_item_15)
'        tTicketPrice(nIndex).sgTotal = FormatDbValue(rsTemp!ticket_price_total)
'        tTicketPrice(nIndex).nTicketType = TicketType
'        rsTemp.MoveNext
'    Next nIndex
'    GetQueryTicketInfo = tTicketPrice
'    Set oDb = Nothing
'End Function
'F7开始日期和结束日期是否在预售日期范围之内
Private Function InPreDate(StartDate As Date, EndDate As Date) As Boolean
Dim oPara As New SystemParam
oPara.Init m_oActiveUser
InPreDate = True
If DateDiff("d", oPara.NowDate, StartDate) < 0 Then InPreDate = False: Exit Function
If DateDiff("d", oPara.NowDate, EndDate) < 0 Then InPreDate = False: Exit Function
If DateDiff("d", oPara.NowDate, EndDate) > oPara.PreSellDate Then InPreDate = False: Exit Function
End Function

'F8获得车次座位信息
Public Function GetStationSeatInfo() As String()
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim aszSeat() As String
    Dim i, j As Integer
    Dim rsTemp As Recordset
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    szSql = "SELECT des_station_id,COUNT(*) AS ticket_count FROM Ticket_sell_lst WHERE bus_id='" & m_szBusID & "' AND bus_date='" & ToDBDate(m_dtRunDate) & "' AND NOT((status & (16|4|8)<>0))  GROUP BY des_station_id"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 0 Then
    ReDim aszSeat(1 To rsTemp.RecordCount, 1 To 2) As String
    
    For i = 1 To rsTemp.RecordCount
        aszSeat(i, 1) = FormatDbValue(rsTemp!des_station_id)
        aszSeat(i, 2) = FormatDbValue(rsTemp!ticket_count)
        rsTemp.MoveNext
    Next i

    szSql = " SELECT a.des_station_id,COUNT(*) AS ticket_count  FROM Ticket_sell_lst a,check_ticket_lst b " _
        & " WHERE a.bus_id='" & m_szBusID & "' AND a.bus_date='" & ToDBDate(m_dtRunDate) & "'" _
        & " AND NOT((a.status & (16|4|8)<>0)) AND (b.status =2) AND a.ticket_id=b.ticket_id " _
        & " GROUP BY des_station_id "
    Set rsTemp = oDb.Execute(szSql)
    

    For i = 1 To rsTemp.RecordCount
        For j = 1 To ArrayLength(aszSeat)
            If aszSeat(j, 1) = FormatDbValue(rsTemp!des_station_id) Then
                aszSeat(j, 2) = aszSeat(j, 2) - FormatDbValue(rsTemp!ticket_count)
            End If
        Next j
        rsTemp.MoveNext
    Next i
    
    GetStationSeatInfo = aszSeat
    End If
    
    Set oDb = Nothing

End Function
'P1
Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
End Property

'P1
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
'P2
Public Property Get ObjStatus() As EObjectStatus
    ObjStatus = m_nObjectStatus
End Property
'P3
Public Property Let BusID(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then ShowError ERR_REBusNotAddStatus
    m_szBusID = vData
End Property
'P3
Public Property Get BusID() As String
    BusID = m_szBusID
End Property
'P4
Public Property Let RunDate(ByVal vData As Date)
    If m_nObjectStatus <> ST_AddObj Then ShowError ERR_REBusNotAddStatus
    m_dtRunDate = vData
End Property
'P4
Public Property Get RunDate() As Date
    RunDate = m_dtRunDate
End Property

'P5
Public Property Get CheckSheet(Optional ByVal BusSerial As Integer = 0) As String
    CheckSheet = m_szCheckSheet
End Property

'P6
Public Property Get EndStationName() As String
    EndStationName = m_szEndStationName
End Property

'P7
'0==FL=}==========>

Public Property Get CheckStatus() As EREBusCheckStatus
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim eCheckStatus As EREBusCheckStatus
    Dim szSql As String
    AssertObjIsValid
    On Error GoTo here
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    szSql = "SELECT *,tbt.status " _
    & "FROM Bus_check_lst tbc,check_ticket_lst tbt " _
    & "WHERE tbc.bus_id = '" & m_szBusID & "' AND " _
    & "tbc.check_gate_id = tbt.check_gate_id AND " _
    & "tbc.bus_date = '" & ToDBDate(m_dtRunDate) & "'AND " _
    & "tbc.check_gate_id = '" & m_szCheckGateID & " 'ORDER BY  tbc.bus_serial_no DESC "
    Set rsTemp = oDb.Execute(szSql)
    '--------------------------
    '记录数为0，则车次未开检
    If rsTemp.RecordCount = 0 Then eCheckStatus = ST_BusCheckWait: CheckStatus = eCheckStatus: Exit Sub
    '-------------------------
    '时间为空则车次正在检票
    If DBTimeIsEmpty(FormatDbValue(rsTemp!check_end_time)) Then eCheckStatus = ST_BusChecking: CheckStatus = eCheckStatus: Exit Sub
    '------------------------
    '开检时间大于停检时间则车次为补检,否则车次为停检
    If FormatDbValue(rsTemp!check_start_time) > FormatDbValue(rsTemp!check_end_time) Then
    eCheckStatus = ST_BusCheckExChecking: CheckStatus = eCheckStatus: Exit Sub
    Else
    eCheckStatus = ST_BusCheckStopped: CheckStatus = eCheckStatus: Exit Sub
    End If
    m_eCheckStatus = eCheckStatus
    CheckStatus = eCheckStatus
Exit Property
here:
End Property
'P8
Public Property Get BusMergedMe() As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    szSql = "SELECT bus_id FROM Mergebus_split_lst WHERE merge_bus_id='" & m_szBusID & "' AND  bus_date='" & ToDBDate(m_dtRunDate) & "' AND status='" & 0 & "'"
    Set rsTemp = oDb.Execute(szSql)
    '--------------------------
    '我是好的，未被并班
    If rsTemp.RecordCount = 0 Then
        m_szBusMergedMe = ""
    Else
        m_szBusMergedMe = FormatDbValue(rsTemp!bus_id)
    End If
    BusMergedMe = m_szBusMergedMe
End Property

'P9被我并的车次，它坏了G
Public Property Get BeMergedBus() As TMergeInfo
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim i As Integer
    'AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    If m_bRefreshBusMerge Then
        szSql = "SELECT DISTINCT merge_bus_id,annotation FROM Mergebus_split_lst WHERE bus_id='" & m_szBusID & "' AND  bus_date='" & ToDBDate(m_dtRunDate) & "'" 'AND status='" & 0 & "'
        Set rsTemp = oDb.Execute(szSql)
        '-----------------------------
        '车次无并班车次，我是好的,其他车也是好的
       If rsTemp.RecordCount = 0 Then
          m_tMergeBus.nMergeType = 0
       End If
       For i = 1 To rsTemp.RecordCount
            m_tMergeBus.nMergeType = 1
            If m_tMergeBus.szBusID = "" Then
              m_tMergeBus.szBusID = FormatDbValue(rsTemp!merge_bus_id)
            Else
              m_tMergeBus.szBusID = m_tMergeBus.szBusID & "," & FormatDbValue(rsTemp!merge_bus_id)
            End If
            rsTemp.MoveNext
            m_bRefreshBusMerge = False
        Next
    End If
    BeMergedBus = m_tMergeBus
End Property
'P10
Public Property Get Company() As String
    Company = m_szCompany
End Property
Public Property Let Company(vData As String)
    Dim oDb As New RTConnection
    Dim szSql As String
    szSql = "SELECT * FROM Company_info WHERE transport_company_id='" & Trim(vData) & "'"
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
   
  '  ERR_NoSlitpCompany = ERR_REBus + 81 '没有拆帐参运公司。
    If oDb.Execute(szSql).RecordCount = 0 Then ShowError ERR_NoCompany
    Set oDb = Nothing
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szCompany = vData
End Property

'P11
Public Property Let Vehicle(vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    
    m_szVehicle = vData
End Property

'P11
Public Property Get Vehicle() As String
    Vehicle = m_szVehicle
End Property
'P12

Public Property Get Owner() As String
  If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj

    Owner = m_szOwner
End Property

Public Property Let Owner(ByVal vData As String)
    Dim oDb As New RTConnection
    Dim szSql As String
    szSql = "SELECT * FROM owner_info WHERE owner_id='" & Trim(vData) & "'"
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    If oDb.Execute(szSql).RecordCount = 0 Then ShowError ERR_NoOwnerInfo
    Set oDb = Nothing
    m_szOwner = vData
End Property
'P13
Public Property Get EndStation() As String
    EndStation = m_szEndStation
End Property

'P14新增车次时可以修改
Public Property Let Route(ByVal vData As String)
    Dim oDb As New RTConnection
    Dim szSql As String
    If m_nObjectStatus <> ST_AddObj Then ShowError ERR_REBusNotAddStatus
    szSql = "SELECT * FROM route_info WHERE Route_id='" & Trim(vData) & "'"
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    If oDb.Execute(szSql).RecordCount = 0 Then ShowError ERR_NoRoute
    Set oDb = Nothing
    m_szRoute = vData
End Property


Public Property Get Route() As String
    Route = m_szRoute
End Property
'顶班用
Public Property Let ReferenceBusID(ByVal vData As String)
   m_szBusCopy = vData
  
End Property
Public Property Get ReferenceBusID() As String
   ReferenceBusID = m_szBusCopy
End Property
Public Property Get busStatus() As EREBusStatus
    busStatus = m_eBusStatus
End Property
Public Property Let AllRefundment(ByVal vData As Boolean)
Attribute AllRefundment.VB_Description = "全额退票"
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_bAllRefundment = vData
End Property
Public Property Get AllRefundment() As Boolean
   AllRefundment = m_bAllRefundment
End Property

'P17
Public Property Get BeingStatus() As EREBusBeingStatus
    BeingStatus = m_eBeingStatus
End Property
'P18

'是否晚点
Public Property Let DelayStatus(ByVal vData As Boolean)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_bDelayStatus = vData
End Property
Public Property Get DelayStatus() As Boolean
   DelayStatus = m_bDelayStatus
End Property

'

Public Property Get BusType() As Integer

    BusType = m_nBusType
End Property
'P18
Public Property Let BusType(vData As Integer)
  '  If m_nObjectStatus <> ST_AddObj Then ShowError ERR_REBusNotAddStatus
    m_nOldBusType = m_nBusType
    Dim oDb As New RTConnection
    Dim szSql As String
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    szSql = "SELECT * FROM bus_type_code WHERE bus_type='" & CInt(vData) & "'"
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    If oDb.Execute(szSql).RecordCount = 0 Then ShowError ERR_NoBusTypeInfo
    Set oDb = Nothing
    m_nBusType = vData
   
End Property
Public Property Get BusTypeName() As String
'    BusType = m_eBusType
    BusTypeName = m_szBusTypeName
End Property

'P19
Public Property Let StartUpTime(ByVal vData As Date)
    Dim dtOldStartupTime As Date
    dtOldStartupTime = m_dtStartupTime
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_dtStartupTime = vData
    m_dtStartupTime = SelfGetFullDateTime(m_dtRunDate, m_dtStartupTime)
    #If SUPPORT_MSG = 1 Then
        If m_nObjectStatus <> ST_AddObj And dtOldStartupTime <> m_dtStartupTime Then
            m_bOffTimeChanged = True
        End If
    #End If
End Property

'P19
Public Property Get StartUpTime() As Date
    StartUpTime = m_dtStartupTime
End Property


Public Property Let InternetStatus(ByVal vData As Integer)
   
'    m_nOldInternetStatus = m_nInternetStatus
'    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_nInternetStatus = vData
'    #If SUPPORT_MSG = 1 Then
'        If m_nObjectStatus <> ST_AddObj AND dtOldStartupTime <> m_dtStartupTime Then
'            m_bOffTimeChanged = True
'        End If
'    #End If
End Property

'P19
Public Property Get InternetStatus() As Integer
    InternetStatus = m_nInternetStatus
End Property
Public Property Get HaveLugge(Optional ByVal BusSerial As Integer = 0) As Boolean
    HaveLugge = m_bHaveLugge
End Property

'P20
Public Property Get StartSeatNo() As Integer
    StartSeatNo = m_nStartSeatNo
End Property

'P22
Public Property Let CheckGate(ByVal vData As String)
    Dim szOldCheckGate As String
    Dim oDb As New RTConnection
    Dim szSql As String
    szOldCheckGate = Trim(m_szCheckGate)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    szSql = "SELECT * FROM Checkgate_info WHERE check_gate_id='" & Trim(vData) & "'"
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    If oDb.Execute(szSql).RecordCount = 0 Then ShowError ERR_NoCheckgate
    Set oDb = Nothing
    m_szCheckGate = Trim(vData)
    #If SUPPORT_MSG = 1 Then
        If m_nObjectStatus <> ST_AddObj And szOldCheckGate <> m_szCheckGate Then
            m_bCheckGateChanged = True
        End If
    #End If
End Property

'P22
Public Property Get CheckGate() As String
    CheckGate = m_szCheckGate
End Property

'P23
Public Property Get VehicleModelID() As String
    VehicleModelID = m_szVehicleModelID
End Property

'P24
Public Property Get VehicleModelName() As String
    VehicleModelName = m_szVehicleModelName
End Property

'P25
Public Property Get TotalSeat() As Integer
    TotalSeat = m_nTotalSeat
End Property

'P26
Public Property Get SaleSeat() As Integer
    SaleSeat = m_nSaleSeat
End Property

'P27
Public Property Get TotalStandSeat() As Integer
    TotalStandSeat = m_nTotalStandSeat
End Property

'P28
Public Property Get SaleStandSeat() As Integer
    SaleStandSeat = m_nSaleStandSeat
End Property
'P31参运公司名称
Public Property Get CompanyName() As String
    CompanyName = m_szCompanyName
End Property

'P31车主名称
Public Property Get OwnerName() As String
    OwnerName = m_szOwnerName
End Property
'P32线路名称
Public Property Get RouteName() As String
    RouteName = m_szRouteName
End Property



'P33拆帐公司代码
Public Property Get SplitCompanyID() As String
    SplitCompanyID = m_szSplitCompanyID
End Property
Public Property Let SplitCompanyID(vData As String)
    
    Dim oDb As New RTConnection
    Dim szSql As String
    
    szSql = "SELECT * FROM Company_info WHERE transport_company_id='" & Trim(vData) & "'"
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    
    
    If oDb.Execute(szSql).RecordCount = 0 Then ShowError ERR_NoSlitpCompany
    Set oDb = Nothing
   
    
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szSplitCompanyID = vData
   
   
End Property

'P34拆帐公司名称
Public Property Get SplitCompanyName() As String
    SplitCompanyName = m_szSplitCompanyName
End Property


Public Property Let ScrollBusCheckTime(vData As Integer)
'   If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
   m_nScrollBusCheckTime = vData
End Property

'P35滚动车次循环剪票时间
Public Property Get ScrollBusCheckTime() As Integer
    ScrollBusCheckTime = m_nScrollBusCheckTime
End Property


Public Property Get BaggageNumber() As Integer
    BaggageNumber = m_nBaggageNumber
End Property

Public Property Get FactWeight() As Double
    FactWeight = m_dbFactWeight
End Property

Public Property Get CalculateWeight() As Double
    CalculateWeight = m_dbCalculateWeight
End Property

Public Property Get OverWeightNumber() As Integer
    OverWeightNumber = m_nOverWeightNumber
End Property

Public Property Get LuggageTotalPrice() As Double
    LuggageTotalPrice = m_dbLuggageTotalPrice
End Property

Public Property Get SeatRemainCount() As Integer

    SeatRemainCount = m_nSeatRemain
End Property
Public Property Get BedSeatCount() As Integer

    BedSeatCount = m_nBedSeatCount
 
End Property
Public Property Get AdditionalSeatCount() As Integer

    AdditionalSeatCount = m_nAdditionalSeatCount
End Property



'S1
Public Sub Init(poAUser As ActiveUser)
    Set m_oActiveUser = poAUser
End Sub

'S2类刚建立时处于不可用状态,此时类不能使用
Private Sub Class_Initialize()
    m_nObjectStatus = ST_NotAvailable
    #If SUPPORT_MSG = 1 Then
        m_oSndMsg.Unit = GetUnitID()
        m_oSndMsg.MsgSource = cszSNRunEnv
    #End If
End Sub

'S3测试对象的状态是否有效（无效状态[ST_NotAvailable]和新增状态[ST_AddObj]则出错）
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then ShowError ERR_NotAvailable
End Sub
'S4
Private Sub AssertObjIsValid() '测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
    AssertActiveUserValid m_oActiveUser, ERR_REBus
    AssertStatusAvailable
End Sub
'S5
Public Sub Identify(REBusID As String, REBusDate As Date, Optional CheckGate As String)
    AssertActiveUserValid m_oActiveUser, ERR_REBus
    m_szBusID = RTrim(REBusID)
    m_dtRunDate = REBusDate
    RefreshMemoryInfo CheckGate
    m_nObjectStatus = ST_NormalObj
       
    #If SUPPORT_MSG = 1 Then
        m_bCheckGateChanged = False
        m_bOffTimeChanged = False
    #End If
    ''..\目标\SNRunEnv.dll
End Sub

'S6调用新增方法后，对象处于新增状态。
Public Sub AddNew()
    AssertActiveUserValid m_oActiveUser, ERR_REBus
    m_nObjectStatus = ST_AddObj
End Sub
'S7
Public Sub Delete() '调用删除方法后，对象处于不可用状态
    AssertObjIsValid
    DeleteObject
    m_nObjectStatus = ST_NotAvailable
    
    #If SUPPORT_MSG = 1 Then
    Dim oParam As New SystemParam
    Dim aszValues(1 To 2) As String
    aszValues(1) = m_szBusID
    aszValues(2) = ToDBDate(m_dtRunDate)
    m_oSndMsg.SendMsg oParam.NowDateTime, cszRemoveBus, aszValues
    #End If
    
End Sub
'S8
Public Sub Update()
    AssertObjIsValid
#If SUPPORT_MSG = 1 Then
    Dim oParam As New SystemParam
    Dim aszValues() As String
#End If
    
    Select Case m_nObjectStatus
           Case ST_AddObj
           AddObject
           If m_bState Then
           m_bStateOpen = False
           Exit Sub
           End If
           
            #If SUPPORT_MSG = 1 Then
            ReDim aszValues(1 To 3)
            aszValues(1) = m_szBusID
            aszValues(2) = ToDBDate(m_dtRunDate)
            aszValues(3) = m_szCheckGateID
            
            m_oSndMsg.SendMsg oParam.NowDateTime, cszAddBus, aszValues
            m_bCheckGateChanged = False
            m_bOffTimeChanged = False
            #End If
           
           Case ST_EditObj
           
           UpdateToDB
            #If SUPPORT_MSG = 1 Then
            ReDim aszValues(1 To 3)
            aszValues(1) = m_szBusID
            aszValues(2) = ToDBDate(m_dtRunDate)
                If m_bCheckGateChanged Then
                    aszValues(3) = m_szCheckGate
                    m_oSndMsg.SendMsg oParam.NowDateTime, cszChangeBusCheckGate, aszValues
                    m_bCheckGateChanged = False
                End If
                
                If m_bOffTimeChanged Then
                    aszValues(3) = ToDBDateTime(m_dtStartupTime)
                    m_oSndMsg.SendMsg oParam.NowDateTime, cszChangeBusTime, aszValues
                    m_bOffTimeChanged = False
                End If
            #End If
           
           Case ST_NormalObj
           Case Else
           ShowError ERR_NotAvailable
    End Select
    m_nObjectStatus = ST_NormalObj
End Sub

'S9
Private Sub RefreshMemoryInfo(Optional pszCheckGate As String) '按照对象现在的ID将数据库中相应的信息读至对象内
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    If pszCheckGate = "" Then   '一般方式
        szSql = "SELECT r.route_name,b.* , c.transport_company_short_name , v.license_tag_no license,v.start_seat_no," _
            & " o.owner_name , c2.transport_company_short_name split_company_short_name ,bt.bus_type_name" _
            & " FROM "
        
        If DateDiff("d", m_dtRunDate, Now) <= 0 Then '如果查询的是环境中的车次则使用工作表
            'Set rsTemp = oDb.Execute("SELECT tbx.route_name,tbe.*,tbc.transport_company_short_name,tbv.license_tag_no license, tbo.owner_name FROM Work_env_bus_info tbe,Vehicle_info tbv,Company_info tbc,owner_info tbo,route_info tbx WHERE bus_id='" & m_szBusID & "' AND" _
            & " bus_date='" & ToDBDate(m_dtRunDate) & "' AND tbe.transport_company_id=tbc.transport_company_id AND tbe.owner_id=tbo.owner_id AND tbe.vehicle_id=tbv.vehicle_id AND tbe.route_id=tbx.route_id")
            szSql = szSql & " Work_env_bus_info b "
        Else
            'Set rsTemp = oDb.Execute("SELECT tbx.route_name,tbe.*,tbc.transport_company_short_name,tbv.license_tag_no license, tbo.owner_name FROM Env_bus_info tbe,Vehicle_info tbv,Company_info tbc,owner_info tbo,route_info tbx WHERE bus_id='" & m_szBusID & "' AND" _
            & " bus_date='" & ToDBDate(m_dtRunDate) & "' AND tbe.transport_company_id=tbc.transport_company_id AND tbe.owner_id=tbo.owner_id AND tbe.vehicle_id=tbv.vehicle_id AND tbe.route_id=tbx.route_id")
            szSql = szSql & " Env_bus_info b "
        End If
        szSql = szSql & ",Vehicle_info v,Company_info c,owner_info o " _
            & " ,route_info r , Company_info c2 ,bus_type_code bt" _
            & " WHERE bus_id='" & m_szBusID & "' AND bus_date='" & ToDBDate(m_dtRunDate) & "' AND b.transport_company_id=c.transport_company_id AND b.owner_id=o.owner_id " _
            & " AND b.vehicle_id=v.vehicle_id AND b.route_id=r.route_id AND b.split_company_id = c2.transport_company_id " _
            & "  AND bt.bus_type=b.bus_type "
    Else    '得到配载方式下的环境车次信息
        szSql = "SELECT r.route_name,l.check_gate_id,l.bus_start_time,l.Status,b.delay_status, b.route_id,b.vehicle_type_code,b.bus_type,b.end_station_id,b.owner_id,b.vehicle_id,b.is_all_refundment" & _
                ",b.register_status,b.sale_seat_quantity,b.sale_stand_seat_quantity,b.total_seat,v.start_seat_no,b.vehicle_type_name,b.total_stand_quantity,b.transport_company_id" & _
                ",b.split_company_id,b.internet_status,b.ScrollBus_Check_Time,b.seat_remain,b.bed_remain,b.additional_remain,b.baggage_number,b.end_station_name" & _
                ",b.fact_weight,b.calculate_weight,b.over_weight_number,b.luggage_total_price, c.transport_company_short_name , v.license_tag_no license,v.start_seat_no," _
            & " o.owner_name , c2.transport_company_short_name split_company_short_name ,bt.bus_type_name" _
            & " FROM "
        
        If DateDiff("d", m_dtRunDate, Now) <= 0 Then '如果查询的是环境中的车次则使用工作表
            'Set rsTemp = oDb.Execute("SELECT tbx.route_name,tbe.*,tbc.transport_company_short_name,tbv.license_tag_no license, tbo.owner_name FROM Work_env_bus_info tbe,Vehicle_info tbv,Company_info tbc,owner_info tbo,route_info tbx WHERE bus_id='" & m_szBusID & "' AND" _
            & " bus_date='" & ToDBDate(m_dtRunDate) & "' AND tbe.transport_company_id=tbc.transport_company_id AND tbe.owner_id=tbo.owner_id AND tbe.vehicle_id=tbv.vehicle_id AND tbe.route_id=tbx.route_id")
            szSql = szSql & " Work_env_bus_info b,Work_env_bus_allot_lst l "
        Else
            'Set rsTemp = oDb.Execute("SELECT tbx.route_name,tbe.*,tbc.transport_company_short_name,tbv.license_tag_no license, tbo.owner_name FROM Env_bus_info tbe,Vehicle_info tbv,Company_info tbc,owner_info tbo,route_info tbx WHERE bus_id='" & m_szBusID & "' AND" _
            & " bus_date='" & ToDBDate(m_dtRunDate) & "' AND tbe.transport_company_id=tbc.transport_company_id AND tbe.owner_id=tbo.owner_id AND tbe.vehicle_id=tbv.vehicle_id AND tbe.route_id=tbx.route_id")
            szSql = szSql & " Env_bus_info b,Env_bus_allot_lst l"
        End If
        szSql = szSql & ",Vehicle_info v,Company_info c,owner_info o " _
            & " ,route_info r , Company_info c2 ,bus_type_code bt" _
            & " WHERE b.bus_id='" & m_szBusID & "' AND b.bus_date='" & ToDBDate(m_dtRunDate) & "' AND b.transport_company_id=c.transport_company_id AND b.owner_id=o.owner_id " _
            & " AND b.vehicle_id=v.vehicle_id AND b.route_id=r.route_id AND b.split_company_id = c2.transport_company_id " _
            & "  AND bt.bus_type=b.bus_type AND b.bus_date=l.bus_date AND b.bus_id=l.bus_id AND l.check_gate_id=" & TransFieldValueToString(pszCheckGate)
    End If
    Set rsTemp = oDb.Execute(szSql)
    
    If rsTemp.RecordCount = 0 Then ShowError ERR_REBusNotExist
    m_szRouteName = FormatDbValue(rsTemp!route_name)
    m_szOwnerName = FormatDbValue(rsTemp!owner_name)
    m_szCheckGate = FormatDbValue(rsTemp!check_gate_id)
    m_szRoute = FormatDbValue(rsTemp!route_id)
    
    m_szVehicleModelID = FormatDbValue(rsTemp!vehicle_type_code)
    m_szVehicleModelName = FormatDbValue(rsTemp!vehicle_type_name)
    m_nBusType = FormatDbValue(rsTemp!bus_type)
    m_szEndStation = FormatDbValue(rsTemp!end_station_id)
    m_szOwner = FormatDbValue(rsTemp!owner_id)
    m_szVehicle = FormatDbValue(rsTemp!vehicle_id)
    m_szVehicleTag = FormatDbValue(rsTemp!license)
    m_szEndStationName = FormatDbValue(rsTemp!end_station_name)
    m_bAllRefundment = FormatDbValue(rsTemp!is_all_refundment)
    m_dtStartupTime = FormatDbValue(rsTemp!bus_start_time)
    m_eBeingStatus = FormatDbValue(rsTemp!register_status)
    m_eBusStatus = FormatDbValue(rsTemp!Status)
    m_nSaleSeat = FormatDbValue(rsTemp!sale_seat_quantity)
    m_nSaleStandSeat = FormatDbValue(rsTemp!sale_stand_seat_quantity)
    m_nTotalSeat = FormatDbValue(rsTemp!total_seat)
    m_nStartSeatNo = Val(FormatDbValue(rsTemp!start_seat_no))
    m_nTotalStandSeat = FormatDbValue(rsTemp!total_stand_quantity)
    
    m_szCompany = FormatDbValue(rsTemp!transport_company_id)
    m_szCompanyName = FormatDbValue(rsTemp!transport_company_short_name)
    
    m_szSplitCompanyID = FormatDbValue(rsTemp!split_company_id)
    m_szSplitCompanyName = FormatDbValue(rsTemp!split_company_short_name)
    
    m_szBusTypeName = FormatDbValue(rsTemp!bus_type_name)
    
    m_nInternetStatus = FormatDbValue(rsTemp!internet_status)
    m_nScrollBusCheckTime = FormatDbValue(rsTemp!ScrollBus_Check_Time)
    
    
     
     m_nSeatRemain = FormatDbValue(rsTemp!seat_remain)
     m_nBedSeatCount = FormatDbValue(rsTemp!bed_remain)
     m_nAdditionalSeatCount = FormatDbValue(rsTemp!additional_remain)
    
    
    m_nBaggageNumber = FormatDbValue(rsTemp!baggage_number)
    m_dbFactWeight = FormatDbValue(rsTemp!fact_weight)
    m_dbCalculateWeight = FormatDbValue(rsTemp!calculate_weight)
    m_nOverWeightNumber = FormatDbValue(rsTemp!over_weight_number)
    m_dbLuggageTotalPrice = FormatDbValue(rsTemp!luggage_total_price)
    '是否晚点
    m_bDelayStatus = FormatDbValue(rsTemp!delay_status)
    
    
    m_bHaveLugge = False
    
    If m_nBaggageNumber > 0 And _
       m_dbFactWeight > 0 And _
       m_dbCalculateWeight > 0 And _
       m_nOverWeightNumber > 0 And _
       m_dbLuggageTotalPrice > 0 _
    Then m_bHaveLugge = True
    
    
    '赋值为写日志用
    m_noldScrollBusCheckTime = m_nScrollBusCheckTime
    m_nOldInternetStatus = m_nInternetStatus
    m_szOldCompanyName = m_szCompanyName
    m_szOldRouteName = m_szRouteName
    m_szOldOwnerName = m_szOwnerName
    m_szOldCheckGate = m_szCheckGate
    m_szOldRoute = m_szRoute
    
    m_szOldVehicleModelID = m_szVehicleModelID
    m_szOldVehicleModelName = m_szVehicleModelName
    m_eOldBusType = m_nBusType
    m_szOldEndStation = m_szEndStation
    
    
    m_szOldCompany = m_szCompany
    m_szOldOwner = m_szOwner
    m_szOldVehicle = m_szVehicle
    m_szOldVehicleTag = m_szVehicleTag
    m_szOldEndStationName = m_szEndStationName
    
    m_bOldAllRefundment = m_bAllRefundment
    m_dtOldStartupTime = m_dtStartupTime
    m_eOldBeingStatus = m_eBeingStatus
    m_eOldBusStatus = m_eBusStatus
    m_nOldSaleSeat = m_nSaleSeat
    m_nOldSaleStandSeat = m_nSaleStandSeat
    m_nOldTotalSeat = m_nTotalSeat
    m_nOldTotalStandSeat = m_nTotalStandSeat
    m_szOldBusTypeName = m_szBusTypeName
    
    m_szOldSplitCompanyID = m_szSplitCompanyID
    m_szOldSplitCompanyName = m_szSplitCompanyName
    '是否晚点
    m_bOldDelayStatus = m_bDelayStatus
    
    '赋值为写日志用
    
    
    m_bRefreshBusMerge = True
    m_bRefreshBusCheck = True
    Set oDb = Nothing
    
End Sub


Private Sub DeleteObject()
    Dim oDb As New RTConnection
    AssertHaveRight m_oActiveUser, RIGHT_REBusDelete
    
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    '--------------------------
    '车次有售出票不能删除
    If BusHaveNotSellTicket(oDb, m_dtRunDate, m_szBusID) Then ShowError ERR_REBusHaveSellTicketNotDelete
    On Error GoTo here
    oDb.Execute ("DELETE Env_bus_info WHERE bus_id='" & m_szBusID & "' AND " _
    & " bus_date='" & ToDBDate(m_dtRunDate) & "'")
    WriteOperateLog m_oActiveUser, RIGHT_REBusDelete, "删除车次" & m_szBusID & ToDBDate(m_dtRunDate)
    Exit Sub
here:
     ShowError err.Number
End Sub

'S12
Private Sub UpdateToDB(Optional poDb As RTConnection = Nothing)
    Dim oDb As RTConnection
    Dim sztempsql As String
    Dim szSqlfield As String
    Dim szsqlWhere As String
    Dim szlog As String
    Dim lRow As Long
     AssertHaveRight m_oActiveUser, RIGHT_REBusUpdate
    If poDb Is Nothing Then
        Set oDb = New RTConnection
        oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    Else
        Set oDb = poDb
    End If
On Error GoTo here
    m_dtStartupTime = SelfGetFullDateTime(m_dtRunDate, m_dtStartupTime)
    oDb.BeginTrans
    sztempsql = "UPDATE Env_bus_info SET "
    szSqlfield = "check_gate_id='" & m_szCheckGate & "',"
    szSqlfield = szSqlfield & "bus_start_time='" & ToDBDateTime(m_dtStartupTime) & "',"
    szSqlfield = szSqlfield & "bus_type=" & m_nBusType & ","
    szSqlfield = szSqlfield & "vehicle_type_name='" & m_szVehicleModelName & "',"
    szSqlfield = szSqlfield & "vehicle_type_code='" & m_szVehicleModelID & "',"
    szSqlfield = szSqlfield & "total_seat=" & m_nTotalSeat & ","
    szSqlfield = szSqlfield & "sale_seat_quantity=" & m_nSaleSeat & ","
    szSqlfield = szSqlfield & "total_stand_quantity=" & m_nTotalStandSeat & ","
    szSqlfield = szSqlfield & "sale_stand_seat_quantity=" & m_nSaleStandSeat & ","
    szSqlfield = szSqlfield & "owner_id='" & m_szOwner & "',"
    szSqlfield = szSqlfield & "transport_company_id='" & m_szCompany & "',"
    szSqlfield = szSqlfield & "split_company_id='" & m_szSplitCompanyID & "',"
    szSqlfield = szSqlfield & "vehicle_id='" & m_szVehicle & "',"
    szSqlfield = szSqlfield & "end_station_name='" & m_szEndStationName & "',"
    szSqlfield = szSqlfield & "route_id='" & m_szRoute & "',"
    szSqlfield = szSqlfield & "is_all_refundment=" & IIf(m_bAllRefundment, cbAllRefundment, cbNotAllRefundment) & ","
    szSqlfield = szSqlfield & "delay_status = " & TransFieldValueToString(m_bDelayStatus) & ","
    szSqlfield = szSqlfield & "register_status=" & m_eBeingStatus & ","
    szSqlfield = szSqlfield & "status=" & Str(m_eBusStatus) & ","
    szSqlfield = szSqlfield & "internet_status=" & m_nInternetStatus & ","
    szSqlfield = szSqlfield & "ScrollBus_Check_Time=" & m_nScrollBusCheckTime & ""
  
    szsqlWhere = " WHERE bus_id='" & m_szBusID & "' AND  bus_date='" & ToDBDate(m_dtRunDate) & "'"
    sztempsql = sztempsql & szSqlfield & szsqlWhere
    oDb.Execute sztempsql, lRow
    
    If m_dtStartupTime <> m_dtOldStartupTime Or m_szCheckGate <> m_szOldCheckGate Then    '对应调整配载信息
        sztempsql = "UPDATE Env_bus_allot_lst SET bus_start_time=" & TransFieldValueToString(m_dtStartupTime) & _
                                                    ",check_gate_id=" & TransFieldValueToString(m_szCheckGate) & _
                " WHERE bus_id=" & TransFieldValueToString(m_szBusID) & " AND bus_date=" & TransFieldValueToString(m_dtRunDate) & _
                " AND check_gate_id=" & TransFieldValueToString(m_szOldCheckGate) & ""
        oDb.Execute sztempsql
    End If
    
    oDb.CommitTrans
    
    '为写日志用
    '赋值为写日志用
    
    If m_noldScrollBusCheckTime <> m_nScrollBusCheckTime Then
       szlog = szlog & ",滚动车次循环检票时间修改"
    End If
    If m_nOldInternetStatus <> m_nInternetStatus Then
        If m_nInternetStatus = 0 Then
           szlog = szlog & ",原不可互联售票改为可售"
           m_nOldInternetStatus = m_nInternetStatus
        Else
          szlog = szlog & ",原可互联售票改为不可售"
          m_nOldInternetStatus = m_nInternetStatus
        End If
    End If
    
    If m_szOldCompanyName <> m_szCompanyName Then
        szlog = szlog & ",公司名[" & m_szOldCompanyName & "]改为[" & m_szCompanyName & "]"
        m_szOldCompanyName = m_szCompanyName
    End If
    If m_szRouteName <> m_szOldRouteName Then
        szlog = szlog & ",线路名[" & m_szOldRouteName & "]改为[" & m_szRouteName & "]"
        m_szOldRouteName = m_szRouteName
    End If
    If m_szOldOwnerName <> m_szOwnerName Then
        szlog = szlog & ",车主名[" & m_szOldOwnerName & "]改为[" & m_szOwnerName & "]"
        m_szOldOwnerName = m_szOwnerName
    End If
    If m_szCheckGate <> m_szOldCheckGate Then
        szlog = szlog & ",检票口[" & m_szOldCheckGate & "]改为[" & m_szCheckGate & "]"
        m_szOldCheckGate = m_szCheckGate
    End If
    If m_szOldRoute <> m_szRoute Then
        szlog = szlog & ",线路代码[" & m_szOldRoute & "]改为[" & m_szRoute & "]"
        m_szOldRoute = m_szRoute
    End If
    If m_szVehicleModelID <> m_szOldVehicleModelID Then
        szlog = szlog & ",车型代码[" & m_szOldVehicleModelID & "]改为[" & m_szVehicleModelID & "]"
        m_szOldVehicleModelID = m_szVehicleModelID
    End If
    If m_szOldVehicleModelName <> m_szVehicleModelName Then
        szlog = szlog & ",车型名[" & m_szOldVehicleModelName & "]改为[" & m_szVehicleModelName & "]"
        m_szOldVehicleModelName = m_szVehicleModelName
    End If
    If m_nBusType <> m_eOldBusType Then
        szlog = szlog & ",车次种类[" & m_eOldBusType & "]改为[" & m_nBusType & "]"
        m_eOldBusType = m_nBusType
    End If
    If m_szOldEndStation <> m_szEndStation Then
        szlog = szlog & ",终点站[" & m_szOldEndStation & "]改为[" & m_szEndStation & "]"
        m_szOldEndStation = m_szEndStation
    End If
    If m_szCompany <> m_szOldCompany Then
        szlog = szlog & ",公司代码[" & m_szOldCompany & "]改为[" & m_szCompany & "]"
        m_szOldCompany = m_szCompany
    End If
    If m_szSplitCompanyID <> m_szOldSplitCompanyID Then
        szlog = szlog & ",拆帐公司代码[" & m_szOldSplitCompanyID & "]改为[" & m_szSplitCompanyID & "]"
        m_szOldSplitCompanyID = m_szSplitCompanyID
    End If
    If m_szOldOwner <> m_szOwner Then
        szlog = szlog & ",车主代码[" & m_szOldOwner & "]改为[" & m_szOwner & "]"
        m_szOldOwner = m_szOwner
    End If
    If m_szVehicle <> m_szOldVehicle Then
        szlog = szlog & ",车辆代码[" & m_szOldVehicle & "]改为[" & m_szVehicle & "]"
        m_szOldVehicle = m_szVehicle
    End If
    If m_szOldVehicleTag <> m_szVehicleTag Then
        szlog = szlog & ",车牌号[" & m_szOldVehicleTag & "]改为[" & m_szVehicleTag & "]"
        m_szOldVehicleTag = m_szVehicleTag
    End If
    If m_szEndStationName <> m_szOldEndStationName Then
        szlog = szlog & ",终点站名[" & m_szOldEndStationName & "]改为[" & m_szEndStationName & "]"
        m_szOldEndStationName = m_szEndStationName
    End If
    If m_bOldAllRefundment <> m_bAllRefundment Then
        szlog = szlog & ",全额退票[" & m_bOldAllRefundment & "]改为[" & m_bAllRefundment & "]"
        m_bOldAllRefundment = m_bAllRefundment
    End If
    If m_dtStartupTime <> m_dtOldStartupTime Then
        szlog = szlog & ",发车时间[" & m_dtOldStartupTime & "]改为[" & m_dtStartupTime & "]"
        m_dtOldStartupTime = m_dtStartupTime
    End If
    '是否晚点
    If m_bDelayStatus <> m_bOldDelayStatus Then
        szlog = szlog & ",晚点状态[" & TransFieldValueToString(m_bOldDelayStatus) & "]改为[" & TransFieldValueToString(m_bDelayStatus) & "]"
        m_bOldDelayStatus = m_bDelayStatus
    End If
    If m_eOldBeingStatus <> m_eBeingStatus Then
        szlog = szlog & ",报班状态[" & m_eOldBeingStatus & "]改为[" & m_eBeingStatus & "]"
        m_eOldBeingStatus = m_eBeingStatus
    End If
    
    Dim nBusStatusTemp As Integer
    Dim strBusStatusTemp As String
    Dim strBusStatus As String
    Dim strOldBusStatus As String
    Dim i As Integer
    If m_eBusStatus <> m_eOldBusStatus Then
        For i = 1 To 2
            
            
            If i = 2 Then
              nBusStatusTemp = m_eOldBusStatus
            Else
              nBusStatusTemp = m_eBusStatus
            End If
            Select Case nBusStatusTemp
                   Case 0: strBusStatusTemp = "车次正常"
                   Case 1: strBusStatusTemp = "车次已被停班"
                   Case 2: strBusStatusTemp = "车次已被并班停班"
                   Case 3: strBusStatusTemp = "车次已经停检"
                   Case 4: strBusStatusTemp = "车次正在检票"
                   Case 5: strBusStatusTemp = "车次正在补检"
                   Case 8: strBusStatusTemp = "车次拆分停班"
                   Case 16: strBusStatusTemp = "车次顶班"
            End Select
            If i = 2 Then
              strOldBusStatus = strBusStatusTemp
            Else
              strBusStatus = strBusStatusTemp
            End If
            
        Next
        szlog = szlog & ",车次状态[" & strOldBusStatus & "]改为[" & strBusStatus & "]"
        m_eOldBusStatus = m_eBusStatus
    End If
    
    
    
    If m_nOldSaleSeat <> m_nSaleSeat Then
        szlog = szlog & ",可售座位数[" & m_nOldSaleSeat & "]改为[" & m_nSaleSeat & "]"
        m_nOldSaleSeat = m_nSaleSeat
    End If
    If m_nSaleStandSeat <> m_nOldSaleStandSeat Then
        szlog = szlog & ",可售站票数[" & m_nOldSaleStandSeat & "]改为[" & m_nSaleStandSeat & "]"
        m_nOldSaleStandSeat = m_nSaleStandSeat
    End If
    If m_nOldTotalSeat <> m_nTotalSeat Then
        szlog = szlog & ",总座位数[" & m_nOldTotalSeat & "]改为[" & m_nTotalSeat & "]"
        m_nOldTotalSeat = m_nTotalSeat
    End If
    
    If m_nTotalStandSeat <> m_nOldTotalStandSeat Then
        szlog = szlog & ",总站票数[" & m_nOldTotalStandSeat & "]改为[" & m_nTotalStandSeat & "]"
        m_nOldTotalStandSeat = m_nTotalStandSeat
    End If
    
    '赋值为写日志用
    Dim szBusdateTemp As String
    szBusdateTemp = ToDBDateTime(m_dtOldStartupTime)
    '截取字符串长度
    If szlog <> "" Then
        szlog = "车次[" & m_szBusID & "]发车时间[" & szBusdateTemp & "]终点站[" & m_szOldEndStationName & "]:基本属性修改：" & szlog
        szlog = GetUnicodeBySize(szlog, 255)
        WriteOperateLog m_oActiveUser, RIGHT_REBusUpdate, szlog
    End If
Exit Sub


here:
    oDb.RollbackTrans
    AssertUpdateObjectError ERR_REBus, oDb
End Sub

'S13
'Public Sub UnMergeBus()
'    Dim szsql As String
'    Dim rsTemp As Recordset
'    Dim odb As New rtConnection
'
'    '''''''
'    Dim nSeatType As Integer
'    Dim tSeatTypeInfo(1 To 5) As TSeatInfoCount
'    Dim szbusID As String
'    Dim i As Integer
'    ''''''
'AssertObjIsValid
'
'   AssertHaveRight m_oActiveUser, RIGHT_REBusUnMergeBus
'    odb.ConnectionString = GetConnectionStr(cszRunEnv)
'    On Error GoTo here
'
'    '查询该车次并入何车次
'
'    szsql = "SELECT * FROM Mergebus_split_lst WHERE merge_bus_id='" & m_szBusID & "' AND " _
'            & " bus_date='" & ToDBDate(m_dtRunDate) & "'"
'    Set rsTemp = odb.Execute(szsql)
'
'    If rsTemp.RecordCount = 0 Then ShowError ERR_REBusNoBeMergedBus
'    szbusID =  FormatDbValue(rstemp!bus_id
'    odb.BeginTrans
'
'    '2将车次置为可运行
'    szsql = "UPDATE Env_bus_info SET status=" & Str(ST_BusNormal) & " WHERE bus_id='" & m_szBusID & "' AND " _
'    & " bus_date='" & ToDBDate(m_dtRunDate) & "'"
'    odb.Execute szsql
'
'    '''''''''''''''
'    '取座位信息
'    szsql = " SELECT Count(seat_type_id) as SeatTypeCount ,Seat_type_id FROM Env_bus_seat_lst WHERE  bus_date='" & ToDBDate(m_dtRunDate) & "' AND " _
'                    & " bus_id='" & szbusID & "' AND " _
'                    & " status='" & ST_SeatMerge & "'GROUP BY Seat_type_id"
'    Set rsTemp = odb.Execute(szsql)
'
'    If rsTemp.RecordCount <> 0 Then
'      For i = 0 To rsTemp.RecordCount
'          nSeatType = Val( FormatDbValue(rstemp!seat_type_id)
'          tSeatTypeInfo(nSeatType).seatCount =  FormatDbValue(rstemp!SeatTypeCount
'      Next
'    End If
'    ''''''''''''''
'    szsql = "UPDATE Env_bus_seat_lst SET status=" & Str(ST_SeatCanSell) & "  WHERE bus_id='" & szbusID & "' AND " _
'            & " bus_date='" & ToDBDate(m_dtRunDate) & "' AND status='" & ST_SeatMerge & "'"
'    odb.Execute szsql
'
'    '3删除并班记录
'    szsql = "DELETE Mergebus_split_lst WHERE merge_bus_id='" & m_szBusID & "' AND " _
'            & " bus_date='" & ToDBDate(m_dtRunDate) & "'"
'    odb.Execute szsql
'
'    UpdateEnviromentSeatCountEx szbusID, m_dtRunDate, tSeatTypeInfo, False, odb
'
'    odb.CommitTrans
'
'    #If SUPPORT_MSG = 1 Then
'       Dim oParam As New SystemParam
'       Dim aszValues(1 To 2) As String
'       aszValues(1) = m_szBusID
'       aszValues(2) = ToDBDate(m_dtRunDate)
'       m_oSndMsg.SendMsg oParam.NowDateTime, cszResumeBus, aszValues
'    #End If
'
'
'
'
'    WriteOperateLog m_oActiveUser, RIGHT_REBusUnMergeBus, "车次" & m_szBusID & "取消并班"
'
'    Exit Sub
'here:
'   odb.RollbackTrans
'   ReBusSlipLockEx m_szBusID, False
'   ReBusSlipLockEx szbusID, False
'   Set odb = Nothing
'   Err.Raise Err.Number
'End Sub

'S14
Public Sub ModifyMergeAnnotation(Annotation As String)
    Dim szSql As String
    Dim oDb As New RTConnection
On Error GoTo here
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    szSql = "UPDATE Mergebus_split_lst SET annotation='" & Annotation & "' WHERE merge_bus_id='" & m_szBusID & "' AND " _
    & " bus_date='" & ToDBDate(m_dtRunDate) & "'"
    oDb.Execute szSql
Exit Sub
here:
End Sub

'S15修改车次票价
Public Sub ModifyBusTicket(BusTicketPrice As TRETicketPrice)
    Dim szSql As String
    Dim sgTotal As Double
    Dim oDb As New RTConnection
    Dim k As Integer
    Dim sStr As String
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_REBusModifyBusTicketPrice
    sgTotal = BusTicketPrice.sgBase
    For k = 1 To 15
        sgTotal = BusTicketPrice.asgPrice(k) + sgTotal
    Next
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    szSql = "UPDATE Env_bus_price_lst SET " _
    & "ticket_price_total=" & sgTotal _
    & ",base_carriage=" & BusTicketPrice.sgBase
    For k = 1 To 15
        szSql = szSql & ",price_item_" & k & "=" & BusTicketPrice.asgPrice(k)
    Next
    
    szSql = szSql _
    & " WHERE bus_id='" & m_szBusID & "' AND  bus_date='" & ToDBDate(m_dtRunDate) & "'" _
        & " AND station_id='" & BusTicketPrice.szStationID & "'  " _
        & " AND ticket_type = " & BusTicketPrice.nTicketType & "  " _
        & " AND seat_type_id= '" & BusTicketPrice.szSeatType & " '" _
        & " AND sell_station_id = " & TransFieldValueToString(BusTicketPrice.szSellStationID)
        
    oDb.Execute szSql
    '修改该站的对应价票价格
    If BusTicketPrice.nTicketType = TP_FullPrice Then
        sStr = "full_price"
    ElseIf BusTicketPrice.nTicketType = TP_HalfPrice Then
        sStr = "half_price"
    ElseIf BusTicketPrice.nTicketType = TP_PreferentialTicket1 Then
        sStr = "preferential_ticket1"
    ElseIf BusTicketPrice.nTicketType = TP_PreferentialTicket2 Then
        sStr = "preferential_ticket2"
    ElseIf BusTicketPrice.nTicketType = TP_PreferentialTicket3 Then
        sStr = "preferential_ticket3"
    End If
        szSql = "UPDATE Env_bus_station_lst SET " & sStr & "=" & sgTotal _
        & " WHERE bus_id='" & m_szBusID & "' AND  bus_date='" & ToDBDate(m_dtRunDate) & "'" _
        & " AND station_id='" & BusTicketPrice.szStationID & "'" _
        & " AND sell_station_id = " & TransFieldValueToString(BusTicketPrice.szSellStationID)
    oDb.Execute szSql
    WriteOperateLog m_oActiveUser, RIGHT_REBusModifyBusTicketPrice, "编辑车次" & Trim(m_szBusID) & "站点" & Trim(BusTicketPrice.szStationID) _
            & "上车站" & Trim(BusTicketPrice.szSellStationID) & "的票价" & sStr & "改为:" & sgTotal
    #If SUPPORT_MSG = 1 Then
        Dim oParam As New SystemParam
        Dim aszValues(1 To 2) As String
        aszValues(1) = m_szBusID
        aszValues(2) = ToDBDate(m_dtRunDate)
        m_oSndMsg.SendMsg oParam.NowDateTime, cszModifyBusPrice, aszValues
    #End If
End Sub
''S16编辑车次站点的限售张数和限售时间
'Public Sub ModifyBusStation(BSInfo As TBusStationSellInfo)
'    Dim szSql As String
'    Dim oDb As New RTConnection
'    Dim rsTemp As New Recordset
'    Dim nOldMaxSaleQuantity As Integer, sgOldStopSaleTime As Single
'    Dim szStationName As String
'    Dim szlog As String
'    Dim szBusdateTemp As String
'    Dim bflg As Boolean
'AssertObjIsValid
'
'     AssertHaveRight m_oActiveUser, RIGHT_REBusModifyBusStation
'
'     oDb.ConnectionString = GetConnectionStr(cszRunEnv)
'
'     szSql = "SELECT e.max_sale_quantity,e.stop_sale_time ,s.station_name FROM Env_bus_station_lst e,station_info s WHERE  bus_id='" & m_szBusID & "' AND  e.bus_date='" & ToDBDate(m_dtRunDate) & "' AND e.station_id='" & BSInfo.szStationID & "'AND e.station_id=s.station_id"
'
'     Set rsTemp = oDb.Execute(szSql)
'        nOldMaxSaleQuantity =  FormatDbValue(rstemp!max_sale_quantity
'        sgOldStopSaleTime =  FormatDbValue(rstemp!stop_sale_time
'        szStationName =  FormatDbValue(rstemp!station_name
'        szlog = "站点[" & szStationName & "]："
'
'     If nOldMaxSaleQuantity <> BSInfo.nLimitedSellCount Then
'        If nOldMaxSaleQuantity = -1 Then
'          If BSInfo.nLimitedSellCount = 0 Then
'             szlog = szlog & "限售张数由[不限]改为[不可售]"
'          Else
'             szlog = szlog & "限售张数由[不限]改为[" & BSInfo.nLimitedSellCount & "]张"
'           End If
'        Else
'          If BSInfo.nLimitedSellCount = 0 Then
'             szlog = szlog & "限售张数由[" & nOldMaxSaleQuantity & "]改为[不可售]"
'          Else
'             szlog = szlog & "限售张数由[" & nOldMaxSaleQuantity & "]改为[" & BSInfo.nLimitedSellCount & "]张"
'          End If
'
'        End If
'        bflg = True
'      End If
'
'
'
'     If sgOldStopSaleTime <> BSInfo.nLimitedSellTime Then
'
'         If BSInfo.nLimitedSellTime <> -1 Then
'
'              If nOldMaxSaleQuantity = -1 Then
'                    szlog = szlog & "限售时间由[不限]改为发车前[" & BSInfo.nLimitedSellTime & "]小时可售"
'              Else
'                    szlog = szlog & "限售时间由[" & sgOldStopSaleTime / 60 & "]改为[" & BSInfo.nLimitedSellTime / 60 & "]小时可售"
'              End If
'
'         Else
'
'              szlog = szlog & "限售时间由[" & sgOldStopSaleTime / 60 & "]改为发车时间[不限]售"
'
'
'        End If
'       bflg = True
'      End If
'
'    If bflg = True Then
'
'     szSql = "UPDATE Env_bus_station_lst SET max_sale_quantity=" & BSInfo.nLimitedSellCount & ",stop_sale_time=" & BSInfo.nLimitedSellTime _
'             & ",sale_ticket_quantity=" & BSInfo.nLimitedSellCount & "  WHERE bus_id='" & m_szBusID & "' AND  bus_date='" & ToDBDate(m_dtRunDate) & "' AND station_id='" & BSInfo.szStationID & "'"
'     oDb.Execute szSql
'    szBusdateTemp = ToDBDateTime(m_dtOldStartupTime)
'    szBusdateTemp = "车次[" & m_szBusID & "]发车时间[" & szBusdateTemp & "]终点站[" & m_szOldEndStationName & "]:"
'    szlog = szBusdateTemp & szlog
'    szlog = GetUnicodeBySize(szlog, 255)
'    WriteOperateLog m_oActiveUser, RIGHT_REBusModifyBusStation, szlog
'   End If
'End Sub
Public Sub ModifyBusStation(BSInfo As TBusStationSellInfo)
'修改车次站点的最大可售张数及限售时间
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As New Recordset
    Dim nOldMaxSaleQuantity As Integer, sgOldStopSaleTime As Single
    Dim szStationName As String
    Dim szlog As String
    Dim bflg As Boolean
    AssertObjIsValid
    
    AssertHaveRight m_oActiveUser, RIGHT_REBusModifyBusStation
    
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    '得到该车次站点的的最大可售张数及限售时间
    szSql = "SELECT e.max_sale_quantity,e.stop_sale_time ,s.station_name" _
        & " FROM Env_bus_station_lst e,station_info s" _
        & " WHERE  bus_id='" & m_szBusID & "' AND  e.bus_date='" & ToDBDate(m_dtRunDate) _
        & "' AND e.station_id='" & BSInfo.szStationID & "'AND e.station_id=s.station_id"
    Set rsTemp = oDb.Execute(szSql)
    nOldMaxSaleQuantity = FormatDbValue(rsTemp!max_sale_quantity)
    sgOldStopSaleTime = FormatDbValue(rsTemp!stop_sale_time)
    szStationName = FormatDbValue(rsTemp!station_name)
    szlog = "站点[" & szStationName & "]"
    
    If nOldMaxSaleQuantity <> BSInfo.nLimitedSellCount Then
        If nOldMaxSaleQuantity = -1 Then
            If BSInfo.nLimitedSellCount = 0 Then
                szlog = szlog & "限售张数由[不限]改为[不可售]"
            Else
                szlog = szlog & "限售张数由[不限]改为[" & BSInfo.nLimitedSellCount & "]张"
            End If
        Else
            If BSInfo.nLimitedSellCount = 0 Then
                szlog = szlog & "限售张数由[" & nOldMaxSaleQuantity & "]改为[不可售]"
            Else
                szlog = szlog & "限售张数由[" & nOldMaxSaleQuantity & "]改为[" & BSInfo.nLimitedSellCount & "]张"
            End If
        
        End If
        bflg = True
    End If
    
    
    
    If sgOldStopSaleTime <> BSInfo.sgLimitedSellTime Then
        
        If BSInfo.sgLimitedSellTime <> -1 Then
        
            If nOldMaxSaleQuantity = -1 Then
                szlog = szlog & "限售时间由[不限]改为发车前[" & BSInfo.sgLimitedSellTime & "]小时可售"
            Else
                szlog = szlog & "限售时间由[" & sgOldStopSaleTime & "]改为[" & BSInfo.sgLimitedSellTime & "]小时可售"
            End If
        Else
            szlog = szlog & "限售时间由[" & sgOldStopSaleTime & "]改为发车时间[不限]售"
        
        
        End If
        bflg = True
    End If
    
    If bflg = True Then
        Dim szBusdateTemp As String
        
        szBusdateTemp = ToDBDateTime(m_dtOldStartupTime)
        
        szBusdateTemp = "车次[" & m_szBusID & "]发车时间[" & szBusdateTemp & "]终点站[" & m_szOldEndStationName & "]:"
        szlog = szBusdateTemp & szlog
        szlog = GetUnicodeBySize(szlog, 255)
        szSql = "UPDATE Env_bus_station_lst SET max_sale_quantity=" & BSInfo.nLimitedSellCount & ",stop_sale_time=" & BSInfo.sgLimitedSellTime _
        & ",sale_ticket_quantity=" & BSInfo.nLimitedSellCount & "  WHERE bus_id='" & m_szBusID & "' AND  bus_date='" & ToDBDate(m_dtRunDate) & "' AND station_id='" & BSInfo.szStationID & "'"
        oDb.Execute szSql
        WriteOperateLog m_oActiveUser, RIGHT_REBusModifyBusStation, szlog
    
    End If
End Sub
'S17
'顶班
Public Sub ReplaceVehicle(RVehicleID As String, Optional AutoTicketPrice As Boolean = False, Optional AutoSeat As Boolean = False, Optional AllRefundment As Boolean = False)
    Dim szSql As String
    Dim rsNewBusPrice As New Recordset
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szNewVehicleModel As String
    Dim szNewVehicleModelName As String
    Dim szNewVehicleBusOwnerID As String
    Dim szNewVehicleCompanyID As String
    Dim szNewSplitCompanyID As String
    Dim nNewVehicleSeatStart As Integer
    Dim nNewVehicleSeatCount As Integer
    Dim szNewLicenseTagNo As String
    Dim szNewTransportCompanyShortName As String
    
    Dim szTemp As String, nTemp As Integer
    Dim nStartSeat As Integer
    Dim i As Integer
    Dim LerrCode As Long
    Dim nTotalSeat As Integer
    Dim nSaleSeatQuantity As Integer
    Dim szSeatNo() As String
    Dim rsTempSaleSeatNo As New Recordset
    Dim n0ldCountSeatNo As Integer
    Dim atSATemp As TSchemeArrangement
    Dim bCopyEnGetBus As Boolean
    
    Dim tSeatTypeInfo(1 To 5) As TSeatInfoCount
    Dim nSeatType As Integer
    
    Dim szPriceTable As String
    Dim rstempSeat As New Recordset
    Dim szSeatWhere As String
    Dim nSeatCountVehicl As Integer
    Dim nSeatCountTemp  As String
    Dim bIsNormBusSeatType As Boolean
    Dim tvTemp() As TVehcileSeatType
    Dim nCountSeaType As Integer
    
    Dim lgRow As Long
    
    AssertObjIsValid
    AssertActiveUserValid m_oActiveUser, ERR_REBus
    AssertHaveRight m_oActiveUser, RIGHT_REBusReplaceBus
    AssertBusCanBeReplaced
    
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
  '  atSATemp = SelfGetExecuteBusProject(m_dtRunDate, LerrCode)
    
    If LerrCode <> 0 Then ShowError ERR_REBusNotExecuteProject
    
    szSql = "SELECT * FROM bus_info WHERE  bus_id='" & m_szBusID & "'"
    Set rsTemp = oDb.Execute(szSql)
    
    If rsTemp.RecordCount = 0 Then bCopyEnGetBus = True '车次由环境拷贝产生
    
    szSql = "SELECT v.*,vtc.vehicle_type_short_name,c.transport_company_short_name " _
        & " FROM Vehicle_info v,vehicle_type_code vtc,company_info c WHERE " _
        & " vehicle_id='" & RVehicleID & "' AND " _
        & " v.vehicle_type_code=vtc.vehicle_type_code AND v.transport_company_id = c.transport_company_id "
    Set rsTemp = oDb.Execute(szSql)
    
    '新车辆基本信息
    If rsTemp.RecordCount = 0 Then ShowError ERR_REBusVehicelNotExist
        
    szNewVehicleModel = FormatDbValue(rsTemp!vehicle_type_code)
    szNewVehicleModelName = FormatDbValue(rsTemp!vehicle_type_short_name)
    szNewVehicleBusOwnerID = FormatDbValue(rsTemp!owner_id)
    szNewVehicleCompanyID = FormatDbValue(rsTemp!transport_company_id)
    
    szNewSplitCompanyID = FormatDbValue(rsTemp!split_company_id)
    nNewVehicleSeatStart = FormatDbValue(rsTemp!start_seat_no)
    nNewVehicleSeatCount = FormatDbValue(rsTemp!seat_quantity)
    
    szNewLicenseTagNo = FormatDbValue(rsTemp!license_tag_no)
    szNewTransportCompanyShortName = FormatDbValue(rsTemp!transport_company_short_name)
    
    
    szPriceTable = BusProjectExecutePrice(m_dtRunDate, LerrCode)
    If LerrCode = 1 Then ShowError ERR_NoExePriceTable
    
    szSql = "SELECT * FROM vehicle_Seat_type_info WHERE vehicle_id='" & RVehicleID & "'"
    Set rstempSeat = oDb.Execute(szSql)
    '新车辆座位类型及该座位类型的数目
    If rstempSeat.RecordCount <> 0 Then
    
        For i = 1 To rstempSeat.RecordCount
            nSeatType = Val(rstempSeat!seat_type_id)
            nSeatCountTemp = Val(rstempSeat!end_seat_no) - Val(rstempSeat!start_seat_no) + 1
            tSeatTypeInfo(nSeatType).SeatTypeID = rstempSeat!seat_type_id
            tSeatTypeInfo(nSeatType).seatCount = tSeatTypeInfo(nSeatType).seatCount + nSeatCountTemp
            nSeatCountVehicl = nSeatCountVehicl + nSeatCountTemp
            rstempSeat.MoveNext
        Next
    
    End If
    If nSeatCountVehicl < nNewVehicleSeatCount Then
    
        If tSeatTypeInfo(1).seatCount = 0 Then
            bIsNormBusSeatType = True
        End If
        
        tSeatTypeInfo(1).seatCount = tSeatTypeInfo(1).seatCount + nNewVehicleSeatCount - nSeatCountVehicl
        tSeatTypeInfo(1).SeatTypeID = cszSeatTypeIsNormal
    
    End If
    nCountSeaType = 0
    For i = 1 To 5
        If tSeatTypeInfo(i).seatCount <> 0 Then
            szSeatWhere = szSeatWhere & " '" & tSeatTypeInfo(i).SeatTypeID & "',"
            nCountSeaType = nCountSeaType + 1
        End If
    Next
    
    szSeatWhere = Left(szSeatWhere, Len(szSeatWhere) - 1)
    
    Set rstempSeat = Nothing
    
        
    '************************************************************************
    
    '车型不一致则要求更新票价,得到新车型的线路票价信息
    If AutoTicketPrice Then
    
        '由车次种类，车型的到票价的车次
        Dim szmyBusID As String
        Dim bHavePrice As Boolean
        Dim bMakePrice As Boolean
        Dim szWhere As String
    
        If m_szBusCopy = "" Then
            m_szBusCopy = m_szBusID
            szmyBusID = m_szBusID
            bMakePrice = False
        Else
            '界面上已生成票价车次
            bMakePrice = True
            szmyBusID = m_szBusCopy
        End If
        '如果需要生成票价
        If bMakePrice = False Then
            If Trim(m_szVehicle) = Trim(RVehicleID) And bCopyEnGetBus = False Then  '同车辆
                szWhere = " AND bus_id='" & m_szBusID & " '"
            End If
            ' 同车型，同线路，同座位，的车次的票价信息
            '如果是有多个起点站车次的车次
            
'            szSql = "SELECT TOP 1 bus_id FROM" _
'                & "(" _
'                & " SELECT bus_id ,COUNT(*) as count_seat_type FROM " _
'                & "    ( " _
'                & "     SELECT DISTINCT bus_id,seat_type_id  FROM bus_price_lst WHERE " _
'                & "     bus_id in " _
'                & "           (" _
'                & "                SELECT DISTINCT bus_id FROM bus_price_lst WHERE price_table_id='" & szPriceTable & "'" _
'                & "                 AND seat_type_id in(" & szSeatWhere & " ) AND  vehicle_type_code='" & szNewVehicleModel & "'" _
'                & "                 AND  bus_id in(SELECT bus_id FROM bus_info WHERE route_id='" & m_szRoute & "'AND project_id='" & atSATemp.szProjectID & "' " & szWHERE & " )    " _
'                & "            )" _
'                & "    )a" _
'                & " GROUP BY a.bus_id" _
'                & ")b" _
'                & "  WHERE b.count_seat_type = " & nCountSeaType & ""
    
    
            szSql = "SELECT TOP 1 bus_id FROM" _
                & "(" _
                & " SELECT bus_id ,COUNT(*) as count_seat_type FROM " _
                & "    ( " _
                & "     SELECT DISTINCT bus_id,seat_type_id  FROM bus_price_lst WHERE " _
                & "     bus_id in " _
                & "           (" _
                & "                SELECT DISTINCT bus_id FROM bus_price_lst WHERE price_table_id='" & szPriceTable & "'" _
                & "                 AND seat_type_id in(" & szSeatWhere & " ) AND  vehicle_type_code='" & szNewVehicleModel & "'" _
                & "                 AND  bus_id in(SELECT bus_id FROM bus_info WHERE route_id='" & m_szRoute & "'" & szWhere & " )    " _
                & "            )" _
                & "    )a" _
                & " GROUP BY a.bus_id" _
                & ")b" _
                & "  WHERE b.count_seat_type = " & nCountSeaType & ""
            Set rsNewBusPrice = oDb.Execute(szSql)
            
            If rsNewBusPrice.RecordCount = 0 Then
            
                If bCopyEnGetBus = False Then ShowError ERR_REBusNewVehicleNowRoutePrice '不是加班车，界面上捕捉该错误
                
'                '同车型，同线路，同座位，的车次不存在  取出同线路车次（m_szBusCopy）由界面生成车次票价
'                szSql = "SELECT Max(Bus_id) as Bus_id FROM bus_info WHERE project_id='" & atSATemp.szProjectID & "'  " _
'                    & " AND  route_id='" & m_szRoute & " ' "
                szSql = "SELECT Max(Bus_id) as Bus_id FROM bus_info WHERE  route_id='" & m_szRoute & " ' "
                Set rsNewBusPrice = oDb.Execute(szSql)
                m_szBusCopy = rsNewBusPrice!bus_id '界面上利用该车次生成新车型车次票价
                Set rsNewBusPrice = Nothing
                ShowError ERR_REBusNewVehicleNowRoutePrice '界面上捕捉该错误
            Else
                '存在车次 rsNewBusPrice!bus_id 票价
                m_szBusCopy = rsNewBusPrice!bus_id
                szmyBusID = m_szBusCopy
                
            End If
                
        End If
    
    
        szSql = "SELECT  bus_id  FROM  bus_price_lst WHERE  bus_id='" & szmyBusID & "' AND price_table_id='" & szPriceTable & "' AND vehicle_type_code='" & szNewVehicleModel & "'AND  seat_type_id In (" & szSeatWhere & " )"
        
        Set rsNewBusPrice = oDb.Execute(szSql)
        If rsNewBusPrice.RecordCount = 0 Then ShowError ERR_REBusNewVehicleNowRoutePrice
    
        
        
    End If
    m_szBusCopy = ""
    bMakePrice = False
    Set rsNewBusPrice = Nothing
    
    
    '*******************************************************************************************
    '事务处理开始
    oDb.BeginTrans
    On Error GoTo here
    '更新环境票价表
    '更新环境站点表
    If AutoTicketPrice Then
        Dim rstemps As Recordset
        Dim szHalfPrice As String
        Dim szPriceItem As String
        Dim j As Integer
        Dim szoldStationId() As String
        Dim bbookflg As Boolean
        Dim szWhereStationID As String
        '保留原车次的预定信息
        szSql = "SELECT * FROM Work_env_bus_station_lst WHERE book_count>0  AND bus_id ='" & m_szBusID & "' AND bus_date='" & ToDBDate(m_dtRunDate) & "'"
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount <> 0 Then
            bbookflg = True
            For i = 1 To rsTemp.RecordCount
                szoldStationId(i, 1) = " &  FormatDbValue(rstemp!book_count & " '"
                szoldStationId(i, 2) = " &  FormatDbValue(rstemp!station_id & " '"
            Next
        End If
        '删除环境站点票价，环境票价表
        If bCopyEnGetBus = False Then
            '计划存在车次，直接 delete
            szSql = "DELETE FROM Env_bus_station_lst WHERE bus_id ='" & m_szBusID & "' AND bus_date='" & ToDBDate(m_dtRunDate) & "'"
            oDb.Execute szSql
        Else
            '计划不存在车次.
            
            '环境拷贝产生车次------保存可售站
            szSql = "SELECT DISTINCT station_id  FROM Work_env_bus_station_lst  WHERE bus_id ='" & m_szBusID & "' AND  " _
                & " bus_date='" & ToDBDate(m_dtRunDate) & "'AND max_sale_quantity=-1"
            Set rsTemp = oDb.Execute(szSql)
            If rsTemp.RecordCount <> 0 Then
                For j = 1 To rsTemp.RecordCount
                    szWhereStationID = szWhereStationID & "'" & FormatDbValue(rsTemp!station_id) & "',"
                    rsTemp.MoveNext
                Next
                szWhereStationID = Left(szWhereStationID, Len(szWhereStationID) - 1)
            End If
            
            szSql = "DELETE FROM Env_bus_station_lst WHERE bus_id ='" & m_szBusID & "' AND bus_date='" & ToDBDate(m_dtRunDate) & "'"
            oDb.Execute szSql
        End If
        
        szSql = "DELETE FROM Env_bus_price_lst WHERE bus_id ='" & m_szBusID & "' AND bus_date='" & ToDBDate(m_dtRunDate) & "'"
        oDb.Execute szSql
        Dim szSeatTypeTemp As String
        Dim nCountSeatType As Integer
        Dim rsTemp1 As New Recordset
        szSql = "SELECT DISTINCT seat_type_id  FROM vehicle_Seat_type_info WHERE vehicle_id='" & RVehicleID & "' ORDER BY seat_type_id"
        Set rsTemp1 = oDb.Execute(szSql)
        nCountSeatType = rsTemp1.RecordCount
        If bIsNormBusSeatType = True Then
            nCountSeatType = nCountSeatType + 1
        End If
        For j = 1 To nCountSeatType
        
            If bIsNormBusSeatType = True Then
                szSeatTypeTemp = cszSeatTypeIsNormal
            End If
            
            If bIsNormBusSeatType = False Then
                szSeatTypeTemp = Trim(rsTemp1!seat_type_id)
            End If
            '写入环境车次站点表
'            szSql = " INSERT Env_bus_station_lst " _
'                & " ( " _
'                & " bus_date , bus_id, sell_station_id,station_id, seat_type_id, station_serial_no, max_sale_quantity, sale_ticket_quantity, book_count " _
'                & " , stop_sale_time , mileage , full_price , half_price , preferential_ticket1 , preferential_ticket2 , preferential_ticket3 " _
'                & " ) " _
'                & " SELECT " _
'                & " '" & ToDBDate(m_dtRunDate) & "'as bus_date , '" & m_szBusID & " ' as bus_id , g.station_id  , '" & m_szSellStationID & "' as sell_station_id,'" & szSeatTypeTemp & "' ,station_serial_no , max_sale_quantity , max_sale_quantity , max_sale_quantity " _
'                & " , stop_sale_time , mileage , ISNULL(f.ticket_price_total,0) , ISNULL(h.ticket_price_total,0) , ISNULL(p1.ticket_price_total,0) , ISNULL(p2.ticket_price_total,0) , ISNULL(p3.ticket_price_total,0) "
'            szSql = szSql & " FROM " _
'                & " ( SELECT project_id ,'" & m_szBusID & " ' as bus_id , station_id, station_serial_no, max_sale_quantity ,stop_sale_time ,mileage " _
'                & "  FROM Bus_station_info WHERE bus_id='" & szmyBusID & "'  AND project_id='" & Trim(atSATemp.szProjectID) & "' " _
'                & " ) g , " _
'                & " ( SELECT '" & m_szBusID & " ' as bus_id , station_id  , base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
'                & "     + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
'                & "     + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
'                & " WHERE ticket_type = " & TP_FullPrice & " AND price_table_id = '" & Trim(szPriceTable) & "' AND vehicle_type_code = '" & szNewVehicleModel & "' AND bus_id = '" & szmyBusID & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' " _
'                & "     ) f , " _
'                & " ( SELECT  '" & m_szBusID & " '  as bus_id , station_id , base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
'                & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
'                & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
'                & " WHERE ticket_type = " & TP_HalfPrice & "  AND price_table_id = '" & Trim(szPriceTable) & "' AND vehicle_type_code = '" & szNewVehicleModel & "' AND bus_id = '" & szmyBusID & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' " _
'                & " ) h , " _
'                & " ( SELECT  '" & m_szBusID & " '  as bus_id , station_id , base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
'                & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
'                & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
'                & " WHERE ticket_type = " & TP_PreferentialTicket1 & "  AND price_table_id = '" & Trim(szPriceTable) & "' AND vehicle_type_code = '" & szNewVehicleModel & "' AND bus_id = '" & szmyBusID & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' " _
'                & " ) p1 , " _
'                & " ( SELECT  '" & m_szBusID & " '  as bus_id , station_id , base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 "
'
'            szSql = szSql _
'                & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
'                & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
'                & " WHERE ticket_type = " & TP_PreferentialTicket2 & "  AND price_table_id = '" & Trim(szPriceTable) & "' AND vehicle_type_code = '" & szNewVehicleModel & "' AND bus_id = '" & szmyBusID & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' " _
'                & " ) p2 , " _
'                & " ( SELECT   '" & m_szBusID & " '  as bus_id , station_id , base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
'                & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
'                & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
'                & " WHERE ticket_type = " & TP_PreferentialTicket3 & "  AND price_table_id = '" & Trim(szPriceTable) & "' AND vehicle_type_code = '" & szNewVehicleModel & "' AND bus_id = '" & szmyBusID & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' " _
'                & " ) p3 " _
'                & " WHERE g.project_id= '" & Trim(atSATemp.szProjectID) & "' AND g.bus_id = '" & m_szBusID & " ' " _
'                & " AND g.bus_id *= f.bus_id AND g.bus_id *= h.bus_id AND g.bus_id *= p1.bus_id AND g.bus_id *= p2.bus_id AND g.bus_id *=p3.bus_id " _
'                & " AND g.station_id *= f.station_id AND g.station_id *= h.station_id AND g.station_id *= p1.station_id AND g.station_id *= p2.station_id AND g.station_id *=p3.station_id "
'
'            oDb.Execute szSql
'            '写入环境车次票价表
'            szSql = "INSERT Env_bus_price_lst " _
'                & " (  " _
'                & " bus_date , bus_id ,sell_station_id,station_id, ticket_type, seat_type_id, station_serial_no, ticket_price_total, base_carriage " _
'                & " , price_item_1 , price_item_2 , price_item_3 , price_item_4 , price_item_5 , price_item_6 , price_item_7 , price_item_8 " _
'                & " , price_item_9 , price_item_10 , price_item_11 , price_item_12 , price_item_13 , price_item_14 , price_item_15 " _
'                & " ) " _
'                & " SELECT '" & ToDBDate(m_dtRunDate) & "' , " _
'                & " '" & m_szBusID & " ' as   bus_id ,sell_station_id,station_id, ticket_type, seat_type_id, station_serial_no " _
'                & " , base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
'                    & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13" _
'                    & " + price_item_14 + price_item_15 ticket_price_total " _
'                & " , base_carriage , price_item_1 , price_item_2 , price_item_3 , price_item_4 , price_item_5 , price_item_6 " _
'                & " , price_item_7 , price_item_8 , price_item_9 , price_item_10 , price_item_11 , price_item_12 , price_item_13 " _
'                & " , price_item_14 , price_item_15 " _
'                & " FROM bus_price_lst " _
'                & " WHERE price_table_id = '" & Trim(szPriceTable) & "' AND vehicle_type_code = '" & szNewVehicleModel & "' AND bus_id = '" & szmyBusID & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' "
'            oDb.Execute szSql
'
            '写入环境车次站点表
            szSql = " INSERT Env_bus_station_lst " _
                & " ( " _
                & " bus_date , bus_id, station_id, sell_station_id,  seat_type_id, station_serial_no, max_sale_quantity, sale_ticket_quantity, book_count " _
                & " , stop_sale_time , mileage , full_price , half_price , preferential_ticket1 , preferential_ticket2 , preferential_ticket3 " _
                & " ) " _
                & " SELECT " _
                & " '" & ToDBDate(m_dtRunDate) & "' , '" & m_szBusID & " ' as bus_id , g.station_id , s.sell_station_id  , '" & szSeatTypeTemp & "' ,station_serial_no , max_sale_quantity , max_sale_quantity , 0 " _
                & " , stop_sale_time , end_station_mileage , ISNULL(f.ticket_price_total,0) , ISNULL(h.ticket_price_total,0) , ISNULL(p1.ticket_price_total,0) , ISNULL(p2.ticket_price_total,0) , ISNULL(p3.ticket_price_total,0) " _
                & " FROM Bus_station_info g ,  Route_section_lst s ,  " _
                & " ( SELECT bus_id , station_id   , sell_station_id, base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
                & "     + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
                & "     + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
                & " WHERE ticket_type = " & TP_FullPrice & " AND price_table_id = '" & Trim(szPriceTable) & "' AND vehicle_type_code = '" & szNewVehicleModel & "' AND bus_id = '" & szmyBusID & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' " _
                & "     ) f , " _
                & " ( SELECT bus_id , station_id  , sell_station_id, base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
                & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
                & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
                & " WHERE ticket_type = " & TP_HalfPrice & "  AND price_table_id = '" & Trim(szPriceTable) & "' AND vehicle_type_code = '" & szNewVehicleModel & "' AND bus_id = '" & szmyBusID & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' " _
                & " ) h , " _
                & " ( SELECT bus_id , station_id  , sell_station_id, base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
                & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
                & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
                & " WHERE ticket_type = " & TP_PreferentialTicket1 & "  AND price_table_id = '" & Trim(szPriceTable) & "' AND vehicle_type_code = '" & szNewVehicleModel & "' AND bus_id = '" & szmyBusID & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' " _
                & " ) p1 , " _
                & " ( SELECT bus_id , station_id  , sell_station_id, base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 "
    
            szSql = szSql _
                & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
                & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
                & " WHERE ticket_type = " & TP_PreferentialTicket2 & "  AND price_table_id = '" & Trim(szPriceTable) & "' AND vehicle_type_code = '" & szNewVehicleModel & "' AND bus_id = '" & szmyBusID & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' " _
                & " ) p2 , " _
                & " ( SELECT bus_id , station_id  , sell_station_id, base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
                & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
                & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
                & " WHERE ticket_type = " & TP_PreferentialTicket3 & "  AND price_table_id = '" & Trim(szPriceTable) & "' AND vehicle_type_code = '" & szNewVehicleModel & "' AND bus_id = '" & szmyBusID & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' " _
                & " ) p3 " _
                & " WHERE g.bus_id = '" & szmyBusID & "' " _
                & " AND g.bus_id *= f.bus_id AND g.bus_id *= h.bus_id AND g.bus_id *= p1.bus_id AND g.bus_id *= p2.bus_id AND g.bus_id *=p3.bus_id " _
                & " AND g.station_id *= f.station_id AND g.station_id *= h.station_id AND g.station_id *= p1.station_id AND g.station_id *= p2.station_id AND g.station_id *=p3.station_id "
            
            szSql = szSql _
                & " AND g.station_id = s.end_station_id AND s.route_id = '" & m_szRoute & "' AND s.sell_station_id *= f.sell_station_id AND s.sell_station_id *= h.sell_station_id " _
                & " AND s.sell_station_id *= p1.sell_station_id AND s.sell_station_id *= p2.sell_station_id AND s.sell_station_id *= p3.sell_station_id"
            
            oDb.Execute szSql, lgRow
            '环境车次票价
            szSql = "INSERT Env_bus_price_lst " _
                & " (  " _
                & " bus_date , bus_id,sell_station_id, station_id, ticket_type, seat_type_id, station_serial_no, ticket_price_total, base_carriage " _
                & " , price_item_1 , price_item_2 , price_item_3 , price_item_4 , price_item_5 , price_item_6 , price_item_7 , price_item_8 " _
                & " , price_item_9 , price_item_10 , price_item_11 , price_item_12 , price_item_13 , price_item_14 , price_item_15 " _
                & " ) " _
                & " SELECT '" & ToDBDate(m_dtRunDate) & "' , " _
                & " '" & m_szBusID & " ' as bus_id ,  sell_station_id,station_id, ticket_type, seat_type_id, station_serial_no " _
                & " , base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
                & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13" _
                & " + price_item_14 + price_item_15 ticket_price_total " _
                & " , base_carriage , price_item_1 , price_item_2 , price_item_3 , price_item_4 , price_item_5 , price_item_6 " _
                & " , price_item_7 , price_item_8 , price_item_9 , price_item_10 , price_item_11 , price_item_12 , price_item_13 " _
                & " , price_item_14 , price_item_15 " _
                & " FROM bus_price_lst " _
                & " WHERE price_table_id = '" & Trim(szPriceTable) & "' AND vehicle_type_code = '" & szNewVehicleModel & "' AND bus_id = '" & szmyBusID & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' "
            oDb.Execute szSql


            If bIsNormBusSeatType = True Then
                bIsNormBusSeatType = False
            Else
                rsTemp1.MoveNext
            End If
                    
        Next
            
        '站点可售
        If bCopyEnGetBus = True Then
            '环境复制得到
            If szWhereStationID <> "" Then
                szSql = "UPDATe Env_bus_station_lst SET max_sale_quantity=0 WHERE station_id  not In( " & szWhereStationID & ") AND " _
                    & "  bus_id='" & m_szBusID & "' AND  bus_date='" & ToDBDate(m_dtRunDate) & " ' "
                oDb.Execute (szSql)
            End If
        
        Else
            '计划中存在的车次
            szSql = "UPDATe Env_bus_station_lst SET max_sale_quantity=0 WHERE station_id not In( SELECT station_id FROM Bus_station_info WHERE   " _
                & "   bus_id='" & m_szBusID & "' AND  " _
                & "  max_sale_quantity= -1) AND " _
                & "  bus_id='" & m_szBusID & "' AND  bus_date='" & ToDBDate(m_dtRunDate) & " ' "
            oDb.Execute (szSql)
        End If
        
        
        '恢复预定信息
        Dim nCount As Integer
        If bbookflg = True Then
            nCount = ArrayLength(szoldStationId)
            For i = 1 To nCount
                szSql = "update Env_bus_station_lst set book_count=" & CInt(szoldStationId(i, 1)) & " WHERE bus_id='" & m_szBusID & "' " _
                    & "  AND  bus_date ='" & ToDBDate(m_dtRunDate) & "' AND station_id='" & szoldStationId(i, 2) & "'AND status='" & SE_SeatBooked & "'"
                oDb.Execute szSql
            Next
        End If
    End If
    '座位处理
    If AutoSeat Then
        UpdateEnvSeatInfo RVehicleID, nNewVehicleSeatCount, nNewVehicleSeatStart, m_szBusID, m_dtRunDate, oDb, True
    End If

    '如果要设置全额退票标志,则设置之
    
    If AllRefundment Then
        m_bAllRefundment = True
    End If
    
    ' 写日志用的
    m_szVehicle = m_szVehicle
    m_szOldVehicleModelID = m_szVehicleModelID
    m_szOldVehicleModelName = m_szVehicleModelName
    m_szOldOwner = m_szOwner
    m_szOldCompany = m_szCompany
    m_szOldSplitCompanyID = m_szSplitCompanyID
    m_eOldBusStatus = ST_BusReplace
    m_nOldTotalSeat = nNewVehicleSeatCount
    '
    m_szVehicle = RVehicleID
    m_szVehicleModelID = szNewVehicleModel
    m_szVehicleModelName = szNewVehicleModelName
    m_szOwner = szNewVehicleBusOwnerID
    m_szCompany = szNewVehicleCompanyID
    m_szSplitCompanyID = szNewSplitCompanyID
    
    m_nObjectStatus = ST_EditObj
    
    m_eBusStatus = ST_BusReplace
    
    m_nTotalSeat = nNewVehicleSeatCount
    m_nStartSeatNo = nNewVehicleSeatStart
    
    '更新座位信息
    m_nSaleSeat = UpdateEnviromentSeatCountEx(m_szBusID, m_dtRunDate, oDb)
    m_nOldSaleSeat = m_nSaleSeat
    '更新基本信息
    UpdateToDB oDb
    
    '更新配载表
    szSql = " UPDATE Env_bus_allot_lst SET status = " & TransFieldValueToString(m_eBusStatus) _
        & " WHERE bus_id = " & TransFieldValueToString(m_szBusID) _
        & " AND bus_date = " & TransFieldValueToString(m_dtRunDate)
    
    oDb.Execute szSql
    
    If AutoSeat Then
        szSql = " UPDATE env_bus_sell_station_lst SET can_sale_quantity = -1 " _
            & " WHERE bus_id = " & TransFieldValueToString(m_szBusID) _
            & " AND bus_date = " & TransFieldValueToString(m_dtRunDate) _
            & " AND can_sale_quantity <> 0 "
        
        oDb.Execute szSql
        
        szSql = " UPDATE Env_bus_allot_lst SET can_sale_quantity = -1 " _
            & " WHERE bus_id = " & TransFieldValueToString(m_szBusID) _
            & " AND bus_date = " & TransFieldValueToString(m_dtRunDate) _
            & " AND can_sale_quantity <> 0 "
        
        oDb.Execute szSql
    End If
    
    oDb.CommitTrans
    'On Error GoTo 0
    
    #If SUPPORT_MSG = 1 Then
        Dim aszValues(1 To 4) As String
        aszValues(1) = m_szBusID
        aszValues(2) = ToDBDate(m_dtRunDate)
        aszValues(3) = szNewLicenseTagNo
        aszValues(4) = szNewTransportCompanyShortName
        m_oSndMsg.SendMsg Date, cszReplaceVehicle, aszValues
    #End If
    
    Dim szlog As String
    
    szlog = ToDBDateTime(m_dtOldStartupTime)
    szlog = "车次[" & m_szBusID & "]发车时间[" & szlog & "]终点站[" & m_szOldEndStationName & "]:"
    szlog = szlog & "车辆[" & m_szVehicle & "]被" & RVehicleID & "顶班" & "票价按照车次" & szmyBusID & "生成"
    szlog = GetUnicodeBySize(szlog, 255)
    WriteOperateLog m_oActiveUser, RIGHT_REBusReplaceBus, szlog
    
    Set rsTemp1 = Nothing
    Set oDb = Nothing
    Set rsTemp = Nothing
    Exit Sub
here:
    oDb.RollbackTrans
    Set rsTemp1 = Nothing
    Set oDb = Nothing
    Set rsTemp = Nothing
    err.Raise err.Number
End Sub

'S18
'注意 当bsitpflg=flase 功能为并班车次
Public Sub MergeBus(MergedBusID As String, Annotation As String, AllRefundment As Boolean)
    Dim szSql As String
    Dim szTemp As String
    Dim szSqlfield As String
    Dim szSqlContext As String
    Dim sztempsql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    
    Dim nSeatCount1 As Integer
    Dim nSeatCount2 As Integer
    Dim szSeatNo1 As String
    Dim szSeatNo2 As String
    Dim lRow As Long
    Dim nSeatType As Integer
    Dim tSeatTypeInfo(1 To 5) As TSeatInfoCount
    Dim i As Integer
    '''''''''''
    
    AssertObjIsValid
    '0检票车次不能并班
    RefreshMemoryInfo
    If Me.CheckStatus <> ST_BusCheckWait Then ShowError ERR_REBusCheckBusNotMerge
    '1流水车次不能并班
    If m_nBusType = TP_ScrollBus Then ShowError ERR_REBusScrollBusNotMerge
    '2先判断该车次的状态是否为正常
    If m_eBusStatus <> ST_BusNormal And m_eBusStatus <> ST_BusSlitpStop Then ShowError ERR_REBusNotNormalNoMerge
    '3然后再判断该车次是否已经并入了一个车次
     AssertHaveRight m_oActiveUser, RIGHT_REBusMergeBus
     
    If Me.BeMergedBus.szBusID <> "" Then ShowError ERR_REBusMerged
        '4再判断指定的车次是否已经被并入了其它车次
    If Me.BusMergedMe <> "" Then ShowError ERR_REBusMerged
    
    If DateDiff("d", m_dtRunDate, Now) > 0 Then ShowError ERR_REBusMergedDate
     oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    On Error GoTo here
    oDb.BeginTrans
        sztempsql = "SELECT bus_type,status AS bus_status FROM Work_env_bus_info WHERE bus_date='" & ToDBDate(m_dtRunDate) & "' AND bus_id='" & Trim(MergedBusID) & "'"
        Set rsTemp = oDb.Execute(sztempsql)
        If FormatDbValue(rsTemp!bus_type) = TP_ScrollBus Then ShowError ERR_REBusScrollBusNotMerge
        If FormatDbValue(rsTemp!bus_status) <> ST_BusNormal And CInt(FormatDbValue(rsTemp!bus_status)) < CnSlitpStatus Then ShowError ERR_REBusNotNormalNoMerge
        '5在并班车次表中新增该并班记录
        '当 bslitpflg=true 时在并班车次表中新增该拆分车次记录
        '并班表同时是车次拆分表
        sztempsql = "INSERT Mergebus_split_lst ("
        szSqlfield = "bus_date,"
        szSqlfield = szSqlfield & "bus_id,"
        szSqlfield = szSqlfield & "merge_bus_id,"
        szSqlfield = szSqlfield & "status,"
        szSqlfield = szSqlfield & "annotation)"
        szSqlContext = " VALUES ('" & ToDBDate(m_dtRunDate) & "',"
        '并班
        '目标车次---m_szBusID,被并班 ----m_szBusID
        szSqlContext = szSqlContext & "'" & m_szBusID & "',"
        szSqlContext = szSqlContext & "'" & MergedBusID & "',"
        szSqlContext = szSqlContext & "'0',"
        szSqlContext = szSqlContext & "'" & Annotation & "')"
        sztempsql = sztempsql & szSqlfield & szSqlContext
        oDb.Execute sztempsql
        If AllRefundment Then
            szTemp = cbAllRefundment
        Else
            szTemp = cbNotAllRefundment
        End If
        Dim szBusSeatStatus() As String
        Dim lRowTemp As Long
        
        szSql = "SELECT * FROM Env_bus_seat_lst WHERE bus_date='" & ToDBDate(m_dtRunDate) & "' AND " _
                      & "bus_id='" & MergedBusID & "' AND " _
                      & "status<>0"
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount <> 0 Then
            nSeatCount1 = rsTemp.RecordCount
            ReDim szBusSeatStatus(1 To rsTemp.RecordCount, 1 To 3)
            
            For i = 1 To rsTemp.RecordCount
                szBusSeatStatus(i, 1) = FormatDbValue(rsTemp!seat_no)
                szBusSeatStatus(i, 2) = FormatDbValue(rsTemp!ticket_no)
                szBusSeatStatus(i, 3) = FormatDbValue(rsTemp!seat_type_id)
                
                szSql = "Update Env_bus_seat_lst set status = " & ST_SeatMerge & "  ,ticket_no='" & szBusSeatStatus(i, 2) & "' WHERE seat_no = " _
                    & " ( SELECT TOP 1  seat_no " _
                    & "  FROM Env_bus_seat_lst WHERE  " _
                    & " bus_date='" & ToDBDate(m_dtRunDate) & "' AND " _
                    & " bus_id='" & m_szBusID & "' AND " _
                    & " status='" & ST_SeatCanSell & "' ) AND bus_date='" & ToDBDate(m_dtRunDate) & "' AND " _
                    & " bus_id='" & m_szBusID & "' AND " _
                    & " status='" & ST_SeatCanSell & "'"
                oDb.Execute szSql, lRow
                lRowTemp = lRowTemp + lRow
            Next
        End If
        
        
        If lRowTemp <> nSeatCount1 Then ShowError 18871 '并班车次座位数已不够
        '6 修改被并车次的状态和是否全额退票
        szSql = " SELECT COUNT(seat_type_id) as SeatTypeCount ,Seat_type_id FROM Env_bus_seat_lst WHERE  bus_date='" & ToDBDate(m_dtRunDate) & "' AND " _
            & " bus_id='" & m_szBusID & "' AND " _
            & " status='" & ST_SeatMerge & "'GROUP BY Seat_type_id"
        Set rsTemp = oDb.Execute(szSql)
        lRow = 0
        For i = 1 To rsTemp.RecordCount
            nSeatType = Val(FormatDbValue(rsTemp!seat_type_id))
            tSeatTypeInfo(nSeatType).seatCount = 0 - FormatDbValue(rsTemp!SeatTypeCount)
            lRow = lRow + FormatDbValue(rsTemp!SeatTypeCount)
            rsTemp.MoveNext
        Next
        If lRow <> nSeatCount1 Then ShowError 18871
        szSql = "UPDATE Env_bus_info SET status=" & ST_BusMergeStopped & ",is_all_refundment=" & szTemp & " WHERE bus_id='" & MergedBusID & "' AND " _
        & " bus_date='" & ToDBDate(m_dtRunDate) & "'"
        oDb.Execute szSql
        
        szSql = " UPDATE Env_bus_allot_lst SET status =" & ST_BusMergeStopped & " WHERE bus_id='" & MergedBusID & "' AND " _
        & " bus_date='" & ToDBDate(m_dtRunDate) & "'"
        oDb.Execute szSql
        
        '修改并班车次环境车次表中可售座位总数
        '      szsql = "UPDATE Env_bus_info SET sale_seat_quantity=sale_seat_quantity -" & nSeatCount1 & " WHERE " _
        '                        & "bus_id='" & m_szBusID & "' AND " _
        '                        & "bus_date='" & ToDBDate(m_dtRunDate) & " '"
        '     odb.Execute szsql
        
        
        
        UpdateEnviromentSeatCountEx m_szBusID, m_dtRunDate, oDb
        
        
     oDb.CommitTrans
    
    
    '''''''''''''''''''
  
    
    
    ''''''''''''''''''
    
    #If SUPPORT_MSG = 1 Then
    Dim oParam As New SystemParam
    Dim aszValues(1 To 2) As String
    aszValues(1) = MergedBusID
    aszValues(2) = ToDBDate(m_dtRunDate)
    m_oSndMsg.SendMsg oParam.NowDateTime, cszMergeBus, aszValues
    #End If
       WriteOperateLog m_oActiveUser, RIGHT_REBusMergeBus, "车次" & m_szBusID & "和车次" & MergedBusID & "并班"
   Exit Sub
here:
    oDb.RollbackTrans
    err.Raise err.Number
End Sub


'S19车次环境内复班
Public Sub ResumeBus(StartDate As Date, EndDate As Date, Optional AllRefundment As Boolean = False)
    Dim szSql As String
    Dim szTemp As String
    Dim oDb As New RTConnection
    Dim lRow As Long
    Dim szlog As String
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_REBusResumeStopBus
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    
    '判断车次停班的开始日期和结束日期是否在运行环境内
    If Me.busStatus <> ST_BusStopped Then ShowError ERR_REBusNotStopNotResume
    'If InPreDate(StartDate, EndDate) = False Then ShowError ERR_REBusResumeBusDateNoOk
    If AllRefundment Then
       szTemp = cbAllRefundment
    Else
        szTemp = cbNotAllRefundment
    End If
    ' 2001-10-24 22
    On Error GoTo ErrHandle
    oDb.BeginTrans
    If StartDate = CDate(cszEmptyDateStr) And StartDate = EndDate Then
        '复班环境所有车次
        szSql = "UPDATE Env_bus_info SET status=" & Str(ST_BusNormal) & ",is_all_refundment=" & Str(szTemp) & " WHERE bus_id='" & m_szBusID & "' AND " _
            & " bus_date>='" & ToDBDate(Now) & "'"
       
        oDb.Execute szSql
        szSql = " UPDATE Env_bus_allot_lst SET status=" & ST_BusNormal & _
                " WHERE bus_id=" & TransFieldValueToString(m_szBusID) & " AND bus_date>='" & ToDBDate(Now) & "'"
        oDb.Execute szSql
        szlog = "发车日期大于等于" & ToDBDate(Now) & "的环境车次[" & m_szBusID & "]全被复班"
    Else
      '5 修改车次的状态和是否全额退票
        szSql = "UPDATE Env_bus_info SET status=" & Str(ST_BusNormal) & ",is_all_refundment=" & Str(szTemp) & " WHERE bus_id='" & m_szBusID & "' AND " _
            & " bus_date>='" & ToDBDate(StartDate) & "' AND bus_date<='" & ToDBDate(EndDate) & "'"
        oDb.Execute szSql
        szSql = " UPDATE Env_bus_allot_lst SET status=" & ST_BusNormal & _
                " WHERE bus_id=" & TransFieldValueToString(m_szBusID) & " AND bus_date>='" & ToDBDate(StartDate) & "' AND bus_date<='" & ToDBDate(EndDate) & "'"
        oDb.Execute szSql
      szlog = "车次" & m_szBusID & "发车时间[" & ToDBDateTime(m_dtOldStartupTime) & "]终点站[" & m_szOldEndStationName & "]:在" & ToDBDate(StartDate) & "--" & ToDBDate(EndDate) & "复班"
    End If
    oDb.CommitTrans
    
    
    #If SUPPORT_MSG = 1 Then
    Dim oParam As New SystemParam
    Dim aszValues(1 To 2) As String
    aszValues(1) = m_szBusID
    aszValues(2) = ToDBDate(m_dtRunDate)
    m_oSndMsg.SendMsg oParam.NowDateTime, cszResumeBus, aszValues
    #End If
    
'    If lRow <> 0 Then
      
    szlog = GetUnicodeBySize(szlog, 255)
    WriteOperateLog m_oActiveUser, RIGHT_REBusResumeStopBus, szlog
'    End If
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    oDb.RollbackTrans
    RaiseError err.Number, , err.Description
End Sub

'S20车次停班
Public Sub StopBus(StartDate As Date, EndDate As Date, Optional AllRefundment As Boolean = True)
    Dim szSql As String
    Dim szTemp As String
    Dim strlog As String
    Dim oDb As New RTConnection
    Dim szlog As String
    Dim lRow As Long
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_REBusStopBus
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    '判断车次停班的开始日期和结束日期是否在运行环境内
'    If Me.busStatus = ST_BusChecking Then ShowError ERR_BusCheckNoStop
'    If Me.busStatus = ST_BusExtraChecking Then ShowError ERR_BusCheckNoStop
'    If Me.busStatus = ST_BusStopCheck Then ShowError ERR_BusCheckNoStop
'    If Me.busStatus = ST_BusStopped Then ShowError ERR_BusStopNoStop
'    If Me.busStatus = ST_BusMergeStopped Then ShowError ERR_MergeNoStop
'    If Me.busStatus = ST_BusSlitpStop Then ShowError ERR_REYesSlitpREBus
   
   Select Case Me.busStatus
          Case ST_BusStopped
             ' ShowError ERR_BusStopNoStop
             Exit Sub
          Case ST_BusMergeStopped
              ShowError ERR_MergeNoStop
          Case ST_BusSlitpStop
              ShowError ERR_REYesSlitpREBus
          Case ST_BusChecking
             strlog = "车次正在检票"
           Case ST_BusExtraChecking
             strlog = "车次正在补检"
          Case ST_BusStopCheck
             strlog = "车次已经停检"
          Case Else
            strlog = "车次正常"
    End Select
    
    '流水车次能停班
    'If Me.BusType = TP_ScrollBus Then ShowError ERR_REBusScrollBusNotStop
'    If InPreDate(StartDate, EndDate) = False Then ShowError ERR_REBusStopBusDateNoOk
    If AllRefundment Then
        szTemp = cbAllRefundment
    Else
        szTemp = cbNotAllRefundment
    End If
    On Error GoTo ErrHandle
    oDb.BeginTrans
     If StartDate = CDate(cszEmptyDateStr) And StartDate = EndDate Then
        szSql = "UPDATE Env_bus_info SET status=" & Str(ST_BusStopped) & ",is_all_refundment=" & Str(szTemp) & " WHERE bus_id='" & m_szBusID & "' AND " _
            & " bus_date>='" & ToDBDate(Now) & "'"
        oDb.Execute szSql
        szSql = " UPDATE Env_bus_allot_lst SET status=" & ST_BusStopped & _
                " WHERE bus_id=" & TransFieldValueToString(m_szBusID) & " AND bus_date>='" & ToDBDate(Now) & "'"
        oDb.Execute szSql
        szlog = "发车日期大于等于" & ToDBDate(Now) & "的环境车次[" & m_szBusID & "]全被停班"
     Else
        '复班环境所有车次
        '5修改被车次的状态和是否全额退票
        szSql = "UPDATE Env_bus_info SET status=" & Str(ST_BusStopped) & ",is_all_refundment=" & Str(szTemp) & " WHERE bus_id='" & m_szBusID & "' AND " _
                & " bus_date>='" & ToDBDate(StartDate) & "' AND bus_date<='" & ToDBDate(EndDate) & "'"
        oDb.Execute szSql
        szSql = " UPDATE Env_bus_allot_lst SET status=" & ST_BusStopped & _
                " WHERE bus_id=" & TransFieldValueToString(m_szBusID) & " AND bus_date>='" & ToDBDate(StartDate) & "' AND bus_date<='" & ToDBDate(EndDate) & "'"
        oDb.Execute szSql
        szlog = "车次[" & m_szBusID & "]发车时间为[" & ToDBDateTime(m_dtOldStartupTime) & "]终点站[" & m_szOldEndStationName & "]车次状态" & strlog & ":在" & ToDBDate(StartDate) & "--" & ToDBDate(EndDate) & "停班"
    End If
    '复制该车次的配载信息
    oDb.CommitTrans
    #If SUPPORT_MSG = 1 Then
        Dim oParam As New SystemParam
        Dim aszValues(1 To 2) As String
        aszValues(1) = m_szBusID
        aszValues(2) = ToDBDate(m_dtRunDate)
        m_oSndMsg.SendMsg oParam.NowDateTime, cszStopBus, aszValues
    #End If
    ' strlog = "车次状态：" & strlog & ",在"
'    If lRow <> 0 Then
     szlog = GetUnicodeBySize(szlog, 255)
     WriteOperateLog m_oActiveUser, RIGHT_REBusStopBus, szlog
'    End If
    
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    oDb.RollbackTrans
    RaiseError err.Number, , err.Description
End Sub

'S21车次报班
Public Sub Being(BeingType As EREBusBeingStatus)
    Dim oDb As New RTConnection
    Dim sztempsql As String
    Dim szSqlfield As String
    Dim szsqlWhere As String
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    sztempsql = "UPDATE Env_bus_info SET "
    szSqlfield = szSqlfield & "register_status=" & Str(BeingType) & ""
    szsqlWhere = " WHERE bus_id='" & m_szBusID & "' AND  bus_date='" & ToDBDate(m_dtRunDate) & "'"
    sztempsql = sztempsql & szSqlfield & szsqlWhere
    oDb.Execute (sztempsql)
    WriteOperateLog m_oActiveUser, RIGHT_REBusManagement, "报班" & m_szBusID
End Sub

'S22取消预留座位
Public Sub UnReserveSeat(SeatNo As String)
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim szSeatNoTemp As String
    Dim oDb As New RTConnection
    
    Dim szSeatNoTemplog As String
AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_ReBusSeatManagement
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    
    '分开座位字符串
    szSeatNoTemp = LeftAndRight(SeatNo, True, "[")
    szSeatNoTemplog = LeftAndRight(SeatNo, False, "[") '预留座位用
    szSeatNoTemplog = "[" & szSeatNoTemplog '写日志用
    
'    If Left(Trim(szSeatNoTemp)) <> "'" Then
'      szSeatNoTemp = "'" & szSeatNoTemp & "'"
'    End If

    '1查询该座位是否可以取消预留
    szSql = "SELECT seat_no FROM Env_bus_seat_lst WHERE bus_id='" & m_szBusID & "'" _
           & " AND  bus_date='" & ToDBDate(m_dtRunDate) & "'" _
           & "AND seat_no IN (" & szSeatNoTemp & " )" _
           & " AND (Status <> " & ST_SeatReserved & " AND Status <> " & ST_SeatProjectBooked & " )"
    Set rsTemp = oDb.Execute(szSql)
    
    If rsTemp.RecordCount > 0 Then
     szSeatNoTemplog = ""
     
      For i = 0 To rsTemp.RecordCount - 1
       szSeatNoTemplog = szSeatNoTemplog & "[" & FormatDbValue(rsTemp!seat_no) & "]"
       rsTemp.MoveNext
      Next
    
    
    ShowError ERR_REBusSeatNoReserved, "座位:" & szSeatNoTemplog
    
    End If
    
    On Error GoTo here
    oDb.BeginTrans
    szSql = "UPDATE Env_bus_seat_lst SET status=" & Str(ST_SeatCanSell) & " ," _
            & " remark='' " _
            & "    WHERE bus_id='" & m_szBusID & "' AND  bus_date='" & ToDBDate(m_dtRunDate) & "' " _
            & " AND seat_no IN (" & szSeatNoTemp & ")" _
            & " AND (Status = " & ST_SeatReserved & " or Status = " & ST_SeatProjectBooked & " )"
    oDb.Execute szSql
    
    UpdateEnviromentSeatCountEx m_szBusID, m_dtRunDate, oDb
    oDb.CommitTrans
    Dim szlog As String
    szlog = "车次[" & m_szBusID & "]发车时间为[" & ToDBDateTime(m_dtOldStartupTime) & "]终点站[" & m_szOldEndStationName & "]:取消预留座位" & szSeatNoTemplog
    WriteOperateLog m_oActiveUser, RIGHT_ReBusSeatManagement, szlog
    Exit Sub
here:

    oDb.RollbackTrans
    err.Raise err.Number
    
End Sub
'S23预留座位
'SeatNo以格式： ''[] 传入
Public Sub ReserveSeat(SeatNo As String, Optional szRemark As String)
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    
    Dim szSeatNoTemp As String
    Dim szSeatNoTemplog As String
    
    Dim i As Integer
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_ReBusSeatManagement
    
    '分开座位字符串
    If SeatNo <> "" Then
        szSeatNoTemp = LeftAndRight(SeatNo, True, "[")
        szSeatNoTemplog = LeftAndRight(SeatNo, False, "[") '预留座位用
        szSeatNoTemplog = "[" & szSeatNoTemplog '写日志用
    End If
    '查询该座位是否可以预留
    szSql = "SELECT seat_no FROM Env_bus_seat_lst WHERE bus_id='" & m_szBusID & "'  " _
            & " AND  bus_date='" & ToDBDate(m_dtRunDate) & "' AND seat_no IN ( " & szSeatNoTemp & ")" _
            & " AND  Status<>" & ST_SeatCanSell & ""
    Set rsTemp = oDb.Execute(szSql)
    
    If rsTemp.RecordCount > 0 Then
     szSeatNoTemplog = ""
      For i = 1 To rsTemp.RecordCount
         szSeatNoTemplog = szSeatNoTemplog & "[" & FormatDbValue(rsTemp!seat_no) & "]"
         rsTemp.MoveNext
      Next
      ShowError ERR_REBusSeatNotReserve, "座位:" & szSeatNoTemplog '& "[" & "不是可售座位不可预留！]"
    End If
    
    
    On Error GoTo here
    
    oDb.BeginTrans
    
    '2预留该座位
    szSql = "UPDATE Env_bus_seat_lst SET status=" & Str(ST_SeatReserved) & "   ," _
          & " remark='" & szRemark & "' " _
          & "  WHERE bus_id='" & m_szBusID & "' " _
          & " AND  bus_date='" & ToDBDate(m_dtRunDate) & "' " _
          & " AND seat_no IN (" & szSeatNoTemp & ")" _
          & " AND  Status=" & ST_SeatCanSell & ""
    oDb.Execute szSql
    
    UpdateEnviromentSeatCountEx m_szBusID, m_dtRunDate, oDb
    
    oDb.CommitTrans
    
    Dim szlog As String
      
    szlog = "车次[" & m_szBusID & "]发车时间为[" & ToDBDateTime(m_dtOldStartupTime) & "]终点站[" & m_szOldEndStationName & "]:预留车次座位" & szSeatNoTemplog
    WriteOperateLog m_oActiveUser, RIGHT_ReBusSeatManagement, szlog
    
    Exit Sub

here:
     oDb.RollbackTrans
     WriteOperateLog m_oActiveUser, RIGHT_ReBusSeatManagement, szlog & "出错 " & err.Number & "]" & err.Description
     err.Raise err.Number
     
End Sub

'S24删除座位
Public Sub DeleteSeat(seatCount As Integer, StartSeatNo As String)
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim TSeatInfoCount(1 To 5) As TSeatInfoCount
    Dim nSerialSeatType As Integer
    Dim oDb As New RTConnection
    Dim szSeatNoTemp As String
    Dim szSeatNoTemplog As String
    Dim lRow As Long
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_ReBusSeatManagement
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    
    
     '分开座位字符串
    If StartSeatNo <> "" Then
        szSeatNoTemp = LeftAndRight(StartSeatNo, True, "[")
        szSeatNoTemplog = LeftAndRight(StartSeatNo, False, "[") '预留座位用
        szSeatNoTemplog = "[" & szSeatNoTemplog '写日志用
   Else
     Exit Sub
   End If
   
   '1查询该座位是否可以删除
   szSql = "SELECT Seat_no FROM Env_bus_seat_lst WHERE bus_id='" & m_szBusID & "' " _
             & " AND  bus_date='" & ToDBDate(m_dtRunDate) & "' " _
             & " AND  seat_no IN (  " & szSeatNoTemp & ")" _
             & " AND Status <> " & ST_SeatCanSell & ""
   Set rsTemp = oDb.Execute(szSql)
   
   If rsTemp.RecordCount > 0 Then
     szSeatNoTemplog = ""
      For i = 0 To rsTemp.RecordCount - 1
         szSeatNoTemplog = szSeatNoTemplog & "[" & FormatDbValue(rsTemp!seat_no) & "]"
         rsTemp.MoveNext
      Next
      ShowError ERR_REBusSeatNotDelete, "座位:" & szSeatNoTemplog
   End If
   
   
   On Error GoTo here
   oDb.BeginTrans
   szSql = "DELETE Env_bus_seat_lst WHERE bus_id='" & m_szBusID & "' " _
                & " AND  bus_date='" & ToDBDate(m_dtRunDate) & "'  " _
                & " AND seat_no IN ( " & szSeatNoTemp & " )" _
                & " AND ticket_no=''"
   oDb.Execute szSql, lRow
        
   szSql = "UPDATE Env_bus_info SET total_seat=total_seat-" & lRow & " WHERE bus_id='" & m_szBusID & "' AND  bus_date='" & ToDBDate(m_dtRunDate) & "'"
   oDb.Execute szSql
   
    UpdateEnviromentSeatCountEx m_szBusID, m_dtRunDate, oDb
    
    oDb.CommitTrans
    
    m_nTotalSeat = m_nTotalSeat - lRow
    m_nOldTotalSeat = m_nTotalSeat
    
    #If SUPPORT_MSG = 1 Then
    Dim oParam As New SystemParam
    Dim aszValues(1 To 3) As String
    aszValues(1) = m_szBusID
    aszValues(2) = ToDBDate(m_dtRunDate)
    aszValues(3) = m_nTotalSeat
    m_oSndMsg.SendMsg oParam.NowDateTime, cszChangeBusSeat, aszValues
    #End If
    Dim szlog As String
    szlog = "删除车次[" & m_szBusID & "]发车时间为[" & ToDBDateTime(m_dtOldStartupTime) & "]终点站[" & m_szOldEndStationName & "]:"
    WriteOperateLog m_oActiveUser, RIGHT_ReBusSeatManagement, szlog & "座位:座位号" & szSeatNoTemplog
Exit Sub
here:
 oDb.RollbackTrans
 ShowError err.Number
 Exit Sub
here1:
oDb.RollbackTrans
ShowError ERR_REBusSeatNotDelete
End Sub



'得到车次站点已售票信息--
'szBusId---车次代码 dtRunBusDate－－－车次发车日期
Public Function GetBusStatinSellInfo(szBusID As String, dtRunBusdate As Date, SellStationID As String) As String()

   Dim oDb As New RTConnection
   Dim szSql As String
   Dim rsTemp As New Recordset
   Dim nCountInfo As Integer
   Dim SaleStaionInfo() As String
   Dim i As Integer

     On Error GoTo here


   AssertHaveRight m_oActiveUser, RIGHT_REBusStopBus
   oDb.ConnectionString = GetConnectionStr(cszRunEnv)

'    ST_TicketNormal = 1 '车票正常售出
'    ST_TicketSellChange = 2 ' 车票改签售出
'    ST_TicketCanceled = 4 '车票已废
'    ST_TicketChanged = 8 '被改签
'    ST_TicketReturned = 16 '车票已退
'    ST_TicketChecked = 32 '车票已检
'
'     szSql = " SELECT  DISTINCT ew.station_id ,staInfo.station_name,ISNULL(a.SellCount,0)as SellCount ," _
'        & " ISNULL(a.ticket_type,'100') as ticket_type "
'
'    If DateDiff("d", dtRunBusdate, Now) <= 0 Then
'          szSql = szSql & "  FROM  Work_env_bus_station_lst ew, station_info stainfo, "
'    Else
'          szSql = szSql & "  FROM  Env_bus_station_lst ew, station_info stainfo, "
'    End If

    szSql = " SELECT COUNT(*) as sellCount ,tbs.sell_station_id,tbi.station_id , max(tbs.sell_station_name) as sell_station_name,max(tbi.station_name) as station_name " _
          & " FROM Ticket_sell_lst sa ,sell_station_info tbs, station_info tbi " _
          & " WHERE sa.sell_station_id=tbs.sell_station_id AND " _
          & " tbi.station_id =sa.des_station_id  AND " _
          & " bus_date='" & ToDBDate(dtRunBusdate) & " ' AND  " _
          & " bus_id='" & szBusID & "' AND " _
          & " (sa.status<>17 and sa.status<>5  and sa.status<>9 ) " _
          & " GROUP by sa.ticket_type, tbi.station_id ,tbs.sell_station_id "
          
          
          
         

   Set rsTemp = oDb.Execute(szSql)
   nCountInfo = rsTemp.RecordCount
   If nCountInfo <> 0 Then
       ReDim SaleStaionInfo(0 To nCountInfo - 1, 0 To 3)

       For i = 0 To nCountInfo - 1
           SaleStaionInfo(i, 0) = FormatDbValue(rsTemp!station_id)
           SaleStaionInfo(i, 1) = FormatDbValue(rsTemp!station_name)
           SaleStaionInfo(i, 2) = FormatDbValue(rsTemp!SellCount)
           SaleStaionInfo(i, 3) = FormatDbValue(rsTemp!sell_station_id)
           SaleStaionInfo(i, 3) = FormatDbValue(rsTemp!sell_station_name)
          ' SaleStaionInfo(i, 4) = rsTemp!Seat_type
          ' SaleStaionInfo(i, 5) = rsTemp!seat_type_name
          rsTemp.MoveNext
      Next

   End If
   GetBusStatinSellInfo = SaleStaionInfo
   If Not rsTemp Is Nothing Then Set rsTemp = Nothing
   Set oDb = Nothing
   Exit Function
here:
    err.Raise err.Number
    Set oDb = Nothing
    If Not rsTemp Is Nothing Then Set rsTemp = Nothing
End Function

'Public Sub EditAllotInfo(SellStationID As String, CheckGateID As String, StartUpTime As Date, CanSaleQuantity As Integer)
'    Dim szSql As String
'    Dim rsTemp As Recordset
'    Dim oDb As New RTConnection
'    Dim szTempSql As String
'    Dim szSqlField As String
'    Dim szSqlContext As String
'     On Error GoTo ErrHandle
'    AssertObjIsValid '验证用户对象是否合法
'    AssertHaveRight m_oActiveUser, RIGHT_ReBusSeatManagement '操作员权限不够
'
'    szSql = "SELECT * " _
'    & "FROM sell_station_id " _
'    & "WHERE sell_station_id ='" & SellStationID & "'"
'    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
'    Set rsTemp = oDb.Execute(szSql)
'
'    If rsTemp.RecordCount = 0 Then
'        ShowError ERR_NoSellStation '错误信息
'    End If
'
'    szSql = "SELECT * " _
'    & "FROM Env_bus_allot_lst tbe " _
'    & "WHERE tbe.sell_station_id='" & SellStationID & "' " _
'    & "AND tbe.check_gate_id='" & CheckGateID & "' " _
'    & "AND tbe.bus_start_time='" & StartUpTime & "' " _
'    & "AND tbe.can_sale_quantity='" & CanSaleQuantity & "'  "
'
'    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
'    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount = 0 Then '如果对应的配载信息不存在则出错
'        ShowError ERR_Allotexistent '配载信息不存在
'    End If
'    '修改该配载信息
'    szTempSql = "UPDATE Env_bus_allot_lst SET " _
'        & "bus_date ='" & ToDBDate(m_dtRunDate) & ",'" _
'        & "bus_id ='" & m_szBusID & ",'" _
'        & "sell_station_id ='" & SellStationID & ",'" _
'        & "check_gate_id ='" & CheckGateID & ",'" _
'        & "bus_start_time ='" & ToDBDateTime(StartUpTime) & ",'" _
'        & "status = " & EREBusStatus.ST_BusNormal & "," _
'        & "can_sale_quantity = " & CanSaleQuantity & "," _
'        & "sell_quantity = " & m_szSellQuantity
'
'
'    oDb.Execute szTempSql
'
'        Set oDb = Nothing
'
'ErrHandle:
'   AssertUpdateObjectError ERR_Bus, oDb
'
'
'End Sub
''S25
Public Sub AddSeat(seatCount As Integer, StartSeatNo As String, Optional szSeatTypeID As String = "")
    Dim szSql As String
    Dim bAddSeatErr As Boolean
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim sztempsql As String
    Dim szSqlfield As String
    Dim tSeatTypeInfo(1 To 5) As TSeatInfoCount
    Dim szSqlContext As String
    Dim nSerialType As Integer
    Dim nRow As Long
    Dim nRowTemp As Long
    Dim oDb As New RTConnection
    
    Dim szSeatNo As String
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_ReBusSeatManagement
    
    '1查询该座位是否可以删除
    If szSeatTypeID = "" Then
       szSeatTypeID = cszSeatTypeIsNormal
    End If
    
    nSerialType = Val(szSeatTypeID)
    
    '2001－12-31
    For i = 0 To seatCount - 1
     szSeatNo = szSeatNo & "'" & Format(StartSeatNo + i, "00") & "',"
    Next
    szSeatNo = Left(szSeatNo, Len(szSeatNo) - 1)
    szSql = "SELECT seat_no FROM Env_bus_seat_lst WHERE bus_id='" & m_szBusID & "' AND  bus_date='" & ToDBDate(m_dtRunDate) & "' AND seat_no IN ( " & szSeatNo & " )"
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
    
        For i = 0 To rsTemp.RecordCount - 1
           
           szSeatNo = szSeatNo & "[" & FormatDbValue(rsTemp!seat_no) & "]"
           rsTemp.MoveNext
        Next
        
        ShowError ERR_REBusSeatExistNotAddNew, "座位：" & szSeatNo
        
    End If
    
    On Error GoTo here
    oDb.BeginTrans
    For i = 0 To seatCount - 1
       'szsql = "SELECT * FROM Env_bus_seat_lst WHERE bus_id='" & m_szBusID & "' AND  bus_date='" & ToDBDate(m_dtRunDate) & "' AND seat_no='" & Format(StartSeatNo + i, "00") & "'"
      '  Set rsTemp = odb.Execute(szsql)
        '-----------------------
        '车次座位已存在，不能新增
        If rsTemp.RecordCount = 0 Then
            sztempsql = "INSERT Env_bus_seat_lst ("
            szSqlfield = "bus_date,"
            szSqlfield = szSqlfield & "bus_id,"
            szSqlfield = szSqlfield & "seat_no,"
            szSqlfield = szSqlfield & "seat_type_id,"
            szSqlfield = szSqlfield & "status,"
            szSqlfield = szSqlfield & "ticket_no)"
            szSqlContext = " VALUES ('" & Format(m_dtRunDate, "YYYY-MM-DD 0:0:0") & "',"
            szSqlContext = szSqlContext & "'" & m_szBusID & "',"
            szSqlContext = szSqlContext & "'" & BuildSeatNo(StartSeatNo + i) & "',"
            szSqlContext = szSqlContext & "'" & szSeatTypeID & "',"
            szSqlContext = szSqlContext & Str(ST_SeatCanSell) & ",'')"
            sztempsql = sztempsql & szSqlfield & szSqlContext
            oDb.Execute sztempsql, nRow
            
            If nRow <> 0 Then
               nRowTemp = nRowTemp + 1
            End If
            
               End If
    Next
    
    sztempsql = "UPDATE Env_bus_info SET total_seat=total_seat+ " & nRowTemp & " WHERE bus_id='" & m_szBusID & "' AND  bus_date='" & ToDBDate(m_dtRunDate) & "'"
    oDb.Execute sztempsql
    
    UpdateEnviromentSeatCountEx m_szBusID, m_dtRunDate, oDb
    oDb.CommitTrans
    #If SUPPORT_MSG = 1 Then
    Dim oParam As New SystemParam
    Dim aszValues(1 To 4) As String
        aszValues(1) = m_szBusID
        aszValues(2) = ToDBDate(m_dtRunDate)
        aszValues(3) = m_nTotalSeat
        aszValues(4) = szSeatTypeID
        m_oSndMsg.SendMsg oParam.NowDateTime, cszChangeBusSeat, aszValues
    #End If
    
    Dim szlog As String
    szlog = "车次[" & m_szBusID & "]发车时间为[" & ToDBDateTime(m_dtOldStartupTime) & "]终点站[" & m_szOldEndStationName & "]:新增座位:座位号从"
    WriteOperateLog m_oActiveUser, RIGHT_ReBusSeatManagement, szlog & Format(StartSeatNo, "00") & "到" & Format(StartSeatNo + seatCount - 1, "00") & ",座型" & szSeatTypeID
    Exit Sub
here:
    oDb.RollbackTrans
    ShowError err.Number
End Sub

'断言该车次可以被顶班
Private Sub AssertBusCanBeReplaced()
    If m_nBusType = TP_ScrollBus Then ShowError ERR_REBusScrollBusNotReplace
    If busStatus = ST_BusMergeStopped Or busStatus = ST_BusSlitpStop Then ShowError ERR_REBusBeMerged
  '  If busStatus = ST_BusStopped Then ShowError ERR_REBusIsStopped
    If BeMergedBus.szBusID <> "" Then ShowError ERR_REBusMergedOtherBus
    If CheckStatus <> ST_BusCheckWait Then ShowError ERR_REBusHaveCheckManner
End Sub

Private Sub GetCheckTime()
    Dim dtTmpTime As Date
    Dim oDb As New RTConnection
    Dim rsTmp As Recordset
    Dim szSql As String
    'AssertObjIsValid
    If m_nObjectStatus = ST_AddObj Then
        ShowError ERR_AddObj
    Else
        oDb.ConnectionString = GetConnectionStr(cszRunEnv)
        If m_bRefreshBusCheck Then
            If m_nBusType = TP_ScrollBus Then
                szSql = "SELECT a.check_start_time,a.check_end_time FROM Bus_check_lst a," & _
                                    "(SELECT MAX(bus_serial_no) bus_serial_no FROM Bus_check_lst" & _
                                    " WHERE bus_id='" & m_szBusID & "' AND bus_start_time='" & ToDBDateTime(m_dtStartupTime) & "') b" & _
                        " WHERE bus_id='" & m_szBusID & "' AND bus_start_time='" & ToDBDateTime(m_dtStartupTime) & "'" & _
                            " AND a.bus_serial_no=b.bus_serial_no"
            Else
                szSql = "SELECT check_start_time,check_end_time FROM Bus_check_lst WHERE bus_id='" & m_szBusID & "' AND bus_start_time='" & ToDBDateTime(m_dtStartupTime) & "'"
            End If
            Set rsTmp = oDb.Execute(szSql)
            If rsTmp.RecordCount = 0 Then
                m_dtStartCheckTime = CDate(cszEmptyDateStr)
                m_dtStopCheckTime = CDate(cszEmptyDateStr)
            Else
                m_dtStartCheckTime = FormatDbValue(rsTmp!check_start_time)
                m_dtStopCheckTime = FormatDbValue(rsTmp!check_end_time)
            End If
            m_bRefreshBusCheck = False
        End If
    End If
End Sub

'删除开检时间
Public Property Get StartCheckTime() As Date
    If m_bRefreshBusCheck Then
        GetCheckTime
    End If
    StartCheckTime = m_dtStartCheckTime
End Property

'删除 停检时间
Public Property Get StopCheckTime() As Date
    'GetCheckTime
    StopCheckTime = m_dtStopCheckTime
End Property

Public Sub ChangeStartupTime(StartDate As Date, EndDate As Date, StartUpTime As Date)
    Dim szSql As String
    Dim szTemp As String
    Dim dtTemp As Date
    Dim i As Integer
    Dim oDb As New RTConnection
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    '判断车次停班的开始日期和结束日期是否在运行环境内
    If InPreDate(StartDate, EndDate) = False Then ShowError ERR_REBusStopBusDateNoOk
    For i = 0 To DateDiff("d", StartDate, EndDate)
    dtTemp = DateAdd("d", i, StartDate)
    StartUpTime = SelfGetFullDateTime(dtTemp, StartUpTime)
    szSql = "UPDATE Env_bus_info SET bus_start_time='" & ToDBDateTime(StartUpTime) & "' WHERE bus_id='" & m_szBusID & "' AND " _
    & " bus_date='" & ToDBDate(dtTemp) & "'"
    oDb.Execute szSql
    Next
    #If SUPPORT_MSG = 1 Then
    Dim aszValues(1 To 3) As String
    Dim oParam As New SystemParam
    oParam.Init m_oActiveUser
    aszValues(1) = m_szBusID
    aszValues(2) = ToDBDate(m_dtRunDate)
    aszValues(3) = ToDBDateTime(SelfGetFullDateTime(m_dtRunDate, StartUpTime))
    m_oSndMsg.SendMsg oParam.NowDateTime, cszChangeBusTime, aszValues
    #End If
    Dim szlog As String
    If Trim(ToDBDateTime(StartUpTime)) <> Trim(m_dtOldStartupTime) Then
        szlog = "车次" & m_szBusID & "终点站[" & m_szOldEndStationName & "]:发车时间由[" & ToDBDateTime(m_dtOldStartupTime) & "]修改为[" & Left(ToDBDateTime(StartUpTime), 19) & "]"
        WriteOperateLog m_oActiveUser, RIGHT_REBusStopBus, szlog
    End If
End Sub



Public Function GetNotCanSellCount(Optional pszSellStationID As String) As Integer
    Dim szSql As String
    Dim i As Integer
    Dim oDb As New RTConnection
    Dim rsTmp As Recordset
    Dim nTmpCount As Integer
    AssertObjIsValid
    If m_nObjectStatus = ST_AddObj Then
        ShowError ERR_AddObj
    Else
        oDb.ConnectionString = GetConnectionStr(cszRunEnv)
        szSql = "SELECT COUNT(*)  AS seatcount FROM  Env_bus_seat_lst e , ticket_sell_lst s WHERE e.status>=3 AND e.status<=5  " & _
                " AND e.bus_id='" & m_szBusID & "' AND e.bus_date='" & ToDBDate(m_dtRunDate) & "' AND s.ticket_id =  e.ticket_no "
        If pszSellStationID <> "" Then szSql = szSql & " AND s.sell_station_id = '" & pszSellStationID & "'"
        
        Set rsTmp = oDb.Execute(szSql)
        nTmpCount = rsTmp!seatCount
    End If
    GetNotCanSellCount = nTmpCount
End Function



Public Property Get VehicleTag() As String
    VehicleTag = m_szVehicleTag
End Property
'已售的座位数
Public Property Get SaledSeatCount() As Integer
    Dim szSql As String
    Dim i As Integer
    Dim oDb As New RTConnection
    Dim rsTmp As Recordset
    Dim nTmpCount As Integer
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    szSql = "SELECT COUNT(*)  AS seatcount FROM  Env_bus_seat_lst WHERE (status=" & ST_SeatSold & " or status=" & ST_SeatReplace & "  or status=" & ST_SeatSlitp & " )" _
             & " AND bus_id='" & m_szBusID & "' AND bus_date='" & ToDBDate(m_dtRunDate) & "'"
    Set rsTmp = oDb.Execute(szSql)
    nTmpCount = rsTmp!seatCount
    SaledSeatCount = nTmpCount
End Property

'预留的座位数
Public Property Get ReserveSeatCount() As Integer
    Dim szSql As String
    Dim i As Integer
    Dim oDb As New RTConnection
    Dim rsTmp As Recordset
    Dim nTmpCount As Integer
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    szSql = "SELECT COUNT(*)  AS seatcount FROM  Env_bus_seat_lst WHERE status=" & ST_SeatReserved & _
                " AND bus_id='" & m_szBusID & "' AND bus_date='" & ToDBDate(m_dtRunDate) & "'"
    Set rsTmp = oDb.Execute(szSql)
    nTmpCount = rsTmp!seatCount
    ReserveSeatCount = nTmpCount
End Property


'复制班次
Public Sub CloneBus(CloneBusDate As Date, CloneBusID As String, Optional VehicleID As String = "", Optional ByVal CloneBusEndDate As Date = cszEmptyDateStr)
    Dim ExeSQL As String
    Dim szSql As String
    Dim szset As String
    Dim oDb As New RTConnection
    Dim szAddBusFix  As String
    Dim bEndStationCanSale As Boolean
    
    Dim szSqlfrm As String
    Dim szSqlfrm2 As String
    Dim rsTemp As Recordset
    Dim szCheckGate As String
    Dim nInternetStatus As Integer
    Dim nString As Integer
    Dim bISAddbus As Boolean
    Dim bISOrignAddbus As Boolean
    Dim nSeatType As Integer
    Dim tSeatTypeInfo(1 To 5) As TSeatInfoCount
    Dim rstempSeat As New Recordset
    Dim i As Integer
    Dim szlog As String
    Dim oParam As New SystemParam
    
    Dim aszValues() As String
    Dim nDays As Integer
    Dim dtBusDate As Date
    
    Const cnMaxCloneDays = 20 '最大的可以复制的天数
    
    
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_REBusAdd
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    If CloneBusEndDate = cszEmptyDateStr Then
        '如果为默认,则设为复制一天
        CloneBusEndDate = CloneBusDate
    End If
    nDays = DateDiff("d", CloneBusDate, CloneBusEndDate)
    If nDays > cnMaxCloneDays Then
        '如果范围大于复制的天数的限制,则出错
        ShowError ERR_BatchCloneEnvBusTooLong
    End If
    
    
    oParam.Init m_oActiveUser
    bEndStationCanSale = oParam.EndStationCanSale
    '加班车前缀
    szAddBusFix = oParam.AdditionBusPreFix
    
    nString = InStr(1, szAddBusFix, "%")
    If nString <> 0 Then
        szAddBusFix = Left(szAddBusFix, nString - 1)
    End If
    
    '加班车的是否允许联网售票
    nInternetStatus = 0
    nString = Len(szAddBusFix)
    If szAddBusFix <> "" Then
        If UCase(Trim(Left(Trim(CloneBusID), nString))) = UCase(Trim(szAddBusFix)) Then
            bISAddbus = True
            nInternetStatus = 1
        End If
        
        '原来的车是不是加班车
        If UCase(Trim(Left(Trim(m_szBusID), nString))) = UCase(Trim(szAddBusFix)) Then
            bISOrignAddbus = True
        End If
        
    End If
    
    
    '判断该时间段内的车次是否均存在
    ExeSQL = "SELECT bus_id " _
        & " FROM Env_bus_info " _
        & " WHERE bus_id='" & CloneBusID & "'" _
        & " AND bus_date>='" & ToDBDate(CloneBusDate) & "'" _
        & " AND bus_date <'" & ToDBDate(DateAdd("d", 1, CloneBusEndDate)) & "'"
    If oDb.Execute(ExeSQL).RecordCount <> 0 Then ShowError ERR_REBusAdd
    
        
    '按日期决定查询不同的表
    If DateDiff("d", m_dtRunDate, Now) <= 0 Then
        szSqlfrm = "work_"
        szSqlfrm2 = "work_"
    Else
        szSqlfrm = ""
        szSqlfrm2 = ""
    End If
    
    
    On Error GoTo here
    oDb.BeginTrans
        
        For i = 0 To nDays '如果只复制一天,则nday = 0
            dtBusDate = DateAdd("d", i, CloneBusDate)
            
            
            '1在计划中添加车次信息表
            szSql = ""
            szset = ""
            szSql = "INSERT INTO Env_bus_info ("
            '添加车次代码
            szSql = szSql & "bus_id,"
            szset = szset & "'" & Trim(CloneBusID) & "' AS bus_id,"
            '添加计划代码
            szSql = szSql & "bus_date,"
            szset = szset & "'" & ToDBDate(dtBusDate) & "' AS bus_date,"
            '互联吗？
            szSql = szSql & "internet_status,"
            szset = szset & "'" & nInternetStatus & "' AS internet_status,"
            
            szSql = szSql & "bus_start_time, "
            szset = szset & "'" & ToDBDate(dtBusDate) & " '+" & "convert(char(5),bus_start_time ,108) bus_start_time , "
            
            '添加线路
            szSql = szSql & "check_gate_id,vehicle_type_name,vehicle_type_code,total_seat,sale_seat_quantity, " _
                & " total_stand_quantity,sale_stand_seat_quantity,bus_type,register_status, " _
                & " is_all_refundment,delay_status,status,route_id,end_station_name,owner_id,transport_company_id, " _
                & " split_company_id,end_station_id,ScrollBus_Check_Time, " _
                & " seat_remain ,bed_remain,additional_remain,other_remain_1,other_remain_2," _
                & " baggage_number,fact_weight,calculate_weight,over_weight_number,luggage_total_price,"
            szset = szset & "check_gate_id,vehicle_type_name,vehicle_type_code,total_seat, total_seat AS sale_seat_quantity," _
                & " total_stand_quantity,total_stand_quantity AS sale_stand_seat_quantity,bus_type,register_status," _
                & " is_all_refundment,0,1,route_id,end_station_name,owner_id,transport_company_id," _
                & " split_company_id,end_station_id,ScrollBus_Check_Time," _
                & " seat_remain  ,bed_remain ,additional_remain  ,other_remain_1 ,other_remain_2 ," _
                & " baggage_number,fact_weight,calculate_weight,over_weight_number,luggage_total_price,"
            If VehicleID = "" Then
                '添加车辆代码
                szSql = szSql & "vehicle_id"
                szset = szset & "vehicle_id"
            Else
                szSql = szSql & "vehicle_id"
                szset = szset & "vehicle_id AS " & VehicleID
            End If
            
            ExeSQL = szSql & ") SELECT " & szset _
                & " FROM " & szSqlfrm & "env_bus_info" & " WHERE bus_id='" & m_szBusID & "' AND bus_date='" & ToDBDate(m_dtRunDate) & "'"
            oDb.Execute ExeSQL
            
            Set rsTemp = oDb.Execute(" SELECT " & szset & " FROM " & szSqlfrm & "env_bus_info" & " WHERE bus_id='" & m_szBusID & "' AND bus_date='" & ToDBDate(m_dtRunDate) & "'")
            szCheckGate = FormatDbValue(rsTemp!check_gate_id)
            Set rsTemp = Nothing
            
            '2复制车次站点表
            szSql = ""
            szset = ""
            szSql = "INSERT INTO Env_bus_station_lst ("
            
            '添加车次代码
            szSql = szSql & "bus_id,"
            szset = szset & "'" & Trim(CloneBusID) & "' AS bus_id,"
            '添加计划代码
            szSql = szSql & "bus_date,"
            szset = szset & "'" & ToDBDate(dtBusDate) & "' AS bus_date,"
            
            szSql = szSql & "sell_station_id,station_id,station_serial_no,seat_type_id,max_sale_quantity,sale_ticket_quantity,book_count,stop_sale_time,mileage,full_price,half_price,preferential_ticket1,preferential_ticket2,preferential_ticket3"
            szset = szset & "sell_station_id,station_id,station_serial_no,seat_type_id,max_sale_quantity,sale_ticket_quantity,0,stop_sale_time,mileage,full_price,half_price,preferential_ticket1,preferential_ticket2,preferential_ticket3"
            ExeSQL = szSql & ") SELECT " & szset _
                & " FROM " & szSqlfrm2 & "env_bus_station_lst" _
                & " WHERE bus_id='" & m_szBusID & "' " _
                & " AND bus_date='" & ToDBDate(m_dtRunDate) & "'"
            oDb.Execute ExeSQL
            
            '判断源车次站点与计划站点数是不是一样，如果不一样的站不用限制终点站可售
            Dim nPlanBusStationCount As Integer
            Dim nEnvBusStationCount As Integer
            Dim bPlanStationEqualEnvStation As Boolean
            Dim bEnvInsertStation As Boolean
            
            If Not bISOrignAddbus Then
                szSql = " SELECT  COUNT(DISTINCT station_id) num FROM bus_station_info WHERE bus_id = '" & m_szBusID & "'"
                Set rsTemp = oDb.Execute(szSql)
                If rsTemp.RecordCount = 0 Then
                    nPlanBusStationCount = 0
                Else
                    nPlanBusStationCount = FormatDbValue(rsTemp!num)
                End If
                
                szSql = " SELECT  COUNT(DISTINCT station_id) num FROM work_env_bus_station_lst WHERE " _
                        & " bus_id = '" & m_szBusID & "' AND bus_date = '" & m_dtRunDate & "'"
                Set rsTemp = oDb.Execute(szSql)
                If rsTemp.RecordCount = 0 Then
                    nEnvBusStationCount = 0
                Else
                    nEnvBusStationCount = FormatDbValue(rsTemp!num)
                End If
                
                bPlanStationEqualEnvStation = IIf(nPlanBusStationCount = nEnvBusStationCount, True, False)
            Else
                bPlanStationEqualEnvStation = True
            End If
            
            szSql = " SELECT DISTINCT station_id FROM work_env_bus_station_lst WHERE station_serial_no = 1" _
                    & " AND bus_id = '" & m_szBusID & "' AND bus_Date = '" & m_dtRunDate & "'"
            Set rsTemp = oDb.Execute(szSql)
            
            If rsTemp.RecordCount > 1 Then
                bEnvInsertStation = True
            End If
            
            If bEndStationCanSale = True And bISAddbus = True And (Not bEnvInsertStation) Then
               
               Dim bDoNotRefresh As Boolean  '加班车是否要刷新，如果有一个以上站点为不可售，自动按所来的复制

               szSql = " SELECT DISTINCT station_id FROM env_bus_station_lst WHERE bus_id = " & TransFieldValueToString(m_szBusID) & " AND " _
                       & " bus_date = " & TransFieldValueToString(m_dtRunDate) & " AND max_sale_quantity = -1 "
                       
                If oDb.Execute(szSql).RecordCount > 1 Then bDoNotRefresh = True
                
                '源班次是加班车的话，如果站点中有大于一个的站点是不限张数的，就不用限制（最后一个站可售），如果源班次不是加班车次，如果计划站点数与环境
                '站点数不一致，也不用限制，除此之外要限制
                If ((Not bISOrignAddbus) And bPlanStationEqualEnvStation) Or (bISOrignAddbus And (Not bDoNotRefresh)) Then
                
                    szSql = "Update Env_bus_station_lst set max_sale_quantity=0 " _
                        & " WHERE station_serial_no < ( " _
                        & " SELECT Max(station_serial_no) " _
                        & " FROM " & szSqlfrm2 & "env_bus_station_lst" _
                        & " WHERE bus_id = '" & CloneBusID & "' " _
                        & " AND bus_date = '" & ToDBDate(dtBusDate) & "' " _
                        & " AND max_sale_quantity=-1) AND bus_id = " & TransFieldValueToString(CloneBusID) & "  AND bus_date = " & TransFieldValueToString(ToDBDate(dtBusDate))
                    oDb.Execute szSql
                End If
            End If
            
            
            '复制该车次的配载信息
            szSql = " INSERT Env_bus_allot_lst (bus_date , bus_id , sell_station_id , check_gate_id, bus_start_time , status,can_sale_quantity ,sell_quantity) " _
                & " SELECT '" & ToDBDate(dtBusDate) & "','" & Trim(CloneBusID) & "' ,  sell_station_id , check_gate_id, " & "'" & ToDBDate(dtBusDate) & " '+" & "convert(char(5),bus_start_time ,108) bus_start_time  , 1,can_sale_quantity,0 " _
                & " FROM Env_bus_allot_lst " _
                & " WHERE bus_date='" & ToDBDate(m_dtRunDate) & "' AND bus_id='" & m_szBusID & "'"
            oDb.Execute szSql
            
            
            '复制该车次的售票点信息
            szSql = " INSERT env_bus_sell_station_lst (bus_date , bus_id , sell_station_id,can_sale_quantity ,sell_quantity) " _
                & " SELECT '" & ToDBDate(dtBusDate) & "','" & Trim(CloneBusID) & "' ,  sell_station_id ,can_sale_quantity,0 " _
                & " FROM env_bus_sell_station_lst " _
                & " WHERE bus_date='" & ToDBDate(m_dtRunDate) & "' AND bus_id='" & m_szBusID & "'"
            oDb.Execute szSql
            
            
            '3复制车次票价表
            szSql = ""
            szset = ""
            szSql = "INSERT INTO Env_bus_price_lst ("
            '添加车次代码
            szSql = szSql & "bus_id,"
            szset = szset & "'" & Trim(CloneBusID) & "' AS bus_id,"
            '添加计划代码
            szSql = szSql & "bus_date,"
            szset = szset & "'" & ToDBDate(dtBusDate) & "' AS bus_date,"
            '添加线路   0==FL=}===========>
            szSql = szSql & "sell_station_id,station_id,ticket_type,seat_type_id,station_serial_no,ticket_price_total,base_carriage,price_item_1,price_item_2,price_item_3,price_item_4,price_item_5,price_item_6,price_item_7,price_item_8,price_item_9,price_item_10,price_item_11,price_item_12,price_item_13,price_item_14,price_item_15"
            szset = szset & "sell_station_id,station_id,ticket_type,seat_type_id,station_serial_no,ticket_price_total,base_carriage,price_item_1,price_item_2,price_item_3,price_item_4,price_item_5,price_item_6,price_item_7,price_item_8,price_item_9,price_item_10,price_item_11,price_item_12,price_item_13,price_item_14,price_item_15"
            ExeSQL = szSql & ") SELECT " & szset _
            & " FROM " & szSqlfrm & "env_bus_price_lst" _
            & " WHERE bus_id='" & m_szBusID & "' " _
            & " AND bus_date='" & ToDBDate(m_dtRunDate) & "'"
            oDb.Execute ExeSQL
            
            
            '2复制车次座位表    0==FL=}===========>
            szSql = ""
            szset = ""
            szSql = "INSERT INTO Env_bus_seat_lst ("
            '添加车次代码
            szSql = szSql & "bus_id,"
            szset = szset & "'" & Trim(CloneBusID) & "' AS bus_id,"
            '添加计划代码
            szSql = szSql & "bus_date,"
            szset = szset & "'" & ToDBDate(dtBusDate) & "' AS bus_date,"
            szSql = szSql & "seat_no,status,ticket_no,seat_type_id"
            szset = szset & "seat_no," & ST_SeatCanSell & " AS status, ' ' as ticket_no ,seat_type_id"
            ExeSQL = szSql & ") SELECT " & szset _
            & " FROM Env_bus_seat_lst " _
            & " WHERE bus_id='" & m_szBusID & "' " _
            & " AND bus_date='" & ToDBDate(m_dtRunDate) & "'"
            
            oDb.Execute ExeSQL
            
            
            UpdateEnviromentSeatCountEx CloneBusID, dtBusDate, oDb
        Next i
    oDb.CommitTrans
    
    szlog = "新增车次[" & CloneBusID & "],发车日期[" & ToDBDate(dtBusDate) & "]，它由车次[" & m_szBusID & "]发车时间为[" & ToDBDateTime(m_dtOldStartupTime) & "]终点站[" & m_szOldEndStationName & "]:拷贝产生"
    WriteOperateLog m_oActiveUser, RIGHT_REBusAdd, szlog
    
    #If SUPPORT_MSG = 1 Then
        
        ReDim aszValues(1 To 3)
        aszValues(1) = CloneBusID
        aszValues(2) = ToDBDate(dtBusDate)
        aszValues(3) = Trim(szCheckGate)
        m_oSndMsg.SendMsg oParam.NowDateTime, cszAddBus, aszValues
        m_bCheckGateChanged = False
        m_bOffTimeChanged = False
    #End If
    Exit Sub
here:
    oDb.RollbackTrans
    ShowError err.Number
End Sub
'返回拆分车次信息
'********************
Public Function GetSlitpInfo(Optional szSplitBusID As String, Optional dtSlitpDate As Date) As String()  ', Optional szSellStationID As String = "") As String()
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim szTemp() As String
    Dim i As Integer
    Dim oDb As New RTConnection
    
    On Error GoTo here
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    '得到指定被并车次的并入车次
'    szSql = " SELECT bus_id FROM Mergebus_split_lst WHERE merge_bus_id='" & szSplitBusID & "'AND status ='" & 1 & "' "
    szSql = " SELECT tbm.* " _
    & " FROM Mergebus_split_lst tbm " _
    & " WHERE " _
    & " tbm.merge_bus_id = '" & szSplitBusID & "'AND " _
    & " status = '" & 1 & "' "

'    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount = 0 Then ShowError ERR_NoSlitpREBus

'    If Not IsEmpty(dtSlitpDate) Then
        szSql = szSql & " AND bus_date = '" & ToDBDate(dtSlitpDate) & "'"
'    End If
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_NoSlitpREBus
'    szSql = ""
    '得到拆分车次信息
    Set rsTemp = Nothing
    szSql = " SELECT DISTINCT bus.* ,mer.SlitpSeat_NO , mer.SeatTypeOldID,tbv.bus_start_time,tbv.vehicle_type_name ,tbv.check_gate_id " _
        & " ,tbv.total_seat ,tbv.end_station_name FROM Env_bus_info tbv," _
        & "( " _
        & "SELECT tbe.bus_id,tbe.bus_date,tbe.ticket_no,tbe.seat_no as SlitpSeat_no ,tbe.seat_type_id as SeatTypeOldId " _
        & " FROM Env_bus_seat_lst tbe, merge_bus_lst tbm WHERE tbe.bus_id=tbm.merge_bus_id AND "

'    '***********************
'    If szSellStationID <> "" Then
'       szSql = szSql
'    End If
'    '****************************

    If szSplitBusID <> "" Then
        szSql = szSql & " tbm.merge_bus_id= '" & szSplitBusID & " ' AND tbm.status='1' AND  (tbe.status='3' or tbe.status='5') "
    Else
        szSql = szSql & "tbm.status='1' AND (tbe.status='3' or tbe.status='5')"
    End If
    If Not IsNull(dtSlitpDate) Then
        szSql = szSql & "  AND tbe.bus_date='" & ToDBDate(dtSlitpDate) & "'AND tbm.bus_date=tbe.bus_date ) mer,"
    Else
        szSql = szSql & ") mer,"
    End If
    szSql = szSql & "(SELECT tbe.* FROM Env_bus_seat_lst tbe, Mergebus_split_lst tbm WHERE tbe.bus_id= tbm.bus_id  AND "
    If szSplitBusID <> "" Then
        szSql = szSql & " tbm.merge_bus_id= '" & szSplitBusID & " ' AND tbm.status='1' AND tbe.status=" & ST_SeatSlitp & ""
    Else
        szSql = szSql & " tbm.status='1' AND tbe.status=" & ST_SeatSlitp & " "
    End If
    If Not IsNull(dtSlitpDate) Then
        szSql = szSql & " AND tbe.bus_date='" & ToDBDate(dtSlitpDate) & "'AND tbm.bus_date='" & ToDBDate(dtSlitpDate) & "' ) bus WHERE"
    Else
        szSql = szSql & ") bus WHERE "
    End If
    If Not IsNull(dtSlitpDate) Then
        szSql = szSql & " mer.ticket_no = bus.ticket_no AND bus.bus_date='" & ToDBDate(dtSlitpDate) & "'AND tbv.bus_id=bus.bus_id  AND tbv.bus_date=bus.bus_date ORDER BY bus.bus_id"
    Else
        szSql = szSql & " mer.ticket_no = bus.ticket_no tbv.bus_id=bus.bus_id  ORDER BY bus.bus_id"
    End If
    
    
    
    
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_REBusSlitpNotInfo
    ReDim szTemp(1 To rsTemp.RecordCount, 1 To 11)
    For i = 1 To rsTemp.RecordCount
        szTemp(i, 1) = FormatDbValue(rsTemp!bus_id)
        szTemp(i, 2) = FormatDbValue(rsTemp!seat_no)
        szTemp(i, 3) = FormatDbValue(rsTemp!SlitpSeat_NO)
        szTemp(i, 4) = FormatDbValue(rsTemp!ticket_no)
        szTemp(i, 5) = FormatDbValue(rsTemp!check_gate_id)
        szTemp(i, 6) = FormatDbValue(rsTemp!vehicle_type_name)
        szTemp(i, 7) = FormatDbValue(rsTemp!total_seat)
        szTemp(i, 8) = FormatDbValue(rsTemp!end_station_name)
        szTemp(i, 9) = FormatDbValue(rsTemp!bus_start_time)
        szTemp(i, 10) = FormatDbValue(rsTemp!SeatTypeOldID)
        szTemp(i, 11) = FormatDbValue(rsTemp!seat_type_id)
        
        
        rsTemp.MoveNext
    Next
    GetSlitpInfo = szTemp
    Set rsTemp = Nothing
    Set oDb = Nothing
    Exit Function
here:
    Set rsTemp = Nothing
    Set oDb = Nothing
    ShowError err.Number
 End Function
'从环境座位表中的到票号和座位号(拆分用到)
Public Function FromEnvGetTicketAndSeatNo(szBusID As String, dtDate As Date) As String()
Dim oDb As New RTConnection
Dim szSql As String
Dim i As Integer
Dim szaTicket() As String
Dim rsTemp As Recordset
Dim nTemp As Integer
nTemp = ST_TicketCanceled Or ST_TicketChanged Or ST_TicketReturned
AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    szSql = "SELECT seat_no,ticket_no FROM Env_bus_seat_lst WHERE bus_id='" & szBusID & "' AND " _
            & " bus_date='" & ToDBDate(dtDate) & "' AND status='" & ST_SeatSold & "'"
    Set rsTemp = oDb.Execute(szSql)
    '---------------------
    '查询该车次售出的有效票
    If rsTemp.RecordCount = 0 Then ShowError ERR_REBusNoSellTicket
    ReDim szaTicket(1 To rsTemp.RecordCount, 1 To 2) As String
    For i = 1 To rsTemp.RecordCount
        szaTicket(i, 1) = FormatDbValue(rsTemp!seat_no)
        szaTicket(i, 2) = FormatDbValue(rsTemp!ticket_no)
        rsTemp.MoveNext
    Next
    FromEnvGetTicketAndSeatNo = szaTicket
    Set oDb = Nothing
End Function
'从环境座位表中的到票号和座位号和被并车次信息
Public Function GetMergeBus(szBusID As String, dtDate As Date) As String()
Dim oDb As New RTConnection
Dim szSql As String
Dim i As Integer
Dim szaTicket() As String
Dim rsTemp As Recordset
Dim nCount As Integer

    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    szSql = "SELECT * FROM  Env_bus_seat_lst tbe  WHERE tbe.bus_id IN " _
            & "(SELECT merge_bus_id FROM Mergebus_split_lst WHERE bus_id='" & szBusID & "' AND bus_date='" & ToDBDate(dtDate) & "')" _
            & " AND tbe.ticket_no IN " _
            & "(SELECT ticket_no FROM Env_bus_seat_lst WHERE bus_id='" & szBusID & "' AND bus_date='" & ToDBDate(dtDate) & "') " _
            & "AND tbe.bus_date='" & ToDBDate(dtDate) & "' AND tbe.ticket_no<>''"
    Set rsTemp = oDb.Execute(szSql)
    '---------------------
    '查询该车次售出的有效票
    If rsTemp.RecordCount = 0 Then Exit Function
    nCount = rsTemp.RecordCount
    ReDim szaTicket(1 To nCount, 1 To 5) As String
    For i = 1 To rsTemp.RecordCount
        szaTicket(i, 1) = FormatDbValue(rsTemp!bus_id)
        szaTicket(i, 2) = FormatDbValue(rsTemp!bus_date)
        szaTicket(i, 3) = FormatDbValue(rsTemp!ticket_no)
        szaTicket(i, 4) = FormatDbValue(rsTemp!seat_no)
        szaTicket(i, 5) = nCount
        rsTemp.MoveNext
    Next
    GetMergeBus = szaTicket
    Set oDb = Nothing
End Function
Public Function GetCanSallSeat(szBusID As String, dtDate As Date) As Integer
Dim oDb As New RTConnection
Dim szSql As String
Dim rsTemp As Recordset
Dim nTemp As Integer
On Error GoTo here
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    szSql = "SELECT COUNT(seat_no) as SeatCount  FROM Env_bus_seat_lst  WHERE bus_id='" & szBusID & "' AND " _
    & " bus_date='" & ToDBDate(dtDate) & "' AND status='" & ST_SeatCanSell & "'"
    Set rsTemp = oDb.Execute(szSql)
       Set oDb = Nothing
       If rsTemp.RecordCount <> 0 Then
       nTemp = FormatDbValue(rsTemp!seatCount)
       End If
     Set rsTemp = Nothing
     GetCanSallSeat = nTemp
    
    Exit Function
here:
    ShowError err.Number
End Function

Public Function BusSlitpMe(szBusID As String) As String()
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szBusSlitpMe() As String
    Dim i As Integer
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    szSql = "SELECT merge_bus_id FROM Mergebus_split_lst WHERE merge_bus_id='" & m_szBusID & "'AND bus_id='" & szBusID & "' AND  bus_date='" & ToDBDate(m_dtRunDate) & "' AND status='" & 1 & "'"
    Set rsTemp = oDb.Execute(szSql)
    '--------------------------
    '我是好的，未被拆分
    If rsTemp.RecordCount = 0 Then
    Else
    ReDim szBusSlitpMe(1 To rsTemp.RecordCount)
        For i = 1 To rsTemp.RecordCount
        szBusSlitpMe(i) = FormatDbValue(rsTemp!bus_id)
        Next
    End If
    BusSlitpMe = szBusSlitpMe
End Function
'得到全票到其它票的转换参数
'nItemPrice--票价表,
'nPriceType---票的类型(全票,半票,......)
Public Function GetPriceTpye(Optional nPriceType As Integer = 2, Optional nItemPrice As Integer = 0) As Single()
    Dim oDb As New RTConnection
    Dim sgPriceItem(1 To 4) As Single
    Dim rsTemp As Recordset
    Dim szSql As String
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    szSql = "SELECT * FROM Ticket_type_formula_info WHERE table_id='" & nItemPrice & "' AND price_item='" & nPriceType & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    sgPriceItem(1) = FormatDbValue(rsTemp!parameter_1)
    sgPriceItem(2) = FormatDbValue(rsTemp!parameter_2)
    GetPriceTpye = sgPriceItem
End Function

Public Function UnSplit(szSplitBusID As String, dtBusDate As Date) ', Optional szSellStationID As String = "")
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szSqlfield As String
    Dim i As Integer
    Dim szaSlitpInfo() As String
    Dim nCount As Integer
    Dim szBusID As String
    
    Dim nSeatType As Integer
    Dim tTSeatInfoCountEx() As TSeatInfoCountEx
    Dim nCountSeatInfo As Integer
    Dim szBusTemp As String
    Dim szAimBusIdlog As String
    
    AssertHaveRight m_oActiveUser, RIGHT_REBusUnSlitp
    
    '得到拆分车次信息
    szaSlitpInfo = GetSlitpInfo(szSplitBusID, dtBusDate) ', szSellStationID)
    nCount = ArrayLength(szaSlitpInfo)
    If nCount = 0 Then ShowError ERR_NoSlitpREBus
        
    On Error GoTo here
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    '开始事务
    oDb.BeginTrans
        
    szBusTemp = szaSlitpInfo(1, 1)
    szAimBusIdlog = szBusTemp
    ReDim tTSeatInfoCountEx(1 To 1)
    tTSeatInfoCountEx(1).szBusID = szBusTemp
    
    For i = 1 To nCount
        nCountSeatInfo = ArrayLength(tTSeatInfoCountEx)
        If Trim(szBusTemp) <> Trim(szaSlitpInfo(i, 1)) Then
            '保存目标车次ＩＤ
             nCountSeatInfo = nCountSeatInfo + 1
             ReDim Preserve tTSeatInfoCountEx(1 To nCountSeatInfo)
             szBusTemp = szaSlitpInfo(1, 1)
             tTSeatInfoCountEx(nCountSeatInfo).szBusID = szBusTemp
             szAimBusIdlog = szAimBusIdlog & "," & szBusTemp
        End If
        nSeatType = Val(szaSlitpInfo(i, 11))
        '保存目标车次ＩＤ座位数保存
        tTSeatInfoCountEx(nCountSeatInfo).tTSeatTypeInfo(nSeatType).seatCount = tTSeatInfoCountEx(nCountSeatInfo).tTSeatTypeInfo(nSeatType).seatCount + 1
        
        '更新目标车次　Env_bus_seat_lst
        szSql = "UPDATE Env_bus_seat_lst  SET  "
        szSqlfield = "status='" & ST_SeatCanSell & "',"
        szSqlfield = szSqlfield & "ticket_no='" & Null & "' "
        szSqlfield = szSqlfield & " WHERE bus_id='" & szaSlitpInfo(i, 1) & "' AND " _
                                & "Bus_date='" & ToDBDate(dtBusDate) & "' AND " _
                                & "seat_no='" & szaSlitpInfo(i, 2) & "'"
        szSql = szSql & szSqlfield
        oDb.Execute szSql
        
        '更新目标车次　Env_bus_info　的可售座位数
        szSql = "UPDATE Env_bus_info " _
        & " SET sale_seat_quantity=sale_seat_quantity+1 " _
        & " WHERE bus_id='" & szaSlitpInfo(i, 1) & "' " _
        & " AND Bus_date='" & ToDBDate(dtBusDate) & "' "
        
        oDb.Execute szSql
'        If i = 1 Then
'           szsql = "UPDATE Env_bus_info Set "
'           szSqlfield = "is_all_refundment='" & 0 & "',"
'           szSqlfield = szSqlfield & "status='" & ST_BusNormal & "' WHERE Bus_id='" & szSplitBusID & "'AND bus_date='" & ToDBDate(dtBusDate) & "' AND status='" & ST_BusSlitpStop & "'"
'           szsql = szsql & szSqlfield
'           odb.Execute szsql
'           WriteOperateLog m_oActiveUser, RIGHT_REBusUnSlitp, "取消车次拆分" & szSplitBusID & "目标车次" & szaSlitpInfo(i, 1)
'        End If
   Next
    '被拆车次复班
    szSql = "UPDATE Env_bus_info set " _
            & " status='" & ST_BusNormal & "' ," _
            & " is_all_refundment='" & 0 & "' " _
            & " WHERE bus_id='" & szSplitBusID & "'AND bus_date='" & ToDBDate(dtBusDate) & "' "
    oDb.Execute szSql
    
    szSql = " UPDATE Env_bus_allot_lst set " _
        & " status=" & ST_BusNormal & " WHERE bus_id='" & szSplitBusID & "' AND " _
        & " bus_date='" & ToDBDate(dtBusDate) & "'"
    oDb.Execute szSql
    '目标车次状态恢复
    szSql = "UPDATE Env_bus_info set status=status - '" & CnSlitpStatus & "' " _
           & " WHERE bus_id In ( SELECT bus_id FROM Mergebus_split_lst WHERE merge_bus_id='" & szSplitBusID & "' " _
           & " AND status = '" & 1 & "' " _
           & " AND bus_date = '" & ToDBDate(dtBusDate) & "') " _
           & " AND bus_date = '" & ToDBDate(dtBusDate) & "' " _
           & " AND status >= '" & CnSlitpStatus & "'"
    oDb.Execute szSql
    'delete 拆分表
    
'    If szSellStationID <> "" Then
        szSql = "delete Mergebus_split_lst WHERE " _
                & "merge_bus_id = '" & szSplitBusID & "'AND " _
                & "bus_date = " & ToDBDate(dtBusDate) & "'AND " _
                & "status = '" & 1 & "'"
    
'    End If
    
    
    
    oDb.Execute szSql
    
    '目标车次座位重新分配
    For i = 1 To nCountSeatInfo
         UpdateEnviromentSeatCountEx tTSeatInfoCountEx(i).szBusID, dtBusDate, oDb
    Next
    oDb.CommitTrans
    WriteOperateLog m_oActiveUser, RIGHT_REBusUnSlitp, "日期[" & Date & "]车次[" & szSplitBusID & "]被取消拆分" & "目标车次:" & szAimBusIdlog
    Set oDb = Nothing
  
    Exit Function
here:
   oDb.RollbackTrans
   Set oDb = Nothing
   WriteOperateLog m_oActiveUser, RIGHT_REBusUnSlitp, "拆分车次" & szSplitBusID & "有误,操作失败"
   ShowError err.Number, "取消拆分失败，可能车次无拆信息或日期不对，请确认后重新进行该操作"
End Function


'*******************************
'取得车次的座位信息
'*******************************
Public Function GetReBusSeatType() As String()
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim i As Integer
    Dim rsTemp As New Recordset
    Dim LerrCode As Long
    Dim bIsNormalSeatType As Boolean
    Dim szNormalSeatName As String
    Dim szSeatTypeIDTemp As String
    Dim nToalSeatNo As Integer
    Dim nCount As Integer
    Dim nCountTemp  As Integer
    
    Dim szSeatPriceType() As String
    Dim szSeatPriceTypeTemp() As String
    On Error GoTo here
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    
     'm_nTotalSeat
     'm_szVehicle
    szSql = "SELECT  seat_type_name FROM Seat_type_code " _
         & "  WHERE  seat_type_id='" & cszSeatTypeIsNormal & "' "
    Set rsTemp = oDb.Execute(szSql)
    
    '取得普通座位名
    If rsTemp.RecordCount <> 0 Then
       szNormalSeatName = FormatDbValue(rsTemp!seat_type_name)
    End If
    
    '取得车辆座位信息
    szSql = "SELECT  v.*,s.seat_type_name FROM  vehicle_seat_type_info  v,Seat_type_code s  " _
         & "  WHERE v.vehicle_id='" & m_szVehicle & "' AND s.seat_type_id=v.seat_type_id  " _
         & " ORDER BY v.seat_type_id "
    Set rsTemp = oDb.Execute(szSql)
    
    
    If rsTemp.RecordCount = 0 Then
       '只有普通座位
         ReDim szSeatPriceType(1 To 1, 1 To 3)
        szSeatPriceType(1, 1) = cszSeatTypeIsNormal
        szSeatPriceType(1, 2) = szNormalSeatName
        szSeatPriceType(1, 3) = m_nTotalSeat
        bIsNormalSeatType = True
    Else
       '可能有其它座位
       Dim j As Integer
       
       
       For i = 1 To rsTemp.RecordCount
            
            nToalSeatNo = nToalSeatNo + (CInt(FormatDbValue(rsTemp!end_seat_no)) - CInt(FormatDbValue(rsTemp!start_seat_no))) + 1
            
            
            If FormatDbValue(rsTemp!seat_type_id) = cszSeatTypeIsNormal Then bIsNormalSeatType = True
            
            
            If Trim(szSeatTypeIDTemp) <> FormatDbValue(rsTemp!seat_type_id) Then
                
                nCountTemp = 0
                
                nCountTemp = CInt(FormatDbValue(rsTemp!end_seat_no)) - CInt(FormatDbValue(rsTemp!start_seat_no))
                
                szSeatPriceTypeTemp = szSeatPriceType
                
                nCount = ArrayLength(szSeatPriceTypeTemp)
                
                ReDim szSeatPriceType(1 To nCount + 1, 1 To 3)
                
                szSeatTypeIDTemp = FormatDbValue(rsTemp!seat_type_id)
                szSeatPriceType(nCount + 1, 1) = FormatDbValue(rsTemp!seat_type_id)
                szSeatPriceType(nCount + 1, 2) = FormatDbValue(rsTemp!seat_type_name)
                szSeatPriceType(nCount + 1, 3) = nCountTemp
                
                
                For j = 1 To nCount
                   szSeatPriceType(j, 1) = szSeatPriceTypeTemp(j, 1)
                   szSeatPriceType(j, 2) = szSeatPriceTypeTemp(j, 2)
                   szSeatPriceType(j, 3) = szSeatPriceTypeTemp(j, 3)
                Next
               
            Else
               '相同座位多条信息，重计算座位数
               nCountTemp = nCountTemp + (CInt(FormatDbValue(rsTemp!end_seat_no)) - CInt(FormatDbValue(rsTemp!start_seat_no))) + 1
               szSeatPriceType(nCount + 1, 3) = nCountTemp
 
            End If
            
            rsTemp.MoveNext
        
        Next
                  
        '是否需要普通座位
        If nToalSeatNo < m_nTotalSeat Then
          nCountTemp = m_nTotalSeat - nToalSeatNo
          If bIsNormalSeatType = False Then
               nCount = ArrayLength(szSeatPriceType)
              '保存原数组
               szSeatPriceTypeTemp = szSeatPriceType
               ReDim szSeatPriceType(1 To nCount + 1, 1 To 3)
               szSeatPriceType(1, 1) = cszSeatTypeIsNormal
               szSeatPriceType(1, 2) = szNormalSeatName
               szSeatPriceType(1, 3) = nCountTemp
               
               For i = 2 To nCount + 1
                  szSeatPriceType(i, 1) = szSeatPriceTypeTemp(i - 1, 1)
                  szSeatPriceType(i, 2) = szSeatPriceTypeTemp(i - 1, 2)
                  szSeatPriceType(i, 3) = szSeatPriceTypeTemp(i - 1, 3)
               Next
               GetReBusSeatType = szSeatPriceTypeTemp
               Set oDb = Nothing
               Set rsTemp = Nothing
               Exit Function
          End If
        End If
      
    End If
    GetReBusSeatType = szSeatPriceType
    Set oDb = Nothing
    Set rsTemp = Nothing
    Exit Function
here:
    Set oDb = Nothing
    Set rsTemp = Nothing
    err.Raise err.Number
 End Function
'获得车次环境票价信息
'Public Function GetBusTicketInfoEx() As TRETicketPriceEx()
'Public Function GetBusTicketInfo() As TBusPriceDetailInfo()
'    Dim szSql As String
'    'Dim atTicketPrice() As TRETicketPriceEx
'    Dim rsTemp As New Recordset
'    Dim oDb As New RTConnection
'    Dim k As Integer
'    Dim i As Integer
'    Dim project As TSchemeArrangement
'    Dim tBusPriceDetail() As TBusPriceDetailInfo
'    Dim plErr As Long
'
'
'    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
'
'    szSql = "SELECT DISTINCT area_code, ebw.route_id,epw.*,sti.seat_type_name,vehicle_type_name,vehicle_type_code,route_name, " _
'          & " station_name , ebs.mileage ,ebw.bus_type,ebw.bus_start_time " _
'          & " FROM Seat_type_code sti,route_info rfo,Station_info sfo, " _
'
'    If DateDiff("d", m_dtRunDate, Now) <= 0 Then '如果查询的是环境中的车次则使用工作表
'        szSql = szSql & " Work_env_bus_station_lst ebs,Work_env_bus_price_lst epw ,Work_env_bus_info ebw "
'    Else
'        szSql = szSql & " Env_bus_price_lst epw, Env_bus_info ebw  , Env_bus_station_lst ebs  "
'    End If
'    project = SelfGetExecuteBusProject(ToDBDate(m_dtRunDate), plErr)
'
'    szSql = szSql & " WHERE epw.seat_type_id = sti.seat_type_id AND ebw.bus_date = epw.bus_date AND ebw.bus_id = epw.bus_id " _
'        & " AND rfo.route_id=ebw.route_id AND sfo.station_id=epw.station_id AND " _
'        & " ebs.station_id=epw.station_id AND ebs.bus_id=epw.bus_id AND ebs.bus_date=epw.bus_date AND " _
'        & " epw.bus_id='" & m_szBusID & "'  AND epw.bus_date='" & ToDBDate(m_dtRunDate) & "'" _
'        & " ORDER BY  epw.bus_date,epw.bus_id,epw.station_serial_no,epw.seat_type_id,ticket_type "
'
'    Set rsTemp = oDb.Execute(szSql)
'
'    If rsTemp.RecordCount = 0 Then
'        ShowError ERR_REBusNotStationPrice
'        Exit Function
'    End If
'    rsTemp.MoveFirst
'    'ReDim atTicketPrice(1 To rsTemp.RecordCount)
'    ReDim tBusPriceDetail(1 To rsTemp.RecordCount)
'    For i = 1 To rsTemp.RecordCount
'        tBusPriceDetail(i).szAreaCode = FormatDbValue(rsTemp!area_Code)
'        tBusPriceDetail(i).szRouteID = FormatDbValue(rsTemp!route_id)
'        tBusPriceDetail(i).szRouteName = FormatDbValue(rsTemp!route_name)
'        tBusPriceDetail(i).szStationID = FormatDbValue(rsTemp!station_id)
'        tBusPriceDetail(i).szStationName = FormatDbValue(rsTemp!station_name)
'        tBusPriceDetail(i).szSeatTypeID = FormatDbValue(rsTemp!seat_type_id)
'        tBusPriceDetail(i).szSeatTypeName = FormatDbValue(rsTemp!seat_type_name)
'        tBusPriceDetail(i).szVehicleModel = FormatDbValue(rsTemp!vehicle_type_code)
'        tBusPriceDetail(i).szVehicleModelName = FormatDbValue(rsTemp!vehicle_type_name)
'        tBusPriceDetail(i).nTicketType = FormatDbValue(rsTemp!ticket_type)
'        tBusPriceDetail(i).sgMileage = FormatDbValue(rsTemp!mileage)
'        tBusPriceDetail(i).dyStartTime = FormatDbValue(rsTemp!bus_start_time)
'        tBusPriceDetail(i).nBusType = FormatDbValue(rsTemp!bus_type)
'        tBusPriceDetail(i).nSerialNo = FormatDbValue(rsTemp!station_serial_no)
'        tBusPriceDetail(i).szBusID = FormatDbValue(rsTemp!bus_id)
'        tBusPriceDetail(i).sgBaseCarriage = FormatDbValue(rsTemp!base_carriage)
'
'        tBusPriceDetail(i).sgTotalPrice = tBusPriceDetail(i).sgBaseCarriage
'        For k = 1 To 15
'            tBusPriceDetail(i).asgItem(k) = FormatDbValue(rsTemp.Fields("price_item_" & k))
'            tBusPriceDetail(i).sgTotalPrice = tBusPriceDetail(i).sgTotalPrice + tBusPriceDetail(i).asgItem(k)
'        Next
'        rsTemp.MoveNext
'    Next i
'
'    GetBusTicketInfo = tBusPriceDetail
'    Set rsTemp = Nothing
'    Set oDb = Nothing
'End Function

'插入环境车次站点
Public Function EnInsertStationAbandon(szStationID As String)
    Dim szSql As String
    Dim rsTemp As New Recordset
    Dim nSerialNo As Integer
    Dim oDb As New RTConnection
    Dim szEndStationId As String
    Dim LerrCode As Long
    Dim nCount As Integer, i As Integer
    Dim bflg As Boolean
    Dim tExePriceTable As TSchemeArrangement
    Dim szEndstationName As String
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_REBusAddStation
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    '取得插入站点序号
    If DateDiff("d", m_dtRunDate, Now) <= 0 Then '如果查询的是环境中的车次则使用工作表
        szSql = "SELECT DISTINCT station_serial_no FROM Work_env_bus_station_lst " _
            & " WHERE bus_id='" & m_szBusID & "'AND " _
            & " bus_date='" & ToDBDate(m_dtRunDate) & "'" _
            & "AND station_id='" & szStationID & "' "
    Else
        szSql = "SELECT DISTINCT station_serial_no FROM Env_bus_station_lst " _
        & " WHERE bus_id='" & m_szBusID & "'AND " _
        & " bus_date='" & ToDBDate(m_dtRunDate) & "'" _
        & "AND station_id='" & szStationID & "' "
    End If
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_NotInsertSationID
    nSerialNo = FormatDbValue(rsTemp!station_serial_no)
    '取得线路终点站代码
    szSql = "SELECT end_station_id,station_name FROM route_info r,station_info s  " _
        & " WHERE r.route_id ='" & m_szRoute & "' AND  s.station_id=r.end_station_id "
    Set rsTemp = oDb.Execute(szSql)
    szEndStationId = FormatDbValue(rsTemp!end_station_id)
    szEndstationName = FormatDbValue(rsTemp!station_name)
    oDb.BeginTrans
    '终点站
    szSql = "update Env_bus_info set end_station_id='" & szEndStationId & "' ," _
        & "  end_station_name ='" & szEndstationName & "' " _
        & " WHERE bus_id='" & m_szBusID & "' " _
        & " AND bus_date='" & ToDBDate(m_dtRunDate) & "'"
    oDb.Execute szSql

    szSql = "update Env_bus_station_lst set station_serial_no=station_serial_no-1 " _
        & "  WHERE bus_id='" & m_szBusID & "' AND bus_date='" & ToDBDate(m_dtRunDate) & "'AND station_serial_no >'" & nSerialNo & "'"
    oDb.Execute szSql
    szSql = "update Env_bus_price_lst  set station_serial_no=station_serial_no-1 WHERE  bus_id='" & m_szBusID & "'  " _
        & "AND bus_date='" & ToDBDate(m_dtRunDate) & "'AND station_serial_no >'" & nSerialNo & "' "
    oDb.Execute szSql


    szSql = "delete Env_bus_price_lst  WHERE  station_id='" & szStationID & "' " _
        & " AND bus_id='" & m_szBusID & "' AND bus_date='" & ToDBDate(m_dtRunDate) & "'"
    oDb.Execute szSql
    szSql = "delete Env_bus_station_lst WHERE station_id='" & szStationID & "'  " _
        & " AND bus_id='" & m_szBusID & "' AND bus_date='" & ToDBDate(m_dtRunDate) & "'"
    oDb.Execute szSql
    oDb.CommitTrans
    Set oDb = Nothing
    Set rsTemp = Nothing

    WriteOperateLog m_oActiveUser, RIGHT_REBusAddStation, "车次" & m_szBusID & "发车时间为[" & Left(ToDBDate(m_dtRunDate), 19) & "]终点站[" & m_szOldEndStationName & "]::环境站点[" & szStationID & "]被删除"
    #If SUPPORT_MSG = 1 Then
        Dim oParam As New SystemParam
        Dim aszValues(1 To 2) As String
        aszValues(1) = m_szBusID
        aszValues(2) = ToDBDate(m_dtRunDate)
        m_oSndMsg.SendMsg oParam.NowDateTime, cszInsertReBusStation, aszValues
    #End If
    Exit Function
here:
    oDb.RollbackTrans
    Set oDb = Nothing
    Set rsTemp = Nothing
 End Function
'********************************************************
'环境中修改座位---
'如环境中无该座位票价,而票价表中有该票价则向环镜新增票价
'**参数：szSeat(1)-座位类型，szSeat(2)-起始座号，szSeat(3)-结束座号
'*********************************************************
Public Sub ReModifySeatType(szSeat() As String)
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim i As Integer
    Dim szPrictTableName As String
    Dim rsTemp As Recordset
    Dim project As TSchemeArrangement
    Dim LerrCode As Long
    Dim lErrCode2 As Long
    Dim szSeatTemp As String
    Dim bflg As Boolean
    
    
    Dim noldSeaType As Integer
    Dim nNewSeaType As Integer
    Dim ttOldSeatType(1 To 5) As TSeatInfoCount
    Dim ttNewSeatType(1 To 5) As TSeatInfoCount
    Dim szSeatWhere As String
    Dim nSeatNOCount As Long
    Dim szlog As String
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_ReBusSeatManagement
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    On Error GoTo here
    '检查环境站点表是否有改座位类型信息
    If DateDiff("d", m_dtRunDate, Now) <= 0 Then
        szSql = "SELECT DISTINCT seat_type_id FROM Work_env_bus_station_lst WHERE bus_date='" & ToDBDate(m_dtRunDate) & "' AND bus_id='" & m_szBusID & "' " ' AND seat_type_id='" & szSeat(1) & "'  "
    Else
        szSql = "SELECT DISTINCT seat_type_id FROM Env_bus_station_lst WHERE bus_date='" & ToDBDate(m_dtRunDate) & "' AND bus_id='" & m_szBusID & "' " 'AND seat_type_id='" & szSeat(1) & "'  "
    End If
    Set rsTemp = oDb.Execute(szSql)
    
    If rsTemp.RecordCount <> 0 Then
        For i = 1 To rsTemp.RecordCount
            szSeatTemp = FormatDbValue(rsTemp!seat_type_id)
            If Trim(szSeatTemp) = Trim(szSeat(1)) Then
                bflg = True
                Exit For
            End If
            rsTemp.MoveNext
        Next
    End If
    
'    project = SelfGetExecuteBusProject(ToDBDate(m_dtRunDate), lErrCode2)
    If bflg = False Then
        '取得执行计划
        szPrictTableName = BusProjectExecutePrice(m_dtRunDate, LerrCode)
        If LerrCode = 1 Then ShowError ERR_NoExePriceTable
        '判断是否有该座位类型的票价
        szSql = "SELECT  * FROM bus_price_lst WHERE price_table_id='" & szPrictTableName & "' " _
            & " AND bus_id='" & m_szBusID & "' AND seat_type_id='" & Trim(szSeat(1)) & "' " _
            & " AND vehicle_type_code='" & m_szVehicleModelID & " '"
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount = 0 Then ShowError ERR_NotPriceSeatType
    End If
    
    '开始事务
    oDb.BeginTrans
    '  odb.CommitTrans
    If bflg = False Then
        '-------
'        '写入环境车次站点表
'        szSql = " INSERT Env_bus_station_lst " _
'            & " ( " _
'            & " bus_date , bus_id,sell_station_id, station_id, seat_type_id, station_serial_no, max_sale_quantity, sale_ticket_quantity, book_count " _
'            & " , stop_sale_time , mileage , full_price , half_price , preferential_ticket1 , preferential_ticket2 , preferential_ticket3 " _
'            & " ) " _
'            & " SELECT " _
'            & " '" & ToDBDate(m_dtRunDate) & "' , g.bus_id ,g.sell_station_id, g.station_id  '" & szSeat(1) & "' ,station_serial_no , max_sale_quantity , max_sale_quantity , max_sale_quantity " _
'            & " , stop_sale_time , mileage , ISNULL(f.ticket_price_total,0) , ISNULL(h.ticket_price_total,0) , ISNULL(p1.ticket_price_total,0) , ISNULL(p2.ticket_price_total,0) , ISNULL(p3.ticket_price_total,0) " _
'            & " FROM Env_bus_station_lst g , " _
'            & " ( SELECT bus_id , station_id ,sell_station_id , base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
'            & "     + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
'            & "     + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
'            & " WHERE ticket_type = " & TP_FullPrice & " AND price_table_id = '" & Trim(szPrictTableName) & "' AND vehicle_type_code = '" & m_szVehicleModelID & "' AND bus_id = '" & m_szBusID & "' AND seat_type_id  =  '" & Trim(szSeat(1)) & "' " _
'            & "     ) f , " _
'            & " ( SELECT bus_id , station_id ,sell_station_id, base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
'            & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
'            & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
'            & " WHERE ticket_type = " & TP_HalfPrice & "  AND price_table_id = '" & Trim(szPrictTableName) & "' AND vehicle_type_code = '" & m_szVehicleModelID & "' AND bus_id = '" & m_szBusID & "' AND seat_type_id  =  '" & Trim(szSeat(1)) & "' " _
'            & " ) h , " _
'            & " ( SELECT bus_id , station_id ,sell_station_id,base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
'            & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
'            & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
'            & " WHERE ticket_type = " & TP_PreferentialTicket1 & "  AND price_table_id = '" & Trim(szPrictTableName) & "' AND vehicle_type_code = '" & m_szVehicleModelID & "' AND bus_id = '" & m_szBusID & "' AND seat_type_id  =  '" & Trim(szSeat(1)) & "' " _
'            & " ) p1 , " _
'            & " ( SELECT bus_id , station_id ,sell_station_id, base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 "
'
'        szSql = szSql _
'            & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
'            & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
'            & " WHERE ticket_type = " & TP_PreferentialTicket2 & "  AND price_table_id = '" & Trim(szPrictTableName) & "' AND vehicle_type_code = '" & m_szVehicleModelID & "' AND bus_id = '" & m_szBusID & "' AND seat_type_id  =  '" & Trim(szSeat(1)) & "' " _
'            & " ) p2 , " _
'            & " ( SELECT bus_id , station_id ,sell_station_id, base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
'            & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
'            & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
'            & " WHERE ticket_type = " & TP_PreferentialTicket3 & "  AND price_table_id = '" & Trim(szPrictTableName) & "' AND vehicle_type_code = '" & m_szVehicleModelID & "' AND bus_id = '" & m_szBusID & "' AND seat_type_id  =  '" & Trim(szSeat(1)) & "' " _
'            & " ) p3 " _
'            & " WHERE g.project_id= '" & Trim(project.szProjectID) & "' AND g.bus_id = '" & m_szBusID & "' AND g.seat_type_id='" & Trim(szSeatTemp) & " '" _
'            & " AND g.bus_id *= f.bus_id AND g.bus_id *= h.bus_id AND g.bus_id *= p1.bus_id AND g.bus_id *= p2.bus_id AND g.bus_id *=p3.bus_id " _
'            & " AND g.station_id *= f.station_id AND g.station_id *= h.station_id AND g.station_id *= p1.station_id AND g.station_id *= p2.station_id AND g.station_id *=p3.station_id "
'
        '********************
        szSql = " INSERT Env_bus_station_lst " _
            & " ( " _
            & " bus_date , bus_id,sell_station_id, station_id, seat_type_id, station_serial_no, max_sale_quantity, sale_ticket_quantity, book_count " _
            & " , stop_sale_time , mileage , full_price , half_price , preferential_ticket1 , preferential_ticket2 , preferential_ticket3 " _
            & " ) " _
            & " SELECT " _
            & " '" & ToDBDate(m_dtRunDate) & "' , g.bus_id ,g.sell_station_id, g.station_id  '" & szSeat(1) & "' ,station_serial_no , max_sale_quantity , max_sale_quantity , max_sale_quantity " _
            & " , stop_sale_time , mileage , ISNULL(f.ticket_price_total,0) , ISNULL(h.ticket_price_total,0) , ISNULL(p1.ticket_price_total,0) , ISNULL(p2.ticket_price_total,0) , ISNULL(p3.ticket_price_total,0) " _
            & " FROM Env_bus_station_lst g , " _
            & " ( SELECT bus_id , station_id ,sell_station_id , base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
            & "     + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
            & "     + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
            & " WHERE ticket_type = " & TP_FullPrice & " AND price_table_id = '" & Trim(szPrictTableName) & "' AND vehicle_type_code = '" & m_szVehicleModelID & "' AND bus_id = '" & m_szBusID & "' AND seat_type_id  =  '" & Trim(szSeat(1)) & "' " _
            & "     ) f , " _
            & " ( SELECT bus_id , station_id ,sell_station_id, base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
            & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
            & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
            & " WHERE ticket_type = " & TP_HalfPrice & "  AND price_table_id = '" & Trim(szPrictTableName) & "' AND vehicle_type_code = '" & m_szVehicleModelID & "' AND bus_id = '" & m_szBusID & "' AND seat_type_id  =  '" & Trim(szSeat(1)) & "' " _
            & " ) h , " _
            & " ( SELECT bus_id , station_id ,sell_station_id,base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
            & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
            & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
            & " WHERE ticket_type = " & TP_PreferentialTicket1 & "  AND price_table_id = '" & Trim(szPrictTableName) & "' AND vehicle_type_code = '" & m_szVehicleModelID & "' AND bus_id = '" & m_szBusID & "' AND seat_type_id  =  '" & Trim(szSeat(1)) & "' " _
            & " ) p1 , " _
            & " ( SELECT bus_id , station_id ,sell_station_id, base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 "
        
        szSql = szSql _
            & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
            & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
            & " WHERE ticket_type = " & TP_PreferentialTicket2 & "  AND price_table_id = '" & Trim(szPrictTableName) & "' AND vehicle_type_code = '" & m_szVehicleModelID & "' AND bus_id = '" & m_szBusID & "' AND seat_type_id  =  '" & Trim(szSeat(1)) & "' " _
            & " ) p2 , " _
            & " ( SELECT bus_id , station_id ,sell_station_id, base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
            & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
            & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
            & " WHERE ticket_type = " & TP_PreferentialTicket3 & "  AND price_table_id = '" & Trim(szPrictTableName) & "' AND vehicle_type_code = '" & m_szVehicleModelID & "' AND bus_id = '" & m_szBusID & "' AND seat_type_id  =  '" & Trim(szSeat(1)) & "' " _
            & " ) p3 " _
            & " WHERE    g.bus_id = '" & m_szBusID & "' AND g.seat_type_id='" & Trim(szSeatTemp) & " '" _
            & " AND g.bus_id *= f.bus_id AND g.bus_id *= h.bus_id AND g.bus_id *= p1.bus_id AND g.bus_id *= p2.bus_id AND g.bus_id *=p3.bus_id " _
            & " AND g.station_id *= f.station_id AND g.station_id *= h.station_id AND g.station_id *= p1.station_id AND g.station_id *= p2.station_id AND g.station_id *=p3.station_id "
        
        oDb.Execute szSql
        '写入环境车次票价表
        
        szSql = "INSERT Env_bus_price_lst " _
            & " (  " _
            & " bus_date , bus_id,,sell_station_id, station_id, ticket_type, seat_type_id, station_serial_no, ticket_price_total, base_carriage " _
            & " , price_item_1 , price_item_2 , price_item_3 , price_item_4 , price_item_5 , price_item_6 , price_item_7 , price_item_8 " _
            & " , price_item_9 , price_item_10 , price_item_11 , price_item_12 , price_item_13 , price_item_14 , price_item_15 " _
            & " ) " _
            & " SELECT '" & ToDBDate(m_dtRunDate) & "' , " _
            & " bus_id , station_id,sell_station_id, ticket_type, seat_type_id, station_serial_no " _
            & " , base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
            & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13" _
            & " + price_item_14 + price_item_15 ticket_price_total " _
            & " , base_carriage , price_item_1 , price_item_2 , price_item_3 , price_item_4 , price_item_5 , price_item_6 " _
            & " , price_item_7 , price_item_8 , price_item_9 , price_item_10 , price_item_11 , price_item_12 , price_item_13 " _
            & " , price_item_14 , price_item_15 " _
            & " FROM bus_price_lst " _
            & " WHERE price_table_id = '" & Trim(szPrictTableName) & "' AND vehicle_type_code = '" & m_szVehicleModelID & "' AND bus_id = '" & m_szBusID & "' AND seat_type_id  =  '" & Trim(szSeat(1)) & "' "
        oDb.Execute szSql
        '-------
    End If
    
    '更新环境座位表
    For i = CInt(szSeat(2)) To CInt(szSeat(3))
        szSeatWhere = szSeatWhere & "'" & Format(i, "00") & "',"
    Next
    szSeatWhere = Left(szSeatWhere, Len(szSeatWhere) - 1)
    
    szSql = "UPDATE Env_bus_seat_lst SET seat_type_id='" & Trim(szSeat(1)) & "' WHERE seat_no IN ( " & szSeatWhere & ") AND " _
        & "bus_date='" & ToDBDate(m_dtRunDate) & "'AND bus_id ='" & m_szBusID & "'" _
        & " AND (status= '" & ST_SeatCanSell & "' or status= '" & ST_SeatReserved & "' or status= '" & ST_SeatProjectBooked & "' )"
    oDb.Execute szSql
    
    
    
    UpdateEnviromentSeatCountEx m_szBusID, m_dtRunDate, oDb
    oDb.CommitTrans
    Set oDb = Nothing
    
    Set rsTemp = Nothing
    szlog = "车次" & m_szBusID & "发车时间为[" & ToDBDateTime(m_dtOldStartupTime) & "]终点站[" & m_szOldEndStationName & "]:修改座位类型:从" & szSeat(2) & "到" & szSeat(3) & "改为" & szSeat(1)
    WriteOperateLog m_oActiveUser, RIGHT_ReBusSeatManagement, szlog
    
    Exit Sub
here:
    oDb.RollbackTrans
    Set oDb = Nothing
    Set rsTemp = Nothing
    WriteOperateLog m_oActiveUser, RIGHT_ReBusSeatManagement, "修改车次" & m_szBusID & "座位类型失败" & szSeat(1)
    err.Raise err.Number
End Sub
Private Function FindRightPriceItem(prsPriceInfo As Recordset, ByVal pszStationID As String, ByVal pnTicketType As ETicketType, Optional pbFROMFirst As Boolean = False) As Integer
    Dim szMsg As String
    Dim m_lTemp As Long
    
    On Error GoTo Error_Handle
    If pbFROMFirst Then
        prsPriceInfo.MoveFirst
    Else
        If prsPriceInfo.Bookmark <> 1 Then prsPriceInfo.MovePrevious
    End If
    
    Do While Not prsPriceInfo.EOF
        If FormatDbValue(prsPriceInfo!station_id) = pszStationID And FormatDbValue(prsPriceInfo!ticket_type) = pnTicketType Then
            Exit Function
        End If
        prsPriceInfo.MoveNext
    Loop
Error_Handle:
    szMsg = "站点" & pszStationID & "票种信息"
'    If pnTicketType = TP_FullPrice Then
'        szMsg = szMsg & "全票"
'    Else
'        szMsg = szMsg & "半票"
'    End If
    ShowError ERR_NotStationTicketTypePrice, , szMsg
End Function
'
'得到车辆运行序号
Public Function GetCurrentBusvehicleNo() As Integer
  Dim rsTemp As New Recordset
  Dim saTemp As TSchemeArrangement
  Dim LerrCode As Long
  Dim nBusRunCycle As Integer
  Dim nCurrentBusvehicleNo As Integer
  Dim nBusRunFirstNo As Integer
  Dim oDb As New RTConnection
  Dim szSql As String
  AssertObjIsValid
  oDb.ConnectionString = GetConnectionStr(cszRunEnv)
'  saTemp = SelfGetExecuteBusProject(m_dtRunDate, LerrCode)
  If LerrCode <> 0 Then ShowError ERR_REBusNotExecuteProject
   '得到车次运行周期
'  szSql = "SELECT bus_run_cycle,run_start_serial FROM bus_info  WHERE bus_id='" & m_szBusID & "' AND " _
'                    & "project_id='" & saTemp.szProjectID & "' "
    
  szSql = "SELECT bus_run_cycle,run_start_serial FROM bus_info  WHERE bus_id='" & m_szBusID & "'"

  Set rsTemp = oDb.Execute(szSql)
  nBusRunCycle = FormatDbValue(rsTemp!bus_run_cycle)
  nBusRunFirstNo = FormatDbValue(rsTemp!run_start_serial)
  nCurrentBusvehicleNo = (DateDiff("d", saTemp.dtBeginDate, m_dtRunDate) Mod nBusRunCycle + nBusRunFirstNo - 1) Mod nBusRunCycle + 1
'  GetCurrentBusvehicleNo = nCurrentBusvehicleNo
End Function

Public Function GetStationPrice(szStationID As String, szSeatType() As String) As TRETicketPriceEx()
Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp   As Recordset
    Dim tTReTicketPrice() As TRETicketPriceEx
    Dim i As Integer, j As Integer, k As Integer
    Dim szPriceItem As String
    Dim nCount As Integer, nSeatCount As Integer
'    Dim saTemp As TSchemeArrangement
    Dim LerrCode As Long
    Dim szBusID As String
    Dim szBusIdTemp As String
    Dim szPriceTable As String
    Dim nCountTemp As Integer
    
    
    AssertObjIsValid
    
    
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    ReDim tTReTicketPrice(0 To 0)
'    saTemp = SelfGetExecuteBusProject(ToDBDate(m_dtRunDate), LerrCode)
    If LerrCode <> 0 Then ShowError 19219
    
    
    szPriceTable = BusProjectExecutePrice(m_dtRunDate, LerrCode)
    If LerrCode <> 0 Then ShowError 19219
    
    nCount = ArrayLength(szSeatType)
    
    For k = 1 To nCount
        szSql = "SELECT p.*,s.seat_type_name,ssi.sell_station_name  FROM bus_price_lst p,sell_station_info ssi, " _
               & " Seat_type_code s WHERE price_table_id='" & szPriceTable & "' AND p.station_id='" & szStationID & "'" _
               & " AND   vehicle_type_code='" & m_szVehicleModelID & "' AND s.seat_type_id=p.seat_type_id and p.sell_station_id=ssi.sell_station_id " _
               & " AND  s.seat_type_id='" & szSeatType(k, 1) & "'"
        Set rsTemp = oDb.Execute(szSql)
        
        If rsTemp.RecordCount = 0 Then ReDim tTReTicketPrice(0 To 0): Exit Function
        
        szBusIdTemp = FormatDbValue(rsTemp!bus_id)
        i = 0
        Dim dbPricetotal As Single
        Do While szBusIdTemp = FormatDbValue(rsTemp!bus_id)
              nSeatCount = ArrayLength(tTReTicketPrice)
              ReDim Preserve tTReTicketPrice(0 To nSeatCount)
              dbPricetotal = 0
              For j = 1 To 15
                 szPriceItem = "price_item_" & j
                 tTReTicketPrice(nSeatCount).asgPrice(j) = FormatDbValue(rsTemp.Fields(szPriceItem))
                 dbPricetotal = dbPricetotal + tTReTicketPrice(nSeatCount).asgPrice(j)
              Next
              
              tTReTicketPrice(nSeatCount).nTicketType = FormatDbValue(rsTemp!ticket_type)
              tTReTicketPrice(nSeatCount).sgBase = FormatDbValue(rsTemp!base_carriage)
              dbPricetotal = dbPricetotal + tTReTicketPrice(nSeatCount).sgBase
              tTReTicketPrice(nSeatCount).sgTotal = dbPricetotal
              tTReTicketPrice(nSeatCount).szStationID = FormatDbValue(rsTemp!station_id)
              tTReTicketPrice(nSeatCount).sgMileage = FormatDbValue(rsTemp!mileage)
              tTReTicketPrice(nSeatCount).szSeatType = FormatDbValue(rsTemp!seat_type_id)
              tTReTicketPrice(nSeatCount).szSeatTypeName = FormatDbValue(rsTemp!seat_type_name)
              tTReTicketPrice(nSeatCount).szSellStationID = FormatDbValue(rsTemp!sell_station_id)
              tTReTicketPrice(nSeatCount).szSellStationName = FormatDbValue(rsTemp!sell_station_name)
              rsTemp.MoveNext
              i = i + 1
              If i >= rsTemp.RecordCount Then
                 Exit Do
              End If
             
            Loop
    Next
    
    GetStationPrice = tTReTicketPrice
    
    Set oDb = Nothing
    Set rsTemp = Nothing
End Function
'该接口不用
'Public Function GetReBusTicketInfoFromStationID(szStationID As String) As TRETicketPriceEx()
'End Function

'插入环境站点
Public Function EnInsertStation(szStationSerial As Integer, tTReTicketPrice() As TRETicketPriceEx, Optional bflg As Boolean = False)
    Dim oDb As New RTConnection
    Dim sztempsql As String
    Dim dtBusDate As Date
    Dim rsTemp As Recordset
    Dim rsTempType As Recordset
    Dim rsTempSellStation As Recordset
    Dim i As Integer
    Dim k As Integer
    Dim szSql As String
    Dim nCount As Integer
    Dim nCountasgPrice As Integer
    Dim tTRestation() As TBusStationSellInfo
    '为了调用
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_REBusAddStation
    On Error GoTo here

    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    nCount = ArrayLength(tTReTicketPrice)
    nCountasgPrice = ArrayLength(tTReTicketPrice(1).asgPrice)
    sztempsql = "SELECT * FROM Work_env_bus_price_lst tbp,Work_env_bus_station_lst tbb WHERE " _
                & " tbb.bus_id='" & m_szBusID & "' AND " _
                & " tbb.bus_date='" & ToDBDate(m_dtRunDate) & "' AND " _
                & " tbp.bus_id=tbb.bus_id AND " _
                & " tbp.bus_date=tbb.bus_date AND " _
                & " tbb.station_id='" & tTReTicketPrice(1).szStationID & "' AND " _
                & " tbp.station_id=tbb.station_id "
    Set rsTemp = oDb.Execute(sztempsql)
    If rsTemp.RecordCount <> 0 Then ShowError ERR_REEXitStation
    
    szSql = "SELECT sell_station_id FROM env_bus_station_lst  WHERE " _
                & " bus_id='" & m_szBusID & "' AND " _
                & " bus_date='" & ToDBDate(m_dtRunDate) & "'" _
                & " GROUP BY sell_station_id "
    Set rsTempSellStation = oDb.Execute(szSql)
    

    oDb.BeginTrans
    If bflg = True Then
         '插入的是终点站
        Dim szStationName As String
        sztempsql = "SELECT station_name FROM station_info WHERE  station_id='" & tTReTicketPrice(1).szStationID & "'"
        Set rsTemp = oDb.Execute(sztempsql)
        szStationName = FormatDbValue(rsTemp!station_name)
        sztempsql = "UPDATE Env_bus_info SET  end_station_id=" & tTReTicketPrice(1).szStationID & "," _
                    & " end_station_name='" & szStationName & "' " _
                    & " WHERE bus_id='" & m_szBusID & "'AND " _
                    & " bus_date='" & ToDBDate(m_dtRunDate) & "'"
        oDb.Execute sztempsql
    Else
        '更新原站点序号
        sztempsql = "UPDATE Env_bus_price_lst SET station_serial_no= station_serial_no+1" _
                    & " WHERE  bus_id='" & m_szBusID & "'  AND " _
                    & "  bus_date='" & ToDBDate(m_dtRunDate) & "' AND  station_serial_no >= '" & szStationSerial & "'"
        oDb.Execute sztempsql

        sztempsql = "UPDATE Env_bus_station_lst SET   station_serial_no= station_serial_no+1" _
                    & " WHERE  bus_id='" & m_szBusID & "'  AND " _
                    & "  bus_date='" & ToDBDate(m_dtRunDate) & "'   AND station_serial_no >= '" & szStationSerial & "'"
        oDb.Execute sztempsql
    End If
    '将站点插入车次站点表
'    For k = 1 To rsTempSellStation.RecordCount
    tTRestation = FormatReTicketToStationTable(tTReTicketPrice)
    For i = 1 To ArrayLength(tTRestation) - 1
    sztempsql = "INSERT Env_bus_station_lst (" _
             & "sell_station_id," _
             & "bus_date," _
             & "bus_id," _
             & " seat_type_id,  " _
             & "mileage," _
             & "station_id, " _
             & "station_serial_no, " _
             & "max_sale_quantity, " _
             & "sale_ticket_quantity, " _
             & "book_count," _
             & "stop_sale_time," _
             & "full_price," _
             & "half_price, " _
             & "preferential_ticket1," _
             & "preferential_ticket2," _
             & "preferential_ticket3)"
     sztempsql = sztempsql & " Values ( " _
              & "'" & tTRestation(i).szSellStationID & "'," _
              & "'" & ToDBDate(m_dtRunDate) & "'," _
              & " '" & m_szBusID & "'," _
              & " '" & tTRestation(i).szSeatTypeID & "'," _
              & " '" & tTRestation(i).nMileage & "'," _
              & "'" & tTRestation(i).szStationID & "', " _
              & "'" & szStationSerial & "', " _
              & "'-1', " _
              & "'-1' , " _
              & "'0' , " _
              & "'-1' , " _
              & "'" & tTRestation(i).sgFullPrice & "', " _
              & "'" & tTRestation(i).sgHalfPrice & "', " _
              & "'" & tTRestation(i).sgPreferentialPrice1 & "', " _
              & "'" & tTRestation(i).sgPreferentialPrice2 & "', " _
              & "'" & tTRestation(i).sgPreferentialPrice3 & "')"
    oDb.Execute sztempsql
    Next i
'    rsTempSellStation.MoveNext
'    Next k
     '将站点插入环境票价表
     Dim szSqlFild As String
     Dim szSqlSet As String
     Dim j As Integer
'    rsTempSellStation.MoveFirst
'    For k = 1 To rsTempSellStation.RecordCount
    For i = 1 To nCount
         szSqlSet = ""
         szSqlFild = ""
         For j = 1 To nCountasgPrice
            szSqlFild = szSqlFild & ",price_item_" & j
            szSqlSet = szSqlSet & ", '" & tTReTicketPrice(i).asgPrice(j) & "'"
         Next
        sztempsql = ""
         sztempsql = "INSERT Env_bus_price_lst (" _
                & "sell_station_id," _
                & "bus_date," _
                & "bus_id," _
                & " seat_type_id , " _
                & "station_serial_no," _
                & "base_carriage," _
                & "ticket_price_total," _
                & "station_id " _
                & " " & szSqlFild & "  ," _
                & "ticket_type)"
        sztempsql = sztempsql & " VALUES ( " _
                & "'" & tTReTicketPrice(i).szSellStationID & "'," _
                & " '" & ToDBDate(m_dtRunDate) & "'," _
                & " '" & m_szBusID & "'," _
                & "'" & tTReTicketPrice(i).szSeatType & "'," _
                & "'" & szStationSerial & "'," _
                & "'" & tTReTicketPrice(i).sgBase & "'," _
                & "'" & tTReTicketPrice(i).sgTotal & "'," _
                & " '" & tTReTicketPrice(i).szStationID & "'  " _
                & "" & szSqlSet & " ," _
                & " '" & tTReTicketPrice(i).nTicketType & "')"
        oDb.Execute sztempsql
   Next
'   rsTempSellStation.MoveNext
'   Next k

    WriteOperateLog m_oActiveUser, RIGHT_REBusAddStation, "车次[" & m_szBusID & "]发车时间为" & ToDBDateTime(m_dtOldStartupTime) & "]终点站[" & m_szOldEndStationName & "]:环境中插入站点:序号为" & szStationSerial

    oDb.CommitTrans
    Set oDb = Nothing
    Set rsTempType = Nothing
    Set rsTemp = Nothing
    EnInsertStation = True
    Exit Function
here:
  oDb.RollbackTrans
  ShowError err.Number
End Function
''*****************************
''
''****************************
Public Function GetItemParam(szExecute As String) As THalfTicketItemParamEX()
    Dim szSql As String, lRowCount As Long
    Dim oDb As New RTConnection
    Dim i As Integer, rsTemp As Recordset, j As Integer
    Dim aTHTemp() As THalfTicketItemParamEX
    Dim eTemp As ETicketType
    Dim nCount As Integer
    szSql = "SELECT * FROM Ticket_type_formula_info  WHERE table_id ='" & szExecute & "' AND ticket_Type IN (  SELECT DISTINCT  ticket_Type_id  FROM Ticket_type_code  WHERE ticket_type_valid = '1' )  ORDER BY ticket_type , price_item"
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    Set rsTemp = oDb.Execute(szSql, lRowCount)
On Error GoTo 0
    If lRowCount = 0 Then
        ShowError ERR_NotPriceItemParam
    Else
        nCount = rsTemp.RecordCount / 16
        ReDim aTHTemp(1 To nCount)
        For j = 1 To nCount
            aTHTemp(j).sTicketType = FormatDbValue(rsTemp!ticket_type)
            For i = 1 To 16
                aTHTemp(j).asgParam1(i) = FormatDbValue(rsTemp!parameter_1)
                aTHTemp(j).asgParam2(i) = FormatDbValue(rsTemp!parameter_2)
                aTHTemp(j).szAnnotation(i) = FormatDbValue(rsTemp!Annotation)
                rsTemp.MoveNext
            Next i
       Next j
     End If
     GetItemParam = aTHTemp
     Set rsTemp = Nothing
     Set oDb = Nothing
End Function

'***************************************************************************************
'***
'*** 功能： 环境新增车次
'***        (1)  如果存在车型、车次种类和线路相同的其他车次票价,复制该票价而不管线路票价
'***        (2)  如果线路票价表中存在该车型，线路票价 (不考虑车次种类不同票价不同的情况)
'***************************************************************************************
Public Function AddObject()
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim sztempsql As String
    Dim i As Integer
    Dim saTemp As TSchemeArrangement
    Dim oPara As New SystemParam
    Dim bEndSationCansale As Boolean
    '车次信息
    Dim szVehicelModelID As String
    Dim szVehicleModelName As String
    Dim nStartSeat As Integer
    Dim szOwnerID As String
    Dim szCorp As String
    Dim szSplitCompanyID As String
    Dim szRouteID As String
    
    '车次票价表
    Dim szPriceTable As String
'    Dim szProjectID As String
    
    '用于插入环境表的变量声明
    Dim nTotalSeat As Integer
    Dim szEndStationId As String
    Dim szEndstationName As String
    
    Dim szCheckGateID As String
    Dim nRegisterStatus As Integer
    Dim nIsAllRefundment As Integer
    Dim szTransportCompanyID As String
    Dim szVehicleID As String
    
    
    Dim szExePriceTable As String
    Dim LerrCode As Long
    Dim szBusTemp As String
    Dim bHaveOldBusPriceflg As Boolean
    Dim bHaveOldRoutePriceflg As Boolean
    
    
    
    ''''''''
    Dim tSeatTypeInfo(1 To 5) As TSeatInfoCount
    Dim nSeatType As Integer
    Dim nSeatCountVehicl As Integer
    Dim rstempSeat As New Recordset
    Dim szSeatWhere As String
    Dim szSeatType() As String
    Dim nSeatTypeUser As Integer
    
    '''''''
    Dim j As Integer, nCount As Integer
    Dim k As Integer
    Dim h As Integer
    Dim tvTemp() As TVehcileSeatType
    ''''
    Dim szAddBusFix  As String '加班车前缀
    Dim nSeatCountTemp As Integer
    Dim bIsNormBusSeatType As Boolean
    Dim szset As String
    Dim nTemp2 As Integer
    Dim szSeatTypeID As String
    Dim szSeatTypeTemp As String
    Dim nCountSeatType As Integer
    Dim rsTemp1 As New Recordset
    Dim nString As Integer
    
    '************************************
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_REBusAdd
    oPara.Init m_oActiveUser
    '加班车前缀
    szAddBusFix = oPara.AdditionBusPreFix
    
    'szAddBusFix = Left(szAddBusFix, 2)
    
    If Trim(m_szBusID) = "" Or DateDiff("d", oPara.NowDate, m_dtRunDate) < 0 Then ShowError ERR_REBusAddNewNotOk
    
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    
    '车次是否存在
    szSql = "SELECT * FROM Work_env_bus_info WHERE bus_id='" & m_szBusID & "' AND bus_date='" & ToDBDate(m_dtRunDate) & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 0 Then ShowError ERR_REBusAdd
    '加班车最后站点可售
    bEndSationCansale = oPara.EndStationCanSale
    '得到计划
   ' saTemp = SelfGetExecuteBusProject(m_dtRunDate, LerrCode)
    If LerrCode <> 0 Then ShowError ERR_REBusNotExecuteProject
    '获得执行的票价表
    szExePriceTable = BusProjectExecutePrice(m_dtRunDate, LerrCode)
    If LerrCode = 1 Then ShowError ERR_NoExePriceTable
    
    '查询新增车辆的车辆信息和确定车次运行的线路及其票价表
    '车辆信息
    szSql = "SELECT tbt.*,tbv.vehicle_type_short_name FROM Vehicle_info tbt ,vehicle_type_code tbv WHERE vehicle_id='" & m_szVehicle & "' AND tbt.vehicle_type_code=tbv.vehicle_type_code"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_REBusVehicelNotExist
    
    '车辆信息
    nStartSeat = FormatDbValue(rsTemp!start_seat_no)
    nTotalSeat = FormatDbValue(rsTemp!seat_quantity)
    szVehicleModelName = FormatDbValue(rsTemp!vehicle_type_short_name)
    szVehicelModelID = FormatDbValue(rsTemp!vehicle_type_code)
    szOwnerID = FormatDbValue(rsTemp!owner_id)
    szCorp = FormatDbValue(rsTemp!transport_company_id)
    szSplitCompanyID = FormatDbValue(rsTemp!split_company_id)
    szSql = "SELECT * FROM vehicle_seat_type_info WHERE vehicle_id='" & m_szVehicle & "'"
    
    Set rstempSeat = oDb.Execute(szSql)
    '车辆座位的类型
    If rstempSeat.RecordCount <> 0 Then
        ReDim tvTemp(1 To rstempSeat.RecordCount)
        
        For i = 1 To rstempSeat.RecordCount
            tvTemp(i).szVehcileID = m_szVehicle
            tvTemp(i).szSeatTypeID = rstempSeat!seat_type_id
            tvTemp(i).szStartSeatNo = rstempSeat!start_seat_no
            tvTemp(i).szEndSeatNo = rstempSeat!end_seat_no
            nSeatType = Val(rstempSeat!seat_type_id)
            nSeatCountTemp = Val(rstempSeat!end_seat_no) - Val(rstempSeat!start_seat_no) + 1
            tSeatTypeInfo(nSeatType).SeatTypeID = rstempSeat!seat_type_id
            tSeatTypeInfo(nSeatType).seatCount = tSeatTypeInfo(nSeatType).seatCount + nSeatCountTemp
            nSeatCountVehicl = nSeatCountVehicl + nSeatCountTemp
            rstempSeat.MoveNext
        Next i
    
    End If
    
    If nSeatCountVehicl < nTotalSeat Then
        If tSeatTypeInfo(1).seatCount = 0 Then
            bIsNormBusSeatType = True
        End If
        tSeatTypeInfo(1).seatCount = nTotalSeat - nSeatCountVehicl
        tSeatTypeInfo(1).SeatTypeID = cszSeatTypeIsNormal
    End If
    
    For i = 1 To 5
        If tSeatTypeInfo(i).seatCount > 0 Then
            nSeatTypeUser = ArrayLength(szSeatType)
            If nSeatTypeUser = 0 Then
                ReDim szSeatType(1 To 1)
                szSeatType(1) = tSeatTypeInfo(i).SeatTypeID
                szSeatWhere = szSeatWhere & " '" & tSeatTypeInfo(i).SeatTypeID & "',"
            Else
                ReDim Preserve szSeatType(1 To nSeatTypeUser + 1)
                szSeatType(nSeatTypeUser + 1) = tSeatTypeInfo(i).SeatTypeID
                szSeatWhere = szSeatWhere & " '" & tSeatTypeInfo(i).SeatTypeID & "',"
            End If
        End If
    Next
    
    szSeatWhere = Left(szSeatWhere, Len(szSeatWhere) - 1)
    Set rstempSeat = Nothing
    szSql = " SELECT * FROM bus_price_lst WHERE bus_id ='" & m_szBusID & "'price_table_id ='" & szExePriceTable & "'AND vehicle_type_code='" & szVehicelModelID & "' AND seat_type_id IN ( " & szSeatWhere & " )"
    If rsTemp.RecordCount = 0 Then ShowError ERR_REBusNewVehicleNowRoutePrice
    szBusTemp = FormatDbValue(rsTemp!bus_id)
    '存在车型、车次种类和线路相同的其他车次票价
    bHaveOldBusPriceflg = True
    '取得线路信息
    szSql = "SELECT r.end_station_id,s.station_name FROM Route_info r ,station_info  s WHERE r.Route_id='" & m_szRoute & "' AND r.status='0' AND s.station_id=r.end_station_id"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_NoRoute
    szEndStationId = FormatDbValue(rsTemp!end_station_id)
    szEndstationName = FormatDbValue(rsTemp!station_name)
    '开始事务
    On Error GoTo here
    oDb.BeginTrans
    '插入环境车次表信息（Env_bus_info）
    szSql = "INSERT Env_bus_info( " _
        & "bus_date," _
        & "bus_id," _
        & "check_gate_id," _
        & "vehicle_type_name," _
        & "vehicle_type_code," _
        & "total_seat," _
        & "sale_seat_quantity," _
        & "total_stand_quantity," _
        & "sale_stand_seat_quantity," _
        & "bus_start_time," _
        & "bus_type," _
        & "status," _
        & "route_id," _
        & "end_station_id," _
        & "end_station_name," _
        & "owner_id," _
        & "transport_company_id," _
        & "split_company_id, " _
        & "ScrollBus_Check_Time, " _
        & "internet_status, " _
        & "vehicle_id) "
    szSql = szSql & " VALUES( '" _
        & ToDBDate(m_dtRunDate) & "','" _
        & m_szBusID & "','" _
        & m_szCheckGate & "','" _
        & szVehicleModelName & "','" _
        & szVehicelModelID & "'," _
        & nTotalSeat & "," _
        & nTotalSeat & "," _
        & 0 & "," _
        & 0 & "," _
        & "'" & ToDBDateTime(m_dtStartupTime) & "'," _
        & m_nBusType & "," _
        & ST_BusStopped & ",'" _
        & m_szRoute & "','" _
        & szEndStationId & "','" _
        & szEndstationName & "','" _
        & Trim(szOwnerID) & "','" _
        & Trim(szCorp) & "','" _
        & Trim(szSplitCompanyID) & "'," _
        & m_nScrollBusCheckTime & ",'" _
        & m_nInternetStatus & "','" _
        & m_szVehicle & "')"
    oDb.Execute szSql
    '插入环境座位表信息（Env_bus_seat_lst）
    If m_nBusType <> TP_ScrollBus Then
    
        For i = 1 To nTotalSeat
            nTemp2 = i - 1 + nStartSeat
            
            '取座位类型
            szSeatTypeID = GetSeatTypeID(nTemp2, tvTemp)
            If szSeatTypeID = "" Then szSeatTypeID = cszSeatTypeIsNormal
            szSql = "INSERT Env_bus_seat_lst( " _
                & "bus_date," _
                & "bus_id," _
                & "seat_no," _
                & "status," _
                & "ticket_no," _
                & "seat_type_id) " _
                & " VALUES('" _
                & ToDBDate(m_dtRunDate) & "','" _
                & m_szBusID & "','" _
                & BuildSeatNo(nTemp2) & "'," _
                & ST_SeatCanSell & ",'','" _
                & szSeatTypeID & "')"
            oDb.Execute szSql
        
        Next i
    End If
    '插入环境车次站点表信息（enviroment_bus_station_info）和环境票价表信息（Env_bus_price_lst）
    szSql = "SELECT DISTINCT seat_type_id  FROM vehicle_seat_type_info WHERE vehicle_id='" & m_szVehicle & "' ORDER BY seat_type_id"
    Set rsTemp1 = oDb.Execute(szSql)
    nCountSeatType = rsTemp1.RecordCount
    If bIsNormBusSeatType = True Then
        nCountSeatType = nCountSeatType + 1
    End If
    For j = 1 To nCountSeatType
        If bIsNormBusSeatType = True Then
            szSeatTypeTemp = cszSeatTypeIsNormal
        End If
        If bIsNormBusSeatType = False Then
            szSeatTypeTemp = Trim(rsTemp1!seat_type_id)
        End If
        '写入环境车次站点表
        szSql = " INSERT Env_bus_station_lst " _
            & " ( " _
            & " bus_date , bus_id, sell_station_id,station_id, seat_type_id, station_serial_no, max_sale_quantity, sale_ticket_quantity, book_count " _
            & " , stop_sale_time , mileage , full_price , half_price , preferential_ticket1 , preferential_ticket2 , preferential_ticket3 " _
            & " ) " _
            & " SELECT " _
            & " '" & ToDBDate(m_dtRunDate) & "'as bus_date , '" & m_szBusID & " ' as bus_id , g.sell_station_id, g.station_id  , '" & cszSeatTypeIsNormal & "' ,station_serial_no , max_sale_quantity , max_sale_quantity , max_sale_quantity " _
            & " , stop_sale_time , mileage , ISNULL(f.ticket_price_total,0) , ISNULL(h.ticket_price_total,0) , ISNULL(p1.ticket_price_total,0) , ISNULL(p2.ticket_price_total,0) , ISNULL(p3.ticket_price_total,0) "
        szSql = szSql & " FROM " _
            & " ( SELECT '" & m_szBusID & " ' as bus_id ,sell_station_id, station_id, station_serial_no, max_sale_quantity ,stop_sale_time ,mileage " _
            & "  FROM Bus_station_info WHERE bus_id='" & szBusTemp & "'" _
            & " ) g , " _
            & " ( SELECT '" & m_szBusID & " ' as bus_id ,sell_station_id, station_id  , base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
            & "     + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
            & "     + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
            & " WHERE ticket_type = " & TP_FullPrice & " AND price_table_id = '" & Trim(szExePriceTable) & "' AND vehicle_type_code = '" & szVehicelModelID & "' AND bus_id = '" & szBusTemp & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' " _
            & "     ) f , " _
            & " ( SELECT  '" & m_szBusID & " '  as bus_id ,sell_station_id, station_id , base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
            & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
            & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
            & " WHERE ticket_type = " & TP_HalfPrice & "  AND price_table_id = '" & Trim(szExePriceTable) & "' AND vehicle_type_code = '" & szVehicelModelID & "' AND bus_id = '" & szBusTemp & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' " _
            & " ) h , " _
            & " ( SELECT  '" & m_szBusID & " '  as bus_id ,sell_station_id, station_id , base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
            & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
            & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
            & " WHERE ticket_type = " & TP_PreferentialTicket1 & "  AND price_table_id = '" & Trim(szExePriceTable) & "' AND vehicle_type_code = '" & szVehicelModelID & "' AND bus_id = '" & szBusTemp & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' " _
            & " ) p1 , " _
            & " ( SELECT  '" & m_szBusID & " '  as bus_id ,sell_station_id, station_id , base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 "
        szSql = szSql _
            & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
            & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
            & " WHERE ticket_type = " & TP_PreferentialTicket2 & "  AND price_table_id = '" & Trim(szExePriceTable) & "' AND vehicle_type_code = '" & szVehicelModelID & "' AND bus_id = '" & szBusTemp & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' " _
            & " ) p2 , " _
            & " ( SELECT   '" & m_szBusID & " '  as bus_id ,sell_station_id, station_id , base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
            & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
            & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
            & " WHERE ticket_type = " & TP_PreferentialTicket3 & "  AND price_table_id = '" & Trim(szExePriceTable) & "' AND vehicle_type_code = '" & szVehicelModelID & "' AND bus_id = '" & szBusTemp & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' " _
            & " ) p3 " _
            & " WHERE  g.bus_id = '" & m_szBusID & " ' " _
            & " AND g.bus_id *= f.bus_id AND g.bus_id *= h.bus_id AND g.bus_id *= p1.bus_id AND g.bus_id *= p2.bus_id AND g.bus_id *=p3.bus_id " _
            & " AND g.station_id *= f.station_id AND g.station_id *= h.station_id AND g.station_id *= p1.station_id AND g.station_id *= p2.station_id AND g.station_id *=p3.station_id "
        
       
        
        
        
        
        
        oDb.Execute szSql
        '写入环境车次票价表
        szSql = "INSERT Env_bus_price_lst " _
            & " (  " _
            & " bus_date , bus_id ,sell_station_id, station_id, ticket_type, seat_type_id, station_serial_no, ticket_price_total, base_carriage " _
            & " , price_item_1 , price_item_2 , price_item_3 , price_item_4 , price_item_5 , price_item_6 , price_item_7 , price_item_8 " _
            & " , price_item_9 , price_item_10 , price_item_11 , price_item_12 , price_item_13 , price_item_14 , price_item_15 " _
            & " ) " _
            & " SELECT '" & ToDBDate(m_dtRunDate) & "' , " _
            & " '" & m_szBusID & " ' as   bus_id ,sell_station_id, station_id, ticket_type, seat_type_id, station_serial_no " _
            & " , base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
            & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13" _
            & " + price_item_14 + price_item_15 ticket_price_total " _
            & " , base_carriage , price_item_1 , price_item_2 , price_item_3 , price_item_4 , price_item_5 , price_item_6 " _
            & " , price_item_7 , price_item_8 , price_item_9 , price_item_10 , price_item_11 , price_item_12 , price_item_13 " _
            & " , price_item_14 , price_item_15 " _
            & " FROM bus_price_lst " _
            & " WHERE price_table_id = '" & Trim(szExePriceTable) & "' AND vehicle_type_code = '" & szVehicelModelID & "' AND bus_id = '" & szBusTemp & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' "
        oDb.Execute szSql
        
        If bIsNormBusSeatType = True Then
            bIsNormBusSeatType = False
        Else
            rsTemp1.MoveNext
        End If
    Next
    
    If bEndSationCansale = True Then
        If szAddBusFix <> "" Then
            nString = Len(szAddBusFix) - 1
            If Trim(Left(Trim(m_szBusID), nString)) = Trim(Left(szAddBusFix, nString)) Then
                '0----站点不可售
                szSql = "Update Env_bus_station_lst set max_sale_quantity=0 WHERE station_serial_no< ( SELECT Max(station_serial_no) FROM Env_bus_station_lst  " _
                    & "  WHERE bus_id='" & m_szBusID & "' AND bus_date='" & ToDBDate(m_dtRunDate) & "'AND max_sale_quantity=-1)  " _
                    & "  bus_id='" & m_szBusID & "' AND bus_date='" & ToDBDate(m_dtRunDate) & "'"
                oDb.Execute szSql
            End If
        End If
    End If
    UpdateEnviromentSeatCountEx m_szBusID, m_dtRunDate, oDb
    '结束事务
    oDb.CommitTrans
    WriteOperateLog m_oActiveUser, RIGHT_REBusAdd, "新增车次" & m_szBusID & "发车时间[" & m_dtStartupTime & "]终点站[" & m_szOldEndStationName & "]:"
    '更新座位信息 Env_bus_info
    Set rsTemp = Nothing
    Set oPara = Nothing
    
    Exit Function
    
here:
    oDb.RollbackTrans
    Set rsTemp = Nothing
    Set oPara = Nothing
    err.Raise err.Number
End Function
'***********************
'  -环境车次票价
'***********************
'Public Function GetREBusPrice() As TRETicketPriceEx()
'Dim i As Integer, j As Integer
'Dim oDb As New RTConnection
'Dim rsTemp As New Recordset
'Dim tREBusPrice() As TRETicketPriceEx
'Dim szSql As String
'Dim nCount As String
'AssertObjIsValid
'oDb.ConnectionString = GetConnectionStr(cszRunEnv)
'
'If DateDiff("d", m_dtRunDate, Now) <= 0 Then
'    szSql = "SELECT e.* ,station_name,seat_type_name,t.ticket_type_name" _
'            & " FROM Work_env_bus_price_lst e ,station_info,Seat_type_code ,Ticket_type_code t " _
'            & "  WHERE  e.bus_date='" & ToDBDate(m_dtRunDate) & "' AND bus_id='" & m_szBusID & "' " _
'            & "  AND e.seat_type_id= Seat_type_code.seat_type_id " _
'            & "  AND e.station_id=station_info.station_id  AND t.ticket_type_id=e.ticket_type" _
'            & "  ORDER BY e.station_serial_no,station_name,e.seat_type_id,ticket_type"
'Else
'    szSql = "SELECT e.* ,station_name,seat_type_name,t.ticket_type_name " _
'            & " FROM Env_bus_price_lst e ,station_info,Seat_type_code ,Ticket_type_code t " _
'            & "  WHERE  e.bus_date='" & ToDBDate(m_dtRunDate) & "' AND bus_id='" & m_szBusID & "'  " _
'            & "  AND e.seat_type_id= Seat_type_code.seat_type_id " _
'            & "  AND e.station_id=station_info.station_id AND t.ticket_type_id=e.ticket_type " _
'            & "  ORDER BY e.station_serial_no,station_name,e.seat_type_id,ticket_type"
'End If
'
'Set rsTemp = oDb.Execute(szSql)
'nCount = rsTemp.RecordCount
'If nCount = 0 Then Exit Function
'ReDim tREBusPrice(1 To nCount)
'
'For i = 1 To nCount
'     For j = 1 To 15
'         tREBusPrice(i).asgPrice(j) = rsTemp.Fields("price_item_" & j)
'     Next j
'     tREBusPrice(i).nTicketType = FormatDbValue(rsTemp!ticket_type)
'     tREBusPrice(i).sgBase = FormatDbValue(rsTemp!base_carriage)
'     tREBusPrice(i).sgTotal = FormatDbValue(rsTemp!ticket_price_total)
'     tREBusPrice(i).szSeatTypeName = FormatDbValue(rsTemp!seat_type_name)
'     tREBusPrice(i).szStationName = FormatDbValue(rsTemp!station_name)
'     tREBusPrice(i).szTicketName = FormatDbValue(rsTemp!ticket_type_name)
'     rsTemp.MoveNext
'Next
'GetREBusPrice = tREBusPrice
'Set rsTemp = Nothing
'Set oDb = Nothing
'End Function
'
' 'addRebus
'Public Function GetBusVehicleSeatType(szVehicle As String) As TVehcileSeatType()
'Dim tvTemp() As TVehcileSeatType
' tvTemp = GetBusVehicleSeatTypeEx(szVehicle)
' GetBusVehicleSeatType = tvTemp
'End Function
 

'Re addbus
Public Function GetPriceTable(dtDate As Date) As String
Dim szPriceTableID  As String
Dim oDb As New RTConnection
Dim szSql As String
Dim rsTemp As Recordset
    
    oDb.ConnectionString = GetConnectionStr("")
    
'    szSql = " SELECT * FROM project_execute_lst WHERE start_time=(SELECT MAX(start_time) FROM " _
'            & " project_execute_lst WHERE start_time<='" & ToDBDateTime(dtDate) & "')"
'    Set rsTemp = oDB.Execute(szSql)
'    If rsTemp.RecordCount = 0 Then ShowError ERR_REBusNotExecuteProject
    
'    szPriceTableID = FormatDbValue(rsTemp!project_id)
    
    szSql = "SELECT price_table_id FROM price_table_abstract_info WHERE  " _
            & "start_run_time=(SELECT MAX(start_run_time) FROM price_table_abstract_info WHERE " _
            & " convert(datetime,convert(char(10),start_run_time,101))<='" & dtDate & "' )"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_NoExePriceTable
    szPriceTableID = FormatDbValue(rsTemp!price_table_id)
    GetPriceTable = Trim(szPriceTableID)
    
End Function
Public Function GetBusProject(pdtDate As Date) As String
'Dim tvTemp As TSchemeArrangement
'Dim plErrCode As Long
'
'tvTemp = SelfGetExecuteBusProject(pdtDate, plErrCode)
'If plErrCode <> 0 Then ShowError ERR_REBusNotExecuteProject
'GetBusProject = tvTemp.szProjectID
'
End Function
Public Function DeleteBus(szBusID As String, szInfo As String)
  If szInfo = "AddNewBus" Then
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szExePricertable As String
    Dim LerrCode As Long
    Dim tvTemp As TSchemeArrangement
    Dim plErrCode As Long
   ' AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_REBusAdd
'    tvTemp = SelfGetExecuteBusProject(Now, plErrCode)
    If plErrCode <> 0 Then ShowError ERR_REBusNotExecuteProject
    oDb.ConnectionString = GetConnectionStr("")
    szExePricertable = BusProjectExecutePrice(Now, LerrCode)
    
'    szSql = "DELETE bus_info WHERE bus_id='" & szBusID & "' AND project_id='" & tvTemp.szProjectID & "'"
    
    szSql = "DELETE bus_info WHERE bus_id='" & szBusID & "'"
    oDb.Execute (szSql)
    szSql = "DELETE bus_price_lst WHERE bus_id='" & szBusID & "' AND price_table_id='" & szExePricertable & "'"
    oDb.Execute (szSql)
            #If SUPPORT_MSG = 1 Then
    Dim oParam As New SystemParam
    Dim aszValues(1 To 3) As String
    aszValues(1) = m_szBusID
    aszValues(2) = ToDBDate(m_dtRunDate)
    aszValues(3) = m_nTotalSeat
    m_oSndMsg.SendMsg oParam.NowDateTime, cszRemoveBus, aszValues
    #End If
    WriteOperateLog m_oActiveUser, RIGHT_REBusAdd, "环境新增车次[2][" & m_szBusID & "]"
  End If
End Function
Public Function ReBusSlipLock(bflg As Boolean) As Boolean
    Dim szSql As String
    Dim szTemp As String
    Dim oDb As New RTConnection
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_REBusStopBus
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    
    If Me.busStatus = ST_BusChecking Or Me.busStatus = ST_BusExtraChecking Or _
        Me.busStatus = ST_BusStopCheck Or _
        Me.busStatus = ST_BusStopped Or _
        Me.busStatus = ST_BusMergeStopped Or _
        Me.busStatus = ST_BusSlitpStop Then ReBusSlipLock = False: Exit Function
    
    On Error GoTo here
    If bflg = True Then
    '拆分锁定
    szSql = "UPDATE Env_bus_info SET status=status + " & CnSlitpStatus & " WHERE bus_id='" & m_szBusID & "' AND " _
    & " bus_date='" & ToDBDate(m_dtRunDate) & "' AND status<" & CnSlitpStatus & ""
     oDb.Execute szSql
       If m_eBusStatus < CnSlitpStatus Then
          m_eBusStatus = m_eBusStatus + 32
       End If
    
    Else
    '拆分锁释放
    szSql = "UPDATE Env_bus_info SET status=status - " & CnSlitpStatus & " WHERE bus_id='" & m_szBusID & "' AND " _
            & " bus_date='" & ToDBDate(m_dtRunDate) & "' AND status>=" & CnSlitpStatus & " "
    oDb.Execute szSql
    If m_eBusStatus >= CnSlitpStatus Then
          m_eBusStatus = m_eBusStatus - 32
     End If
    End If
    #If SUPPORT_MSG = 1 Then
    Dim oParam As New SystemParam
    Dim aszValues(1 To 2) As String
    aszValues(1) = m_szBusID
    aszValues(2) = ToDBDate(m_dtRunDate)
    m_oSndMsg.SendMsg oParam.NowDateTime, cszSlipBusLock, aszValues
    
    #End If
    
    Set oDb = Nothing
    
    
    ReBusSlipLock = True
    
    Exit Function
here:
    Set oDb = Nothing
   
   
    err.Raise err.Number
End Function

Private Function ReBusSlipLockEx(szBusID As String, bflg As Boolean) As Boolean
  Dim szSql As String
    Dim szTemp As String
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    

    On Error GoTo here
    If bflg = True Then
    '拆分锁定
    szSql = "UPDATE Env_bus_info SET status=status + " & CnSlitpStatus & " WHERE bus_id='" & szBusID & "' AND " _
    & " bus_date='" & ToDBDate(m_dtRunDate) & "' AND status<" & CnSlitpStatus & ""
       oDb.Execute szSql
       If m_eBusStatus < CnSlitpStatus Then
          m_eBusStatus = m_eBusStatus + 32
       End If
    
    Else
    
    '拆分锁释放
    szSql = "UPDATE Env_bus_info SET status=status - " & CnSlitpStatus & " WHERE bus_id='" & szBusID & "' AND " _
            & " bus_date='" & ToDBDate(m_dtRunDate) & "' AND status>=" & CnSlitpStatus & " "
      
      oDb.Execute szSql
      If m_eBusStatus >= CnSlitpStatus Then
          m_eBusStatus = m_eBusStatus - 32
       End If
    
    End If
    
    
    
    
    #If SUPPORT_MSG = 1 Then
    
    Dim oParam As New SystemParam
    Dim aszValues(1 To 2) As String
    aszValues(1) = m_szBusID
    aszValues(2) = ToDBDate(m_dtRunDate)
    m_oSndMsg.SendMsg oParam.NowDateTime, cszSlipBusLock, aszValues
    
    #End If
    
    Set oDb = Nothing
    ReBusSlipLockEx = True
    Exit Function
here:
    Set oDb = Nothing
    err.Raise err.Number
End Function

''分配座位
'Private Function MaketSeatTypeInfo(ByRef tSeatTypeInfo() As TSeatInfoCount) As Boolean
'End Function

Public Function GetBusSaleSeatInfo() As String()
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As New Recordset
    Dim szDate() As String
    Dim i As Integer
    Dim nCount As Integer
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    On Error GoTo here
    szSql = "    SELECT s.*,tbs.sell_station_name, u.user_name ,si.station_name,t.ticket_type_name  ,s.des_station_id,st.seat_type_name" _
            & "  FROM Ticket_sell_lst s,user_info u ,station_info si ,Seat_type_code st,Ticket_type_code t,sell_station_info tbs" _
            & "  WHERE s.bus_id = '" & m_szBusID & "' " _
            & "  AND t.ticket_type_id =s.ticket_type " _
            & "  AND s.sell_station_id = tbs.sell_station_id " _
            & "  AND  s.bus_date = '" & ToDBDate(m_dtRunDate) & "' " _
            & "  AND s.des_station_id = si.station_id " _
            & "  AND st.seat_type_id = s.seat_type_id " _
            & "  AND u.user_id = s.user_id " _
            & "  AND ( s.status=" & ST_TicketSellChange & " or s.status=" & ST_TicketNormal & " or s.status=" & ST_TicketNormal & " | " & ST_TicketChecked & "  or s.status=" & ST_TicketSellChange & " | " & ST_TicketChecked & "  ) " _
            & "  ORDER BY s.seat_no"
            
      Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    
    If nCount <> 0 Then
       ReDim szDate(1 To nCount, 1 To 19)
        For i = 1 To nCount
             szDate(i, 1) = FormatDbValue(rsTemp!station_name)
             szDate(i, 2) = FormatDbValue(rsTemp!user_name)
             szDate(i, 3) = FormatDbValue(rsTemp!operation_time)
             szDate(i, 4) = FormatDbValue(rsTemp!seat_type_name)
             szDate(i, 5) = FormatDbValue(rsTemp!seat_no)
             szDate(i, 6) = FormatDbValue(rsTemp!ticket_price)
             szDate(i, 7) = FormatDbValue(rsTemp!ticket_type_name)
             szDate(i, 8) = FormatDbValue(rsTemp!ticket_id)
             szDate(i, 9) = FormatDbValue(rsTemp!sale_ticket_type)
             szDate(i, 10) = FormatDbValue(rsTemp!des_station_id)
             szDate(i, 11) = FormatDbValue(rsTemp!Status)

             szDate(i, 12) = FormatDbValue(rsTemp!sell_station_id)
             szDate(i, 13) = FormatDbValue(rsTemp!sell_station_name)
           rsTemp.MoveNext
        Next
        
        GetBusSaleSeatInfo = szDate
    
    End If
    
    Set oDb = Nothing
    Exit Function
here:
    err.Raise err.Number
    Set oDb = Nothing
End Function
'Public Function GetFromEnvBusBusID(szStationID As String, Optional bflg As Boolean) As String()
'    Dim szSql As String
'    Dim oDb As New RTConnection
'    Dim rsTemp As New Recordset
'    Dim szDate() As String
'    Dim i As Integer
'    Dim nCount As Integer
'    Dim szWHERE As String
'    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
'    On Error GoTo here
'    If bflg = True Then
'        szWHERE = " bw.status='" & ST_BusNormal & "'ORDER BY esw.bus_id"
'    Else
'        szWHERE = " ORDER BY esw.bus_id"
'    End If
'    szSql = " SELECT esw.bus_id,bw.vehicle_type_name,bw.route_id,r.route_name,bw.vehicle_type_code,bw.total_seat," _
'        & " bw.sale_seat_quantity,bw.seat_remain,bw.bed_remain , " _
'        & " bw.additional_remain , bw.other_remain_1, bw.other_remain_2 " _
'        & " FROM Work_env_bus_station_lst  esw,Work_env_bus_info bw , route_info r " _
'        & " WHERE r.route_id=bw.route_id AND  esw.bus_id=bw.bus_id " _
'        & " AND esw.bus_date=bw.bus_date AND esw.bus_date='" & ToDBDate(m_dtRunDate) & "' " _
'        & " AND station_id='" & szStationID & "' "
'    szSql = szSql & szWHERE
'    Set rsTemp = oDb.Execute(szSql)
'    nCount = rsTemp.RecordCount
'    If nCount <> 0 Then
'        ReDim szDate(1 To nCount, 1 To 12)
'        For i = 1 To nCount
'            szDate(i, 1) = FormatDbValue(rsTemp!bus_id)
'            szDate(i, 2) = FormatDbValue(rsTemp!vehicle_type_name)
'            szDate(i, 3) = FormatDbValue(rsTemp!route_id)
'            szDate(i, 4) = FormatDbValue(rsTemp!route_name)
'            szDate(i, 5) = FormatDbValue(rsTemp!vehicle_type_code)
'            szDate(i, 6) = FormatDbValue(rsTemp!total_seat)
'            szDate(i, 7) = FormatDbValue(rsTemp!sale_seat_quantity)
'            szDate(i, 8) = FormatDbValue(rsTemp!seat_remain)
'            szDate(i, 9) = FormatDbValue(rsTemp!bed_remain)
'            szDate(i, 10) = FormatDbValue(rsTemp!additional_remain)
'            szDate(i, 11) = FormatDbValue(rsTemp!other_remain_1)
'            szDate(i, 12) = FormatDbValue(rsTemp!other_remain_2)
'            rsTemp.MoveNext
'        Next
'        GetFromEnvBusBusID = szDate
'    End If
'
'    Set oDb = Nothing
'    Exit Function
'here:
'
'    err.Raise err.Number
'    Set oDb = Nothing
'
'End Function
'
'
'Public Function GetEnBusStationInfo() As String()
'    Dim szSql As String
'    Dim oDb As New RTConnection
'    Dim rsTemp As New Recordset
'    Dim szDate() As String
'    Dim i As Integer
'    Dim nCount As Integer
'    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
'    On Error GoTo here
'   szSql = "SELECT w.bus_id,w.station_id,s.station_input_code,s.station_name ,w.max_sale_quantity FROM Work_env_bus_station_lst w," _
'            & " station_info s WHERE w.bus_id='" & m_szBusID & "' AND w.bus_date='" & ToDBDate(m_dtRunDate) & "' " _
'            & " AND w.station_id =s.station_id order BY w.station_serial_no "
'   Set rsTemp = oDb.Execute(szSql)
'   nCount = rsTemp.RecordCount
'
'   If nCount <> 0 Then
'       ReDim szDate(1 To nCount, 1 To 5)
'
'        For i = 1 To nCount
'
'             szDate(i, 1) = FormatDbValue(rsTemp!bus_id)
'             szDate(i, 2) = FormatDbValue(rsTemp!station_id)
'             szDate(i, 3) = FormatDbValue(rsTemp!station_name)
'             szDate(i, 4) = FormatDbValue(rsTemp!max_sale_quantity)
'             szDate(i, 5) = FormatDbValue(rsTemp!station_input_code)
'             rsTemp.MoveNext
'        Next
'        GetEnBusStationInfo = szDate
'    End If
'    Set oDb = Nothing
'    Exit Function
'here:
'    err.Raise err.Number
'    Set oDb = Nothing
'End Function


Public Function MegerBusAndSlitpBus(szBusInfo() As String, szBusID As String) As String()
    Dim szSql As String
    Dim szSqlfield As String
    Dim szSqlContext As String
    Dim oDb As New RTConnection
    Dim rsTemp As New Recordset
    Dim rsTemp1 As New Recordset
    Dim szDate() As String
    Dim i As Integer
    Dim nCount As Integer
    Dim szlog As String
    Dim lRow As Long
    Dim szRutunData() As String
    Dim nSeatType As Integer
    Dim tSeatTypeInfo(1 To 5) As TSeatInfoCount

    #If SUPPORT_MSG = 1 Then
    Dim oParam As New SystemParam
    Dim aszValues() As String
    #End If
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_ReBusSlitp
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    On Error GoTo here
    nCount = ArrayLength(szBusInfo)
    ReDim szRutunData(1 To nCount, 1 To 4)
    
    szSql = "SELECT * FROM Work_env_bus_info WHERE bus_id = " & TransFieldValueToString(szBusID) & " AND bus_date = " & TransFieldValueToString(m_dtRunDate)
    Set rsTemp = oDb.Execute(szSql)
    
    oDb.BeginTrans
        '更新车次状态
        szSql = "UPDATE Env_bus_info set status='" & ST_BusSlitpStop & "' WHERE bus_id ='" & m_szBusID & "' AND bus_date='" & ToDBDate(m_dtRunDate) & "'"
        oDb.Execute szSql
        '得到座位数
        szSql = "  SELECT TOP " & nCount & "  s.*,t.seat_type_name FROM Env_bus_seat_lst s,Seat_type_code t WHERE s.bus_id ='" & szBusID & "' AND  " _
        & " s.bus_date='" & ToDBDate(m_dtRunDate) & "' AND s.status='" & ST_SeatCanSell & "' AND t.seat_type_id =s.seat_type_id"
        Set rsTemp = oDb.Execute(szSql)
        '如果座位数不正确,则出错
        If rsTemp.RecordCount < nCount Then ShowError ERR_SeSeatErr
        For i = 1 To nCount
            szlog = szlog & szBusInfo(i) & ","
            szRutunData(i, 1) = FormatDbValue(rsTemp!seat_no)
            'szRutunData(i, 2) = FormatDbValue(szBusInfo(i, 2))
            szRutunData(i, 3) = szBusInfo(i)
            szRutunData(i, 4) = FormatDbValue(rsTemp!seat_type_name)
            nSeatType = Val(FormatDbValue(rsTemp!seat_type_id))
            tSeatTypeInfo(nSeatType).seatCount = tSeatTypeInfo(nSeatType).seatCount - 1
            szSql = " SELECT COUNT(*) as countTemp FROM Env_bus_seat_lst WHERE  bus_id ='" & szBusID & "' AND  " _
            & " bus_date='" & ToDBDate(m_dtRunDate) & "' AND  ticket_no='" & szBusInfo(i) & "'"
            Set rsTemp1 = oDb.Execute(szSql)
            If rsTemp1!countTemp = 0 Then
                szSql = "UPDATE Env_bus_seat_lst set ticket_no='" & szBusInfo(i) & "',status='" & ST_SeatSlitp & "' WHERE bus_id ='" & szBusID & "' AND  " _
                & " bus_date='" & ToDBDate(m_dtRunDate) & "' AND seat_no = '" & FormatDbValue(rsTemp!seat_no) & "' "


                oDb.Execute szSql
                rsTemp.MoveNext
            End If
        Next
        szSql = "SELECT * FROM Mergebus_split_lst WHERE bus_id ='" & szBusID & " ' AND merge_bus_id ='" & m_szBusID & "' AND bus_date= '" & ToDBDate(m_dtRunDate) & "'"
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount <> 0 Then
            szSql = "delete  Mergebus_split_lst WHERE bus_id ='" & szBusID & " ' AND merge_bus_id ='" & m_szBusID & "' AND bus_date= '" & ToDBDate(m_dtRunDate) & "'"
            oDb.Execute szSql
        End If
        szlog = Left(szlog, Len(szlog) - 1)
        szSql = "INSERT Mergebus_split_lst ("
        szSqlfield = "bus_date,"
        szSqlfield = szSqlfield & "bus_id,"
        szSqlfield = szSqlfield & "merge_bus_id,"
        szSqlfield = szSqlfield & "status ,"
        'szSqlfield = szSqlfield & "slitp_Bus_Info,"
        szSqlfield = szSqlfield & "annotation )"
        szSqlContext = " VALUES ('" & ToDBDate(m_dtRunDate) & "',"
        szSqlContext = szSqlContext & "'" & szBusID & "',"
        szSqlContext = szSqlContext & "'" & m_szBusID & "',"
        szSqlContext = szSqlContext & "'" & "1 " & " ',"
        'szSqlContext = szSqlContext & "'" & szLog & "  ',"
        szSqlContext = szSqlContext & "'" & "车次" & m_szBusID & "被拆入到" & szBusID & Now & "  '  )"
        szSql = szSql & szSqlfield & szSqlContext
        oDb.Execute szSql, lRow
        UpdateEnviromentSeatCountEx szBusID, m_dtRunDate, oDb
    oDb.CommitTrans
    #If SUPPORT_MSG = 1 Then
        ReDim aszValues(1 To 2)
        aszValues(1) = m_szBusID
        aszValues(2) = ToDBDate(m_dtRunDate)
        oParam.Init m_oActiveUser
        m_oSndMsg.SendMsg oParam.NowDateTime, cszMergeBus, aszValues
    #End If
    WriteOperateLog m_oActiveUser, RIGHT_ReBusSlitp, "车次" & m_szBusID & "]售出的[" & nCount & "]张票被拆分到" & szBusID
    'ReBusSlipLockEx szbusID, False
    MegerBusAndSlitpBus = szRutunData
    Exit Function
here:
    oDb.RollbackTrans
    Set oDb = Nothing
    err.Raise err.Number
End Function

'取消并班
Public Function UnMegerBusAndSlitpBus(szBusInfo() As String, szBusID As String)
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim nCount As Integer
    Dim lEffect As Long
    Dim szTicketsList As String
    Dim tSeatTypeInfo(1 To 5) As TSeatInfoCount

    #If SUPPORT_MSG = 1 Then
    Dim oParam As New SystemParam
    Dim aszValues() As String
    #End If
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_ReBusSlitp
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    On Error GoTo here
    

    oDb.BeginTrans
        '更新车次状态
        szSql = "UPDATE Env_bus_info set status='" & ST_BusStopped & "' WHERE bus_id ='" & m_szBusID & "' AND bus_date='" & ToDBDate(m_dtRunDate) & "'"
        oDb.Execute szSql

        For i = 1 To ArrayLength(szBusInfo)
             szTicketsList = szTicketsList & "'" & szBusInfo(i) & "',"
        Next i
        
        szTicketsList = Left(szTicketsList, Len(szTicketsList) - 1)
        
        szSql = " SELECT COUNT(seat_no) num FROM env_bus_seat_lst WHERE bus_id ='" & szBusID & "' AND  " _
                & " bus_date='" & ToDBDate(m_dtRunDate) & "' AND  status = '" & ST_SeatSlitp & "' AND ticket_no IN ( " & szTicketsList & ") "
        Set rsTemp = oDb.Execute(szSql)
        
        If rsTemp.RecordCount > 0 Then
            nCount = FormatDbValue(rsTemp!num)
        Else
            ShowError ERR_UnSlipCoutNotEqualSlipCount
        End If
           
        szSql = "UPDATE Env_bus_seat_lst set ticket_no='', status='" & ST_SeatCanSell & "' WHERE bus_id ='" & szBusID & "' AND  " _
                & " bus_date='" & ToDBDate(m_dtRunDate) & "' AND ticket_no IN ( " & szTicketsList & ") "
        oDb.Execute szSql, lEffect
        
        If lEffect <> nCount Then ShowError ERR_UnSlipCoutNotEqualSlipCount
        
        szSql = " DELETE Mergebus_split_lst WHERE bus_id ='" & szBusID & " ' AND merge_bus_id ='" & m_szBusID & "' AND bus_date= '" & ToDBDate(m_dtRunDate) & "'"
        oDb.Execute szSql

        UpdateEnviromentSeatCountEx szBusID, m_dtRunDate, oDb
    oDb.CommitTrans
    #If SUPPORT_MSG = 1 Then
        ReDim aszValues(1 To 2)
        aszValues(1) = m_szBusID
        aszValues(2) = ToDBDate(m_dtRunDate)
        oParam.Init m_oActiveUser
        m_oSndMsg.SendMsg oParam.NowDateTime, cszStopBus, aszValues
    #End If
    WriteOperateLog m_oActiveUser, RIGHT_REBusUnSlitp, "车次" & m_szBusID & "]拆分的[" & nCount & "]张票被取消"
    'ReBusSlipLockEx szbusID, False
'    UnMegerBusAndSlitpBus = szRutunData
    Exit Function
here:
    oDb.RollbackTrans
    Set oDb = Nothing
    err.Raise err.Number
End Function

'得于并班车次的信息 zyw 玉环
Public Function GetMegerBusInfo(szBusID As String, szBusDate As Date) As Recordset
On Error GoTo here

    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As New Recordset
       
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    
        
    szSql = " SELECT * FROM Mergebus_split_lst WHERE  merge_bus_id ='" & szBusID & "' AND bus_date= '" & ToDBDate(szBusDate) & "'"
    
    Set rsTemp = oDb.Execute(szSql)
    
    Set GetMegerBusInfo = rsTemp
       
    Exit Function
here:
    oDb.RollbackTrans
    Set oDb = Nothing
    err.Raise err.Number
End Function



Public Function GetSlitpBusTicketNo() As String()
    '得到拆分车次的票号
    '************注**************
    '************此处的SQL语句需要修改,写得很乱
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As New Recordset
    Dim szDate() As String
    Dim i As Integer
    Dim nCount As Integer
    
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    
    On Error GoTo here
    szSql = "SELECT e2.ticket_no,e2.bus_id,e1.seat_no ,e1.bus_id as mergebusid " _
    & " FROM Env_bus_seat_lst e1, Env_bus_seat_lst e2 " _
    & " WHERE e2.status='" & ST_SeatSlitp & "' AND e2.bus_id in " _
    & " (SELECT  bus_id FROM Mergebus_split_lst WHERE merge_bus_id='" & m_szBusID & "' " _
    & " AND bus_date='" & ToDBDate(m_dtRunDate) & "') " _
    & " AND e2.bus_date='" & ToDBDate(m_dtRunDate) & "'" _
    & " AND e1.bus_date=e2.bus_date AND e1.bus_id='" & m_szBusID & "'" _
    & " AND e1.ticket_no=e2.ticket_no"
    
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    ReDim szDate(1 To nCount, 1 To 3)
    For i = 1 To nCount
    szDate(i, 1) = FormatDbValue(rsTemp!mergebusid)
    szDate(i, 2) = FormatDbValue(rsTemp!ticket_no)
    szDate(i, 3) = FormatDbValue(rsTemp!seat_no)
    rsTemp.MoveNext
    Next
    End If
    Set oDb = Nothing
    Set rsTemp = Nothing
    GetSlitpBusTicketNo = szDate
    Exit Function
here:
    Set oDb = Nothing
    Set rsTemp = Nothing
    
End Function



'得到配载信息
Public Function GetAllotInfo(Optional SellStationID As String) As TReBusAllotInfo() '无定义
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rzTemp As Recordset
    Dim atResult() As TReBusAllotInfo
      
    AssertObjIsValid
    
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    If DateDiff("d", m_dtRunDate, Now) <= 0 Then
        szSql = "SELECT a.*,b.sell_station_name,b.sell_station_full_name,c.check_gate_name" & _
                " FROM Work_env_bus_allot_lst a,sell_station_info b,Checkgate_info c" & _
                " WHERE a.bus_id=" & TransFieldValueToString(m_szBusID) & " AND a.bus_date=" & TransFieldValueToString(m_dtRunDate) & _
                " AND a.check_gate_id=c.check_gate_id AND a.sell_station_id=b.sell_station_id"
    Else
        szSql = "SELECT a.*,b.sell_station_name,b.sell_station_full_name,c.check_gate_name" & _
                " FROM Env_bus_allot_lst a,sell_station_info b,Checkgate_info c" & _
                " WHERE a.bus_id=" & TransFieldValueToString(m_szBusID) & " AND a.bus_date=" & TransFieldValueToString(m_dtRunDate) & _
                " AND a.check_gate_id=c.check_gate_id AND a.sell_station_id=b.sell_station_id"
    End If
    
    If SellStationID <> "" Then
      szSql = szSql & " AND a.sell_station_id='" & SellStationID & "'"
    End If
    szSql = szSql & " ORDER BY a.bus_start_time"
    Set rzTemp = oDb.Execute(szSql)
    If rzTemp.RecordCount = 0 Then Exit Function
    ReDim atResult(1 To rzTemp.RecordCount)
    Dim i As Integer
    For i = 1 To rzTemp.RecordCount
        atResult(i).dtRunTime = FormatDbValue(rzTemp!bus_start_time)
        atResult(i).nCanSellQuantity = FormatDbValue(rzTemp!can_sale_quantity)
        atResult(i).szBusID = FormatDbValue(rzTemp!bus_id)
        atResult(i).szCheckGateID = FormatDbValue(rzTemp!check_gate_id)
        atResult(i).szCheckGateName = FormatDbValue(rzTemp!check_gate_name)
        atResult(i).dyBusDate = FormatDbValue(rzTemp!bus_date)
        atResult(i).szSellStationID = FormatDbValue(rzTemp!sell_station_id)
        atResult(i).szSellStationName = FormatDbValue(rzTemp!sell_station_name)
        atResult(i).nStatus = FormatDbValue(rzTemp!Status)
        atResult(i).nSellQuantity = FormatDbValue(rzTemp!sell_quantity)
        
        rzTemp.MoveNext
    Next i
    
    GetAllotInfo = atResult
    Set oDb = Nothing
End Function


'保存车次配载信息
Public Sub SaveAllot(patAllotInfo() As TReBusAllotInfo)
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_REBusAllot
    
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr
    Dim szSql As String
        
    Dim i As Integer, j As Integer, nCount As Integer
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

    Dim atTmp() As TReBusAllotInfo
    atTmp = GetAllotInfo
    
    '分别处理哪些进行删除，哪些进行修改，哪些需新增
    Dim atDelete() As TReBusAllotInfo, nCountDelete As Integer
    Dim atUpdate() As TReBusAllotInfo, nCountUpdate As Integer
    Dim atInsert() As TReBusAllotInfo, nCountInsert As Integer
    
    '利用最小的发车时间，更新环境车次中的检票口及发车时间
    Dim dtMinRunTime As Date, nMinIndex As Integer
    If nCount > 0 Then
        nMinIndex = 1
        dtMinRunTime = patAllotInfo(1).dtRunTime
    End If
    
    If nCount = 0 Then atDelete = atTmp
    '得到需要更改的和新增的
    For i = 1 To nCount
        For j = 1 To ArrayLength(atTmp)
            If atTmp(j).szSellStationID = patAllotInfo(i).szSellStationID Then
                Exit For
            End If
        Next j
        If j > ArrayLength(atTmp) Then  '新增的
            nCountInsert = nCountInsert + 1
            ReDim Preserve atInsert(1 To nCountInsert)
            atInsert(nCountInsert) = patAllotInfo(i)
        Else    '更改的
            If patAllotInfo(i).nCanSellQuantity < atTmp(j).nSellQuantity And patAllotInfo(i).nCanSellQuantity <> -1 Then RaiseError ERR_SellTicketsOut
            If patAllotInfo(i).dtRunTime <> atTmp(j).dtRunTime Or patAllotInfo(i).nCanSellQuantity <> atTmp(j).nCanSellQuantity Or patAllotInfo(i).szCheckGateID <> atTmp(j).szCheckGateID Then    '有更改
                nCountUpdate = nCountUpdate + 1
                ReDim Preserve atUpdate(1 To nCountUpdate)
                atUpdate(nCountUpdate) = patAllotInfo(i)
                '以下将两个无用的属性用于临时存放是否更改了发车时间和是否更改了检票口的标志，用于后面的消息发送
                atUpdate(nCountUpdate).szSellStationName = ""
                atUpdate(nCountUpdate).szCheckGateName = ""
                If patAllotInfo(i).dtRunTime <> atTmp(j).dtRunTime Then
                    atUpdate(nCountUpdate).szSellStationName = "1"
                End If
                If patAllotInfo(i).szCheckGateID <> atTmp(j).szCheckGateID Then
                    atUpdate(nCountUpdate).szCheckGateName = "1"
                End If
            End If
        End If
        
        If patAllotInfo(i).dtRunTime < dtMinRunTime Then
            nMinIndex = i
            dtMinRunTime = patAllotInfo(i).dtRunTime
        End If
    Next i
    '得到需要删除的
    For i = 1 To ArrayLength(atTmp)
        For j = 1 To nCount
            If atTmp(i).szSellStationID = patAllotInfo(j).szSellStationID Then Exit For
        Next j
        If j > nCount Then  '未找到
            nCountDelete = nCountDelete + 1
            ReDim Preserve atDelete(1 To nCountDelete)
            atDelete(nCountDelete) = atTmp(i)
        End If
    Next i
    
On Error GoTo ErrHandle
    
    oDb.BeginTrans
    
    Dim szDeleteLog As String
    Dim szUpdateLog As String
    Dim szInsertLog As String
    
    '处理删除
    For i = 1 To ArrayLength(atDelete)
        If atDelete(i).nSellQuantity > 0 Then RaiseError ERR_REBusHaveSellTicketNotDelete
        szSql = "DELETE Env_bus_allot_lst WHERE bus_id=" & TransFieldValueToString(m_szBusID) & " AND bus_date=" & TransFieldValueToString(m_dtRunDate) & _
                " AND sell_station_id=" & TransFieldValueToString(atDelete(i).szSellStationID)
        oDb.Execute szSql
        szDeleteLog = szDeleteLog & "删除车次配载，车次代码为[" & m_szBusID & "],发车日期为[" & m_dtRunDate & "],上车站为[" & atDelete(i).szSellStationID & "] "
    Next i
    
    '处理更新
    For i = 1 To ArrayLength(atUpdate)
        szSql = "UPDATE Env_bus_allot_lst SET can_sale_quantity=" & TransFieldValueToString(atUpdate(i).nCanSellQuantity) & _
                                                    ",bus_start_time=" & TransFieldValueToString(atUpdate(i).dtRunTime) & ",check_gate_id=" & TransFieldValueToString(atUpdate(i).szCheckGateID) & _
                " WHERE bus_id=" & TransFieldValueToString(m_szBusID) & " AND bus_date=" & TransFieldValueToString(m_dtRunDate) & _
                " AND sell_station_id=" & TransFieldValueToString(atUpdate(i).szSellStationID)
        oDb.Execute szSql
        szUpdateLog = szUpdateLog & "修改车次配载，车次[" & m_szBusID & "],发车日期[" & m_dtRunDate & "],上车站为[" & atUpdate(i).szSellStationID & "]，修改为：发车时间[" & atUpdate(i).dtRunTime & "],检票口[" & atUpdate(i).szCheckGateID & "] "
    Next i
                                                    '",check_gate_id=" & TransFieldValueToString(atUpdate(i).szCheckGateID) & _

    '处理插入,缺省正常状态和未售票
    For i = 1 To ArrayLength(atInsert)
        szSql = "INSERT Env_bus_allot_lst (bus_date,bus_id,sell_station_id,check_gate_id,bus_start_time,status,can_sale_quantity,sell_quantity)" & _
                " VALUES(" & TransFieldValueToString(m_dtRunDate) & "," & TransFieldValueToString(m_szBusID) & "," & TransFieldValueToString(atInsert(i).szSellStationID) & _
                "," & TransFieldValueToString(atInsert(i).szCheckGateID) & "," & TransFieldValueToString(atInsert(i).dtRunTime) & ",0" & _
                "," & TransFieldValueToString(atInsert(i).nCanSellQuantity) & ",0)"
        oDb.Execute szSql
        szInsertLog = szInsertLog & "新增车次配载，车次代码为[" & m_szBusID & "],发车日期为[" & m_dtRunDate & "],上车站为[" & atInsert(i).szSellStationID & "],检票口为[" & atInsert(i).szCheckGateID & "],发车时间为[" & atInsert(i).dtRunTime & "可售张数为[" & atInsert(i).nCanSellQuantity & "] "
    Next i
    
    
        '如果只修改配载站的检票口,不修改车次信息的检票口。
        '这样的话,如果修改的检票口正好当前车次信息表的检票口的话，会导致车次关联不到。（新的配载站权限中用到关联）
        '所以，需要加上以下的语句，将车次信息表中的相应信息也进行修改。
        
        '将车次表中的发车时间与配载的发车时间进行对应
'        szSql = " UPDATE bus_info " _
'            & " SET bus_start_time = a.bus_start_time , check_gate_id = a.check_gate_id  " _
'            & " FROM bus_info b , checkgate_info c , bus_allot_info a " _
'            & " WHERE  b.check_gate_id = c.check_gate_id " _
'            & " AND b.bus_id = a.bus_id " _
'            & " AND a.sell_station_id  = c.sell_station_id " _
'            & " AND b.bus_id = " & TransFieldValueToString(m_szBusID)
'
'        oDb.Execute szSql
'

        szSql = " UPDATE env_bus_info " _
            & " SET bus_start_time = a.bus_start_time , check_gate_id = a.check_gate_id  " _
            & " FROM env_bus_info b , env_bus_allot_lst a  , checkgate_info c " _
            & " WHERE b.bus_id = a.bus_id AND b.bus_date = a.bus_date " _
            & " AND a.sell_station_id  = c.sell_station_id " _
            & " AND b.check_gate_id = c.check_gate_id " _
            & " AND b.bus_id = " & TransFieldValueToString(m_szBusID) _
            & " AND b.bus_date = " & TransFieldValueToString(m_dtRunDate)
            
        oDb.Execute szSql
        
        
    oDb.CommitTrans
    
    WriteOperateLog m_oActiveUser, RIGHT_REBusAllot, szDeleteLog & szUpdateLog & szInsertLog
    
    Set oDb = Nothing
    
    '消息处理
     #If SUPPORT_MSG = 1 Then
     Dim aszValues() As String
     For i = 1 To ArrayLength(atInsert)     '新增的车次
        ReDim aszValues(1 To 3)
        m_oSndMsg.SellStation = atInsert(i).szSellStationID
        aszValues(1) = atInsert(i).szBusID
        aszValues(2) = ToDBDate(atInsert(i).dyBusDate)
        aszValues(3) = atInsert(i).szCheckGateID
        m_oSndMsg.SendMsg Now, cszAddBus, aszValues
     Next
     For i = 1 To ArrayLength(atUpdate)     '更改的车次
        If atUpdate(i).szSellStationName <> "" Or atUpdate(i).szCheckGateName <> "" Then '更改了发车时间或检票口
            ReDim aszValues(1 To 3)
            m_oSndMsg.SellStation = atUpdate(i).szSellStationID
            aszValues(1) = atUpdate(i).szBusID
            aszValues(2) = ToDBDate(atUpdate(i).dyBusDate)
            If atUpdate(i).szSellStationName <> "" Then '更改了发车时间
                aszValues(3) = ToDBDateTime(atUpdate(i).dtRunTime)
                m_oSndMsg.SendMsg Now, cszChangeBusTime, aszValues
            End If
            If atUpdate(i).szCheckGateName <> "" Then  '更改了检票口时间
                aszValues(3) = atUpdate(i).szCheckGateID
                m_oSndMsg.SendMsg Now, cszChangeBusCheckGate, aszValues
            End If
        End If
    Next i
    For i = 1 To ArrayLength(atDelete)      '删除的车次
        ReDim aszValues(1 To 2)
        m_oSndMsg.SellStation = atDelete(i).szSellStationID
        aszValues(1) = atDelete(i).szBusID
        aszValues(2) = ToDBDate(atDelete(i).dyBusDate)
        m_oSndMsg.SendMsg Now, cszRemoveBus, aszValues
    Next i
    m_oSndMsg.SellStation = ""
     
     #End If
    
    Exit Sub
ErrHandle:
    Set oDb = Nothing
    oDb.RollbackTrans
    RaiseError err.Number, , err.Description
End Sub




'=======================================================================
'=车次售票点管理                                                       =
'=得到配载信息                                                         =
'=======================================================================

Public Function GetSellStationInfo(Optional SellStationID As String) As TReBusAllotInfo()
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rzTemp As Recordset
    Dim atResult() As TReBusAllotInfo
    
    AssertObjIsValid
    
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    If DateDiff("d", m_dtRunDate, Now) <= 0 Then
        szSql = "SELECT a.*,b.sell_station_name,b.sell_station_full_name" & _
                " FROM work_env_bus_sell_station_lst a,sell_station_info b" & _
                " WHERE a.bus_id=" & TransFieldValueToString(m_szBusID) & " AND a.bus_date=" & TransFieldValueToString(m_dtRunDate) & _
                " AND a.sell_station_id=b.sell_station_id"
    Else
        szSql = "SELECT a.*,b.sell_station_name,b.sell_station_full_name" & _
                " FROM env_bus_sell_station_lst a,sell_station_info b" & _
                " WHERE a.bus_id=" & TransFieldValueToString(m_szBusID) & " AND a.bus_date=" & TransFieldValueToString(m_dtRunDate) & _
                " AND a.sell_station_id=b.sell_station_id"
    End If
    
    If SellStationID <> "" Then
      szSql = szSql & " AND a.sell_station_id='" & SellStationID & "'"
    End If
    szSql = szSql & " ORDER BY a.sell_station_id"
    Set rzTemp = oDb.Execute(szSql)
    If rzTemp.RecordCount = 0 Then Exit Function
    ReDim atResult(1 To rzTemp.RecordCount)
    Dim i As Integer
    For i = 1 To rzTemp.RecordCount
        atResult(i).szBusID = FormatDbValue(rzTemp!bus_id)
        atResult(i).dyBusDate = FormatDbValue(rzTemp!bus_date)
        atResult(i).szSellStationID = FormatDbValue(rzTemp!sell_station_id)
        atResult(i).szSellStationName = FormatDbValue(rzTemp!sell_station_name)
        atResult(i).nCanSellQuantity = FormatDbValue(rzTemp!can_sale_quantity)
        atResult(i).nSellQuantity = FormatDbValue(rzTemp!sell_quantity)
        
        rzTemp.MoveNext
    Next i
    
    GetSellStationInfo = atResult
    Set oDb = Nothing
End Function



'=======================================================================
'=车次售票点管理                                                       =
'=保存车次配载信息                                                     =
'=======================================================================

Public Sub SaveSellStationInfo(patAllotInfo() As TReBusAllotInfo)


On Error GoTo ErrHandle
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_REBusAllot
    
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim i As Integer, j As Integer, nCount As Integer
    Dim nCanSell As Integer
    Dim atTmp() As TReBusAllotInfo
    '分别处理哪些进行删除，哪些进行修改，哪些需新增
    Dim atDelete() As TReBusAllotInfo, nCountDelete As Integer
    Dim atUpdate() As TReBusAllotInfo, nCountUpdate As Integer
    Dim atInsert() As TReBusAllotInfo, nCountInsert As Integer
    
    oDb.ConnectionString = GetConnectionStr
    
    nCount = ArrayLength(patAllotInfo)
    atTmp = GetSellStationInfo
    If nCount = 0 Then atDelete = atTmp
    '得到需要更改的和新增的
    
    For i = 1 To nCount
        For j = 1 To ArrayLength(atTmp)
            If atTmp(j).szSellStationID = patAllotInfo(i).szSellStationID Then
                Exit For
            End If
        Next j
        If j > ArrayLength(atTmp) Then  '新增的
            nCountInsert = nCountInsert + 1
            ReDim Preserve atInsert(1 To nCountInsert)
            atInsert(nCountInsert) = patAllotInfo(i)
        Else    '更改的
            If patAllotInfo(i).nCanSellQuantity < atTmp(j).nSellQuantity And patAllotInfo(i).nCanSellQuantity <> -1 Then RaiseError ERR_SellTicketsOut '出有票售出的错
            If patAllotInfo(i).nCanSellQuantity <> atTmp(j).nCanSellQuantity Then     '有更改
                nCountUpdate = nCountUpdate + 1
                ReDim Preserve atUpdate(1 To nCountUpdate)
                atUpdate(nCountUpdate) = patAllotInfo(i)
            End If
        End If
        
    Next i
    '得到需要删除的
    For i = 1 To ArrayLength(atTmp)
        For j = 1 To nCount
            If atTmp(i).szSellStationID = patAllotInfo(j).szSellStationID Then Exit For
        Next j
        If j > nCount Then  '未找到
            nCountDelete = nCountDelete + 1
            ReDim Preserve atDelete(1 To nCountDelete)
            atDelete(nCountDelete) = atTmp(i)
        End If
    Next i
    
    
    oDb.BeginTrans
        '处理删除
        For i = 1 To ArrayLength(atDelete)
            If atDelete(i).nSellQuantity > 0 Then RaiseError ERR_REBusHaveSellTicketNotDelete
            szSql = "DELETE env_bus_sell_station_lst WHERE bus_id=" & TransFieldValueToString(m_szBusID) & " AND bus_date=" & TransFieldValueToString(m_dtRunDate) & _
                    " AND sell_station_id=" & TransFieldValueToString(atDelete(i).szSellStationID)
            oDb.Execute szSql
        Next i
        
        '处理更新
        For i = 1 To ArrayLength(atUpdate)
            szSql = "UPDATE env_bus_sell_station_lst SET can_sale_quantity=" & TransFieldValueToString(atUpdate(i).nCanSellQuantity) & _
                    " WHERE bus_id=" & TransFieldValueToString(m_szBusID) & " AND bus_date=" & TransFieldValueToString(m_dtRunDate) & _
                    " AND sell_station_id=" & TransFieldValueToString(atUpdate(i).szSellStationID)
            oDb.Execute szSql
        Next i
        
        '处理插入,缺省正常状态和未售票
        For i = 1 To ArrayLength(atInsert)
            szSql = "INSERT env_bus_sell_station_lst (bus_date,bus_id,sell_station_id,can_sale_quantity,sell_quantity)" _
                    & " VALUES(" & TransFieldValueToString(m_dtRunDate) & "," & TransFieldValueToString(m_szBusID) & "," _
                    & TransFieldValueToString(atInsert(i).szSellStationID) & "," _
                    & TransFieldValueToString(atInsert(i).nCanSellQuantity) & ",0)"
            oDb.Execute szSql
        Next i
        
    oDb.CommitTrans
    
    Set oDb = Nothing
    
    
    Exit Sub
ErrHandle:
    Set oDb = Nothing
    oDb.RollbackTrans
    RaiseError err.Number, , err.Description
End Sub




