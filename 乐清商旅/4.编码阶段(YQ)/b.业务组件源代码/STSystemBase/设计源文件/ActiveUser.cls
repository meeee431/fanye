VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ActiveUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************
'* Source File Name:ActiveUser.cls
'* Project Name:STSystem.vbp
'* Engineer:陈峰
'* Data Generated:2002-08-15
'* Last Revision Date:2003-06-18
'* Brief Description:活动用户类
'* Relational Document:
'**********************************************************


'***开头的注意一下


Public Type TInterfaceCheckTicketEx
    BusID As String
    BusDate As Date
    TicketStatus As ETicketStatus
    TicketType As ETicketType
    SeatNo As String
    StationName As String
    Price As Double
    TicketID As String
    BusStartTime As Date
    ChildrenTicket As Boolean
    SellstationID As String
    SellStationName As String
    
    CardType As String      '证件类型
    IDCardNo As String      '证件号
    PersonName As String    '姓名
    Sex As String           '性别
    National As String      '民族(本国)'国籍(国外)
    Address As String       '地址
    PersonPicture As String '证件照片
End Type

Public Type TSeatInfoEx2
    szSellStationID As String
    szSellStationName As String
    szSeatNo As String
    szSeatStatus As ERESeatStatus
    szSeatType As String
    szSeatTypeName As String
    szTicketNo As String
    szDestStationID As String
    szDestName As String
    szTicketPrice As Currency
    szRemark As String '备注
    szTicketType As String
    szTicketTypeName As String
    
    szCardType As String      '证件类型
    szIDCardNo As String      '证件号
    szPersonName As String    '姓名
    szSex As String           '性别
    szNational As String      '民族(本国)'国籍(国外)
    szAddress As String       '地址
    szPersonPicture As String '证件照片
End Type

'实名制信息
Public Type TCardInfo
    szCardType As String        '证件类型
    szIDCardNo As String        '证件号码
    szPersonName As String      '姓名
    szSex As String             '性别
    szPersonPicture As String   '照片
    szNation As String          '民族(本国)/国籍(外国)
    szAddress As String         '地址
    szReserve1 As String        '预留备用字段1
    szReserve2 As String        '预留备用字段2
End Type


Option Explicit
'中间层各类的类号
Public Enum ECOMID
    CID_ActiveUser = 1
    CID_SystemAgent = 2
    CID_RTConnection = 3
    CID_ServerProxy = 4
    CID_SellTicketClient = 5
    CID_SellTicketService = 6
    CID_Company = 7
    CID_Vehicle = 8
    CID_Owner = 9
    CID_BusProject = 10
    CID_Bus = 11
    CID_BaseInfo = 12
    CID_Area = 13
    CID_ChargeRatio = 14
    CID_Station = 15
    CID_Section = 16
    CID_Route = 17
    CID_RoadLevel = 18
    CID_CheckGate = 19
    CID_VehicleModel = 20
    CID_RegularScheme = 21
    CID_REBus = 22
    CID_REScheme = 23
    CID_ServiceTicket = 24
    CID_ClientTicket = 25
    CID_ClientReturnedTicket = 26
    CID_CheckTicket = 27
    CID_Finance = 28
    CID_SystemMan = 29
    
    CID_SystemParam = 30
    CID_Unit = 31
    CID_COMFunction = 32
    CID_Component = 33
    CID_User = 34
    CID_UserGroup = 35
    CID_RouteTicketPrice = 36
    CID_TicketPriceFormula = 37
    CID_TicketItem = 38
    CID_HalfTicketPrice = 39
    CID_TicketPriceMan = 40
    CID_PriceItemFunLib = 41
    
    CID_CheckSheet = 42
    CID_Protocol = 43
    CID_SettleSheet = 44
    CID_Split = 45
    
    CID_TicketSellerDim = 46
    CID_TicketStationDim = 47
    CID_SellerFinance = 48
    
    CID_TicketType = 49
    '以下部分是行包用的
    CID_Kinds = 50
    CID_LugInfo = 51
    CID_LugReceive = 52
    CID_LuggageItem = 53
    CID_LuggageSheet = 54
    
    CID_CheckIn = 55 '报班
    CID_Book = 56 '预定
    
    CID_SeatType = 57 '座位类型
    CID_BusType = 58 '车次类型
    '语音播放
    CID_Chunnul = 59
    CID_Room = 60
    CID_File = 61
    CID_Rule = 62
    CID_Task = 63
    CID_Project = 64
    CID_SoundBase = 65
    CID_SellStation = 66
    End Enum

'各类的错误的起始号(实际也是其所有资源的起始号)
Public Enum EClassErrBegin
    ERR_ActiveUser = CID_ActiveUser * cnMidErrStep + cnMidErrBegin
    ERR_SystemAgent = CID_SystemAgent * cnMidErrStep + cnMidErrBegin
    ERR_RTConnection = CID_RTConnection * cnMidErrStep + cnMidErrBegin
    ERR_ServerProxy = CID_ServerProxy * cnMidErrStep + cnMidErrBegin
    ERR_SellTicketClient = CID_SellTicketClient * cnMidErrStep + cnMidErrBegin
    ERR_SellTicketService = CID_SellTicketService * cnMidErrStep + cnMidErrBegin
    ERR_Company = CID_Company * cnMidErrStep + cnMidErrBegin
    ERR_Vehicle = CID_Vehicle * cnMidErrStep + cnMidErrBegin
    ERR_Owner = CID_Owner * cnMidErrStep + cnMidErrBegin
    ERR_BusProject = CID_BusProject * cnMidErrStep + cnMidErrBegin
    ERR_Bus = CID_Bus * cnMidErrStep + cnMidErrBegin
    ERR_BaseInfo = CID_BaseInfo * cnMidErrStep + cnMidErrBegin
    ERR_Area = CID_Area * cnMidErrStep + cnMidErrBegin
    ERR_ChargeRatio = CID_ChargeRatio * cnMidErrStep + cnMidErrBegin
    ERR_Station = CID_Station * cnMidErrStep + cnMidErrBegin
    ERR_Section = CID_Section * cnMidErrStep + cnMidErrBegin
    ERR_Route = CID_Route * cnMidErrStep + cnMidErrBegin
    ERR_RoadLevel = CID_RoadLevel * cnMidErrStep + cnMidErrBegin
    ERR_CheckGate = CID_CheckGate * cnMidErrStep + cnMidErrBegin
    ERR_VehicleModel = CID_VehicleModel * cnMidErrStep + cnMidErrBegin
    ERR_RegularScheme = CID_RegularScheme * cnMidErrStep + cnMidErrBegin
    ERR_REBus = CID_REBus * cnMidErrStep + cnMidErrBegin
    ERR_REScheme = CID_REScheme * cnMidErrStep + cnMidErrBegin
    ERR_ServiceTicket = CID_ServiceTicket * cnMidErrStep + cnMidErrBegin
    ERR_ClientTicket = CID_ClientTicket * cnMidErrStep + cnMidErrBegin
    ERR_ClientReturnedTicket = CID_ClientReturnedTicket * cnMidErrStep + cnMidErrBegin
    ERR_CheckTicket = CID_CheckTicket * cnMidErrStep + cnMidErrBegin
    ERR_Finance = CID_Finance * cnMidErrStep + cnMidErrBegin
    ERR_SystemMan = CID_SystemMan * cnMidErrStep + cnMidErrBegin
    ERR_SystemParam = CID_SystemParam * cnMidErrStep + cnMidErrBegin
    ERR_Unit = CID_Unit * cnMidErrStep + cnMidErrBegin
    ERR_COMFunction = CID_COMFunction * cnMidErrStep + cnMidErrBegin
    ERR_Component = CID_Component * cnMidErrStep + cnMidErrBegin
    ERR_User = CID_User * cnMidErrStep + cnMidErrBegin
    ERR_UserGroup = CID_UserGroup * cnMidErrStep + cnMidErrBegin
    ERR_RouteTicketPrice = CID_RouteTicketPrice * cnMidErrStep + cnMidErrBegin
    ERR_TicketPriceFormula = CID_TicketPriceFormula * cnMidErrStep + cnMidErrBegin
    ERR_TicketItem = CID_TicketItem * cnMidErrStep + cnMidErrBegin
    ERR_HalfTicketPrice = CID_HalfTicketPrice * cnMidErrStep + cnMidErrBegin
    ERR_TicketPriceMan = CID_TicketPriceMan * cnMidErrStep + cnMidErrBegin
    ERR_PriceItemFunLib = CID_PriceItemFunLib * cnMidErrStep + cnMidErrBegin
    
    ERR_CheckSheet = CID_CheckSheet * cnMidErrStep + cnMidErrBegin
    ERR_Protocol = CID_Protocol * cnMidErrStep + cnMidErrBegin
    ERR_SettleSheet = CID_SettleSheet * cnMidErrStep + cnMidErrBegin
    ERR_Split = CID_Split * cnMidErrStep + cnMidErrBegin
    
    ERR_TicketSellerDim = CID_TicketSellerDim * cnMidErrStep + cnMidErrBegin
    ERR_TicketStationDim = CID_TicketStationDim * cnMidErrStep + cnMidErrBegin
    ERR_SellerFinance = CID_SellerFinance * cnMidErrStep + cnMidErrBegin
    
    ERR_TicketType = CID_TicketType * cnMidErrStep + cnMidErrBegin
    
    '以下部分是行包用的
    ERR_Kinds = CID_Kinds * cnMidErrStep + cnMidErrBegin
    ERR_LugInfo = CID_LugInfo * cnMidErrStep + cnMidErrBegin
    ERR_LugReceive = CID_LugReceive * cnMidErrStep + cnMidErrBegin
    ERR_LuggageItem = CID_LuggageItem * cnMidErrStep + cnMidErrBegin
    ERR_LuggageSheet = CID_LuggageSheet * cnMidErrStep + cnMidErrBegin
    
    ERR_CheckIn = CID_CheckIn * cnMidErrStep + cnMidErrBegin '报班
    ERR_Book = CID_Book * cnMidErrStep + cnMidErrBegin
    
    '
    ERR_SeatType = CID_SeatType * cnMidErrStep + cnMidErrBegin
    ERR_BusType = CID_BusType * cnMidErrStep + cnMidErrBegin
    '语音播放
    ERR_Chunnul = CID_Chunnul * cnMidErrStep + cnMidErrBegin
    ERR_Room = CID_Room * cnMidErrStep + cnMidErrBegin
    ERR_File = CID_File * cnMidErrStep + cnMidErrBegin
    ERR_Rule = CID_Rule * cnMidErrStep + cnMidErrBegin
    ERR_Task = CID_Task * cnMidErrStep + cnMidErrBegin
    ERR_Project = CID_Project * cnMidErrStep + cnMidErrBegin
    ERR_SoundBase = CID_SoundBase * cnMidErrStep + cnMidErrBegin
    ERR_SellStation = CID_SellStation * cnMidErrStep + cnMidErrBegin
End Enum
'以下是公用的枚举类型

'预定
Public Enum EBookInfo
   SE_SeatBooked = 0 '已被预定,
   SE_CancelBook = 1 '已取消预定
   SE_SellSeatBooked = 2 ',2已被预人购买
End Enum
Public Enum ECityProvince
   EA_nInCity = 0  '市内
   EA_nOutCity = 1      '市外
   EA_nOutProvince = 2  '省外
End Enum



'售出票票种
Public Enum ESellTicketType
    TP_NormalSold = 1     '正常售出
    TP_ExtraSold = 2      '补票售出
    TP_ChangeSold = 3     '改签售出
End Enum
'车次类型
Public Enum EBusType
    TP_RegularBus = 0 '固定车次
    TP_ScrollBus = 1 ' 流水车次（滚动车次）
End Enum

'车辆状态
Public Enum EVehicleStatus
    ST_VehicleRun = 0 '未停车（正常运行）
    ST_VehicleStop = 1 ' 已停车
End Enum
'线路状态
Public Enum ERouteStatus
    ST_RouteAvailable = 0 '线路可用
    ST_RouteNotAvailable = 1 ' 线路不可用
End Enum
'线路上的站点类型
Public Enum ERouteStationType
    TP_CanSellTicket = 1 '此车站可售票
    TP_CanNotSellTicket = 2 ' 此车站不可售票，只用于连接线路
End Enum
'环境座位状态
Public Enum ERESeatStatus
    ST_SeatCanSell = 0 '此座位可售
    ST_SeatReserved = 1 ' 此座位已被预留
    ST_SeatBooked = 2 '此座位已被预定
    ST_SeatSold = 3 '此座位已被售出
    ST_SeatSlitp = 4 '此座位已被售出,拆分得到
    ST_SeatReplace = 5 '此座位已被售出，顶班得到
    ST_SeatMerge = 6 '此座位已被售出，并班得到
    ST_SeatProjectBooked = 64 '计划预定
End Enum
'环境车次状态
Public Enum EREBusStatus
    ST_BusNormal = 0 '车次正常
    ST_BusStopped = 1 ' 车次已被停班
    ST_BusMergeStopped = 2 '车次已被并班停班
    ST_BusStopCheck = 3 '车次已经停检
    ST_BusChecking = 4 '车次正在检票
    ST_BusExtraChecking = 5 '车次正在补检
    ST_BusSlitpStop = 8 '车次拆分停班
    ST_BusReplace = 16 '车次顶班
    'ST_BusLock = 32+ x   other +32
End Enum




'车票状态
Public Enum ETicketStatus
    ST_TicketNormal = 1 '车票正常售出
    ST_TicketSellChange = 2 ' 车票改签售出
    ST_TicketCanceled = 4 '车票已废
    ST_TicketChanged = 8 '被改签
    ST_TicketReturned = 16 '车票已退
    ST_TicketChecked = 32 '车票已检
End Enum
'车票种类
Public Enum ETicketType
    TP_FullPrice = 1 '全票
    TP_HalfPrice = 2 ' 半票
    TP_FreeTicket = 3 '免票
    TP_PreferentialTicket1 = 4 '优惠票1
    TP_PreferentialTicket2 = 5 '优惠票2
    TP_PreferentialTicket3 = 6 '优惠票3
End Enum

'退票方式
Public Enum EReturnTicketType
    TP_NormalReturn = 0 '正常退票
    TP_ForceReturn = 1 ' 强行退票
End Enum

'废票方式
Public Enum ECancelTicketType
    TP_NormalCancel = 0 '正常废票
    TP_ForceCancel = 1 ' 强行废票
End Enum

'环境车次报班状态
Public Enum EREBusBeingStatus
    ST_BusNotBeing = 0 '车次报班
    ST_BusNormalBeing = 1 ' 车次正常报班
    ST_BusLate = 2 '车次误班
    ST_BusMiss = 3 '车次错班
End Enum

'车次计划状态
Public Enum EBusProjectStatus
    ST_BusProjectNormal = 0 '车次计划未被逻辑删除
    ST_BusProjectDeleted = 1 ' 车次计划被逻辑删除
End Enum

'活动用户的状态
Public Enum EActiveUserStatus
    ST_AUNotLogin = 0
    ST_AUTimeOut = 1
    ST_AUValid = 2
End Enum
'用户组状态
Public Enum EUserGroupType
    TP_InnerUserGroup = 1 '此用户组是内置用户组
    TP_NormalUserGroup = 0 '此用户组不是内置用户组
End Enum
'单位类型
Public Enum EUnitType
    TP_UnitServer = 0 '此单位提供车票出售
    TP_UnitClient = 1 '此单位是代售点
    TP_UnitSC = 2 '此单位既提供车票出售也是代售点
End Enum

'线路票价表的执行状态
Public Enum ERoutePriceTableStatus
    ST_PriceTableExecute = 0 '此线路票价表在执行
    ST_PriceTableNotExecute = 1 '此线路票价表未执行
End Enum

'检票口状态
Public Enum ECheckGateStatus
    ST_CheckGateWait = 1
    ST_CheckGateChecking = 2
End Enum

'环境车次检票状态
Public Enum EREBusCheckStatus
    ST_BusCheckWait = 1 '还未开检（等待检票）
    ST_BusChecking2 = 2 '正在检票
    ST_BusCheckExChecking = 3 '正在补检
    ST_BusCheckStopped = 4 '已经停检
End Enum

'用户类型
Public Enum EUserType
    TP_InnerUser = 0 '内置用户
    TP_NormalUser = 1 '非内置用户
End Enum

'退出登录类型
Public Enum ELogoutType
    TP_NormalLogout = 0
    TP_ForceLogout = 1
    TP_TimeOutLogout = 3
End Enum

'权限描述
Public Enum ERightDesp
    RD_RightID = 1
    RD_RightGroup = 2
    RD_RightName = 3
    RD_WriteLog = 4
    RD_Annotation = 5
End Enum
'工作日期
Public Enum EUserWorkWeek
    WD_Sunday = 1 '星期日
    WD_Monday = 2 '星期一
    WD_Tuesday = 4 '星期二
    WD_Wednesday = 8 '星期三
    WD_Thursday = 16 '星期四
    WD_Friday = 32 '星期五
    WD_Saturday = 64 '星期六
    WD_AllDay = 127 '所有日期均可
    WD_AllWorkDay = 62 '正常工作日（除去星期日和星期六）
End Enum


Public Enum ETicketTypeValid
    TP_TicketTypeValid = 1 '可用
    TP_TicketTypeNotValid = 0 '不可用
    TP_TicketTypeAll = 2 '所有票种
End Enum

Public Enum EPriceItemUse
    TP_PriceItemUse = 1 '使用
    TP_PriceItemNotUse = 0 '不使用
    TP_PriceItemAll = 2 '所有票价项
End Enum

Public Enum ECheckLuggageIDValidMark
    TP_CheckLuggageIDValid = 1
    TP_CheckLuggageIDInValid = 0
End Enum

Public Enum EPlayMode
    TP_ContinuePlay = 1 '连续播放
    TP_OrderPlay = 0    '顺序播放
End Enum



'实体类的状态
Public Enum EObjectStatus
    ST_NotAvailable = 0
    ST_AddObj = 1
    ST_NormalObj = 2
    ST_EditObj = 3
End Enum

'**************************
'修改者:陈峰
'修改日期:2002-11-06
'是否写日志
'**************************
Public Enum EIsWritelog
    NotWirteLog = 0
    WriteLog = 1
End Enum

'改签类型

'公用的枚举类型到此为止


'以下是公用的数据类型
Public Type TBusOrderCount
    szStatioinID As String
    dbCount As Double
End Type

'座位信息
Public Type TSeatInfoCount
    SeatTypeID As String
    SeatTypeName As String
    SeatCount As Integer
    bFlg As Boolean
End Type


'退票费信息
Public Type TReturnTicketCharge
    sgTicketPrice As Double
    sgReturnRatio As Double
End Type

'通行费信息
Public Type TTransitChargeInfo
    szSection As String
    szSeatType As String
    szSeatTypeName As String
    szVehicleType As String
    szVehicleTypeName As String
    szPassCharge As Double
    szAnnotation As String
End Type
'车辆座位信息
Public Type TVehcileSeatType
    szSerial As String
    szVehcileID As String
    szVehcileTypeName As String
    szSeatTypeID As String
    szSeatTypeName As String
    szStartSeatNo As String
    szEndSeatNo As String
End Type
'**************************
'修改者:陈峰
'修改日期:2002-11-10
'车次票价主键信息(在删除票价时用到)
'**************************
Public Type TBusVehicleSeatType 'TBusVehicleSeatTypeEx
    szBusID As String
    szVehicleID As String
    szSeatTypeID As String
    szSeatTypeName As String
    szVehicleTypeCode As String
    szVehicleTypeShortName As String
    szVehicleTypeName As String
    
    szSellStationID As String
    szSellStationName As String
    
End Type
''环境车次站点信息
Public Type TREBusStationInfo
    szStationID  As String
    szStationName As String
    szStationSerialNo As String
    nMileage As Double
    sgLimitedSellTime As Single
    nLimitedSellCount As Integer
    sgFullPrice As Double
    sgHalfPrice As Double
    sgPreferentialPrice1 As Double '优惠票1
    sgPreferentialPrice2 As Double '优惠票2
    sgPreferentialPrice3 As Double '优惠票3
    sgBasePrice As Double
    szSeatTypeID As String         '座位类型代码
    szSeatTypeName As String       '座位类型名称
    
    
    szSellStationID As String
    szSellStationName As String


End Type

'**************************
'修改者:陈峰
'修改日期:2002-11-06
'环境车次站点信息
'**************************
Public Type TBusStationSellInfo
    szStationID  As String
    szStationName As String
    nMileage As Double
    sgLimitedSellTime As Single
    nLimitedSellCount As Integer
    sgFullPrice As Double
    sgHalfPrice As Double
    sgPreferentialPrice1 As Double '优惠票1
    sgPreferentialPrice2 As Double '优惠票2
    sgPreferentialPrice3 As Double '优惠票3
    sgBasePrice As Double
    szSeatTypeID As String         '座位类型代码
    szSeatTypeName As String       '座位类型名称
    
    
    szSellStationID As String
    szSellStationName As String
    
End Type

'车次车辆信息
Public Type TBusVehicleInfo
    nSerialNo As Integer
    szVehicleID As String
    szVehicleModelID As String
    dtBeginStopDate As Date
    dtEndStopDate As Date
    nStandTicketCount As Integer
End Type
''有不同座位类型的车次信息
'Public Type TMultiSeatTypeBus
'    adBusDate() As Date
'    aszBusID() As String
'End Type

'费率信息（基本运价率、公建金率等等）
Public Type TChargeRatio
    szAreaCode As String
    szVehicleModel As String
    szRoadLevel As String
    sgRoadConstructFundRatio As Double
    sgBaseCarriageRatio As Double
    szAnnotation As String
    szSeatType As String
End Type
'线路路段信息
Public Type TRouteSectionInfo
    szSectionID As String
    szSectionName As String
    nEndStationType As ERouteStationType
    sgEndStationMileage As Double
End Type
'座位状态
Public Type TSeatStatus
    szSeatNo As String
    nSeatStatus As Integer
    szTicketId As String
End Type
'并班信息
Public Type TMergeInfo
    szBusID As String
    nMergeType As Integer
    szAnonation As String
End Type
'**************************
'修改者:陈峰
'修改日期:2002-11-06
'运行环境票价信息
'**************************
Public Type TRETicketPrice
    szStationID As String
    szStationName As String
    nTicketType As Integer
    szTicketTypeName As String
    sgTotal As Double
    sgBase As Double
    szSeatType As String
    asgPrice(1 To 15) As Double
    
    szSellStationID As String
    szSellStationName As String
End Type
'**************************
'修改者:陈峰
'修改日期:2002-11-06
'运行环境票价信息
'**************************
Public Type TRETicketPriceEx
    szRouteID As String
    szRouteName As String
    szVehicleModelID As String
    szVehicleModelName As String
    szStationID As String
    szStationName As String
    sgMileage As Double
    nTicketType As Integer
    szTicketName As String
    sgTotal As Double
    sgBase As Double
    szSeatType As String
    szSeatTypeName As String
    asgPrice(1 To 15) As Double
    
    szSellStationID As String
    szSellStationName As String
End Type


'退票信息
Public Type TReturnedTicketInfo
    dtReturnTime As Date
    szCredenceID As String
    sgReturnCharge As Double
    szOperatorID As String
    nReturnType As Integer
End Type
'**************************
'修改者:陈峰
'修改日期:2002-11-06
'改签票信息
'**************************
Public Type TChangedTicketInfo
    szTicketId As String
    sgTicketPrice As Double
    szCredenceID As String
'    nChangeType As EChangeType
    sgChangeCharge As Double
    dtDate As Date
    dtBusID As String
    szBeginStationID As String
    szBeginStationName As String
    szEndStationID As String
    szEndStationName As String
    nTicketType As ETicketType
    szSeatNo As String
    szSellStationID As String
    szSellStationName As String
End Type
'废票信息
Public Type TCanceledTicketInfo
    dtCancelTime As Date
    szOperatorID As String
    nCancelType As Integer
End Type
''检票信息
'Public Type TCheckedTicketInfo
'    szTicketId As String
'    szBusID As String
'    dtBusDate As Date
'    nBusSerialNo As Integer '车次序号，对流水车次有用
'    dtCheckTime As Date
'    nCheckTicketType As Integer
'End Type
''扩展检票信息
'Public Type TCheckedTicketInfoEx
'    szTicketId As String
'    szBusID As String
'    dtBusDate As Date
'    nBusSerialNo As Integer '车次序号，对流水车次有用
'    dtCheckTime As Date
'    nCheckTicketType As Integer
'    szChecker As String
'    szCheckGate As String
'End Type
'检票信息
Public Type TCheckedTicketInfo
    szTicketId As String
    szBusID As String
    dtBusDate As Date
    nBusSerialNo As Integer '车次序号，对流水车次有用
    dtCheckTime As Date
    nCheckTicketType As Integer
    szCheckerID As String
    szCheckGateID As String
    
    szCheckerName As String
    szCheckGateName As String
    szSellStatonID As String
    szSellStatonName As String
End Type
'系统参数信息
Public Type TParamInfo
    szParamName As String
    szParamChineseName As String
    szAnnotation As String
    szParamValue As String
End Type
''票价信息（非运行环境）
'Public Type TPriceDetailInfo
'    szStationID As String '站点代码
'    szStationName As String '站点名称
'    szVehicleModel As String '车型代码
'    szVehicleModelName As String '车型名称
'    szRouteID As String '线路代码
'    szRouteName As String '线路名称
'    nTicketType As ETicketType '票种
'    szSeatTypeID As String '座位类型代码
'    szSeatTypeName As String '座位类型名称
'    sgMileage As Double '里程
'    sgTotalPrice As Double '总票价
'    nSerialNo As Integer '站点序号
'    sgBaseCarriage As Double '基本运价
'    asgItem(1 To 12) As Double '其它票价项
'End Type

'**************************
'修改者:陈峰
'修改日期:2002-11-06
'线路车次票价明细信息
'**************************
Public Type TBusPriceDetailInfo
    szBusID As String '车次代码
    dyStartTime As Date '发车时间
    nBusType As EBusType '车次类型
    szRouteID As String '线路代码
    szRouteName As String '线路名称
    szVehicleModel As String '车型代码
    szVehicleModelName As String '车型名称
    szAreaCode As String '区域代码
    szStationID As String '站点代码
    szStationName As String '站点名称
    nTicketType As ETicketType '票种
    szSeatTypeID As String '座位类型代码
    szSeatTypeName As String '座位类型名称
    sgMileage As Double '里程
    sgTotalPrice As Double '总票价
    nSerialNo As Integer '站点序号
    sgBaseCarriage As Double '基本运价
    asgItem(1 To 15) As Double '其它票价项


    szSellStationID As String
    szSellStationName As String
End Type

'线路票价项的公式和参数信息
Public Type TItemFAndP
    szPriceItem As String
    szFormula As String
    sgParam1 As String
    sgParam2 As String
    sgParam3 As String
    sgParam4 As String
    sgParam5 As String
End Type
'半票票价项信息
Public Type THalfTicketItemParam
    szTicketItem As String
    szTicketType As String
    szSeatType As String
    sgParam1 As Double
    sgParam2 As Double
    szAnnotation As String
End Type
'外部票价计算公式（组件）的信息
Public Type TPriceFormulaInfo
    szFunName As String  '函数名字
    szFunChineseName As String '函数中文名称
    nFunParamCount As Integer  '参数个数
    szFunIntroduce As String   '函数介绍
    aszParamIntroduce() As String '参数介绍
    szCheckParamValidFunName As String '此函数的原型为（参数1, 参数2…） as Boolean
    bBaseCarriage As Boolean '是否是计算基本运价的公式
End Type

'组件的函数信息
Public Type TCOMFunctionInfo
    szFunctionCode As String
    szFunctionName As String
    szFunctionGroup As String
    bIfWriteLog As Boolean
    szAnnotation As String
End Type

'组件的自我信息
Public Type TCOMSelfInfo
    szCOMCode As String
    szCOMName As String
    szCOMVersion As String
    nFunctionCount As Integer
    acfiFunction() As TCOMFunctionInfo
End Type

'**************************
'修改者:陈峰
'修改日期:2002-11-06
'购买车票信息
'**************************
Public Type TBuyTicketInfo
    nTicketType As ETicketType
    szTicketNo As String
    szSeatNo As String '空表示系统自已给座位号，'ST'表示站票
    szReserved As String
    szSeatTypeID As String  '座位类型代码
    szSeatTypeName As String '座位类型名称
    
End Type

'车次车辆预留座位信息
Public Type TReserveSeatInfo
    nSerialNo As Integer
    dtBeginDate  As Date
    dtEndDate As Date
    nStartSeatNo As Integer
    nSeatCount As Integer
End Type

'车次票价和座位信息
Public Type TBusSeatPriceInfo
    sgTicketPrice() As Single
    szSeatType() As String
End Type

'车次计划安排信息
Public Type TSchemeArrangement
    nSerialNo As Integer
    dtBeginDate  As Date
    szProjectID As String
    szAnnotation As String
End Type

'线路车型信息
Public Type TRouteVehicleModel
    szRouteID As String
    aszVehicleModel() As String
End Type

'售票结果
Public Type TSellTicketResult
    asgTicketPrice() As Double
    aszSeatNo() As String
    aszTicketType() As ETicketType
    aszSeatType() As String
    
    szDesStationName As String
    
    szBeginStationID As String
    szBeginStationName As String
    
    szReserved As String
End Type

'**************************
'修改者:陈峰
'修改日期:2002-11-06
'车次检票信息
'**************************
Public Type TBusCheckInfo
    dtDate As Date
    szBusID As String
    nBusSerialNo As Integer
    szVehicleID As String
    szCheckSheet As String
    dtStartupTime As Date
    dtBeginCheckTime As Date
    dtEndCheckTime As Date
    szChecker As String
    szCompanyID As String
    szSplitCompanyID As String
    szOwnerID As String
    
    szCheckGateID As String
    szCheckGateName As String
    szSellStationID As String
    szSellStationName As String

End Type

'**************************
'修改者:陈峰
'修改日期:2002-11-06
'路单信息
'**************************
Public Type TCheckSheetInfo
    szCheckSheet As String
    szBusID As String
    dtDate As Date
    nBusSerialNo As String
    szVehicleID As String
    szCompanyID As String
    szSplitCompanyID As String
    szRouteID As String
    szVehicleModelID As String
    szOwnerID As String
    dtStartupTime As Date
    bValid As Boolean
    dtSettleDateTime As Date
    szSettleSheetID As String
    nSettleSheetSerialNo As Integer
    szMakeSheetUser As String
    dtMakeSheetDateTime As Date
    
    szCheckGateID As String
    szCheckGateName As String
    szSellStationID As String
    szSellStationName As String

End Type

'路单车站信息
Public Type TCheckSheetStationInfo
    szCheckSheet As String
    szStationID As String
    nTicketType As ETicketType
    szSeatTypeID As String
    nManCount As Integer
    sgMileage As Double
    sgTicketPrice As Double

End Type



'路单车站信息
'因TCheckSheetStationInfo少返回了一个检票状态,即是否改乘,所以又加了这个类型
Public Type TCheckSheetStationInfoEx
    szCheckSheet As String
    szStationID As String
    nTicketType As ETicketType
    nCheckStatus As Integer '检票状态
    szSeatTypeID As String
    nManCount As Integer
    sgMileage As Double
    sgTicketPrice As Double
End Type

'路单车站信息
'添加了一个单价类型
Public Type TCheckSheetStationInfoEx2
    szCheckSheet As String
    szStationID As String
    nTicketType As ETicketType
    nCheckStatus As Integer '检票状态
    szSeatTypeID As String
    nManCount As Integer
    sgMileage As Double
    sgTicketPrice As Double
    szSinglePrice As Double
End Type



'访问控制信息
Public Type TRemoteUserInfo
    szRemoteUserID As String '远程用户帐号
    szAnnotation As String '注释
End Type

'公用的数据类型到此为止



Public Type TActiveUserInfo
    lLoginEventID As Long '登录事件ID
    lLoginEventID2 As Long '登录事件ID2
    szUserID As String '登录用户ID
    dtLoginTime As Date '登录时间
    szLoginHost As String '登录工作站
    dtLastTime As Date '最后使用时间
End Type

'------------------------------------
''单位信息
'Public Type TUnit
'    szUnitID As String
'    szUnitShortName As String
'    szUnitFullName As String
'    nUnitType As EUnitType
'    szIPAddress As String
'    szAnnotation As String
'    szStationID As String '本单位的车站代码
'
'End Type
''单位信息
'Public Type TUnitEx
'    szUnitID As String
'    szUnitShortName As String
'    szUnitFullName As String
'    nUnitType As EUnitType
'    szIPAddress As String
'    szAnnotation As String
'    dbSellCharge As Double '售票手续费
'    szStationID As String '本单位的车站代码
'
'
'End Type
'单位信息
Public Type TUnit
    szUnitID As String
    szUnitShortName As String
    szUnitFullName As String
    nUnitType As EUnitType
    szIPAddress As String
    szAnnotation As String
    dbSellCharge As Double '售票手续费
    szStationID As String '本单位的车站代码
End Type

Public Type TTicketType
    nTicketTypeID As Integer
    szTicketTypeName As String
    nTicketTypeValid As Integer
    szAnnotation As String
End Type


'**************************
'修改者:陈峰
'修改日期:2002-11-13
'车次配载信息
'**************************
Public Type TBusAllotInfo
    szBusID As String
    dtRunTime As Date
    szProjectID As String
    szCheckGateID As String
    szCheckGateName As String
    szSellStationID As String
    szSellStationName As String
    nCanSellQuantity As Integer '允许售票张数
End Type

'**************************
'修改者:陈峰
'修改日期:2002-11-13
'环境车次配载信息
'**************************
Public Type TReBusAllotInfo
    dyBusDate As Date
    szBusID As String
    szSellStationID As String
    szSellStationName As String
    dtRunTime As Date
    szCheckGateID As String
    szCheckGateName As String
    nStatus As Integer  '车次状态
    nCanSellQuantity As Integer '允许售票张数
    nSellQuantity As Integer  '已售票数量
End Type


'**************************
'修改者:陈峰
'修改日期:2002-11-06
'部门信息
'**************************
Public Type TDepartmentInfo
    szSellStationID As String
    szSellStationName As String
    szSellStationFullName As String
    szUnitID As String
    szUnitName As String
    szUnitFullName As String
    szStationID As String '对应的站点代码
    szStationName As String
    szAnnotation As String
End Type

'**************************
'修改者:陈峰
'修改日期:2002-11-06
'用户信息
'**************************
Public Type TUserInfo
    UserID As String
    UserName As String
    InnerUser As Boolean
    UnitID As String
    LoginIP As String
    CanSellDay As Integer
    MultLogin As Integer
    StartTime As Date
    EndTime As Date
    Week As Integer
    Annotation As String
    IsInternetUser As Boolean
    
    SellstationID As String '售票站代码
    UnitName As String '单位名称
    SellStationName As String '售票站名称
End Type
'**************************
'修改者:陈峰
'修改日期:2002-11-06
'用户组信息
'**************************
Public Type TUserGroupInfo
    UserGroupID As String
    GroupName As String
    InnerGroup As Boolean
    Annotation As String
End Type
'**************************
'修改者:陈峰
'修改日期:2002-11-06
'组件功能信息
'**************************
Public Type TCOMFunctionInfo1
    szFunctionCode As String
    szFunctionGroup As String
    szFunctionName As String
    szAnnotation As String
    bIfWriteLog As Boolean
    szcomponentID As String
End Type
'**************************
'修改者:陈峰
'修改日期:2002-11-06
'组件信息
'**************************
Public Type TCOMSelfInfo1
    szCOMCode As String
    szCOMName As String
    szCOMVersion As String
    dtLoadTime As Date
End Type

'**************************
'修改者:陈峰
'修改日期:2002-11-07
'扩展的线路路段信息
'**************************
Public Type TRouteSectionInfoEx
    nEndStationType As Integer
    sgEndStationMileage As Double
    szSectionID As String
    nSectionSerial As Integer
    szEndStationID As String
    szEndStationName As String
    szStartStationID As String
    szStartStationName As String
    szSectionName As String
End Type


'**************************
'修改者:陈峰
'修改日期:2002-11-07
'查询票价信息
'**************************
Public Type TQueryTicketPrice
    szStationID As String
    szStationName As String
    szVehicleModel As String
    nTicketType As Integer
    sgMileage As Double
    sgTotal As Double
    sgBase As Double
    asgPrice(1 To 15) As Double
End Type

''**************************
''修改者:陈峰
''修改日期:2002-11-07
''查询环境站点信息
''**************************
'Public Type TREBusStationSellInfoEx
'        szStationID  As String
'        szSeatTypeID As String
'        szStationName As String
'        nMileage As Double
'        nLimitedSellTime As Integer
'        nLimitedSellCount As Integer
'        sgFullPrice As Double
'        sgHalfPrice As Double
'        sgPreferentialPrice1 As Double '优惠票1
'        sgPreferentialPrice2 As Double '优惠票2
'        sgPreferentialPrice3 As Double '优惠票3
'        sgBasePrice As Double
'End Type
'**************************
'修改者:陈峰
'修改日期:2002-11-07
'半价参数信息
'**************************
Public Type THalfTicketItemParamEX
   asgParam1(1 To 16) As Double
   asgParam2(1 To 16) As Double
   sTicketType As Integer
   szAnnotation(1 To 16) As String
End Type
'**************************
'修改者:陈峰
'修改日期:2002-11-07
'座位类型信息
'**************************
Public Type TSeatInfoCountEx
  tTSeatTypeInfo(1 To 5) As TSeatInfoCount
  szBusID As String
End Type
'**************************
'修改者:陈峰
'修改日期:2002-11-07
'座位信息
'**************************
Public Type TSeatInfo
    szSeatNo As String
    szSeatStatus As ERESeatStatus
    szSeatType As String
    szSeatTypeName As String
    szTicketNo As String
    szDestName As String
    szTicketPrice As Currency
End Type

Public Type TSeatInfoEx
    szSeatNo As String
    szSeatStatus As ERESeatStatus
    szSeatType As String
    szSeatTypeName As String
    szTicketNo As String
    szDestName As String
    szTicketPrice As Currency
    szRemark As String
End Type





'财务系统路单拆算的数据结构。


Public Type TSelletStation
    szStationID As String '站点代码
    szStationName As String '站点名称
    nPriceIdentifyNo As Integer '票价号
    nChangeMergeStatus As Integer '改并状态
    sgKilometre As Single '里程数
    nFullHumenCount As Integer '全票人数
    nHalfHumenCount As Integer '半票人数
    nPreferentialCount1 As Integer '优惠票1人数
    nPreferentialCount2 As Integer '优惠票2人数
    nPreferentialCount3 As Integer '优惠票3人数
    dbFullTotalPrice As Double '全票价
    dbHalfTotalPrice As Double '半票价
    dbPreferentialPrice1 As Double '优惠票1价
    dbPreferentialPrice2 As Double '优惠票2价
    dbPreferentialPrice3 As Double '优惠票3价
    dbFullPriceItem(1 To 16) As Double '全票价票价项
    dbHalfPriceItem(1 To 16) As Double '半票价票价项
    dbPreferentialItem1(1 To 16) As Double '半票价票价项
    dbPreferentialItem2(1 To 16) As Double '半票价票价项
    dbPreferentialItem3(1 To 16) As Double '半票价票价项
    dtDate As Date '日期
    vData As Variant
    
    
    szSellStationID As String  '上车站代码
    szSellStationName As String  '上车站名称
    
End Type


'财务系统行包拆算的数据结构。
Public Type TSplitLuggage
    szLuggageSheetID As String '行包签发单号
    dbTotalCarriage As Double  '总运价
    dbLuggageCharge As Double '拆算票款
    dbTotalPiece  As Long  '总件数
    dbPriceItem(1 To 10) As Double '行包票价项
    vData As Variant '储存其他
End Type

'路单拆算明细
Public Type TSplitCsDetail
    szCheckSheetID As String '路单编号
    szBusID As String '车次代码
    nBusSerial As Integer ' 车次序号
    dtBusDate As Date '车次日期
    szVehicleID As String '车辆代码
    szLicenseTagNo As String '车牌
    szRouteID As String '线路代码
    szRouteName As String '线路名称
    szCompanyName As String '参运公司名称
    szCompanyID As String '参运公司代码
    szSplitCompanyName As String '拆帐公司名称
    szSplitCompanyID As String '拆帐公司代码
    szOwnerID As String '车主代码
    szOwnerName As String '车主名称
    dtStartTime As Date '发车时间
    szItemID As String  '代码
    szItemName As String '名称
    dbItemMoney As Double '拆算项票款
    szItemFormula As String '拆算公式
    eItemType As EFormulaOrMoney '费用类型
    dtSettleDate As Date         '结算日期
End Type

'行包拆算明细
Public Type TSplitLsDetail
    szCheckLuggageID As String '签发单编号
    szBusID As String '车次代码
    nBusSerial As Integer ' 车次序号
    dtBusDate As Date '车次日期
    szVehicleID As String '车辆代码
    szLicenseTagNo As String '车牌
    szRouteID As String '线路代码
    szRouteName As String '线路名称
    szCompanyName As String '参运公司名称
    szCompanyID As String '参运公司代码
    szSplitCompanyName As String '拆帐公司名称
    szSplitCompanyID As String '拆帐公司代码
    szOwnerID As String '车主代码
    szOwnerName As String '车主名称
    dtStartTime As Date '发车时间
    szItemID As String  '代码
    szItemName As String '名称
    dbItemMoney As Double '拆算项票款
    szItemFormula As String '拆算公式
    eItemType As EFormulaOrMoney '费用类型
    dtSettleDate As Date         '结算日期
End Type


    
'路单信息类型TCheckSheet
Public Type TCheckSheet
    szCheckSheetID As String '路单编号
    szBusID As String '车次代码
    nBusSerial As Integer ' 车次序号
    dtBusDate As Date '车次日期
    szVehicleID As String '车辆代码
    szLicenseTagNo As String '车牌
    szRouteID As String '线路代码
    szRouteName As String '线路名称
    szCompanyName As String '参运公司名称
    szCompanyID As String '参运公司代码
    szSplitCompanyName As String '拆帐公司名称
    szSplitCompanyID As String '拆帐公司代码
    szOwnerID As String '车主代码
    szOwnerName As String '车主名称
    dtStartTime As Date '发车时间
    tdStartDate As Date '结算开始日期
    dtEndDate As Date '结算结束日期
    Availability As ECheckStatus '路单状态1有效 2无效 3已结算 4多次结算
    BusType As EBusType '车次类型
    szSettleSheetID As String '结算单号
    szSettleSheetDate As Date '结算日期
    szSettleUnit As String    '结算单位
    szChecker As String         '结算人
    szProtocol As String '协议代码
    dbTotalPrice As Double '总票款
    dbSelletPrice As Double  '结算票款
    nStatus As Integer       '结算状态
    dbCheckCharge As Double '拆算票款
    nQuantity As Long  '总人数
    vData As Variant '储存其他
    
    szCheckGateID As String
    szCheckGateName As String
    szSellStationID As String
    szSellStationName As String
    
End Type

'签发单信息类型TCheckSheet
Public Type TLuggageSheet
    szLuggageSheetID As String '签发单编号
    szBusID As String '车次代码
    nBusSerial As Integer ' 车次序号
    dtBusDate As Date '车次日期
    szVehicleID As String '车辆代码
    szLicenseTagNo As String '车牌
    szRouteID As String '线路代码
    szRouteName As String '线路名称
    szCompanyName As String '参运公司名称
    szCompanyID As String '参运公司代码
    szSplitCompanyName As String '拆帐公司名称
    szSplitCompanyID As String '拆帐公司代码
    szOwnerID As String '车主代码
    szOwnerName As String '车主名称
    dtStartTime As Date '发车时间
    tdStartDate As Date '结算开始日期
    dtEndDate As Date '结算结束日期
    Availability As ECheckStatus '签发单状态1有效 2无效 3已结算 4多次结算
    BusType As EBusType '车次类型
    szSettleSheetID As String '结算单号
    szSettleSheetDate As Date '结算日期
    szSettleUnit As String    '结算单位
    szChecker As String         '结算人
    szProtocol As String '协议代码
    dbTotalCarriage As Double '总票款
    dbTotalPiece As Double  '总件数
    dbLuggageCharge As Double '拆算票款
     nStatus As Integer       '结算状态
    vData As Variant '储存其他
End Type

''结算路单信息结构
'Public Type TSellet
'    szSettleSheetID As String
'    nSettleObject As Integer
'    szObjectID As String
'    szBusID As String '车次代码
'    nBusSerial As Integer ' 车次序号
'    dtBusDate As Date '车次日期
'    szVehicleID As String '车辆代码
'    szLicenseTagNo As String '车牌
'    szRouteID As String '线路代码
'    szRouteName As String '线路名称
'    szCompanyName As String '参运公司名称
'    szCompanyID As String '参运公司代码
'    szSplitCompanyName As String '拆帐公司名称
'    szSplitCompanyID As String '拆帐公司代码
'    szOwnerID As String '车主代码
'    szOwnerName As String '车主名称
'    dtStartTime As Date '发车时间
'    dtStartDate As Date '结算开始日期
'    dtEndDate As Date '结算结束日期
'    szSettlementUnit As String '结算单位
'    szMoneyReceiver As String '收款人
'    szSettlementChecker As String '复核
'    szSettlementExecutive As String '拆算人
'    szProtocol As String '拆算协议
'    nProtocolType As EProtocolType  '协议类型
'    dtSettlementDate As Date '结算日期
'    dbTotalTicketPrice As Double '总票款
'    nTotalQuantity As Integer '总人数
'    dbCheckCharge As Double '拆算路单票款
'    dbCheckCompensate As Double
'    dbTotalPiece As Double  '总件数
'    dbTotalCarriage As Double
'    dbLuggageCharge As Double '拆算行包票款
'    dbLuggageCompensate As Double
'    dbSettlementPrice As Double '结算票款
'    szAnnotation As String '注释
'    szAccount As String    '帐号
'    vData As Variant
'End Type

'结算路单信息结构
Public Type TSellet
    szSettleSheetID As String
    nSettleMethod As Integer        '结算方法： 0 － 自动；1 －手工
    szDefaultBusID As String        '默认车次（手工结算使用）
    dtDefaultBusDate As Date        '默认车次日期（手工结算使用）
    nSettleObject As Integer
    szObjectID As String
    szBusID As String               '车次代码
    nBusSerial As Integer           ' 车次序号
    dtBusDate As Date               '车次日期
    szVehicleID As String           '车辆代码
    szLicenseTagNo As String        '车牌
    szRouteID As String             '线路代码
    szRouteName As String           '线路名称
    szCompanyName As String         '参运公司名称
    szCompanyID As String           '参运公司代码
    szSplitCompanyName As String    '拆帐公司名称
    szSplitCompanyID As String      '拆帐公司代码
    szOwnerID As String             '车主代码
    szOwnerName As String           '车主名称
    dtStartTime As Date             '发车时间
    dtStartDate As Date             '结算开始日期
    dtEndDate As Date               '结算结束日期
    szSettlementUnit As String      '结算单位
    szMoneyReceiver As String       '收款人
    szSettlementChecker As String   '复核
    szSettlementExecutive As String '拆算人
    szProtocol As String            '拆算协议
    nProtocolType As EProtocolType  '协议类型
    dtSettlementDate As Date        '结算日期
    dbTotalTicketPrice As Double    '总票款
    nTotalQuantity As Long          '总人数
    dbCheckCharge As Double         '拆算路单票款
    dbCheckCompensate As Double
    dbTotalPiece As Double          '总件数
    dbTotalCarriage As Double
    dbLuggageCharge As Double       '拆算行包票款
    dbLuggageCompensate As Double
    dbSettlementPrice As Double     '结算票款
    szPaymentType As String         '结算方式（现金/支票/电汇等）
    szAnnotation As String          '注释
    szAccount As String             '帐号
    vData As Variant
    nSettleTag As Integer           '结算标记
    dbTotalLsSplitOut As Double     '拆出行包款(拆出)
    dbTotalCsSplitOut As Double     '拆出路单款(拆出)
End Type


'结算行包信息结构
Public Type TSelletLuggage
    szSettleSheetID As String
    nSettleType As Integer          '结算类型： 0 － 签发单结算；1 －非旅客同行行包结算
    szDefaultBusID As String        '默认车次（手工结算使用）
    dtDefaultBusDate As Date        '默认车次日期（手工结算使用）
    szBusID As String               '车次代码
    nBusSerial As Integer           '车次序号
    dtBusDate As Date               '车次日期
    szVehicleID As String           '车辆代码
    szLicenseTagNo As String        '车牌
    szRouteID As String             '线路代码
    szRouteName As String           '线路名称
    szCompanyName As String         '参运公司名称
    szCompanyID As String           '参运公司代码
    szSplitCompanyName As String    '拆帐公司名称
    szSplitCompanyID As String      '拆帐公司代码
    szOwnerID As String             '车主代码
    szOwnerName As String           '车主名称
    dtStartTime As Date             '发车时间
    dtStartDate As Date             '结算开始日期
    dtEndDate As Date               '结算结束日期
    szSettlementUnit As String      '结算单位
    szMoneyReceiver As String       '收款人
    szSettlementChecker As String   '复核
    szSettlementExecutive As String '拆算人
    szProtocol As String            '拆算协议
    dtSettlementDate As Date        '结算日期
    dbTotalCarriage As Double       '总票款
    dbTotalPiece As Double          '总件数
    dbLuggageCharge As Double       '拆算行包票款
    dbSettlementPrice As Double     '结算票款(拆入金额)
    szPaymentType As String         '结算方式（现金/支票/电汇等）
    szAnnotation As String          '注释
    szAccount As String             '帐号
    vData As Variant
    nSettleTag As Integer           ' 结算标记
    dbTotalLsSplitOut As Double     '拆出行包款(拆出)
    dbTotalCsSplitOut As Double     '拆出路单款(拆出)
End Type


'路单状态及行包状态
Public Enum ECheckStatus
    NoAvailability = 0 '无效
    Availability = 1 '有效
    SettleFinished = 2 '已结算
    RepetitiousSettle = 3 '多次结算
    AllCheckSheet = 4
End Enum

'路单站点信息结构
Public Type TFinStation
    szStationID As String '站点代码
    nStationSerial As Integer '站点序号
    szStationName As String '站点名称
    eCSTicketType As ETicketType '检票类型
    nPriceIdentifyNo As Integer '票价唯一标示
    nChangeMergeStatus As Integer '改并状态
    dbKilometre As Single '里程数
    nHumenCount As Integer '人数
    dbTotalPrice As Double '总票价
    dbPriceItem(1 To 16) As Double '票价项
    szSeatTypeID As String   '座位类型
    szSeatTypeName As String   '座位类型
    
    
    szSellStationID As String  '上车站代码
    szSellStationName As String  '上车站名称
    
End Type

Public Enum ESplitTag
    NormalCheck = 0
    GrantCheck = 1
    Sale = 2
    Luggage = 3
End Enum

Public Enum EDefault
    Default = 0
    NotDefaule = 1
End Enum

Public Type TProtocol '协议信息结构
    szProtocolID As String
    szAnnotation As String
    szProtocolName As String
    eDefaultMark As EDefault
    nProtocolType As EProtocolType
    vData As Variant
End Type

Public Enum EProtocolType
    SocietyBus = 0             '社会车辆
    CommonBus = 1              '常运车辆
End Enum

Public Type TSplitChargeItem   '拆算项信息结构
    szItemID As String
    szItemName As String
    dbItemMoney As Double
    dbCompensate As Double     '补差
    szItemFormula As String
    szFormularName As String   '公式名
    eItemType As EFormulaOrMoney
    eType As ESplitType
    ItemStatus As ESplitStatus
End Type

Public Type TSplitItem   '拆算项结构
    szItemID As String
    szItemName As String
    eType As Integer
    eStatus As ESplitStatus
End Type

Public Enum EFormulaOrMoney
    MonthMoney = 0       '月费用
    DayMoney = 1         '日费用
    UnUseful = 2
End Enum

Public Enum ESplitStatus
    GetIn = 0           '拆入
    GetOut = 1          '拆出
End Enum

Public Enum ESplitType
    CheckType = 0           '路单信息项
    LuggageType = 1         '行包信息项
End Enum

Public Type TBusCompensateInfo
    szBusID As String                       '补差车次
    dtBusStartDate As Date                  '发车时间
    dbMoneyCompensateToBus As Double        '车次补差金额
    dbMoneyCompensateToUnit As Double       '车主补差金额
    szCompensateOwnerID As String           '车主ID
    szCompensateUnitID As String            '用户单位
    dtCompensateDate As Date                '补差日期
    dbCompensateMoney As Double             '结算金额
    nSerialNumber As Integer                '车次序号
End Type

Public Type TInputBusInfo
    szBusID As String                       '车次ID
    dtBusStartDate As Date                  '发车时间
    nSerialNumber As Integer                '车次序号
    dbSettleMoney As Double                 '余额
End Type

'***此类型可以删除,已不用
Public Type TBusStationInfo         '站点信息
    szStationID As String           '站点代码
    szStationName As String         '站点名称
    szStationInputCode As String    '站点
End Type

Public Type TSplitConditionInfo     '拆帐条件信息（票价条件）
    szBusID As String               '车次           y
    szBusDate As Date               '车次日期       y
    szStationID As String           '站点代码       y
    szStationName As String         '站点名称
    szTicketTypeID As String        '票种代码       y
    szTicketTypeName As String      '票种名称
    szSeatTypeID As String          '座位类型代码   y
    szSeatTypeName As String        '座位类型名称
    szProtocolID As String          '协议代码
    szProtocolName As String        '协议名称
    nProtocolType As Integer        '协议类型
    szProtocolRem  As String        '协议注释（暂时未使用）
    dbQuantity As Double            '统计人数
    dbTicketPrice As Double         '票价
    nPriceIdentify As Integer       '票价表序号（暂时未使用）
    szVehicleTypeID As String       '车辆类型代码
    szVehicleTypeName As String     '车辆类型名称
    szSellStationID As String       '上车站代码
    szSellStationName As String     '上车站名称
End Type










'本类（活动用户的错误号）
Public Enum EErrActiveUser
    ERR_ActiveUserNotExist = ERR_ActiveUser + 10
    ERR_ActiveUserNotLogin = ERR_ActiveUser + 11
    ERR_ActiveUserTimeOut = ERR_ActiveUser + 12
    ERR_ActiveUserTwoPasswordNotSame = ERR_ActiveUser + 13
    ERR_ActiveUserOldPasswordNotOk = ERR_ActiveUser + 14
    ERR_ActiveUserReachMaxLoginCount = ERR_ActiveUser + 15
    ERR_ActiveUserNotLegalWorkstation = ERR_ActiveUser + 16
    ERR_ActiveUserMustHaveLoginWorkstation = ERR_ActiveUser + 17
    ERR_ActiveUserUserNotExistOrPasswordNotOk = ERR_ActiveUser + 18 '用户不存在或密码错
    ERR_ActiveUserUserCanNotLoginAtThisTime = ERR_ActiveUser + 19 '用户不能在这个时间登录
    ERR_ActiveUserUserCanNotLoginAtThisDate = ERR_ActiveUser + 20 '用户不能在这天登录
    
End Enum
'------------------------------------
Private m_oActiveUserData As ActiveUserData
Private m_szConnectionString As String '亦或DBUse.RTConnection

Private Sub WriteLoginLog(poDb As RTConnection)
'    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim szIPAdd As String
    szIPAdd = GetComputerIP(m_oActiveUserData.m_szLoginHost)
    szSql = "INSERT INTO  login_log_lst(login_start_time,user_id,computer_name,ip_address) " _
    & " VALUES('" & ToDBDateTime(m_oActiveUserData.m_dtLoginTime) & "','" _
    & m_oActiveUserData.m_szUserID & "','" _
    & m_oActiveUserData.m_szLoginHost & "','" _
    & szIPAdd & "')"
    poDb.Execute szSql
    
    szSql = "SELECT @@IDENTITY AS test"
    
    Set rsTemp = poDb.Execute(szSql)
    m_oActiveUserData.m_lLoginEventID2 = FormatDbValue(rsTemp!test)
    
#If ACTIVEUSER_USE_MEMORY = 0 Then
    szSql = "UPDATE active_user_info SET login_id2=" & m_oActiveUserData.m_lLoginEventID2 & " WHERE login_id=" & m_oActiveUserData.m_lLoginEventID
    poDb.Execute szSql
#End If

End Sub

Private Sub WriteLogoutLog(Optional pnLogoutType As ELogoutType = TP_NormalLogout)

    Dim oDb As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "UPDATE login_log_lst SET " _
    & " login_off_type =" & pnLogoutType _
    & ",login_off_time='" & ToDBDateTime(Now) & "' " _
    & " WHERE login_event_id=" & m_oActiveUserData.m_lLoginEventID2
    oDb.Execute szSql

End Sub

Private Function IsLoginEventIDExist(Optional ByVal plLoginEventID As Long = 0) As Boolean '当判断自己时，不但要判断m_lLoginEventID是否存在，还要比较m_dtLoginTime是否相同

    Dim oTemp As ActiveUserData
    Dim lEventID As Long
    
    If plLoginEventID = 0 Then
        lEventID = m_oActiveUserData.m_lLoginEventID
    Else
        lEventID = plLoginEventID
    End If

#If ACTIVEUSER_USE_MEMORY Then
    
    On Error GoTo here
    Set oTemp = m_clActiveUserData(cszCLPrefix & lEventID)
    If plLoginEventID = 0 Then
        If m_oActiveUserData.m_dtLoginTime = oTemp.m_dtLoginTime Then
            IsLoginEventIDExist = True
        Else
            IsLoginEventIDExist = False
        End If
    Else
        IsLoginEventIDExist = True
    End If
    Exit Function
    
here:
    IsLoginEventIDExist = False
#Else
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "SELECT * FROM active_user_info WHERE login_id=" & lEventID
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 1 Then
        IsLoginEventIDExist = False
    Else
        If plLoginEventID = 0 Then
            If m_oActiveUserData.m_dtLoginTime = FormatDbValue(rsTemp!login_time) Then
                IsLoginEventIDExist = True
            Else
                IsLoginEventIDExist = False
            End If
        Else
            IsLoginEventIDExist = True
        End If
    End If
    Exit Function
#End If

End Function


Public Property Get ActiveUserStatus() As EActiveUserStatus
    
    If m_oActiveUserData Is Nothing Then
        ActiveUserStatus = ST_AUNotLogin
    ElseIf Not IsLoginEventIDExist() Then
        ActiveUserStatus = ST_AUTimeOut
    Else
        ActiveUserStatus = ST_AUValid
    End If
    
End Property


Public Property Get LogEventID() As Long
    Me.AssertActiveUserValid
    LogEventID = m_oActiveUserData.m_lLoginEventID
End Property


Public Property Get UserID() As String
    Me.AssertActiveUserValid
    UserID = m_oActiveUserData.m_szUserID
    
End Property

Public Property Get ServerUnit() As String
    Me.AssertActiveUserValid
'    Dim oParam As New SystemParam
'
'    oParam.Init Me
'    ServerUnit = oParam.UnitID
    ServerUnit = GetUnitID()
End Property

Public Property Get LoginTime() As Date
    Me.AssertActiveUserValid
    LoginTime = m_oActiveUserData.m_dtLoginTime
End Property

Public Property Get LoginHost() As String
    Me.AssertActiveUserValid
    LoginHost = m_oActiveUserData.m_szLoginHost
End Property

Public Property Get LastTime() As Date
    If Not m_oActiveUserData Is Nothing Then
        LastTime = m_oActiveUserData.m_dtLastTime
    End If
    Me.AssertActiveUserValid
End Property


Public Sub Identify(ByVal plLoginEventID As Long)
    
    If Not m_oActiveUserData Is Nothing Then
        If m_oActiveUserData.m_lLoginEventID = plLoginEventID Then Exit Sub
    End If
    
    If Not IsLoginEventIDExist(plLoginEventID) Then
        ShowError ERR_ActiveUserNotExist
    End If
    SelfLogout
    RefreshInfo plLoginEventID
    
End Sub

Public Function Login(ByVal pszUserID As String, ByVal pszPassword As String, ByVal pszLoginHost As String) As Long
    Dim dtNow As Date
    Dim bInternetUser As Boolean
    
    Dim rsTemp As Recordset
    
    bInternetUser = AssertCanLogin(pszUserID, pszPassword, pszLoginHost)
    
    SelfLogout
    Set m_oActiveUserData = New ActiveUserData
    dtNow = SelfNowDateTime()
    m_oActiveUserData.m_lLoginEventID = GetNextEventID(Asc(Left(pszUserID, 1)))
    
    m_oActiveUserData.m_szUserID = pszUserID
    m_oActiveUserData.m_szLoginHost = pszLoginHost
    'm_oActiveUserData.m
    m_oActiveUserData.m_dtLoginTime = dtNow
    m_oActiveUserData.m_dtLastTime = dtNow
    m_oActiveUserData.m_bInternetUser = bInternetUser
    
#If ACTIVEUSER_USE_MEMORY Then
    'Dim oActiveUserData As New ActiveUserData
    m_clActiveUserData.Add m_oActiveUserData, cszCLPrefix & m_oActiveUserData.m_lLoginEventID
    Login = m_oActiveUserData.m_lLoginEventID
#Else
    Dim oDb As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    '得到用户所属售票站,2002-11-12陈峰
    szSql = " SELECT sell_station_id FROM user_info WHERE user_id = " & TransFieldValueToString(pszUserID)
    Set rsTemp = oDb.Execute(szSql)
    m_oActiveUserData.m_szSellStationID = FormatDbValue(rsTemp!sell_station_id)
    '得到用户所属售票站,2002-11-12陈峰
    oDb.BeginTrans
    On Error GoTo here
    
    szSql = "INSERT INTO active_user_info(login_id,user_id,host_name,login_time,last_time,internet_user_status) VALUES(" _
    & m_oActiveUserData.m_lLoginEventID & ",'" _
    & m_oActiveUserData.m_szUserID & "','" _
    & m_oActiveUserData.m_szLoginHost & "','" _
    & ToDBDateTime(dtNow) & "','" _
    & ToDBDateTime(dtNow) & "'," _
    & IIf(bInternetUser, 1, 0) & ")"
    
    oDb.Execute szSql
    
    WriteLoginLog oDb
    oDb.CommitTrans
    Login = m_oActiveUserData.m_lLoginEventID
    Exit Function
here:
    oDb.RollbackTrans
    err.Raise err.Number
  
#End If

End Function

Public Sub Logout(Optional ByVal pnLogoutType As ELogoutType = TP_NormalLogout)

    Me.AssertActiveUserValid

    #If ACTIVEUSER_USE_MEMORY Then
        m_clActiveUserData.Remove cszCLPrefix & m_oActiveUserData.m_lLoginEventID
        WriteLogoutLog
    #End If

    SelfLogoutActiveUser m_oActiveUserData.m_lLoginEventID, m_oActiveUserData.m_lLoginEventID2, pnLogoutType
    Set m_oActiveUserData = Nothing
End Sub

Public Sub ChangePassword(ByVal pszOldPassword As String, ByVal pszNewPassword As String, ByVal pszReNewPassword As String)
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim lRowAffect As Long
    Me.AssertActiveUserValid
    If pszNewPassword <> pszReNewPassword Then ShowError ERR_ActiveUserTwoPasswordNotSame
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "UPDATE user_info SET " _
    & " user_password='" & EncodePassword(pszNewPassword) _
    & "' WHERE user_id='" & m_oActiveUserData.m_szUserID & "' AND " _
    & " user_password='" & EncodePassword(pszOldPassword) & "'"
    oDb.Execute szSql, lRowAffect
    If lRowAffect <> 1 Then ShowError ERR_ActiveUserOldPasswordNotOk
End Sub

Private Sub RefreshInfo(ByVal plLogEventID As Long)

#If ACTIVEUSER_USE_MEMORY Then
    Set m_oActiveUserData = m_clActiveUserData(cszCLPrefix & plLogEventID)
#Else
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "SELECT * FROM active_user_info WHERE login_id=" & plLogEventID
    Set rsTemp = oDb.Execute(szSql)
    Set m_oActiveUserData = New ActiveUserData
    m_oActiveUserData.m_lLoginEventID = FormatDbValue(rsTemp!login_id)
    m_oActiveUserData.m_dtLastTime = FormatDbValue(rsTemp!last_time)
    m_oActiveUserData.m_dtLoginTime = FormatDbValue(rsTemp!login_time)
    m_oActiveUserData.m_lLoginEventID2 = FormatDbValue(rsTemp!login_id2)
    m_oActiveUserData.m_szLoginHost = FormatDbValue(rsTemp!host_name)
    m_oActiveUserData.m_szUserID = FormatDbValue(rsTemp!user_id)
#End If

End Sub


Private Sub Class_Terminate()
    SelfLogout
End Sub

Private Sub SelfLogout()
    If ActiveUserStatus = ST_AUValid Then
'        OutDebugInfo m_oActiveUserData.m_lLoginEventID
        Logout
    End If
End Sub

'判断指定的用户是否能够登录(即密码正确否，判断是否可以同时登录，和登录工作站)
Private Function AssertCanLogin(pszUserID As String, pszPassword As String, pszLoginHost As String) As Boolean
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szWorkstation As String
    Dim nLoginCount As Integer
    Dim szPassword As String
    Dim rsTemp As Recordset
    Dim nLoginWeek As Integer
    Dim dtBeginTime As Date
    Dim dtEndTime As Date
    Dim dtNow As Date
    Dim nTemp As Integer
    
    If Trim(pszLoginHost) = "" Then ShowError ERR_ActiveUserMustHaveLoginWorkstation
    
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "SELECT * FROM user_info WHERE user_id='" & pszUserID & "' AND " _
    & " user_password='" & EncodePassword(pszPassword) & "' AND " _
    & " user_delete_status=1"
    
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 1 Then ShowError ERR_ActiveUserUserNotExistOrPasswordNotOk
    AssertCanLogin = IIf(FormatDbValue(rsTemp!internet_user_status) = 0, False, True)
    
    nLoginCount = FormatDbValue(rsTemp!is_simultaneity)
    szWorkstation = FormatDbValue(rsTemp!login_ip)
    
    nLoginWeek = FormatDbValue(rsTemp!Week)
    dtBeginTime = FormatDbValue(rsTemp!login_start_time)
    dtEndTime = FormatDbValue(rsTemp!login_end_time)
    
    dtNow = SelfNowDateTime()
    
    If dtBeginTime <> cdtEmptyTime Then
        If TimeDiff("s", dtNow, dtEndTime) < 0 Then ShowError ERR_ActiveUserUserCanNotLoginAtThisTime
    End If
    
    If dtEndTime <> cdtEmptyDate Then
        If TimeDiff("s", dtNow, dtBeginTime) > 0 Then ShowError ERR_ActiveUserUserCanNotLoginAtThisTime
    End If
'    dtNow = SelfNowDate()

    If nLoginWeek <> 0 Then
        nTemp = Weekday(dtNow) - 1
        If (2 ^ nTemp And nLoginWeek) = 0 Then ShowError ERR_ActiveUserUserCanNotLoginAtThisDate
    End If
    
    
    If nLoginCount >= 0 Then '小于0 不限,为0禁用,大于0,限制相应的登录次数
        If nLoginCount = 0 Then
            ShowError ERR_ActiveUserReachMaxLoginCount
        Else
            If nLoginCount <= GetSpecialActiveUserCount(pszUserID) Then
                ShowError ERR_ActiveUserReachMaxLoginCount
            End If
        End If
    End If
    
    If szWorkstation <> "" Then
        If InStr(1, szWorkstation, pszLoginHost, vbTextCompare) = 0 Then ShowError ERR_ActiveUserNotLegalWorkstation
    End If
    
End Function

Public Sub IAmLive()
    'Me.AssertActiveUserValid
    If Not m_oActiveUserData Is Nothing Then
        Dim dtTemp As Date
        dtTemp = SelfNowDateTime()
        If m_oActiveUserData.m_dtLastTime <> dtTemp Then
            m_oActiveUserData.m_dtLastTime = dtTemp
            
            #If ACTIVEUSER_USE_MEMORY = 0 Then
                Dim szSql As String
                Dim oDb As New RTConnection
                oDb.ConnectionString = GetConnectionStr(cszSystemMan)
                szSql = "UPDATE active_user_info SET last_time='" & ToDBDateTime(dtTemp) & "' WHERE " _
                & " login_id=" & m_oActiveUserData.m_lLoginEventID
                
                oDb.Execute szSql
            #End If
        End If
    End If
End Sub

'得到用户的名称
Public Property Get UserName() As String
    Me.AssertActiveUserValid
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    szSql = "SELECT user_name FROM user_info WHERE " _
    & " user_id='" & m_oActiveUserData.m_szUserID & "'"
    Set rsTemp = oDb.Execute(szSql)
    UserName = FormatDbValue(rsTemp!user_name)
End Property

'得到用户所属的单位代码
Public Property Get UserUnitID() As String
    Me.AssertActiveUserValid
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    szSql = "SELECT unit_id FROM user_info WHERE " _
    & " user_id='" & m_oActiveUserData.m_szUserID & "'"
    Set rsTemp = oDb.Execute(szSql)
    UserUnitID = FormatDbValue(rsTemp!unit_id)
End Property

'得到用户所属售票站

Public Property Get SellstationID() As String
    Me.AssertActiveUserValid
    SellstationID = m_oActiveUserData.m_szSellStationID
End Property



'得到用户所属的单位名称
Public Property Get UserUnitName() As String
    Me.AssertActiveUserValid
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)

    szSql = "SELECT unit_short_name FROM user_info,connect_unit_info WHERE " _
    & " user_id='" & m_oActiveUserData.m_szUserID & "' AND " _
    & " user_info.unit_id=connect_unit_info.unit_id "
    Set rsTemp = oDb.Execute(szSql)
    UserUnitName = FormatDbValue(rsTemp!unit_short_name)
End Property

'判断活动用户对象是有效的
Public Sub AssertActiveUserValid()
    Dim ausTemp As EActiveUserStatus
    Dim bIsInternetUser As Boolean
    
    If m_oActiveUserData Is Nothing Then
        ShowError ERR_ActiveUserNotLogin
    Else
        bIsInternetUser = m_oActiveUserData.m_bInternetUser
    End If
    
    If bIsInternetUser Then
        ausTemp = ActiveUserStatus
        If ausTemp = ST_AUNotLogin Then
            ShowError ERR_ActiveUserNotLogin
        ElseIf ausTemp = ST_AUTimeOut Then
            ShowError ERR_ActiveUserTimeOut
        End If
        IAmLive
    End If
End Sub

'返回此活动用户是否是Internet用户
Public Property Get IsInternetUser() As Boolean
    Me.AssertActiveUserValid
    IsInternetUser = m_oActiveUserData.m_bInternetUser
End Property


