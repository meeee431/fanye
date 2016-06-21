Attribute VB_Name = "mdlDefine"
Option Explicit

'模块定义
'这里的定义是从原先的中间层定义复制过来的


'车票种类
Public Enum ETicketType
    TP_FullPrice = 1 '全票
    TP_HalfPrice = 2 ' 半票
    TP_FreeTicket = 3 '免票
    TP_PreferentialTicket1 = 4 '优惠票1
    TP_PreferentialTicket2 = 5 '优惠票2
    TP_PreferentialTicket3 = 6 '优惠票3
End Enum


Public Enum ETicketTypeValid
    TP_TicketTypeValid = 1 '可用
    TP_TicketTypeNotValid = 0 '不可用
    TP_TicketTypeAll = 2 '所有票种
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
'车次类型
Public Enum EBusType
    TP_RegularBus = 0 '固定车次
    TP_ScrollBus = 1 ' 流水车次（滚动车次）
End Enum

Public Type TBuyTicketInfo
    nTicketType As ETicketType
    szTicketNo As String
    szSeatNo As String '空表示系统自已给座位号，'ST'表示站票
    szReserved As String
    szSeatTypeID As String  '座位类型代码
    szSeatTypeName As String '座位类型名称
    
End Type



'此定义售票传递参数
Public Type TSellTicketParam
    BuyTicketInfo() As TBuyTicketInfo
    pasgSellTicketPrice() As Single
    aszOrgTicket() As String    '改签用
    aszChangeSheetID() As String   '改签用
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


Public Type TTicketType
    nTicketTypeID As Integer
    szTicketTypeName As String
    nTicketTypeValid As Integer
    szAnnotation As String
End Type

'以下是公用的数据类型
Public Type TBusOrderCount
    szStatioinID As String
    dbCount As Double
End Type

'得到车次信息
Public Function GetBusRs(pdyBusDate As Date, pszStationID As String) As Recordset
    
End Function

''得到所有站点
'Public Function GetAllStationRs() As Recordset
'
'End Function

'得到所有票种
Public Function GetAllTicketType(Optional pnTicketTypeID As Integer) As TTicketType()
    Dim atTemp() As TTicketType
    ReDim atTemp(1 To 6)
    
    
    atTemp(1).nTicketTypeID = 1
    atTemp(1).szTicketTypeName = "全票"
    atTemp(1).nTicketTypeValid = 1
    atTemp(1).szAnnotation = ""

    atTemp(2).nTicketTypeID = 2
    atTemp(2).szTicketTypeName = "半票"
    atTemp(2).nTicketTypeValid = 1
    atTemp(2).szAnnotation = ""
    
    atTemp(3).nTicketTypeID = 3
    atTemp(3).szTicketTypeName = "免票"
    atTemp(3).nTicketTypeValid = 0
    atTemp(3).szAnnotation = ""
    
    
    atTemp(4).nTicketTypeID = 4
    atTemp(4).szTicketTypeName = "优惠票1"
    atTemp(4).nTicketTypeValid = 0
    atTemp(4).szAnnotation = ""
    
    
    atTemp(5).nTicketTypeID = 5
    atTemp(5).szTicketTypeName = "优惠票2"
    atTemp(5).nTicketTypeValid = 0
    atTemp(5).szAnnotation = ""
    
    
    atTemp(6).nTicketTypeID = 6
    atTemp(6).szTicketTypeName = "优惠票3"
    atTemp(6).nTicketTypeValid = 0
    atTemp(6).szAnnotation = ""
    
    
    
    GetAllTicketType = atTemp
    
    
End Function

'得到所有座位类型
Public Function GetAllSeatType() As String()
    Dim aszSeatType() As String
    ReDim aszSeatType(1 To 3, 1 To 3)
    
    aszSeatType(1, 1) = "01"
    aszSeatType(1, 2) = "普通"
    aszSeatType(1, 3) = ""
    
    aszSeatType(2, 1) = "02"
    aszSeatType(2, 2) = "卧铺"
    aszSeatType(2, 3) = ""
    
    aszSeatType(3, 1) = "03"
    aszSeatType(3, 2) = "加座"
    aszSeatType(3, 3) = ""
    
    GetAllSeatType = aszSeatType
    
    
End Function



 '得到票种个数
Public Function GetTicketTypeCount() As Integer
    GetTicketTypeCount = 6
End Function

'得到座位类型数目
Public Function GetSeatTypeCount() As Integer
    GetSeatTypeCount = 3
End Function

Public Function GetActiveUserID() As String
    GetActiveUserID = m_cszOperatorID
End Function

'Public Function GetAllStartStation() As String()
'    Dim aszTemp() As String
'    ReDim aszTemp(1 To 5, 1 To 2) As String
'    aszTemp(1, 1) = "0"
'    aszTemp(1, 2) = "客运中心"
'    aszTemp(2, 1) = "1"
'    aszTemp(2, 2) = "新城站"
'    aszTemp(3, 1) = "2"
'    aszTemp(3, 2) = "新南站"
'    aszTemp(4, 1) = "3"
'    aszTemp(4, 2) = "西站"
'    aszTemp(5, 1) = "4"
'    aszTemp(5, 2) = "东站"
'    GetAllStartStation = aszTemp
'
'End Function



