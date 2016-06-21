VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CheckSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E81E8B001C5"
'
Option Base 0
Option Explicit



Public Type TSplitItemInfo     '拆算费用项结构
    SplitItemID As String    '折算项代码
    SplitItemName As String    '折算项名称
    SplitStatus As ESplitItemUseStatus   '使用状态 0-使用,1-未用
    SplitType As ESplitItemType   '折算类型 0-拆给对方公司,1-拆给站方,2-留给本公司
    AllowModify As ESplitItemAllowModify  '是否允许修改 0-不允许修改,1-允许修改如果允许修改,则在结算过程中,可以修改该结算项的值
End Type

Public Type TFinChargeItemInfo     '费用名结构
    ProtocolID As String    '协议号
    ProtocolName As String '协议名
    SplitItemID As String    '折算项代码
    SplitItemName As String    '折算项名称
    SplitStatus As ESplitItemUseStatus '使用状态 0-使用,1-未用
    SplitType As ESplitItemType '折算类型 0-拆给对方公司,1-拆给站方,2-留给本公司
    AllowModify As ESplitItemAllowModify '是否允许修改 0-不允许修改,1-允许修改
    FormularName As String  '公式名称LimitCharge as doublic  '底限费用
    FormulaComment As String '公式描述
    LimitCharge As Double  '底限费用
    UpCharge As Double '上限费用
End Type
Public Type TSettleSheet      '结算单基本信息
    SettleSheetID As String       '结算单号
    CheckSheetCount As Integer  '路单数
    TotalTicketPrice  As Double  '总票款
    TotalQuantity   As Double '总人数
    SettleOtherCompanyPrice  As Double   '结算给对方公司金额
    SettleStationPrice   As Double '结给站方金额
    SettleLocalCompanyPrice  As Double  '结给本公司金额
    ObjectID   As String  '对象代码
    
    SettleObject As ESettleObjectType '结算对象
    TransportCompanyID As String   '参运公司代码
    TransportCompanyName As String  '参运公司名称
    Settler  As String  '拆帐人
    Checker As String  '复核人
    SettleDate As Date   '拆帐日期
    SettleStartDate  As Date  '结算开始时间
    SettleEndDate As Date  '结算结束时间
    UnitName As String  '结算单位名称
    Status  As ESettleSheetStatus '结算单状态
    Annotation As String  '注释
    
    
    
    '*****新增的,导致不兼容
    ObjectName As String '对象名称
    LuggageSettleIDs As String '行包结算单号,以逗号分隔
    LuggageTotalBaseCarriage As Double '行包结算总运费
    LuggageTotalSettlePrice As Double '行包结算总结算款
    LuggageProtocolName As String '行包结算协议
    
    
    
    '2005-07-15号陈峰加
    RouteID As String '线路代码
    RouteName As String '线路名称
        
End Type

Public Type TCheckSheetStation       '路单信息
    CheckSheetID As String           '路单代码
    SellstationID As String          '上车站代码
    StationID As String     '站点代码
    StationName As String  '站点名称
    PriceIdentify As Integer   '票价识别序号
    TicketType As Integer  '票种
    SeatTypeID As String   '座位类型代码
    Status As Integer  '改并状态
    BusDate As Date    '日期
    BusID As String    '车次代码
    BusSerialNo As Integer '车次序号
    StationSerial As Integer   '站点序号
    Quantity As Long   '人数
    Mileage As Double  '里程数
    TicketPrice As Double    '票价
    BaseCarriage As Double '基本运价
    PriceItem(1 To 15) As Double    '票价项
End Type

Public Type TVehilceSettle   '车辆结算明细类数组
    SettlementSheetID As String      '结算单号
    VehicleID As String  '车辆代码
    VehicleTypeCode  As String      '车型代码
    LicenseTagNo    As String    '车牌号
    ProtocolID   As String         '协议号
    ProtocolName      As String   '协议名称
    SettlePrice As Double '应结票款
    SettleStationPrice As Double '结给车站票款
    PassengerNumber As Double '人数
    Mileage As Double '总人公里
    SplitItem(1 To 20)   As Double        ' 拆算项
End Type

Public Type TBusSettle   '车次结算明细类数组
    SettlementSheetID As String      '结算单号
    BusID As String
    VehicleID As String  '车辆代码
    VehicleTypeCode  As String      '车型代码
    LicenseTagNo    As String    '车牌号
    CompanyID  As String  '参运公司代码
    CompanyName As String    '公司简称
    ProtocolID   As String         '协议号
    ProtocolName      As String   '协议名称
    SettlePrice As Double '应结票款
    SettleStationPrice As Double '结给车站票款
    PassengerNumber As Double '人数
    Mileage As Double '总人公里
    SplitItem(1 To 20)   As Double        ' 拆算项
End Type

Public Type TCompnaySettle    '公司结算明细类数组
    SettlementSheetID  As String    '结算单号
    CompanyID  As String  '参运公司代码
    CompanyName As String    '公司简称
    ProtocolID      As String      '协议号
    ProtocolName  As String      '协议名称
    SettlePrice As Double '应结票款
    SettleStationPrice As Double '结给车站票款
    PassengerNumber As Double '人数
    Mileage As Double '总人公里
    SplitItem(1 To 20)      As Double     ' 拆算项
End Type
Public Type TSettleExtraInfo '结算手工补票信息
    SettlementSheetID  As String    '结算单号
    PassengerNumber As Double '人数
    TotalTicketPrice As Double '总票款
    Ratio As Double '劳务费率
    ServicePrice As Double '劳务费
    SettleOutPrice As Double '拆出金额
    
End Type



Public Type TSettleSheetStation     '结算站点类数组
    SettlementSheetID   As String   '结算单代码
    SellSationID          As String '上车站代码
    SellStationName As String       '上车站名称
    RouteID             As String   '线路代码
    RouteName As String             '线路名称
    StationID As String             '站点代码
    StationName    As String        '站点名称
    TicketType      As String       '票种代码
    TicketTypeName  As String       '票种名称
    
    VehicleTypeCode As String       '车型代码
    VehicleTypeName As String       '车型名称
    AreaRatio As Double             '区域费率
'    LicenseTagNo As String          '车牌号
'    TransportCompanyID As String    '参运公司代码
'    TransportCompanyName As String  '参运公司名称
    
    Quantity          As Double     '站点人数
End Type
Public Type TCompanySettlePrice '公司结算价信息
    CompanyID As String     '参运公司代码
    CompanyName As String  '参运公司名称
    VehicleTypeCode As String   '车型代码
    VehicleTypeName As String  '车型名称
    RouteID   As String      '线路代码
    RouteName As String   '线路名称
    SellstationID As String    '上车站代码
    SellStationName As String  '上车站名称
    StationID  As String         '站点代码
    StationName As String  '站点名称
    Mileage   As Double    '里程数
    PassCharge    As Double      '通行费
    SettlefullPrice As Double  '结算全价
    SettleHalfPrice As Double   '结算半价
    HalveFullPrice As Double '平分结算全价
    HalveHalfPrice As Double '平分结算半价
    ServiceFullPrice As Double '劳务费结算全价
    ServiceHalfPrice As Double '劳务费结算半价
    SpringFullPrice As Double '春运费结算全价
    SpringHalfPrice As Double '春运费结算半价
    
    Annotation    As String  '计算说明
End Type

'Public Type TCompanySettlePriceEX '公司结算价信息
'    CompanyID As String     '参运公司代码
'    CompanyName As String  '参运公司名称
'    VehicleTypeCode As String   '车型代码
'    VehicleTypeName As String  '车型名称
'    RouteID   As String      '线路代码
'    RouteName As String   '线路名称
'    SellstationID As String    '上车站代码
'    SellStationName As String  '上车站名称
'    StationID  As String         '站点代码
'    StationName As String  '站点名称
'    Mileage   As Double    '里程数
'    PassCharge    As Double      '通行费
'    SettlefullPrice As Double  '结算全价
'    SettleHalfPrice As Double   '结算半价
'    HalveFullPrice As Double '平分结算全价
'    HalveHalfPrice As Double '平分结算半价
'    ServiceFullPrice As Double '劳务费结算全价
'    ServiceHalfPrice As Double '劳务费结算半价
'
'    Annotation    As String  '计算说明
'End Type

Public Type TVehcileSettlePrice '车辆结算价信息
    VehcileID As String    '车辆代码
    LicenseTagNo As String  '车牌号
    RouteID    As String     '线路代码
    RouteName  As String     '线路名称
    SellstationID As String    '上车站代码
    SellStationName As String   '上车站名称
    StationID   As String        '站点代码
    StationName As String  '站点名称
    Mileage  As Double    '里程数
    PassCharge    As Double     '通行费
    SettlefullPrice As Double '结算全价
    SettleHalfPrice  As Double '结算半价
    HalveFullPrice As Double '平分结算全价
    HalveHalfPrice As Double '平分结算半价
    ServiceFullPrice As Double '劳务费结算全价
    ServiceHalfPrice As Double '劳务费结算半价
    SpringFullPrice As Double '春运费结算全价
    SpringHalfPrice As Double '春运费结算半价
    Annotation  As String   '计算说明
End Type


Public Type TBusSettlePrice '车次结算价信息
    BusID As String    '车次代码
    RouteID  As String     '线路代码
    RouteName  As String     '线路名称
    TransportCompanyID  As String     '参运公司代码
    TransportCompanyName  As String     '参运公司名称
    SellstationID As String    '上车站代码
    SellStationName As String   '上车站名称
    StationID   As String        '站点代码
    StationName As String  '站点名称
    Mileage  As Double    '里程数
    PassCharge    As Double     '通行费
    SettlefullPrice As Double '结算全价
    SettleHalfPrice  As Double '结算半价
    HalveFullPrice As Double '平分结算全价
    HalveHalfPrice As Double '平分结算半价
    ServiceFullPrice As Double '劳务费结算全价
    ServiceHalfPrice As Double '劳务费结算半价
    SpringFullPrice As Double '春运费结算全价
    SpringHalfPrice As Double '春运费结算半价
    Annotation  As String   '计算说明
End Type


Public Type TSplitResult '结算结果
    SettleSheetInfo As TSettleSheet
    SheetStationInfo() As TSettleSheetStation
    CompanyInfo() As TCompnaySettle
    VehicleInfo() As TVehilceSettle
    BusInfo() As TBusSettle
    SettleExtraInfo() As TSettleExtraInfo '结算手工补票信息
    
End Type



Public Enum ESplitItemUseStatus
    CS_SplitItemUse = 1 '使用
    CS_SplitItemNotUse = 0 '未用
End Enum

Public Enum ESplitItemAllowModify
    CS_SplitItemAllowModify = 1 '允许修改
    CS_SplitItemNotAllowModify = 0 '不允许修改
End Enum

Public Enum ESplitItemType
    CS_SplitOtherCompany = 0 '拆给对方公司
    CS_SplitStation = 1 '拆给站方
    CS_SplitLocalCompany = 2 '留给本公司
End Enum

Public Enum ECheckSheetValidMark
    CS_CheckSheetValid = 1 '正常
    CS_CheckSheetInvalid = 0 '作废
End Enum

Public Enum ECheckSheetSettleStatus
    CS_CheckSheetSettled = 1 '已经结算
    CS_CheckSheetNotSettled = 0 '未结算
    CS_CheckSheetSettling = 2 '正在结算
    
End Enum

Public Enum EQueryNegativeType
    CS_QueryAll = -1 '全部
    CS_QueryNegative = 0 '应结款为负
    CS_QueryNotNegative = 1 '应结款为正
    
End Enum


'0正常;1作废;2已汇款;3扣款未结清;4扣款已结清
Public Enum ESettleSheetStatus
    CS_SettleSheetValid = 0 '未结算
    CS_SettleSheetInvalid = 1 '作废
    CS_SettleSheetSettled = 2 '已结算
    CS_SettleSheetNotInvalid = 3 '未作废
'    CS_SettleSheetNegativeHasPayed = 4 '扣款已结清
'
End Enum

Public Enum ESettleObjectType
    CS_SettleByBus = 0   '车次
    CS_SettleByVehicle = 1 '车辆
    CS_SettleByTransportCompany = 2 '公司
    CS_SettleByOwner = 3 '车主
    CS_SettleBySplitCompany = 4 '拆账公司
End Enum



Public Enum EErrSettle
    '错误枚举
    ERR_ChargeItem = ERR_Split + 100 '28100 收费项
    ERR_Formular = ERR_Split + 150 '28150 公式管理
    ERR_CheckSheet = ERR_Split + 200 '28200 路单管理
    ERR_CompanySettlePrice = ERR_Split + 250 '28250 公司结算价管理
    ERR_HalveCompany = ERR_Split + 300 '28300 公司加总平分管理
    ERR_Report = ERR_Split + 350 '28350 查询及报表管理
    ERR_SettleSheet = ERR_Split + 400 '28400 结算单管理
    ERR_Protocol = ERR_Split + 450 '28450 协议管理
    ERR_SplitItem = ERR_Split + 500 '28500 拆算项管理
    ERR_VehicleSettlePrice = ERR_Split + 550 '28550 车辆结算价管理
    ERR_BusSettlePrice = ERR_Split + 600 '28600 车次结算价管理
    
    
    
    ERR_ProtocolExistNotAdd = ERR_Split + 21  ' 28021 协议已经存在不能新增
    ERR_ProtocolNotExist = ERR_Split + 22     ' 28022 协议不存在
    ERR_ProtocolUse = ERR_Split + 25          ' 28025 协议已经使用
    ERR_ChargeItemExitNotAdd = ERR_Split + 23 ' 28023 费用项已经存在不能新增
    ERR_ChargeItemNotExit = ERR_Split + 24    ' 28024 费用项不存在
    ERR_FormularExistNotAdd = ERR_Split + 26  ' 28026 公式已经存在不能新增
    ERR_FormularNotExist = ERR_Split + 27     ' 28027 公式不存在
    ERR_CompanySettlePriceExistNotAdd = ERR_Split + 28  ' 28028 公司结算价信息已经存在不能新增
    ERR_CompanySettlePriceNotExist = ERR_Split + 29     ' 28029 公司结算价信息不存在
    ERR_VehicleSettlePriceExistNotAdd = ERR_Split + 30  ' 28030 车辆结算价信息已经存在不能新增
    ERR_VehicleSettlePriceNotExist = ERR_Split + 31     ' 28031 车辆结算价信息不存在
    ERR_HalveCompanyExistNotAdd = ERR_Split + 32  ' 28032 加总平分公司已经存在不能新增
    ERR_HalveCompanyNotExist = ERR_Split + 33     ' 28033 加总平分公司不存在
    ERR_SplitInputParam = ERR_Split + 34 '28034 结算调用参数不正确
    ERR_PreviewSplitInputParam = ERR_Split + 35 '28035 结算预览调用参数不正确
    ERR_SomeSheetNotExist = ERR_Split + 36 '28036 有部分要结算的路单不存在
    ERR_SomeSheetHasCancelORSettle = ERR_Split + 37 '28037 有部分要结算的路单已废或已结算
    ERR_SplitSheetMustSameCompany = ERR_Split + 38 '28038 按公司拆时所有的路单必须是同一个公司
    ERR_SplitSheetMustSameVehicle = ERR_Split + 39 '28039 按车辆拆时所有的路单必须是同一辆车
    ERR_CompanyNotSetProtocol = ERR_Split + 40 '28040 公司协议未设置
    ERR_CancelSettleSheet = ERR_Split + 41 '28041 作废结算单
    ERR_VehicleNotSetProtocol = ERR_Split + 42 '28042 车辆协议未设置
    ERR_NoVehicle = ERR_Split + 43  '28043 未传入需设置协议的车辆
    ERR_SplitNotAvailable = ERR_Split + 44 '28044 拆算对象不可用
    ERR_SettleSheetHasExist = ERR_Split + 45 '28045 结算单已存在
    ERR_SettleSheetNotExist = ERR_Split + 46 '28046 结算单不存在
    ERR_CheckSheetNotAvailable = ERR_Split + 47 '28047 路单对象不可用
    ERR_CompanySettlePriceNotAvailable = ERR_Split + 48 '28048 公司结算价对象不可用
    ERR_FormularNotAvailable = ERR_Split + 49 '28049 公式对象不可用
    ERR_HalveCompanyNotAvailable = ERR_Split + 50 '28050 公司加总平分对象不可用
    ERR_ProtocloNotAvailable = ERR_Split + 51 '28051 协议对象不可用
    ERR_ReportNotAvailable = ERR_Split + 52 '28052 查询对象不可用
    ERR_SettleSheetNotAvailable = ERR_Split + 53 '28053 结算单对象不可用
    ERR_SplitItemAvailable = ERR_Split + 54 '结算项对象不可用
    ERR_VehicleSettlePriceAvailable = ERR_Split + 55 '28055 车辆结算价对象不可用
    ERR_ManualSettleCannotUseTicketPriceStr = ERR_Split + 56 '如果使用手工结算,不能用到带[票价]字符的项目
    ERR_CompanyNotExist = ERR_Split + 57 '参运公司不存在
    ERR_VehicleNotExist = ERR_Split + 58 '车辆不存在
    ERR_NotAllowSplitAboveFactQuantity = ERR_Split + 59 '不允许实结人数大于总的路单人数
    ERR_VehicleFixFeeHasExist = ERR_Split + 60 '车辆固定费用已存在
    ERR_VehicleFixFeeNotExist = ERR_Split + 61 '车辆固定费用不存在
    ERR_VehicleFixFeeHasDesCannotEdit = ERR_Split + 62 '费用已经扣过,不就允许修改或删除
    ERR_SettleSheetHasCanceled = ERR_Split + 63 '结算单已作废,不允许作废
    ERR_SettleSheetHasSettled = ERR_Split + 64 '结算单已汇款,不允许作废
    ERR_NotAllowRemitHasCancel = ERR_Split + 65 '结算单已作废,则不允许再次汇款
    ERR_NotAllowRemitHasRemitted = ERR_Split + 66 '结算单已汇款,则不允许再次汇款
    ERR_OnlyAllowUpdateOneSettleSheet = ERR_Split + 67 '同一时间只允许修改一张结算单的费用
    ERR_NotAllowUpdateHasCancel = ERR_Split + 68 '已作废,则不允许修改费用
    ERR_NotAllowUpdateHasRemitted = ERR_Split + 69 '已汇款,则不允许修改费用
    ERR_BusFixFeeHasExist = ERR_Split + 70 '车次固定费用已存在
    ERR_BusFixFeeHasDesCannotEdit = ERR_Split + 71 '车次固定费用已经扣过,不就允许修改或删除
    ERR_BusFixFeeNotExist = ERR_Split + 72 '车次固定费用不存在
    ERR_NotAllowCancelRemitHasCancel = ERR_Split + 73 '已作废,则不允许作废汇款
    ERR_NotAllowCancelRemitHasNotRemitted = ERR_Split + 74 '未汇款,则不允许作废汇款
    ERR_NoFixFeeItem = ERR_Split + 75 '系统参数中未设置固定费用项
    ERR_BusNotSetProtocol = ERR_Split + 76 '28076 车次协议未设置
    
End Enum


Public Enum ERightSettle
    
    '=================================================================================================
    '此权限没用
    RIGHT_ChargeItem = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 2 '28662 结算项管理
    RIGHT_StatSettleSheetCompany = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 14 '汇总公司拆算明细表
    RIGHT_StatSettleSheetVehicle = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 15 '汇总车辆拆算明细表
    '=================================================================================================
    
    
    RIGHT_Protocol = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 1 '28656 编辑协议
    RIGHT_Formular = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 3 '编辑公式
    RIGHT_CheckSheet = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 4 '路单管理
    RIGHT_CompanySettlePrice = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 5 '编辑公司结算价
    RIGHT_HalveCompany = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 6 '编辑公司加总平分
    RIGHT_Report = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 7 '综合查询
    RIGHT_SettleSheet = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 8 '结算单管理
    RIGHT_Split = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 9 '路单结算
    RIGHT_SplitItem = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 10 '结算项管理
    RIGHT_VehicleSettlePrice = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 11 '编辑车辆结算价
    RIGHT_SplitCheckSheet = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 12 '路单结算
    RIGHT_CancelSettleSheet = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 13 '作废结算单
    RIGHT_SetVehicleProtocol = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 16 '设置车辆协议
    RIGHT_SetCompanyProtocol = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 17 '设置公司协议
    RIGHT_PrintSettleSheet = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 18 '打印结算单
    RIGHT_ReprintSettleSheet = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 19 '重打结算单
    RIGHT_SplitCheckSheetManual = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 20 '路单手工结算
    RIGHT_ModifySheetVehicle = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 21 '修改路单的运行车辆
    RIGHT_VehicleFixFee = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 22 '编辑车辆固定费用
    RIGHT_SetRemit = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 23 '设置汇款
    
    
    RIGHT_UpdateHasSettledItem = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 24 '修改已结算的结算项值 28794
    RIGHT_GetProtocol = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 25 '查询协议 28800
    RIGHT_GetFormula = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 26 '查询公式 28806
    RIGHT_GetCompanySettlePrice = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 27 '查询公司结算价 28812
    RIGHT_GetHalveCompany = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 28 '查询公司加总平分 28818
    RIGHT_UpdateSettleSheet = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 29 '编辑结算单 28824
    RIGHT_GetVehicleSettlePrice = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 30 '查询车辆结算价 28830
    RIGHT_GetVehicleProtocol = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 31 '查询车辆协议 28836
    RIGHT_GetCompanyProtocol = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 32 '查询公司协议 28842
    RIGHT_GetVehicleFixFee = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 33 '查询车辆固定费用 28848
    RIGHT_GetRemit = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 34 '查询汇款 28854
    
    RIGHT_SettleSheetStat = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 35 '路单结算明细 28860
    RIGHT_CompanySettleDetail = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 36 '公司结算明细 28866
    RIGHT_VehicleSettleDetail = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 37 '车辆结算明细 28872
    RIGHT_CompanySettleStat = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 38 '公司结算汇总 28878
    RIGHT_VehicleSettleStat = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 39 '车辆结算汇总 28884
    RIGHT_VehicleSettleStatByMonth = ERR_VehicleSettlePrice + cnMidRightBegin + cnMidRightStep * 40 '车辆结算汇总(按月汇总) 28890
    
    RIGHT_BusFixFee = ERR_BusSettlePrice + cnMidRightBegin + cnMidRightStep * 41 '编辑车次固定费用 28946
    RIGHT_GetBusFixFee = ERR_BusSettlePrice + cnMidRightBegin + cnMidRightStep * 42 '查询车次固定费用 28952
    RIGHT_BusSettleStat = ERR_BusSettlePrice + cnMidRightBegin + cnMidRightStep * 43 '车次结算汇总 28958
    RIGHT_BusSettleDetail = ERR_BusSettlePrice + cnMidRightBegin + cnMidRightStep * 44 '车次结算明细 28964
    RIGHT_CancelRemit = ERR_BusSettlePrice + cnMidRightBegin + cnMidRightStep * 45 '作废汇款
    
End Enum






'set this to 0 to disable debug code in this class
#Const DebugMode = 1
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3E8219CB00BB
    Private mlClassDebugID As Long
#End If

'##ModelId=3E81E8EC01F4
Private m_szCheckSheetID As String

'##ModelId=3E81E8F100FA
Private m_dyBusDate As Date

'##ModelId=3E81E8F6004E
Private m_szBusID As String

'##ModelId=3E81E8FB01D4
Private m_szBusSerialNo As String

'##ModelId=3E81E8FF02FD
Private m_szCheckGateID As String

'##ModelId=3E81E9040000
Private m_szVehicleID As String

'##ModelId=3E81E9080157
Private m_szLicenseTagNo As String

'##ModelId=3E81E90C0261
Private m_szTransportCompanyID As String

'##ModelId=3E81E9120000
Private m_szTransportCompanyName As String

'##ModelId=3E81E91600CB
Private m_szSplitCompanyID As String

'##ModelId=3E81E91903B9
Private m_szSplitCompanyName As String

'##ModelId=3E81E91F02AF
Private m_szRouteID As String

'##ModelId=3E81E9230251
Private m_szRouteName As String

'##ModelId=3E81E92702DE
Private m_szVehicleTypeCode As String

'##ModelId=3E81E92C0000
Private m_szVehicleTypeName As String

'##ModelId=3E81E93202DE
Private m_szOwnerID As String

'##ModelId=3E81E9370128
Private m_szOwnerName As String

'##ModelId=3E81E93B004E
Private m_dyBusStartTime As Date

'##ModelId=3E81E93F0167
Private m_EValidMark As ECheckSheetValidMark

'##ModelId=3E81E9430167
Private m_szChecker As String

'##ModelId=3E81E9470109
Private m_dySheetMakeDate As Date

'##ModelId=3E81E94B0261
Private m_szSettleSheetID As String

'##ModelId=3E81E94F0280
Private m_ESettleStatus As ECheckSheetSettleStatus

'##ModelId=3E81E9550186
Private m_dySettleDate As Date
Private m_oActiveUser As ActiveUser
Private m_nObjectStatus As EObjectStatus '对象状态

'##ModelId=3E81E95D035B
Public Sub Init(poAUser As ActiveUser)
    On Error GoTo InitErr

    Set m_oActiveUser = poAUser
    m_nObjectStatus = ST_NormalObj
    Exit Sub
InitErr:
    Call RaiseError(MyUnhandledError, "CheckSheet:Init Method")
End Sub

'##ModelId=3E81E962030D
Public Sub Identify(pszSheetID As String)
    On Error GoTo IdentifyErr
    m_szCheckSheetID = pszSheetID
    RefreshObject
    m_nObjectStatus = ST_NormalObj
    Exit Sub
IdentifyErr:
    Call RaiseError(MyUnhandledError, "CheckSheet:Identify Method")
End Sub
Private Sub RefreshObject()
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "SELECT s.*,v.license_tag_no,c.transport_company_name,r.route_name,t.vehicle_type_short_name,o.owner_name " _
      & " FROM " & cszTableCheckSheet & " s," & cszTableVehicle & " v, " & cszTableCompany & " c," & cszTableRouteInfo & " r," & cszTableVehicleTypeCode & " t," & cszTableOwnerInfo & " o " _
      & " WHERE s.transport_company_id=c.transport_company_id AND s.vehicle_id=v.vehicle_id AND s.route_id =r.route_id " _
      & " AND v.owner_id=o.owner_id AND v.vehicle_type_code=t.vehicle_type_code AND check_sheet_id = " & TransFieldValueToString(m_szCheckSheetID)
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then RaiseError ERR_CheckSheetNotAvailable
    m_szCheckSheetID = FormatDbValue(rsTemp!check_sheet_id)
    m_dyBusDate = FormatDbValue(rsTemp!Bus_date)
    m_szBusID = FormatDbValue(rsTemp!bus_id)
    m_szBusSerialNo = FormatDbValue(rsTemp!bus_serial_no)
    m_szCheckGateID = FormatDbValue(rsTemp!check_gate_id)
    m_szVehicleID = FormatDbValue(rsTemp!vehicle_id)
    m_szLicenseTagNo = FormatDbValue(rsTemp!license_tag_no)
    m_szTransportCompanyID = FormatDbValue(rsTemp!transport_company_id)
    m_szTransportCompanyName = FormatDbValue(rsTemp!transport_company_name)
    m_szSplitCompanyID = FormatDbValue(rsTemp!split_company_id)
    m_szSplitCompanyName = FormatDbValue(rsTemp!transport_company_name)
    m_szRouteID = FormatDbValue(rsTemp!route_id)
    m_szRouteName = FormatDbValue(rsTemp!route_name)
    m_szVehicleTypeCode = FormatDbValue(rsTemp!vehicle_type_code)
    m_szVehicleTypeName = FormatDbValue(rsTemp!vehicle_type_short_name)
    m_szOwnerID = FormatDbValue(rsTemp!owner_id)
    m_szOwnerName = FormatDbValue(rsTemp!owner_name)
    m_dyBusStartTime = FormatDbValue(rsTemp!bus_start_time)
    m_EValidMark = FormatDbValue(rsTemp!valid_mark)
    m_szChecker = FormatDbValue(rsTemp!Checker)
    m_dySheetMakeDate = FormatDbValue(rsTemp!sheet_make_date)
    m_szSettleSheetID = FormatDbValue(rsTemp!settlement_sheet_id)
    m_ESettleStatus = FormatDbValue(rsTemp!settlement_status)
End Sub

'验证路单有效性
Public Function CheckSheetAvailable(szCheckSheetID As String) As ADODB.Recordset
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "SELECT * FROM " & cszTableCheckSheet & " WHERE check_sheet_id=" & TransFieldValueToString(szCheckSheetID)
    Set rsTemp = oDB.Execute(szSql)
    Set CheckSheetAvailable = rsTemp
    Set rsTemp = Nothing
End Function
'##ModelId=3E81E967002E
Public Function GetStationInfo(Optional pszStationID As String = "") As TCheckSheetStation()
    On Error GoTo GetStationInfoErr
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset, i As Integer, j As Integer
    Dim TTemp() As TCheckSheetStation
    Dim szSqlTemp As String
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If pszStationID <> "" Then
        szSqlTemp = " AND f.station_id "
    End If
    szSql = "SELECT * FROM " & cszTableSheetStation & " f " _
        & " WHERE check_sheet_id = " & TransFieldValueToString(m_szCheckSheetID) & szSqlTemp
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    For i = 1 To rsTemp.RecordCount
        TTemp(i).CheckSheetID = FormatDbValue(rsTemp!check_sheet_id)
        TTemp(i).SellstationID = FormatDbValue(rsTemp!sell_station_id)
        TTemp(i).StationID = FormatDbValue(rsTemp!station_id)
        TTemp(i).PriceIdentify = FormatDbValue(rsTemp!price_identify)
        TTemp(i).TicketType = FormatDbValue(rsTemp!ticket_type)
        TTemp(i).SeatTypeID = FormatDbValue(rsTemp!seat_type_id)
        TTemp(i).Status = FormatDbValue(rsTemp!Status)
        TTemp(i).BusDate = FormatDbValue(rsTemp!Bus_date)
        TTemp(i).BusID = FormatDbValue(rsTemp!bus_id)
        TTemp(i).BusSerialNo = FormatDbValue(rsTemp!bus_serial_no)
        TTemp(i).StationSerial = FormatDbValue(rsTemp!station_serial)
        TTemp(i).StationName = FormatDbValue(rsTemp!station_name)
        TTemp(i).Quantity = FormatDbValue(rsTemp!Quantity)
        TTemp(i).Mileage = FormatDbValue(rsTemp!Mileage)
        TTemp(i).TicketPrice = FormatDbValue(rsTemp!ticket_price)
        TTemp(i).BaseCarriage = FormatDbValue(rsTemp!base_carriage)
        For j = 1 To 15
            TTemp(i).PriceItem(j) = FormatDbValue(rsTemp(("Price_item_" & Trim(Str(j)))))
        Next j
    Next i
    Exit Function
GetStationInfoErr:
    Call RaiseError(MyUnhandledError, "CheckSheet:GetStationInfo Method")
End Function

'##ModelId=3E8219CD02CE
Private Sub Class_Initialize()
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
'        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'##ModelId=3E8219CD036B
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

#If DebugMode Then
    '##ModelId=3E8225DA0399
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

'##ModelId=3E822938006D
Public Property Get SettleDate() As Date
    SettleDate = m_dySettleDate
End Property

'##ModelId=3E822937032C
Public Property Get SettleStatus() As ECheckSheetSettleStatus
    SettleStatus = m_ESettleStatus
End Property

'##ModelId=3E82293701D4
Public Property Get SettleSheetID() As String
    SettleSheetID = m_szSettleSheetID
End Property

'##ModelId=3E822937008C
Public Property Get SheetMakeDate() As Date
    SheetMakeDate = m_dySheetMakeDate
End Property

'##ModelId=3E822936032C
Public Property Get Checker() As String
    Checker = m_szChecker
End Property

'##ModelId=3E8229360213
Public Property Get ValidMark() As ECheckSheetValidMark
    ValidMark = m_EValidMark
End Property

'##ModelId=3E82293600EA
Public Property Get BusStartTime() As Date
    BusStartTime = m_dyBusStartTime
End Property

'##ModelId=3E82293503C8
Public Property Get OwnerName() As String
    OwnerName = m_szOwnerName
End Property

'##ModelId=3E822935029F
Public Property Get OwnerID() As String
    OwnerID = m_szOwnerID
End Property

'##ModelId=3E82293501B5
Public Property Get VehicleTypeName() As String
    VehicleTypeName = m_szVehicleTypeName
End Property

'##ModelId=3E82293500CB
Public Property Get VehicleTypeCode() As String
    VehicleTypeCode = m_szVehicleTypeCode
End Property

'##ModelId=3E82293403D8
Public Property Get RouteName() As String
    RouteName = m_szRouteName
End Property

'##ModelId=3E822934030D
Public Property Get RouteID() As String
    RouteID = m_szRouteID
End Property

'##ModelId=3E8229340232
Public Property Get SplitCompanyName() As String
    SplitCompanyName = m_szSplitCompanyName
End Property

'##ModelId=3E8229340167
Public Property Get SplitCompanyID() As String
    SplitCompanyID = m_szSplitCompanyID
End Property

'##ModelId=3E8229330148
Public Property Get TransportCompanyName() As String
    TransportCompanyName = m_szTransportCompanyName
End Property

'##ModelId=3E822933008C
Public Property Get TransportCompanyID() As String
    TransportCompanyID = m_szTransportCompanyID
End Property

'##ModelId=3E82293203A9
Public Property Get LicenseTagNo() As String
    LicenseTagNo = m_szLicenseTagNo
End Property

'##ModelId=3E82293202EE
Public Property Get VehicleID() As String
    VehicleID = m_szVehicleID
End Property

'##ModelId=3E8229320232
Public Property Get CheckGateID() As String
    CheckGateID = m_szCheckGateID
End Property

'##ModelId=3E8229320186
Public Property Get BusSerialNo() As String
    BusSerialNo = m_szBusSerialNo
End Property

'##ModelId=3E82293200CB
Public Property Get BusID() As Single
    BusID = m_szBusID
End Property

'##ModelId=3E822932001F
Public Property Get BusDate() As Date
    BusDate = m_dyBusDate
End Property


'##ModelId=3E822931034B
Public Property Get CheckSheetID() As String
    CheckSheetID = m_szCheckSheetID
End Property






