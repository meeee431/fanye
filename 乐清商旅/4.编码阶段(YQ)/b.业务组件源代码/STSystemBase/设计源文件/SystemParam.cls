VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************
'* Source File Name:SystemParam.cls
'* Project Name:STSystem.vbp
'* Engineer:陈峰
'* Data Generated:2002-08-15
'* Last Revision Date:2002-08-15
'* Brief Description:系统参数
'* Relational Document:
'**********************************************************
Option Explicit
'售票
Const cszPreSellDate = "PreSellDate" '预售天数
Const cszStopSellTime = "StopSellTime" '停售时间
Const cszCancelTicketTime = "CancelTicketTime" '废票时间
Const cszTicketNumberLen = "TicketNumberLen" '票号数字部分长
Const cszTicketPrefixLen = "TicketPrefixLen" '票号前辍部分长
Const cszReturnTicketRatio1 = "ReturnTicketRatio1" '第一退票费率
Const cszReturnTicketRatio2 = "ReturnTicketRatio2" '第二退票费率
Const cszReturnTicketRatio3 = "ReturnTicketRatio3" '第三退票费率
Const cszReturnTicketTime1 = "ReturnTicketTime1" '第一退票时间
Const cszReturnTicketTime2 = "ReturnTicketTime2" '第二退票时间
Const cszReturnTicketTime3 = "ReturnTicketTime3" '第三退票时间
Const cszChangeCharge = "ChangeCharge" '改签费用
Const cszScrollBusCanReturnTicket = "ScrollBusCanReturnTicket" '流水车次能否退票
Const cszScrollBusReturnRatio = "ScrollBusReturnRatio" '流水车次的退票费率
Const cszScrollBusReturnTime = "ScrollBusReturnTime" '流水车次退票时间段
Const cszPrintBarCode = "PrintBarCode" '车票要否打印条形码
Const cszPrintAid = "PrintAid" '车票要否打印副联
Const cszQueryTime = "QueryTime" '查询时间
Const cszIsPrintReturnChangeSheet = "IsPrintReturnChangeSheet" '是否打印退票手续费单据
Const cszRTRoundingBlanceValue1 = "RTRoundingBlanceValue1" '退票余额收取参数1
Const cszRTRoundingBlanceValue2 = "RTRoundingBlanceValue2" '退票余额收取参数2
Const cszRTRoundingBlanceValue3 = "RTRoundingBlanceValue3" '退票余额收取参数3
Const cszScrollBusLatestTime = "ScrollBusLatestTime" '滚动车次最迟上车时间
Const cszSellStationCanSellEachOther = "SellStationCanSellEachOther" '售票站之间是否可以相互售票
Const cszInternetBusShow = "InternetBusShow" '互联售票车次显示时间
Const cszExtraBusShowBefore = "ExtraBusShowBefore"  '补票车次显示前几小时
Const cszExtraBusShowAfter = "ExtraBusShowAfter"    '补票车次显示后几小时
Const cszPrintBusIDLen = "PrintBusIDLen" '车次打印长度
Const cszPrintScrollBusMode = "PrintScrollBusMode" '滚动车次发车时间打印方式

Const cszIsPrintZeroValueReturnSheet = "IsPrintZeroValueReturnSheet" '是否打印手续费为0的退票手续费
Const cszAllowScrollBusSaleForever = "AllowScrollBusSaleForever" '滚动班次是否永远可售

Const cszDiscountTicketInTicketTypePosition = "DiscountTicketInTicketTypePosition" '折扣票在票种项的位置
Const cszAllowSellScreenShow = "AllowSellScreenShow" '分屏显示
Const cszWantListNoSeatBus = "WantListNoSeatBus" '是否列出无座位车次
Const cszWantDirectSheetPrint = "WantDirectSheetPrint" '普通打印，高速打印

Const cszCardBuyTicketInterval = "CardBuyTicketInterval" '同一证件购票时间间隔(分钟)

'票价管理参数设置
Const cszBusSpeed = "BusSpeed" '车辆平均速度
Const cszNightShiftTime1 = "NightShiftTime1" '夜班费第一时间段
Const cszNightShiftTime2 = "NightShiftTime2" '夜班费第二时间段
Const cszPriceDetailKeepBit = "PriceDetailKeepBit " '票价项小数点后位数
'调度
Const cszAdditionBusPreFix = "AdditionBusPreFix" '加班车次的前缀
Const cszEndStatinCanSale = "EndStationCanSale" '是否加班车只售终点站
Const cszAllowSlitpNotPassStation = "AllowSlitpNotPassStation" '是否允许不停靠站点拆分
Const cszAllowMakeEnviromentSaleBus = "AllowMakeEnviromentSaleBus"  '是否允许生成已售票车次的环境
Const cszImmediatelyQuery = "stationImmediatelyQuery"  '以站点查询
Const cszMakeStopEnviroment = "MakeStopEnviroment" '环境中是否生成停班车次
'行包参数
Const cszLuggageIDNumberLen = "LuggageIDNumberLen" '行包单号数字部分长
Const cszLuggageIDPrefixLen = "LuggageIDPrefixLen" '行包单号前辍部分长
Const cszCheckLuggageIDNumberLen = "CheckLuggageIDNumberLen" '签发单号数字部分长
'Const cszCheckLuggageIDPrefixLen = "CheckLuggageIDPrefixLen" '签发单号前辍部分长

Const cszCheckLimitTime = "CheckLimitTime" '签发时限

Const cszAcceptLimitTime = "AcceptLimitTime" '受理时限

Const cszStopBusCanCheck = "StopBusCanCheck" '停班车可以签发

Const cszProfessionalBusPreFix = "ProfessionalBusPreFix" '专业车前缀
Const cszLuggageRefreshFre = "LuggageRefreshFre" '行包刷新频率
Const cszLuggageLoadItemNo = "LuggageLoadItemNo" '装卸费为第几个收费项
Const cszLuggageMoveItemNo = "LuggageMoveItemNo" '搬运费为第几个收费项
Const cszStopBusCanAccept = "StopBusCanAccept" '停班车可以受理

Const cszIsOneFormulaEachStation = "IsOneFormulaEachStation" '是否每站点一公式,0为所有站点相同公司,1为每站点一公式
Const cszIsRelationWithVehicleType = "IsRelationWithVehicleType" '计算公式是否与车型有关


Const cszIsPrintCheckSheet = "IsPrintCheckSheet" '是否打印签发单,0不打印,1打印
Const cszIsDispSettlePriceInCheck = "IsDispSettlePriceInCheck" '是否在签发时显示应结费用,0不显示,1显示
Const cszIsSettlePriceFromAcceptInCheck = "IsSettlePriceFromAcceptInCheck" '结算运费是不是从受理的结算运费中汇总得到,0直接在签发中选择费率,1从受理中汇总得到
Const cszIsDispSettlePriceInAccept = "IsDispSettlePriceInAccept" '是否在受理时显示应结费用,0不显示,1显示




'票款结算参数

Const cszSubTotalItem1 = "SubTotalItem1" '小计项1
Const cszBasePriceTotalItem = "BasePriceTotalItem" '基本运价合计项

Const cszTranNeedStationDetail = "TranNeedStationDetail" '运量统计需要列出的站点明细
Const cszInterProvinceStationDetail = "InterProvinceStationDetail" '省际班车需列出的站点明细
Const cszInCityStationDetail = "InCityStationDetail" '市内班车需列出的站点明细
Const cszExtramuralStationDetail = "ExtramuralStationDetail" '市外班车需列出的站点明细

'票价

Const cszAdvanceDistance1 = "AdvanceDistance1"  '春运费计算1
Const cszAdvanceDistance2 = "AdvanceDistance2"  '春运费计算2
Const cszRoadBuildChargeItem = "RoadBuildChargeItem" '公建金为第几个票价项
Const cszFarDistanceAddChargeItem = "FarDistanceAddChargeItem" '250K加成费为第几个票价项
Const cszSpringItem = "SpringItem" '春运费为第几个票价项

'检票

Const cszCheckSheetLen = "CheckSheetLen" '路单编号长度
Const cszAllowChangeRide = "AllowChangeRide" '是否允许改乘
Const cszAllowChangeRideLowerPrice = "AllowChangeRideLowerPrice" '是否允许从低票价的车次改乘到高票价的车次

Const cszBeginCheckTime = "BeginCheckTime" '开始检票时间一时间间隔
Const cszCheckTicketTime = "CheckTicketTime" '检票时间(分钟)
Const cszLatestExtraCheckTime = "LatestExtraCheckTime" '补检时间最晚开检时间(分钟)
Const cszExtraCheckTime = "ExtraCheckTime" '补检时间(分钟)

Const cszWantPrintRoadSheetID = "WantPrintRoadSheetID" '是否打印路单号
Const cszWantPrintRoadSheetTitle = "WantPrintRoadSheetTitle" '是否打印标题
Const cszRoadSheetTitle = "RoadSheetTitle" '路单标题
Const cszWantTDPrintType = "WantTDPrintType" '路单是否套打方式


Const cszAllowStartChectNotRearchTime = "AllowStartChectNotRearchTime" '允许未到开检时间直接开检


'预定
Const cszBookTime = "BookTime" '预定时间
Const cszIsBookValid = "IsBookValid" '是否上预定系统
Const cszBookNumberLen = "BookNumberLen" '预定号长度



'报班参数设置
Const cszPreviousRegistTime = "PreviousRegistTime" '预览报班时间
Const cszStartRegistTime = "StartRegistTime" '报班开始时间
Const cszEndRegistTime = "EndRegistTime" '报班结束时间
Const cszOutRegistTime = "OutRegistTime" '误班结束时间
Const cszLogoutRegistTime = "LogoutRegistTime" '自动脱班时间
Const cszRegistCardNoLen = "RegistCardNoLen" '报班卡号长度

'语音播放参数设置
Const cszRecordPath = "RecordPath" '声音文件存放路径
Const cszAllowLength = "AllowLength" '是否限制文件长度
Const cszFileLength = "FileLength" '声音文件长度
Const cszPlayTime = "PlayTime" '播放时间
Const cszPlayMode = "PlayMode" '播放方式
Const cszRefreshFre = "RefreshFre" '刷新频率
Const cszPlayNumber = "PlayNumber" '播放次数

'路单结算
Const cszExtraCsRatio = "ExtraCsRatio"                  '补票结算费率
Const cszExtraLsRatio = "ExtraLsRatio"                  '补行包票结算费率
Const cszExtraCsRatio1 = "ExtraCsRatio1"                '补票结算费率1
Const cszExtraCsRatio2 = "ExtraCsRatio2"                '补票结算费率2
Const cszExtraCsRatio3 = "ExtraCsRatio3"                '补票结算费率3
Const cszExtraLsRatio1 = "ExtraLsRatio1"                '补行包票结算费率1
Const cszExtraLsRatio2 = "ExtraLsRatio2"                '补行包票结算费率2
Const cszExtraLsRatio3 = "ExtraLsRatio3"                '补行包票结算费率3
Const cszExtraCsDisplay1 = "ExtraCsDisplay1"            '路单结算中显示补客票1
Const cszExtraCsDisplay2 = "ExtraCsDisplay2"            '路单结算中显示补客票2
Const cszExtraCsDisplay3 = "ExtraCsDisplay3"            '路单结算中显示补行包票1
Const cszWantPrintCsSplitIn = "WantPrintCsSplitIn"      '路单结算中打印拆入项
Const cszWantPrintCsSplitOut = "WantPrintCsSplitOut"    '路单结算中打印拆出项
Const cszWantCsPrintLuggage = "WantCsPrintLuggage"      '路单结算中打印行包运费
Const cszWantStatCheckInfo = "WantStatCheckInfo"        '手工结算前统计检票信息
Const cszWantModifySplitResult = "ModifySplitResult"    '是否允许修改路单结算的结果
Const cszWantLsSettleTogether = "LsSettleTogether"      '是否签发单和非旅客同行行包一起结算
Const cszWantSelectSomeCsSheet = "SelectSomeCsSheet"    '是否允许选择部分路单进行结算


Const cszAllowSplitBySomeTimes = "AllowSplitBySomeTimes" '是否允许路单分几次来结算 0,不允许,1允许
Const cszAllowSplitAboveFactQuantity = "AllowSplitAboveFactQuantity" '是否允许路单结算人数超过实际人数? 0不允许,1允许
Const cszFixFeeItem = "FixFeeItem"    '结算的固定费用项
Const cszAllowSettleTotalNegative = "AllowSettleTotalNegative" '是否将上月结算的负数算到下月 0,不允许,1允许
Const cszSettleNegativeSplitItem = "SettleNegativeSplitItem" '将上月结算的负数项放在结算项的位置


Const cszServiceItemPosition = "ServiceItemPosition" '劳务费的项目位置
Const cszIsFixFeeUpdateEachMonth = "IsFixFeeUpdateEachMonth" '是否固定费用是每个月更新的
Const cszCarriageItemPosition = "CarriageItemPosition" '运费的项目位置
Const cszSpringChargeItem = "SpringChargeItem" '春运费项目位置

'其他
Const cszKeepDate = "KeepDate"  '导出数据保留的日期
Const cszActiveUserTimerOut = "ActiveUserTimeOut" '活动用户超时的时间限制
Const cszCheckActiveUserInterval = "CheckActiveUserInterval" '检查活动用户的时间间隔

Const cszSpecialTicketTypePosition = "SpecialTicketTypePosition" '特种票所在票种代码
Const cszSpecialTicketTypeRatio = "SpecialTicketTypeRatio" '特种票百分比
Const cszHalfTicketTypeRatio = "HalfTicketTypeRatio" '半票百分比

Public Enum EErrSystemParam
    ERR_SystemParamNoActiveUser = ERR_SystemParam + ERR_NoActiveUser
    ERR_SpecialParamNotExist = ERR_SystemParam + 21
End Enum

'退票费率类型
Public Type RETURNFEES
    iReturnNum As Integer
    iReturnTime As Integer
    sgReturnRate As Single
    sgLeastMoney As Single
End Type

Public Enum ERightSystemParam
    RIGHT_SetSysParam = ERR_SystemParam + cnMidRightBegin + cnMidRightStep * 1
End Enum

'增加消息支持
'-------------------
#If SUPPORT_MSG = 1 Then
Private m_oSndMsg As New CSendMsg
#End If
'-------------------

'所有的类都有的私有类充量
Private m_oActiveUser As ActiveUser
'--------------------------->>
Public Sub Init(poAUser As ActiveUser)
    Set SelfUser = poAUser
End Sub

Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
End Property

Public Property Set SelfUser(vNewValue As ActiveUser)
        Set m_oActiveUser = vNewValue
End Property

'得到系统参数
Public Function GetParam(ByVal pszParamName As String) As TParamInfo
    AssertActiveUserValid m_oActiveUser, ERR_SystemParam
    'GetParam = SelfGetParam(pszParamName)
    
    Dim piTemp As TParamInfo
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "SELECT * FROM  " & cszSystemParam & "  WHERE parameter_name='" & FormatStr(pszParamName) & "'"
    Set rsTemp = oDb.Execute(szSql)
    
    If rsTemp.RecordCount <> 0 Then
        piTemp.szParamName = FormatDbValue(rsTemp!parameter_name)
        piTemp.szParamChineseName = FormatDbValue(rsTemp!parameter_chinese_name)
        piTemp.szAnnotation = FormatDbValue(rsTemp!Annotation)
        piTemp.szParamValue = FormatDbValue(rsTemp!parameter_value)
    Else
        ShowError ERR_SpecialParamNotExist, , "系统参数=" & pszParamName
    End If
    GetParam = piTemp
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

'修改系统参数
Public Sub ModifyParam(ppiInfo As TParamInfo, Optional ByVal pbOnlyValue As Boolean = False)
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim lRowAffect As Long
    Dim oDb As New RTConnection
    
    AssertActiveUserValid m_oActiveUser, ERR_SystemParam
    AssertHaveRight m_oActiveUser, RIGHT_SetSysParam
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    If Not pbOnlyValue Then
        szSql = "UPDATE  " & cszSystemParam & "  " _
        & "SET parameter_chinese_name='" & ppiInfo.szParamChineseName _
        & "',parameter_value='" & ppiInfo.szParamValue _
        & "',annotation='" & ppiInfo.szAnnotation _
        & "' WHERE parameter_name='" & ppiInfo.szParamName & "'"
    Else
        szSql = "UPDATE  " & cszSystemParam & "  " _
        & " SET parameter_value='" & ppiInfo.szParamValue & "' " _
        & " WHERE parameter_name='" & ppiInfo.szParamName & "'"
    End If
    
    oDb.Execute szSql, lRowAffect
    If lRowAffect = 0 Then ShowError ERR_SpecialParamNotExist, , "系统参数=" & ppiInfo.szParamName
    
    #If SUPPORT_MSG Then
    Dim aszValues(1 To 2) As String
    aszValues(1) = ppiInfo.szParamName
    aszValues(2) = ppiInfo.szParamValue
    m_oSndMsg.SendMsg NowDateTime, cszChangeParam, aszValues
    #End If
   
    WriteOperateLog m_oActiveUser, RIGHT_SetSysParam, ppiInfo.szParamName & " --> " & ppiInfo.szParamValue
    Set rsTemp = Nothing
    Set oDb = Nothing
End Sub

'特种票所在票种代码
Public Property Get SpecialTicketTypePosition() As String
    SpecialTicketTypePosition = GetParam(cszSpecialTicketTypePosition).szParamValue
End Property

'特种票百分比
Public Property Get SpecialTicketTypeRatio() As Double
    SpecialTicketTypeRatio = CSng(GetParam(cszSpecialTicketTypeRatio).szParamValue)
End Property

'半票百分比
Public Property Get HalfTicketTypeRatio() As Double
    HalfTicketTypeRatio = CSng(GetParam(cszHalfTicketTypeRatio).szParamValue)
End Property

Public Property Let HalfTicketTypeRatio(ByVal vNewValue As Double)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszHalfTicketTypeRatio
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到本站单位码
Public Property Get UnitID() As String
    UnitID = GetParam(cszLocalUnitID).szParamValue
End Property

Public Property Let UnitID(ByVal vNewValue As String)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszLocalUnitID
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到本站站码
Public Property Get StationID() As String
    StationID = GetParam(cszLocalStationID).szParamValue
End Property

Public Property Let StationID(ByVal vNewValue As String)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszLocalStationID
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到活动用户超时的分钟
Public Property Get ActiveUserTimeOut() As Integer
    ActiveUserTimeOut = CInt(GetParam(cszActiveUserTimerOut).szParamValue)
End Property

Public Property Let ActiveUserTimeOut(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszActiveUserTimerOut
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到检查活动用户的时间间隔
Public Property Get CheckActiveUserInterval() As Integer
    CheckActiveUserInterval = CInt(GetParam(cszCheckActiveUserInterval).szParamValue)
End Property

Public Property Let CheckActiveUserInterval(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszCheckActiveUserInterval
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到当前日期时间
Public Property Get NowDateTime() As Date
    NowDateTime = SelfNowDateTime()
End Property

'得到当前日期
Public Property Get NowDate() As Date
    NowDate = SelfNowDate()
End Property

'得到当前时间
Public Property Get NowTime() As Date
    NowTime = SelfNowTime()
End Property

'得到预售天数
Public Property Get PreSellDate() As Integer
    PreSellDate = CInt(GetParam(cszPreSellDate).szParamValue)
End Property

Public Property Let PreSellDate(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszPreSellDate
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到分屏显示
Public Property Get AllowSellScreenShow() As Integer
    AllowSellScreenShow = CInt(GetParam(cszAllowSellScreenShow).szParamValue)
End Property

Public Property Let AllowSellScreenShow(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszAllowSellScreenShow
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到停售时间
Public Property Get StopSellTime() As Integer
    StopSellTime = CInt(GetParam(cszStopSellTime).szParamValue)
End Property

Public Property Let StopSellTime(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszStopSellTime
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property


'得到报班卡号长度
Public Property Get RegistCardNoLen() As Integer
   RegistCardNoLen = CInt(GetParam(cszRegistCardNoLen).szParamValue)
End Property

Public Property Let RegistCardNoLen(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszRegistCardNoLen
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到自动脱班时间
Public Property Get LogoutRegistTime() As Integer
    LogoutRegistTime = CInt(GetParam(cszLogoutRegistTime).szParamValue)
End Property

Public Property Let LogoutRegistTime(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszLogoutRegistTime
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到误班结束时间
Public Property Get OutRegistTime() As Integer
    OutRegistTime = CInt(GetParam(cszOutRegistTime).szParamValue)
End Property

Public Property Let OutRegistTime(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszOutRegistTime
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

Public Property Get EndRegistTime() As Integer
    EndRegistTime = CInt(GetParam(cszEndRegistTime).szParamValue)
End Property

Public Property Let EndRegistTime(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszEndRegistTime
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到报班开始时间
Public Property Get StartRegistTime() As Integer
    StartRegistTime = CInt(GetParam(cszStartRegistTime).szParamValue)
End Property

Public Property Let StartRegistTime(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszStartRegistTime
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到预览报班时间
Public Property Get PreviousRegistTime() As Integer
   PreviousRegistTime = CInt(GetParam(cszPreviousRegistTime).szParamValue)
End Property

Public Property Let PreviousRegistTime(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszPreviousRegistTime
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到正常废票时间
Public Property Get CancelTicketTime() As Integer
    CancelTicketTime = CInt(GetParam(cszCancelTicketTime).szParamValue)
End Property

Public Property Let CancelTicketTime(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszCancelTicketTime
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到正常退票时间1
Public Property Get ReturnTicketTime1() As Integer
    ReturnTicketTime1 = CInt(GetParam(cszReturnTicketTime1).szParamValue)
End Property

Public Property Let ReturnTicketTime1(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszReturnTicketTime1
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到正常退票时间2
Public Property Get ReturnTicketTime2() As Integer
    ReturnTicketTime2 = CInt(GetParam(cszReturnTicketTime2).szParamValue)
End Property

Public Property Let ReturnTicketTime2(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszReturnTicketTime2
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到正常退票时间3
Public Property Get ReturnTicketTime3() As Integer
    ReturnTicketTime3 = CInt(GetParam(cszReturnTicketTime3).szParamValue)
End Property

Public Property Let ReturnTicketTime3(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszReturnTicketTime3
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到滚动车次退票时间段
Public Property Get ScrollBusReturnTime() As Double
        ScrollBusReturnTime = CSng(GetParam(cszScrollBusReturnTime).szParamValue)
End Property

Public Property Let ScrollBusReturnTime(ByVal vNewValue As Double)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszScrollBusReturnTime
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到正常退票费率1
Public Property Get ReturnTicketRatio1() As Double
    ReturnTicketRatio1 = CSng(GetParam(cszReturnTicketRatio1).szParamValue)
End Property

Public Property Let ReturnTicketRatio1(ByVal vNewValue As Double)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszReturnTicketRatio1
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property
'得到正常退票费率2
Public Property Get ReturnTicketRatio2() As Double
    ReturnTicketRatio2 = CSng(GetParam(cszReturnTicketRatio2).szParamValue)
End Property

Public Property Let ReturnTicketRatio2(ByVal vNewValue As Double)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszReturnTicketRatio2
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到正常退票费率3
Public Property Get ReturnTicketRatio3() As Double
    ReturnTicketRatio3 = CSng(GetParam(cszReturnTicketRatio3).szParamValue)
End Property

Public Property Let ReturnTicketRatio3(ByVal vNewValue As Double)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszReturnTicketRatio3
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到指定发车时间的退票费率
Public Function GetReturnChargeRatio(ByVal pdtBusStartupDateTime As Date) As Double
    Dim lInterval As Long
    Dim sgTemp As Double
    Dim rfTemp() As RETURNFEES
    Dim iLen, iCount As Integer
    iLen = 0
    sgTemp = 100
    rfTemp = GetReturnFees()
    iLen = ArrayLength(rfTemp)
    If iLen = 0 Then
        GetReturnChargeRatio = sgTemp
        Exit Function
    End If
    
    lInterval = DateDiff("n", SelfNowDateTime(), pdtBusStartupDateTime)
    For iCount = 1 To iLen
        If lInterval >= rfTemp(iCount).iReturnTime Then
            sgTemp = rfTemp(iCount).sgReturnRate
            Exit For
        End If
    Next iCount
    
'    If lInterval >= ReturnTicketTime1 Then
'        sgTemp = ReturnTicketRatio1
'    ElseIf lInterval >= ReturnTicketTime2 Then
'        sgTemp = ReturnTicketRatio2
'    ElseIf lInterval >= ReturnTicketTime3 Then
'        sgTemp = ReturnTicketRatio3
'    End If
    GetReturnChargeRatio = sgTemp
End Function

'得到改签费
Public Property Get ChangeCharge() As Double
    ChangeCharge = CSng(GetParam(cszChangeCharge).szParamValue)
End Property

Public Property Let ChangeCharge(ByVal vNewValue As Double)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszChangeCharge
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到开始检票时间
Public Property Get BeginCheckTime() As Integer
    BeginCheckTime = CInt(GetParam(cszBeginCheckTime).szParamValue)
End Property

Public Property Let BeginCheckTime(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszBeginCheckTime
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到检票时间
Public Property Get CheckTicketTime() As Integer
    CheckTicketTime = CInt(GetParam(cszCheckTicketTime).szParamValue)
End Property

Public Property Let CheckTicketTime(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszCheckTicketTime
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到补检时间最晚开检时间(分钟)
Public Property Get LatestExtraCheckTime() As Integer
    LatestExtraCheckTime = CInt(GetParam(cszLatestExtraCheckTime).szParamValue)
End Property

Public Property Let LatestExtraCheckTime(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszLatestExtraCheckTime
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到补检时间(分钟)
Public Property Get ExtraCheckTime() As Integer
    ExtraCheckTime = CInt(GetParam(cszExtraCheckTime).szParamValue)
End Property

Public Property Let ExtraCheckTime(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszExtraCheckTime
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到是否打印路单号
Public Property Get WantPrintRoadSheetID() As Boolean
    Dim nTemp As Integer
    nTemp = CInt(GetParam(cszWantPrintRoadSheetID).szParamValue)
    WantPrintRoadSheetID = IIf(nTemp = 0, False, True)
End Property

Public Property Let WantPrintRoadSheetID(ByVal vNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszWantPrintRoadSheetID
    piTemp.szParamValue = IIf(vNewValue, "1", "0")
    ModifyParam piTemp, True
End Property

'得到是否打印路标题
Public Property Get WantPrintRoadSheetTitle() As Boolean
    Dim nTemp As Integer
    nTemp = CInt(GetParam(cszWantPrintRoadSheetTitle).szParamValue)
    WantPrintRoadSheetTitle = IIf(nTemp = 0, False, True)
End Property

Public Property Let WantPrintRoadSheetTitle(ByVal vNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszWantPrintRoadSheetTitle
    piTemp.szParamValue = IIf(vNewValue, "1", "0")
    ModifyParam piTemp, True
End Property

'得到路单标题
Public Property Get RoadSheetTitle() As String
    RoadSheetTitle = GetParam(cszRoadSheetTitle).szParamValue
End Property

Public Property Let RoadSheetTitle(ByVal vNewValue As String)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszRoadSheetTitle
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到票号数字部分的长度
Public Property Get TicketNumberLen() As Integer
    TicketNumberLen = GetParam(cszTicketNumberLen).szParamValue
End Property

Public Property Let TicketNumberLen(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszTicketNumberLen
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到票号前辍部分的长度
Public Property Get TicketPrefixLen() As Integer
    TicketPrefixLen = GetParam(cszTicketPrefixLen).szParamValue
End Property

Public Property Let TicketPrefixLen(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszTicketPrefixLen
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'检票是否使用快速打印
Public Property Get WantTDPrintType() As Boolean
    WantTDPrintType = IIf(GetParam(cszWantTDPrintType).szParamValue = 0, False, True)
End Property

Public Property Let WantTDPrintType(ByVal vNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszWantTDPrintType
    piTemp.szParamValue = IIf(vNewValue, 1, 0)
    ModifyParam piTemp, True
End Property

'流水车次是否可以退票
Public Property Get ScrollBusCanReturnTicket() As Boolean
    ScrollBusCanReturnTicket = IIf(GetParam(cszScrollBusCanReturnTicket).szParamValue = 0, False, True)
End Property

Public Property Let ScrollBusCanReturnTicket(ByVal vNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszScrollBusCanReturnTicket
    piTemp.szParamValue = IIf(vNewValue, 1, 0)
    ModifyParam piTemp, True
End Property

'流水车次的退票费率
Public Property Get ScrollBusReturnRatio() As Double
    ScrollBusReturnRatio = CSng(GetParam(cszScrollBusReturnRatio).szParamValue)
End Property

Public Property Let ScrollBusReturnRatio(ByVal vNewValue As Double)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszScrollBusReturnRatio
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'是否打印条形码
Public Property Get PrintBarCode() As Boolean
    PrintBarCode = IIf(GetParam(cszPrintBarCode).szParamValue = 0, False, True)
End Property

Public Property Let PrintBarCode(ByVal vNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszPrintBarCode
    piTemp.szParamValue = IIf(vNewValue, 1, 0)
    ModifyParam piTemp, True
End Property

'是否打印副联
Public Property Get PrintAid() As Boolean
    PrintAid = IIf(GetParam(cszPrintAid).szParamValue = 0, False, True)
End Property

Public Property Let PrintAid(ByVal vNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszPrintAid
    piTemp.szParamValue = IIf(vNewValue, 1, 0)
    ModifyParam piTemp, True
End Property

'查询时间
Public Property Get QueryTime() As Long
    QueryTime = CLng(GetParam(cszQueryTime).szParamValue)
End Property

Public Property Let QueryTime(ByVal vNewValue As Long)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszQueryTime
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'路单号长度
Public Property Get CheckSheetLen() As Integer
    CheckSheetLen = CInt(GetParam(cszCheckSheetLen).szParamValue)
End Property

Public Property Let CheckSheetLen(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszCheckSheetLen
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'是否允许改乘
Public Property Get AllowChangeRide() As Boolean
    AllowChangeRide = IIf(GetParam(cszAllowChangeRide).szParamValue = 0, False, True)
End Property

Public Property Let AllowChangeRide(ByVal vNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszAllowChangeRide
    piTemp.szParamValue = IIf(vNewValue, 1, 0)
    ModifyParam piTemp, True
End Property

'是否允许票价低于当检车次票价的车票改乘
Public Property Get AllowChangeRideLowerPrice() As Boolean
    AllowChangeRideLowerPrice = IIf(GetParam(cszAllowChangeRideLowerPrice).szParamValue = 0, False, True)
End Property

Public Property Let AllowChangeRideLowerPrice(ByVal vNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszAllowChangeRideLowerPrice
    piTemp.szParamValue = IIf(vNewValue, 1, 0)
    ModifyParam piTemp, True
End Property

'*********语音播放系统参数
'得到录音文件存放路径
Public Property Get RecordPath() As String
    RecordPath = GetParam(cszRecordPath).szParamValue
End Property

Public Property Let RecordPath(ByVal vNewValue As String)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszRecordPath
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到是否限制声音文件长度
Public Property Get IsLimitLength() As Boolean
    Dim nTemp As Integer
    nTemp = CInt(GetParam(cszAllowLength).szParamValue)
    IsLimitLength = IIf(nTemp = 0, False, True)
End Property

Public Property Let IsLimitLength(ByVal vNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszAllowLength
    piTemp.szParamValue = IIf(vNewValue, "1", "0")
    ModifyParam piTemp, True
End Property

'得到文件长度
Public Property Get FileLength() As String
    FileLength = GetParam(cszFileLength).szParamValue
End Property

Public Property Let FileLength(ByVal vNewValue As String)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszFileLength
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到播放时间(分钟)
Public Property Get PlayTime() As Integer
    PlayTime = CInt(GetParam(cszPlayTime).szParamValue)
End Property

Public Property Let PlayTime(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszPlayTime
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到播放方式
Public Property Get PlayMode() As Integer
    PlayMode = CInt(GetParam(cszPlayMode).szParamValue)
End Property

Public Property Let PlayMode(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszPlayMode
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到刷新频率
Public Property Get RefrshFre() As Integer
    RefrshFre = CInt(GetParam(cszRefreshFre).szParamValue)
End Property

Public Property Let RefrshFre(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszRefreshFre
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到播放次数
Public Property Get PlayNumber() As Integer
    PlayNumber = CInt(GetParam(cszPlayNumber).szParamValue)
End Property

Public Property Let PlayNumber(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszPlayNumber
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到路单结算费率1
Public Property Get ExtraCsRatio1() As Double
     ExtraCsRatio1 = GetParam(cszExtraCsRatio1).szParamValue
End Property

Public Property Let ExtraCsRatio1(ByVal vNewValue As Double)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszExtraCsRatio1
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到路单结算费率2
Public Property Get ExtraCsRatio2() As Double
     ExtraCsRatio2 = GetParam(cszExtraCsRatio2).szParamValue
End Property

Public Property Let ExtraCsRatio2(ByVal vNewValue As Double)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszExtraCsRatio2
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到路单结算费率3
Public Property Get ExtraCsRatio3() As Double
     ExtraCsRatio3 = GetParam(cszExtraCsRatio3).szParamValue
End Property

Public Property Let ExtraCsRatio3(ByVal vNewValue As Double)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszExtraCsRatio3
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到行包结算费率1
Public Property Get ExtraLsRatio1() As Double
    ExtraLsRatio1 = GetParam(cszExtraLsRatio1).szParamValue
End Property

Public Property Let ExtraLsRatio1(ByVal vNewValue As Double)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszExtraLsRatio1
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到行包结算费率2
Public Property Get ExtraLsRatio2() As Double
    ExtraLsRatio2 = GetParam(cszExtraLsRatio2).szParamValue
End Property

Public Property Let ExtraLsRatio2(ByVal vNewValue As Double)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszExtraLsRatio2
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到行包结算费率3
Public Property Get ExtraLsRatio3() As Double
    ExtraLsRatio3 = GetParam(cszExtraLsRatio3).szParamValue
End Property

Public Property Let ExtraLsRatio3(ByVal vNewValue As Double)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszExtraLsRatio3
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到路单结算费率
Public Property Get ExtraCsRatio() As Double
     ExtraCsRatio = GetParam(cszExtraCsRatio).szParamValue
End Property

Public Property Let ExtraCsRatio(ByVal vNewValue As Double)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszExtraCsRatio
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到行包结算费率
Public Property Get ExtraLsRatio() As Double
    ExtraLsRatio = GetParam(cszExtraLsRatio).szParamValue
End Property

Public Property Let ExtraLsRatio(ByVal vNewValue As Double)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszExtraLsRatio
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到路单结算中显示补客票1
Public Property Get IsExtraCsDisplay1() As Boolean
    IsExtraCsDisplay1 = IIf(GetParam(cszExtraCsDisplay1).szParamValue = 0, False, True)
End Property

Public Property Let IsExtraCsDisplay1(ByVal vNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszExtraCsDisplay1
    piTemp.szParamValue = IIf(vNewValue, 1, 0)
    ModifyParam piTemp, True
End Property

'得到路单结算中显示补客票2
Public Property Get IsExtraCsDisplay2() As Boolean
    IsExtraCsDisplay2 = IIf(GetParam(cszExtraCsDisplay2).szParamValue = 0, False, True)
End Property

Public Property Let IsExtraCsDisplay2(ByVal vNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszExtraCsDisplay2
    piTemp.szParamValue = IIf(vNewValue, 1, 0)
    ModifyParam piTemp, True
End Property

'得到路单结算中显示补客票3
Public Property Get IsExtraCsDisplay3() As Boolean
    IsExtraCsDisplay3 = IIf(GetParam(cszExtraCsDisplay3).szParamValue = 0, False, True)
End Property

Public Property Let IsExtraCsDisplay3(ByVal vNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszExtraCsDisplay3
    piTemp.szParamValue = IIf(vNewValue, 1, 0)
    ModifyParam piTemp, True
End Property


'得到路单结算中打印拆入项
Public Property Get IsPrintCsSplitIn() As Boolean
    IsPrintCsSplitIn = IIf(GetParam(cszWantPrintCsSplitIn).szParamValue = 0, False, True)
End Property

Public Property Let IsPrintCsSplitIn(ByVal vNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszWantPrintCsSplitIn
    piTemp.szParamValue = IIf(vNewValue, 1, 0)
    ModifyParam piTemp, True
End Property

'得到路单结算中打印拆出项
Public Property Get IsPrintCsSplitOut() As Boolean
    IsPrintCsSplitOut = IIf(GetParam(cszWantPrintCsSplitOut).szParamValue = 0, False, True)
End Property

Public Property Let IsPrintCsSplitOut(ByVal vNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszWantPrintCsSplitOut
    piTemp.szParamValue = IIf(vNewValue, 1, 0)
    ModifyParam piTemp, True
End Property

'得到路单结算中打印行包运费
Public Property Get IsCsPrintLuggage() As Boolean
    IsCsPrintLuggage = IIf(GetParam(cszWantCsPrintLuggage).szParamValue = 0, False, True)
End Property

Public Property Let IsCsPrintLuggage(ByVal vNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszWantCsPrintLuggage
    piTemp.szParamValue = IIf(vNewValue, 1, 0)
    ModifyParam piTemp, True
End Property

'得到是否路单结算前统计检票信息
Public Property Get IsStatCheckInfo() As Boolean
    IsStatCheckInfo = IIf(GetParam(cszWantStatCheckInfo).szParamValue = 0, False, True)
End Property

Public Property Let IsStatCheckInfo(ByVal vNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszWantStatCheckInfo
    piTemp.szParamValue = IIf(vNewValue, 1, 0)
    ModifyParam piTemp, True
End Property

'得到是否允许修改结算后部分结算项的值
Public Property Get IsModifySplitResult() As Boolean
    IsModifySplitResult = IIf(GetParam(cszWantModifySplitResult).szParamValue = 0, False, True)
End Property

Public Property Let IsModifySplitResult(ByVal vNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszWantModifySplitResult
    piTemp.szParamValue = IIf(vNewValue, 1, 0)
    ModifyParam piTemp, True
End Property

'得到是否签发单和非旅客同行行包一起结算
Public Property Get IsLsSettleTogether() As Boolean
    IsLsSettleTogether = IIf(GetParam(cszWantLsSettleTogether).szParamValue = 0, False, True)
End Property

Public Property Let IsLsSettleTogether(ByVal vNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszWantLsSettleTogether
    piTemp.szParamValue = IIf(vNewValue, 1, 0)
    ModifyParam piTemp, True
End Property

'得到是否允许选择部分路单结算
Public Property Get IsSelectSomeCsSheet() As Boolean
    IsSelectSomeCsSheet = IIf(GetParam(cszWantSelectSomeCsSheet).szParamValue = 0, False, True)
End Property

Public Property Let IsSelectSomeCsSheet(ByVal vNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszWantSelectSomeCsSheet
    piTemp.szParamValue = IIf(vNewValue, 1, 0)
    ModifyParam piTemp, True
End Property

Private Sub Class_Initialize()
    #If SUPPORT_MSG = 1 Then
    m_oSndMsg.MsgSource = cszSNSysMan
    m_oSndMsg.Unit = GetUnitID()
    #End If
End Sub

'得到是否打印退票手续费单据
Public Property Get IsPrintReturnChangeSheet() As Boolean
    IsPrintReturnChangeSheet = IIf(GetParam(cszIsPrintReturnChangeSheet).szParamValue = 0, False, True)
End Property

Public Property Let IsPrintReturnChangeSheet(ByVal vNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszIsPrintReturnChangeSheet
    piTemp.szParamValue = IIf(vNewValue, 1, 0)
    ModifyParam piTemp, True
End Property

'得到退票余额收取参数3
Public Property Get RTRoundingBlanceValue3() As Double
    RTRoundingBlanceValue3 = GetParam(cszRTRoundingBlanceValue3).szParamValue
End Property

Public Property Let RTRoundingBlanceValue3(ByVal vNewValue As Double)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszRTRoundingBlanceValue3
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到退票余额收取参数2
Public Property Get RTRoundingBlanceValue2() As Double
    RTRoundingBlanceValue2 = GetParam(cszRTRoundingBlanceValue2).szParamValue
End Property

Public Property Let RTRoundingBlanceValue2(ByVal vNewValue As Double)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszRTRoundingBlanceValue2
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到退票余额收取参数1
Public Property Get RTRoundingBlanceValue1() As Double
    RTRoundingBlanceValue1 = GetParam(cszRTRoundingBlanceValue1).szParamValue
End Property

Public Property Let RTRoundingBlanceValue1(ByVal vNewValue As Double)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszRTRoundingBlanceValue1
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到退票费
Public Function GetReturnCharge(ByVal pdtBusStartupDateTime As Date, ByVal dbTicketPrice As Double) As Double
    Dim lInterval As Long
    Dim sgTemp As Double
    Dim rfTemp() As RETURNFEES
    Dim iLen, iCount As Integer
    Dim dbLeastMoney As Double
    iLen = 0
    iLen = ArrayLength(GetReturnFees())
    sgTemp = dbTicketPrice
    If iLen = 0 Then
        GetReturnCharge = Round(sgTemp + 0.01, 1)
        Exit Function
    End If
    ReDim rfTemp(1 To iLen)
    rfTemp = GetReturnFees()
    
    lInterval = DateDiff("n", SelfNowDateTime(), pdtBusStartupDateTime)
    
    For iCount = 1 To iLen
        If lInterval >= rfTemp(iCount).iReturnTime Then
            sgTemp = rfTemp(iCount).sgReturnRate * dbTicketPrice / 100
            dbLeastMoney = rfTemp(iCount).sgLeastMoney
            sgTemp = IIf(sgTemp > dbLeastMoney, sgTemp, dbLeastMoney)
            Exit For
        End If
    Next iCount
    
'    If lInterval >= ReturnTicketTime1 Then
'        sgTemp = ReturnTicketRatio1 * dbTicketPrice
'        dbLeastMoney = RTRoundingBlanceValue1
'        sgTemp = IIf(sgTemp > dbLeastMoney, sgTemp, dbLeastMoney)
'
'    ElseIf lInterval >= ReturnTicketTime2 Then
'        sgTemp = ReturnTicketRatio2 * dbTicketPrice
'        dbLeastMoney = RTRoundingBlanceValue2
'        sgTemp = IIf(sgTemp > dbLeastMoney, sgTemp, dbLeastMoney)
'
'    ElseIf lInterval >= ReturnTicketTime3 Then
'        sgTemp = ReturnTicketRatio3 * dbTicketPrice
'        dbLeastMoney = RTRoundingBlanceValue3
'        sgTemp = IIf(sgTemp > dbLeastMoney, sgTemp, dbLeastMoney)
'    End If
    
    GetReturnCharge = Round(sgTemp + 0.01, 1)
End Function


'得到所有票种
Public Function GetAllTicketType(Optional nValid As Integer = TP_TicketTypeAll, Optional bnFreeTicketType As Boolean = True) As TTicketType()
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim szSql  As String
    Dim atTemp() As TTicketType
    
    szSql = " SELECT * FROM Ticket_type_code "
    '默认为所有的票种
    If nValid <> TP_TicketTypeAll Then
        szSql = szSql & " WHERE ticket_type_valid = " & nValid
     
     If bnFreeTicketType = False Then
       szSql = szSql & " AND ticket_type_id<>'" & TP_FreeTicket & "'"
     End If
    
    Else
     
     If bnFreeTicketType = False Then
       szSql = szSql & " where ticket_type_id<>'" & TP_FreeTicket & "'"
     End If
    
    End If
    
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    Set rsTemp = oDb.Execute(szSql)
    
    If rsTemp.RecordCount > 0 Then ReDim atTemp(1 To rsTemp.RecordCount)
        
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
End Function

'得到所有的票种
Public Function GetAllTicketTypeRS(Optional nValid As Integer = TP_TicketTypeAll) As Recordset
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim szSql  As String
    Dim atTemp() As TTicketType
    
    szSql = " SELECT ticket_type_id , ticket_type_name FROM Ticket_type_code "
    '默认为所有的票种
    If nValid <> TP_TicketTypeAll Then
        szSql = szSql & " WHERE ticket_type_valid = " & nValid
    End If
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    Set rsTemp = oDb.Execute(szSql)
    Set GetAllTicketTypeRS = rsTemp
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

'得到预定超时时间
Public Property Get BookTime() As Long
    BookTime = GetParam(cszBookTime).szParamValue
End Property

Public Property Let BookTime(ByVal vNewValue As Long)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszBookTime
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到是否上预定系统
Public Property Get IsBookValid() As Boolean
    IsBookValid = IIf(GetParam(cszIsBookValid).szParamValue = 0, False, True)
End Property

Public Property Let IsBookValid(ByVal vNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszIsBookValid
    piTemp.szParamValue = IIf(vNewValue, 1, 0)
    ModifyParam piTemp, True
End Property

'得到预定号长度
Public Property Get BookNumberLen() As Integer
    BookNumberLen = GetParam(cszBookNumberLen).szParamValue
End Property

Public Property Let BookNumberLen(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszBookNumberLen
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到退票费率信息
Public Function GetReturnFees() As RETURNFEES()
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim iCount As Integer
    Dim rfTemp() As RETURNFEES
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "SELECT * FROM Return_rate_code order by returntkt_number"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 0 Then
        
        If rsTemp.RecordCount > 0 Then ReDim rfTemp(1 To rsTemp.RecordCount)
        For iCount = 1 To rsTemp.RecordCount
            rfTemp(iCount).iReturnNum = rsTemp!returntkt_number
            rfTemp(iCount).iReturnTime = rsTemp!returntkt_time
            rfTemp(iCount).sgLeastMoney = rsTemp!least_money
            rfTemp(iCount).sgReturnRate = rsTemp!returntkt_rate
            rsTemp.MoveNext
        Next iCount
'    Else
'        ReDim rfTemp(1 To 1)
'        rfTemp(1).iReturnNum = -1
    End If
    GetReturnFees = rfTemp
    Set rsTemp = Nothing

End Function

'设置退票费率信息
Public Sub SetReturnFees(FeesNum As Integer, rfReturnFees() As RETURNFEES)
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim iCount As Integer
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    oDb.BeginTrans
        szSql = "SELECT * FROM Return_rate_code order by returntkt_number"
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount <> 0 Then
          szSql = "DELETE  Return_rate_code "
          oDb.Execute szSql
        End If
        If FeesNum > 0 Then
            For iCount = 1 To FeesNum
                szSql = "INSERT INTO Return_rate_code(" _
                & "returntkt_number," _
                & "returntkt_time," _
                & "returntkt_rate," _
                & "least_money)"
                szSql = szSql & " VALUES(" _
                & CInt(rfReturnFees(iCount).iReturnNum) & "," _
                & CInt(rfReturnFees(iCount).iReturnTime) & ",'" _
                & CSng(rfReturnFees(iCount).sgReturnRate) & "','" _
                & CSng(rfReturnFees(iCount).sgLeastMoney) & "')"
                oDb.Execute szSql
            Next iCount
        End If
    oDb.CommitTrans
    Set rsTemp = Nothing
End Sub

'得到车次打印长度
Public Property Get PrintBusIDLen() As Integer
    PrintBusIDLen = GetParam(cszPrintBusIDLen).szParamValue
End Property

Public Property Let PrintBusIDLen(ByVal pnLen As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszPrintBusIDLen
    piTemp.szParamValue = pnLen
    ModifyParam piTemp, True
End Property

'得到滚动车次发车时间打印方式
Public Property Get PrintScrollBusMode() As Boolean
    PrintScrollBusMode = GetParam(cszPrintScrollBusMode).szParamValue
End Property

Public Property Let PrintScrollBusMode(ByVal pbMode As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszPrintScrollBusMode
    piTemp.szParamValue = IIf(pbMode, 1, 0)
    ModifyParam piTemp, True
End Property

'得到环境停班车次是否生成
Public Property Get MekeStopEnviroment() As Boolean
    Dim nTemp As Integer
    nTemp = GetParam(cszMakeStopEnviroment).szParamValue
    If nTemp = 1 Then
       MekeStopEnviroment = True  '生成默认
    Else
       MekeStopEnviroment = False
    End If
End Property

Public Property Let MekeStopEnviroment(ByVal pbMode As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszMakeStopEnviroment
    piTemp.szParamValue = IIf(pbMode, 1, 0)
    ModifyParam piTemp, True
End Property

'设置环境加班车是否最后一站可售
Public Property Get EndStationCanSale() As Boolean
    Dim nTemp As Integer
    nTemp = GetParam(cszEndStatinCanSale).szParamValue
    If nTemp = 1 Then
       EndStationCanSale = True  '生成默认
    Else
       EndStationCanSale = False
    End If
End Property

'以站点查询
Public Property Let StationImmediatelyQuery(ByVal pbMode As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszImmediatelyQuery
    piTemp.szParamValue = IIf(pbMode, 1, 0)
    ModifyParam piTemp, True
End Property

Public Property Get StationImmediatelyQuery() As Boolean
    Dim nTemp As Integer
    On Error GoTo here
    nTemp = GetParam(cszImmediatelyQuery).szParamValue
    If nTemp = 1 Then
       StationImmediatelyQuery = True  '生成默认
    Else
       StationImmediatelyQuery = False
    End If
   Exit Property
here:
    StationImmediatelyQuery = False
End Property

'允许生成已售票车次的环境
Public Property Let MakeEnviromentSaleBus(ByVal pbMode As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszAllowMakeEnviromentSaleBus
    piTemp.szParamValue = IIf(pbMode, 1, 0)
    ModifyParam piTemp, True
End Property

'允许生成已售票车次的环境
Public Property Get MakeEnviromentSaleBus() As Boolean
    Dim nTemp As Integer
    nTemp = GetParam(cszAllowMakeEnviromentSaleBus).szParamValue
    If nTemp = 1 Then
       MakeEnviromentSaleBus = True  '生成默认
    Else
       MakeEnviromentSaleBus = False
    End If
End Property

Public Property Let EndStationCanSale(ByVal pbMode As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszEndStatinCanSale
    piTemp.szParamValue = IIf(pbMode, 1, 0)
    ModifyParam piTemp, True
End Property

'允许拆分不过站 1--允许
Public Property Get AllowSlitp() As Boolean
    AllowSlitp = IIf(CInt(GetParam(cszAllowSlitpNotPassStation).szParamValue) = 1, True, False)
End Property

Public Property Let AllowSlitp(ByVal bNewValue As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszAllowSlitpNotPassStation
    piTemp.szParamValue = IIf(bNewValue, 1, 0)
    ModifyParam piTemp, True
End Property

'得到票价项小数点后位数
Public Property Get PriceDetailKeepBit() As Integer
    PriceDetailKeepBit = GetParam(cszPriceDetailKeepBit).szParamValue
  End Property

Public Property Let PriceDetailKeepBit(ByVal pbMode As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszPriceDetailKeepBit
    piTemp.szParamValue = pbMode
    ModifyParam piTemp, True
End Property

'得到车辆平均速度
Public Property Get BusSpeed() As Single
    BusSpeed = GetParam(cszBusSpeed).szParamValue
  End Property

Public Property Let BusSpeed(ByVal sgValue As Single)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszBusSpeed
    piTemp.szParamValue = sgValue
    ModifyParam piTemp, True
End Property

'得到夜班费第一时间段
Public Property Get NightShiftTime1() As Date
    NightShiftTime1 = GetParam(cszNightShiftTime1).szParamValue
  End Property

Public Property Let NightShiftTime1(ByVal dtNewTime As Date)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszNightShiftTime1
    piTemp.szParamValue = dtNewTime
    ModifyParam piTemp, True
End Property

'得到夜班费第二时间段
Public Property Get NightShiftTime2() As Date
    NightShiftTime2 = GetParam(cszNightShiftTime2).szParamValue
  End Property

Public Property Let NightShiftTime2(ByVal dtNewTime As Date)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszNightShiftTime2
    piTemp.szParamValue = dtNewTime
    ModifyParam piTemp, True
End Property

'得到加班车前缀
Public Property Get AdditionBusPreFix() As String
    AdditionBusPreFix = GetParam(cszAdditionBusPreFix).szParamValue
End Property

Public Property Let AdditionBusPreFix(ByVal pszBusPreFix As String)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszAdditionBusPreFix
    piTemp.szParamValue = pszBusPreFix
    ModifyParam piTemp, True
End Property

'得到滚动车次最迟上车时间
Public Property Get ScrollBusLatestTime() As Integer
    ScrollBusLatestTime = GetParam(cszScrollBusLatestTime).szParamValue
End Property

Public Property Let ScrollBusLatestTime(pszScrollBusLatestTime As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszScrollBusLatestTime
    piTemp.szParamValue = pszScrollBusLatestTime
    ModifyParam piTemp, True
End Property

'得到行包单号数字部分长
Public Property Get LuggageIDNumberLen() As Integer
    LuggageIDNumberLen = GetParam(cszLuggageIDNumberLen).szParamValue
End Property

Public Property Let LuggageIDNumberLen(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszLuggageIDNumberLen
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到行包单号前辍部分长
Public Property Get LuggageIDPrefixLen() As Integer
    LuggageIDPrefixLen = GetParam(cszLuggageIDPrefixLen).szParamValue
End Property

Public Property Let LuggageIDPrefixLen(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszLuggageIDPrefixLen
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到签发单号数字部分长
Public Property Get CheckLuggageIDNumberLen() As Integer
    CheckLuggageIDNumberLen = GetParam(cszCheckLuggageIDNumberLen).szParamValue
End Property

Public Property Let CheckLuggageIDNumberLen(ByVal vNewValue As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszCheckLuggageIDNumberLen
    piTemp.szParamValue = vNewValue
    ModifyParam piTemp, True
End Property

'得到签发时限
Public Property Get CheckLimitTime() As Integer
    CheckLimitTime = GetParam(cszCheckLimitTime).szParamValue
End Property

Public Property Let CheckLimitTime(pnCheckLimitTime As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszCheckLimitTime
    piTemp.szParamValue = pnCheckLimitTime
    ModifyParam piTemp, True
End Property

'得到受理时限
Public Property Get AcceptLimitTime() As Integer
    AcceptLimitTime = GetParam(cszAcceptLimitTime).szParamValue
End Property

Public Property Let AcceptLimitTime(pnAcceptLimitTime As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszAcceptLimitTime
    piTemp.szParamValue = pnAcceptLimitTime
    ModifyParam piTemp, True
End Property

'得到停班车可以签发
Public Property Get StopBusCanCheck() As Boolean
    StopBusCanCheck = IIf(GetParam(cszStopBusCanCheck).szParamValue = 0, False, True)
End Property

Public Property Let StopBusCanCheck(pbStopBusCanCheck As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszStopBusCanCheck
    piTemp.szParamValue = IIf(pbStopBusCanCheck, 1, 0)
    ModifyParam piTemp, True
End Property

'得到停班车可以受理
Public Property Get StopBusCanAccept() As Boolean
    StopBusCanAccept = IIf(GetParam(cszStopBusCanAccept).szParamValue = 0, False, True)
End Property

Public Property Let StopBusCanAccept(pbStopBusCanAccept As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszStopBusCanAccept
    piTemp.szParamValue = IIf(pbStopBusCanAccept, 1, 0)
    ModifyParam piTemp, True
End Property

'得到专业车前缀
Public Property Get ProfessionalBusPreFix() As String
    ProfessionalBusPreFix = GetParam(cszProfessionalBusPreFix).szParamValue
End Property

Public Property Let ProfessionalBusPreFix(pszProfessionalBusPreFix As String)
'期间是用逗号分隔的
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszProfessionalBusPreFix
    piTemp.szParamValue = pszProfessionalBusPreFix
    ModifyParam piTemp, True
End Property

'得到行包刷新频率
Public Property Get LuggageRefreshFre() As Integer
    LuggageRefreshFre = GetParam(cszLuggageRefreshFre).szParamValue
End Property

Public Property Let LuggageRefreshFre(pnLuggageRefreshFre As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszLuggageRefreshFre
    piTemp.szParamValue = pnLuggageRefreshFre
    ModifyParam piTemp, True
End Property

'得到装卸费为第几个收费项
Public Property Get LuggageLoadItemNo() As Integer
    LuggageLoadItemNo = GetParam(cszLuggageLoadItemNo).szParamValue
End Property

Public Property Let LuggageLoadItemNo(pnLuggageLoadItemNo As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszLuggageLoadItemNo
    piTemp.szParamValue = pnLuggageLoadItemNo
    ModifyParam piTemp, True
End Property

'得到搬运为第几个收费项
Public Property Get LuggageMoveItemNo() As Integer
    LuggageMoveItemNo = GetParam(cszLuggageMoveItemNo).szParamValue
End Property

Public Property Let LuggageMoveItemNo(pnLuggageMoveItemNo As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszLuggageMoveItemNo
    piTemp.szParamValue = pnLuggageMoveItemNo
    ModifyParam piTemp, True
End Property

'补票车次显示前几小时
Public Property Get ExtraBusShowBefore() As Double
    ExtraBusShowBefore = GetParam(cszExtraBusShowBefore).szParamValue
End Property

Public Property Let ExtraBusShowBefore(pExtraBusShowBefore As Double)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszExtraBusShowBefore
    piTemp.szParamValue = pExtraBusShowBefore
    ModifyParam piTemp, True
End Property

'补票车次显示后几小时
Public Property Get ExtraBusShowAfter() As Double
    ExtraBusShowAfter = GetParam(cszExtraBusShowAfter).szParamValue
End Property

Public Property Let ExtraBusShowAfter(pExtraBusShowAfter As Double)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszExtraBusShowAfter
    piTemp.szParamValue = pExtraBusShowAfter
    ModifyParam piTemp, True
End Property

'互联售票车次显示时间
Public Property Get InternetBusShow() As Double
    InternetBusShow = GetParam(cszInternetBusShow).szParamValue
End Property

Public Property Let InternetBusShow(pInternetBusShow As Double)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszInternetBusShow
    piTemp.szParamValue = pInternetBusShow
    ModifyParam piTemp, True
End Property

'小计项1
Public Property Get SubTotalItem1() As String
    SubTotalItem1 = GetParam(cszSubTotalItem1).szParamValue
End Property

Public Property Let SubTotalItem1(pszSubTotalItem1 As String)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszSubTotalItem1
    piTemp.szParamValue = pszSubTotalItem1
    ModifyParam piTemp, True
End Property

'基本运价合计项
Public Property Get BasePriceTotalItem() As String
    BasePriceTotalItem = GetParam(cszBasePriceTotalItem).szParamValue
End Property

Public Property Let BasePriceTotalItem(pszBasePriceTotalItem As String)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszBasePriceTotalItem
    piTemp.szParamValue = pszBasePriceTotalItem
    ModifyParam piTemp, True
End Property

'运量统计需要列出的站点明细
Public Property Get TranNeedStationDetail() As String
    TranNeedStationDetail = GetParam(cszTranNeedStationDetail).szParamValue
End Property

Public Property Let TranNeedStationDetail(pszTranNeedStationDetail As String)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszTranNeedStationDetail
    piTemp.szParamValue = pszTranNeedStationDetail
    ModifyParam piTemp, True
End Property

'省际班车需列出的站点明细
Public Property Get InterProvinceStationDetail() As String
    InterProvinceStationDetail = GetParam(cszInterProvinceStationDetail).szParamValue
End Property

Public Property Let InterProvinceStationDetail(pszInterProvinceStationDetail As String)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszInterProvinceStationDetail
    piTemp.szParamValue = pszInterProvinceStationDetail
    ModifyParam piTemp, True
End Property

'市内班车需列出的站点明细
Public Property Get InCityStationDetail() As String
    InCityStationDetail = GetParam(cszInCityStationDetail).szParamValue
End Property

Public Property Let InCityStationDetail(pszInCityStationDetail As String)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszInCityStationDetail
    piTemp.szParamValue = pszInCityStationDetail
    ModifyParam piTemp, True
End Property

'市外班车需列出的站点明细
Public Property Get ExtramuralStationDetail() As String
    ExtramuralStationDetail = GetParam(cszExtramuralStationDetail).szParamValue
End Property

Public Property Let ExtramuralStationDetail(pszExtramuralStationDetail As String)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszExtramuralStationDetail
    piTemp.szParamValue = pszExtramuralStationDetail
    ModifyParam piTemp, True
End Property

'导出数据保留的日期
Public Property Get KeepDate() As Integer
    KeepDate = GetParam(cszKeepDate).szParamValue
End Property

Public Property Let KeepDate(pnKeepDate As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszKeepDate
    piTemp.szParamValue = pnKeepDate
    ModifyParam piTemp, True
End Property

'春运费计算距离1
Public Property Get AdvanceDistance1() As Single
    AdvanceDistance1 = GetParam(cszAdvanceDistance1).szParamValue
End Property

Public Property Let AdvanceDistance1(sgDistance As Single)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszAdvanceDistance1
    piTemp.szParamValue = sgDistance
    ModifyParam piTemp, True
End Property

'春运费计算距离2
Public Property Get AdvanceDistance2() As Single
    AdvanceDistance2 = GetParam(cszAdvanceDistance2).szParamValue
End Property

Public Property Let AdvanceDistance2(sgDistance As Single)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszAdvanceDistance2
    piTemp.szParamValue = sgDistance
    ModifyParam piTemp, True
End Property

'公建金为第几个票价项
Public Property Get RoadBuildChargeItem() As Integer
    RoadBuildChargeItem = GetParam(cszRoadBuildChargeItem).szParamValue
End Property

Public Property Let RoadBuildChargeItem(pnRoadBuildChargeItem As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszRoadBuildChargeItem
    piTemp.szParamValue = pnRoadBuildChargeItem
    ModifyParam piTemp, True
End Property

'250K加成费为第几个票价项
Public Property Get FarDistanceAddChargeItem() As Integer
    FarDistanceAddChargeItem = GetParam(cszFarDistanceAddChargeItem).szParamValue
End Property

Public Property Let FarDistanceAddChargeItem(pnFarDistanceAddChargeItem As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszFarDistanceAddChargeItem
    piTemp.szParamValue = pnFarDistanceAddChargeItem
    ModifyParam piTemp, True
End Property

 '春运费为第几个票价项
Public Property Get SpringItem() As Integer
    SpringItem = GetParam(cszSpringItem).szParamValue
End Property

Public Property Let SpringItem(pnSpringItem As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszSpringItem
    piTemp.szParamValue = pnSpringItem
    ModifyParam piTemp, True
End Property

'得到用户的可售天数
Public Property Get UserCanSellDate() As Integer
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)

    szSql = "SELECT can_sell_day FROM user_info WHERE " _
    & " user_id='" & m_oActiveUser.UserID & "' "
    Set rsTemp = oDb.Execute(szSql)
    UserCanSellDate = FormatDbValue(rsTemp!can_sell_day)
End Property

'售票站之间是否相互可售
Public Property Get SellStationCanSellEachOther() As Boolean
    SellStationCanSellEachOther = IIf(GetParam(cszSellStationCanSellEachOther).szParamValue = "1", True, False)
End Property

Public Property Let SellStationCanSellEachOther(ByVal value As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszSellStationCanSellEachOther
    If value = True Then
        piTemp.szParamValue = "1"
    Else
        piTemp.szParamValue = "0"
    End If
    ModifyParam piTemp, True
End Property

'售票是否使用快速打印
Public Property Get WantDirectSheetPrint() As Boolean
    WantDirectSheetPrint = IIf(GetParam(cszWantDirectSheetPrint).szParamValue = "1", True, False)
End Property

Public Property Let WantDirectSheetPrint(ByVal value As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszWantDirectSheetPrint
    If value = True Then
        piTemp.szParamValue = "1"
    Else
        piTemp.szParamValue = "0"
    End If
    ModifyParam piTemp, True
End Property

'售票时是否列出无座位的车次
Public Property Get WantListNoSeatBus() As Boolean
    WantListNoSeatBus = IIf(GetParam(cszWantListNoSeatBus).szParamValue = "1", True, False)
End Property

Public Property Let WantListNoSeatBus(ByVal value As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszWantListNoSeatBus
    If value = True Then
        piTemp.szParamValue = "1"
    Else
        piTemp.szParamValue = "0"
    End If
    ModifyParam piTemp, True
End Property

'是否每站点一公式,0为所有站点相同公司1为每站点一公式)
Public Property Get IsOneFormulaEachStation() As Boolean
    IsOneFormulaEachStation = IIf(GetParam(cszIsOneFormulaEachStation).szParamValue = "1", True, False)
End Property

Public Property Let IsOneFormulaEachStation(ByVal value As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszIsOneFormulaEachStation
    If value = True Then
        piTemp.szParamValue = "1"
    Else
        piTemp.szParamValue = "0"
    End If
    ModifyParam piTemp, True
End Property

'计算公式是否与车型有关
Public Property Get IsRelationWithVehicleType() As Boolean
    IsRelationWithVehicleType = IIf(GetParam(cszIsRelationWithVehicleType).szParamValue = "1", True, False)
End Property

Public Property Let IsRelationWithVehicleType(ByVal value As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszIsRelationWithVehicleType
    If value = True Then
        piTemp.szParamValue = "1"
    Else
        piTemp.szParamValue = "0"
    End If
    ModifyParam piTemp, True
End Property

'是否打印签发单
Public Property Get IsPrintCheckSheet() As Boolean
    IsPrintCheckSheet = IIf(GetParam(cszIsPrintCheckSheet).szParamValue = "1", True, False)
End Property

Public Property Let IsPrintCheckSheet(ByVal value As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszIsPrintCheckSheet
    If value = True Then
        piTemp.szParamValue = "1"
    Else
        piTemp.szParamValue = "0"
    End If
    ModifyParam piTemp, True
End Property

'是否在签发时显示应结费用
Public Property Get IsDispSettlePriceInCheck() As Boolean
    IsDispSettlePriceInCheck = IIf(GetParam(cszIsDispSettlePriceInCheck).szParamValue = "1", True, False)
End Property

Public Property Let IsDispSettlePriceInCheck(ByVal value As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszIsDispSettlePriceInCheck
    If value = True Then
        piTemp.szParamValue = "1"
    Else
        piTemp.szParamValue = "0"
    End If
    ModifyParam piTemp, True
End Property

'结算运费是不是从受理的结算运费中汇总得到
Public Property Get IsSettlePriceFromAcceptInCheck() As Boolean
    IsSettlePriceFromAcceptInCheck = IIf(GetParam(cszIsSettlePriceFromAcceptInCheck).szParamValue = "1", True, False)
End Property

Public Property Let IsSettlePriceFromAcceptInCheck(ByVal value As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszIsSettlePriceFromAcceptInCheck
    If value = True Then
        piTemp.szParamValue = "1"
    Else
        piTemp.szParamValue = "0"
    End If
    ModifyParam piTemp, True
End Property

'是否在受理时显示应结费用
Public Property Get IsDispSettlePriceInAccept() As Boolean
    IsDispSettlePriceInAccept = IIf(GetParam(cszIsDispSettlePriceInAccept).szParamValue = "1", True, False)
End Property

Public Property Let IsDispSettlePriceInAccept(ByVal value As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszIsDispSettlePriceInAccept
    If value = True Then
        piTemp.szParamValue = "1"
    Else
        piTemp.szParamValue = "0"
    End If
    ModifyParam piTemp, True
End Property

'是否允许路单分几次来结算 0,不允许,1允许
Public Property Get AllowSplitBySomeTimes() As Boolean
    AllowSplitBySomeTimes = IIf(GetParam(cszAllowSplitBySomeTimes).szParamValue = "1", True, False)
End Property

Public Property Let AllowSplitBySomeTimes(ByVal value As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszAllowSplitBySomeTimes
    If value = True Then
        piTemp.szParamValue = "1"
    Else
        piTemp.szParamValue = "0"
    End If
    ModifyParam piTemp, True
End Property

'是否允许路单结算人数超过实际人数? 0不允许,1允许
Public Property Get AllowSplitAboveFactQuantity() As Boolean
    AllowSplitAboveFactQuantity = IIf(GetParam(cszAllowSplitAboveFactQuantity).szParamValue = "1", True, False)
End Property

Public Property Let AllowSplitAboveFactQuantity(ByVal value As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszAllowSplitAboveFactQuantity
    If value = True Then
        piTemp.szParamValue = "1"
    Else
        piTemp.szParamValue = "0"
    End If
    ModifyParam piTemp, True
End Property

'是否将上月结算的负数算到下月 0,不允许,1允许
Public Property Get AllowSettleTotalNegative() As Boolean
    AllowSettleTotalNegative = IIf(GetParam(cszAllowSettleTotalNegative).szParamValue = "1", True, False)
End Property

Public Property Let AllowSettleTotalNegative(ByVal value As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszAllowSettleTotalNegative
    If value = True Then
        piTemp.szParamValue = "1"
    Else
        piTemp.szParamValue = "0"
    End If
    ModifyParam piTemp, True
End Property

'将上月结算的负数项放在结算项的位置
Public Property Get SettleNegativeSplitItem() As Integer
    SettleNegativeSplitItem = GetParam(cszSettleNegativeSplitItem).szParamValue
End Property

Public Property Let SettleNegativeSplitItem(value As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszSettleNegativeSplitItem
    piTemp.szParamValue = value
    ModifyParam piTemp, True
End Property

'结算项的固定费用项  ,其值为项目编号用逗号分隔 如:1,3,5,11
Public Property Get FixFeeItem() As String
    FixFeeItem = GetParam(cszFixFeeItem).szParamValue
End Property

Public Property Let FixFeeItem(value As String)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszFixFeeItem
    piTemp.szParamValue = value
    ModifyParam piTemp, True
End Property

'劳务费的项目位置
Public Property Get ServiceItemPosition() As Integer
    ServiceItemPosition = GetParam(cszServiceItemPosition).szParamValue
End Property

Public Property Let ServiceItemPosition(value As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszServiceItemPosition
    piTemp.szParamValue = value
    ModifyParam piTemp, True
End Property

'是否固定费用是每个月更新的 0,不允许,1允许
Public Property Get IsFixFeeUpdateEachMonth() As Boolean
    IsFixFeeUpdateEachMonth = IIf(GetParam(cszIsFixFeeUpdateEachMonth).szParamValue = "1", True, False)
End Property

Public Property Let IsFixFeeUpdateEachMonth(ByVal value As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszIsFixFeeUpdateEachMonth
    If value = True Then
        piTemp.szParamValue = "1"
    Else
        piTemp.szParamValue = "0"
    End If
    ModifyParam piTemp, True
End Property

'运费的项目位置
Public Property Get CarriageItemPosition() As Integer
    CarriageItemPosition = GetParam(cszCarriageItemPosition).szParamValue
End Property

Public Property Let CarriageItemPosition(value As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszCarriageItemPosition
    piTemp.szParamValue = value
    ModifyParam piTemp, True
End Property

'春运费项目位置
Public Property Get SpringChargeItem() As Integer
    SpringChargeItem = Val(GetParam(cszSpringChargeItem).szParamValue)
End Property

Public Property Let SpringChargeItem(value As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszSpringChargeItem
    piTemp.szParamValue = value
    ModifyParam piTemp, True
End Property

'是否打印手续费为0的退票手续费 0不允许,1允许
Public Property Get IsPrintZeroValueReturnSheet() As Boolean
    IsPrintZeroValueReturnSheet = IIf(GetParam(cszIsPrintZeroValueReturnSheet).szParamValue = "1", True, False)
End Property

Public Property Let IsPrintZeroValueReturnSheet(ByVal value As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszIsPrintZeroValueReturnSheet
    If value = True Then
        piTemp.szParamValue = "1"
    Else
        piTemp.szParamValue = "0"
    End If
    ModifyParam piTemp, True
End Property

'滚动班次是否永远可售 0,不允许,1允许
Public Property Get AllowScrollBusSaleForever() As Boolean
    AllowScrollBusSaleForever = IIf(GetParam(cszAllowScrollBusSaleForever).szParamValue = "1", True, False)
End Property

Public Property Let AllowScrollBusSaleForever(ByVal value As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszAllowScrollBusSaleForever
    If value = True Then
        piTemp.szParamValue = "1"
    Else
        piTemp.szParamValue = "0"
    End If
    ModifyParam piTemp, True
End Property

'允许未到开检时间直接开检 0,不允许,1允许
Public Property Get AllowStartChectNotRearchTime() As Boolean
    AllowStartChectNotRearchTime = IIf(GetParam(cszAllowStartChectNotRearchTime).szParamValue = "1", True, False)
End Property

Public Property Let AllowStartChectNotRearchTime(ByVal value As Boolean)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszAllowStartChectNotRearchTime
    If value = True Then
        piTemp.szParamValue = "1"
    Else
        piTemp.szParamValue = "0"
    End If
    ModifyParam piTemp, True
End Property

'折扣票在票种项的位置
Public Property Get DiscountTicketInTicketTypePosition() As Integer
    DiscountTicketInTicketTypePosition = GetParam(cszDiscountTicketInTicketTypePosition).szParamValue
End Property

Public Property Let DiscountTicketInTicketTypePosition(value As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszDiscountTicketInTicketTypePosition
    piTemp.szParamValue = value
    ModifyParam piTemp, True
End Property

'同一证件购票时间间隔
Public Property Get CardBuyTicketInterval() As Integer
    CardBuyTicketInterval = GetParam(cszCardBuyTicketInterval).szParamValue
End Property

Public Property Let CardBuyTicketInterval(ByVal value As Integer)
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszCardBuyTicketInterval
    piTemp.szParamValue = value
    ModifyParam piTemp, True
End Property
