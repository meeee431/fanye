VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LuggageSvr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E03A51101FF"
'行包托运服务
Option Base 0
Option Explicit


'以下定义枚举类型,公用部分
'======================================================
'中间层行包类的类号
Public Enum ECOMID_Luggage
    CID_LuggageAcceptSheet = 50
    CID_LuggageCarrySheet = 51
    CID_LuggageKinds = 52
    CID_LuggageParam = 53
    CID_LuggageSvr = 54
    CID_LuggageProtocol = 55
    CID_LuggageFromula = 56
    CID_LuggageFinSplitSvr = 57
    CID_LuggageDataStatSvr = 58
    CID_LuggageFinanceSheet = 59
End Enum


'运费项信息
Public Type TLuggagePriceItem
    PriceID As String        '运费项代码
    PriceName As String      '运费项名称
    PriceValue As Double     '运费项值
End Type
    
'行包收费项结构
Public Type TLuggagePriceItemFormula
    PriceItem As String                 '代码
    PriceItemName As String             '项目名称
    AcceptType As ELuggageAcceptType    '托运类型
    UsedMark As Boolean                 '是否使用
'    Formula As String                   '公式
'    szParam1 As String                  '参数...
'    szParam2 As String
'    szParam3 As String
'    szParam4 As String
'    szParam5 As String
End Type
'行包收费项结构
Public Type TLuggagePriceItemFormulaEx
    PriceItem As String                 '代码
    PriceItemName As String             '项目名称
    AcceptType As ELuggageAcceptType    '托运类型
    UsedMark As Boolean                 '是否使用
    Formula As String                   '公式
    szParam1 As String                  '参数...
    szParam2 As String
    szParam3 As String
    szParam4 As String
    szParam5 As String
    szParam6 As String
    szParam7 As String
    szParam8 As String
    szParam9 As String
    szParam10 As String
End Type
    


'行包明细信息
Public Type TLuggageItemInfo
    LabelID As String       '标签号
    LuggageID As String     '受理单号
    LuggageName As String   '行包名称
    LuggageType As String   '行包类型
    LuggageTypeName As String '行包类型名称
    Number As Integer       '件数
    ActWeight As Double     '实重
    CalWeight As Double     '计重
    luggage_bulk As String          '体积
    PackType As String      '包装形势
End Type

'行包种类
Public Type TLuggageKinds
    KindsCode As String '种类代码
    KindsName As String '各类名称
    Annotation As String '备注
End Type



'2004-8-3陈峰新加

Public Type TLuggageFormulaInfo '行包的公式结构
    FormulaID As String
    FormulaName As String
    ItemFormula As String
    Param(1 To 10) As Double
'    Param2 As Double
'    Param3 As Double
'    Param4 As Double
'    Param5 As Double
'    Param6 As Double
'    Param7 As Double
'    Param8 As Double
'    Param9 As Double
'    Param10 As Double
End Type
Public Type TLuggageItemFormulaInfo '行包的站点计算公式
    PriceItemID As String '行包收费项代码
    PriceItemName As String '行包收费项名称
    AcceptType As Integer '托运方式
    AcceptTypeName As String '托运方式名称
    StationID As String '站点代码
    StationName As String '站点名称
    FormulaID As String '公式代码
    FormulaName As String '公式名称
    
End Type

Public Type TConnectUnit
   szUnitID As String '车站编号
   szUnitShortName As String '车站名称简称
   szUnitFullName As String '车站名称全称
   szIPAddress As String '车站服务器IP
   szDatabase As String  '数据库名称
   szDBUser As String '数据库用户
   szUserPassword As String '用户密码
End Type

Public Enum ELuggageAcceptType      '行包类型
    NormalAccept = 0        '普通类型(快件类型)
    CarryAccept = 1         '随行行包
End Enum

Public Enum ELuggagePriceItemUseMark '行包收费项使用标志
    LuggagePriceItemUsed = 0 '行包收费项使用
    LuggagePriceItemNotUsed = 1 '不使用
End Enum

Public Enum ELuggageCarrySettlementStatus
    ST_LuggageCarrySettlementNo = 0
    ST_LuggageCarrySettlement = 1
    ST_LuggageCarrySettlementMuch = 2
    
End Enum
    
Public Enum ELuggageAcceptStatus
    ST_LuggageAcceptNormal = 0    '正常待受理
    ST_LuggageAcceptCancel = 1    '行包单已作废
    ST_LuggageAcceptReturn = 2     '行包单已退运
    ST_LuggageAcceptOver = 3        '行包单已签发
End Enum

Public Enum ELuggageSheetValidStatus
    ST_LuggageSheetValidStatusNormal = 1 '签发单正常
    ST_LuggageSheetValidStatusCancel = 0 ' 签发单作废
End Enum


'行包类的错误的起始号(实际也是其所有资源的起始号)
Public Enum EClassErrBegin_Luggage
    ERR_LuggageAcceptSheet = CID_LuggageAcceptSheet * cnMidErrStep + cnMidErrBegin  '30000
    ERR_LuggageCarrySheet = CID_LuggageCarrySheet * cnMidErrStep + cnMidErrBegin    '30400
    ERR_LuggageKinds = CID_LuggageKinds * cnMidErrStep + cnMidErrBegin      '30800
    ERR_LuggageParam = CID_LuggageParam * cnMidErrStep + cnMidErrBegin      '31200
    ERR_LuggageSvr = CID_LuggageSvr * cnMidErrStep + cnMidErrBegin      '31600
    ERR_LuggageProtocol = CID_LuggageProtocol * cnMidErrStep + cnMidErrBegin '32000
    ERR_LuggageFromula = CID_LuggageFromula * cnMidErrStep + cnMidErrBegin '32400
    ERR_LuggageFinSplitSvr = CID_LuggageFinSplitSvr * cnMidErrStep + cnMidErrBegin '32800
    ERR_LuggageDataStatSvr = CID_LuggageDataStatSvr * cnMidErrStep + cnMidErrBegin '33200
    ERR_LuggageFinanceSheet = CID_LuggageFinanceSheet * cnMidErrStep + cnMidErrBegin '33400
    
    
End Enum

    

'本模块私有定义部分
'==========================================================================
'错误号
Public Enum EErrLuggageSvr
    ERR_LuggageSvrHaveNo = ERR_LuggageSvr + 21  '无此行包受理单     '31621开始
    ERR_LuggageSvrSheetAcceptted = ERR_LuggageSvr + 22 '行包已受理
    ERR_LuggageSvrSheetCanceled = ERR_LuggageSvr + 23 '行包已作废
    ERR_LuggageSvrSheetReturned = ERR_LuggageSvr + 24 '行包已退
    ERR_LuggageSvrSheetCarryed = ERR_LuggageSvr + 25 '行包已签发
    ERR_LuggageSvrCarrySheetCancel = ERR_LuggageSvr + 26 '签发单已作废
    ERR_LuggageSvrCarrySheetHave = ERR_LuggageSvr + 27 '签发单已存在
    ERR_LuggageSvrStationNoInfo = ERR_LuggageSvr + 28 '没有此车次信息
    ERR_LuggageSvrHasExist = ERR_LuggageSvr + 29 '行包受理单已存在
    
    ERR_LuggagePriceFormulaHasExist = ERR_LuggageSvr + 31 '行包计算公式已存在
    ERR_LuggagePriceFormulaNotExist = ERR_LuggageSvr + 32 '行包计算公式不存在
    
    ERR_StationNotExist = ERR_LuggageSvr + 33 '站点不存在
    
End Enum
'权限号
Public Enum ERightLuggageSvr
    RIGHT_LuggageSvr = ERR_LuggageSvr + cnMidRightBegin             '31700开始
    RIGHT_LuggageSvrNormalAccept = RIGHT_LuggageSvr + cnMidRightStep * 1 '行包受理
    RIGHT_LuggageSvrNormalCarry = RIGHT_LuggageSvr + cnMidRightStep * 2 '行包单签发
    RIGHT_LuggageSvrCancel = RIGHT_LuggageSvr + cnMidRightStep * 3 '废行包受理单
    RIGHT_LuggageSvrReturn = RIGHT_LuggageSvr + cnMidRightStep * 4 '退行包受理单
    RIGHT_LuggageSvrReprintCarry = RIGHT_LuggageSvr + cnMidRightStep * 5 '重打行包签发单
    RIGHT_LuggageSvrTotalAccept = RIGHT_LuggageSvr + cnMidRightStep * 6 '统计受理情况
End Enum

'以下定义本模块变量
'------------------------------------------
Private mszSellStationID As String      '当前售票站
Private m_oActiveUser As ActiveUser



'初始化活动对象和当前售票站点
'##ModelId=3E03A6A700B8
Public Sub Init(poAUser As ActiveUser)
    On Error GoTo InitErr
    Set m_oActiveUser = poAUser
    
    mszSellStationID = poAUser.SellStationID

    Exit Sub
InitErr:
    Call RaiseError(err.Number, "LuggageSvr:Init Method", err.Description)
End Sub

'行包受理
'##ModelId=3E03AB1D0088
Public Sub AcceptLuggage(poAcceptSheet As AcceptSheet)
    '验证活动用户及权限
'    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
'    AssertHaveRight m_oActiveUser, RIGHT_LuggageSvrNormalAccept
'
    On Error GoTo AcceptLuggageErr
    
    Dim szSql As String
    Dim szTempPriceSql As String
    Dim oDb As New RTConnection
    Dim Count As Integer
    Dim i As Integer, j As Integer
    Dim atPriceItems() As TLuggagePriceItem
    Dim nCountPriceID As Integer
    Dim szTempPriceID  As String
    Dim szTempPriceID1 As String
    Dim szTempPriceItem As String
    Dim nLuggageType As Integer
    Dim rsTemp As Recordset
    atPriceItems = poAcceptSheet.PriceItems
    nLuggageType = GetLuggageTypeInt(poAcceptSheet.AcceptType)
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    szSql = "SELECT * FROM luggage_accept_info WHERE luggage_id=" & TransFieldValueToString(poAcceptSheet.SheetID)
    Set rsTemp = oDb.Execute(szSql)
    '如无此行包单　则出错
    If rsTemp.RecordCount <> 0 Then
        ShowError ERR_LuggageSvrHasExist  '行包单已存在
    End If
    oDb.BeginTrans
        'INSERT 受理信息表
        
        szSql = "INSERT luggage_accept_info (luggage_id,accept_type,start_station_id,start_station_name, " _
        & " des_station_id,des_station_name,mileage,bus_id,bus_date,luggage_name,cal_weight,fact_weight,start_label_id," _
        & " end_label_id,baggage_number,over_weight_number,status,shipper,shipper_phone,pick_type,picker,picker_address," _
        & " picker_phone,user_id,operation_time,pack,license_tag_no,trans_ticket_id , insurance_id , settle_ratio , settle_price , annotation1 , annotation2, sell_station_id ) " _
        & " VALUES ( " _
        & TransFieldValueToString(poAcceptSheet.SheetID) & "," & TransFieldValueToString(nLuggageType) & "," _
        & TransFieldValueToString(poAcceptSheet.StartStationID) & "," & TransFieldValueToString(poAcceptSheet.StartStationName) & "," _
        & TransFieldValueToString(poAcceptSheet.DesStationID) & "," & TransFieldValueToString(poAcceptSheet.DesStationName) & "," _
        & TransFieldValueToString(poAcceptSheet.Mileage) & "," & TransFieldValueToString(poAcceptSheet.BusID) & "," _
        & TransFieldValueToString(ToDBDate(poAcceptSheet.BusDate)) & "," & TransFieldValueToString(poAcceptSheet.LuggageName) & "," _
        & TransFieldValueToString(poAcceptSheet.CalWeight) & "," _
        & TransFieldValueToString(poAcceptSheet.ActWeight) & "," & TransFieldValueToString(poAcceptSheet.StartLabelID) & "," _
        & TransFieldValueToString(poAcceptSheet.EndLabelID) & "," & TransFieldValueToString(poAcceptSheet.Number) & "," _
        & TransFieldValueToString(poAcceptSheet.OverNumber) & "," & TransFieldValueToString(ST_LuggageAcceptNormal) & "," _
        & TransFieldValueToString(poAcceptSheet.Shipper) & "," & TransFieldValueToString(poAcceptSheet.LuggageShipperPhone) & "," _
        & TransFieldValueToString(poAcceptSheet.PickType) & "," _
        & TransFieldValueToString(poAcceptSheet.Picker) & "," & TransFieldValueToString(poAcceptSheet.PickerAddress) & "," _
        & TransFieldValueToString(poAcceptSheet.PickerPhone) & "," & TransFieldValueToString(m_oActiveUser.UserID) & "," _
        & TransFieldValueToString(ToDBDateTime(Now)) & "," _
        & TransFieldValueToString(poAcceptSheet.Pack) & "," & TransFieldValueToString(poAcceptSheet.LicenseTagNo) & "," _
        & TransFieldValueToString(poAcceptSheet.TransTicketID) & "," _
        & TransFieldValueToString(poAcceptSheet.InsuranceID) & "," _
        & TransFieldValueToString(poAcceptSheet.SettleRatio) & "," _
        & TransFieldValueToString(poAcceptSheet.SettlePrice) & " ," _
        & TransFieldValueToString(poAcceptSheet.Annotation1) & ", "
        
        szSql = szSql & TransFieldValueToString(poAcceptSheet.Annotation2) & " ," _
                          & TransFieldValueToString(poAcceptSheet.SellStationID) & ")"
        
        oDb.Execute szSql
        
        'INSERT 行包票价明细表
        '    Count = ArrayLength(poAcceptSheet.PriceItems)
        szTempPriceSql = ""
        nCountPriceID = ArrayLength(atPriceItems)
        szTempPriceItem = ""
        szTempPriceID = "000"
        szTempPriceID1 = "001"
        For i = 1 To nCountPriceID
            For j = 1 To 10
                If atPriceItems(i).PriceID = szTempPriceID + Trim(Str(j - 1)) Then '1000 暂时定义
                    If i <> nCountPriceID Then
                        szTempPriceSql = szTempPriceSql + TransFieldValueToString(atPriceItems(i).PriceValue) + ","
                        szTempPriceItem = szTempPriceItem + "price_item_" + Trim(Str(i)) + ","
                    Else
                        szTempPriceSql = szTempPriceSql + TransFieldValueToString(atPriceItems(i).PriceValue)
                        szTempPriceItem = szTempPriceItem + "price_item_" + Trim(Str(i))
                    End If
                    Exit For
                End If
                If atPriceItems(i).PriceID = szTempPriceID1 + Trim(Str(j - 1)) Then '1000 暂时定义
                    If i <> nCountPriceID Then
                        szTempPriceSql = szTempPriceSql + TransFieldValueToString(atPriceItems(i).PriceValue) + ","
                        szTempPriceItem = szTempPriceItem + "price_item_" + Trim(Str(i)) + ","
                    Else
                        szTempPriceSql = szTempPriceSql + TransFieldValueToString(atPriceItems(i).PriceValue)
                        szTempPriceItem = szTempPriceItem + "price_item_" + Trim(Str(i))
                    End If
                    Exit For
                
                End If
            Next j
        Next i
        
        szTempPriceSql = TransFieldValueToString(poAcceptSheet.SheetID) & "," & TransFieldValueToString(poAcceptSheet.TotalPrice) & "," & szTempPriceSql
        szSql = " INSERT luggage_price_lst(luggage_id, price_total," & szTempPriceItem & " ,cal_base_price ) " _
            & " VALUES (" & szTempPriceSql & ", " & TransFieldValueToString(poAcceptSheet.CalBasePrice) & ")"
        oDb.Execute szSql

    oDb.CommitTrans
    
    Set oDb = Nothing
    Exit Sub
AcceptLuggageErr:
    oDb.RollbackTrans
    Call RaiseError(err.Number, "LuggageSvr:AcceptLuggage Method", err.Description)
End Sub

'查询受理单信息,返回受理单记录集
'##ModelId=3E03C2AF020B
Public Function GetAcceptSheetRS(Optional StartTime As Date = cszEmptyDateStr, Optional EndTime As Date = cszEmptyDateStr, Optional Status As Integer = -1, Optional LikeSheetID As String, Optional ToStationID As String, Optional CarrySheetID As String = "", Optional bIsDiferent As Boolean = False, Optional bOrderByDate As Boolean = False, Optional LikeTransTicketID As String, Optional LikeShipper As String, Optional TransportCompanyID As String, Optional RouteID As String) As ADODB.Recordset
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
    
    On Error GoTo GetAcceptSheetRSErr
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szTempSql As String
    Dim szTempFromSql As String
    
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szTempSql = ""
    If StartTime <> "1900-1-1" Then
        szTempSql = " AND l.operation_time >= " & TransFieldValueToString(StartTime) & "  "
    End If
    
    If EndTime <> "1900-1-1" Then
        szTempSql = szTempSql & " AND l.operation_time <= " & TransFieldValueToString(EndTime) & "  "
    End If
    
    If Status <> -1 Then
        szTempSql = szTempSql & " AND l.status = " & TransFieldValueToString(Status) & " "
    End If
    
    If LikeSheetID <> "" Then
        szTempSql = szTempSql & " AND l.luggage_id LIKE '%" & LikeSheetID & "%'  "
    End If
    
    If ToStationID <> "" Then
        szTempSql = szTempSql & " AND l.des_station_id LIKE '%" & ToStationID & "%'  "
    End If
    
    If LikeTransTicketID <> "" Then
        szTempSql = szTempSql & " AND l.trans_ticket_id LIKE '%" & LikeTransTicketID & "%'  "
    End If
    
    If LikeShipper <> "" Then
        szTempSql = szTempSql & " AND l.shipper LIKE '%" & LikeShipper & "%'  "
    End If
    
    If CarrySheetID <> "" Then
        szTempFromSql = " , luggage_carry_accept_lst c"
        szTempSql = szTempSql & " AND c.luggage_id=l.luggage_id AND c.check_luggage_id =" & TransFieldValueToString(CarrySheetID)
    Else
        szTempFromSql = szTempFromSql + ""
    End If
    
    If TransportCompanyID <> "" Then
        szTempFromSql = " , luggage_carry_accept_lst c,  luggage_sheet_info t "
        szTempSql = szTempSql & " AND c.luggage_id=l.luggage_id AND t.sheet_id = c.check_luggage_id AND t.transport_company_id =" & TransFieldValueToString(TransportCompanyID)
    Else
        szTempFromSql = szTempFromSql + ""
    End If
    
    If RouteID <> "" Then
        szTempFromSql = " , luggage_carry_accept_lst c,  luggage_sheet_info t "
        szTempSql = szTempSql & " AND c.luggage_id=l.luggage_id AND t.sheet_id = c.check_luggage_id AND t.route_id =" & TransFieldValueToString(RouteID)
    Else
        szTempFromSql = szTempFromSql + ""
    End If
    
    If bIsDiferent <> False Then
        szTempSql = szTempSql & " AND p.cal_base_price <>  p.price_item_1 "
    End If
    
    If m_oActiveUser.SellStationID <> "" Then
        szTempSql = szTempSql & " AND l.sell_station_id = " & TransFieldValueToString(m_oActiveUser.SellStationID)
    End If
    
    
    '如果工作表中数据还在,则从工作表中取,以提高速度
    Dim oParam As New SystemParam
    Dim nPreDay As Integer
    Dim szTableEnvBus As String
    Dim szTableEnvBusAllot As String
    
    oParam.Init m_oActiveUser
    nPreDay = oParam.PreSellDate '预售天数
    If StartTime <> cszEmptyDateStr And DateAdd("d", nPreDay, StartTime) > oParam.NowDate Then
        szTableEnvBus = cszTableEnvBusWork '从工作表中取"
        szTableEnvBusAllot = cszTableEnvAllotWork
    Else
        szTableEnvBus = cszTableEnvBusInfo
        szTableEnvBusAllot = cszTableEnvAllotInfo
    End If
    
    szSql = "SELECT a.bus_start_time,l.*, f.cancel_time, f.user_id cancel_user_id " _
        & " , CASE l.accept_type WHEN " & TransFieldValueToString(ELuggageAcceptType.NormalAccept) & " THEN " & TransFieldValueToString(cszLuggageCarryAcceptName) _
        & " WHEN " & TransFieldValueToString(ELuggageAcceptType.CarryAccept) & " THEN " & TransFieldValueToString(cszLuggageNormalAcceptName) & " END accept_type_name " _
        & " , CASE l.status WHEN 0 THEN '未签发' WHEN 1 THEN '废票' WHEN 2 THEN '退票' WHEN 3 THEN '已签发' END status_name " _
        & " ,p.price_total,p.price_item_1 , p.price_item_2 , p.price_item_3 , p.price_item_4 , p.price_item_5 " _
        & " ,p.price_item_6 , p.price_item_7 , p.price_item_8 , p.price_item_9 , p.price_item_10 , p.cal_base_price " _
        & " FROM luggage_accept_info l,luggage_price_lst p, luggage_cancel_lst f, " & szTableEnvBus & " e, " & szTableEnvBusAllot & " a " & szTempFromSql _
        & " WHERE l.luggage_id=p.luggage_id AND e.bus_id=l.bus_id AND l.luggage_id *= f.luggage_id  AND datediff(day,e.bus_date,l.bus_date )=0  " _
        & " AND e.bus_id = a.bus_id AND e.bus_date = a.bus_date AND a.sell_station_id = " & TransFieldValueToString(m_oActiveUser.SellStationID) _
        & szTempSql
    If bOrderByDate Then
        szSql = szSql & " ORDER BY e.bus_start_time , l.bus_id, l.luggage_id"
    Else
    
        szSql = szSql & " ORDER BY l.des_station_name,l.luggage_id "
    End If
    Set rsTemp = oDb.Execute(szSql)
    Set GetAcceptSheetRS = rsTemp
    

    Exit Function
GetAcceptSheetRSErr:
    Call RaiseError(err.Number, "LuggageSvr:GetAcceptSheetRS Method", err.Description)
End Function

Public Function GetCarrySheetRS(Optional StartTime As Date = "1900-1-1", Optional EndTime As Date = "1900-1-1", Optional LikeSheetID As String, Optional BusID As String, Optional Status As Integer = -1, Optional LikeWorker As String) As ADODB.Recordset
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
    
    On Error GoTo GetCarrySheetRSErr
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szTempSql As String
    Dim szTempFromSql As String
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szTempFromSql = ""
    szTempSql = ""
    If StartTime <> "1900-1-1" Then
       szTempSql = szTempSql & " AND l.sheet_make_time >= " & TransFieldValueToString(StartTime)
    End If
    
    If EndTime <> "1900-1-1" Then
       szTempSql = szTempSql & " AND l.sheet_make_time <= " & TransFieldValueToString(EndTime)
    End If
    
    If LikeSheetID <> "" Then
        szTempSql = szTempSql & " AND l.sheet_id LIKE '%" & LikeSheetID & "%' "
    End If
    If LikeWorker <> "" Then
        szTempSql = szTempSql & " AND l.worker LIKE '%" & LikeWorker & "%' "
    End If
    
    If BusID <> "" Then
        szTempSql = szTempSql + " AND l.license_tag_no ='" & BusID & "'"
    End If
    If Status <> -1 Then
        szTempSql = szTempSql + " AND l.valid_mark = " & TransFieldValueToString(Status)
    End If
    
    If m_oActiveUser.SellStationID <> "" Then
        szTempSql = szTempSql & " AND l.sell_station_id = " & TransFieldValueToString(m_oActiveUser.SellStationID)
    End If
    
    szSql = "SELECT l.* " _
        & " FROM luggage_sheet_info l " _
        & " WHERE 1=1 " & szTempSql _
        & " ORDER BY sheet_id "
    Set rsTemp = oDb.Execute(szSql)
    Set GetCarrySheetRS = rsTemp

    Exit Function
GetCarrySheetRSErr:
    Call RaiseError(err.Number, "LuggageSvr:GetCarrySheetRS Method", err.Description)
End Function

'根据受理单查询行包签发单信息,返回签发单记录集
Public Function GetCarrySheetRSEX(LikeSheetID As String) As ADODB.Recordset
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
    
    On Error GoTo GetCarrySheetRSErr
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szTempSql As String
    Dim szTempFromSql As String
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    If LikeSheetID <> "" Then
    
        If m_oActiveUser.SellStationID <> "" Then
            szSql = " AND l.sell_station_id = " & TransFieldValueToString(m_oActiveUser.SellStationID)
        End If
   
        szSql = "SELECT l.* " _
            & " FROM luggage_carry_accept_lst ca,luggage_accept_info a, luggage_sheet_info l " _
            & " WHERE ca.luggage_id=a.luggage_id and ca.check_luggage_id=l.sheet_id and a.luggage_id='" & LikeSheetID & "'" _
            & szSql & " ORDER BY sheet_id "
        Set rsTemp = oDb.Execute(szSql)
        Set GetCarrySheetRSEX = rsTemp
    'your code goes here...
    End If
    Exit Function
GetCarrySheetRSErr:
    Call RaiseError(err.Number, "LuggageSvr:GetCarrySheetRS Method", err.Description)
End Function

'得到最新的行包单号
Public Function GetLastLuggageID(UserID As String) As String
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    szSql = "SELECT MAX(luggage_id) AS luggage_id FROM luggage_accept_info " _
    & " WHERE operation_time=(" _
    & " SELECT MAX(operation_time) FROM luggage_accept_info WHERE " _
    & " user_id='" & UserID & "') " _
    & " AND user_id='" & UserID & "' "
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 1 Then
        GetLastLuggageID = FormatDbValue(rsTemp!luggage_id)
    Else
        Dim oParam As New STLuggage.LuggageParam
        oParam.Init m_oActiveUser
        GetLastLuggageID = String(oParam.LuggageIDNumberLen, "0")
    End If
    Set oDb = Nothing
    
End Function

'得到最新的行包签发单号
Public Function GetLastSheetID(UserID As String) As String
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    szSql = "SELECT MAX(sheet_id) AS sheet_id FROM luggage_sheet_info " _
    & " WHERE sheet_make_time=(" _
    & " SELECT MAX(sheet_make_time) FROM luggage_sheet_info WHERE " _
    & " checker='" & UserID & "') " _
    & " AND checker='" & UserID & "' "
    
'    szSql = "SELECT MAX(sheet_id) AS sheet_id FROM luggage_sheet_info " _
'    & " WHERE sheet_make_time=(" _
'    & " SELECT MAX(sheet_make_time) FROM luggage_sheet_info) "
    
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 1 Then
        GetLastSheetID = FormatDbValue(rsTemp!sheet_id)
    Else
        Dim oParam As New STLuggage.LuggageParam
        oParam.Init m_oActiveUser
        GetLastSheetID = String(oParam.CarrySheetIDNumberLen, "0")
    End If
    Set oDb = Nothing
End Function

'得到所有的站点信息
'参见SellTicketClient.GetAllStationRS方法
'##ModelId=3E03E085010A
Public Function GetAllStationRS() As ADODB.Recordset
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
    
    On Error GoTo GetAllStationRSErr
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    szSql = "SELECT station_id,station_name ,station_input_code  FROM station_info  ORDER BY station_id"
    Set rsTemp = oDb.Execute(szSql)
    Set GetAllStationRS = rsTemp
    'your code goes here...

    Exit Function
GetAllStationRSErr:
    Call RaiseError(err.Number, "LuggageSvr:GetAllStationRS Method", err.Description)
End Function

'得到到达指定站点的车次信息
'
'##ModelId=3E03E166028A
Public Function GetToStationBusRS(ToStationID As String, BusDate As Date) As ADODB.Recordset
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
    
    On Error GoTo GetToStationBusRSErr
    
    Dim szSql As String
    Dim szTemp As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    Dim szTables As String
    If BusDate >= Date Then       '工作表
        szTables = "" & cszTableEnvBusStationWork & " e, " & cszTableEnvBusWork & " b , " & cszTableVehicle & " t  , " & cszTableCompany & " c , " & cszTableEnvAllotWork & " a "
    Else
        szTables = "" & cszTableEnvBusStationInfo & " e, " & cszTableEnvBusInfo & " b, " & cszTableVehicle & " t  ,  " & cszTableCompany & " c , " & cszTableEnvAllotInfo & " a "
    End If
    
    szSql = "SELECT DISTINCT t.license_tag_no,e.*,convert(char(20),a.bus_start_time,108) bus_start_time,b.vehicle_type_code, b.vehicle_type_name,b.end_station_name , b.bus_type , b.transport_company_id , c.transport_company_short_name " _
            & " FROM " & szTables _
            & " WHERE e.station_id = " & TransFieldValueToString(ToStationID) _
            & " AND b.bus_date=" & TransFieldValueToString(BusDate) _
            & " AND e.bus_date = " & TransFieldValueToString(BusDate) _
            & " AND e.bus_id = b.bus_id AND a.sell_station_id = " & TransFieldValueToString(m_oActiveUser.SellStationID) _
            & " AND a.bus_id = e.bus_id AND a.bus_date = e.bus_date AND a.sell_station_id = e.sell_station_id " _
            & " AND ( a.bus_start_time > " & TransFieldValueToString(ToDBDateTime(Now)) & " OR a.bus_start_time=" & TransFieldValueToString(BusDate) & ")" _
            & " AND b.status not in(1,2,8) AND e.max_sale_quantity<>0 AND a.can_sale_quantity <> 0 " _
            & " AND t.vehicle_id=b.vehicle_id AND b.transport_company_id = c.transport_company_id ORDER BY b.bus_type,convert(char(20),a.bus_start_time,108)"
            
            
    Set rsTemp = oDb.Execute(szSql)
    Set GetToStationBusRS = rsTemp
    Set oDb = Nothing
    'your code goes here...

    Exit Function
GetToStationBusRSErr:
    Call RaiseError(err.Number, "LuggageSvr:GetToStationBusRS Method", err.Description)
End Function

'显示指定车次的当前车次
Public Function GetBusVehicle(szBusID As String) As ADODB.Recordset
    On Error GoTo Here:
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT DISTINCT t.license_tag_no,t.vehicle_id FROM " & cszTableBusVehicle & " b," & cszTableVehicle & " t WHERE b.vehicle_id=t.vehicle_id AND " _
            & " b.bus_id= " & TransFieldValueToString(szBusID)
    Set rsTemp = oDb.Execute(szSql)
    Set GetBusVehicle = rsTemp
    
    Set rsTemp = Nothing
  Exit Function
  Set rsTemp = Nothing
Here:
End Function


'作废行包受理单
'##ModelId=3E03E72C02CB
Public Sub CancelAcceptSheet(SheetID As String)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
    AssertHaveRight m_oActiveUser, RIGHT_LuggageSvrCancel
    
    On Error GoTo CancelAcceptSheetErr
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
'
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    szSql = " SELECT * FROM luggage_accept_info WHERE luggage_id = " & TransFieldValueToString(SheetID)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_LuggageSvrHaveNo '出错
    
    If FormatDbValue(rsTemp!Status) = 2 Then
        ShowError ERR_LuggageSvrSheetCanceled  '行包单已废
    ElseIf FormatDbValue(rsTemp!Status) = 3 Then
        ShowError ERR_LuggageSvrSheetReturned '行包单已退
    End If
    
    
    szSql = " SELECT 1 FROM luggage_cancel_lst WHERE luggage_id=" & TransFieldValueToString(SheetID)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 0 Then ShowError ERR_LuggageSvrSheetCanceled
    
    
    oDb.BeginTrans
    
    '修改废行包票明细表
    szSql = "INSERT luggage_cancel_lst (luggage_id,user_id,cancel_time) " _
        & " SELECT luggage_id,user_id ,getdate() FROM luggage_accept_info " _
        & " WHERE luggage_id = " & TransFieldValueToString(SheetID)
    oDb.Execute szSql
    
    '修改受理行包信息表的行包单状态
    szSql = "UPDATE luggage_accept_info SET status = " & TransFieldValueToString(ST_LuggageAcceptCancel) _
        & " WHERE luggage_id = " & TransFieldValueToString(SheetID)
    oDb.Execute szSql
   
    oDb.CommitTrans
    
    Set oDb = Nothing
    'your code goes here...

    Exit Sub
CancelAcceptSheetErr:
    oDb.RollbackTrans
    Call RaiseError(err.Number, "LuggageSvr:CancelAcceptSheet Method", err.Description)
End Sub

'行包退运
'##ModelId=3E03E79C01DC
Public Sub ReturnAcceptSheet(SheetID As String, CredenceID As String, ReturnCharge As Double)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
    AssertHaveRight m_oActiveUser, RIGHT_LuggageSvrCancel
    
    On Error GoTo ReturnAcceptSheetErr
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    'Dim szUserID As String '退票员ID
    
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT * FROM luggage_accept_info WHERE luggage_id = " & TransFieldValueToString(SheetID)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then
        ShowError ERR_LuggageSvrHaveNo
    Else
        If FormatDbValue(rsTemp!Status) = 2 Then
            ShowError ERR_LuggageSvrSheetCanceled '行包单已废票
        ElseIf FormatDbValue(rsTemp!Status) = 3 Then
            ShowError ERR_LuggageSvrSheetReturned '行包单已退票
        ElseIf FormatDbValue(rsTemp!Status) = 4 Then
            ShowError ERR_LuggageSvrSheetCarryed '行包单已签发
        End If
    End If
    
    szSql = "SELECT 1 FROM luggage_return_lst WHERE luggage_id = " & TransFieldValueToString(SheetID)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 0 Then ShowError ERR_LuggageSvrSheetReturned
    
    oDb.BeginTrans
    'INSERT 退行包票明细表
    szSql = "INSERT luggage_return_lst(luggage_id,credence_sheet_id,return_charge,user_id,operation_time)" _
        & " VALUES ( " _
        & TransFieldValueToString(SheetID) & "," _
        & TransFieldValueToString(CredenceID) & "," _
        & TransFieldValueToString(ReturnCharge) & "," _
        & TransFieldValueToString(m_oActiveUser.UserID) & "," _
        & TransFieldValueToString(ToDBDateTime(Now)) & ")"
    
    oDb.Execute szSql
    'UPDATE 行包受理信息表 的　行包单状态
    szSql = "UPDATE luggage_accept_info SET status =" & TransFieldValueToString(ST_LuggageAcceptReturn) _
        & " WHERE luggage_id = " & TransFieldValueToString(SheetID)
                
    oDb.Execute szSql
    
    oDb.CommitTrans
    Set oDb = Nothing
    
    
    'your code goes here...

    Exit Sub
ReturnAcceptSheetErr:
    oDb.RollbackTrans
    Call RaiseError(err.Number, "LuggageSvr:ReturnAcceptSheet Method", err.Description)
End Sub

'统计指定用户的行包托运信息
'##ModelId=3E03EDF501D8
Public Function TotalAcceptRS(UserID() As String, StartTime As Date, EndTime As Date) As ADODB.Recordset
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
    AssertHaveRight m_oActiveUser, RIGHT_LuggageSvrTotalAccept
    
    On Error GoTo TotalAcceptRSErr
    
    Dim szSql As String
    Dim szUsers As String
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr
    
    Dim rsData As New Recordset
    rsData.CursorLocation = adUseClient
    
    If ArrayLength(UserID) Then
        szUsers = " AND u.user_id IN("
        Dim i As Integer
        For i = 1 To ArrayLength(UserID) - 1
            szUsers = szUsers & TransFieldValueToString(UserID(i)) & ","
        Next i
        szUsers = szUsers & TransFieldValueToString(UserID(i)) & ")"
    End If
    
    Dim rsTmp As Recordset
    '得到正常售票
    szSql = "SELECT RTRIM(u.user_id) +'['+RTRIM(MAX(user_name))+']' AS user_name ,(CASE accept_type WHEN 0 THEN '" & GetLuggageTypeString(0) & "' WHEN 1 THEN '" & GetLuggageTypeString(1) & "' END) accept_type " & _
            ",SUM(cal_weight) cal_weight,SUM(fact_weight) fact_weight,SUM(over_weight_number) over_weight_number,SUM(baggage_number) baggage_number" & _
            ",SUM(price_total) price_total,SUM(price_item_1) price_item_1,SUM(price_item_2) price_item_2,SUM(price_item_3) price_item_3" & _
            ",SUM(price_item_4) price_item_4,SUM(price_item_5) price_item_5,SUM(price_item_6) price_item_6,SUM(price_item_7) price_item_7" & _
            ",SUM(price_item_8) price_item_8,SUM(price_item_9) price_item_9,SUM(price_item_10) price_item_10" & _
            ",SUM(1) sheet_number,0 cancel_number, 0 cancel_price,0 return_number,0 return_price,0 return_charge" & _
            " FROM luggage_accept_info l,luggage_price_lst p,user_info u" & _
            " WHERE p.luggage_id=l.luggage_id AND l.user_id=u.user_id" & szUsers & _
            " AND operation_time>='" & ToDBDateTime(StartTime) & "' AND operation_time<='" & ToDBDateTime(EndTime) & "'" & _
            " GROUP BY u.user_id,accept_type"
    Set rsTmp = oDb.Execute(szSql)
    
    
    '创建rsData，复制数据
    Dim oField As Field
    For Each oField In rsTmp.Fields
        If oField.Type = adNumeric Then
            rsData.Fields.Append oField.Name, adDouble
        Else
            rsData.Fields.Append oField.Name, oField.Type, oField.DefinedSize
        End If
    Next
    rsData.Open
    While Not rsTmp.EOF
        rsData.AddNew
        For i = 0 To rsData.Fields.Count - 1
            If Not IsNull(rsTmp.Fields(i).Value) Then
                rsData.Fields(i).Value = rsTmp.Fields(i).Value
            End If
            
        Next
        rsData.Update
        rsTmp.MoveNext
    Wend
    
    '得到废票,并更改返回的记录集
    szSql = "SELECT RTRIM(u.user_id) +'['+RTRIM(MAX(user_name))+']' AS user_name,(CASE accept_type WHEN 0 THEN '" & GetLuggageTypeString(0) & "' WHEN 1 THEN '" & GetLuggageTypeString(1) & "' END) accept_type " & _
            ",SUM(1) cancel_number, SUM(price_total) cancel_price" & _
            " FROM luggage_accept_info l,luggage_price_lst p,luggage_cancel_lst c,user_info u" & _
            " WHERE p.luggage_id=l.luggage_id AND c.luggage_id=l.luggage_id AND c.user_id=u.user_id" & szUsers & _
            " AND cancel_time>='" & ToDBDateTime(StartTime) & "' AND cancel_time<='" & ToDBDateTime(EndTime) & "' " & _
            " GROUP BY u.user_id,accept_type"
    Set rsTmp = oDb.Execute(szSql)
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    If rsData.RecordCount > 0 Then rsData.MoveFirst
    Do While Not rsTmp.EOF     '将对应用户的返回记录集的废票进行更改
        Do While Not rsData.EOF
            If rsTmp!user_name = rsData!user_name And rsTmp!Accept_type = rsData!Accept_type Then
                rsData!cancel_number = rsTmp!cancel_number
                rsData!cancel_price = rsTmp!cancel_price
                rsData.Update
                Exit Do
            End If
            rsData.MoveNext
        Loop
        If rsData.EOF Then      '废票记录集中有,返回记录集中没有,则添加一行
            rsData.AddNew
            rsData!user_name = rsTmp!user_name
            rsData!Accept_type = rsTmp!Accept_type
            rsData!cancel_number = rsTmp!cancel_number
            rsData!cancel_price = rsTmp!cancel_price
            rsData!cal_weight = 0: rsData!fact_weight = 0: rsData!over_weight_number = 0: rsData!baggage_number = 0
            rsData!return_number = 0: rsData!return_price = 0: rsData!return_charge = 0
            rsData!price_total = 0
            For i = 1 To 10
                rsData.Fields("price_item_" & i) = 0
            Next
            rsData.Update
        End If
        rsTmp.MoveNext
    Loop

    '得到退票,并更改返回的记录集
    szSql = "SELECT RTRIM(u.user_id) +'['+RTRIM(MAX(user_name))+']' AS user_name,(CASE accept_type WHEN 0 THEN '" & GetLuggageTypeString(0) & "' WHEN 1 THEN '" & GetLuggageTypeString(1) & "' END) accept_type " & _
            ",SUM(1) return_number, SUM(price_total) return_price,SUM(return_charge) return_charge" & _
            " FROM luggage_accept_info l,luggage_price_lst p,luggage_return_lst c,user_info u" & _
            " WHERE p.luggage_id=l.luggage_id AND c.luggage_id=l.luggage_id AND c.user_id=u.user_id" & szUsers & _
            " AND c.operation_time>='" & ToDBDateTime(StartTime) & "' AND c.operation_time<='" & ToDBDateTime(EndTime) & "' " & _
            " GROUP BY u.user_id,accept_type"
    Set rsTmp = oDb.Execute(szSql)
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Do While Not rsTmp.EOF     '将对应用户的返回记录集的废票进行更改
        Do While Not rsData.EOF
            If rsTmp!user_name = rsData!user_name And rsTmp!Accept_type = rsData!Accept_type Then
                rsData!return_number = rsTmp!return_number
                rsData!return_price = rsTmp!return_price
                rsData!return_charge = rsTmp!return_charge
                rsData.Update
                Exit Do
            End If
            rsData.MoveNext
        Loop
        If rsData.EOF Then      '退票记录集中有,返回记录集中没有,则添加一行
            rsData.AddNew
            rsData!user_name = rsTmp!user_name
            rsData!Accept_type = rsTmp!Accept_type
            rsData!return_number = rsTmp!return_number
            rsData!return_price = rsTmp!return_price
            rsData!return_charge = rsTmp!return_charge
            rsData!cal_weight = 0: rsData!fact_weight = 0: rsData!over_weight_number = 0: rsData!baggage_number = 0
            rsData!cancel_number = 0
            rsData!cancel_price = 0
            rsData!price_total = 0
            For i = 1 To 10
                rsData.Fields("price_item_" & Trim(Str(i))) = 0
            Next
            rsData.Update
        End If
        rsTmp.MoveNext
    Loop

    If rsData.RecordCount > 0 Then rsData.MoveFirst
    Set TotalAcceptRS = rsData
    Set oDb = Nothing

    Exit Function
TotalAcceptRSErr:
    Call RaiseError(err.Number, "LuggageSvr:TotalAcceptRS Method", err.Description)
End Function

'行包签发
'
'##ModelId=3E04035A03DD
Public Sub CarryLuggage(poCarrySheet As CarrySheet, mAcceptType As Integer)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
    AssertHaveRight m_oActiveUser, RIGHT_LuggageSvrNormalCarry
    
    
    On Error GoTo CarryLuggageErr
    
    Dim szSql As String
    Dim szTempPriceSql As String
    Dim szTempSql As String
    Dim oDb As New RTConnection
    Dim i As Integer, j As Integer
    Dim atPriceItems() As TLuggagePriceItem
    Dim atAcceptSheetIDs() As String
    Dim nCountPriceID As Integer
    Dim szTempPriceID As String
    Dim szTempPriceItem As String
    Dim Count As Integer
    Dim rsTemp As Recordset
    
    Dim nCount As Integer
    Dim dbCalWeight As Double
    Dim dbActWeight As Double
    Dim nNumber As Integer
    Dim nOverNumber As Integer
    Dim nAcceptSheetNumber As Integer
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    szSql = "SELECT * FROM luggage_sheet_info WHERE sheet_id=" & TransFieldValueToString(poCarrySheet.SheetID)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 0 Then
        If FormatDbValue(rsTemp!valid_mark) = 0 Then
            ShowError ERR_LuggageSvrCarrySheetCancel '签发单已废
        Else
            ShowError ERR_LuggageSvrCarrySheetHave '签发单已存在
        End If
    End If
    
    
    atPriceItems = poCarrySheet.PriceItems
    atAcceptSheetIDs = poCarrySheet.AcceptSheetIDs
    oDb.BeginTrans
        '得到票价信息
        
        szTempPriceSql = ""
        nCountPriceID = ArrayLength(atPriceItems)
        szTempPriceItem = ""
        szTempPriceID = "000"
        
        For i = 1 To nCountPriceID
            For j = 1 To 10
                If atPriceItems(i).PriceID = Trim(szTempPriceID & Trim(Str(i - 1))) Then '1000 暂时定义
                        szTempPriceSql = szTempPriceSql + TransFieldValueToString(atPriceItems(i).PriceValue) + ","
                        szTempPriceItem = szTempPriceItem + "price_item_" + Trim(Str(i)) + ","
                    Exit For
                End If
                If atPriceItems(i).PriceID = Trim("001" & Trim(Str(i - 1))) Then '1000 暂时定义
                        szTempPriceSql = szTempPriceSql + TransFieldValueToString(atPriceItems(i).PriceValue) + ","
                        szTempPriceItem = szTempPriceItem + "price_item_" + Trim(Str(i)) + ","
                    Exit For
                End If
            Next j
        Next i
        
        
        szTempPriceSql = TransFieldValueToString(poCarrySheet.TotalPrice) + "," + szTempPriceSql + TransFieldValueToString(ST_LuggageCarrySettlementNo)
        'INSERT 签发单信息表
        szSql = "INSERT luggage_sheet_info(sheet_id,sell_station_id,accept_type,bus_id,bus_date,vehicle_id,license_tag_no,transport_company_id,transport_company_short_name," _
            & " split_company_id,splict_company_short_name,route_id,route_name,vehicle_type_code,vehicle_type_name,owner_id, " _
            & " owner_name,bus_start_time,valid_mark,worker,protocol_id,protocol_name,checker,sheet_make_time,luggage_number,baggage_number," _
            & " cal_weight,fact_weight,over_number,total_price," & szTempPriceItem & " settlement_status , print_settle_ratio , print_settle_price ) " _
            & " VALUES ( " _
            & TransFieldValueToString(poCarrySheet.SheetID) & "," & TransFieldValueToString(poCarrySheet.SellStationID) & "," _
            & TransFieldValueToString(mAcceptType) & "," & TransFieldValueToString(poCarrySheet.BusID) & "," _
            & TransFieldValueToString(ToDBDate(poCarrySheet.BusDate)) & "," _
            & TransFieldValueToString(poCarrySheet.VehicleID) & "," & TransFieldValueToString(poCarrySheet.VehicleLicense) & "," _
            & TransFieldValueToString(poCarrySheet.CompanyID) & "," & TransFieldValueToString(poCarrySheet.CompanyName) & "," _
            & TransFieldValueToString(poCarrySheet.SplitCompanyID) & "," & TransFieldValueToString(poCarrySheet.SplitCompanyName) & "," _
            & TransFieldValueToString(poCarrySheet.RouteID) & "," & TransFieldValueToString(poCarrySheet.RouteName) & "," _
            & TransFieldValueToString(poCarrySheet.VehicleTypeID) & "," & TransFieldValueToString(poCarrySheet.VehicleTypeName) & "," _
            & TransFieldValueToString(poCarrySheet.BusOwnerID) & "," & TransFieldValueToString(poCarrySheet.BusOwnerName) & "," _
            & TransFieldValueToString(ToDBDateTime(poCarrySheet.BusStartOffTime)) & "," & TransFieldValueToString(1) & "," _
            & TransFieldValueToString(poCarrySheet.MoveWorker) & " , " _
            & TransFieldValueToString(poCarrySheet.ProtocolID) & "," & TransFieldValueToString(poCarrySheet.ProtocolName) & "," _
            & TransFieldValueToString(poCarrySheet.OperatorID) & "," & TransFieldValueToString(SelfNowDateTime) & "," _
            & TransFieldValueToString(poCarrySheet.AcceptSheetNumber) & "," & TransFieldValueToString(poCarrySheet.Number) & "," _
            & TransFieldValueToString(poCarrySheet.CalWeight) & "," & TransFieldValueToString(poCarrySheet.ActWeight) & "," _
            & TransFieldValueToString(poCarrySheet.OverNumber) & "," & szTempPriceSql & " , " _
            & TransFieldValueToString(poCarrySheet.PrintSettleRatio) & "," & TransFieldValueToString(poCarrySheet.PrintSettlePrice) & ")"
            
        oDb.Execute szSql
        'INSERT 行包签发明细表
        
        
        
        If poCarrySheet.AcceptSheetNumber <> 0 Then
            Count = ArrayLength(poCarrySheet.AcceptSheetIDs)
            For i = 1 To Count
                szSql = "INSERT luggage_carry_accept_lst(check_luggage_id,luggage_id) " _
                    & " VALUES ( " _
                    & TransFieldValueToString(poCarrySheet.SheetID) & "," _
                    & TransFieldValueToString(atAcceptSheetIDs(i)) & ")"
                oDb.Execute szSql
                
                
                
                
                '修改　行包受理信息表中的　行包单状态
                szSql = " UPDATE luggage_accept_info SET status = " & TransFieldValueToString(ST_LuggageAcceptOver) & "," _
                    & " settle_ratio = " & TransFieldValueToString(poCarrySheet.PrintSettleRatio / 100) & "," _
                    & " settle_price = settle_price/settle_ratio/100*" & Val(poCarrySheet.PrintSettleRatio) & "," _
                    & " bus_id = " & TransFieldValueToString(poCarrySheet.BusID) & "," _
                    & " bus_date = " & TransFieldValueToString(poCarrySheet.BusDate) & "," _
                    & " license_tag_no = " & TransFieldValueToString(poCarrySheet.VehicleLicense) _
                    & " WHERE luggage_id =" & TransFieldValueToString(atAcceptSheetIDs(i))
                
                oDb.Execute szSql
                
            Next i
        Else
            '不做插入工作
        End If
        '更新luggage_sheet_info表总的总计重，总实重等
    '    poCarrySheet.TotalCarryLuggage
        szSql = "SELECT COUNT(l.luggage_id) AS luggage_number,SUM(a.cal_weight) AS cal_weight,SUM(a.fact_weight) AS fact_weight,SUM(a.baggage_number) AS baggage_number, " _
            & " SUM(a.over_weight_number) AS over_weight_number " _
            & " FROM luggage_carry_accept_lst l , luggage_accept_info a" _
            & " WHERE l.luggage_id = a.luggage_id  AND check_luggage_id =" & TransFieldValueToString(poCarrySheet.SheetID) _
            & " GROUP BY l.check_luggage_id "
        
        
        Set rsTemp = oDb.Execute(szSql)
        nCount = rsTemp.RecordCount
        If nCount <> 0 Then
    '        mszSheetID = FormatDbValue(rsTemp!luggage_id)
            dbCalWeight = FormatDbValue(rsTemp!cal_weight)             '计重
            dbActWeight = FormatDbValue(rsTemp!fact_weight)            '实重
            nNumber = FormatDbValue(rsTemp!baggage_number)             '件数
            nOverNumber = FormatDbValue(rsTemp!over_weight_number)     '超重件数
            nAcceptSheetNumber = FormatDbValue(rsTemp!luggage_number)  '受理单数
        End If
        
        
    
        szSql = "UPDATE luggage_sheet_info SET baggage_number =" & TransFieldValueToString(nNumber) & "," _
            & " cal_weight =" & TransFieldValueToString(dbCalWeight) & "," _
            & " fact_weight = " & TransFieldValueToString(dbActWeight) & "," _
            & " over_number =" & TransFieldValueToString(nOverNumber) _
            & " WHERE sheet_id= " & TransFieldValueToString(poCarrySheet.SheetID)
        oDb.Execute szSql
        
    oDb.CommitTrans
    
    

    Exit Sub
CarryLuggageErr:
    oDb.RollbackTrans
    Call RaiseError(err.Number, "LuggageSvr:CarryLuggage Method", err.Description)
End Sub

'作废签发单
Public Sub CancelSheet(CarrySheet As String)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
    AssertHaveRight m_oActiveUser, RIGHT_LuggageSvrReprintCarry
    
    On Error GoTo CancelSheetErr
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    '判断是否有记录
    szSql = "SELECT * FROM luggage_sheet_info WHERE sheet_id =" & TransFieldValueToString(CarrySheet)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_LuggageSvrHaveNo
    
    If FormatDbValue(rsTemp!valid_mark) = ST_LuggageSheetValidStatusCancel Then
        ShowError ERR_LuggageSvrCarrySheetCancel '签发单已废
    End If
    
    oDb.BeginTrans
    
    szSql = "UPDATE luggage_sheet_info SET valid_mark =" & TransFieldValueToString(ST_LuggageSheetValidStatusCancel) _
        & " WHERE sheet_id =" & TransFieldValueToString(CarrySheet)
        
    oDb.Execute szSql
    
    szSql = "UPDATE luggage_accept_info SET status=" & TransFieldValueToString(ST_LuggageAcceptNormal) _
         & " WHERE luggage_id IN ( SELECT luggage_id FROM luggage_carry_accept_lst WHERE check_luggage_id=" & TransFieldValueToString(CarrySheet) & ")"
    oDb.Execute szSql
    
    
    szSql = "DELETE luggage_carry_accept_lst WHERE check_luggage_id=" & TransFieldValueToString(CarrySheet)
    oDb.Execute szSql
    
    oDb.CommitTrans
    Set oDb = Nothing
    
    Exit Sub
CancelSheetErr:
    oDb.RollbackTrans
    Call RaiseError(err.Number, "LuggageSvr:ReprintCarrySheet Method", err.Description)
End Sub

'重打行包签发单
'
'##ModelId=3E0403B2039D
Public Sub ReprintCarrySheet(CarrySheet As String, NewCarrySheet As String)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
    AssertHaveRight m_oActiveUser, RIGHT_LuggageSvrReprintCarry
    
    On Error GoTo ReprintCarrySheetErr
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    '判断是否有记录
    szSql = "SELECT * FROM luggage_sheet_info WHERE sheet_id =" & TransFieldValueToString(CarrySheet)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_LuggageSvrHaveNo
    
    If FormatDbValue(rsTemp!valid_mark) = 0 Then
        ShowError ERR_LuggageSvrCarrySheetCancel '签发单已废
    End If
    
    oDb.BeginTrans
    'INSERT　签发单信息表
    szSql = "INSERT luggage_sheet_info (sheet_id,accept_type,sell_station_id,bus_id,bus_date,vehicle_id,license_tag_no,transport_company_id,transport_company_short_name, " _
        & " split_company_id,splict_company_short_name,route_id,route_name,vehicle_type_name,vehicle_type_code,owner_id,owner_name,bus_start_time,valid_mark ,worker,protocol_id, " _
        & " protocol_name,checker,sheet_make_time,luggage_number,baggage_number,cal_weight,fact_weight,over_number,total_price,price_item_1,price_item_2,price_item_3,price_item_4, " _
        & " price_item_5,price_item_6,price_item_7,price_item_8,price_item_9,price_item_10,settlement_status,print_settle_ratio,print_settle_price) " _
        & " SELECT " & TransFieldValueToString(NewCarrySheet) & " AS sheet_id ,accept_type,sell_station_id,bus_id,bus_date,vehicle_id,license_tag_no,transport_company_id,transport_company_short_name, " _
        & " split_company_id,splict_company_short_name,route_id,route_name,vehicle_type_name,vehicle_type_code,owner_id,owner_name,bus_start_time," _
        & TransFieldValueToString(ST_LuggageSheetValidStatusNormal) & " AS valid_mark,worker,protocol_id, " _
        & " protocol_name,checker,GETDATE(),luggage_number,baggage_number,cal_weight,fact_weight,over_number,total_price,price_item_1,price_item_2,price_item_3,price_item_4, " _
        & " price_item_5,price_item_6,price_item_7,price_item_8,price_item_9,price_item_10,settlement_status,print_settle_ratio,print_settle_price" _
        & " FROM luggage_sheet_info " _
        & " WHERE sheet_id =" & TransFieldValueToString(CarrySheet)
    oDb.Execute szSql
    
    szSql = "UPDATE luggage_sheet_info SET valid_mark =" & TransFieldValueToString(ST_LuggageSheetValidStatusCancel) _
        & " WHERE sheet_id =" & TransFieldValueToString(CarrySheet)
    oDb.Execute szSql
    
    szSql = "INSERT luggage_carry_accept_lst (check_luggage_id,luggage_id) SELECT " & TransFieldValueToString(NewCarrySheet) & " AS check_luggage_id,luggage_id " _
          & "FROM luggage_carry_accept_lst WHERE check_luggage_id= " & TransFieldValueToString(CarrySheet)
    oDb.Execute szSql
    oDb.CommitTrans
    'your code goes here...
    Set oDb = Nothing
    
    Exit Sub
ReprintCarrySheetErr:
    Call RaiseError(err.Number, "LuggageSvr:ReprintCarrySheet Method", err.Description)
End Sub

'得到车次途经站点信息，以后可能要删除
'
'##ModelId=3E0404B20369
Public Function GetBusStationNames(BusDate As Date, BusID As String) As String()
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
    
    On Error GoTo GetBusStationNamesErr
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim aszTemp() As String
    Dim i As Integer
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    Dim szTables As String
    If BusDate = Date Then       '工作表
        szTables = "" & cszTableEnvBusStationWork & " e"
    Else
        szTables = "" & cszTableEnvBusStationInfo & " e"
    End If
    
    szSql = "SELECT distinct s.station_name,e.station_serial_no FROM station_info s," & szTables _
    & " WHERE e.bus_date = " & TransFieldValueToString(BusDate) _
    & " AND e.bus_id = " & TransFieldValueToString(BusID) _
    & " AND e.station_id = s.station_id ORDER BY e.station_serial_no"
    
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        ReDim aszTemp(1 To rsTemp.RecordCount)
        For i = 1 To rsTemp.RecordCount
            aszTemp(i) = FormatDbValue(rsTemp!station_name)
            rsTemp.MoveNext
        Next
    Else
        ShowError ERR_LuggageSvrStationNoInfo
    End If
    GetBusStationNames = aszTemp
    

    Exit Function
GetBusStationNamesErr:
    Call RaiseError(err.Number, "LuggageSvr:GetBusStationNames Method", err.Description)
End Function

'得到车次途经站点信息，以后可能要删除(楚门加的，有站点代码，用于修改签发单)
'
'##ModelId=3E0404B20369
Public Function GetBusStationNamesEx(BusDate As Date, BusID As String) As String()
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
    
    On Error GoTo GetBusStationNamesExErr
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim aszTemp() As String
    Dim i As Integer
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    Dim szTables As String
    If BusDate = Date Then       '工作表
        szTables = "" & cszTableEnvBusStationWork & " e"
    Else
        szTables = "" & cszTableEnvBusStationInfo & " e"
    End If
    
    szSql = "SELECT distinct s.station_id, s.station_name,e.station_serial_no FROM station_info s," & szTables _
    & " WHERE e.bus_date = " & TransFieldValueToString(BusDate) _
    & " AND e.bus_id = " & TransFieldValueToString(BusID) _
    & " AND e.station_id = s.station_id ORDER BY e.station_serial_no"
    
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        ReDim aszTemp(1 To rsTemp.RecordCount)
        For i = 1 To rsTemp.RecordCount
            aszTemp(i) = MakeDisplayString(FormatDbValue(rsTemp!station_id), FormatDbValue(rsTemp!station_name))
            rsTemp.MoveNext
        Next
    Else
        ShowError ERR_LuggageSvrStationNoInfo
    End If
    GetBusStationNamesEx = aszTemp
    

    Exit Function
GetBusStationNamesExErr:
    Call RaiseError(err.Number, "LuggageSvr:GetBusStationNames Method", err.Description)
End Function


'得到车号,取得车主,参运公司,拆帐公司
Public Function GetBusOtherInfo(szVehicleID As String) As ADODB.Recordset
    On Error GoTo err
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTmp As Recordset
    Dim i As Integer
            oDb.ConnectionString = GetConnectionStr
            szSql = "SELECT t.owner_id,o.owner_name,t.split_company_id,s.transport_company_short_name,t.transport_company_id,s.transport_company_short_name " _
                & " FROM " & cszTableVehicle & " t," & cszTableOwnerInfo & " o," & cszTableCompany & " s" _
                & " Where t.owner_id = o.owner_id And t.transport_company_id = s.transport_company_id " _
                & " AND t.vehicle_id= " & TransFieldValueToString(szVehicleID)
            Set rsTmp = oDb.Execute(szSql)
            Set GetBusOtherInfo = rsTmp
            Set rsTmp = Nothing
    Exit Function
err:
    Set rsTmp = Nothing
End Function

'得到车次的运行车辆
'##ModelId=3E040B7A036F
Public Function GetBusRunVehicles(BusID As String) As String()
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
    
    On Error GoTo GetBusRunVehiclesErr


'得到某一车次的运行车辆
    Dim aszVehicle() As String
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTmp As Recordset
    Dim i As Integer
    Dim szProjectID As String
    szProjectID = "0001"
    
    oDb.ConnectionString = GetConnectionStr
    szSql = "SELECT DISTINCT  B.vehicle_id vehicleId,I.license_tag_no license,O.owner_name owner,V.vehicle_type_short_name vehicleType" _
        & " ,T.transport_company_name company,I.status status FROM " _
        & "" & cszTableOwnerInfo & " O," & cszTableCompany & " T,vehicle_type_code V," & cszTableVehicle & " I," & cszTableBusVehicle & " B " _
        & " WHERE B.bus_id='" & BusID & "' AND B.vehicle_id=I.vehicle_id AND I.owner_id=O.owner_id " _
        & " AND I.transport_company_id=T.transport_company_id AND I.vehicle_type_code=V.vehicle_type_code AND B.project_id='" & szProjectID & "'"
    
    Set rsTmp = oDb.Execute(szSql)
    If rsTmp.RecordCount <> 0 Then
        ReDim aszVehicle(1 To rsTmp.RecordCount, 1 To 2)
        For i = 1 To rsTmp.RecordCount
            aszVehicle(i, 1) = FormatDbValue(rsTmp!VehicleID)
            aszVehicle(i, 2) = FormatDbValue(rsTmp!license)
            rsTmp.MoveNext
        Next i
    End If
    GetBusRunVehicles = aszVehicle

    Exit Function
GetBusRunVehiclesErr:
    Call RaiseError(err.Number, "LuggageSvr:GetBusRunVehicles Method", err.Description)
End Function


'查询行包签发单信息,返回签发单记录集
'##ModelId=3E040C4801E5
Public Function GetLuggageKinds() As ADODB.Recordset
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
    
    On Error GoTo GetLuggageKindsErr
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
        
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    szSql = "SELECT * FROM luggage_kinds_code ORDER BY kinds_code "
    Set rsTemp = oDb.Execute(szSql)
    Set GetLuggageKinds = rsTemp
    'your code goes here...

    Exit Function
GetLuggageKindsErr:
    Call RaiseError(err.Number, "LuggageSvr:GetLuggageKinds Method", err.Description)
End Function

'取得还未车次填入的签发单信息(楚门改为没有装车时间的行包单)
Public Function PreCarryLuggage() As ADODB.Recordset
      '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
    AssertHaveRight m_oActiveUser, RIGHT_LuggageSvrNormalCarry
    
    
    On Error GoTo CarryLuggageErr
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
      
      szSql = " SELECT sheet_id FROM luggage_sheet_info WHERE (up_load_time is NULL OR up_load_time = '1900-01-01') " _
            & " AND valid_mark=1 AND sell_station_id = " & TransFieldValueToString(m_oActiveUser.SellStationID) _
            & " ORDER BY sheet_id "
            
    Set rsTemp = oDb.Execute(szSql)
    Set PreCarryLuggage = rsTemp
    Set rsTemp = Nothing
    Exit Function
CarryLuggageErr:
      Call RaiseError(err.Number, "LuggageSvr:CarryLuggage Method", err.Description)
End Function

'old sheet list
Public Function GetOldSheetRs(SheetID As String) As ADODB.Recordset
  AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
    
    On Error GoTo GetCarrySheetRSErr
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szTempSql As String
    Dim szTempFromSql As String
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
   
      
    szSql = "SELECT l.*,a.*,p.price_item_1 as shipperprice,l.bus_id as busid " _
        & " FROM luggage_sheet_info l,luggage_accept_info a,luggage_carry_accept_lst sa,luggage_price_lst p " _
        & " WHERE l.sheet_id=sa.check_luggage_id AND sa.luggage_id=a.luggage_id AND p.luggage_id=a.luggage_id " _
        & " AND l.sheet_id='" & Trim(SheetID) & "'"
          
    Set rsTemp = oDb.Execute(szSql)
    Set GetOldSheetRs = rsTemp
   

    Exit Function
GetCarrySheetRSErr:
    Call RaiseError(err.Number, "LuggageSvr:GetCarrySheetRS Method", err.Description)
End Function


'指定车辆得到车主信息
Public Function GetOwnerName(szVehicle As String) As ADODB.Recordset
    On Error GoTo GetCarrySheetRSErr
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
   
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
   
      
    szSql = "SELECT o.owner_name FROM " & cszTableVehicle & " t," & cszTableOwnerInfo & " o WHERE t.owner_id=o.owner_id AND t.vehicle_id =" & TransFieldValueToString(szVehicle)
          
    Set rsTemp = oDb.Execute(szSql)
    Set GetOwnerName = rsTemp
    
Exit Function
GetCarrySheetRSErr:
End Function

Public Sub UpdateCarryLuggage(poSheetID As String, poVehicleID As String, poVehicleLicense As String, pszStation As String, pdtUpLoadTime As Date)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
    AssertHaveRight m_oActiveUser, RIGHT_LuggageSvrNormalCarry
    
    
    On Error GoTo CarryLuggageErr
    
    Dim szSql As String
    Dim szTempPriceSql As String
    Dim szTempSql As String
    Dim rsTmp As Recordset
    Dim oDb As New RTConnection
    Dim i As Integer, j As Integer
    Dim Count As Integer
    Dim rsTemp As Recordset
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    oDb.BeginTrans
    szSql = "SELECT t.owner_id,o.owner_name,t.split_company_id,s.transport_company_short_name,t.transport_company_id,s.transport_company_short_name " _
                & " FROM " & cszTableVehicle & " t," & cszTableOwnerInfo & " o," & cszTableCompany & " s" _
                & " Where t.owner_id = o.owner_id And t.transport_company_id = s.transport_company_id " _
                & " AND t.vehicle_id= " & TransFieldValueToString(poVehicleID)
            Set rsTmp = oDb.Execute(szSql)
            If rsTmp.RecordCount > 0 Then
                
            
                szSql = "UPDATE luggage_sheet_info SET  vehicle_id=" & TransFieldValueToString(poVehicleID) & "," _
                      & " license_tag_no=" & TransFieldValueToString(poVehicleLicense) _
                      & " , owner_id=" & TransFieldValueToString(Trim(rsTmp!owner_id)) _
                      & " , owner_name=" & TransFieldValueToString(Trim(rsTmp!owner_name)) _
                      & " , transport_company_id= " & TransFieldValueToString(Trim(rsTmp!transport_company_id)) _
                      & " , transport_company_short_name= " & TransFieldValueToString(Trim(rsTmp!transport_company_short_name)) _
                      & ", split_company_id=" & TransFieldValueToString(Trim(rsTmp!split_company_id)) _
                      & ", splict_company_short_name= " & TransFieldValueToString(Trim(rsTmp!transport_company_short_name)) _
                      & ", up_load_time= " & TransFieldValueToString(pdtUpLoadTime) _
                      & " WHERE sheet_id =" & TransFieldValueToString(poSheetID)
                          
                oDb.Execute szSql
                
                
                szSql = " UPDATE luggage_accept_info SET  license_tag_no = " & TransFieldValueToString(poVehicleLicense) _
                            & " WHERE luggage_id IN (SELECT luggage_id FROM luggage_carry_accept_lst " _
                            & " WHERE check_luggage_id = " & TransFieldValueToString(poSheetID) & ")"
                    
                oDb.Execute szSql
                
                
                If pszStation <> "" Then
                    szSql = " UPDATE luggage_accept_info SET des_station_id = " & TransFieldValueToString(ResolveDisplay(pszStation)) _
                            & " , des_station_name = " & TransFieldValueToString(ResolveDisplayEx(pszStation)) _
                            & " WHERE luggage_id IN (SELECT luggage_id FROM luggage_carry_accept_lst " _
                            & " WHERE check_luggage_id = " & TransFieldValueToString(poSheetID) & ")"
                    
                    oDb.Execute szSql
                End If
           End If
           
           oDb.CommitTrans
    Exit Sub
CarryLuggageErr:
      oDb.RollbackTrans
      Call RaiseError(err.Number, "LuggageSvr:CarryLuggage Method", err.Description)
End Sub


'到站的所有车辆
Public Function GetEndStationVehicle(EndStationVehicle As String) As ADODB.Recordset
 On Error GoTo ErrHandle:
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT  v.license_tag_no  FROM " & cszTableEnvBusWork & " e, " & cszTableVehicle & " v " _
            & " WHERE  v.vehicle_id = e.vehicle_id AND e.end_station_name= " & TransFieldValueToString(EndStationVehicle) _
            & " GROUP BY v.license_tag_no "
    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount = 0 Then Exit Function
    Set GetEndStationVehicle = rsTemp
    
    Set rsTemp = Nothing
 Exit Function
    
ErrHandle:
  Set rsTemp = Nothing
 End Function

'更改车辆信息((((((((((((((((((((((((((((((((((((((((((((((((
'Public Sub UpdateVehicle()
'    Dim szsql As String
'    Dim rsTemp As Recordset
'    Dim odb As New RTConnection
'    Dim i As Integer
'    odb.ConnectionString = GetConnectionStr(cszLuggage)
'    szsql = "select distinct vehicle_id,license_tag_no from luggage_sheet_info"
'    Set rsTemp = odb.Execute(szsql)
'    If rsTemp.RecordCount = 0 Then Exit Sub
'    For i = 0 To rsTemp.RecordCount - 1
'    If Trim(rsTemp!vehicle_id) <> "" Then
'    szsql = "update luggage_sheet_info set " _
'& " owner_id=(select distinct o.owner_id from " & cszTableOwnerInfo & " o," & csztablevehicle & " t,luggage_sheet_info l where l.vehicle_id=t.vehicle_id and" _
'& "       t.owner_id=o.owner_id and l.vehicle_id= " & TransFieldValueToString(rsTemp!vehicle_id) & "), " _
'& "owner_name=(select distinct o.owner_name from " & cszTableOwnerInfo & " o," & csztablevehicle & " t,luggage_sheet_info l where l.vehicle_id=t.vehicle_id and" _
'& "       t.owner_id=o.owner_id and l.vehicle_id=" & TransFieldValueToString(rsTemp!vehicle_id) & "), " _
'& "transport_company_id=(select distinct f.transport_company_id from " & csztablecompany & " f," & csztablevehicle & " t,luggage_sheet_info l where l.vehicle_id=t.vehicle_id and" _
'& "       f.transport_company_id=t.transport_company_id and l.vehicle_id=" & TransFieldValueToString(rsTemp!vehicle_id) & "), " _
'& "transport_company_short_name=(select distinct f.transport_company_short_name from " & csztablecompany & " f," & csztablevehicle & " t,luggage_sheet_info l where l.vehicle_id=t.vehicle_id and" _
'& "       f.transport_company_id=t.transport_company_id and l.vehicle_id=" & TransFieldValueToString(rsTemp!vehicle_id) & "), " _
'& " split_company_id=(select distinct t.split_company_id from " & csztablecompany & " f," & csztablevehicle & " t,luggage_sheet_info l where l.vehicle_id=t.vehicle_id and" _
'& "       f.transport_company_id=t.split_company_id and l.vehicle_id=" & TransFieldValueToString(rsTemp!vehicle_id) & "), " _
'& " splict_company_short_name=(select distinct f.transport_company_short_name from " & csztablecompany & " f," & csztablevehicle & " t,luggage_sheet_info l where l.vehicle_id=t.vehicle_id and" _
'& "          f.transport_company_id=t.split_company_id and l.vehicle_id=" & TransFieldValueToString(rsTemp!vehicle_id) & ") " _
'& " where vehicle_id=" & TransFieldValueToString(rsTemp!vehicle_id)
'    odb.Execute szsql
'    End If
'    rsTemp.MoveNext
'    Next i
'    Set rsTemp = Nothing
' Exit Sub
'End Sub


'得到行包公式信息
Public Function GetLugFormulaInfo(Optional pszFormulaID As String = "", Optional pszFormulaName As String = "") As TLuggageFormulaInfo()
    On Error GoTo ErrorHandle
    Dim szSql As String
    Dim oDb As New RTConnection
    
    Dim atFormulaInfo() As TLuggageFormulaInfo
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim j As Integer
    oDb.ConnectionString = GetConnectionStr
    
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
'    AssertHaveRight m_oActiveUser, RIGHT_LuggageSvrNormalCarry
    
    szSql = "SELECT * FROM luggage_formula_info " _
        & " WHERE 1=1"
        
    If pszFormulaID <> "" Then
        szSql = szSql & " AND formula_id = " & TransFieldValueToString(pszFormulaID)
    End If
    If pszFormulaName <> "" Then
        szSql = szSql & " AND formula_name = " & TransFieldValueToString(pszFormulaName)
    End If
    
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        ReDim atFormulaInfo(1 To rsTemp.RecordCount)
        For i = 1 To rsTemp.RecordCount
            atFormulaInfo(i).FormulaID = FormatDbValue(rsTemp!formula_id)
            atFormulaInfo(i).FormulaName = FormatDbValue(rsTemp!formula_name)
            atFormulaInfo(i).ItemFormula = FormatDbValue(rsTemp!item_formula)
            For j = 1 To 10
                atFormulaInfo(i).Param(j) = FormatDbValue(rsTemp.Fields("parameter_" & j))
            Next j
            rsTemp.MoveNext
        Next i
    End If
    GetLugFormulaInfo = atFormulaInfo
    Exit Function
ErrorHandle:
    err.Raise err.Number, err.Source, err.Description
    
End Function

'新增行包公式
Public Sub AddLugFormulaInfo(patFormulaInfo() As TLuggageFormulaInfo)
On Error GoTo ErrorHandle
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim nCount As Integer
    Dim j As Integer
    
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
'    AssertHaveRight m_oActiveUser, RIGHT_LuggageSvrNormalCarry
    
    oDb.ConnectionString = GetConnectionStr
    nCount = ArrayLength(patFormulaInfo)
    oDb.BeginTrans
        For i = 1 To nCount
            szSql = " SELECT * FROM luggage_formula_info WHERE formula_id = " & TransFieldValueToString(patFormulaInfo(i).FormulaID)
            Set rsTemp = oDb.Execute(szSql)
            If rsTemp.RecordCount > 0 Then
                RaiseError ERR_LuggagePriceFormulaHasExist
            End If
            szSql = " INSERT luggage_formula_info (" _
                & " formula_id , " _
                & " formula_name , " _
                & " item_formula , " _
                & " parameter_1 , " _
                & " parameter_2 , " _
                & " parameter_3 , " _
                & " parameter_4 , " _
                & " parameter_5 , " _
                & " parameter_6 , " _
                & " parameter_7 , " _
                & " parameter_8 , " _
                & " parameter_9 , " _
                & " parameter_10 ) " _
                & " VALUES ( "
            szSql = szSql & TransFieldValueToString(patFormulaInfo(i).FormulaID) & " ," _
                & TransFieldValueToString(patFormulaInfo(i).FormulaName) & " , " _
                & TransFieldValueToString(patFormulaInfo(i).ItemFormula) & ","
                
                
            For j = 1 To 9
                szSql = szSql & TransFieldValueToString(patFormulaInfo(i).Param(j)) & ","
            Next j
            szSql = szSql & TransFieldValueToString(patFormulaInfo(i).Param(10)) & ")" _

            oDb.Execute (szSql)
                
            
        Next i
    oDb.CommitTrans
    
    Exit Sub
ErrorHandle:
    oDb.RollbackTrans
    err.Raise err.Number, err.Source, err.Description
    
End Sub

'修改行包公式
Public Sub EditLugFormulaInfo(patFormulaInfo() As TLuggageFormulaInfo)
On Error GoTo ErrorHandle
    Dim szSql As String
    Dim oDb As New RTConnection
'    Dim rsTemp As Recordset
    Dim i As Integer
    Dim nCount As Integer
    Dim lAffect As Long
    Dim j As Integer
    
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
'    AssertHaveRight m_oActiveUser, RIGHT_LuggageSvrNormalCarry
    
    oDb.ConnectionString = GetConnectionStr
    nCount = ArrayLength(patFormulaInfo)
    oDb.BeginTrans
        For i = 1 To nCount
            szSql = " UPDATE luggage_formula_info SET " _
                & " formula_name = " & TransFieldValueToString(patFormulaInfo(i).FormulaName) & " , " _
                & " item_formula = " & TransFieldValueToString(patFormulaInfo(i).ItemFormula) & " , "
                
                
            For j = 1 To 9
                szSql = szSql & " parameter_" & j & " = " & TransFieldValueToString(patFormulaInfo(i).Param(j)) & " , "
            Next j
            szSql = szSql & " parameter_10 = " & TransFieldValueToString(patFormulaInfo(i).Param(10)) & "  " _
                & " WHERE formula_id = " & TransFieldValueToString(patFormulaInfo(i).FormulaID)
                
            
            oDb.Execute szSql, lAffect
            If lAffect <> 1 Then
                RaiseError ERR_LuggagePriceFormulaNotExist
            End If
            
        Next i
    oDb.CommitTrans
    
    Exit Sub
ErrorHandle:
    oDb.RollbackTrans
    err.Raise err.Number, err.Source, err.Description
    
End Sub


'删除行包公式


Public Sub DelLugFormulaInfo(pszFormulaID As String) 'patFormulaInfo() As TLuggageFormulaInfo)
On Error GoTo ErrorHandle
    Dim szSql As String
    Dim oDb As New RTConnection
'    Dim rsTemp As Recordset
    Dim i As Integer
    Dim nCount As Integer
    Dim lAffect As Long
    
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
'    AssertHaveRight m_oActiveUser, RIGHT_LuggageSvrNormalCarry
    
    oDb.ConnectionString = GetConnectionStr
    szSql = " DELETE luggage_formula_info  " _
        & " WHERE formula_id = " & TransFieldValueToString(pszFormulaID)
        
    
    oDb.Execute szSql, lAffect
    If lAffect <> 1 Then
        RaiseError ERR_LuggagePriceFormulaNotExist
    End If
            
    
    Exit Sub
ErrorHandle:
    err.Raise err.Number, err.Source, err.Description
    
End Sub

'查询各站点的行包收费项公式
Public Function GetLugItemFormulaInfo(Optional pszStationId As String = "", Optional pszStationName As String = "", Optional pnAcceptType As Integer = -1, Optional pszFormulaID As String = "", Optional pszFormulaName As String = "") As TLuggageItemFormulaInfo()
    On Error GoTo ErrorHandle
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim i As Integer
    Dim rsTemp As Recordset
    Dim atItemFormulaInfo() As TLuggageItemFormulaInfo
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
'    AssertHaveRight m_oActiveUser, RIGHT_LuggageSvrNormalCarry

    oDb.ConnectionString = GetConnectionStr
    '查询时,只查询出使用的行包收费项,新增时,则把所有的行包收费项都写入
    
    
    
    szSql = " SELECT l.* , f.formula_name , p.chinese_name , s.station_name " _
        & " FROM luggage_item_formula_lst l , luggage_price_item_info p , luggage_formula_info f , station_info s " _
        & " WHERE l.charge_item = p.charge_item " _
        & " AND l.accept_type = p.accept_type " _
        & " AND l.station_id = s.station_id " _
        & " AND l.formula_id = f.formula_id AND p.use_mark = " & TransFieldValueToString(ELuggagePriceItemUseMark.LuggagePriceItemUsed)
         
    If pszStationId <> "" Then
        szSql = szSql & " AND l.station_id = " & TransFieldValueToString(pszStationId)
    End If
    If pszStationName <> "" Then
        szSql = szSql & " AND s.station_name LIKE " & TransFieldValueToString(pszStationName & "%")
        
    End If
    ' Optional pnAcceptType As Integer = -1, Optional FormulaID As String = "", Optional FormulaName As String = "") As TLuggageItemFormulaInfo()
    If pnAcceptType <> -1 Then
        szSql = szSql & " AND l.accept_type = " & TransFieldValueToString(pnAcceptType)
    End If
    If pszFormulaID <> "" Then
        szSql = szSql & " AND l.formula_id = " & TransFieldValueToString(pszFormulaID)
        
    End If
    If pszFormulaName <> "" Then
        szSql = szSql & " AND f.formula_name = " & TransFieldValueToString(pszFormulaName)
    End If
    szSql = szSql & " ORDER BY l.station_id , l.accept_type , l.charge_item "
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        ReDim atItemFormulaInfo(1 To rsTemp.RecordCount)
        For i = 1 To rsTemp.RecordCount
            atItemFormulaInfo(i).AcceptType = FormatDbValue(rsTemp!Accept_type)
            atItemFormulaInfo(i).AcceptTypeName = GetLuggageTypeString(atItemFormulaInfo(i).AcceptType)
            atItemFormulaInfo(i).PriceItemID = FormatDbValue(rsTemp!charge_item)
            atItemFormulaInfo(i).PriceItemName = FormatDbValue(rsTemp!chinese_name)
            atItemFormulaInfo(i).FormulaID = FormatDbValue(rsTemp!formula_id)
            atItemFormulaInfo(i).FormulaName = FormatDbValue(rsTemp!formula_name)
            atItemFormulaInfo(i).StationID = FormatDbValue(rsTemp!station_id)
            atItemFormulaInfo(i).StationName = FormatDbValue(rsTemp!station_name)
            rsTemp.MoveNext
        Next i
        
    End If
    GetLugItemFormulaInfo = atItemFormulaInfo
    Exit Function
ErrorHandle:
    err.Raise err.Number, err.Source, err.Description
    
End Function

Public Sub SetLugItemFormulaInfo(patLuggageItemInfo() As TLuggageItemFormulaInfo)
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim i As Integer
    Dim nCount As Integer
    Dim lAffect As Long
    Dim szStation As String
    Dim rsTemp As Recordset
    Dim nStationCount As Integer
    Dim bSame As Boolean
    Dim j As Integer
    On Error GoTo ErrorHandle
    
    
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
'    AssertHaveRight m_oActiveUser, RIGHT_LuggageSvrNormalCarry
    
    oDb.ConnectionString = GetConnectionStr
    nCount = ArrayLength(patLuggageItemInfo)
    oDb.BeginTrans
        '判断该站点是否存在
        szSql = "SELECT * FROM station_info WHERE station_id IN ( "
        For i = 1 To nCount - 1
            bSame = False
            For j = i - 1 To 1 Step -1
                If patLuggageItemInfo(i).StationID = patLuggageItemInfo(j).StationID Then
                    bSame = True
                    Exit For
                End If
            Next j
            
            If Not bSame Then
                nStationCount = nStationCount + 1
                szStation = szStation & TransFieldValueToString(patLuggageItemInfo(i).StationID) & " , "
            End If
        Next i
        If nCount > 0 Then
            bSame = False
            For j = i - 1 To 1 Step -1
                If patLuggageItemInfo(i).StationID = patLuggageItemInfo(j).StationID Then
                    bSame = True
                    Exit For
                End If
            Next j
            If Not bSame Then
                nStationCount = nStationCount + 1
            End If
            szStation = szStation & TransFieldValueToString(patLuggageItemInfo(i).StationID) & " ) "
             
        End If
        szSql = szSql & szStation
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount <> nStationCount Then
            '说明有一个或多个站点不存在,则出错.
            RaiseError ERR_StationNotExist
        End If
        For i = 1 To nCount
            '暂不判断行包的收费项及受理类型,行包公式是否存在.
            
            
            szSql = " UPDATE luggage_item_formula_lst " _
                & " SET formula_id = " & TransFieldValueToString(patLuggageItemInfo(i).FormulaID) _
                & " WHERE charge_item = " & TransFieldValueToString(patLuggageItemInfo(i).PriceItemID) _
                & " AND accept_type = " & TransFieldValueToString(patLuggageItemInfo(i).AcceptType) _
                & " AND station_id = " & TransFieldValueToString(patLuggageItemInfo(i).StationID)
            oDb.Execute szSql, lAffect
            If lAffect = 0 Then
                szSql = " INSERT luggage_item_formula_lst (charge_item , accept_type , station_id , formula_id) VALUES ( " _
                    & TransFieldValueToString(patLuggageItemInfo(i).PriceItemID) & " , " _
                    & TransFieldValueToString(patLuggageItemInfo(i).AcceptType) & " , " _
                    & TransFieldValueToString(patLuggageItemInfo(i).StationID) & " ," _
                    & TransFieldValueToString(patLuggageItemInfo(i).FormulaID) & " ) "
                oDb.Execute szSql
                
            End If
        Next i
        
    oDb.CommitTrans
    
    Exit Sub
ErrorHandle:
    oDb.RollbackTrans
    err.Raise err.Number, err.Source, err.Description

End Sub



'得到环境车次的运行车辆
'##ModelId=3E040B7A036F
Public Function GetEnvBusRunVehicles(BusDate As Date, BusID As String) As String()
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
    
    On Error GoTo GetBusRunVehiclesErr


'得到某一车次的运行车辆
    Dim aszVehicle() As String
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTmp As Recordset
    Dim i As Integer
    Dim szProjectID As String
    szProjectID = "0001"
    
    oDb.ConnectionString = GetConnectionStr
    szSql = "SELECT B.vehicle_id vehicleId,I.license_tag_no license  FROM " _
        & " " & cszTableVehicle & " I," & cszTableEnvBusWork & " B " _
        & " WHERE B.bus_id='" & BusID & "' AND B.vehicle_id=I.vehicle_id  " _
        & " and b.bus_date = '" & BusDate & "'"
    
    Set rsTmp = oDb.Execute(szSql)
    If rsTmp.RecordCount <> 0 Then
        ReDim aszVehicle(1 To rsTmp.RecordCount, 1 To 2)
        For i = 1 To rsTmp.RecordCount
            aszVehicle(i, 1) = FormatDbValue(rsTmp!VehicleID)
            aszVehicle(i, 2) = FormatDbValue(rsTmp!license)
            rsTmp.MoveNext
        Next i
    End If
    GetEnvBusRunVehicles = aszVehicle

    Exit Function
GetBusRunVehiclesErr:
    Call RaiseError(err.Number, "LuggageSvr:GetEnvBusRunVehicles Method", err.Description)
End Function

'得到行包互联单位
Public Function GetConnectUnit() As TConnectUnit()
On Error GoTo ErrorHandle
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim TConnectUnit() As TConnectUnit
    AssertActiveUserValid m_oActiveUser, ERR_LuggageAcceptSheet
    oDb.ConnectionString = GetConnectionStr
    szSql = "select * from connect_unit_info" _
            & " Where (service_type = 2 Or service_type = 0) And unit_delete_status = 1 "
    
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then ReDim TConnectUnit(1 To rsTemp.RecordCount)
    
    Do While Not rsTemp.EOF
           i = i + 1
           TConnectUnit(i).szUnitID = Trim(rsTemp!unit_id)
           TConnectUnit(i).szUnitShortName = Trim(rsTemp!unit_short_name)
           TConnectUnit(i).szUnitFullName = Trim(rsTemp!unit_name)
           TConnectUnit(i).szIPAddress = Trim(rsTemp!ip_address)
           TConnectUnit(i).szDatabase = "" 'IIf(IsNull(rsTemp!sql_database), "", Trim(rsTemp!sql_database))
           TConnectUnit(i).szDBUser = "" ' IIf(IsNull(rsTemp!sql_user), "", Trim(rsTemp!sql_user))
           TConnectUnit(i).szUserPassword = "" ' IIf(IsNull(rsTemp!sql_pass), "", Trim(rsTemp!sql_pass))
           rsTemp.MoveNext
    Loop
    GetConnectUnit = TConnectUnit
    Exit Function
ErrorHandle:
    err.Raise err.Number, err.Source, err.Description
End Function


'得到指定车次的配载站发车时间
'
'##ModelId=3E03E166028A
Public Function GetAllotStationBusStartTime(BusID As String, BusDate As Date) As ADODB.Recordset
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageSvr
    
    On Error GoTo GetAllotStationBusStartTimeErr
    
    Dim szSql As String
    Dim szTemp As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    Dim szTables As String
    If BusDate >= Date Then       '工作表
        szTables = "" & cszTableEnvBusWork & " b , " & cszTableVehicle & " t  , " & cszTableCompany & " c , " & cszTableEnvAllotWork & " a "
    Else
        szTables = "" & cszTableEnvBusInfo & " b, " & cszTableVehicle & " t  ,  " & cszTableCompany & " c , " & cszTableEnvAllotInfo & " a "
    End If
    
    szSql = "SELECT DISTINCT t.license_tag_no,convert(char(20),a.bus_start_time,108) bus_start_time,b.bus_date,b.vehicle_type_code, b.vehicle_type_name,b.end_station_name , b.bus_type , b.transport_company_id , c.transport_company_short_name " _
            & " FROM " & szTables _
            & " WHERE b.bus_date=" & TransFieldValueToString(BusDate) _
            & " AND a.sell_station_id = " & TransFieldValueToString(m_oActiveUser.SellStationID) _
            & " AND a.bus_id = b.bus_id AND a.bus_date = b.bus_date " _
            & " AND b.status not in(1,2,8) AND a.can_sale_quantity<>0 " _
            & " AND b.bus_id = " & TransFieldValueToString(BusID) _
            & " AND t.vehicle_id=b.vehicle_id AND b.transport_company_id = c.transport_company_id ORDER BY b.bus_type,convert(char(20),a.bus_start_time,108)"
            
            
    Set rsTemp = oDb.Execute(szSql)
    Set GetAllotStationBusStartTime = rsTemp
    Set oDb = Nothing
    'your code goes here...

    Exit Function
GetAllotStationBusStartTimeErr:
    Call RaiseError(err.Number, "LuggageSvr:GetAllotStationBusStartTime Method", err.Description)
End Function

