VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Split"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E81F41E03D8"

Option Base 0
Option Explicit


Const cszSettleStr = "[拆算]"
Const cszTicketPriceStr = "[票价]"
Const cnFixFeeCount = 20
Const cnSplitItemCount = 20
'set this to 0 to disable debug code in this class
#Const DebugMode = 1
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3E8219D1002E
    Private mlClassDebugID As Long
#End If

Private m_oActiveUser As ActiveUser
Private m_nObjectStatus As EObjectStatus '对象状态
Private moCalc As New Calculater

Private m_bIsNew As Boolean '是否为同一路单

Private m_aszFixFeeItem() As String '有哪几个固定费用项
Private m_nFixFeeItemCount As Integer '固定费用的项目个数
Private m_bAllowSettleTotalNegative As Boolean '是否将上月结算的负数算到下月
Private m_nSettleNegativeSplitItem As Integer   '将上月结算的负数项放在结算项的位置
Public m_nServiceItemPosition As Integer '劳务费项目所在位置
Private m_nCarriageItemPosition As Integer '运费项目所在位置
Public m_nServiceItem As Integer



'##ModelId=3E81F46B0222
Public Sub Init(opAUser As ActiveUser)
    On Error GoTo InitErr
    
    Dim oSystemParam As SystemParam
    
    Set m_oActiveUser = opAUser
    m_nObjectStatus = ST_NormalObj
    
    Set oSystemParam = New SystemParam
    oSystemParam.Init m_oActiveUser
    m_aszFixFeeItem = Split(oSystemParam.FixFeeItem, ",")
    m_nFixFeeItemCount = ArrayLength(m_aszFixFeeItem)
    m_bAllowSettleTotalNegative = oSystemParam.AllowSettleTotalNegative '是否将上月结算的负数算到下月
    m_nSettleNegativeSplitItem = oSystemParam.SettleNegativeSplitItem '将上月结算的负数项放在结算项的位置
    m_nServiceItemPosition = oSystemParam.ServiceItemPosition '劳务费项目所在位置
    m_nCarriageItemPosition = oSystemParam.CarriageItemPosition '运费项目所在位置
    m_nServiceItem = m_nServiceItemPosition
    
    
    Exit Sub
InitErr:
    Call RaiseError(err.Number, "Split:Init Method")
End Sub

'##ModelId=3E81F47000EA
Public Function GetLastSettleSheetID() As String
    '得到最后一张结算单号
    Dim pszDate As String
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    On Error GoTo GetLastSettleSheetIDErr
    AssertObjIsValid
    
    
    oDB.ConnectionString = GetConnectionStr()
    
    pszDate = Right(Year(Now), 2) & IIf(Len(Month(Now)) = 2, Month(Now), 0 & Left(Month(Now), 2))
    
    'szSql = "SELECT settlement_sheet_id FROM fin_settle_sheet_lst WHERE settlement_sheet_id = " _
        & " (SELECT MAX(settlement_sheet_id) AS settlement_sheet_id FROM fin_settle_sheet_lst where left(settlement_sheet_id,6) = '" & pszDate & "'  )"
    
    '结算单号按结算人来得到
    'szSql = "SELECT MAX(settlement_sheet_id) settlement_sheet_id AS settlement_sheet_id FROM fin_settle_sheet_lst where user_id  = '" & m_oActiveUser.UserID & "'"
'    szSql = "SELECT MAX(settlement_sheet_id) AS settlement_sheet_id FROM fin_settle_sheet_lst WHERE " _
'        & " settle_date>=" & TransFieldValueToString(GetFirstMonthDay(Now)) & " and  settler= " & TransFieldValueToString(m_oActiveUser.UserName) & " "
        
    szSql = "SELECT settlement_sheet_id FROM fin_settle_sheet_lst WHERE settle_date = " _
        & " (SELECT MAX(settle_date) FROM fin_settle_sheet_lst )"
        
        
    Set rsTemp = oDB.Execute(szSql)
    
    Dim szUser As String
    If Len(Trim(m_oActiveUser.UserID)) >= 3 Then
        szUser = Right(Trim(m_oActiveUser.UserID), 3)
    ElseIf Len(Trim(m_oActiveUser.UserID)) = 2 Then
        szUser = "0" & Trim(m_oActiveUser.UserID)
    ElseIf Len(Trim(m_oActiveUser.UserID)) = 1 Then
        szUser = "00" & Trim(m_oActiveUser.UserID)
    End If
    
    If IsNull(rsTemp!settlement_sheet_id) Then
        GetLastSettleSheetID = pszDate & szUser & "001"
    Else
'        GetLastSettleSheetID = pszDate & szUser & Format(Right(FormatDbValue(rsTemp!settlement_sheet_id), 3) + 1, "000")
        GetLastSettleSheetID = FormatDbValue(rsTemp!settlement_sheet_id) + 1
    End If
    

    Exit Function
GetLastSettleSheetIDErr:
    Call RaiseError(err.Number, "Split:GetLastSettleSheetID Method")
End Function



Public Function TotalCheckSheetStationInfo(paszCheckSheetID() As String, Optional pbIsToday As Boolean = False) As Recordset 'TSettleSheetStation()
    '汇总路单站点信息
    On Error GoTo ErrorHandle
    
'    Dim atStationInfo() As TSettleSheetStation
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim szWhere As String
    Dim nSheetCount As Integer
    Dim i As Integer
    AssertObjIsValid
    
    
    oDB.ConnectionString = GetConnectionStr()
    
    nSheetCount = ArrayLength(paszCheckSheetID)
    
    szWhere = " ( "
    For i = 1 To nSheetCount - 1
            szWhere = szWhere & "'" & paszCheckSheetID(i) & "',"
    Next i
    szWhere = szWhere & "'" & paszCheckSheetID(i) & "')"
    
    Dim oSysParm As New SystemParam
    oSysParm.Init m_oActiveUser
    
    szSql = " SELECT (i.station_id) sell_station_id,(s.route_id) route_id ,(f.station_id) station_id," _
                & " (f.ticket_type) ticket_type"
    If Not oSysParm.AllowSplitBySomeTimes Then
        '如果不允许分多次结算
                    
        szSql = szSql & ",SUM(f.quantity) quantity "
    Else
        szSql = szSql & ",SUM(f.quantity-f.fact_quantity) quantity "
    End If
    szSql = szSql & " ,MAX(s.vehicle_type_code) vehicle_type_code " _
                & " ,MAX(s.vehicle_id) vehicle_id,MAX(s.transport_company_id) transport_company_id " _
                & " ,MAX(r.route_name) route_name , MAX(t.station_name) station_name , MAX(ticket_type_name) ticket_type_name " _
                & " ,MAX(v.vehicle_type_short_name) vehicle_type_name , MAX(i.station_name) sell_station_name , MAX(a.annotation) annotation " _
                & " FROM " & IIf(Not pbIsToday, cszTableSheetStation, cszTableSheetStationTemp) & " f, " & cszTableCheckSheet & " s , " & cszTableRouteInfo & " r " _
                & " , " & cszTableStationInfo & " t , " & cszTableTicketType & " y , " & cszTableVehicleTypeCode & " v " _
                & " ,  " & cszTableStationInfo & " i , sell_station_info z , area_code a " _
                & " WHERE  s.route_id = r.route_id AND f.station_id = t.station_id AND f.ticket_type = y.ticket_type_id " _
                & " AND z.station_id = i.station_id AND f.sell_station_id = z.sell_station_id AND t.area_code = a.area_code  " _
                & " AND s.vehicle_type_code = v.vehicle_type_code AND s.check_sheet_id=f.check_sheet_id AND s.check_sheet_id IN" & szWhere _
                & " GROUP BY i.station_id , s.route_id  ,s.vehicle_type_code, f.station_id , f.ticket_type "
'    Else
'        szSql = " SELECT (i.station_id) sell_station_id,(s.route_id) route_id ,(f.station_id) station_id," _
'                    & " (f.ticket_type) ticket_type,SUM(f.quantity-f.fact_quantity) quantity,MAX(s.vehicle_type_code) vehicle_type_code " _
'                    & " ,MAX(s.vehicle_id) vehicle_id,MAX(s.transport_company_id) transport_company_id " _
'                    & " ,MAX(r.route_name) route_name , MAX(t.station_name) station_name , MAX(ticket_type_name) ticket_type_name " _
'                    & " ,MAX(v.vehicle_type_short_name) vehicle_type_name , MAX(i.station_name) sell_station_name , MAX(a.annotation) annotation " _
'                    & " FROM " & cszTableSheetStation & " f, " & cszTableCheckSheet & " s , " & cszTableRouteInfo & " r " _
'                    & " , " & cszTableStationInfo & " t , " & cszTableTicketType & " y , " & cszTableVehicleTypeCode & " v " _
'                    & " ,  " & cszTableStationInfo & " i , sell_station_info z , area_code a " _
'                    & " WHERE  s.route_id = r.route_id AND f.station_id = t.station_id AND f.ticket_type = y.ticket_type_id " _
'                    & " AND z.station_id = i.station_id AND f.sell_station_id = z.sell_station_id AND t.area_code = a.area_code  " _
'                    & " AND s.vehicle_type_code = v.vehicle_type_code AND s.check_sheet_id=f.check_sheet_id AND s.check_sheet_id IN" & szWhere _
'                    & " GROUP BY i.station_id , s.route_id  ,s.vehicle_type_code, f.station_id , f.ticket_type "
'
'
'    End If
    Set rsTemp = oDB.Execute(szSql)
    
'    If rsTemp.RecordCount > 0 Then
'        ReDim atStationInfo(1 To rsTemp.RecordCount)
'        For i = 1 To rsTemp.RecordCount
'            atStationInfo(i).SellSationID = FormatDbValue(rsTemp!sell_station_id)
'            atStationInfo(i).SellStationName = FormatDbValue(rsTemp!sell_station_name)
'            atStationInfo(i).RouteID = FormatDbValue(rsTemp!route_id)
'            atStationInfo(i).RouteName = FormatDbValue(rsTemp!route_name)
'            atStationInfo(i).StationID = FormatDbValue(rsTemp!station_id)
'            atStationInfo(i).StationName = FormatDbValue(rsTemp!station_name)
'            atStationInfo(i).TicketType = FormatDbValue(rsTemp!ticket_type)
'            atStationInfo(i).TicketTypeName = FormatDbValue(rsTemp!ticket_type_name)
'            atStationInfo(i).Quantity = FormatDbValue(rsTemp!Quantity)
'            rsTemp.MoveNext
'        Next i
'    End If
    Set TotalCheckSheetStationInfo = rsTemp
    
    Exit Function
ErrorHandle:
    Call err.Raise(err.Number, "Split:TotalCheckSheetStationInfo Method", err.Description)

End Function

Public Function TotalCheckSheetStationInfoEx(paszCheckSheetID() As String, Optional pbIsToday As Boolean = False) As Recordset
    '汇总路单站点信息
    '按日期\票种\站点来汇总
    On Error GoTo ErrorHandle
    
'    Dim atStationInfo() As TSettleSheetStation
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim szWhere As String
    Dim nSheetCount As Integer
    Dim i As Integer
    AssertObjIsValid
    
    
    oDB.ConnectionString = GetConnectionStr()
    
    nSheetCount = ArrayLength(paszCheckSheetID)
    
    szWhere = " ( "
    For i = 1 To nSheetCount - 1
            szWhere = szWhere & "'" & paszCheckSheetID(i) & "',"
    Next i
    szWhere = szWhere & "'" & paszCheckSheetID(i) & "')"
    
    Dim oSysParm As New SystemParam
    oSysParm.Init m_oActiveUser
    
    szSql = " SELECT s.bus_date , s.station_id , s.ticket_type , MAX(t.ticket_type_name) ticket_type_name , MAX(n.station_name) station_name "
    
    If Not oSysParm.AllowSplitBySomeTimes Then
        '如果不允许分多次结算
        szSql = szSql & " , sum(quantity) quantity "
        
    Else
        szSql = szSql & " , sum(quantity-fact_quantity) quantity  "
    End If
    szSql = szSql & " , sum(mileage) mileage , sum(ticket_price) ticket_price" _
            & "  FROM " & IIf(Not pbIsToday, cszTableSheetStation, cszTableSheetStationTemp) & " s , sheet_lst h , ticket_type_code t , station_info n " _
            & " WHERE s.check_sheet_id = h.check_sheet_id  AND s.ticket_type = t.ticket_type_id AND s.station_id = n.station_id " _
            & " AND s.check_sheet_id IN" & szWhere _
            & " GROUP BY s.bus_date , s.station_id , s.ticket_type " _
            & " ORDER BY s.bus_date , s.station_id , s.ticket_type "

'
'        szSql = " SELECT s.bus_date , s.station_id , s.ticket_type , MAX(t.ticket_type_name) ticket_type_name , MAX(n.station_name) station_name " _
'            & " , sum(quantity-fact_quantity) quantity , sum(mileage) mileage , sum(ticket_price) ticket_price" _
'            & "  FROM fin_sheet_station_lst s , sheet_lst h , ticket_type_code t , station_info n " _
'            & " WHERE s.check_sheet_id = h.check_sheet_id  AND s.ticket_type = t.ticket_type_id AND s.station_id = n.station_id " _
'            & " AND s.check_sheet_id IN" & szWhere _
'            & " GROUP BY s.bus_date , s.station_id , s.ticket_type " _
'            & " ORDER BY s.bus_date , s.station_id , s.ticket_type "
'
'    End If
'
    Set rsTemp = oDB.Execute(szSql)
    
    
    Set TotalCheckSheetStationInfoEx = rsTemp
    
    Exit Function
ErrorHandle:
    Call err.Raise(err.Number, "Split:TotalCheckSheetStationInfoEx Method", err.Description)

End Function

Public Function GetCheckSheetDistinctStation(paszCheckSheetID() As String, Optional pbIsToday As Boolean = False) As Recordset
    '得到唯一的站点及票种信息。
    On Error GoTo ErrorHandle
    
    Dim atStationInfo() As TSettleSheetStation
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim szWhere As String
    Dim nSheetCount As Integer
    Dim i As Integer
    
    AssertObjIsValid
    
    
    oDB.ConnectionString = GetConnectionStr()
    
    nSheetCount = ArrayLength(paszCheckSheetID)
    
    szWhere = " ( "
    For i = 1 To nSheetCount - 1
            szWhere = szWhere & "'" & paszCheckSheetID(i) & "',"
    Next i
    szWhere = szWhere & "'" & paszCheckSheetID(i) & "')"
    
    
    szSql = " SELECT DISTINCT s.station_id , s.ticket_type , t.ticket_type_name , n.station_name " _
        & "  FROM " & IIf(Not pbIsToday, cszTableSheetStation, cszTableSheetStationTemp) & " s , sheet_lst h , ticket_type_code t , station_info n " _
        & " WHERE s.check_sheet_id = h.check_sheet_id AND s.ticket_type = t.ticket_type_id AND s.station_id = n.station_id " _
        & " AND s.check_sheet_id IN" & szWhere _
        & " ORDER BY s.station_id , s.ticket_type "
    
    
    Set rsTemp = oDB.Execute(szSql)
    

    Set GetCheckSheetDistinctStation = rsTemp
    
    Exit Function
ErrorHandle:
    Call err.Raise(err.Number, "Split:GetCheckSheetDistinctStation Method", err.Description)

End Function




'##ModelId=3E81F475006D
'此接口暂时不用

Public Function PreviewSplitCheckSheet(pszCompanyID As String, pszVehicleID As String, paszCheckSheetID() As String, Optional pbIsSingle As Boolean = False) As TSplitResult
'人数根据路单中的汇总人数得到
'返回预结算后的信息,此接口并不写数据库,只是将拆算结果返回.
'    '******注  pbIsSingle 参数表示是否强行按照单行结算 true按单行, false按设置结算
'
'    On Error GoTo PreviewSplitCheckSheetErr
'
'    Dim szSql As String
'    Dim rsTemp As Recordset
'    Dim oDb As New RTConnection
'
'
'    Dim bSplitByCompany As Boolean
'    Dim nSheetCount As Integer '路单数
'    Dim szWhere As String
'
'
'    Dim rsPassageNumber As Recordset '各公司、上车站、线路、站点、票种的人数及各票价项信息
'    Dim rsProtocol As Recordset
'    Dim rsHalve As Recordset
'    Dim rsCompanySettlePrice As Recordset
'    Dim rsVehicleSettlePrice As Recordset
'    Dim rsDefaultProtocol As Recordset '默认协议
'    Dim rsHalveProtocol As Recordset '加总平分公司的协议
'
'
'    Dim szHalveCompanyID As String '平分给对方的公司代码
'    Dim szHalveCompanyName As String
'    Dim dbHalveRatio As Double '加总平分的比率
'
'    Dim i As Integer
'    Dim j As Integer
'    Dim k As Integer
'    Dim l As Integer
'
'    Dim aszFormulaValue() As String
'    Dim nCount As Integer '票价项数
'
'
''    Dim szFormuarName As String '公式名称
'
'
'    Dim tResultInfo As TSplitResult '返回的结果信息
'
'    Dim atCompanyInfo() As TCompnaySettle '公司结算明细
'    Dim atVehicleInfo() As TVehilceSettle '车辆结算明细
''    Dim atStationInfo() As TSettleSheetStation '站点结算明细
'
'    Dim dbSplitItem As Double '临时拆出来的某一项的金额
'    Dim dbTotalTicketPrice As Double '总票款
'    Dim dbTotalQuantity As Double '总人数
'
'    Dim dbSettleOtherCompanyPrice As Double '总的结给对方公司金额
'    Dim dbSettleStationPrice As Double '总的结给车站的金额
'
'    Dim dbSplitlimitCharge As Double '底限额
'    Dim dbSplitUpCharge As Double '上限额
'
'
'    Dim bIsDefaut As Boolean '是否默认协议
'
'    Dim nVailableNumber As Integer '可用的结算项的数目
'
'
'    Dim nItem As Integer
'    Dim szItem As String
'
'
'
'
'    AssertObjIsValid
'    AssertHaveRight m_oActiveUser, RIGHT_SplitCheckSheet '验证用户是否有结算的权限
'
'    oDb.ConnectionString = GetConnectionStr()
'
'
'    'if 公司代码不为'' then 按公司结算elseif 车辆代码不为'' then 按车辆结算else      出调用参数不正确错endif
'    If pszCompanyID = "" And pszVehicleID = "" Then
'        '出传入参数错
'        RaiseError ERR_PreviewSplitInputParam
'    ElseIf pszCompanyID <> "" Then
'        bSplitByCompany = True
'    Else
'        bSplitByCompany = False
'    End If
'    nSheetCount = ArrayLength(paszCheckSheetID)
'    If nSheetCount = 0 Then RaiseError ERR_PreviewSplitInputParam '出传入参数错
'
'    szWhere = " ( "
'    For i = 1 To nSheetCount - 1
'            szWhere = szWhere & "'" & paszCheckSheetID(i) & "',"
'    Next i
'    szWhere = szWhere & "'" & paszCheckSheetID(i) & "')"
'
'
'
'    '判断每一张路单是否均存在，如不存在则出"路单不存在"的错。如无效或经过结算。则出"路单无效或路单已经结算"的错。
'    szSql = "SELECT * FROM " & cszTableCheckSheet & "  WHERE check_sheet_id IN " & szWhere
'    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount <> nSheetCount Then
'        '如果记录数不为传入的路单数,则出错
'        RaiseError ERR_SomeSheetNotExist
'    Else
'
'        '因为允许把其他车辆的路单使用电脑结算,故不需要再判断.
'
''        '判断所有的路单是不是同一公司或同一车辆
''        For i = 1 To rsTemp.RecordCount
''            If bSplitByCompany Then
''                If FormatDbValue(rsTemp!transport_company_id) <> pszCompanyID Then
''                    '出错
''                    RaiseError ERR_SplitSheetMustSameCompany
''                End If
''            Else
''                If UCase(FormatDbValue(rsTemp!vehicle_id)) <> UCase(pszVehicleID) Then
''                    '出错
''                    RaiseError ERR_SplitSheetMustSameVehicle
''                End If
''            End If
''            rsTemp.MoveNext
''        Next i
'    End If
'
'
'    '因此句语句在路单有300张左右时,效率太低,故暂时不验证
''    szSql = "SELECT * FROM " & cszTableCheckSheet & " WHERE (settlement_status = " & TransFieldValueToString(CS_CheckSheetSettled) _
''        & " OR valid_mark = " & TransFieldValueToString(CS_CheckSheetInvalid) & " )" _
''        & " AND check_sheet_id IN " & szWhere
''
''    Set rsTemp = odb.Execute(szSql)
''    If rsTemp.RecordCount <> 0 Then
''        '如果存在已废或已结算的路单,则出错
''        RaiseError ERR_SomeSheetHasCancelORSettle
''    End If
'
'
'    szSql = " SELECT * FROM price_item_info "
'    Set rsTemp = oDb.Execute(szSql)
'
'    nCount = rsTemp.RecordCount
'
'    aszFormulaValue = GetFormulaItem
'
'
'    '取出默认的协议
'    bIsDefaut = False
'    Set rsDefaultProtocol = GetProtocolFormulaCompany(, True)
'
'
'    szSql = " SELECT * FROM split_item_info WHERE split_status = " & TransFieldValueToString(ESplitItemUseStatus.CS_SplitItemUse)
'    Set rsTemp = oDb.Execute(szSql)
'    nVailableNumber = rsTemp.RecordCount
'
'
'    m_bIsNew = False
'
'    '如果按公司结算
'    If bSplitByCompany Then
'        '1．   汇总出各公司、车型、上车站、线路、站点、票种的人数及各票价项信息rsPassageNumber
'        szSql = "SELECT " & TransFieldValueToString(pszCompanyID) & "  transport_company_id , c.vehicle_type_code , c.route_id , i.station_id sell_station_id , s.station_id , s.ticket_type " _
'            & " , MAX(transport_company_short_name) transport_company_short_name , MAX(route_name) route_name  " _
'            & " , MAX(i.station_name) sell_station_name , MAX(f.station_name) station_name , MAX(vehicle_type_short_name) vehicle_type_short_name  " _
'            & " , MAX(ticket_type_name) ticket_type_name , MAX(a.annotation) annotation " _
'            & " , SUM(quantity) quantity , SUM(s.mileage*quantity) mileage , SUM(ticket_price*quantity) ticket_price , SUM(base_carriage*quantity) base_carriage " _
'            & " , SUM(price_item_1*quantity) price_item_1 , SUM(price_item_2*quantity) price_item_2 , SUM(price_item_3*quantity) price_item_3 , SUM(price_item_4*quantity) price_item_4 " _
'            & " , SUM(price_item_5*quantity) price_item_5 , SUM(price_item_6*quantity) price_item_6 , SUM(price_item_7*quantity) price_item_7 , SUM(price_item_8*quantity) price_item_8 " _
'            & " , SUM(price_item_9*quantity) price_item_9 , SUM(price_item_10*quantity) price_item_10 , SUM(price_item_11*quantity) price_item_11 , SUM(price_item_12*quantity) price_item_12 " _
'            & " , SUM(price_item_13*quantity) price_item_13 , SUM(price_item_14*quantity) price_item_14 , SUM(price_item_15*quantity) price_item_15 " _
'            & " FROM " & cszTableCheckSheet & " c, " & cszTableSheetStation & " s , " & cszTableCompany & " p , " & cszTableRouteInfo & " r " _
'            & " , " & cszTableStationInfo & "  i , " & cszTableStationInfo & " f , " & cszTableTicketType & " k , " & cszTableVehicleTypeCode & " d " _
'            & " , area_code a, sell_station_info t  " _
'            & " WHERE c.check_sheet_id = s.check_sheet_id AND c.vehicle_type_code = d.vehicle_type_code AND c.route_id = r.route_id " _
'            & " AND s.sell_station_id = t.sell_station_id AND t.station_id = i.station_id AND s.station_id = f.station_id AND s.ticket_type = k.ticket_type_id " _
'            & " AND p.transport_company_id = " & TransFieldValueToString(pszCompanyID) & "  AND f.area_code = a.area_code " _
'            & " AND c.check_sheet_id IN  " & szWhere _
'            & " GROUP BY c.vehicle_type_code , c.route_id , i.station_id , s.station_id , s.ticket_type ORDER BY c.route_id, c.vehicle_type_code  , i.station_id  , s.station_id , s.ticket_type "
'
'        Set rsPassageNumber = oDb.Execute(szSql)
'
'        '2．   得到所有该公司需要与之加总平分的公司线路信息rsHalve
'        szSql = " SELECT h.*, c.transport_company_short_name FROM fin_halve_company_lst h, " & cszTableCompany & " c " _
'            & " WHERE c.transport_company_id = h.transport_company_id AND h.transport_company_id = " & TransFieldValueToString(pszCompanyID)
'
'        Set rsHalve = oDb.Execute(szSql)
'
'        '3． 得到该公司的协议
'        Set rsProtocol = GetProtocolFormulaCompany(pszCompanyID)
'        If rsProtocol.RecordCount = 0 Then
'            '取出默认的协议
'            If rsDefaultProtocol.RecordCount <> 0 Then
'                Set rsProtocol = rsDefaultProtocol
'                bIsDefaut = True
'            Else
'                '该公司协议未设置
'                RaiseError ERR_CompanyNotSetProtocol
'            End If
'        Else
'            bIsDefaut = False
'        End If
'
'
'        For i = 1 To rsPassageNumber.RecordCount
'            szHalveCompanyID = ""
'            If rsHalve.RecordCount > 0 Then rsHalve.MoveFirst
'            For j = 1 To rsHalve.RecordCount
'                If FormatDbValue(rsPassageNumber!route_id) = FormatDbValue(rsHalve!route_id) And FormatDbValue(rsPassageNumber!transport_company_id) = FormatDbValue(rsHalve!transport_company_id) Then
'                    '找到该公司在该线路上需要加总平分
'
'                    szHalveCompanyID = FormatDbValue(rsHalve!other_transport_company_id)
'                    szHalveCompanyName = GetHavleOtherCompanyName(szHalveCompanyID)
'
'                    dbHalveRatio = FormatDbValue(rsHalve!Ratio)
'                    Exit For
'                End If
'                rsHalve.MoveNext
'            Next j
'            'If 该公司在rsHalve中找到需要加总平分then
'
'            If szHalveCompanyID = "" Or pbIsSingle Then
'                '不需要加总平分
'                '找到该线路的起始协议
'
'                FindRouteProtocol rsProtocol, FormatDbValue(rsPassageNumber!route_id), bIsDefaut
'
'                '如果公式中没有用到结算价,则不进行查询结算价
'                If FindSettlePriceStr(rsProtocol, nVailableNumber, cszSettleStr) Then
'                    Set rsCompanySettlePrice = GetCompanySettlePrice(rsPassageNumber!transport_company_id, rsPassageNumber!vehicle_type_code, rsPassageNumber!route_id, rsPassageNumber!sell_station_id, rsPassageNumber!station_id)
'                End If
'
'                For j = 1 To nVailableNumber
'
'                    SetFormulaValueCompany rsPassageNumber, aszFormulaValue, nCount, rsCompanySettlePrice, rsProtocol, atCompanyInfo, pszCompanyID, FormatDbValue(rsPassageNumber!transport_company_short_name), False, bIsDefaut, 1
'
'                    rsProtocol.MoveNext
'                    m_bIsNew = True
'                Next j
''                atCompanyInfo(1).PassengerNumber = atCompanyInfo(1).PassengerNumber + FormatDbValue(rsPassageNumber!Quantity)
'
'            Else
'                '将这些人数及售票的票价项分成两半，一半(此比率根据fin_halve_company_lst中的设定比率)拆给该公司，剩下部分拆给加总平分的公司。
'
'                '先将应拆给该公司的部分拆给该公司
'                '找到该线路的起始协议
'                FindRouteProtocol rsProtocol, FormatDbValue(rsPassageNumber!route_id), bIsDefaut
'                '如果公式中没有用到结算价,则不进行查询结算价
'                If FindSettlePriceStr(rsProtocol, nVailableNumber, cszSettleStr) Then
'                    Set rsCompanySettlePrice = GetCompanySettlePrice(rsPassageNumber!transport_company_id, rsPassageNumber!vehicle_type_code, rsPassageNumber!route_id, rsPassageNumber!sell_station_id, rsPassageNumber!station_id)
'                End If
'
'
'
'                For j = 1 To nVailableNumber
'                    SetFormulaValueCompany rsPassageNumber, aszFormulaValue, nCount, rsCompanySettlePrice, rsProtocol, atCompanyInfo, pszCompanyID, FormatDbValue(rsPassageNumber!transport_company_short_name), True, bIsDefaut, dbHalveRatio
'
'                    rsProtocol.MoveNext
'                    m_bIsNew = True
'                Next j
'
'
'                '将，剩下部分拆给加总平分的公司。
'                Dim bIsHalveDefault As Boolean
'                '取出协议
'                Set rsHalveProtocol = GetProtocolFormulaCompany(szHalveCompanyID)
'                If rsHalveProtocol.RecordCount = 0 Then
'                    '取出默认的协议
'                    If rsDefaultProtocol.RecordCount <> 0 Then
'                        Set rsHalveProtocol = rsDefaultProtocol
'                        bIsHalveDefault = True
'                    Else
'                        '该公司协议未设置
'                        RaiseError ERR_CompanyNotSetProtocol
'                    End If
'                Else
'                    bIsHalveDefault = False
'                End If
'
'
'
'                m_bIsNew = False
'                '找到该线路的起始协议
'                FindRouteProtocol rsHalveProtocol, FormatDbValue(rsPassageNumber!route_id), bIsHalveDefault
'
'
'                '如果公式中没有用到结算价,则不进行查询结算价
'                If FindSettlePriceStr(rsHalveProtocol, nVailableNumber, cszSettleStr) Then
'                    Set rsCompanySettlePrice = GetCompanySettlePrice(szHalveCompanyID, rsPassageNumber!vehicle_type_code, rsPassageNumber!route_id, rsPassageNumber!sell_station_id, rsPassageNumber!station_id)
'                End If
'
'
'                For j = 1 To nVailableNumber
'                    SetFormulaValueCompany rsPassageNumber, aszFormulaValue, nCount, rsCompanySettlePrice, rsHalveProtocol, atCompanyInfo, szHalveCompanyID, szHalveCompanyName, True, bIsHalveDefault, 1 - dbHalveRatio
'                    rsHalveProtocol.MoveNext
'                    m_bIsNew = True
'                Next j
'
'            End If
'
''            dbTotalTicketPrice = dbTotalTicketPrice + FormatDbValue(rsPassageNumber!ticket_price)
''            dbTotalQuantity = dbTotalQuantity + FormatDbValue(rsPassageNumber!Quantity)
'
'            dbTotalTicketPrice = FormatDbValue(rsPassageNumber!ticket_price)
'            dbTotalQuantity = FormatDbValue(rsPassageNumber!Quantity)
'            tResultInfo.SettleSheetInfo.TotalTicketPrice = tResultInfo.SettleSheetInfo.TotalTicketPrice + dbTotalTicketPrice
'            tResultInfo.SettleSheetInfo.TotalQuantity = tResultInfo.SettleSheetInfo.TotalQuantity + dbTotalQuantity
'
'            rsPassageNumber.MoveNext
'            m_bIsNew = False
'        Next i
'
'
'
'    Else
'
'        'else 按车辆结算
'        Dim szCompanyID As String
'
''        '1.汇总出各车辆、上车站、线路、站点、票种的人数及各票价项rsPassageNumber。
''        szSql = "SELECT c.vehicle_id  , c.route_id , i.station_id sell_station_id , s.station_id , s.ticket_type " _
''            & " , MAX(license_tag_no) license_tag_no , MAX(route_name) route_name  " _
''            & " , MAX(i.station_name) sell_station_name , MAX(f.station_name) station_name,MAX(d.vehicle_type_code) vehicle_type_code  " _
''            & " , MAX(ticket_type_name) ticket_type_name , MAX(c.transport_company_id) transport_company_id " _
''            & " , MAX(m.transport_company_short_name) transport_company_short_name , MAX(a.annotation) annotation " _
''            & " , SUM(quantity) quantity , SUM(s.mileage*quantity) mileage, SUM(ticket_price) ticket_price , SUM(base_carriage) base_carriage " _
''            & " , SUM(price_item_1) price_item_1 , SUM(price_item_2) price_item_2 , SUM(price_item_3) price_item_3 , SUM(price_item_4) price_item_4 " _
''            & " , SUM(price_item_5) price_item_5 , SUM(price_item_6) price_item_6 , SUM(price_item_7) price_item_7 , SUM(price_item_8) price_item_8 " _
''            & " , SUM(price_item_9) price_item_9 , SUM(price_item_10) price_item_10 , SUM(price_item_11) price_item_11 , SUM(price_item_12) price_item_12 " _
''            & " , SUM(price_item_13) price_item_13 , SUM(price_item_14) price_item_14 , SUM(price_item_15) price_item_15 " _
''            & " FROM " & cszTableCheckSheet & " c, " & cszTableSheetStation & " s , " & cszTableVehicle & " p , " & cszTableRouteInfo & " r " _
''            & " , " & cszTableStationInfo & "  i , " & cszTableStationInfo & " f , " & cszTableTicketType & " k ," & cszTableCompany & " m " _
''            & " , " & cszTableVehicleTypeCode & " d , area_code a , sell_station_info t  " _
''            & " WHERE c.vehicle_type_code=d.vehicle_type_code AND p.vehicle_id=c.vehicle_id AND c.check_sheet_id = s.check_sheet_id " _
''            & " AND c.route_id = r.route_id AND m.transport_company_id=c.transport_company_id " _
''            & " AND s.sell_station_id = t.sell_station_id AND t.station_id = i.station_id AND s.station_id = f.station_id AND s.ticket_type = k.ticket_type_id " _
''            & " AND f.area_code = a.area_code " _
''            & " AND c.check_sheet_id IN  " & szWhere _
''            & " GROUP BY c.vehicle_id  , c.route_id , i.station_id , s.station_id , s.ticket_type ORDER BY c.vehicle_id  , c.route_id , i.station_id , s.station_id , s.ticket_type "
'        Debug.Print pszCompanyID
'
'        '1.汇总出各车辆、上车站、线路、站点、票种的人数及各票价项rsPassageNumber。
'        szSql = "SELECT " & TransFieldValueToString(pszVehicleID) & " vehicle_id  , c.route_id , i.station_id sell_station_id , s.station_id , s.ticket_type " _
'            & " , MAX(license_tag_no) license_tag_no , MAX(route_name) route_name  " _
'            & " , MAX(i.station_name) sell_station_name , MAX(f.station_name) station_name,MAX(d.vehicle_type_code) vehicle_type_code  " _
'            & " , MAX(ticket_type_name) ticket_type_name , MAX(c.transport_company_id) transport_company_id " _
'            & " , MAX(m.transport_company_short_name) transport_company_short_name , MAX(a.annotation) annotation " _
'            & " , SUM(quantity) quantity , SUM(s.mileage*quantity) mileage , SUM(ticket_price*quantity) ticket_price , SUM(base_carriage*quantity) base_carriage " _
'            & " , SUM(price_item_1*quantity) price_item_1 , SUM(price_item_2*quantity) price_item_2 , SUM(price_item_3*quantity) price_item_3 , SUM(price_item_4*quantity) price_item_4 " _
'            & " , SUM(price_item_5*quantity) price_item_5 , SUM(price_item_6*quantity) price_item_6 , SUM(price_item_7*quantity) price_item_7 , SUM(price_item_8*quantity) price_item_8 " _
'            & " , SUM(price_item_9*quantity) price_item_9 , SUM(price_item_10*quantity) price_item_10 , SUM(price_item_11*quantity) price_item_11 , SUM(price_item_12*quantity) price_item_12 " _
'            & " , SUM(price_item_13*quantity) price_item_13 , SUM(price_item_14*quantity) price_item_14 , SUM(price_item_15*quantity) price_item_15 " _
'            & " FROM " & cszTableCheckSheet & " c, " & cszTableSheetStation & " s , " & cszTableVehicle & " p , " & cszTableRouteInfo & " r " _
'            & " , " & cszTableStationInfo & "  i , " & cszTableStationInfo & " f , " & cszTableTicketType & " k ," & cszTableCompany & " m " _
'            & " , " & cszTableVehicleTypeCode & " d , area_code a , sell_station_info t  " _
'            & " WHERE c.vehicle_type_code=d.vehicle_type_code AND p.vehicle_id=" & TransFieldValueToString(pszVehicleID) & " AND c.check_sheet_id = s.check_sheet_id " _
'            & " AND c.route_id = r.route_id AND m.transport_company_id=c.transport_company_id " _
'            & " AND s.sell_station_id = t.sell_station_id AND t.station_id = i.station_id AND s.station_id = f.station_id AND s.ticket_type = k.ticket_type_id " _
'            & " AND f.area_code = a.area_code " _
'            & " AND c.check_sheet_id IN  " & szWhere _
'            & " GROUP BY c.route_id , i.station_id , s.station_id , s.ticket_type ORDER BY c.route_id , i.station_id , s.station_id , s.ticket_type "
'
'
'        Set rsPassageNumber = oDb.Execute(szSql)
'
'        szCompanyID = FormatDbValue(rsPassageNumber!transport_company_id)
'        '2.  得到需要结算的车辆所属公司的加总平分公司信息。
'
'        szSql = " SELECT h.*, c.transport_company_short_name FROM fin_halve_company_lst h, " & cszTableCompany & " c " _
'            & " WHERE c.transport_company_id = h.transport_company_id AND c.transport_company_id = " & TransFieldValueToString(szCompanyID)
'
'        Set rsHalve = oDb.Execute(szSql)
'        '3． 得到该车辆的协议
'        Set rsProtocol = GetProtocolFormulaVehicle(pszVehicleID, False)
'        bIsDefaut = False
'        If rsProtocol.RecordCount = 0 Then
'            '取出默认的协议
'            If rsDefaultProtocol.RecordCount <> 0 Then
'                Set rsProtocol = rsDefaultProtocol
'                bIsDefaut = True
'            Else
'                '该车辆协议未设置
'                RaiseError ERR_VehicleNotSetProtocol
'            End If
'        End If
'
'        '4． For I = 1 to rsPassageNumber.RecordCount
'        For i = 1 To rsPassageNumber.RecordCount
'
'
'            szHalveCompanyID = ""
'            For j = 1 To rsHalve.RecordCount
'                If FormatDbValue(rsPassageNumber!route_id) = FormatDbValue(rsHalve!route_id) And FormatDbValue(rsPassageNumber!transport_company_id) = FormatDbValue(rsHalve!transport_company_id) Then
'                    '找到该公司在该线路上需要加总平分
'                    szHalveCompanyID = FormatDbValue(rsHalve!other_transport_company_id)
'                    szHalveCompanyName = GetHavleOtherCompanyName(szHalveCompanyID)
'                    dbHalveRatio = FormatDbValue(rsHalve!Ratio)
'                    Exit For
'                End If
'            Next j
'            If szHalveCompanyID = "" Then
'                '不需要加总平分
'                '将相应的人数及各售票票价只算给该车辆。
'                rsProtocol.MoveFirst
'
'                '如果公式中没有用到结算价,则不进行查询结算价
'                If FindSettlePriceStr(rsProtocol, nVailableNumber, cszSettleStr) Then
'                    Set rsVehicleSettlePrice = GetVehicleSettlePrice(rsPassageNumber!vehicle_id, rsPassageNumber!route_id, rsPassageNumber!sell_station_id, rsPassageNumber!station_id)
'                End If
'
'                For j = 1 To nVailableNumber
'                    SetFormulaValueVehicle rsPassageNumber, aszFormulaValue, nCount, rsVehicleSettlePrice, rsProtocol, atVehicleInfo, pszVehicleID, FormatDbValue(rsPassageNumber!license_tag_no), False, 0
'                    rsProtocol.MoveNext
'                    m_bIsNew = True
'                Next j
'            Else
'                'Else累加该车辆的结算费用明细将这些人数及各售票票价项分为两半，一半算给该车辆，另一半算给加总平分中的另一公司。
'
'
'                '先结给该车辆，方法与上同。
'
'                '如果公式中没有用到结算价,则不进行查询结算价
'                If FindSettlePriceStr(rsProtocol, nVailableNumber, cszSettleStr) Then
'                    Set rsVehicleSettlePrice = GetVehicleSettlePrice(rsPassageNumber!vehicle_id, rsPassageNumber!route_id, rsPassageNumber!sell_station_id, rsPassageNumber!station_id)
'                End If
'
'                rsProtocol.MoveFirst
'                For j = 1 To nVailableNumber
'                    SetFormulaValueVehicle rsPassageNumber, aszFormulaValue, nCount, rsVehicleSettlePrice, rsProtocol, atVehicleInfo, pszVehicleID, FormatDbValue(rsPassageNumber!license_tag_no), False, dbHalveRatio
'                    rsProtocol.MoveNext
'                    m_bIsNew = True
'                Next j
'                '再结给该公司，与按公司结算类似
'                '取出协议
'                bIsHalveDefault = False
'                Set rsHalveProtocol = GetProtocolFormulaCompany(szHalveCompanyID)
'                If rsHalveProtocol.RecordCount = 0 Then
'                    '取出默认的协议
'                    If rsDefaultProtocol.RecordCount <> 0 Then
'                        Set rsHalveProtocol = rsDefaultProtocol
'                        bIsHalveDefault = True
'                    Else
'                        '该公司协议未设置
'                        RaiseError ERR_VehicleNotSetProtocol
'                    End If
'                End If
'
'                '如果公式中没有用到结算价,则不进行查询结算价
'                If FindSettlePriceStr(rsHalveProtocol, nVailableNumber, cszSettleStr) Then
'                    Set rsVehicleSettlePrice = GetVehicleSettlePrice(rsPassageNumber!vehicle_id, rsPassageNumber!route_id, rsPassageNumber!sell_station_id, rsPassageNumber!station_id)
'                End If
'
'                m_bIsNew = False
'                rsHalveProtocol.MoveFirst
'                FindRouteProtocol rsHalveProtocol, FormatDbValue(rsPassageNumber!route_id), bIsHalveDefault
'                For j = 1 To nVailableNumber
'                    SetFormulaValueCompany rsPassageNumber, aszFormulaValue, nCount, rsCompanySettlePrice, rsHalveProtocol, atCompanyInfo, szHalveCompanyID, szHalveCompanyName, True, bIsHalveDefault, 1 - dbHalveRatio
'                    rsHalveProtocol.MoveNext
'                    m_bIsNew = True
'                Next j
'
'            End If
'
'            dbTotalTicketPrice = FormatDbValue(rsPassageNumber!ticket_price)
'            dbTotalQuantity = FormatDbValue(rsPassageNumber!Quantity)
'            tResultInfo.SettleSheetInfo.TotalTicketPrice = tResultInfo.SettleSheetInfo.TotalTicketPrice + dbTotalTicketPrice
'            tResultInfo.SettleSheetInfo.TotalQuantity = tResultInfo.SettleSheetInfo.TotalQuantity + dbTotalQuantity
'            rsPassageNumber.MoveNext
'            m_bIsNew = False
'        Next i
'    End If
'    '将总路单数、总票款、总人数、应结票款、结给车站票款以及公司的费用明细和车辆明细均返回
'    With tResultInfo.SettleSheetInfo
'
'        .CheckSheetCount = ArrayLength(paszCheckSheetID) '路单数
'        For i = 1 To ArrayLength(atCompanyInfo)
'            rsProtocol.MoveFirst
''            当结算有底限时 , 现在的处理方式是不对的, 未判断不同公司, 线路不同, 其协议不同, 协议不同时当然其底限也是不同的
''            If bSplitByCompany Then
''               FindRouteProtocol rsProtocol, rsCompanySettlePrice, bIsDefaut
''            End If
'            For j = 1 To nVailableNumber 'rsProtocol.RecordCount
'                szItem = FormatDbValue(rsProtocol!split_item_id)
'                nItem = CInt(szItem)
'                dbSplitlimitCharge = GetlimitCharge(FormatDbValue(rsProtocol!protocol_id), szItem, dbSplitUpCharge)
'                '如果上限不为0,则   该项值=上限值-该项值
'                If dbSplitUpCharge <> 0 Then
'                    atCompanyInfo(i).SplitItem(nItem) = FormatMoney(dbSplitUpCharge - atCompanyInfo(i).SplitItem(nItem))
'                End If
'                '当下限值大于该项值时,该值为下限值
'                If dbSplitlimitCharge > atCompanyInfo(i).SplitItem(nItem) Then
'                    atCompanyInfo(i).SplitItem(nItem) = dbSplitlimitCharge
'                Else
''                    atCompanyInfo(i).SplitItem(nItem) = atCompanyInfo(i).SplitItem(nItem)
'                End If
'                atCompanyInfo(i).SplitItem(nItem) = FormatMoney(atCompanyInfo(i).SplitItem(nItem))
'                atCompanyInfo(i).SettlePrice = atCompanyInfo(i).SettlePrice + GetSettlePrice(rsProtocol, atCompanyInfo(i).SplitItem(nItem))
'                atCompanyInfo(i).SettleStationPrice = atCompanyInfo(i).SettleStationPrice + GetSettleStationPrice(rsProtocol, atCompanyInfo(i).SplitItem(nItem))
'                rsProtocol.MoveNext
'            Next j
'            atCompanyInfo(i).SettlePrice = FormatMoney(atCompanyInfo(i).SettlePrice) ' FormatMoney (atCompanyInfo(i).SettlePrice + 0.01, 1)
'            atCompanyInfo(i).SettleStationPrice = FormatMoney(atCompanyInfo(i).SettleStationPrice) ' Round(atCompanyInfo(i).SettleStationPrice + 0.01, 1)
'
'            dbSettleOtherCompanyPrice = dbSettleOtherCompanyPrice + atCompanyInfo(i).SettlePrice
'            dbSettleStationPrice = dbSettleStationPrice + atCompanyInfo(i).SettleStationPrice
'        Next i
'
'        For i = 1 To ArrayLength(atVehicleInfo)
'            rsProtocol.MoveFirst
'            For j = 1 To nVailableNumber 'rsProtocol.RecordCount
'                szItem = FormatDbValue(rsProtocol!split_item_id)
'                nItem = CInt(szItem)
'                dbSplitlimitCharge = GetlimitCharge(FormatDbValue(rsProtocol!protocol_id), szItem, dbSplitUpCharge)
'                '如果上限不为0,则   该项值=上限值-该项值
'                If dbSplitUpCharge <> 0 Then
'                    atVehicleInfo(i).SplitItem(nItem) = FormatMoney(dbSplitUpCharge - atVehicleInfo(i).SplitItem(nItem))
'                End If
'                '当下限值大于该项值时,该值为下限值
'                If dbSplitlimitCharge > atVehicleInfo(i).SplitItem(nItem) Then
'                    atVehicleInfo(i).SplitItem(nItem) = dbSplitlimitCharge
'                Else
''                    atVehicleInfo(i).SplitItem(nItem) = atVehicleInfo(i).SplitItem(nItem)
'                End If
'                atVehicleInfo(i).SplitItem(nItem) = FormatMoney(atVehicleInfo(i).SplitItem(nItem))
'                atVehicleInfo(i).SettlePrice = atVehicleInfo(i).SettlePrice + GetSettlePrice(rsProtocol, atVehicleInfo(i).SplitItem(nItem))
'                atVehicleInfo(i).SettleStationPrice = atVehicleInfo(i).SettleStationPrice + GetSettleStationPrice(rsProtocol, atVehicleInfo(i).SplitItem(nItem))
'                rsProtocol.MoveNext
'            Next j
'            atVehicleInfo(i).SettlePrice = FormatMoney(atVehicleInfo(i).SettlePrice)  ' Round(atVehicleInfo(i).SettlePrice + 0.01, 1)
'            atVehicleInfo(i).SettleStationPrice = FormatMoney(atVehicleInfo(i).SettleStationPrice)  ' Round(atVehicleInfo(i).SettleStationPrice + 0.01, 1)
'
'
'            dbSettleOtherCompanyPrice = dbSettleOtherCompanyPrice + atVehicleInfo(i).SettlePrice
'            dbSettleStationPrice = dbSettleStationPrice + atVehicleInfo(i).SettleStationPrice
'        Next i
'        dbSettleOtherCompanyPrice = dbSettleOtherCompanyPrice ' Round(dbSettleOtherCompanyPrice + 0.01, 1)
'        dbSettleStationPrice = dbSettleStationPrice ' Round(dbSettleStationPrice + 0.01, 1)
'
'        .SettleOtherCompanyPrice = dbSettleOtherCompanyPrice '结算给对方公司金额
'        .SettleStationPrice = dbSettleStationPrice '结给站方金额
'        If pszCompanyID <> "" Then
'            .ObjectID = pszCompanyID '对象代码
'            .SettleObject = CS_SettleByTransportCompany  '结算对象
'
'        Else
'            .ObjectID = pszVehicleID
'            .SettleObject = CS_SettleByVehicle
'        End If
'        If rsPassageNumber.RecordCount > 0 Then
'        rsPassageNumber.MoveFirst
'        .TransportCompanyID = FormatDbValue(rsPassageNumber!transport_company_id)
'        .TransportCompanyName = FormatDbValue(rsPassageNumber!transport_company_short_name)
'        End If
'    End With
'
'    tResultInfo.CompanyInfo = atCompanyInfo
'    tResultInfo.VehicleInfo = atVehicleInfo
'    PreviewSplitCheckSheet = tResultInfo
'
'    Exit Function
'PreviewSplitCheckSheetErr:
'    err.Raise err.Number, err.Source, err.Description

    
End Function



'##ModelId=3E81F475006D

'人数根据传入的记录集来
'返回预结算后的信息,此接口并不写数据库,只是将拆算结果返回.
Public Function PreviewSplitCheckSheetEx(pszCompanyID As String, pszVehicleID As String, paszCheckSheetID() As String, prsPassageNumber As Recordset _
        , pdtStartDate As Date, pdtEndDate As Date, patExtraInfo() As TSettleExtraInfo, Optional pbIsSingle As Boolean = False _
        , Optional pszBusID As String) As TSplitResult
        
    'patExtraInfo() 里面是手工补票人数信息及结算信息
        
    '******注  pbIsSingle 参数表示是否强行按照单行结算 true按单行, false按设置结算
    'prsPassageNumber 与PreviewSplitCheckSheet   的汇总记录集字段 相同
    
    On Error GoTo PreviewSplitCheckSheetErr
    
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    
    
    Dim bSplitByCompany As Boolean
    Dim bSplitByVehicle As Boolean
    Dim nSheetCount As Integer '路单数
    Dim szWhere As String
    
    
    Dim rsProtocol As Recordset
    Dim rsHalve As Recordset
    Dim rsCompanySettlePrice As Recordset
    Dim rsVehicleSettlePrice As Recordset
    Dim rsBusSettlePrice As Recordset
    Dim rsDefaultProtocol As Recordset '默认协议
    Dim rsHalveProtocol As Recordset '加总平分公司的协议
    
    
    Dim szHalveCompanyID As String '平分给对方的公司代码
    Dim szHalveCompanyName As String
    Dim dbHalveRatio As Double '加总平分的比率
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    
    Dim aszFormulaValue() As String
    Dim nCount As Integer '票价项数
    
    
'    Dim szFormuarName As String '公式名称
    
    
    Dim tResultInfo As TSplitResult '返回的结果信息
    
    Dim atCompanyInfo() As TCompnaySettle '公司结算明细
    Dim atVehicleInfo() As TVehilceSettle '车辆结算明细
    Dim atBusInfo() As TBusSettle '车次结算明细
'    Dim atStationInfo() As TSettleSheetStation '站点结算明细
    
    Dim dbSplitItem As Double '临时拆出来的某一项的金额
    Dim dbTotalTicketPrice As Double '总票款
    Dim dbTotalQuantity As Double '总人数
    
    Dim dbSettleOtherCompanyPrice As Double '总的结给对方公司金额
    Dim dbSettleStationPrice As Double '总的结给车站的金额
        
    Dim dbSplitlimitCharge As Double '底限额
    Dim dbSplitUpCharge As Double '上限额
        
        
    Dim bIsDefaut As Boolean '是否默认协议
    
    Dim nVailableNumber As Integer '可用的结算项的数目
    
    
    Dim nItem As Integer
    Dim szitem As String
    
    
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_SplitCheckSheet '验证用户是否有结算的权限
    
    oDB.ConnectionString = GetConnectionStr()
    
    
    'if 公司代码不为'' then 按公司结算elseif 车辆代码不为'' then 按车辆结算else      出调用参数不正确错endif
    If pszCompanyID = "" And pszVehicleID = "" And pszBusID = "" Then
        '出传入参数错
        RaiseError ERR_PreviewSplitInputParam
    ElseIf pszCompanyID <> "" Then
        bSplitByCompany = True
    ElseIf pszVehicleID <> "" Then
        bSplitByVehicle = True
    ElseIf pszBusID <> "" Then
        '按车次结算
        bSplitByCompany = False
        bSplitByVehicle = False
    End If
    nSheetCount = ArrayLength(paszCheckSheetID)
    If nSheetCount = 0 Then RaiseError ERR_PreviewSplitInputParam '出传入参数错
    
    szWhere = " ( "
    For i = 1 To nSheetCount - 1
            szWhere = szWhere & "'" & paszCheckSheetID(i) & "',"
    Next i
    szWhere = szWhere & "'" & paszCheckSheetID(i) & "')"
    
    
    
    '判断每一张路单是否均存在，如不存在则出"路单不存在"的错。如无效或经过结算。则出"路单无效或路单已经结算"的错。
    szSql = "SELECT * FROM " & cszTableCheckSheet & "  WHERE check_sheet_id IN " & szWhere
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount <> nSheetCount Then
        '如果记录数不为传入的路单数,则出错
        RaiseError ERR_SomeSheetNotExist
    Else
    
        '因为允许把其他车辆的路单使用电脑结算,故不需要再判断.
    
'        '判断所有的路单是不是同一公司或同一车辆
'        For i = 1 To rsTemp.RecordCount
'            If bSplitByCompany Then
'                If FormatDbValue(rsTemp!transport_company_id) <> pszCompanyID Then
'                    '出错
'                    RaiseError ERR_SplitSheetMustSameCompany
'                End If
'            Else
'                If UCase(FormatDbValue(rsTemp!vehicle_id)) <> UCase(pszVehicleID) Then
'                    '出错
'                    RaiseError ERR_SplitSheetMustSameVehicle
'                End If
'            End If
'            rsTemp.MoveNext
'        Next i
    End If
    
    
    '因此句语句在路单有300张左右时,效率太低,故暂时不验证
'    szSql = "SELECT * FROM " & cszTableCheckSheet & " WHERE (settlement_status = " & TransFieldValueToString(CS_CheckSheetSettled) _
'        & " OR valid_mark = " & TransFieldValueToString(CS_CheckSheetInvalid) & " )" _
'        & " AND check_sheet_id IN " & szWhere
'
'    Set rsTemp = odb.Execute(szSql)
'    If rsTemp.RecordCount <> 0 Then
'        '如果存在已废或已结算的路单,则出错
'        RaiseError ERR_SomeSheetHasCancelORSettle
'    End If
    
    
    szSql = " SELECT * FROM price_item_info "
    Set rsTemp = oDB.Execute(szSql)
    
    nCount = rsTemp.RecordCount
    
    aszFormulaValue = GetFormulaItem
    
                          
    '取出默认的协议
    bIsDefaut = False
    Set rsDefaultProtocol = GetProtocolFormulaCompany(, True)
    
    
    szSql = " SELECT * FROM split_item_info WHERE split_status = " & TransFieldValueToString(ESplitItemUseStatus.CS_SplitItemUse)
    Set rsTemp = oDB.Execute(szSql)
    nVailableNumber = rsTemp.RecordCount
    
    
    m_bIsNew = False
    
    
    '如果按公司结算
    If bSplitByCompany Then
'        '1．   汇总出各公司、车型、上车站、线路、站点、票种的人数及各票价项信息prsPassageNumber
'        szSql = "SELECT " & TransFieldValueToString(pszCompanyID) & "  transport_company_id , c.vehicle_type_code , c.route_id , i.station_id sell_station_id , s.station_id , s.ticket_type " _
'            & " , MAX(transport_company_short_name) transport_company_short_name , MAX(route_name) route_name  " _
'            & " , MAX(i.station_name) sell_station_name , MAX(f.station_name) station_name , MAX(vehicle_type_short_name) vehicle_type_short_name  " _
'            & " , MAX(ticket_type_name) ticket_type_name , MAX(a.annotation) annotation " _
'            & " , SUM(quantity) quantity , SUM(s.mileage*quantity) mileage , SUM(ticket_price*quantity) ticket_price , SUM(base_carriage*quantity) base_carriage " _
'            & " , SUM(price_item_1*quantity) price_item_1 , SUM(price_item_2*quantity) price_item_2 , SUM(price_item_3*quantity) price_item_3 , SUM(price_item_4*quantity) price_item_4 " _
'            & " , SUM(price_item_5*quantity) price_item_5 , SUM(price_item_6*quantity) price_item_6 , SUM(price_item_7*quantity) price_item_7 , SUM(price_item_8*quantity) price_item_8 " _
'            & " , SUM(price_item_9*quantity) price_item_9 , SUM(price_item_10*quantity) price_item_10 , SUM(price_item_11*quantity) price_item_11 , SUM(price_item_12*quantity) price_item_12 " _
'            & " , SUM(price_item_13*quantity) price_item_13 , SUM(price_item_14*quantity) price_item_14 , SUM(price_item_15*quantity) price_item_15 " _
'            & " FROM " & cszTableCheckSheet & " c, " & cszTableSheetStation & " s , " & cszTableCompany & " p , " & cszTableRouteInfo & " r " _
'            & " , " & cszTableStationInfo & "  i , " & cszTableStationInfo & " f , " & cszTableTicketType & " k , " & cszTableVehicleTypeCode & " d " _
'            & " , area_code a, sell_station_info t  " _
'            & " WHERE c.check_sheet_id = s.check_sheet_id AND c.vehicle_type_code = d.vehicle_type_code AND c.route_id = r.route_id " _
'            & " AND s.sell_station_id = t.sell_station_id AND t.station_id = i.station_id AND s.station_id = f.station_id AND s.ticket_type = k.ticket_type_id " _
'            & " AND p.transport_company_id = " & TransFieldValueToString(pszCompanyID) & "  AND f.area_code = a.area_code " _
'            & " AND c.check_sheet_id IN  " & szWhere _
'            & " GROUP BY c.vehicle_type_code , c.route_id , i.station_id , s.station_id , s.ticket_type ORDER BY c.route_id, c.vehicle_type_code  , i.station_id  , s.station_id , s.ticket_type "
'
'        Set prsPassageNumber = oDb.Execute(szSql)
        
        '2．   得到所有该公司需要与之加总平分的公司线路信息rsHalve
        szSql = " SELECT h.*, c.transport_company_short_name FROM fin_halve_company_lst h, " & cszTableCompany & " c " _
            & " WHERE c.transport_company_id = h.transport_company_id AND h.transport_company_id = " & TransFieldValueToString(pszCompanyID)
        
        Set rsHalve = oDB.Execute(szSql)
        
        '3． 得到该公司的协议
        Set rsProtocol = GetProtocolFormulaCompany(pszCompanyID)
        If rsProtocol.RecordCount = 0 Then
            '取出默认的协议
            If rsDefaultProtocol.RecordCount <> 0 Then
                Set rsProtocol = rsDefaultProtocol
                bIsDefaut = True
            Else
                '该公司协议未设置
                RaiseError ERR_CompanyNotSetProtocol
            End If
        Else
            bIsDefaut = False
        End If
        
        
        For i = 1 To prsPassageNumber.RecordCount
            szHalveCompanyID = ""
            If rsHalve.RecordCount > 0 Then rsHalve.MoveFirst
            For j = 1 To rsHalve.RecordCount
                If FormatDbValue(prsPassageNumber!route_id) = FormatDbValue(rsHalve!route_id) And FormatDbValue(prsPassageNumber!transport_company_id) = FormatDbValue(rsHalve!transport_company_id) Then
                    '找到该公司在该线路上需要加总平分
                    
                    szHalveCompanyID = FormatDbValue(rsHalve!other_transport_company_id)
                    szHalveCompanyName = GetHavleOtherCompanyName(szHalveCompanyID)
                    
                    dbHalveRatio = FormatDbValue(rsHalve!Ratio)
                    Exit For
                End If
                rsHalve.MoveNext
            Next j
            'If 该公司在rsHalve中找到需要加总平分then
            
            If szHalveCompanyID = "" Or pbIsSingle Then
                '不需要加总平分
                '找到该线路的起始协议
                
                FindRouteProtocol rsProtocol, FormatDbValue(prsPassageNumber!route_id), bIsDefaut
                
                '如果公式中没有用到结算价,则不进行查询结算价
                If FindSettlePriceStr(rsProtocol, nVailableNumber, cszSettleStr) Then
                    Set rsCompanySettlePrice = GetCompanySettlePrice(prsPassageNumber!transport_company_id, prsPassageNumber!vehicle_type_code, prsPassageNumber!route_id, prsPassageNumber!sell_station_id, prsPassageNumber!station_id)
                End If
                
                For j = 1 To nVailableNumber
                    
                    SetFormulaValueCompany prsPassageNumber, aszFormulaValue, nCount, rsCompanySettlePrice, rsProtocol, atCompanyInfo, pszCompanyID, FormatDbValue(prsPassageNumber!transport_company_short_name), False, bIsDefaut, 1
                    
                    rsProtocol.MoveNext
                    m_bIsNew = True
                Next j
'                atCompanyInfo(1).PassengerNumber = atCompanyInfo(1).PassengerNumber + FormatDbValue(prsPassageNumber!Quantity)

            Else
                '将这些人数及售票的票价项分成两半，一半(此比率根据fin_halve_company_lst中的设定比率)拆给该公司，剩下部分拆给加总平分的公司。
                
                '先将应拆给该公司的部分拆给该公司
                '找到该线路的起始协议
                FindRouteProtocol rsProtocol, FormatDbValue(prsPassageNumber!route_id), bIsDefaut
                '如果公式中没有用到结算价,则不进行查询结算价
                If FindSettlePriceStr(rsProtocol, nVailableNumber, cszSettleStr) Then
                    Set rsCompanySettlePrice = GetCompanySettlePrice(prsPassageNumber!transport_company_id, prsPassageNumber!vehicle_type_code, prsPassageNumber!route_id, prsPassageNumber!sell_station_id, prsPassageNumber!station_id)
                End If
                
                For j = 1 To nVailableNumber
                    SetFormulaValueCompany prsPassageNumber, aszFormulaValue, nCount, rsCompanySettlePrice, rsProtocol, atCompanyInfo, pszCompanyID, FormatDbValue(prsPassageNumber!transport_company_short_name), True, bIsDefaut, dbHalveRatio
                    
                    rsProtocol.MoveNext
                    m_bIsNew = True
                Next j
                
                
                '将剩下部分拆给加总平分的公司。
                Dim bIsHalveDefault As Boolean
                '取出协议
                Set rsHalveProtocol = GetProtocolFormulaCompany(szHalveCompanyID)
                If rsHalveProtocol.RecordCount = 0 Then
                    '取出默认的协议
                    If rsDefaultProtocol.RecordCount <> 0 Then
                        Set rsHalveProtocol = rsDefaultProtocol
                        bIsHalveDefault = True
                    Else
                        '该公司协议未设置
                        RaiseError ERR_CompanyNotSetProtocol
                    End If
                Else
                    bIsHalveDefault = False
                End If
                
                
                
                m_bIsNew = False
                '找到该线路的起始协议
                FindRouteProtocol rsHalveProtocol, FormatDbValue(prsPassageNumber!route_id), bIsHalveDefault
                
                
                '如果公式中没有用到结算价,则不进行查询结算价
                If FindSettlePriceStr(rsHalveProtocol, nVailableNumber, cszSettleStr) Then
                    Set rsCompanySettlePrice = GetCompanySettlePrice(szHalveCompanyID, prsPassageNumber!vehicle_type_code, prsPassageNumber!route_id, prsPassageNumber!sell_station_id, prsPassageNumber!station_id)
                End If
                

                For j = 1 To nVailableNumber
                    SetFormulaValueCompany prsPassageNumber, aszFormulaValue, nCount, rsCompanySettlePrice, rsHalveProtocol, atCompanyInfo, szHalveCompanyID, szHalveCompanyName, True, bIsHalveDefault, 1 - dbHalveRatio
                    rsHalveProtocol.MoveNext
                    m_bIsNew = True
                Next j
                   
            End If
            
'            dbTotalTicketPrice = dbTotalTicketPrice + FormatDbValue(prsPassageNumber!ticket_price)
'            dbTotalQuantity = dbTotalQuantity + FormatDbValue(prsPassageNumber!Quantity)
            
            dbTotalTicketPrice = FormatDbValue(prsPassageNumber!ticket_price)
            dbTotalQuantity = FormatDbValue(prsPassageNumber!Quantity)
            tResultInfo.SettleSheetInfo.TotalTicketPrice = tResultInfo.SettleSheetInfo.TotalTicketPrice + dbTotalTicketPrice
            tResultInfo.SettleSheetInfo.TotalQuantity = tResultInfo.SettleSheetInfo.TotalQuantity + dbTotalQuantity
            
            prsPassageNumber.MoveNext
            m_bIsNew = False
        Next i
    
    
    
    ElseIf bSplitByVehicle Then
        
        'else 按车辆结算
        Dim szCompanyID As String
        
            
'        '1.汇总出各车辆、上车站、线路、站点、票种的人数及各票价项prsPassageNumber。
'        szSql = "SELECT " & TransFieldValueToString(pszVehicleID) & " vehicle_id  , c.route_id , i.station_id sell_station_id , s.station_id , s.ticket_type " _
'            & " , MAX(license_tag_no) license_tag_no , MAX(route_name) route_name  " _
'            & " , MAX(i.station_name) sell_station_name , MAX(f.station_name) station_name,MAX(d.vehicle_type_code) vehicle_type_code  " _
'            & " , MAX(ticket_type_name) ticket_type_name , MAX(c.transport_company_id) transport_company_id " _
'            & " , MAX(m.transport_company_short_name) transport_company_short_name , MAX(a.annotation) annotation " _
'            & " , SUM(quantity) quantity , SUM(s.mileage*quantity) mileage , SUM(ticket_price*quantity) ticket_price , SUM(base_carriage*quantity) base_carriage " _
'            & " , SUM(price_item_1*quantity) price_item_1 , SUM(price_item_2*quantity) price_item_2 , SUM(price_item_3*quantity) price_item_3 , SUM(price_item_4*quantity) price_item_4 " _
'            & " , SUM(price_item_5*quantity) price_item_5 , SUM(price_item_6*quantity) price_item_6 , SUM(price_item_7*quantity) price_item_7 , SUM(price_item_8*quantity) price_item_8 " _
'            & " , SUM(price_item_9*quantity) price_item_9 , SUM(price_item_10*quantity) price_item_10 , SUM(price_item_11*quantity) price_item_11 , SUM(price_item_12*quantity) price_item_12 " _
'            & " , SUM(price_item_13*quantity) price_item_13 , SUM(price_item_14*quantity) price_item_14 , SUM(price_item_15*quantity) price_item_15 " _
'            & " FROM " & cszTableCheckSheet & " c, " & cszTableSheetStation & " s , " & cszTableVehicle & " p , " & cszTableRouteInfo & " r " _
'            & " , " & cszTableStationInfo & "  i , " & cszTableStationInfo & " f , " & cszTableTicketType & " k ," & cszTableCompany & " m " _
'            & " , " & cszTableVehicleTypeCode & " d , area_code a , sell_station_info t  " _
'            & " WHERE c.vehicle_type_code=d.vehicle_type_code AND p.vehicle_id=" & TransFieldValueToString(pszVehicleID) & " AND c.check_sheet_id = s.check_sheet_id " _
'            & " AND c.route_id = r.route_id AND m.transport_company_id=c.transport_company_id " _
'            & " AND s.sell_station_id = t.sell_station_id AND t.station_id = i.station_id AND s.station_id = f.station_id AND s.ticket_type = k.ticket_type_id " _
'            & " AND f.area_code = a.area_code " _
'            & " AND c.check_sheet_id IN  " & szWhere _
'            & " GROUP BY c.route_id , i.station_id , s.station_id , s.ticket_type ORDER BY c.route_id , i.station_id , s.station_id , s.ticket_type "
'
            
'        Set prsPassageNumber = oDb.Execute(szSql)
        
        szCompanyID = FormatDbValue(prsPassageNumber!transport_company_id)
        '2.  得到需要结算的车辆所属公司的加总平分公司信息。
        
        szSql = " SELECT h.*, c.transport_company_short_name FROM fin_halve_company_lst h, " & cszTableCompany & " c " _
            & " WHERE c.transport_company_id = h.transport_company_id AND c.transport_company_id = " & TransFieldValueToString(szCompanyID)
            
        Set rsHalve = oDB.Execute(szSql)
        '3． 得到该车辆的协议
        Set rsProtocol = GetProtocolFormulaVehicle(pszVehicleID, False)
        bIsDefaut = False
        If rsProtocol.RecordCount = 0 Then
            '取出默认的协议
            If rsDefaultProtocol.RecordCount <> 0 Then
                Set rsProtocol = rsDefaultProtocol
                bIsDefaut = True
            Else
                '该车辆协议未设置
                RaiseError ERR_VehicleNotSetProtocol
            End If
        End If
    prsPassageNumber.MoveFirst
        '4． For I = 1 to prsPassageNumber.RecordCount
        For i = 1 To prsPassageNumber.RecordCount
            
            
            szHalveCompanyID = ""
            For j = 1 To rsHalve.RecordCount
                If FormatDbValue(prsPassageNumber!route_id) = FormatDbValue(rsHalve!route_id) And FormatDbValue(prsPassageNumber!transport_company_id) = FormatDbValue(rsHalve!transport_company_id) Then
                    '找到该公司在该线路上需要加总平分
                    szHalveCompanyID = FormatDbValue(rsHalve!other_transport_company_id)
                    szHalveCompanyName = GetHavleOtherCompanyName(szHalveCompanyID)
                    dbHalveRatio = FormatDbValue(rsHalve!Ratio)
                    Exit For
                End If
            Next j
            
            If szHalveCompanyID = "" Then
                '不需要加总平分
                '将相应的人数及各售票票价只算给该车辆。
                rsProtocol.MoveFirst
                
                '如果公式中没有用到结算价,则不进行查询结算价
                If FindSettlePriceStr(rsProtocol, nVailableNumber, cszSettleStr) Then
                    Set rsVehicleSettlePrice = GetVehicleSettlePrice(prsPassageNumber!vehicle_id, prsPassageNumber!route_id, prsPassageNumber!sell_station_id, prsPassageNumber!station_id)
                End If
                                
                For j = 1 To nVailableNumber
                                
'    MsgBox "5,j: " & j & "  i," & i
                    SetFormulaValueVehicle prsPassageNumber, aszFormulaValue, nCount, rsVehicleSettlePrice, rsProtocol, atVehicleInfo, pszVehicleID, FormatDbValue(prsPassageNumber!license_tag_no), False, 0
                    rsProtocol.MoveNext
                    m_bIsNew = True
                Next j
            Else
                'Else累加该车辆的结算费用明细将这些人数及各售票票价项分为两半，一半算给该车辆，另一半算给加总平分中的另一公司。

                
                '先结给该车辆，方法与上同。
                
                '如果公式中没有用到结算价,则不进行查询结算价
                If FindSettlePriceStr(rsProtocol, nVailableNumber, cszSettleStr) Then
                    Set rsVehicleSettlePrice = GetVehicleSettlePrice(prsPassageNumber!vehicle_id, prsPassageNumber!route_id, prsPassageNumber!sell_station_id, prsPassageNumber!station_id)
                End If
                
                rsProtocol.MoveFirst
                For j = 1 To nVailableNumber
                    SetFormulaValueVehicle prsPassageNumber, aszFormulaValue, nCount, rsVehicleSettlePrice, rsProtocol, atVehicleInfo, pszVehicleID, FormatDbValue(prsPassageNumber!license_tag_no), False, dbHalveRatio
                    rsProtocol.MoveNext
                    m_bIsNew = True
                Next j
                '再结给该公司，与按公司结算类似
                '取出协议
                bIsHalveDefault = False
                Set rsHalveProtocol = GetProtocolFormulaCompany(szHalveCompanyID)
                If rsHalveProtocol.RecordCount = 0 Then
                    '取出默认的协议
                    If rsDefaultProtocol.RecordCount <> 0 Then
                        Set rsHalveProtocol = rsDefaultProtocol
                        bIsHalveDefault = True
                    Else
                        '该公司协议未设置
                        RaiseError ERR_VehicleNotSetProtocol
                    End If
                End If
                
                '如果公式中没有用到结算价,则不进行查询结算价
                If FindSettlePriceStr(rsHalveProtocol, nVailableNumber, cszSettleStr) Then
                    Set rsVehicleSettlePrice = GetVehicleSettlePrice(prsPassageNumber!vehicle_id, prsPassageNumber!route_id, prsPassageNumber!sell_station_id, prsPassageNumber!station_id)
                End If
                
                m_bIsNew = False
                rsHalveProtocol.MoveFirst
                FindRouteProtocol rsHalveProtocol, FormatDbValue(prsPassageNumber!route_id), bIsHalveDefault
                For j = 1 To nVailableNumber
                    SetFormulaValueCompany prsPassageNumber, aszFormulaValue, nCount, rsCompanySettlePrice, rsHalveProtocol, atCompanyInfo, szHalveCompanyID, szHalveCompanyName, True, bIsHalveDefault, 1 - dbHalveRatio
                    rsHalveProtocol.MoveNext
                    m_bIsNew = True
                Next j
            End If
            
            dbTotalTicketPrice = FormatDbValue(prsPassageNumber!ticket_price)
            dbTotalQuantity = FormatDbValue(prsPassageNumber!Quantity)
            tResultInfo.SettleSheetInfo.TotalTicketPrice = tResultInfo.SettleSheetInfo.TotalTicketPrice + dbTotalTicketPrice
            tResultInfo.SettleSheetInfo.TotalQuantity = tResultInfo.SettleSheetInfo.TotalQuantity + dbTotalQuantity
            prsPassageNumber.MoveNext
            m_bIsNew = False
        Next i
                
    ElseIf bSplitByCompany = False And bSplitByVehicle = False Then
        'else 按车次结算
'        Dim szCompanyID As String
        
            
'        '1.汇总出各车辆、上车站、线路、站点、票种的人数及各票价项prsPassageNumber。
'        szSql = "SELECT " & TransFieldValueToString(pszVehicleID) & " vehicle_id  , c.route_id , i.station_id sell_station_id , s.station_id , s.ticket_type " _
'            & " , MAX(license_tag_no) license_tag_no , MAX(route_name) route_name  " _
'            & " , MAX(i.station_name) sell_station_name , MAX(f.station_name) station_name,MAX(d.vehicle_type_code) vehicle_type_code  " _
'            & " , MAX(ticket_type_name) ticket_type_name , MAX(c.transport_company_id) transport_company_id " _
'            & " , MAX(m.transport_company_short_name) transport_company_short_name , MAX(a.annotation) annotation " _
'            & " , SUM(quantity) quantity , SUM(s.mileage*quantity) mileage , SUM(ticket_price*quantity) ticket_price , SUM(base_carriage*quantity) base_carriage " _
'            & " , SUM(price_item_1*quantity) price_item_1 , SUM(price_item_2*quantity) price_item_2 , SUM(price_item_3*quantity) price_item_3 , SUM(price_item_4*quantity) price_item_4 " _
'            & " , SUM(price_item_5*quantity) price_item_5 , SUM(price_item_6*quantity) price_item_6 , SUM(price_item_7*quantity) price_item_7 , SUM(price_item_8*quantity) price_item_8 " _
'            & " , SUM(price_item_9*quantity) price_item_9 , SUM(price_item_10*quantity) price_item_10 , SUM(price_item_11*quantity) price_item_11 , SUM(price_item_12*quantity) price_item_12 " _
'            & " , SUM(price_item_13*quantity) price_item_13 , SUM(price_item_14*quantity) price_item_14 , SUM(price_item_15*quantity) price_item_15 " _
'            & " FROM " & cszTableCheckSheet & " c, " & cszTableSheetStation & " s , " & cszTableVehicle & " p , " & cszTableRouteInfo & " r " _
'            & " , " & cszTableStationInfo & "  i , " & cszTableStationInfo & " f , " & cszTableTicketType & " k ," & cszTableCompany & " m " _
'            & " , " & cszTableVehicleTypeCode & " d , area_code a , sell_station_info t  " _
'            & " WHERE c.vehicle_type_code=d.vehicle_type_code AND p.vehicle_id=" & TransFieldValueToString(pszVehicleID) & " AND c.check_sheet_id = s.check_sheet_id " _
'            & " AND c.route_id = r.route_id AND m.transport_company_id=c.transport_company_id " _
'            & " AND s.sell_station_id = t.sell_station_id AND t.station_id = i.station_id AND s.station_id = f.station_id AND s.ticket_type = k.ticket_type_id " _
'            & " AND f.area_code = a.area_code " _
'            & " AND c.check_sheet_id IN  " & szWhere _
'            & " GROUP BY c.route_id , i.station_id , s.station_id , s.ticket_type ORDER BY c.route_id , i.station_id , s.station_id , s.ticket_type "
'
            
'        Set prsPassageNumber = oDb.Execute(szSql)
        
        szCompanyID = FormatDbValue(prsPassageNumber!transport_company_id)
        '2.  得到需要结算的车辆所属公司的加总平分公司信息。
        
        szSql = " SELECT h.*, c.transport_company_short_name FROM fin_halve_company_lst h, " & cszTableCompany & " c " _
            & " WHERE c.transport_company_id = h.transport_company_id AND c.transport_company_id = " & TransFieldValueToString(szCompanyID)
            
        Set rsHalve = oDB.Execute(szSql)
        '3． 得到该车辆的协议
        Set rsProtocol = GetProtocolFormulaBus(pszBusID, szCompanyID, False)
        bIsDefaut = False
        If rsProtocol.RecordCount = 0 Then
            '取出默认的协议
            If rsDefaultProtocol.RecordCount <> 0 Then
                Set rsProtocol = rsDefaultProtocol
                bIsDefaut = True
            Else
                '该车次协议未设置
                RaiseError ERR_VehicleNotSetProtocol
            End If
        End If
        prsPassageNumber.MoveFirst
        '4． For I = 1 to prsPassageNumber.RecordCount
        For i = 1 To prsPassageNumber.RecordCount
            
            
            szHalveCompanyID = ""
            For j = 1 To rsHalve.RecordCount
                If FormatDbValue(prsPassageNumber!route_id) = FormatDbValue(rsHalve!route_id) And FormatDbValue(prsPassageNumber!transport_company_id) = FormatDbValue(rsHalve!transport_company_id) Then
                    '找到该公司在该线路上需要加总平分
                    szHalveCompanyID = FormatDbValue(rsHalve!other_transport_company_id)
                    szHalveCompanyName = GetHavleOtherCompanyName(szHalveCompanyID)
                    dbHalveRatio = FormatDbValue(rsHalve!Ratio)
                    Exit For
                End If
            Next j
            
            If szHalveCompanyID = "" Then
                '不需要加总平分
                '将相应的人数及各售票票价只算给该车辆。
                rsProtocol.MoveFirst
                
                '如果公式中没有用到结算价,则不进行查询结算价
                If FindSettlePriceStr(rsProtocol, nVailableNumber, cszSettleStr) Then
                    Set rsBusSettlePrice = GetBusSettlePrice(prsPassageNumber!bus_id, prsPassageNumber!route_id, prsPassageNumber!sell_station_id, prsPassageNumber!station_id)
                End If
                                
                For j = 1 To nVailableNumber
                                
                    SetFormulaValueBus prsPassageNumber, aszFormulaValue, nCount, rsBusSettlePrice, rsProtocol, atBusInfo, pszBusID, FormatDbValue(prsPassageNumber!license_tag_no), False, 0
                    rsProtocol.MoveNext
                    m_bIsNew = True
                Next j
            Else
                'Else累加该车辆的结算费用明细将这些人数及各售票票价项分为两半，一半算给该车辆，另一半算给加总平分中的另一公司。

                
                '先结给该车辆，方法与上同。
                
                '如果公式中没有用到结算价,则不进行查询结算价
                If FindSettlePriceStr(rsProtocol, nVailableNumber, cszSettleStr) Then
                    Set rsBusSettlePrice = GetBusSettlePrice(prsPassageNumber!bus_id, prsPassageNumber!route_id, prsPassageNumber!sell_station_id, prsPassageNumber!station_id)
                End If
                
                rsProtocol.MoveFirst
                For j = 1 To nVailableNumber
                    SetFormulaValueBus prsPassageNumber, aszFormulaValue, nCount, rsBusSettlePrice, rsProtocol, atBusInfo, pszBusID, FormatDbValue(prsPassageNumber!license_tag_no), False, dbHalveRatio
                    rsProtocol.MoveNext
                    m_bIsNew = True
                Next j
                '再结给该公司，与按公司结算类似
                '取出协议
                bIsHalveDefault = False
                Set rsHalveProtocol = GetProtocolFormulaCompany(szHalveCompanyID)
                If rsHalveProtocol.RecordCount = 0 Then
                    '取出默认的协议
                    If rsDefaultProtocol.RecordCount <> 0 Then
                        Set rsHalveProtocol = rsDefaultProtocol
                        bIsHalveDefault = True
                    Else
                        '该公司协议未设置
                        RaiseError ERR_VehicleNotSetProtocol
                    End If
                End If
                
                '如果公式中没有用到结算价,则不进行查询结算价
                If FindSettlePriceStr(rsHalveProtocol, nVailableNumber, cszSettleStr) Then
                    Set rsBusSettlePrice = GetBusSettlePrice(prsPassageNumber!bus_id, prsPassageNumber!route_id, prsPassageNumber!sell_station_id, prsPassageNumber!station_id)
                End If
                
                m_bIsNew = False
                rsHalveProtocol.MoveFirst
                FindRouteProtocol rsHalveProtocol, FormatDbValue(prsPassageNumber!route_id), bIsHalveDefault
                For j = 1 To nVailableNumber
                    SetFormulaValueCompany prsPassageNumber, aszFormulaValue, nCount, rsCompanySettlePrice, rsHalveProtocol, atCompanyInfo, szHalveCompanyID, szHalveCompanyName, True, bIsHalveDefault, 1 - dbHalveRatio
                    rsHalveProtocol.MoveNext
                    m_bIsNew = True
                Next j
                
            End If
            
            dbTotalTicketPrice = FormatDbValue(prsPassageNumber!ticket_price)
            dbTotalQuantity = FormatDbValue(prsPassageNumber!Quantity)
            tResultInfo.SettleSheetInfo.TotalTicketPrice = tResultInfo.SettleSheetInfo.TotalTicketPrice + dbTotalTicketPrice
            tResultInfo.SettleSheetInfo.TotalQuantity = tResultInfo.SettleSheetInfo.TotalQuantity + dbTotalQuantity
            prsPassageNumber.MoveNext
            m_bIsNew = False
        Next i
                
    End If
    '=========================
    '以下为处理计算好的结果集
    '=========================
    
    
    
    '将总路单数、总票款、总人数、应结票款、结给车站票款以及公司的费用明细和车辆明细均返回
    With tResultInfo.SettleSheetInfo
        
        .CheckSheetCount = ArrayLength(paszCheckSheetID) '路单数
        
        
        
        '==============================================================
        '以下为处理手工补票的操作
        Dim dbExtraTotalTicketPrice  As Double '补票金额
        Dim dbExtraServicePrice As Double '劳务费
        Dim dbExtraSettleOutPrice As Double '拆出金额
        Dim nTotalCount As Integer '
        
        dbExtraTotalTicketPrice = 0 '补票金额
        dbExtraServicePrice = 0 '劳务费
        dbExtraSettleOutPrice = 0 '拆出金额
        
        '汇总得到总的 补票金额、劳务费、拆出金额
        For i = 1 To ArrayLength(patExtraInfo)
            dbExtraTotalTicketPrice = dbExtraTotalTicketPrice + patExtraInfo(i).TotalTicketPrice
            '直接把总价当作运费来算
            
            dbExtraServicePrice = dbExtraServicePrice + patExtraInfo(i).ServicePrice
            dbExtraSettleOutPrice = dbExtraSettleOutPrice + patExtraInfo(i).SettleOutPrice
        Next i
        '此处偷懒下，按结果集的车次数、公司数、车辆数，进行累加，再平均分配
        
        '汇总共有的数目
        nTotalCount = ArrayLength(atCompanyInfo) + ArrayLength(atVehicleInfo) + ArrayLength(atBusInfo)
        If nTotalCount > 0 Then
'            dbExtraTotalTicketPrice = FormatMoney(dbExtraTotalTicketPrice / nTotalCount)
            dbExtraServicePrice = FormatMoney(dbExtraServicePrice / nTotalCount)
            dbExtraSettleOutPrice = FormatMoney(dbExtraTotalTicketPrice / nTotalCount)
        End If
        
        '以上为处理手工补票的操作
        '==============================================================
        
        
        
        
        
        
        
        
        
        
        If pszCompanyID <> "" Then
            .ObjectID = pszCompanyID '对象代码
            .SettleObject = CS_SettleByTransportCompany  '结算对象
            
        ElseIf pszVehicleID <> "" Then
            .ObjectID = pszVehicleID
            .SettleObject = CS_SettleByVehicle
        ElseIf pszBusID <> "" Then
            .ObjectID = pszBusID
            .SettleObject = CS_SettleByBus
        End If
        
        
        
        '========================================================================
        '处理固定费用
        '========================================================================
        
        
        
        If .SettleObject = CS_SettleByVehicle And m_nFixFeeItemCount > 0 Then
            '如果是按照车辆结算,并且固定费用项存在 ,则读取该车辆的所有的固定费用项
            '============================
            Dim rsFixFee As Recordset
            szSql = " SELECT SUM(split_item_1) split_item_1 , SUM(split_item_2) split_item_2 , SUM(split_item_3) split_item_3 " _
                & " , SUM(split_item_4) split_item_4, SUM(split_item_5) split_item_5, SUM(split_item_6) split_item_6, SUM(split_item_7) split_item_7 " _
                & " , SUM(split_item_8) split_item_8, SUM(split_item_9) split_item_9, SUM(split_item_10) split_item_10, SUM(split_item_11) split_item_11 " _
                & " , SUM(split_item_12) split_item_12, SUM(split_item_13) split_item_13, SUM(split_item_14) split_item_14, SUM(split_item_15) split_item_15 " _
                & " , SUM(split_item_16) split_item_16, SUM(split_item_17) split_item_17, SUM(split_item_18) split_item_18, SUM(split_item_19) split_item_19 " _
                & " , SUM(split_item_20) split_item_20 " _
                & " FROM fin_vehicle_fix_fee_lst" _
                & " WHERE vehicle_id = " & TransFieldValueToString(pszVehicleID) _
                & " AND bus_date <" & TransFieldValueToString(pdtEndDate) _
                & " AND is_dec = " & TransFieldValueToString(False)

            '& " AND bus_date >= " & TransFieldValueToString(pdtStartDate)

            Set rsFixFee = oDB.Execute(szSql)
            
            
        ElseIf .SettleObject = CS_SettleByBus And m_nFixFeeItemCount > 0 Then
            '如果是按照车次结算,并且固定费用项存在 ,则读取该车次的所有的固定费用项
            '============================
'            Dim rsFixFee As Recordset
            szSql = " SELECT SUM(split_item_1) split_item_1 , SUM(split_item_2) split_item_2 , SUM(split_item_3) split_item_3 " _
                & " , SUM(split_item_4) split_item_4, SUM(split_item_5) split_item_5, SUM(split_item_6) split_item_6, SUM(split_item_7) split_item_7 " _
                & " , SUM(split_item_8) split_item_8, SUM(split_item_9) split_item_9, SUM(split_item_10) split_item_10, SUM(split_item_11) split_item_11 " _
                & " , SUM(split_item_12) split_item_12, SUM(split_item_13) split_item_13, SUM(split_item_14) split_item_14, SUM(split_item_15) split_item_15 " _
                & " , SUM(split_item_16) split_item_16, SUM(split_item_17) split_item_17, SUM(split_item_18) split_item_18, SUM(split_item_19) split_item_19 " _
                & " , SUM(split_item_20) split_item_20 " _
                & " FROM fin_bus_fix_fee_lst" _
                & " WHERE bus_id = " & TransFieldValueToString(pszBusID) _
                & " AND bus_date <" & TransFieldValueToString(pdtEndDate) _
                & " AND is_dec = " & TransFieldValueToString(False)

            '& " AND bus_date >= " & TransFieldValueToString(pdtStartDate)

            Set rsFixFee = oDB.Execute(szSql)


        End If
        
        
        
        
        '赋公司结算结果
        For i = 1 To ArrayLength(atCompanyInfo)
            rsProtocol.MoveFirst
'            当结算有底限时 , 现在的处理方式是不对的, 未判断不同公司, 线路不同, 其协议不同, 协议不同时当然其底限也是不同的
'            If bSplitByCompany Then
'               FindRouteProtocol rsProtocol, rsCompanySettlePrice, bIsDefaut
'            End If
            For j = 1 To nVailableNumber 'rsProtocol.RecordCount
                szitem = FormatDbValue(rsProtocol!split_item_id)
                nItem = CInt(szitem)
                
                
                
                '将手工补票的数据进行累加
                If nItem = m_nCarriageItemPosition Then
                    '运费 拆出款加到运费的项目处
                    atCompanyInfo(i).SplitItem(nItem) = FormatMoney(atCompanyInfo(i).SplitItem(nItem) + dbExtraSettleOutPrice)
                    
                ElseIf nItem = m_nServiceItemPosition Then
                    '劳务费 劳务费加到劳务费的对应项目处
                    atCompanyInfo(i).SplitItem(nItem) = FormatMoney(atCompanyInfo(i).SplitItem(nItem) + dbExtraServicePrice)
                End If
                
                
                
                dbSplitlimitCharge = GetlimitCharge(FormatDbValue(rsProtocol!protocol_id), szitem, dbSplitUpCharge)
                '如果上限不为0,则   该项值=上限值-该项值
                If dbSplitUpCharge <> 0 Then
                    
                    'atCompanyInfo(i).SplitItem(nItem) = FormatMoney(dbSplitUpCharge - atCompanyInfo(i).SplitItem(nItem))
                    '针对椒江的情况进行修改
                    atCompanyInfo(i).SplitItem(nItem) = FormatMoney(atCompanyInfo(i).SplitItem(nItem) - dbSplitUpCharge)
                    If atCompanyInfo(i).SplitItem(nItem) < 0 Then
                        atCompanyInfo(i).SplitItem(nItem) = 0
                    End If
                End If
                '当下限值大于该项值时,该值为下限值
                If dbSplitlimitCharge > atCompanyInfo(i).SplitItem(nItem) Then
                    atCompanyInfo(i).SplitItem(nItem) = dbSplitlimitCharge
                Else
'                    atCompanyInfo(i).SplitItem(nItem) = atCompanyInfo(i).SplitItem(nItem)
                End If
                atCompanyInfo(i).SplitItem(nItem) = FormatMoney(atCompanyInfo(i).SplitItem(nItem))
                atCompanyInfo(i).SettlePrice = atCompanyInfo(i).SettlePrice + GetSettlePrice(rsProtocol, atCompanyInfo(i).SplitItem(nItem))
                atCompanyInfo(i).SettleStationPrice = atCompanyInfo(i).SettleStationPrice + GetSettleStationPrice(rsProtocol, atCompanyInfo(i).SplitItem(nItem))
                rsProtocol.MoveNext
            Next j
            atCompanyInfo(i).SettlePrice = FormatMoney(atCompanyInfo(i).SettlePrice) ' FormatMoney (atCompanyInfo(i).SettlePrice + 0.01, 1)
            atCompanyInfo(i).SettleStationPrice = FormatMoney(atCompanyInfo(i).SettleStationPrice) ' Round(atCompanyInfo(i).SettleStationPrice + 0.01, 1)
            
            dbSettleOtherCompanyPrice = dbSettleOtherCompanyPrice + atCompanyInfo(i).SettlePrice
            dbSettleStationPrice = dbSettleStationPrice + atCompanyInfo(i).SettleStationPrice
        Next i
        
        
'        ******************
        '赋车辆结算结果
        For i = 1 To ArrayLength(atVehicleInfo)
            
            rsProtocol.MoveFirst
            For j = 1 To nVailableNumber 'rsProtocol.RecordCount

                szitem = FormatDbValue(rsProtocol!split_item_id)
                nItem = Val(szitem)
                
                
                
                '将手工补票的数据进行累加
                If nItem = m_nCarriageItemPosition Then
                    '运费 拆出款加到运费的项目处
                    atVehicleInfo(i).SplitItem(nItem) = FormatMoney(atVehicleInfo(i).SplitItem(nItem) + dbExtraSettleOutPrice)
                    
                ElseIf nItem = m_nServiceItemPosition Then
                    '劳务费 劳务费加到劳务费的对应项目处
                    atVehicleInfo(i).SplitItem(nItem) = FormatMoney(atVehicleInfo(i).SplitItem(nItem) + dbExtraServicePrice)
                End If
                
                
                
                
                If m_nFixFeeItemCount > 0 And Not (rsFixFee Is Nothing) Then
                    If rsFixFee.RecordCount > 0 Then
                    '如果存在固定费用项
                        For k = 0 To m_nFixFeeItemCount - 1
                            If nItem = m_aszFixFeeItem(k) Then
                                atVehicleInfo(i).SplitItem(nItem) = atVehicleInfo(i).SplitItem(nItem) + FormatDbValue(rsFixFee.Fields("split_item_" & nItem))
                                Exit For
                            End If
                        Next k
                    End If
                End If
                

'                dbSplitlimitCharge = GetlimitCharge(FormatDbValue(rsProtocol!protocol_id), szItem, dbSplitUpCharge)
'                '如果上限不为0,则   该项值=上限值-该项值
'                If dbSplitUpCharge <> 0 Then
'                    atVehicleInfo(i).SplitItem(nItem) = FormatMoney(dbSplitUpCharge - atVehicleInfo(i).SplitItem(nItem))
'                End If

                dbSplitlimitCharge = GetlimitCharge(FormatDbValue(rsProtocol!protocol_id), szitem, dbSplitUpCharge)
                '如果上限不为0,则   该项值=上限值-该项值
                If dbSplitUpCharge <> 0 Then
                    
                    'atCompanyInfo(i).SplitItem(nItem) = FormatMoney(dbSplitUpCharge - atCompanyInfo(i).SplitItem(nItem))
                    '针对椒江的情况进行修改
                    atVehicleInfo(i).SplitItem(nItem) = FormatMoney(atVehicleInfo(i).SplitItem(nItem) - dbSplitUpCharge)
                    If atVehicleInfo(i).SplitItem(nItem) < 0 Then
                        atVehicleInfo(i).SplitItem(nItem) = 0
                    End If
                End If
                
                
                '当下限值大于该项值时,该值为下限值
                If dbSplitlimitCharge > atVehicleInfo(i).SplitItem(nItem) Then
                    atVehicleInfo(i).SplitItem(nItem) = dbSplitlimitCharge
                Else
'                    atVehicleinfo(i).SplitItem(nItem) = atVehicleinfo(i).SplitItem(nItem)
                End If
                atVehicleInfo(i).SplitItem(nItem) = FormatMoney(atVehicleInfo(i).SplitItem(nItem))
'                End If

                atVehicleInfo(i).SettlePrice = atVehicleInfo(i).SettlePrice + GetSettlePrice(rsProtocol, atVehicleInfo(i).SplitItem(nItem))
                atVehicleInfo(i).SettleStationPrice = atVehicleInfo(i).SettleStationPrice + GetSettleStationPrice(rsProtocol, atVehicleInfo(i).SplitItem(nItem))
                rsProtocol.MoveNext
            Next j
            atVehicleInfo(i).SettlePrice = FormatMoney(atVehicleInfo(i).SettlePrice)  ' Round(atVehicleinfo(i).SettlePrice + 0.01, 1)
            atVehicleInfo(i).SettleStationPrice = FormatMoney(atVehicleInfo(i).SettleStationPrice)  ' Round(atVehicleinfo(i).SettleStationPrice + 0.01, 1)
            
            
            dbSettleOtherCompanyPrice = dbSettleOtherCompanyPrice + atVehicleInfo(i).SettlePrice
            dbSettleStationPrice = dbSettleStationPrice + atVehicleInfo(i).SettleStationPrice
        Next i
'        ******************
        
        '赋车次结算结果
        For i = 1 To ArrayLength(atBusInfo)
            
            rsProtocol.MoveFirst
            For j = 1 To nVailableNumber 'rsProtocol.RecordCount

                szitem = FormatDbValue(rsProtocol!split_item_id)
                nItem = Val(szitem)
                
                
                
                '将手工补票的数据进行累加
                If nItem = m_nCarriageItemPosition Then
                    '运费 拆出款加到运费的项目处
                    atBusInfo(i).SplitItem(nItem) = FormatMoney(atBusInfo(i).SplitItem(nItem) + dbExtraSettleOutPrice)
                    
                ElseIf nItem = m_nServiceItemPosition Then
                    '劳务费 劳务费加到劳务费的对应项目处
                    atBusInfo(i).SplitItem(nItem) = FormatMoney(atBusInfo(i).SplitItem(nItem) + dbExtraServicePrice)
                End If
                
                
                
                
                
                
                
                If m_nFixFeeItemCount > 0 And Not (rsFixFee Is Nothing) Then
                    If rsFixFee.RecordCount > 0 Then
                    '如果存在固定费用项
                        For k = 0 To m_nFixFeeItemCount - 1
                            If nItem = m_aszFixFeeItem(k) Then
                                atBusInfo(i).SplitItem(nItem) = atBusInfo(i).SplitItem(nItem) + FormatDbValue(rsFixFee.Fields("split_item_" & nItem))
                                Exit For
                            End If
                        Next k
                    End If
                End If
'                dbSplitlimitCharge = GetlimitCharge(FormatDbValue(rsProtocol!protocol_id), szItem, dbSplitUpCharge)
'                '如果上限不为0,则   该项值=上限值-该项值
'                If dbSplitUpCharge <> 0 Then
'                    atBusInfo(i).SplitItem(nItem) = FormatMoney(dbSplitUpCharge - atBusInfo(i).SplitItem(nItem))
'                End If
'
'
                
                dbSplitlimitCharge = GetlimitCharge(FormatDbValue(rsProtocol!protocol_id), szitem, dbSplitUpCharge)
                '如果上限不为0,则   该项值=上限值-该项值
                If dbSplitUpCharge <> 0 Then
                    
                    '针对椒江的情况进行修改
                    atBusInfo(i).SplitItem(nItem) = FormatMoney(atBusInfo(i).SplitItem(nItem) - dbSplitUpCharge)
                    If atBusInfo(i).SplitItem(nItem) < 0 Then
                        atBusInfo(i).SplitItem(nItem) = 0
                    End If
                End If
                
                
                
                
                '当下限值大于该项值时,该值为下限值
                If dbSplitlimitCharge > atBusInfo(i).SplitItem(nItem) Then
                    atBusInfo(i).SplitItem(nItem) = dbSplitlimitCharge
                Else
'                    atBusinfo(i).SplitItem(nItem) = atBusinfo(i).SplitItem(nItem)
                End If
                atBusInfo(i).SplitItem(nItem) = FormatMoney(atBusInfo(i).SplitItem(nItem))
'                End If
                atBusInfo(i).SettlePrice = atBusInfo(i).SettlePrice + GetSettlePrice(rsProtocol, atBusInfo(i).SplitItem(nItem))
                atBusInfo(i).SettleStationPrice = atBusInfo(i).SettleStationPrice + GetSettleStationPrice(rsProtocol, atBusInfo(i).SplitItem(nItem))
                rsProtocol.MoveNext
            Next j
            atBusInfo(i).SettlePrice = FormatMoney(atBusInfo(i).SettlePrice)  ' Round(atBusinfo(i).SettlePrice + 0.01, 1)
            atBusInfo(i).SettleStationPrice = FormatMoney(atBusInfo(i).SettleStationPrice)  ' Round(atBusinfo(i).SettleStationPrice + 0.01, 1)
            
            
            dbSettleOtherCompanyPrice = dbSettleOtherCompanyPrice + atBusInfo(i).SettlePrice
            dbSettleStationPrice = dbSettleStationPrice + atBusInfo(i).SettleStationPrice
        Next i
        dbSettleOtherCompanyPrice = dbSettleOtherCompanyPrice ' Round(dbSettleOtherCompanyPrice + 0.01, 1)
        dbSettleStationPrice = dbSettleStationPrice ' Round(dbSettleStationPrice + 0.01, 1)
        
        .SettleOtherCompanyPrice = dbSettleOtherCompanyPrice '结算给对方公司金额
        .SettleStationPrice = dbSettleStationPrice '结给站方金额

        If prsPassageNumber.RecordCount > 0 Then
            prsPassageNumber.MoveFirst
            If .SettleObject = CS_SettleByBus Then
                szSql = " SELECT route_id FROM bus_info WHERE bus_id = " & TransFieldValueToString(.ObjectID)
                Set rsTemp = oDB.Execute(szSql)
                If rsTemp.RecordCount > 0 Then
                    .RouteID = FormatDbValue(rsTemp!route_id)
                Else
                    .RouteID = FormatDbValue(prsPassageNumber!route_id)
                End If
            Else
                szSql = " SELECT route_id FROM bus_check_lst WHERE vehicle_id = " & TransFieldValueToString(.ObjectID) _
                    & " AND bus_date <" & TransFieldValueToString(pdtEndDate) _
                    & " AND bus_date >=" & TransFieldValueToString(pdtStartDate)
                Set rsTemp = oDB.Execute(szSql)
                If rsTemp.RecordCount > 0 Then
                    .RouteID = FormatDbValue(rsTemp!route_id)
                Else
                    .RouteID = FormatDbValue(prsPassageNumber!route_id)
                End If
            End If
            .TransportCompanyID = FormatDbValue(prsPassageNumber!transport_company_id)
            .TransportCompanyName = FormatDbValue(prsPassageNumber!transport_company_short_name)
            szSql = " SELECT route_name FROM route_info WHERE route_id = " & TransFieldValueToString(.RouteID)
            Set rsTemp = oDB.Execute(szSql)
            If rsTemp.RecordCount > 0 Then
                .RouteName = FormatDbValue(rsTemp!route_name)
            End If
        End If
    End With
    
    tResultInfo.CompanyInfo = atCompanyInfo
    tResultInfo.VehicleInfo = atVehicleInfo
    tResultInfo.BusInfo = atBusInfo
    PreviewSplitCheckSheetEx = tResultInfo
    
    Exit Function
PreviewSplitCheckSheetErr:
    err.Raise err.Number, err.Source, err.Description
End Function



'##ModelId=3E81F4790196
Public Sub SplitCheckSheet(pszSettleSheetID As String, pszCompanyID As String, pszVehicleID As String, paszCheckSheetID() As String _
        , ptResultInfo As TSplitResult, pdyStartDate As Date, pdyEndDate As Date, prsPassageNumber As Recordset, Optional pszAdditionPrice As Integer, Optional pszBusID As String, Optional pbIsToday As Boolean = False)
        ', pszLuggageIDs As String, pdbTotalLuggageBaseCarriage As Double         , pdbTotalLuggageSettlePrice As Double)
        
        
        
    On Error GoTo SplitCheckSheetErr
    
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    
    Dim szWhere As String
    Dim nSheetCount As Integer
    Dim i As Integer
    Dim j As Integer
    
    Dim bSplitByCompany As Boolean
    Dim bSplitByVehicle As Boolean
    Dim dyNow As String
    Dim oSysParam As New SystemParam
    Dim bAllowSplitBySomeTimes As Boolean '是否允许一张路单结算多次.

    '验证用户是否合法
    '验证对象是否合法
    '验证用户是否有结算的权限
    
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_SplitCheckSheet '验证用户是否有结算的权限
    
    oDB.ConnectionString = GetConnectionStr()
    
    If pszCompanyID = "" And pszVehicleID = "" And pszBusID = "" Then
        '出传入参数错
        RaiseError ERR_PreviewSplitInputParam
    ElseIf pszCompanyID <> "" Then
        bSplitByCompany = True
    ElseIf pszVehicleID <> "" Then
        bSplitByVehicle = True
    Else: bSplitByCompany = False And bSplitByVehicle = False
    End If
    
    oSysParam.Init m_oActiveUser
    bAllowSplitBySomeTimes = oSysParam.AllowSplitBySomeTimes '是否允许一张路单结算多次
    
    nSheetCount = ArrayLength(paszCheckSheetID)
    If nSheetCount = 0 Then RaiseError ERR_PreviewSplitInputParam '出传入参数错
    
    szWhere = " ( "
    For i = 1 To nSheetCount - 1
            szWhere = szWhere & "'" & paszCheckSheetID(i) & "',"
    Next i
    szWhere = szWhere & "'" & paszCheckSheetID(i) & "')"
    
    
    '判断每一张路单是否均存在，如不存在则出"路单不存在"的错。如无效或经过结算。则出"路单无效或路单已经结算"的错。
    szSql = "SELECT * FROM " & cszTableCheckSheet & "  WHERE check_sheet_id IN " & szWhere
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount <> nSheetCount Then
        '如果记录数不为传入的路单数,则出错
        RaiseError ERR_SomeSheetNotExist
    Else
'        '判断所有的路单是不是同一公司或同一车辆
'
'        For i = 1 To rsTemp.RecordCount
'            If bSplitByCompany Then
'                If FormatDbValue(rsTemp!transport_company_id) <> pszCompanyID Then
'                    '出错
'                    RaiseError ERR_SplitSheetMustSameCompany
'                End If
'            Else
'                If UCase(FormatDbValue(rsTemp!vehicle_id)) <> UCase(pszVehicleID) Then
'                    '出错
'                    RaiseError ERR_SplitSheetMustSameVehicle
'                End If
'            End If
'            rsTemp.MoveNext
'        Next i
    End If
    
'    szSql = "SELECT * FROM " & cszTableCheckSheet & " WHERE (settlement_status = " & TransFieldValueToString(CS_CheckSheetSettled) _
'        & " OR valid_mark = " & TransFieldValueToString(CS_CheckSheetInvalid) & " )" _
'        & " AND check_sheet_id IN " & szWhere
'
'    Set rsTemp = odb.Execute(szSql)
'    If rsTemp.RecordCount <> 0 Then
'        '如果存在已废或已结算的路单,则出错
'        RaiseError ERR_SomeSheetHasCancelORSettle
'    End If
    
    szSql = " SELECT * FROM fin_settle_sheet_lst WHERE settlement_sheet_id = " & TransFieldValueToString(pszSettleSheetID)
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount <> 0 Then
        RaiseError ERR_SettleSheetHasExist
    End If
    
    
'    Dim szObjectName As String
    
'    If ptResultInfo.SettleSheetInfo.SettleObject = CS_SettleByTransportCompany Then
'        szSql = " SELECT transport_company_short_name object_name FROM " & cszTableCompany & " WHERE transport_company_id = '" & ptResultInfo.SettleSheetInfo.ObjectID & "'"
'    ElseIf ptResultInfo.SettleSheetInfo.SettleObject = CS_SettleByVehicle Then
'        szSql = "SELECT license_tag_no object_name FROM " & cszTableVehicle & " WHERE vehicle_id = '" & ptResultInfo.SettleSheetInfo.ObjectID & "'"
'    End If
'
'    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount > 0 Then szObjectName = FormatDbValue(rsTemp!object_name)
'
    
    oDB.BeginTrans

        dyNow = ToDBDateTime(SelfNowDateTime())
        '开始事务
        '根据公司结算明细及车辆结算明细，累加出总人数、总应结    票款、总应结给车站票款
        '汇总出所有这些路单的总的售票票款、总路单数
        '写入"路单结算单表"，拆帐人写入用户代码，拆帐日期写当前日期，起始日期与结束日期用传入的参数写入
        
        szSql = " INSERT fin_settle_sheet_lst ( " _
            & " settlement_sheet_id , " _
            & " sheet_quantity ," _
            & " total_ticket_price , " _
            & " total_quantity , " _
            & " settle_price , " _
            & " settle_station_price , " _
            & " object_id , " _
            & " object_name , " _
            & " settle_object , " _
            & " transport_company_id , " _
            & " transport_company_name , " _
            & " route_id , " _
            & " route_name , " _
            & " settler , " _
            & " checker , " _
            & " settle_date , " _
            & " start_date , " _
            & " end_date , " _
            & " unit_name , " _
            & " status , " _
            & " is_print , "
            
        szSql = szSql & " luggage_settle_id , " _
            & " luggage_base_carriage , " _
            & " luggage_settle_price , " _
            & " luggage_protocol_name , "
            
            
        szSql = szSql & " annotation) VALUES (" _
            & TransFieldValueToString(pszSettleSheetID) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.CheckSheetCount) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.TotalTicketPrice) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.TotalQuantity) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.SettleOtherCompanyPrice) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.SettleStationPrice) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.ObjectID) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.ObjectName) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.SettleObject) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.TransportCompanyID) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.TransportCompanyName) & ","
            
        szSql = szSql & TransFieldValueToString(ptResultInfo.SettleSheetInfo.RouteID) & " , " _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.RouteName) & " , " _
            & TransFieldValueToString(m_oActiveUser.UserName) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.Checker) & "," _
            & TransFieldValueToString(dyNow) & "," _
            & TransFieldValueToString(pdyStartDate) & "," _
            & TransFieldValueToString(pdyEndDate) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.UnitName) & "," _
            & TransFieldValueToString(CS_SettleSheetValid) & "," _
            & TransFieldValueToString(False) & " ," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.LuggageSettleIDs) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.LuggageTotalBaseCarriage) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.LuggageTotalSettlePrice) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.LuggageProtocolName) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.Annotation) & ")"
        
        oDB.Execute szSql
        
        '如果结算对象为车辆,且存在固定费用
        If m_nFixFeeItemCount > 0 And ptResultInfo.SettleSheetInfo.SettleObject = CS_SettleByVehicle Then
            
            '修改<结算单车辆固定费用清单>,将在时间段内的,未扣的状态改为已扣,结算单号改为对应的结算单号
            szSql = " UPDATE fin_vehicle_fix_fee_lst SET is_dec = " & TransFieldValueToString(True) _
                & " , settlement_sheet_id =" & TransFieldValueToString(pszSettleSheetID) _
                & " WHERE vehicle_id = " & TransFieldValueToString(ptResultInfo.SettleSheetInfo.ObjectID) _
                & " AND bus_date >= " & TransFieldValueToString(pdyStartDate) _
                & " AND bus_date < " & TransFieldValueToString(pdyEndDate) _
                & " AND is_dec = " & TransFieldValueToString(False)
                
            oDB.Execute szSql
            
'            '**************************************************
'            '将结算单的状态,应设置为已结清的状态
'            '条件可能不对,因为现在是所有的将小于本次结算起始时间的结算单的未结清状态改为已结清 " AND end_date < " & TransFieldValueToString(pdyStartDate)
'            '**************************************************
'            szSql = " UPDATE fin_settle_sheet_lst SET status = " & TransFieldValueToString(CS_SettleSheetNegativeHasPayed) _
'                & " WHERE settle_object = " & TransFieldValueToString(CS_SettleByVehicle) _
'                & " AND object_id = " & TransFieldValueToString(ptResultInfo.SettleSheetInfo.ObjectID) _
'                & " AND settle_price -settle_station_price < 0 AND status = " & TransFieldValueToString(CS_SettleSheetNegativeNotPay) _
'                & " AND end_date < " & TransFieldValueToString(pdyStartDate)
'
'            oDB.Execute (szSql)
        ElseIf m_nFixFeeItemCount > 0 And ptResultInfo.SettleSheetInfo.SettleObject = CS_SettleByBus Then
        '如果结算对象为车次，且存在固定费用
            '修改<结算单车次固定费用清单>,将在时间段内的,未扣的状态改为已扣,结算单号改为对应的结算单号
            szSql = " UPDATE fin_bus_fix_fee_lst SET is_dec = " & TransFieldValueToString(True) _
                & " , settlement_sheet_id =" & TransFieldValueToString(pszSettleSheetID) _
                & " WHERE bus_id = " & TransFieldValueToString(ptResultInfo.SettleSheetInfo.ObjectID) _
                & " AND bus_date >= " & TransFieldValueToString(pdyStartDate) _
                & " AND bus_date < " & TransFieldValueToString(pdyEndDate) _
                & " AND is_dec = " & TransFieldValueToString(False)
                
            oDB.Execute szSql
        End If
        
        
        
        '是否允许一张路单结算多次?
        If Not bAllowSplitBySomeTimes Then
            
            
            '不允许
            
            '修改每一张路单的结算单号、拆算协议、结算时间、结算标记。            '拆算协议这个字段不写入            '因为一个协议不一定是只拆给一下公司的，或一个车辆的
            szSql = " UPDATE " & cszTableCheckSheet & " SET settlement_sheet_id = " & TransFieldValueToString(pszSettleSheetID) _
                & " , settlement_time = " & TransFieldValueToString(dyNow) _
                & " , settlement_status = " & TransFieldValueToString(CS_CheckSheetSettled) _
                & " WHERE check_sheet_id IN " & szWhere
                
            oDB.Execute szSql
            
            
            
            '使用实拆数累加的方式,因为考虑,如果未拆完,还需要再拆.  因为有小票根拉下的情况的.
            prsPassageNumber.MoveFirst
            For i = 1 To prsPassageNumber.RecordCount
                szSql = " UPDATE " & IIf(Not pbIsToday, cszTableSheetStation, cszTableSheetStationTemp) & " SET fact_quantity =fact_quantity + " & FormatDbValue(prsPassageNumber!fact_quantity) _
                    & " WHERE check_sheet_id = " & TransFieldValueToString(prsPassageNumber!check_sheet_id) _
                    & " AND sell_station_id = " & TransFieldValueToString(prsPassageNumber!sell_station_id) _
                    & " AND station_id = " & TransFieldValueToString(prsPassageNumber!station_id) _
                    & " AND price_identify = " & TransFieldValueToString(prsPassageNumber!price_identify) _
                    & " AND ticket_type = " & TransFieldValueToString(prsPassageNumber!ticket_type) _
                    & " AND seat_type_id = " & TransFieldValueToString(prsPassageNumber!seat_type_id)
                oDB.Execute szSql
                
                '写入<拆帐结算单实际车次站点拆算表>   为了查询及作废时恢复用
                szSql = " INSERT fin_settle_bus_station_lst ( " _
                    & " settlement_sheet_id , " _
                    & " check_sheet_id , " _
                    & " sell_station_id , " _
                    & " station_id , " _
                    & " price_identify , " _
                    & " ticket_type , " _
                    & " seat_type_id , " _
                    & " status , " _
                    & " bus_date , " _
                    & " bus_id , " _
                    & " bus_serial_no , " _
                    & " station_serial , " _
                    & " station_name , " _
                    & " quantity , " _
                    & " mileage , " _
                    & " ticket_price , " _
                    & " base_carriage , " _
                    
                For j = 1 To cnTotalPriceItem - 1
                    szSql = szSql & " price_item_" & j & " , "
                Next j
                szSql = szSql & " price_item_" & j & " ) "
                
                
                szSql = szSql & " VALUES (" _
                    & TransFieldValueToString(pszSettleSheetID) & " , " _
                    & TransFieldValueToString(FormatDbValue((prsPassageNumber!check_sheet_id))) & " , " _
                    & TransFieldValueToString(FormatDbValue((prsPassageNumber!sell_station_id))) & " , " _
                    & TransFieldValueToString(FormatDbValue((prsPassageNumber!station_id))) & " , " _
                    & TransFieldValueToString(FormatDbValue((prsPassageNumber!price_identify))) & " , " _
                    & TransFieldValueToString(FormatDbValue((prsPassageNumber!ticket_type))) & " , " _
                    & TransFieldValueToString(FormatDbValue((prsPassageNumber!seat_type_id))) & " , " _
                    & TransFieldValueToString(FormatDbValue((prsPassageNumber!Status))) & " , " _
                    & TransFieldValueToString(FormatDbValue((prsPassageNumber!Bus_date))) & " , " _
                    & TransFieldValueToString(FormatDbValue((prsPassageNumber!bus_id))) & " , " _
                    & TransFieldValueToString(FormatDbValue((prsPassageNumber!bus_serial_no))) & " , " _
                    & TransFieldValueToString(FormatDbValue((prsPassageNumber!station_serial))) & " , " _
                    & TransFieldValueToString(FormatDbValue((prsPassageNumber!station_name))) & " , " _
                    & TransFieldValueToString(FormatDbValue((prsPassageNumber!fact_quantity))) & " , " _
                    & TransFieldValueToString(FormatDbValue((prsPassageNumber!Mileage))) & " , " _
                    & TransFieldValueToString(FormatDbValue((prsPassageNumber!ticket_price))) & " , " _
                    & TransFieldValueToString(FormatDbValue((prsPassageNumber!base_carriage))) & " , "
                
                For j = 1 To cnTotalPriceItem - 1
                    szSql = szSql & TransFieldValueToString(FormatDbValue((prsPassageNumber.Fields("price_item_" & j)))) & " , "
                Next j
                szSql = szSql & TransFieldValueToString(FormatDbValue((prsPassageNumber.Fields("price_item_" & j)))) & " ) "
                oDB.Execute szSql
                prsPassageNumber.MoveNext
            Next i
            
            
            
            '张数从实拆中统计
            '现在是不考虑一张路单拆多次的情况,故可以直接从fin_sheet_station_lst表中汇总得到.
            ',如果允许拆多次,就不能这样写了
            
            szSql = "INSERT fin_settle_station_lst(settlement_sheet_id,sell_station_id,route_id,route_name " _
                & "     , station_id,ticket_type_id,quantity,bus_id,vehicle_type_code ,vehicle_id,transport_company_id) " _
                & " SELECT " & TransFieldValueToString(pszSettleSheetID) & ",(f.sell_station_id) sell_station_id,(s.route_id) route_id, MAX(r.route_name) route_name " _
                & "     , (f.station_id) station_id,(f.ticket_type) ticket_type,SUM(f.fact_quantity) quantity,MAX(s.bus_id) bus_id" _
                & "     , MAX(s.vehicle_type_code) vehicle_type_code," & TransFieldValueToString(pszVehicleID) & " vehicle_id" _
                & "     , " & TransFieldValueToString(pszCompanyID) & " transport_company_id " _
                & " FROM " & IIf(Not pbIsToday, cszTableSheetStation, cszTableSheetStationTemp) & " f, " & cszTableCheckSheet & " s , " & cszTableRouteInfo & " r " _
                & " WHERE s.route_id = r.route_id AND s.check_sheet_id=f.check_sheet_id AND s.check_sheet_id IN" & szWhere _
                & " GROUP BY f.sell_station_id , s.route_id , f.station_id , f.ticket_type "
            
            
            oDB.Execute szSql
        Else
            Dim szCheckSheetID As String '临时保存路单,只为比较时用
            Dim bAllowSplitAboveFactQuantity As Boolean '是否允许结算人数大于实际路单站点人数
            Dim bIsSettled As Boolean
            Dim bFind As Boolean '是否找到该条件的人数
            Dim aszSettledSheet() As String
            Dim aszNotSettledSheet() As String
            Dim k As Integer
            Dim szSettledSheet As String
            Dim szNotSettledSheet As String
            
            
            bAllowSplitAboveFactQuantity = oSysParam.AllowSplitAboveFactQuantity
            bIsSettled = True
            
            '得到要结算的站点的所有的总人数,已结人数
            szSql = " SELECT s.*, t.station_id sell_station_id2  FROM " & IIf(Not pbIsToday, cszTableSheetStation, cszTableSheetStationTemp) & " s , sell_station_info t  " _
                & " WHERE s.sell_station_id = t.sell_station_id AND check_sheet_id IN " & szWhere _
                & " ORDER BY check_sheet_id "
            Set rsTemp = oDB.Execute(szSql)
            
            
'            szCheckSheetID = FormatDbValue(rsTemp!check_sheet_id)
            bIsSettled = True
            k = 0
            For i = 1 To rsTemp.RecordCount
                prsPassageNumber.MoveFirst
                bFind = False
                bIsSettled = True
                For j = 1 To prsPassageNumber.RecordCount
                
                
                    If FormatDbValue(rsTemp!check_sheet_id) = FormatDbValue(prsPassageNumber!check_sheet_id) _
                        And FormatDbValue(prsPassageNumber!sell_station_id) = FormatDbValue(rsTemp!sell_station_id2) _
                        And FormatDbValue(prsPassageNumber!station_id) = FormatDbValue(rsTemp!station_id) _
                        And FormatDbValue(prsPassageNumber!price_identify) = FormatDbValue(rsTemp!price_identify) _
                        And FormatDbValue(prsPassageNumber!ticket_type) = FormatDbValue(rsTemp!ticket_type) _
                        And FormatDbValue(prsPassageNumber!seat_type_id) = FormatDbValue(rsTemp!seat_type_id) Then
                            
                            
                        
                        '当条件相等时
                        bFind = True
                        
                        '需判断是否总人数大于已结人数
            
                        If FormatDbValue(prsPassageNumber!fact_quantity) < FormatDbValue(rsTemp!Quantity) - FormatDbValue(rsTemp!fact_quantity) Then
                            '实结人数与剩余的人数不等.
                            
                            '如果还有剩余人数,则设为正在结算 ,已将<>改为了<
                            If Not bAllowSplitAboveFactQuantity And FormatDbValue(prsPassageNumber!fact_quantity) > FormatDbValue(rsTemp!Quantity) - FormatDbValue(rsTemp!fact_quantity) Then
                                RaiseError ERR_NotAllowSplitAboveFactQuantity '大于则,出错
                            End If
                            bIsSettled = False
                        
                        End If
                        Exit For
                    Else
                        '将该路单的状态保存下来
                        
                        
                    End If
                    prsPassageNumber.MoveNext
                Next j
                If Not bFind Then
                    '未找到
                    If FormatDbValue(rsTemp!Quantity) > FormatDbValue(rsTemp!fact_quantity) Then
                        '但总人数与已结人数不相同
                        '如果还有剩余人数,则设为正在结算 ,已将<>改为了>
                        bIsSettled = False
                    End If
                End If
                
                If bIsSettled Then
                    k = k + 1
                    ReDim Preserve aszSettledSheet(1 To k)
                    aszSettledSheet(k) = FormatDbValue(rsTemp!check_sheet_id)
                Else
                    k = k + 1
                    ReDim Preserve aszNotSettledSheet(1 To k)
                    aszNotSettledSheet(k) = FormatDbValue(rsTemp!check_sheet_id)
                End If
                rsTemp.MoveNext
            Next i
            '判断是否该路单的所有的人数均结算完毕.
            '保存已结算完毕的路单
            If ArrayLength(aszSettledSheet) > 0 Then
                szSettledSheet = " ( "
                For i = 1 To ArrayLength(aszSettledSheet) - 1
                        szSettledSheet = szSettledSheet & "'" & aszSettledSheet(i) & "',"
                Next i
                szSettledSheet = szSettledSheet & "'" & aszSettledSheet(i) & "')"
                
                '修改每一张路单的结算单号、拆算协议、结算时间、结算标记。            '拆算协议这个字段不写入            '因为一个协议不一定是只拆给一下公司的，或一个车辆的
                szSql = " UPDATE " & cszTableCheckSheet & " SET settlement_sheet_id = " & TransFieldValueToString(pszSettleSheetID) _
                    & " , settlement_time = " & TransFieldValueToString(dyNow) _
                    & " , settlement_status = " & TransFieldValueToString(CS_CheckSheetSettled) _
                    & " WHERE check_sheet_id IN " & szSettledSheet
                    
                oDB.Execute szSql
                           
            End If
            '保存未结算完毕的路单
            If ArrayLength(aszNotSettledSheet) > 0 Then
                szNotSettledSheet = " ( "
                For i = 1 To ArrayLength(aszNotSettledSheet) - 1
                        szNotSettledSheet = szNotSettledSheet & "'" & aszNotSettledSheet(i) & "',"
                Next i
                szNotSettledSheet = szNotSettledSheet & "'" & aszNotSettledSheet(i) & "')"
            
                '修改每一张路单的结算单号、拆算协议、结算时间、结算标记。            '拆算协议这个字段不写入            '因为一个协议不一定是只拆给一下公司的，或一个车辆的
                szSql = " UPDATE " & cszTableCheckSheet & " SET settlement_sheet_id = " & TransFieldValueToString(pszSettleSheetID) _
                    & " , settlement_time = " & TransFieldValueToString(dyNow) _
                    & " , settlement_status = " & TransFieldValueToString(CS_CheckSheetSettling) _
                    & " WHERE check_sheet_id IN " & szNotSettledSheet
                    
                oDB.Execute szSql
                           
            End If
            
            
            
            '保存实拆人数时,应按实拆人数来算,不能按sheet_lst表的fact_quantity字段汇总得到.
            
            '使用实拆数累加的方式,因为考虑,如果未拆完,还需要再拆.  因为有小票根拉下的情况的.
            prsPassageNumber.MoveFirst
            For i = 1 To prsPassageNumber.RecordCount
                szSql = " UPDATE " & IIf(Not pbIsToday, cszTableSheetStation, cszTableSheetStationTemp) & " SET fact_quantity =fact_quantity + " & FormatDbValue(prsPassageNumber!fact_quantity) _
                    & "  FROM " & IIf(Not pbIsToday, cszTableSheetStation, cszTableSheetStationTemp) & " s, sell_station_info t WHERE check_sheet_id = " & TransFieldValueToString(prsPassageNumber!check_sheet_id) _
                    & " AND s.sell_station_id = t.sell_station_id AND t.station_id = " & TransFieldValueToString(prsPassageNumber!sell_station_id) _
                    & " AND s.station_id = " & TransFieldValueToString(prsPassageNumber!station_id) _
                    & " AND price_identify = " & TransFieldValueToString(prsPassageNumber!price_identify) _
                    & " AND ticket_type = " & TransFieldValueToString(prsPassageNumber!ticket_type) _
                    & " AND seat_type_id = " & TransFieldValueToString(prsPassageNumber!seat_type_id)
                    
                oDB.Execute szSql
                
                '写入<拆帐结算单实际车次站点拆算表>   为了查询及作废时恢复用
                szSql = " INSERT fin_settle_bus_station_lst ( " _
                    & " settlement_sheet_id , " _
                    & " check_sheet_id , " _
                    & " sell_station_id , " _
                    & " station_id , " _
                    & " price_identify , " _
                    & " ticket_type , " _
                    & " seat_type_id , " _
                    & " status , " _
                    & " bus_date , " _
                    & " bus_id , " _
                    & " bus_serial_no , " _
                    & " station_serial , " _
                    & " station_name , " _
                    & " quantity , " _
                    & " mileage , " _
                    & " ticket_price , " _
                    & " base_carriage , " _
                    
                For j = 1 To cnTotalPriceItem - 1
                    szSql = szSql & " price_item_" & j & " , "
                Next j
                szSql = szSql & " price_item_" & j & " ) "
                
                
                szSql = szSql & " VALUES (" _
                    & TransFieldValueToString(pszSettleSheetID) & " , " _
                    & TransFieldValueToString(FormatDbValue(prsPassageNumber!check_sheet_id)) & " , " _
                    & TransFieldValueToString(FormatDbValue(prsPassageNumber!sell_station_id)) & " , " _
                    & TransFieldValueToString(FormatDbValue(prsPassageNumber!station_id)) & " , " _
                    & TransFieldValueToString(FormatDbValue(prsPassageNumber!price_identify)) & " , " _
                    & TransFieldValueToString(FormatDbValue(prsPassageNumber!ticket_type)) & " , " _
                    & TransFieldValueToString(FormatDbValue(prsPassageNumber!seat_type_id)) & " , " _
                    & TransFieldValueToString(FormatDbValue(prsPassageNumber!Status)) & " , " _
                    & TransFieldValueToString(FormatDbValue(prsPassageNumber!Bus_date)) & " , " _
                    & TransFieldValueToString(FormatDbValue(prsPassageNumber!bus_id)) & " , " _
                    & TransFieldValueToString(FormatDbValue(prsPassageNumber!bus_serial_no)) & " , " _
                    & TransFieldValueToString(FormatDbValue(prsPassageNumber!station_serial)) & " , " _
                    & TransFieldValueToString(FormatDbValue(prsPassageNumber!station_name)) & " , " _
                    & TransFieldValueToString(FormatDbValue(prsPassageNumber!fact_quantity)) & " , " _
                    & TransFieldValueToString(FormatDbValue(prsPassageNumber!Mileage)) & " , " _
                    & TransFieldValueToString(FormatDbValue(prsPassageNumber!ticket_price)) & " , " _
                    & TransFieldValueToString(FormatDbValue(prsPassageNumber!base_carriage)) & " , "
                
                For j = 1 To cnTotalPriceItem - 1
                    szSql = szSql & TransFieldValueToString(FormatDbValue(prsPassageNumber.Fields("price_item_" & j))) & " , "
                Next j
                szSql = szSql & TransFieldValueToString(FormatDbValue(prsPassageNumber.Fields("price_item_" & j))) & " ) "
                oDB.Execute szSql
                
                prsPassageNumber.MoveNext
            
            Next i
            
            '人数需要汇总
                
    '        将汇总的站点人数信息写入 "拆帐结算单站点表"
            For i = 1 To ArrayLength(ptResultInfo.SheetStationInfo)
                szSql = "INSERT fin_settle_station_lst( " _
                    & " settlement_sheet_id, " _
                    & " sell_station_id," _
                    & " route_id," _
                    & " route_name , " _
                    & " station_id, " _
                    & " ticket_type_id," _
                    & " quantity, " _
                    & " bus_id, " _
                    & " vehicle_type_code ," _
                    & " vehicle_id," _
                    & " transport_company_id) "
                szSql = szSql & " VALUES ( " _
                    & TransFieldValueToString(pszSettleSheetID) & " , " _
                    & TransFieldValueToString(ptResultInfo.SheetStationInfo(i).SellSationID) & " , " _
                    & TransFieldValueToString(ptResultInfo.SheetStationInfo(i).RouteID) & " , " _
                    & TransFieldValueToString(ptResultInfo.SheetStationInfo(i).RouteName) & " , " _
                    & TransFieldValueToString(ptResultInfo.SheetStationInfo(i).StationID) & " , " _
                    & TransFieldValueToString(ptResultInfo.SheetStationInfo(i).TicketType) & " , " _
                    & TransFieldValueToString(ptResultInfo.SheetStationInfo(i).Quantity) & " , " _
                    & TransFieldValueToString("") & " , " _
                    & TransFieldValueToString(ptResultInfo.SheetStationInfo(i).VehicleTypeCode) & " , " _
                    & TransFieldValueToString(IIf(bSplitByCompany, "", pszVehicleID)) & " , " _
                    & TransFieldValueToString(IIf(bSplitByCompany, pszCompanyID, "")) & " ) "
                    
                oDB.Execute szSql
            Next i

            
'
''
''            '张数从实拆中统计
''            '现在是不考虑一张路单拆多次的情况,故可以直接从fin_sheet_station_lst表中汇总得到.
''            ',如果允许拆多次,就不能这样写了
''
'            szSql = "INSERT fin_settle_station_lst(settlement_sheet_id,sell_station_id,route_id,route_name " _
'                & "     , station_id,ticket_type_id,quantity,bus_id,vehicle_type_code ,vehicle_id,transport_company_id) " _
'                & " SELECT " & TransFieldValueToString(pszSettleSheetID) & ",(f.sell_station_id) sell_station_id,(s.route_id) route_id, MAX(r.route_name) route_name " _
'                & "     , (f.station_id) station_id,(f.ticket_type) ticket_type,SUM(f.fact_quantity) quantity,MAX(s.bus_id) bus_id" _
'                & "     , MAX(s.vehicle_type_code) vehicle_type_code," & TransFieldValueToString(pszVehicleID) & " vehicle_id" _
'                & "     , " & TransFieldValueToString(pszCompanyID) & " transport_company_id " _
'                & " FROM " & IIf(Not pbIsToday, cszTableSheetStation, cszTableSheetStationTemp) & " f, " & cszTableCheckSheet & " s , " & cszTableRouteInfo & " r " _
'                & " WHERE s.route_id = r.route_id AND s.check_sheet_id=f.check_sheet_id AND s.check_sheet_id IN" & szWhere _
'                & " GROUP BY f.sell_station_id , s.route_id , f.station_id , f.ticket_type "
'            oDb.Execute szSql
    
        End If
        '写入"结算单公司拆算明细表"
        For i = 1 To ArrayLength(ptResultInfo.CompanyInfo)
            szSql = " INSERT fin_settle_company_price_lst ( " _
                & " settlement_sheet_id , " _
                & " transport_company_id , " _
                & " protocol_id, " _
                & " settle_price, " _
                & " settle_station_price, " _
                & " mileage , " _
                & " passenger_number, "
            
            For j = 1 To 19
                szSql = szSql & "split_item_" & CInt(j) & " , "
            Next j
            szSql = szSql & "split_item_20 ) "
            szSql = szSql & " VALUES ( " _
                & TransFieldValueToString(pszSettleSheetID) & " , " _
                & TransFieldValueToString(ptResultInfo.CompanyInfo(i).CompanyID) & " , " _
                & TransFieldValueToString(ptResultInfo.CompanyInfo(i).ProtocolID) & " , " _
                & TransFieldValueToString(ptResultInfo.CompanyInfo(i).SettlePrice) & " , " _
                & TransFieldValueToString(ptResultInfo.CompanyInfo(i).SettleStationPrice) & " , " _
                & TransFieldValueToString(ptResultInfo.CompanyInfo(i).Mileage) & "," _
                & TransFieldValueToString(ptResultInfo.CompanyInfo(i).PassengerNumber) & " , "
                
                
            For j = 1 To 19
                szSql = szSql & TransFieldValueToString(ptResultInfo.CompanyInfo(i).SplitItem(j)) & " , "
            Next j
            szSql = szSql & TransFieldValueToString(ptResultInfo.CompanyInfo(i).SplitItem(20)) & " ) "
            oDB.Execute szSql
            
        Next i
        
        '写入"结算单车辆拆算明细表"
        For i = 1 To ArrayLength(ptResultInfo.VehicleInfo)
            With ptResultInfo.VehicleInfo(i)
                szSql = " INSERT fin_settle_vehicle_lst ( " _
                    & " settlement_sheet_id , " _
                    & " vehicle_id , " _
                    & " protocol_id , " _
                    & " settle_price , " _
                    & " settle_station_price , " _
                    & " addition_price , " _
                    & " mileage , " _
                    & " passenger_number , "
                For j = 1 To 19
                    szSql = szSql & "split_item_" & CInt(j) & " , "
                Next j
                szSql = szSql & "split_item_20 ) "
                szSql = szSql & " VALUES ( " _
                    & TransFieldValueToString(pszSettleSheetID) & " , " _
                    & TransFieldValueToString(.VehicleID) & " , " _
                    & TransFieldValueToString(.ProtocolID) & " , " _
                    & TransFieldValueToString(.SettlePrice) & " , " _
                    & TransFieldValueToString(.SettleStationPrice) & " , " _
                    & pszAdditionPrice & " , " _
                    & TransFieldValueToString(.Mileage) & " , " _
                    & TransFieldValueToString(.PassengerNumber) & " , "
                    
                For j = 1 To 19
                    szSql = szSql & TransFieldValueToString(.SplitItem(j)) & " , "
                Next j
            
                szSql = szSql & TransFieldValueToString(.SplitItem(20)) & " ) "
                
                oDB.Execute szSql
            End With
        Next i
        
        '写入"结算单车次拆算明细表"
        For i = 1 To ArrayLength(ptResultInfo.BusInfo)
            With ptResultInfo.BusInfo(i)
                szSql = " INSERT fin_settle_bus_lst ( " _
                    & " settlement_sheet_id , " _
                    & " bus_id , " _
                    & " transport_company_id , " _
                    & " protocol_id , " _
                    & " settle_price , " _
                    & " settle_station_price , " _
                    & " addition_price , " _
                    & " passenger_number , " _
                    & " mileage , "
                For j = 1 To 19
                    szSql = szSql & "split_item_" & CInt(j) & " , "
                Next j
                szSql = szSql & "split_item_20 ) "
                szSql = szSql & " VALUES ( " _
                    & TransFieldValueToString(pszSettleSheetID) & " , " _
                    & TransFieldValueToString(.BusID) & " , " _
                    & TransFieldValueToString(pszCompanyID) & " , " _
                    & TransFieldValueToString(.ProtocolID) & " , " _
                    & TransFieldValueToString(.SettlePrice) & " , " _
                    & TransFieldValueToString(.SettleStationPrice) & " , " _
                    & pszAdditionPrice & " , " _
                    & TransFieldValueToString(.PassengerNumber) & " , " _
                    & TransFieldValueToString(.Mileage) & " , "
                    
                For j = 1 To 19
                    szSql = szSql & TransFieldValueToString(.SplitItem(j)) & " , "
                Next j
            
                szSql = szSql & TransFieldValueToString(.SplitItem(20)) & " ) "
                
                oDB.Execute szSql
            End With
        Next i
        
        '写入[结算单手工补票表]
        For i = 1 To ArrayLength(ptResultInfo.SettleExtraInfo)
            With ptResultInfo.SettleExtraInfo(i)
                szSql = " INSERT fin_settle_extra_lst ( " _
                    & " settlement_sheet_id , " _
                    & " passenger_number , " _
                    & " total_ticket_price , " _
                    & " ratio , " _
                    & " service_price , " _
                    & " settle_out_price ) VALUES ( " _
                    & TransFieldValueToString(pszSettleSheetID) & " , " _
                    & TransFieldValueToString(.PassengerNumber) & " , " _
                    & TransFieldValueToString(.TotalTicketPrice) & "," _
                    & TransFieldValueToString(.Ratio) & " ," _
                    & TransFieldValueToString(.ServicePrice) & "," _
                    & TransFieldValueToString(.SettleOutPrice) & ")"
                oDB.Execute szSql
            End With
            
        Next i
        
        
    '结束事务
    oDB.CommitTrans
    Dim szlog As String
    szlog = "结算单号为" & pszSettleSheetID & ","
    If bSplitByCompany Then
        szlog = szlog & "按公司结算,结算对象为" & pszCompanyID & ","
    ElseIf bSplitByVehicle Then
        szlog = szlog & "按车辆结算,结算对象为" & pszVehicleID & ","
    Else: szlog = szlog & "按车次结算,结算对象为" & pszBusID & ","
    End If
    szlog = szlog & "结算日期段为" & ToDBDate(pdyStartDate) & "-" & ToDBDate(pdyEndDate)
    '写入日志
    WriteOperateLog m_oActiveUser, RIGHT_SplitCheckSheet, szlog

    Exit Sub
SplitCheckSheetErr:
    oDB.RollbackTrans
    Call err.Raise(err.Number, "Split:SplitCheckSheet Method", err.Description)
End Sub

'##ModelId=3E81F47D0128
Public Sub SetVehicleProtocol(paszVelicleID() As String, pszProtocolID As String)
    On Error GoTo SetVehicleProtocolErr
    
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    Dim szWhere As String
    Dim i As Integer
    
    Dim nVehicleCount  As Integer
    
    '验证用户是否合法
    '验证对象是否合法
    '验证用户是否有结算的权限
    
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_SetVehicleProtocol  '验证用户是否有设置公司协议的权限
    
    oDB.ConnectionString = GetConnectionStr()
    
    
    
    
    
    nVehicleCount = ArrayLength(paszVelicleID)
    If nVehicleCount = 0 Then RaiseError ERR_NoVehicle '没有车辆要设置协议
    
    szWhere = " ( "
    For i = 1 To nVehicleCount - 1
            szWhere = szWhere & "'" & paszVelicleID(i) & "',"
    Next i
    szWhere = szWhere & "'" & paszVelicleID(i) & "')"
    
    If pszProtocolID <> "" Then
        '判断该协议是否存在，如果该协议不存在，则出错
        szSql = " SELECT * FROM fin_settle_protocol_info WHERE protocol_id = " & TransFieldValueToString(pszProtocolID)
        Set rsTemp = oDB.Execute(szSql)
        If rsTemp.RecordCount = 0 Then
            RaiseError ERR_ProtocolNotExist '协议不存在
        End If
    Else
        '允许传入协议为空
    End If
    '将这些车辆设置为该协议
    szSql = " UPDATE " & cszTableVehicle & " SET protocol_id = " & TransFieldValueToString(pszProtocolID) _
        & " WHERE vehicle_id IN " & szWhere
    oDB.Execute szSql
    
    szWhere = ""
    For i = 1 To nVehicleCount - 1
            szWhere = szWhere & paszVelicleID(i) & ","
    Next i
    szWhere = szWhere & paszVelicleID(i)
    '写入日志
    WriteOperateLog m_oActiveUser, RIGHT_SetCompanyProtocol, "设置车辆协议" & szWhere
    
    Exit Sub
SetVehicleProtocolErr:
    Call RaiseError(err.Number, "Split:SetVehicleProtocol Method")
End Sub

'##ModelId=3E81F481000F
Public Sub SetCompanyProtocol(paszCompanyID As String, pszRouteID As String, pszProtocolID As String, pszRouteName As String)
    On Error GoTo SetCompanyProtocolErr

    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    Dim szWhere As String
    Dim i As Integer
    
    Dim nCompanyCount  As Integer
    
    '验证用户是否合法
    '验证对象是否合法
    '验证用户是否有结算的权限
    
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_SetCompanyProtocol    '验证用户是否有设置车辆协议的权限
    
    
    oDB.ConnectionString = GetConnectionStr()
    
    szSql = "DELETE fin_compnay_protocol_lst WHERE transport_company_id = '" & paszCompanyID & "' AND route_id  = '" & pszRouteID & "'"
    oDB.Execute szSql
    szSql = "INSERT fin_compnay_protocol_lst  ( transport_company_id , route_id ,route_name ,protocol_id ) " _
        & " VALUES(" _
        & TransFieldValueToString(paszCompanyID) & "," _
        & TransFieldValueToString(pszRouteID) & "," _
        & TransFieldValueToString(pszRouteName) & "," _
        & TransFieldValueToString(pszProtocolID) & ") "
        
    oDB.Execute szSql
    
    
    Exit Sub
SetCompanyProtocolErr:
    Call RaiseError(err.Number, "Split:SetCompanyProtocol Method")
End Sub

'##ModelId=3E8219D10109
Private Sub Class_Initialize()
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
'        mlClassDebugID = GetNextClassDebugID()
'        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'##ModelId=3E8219D10128
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
'    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

#If DebugMode Then
    '##ModelId=3E8219D1006D
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If



'**************************************************
'Member Code:S3
'Brief Description:测试对象是否有效
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then RaiseError ERR_SplitNotAvailable
End Sub

'**************************************************
'Member Code:S4
'Brief Description:测试对象是否有效
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub AssertObjIsValid()
    AssertActiveUserValid m_oActiveUser, ERR_Split
    AssertStatusAvailable
End Sub

Public Function GetFormulaItem() As String()
    '得到设置公式用的项
    
    
    '包括各拆算项及拆算运费、拆算通行费、人数、票价项等
    Dim aszTemp() As String
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    
    Dim i As Integer
    Dim nCount As Integer
    
    oDB.ConnectionString = GetConnectionStr
    szSql = " SELECT * FROM price_item_info "
    Set rsTemp = oDB.Execute(szSql)
    
    nCount = rsTemp.RecordCount
    
    ReDim aszTemp(1 To nCount + 11, 1 To 2)
    aszTemp(1, 1) = "总人数"
    aszTemp(2, 1) = "总金额"
    For i = 3 To nCount + 2
        aszTemp(i, 1) = cszTicketPriceStr & FormatDbValue(rsTemp!chinese_name)
        rsTemp.MoveNext
    Next i
    
    aszTemp(nCount + 3, 1) = cszSettleStr & "结算价"
    aszTemp(nCount + 4, 1) = cszSettleStr & "通行费"
    aszTemp(nCount + 5, 1) = cszSettleStr & "劳务费"
    aszTemp(nCount + 6, 1) = cszTicketPriceStr & "人公里"
    aszTemp(nCount + 7, 1) = cszSettleStr & "公里"
    aszTemp(nCount + 8, 1) = cszSettleStr & "春运费"
    aszTemp(nCount + 9, 1) = "区域费率"
    aszTemp(nCount + 10, 1) = cszSettleStr & "250K加成"
    aszTemp(nCount + 11, 1) = "路单数"
    
    
    
    GetFormulaItem = aszTemp
    
End Function



Public Function GetProtocolFormulaCompany(Optional pszCompanyID As String, Optional pbIsDefault As Boolean) As Recordset
    '得到协议公式
    '得到所有的该公司的协议


    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szWhere As String
    oDB.ConnectionString = GetConnectionStr()

'    AssertHaveRight RIGHT_GetCompanyProtocol

    If pszCompanyID <> "" Then
        szWhere = szWhere & " AND c.transport_company_id = " & TransFieldValueToString(pszCompanyID)
        szSql = " SELECT p.* , s.split_item_name , s.split_status , s.split_type , s.allow_modify,f.formula_comment,f.split_item_id,l.route_id " _
        & " FROM  " & cszTableCompany & " c, fin_settle_protocol_info p , split_item_info s,fin_charge_item_info f,fin_compnay_protocol_lst l  " _
        & " WHERE f.protocol_id=p.protocol_id AND f.split_item_id=s.split_item_id AND l.transport_company_id=c.transport_company_id AND l.protocol_id=p.protocol_id " _
        & " AND s.split_status = " & TransFieldValueToString(CS_SplitItemUse) _
        & szWhere & " ORDER BY f.protocol_id , l.route_id , f.split_item_id "


    ElseIf pbIsDefault Then
        szWhere = szWhere & " AND default_mark = " & TransFieldValueToString(EDefault.Default)
        szSql = " SELECT p.* , s.split_item_name , s.split_status , s.split_type , s.allow_modify,f.formula_comment,f.split_item_id " _
        & " FROM  fin_settle_protocol_info p , split_item_info s,fin_charge_item_info f " _
        & " WHERE f.protocol_id=p.protocol_id AND f.split_item_id=s.split_item_id AND p.default_mark=0 " _
        & " AND s.split_status = " & TransFieldValueToString(CS_SplitItemUse)
    End If

    Set rsTemp = oDB.Execute(szSql)
    Set GetProtocolFormulaCompany = rsTemp
    Set oDB = Nothing
    Set rsTemp = Nothing
    
End Function



Public Function GetProtocolFormulaVehicle(Optional pszVehicleID As String, Optional pbIsDefault As Boolean) As Recordset
    '得到协议公式
    '得到所有的该车辆的协议


    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szWhere As String

'    AssertHaveRight RIGHT_GetCompanyProtocol


    If pszVehicleID <> "" Then
        szWhere = szWhere & " AND c.vehicle_id = " & TransFieldValueToString(pszVehicleID)
    End If
    If pbIsDefault Then
        szWhere = szWhere & " AND default_mark = " & TransFieldValueToString(EDefault.Default)
    End If
    oDB.ConnectionString = GetConnectionStr()
    szSql = " SELECT p.* , s.split_item_name , s.split_status , s.split_type , s.allow_modify ,f.formula_comment,f.split_item_id " _
        & " FROM " & cszTableVehicle & "  c , fin_settle_protocol_info p , split_item_info s ,fin_charge_item_info f " _
        & " WHERE c.protocol_id = p.protocol_id AND f.protocol_id=p.protocol_id AND  f.split_item_id = s.split_item_id AND s.split_status = " & TransFieldValueToString(CS_SplitItemUse) _
        & szWhere

    Set rsTemp = oDB.Execute(szSql)
    Set GetProtocolFormulaVehicle = rsTemp
    Set oDB = Nothing
    Set rsTemp = Nothing
    
End Function

Public Function GetProtocolFormulaBus(Optional pszBusID As String, Optional pszCompanyID As String, Optional pbIsDefault As Boolean) As Recordset
    '得到协议公式
    '得到所有的该车次的协议


    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szWhere As String
    Dim szWhere2 As String
    
'    AssertHaveRight RIGHT_GetCompanyProtocol


    If pszBusID <> "" Then
        szWhere = szWhere & " AND p.bus_id = " & TransFieldValueToString(pszBusID)
    End If
    If pbIsDefault Then
        szWhere = szWhere & " AND default_mark = " & TransFieldValueToString(EDefault.Default)
    End If
    If pszCompanyID <> "" Then
        szWhere2 = szWhere2 & " AND p.transport_company_id = " & TransFieldValueToString(pszCompanyID)
    End If
    oDB.ConnectionString = GetConnectionStr()
'    szSql = " SELECT p.* , s.split_item_name , s.split_status , s.split_type , s.allow_modify ,f.formula_comment,f.split_item_id " _
'        & " FROM " & cszTableVehicle & "  c , fin_settle_protocol_info p , split_item_info s ,fin_charge_item_info f " _
'        & " WHERE c.protocol_id = p.protocol_id AND f.protocol_id=p.protocol_id AND  f.split_item_id = s.split_item_id AND s.split_status = " & TransFieldValueToString(CS_SplitItemUse) _
'        & szWhere
'
'    szSql = "SELECT p.* , s.protocol_name , s.default_mark , r.route_name , c.transport_company_short_name , l.split_item_name , l.split_status , l.split_type , l.allow_modify , f.formula_comment,f.split_item_id  " _
'        & " FROM fin_bus_protocol_lst p, bus_info b , fin_settle_protocol_info s,route_info r,company_info c,split_item_info l ,fin_charge_item_info f " _
'        & " WHERE p.bus_id = b.bus_id AND p.protocol_id = s.protocol_id AND r.route_id=b.route_id AND c.transport_company_id=p.transport_company_id and f.protocol_id=p.protocol_id AND  f.split_item_id = l.split_item_id AND l.split_status = " & TransFieldValueToString(CS_SplitItemUse) _
'        & szWhere
    szSql = "SELECT p.* , s.protocol_name , s.default_mark ,l.split_item_name , l.split_status , l.split_type , l.allow_modify , f.formula_comment,f.split_item_id  " _
        & " FROM fin_bus_protocol_lst p,  fin_settle_protocol_info s,split_item_info l ,fin_charge_item_info f " _
        & " WHERE p.protocol_id = s.protocol_id  and f.protocol_id=p.protocol_id AND  f.split_item_id = l.split_item_id AND l.split_status = " & TransFieldValueToString(CS_SplitItemUse) _
        & szWhere
    Set rsTemp = oDB.Execute(szSql & szWhere2)
    If rsTemp.RecordCount = 0 Then
        '如果该公司的协议不存在, 则查询该车次的协议.不加公司条件
        
        Set rsTemp = oDB.Execute(szSql)
    End If
    Set GetProtocolFormulaBus = rsTemp
    Set oDB = Nothing
    Set rsTemp = Nothing
    
End Function

Private Function GetSettlePrice(rsProtocol As Recordset, adbSplitItemPrice As Double) As Double
    '得到应结票款
    '根据传入的结算项的结算类型,及各结算项的金额,算出应结票款
    Dim i As Integer
    Dim dbTemp As Double
    Dim dblimitCharge As Double
    
        If FormatDbValue(rsProtocol!split_type) = CS_SplitOtherCompany Then
            dbTemp = adbSplitItemPrice
        End If
    GetSettlePrice = dbTemp
End Function

Private Function GetSettleStationPrice(rsProtocol As Recordset, adbSplitItemPrice As Double) As Double
    '得到结给车站票款
    '根据传入的结算项的结算类型,及各结算项的金额,算出结给车站票款
    Dim i As Integer
    Dim dbTemp As Double
    Dim dblimitCharge As Double

        If FormatDbValue(rsProtocol!split_type) = CS_SplitStation Then
             dbTemp = adbSplitItemPrice
        End If

    GetSettleStationPrice = dbTemp
    
End Function

'取得底限费用
Private Function GetlimitCharge(ByVal pszProtocolID As String, ByVal pszSplitItemID As String, ByRef pdbUpCharge As Double) As Double
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr()
    szSql = "SELECT * FROM fin_charge_item_info WHERE protocol_id=" & TransFieldValueToString(pszProtocolID) _
        & " AND split_item_id=" & TransFieldValueToString(pszSplitItemID)
    Set rsTemp = oDB.Execute(szSql)
    pdbUpCharge = FormatDbValue(rsTemp!up_charge)
    GetlimitCharge = FormatDbValue(rsTemp!limit_charge)
    
End Function

Private Sub SetFormulaValueCompany(rsPassageNumber As Recordset, aszFormulaValue() As String, nCount As Integer, rsCompanySettlePrice As Recordset _
 , rsProtocol As Recordset, atCompanyInfo() As TCompnaySettle, szCompanyID As String, szCompanyName As String, bHalve As Boolean, bDefault As Boolean, dbHalveRatio As Double)
    'bHalve 为TRUE 表示是否加总平分
    'dbHalveRatio表示加总平分时取的比率
    
    
    
    Dim k As Integer
    Dim dbSplitItem As Double
    Dim l As Integer
    Dim j As Integer '判断是否为同一公司
    Dim i As Integer
    Dim bUseSettlePrice As Boolean '是否使用结算价相关项
    
    aszFormulaValue(1, 2) = FormatDbValue(rsPassageNumber!Quantity) '总人数
    aszFormulaValue(2, 2) = FormatDbValue(rsPassageNumber!ticket_price) '票价总金额
    aszFormulaValue(3, 2) = FormatDbValue(rsPassageNumber!base_carriage)
    aszFormulaValue(nCount + 9, 2) = Val(FormatDbValue(rsPassageNumber!Annotation)) '区域费率
    aszFormulaValue(nCount + 10, 2) = 0 '初始化
    aszFormulaValue(nCount + 11, 2) = Val(FormatDbValue(rsPassageNumber!check_sheet_count)) '路单数
    For k = 4 To nCount + 2
        aszFormulaValue(k, 2) = rsPassageNumber("price_item_" & CInt(k - 3)).Value
    Next k
    
    bUseSettlePrice = False
    '如果未使用结算价相关项
    If rsCompanySettlePrice Is Nothing Then
        bUseSettlePrice = True
    ElseIf rsCompanySettlePrice.RecordCount = 0 Then
        bUseSettlePrice = True
    Else
        bUseSettlePrice = False
        '当公式设置与结算价相关时，如果票种为半价或各优惠票时，按照结算价半价来拆。
        If FormatDbValue(rsPassageNumber!ticket_type) = TP_FullPrice Or FormatDbValue(rsPassageNumber!ticket_type) = TP_FreeTicket Then
            '如果为全票或免票,则取全价
    
            If bHalve Then
                aszFormulaValue(nCount + 3, 2) = FormatDbValue(rsCompanySettlePrice!halve_full_price)
            Else
                aszFormulaValue(nCount + 3, 2) = FormatDbValue(rsCompanySettlePrice!settle_full_price)
            End If
            aszFormulaValue(nCount + 5, 2) = FormatDbValue(rsCompanySettlePrice!service_full_price)
            '拆算公里
            aszFormulaValue(nCount + 7, 2) = FormatDbValue(rsCompanySettlePrice!Mileage)
            aszFormulaValue(nCount + 8, 2) = FormatDbValue(rsCompanySettlePrice!spring_full_price)
            
        Else
            '否则取半价
            If bHalve Then
                aszFormulaValue(nCount + 3, 2) = FormatDbValue(rsCompanySettlePrice!halve_half_price)
            Else
                aszFormulaValue(nCount + 3, 2) = FormatDbValue(rsCompanySettlePrice!settle_half_price)
            End If
            aszFormulaValue(nCount + 5, 2) = FormatDbValue(rsCompanySettlePrice!service_half_price)
            '拆算公里
            aszFormulaValue(nCount + 7, 2) = FormatDbValue(rsCompanySettlePrice!Mileage) / 2
            aszFormulaValue(nCount + 8, 2) = FormatDbValue(rsCompanySettlePrice!spring_half_price)
        End If
        
        aszFormulaValue(nCount + 4, 2) = FormatDbValue(rsCompanySettlePrice!pass_charge)
        aszFormulaValue(nCount + 10, 2) = IIf(FormatDbValue(rsCompanySettlePrice!Mileage) >= 250, 1, 0)
        
    End If
    
    If bUseSettlePrice Then
        aszFormulaValue(nCount + 3, 2) = 0
        aszFormulaValue(nCount + 5, 2) = 0
        aszFormulaValue(nCount + 4, 2) = 0
    End If
    aszFormulaValue(nCount + 6, 2) = FormatDbValue(rsPassageNumber!Mileage)
    
    If FormatDbValue(rsPassageNumber!ticket_type) = TP_FullPrice Or FormatDbValue(rsPassageNumber!ticket_type) = TP_FreeTicket Then
        
    Else
        '半票的[票价]人公里为一半
        aszFormulaValue(nCount + 6, 2) = aszFormulaValue(nCount + 6, 2) / 2
        
    End If
    
    
    dbSplitItem = CaleLugValue(FormatDbValue(rsProtocol!formula_comment), aszFormulaValue) '根据协议得出相应的款额
    
    
    
    If bHalve Then
        '加总平分的取相应比率的里程
        '票价公里
        aszFormulaValue(nCount + 6, 2) = aszFormulaValue(nCount + 6, 2) * dbHalveRatio
        aszFormulaValue(nCount + 7, 2) = aszFormulaValue(nCount + 7, 2) * dbHalveRatio
    End If
    

    
    If bHalve Then
        '如果未用到结算价,且使用加总平分
        dbSplitItem = Round(dbSplitItem * dbHalveRatio + 0.0001, 3)
    Else
        dbSplitItem = Round(dbSplitItem + 0.0001, 3)
    End If
        
    For l = 1 To ArrayLength(atCompanyInfo)
        If atCompanyInfo(l).CompanyID = szCompanyID Then
            '如果该类型中已经有该公司
            Exit For
        End If
    Next l
    
    If l > ArrayLength(atCompanyInfo) Then
        If l = 1 Then
            ReDim atCompanyInfo(1 To l)
        Else
            ReDim Preserve atCompanyInfo(1 To l)
        End If
        atCompanyInfo(l).CompanyID = szCompanyID
        atCompanyInfo(l).CompanyName = szCompanyName
        atCompanyInfo(l).ProtocolID = FormatDbValue(rsProtocol!protocol_id)
        atCompanyInfo(l).ProtocolName = FormatDbValue(rsProtocol!protocol_name)
    End If
    
'     dblimitCharge = GetlimitCharge(FormatDbValue(rsProtocol!protocol_id), FormatDbValue(rsProtocol!split_item_id))
'     If dblimitCharge > dbSplitItem Then
'        dbSplitItem = dblimitCharge
'     Else
'        dbSplitItem = dbSplitItem
'     End If
    atCompanyInfo(l).SplitItem(CInt(rsProtocol!split_item_id)) = atCompanyInfo(l).SplitItem(CInt(rsProtocol!split_item_id)) + dbSplitItem
'    atCompanyInfo(l).SettlePrice = atCompanyInfo(l).SettlePrice + GetSettlePrice(rsProtocol, dbSplitItem)
'    atCompanyInfo(l).SettleStationPrice = atCompanyInfo(l).SettleStationPrice + GetSettleStationPrice(rsProtocol, dbSplitItem)
    If m_bIsNew = False Then
        atCompanyInfo(l).PassengerNumber = atCompanyInfo(l).PassengerNumber + aszFormulaValue(1, 2)  '人数
        atCompanyInfo(l).Mileage = atCompanyInfo(l).Mileage + aszFormulaValue(nCount + 6, 2) '人数
        
    End If
    
End Sub


Private Sub SetFormulaValueVehicle(rsPassageNumber As Recordset, aszFormulaValue() As String, nCount As Integer, rsVehicleSettlePrice As Recordset _
 , rsProtocol As Recordset, atVehicleInfo() As TVehilceSettle, szVehicleID As String, szLicenseTagNo As String, bHalve As Boolean, dbHalveRatio As Double)
    'bHalve 为TRUE 表示是否只取其中的一部分进行拆分
    Dim k As Integer
    Dim dbSplitItem As Double
    Dim l As Integer
    Dim i As Integer
    
    
    Dim bUseSettlePrice As Boolean '是否使用结算价相关项
    
    
    aszFormulaValue(1, 2) = FormatDbValue(rsPassageNumber!Quantity) '总人数
    aszFormulaValue(2, 2) = FormatDbValue(rsPassageNumber!ticket_price) '票价总金额
    aszFormulaValue(3, 2) = FormatDbValue(rsPassageNumber!base_carriage)
    aszFormulaValue(nCount + 9, 2) = Val(FormatDbValue(rsPassageNumber!Annotation)) '区域费率
    aszFormulaValue(nCount + 10, 2) = 0 '初始化
    aszFormulaValue(nCount + 11, 2) = Val(FormatDbValue(rsPassageNumber!check_sheet_count)) '路单数
    For k = 4 To nCount + 2
        aszFormulaValue(k, 2) = FormatDbValue(rsPassageNumber("price_item_" & CInt(k - 3)))
    Next k
    
    
    bUseSettlePrice = False
    If rsVehicleSettlePrice Is Nothing Then
        bUseSettlePrice = True
    ElseIf rsVehicleSettlePrice.RecordCount = 0 Then
        bUseSettlePrice = True
    Else
        bUseSettlePrice = False

        '当公式设置与结算价相关时，如果票种为半价或各优惠票时，按照结算价半价来拆。
        If FormatDbValue(rsPassageNumber!ticket_type) = TP_FullPrice Or FormatDbValue(rsPassageNumber!ticket_type) = TP_FreeTicket Then
            '如果为全票或免票,则取全价
            If bHalve Then
                aszFormulaValue(nCount + 3, 2) = FormatDbValue(rsVehicleSettlePrice!halve_full_price)
            Else
                aszFormulaValue(nCount + 3, 2) = FormatDbValue(rsVehicleSettlePrice!settle_full_price)
            End If
            aszFormulaValue(nCount + 5, 2) = FormatDbValue(rsVehicleSettlePrice!service_full_price)
            '拆算公里
            aszFormulaValue(nCount + 7, 2) = FormatDbValue(rsVehicleSettlePrice!Mileage)
            aszFormulaValue(nCount + 8, 2) = FormatDbValue(rsVehicleSettlePrice!spring_full_price)
        Else
            '否则取半价
            If bHalve Then
                aszFormulaValue(nCount + 3, 2) = FormatDbValue(rsVehicleSettlePrice!halve_half_price)
            Else
                aszFormulaValue(nCount + 3, 2) = FormatDbValue(rsVehicleSettlePrice!settle_half_price)
            End If
            aszFormulaValue(nCount + 5, 2) = FormatDbValue(rsVehicleSettlePrice!service_half_price)
            '拆算公里
            aszFormulaValue(nCount + 7, 2) = FormatDbValue(rsVehicleSettlePrice!Mileage) / 2
            aszFormulaValue(nCount + 8, 2) = FormatDbValue(rsVehicleSettlePrice!spring_half_price)
        End If
        aszFormulaValue(nCount + 4, 2) = FormatDbValue(rsVehicleSettlePrice!pass_charge)
        aszFormulaValue(nCount + 10, 2) = IIf(FormatDbValue(rsVehicleSettlePrice!Mileage) >= 250, 1, 0)
    End If
    If bUseSettlePrice Then
        aszFormulaValue(nCount + 3, 2) = 0
        aszFormulaValue(nCount + 5, 2) = 0
        aszFormulaValue(nCount + 4, 2) = 0
    End If
    aszFormulaValue(nCount + 6, 2) = FormatDbValue(rsPassageNumber!Mileage)
    
    If FormatDbValue(rsPassageNumber!ticket_type) = TP_FullPrice Or FormatDbValue(rsPassageNumber!ticket_type) = TP_FreeTicket Then
        
    Else
        '半票的[票价]人公里为一半
        aszFormulaValue(nCount + 6, 2) = aszFormulaValue(nCount + 6, 2) / 2
    End If
'    dbSplitItem = Round(CaleLugValue(FormatDbValue(rsProtocol!formula_comment), aszFormulaValue) + 0.001, 2) '跟据协议得出相应的款额


    dbSplitItem = CaleLugValue(FormatDbValue(rsProtocol!formula_comment), aszFormulaValue) '根据协议得出相应的款额
    
    If bHalve Then
        '票价公里
        aszFormulaValue(nCount + 6, 2) = aszFormulaValue(nCount + 6, 2) * dbHalveRatio
        aszFormulaValue(nCount + 7, 2) = aszFormulaValue(nCount + 7, 2) * dbHalveRatio
        
    End If
    
    If bHalve Then
        '如果使用加总平分
        dbSplitItem = Round(dbSplitItem * dbHalveRatio + 0.0001, 3)
    Else
        dbSplitItem = Round(dbSplitItem + 0.0001, 3)
    End If
    For l = 1 To ArrayLength(atVehicleInfo)
        If atVehicleInfo(l).VehicleID = szVehicleID Then
            '如果该类型中已经有该公司
            Exit For
        End If
    Next l
    If l > ArrayLength(atVehicleInfo) Then
        If l = 1 Then
            ReDim atVehicleInfo(1 To l)
        Else
            ReDim Preserve atVehicleInfo(1 To l)
        End If
        atVehicleInfo(l).VehicleID = szVehicleID
        atVehicleInfo(l).LicenseTagNo = szLicenseTagNo
        atVehicleInfo(l).ProtocolID = FormatDbValue(rsProtocol!protocol_id)
        atVehicleInfo(l).ProtocolName = FormatDbValue(rsProtocol!protocol_name)
        
    End If
    
'    dblimitCharge = GetlimitCharge(FormatDbValue(rsProtocol!protocol_id), FormatDbValue(rsProtocol!split_item_id))
'     If dblimitCharge > dbSplitItem Then
'        dbSplitItem = dblimitCharge
'     Else
'        dbSplitItem = dbSplitItem
'     End If
    atVehicleInfo(l).SplitItem(CInt(rsProtocol!split_item_id)) = atVehicleInfo(l).SplitItem(CInt(rsProtocol!split_item_id)) + dbSplitItem
'    atVehicleInfo(l).SettlePrice = atVehicleInfo(1).SettlePrice + GetSettlePrice(rsProtocol, dbSplitItem)
'    atVehicleInfo(l).SettleStationPrice = atVehicleInfo(l).SettleStationPrice + GetSettleStationPrice(rsProtocol, dbSplitItem)
    If m_bIsNew = False Then
        atVehicleInfo(l).PassengerNumber = atVehicleInfo(l).PassengerNumber + aszFormulaValue(1, 2) '人数
        atVehicleInfo(l).Mileage = atVehicleInfo(l).Mileage + aszFormulaValue(nCount + 6, 2) '人数
    End If
End Sub

'得到公司名字
Private Function GetHavleOtherCompanyName(pszCompanyID As String) As String
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr()
    szSql = "SELECT transport_company_short_name FROM " & cszTableCompany & " WHERE transport_company_id= " & TransFieldValueToString(pszCompanyID)
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    GetHavleOtherCompanyName = FormatDbValue(rsTemp!transport_company_short_name)
    
End Function

Private Sub SetFormulaValueBus(rsPassageNumber As Recordset, aszFormulaValue() As String, nCount As Integer, rsBusSettlePrice As Recordset _
 , rsProtocol As Recordset, atBusInfo() As TBusSettle, szBusID As String, szLicenseTagNo As String, bHalve As Boolean, dbHalveRatio As Double)
    'bHalve 为TRUE 表示是否只取其中的一部分进行拆分
    Dim k As Integer
    Dim dbSplitItem As Double
    Dim l As Integer
    Dim i As Integer
    
    
    Dim bUseSettlePrice As Boolean '是否使用结算价相关项
    
    
    aszFormulaValue(1, 2) = FormatDbValue(rsPassageNumber!Quantity) '总人数
    aszFormulaValue(2, 2) = FormatDbValue(rsPassageNumber!ticket_price) '票价总金额
    aszFormulaValue(3, 2) = FormatDbValue(rsPassageNumber!base_carriage)
    aszFormulaValue(nCount + 9, 2) = Val(FormatDbValue(rsPassageNumber!Annotation)) '区域费率
    aszFormulaValue(nCount + 10, 2) = 0 '初始化
    aszFormulaValue(nCount + 11, 2) = Val(FormatDbValue(rsPassageNumber!check_sheet_count)) '路单数
    For k = 4 To nCount + 2
        aszFormulaValue(k, 2) = FormatDbValue(rsPassageNumber("price_item_" & CInt(k - 3)))
    Next k
    
    
    bUseSettlePrice = False
    If rsBusSettlePrice Is Nothing Then
        bUseSettlePrice = True
    ElseIf rsBusSettlePrice.RecordCount = 0 Then
        bUseSettlePrice = True
    Else
        bUseSettlePrice = False

        '当公式设置与结算价相关时，如果票种为半价或各优惠票时，按照结算价半价来拆。
        If FormatDbValue(rsPassageNumber!ticket_type) = TP_FullPrice Or FormatDbValue(rsPassageNumber!ticket_type) = TP_FreeTicket Then
            '如果为全票或免票,则取全价
            If bHalve Then
                aszFormulaValue(nCount + 3, 2) = FormatDbValue(rsBusSettlePrice!halve_full_price)
            Else
                aszFormulaValue(nCount + 3, 2) = FormatDbValue(rsBusSettlePrice!settle_full_price)
            End If
            aszFormulaValue(nCount + 5, 2) = FormatDbValue(rsBusSettlePrice!service_full_price)
            '拆算公里
            aszFormulaValue(nCount + 7, 2) = FormatDbValue(rsBusSettlePrice!Mileage)
            aszFormulaValue(nCount + 8, 2) = FormatDbValue(rsBusSettlePrice!spring_full_price)
        Else
            '否则取半价
            If bHalve Then
                aszFormulaValue(nCount + 3, 2) = FormatDbValue(rsBusSettlePrice!halve_half_price)
            Else
                aszFormulaValue(nCount + 3, 2) = FormatDbValue(rsBusSettlePrice!settle_half_price)
            End If
            aszFormulaValue(nCount + 5, 2) = FormatDbValue(rsBusSettlePrice!service_half_price)
            '拆算公里
            aszFormulaValue(nCount + 7, 2) = FormatDbValue(rsBusSettlePrice!Mileage) / 2
            aszFormulaValue(nCount + 8, 2) = FormatDbValue(rsBusSettlePrice!spring_half_price)
        End If
        aszFormulaValue(nCount + 4, 2) = FormatDbValue(rsBusSettlePrice!pass_charge)
        aszFormulaValue(nCount + 10, 2) = IIf(FormatDbValue(rsBusSettlePrice!Mileage) >= 250, 1, 0)
    End If
    If bUseSettlePrice Then
        aszFormulaValue(nCount + 3, 2) = 0
        aszFormulaValue(nCount + 5, 2) = 0
        aszFormulaValue(nCount + 4, 2) = 0
    End If
    aszFormulaValue(nCount + 6, 2) = FormatDbValue(rsPassageNumber!Mileage)
    
    If FormatDbValue(rsPassageNumber!ticket_type) = TP_FullPrice Or FormatDbValue(rsPassageNumber!ticket_type) = TP_FreeTicket Then
        
    Else
        '半票的[票价]人公里为一半
        aszFormulaValue(nCount + 6, 2) = aszFormulaValue(nCount + 6, 2) / 2
    End If
'    dbSplitItem = Round(CaleLugValue(FormatDbValue(rsProtocol!formula_comment), aszFormulaValue) + 0.001, 2) '跟据协议得出相应的款额


    dbSplitItem = CaleLugValue(FormatDbValue(rsProtocol!formula_comment), aszFormulaValue) '根据协议得出相应的款额
    
    If bHalve Then
        '票价公里
        aszFormulaValue(nCount + 6, 2) = aszFormulaValue(nCount + 6, 2) * dbHalveRatio
        aszFormulaValue(nCount + 7, 2) = aszFormulaValue(nCount + 7, 2) * dbHalveRatio
        
    End If
    
    If bHalve Then
        '如果使用加总平分
        dbSplitItem = Round(dbSplitItem * dbHalveRatio + 0.0001, 3)
    Else
        dbSplitItem = Round(dbSplitItem + 0.0001, 3)
    End If
    For l = 1 To ArrayLength(atBusInfo)
        If atBusInfo(l).BusID = szBusID Then
            '如果该类型中已经有该公司
            Exit For
        End If
    Next l
    If l > ArrayLength(atBusInfo) Then
        If l = 1 Then
            ReDim atBusInfo(1 To l)
        Else
            ReDim Preserve atBusInfo(1 To l)
        End If
        atBusInfo(l).BusID = szBusID
        atBusInfo(l).LicenseTagNo = szLicenseTagNo
        atBusInfo(l).ProtocolID = FormatDbValue(rsProtocol!protocol_id)
        atBusInfo(l).ProtocolName = FormatDbValue(rsProtocol!protocol_name)
        
    End If
    
'    dblimitCharge = GetlimitCharge(FormatDbValue(rsProtocol!protocol_id), FormatDbValue(rsProtocol!split_item_id))
'     If dblimitCharge > dbSplitItem Then
'        dbSplitItem = dblimitCharge
'     Else
'        dbSplitItem = dbSplitItem
'     End If
    atBusInfo(l).SplitItem(CInt(rsProtocol!split_item_id)) = atBusInfo(l).SplitItem(CInt(rsProtocol!split_item_id)) + dbSplitItem
'    atBusInfo(l).SettlePrice = atBusInfo(1).SettlePrice + GetSettlePrice(rsProtocol, dbSplitItem)
'    atBusInfo(l).SettleStationPrice = atBusInfo(l).SettleStationPrice + GetSettleStationPrice(rsProtocol, dbSplitItem)
    If m_bIsNew = False Then
        atBusInfo(l).PassengerNumber = atBusInfo(l).PassengerNumber + aszFormulaValue(1, 2) '人数
        atBusInfo(l).Mileage = atBusInfo(l).Mileage + aszFormulaValue(nCount + 6, 2) '人数
    End If
End Sub


'===================================================以下是计算值的函数

Private Function CaleLugValue(Formula As String, atProtocolChargeItem() As String) As Double
    Dim szTemp As String
    Dim dbValue As Double
    
    On Error GoTo ErrorHandle
    
    szTemp = FullStr(Trim(Formula), atProtocolChargeItem)
    'szTemp = "(9*10+12*10)*0.8"
    If Trim(szTemp) <> "" Then
        dbValue = moCalc.CalculateIt(szTemp)
    Else
        dbValue = 0
    End If
    CaleLugValue = dbValue
    Exit Function
ErrorHandle:
    err.Raise err.Number, err.Source, "公式错误,公式名为:" & szTemp
    
    
End Function

Private Function FullStr(pszFormula As String, paszKeyValue() As String) As String
    '将公式字符串转换为值
    
    Dim nCount As Integer
    Dim i As Integer
    Dim szTemp As String, nLastSeat As Integer
    Dim nFormulaLen As Integer
    Dim nTemp As Integer, j As Integer
    Dim bContinue As Boolean
    Dim nTag As Integer
    pszFormula = pszFormula
    nCount = UBound(paszKeyValue, 1)
    nFormulaLen = Len(pszFormula)
    nLastSeat = 1
    bContinue = True
    nTag = 1
    Do While bContinue
        bContinue = False

        nTemp = GetLowestSeat(nLastSeat, pszFormula, paszKeyValue, j)

        If nTemp > 0 Then
            szTemp = szTemp & Mid(pszFormula, nLastSeat, nTemp - nLastSeat)
            If paszKeyValue(j, 2) < 0 Then
'               nTag = -1
'               paszKeyValue(j, 2) = Abs(paszKeyValue(j, 2))
                szTemp = szTemp & "(0-" & Abs(paszKeyValue(j, 2)) & ")"
            Else
                szTemp = szTemp & paszKeyValue(j, 2)
            End If
'            szTemp = szTemp & paszKeyValue(j, 2)
            nTemp = Len(paszKeyValue(j, 1)) + nTemp
            nLastSeat = nTemp
            bContinue = True
        End If

    Loop
    If nLastSeat <= nFormulaLen Then
        szTemp = szTemp & Right(pszFormula, nFormulaLen - nLastSeat + 1)
    End If
    FullStr = szTemp
End Function

'##ModelId=3DDD20600241
Private Function GetLowestSeat(pnStart As Integer, pszFormula As String, paszKeyValue() As String, pnIndex As Integer) As Integer '得到最小字符串
    Dim nCount As Integer
    Dim nResult As Integer, i As Integer
    Dim nLen As Integer
    Dim nIndex As Integer
    nCount = UBound(paszKeyValue, 1)
    nIndex = 0
    nResult = Len(pszFormula)
    nLen = Len(pszFormula)
    For i = 1 To nCount
        nLen = InStr(pnStart, pszFormula, paszKeyValue(i, 1), vbTextCompare)
        If nLen < nResult And nLen <> 0 And IsLastCharLegalSeperator(pszFormula, Len(paszKeyValue(i, 1)) + nLen) Then
            nResult = nLen
            nIndex = i
        End If
    Next
    If nIndex = 0 Then
        GetLowestSeat = 0
    Else
        GetLowestSeat = nResult
    End If
    pnIndex = nIndex
End Function

'##ModelId=3DDD2060030C
Private Function IsLastCharLegalSeperator(ByVal pszString As String, ByVal pnIndex As Integer) As Boolean   '判断是否是最后公式项
    Dim szChar As String
    On Error GoTo Error_HANDle
    szChar = Mid(pszString, pnIndex, 1)
    If szChar = " " Or szChar = "+" Or szChar = "-" Or szChar = "*" Or szChar = "/" Or szChar = ")" Or szChar = "" Then
        IsLastCharLegalSeperator = True
    Else
        IsLastCharLegalSeperator = False
    End If
    Exit Function
Error_HANDle:
    IsLastCharLegalSeperator = True
End Function




'===================================================以上是计算值的函数

Public Function DeleteCompanyProtocol(CompanyID As String, RouteID As String)
    On Error GoTo err
    Dim szSql As String
    Dim oDB As New RTConnection
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_SetCompanyProtocol
    
    oDB.ConnectionString = GetConnectionStr("")
    szSql = "DELETE FROM fin_compnay_protocol_lst " _
        & " WHERE transport_company_id = " & TransFieldValueToString(CompanyID) _
        & " AND route_id = " & TransFieldValueToString(RouteID)
    oDB.Execute (szSql)
    Exit Function
err:
    Exit Function
End Function



Private Sub FindRouteProtocol(prsProtocol As Recordset, pszRouteID As String, pbDefault As Boolean)
    '查找某线路的协议的位置
    Dim j As Integer
    
    If Not pbDefault Then
        '找到该线路的起始协议
        prsProtocol.MoveFirst
        For j = 1 To prsProtocol.RecordCount
            If FormatDbValue(prsProtocol!route_id) = Trim(pszRouteID) Then
                
                Exit For
            End If
            prsProtocol.MoveNext
        Next j
        If j > prsProtocol.RecordCount Then RaiseError ERR_CompanyNotSetProtocol
    End If
End Sub

Private Function GetCompanySettlePrice(pszCompanyID As String, pszVehicleTypeCode As String, pszRouteID As String, pszSellStationID As String, pszStationID As String) As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsCompanySettlePrice As Recordset

    oDB.ConnectionString = GetConnectionStr
    
    
    
    '得到这些路单的公司、上车站、线路、站点条件的公司结算全价、结算半价、结算通行费rsCompanySettlePrice
    szSql = " SELECT * FROM fin_company_settle_price_lst WHERE transport_company_id = " & TransFieldValueToString(pszCompanyID) _
        & " AND vehicle_type_code = " & TransFieldValueToString(pszVehicleTypeCode) _
        & " AND route_id = " & TransFieldValueToString(pszRouteID) _
        & " AND sell_station_id = " & TransFieldValueToString(pszSellStationID) _
        & " AND station_id = " & TransFieldValueToString(pszStationID)
    Set rsCompanySettlePrice = oDB.Execute(szSql)
    
    '判断结算价是否存在,没有触发错误
    
    If rsCompanySettlePrice.RecordCount = 0 Then err.Raise ERR_Split, "结算预览", " 线路=" & Trim(TransFieldValueToString(pszRouteID)) & ";    公司=" & Trim(TransFieldValueToString(pszCompanyID)) & ";    车型=" & Trim(TransFieldValueToString(pszVehicleTypeCode)) & Chr(10) & Chr(13) & "    站点= " & Trim(TransFieldValueToString(pszStationID)) & "   的结算价不存在!"
    
    
    Set GetCompanySettlePrice = rsCompanySettlePrice

End Function
                    

Private Function GetVehicleSettlePrice(pszVehicle As String, pszRouteID As String, pszSellStationID As String, pszStationID As String) As Recordset

    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsVehicleSettlePrice As Recordset

    oDB.ConnectionString = GetConnectionStr
    
    
    
                
    '查找该车辆、上车站、线路、站点、票种的售票票价项及结算全价、结算运费、结算半价。
    szSql = " SELECT * FROM fin_vehicle_settle_price_lst WHERE vehicle_id = " & TransFieldValueToString(pszVehicle) _
        & " AND route_id = " & TransFieldValueToString(pszRouteID) _
        & " AND sell_station_id = " & TransFieldValueToString(pszSellStationID) _
        & " AND station_id = " & TransFieldValueToString(pszStationID)
    Set rsVehicleSettlePrice = oDB.Execute(szSql)
    
    '判断结算价是否存在,没有触发错误
    If rsVehicleSettlePrice.RecordCount = 0 Then err.Raise ERR_Split, "结算预览", " 线路= " & Trim(TransFieldValueToString(pszRouteID)) & ";   车辆=" & Trim(TransFieldValueToString(pszVehicle)) & Chr(10) & Chr(13) & "  站点= " & Trim(TransFieldValueToString(pszStationID)) & "   的结算价不存在!"
                
    Set GetVehicleSettlePrice = rsVehicleSettlePrice
    
    
End Function

Private Function GetBusSettlePrice(pszBus As String, pszRouteID As String, pszSellStationID As String, pszStationID As String) As Recordset

    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsBusSettlePrice As Recordset

    oDB.ConnectionString = GetConnectionStr
    
    
    
                
    '查找该车辆、上车站、线路、站点、票种的售票票价项及结算全价、结算运费、结算半价。
    szSql = " SELECT * FROM fin_bus_settle_price_lst WHERE bus_id = " & TransFieldValueToString(pszBus) _
        & " AND sell_station_id = " & TransFieldValueToString(pszSellStationID) _
        & " AND station_id = " & TransFieldValueToString(pszStationID)
    Set rsBusSettlePrice = oDB.Execute(szSql)
    
    
    '& " AND route_id = " & TransFieldValueToString(pszRouteID)

    

    '判断结算价是否存在,没有触发错误
    If rsBusSettlePrice.RecordCount = 0 Then err.Raise ERR_Split, "结算预览", " 线路= " & Trim(TransFieldValueToString(pszRouteID)) & ";   车辆=" & Trim(TransFieldValueToString(pszBus)) & Chr(10) & Chr(13) & "  站点= " & Trim(TransFieldValueToString(pszStationID)) & "   的结算价不存在!"
                
    Set GetBusSettlePrice = rsBusSettlePrice
    
    
End Function

Private Function FindSettlePriceStr(prsProtocol As Recordset, pnVailableNumber As Integer, cszSettlePrice As String) As Boolean
    Dim rsTemp As Recordset
    Dim bExist As Boolean
    Dim i As Integer
    
    Set rsTemp = prsProtocol
    bExist = False
    For i = 1 To pnVailableNumber
        
        If InStr(1, FormatDbValue(rsTemp!formula_comment), cszSettlePrice) > 0 Then
            bExist = True
            rsTemp.MoveNext
            Exit For
        End If
        rsTemp.MoveNext
    Next i
    If i > pnVailableNumber Then
        rsTemp.Move -i + 1
    Else
        rsTemp.Move -i
    End If
    FindSettlePriceStr = bExist
    
End Function

'##ModelId=3E81F475006D
'按照传入的站点人数来进行结算预览
'返回预结算后的信息,此接口并不写数据库,只是将拆算结果返回.
Public Function PreviewSplitCheckSheetManual(pszCompanyID As String, pszVehicleID As String, paszCheckSheetID() As String, patStationQuantity() As TSettleSheetStation, Optional pbIsSingle As Boolean = False) As TSplitResult
'    ******注  pbIsSingle 参数表是是否强行按照单行结算 true按单行, false按设置结算
    
    On Error GoTo PreviewSplitCheckSheetManualErr


    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection


    Dim bSplitByCompany As Boolean
'    Dim nSheetCount As Integer '路单数
    Dim szWhere As String


'    Dim rsPassageNumber As Recordset '各公司、上车站、线路、站点、票种的人数及各票价项信息
    Dim rsProtocol As Recordset
    Dim rsHalve As Recordset
    Dim rsCompanySettlePrice As Recordset
    Dim rsVehicleSettlePrice As Recordset
    Dim rsDefaultProtocol As Recordset '默认协议
    Dim rsHalveProtocol As Recordset '加总平分公司的协议


    Dim szHalveCompanyID As String '平分给对方的公司代码
    Dim szHalveCompanyName As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer

    Dim aszFormulaValue() As String
    Dim nCount As Integer '票价项数


'    Dim szFormuarName As String '公式名称


    Dim tResultInfo As TSplitResult '返回的结果信息

    Dim atCompanyInfo() As TCompnaySettle '公司结算明细
    Dim atVehicleInfo() As TVehilceSettle '车辆结算明细
'    Dim atStationInfo() As TSettleSheetStation '站点结算明细

    Dim dbSplitItem As Double '临时拆出来的某一项的金额
    Dim dbTotalTicketPrice As Double '总票款
    Dim dbTotalQuantity As Double '总人数

    Dim dbSettleOtherCompanyPrice As Double '总的结给对方公司金额
    Dim dbSettleStationPrice As Double '总的结给车站的金额

    Dim dbSplitlimitCharge As Double '底限额
    Dim dbSplitUpCharge As Double '上限额
    
    
    Dim bIsDefaut As Boolean '是否默认协议

    Dim nVailableNumber As Integer '可用的结算项的数目


    Dim nItem As Integer
    Dim szitem As String
    
    
    Dim szCompanyID As String
    Dim szCompanyName As String
    Dim rsCompanyInfo As Recordset



    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_SplitCheckSheet 'RIGHT_SplitCheckSheetManual '验证用户是否有手工结算的权限

    oDB.ConnectionString = GetConnectionStr()


    'if 公司代码不为'' then 按公司结算elseif 车辆代码不为'' then 按车辆结算else      出调用参数不正确错endif
    If pszCompanyID = "" And pszVehicleID = "" Then
        '出传入参数错
        RaiseError ERR_PreviewSplitInputParam
    ElseIf pszCompanyID <> "" Then
        bSplitByCompany = True
    Else
        bSplitByCompany = False
    End If
'    不判断路单数
'    nSheetCount = ArrayLength(paszCheckSheetID)
'    If nSheetCount = 0 Then RaiseError ERR_PreviewSplitInputParam '出传入参数错
'
'    szWhere = " ( "
'    For i = 1 To nSheetCount - 1
'            szWhere = szWhere & "'" & paszCheckSheetID(i) & "',"
'    Next i
'    szWhere = szWhere & "'" & paszCheckSheetID(i) & "')"
'
'
'
'    '判断每一张路单是否均存在，如不存在则出"路单不存在"的错。如无效或经过结算。则出"路单无效或路单已经结算"的错。
'    szSql = "SELECT * FROM " & cszTableCheckSheet & "  WHERE check_sheet_id IN " & szWhere
'    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount <> nSheetCount Then
'        '如果记录数不为传入的路单数,则出错
'        RaiseError ERR_SomeSheetNotExist
'    Else
        '判断所有的路单是不是同一公司或同一车辆
'        For i = 1 To rsTemp.RecordCount
'            If bSplitByCompany Then
'                If FormatDbValue(rsTemp!transport_company_id) <> pszCompanyID Then
'                    '出错
'                    RaiseError ERR_SplitSheetMustSameCompany
'                End If
'            Else
'                If FormatDbValue(rsTemp!vehicle_id) <> pszVehicleID Then
'                    '出错
'                    RaiseError ERR_SplitSheetMustSameVehicle
'                End If
'            End If
'            rsTemp.MoveNext
'        Next i
'    End If

    szSql = " SELECT * FROM price_item_info "
    Set rsTemp = oDB.Execute(szSql)

    nCount = rsTemp.RecordCount

    aszFormulaValue = GetFormulaItem


    '取出默认的协议
    bIsDefaut = False
    Set rsDefaultProtocol = GetProtocolFormulaCompany(, True)


    szSql = " SELECT * FROM split_item_info WHERE split_status = " & TransFieldValueToString(ESplitItemUseStatus.CS_SplitItemUse)
    Set rsTemp = oDB.Execute(szSql)
    nVailableNumber = rsTemp.RecordCount
    dbTotalQuantity = 0

    m_bIsNew = False

    '如果按公司结算
    If bSplitByCompany Then
        szSql = " SELECT * FROM " & cszTableCompany & " WHERE transport_company_id = " & TransFieldValueToString(pszCompanyID)
        Set rsCompanyInfo = oDB.Execute(szSql)
        If rsCompanyInfo.RecordCount = 0 Then
            RaiseError ERR_CompanyNotExist
        Else
            szCompanyName = FormatDbValue(rsCompanyInfo!transport_company_short_name)
        End If
        '2．   得到所有该公司需要与之加总平分的公司线路信息rsHalve
        szSql = " SELECT h.*, c.transport_company_short_name FROM fin_halve_company_lst h, " & cszTableCompany & " c " _
            & " WHERE c.transport_company_id = h.transport_company_id AND h.transport_company_id = " & TransFieldValueToString(pszCompanyID)
        
        Set rsHalve = oDB.Execute(szSql)
        
        '3． 得到该公司的协议
        Set rsProtocol = GetProtocolFormulaCompany(pszCompanyID)
        If rsProtocol.RecordCount = 0 Then
            '取出默认的协议
            If rsDefaultProtocol.RecordCount <> 0 Then
                Set rsProtocol = rsDefaultProtocol
                bIsDefaut = True
            Else
                '该公司协议未设置
                RaiseError ERR_CompanyNotSetProtocol
            End If
        Else
            bIsDefaut = False
        End If

        
        For i = 1 To ArrayLength(patStationQuantity)
            dbTotalQuantity = patStationQuantity(i).Quantity
            szHalveCompanyID = ""
            If rsHalve.RecordCount > 0 Then rsHalve.MoveFirst
            For j = 1 To rsHalve.RecordCount
                If patStationQuantity(i).RouteID = FormatDbValue(rsHalve!route_id) And pszCompanyID = FormatDbValue(rsHalve!transport_company_id) Then
                    '找到该公司在该线路上需要加总平分

                    szHalveCompanyID = FormatDbValue(rsHalve!other_transport_company_id)
                    szHalveCompanyName = GetHavleOtherCompanyName(szHalveCompanyID)
                    Exit For
                End If
                rsHalve.MoveNext
            Next j
            'If 该公司在rsHalve中找到需要加总平分then

            If szHalveCompanyID = "" Or pbIsSingle Then
                '不需要加总平分
                '找到该线路的起始协议

                FindRouteProtocol rsProtocol, patStationQuantity(i).RouteID, bIsDefaut

                If FindSettlePriceStr(rsProtocol, nVailableNumber, cszTicketPriceStr) Then
                    '出错   如果使用手工结算,不能用到"[票价]"的项目
                    RaiseError ERR_ManualSettleCannotUseTicketPriceStr
                End If


                '如果公式中没有用到结算价,则不进行查询结算价
                If FindSettlePriceStr(rsProtocol, nVailableNumber, cszSettleStr) Then
                    Set rsCompanySettlePrice = GetCompanySettlePrice(pszCompanyID, patStationQuantity(i).VehicleTypeCode, patStationQuantity(i).RouteID, patStationQuantity(i).SellSationID, patStationQuantity(i).StationID)
                End If

                For j = 1 To nVailableNumber

                    SetFormulaValueCompanyManual patStationQuantity(i), aszFormulaValue, nCount, rsCompanySettlePrice, rsProtocol, atCompanyInfo, pszCompanyID, szCompanyName, False, bIsDefaut
                    
                    rsProtocol.MoveNext
                    m_bIsNew = True
                Next j
'                atCompanyInfo(1).PassengerNumber = atCompanyInfo(1).PassengerNumber + FormatDbValue(rsPassageNumber!Quantity)
            Else
                '将这些人数及售票的票价项分成两半，一半拆给该公司，另一半拆给加总平分的公司。

                '先将一半拆给该公司
                '找到该线路的起始协议
                FindRouteProtocol rsProtocol, patStationQuantity(i).RouteID, bIsDefaut
                '如果公式中没有用到结算价,则不进行查询结算价
                If FindSettlePriceStr(rsProtocol, nVailableNumber, cszSettleStr) Then
                    Set rsCompanySettlePrice = GetCompanySettlePrice(pszCompanyID, patStationQuantity(i).VehicleTypeCode, patStationQuantity(i).RouteID, patStationQuantity(i).SellSationID, patStationQuantity(i).StationID)
                End If
                For j = 1 To nVailableNumber
                    SetFormulaValueCompanyManual patStationQuantity(i), aszFormulaValue, nCount, rsCompanySettlePrice, rsProtocol, atCompanyInfo, pszCompanyID, szCompanyName, True, bIsDefaut

                    rsProtocol.MoveNext
                    m_bIsNew = True
                Next j


                '将一半拆给与该公司加总平分的公司
                Dim bIsHalveDefault As Boolean
                '取出协议
                Set rsHalveProtocol = GetProtocolFormulaCompany(szHalveCompanyID)
                If rsHalveProtocol.RecordCount = 0 Then
                    '取出默认的协议
                    If rsDefaultProtocol.RecordCount <> 0 Then
                        Set rsHalveProtocol = rsDefaultProtocol
                        bIsHalveDefault = True
                    Else
                        '该公司协议未设置
                        RaiseError ERR_CompanyNotSetProtocol
                    End If
                Else
                    bIsHalveDefault = False
                End If



                m_bIsNew = False
                '找到该线路的起始协议
                FindRouteProtocol rsHalveProtocol, patStationQuantity(i).RouteID, bIsHalveDefault


                '如果公式中没有用到结算价,则不进行查询结算价
                If FindSettlePriceStr(rsHalveProtocol, nVailableNumber, cszSettleStr) Then
                    Set rsCompanySettlePrice = GetCompanySettlePrice(szHalveCompanyID, patStationQuantity(i).VehicleTypeCode, patStationQuantity(i).RouteID, patStationQuantity(i).SellSationID, patStationQuantity(i).StationID)
                End If


                For j = 1 To nVailableNumber
                    SetFormulaValueCompanyManual patStationQuantity(i), aszFormulaValue, nCount, rsCompanySettlePrice, rsHalveProtocol, atCompanyInfo, szHalveCompanyID, szHalveCompanyName, True, bIsHalveDefault
                    rsHalveProtocol.MoveNext
                    m_bIsNew = True
                Next j

            End If

'            dbTotalTicketPrice = dbTotalTicketPrice + FormatDbValue(rsPassageNumber!ticket_price)
'            dbTotalQuantity = dbTotalQuantity + FormatDbValue(rsPassageNumber!Quantity)
            
            dbTotalTicketPrice = 0 'FormatDbValue(rsPassageNumber!ticket_price)
'            dbTotalQuantity = 0 ' FormatDbValue(rsPassageNumber!Quantity)
            tResultInfo.SettleSheetInfo.TotalTicketPrice = tResultInfo.SettleSheetInfo.TotalTicketPrice + dbTotalTicketPrice
            tResultInfo.SettleSheetInfo.TotalQuantity = tResultInfo.SettleSheetInfo.TotalQuantity + dbTotalQuantity
            
            'rsPassageNumber.MoveNext
            m_bIsNew = False
        Next i

    Else

        'else 按车辆结算
        Dim rsVehicleInfo As Recordset
        Dim szLicenseTagNo As String
        szSql = " SELECT c.* , v.license_tag_no FROM " & cszTableVehicle & " v , " & cszTableCompany & " c " _
            & " WHERE v.transport_company_id = c.transport_company_id AND v.vehicle_id = " & TransFieldValueToString(pszVehicleID)
        Set rsVehicleInfo = oDB.Execute(szSql)
        If rsVehicleInfo.RecordCount = 0 Then
            RaiseError ERR_VehicleNotExist
        Else
            szCompanyID = FormatDbValue(rsVehicleInfo!transport_company_id)
            szLicenseTagNo = FormatDbValue(rsVehicleInfo!license_tag_no)
            szCompanyName = FormatDbValue(rsVehicleInfo!transport_company_short_name)
        End If
'        szCompanyID = FormatDbValue(rsPassageNumber!transport_company_id)
        '2.  得到需要结算的车辆所属公司的加总平分公司信息。

        szSql = " SELECT h.*, c.transport_company_short_name FROM fin_halve_company_lst h, " & cszTableCompany & " c " _
            & " WHERE c.transport_company_id = h.transport_company_id AND c.transport_company_id = " & TransFieldValueToString(szCompanyID)

        Set rsHalve = oDB.Execute(szSql)
        '3． 得到该车辆的协议
        Set rsProtocol = GetProtocolFormulaVehicle(pszVehicleID, False)
        bIsDefaut = False
        If rsProtocol.RecordCount = 0 Then
            '取出默认的协议
            If rsDefaultProtocol.RecordCount <> 0 Then
                Set rsProtocol = rsDefaultProtocol
                bIsDefaut = True
            Else
                '该车辆协议未设置
                RaiseError ERR_VehicleNotSetProtocol
            End If
        End If

        '4． For I = 1 to rsPassageNumber.RecordCount
        For i = 1 To ArrayLength(patStationQuantity)

            dbTotalQuantity = patStationQuantity(i).Quantity
            szHalveCompanyID = ""
            For j = 1 To rsHalve.RecordCount
                If patStationQuantity(i).RouteID = FormatDbValue(rsHalve!route_id) And szCompanyID = FormatDbValue(rsHalve!transport_company_id) Then
                    '找到该公司在该线路上需要加总平分
                    szHalveCompanyID = FormatDbValue(rsHalve!other_transport_company_id)
                    szHalveCompanyName = GetHavleOtherCompanyName(szHalveCompanyID)
                    Exit For
                End If
            Next j
            If szHalveCompanyID = "" Or pbIsSingle Then
                '不需要加总平分
                '将相应的人数及各售票票价只算给该车辆。
                rsProtocol.MoveFirst

                '如果公式中没有用到结算价,则不进行查询结算价
                If FindSettlePriceStr(rsProtocol, nVailableNumber, cszSettleStr) Then
                    Set rsVehicleSettlePrice = GetVehicleSettlePrice(pszVehicleID, patStationQuantity(i).RouteID, patStationQuantity(i).SellSationID, patStationQuantity(i).StationID)
                End If
                For j = 1 To nVailableNumber
                    SetFormulaValueVehicleManual patStationQuantity(i), aszFormulaValue, nCount, rsVehicleSettlePrice, rsProtocol, atVehicleInfo, pszVehicleID, szLicenseTagNo, False
                    rsProtocol.MoveNext
                    m_bIsNew = True
                Next j
            Else
                'Else累加该车辆的结算费用明细将这些人数及各售票票价项分为两半，一半算给该车辆，另一半算给加总平分中的另一公司。


                '先结给该车辆，方法与上同。

                '如果公式中没有用到结算价,则不进行查询结算价
                If FindSettlePriceStr(rsProtocol, nVailableNumber, cszSettleStr) Then
                    Set rsVehicleSettlePrice = GetVehicleSettlePrice(pszVehicleID, patStationQuantity(i).RouteID, patStationQuantity(i).SellSationID, patStationQuantity(i).StationID)
                End If

                rsProtocol.MoveFirst
                For j = 1 To nVailableNumber
                    SetFormulaValueVehicleManual patStationQuantity(i), aszFormulaValue, nCount, rsVehicleSettlePrice, rsProtocol, atVehicleInfo, pszVehicleID, szLicenseTagNo, False
                    rsProtocol.MoveNext
                    m_bIsNew = True
                Next j
                '再结给该公司，与按公司结算类似
                '取出协议
                bIsHalveDefault = False
                Set rsHalveProtocol = GetProtocolFormulaCompany(szHalveCompanyID)
                If rsHalveProtocol.RecordCount = 0 Then
                    '取出默认的协议
                    If rsDefaultProtocol.RecordCount <> 0 Then
                        Set rsHalveProtocol = rsDefaultProtocol
                        bIsHalveDefault = True
                    Else
                        '该公司协议未设置
                        RaiseError ERR_VehicleNotSetProtocol
                    End If
                End If


                '如果公式中没有用到结算价,则不进行查询结算价
                If FindSettlePriceStr(rsHalveProtocol, nVailableNumber, cszSettleStr) Then
                    Set rsVehicleSettlePrice = GetVehicleSettlePrice(pszVehicleID, patStationQuantity(i).RouteID, patStationQuantity(i).SellSationID, patStationQuantity(i).StationID)
                End If

                m_bIsNew = False
                rsHalveProtocol.MoveFirst
                FindRouteProtocol rsHalveProtocol, patStationQuantity(i).RouteID, bIsHalveDefault
                For j = 1 To nVailableNumber
                    SetFormulaValueCompanyManual patStationQuantity(i), aszFormulaValue, nCount, rsCompanySettlePrice, rsHalveProtocol, atCompanyInfo, szHalveCompanyID, szHalveCompanyName, True, bIsHalveDefault
                    rsHalveProtocol.MoveNext
                    m_bIsNew = True
                Next j

            End If

            dbTotalTicketPrice = 0 ' FormatDbValue(rsPassageNumber!ticket_price)
'            dbTotalQuantity = 0 'FormatDbValue(rsPassageNumber!Quantity)
            tResultInfo.SettleSheetInfo.TotalTicketPrice = tResultInfo.SettleSheetInfo.TotalTicketPrice + dbTotalTicketPrice
            tResultInfo.SettleSheetInfo.TotalQuantity = tResultInfo.SettleSheetInfo.TotalQuantity + dbTotalQuantity
'            rsPassageNumber.MoveNext
            m_bIsNew = False
        Next i
    End If
    '将总路单数、总票款、总人数、应结票款、结给车站票款以及公司的费用明细和车辆明细均返回
    With tResultInfo.SettleSheetInfo

        .CheckSheetCount = ArrayLength(paszCheckSheetID) '路单数
        For i = 1 To ArrayLength(atCompanyInfo)
            rsProtocol.MoveFirst
'            当结算有底限时 , 现在的处理方式是不对的, 未判断不同公司, 线路不同, 其协议不同, 协议不同时当然其底限也是不同的
'            If bSplitByCompany Then
'               FindRouteProtocol rsProtocol, rsCompanySettlePrice, bIsDefaut
'            End If
            For j = 1 To nVailableNumber 'rsProtocol.RecordCount
                szitem = FormatDbValue(rsProtocol!split_item_id)
                nItem = CInt(szitem)
                dbSplitlimitCharge = GetlimitCharge(FormatDbValue(rsProtocol!protocol_id), szitem, dbSplitUpCharge)
                '如果上限不为0,则   该项值=上限值-该项值
                If dbSplitUpCharge <> 0 Then
                    atCompanyInfo(i).SplitItem(nItem) = dbSplitUpCharge - atCompanyInfo(i).SplitItem(nItem)
                End If
                '当下限值大于该项值时,该值为下限值
                If dbSplitlimitCharge > atCompanyInfo(i).SplitItem(nItem) Then
                    atCompanyInfo(i).SplitItem(nItem) = dbSplitlimitCharge
                Else
                End If
                atCompanyInfo(i).SplitItem(nItem) = FormatMoney(atCompanyInfo(i).SplitItem(nItem))
'                atCompanyInfo(i).SplitItem(nItem) = IIf(dbSplitlimitCharge > atCompanyInfo(i).SplitItem(nItem), dbSplitlimitCharge, atCompanyInfo(i).SplitItem(nItem))
                atCompanyInfo(i).SettlePrice = atCompanyInfo(i).SettlePrice + GetSettlePrice(rsProtocol, atCompanyInfo(i).SplitItem(nItem))
                atCompanyInfo(i).SettleStationPrice = atCompanyInfo(i).SettleStationPrice + GetSettleStationPrice(rsProtocol, atCompanyInfo(i).SplitItem(nItem))
                rsProtocol.MoveNext
            Next j
            atCompanyInfo(i).SettlePrice = FormatMoney(atCompanyInfo(i).SettlePrice) 'Round(atCompanyInfo(i).SettlePrice + 0.01, 1)
            atCompanyInfo(i).SettleStationPrice = FormatMoney(atCompanyInfo(i).SettleStationPrice) 'Round(atCompanyInfo(i).SettleStationPrice + 0.01, 1)

            dbSettleOtherCompanyPrice = FormatMoney(dbSettleOtherCompanyPrice + atCompanyInfo(i).SettlePrice)
            dbSettleStationPrice = FormatMoney(dbSettleStationPrice + atCompanyInfo(i).SettleStationPrice)
        Next i

        For i = 1 To ArrayLength(atVehicleInfo)
            rsProtocol.MoveFirst
            For j = 1 To nVailableNumber 'rsProtocol.RecordCount
                szitem = FormatDbValue(rsProtocol!split_item_id)
                nItem = CInt(szitem)
                dbSplitlimitCharge = GetlimitCharge(FormatDbValue(rsProtocol!protocol_id), szitem, dbSplitUpCharge)
                '如果上限不为0,则   该项值=上限值-该项值
                If dbSplitUpCharge <> 0 Then
                    atVehicleInfo(i).SplitItem(nItem) = dbSplitUpCharge - atVehicleInfo(i).SplitItem(nItem)
                End If
                '当下限值大于该项值时,该值为下限值
                If dbSplitlimitCharge > atVehicleInfo(i).SplitItem(nItem) Then
                    atVehicleInfo(i).SplitItem(nItem) = dbSplitlimitCharge
                Else
                End If
                atVehicleInfo(i).SplitItem(nItem) = FormatMoney(atVehicleInfo(i).SplitItem(nItem))

'                atVehicleInfo(i).SplitItem(nItem) = IIf(dbSplitlimitCharge > atVehicleInfo(i).SplitItem(nItem), dbSplitlimitCharge, atVehicleInfo(i).SplitItem(nItem))
                atVehicleInfo(i).SettlePrice = atVehicleInfo(i).SettlePrice + GetSettlePrice(rsProtocol, atVehicleInfo(i).SplitItem(nItem))
                atVehicleInfo(i).SettleStationPrice = atVehicleInfo(i).SettleStationPrice + GetSettleStationPrice(rsProtocol, atVehicleInfo(i).SplitItem(nItem))
                rsProtocol.MoveNext
            Next j
            atVehicleInfo(i).SettlePrice = atVehicleInfo(i).SettlePrice ' Round(atVehicleInfo(i).SettlePrice + 0.01, 1)
            atVehicleInfo(i).SettleStationPrice = atVehicleInfo(i).SettleStationPrice 'Round(atVehicleInfo(i).SettleStationPrice + 0.01, 1)


            dbSettleOtherCompanyPrice = dbSettleOtherCompanyPrice + atVehicleInfo(i).SettlePrice
            dbSettleStationPrice = dbSettleStationPrice + atVehicleInfo(i).SettleStationPrice
        Next i
        dbSettleOtherCompanyPrice = dbSettleOtherCompanyPrice 'Round(dbSettleOtherCompanyPrice + 0.01, 1)
        dbSettleStationPrice = dbSettleStationPrice 'Round(dbSettleStationPrice + 0.01, 1)

        .SettleOtherCompanyPrice = FormatMoney(dbSettleOtherCompanyPrice) '结算给对方公司金额
        .SettleStationPrice = FormatMoney(dbSettleStationPrice) '结给站方金额
        If pszCompanyID <> "" Then
            .ObjectID = pszCompanyID '对象代码
            .SettleObject = CS_SettleByTransportCompany  '结算对象
        Else
            .ObjectID = pszVehicleID
            .SettleObject = CS_SettleByVehicle
        End If
        .TransportCompanyID = szCompanyID ' FormatDbValue(rsPassageNumber!transport_company_id)
        .TransportCompanyName = szCompanyName ' FormatDbValue(rsPassageNumber!transport_company_short_name)
    
    End With

    tResultInfo.CompanyInfo = atCompanyInfo
    tResultInfo.VehicleInfo = atVehicleInfo
    PreviewSplitCheckSheetManual = tResultInfo

    Exit Function
PreviewSplitCheckSheetManualErr:
    err.Raise err.Number, err.Source, err.Description
End Function



'手工结算时,按参数来统计
Private Sub SetFormulaValueCompanyManual(tStationQuantity As TSettleSheetStation, aszFormulaValue() As String, nCount As Integer, rsCompanySettlePrice As Recordset _
 , rsProtocol As Recordset, atCompanyInfo() As TCompnaySettle, szCompanyID As String, szCompanyName As String, bHalf As Boolean, bDefault As Boolean)
    'bHalf 为TRUE 表示是否只取其中的一半进行拆分
    Dim k As Integer
    Dim dbSplitItem As Double
    Dim l As Integer
    Dim j As Integer '判断是否为同一公司
    Dim i As Integer
    Dim bUseSettlePrice As Boolean '是否使用结算价相关项
    
    aszFormulaValue(1, 2) = tStationQuantity.Quantity '总人数
    aszFormulaValue(2, 2) = 0 '票价总金额
    aszFormulaValue(3, 2) = 0
    
    aszFormulaValue(nCount + 9, 2) = tStationQuantity.AreaRatio '区域费率
    aszFormulaValue(nCount + 10, 2) = 0 '初始化
    
    For k = 4 To nCount + 2
        aszFormulaValue(k, 2) = 0
    Next k
    
    If bHalf Then
        For k = 1 To nCount + 2
            aszFormulaValue(k, 2) = aszFormulaValue(k, 2)
        Next k
    End If
    
    bUseSettlePrice = False
    '如果未使用结算价相关项
    If rsCompanySettlePrice Is Nothing Then
        bUseSettlePrice = True
    ElseIf rsCompanySettlePrice.RecordCount = 0 Then
        bUseSettlePrice = True
    Else
        bUseSettlePrice = False
        '当公式设置与结算价相关时，如果票种为半价或各优惠票时，按照结算价半价来拆。
        If tStationQuantity.TicketType = TP_FullPrice Or tStationQuantity.TicketType = TP_FreeTicket Then
            '如果为全票或免票,则取全价
    
            If bHalf Then
                aszFormulaValue(nCount + 3, 2) = FormatDbValue(rsCompanySettlePrice!halve_full_price)
            Else
                aszFormulaValue(nCount + 3, 2) = FormatDbValue(rsCompanySettlePrice!settle_full_price)
            End If
            aszFormulaValue(nCount + 5, 2) = FormatDbValue(rsCompanySettlePrice!service_full_price)
            '拆算公里
            aszFormulaValue(nCount + 7, 2) = FormatDbValue(rsCompanySettlePrice!Mileage)
            aszFormulaValue(nCount + 8, 2) = FormatDbValue(rsCompanySettlePrice!spring_full_price)
            
        Else
            '否则取半价
            If bHalf Then
                aszFormulaValue(nCount + 3, 2) = FormatDbValue(rsCompanySettlePrice!halve_half_price)
            Else
                aszFormulaValue(nCount + 3, 2) = FormatDbValue(rsCompanySettlePrice!settle_half_price)
            End If
            aszFormulaValue(nCount + 5, 2) = FormatDbValue(rsCompanySettlePrice!service_half_price)
            '拆算公里
            aszFormulaValue(nCount + 7, 2) = FormatDbValue(rsCompanySettlePrice!Mileage) / 2
            aszFormulaValue(nCount + 8, 2) = FormatDbValue(rsCompanySettlePrice!spring_half_price)
        End If
        aszFormulaValue(nCount + 4, 2) = FormatDbValue(rsCompanySettlePrice!pass_charge)
        aszFormulaValue(nCount + 10, 2) = IIf(FormatDbValue(rsCompanySettlePrice!Mileage) >= 250, 1, 0)
    End If
    
    If bUseSettlePrice Then
        aszFormulaValue(nCount + 3, 2) = 0
        aszFormulaValue(nCount + 5, 2) = 0
        aszFormulaValue(nCount + 4, 2) = 0
    End If
'    If bHalf Then
'    Else
'    End If
    If bHalf Then
        '票价公里
        aszFormulaValue(nCount + 6, 2) = 0
    Else
        '票价公里
        aszFormulaValue(nCount + 6, 2) = 0
    End If
    
    If tStationQuantity.TicketType = TP_FullPrice Or tStationQuantity.TicketType = TP_FreeTicket Then
        
    Else
        '半票的[票价]人公里为一半
        aszFormulaValue(nCount + 6, 2) = 0
    End If
    dbSplitItem = CaleLugValue(FormatDbValue(rsProtocol!formula_comment), aszFormulaValue) '根据协议得出相应的款额
    
    
    If bUseSettlePrice And bHalf Then
        '如果未用到结算价,且使用加总平分
        dbSplitItem = Round(dbSplitItem / 2 + 0.0001, 3)
    Else
        dbSplitItem = Round(dbSplitItem + 0.0001, 3)
    End If
    For l = 1 To ArrayLength(atCompanyInfo)
        If atCompanyInfo(l).CompanyID = szCompanyID Then
            '如果该类型中已经有该公司
            Exit For
        End If
    Next l
    If l > ArrayLength(atCompanyInfo) Then
        If l = 1 Then
            ReDim atCompanyInfo(1 To l)
        Else
            ReDim Preserve atCompanyInfo(1 To l)
        End If
        atCompanyInfo(l).CompanyID = szCompanyID
        atCompanyInfo(l).CompanyName = szCompanyName
        atCompanyInfo(l).ProtocolID = FormatDbValue(rsProtocol!protocol_id)
        atCompanyInfo(l).ProtocolName = FormatDbValue(rsProtocol!protocol_name)
        
    End If
'     dblimitCharge = GetlimitCharge(FormatDbValue(rsProtocol!protocol_id), FormatDbValue(rsProtocol!split_item_id))
'     If dblimitCharge > dbSplitItem Then
'        dbSplitItem = dblimitCharge
'     Else
'        dbSplitItem = dbSplitItem
'     End If
    atCompanyInfo(l).SplitItem(CInt(rsProtocol!split_item_id)) = atCompanyInfo(l).SplitItem(CInt(rsProtocol!split_item_id)) + dbSplitItem
'    atCompanyInfo(l).SettlePrice = atCompanyInfo(l).SettlePrice + GetSettlePrice(rsProtocol, dbSplitItem)
'    atCompanyInfo(l).SettleStationPrice = atCompanyInfo(l).SettleStationPrice + GetSettleStationPrice(rsProtocol, dbSplitItem)
    If m_bIsNew = False Then
        atCompanyInfo(l).PassengerNumber = atCompanyInfo(l).PassengerNumber + aszFormulaValue(1, 2)  '人数
        atCompanyInfo(l).Mileage = atCompanyInfo(l).Mileage + aszFormulaValue(nCount + 7, 2) * aszFormulaValue(1, 2) '里程
        
    End If
    
End Sub




Private Sub SetFormulaValueVehicleManual(tStationQuantity As TSettleSheetStation, aszFormulaValue() As String, nCount As Integer, rsVehicleSettlePrice As Recordset _
 , rsProtocol As Recordset, atVehicleInfo() As TVehilceSettle, szVehicleID As String, szLicenseTagNo As String, bHalf As Boolean)
    'bHalf 为TRUE 表示是否只取其中的一半进行拆分
    Dim k As Integer
    Dim dbSplitItem As Double
    Dim l As Integer
    Dim i As Integer
    
    
    Dim bUseSettlePrice As Boolean '是否使用结算价相关项
    
    
    aszFormulaValue(1, 2) = tStationQuantity.Quantity ' FormatDbValue(rsPassageNumber!Quantity) '总人数
    aszFormulaValue(2, 2) = 0 'FormatDbValue(rsPassageNumber!ticket_price) '票价总金额
    aszFormulaValue(3, 2) = 0 'FormatDbValue(rsPassageNumber!base_carriage)
    
    
    aszFormulaValue(nCount + 9, 2) = tStationQuantity.AreaRatio '区域费率
    aszFormulaValue(nCount + 10, 2) = 0 '初始化
    
    For k = 4 To nCount + 2
        aszFormulaValue(k, 2) = 0 ' FormatDbValue(rsPassageNumber("price_item_" & CInt(k - 3)))
    Next k
    
    If bHalf Then
        For k = 1 To nCount + 2
            aszFormulaValue(k, 2) = aszFormulaValue(k, 2)
        Next k
    End If
    bUseSettlePrice = False
    If rsVehicleSettlePrice Is Nothing Then
        bUseSettlePrice = True
'        aszFormulaValue(nCount + 3, 2) = 0
'        aszFormulaValue(nCount + 5, 2) = 0
'        aszFormulaValue(nCount + 4, 2) = 0
    ElseIf rsVehicleSettlePrice.RecordCount = 0 Then
        bUseSettlePrice = True
'        aszFormulaValue(nCount + 3, 2) = 0
'        aszFormulaValue(nCount + 5, 2) = 0
'        aszFormulaValue(nCount + 4, 2) = 0
    Else
        bUseSettlePrice = False

        '当公式设置与结算价相关时，如果票种为半价或各优惠票时，按照结算价半价来拆。
        If tStationQuantity.TicketType = TP_FullPrice Or tStationQuantity.TicketType = TP_FreeTicket Then
            '如果为全票或免票,则取全价
            If bHalf Then
                aszFormulaValue(nCount + 3, 2) = FormatDbValue(rsVehicleSettlePrice!halve_full_price)
            Else
                aszFormulaValue(nCount + 3, 2) = FormatDbValue(rsVehicleSettlePrice!settle_full_price)
            End If
            aszFormulaValue(nCount + 5, 2) = FormatDbValue(rsVehicleSettlePrice!service_full_price)
            '拆算公里
            aszFormulaValue(nCount + 7, 2) = FormatDbValue(rsVehicleSettlePrice!Mileage)
            aszFormulaValue(nCount + 8, 2) = FormatDbValue(rsVehicleSettlePrice!spring_full_price)
        Else
            '否则取半价
            If bHalf Then
                aszFormulaValue(nCount + 3, 2) = FormatDbValue(rsVehicleSettlePrice!halve_half_price)
            Else
                aszFormulaValue(nCount + 3, 2) = FormatDbValue(rsVehicleSettlePrice!settle_half_price)
            End If
            aszFormulaValue(nCount + 5, 2) = FormatDbValue(rsVehicleSettlePrice!service_half_price)
            '拆算公里
            aszFormulaValue(nCount + 7, 2) = FormatDbValue(rsVehicleSettlePrice!Mileage) / 2
            aszFormulaValue(nCount + 8, 2) = FormatDbValue(rsVehicleSettlePrice!spring_half_price)
        End If
        aszFormulaValue(nCount + 4, 2) = FormatDbValue(rsVehicleSettlePrice!pass_charge)
        aszFormulaValue(nCount + 10, 2) = IIf(FormatDbValue(rsVehicleSettlePrice!Mileage) >= 250, 1, 0)
    End If
    If bUseSettlePrice Then
        aszFormulaValue(nCount + 3, 2) = 0
        aszFormulaValue(nCount + 5, 2) = 0
        aszFormulaValue(nCount + 4, 2) = 0
    End If
    
    
    If bHalf Then
        '票价公里
        aszFormulaValue(nCount + 6, 2) = 0 ' FormatDbValue(rsPassageNumber!Mileage) / 2
    Else
        '票价公里
        aszFormulaValue(nCount + 6, 2) = 0 'FormatDbValue(rsPassageNumber!Mileage)
    End If
    If tStationQuantity.TicketType = TP_FullPrice Or tStationQuantity.TicketType = TP_FreeTicket Then
        
    Else
        '半票的[票价]人公里为一半
        aszFormulaValue(nCount + 6, 2) = aszFormulaValue(nCount + 6, 2) / 2
    End If
'    dbSplitItem = Round(CaleLugValue(FormatDbValue(rsProtocol!formula_comment), aszFormulaValue) + 0.001, 2) '跟据协议得出相应的款额


    dbSplitItem = CaleLugValue(FormatDbValue(rsProtocol!formula_comment), aszFormulaValue) '根据协议得出相应的款额
    
    
    If bUseSettlePrice And bHalf Then
        '如果未用到结算价,且使用加总平分
        dbSplitItem = Round(dbSplitItem / 2 + 0.0001, 3)
    Else
        dbSplitItem = Round(dbSplitItem + 0.0001, 3)
    End If



    For l = 1 To ArrayLength(atVehicleInfo)
        If atVehicleInfo(l).VehicleID = szVehicleID Then
            '如果该类型中已经有该公司
            Exit For
        End If
    Next l
    If l > ArrayLength(atVehicleInfo) Then
        If l = 1 Then
            ReDim atVehicleInfo(1 To l)
        Else
            ReDim Preserve atVehicleInfo(1 To l)
        End If
        atVehicleInfo(l).VehicleID = szVehicleID
        atVehicleInfo(l).LicenseTagNo = szLicenseTagNo
        atVehicleInfo(l).ProtocolID = FormatDbValue(rsProtocol!protocol_id)
        atVehicleInfo(l).ProtocolName = FormatDbValue(rsProtocol!protocol_name)
        
    End If
    
'    dblimitCharge = GetlimitCharge(FormatDbValue(rsProtocol!protocol_id), FormatDbValue(rsProtocol!split_item_id))
'     If dblimitCharge > dbSplitItem Then
'        dbSplitItem = dblimitCharge
'     Else
'        dbSplitItem = dbSplitItem
'     End If
    atVehicleInfo(l).SplitItem(CInt(rsProtocol!split_item_id)) = atVehicleInfo(l).SplitItem(CInt(rsProtocol!split_item_id)) + dbSplitItem
'    atVehicleInfo(l).SettlePrice = atVehicleInfo(1).SettlePrice + GetSettlePrice(rsProtocol, dbSplitItem)
'    atVehicleInfo(l).SettleStationPrice = atVehicleInfo(l).SettleStationPrice + GetSettleStationPrice(rsProtocol, dbSplitItem)
    If m_bIsNew = False Then
        atVehicleInfo(l).PassengerNumber = atVehicleInfo(l).PassengerNumber + aszFormulaValue(1, 2) '人数
        atVehicleInfo(l).Mileage = atVehicleInfo(l).Mileage + aszFormulaValue(nCount + 7, 2) * aszFormulaValue(1, 2) '里程
    End If
End Sub



'手工结算,站点人数信息按照参数来插入
'##ModelId=3E81F4790196
Public Sub SplitCheckSheetManual(pszSettleSheetID As String, pszCompanyID As String, pszVehicleID As String, paszCheckSheetID() As String _
        , ptResultInfo As TSplitResult, pdyStartDate As Date, pdyEndDate As Date)
        ', pszLuggageIDs As String, pdbTotalLuggageBaseCarriage As Double         , pdbTotalLuggageSettlePrice As Double)
        
        
        
    On Error GoTo SplitCheckSheetErr
    
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    
    Dim szWhere As String
    Dim nSheetCount As Integer
    Dim i As Integer
    Dim j As Integer
    
    Dim bSplitByCompany As Boolean
    Dim dyNow As String
    '验证用户是否合法
    '验证对象是否合法
    '验证用户是否有结算的权限
    
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_SplitCheckSheet '验证用户是否有结算的权限
    
    oDB.ConnectionString = GetConnectionStr()
    
    If pszCompanyID = "" And pszVehicleID = "" Then
        '出传入参数错
        RaiseError ERR_PreviewSplitInputParam
    ElseIf pszCompanyID <> "" Then
        bSplitByCompany = True
    Else
        bSplitByCompany = False
    End If
    
    
'    nSheetCount = ArrayLength(paszCheckSheetID)
'    If nSheetCount = 0 Then RaiseError ERR_PreviewSplitInputParam '出传入参数错
    
'    szWhere = " ( "
'    For i = 1 To nSheetCount - 1
'            szWhere = szWhere & "'" & paszCheckSheetID(i) & "',"
'    Next i
'    szWhere = szWhere & "'" & paszCheckSheetID(i) & "')"
'
'
'    '判断每一张路单是否均存在，如不存在则出"路单不存在"的错。如无效或经过结算。则出"路单无效或路单已经结算"的错。
'    szSql = "SELECT * FROM " & cszTableCheckSheet & "  WHERE check_sheet_id IN " & szWhere
'    Set rsTemp = oDB.Execute(szSql)
'    If rsTemp.RecordCount <> nSheetCount Then
'        '如果记录数不为传入的路单数,则出错
'        RaiseError ERR_SomeSheetNotExist
'    Else
        '判断所有的路单是不是同一公司或同一车辆
        
'        For i = 1 To rsTemp.RecordCount
'            If bSplitByCompany Then
'                If FormatDbValue(rsTemp!transport_company_id) <> pszCompanyID Then
'                    '出错
'                    RaiseError ERR_SplitSheetMustSameCompany
'                End If
'            Else
'                If ucase(FormatDbValue(rsTemp!vehicle_id)) <> ucase(pszVehicleID) Then
'                    '出错
'                    RaiseError ERR_SplitSheetMustSameVehicle
'                End If
'            End If
'            rsTemp.MoveNext
'        Next i
'    End If
    
'    szSql = "SELECT * FROM " & cszTableCheckSheet & " WHERE (settlement_status = " & TransFieldValueToString(CS_CheckSheetSettled) _
'        & " OR valid_mark = " & TransFieldValueToString(CS_CheckSheetInvalid) & " )" _
'        & " AND check_sheet_id IN " & szWhere
'
'    Set rsTemp = odb.Execute(szSql)
'    If rsTemp.RecordCount <> 0 Then
'        '如果存在已废或已结算的路单,则出错
'        RaiseError ERR_SomeSheetHasCancelORSettle
'    End If

    szSql = " SELECT * FROM fin_settle_sheet_lst WHERE settlement_sheet_id = " & TransFieldValueToString(pszSettleSheetID)
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount <> 0 Then
        RaiseError ERR_SettleSheetHasExist
    End If
    
    
'    Dim szObjectName As String
    
'    If ptResultInfo.SettleSheetInfo.SettleObject = CS_SettleByTransportCompany Then
'        szSql = " SELECT transport_company_short_name object_name FROM " & cszTableCompany & " WHERE transport_company_id = '" & ptResultInfo.SettleSheetInfo.ObjectID & "'"
'    ElseIf ptResultInfo.SettleSheetInfo.SettleObject = CS_SettleByVehicle Then
'        szSql = "SELECT license_tag_no object_name FROM " & cszTableVehicle & " WHERE vehicle_id = '" & ptResultInfo.SettleSheetInfo.ObjectID & "'"
'    End If
'
'    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount > 0 Then szObjectName = FormatDbValue(rsTemp!object_name)
'
    
    oDB.BeginTrans

        dyNow = ToDBDateTime(SelfNowDateTime())
        '开始事务
        '根据公司结算明细及车辆结算明细，累加出总人数、总应结    票款、总应结给车站票款
        '汇总出所有这些路单的总的售票票款、总路单数
        '写入"路单结算单表"，拆帐人写入用户代码，拆帐日期写当前日期，起始日期与结束日期用传入的参数写入
        
        szSql = " INSERT fin_settle_sheet_lst ( " _
            & " settlement_sheet_id , " _
            & " sheet_quantity ," _
            & " total_ticket_price , " _
            & " total_quantity , " _
            & " settle_price , " _
            & " settle_station_price , " _
            & " object_id , " _
            & " object_name , " _
            & " settle_object , " _
            & " transport_company_id , " _
            & " transport_company_name , " _
            & " settler , " _
            & " checker , " _
            & " settle_date , " _
            & " start_date , " _
            & " end_date , " _
            & " unit_name , " _
            & " status , " _
            & " luggage_settle_id , " _
            & " luggage_base_carriage , " _
            & " luggage_settle_price , " _
            & " luggage_protocol_name , " _
            & " annotation) "
        szSql = szSql & " VALUES (" _
            & TransFieldValueToString(pszSettleSheetID) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.CheckSheetCount) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.TotalTicketPrice) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.TotalQuantity) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.SettleOtherCompanyPrice) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.SettleStationPrice) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.ObjectID) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.ObjectName) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.SettleObject) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.TransportCompanyID) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.TransportCompanyName) & "," _
            & TransFieldValueToString(m_oActiveUser.UserName) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.Checker) & "," _
            & TransFieldValueToString(dyNow) & "," _
            & TransFieldValueToString(pdyStartDate) & "," _
            & TransFieldValueToString(pdyEndDate) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.UnitName) & "," _
            & TransFieldValueToString(CS_SettleSheetValid) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.LuggageSettleIDs) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.LuggageTotalBaseCarriage) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.LuggageTotalSettlePrice) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.LuggageProtocolName) & "," _
            & TransFieldValueToString(ptResultInfo.SettleSheetInfo.Annotation) & ")"
        
        oDB.Execute szSql
        
        
        '修改每一张路单的结算单号、拆算协议、结算时间、结算标记。
        
        '拆算协议这个字段不写入
        '因为一个协议不一定是只拆给一下公司的，或一个车辆的

'        szSql = " UPDATE " & cszTableCheckSheet & " SET settlement_sheet_id = " & TransFieldValueToString(pszSettleSheetID) _
'            & " , settlement_time = " & TransFieldValueToString(dyNow) _
'            & " , settlement_status = " & TransFieldValueToString(CS_CheckSheetSettled) _
'            & " WHERE check_sheet_id IN " & szWhere
'
'        oDB.Execute szSql
        
        
        
'        将汇总的站点人数信息写入 "拆帐结算单站点表"
        For i = 1 To ArrayLength(ptResultInfo.SheetStationInfo)
            szSql = "INSERT fin_settle_station_lst( " _
                & " settlement_sheet_id, " _
                & " sell_station_id," _
                & " route_id," _
                & " route_name , " _
                & " station_id, " _
                & " ticket_type_id," _
                & " quantity, " _
                & " bus_id, " _
                & " vehicle_type_code ," _
                & " vehicle_id," _
                & " transport_company_id) "
            szSql = szSql & " VALUES ( " _
                & TransFieldValueToString(pszSettleSheetID) & " , " _
                & TransFieldValueToString(ptResultInfo.SheetStationInfo(i).SellSationID) & " , " _
                & TransFieldValueToString(ptResultInfo.SheetStationInfo(i).RouteID) & " , " _
                & TransFieldValueToString(ptResultInfo.SheetStationInfo(i).RouteName) & " , " _
                & TransFieldValueToString(ptResultInfo.SheetStationInfo(i).StationID) & " , " _
                & TransFieldValueToString(ptResultInfo.SheetStationInfo(i).TicketType) & " , " _
                & TransFieldValueToString(ptResultInfo.SheetStationInfo(i).Quantity) & " , " _
                & TransFieldValueToString("") & " , " _
                & TransFieldValueToString(ptResultInfo.SheetStationInfo(i).VehicleTypeCode) & " , " _
                & TransFieldValueToString(IIf(bSplitByCompany, "", pszVehicleID)) & " , " _
                & TransFieldValueToString(IIf(bSplitByCompany, pszCompanyID, "")) & " ) "
                
            oDB.Execute szSql
        Next i

            
            
            
            '& " SELECT " & TransFieldValueToString(pszSettleSheetID) & ",(f.sell_station_id) sell_station_id,(s.route_id) route_id, MAX(r.route_name) route_name " _
            & "     , (f.station_id) station_id,(f.ticket_type) ticket_type,SUM(f.quantity) quantity,MAX(s.bus_id) bus_id" _
            & "     ,MAX(s.vehicle_type_code) vehicle_type_code,MAX(s.vehicle_id) vehicle_id,MAX(s.transport_company_id) transport_company_id" _
            & " FROM " & cszTableSheetStation & " f, " & cszTableCheckSheet & " s , " & cszTableRouteInfo & " r " _
            & " WHERE s.route_id = r.route_id AND s.check_sheet_id=f.check_sheet_id AND s.check_sheet_id IN" & szWhere _
            & " GROUP BY f.sell_station_id , s.route_id , f.station_id , f.ticket_type "
'        oDb.Execute szSql
        
        '写入"结算单公司拆算明细表"
        For i = 1 To ArrayLength(ptResultInfo.CompanyInfo)
            szSql = " INSERT fin_settle_company_price_lst ( " _
                & " settlement_sheet_id , " _
                & " transport_company_id , " _
                & " protocol_id, " _
                & " settle_price, " _
                & " settle_station_price, " _
                & " mileage , " _
                & " passenger_number, "
            For j = 1 To 19
                szSql = szSql & "split_item_" & CInt(j) & " , "
            Next j
            szSql = szSql & "split_item_20 ) "
            szSql = szSql & " VALUES ( " _
                & TransFieldValueToString(pszSettleSheetID) & " , " _
                & TransFieldValueToString(ptResultInfo.CompanyInfo(i).CompanyID) & " , " _
                & TransFieldValueToString(ptResultInfo.CompanyInfo(i).ProtocolID) & " , " _
                & TransFieldValueToString(ptResultInfo.CompanyInfo(i).SettlePrice) & " , " _
                & TransFieldValueToString(ptResultInfo.CompanyInfo(i).SettleStationPrice) & " , " _
                & TransFieldValueToString(ptResultInfo.CompanyInfo(i).Mileage) & "," _
                & TransFieldValueToString(ptResultInfo.CompanyInfo(i).PassengerNumber) & " , "
                
                
            For j = 1 To 19
                szSql = szSql & TransFieldValueToString(ptResultInfo.CompanyInfo(i).SplitItem(j)) & " , "
            Next j
            szSql = szSql & TransFieldValueToString(ptResultInfo.CompanyInfo(i).SplitItem(20)) & " ) "
            oDB.Execute szSql
            
        Next i
        
        '写入"结算单车辆拆算明细表"
        For i = 1 To ArrayLength(ptResultInfo.VehicleInfo)
            With ptResultInfo.VehicleInfo(i)
                szSql = " INSERT fin_settle_vehicle_lst ( " _
                    & " settlement_sheet_id , " _
                    & " vehicle_id , " _
                    & " protocol_id , " _
                    & " settle_price , " _
                    & " settle_station_price , " _
                    & " mileage , " _
                    & " passenger_number , "
                For j = 1 To 19
                    szSql = szSql & "split_item_" & CInt(j) & " , "
                Next j
                szSql = szSql & "split_item_20 ) "
                szSql = szSql & " VALUES ( " _
                    & TransFieldValueToString(pszSettleSheetID) & " , " _
                    & TransFieldValueToString(.VehicleID) & " , " _
                    & TransFieldValueToString(.ProtocolID) & " , " _
                    & TransFieldValueToString(.SettlePrice) & " , " _
                    & TransFieldValueToString(.SettleStationPrice) & " , " _
                    & TransFieldValueToString(.Mileage) & " , " _
                    & TransFieldValueToString(.PassengerNumber) & " , "
                    
                For j = 1 To 19
                    szSql = szSql & TransFieldValueToString(.SplitItem(j)) & " , "
                Next j
            
                szSql = szSql & TransFieldValueToString(.SplitItem(20)) & " ) "
                
                oDB.Execute szSql
            End With
        Next i
    '结束事务
    oDB.CommitTrans
    Dim szlog As String
    szlog = "结算单号为" & pszSettleSheetID & ","
    If bSplitByCompany Then
        szlog = szlog & "按公司结算,结算对象为" & pszCompanyID & ","
    Else
        szlog = szlog & "按车辆结算,结算对象为" & pszVehicleID & ","
    End If
    szlog = szlog & "结算日期段为" & ToDBDate(pdyStartDate) & "-" & ToDBDate(pdyEndDate)
    '写入日志
    WriteOperateLog m_oActiveUser, RIGHT_SplitCheckSheet, szlog

    Exit Sub
SplitCheckSheetErr:
    oDB.RollbackTrans
    Call err.Raise(err.Number, "Split:SplitCheckSheet Method", err.Description)
End Sub


'检票时,更改车次车辆
Public Sub ChangeSheetVehicle(pszSheetID As String, pszVehicle As String)
    '更改车辆信息
    Dim oDB As New RTConnection
    Dim szSql As String
    
    Dim szVehicleTypeCode As String
    Dim szOwnerID As String
    Dim szTransportCompanyID As String
    Dim szVehicleTypeName As String
    Dim szSplitCompanyID As String
    
    Dim rsVehicle As Recordset

    On Error GoTo ErrorHandle
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_ModifySheetVehicle
    
    
    oDB.ConnectionString = GetConnectionStr
    szSql = "SELECT v.*,t.vehicle_type_short_name" & _
            " FROM Vehicle_info v,vehicle_type_code t" & _
            " WHERE v.vehicle_id='" & pszVehicle & "' AND v.vehicle_type_code=t.vehicle_type_code"
    Set rsVehicle = oDB.Execute(szSql)
    If rsVehicle.RecordCount = 0 Then err.Raise 33333, , "车辆不存在"
    
    
    szVehicleTypeCode = FormatDbValue(rsVehicle!vehicle_type_code)
    szOwnerID = FormatDbValue(rsVehicle!owner_id)
    szTransportCompanyID = FormatDbValue(rsVehicle!transport_company_id)
    szVehicleTypeName = FormatDbValue(rsVehicle!vehicle_type_short_name)
    szSplitCompanyID = FormatDbValue(rsVehicle!split_company_id)
    
    '开始修改
'    oDb.BeginTrans
'        'Bus_check_lst
'        '修改检票车次表
'        szSql = " UPDATE bus_check_lst " & _
'            " SET vehicle_id='" & pszVehicle & "'" & _
'            "    ,transport_company_id='" & szTransportCompanyID & "'" & _
'            "    ,owner_id='" & szOwnerID & "'" & _
'            "    ,split_company_id='" & szSplitCompanyID & "'" & _
'            "    ,vehicle_type_code='" & szVehicleTypeCode & "'" & _
'            " WHERE bus_id='" & pszBusID & "' AND " & DBDateCompare(pdyBusDate, "bus_date") _
'            & " AND bus_serial_no = '" & pnBusSerialNo & "'" & " AND check_gate_id = '" & m_sCheckGateNo & "'"
'             'SET vehicle_id = , transport_company_id = , vehicle_type_code = , owner_id = , split_company_id =  "
'        oDb.Execute szSql
'
'
'        '环境车次表
'        szSql = "UPDATE Env_bus_info" & _
'             " SET vehicle_id='" & pszVehicle & "'" & _
'             "    ,transport_company_id='" & szTransportCompanyID & "'" & _
'             "    ,owner_id='" & szOwnerID & "'" & _
'             "    ,split_company_id='" & szSplitCompanyID & "'" & _
'             "    ,vehicle_type_code='" & szVehicleTypeCode & "'" & _
'             "    ,vehicle_type_name='" & szVehicleTypeName & "'" & _
'             " WHERE bus_id='" & pszBusID & "' AND " & DBDateCompare(pdyBusDate, "bus_date")
'    oDb.Execute szSql '固定车次回写环境车次表状


        '只修改路单表 ，先暂时其他表都不修改
        szSql = "UPDATE " & cszTableCheckSheet & _
             " SET vehicle_id='" & pszVehicle & "'" & _
             "    ,transport_company_id='" & szTransportCompanyID & "'" & _
             "    ,owner_id='" & szOwnerID & "'" & _
             "    ,split_company_id='" & szSplitCompanyID & "'" & _
             "    ,vehicle_type_code='" & szVehicleTypeCode & "'" & _
             " WHERE check_sheet_id ='" & pszSheetID & "'"
        oDB.Execute szSql
        WriteOperateLog m_oActiveUser, RIGHT_ModifySheetVehicle, "修改路单的所属车辆改为:" & pszVehicle
'    oDb.CommitTrans
    Exit Sub
ErrorHandle:
'    oDb.RollbackTrans
    err.Raise err.Number, err.Source, err.Description
    
End Sub


Public Function GetCheckSheetStationList(paszCheckSheetID() As String, Optional pbIsToday As Boolean = False) As Recordset
    '得到路单站点明细
    
    On Error GoTo ErrorHandle
    
'    Dim atStationInfo() As TSettleSheetStation
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim szWhere As String
    Dim nSheetCount As Integer
    Dim i As Integer
    
    AssertObjIsValid
    
    
    oDB.ConnectionString = GetConnectionStr()
    
    nSheetCount = ArrayLength(paszCheckSheetID)
    
    szWhere = " ( "
    For i = 1 To nSheetCount - 1
            szWhere = szWhere & "'" & paszCheckSheetID(i) & "',"
    Next i
    szWhere = szWhere & "'" & paszCheckSheetID(i) & "')"
    
    
    
    '此处的annotation,是为了调用PreviewSplitCheckSheetEx接口,而留 的.

    szSql = " SELECT s.* , i.sell_station_name , ticket_type_name , seat_type_name , c.route_id , c.vehicle_type_code , i.station_id sell_station_id2 , r.route_name " _
    & " , a.annotation " _
    & " FROM " & IIf(Not pbIsToday, cszTableSheetStation, cszTableSheetStationTemp) & " s , sell_station_info i ,  ticket_type_code y , seat_type_code d , sheet_lst c , station_info t , area_code a , route_info r " _
    & " WHERE s.sell_station_id = i.sell_station_id AND s.ticket_type = y.ticket_type_id " _
    & " AND s.seat_type_id = d.seat_type_id and s.check_sheet_id = c.check_sheet_id AND s.station_id = t.station_id AND t.area_code = a.area_code  AND r.route_id = c.route_id  " _
    & " AND s.check_sheet_id IN " & szWhere _
    & " ORDER BY s.bus_date , s.bus_id , s.sell_station_id , s.station_id , s.ticket_type , price_identify "
    
    Set rsTemp = oDB.Execute(szSql)
    
    Set GetCheckSheetStationList = rsTemp
    
    Exit Function
ErrorHandle:
    Call err.Raise(err.Number, "Split:GetCheckSheetStationList Method", err.Description)

End Function

Public Function GetLicenseTagNo(paszCheckSheetID() As String, Optional pbIsToday As Boolean = False) As Recordset
    '得到结算车辆
    
    On Error GoTo ErrorHandle
    
'    Dim atStationInfo() As TSettleSheetStation
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim szWhere As String
    Dim nSheetCount As Integer
    Dim i As Integer
    
    AssertObjIsValid
    
    
    oDB.ConnectionString = GetConnectionStr()
    
    nSheetCount = ArrayLength(paszCheckSheetID)
    
    szWhere = " ( "
    For i = 1 To nSheetCount - 1
            szWhere = szWhere & "'" & paszCheckSheetID(i) & "',"
    Next i
    szWhere = szWhere & "'" & paszCheckSheetID(i) & "')"

    szSql = " SELECT license_tag_no " _
            & " FROM " & IIf(Not pbIsToday, cszTableSheetStation, cszTableSheetStationTemp) & " s , vehicle_info v , sheet_lst c " _
            & " WHERE c.vehicle_id=v.vehicle_id " _
            & " AND s.check_sheet_id = c.check_sheet_id " _
            & " AND s.check_sheet_id IN " & szWhere _
            & " group BY license_tag_no "
    
    Set rsTemp = oDB.Execute(szSql)
    
    Set GetLicenseTagNo = rsTemp
    
    Exit Function
ErrorHandle:
    Call err.Raise(err.Number, "Split:GetLicenseTagNo Method", err.Description)

End Function


'新增车辆固定费用
Public Sub AddVehicleFixFee(prsFixFee As Recordset)
    
    On Error GoTo ErrorHandle
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim i As Integer
    Dim j As Integer
    
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_VehicleFixFee '验证用户是否有车辆固定费用的权限
    
    oDB.ConnectionString = GetConnectionStr()
    
    If prsFixFee.RecordCount = 0 Then Exit Sub
    prsFixFee.MoveFirst
    For i = 1 To prsFixFee.RecordCount
    
        
        szSql = "SELECT 1 FROM fin_vehicle_fix_fee_lst  " _
            & " WHERE vehicle_id = " & TransFieldValueToString(prsFixFee!vehicle_id) _
            & " AND bus_date = " & TransFieldValueToString(prsFixFee!Bus_date)
        Set rsTemp = oDB.Execute(szSql)
        If rsTemp.RecordCount > 0 Then
            '出车辆固定费用已存在的错
            RaiseError ERR_VehicleFixFeeHasExist
            
        End If
        '如果该车辆当月的信息已经输入过一次,则出错
        
        szSql = "INSERT fin_vehicle_fix_fee_lst  ( " _
            & " vehicle_id , " _
            & " bus_date , "
        For j = 1 To cnFixFeeCount - 1
            szSql = szSql & "split_item_" & j & " , "
        Next j
        szSql = szSql & "split_item_" & j & " , " _
            & " is_dec ) VALUES ( " _
        
        szSql = szSql & TransFieldValueToString(prsFixFee!vehicle_id) & " , " _
            & TransFieldValueToString(prsFixFee!Bus_date) & " , " _
            
        For j = 1 To cnFixFeeCount - 1
            
            szSql = szSql & TransFieldValueToString(FormatDbValue(prsFixFee.Fields("split_item_" & j))) & " , "
        Next j
        szSql = szSql & TransFieldValueToString(FormatDbValue(prsFixFee.Fields("split_item_" & j))) & " , " _
            & TransFieldValueToString(False) & " ) "
        
        oDB.Execute szSql
        
        prsFixFee.MoveNext
    Next i
    
    Exit Sub
ErrorHandle:
    err.Raise err.Number, err.Source
    
End Sub

'修改车辆固定费用
Public Sub EditVehicleFixFee(prsFixFee As Recordset)
    On Error GoTo ErrorHandle
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim i As Integer
    Dim lEffect As Long
    Dim oField As Field
    Dim szField As String
    Const cszProcName = "EditVehicleFixFee"
    
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_VehicleFixFee '验证用户是否有车辆固定费用的权限
    
    oDB.ConnectionString = GetConnectionStr()
    
    If prsFixFee.RecordCount = 0 Then Exit Sub
    
    szSql = "SELECT * FROM fin_vehicle_fix_fee_lst  " _
        & " WHERE vehicle_id = " & TransFieldValueToString(prsFixFee!vehicle_id) _
        & " AND bus_date = " & TransFieldValueToString(prsFixFee!Bus_date)
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then
        '出车辆固定费用已存在的错
        RaiseError ERR_VehicleFixFeeNotExist
    Else
        '如果费用已经扣过就不允许修改或删除
        
        If TransFieldValueToString(FormatDbValue(rsTemp!is_dec)) = TransFieldValueToString(True) Then
            RaiseError ERR_VehicleFixFeeHasDesCannotEdit
        End If
    End If
    
    For Each oField In prsFixFee.Fields
        Select Case LCase(oField.Name)
            Case "vehicle_id", "is_dec" '忽略更改
            Case "bus_date", "split_item_1", "split_item_2", "split_item_3", "split_item_4", "split_item_5", "split_item_6", "split_item_7" _
                , "split_item_8", "split_item_9", "split_item_10", "split_item_11", "split_item_12" _
                , "split_item_13", "split_item_14", "split_item_15", "split_item_16", "split_item_17" _
                , "split_item_18", "split_item_19", "split_item_20"
                
                
                
                szField = szField & "," & oField.Name & "=" & TransFieldValueToString(oField.Value)
            Case Else
                RaiseError "列名不正确", cszProcName
        End Select
        
    Next
    oDB.BeginTrans
        Do While Not prsFixFee.EOF
            szSql = "UPDATE fin_vehicle_fix_fee_lst SET vehicle_id=vehicle_id " _
                & szField
                
            szSql = szSql & " WHERE vehicle_id = " & TransFieldValueToString(prsFixFee!vehicle_id) _
                & " AND bus_date = " & TransFieldValueToString(prsFixFee!Bus_date)
                
            oDB.Execute szSql
'            If lEffect = 0 Then
'                '出记录未找到的错
'
'                RaiseError ERR_VehicleFixFeeHasExist
'            End If
            prsFixFee.MoveNext
        Loop
    oDB.CommitTrans
    Set oDB = Nothing
'    Set rsTemp = Nothing

    
    
    Exit Sub
UpdDriverErr:
    oDB.RollbackTrans
    Set oDB = Nothing
'    Set rsTemp = Nothing
    
    Call RaiseError(err.Number, "split:EditVehicleFixFee Method", err.Description)
    
    
    
    Exit Sub
ErrorHandle:
    err.Raise err.Number, err.Source
End Sub

'删除车辆固定费用
Public Sub DelVehicleFixFee(pszVehicleID As String, pdtDate As Date)
    On Error GoTo ErrorHandle
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim i As Integer
    Dim lEffect As Long
    
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_VehicleFixFee '验证用户是否有车辆固定费用的权限
    
    oDB.ConnectionString = GetConnectionStr()
    szSql = "SELECT * FROM fin_vehicle_fix_fee_lst  " _
        & " WHERE vehicle_id = " & TransFieldValueToString(pszVehicleID) _
        & " AND bus_date = " & TransFieldValueToString(pdtDate)
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then
        '出车辆固定费用已存在的错
        RaiseError ERR_VehicleFixFeeNotExist
    Else
        '如果费用已经扣过不就允许修改或删除
        
        If TransFieldValueToString(FormatDbValue(rsTemp!is_dec)) = TransFieldValueToString(True) Then
            RaiseError ERR_VehicleFixFeeHasDesCannotEdit
        End If
    End If
    
    
    szSql = " DELETE fin_vehicle_fix_fee_lst " _
        & " WHERE vehicle_id = " & TransFieldValueToString(pszVehicleID) _
                & " AND bus_date = " & TransFieldValueToString(pdtDate)
    oDB.Execute szSql
'    If lEffect = 0 Then
'        '出记录未找到的错
'        RaiseError ERR_VehicleFixFeeHasExist
'    End If
    
    
    Exit Sub
ErrorHandle:
    err.Raise err.Number, err.Source
    
End Sub


'汇总出车辆在某段时间内的所有结算单的应结金额
'该段时间为结算起始日期小于查询的起始日期,并且结算的结束日期大于查询的结束日期

Public Function TotalVehicleSettlePrice(pszVehicle As String, pdtStartDate As Date, pdtEndDate As Date)
    On Error GoTo ErrorHandle
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim i As Integer
    
    AssertObjIsValid
    
    oDB.ConnectionString = GetConnectionStr()
    
    
    szSql = " SELECT SUM(settle_price-settle_station_price ) settle_other_price FROM fin_settle_sheet_lst " _
        & " WHERE object_id = " & TransFieldValueToString(pszVehicle) _
        & " AND settle_object = " & TransFieldValueToString(CS_SettleByVehicle) _
        & " AND start_date <= " & TransFieldValueToString(pdtStartDate) _
        & " AND end_date >= " & TransFieldValueToString(pdtEndDate) _
        & " AND status <> " & TransFieldValueToString(CS_SettleSheetInvalid)
    
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    TotalVehicleSettlePrice = FormatDbValue(rsTemp!settle_other_price)
    
    Exit Function
ErrorHandle:
    err.Raise err.Number, err.Source
    
End Function



'根据计算应结费用,判断是否是负数,如果是负数,则结算单状态设为扣款未结清,如果为正数,则设为已汇
'如果为负数,则还要将负数款,生成到下个月
Public Sub SetRemit(paszSettleSheet() As String)
    '设置某些结算单汇款
    On Error GoTo ErrorHandle
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim i As Integer
    Dim szWhere As String
    Dim nCount As Integer
    Dim rsTemp As Recordset
    Dim j As Integer
    
'    Dim nSettleNegativeSplitItem As Integer
'    Dim oSystemParam As New SystemParam
    Dim lEffect As Long
    Dim dbNegative As Double
    
    
    AssertObjIsValid
    '需要加一个汇款的权限
    AssertHaveRight m_oActiveUser, RIGHT_SetRemit
    
    oDB.ConnectionString = GetConnectionStr()
    
    nCount = ArrayLength(paszSettleSheet)
    
    
    
    
    For i = 1 To nCount - 1
        szWhere = szWhere & TransFieldValueToString(paszSettleSheet(i)) & " , "
    Next i
    
    szWhere = " ( " & szWhere & TransFieldValueToString(paszSettleSheet(i)) & " ) "
    
    
    '需判断一下,如果已作废或已汇款,则不允许汇款
    szSql = " SELECT * FROM fin_settle_sheet_lst " _
        & " WHERE settlement_sheet_id IN " & szWhere
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount <> nCount Then
        '说明有几张结算单并不存在,出结算单不存在的错
        RaiseError ERR_SettleSheetNotExist
    Else
        For i = 1 To rsTemp.RecordCount
            If FormatDbValue(rsTemp!Status) = CS_SettleSheetInvalid Then
                '如果已作废则不允许再次汇款
                RaiseError ERR_NotAllowRemitHasCancel
            
            ElseIf FormatDbValue(rsTemp!Status) = CS_SettleSheetSettled Then
                '如果已汇款,则不允许再次汇款
                RaiseError ERR_NotAllowRemitHasRemitted

            End If
            rsTemp.MoveNext
        Next i
    End If
    
    oDB.BeginTrans
        

        '设置正数款,则状态设为已汇
        szSql = " UPDATE fin_settle_sheet_lst SET status =  " & TransFieldValueToString(CS_SettleSheetSettled) _
            & " , remit_date = getdate() " _
            & " WHERE settlement_sheet_id IN " & szWhere
        oDB.Execute szSql
        
        
        
        '如果负数款需要累加到下个月
        If m_bAllowSettleTotalNegative Then
        
            '取出所有的为负数的款项,累计成为附加款
            '*********************************************************************************************************************************************
            '此处不考虑,当为负数的车辆,结算过几次.有几次是负的,有几次是正的的情况,因为象东阳,他是车主未来结算,他们把结算单的负数款一次性的累加到下个月而已。
            '难道为负数的款与银行的款汇是同一步的操作？
            '*********************************************************************************************************************************************
            '查询为负数的，并且按车辆结算的单据
            szSql = " SELECT * FROM fin_settle_sheet_lst WHERE (settle_price - settle_station_price < 0) " _
                & " AND settle_object = " & CS_SettleByVehicle _
                & " AND settlement_sheet_id IN " & szWhere
            Set rsTemp = oDB.Execute(szSql)
            For i = 1 To rsTemp.RecordCount
                '把附加款算到结束日期的下个月1号
                '并且状态需要是未扣的
                '并需要将减改为加,因为放在结给站方处,所以是减的
                If FormatDbValue(rsTemp!settle_object) = CS_SettleByVehicle Then
                    dbNegative = FormatMoney(TransFieldValueToString(FormatDbValue(rsTemp!settle_station_price) - FormatDbValue(rsTemp!settle_price)))
                    szSql = " UPDATE fin_vehicle_fix_fee_lst " _
                        & " SET split_item_" & m_nSettleNegativeSplitItem & " = split_item_" & m_nSettleNegativeSplitItem & " + " & dbNegative _
                        & " WHERE vehicle_id = " & TransFieldValueToString(FormatDbValue(rsTemp!object_id)) _
                        & " AND bus_date = " & TransFieldValueToString(GetFirstMonthDay(DateAdd("m", 1, FormatDbValue(rsTemp!end_date)))) _
                        & " AND is_dec = " & TransFieldValueToString(False)
                    oDB.Execute szSql, lEffect
                
                    '先进行修改,如果没有作修改,则进行新增
                    If lEffect = 0 Then
                    
                        szSql = "INSERT fin_vehicle_fix_fee_lst  ( " _
                            & " vehicle_id , " _
                            & " bus_date , "
                        For j = 1 To cnFixFeeCount - 1
                            szSql = szSql & "split_item_" & j & " , "
                        Next j
                        szSql = szSql & "split_item_" & j & " , " _
                            & " is_dec ) VALUES ( " _
                        
                        szSql = szSql & TransFieldValueToString(FormatDbValue(rsTemp!object_id)) & " , " _
                            & TransFieldValueToString(GetFirstMonthDay(DateAdd("m", 1, FormatDbValue(rsTemp!end_date)))) _
                        '只保存负数项,其他均设为0
                        For j = 1 To cnFixFeeCount
                            If j = m_nSettleNegativeSplitItem Then
                                szSql = szSql & " , " & dbNegative
                            Else
                                szSql = szSql & " , " & TransFieldValueToString(0)
                            End If
                        Next j
                        
                        szSql = szSql & " , " & TransFieldValueToString(False) & " ) "
                        
                        oDB.Execute szSql
                    End If
                ElseIf FormatDbValue(rsTemp!settle_object) = CS_SettleByBus Then
                    dbNegative = FormatMoney(TransFieldValueToString(FormatDbValue(rsTemp!settle_station_price) - FormatDbValue(rsTemp!settle_price)))
                    szSql = " UPDATE fin_bus_fix_fee_lst " _
                        & " SET split_item_" & m_nSettleNegativeSplitItem & " = split_item_" & m_nSettleNegativeSplitItem & " + " & dbNegative _
                        & " WHERE bus_id = " & TransFieldValueToString(FormatDbValue(rsTemp!object_id)) _
                        & " AND bus_date = " & TransFieldValueToString(GetFirstMonthDay(DateAdd("m", 1, FormatDbValue(rsTemp!end_date)))) _
                        & " AND is_dec = " & TransFieldValueToString(False)
                    oDB.Execute szSql, lEffect
                
                    '先进行修改,如果没有作修改,则进行新增
                    If lEffect = 0 Then
                    
                        szSql = "INSERT fin_bus_fix_fee_lst  ( " _
                            & " bus_id , " _
                            & " bus_date , "
                        For j = 1 To cnFixFeeCount - 1
                            szSql = szSql & "split_item_" & j & " , "
                        Next j
                        szSql = szSql & "split_item_" & j & " , " _
                            & " is_dec ) VALUES ( " _
                        
                        szSql = szSql & TransFieldValueToString(FormatDbValue(rsTemp!object_id)) & " , " _
                            & TransFieldValueToString(GetFirstMonthDay(DateAdd("m", 1, FormatDbValue(rsTemp!end_date)))) _
                        '只保存负数项,其他均设为0
                        For j = 1 To cnFixFeeCount
                            If j = m_nSettleNegativeSplitItem Then
                                szSql = szSql & " , " & dbNegative
                            Else
                                szSql = szSql & " , " & TransFieldValueToString(0)
                            End If
                        Next j
                        
                        szSql = szSql & " , " & TransFieldValueToString(False) & " ) "
                        
                        oDB.Execute szSql
                    End If
                End If
            Next i
        End If
        
    oDB.CommitTrans
    Exit Sub
    
ErrorHandle:
    oDB.RollbackTrans
    err.Raise err.Number, err.Source

    
End Sub

'修改车辆结算后的费用项
Public Sub UpdateVehicleSettleItem(patVehicleSettle() As TVehilceSettle)
    
    
    On Error GoTo ErrorHandle
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim i As Integer
    Dim szWhere As String
    Dim nCount As Integer
    Dim j As Integer
    Dim dbSettlePrice As Double
    Dim dbSettleStationPrice As Double
    Dim szSettlementSheetID As String
    Dim rsTemp As Recordset
    
    
    AssertObjIsValid
    
    AssertHaveRight m_oActiveUser, RIGHT_UpdateHasSettledItem
    '需要加一个权限
'    AssertHaveRight RIGHT_SetRemit
    
    oDB.ConnectionString = GetConnectionStr()
    
    nCount = ArrayLength(patVehicleSettle)
    If nCount = 0 Then Exit Sub
    
    '该车辆的费用更新只允许一张张结算单修改
    
    szSettlementSheetID = patVehicleSettle(1).SettlementSheetID
        
    For i = 1 To nCount
        If patVehicleSettle(i).SettlementSheetID <> szSettlementSheetID Then
            '出 同一时间只允许修改一张结算单的费用,的错
            
            RaiseError ERR_OnlyAllowUpdateOneSettleSheet
        End If
    Next i
    

    '需判断一下,如果已作废或已汇款,则不允许汇款
    szSql = " SELECT * FROM fin_settle_sheet_lst " _
        & " WHERE settlement_sheet_id = " & TransFieldValueToString(szSettlementSheetID)
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount <> nCount Then
        '说明结算单并不存在,出结算单不存在的错
        RaiseError ERR_SettleSheetNotExist
    Else
        If FormatDbValue(rsTemp!Status) = CS_SettleSheetInvalid Then
            '如果已作废,则不允许修改费用
            RaiseError ERR_NotAllowUpdateHasCancel
        
        ElseIf FormatDbValue(rsTemp!Status) = CS_SettleSheetSettled Then
            '如果已汇款,则不允许修改费用
            RaiseError ERR_NotAllowUpdateHasRemitted
            
        End If
    End If
    
    oDB.BeginTrans
    
        
        
        For i = 1 To nCount
            szSql = " UPDATE fin_settle_vehicle_lst " _
                & " SET settle_price = " & TransFieldValueToString(patVehicleSettle(i).SettlePrice) _
                & " ,  settle_station_price = " & TransFieldValueToString(patVehicleSettle(i).SettleStationPrice) _
                
            For j = 1 To cnSplitItemCount
                szSql = szSql & " , split_item_" & j & " = " & TransFieldValueToString(patVehicleSettle(i).SplitItem(j))
            Next j
            szSql = szSql & " WHERE settlement_sheet_id = " & TransFieldValueToString(patVehicleSettle(i).SettlementSheetID) _
                & " AND vehicle_id = " & TransFieldValueToString(patVehicleSettle(i).VehicleID)
            oDB.Execute szSql
        Next i
        
        
        '得到车辆的汇总数
        szSql = " SELECT SUM(settle_price) settle_price , SUM(settle_station_price) settle_station_price " _
            & " FROM fin_settle_vehicle_lst " _
            & " WHERE  settlement_sheet_id = " & TransFieldValueToString(szSettlementSheetID)
        Set rsTemp = oDB.Execute(szSql)
        dbSettlePrice = dbSettlePrice + FormatDbValue(rsTemp!settle_price)
        dbSettleStationPrice = dbSettleStationPrice + FormatDbValue(rsTemp!settle_station_price)
        '得到公司的汇总数
        szSql = " SELECT SUM(settle_price) settle_price , SUM(settle_station_price) settle_station_price " _
            & " FROM fin_settle_company_price_lst " _
            & " WHERE  settlement_sheet_id = " & TransFieldValueToString(szSettlementSheetID)
        Set rsTemp = oDB.Execute(szSql)
        
        dbSettlePrice = dbSettlePrice + FormatDbValue(rsTemp!settle_price)
        dbSettleStationPrice = dbSettleStationPrice + FormatDbValue(rsTemp!settle_station_price)
        
        '更新<路单结算单表>
        szSql = " UPDATE fin_settle_sheet_lst " _
            & " SET settle_price =" & TransFieldValueToString(dbSettlePrice) _
            & " , settle_station_price = " & TransFieldValueToString(dbSettleStationPrice) _
            & " WHERE  settlement_sheet_id = " & TransFieldValueToString(szSettlementSheetID)
        oDB.Execute szSql
        
        
    oDB.CommitTrans
    Exit Sub
ErrorHandle:
    oDB.RollbackTrans
    err.Raise err.Number, err.Source
    
    
End Sub

'修改车次结算后的费用项
Public Sub UpdateBusSettleItem(patBusSettle() As TBusSettle)
    
    
    On Error GoTo ErrorHandle
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim i As Integer
    Dim szWhere As String
    Dim nCount As Integer
    Dim j As Integer
    Dim dbSettlePrice As Double
    Dim dbSettleStationPrice As Double
    Dim szSettlementSheetID As String
    Dim rsTemp As Recordset
    
    
    AssertObjIsValid
    
    AssertHaveRight m_oActiveUser, RIGHT_UpdateHasSettledItem
    '需要加一个权限
'    AssertHaveRight RIGHT_SetRemit
    
    oDB.ConnectionString = GetConnectionStr()
    
    nCount = ArrayLength(patBusSettle)
    If nCount = 0 Then Exit Sub
    
    '该车次的费用更新只允许一张张结算单修改
    
    szSettlementSheetID = patBusSettle(1).SettlementSheetID
        
    For i = 1 To nCount
        If patBusSettle(i).SettlementSheetID <> szSettlementSheetID Then
            '出 同一时间只允许修改一张结算单的费用,的错
            
            RaiseError ERR_OnlyAllowUpdateOneSettleSheet
        End If
    Next i
    

    '需判断一下,如果已作废或已汇款,则不允许汇款
    szSql = " SELECT * FROM fin_settle_sheet_lst " _
        & " WHERE settlement_sheet_id = " & TransFieldValueToString(szSettlementSheetID)
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount <> nCount Then
        '说明结算单并不存在,出结算单不存在的错
        RaiseError ERR_SettleSheetNotExist
    Else
        If FormatDbValue(rsTemp!Status) = CS_SettleSheetInvalid Then
            '如果已作废,则不允许修改费用
            RaiseError ERR_NotAllowUpdateHasCancel
        
        ElseIf FormatDbValue(rsTemp!Status) = CS_SettleSheetSettled Then
            '如果已汇款,则不允许修改费用
            RaiseError ERR_NotAllowUpdateHasRemitted
            
        End If
    End If
    
    oDB.BeginTrans
    
        
        
        For i = 1 To nCount
            szSql = " UPDATE fin_settle_bus_lst " _
                & " SET settle_price = " & TransFieldValueToString(patBusSettle(i).SettlePrice) _
                & " ,  settle_station_price = " & TransFieldValueToString(patBusSettle(i).SettleStationPrice) _
                
            For j = 1 To cnSplitItemCount
                szSql = szSql & " , split_item_" & j & " = " & TransFieldValueToString(patBusSettle(i).SplitItem(j))
            Next j
            szSql = szSql & " WHERE settlement_sheet_id = " & TransFieldValueToString(patBusSettle(i).SettlementSheetID) _
                & " AND bus_id = " & TransFieldValueToString(patBusSettle(i).BusID)
            oDB.Execute szSql
        Next i
        
        
        '得到车辆的汇总数
        szSql = " SELECT SUM(settle_price) settle_price , SUM(settle_station_price) settle_station_price " _
            & " FROM fin_settle_bus_lst " _
            & " WHERE  settlement_sheet_id = " & TransFieldValueToString(szSettlementSheetID)
        Set rsTemp = oDB.Execute(szSql)
        dbSettlePrice = dbSettlePrice + FormatDbValue(rsTemp!settle_price)
        dbSettleStationPrice = dbSettleStationPrice + FormatDbValue(rsTemp!settle_station_price)
        '得到公司的汇总数
        szSql = " SELECT SUM(settle_price) settle_price , SUM(settle_station_price) settle_station_price " _
            & " FROM fin_settle_company_price_lst " _
            & " WHERE  settlement_sheet_id = " & TransFieldValueToString(szSettlementSheetID)
        Set rsTemp = oDB.Execute(szSql)
        
        dbSettlePrice = dbSettlePrice + FormatDbValue(rsTemp!settle_price)
        dbSettleStationPrice = dbSettleStationPrice + FormatDbValue(rsTemp!settle_station_price)
        
        '更新<路单结算单表>
        szSql = " UPDATE fin_settle_sheet_lst " _
            & " SET settle_price =" & TransFieldValueToString(dbSettlePrice) _
            & " , settle_station_price = " & TransFieldValueToString(dbSettleStationPrice) _
            & " WHERE  settlement_sheet_id = " & TransFieldValueToString(szSettlementSheetID)
        oDB.Execute szSql
        
        
    oDB.CommitTrans
    Exit Sub
ErrorHandle:
    oDB.RollbackTrans
    err.Raise err.Number, err.Source
    
    
End Sub

'修改公司结算后的费用项
Public Sub UpdateCompanySettleItem(patCompanySettle() As TCompnaySettle)
    
    
    On Error GoTo ErrorHandle
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim i As Integer
    Dim szWhere As String
    Dim nCount As Integer
    Dim j As Integer
    Dim dbSettlePrice As Double
    Dim dbSettleStationPrice As Double
    Dim szSettlementSheetID As String
    
    Dim rsTemp As Recordset
    
    
    
    AssertObjIsValid
    
    AssertHaveRight m_oActiveUser, RIGHT_UpdateHasSettledItem
    
    '需要加一个权限
'    AssertHaveRight RIGHT_SetRemit
    
    oDB.ConnectionString = GetConnectionStr()
    
    nCount = ArrayLength(patCompanySettle)
    If nCount = 0 Then Exit Sub
    
    '该公司的费用更新只允许一张张结算单修改
    
    
    szSettlementSheetID = patCompanySettle(1).SettlementSheetID
        
    For i = 1 To nCount
        If patCompanySettle(i).SettlementSheetID <> szSettlementSheetID Then
            '出 同一时间只允许修改一张结算单的费用,的错
            
            RaiseError ERR_OnlyAllowUpdateOneSettleSheet
        End If
    Next i
    

    '需判断一下,如果已作废或已汇款,则不允许汇款
    szSql = " SELECT * FROM fin_settle_sheet_lst " _
        & " WHERE settlement_sheet_id = " & TransFieldValueToString(szSettlementSheetID)
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount <> nCount Then
        '说明结算单并不存在,出结算单不存在的错
        RaiseError ERR_SettleSheetNotExist
    Else
        If FormatDbValue(rsTemp!Status) = CS_SettleSheetInvalid Then
            '如果已作废,则不允许修改费用
            RaiseError ERR_NotAllowUpdateHasCancel
        
        ElseIf FormatDbValue(rsTemp!Status) = CS_SettleSheetSettled Then
            '如果已汇款,则不允许修改费用
            RaiseError ERR_NotAllowUpdateHasRemitted
            
        End If
    End If
    
    oDB.BeginTrans
        
        For i = 1 To nCount
            szSql = " UPDATE fin_settle_company_price_lst " _
                & " SET settle_price = " & TransFieldValueToString(patCompanySettle(i).SettlePrice) _
                & " ,  settle_station_price = " & TransFieldValueToString(patCompanySettle(i).SettleStationPrice) _
                
            For j = 1 To cnSplitItemCount
                szSql = szSql & " , split_item_" & j & " = " & TransFieldValueToString(patCompanySettle(i).SplitItem(j))
            Next j
            szSql = szSql & " WHERE settlement_sheet_id = " & TransFieldValueToString(patCompanySettle(i).SettlementSheetID) _
                & " AND transport_company_id = " & TransFieldValueToString(patCompanySettle(i).CompanyID)
            oDB.Execute szSql
        Next i
        
        
        szSettlementSheetID = patCompanySettle(1).SettlementSheetID
        '得到车辆的汇总数
        szSql = " SELECT SUM(settle_price) settle_price , SUM(settle_station_price) settle_station_price " _
            & " FROM fin_settle_vehicle_lst " _
            & " WHERE  settlement_sheet_id = " & TransFieldValueToString(szSettlementSheetID)
        Set rsTemp = oDB.Execute(szSql)
        dbSettlePrice = dbSettlePrice + FormatDbValue(rsTemp!settle_price)
        dbSettleStationPrice = dbSettleStationPrice + FormatDbValue(rsTemp!settle_station_price)
        '得到公司的汇总数
        szSql = " SELECT SUM(settle_price) settle_price , SUM(settle_station_price) settle_station_price " _
            & " FROM fin_settle_company_price_lst " _
            & " WHERE  settlement_sheet_id = " & TransFieldValueToString(szSettlementSheetID)
        Set rsTemp = oDB.Execute(szSql)
        
        dbSettlePrice = dbSettlePrice + FormatDbValue(rsTemp!settle_price)
        dbSettleStationPrice = dbSettleStationPrice + FormatDbValue(rsTemp!settle_station_price)
        
        '更新<路单结算单表>
        szSql = " UPDATE fin_settle_sheet_lst " _
            & " settle_price =" & TransFieldValueToString(dbSettlePrice) _
            & " , settle_station_price = " & TransFieldValueToString(dbSettleStationPrice) _
            & " WHERE  settlement_sheet_id = " & TransFieldValueToString(szSettlementSheetID)
        oDB.Execute szSql
        
        
    oDB.CommitTrans
    Exit Sub
ErrorHandle:
    oDB.RollbackTrans
    err.Raise err.Number, err.Source
    
    
End Sub





''=========================================================
''理论结算,从检票数据得到
''此理论结算,只按照车辆来结算
''返回预结算后的信息,此接口并不写数据库,只是将拆算结果返回.
''=========================================================
'Public Function PreviewSplitCheckSheet_TheoryCheck(pszVehicleID As String, pdtStartDate As Date, pdtEndDate As Date, Optional pszRouteID As String = "") As TSplitResult
'
'
'
'
'    On Error GoTo PreviewSplitCheckSheet_TheoryCheckErr
'
'    Dim szSql As String
'    Dim rsTemp As Recordset
'    Dim oDB As New RTConnection
'
'
'    Dim bSplitByCompany As Boolean
'    Dim nSheetCount As Integer '路单数
'    Dim szWhere As String
'
'
'    Dim rsProtocol As Recordset
'    Dim rsHalve As Recordset
'    Dim rsCompanySettlePrice As Recordset
'    Dim rsVehicleSettlePrice As Recordset
'    Dim rsDefaultProtocol As Recordset '默认协议
'    Dim rsHalveProtocol As Recordset '加总平分公司的协议
'
'
'    Dim szHalveCompanyID As String '平分给对方的公司代码
'    Dim szHalveCompanyName As String
'    Dim dbHalveRatio As Double '加总平分的比率
'
'    Dim i As Integer
'    Dim j As Integer
'    Dim k As Integer
'    Dim l As Integer
'
'    Dim aszFormulaValue() As String
'    Dim nCount As Integer '票价项数
'
'
'    Dim tResultInfo As TSplitResult '返回的结果信息
'
'    Dim atCompanyInfo() As TCompnaySettle '公司结算明细
'    Dim atVehicleInfo() As TVehilceSettle '车辆结算明细
''    Dim atStationInfo() As TSettleSheetStation '站点结算明细
'
'    Dim dbSplitItem As Double '临时拆出来的某一项的金额
'    Dim dbTotalTicketPrice As Double '总票款
'    Dim dbTotalQuantity As Double '总人数
'
'    Dim dbSettleOtherCompanyPrice As Double '总的结给对方公司金额
'    Dim dbSettleStationPrice As Double '总的结给车站的金额
'
'    Dim dbSplitlimitCharge As Double '底限额
'    Dim dbSplitUpCharge As Double '上限额
'
'
'    Dim bIsDefaut As Boolean '是否默认协议
'
'    Dim nVailableNumber As Integer '可用的结算项的数目
'
'
'    Dim nItem As Integer
'    Dim szItem As String
'
'
'    Dim prsPassageNumber As Recordset
'
'
'    Dim bIsHalveDefault As Boolean
'
'
'
'
'
'    AssertObjIsValid
'    AssertHaveRight m_oActiveUser, RIGHT_SplitCheckSheet '验证用户是否有结算的权限
'
'    oDB.ConnectionString = GetConnectionStr()
'
'
'    bSplitByCompany = False
'
'
'    szSql = " SELECT * FROM price_item_info "
'    Set rsTemp = oDB.Execute(szSql)
'
'    nCount = rsTemp.RecordCount
'
'    aszFormulaValue = GetFormulaItem
'
'
'    '取出默认的协议
'    bIsDefaut = False
'    Set rsDefaultProtocol = GetProtocolFormulaCompany(, True)
'
'
'    szSql = " SELECT * FROM split_item_info WHERE split_status = " & TransFieldValueToString(ESplitItemUseStatus.CS_SplitItemUse)
'    Set rsTemp = oDB.Execute(szSql)
'    nVailableNumber = rsTemp.RecordCount
'
'
'    m_bIsNew = False
'
'
'        '按车辆结算
'        Dim szCompanyID As String
'
'
'        '1.汇总出各车辆、上车站、线路、站点、票种的人数及各票价项prsPassageNumber。
'        szSql = "SELECT " & TransFieldValueToString(pszVehicleID) & " vehicle_id  , c.route_id , i.station_id sell_station_id , s.station_id , s.ticket_type " _
'            & " , MAX(license_tag_no) license_tag_no , MAX(route_name) route_name  " _
'            & " , MAX(i.station_name) sell_station_name , MAX(f.station_name) station_name,MAX(d.vehicle_type_code) vehicle_type_code  " _
'            & " , MAX(ticket_type_name) ticket_type_name , MAX(c.transport_company_id) transport_company_id " _
'            & " , MAX(m.transport_company_short_name) transport_company_short_name , MAX(a.annotation) annotation " _
'            & " , SUM(quantity) quantity , SUM(s.mileage*quantity) mileage , SUM(ticket_price*quantity) ticket_price , SUM(base_carriage*quantity) base_carriage " _
'            & " , SUM(price_item_1*quantity) price_item_1 , SUM(price_item_2*quantity) price_item_2 , SUM(price_item_3*quantity) price_item_3 , SUM(price_item_4*quantity) price_item_4 " _
'            & " , SUM(price_item_5*quantity) price_item_5 , SUM(price_item_6*quantity) price_item_6 , SUM(price_item_7*quantity) price_item_7 , SUM(price_item_8*quantity) price_item_8 " _
'            & " , SUM(price_item_9*quantity) price_item_9 , SUM(price_item_10*quantity) price_item_10 , SUM(price_item_11*quantity) price_item_11 , SUM(price_item_12*quantity) price_item_12 " _
'            & " , SUM(price_item_13*quantity) price_item_13 , SUM(price_item_14*quantity) price_item_14 , SUM(price_item_15*quantity) price_item_15 " _
'            & " FROM " & cszTableCheckSheet & " c, " & cszTableSheetStation & " s , " & cszTableVehicle & " p , " & cszTableRouteInfo & " r " _
'            & " , " & cszTableStationInfo & "  i , " & cszTableStationInfo & " f , " & cszTableTicketType & " k ," & cszTableCompany & " m " _
'            & " , " & cszTableVehicleTypeCode & " d , area_code a , sell_station_info t  " _
'            & " WHERE c.vehicle_type_code=d.vehicle_type_code AND p.vehicle_id=" & TransFieldValueToString(pszVehicleID) & " AND c.check_sheet_id = s.check_sheet_id " _
'            & " AND c.route_id = r.route_id AND m.transport_company_id=c.transport_company_id " _
'            & " AND s.sell_station_id = t.sell_station_id AND t.station_id = i.station_id AND s.station_id = f.station_id AND s.ticket_type = k.ticket_type_id " _
'            & " AND f.area_code = a.area_code " _
'            & " AND c." _
'            & " GROUP BY c.route_id , i.station_id , s.station_id , s.ticket_type " _
'            & " ORDER BY c.route_id , i.station_id , s.station_id , s.ticket_type "
'
'
'        Set prsPassageNumber = oDB.Execute(szSql)
'
'        szCompanyID = FormatDbValue(prsPassageNumber!transport_company_id)
'        '2.  得到需要结算的车辆所属公司的加总平分公司信息。
'
'        szSql = " SELECT h.*, c.transport_company_short_name FROM fin_halve_company_lst h, " & cszTableCompany & " c " _
'            & " WHERE c.transport_company_id = h.transport_company_id AND c.transport_company_id = " & TransFieldValueToString(szCompanyID)
'
'        Set rsHalve = oDB.Execute(szSql)
'        '3． 得到该车辆的协议
'        Set rsProtocol = GetProtocolFormulaVehicle(pszVehicleID, False)
'        bIsDefaut = False
'        If rsProtocol.RecordCount = 0 Then
'            '取出默认的协议
'            If rsDefaultProtocol.RecordCount <> 0 Then
'                Set rsProtocol = rsDefaultProtocol
'                bIsDefaut = True
'            Else
'                '该车辆协议未设置
'                RaiseError ERR_VehicleNotSetProtocol
'            End If
'        End If
'
'        '4． For I = 1 to prsPassageNumber.RecordCount
'        For i = 1 To prsPassageNumber.RecordCount
'
'
'            szHalveCompanyID = ""
'            For j = 1 To rsHalve.RecordCount
'                If FormatDbValue(prsPassageNumber!route_id) = FormatDbValue(rsHalve!route_id) And FormatDbValue(prsPassageNumber!transport_company_id) = FormatDbValue(rsHalve!transport_company_id) Then
'                    '找到该公司在该线路上需要加总平分
'                    szHalveCompanyID = FormatDbValue(rsHalve!other_transport_company_id)
'                    szHalveCompanyName = GetHavleOtherCompanyName(szHalveCompanyID)
'                    dbHalveRatio = FormatDbValue(rsHalve!Ratio)
'                    Exit For
'                End If
'            Next j
'            If szHalveCompanyID = "" Then
'                '不需要加总平分
'                '将相应的人数及各售票票价只算给该车辆。
'                rsProtocol.MoveFirst
'
'                '如果公式中没有用到结算价,则不进行查询结算价
'                If FindSettlePriceStr(rsProtocol, nVailableNumber, cszSettleStr) Then
'                    Set rsVehicleSettlePrice = GetVehicleSettlePrice(prsPassageNumber!vehicle_id, prsPassageNumber!route_id, prsPassageNumber!sell_station_id, prsPassageNumber!station_id)
'                End If
'
'                For j = 1 To nVailableNumber
'                    SetFormulaValueVehicle prsPassageNumber, aszFormulaValue, nCount, rsVehicleSettlePrice, rsProtocol, atVehicleInfo, pszVehicleID, FormatDbValue(prsPassageNumber!license_tag_no), False, 0
'                    rsProtocol.MoveNext
'                    m_bIsNew = True
'                Next j
'            Else
'                'Else累加该车辆的结算费用明细将这些人数及各售票票价项分为两半，一半算给该车辆，另一半算给加总平分中的另一公司。
'
'
'                '先结给该车辆，方法与上同。
'
'                '如果公式中没有用到结算价,则不进行查询结算价
'                If FindSettlePriceStr(rsProtocol, nVailableNumber, cszSettleStr) Then
'                    Set rsVehicleSettlePrice = GetVehicleSettlePrice(prsPassageNumber!vehicle_id, prsPassageNumber!route_id, prsPassageNumber!sell_station_id, prsPassageNumber!station_id)
'                End If
'
'                rsProtocol.MoveFirst
'                For j = 1 To nVailableNumber
'                    SetFormulaValueVehicle prsPassageNumber, aszFormulaValue, nCount, rsVehicleSettlePrice, rsProtocol, atVehicleInfo, pszVehicleID, FormatDbValue(prsPassageNumber!license_tag_no), False, dbHalveRatio
'                    rsProtocol.MoveNext
'                    m_bIsNew = True
'                Next j
'                '再结给该公司，与按公司结算类似
'                '取出协议
'                bIsHalveDefault = False
'                Set rsHalveProtocol = GetProtocolFormulaCompany(szHalveCompanyID)
'                If rsHalveProtocol.RecordCount = 0 Then
'                    '取出默认的协议
'                    If rsDefaultProtocol.RecordCount <> 0 Then
'                        Set rsHalveProtocol = rsDefaultProtocol
'                        bIsHalveDefault = True
'                    Else
'                        '该公司协议未设置
'                        RaiseError ERR_VehicleNotSetProtocol
'                    End If
'                End If
'
'                '如果公式中没有用到结算价,则不进行查询结算价
'                If FindSettlePriceStr(rsHalveProtocol, nVailableNumber, cszSettleStr) Then
'                    Set rsVehicleSettlePrice = GetVehicleSettlePrice(prsPassageNumber!vehicle_id, prsPassageNumber!route_id, prsPassageNumber!sell_station_id, prsPassageNumber!station_id)
'                End If
'
'                m_bIsNew = False
'                rsHalveProtocol.MoveFirst
'                FindRouteProtocol rsHalveProtocol, FormatDbValue(prsPassageNumber!route_id), bIsHalveDefault
'                For j = 1 To nVailableNumber
'                    SetFormulaValueCompany prsPassageNumber, aszFormulaValue, nCount, rsCompanySettlePrice, rsHalveProtocol, atCompanyInfo, szHalveCompanyID, szHalveCompanyName, True, bIsHalveDefault, 1 - dbHalveRatio
'                    rsHalveProtocol.MoveNext
'                    m_bIsNew = True
'                Next j
'
'            End If
'
'            dbTotalTicketPrice = FormatDbValue(prsPassageNumber!ticket_price)
'            dbTotalQuantity = FormatDbValue(prsPassageNumber!Quantity)
'            tResultInfo.SettleSheetInfo.TotalTicketPrice = tResultInfo.SettleSheetInfo.TotalTicketPrice + dbTotalTicketPrice
'            tResultInfo.SettleSheetInfo.TotalQuantity = tResultInfo.SettleSheetInfo.TotalQuantity + dbTotalQuantity
'            prsPassageNumber.MoveNext
'            m_bIsNew = False
'        Next i
'
'
'
'    '将总路单数、总票款、总人数、应结票款、结给车站票款以及公司的费用明细和车辆明细均返回
'    With tResultInfo.SettleSheetInfo
'        '=========================================
'        Dim rsCheckSheetCount As Recordset
'        szSql = ""
'        Set rsCheckSheetCount = oDB.Execute(szSql)
'
'
'
'        .CheckSheetCount = 0 'ArrayLength(paszCheckSheetID) '路单数
'        For i = 1 To ArrayLength(atCompanyInfo)
'            rsProtocol.MoveFirst
''            当结算有底限时 , 现在的处理方式是不对的, 未判断不同公司, 线路不同, 其协议不同, 协议不同时当然其底限也是不同的
''            If bSplitByCompany Then
''               FindRouteProtocol rsProtocol, rsCompanySettlePrice, bIsDefaut
''            End If
'            For j = 1 To nVailableNumber 'rsProtocol.RecordCount
'                szItem = FormatDbValue(rsProtocol!split_item_id)
'                nItem = CInt(szItem)
'                dbSplitlimitCharge = GetlimitCharge(FormatDbValue(rsProtocol!protocol_id), szItem, dbSplitUpCharge)
'                '如果上限不为0,则   该项值=上限值-该项值
'                If dbSplitUpCharge <> 0 Then
'                    atCompanyInfo(i).SplitItem(nItem) = FormatMoney(dbSplitUpCharge - atCompanyInfo(i).SplitItem(nItem))
'                End If
'                '当下限值大于该项值时,该值为下限值
'                If dbSplitlimitCharge > atCompanyInfo(i).SplitItem(nItem) Then
'                    atCompanyInfo(i).SplitItem(nItem) = dbSplitlimitCharge
'                Else
''                    atCompanyInfo(i).SplitItem(nItem) = atCompanyInfo(i).SplitItem(nItem)
'                End If
'                atCompanyInfo(i).SplitItem(nItem) = FormatMoney(atCompanyInfo(i).SplitItem(nItem))
'                atCompanyInfo(i).SettlePrice = atCompanyInfo(i).SettlePrice + GetSettlePrice(rsProtocol, atCompanyInfo(i).SplitItem(nItem))
'                atCompanyInfo(i).SettleStationPrice = atCompanyInfo(i).SettleStationPrice + GetSettleStationPrice(rsProtocol, atCompanyInfo(i).SplitItem(nItem))
'                rsProtocol.MoveNext
'            Next j
'            atCompanyInfo(i).SettlePrice = FormatMoney(atCompanyInfo(i).SettlePrice) ' FormatMoney (atCompanyInfo(i).SettlePrice + 0.01, 1)
'            atCompanyInfo(i).SettleStationPrice = FormatMoney(atCompanyInfo(i).SettleStationPrice) ' Round(atCompanyInfo(i).SettleStationPrice + 0.01, 1)
'
'            dbSettleOtherCompanyPrice = dbSettleOtherCompanyPrice + atCompanyInfo(i).SettlePrice
'            dbSettleStationPrice = dbSettleStationPrice + atCompanyInfo(i).SettleStationPrice
'        Next i
''        If pszCompanyID <> "" Then
''            .ObjectID = pszCompanyID '对象代码
''            .SettleObject = CS_SettleByTransportCompany  '结算对象
''
''        Else
'            .ObjectID = pszVehicleID
'            .SettleObject = CS_SettleByVehicle
''        End If
'        If .SettleObject = CS_SettleByVehicle And m_nFixFeeItemCount > 0 Then
'            '如果是按照车辆结算,并且固定费用项存在 ,则读取该车辆的所有的固定费用项
'            '============================
'            Dim rsFixFee As Recordset
'            szSql = " SELECT SUM(split_item_1) split_item_1 , SUM(split_item_2) split_item_2 , SUM(split_item_3) split_item_3 " _
'                & " , SUM(split_item_4) split_item_4, SUM(split_item_5) split_item_5, SUM(split_item_6) split_item_6, SUM(split_item_7) split_item_7 " _
'                & " , SUM(split_item_8) split_item_8, SUM(split_item_9) split_item_9, SUM(split_item_10) split_item_10, SUM(split_item_11) split_item_11 " _
'                & " , SUM(split_item_12) split_item_12, SUM(split_item_13) split_item_13, SUM(split_item_14) split_item_14, SUM(split_item_15) split_item_15 " _
'                & " , SUM(split_item_16) split_item_16, SUM(split_item_17) split_item_17, SUM(split_item_18) split_item_18, SUM(split_item_19) split_item_19 " _
'                & " , SUM(split_item_20) split_item_20 " _
'                & " FROM fin_vehicle_fix_fee_lst" _
'                & " WHERE vehicle_id = " & TransFieldValueToString(pszVehicleID) _
'                & " AND bus_date <" & TransFieldValueToString(pdtEndDate) _
'                & " AND is_dec = " & TransFieldValueToString(False)
'
'            '& " AND bus_date >= " & TransFieldValueToString(pdtStartDate)
'
'            Set rsFixFee = oDB.Execute(szSql)
'
'
'
'        End If
'        For i = 1 To ArrayLength(atVehicleInfo)
'
'            rsProtocol.MoveFirst
'            For j = 1 To nVailableNumber 'rsProtocol.RecordCount
'
'                szItem = FormatDbValue(rsProtocol!split_item_id)
'                nItem = Val(szItem)
'
'                If m_nFixFeeItemCount > 0 Then
'                    '如果存在固定费用项
'                    For k = 0 To m_nFixFeeItemCount - 1
'                        If nItem = m_aszFixFeeItem(k) Then
'                            atVehicleInfo(i).SplitItem(nItem) = FormatDbValue(rsFixFee.Fields("split_item_" & nItem))
'                            Exit For
'                        End If
'                    Next k
'                End If
'                If m_nFixFeeItemCount <= 0 Or k > m_nFixFeeItemCount Then
'                    '如果未找到,或不存在固定费用项
'
'                    dbSplitlimitCharge = GetlimitCharge(FormatDbValue(rsProtocol!protocol_id), szItem, dbSplitUpCharge)
'                    '如果上限不为0,则   该项值=上限值-该项值
'                    If dbSplitUpCharge <> 0 Then
'                        atVehicleInfo(i).SplitItem(nItem) = FormatMoney(dbSplitUpCharge - atVehicleInfo(i).SplitItem(nItem))
'                    End If
'                    '当下限值大于该项值时,该值为下限值
'                    If dbSplitlimitCharge > atVehicleInfo(i).SplitItem(nItem) Then
'                        atVehicleInfo(i).SplitItem(nItem) = dbSplitlimitCharge
'                    Else
'    '                    atVehicleInfo(i).SplitItem(nItem) = atVehicleInfo(i).SplitItem(nItem)
'                    End If
'                    atVehicleInfo(i).SplitItem(nItem) = FormatMoney(atVehicleInfo(i).SplitItem(nItem))
'                End If
'                atVehicleInfo(i).SettlePrice = atVehicleInfo(i).SettlePrice + GetSettlePrice(rsProtocol, atVehicleInfo(i).SplitItem(nItem))
'                atVehicleInfo(i).SettleStationPrice = atVehicleInfo(i).SettleStationPrice + GetSettleStationPrice(rsProtocol, atVehicleInfo(i).SplitItem(nItem))
'                rsProtocol.MoveNext
'            Next j
'            atVehicleInfo(i).SettlePrice = FormatMoney(atVehicleInfo(i).SettlePrice)  ' Round(atVehicleInfo(i).SettlePrice + 0.01, 1)
'            atVehicleInfo(i).SettleStationPrice = FormatMoney(atVehicleInfo(i).SettleStationPrice)  ' Round(atVehicleInfo(i).SettleStationPrice + 0.01, 1)
'
'
'            dbSettleOtherCompanyPrice = dbSettleOtherCompanyPrice + atVehicleInfo(i).SettlePrice
'            dbSettleStationPrice = dbSettleStationPrice + atVehicleInfo(i).SettleStationPrice
'        Next i
'        dbSettleOtherCompanyPrice = dbSettleOtherCompanyPrice ' Round(dbSettleOtherCompanyPrice + 0.01, 1)
'        dbSettleStationPrice = dbSettleStationPrice ' Round(dbSettleStationPrice + 0.01, 1)
'
'        .SettleOtherCompanyPrice = dbSettleOtherCompanyPrice '结算给对方公司金额
'        .SettleStationPrice = dbSettleStationPrice '结给站方金额
'
'        If prsPassageNumber.RecordCount > 0 Then
'            prsPassageNumber.MoveFirst
'            .TransportCompanyID = FormatDbValue(prsPassageNumber!transport_company_id)
'            .TransportCompanyName = FormatDbValue(prsPassageNumber!transport_company_short_name)
'            .RouteID = FormatDbValue(prsPassageNumber!route_id)
'            szSql = " SELECT route_name FROM route_info WHERE route_id = " & TransFieldValueToString(.RouteID)
''            .RouteName = FormatDbValue(prsPassageNumber!route_name)
'            Set rsTemp = oDB.Execute(szSql)
'            If rsTemp.RecordCount > 0 Then
'                .RouteName = FormatDbValue(rsTemp!route_name)
'            End If
'        End If
'    End With
'
'    tResultInfo.CompanyInfo = atCompanyInfo
'    tResultInfo.VehicleInfo = atVehicleInfo
'    PreviewSplitCheckSheet_TheoryCheck = tResultInfo
'
'    Exit Function
'PreviewSplitCheckSheet_TheoryCheckErr:
'    err.Raise err.Number, err.Source, err.Description
'End Function




Public Sub ReTotalVehicleFixFee(pszSettleSheetID As String)
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsFixFee As Recordset
    Dim rsItem As Recordset
    

    Dim adbSplitItem() As Double
    Dim dbSettlePrice As Double
    Dim dbSettleStationPrice As Double
    Dim i As Integer
        
    On Error GoTo ErrorHandle
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_CancelSettleSheet
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    '查询有无该结算单
    szSql = " SELECT * FROM fin_settle_sheet_lst " _
        & " WHERE settlement_sheet_id = " & TransFieldValueToString(pszSettleSheetID)
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then RaiseError ERR_SettleSheetNotExist

    If FormatDbValue(rsTemp!Status) = CS_SettleSheetInvalid Then
        '如果已作废,则出错
        RaiseError ERR_SettleSheetHasCanceled

    ElseIf FormatDbValue(rsTemp!Status) = CS_SettleSheetSettled Then
        '如果已结算,则出错
        RaiseError ERR_SettleSheetHasSettled
    End If

    If FormatDbValue(rsTemp!settle_object) <> CS_SettleByVehicle And FormatDbValue(rsTemp!settle_object) <> CS_SettleByBus Then
        '出错,未按车辆在结算
        err.Raise 3343, err.Source, "未按车辆或车次在结算,不允许重新统计固定费用"
        Exit Sub
        
    End If
    
    oDB.BeginTrans
        '更新车辆表中的汇总数字
        If FormatDbValue(rsTemp!settle_object) = CS_SettleByVehicle Then
            szSql = " UPDATE fin_settle_vehicle_lst " _
                & " SET split_item_1 = v.split_item_1 + a.split_item_1 , split_item_2 = v.split_item_2 + a.split_item_2 " _
                & " , split_item_3 = v.split_item_3 + a.split_item_3  , split_item_4 = v.split_item_4 + a.split_item_4 " _
                & " , split_item_5 = v.split_item_5 + a.split_item_5 , split_item_6 = v.split_item_6 + a.split_item_6 " _
                & " , split_item_7 = v.split_item_7 + a.split_item_7  , split_item_8 = v.split_item_8 + a.split_item_8 " _
                & " , split_item_9 = v.split_item_9 + a.split_item_9 , split_item_10 = v.split_item_10 + a.split_item_10 " _
                & " , split_item_11 = v.split_item_11 + a.split_item_11 , split_item_12 = v.split_item_12 + a.split_item_12 " _
                & " , split_item_13 = v.split_item_13 + a.split_item_13 , split_item_14 = v.split_item_14 + a.split_item_14 " _
                & " , split_item_15 = v.split_item_15 + a.split_item_15 , split_item_16 = v.split_item_16 + a.split_item_16 " _
                & " , split_item_17 = v.split_item_17 + a.split_item_17 , split_item_18 = v.split_item_18 + a.split_item_18 " _
                & " , split_item_19 = v.split_item_19 + a.split_item_19 , split_item_20 = v.split_item_20 + a.split_item_20 " _
                & " FROM fin_settle_vehicle_lst v , ( " _
                & "     SELECT ISNULL(SUM(split_item_1),0) split_item_1 , ISNULL(SUM(split_item_2),0) split_item_2 , ISNULL(SUM(split_item_3),0) split_item_3 " _
                & " , ISNULL(SUM(split_item_4),0) split_item_4, ISNULL(SUM(split_item_5),0) split_item_5, ISNULL(SUM(split_item_6),0) split_item_6 " _
                & " , ISNULL(SUM(split_item_7),0) split_item_7  , ISNULL(SUM(split_item_8),0) split_item_8, ISNULL(SUM(split_item_9),0) split_item_9  " _
                & " , ISNULL(SUM(split_item_10),0) split_item_10, ISNULL(SUM(split_item_11),0) split_item_11  , ISNULL(SUM(split_item_12),0) split_item_12 " _
                & " , ISNULL(SUM(split_item_13),0) split_item_13, ISNULL(SUM(split_item_14),0) split_item_14, ISNULL(SUM(split_item_15),0) split_item_15 " _
                & " , ISNULL(SUM(split_item_16),0) split_item_16, ISNULL(SUM(split_item_17),0) split_item_17, ISNULL(SUM(split_item_18),0) split_item_18 " _
                & " , ISNULL(SUM(split_item_19),0) split_item_19  , ISNULL(SUM(split_item_20),0) split_item_20 " _
                & " FROM fin_vehicle_fix_fee_lst WHERE vehicle_id = " & TransFieldValueToString(rsTemp!object_id) & "  AND bus_date < " & TransFieldValueToString(rsTemp!end_date) _
                & " AND is_dec = " & TransFieldValueToString(False) _
                & " ) a " _
                & " WHERE v.settlement_sheet_id =" & TransFieldValueToString(pszSettleSheetID)
            
            Set rsFixFee = oDB.Execute(szSql)
            
            '修改<结算单车辆固定费用清单>,改为已扣
            szSql = " UPDATE fin_vehicle_fix_fee_lst SET is_dec = " & TransFieldValueToString(True) _
                & " WHERE vehicle_id = " & TransFieldValueToString(rsTemp!object_id) _
                & " AND bus_date < " & TransFieldValueToString(rsTemp!end_date)
                
            oDB.Execute szSql
            
            
            
            szSql = " SELECT  s.split_item_name , s.split_status , s.split_type , s.allow_modify  " _
                & " FROM  split_item_info s "
                
            Set rsItem = oDB.Execute(szSql)
            
            '得到结算单车辆拆算明细表 中的数据
            szSql = "SELECT * FROM fin_settle_vehicle_lst WHERE settlement_sheet_id =" & TransFieldValueToString(pszSettleSheetID)
            Set rsTemp = oDB.Execute(szSql)
            '如果有多条记录或无记录则不执行下去,
            
            If rsTemp.RecordCount <> 1 Then Exit Sub
            
            
            '汇总到 应结票款 \ 结给车站票款
            
            ReDim adbSplitItem(1 To cnSplitItemCount)
            dbSettlePrice = 0
            dbSettleStationPrice = 0
                
                For i = 1 To rsItem.RecordCount
                    adbSplitItem(i) = FormatDbValue(rsTemp.Fields("split_item_" & i))
                    
                    dbSettlePrice = dbSettlePrice + GetSettlePrice(rsItem, adbSplitItem(i))
                    dbSettleStationPrice = dbSettleStationPrice + GetSettleStationPrice(rsItem, adbSplitItem(i))
                    
                    rsItem.MoveNext
                Next i
                rsTemp.MoveNext
                
                
            dbSettleStationPrice = FormatMoney(dbSettleStationPrice)
            dbSettleStationPrice = FormatMoney(dbSettleStationPrice)
            
            '更新车辆结算明细表
            
            szSql = " UPDATE fin_settle_vehicle_lst SET settle_price = " & TransFieldValueToString(dbSettlePrice) _
                & " , settle_station_price = " & TransFieldValueToString(dbSettleStationPrice) _
                & " WHERE settlement_sheet_id =" & TransFieldValueToString(pszSettleSheetID)
            oDB.Execute szSql
            '更新路单结算表
            szSql = " UPDATE fin_settle_sheet_lst SET settle_price = " & TransFieldValueToString(dbSettlePrice) _
                & " , settle_station_price = " & TransFieldValueToString(dbSettleStationPrice) _
                & " WHERE settlement_sheet_id =" & TransFieldValueToString(pszSettleSheetID)
            oDB.Execute szSql
        ElseIf FormatDbValue(rsTemp!settle_object) = CS_SettleByBus Then
            szSql = " UPDATE fin_settle_bus_lst " _
                & " SET split_item_1 = v.split_item_1 + a.split_item_1 , split_item_2 = v.split_item_2 + a.split_item_2 " _
                & " , split_item_3 = v.split_item_3 + a.split_item_3  , split_item_4 = v.split_item_4 + a.split_item_4 " _
                & " , split_item_5 = v.split_item_5 + a.split_item_5 , split_item_6 = v.split_item_6 + a.split_item_6 " _
                & " , split_item_7 = v.split_item_7 + a.split_item_7  , split_item_8 = v.split_item_8 + a.split_item_8 " _
                & " , split_item_9 = v.split_item_9 + a.split_item_9 , split_item_10 = v.split_item_10 + a.split_item_10 " _
                & " , split_item_11 = v.split_item_11 + a.split_item_11 , split_item_12 = v.split_item_12 + a.split_item_12 " _
                & " , split_item_13 = v.split_item_13 + a.split_item_13 , split_item_14 = v.split_item_14 + a.split_item_14 " _
                & " , split_item_15 = v.split_item_15 + a.split_item_15 , split_item_16 = v.split_item_16 + a.split_item_16 " _
                & " , split_item_17 = v.split_item_17 + a.split_item_17 , split_item_18 = v.split_item_18 + a.split_item_18 " _
                & " , split_item_19 = v.split_item_19 + a.split_item_19 , split_item_20 = v.split_item_20 + a.split_item_20 " _
                & " FROM fin_settle_bus_lst v , ( " _
                & "     SELECT ISNULL(SUM(split_item_1),0) split_item_1 , ISNULL(SUM(split_item_2),0) split_item_2 , ISNULL(SUM(split_item_3),0) split_item_3 " _
                & " , ISNULL(SUM(split_item_4),0) split_item_4, ISNULL(SUM(split_item_5),0) split_item_5, ISNULL(SUM(split_item_6),0) split_item_6 " _
                & " , ISNULL(SUM(split_item_7),0) split_item_7  , ISNULL(SUM(split_item_8),0) split_item_8, ISNULL(SUM(split_item_9),0) split_item_9  " _
                & " , ISNULL(SUM(split_item_10),0) split_item_10, ISNULL(SUM(split_item_11),0) split_item_11  , ISNULL(SUM(split_item_12),0) split_item_12 " _
                & " , ISNULL(SUM(split_item_13),0) split_item_13, ISNULL(SUM(split_item_14),0) split_item_14, ISNULL(SUM(split_item_15),0) split_item_15 " _
                & " , ISNULL(SUM(split_item_16),0) split_item_16, ISNULL(SUM(split_item_17),0) split_item_17, ISNULL(SUM(split_item_18),0) split_item_18 " _
                & " , ISNULL(SUM(split_item_19),0) split_item_19  , ISNULL(SUM(split_item_20),0) split_item_20 " _
                & " FROM fin_bus_fix_fee_lst WHERE bus_id = " & TransFieldValueToString(rsTemp!object_id) & "  AND bus_date < " & TransFieldValueToString(rsTemp!end_date) _
                & " AND is_dec = " & TransFieldValueToString(False) _
                & " ) a " _
                & " WHERE v.settlement_sheet_id =" & TransFieldValueToString(pszSettleSheetID)
            
            Set rsFixFee = oDB.Execute(szSql)
            
            '修改<结算单车辆固定费用清单>,改为已扣
            szSql = " UPDATE fin_bus_fix_fee_lst SET is_dec = " & TransFieldValueToString(True) _
                & " WHERE bus_id = " & TransFieldValueToString(rsTemp!object_id) _
                & " AND bus_date < " & TransFieldValueToString(rsTemp!end_date)
                
            oDB.Execute szSql
            
            
            
            szSql = " SELECT  s.split_item_name , s.split_status , s.split_type , s.allow_modify  " _
                & " FROM  split_item_info s "
                
            Set rsItem = oDB.Execute(szSql)
            
            '得到结算单车辆拆算明细表 中的数据
            szSql = "SELECT * FROM fin_settle_bus_lst WHERE settlement_sheet_id =" & TransFieldValueToString(pszSettleSheetID)
            Set rsTemp = oDB.Execute(szSql)
            '如果有多条记录或无记录则不执行下去,
            
            If rsTemp.RecordCount <> 1 Then Exit Sub
            
            
            '汇总到 应结票款 \ 结给车站票款
            
            ReDim adbSplitItem(1 To cnSplitItemCount)
            dbSettlePrice = 0
            dbSettleStationPrice = 0
                
                For i = 1 To rsItem.RecordCount
                    adbSplitItem(i) = FormatDbValue(rsTemp.Fields("split_item_" & i))
                    
                    dbSettlePrice = dbSettlePrice + GetSettlePrice(rsItem, adbSplitItem(i))
                    dbSettleStationPrice = dbSettleStationPrice + GetSettleStationPrice(rsItem, adbSplitItem(i))
                    
                    rsItem.MoveNext
                Next i
                rsTemp.MoveNext
                
                
            dbSettleStationPrice = FormatMoney(dbSettleStationPrice)
            dbSettleStationPrice = FormatMoney(dbSettleStationPrice)
            
            '更新车辆结算明细表
            
            szSql = " UPDATE fin_settle_bus_lst SET settle_price = " & TransFieldValueToString(dbSettlePrice) _
                & " , settle_station_price = " & TransFieldValueToString(dbSettleStationPrice) _
                & " WHERE settlement_sheet_id =" & TransFieldValueToString(pszSettleSheetID)
            oDB.Execute szSql
            '更新路单结算表
            szSql = " UPDATE fin_settle_sheet_lst SET settle_price = " & TransFieldValueToString(dbSettlePrice) _
                & " , settle_station_price = " & TransFieldValueToString(dbSettleStationPrice) _
                & " WHERE settlement_sheet_id =" & TransFieldValueToString(pszSettleSheetID)
            oDB.Execute szSql
        End If
        
        
        
        '此处忽略存在车辆结算,同时结算加总平分的情况。
        '懒得写
        '语句不写了
        '2005-8-11 陈峰
        
        
    oDB.CommitTrans


    Exit Sub
ErrorHandle:
    oDB.RollbackTrans

    err.Raise err.Number, err.Source, err.Description

End Sub

Public Sub SetBusProtocol(pszBusID As String, Optional pszCompanyID As String, Optional pszProtocolID As String)
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim lEffect As Long
    On Error GoTo ErrorHandle
    
    oDB.ConnectionString = GetConnectionStr
    szSql = " UPDATE fin_bus_protocol_lst SET protocol_id =  " & TransFieldValueToString(pszProtocolID) _
        & " WHERE bus_id = " & TransFieldValueToString(pszBusID) _
        & " AND transport_company_id = " & TransFieldValueToString(pszCompanyID)
        
    oDB.Execute szSql, lEffect
    If lEffect = 0 Then
        szSql = " INSERT fin_bus_protocol_lst ( bus_id , " _
            & " transport_company_id , " _
            & " protocol_id ) VALUES ( " _
            & TransFieldValueToString(pszBusID) & " , " _
            & TransFieldValueToString(pszCompanyID) & " , " _
            & TransFieldValueToString(pszProtocolID) & " ) "
        oDB.Execute szSql
        
    End If
    
    Set oDB = Nothing

    Exit Sub
ErrorHandle:

    Set oDB = Nothing
    err.Raise err.Number, err.Source, err.Description

    
End Sub

'新增车次固定费用
Public Sub AddBusFixFee(prsFixFee As Recordset)
    
    On Error GoTo ErrorHandle
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim i As Integer
    Dim j As Integer
    
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_BusFixFee '验证用户是否有车次固定费用的权限
    
    oDB.ConnectionString = GetConnectionStr()
    
    If prsFixFee.RecordCount = 0 Then Exit Sub
    prsFixFee.MoveFirst
    oDB.BeginTrans
        For i = 1 To prsFixFee.RecordCount
        
            
            szSql = "SELECT 1 FROM fin_bus_fix_fee_lst  " _
                & " WHERE bus_id = " & TransFieldValueToString(prsFixFee!bus_id) _
                & " AND transport_company_id = " & TransFieldValueToString(prsFixFee!transport_company_id) _
                & " AND bus_date = " & TransFieldValueToString(prsFixFee!Bus_date)
            Set rsTemp = oDB.Execute(szSql)
            If rsTemp.RecordCount > 0 Then
                '车次固定费用已存在的出错
                RaiseError ERR_BusFixFeeHasExist
                
            End If
            '如果该车次当月的信息已经输入过一次,则出错
            
            szSql = "INSERT fin_bus_fix_fee_lst  ( " _
                & " bus_id , " _
                & " transport_company_id , " _
                & " bus_date , "
            For j = 1 To cnFixFeeCount - 1
                szSql = szSql & "split_item_" & j & " , "
            Next j
            szSql = szSql & "split_item_" & j & " , " _
                & " is_dec ) VALUES ( " _
            
            szSql = szSql & TransFieldValueToString(prsFixFee!bus_id) & " , " _
                & TransFieldValueToString(prsFixFee!transport_company_id) & " , " _
                & TransFieldValueToString(prsFixFee!Bus_date) & " , " _
    
            For j = 1 To cnFixFeeCount - 1
                
                szSql = szSql & TransFieldValueToString(FormatDbValue(prsFixFee.Fields("split_item_" & j))) & " , "
            Next j
            szSql = szSql & TransFieldValueToString(FormatDbValue(prsFixFee.Fields("split_item_" & j))) & " , " _
                & TransFieldValueToString(False) & " ) "
            
            oDB.Execute szSql
            
            prsFixFee.MoveNext
        Next i
    oDB.CommitTrans
    Exit Sub
ErrorHandle:
    oDB.RollbackTrans
    err.Raise err.Number, err.Source
    
End Sub

Public Sub DelBusProtocol(pszBusID As String)
On Error GoTo ErrorHandle
Dim szSql As String
Dim oDB As New RTConnection
oDB.ConnectionString = GetConnectionStr()
    szSql = " DELETE fin_bus_protocol_lst " _
        & " WHERE bus_id = " & TransFieldValueToString(pszBusID)
    oDB.Execute szSql
Exit Sub
ErrorHandle:
    err.Raise err.Number, err.Source
End Sub
'删除车次固定费用
Public Sub DelBusFixFee(pszVehicleID As String, pdtDate As Date)
    On Error GoTo ErrorHandle
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim i As Integer
    Dim lEffect As Long
    
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_BusFixFee '验证用户是否有车辆固定费用的权限
    
    oDB.ConnectionString = GetConnectionStr()
    szSql = "SELECT * FROM fin_bus_fix_fee_lst  " _
        & " WHERE bus_id = " & TransFieldValueToString(pszVehicleID) _
        & " AND bus_date = " & TransFieldValueToString(pdtDate)
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then
        '出车次固定费用已存在的错
        RaiseError ERR_BusFixFeeNotExist
    Else
        '如果费用已经扣过不就允许修改或删除
        
        If TransFieldValueToString(FormatDbValue(rsTemp!is_dec)) = TransFieldValueToString(True) Then
            RaiseError ERR_BusFixFeeHasDesCannotEdit
        End If
    End If
    
    
    szSql = " DELETE fin_bus_fix_fee_lst " _
        & " WHERE bus_id = " & TransFieldValueToString(pszVehicleID) _
                & " AND bus_date = " & TransFieldValueToString(pdtDate)
    oDB.Execute szSql
'    If lEffect = 0 Then
'        '出记录未找到的错
'        RaiseError ERR_VehicleFixFeeHasExist
'    End If
    
    
    Exit Sub
ErrorHandle:
    err.Raise err.Number, err.Source
    
End Sub

'修改车次固定费用
Public Sub EditBusFixFee(prsFixFee As Recordset)
    On Error GoTo ErrorHandle
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim i As Integer
    Dim lEffect As Long
    Dim oField As Field
    Dim szField As String
    Const cszProcName = "EditBusFixFee"
    
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_BusFixFee '验证用户是否有车次固定费用的权限
    
    oDB.ConnectionString = GetConnectionStr()
    
    If prsFixFee.RecordCount = 0 Then Exit Sub
    
    szSql = "SELECT * FROM fin_bus_fix_fee_lst  " _
        & " WHERE bus_id = " & TransFieldValueToString(prsFixFee!bus_id) _
        & " AND bus_date = " & TransFieldValueToString(prsFixFee!Bus_date)
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then
        '出车次固定费用已存在的错
        RaiseError ERR_BusFixFeeNotExist
    Else
        '如果费用已经扣过就不允许修改或删除
        
        If TransFieldValueToString(FormatDbValue(rsTemp!is_dec)) = TransFieldValueToString(True) Then
            RaiseError ERR_BusFixFeeHasDesCannotEdit
        End If
    End If
    
    For Each oField In prsFixFee.Fields
        Select Case LCase(oField.Name)
            Case "bus_id", "is_dec" '忽略更改
            Case "transport_company_id", "bus_date", "split_item_1", "split_item_2", "split_item_3", "split_item_4", "split_item_5", "split_item_6", "split_item_7" _
                , "split_item_8", "split_item_9", "split_item_10", "split_item_11", "split_item_12" _
                , "split_item_13", "split_item_14", "split_item_15", "split_item_16", "split_item_17" _
                , "split_item_18", "split_item_19", "split_item_20"
                
                
                
                szField = szField & "," & oField.Name & "=" & TransFieldValueToString(oField.Value)
            Case Else
                RaiseError "列名不正确", cszProcName
        End Select
        
    Next
    oDB.BeginTrans
        Do While Not prsFixFee.EOF
            szSql = "UPDATE fin_bus_fix_fee_lst SET bus_id=bus_id " _
                & szField
                
            szSql = szSql & " WHERE bus_id = " & TransFieldValueToString(prsFixFee!bus_id) _
                & " AND bus_date = " & TransFieldValueToString(prsFixFee!Bus_date)
                
            oDB.Execute szSql
'            If lEffect = 0 Then
'                '出记录未找到的错
'
'                RaiseError ERR_VehicleFixFeeHasExist
'            End If
            prsFixFee.MoveNext
        Loop
    oDB.CommitTrans
    Set oDB = Nothing
'    Set rsTemp = Nothing

    
    
    Exit Sub
UpdDriverErr:
    oDB.RollbackTrans
    Set oDB = Nothing
'    Set rsTemp = Nothing
    
    Call RaiseError(err.Number, "split:EditBusFixFee Method", err.Description)
    
    
    
    Exit Sub
ErrorHandle:
    err.Raise err.Number, err.Source
End Sub

Public Sub MakeSheetStatTemp(pdtBusDate As Date)
    '把路单统计数据生成到临时表中
    
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    
    On Error GoTo ErrorHandle
    AssertObjIsValid
'    AssertHaveRight m_oActiveUser, RIGHT_BusFixFee
    
    oDB.ConnectionString = GetConnectionStr()
    oDB.BeginTrans
        '删除路单统计的临时表
        szSql = " DELETE fin_sheet_station_lst_temp "
        oDB.Execute szSql
        '生成统计数据
        szSql = " EXEC STDayFinStat_SheetNormal_temp " & TransFieldValueToString(ToDBDate(pdtBusDate))
        oDB.Execute szSql
        szSql = " EXEC STDayFinStat_SheetMerge_temp " & TransFieldValueToString(ToDBDate(pdtBusDate))
        oDB.Execute szSql
        szSql = " EXEC STDayFinStat_SheetChange_temp " & TransFieldValueToString(ToDBDate(pdtBusDate))
        oDB.Execute szSql
    oDB.CommitTrans
    
    Exit Sub
ErrorHandle:
    err.Raise err.Number, err.Source, err.Description
    
    
End Sub

'汇款作废
Public Sub CancelRemit(paszSettleSheet() As String)
    '设置某些结算单汇款
    On Error GoTo ErrorHandle
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim i As Integer
    Dim szWhere As String
    Dim nCount As Integer
    Dim rsTemp As Recordset

    Dim lEffect As Long
    Dim dbNegative As Double
    
    AssertObjIsValid
    '需要加一个汇款作废的权限
    AssertHaveRight m_oActiveUser, RIGHT_CancelRemit
    
    oDB.ConnectionString = GetConnectionStr()
    
    nCount = ArrayLength(paszSettleSheet)
    
    For i = 1 To nCount - 1
        szWhere = szWhere & TransFieldValueToString(paszSettleSheet(i)) & " , "
    Next i
    
    szWhere = " ( " & szWhere & TransFieldValueToString(paszSettleSheet(i)) & " ) "
    
    
    '需判断一下,如果未结或作废,则不允许作废汇款
    szSql = " SELECT * FROM fin_settle_sheet_lst " _
        & " WHERE settlement_sheet_id IN " & szWhere
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount <> nCount Then
        '说明有几张结算单并不存在,出结算单不存在的错
        RaiseError ERR_SettleSheetNotExist
    Else
        For i = 1 To rsTemp.RecordCount
            If FormatDbValue(rsTemp!Status) = CS_SettleSheetInvalid Then
                '如果作废则不允许作废汇款
                RaiseError ERR_NotAllowCancelRemitHasCancel
            
            ElseIf FormatDbValue(rsTemp!Status) = CS_SettleSheetValid Then
                '如果未汇款(未结),则不允许作废汇款
                RaiseError ERR_NotAllowCancelRemitHasNotRemitted

            End If
            rsTemp.MoveNext
        Next i
    End If
    
    oDB.BeginTrans
        

        '状态设为未结
        szSql = " UPDATE fin_settle_sheet_lst SET status =  " & TransFieldValueToString(CS_SettleSheetValid) _
            & " , remit_date = '1900-01-01' " _
            & " WHERE settlement_sheet_id IN " & szWhere
        oDB.Execute szSql
        
    oDB.CommitTrans
    Exit Sub
    
ErrorHandle:
    oDB.RollbackTrans
    err.Raise err.Number, err.Source
    
End Sub




