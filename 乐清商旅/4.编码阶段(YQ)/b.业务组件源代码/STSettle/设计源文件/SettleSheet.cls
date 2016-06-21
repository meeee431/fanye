VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SettleSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E81EE4E035B"
'
Option Base 0
Option Explicit

'##ModelId=3E81EE5A0280
Private m_szOwnerID As String

'##ModelId=3E81EE6700FA
Private m_szSettleSheetID As String

'##ModelId=3E81EE79008C
Private m_nCheckSheetCount As Integer

'##ModelId=3E81EE7D02FD
Private m_dbTotalTicketPrice As Double

'##ModelId=3E81EE820157
Private m_dbTotalQuantity As Double

'##ModelId=3E81EE8602BF
Private m_dbSettleOtherCompanyPrice As Double

'##ModelId=3E81EE8A0242
Private m_dbSettleStationPrice As Double

'##ModelId=3E81EE8F0261
Private m_dbSettleLocalCompanyPrice As Double

'##ModelId=3E81EE9401E4
Private m_szObjectID As String

'##ModelId=3E81EE98006D
Private m_ESettleObject As ESettleObjectType

'##ModelId=3E81EE9C005D
Private m_szTransportCompanyID As String

'##ModelId=3E81EE9F031C
Private m_szTransportCompanyName As String

'##ModelId=3E81EEA50177
Private m_szSettler As String

'##ModelId=3E81EEA900DA
Private m_szChecker As String

'##ModelId=3E81EEAC02DE
Private m_dtSettleDate As Date

'##ModelId=3E81EEB0005D
Private m_dtSettleStartDate As Date

'##ModelId=3E81EEB4009C
Private m_dtSettleEndDate As Date

'##ModelId=3E81EEB7035B
Private m_szUnitName As String

'##ModelId=3E81EEBB02BF
Private m_EStatus As ESettleSheetStatus

'##ModelId=3E81EEBF031C
Private m_szAnnotation As String

'结算对象名称
Private m_szObjectName As String

Private m_szRouteID As String '线路代码
Private m_szRouteName As String '线路名称
Private m_bIsPrint As Integer  '是否该结算单已打印过,因为这些需要控制结算单,只能打印一次



Private m_szLuggageSettleIDs As String '结算单对应的行包结算单号
Private m_dbLuggageTotalBaseCarriage As Double '行包结算单的总运价
Private m_dbLuggageTotalSettlePrice As Double '行包结算单的总应拆金额
Private m_szLuggageProtocolName As String '行包结算单的结算协议

'local variable to hold the serialized class ID that was created in  Class_Initi-
'alize
'##ModelId=3E8219CE029F
Private mmlClassDebugID As Long
Private m_oActiveUser As ActiveUser
Private m_nObjectStatus As EObjectStatus '对象状态

'##ModelId=3E81EE650128
Public Sub Init(opAUser As ActiveUser)
    On Error GoTo InitErr

    Set m_oActiveUser = opAUser
    m_nObjectStatus = ST_NotAvailable
    Exit Sub
InitErr:
    Call RaiseError(MyUnhandledError, "SettleSheet:Init Method")
End Sub

'##ModelId=3E81EECD034B
Public Sub Identify(pszSettleSheet As String)
    m_szSettleSheetID = pszSettleSheet
    m_nObjectStatus = ST_NormalObj
    RefreshObject
End Sub
Private Sub RefreshObject()
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    AssertObjIsValid
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "SELECT * FROM fin_settle_sheet_lst WHERE settlement_sheet_id=" & TransFieldValueToString(m_szSettleSheetID)
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then RaiseError ERR_SettleSheetNotExist
    m_szAnnotation = FormatDbValue(rsTemp!Annotation)
    m_szChecker = FormatDbValue(rsTemp!Checker)
    m_szObjectID = FormatDbValue(rsTemp!object_id)
    m_szSettler = FormatDbValue(rsTemp!Settler)
    m_szSettleSheetID = FormatDbValue(rsTemp!settlement_sheet_id)
    m_szTransportCompanyID = FormatDbValue(rsTemp!transport_company_id)
    m_szTransportCompanyName = FormatDbValue(rsTemp!transport_company_name)
    m_szUnitName = FormatDbValue(rsTemp!unit_name)
    m_nCheckSheetCount = FormatDbValue(rsTemp!sheet_quantity)
    m_dbSettleOtherCompanyPrice = FormatDbValue(rsTemp!settle_price)
    m_dbSettleStationPrice = FormatDbValue(rsTemp!settle_station_price)
    m_dbTotalQuantity = FormatDbValue(rsTemp!total_quantity)
    m_dbTotalTicketPrice = FormatDbValue(rsTemp!total_ticket_price)
    m_dtSettleDate = FormatDbValue(rsTemp!settle_date)
    m_dtSettleEndDate = FormatDbValue(rsTemp!end_date)
    m_dtSettleStartDate = FormatDbValue(rsTemp!start_date)
    m_ESettleObject = FormatDbValue(rsTemp!settle_object)
    m_EStatus = FormatDbValue(rsTemp!Status)
    
    m_dbSettleLocalCompanyPrice = m_dbSettleOtherCompanyPrice - m_dbSettleStationPrice
    
    
    m_szObjectName = FormatDbValue(rsTemp!object_name)
    m_szLuggageSettleIDs = FormatDbValue(rsTemp!luggage_settle_id)
    m_dbLuggageTotalBaseCarriage = FormatDbValue(rsTemp!luggage_base_carriage)
    m_dbLuggageTotalSettlePrice = FormatDbValue(rsTemp!luggage_settle_price)
    
    m_szRouteID = FormatDbValue(rsTemp!route_id)
    m_szRouteName = FormatDbValue(rsTemp!route_name)
    
'    m_bIsPrint = IIf(FormatDbValue(rsTemp!is_print) = 0, False, True)
    m_bIsPrint = FormatDbValue(rsTemp!is_print)
    
    
End Sub

'##ModelId=3E81F074000F
Public Function GetCompanyInfo(Optional pszCompanyID As String = "") As TCompnaySettle()
    On Error GoTo GetCompanyInfoErr
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim szSqlTemp As String
    Dim TTemp() As TCompnaySettle
    Dim i As Integer, j As Integer
    AssertObjIsValid
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If pszCompanyID <> "" Then
        szSqlTemp = " AND transport_company_id = " & TransFieldValueToString(pszCompanyID)
    End If
    szSql = "SELECT f.*,p.protocol_name,c.transport_company_name " _
        & " FROM fin_settle_company_price_lst f,fin_settle_protocol_info p," & cszTableCompany & " c " _
        & " WHERE settlement_sheet_id =" & TransFieldValueToString(m_szSettleSheetID) _
        & " AND f.protocol_id = p.protocol_id AND c.transport_company_id =f.transport_company_id " _
        & szSqlTemp
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    For i = 1 To rsTemp.RecordCount
        TTemp(i).CompanyID = FormatDbValue(rsTemp!transport_company_id)
        TTemp(i).CompanyName = FormatDbValue(rsTemp!transport_company_name)
        TTemp(i).ProtocolID = FormatDbValue(rsTemp!protcol_id)
        TTemp(i).ProtocolName = FormatDbValue(rsTemp!protocl_name)
        TTemp(i).SettlementSheetID = FormatDbValue(rsTemp!settlement_sheet_id)
        For j = 1 To 20
            TTemp(i).SplitItem(j) = FormatDbValue(rsTemp("split_item_" & Trim(Str(j))))
        Next j
    Next i
    GetCompanyInfo = TTemp
    Exit Function
GetCompanyInfoErr:
    Call RaiseError(MyUnhandledError, "SettleSheet:GetCompanyInfo Method")
End Function

'##ModelId=3E81EED5003E
Public Function GetVehicleInfo(Optional pszVehicleID As String = "") As TVehilceSettle()
    On Error GoTo GetVehicleInfoErr
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim szSqlTemp As String
    Dim TTemp() As TVehilceSettle
    Dim i As Integer, j As Integer
    AssertObjIsValid
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If pszVehicleID <> "" Then
        szSqlTemp = " AND vehicle_id = " & TransFieldValueToString(pszVehicleID)
    End If
    szSql = "SELECT f.*,vi.license_tag_no,p.protocol_name,v.vehilce_type_code " _
        & " FROM fin_settle_vehicle_lst f," & cszTableVehicleTypeCode & " v,fin_settle_protocol_info p,vehilce_info vi" _
        & " WHERE settlement_sheet_id =" & TransFieldValueToString(m_szSettleSheetID) _
        & " AND vi.vehilce_id = f.vehilce_id " _
        & " AND p.protocol_id = f.protocol_id AND v.vehicle_type_code = vi.vehicle_type_code " & szSqlTemp
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    For i = 1 To rsTemp.RecordCount
        TTemp(i).VehicleID = FormatDbValue(rsTemp!vehicle_id)
        TTemp(i).VehicleTypeCode = FormatDbValue(rsTemp!vehicle_type_code)
        TTemp(i).LicenseTagNo = FormatDbValue(rsTemp!license_tag_no)
        TTemp(i).ProtocolID = FormatDbValue(rsTemp!protcol_id)
        TTemp(i).ProtocolName = FormatDbValue(rsTemp!protocl_name)
        TTemp(i).SettlementSheetID = FormatDbValue(rsTemp!settlement_sheet_id)
        For j = 1 To 20
            TTemp(i).SplitItem(j) = FormatDbValue(rsTemp("split_item_" & Trim(Str(j))))
        Next j
    Next i
    GetVehicleInfo = TTemp
    Exit Function
GetVehicleInfoErr:
    Call RaiseError(MyUnhandledError, "SettleSheet:GetVehicleInfo Method")
End Function

'##ModelId=3E81EED9003E
Public Function GetSettleSheetStationInfo() As TSettleSheetStation()
    AssertActiveUserValid m_oActiveUser, ERR_SettleSheet
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim szSqlTemp As String
    Dim aTTemp() As TSettleSheetStation
    Dim i As Integer, j As Integer
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "SELECT f.settlement_sheet_id,f.sell_station_id,ss.station_name sell_station_name , f.route_id,r.route_name,f.station_id,s.station_name," _
        & " f.ticket_type_id,t.ticket_type_name,f.quantity , a.annotation " _
        & " FROM fin_settle_station_lst f," & cszTableRouteInfo & " r ," & cszTableStationInfo & " s ," & cszTableStationInfo & " ss," _
        & cszTableTicketType & " t , area_code a , sell_station_info z " _
        & " WHERE f.route_id =r.route_id AND f.sell_station_id = z.sell_station_id AND z.station_id = ss.station_id " _
        & " AND f.station_id = s.station_id AND f.ticket_type_id = t.ticket_type_id AND a.area_code = s.area_code  "
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    ReDim TTemp(1 To rsTemp.RecordCount)
    For i = 1 To rsTemp.RecordCount
        aTTemp(i).RouteID = FormatDbValue(rsTemp!route_id)
        aTTemp(i).RouteName = FormatDbValue(rsTemp!route_name)
        aTTemp(i).SellSationID = FormatDbValue(rsTemp!sell_station_id)
        aTTemp(i).SellStationName = FormatDbValue(rsTemp!sell_station_name)
        aTTemp(i).SettlementSheetID = FormatDbValue(rsTemp!settlement_sheet_id)
        aTTemp(i).StationID = FormatDbValue(rsTemp!station_id)
        aTTemp(i).StationName = FormatDbValue(rsTemp!station_name)
        aTTemp(i).TicketType = FormatDbValue(rsTemp!ticket_type_id)
        aTTemp(i).TicketTypeName = FormatDbValue(rsTemp!ticket_type_name)
        aTTemp(i).Quantity = FormatDbValue(rsTemp!Quantity)
        aTTemp(i).AreaRatio = Val(FormatDbValue(rsTemp!Annotation))
    Next i
    GetSettleSheetStationInfo = aTTemp
    

    Exit Function
GetSettleSheetStationInfoErr:
    Call RaiseError(MyUnhandledError, "SettleSheet:GetSettleSheetStationInfo Method")
End Function

'##ModelId=3E81EEDD008C
Public Function GetCheckSheetInfo() As TCheckSheet()
    On Error GoTo GetCheckSheetInfoErr
'    Dim oDb As New RTConnection
'    Dim szSql As String
'    Dim rsTemp As Recordset
'    Dim szSqlTemp As String
'    Dim TTemp() As TCheckSheet
'    Dim i As Integer, j As Integer
'    oDb.ConnectionString = GetConnectionStr(cszSettle)
'    szSql = "SELECT * " _
'        & " FROM " & cszTableCheckSheet & " s," & cszTableRouteInfo & " r,sell_station_info s," & cszTableStationInfo & " si," & cszTableOwnerInfo & " o," & cszTableCompany & " c,vehilce_info v " _
'        & " WHERE "
'
'    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount = 0 Then Exit Function
'    For i = 1 To rsTemp.RecordCount
'        TTemp(i).Availability = FormatDbValue(rsTemp!route_id)
'        TTemp(i).BusType = FormatDbValue(rsTemp!route_name)
'        TTemp(i).dbCheckCharge = FormatDbValue(rsTemp!sell_station_id)
'        TTemp(i).dbSelletPrice = FormatDbValue(rsTemp!sell_station_name)
'        TTemp(i).dbTotalPrice = FormatDbValue(rsTemp!settlement_sheet_id)
'        TTemp(i).dtBusDate = FormatDbValue(rsTemp!station_id)
'        TTemp(i).dtEndDate = FormatDbValue(rsTemp!station_name)
'        TTemp(i).dtStartTime = FormatDbValue(rsTemp!ticket_type_id)
'        TTemp(i).nBusSerial = FormatDbValue(rsTemp!ticket_type_name)
'        TTemp(i).nQuantity = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).nStatus = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).szBusID = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).szChecker = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).szCheckGateID = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).szCheckGateName = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).szCheckSheetID = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).szCompanyID = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).szCompanyName = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).szLicenseTagNo = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).szOwnerID = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).szOwnerName = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).szProtocol = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).szRouteID = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).szRouteName = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).szSellStationID = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).szSellStationName = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).szSettleSheetDate = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).szSettleSheetID = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).szSettleUnit = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).szSplitCompanyID = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).szSplitCompanyName = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).szVehicleID = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).tdStartDate = FormatDbValue(rsTemp!Quantity)
'        TTemp(i).vData = FormatDbValue(rsTemp!Quantity)
'    Next i
'    GetCheckSheetInfo = TTemp
    Exit Function
GetCheckSheetInfoErr:
    Call RaiseError(MyUnhandledError, "SettleSheet:GetCheckSheetInfo Method")
End Function

'##ModelId=3E81EEE0034B
Public Sub CancelSettleSheet(szSettleSheetID As String, Optional pszStatus As String)
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    Dim szSql As String
    On Error GoTo ErrorHandle
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_CancelSettleSheet
    '************************
    '此接口以后需要修改
    m_nObjectStatus = ST_NormalObj
    '************************
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    '查询有无该结算单
    szSql = " SELECT * FROM fin_settle_sheet_lst " _
        & " WHERE settlement_sheet_id = " & TransFieldValueToString(szSettleSheetID)
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then RaiseError ERR_SettleSheetNotExist
    
    If FormatDbValue(rsTemp!Status) = CS_SettleSheetInvalid Then
        '如果已作废,则出错
        RaiseError ERR_SettleSheetHasCanceled
        
    ElseIf FormatDbValue(rsTemp!Status) = CS_SettleSheetSettled Then
        '如果已结算,则出错
        RaiseError ERR_SettleSheetHasSettled
    End If
    oDB.BeginTrans
        
        szSql = "UPDATE fin_settle_sheet_lst SET status = " & CS_SettleSheetInvalid _
            & " WHERE settlement_sheet_id = " & TransFieldValueToString(szSettleSheetID)   '将"路单结算单表"中的状态改为作废
        oDB.Execute szSql
        
        
        
        
        
        '将"路单信息表"中的将所有该结算的路单的结算单号、拆算协议、结算时间还原空,但状态为正在结算状态。
        szSql = "UPDATE " & cszTableCheckSheet & " " _
            & " SET settlement_status = " & TransFieldValueToString(CS_CheckSheetSettling) & " , " _
            & " settlement_sheet_id ='' ," _
            & " protocol_id ='' , " _
            & " settlement_time = " & TransFieldValueToString(cszEmptyDateStr) _
            & " WHERE check_sheet_id IN ( " _
            & " SELECT check_sheet_id FROM fin_settle_bus_station_lst WHERE settlement_sheet_id = " & TransFieldValueToString(szSettleSheetID) & " ) "
        oDB.Execute szSql
        
        '还原路单站点表中的人数
        szSql = " UPDATE " & cszTableSheetStation & " " _
            & " SET fact_quantity = fact_quantity - s.quantity " _
            & " FROM " & cszTableSheetStation & " t , fin_settle_bus_station_lst s , sell_station_info i  " _
            & " WHERE s.settlement_sheet_id = " & TransFieldValueToString(szSettleSheetID) _
            & " AND t.check_sheet_id = s.check_sheet_id " _
            & " AND t.sell_station_id = i.sell_station_id AND s.sell_station_id = i.station_id " _
            & " AND t.station_id = s.station_id " _
            & " AND t.price_identify = s.price_identify " _
            & " AND t.ticket_type = s.ticket_type " _
            & " AND t.seat_type_id = s.seat_type_id "
            
        oDB.Execute szSql
        
        If pszStatus = "车辆" Then
        '还原车辆固定费用的信息
        szSql = " UPDATE fin_vehicle_fix_fee_lst " _
            & " SET is_dec = " & TransFieldValueToString(False) _
            & "     , settlement_sheet_id = " & TransFieldValueToString("") _
            & " WHERE settlement_sheet_id = " & TransFieldValueToString(szSettleSheetID) _
            
        oDB.Execute szSql
        ElseIf pszStatus = "车次" Then
        szSql = " UPDATE fin_bus_fix_fee_lst " _
            & " SET is_dec = " & TransFieldValueToString(False) _
            & "     , settlement_sheet_id = " & TransFieldValueToString("") _
            & " WHERE settlement_sheet_id = " & TransFieldValueToString(szSettleSheetID) _
            
        oDB.Execute szSql
        End If
        
        
    oDB.CommitTrans
    WriteOperateLog m_oActiveUser, RIGHT_CancelSettleSheet, szSettleSheetID & "结算单作废"
    
    Exit Sub
ErrorHandle:
    oDB.RollbackTrans
    err.Raise err.Number, err.Source, err.Description
    
    
    
End Sub


'##ModelId=3E822941003E
Public Property Get mlClassDebugID() As Long
    mlClassDebugID = mmlClassDebugID
End Property

'##ModelId=3E82294002EE
Public Property Get Annotation() As String
    Annotation = m_szAnnotation
End Property

'##ModelId=3E82294001B5
Public Property Get Status() As ESettleSheetStatus
    Status = m_EStatus
End Property

'##ModelId=3E822940008C
Public Property Get UnitName() As String
    UnitName = m_szUnitName
End Property

'##ModelId=3E82293F034B
Public Property Get SettleEndDate() As Date
    SettleEndDate = m_dtSettleEndDate
End Property

'##ModelId=3E82293F0213
Public Property Get SettleStartDate() As Date
    SettleStartDate = m_dtSettleStartDate
End Property

'##ModelId=3E82293F00FA
Public Property Get SettleDate() As Date
    SettleDate = m_dtSettleDate
End Property

'##ModelId=3E82293E03C8
Public Property Get Checker() As String
    Checker = m_szChecker
End Property

'##ModelId=3E82293E02BF
Public Property Get Settler() As String
    Settler = m_szSettler
End Property

'##ModelId=3E82293E0196
Public Property Get TransportCompanyName() As String
    TransportCompanyName = m_szTransportCompanyName
End Property

'##ModelId=3E82293E008C
Public Property Get TransportCompanyID() As String
    TransportCompanyID = m_szTransportCompanyID
End Property

'##ModelId=3E82293D036B
Public Property Get SettleObject() As ESettleObjectType
    SettleObject = m_ESettleObject
End Property

'##ModelId=3E82293D0242
Public Property Get ObjectID() As String
    ObjectID = m_szObjectID
End Property

Public Property Get ObjectName() As String
'    Dim szSql As String
'    Dim oDb As New RTConnection
'    Dim rsTemp As Recordset
'    oDb.ConnectionString = GetConnectionStr
'
'    If m_ESettleObject = CS_SettleByTransportCompany Then
'        szSql = " SELECT transport_company_short_name object_name FROM " & cszTableCompany & " WHERE transport_company_id = '" & m_szObjectID & "'"
'    ElseIf m_ESettleObject = CS_SettleByVehicle Then
'        szSql = "SELECT license_tag_no object_name FROM " & cszTableVehicle & " WHERE vehicle_id = '" & m_szObjectID & "'"
'    End If
'
'    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount > 0 Then ObjectName = FormatDbValue(rsTemp!object_name)
    ObjectName = m_szObjectName
    
End Property


'##ModelId=3E82293D0128
Public Property Get SettleLocalCompanyPrice() As Double
    SettleLocalCompanyPrice = m_dbSettleLocalCompanyPrice
End Property

'##ModelId=3E82293D002E
Public Property Get SettleStationPrice() As Double
    SettleStationPrice = m_dbSettleStationPrice
End Property

'##ModelId=3E82293C030D
Public Property Get SettleOtherCompanyPrice() As Double
    SettleOtherCompanyPrice = m_dbSettleOtherCompanyPrice
End Property

'##ModelId=3E82293C0203
Public Property Get TotalQuantity() As Double
    TotalQuantity = m_dbTotalQuantity
End Property

'##ModelId=3E82293C0148
Public Property Get TotalTicketPrice() As Double
    TotalTicketPrice = m_dbTotalTicketPrice
End Property

'##ModelId=3E82293C005D
Public Property Get CheckSheetCount() As Integer
    CheckSheetCount = m_nCheckSheetCount
End Property

'##ModelId=3E82293B038A
Public Property Get SettleSheetID() As String
    SettleSheetID = m_szSettleSheetID
End Property


'##ModelId=3E82293B02AF
Public Property Get OwnerID() As String
    OwnerID = m_szOwnerID
End Property


Public Property Get LuggageSettleIDs() As String
    LuggageSettleIDs = m_szLuggageSettleIDs
End Property



Public Property Get LuggageTotalBaseCarriage() As Double
    LuggageTotalBaseCarriage = m_dbLuggageTotalBaseCarriage
End Property


Public Property Get LuggageTotalSettlePrice() As Double
    LuggageTotalSettlePrice = m_dbLuggageTotalSettlePrice
End Property


Public Property Get LuggageProtocolName() As String
    LuggageProtocolName = m_szLuggageProtocolName
End Property


Public Property Get RouteID() As String
    RouteID = m_szRouteID
End Property


Public Property Get RouteName() As String
    RouteName = m_szRouteName
End Property

Public Property Get IsPrint() As Integer
    IsPrint = m_bIsPrint
End Property



'**************************************************
'Member Code:S3
'Brief Description:测试对象是否有效
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then RaiseError ERR_SettleSheetNotAvailable
End Sub

'**************************************************
'Member Code:S4
'Brief Description:测试对象是否有效
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub AssertObjIsValid()
    AssertActiveUserValid m_oActiveUser, ERR_SettleSheet
    AssertStatusAvailable
End Sub




Public Sub SetPrint()
    '设置该结算单已打印
    Dim szSql As String
    Dim oDB As New RTConnection
    AssertObjIsValid
    
    On Error GoTo ErrorHandle
    oDB.ConnectionString = GetConnectionStr
    
    szSql = " UPDATE fin_settle_sheet_lst SET is_print =is_print+1" _
        & " WHERE settlement_sheet_id = " & TransFieldValueToString(m_szSettleSheetID)
    oDB.Execute szSql
    Exit Sub
ErrorHandle:
    err.Raise err.Number, err.Source, err.Description
    
End Sub




