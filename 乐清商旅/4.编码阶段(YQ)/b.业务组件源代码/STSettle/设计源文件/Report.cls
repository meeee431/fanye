VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E81F6A802AF"
'
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 1
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3E8219D10222
    Private mlClassDebugID As Long
#End If
Private m_oActiveUser As ActiveUser
Private m_nObjectStatus As EObjectStatus '对象状态


'##ModelId=3E81F6F600DA
Public Sub Init(opAUser As ActiveUser)
    On Error GoTo InitErr

    Set m_oActiveUser = opAUser
    m_nObjectStatus = ST_NormalObj
    Exit Sub
InitErr:
    Call RaiseError(MyUnhandledError, "Report:Init Method")
End Sub

'##ModelId=3E81F6FB008C
Public Function GetSplitItemInfo(Optional pszSplitItemID As String = "", Optional pbSplitStatus As Boolean = False) As TSplitItemInfo()
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szTemp As String
    Dim aszTemp() As TSplitItemInfo, i As Integer
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_SplitItem
    
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If pszSplitItemID <> "" Then
        szTemp = " AND split_item_id = " & TransFieldValueToString(pszSplitItemID)
    End If
    
    If pbSplitStatus Then
        '只查询可用项
        szTemp = szTemp & " AND split_status = " & TransFieldValueToString(CS_SplitItemUse)
    End If
    
    szSql = "SELECT * FROM split_item_info WHERE 1=1 " & szTemp
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    ReDim aszTemp(1 To rsTemp.RecordCount)
    For i = 1 To rsTemp.RecordCount
        aszTemp(i).SplitItemID = FormatDbValue(rsTemp!split_item_id)
        aszTemp(i).SplitItemName = FormatDbValue(rsTemp!split_item_name)
        aszTemp(i).SplitStatus = FormatDbValue(rsTemp!split_status)
        aszTemp(i).SplitType = FormatDbValue(rsTemp!split_type)
        aszTemp(i).AllowModify = FormatDbValue(rsTemp!allow_modify)
        rsTemp.MoveNext
    Next i
    GetSplitItemInfo = aszTemp

End Function
 

 Public Function GetSplitItemInfoRs() As ADODB.Recordset
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_SplitItem
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "SELECT * FROM split_item_info WHERE 1=1 "
    Set rsTemp = oDB.Execute(szSql)
    Set GetSplitItemInfoRs = rsTemp
End Function

'得到固定费用项
 Public Function GetSplitItemInfoRs2(g_szFixFeeItem As String) As ADODB.Recordset
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim i, nCols As Integer
    Dim m_aszFixFeeItem() As String
    Dim szWhere As String
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_SplitItem
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    
    m_aszFixFeeItem = Split(g_szFixFeeItem, ",")
    nCols = ArrayLength(m_aszFixFeeItem)
    
    For i = 0 To nCols - 1
        If Len(m_aszFixFeeItem(i)) = 1 Then
            m_aszFixFeeItem(i) = "000" & m_aszFixFeeItem(i)
        Else
            m_aszFixFeeItem(i) = "00" & m_aszFixFeeItem(i)
        End If
    Next i
    
    For i = 0 To nCols - 2
        szWhere = szWhere & " f.split_item_id = " & TransFieldValueToString(m_aszFixFeeItem(i)) & " or"
    Next i
    
    If i = nCols - 1 Then
        szWhere = szWhere & " f.split_item_id = " & TransFieldValueToString(m_aszFixFeeItem(i))
    End If
     If szWhere = "" Then RaiseError ERR_NoFixFeeItem
    szSql = "SELECT * FROM split_item_info as f WHERE 1=1  " & "and" & "(" & szWhere & ")"
    Set rsTemp = oDB.Execute(szSql)
    Set GetSplitItemInfoRs2 = rsTemp
End Function
 
'##ModelId=3E81F6FF03C8
Public Function GetAllFormula(Optional pszFormulaID As String = "") As String()
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szTemp As String
    Dim aszTemp() As String, i As Integer
    AssertObjIsValid

    AssertHaveRight m_oActiveUser, RIGHT_GetFormula
    
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If pszFormulaID <> "" Then
        szTemp = " WHERE formuar_id = " & TransFieldValueToString(pszFormulaID)
    End If
    szSql = "SELECT * FROM settle_formula_info " & szTemp
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 3)
    For i = 1 To rsTemp.RecordCount
        aszTemp(i, 1) = FormatDbValue(rsTemp!formuar_id)
        aszTemp(i, 2) = FormatDbValue(rsTemp!formular_name)
        aszTemp(i, 3) = FormatDbValue(rsTemp!formular_content)
        rsTemp.MoveNext
    Next i
    GetAllFormula = aszTemp
End Function

Public Function GetAllProtocol(Optional pszProtocolID As String = "") As String()
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szTemp As String
    Dim aszTemp() As String, i As Integer
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_GetProtocol
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If pszProtocolID <> "" Then
        szTemp = " WHERE protocol_id = " & TransFieldValueToString(pszProtocolID)
    End If
    szSql = "SELECT * FROM fin_settle_protocol_info  " & szTemp
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 4)
    For i = 1 To rsTemp.RecordCount
        aszTemp(i, 1) = FormatDbValue(rsTemp!protocol_id)
        aszTemp(i, 2) = FormatDbValue(rsTemp!protocol_name)
        aszTemp(i, 3) = FormatDbValue(rsTemp!Annotation)
        aszTemp(i, 4) = FormatDbValue(rsTemp!default_mark)
        rsTemp.MoveNext
    Next i
    GetAllProtocol = aszTemp
End Function


Public Function GetAllProtocolRS(Optional pszProtocolID As String = "") As Recordset
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szTemp As String
    Dim aszTemp() As String, i As Integer
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_GetProtocol
    
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If pszProtocolID <> "" Then
        szTemp = " AND protocol_id = " & TransFieldValueToString(pszProtocolID)
    End If
    szSql = "SELECT i.* , t.split_item_name , p.protocol_name" _
        & " FROM fin_settle_protocol_info p ,fin_charge_item_info i , split_item_info t " _
        & " WHERE p.protocol_id = i.protocol_id  AND i.split_item_id=t.split_item_id " _
        & " AND t.split_status = " & TransFieldValueToString(ESplitItemUseStatus.CS_SplitItemUse) _
        & szTemp
    
    Set rsTemp = oDB.Execute(szSql)
    
    Set GetAllProtocolRS = rsTemp
        
End Function

Public Function GetAllCompanyProtocolRS(Optional CompanyID As String = "", Optional RouteID As String = "", Optional ProtocolID As String = "") As Recordset
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim aszTemp() As String
    Dim szTemp As String, i As Integer
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_GetCompanyProtocol
    
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If Trim(CompanyID) <> "" Then
        szTemp = " AND f.transport_company_id = " & TransFieldValueToString(CompanyID)
    End If
    If Trim(ProtocolID) <> "" Then
        szTemp = szTemp + " AND f.protocol_id = " & TransFieldValueToString(ProtocolID)
    End If
    If Trim(RouteID) <> "" Then
        szTemp = szTemp + " AND f.route_id = " & TransFieldValueToString(RouteID)
    End If
    szSql = "SELECT c.transport_company_id ,c.transport_company_short_name ,r.route_id,r.route_name ,f.protocol_id ,p.protocol_name" _
        & " FROM fin_compnay_protocol_lst f , route_info r, " & cszTableCompany & " c ,fin_settle_protocol_info p" _
        & " WHERE f.transport_company_id = c.transport_company_id AND f.route_id =r.route_id AND f.protocol_id = p.protocol_id " & szTemp
    Set rsTemp = oDB.Execute(szSql)
    Set GetAllCompanyProtocolRS = rsTemp
    
End Function


Public Function GetVehicleProtocolRS(Optional pszVehicleID As String = "", Optional pszCompanyID As String = "", Optional pszProtocolID As String = "") As Recordset

    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim aszTemp() As String
    Dim szTemp As String, i As Integer
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_GetVehicleProtocol
    
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If pszVehicleID <> "" Then
        szTemp = " AND v.vehicle_id = " & TransFieldValueToString(pszVehicleID)
    End If
    If pszProtocolID <> "" Then
        szTemp = szTemp & " AND v.protocol_id = " & TransFieldValueToString(pszProtocolID)
    End If
    If pszCompanyID <> "" Then
        szTemp = szTemp & " AND v.transport_company_id = " & TransFieldValueToString(pszCompanyID)
    End If
    szSql = "SELECT ISNULL( v.protocol_id,'') protocol_id ,ISNULL(p.protocol_name,'') protocol_name , " _
        & " v.transport_company_id company_id,c.transport_company_short_name company_name,ISNULL( p.default_mark,0) default_mark,v.vehicle_id,v.license_tag_no" _
        & " FROM " & cszTableVehicle & " v LEFT JOIN fin_settle_protocol_info p ON v.protocol_id=p.protocol_id " _
        & " JOIN " & cszTableCompany & " c ON c.transport_company_id=v.transport_company_id WHERE 1=1 " & szTemp
    Set rsTemp = oDB.Execute(szSql)
    Set GetVehicleProtocolRS = rsTemp
    
End Function

Public Function GetCompanySettlePriceLstRS(Optional TransportCompanyID As String = "", Optional VehicleType As String = "", _
                                 Optional RouteID As String = "", Optional SellstationID As String = "", Optional StationID As String = "") As Recordset
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szTemp As String
    Dim sztemp1 As String, i As Integer
    Dim aszTemp() As TCompanySettlePrice
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_GetCompanySettlePrice
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If TransportCompanyID <> "" Then
        szTemp = " AND f.transport_company_id = " & TransFieldValueToString(TransportCompanyID)
    Else
        szTemp = ""
    End If
    If VehicleType <> "" Then
        szTemp = szTemp + " AND f.vehicle_type_code  = " & TransFieldValueToString(VehicleType)
    Else
        szTemp = szTemp + ""
    End If
    If RouteID <> "" Then
        szTemp = szTemp + " AND f.route_id = " & TransFieldValueToString(RouteID)
    Else
        szTemp = szTemp + ""
    End If
    If SellstationID <> "" Then
        szTemp = szTemp + " AND f.sell_station_id = " & TransFieldValueToString(SellstationID)
    Else
        szTemp = szTemp + ""
    End If
    If StationID <> "" Then
        szTemp = szTemp + " AND f.station_id = " & TransFieldValueToString(StationID)
    Else
        szTemp = szTemp + ""
    End If
    szSql = "SELECT f.transport_company_id ,c.transport_company_short_name,f.vehicle_type_code ,vt.vehicle_type_short_name, " _
        & " f.route_id,r.route_name,f.sell_station_id,s.station_name sell_station_name,f.station_id,si.station_name ,f.mileage,f.pass_charge, " _
        & " f.settle_full_price,f.settle_half_price,f.annotation ,f.halve_full_price,f.halve_half_price,f.service_full_price," _
        & " f.service_half_price ,f.spring_full_price,f.spring_half_price " _
        & " FROM fin_company_settle_price_lst f ," & cszTableCompany & " c ," & cszTableVehicleTypeCode & " vt," & cszTableRouteInfo _
        & " r ," & cszTableStationInfo & " s ," & cszTableStationInfo & " si " _
        & " WHERE f.transport_company_id =c.transport_company_id AND f.vehicle_type_code =vt.vehicle_type_code " _
        & " AND r.route_id=f.route_id AND f.sell_station_id=s.station_id " _
        & " AND f.station_id = si.station_id " & szTemp _
        & " ORDER BY f.sell_station_id , f.transport_company_id , f.route_id ,f.vehicle_type_code , f.station_id "
    Set rsTemp = oDB.Execute(szSql)
    Set GetCompanySettlePriceLstRS = rsTemp
    
    
End Function

Public Function GetVehicleSettlePriceLstRS(Optional VehicleID As String = "", Optional CompanyID As String = "", Optional RouteID As String = "", Optional SellstationID As String = "", Optional StationID As String = "") As Recordset
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szTemp As String
    Dim sztemp1 As String, i As Integer
    
    Dim aszTemp() As TVehcileSettlePrice
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_GetVehicleSettlePrice
    
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If VehicleID <> "" Then
        szTemp = " AND f.vehicle_id = " & TransFieldValueToString(VehicleID)
    Else
        szTemp = ""
        
    End If
    If RouteID <> "" Then
        szTemp = szTemp + " AND f.route_id = " & TransFieldValueToString(RouteID)
    Else
        szTemp = szTemp + ""
    End If
    If SellstationID <> "" Then
        szTemp = szTemp + " AND f.sell_station_id = " & TransFieldValueToString(SellstationID)
    Else
        szTemp = szTemp + ""
    End If
    If StationID <> "" Then
        szTemp = szTemp + " AND f.station_id = " & TransFieldValueToString(StationID)
    Else
        szTemp = szTemp + ""
    End If
    If CompanyID <> "" Then
        szTemp = szTemp + " AND v.transport_company_id  = " & TransFieldValueToString(CompanyID)
    Else
        szTemp = szTemp + ""
    End If
    szSql = "SELECT f.vehicle_id ,f.route_id,r.route_name,f.sell_station_id,s.station_name sell_station_name,f.station_id,si.station_name , " _
        & " f.mileage,f.pass_charge,f.settle_full_price,f.settle_half_price,f.annotation ,v.license_tag_no ,f.halve_full_price, " _
        & " f.halve_half_price,f.service_full_price,f.service_half_price,f.spring_full_price,f.spring_half_price " _
        & " FROM fin_vehicle_settle_price_lst f ," & cszTableVehicle & " v," & cszTableRouteInfo & " r ," & cszTableStationInfo & " s ," & cszTableStationInfo & " si " _
        & " WHERE f.vehicle_id=v.vehicle_id AND r.route_id=f.route_id AND f.sell_station_id=s.station_id " _
        & " AND f.station_id = si.station_id " & szTemp _
        & " ORDER BY f.sell_station_id , f.route_id , f.vehicle_id , f.station_id "
    Set rsTemp = oDB.Execute(szSql)
    Set GetVehicleSettlePriceLstRS = rsTemp
End Function



'##ModelId=3E81F70500BB
Public Function GetAllHalveCompany(Optional RouteID As String = "", Optional CompanyID As String = "", Optional CompnayIDOther As String = "") As String()
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szTemp As String
    Dim aszTemp() As String, i As Integer
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_GetHalveCompany
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If RouteID <> "" Then
        szTemp = " AND f.route_id = " & TransFieldValueToString(RouteID)
    Else
        szTemp = ""
    End If
    If CompanyID <> "" Then
        szTemp = szTemp + " AND f.transport_company_id = " & TransFieldValueToString(CompanyID)
    Else
        szTemp = szTemp + ""
    End If
    If CompnayIDOther <> "" Then
        szTemp = szTemp + "  AND f.other_transport_company_id = " & TransFieldValueToString(CompnayIDOther)
    Else
        szTemp = szTemp + ""
    End If
    szSql = "SELECT f.route_id ,r.route_name,f.transport_company_id ,f.ratio , c.transport_company_short_name,f.other_transport_company_id,c1.transport_company_short_name other_transport_company_name " _
        & " FROM fin_halve_company_lst f," & cszTableRouteInfo & " r, " & cszTableCompany & " c," & cszTableCompany & " c1 WHERE f.route_id=r.route_id AND f.transport_company_id = c.transport_company_id AND f.other_transport_company_id = c1.transport_company_id " & szTemp
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 7)
    For i = 1 To rsTemp.RecordCount
        aszTemp(i, 1) = FormatDbValue(rsTemp!route_id)
        aszTemp(i, 2) = FormatDbValue(rsTemp!route_name)
        aszTemp(i, 3) = FormatDbValue(rsTemp!transport_company_id)
        aszTemp(i, 4) = FormatDbValue(rsTemp!transport_company_short_name)
        aszTemp(i, 5) = FormatDbValue(rsTemp!other_transport_company_id)
        aszTemp(i, 6) = FormatDbValue(rsTemp!other_transport_company_name)
        aszTemp(i, 7) = FormatDbValue(rsTemp!Ratio)
        rsTemp.MoveNext
    Next i
    GetAllHalveCompany = aszTemp
End Function

'##ModelId=3E81F7090222
Public Function GetCompanySettlePriceLst(Optional TransportCompanyID As String = "", Optional VehicleType As String = "", _
                                 Optional RouteID As String = "", Optional SellstationID As String = "", Optional StationID As String = "") As TCompanySettlePrice()
Attribute GetCompanySettlePriceLst.VB_Description = "slfjklsdjf"
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szTemp As String
    Dim sztemp1 As String, i As Integer
    Dim aszTemp() As TCompanySettlePrice
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_GetCompanySettlePrice
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If TransportCompanyID <> "" Then
        szTemp = " AND f.transport_company_id = " & TransFieldValueToString(TransportCompanyID)
    Else
        szTemp = ""
    End If
    If VehicleType <> "" Then
        szTemp = szTemp + " AND f.vehicle_type_code  = " & TransFieldValueToString(VehicleType)
    Else
        szTemp = szTemp + ""
    End If
    If RouteID <> "" Then
        szTemp = szTemp + " AND f.route_id = " & TransFieldValueToString(RouteID)
    Else
        szTemp = szTemp + ""
    End If
    If SellstationID <> "" Then
        szTemp = szTemp + " AND f.sell_station_id = " & TransFieldValueToString(SellstationID)
    Else
        szTemp = szTemp + ""
    End If
    If StationID <> "" Then
        szTemp = szTemp + " AND f.station_id = " & TransFieldValueToString(StationID)
    Else
        szTemp = szTemp + ""
    End If
    szSql = "SELECT f.transport_company_id ,c.transport_company_short_name,f.vehicle_type_code ,vt.vehicle_type_short_name, " _
        & " f.route_id,f.route_name,f.sell_station_id,s.station_name sell_station_name,f.station_id,si.station_name ,f.mileage,f.pass_charge, " _
        & " f.settle_full_price,f.settle_half_price,f.annotation ,f.halve_full_price,f.halve_half_price,f.service_full_price, " _
        & " f.service_half_price ,f.spring_full_price,f.spring_half_price " _
        & " FROM fin_company_settle_price_lst f ," & cszTableCompany & " c ," & cszTableVehicleTypeCode & " vt," _
        & cszTableStationInfo & " s ," & cszTableStationInfo & " si " _
        & " WHERE f.transport_company_id =c.transport_company_id AND f.vehicle_type_code =vt.vehicle_type_code  " _
        & " AND f.sell_station_id=s.station_id " _
        & " AND f.station_id = si.station_id " & szTemp _
        & " ORDER BY f.sell_station_id , f.transport_company_id , f.route_id ,f.vehicle_type_code , f.mileage "
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    ReDim aszTemp(1 To rsTemp.RecordCount)
    For i = 1 To rsTemp.RecordCount
        aszTemp(i).RouteID = FormatDbValue(rsTemp!route_id)
        aszTemp(i).RouteName = FormatDbValue(rsTemp!route_name)
        aszTemp(i).CompanyID = FormatDbValue(rsTemp!transport_company_id)
        aszTemp(i).CompanyName = FormatDbValue(rsTemp!transport_company_short_name)
        aszTemp(i).Mileage = FormatDbValue(rsTemp!Mileage)
        aszTemp(i).Annotation = FormatDbValue(rsTemp!Annotation)
        aszTemp(i).PassCharge = FormatDbValue(rsTemp!pass_charge)
        aszTemp(i).SellstationID = FormatDbValue(rsTemp!sell_station_id)
        aszTemp(i).SellStationName = FormatDbValue(rsTemp!sell_station_name)
        aszTemp(i).SettlefullPrice = FormatDbValue(rsTemp!settle_full_price)
        aszTemp(i).SettleHalfPrice = FormatDbValue(rsTemp!settle_half_price)
        aszTemp(i).StationID = FormatDbValue(rsTemp!station_id)
        aszTemp(i).StationName = FormatDbValue(rsTemp!station_name)
        aszTemp(i).VehicleTypeCode = FormatDbValue(rsTemp!vehicle_type_code)
        aszTemp(i).VehicleTypeName = FormatDbValue(rsTemp!vehicle_type_short_name)
        aszTemp(i).HalveFullPrice = FormatDbValue(rsTemp!halve_full_price)
        aszTemp(i).HalveHalfPrice = FormatDbValue(rsTemp!halve_half_price)
        aszTemp(i).ServiceFullPrice = FormatDbValue(rsTemp!service_full_price)
        aszTemp(i).ServiceHalfPrice = FormatDbValue(rsTemp!service_half_price)
        aszTemp(i).SpringFullPrice = FormatDbValue(rsTemp!spring_full_price)
        aszTemp(i).SpringHalfPrice = FormatDbValue(rsTemp!spring_half_price)
        rsTemp.MoveNext
    Next i
    GetCompanySettlePriceLst = aszTemp
End Function

'##ModelId=3E81F70D0119
Public Function GetVehicleSettlePriceLst(Optional VehicleID As String = "", Optional CompanyID As String = "", Optional RouteID As String = "", Optional SellstationID As String = "", Optional StationID As String = "") As TVehcileSettlePrice()
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szTemp As String
    Dim sztemp1 As String, i As Integer
    
    Dim aszTemp() As TVehcileSettlePrice
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_GetVehicleSettlePrice
    
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If VehicleID <> "" Then
        szTemp = " AND f.vehicle_id = " & TransFieldValueToString(VehicleID)
    Else
        szTemp = ""
        
    End If
    If RouteID <> "" Then
        szTemp = szTemp + " AND f.route_id = " & TransFieldValueToString(RouteID)
    Else
        szTemp = szTemp + ""
    End If
    If SellstationID <> "" Then
        szTemp = szTemp + " AND f.sell_station_id = " & TransFieldValueToString(SellstationID)
    Else
        szTemp = szTemp + ""
    End If
    If StationID <> "" Then
        szTemp = szTemp + " AND f.station_id = " & TransFieldValueToString(StationID)
    Else
        szTemp = szTemp + ""
    End If
    If CompanyID <> "" Then
        szTemp = szTemp + " AND v.transport_company_id  = " & TransFieldValueToString(CompanyID)
    Else
        szTemp = szTemp + ""
    End If
    szSql = "SELECT f.vehicle_id ,f.route_id,r.route_name,f.sell_station_id,s.station_name sell_station_name,f.station_id,si.station_name , " _
        & " f.mileage,f.pass_charge,f.settle_full_price,f.settle_half_price,f.annotation ,v.license_tag_no ,f.halve_full_price, " _
        & " f.halve_half_price,f.service_full_price,f.service_half_price,f.spring_full_price,f.spring_half_price " _
        & " FROM fin_vehicle_settle_price_lst f ," & cszTableVehicle & " v," & cszTableRouteInfo & " r ," & cszTableStationInfo & " s ," & cszTableStationInfo & " si " _
        & " WHERE f.vehicle_id=v.vehicle_id AND r.route_id=f.route_id AND f.sell_station_id=s.station_id " _
        & " AND f.station_id = si.station_id " & szTemp _
        & " ORDER BY f.sell_station_id , f.route_id , f.vehicle_id , f.station_id "
        
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    ReDim aszTemp(1 To rsTemp.RecordCount)
    For i = 1 To rsTemp.RecordCount
        aszTemp(i).RouteID = FormatDbValue(rsTemp!route_id)
        aszTemp(i).RouteName = FormatDbValue(rsTemp!route_name)
        aszTemp(i).VehcileID = FormatDbValue(rsTemp!vehicle_id)
        aszTemp(i).LicenseTagNo = FormatDbValue(rsTemp!license_tag_no)
        aszTemp(i).Mileage = FormatDbValue(rsTemp!Mileage)
        aszTemp(i).Annotation = FormatDbValue(rsTemp!Annotation)
        aszTemp(i).PassCharge = FormatDbValue(rsTemp!pass_charge)
        aszTemp(i).SellstationID = FormatDbValue(rsTemp!sell_station_id)
        aszTemp(i).SellStationName = FormatDbValue(rsTemp!sell_station_name)
        aszTemp(i).SettlefullPrice = FormatDbValue(rsTemp!settle_full_price)
        aszTemp(i).SettleHalfPrice = FormatDbValue(rsTemp!settle_half_price)
        aszTemp(i).StationID = FormatDbValue(rsTemp!station_id)
        aszTemp(i).StationName = FormatDbValue(rsTemp!station_name)
        aszTemp(i).HalveFullPrice = FormatDbValue(rsTemp!halve_full_price)
        aszTemp(i).HalveHalfPrice = FormatDbValue(rsTemp!halve_half_price)
        aszTemp(i).ServiceFullPrice = FormatDbValue(rsTemp!service_full_price)
        aszTemp(i).ServiceHalfPrice = FormatDbValue(rsTemp!service_half_price)
        aszTemp(i).SpringFullPrice = FormatDbValue(rsTemp!spring_full_price)
        aszTemp(i).SpringHalfPrice = FormatDbValue(rsTemp!spring_half_price)
        rsTemp.MoveNext
    Next i
    GetVehicleSettlePriceLst = aszTemp
End Function

Public Function GetBusSettlePriceLst(Optional BusID As String = "", Optional TransportCompanyID As String = "", Optional SellstationID As String = "", Optional StationID As String = "") As TBusSettlePrice()
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szTemp As String
    Dim sztemp1 As String, i As Integer
    
    Dim aszTemp() As TBusSettlePrice
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_GetVehicleSettlePrice
    
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If BusID <> "" Then
        szTemp = " AND f.bus_id = " & TransFieldValueToString(BusID)
    Else
        szTemp = ""
        
    End If
    If TransportCompanyID <> "" Then
        szTemp = szTemp + " AND f.transport_company_id = " & TransFieldValueToString(TransportCompanyID)
    Else
        szTemp = szTemp + ""
    End If
    If SellstationID <> "" Then
        szTemp = szTemp + " AND f.sell_station_id = " & TransFieldValueToString(SellstationID)
    Else
        szTemp = szTemp + ""
    End If
    If StationID <> "" Then
        szTemp = szTemp + " AND f.station_id = " & TransFieldValueToString(StationID)
    Else
        szTemp = szTemp + ""
    End If
    szSql = "SELECT f.bus_id ,f.transport_company_id,f.sell_station_id,s.station_name sell_station_name,f.station_id,si.station_name , " _
        & " f.mileage,f.pass_charge,f.settle_full_price,f.settle_half_price,f.annotation ,f.halve_full_price,r.route_id,r.route_name, " _
        & " f.halve_half_price,f.service_full_price,f.service_half_price,f.spring_full_price,f.spring_half_price,c.transport_company_short_name " _
        & " FROM fin_bus_settle_price_lst f ,bus_info b," & cszTableCompany & " c ," & cszTableRouteInfo & " r ," & cszTableStationInfo & " s ," & cszTableStationInfo & " si " _
        & " WHERE f.bus_id=b.bus_id AND b.route_id=r.route_id AND f.sell_station_id=s.station_id " _
        & " AND f.station_id = si.station_id AND f.transport_company_id=c.transport_company_id " & szTemp _
        & " ORDER BY f.sell_station_id , f.transport_company_id , f.bus_id , f.station_id "
        
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    ReDim aszTemp(1 To rsTemp.RecordCount)
    For i = 1 To rsTemp.RecordCount
        aszTemp(i).RouteID = FormatDbValue(rsTemp!route_id)
        aszTemp(i).RouteName = FormatDbValue(rsTemp!route_name)
        aszTemp(i).BusID = FormatDbValue(rsTemp!bus_id)
        aszTemp(i).TransportCompanyID = FormatDbValue(rsTemp!transport_company_id)
        aszTemp(i).TransportCompanyName = FormatDbValue(rsTemp!transport_company_short_name)
        aszTemp(i).Mileage = FormatDbValue(rsTemp!Mileage)
        aszTemp(i).Annotation = FormatDbValue(rsTemp!Annotation)
        aszTemp(i).PassCharge = FormatDbValue(rsTemp!pass_charge)
        aszTemp(i).SellstationID = FormatDbValue(rsTemp!sell_station_id)
        aszTemp(i).SellStationName = FormatDbValue(rsTemp!sell_station_name)
        aszTemp(i).SettlefullPrice = FormatDbValue(rsTemp!settle_full_price)
        aszTemp(i).SettleHalfPrice = FormatDbValue(rsTemp!settle_half_price)
        aszTemp(i).StationID = FormatDbValue(rsTemp!station_id)
        aszTemp(i).StationName = FormatDbValue(rsTemp!station_name)
        aszTemp(i).HalveFullPrice = FormatDbValue(rsTemp!halve_full_price)
        aszTemp(i).HalveHalfPrice = FormatDbValue(rsTemp!halve_half_price)
        aszTemp(i).ServiceFullPrice = FormatDbValue(rsTemp!service_full_price)
        aszTemp(i).ServiceHalfPrice = FormatDbValue(rsTemp!service_half_price)
        aszTemp(i).SpringFullPrice = FormatDbValue(rsTemp!spring_full_price)
        aszTemp(i).SpringHalfPrice = FormatDbValue(rsTemp!spring_half_price)
        rsTemp.MoveNext
    Next i
    GetBusSettlePriceLst = aszTemp
End Function

'##ModelId=3E81F716006D
Public Function GetSettleSheetInfo(Optional Status As ESettleSheetStatus = -1, Optional StartDate As Date = cszEmptyDateStr, _
                            Optional EndDate As Date = cszForeverDateStr, Optional SettleObject As ESettleObjectType = -1, Optional ObjectID As String = "", Optional SettleSheetID As String = "", Optional ObjectName As String = "", Optional CheckSheetID As String = "") As TSettleSheet()
    'CheckSheetID为按照路单号来查结算单号
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim aTTemp() As TSettleSheet
    Dim szTemp As String, i As Integer
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_SettleSheet
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szTemp = ""
    
    If Status <> -1 Then
        szTemp = szTemp + " AND s.status = " & TransFieldValueToString(Status)
    End If
    If StartDate <> cszEmptyDateStr Then
        szTemp = szTemp + " AND s.start_date >=" & TransFieldValueToString(StartDate)
    End If
    If EndDate <> cszForeverDateStr Then
        szTemp = szTemp + " AND s.end_date <= " & TransFieldValueToString(EndDate)
    End If
    
    If ObjectID <> "" Then
        szTemp = szTemp + " AND s.object_id = " & TransFieldValueToString(ObjectID)
    End If
    If SettleObject <> -1 Then
        szTemp = szTemp + " AND s.settle_object = " & TransFieldValueToString(SettleObject)
    End If
    If SettleSheetID <> "" Then
        szTemp = szTemp + " AND s.settlement_sheet_id = " & TransFieldValueToString(SettleSheetID)
    End If
    If ObjectName <> "" Then
        szTemp = szTemp & " AND s.object_name LIKE " & TransFieldValueToString("%" & ObjectName & "%")
    End If
    If CheckSheetID <> "" Then
        szTemp = szTemp & " AND settlement_sheet_id IN ( SELECT settlement_sheet_id FROM fin_settle_bus_station_lst WHERE check_sheet_id =" & TransFieldValueToString(CheckSheetID) & " ) "
    End If
    szSql = "SELECT * FROM fin_settle_sheet_lst s WHERE 1=1 " & szTemp
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    ReDim aTTemp(1 To rsTemp.RecordCount)
    For i = 1 To rsTemp.RecordCount
        aTTemp(i).Annotation = FormatDbValue(rsTemp!Annotation)
        aTTemp(i).Checker = FormatDbValue(rsTemp!Checker)
        aTTemp(i).CheckSheetCount = FormatDbValue(rsTemp!sheet_quantity)
        aTTemp(i).ObjectID = FormatDbValue(rsTemp!object_id)
        aTTemp(i).SettleDate = FormatDbValue(rsTemp!settle_date)
        aTTemp(i).SettleEndDate = FormatDbValue(rsTemp!end_date)
        aTTemp(i).SettleObject = FormatDbValue(rsTemp!settle_object)
        aTTemp(i).SettleOtherCompanyPrice = FormatDbValue(rsTemp!settle_price)
        aTTemp(i).Settler = FormatDbValue(rsTemp!Settler)
        aTTemp(i).SettleSheetID = FormatDbValue(rsTemp!settlement_sheet_id)
        aTTemp(i).SettleStartDate = FormatDbValue(rsTemp!start_date)
        aTTemp(i).SettleStationPrice = FormatDbValue(rsTemp!settle_station_price)
        aTTemp(i).Status = FormatDbValue(rsTemp!Status)
        aTTemp(i).TotalQuantity = FormatDbValue(rsTemp!total_quantity)
        aTTemp(i).TotalTicketPrice = FormatDbValue(rsTemp!total_ticket_price)
        aTTemp(i).TransportCompanyID = FormatDbValue(rsTemp!transport_company_id)
        aTTemp(i).TransportCompanyName = FormatDbValue(rsTemp!transport_company_name)
        aTTemp(i).UnitName = FormatDbValue(rsTemp!Annotation)
        
        aTTemp(i).ObjectName = FormatDbValue(rsTemp!object_name)
        aTTemp(i).LuggageSettleIDs = FormatDbValue(rsTemp!luggage_settle_id)
        aTTemp(i).LuggageTotalBaseCarriage = FormatDbValue(rsTemp!luggage_base_carriage)
        aTTemp(i).LuggageTotalSettlePrice = FormatDbValue(rsTemp!luggage_settle_price)
        aTTemp(i).LuggageProtocolName = FormatDbValue(rsTemp!luggage_protocol_name)
        
        aTTemp(i).RouteID = FormatDbValue(rsTemp!route_id)
        aTTemp(i).RouteName = FormatDbValue(rsTemp!route_name)
        
        rsTemp.MoveNext
    Next i
    GetSettleSheetInfo = aTTemp
End Function

'##ModelId=3E81F721005D
Public Function GetVehicleProtocol(Optional pszVehicleID As String = "", Optional pszCompanyID As String = "", Optional pszProtocolID As String = "", Optional pszLicenseTagNo As String = "") As String()
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim aszTemp() As String
    Dim szTemp As String, i As Integer
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_GetVehicleProtocol
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If pszVehicleID <> "" Then
        szTemp = " AND v.vehicle_id = " & TransFieldValueToString(pszVehicleID)
    End If
    If pszProtocolID <> "" Then
        szTemp = szTemp & " AND v.protocol_id = " & TransFieldValueToString(pszProtocolID)
    End If
    If pszCompanyID <> "" Then
        szTemp = szTemp & " AND v.transport_company_id = " & TransFieldValueToString(pszCompanyID)
    End If
    If pszLicenseTagNo <> "" Then
        szTemp = szTemp & " AND v.license_tag_no LIKE " & TransFieldValueToString("%" & pszLicenseTagNo & "%")
    End If
    szSql = "SELECT ISNULL( v.protocol_id,'') protocol_id ,ISNULL(p.protocol_name,'') protocol_name , " _
        & " v.transport_company_id company_id,c.transport_company_short_name company_name,ISNULL( p.default_mark,0) default_mark,v.vehicle_id,v.license_tag_no" _
        & " FROM " & cszTableVehicle & " v LEFT JOIN fin_settle_protocol_info p ON v.protocol_id=p.protocol_id " _
        & " JOIN " & cszTableCompany & " c ON c.transport_company_id=v.transport_company_id WHERE 1=1 " & szTemp
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 7)
    For i = 1 To rsTemp.RecordCount
        aszTemp(i, 1) = FormatDbValue(rsTemp!protocol_id)
        aszTemp(i, 2) = FormatDbValue(rsTemp!protocol_name)
        aszTemp(i, 3) = FormatDbValue(rsTemp!vehicle_id)
        aszTemp(i, 4) = FormatDbValue(rsTemp!license_tag_no)
        aszTemp(i, 5) = FormatDbValue(rsTemp!company_id)
        aszTemp(i, 6) = FormatDbValue(rsTemp!company_name)
        aszTemp(i, 7) = FormatDbValue(rsTemp!default_mark)
        rsTemp.MoveNext
    Next i
    GetVehicleProtocol = aszTemp
End Function


'##ModelId=3E81F72403D8
Public Function GetCompanyProtocol(Optional pszCompanyID As String = "", Optional pszProtocolID As String = "") As String()
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim aszTemp() As String
    Dim szTemp As String, i As Integer
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_GetCompanyProtocol
    
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If Trim(pszCompanyID) <> "" Then
        szTemp = " AND c.transport_company_id = " & TransFieldValueToString(pszCompanyID)
    End If
    If Trim(pszProtocolID) <> "" Then
        szTemp = szTemp + " AND c.protocol_id = " & TransFieldValueToString(pszProtocolID)
    End If
    szSql = "SELECT ISNULL( c.protocol_id,'') protocol_id ,ISNULL( p.protocol_name,'')protocol_name  ,ISNULL( c.transport_company_id ,'') transport_company_id  ,ISNULL( p.default_mark,'') default_mark ,ISNULL( c.transport_company_short_name,'') transport_company_name  " _
        & " FROM " & cszTableCompany & " c LEFT JOIN fin_settle_protocol_info p ON c.protocol_id=p.protocol_id  WHERE 1=1 " & szTemp
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 5)
    For i = 1 To rsTemp.RecordCount
        aszTemp(i, 1) = FormatDbValue(rsTemp!protocol_id)
        aszTemp(i, 2) = FormatDbValue(rsTemp!protocol_name)
        aszTemp(i, 3) = FormatDbValue(rsTemp!transport_company_id)
        aszTemp(i, 4) = FormatDbValue(rsTemp!transport_company_name)
        aszTemp(i, 5) = FormatDbValue(rsTemp!default_mark)
        rsTemp.MoveNext
    Next i
    GetCompanyProtocol = aszTemp
End Function


'##ModelId=3E81F72D0232
Public Function GetNeedSplitCheckSheet(peSettleObject As ESettleObjectType, pszObjectID As String, pdtStartDate As Date, pdtEndDate As Date, pszRouteID As String, paszCheckSheetID() As String, Optional pszCompanyID As String, Optional pbIsToday As Boolean = False) As ADODB.Recordset
    'pbIsToday 表示是否是当天结算,true 为当天结算,则不从统计表(fin_sheet_station_lst)中取数据,而从sheet_station_lst表中取
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim szWhere As String
    Dim nSheetCount As Integer
    
    
    Dim oSysParm As New SystemParam
    Dim bAllowSplitBySomeTimes As Boolean
    oSysParm.Init m_oActiveUser
    
    bAllowSplitBySomeTimes = oSysParm.AllowSplitBySomeTimes
    AssertObjIsValid
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    '得到路单
    nSheetCount = ArrayLength(paszCheckSheetID)
    
    
    szSql = "SELECT s.check_sheet_id,MAX(s.bus_date) bus_date,MAX(s.bus_id) bus_id,MAX(s.bus_serial_no) bus_serial_no ,MAX(s.check_gate_id) check_gate_id, " _
        & " MAX(c.transport_company_short_name) transport_company_short_name,MAX(r.route_name) route_name,MAX(v.license_tag_no) license_tag_no ," _
        & " MAX(t.station_name) station_name,MAX(s.transport_company_id) transport_company_id,MAX(s.vehicle_id) vehicle_id,MAX(r.mileage) mileage"
        
        
    '注意,此处的路单条件:  结算状态为不等于已结算,表示,未结算及正在结算的路单均算, 且人数为总人数减去已结人数
    
    If (Not bAllowSplitBySomeTimes) Then
        '如果只能一次性结算
        szSql = szSql & " ,SUM(l.quantity) quantity,SUM(l.ticket_price*l.quantity) ticket_price "
        szWhere = szWhere & "  AND s.settlement_status= " & TransFieldValueToString(CS_CheckSheetNotSettled)
            
    Else
        '如果可以分多次结算
        szSql = szSql & " ,SUM(l.quantity-l.fact_quantity)  quantity,SUM(l.ticket_price*(l.quantity-l.fact_quantity)) ticket_price "
        szWhere = szWhere & " and s.settlement_status<>" & TransFieldValueToString(CS_CheckSheetSettled)
        
    End If
    
    
    If nSheetCount > 0 Then
        '如果存在查询某几张路单的条件,则
        '此条件在添加路单时用到
        
        
        szWhere = szWhere & " AND s.check_sheet_id IN ( "
        For i = 1 To nSheetCount - 1
                szWhere = szWhere & "'" & paszCheckSheetID(i) & "',"
        Next i
        szWhere = szWhere & " '" & paszCheckSheetID(i) & "')"
        
        
    Else
        
        szWhere = szWhere & " AND s.bus_date >= " & TransFieldValueToString(ToDBDate(pdtStartDate)) _
            & " AND s.bus_date <= " & TransFieldValueToString(ToDBDate(pdtEndDate))
        '对象类型为参运公司
        If pszRouteID <> "" Then
            szWhere = szWhere & " AND s.route_id = " & TransFieldValueToString(pszRouteID)
        End If
        
        If pszCompanyID <> "" Then
            szWhere = szWhere & " AND s.transport_company_id = " & TransFieldValueToString(pszCompanyID)
        End If
        
        '按公司结算与按车辆结算条件不一样,其他都相同
        If peSettleObject = CS_SettleByTransportCompany Then
            '按公司结算
            szWhere = szWhere & " AND s.transport_company_id= " & TransFieldValueToString(pszObjectID)

        ElseIf peSettleObject = CS_SettleByVehicle Then
            '按车辆结算
            szWhere = szWhere & " AND s.vehicle_id= " & TransFieldValueToString(pszObjectID)
            
        ElseIf peSettleObject = CS_SettleByBus Then
            '按车次结算
            szWhere = szWhere & " AND s.bus_id= " & TransFieldValueToString(pszObjectID)
        End If
    End If
        
    szSql = szSql & " FROM  " & cszTableCheckSheet & " s," & cszTableCompany & "  c," & cszTableRouteInfo & " r," & cszTableVehicle & " v," & IIf(Not pbIsToday, cszTableSheetStation, cszTableSheetStationTemp) & " l ,station_info t " _
        & " WHERE s.transport_company_id=c.transport_company_id AND s.vehicle_id=v.vehicle_id AND l.check_sheet_id=s.check_sheet_id AND r.end_station_id=t.station_id  " _
        & " AND s.route_id=r.route_id AND UPPER(l.sell_station_id) NOT IN ('CM' ,'QG') " _
        & " AND s.valid_mark= " & TransFieldValueToString(CS_CheckSheetValid) _
        & szWhere _
        & " GROUP BY s.check_sheet_id"
    
    
    Set rsTemp = oDB.Execute(szSql)
    Set GetNeedSplitCheckSheet = rsTemp
    
    Set oSysParm = Nothing
    Set rsTemp = Nothing
    
End Function

'**************************************************
'Member Code:S3
'Brief Description:测试对象是否有效
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then RaiseError ERR_ReportNotAvailable
End Sub

'**************************************************
'Member Code:S4
'Brief Description:测试对象是否有效
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub AssertObjIsValid()
    AssertActiveUserValid m_oActiveUser, ERR_Report
    AssertStatusAvailable
End Sub

Public Function ReprintSettleSheet(OldSettlSheetID As String, NewSettlSheetID As String) As ADODB.Recordset  '重打结算单
On Error GoTo Here
'    Dim oDB As New RTConnection
'    Dim rsTemp As Recordset
'    Dim rsTemp1 As Recordset
'    Dim rsTemp2 As Recordset
'    Dim szSql As String
'    AssertObjIsValid
'    AssertHaveRight m_oActiveUser, RIGHT_ReprintSettleSheet
'
'    oDB.ConnectionString = GetConnectionStr(cszSettle)
'
'    oDB.BeginTrans
'    szSql = "SELECT * FROM fin_settle_sheet_lst WHERE settlement_sheet_id=" & TransFieldValueToString(OldSettlSheetID)
'    Set rsTemp = oDB.Execute(szSql)
'    If FormatDbValue(rsTemp!settle_object) = CS_SettleByVehicle Then
'        '新增结算单
'        szSql = "INSERT INTO fin_settle_sheet_lst ( settlement_sheet_id ,sheet_quantity," _
'            & " total_ticket_price,total_quantity,settle_price,settle_station_price,object_id , object_name , settle_object,transport_company_id, " _
'            & " transport_company_name,settler,checker,settle_date,start_date,end_date,unit_name,status,luggage_settle_id,luggage_base_carriage, " _
'            & " luggage_settle_price,luggage_protocol_name,annotation ) " _
'            & " SELECT " & TransFieldValueToString(NewSettlSheetID) & ",sheet_quantity," _
'            & " total_ticket_price,total_quantity,settle_price,settle_station_price,object_id , object_name , settle_object,transport_company_id, " _
'            & " transport_company_name,settler,checker,settle_date,start_date,end_date,unit_name,status,luggage_settle_id,luggage_base_carriage, " _
'            & " luggage_settle_price,luggage_protocol_name,annotation FROM fin_settle_sheet_lst " _
'            & " WHERE settlement_sheet_id=" & TransFieldValueToString(OldSettlSheetID)
'        oDB.Execute szSql
'
'        szSql = "INSERT INTO fin_settle_vehicle_lst (settlement_sheet_id,vehicle_id," _
'            & " protocol_id,settle_price,settle_station_price,passenger_number,split_item_1,split_item_2,split_item_3,split_item_4," _
'            & " split_item_5,split_item_6,split_item_7,split_item_8,split_item_9,split_item_10,split_item_11,split_item_12,split_item_13,split_item_14," _
'            & " split_item_15,split_item_16,split_item_17,split_item_18,split_item_19,split_item_20)" _
'            & " SELECT " & TransFieldValueToString(NewSettlSheetID) & " ,vehicle_id," _
'            & " protocol_id,settle_price,settle_station_price,passenger_number,split_item_1,split_item_2,split_item_3,split_item_4," _
'            & " split_item_5,split_item_6,split_item_7,split_item_8,split_item_9,split_item_10,split_item_11,split_item_12,split_item_13,split_item_14," _
'            & " split_item_15,split_item_16,split_item_17,split_item_18,split_item_19,split_item_20 FROM fin_settle_vehicle_lst " _
'            & " WHERE settlement_sheet_id=" & TransFieldValueToString(OldSettlSheetID)
'        oDB.Execute szSql
'
'        szSql = "INSERT INTO fin_settle_station_lst (settlement_sheet_id,sell_station_id,  " _
'            & " route_id,station_id,ticket_type_id,quantity,bus_id,transport_company_id,vehicle_type_code,vehicle_id) " _
'            & " SELECT " & TransFieldValueToString(NewSettlSheetID) & ",sell_station_id,  " _
'            & " route_id,station_id,ticket_type_id,quantity,bus_id,transport_company_id,vehicle_type_code,vehicle_id  FROM fin_settle_station_lst " _
'            & " WHERE settlement_sheet_id=" & TransFieldValueToString(OldSettlSheetID)
'        oDB.Execute szSql
'
'    ElseIf FormatDbValue(rsTemp!settle_object) = CS_SettleByTransportCompany Then
'        szSql = "INSERT INTO fin_settle_sheet_lst(settlement_sheet_id,sheet_quantity," _
'            & " total_ticket_price,total_quantity,settle_price,settle_station_price,object_id,object_name,settle_object,transport_company_id, " _
'            & " transport_company_name,settler,checker,settle_date,start_date,end_date,unit_name,status,luggage_settle_id,luggage_base_carriage, " _
'            & " luggage_settle_price,luggage_protocol_name,annotation) " _
'            & " SELECT " & TransFieldValueToString(NewSettlSheetID) & ",sheet_quantity," _
'            & " total_ticket_price,total_quantity,settle_price,settle_station_price,object_id,object_name,settle_object,transport_company_id, " _
'            & " transport_company_name,settler,checker,settle_date,start_date,end_date,unit_name,status,luggage_settle_id,luggage_base_carriage, " _
'            & " luggage_settle_price,luggage_protocol_name,annotation FROM fin_settle_sheet_lst " _
'            & " WHERE settlement_sheet_id=" & TransFieldValueToString(OldSettlSheetID)
'        oDB.Execute szSql
'
'        szSql = "INSERT INTO fin_settle_company_price_lst(settlement_sheet_id,transport_company_id," _
'            & " protocol_id,settle_price,settle_station_price,passenger_number,split_item_1,split_item_2,split_item_3,split_item_4," _
'            & " split_item_5,split_item_6,split_item_7,split_item_8,split_item_9,split_item_10,split_item_11,split_item_12,split_item_13,split_item_14," _
'            & " split_item_15,split_item_16,split_item_17,split_item_18,split_item_19,split_item_20) " _
'            & " SELECT " & TransFieldValueToString(NewSettlSheetID) & " ,transport_company_id," _
'            & " protocol_id,settle_price,settle_station_price,passenger_number,split_item_1,split_item_2,split_item_3,split_item_4," _
'            & " split_item_5,split_item_6,split_item_7,split_item_8,split_item_9,split_item_10,split_item_11,split_item_12,split_item_13,split_item_14," _
'            & " split_item_15,split_item_16,split_item_17,split_item_18,split_item_19,split_item_20 FROM fin_settle_company_price_lst " _
'            & " WHERE settlement_sheet_id=" & TransFieldValueToString(OldSettlSheetID)
'        oDB.Execute szSql
'
'        szSql = "SELECT * FROM fin_settle_station_lst WHERE settlement_sheet_id=" & TransFieldValueToString(OldSettlSheetID)
'        Set rsTemp2 = oDB.Execute(szSql)
'        If rsTemp2.RecordCount > 0 Then
'            szSql = "INSERT INTO fin_settle_station_lst (settlement_sheet_id,sell_station_id, " _
'                & " route_id,station_id,ticket_type_id,quantity,bus_id,transport_company_id,vehicle_type_code,vehicle_id) " _
'                & " SELECT " & TransFieldValueToString(NewSettlSheetID) & ",sell_station_id, " _
'                & " route_id,station_id,ticket_type_id,quantity,bus_id,transport_company_id,vehicle_type_code,vehicle_id FROM fin_settle_station_lst " _
'                & " WHERE settlement_sheet_id=" & TransFieldValueToString(OldSettlSheetID)
'            oDB.Execute szSql
'        End If
'
'    End If
'        '作废原结算单
'        szSql = "UPDATE fin_settle_sheet_lst SET status=" & TransFieldValueToString(CS_SettleSheetInvalid) _
'            & " WHERE settlement_sheet_id=" & TransFieldValueToString(OldSettlSheetID)
'        oDB.Execute szSql
'
'        '改变对应原路单的结算单
'
'        szSql = "UPDATE " & cszTableCheckSheet & " " _
'            & " SET settlement_sheet_id =" & TransFieldValueToString(NewSettlSheetID) _
'            & " WHERE settlement_sheet_id = " & TransFieldValueToString(OldSettlSheetID)
'        oDB.Execute szSql
'
'
'    oDB.CommitTrans
'
'    WriteOperateLog m_oActiveUser, RIGHT_ReprintSettleSheet, "作废结算单" & OldSettlSheetID & "生成新的结算单号为" & NewSettlSheetID
'
'    Set rsTemp = Nothing
'    Set rsTemp1 = Nothing
'    Exit Function
Here:
'    oDB.RollbackTrans
'    err.Raise err.Number, err.Source, err.Description
'    Set rsTemp = Nothing
'    Set rsTemp1 = Nothing
End Function

'打印结算单
Public Function PrintSettleSheet(pszSheetID As String) As Recordset
On Error GoTo Here
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
'    Dim rsTemp1 As Recordset
    Dim szSql As String
    
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_PrintSettleSheet
    
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "SELECT * FROM fin_settle_sheet_lst WHERE settlement_sheet_id=" & TransFieldValueToString(pszSheetID)
    Set rsTemp = oDB.Execute(szSql)
            If rsTemp.RecordCount = 0 Then RaiseError ERR_SettleSheetNotExist
        If FormatDbValue(rsTemp!settle_object) = CS_SettleByVehicle Then
            szSql = "SELECT c.vehicle_id object_id ,MAX(f.object_name) object_name , MAX(f.settlement_sheet_id) settlement_sheet_id ,MAX(f.sheet_quantity) sheet_quantity ,MAX(f.total_ticket_price) total_ticket_price , " _
                & " MAX(f.luggage_settle_id) luggage_settle_id , MAX(f.luggage_base_carriage) luggage_base_carriage , MAX(f.luggage_settle_price) luggage_settle_price, " _
                & " MAX(f.luggage_protocol_name) luggage_protocol_name, " _
                & " MAX(f.total_quantity) total_quantity,MAX(f.settle_price-f.settle_station_price) settle_price,MAX(f.settler) settler,MAX(f.settle_date) settle_date, " _
                & " MAX(f.start_date) start_date,MAX(f.end_date) end_date,MAX(f.unit_name) unit_name, " _
                & " MAX(c.split_item_1) split_item_1 ,MAX(c.split_item_2) split_item_2,MAX(c.split_item_3) split_item_3, " _
                & " MAX(c.split_item_4) split_item_4,MAX(c.split_item_5) split_item_5,MAX(c.split_item_6) split_item_6,MAX(c.split_item_7) split_item_7, " _
                & " MAX(c.split_item_8) split_item_8,MAX(c.split_item_9) split_item_9,MAX(c.split_item_10) split_item_10 , " _
                & " MAX(c.split_item_11) split_item_11 ,MAX(c.split_item_12) split_item_12,MAX(c.split_item_13) split_item_13, " _
                & " MAX(c.split_item_14) split_item_14,MAX(c.split_item_15) split_item_15,MAX(c.split_item_16) split_item_16,MAX(c.split_item_17) split_item_17, " _
                & " MAX(c.split_item_18) split_item_18,MAX(c.split_item_19) split_item_19,MAX(c.split_item_20) split_item_20 FROM  fin_settle_sheet_lst f, " _
                & " fin_settle_vehicle_lst c , transport_vehicle_info v WHERE f.settlement_sheet_id=c.settlement_sheet_id AND c.vehicle_id = v.vehicle_id AND f.settlement_sheet_id=" & TransFieldValueToString(pszSheetID) _
                & " GROUP BY c.vehicle_id "
        ElseIf FormatDbValue(rsTemp!settle_object) = CS_SettleByTransportCompany Then
            szSql = "SELECT c.transport_company_id object_id , MAX(f.object_name) object_name ,MAX(f.settlement_sheet_id) settlement_sheet_id ,MAX(f.sheet_quantity) sheet_quantity ,MAX(f.total_ticket_price) total_ticket_price , " _
                & " MAX(f.luggage_settle_id) luggage_settle_id , MAX(f.luggage_base_carriage) luggage_base_carriage , MAX(f.luggage_settle_price) luggage_settle_price,  " _
                & " MAX(f.luggage_protocol_name) luggage_protocol_name, " _
                & " MAX(f.total_quantity) total_quantity,MAX(f.settle_price-f.settle_station_price) settle_price,MAX(f.settler) settler,MAX(f.settle_date) settle_date, " _
                & " MAX(f.start_date) start_date,MAX(f.end_date) end_date,MAX(f.unit_name) unit_name, " _
                & " MAX(c.split_item_1) split_item_1 ,MAX(c.split_item_2) split_item_2,MAX(c.split_item_3) split_item_3, " _
                & " MAX(c.split_item_4) split_item_4,MAX(c.split_item_5) split_item_5,MAX(c.split_item_6) split_item_6,MAX(c.split_item_7) split_item_7, " _
                & " MAX(c.split_item_8) split_item_8,MAX(c.split_item_9) split_item_9,MAX(c.split_item_10) split_item_10 , " _
                & " MAX(c.split_item_11) split_item_11 ,MAX(c.split_item_12) split_item_12,MAX(c.split_item_13) split_item_13, " _
                & " MAX(c.split_item_14) split_item_14,MAX(c.split_item_15) split_item_15,MAX(c.split_item_16) split_item_16,MAX(c.split_item_17) split_item_17, " _
                & " MAX(c.split_item_18) split_item_18,MAX(c.split_item_19) split_item_19,MAX(c.split_item_20) split_item_20  FROM  fin_settle_sheet_lst f, " _
                & " fin_settle_company_price_lst c , transport_company_info p WHERE f.settlement_sheet_id=c.settlement_sheet_id AND c.transport_company_id = p.transport_company_id AND f.settlement_sheet_id=" & TransFieldValueToString(pszSheetID) _
                & " GROUP BY c.transport_company_id "
        End If
        
        Set rsTemp = oDB.Execute(szSql)
        Set PrintSettleSheet = rsTemp
        
        
        Set rsTemp = Nothing
    Exit Function
Here:
    err.Raise err.Number, err.Source, err.Description
    Set rsTemp = Nothing
End Function

'随路单结算打印行包结算单
Public Function GetLugFinSheet(pszSplitType As Integer, pszObjectID As String, pdtStartDate As Date, pdtEndDate As Date) As ADODB.Recordset
    On Error GoTo err
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "SELECT MAX(i.split_object_name) split_object_name,SUM(i.total_price) total_price,SUM(i.need_split_out) need_split_out,MAX(l.protocol_name) protocol_name " _
        & " FROM  fin_luggage_sheet_info i,fin_luggage_sheet_lst l " _
        & " WHERE i.fin_sheet_id=l.fin_sheet_id AND i.split_object_type=" & TransFieldValueToString(pszSplitType) _
        & " AND i.split_object_id=" & TransFieldValueToString(pszObjectID) _
        & " AND i.settlement_start_time=" & TransFieldValueToString(ToDBDate(pdtStartDate)) _
        & " AND i.settlement_end_time=" & TransFieldValueToString(ToDBDate(pdtEndDate)) _
        & " AND i.status=1 "
    Set rsTemp = oDB.Execute(szSql)
    Set GetLugFinSheet = rsTemp
    
    Exit Function
err:
    err.Raise err.Number, err.Source, err.Description
End Function


Public Function PreLugFinSheet(pszSheetID() As String) As ADODB.Recordset
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim szTemp As String
    Dim i As Integer
    szTemp = " ("
    For i = 1 To ArrayLength(pszSheetID) - 1
        szTemp = szTemp & pszSheetID(i) & ","
    Next i
    szTemp = szTemp & pszSheetID(i) & ")"
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "SELECT MAX(i.split_object_type) split_object_type,MAX(i.split_object_name) split_object_name,SUM(l.total_price) total_price,SUM(l.need_split_out) need_split_out,MAX(l.protocol_name) protocol_name " _
        & " FROM  fin_luggage_sheet_info i,fin_luggage_sheet_lst l " _
        & " WHERE i.fin_sheet_id=l.fin_sheet_id AND i.fin_sheet_id IN " & szTemp _
        & " GROUP BY i.split_object_name"
        
    Set rsTemp = oDB.Execute(szSql)
    Set PreLugFinSheet = rsTemp
End Function

'判断行包结算单的有效性
Public Function GetLugSheet(pszSheetID) As ADODB.Recordset
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "SELECT split_object_name FROM fin_luggage_sheet_info WHERE " _
        & " fin_sheet_id=" & TransFieldValueToString(pszSheetID) _
        & " AND status=1 "
        
    Set rsTemp = oDB.Execute(szSql)
    Set GetLugSheet = rsTemp
End Function

Public Function GetSettleRouteQuantity(pszSheetID As String) As ADODB.Recordset
On Error GoTo Here
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim rsVehicleType As Recordset
    Dim szSql As String
    
    AssertObjIsValid
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "SELECT * FROM fin_settle_sheet_lst WHERE settlement_sheet_id=" & TransFieldValueToString(pszSheetID)
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then RaiseError ERR_SettleSheetNotExist
'    If FormatDbValue(rsTemp!settle_object) = CS_SettleByTransportCompany Then
'
'
'        szSql = " SELECT f.*,p.mileage,p.mileage*quantity man_mul_mileage,p.settle_full_price, p.settle_half_price " _
'            & " , p.halve_full_price , p.halve_half_price , p.service_full_price ,p.service_half_price,p.pass_charge " _
'            & " , p.spring_full_price , p.spring_half_price " _
'            & " , f.route_name,s.station_name, t.ticket_type_name , h.vehicle_type_short_name vehicle_type_name , a.station_name sell_station_name " _
'            & " , ltrim(convert(char(15), p.settle_full_price)) + '/' +ltrim(convert(char(15), p.settle_half_price)) settle_price " _
'            & " , ltrim(convert(char(15), p.halve_full_price)) + '/' +ltrim(convert(char(15), p.halve_half_price)) halve_price " _
'            & " , ltrim(convert(char(15), p.service_full_price)) + '/' +ltrim(convert(char(15), p.service_half_price)) service_price " _
'            & " , ltrim(convert(char(15), p.spring_full_price)) + '/' +ltrim(convert(char(15), p.spring_half_price)) spring_price " _
'            & " FROM fin_settle_station_lst f,fin_company_settle_price_lst p,station_info s, " & cszTableTicketType & " t " _
'            & "     , " & cszTableVehicleTypeCode & " h  , station_info a " _
'            & " WHERE f.station_id=p.station_id AND f.route_id=p.route_id AND f.sell_station_id=p.sell_station_id " _
'            & " AND settlement_sheet_id=" & TransFieldValueToString(pszSheetID) _
'            & " AND s.station_id=f.station_id AND f.ticket_type_id = t.ticket_type_id " _
'            & " AND p.transport_company_id=" & TransFieldValueToString(FormatDbValue(rsTemp!object_id)) _
'            & " AND p.vehicle_type_code = f.vehicle_type_code AND p.sell_station_id = a.station_id AND p.vehicle_type_code = h.vehicle_type_code  "
'            'IN ( SELECT p.vehicle_type_code  FROM fin_settle_station_lst f,fin_company_settle_price_lst p  WHERE f.station_id=p.station_id AND f.route_id=p.route_id AND f.sell_station_id=p.sell_station_id  AND settlement_sheet_id=" & TransFieldValueToString(pszSheetID) & " AND p.transport_company_id=" & TransFieldValueToString(FormatDbValue(rsTemp!object_id)) & ")"
'        Set rsTemp = oDB.Execute(szSql)

        
        szSql = "  SELECT f.*,s.station_name, t.ticket_type_name  " _
            & " FROM fin_settle_station_lst f,station_info s,  ticket_type_code  t   " _
            & " WHERE s.station_id=f.station_id AND f.ticket_type_id = t.ticket_type_id   " _
            & " AND settlement_sheet_id=" & TransFieldValueToString(pszSheetID)

        Set rsTemp = oDB.Execute(szSql)



'    Else
'        szSql = " SELECT f.*,p.mileage,p.mileage*quantity man_mul_mileage,p.settle_full_price, p.settle_half_price " _
'            & " , p.halve_full_price , p.halve_half_price , p.service_full_price ,p.service_half_price,p.pass_charge " _
'            & " , p.spring_full_price , p.spring_half_price " _
'            & " , f.route_name,s.station_name, t.ticket_type_name  , h.vehicle_type_short_name vehicle_type_name  " _
'            & " , ltrim(convert(char(15), p.settle_full_price)) + '/' +ltrim(convert(char(15), p.settle_half_price)) settle_price " _
'            & " , ltrim(convert(char(15), p.halve_full_price)) + '/' +ltrim(convert(char(15), p.halve_half_price)) halve_price " _
'            & " , ltrim(convert(char(15), p.service_full_price)) + '/' +ltrim(convert(char(15), p.service_half_price)) service_price " _
'            & " , ltrim(convert(char(15), p.spring_full_price)) + '/' +ltrim(convert(char(15), p.spring_half_price)) spring_price " _
'            & " FROM fin_settle_station_lst f,fin_vehicle_settle_price_lst p,station_info s, ticket_type_code t, " & cszTableVehicleTypeCode & " h " _
'            & " WHERE f.station_id=p.station_id AND f.route_id=p.route_id AND f.sell_station_id=p.sell_station_id " _
'            & " AND settlement_sheet_id=" & TransFieldValueToString(pszSheetID) _
'            & " AND s.station_id=f.station_id AND f.ticket_type_id = t.ticket_type_id  " _
'            & " AND f.vehicle_type_code = h.vehicle_type_code AND p.vehicle_id=" & TransFieldValueToString(FormatDbValue(rsTemp!object_id))
'          Set rsTemp = oDB.Execute(szSql)
'    End If
    Set GetSettleRouteQuantity = rsTemp
    
    Exit Function
Here:
    err.Raise err.Number, err.Source, err.Description
    
End Function

Public Function GetSettleRouteCalRemark(pszSheetID As String) As ADODB.Recordset
On Error GoTo Here
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim rsVehicleType As Recordset
    Dim szSql As String
    
    AssertObjIsValid
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "SELECT * FROM fin_settle_sheet_lst WHERE settlement_sheet_id=" & TransFieldValueToString(pszSheetID)
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then RaiseError ERR_SettleSheetNotExist
    If FormatDbValue(rsTemp!settle_object) = CS_SettleByTransportCompany Then
        szSql = " SELECT p.vehicle_type_code, MAX(v.vehicle_type_name) vehicle_type_name  , p.transport_company_id , c.route_id , max(r.route_name) route_name , max(p.annotation ) annotation " _
            & " FROM " & cszTableCheckSheet & " c , fin_company_settle_price_lst p , route_info r , vehicle_type_code v " _
            & " WHERE c.route_id = r.route_id AND c.route_id = p.route_id  AND p.annotation<>'' " _
            & " AND c.settlement_sheet_id =" & TransFieldValueToString(pszSheetID) _
            & " AND p.vehicle_type_code = v.vehicle_type_code " _
            & " AND p.transport_company_id = " & TransFieldValueToString(rsTemp!object_id) _
            & " GROUP BY p.vehicle_type_code, p.transport_company_id , c.route_id "
            
        '****此处应该需要加上车型
    Else
        szSql = " SELECT p.vehicle_id ,c.route_id , max(r.route_name) route_name , max(p.annotation ) annotation " _
            & " FROM " & cszTableCheckSheet & " c , fin_vehicle_settle_price_lst p , route_info r " _
            & " WHERE c.route_id = r.route_id AND c.route_id = p.route_id AND p.annotation<>'' " _
            & " AND c.settlement_sheet_id =" & TransFieldValueToString(pszSheetID) _
            & " AND p.vehicle_id = " & TransFieldValueToString(rsTemp!object_id) _
            & " GROUP BY p.vehicle_id, c.route_id "
    End If
    
    
    Set rsTemp = oDB.Execute(szSql)
    Set GetSettleRouteCalRemark = rsTemp
    Exit Function
Here:

End Function

'*************************************************************************
'*************************************************************************
'************************** 以下为报表接口********************************
'*************************************************************************
'*************************************************************************



'路单结算明细
Public Function SettleSheetStat(dtStartDate As Date, dtEndDate As Date) As ADODB.Recordset
    Dim i As Integer
    Dim rsObject As Recordset
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    oDB.ConnectionString = GetConnectionStr()
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_SettleSheetStat
    
    
    
    szSql = "SELECT settlement_sheet_id,(SELECT CASE settle_object WHEN 1 THEN '车辆' WHEN 2 THEN '公司' END) settle_object, object_id , object_name  " _
        & " , total_ticket_price,total_quantity,(settle_price-settle_station_price) settle_other_price,settle_price , settle_station_price" _
        & " , convert(char(10),start_date,120) start_date ,convert(char(10),end_date,120) end_date  , convert(char(10),settle_date,120) settle_date  " _
        & " , luggage_settle_id , luggage_base_carriage , luggage_settle_price " _
        & " , f.luggage_protocol_name luggage_protocol_name " _
        & " FROM fin_settle_sheet_lst f WHERE start_date>=" & TransFieldValueToString(dtStartDate) _
        & " AND end_date<=" & TransFieldValueToString(dtEndDate) _
        & " AND status=" & TransFieldValueToString(CS_SettleSheetValid)
    Set rsTemp = oDB.Execute(szSql)
    Set SettleSheetStat = rsTemp
    
End Function

'公司结算明细
Public Function CompanySettleDetail(dtStartDate As Date, dtEndDate As Date, Optional pszComapanyID As String = "") As ADODB.Recordset
    Dim i As Integer
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szWhere As String
    Dim oDB As New RTConnection
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_CompanySettleDetail
    
    
    oDB.ConnectionString = GetConnectionStr()
    If pszComapanyID <> "" Then
        szWhere = " AND s.transport_company_id=" & TransFieldValueToString(pszComapanyID)
'    Else
'        szWhere = " AND 1=1"
    End If
    szSql = "SELECT f.settlement_sheet_id,c.transport_company_short_name,s.passenger_number,s.settle_price,(s.settle_price-s.settle_station_price) settle_other_price, s.settle_station_price  " _
        & " , split_item_1,split_item_2,split_item_3,split_item_4,split_item_5,split_item_6,split_item_7,split_item_8,split_item_9,split_item_10,split_item_11,split_item_12,split_item_13,split_item_14 " _
        & " , split_item_15,split_item_16,split_item_17,split_item_18,split_item_19,split_item_20 " _
        & " , convert(char(10),f.start_date,120) start_date ,convert(char(10),f.end_date,120) end_date  , convert(char(10),f.settle_date,120) settle_date " _
        & " , f.luggage_settle_id , f.luggage_base_carriage , f.luggage_settle_price " _
        & " , f.luggage_protocol_name luggage_protocol_name " _
        & " FROM fin_settle_sheet_lst f,fin_settle_company_price_lst s, " & cszTableCompany & " c" _
        & " WHERE f.settlement_sheet_id=s.settlement_sheet_id AND s.transport_company_id=c.transport_company_id AND f.status=" & TransFieldValueToString(CS_SettleSheetValid) _
        & " AND start_date>=" & TransFieldValueToString(dtStartDate) & " AND end_date<=" & TransFieldValueToString(dtEndDate) _
        & szWhere
    Set rsTemp = oDB.Execute(szSql)
    Set CompanySettleDetail = rsTemp
End Function

'车辆结算明细
Public Function VehicleSettleDetail(dtStartDate As Date, dtEndDate As Date, Optional pszVehicleID As String = "", Optional pszCompanyID As String = "", Optional pnStatus As Integer = CS_SettleSheetValid, Optional peQueryNegative As EQueryNegativeType = CS_QueryAll, Optional pbStatBySettleDate As Boolean = True) As ADODB.Recordset
    Dim i As Integer
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szWhere As String
    Dim oDB As New RTConnection
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_VehicleSettleDetail
    oDB.ConnectionString = GetConnectionStr()
    If pszVehicleID <> "" Then
        szWhere = " AND v.vehicle_id  IN ( " & pszVehicleID & ")"
    End If
    If pszCompanyID <> "" Then
        szWhere = szWhere & " AND c.transport_company_id IN ( " & pszCompanyID & ")"
    End If
    If peQueryNegative = CS_QueryNegative Then
        szWhere = szWhere & " and (v.settle_price - v.settle_station_price < 0 )"
    ElseIf peQueryNegative = CS_QueryNotNegative Then
        szWhere = szWhere & " and (v.settle_price - v.settle_station_price >= 0) "
    End If
    If pnStatus = CS_SettleSheetNotInvalid Then
        szWhere = szWhere & " AND f.status <> " & TransFieldValueToString(CS_SettleSheetInvalid)
    Else
        szWhere = szWhere & " AND f.status = " & TransFieldValueToString(pnStatus)
    End If
    
    If Not pbStatBySettleDate Then
        szWhere = szWhere & " AND start_date>= " & TransFieldValueToString(dtStartDate) & " AND  end_date< " & TransFieldValueToString(dtEndDate)
    Else
        szWhere = szWhere & " AND settle_date>= " & TransFieldValueToString(dtStartDate) & " AND  settle_date< " & TransFieldValueToString(dtEndDate)
    End If
    
    '& " AND start_date>=" & TransFieldValueToString(dtStartDate) & " AND  end_date<=" & TransFieldValueToString(dtEndDate)
    szSql = "SELECT h.license_tag_no,v.passenger_number,v.settle_price,(v.settle_price- v.settle_station_price) settle_other_price, v.settle_station_price , " _
        & " split_item_1,split_item_2,split_item_3,split_item_4,split_item_5,split_item_6,split_item_7,split_item_8,split_item_9,split_item_10,split_item_11,split_item_12,split_item_13,split_item_14," _
        & " split_item_15,split_item_16,split_item_17,split_item_18,split_item_19,split_item_20" _
        & "  ,convert(char(10),f.start_date,120) start_date ,convert(char(10),f.end_date,120) end_date  , convert(char(10),f.settle_date,120) settle_date  " _
        & " , f.luggage_settle_id , f.luggage_base_carriage , f.luggage_settle_price " _
        & " , f.luggage_protocol_name luggage_protocol_name , o.owner_id , o.owner_name " _
        & " , o.account_id owner_account_id , f.route_id , f.route_name , f.settlement_sheet_id ,case convert(char(10),isnull(f.remit_date,''),120) when '1900-01-01' then '' else convert(char(10),isnull(f.remit_date,''),120) end remit_date  " _
        & " FROM fin_settle_sheet_lst f,fin_settle_vehicle_lst v," & cszTableVehicle & " h , " & cszTableCompany & " c , owner_info o " _
        & " WHERE f.settlement_sheet_id=v.settlement_sheet_id AND v.vehicle_id=h.vehicle_id " _
        & " AND h.transport_company_id = c.transport_company_id AND h.owner_id = o.owner_id " _
        & szWhere
    Set rsTemp = oDB.Execute(szSql)
    Set VehicleSettleDetail = rsTemp
End Function

'公司结算汇总
Public Function CompanySettleStat(dtStartDate As Date, dtEndDate As Date, Optional pszComapanyID As String = "") As ADODB.Recordset
    Dim i As Integer
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szWhere As String
    Dim oDB As New RTConnection
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_CompanySettleStat
    
    
    oDB.ConnectionString = GetConnectionStr()
    If pszComapanyID <> "" Then
        szWhere = " AND s.transport_company_id=" & TransFieldValueToString(pszComapanyID)
'    Else
'        szWhere = " AND 1=1"
    End If
    szSql = "SELECT MAX(c.transport_company_short_name) transport_company_short_name ,SUM(s.passenger_number) passenger_number ,SUM(s.settle_price) settle_price,SUM(s.settle_price- s.settle_station_price) settle_other_price ,SUM(s.settle_station_price) settle_station_price , " _
        & " SUM(split_item_1) split_item_1,SUM(split_item_2) split_item_2,SUM(split_item_3) split_item_3 ,SUM(split_item_4) split_item_4 , " _
        & " SUM(split_item_5) split_item_5,SUM(split_item_6) split_item_6 ,SUM(split_item_7) split_item_7,SUM(split_item_8) split_item_8,SUM(split_item_9) split_item_9,SUM(split_item_10) split_item_10 ," _
        & " SUM(split_item_11) split_item_11,SUM(split_item_12) split_item_12,SUM(split_item_13) split_item_13 ,SUM(split_item_14) split_item_14 , " _
        & " SUM(split_item_15) split_item_15,SUM(split_item_16) split_item_16 ,SUM(split_item_17) split_item_17,SUM(split_item_18) split_item_18,SUM(split_item_19) split_item_19,SUM(split_item_20) split_item_20, " _
        & " SUM(f.luggage_base_carriage) luggage_base_carriage , SUM(f.luggage_settle_price) luggage_base_carriage " _
        & " FROM fin_settle_sheet_lst f,fin_settle_company_price_lst s, " & cszTableCompany & " c" _
        & " WHERE f.settlement_sheet_id=s.settlement_sheet_id AND s.transport_company_id=c.transport_company_id AND f.status=" & TransFieldValueToString(CS_SettleSheetValid) _
        & " AND start_date>=" & TransFieldValueToString(dtStartDate) & " AND  end_date<= " & TransFieldValueToString(dtEndDate) _
        & szWhere _
        & " GROUP BY s.transport_company_id"
    Set rsTemp = oDB.Execute(szSql)
    Set CompanySettleStat = rsTemp
End Function

'车辆结算汇总
Public Function VehicleSettleStat(dtStartDate As Date, dtEndDate As Date, Optional pszVehicleID As String = "", Optional pszCompanyID As String = "", Optional peQueryNegative As EQueryNegativeType = CS_QueryAll, Optional pnStatus As ESettleSheetStatus = CS_SettleSheetNotInvalid, Optional pbStatBySettleDate As Boolean = True) As ADODB.Recordset
    Dim i As Integer
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szWhere As String
    Dim oDB As New RTConnection
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_VehicleSettleStat
    
    
    oDB.ConnectionString = GetConnectionStr()
    If pszVehicleID <> "" Then
        szWhere = " AND v.vehicle_id  IN ( " & pszVehicleID & ")"
    End If
    If pszCompanyID <> "" Then
        szWhere = szWhere & " AND c.transport_company_id IN ( " & pszCompanyID & ")"
    End If
    

    If peQueryNegative = CS_QueryNegative Then
        szWhere = szWhere & " and (v.settle_price - v.settle_station_price < 0 )"
    ElseIf peQueryNegative = CS_QueryNotNegative Then
        szWhere = szWhere & " and (v.settle_price - v.settle_station_price >= 0) "
    End If
    If pnStatus = CS_SettleSheetNotInvalid Then
        szWhere = szWhere & " AND f.status <> " & TransFieldValueToString(CS_SettleSheetInvalid)
    Else
        szWhere = szWhere & " AND f.status = " & TransFieldValueToString(pnStatus)
    End If
    If Not pbStatBySettleDate Then
        szWhere = szWhere & " AND start_date>= " & TransFieldValueToString(dtStartDate) & " AND  end_date< " & TransFieldValueToString(dtEndDate)
    Else
        szWhere = szWhere & " AND settle_date>= " & TransFieldValueToString(dtStartDate) & " AND  settle_date< " & TransFieldValueToString(dtEndDate)
    End If
    szSql = "SELECT v.vehicle_id, MAX(o.owner_name) owner_name, MAX(f.route_id) route_id ,MAX(f.route_name) route_name , MAX(h.license_tag_no) license_tag_no,SUM(v.passenger_number) passenger_number ,SUM(v.settle_price) settle_price , " _
        & " SUM(v.settle_price- v.settle_station_price) settle_other_price ,SUM(v.settle_station_price) settle_station_price ," _
        & " SUM(split_item_1) split_item_1,SUM(split_item_2) split_item_2,SUM(split_item_3) split_item_3 ,SUM(split_item_4) split_item_4 , " _
        & " SUM(split_item_5) split_item_5,SUM(split_item_6) split_item_6 ,SUM(split_item_7) split_item_7,SUM(split_item_8) split_item_8, " _
        & " SUM(split_item_9) split_item_9,SUM(split_item_10) split_item_10 , " _
        & " SUM(split_item_11) split_item_11,SUM(split_item_12) split_item_12,SUM(split_item_13) split_item_13 ,SUM(split_item_14) split_item_14 , " _
        & " SUM(split_item_15) split_item_15,SUM(split_item_16) split_item_16 ,SUM(split_item_17) split_item_17,SUM(split_item_18) split_item_18, " _
        & " SUM(split_item_19) split_item_19,SUM(split_item_20) split_item_20 , " _
        & " SUM(f.luggage_base_carriage) luggage_base_carriage , SUM(f.luggage_settle_price) luggage_base_carriage " _
        & " FROM fin_settle_sheet_lst f,fin_settle_vehicle_lst v," & cszTableVehicle & " h, " & cszTableCompany & " c , " & cszTableOwnerInfo & " o " _
        & " WHERE f.settlement_sheet_id=v.settlement_sheet_id AND v.vehicle_id=h.vehicle_id " _
        & " AND h.transport_company_id = c.transport_company_id AND h.owner_id = o.owner_id " _
        & szWhere _
        & " GROUP BY v.vehicle_id"
    Set rsTemp = oDB.Execute(szSql)
    Set VehicleSettleStat = rsTemp
End Function

'车次结算汇总
Public Function BusSettleStat(dtStartDate As Date, dtEndDate As Date, Optional pszBusID As String = "", Optional pszCompanyID As String = "", Optional peQueryNegative As EQueryNegativeType = CS_QueryAll, Optional pnStatus As ESettleSheetStatus = CS_SettleSheetNotInvalid, Optional pbStatBySettleDate As Boolean = True) As ADODB.Recordset
    Dim i As Integer
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szWhere As String
    Dim oDB As New RTConnection
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_BusSettleStat
    
    
    oDB.ConnectionString = GetConnectionStr()
    If pszBusID <> "" Then
        szWhere = " AND v.bus_id  IN ( " & pszBusID & ")"
    End If
    If pszCompanyID <> "" Then
        szWhere = szWhere & " AND c.transport_company_id IN ( " & pszCompanyID & ")"
    End If
    

    If peQueryNegative = CS_QueryNegative Then
        szWhere = szWhere & " and (v.settle_price - v.settle_station_price < 0 )"
    ElseIf peQueryNegative = CS_QueryNotNegative Then
        szWhere = szWhere & " and (v.settle_price - v.settle_station_price >= 0) "
    End If
    If pnStatus = CS_SettleSheetNotInvalid Then
        szWhere = szWhere & " AND f.status <> " & TransFieldValueToString(CS_SettleSheetInvalid)
    Else
        szWhere = szWhere & " AND f.status = " & TransFieldValueToString(pnStatus)
    End If
    If Not pbStatBySettleDate Then
        szWhere = szWhere & " AND start_date>= " & TransFieldValueToString(dtStartDate) & " AND  end_date< " & TransFieldValueToString(dtEndDate)
    Else
        szWhere = szWhere & " AND settle_date>= " & TransFieldValueToString(dtStartDate) & " AND  settle_date< " & TransFieldValueToString(dtEndDate)
    End If
    szSql = "SELECT v.bus_id,v.transport_company_id,MAX(f.route_id) route_id ,MAX(f.route_name) route_name , SUM(v.passenger_number) passenger_number ,SUM(v.settle_price) settle_price , " _
        & " SUM(v.settle_price- v.settle_station_price) settle_other_price ,SUM(v.settle_station_price) settle_station_price ," _
        & " SUM(split_item_1) split_item_1,SUM(split_item_2) split_item_2,SUM(split_item_3) split_item_3 ,SUM(split_item_4) split_item_4 , " _
        & " SUM(split_item_5) split_item_5,SUM(split_item_6) split_item_6 ,SUM(split_item_7) split_item_7,SUM(split_item_8) split_item_8, " _
        & " SUM(split_item_9) split_item_9,SUM(split_item_10) split_item_10 , " _
        & " SUM(split_item_11) split_item_11,SUM(split_item_12) split_item_12,SUM(split_item_13) split_item_13 ,SUM(split_item_14) split_item_14 , " _
        & " SUM(split_item_15) split_item_15,SUM(split_item_16) split_item_16 ,SUM(split_item_17) split_item_17,SUM(split_item_18) split_item_18, " _
        & " SUM(split_item_19) split_item_19,SUM(split_item_20) split_item_20 , " _
        & " SUM(f.luggage_base_carriage) luggage_base_carriage , SUM(f.luggage_settle_price) luggage_base_carriage " _
        & " FROM fin_settle_sheet_lst f,fin_settle_bus_lst v, " & cszTableCompany & " c " _
        & " WHERE f.settlement_sheet_id=v.settlement_sheet_id " _
        & " AND v.transport_company_id = c.transport_company_id " _
        & szWhere _
        & " GROUP BY v.bus_id,v.transport_company_id order by v.bus_id,v.transport_company_id "
    Set rsTemp = oDB.Execute(szSql)
    Set BusSettleStat = rsTemp
End Function

'得到指这路单信息
Public Function GetCheckSheetInfo(pszSettleSheetID As String) As ADODB.Recordset
    Dim i As Integer
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    
    AssertObjIsValid
'    AssertHaveRight m_oActiveUser, RIGHT_CheckSheet
    
    
    oDB.ConnectionString = GetConnectionStr()
    szSql = "SELECT k.check_sheet_id,k.bus_date,k.bus_serial_no , k.bus_id,v.license_tag_no,c.transport_company_short_name, " _
        & " r.route_name,t.vehicle_type_name,o.owner_name, k.checker " _
        & " FROM " & cszTableCheckSheet & " k, " & cszTableCompany & " c," & cszTableRouteInfo & " r," _
        & cszTableOwnerInfo & " o," & cszTableVehicle & " v," & cszTableVehicleTypeCode & " t " _
        & " WHERE k.transport_company_id=c.transport_company_id AND k.route_id=r.route_id AND k.owner_id=o.owner_id " _
        & " AND k.vehicle_id=v.vehicle_id AND k.vehicle_type_code=t.vehicle_type_code " _
        & " AND k.settlement_sheet_id=" & TransFieldValueToString(pszSettleSheetID)
    Set rsTemp = oDB.Execute(szSql)
    Set GetCheckSheetInfo = rsTemp
End Function


Public Function GetVehileType(CompnayID As String, RouteID As String) As String()
    Dim i As Integer
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    Dim szTemp As String
    Dim aszVT() As String
    oDB.ConnectionString = GetConnectionStr()
    If CompnayID <> "" And RouteID <> "" Then
        szTemp = " , bus_info b , " & cszTableBusVehicle & " bv , " & cszTableVehicle & " v , " & cszTableCompany & "  c " _
            & " WHERE b.bus_id = bv.bus_id And bv.vehicle_id = v.vehicle_id And " _
            & " v.transport_company_id = C.transport_company_id " _
            & " AND b.route_id = '" & RouteID & "'  AND v.vehicle_type_code=vt.vehicle_type_code " _
            & " AND C.transport_company_id = '" & CompnayID & "' "
            
    End If
    szSql = "SELECT DISTINCT  vt.vehicle_type_code,vt.vehicle_type_short_name " _
        & " FROM vehicle_type_code vt  " & szTemp
    Dim szProjectID As String
    szProjectID = GetExecuteProject
    If szProjectID <> "" And InStr(1, szSql, "WHERE") > 0 Then
        szSql = szSql & " AND b.project_id = " & TransFieldValueToString(szProjectID)
    End If
        
        
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        ReDim aszVT(1 To rsTemp.RecordCount, 1 To 2)
        For i = 1 To rsTemp.RecordCount
            aszVT(i, 1) = FormatDbValue(rsTemp!vehicle_type_code)
            aszVT(i, 2) = FormatDbValue(rsTemp!vehicle_type_short_name)
            rsTemp.MoveNext
        Next i
    End If
    GetVehileType = aszVT
End Function


Public Function GetAllCompanyProtocol(Optional CompanyID As String = "", Optional RouteID As String = "", Optional ProtocolID As String = "") As String()
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim aszTemp() As String
    Dim szTemp As String, i As Integer
    
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_GetCompanyProtocol
    
    
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If Trim(CompanyID) <> "" Then
        szTemp = " AND f.transport_company_id = " & TransFieldValueToString(CompanyID)
    End If
    If Trim(ProtocolID) <> "" Then
        szTemp = szTemp + " AND f.protocol_id = " & TransFieldValueToString(ProtocolID)
    End If
    If Trim(RouteID) <> "" Then
        szTemp = szTemp + " AND f.route_id = " & TransFieldValueToString(RouteID)
    End If
'    szSql = "SELECT ISNULL( c.protocol_id,'') protocol_id ,ISNULL( p.protocol_name,'')protocol_name  ,ISNULL( c.transport_company_id ,'') transport_company_id  ,ISNULL( p.default_mark,'') default_mark ,ISNULL( c.transport_company_short_name,'') transport_company_name  " _
'        & " FROM " & cszTableCompany & " c LEFT JOIN fin_settle_protocol_info p ON c.protocol_id=p.protocol_id  WHERE 1=1 " & szTemp
'    szSql = "SELECT C.transport_company_id,ISNULL(f.route_id,'') route_id,ISNULL (f.protocol_id,'') protocol_id  ,c.transport_company_name" _
'        & " FROM fin_compnay_protocol_lst f right join transport_company_info c " _
'        & " ON f.transport_company_id = c.transport_company_id  " & szTemp
    szSql = "SELECT c.transport_company_id ,c.transport_company_short_name ,f.route_id,f.route_name ,f.protocol_id ,p.protocol_name" _
        & " FROM fin_compnay_protocol_lst f ,  " & cszTableCompany & " c ,fin_settle_protocol_info p" _
        & " WHERE f.transport_company_id = c.transport_company_id AND f.protocol_id = p.protocol_id " & szTemp
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 6)
    For i = 1 To rsTemp.RecordCount
        aszTemp(i, 1) = FormatDbValue(rsTemp!transport_company_id)
        aszTemp(i, 2) = FormatDbValue(rsTemp!transport_company_short_name)
        aszTemp(i, 3) = FormatDbValue(rsTemp!protocol_id)
        aszTemp(i, 4) = FormatDbValue(rsTemp!protocol_name)
        aszTemp(i, 5) = FormatDbValue(rsTemp!route_id)
        aszTemp(i, 6) = FormatDbValue(rsTemp!route_name)
        rsTemp.MoveNext
    Next i
    GetAllCompanyProtocol = aszTemp
End Function

Public Function GetCopmanyRoute(CompanyID As String) As String()
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim aszTemp() As String, szWhere As String
    Dim szTemp As String, i As Integer, nVehicleCount As Integer
    
    
    
    oDB.ConnectionString = GetConnectionStr("")
    
'    nVehicleCount = ArrayLength(CompanyID)
'    szWhere = " ( "
'    For i = 1 To nVehicleCount - 1
'
'            szWhere = szWhere & "'" & ResolveDisplay(CompanyID(i)) & "',"
'    Next i
'    szWhere = szWhere & "'" & ResolveDisplay(CompanyID(i)) & "')"
    
    szSql = " SELECT DISTINCT  b.route_id ,r.route_name " _
        & " FROM  " & cszTableCompany & " c , " & cszTableVehicle & " v ," & cszTableBusVehicle & " bv , bus_info b ,route_info r " _
        & " WHERE c.transport_company_id = v.transport_company_id AND v.vehicle_id = bv.vehicle_id AND b.route_id = r.route_id " _
        & " AND bv.bus_id = b.bus_id AND c.transport_company_id  = '" & ResolveDisplay(CompanyID) & "'"
    Dim szProjectID As String
    szProjectID = GetExecuteProject
    If szProjectID <> "" Then
        szSql = szSql & " AND b.project_id = " & TransFieldValueToString(szProjectID)
    End If
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 2)
    For i = 1 To rsTemp.RecordCount
        aszTemp(i, 1) = FormatDbValue(rsTemp!route_id)
        aszTemp(i, 2) = FormatDbValue(rsTemp!route_name)
        rsTemp.MoveNext
    Next i
    GetCopmanyRoute = aszTemp
End Function

''##ModelId=3E81F71A004E
'Public Function StatSettleSheetCompany(pszSheetID As String, pszCompnayID As String) As Recordset
'
'    AssertHaveRight m_oActiveUser, RIGHT_StatSettleSheetCompany
'    AssertObjIsValid
'    Dim oDB As New RTConnection
'    Dim rsTemp As Recordset
'    Dim szSql As String
'    oDB.ConnectionString = GetConnectionStr(cszSettle)
'    szSql = "SELECT f.*,c.company_name,c.protocol_id,p.protocol_name " _
'        & " FROM fin_settle_company_price_lst f," & cszTableCompany & " c,fin_settle_sheet_lst s,fin_settle_protocol_info p  " _
'        & " WHERE f.settlement_sheet_id=s.settlement_sheet_id AND f.transport_company_id = " & TransFieldValueToString(pszCompnayID) _
'        & " AND c.transport_company_id =f.transport_company_id AND f.settlement_sheet_id=" & TransFieldValueToString(pszSheetID) _
'        & " GROUP BY f.transport_company_id "
'    Set rsTemp = oDB.Execute(szSql)
'End Function
'
''##ModelId=3E81F71D029F
'Public Function StatSettleSheetVehicle(pszSheetID As String, pszVehiceID As String) As ADoDB.Recordset
'
'    AssertHaveRight m_oActiveUser, RIGHT_StatSettleSheetVehicle
'    AssertObjIsValid
'    Dim oDB As New RTConnection
'    Dim rsTemp As Recordset
'    Dim szSql As String
'    oDB.ConnectionString = GetConnectionStr(cszSettle)
'    szSql = "SELECT f.*,v.license_tag_no,c.protocol_id,p.protocol_name " _
'        & " FROM fin_settle_vehicle_lst f," & cszTableVehicle & " v,fin_settle_sheet_lst s,fin_settle_protocol_info p  " _
'        & " WHERE f.settlement_sheet_id=s.settlement_sheet_id AND f.vehicle_id = " & TransFieldValueToString(pszVehiceID) _
'        & " AND v.vehicle_id =f.vehicle_id AND f.settlement_sheet_id=" & TransFieldValueToString(pszSheetID) _
'        & " GROUP BY f.vehicle_id "
'    Set rsTemp = oDB.Execute(szSql)
'End Function



'##ModelId=3E81F73102DE
Public Function GetSettleVehicleLst(pszSettlementSheetID As String, Optional pszVehicleID As String = "") As ADODB.Recordset
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_SettleSheet
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If pszVehicleID = "" Then
    
        szSql = "SELECT i.license_tag_no, f.object_name object_name,p.protocol_name,v.* , (v.settle_price-v.settle_station_price) settle_other_price " _
            & " , f.luggage_settle_id , f.luggage_base_carriage , f.luggage_settle_price  " _
            & " ,f.luggage_protocol_name luggage_protocol_name , v.mileage " _
            & " FROM fin_settle_sheet_lst f,fin_settle_vehicle_lst v," & cszTableVehicle & " i,fin_settle_protocol_info p " _
            & " WHERE v.settlement_sheet_id = " & TransFieldValueToString(pszSettlementSheetID) _
            & " AND f.settlement_sheet_id=v.settlement_sheet_id AND v.vehicle_id=i.vehicle_id AND p.protocol_id=v.protocol_id "
    Else
        szSql = "SELECT i.license_tag_no, f.object_name object_name ,p.protocol_name,v.* , (v.settle_price-v.settle_station_price) settle_other_price " _
            & " , f.luggage_settle_id , f.luggage_base_carriage , f.luggage_settle_price " _
            & " ,f.luggage_protocol_name luggage_protocol_name, v.mileage " _
            & " FROM fin_settle_sheet_lst f,fin_settle_vehicle_lst v ," & cszTableVehicle & " i,fin_settle_protocol_info p" _
            & " WHERE v.settlement_sheet_id = " & TransFieldValueToString(pszSettlementSheetID) _
            & " AND v.vehicle_id = " & TransFieldValueToString(pszVehicleID) _
            & " AND f.settlement_sheet_id=v.settlement_sheet_id AND v.vehicle_id=f.vehicle_id AND p.protocol_id=v.protocol_id "
    End If
    Set rsTemp = oDB.Execute(szSql)
    Set GetSettleVehicleLst = rsTemp
    Set rsTemp = Nothing
End Function

Public Function GetSettleBusLst(pszSettlementSheetID As String, Optional pszBusID As String = "") As ADODB.Recordset
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_SettleSheet
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If pszBusID = "" Then
    
        szSql = "SELECT i.license_tag_no, f.object_name object_name,p.protocol_name,v.* , (v.settle_price-v.settle_station_price) settle_other_price " _
            & " , f.luggage_settle_id , f.luggage_base_carriage , f.luggage_settle_price  " _
            & " ,f.luggage_protocol_name luggage_protocol_name , v.mileage " _
            & " FROM fin_settle_sheet_lst f,fin_settle_bus_lst v," & cszTableVehicle & " i,fin_settle_protocol_info p " _
            & " WHERE v.settlement_sheet_id = " & TransFieldValueToString(pszSettlementSheetID) _
            & " AND f.settlement_sheet_id=v.settlement_sheet_id /*AND v.bus_id=i.vehicle_id*/ AND p.protocol_id=v.protocol_id "
    Else
        szSql = "SELECT i.license_tag_no, f.object_name object_name ,p.protocol_name,v.* , (v.settle_price-v.settle_station_price) settle_other_price " _
            & " , f.luggage_settle_id , f.luggage_base_carriage , f.luggage_settle_price " _
            & " ,f.luggage_protocol_name luggage_protocol_name, v.mileage " _
            & " FROM fin_settle_sheet_lst f,fin_settle_bus_lst v ," & cszTableVehicle & " i,fin_settle_protocol_info p" _
            & " WHERE v.settlement_sheet_id = " & TransFieldValueToString(pszSettlementSheetID) _
            & " AND v.bus_id = " & TransFieldValueToString(pszBusID) _
            & " AND f.settlement_sheet_id=v.settlement_sheet_id /*AND v.bus_id=f.vehicle_id*/ AND p.protocol_id=v.protocol_id "
    End If
    Set rsTemp = oDB.Execute(szSql)
    Set GetSettleBusLst = rsTemp
    Set rsTemp = Nothing
End Function

Public Function GetSettleBusLstSimple(pszSettlementSheetID As String, Optional pszBusID As String = "") As ADODB.Recordset
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_SettleSheet
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If pszBusID = "" Then
    
        szSql = "SELECT /*i.license_tag_no,*/ f.object_name object_name,p.protocol_name,v.* , (v.settle_price-v.settle_station_price) settle_other_price " _
            & " , f.luggage_settle_id , f.luggage_base_carriage , f.luggage_settle_price  " _
            & " ,f.luggage_protocol_name luggage_protocol_name , v.mileage " _
            & " FROM fin_settle_sheet_lst f,fin_settle_bus_lst v,/*" & cszTableVehicle & " i,*/fin_settle_protocol_info p " _
            & " WHERE v.settlement_sheet_id = " & TransFieldValueToString(pszSettlementSheetID) _
            & " AND f.settlement_sheet_id=v.settlement_sheet_id /*AND v.bus_id=i.vehicle_id*/ AND p.protocol_id=v.protocol_id "
    Else
        szSql = "SELECT i.license_tag_no, f.object_name object_name ,p.protocol_name,v.* , (v.settle_price-v.settle_station_price) settle_other_price " _
            & " , f.luggage_settle_id , f.luggage_base_carriage , f.luggage_settle_price " _
            & " ,f.luggage_protocol_name luggage_protocol_name, v.mileage " _
            & " FROM fin_settle_sheet_lst f,fin_settle_bus_lst v ," & cszTableVehicle & " i,fin_settle_protocol_info p" _
            & " WHERE v.settlement_sheet_id = " & TransFieldValueToString(pszSettlementSheetID) _
            & " AND v.bus_id = " & TransFieldValueToString(pszBusID) _
            & " AND f.settlement_sheet_id=v.settlement_sheet_id /*AND v.bus_id=f.vehicle_id*/ AND p.protocol_id=v.protocol_id "
    End If
    Set rsTemp = oDB.Execute(szSql)
    Set GetSettleBusLstSimple = rsTemp
    Set rsTemp = Nothing
End Function

'##ModelId=3E81F7350271
Public Function GetSettleCompanyLst(pszSettlementSheetID As String, Optional pszCompanyID As String = "") As ADODB.Recordset
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_SettleSheet
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If pszCompanyID = "" Then
        szSql = "SELECT f.object_name object_name,p.protocol_name,c.*,(c.settle_price-c.settle_station_price) settle_other_price " _
            & " , f.luggage_settle_id , f.luggage_base_carriage , f.luggage_settle_price " _
            & " ,f.luggage_protocol_name luggage_protocol_name , c.mileage , i.transport_company_short_name " _
            & " FROM fin_settle_sheet_lst f,fin_settle_company_price_lst c," & cszTableCompany & " i,fin_settle_protocol_info p " _
            & " WHERE c.settlement_sheet_id = " & TransFieldValueToString(pszSettlementSheetID) _
            & " AND f.settlement_sheet_id=c.settlement_sheet_id AND i.transport_company_id=c.transport_company_id AND p.protocol_id=c.protocol_id "
    Else
        szSql = "SELECT f.object_name object_name,p.protocol_name,c.* ,(c.settle_price-c.settle_station_price) settle_other_price  " _
            & " , f.luggage_settle_id , f.luggage_base_carriage , f.luggage_settle_price " _
            & " ,f.luggage_protocol_name luggage_protocol_name, c.mileage  , i.transport_company_short_name " _
            & " FROM fin_settle_sheet_lst f,fin_settle_company_price_lst c " & cszTableCompany & " i,fin_settle_protocol_info p" _
            & " WHERE c.settlement_sheet_id = " & TransFieldValueToString(pszSettlementSheetID) _
            & " AND c.transport_company_id = " & TransFieldValueToString(pszCompanyID) _
            & " AND f.settlement_sheet_id=c.settlement_sheet_id AND i.transport_company_id=c.transport_company_id AND p.protocol_id=c.protocol_id"
    End If
    Set rsTemp = oDB.Execute(szSql)
    Set GetSettleCompanyLst = rsTemp
    Set rsTemp = Nothing
End Function


Public Function GetAllSectionInfo(pszRouteID As String, Optional pszSellStationID As String = "", Optional pszSectionID As String = "") As String()
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset, i As Integer
    Dim szTemp As String
    Dim aszTemp() As String
    oDB.ConnectionString = GetConnectionStr("")
    If Trim(pszSellStationID) <> "" Then
        szTemp = " AND r.Sell_station_id = " & TransFieldValueToString(pszSectionID)
    End If
    If Trim(pszSectionID) <> "" Then
        szTemp = szTemp & " AND r.section_id = " & TransFieldValueToString(pszSectionID)
    End If
    szSql = "SELECT r.route_id,ri.route_name,s.station_id sell_station_id,s.sell_station_name ,r.end_station_id,si.station_name,r.section_id,ss.section_name,r.end_station_mileage" _
        & " FROM Route_section_lst r,route_info ri,sell_station_info s,station_info si,section_info ss " _
        & " WHERE  r.route_id= " & TransFieldValueToString(pszRouteID) _
        & " AND r.route_id = ri.route_id AND r.end_station_id =si.station_id AND r.sell_station_id = s.sell_station_id AND r.section_id = ss.section_id " _
        & szTemp _
        & " ORDER BY r.sell_station_id,r.section_serial"
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 10)
    For i = 1 To rsTemp.RecordCount
        aszTemp(i, 1) = FormatDbValue(rsTemp!route_id)
        aszTemp(i, 2) = FormatDbValue(rsTemp!route_name)
        aszTemp(i, 3) = FormatDbValue(rsTemp!sell_station_id)
        aszTemp(i, 4) = FormatDbValue(rsTemp!sell_station_name)
        aszTemp(i, 5) = FormatDbValue(rsTemp!end_station_id)
        aszTemp(i, 6) = FormatDbValue(rsTemp!station_name)
        aszTemp(i, 7) = FormatDbValue(rsTemp!section_id)
        aszTemp(i, 8) = FormatDbValue(rsTemp!section_name)
        aszTemp(i, 9) = FormatDbValue(rsTemp!end_station_mileage)
        rsTemp.MoveNext
    Next i
    GetAllSectionInfo = aszTemp
End Function



'查询路单
'pnStatus =路单状态
'pbNotIsSettled =1 表示,只取未结算或正在结算的路单   , 0 表示,取所有的结算状态路单
Public Function GetAllCheckSheetRS(pszCheckSheetID As String, pszBusID As String, pszVehicleID As String, pdtStartDate As Date, pdtEndDate As Date, Optional pnValidMark As Integer = -1, Optional pnBusSerial As Integer = -1, Optional pbNotIsSettled = 0) As Recordset


    Dim szWhere As String
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    Dim i As Integer
    On Error GoTo ErrorHandle
    
    
    AssertObjIsValid
'    AssertHaveRight m_oActiveUser, RIGHT_CheckSheet
    
    
    
    oDB.ConnectionString = GetConnectionStr("")
    
    If pnValidMark <> -1 Then
        szWhere = szWhere & " AND s.valid_mark = " & TransFieldValueToString(pnValidMark)
    End If
    If pnBusSerial <> -1 Then
        szWhere = szWhere & " AND s.bus_serial_no = " & TransFieldValueToString(pnBusSerial)
    End If
    If pszCheckSheetID <> "" Then
        szWhere = szWhere & " AND s.check_sheet_id like " & TransFieldValueToString(pszCheckSheetID & "%")
    End If
    If pszBusID <> "" Then
        szWhere = szWhere & " AND s.bus_id like " & TransFieldValueToString(pszBusID & "%")
    End If
    If pszVehicleID <> "" Then
        szWhere = szWhere & " AND s.vehicle_id = " & TransFieldValueToString(pszVehicleID)
    End If
    If pbNotIsSettled <> 0 Then
        szWhere = szWhere & " AND s.settlement_status <> " & TransFieldValueToString(ECheckSheetSettleStatus.CS_CheckSheetSettled)
    End If

    
    szSql = "SELECT s.* , c.transport_company_short_name transport_company_name , v.license_tag_no " _
        & " , CASE s.valid_mark WHEN 1 THEN '有效' WHEN 0 THEN '作废' END  valid_mark_name " _
        & " , CASE s.settlement_status WHEN 1 THEN '已经结算' WHEN 0 THEN '未结算' WHEN 2 THEN '正在结算' END settlement_status_name " _
        & " FROM " & cszTableCheckSheet & " s," & cszTableCompany & "  c ," & cszTableVehicle & " v " _
        & " WHERE bus_date >= " & TransFieldValueToString(pdtStartDate) & " AND bus_date <  " & TransFieldValueToString(pdtEndDate) _
        & " AND s.transport_company_id = c.transport_company_id AND s.vehicle_id = v.vehicle_id " & szWhere _
        
    Set rsTemp = oDB.Execute(szSql)
    
    
    Set GetAllCheckSheetRS = rsTemp
    Set rsTemp = Nothing
    Exit Function
ErrorHandle:
    err.Raise err.Number, err.Source, err.Description
    
    
    
End Function



'查询车辆的固定费用
Public Function GetAllVehicleFixFee(pszVehicleID As String, pszCompanyID As String, pdtStartDate As Date, pdtEndDate As Date, Optional pszLicenseTagNo As String, Optional pnIsDec As Integer) As Recordset


    Dim szWhere As String
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    Dim i As Integer
    On Error GoTo ErrorHandle
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_GetVehicleFixFee
    
    
    oDB.ConnectionString = GetConnectionStr("")
    
    If pszVehicleID <> "" Then
        szWhere = szWhere & " AND f.vehicle_id = " & TransFieldValueToString(pszVehicleID)
    End If
    If pszLicenseTagNo <> "" Then
        szWhere = szWhere & " AND v.license_tag_no LIKE " & TransFieldValueToString("%" & pszLicenseTagNo & "%")
    End If
    If pszCompanyID <> "" Then
        szWhere = szWhere & " AND v.transport_company_id = " & TransFieldValueToString(pszCompanyID)
    End If
    If pnIsDec <> -1 Then
        szWhere = szWhere & " AND f.is_dec = " & TransFieldValueToString(pnIsDec)
    
    End If
    
    szSql = "SELECT f.* , CASE f.is_dec WHEN 0 THEN '未扣' WHEN 1 THEN '已扣' END status_name , c.transport_company_short_name transport_company_name , v.license_tag_no " _
        & " FROM fin_vehicle_fix_fee_lst f," & cszTableCompany & "  c ," & cszTableVehicle & " v " _
        & " WHERE bus_date >= " & TransFieldValueToString(pdtStartDate) & " AND bus_date <  " & TransFieldValueToString(pdtEndDate) _
        & " AND v.transport_company_id = c.transport_company_id AND f.vehicle_id = v.vehicle_id " & szWhere _
        
    Set rsTemp = oDB.Execute(szSql)
    
    Set GetAllVehicleFixFee = rsTemp
    Set rsTemp = Nothing
    Exit Function
ErrorHandle:
    err.Raise err.Number, err.Source, err.Description
    
End Function

'车次结算明细
Public Function BusSettleDetail(dtStartDate As Date, dtEndDate As Date, Optional pszBusID As String = "", Optional pszCompanyID As String = "", Optional pnStatus As Integer = CS_SettleSheetValid, Optional peQueryNegative As EQueryNegativeType = CS_QueryAll, Optional pbStatBySettleDate As Boolean = True) As ADODB.Recordset
    Dim i As Integer
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szWhere As String
    Dim oDB As New RTConnection
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_BusSettleDetail
    oDB.ConnectionString = GetConnectionStr()
    If pszBusID <> "" Then
        szWhere = " AND v.bus_id  IN ( " & pszBusID & ")"
    End If
    If pszCompanyID <> "" Then
        szWhere = szWhere & " AND c.transport_company_id IN ( " & pszCompanyID & ")"
    End If
    If peQueryNegative = CS_QueryNegative Then
        szWhere = szWhere & " and (v.settle_price - v.settle_station_price < 0 )"
    ElseIf peQueryNegative = CS_QueryNotNegative Then
        szWhere = szWhere & " and (v.settle_price - v.settle_station_price >= 0) "
    End If
    If pnStatus = CS_SettleSheetNotInvalid Then
        szWhere = szWhere & " AND f.status <> " & TransFieldValueToString(CS_SettleSheetInvalid)
    Else
        szWhere = szWhere & " AND f.status = " & TransFieldValueToString(pnStatus)
    End If
    
    If Not pbStatBySettleDate Then
        szWhere = szWhere & " AND start_date>= " & TransFieldValueToString(dtStartDate) & " AND  end_date< " & TransFieldValueToString(dtEndDate)
    Else
        szWhere = szWhere & " AND settle_date>= " & TransFieldValueToString(dtStartDate) & " AND  settle_date< " & TransFieldValueToString(dtEndDate)
    End If
    
    '& " AND start_date>=" & TransFieldValueToString(dtStartDate) & " AND  end_date<=" & TransFieldValueToString(dtEndDate)
    szSql = "SELECT bus_id,v.passenger_number,v.settle_price,(v.settle_price- v.settle_station_price) settle_other_price, v.settle_station_price , " _
        & " split_item_1,split_item_2,split_item_3,split_item_4,split_item_5,split_item_6,split_item_7,split_item_8,split_item_9,split_item_10,split_item_11,split_item_12,split_item_13,split_item_14," _
        & " split_item_15,split_item_16,split_item_17,split_item_18,split_item_19,split_item_20" _
        & "  ,convert(char(10),f.start_date,120) start_date ,convert(char(10),f.end_date,120) end_date  , convert(char(10),f.settle_date,120) settle_date  " _
        & " , f.luggage_settle_id , f.luggage_base_carriage , f.luggage_settle_price " _
        & " , f.luggage_protocol_name luggage_protocol_name " _
        & " , f.route_id , f.route_name , f.settlement_sheet_id ,case convert(char(10),isnull(f.remit_date,''),120) when '1900-01-01' then '' else convert(char(10),isnull(f.remit_date,''),120) end remit_date  " _
        & " FROM fin_settle_sheet_lst f,fin_settle_bus_lst v, " & cszTableCompany & " c " _
        & " WHERE f.settlement_sheet_id=v.settlement_sheet_id " _
        & " AND v.transport_company_id = c.transport_company_id " _
        & szWhere
    Set rsTemp = oDB.Execute(szSql)
    Set BusSettleDetail = rsTemp
End Function

'车辆结算明细
Public Function VehicleSettleDetailEX(paszSettleSheet() As String) As ADODB.Recordset

    On Error GoTo ErrorHandle
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim i As Integer
    Dim szWhere As String
    Dim nCount As Integer
    Dim rsTemp As Recordset
    
    
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_VehicleSettleDetail
    
    
    oDB.ConnectionString = GetConnectionStr()
    
    nCount = ArrayLength(paszSettleSheet)
        
    For i = 1 To nCount - 1
        szWhere = szWhere & TransFieldValueToString(paszSettleSheet(i)) & " , "
    Next i
    
    szWhere = " AND f.settlement_sheet_id IN ( " & szWhere & TransFieldValueToString(paszSettleSheet(i)) & " ) "

    
    
    szSql = "SELECT h.license_tag_no,v.passenger_number,v.settle_price,(v.settle_price- v.settle_station_price) settle_other_price, v.settle_station_price , " _
        & " split_item_1,split_item_2,split_item_3,split_item_4,split_item_5,split_item_6,split_item_7,split_item_8,split_item_9,split_item_10,split_item_11,split_item_12,split_item_13,split_item_14," _
        & " split_item_15,split_item_16,split_item_17,split_item_18,split_item_19,split_item_20" _
        & "  ,convert(char(10),f.start_date,120) start_date ,convert(char(10),f.end_date,120) end_date  , convert(char(10),f.settle_date,120) settle_date  " _
        & " , f.luggage_settle_id , f.luggage_base_carriage , f.luggage_settle_price " _
        & " , f.luggage_protocol_name luggage_protocol_name , o.owner_id , o.owner_name " _
        & " , o.account_id owner_account_id , f.route_id , f.route_name , f.settlement_sheet_id  " _
        & " FROM fin_settle_sheet_lst f,fin_settle_vehicle_lst v," & cszTableVehicle & " h , " & cszTableCompany & " c , owner_info o " _
        & " WHERE f.settlement_sheet_id=v.settlement_sheet_id AND v.vehicle_id=h.vehicle_id  " _
        & " AND h.transport_company_id = c.transport_company_id AND h.owner_id = o.owner_id " _
        & " AND (v.settle_price - v.settle_station_price > 0 ) " _
        & szWhere
    Set rsTemp = oDB.Execute(szSql)
    Set VehicleSettleDetailEX = rsTemp
    Exit Function
ErrorHandle:
    err.Raise err.Number, err.Source, err.Description

End Function

'车次结算明细
Public Function BusSettleDetailEX(paszSettleSheet() As String) As ADODB.Recordset

    On Error GoTo ErrorHandle
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim i As Integer
    Dim szWhere As String
    Dim nCount As Integer
    Dim rsTemp As Recordset
    
    
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_VehicleSettleDetail
    
    
    oDB.ConnectionString = GetConnectionStr()
    
    nCount = ArrayLength(paszSettleSheet)
        
    For i = 1 To nCount - 1
        szWhere = szWhere & TransFieldValueToString(paszSettleSheet(i)) & " , "
    Next i
    
    szWhere = " AND f.settlement_sheet_id IN ( " & szWhere & TransFieldValueToString(paszSettleSheet(i)) & " ) "

    
    
    szSql = "SELECT b.bus_start_time,h.license_tag_no,v.passenger_number,v.settle_price,(v.settle_price- v.settle_station_price) settle_other_price, v.settle_station_price , " _
        & " split_item_1,split_item_2,split_item_3,split_item_4,split_item_5,split_item_6,split_item_7,split_item_8,split_item_9,split_item_10,split_item_11,split_item_12,split_item_13,split_item_14," _
        & " split_item_15,split_item_16,split_item_17,split_item_18,split_item_19,split_item_20" _
        & "  ,convert(char(10),f.start_date,120) start_date ,convert(char(10),f.end_date,120) end_date  , convert(char(10),f.settle_date,120) settle_date  " _
        & " , f.luggage_settle_id , f.luggage_base_carriage , f.luggage_settle_price " _
        & " , f.luggage_protocol_name luggage_protocol_name , o.owner_id , o.owner_name " _
        & " , o.account_id owner_account_id , f.route_id , f.route_name , f.settlement_sheet_id  " _
        & " FROM bus_info b,fin_settle_sheet_lst f,fin_settle_bus_lst v," & cszTableVehicle & " h , " & cszTableCompany & " c , owner_info o " _
        & " WHERE f.settlement_sheet_id=v.settlement_sheet_id /*AND v.vehicle_id=h.vehicle_id*/  " _
        & " AND h.transport_company_id = c.transport_company_id AND h.owner_id = o.owner_id and b.bus_id=v.bus_id " _
        & " AND (v.settle_price - v.settle_station_price > 0 ) " _
        & szWhere
    Set rsTemp = oDB.Execute(szSql)
    Set BusSettleDetailEX = rsTemp
    Exit Function
ErrorHandle:
    err.Raise err.Number, err.Source, err.Description

End Function



Public Function TotalSettleStationQuantity(pszSettleSheetID As String) As ADODB.Recordset
    '汇总结算单的站点人数
    
    On Error GoTo ErrorHandle
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    
    
    AssertObjIsValid
    oDB.ConnectionString = GetConnectionStr()
    szSql = " SELECT s.station_id , s.ticket_type_id , MAX(t.station_name) station_name  , MAX(ticket_type_name) ticket_type_name , SUM(quantity) quantity " _
        & " FROM fin_settle_station_lst s , station_info t , ticket_type_code y " _
        & " WHERE settlement_sheet_id = " & TransFieldValueToString(pszSettleSheetID) _
        & " AND s.station_id = t.station_id AND s.ticket_type_id = y.ticket_type_id " _
        & " GROUP BY s.station_id , s.ticket_type_id ORDER BY s.station_id , s.ticket_type_id "
    Set rsTemp = oDB.Execute(szSql)
    
    Set TotalSettleStationQuantity = rsTemp
    Exit Function
ErrorHandle:
    err.Raise err.Number, err.Source, err.Description
    
End Function



'车辆结算汇总 (按月汇总)
Public Function VehicleSettleStatByMonth(dtStartDate As Date, dtEndDate As Date, Optional pszVehicleID As String = "", Optional pszCompanyID As String = "", Optional peQueryNegative As EQueryNegativeType = CS_QueryAll, Optional pnStatus As ESettleSheetStatus = CS_SettleSheetNotInvalid) As ADODB.Recordset
    Dim i As Integer
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szWhere As String
    Dim oDB As New RTConnection
    On Error GoTo ErrorHandle
    
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_VehicleSettleStatByMonth
    
    
    oDB.ConnectionString = GetConnectionStr()
    If pszVehicleID <> "" Then
        szWhere = " AND v.vehicle_id  IN ( " & pszVehicleID & ")"
    End If
    If pszCompanyID <> "" Then
        szWhere = szWhere & " AND c.transport_company_id IN ( " & pszCompanyID & ")"
    End If
    
'    If peQueryNegative = CS_QueryNegative Then
'        szWhere = szWhere & " and (v.settle_price - v.settle_station_price < 0 ) "
'    ElseIf peQueryNegative = CS_QueryNotNegative Then
'        szWhere = szWhere & " and (v.settle_price - v.settle_station_price >= 0 ) "
'    End If
    
    If peQueryNegative = CS_QueryNegative Then
        szWhere = szWhere & " and (v.settle_price - v.settle_station_price < 0 )"
    ElseIf peQueryNegative = CS_QueryNotNegative Then
        szWhere = szWhere & " and (v.settle_price - v.settle_station_price >= 0) "
    End If
    If pnStatus = CS_SettleSheetNotInvalid Then
        szWhere = szWhere & " AND f.status <> " & TransFieldValueToString(CS_SettleSheetInvalid)
    Else
        szWhere = szWhere & " AND f.status = " & TransFieldValueToString(pnStatus)
    End If
    szSql = "SELECT v.vehicle_id , convert(char(7) , f.end_date ,120) end_date , MAX(o.owner_name) owner_name , MAX(h.license_tag_no) license_tag_no,SUM(v.passenger_number) passenger_number ,SUM(v.settle_price) settle_price , " _
        & " SUM(v.settle_price- v.settle_station_price) settle_other_price ,SUM(v.settle_station_price) settle_station_price ," _
        & " SUM(split_item_1) split_item_1,SUM(split_item_2) split_item_2,SUM(split_item_3) split_item_3 ,SUM(split_item_4) split_item_4 , " _
        & " SUM(split_item_5) split_item_5,SUM(split_item_6) split_item_6 ,SUM(split_item_7) split_item_7,SUM(split_item_8) split_item_8, " _
        & " SUM(split_item_9) split_item_9,SUM(split_item_10) split_item_10 , " _
        & " SUM(split_item_11) split_item_11,SUM(split_item_12) split_item_12,SUM(split_item_13) split_item_13 ,SUM(split_item_14) split_item_14 , " _
        & " SUM(split_item_15) split_item_15,SUM(split_item_16) split_item_16 ,SUM(split_item_17) split_item_17,SUM(split_item_18) split_item_18, " _
        & " SUM(split_item_19) split_item_19,SUM(split_item_20) split_item_20 , " _
        & " SUM(f.luggage_base_carriage) luggage_base_carriage , SUM(f.luggage_settle_price) luggage_base_carriage " _
        & " FROM fin_settle_sheet_lst f,fin_settle_vehicle_lst v," & cszTableVehicle & " h, " & cszTableCompany & " c , " & cszTableOwnerInfo & " o " _
        & " WHERE f.settlement_sheet_id=v.settlement_sheet_id AND v.vehicle_id=h.vehicle_id " _
        & " AND start_date>= " & TransFieldValueToString(dtStartDate) & " AND  end_date<= " & TransFieldValueToString(dtEndDate) _
        & " AND h.transport_company_id = c.transport_company_id AND h.owner_id = o.owner_id " _
        & szWhere _
        & " GROUP BY v.vehicle_id , convert(char(7) , f.end_date ,120) " _
        & " ORDER BY v.vehicle_id , convert(char(7) , f.end_date ,120) "
        
    Set rsTemp = oDB.Execute(szSql)
    Set VehicleSettleStatByMonth = rsTemp
    Exit Function
ErrorHandle:
    err.Raise err.Number, err.Source, err.Description
    
End Function



'========================================
'此接口为查询车辆结算平衡表中的检票部分用
'========================================
'车辆统计(按车次日期进行汇总)
Public Function GetCheckVehicleStatByBusDate(pszCompanyID As String, pdtFromDate As Date, pdtToDate As Date, Optional pszVehicleID As String = "", Optional pszSellStationID As String = "") As Recordset
'pnExtraStatus =1 表示售票，2表示补票，3表示全部

    AssertActiveUserValid m_oActiveUser, ERR_Report
'    AssertHaveRight m_oActiveUser, RIGHT_GetBusStatByBusDate
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If pszCompanyID <> "" Then
        szWhere = " AND split_company_id IN ( " & pszCompanyID & ")"
    End If
    '如果要查询的售票站不为空
    If pszSellStationID <> "" Then
'        szWhere = szWhere & " AND b.sell_station_id='" & pszSellStationID & "'"
        szWhere = szWhere & " AND b.sell_station_id IN ( " & pszSellStationID & ")"
    End If
'    If pnExtraStatus <> TP_AllSold Then
'    '如果不是所有的状态
'        szWhere = szWhere & " AND extra_status = " & pnExtraStatus
'
'    End If
    If pszVehicleID <> "" Then
        szWhere = szWhere & " AND vehicle_id= '" & pszVehicleID & "'"
        
    End If
    
    szSql = "   SELECT s.vehicle_id , b.status , MAX(v.license_tag_no) license_tag_no , MAX(s.route_id) route_id , MAX(route_name) route_name " _
            & "    , SUM(base_carriage*quantity) base_price , SUM (price_item_1*quantity) price_item_1, SUM(price_item_2*quantity) price_item_2   " _
            & "    , SUM(price_item_3*quantity) price_item_3 , SUM(price_item_4*quantity) price_item_4 , SUM(price_item_5*quantity) price_item_5  " _
            & "    , SUM(price_item_6*quantity) price_item_6 , SUM(price_item_7*quantity) price_item_7 ,SUM(price_item_8*quantity) price_item_8  " _
            & "    , SUM(price_item_9*quantity) price_item_9 , SUM(price_item_10*quantity) price_item_10 , SUM(price_item_11*quantity) price_item_11 " _
            & "    , SUM(price_item_12*quantity) price_item_12 , SUM(price_item_13*quantity) price_item_13 , SUM(price_item_14*quantity) price_item_14 " _
            & "    , SUM(price_item_15*quantity) price_item_15 " _
            & "    , SUM(quantity) AS quantity ,SUM(fact_quantity) AS fact_quantity " _
            & "    , SUM(ticket_price*quantity) AS ticket_price " _
            & " FROM  fin_sheet_station_lst b , route_info r , vehicle_info v , sheet_lst s  " _
            & " WHERE b.check_sheet_id = s.check_sheet_id AND s.vehicle_id = v.vehicle_id AND r.route_id = s.route_id AND s.valid_mark = " & TransFieldValueToString(ECheckSheetValidMark.CS_CheckSheetValid) _
            & " AND s.bus_date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND s.bus_date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "'" & szWhere & ExceptZeroBus _
            & " GROUP BY s.vehicle_id , b.status ORDER BY s.vehicle_id , b.status "
        
        '按照车辆及改并状态进行排序
            
    Set rsTemp = oDB.Execute(szSql)
    Set GetCheckVehicleStatByBusDate = rsTemp
    
End Function


'========================================
'此接口为查询车次结算平衡表中的检票部分用
'========================================
'车次统计(按车次日期进行汇总)
Public Function GetCheckBusStatByBusDate(pszCompanyID As String, pdtFromDate As Date, pdtToDate As Date, Optional pszLike As String = "", Optional pszSellStationID As String = "") As Recordset
'pnExtraStatus =1 表示售票，2表示补票，3表示全部


    AssertActiveUserValid m_oActiveUser, ERR_Report
'    AssertHaveRight m_oActiveUser, RIGHT_GetBusStatByBusDate
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim szWhere As String
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If pszCompanyID <> "" Then
        szWhere = " AND s.split_company_id IN ( " & pszCompanyID & ")"
    End If
    '如果要查询的售票站不为空
    If pszSellStationID <> "" Then
        szWhere = szWhere & " AND b.sell_station_id IN ( " & pszSellStationID & ")"
    End If
'    If pnExtraStatus <> TP_AllSold Then
'    '如果不是所有的状态
'        szWhere = szWhere & " AND extra_status = " & pnExtraStatus
'
'    End If
    If pszLike <> "" Then
        szWhere = szWhere & " AND b.bus_id like '" & pszLike & "%'"
    End If

    szSql = "   SELECT b.bus_id, s.transport_company_id ,b.status, MAX(s.bus_start_time) bus_start_time , MAX(s.route_id) route_id , MAX(route_name) route_name " _
            & "    , SUM(base_carriage*quantity) base_price , SUM (price_item_1*quantity) price_item_1, SUM(price_item_2*quantity) price_item_2   " _
            & "    , SUM(price_item_3*quantity) price_item_3 , SUM(price_item_4*quantity) price_item_4 , SUM(price_item_5*quantity) price_item_5  " _
            & "    , SUM(price_item_6*quantity) price_item_6 , SUM(price_item_7*quantity) price_item_7 ,SUM(price_item_8*quantity) price_item_8  " _
            & "    , SUM(price_item_9*quantity) price_item_9 , SUM(price_item_10*quantity) price_item_10 , SUM(price_item_11*quantity) price_item_11 " _
            & "    , SUM(price_item_12*quantity) price_item_12 , SUM(price_item_13*quantity) price_item_13 , SUM(price_item_14*quantity) price_item_14 " _
            & "    , SUM(price_item_15*quantity) price_item_15 " _
            & "    , SUM(quantity) AS quantity ,SUM(fact_quantity) AS fact_quantity " _
            & "    , SUM(ticket_price*quantity) AS ticket_price " _
            & " FROM  fin_sheet_station_lst b , route_info r , sheet_lst s  " _
            & " WHERE b.check_sheet_id = s.check_sheet_id AND r.route_id = s.route_id AND s.valid_mark = " & TransFieldValueToString(ECheckSheetValidMark.CS_CheckSheetValid) _
            & " AND s.bus_date >= '" & ToDBDateTime(pdtFromDate) & "'" _
            & " AND s.bus_date <'" & ToDBDateTime(DateAdd("d", 1, pdtToDate)) & "'" & szWhere & ExceptZeroBus _
            & " GROUP BY b.bus_id, s.transport_company_id,b.status  ORDER BY b.bus_id, s.transport_company_id,b.status  "
        '按照车次、公司及改并状态进行排序
        
        
    Set rsTemp = oDB.Execute(szSql)
    Set GetCheckBusStatByBusDate = rsTemp
    
End Function


'得到不是零的条件
Private Function ExceptZeroBus() As String
    ExceptZeroBus = "  AND NOT ( quantity  = 0  AND  ticket_price = 0 ) "
End Function






'得到所有的计划车次及所属公司
'在设置车辆公司协议时用到

Public Function GetAllBusCompany(Optional pszBusID As String, Optional pszCompanyID As String, Optional pszRouteID As String) As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szWhere As String
    
    
    On Error GoTo ErrorHandle
    
    AssertObjIsValid
'    AssertHaveRight m_oActiveUser, RIGHT_GetCompanyProtocol
    
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If pszBusID <> "" Then
        szWhere = szWhere & " AND b.bus_id LIKE " & TransFieldValueToString(pszBusID & "%")
    End If
    If pszCompanyID <> "" Then
        szWhere = szWhere & " AND v.transport_company_id = " & TransFieldValueToString(pszCompanyID)
    End If
    
    If pszRouteID <> "" Then
        szWhere = szWhere & " AND b.route_id= " & TransFieldValueToString(pszRouteID)
    End If

    szSql = "SELECT b.bus_id , v.transport_company_id , MAX(b.bus_start_time ) bus_start_time " _
        & " , MAX(b.route_id) route_id , MAX(c.transport_company_short_name) transport_company_short_name " _
        & " , MAX(r.route_name) route_name " _
        & " FROM bus_info b , bus_vehicle_code i , vehicle_info v,  company_info c , route_info r " _
        & " WHERE b.bus_id = i.bus_id And i.vehicle_id = v.vehicle_id " _
        & " AND v.transport_company_id =c.transport_company_id " _
        & " AND b.route_id = r.route_id " _
        & szWhere _
        & " GROUP BY b.bus_id , v.transport_company_id " _
        & " ORDER BY b.bus_id , v.transport_company_id "
    
    Set rsTemp = oDB.Execute(szSql)
    
    Set GetAllBusCompany = rsTemp
    Set rsTemp = Nothing
    Set oDB = Nothing
    
    Exit Function
    
ErrorHandle:

    Set rsTemp = Nothing
    Set oDB = Nothing
    err.Raise err.Number, err.Source, err.Description
End Function

'得到所有的计划车次协议
'在设置车辆公司协议时用到
'此接口中查询出来的记录中漏掉了,无公司的协议设置,以及协议设置后删除协议后的车次协议
Public Function GetAllBusProtocol(Optional pszBusID As String, Optional pszCompanyID As String, Optional pszRouteID As String, Optional pszProtocolID As String) As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szWhere As String
    
    
    On Error GoTo ErrorHandle
    
    AssertObjIsValid
'    AssertHaveRight m_oActiveUser, RIGHT_GetCompanyProtocol
    
    
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If pszBusID <> "" Then
        szWhere = szWhere & " AND p.bus_id LIKE " & TransFieldValueToString(pszBusID & "%")
    End If
    If pszCompanyID <> "" Then
        szWhere = szWhere & " AND p.transport_company_id = " & TransFieldValueToString(pszCompanyID)
    End If
    
    If pszRouteID <> "" Then
        szWhere = szWhere & " AND b.route_id= " & TransFieldValueToString(pszRouteID)
    End If
    
    If pszProtocolID <> "" Then
        szWhere = szWhere & " AND p.protocol_id= " & TransFieldValueToString(pszProtocolID)
    End If

    szSql = "SELECT p.* , s.protocol_name , s.default_mark  , c.transport_company_short_name  " _
        & " FROM fin_bus_protocol_lst p, bus_info b , fin_settle_protocol_info s,company_info c " _
        & " WHERE p.bus_id *= b.bus_id AND p.protocol_id = s.protocol_id AND c.transport_company_id=p.transport_company_id " _
        & szWhere _
    
    Set rsTemp = oDB.Execute(szSql)
    Set GetAllBusProtocol = rsTemp
    Set rsTemp = Nothing
    Set oDB = Nothing
    
    Exit Function
    
ErrorHandle:
    Set rsTemp = Nothing
    Set oDB = Nothing
    
    err.Raise err.Number, err.Source, err.Description
    

End Function

Public Function GetOtherBusProtocol(Optional pszBusID As String, Optional pszCompanyID As String, Optional pszProtocolID As String) As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szWhere As String


    On Error GoTo ErrorHandle

    AssertObjIsValid
'    AssertHaveRight m_oActiveUser, RIGHT_GetCompanyProtocol


    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If pszBusID <> "" Then
        szWhere = szWhere & " AND p.bus_id LIKE " & TransFieldValueToString(pszBusID & "%")
    End If
    If pszCompanyID <> "" Then
        szWhere = szWhere & " AND p.transport_company_id = " & TransFieldValueToString(pszCompanyID)
    End If

    If pszProtocolID <> "" Then
        szWhere = szWhere & " AND p.protocol_id= " & TransFieldValueToString(pszProtocolID)
    End If

    szSql = "select * " _
        & " from fin_bus_protocol_lst p where not bus_id in ( select bus_id from bus_info) " _
        & " or not transport_company_id in ( select transport_company_id from company_info ) " _
        & " or not protocol_id in (select protocol_id from fin_settle_protocol_info) " _
        & szWhere

    Set rsTemp = oDB.Execute(szSql)
    Set GetOtherBusProtocol = rsTemp
    Set rsTemp = Nothing
    Set oDB = Nothing

    Exit Function

ErrorHandle:
    Set rsTemp = Nothing
    Set oDB = Nothing

    err.Raise err.Number, err.Source, err.Description


End Function

'查询车次的固定费用
Public Function GetAllBusFixFee(pszBusID As String, Optional pszCompanyID As String, Optional pdtStartDate As Date, Optional pdtEndDate As Date, Optional pszLicenseTagNo As String, Optional pnIsDec As Integer) As Recordset


    Dim szWhere As String
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    Dim i As Integer
    On Error GoTo ErrorHandle
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_GetBusFixFee
    
    
    oDB.ConnectionString = GetConnectionStr("")
    
    If pszBusID <> "" Then
        szWhere = szWhere & " AND f.bus_id = " & TransFieldValueToString(pszBusID)
    End If
    If pszLicenseTagNo <> "" Then
        szWhere = szWhere & " AND v.license_tag_no LIKE " & TransFieldValueToString("%" & pszLicenseTagNo & "%")
    End If
    If pszCompanyID <> "" Then
        szWhere = szWhere & " AND f.transport_company_id = " & TransFieldValueToString(pszCompanyID)
    End If
    If pnIsDec <> -1 Then
        szWhere = szWhere & " AND f.is_dec = " & TransFieldValueToString(pnIsDec)
    
    End If
    
    szSql = "SELECT f.* , CASE f.is_dec WHEN 0 THEN '未扣' WHEN 1 THEN '已扣' END status_name , c.transport_company_short_name transport_company_name ,b.bus_start_time ,r.route_name " _
        & " FROM fin_bus_fix_fee_lst f,bus_info b," & cszTableCompany & "  c,route_info r " _
        & " WHERE bus_date >= " & TransFieldValueToString(pdtStartDate) & " AND bus_date <  " & TransFieldValueToString(pdtEndDate) _
        & " AND b.bus_id = f.bus_id" _
        & " AND b.route_id = r.route_id" _
        & " AND f.transport_company_id = c.transport_company_id " & szWhere

    Set rsTemp = oDB.Execute(szSql)
    
    Set GetAllBusFixFee = rsTemp
    Set rsTemp = Nothing
    Exit Function
ErrorHandle:
    err.Raise err.Number, err.Source, err.Description
    
End Function


'得到该结算单的手工补票人数信息
Public Function GetExtraInfo(pszSettleSheetID As String) As ADODB.Recordset
    Dim i As Integer
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    
    AssertObjIsValid
'    AssertHaveRight m_oActiveUser, RIGHT_CheckSheet
    
    
    oDB.ConnectionString = GetConnectionStr()
    szSql = "SELECT * " _
        & " FROM fin_settle_extra_lst  " _
        & " WHERE settlement_sheet_id= " & TransFieldValueToString(pszSettleSheetID)
    Set rsTemp = oDB.Execute(szSql)
    Set GetExtraInfo = rsTemp
    
End Function

Public Function GetCheckSheetStationListEx(paszCheckSheetID() As String, Optional pbIsToday As Boolean = False) As Recordset
    '得到路单站点明细玉环
    
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
    

    szSql = " SELECT sell_station_id,station_id,max(station_name) station_name,ticket_price,sum(quantity) quantity " _
    & " FROM  fin_sheet_station_lst f" _
    & " WHERE " _
    & " check_sheet_id IN " & szWhere _
    & " group by sell_station_id,station_id,station_name,ticket_price order by station_name,ticket_price "
    
    Set rsTemp = oDB.Execute(szSql)
    
    Set GetCheckSheetStationListEx = rsTemp
    
    Exit Function
ErrorHandle:
    Call err.Raise(err.Number, "Split:GetCheckSheetStationListEx Method", err.Description)

End Function

