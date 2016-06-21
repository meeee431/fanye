VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LugFinSplitSvr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'行包财务拆算服务
Option Base 0
Option Explicit
Public Enum RIGHTLuggageFinSplitSvr
'32800
    RIGHT_LugFinanceDetailStat = ERR_LuggageFinSplitSvr + cnMidRightBegin + cnMidRightStep * 1    '32906  行包营收拆算报表统计
    RIGHT_LugFinanceStat = ERR_LuggageFinSplitSvr + cnMidRightBegin + cnMidRightStep * 2    '  32912 行包营收拆算一览表统计
End Enum

Public Enum ObjectType '对象类型 0-拆帐公司 1-车辆 2-参运公司 3-车主 4-车次
    SplitCompanyType = 0
    VehicleType = 1
    TranportCompanyType = 2
    OwnerType = 3
    BusType = 4
End Enum

Const MyUnhandledError = 234
Private m_oActiveUser As ActiveUser
'初始化活动对象
'##ModelId=3E28DD4F0068
Public Sub Init(poAUser As ActiveUser)
    On Error GoTo InitErr
    Set m_oActiveUser = poAUser
    Exit Sub
InitErr:
    Call RaiseError(MyUnhandledError, "LugFinSplitSvr:Init Method")
End Sub
'查询待结算的行包签发单信息,返回签发单记录集
'##ModelId=3E2850C50273
Public Function GetWillSplitSheetRS(StartTime As Date, EndTime As Date, SellStationID As String, Optional AcceptType As Integer = -1, Optional BusID As String, _
                                    Optional VehicleID As String, Optional CompanyID As String, Optional SplitCompanyID As String, _
                                    Optional OwnerID As String) As ADODB.Recordset
    AssertActiveUserValid m_oActiveUser, ERR_LuggageFinSplitSvr
    On Error GoTo GetWillSplitSheetRSErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szTempWheSql As String
    Dim rsTemp As Recordset
'    StartTime = CDate(CStr(StartTime) + " 00:00:01")
'    EndTime = CDate(CStr(EndTime) + " 23:59:59")
    If AcceptType <> -1 Then
        szTempWheSql = " AND l.accept_type = " & TransFieldValueToString(AcceptType)
    Else
        szTempWheSql = ""
    End If
    If BusID <> "" Then
        szTempWheSql = szTempWheSql + " AND l.bus_id = " & TransFieldValueToString(BusID)
    Else
        szTempWheSql = szTempWheSql + ""
    End If
    
    If VehicleID <> "" Then
        szTempWheSql = szTempWheSql + " AND l.vehicle_id = " & TransFieldValueToString(VehicleID)
    Else
        szTempWheSql = szTempWheSql + ""
    End If
    
    If CompanyID <> "" Then
        szTempWheSql = szTempWheSql + " AND l.transport_company_id = " & TransFieldValueToString(CompanyID)
    Else
        szTempWheSql = szTempWheSql + ""
    End If
    
    If SplitCompanyID <> "" Then
        szTempWheSql = szTempWheSql + " AND l.split_company_id = " & TransFieldValueToString(SplitCompanyID)
    Else
        szTempWheSql = szTempWheSql + ""
    End If
    
    If OwnerID <> "" Then
        szTempWheSql = szTempWheSql + " AND l.owner_id = " & TransFieldValueToString(OwnerID)
    Else
        szTempWheSql = szTempWheSql + ""
    End If
    
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT l.sheet_id,l.accept_type,l.sell_station_id,l.bus_id,l.bus_date,l.vehicle_id,l.license_tag_no,l.transport_company_id,transport_company_short_name," _
    & "l.split_company_id,l.splict_company_short_name,l.route_id,l.route_name,l." & cszTableVehicleTypeCode & ",l.vehicle_type_name,l.owner_id,l.owner_name," _
    & "l.bus_start_time,l.valid_mark,p.protocol_id,p.protocol_name,l.checker, " _
    & "l.sheet_make_time,l.luggage_number,l.baggage_number,l.cal_weight,l.fact_weight,l.over_number,l.total_price," _
    & "l.price_item_1,l.price_item_2,l.price_item_3,l.price_item_4,l.price_item_5, " _
    & "l.price_item_6 , l.price_item_7, l.price_item_8, l.price_item_9, l.price_item_10, l.settlement_status " _
    & " FROM luggage_sheet_info l,luggage_settle_protocol_info p,vehicle_lug_protocol_info v " _
    & " WHERE l.settlement_status = 0 AND l.valid_mark=1 AND v.accept_type=l.accept_type " _
    & " AND l.vehicle_id=v.vehicle_id AND p.protocol_id=v.protocol_id " _
    & " AND l.sell_station_id = " & TransFieldValueToString(SellStationID) _
    & " AND l.sheet_make_time >= " & TransFieldValueToString(ToDBDateTime(StartTime)) _
    & " AND l.sheet_make_time <= " & TransFieldValueToString(ToDBDateTime(EndTime)) & szTempWheSql
    Set rsTemp = oDb.Execute(szSql)
    Set GetWillSplitSheetRS = rsTemp

    Exit Function
GetWillSplitSheetRSErr:
    Call RaiseError(MyUnhandledError, "LugFinSplitSvr:GetWillSplitSheetRS Method")
End Function

'查询无车辆协议的签发单
Public Function GetWillSplitSheetRSTemp(StartTime As Date, EndTime As Date, SellStationID As String, Optional AcceptType As Integer = -1, Optional BusID As String, _
                                    Optional VehicleID As String, Optional CompanyID As String, Optional SplitCompanyID As String, _
                                    Optional OwnerID As String) As ADODB.Recordset
    AssertActiveUserValid m_oActiveUser, ERR_LuggageFinSplitSvr
    On Error GoTo GetWillSplitSheetRSErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szTempWheSql As String
    Dim rsTemp As Recordset
'    StartTime = CDate(CStr(StartTime) + " 00:00:01")
'    EndTime = CDate(CStr(EndTime) + " 23:59:59")
    If AcceptType <> -1 Then
        szTempWheSql = " AND l.accept_type = " & TransFieldValueToString(AcceptType)
    Else
        szTempWheSql = ""
    End If
    If BusID <> "" Then
        szTempWheSql = szTempWheSql + " AND l.bus_id = " & TransFieldValueToString(BusID)
    Else
        szTempWheSql = szTempWheSql + ""
    End If
    
    If VehicleID <> "" Then
        szTempWheSql = szTempWheSql + " AND l.vehicle_id = " & TransFieldValueToString(VehicleID)
    Else
        szTempWheSql = szTempWheSql + ""
    End If
    
    If CompanyID <> "" Then
        szTempWheSql = szTempWheSql + " AND l.transport_company_id = " & TransFieldValueToString(CompanyID)
    Else
        szTempWheSql = szTempWheSql + ""
    End If
    
    If SplitCompanyID <> "" Then
        szTempWheSql = szTempWheSql + " AND l.split_company_id = " & TransFieldValueToString(SplitCompanyID)
    Else
        szTempWheSql = szTempWheSql + ""
    End If
    
    If OwnerID <> "" Then
        szTempWheSql = szTempWheSql + " AND l.owner_id = " & TransFieldValueToString(OwnerID)
    Else
        szTempWheSql = szTempWheSql + ""
    End If
    
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT l.sheet_id,l.accept_type,l.sell_station_id,l.bus_id,l.bus_date,l.vehicle_id,l.license_tag_no,l.transport_company_id,transport_company_short_name," _
    & "l.split_company_id,l.splict_company_short_name,l.route_id,l.route_name,l." & cszTableVehicleTypeCode & ",l.vehicle_type_name,l.owner_id,l.owner_name," _
    & "l.bus_start_time,l.valid_mark,l.checker, " _
    & "l.sheet_make_time,l.luggage_number,l.baggage_number,l.cal_weight,l.fact_weight,l.over_number,l.total_price," _
    & "l.price_item_1,l.price_item_2,l.price_item_3,l.price_item_4,l.price_item_5, " _
    & "l.price_item_6 , l.price_item_7, l.price_item_8, l.price_item_9, l.price_item_10, l.settlement_status " _
    & " FROM luggage_sheet_info l,vehicle_lug_protocol_info v " _
    & " WHERE l.settlement_status = 0 AND l.valid_mark=1 AND v.accept_type=l.accept_type " _
    & " AND l.vehicle_id=v.vehicle_id AND  v.protocol_id='' " _
    & " AND l.sell_station_id = " & TransFieldValueToString(SellStationID) _
    & " AND l.sheet_make_time >= " & TransFieldValueToString(ToDBDateTime(StartTime)) _
    & " AND l.sheet_make_time <= " & TransFieldValueToString(ToDBDateTime(EndTime)) & szTempWheSql
    Set rsTemp = oDb.Execute(szSql)
    Set GetWillSplitSheetRSTemp = rsTemp

    Exit Function
GetWillSplitSheetRSErr:
    Call RaiseError(MyUnhandledError, "LugFinSplitSvr:GetWillSplitSheetRS Method")
End Function

'判断是否启用默认协议
Public Function GetHaveProtocol() As ADODB.Recordset
  AssertActiveUserValid m_oActiveUser, ERR_LuggageFinSplitSvr
  On Error GoTo err
    Dim i As Integer
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT * FROM luggage_settle_protocol_info WHERE default_mark!=2"
    Set rsTemp = oDb.Execute(szSql)
    Set GetHaveProtocol = rsTemp
 Exit Function
err:
   Call RaiseError(MyUnhandledError, "LugFinSplitSvr:GetWillSplitSheetRS Method")
End Function

'查询结算单记录集
'##ModelId=3E28A9200174
Public Function GetFinSheetDetailRS(StartDate As Date, EndDate As Date, Optional SellStationID As String, Optional Status As Integer = -1, _
                                Optional VehicleID As String, Optional CompanyID As String, Optional SplitCompanyID As String, _
                                Optional OwnerID As String) As ADODB.Recordset
    AssertActiveUserValid m_oActiveUser, ERR_LuggageFinSplitSvr
    On Error GoTo GetFinanceSheetRSErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szTempWheSql As String
    Dim rsTemp As Recordset
    If SellStationID <> "" Then
        szTempWheSql = "AND sell_station_id = " & TransFieldValueToString(SellStationID)
    Else
        szTempWheSql = ""
    End If
    If Status <> -1 Then
        szTempWheSql = szTempWheSql + " AND status = " & TransFieldValueToString(Status)
    Else
        szTempWheSql = szTempWheSql + ""
    End If
    
    If VehicleID <> -1 Then
        szTempWheSql = szTempWheSql + " AND vehicle_id = " & TransFieldValueToString(VehicleID)
    Else
        szTempWheSql = szTempWheSql + ""
    End If
    If CompanyID <> -1 Then
        szTempWheSql = szTempWheSql + " AND transport_company_id = " & TransFieldValueToString(CompanyID)
    Else
        szTempWheSql = ""
    End If
    If SplitCompanyID <> -1 Then
        szTempWheSql = " AND split_company_id = " & TransFieldValueToString(SplitCompanyID)
    Else
        szTempWheSql = szTempWheSql + ""
    End If
            
    If OwnerID <> -1 Then
        szTempWheSql = szTempWheSql + " AND owner_id = " & TransFieldValueToString(OwnerID)
    Else
        szTempWheSql = szTempWheSql + ""
    End If
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT * FROM fin_luggage_sheet_lst " _
        & " WHERE settle_month >= " & TransFieldValueToString(ToDBDate(StartDate)) _
        & " AND settle_month < " & TransFieldValueToString(ToDBDate(EndDate)) & szTempWheSql
    Set rsTemp = oDb.Execute(szSql)
    Set GetFinSheetDetailRS = rsTemp
    

    Exit Function
GetFinanceSheetRSErr:
    Call RaiseError(MyUnhandledError, "LugFinSplitSvr:GetFinanceSheetRS Method")
End Function

Public Function GetFinanceSheetRS(StartDate As Date, EndDate As Date, Optional SellStationID, Optional Status As Integer = -1, _
                            Optional pnAcceptType As Integer = -1, Optional pszObjectID As String) As ADODB.Recordset
    AssertActiveUserValid m_oActiveUser, ERR_LuggageFinSplitSvr
    
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szTempWheSql As String
    Dim rsTemp As Recordset
    
    If SellStationID <> "" Then
        szTempWheSql = " AND s.sell_station_id = " & TransFieldValueToString(SellStationID)
    Else
        szTempWheSql = ""
    End If
    
    If Status <> -1 Then
        szTempWheSql = szTempWheSql + " AND status = " & TransFieldValueToString(Status)
    Else
        szTempWheSql = szTempWheSql + ""
    End If
    
    If pnAcceptType <> -1 Then
        szTempWheSql = szTempWheSql + " AND accept_type = " & TransFieldValueToString(pnAcceptType)
    Else
        szTempWheSql = szTempWheSql + ""
    End If
    
    If pszObjectID <> "" Then
        szTempWheSql = szTempWheSql + " AND split_object_id = " & TransFieldValueToString(pszObjectID)
    Else
        szTempWheSql = szTempWheSql + ""
    End If
    
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT f.*,s.sell_station_name FROM fin_luggage_sheet_info f,sell_station_info s " _
        & " WHERE f.sell_station_id=s.sell_station_id AND settle_month >= " & TransFieldValueToString(ToDBDate(StartDate)) _
        & " AND   settle_month <= " & TransFieldValueToString(ToDBDate(EndDate)) & szTempWheSql
        
    Set rsTemp = oDb.Execute(szSql)
    Set GetFinanceSheetRS = rsTemp
    
    

End Function

'行包营收拆算一览表统计，返回统计记录集
'统计结果以每个拆帐公司汇总一条记录（行包类型、总运费、协议名称、应拆费用等元素）
'##ModelId=3E28D3980094
Public Function LugFinanceStat(StartMonth As Date, EndMonth As Date, SplitCompanyID() As String, Optional SellStationID As String, _
                                Optional AcceptType As ELuggageAcceptType = -1) As ADODB.Recordset
    AssertActiveUserValid m_oActiveUser, ERR_LuggageFinSplitSvr
    AssertHaveRight m_oActiveUser, RIGHT_LugFinanceStat
    On Error GoTo LugFinanceStatErr

    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szTempWheSql As String
    Dim szTempGroSql As String
    Dim rsTemp As Recordset
    Dim aszSplitCompany() As String
    Dim szSplitCompany As String
    aszSplitCompany = SplitCompanyID
    If ArrayLength(aszSplitCompany) Then
        szSplitCompany = "("
        Dim i As Integer
        For i = 1 To ArrayLength(aszSplitCompany) - 1
            szSplitCompany = szSplitCompany & TransFieldValueToString(aszSplitCompany(i)) & ","
        Next i
        szSplitCompany = szSplitCompany & TransFieldValueToString(aszSplitCompany(i)) & ")"
    Else
        '出错
    End If
    
    If SellStationID <> "" Then
        szTempWheSql = " AND i.sell_station_id = " & TransFieldValueToString(SellStationID)
        szTempGroSql = " ,i.sell_station_id "
    Else
        szTempWheSql = ""
        szTempGroSql = ""
    End If
    
    If AcceptType <> -1 Then
        szTempWheSql = szTempWheSql + " AND l.accept_type = " & TransFieldValueToString(AcceptType)
        szTempGroSql = szTempGroSql + " ,i.accept_type "
    Else
        szTempWheSql = szTempWheSql + ""
        szTempGroSql = szTempGroSql + " ,i.accept_type "
    End If
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT MAX(i.settle_month) as settle_month,MAX(i.split_object_name) split_object_name,(case i.accept_type when 0 then '快件' when 1 then '随行' end) as accept_type, " _
        & " MAX(s.sell_station_name) as sell_station_name,sum(i.total_price) as total_price,sum(i.need_split_out) as need_split_out  " _
        & " FROM fin_luggage_sheet_info i,sell_station_info s " _
        & " WHERE i.sell_station_id=s.sell_station_id AND i.operate_date >= " & TransFieldValueToString(StartMonth) _
        & " AND i.operate_date <= " & TransFieldValueToString(EndMonth) _
        & " AND i.split_object_id IN " & szSplitCompany & szTempWheSql _
        & " AND i.status=1 AND i.split_object_type=2 " _
        & " GROUP BY i.split_object_id " & szTempGroSql
        Set rsTemp = oDb.Execute(szSql)
    Set LugFinanceStat = rsTemp
    Exit Function
LugFinanceStatErr:
    Call RaiseError(MyUnhandledError, "LugFinSplitSvr:LugFinanceStat Method")
End Function

'打印结算单
Public Function PrintLugFinSheet(SheetID As String) As ADODB.Recordset
    On Error GoTo err
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT distinct i.*,l.protocol_name " _
        & " FROM  fin_luggage_sheet_info i,fin_luggage_sheet_lst l WHERE i.fin_sheet_id=l.fin_sheet_id AND i.fin_sheet_id=" & TransFieldValueToString(SheetID)
    Set rsTemp = oDb.Execute(szSql)
    Set PrintLugFinSheet = rsTemp

    Exit Function
err:
    Call RaiseError(MyUnhandledError, "LugFinSplitSvr:LugFinanceStat Method")
End Function
'重打结算单
Public Function RSPrintLugFinSheet(SheetID As String, OldSheetID As String) As ADODB.Recordset
    On Error GoTo err
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    oDb.BeginTrans
    '新增新结算单
    szSql = "INSERT INTO fin_luggage_sheet_info SELECT " & TransFieldValueToString(SheetID) & ",sell_station_id," _
        & " settle_month,accept_type,settlement_start_time,settlement_end_time,total_price,split_object_type," _
        & " split_object_id,split_object_name,need_split_out,act_split_out,status,operator,operate_date,annotation " _
        & " FROM fin_luggage_sheet_info WHERE fin_sheet_id=" & TransFieldValueToString(OldSheetID)
    oDb.Execute szSql
    
    szSql = "INSERT INTO fin_luggage_sheet_lst SELECT " & TransFieldValueToString(SheetID) & ",vehicle_id,sell_station_id," _
        & " settle_month,accept_type,status,total_price,need_split_out, " _
        & " protocol_id,protocol_name,license_tag_no,transport_company_id,transport_company_short_name,owner_id,owner_name, " _
        & " split_company_id,split_company_name,route_id,route_name " _
        & " FROM fin_luggage_sheet_lst WHERE fin_sheet_id=" & TransFieldValueToString(OldSheetID)
    oDb.Execute szSql
    
    szSql = "INSERT INTO fin_luggage_split_detail_lst SELECT " & TransFieldValueToString(SheetID) & ",vehicle_id," _
        & " charge_code,charge_name,split_out_money " _
        & " FROM fin_luggage_split_detail_lst WHERE fin_sheet_id=" & TransFieldValueToString(OldSheetID)
    oDb.Execute szSql
    
    szSql = "INSERT INTO fin_sheet_to_luggage_lst SELECT " & TransFieldValueToString(SheetID) & ",luggage_sheet_id " _
        & " FROM fin_sheet_to_luggage_lst WHERE fin_sheet_id=" & TransFieldValueToString(OldSheetID)
    oDb.Execute szSql
    
    
    '作废原结算单
    
    szSql = "UPDATE fin_luggage_sheet_info SET status=" & TransFieldValueToString(ELuggageSettleValidMark.LuggageNotValid) _
        & " WHERE fin_sheet_id=" & TransFieldValueToString(OldSheetID)
    oDb.Execute szSql
        
    szSql = "UPDATE fin_luggage_sheet_lst SET status=" & TransFieldValueToString(ELuggageSettleValidMark.LuggageNotValid) _
        & " WHERE fin_sheet_id=" & TransFieldValueToString(OldSheetID)
    oDb.Execute szSql
    
    
    '显示打印信息
        szSql = "SELECT distinct i.*,l.protocol_name " _
        & " FROM  fin_luggage_sheet_info i,fin_luggage_sheet_lst l WHERE i.fin_sheet_id=l.fin_sheet_id AND i.fin_sheet_id=" & TransFieldValueToString(SheetID)
    Set rsTemp = oDb.Execute(szSql)
    Set RSPrintLugFinSheet = rsTemp
    oDb.CommitTrans
    Set rsTemp = Nothing
    Exit Function
err:
    oDb.RollbackTrans
    Set rsTemp = Nothing
    Call RaiseError(MyUnhandledError, "LugFinSplitSvr:LugFinanceStat Method")
End Function

'查询结算单信息
Public Function GetFinSheetInfo(SheetID As String) As ADODB.Recordset
    On Error GoTo err
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT * FROM fin_luggage_sheet_info WHERE fin_sheet_id=" & TransFieldValueToString(SheetID)
    Set rsTemp = oDb.Execute(szSql)
    Set GetFinSheetInfo = rsTemp
    
    Set rsTemp = Nothing
        Exit Function
err:
    Set rsTemp = Nothing
    Call RaiseError(MyUnhandledError, "LugFinSplitSvr:LugFinanceStat Method")
End Function

'行包营收拆算报表统计，返回统计记录集
'统计结果以每辆车汇总一条记录（车牌号、行包类型、总运费、协议名称、应拆费用、是否已结等元素）
'##ModelId=3E28E11B0201
Public Function LugFinanceDetailStat(StartMonth As Date, EndMonth As Date, VehicleID() As String, Optional SellStationID As String, _
                                    Optional AcceptType As ELuggageAcceptType = -1, Optional SplitCompanyID As String) As ADODB.Recordset
    AssertActiveUserValid m_oActiveUser, ERR_LuggageFinSplitSvr
    AssertHaveRight m_oActiveUser, RIGHT_LugFinanceDetailStat
    On Error GoTo LugFinanceDetailStatErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim szTempWheSql As String
    Dim aszTempVehicle() As String
    Dim szVehicle As String
    aszTempVehicle = VehicleID
    If ArrayLength(aszTempVehicle) Then
        szVehicle = "("
        Dim i As Integer
        For i = 1 To ArrayLength(aszTempVehicle) - 1
            szVehicle = szVehicle & TransFieldValueToString(aszTempVehicle(i)) & ","
        Next i
        szVehicle = szVehicle & TransFieldValueToString(aszTempVehicle(i)) & ")"
    Else
        '出错
    End If
    
    If SellStationID <> "" Then
        szTempWheSql = " AND l.sell_station_id = " & TransFieldValueToString(SellStationID)
    Else
        szTempWheSql = ""
    End If
    
    If AcceptType <> -1 Then
        szTempWheSql = szTempWheSql + " AND l.accept_type = " & TransFieldValueToString(AcceptType)
    Else
        szTempWheSql = szTempWheSql + ""
    End If
    
    If SplitCompanyID <> "" Then
        szTempWheSql = szTempWheSql + " AND l.split_company_id = " & TransFieldValueToString(SplitCompanyID)
    Else
        szTempWheSql = szTempWheSql + ""
    End If
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT MAX(l.settle_month) as settle_month,MAX(l.license_tag_no) AS license_tag_no,SUM(l.total_price) total_price , " _
        & " MAX(l.transport_company_short_name) as transport_company_short_name , " _
        & " (case l.accept_type when 0 then '快件' when 1 then '随行' end) as accept_type," _
        & " MAX(l.protocol_name) as protocol_name,SUM(l.need_split_out) as need_split_out," _
        & " (case l.status when 0 then '作废' when 1 then '已结'  end) as status " _
        & " FROM fin_luggage_sheet_lst l,fin_luggage_sheet_info i " _
        & " WHERE l.fin_sheet_id = i.fin_sheet_id AND i.operate_date >= " & TransFieldValueToString(ToDBDate(StartMonth)) _
        & " AND i.operate_date <= " & TransFieldValueToString(ToDBDate(EndMonth)) _
        & " AND l.vehicle_id IN " & szVehicle & szTempWheSql _
        & " AND i.status=1 " _
        & " GROUP BY l.vehicle_id,l.accept_type,l.status "
        
    Set rsTemp = oDb.Execute(szSql)
    Set LugFinanceDetailStat = rsTemp

    Exit Function
LugFinanceDetailStatErr:
    Call RaiseError(MyUnhandledError, "LugFinSplitSvr:LugFinanceDetailStat Method")
End Function


'返回拆帐公司集合
Public Function GetSplitCompany() As ADODB.Recordset
 On Error GoTo ErrHandle
  Dim i As Integer
  Dim oDb As New RTConnection
  Dim szSql As String
  Dim rsTemp As Recordset
  oDb.ConnectionString = GetConnectionStr(cszSplit)
  szSql = "Select split_company_id,split_company_name FROM fin_luggage_sheet_lst GROUP BY split_company_id,split_company_name "
  Set rsTemp = oDb.Execute(szSql)
  Set GetSplitCompany = rsTemp
  Set rsTemp = Nothing
 Exit Function
ErrHandle:
 Set oDb = Nothing
 Set rsTemp = Nothing
 Call RaiseError(MyUnhandledError, "")
End Function

'返回结算单中车辆集合
Public Function GetVehicleInfo(mCompanyID As String) As ADODB.Recordset
   On Error GoTo ErrHandle
  Dim i As Integer
  Dim oDb As New RTConnection
  Dim szSql As String
  Dim rsTemp As Recordset
  oDb.ConnectionString = GetConnectionStr(cszSplit)
  szSql = "Select vehicle_id,license_tag_no FROM fin_luggage_sheet_lst WHERE transport_company_id=" & TransFieldValueToString(mCompanyID) _
         & " GROUP BY vehicle_id,license_tag_no "
  Set rsTemp = oDb.Execute(szSql)
  Set GetVehicleInfo = rsTemp
  Set rsTemp = Nothing
 Exit Function
ErrHandle:
 Set oDb = Nothing
 Set rsTemp = Nothing
 Call RaiseError(MyUnhandledError, "")
End Function

'自动增加结算单号用
Public Function GetFinSheetID() As ADODB.Recordset
 On Error GoTo ErrHandle
  Dim i As Integer
  Dim oDb As New RTConnection
  Dim szSql As String
  Dim rsTemp As Recordset
  oDb.ConnectionString = GetConnectionStr(cszSplit)
  szSql = "SELECT fin_sheet_id FROM fin_luggage_sheet_lst WHERE fin_sheet_id =(SELECT MAX(fin_sheet_id) AS fin_sheet_id FROM fin_luggage_sheet_lst )"
        
  Set rsTemp = oDb.Execute(szSql)
  Set GetFinSheetID = rsTemp
  Set rsTemp = Nothing
 Exit Function
ErrHandle:
 Set oDb = Nothing
 Set rsTemp = Nothing
 Call RaiseError(MyUnhandledError, "")
End Function


''行包员每日结算报表
'Public Function StatDayAccept(UserID() As String, StartTime As Date, EndTime As Date) As ADODB.Recordset
'    On Error GoTo StatDayAcceptErr
'
'    AssertActiveUserValid m_oActiveUser, ERR_LuggageSheet
''    AssertHaveRight m_oActiveUser, RIGHT_LuggageSvrNormalAccept
'
'    Dim szSql As String
'    Dim szUsers As String
'    If ArrayLength(UserID) Then
'        szUsers = "("
'        Dim i As Integer
'        For i = 1 To ArrayLength(UserID) - 1
'            szUsers = szUsers & TransFieldValueToString(UserID(i)) & ","
'        Next i
'        szUsers = szUsers & TransFieldValueToString(UserID(i)) & ")"
'    End If
'
'    Dim oDb As New RTConnection
'    oDb.ConnectionString = GetConnectionStr
'    '统计正常受理的行包单
'    szSql = "(SELECT i.user_id,i.accept_type,SUM(price_total) price_total,SUM(price_item_1) price_item_1,SUM(price_item_2) price_item_2,SUM(price_item_3) price_item_3" & _
'            ",SUM(price_item_4) price_item_4,SUM(price_item_5) price_item_5,SUM(price_item_6) price_item_6,SUM(price_item_7) price_item_7" & _
'            ",SUM(price_item_8) price_item_8,SUM(price_item_9) price_item_9,SUM(price_item_10) price_item_10" & _
'            " FROM luggage_accept_info i,luggage_price_lst p" & _
'            " WHERE i.luggage_id=p.luggage_id AND i.status=" & ST_LuggageAcceptNormal & IIf(szUsers <> "", " AND i.user_id IN" & szUsers, "") & _
'            " AND i.operation_time>=" & TransFieldValueToString(ToDBDateTime(StartTime)) & " AND i.operation_time<=" & TransFieldValueToString(ToDBDateTime(EndTime)) & _
'            " GROUP BY i.user_id,i.accept_type) a"
'    '统计退理的行包手续费
'    szSql = szSql & " FULL JOIN " & _
'            "(SELECT i.user_id,i.accept_type,SUM(r.return_charge) return_charge" & _
'            " FROM luggage_accept_info i,luggage_return_lst r" & _
'            " WHERE i.luggage_id=r.luggage_id AND i.status=" & ST_LuggageAcceptReturn & IIf(szUsers <> "", " AND r.user_id IN" & szUsers, "") & _
'            " AND i.operation_time>=" & TransFieldValueToString(ToDBDateTime(StartTime)) & " AND i.operation_time<=" & TransFieldValueToString(ToDBDateTime(EndTime)) & _
'            " GROUP BY i.user_id,accept_type) b"
'
'    '将正常受理和退运的按照操作员进行全联接
'    szSql = " SELECT a.user_id,(CASE a.accept_type WHEN 0 THEN '" & GetLuggageTypeString(0) & "' WHEN 1 THEN '" & GetLuggageTypeString(1) & "' END) accept_type ,ISNULL(price_total,-return_charge) price_total,ISNULL(price_item_1,0) price_item_1,ISNULL(price_item_1,0) price_item_1" & _
'            " ,ISNULL(price_item_2,0) price_item_2,ISNULL(price_item_3,0) price_item_3,ISNULL(price_item_4,0) price_item_4,ISNULL(price_item_5,0) price_item_5" & _
'            ",ISNULL(price_item_6,0) price_item_6,ISNULL(price_item_7,0) price_item_7,ISNULL(price_item_8,0) price_item_8,ISNULL(price_item_9,0) price_item_9" & _
'            ",ISNULL(price_item_10,0) price_item_10,ISNULL(return_charge,0) return_charge" & _
'            " FROM " & szSql & " ON a.user_id=b.user_id AND a.accept_type=b.accept_type"
'
'    Set StatDayAccept = oDb.Execute(szSql)
'
'    Exit Function
'StatDayAcceptErr:
'    Call RaiseError(err.Number, "LuggageSheet:StatDayAccept Method", err.Description)
'End Function

Public Function GetAllProtocol() As TLugProtocol()
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim atTemp() As TLugProtocol
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT * FROM luggage_settle_protocol_info "
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    ReDim atTemp(1 To rsTemp.RecordCount)
    For i = 1 To rsTemp.RecordCount
        atTemp(i).ProtocolID = FormatDbValue(rsTemp!protocol_id)
        atTemp(i).ProtocolName = FormatDbValue(rsTemp!protocol_name)
        atTemp(i).IsDefault = FormatDbValue(rsTemp!default_mark)
        atTemp(i).Annotation = FormatDbValue(rsTemp!Annotation)
        rsTemp.MoveNext
    Next i
    GetAllProtocol = atTemp
End Function
























