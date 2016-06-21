VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompanySettlePrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E81DEBB02EE"
'
Option Base 0
Option Explicit

'set this to 0 to disable debug code in this class
#Const DebugMode = 1
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3E8219C6000F
    Private mlClassDebugID As Long
#End If

'##ModelId=3E81DF12000F
Private m_szCompanyID As String

'##ModelId=3E81DF18005D
Private m_szCompanyName As String

'##ModelId=3E81DF1D01F4
Private m_szVehicleTypeCode As String

'##ModelId=3E81DF230157
Private m_szVehicleTypeName As String

'##ModelId=3E81DF29005D
Private m_szRouteID As String

'##ModelId=3E81DF2E0399
Private m_szRouteName As String

'##ModelId=3E81DF50035B
Private m_szSellStationID As String

'##ModelId=3E81DF55001F
Private m_szSellStationName As String

'##ModelId=3E81DF5901A5
Private m_szStationID As String

'##ModelId=3E81DF5D01B5
Private m_szStationName As String

'##ModelId=3E81DF610157
Private m_dbMileage As Double

'##ModelId=3E81DF65038A
Private m_dbPassCharge As Double

'##ModelId=3E81DF6A0242
Private m_dbSettlefullPrice As Double

'##ModelId=3E81DF700399
Private m_dbSettleHalfPrice As Double

'##ModelId=3E81DF7502BF
Private m_szAnnotation As String
Private m_dbHalveFullPrice As Double
Private m_dbHalveHalfPrice As Double

Private m_dbServiceFullPrice As Double
Private m_dbServiceHalfPrice As Double

Private m_dbSpringFullPrice As Double
Private m_dbSpringHalfPrice As Double


Private m_oActiveUser As ActiveUser
Private m_nObjectStatus As EObjectStatus '对象状态
'##ModelId=3E81DF86038A
Public Sub Init(opAUser As ActiveUser)
    Set m_oActiveUser = opAUser
    m_nObjectStatus = ST_NormalObj
End Sub

'##ModelId=3E81DFA10290
Public Sub Identify(pszCompanyID As String, pszRouteID As String, pszVehicleTypeCode As String, pszSellstationID As String, pszStationID As String)
    AssertActiveUserValid m_oActiveUser, ERR_Company
    m_szCompanyID = pszCompanyID
    m_szRouteID = pszRouteID
    m_szVehicleTypeCode = pszVehicleTypeCode
    m_szSellStationID = pszSellstationID
    m_szStationID = pszStationID
    
    RefreshObject
    m_nObjectStatus = ST_NormalObj

End Sub
Private Sub RefreshObject()
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "SELECT f.transport_company_id ,c.transport_company_name,f." & cszTableVehicleTypeCode & ",v.vehicle_type_name ,f.route_id ,f.route_name  " _
        & " ,f.sell_station_id,s.station_name sell_station_name,f.station_id,si.station_name ,f.mileage,f.pass_charge,f.settle_full_price,f.settle_half_price " _
        & " ,f.service_full_price,f.service_half_price,f.halve_full_price,f.halve_half_price,f.spring_full_price,f.spring_half_price,f.annotation" _
        & " FROM fin_company_settle_price_lst f," & cszTableCompany & "  c," & cszTableVehicleTypeCode & " v, " & cszTableStationInfo & " s," & cszTableStationInfo & " si" _
        & " WHERE f.transport_company_id = " & TransFieldValueToString(m_szCompanyID) _
        & " AND f.vehicle_type_code = " & TransFieldValueToString(m_szVehicleTypeCode) _
        & " AND f.route_id = " & TransFieldValueToString(m_szRouteID) _
        & " AND f.sell_station_id = " & TransFieldValueToString(m_szSellStationID) _
        & " AND f.station_id = " & TransFieldValueToString(m_szStationID) _
        & " AND f.transport_company_id = c.transport_company_id AND f.vehicle_type_code=v.vehicle_type_code " _
        & " AND f.sell_station_id =s.station_id AND f.station_id = si.station_id "
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then RaiseError ERR_CompanySettlePriceNotExist
    m_szCompanyID = FormatDbValue(rsTemp!transport_company_id)
    m_szCompanyName = FormatDbValue(rsTemp!transport_company_name)
    m_szVehicleTypeCode = FormatDbValue(rsTemp!vehicle_type_code)
    m_szVehicleTypeName = FormatDbValue(rsTemp!vehicle_type_name)
    m_szRouteID = FormatDbValue(rsTemp!route_id)
    m_szRouteName = FormatDbValue(rsTemp!route_name)
    m_szSellStationID = FormatDbValue(rsTemp!sell_station_id)
    m_szSellStationName = FormatDbValue(rsTemp!sell_station_name)
    m_szStationID = FormatDbValue(rsTemp!station_id)
    m_szStationName = FormatDbValue(rsTemp!station_name)
    m_dbMileage = FormatDbValue(rsTemp!Mileage)
    m_dbPassCharge = FormatDbValue(rsTemp!pass_charge)
    m_dbSettlefullPrice = FormatDbValue(rsTemp!settle_full_price)
    m_dbSettleHalfPrice = FormatDbValue(rsTemp!settle_half_price)
    m_dbServiceFullPrice = FormatDbValue(rsTemp!service_full_price)
    m_dbServiceHalfPrice = FormatDbValue(rsTemp!service_half_price)
    m_dbHalveFullPrice = FormatDbValue(rsTemp!halve_full_price)
    m_dbHalveHalfPrice = FormatDbValue(rsTemp!halve_half_price)
    m_dbSpringFullPrice = FormatDbValue(rsTemp!spring_full_price)
    m_dbSpringHalfPrice = FormatDbValue(rsTemp!spring_half_price)
    
    m_szAnnotation = FormatDbValue(rsTemp!Annotation)
    
End Sub
'##ModelId=3E81DFA602CE
Public Sub AddNew()
    AssertActiveUserValid m_oActiveUser, ERR_CompanySettlePrice
    AssertObjIsValid
    m_nObjectStatus = ST_AddObj
    m_szCompanyID = ""
    m_szVehicleTypeCode = ""
    m_szRouteID = ""
    m_szRouteName = ""
    m_szSellStationID = ""
    m_szStationID = ""
    m_dbMileage = 0
    m_dbPassCharge = 0
    m_dbSettlefullPrice = 0
    m_dbSettleHalfPrice = 0
    m_dbServiceFullPrice = 0
    m_dbServiceHalfPrice = 0
    m_dbHalveFullPrice = 0
    m_dbHalveHalfPrice = 0
    m_dbSpringFullPrice = 0
    m_dbSpringHalfPrice = 0
    m_szAnnotation = ""
    Exit Sub
AddNewErr:
    Call RaiseError(MyUnhandledError, "CompanySettlePrice:AddNew Method")
End Sub

'##ModelId=3E81DFAC0186
Private Sub AddObject()
    On Error GoTo AddObjectErr
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "INSERT fin_company_settle_price_lst (transport_company_id,vehicle_type_code,route_id,route_name,sell_station_id " _
        & " , station_id,mileage,pass_charge,settle_full_price,settle_half_price,halve_full_price,halve_half_price " _
        & " , service_full_price,service_half_price,spring_full_price,spring_half_price ,annotation)" _
        & " VALUES ( " _
        & TransFieldValueToString(m_szCompanyID) & "," _
        & TransFieldValueToString(m_szVehicleTypeCode) & " ," _
        & TransFieldValueToString(m_szRouteID) & "," _
        & TransFieldValueToString(m_szRouteName) & "," _
        & TransFieldValueToString(m_szSellStationID) & "," _
        & TransFieldValueToString(m_szStationID) & "," _
        & TransFieldValueToString(m_dbMileage) & "," _
        & TransFieldValueToString(m_dbPassCharge) & "," _
        & TransFieldValueToString(m_dbSettlefullPrice) & "," _
        & TransFieldValueToString(m_dbSettleHalfPrice) & "," _
        & TransFieldValueToString(m_dbHalveFullPrice) & "," _
        & TransFieldValueToString(m_dbHalveHalfPrice) & "," _
        & TransFieldValueToString(m_dbServiceFullPrice) & "," _
        & TransFieldValueToString(m_dbServiceHalfPrice) & "," _
        & TransFieldValueToString(m_dbSpringFullPrice) & "," _
        & TransFieldValueToString(m_dbSpringHalfPrice) & "," _
        & TransFieldValueToString(m_szAnnotation) & ")"
    oDB.Execute (szSql)
    Exit Sub
AddObjectErr:
    Call RaiseError(ERR_CompanySettlePriceExistNotAdd)
End Sub

'##ModelId=3E81DFB003C8
Public Sub UpDate()
    AssertActiveUserValid m_oActiveUser, ERR_CompanySettlePrice
    AssertHaveRight m_oActiveUser, RIGHT_CompanySettlePrice
    AssertObjIsValid
    Select Case m_nObjectStatus
        Case ST_AddObj
            AddObject
        Case ST_EditObj
            UpDateToDB
        Case ST_NormalObj
        Case Else
        RaiseError ERR_NotAvailable
    End Select
    m_nObjectStatus = ST_NormalObj

End Sub

'##ModelId=3E81DFB503D8
Private Sub UpDateToDB()
    On Error GoTo UpDateDBErr
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "UPDATE fin_company_settle_price_lst SET " _
        & " mileage = " & TransFieldValueToString(m_dbMileage) & "," _
        & " pass_charge = " & TransFieldValueToString(m_dbPassCharge) & "," _
        & " route_name = " & TransFieldValueToString(m_szRouteName) & "," _
        & " settle_full_price = " & TransFieldValueToString(m_dbSettlefullPrice) & "," _
        & " settle_half_price = " & TransFieldValueToString(m_dbSettleHalfPrice) & "," _
        & " halve_full_price = " & TransFieldValueToString(m_dbHalveFullPrice) & "," _
        & " halve_half_price = " & TransFieldValueToString(m_dbHalveHalfPrice) & "," _
        & " service_full_price = " & TransFieldValueToString(m_dbServiceFullPrice) & "," _
        & " service_half_price =" & TransFieldValueToString(m_dbServiceHalfPrice) & "," _
        & " spring_full_price = " & TransFieldValueToString(m_dbSpringFullPrice) & "," _
        & " spring_half_price =" & TransFieldValueToString(m_dbSpringHalfPrice) & "," _
        & " annotation =" & TransFieldValueToString(m_szAnnotation) _
        & " WHERE transport_company_id = " & TransFieldValueToString(m_szCompanyID) & "AND" _
        & " vehicle_type_code = " & TransFieldValueToString(m_szVehicleTypeCode) & "AND" _
        & " route_id = " & TransFieldValueToString(m_szRouteID) & "AND" _
        & " sell_station_id = " & TransFieldValueToString(m_szSellStationID) & "AND" _
        & " station_id = " & TransFieldValueToString(m_szStationID)

    oDB.Execute (szSql)
    Exit Sub
UpDateDBErr:
    Call RaiseError(ERR_CompanySettlePriceExistNotAdd, "CompanySettlePrice:UpDateDB Method")
End Sub

'##ModelId=3E81DFBA0242
Public Sub Delete()
    AssertActiveUserValid m_oActiveUser, ERR_CompanySettlePrice
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_CompanySettlePrice
    DeleteObject
End Sub

'##ModelId=3E81DFBE02CE
Private Sub DeleteObject()
    On Error GoTo DeleteObjectErr
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "DELETE FROM fin_company_settle_price_lst WHERE  " _
        & " transport_company_id = " & TransFieldValueToString(m_szCompanyID) & "AND" _
        & " vehicle_type_code = " & TransFieldValueToString(m_szVehicleTypeCode) & " AND" _
        & " route_id = " & TransFieldValueToString(m_szRouteID) & "AND" _
        & " sell_station_id = " & TransFieldValueToString(m_szSellStationID) & "AND" _
        & " station_id = " & TransFieldValueToString(m_szStationID)
    oDB.Execute (szSql)
    Exit Sub
DeleteObjectErr:
    Call RaiseError(ERR_CompanySettlePriceNotExist)
End Sub

'##ModelId=3E82292C002E
Public Property Get Annotation() As String
    Annotation = m_szAnnotation
End Property

'##ModelId=3E82292B0196
Public Property Let Annotation(ByVal Value As String)
    m_szAnnotation = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E82292B003E
Public Property Get SettleHalfPrice() As Double
    SettleHalfPrice = m_dbSettleHalfPrice
End Property

'##ModelId=3E82292A01B5
Public Property Let SettleHalfPrice(ByVal Value As Double)
    m_dbSettleHalfPrice = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E82292A005D
Public Property Get SettlefullPrice() As Double
    SettlefullPrice = m_dbSettlefullPrice
End Property

'##ModelId=3E82292901D4
Public Property Let SettlefullPrice(ByVal Value As Double)
    m_dbSettlefullPrice = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E82292900CB
Public Property Get PassCharge() As Double
    PassCharge = m_dbPassCharge
End Property

'##ModelId=3E8229280232
Public Property Let PassCharge(ByVal Value As Double)
    m_dbPassCharge = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E8229280128
Public Property Get Mileage() As Double
    Mileage = m_dbMileage
End Property

'##ModelId=3E822927035B
Public Property Let Mileage(ByVal Value As Double)
    m_dbMileage = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E8229270271
Public Property Get StationName() As String
    StationName = m_szStationName
End Property

'##ModelId=3E8229270138
Public Property Get StationID() As String
    StationID = m_szStationID
End Property

'##ModelId=3E82292602EE
Public Property Let StationID(ByVal Value As String)
    m_szStationID = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E82292601C5
Public Property Get SellStationName() As String
    SellStationName = m_szSellStationName
End Property

'##ModelId=3E822926008C
Public Property Get SellstationID() As String
    SellstationID = m_szSellStationID
End Property

'##ModelId=3E8229250251
Public Property Let SellstationID(ByVal Value As String)
    m_szSellStationID = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E8229250167
Public Property Get RouteName() As String
    RouteName = m_szRouteName
End Property



Public Property Let RouteName(ByVal Value As String)
    m_szRouteName = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E822925003E
Public Property Get RouteID() As String
    RouteID = m_szRouteID
End Property

'##ModelId=3E8229240213
Public Property Let RouteID(ByVal Value As String)
    m_szRouteID = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E82292400FA
Public Property Get VehicleTypeName() As String
    VehicleTypeName = m_szVehicleTypeName
End Property

'##ModelId=3E82292303D8
Public Property Get VehicleTypeCode() As String
    VehicleTypeCode = m_szVehicleTypeCode
End Property

'##ModelId=3E82292301D4
Public Property Let VehicleTypeCode(ByVal Value As String)
    m_szVehicleTypeCode = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E82292300CB
Public Property Get CompanyName() As String
    CompanyName = m_szCompanyName
End Property

'##ModelId=3E82292203A9
Public Property Get CompanyID() As String
    CompanyID = m_szCompanyID
End Property


'##ModelId=3E82292201A5
Public Property Let CompanyID(ByVal Value As String)
    m_szCompanyID = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

Public Property Let HalveFullPrice(ByVal Value As Double)
    m_dbHalveFullPrice = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property
Public Property Get HalveFullPrice() As Double
    HalveFullPrice = m_dbHalveFullPrice
End Property

Public Property Let HalveHalfPrice(ByVal Value As Double)
    m_dbHalveHalfPrice = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

Public Property Get HalveHalfPrice() As Double
    HalveHalfPrice = m_dbHalveHalfPrice
End Property

Public Property Let ServiceFullPrice(ByVal Value As Double)
    m_dbServiceFullPrice = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

Public Property Get ServiceFullPrice() As Double
    ServiceFullPrice = m_dbServiceHalfPrice
End Property

Public Property Let ServiceHalfPrice(ByVal Value As Double)
    m_dbServiceHalfPrice = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

Public Property Get ServiceHalfPrice() As Double
    ServiceHalfPrice = m_dbHalveHalfPrice
End Property



Public Property Let SpringFullPrice(ByVal Value As Double)
    m_dbSpringFullPrice = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

Public Property Get SpringFullPrice() As Double
    SpringFullPrice = m_dbSpringFullPrice
End Property

Public Property Let SpringHalfPrice(ByVal Value As Double)
    m_dbSpringHalfPrice = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

Public Property Get SpringHalfPrice() As Double
    SpringHalfPrice = m_dbSpringHalfPrice
End Property



Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then RaiseError ERR_CompanySettlePriceNotAvailable
End Sub

'**************************************************
'Member Code:S4
'Brief Description:测试对象是否有效
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub AssertObjIsValid()
    AssertActiveUserValid m_oActiveUser, ERR_Protocol
    AssertStatusAvailable
End Sub
Public Sub DeleteRoute(CompnayID As String, RouteID As String)
    On Error GoTo DeleteObjectErr
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "DELETE FROM fin_company_settle_price_lst WHERE  " _
        & " transport_company_id = " & TransFieldValueToString(CompnayID) & "AND" _
        & " route_id = " & TransFieldValueToString(RouteID)
    oDB.Execute (szSql)
    Exit Sub
DeleteObjectErr:
    Call RaiseError(ERR_CompanySettlePriceNotExist)
End Sub


