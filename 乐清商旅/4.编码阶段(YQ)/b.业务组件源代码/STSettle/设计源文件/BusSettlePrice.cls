VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BusSettlePrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
Option Base 0
Option Explicit

'set this to 0 to disable debug code in this class
#Const DebugMode = 1
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3E8219C20196
    Private mlClassDebugID As Long
#End If

'##ModelId=3E80FC5D0222
Private m_szBusID As String

'##ModelId=3E80FC630148
Private m_szTransportCompanyID As String
Private m_szTransportCompanyName As String

'##ModelId=3E80FC6901C5
Private m_szRouteID As String

'##ModelId=3E80FC6F00EA
Private m_szRouteName As String

'##ModelId=3E80FC730261
Private m_szSellStationID As String

'##ModelId=3E80FC7801C5
Private m_szSellStationName As String

'##ModelId=3E80FC7D002E
Private m_szStationID As String

'##ModelId=3E80FC8102FD
Private m_szStationName As String

'##ModelId=3E80FC86037A
Private m_dbMileage As Double

'##ModelId=3E80FC8B037A
Private m_dbPassCharge As Double

'##ModelId=3E80FC91002E
Private m_dbSettlefullPrice As Double

'##ModelId=3E80FC96001F
Private m_dbSettleHalfPrice As Double

'##ModelId=3E80FC9A00AB
Private m_szAnnotation As String
Private m_oActiveUser As ActiveUser
Private m_nObjectStatus As EObjectStatus '对象状态


Private m_dbHalveFullPrice As Double
Private m_dbHalveHalfPrice As Double
Private m_dbServiceFullPrice As Double
Private m_dbServiceHalfPrice As Double

Private m_dbSpringFullPrice As Double
Private m_dbSpringHalfPrice As Double



'##ModelId=3E80FD37006D
Public Sub Init(opAUser As ActiveUser)
    On Error GoTo InitErr

    Set m_oActiveUser = opAUser
    m_nObjectStatus = ST_NormalObj
    Exit Sub
InitErr:
    Call RaiseError(MyUnhandledError, "VehicleSettlePrice:Init Method")
End Sub

'##ModelId=3E80FD3C0119
Public Sub Identify(pszBusID As String, pszTransportCompanyID As String, pszSellStationID As String, pszStationID As String)
    On Error GoTo IdentifyErr
    m_szBusID = pszBusID
    m_szTransportCompanyID = pszTransportCompanyID
    m_szSellStationID = pszSellStationID
    m_szStationID = pszStationID
    RefreshObject
    m_nObjectStatus = ST_NormalObj
    Exit Sub
IdentifyErr:
    Call RaiseError(MyUnhandledError, "VehicleSettlePrice:Identify Method")
End Sub
Private Sub RefreshObject()
        Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "SELECT f.bus_id ,f.transport_company_id, c.transport_company_short_name " _
        & " ,f.sell_station_id,s.station_name sell_station_name,f.station_id,si.station_name ,f.mileage,f.pass_charge,f.settle_full_price,f.settle_half_price,f.service_full_price , f.service_half_price,f.halve_full_price , f.halve_half_price , f.annotation , f.spring_full_price , f.spring_half_price " _
        & " FROM fin_bus_settle_price_lst f," & cszTableCompany & " c," & cszTableStationInfo & " si , " & cszTableStationInfo & " s  " _
        & " WHERE f.bus_id = " & TransFieldValueToString(m_szBusID) _
        & " AND f.transport_company_id = " & TransFieldValueToString(m_szTransportCompanyID) _
        & " AND f.sell_station_id = " & TransFieldValueToString(m_szSellStationID) _
        & " AND f.station_id = " & TransFieldValueToString(m_szStationID) _
        & " AND f.transport_company_id = c.transport_company_id  " _
        & " AND f.sell_station_id =s.station_id AND f.station_id = si.station_id "
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then RaiseError ERR_VehicleSettlePriceExistNotAdd
    m_szBusID = FormatDbValue(rsTemp!bus_id)
    m_szTransportCompanyID = FormatDbValue(rsTemp!transport_company_id)
    m_szTransportCompanyName = FormatDbValue(rsTemp!transport_company_short_name)
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
'##ModelId=3E80FD410271
Public Sub AddNew()
    On Error GoTo AddNewErr
    AssertActiveUserValid m_oActiveUser, ERR_VehicleSettlePrice
    m_nObjectStatus = ST_AddObj
    m_szBusID = ""
    
    m_szTransportCompanyID = ""
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
    Call RaiseError(MyUnhandledError, "VehicleSettlePrice:AddNew Method")
End Sub

'##ModelId=3E80FD480177
Private Sub AddObject()
    On Error GoTo AddObjecErr
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "INSERT fin_bus_settle_price_lst (bus_id,transport_company_id,sell_station_id,station_id,mileage,pass_charge,settle_full_price," _
        & " settle_half_price,halve_full_price,halve_half_price,service_full_price,service_half_price,annotation)" _
        & " VALUES ( " _
        & TransFieldValueToString(m_szBusID) & "," _
        & TransFieldValueToString(m_szTransportCompanyID) & "," _
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
        & TransFieldValueToString(m_szAnnotation) & ")"
    oDB.Execute (szSql)
    Exit Sub
AddObjecErr:
    Call RaiseError(ERR_VehicleSettlePriceExistNotAdd)
End Sub

'##ModelId=3E80FD4D01A5
Public Sub UpDate()
    AssertActiveUserValid m_oActiveUser, ERR_VehicleSettlePrice
    AssertHaveRight m_oActiveUser, RIGHT_VehicleSettlePrice
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

'##ModelId=3E80FD51032C
Private Sub UpDateToDB()
    On Error GoTo UpDateDBErr
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "UPDATE fin_bus_settle_price_lst SET " _
        & " mileage = " & TransFieldValueToString(m_dbMileage) & "," _
        & " pass_charge = " & TransFieldValueToString(m_dbPassCharge) & "," _
        & " settle_full_price = " & TransFieldValueToString(m_dbSettlefullPrice) & "," _
        & " settle_half_price = " & TransFieldValueToString(m_dbSettleHalfPrice) & "," _
        & " annotation =" & TransFieldValueToString(m_szAnnotation) & "," _
        & " halve_full_price = " & TransFieldValueToString(m_dbHalveFullPrice) & "," _
        & " halve_half_price = " & TransFieldValueToString(m_dbHalveHalfPrice) & "," _
        & " service_full_price = " & TransFieldValueToString(m_dbServiceFullPrice) & "," _
        & " service_half_price = " & TransFieldValueToString(m_dbServiceHalfPrice) & " " _
        & " WHERE  bus_id = " & TransFieldValueToString(m_szBusID) & "AND" _
        & " transport_company_id = " & TransFieldValueToString(m_szTransportCompanyID) & "AND" _
        & " sell_station_id = " & TransFieldValueToString(m_szSellStationID) & "AND" _
        & " station_id = " & TransFieldValueToString(m_szStationID)
    oDB.Execute (szSql)
    Exit Sub
UpDateDBErr:
    Call RaiseError(ERR_VehicleSettlePriceExistNotAdd)
End Sub

'##ModelId=3E80FD5600BB
Public Sub Delete()
    AssertActiveUserValid m_oActiveUser, ERR_VehicleSettlePrice
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_VehicleSettlePrice
    DeleteObject
End Sub

'##ModelId=3E80FD5A02FD
Private Sub DeleteObject()
    On Error GoTo DeleteObjectErr
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "DELETE FROM fin_bus_settle_price_lst " _
        & " WHERE bus_id = " & TransFieldValueToString(m_szBusID) & "AND" _
        & " transport_company_id = " & TransFieldValueToString(m_szTransportCompanyID) & "AND" _
        & " sell_station_id = " & TransFieldValueToString(m_szSellStationID) & "AND" _
        & " station_id = " & TransFieldValueToString(m_szStationID)

    oDB.Execute (szSql)
    Exit Sub
DeleteObjectErr:
    Call RaiseError(ERR_VehicleSettlePriceNotExist)
End Sub

'##ModelId=3E8219C403B9
'Private Sub Class_Initialize()
'    #If DebugMode Then
'        'get the next available class ID, and print out
'        'that the class was created successfully
'        mlClassDebugID = GetNextClassDebugID()
'        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
'    #End If
'End Sub
'
''##ModelId=3E8219C5007D
'Private Sub Class_Terminate()
'    #If DebugMode Then
'    'the class is being destroyed
'    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
'    #End If
'End Sub
'
'#If DebugMode Then
'    '##ModelId=3E8225D900EA
'    Public Property Get ClassDebugID() As Long
'        'if we are in debug mode, surface this property that consumers can query
'        ClassDebugID = mlClassDebugID
'    End Property
'#End If

'##ModelId=3E82292001C5
Public Property Get Annotation() As String
    Annotation = m_szAnnotation
End Property

'##ModelId=3E82291F032C
Public Property Let Annotation(ByVal Value As String)
    m_szAnnotation = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E82291F01E4
Public Property Get SettleHalfPrice() As Double
    SettleHalfPrice = m_dbSettleHalfPrice
End Property

'##ModelId=3E82291E037A
Public Property Let SettleHalfPrice(ByVal Value As Double)
    m_dbSettleHalfPrice = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E82291E0251
Public Property Get SettlefullPrice() As Double
    SettlefullPrice = m_dbSettlefullPrice
End Property

'##ModelId=3E82291E003E
Public Property Let SettlefullPrice(ByVal Value As Double)
    m_dbSettlefullPrice = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E82291D02FD
Public Property Get PassCharge() As Double
    PassCharge = m_dbPassCharge
End Property

'##ModelId=3E82291D00DA
Public Property Let PassCharge(ByVal Value As Double)
    m_dbPassCharge = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E82291C038A
Public Property Get Mileage() As Double
    Mileage = m_dbMileage
End Property

'##ModelId=3E82291C0186
Public Property Let Mileage(ByVal Value As Double)
    m_dbMileage = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E82291C004E
Public Property Get StationName() As String
    StationName = m_szStationName
End Property

'##ModelId=3E82291B030D
Public Property Get StationID() As String
    StationID = m_szStationID
End Property

'##ModelId=3E82291B0109
Public Property Let StationID(ByVal Value As String)
    m_szStationID = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E82291A0157
Public Property Get SellStationName() As String
    SellStationName = m_szSellStationName
End Property

'##ModelId=3E82291901F4
Public Property Get SellstationID() As String
    SellstationID = m_szSellStationID
End Property

'##ModelId=3E822919003E
Public Property Let SellstationID(ByVal Value As String)
    m_szSellStationID = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E822918033C
Public Property Get RouteName() As String
    RouteName = m_szRouteName
End Property

'##ModelId=3E82291C0186
Public Property Let RouteName(ByVal Value As String)
    m_szRouteName = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E8229180251
Public Property Get RouteID() As String
    RouteID = m_szRouteID
End Property

'##ModelId=3E822918008C
Public Property Let RouteID(ByVal Value As String)
    m_szRouteID = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property


'##ModelId=3E822917029F
Public Property Get BusID() As String
    BusID = m_szBusID
End Property


'##ModelId=3E82291700EA
Public Property Let BusID(ByVal Value As String)
    m_szBusID = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E822917029F
Public Property Get TransportCompanyID() As String
    TransportCompanyID = m_szTransportCompanyID
End Property


'##ModelId=3E82291700EA
Public Property Let TransportCompanyID(ByVal Value As String)
    m_szTransportCompanyID = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E822917029F
Public Property Get TransportCompanyName() As String
    TransportCompanyName = m_szTransportCompanyName
End Property


'##ModelId=3E82291700EA
Public Property Let TransportCompanyName(ByVal Value As String)
    m_szTransportCompanyName = Value
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


'**************************************************
'Member Code:S3
'Brief Description:测试对象是否有效
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then RaiseError ERR_VehicleSettlePriceAvailable
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
Public Sub DeleteRoute(BusID As String, TransportCompanyID As String)
    On Error GoTo DeleteObjectErr
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "DELETE FROM fin_bus_settle_price_lst " _
        & " WHERE bus_id = " & TransFieldValueToString(BusID) & "AND" _
        & " transport_company_id = " & TransFieldValueToString(TransportCompanyID) _

    oDB.Execute (szSql)
    Exit Sub
DeleteObjectErr:
    Call RaiseError(ERR_VehicleSettlePriceNotExist)
End Sub



