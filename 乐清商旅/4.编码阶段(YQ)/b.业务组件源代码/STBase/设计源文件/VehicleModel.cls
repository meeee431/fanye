VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VehicleModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**********************************************************
'* Source File Name:VehicleModel.cls
'* Project Name:STBase.vbp
'* Engineer:
'* Data Generated:
'* Last Revision Date:
'* Brief Description:公路等级对象
'* Relational Document:
'**********************************************************

Option Explicit

'出错
Public Enum EErrVehicleModel
    ERR_VehicleModelExist = ERR_VehicleModel + ERR_AddDuplicate '58001车型已存在不能新增
    ERR_VehicleModelNoDelete = ERR_VehicleModel + ERR_DeleteChildExist '58007车型已车次，不能删除
    ERR_VehicleModelNotAvailable = ERR_VehicleModel + ERR_NotAvailable '58012车型对象处于无状态
    ERR_VehicleModelNotAddStatus = ERR_VehicleModel + ERR_NotAddObj '58016车型对象不在新增状态
    ERR_VehicleModelNotExist = ERR_VehicleModel + 21 '58021车型无存在
    ERR_VehicleModelIDNotNull = ERR_VehicleModel + 22 '58022车型代码不能为空
End Enum

'权限
Public Enum ERightVehicleModel
    RIGHT_VehicleModelManagement = ERR_VehicleModel + cnMidRightBegin + cnMidRightStep * 1 '58106车型管理
End Enum

Private m_oActiveUser As ActiveUser 'P1所有的类都有的私有类变量
Private m_nObjectStatus As EObjectStatus 'P2所有的实体类都有的私有类变量
Private m_szVehicleModelCode As String 'P3车型代码
Private m_szVehicleModelName As String 'P4车型名称
Private m_szVehicleModelShortName As String 'P5车型简称
Private m_nStartSeatNumber As Integer 'P6开始座位号
Private m_nSeatCount As Integer 'P7座位数
Private m_nLugRatio As Double  'P8行包费率


'F1获得该车型的所有车辆,通过查型表transport_vehicle_info
Public Function GetAllVehicle() As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute("SELECT * FROM Vehicle_info WHERE vehicle_type_code='" & m_szVehicleModelCode & "'")
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    ReDim szaTemp(1 To nCount, 1 To 2) As String
    For i = 1 To nCount
        szaTemp(i, 1) = FormatDbValue(rsTemp!vehicle_id)
        szaTemp(i, 2) = FormatDbValue(rsTemp!license_tag_no)
        rsTemp.MoveNext
    Next
    GetAllVehicle = szaTemp
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function
'P1
Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
End Property
'P1
Public Property Set SelfUser(vNewValue As ActiveUser)
    If m_nObjectStatus = ST_AddObj Then
        ShowError ERR_AddObj
    ElseIf m_nObjectStatus = ST_EditObj Then
        ShowError ERR_EditObj
    ElseIf m_nObjectStatus = ST_NormalObj Then
        ShowError ERR_NormalObj
    Else
        Set m_oActiveUser = vNewValue
    End If
End Property
'P2
Public Property Get ObjStatus() As EObjectStatus
    ObjStatus = m_nObjectStatus
End Property
'P3
Public Property Let VehicleModelCode(ByVal vData As String)
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    AssertActiveUserValid m_oActiveUser, ERR_VehicleModel
    '查询是否是新增状态
    If m_nObjectStatus <> ST_AddObj Then ShowError ERR_VehicleModelNotAddStatus
    '输入的主键是否为空
    If vData = "" Then ShowError ERR_VehicleModelIDNotNull
    '查询该对象是否已有记录
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT vehicle_type_code FROM vehicle_type_code WHERE vehicle_type_code='" & RTrim(vData) & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 1 Then ShowError ERR_VehicleModelExist
    m_szVehicleModelCode = vData
    Set rsTemp = Nothing
    Set oDb = Nothing
End Property
'P3
Public Property Get VehicleModelCode() As String
   VehicleModelCode = m_szVehicleModelCode
End Property
'P4
Public Property Let VehicleModelName(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szVehicleModelName = vData
End Property
'P4
Public Property Get VehicleModelName() As String
    VehicleModelName = m_szVehicleModelName
End Property
'P5
Public Property Let VehicleModelShortName(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szVehicleModelShortName = vData
End Property
'P5
Public Property Get VehicleModelShortName() As String
    VehicleModelShortName = m_szVehicleModelShortName
End Property
'P6
Public Property Let StartSeatNumber(ByVal vData As Integer)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_nStartSeatNumber = vData
End Property
'P6
Public Property Get StartSeatNumber() As Integer
    StartSeatNumber = m_nStartSeatNumber
End Property
'P7
Public Property Let SeatCount(ByVal vData As Integer)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_nSeatCount = vData
End Property
'P7
Public Property Get SeatCount() As Integer
    SeatCount = m_nSeatCount
End Property

'P8
Public Property Let LugRatio(ByVal vData As Double)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_nLugRatio = vData
End Property
'P8
Public Property Get LugRatio() As Double
    LugRatio = m_nLugRatio
End Property

'S1
Public Sub Init(poAUser As ActiveUser)
    Set m_oActiveUser = poAUser
    'AssertHaveRight m_oActiveUser, RIGHT_VehicleModelManagement
End Sub

'S2类刚建立时处于不可用状态,此时类不能使用
Private Sub Class_Initialize()
    m_nObjectStatus = ST_NotAvailable
End Sub

'S3测试对象的状态是否有效（无效状态[ST_NotAvailable]和新增状态[ST_AddObj]则出错）
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then ShowError ERR_VehicleModelNotAvailable
End Sub
'S4
Private Sub AssertObjIsValid() '测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
    AssertActiveUserValid m_oActiveUser, ERR_VehicleModel
    AssertStatusAvailable
End Sub
'S5
Public Sub Identify(VehicleModel As String)
    AssertActiveUserValid m_oActiveUser, ERR_VehicleModel
    m_szVehicleModelCode = RTrim(VehicleModel)
    RefreshMemoryInfo
    m_nObjectStatus = ST_NormalObj
End Sub

'S6调用新增方法后，对象处于新增状态。
Public Sub AddNew()
    AssertHaveRight m_oActiveUser, RIGHT_VehicleModelManagement
    AssertActiveUserValid m_oActiveUser, ERR_VehicleModel
    m_nObjectStatus = ST_AddObj
End Sub
'S7
Public Sub Delete() '调用删除方法后，对象处于不可用状态
    AssertHaveRight m_oActiveUser, RIGHT_VehicleModelManagement
    AssertObjIsValid
    DeleteObject
    m_nObjectStatus = ST_NotAvailable
End Sub
'S8
Public Sub Update()
    AssertHaveRight m_oActiveUser, RIGHT_VehicleModelManagement
    AssertObjIsValid
    Select Case m_nObjectStatus
           Case ST_AddObj
           AddObject
           Case ST_EditObj
           UpdateToDB
           Case ST_NormalObj
           Case Else
           ShowError ERR_NotAvailable
    End Select
    m_nObjectStatus = ST_NormalObj
End Sub

'S9
Private Sub RefreshMemoryInfo() '按照对象现在的ID将数据库中相应的信息读至对象内
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute("SELECT * FROM vehicle_type_code WHERE vehicle_type_code='" & m_szVehicleModelCode & "'")
    If rsTemp.RecordCount = 0 Then ShowError ERR_VehicleModelNotExist
    m_szVehicleModelName = FormatDbValue(rsTemp!vehicle_type_name)
    m_szVehicleModelShortName = FormatDbValue(rsTemp!vehicle_type_short_name)
    m_nSeatCount = FormatDbValue(rsTemp!seat_quantity)
    m_nStartSeatNumber = FormatDbValue(rsTemp!start_seat_no)
    '行包费率
    m_nLugRatio = FormatDbValue(rsTemp!lug_base_transit_ratio)
    Set rsTemp = Nothing
    Set oDb = Nothing
End Sub

'S10
Private Sub AddObject()
    Dim oDb As New RTConnection
    Dim szTempSql As String
    Dim szSqlField As String
    Dim szSqlContext As String
    If Trim(m_szVehicleModelCode) = "" Then ShowError ERR_VehicleModelIDNotNull
    
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
On Error GoTo Here
    szTempSql = "INSERT vehicle_type_code ("
    szSqlField = szSqlField & "vehicle_type_code,"
    szSqlField = szSqlField & "vehicle_type_name,"
    szSqlField = szSqlField & "vehicle_type_short_name,"
    szSqlField = szSqlField & "seat_quantity,"
    szSqlField = szSqlField & "start_seat_no,"
    '行包费率
    szSqlField = szSqlField & "lug_base_transit_ratio)"
    
    szSqlContext = " VALUES ('" & m_szVehicleModelCode & "',"
    szSqlContext = szSqlContext & " '" & m_szVehicleModelName & "',"
    szSqlContext = szSqlContext & " '" & m_szVehicleModelShortName & "',"
    szSqlContext = szSqlContext & " " & Str(m_nSeatCount) & ","
    szSqlContext = szSqlContext & " " & Str(m_nStartSeatNumber) & ","
    '行包费率
    szSqlContext = szSqlContext & Str(m_nLugRatio) & ")"
    
    szTempSql = szTempSql & szSqlField & szSqlContext
    oDb.Execute (szTempSql)
    Set oDb = Nothing
Exit Sub
Here:
    AssertAddObjectError ERR_VehicleModel, oDb
    Set oDb = Nothing
End Sub
'S11
Private Sub DeleteObject()
    Dim oDb As New RTConnection
On Error GoTo Here
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    oDb.Execute ("DELETE vehicle_type_code WHERE vehicle_type_code='" & m_szVehicleModelCode & "'")
    Set oDb = Nothing
Exit Sub
Here:
    AssertDeleteObjectError ERR_VehicleModel, oDb
    Set oDb = Nothing
End Sub

'S12
Private Sub UpdateToDB()
    Dim oDb As New RTConnection
    Dim szTempSql As String
    Dim szSqlField As String
    Dim szSqlWhere As String
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
On Error GoTo Here
    szTempSql = "UPDATE vehicle_type_code SET "
    szSqlField = szSqlField & "vehicle_type_name='" & m_szVehicleModelName & "',"
    szSqlField = szSqlField & "vehicle_type_short_name='" & m_szVehicleModelShortName & "',"
    szSqlField = szSqlField & "seat_quantity=" & m_nSeatCount & ","
    szSqlField = szSqlField & "start_seat_no='" & m_nStartSeatNumber & "',"
    '行包费率
    szSqlField = szSqlField & "lug_base_transit_ratio=" & m_nLugRatio
    
    szSqlWhere = " WHERE vehicle_type_code='" & m_szVehicleModelCode & "'"
    szTempSql = szTempSql & szSqlField & szSqlWhere
    oDb.Execute (szTempSql)
    Set oDb = Nothing
Exit Sub
Here:
    AssertUpdateObjectError ERR_VehicleModel, oDb
    Set oDb = Nothing
End Sub




