VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Package"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E039C23002B"
'行包受理单
Option Base 0

Option Explicit



'
'##ModelId=3E03B0D702F3
Private m_lPackageID As Long


'货物名称
'##ModelId=3E03B10C0000
Private m_szPackageName As String
Private m_szPackType As String

'地区类型
Private m_szAreaType As String

'到达时间
Private m_dtArrivedTime As Date

'计重
'##ModelId=3E03B29D0314
Private m_szCalWeight As String

'起点站名称
'##ModelId=3E03B14F01DD
Private m_szStartStationName As String

Private m_szLicenseTagNo As String

'件数
'##ModelId=3E03B3120286
Private m_nNumber As Integer

'存放位置
Private m_szSavePosition As String

'受理单状态
'##ModelId=3E03B37B01DD
Private m_nStatus As Integer

'受理单状态
'##ModelId=3E03B3940251
Private m_szStatusString As String

'托运人
'##ModelId=3E03B3A40268
Private m_szShipper As String
Private m_szShipperUnit As String
Private m_szShipperPhone As String

'交货方式
'##ModelId=3E03B3E300B0
Private m_szPickType As String

'收件人联系电话
'##ModelId=3E03B40301CE
Private m_szPicker As String
Private m_szPickerPhone As String
Private m_szPickerUnit As String
'收件人联系地址
'##ModelId=3E03B5870302
Private m_szPickerAddress As String
Private m_szPickerCreditID As String
Private m_dtPickTime As Date
'提货受理人
Private m_szPickOperator As String

'受理人姓名
'##ModelId=3E03B5FD019A
Private m_szOperator As String

'受理时间
'##ModelId=3E03B5C40314
Private m_dtOperateTime As Date
Private m_szUserID As String

Private m_szRemark As String '备注

'以下是结算时的信息
'受理单号
Private m_szSheetID As String
Private m_dbTransitCharge As Double
Private m_dbLoadCharge As Double
Private m_dbMoveCharge As Double
Private m_dbSendCharge As Double
Private m_dbKeepCharge As Double
Private m_dbOtherCharge As Double
Private m_szLoader As String
Private m_szDrawerPhone As String
Private m_szDrawer As String
Private m_szSellStationID As String

'活动用户
Private m_oActiveUser As ActiveUser
Private m_nObjectStatus As EObjectStatus



'本模块私有定义部分
'==========================================================================
'错误号
Public Enum EErrPackage
    ERR_AcceptPackageWithout = ERR_Package + 21   '没有些受理单
    ERR_AcceptPackageCanced = ERR_Package + 22      '受理单已作废
    ERR_PackageSheetExist = ERR_Package + 23           '行包提取单据已存在
End Enum
'权限号

Public Enum ERightPackage
    RIGHT_AcceptPackage = ERR_Package + cnMidRightBegin             '30100开始
    RIGHT_AcceptPackageEdit = RIGHT_AcceptPackage + cnMidRightStep * 1    '更改行包单
    RIGHT_AcceptPackageDelete = RIGHT_AcceptPackage + cnMidRightStep * 2    '删除行包单
End Enum

'初始化活动对象
'##ModelId=3E03A0940076
'**************************************************
'Member Code:S1
'Brief Description:初始化对象
'Engineer:
'Date Generated:2002/12/21
'Last Revision Date:
'**************************************************
Public Sub Init(poAUser As ActiveUser)
    On Error GoTo InitErr
    Dim oParam As New SystemParam
    
    
    Set m_oActiveUser = poAUser
    m_szUserID = poAUser.UserID
    
    Exit Sub
InitErr:
    Call RaiseError(err.Number, "AcceptSheet:Init Method", err.Description)
End Sub

'指定受理单
'##ModelId=3E03A11D0023
Public Sub Identify(PackageID As Long)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_Package
    
    On Error GoTo IdentifyErr

       m_lPackageID = PackageID
       RefreshMemoryInfo

    Exit Sub
IdentifyErr:
    Call RaiseError(err.Number, "AcceptSheet:Identify Method", err.Description)
End Sub

Private Sub RefreshMemoryInfo() '按照对象现在的ID将数据库中相应的信息读至对象内
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim i As Integer, j As Integer
    
    
    oDb.ConnectionString = GetConnectionStr(cszPackage)
    szSql = "SELECT * " _
        & " FROM package_arrived_info  " _
        & " WHERE package_id=" & m_lPackageID
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_AcceptPackageWithout
    
    m_szSheetID = FormatDbValue(rsTemp!sheet_id)
    m_szPackageName = FormatDbValue(rsTemp!package_name)
    m_szAreaType = FormatDbValue(rsTemp!area_type)
    m_szPackType = FormatDbValue(rsTemp!Pack)
    m_dtArrivedTime = FormatDbValue(rsTemp!arrive_time)
    m_szCalWeight = FormatDbValue(rsTemp!Weight)
    m_szStartStationName = FormatDbValue(rsTemp!start_station_name)
    m_szLicenseTagNo = FormatDbValue(rsTemp!license_tag_no)
    m_nNumber = FormatDbValue(rsTemp!package_number)
    m_szSavePosition = FormatDbValue(rsTemp!save_position)
    m_nStatus = FormatDbValue(rsTemp!Status)
    Select Case m_nStatus
        Case 0
            m_szStatusString = "未提"
        Case 1
            m_szStatusString = "已结"
        Case 2
            m_szStatusString = "已废"
    End Select
    m_szShipper = FormatDbValue(rsTemp!send_name)
    m_szShipperUnit = FormatDbValue(rsTemp!send_unit)
    m_szShipperPhone = FormatDbValue(rsTemp!send_phone)
    m_szPickType = FormatDbValue(rsTemp!pick_type)
    m_szPicker = FormatDbValue(rsTemp!Picker)
    m_szPickerAddress = FormatDbValue(rsTemp!picker_address)
    m_szPickerPhone = FormatDbValue(rsTemp!picker_phone)
    m_szPickerUnit = FormatDbValue(rsTemp!picker_unit)
    m_szPickerCreditID = FormatDbValue(rsTemp!picker_credit)
    m_dtPickTime = FormatDbValue(rsTemp!pick_time)
    m_szOperator = FormatDbValue(rsTemp!Operator)
    m_szPickOperator = FormatDbValue(rsTemp!user_id)
    m_dtOperateTime = FormatDbValue(rsTemp!operation_time)
    m_szLoader = FormatDbValue(rsTemp!Loader)
        
    m_dbTransitCharge = FormatDbValue(rsTemp!transit_charge)
    m_dbLoadCharge = FormatDbValue(rsTemp!load_charge)
    m_dbKeepCharge = FormatDbValue(rsTemp!keep_charge)
    m_dbMoveCharge = FormatDbValue(rsTemp!move_charge)
    m_dbSendCharge = FormatDbValue(rsTemp!send_charge)
    m_dbOtherCharge = FormatDbValue(rsTemp!other_charge)
    m_szDrawer = FormatDbValue(rsTemp!Drawer)
    m_szDrawerPhone = FormatDbValue(rsTemp!drawer_phone)
    m_dbOtherCharge = FormatDbValue(rsTemp!other_charge)
    m_szRemark = FormatDbValue(rsTemp!Memo)
    m_szSellStationID = FormatDbValue(rsTemp!sell_station_id)
    
    
    
    Set rsTemp = Nothing
    
End Sub

'新增受理单
'##ModelId=3E03A16D00E6
Public Sub AddNew()
    '验证活动用户及权限
'    AssertActiveUserValid m_oActiveUser, ERR_LuggageAcceptSheet
    
    On Error GoTo AddNewErr
    m_nObjectStatus = ST_AddObj
    m_nStatus = 0
    '把属性值清空
    m_szSheetID = ""
    m_szPackageName = ""
    m_szAreaType = ""
    m_szPackType = ""
    m_dtArrivedTime = cszEmptyDateStr
    m_szCalWeight = ""
    m_szStartStationName = ""
    m_szLicenseTagNo = ""
    m_nNumber = 0
    m_szSavePosition = ""
    m_nStatus = 0
    m_szShipper = ""
    m_szShipperUnit = ""
    m_szShipperPhone = ""
    m_szPickType = ""
    m_szPicker = ""
    m_szPickerAddress = ""
    m_szPickerPhone = ""
    m_szPickerUnit = ""
    m_szPickerCreditID = ""
    m_dtPickTime = cszEmptyDateStr
    m_szOperator = ""
    m_dtOperateTime = cszEmptyDateStr
    m_szLoader = ""
        
    m_dbTransitCharge = 0
    m_dbLoadCharge = 0
    m_dbKeepCharge = 0
    m_dbMoveCharge = 0
    m_dbSendCharge = 0
    m_dbOtherCharge = 0
    m_szRemark = ""
    m_szSellStationID = m_oActiveUser.SellStationID
    
    
    Exit Sub
AddNewErr:
    Call RaiseError(err.Number, "AcceptSheet:AddNew Method", err.Description)
End Sub



'更新受理单
'##ModelId=3E03A179024C
Public Sub Update(Optional m_nObjectStatus As EObjectStatus)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_Package
    
    On Error GoTo UpdateErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szSqlField As String
    
    Select Case m_nObjectStatus
    Case ST_AddObj
'        AddObject
    Case ST_EditObj
        AssertHaveRight m_oActiveUser, RIGHT_AcceptPackageEdit
        
        UpdateToDB
    Case ST_NormalObj
    Case Else
        ShowError ERR_NotAvailable
    End Select

    Exit Sub
UpdateErr:
    Call RaiseError(err.Number, "AcceptSheet:Update Method", err.Description)
End Sub
Private Function AddObject()
'    Dim szSql As String
'    Dim oDb As New RTConnection
'    Dim rsTemp As Recordset
'    Dim szSqlField As String
'    On Error GoTo here
'    oDb.ConnectionString = GetConnectionStr(cszPackage)
'    szSql = "SELECT ISNULL(MAX(package_id),0) package_id FROM package_arrived_info"
'    Set rsTemp = oDb.Execute(szSql)
'    Dim lTmpID As Long
'    lTmpID = FormatDbValue(rsTemp!package_id) + 1
'
'    szSql = "INSERT package_arrived_info (package_id,sheet_id,package_name,area_type,pack,arrive_time,weight,start_station_name,license_tag_no,package_number,save_position," _
'                                        & "status,send_name,send_unit,send_phone,pick_type,picker,picker_unit,picker_address,picker_phone,picker_credit,pick_time,user_id," _
'                                        & "operation_time,loader,transit_charge,load_charge,keep_charge,move_charge,send_charge,other_charge,remark) " & _
'            " VALUES(" & TransFieldValueToString(lTmpID) & _
'            "," & TransFieldValueToString(m_szSheetID) & _
'            "," & TransFieldValueToString(m_szPackageName) & _
'            "," & TransFieldValueToString(m_szAreaType) & _
'            "," & TransFieldValueToString(m_szPackType) & _
'            "," & TransFieldValueToString(m_dtArrivedTime) & _
'            "," & TransFieldValueToString(m_szCalWeight) & _
'            "," & TransFieldValueToString(m_szStartStationName) & _
'            "," & TransFieldValueToString(m_szLicenseTagNo) & _
'            "," & TransFieldValueToString(m_nNumber) & _
'            "," & TransFieldValueToString(m_szSavePosition) & _
'            "," & TransFieldValueToString(m_nStatus) & _
'            "," & TransFieldValueToString(m_szShipper) & _
'            "," & TransFieldValueToString(m_szShipperUnit) & _
'            "," & TransFieldValueToString(m_szShipperPhone) & _
'            "," & TransFieldValueToString(m_szPickType) & _
'            "," & TransFieldValueToString(m_szPicker) & _
'            "," & TransFieldValueToString(m_szPickerUnit) & _
'            "," & TransFieldValueToString(m_szPickerAddress)
'        szSql = szSql & _
'            "," & TransFieldValueToString(m_szPickerPhone) & _
'            "," & TransFieldValueToString(m_szPickerCreditID) & _
'            "," & TransFieldValueToString(m_dtPickTime) & _
'            "," & TransFieldValueToString(m_szOperator) & _
'            "," & TransFieldValueToString(Format(Now, "yyyy-MM-dd hh:mm")) & _
'            "," & TransFieldValueToString(m_szLoader) & _
'            "," & TransFieldValueToString(m_dbTransitCharge) & _
'            "," & TransFieldValueToString(m_dbLoadCharge) & _
'            "," & TransFieldValueToString(m_dbKeepCharge) & _
'            "," & TransFieldValueToString(m_dbMoveCharge) & _
'            "," & TransFieldValueToString(m_dbSendCharge) & _
'            "," & TransFieldValueToString(m_dbOtherCharge) & _
'            "," & TransFieldValueToString(m_szRemark)
'
'    oDb.Execute szSql
'    m_lPackageID = lTmpID
'    Set oDb = Nothing
'    Exit Function
'here:
'    Call RaiseError(err.Number, "AcceptSheet:GetLugItemDetail Method", err.Description)
End Function

Private Sub UpdateToDB()
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSqlField As String
    On Error GoTo here
    oDb.ConnectionString = GetConnectionStr(cszPackage)
    szSql = "SELECT * FROM package_arrived_info WHERE package_id =" & TransFieldValueToString(m_lPackageID)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_AcceptPackageWithout
    
    oDb.BeginTrans
    
    szSql = "UPDATE package_arrived_info SET " _
        & "sheet_id='" & m_szSheetID & "'," _
        & "package_name='" & m_szPackageName & "'," _
        & "area_type='" & m_szAreaType & "'," _
        & "pack='" & m_szPickType & "'," _
        & "arrive_time=" & TransFieldValueToString(m_dtArrivedTime) & "," _
        & "weight='" & m_szCalWeight & "'," _
        & "start_station_name='" & m_szStartStationName & "'," _
        & "license_tag_no='" & m_szLicenseTagNo & "'," _
        & "" _
        & "package_number=" & m_nNumber & "," _
        & "save_position='" & m_szSavePosition & "'," _
        & "status=" & m_nStatus & "," _
        & "send_name='" & m_szShipper & "'," _
        & "send_unit='" & m_szShipperUnit & "'," _
        & "send_phone='" & m_szShipperPhone & "'," _
        & "pick_type='" & m_szPickType & "'," _
        & "picker='" & m_szPicker & "'," _
        & "picker_unit='" & m_szPickerUnit & "'," _
        & "picker_address='" & m_szPickerAddress & "'," _
        & "picker_phone='" & m_szPickerPhone & "'," _
        & "picker_credit='" & m_szPickerCreditID & "'," _
        & "pick_time=" & TransFieldValueToString(m_dtPickTime) & "," _
        & "memo=" & TransFieldValueToString(m_szRemark) & "," _
        & "operator='" & m_szOperator & "',"

    szSql = szSql & " user_id = user_id " & ", " _
        & " operation_time = " & TransFieldValueToString(Now) & " , " _
        & " loader = " & TransFieldValueToString(m_szLoader) & " , " _
        & " load_charge = " & TransFieldValueToString(m_dbLoadCharge) & " , " _
        & " keep_charge = " & TransFieldValueToString(m_dbKeepCharge) & " , " _
        & " move_charge = " & TransFieldValueToString(m_dbMoveCharge) & " , " _
        & " send_charge= " & TransFieldValueToString(m_dbSendCharge) & ", " _
        & " other_charge= " & TransFieldValueToString(m_dbOtherCharge) & ", " _
        & " drawer= " & TransFieldValueToString(m_szDrawer) & ", " _
        & " drawer_phone= " & TransFieldValueToString(m_szDrawerPhone) & ", " _
        & " transit_charge = " & TransFieldValueToString(m_dbTransitCharge) _
        & " WHERE package_id =" & TransFieldValueToString(m_lPackageID)
        
    oDb.Execute szSql
    
    oDb.CommitTrans
    
    Set oDb = Nothing
    Exit Sub
here:
    oDb.RollbackTrans
    Call RaiseError(err.Number, "AcceptSheet:GetLugItemDetail Method", err.Description)
End Sub


Private Sub DeleteObject()
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszPackage)
    oDb.Execute ("DELETE package_arrived_info WHERE package_id='" & m_lPackageID & "'")
    Set oDb = Nothing
End Sub
'删除受理单
'##ModelId=3E03A52A0345
Public Sub Delete()
    '验证活动用户及权限
'    AssertActiveUserValid m_oActiveUser, ERR_LuggageAcceptSheet
'    AssertHaveRight m_oActiveUser, RIGHT_AcceptSheetDelete
    
    On Error GoTo DeleteErr
    
    DeleteObject
    m_nObjectStatus = ST_NotAvailable

    Exit Sub
DeleteErr:
    Call RaiseError(err.Number, "AcceptSheet:Delete Method", err.Description)
End Sub
'##ModelId=3E04A7B700DD
Public Property Get PackType() As String
   Let PackType = m_szPackType
End Property

'##ModelId=3E04A7B6021C
Public Property Let PackType(ByVal Value As String)
    m_szPackType = Value
End Property '##ModelId=3E04A7B700DD
Public Property Get PackageID() As Long
   Let PackageID = m_lPackageID
End Property
'##ModelId=3E04A7B6021C
Public Property Let PackageID(ByVal Value As Long)
    m_lPackageID = Value
End Property '##ModelId=3E04A7B700DD
Public Property Get SheetID() As String
   Let SheetID = m_szSheetID
End Property

'##ModelId=3E04A7B6021C
Public Property Let SheetID(ByVal Value As String)
    m_szSheetID = Value
End Property
'##ModelId=3E04A7B700DD
Public Property Get PackageName() As String
   Let PackageName = m_szPackageName
End Property

'##ModelId=3E04A7B6021C
Public Property Let PackageName(ByVal Value As String)
    m_szPackageName = Value
End Property
'##ModelId=3E04A7B700DD
Public Property Get AreaType() As String
   Let AreaType = m_szAreaType
End Property

'##ModelId=3E04A7B6021C
Public Property Let AreaType(ByVal Value As String)
    m_szAreaType = Value
End Property

'##ModelId=3E04A7B700DD
Public Property Get ArrivedTime() As Date
   Let ArrivedTime = m_dtArrivedTime
End Property

'##ModelId=3E04A7B6021C
Public Property Let ArrivedTime(ByVal Value As Date)
    m_dtArrivedTime = Value
End Property
'##ModelId=3E04A7B700DD
Public Property Get PickTime() As Date
   Let PickTime = m_dtPickTime
End Property

'##ModelId=3E04A7B6021C
Public Property Let PickTime(ByVal Value As Date)
    m_dtPickTime = Value
End Property

'##ModelId=3E04A7B700DD
Public Property Get PickOperator() As String
   Let PickOperator = m_szPickOperator
End Property

'##ModelId=3E04A7B6021C
Public Property Let PickOperator(ByVal Value As String)
    m_szPickOperator = Value
End Property

'##ModelId=3E04A7AC022C
Public Property Get CalWeight() As String
   Let CalWeight = m_szCalWeight
End Property

'##ModelId=3E04A7AB03E3
Public Property Let CalWeight(ByVal Value As String)
m_szCalWeight = Value
End Property
'##ModelId=3E04A7A80280
Public Property Get StartStationName() As String

    
StartStationName = m_szStartStationName
End Property

'##ModelId=3E04A7AB03E3
Public Property Let StartStationName(ByVal Value As String)
m_szStartStationName = Value
End Property

'车牌
Public Property Let LicenseTagNo(ByVal Value As String)
    m_szLicenseTagNo = Value
End Property

Public Property Get LicenseTagNo() As String
    LicenseTagNo = m_szLicenseTagNo
End Property
'##ModelId=3E04A7B00105
Public Property Get PackageNumber() As Integer
   Let PackageNumber = m_nNumber
End Property

'##ModelId=3E04A7AF0294
Public Property Let PackageNumber(ByVal Value As Integer)
m_nNumber = Value
End Property
'##ModelId=3E04A7B00105
Public Property Get SavePosition() As String
   Let SavePosition = m_szSavePosition
End Property
'##ModelId=3E04A7B00105
Public Property Get UserID() As String
   Let UserID = m_szUserID
End Property
'##ModelId=3E04A7AF0294
Public Property Let SavePosition(ByVal Value As String)
m_szSavePosition = Value
End Property
'##ModelId=3E04A7B103BA
Public Property Get StatusString() As String

    Select Case m_nStatus
        Case 0
            m_szStatusString = "未提取"
        Case 1
            m_szStatusString = "已提"
        Case 2
            m_szStatusString = "已废"
    End Select
        
StatusString = m_szStatusString
End Property

'##ModelId=3E04A7B1025B
Public Property Get Status() As Integer

        
Status = m_nStatus
End Property
Public Property Let Status(ByVal Value As Integer)
m_nStatus = Value
End Property
'##ModelId=3E04A7B203C5
Public Property Get Shipper() As String
Shipper = m_szShipper
End Property

'##ModelId=3E04A7B20144
Public Property Let Shipper(ByVal Value As String)
m_szShipper = Value
End Property
'##ModelId=3E04A7B203C5
Public Property Get ShipperUnit() As String
ShipperUnit = m_szShipperUnit
End Property

'##ModelId=3E04A7B20144
Public Property Let ShipperUnit(ByVal Value As String)
m_szShipperUnit = Value
End Property
'##ModelId=3E04A7B203C5
Public Property Get ShipperPhone() As String
ShipperPhone = m_szShipperPhone
End Property


'##ModelId=3E04A7B20144
Public Property Let ShipperPhone(ByVal Value As String)
m_szShipperPhone = Value
End Property
'##ModelId=3E04A7B5003A
Public Property Get PickType() As String
   Let PickType = m_szPickType
End Property

'##ModelId=3E04A7B4018D
Public Property Let PickType(ByVal Value As String)
m_szPickType = Value
End Property
'##ModelId=3E04A7B700DD
Public Property Get PickerAddress() As String
   Let PickerAddress = m_szPickerAddress
End Property

'##ModelId=3E04A7B6021C
Public Property Let PickerAddress(ByVal Value As String)
m_szPickerAddress = Value
End Property

'##ModelId=3E04A7B60082
Public Property Get PickerPhone() As String
   Let PickerPhone = m_szPickerPhone
End Property

'##ModelId=3E04A7B501CB
Public Property Let PickerPhone(ByVal Value As String)
m_szPickerPhone = Value
End Property



'##ModelId=3E04A7B303E5
Public Property Get Picker() As String
Picker = m_szPicker
End Property

'##ModelId=3E04A7B3015A
Public Property Let Picker(ByVal Value As String)
m_szPicker = Value
End Property

'##ModelId=3E04A7B303E5
Public Property Get PickerUnit() As String
PickerUnit = m_szPickerUnit
End Property

'##ModelId=3E04A7B3015A
Public Property Let PickerUnit(ByVal Value As String)
m_szPickerUnit = Value
End Property

'##ModelId=3E04A7B303E5
Public Property Get PickerCreditID() As String
PickerCreditID = m_szPickerCreditID
End Property

'##ModelId=3E04A7B3015A
Public Property Let PickerCreditID(ByVal Value As String)
m_szPickerCreditID = Value
End Property
'##ModelId=3E04A7B801D9
Public Property Get OperateTime() As Date

    
OperateTime = m_dtOperateTime
End Property

'##ModelId=3E04A7B80034
Public Property Get Operator() As String
'    If m_szOperator = "" Then
'        m_szOperator = m_oActiveUser.UserName
'    End If
Operator = m_szOperator
End Property
'##ModelId=3E04A7B303E5
Public Property Let Operator(ByVal Value As String)
m_szOperator = Value
End Property '##ModelId=3E04A7B303E5
Public Property Get Loader() As String
Loader = m_szLoader
End Property

'##ModelId=3E04A7B3015A
Public Property Let Loader(ByVal Value As String)
m_szLoader = Value
End Property



'##ModelId=3E04A7AD01BF
Public Property Get TransitCharge() As Double
   Let TransitCharge = m_dbTransitCharge
End Property

'##ModelId=3E04A7AC0376
Public Property Let TransitCharge(ByVal Value As Double)
m_dbTransitCharge = Value
End Property

'##ModelId=3E04A7AD01BF
Public Property Get LoadCharge() As Double
   Let LoadCharge = m_dbLoadCharge
End Property

'##ModelId=3E04A7AC0376
Public Property Let LoadCharge(ByVal Value As Double)
m_dbLoadCharge = Value
End Property
'##ModelId=3E04A7AD01BF
Public Property Get KeepCharge() As Double
   Let KeepCharge = m_dbKeepCharge
End Property

'##ModelId=3E04A7AC0376
Public Property Let KeepCharge(ByVal Value As Double)
m_dbKeepCharge = Value
End Property
'##ModelId=3E04A7AD01BF
Public Property Get MoveCharge() As Double
   Let MoveCharge = m_dbMoveCharge
End Property

'##ModelId=3E04A7AC0376
Public Property Let MoveCharge(ByVal Value As Double)
m_dbMoveCharge = Value
End Property

'##ModelId=3E04A7AD01BF
Public Property Get SendCharge() As Double
   Let SendCharge = m_dbSendCharge
End Property

'##ModelId=3E04A7AC0376
Public Property Let SendCharge(ByVal Value As Double)
m_dbSendCharge = Value
End Property
'##ModelId=3E04A7AD01BF
Public Property Get OtherCharge() As Double
   Let OtherCharge = m_dbOtherCharge
End Property

'##ModelId=3E04A7AC0376
Public Property Let OtherCharge(ByVal Value As Double)
m_dbOtherCharge = Value
End Property

'备注1
Public Property Let Remark(ByVal Value As String)
    m_szRemark = Value
End Property

Public Property Get Remark() As String
    Remark = m_szRemark
End Property

'提件人电话(嵊州新加的)
Public Property Let DrawerPhone(ByVal Value As String)
    m_szDrawerPhone = Value
End Property

Public Property Get DrawerPhone() As String
    DrawerPhone = m_szDrawerPhone
End Property

'提件人姓名(嵊州新加的)
Public Property Let Drawer(ByVal Value As String)
    m_szDrawer = Value
End Property

Public Property Get Drawer() As String
    Drawer = m_szDrawer
End Property

'受理站(楚门新加的) 只读
Public Property Get SellStationID() As String
    SellStationID = m_szSellStationID
End Property


