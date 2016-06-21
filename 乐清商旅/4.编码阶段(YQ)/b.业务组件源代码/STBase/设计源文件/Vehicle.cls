VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"VehicleModel"
Attribute VB_Ext_KEY = "Member1" ,"Owner"
Attribute VB_Ext_KEY = "Member2" ,"Company"

'**********************************************************
'* Source File Name:Company.cls
'* Project Name:STBase.vbp
'* Engineer:
'* Data Generated:
'* Last Revision Date:
'* Brief Description:车辆对象
'* Relational Document:
'**********************************************************
Option Explicit

'出错
Public Enum EErrVehicle
    ERR_VehicleExist = ERR_Vehicle + ERR_AddDuplicate '13201车辆已存在不能新增
    ERR_VehicleNotParent = ERR_Vehicle + ERR_AddNoParent '13202参运公司、车型车主不确定
    ERR_VehicleNoDelete = ERR_Vehicle + ERR_DeleteChildExist '13207车辆已车次或有检票记录
    ERR_VehicleNotEdit = ERR_Vehicle + ERR_EditChildExist '13205车辆参运公司、车型、车主
    ERR_VehicleNotAvailable = ERR_Vehicle + ERR_NotAvailable '13212车辆对象处于无状态
    ERR_VehicleNotAddStatus = ERR_Vehicle + ERR_NotAddObj '13216车辆对象不在新增状态
    ERR_VehicleNotExist = ERR_Vehicle + 21 '13221车辆无存在
    ERR_VehicleIDNotNull = ERR_Vehicle + 22 '13222车辆代码不能为空
    ERR_NOBusInEniroment = ERR_Vehicle + 23 '13223该车辆环境不存在车次
End Enum

'权限
Public Enum ERightVehicle
    RIGHT_VehicleManagement = ERR_Vehicle + cnMidRightBegin + cnMidRightStep * 1 '53306车辆管理
    RIGHT_VehicleStop = ERR_Vehicle + cnMidRightBegin + cnMidRightStep * 2 '53312车辆停班
    RIGHT_VehicleAllBusStop = ERR_Vehicle + cnMidRightBegin + cnMidRightStep * 3 '53318车辆的所有车次停班
End Enum

'保持属性值的局部变量
Private m_oActiveUser As ActiveUser 'P1所有的类都有的私有类变量
Private m_nObjectStatus As EObjectStatus 'P2所有的实体类都有的私有类变量
Private m_szVehicleID As String 'P3车辆代码
Private m_szVehicleModel As String 'P4车型
Private m_szCardID As String 'P5报班卡号
Private m_szLicenseTag As String 'P6车牌
Private m_szOwner As String 'P7车主
Private m_nStartSeatNo As Integer 'P8开始座位号
Private m_nSeatCount As Integer 'P9座位数
Private m_eStatus As EVehicleStatus 'P10车辆状态
Private m_szCompany As String 'P11参运公司
Private m_szAnnotation As String 'P12注释
Private m_szCompanyName As String 'P13参运公司名称
Private m_szOwnerName As String 'P14车主名称
Private m_szVehicleModelName As String 'P15车型名称

Private m_szSplitCompanyID As String 'P16拆帐公司代码
Private m_szSplitCompanyName As String 'P17拆帐公司名称
Private m_szBusID As String


''''''''''''''''''


Private m_oldszAnnotation   As String
Private m_OldszVehicleModel As String 'P4车型
Private m_OldszCardID As String 'P5报班卡号
Private m_OldszLicenseTag As String 'P6车牌
Private m_OldszOwner As String 'P7车主
Private m_OldnStartSeatNo As Integer 'P8开始座位号
Private m_oldnSeatCount As Integer 'P9座位数
Private m_oldszCompany As String 'P11参运公司
Private m_odlszAnnotation As String 'P12注释
Private m_oldszCompanyName As String 'P13参运公司名称
Private m_oldszOwnerName As String 'P14车主名称
Private m_oldszVehicleModelName As String 'P15车型名称
Private m_eOldStatus As EVehicleStatus '旧的车辆状态



Private m_oldszSplitCompanyID As String 'P16拆帐公司代码
Private m_oldszSplitCompanyName As String 'P17拆帐公司名称

'计划车次----写日志用
Public Property Let ProjectBusID(szBusId As String)
 m_szBusID = szBusId
End Property

'F1获得该车辆在指定计划中运行的所有车次
Public Function GetAllBus(PlanID As String) As String()
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim szaBus() As String
    Dim rsTemp As Recordset
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    If PlanID <> "" Then
    Set rsTemp = oDb.Execute("SELECT DISTINCT tbv.bus_id " _
                            & " FROM bus_vehicle_code tbv " _
                            & " WHERE tbv.vehicle_id='" & m_szVehicleID & "'")
    Else
    Set rsTemp = oDb.Execute("SELECT DISTINCT tbv.bus_id " _
                            & " FROM bus_vehicle_code tbv " _
                            & " WHERE tbv.vehicle_id='" & m_szVehicleID & "'")
    End If
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    ReDim szaBus(1 To nCount) As String
    For i = 1 To nCount
        szaBus(i) = FormatDbValue(rsTemp!Bus_ID)
        rsTemp.MoveNext
    Next
    GetAllBus = szaBus
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
Public Property Let VehicleID(ByVal vData As String)
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    AssertActiveUserValid m_oActiveUser, ERR_Vehicle
    '查询是否是新增状态
    If m_nObjectStatus <> ST_AddObj Then ShowError ERR_VehicleNotAddStatus
    '输入的主键是否为空
    If vData = "" Then ShowError ERR_VehicleIDNotNull
    '查询该对象是否已有记录
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT vehicle_id FROM Vehicle_info WHERE vehicle_id='" & RTrim(vData) & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 1 Then ShowError ERR_VehicleExist
    m_szVehicleID = Trim(vData)
    Set rsTemp = Nothing
    Set oDb = Nothing
End Property
'P3
Public Property Get VehicleID() As String
    VehicleID = m_szVehicleID
End Property
'P4
Public Property Let VehicleModel(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szVehicleModel = vData
End Property
'P4
Public Property Get VehicleModel() As String
    VehicleModel = m_szVehicleModel
End Property
'P5
Public Property Let CardID(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szCardID = vData
End Property
'P5
Public Property Get CardID() As String
    CardID = m_szCardID
End Property
'P6
Public Property Let LicenseTag(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szLicenseTag = vData
End Property
'P6
Public Property Get LicenseTag() As String
    LicenseTag = m_szLicenseTag
End Property
'P7
Public Property Let Owner(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szOwner = vData
End Property
'P7
Public Property Get Owner() As String
    Owner = m_szOwner
End Property
'P8
Public Property Let StartSeatNo(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_nStartSeatNo = vData
End Property
'P8
Public Property Get StartSeatNo() As String
    StartSeatNo = m_nStartSeatNo
End Property
'P9
Public Property Let SeatCount(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_nSeatCount = vData
End Property
'P9
Public Property Get SeatCount() As String
    SeatCount = m_nSeatCount
End Property
'P10
Public Property Let Status(ByVal vData As EVehicleStatus)
    AssertHaveRight m_oActiveUser, RIGHT_VehicleStop
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_eStatus = vData
End Property
'P10
Public Property Get Status() As EVehicleStatus
    Status = m_eStatus
End Property
'P11
Public Property Let Company(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szCompany = vData
End Property
'P11
Public Property Get Company() As String
    Company = m_szCompany
End Property


'P12
Public Property Let Annotation(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szAnnotation = vData
End Property
'P12
Public Property Get Annotation() As String
    Annotation = m_szAnnotation
End Property

'P13参运公司名称
Public Property Get CompanyName() As String
    CompanyName = m_szCompanyName
End Property

'P14车主名称
Public Property Get OwnerName() As String
    OwnerName = m_szOwnerName
End Property
'P15车型名称
Public Property Get VehicleModelName() As Variant
       VehicleModelName = m_szVehicleModelName
End Property

'P16
Public Property Let SplitCompanyID(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szSplitCompanyID = vData
End Property

'P16
Public Property Get SplitCompanyID() As String
    SplitCompanyID = m_szSplitCompanyID
End Property

'P17
Public Property Get SplitCompanyName() As String
    SplitCompanyName = m_szSplitCompanyName
End Property




'S1
Public Sub Init(poAUser As ActiveUser)
    Set m_oActiveUser = poAUser
    'AssertHaveRight m_oActiveUser, RIGHT_VehicleManagement
End Sub

'S2类刚建立时处于不可用状态,此时类不能使用
Private Sub Class_Initialize()
    m_nObjectStatus = ST_NotAvailable
End Sub

'S3测试对象的状态是否有效（无效状态[ST_NotAvailable]和新增状态[ST_AddObj]则出错）
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then ShowError ERR_VehicleNotAvailable
End Sub
'S4
Private Sub AssertObjIsValid() '测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
    AssertActiveUserValid m_oActiveUser, ERR_Vehicle
    AssertStatusAvailable
End Sub
'S5
Public Sub Identify(pszVehicleID As String)
    AssertActiveUserValid m_oActiveUser, ERR_Vehicle
    m_szVehicleID = RTrim(pszVehicleID)
    RefreshMemoryInfo
    m_nObjectStatus = ST_NormalObj
End Sub

'S6调用新增方法后，对象处于新增状态。
Public Sub AddNew()
    AssertHaveRight m_oActiveUser, RIGHT_VehicleManagement
    AssertActiveUserValid m_oActiveUser, ERR_Vehicle
    m_nObjectStatus = ST_AddObj
End Sub
'S7
Public Sub Delete() '调用删除方法后，对象处于不可用状态
    AssertHaveRight m_oActiveUser, RIGHT_VehicleManagement
    AssertObjIsValid
    DeleteObject
    m_nObjectStatus = ST_NotAvailable
End Sub
'S8
Public Sub Update()
    AssertHaveRight m_oActiveUser, RIGHT_VehicleManagement
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
'S13 2002/11/10 wjb
Public Sub ReCover() '恢复删除的车辆
    AssertHaveRight m_oActiveUser, RIGHT_VehicleManagement
    AssertObjIsValid
    ReCoverObject
    m_nObjectStatus = ST_NormalObj
End Sub
Private Sub ReCoverObject()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    oDb.Execute ("UPDATE Vehicle_info SET delete_status=0 WHERE vehicle_id='" & m_szVehicleID & "'")
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertReCoverObjectError ERR_Vehicle, oDb  '恢复出错
End Sub
'S9
Private Sub RefreshMemoryInfo() '按照对象现在的ID将数据库中相应的信息读至对象内
    Dim oDb As New RTConnection
    Dim rsVehicle As Recordset
    Dim szSql As String
    
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT v.*, c.transport_company_short_name, o.owner_name,t.vehicle_type_short_name  , " _
        & " c2.transport_company_short_name AS splitcompanyname " _
        & " FROM Vehicle_info v,vehicle_type_code t , Company_info c ,owner_info o , Company_info c2 " _
        & " WHERE vehicle_id='" & m_szVehicleID & "' AND v.transport_company_id=c.transport_company_id AND v.owner_id=o.owner_id " _
        & " AND t.vehicle_type_code=v.vehicle_type_code AND c2.transport_company_id=V.split_company_id "

    Set rsVehicle = oDb.Execute(szSql) '"SELECT tbe.*,tbc.transport_company_short_name,tbo.owner_name,tbv.vehicle_type_short_name FROM transport_vehicle_info tbe,vehicle_type_code tbv,transport_company_info tbc,owner_info tbo  WHERE vehicle_id='" & m_szVehicleID & "' AND tbe.transport_company_id=tbc.transport_company_id AND tbe.owner_id=tbo.owner_id AND tbv.vehicle_type_code=tbe.vehicle_type_code")
    If rsVehicle.RecordCount = 0 Then ShowError ERR_VehicleNotExist
    m_szCompanyName = Trim(rsVehicle!transport_company_short_name)
    m_szOwnerName = Trim(rsVehicle!owner_name)
    m_szVehicleModelName = Trim(rsVehicle!vehicle_type_short_name)
    m_szVehicleModel = Trim(rsVehicle!vehicle_type_code)   'P11车型
    m_szOwner = Trim(rsVehicle!owner_id)   'P8车主
    m_szAnnotation = FormatDbValue(rsVehicle!Annotation)    'P3注释
    m_nSeatCount = FormatDbValue(rsVehicle!seat_quantity) 'P6座位数
    m_nStartSeatNo = FormatDbValue(rsVehicle!start_seat_no) 'P7开始座位号
    m_szCardID = FormatDbValue(rsVehicle!register_card_id) 'P10报班卡号
    m_szCompany = Trim(rsVehicle!transport_company_id) 'P4参运公司
    m_szLicenseTag = FormatDbValue(rsVehicle!license_tag_no)    'P9车牌
    m_szVehicleModel = Trim(rsVehicle!vehicle_type_code) 'P11车型
    m_eStatus = rsVehicle!Status 'P5车辆状态
    m_eOldStatus = m_eStatus
    
      m_szSplitCompanyID = FormatDbValue(rsVehicle!split_company_id)
    m_szSplitCompanyName = FormatDbValue(rsVehicle!SplitCompanyName)
    
    
    
    m_oldszCompanyName = m_szCompanyName
    m_oldszOwnerName = m_szOwnerName
    m_oldszVehicleModelName = m_szVehicleModelName
    m_OldszVehicleModel = m_szVehicleModel
    m_oldszAnnotation = m_szAnnotation
    m_oldnSeatCount = m_nSeatCount
    m_OldnStartSeatNo = m_nStartSeatNo
    m_OldszCardID = m_szCardID
    m_oldszCompany = m_szCompany
    m_OldszLicenseTag = m_szLicenseTag
    m_OldszVehicleModel = m_szVehicleModel
    m_oldszSplitCompanyID = m_szSplitCompanyID
    m_oldszSplitCompanyName = m_szSplitCompanyName
    
    Set rsVehicle = Nothing
    Set oDb = Nothing
End Sub

'S10
Private Sub AddObject()
    Dim oDb As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    If Trim(m_szVehicleID) = "" Then ShowError ERR_VehicleIDNotNull
On Error GoTo Here
    Dim rsTemp As Recordset
    szSql = " SELECT * FROM Company_info "
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_SplitCompanyNotExist
    
    szSql = "SELECT license_tag_no FROM Vehicle_info WHERE license_tag_no='" & RTrim(m_szLicenseTag) & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then ShowError ERR_VehicleExist

    szSql = "INSERT Vehicle_info (" _
        & " vehicle_id," _
        & " vehicle_type_code," _
        & " owner_id," _
        & " Annotation," _
        & " seat_quantity," _
        & " start_seat_no," _
        & " register_card_id," _
        & " transport_company_id," _
        & " split_company_id , " _
        & " license_tag_no," _
        & " Status," _
        & " delete_status" _
        & " )  VALUES ('" _
        & m_szVehicleID & "','" _
        & m_szVehicleModel & "','" _
        & m_szOwner & "','" _
        & m_szAnnotation & "', " _
        & m_nSeatCount & " ,'" _
        & m_nStartSeatNo & "','" _
        & m_szCardID & "','" _
        & m_szCompany & "','" _
        & m_szSplitCompanyID & "','" _
        & m_szLicenseTag & "'," _
        & m_eStatus & ",1)"
    
    oDb.Execute (szSql)
    
    WriteOperateLog m_oActiveUser, RIGHT_VehicleManagement, "新增车辆[" & m_szVehicleID & "]"
    
    Set oDb = Nothing
Exit Sub
Here:
    AssertAddObjectError ERR_Vehicle, oDb
    Set oDb = Nothing
End Sub
'S11
Private Sub DeleteObject()
    Dim oDb As New RTConnection
On Error GoTo Here
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    oDb.Execute ("UPDATE Vehicle_info SET delete_status=0 WHERE vehicle_id='" & m_szVehicleID & "'")
    Set oDb = Nothing
Exit Sub
Here:
    AssertDeleteObjectError ERR_Vehicle, oDb
    Set oDb = Nothing
End Sub
'S12
Private Sub UpdateToDB()
    Dim oDb As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
On Error GoTo Here
    Dim rsTemp As Recordset
    Dim szTemp As String
    szSql = " SELECT * FROM Company_info "
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_SplitCompanyNotExist
    
    szSql = "UPDATE Vehicle_info SET vehicle_type_code='" & m_szVehicleModel & "'," _
    & " owner_id='" & m_szOwner & "', " _
    & "  Annotation='" & m_szAnnotation & "', " _
    & " seat_quantity=" & m_nSeatCount & "," _
    & " start_seat_no='" & m_nStartSeatNo & "'," _
    & " register_card_id='" & m_szCardID & "'," _
    & " transport_company_id='" & m_szCompany & "'," _
    & " split_company_id ='" & m_szSplitCompanyID & "', " _
    & " license_tag_no='" & m_szLicenseTag & "'," _
    & " status=" & m_eStatus & " WHERE vehicle_id='" & m_szVehicleID & "'"
    oDb.Execute szSql

If m_szBusID <> "" Then
szTemp = "在车次[" & m_szBusID & "]中，"
End If
szTemp = szTemp & "车辆 [ " & m_szVehicleID & "]"

If m_eOldStatus <> m_eStatus Then
        If m_eOldStatus = ST_VehicleRun Then
            
            szTemp = szTemp & "车牌为[" & m_szLicenseTag & "] 车辆停班 "
'           WriteOperateLog m_oActiveUser, RIGHT_VehicleStop, m_szLicenseTag & ""
        Else
            
            szTemp = szTemp & "车牌为[" & m_szLicenseTag & "] 重新开始运行 "
           
'           WriteOperateLog m_oActiveUser, RIGHT_VehicleStop, m_szLicenseTag & "车辆停班"
        End If
    End If
 
 
 If m_OldszVehicleModel <> m_szVehicleModel And m_oldszVehicleModelName <> m_szVehicleModelName Then
   szTemp = szTemp & "车型[" & m_oldszVehicleModelName & "]改为" & m_szVehicleModelName
 End If
 If m_OldszCardID <> m_szCardID Then
    szTemp = szTemp & "报班卡号由[" & m_OldszCardID & "]改为" & m_szCardID
 End If
 If m_OldszLicenseTag <> m_szLicenseTag Then
   szTemp = szTemp & "车牌由[" & m_OldszLicenseTag & "]改为" & m_szLicenseTag
 End If
 If m_OldszOwner <> m_szOwner And m_oldszOwnerName <> m_szOwnerName Then
    szTemp = szTemp & "车主由[" & m_OldszOwner & "]改为[" & m_szOwner & "]"
 End If
 If m_OldnStartSeatNo <> m_nStartSeatNo Then
   szTemp = szTemp & "座位起始号由[" & m_OldnStartSeatNo & "]改为[" & m_nStartSeatNo & "]"
 End If
 If m_oldnSeatCount <> m_nSeatCount Then
  szTemp = szTemp & "总座位由[" & m_oldnSeatCount & "]改为" & m_nSeatCount
 End If
 If m_odlszAnnotation <> m_szAnnotation Then
  szTemp = szTemp & "注释由[" & m_odlszAnnotation & "]改为" & m_szAnnotation
 End If
 
 If m_oldszCompanyName <> m_szCompanyName Then
   szTemp = szTemp & " 参运公司由[" & m_oldszCompanyName & "]改为" & m_szCompanyName
 End If
 
 If m_oldszSplitCompanyID <> m_szSplitCompanyID Then
      szTemp = szTemp & " 拆帐公司由[" & m_oldszSplitCompanyID & "]改为" & m_szSplitCompanyID
 End If
 
 If Trim(szTemp) <> Trim("车辆 [ " & m_szVehicleID & "]") Then
     
     szTemp = szTemp & "影响车辆对应的所有车次"
     szTemp = GetUnicodeBySize(szTemp, 255)
     WriteOperateLog m_oActiveUser, RIGHT_VehicleManagement, szTemp
 End If
 
 
 Set oDb = Nothing


Exit Sub

Here:
    
    AssertUpdateObjectError ERR_Vehicle, oDb
    Set oDb = Nothing

End Sub

'S13该运行车辆的所有车次都停班
Public Sub AllBusStop(PlanID As String, StartDate As Date, EndDate As Date)
   Dim oDb As New RTConnection
   Dim szSql As String
   AssertHaveRight m_oActiveUser, RIGHT_VehicleAllBusStop
   oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
   szSql = "UPDATE bus_vehicle_code " _
            & " SET stop_start_date='" & ToDBDate(StartDate) & "',stop_end_date='" & ToDBDate(EndDate) & "' " _
            & " WHERE vehicle_id='" & m_szVehicleID & "'"
   oDb.Execute szSql
   szSql = "SELECT Bus_ID " _
            & " FROM bus_vehicle_code " _
            & " WHERE vehicle_id='" & m_szVehicleID & "'"
   szSql = "UPDATE bus_info " _
            & " SET stop_start_date='" & ToDBDate(StartDate) & "',stop_end_date='" & ToDBDate(EndDate) & "' " _
            & " FROM bus_info where Bus_ID in (select Bus_ID From bus_vehicle_code " _
            & " WHERE vehicle_id='" & m_szVehicleID & "'"
   oDb.Execute szSql
   WriteOperateLog m_oActiveUser, RIGHT_VehicleAllBusStop, "计划[" & PlanID & "]中运行车辆" & m_szLicenseTag & "的所有车次在" & Format(StartDate, "YYYY年MM月DD日") & "至" & Format(EndDate, "YYYY年MM月DD日") & "停班"
   Set oDb = Nothing
End Sub



Public Sub UpdateReBus()
 Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim szData() As String
    Dim i As Integer
    Dim szlog As String
    
    
    AssertHaveRight m_oActiveUser, RIGHT_VehicleManagement
    
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "select Bus_id, bus_date FROM enviroment_bus_work WHERE vehicle_id='" & m_szVehicleID & "' and Bus_date>='" & ToDBDate(Date) & "'"
    Set rsTemp = oDb.Execute(szSql)
    
    If rsTemp.RecordCount = 0 Then ShowError ERR_NOBusInEniroment
    For i = 1 To rsTemp.RecordCount
          WriteOperateLog m_oActiveUser, RIGHT_VehicleStop, "车辆 <" & m_szVehicleID & " >信息改变，并更新环境内车次 " & FormatDbValue(rsTemp!Bus_ID) & "发车时间" & FormatDbValue(rsTemp!bus_date)
    Next
    
    szSql = "UPDATE enviroment_bus_info SET vehicle_type_code='" & m_szVehicleModel & "'," _
           & "owner_id='" & m_szOwner & "',total_seat=" & m_nSeatCount & "," _
           & "transport_company_id='" & m_szCompany & "'," _
           & "split_company_id= '" & m_szSplitCompanyID & "'," _
           & "status=" & m_eStatus & " WHERE vehicle_id='" & m_szVehicleID & "' and Bus_date>='" & ToDBDate(Date) & "'"
    oDb.Execute szSql
    
    szlog = "车次 " & FormatDbValue(rsTemp!Bus_ID) & "车辆信息改变："
    
    If m_szVehicleModel <> m_OldszVehicleModel Then
     szlog = szlog & "车型由[" & m_OldszVehicleModel & "]改为" & m_szVehicleModel
    End If
    
    If m_nSeatCount <> m_oldnSeatCount Then
       szlog = szlog & "车辆座位由[" & m_oldnSeatCount & "]改为" & m_nSeatCount
    End If
    
    If m_szCompany <> m_oldszCompany Then
      szlog = szlog & "参运公司由[" & m_oldszCompany & "]改为" & m_szCompany
    End If
    
    If m_szSplitCompanyID <> m_oldszSplitCompanyID Then
       szlog = szlog & "拆帐公司[" & m_oldszSplitCompanyID & "]改为" & m_szSplitCompanyID
    End If
    szlog = GetUnicodeBySize(szlog, 255)
    
    
    WriteOperateLog m_oActiveUser, RIGHT_VehicleManagement, szlog


  Set oDb = Nothing
End Sub
Public Sub UpdateVehicleSeatType(paszSeatType() As String)
    Dim oDb As New RTConnection
    Dim i As Integer
    Dim szSql As String
    Dim nCount As Integer
    
    On Error GoTo Here
    '先删除后重新插入
    AssertObjIsValid
    
    AssertHaveRight m_oActiveUser, RIGHT_VehicleManagement
    oDb.ConnectionString = GetConnectionStr("")
    
    oDb.BeginTrans
    
        szSql = "DELETE vehicle_seat_type_info WHERE vehicle_id = '" & m_szVehicleID & "'"
        oDb.Execute szSql
        
        nCount = ArrayLength(paszSeatType)
        For i = 1 To nCount
            szSql = "INSERT vehicle_seat_type_info ( " _
                & " vehicle_id , " _
                & " seat_type_id , " _
                & " start_seat_no , " _
                & " end_seat_no ) " _
                & " VALUES ('" _
                & m_szVehicleID & "','" _
                & paszSeatType(i, 1) & "','" _
                & paszSeatType(i, 2) & "','" _
                & paszSeatType(i, 3) & "')"
            oDb.Execute szSql
        Next i
    WriteOperateLog m_oActiveUser, RIGHT_VehicleManagement, "车辆 <" & m_szVehicleID & " >座位信息改变"
    oDb.CommitTrans
    Exit Sub
Here:
    oDb.RollbackTrans
    err.Raise err.Number
End Sub
Public Sub DeleteVehicleSeatType()
    Dim oDb As New RTConnection
    Dim i As Integer
    Dim rsTemp As New Recordset
    Dim szSql As String
    Dim nCount As Integer
    Dim lAow As Long
    On Error GoTo Here
    '删除
    AssertObjIsValid
    
    AssertHaveRight m_oActiveUser, RIGHT_VehicleManagement
    
    oDb.ConnectionString = GetConnectionStr("")
    
    oDb.BeginTrans
      
        szSql = "select *  From  vehicle_seat_type_info WHERE vehicle_id = '" & m_szVehicleID & "'"
        Set rsTemp = oDb.Execute(szSql)

        If rsTemp.RecordCount > 0 Then
          szSql = "DELETE vehicle_seat_type_info WHERE vehicle_id = '" & m_szVehicleID & "'"
          oDb.Execute szSql
          WriteOperateLog m_oActiveUser, RIGHT_VehicleManagement, "车辆 <" & m_szVehicleID & " >座位信息被删除"
        End If
    
    Set rsTemp = Nothing
    
    oDb.CommitTrans
    
    Set oDb = Nothing
    
    Exit Sub
Here:
    oDb.RollbackTrans
    Set rsTemp = Nothing
    Set oDb = Nothing
    err.Raise err.Number
    
End Sub


Public Sub AllBusStopEx(PlanID As String, StartDate As Date, EndDate As Date, Optional szBusId As String)
   
   Dim oDb As New RTConnection
   Dim szSql As String
   AssertHaveRight m_oActiveUser, RIGHT_VehicleAllBusStop
   oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
   szSql = "UPDATE bus_vehicle_code " _
            & " SET stop_start_date='" & ToDBDate(StartDate) & "',stop_end_date='" & ToDBDate(EndDate) & "' " _
            & " WHERE vehicle_id='" & m_szVehicleID & "' bus_id='" & szBusId & "'"
   oDb.Execute szSql
   'szSql = "select Bus_ID From bus_vehicle_code WHERE vehicle_id='" & m_szVehicleID & "' AND project_id='" & Trim(PlanID) & "'"
   szSql = "UPDATE bus_info SET stop_start_date='" & ToDBDate(StartDate) & "',stop_end_date='" & ToDBDate(EndDate) & "' where  bus_id='" & szBusId & "' " 'FROM bus_info where Bus_ID in (select Bus_ID From bus_vehicle_code WHERE vehicle_id='" & m_szVehicleID & "' AND project_id='" & Trim(PlanID) & "') and  project_id='" & Trim(PlanID) & "'and bus_id='" & szBusID & "'"
   oDb.Execute szSql
   WriteOperateLog m_oActiveUser, RIGHT_VehicleAllBusStop, "计划[" & PlanID & "]中运行车辆" & m_szLicenseTag & "的所有车次[" & szBusId & "]在" & Format(StartDate, "YYYY年MM月DD日") & "至" & Format(EndDate, "YYYY年MM月DD日") & "停班"
   Set oDb = Nothing

End Sub



