VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Owner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Company"

'**********************************************************
'* Source File Name:Owner.cls
'* Project Name:STBase.vbp
'* Engineer:
'* Data Generated:
'* Last Revision Date:
'* Brief Description:车主对象
'* Relational Document:
'**********************************************************
Option Explicit

'出错
Public Enum EErrOwner
    ERR_OwnerExistNotAddNew = ERR_Owner + ERR_AddDuplicate '53601已存在该车主不能新增
    ERR_OwnerNotDelete = ERR_Owner + ERR_DeleteChildExist '53607车主已有车辆所以不能删除
    ERR_OwnerNotCompany = ERR_Owner + ERR_EditChildExist '53605参运公司不存在，不能修改
    ERR_OwnerNotAddStatus = ERR_Owner + ERR_NotAddObj '53616车主对象处于无状态
    ERR_OwnerIDNotNull = ERR_Owner + 21 '53621车主代码不能为空
    ERR_OwnerNotExist = ERR_Owner + 22 ' 53622车主不存在
    
    ERR_SplitCompanyNotExist = ERR_Owner + 23 '53623拆帐公司不存在
    
End Enum

'权限
Public Enum ERightOwner
    RIGHT_OwnerManagement = ERR_Owner + cnMidRightBegin + cnMidRightStep * 1 '53706车主管理
End Enum

Private m_oActiveUser As ActiveUser 'P1所有的类都有的私有类变量
Private m_nObjectStatus As EObjectStatus 'P2所有的实体类都有的私有类变量
Private m_szOwnerID As String 'P3车主代码
Private m_szOwnerName As String 'P4车主名称
Private m_szAddress As String 'P5车主地址
Private m_szCompany As String 'P6车主参运公司
Private m_szContact As String 'P7车主联系方式
Private m_szAnnotation As String 'P8注释
Private m_szIDCard As String 'P9身份证号
Private m_szCompanyName As String 'P10参运公司名称
Private m_szAccountID As String 'P11账号
Private m_szSplitCompanyName As String '拆帐公司名称
Private m_szSplitCompanyID As String '拆帐公司代码

'F1获得该车主的所有的车辆
Public Function GetAllVehicle() As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute("SELECT vehicle_id,license_tag_no FROM Vehicle_info WHERE owner_id='" & m_szOwnerID & "'")
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
'P2
Public Property Let ObjStatus(ByVal vData As EObjectStatus)
    m_nObjectStatus = vData
End Property
'P3
Public Property Let OwnerID(ByVal vData As String)
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    AssertActiveUserValid m_oActiveUser, ERR_Owner
    '查询是否是新增状态
    If m_nObjectStatus <> ST_AddObj Then ShowError ERR_OwnerNotAddStatus
    '输入的主键是否为空
    If vData = "" Then ShowError ERR_OwnerIDNotNull
    '查询该对象是否已有记录
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT * FROM owner_info WHERE owner_id='" & RTrim(vData) & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 1 Then ShowError ERR_OwnerExistNotAddNew
    m_szOwnerID = RTrim(vData)
    Set rsTemp = Nothing
    Set oDb = Nothing
End Property
'P3
Public Property Get OwnerID() As String
    OwnerID = m_szOwnerID
End Property
'P4
Public Property Let OwnerName(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szOwnerName = vData
End Property
'P4
Public Property Get OwnerName() As String
    OwnerName = m_szOwnerName
End Property
'P5
Public Property Let Address(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szAddress = vData
End Property
'P5
Public Property Get Address() As String
    Address = m_szAddress
End Property
'P6
Public Property Let Company(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szCompany = vData
End Property
'P6
Public Property Get Company() As String
    Company = m_szCompany
End Property


'P7
Public Property Let Contact(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szContact = vData
End Property
'P7
Public Property Get Contact() As String
    Contact = m_szContact
End Property
'P8
Public Property Let Annotation(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szAnnotation = vData
End Property
'P8
Public Property Get Annotation() As String
    Annotation = m_szAnnotation
End Property
'P9
Public Property Get IDCard() As String
    IDCard = m_szIDCard
End Property
'P9
Public Property Let IDCard(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szIDCard = vData
End Property
'P10
Public Property Get CompanyName() As String
       CompanyName = m_szCompanyName
End Property
'P11
Public Property Let AccountID(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szAccountID = RTrim(vData)
End Property

'P11
Public Property Get AccountID() As String
    AccountID = m_szAccountID
End Property

'P12
Public Property Let SplitCompanyID(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szSplitCompanyID = vData
End Property
'P12
Public Property Get SplitCompanyID() As String
    SplitCompanyID = m_szSplitCompanyID
End Property
'P13
Public Property Get SplitCompanyName() As String
       SplitCompanyName = m_szSplitCompanyName
End Property

'S1
Public Sub Init(poAUser As ActiveUser)
    Set m_oActiveUser = poAUser
    'AssertHaveRight m_oActiveUser, RIGHT_OwnerManagement
End Sub

'S2类刚建立时处于不可用状态,此时类不能使用
Private Sub Class_Initialize()
    m_nObjectStatus = ST_NotAvailable
End Sub

'S3测试对象的状态是否有效（无效状态[ST_NotAvailable]和新增状态[ST_AddObj]则出错）
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then ShowError ERR_NotAvailable
End Sub
'S4
Private Sub AssertObjIsValid() '测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
    AssertActiveUserValid m_oActiveUser, ERR_Owner
    AssertStatusAvailable
End Sub
'S5
Public Sub Identify(OwnerID As String)
    AssertActiveUserValid m_oActiveUser, ERR_Owner
    m_szOwnerID = RTrim(OwnerID)
    RefreshMemoryInfo
    m_nObjectStatus = ST_NormalObj
End Sub
'**********************************************************
'* Engineer:wjb
'* Data Generated:
'* Last Revision Date:2002/11/9
'**********************************************************
'S6调用新增方法后，对象处于新增状态。
Public Sub AddNew()
    AssertHaveRight m_oActiveUser, RIGHT_OwnerManagement
    AssertActiveUserValid m_oActiveUser, ERR_Owner
    m_nObjectStatus = ST_AddObj
End Sub
'S7
Public Sub Delete() '调用删除方法后，对象处于不可用状态
    AssertHaveRight m_oActiveUser, RIGHT_OwnerManagement
    AssertObjIsValid
    DeleteObject
    m_nObjectStatus = ST_NotAvailable
End Sub
'S13
Public Sub ReCover() '恢复删除的车主
    AssertHaveRight m_oActiveUser, RIGHT_CompanyManagement
    AssertObjIsValid
    ReCoverObject
    m_nObjectStatus = ST_NormalObj
End Sub
Private Sub ReCoverObject()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    oDb.Execute ("UPDATE owner_info SET delete_status=1 WHERE owner_id='" & m_szOwnerID & "'")
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertReCoverObjectError ERR_Owner, oDb '恢复出错
End Sub
'S8
Public Sub Update()
    AssertHaveRight m_oActiveUser, RIGHT_OwnerManagement
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
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = " SELECT o.* , c.transport_company_short_name AS companyname , " _
        & " c2.transport_company_short_name AS splitcompanyname " _
        & " FROM owner_info o,Company_info c , Company_info c2 " _
        & " WHERE owner_id='" & m_szOwnerID & "' AND c.transport_company_id=o.transport_company_id " _
        & " AND c2.transport_company_id=o.split_company_id "
            
    Set rsTemp = oDb.Execute(szSql) '"SELECT tbo.*,tbc.transport_company_short_name AS companyname  FROM owner_info tbo,transport_company_info tbc WHERE owner_id='" & m_szOwnerID & "' AND tbc.transport_company_id=tbo.transport_company_id")
    If rsTemp.RecordCount = 0 Then ShowError ERR_OwnerNotExist
    m_szAddress = FormatDbValue(rsTemp!Address)
    m_szAnnotation = FormatDbValue(rsTemp!Annotation)
    m_szCompany = FormatDbValue(rsTemp!transport_company_id)
    m_szSplitCompanyID = FormatDbValue(rsTemp!split_company_id) '拆帐公司代码
    m_szSplitCompanyName = FormatDbValue(rsTemp!SplitCompanyName) '拆帐公司名称
    m_szContact = FormatDbValue(rsTemp!Contact)
    m_szOwnerName = FormatDbValue(rsTemp!owner_name)
    m_szIDCard = FormatDbValue(rsTemp!id_card)
    m_szCompanyName = FormatDbValue(rsTemp!CompanyName)
    m_szAccountID = FormatDbValue(rsTemp!account_id)
    
    

    Set rsTemp = Nothing
    Set oDb = Nothing
End Sub

'S10
Private Sub AddObject()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    If Trim(m_szOwnerID) = "" Then ShowError ERR_OwnerIDNotNull
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)

    szSql = " SELECT * FROM Company_info WHERE transport_company_id = '" & m_szCompany & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_SplitCompanyNotExist
    If m_szCompany <> m_szSplitCompanyID Then
        szSql = " SELECT * FROM Company_info WHERE transport_company_id = '" & m_szSplitCompanyID & "'"
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount = 0 Then ShowError ERR_SplitCompanyNotExist
    End If
    szSql = "INSERT owner_info (" _
        & "owner_id," _
        & "owner_name," _
        & "address," _
        & "transport_company_id," _
        & "split_company_id," _
        & "id_card," _
        & "contact," _
        & "Annotation," _
        & "account_id," _
        & "delete_status)" _
        & " VALUES ('" & m_szOwnerID & "'," _
        & " '" & m_szOwnerName & "'," _
        & " '" & m_szAddress & "'," _
        & " '" & m_szCompany & "'," _
        & " '" & m_szSplitCompanyID & "'," _
        & " '" & m_szIDCard & "'," _
        & " '" & m_szContact & "'," _
        & " '" & m_szAnnotation & "'," _
        & " '" & m_szAccountID & "',1)"
    
    oDb.Execute (szSql)
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertAddObjectError ERR_Owner, oDb
End Sub
'S11
Private Sub DeleteObject()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    oDb.Execute ("UPDATE owner_info SET delete_status=0 WHERE owner_id='" & m_szOwnerID & "'")
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertDeleteObjectError ERR_Owner, oDb
End Sub

'S12
Private Sub UpdateToDB()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = " SELECT * FROM Company_info "
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_SplitCompanyNotExist
    
    szSql = "UPDATE owner_info SET " _
        & "owner_name='" & m_szOwnerName & "'," _
        & "address='" & m_szAddress & "'," _
        & "transport_company_id='" & m_szCompany & "'," _
        & "split_company_id='" & m_szSplitCompanyID & "'," _
        & "contact='" & m_szContact & "'," _
        & "id_card='" & m_szIDCard & "'," _
        & "Annotation='" & m_szAnnotation & "'," _
        & "account_id='" & m_szAccountID & "'" _
        & " WHERE owner_id='" & m_szOwnerID & "'" _
    
    oDb.Execute (szSql)
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertUpdateObjectError ERR_Owner, oDb
End Sub




