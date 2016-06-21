VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Company"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'**********************************************************
'* Source File Name:Company.cls
'* Project Name:STBase.vbp
'* Engineer:
'* Data Generated:
'* Last Revision Date:
'* Brief Description:公司对象
'* Relational Document:
'**********************************************************
Option Explicit

'出错
Public Enum EErrCompany
    ERR_CompanyExistNotAddNew = ERR_Company + ERR_AddDuplicate '52801已存在该参运公司不能新增
    ERR_CompanyNotDelete = ERR_Company + ERR_DeleteChildExist '52807参运公司已有车辆或车主，所以不能删除
    ERR_CompanyNotAddStatus = ERR_Company + ERR_NotAddObj '52816参运公司对象处于无状态
    ERR_CompanyIDNotNull = ERR_Company + 21 '52821参运公司代码不能为空
    ERR_CompanyNotExist = ERR_Company + 22 ' 52822参运公司不存在
  
End Enum

'权限
Public Enum ERightCompany
    RIGHT_CompanyManagement = ERR_Company + cnMidRightBegin + cnMidRightStep * 1 ' 52906参运公司管理
End Enum

Private m_oActiveUser As ActiveUser 'P1所有的类都有的私有类变量
Private m_nObjectStatus As EObjectStatus 'P2所有的实体类都有的私有类变量
Private m_szCompanyId As String 'P3公司代码
Private m_szCompanyName As String 'P4公司名称
Private m_szCompanyShortName As String 'P5公司简称
Private m_szPrincipal As String 'P6负责人
Private m_szContact As String 'P7联系方式
Private m_szAnnotation As String 'P8注释
Private m_szAccountID As String 'P9账号

'F1获得该参运公司的所有车主
Public Function GetAllOwner() As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute("SELECT * FROM owner_info tbo WHERE transport_company_id='" & m_szCompanyId & "'")
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
        ReDim szaTemp(1 To nCount, 1 To 2) As String
        For i = 1 To nCount
            szaTemp(i, 1) = FormatDbValue(rsTemp!owner_id)
            szaTemp(i, 2) = FormatDbValue(rsTemp!owner_name)
            rsTemp.MoveNext
        Next
        GetAllOwner = szaTemp
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
        RaiseError ERR_AddObj
    ElseIf m_nObjectStatus = ST_EditObj Then
        RaiseError ERR_EditObj
    ElseIf m_nObjectStatus = ST_NormalObj Then
        RaiseError ERR_NormalObj
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
'P3仅用于新增时才能修改对象的标示CompanyID
Public Property Let CompanyId(ByVal vData As String)
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    AssertActiveUserValid m_oActiveUser, ERR_Company
    If m_nObjectStatus <> ST_AddObj Then RaiseError ERR_CompanyNotAddStatus
    If vData = "" Then RaiseError ERR_CompanyIDNotNull
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT * FROM Company_info WHERE transport_company_id='" & RTrim(vData) & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 1 Then RaiseError ERR_CompanyExistNotAddNew
    m_szCompanyId = RTrim(vData)
    Set rsTemp = Nothing
    Set oDb = Nothing
End Property
'P3
Public Property Get CompanyId() As String
    CompanyId = m_szCompanyId
End Property
'P4
Public Property Let CompanyName(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szCompanyName = RTrim(vData)
End Property
'P4
Public Property Get CompanyName() As String
    CompanyName = m_szCompanyName
End Property
'P5
Public Property Let CompanyShortName(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szCompanyShortName = RTrim(vData)
End Property
'P5
Public Property Get CompanyShortName() As String
    CompanyShortName = m_szCompanyShortName
End Property
'P6
Public Property Let Principal(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szPrincipal = RTrim(vData)
End Property
'P6
Public Property Get Principal() As String
    Principal = m_szPrincipal
End Property

'P7
Public Property Let Contact(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szContact = RTrim(vData)
End Property

'P7
Public Property Get Contact() As String
    Contact = m_szContact
End Property

'P8
Public Property Let Annotation(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szAnnotation = RTrim(vData)
End Property

'P8
Public Property Get Annotation() As String
    Annotation = m_szAnnotation
End Property
'P9
Public Property Let AccountID(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szAccountID = RTrim(vData)
End Property

'P9
Public Property Get AccountID() As String
    AccountID = m_szAccountID
End Property
'S1
Public Sub Init(poAUser As ActiveUser)
    Set m_oActiveUser = poAUser
    'AssertHaveRight m_oActiveUser, RIGHT_CompanyManagement
End Sub

'S2类刚建立时处于不可用状态,此时类不能使用
Private Sub Class_Initialize()
    m_nObjectStatus = ST_NotAvailable
End Sub

'S3测试对象的状态是否有效（无效状态[ST_NotAvailable]和新增状态[ST_AddObj]则出错）
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then RaiseError ERR_NotAvailable
End Sub
'S4
Private Sub AssertObjIsValid() '测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
    AssertActiveUserValid m_oActiveUser, ERR_Company
    AssertStatusAvailable
End Sub
'S5
Public Sub Identify(pszCompanyID As String)
    AssertActiveUserValid m_oActiveUser, ERR_Company
    m_szCompanyId = RTrim(pszCompanyID)
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
    AssertHaveRight m_oActiveUser, RIGHT_CompanyManagement
    AssertActiveUserValid m_oActiveUser, ERR_Company
    m_nObjectStatus = ST_AddObj
End Sub
'S7
Public Sub Delete() '调用删除方法后，对象处于不可用状态
    AssertHaveRight m_oActiveUser, RIGHT_CompanyManagement
    AssertObjIsValid
    DeleteObject
    m_nObjectStatus = ST_NotAvailable
End Sub
'S13
Public Sub ReCover() '恢复删除的公司
    AssertHaveRight m_oActiveUser, RIGHT_CompanyManagement
    AssertObjIsValid
    ReCoverObject
    m_nObjectStatus = ST_NormalObj
End Sub
Private Sub ReCoverObject()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    oDb.Execute ("UPDATE Company_info SET delete_status=1 WHERE transport_company_id='" & m_szCompanyId & "'")
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertReCoverObjectError ERR_Company, oDb '恢复出错
End Sub
'S8
Public Sub Update()
    AssertHaveRight m_oActiveUser, RIGHT_CompanyManagement
    AssertObjIsValid
    Select Case m_nObjectStatus
           Case ST_AddObj
           AddObject
           Case ST_EditObj
           UpdateToDB
           Case ST_NormalObj
           Case Else
           RaiseError ERR_NotAvailable
    End Select
    m_nObjectStatus = ST_NormalObj
End Sub

'S9
Private Sub RefreshMemoryInfo() '按照对象现在的ID将数据库中相应的信息读至对象内
    Dim oDb As New RTConnection
    Dim rsCompany As Recordset
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsCompany = oDb.Execute("SELECT * FROM Company_info WHERE transport_company_id='" & m_szCompanyId & "'")
    If rsCompany.RecordCount = 0 Then RaiseError ERR_CompanyNotExist
    m_szCompanyShortName = FormatDbValue(rsCompany!transport_company_short_name)
    m_szCompanyName = FormatDbValue(rsCompany!transport_company_name)
    m_szContact = FormatDbValue(rsCompany!Contact)
    m_szAnnotation = FormatDbValue(rsCompany!Annotation)
    m_szPrincipal = FormatDbValue(rsCompany!manager)
    m_szAccountID = Trim(rsCompany!account_id)
    Set rsCompany = Nothing
    Set oDb = Nothing
End Sub

'S10
Private Sub AddObject()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    If Trim(m_szCompanyId) = "" Then RaiseError ERR_CompanyIDNotNull
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    oDb.Execute ("INSERT Company_info (transport_company_id,transport_company_short_name," _
    & " transport_company_name,manager,contact,annotation,account_id,delete_status) VALUES ('" & m_szCompanyId & "','" & m_szCompanyShortName _
    & "','" & m_szCompanyName & "','" & m_szPrincipal & "',' " & m_szContact & " ','" & m_szAnnotation & "','" & m_szAccountID & "',1)")
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertAddObjectError ERR_Company, oDb
End Sub
'S11
Private Sub DeleteObject()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    oDb.Execute ("UPDATE Company_info SET delete_status=0 WHERE transport_company_id='" & m_szCompanyId & "'")
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertDeleteObjectError ERR_Company, oDb
End Sub
'S12
Private Sub UpdateToDB()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "UPDATE Company_info SET transport_company_short_name='" & m_szCompanyShortName & "'," _
    & "transport_company_name='" & m_szCompanyName & "',manager='" & m_szPrincipal & "',contact='" & m_szContact & "'," _
    & "annotation='" & m_szAnnotation & "',account_id='" & m_szAccountID & "'" _
    & "WHERE transport_company_id='" & m_szCompanyId & "'"
    oDb.Execute szSql
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertUpdateObjectError ERR_Company, oDb
End Sub

