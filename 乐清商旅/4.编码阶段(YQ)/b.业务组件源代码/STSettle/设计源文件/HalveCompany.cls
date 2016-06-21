VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HalveCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E81E4E702FD"
'
Option Base 0
Option Explicit

'set this to 0 to disable debug code in this class
#Const DebugMode = 1
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3E8219C90203
    Private mlClassDebugID As Long
#End If

'##ModelId=3E81E4F6006D
Private m_szRouteID As String

'##ModelId=3E81E4FA0290
Private m_szRouteName As String

'##ModelId=3E81E4FF0203
Private m_szCompanyID As String

'##ModelId=3E81E505000F
Private m_szCompanyName As String

'##ModelId=3E81E50803B9
Private m_szOtherCompanyId As String
Private m_szOtherCompanyName As String
'##ModelId=3E81E50E00CB
Private mOtherCompanyName As String

Private m_dbRatio As Double


Private m_oActiveUser As ActiveUser
Private m_nObjectStatus As EObjectStatus '对象状态
'##ModelId=3E81E6F3031C
Public Sub Init(opAUser As ActiveUser)
    Set m_oActiveUser = opAUser
End Sub

'##ModelId=3E81E6F9038A
Public Sub Identify(pszRouteID As String, pszCompanyID As String)
    AssertActiveUserValid m_oActiveUser, ERR_HalfTicketPrice
    m_szRouteID = pszRouteID
    m_szCompanyID = pszCompanyID
    RefreshObject
    m_nObjectStatus = ST_NormalObj
End Sub
Private Sub RefreshObject()
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "SELECT h.route_id,h.route_name,h.comapny_id,c.company_name,h.other_transport_company_id,c1.company_name , h.ratio " _
        & " FROM fin_halve_company_lst h," & cszTableCompany & " c," & cszTableCompany & " c1" _
        & " WHERE h.route_id = " & TransFieldValueToString(m_szRouteID) _
        & " AND h.company_id = " & TransFieldValueToString(m_szCompanyID) _
        & " AND h.other_transport_company_id =c1.company_id " _
        & " AND h.transport_company_id =c.company_id"
    
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then RaiseError ERR_HalveCompanyNotExist  '指定的线路号和公司号不存在，出错
    m_szCompanyID = FormatDbValue(rsTemp!comapny_id)
    m_szCompanyName = FormatDbValue(rsTemp!company_name)
    m_szRouteID = FormatDbValue(rsTemp!route_id)
    m_szRouteName = FormatDbValue(rsTemp!route_name)
    m_szOtherCompanyId = FormatDbValue(rsTemp!other_transport_company_id)
    m_szOtherCompanyName = FormatDbValue(rsTemp!company_name)
    m_dbRatio = FormatDbValue(rsTemp!Ratio)
    
End Sub
'##ModelId=3E81E6FE0186
Public Sub AddNew()
    AssertActiveUserValid m_oActiveUser, ERR_HalfTicketPrice
    m_nObjectStatus = ST_AddObj
    m_szRouteID = ""
    m_szCompanyID = ""
    m_szOtherCompanyId = ""
End Sub

'##ModelId=3E81E703003E
Private Sub AddObject()
    On Error GoTo AddObjectErr
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "INSERT fin_halve_company_lst (route_id,route_name , transport_company_id,other_transport_company_id,ratio) " _
        & " VALUES ( " _
        & TransFieldValueToString(m_szRouteID) & "," _
        & TransFieldValueToString(m_szRouteName) & "," _
        & TransFieldValueToString(m_szCompanyID) & "," _
        & TransFieldValueToString(m_szOtherCompanyId) & "," _
        & TransFieldValueToString(m_dbRatio) & ")"
    oDb.Execute (szSql)
    Exit Sub
AddObjectErr:
    Call RaiseError(ERR_HalveCompanyExistNotAdd)
End Sub

'##ModelId=3E81E7070232
Public Sub UpDate()
    AssertActiveUserValid m_oActiveUser, ERR_HalfTicketPrice
    AssertHaveRight m_oActiveUser, RIGHT_HalveCompany
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

'##ModelId=3E81E70E005D
Private Sub UpDateToDB()
    On Error GoTo UpDateDBErr
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "UPDATE fin_halve_company_lst SET " _
        & " route_id = " & TransFieldValueToString(m_szRouteID) & "," _
        & " route_name = " & TransFieldValueToString(m_szRouteName) & "," _
        & " transport_company_id = " & TransFieldValueToString(m_szCompanyID) & "," _
        & " other_transport_company_id = " & TransFieldValueToString(m_szOtherCompanyId) & "," _
        & " ratio =  " & TransFieldValueToString(m_dbRatio) _
        & " WHERE route_id = " & TransFieldValueToString(m_szRouteID) _
        & " AND transport_company_id = " & TransFieldValueToString(m_szCompanyID)
    oDb.Execute (szSql)
    Exit Sub
UpDateDBErr:
    Call RaiseError(ERR_HalveCompanyExistNotAdd)
End Sub

'##ModelId=3E81E7120232
Public Sub Delete()
    AssertActiveUserValid m_oActiveUser, ERR_HalveCompany
    AssertHaveRight m_oActiveUser, RIGHT_HalveCompany
    AssertObjIsValid
    DeleteObject
End Sub

'##ModelId=3E81E7170157
Private Sub DeleteObject()
    On Error GoTo DeleteObjectErr
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "DELETE fin_halve_company_lst " _
        & " WHERE route_id = " & TransFieldValueToString(m_szRouteID) _
        & " AND transport_company_id = " & TransFieldValueToString(m_szCompanyID)
    oDb.Execute (szSql)

    Exit Sub
DeleteObjectErr:
    Call RaiseError(ERR_HalveCompanyNotExist)
End Sub

''##ModelId=3E8219CA0148
'Private Sub Class_Initialize()
'    #If DebugMode Then
'        'get the next available class ID, and print out
'        'that the class was created successfully
'        mlClassDebugID = GetNextClassDebugID()
'        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
'    #End If
'End Sub
'
''##ModelId=3E8219CA01A5
'Private Sub Class_Terminate()
'    #If DebugMode Then
'    'the class is being destroyed
'    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
'    #End If
'End Sub
'
'#If DebugMode Then
'    '##ModelId=3E8225DA01A5
'    Public Property Get ClassDebugID() As Long
'        'if we are in debug mode, surface this property that consumers can query
'        ClassDebugID = mlClassDebugID
'    End Property
'#End If

'##ModelId=3E822930034B
Public Property Get OtherCompanyName() As String
    OtherCompanyName = mOtherCompanyName
End Property

'##ModelId=3E82293002AF
Public Property Get OtherCompanyId() As String
    OtherCompanyId = m_szOtherCompanyId
End Property

'##ModelId=3E8229300186
Public Property Let OtherCompanyId(ByVal Value As String)
    m_szOtherCompanyId = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E82293000FA
Public Property Get CompanyName() As String
    CompanyName = m_szCompanyName
End Property

'##ModelId=3E822930005D
Public Property Get CompanyID() As String
    CompanyID = m_szCompanyID
End Property

'##ModelId=3E82292F034B
Public Property Let CompanyID(ByVal Value As String)
    m_szCompanyID = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E82292F02BF
Public Property Get RouteName() As String
    RouteName = m_szRouteName
End Property

'##ModelId=3E82292F034B
Public Property Let RouteName(ByVal Value As String)
    m_szRouteName = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property
'##ModelId=3E82292E037A
Public Property Get RouteID() As String
    RouteID = m_szRouteID
End Property
'##ModelId=3E82293002AF
Public Property Get Ratio() As Double
    Ratio = m_dbRatio
End Property

'##ModelId=3E8229300186
Public Property Let Ratio(ByVal Value As Double)
    m_dbRatio = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property


'##ModelId=3E82292E01E4
Public Property Let RouteID(ByVal Value As String)
    m_szRouteID = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'**************************************************
'Member Code:S3
'Brief Description:测试对象是否有效
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then RaiseError ERR_HalveCompanyNotAvailable
End Sub

'**************************************************
'Member Code:S4
'Brief Description:测试对象是否有效
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub AssertObjIsValid()
    AssertActiveUserValid m_oActiveUser, ERR_HalfTicketPrice
    AssertStatusAvailable
End Sub

