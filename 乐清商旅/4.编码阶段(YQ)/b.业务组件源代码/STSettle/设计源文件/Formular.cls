VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Formular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E80E98703A9"
'此接口为设置公式使用
Option Base 0
Option Explicit

'set this to 0 to disable debug code in this class
#Const DebugMode = 1
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3E8219BF0177
    Private mlClassDebugID As Long
#End If

'公式代码
'##ModelId=3E80E9970148
Private m_szFormularID As String

'公式名称
'##ModelId=3E80E9AA001F
Private m_szFormularName As String

'公式内容
'##ModelId=3E80E9AF02BF
Private m_szFormularContent As String
Private m_oActiveUser As ActiveUser
Private m_nObjectStatus As EObjectStatus '对象状态

'##ModelId=3E80E9E001C5
Public Sub Init(opAUser As ActiveUser)
    Set m_oActiveUser = opAUser
End Sub

'##ModelId=3E80E9EA000F
Public Sub Identify(pszFormularID As String)
    AssertActiveUserValid m_oActiveUser, ERR_Formular
    m_szFormularID = pszFormularID
    RefreshObject
    m_nObjectStatus = ST_NormalObj
    Exit Sub
End Sub
Private Sub RefreshObject()
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr()
    szSql = "SELECT * FROM settle_formula_info WHERE formuar_id = " & TransFieldValueToString(m_szFormularID)
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then RaiseError ERR_FormularNotExist  '公式号不存在
    m_szFormularID = FormatDbValue(rsTemp!formuar_id)
    m_szFormularName = FormatDbValue(rsTemp!formular_name)
    m_szFormularContent = FormatDbValue(rsTemp!formular_content)
End Sub
'##ModelId=3E80E9EF0167
Public Sub AddNew()
    AssertActiveUserValid m_oActiveUser, ERR_Formular
    m_nObjectStatus = ST_AddObj
    m_szFormularContent = ""
    m_szFormularID = ""
    m_szFormularName = ""
End Sub

'##ModelId=3E80E9F4008C
Private Sub AddObject()
    On Error GoTo AddObjectErr
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "INSERT settle_formula_info (formuar_id,formular_name,formular_content) " _
        & " VALUES( " _
        & TransFieldValueToString(m_szFormularID) & "," _
        & TransFieldValueToString(m_szFormularName) & "," _
        & TransFieldValueToString(m_szFormularContent) & ")"
    oDB.Execute (szSql)
    Exit Sub
AddObjectErr:
    err.Raise err.Number, err.Source, err.Description   '主键重复，出错
End Sub

'##ModelId=3E80E9F9002E
Public Sub UpDate()
    AssertActiveUserValid m_oActiveUser, ERR_Formular
    AssertHaveRight m_oActiveUser, RIGHT_Formular
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

'##ModelId=3E80E9FD02AF
Private Sub UpDateToDB()
    On Error GoTo UpDateDBErr
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = " UPDATE settle_formula_info SET " _
        & " formular_name = " & TransFieldValueToString(m_szFormularName) & "," _
        & " formular_content =" & TransFieldValueToString(m_szFormularContent) _
        & " WHERE formuar_id = " & TransFieldValueToString(m_szFormularID)
    oDB.Execute (szSql)
    Exit Sub
UpDateDBErr:
    err.Raise err.Number, err.Source, err.Description
End Sub

'##ModelId=3E80EA0202CE
Public Sub Delete()
    AssertActiveUserValid m_oActiveUser, ERR_Formular
    AssertHaveRight m_oActiveUser, RIGHT_Formular
    AssertObjIsValid
    DeleteObject
End Sub

'##ModelId=3E80EA060280
Private Sub DeleteObject()
    On Error GoTo DeleteObjectErr
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "DELETE settle_formula_info " _
        & " WHERE  formuar_id = " & TransFieldValueToString(m_szFormularID)
    oDB.Execute (szSql)
    Exit Sub
DeleteObjectErr:
    err.Raise err.Number, err.Source, err.Description
End Sub
'
''##ModelId=3E8219C0006D
'Private Sub Class_Initialize()
'    #If DebugMode Then
'        'get the next available class ID, and print out
'        'that the class was created successfully
'        mlClassDebugID = GetNextClassDebugID()
'        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
'    #End If
'End Sub
'
''##ModelId=3E8219C000CB
'Private Sub Class_Terminate()
'    #If DebugMode Then
'    'the class is being destroyed
'    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
'    #End If
'End Sub
'
'#If DebugMode Then
'    '##ModelId=3E8225D801A5
'    Public Property Get ClassDebugID() As Long
'        'if we are in debug mode, surface this property that consumers can query
'        ClassDebugID = mlClassDebugID
'    End Property
'#End If

'##ModelId=3E822915003E
Public Property Get FormularContent() As String
    FormularContent = m_szFormularContent
End Property

'##ModelId=3E82291402FD
Public Property Let FormularContent(ByVal Value As String)
    m_szFormularContent = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E8229140251
Public Property Get FormularName() As String
    FormularName = m_szFormularName
End Property

'##ModelId=3E8229140138
Public Property Let FormularName(ByVal Value As String)
    m_szFormularName = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E82291400AB
Public Property Get FormularID() As String
    FormularID = m_szFormularID
End Property


'##ModelId=3E822913038A
Public Property Let FormularID(ByVal Value As String)
    m_szFormularID = Value
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
    If m_nObjectStatus = ST_NotAvailable Then RaiseError ERR_FormularNotAvailable
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

Public Function GetPriceItem() As String()
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset, i As Integer
    Dim szSql As String
    Dim aszTemp() As String
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "SELECT * FROM price_item_info WHERE use_mark =1 "
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 3)
    For i = 1 To rsTemp.RecordCount
        aszTemp(i, 1) = FormatDbValue(rsTemp!price_item)
        aszTemp(i, 2) = FormatDbValue(rsTemp!chinese_name)
        aszTemp(i, 3) = FormatDbValue(rsTemp!use_mark)
        rsTemp.MoveNext
    Next i
    GetPriceItem = aszTemp
End Function
