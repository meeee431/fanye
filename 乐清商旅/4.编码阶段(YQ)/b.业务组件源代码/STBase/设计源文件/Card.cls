VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Card"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************
'Member Code:
'Brief Description:IC卡
'Engineer:wjb
'Date Generated:1999/10/11
'Last Revision Date:2002/11/9
'**************************************************
'Option Explicit
'
'Private m_oActiveUser As ActiveUser 'P1所有的类都有的私有类变量
'Private m_nObjectStatus As EObjectStatus 'P2所有的实体类都有的私有类变量
'
''S1
'Public Sub Init(poAUser As ActiveUser)
'    Set m_oActiveUser = poAUser
'    'AssertHaveRight m_oActiveUser, RIGHT_RouteManagement
'End Sub
'
''S2类刚建立时处于不可用状态,此时类不能使用
'Private Sub Class_Initialize()
'    m_nObjectStatus = ST_NotAvailable
'End Sub
'
''S3测试对象的状态是否有效（无效状态[ST_NotAvailable]和新增状态[ST_AddObj]则出错）
'Private Sub AssertStatusAvailable()
'    If m_nObjectStatus = ST_NotAvailable Then ShowError ERR_RouteNotAvailable
'End Sub
''S4
'Private Sub AssertObjIsValid() '测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
'    AssertActiveUserValid m_oActiveUser, ERR_Route
'    AssertStatusAvailable
'End Sub
''S5
'Public Sub Identify(CardID As String)
'    AssertActiveUserValid m_oActiveUser, ERR_Route
'    'RefreshMemoryInfo
'    m_nObjectStatus = ST_NormalObj
'End Sub
'
''S6调用新增方法后，对象处于新增状态。
'Public Sub AddNew()
'    AssertHaveRight m_oActiveUser, RIGHT_RouteManagement
'    AssertActiveUserValid m_oActiveUser, ERR_Route
'    m_nObjectStatus = ST_AddObj
'End Sub
''S7
'Public Sub Delete() '调用删除方法后，对象处于不可用状态
'    AssertHaveRight m_oActiveUser, RIGHT_RouteManagement
'    AssertObjIsValid
'    'DeleteObject
'    m_nObjectStatus = ST_NotAvailable
'End Sub
''S8
'Public Sub Update()
'    AssertHaveRight m_oActiveUser, RIGHT_RouteManagement
'    AssertObjIsValid
'    Select Case m_nObjectStatus
'           Case ST_AddObj
'           'AddObject
'           Case ST_EditObj
'           'UpdateToDB
'           Case ST_NormalObj
'           Case Else
'           ShowError ERR_NotAvailable
'    End Select
'    m_nObjectStatus = ST_NormalObj
'End Sub
'
'Public Function GetAllCard(Optional nStatus As Integer = 0) As String()
'    Dim szaTemp() As String
'    Dim oDb As New RTConnection
'    Dim i As Integer
'    Dim rsTemp As Recordset
'    Dim szSql As String
'
'    'AssertObjIsValid
'    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
'    szSql = "SELECT card_no,vehicle_id FROM card_info"
'    If nStatus <> 5 Then szSql = szSql & " WHERE use_status='" & nStatus & "'"
'    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount <> 0 Then
'    ReDim szaTemp(1 To rsTemp.RecordCount, 1 To 2) As String
'    For i = 1 To rsTemp.RecordCount
'        szaTemp(i, 1) = FormatDbValue(rsTemp!card_no)
'        szaTemp(i, 2) = FormatDbValue(rsTemp!vehicle_id)
'        rsTemp.MoveNext
'    Next
'    GetAllCard = szaTemp
'    End If
'    Set rsTemp = Nothing
'    Set oDb = Nothing
'End Function
