VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TicketType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************
'* Source File Name:TicketType.cls
'* Project Name:STSystem.vbp
'* Engineer:陈峰
'* Data Generated:2002-08-15
'* Last Revision Date:2002-08-15
'* Brief Description:票种
'* Relational Document:
'**********************************************************
Option Explicit
'出错
Public Enum EErrTicketType
    ERR_TicketTypeNotAvailable = ERR_TicketType + ERR_NotAvailable   '29612票种对象处于无状态
    ERR_TicketTypeNotExist = ERR_TicketType + 21 '29621票种无存在
End Enum

'权限
Public Enum ERightTicketType
    RIGHT_TicketType = ERR_TicketType + cnMidRightBegin + cnMidRightStep * 1 '29706 票种 ...权限
End Enum


Private m_oActiveUser As ActiveUser 'P1所有的类都有的私有类变量
Private m_nObjectStatus As EObjectStatus 'P2所有的实体类都有的私有类变量
Private m_nTicketTypeID As String 'P3票种代码
Private m_szTicketTypeName As String 'P4票种名称
Private m_nTicketTypeValid As Integer 'P5票种是否可用
Private m_szAnnotation As String 'P6注释

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
Public Property Get TicketTypeID() As String
    TicketTypeID = m_nTicketTypeID
End Property
'P4
Public Property Let TicketTypeName(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szTicketTypeName = vData
End Property
'P4
Public Property Get TicketTypeName() As String
    TicketTypeName = m_szTicketTypeName
End Property
'P5
Public Property Let TicketTypeValid(ByVal vData As Integer)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_nTicketTypeValid = vData
End Property
'P5
Public Property Get TicketTypeValid() As Integer
    TicketTypeValid = m_nTicketTypeValid
End Property

'P6
Public Property Let Annotation(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szAnnotation = vData
End Property
'P6
Public Property Get Annotation() As String
    Annotation = m_szAnnotation
End Property


'S1
Public Sub Init(poAUser As ActiveUser)
    Set m_oActiveUser = poAUser
    'AssertHaveRight m_oActiveUser, RIGHT_TicketType
End Sub

'S2类刚建立时处于不可用状态,此时类不能使用
Private Sub Class_Initialize()
    m_nObjectStatus = ST_NotAvailable
End Sub

'S3测试对象的状态是否有效（无效状态[ST_NotAvailable]和新增状态[ST_AddObj]则出错）
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then ShowError ERR_TicketTypeNotAvailable
End Sub
'S4
Private Sub AssertObjIsValid() '测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
    AssertActiveUserValid m_oActiveUser, ERR_TicketType
    AssertStatusAvailable
End Sub
'S5
Public Sub Identify(Optional TicketTypeID As String = "", Optional TicketTypeInput As String = "")
    AssertActiveUserValid m_oActiveUser, ERR_TicketType
    m_nTicketTypeID = RTrim(TicketTypeID)
    RefreshMemoryInfo
    m_nObjectStatus = ST_NormalObj
End Sub

'S6
Public Sub Update()
    AssertHaveRight m_oActiveUser, RIGHT_TicketType
    AssertObjIsValid
    Select Case m_nObjectStatus
           Case ST_EditObj
           UpdateToDB
    End Select
    m_nObjectStatus = ST_NormalObj
End Sub

'S7
Private Sub RefreshMemoryInfo() '按照对象现在的ID将数据库中相应的信息读至对象内
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    Set rsTemp = oDb.Execute("SELECT * FROM Ticket_type_code  WHERE Ticket_Type_id=" & m_nTicketTypeID)
    
    If rsTemp.RecordCount = 0 Then ShowError ERR_TicketTypeNotExist
    m_szTicketTypeName = Trim(rsTemp!Ticket_Type_name)
    m_nTicketTypeValid = FormatDbValue(rsTemp!ticket_type_valid)
    m_szAnnotation = FormatDbValue(rsTemp!Annotation)
    Set oDb = Nothing
End Sub


'S8
Private Sub UpdateToDB()
    Dim oDb As New RTConnection
    Dim szTempSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
On Error GoTo Here
    If m_nTicketTypeValid <> TP_TicketTypeNotValid Then m_nTicketTypeValid = TP_TicketTypeValid
    szTempSql = "UPDATE Ticket_type_code SET " _
            & "  Ticket_Type_name='" & m_szTicketTypeName & "'," _
            & "  Ticket_Type_valid=" & m_nTicketTypeValid & "," _
            & "  Annotation='" & m_szAnnotation & "'" _
            & " WHERE Ticket_Type_id='" & m_nTicketTypeID & "'"
    oDb.Execute (szTempSql)
Exit Sub
Here:
    RaiseError ERR_TicketType, oDb
End Sub


