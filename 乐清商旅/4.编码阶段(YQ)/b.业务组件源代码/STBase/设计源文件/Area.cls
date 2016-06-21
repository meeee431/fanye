VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Area"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**********************************************************
'* Source File Name:Area.cls
'* Project Name:STBase.vbp
'* Engineer:
'* Data Generated:
'* Last Revision Date:
'* Brief Description:地区对象
'* Relational Document:
'**********************************************************

Option Explicit
'出错
Public Enum EErrArea
    ERR_AreaExistNotAdd = ERR_Area + ERR_AddDuplicate '55201不能新增，因为该地区已经存在
    ERR_AreaNoDelete = ERR_Area + ERR_DeleteChildExist '55207该地区已安排了路段或站点，不能删除
    ERR_AreaNotAddStatus = ERR_Area + ERR_NotAddObj '55216地区对象不处于新增状态
    ERR_AreaNotAvailable = ERR_Route + ERR_NotAvailable '55212地区对象处于无状态
    ERR_AreaNotExist = ERR_Area + 21 '55221没有该地区
    ERR_AreaIDNotNull = ERR_Area + 22 '55222地区代码不能为空
End Enum

'权限
Public Enum ERightArea
    RIGHT_AreaManagement = ERR_Area + cnMidRightBegin + cnMidRightStep * 1 '55306管理地区的权限
End Enum

Private m_oActiveUser As ActiveUser 'P1所有的类都有的私有类变量
Private m_nObjectStatus As EObjectStatus 'P2所有的实体类都有的私有类变量
Private m_szAreaCode As String 'P3地区代码
Private m_szAreaName As String 'P4地区名称
Private m_szAnnotation As String 'P5注释
Private m_nFlgProvince As Integer '省内或省外
'**************************************************
'Member Code:F1
'Brief Description:获得该地区的所有路段
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/14
'**************************************************
Public Function GetAllSection() As String()
    Dim szaSection() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute("SELECT section_id,section_name FROM section_info WHERE area_code='" & m_szAreaCode & "'")
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    ReDim szaSection(1 To nCount, 1 To 2) As String
    For i = 1 To nCount
        szaSection(i, 1) = FormatDbValue(rsTemp!section_id)
        szaSection(i, 2) = FormatDbValue(rsTemp!section_name)
        rsTemp.MoveNext
    Next
    GetAllSection = szaSection
    End If
    If Not rsTemp Is Nothing Then Set rsTemp = Nothing
    If Not oDb Is Nothing Then Set oDb = Nothing
End Function

'**************************************************
'Member Code:F2
'Brief Description:获得该地区的所有站点
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/14
'**************************************************
Public Function GetAllStation() As String()
    Dim szaStation() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute("SELECT *  FROM station_info WHERE area_code='" & m_szAreaCode & "'")
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
        ReDim szaStation(1 To nCount, 1 To 3) As String
        For i = 1 To nCount
            szaStation(i, 1) = FormatDbValue(rsTemp!station_id)
            szaStation(i, 2) = FormatDbValue(rsTemp!station_name)
            szaStation(i, 3) = FormatDbValue(rsTemp!station_input_code)
            rsTemp.MoveNext
        Next
        GetAllStation = szaStation
    End If
    If Not rsTemp Is Nothing Then Set rsTemp = Nothing
    If Not oDb Is Nothing Then Set oDb = Nothing
End Function

'P1
Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
End Property

'**************************************************
'Member Code:P1
'Brief Description:活动用户
'**************************************************
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
'**************************************************
'Member Code:P2
'Brief Description:对象状态
'**************************************************
Public Property Get ObjStatus() As EObjectStatus
    ObjStatus = m_nObjectStatus
End Property
'**************************************************
'Member Code:P3
'Brief Description:地区代码
'**************************************************
Public Property Let AreaCode(ByVal vData As String)
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    AssertActiveUserValid m_oActiveUser, ERR_Area
    '查询是否是新增状态
    If m_nObjectStatus <> ST_AddObj Then RaiseError ERR_AreaNotAddStatus
    '输入的主键是否为空
    If vData = "" Then RaiseError ERR_AreaIDNotNull
    '查询该对象是否已有记录
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT * FROM area_code WHERE area_code='" & RTrim(vData) & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 1 Then RaiseError ERR_AreaExistNotAdd
    m_szAreaCode = RTrim(vData)
    If Not rsTemp Is Nothing Then Set rsTemp = Nothing
    If Not oDb Is Nothing Then Set oDb = Nothing
End Property
'**************************************************
'Member Code:P6
'Brief Description:标示省内外 0----'市内  ,1-----市外，2--省外
'**************************************************
Public Property Get FlgProvince() As String
    AssertObjIsValid
    FlgProvince = m_nFlgProvince
End Property

Public Property Let FlgProvince(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_nFlgProvince = vData
End Property


Public Property Get AreaCode() As String
    AssertObjIsValid
    AreaCode = m_szAreaCode
End Property
'**************************************************
'Member Code:P4
'Brief Description:地区名称
'**************************************************
Public Property Let AreaName(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szAreaName = vData
End Property
Public Property Get AreaName() As String
    AreaName = m_szAreaName
End Property
'**************************************************
'Member Code:P5
'Brief Description:地区注释
'**************************************************
Public Property Let Annotation(ByVal vData As String)
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
    m_szAnnotation = vData
End Property

Public Property Get Annotation() As String
    Annotation = m_szAnnotation
End Property

'**************************************************
'Member Code:S1
'Brief Description:初始化对象
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/14
'**************************************************
Public Sub Init(poAUser As ActiveUser)
    Set m_oActiveUser = poAUser
End Sub

'**************************************************
'Member Code:S2
'Brief Description:类初始化
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub Class_Initialize()
    m_nObjectStatus = ST_NotAvailable
End Sub

'**************************************************
'Member Code:S3
'Brief Description:测试对象是否有效
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then RaiseError ERR_AreaNotAvailable
End Sub

'**************************************************
'Member Code:S4
'Brief Description:测试对象是否有效
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub AssertObjIsValid()
    AssertActiveUserValid m_oActiveUser, ERR_Area
    AssertStatusAvailable
End Sub
'**************************************************
'Member Code:S5
'Brief Description:标示地区对象
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Public Sub Identify(AreaCode As String)
    AssertActiveUserValid m_oActiveUser, ERR_Area
    m_szAreaCode = RTrim(AreaCode)
    RefreshMemoryInfo
    m_nObjectStatus = ST_NormalObj
End Sub

'**************************************************
'Member Code:S6
'Brief Description:新增地区
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Public Sub AddNew()
    AssertActiveUserValid m_oActiveUser, ERR_Area
    m_nObjectStatus = ST_AddObj
    AssertHaveRight m_oActiveUser, RIGHT_AreaManagement '校对用户是否有权限管理地区对象
End Sub
'**************************************************
'Member Code:S7
'Brief Description:删除地区
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Public Sub Delete()
    AssertObjIsValid
    DeleteObject
    m_nObjectStatus = ST_NotAvailable
    AssertHaveRight m_oActiveUser, RIGHT_AreaManagement '校对用户是否有权限管理地区对象
End Sub
'**************************************************
'Member Code:S8
'Brief Description:修改地区
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Public Sub Update()
    AssertHaveRight m_oActiveUser, RIGHT_AreaManagement '校对用户是否有权限管理地区对象
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

'**************************************************
'Member Code:S9
'Brief Description:从数据库中将地区属性读入内存
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub RefreshMemoryInfo()
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsArea As Recordset
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "SELECT * FROM area_code WHERE area_code='" & m_szAreaCode & "'"
    Set rsArea = oDb.Execute(szSql)
    If rsArea.RecordCount = 0 Then RaiseError ERR_AreaNotExist
    m_szAreaName = rsArea!area_name
    m_szAnnotation = FormatDbValue(rsArea!Annotation)
    m_nFlgProvince = FormatDbValue(rsArea!province_in_out)
    Set oDb = Nothing
    m_nObjectStatus = ST_NormalObj
End Sub

'**************************************************
'Member Code:S10
'Brief Description:向数据库中新增地区
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub AddObject()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    
    '----------------
    '新增地区代码不能未空
    If Trim(m_szAreaCode) = "" Then RaiseError ERR_AreaIDNotNull
    
    szSql = "INSERT area_code (area_code,area_name,annotation,province_in_out) " _
    & " VALUES ('" & m_szAreaCode & "','" & m_szAreaName & "','" & m_szAnnotation & "'," & m_nFlgProvince & ")"
    oDb.Execute szSql
    If Not oDb Is Nothing Then Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertAddObjectError ERR_Area, oDb
End Sub
'**************************************************
'Member Code:S11
'Brief Description:丛数据库中删除该地区
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub DeleteObject()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "DELETE area_code WHERE area_code='" & m_szAreaCode & "'"
    oDb.Execute szSql
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertDeleteObjectError ERR_Area, oDb
End Sub

'**************************************************
'Member Code:S11
'Brief Description:丛数据库中修改地区属性
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub UpdateToDB()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szSql = "UPDATE area_code SET area_name='" & m_szAreaName & "',province_in_out=" & m_nFlgProvince & " ," _
    & "Annotation='" & m_szAnnotation & "' WHERE area_code='" & m_szAreaCode & "'"
    oDb.Execute szSql
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertUpdateObjectError ERR_Area, oDb
End Sub



