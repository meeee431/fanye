VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BackSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'**********************************************************
'* Source File Name:BackSection.cls
'* Project Name:STSettle.vbp
'* Engineer:
'* Data Generated:2004/03/12
'* Last Revision Date:2004/03/12
'* Brief Description:回程路段对象
'* Relational Document:
'**********************************************************
Option Explicit
'出错
Public Enum EErrSection
    ERR_SectionExist = ERR_Section + ERR_AddDuplicate '16401 路段已存在不能新增
    ERR_SectionNoDelete = ERR_Section + ERR_DeleteChildExist '16407 路段已安排线路不能删除
    ERR_SectionNotRoadLevelOrStation = ERR_Section + ERR_EditChildExist '16405 路段无该公路等级或站点
    ERR_SectionNotAvailable = ERR_Section + ERR_NotAvailable '16412 路段对象处于无状态
    ERR_SectionNotAddStatus = ERR_Section + ERR_NotAddObj '16416 路段对象不在新增状态
    ERR_SectionNotExist = ERR_Section + 21 '16421 路段无存在
    ERR_SectionIDNotNull = ERR_Section + 22 '16422 路段代码不能为空
    ERR_SectionStartNoEnd = ERR_Section + 23 '16423 路段起点站不能同终点站
    
    
End Enum

'权限
Public Enum ERightSection
    RIGHT_SectionManagement = ERR_Section + cnMidRightBegin + cnMidRightStep * 1 '16506 路段管理
End Enum

Private m_oActiveUser As ActiveUser 'P1所有的类都有的私有类变量
Private m_nObjectStatus As EObjectStatus 'P2所有的实体类都有的私有类变量

'Private pszSectionID As String

Public Sub Init(poAUser As ActiveUser)
    Set m_oActiveUser = poAUser
    'AssertHaveRight m_oActiveUser, RIGHT_RouteManagement
End Sub

'S2类刚建立时处于不可用状态,此时类不能使用
Private Sub Class_Initialize()
    m_nObjectStatus = ST_NotAvailable
End Sub

'S3测试对象的状态是否有效（无效状态[ST_NotAvailable]和新增状态[ST_AddObj]则出错）
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then RaiseError ERR_RouteNotAvailable
End Sub
'S4
Private Sub AssertObjIsValid() '测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
    AssertActiveUserValid m_oActiveUser, ERR_Route
    AssertStatusAvailable
End Sub
'F1
Public Function GetAllRoute(pszSectionID As String) As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    Dim szSql As String
    AssertObjIsValid
    szSql = "SELECT DISTINCT route_id FROM back_route_section_lst  WHERE section_id='" & pszSectionID & "'"
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    ReDim szaTemp(1 To nCount) As String
    For i = 1 To nCount
        szaTemp(i) = FormatDbValue(rsTemp!route_id)
        rsTemp.MoveNext
    Next
    GetAllRoute = szaTemp
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

'F3
Public Function GetAllRouteName(pszSectionID As String) As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    Dim szSql As String
    AssertObjIsValid
    szSql = "SELECT  route_name FROM back_route_section_lst tbc ,back_route_info tbr  WHERE section_id='" & pszSectionID & "' AND tbr.route_id =tbc.route_id"
    
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
    ReDim szaTemp(1 To nCount) As String
    For i = 1 To nCount
        szaTemp(i) = FormatDbValue(rsTemp!route_name)
        rsTemp.MoveNext
    Next
    GetAllRouteName = szaTemp
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function


'因为里程数的修改将修改使用该路段的所有线路
Public Sub UpdateRouteMileage(pszSectionID As String, psgMileage As Single, psgOldSectionMileage As Single)
    On Error GoTo ErrHandle

    Dim oDb As New RTConnection
    Dim szTempSql As String
    Dim szSqlField As String
    Dim szSqlWhere As String
    Dim nCount As Integer
    Dim i As Integer
    Dim rsTemp As Recordset
    Dim sgChangeMileage As Double
    Dim szlog As String


    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)

    '因为里程数的修改将修改使用该路段的所有线路和票价表中的站点里程数
    sgChangeMileage = psgMileage - psgOldSectionMileage
    If sgChangeMileage = 0 Then Exit Sub
'    szaRouteID = Me.GetAllRoute
    szTempSql = "SELECT DISTINCT sell_station_id , route_id FROM back_route_section_lst  WHERE section_id='" & pszSectionID & "'"
    Set rsTemp = oDb.Execute(szTempSql)
    nCount = rsTemp.RecordCount
    For i = 1 To nCount
        '修改线路路段中的里程数back_route_section_lst
        szTempSql = "UPDATE back_route_section_lst " _
            & " SET end_station_mileage=end_station_mileage+ " & Str(sgChangeMileage) _
            & " WHERE route_id='" & FormatDbValue(rsTemp!route_id) & "' " _
            & " AND section_serial>=( " _
            & " SELECT min(section_serial) FROM back_route_section_lst " _
            & " WHERE section_id='" & pszSectionID & "' AND route_id='" & FormatDbValue(rsTemp!route_id) & "' AND sell_station_id = '" & FormatDbValue(rsTemp!sell_station_id) & "'" _
            & " ) AND sell_station_id = '" & FormatDbValue(rsTemp!sell_station_id) & "'"
        oDb.Execute szTempSql
        '修改线路的里程数route_info
        szTempSql = "UPDATE back_route_info SET mileage=mileage+" & Str(sgChangeMileage) & " WHERE route_id='" & FormatDbValue(rsTemp!route_id) & "'"
        oDb.Execute szTempSql
        rsTemp.MoveNext
    Next i
    Set rsTemp = Nothing
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    AssertUpdateObjectError ERR_Section, oDb

End Sub



