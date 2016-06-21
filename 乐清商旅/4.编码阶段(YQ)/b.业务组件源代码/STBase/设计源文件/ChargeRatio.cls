VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChargeRatio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**********************************************************
'* Source File Name:ChargeRatio.cls
'* Project Name:STBase.vbp
'* Engineer:
'* Data Generated:
'* Last Revision Date:
'* Brief Description:费率对象
'* Relational Document:
'**********************************************************
Option Explicit
'出错
Public Enum EErrChargeRatio
    ERR_ChargeRatioExistNotAdd = ERR_ChargeRatio + ERR_AddDuplicate '55601已有该费率不能新增
    ERR_ChargeRatioNotAvailable = ERR_Route + ERR_NotAvailable '55612费率对象处于无状态
    ERR_ChargeRatioNotExist = ERR_ChargeRatio + 21 '55621给定的条件不能确定费率
    ERR_ChargeRatioItemNotNull = ERR_ChargeRatio + 22 '55622新增费率的三项中有空项
End Enum

'权限
Public Enum ERightChargeRatio
    RIGHT_ChargeRatioManagement = ERR_ChargeRatio + cnMidRightBegin + cnMidRightStep * 1 '55706费率管理
End Enum

'费率扩展接口
Public Type TChargeRatioEx
       sgBaseCarriageRatio As Double
       sgRoadConstructFundRatio As Double
       szAnnotation As String
       szAreaCode As String
       szAreaName As String
       szRoadLevel As String
       szSeatType As String
       szRoadLevelName As String
       szVehicleModel As String
       szVehicleModelName As String
End Type
Private m_oActiveUser As ActiveUser 'P1所有的类都有的私有类变量
'F2根据不同的条件查询出不同的费率数组
Public Function GetAllChargeRatio(Optional VehicleModel As String = "", Optional Area As String = "", Optional RoadLevel As String = "", Optional szSeatType As String) As TChargeRatioEx()
    Dim taChargeRatio() As TChargeRatioEx
    Dim oDb As New RTConnection
    Dim szTempSql As String
    Dim szWhere As String
    Dim i As Integer, nCount As Integer
    Dim rsTemp As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_ChargeRatio
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    If VehicleModel <> "" Then
        szWhere = " AND tbc.vehicle_type_code='" & VehicleModel & "'"
    End If
    If Area <> "" Then
        szWhere = szWhere & " AND tbc.area_code='" & Area & "'"
    End If
    If RoadLevel <> "" Then
        szWhere = szWhere & " AND tbc.road_rank_code='" & RoadLevel & "'"
    End If
    If szSeatType <> "" Then
      szWhere = szWhere & " AND tbc.seat_type_id='" & szSeatType & "'"
    End If
    
    szTempSql = "SELECT tbc.*,tba.area_name,tbr.road_rank_name,tbv.vehicle_type_short_name FROM charge_ratio_info tbc,road_level_code tbr,area_code tba,vehicle_type_code tbv WHERE " _
    & " tbc.area_code=tba.area_code AND tbr.road_rank_code=tbc.road_rank_code AND tbc.vehicle_type_code =tbv.vehicle_type_code" & szWhere
    Set rsTemp = oDb.Execute(szTempSql)
    nCount = rsTemp.RecordCount
    
    If nCount <> 0 Then
        ReDim taChargeRatio(1 To nCount) As TChargeRatioEx
        For i = 1 To nCount
            taChargeRatio(i).szVehicleModel = FormatDbValue(rsTemp!vehicle_type_code)
            taChargeRatio(i).szRoadLevel = FormatDbValue(rsTemp!road_rank_code)
            taChargeRatio(i).szAreaCode = FormatDbValue(rsTemp!area_code)
            taChargeRatio(i).sgBaseCarriageRatio = FormatDbValue(rsTemp!base_transit_ratio)
            taChargeRatio(i).sgRoadConstructFundRatio = FormatDbValue(rsTemp!road_construct_fund_ratio)
            taChargeRatio(i).szAreaName = FormatDbValue(rsTemp!area_name)
            taChargeRatio(i).szSeatType = FormatDbValue(rsTemp!seat_type_id)
            taChargeRatio(i).szRoadLevelName = FormatDbValue(rsTemp!road_rank_name)
            taChargeRatio(i).szVehicleModelName = FormatDbValue(rsTemp!vehicle_type_short_name)
            taChargeRatio(i).szAnnotation = FormatDbValue(rsTemp!Annotation)
            rsTemp.MoveNext
        Next
        GetAllChargeRatio = taChargeRatio
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
    Set m_oActiveUser = vNewValue
End Property
'S1
Public Sub Init(poAUser As ActiveUser)
    Set m_oActiveUser = poAUser
End Sub
'S2新增费率
Public Sub AddChargeRatio(ChargeRatio As TChargeRatio)
    Dim oDb As New RTConnection
    Dim szTempSql As String
    Dim szSqlField As String
    Dim szSqlContext As String
    Dim szlog As String
    AssertHaveRight m_oActiveUser, RIGHT_ChargeRatioManagement
    AssertActiveUserValid m_oActiveUser, ERR_ChargeRatio
    If Trim(ChargeRatio.szAreaCode) = "" Or Trim(ChargeRatio.szRoadLevel = "") Or Trim(ChargeRatio.szAreaCode) = "" Then RaiseError ERR_ChargeRatioItemNotNull
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szTempSql = "INSERT charge_ratio_info ("
    szSqlField = szSqlField & "area_code,"
    szSqlField = szSqlField & "road_rank_code,"
    szSqlField = szSqlField & "vehicle_type_code,"
    szSqlField = szSqlField & "base_transit_ratio,"
    szSqlField = szSqlField & "seat_type_id,"
    szSqlField = szSqlField & "road_construct_fund_ratio,"
    szSqlField = szSqlField & "annotation)"
    szSqlContext = " VALUES ('" & ChargeRatio.szAreaCode & "',"
    szSqlContext = szSqlContext & " '" & ChargeRatio.szRoadLevel & "',"
    szSqlContext = szSqlContext & " '" & ChargeRatio.szVehicleModel & "',"
    szSqlContext = szSqlContext & " " & ChargeRatio.sgBaseCarriageRatio & ","
    szSqlContext = szSqlContext & " '" & ChargeRatio.szSeatType & "',"
    szSqlContext = szSqlContext & " " & ChargeRatio.sgRoadConstructFundRatio & ","
    szSqlContext = szSqlContext & " '" & ChargeRatio.szAnnotation & "')"
    szTempSql = szTempSql & szSqlField & szSqlContext
    oDb.Execute (szTempSql)
    Set oDb = Nothing
    
    szlog = szlog & "座位类型[" & ChargeRatio.szSeatType & "]"
    szlog = szlog & "地区[" & ChargeRatio.szAreaCode & "费率为" & ChargeRatio.szAnnotation
    szlog = GetUnicodeBySize(szlog, 255)
    WriteOperateLog m_oActiveUser, RIGHT_ChargeRatioManagement, szlog
End Sub
'S3删除费率
Public Sub DeleteChargeRatio(ChargeRatio As TChargeRatio)
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim szlog As String
    
    AssertActiveUserValid m_oActiveUser, ERR_ChargeRatio
    AssertHaveRight m_oActiveUser, RIGHT_ChargeRatioManagement
    
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    If ChargeRatio.szSeatType <> "" Then
      szSql = "DELETE charge_ratio_info WHERE area_code='" & ChargeRatio.szAreaCode & "' AND road_rank_code='" & ChargeRatio.szRoadLevel & "' AND vehicle_type_code='" & ChargeRatio.szVehicleModel & "' and seat_type_id='" & ChargeRatio.szSeatType & "'"
    Else
      szSql = "DELETE charge_ratio_info WHERE area_code='" & ChargeRatio.szAreaCode & "' AND road_rank_code='" & ChargeRatio.szRoadLevel & "' AND vehicle_type_code='" & ChargeRatio.szVehicleModel & "'"
    End If
    oDb.Execute (szSql)
    If ChargeRatio.szSeatType <> "" Then
        szlog = Now & "地区[" & " & ChargeRatio.szAreaCode & " & "]公路等级[ " & " & ChargeRatio.szRoadLevel & " & " ]车型[" & " &  ChargeRatio.szVehicleModel & " & "]座位类型[" & " & ChargeRatio.szSeatType & " & "费率被删除"
        szlog = GetUnicodeBySize(szlog, 255)
        WriteOperateLog m_oActiveUser, RIGHT_ChargeRatioManagement, szlog
    Else
        szlog = Now & "地区[" & " & ChargeRatio.szAreaCode & " & "]公路等级[ " & " & ChargeRatio.szRoadLevel & " & " ]车型[" & " &  ChargeRatio.szVehicleModel & " & "]全部座位类型" & "费率被删除"
        szlog = GetUnicodeBySize(szlog, 255)
        WriteOperateLog m_oActiveUser, RIGHT_ChargeRatioManagement, szlog
    End If
    Set oDb = Nothing
End Sub
'S4编辑费率
Public Sub ModifyChargeRatio(ChargeRatio As TChargeRatio)
    Dim oDb As New RTConnection
    Dim szTempSql As String
    Dim szSqlField As String
    Dim lEdit As Long
    Dim szSqlWhere As String
    Dim szSqlWhere2 As String
    
    Dim szlog As String
    AssertHaveRight m_oActiveUser, RIGHT_ChargeRatioManagement
    AssertActiveUserValid m_oActiveUser, ERR_ChargeRatio
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    szTempSql = "UPDATE charge_ratio_info SET "
    szSqlField = "base_transit_ratio=" & ChargeRatio.sgBaseCarriageRatio & ","
    '当 ChargeRatio.szSeatType= "" 更新所有座位类型
    If ChargeRatio.szSeatType <> "" Then
          szlog = szlog & "座位类型[" & ChargeRatio.szSeatType & "]"
         szSqlField = szSqlField & "seat_type_id ='" & ChargeRatio.szSeatType & "',"
         szSqlWhere2 = " AND seat_type_id ='" & ChargeRatio.szSeatType & "'"
    Else
        szlog = szlog & "所有座位类型"
    End If
    
    
    If ChargeRatio.szAnnotation <> "" Then
        
        szSqlField = szSqlField & "annotation ='" & ChargeRatio.szAnnotation & "',"
    
    End If
         szSqlField = szSqlField & "road_construct_fund_ratio=" & ChargeRatio.sgRoadConstructFundRatio & ""
         szSqlWhere = " WHERE area_code='" & ChargeRatio.szAreaCode & "'  AND road_rank_code='" & ChargeRatio.szRoadLevel & "' AND vehicle_type_code='" & ChargeRatio.szVehicleModel & "'"
         szTempSql = szTempSql & szSqlField & szSqlWhere & szSqlWhere2
    oDb.Execute szTempSql, lEdit
    
    szlog = szlog & "地区[" & ChargeRatio.szAreaCode & "费率为" & ChargeRatio.szAnnotation
    If lEdit <> 0 Then
       szlog = GetUnicodeBySize(szlog, 255)
       WriteOperateLog m_oActiveUser, RIGHT_ChargeRatioManagement, szlog
    End If
    
    If lEdit = 0 Then
        AddChargeRatio ChargeRatio
    End If
    Set oDb = Nothing
End Sub
