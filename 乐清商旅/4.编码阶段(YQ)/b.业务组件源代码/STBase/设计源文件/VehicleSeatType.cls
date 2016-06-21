VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VehicleSeatType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum ERightVehiclSeatType
    RIGHT_VechileSeatTypeManagement = ERR_VehicleSeatType + cnMidRightBegin + cnMidRightStep * 1 '57306公路等级管理
End Enum
Public Enum EErrVehiclSeatType
    ERR_VechileSeatTypeINFOExistNotAddNew = ERR_VehicleSeatType + ERR_AddDuplicate '信息已存在不能重复增加
    ERR_VechileSeatTypeINFONotAvailable = ERR_VehicleSeatType + ERR_NotAvailable '车辆对象处于无状态
    ERR_NoVehicleSeatTypeInfo = ERR_VehicleSeatType + 21 '数据库车辆座位信息不存在
End Enum

Private m_nObjectStatus As EObjectStatus
Private m_oActiveUser As New ActiveUser
Private m_szVehicleID As String
Private m_szEndSeatNO As String
Private m_szStartSeatNO As String
Private m_szSeatTypeID As String
Private m_szSerial As Integer
Public Property Let serial(ByVal vData As Long)
    m_szSerial = vData
End Property
Public Property Get serial() As Long
    serial = m_szSerial
End Property
Public Property Let SeatTypeID(ByVal vData As String)
   m_szSeatTypeID = vData
End Property
Public Property Get SeatTypeID() As String
    SeatTypeID = m_szSeatTypeID
End Property
Public Property Let startSeatNo(ByVal vData As String)
    m_szStartSeatNO = vData
End Property
Public Property Get startSeatNo() As String
    startSeatNo = m_szStartSeatNO
End Property
Public Property Let EndSeatNO(ByVal vData As String)
    m_szEndSeatNO = vData
End Property
Public Property Get EndSeatNO() As String
    EndSeatNO = m_szEndSeatNO
End Property
Public Property Let VehicleID(ByVal vData As String)
    m_szVehicleID = vData
End Property
Public Property Get VehicleID() As String
    VehicleID = m_szVehicleID
End Property
Public Property Let ObjStatus(ByVal vData As EObjectStatus)
    m_nObjectStatus = vData
End Property
Public Property Get ObjStatus() As EObjectStatus
    ObjStatus = m_nObjectStatus
End Property
Public Function GetAll(Optional szVehicleID As String) As String()
    Dim oDb As New SNConnection
    Dim szTempSql As String
    Dim szSql As String
    Dim szSqlContext As String
    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim nCount As Integer
    Dim szTemp() As String
     oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
      If szVehicleID <> "" Then
     szSql = "Select distinct tve.* ,sti.seat_type_name,vtc.vehicle_type_name from vehicle_seat_type_info tve ,seat_type_info sti ,transport_vehicle_info trv , vehicle_type_code vtc where tve.vehicle_id='" & szVehicleID & " 'and " _
              & " sti.seat_type_id=tve.Seat_type_id and " _
              & " trv.vehicle_id=tve.vehicle_id  and vtc.vehicle_type_code=trv.vehicle_type_code "
     Else
     szSql = "Select distinct tve.* ,sti.seat_type_name,vtc.vehicle_type_name from vehicle_seat_type_info tve ,seat_type_info sti ,transport_vehicle_info trv , vehicle_type_code vtc where    " _
              & " sti.seat_type_id=tve.Seat_type_id and " _
              & " trv.vehicle_id=tve.vehicle_id  and vtc.vehicle_type_code=trv.vehicle_type_code "
     End If
     
     Set rsTemp = oDb.Execute(szSql)
     nCount = rsTemp.RecordCount
     If nCount <> 0 Then
     ReDim szTemp(1 To nCount, 1 To 7)
        For i = 1 To nCount
            szTemp(i, 1) = rsTemp!serial_no
            szTemp(i, 2) = FormatDbValue(rsTemp!vehicle_type_name)
            szTemp(i, 3) = rsTemp!vehicle_id
            szTemp(i, 4) = rsTemp!seat_type_id
            szTemp(i, 5) = rsTemp!seat_type_name
            szTemp(i, 6) = FormatDbValue(rsTemp!start_seat_no)
            szTemp(i, 7) = FormatDbValue(rsTemp!End_seat_no)
            rsTemp.MoveNext
        Next
     End If
          GetAll = szTemp
          Set rsTemp = Nothing
          Set oDb = Nothing
End Function
Public Sub Init(poUser As ActiveUser)
 Set m_oActiveUser = poUser
End Sub
Private Sub AssertObjIsValid() '测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
    AssertActiveUserValid m_oActiveUser, RIGHT_VechileSeatTypeManagement
    AssertStatusAvailable
End Sub
Private Sub AssertStatusAvailable()
  If m_nObjectStatus = ST_NotAvailable Then ShowError ERR_VechileSeatTypeINFONotAvailable
End Sub
Public Function Delete(szSerial As String)
    Dim oDb As New SNConnection
    Dim szTempSql As String
    Dim rsTemp As New Recordset
     On Error GoTo here
     oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
     oDb.BeginTrans
      szTempSql = "Delete vehicle_seat_type_info where serial_no='" & szSerial & " '"
     oDb.Execute szTempSql
     oDb.CommitTrans

     Set rsTemp = Nothing
     Set oDb = Nothing
     Exit Function
here:
    oDb.RollbackTrans
     Set rsTemp = Nothing
     Set oDb = Nothing
     
End Function
Public Function AddNew(szTemp() As String, szVehicleID As String) As Integer
    Dim oDb As New SNConnection
    Dim szTempSql As String
    Dim szSqlField As String
    Dim szSqlContext As String
    Dim rsTemp As New Recordset
    Dim nCount As Integer
    Dim i As Integer
    Dim nSearial As Integer
    On Error GoTo here
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    nCount = ArrayLength(szTemp)
    If nCount = 0 Then Exit Function
    oDb.BeginTrans
       szTempSql = "Select serial_no from vehicle_seat_type_info where vehicle_id='" & szVehicleID & " 'and ((start_seat_no<= '" & szTemp(2) & "' and end_seat_no >= '" & szTemp(2) & "')  or " _
               & "(start_seat_no <= '" & szTemp(3) & "' and end_seat_no >= '" & szTemp(3) & "') )"
      Set rsTemp = oDb.Execute(szTempSql)
      If rsTemp.RecordCount <> 0 Then
        oDb.RollbackTrans
        ShowError ERR_VechileSeatTypeINFOExistNotAddNew
       End If
         szTempSql = ""
         szSqlField = "INSERT vehicle_seat_type_info ("
         szSqlField = szSqlField & " seat_type_id,"
         szSqlField = szSqlField & "vehicle_id,"
         szSqlField = szSqlField & "start_seat_no,"
         szSqlField = szSqlField & "end_seat_no)"
         szSqlContext = " VALUES ('" & szTemp(1) & "',"
         szSqlContext = szSqlContext & " '" & szVehicleID & "',"
         szSqlContext = szSqlContext & " '" & szTemp(2) & "',"
         szSqlContext = szSqlContext & " '" & szTemp(3) & "')"
         szTempSql = szTempSql & szSqlField & szSqlContext
    oDb.Execute szTempSql
         szSqlContext = ""
         szTempSql = "Select serial_no from vehicle_seat_type_info where vehicle_id='" & szVehicleID & " 'and start_seat_no= '" & szTemp(2) & "' and end_seat_no = '" & szTemp(3) & "'   "
         Set rsTemp = oDb.Execute(szTempSql)
         If rsTemp.RecordCount = 0 Then nSearial = -1: GoTo here
         nSearial = rsTemp!serial_no
    AddNew = nSearial
    oDb.CommitTrans
    Set oDb = Nothing
    Set rsTemp = Nothing
    Exit Function
here:
    oDb.RollbackTrans
    AddNew = nSearial
    Set oDb = Nothing
    Set rsTemp = Nothing
    ShowError Err.Number, "数据库更新时出错，请检查网络"
End Function
Public Function UpdateToDB(szTemp() As String, szVehicleID As String)
    Dim oDb As New SNConnection
    Dim szTempSql As String
    Dim szSqlField As String
    Dim szSqlContext As String
    Dim rsTemp As New Recordset
    Dim szSqlWhere As String
    Dim nCount As Integer
    Dim i As Integer
    On Error GoTo here
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    nCount = ArrayLength(szTemp)
    If nCount = 0 Then Exit Function
    oDb.BeginTrans
   For i = 1 To nCount
      szTempSql = "Select serial_no from vehicle_seat_type_info where  serial_no='" & szTemp(i, 0) & "'"
      Set rsTemp = oDb.Execute(szTempSql)
      If rsTemp.RecordCount = 0 Then
        oDb.RollbackTrans
        ShowError ERR_NoVehicleSeatTypeInfo
      End If
         szTempSql = ""
         szSqlField = ""
         szSqlWhere = ""
         szTempSql = "UPDATE vehicle_seat_type_info SET "
         szSqlField = szSqlField & "vehicle_id='" & szVehicleID & "',"
         szSqlField = szSqlField & "seat_type_id='" & szTemp(i, 1) & "',"
         szSqlField = szSqlField & "start_seat_no='" & szTemp(i, 2) & "',"
         szSqlField = szSqlField & "end_seat_no='" & szTemp(i, 3) & "'"
         szSqlWhere = " WHERE serial_no='" & szTemp(i, 0) & "'"
         szTempSql = szTempSql & szSqlField & szSqlWhere
         oDb.Execute szTempSql
   Next
   oDb.CommitTrans
   Set oDb = Nothing
   Set rsTemp = Nothing
   Exit Function
here:
   oDb.RollbackTrans
   TreatUpdateError ERR_VehicleSeatType, oDb
   Set oDb = Nothing
   Set oDb = Nothing
   Set rsTemp = Nothing
End Function
