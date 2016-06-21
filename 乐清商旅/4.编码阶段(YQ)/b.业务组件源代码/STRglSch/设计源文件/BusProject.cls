VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BusProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**********************************************************
'* Source File Name:BusProject.cls
'* Project Name:STRglSch.vbp
'* Engineer:
'* Data Generated:
'* Last Revision Date:
'* Brief Description:车次计划
'* Relational Document:
'**********************************************************

Option Explicit

'出错
Public Enum EErrBusProject
    ERR_BusProjectAdd = ERR_BusProject + ERR_AddDuplicate '14001新增计划已存在
    ERR_BusProjectNotAddStatus = ERR_BusProject + ERR_NotAddObj '14016计划不处于新增
    ERR_BusProjectNotExist = ERR_BusProject + 21 '14021车次计划不存在
    ERR_BusProjectNotBus = ERR_BusProject + 22 '14022车次计划无运行车辆
    ERR_BsuProjectIDNotNull = ERR_BusProject + 23 '14023车次计划代码不能为空
    ERR_BusProjectExistNotAddNew = ERR_BusProject + 24 '14024该车次计划已存在，不能新增
    ERR_BusProjectNotPriceTable = ERR_BusProject + 25 '14025车次计划无票价
    ERR_BusProjectPropertyNotEdit = ERR_BusProject + 26 '14026车次对象的该属性不能修改
    ERR_ExecutePriceTableNotExist = ERR_BusProject + 27 '14027车次无执行票价
    ERR_BusProjectReplecProjectIDorName = ERR_BusProject + 28 '14028 计划代码或名重复
    ERR_alreadyExitstProject = ERR_BusProject + 29 '14029计划已按排
End Enum

'权限
Public Enum ERightBusProject
    RIGHT_BusProjectDelete = ERR_BusProject + cnMidRightBegin + cnMidRightStep * 1 '14107删除计划
    RIGHT_BusProjectAdd = ERR_BusProject + cnMidRightBegin + cnMidRightStep * 2 '14113新增计划
    RIGHT_BusProjectManagement = ERR_BusProject + cnMidRightBegin + cnMidRightStep * 4 '14125计划管理
    RIGHT_BusProjectRecover = ERR_BusProject + cnMidRightBegin + cnMidRightStep * 5 '14131恢复删除计划
    RIGHT_BusProjectPurge = ERR_BusProject + cnMidRightBegin + cnMidRightStep * 6 '14137物理删除计划
    RIGHT_BusProjectBooked = ERR_BusProject + cnMidRightBegin + cnMidRightStep * 7 '14137物理删除计划

End Enum
'
Private m_oActiveUser As ActiveUser 'P1所有的类都有的私有类变量
Private m_nObjectStatus As EObjectStatus 'P2所有的实体类都有的私有类变量
'Private m_szProjectID As String 'P3计划代码
Private m_szExecutePriceTable As String 'P4执行票价票
Private m_szAnnotation As String 'P5注释
Private m_dtLastUpdateTime As Date 'P6最后一次修改时间
Private m_dtCreateTime As Date 'P7创建时间
Private m_ePrjectStatus As EBusProjectStatus  'P8计划状态
Private m_szProjectName As String 'P9计划名称
Private m_szExecutePriceTableName As String 'P10执行票价票名
'
''**************************************************
''Member Code:F1
''Brief Description:得到该车次计划中安排了车次及车辆但没有生成票价信息的线路和相应的车型信息
''Engineer:
''Date Generated:
''Last Revision Date:
''**************************************************
Public Function GetNoPriceRouteVehicleModel(Optional ByVal RoutePriceTableID As String = "") As TRouteVehicleModel()
    AssertObjIsValid
    Dim arvmReturnData() As TRouteVehicleModel
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    Dim i As Integer, j As Integer, k As Integer, nTempEnd As Integer, nTempBegin As Integer
    Dim szLastRouteID As String
    szSql = "SELECT DISTINCT route_id,vehicle_type_code FROM " _
    & " bus_info,bus_vehicle_code,Vehicle_info WHERE " _
    & " bus_info.bus_id=bus_vehicle_code.bus_id AND " _
    & " bus_vehicle_code.vehicle_id=Vehicle_info.vehicle_id AND " _
    & " (SELECT COUNT(*) FROM  price_table_info ,price_table_lst WHERE " _
    & " price_table_info.price_table_id =price_table_lst.price_table_id AND " _
    & " price_table_lst.route_id=bus_info.route_id AND " _
    & " price_table_lst.vehicle_type_code=Vehicle_info.vehicle_type_code "
    If RoutePriceTableID <> "" Then
        szSql = szSql & " AND price_table_info.price_table_id='" & RoutePriceTableID & "' "
    End If


    szSql = szSql _
    & ")=0  ORDER BY " _
    & " route_id"


    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        ReDim arvmReturnData(1 To rsTemp.RecordCount)
        j = 0
        k = 1
        szLastRouteID = ""
        nTempBegin = 1
        For i = 1 To rsTemp.RecordCount
            If RTrim(rsTemp!route_id) <> szLastRouteID Then
                If szLastRouteID <> "" Then
                    nTempEnd = i - 1
                    ReDim arvmReturnData(j).aszVehicleModel(1 To nTempEnd - nTempBegin + 1)
                    arvmReturnData(j).szRouteID = szLastRouteID
                    rsTemp.Move -(nTempEnd - nTempBegin + 1)
                    For k = nTempBegin To nTempEnd
                        arvmReturnData(j).aszVehicleModel(k - nTempBegin + 1) = RTrim(rsTemp!vehicle_type_code)
                        rsTemp.MoveNext
                    Next
                    nTempBegin = i
                End If
                j = j + 1
                szLastRouteID = RTrim(rsTemp!route_id)
            End If
            rsTemp.MoveNext
        Next

        nTempEnd = i
        ReDim arvmReturnData(j).aszVehicleModel(1 To nTempEnd - nTempBegin + 1)
        arvmReturnData(j).szRouteID = szLastRouteID
        rsTemp.Move -(nTempEnd - nTempBegin + 1)
        For k = nTempBegin To nTempEnd
            arvmReturnData(j).aszVehicleModel(k - nTempBegin + 1) = RTrim(rsTemp!vehicle_type_code)
            rsTemp.MoveNext
        Next
        ReDim Preserve arvmReturnData(1 To j)
    End If
    GetNoPriceRouteVehicleModel = arvmReturnData
End Function

''**************************************************
''Member Code:F2
''Brief Description:得到该车次计划中的安排了车次及车辆的线路和相应的车型信息。
''Engineer:
''Date Generated:
''Last Revision Date:
''**************************************************
Public Function GetRouteVehicleModel() As TRouteVehicleModel()
    AssertObjIsValid
    Dim arvmReturnData() As TRouteVehicleModel
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    Dim i As Integer, j As Integer, k As Integer, nTempEnd As Integer, nTempBegin As Integer
    Dim szLastRouteID As String
    szSql = "SELECT  DISTINCT route_id,vehicle_type_code FROM " _
    & " bus_info,bus_vehicle_code,Vehicle_info WHERE " _
    & " bus_info.bus_id=bus_vehicle_code.bus_id AND " _
    & " bus_vehicle_code.vehicle_id=Vehicle_info.vehicle_id " _
    & " ORDER BY route_id "

    oDb.ConnectionString = GetConnectionStr(cszSystemMan)

    Set rsTemp = oDb.Execute(szSql)

    If rsTemp.RecordCount > 0 Then
        ReDim arvmReturnData(1 To rsTemp.RecordCount)
        j = 0
        k = 1
        szLastRouteID = ""
        nTempBegin = 1
        For i = 1 To rsTemp.RecordCount
            If RTrim(rsTemp!route_id) <> szLastRouteID Then
                If szLastRouteID <> "" Then
                    nTempEnd = i - 1
                    ReDim arvmReturnData(j).aszVehicleModel(1 To nTempEnd - nTempBegin + 1)
                    arvmReturnData(j).szRouteID = szLastRouteID
                    rsTemp.Move -(nTempEnd - nTempBegin + 1)
                    For k = nTempBegin To nTempEnd
                        arvmReturnData(j).aszVehicleModel(k - nTempBegin + 1) = RTrim(rsTemp!vehicle_type_code)
                        rsTemp.MoveNext
                    Next
                    nTempBegin = i
                End If
                j = j + 1
                szLastRouteID = RTrim(rsTemp!route_id)
            End If
            rsTemp.MoveNext
        Next

        nTempEnd = i - 1
        ReDim arvmReturnData(j).aszVehicleModel(1 To nTempEnd - nTempBegin + 1)
        arvmReturnData(j).szRouteID = szLastRouteID
        rsTemp.Move -(nTempEnd - nTempBegin + 1)
        'rsTemp.Move -(nTempEnd - nTempBegin)
        For k = nTempBegin To nTempEnd
            'arvmReturnData(j).aszVehicleModel(k - nTempBegin + 1) = RTrim(rsTemp!vehicle_type_code)
            arvmReturnData(j).aszVehicleModel(k - nTempBegin + 1) = RTrim(rsTemp!vehicle_type_code)

            rsTemp.MoveNext
        Next
        ReDim Preserve arvmReturnData(1 To j)
    End If

    GetRouteVehicleModel = arvmReturnData

End Function
''**************************************************
''Member Code:F3
''Brief Description:得到计划的所有车次
''Engineer:陈峰
''Date Generated:2001/7/11
''Last Revision Date:2001/7/11
''**************************************************
Public Function GetAllBus(Optional LikeBusID As String, Optional RouteID As String, Optional szBusStationID As String, Optional bCanSale As Boolean, Optional szSellStationID As String, Optional EqualBusID As String) As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim szWhere As String
    Dim szSql As String
    Dim i As Long
    Dim szFild As String
    Dim szWhere2 As String
    Dim rsTemp As Recordset
    
    
    '''''''''
    Dim rstempBusVehicle As New Recordset
    Dim j As Integer
    Dim szDate As String
    Dim szWhereBusVehicle As String
    '''''''''
    AssertObjIsValid
    On Error GoTo ErrHere
    
    If LikeBusID <> "" Then
        szWhere = " AND tbb.bus_id LIKE '" & LikeBusID & "%' "
        szWhereBusVehicle = " AND bv.bus_id LIKE '" & LikeBusID & "%' "
    End If

    If RouteID <> "" Then
        szWhere = szWhere & " AND tbb.route_id='" & RouteID & "'"
    End If
    
    If EqualBusID <> "" Then
        szWhere = szWhere & " AND tbb.bus_id = '" & EqualBusID & "'"
    End If
    
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    If szBusStationID = "" Or szBusStationID = "(全部)" Then
    
    '    szSql = "SELECT tbb.*,tbr.route_name ,bti.bus_type_name,ci.check_gate_name , tbr.end_station_id " _
    '            & " FROM bus_info tbb,route_info tbr, " _
    '            & " Bus_type_code bti,Checkgate_info ci " _
    '            & " WHERE tbb.project_id='" & m_szProjectID & "'  " _
    '            & " AND tbb.route_id=tbr.route_id AND bti.bus_type=tbb.bus_type and tbb.check_gate_id=ci.check_gate_id "
    
        szSql = "SELECT tbb.*,tbr.route_name ,bti.bus_type_name,ci.check_gate_name , tbr.end_station_id " _
        & " FROM bus_info tbb,bus_allot_info tba,route_info tbr, " _
        & " Bus_type_code bti,Checkgate_info ci " _
        & " WHERE  tbb.route_id=tbr.route_id AND bti.bus_type=tbb.bus_type and tbb.check_gate_id=ci.check_gate_id AND tbb.bus_id=tba.bus_id AND tbb.check_gate_id=tba.check_gate_id "
        szSql = szSql & szWhere
    Else
        If bCanSale = True Then
            szWhere = szWhere & " and g.max_sale_quantity <> 0"
        End If
        '   szSql = "SELECT tbb.*,tbr.route_name ,bti.bus_type_name,ci.check_gate_name, tbr.end_station_id" _
        '            & " FROM bus_info tbb,route_info tbr, " _
        '            & " Bus_type_code bti,Checkgate_info ci ,Bus_station_info g,station_info s" _
        '            & " WHERE tbb.project_id='" & m_szProjectID & "' AND g.project_id =tbb.project_id and g.station_id=s.station_id and g.bus_id=tbb.bus_id  " _
        '            & " AND tbb.route_id=tbr.route_id AND bti.bus_type=tbb.bus_type and tbb.check_gate_id=ci.check_gate_id "
        '  szSql = szSql & szWhere _
        '             & " AND  (ltrim(s.station_name)='" & Trim(szBusStationID) & "'  " _
        '             & " OR ltrim(s.station_input_code)='" & Trim(szBusStationID) & "' OR ltrim(s.station_id)='" & Trim(szBusStationID) & "')  order by tbb.bus_id"
        
        
        
        szSql = "SELECT tbb.*,tbr.route_name ,bti.bus_type_name,ci.check_gate_name, tbr.end_station_id" _
        & " FROM bus_info tbb,bus_allot_info tba,route_info tbr, " _
        & " Bus_type_code bti,Checkgate_info ci ,Bus_station_info g,station_info s" _
        & " WHERE  g.station_id=s.station_id and g.bus_id=tbb.bus_id  " _
        & " AND tbb.route_id=tbr.route_id AND bti.bus_type=tbb.bus_type and tbb.check_gate_id=ci.check_gate_id AND tbb.bus_id=tba.bus_id AND tbb.check_gate_id=tba.check_gate_id "
        szSql = szSql & szWhere _
        & " AND  (ltrim(s.station_name)='" & Trim(szBusStationID) & "'  " _
        & " OR ltrim(s.station_input_code)='" & Trim(szBusStationID) & "' OR ltrim(s.station_id)='" & Trim(szBusStationID) & "')"
    
    End If
    
    If szSellStationID <> "" Then
        szSql = szSql & " AND tba.sell_station_id=" & TransFieldValueToString(szSellStationID)
    End If
    
    szSql = szSql & "  order by tbb.bus_id"
    
    ' , tbb.bus_start_time
    Set rsTemp = oDb.Execute(szSql)
    
    '''''''''''''
    Dim nSerial As String
    Dim lerror As Long
    Dim dtBeginDate As Date
    '取车次车辆
    '    dtBeginDate = SelfGetExecuteBusProject(Now, lerror).dtBeginDate
    szSql = "SELECT bv.*, t.status ,t.license_tag_no " _
        & " FROM bus_vehicle_code bv, Vehicle_info t " _
        & " WHERE bv.vehicle_id  =t.vehicle_id " & szWhereBusVehicle _
        & " ORDER BY bus_id ,vehicle_serial "
    Set rstempBusVehicle = oDb.Execute(szSql)
    '''''''''''''
    '判断计划中是否有车次
    If rsTemp.RecordCount <> 0 Then
        ReDim szaTemp(1 To rsTemp.RecordCount, 1 To 13) As String
        For i = 1 To rsTemp.RecordCount
            szaTemp(i, 1) = rsTemp!BUS_ID
            If rsTemp!bus_type <> TP_ScrollBus Then
                szaTemp(i, 2) = rsTemp!bus_start_time
            Else
                szaTemp(i, 2) = "滚动车次"
            End If
            szaTemp(i, 3) = FormatDbValue(rsTemp!check_gate_id)
            szaTemp(i, 4) = FormatDbValue(rsTemp!route_name)
            szaTemp(i, 5) = rsTemp!bus_type
            szaTemp(i, 6) = FormatDbValue(rsTemp!stop_start_date)
            szaTemp(i, 7) = FormatDbValue(rsTemp!stop_end_date)
            szaTemp(i, 8) = FormatDbValue(rsTemp!bus_Type_name)
            szaTemp(i, 9) = FormatDbValue(rsTemp!check_gate_name)
            szaTemp(i, 12) = FormatDbValue(rsTemp!end_station_id) '终点站代码
        
            nSerial = (DateDiff("d", "2003-01-02", Date) Mod rsTemp!bus_run_cycle + rsTemp!run_start_serial - 1) Mod rsTemp!bus_run_cycle + 1  '(DateDiff("d", dtBeginDate, Now) Mod rsTemp!bus_run_cycle + rsTemp!run_start_serial - 1) Mod rsTemp!bus_run_cycle + 1
            '  rstempBusVehicle.MoveFirst
            
            If CInt(rsTemp!bus_type) <> TP_ScrollBus Then
                rstempBusVehicle.Find "bus_id=" & " '" & Trim(szaTemp(i, 1)) & " '"
                rstempBusVehicle.Find "vehicle_serial = " & nSerial
                If rstempBusVehicle.EOF Then
                    If rstempBusVehicle.RecordCount > 0 Then
                        rstempBusVehicle.MoveFirst
                    End If
                    szaTemp(i, 11) = "(停)"
                    szaTemp(i, 10) = "没安排车辆"
                ElseIf FormatDbValue(rstempBusVehicle!BUS_ID) <> Trim(szaTemp(i, 1)) Then
                    If rstempBusVehicle.RecordCount > 0 Then
                        rstempBusVehicle.MoveFirst
                    End If
                    szaTemp(i, 11) = "(停)"
                    szaTemp(i, 10) = "没安排车辆"
                ElseIf rstempBusVehicle.EOF = False Then
                    '  If nSerial > 1 Then
                    rstempBusVehicle.Move nSerial - 1
                    '  End If
Here:
                    If rstempBusVehicle.EOF = True Then
                        rstempBusVehicle.Move -1
                    End If
                    If rstempBusVehicle!BUS_ID = szaTemp(i, 1) Then
                        If rstempBusVehicle!Status = ST_VehicleStop Then
                            szaTemp(i, 11) = "(停)"
                        Else
                            If Format(rstempBusVehicle!stop_start_date, "YYYY-MM-dd") = cszForeverDateStr Then
                                szaTemp(i, 11) = "(停)"
                            Else
                                If Format(rstempBusVehicle!stop_start_date, "YYYY-MM-dd") <> cszEmptyDateStr And Format(rstempBusVehicle!stop_start_date, "YYYY-MM-dd") <> cszEmptyDateStr Then
                                    If CDate(rstempBusVehicle!stop_start_date) >= CDate(Format(Now, "YYYY-MM-DD")) And CDate(rstempBusVehicle!stop_start_date) <= CDate(Format(Now, "YYYY-MM-DD")) Then
                                        szaTemp(i, 11) = "(停)"
                                    End If
                                Else
                                    szaTemp(i, 11) = ""
                                End If
                            End If
                        End If
                    Else
                        rstempBusVehicle.Move -1
                        GoTo Here
                    End If
                    szaTemp(i, 10) = Trim(FormatDbValue(rstempBusVehicle!vehicle_id)) & "[" & Trim(FormatDbValue(rstempBusVehicle!license_tag_no)) & "]"
'                Else
'                    If rstempBusVehicle.RecordCount > 0 Then
'                        rstempBusVehicle.MoveFirst
'                    End If
'                    szaTemp(i, 11) = "(停)"
'                    szaTemp(i, 10) = "没安排车辆"
                End If
            End If
            rsTemp.MoveNext
        Next
        GetAllBus = szaTemp
    End If
    
    Set rstempBusVehicle = Nothing
    Set rsTemp = Nothing
    Set oDb = Nothing
    
    Exit Function
    
ErrHere:
    
    Set rstempBusVehicle = Nothing
    Set rsTemp = Nothing
    Set oDb = Nothing
    
    err.Raise err.Number

End Function
'Private Function FindVehile(rsTemp As Recordset, szBusID As String, nSeral As Integer)
'
'End Function
'**************************************************
'Member Code:F4
'Brief Description:获得车次的所有票价
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Public Function GetAllPriceTable() As String()
    Dim szaTemp() As String
    Dim oDb As New RTConnection
    Dim i As Integer

    Dim rsTemp As Recordset
AssertObjIsValid
'On Error GoTo here
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)

    Set rsTemp = oDb.Execute("SELECT * FROM price_table_info ")
    '-----------------------
    '判断计划中是否有车次
    If rsTemp.RecordCount <> 0 Then
        ReDim szaTemp(1 To rsTemp.RecordCount) As String
        For i = 1 To rsTemp.RecordCount
        szaTemp(i) = FormatDbValue(rsTemp!price_table_id)
        rsTemp.MoveNext
        Next
        GetAllPriceTable = szaTemp
    End If
End Function
'
''**************************************************
''Member Code:P1
''Brief Description:活动用户
''**************************************************
Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
End Property
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
''**************************************************
''Member Code:P2
''Brief Description:对象状态
''**************************************************
Public Property Get ObjStatus() As EObjectStatus
    ObjStatus = m_nObjectStatus
End Property
'
''**************************************************
''Member Code:P3
''Brief Description:计划代码
''**************************************************
'Public Property Let ProjectID(ByVal vData As String)
'    Dim oDb As New RTConnection
'    Dim rsTemp As Recordset
'    Dim szSql As String
'    AssertActiveUserValid m_oActiveUser, ERR_BusProject
'    '查询是否是新增状态
'    If m_nObjectStatus = ST_AddObj Then m_szProjectID = vData: Exit Property
'    If m_nObjectStatus <> ST_AddObj Then ShowError ERR_BusProjectNotAddStatus
'    '输入的主键是否为空
'    If vData = "" Then ShowError ERR_BsuProjectIDNotNull
'    '查询该对象是否已有记录
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'    szSql = "SELECT * FROM bus_project_info  WHERE project_id='" & RTrim(vData) & "'"
'    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount = 1 Then ShowError ERR_BusProjectExistNotAddNew
'    m_szProjectID = RTrim(vData)
'End Property
'Public Property Get ProjectID() As String
'    ProjectID = m_szProjectID
'End Property
'
''**************************************************
''Member Code:P4
''Brief Description:计划的执行票价表
''**************************************************
Public Property Let ExecutePriceTable(ByVal vData As String)
    m_szExecutePriceTable = vData
    If m_nObjectStatus = ST_NormalObj Then m_nObjectStatus = ST_EditObj
End Property
Public Property Get ExecutePriceTable() As String
    ExecutePriceTable = m_szExecutePriceTable
End Property
''**************************************************
''Member Code:P5
''Brief Description:计划注释
''**************************************************
Public Property Let Annotation(ByVal vData As String)
    m_szAnnotation = vData
    If m_nObjectStatus = ST_NormalObj Then m_nObjectStatus = ST_EditObj
End Property
Public Property Get Annotation() As String
    Annotation = m_szAnnotation
End Property
'**************************************************
'Member Code:P6
'Brief Description:计划的最后修改日期
'**************************************************
Public Property Let LastUpdateTime(ByVal vData As Date)
    ShowError ERR_BusProjectPropertyNotEdit
End Property
Public Property Get LastUpdateTime() As Date
    LastUpdateTime = m_dtLastUpdateTime
End Property
'**************************************************
'Member Code:P7
'Brief Description:计划的创建日期
'**************************************************
Public Property Let CreateTime(ByVal vData As Date)
    ShowError ERR_BusProjectPropertyNotEdit
End Property
Public Property Get CreateTime() As Date
    CreateTime = m_dtCreateTime
End Property
'**************************************************
'Member Code:P8
'Brief Description:计划状态
'**************************************************
Public Property Let PrjectStatus(ByVal vData As EBusProjectStatus)
    m_ePrjectStatus = vData
    If m_nObjectStatus = ST_NormalObj Then m_nObjectStatus = ST_EditObj
End Property
Public Property Get PrjectStatus() As EBusProjectStatus
    PrjectStatus = m_ePrjectStatus
End Property

'**************************************************
'Member Code:P9
'Brief Description:计划名称
'**************************************************
Public Property Let ProjectName(ByVal vData As String)
    m_szProjectName = Trim(vData)
    If m_nObjectStatus = ST_NormalObj Then m_nObjectStatus = ST_EditObj
End Property
Public Property Get ProjectName() As String
    ProjectName = m_szProjectName
End Property

'P10执行票价表名
Public Property Get ExecutePriceTableName() As String
    ExecutePriceTableName = m_szExecutePriceTableName
End Property
''**************************************************
''Member Code:S1
''Brief Description:初始化对象
''Engineer:
''Date Generated:
''Last Revision Date:
'**************************************************
Public Sub Init(poAUser As ActiveUser)
'**************************************************
'poAUser(活动用户)
'**************************************************
    Set m_oActiveUser = poAUser
    'AssertHaveRight m_oActiveUser, RIGHT_BusProjectManagement
End Sub
'
''**************************************************
''Member Code:S2
''Brief Description:初始化类
''Engineer:
''Date Generated:
''Last Revision Date:
''**************************************************
'Private Sub Class_Initialize()
'    m_nObjectStatus = ST_NotAvailable
'End Sub
'
''**************************************************
''Member Code:S3
''Brief Description:测试对象是是否有效
''Engineer:
''Date Generated:
''Last Revision Date:
''**************************************************
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then ShowError ERR_NotAvailable
End Sub
'
''**************************************************
''Member Code:S4
''Brief Description:测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
''Engineer:
''Date Generated:
''Last Revision Date:
''**************************************************
Private Sub AssertObjIsValid()
    AssertActiveUserValid m_oActiveUser, ERR_BusProject
    AssertStatusAvailable
End Sub
''**************************************************
''Member Code:S5
''Brief Description:标示对象
''Engineer:
''Date Generated:
''Last Revision Date:
''**************************************************
Public Sub Identify()
'**************************************************
'ProjectID(计划代码)
'**************************************************
AssertActiveUserValid m_oActiveUser, ERR_BusProject
    RefreshMemoryInfo
    m_nObjectStatus = ST_NormalObj
End Sub
''
''**************************************************
''Member Code:S6
''Brief Description:新增对象
''Engineer:
''Date Generated:
''Last Revision Date:
''**************************************************
Public Sub AddNew()
    AssertActiveUserValid m_oActiveUser, ERR_BusProject
    m_nObjectStatus = ST_AddObj
End Sub
''**************************************************
''Member Code:S7
''Brief Description:删除对象
''Engineer:
''Date Generated:
''Last Revision Date:
''**************************************************
Public Sub Delete()
    AssertObjIsValid
  '  DeleteObject
    m_nObjectStatus = ST_NotAvailable
End Sub
'
''**************************************************
''Member Code:S8
''Brief Description:修改对象属性
''Engineer:
''Date Generated:
''Last Revision Date:
''**************************************************
Public Sub Update()
    AssertObjIsValid
    Select Case m_nObjectStatus
           Case ST_AddObj
'           AddObject
           Case ST_EditObj
'           UpdateToDB
           Case ST_NormalObj
           Case Else
           ShowError ERR_NotAvailable
    End Select
    m_nObjectStatus = ST_NormalObj
End Sub
'
'
''**************************************************
''Member Code:S9
''Brief Description:从数据库中提取对象属性
''Engineer:
''Date Generated:
''Last Revision Date:
'''**************************************************
Private Sub RefreshMemoryInfo()
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_BusProject

    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    '1获得车次的基本属性
'    szSql = "SELECT * FROM bus_project_info  WHERE project_id='" & m_szProjectID & "'"
'    Set rsTemp = oDb.Execute(szSql)
'
'    If rsTemp.RecordCount = 0 Then ShowError ERR_ProjectNotExist
'        m_szAnnotation = FormatDbValue(rsTemp!Annotation)
'        m_szProjectName = FormatDbValue(rsTemp!project_name)
'        m_ePrjectStatus = FormatDbValue(rsTemp!Status)
'        m_dtCreateTime = FormatDbValue(rsTemp!create_time)
'        m_dtLastUpdateTime = FormatDbValue(rsTemp!last_modify_time)

    '2查询车次的执行票价表
    Dim lErrCode As Long
    m_szExecutePriceTable = BusProjectExecutePrice(Now, lErrCode)
    If lErrCode = 1 Then ShowError ERR_ExecutePriceTableNotExist

    '********* add **************
    szSql = "SELECT price_table_name FROM price_table_info  WHERE price_table_id='" & m_szExecutePriceTable & "' "
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 0 Then
'*********** remark ************************
'       m_szExecutePriceTable = ""
        m_szExecutePriceTableName = ""

'        m_szExecutePriceTable = Trim(rsTemp!price_table_id)
        m_szExecutePriceTableName = FormatDbValue(rsTemp!price_table_name)
    End If
    Set oDb = Nothing
    m_nObjectStatus = ST_NormalObj
End Sub
'
''**************************************************
''Member Code:S10
''Brief Description:新增对象
''Engineer:
''Date Generated:
''Last Revision Date:
''**************************************************
''Private Sub AddObject()
''    Dim OPara As New SystemParam
''    Dim oDb As New RTConnection
''    Dim szTempSql As String
''    Dim szSqlField As String
''    Dim szSqlContext As String
''    Dim rsTemp As Recordset
''AssertActiveUserValid m_oActiveUser, ERR_BusProject
''    If Trim(m_szProjectID) = "" Then ShowError ERR_BusIDNotNull
''On Error GoTo Here
''    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
''    szTempSql = "select * from Bus_project_info where project_id= '" & m_szProjectID & "'or project_name='" & m_szProjectName & "'"
''    Set rsTemp = oDb.Execute(szTempSql)
''    If rsTemp.RecordCount <> 0 Then ShowError ERR_BusProjectReplecProjectIDorName
''
''    szTempSql = "INSERT bus_project_info ("
''    szSqlField = szSqlField & "project_id,"
''    szSqlField = szSqlField & "project_name,"
''    szSqlField = szSqlField & "Annotation,"
''    szSqlField = szSqlField & "Status,"
''    szSqlField = szSqlField & "create_time,"
''    szSqlField = szSqlField & "last_modify_time)"
''    OPara.Init m_oActiveUser
''    szSqlContext = " VALUES ('" & m_szProjectID & "',"
''    szSqlContext = szSqlContext & " '" & m_szProjectName & "',"
''    szSqlContext = szSqlContext & " '" & m_szAnnotation & "',"
''    szSqlContext = szSqlContext & "" & Str(m_ePrjectStatus) & ","
''    szSqlContext = szSqlContext & " '" & ToDBDateTime(OPara.NowDate) & "',"
''    szSqlContext = szSqlContext & " '" & ToDBDateTime(OPara.NowDate) & "')"
''    szTempSql = szTempSql & szSqlField & szSqlContext
''    oDb.Execute (szTempSql)
''    WriteOperateLog m_oActiveUser, RIGHT_BusProjectAdd, "新增计划" & m_szProjectName
''Exit Sub
''Here:
''    AssertAddObjectError ERR_BusProject, oDb
''End Sub
'
''**************************************************
''Member Code:S11
''Brief Description:删除对象
''Engineer:
''Date Generated:
''Last Revision Date:
''**************************************************
''Private Sub DeleteObject()
''    Dim oDb As New RTConnection
''    Dim szSql As String
''    Dim rsTemp As Recordset
''AssertObjIsValid
''    AssertHaveRight m_oActiveUser, RIGHT_BusProjectDelete
''    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
''    szSql = "select * from project_execute_lst where project_id='" & m_szProjectID & "'"
''    Set rsTemp = oDb.Execute(szSql)
''    If rsTemp.RecordCount <> 0 Then ShowError ERR_alreadyExitstProject
''    oDb.Execute ("UPDATE bus_project_info SET status=" & Str(ST_BusProjectDeleted) & " WHERE project_id='" & m_szProjectID & "'")
''    WriteOperateLog m_oActiveUser, RIGHT_BusProjectDelete, "逻辑删除计划" & m_szProjectID
''End Sub
'
''**************************************************
''Member Code:S12
''Brief Description:修改对象属性
''Engineer:
''Date Generated:
''Last Revision Date:
''**************************************************
''Private Sub UpdateToDB()
''    Dim oDb As New RTConnection
''    Dim OPara As New SystemParam
''    Dim rsTemp As Recordset
''    Dim szTempSql As String
''    Dim szSqlField As String
''    Dim szSqlWhere As String
''AssertObjIsValid
''     AssertHaveRight m_oActiveUser, RIGHT_BusProjectManagement
''    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
''    '------------------------------------
''    '第一步：修改计划基本信息
''    szTempSql = "UPDATE bus_project_info SET "
''    szSqlField = "project_name='" & m_szProjectName & "',"
''    szSqlField = szSqlField & "Annotation='" & m_szAnnotation & "',"
''    szSqlField = szSqlField & "Status=" & Str(m_ePrjectStatus) & ","
''    OPara.Init m_oActiveUser
''    szSqlField = szSqlField & "last_modify_time='" & ToDBDateTime(OPara.NowDate) & "'"
''    szSqlWhere = " WHERE project_id='" & m_szProjectID & "'"
''    szTempSql = szTempSql & szSqlField & szSqlWhere
''    oDb.Execute (szTempSql)
''    '-----------------------------------
''    '第二步：修改计划的对象票价表
''    '查询该票价表是否正确
''    'status--remarke
'''    szTempSql = "UPDATE price_table_info SET status=" & Str(ST_PriceTableNotExecute) & " WHERE project_id='" & m_szProjectID & "'"
'''    odb.Execute (szTempSql)
'''    szTempSql = "UPDATE price_table_info SET status=" & Str(ST_PriceTableExecute) & " WHERE project_id='" & m_szProjectID & "' AND price_table_id='" & m_szExecutePriceTable & "' "
'''    odb.Execute (szTempSql)
''   WriteOperateLog m_oActiveUser, RIGHT_BusProjectManagement, "修改计划基本信息"
''End Sub
'
''**************************************************
''Member Code:S13
''Brief Description:将逻辑上删除的车次计划从物理上删除
''Engineer:
''Date Generated:
''Last Revision Date:
''**************************************************
Public Sub Purge()
'    Dim oDb As New RTConnection
'AssertObjIsValid
'    AssertHaveRight m_oActiveUser, RIGHT_BusProjectPurge
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'    If m_ePrjectStatus = ST_BusProjectDeleted Then
'        oDb.Execute "DELETE project_execute_lst WHERE project_id='" & m_szProjectID & "'"
'        oDb.Execute "DELETE bus_project_info WHERE status=" & Str(ST_BusProjectDeleted) & " AND project_id='" & m_szProjectID & "'"
'    End If
'    WriteOperateLog m_oActiveUser, RIGHT_BusProjectPurge, "将逻辑上删除的车次计划[" & m_szProjectID & "]从物理上删除"
End Sub
'
''**************************************************
''Member Code:S14
''Brief Description:将逻辑上删除的车次计划复原
''Engineer:
''Date Generated:
''Last Revision Date:
'**************************************************
Public Sub Recover()
'    Dim oDb As New RTConnection
'    AssertObjIsValid
'    AssertHaveRight m_oActiveUser, RIGHT_BusProjectRecover
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'    oDb.Execute ("UPDATE bus_project_info SET status=" & Str(ST_BusProjectNormal) & " WHERE project_id='" & m_szProjectID & "'")
'    WriteOperateLog m_oActiveUser, RIGHT_BusProjectRecover, "将逻辑上删除的车次计划[" & m_szProjectID & "]复原"
End Sub
'
Public Function GetAllBusReport() As Recordset
    Dim oDb As New RTConnection
    Dim szSql As String
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "SELECT tbb.*,tbc.check_gate_id AS check_gate_name,tbr.route_name,bvc.vehicle_id,vi.license_tag_no " _
            & " FROM bus_info tbb,Checkgate_info tbc,route_info tbr,bus_vehicle_code bvc,Vehicle_info vi " _
            & " WHERE  tbb.check_gate_id=tbc.check_gate_id AND tbr.route_id=tbb.route_id AND bvc.vehicle_id=vi.vehicle_id AND bvc.bus_id = tbb.bus_id"
    Set GetAllBusReport = oDb.Execute(szSql)
End Function

Public Function GetAllBusVehicleReport()
    Dim oDb As New RTConnection
    Dim szSql As String
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "SELECT tbr.route_name,tbc.bus_start_time,tbc.check_gate_id,cgi.check_gate_name, tbv.start_seat_no," _
            & " tbc.bus_type,vcd.vehicle_type_short_name,bti.bus_type_name, tbc.bus_run_cycle,tbc.run_start_serial," _
            & " tbb.bus_id,tbb.vehicle_id,tbb.vehicle_serial,tbv.license_tag_no,tbv.seat_quantity,tbc.stop_start_date," _
            & " tbc.stop_end_date,tbb.sale_stand_ticket_quantity,tbv.status ," _
            & " tbf.transport_company_short_name,owner_name,tbf2.transport_company_short_name split_company_short_name " _
            & " FROM bus_vehicle_code tbb,bus_info tbc,route_info tbr,Vehicle_info" _
            & " tbv,owner_info tbj,Company_info tbf , Company_info tbf2, " _
             & " vehicle_type_code vcd ,Bus_type_code bti ,Checkgate_info cgi WHERE tbc.route_id=tbr.route_id AND tbb.vehicle_id=tbv.vehicle_id AND " _
            & " tbb.bus_id = tbc.bus_id And " _
            & " tbv.owner_id = tbj.owner_id And tbv.transport_company_id = tbf.transport_company_id " _
            & " AND tbv.split_company_id = tbf2.transport_company_id  " _
            & " and cgi.check_gate_id=tbc.check_gate_id" _
            & " and vcd.vehicle_type_code=tbv.vehicle_type_code and bti.bus_type=tbc.bus_type" _
            & " ORDER BY tbc.bus_start_time,tbb.bus_id,tbb.vehicle_serial"
    Set GetAllBusVehicleReport = oDb.Execute(szSql)
End Function
'
Public Function GetBusVehicleReport(Optional szAereID As String)
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szMsg As String
    Dim szMsgTable As String
    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)

    szSql = "SELECT tbr.route_name,tbc.bus_start_time,tbc.check_gate_id,cgi.check_gate_name, tbv.start_seat_no," _
            & " tbc.bus_type,vcd.vehicle_type_short_name,bti.bus_type_name, tbc.bus_run_cycle,tbc.run_start_serial," _
            & " tbb.bus_id,tbb.vehicle_id,tbb.vehicle_serial,tbv.license_tag_no,tbv.seat_quantity,tbc.stop_start_date," _
            & " tbc.stop_end_date,tbb.sale_stand_ticket_quantity,tbv.status ," _
            & " tbf.transport_company_short_name,owner_name,tbf2.transport_company_short_name split_company_short_name "
    szSql = szSql & "FROM bus_vehicle_code tbb,bus_info tbc,route_info tbr,Vehicle_info" _
            & " tbv,owner_info tbj,Company_info tbf , Company_info tbf2, " _
            & " vehicle_type_code vcd ,Bus_type_code bti ,Checkgate_info cgi"

    If szAereID <> "" Then
       szSql = szSql & ",Station_info s"
       szMsg = " and ltrim(s.area_code) = '" & Trim(szAereID) & "' and tbr.end_station_id=s.station_id   "
    End If

    szSql = szSql & " WHERE tbc.route_id=tbr.route_id AND tbb.vehicle_id=tbv.vehicle_id AND " _
            & " tbb.bus_id = tbc.bus_id And " _
            & " tbv.owner_id = tbj.owner_id And tbv.transport_company_id = tbf.transport_company_id " _
            & " AND tbv.split_company_id = tbf2.transport_company_id  " _
            & " and cgi.check_gate_id=tbc.check_gate_id" _
            & " and vcd.vehicle_type_code=tbv.vehicle_type_code and bti.bus_type=tbc.bus_type"
           
    szSql = szSql & szMsg & " ORDER BY tbc.bus_start_time,tbb.bus_id,tbb.vehicle_serial"
    Set GetBusVehicleReport = oDb.Execute(szSql)
End Function

'计划车次预订
Public Function ProjectBook(dtStartBusDate As Date, dtEndBusDate As Date, szBookInfo() As String, szBusID() As String) As Boolean
Dim oDb As New RTConnection
Dim szSql As String
Dim nCount As Integer
Dim szWhereBusID As String
Dim i As Integer
Dim szlog As String
Dim szLogBus As String
On Error GoTo Here

AssertHaveRight m_oActiveUser, RIGHT_BusProjectBooked
oDb.ConnectionString = GetConnectionStr(cszSystemMan)
nCount = ArrayLength(szBusID)

'没有车次信息，没有总座位数
If nCount = 0 Or CInt(szBookInfo(1)) = 0 Then Exit Function
szWhereBusID = ""
For i = 0 To nCount - 1
   szWhereBusID = szWhereBusID & "'" & szBusID(i) & "',"
   szLogBus = szLogBus & "[" & szBusID(i) & "]"
Next

szWhereBusID = Left(szWhereBusID, Len(szWhereBusID) - 1)

szSql = "delete project_book_info where " _
        & "( " _
        & " ( start_bus_date<='" & ToDBDate(dtEndBusDate) & "' and end_bus_date>='" & ToDBDate(dtEndBusDate) & "')" _
        & " or " _
        & " (start_bus_date<='" & ToDBDate(dtStartBusDate) & "' and end_bus_date>='" & ToDBDate(dtEndBusDate) & "') " _
        & " or " _
        & " (start_bus_date<='" & ToDBDate(dtStartBusDate) & "'and end_bus_date>='" & ToDBDate(dtStartBusDate) & "')" _
        & " or " _
        & " (start_bus_date>='" & ToDBDate(dtStartBusDate) & "' and end_bus_date<='" & ToDBDate(dtEndBusDate) & "')" _
        & ")" _
        & " and bus_id in(" & szWhereBusID & ")"
oDb.Execute szSql

For i = 0 To nCount - 1
    szSql = "insert project_book_info " _
            & " (start_bus_date,end_bus_date,bus_id,seat_start_no,Seat_book_count,remark_info)" _
            & "  values ( " _
            & " '" & ToDBDate(dtStartBusDate) & "' ," _
            & "  '" & ToDBDate(dtEndBusDate) & "'," _
            & " '" & szBusID(i) & "'," _
            & " '" & szBookInfo(0) & "'," _
            & " '" & szBookInfo(1) & "'," _
            & " '" & szBookInfo(2) & "')"
    oDb.Execute szSql
Next

szlog = "在时间段" & ToDBDate(dtStartBusDate) & "-" & ToDBDate(dtEndBusDate) & "计划预车次" & szLogBus

If Len(szlog) > 255 Then
   szlog = Left(szlog, 254) & "等"
End If

WriteOperateLog m_oActiveUser, RIGHT_BusProjectManagement, szlog
ProjectBook = True

Exit Function
Here:
    err.Raise err.Number
End Function

''计划车次预订,
''bIsNotAllBus=false 将删除全部满足时间条件的预定信息
''bIsNotAllBus=true  将删除全部满足时间条件的,且满足车次条件预定信息
Public Function ProjectBookDelete(dtStartBusDate As Date, dtEndBusDate As Date, szBusID() As String, Optional bIsNotAllBus As Boolean = True) As Boolean
Dim oDb As New RTConnection
Dim szSql As String
Dim nCount As Integer
Dim szWhereBusID As String
Dim Lrow As Long
Dim i As Integer
Dim szlog As String
Dim szLogBus As String
On Error GoTo Here
AssertHaveRight m_oActiveUser, RIGHT_BusProjectBooked
oDb.ConnectionString = GetConnectionStr(cszSystemMan)

If bIsNotAllBus = True Then

    nCount = ArrayLength(szBusID)

    If nCount = 0 Then Exit Function '没有车次信息

    For i = 0 To nCount - 1
        szWhereBusID = szWhereBusID & "'" & szBusID(i) & "',"
        szLogBus = szLogBus & "[" & szBusID(i) & "]"
    Next

End If

szWhereBusID = Left(szWhereBusID, Len(szWhereBusID) - 1)

szSql = "delete project_book_info where " _
        & " start_bus_date='" & ToDBDate(dtStartBusDate) & "' and end_bus_date='" & ToDBDate(dtEndBusDate) & "' "


szlog = "在时间段" & ToDBDate(dtStartBusDate) & "--" & ToDBDate(dtEndBusDate) & "的车次"

If bIsNotAllBus = True Then
   szSql = szSql & " and bus_id in(" & szWhereBusID & ")"
   szlog = szlog & szLogBus
End If

oDb.Execute szSql, Lrow

If Len(szlog) > 255 Then
    szlog = Left(szlog, 245) & "等"
End If

szlog = szlog & "计划预定信息被删除"


If Lrow > 0 Then
    ProjectBookDelete = True
    WriteOperateLog m_oActiveUser, RIGHT_BusProjectManagement, szlog
Else
    ProjectBookDelete = False
End If

Set oDb = Nothing

Exit Function

Here:
    err.Raise err.Number

End Function
''计划车次预订,
''bIsNotAllBus=false 将删除全部满足时间条件的预定信息
''bIsNotAllBus=true  将删除全部满足时间条件的,且满足车次条件预定信息
Public Function ProjectBookQuery(dtStartBusDate As Date, dtEndBusDate As Date, Optional szBusID As String, Optional bIsAllBus As Boolean = True) As String()
Dim oDb As New RTConnection
Dim szSql As String
Dim nCount As Integer
Dim rsTemp As New Recordset
Dim szBookInfo() As String
Dim i As Integer
On Error GoTo Here

AssertHaveRight m_oActiveUser, RIGHT_BusProjectBooked

oDb.ConnectionString = GetConnectionStr(cszSystemMan)


'If bIsSelect = False Then
        szSql = "SELECT p.*,r.route_name FROM  project_book_info p,route_info r ,bus_info b WHERE " _
                & "( " _
                & " ( p.start_bus_date<='" & ToDBDate(dtEndBusDate) & "' and p.end_bus_date>='" & ToDBDate(dtEndBusDate) & "')" _
                & " or " _
                & " (p.start_bus_date<='" & ToDBDate(dtStartBusDate) & "' and p.end_bus_date>='" & ToDBDate(dtEndBusDate) & "') " _
                & " or " _
                & " (p.start_bus_date<='" & ToDBDate(dtStartBusDate) & "'and p.end_bus_date>='" & ToDBDate(dtStartBusDate) & "')" _
                & " or " _
                & " (p.start_bus_date>='" & ToDBDate(dtStartBusDate) & "' and p.end_bus_date<='" & ToDBDate(dtEndBusDate) & "')" _
                & ")"
        If szBusID <> "" Then
            szSql = szSql & " and p.bus_id ='" & szBusID & "'"
        End If

        szSql = szSql & " and p.bus_id=b.Bus_id and r.route_id=b.route_id"

'Else
'      szSql = "select p.*,r.route_name from  project_book_info p,route_info r ,bus_info b where " _
'                & "( " _
'                & " ( p.start_bus_date<='" & ToDBDate(dtEndBusDate) & "' and p.end_bus_date>='" & ToDBDate(dtEndBusDate) & "')" _
'       If szBusID <> "" Then
'            szSql = szSql & " and p.bus_id ='" & szBusID & "'"
'       End If
'
'       szSql = szSql & " and p.bus_id=b.Bus_id and r.route_id=b.route_id"
'End If
Set rsTemp = oDb.Execute(szSql)

nCount = rsTemp.RecordCount

If nCount = 0 Then Exit Function '无记录

ReDim szBookInfo(0 To nCount - 1, 0 To 6)
For i = 0 To nCount - 1
    szBookInfo(i, 0) = rsTemp!BUS_ID
    szBookInfo(i, 1) = rsTemp!route_name
    szBookInfo(i, 2) = rsTemp!start_bus_date
    szBookInfo(i, 3) = rsTemp!End_bus_date
    szBookInfo(i, 4) = rsTemp!seat_start_no
    szBookInfo(i, 5) = rsTemp!seat_book_count
    szBookInfo(i, 6) = rsTemp!remark_info
    rsTemp.MoveNext
Next

ProjectBookQuery = szBookInfo
Set oDb = Nothing
Exit Function

Here:
    err.Raise err.Number
End Function
'




