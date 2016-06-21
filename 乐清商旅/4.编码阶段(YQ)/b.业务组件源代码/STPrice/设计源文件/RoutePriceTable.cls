VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RoutePriceTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**********************************************************
'* Source File Name:RoutePriceTable.cls
'* Project Name:StPrice.vbp
'* Engineer:
'* Data Generated:
'* Last Revision Date:
'* Brief Description:线路票价表
'* Relational Document:
'**********************************************************

Option Explicit

Const cnTotalBusPriceItem = 16

Const cszBaseCharge = "0000"
Const cszPriceItem1 = "0001"
Const cszPriceItem2 = "0002"
Const cszPriceItem3 = "0003"
Const cszPriceItem4 = "0004"
Const cszPriceItem5 = "0005"
Const cszPriceItem6 = "0006"
Const cszPriceItem7 = "0007"
Const cszPriceItem8 = "0008"
Const cszPriceItem9 = "0009"
Const cszPriceItem10 = "0010"
Const cszPriceItem11 = "0011"
Const cszPriceItem12 = "0012"
Const cszPriceItem13 = "0013"
Const cszPriceItem14 = "0014"
Const cszPriceItem15 = "0015"

Const cszDisCountFormula = "不计算"

Public Enum EErrRoutePriceTable
    ERR_SpecialRoutePriceTableNotExist = ERR_RouteTicketPrice + 21  '对象不存在错误
    ERR_RoutePriceTableNoRoute = ERR_RouteTicketPrice + 22      '无线路
    ERR_RoutePriceTableNoStation = ERR_RouteTicketPrice + 23    '无指定的车站
    ERR_RoutePriceTableNoVehicle = ERR_RouteTicketPrice + 24    '无指定的车型
    ERR_RoutePriceTableStationNotInRoute = ERR_RouteTicketPrice + 25    '指定的车站不在此线路上
    ERR_RoutePriceTableNoTicketPrice = ERR_RouteTicketPrice + 26   '无指定的票价项
    
    ERR_RoutePriceTableNoMakeRows = ERR_RouteTicketPrice + 27     '没有指定票价项生成错误
    ERR_RoutePriceTableNoModifyRows = ERR_RouteTicketPrice + 28     '没有指定票价项更改错误
    ERR_SpecifyRouteNoFormulaAndNoDefaultFormula = ERR_RouteTicketPrice + 29 '指定的线路无票价计算公式且当前无缺省的执行的票价计算公式
    ERR_NoSpecifyRoute = ERR_RouteTicketPrice + 30 '无指定的线路
    ERR_FormulaNotRight = ERR_RouteTicketPrice + 31 '票价计算公式不全
    ERR_RoutePriceTableSetExec = ERR_RouteTicketPrice + 32 '设定执行票价表错误
    ERR_CopyRoutePriceTable = ERR_RouteTicketPrice + 33 '复制线路票价表错误
    ERR_TicketFunLibNotSpecifyFun = ERR_RouteTicketPrice + 34 '票价计算公式库中无指定的计算方法
    ERR_RoutePriceHaveExist = ERR_RouteTicketPrice + 35 '此线路票价中已存在此票价项
    ERR_ReMakeHalfPrice = ERR_RouteTicketPrice + 36 '重新生成线路票价表的半票错
    ERR_ModifyBusPrice = ERR_RouteTicketPrice + 45 '保存车次票价错
    ERR_SeatType = ERR_RouteTicketPrice + 46 '座位类型为空错
    ERR_BusPriceTableNoBus = ERR_RouteTicketPrice + 47      '无车次
    ERR_BusPriceMakeExit = ERR_RouteTicketPrice + 48      '退出车次票价生成
    ERR_BusSpeedNoExist = ERR_RouteTicketPrice + 49      '车次平均速度没设置（不能为零）
End Enum
    
Public Enum ERightEveryObject
    RIGHT_MakeSpecifyPrice = ERR_RouteTicketPrice + cnMidRightBegin + cnMidRightStep * 1 '生成指定的票价信息
    RIGHT_ModifySpecifyPrice = ERR_RouteTicketPrice + cnMidRightBegin + cnMidRightStep * 2  '修改指定的票价信息权限
    RIGHT_DeleteSpecifyPrice = ERR_RouteTicketPrice + cnMidRightBegin + cnMidRightStep * 3  '删除指定的票价项信息权限
    RIGHT_SetExecutePrice = ERR_RouteTicketPrice + cnMidRightBegin + cnMidRightStep * 4 '设置运行状态的权限
    RIGHT_QueryPrice = ERR_RouteTicketPrice + cnMidRightBegin + cnMidRightStep * 5 '查询票价的权限
End Enum
    
'车次票价公式
Public Type TFormula
    szBusID As String
    szFormula As String
End Type

'----------------------
'各类特有的变量
Private m_szPriceTableID As String
'Private m_szProjectID As String
Private m_szPriceTableName As String
Private m_dtCreateTime As Date
Private m_dtLastModifyTime As Date
Private m_dtStartRunTime As Date

Private m_tTailDeal() As TAreaTailDealMethod
Private m_tTailDealMethod() As TTailDealMethod

Public Event SetProgressRange(ByVal lRange)
Public Event SetProgressValue(ByVal lValue)
Public Event SetMakeBusPriceStatus(ByVal lStatus As String)
Private m_CancelMakeBusPrice As Boolean '取消生成车次票价

'---------------------
'所有的实体类都有的私有类变量<<
Private m_nObjectStatus As EObjectStatus

'所有的类都有的私有类变量
Private m_oActiveUser As ActiveUser
'--------------------------->>
Private m_oTicketPriceMan As New TicketPriceMan

Public Sub Init(poAUser As ActiveUser)
    Set SelfUser = poAUser
    m_oTicketPriceMan.Init m_oActiveUser
    m_oTicketPriceMan.ObjStatus = ST_NormalObj
    m_tTailDealMethod = m_oTicketPriceMan.GetAllTailDealMethod()
End Sub

Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
    m_oActiveUser.IAmLive
End Property

Public Property Set SelfUser(vNewValue As ActiveUser)
    If m_nObjectStatus = ST_AddObj Then
        ShowError ERR_AddObj
    ElseIf m_nObjectStatus = ST_EditObj Then
        ShowError ERR_EditObj
    ElseIf m_nObjectStatus = ERR_NormalObj Then
        ShowError ERR_NormalObj
    Else
        Set m_oActiveUser = vNewValue
    End If
    m_oActiveUser.IAmLive
End Property

Private Sub Class_Initialize()
    m_nObjectStatus = ST_NotAvailable '类刚建立时处于不可用状态
End Sub
Private Sub AssertStatusAvailable() '测试对象的状态是否有效（不为无效ECNotAvailable状态）
    If m_nObjectStatus = ST_NotAvailable Then ShowError ERR_NotAvailable
End Sub
Public Sub AssertObjIsValid() '测试对象是是否有效（活动用户对象有效且对象的状态有效）
    AssertActiveUserValid m_oActiveUser, ERR_RouteTicketPrice
    AssertStatusAvailable
End Sub

Public Sub Identify(ByVal PriceTableID As String)
    AssertActiveUserValid m_oActiveUser, ERR_RouteTicketPrice
    m_szPriceTableID = PriceTableID
    m_tTailDeal = m_oTicketPriceMan.GetAreaTailMethod(m_szPriceTableID)
    RefreshMemoryInfo
    m_nObjectStatus = ST_NormalObj
End Sub

Public Sub AddNew() '调用新增方法后，对象处于新增状态。
    AssertActiveUserValid m_oActiveUser, ERR_RouteTicketPrice
    m_nObjectStatus = ST_AddObj
End Sub

Public Sub Delete() '调用删除方法后，对象处于不可用状态
    AssertObjIsValid
    
    If m_nObjectStatus <> ST_AddObj Then
        DeleteObject
    End If
    
    m_nObjectStatus = ST_NotAvailable
End Sub

Public Sub Update()
    AssertObjIsValid
    
    Select Case m_nObjectStatus
        Case ST_AddObj '如对象处于新增状态，则新增一对象
            AddObject
        Case ST_EditObj '如对象处于修改状态，则更新它
            UpdateToDB
        Case ST_NotAvailable '如果对象不可用则出错
            ShowError ERR_NotAvailable
    End Select
    
    m_nObjectStatus = ST_NormalObj
End Sub

Public Property Get ObjStatus() As EObjectStatus
    ObjStatus = m_nObjectStatus
End Property

Public Property Get CancelMakeBusPrice() As Boolean
    CancelMakeBusPrice = m_CancelMakeBusPrice
End Property

Public Property Let CancelMakeBusPrice(ByVal vNewValue As Boolean)
   m_CancelMakeBusPrice = vNewValue
End Property

Private Sub RefreshMemoryInfo() '按照对象现在的ID将数据库中相应的信息读至对象内
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    Dim szSql As String
    oDB.ConnectionString = GetConnectionStr(cszSNPrice)

    szSql = "SELECT * FROM price_table_info WHERE " _
    & " price_table_id='" & m_szPriceTableID & "' "

    Set rsTemp = oDB.Execute(szSql)

    If rsTemp.RecordCount <> 1 Then ShowError ERR_SpecialRoutePriceTableNotExist
    m_dtCreateTime = FormatDbValue(rsTemp!create_time)
    m_dtLastModifyTime = FormatDbValue(rsTemp!last_modify_time)
    m_szPriceTableID = FormatDbValue(rsTemp!price_table_id)
    m_szPriceTableName = FormatDbValue(rsTemp!price_table_name)
    m_dtStartRunTime = FormatDbValue(rsTemp!start_run_time)
'    m_szProjectID = FormatDbValue(rsTemp!project_id)
    Set rsTemp = Nothing
    Set oDB = Nothing
End Sub

Private Sub AddObject()
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim szNow As String
    Dim oParam  As New SystemParam

    oParam.Init m_oActiveUser
    szNow = ToDBDateTime(oParam.NowDateTime)

    oDB.ConnectionString = GetConnectionStr(cszSNPrice)
    szSql = "INSERT INTO price_table_info(" _
    & "price_table_id," _
    & " price_table_name," _
    & " create_time," _
    & " last_modify_time," _
    & " start_run_time" _
    & " ) VALUES ('" _
    & m_szPriceTableID & "','" _
    & m_szPriceTableName & "','" _
    & szNow & "','" _
    & szNow & "','" _
    & ToDBDateTime(m_dtStartRunTime) & "')"

    On Error GoTo Here
        oDB.Execute szSql
         Set oDB = Nothing
         Set oParam = Nothing
    Exit Sub
Here:
    AssertAddObjectError ERR_RouteTicketPrice, oDB
    Set oParam = Nothing
     Set oDB = Nothing
End Sub

Private Sub DeleteObject()
    Dim oDB As New RTConnection
    Dim szSql As String
    oDB.ConnectionString = GetConnectionStr(cszSNPrice)
    szSql = "DELETE  price_table_info WHERE " _
    & " price_table_id='" & m_szPriceTableID & "'"
On Error GoTo Here
    oDB.Execute szSql
     Set oDB = Nothing
    Exit Sub
Here:
    AssertDeleteObjectError ERR_RouteTicketPrice, oDB  '开始错误代码
     Set oDB = Nothing
End Sub

Private Sub UpdateToDB()
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim dtNow As Date
    Dim oParam As New SystemParam
    oParam.Init m_oActiveUser
    dtNow = oParam.NowDateTime

    oDB.ConnectionString = GetConnectionStr(cszSNPrice)

    oDB.BeginTrans
On Error GoTo Here
    szSql = "UPDATE price_table_info SET " _
    & " price_table_name='" & m_szPriceTableName & "'," _
    & " start_run_time='" & ToDBDateTime(m_dtStartRunTime) & "'," _
    & " last_modify_time='" & ToDBDateTime(dtNow) & "' WHERE" _
    & " price_table_id='" & m_szPriceTableID & "'"
    oDB.Execute szSql

    oDB.CommitTrans
    m_dtLastModifyTime = dtNow
    Set oParam = Nothing
     Set oDB = Nothing
    Exit Sub
Here:
    oDB.RollbackTrans
    AssertUpdateObjectError ERR_RouteTicketPrice, oDB  '开始错误代码
     Set oParam = Nothing
     Set oDB = Nothing
End Sub

Public Property Get RoutePriceTableID() As String
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_QueryPrice
    RoutePriceTableID = m_szPriceTableID
End Property

Public Property Let RoutePriceTableID(ByVal vNewValue As String)
    AssertObjIsValid
    If m_nObjectStatus <> ST_AddObj Then ShowError ERR_NotAddObj
    m_szPriceTableID = vNewValue
End Property

'------------------------
'返回所属的车次计划代码
'Public Property Get BusProject() As String
'    AssertObjIsValid
'    AssertHaveRight m_oActiveUser, RIGHT_QueryPrice
'    BusProject = m_szProjectID
'End Property

'设置所属的车次计划代码
'Public Property Let BusProject(ByVal vNewValue As String)
'    AssertObjIsValid
'    m_szProjectID = vNewValue
'    If m_nObjectStatus = ST_NormalObj Then
'        m_nObjectStatus = ST_EditObj
'    End If
'End Property

'返回线路票价名称
Public Property Get RoutePriceTableName() As String
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_QueryPrice
    RoutePriceTableName = m_szPriceTableName
End Property

'设置线路票价名称
Public Property Let RoutePriceTableName(ByVal vNewValue As String)
    AssertObjIsValid
    m_szPriceTableName = vNewValue
    
    If m_nObjectStatus = ST_NormalObj Then
        m_nObjectStatus = ST_EditObj
    End If
End Property

'返回此线路票价的生成日期时间
Public Property Get CreateTime() As Date
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_QueryPrice
    CreateTime = m_dtCreateTime
End Property

'返回最后修改日期时间
Public Property Get LastModifyTime() As Date
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_QueryPrice
    LastModifyTime = m_dtLastModifyTime
End Property

'返回开始执行日期时间
Public Property Get StartRunTime() As Date
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_QueryPrice
    StartRunTime = m_dtStartRunTime
End Property

'返回开始执行日期时间
Public Property Let StartRunTime(dtNewDate As Date)
    AssertObjIsValid
    m_dtStartRunTime = dtNewDate
End Property

'返回车次票价表是否是空的
Public Property Get IsContentEmpty() As Boolean
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_QueryPrice
    oDB.ConnectionString = GetConnectionStr(cszSNPrice)
    
    If m_nObjectStatus = ST_AddObj Then
        IsContentEmpty = True
    Else
        szSql = "SELECT COUNT(*) AS countx FROM bus_price_lst WHERE " _
        & " price_table_id='" & m_szPriceTableID & "'"
        Set rsTemp = oDB.Execute(szSql)
        IsContentEmpty = IIf(rsTemp!countx = 0, True, False)
    End If
    Set rsTemp = Nothing
     Set oDB = Nothing
End Property

'删除车次票价
Public Sub DeleteBusPrice(BusVehicleSeat() As TBusVehicleSeatType)
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim szCondition As String
    Dim i As Integer
    Dim nTemp As Integer
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_DeleteSpecifyPrice
    If m_nObjectStatus = ST_AddObj Then
        ShowError ERR_AddObj
    Else
        On Error GoTo DbError
        '得到删除条件
        szCondition = ""
        nTemp = ArrayLength(BusVehicleSeat)
        For i = 1 To nTemp
            If i = 1 Then
                szCondition = " AND ("
            End If
            szCondition = szCondition & " (bus_id= '" & BusVehicleSeat(i).szBusID & "' and seat_type_id='" & BusVehicleSeat(i).szSeatTypeID & "' and vehicle_type_code='" & BusVehicleSeat(i).szVehicleTypeCode & "')"
            If i = nTemp Then
                szCondition = szCondition & ") "
            Else
                szCondition = szCondition & " OR "
            End If
        Next
        '组合成SQL
        oDB.ConnectionString = GetConnectionStr(cszPriceMan)
'        odb.BeginTrans
        If nTemp = 0 Then ShowError ERR_BusPriceTableNoBus
        szSql = "DELETE FROM bus_price_lst WHERE price_table_id= '" & m_szPriceTableID & "'" & szCondition
    
        oDB.Execute szSql
'        odb.CommitTrans
        WriteOperateLog m_oActiveUser, RIGHT_DeleteSpecifyPrice, "删除票价表[" & m_szPriceTableID & "]条件为" & szCondition & "的部分车次票价"
    End If
    Set oDB = Nothing
    Exit Sub
DbError:
'    odb.RollbackTrans
    
    AssertDeleteObjectError ERR_RouteTicketPrice, oDB
    Set oDB = Nothing
End Sub

'Private Function FindRightPriceItem(prsPriceItem As Recordset, pszItem As String)
'    prsPriceItem.MoveFirst
'    Do While Not prsPriceItem.EOF
'        If RTrim(prsPriceItem!price_item) = RTrim(pszItem) Then
'            Exit Do
'        End If
'        prsPriceItem.MoveNext
'    Loop
'End Function

'重新生成半票价
Public Sub ReMakeHalfPrice()
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim dbParam() As Double
    Dim rsTemp As Recordset
    Dim j As Integer
    Dim szField As String
    Dim rsUseTicket As Recordset
    Dim UseTicketNo As Integer
    Dim i, nTemp As Long
    Dim BusID() As String
    Dim rsBus As Recordset
    Dim nValue As Integer
    
On Error GoTo Error_Handle
    oDB.ConnectionString = GetConnectionStr(cszPriceMan)
    '得到车次票价表中有票价的车次
    szSql = "SELECT DISTINCT bus_id FROM  bus_price_lst WHERE ticket_type=1 AND price_table_id='" & m_szPriceTableID & "'"
    Set rsBus = oDB.Execute(szSql)
    '得到票种信息
    szSql = "SELECT * FROM Ticket_type_code WHERE ticket_type_valid=" & TP_TicketTypeValid & " And ticket_type_id <> " & TP_FreeTicket & " And ticket_type_id <> " & TP_FullPrice
    Set rsUseTicket = oDB.Execute(szSql)
    UseTicketNo = rsUseTicket.RecordCount
    
    rsUseTicket.MoveFirst
    ReDim dbParam(1 To UseTicketNo, 1 To cnTotalBusPriceItem, 1 To 2)
    '得到特殊票计算公式
    szSql = "SELECT a.* FROM Ticket_type_formula_info a,Ticket_type_code b WHERE a.ticket_type=b.ticket_type_id and ticket_type_valid=" & TP_TicketTypeValid & " AND table_id='" & m_szPriceTableID & "' ORDER BY ticket_type,price_item"
    Set rsTemp = oDB.Execute(szSql)
    
    j = 0
    rsUseTicket.MoveFirst
    rsTemp.MoveFirst
    Do While Not rsUseTicket.EOF
        j = j + 1
        For i = 1 To cnTotalBusPriceItem
            dbParam(j, i, 1) = FormatDbValue(rsTemp!parameter_1)
            dbParam(j, i, 2) = FormatDbValue(rsTemp!parameter_2)
            rsTemp.MoveNext
        Next
        rsUseTicket.MoveNext
    Loop
'    odb.BeginTrans
        If rsBus.RecordCount > 0 Then
            rsBus.MoveFirst
            RaiseEvent SetProgressRange(rsBus.RecordCount)
        End If
        Do While Not rsBus.EOF
            nValue = nValue + 1
            '删除特殊票的票价
            oDB.Execute ("DELETE FROM bus_price_lst WHERE ticket_type<>" & TP_FullPrice & " AND price_table_id='" & m_szPriceTableID & "' AND Bus_id='" & rsBus!Bus_id & "'")
            
            rsUseTicket.MoveFirst
            Do While Not rsUseTicket.EOF
                '用全票写入各特殊票的票价
                szSql = "INSERT INTO bus_price_lst " _
                    & " (sell_station_id,bus_id,station_id,vehicle_type_code,price_table_id, ticket_type,seat_type_id,station_serial_no, " _
                    & " mileage,base_carriage,price_item_1,price_item_2,price_item_3,price_item_4,price_item_5, " _
                    & " price_item_6,price_item_7,price_item_8,price_item_9,price_item_10,price_item_11,price_item_12, " _
                    & " price_item_13,price_item_14,price_item_15) " _
                    & " SELECT sell_station_id,bus_id,station_id,vehicle_type_code,price_table_id, " & rsUseTicket!ticket_type_id & "," _
                    & " seat_type_id,station_serial_no,mileage,base_carriage,price_item_1,price_item_2,price_item_3," _
                    & " price_item_4,price_item_5,price_item_6,price_item_7,price_item_8,price_item_9,price_item_10," _
                    & " price_item_11,price_item_12,price_item_13,price_item_14,price_item_15 " _
                    & " FROM bus_price_lst WHERE ticket_type=" & TP_FullPrice & " AND price_table_id='" & m_szPriceTableID & "' AND Bus_id='" & rsBus!Bus_id & "'"
                
                oDB.Execute szSql 'odb.Execute szSql, , , 0
                rsUseTicket.MoveNext
            Loop
            '得到所有的全票价
            szSql = "SELECT * FROM  bus_price_lst WHERE ticket_type=" & TP_FullPrice & " AND price_table_id='" & m_szPriceTableID & "' AND Bus_id='" & rsBus!Bus_id & "'"
            Set rsTemp = oDB.Execute(szSql)
            '更新所有的特殊票的票价
            Do While Not rsTemp.EOF
                rsUseTicket.MoveFirst
                j = 1
                Do While Not rsUseTicket.EOF
                    szSql = "UPDATE bus_price_lst SET base_carriage=" & Round(FormatDbValue(rsTemp!base_carriage) * dbParam(j, 1, 1) + dbParam(j, 1, 2) + 0.001, 3)
                    For i = 1 To cnTotalBusPriceItem - 1
                        szField = "price_item_" & i
                        szSql = szSql & "," & szField & "=" & Round(FormatDbValue(rsTemp(szField)) * dbParam(j, i + 1, 1) + dbParam(j, i + 1, 2) + 0.001, 3)
                    Next
                    szSql = szSql & " WHERE " _
                        & "bus_id='" & rsTemp!Bus_id & "' AND " _
                        & "station_id='" & rsTemp!station_id & "' AND " _
                        & "vehicle_type_code='" & rsTemp!vehicle_type_code & "' AND " _
                        & "price_table_id='" & m_szPriceTableID & "' AND " _
                        & " seat_type_id='" & rsTemp!seat_type_id & "'" _
                        & " AND ticket_type=" & rsUseTicket!ticket_type_id _
                        & " AND station_serial_no='" & rsTemp!station_serial_no & "'" _
                        & " AND Bus_id='" & rsBus!Bus_id & "'"
                    oDB.Execute szSql 'odb.Execute szSql, , , 0
                    rsUseTicket.MoveNext
                    j = j + 1
                Loop
                rsTemp.MoveNext
            Loop
            '******************尾数处理
            DealTableTail m_szPriceTableID, rsBus!Bus_id
        '******************尾数处理结束
            rsBus.MoveNext
            RaiseEvent SetProgressValue(nValue)
        Loop
        
'    odb.CommitTrans
    Set rsUseTicket = Nothing
    Set rsTemp = Nothing
    Set rsBus = Nothing
    Set oDB = Nothing
    Exit Sub
Error_Handle:
'    odb.RollbackTrans
    ShowError ERR_ReMakeHalfPrice, , , True
    If Not rsUseTicket Is Nothing Then Set rsUseTicket = Nothing
    If Not rsTemp Is Nothing Then Set rsTemp = Nothing
    Set oDB = Nothing
End Sub

'生成车次票价
Public Function MakeSpecifyBusPrice(BusVehicleSeat() As TBusVehicleSeatType) As TBusPriceDetailInfo()
    Static oCaculator As Object ' PriceItemFunLib
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    Dim aPriceDetail() As TBusPriceDetailInfo
    Dim nStationCount As Long
    Dim i As Long
    Dim rsStation As Recordset
    Dim sgBase As Double
    Dim nItemIndex As Integer
    Dim RowCount As Long
    Dim sgTemp As Single
    Dim szBus As String
    Dim szRouteOld As String
    Dim szBusOld As String
    Dim szItemFormula As String
    Dim nRecPosition As Long
    Dim szTemp As String
    Dim nTemp As Integer
    Dim n As Long
    Dim k As Long
    Dim szSeat As String

    Dim ttFormula() As TFormula
    Dim nUseTicket As Integer
    Dim szBusID() As String
    Dim szCondition As String
    Dim szConditionTemp As String
    Dim nFullRow As Long

    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_MakeSpecifyPrice

    On Error GoTo Error_Handle
    oDB.ConnectionString = GetConnectionStr(cszPriceMan)
    '得到所有的可用的非免票的票种
    Set rsTemp = oDB.Execute("SELECT * FROM Ticket_type_code WHERE ticket_type_valid='" & TP_TicketTypeValid & "' AND ticket_type_id<>'" & TP_FreeTicket & "'")
    nUseTicket = rsTemp.RecordCount - 1
    '得到查询条件
    szCondition = ""
    nTemp = ArrayLength(BusVehicleSeat)
    For i = 1 To nTemp
        If i = 1 Then
            szCondition = " AND ("
            szConditionTemp = " AND ("
        End If
            szCondition = szCondition & " (bfo.bus_id= '" & BusVehicleSeat(i).szBusID & "' and sti.seat_type_id='" & BusVehicleSeat(i).szSeatTypeID & "' and vtc.vehicle_type_code='" & BusVehicleSeat(i).szVehicleTypeCode & "'" & " AND rsa.sell_station_id='" & BusVehicleSeat(i).szSellStationID & "')"
            szConditionTemp = szConditionTemp & " (bfo.bus_id= '" & BusVehicleSeat(i).szBusID & "' and ptl.seat_type_id='" & BusVehicleSeat(i).szSeatTypeID & "' and ptl.vehicle_type_code='" & BusVehicleSeat(i).szVehicleTypeCode & "'" & " AND rsa.sell_station_id='" & BusVehicleSeat(i).szSellStationID & "')"
            If i = nTemp Then
                szCondition = szCondition & ") "
                szConditionTemp = szConditionTemp & ") "
            Else
                szCondition = szCondition & " OR "
                szConditionTemp = szConditionTemp & " OR "
            End If
    Next

    szBus = ""
    nTemp = ArrayLength(BusVehicleSeat)
    If nTemp > 0 Then
        szBus = szBus & " AND ( bfo.bus_id IN('" & BusVehicleSeat(1).szBusID & "'"
        For n = 2 To nTemp
            szBus = szBus & ",'" & BusVehicleSeat(n).szBusID & "'"
        Next
        szBus = szBus & ")) "
    End If
    '得到线路的所用的票价公式
    szSql = "SELECT formula_name FROM route_info WHERE formula_name=''"
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        szSql = "SELECT formula_name FROM Price_formula_info WHERE default_execute_mark=1"
        Set rsTemp = oDB.Execute(szSql)
        If rsTemp.RecordCount <> 1 Then ShowError ERR_SpecifyRouteNoFormulaAndNoDefaultFormula
        oDB.Execute ("UPDATE route_info SET  formula_name=(SELECT formula_name FROM Price_formula_info WHERE default_execute_mark=1) WHERE formula_name=''")
    End If
    szSql = "SELECT bus_id,formula_name FROM route_info rfo,bus_info bfo WHERE formula_name<>'" & cszDisCountFormula & "' AND rfo.route_id=bfo.route_id "
    szSql = szSql & szBus

    Set rsTemp = oDB.Execute(szSql)
    nTemp = rsTemp.RecordCount
    If nTemp > 0 Then
        ReDim ttFormula(1 To nTemp)
    End If

    i = 1
    Do While Not rsTemp.EOF
        ttFormula(i).szBusID = rsTemp!Bus_id
        ttFormula(i).szFormula = FormatDbValue(rsTemp!formula_name)
        i = i + 1
        rsTemp.MoveNext
    Loop
    szTemp = ""
    szBus = ""

    If nTemp > 0 Then
        rsTemp.MoveFirst

        ReDim szBusID(1 To nTemp)

        szBus = szBus & " AND ( bfo.bus_id IN('" & rsTemp!Bus_id & "'"
        szBusID(1) = rsTemp!Bus_id
        rsTemp.MoveNext
        For n = 2 To nTemp
            szBus = szBus & ",'" & rsTemp!Bus_id & "'"
            szBusID(n) = rsTemp!Bus_id
            rsTemp.MoveNext
        Next
        szBus = szBus & ")) "
    Else
        szBus = szBus & " AND  bfo.bus_id IN('')"
        szTemp = "''"
    End If

    For n = 1 To nTemp
        If n = nTemp Then
            szTemp = szTemp & "'" & ttFormula(n).szFormula & "'"
        Else
            szTemp = szTemp & "'" & ttFormula(n).szFormula & "',"
        End If
    Next
    '得到特殊票公式及票价项公式信息
    szSql = "SELECT rfo.route_id,bfo.bus_id,ticket_type_id,pfl.*,use_mark, tcf.parameter_1 AS p1,tcf.parameter_2 AS p2 " _
        & " FROM price_formula_detail_lst pfl,price_item_info pii,Ticket_type_formula_info tcf, " _
        & " Ticket_type_code tti,route_info rfo,Bus_info bfo " _
        & " WHERE  pfl.formula_name in (" & szTemp & ") AND  pfl.price_item=pii.price_item AND pfl.price_table_id=table_id AND " _
        & " bfo.route_id = rfo.route_id And tcf.price_item = pii.price_item And tcf.Ticket_type = tti.ticket_type_id " _
        & " AND tti.ticket_type_valid='" & TP_TicketTypeValid & "'  AND table_id='" & m_szPriceTableID & "' " _
        & szBus & " AND rfo.formula_name=pfl.formula_name " _
        & " ORDER BY rfo.route_id,bus_id,pii.price_item,ticket_type_id"
    Set rsTemp = oDB.Execute(szSql) 'odb.Execute(szSql, , , 0)

    If rsTemp.RecordCount Mod cnTotalBusPriceItem <> 0 Then ShowError ERR_FormulaNotRight
    '得到各途经站与计算票价相关的参数信息

    '此句SQL语句需要重点看一下,应该不需要使用到DISTINCT的.
    '***陈峰   2002-11-10
    szSql = "SELECT DISTINCT ssi.sell_station_name,bfo.bus_id ,route_name ,rsa.*,sfo.station_name,area_code,bus_start_time,bus_type," _
        & " vtc.vehicle_type_code,sti.seat_type_id, seat_type_name,ticket_type_id, ticket_type_name ,vtc.vehicle_type_short_name " _
        & " FROM Route_section_lst rsa,station_info sfo,bus_info bfo,Vehicle_info tvi," _
        & " bus_vehicle_code bva,Seat_type_code sti, Ticket_type_code tti,Route_info rfo,vehicle_type_code vtc,sell_station_info ssi " _
        & " WHERE rsa.end_station_id=sfo.station_id AND station_type=1 AND " _
        & " bfo.route_id=rsa.route_id  AND bfo.bus_id=bva.bus_id AND " _
        & " ticket_type_valid='" & TP_TicketTypeValid & "' " & szBus & szCondition _
        & " AND rsa.sell_station_id=ssi.sell_station_id" _
        & " AND rfo.route_id=rsa.route_id  AND tti.ticket_type_id<>'" & TP_FreeTicket & "'" _
        & " AND bva.vehicle_id=tvi.vehicle_id AND bfo.bus_id=bva.bus_id " _
        & " ORDER BY bfo.bus_id,rsa.route_id,vtc.vehicle_type_code,rsa.section_serial, " _
        & " sti.seat_type_id,ticket_type_id "
    Set rsStation = oDB.Execute(szSql) 'odb.Execute(szSql, , , 0)
    nStationCount = rsStation.RecordCount

    If oCaculator Is Nothing Then
        Set oCaculator = CreateObject("STPriLib.PriceItemFunLib")
        oCaculator.Init m_oActiveUser
    End If

    If nStationCount > 0 Then ReDim aPriceDetail(1 To nStationCount)
    i = 0

    szBusOld = ""
    nRecPosition = 0

    Do While Not rsStation.EOF
        If szBusOld <> rsStation!Bus_id Then
            szBusOld = rsStation!Bus_id
            oCaculator.Identify szBusOld
        End If

        i = i + 1
        '得到外围的信息
        aPriceDetail(i).szAreaCode = Trim(rsStation!Area_Code)
        aPriceDetail(i).szRouteID = Trim(rsStation!route_id)
        aPriceDetail(i).szRouteName = Trim(rsStation!route_name)
        aPriceDetail(i).szBusID = Trim(rsStation!Bus_id)
        aPriceDetail(i).dyStartTime = Format(rsStation!bus_start_time, "HH:MM:SS")
        aPriceDetail(i).nBusType = rsStation!bus_type
        aPriceDetail(i).szSeatTypeID = Trim(rsStation!seat_type_id)
        aPriceDetail(i).szSeatTypeName = Trim(rsStation!seat_type_name)
        aPriceDetail(i).szVehicleModel = Trim(rsStation!vehicle_type_code)
        aPriceDetail(i).szVehicleModelName = Trim(rsStation!vehicle_type_short_name)
        aPriceDetail(i).szStationID = RTrim(rsStation!end_station_id)
        aPriceDetail(i).szStationName = RTrim(rsStation!station_name)
        aPriceDetail(i).sgMileage = Trim(rsStation!end_station_mileage)
        aPriceDetail(i).nTicketType = rsStation!ticket_type_id
        aPriceDetail(i).nSerialNo = rsStation!section_serial
        aPriceDetail(i).szSellStationID = FormatDbValue(rsStation!sell_station_id)
        aPriceDetail(i).szSellStationName = FormatDbValue(rsStation!sell_station_name)

        If rsStation!ticket_type_id = TP_FullPrice Then
            RowCount = 0
        Else
            RowCount = RowCount + 1
        End If

        If szBusOld = rsStation!Bus_id Then
            rsTemp.MoveFirst
            rsTemp.Move nRecPosition
        End If
        
        Do While Not rsTemp.EOF
            If rsTemp!route_id = rsStation!route_id And rsTemp!Bus_id = rsStation!Bus_id And (rsTemp!ticket_type_id = rsStation!ticket_type_id Or rsStation!ticket_type_id = 1) Then
                If RowCount = 0 Then
                    nRecPosition = rsTemp.AbsolutePosition - 1
                Else
                    'rsTemp.Move RowCount - 1
                End If
            Exit Do
            Else
                rsTemp.MoveNext
            End If
        Loop
        '得到基本运价计算公式
        szItemFormula = FormatDbValue(rsTemp!item_formula)
        '算出基本运价
        If RowCount = 0 Then
            sgBase = CallByName(oCaculator, szItemFormula, _
            VbMethod, Trim(FormatDbValue(rsStation!sell_station_id)), RTrim(rsStation!end_station_id), Trim(rsStation!vehicle_type_code), _
            Trim(rsStation!seat_type_id), _
            rsTemp!parameter_1, rsTemp!parameter_2, rsTemp!parameter_3, _
            rsTemp!parameter_4, rsTemp!parameter_5)
        End If
        '四舍五入精确到厘
        If RowCount = 0 Then
            aPriceDetail(i).sgBaseCarriage = Round(sgBase + 0.0001, 3)
        Else
            aPriceDetail(i).sgBaseCarriage = Round(sgBase * rsTemp!p1 + rsTemp!p2 + 0.0001, 3)
        End If
        '算出其他各使用票价项的票价
        For k = 1 To 15
            rsTemp.Move nUseTicket
            If rsTemp!use_mark <> 0 Then
                nItemIndex = CInt(rsTemp!price_item)
                szItemFormula = FormatDbValue(rsTemp!item_formula)
                If szItemFormula <> "" Then
                    If RowCount = 0 Then
                        '如果是全票
                        aPriceDetail(i).asgItem(nItemIndex) = Round(CallByName(oCaculator, szItemFormula, _
                        VbMethod, Trim(FormatDbValue(rsStation!sell_station_id)), RTrim(rsStation!end_station_id), Trim(rsStation!vehicle_type_code), _
                        Trim(rsStation!seat_type_id), sgBase, rsTemp!parameter_1, _
                        rsTemp!parameter_2, rsTemp!parameter_3, rsTemp!parameter_4, rsTemp!parameter_5) + 0.0001, 3)
                        nFullRow = i
                        aPriceDetail(i).asgItem(nItemIndex) = Round(aPriceDetail(i).asgItem(nItemIndex) + 0.0001, 3)
                    Else
                        '如果不是全票,将全票行的数据赋给   特殊票种行
                        aPriceDetail(i).asgItem(nItemIndex) = aPriceDetail(nFullRow).asgItem(nItemIndex)
                        '根据参数算出特殊票种的票价
                        aPriceDetail(i).asgItem(nItemIndex) = Round(aPriceDetail(i).asgItem(nItemIndex) * rsTemp!p1 + rsTemp!p2 + 0.0001, 3)
                    End If
                End If
            End If
        Next
        '计算总票价
        aPriceDetail(i).sgTotalPrice = aPriceDetail(i).sgBaseCarriage
        For k = 1 To 15
            aPriceDetail(i).sgTotalPrice = aPriceDetail(i).sgTotalPrice + aPriceDetail(i).asgItem(k)
        Next
        aPriceDetail(i).sgTotalPrice = Format(aPriceDetail(i).sgTotalPrice, "0.000")

        szBusOld = rsStation!Bus_id
        rsStation.MoveNext
    Loop

    MakeSpecifyBusPrice = aPriceDetail

    Set rsTemp = Nothing
    Set rsStation = Nothing
    Set oDB = Nothing

    Exit Function

Error_Handle:
    If err.Number = 438 Then
    ShowError ERR_TicketFunLibNotSpecifyFun, , FormatDbValue(rsTemp!item_formula)
    ElseIf err.Number = 11 Then
    ShowError ERR_BusSpeedNoExist
    Else
    If Not oCaculator Is Nothing Then
    szTemp = oCaculator.GetLastError()
    If szTemp <> "" Then
    err.Raise err.Number, , szTemp
    Else
    err.Raise err.Number
    End If
    Else
    err.Raise err.Number
    End If
    End If
    If Not rsTemp Is Nothing Then Set rsTemp = Nothing
    If Not rsStation Is Nothing Then Set rsStation = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
End Function


'生成车次票价,返回记录集
Public Function MakeSpecifyBusPriceRS(BusVehicleSeat() As TBusVehicleSeatType) As Recordset
    Static oCaculator As Object ' PriceItemFunLib
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    '    Dim aPriceDetail() As TBusPriceDetailInfo
    Dim nStationCount As Long
    Dim i As Long
    Dim rsStation As Recordset
    Dim sgBase As Double
    Dim nItemIndex As Integer
    Dim RowCount As Long
    Dim sgTemp As Single
    Dim szBus As String
    Dim szRouteOld As String
    Dim szBusOld As String
    Dim szItemFormula As String
    Dim nRecPosition As Long
    Dim szTemp As String
    Dim nTemp As Integer
    Dim n As Long
    Dim k As Long
    Dim szSeat As String
    
    Dim ttFormula() As TFormula
    Dim nUseTicket As Integer
    Dim szBusID() As String
    Dim szCondition As String
    Dim szConditionTemp As String
    Dim dbFullGeneralData As Double '全票数据
    
    Dim rsResult As New Recordset '存放手工生成的记录集
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_MakeSpecifyPrice
    
    On Error GoTo Error_Handle
    oDB.ConnectionString = GetConnectionStr(cszPriceMan)
    '得到所有的可用的非免票的票种
    Set rsTemp = oDB.Execute("SELECT * FROM Ticket_type_code WHERE ticket_type_valid='" & TP_TicketTypeValid & "' AND ticket_type_id<>'" & TP_FreeTicket & "'")
    nUseTicket = rsTemp.RecordCount - 1
    '得到查询条件
    szCondition = ""
    nTemp = ArrayLength(BusVehicleSeat)
    For i = 1 To nTemp
        If i = 1 Then
            szCondition = " AND ("
            szConditionTemp = " AND ("
        End If
        szCondition = szCondition & " (bfo.bus_id= '" & BusVehicleSeat(i).szBusID & "' and sti.seat_type_id='" & BusVehicleSeat(i).szSeatTypeID & "' and vtc.vehicle_type_code='" & BusVehicleSeat(i).szVehicleTypeCode & "') "
        szConditionTemp = szConditionTemp & " (bfo.bus_id= '" & BusVehicleSeat(i).szBusID & "' and ptl.seat_type_id='" & BusVehicleSeat(i).szSeatTypeID & "' and ptl.vehicle_type_code='" & BusVehicleSeat(i).szVehicleTypeCode & "') "
        If i = nTemp Then
            szCondition = szCondition & ") "
            szConditionTemp = szConditionTemp & ") "
        Else
            szCondition = szCondition & " OR "
            szConditionTemp = szConditionTemp & " OR "
        End If
    Next
    
    szBus = ""
    nTemp = ArrayLength(BusVehicleSeat)
    If nTemp > 0 Then
        szBus = szBus & " AND ( bfo.bus_id IN('" & BusVehicleSeat(1).szBusID & "'"
        For n = 2 To nTemp
            szBus = szBus & ",'" & BusVehicleSeat(n).szBusID & "'"
        Next
        szBus = szBus & ")) "
    End If
    '得到线路的所用的票价公式
    szSql = "SELECT formula_name FROM route_info WHERE formula_name=''"
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        szSql = "SELECT formula_name FROM Price_formula_info WHERE default_execute_mark=1"
        Set rsTemp = oDB.Execute(szSql)
        If rsTemp.RecordCount <> 1 Then ShowError ERR_SpecifyRouteNoFormulaAndNoDefaultFormula
        oDB.Execute ("UPDATE route_info SET  formula_name=(SELECT formula_name FROM Price_formula_info WHERE default_execute_mark=1) WHERE formula_name=''")
    End If
    szSql = "SELECT bus_id,formula_name FROM route_info rfo,bus_info bfo WHERE formula_name<>'" & cszDisCountFormula & "' AND rfo.route_id=bfo.route_id "
    szSql = szSql & szBus

    Set rsTemp = oDB.Execute(szSql)
    nTemp = rsTemp.RecordCount
    If nTemp > 0 Then
        ReDim ttFormula(1 To nTemp)
    End If

    i = 1
    Do While Not rsTemp.EOF
        ttFormula(i).szBusID = rsTemp!Bus_id
        ttFormula(i).szFormula = FormatDbValue(rsTemp!formula_name)
        i = i + 1
        rsTemp.MoveNext
    Loop
    szTemp = ""
    szBus = ""
    
    If nTemp > 0 Then
        rsTemp.MoveFirst
        
        ReDim szBusID(1 To nTemp)
        
        szBus = szBus & " AND ( bfo.bus_id IN('" & rsTemp!Bus_id & "'"
        szBusID(1) = rsTemp!Bus_id
        rsTemp.MoveNext
        For n = 2 To nTemp
            szBus = szBus & ",'" & rsTemp!Bus_id & "'"
            szBusID(n) = rsTemp!Bus_id
            rsTemp.MoveNext
        Next
        szBus = szBus & ")) "
    Else
        szBus = szBus & " AND  bfo.bus_id IN('')"
        szTemp = "''"
    End If

    For n = 1 To nTemp
        If n = nTemp Then
            szTemp = szTemp & "'" & ttFormula(n).szFormula & "'"
        Else
            szTemp = szTemp & "'" & ttFormula(n).szFormula & "',"
        End If
    Next
    '得到特殊票公式及票价项公式信息
'    szSql = "SELECT rfo.route_id,bfo.bus_id,ticket_type_id,pfl.*,use_mark, tcf.parameter_1 AS p1,tcf.parameter_2 AS p2 " _
'    & " FROM price_formula_detail_lst pfl,price_item_info pii,Ticket_type_formula_info tcf, " _
'    & " Ticket_type_code tti,route_info rfo,Bus_info bfo " _
'    & " WHERE  pfl.formula_name in (" & szTemp & ") AND  pfl.price_item=pii.price_item AND pfl.price_table_id=table_id AND " _
'    & " bfo.route_id = rfo.route_id And tcf.price_item = pii.price_item And tcf.Ticket_type = tti.ticket_type_id " _
'    & " AND tti.ticket_type_valid='" & TP_TicketTypeValid & "'  AND table_id='" & m_szPriceTableID & "' " _
'    & szBus & " AND rfo.formula_name=pfl.formula_name AND project_id='" & m_szProjectID & "'" _
'    & " ORDER BY rfo.route_id,bus_id,pii.price_item,ticket_type_id"
'    Set rsTemp = odb.Execute(szSql) 'odb.Execute(szSql, , , 0)
    
    szSql = "SELECT rfo.route_id,bfo.bus_id,ticket_type_id,pfl.*,use_mark, tcf.parameter_1 AS p1,tcf.parameter_2 AS p2 " _
        & " FROM price_formula_detail_lst pfl,price_item_info pii,Ticket_type_formula_info tcf, " _
        & " Ticket_type_code tti,route_info rfo,Bus_info bfo " _
        & " WHERE  pfl.formula_name in (" & szTemp & ") AND  pfl.price_item=pii.price_item AND pfl.price_table_id=table_id AND " _
        & " bfo.route_id = rfo.route_id And tcf.price_item = pii.price_item And tcf.Ticket_type = tti.ticket_type_id " _
        & " AND tti.ticket_type_valid='" & TP_TicketTypeValid & "' AND table_id='" & m_szPriceTableID & "' " _
        & szBus & " AND rfo.formula_name=pfl.formula_name " _
        & " ORDER BY bus_id,pii.price_item,ticket_type_id"
    Set rsTemp = oDB.Execute(szSql)
    
    
    If rsTemp.RecordCount Mod cnTotalBusPriceItem <> 0 Then ShowError ERR_FormulaNotRight
    '得到各途经站与计算票价相关的参数信息
    
    '此句SQL语句需要重点看一下,应该不需要使用到DISTINCT的.
    '***陈峰   2002-11-10
    szSql = "SELECT DISTINCT ssi.sell_station_name,bfo.bus_id ,route_name ,rsa.*,sfo.station_name,area_code,bus_start_time,bus_type," _
    & " vtc.vehicle_type_code,sti.seat_type_id, seat_type_name,ticket_type_id, ticket_type_name ,vtc.vehicle_type_short_name " _
    & " FROM Route_section_lst rsa,station_info sfo,bus_info bfo,Vehicle_info tvi," _
    & " bus_vehicle_code bva,Seat_type_code sti, Ticket_type_code tti,Route_info rfo,vehicle_type_code vtc,sell_station_info ssi " _
    & " WHERE rsa.end_station_id=sfo.station_id AND station_type=1 AND " _
    & " bfo.route_id=rsa.route_id  AND bfo.bus_id=bva.bus_id AND " _
    & " ticket_type_valid='" & TP_TicketTypeValid & "' " & szBus & szCondition _
    & " AND rsa.sell_station_id=ssi.sell_station_id" _
    & " AND rfo.route_id=rsa.route_id  AND tti.ticket_type_id<>'" & TP_FreeTicket & "'" _
    & " AND bva.vehicle_id=tvi.vehicle_id AND bfo.bus_id=bva.bus_id " _
    & " ORDER BY bfo.bus_id,rsa.sell_station_id,rsa.route_id,vtc.vehicle_type_code,rsa.section_serial, " _
    & " sti.seat_type_id,ticket_type_id "
     
'        szSql = "SELECT DISTINCT ssi.sell_station_name,bfo.bus_id ,rsa.*,sfo.station_name,area_code,bus_start_time,bus_type," _
'            & " bva.project_id , vtc.vehicle_type_code,sti.seat_type_id, ticket_type_id  " _
'            & " FROM Route_section_lst rsa,station_info sfo,bus_info bfo,vehicle_type_code vtc," _
'            & " bus_vehicle_code bva,Seat_type_code sti, Ticket_type_code tti ,Bus_station_info gbs ,sell_station_info ssi" _
'            & " WHERE rsa.end_station_id=sfo.station_id AND station_type=1 AND " _
'            & " bfo.route_id=rsa.route_id AND bva.project_id='" & m_szProjectID & "' AND bfo.bus_id=bva.bus_id AND " _
'            & " ticket_type_valid='" & TP_TicketTypeValid & "' " & szBus & szCondition _
'            & " AND tti.ticket_type_id<>'" & TP_FreeTicket & "'" _
'            & " AND rsa.sell_station_id=ssi.sell_station_id" _
'            & " AND bfo.bus_id=bva.bus_id AND bfo.project_id='" & m_szProjectID & "' " _
'            & " and gbs.station_id=sfo.station_id and gbs.bus_id=bfo.bus_id and bva.project_id=gbs.project_id "
'
'        If StopStation = True Then szSql = szSql & " and max_sale_quantity<>0"
'        szSql = szSql & " ORDER BY rsa.sell_station_id,bfo.bus_id,vtc.vehicle_type_code,rsa.section_serial, " _
'            & " sti.seat_type_id,ticket_type_id "
     
    
    Set rsStation = oDB.Execute(szSql) 'odb.Execute(szSql, , , 0)
    nStationCount = rsStation.RecordCount
    
    If oCaculator Is Nothing Then
        Set oCaculator = CreateObject("STPriLib.PriceItemFunLib")
        oCaculator.Init m_oActiveUser
    End If
    '初始化结果记录集
    rsResult.Fields.Append "Area_Code", adChar, 6
    rsResult.Fields.Append "route_id", adChar, 4
    rsResult.Fields.Append "route_name", adChar, 16
    rsResult.Fields.Append "Bus_id", adChar, 5
    rsResult.Fields.Append "bus_start_time", adDate
    rsResult.Fields.Append "bus_type", adSmallInt
    rsResult.Fields.Append "seat_type_id", adChar, 3
    rsResult.Fields.Append "seat_type_name", adChar, 10
    rsResult.Fields.Append "vehicle_type_code", adChar, 3
    rsResult.Fields.Append "vehicle_type_short_name", adVarChar, 10
    rsResult.Fields.Append "station_id", adChar, 9
    rsResult.Fields.Append "station_name", adVarChar, 10
    rsResult.Fields.Append "mileage", adDouble
    rsResult.Fields.Append "ticket_type", adSmallInt
    rsResult.Fields.Append "ticket_type_name", adChar, 12
    rsResult.Fields.Append "station_serial_no", adSmallInt
    rsResult.Fields.Append "base_carriage", adDouble '基本运价
    rsResult.Fields.Append "sell_station_id", adChar, 9
    rsResult.Fields.Append "sell_station_name", adVarChar, 10
    
    For i = 1 To 15
        rsResult.Fields.Append "price_item_" & i, adDouble '各票价项
    Next i
    rsResult.Open
    
    i = 0
    szBusOld = ""
    nRecPosition = 0
    Do While Not rsStation.EOF
        If szBusOld <> rsStation!Bus_id Then
            szBusOld = rsStation!Bus_id
            oCaculator.Identify szBusOld
        End If
        rsResult.AddNew
        i = i + 1
        '得到外围的信息
        rsResult!Area_Code = FormatDbValue(rsStation!Area_Code)
        rsResult!route_id = FormatDbValue(rsStation!route_id)
        rsResult!route_name = FormatDbValue(rsStation!route_name)
        rsResult!Bus_id = FormatDbValue(rsStation!Bus_id)
        rsResult!bus_start_time = Format(rsStation!bus_start_time, "HH:MM:SS")
        rsResult!bus_type = FormatDbValue(rsStation!bus_type)
        rsResult!seat_type_id = FormatDbValue(rsStation!seat_type_id)
        rsResult!seat_type_name = FormatDbValue(rsStation!seat_type_name)
        rsResult!vehicle_type_code = FormatDbValue(rsStation!vehicle_type_code)
        rsResult!vehicle_type_short_name = FormatDbValue(rsStation!vehicle_type_short_name)
        rsResult!station_id = FormatDbValue(rsStation!end_station_id)
        rsResult!station_name = FormatDbValue(rsStation!station_name)
        rsResult!Mileage = FormatDbValue(rsStation!end_station_mileage)
        rsResult!Ticket_Type = FormatDbValue(rsStation!ticket_type_id)
        rsResult!ticket_type_name = FormatDbValue(rsStation!ticket_type_name)
        rsResult!station_serial_no = FormatDbValue(rsStation!section_serial)
        rsResult!sell_station_id = FormatDbValue(rsStation!sell_station_id)
        rsResult!sell_station_name = FormatDbValue(rsStation!sell_station_name)
        
        If rsStation!ticket_type_id = TP_FullPrice Then
            RowCount = 0
        Else
            RowCount = RowCount + 1
        End If
        If szBusOld = rsStation!Bus_id Then
            rsTemp.MoveFirst
            rsTemp.Move nRecPosition
        End If
        Do While Not rsTemp.EOF
            If rsTemp!route_id = rsStation!route_id And rsTemp!Bus_id = rsStation!Bus_id And (rsTemp!ticket_type_id = rsStation!ticket_type_id Or rsStation!ticket_type_id = 1) Then
                If RowCount = 0 Then
                    nRecPosition = rsTemp.AbsolutePosition - 1
                Else
                    'rsTemp.Move RowCount - 1
                End If
            Exit Do
            Else
                rsTemp.MoveNext
            End If
        Loop
        
        
        '得到基本运价计算公式
        szItemFormula = FormatDbValue(rsTemp!item_formula)
        '算出基本运价
        If RowCount = 0 Then
            sgBase = CallByName(oCaculator, szItemFormula, _
            VbMethod, Trim(FormatDbValue(rsStation!sell_station_id)), FormatDbValue(rsStation!end_station_id), FormatDbValue(rsStation!vehicle_type_code), _
            FormatDbValue(rsStation!seat_type_id), _
            rsTemp!parameter_1, rsTemp!parameter_2, rsTemp!parameter_3, _
            rsTemp!parameter_4, rsTemp!parameter_5)
        End If
        '四舍五入精确到厘
        If RowCount = 0 Then
            rsResult!base_carriage = Round(sgBase + 0.0001, 3)
        Else
            rsResult!base_carriage = Round(sgBase * rsTemp!p1 + rsTemp!p2 + 0.0001, 3)
        End If
        '算出其他各使用票价项的票价
        'rsTemp.MoveFirst
        For k = 1 To 15
            rsTemp.Move nUseTicket
            If rsTemp!use_mark <> 0 Then
                nItemIndex = CInt(rsTemp!price_item)
                szItemFormula = FormatDbValue(rsTemp!item_formula)
                If szItemFormula <> "" Then
                    If RowCount = 0 Then
                        '如果是全票
                        rsResult.Fields("price_item_" & nItemIndex).Value = Round(CallByName(oCaculator, szItemFormula, _
                            VbMethod, Trim(FormatDbValue(rsStation!sell_station_id)), FormatDbValue(rsStation!end_station_id), FormatDbValue(rsStation!vehicle_type_code), _
                            FormatDbValue(rsStation!seat_type_id), sgBase, rsTemp!parameter_1, _
                        rsTemp!parameter_2, rsTemp!parameter_3, rsTemp!parameter_4, rsTemp!parameter_5) + 0.0001, 3)
                        rsResult.Fields("price_item_" & nItemIndex).Value = Round(rsResult.Fields("price_item_" & nItemIndex).Value + 0.0001, 3)
'                        dbFullGeneralData = rsResult.Fields("price_item_" & nItemIndex).Value
                    Else
                        '如果不是全票,将全票行的数据赋给   特殊票种行
                        rsResult.Move -RowCount
                        dbFullGeneralData = rsResult.Fields("price_item_" & nItemIndex).Value
                        rsResult.Move RowCount
                        '根据参数算出特殊票种的票价
                        rsResult.Fields("price_item_" & nItemIndex).Value = Round(dbFullGeneralData * rsTemp!p1 + rsTemp!p2 + 0.0001, 3)
                    End If
                End If
            End If
        Next
'        '计算总票价
'        aPriceDetail(i).sgTotalPrice = aPriceDetail(i).sgBaseCarriage
'        For k = 1 To 15
'            aPriceDetail(i).sgTotalPrice = aPriceDetail(i).sgTotalPrice + aPriceDetail(i).asgItem(k)
'        Next
'        aPriceDetail(i).sgTotalPrice = Format(aPriceDetail(i).sgTotalPrice, "0.000")
        szBusOld = rsStation!Bus_id
        rsStation.MoveNext
        rsResult.Update
    Loop
    Set MakeSpecifyBusPriceRS = rsResult
    Set rsTemp = Nothing
    Set rsStation = Nothing
    Set oDB = Nothing
    Set rsResult = Nothing
    Exit Function
Error_Handle:
    If err.Number = 438 Then
        ShowError ERR_TicketFunLibNotSpecifyFun, , FormatDbValue(rsTemp!item_formula)
    ElseIf err.Number = 11 Then
        ShowError ERR_BusSpeedNoExist
    Else
        If Not oCaculator Is Nothing Then
            szTemp = oCaculator.GetLastError()
            If szTemp <> "" Then
                err.Raise err.Number, , szTemp
            Else
                err.Raise err.Number
            End If
        Else
            err.Raise err.Number
        End If
    End If
    If Not rsTemp Is Nothing Then Set rsTemp = Nothing
    If Not rsStation Is Nothing Then Set rsStation = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
End Function

'生成空的车次票价
Public Function MakeEmptyBusPrice(BusVehicleSeat() As TBusVehicleSeatType) As TBusPriceDetailInfo()
'    Dim szSql As String
'    Dim odb As New RTConnection
'    Dim aPriceDetail() As TBusPriceDetailInfo
'    Dim nStationCount As Long
'    Dim i As Long
'    Dim rsStation As Recordset
'    Dim k As Long
'    Dim nTemp As Integer
'    Dim szCondition As String
'
'    AssertObjIsValid
'    AssertHaveRight m_oActiveUser, RIGHT_MakeSpecifyPrice
'    On Error GoTo Error_Handle
'    odb.ConnectionString = GetConnectionStr(cszPriceMan)
'    '得到查询条件
'    szCondition = ""
'    nTemp = ArrayLength(BusVehicleSeat)
'    For i = 1 To nTemp
'        If i = 1 Then
'            szCondition = " AND ("
'        End If
'        szCondition = szCondition & " (bfo.bus_id= '" & BusVehicleSeat(i).szBusID & "' and sti.seat_type_id='" & BusVehicleSeat(i).szSeatTypeID & "'" _
'            & " AND vtc.vehicle_type_code='" & BusVehicleSeat(i).szVehicleTypeCode & "'" & " AND rsa.sell_station_id='" & BusVehicleSeat(i).szSellStationID & "')"
'        If i = nTemp Then
'            szCondition = szCondition & ") "
'        Else
'            szCondition = szCondition & " OR "
'        End If
'    Next
'    '得到各途经站与计算票价相关的参数信息
'    szSql = "SELECT DISTINCT ssi.sell_station_name,bfo.bus_id ,route_name ,rsa.*,sfo.station_name,area_code,bus_start_time,bus_type," _
'        & " bva.project_id , vtc.vehicle_type_code,sti.seat_type_id, seat_type_name,ticket_type_id, ticket_type_name ,vtc.vehicle_type_short_name " _
'        & " FROM Route_section_lst rsa,station_info sfo,bus_info bfo,Vehicle_info tvi," _
'        & " bus_vehicle_code bva,Seat_type_code sti, Ticket_type_code tti,Route_info rfo,vehicle_type_code vtc,sell_station_info ssi " _
'        & " WHERE rsa.end_station_id=sfo.station_id AND station_type=1 AND " _
'        & " bfo.route_id=rsa.route_id AND bva.project_id='" & m_szProjectID & "' AND bfo.bus_id=bva.bus_id AND " _
'        & " ticket_type_valid='" & TP_TicketTypeValid & "' " & szCondition _
'        & " AND rfo.route_id=rsa.route_id  AND tti.ticket_type_id<>'" & TP_FreeTicket & "'" _
'        & " AND rsa.sell_station_id=ssi.sell_station_id" _
'        & " AND bva.vehicle_id=tvi.vehicle_id AND bfo.bus_id=bva.bus_id AND bfo.project_id='" & m_szProjectID & "' " _
'        & " ORDER BY bva.project_id,rsa.route_id,bfo.bus_id,vtc.vehicle_type_code,rsa.section_serial, " _
'        & " sti.seat_type_id,ticket_type_id "
'    Set rsStation = odb.Execute(szSql) 'odb.Execute(szSql, , , 0)
'
'    nStationCount = rsStation.RecordCount
'    If nStationCount > 0 Then ReDim aPriceDetail(1 To nStationCount)
'    i = 0
'    Do While Not rsStation.EOF
'        i = i + 1
'        '返回明细信息
'        aPriceDetail(i).szAreaCode = Trim(rsStation!Area_Code)
'        aPriceDetail(i).szRouteID = Trim(rsStation!route_id)
'        aPriceDetail(i).szRouteName = Trim(rsStation!route_name)
'        aPriceDetail(i).szBusID = Trim(rsStation!Bus_id)
'        aPriceDetail(i).dyStartTime = Format(rsStation!bus_start_time, "HH:MM:SS")
'        aPriceDetail(i).nBusType = rsStation!bus_type
'        aPriceDetail(i).szSeatTypeID = Trim(rsStation!seat_type_id)
'        aPriceDetail(i).szSeatTypeName = Trim(rsStation!seat_type_name)
'        aPriceDetail(i).szVehicleModel = Trim(rsStation!vehicle_type_code)
'        aPriceDetail(i).szVehicleModelName = Trim(rsStation!vehicle_type_short_name)
'        aPriceDetail(i).szStationID = RTrim(rsStation!end_station_id)
'        aPriceDetail(i).szStationName = RTrim(rsStation!station_name)
'        aPriceDetail(i).sgMileage = Trim(rsStation!end_station_mileage)
'        aPriceDetail(i).nTicketType = rsStation!ticket_type_id
'        aPriceDetail(i).nSerialNo = rsStation!section_serial
'        aPriceDetail(i).sgBaseCarriage = 0
'        aPriceDetail(i).szSellStationID = Trim(rsStation!sell_station_id)
'        aPriceDetail(i).szSellStationName = Trim(rsStation!sell_station_name)
'        For k = 1 To 15
'            aPriceDetail(i).asgItem(k) = 0
'        Next
'
'        '计算总票价
'        aPriceDetail(i).sgTotalPrice = 0
'        rsStation.MoveNext
'    Loop
'
'    MakeEmptyBusPrice = aPriceDetail
'
'    Set rsStation = Nothing
'    Set odb = Nothing
'    Exit Function
'
'Error_Handle:
'    err.Raise err.Number
'
'    If Not rsStation Is Nothing Then Set rsStation = Nothing
'    If Not odb Is Nothing Then Set odb = Nothing
End Function

Public Function MakeEmptyBusPriceRS(BusVehicleSeat() As TBusVehicleSeatType) As Recordset
    '生成空白的票价
    '返回记录集
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim aPriceDetail() As TBusPriceDetailInfo
    Dim nStationCount As Long
    Dim i As Long
    Dim rsStation As Recordset
    Dim j As Long
    Dim nTemp As Integer
    Dim szCondition As String
    Dim rsResult As New Recordset
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_MakeSpecifyPrice
    
    On Error GoTo Error_Handle
    
    oDB.ConnectionString = GetConnectionStr(cszPriceMan)
    '得到查询条件
    szCondition = ""
    nTemp = ArrayLength(BusVehicleSeat)
    For i = 1 To nTemp
        If i = 1 Then
            szCondition = " AND ("
        End If
        szCondition = szCondition & " (bfo.bus_id= '" & BusVehicleSeat(i).szBusID & "' and sti.seat_type_id='" & BusVehicleSeat(i).szSeatTypeID & "'" _
           & " AND vtc.vehicle_type_code='" & BusVehicleSeat(i).szVehicleTypeCode & "') "
        If i = nTemp Then
            szCondition = szCondition & ") "
        Else
            szCondition = szCondition & " OR "
        End If
    Next i
    '得到各途经站与计算票价相关的参数信息
    szSql = "SELECT DISTINCT ssi.sell_station_name,bfo.bus_id ,route_name ,rsa.*,sfo.station_name,area_code,bus_start_time,bus_type," _
        & "  vtc.vehicle_type_code,sti.seat_type_id, seat_type_name,ticket_type_id, ticket_type_name ,vtc.vehicle_type_short_name " _
        & " FROM Route_section_lst rsa,station_info sfo,bus_info bfo,Vehicle_info tvi," _
        & " bus_vehicle_code bva,Seat_type_code sti, Ticket_type_code tti,Route_info rfo,vehicle_type_code vtc,sell_station_info ssi " _
        & " WHERE rsa.end_station_id=sfo.station_id AND station_type=1 AND " _
        & " bfo.route_id=rsa.route_id AND  bfo.bus_id=bva.bus_id AND " _
        & " ticket_type_valid='" & TP_TicketTypeValid & "' " & szCondition _
        & " AND rfo.route_id=rsa.route_id  AND tti.ticket_type_id<>'" & TP_FreeTicket & "'" _
        & " AND rsa.sell_station_id=ssi.sell_station_id" _
        & " AND bva.vehicle_id=tvi.vehicle_id AND bfo.bus_id=bva.bus_id  " _
        & " ORDER BY bfo.bus_id,rsa.sell_station_id,rsa.route_id,vtc.vehicle_type_code,rsa.section_serial, " _
        & " sti.seat_type_id,ticket_type_id "
    Set rsStation = oDB.Execute(szSql) 'odb.Execute(szSql, , , 0)
    
    nStationCount = rsStation.RecordCount
    '初始化结果记录集
    rsResult.Fields.Append "Area_Code", adChar, 6
    rsResult.Fields.Append "route_id", adChar, 4
    rsResult.Fields.Append "route_name", adChar, 16
    rsResult.Fields.Append "Bus_id", adChar, 5
    rsResult.Fields.Append "bus_start_time", adDate
    rsResult.Fields.Append "bus_type", adSmallInt
    rsResult.Fields.Append "seat_type_id", adChar, 3
    rsResult.Fields.Append "seat_type_name", adChar, 10
    rsResult.Fields.Append "vehicle_type_code", adChar, 3
    rsResult.Fields.Append "vehicle_type_short_name", adVarChar, 10
    rsResult.Fields.Append "station_id", adChar, 9
    rsResult.Fields.Append "station_name", adVarChar, 10
    rsResult.Fields.Append "mileage", adDouble
    rsResult.Fields.Append "ticket_type", adSmallInt
    rsResult.Fields.Append "ticket_type_name", adChar, 12
    rsResult.Fields.Append "station_serial_no", adSmallInt
    rsResult.Fields.Append "base_carriage", adDouble '基本运价
    rsResult.Fields.Append "sell_station_id", adChar, 9
    rsResult.Fields.Append "sell_station_name", adVarChar, 10
    For i = 1 To 15
        rsResult.Fields.Append "price_item_" & i, adDouble '各票价项
    Next i
    rsResult.Open
    
    i = 0
    Do While Not rsStation.EOF
        rsResult.AddNew
        i = i + 1
        '返回明细信息
        rsResult!Area_Code = FormatDbValue(rsStation!Area_Code)
        rsResult!route_id = FormatDbValue(rsStation!route_id)
        rsResult!route_name = FormatDbValue(rsStation!route_name)
        rsResult!Bus_id = FormatDbValue(rsStation!Bus_id)
        rsResult!bus_start_time = Format(rsStation!bus_start_time, "HH:MM:SS")
        rsResult!bus_type = FormatDbValue(rsStation!bus_type)
        rsResult!seat_type_id = FormatDbValue(rsStation!seat_type_id)
        rsResult!seat_type_name = FormatDbValue(rsStation!seat_type_name)
        rsResult!vehicle_type_code = FormatDbValue(rsStation!vehicle_type_code)
        rsResult!vehicle_type_short_name = FormatDbValue(rsStation!vehicle_type_short_name)
        rsResult!station_id = FormatDbValue(rsStation!end_station_id)
        rsResult!station_name = FormatDbValue(rsStation!station_name)
        rsResult!Mileage = FormatDbValue(rsStation!end_station_mileage)
        rsResult!Ticket_Type = FormatDbValue(rsStation!ticket_type_id)
        rsResult!ticket_type_name = FormatDbValue(rsStation!ticket_type_name)
        rsResult!station_serial_no = FormatDbValue(rsStation!section_serial)
        rsResult!base_carriage = 0
        rsResult!sell_station_id = FormatDbValue(rsStation!sell_station_id)
        rsResult!sell_station_name = FormatDbValue(rsStation!sell_station_name)
        For j = 1 To 15
        rsResult.Fields("price_item_" & j).Value = 0
        Next j
        rsStation.MoveNext
        rsResult.Update
    Loop
    
    Set MakeEmptyBusPriceRS = rsResult
    
    Set rsStation = Nothing
    Set oDB = Nothing
    Set rsResult = Nothing
    Exit Function
    
Error_Handle:
    err.Raise err.Number
    Set rsStation = Nothing
    Set oDB = Nothing
End Function


'计算环境票价以数组形式传出
Public Function MakeEnvBusPrice(RouteId As String, BusID As String, VehicleID As String, BusStartTime As Date, CheckGate As String, StationID() As String) As TBusPriceDetailInfo()
    Static oCaculator As Object
    
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    Dim aPriceDetail() As TBusPriceDetailInfo
    Dim nStationCount As Long
    Dim i As Long
    Dim rsStation As Recordset
    Dim sgBase As Double
    Dim nItemIndex As Integer
    Dim RowCount As Long
    Dim sgTemp As Single
    
    Dim szItemFormula As String
    Dim szTemp As String
    Dim nTemp As Integer
    Dim n As Long
    Dim k As Long
    Dim szSeat As String
    
    Dim ttFormula() As TFormula
    Dim nUseTicket As Integer
    Dim nFullRow As Long
    Dim ttDealValue As TDealValue
    Dim nRow As Long
    Dim nAffectRow As Long
    Dim sgDealValue As Single
    Dim sgLastValue As Single
    Dim szitem As String
    Dim szItemValue As String
    Dim j As Long
    Dim cszPriceItem As String
    Dim szCondition As String '停靠站点条件和车辆座位条件
    Dim rsSection As Recordset
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_MakeSpecifyPrice
    
    On Error GoTo Error_Handle
    
    oDB.ConnectionString = GetConnectionStr(cszPriceMan)
    '***********************060930xxd,新增环境中列出一条路线路的多个站点
    Set rsSection = oDB.Execute("select * from Route_section_lst where route_id='" & RouteId & "'")
    nTemp = rsSection.RecordCount
    For i = 1 To nTemp
        If i = 1 Then szCondition = szCondition & " AND sfo.station_id in ("
        szCondition = szCondition & "'" & Trim(rsSection!end_station_id) & "'"
        If i = nTemp Then
            szCondition = szCondition & ")"
        Else
            szCondition = szCondition & ","
        End If
        rsSection.MoveNext
    Next
    '***********************
    '得到车辆的座位类型信息
    szSql = "SELECT vsi.*,tvi.start_seat_no as vehicle_start_seat_no,tvi.seat_quantity " _
        & " FROM vehicle_seat_type_info vsi,Vehicle_info tvi " _
        & " WHERE vsi.vehicle_id=tvi.vehicle_id and tvi.vehicle_id = '" & VehicleID & "'"
    Set rsTemp = oDB.Execute(szSql)
    j = 0
    If rsTemp.RecordCount > 0 Then
        '判断是否有没有所有的座位均设置了座位类型
        nTemp = rsTemp!seat_quantity
        For i = 1 To rsTemp.RecordCount
            If i = 1 Then szCondition = szCondition & " AND sti.seat_type_id in ("
            j = j + rsTemp!end_seat_no - rsTemp!start_seat_no + 1
            szCondition = szCondition & " '" & rsTemp!seat_type_id & "'"
            If nTemp > j Then
                szCondition = szCondition & ","
            Else
                szCondition = szCondition & ")"
            End If
            rsTemp.MoveNext
        Next
        If nTemp > j Then szCondition = szCondition & " '" & cszSeatTypeIsNormal & "')"
    Else
        szCondition = szCondition & " AND sti.seat_type_id IN ('" & cszSeatTypeIsNormal & "')"
    End If
    '得到所有的使用的非免票的票种
    Set rsTemp = oDB.Execute("SELECT * FROM Ticket_type_code WHERE ticket_type_valid='" & TP_TicketTypeValid & "' AND ticket_type_id<>'" & TP_FreeTicket & "'")
    nUseTicket = rsTemp.RecordCount - 1
    '得到线路的计算公式
    szSql = "SELECT formula_name FROM route_info WHERE Route_id = '" & RouteId & "' and formula_name=''"
    Set rsTemp = oDB.Execute(szSql)
    
    If rsTemp.RecordCount > 0 Then
        szSql = "SELECT formula_name FROM Price_formula_info WHERE default_execute_mark=1"
        Set rsTemp = oDB.Execute(szSql)
        If rsTemp.RecordCount <> 1 Then ShowError ERR_SpecifyRouteNoFormulaAndNoDefaultFormula
        oDB.Execute ("UPDATE route_info SET  formula_name=(SELECT formula_name FROM Price_formula_info WHERE default_execute_mark=1) WHERE Route_id = '" & RouteId & "'  AND formula_name=''")
    End If
    '得到所有计算公式不是"不计算"的线路
    szSql = "SELECT rfo.route_id,'" & BusID & "' as bus_id,formula_name FROM route_info rfo WHERE formula_name<>'" & cszDisCountFormula & "' AND rfo.route_id='" & RouteId & "' "
    
    Set rsTemp = oDB.Execute(szSql)
    
    nTemp = rsTemp.RecordCount
    If nTemp > 0 Then
        ReDim ttFormula(1 To 1)
        ttFormula(1).szBusID = rsTemp!Bus_id
        ttFormula(1).szFormula = Trim(rsTemp!formula_name)
    End If
    
    szTemp = ""
    szTemp = szTemp & "'" & ttFormula(1).szFormula & "'"
    '得到所有的特殊票计算公式
    szSql = "SELECT rfo.route_id,'" & BusID & "' as bus_id,ticket_type_id,pfl.*,use_mark, tcf.parameter_1 AS p1, " _
        & " tcf.parameter_2 AS p2 " _
        & " FROM price_formula_detail_lst pfl,price_item_info pii,Ticket_type_formula_info tcf, " _
        & " Ticket_type_code tti,route_info rfo " _
        & " WHERE  pfl.formula_name in (" & szTemp & ") AND  pfl.price_item=pii.price_item AND " _
        & " pfl.price_table_id=table_id AND rfo.route_id='" & RouteId & "' And " _
        & " tcf.price_item = pii.price_item And tcf.Ticket_Type = tti.ticket_type_id " _
        & " AND tti.ticket_type_valid='" & TP_TicketTypeValid & "' AND table_id='" & m_szPriceTableID & " ' AND rfo.formula_name=pfl.formula_name " _
        & " ORDER BY bus_id,pii.price_item,ticket_type_id"
    
    Set rsTemp = oDB.Execute(szSql)
    
    If rsTemp.RecordCount Mod cnTotalBusPriceItem <> 0 Then ShowError ERR_FormulaNotRight
    '得到各途经站与计算票价相关的参数信息
    szSql = " SELECT '" & BusID & "' as  bus_id ,ssi.sell_station_name,rsa.*,area_code,'" & ToDBTime(BusStartTime) & "' as bus_start_time, " _
        & " tvi.vehicle_type_code , sti.seat_type_id, ticket_type_id,station_name,seat_type_name,vehicle_type_short_name " _
        & " FROM Route_section_lst rsa,station_info sfo,Vehicle_info tvi,Seat_type_code sti, Ticket_type_code tti,vehicle_type_code vtc,sell_station_info ssi " _
        & " WHERE rsa.end_station_id=sfo.station_id AND station_type=1 AND rsa.route_id='" & RouteId & "' and vtc.vehicle_type_code=tvi.vehicle_type_code AND " _
        & " ticket_type_valid='" & TP_TicketTypeValid & "' AND tti.ticket_type_id<>'" & TP_FreeTicket & "' AND tvi.vehicle_id='" & VehicleID & "' " _
        & " AND rsa.sell_station_id=ssi.sell_station_id" _
        & szCondition _
        & " ORDER BY bfo.bus_id,tvi.vehicle_type_code,rsa.sell_station_id,rsa.section_serial, sti.seat_type_id,ticket_type_id "
    
    Set rsStation = oDB.Execute(szSql) 'odb.Execute(szSql, , , 0)
    nStationCount = rsStation.RecordCount
    '初始化对象
    If oCaculator Is Nothing Then
        Set oCaculator = CreateObject("STPriLib.PriceItemFunLib")
        oCaculator.Init m_oActiveUser
        oCaculator.EnvIdentify RouteId, BusID, BusStartTime, CheckGate, VehicleID
    End If
    
    '对每个车次生成和保存
    nStationCount = rsStation.RecordCount
    If nStationCount > 0 Then ReDim aPriceDetail(1 To nStationCount)
    i = 0
    If nStationCount > 0 Then
        rsStation.MoveFirst
        Do While Not rsStation.EOF
            i = i + 1
            '得到外围信息
            aPriceDetail(i).szAreaCode = Trim(rsStation!Area_Code)
            aPriceDetail(i).szRouteID = Trim(rsStation!route_id)
            aPriceDetail(i).szBusID = Trim(rsStation!Bus_id)
            aPriceDetail(i).dyStartTime = Format(rsStation!bus_start_time, "HH:MM:SS")
            aPriceDetail(i).szSeatTypeID = Trim(rsStation!seat_type_id)
            aPriceDetail(i).szSeatTypeName = rsStation!seat_type_name
            aPriceDetail(i).szVehicleModel = Trim(rsStation!vehicle_type_code)
            aPriceDetail(i).szVehicleModelName = rsStation!vehicle_type_short_name
            aPriceDetail(i).szStationID = RTrim(rsStation!end_station_id)
            aPriceDetail(i).szStationName = rsStation!station_name
            aPriceDetail(i).sgMileage = Trim(rsStation!end_station_mileage)
            aPriceDetail(i).nTicketType = rsStation!ticket_type_id
            aPriceDetail(i).nSerialNo = rsStation!section_serial
            aPriceDetail(i).szSellStationID = Trim(rsStation!sell_station_id)
            aPriceDetail(i).szSellStationName = Trim(rsStation!sell_station_name)
            
            If rsStation!ticket_type_id = TP_FullPrice Then
                RowCount = 0
            Else
                RowCount = RowCount + 1
            End If
        
            If RowCount = 0 Then
                rsTemp.MoveFirst
            Else
                rsTemp.MoveFirst
                rsTemp.Move RowCount - 1
            End If
            '得到基本运价项的计算公式
            szItemFormula = FormatDbValue(rsTemp!item_formula)
            If RowCount = 0 Then
                '计算基本运价
                sgBase = CallByName(oCaculator, szItemFormula, _
                VbMethod, Trim(rsStation!sell_station_id), RTrim(rsStation!end_station_id), Trim(rsStation!vehicle_type_code), _
                Trim(rsStation!seat_type_id), _
                rsTemp!parameter_1, rsTemp!parameter_2, rsTemp!parameter_3, _
                rsTemp!parameter_4, rsTemp!parameter_5)
            End If
            '精确到厘
            If RowCount = 0 Then
                aPriceDetail(i).sgBaseCarriage = Round(sgBase + 0.0001, 3)
            Else
                aPriceDetail(i).sgBaseCarriage = Round(sgBase * rsTemp!p1 + rsTemp!p2 + 0.0001, 3)
            End If
            '计算其他票价项的计算公式
            For k = 1 To 15
                rsTemp.Move nUseTicket
                If rsTemp!use_mark <> 0 Then
                    nItemIndex = CInt(rsTemp!price_item)
                    szItemFormula = FormatDbValue(rsTemp!item_formula)
                    If szItemFormula <> "" Then
                        If RowCount = 0 Then
                            aPriceDetail(i).asgItem(nItemIndex) = Round(CallByName(oCaculator, szItemFormula, _
                            VbMethod, Trim(rsStation!sell_station_id), RTrim(rsStation!end_station_id), Trim(rsStation!vehicle_type_code), _
                            Trim(rsStation!seat_type_id), sgBase, rsTemp!parameter_1, _
                            rsTemp!parameter_2, rsTemp!parameter_3, rsTemp!parameter_4, rsTemp!parameter_5) + 0.0001, 3)
                            nFullRow = i
                        Else
                            aPriceDetail(i).asgItem(nItemIndex) = aPriceDetail(nFullRow).asgItem(nItemIndex)
                        End If
                
                        If RowCount = 0 Then
                            aPriceDetail(i).asgItem(nItemIndex) = Round(aPriceDetail(i).asgItem(nItemIndex) + 0.0001, 3)
                        Else
                            aPriceDetail(i).asgItem(nItemIndex) = Round(aPriceDetail(i).asgItem(nItemIndex) * rsTemp!p1 + rsTemp!p2 + 0.0001, 3)
                        End If
                    End If
                End If
            Next
            
            If i = nStationCount Then Exit Do
            rsStation.MoveNext
        Loop
        
        '尾数处理
        
        k = ArrayLength(aPriceDetail)
        For i = 1 To k
            sgDealValue = aPriceDetail(i).sgBaseCarriage
            sgLastValue = aPriceDetail(i).asgItem(cnTotalBusPriceItem - 1)
            ttDealValue = GetColCarryValue(sgDealValue, aPriceDetail(i).szAreaCode, cszBaseCharge, sgLastValue, False, Trim(aPriceDetail(i).nBusType))
            aPriceDetail(i).sgBaseCarriage = ttDealValue.sgValue
            aPriceDetail(i).sgTotalPrice = Format(aPriceDetail(i).sgBaseCarriage, "0.00")
            For n = 1 To cnTotalBusPriceItem - 1
                cszPriceItem = n
                Do While Len(cszPriceItem) < 4
                    cszPriceItem = "0" & cszPriceItem
                Loop
                sgDealValue = aPriceDetail(i).asgItem(n)
                sgLastValue = aPriceDetail(i).asgItem(cnTotalBusPriceItem - 1)
                ttDealValue = GetColCarryValue(sgDealValue, aPriceDetail(i).szAreaCode, cszPriceItem, sgLastValue, False, Trim(aPriceDetail(i).nBusType))
                aPriceDetail(i).asgItem(n) = ttDealValue.sgValue
                aPriceDetail(i).sgTotalPrice = aPriceDetail(i).sgTotalPrice + Format(aPriceDetail(i).asgItem(n), "0.00")
            Next
            sgDealValue = aPriceDetail(i).sgTotalPrice
            sgLastValue = aPriceDetail(i).asgItem(cnTotalBusPriceItem - 1)
            ttDealValue = GetColCarryValue(sgDealValue, aPriceDetail(i).szAreaCode, cszBaseCharge, sgLastValue, True, Trim(aPriceDetail(i).nBusType))
            aPriceDetail(i).sgTotalPrice = Format(ttDealValue.sgValue, "0.00")
            aPriceDetail(i).asgItem(cnTotalBusPriceItem - 1) = ttDealValue.sgLastColValue
        Next
    End If
    
    Set rsTemp = Nothing
    Set rsStation = Nothing
    Set oDB = Nothing
    MakeEnvBusPrice = aPriceDetail
    Exit Function
    
Error_Handle:
    If err.Number = 11 Then
        err.Number = ERR_BusSpeedNoExist
    ElseIf err.Number = 438 Then
        err.Number = ERR_TicketFunLibNotSpecifyFun
    Else
        If Not oCaculator Is Nothing Then
        szTemp = oCaculator.GetLastError()
        If szTemp <> "" Then err.Description = szTemp
        End If
    End If
    err.Raise err.Number, "", err.Description
End Function

'计算后自动写入数据库
Public Sub MakeBusPrice(BusVehicleSeat() As TBusVehicleSeatType, Optional SELECTVehicleSeat As Boolean, Optional StopStation As Boolean)
    Static oCaculator As Object
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    Dim aPriceDetail() As TBusPriceDetailInfo
    Dim nStationCount As Long
    Dim i As Long
    Dim rsStation As Recordset
    Dim sgBase As Double
    Dim nItemIndex As Integer
    Dim RowCount As Long
    Dim sgTemp As Single
    Dim szBus As String
    Dim szBusOld As String
    Dim szItemFormula As String
    Dim nRecPosition As Long
    Dim szTemp As String
    Dim nTemp As Integer
    Dim n As Long
    Dim k As Long
    Dim szSeat As String
    
    Dim ttFormula() As TFormula
    Dim nUseTicket As Integer
    Dim szBusID() As String
    Dim szCondition As String
    Dim nFullRow As Long
    Dim ttDealValue As TDealValue
    Dim nRow As Long
    Dim nAffectRow As Long
    Dim sgDealValue As Single
    Dim sgLastValue As Single
    Dim szitem As String
    Dim szItemValue As String
    Dim nBusStartPosition() As Long '车次记录开始位置
    Dim j As Long
    Dim nValue As Long
    Dim cszPriceItem As String
    
    Dim szMakeBusPriceStatus As String '标志生成状态
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_MakeSpecifyPrice
    
    On Error GoTo Error_Handle
    
    oDB.ConnectionString = GetConnectionStr(cszPriceMan)
    
    Set rsTemp = oDB.Execute("SELECT * FROM Ticket_type_code WHERE ticket_type_valid='" & TP_TicketTypeValid & "' AND ticket_type_id<>'" & TP_FreeTicket & "'")
    nUseTicket = rsTemp.RecordCount - 1
    '得到查询条件
    szCondition = ""
    nTemp = ArrayLength(BusVehicleSeat)
    If SELECTVehicleSeat = False Then
        '如果选择座位类型
        For i = 1 To nTemp
            If i = 1 Then
                szCondition = " AND ("
            End If
            szCondition = szCondition & " (bfo.bus_id= '" & BusVehicleSeat(i).szBusID & "' and sti.seat_type_id='" & BusVehicleSeat(i).szSeatTypeID & "' and tvi.vehicle_type_code='" & BusVehicleSeat(i).szVehicleTypeCode & "') "
            If i = nTemp Then
                szCondition = szCondition & ") "
            Else
                szCondition = szCondition & " OR "
            End If
        Next
    Else
        For i = 1 To nTemp
            If i = 1 Then
                szCondition = " AND ("
            End If
            szCondition = szCondition & " (bfo.bus_id= '" & BusVehicleSeat(i).szBusID & "' and sti.seat_type_id='" & BusVehicleSeat(i).szSeatTypeID & "' and vtc.vehicle_type_code='" & BusVehicleSeat(i).szVehicleTypeCode & "') "
            If i = nTemp Then
                szCondition = szCondition & ") "
            Else
                szCondition = szCondition & " OR "
            End If
        Next
    End If
    
    szBus = ""
    nTemp = ArrayLength(BusVehicleSeat)
    If nTemp > 0 Then
        szBus = szBus & " AND ( bfo.bus_id IN('" & BusVehicleSeat(1).szBusID & "'"
        For n = 2 To nTemp
            szBus = szBus & ",'" & BusVehicleSeat(n).szBusID & "'"
        Next
        szBus = szBus & ")) "
    End If
    '得到线路公式
    szSql = "SELECT formula_name FROM route_info WHERE formula_name=''"
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        szSql = "SELECT formula_name FROM Price_formula_info WHERE default_execute_mark=1"
        Set rsTemp = oDB.Execute(szSql)
        If rsTemp.RecordCount <> 1 Then ShowError ERR_SpecifyRouteNoFormulaAndNoDefaultFormula
        oDB.Execute ("UPDATE route_info SET  formula_name=(SELECT formula_name FROM Price_formula_info WHERE default_execute_mark=1) WHERE formula_name=''")
    End If
    '得到车次对应的公式
    szSql = "SELECT rfo.route_id,bus_id,formula_name FROM route_info rfo,bus_info bfo WHERE formula_name<>'" & cszDisCountFormula & "' AND rfo.route_id=bfo.route_id "
    
    szSql = szSql & szBus

    Set rsTemp = oDB.Execute(szSql)
    
    nTemp = rsTemp.RecordCount
    If nTemp > 0 Then
        ReDim ttFormula(1 To nTemp)
    End If
    
    i = 1
    Do While Not rsTemp.EOF
        ttFormula(i).szBusID = rsTemp!Bus_id
        ttFormula(i).szFormula = FormatDbValue(rsTemp!formula_name)
        i = i + 1
        rsTemp.MoveNext
    Loop
    szTemp = ""
    
    szBus = ""
    
    If nTemp > 0 Then
        rsTemp.MoveFirst
        
        ReDim szBusID(1 To nTemp)
        
        szBus = szBus & " AND ( bfo.bus_id IN('" & rsTemp!Bus_id & "'"
        szBusID(1) = rsTemp!Bus_id
        rsTemp.MoveNext
        For n = 2 To nTemp
            szBus = szBus & ",'" & rsTemp!Bus_id & "'"
            szBusID(n) = rsTemp!Bus_id
            rsTemp.MoveNext
        Next
        szBus = szBus & ")) "
    Else
        szBus = szBus & " AND  bfo.bus_id IN('')"
        szTemp = "''"
    End If

    For n = 1 To nTemp
        If n = nTemp Then
            szTemp = szTemp & "'" & ttFormula(n).szFormula & "'"
        Else
            szTemp = szTemp & "'" & ttFormula(n).szFormula & "',"
        End If
    Next
    '得到特殊票计算公式
    szSql = "SELECT rfo.route_id,bfo.bus_id,ticket_type_id,pfl.*,use_mark, tcf.parameter_1 AS p1,tcf.parameter_2 AS p2 " _
        & " FROM price_formula_detail_lst pfl,price_item_info pii,Ticket_type_formula_info tcf, " _
        & " Ticket_type_code tti,route_info rfo,Bus_info bfo " _
        & " WHERE  pfl.formula_name in (" & szTemp & ") AND  pfl.price_item=pii.price_item AND pfl.price_table_id=table_id AND " _
        & " bfo.route_id = rfo.route_id And tcf.price_item = pii.price_item And tcf.Ticket_type = tti.ticket_type_id " _
        & " AND tti.ticket_type_valid='" & TP_TicketTypeValid & "' AND table_id='" & m_szPriceTableID & "' " _
        & szBus & " AND rfo.formula_name=pfl.formula_name " _
        & " ORDER BY bus_id,pii.price_item,ticket_type_id"
    Set rsTemp = oDB.Execute(szSql)

    If rsTemp.RecordCount Mod cnTotalBusPriceItem <> 0 Then ShowError ERR_FormulaNotRight
    '得到各途经站与计算票价相关的参数信息时，同时返回上车站的相关信息
    If SELECTVehicleSeat = False Then
        szSql = "SELECT DISTINCT ssi.sell_station_name,bfo.bus_id ,rsa.*,sfo.station_name,area_code,bus_start_time,bus_type," _
            & " tvi.vehicle_type_code,sti.seat_type_id, ticket_type_id  " _
            & " FROM Route_section_lst rsa,station_info sfo,bus_info bfo,Vehicle_info tvi," _
            & " bus_vehicle_code bva,Seat_type_code sti, Ticket_type_code tti,Bus_station_info gbs,sell_station_info ssi " _
            & " WHERE rsa.end_station_id=sfo.station_id AND station_type=1 AND " _
            & " bfo.route_id=rsa.route_id AND  bfo.bus_id=bva.bus_id AND " _
            & " ticket_type_valid='" & TP_TicketTypeValid & "' " & szBus & szCondition _
            & " AND tti.ticket_type_id<>'" & TP_FreeTicket & "'" _
            & " AND rsa.sell_station_id=ssi.sell_station_id" _
            & " AND bva.vehicle_id=tvi.vehicle_id AND bfo.bus_id=bva.bus_id " _
            & " and gbs.station_id=sfo.station_id and gbs.bus_id=bfo.bus_id "
    
        If StopStation = True Then szSql = szSql & " and max_sale_quantity<>0"
        szSql = szSql & " ORDER BY bfo.bus_id,rsa.sell_station_id,tvi.vehicle_type_code,rsa.section_serial, " _
            & " sti.seat_type_id,ticket_type_id "
    Else
        szSql = "SELECT DISTINCT ssi.sell_station_name,bfo.bus_id ,rsa.*,sfo.station_name,area_code,bus_start_time,bus_type," _
            & "  vtc.vehicle_type_code,sti.seat_type_id, ticket_type_id  " _
            & " FROM Route_section_lst rsa,station_info sfo,bus_info bfo,vehicle_type_code vtc," _
            & " bus_vehicle_code bva,Seat_type_code sti, Ticket_type_code tti ,Bus_station_info gbs ,sell_station_info ssi" _
            & " WHERE rsa.end_station_id=sfo.station_id AND station_type=1 AND " _
            & " bfo.route_id=rsa.route_id  AND bfo.bus_id=bva.bus_id AND " _
            & " ticket_type_valid='" & TP_TicketTypeValid & "' " & szBus & szCondition _
            & " AND tti.ticket_type_id<>'" & TP_FreeTicket & "'" _
            & " AND rsa.sell_station_id=ssi.sell_station_id" _
            & " AND bfo.bus_id=bva.bus_id " _
            & " and gbs.station_id=sfo.station_id and gbs.bus_id=bfo.bus_id "
        
        If StopStation = True Then szSql = szSql & " and max_sale_quantity<>0"
        szSql = szSql & " ORDER BY bfo.bus_id,rsa.sell_station_id,vtc.vehicle_type_code,rsa.section_serial, " _
            & " sti.seat_type_id,ticket_type_id "
    End If
        
    Set rsStation = oDB.Execute(szSql) 'odb.Execute(szSql, , , 0)
    nStationCount = rsStation.RecordCount
    If nStationCount > 0 Then RaiseEvent SetProgressRange(nStationCount)
    '初始化对象
    If oCaculator Is Nothing Then
        Set oCaculator = CreateObject("STPriLib.PriceItemFunLib")
        oCaculator.Init m_oActiveUser
    End If
    
    i = 0
    szBusOld = ""
    nRecPosition = 0
    
    '查找各车次在记录集rsStation中的开始位置
    nTemp = ArrayLength(szBusID)
    If nTemp >= 1 Then ReDim nBusStartPosition(1 To nTemp)
    For i = 1 To nTemp
    
        If nStationCount > 0 Then rsStation.MoveFirst
        rsStation.Find "bus_id='" & szBusID(i) & "'"
        If rsStation.AbsolutePosition > 0 Then nBusStartPosition(i) = rsStation.AbsolutePosition
    Next
    
    '对每个车次生成和保存
    
    For j = 1 To nTemp
nextBus:
        If j < nTemp Then
            nStationCount = nBusStartPosition(j + 1) - nBusStartPosition(j)
        Else
            nStationCount = rsStation.RecordCount - nBusStartPosition(j) + 1
        End If
        If nStationCount > 0 Then ReDim aPriceDetail(1 To nStationCount)
        i = 0
        
        If nBusStartPosition(j) > 0 Then
            Do While Not rsStation.EOF
                rsStation.MoveFirst
                rsStation.Move nBusStartPosition(j) - 1
                rsStation.Move i
                If szBusOld <> rsStation!Bus_id Then
                    szBusOld = rsStation!Bus_id
                    
                    oCaculator.Identify szBusOld
                End If
                '得到外围信息
                i = i + 1
                aPriceDetail(i).szAreaCode = Trim(rsStation!Area_Code)
                aPriceDetail(i).szRouteID = Trim(rsStation!route_id)
                aPriceDetail(i).szBusID = Trim(rsStation!Bus_id)
                aPriceDetail(i).dyStartTime = Format(rsStation!bus_start_time, "HH:MM:SS")
                aPriceDetail(i).nBusType = rsStation!bus_type
                aPriceDetail(i).szSeatTypeID = Trim(rsStation!seat_type_id)
                aPriceDetail(i).szVehicleModel = Trim(rsStation!vehicle_type_code)
                aPriceDetail(i).szStationID = RTrim(rsStation!end_station_id)
                aPriceDetail(i).sgMileage = Trim(rsStation!end_station_mileage)
                aPriceDetail(i).nTicketType = rsStation!ticket_type_id
                aPriceDetail(i).nSerialNo = rsStation!section_serial
                aPriceDetail(i).szSellStationID = Trim(rsStation!sell_station_id)
                aPriceDetail(i).szSellStationName = Trim(rsStation!sell_station_name)
        
                If rsStation!ticket_type_id = TP_FullPrice Then
                    RowCount = 0
                Else
                    RowCount = RowCount + 1
                End If
    
                If szBusOld = rsStation!Bus_id Then
                    rsTemp.MoveFirst
                    rsTemp.Move nRecPosition
                End If
                Do While Not rsTemp.EOF
                    If rsTemp!route_id = rsStation!route_id And rsTemp!Bus_id = rsStation!Bus_id Then
                        If RowCount = 0 Then
                            nRecPosition = rsTemp.AbsolutePosition - 1
                        Else
                            rsTemp.Move RowCount - 1
                        End If
                        Exit Do
                    Else
                        rsTemp.MoveNext
                    End If
                Loop
                '得到基本运价计算公式
                szItemFormula = FormatDbValue(rsTemp!item_formula)
                '计算基本运价票价
                If RowCount = 0 Then
                    sgBase = CallByName(oCaculator, szItemFormula, _
                    VbMethod, Trim(rsStation!sell_station_id), RTrim(rsStation!end_station_id), Trim(rsStation!vehicle_type_code), _
                    Trim(rsStation!seat_type_id), _
                    rsTemp!parameter_1, rsTemp!parameter_2, rsTemp!parameter_3, _
                    rsTemp!parameter_4, rsTemp!parameter_5)
                End If
    
                If RowCount = 0 Then
                    aPriceDetail(i).sgBaseCarriage = Round(sgBase + 0.0001, 3)
                Else
                    aPriceDetail(i).sgBaseCarriage = Round(sgBase * rsTemp!p1 + rsTemp!p2 + 0.0001, 3)
                End If
                '计算其他票价项的票价
'                rsTemp.MoveFirst
                For k = 1 To 15
                    rsTemp.Move nUseTicket
                    If rsTemp!use_mark <> 0 Then
                        nItemIndex = CInt(rsTemp!price_item)
                        szItemFormula = FormatDbValue(rsTemp!item_formula)
                        If szItemFormula <> "" Then
                            If RowCount = 0 Then
                                aPriceDetail(i).asgItem(nItemIndex) = Round(CallByName(oCaculator, szItemFormula, _
                                 VbMethod, Trim(rsStation!sell_station_id), RTrim(rsStation!end_station_id), Trim(rsStation!vehicle_type_code), _
                                 Trim(rsStation!seat_type_id), sgBase, rsTemp!parameter_1, _
                                 rsTemp!parameter_2, rsTemp!parameter_3, rsTemp!parameter_4, rsTemp!parameter_5) + 0.0001, 3)
                                 nFullRow = i
                            Else
                                aPriceDetail(i).asgItem(nItemIndex) = aPriceDetail(nFullRow).asgItem(nItemIndex)
                            End If
                        
                            If RowCount = 0 Then
                                aPriceDetail(i).asgItem(nItemIndex) = Round(aPriceDetail(i).asgItem(nItemIndex) + 0.0001, 3)
                            Else
                                aPriceDetail(i).asgItem(nItemIndex) = Round(aPriceDetail(i).asgItem(nItemIndex) * rsTemp!p1 + rsTemp!p2 + 0.0001, 3)
                            End If
                        End If
                    End If
                Next
                szBusOld = rsStation!Bus_id
                If i = nStationCount Then Exit Do
                rsStation.MoveNext
            Loop
            '尾数处理
            k = ArrayLength(aPriceDetail)
            For i = 1 To k
                sgDealValue = aPriceDetail(i).sgBaseCarriage
                sgLastValue = aPriceDetail(i).asgItem(cnTotalBusPriceItem - 1)
                ttDealValue = GetColCarryValue(sgDealValue, aPriceDetail(i).szAreaCode, cszBaseCharge, sgLastValue, False, Trim(aPriceDetail(i).nBusType))
                aPriceDetail(i).sgBaseCarriage = ttDealValue.sgValue
                aPriceDetail(i).sgTotalPrice = Format(aPriceDetail(i).sgBaseCarriage, "0.00")
                For n = 1 To cnTotalBusPriceItem - 1
                    cszPriceItem = n
                    Do While Len(cszPriceItem) < 4
                        cszPriceItem = "0" & cszPriceItem
                    Loop
                    sgDealValue = aPriceDetail(i).asgItem(n)
                    sgLastValue = aPriceDetail(i).asgItem(cnTotalBusPriceItem - 1)
                    ttDealValue = GetColCarryValue(sgDealValue, aPriceDetail(i).szAreaCode, cszPriceItem, sgLastValue, False, Trim(aPriceDetail(i).nBusType))
                    aPriceDetail(i).asgItem(n) = ttDealValue.sgValue
                    aPriceDetail(i).sgTotalPrice = aPriceDetail(i).sgTotalPrice + Format(aPriceDetail(i).asgItem(n), "0.00")
                Next
                sgDealValue = aPriceDetail(i).sgTotalPrice
                sgLastValue = aPriceDetail(i).asgItem(cnTotalBusPriceItem - 1)
                ttDealValue = GetColCarryValue(sgDealValue, aPriceDetail(i).szAreaCode, cszBaseCharge, sgLastValue, True, Trim(aPriceDetail(i).nBusType))
                aPriceDetail(i).sgTotalPrice = Format(ttDealValue.sgValue, "0.00")
                aPriceDetail(i).asgItem(cnTotalBusPriceItem - 1) = ttDealValue.sgLastColValue
            Next
            '写入数据库中
            nRow = ArrayLength(aPriceDetail)
            oDB.BeginTrans
            For i = 1 To nRow
                szitem = " price_item_1='" & Round(aPriceDetail(i).asgItem(1), 2) & "'," & " price_item_2='" & Round(aPriceDetail(i).asgItem(2), 2) & "'," & " price_item_3='" & Round(aPriceDetail(i).asgItem(3), 2) & "'," _
                    & " price_item_4='" & Round(aPriceDetail(i).asgItem(4), 2) & "'," & " price_item_5='" & Round(aPriceDetail(i).asgItem(5), 2) & "'," & " price_item_6='" & Round(aPriceDetail(i).asgItem(6), 2) & "'," _
                    & " price_item_7='" & Round(aPriceDetail(i).asgItem(7), 2) & "'," & " price_item_8='" & Round(aPriceDetail(i).asgItem(8), 2) & "'," & " price_item_9='" & Round(aPriceDetail(i).asgItem(9), 2) & "'," _
                    & " price_item_10='" & Round(aPriceDetail(i).asgItem(10), 2) & "'," & " price_item_11='" & Round(aPriceDetail(i).asgItem(11), 2) & "'," & " price_item_12='" & Round(aPriceDetail(i).asgItem(12), 2) & "'," _
                    & " price_item_13='" & Round(aPriceDetail(i).asgItem(13), 2) & "'," & " price_item_14='" & Round(aPriceDetail(i).asgItem(14), 2) & "'," & " price_item_15='" & Round(aPriceDetail(i).asgItem(15), 2) & "'"
'                If aPriceDetail(i).szSellStationID = "hq" Then
'                    Debug.Print 1
'                End If
                szSql = "UPDATE bus_price_lst SET " & "station_serial_no='" & aPriceDetail(i).nSerialNo & "', base_carriage='" & Round(aPriceDetail(i).sgBaseCarriage, 2) & "'," & szitem _
                    & ",mileage=" & aPriceDetail(i).sgMileage _
                    & " WHERE bus_id='" & aPriceDetail(i).szBusID & "' AND Station_id='" & aPriceDetail(i).szStationID & "'" _
                    & " AND price_table_id='" & m_szPriceTableID & "' AND ticket_type='" & aPriceDetail(i).nTicketType & "'" _
                    & " AND seat_type_id= '" & aPriceDetail(i).szSeatTypeID & "' AND vehicle_type_code='" & aPriceDetail(i).szVehicleModel & "'" _
                    & " AND sell_station_id=" & TransFieldValueToString(aPriceDetail(i).szSellStationID)
                oDB.Execute szSql, nAffectRow 'odb.Execute szSql, nAffectRow, , 0
                If nAffectRow < 1 Then
                
                    szitem = " price_item_1,price_item_2,price_item_3,price_item_4,price_item_5,price_item_6,price_item_7,price_item_8," _
                        & " price_item_9,price_item_10,price_item_11,price_item_12,price_item_13,price_item_14,price_item_15"
                    szItemValue = Round(aPriceDetail(i).asgItem(1), 2) & "," & Round(aPriceDetail(i).asgItem(2), 2) & "," & Round(aPriceDetail(i).asgItem(3), 2) & "," & Round(aPriceDetail(i).asgItem(4), 2) _
                        & "," & Round(aPriceDetail(i).asgItem(5), 2) & "," & Round(aPriceDetail(i).asgItem(6), 2) & "," & Round(aPriceDetail(i).asgItem(7), 2) & "," & Round(aPriceDetail(i).asgItem(8), 2) _
                        & "," & Round(aPriceDetail(i).asgItem(9), 2) & "," & Round(aPriceDetail(i).asgItem(10), 2) & "," & Round(aPriceDetail(i).asgItem(11), 2) & "," & Round(aPriceDetail(i).asgItem(12), 2) _
                        & "," & Round(aPriceDetail(i).asgItem(13), 2) & "," & Round(aPriceDetail(i).asgItem(14), 2) & "," & Round(aPriceDetail(i).asgItem(15), 2)
                    szSql = "INSERT INTO bus_price_lst (sell_station_id,bus_id,station_id,vehicle_type_code,price_table_id,ticket_type,seat_type_id,station_serial_no,mileage,base_carriage," & szitem & ") " _
                        & " VALUES ( '" & aPriceDetail(i).szSellStationID & "','" & aPriceDetail(i).szBusID & "','" & aPriceDetail(i).szStationID & "','" & aPriceDetail(i).szVehicleModel & "','" _
                        & m_szPriceTableID & "','" & aPriceDetail(i).nTicketType & "','" _
                        & aPriceDetail(i).szSeatTypeID & "'," _
                        & aPriceDetail(i).nSerialNo & "," & aPriceDetail(i).sgMileage & "," & Round(aPriceDetail(i).sgBaseCarriage, 2) & "," & szItemValue & ") "
                    oDB.Execute szSql 'odb.Execute szSql, , , 0
                End If
            Next
            oDB.CommitTrans
            nValue = nValue + nStationCount
            szMakeBusPriceStatus = "[" & szBusOld & "]" & "成功"
            RaiseEvent SetMakeBusPriceStatus(szMakeBusPriceStatus)
            RaiseEvent SetProgressValue(nValue)
            WriteOperateLog m_oActiveUser, RIGHT_MakeSpecifyPrice, "生成票价表［" & m_szPriceTableID & "］的[" & szBusOld & "]车次的票价"
            DoEvents
            If m_CancelMakeBusPrice = True Then
                szMakeBusPriceStatus = "退出车次票价生成"
                RaiseEvent SetMakeBusPriceStatus(szMakeBusPriceStatus)
                Exit For
            End If
        Else
            szMakeBusPriceStatus = "[" & szBusID(j) & "]" & "没有停靠站点，生成失败"
            RaiseEvent SetMakeBusPriceStatus(szMakeBusPriceStatus)
        End If
    Next
    
    Set rsTemp = Nothing
    Set rsStation = Nothing
    Set oDB = Nothing
    Exit Sub
    
Error_Handle:
    nValue = nValue + nStationCount
    RaiseEvent SetProgressValue(nValue)
    If err.Number = 11 Then
        err.Number = ERR_BusSpeedNoExist
    ElseIf err.Number = 438 Then
        err.Number = ERR_TicketFunLibNotSpecifyFun
    Else
        If Not oCaculator Is Nothing Then
            szTemp = oCaculator.GetLastError()
            If szTemp <> "" Then err.Description = szTemp
        End If
    End If
    szMakeBusPriceStatus = "[" & szBusOld & "]" & err.Description
    
    RaiseEvent SetMakeBusPriceStatus(szMakeBusPriceStatus)
    j = j + 1
    If m_CancelMakeBusPrice = True Then
        szMakeBusPriceStatus = "退出车次票价生成"
        RaiseEvent SetMakeBusPriceStatus(szMakeBusPriceStatus)
    Else
        If j <= nTemp Then Resume nextBus
    End If
End Sub

'修改车次票价
Public Sub ModifySpecifyBusPrice(PriceDetailInfo() As TBusPriceDetailInfo)
    Dim i As Long
    Dim nRow As Long
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim n As Integer
    Dim szitem As String
    Dim szItemValue As String
    Dim nAffectRow As Long
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_ModifySpecifyPrice
    On Error GoTo errorHander
    
    nRow = ArrayLength(PriceDetailInfo)
    oDB.ConnectionString = GetConnectionStr(cszPriceMan)
    oDB.BeginTrans
    For i = 1 To nRow
        '修改车次票价
        szitem = " price_item_1='" & PriceDetailInfo(i).asgItem(1) & "'," & " price_item_2='" & PriceDetailInfo(i).asgItem(2) & "'," & " price_item_3='" & PriceDetailInfo(i).asgItem(3) & "'," _
            & " price_item_4='" & PriceDetailInfo(i).asgItem(4) & "'," & " price_item_5='" & PriceDetailInfo(i).asgItem(5) & "'," & " price_item_6='" & PriceDetailInfo(i).asgItem(6) & "'," _
            & " price_item_7='" & PriceDetailInfo(i).asgItem(7) & "'," & " price_item_8='" & PriceDetailInfo(i).asgItem(8) & "'," & " price_item_9='" & PriceDetailInfo(i).asgItem(9) & "'," _
            & " price_item_10='" & PriceDetailInfo(i).asgItem(10) & "'," & " price_item_11='" & PriceDetailInfo(i).asgItem(11) & "'," & " price_item_12='" & PriceDetailInfo(i).asgItem(12) & "'," _
            & " price_item_13='" & PriceDetailInfo(i).asgItem(13) & "'," & " price_item_14='" & PriceDetailInfo(i).asgItem(14) & "'," & " price_item_15='" & PriceDetailInfo(i).asgItem(15) & "'"
        szSql = "UPDATE bus_price_lst SET base_carriage='" & PriceDetailInfo(i).sgBaseCarriage & "'," & szitem _
            & " WHERE bus_id='" & PriceDetailInfo(i).szBusID & "' AND Station_id='" & PriceDetailInfo(i).szStationID & "'" _
            & " AND price_table_id='" & m_szPriceTableID & "' AND ticket_type='" & PriceDetailInfo(i).nTicketType & "'" _
            & " AND sell_station_id='" & PriceDetailInfo(i).szSellStationID & "'" _
            & " AND seat_type_id= '" & PriceDetailInfo(i).szSeatTypeID & "' AND vehicle_type_code='" & PriceDetailInfo(i).szVehicleModel & "'"
        
        oDB.Execute szSql, nAffectRow
        If nAffectRow < 1 Then
            '如果不存在,则新增车次票价
            szitem = " price_item_1,price_item_2,price_item_3,price_item_4,price_item_5,price_item_6,price_item_7,price_item_8," _
                & " price_item_9,price_item_10,price_item_11,price_item_12,price_item_13,price_item_14,price_item_15"
            szItemValue = PriceDetailInfo(i).asgItem(1) & "," & PriceDetailInfo(i).asgItem(2) & "," & PriceDetailInfo(i).asgItem(3) & "," & PriceDetailInfo(i).asgItem(4) _
                & "," & PriceDetailInfo(i).asgItem(5) & "," & PriceDetailInfo(i).asgItem(6) & "," & PriceDetailInfo(i).asgItem(7) & "," & PriceDetailInfo(i).asgItem(8) _
                & "," & PriceDetailInfo(i).asgItem(9) & "," & PriceDetailInfo(i).asgItem(10) & "," & PriceDetailInfo(i).asgItem(11) & "," & PriceDetailInfo(i).asgItem(12) _
                & "," & PriceDetailInfo(i).asgItem(13) & "," & PriceDetailInfo(i).asgItem(14) & "," & PriceDetailInfo(i).asgItem(15)
            
            szSql = "INSERT INTO bus_price_lst (sell_station_id,bus_id,station_id,vehicle_type_code,price_table_id,ticket_type,seat_type_id,station_serial_no,mileage,base_carriage," & szitem & ") " _
                & " VALUES ( '" & PriceDetailInfo(i).szSellStationID & "','" & PriceDetailInfo(i).szBusID & "','" & PriceDetailInfo(i).szStationID & "','" & PriceDetailInfo(i).szVehicleModel & "','" _
                & m_szPriceTableID & "','" & PriceDetailInfo(i).nTicketType & "','" _
                & PriceDetailInfo(i).szSeatTypeID & "'," _
                & PriceDetailInfo(i).nSerialNo & "," & PriceDetailInfo(i).sgMileage & "," & PriceDetailInfo(i).sgBaseCarriage & "," & szItemValue & ") "
            oDB.Execute szSql
        End If
    Next
    oDB.CommitTrans
    WriteOperateLog m_oActiveUser, RIGHT_ModifySpecifyPrice, "增加或修改票价表[" & m_szPriceTableID & "]的部分车次票价"
    Set oDB = Nothing
    Exit Sub
    
errorHander:
    ShowError ERR_ModifyBusPrice
End Sub

'得到车次票价
Public Function GetSpecifyBusPrice(paszBusID() As String, paszVehicleModel() As String, paszSeatType() As String, paszSellStationID() As String) As TBusPriceDetailInfo()
Dim nSeatType As Integer
Dim szSeatType() As String
Dim i As Integer
Dim RowCount As Integer
Dim szTempSeat As String
Dim szTemp As String
Dim szTempVehicle As String
Dim nVehicleType As Integer
Dim szBus As String
Dim nTemp As Integer
Dim szSql As String
Dim aPriceDetail() As TBusPriceDetailInfo
Dim oDB As New RTConnection
Dim rsTemp As Recordset
Dim n As Integer
Dim rsStation As Recordset

On Error GoTo errorHander
oDB.ConnectionString = GetConnectionStr(cszPriceMan)

    nSeatType = ArrayLength(paszSeatType)
    szSeatType = paszSeatType

    szTempSeat = ""
    For i = 1 To nSeatType
        If i = 1 Then szTempSeat = " AND bpl.seat_type_id IN ("
        szTempSeat = szTempSeat & "'" & szSeatType(i) & "'"
        If i = nSeatType Then
           szTempSeat = szTempSeat & ")"
        Else
           szTempSeat = szTempSeat & ","
        End If
    Next

    szTempVehicle = ""
    nVehicleType = ArrayLength(paszVehicleModel)
    For i = 1 To nVehicleType
        If i = 1 Then szTempVehicle = " AND bpl.vehicle_type_code IN ("
        szTempVehicle = szTempVehicle & "'" & paszVehicleModel(i) & "'"
        If i = nVehicleType Then
           szTempVehicle = szTempVehicle & ")"
        Else
           szTempVehicle = szTempVehicle & ","
        End If
    Next

    szBus = ""
    nTemp = ArrayLength(paszBusID)
    If nTemp > 0 Then
       ReDim ttFormula(1 To nTemp)
       szBus = szBus & " (bpl.bus_id IN('" & paszBusID(1) & "'"
       For i = 2 To nTemp
           szBus = szBus & ",'" & paszBusID(i) & "'"
       Next
       szBus = szBus & ")) "
    End If
    
    '加入上车站查询条件
    Dim szTmpSellStation As String
    For i = 1 To ArrayLength(paszSellStationID)
        If i = 1 Then szTmpSellStation = " AND bpl.sell_station_id IN ("
        szTmpSellStation = szTmpSellStation & TransFieldValueToString(paszSellStationID(i))
        If i = ArrayLength(paszSellStationID) Then
            szTmpSellStation = szTmpSellStation & ")"
        Else
            szTmpSellStation = szTmpSellStation & ","
        End If
    Next i
        
    '查询票价信息
    szSql = "SELECT bpl.*,ssi.sell_station_name,ssi.sell_station_full_name,bfo.route_id,bfo.bus_start_time,rfo.route_name,vtc.vehicle_type_short_name," _
          & " sti.seat_type_name,sfo.station_name,bfo.bus_type,area_code " _
          & " FROM Bus_price_lst bpl,bus_info bfo,Route_info rfo,vehicle_type_code vtc,Seat_type_code sti," _
          & " Ticket_type_code tti,station_info sfo,sell_station_info ssi " _
          & " WHERE bpl.bus_id = bfo.bus_id AND rfo.route_id = bfo.route_id AND sti.seat_type_id = bpl.seat_type_id " _
          & " AND " & szBus & szTempSeat & szTempVehicle & szTmpSellStation _
          & " AND tti.ticket_type_id=bpl.ticket_type AND sfo.station_id=bpl.station_id " _
          & " AND bpl.sell_station_id=ssi.sell_station_id" _
          & " AND vtc.vehicle_type_code=bpl.vehicle_type_code AND bpl.price_table_id='" & m_szPriceTableID & "'" _
          & " ORDER BY bpl.sell_station_id,bpl.bus_id,bfo.route_id,bpl.vehicle_type_code,bpl.station_serial_no,bpl.station_id,bpl.seat_type_id," _
          & "bpl.Ticket_type "

        Set rsStation = oDB.Execute(szSql)
        If rsStation.RecordCount > 0 Then
           ReDim aPriceDetail(1 To rsStation.RecordCount)
        Else
           GoTo errorHander
        End If
        i = 0
        If rsStation.RecordCount > 0 Then
           Do While Not rsStation.EOF
                i = i + 1
                 aPriceDetail(i).szAreaCode = Trim(rsStation!Area_Code)
                 aPriceDetail(i).szRouteID = Trim(rsStation!route_id)
                 aPriceDetail(i).szRouteName = Trim(rsStation!route_name)
                 aPriceDetail(i).szBusID = Trim(rsStation!Bus_id)
                 aPriceDetail(i).dyStartTime = Format(rsStation!bus_start_time, "HH:MM:SS")
                 aPriceDetail(i).nBusType = rsStation!bus_type
                 aPriceDetail(i).szSeatTypeID = Trim(rsStation!seat_type_id)
                 aPriceDetail(i).szSeatTypeName = Trim(rsStation!seat_type_name)
                 aPriceDetail(i).szVehicleModel = Trim(rsStation!vehicle_type_code)
                 aPriceDetail(i).szVehicleModelName = Trim(rsStation!vehicle_type_short_name)
                 aPriceDetail(i).szStationID = RTrim(rsStation!station_id)
                 aPriceDetail(i).szStationName = RTrim(rsStation!station_name)
                 aPriceDetail(i).sgMileage = Trim(rsStation!Mileage)
                 aPriceDetail(i).nTicketType = rsStation!Ticket_Type
                 aPriceDetail(i).nSerialNo = rsStation!station_serial_no
                 aPriceDetail(i).sgBaseCarriage = rsStation!base_carriage
                 aPriceDetail(i).sgTotalPrice = rsStation!base_carriage
                 aPriceDetail(i).szSellStationID = Trim(rsStation!sell_station_id)
                 aPriceDetail(i).szSellStationName = Trim(rsStation!sell_station_name)
                                  
                 For n = 1 To 15
                     aPriceDetail(i).asgItem(n) = rsStation("price_item_" & n)
                     aPriceDetail(i).sgTotalPrice = aPriceDetail(i).sgTotalPrice + rsStation("price_item_" & n)
                 Next
                 rsStation.MoveNext
            Loop
        End If
        GetSpecifyBusPrice = aPriceDetail
    Exit Function
errorHander:
End Function
'直接修改票价，春运加价时用
Public Sub DirectModifyPrice(paszStationID() As String, panTicketType() As Integer, paszSeatType() As String, paszVehicleModel() As String, pbMantissa As Boolean, pdbRatio As Double, pszSourceItem As String, pszDesItem As String, pdbAdd As Double)
    Dim oDB As New RTConnection
    Dim szSql As String
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_ModifySpecifyPrice
    
    oDB.ConnectionString = GetConnectionStr(cszPriceMan)
    
    Dim i As Integer
    Dim szCondition As String
    
    '加入途经站类型条件
    If ArrayLength(paszStationID) > 0 Then
        For i = 1 To ArrayLength(paszStationID)
            If i = 1 Then szCondition = szCondition & " AND station_id IN ("
            If i = ArrayLength(paszStationID) Then
                szCondition = szCondition & TransFieldValueToString(paszStationID(i)) & ")"
            Else
                szCondition = szCondition & TransFieldValueToString(paszStationID(i)) & ","
            End If
        Next
    End If
        
    '加入车票类型条件
    If ArrayLength(panTicketType) > 0 Then
        For i = 1 To ArrayLength(panTicketType)
            If i = 1 Then szCondition = szCondition & " AND ticket_type IN ("
            If i = ArrayLength(panTicketType) Then
                szCondition = szCondition & TransFieldValueToString(panTicketType(i)) & ")"
            Else
                szCondition = szCondition & TransFieldValueToString(panTicketType(i)) & ","
            End If
        Next
    End If
    
    '加入座位类型
    If ArrayLength(paszSeatType) > 0 Then
        For i = 1 To ArrayLength(paszSeatType)
            If i = 1 Then szCondition = szCondition & " AND seat_type_id IN ("
            If i = ArrayLength(paszSeatType) Then
                szCondition = szCondition & TransFieldValueToString(paszSeatType(i)) & ")"
            Else
                szCondition = szCondition & TransFieldValueToString(paszSeatType(i)) & ","
            End If
        Next
    End If
        
    '加入车型
    If ArrayLength(paszVehicleModel) > 0 Then
        For i = 1 To ArrayLength(paszVehicleModel)
            If i = 1 Then szCondition = szCondition & " AND vehicle_type_code IN ("
            If i = ArrayLength(paszVehicleModel) Then
                szCondition = szCondition & TransFieldValueToString(paszVehicleModel(i)) & ")"
            Else
                szCondition = szCondition & TransFieldValueToString(paszVehicleModel(i)) & ","
            End If
        Next
    End If
    
    
'    szSql = "UPDATE bus_price_lst SET " & ConvertPriceItemName(Val(pszDesItem)) & "=" & ConvertPriceItemName(Val(pszSourceItem)) & "*" & pdbRatio & "+" & pdbAdd & _
'            " WHERE price_table_id=" & TransFieldValueToString(m_szPriceTableID) & szCondition
    
    szSql = "UPDATE bus_price_lst SET " & ConvertPriceItemName(Val(pszDesItem)) & "=" & ConvertPriceItemName(Val(pszSourceItem)) & "*" & pdbRatio & "+" & pdbAdd _
        & " FROM bus_price_lst p , route_info r , bus_info b " _
        & " WHERE p.bus_id = b.bus_id AND r.route_id = b.route_id AND r.formula_name <> '不计算'" _
        & " AND price_table_id=" & TransFieldValueToString(m_szPriceTableID) & szCondition

    
    
    oDB.Execute szSql
    
    Dim rsTmp As Recordset
    If pbMantissa Then  '尾数处理
        If szCondition = "" Then    '全部票价表进行尾数处理
            DealTableTail m_szPriceTableID
        Else
            szCondition = Mid(szCondition, 5, Len(szCondition) - 4)   'fule add
            szSql = "SELECT DISTINCT bus_id FROM bus_price_lst WHERE " & szCondition        '选中的车次进行处理
            Set rsTmp = oDB.Execute(szSql)
            While Not rsTmp.EOF
                DealTableTail m_szPriceTableID, FormatDbValue(rsTmp!Bus_id)
                rsTmp.MoveNext
            Wend
        End If
    End If
    
    
    WriteOperateLog m_oActiveUser, RIGHT_ModifySpecifyPrice, "直接修改票价表["" & m_szPriceTableID & ""]的部分车次票价"
    Set oDB = Nothing
    Exit Sub
    
errorHander:
    ShowError ERR_ModifyBusPrice
End Sub
Private Function ConvertPriceItemName(pnItemNo As Integer) As String
    If pnItemNo = 0 Then
        ConvertPriceItemName = "base_carriage"
    Else
        ConvertPriceItemName = "price_item_" & pnItemNo
    End If
End Function

'得到环境车次票价记录集
Public Function GetEnvBusPriceRS(pdtQueryDate As Date, paszBusID() As String) As Recordset
Dim nSeatType As Integer
Dim szSeatType() As String
Dim i As Integer
Dim RowCount As Integer
Dim szTempSeat As String
Dim szTemp As String
Dim szTempVehicle As String
Dim nVehicleType As Integer
Dim szBus As String
Dim nTemp As Integer
Dim szSql As String
Dim oDB As New RTConnection
Dim rsTemp As Recordset
Dim n As Integer

On Error GoTo errorHander
oDB.ConnectionString = GetConnectionStr(cszPriceMan)


    szBus = ""
    nTemp = ArrayLength(paszBusID)
    If nTemp > 0 Then
       ReDim ttFormula(1 To nTemp)
       szBus = szBus & " (bpl.bus_id IN('" & paszBusID(1) & "'"
       For i = 2 To nTemp
           szBus = szBus & ",'" & paszBusID(i) & "'"
       Next
       szBus = szBus & ")) "
    End If
    
    szSql = "SELECT bpl.*,ssi.sell_station_name,bfo.route_id,bfo.bus_start_time,rfo.route_name ," _
          & " sti.seat_type_name,sfo.station_name,bfo.bus_type,area_code,tti.ticket_type_name " _
          & " FROM Env_bus_price_lst bpl,Work_env_bus_info bfo,Route_info rfo,Seat_type_code sti," _
          & " Ticket_type_code tti,station_info sfo,sell_station_info ssi " _
          & " WHERE bpl.bus_id = bfo.bus_id AND bpl.bus_date=bfo.bus_date AND rfo.route_id = bfo.route_id AND sti.seat_type_id = bpl.seat_type_id " _
          & " AND bpl.bus_date='" & ToDBDate(pdtQueryDate) & "' AND " & szBus _
          & " AND tti.ticket_type_id=bpl.ticket_type AND sfo.station_id=bpl.station_id " _
          & " AND bpl.sell_station_id=ssi.sell_station_id" _
          & " ORDER BY bpl.bus_date,bpl.bus_id,bfo.route_id,bpl.sell_station_id , bpl.station_serial_no,bpl.station_id,bpl.seat_type_id," _
          & "bpl.Ticket_type "

        Set GetEnvBusPriceRS = oDB.Execute(szSql)
    Exit Function
errorHander:
End Function

'得到车次票价记录集
Public Function GetSpecifyBusPriceRS(paszBusID() As String, paszVehicleModel() As String, paszSeatType() As String, paszSellStation() As String) As Recordset
Dim nSeatType As Integer
Dim szSeatType() As String
Dim i As Integer
Dim RowCount As Integer
Dim szTempSeat As String
Dim szTemp As String
Dim szTempVehicle As String
Dim nVehicleType As Integer
Dim szBus As String
Dim nTemp As Integer
Dim szSql As String
Dim oDB As New RTConnection
Dim rsTemp As Recordset
Dim n As Integer

On Error GoTo errorHander
oDB.ConnectionString = GetConnectionStr(cszPriceMan)

    nSeatType = ArrayLength(paszSeatType)
    szSeatType = paszSeatType

    szTempSeat = ""
    For i = 1 To nSeatType
        If i = 1 Then szTempSeat = " AND bpl.seat_type_id IN ("
        szTempSeat = szTempSeat & "'" & szSeatType(i) & "'"
        If i = nSeatType Then
           szTempSeat = szTempSeat & ")"
        Else
           szTempSeat = szTempSeat & ","
        End If
    Next

    szTempVehicle = ""
    nVehicleType = ArrayLength(paszVehicleModel)
    For i = 1 To nVehicleType
        If i = 1 Then szTempVehicle = " AND bpl.vehicle_type_code IN ("
        szTempVehicle = szTempVehicle & "'" & paszVehicleModel(i) & "'"
        If i = nVehicleType Then
           szTempVehicle = szTempVehicle & ")"
        Else
           szTempVehicle = szTempVehicle & ","
        End If
    Next

    szBus = ""
    nTemp = ArrayLength(paszBusID)
    If nTemp > 0 Then
       ReDim ttFormula(1 To nTemp)
       szBus = szBus & " (bpl.bus_id IN('" & paszBusID(1) & "'"
       For i = 2 To nTemp
           szBus = szBus & ",'" & paszBusID(i) & "'"
       Next
       szBus = szBus & ")) "
    End If
        
    Dim szSellStation As String
    nTemp = ArrayLength(paszSellStation)
    If nTemp > 0 Then
        If paszSellStation(1) <> "" Then
            szSellStation = szSellStation & " AND bpl.sell_station_id IN('" & paszSellStation(1) & "'"
            For i = 2 To nTemp
                szSellStation = szSellStation & ",'" & paszSellStation(i) & "'"
            Next
            szSellStation = szSellStation & ")"
        End If
    End If
    
    szSql = "SELECT bpl.*,bfo.route_id,CONVERT(char(5),bfo.bus_start_time,108) bus_start_time,rfo.route_name ,vtc.vehicle_type_short_name," _
          & " sti.seat_type_name,sfo.station_name,bfo.bus_type,area_code,tti.ticket_type_name , s.sell_station_name ,s.sell_station_full_name" _
          & " FROM Bus_price_lst bpl,bus_info bfo,Route_info rfo,vehicle_type_code vtc,Seat_type_code sti," _
          & " Ticket_type_code tti,station_info sfo , sell_station_info s " _
          & " WHERE bpl.bus_id = bfo.bus_id AND rfo.route_id = bfo.route_id AND sti.seat_type_id = bpl.seat_type_id " _
          & " AND " & szBus & szTempSeat & szTempVehicle & szSellStation _
          & " AND tti.ticket_type_id=bpl.ticket_type AND sfo.station_id=bpl.station_id  " _
          & " AND bpl.sell_station_id=s.sell_station_id" _
          & " AND vtc.vehicle_type_code=bpl.vehicle_type_code AND bpl.price_table_id='" & m_szPriceTableID & "'" _
          & " ORDER BY bpl.sell_station_id,bpl.bus_id,bfo.route_id,bpl.station_serial_no,bpl.station_id,bpl.vehicle_type_code,bpl.seat_type_id," _
          & "bpl.Ticket_type "

        Set GetSpecifyBusPriceRS = oDB.Execute(szSql)
    Exit Function
errorHander:
    ShowError err.Number, , err.Description
End Function


'以下是尾数处理

'处理某个数值的尾数
Public Function GetColCarryValue(ByVal psgValue As Single, szStationID As String, szPriceItem As String, sgLastColValue As Single, pbTotalCol As Boolean, Optional szBusType As String = nAllBusType) As TDealValue
    '得到某一列进位的值
    'psgValue 某一列的值, pbTotalCol 是否是总计列  szStationID 站点编号
    Dim ttDealValue As TDealValue
    On Error GoTo Error_Handle
    
    ttDealValue.pbTotalCol = pbTotalCol
    If pbTotalCol Then
        '如果是总计列则进行总计的四舍五入方法
        ttDealValue = m_oTicketPriceMan.TotalTailDeal(psgValue, szStationID, szBusType, szPriceItem, m_tTailDeal, m_tTailDealMethod, sgLastColValue)
    Else
        '如果是某一列则进行某一列的四舍五入方法
        ttDealValue.sgValue = m_oTicketPriceMan.DetailTailDeal(psgValue, szStationID, szBusType, szPriceItem, m_tTailDeal, m_tTailDealMethod)
        ttDealValue.pbTotalCol = False
        ttDealValue.sgLastColValue = sgLastColValue
    End If
    GetColCarryValue = ttDealValue
    Exit Function
    
Error_Handle:
End Function

'处理整个票价表的尾数，并更新数据库
Public Sub DealTableTail(PriceTableID As String, Optional BusID As String)
Dim szSql As String
Dim i As Long
Dim j As Integer
Dim aPriceDetail() As TBusPriceDetailInfo
Dim szitem As String
Dim ttDealValue As TDealValue
Dim sgDealValue As Single
Dim sgLastValue As Single
Dim nTemp As Long
Dim oDB As New RTConnection
Dim rsTemp As Recordset
Dim szTemp As String
Dim szBusID As String

On Error GoTo errorHander
oDB.ConnectionString = GetConnectionStr(cszSNPrice)


If BusID <> "" Then szBusID = szBusID & " AND  bfo.bus_id='" & BusID & "'"
szSql = "SELECT area_code,bus_type,bpl.*  FROM bus_price_lst bpl,bus_info bfo,station_info sfo " _
     & " WHERE sfo.station_id=bpl.station_id and bfo.bus_id=bpl.bus_id  AND price_table_id='" & PriceTableID & "'"
szSql = szSql & szBusID

    Set rsTemp = oDB.Execute(szSql)
    'odb.Execute(szSql, , , 0)
    nTemp = rsTemp.RecordCount
    
    If nTemp > 0 Then
        ReDim aPriceDetail(1 To nTemp)
    End If
    i = 0

    Do While Not rsTemp.EOF
         i = i + 1
         aPriceDetail(i).szAreaCode = Trim(rsTemp!Area_Code)
         aPriceDetail(i).szBusID = Trim(rsTemp!Bus_id)
         aPriceDetail(i).nBusType = rsTemp!bus_type
         aPriceDetail(i).szSeatTypeID = Trim(rsTemp!seat_type_id)
         aPriceDetail(i).szVehicleModel = Trim(rsTemp!vehicle_type_code)
         aPriceDetail(i).szStationID = RTrim(rsTemp!station_id)
         aPriceDetail(i).sgMileage = Trim(rsTemp!Mileage)
         aPriceDetail(i).nTicketType = rsTemp!Ticket_Type
         aPriceDetail(i).nSerialNo = rsTemp!station_serial_no
         aPriceDetail(i).sgBaseCarriage = rsTemp!base_carriage
         aPriceDetail(i).szSellStationID = rsTemp!sell_station_id
         For j = 1 To cnTotalBusPriceItem - 1
            aPriceDetail(i).asgItem(j) = rsTemp.Fields("price_item_" & j)
         Next
         rsTemp.MoveNext
      Loop

      For i = 1 To nTemp
          If aPriceDetail(i).sgBaseCarriage <> 0 Then
            sgDealValue = aPriceDetail(i).sgBaseCarriage
            sgLastValue = aPriceDetail(i).asgItem(cnTotalBusPriceItem - 1)
            ttDealValue = GetColCarryValue(sgDealValue, aPriceDetail(i).szAreaCode, cszBaseCharge, sgLastValue, False, Trim(aPriceDetail(i).nBusType))
            aPriceDetail(i).sgBaseCarriage = Format(ttDealValue.sgValue, "0.00")
            aPriceDetail(i).sgTotalPrice = Format(aPriceDetail(i).sgBaseCarriage, "0.00")
          End If
          For j = 1 To cnTotalBusPriceItem - 1
              szTemp = j
              Do While Len(szTemp) < 4
                 szTemp = "0" & szTemp
              Loop
              If aPriceDetail(i).asgItem(j) <> 0 Then
                sgDealValue = aPriceDetail(i).asgItem(j)
                sgLastValue = aPriceDetail(i).asgItem(cnTotalBusPriceItem - 1)
                ttDealValue = GetColCarryValue(sgDealValue, aPriceDetail(i).szAreaCode, szTemp, sgLastValue, False, Trim(aPriceDetail(i).nBusType))
                aPriceDetail(i).asgItem(j) = Format(ttDealValue.sgValue, "0.00")
                aPriceDetail(i).sgTotalPrice = aPriceDetail(i).sgTotalPrice + Format(aPriceDetail(i).asgItem(j), "0.00")
              End If
          Next
            sgDealValue = aPriceDetail(i).sgTotalPrice
            sgLastValue = aPriceDetail(i).asgItem(cnTotalBusPriceItem - 1)
            ttDealValue = GetColCarryValue(sgDealValue, aPriceDetail(i).szAreaCode, cszBaseCharge, sgLastValue, True, Trim(aPriceDetail(i).nBusType))
            aPriceDetail(i).sgTotalPrice = Format(ttDealValue.sgValue, "0.00")
            aPriceDetail(i).asgItem(cnTotalBusPriceItem - 1) = Format(ttDealValue.sgLastColValue, "0.00")
      Next

    oDB.BeginTrans

    For i = 1 To nTemp
         szitem = " price_item_1='" & Round(aPriceDetail(i).asgItem(1), 2) & "'," & " price_item_2='" & Round(aPriceDetail(i).asgItem(2), 2) & "'," & " price_item_3='" & Round(aPriceDetail(i).asgItem(3), 2) & "'," _
         & " price_item_4='" & Round(aPriceDetail(i).asgItem(4), 2) & "'," & " price_item_5='" & Round(aPriceDetail(i).asgItem(5), 2) & "'," & " price_item_6='" & Round(aPriceDetail(i).asgItem(6), 2) & "'," _
         & " price_item_7='" & Round(aPriceDetail(i).asgItem(7), 2) & "'," & " price_item_8='" & Round(aPriceDetail(i).asgItem(8), 2) & "'," & " price_item_9='" & Round(aPriceDetail(i).asgItem(9), 2) & "'," _
         & " price_item_10='" & Round(aPriceDetail(i).asgItem(10), 2) & "'," & " price_item_11='" & Round(aPriceDetail(i).asgItem(11), 2) & "'," & " price_item_12='" & Round(aPriceDetail(i).asgItem(12), 2) & "'," _
         & " price_item_13='" & Round(aPriceDetail(i).asgItem(13), 2) & "'," & " price_item_14='" & Round(aPriceDetail(i).asgItem(14), 2) & "'," & " price_item_15='" & Round(aPriceDetail(i).asgItem(15), 2) & "'"
        szSql = "UPDATE bus_price_lst SET base_carriage='" & Round(aPriceDetail(i).sgBaseCarriage, 2) & "'," & szitem _
                & " WHERE bus_id='" & aPriceDetail(i).szBusID & "' AND Station_id='" & aPriceDetail(i).szStationID & "'" _
                & " AND price_table_id='" & m_szPriceTableID & "' AND ticket_type='" & aPriceDetail(i).nTicketType & "'" _
                & " AND seat_type_id= '" & aPriceDetail(i).szSeatTypeID & "' AND vehicle_type_code='" & aPriceDetail(i).szVehicleModel & "'" _
                & " AND sell_station_id = '" & aPriceDetail(i).szSellStationID & "' "
                
        oDB.Execute szSql 'odb.Execute szSql, , , 0
    Next
    oDB.CommitTrans
    Set oDB = Nothing
    Exit Sub
    
errorHander:
   ShowError err.Number
End Sub

'同步车次的票价
'此接口与同步车次站点一起用.
'主要作用是,当新增站点后,保留原先的票价.
Public Sub RefreshBusPriceRS(BusVehicleSeat() As TBusVehicleSeatType)
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsPrice As Recordset
    Dim rsResult As Recordset
    Dim szBusID As String
    On Error GoTo ErrorHandle
    
    AssertObjIsValid
'    AssertHaveRight m_oActiveUser, RIGHT_BusStationManagement
     
    oDB.ConnectionString = GetConnectionStr
    If ArrayLength(BusVehicleSeat) = 0 Then Exit Sub
    szBusID = BusVehicleSeat(1).szBusID
    
    '如果存在#bus_price_lst,则删除它
    szSql = "if exists (select 1" _
        & " From sysobjects " _
        & " where  id = object_id('#bus_price_lst')" _
        & " and   type = 'U')" _
        & " drop table #bus_price_lst "
    oDB.Execute szSql
        
    oDB.BeginTrans
        
        '得到原来保存的车次的站点票价信息
        szSql = " SELECT * INTO #bus_price_lst FROM bus_price_lst " _
            & " WHERE bus_id =  " & TransFieldValueToString(szBusID) _
            & " AND price_table_id = " & TransFieldValueToString(m_szPriceTableID)
        oDB.Execute szSql
        
        '重新计算车次的站点票价
        MakeBusPrice BusVehicleSeat, True
        
        '把原先保存着的票价信息替换到新计算出来的票价上去。
        szSql = " UPDATE bus_price_lst " _
            & " SET " _
            & " base_carriage = b.base_carriage " _
            & " ,price_item_1 = b.price_item_1 " _
            & " ,price_item_2 = b.price_item_2 " _
            & " ,price_item_3 = b.price_item_3 " _
            & " ,price_item_4 = b.price_item_4 " _
            & " ,price_item_5 = b.price_item_5 " _
            & " ,price_item_6 = b.price_item_6 " _
            & " ,price_item_7 = b.price_item_7 " _
            & " ,price_item_8 = b.price_item_8 " _
            & " ,price_item_9 = b.price_item_9 " _
            & " ,price_item_10 = b.price_item_10 " _
            & " ,price_item_11 = b.price_item_11 " _
            & " ,price_item_12 = b.price_item_12 " _
            & " ,price_item_13 = b.price_item_13 " _
            & " ,price_item_14 = b.price_item_14 " _
            & " ,price_item_15 = b.price_item_15 " _
            & " FROM bus_price_lst a ,#bus_price_lst b  WHERE a.bus_id = b.bus_id AND a.price_table_id = b.price_table_id AND a.sell_station_id = b.sell_station_id AND a.station_id = b.station_id " _
            & " AND a.ticket_type = b.ticket_type AND a.vehicle_type_code = b.vehicle_type_code AND a.seat_type_id = b.seat_type_id"
        
        oDB.Execute szSql
        
    oDB.CommitTrans
    
    Exit Sub
ErrorHandle:
    oDB.RollbackTrans
    err.Raise err.Number, err.Source, err.Description

End Sub



