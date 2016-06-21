VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CarrySheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E0390670375"
'行包签发单
Option Base 0

Option Explicit





'##ModelId=3E04A79E0037

'受理单号
'##ModelId=3E03F3C702B7
Private mszSheetID As String

'托运方方式
Private mszAcceptType As String

'承运车次
'##ModelId=3E03F3C703A8
Private mszBusID As String

'车次日期
'##ModelId=3E03F3C703D0
Private mdtBusDate As Date

'发车时间
'##ModelId=3E03F7A50058
Private mdtBusStartOffTime As Date

'承运车辆代码
'##ModelId=3E03F3C702D5
Private mszVehicleID As String

'承运车辆牌号
'##ModelId=3E03F3C702F3
Private mszVehicleLicense As String

'参运公司代码
'##ModelId=3E03F3C7031B
Private mszCompanyID As String

'参运公司名称
'##ModelId=3E03F3C70339
Private mszCompanyName As String

'拆帐公司代码
'##ModelId=3E03F3C70361
Private mszSplitCompanyID As String

'拆帐公司名称
'##ModelId=3E03F3C80060
Private mszSplitCompanyName As String

'线路代码
'##ModelId=3E03F3C80088
Private mszRouteID As String

'线路名称
'##ModelId=3E03F3C80164
Private mszRouteName As String

'承运车型代码
'##ModelId=3E03F6C00045
Private mszVehicleTypeID As String

'承运车型名称
'##ModelId=3E03F6C00077
Private mszVehicleTypeName As String

'车主代码
'##ModelId=3E03F5CA02E0
Private mszBusOwnerID As String

'车主名称
'##ModelId=3E03F73A013A
Private mszBusOwnerName As String

'是否有效
'##ModelId=3E03F8260298
Private mbValid As Boolean

'拆算协议代码
'##ModelId=3E03F8630187
Private mszProtocolID As String

'拆算协议名称
'##ModelId=3E03F87F0138
Private mszProtocolName As String

'装卸工
'##ModelId=3E03F3C802B9
Private mszMoveWorker As String

'签发员
'##ModelId=3E03F3C802B9
Private mszOperatorID As String

'签发员姓名
'##ModelId=3E03FB0E00A3
Private mszOperatorName As String

'签发时间
'##ModelId=3E03F3C80287
Private mdtOperateTime As Date

'受理单数
'##ModelId=3E03F9E1017D
Private mnAcceptSheetNumber As Integer

'计重
'##ModelId=3E03F3C80010
Private mdbCalWeight As Double

'实重
'##ModelId=3E03F3C80038
Private mdbActWeight As Double

'件数
'##ModelId=3E03F3C800B0
Private mnNumber As Integer

'超重件数
'##ModelId=3E03F3C800E2
Private mnOverNumber As Integer

'签发单状态
'##ModelId=3E03F3C8010A
Private mnStatus As Integer

'签发单状态
'##ModelId=3E03F3C8013C
Private mszStatusString As String

'总运价
'##ModelId=3E03F3C802EB
Private mdbTotalPrice As Double

'运费项数组
'##ModelId=3E03F3C8031D
Private matPriceItems() As TLuggagePriceItem

Private m_dbPrintSettleRatio As Double '打印的应结费率
Private m_dbPrintSettlePrice As Double '打印的应结运费



'行包单集
'##ModelId=3E03FD8101D1
Private mszSellStationID As String
Private mAcceptSheetIDs() As String
Private m_oActiveUser As ActiveUser
Private m_nObjectStatus As EObjectStatus
Private mszaTemp() As String

'本模块私有定义部分
'==========================================================================
'错误号
Private Enum EErrCarrySheet
    ERR_CarrySheetWithout = ERR_LuggageCarrySheet + 21  '无此受理单     '30421开始
    ERR_CarrySheetLuggageExist = ERR_LuggageCarrySheet + 22 '此行包单以存在
    ERR_CarrySheetLuggageNoBus = ERR_LuggageCarrySheet + 23 '无此车次信息
    ERR_CarrySheetNoExist = ERR_LuggageCarrySheet + 24 '此签发单不存在
    ERR_AcceptSheetNoFit = ERR_LuggageCarrySheet + 25 '此受理单不合法,可能是已废或已退
End Enum

'权限号
Private Enum ERightCarrySeet
    RIGHT_CarrySheet = ERR_LuggageCarrySheet + cnMidRightBegin             '30500开始
    RIGHT_CarrySheetEdit = RIGHT_CarrySheet + cnMidRightStep * 1            '更改签发单
    RIGHT_CarrySheetDelete = RIGHT_CarrySheet + cnMidRightStep * 2          '删除签发单
End Enum

Public Property Let PriceItems(Value() As TLuggagePriceItem)
    Let matPriceItems = Value
End Property


'初始化活动对象
'##ModelId=3E03F3B9034D
'**************************************************
'Member Code:S1
'Brief Description:初始化对象
'Engineer:王建波
'Date Generated:2002/12/21
'Last Revision Date:
'**************************************************
Public Sub Init(poAUser As ActiveUser)
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szSqlField As String
    Dim szSqlContext As String
    Dim rsTemp As Recordset, i As Integer
    On Error GoTo InitErr
    Set m_oActiveUser = poAUser
    mszSellStationID = m_oActiveUser.SellStationID
    mszOperatorID = m_oActiveUser.UserID
    mszOperatorName = m_oActiveUser.UserName
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT * FROM luggage_price_item_info WHERE use_mark=0 "
    Set rsTemp = oDb.Execute(szSql)
    
    ReDim matPriceItems(1 To rsTemp.RecordCount) As TLuggagePriceItem
    For i = 1 To rsTemp.RecordCount
        matPriceItems(i).PriceValue = 0
    Next
     
    Exit Sub
InitErr:
    Call RaiseError(err.Number, "CarrySheet:Init Method", err.Description)
End Sub

'指定受理单
'##ModelId=3E03F3B9036B
Public Sub Identify(SheetID As String)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageAcceptSheet
    
    On Error GoTo IdentifyErr
       mszSheetID = SheetID
       RefreshMemoryInfo
    Exit Sub
IdentifyErr:
    Call RaiseError(err.Number, "CarrySheet:Identify Method", err.Description)
End Sub
Private Sub RefreshMemoryInfo() '按照对象现在的ID将数据库中相应的信息读至对象内
 On Error GoTo ErrHandle
    Dim rsTemp As Recordset
    Dim rstemp1 As Recordset
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim i As Integer
    Dim j As Integer
    Dim szSqlProt As String
    Dim rsTempProt As Recordset
    Dim rsTemp2 As Recordset
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT * " _
        & " FROM luggage_sheet_info  " _
        & " WHERE  sheet_id='" & mszSheetID & "'"
    
    Set rsTemp = oDb.Execute(szSql)
      If rsTemp.RecordCount = 1 Then
        mszAcceptType = GetLuggageTypeString(FormatDbValue(rsTemp!Accept_type))
        mszBusID = FormatDbValue(rsTemp!bus_id)
        mdtBusDate = FormatDbValue(rsTemp!bus_date)
        mdtBusStartOffTime = FormatDbValue(rsTemp!bus_start_time)
        mszVehicleID = FormatDbValue(rsTemp!vehicle_id)
        mszVehicleLicense = FormatDbValue(rsTemp!license_tag_no)
        mszCompanyID = FormatDbValue(rsTemp!transport_company_id)
        mszCompanyName = FormatDbValue(rsTemp!transport_company_short_name)
        mszSplitCompanyID = FormatDbValue(rsTemp!split_company_id)
        mszSplitCompanyName = FormatDbValue(rsTemp!splict_company_short_name)
        mszRouteID = FormatDbValue(rsTemp!route_id)
        mszRouteName = FormatDbValue(rsTemp!route_name)
        mszVehicleTypeID = FormatDbValue(rsTemp!vehicle_type_code)
        mszVehicleTypeName = FormatDbValue(rsTemp!vehicle_type_name)
        mszBusOwnerID = FormatDbValue(rsTemp!owner_id)
        mszBusOwnerName = FormatDbValue(rsTemp!owner_name)
        mbValid = FormatDbValue(rsTemp!valid_mark)
        mszProtocolID = FormatDbValue(rsTemp!protocol_id)
        mszProtocolName = FormatDbValue(rsTemp!protocol_name)
        mszMoveWorker = FormatDbValue(rsTemp!worker)
        mdtOperateTime = FormatDbValue(rsTemp!sheet_make_time)
        mnAcceptSheetNumber = FormatDbValue(rsTemp!luggage_number)
        mdbCalWeight = FormatDbValue(rsTemp!cal_weight)
        mdbActWeight = FormatDbValue(rsTemp!fact_weight)
        mnNumber = FormatDbValue(rsTemp!baggage_number)
        mnOverNumber = FormatDbValue(rsTemp!over_number)
        mnStatus = FormatDbValue(rsTemp!valid_mark)
        mdbTotalPrice = FormatDbValue(rsTemp!total_price)
                
        mszOperatorID = FormatDbValue(rsTemp!checker)
        m_dbPrintSettlePrice = FormatDbValue(rsTemp!print_settle_price)
        m_dbPrintSettleRatio = FormatDbValue(rsTemp!print_settle_ratio)
        
        '如果签发单车号不为空,车辆协议从协议表中取
        szSqlProt = "SELECT distinct p.protocol_id,p.protocol_name" _
        & " FROM luggage_sheet_info l,luggage_settle_protocol_info p,vehicle_lug_protocol_info v " _
        & " WHERE l.vehicle_id=v.vehicle_id AND p.protocol_id=v.protocol_id " _
        & " AND l.sheet_id='" & mszSheetID & "'"
        Set rsTempProt = oDb.Execute(szSqlProt)
        If rsTempProt.RecordCount <> 0 Then
            mszProtocolID = FormatDbValue(rsTempProt!protocol_id)
            mszProtocolName = FormatDbValue(rsTempProt!protocol_name)
        End If
        
        If mszOperatorID = m_oActiveUser.UserID Then
            mszOperatorName = m_oActiveUser.UserName
        Else
            szSql = "SELECT user_name FROM user_info WHERE user_id =" & TransFieldValueToString(mszOperatorID)
            Set rsTemp2 = oDb.Execute(szSql)
            mszOperatorName = FormatDbValue(rsTemp2!user_name)
        End If
        
      Else
        ShowError ERR_CarrySheetNoExist
      End If
      
    szSql = "SELECT * FROM luggage_price_item_info WHERE use_mark = 0 AND accept_type= " & TransFieldValueToString(GetLuggageTypeInt(mszAcceptType))
    Set rstemp1 = oDb.Execute(szSql)
    If rstemp1.RecordCount > 0 Then
      ReDim Preserve matPriceItems(1 To rstemp1.RecordCount) As TLuggagePriceItem
        For j = 1 To rstemp1.RecordCount
            matPriceItems(j).PriceID = FormatDbValue(rstemp1!charge_item)
            matPriceItems(j).PriceName = FormatDbValue(rstemp1!chinese_name)
            matPriceItems(j).PriceValue = FormatDbValue(rsTemp.Fields(Trim("price_item_" + Trim(Str(j)))))
            rstemp1.MoveNext
       Next j
    End If
    Set rsTemp = Nothing
    Set rstemp1 = Nothing
    
    Exit Sub
ErrHandle:
    Set rsTemp = Nothing
    Set rstemp1 = Nothing
    err.Raise err.Number
End Sub
'新增签发单
'##ModelId=3E03F3B90389
Public Sub AddNew()
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageCarrySheet
    
    On Error GoTo AddNewErr
    Dim atTempPriceItem() As TLuggagePriceItem
    Dim atTempAcceptID() As String
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
'    AssertActiveUserValid m_oActiveUser, ERR_LugReceive
'    AssertHaveRight m_oActiveUser, RIGHT_CarrySheet
    
    m_nObjectStatus = ST_AddObj
    
    mnStatus = 1
    
    mszBusID = ""
    mszBusOwnerID = ""
    mszBusOwnerName = ""
    mszCompanyID = ""
    mszCompanyName = ""
    
    mszCompanyName = ""
    mszProtocolID = ""
    mszProtocolName = ""
    mszMoveWorker = ""
    mszRouteID = ""
    mszRouteName = ""
'    mszSellStationID = ""
    mszSheetID = ""
    mszSplitCompanyID = ""
    mszSplitCompanyName = ""
    mszStatusString = ""
    mszVehicleID = ""
    mszVehicleLicense = ""
    mszVehicleTypeID = ""
    mszVehicleTypeName = ""
    
    mnAcceptSheetNumber = 0
    mnNumber = 0
    mnOverNumber = 0
    mdbActWeight = 0
    mdbCalWeight = 0
    mdbTotalPrice = 0
    mdtBusDate = Date
    mdtBusStartOffTime = Now
    mdtOperateTime = Now
    mAcceptSheetIDs = atTempAcceptID
    m_dbPrintSettlePrice = 0
    m_dbPrintSettleRatio = 0
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT * FROM luggage_price_item_info WHERE use_mark=0 "
    Set rsTemp = oDb.Execute(szSql)
    ReDim matPriceItems(1 To rsTemp.RecordCount) As TLuggagePriceItem
    Exit Sub
AddNewErr:
    Call RaiseError(err.Number, "CarrySheet:AddNew Method", err.Description)
End Sub

'更新签发单
'##ModelId=3E03F3B903A7
Public Sub Update()
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageCarrySheet
    
    On Error GoTo UpdateErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szSqlField As String
    
    Select Case m_nObjectStatus
    Case ST_AddObj
        AddObject
    Case ST_EditObj
        AssertHaveRight m_oActiveUser, RIGHT_CarrySheetEdit
        
        UpdateToDB
    Case ST_NormalObj
    Case Else
        ShowError ERR_NotAvailable
    End Select
    
    
'    oDb.ConnectionString = GetConnectionStr(cszLuggage)

    Exit Sub
UpdateErr:
    Call RaiseError(err.Number, "CarrySheet:Update Method", err.Description)
End Sub

Private Sub AddObject()

End Sub
Private Sub UpdateToDB()
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSqlField As String
    On Error GoTo UpdateErr
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT * FROM luggage_sheet_info WHERE luggage_id =" & TransFieldValueToString(mszSheetID)
    
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_CarrySheetWithout
    szSql = "UPDATE luggage_accept_info SET"
'        szSqlField = szSqlField & "sheet_id='" & mszSheetID & "',"
        szSqlField = szSqlField & "bus_id='" & mszBusID & "',"
        szSqlField = szSqlField & "bus_date='" & mdtBusDate & "',"
        szSqlField = szSqlField & "vehicle_id='" & mszVehicleID & "',"
        szSqlField = szSqlField & "license_tag_no='" & mszVehicleLicense & "',"
        szSqlField = szSqlField & "transport_company_id='" & mszCompanyID & "',"
        szSqlField = szSqlField & "transport_company_short_name='" & mszCompanyName & "',"
        szSqlField = szSqlField & "split_company_id='" & mszSplitCompanyID & "',"
        szSqlField = szSqlField & "splict_company_short_name='" & mszSplitCompanyName & "',"
        szSqlField = szSqlField & "route_id='" & mszRouteID & "',"
        szSqlField = szSqlField & "route_name='" & mszRouteName & "',"
        szSqlField = szSqlField & "vehicle_type_code='" & mszVehicleTypeID & "',"
        szSqlField = szSqlField & "vehicle_type_name='" & mszVehicleTypeName & "',"
        szSqlField = szSqlField & "owner_id='" & mszBusOwnerID & "',"
        szSqlField = szSqlField & "owner_name='" & mszBusOwnerName & "',"
        szSqlField = szSqlField & "bus_start_time='" & mdtBusStartOffTime & "',"
        szSqlField = szSqlField & "valid_mark=" & mbValid & ","
        szSqlField = szSqlField & "protocol_id='" & mszProtocolID & "',"
        szSqlField = szSqlField & "protocol_name='" & mszProtocolName & "',"
        szSqlField = szSqlField & "checker='" & mszOperatorID & "',"
        szSqlField = szSqlField & "worker='" & mszMoveWorker & "',"
        szSqlField = szSqlField & "sheet_make_time='" & mdtOperateTime & "',"
        szSqlField = szSqlField & "luggage_number=" & mnAcceptSheetNumber & ","
        szSqlField = szSqlField & "baggage_number=" & mnNumber & ","
        szSqlField = szSqlField & "cal_weight=" & mdbCalWeight & ","
        szSqlField = szSqlField & "fact_weight=" & mdbActWeight & ","
        szSqlField = szSqlField & "over_number=" & mnOverNumber & ","
        szSqlField = szSqlField & "total_price=" & mdbTotalPrice & ","
        szSqlField = szSqlField & "price_item_1=" & 0 & ","
        szSqlField = szSqlField & "price_item_2=" & 0 & ","
        szSqlField = szSqlField & "price_item_3=" & 0 & ","
        szSqlField = szSqlField & "price_item_4=" & 0 & ","
        szSqlField = szSqlField & "price_item_5=" & 0 & ","
        szSqlField = szSqlField & "price_item_6=" & 0 & ","
        szSqlField = szSqlField & "price_item_7=" & 0 & ","
        szSqlField = szSqlField & "price_item_8=" & 0 & ","
        szSqlField = szSqlField & "price_item_9=" & 0 & ","
        szSqlField = szSqlField & "price_item_10=" & 0 & ","
        szSqlField = szSqlField & "settlement_status=" & mnStatus & ","
    
    oDb.Execute (szSql)
    Set oDb = Nothing
    Exit Sub
UpdateErr:
    Call RaiseError(err.Number, "CarrySheet:Update Method", err.Description)
End Sub


Private Sub DeleteObject()
    
    On Error GoTo DeleteErr
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    oDb.Execute ("DELETE luggage_sheet_info WHERE sheet_id='" & mszSheetID & "'")
    Set oDb = Nothing
    
    Exit Sub
DeleteErr:
    Call RaiseError(err.Number, "CarrySheet:Delete Method", err.Description)
End Sub

'删除签发单
'##ModelId=3E03F3B903C5
Public Sub Delete()
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageCarrySheet
    AssertHaveRight m_oActiveUser, RIGHT_CarrySheetDelete
    
    On Error GoTo DeleteErr
    
'    AssertActiveUserValid m_oActiveUser, ERR_LugReceive
'    AssertHaveRight m_oActiveUser, RIGHT_CarrySheet
'
    DeleteObject
    m_nObjectStatus = ST_NotAvailable
    Exit Sub
DeleteErr:
    Call RaiseError(err.Number, "CarrySheet:Delete Method", err.Description)
End Sub

'得到当前车次的预装行包受理单
'返回受理单代码数组
'
'##ModelId=3E03F3B903E3
Public Function GetBusPreLoadLuggage(mAcceptType As Integer) As String()
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageCarrySheet
    
    On Error GoTo GetBusPreLoadLuggageErr
    Dim szaTemp() As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    Dim mnTempAcceptID() As String
    Dim szSql As String
    Dim i As Integer
    
'    AssertActiveUserValid m_oActiveUser, ERR_LugReceive
'    AssertHaveRight m_oActiveUser, RIGHT_CarrySheet
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
'    szSql = "SELECT DISTINCT c.luggage_id " _
'        & " FROM luggage_sheet_info l,luggage_carry_accept_lst c " _
'        & " WHERE l.bus_id='" & mszBusID & "' AND l.bus_date =" & TransFieldValueToString(Date) _
'        & " AND c.check_luggage_id = l.sheet_id "
    
    szSql = "SELECT DISTINCT l.*,p.* " _
        & " FROM luggage_accept_info l,luggage_price_lst p " _
        & " WHERE l.bus_id='" & mszBusID & "' AND l.bus_date =" & TransFieldValueToString(Date) _
        & " AND l.status = 0 AND l.luggage_id=p.luggage_id "
        
        
    If mAcceptType <> -1 Then
        szSql = szSql & " AND l.accept_type=" & TransFieldValueToString(mAcceptType)
    End If
    Set rsTemp = oDb.Execute(szSql)
    
    If rsTemp.RecordCount <> 0 Then
        ReDim mnTempAcceptID(1 To rsTemp.RecordCount)
        For i = 1 To rsTemp.RecordCount
            mnTempAcceptID(i) = FormatDbValue(rsTemp!luggage_id)
            rsTemp.MoveNext
        Next
        GetBusPreLoadLuggage = mnTempAcceptID
        
        
        
        Exit Function
    End If
    GetBusPreLoadLuggage = mAcceptSheetIDs
    Exit Function
GetBusPreLoadLuggageErr:
    Call RaiseError(err.Number, "CarrySheet:GetBusPreLoadLuggage Method", err.Description)
End Function

'汇总签发的行包单
'
'##ModelId=3E03F3BA001A
Public Sub TotalCarryLuggage()
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageCarrySheet
    
    On Error GoTo TotalCarryLuggageErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim nCount As Integer
    Dim szaTemp() As TLuggageItemInfo
    Dim i As Integer
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT COUNT(l.luggage_id) AS luggage_number,SUM(a.cal_weight) AS cal_weight,SUM(a.fact_weight) AS fact_weight,SUM(a.baggage_number) AS baggage_number, " _
        & " SUM(a.over_weight_number) AS over_weight_number " _
        & " FROM luggage_carry_accept_lst l , luggage_accept_info a" _
        & " WHERE l.luggage_id = a.luggage_id  AND check_luggage_id =" & TransFieldValueToString(SheetID) _
        & " GROUP BY l.check_luggage_id "
    
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
'        mszSheetID = FormatDbValue(rsTemp!luggage_id)
        mdbCalWeight = FormatDbValue(rsTemp!cal_weight)             '计重
        mdbActWeight = FormatDbValue(rsTemp!fact_weight)            '实重
        mnNumber = FormatDbValue(rsTemp!baggage_number)             '件数
        mnOverNumber = FormatDbValue(rsTemp!over_weight_number)     '超重件数
        mnAcceptSheetNumber = FormatDbValue(rsTemp!luggage_number)  '受理单数
        
        
        
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing

    Exit Sub
TotalCarryLuggageErr:
    Call RaiseError(err.Number, "CarrySheet:TotalCarryLuggage Method", err.Description)
End Sub

Public Sub ClearLuggage()
    Dim i As Integer
    Dim szTemp As String
    Dim aszAcceptSheets() As String
    Dim atTemp() As TLuggagePriceItem
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr
    
    
    mAcceptSheetIDs = aszAcceptSheets
'    matPriceItems = atTemp
    szSql = "SELECT * FROM luggage_price_item_info WHERE use_mark=0 "
    Set rsTemp = oDb.Execute(szSql)
    ReDim matPriceItems(1 To rsTemp.RecordCount) As TLuggagePriceItem
    mnAcceptSheetNumber = 0
    
    
End Sub



'新增行包
'##ModelId=3E03F3BA0038
Public Sub AddLuggage(LuggageID As String, AcceptType As Integer)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageCarrySheet
    
    On Error GoTo AddLuggageErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szSql1 As String
    Dim rsTemp As Recordset
    Dim rstemp1 As Recordset
    Dim Count As Integer
    Dim Count1 As Integer
    Dim i As Integer, j As Integer
    Dim UseStatus(1 To 10) As Integer
    Dim szSql2 As String
    
'    AssertActiveUserValid m_oActiveUser, ERR_LugReceive
'    AssertHaveRight m_oActiveUser, RIGHT_CarrySheet
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    '判断受理单状态 正常状态 status=0
    szSql = "SELECT * FROM luggage_accept_info WHERE status=0 AND luggage_id=" & TransFieldValueToString(LuggageID)
    If oDb.Execute(szSql).RecordCount = 0 Then ShowError ERR_AcceptSheetNoFit
    
    Count = ArrayLength(mAcceptSheetIDs) + 1
    ReDim Preserve mAcceptSheetIDs(1 To Count) As String
    
    mAcceptSheetIDs(Count) = LuggageID
    mnAcceptSheetNumber = mnAcceptSheetNumber + 1
    
    '读出行包票价明细表中信息
    szSql = "SELECT * FROM luggage_price_lst WHERE luggage_id=" & TransFieldValueToString(LuggageID)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_CarrySheetLuggageExist
    
    szSql1 = "SELECT  * FROM luggage_price_item_info WHERE accept_type= " & TransFieldValueToString(AcceptType)
    Set rstemp1 = oDb.Execute(szSql1)
    
    
    j = 1
    For i = 1 To rstemp1.RecordCount
        
        If (FormatDbValue(rstemp1!use_mark) = LuggagePriceItemUsed) And rsTemp.Fields.Item(i + 1).Name = Trim("price_item_" & Trim(Str(i))) Then
            matPriceItems(j).PriceID = Trim("000" & Trim(Str(i - 1)))
            matPriceItems(j).PriceName = rstemp1!chinese_name
            '分项总价
            matPriceItems(j).PriceValue = matPriceItems(j).PriceValue + FormatDbValue(rsTemp.Fields(Trim("price_item_" & Trim(Str(i)))))
            j = j + 1
        End If
        rstemp1.MoveNext
    Next i
    
    
    Set oDb = Nothing
    Exit Sub
AddLuggageErr:
    Call RaiseError(err.Number, "CarrySheet:AddLuggage Method", err.Description)
End Sub

'删除行包
'##ModelId=3E0401DA01B4
Public Sub RemoveLuggage(LuggageID As String)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageCarrySheet
    
    Dim oDb As New RTConnection
    Dim szSql  As String, szSql1 As String
    Dim Count  As Integer
    Dim rsTemp As Recordset, rstemp1 As Recordset
    Dim atTempSheetID() As String, i As Integer, j As Integer
    On Error GoTo RemoveLuggageErr
    atTempSheetID = mAcceptSheetIDs
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    szSql = "SELECT * FROM luggage_accept_info WHERE luggage_id =" & TransFieldValueToString(LuggageID)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_CarrySheetLuggageExist
    
    
    
    Count = ArrayLength(mAcceptSheetIDs)
    
    '删除mAcceptSheetIDs中的受理单
    If Count <> 1 Then
        ReDim Preserve mAcceptSheetIDs(1 To Count) As String
        
        For i = 1 To Count - 1
            If mAcceptSheetIDs(i) = LuggageID Or atTempSheetID(Count) = LuggageID Then
                ReDim Preserve mAcceptSheetIDs(1 To Count - 1) As String
                mAcceptSheetIDs(i) = atTempSheetID(Count)
                mnAcceptSheetNumber = mnAcceptSheetNumber - 1
                Exit For
            End If
        Next
    Else
        mnAcceptSheetNumber = 0
    End If
    
    
    '读出行包票价明细表中信息
    If Count <> 1 Then
        szSql = "SELECT * FROM luggage_price_lst WHERE luggage_id=" & TransFieldValueToString(LuggageID)
        
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount = 0 Then ShowError ERR_CarrySheetLuggageExist
    
        szSql1 = "SELECT * FROM luggage_price_item_info  "
        Set rstemp1 = oDb.Execute(szSql1)
    
        j = 1
        For i = 1 To rstemp1.RecordCount
            If (rstemp1!use_mark = 0) And rsTemp.Fields.Item(Val(rstemp1!charge_item)).Name = Trim("price_item_" & Trim(Str(i))) Then
                
                matPriceItems(j).PriceID = Trim("000" & Trim(Str(i - 1)))
                matPriceItems(j).PriceName = rstemp1!chinese_name
                matPriceItems(j).PriceValue = matPriceItems(i).PriceValue - FormatDbValue(rsTemp("price_item_" & Trim(Str(i))))
                j = j + 1
            End If
            rstemp1.MoveNext
        Next i
    Else
        szSql1 = "SELECT * FROM luggage_price_item_info WHERE use_mark=0 "
        Set rsTemp = oDb.Execute(szSql1)
        ReDim matPriceItems(1 To rsTemp.RecordCount) As TLuggagePriceItem
        For i = 1 To rsTemp.RecordCount
            matPriceItems(i).PriceValue = 0
        Next
        
    End If
    
    Set oDb = Nothing

    Exit Sub
RemoveLuggageErr:
    Call RaiseError(err.Number, "CarrySheet:RemoveLuggage Method", err.Description)
End Sub

'根据车号显示发车时间
Public Function RefreshBusStartTime(mVehicle As String) As ADODB.Recordset
     AssertActiveUserValid m_oActiveUser, ERR_LuggageCarrySheet
    On Error GoTo RefreshBusVehicleErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim nCount As Integer
    Dim i As Integer
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
     Dim szTables As String

    If mdtBusDate = Date Then       '工作表
        szTables = cszTableEnvBusWork & " e, " & cszTableVehicle & " v , " & cszTableRouteInfo & " r"
    Else
        szTables = cszTableEnvBusInfo & " e, " & cszTableVehicle & " v , " & cszTableRouteInfo & " r"
    End If

    
    'RTStation数据库
    szSql = "SELECT DISTINCT e.bus_start_time " _
        & " FROM " & szTables _
        & " WHERE e.bus_date =" & TransFieldValueToString(mdtBusDate) _
        & " AND v.license_tag_no = " & TransFieldValueToString(mVehicle) _
        & " AND v.vehicle_id = e.vehicle_id AND  r.route_id = e.route_id "
    
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_CarrySheetLuggageNoBus '无此车次信息
    Set RefreshBusStartTime = rsTemp
    Set rsTemp = Nothing
    Set oDb = Nothing

    Exit Function
RefreshBusVehicleErr:
    Call RaiseError(err.Number, "CarrySheet:RefreshBusVehicle Method", err.Description)
End Function


'根据发车时间和车次显示车次信息
Public Sub RefreshBusInfoEX(mBusID As String, mBusStartTime As Date, AcceptType As String)
     AssertActiveUserValid m_oActiveUser, ERR_LuggageCarrySheet
    On Error GoTo RefreshBusVehicleErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim nCount As Integer
    Dim i As Integer
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
     Dim szTables As String
    If mdtBusDate = Date Then       '工作表
        szTables = cszTableEnvBusWork & " e, " & cszTableVehicle & " v ," & cszTableCompany & " c ," & cszTableCompany & " c1, " & cszTableRouteInfo & " r," & cszTableOwnerInfo & " o, " & cszTableEnvAllotWork & " a "
    Else
        szTables = cszTableEnvBusInfo & " e, " & cszTableVehicle & " v ," & cszTableCompany & " c ," & cszTableCompany & " c1, " & cszTableRouteInfo & " r," & cszTableOwnerInfo & " o, " & cszTableEnvAllotInfo & " a "
    End If

    
    'RTStation数据库
    szSql = "SELECT e.bus_id,e.bus_date,e.vehicle_id,v.license_tag_no,e.transport_company_id,c.transport_company_short_name," _
        & " e.transport_company_id as split_company_id,c1.transport_company_short_name as split_company_name, e.route_id,r.route_name,e.vehicle_type_code,e.vehicle_type_name,e.owner_id ," _
        & " o.owner_name,a.bus_start_time " _
        & " FROM " & szTables _
        & " WHERE a.bus_start_time =" & TransFieldValueToString(mBusStartTime) _
        & " AND e.bus_id = " & TransFieldValueToString(mBusID) _
        & " AND v.vehicle_id = e.vehicle_id AND c.transport_company_id = e.transport_company_id " _
        & " AND a.bus_id = e.bus_id AND a.bus_date = e.bus_date AND a.sell_station_id = " & TransFieldValueToString(m_oActiveUser.SellStationID) _
        & " AND e.transport_company_id = c1.transport_company_id AND r.route_id = e.route_id AND o.owner_id = e.owner_id "
    
    Set rsTemp = oDb.Execute(szSql)
    



    If rsTemp.RecordCount = 0 Then ShowError ERR_CarrySheetLuggageNoBus '无此车次信息
    
    '以上SZSQL无拆帐信息
    
    mszBusID = FormatDbValue(rsTemp!bus_id)                                     '车次代码
'    mszSellStationID = FormatDbValue(rsTemp!sell_station_id)                    '起点站代码
    mdtBusDate = FormatDbValue(rsTemp!bus_date)                                 '车次日期
'    If mszVehicleID = "" Then
        mszVehicleID = FormatDbValue(rsTemp!vehicle_id)                             '车辆代码
        mszVehicleLicense = FormatDbValue(rsTemp!license_tag_no)                    '车牌
'    End If
    mszCompanyID = FormatDbValue(rsTemp!transport_company_id)                   '参运公司代码
    mszCompanyName = FormatDbValue(rsTemp!transport_company_short_name)         '参运公司简称
    mszBusOwnerID = FormatDbValue(rsTemp!owner_id)                              '车主代码
    mszBusOwnerName = FormatDbValue(rsTemp!owner_name)                          '车主名称
    mdtBusStartOffTime = FormatDbValue(rsTemp!bus_start_time)                   '发车时间
    mszVehicleTypeID = FormatDbValue(rsTemp!vehicle_type_code)                  '车型代码
    mszVehicleTypeName = FormatDbValue(rsTemp!vehicle_type_name)                '车型名称
    mszRouteID = FormatDbValue(rsTemp!route_id)                                 '线路代码
    mszRouteName = FormatDbValue(rsTemp!route_name)                             '线路名称
    mszSplitCompanyID = FormatDbValue(rsTemp!split_company_id)                  '拆帐公司代码
    mszSplitCompanyName = FormatDbValue(rsTemp!split_company_name)    '拆账公司简称
    szSql = "SELECT s.protocol_name,s.protocol_id " _
        & " FROM luggage_settle_protocol_info s ,vehicle_lug_protocol_info v " _
        & " WHERE v.vehicle_id =" & TransFieldValueToString(mszVehicleID) _
        & " AND v.protocol_id = s.protocol_id AND v.accept_type = " & TransFieldValueToString(GetLuggageTypeInt(AcceptType))

    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        mszProtocolID = FormatDbValue(rsTemp!protocol_id)                           '拆算协议代码
        mszProtocolName = FormatDbValue(rsTemp!protocol_name)                       '拆算协议名称
    End If
        
    Set rsTemp = Nothing
    Set oDb = Nothing

    Exit Sub
RefreshBusVehicleErr:
    Call RaiseError(err.Number, "CarrySheet:RefreshBusVehicle Method", err.Description)
End Sub

'根据发车时间和车号显示车次信息
Public Sub RefreshBusInfo(mVehicle As String, mBusStartTime As Date, AcceptType As String)
     AssertActiveUserValid m_oActiveUser, ERR_LuggageCarrySheet
    On Error GoTo RefreshBusVehicleErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim nCount As Integer
    Dim i As Integer
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
     Dim szTables As String
    If mdtBusDate = Date Then       '工作表
        szTables = "" & cszTableEnvBusWork & " e, " & cszTableVehicle & " v ," & cszTableCompany & " c ," & cszTableCompany & " c1, " & cszTableRouteInfo & " r," & cszTableOwnerInfo & " o"
    Else
        szTables = "" & cszTableEnvBusInfo & " e, " & cszTableVehicle & " v ," & cszTableCompany & " c ," & cszTableCompany & " c1, " & cszTableRouteInfo & " r," & cszTableOwnerInfo & " o"
    End If

    
    szSql = "SELECT e.bus_id,e.bus_date,e.vehicle_id,v.license_tag_no,e.transport_company_id,c.transport_company_short_name," _
        & " e.transport_company_id as split_company_id,c1.transport_company_short_name as split_company_name, e.route_id,r.route_name,e.vehicle_type_code,e.vehicle_type_name,e.owner_id ," _
        & " o.owner_name,e.bus_start_time " _
        & " FROM " & szTables _
        & " WHERE e.bus_start_time =" & TransFieldValueToString(mBusStartTime) _
        & " AND v.license_tag_no = " & TransFieldValueToString(mVehicle) _
        & " AND v.vehicle_id = e.vehicle_id AND c.transport_company_id = e.transport_company_id " _
        & " AND e.transport_company_id = c1.transport_company_id AND r.route_id = e.route_id AND o.owner_id = e.owner_id "
    
    Set rsTemp = oDb.Execute(szSql)
    



    If rsTemp.RecordCount = 0 Then ShowError ERR_CarrySheetLuggageNoBus '无此车次信息
    
    '以上SZSQL无拆帐信息
    
    mszBusID = FormatDbValue(rsTemp!bus_id)                                     '车次代码
'    mszSellStationID = FormatDbValue(rsTemp!sell_station_id)                    '起点站代码
    mdtBusDate = FormatDbValue(rsTemp!bus_date)                                 '车次日期
'    If mszVehicleID = "" Then
        mszVehicleID = FormatDbValue(rsTemp!vehicle_id)                             '车辆代码
        mszVehicleLicense = FormatDbValue(rsTemp!license_tag_no)                    '车牌
'    End If
    mszCompanyID = FormatDbValue(rsTemp!transport_company_id)                   '参运公司代码
    mszCompanyName = FormatDbValue(rsTemp!transport_company_short_name)         '参运公司简称
    mszBusOwnerID = FormatDbValue(rsTemp!owner_id)                              '车主代码
    mszBusOwnerName = FormatDbValue(rsTemp!owner_name)                          '车主名称
    mdtBusStartOffTime = FormatDbValue(rsTemp!bus_start_time)                   '发车时间
    mszVehicleTypeID = FormatDbValue(rsTemp!vehicle_type_code)                  '车型代码
    mszVehicleTypeName = FormatDbValue(rsTemp!vehicle_type_name)                '车型名称
    mszRouteID = FormatDbValue(rsTemp!route_id)                                 '线路代码
    mszRouteName = FormatDbValue(rsTemp!route_name)                             '线路名称
    mszSplitCompanyID = FormatDbValue(rsTemp!split_company_id)                  '拆帐公司代码
    mszSplitCompanyName = FormatDbValue(rsTemp!split_company_name)    '拆账公司简称
    '嘉兴数据庫
  szSql = "SELECT s.protocol_name,s.protocol_id " _
        & " FROM luggage_settle_protocol_info s ,vehicle_lug_protocol_info v " _
        & " WHERE v.vehicle_id =" & TransFieldValueToString(mszVehicleID) _
        & " AND v.protocol_id = s.protocol_id AND v.accept_type = " & TransFieldValueToString(GetLuggageTypeInt(AcceptType))

    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        mszProtocolID = FormatDbValue(rsTemp!protocol_id)                           '拆算协议代码
        mszProtocolName = FormatDbValue(rsTemp!protocol_name)                       '拆算协议名称
    End If
        
    Set rsTemp = Nothing
    Set oDb = Nothing

    Exit Sub
RefreshBusVehicleErr:
    Call RaiseError(err.Number, "CarrySheet:RefreshBusVehicle Method", err.Description)
End Sub


'刷新承运车次、车辆及财务拆算信息
'
'##ModelId=3E04029A028C
Public Sub RefreshBusVehicle(AcceptType As String)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageCarrySheet
    
    On Error GoTo RefreshBusVehicleErr
 ' 没有完成财务拆算信息写入
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim nCount As Integer
    Dim i As Integer
    
'    AssertActiveUserValid m_oActiveUser, ERR_LugReceive
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    Dim szTables As String
    If mdtBusDate = Date Then       '工作表
        szTables = "" & cszTableEnvBusWork & " e, " & cszTableVehicle & " v ," & cszTableCompany & " c ," & cszTableCompany & " c1, " & cszTableRouteInfo & " r," & cszTableOwnerInfo & " o"
    Else
        szTables = "" & cszTableEnvBusInfo & " e, " & cszTableVehicle & " v ," & cszTableCompany & " c ," & cszTableCompany & " c1, " & cszTableRouteInfo & " r," & cszTableOwnerInfo & " o"
    End If
    
    szSql = "SELECT e.bus_id,e.bus_date,e.vehicle_id,v.license_tag_no,e.transport_company_id,c.transport_company_short_name," _
        & " e.split_company_id,c1.transport_company_short_name, e.route_id,r.route_name,e.vehicle_type_code,e.vehicle_type_name,e.owner_id ," _
        & " o.owner_name,e.bus_start_time " _
        & " FROM " & szTables _
        & " WHERE e.bus_id =" & TransFieldValueToString(mszBusID) _
        & " AND e.bus_date = " & TransFieldValueToString(ToDBDate(mdtBusDate)) _
        & " AND v.vehicle_id = e.vehicle_id AND c.transport_company_id = e.transport_company_id " _
        & " AND e.transport_company_id = c1.transport_company_id AND r.route_id = e.route_id AND o.owner_id = e.owner_id "
    
    Set rsTemp = oDb.Execute(szSql)
    



    If rsTemp.RecordCount = 0 Then ShowError ERR_CarrySheetLuggageNoBus '无此车次信息
    
    '以上SZSQL无拆帐信息
    
    mszBusID = FormatDbValue(rsTemp!bus_id)                                     '车次代码
'    mszSellStationID = FormatDbValue(rsTemp!sell_station_id)                    '起点站代码
    mdtBusDate = FormatDbValue(rsTemp!bus_date)                                 '车次日期
    If mszVehicleID = "" Then
        mszVehicleID = FormatDbValue(rsTemp!vehicle_id)                             '车辆代码
        mszVehicleLicense = FormatDbValue(rsTemp!license_tag_no)                    '车牌
    End If
    mszCompanyID = FormatDbValue(rsTemp!transport_company_id)                   '参运公司代码
    mszCompanyName = FormatDbValue(rsTemp!transport_company_short_name)         '参运公司简称
    mszBusOwnerID = FormatDbValue(rsTemp!owner_id)                              '车主代码
    mszBusOwnerName = FormatDbValue(rsTemp!owner_name)                          '车主名称
    mdtBusStartOffTime = FormatDbValue(rsTemp!bus_start_time)                   '发车时间
    mszVehicleTypeID = FormatDbValue(rsTemp!vehicle_type_code)                  '车型代码
    mszVehicleTypeName = FormatDbValue(rsTemp!vehicle_type_name)                '车型名称
    mszRouteID = FormatDbValue(rsTemp!route_id)                                 '线路代码
    mszRouteName = FormatDbValue(rsTemp!route_name)                             '线路名称
    mszSplitCompanyID = FormatDbValue(rsTemp!split_company_id)                  '拆帐公司代码
    mszSplitCompanyName = FormatDbValue(rsTemp!transport_company_short_name)    '拆账公司简称
    '======================================
    'RTStation
'    '得到拆算信息
'    szSql = "SELECT s.protocol_name,l.protocol_id " _
'        & " FROM settle_protocol_info s ,vehicle_lug_protocol_info v " _
'        & " WHERE v.vehicle_id =" & TransFieldValueToString(mszVehicleID) _
'        & " AND v.protocol_id = s.protocol_id "
''
    '=======================================
    '嘉兴数据庫
    szSql = "SELECT s.protocol_name,s.protocol_id " _
        & " FROM luggage_settle_protocol_info s ,vehicle_lug_protocol_info v " _
        & " WHERE v.vehicle_id =" & TransFieldValueToString(mszVehicleID) _
        & " AND v.protocol_id = s.protocol_id AND v.accept_type = " & TransFieldValueToString(GetLuggageTypeInt(AcceptType))
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        mszProtocolID = FormatDbValue(rsTemp!protocol_id)                           '拆算协议代码
        mszProtocolName = FormatDbValue(rsTemp!protocol_name)                       '拆算协议名称
    End If
        
    Set rsTemp = Nothing
    Set oDb = Nothing

    Exit Sub
RefreshBusVehicleErr:
    Call RaiseError(err.Number, "CarrySheet:RefreshBusVehicle Method", err.Description)
End Sub

'##ModelId=3E04A79F00F7
Public Property Get AcceptSheetIDs() As String()
    Let AcceptSheetIDs = mAcceptSheetIDs
End Property

'##ModelId=3E04A79E031C
Public Property Get PriceItems() As TLuggagePriceItem()
    Let PriceItems = matPriceItems
End Property

'##ModelId=3E04A79D0266
Public Property Get TotalPrice() As Double
    Dim Count As Integer, i As Integer
    Count = ArrayLength(matPriceItems)
    mdbTotalPrice = 0

    For i = 1 To Count
        mdbTotalPrice = mdbTotalPrice + matPriceItems(i).PriceValue
    Next
    Let TotalPrice = mdbTotalPrice
End Property

'##ModelId=3E04A79D00B8
Public Property Get StatusString() As String
    Select Case mnStatus
        Case 0
            mszStatusString = "已作废"
        Case 1
            mszStatusString = "正常"
    End Select
        
    Let StatusString = mszStatusString
End Property

'##ModelId=3E04A79C02F1
Public Property Get Status() As Integer
'    Dim szSql As String
'    Dim oDb As New RTConnection
'    Dim rsTemp As Recordset
'    oDb.ConnectionString = GetConnectionStr(cszLuggage)
'    szSql = "SELECT valid_mark FROM luggage_sheet_info WHERE sheet_id =" & TransFieldValueToString(mszSheetID)
'    Set rsTemp = oDb.Execute(szSql)
'    mnStatus = FormatDbValue(rsTemp!valid_mark)
    Let Status = mnStatus
End Property

'##ModelId=3E04A79C0143
Public Property Get OverNumber() As Integer
   Let OverNumber = mnOverNumber
End Property

'##ModelId=3E04A79B025A
Public Property Let OverNumber(ByVal Value As Integer)
    Let mnOverNumber = Value
End Property

'##ModelId=3E04A79B00B5
Public Property Get Number() As Integer
   Let Number = mnNumber
End Property

'##ModelId=3E04A79A01EA
Public Property Let Number(ByVal Value As Integer)
    Let mnNumber = Value
End Property

'##ModelId=3E04A79A004F
Public Property Get ActWeight() As Double
   Let ActWeight = mdbActWeight
End Property

'##ModelId=3E04A799018E
Public Property Let ActWeight(ByVal Value As Double)
    Let mdbActWeight = Value
End Property

'##ModelId=3E04A79803DC
Public Property Get CalWeight() As Double
   Let CalWeight = mdbCalWeight
End Property

'##ModelId=3E04A798013D
Public Property Let CalWeight(ByVal Value As Double)
    Let mdbCalWeight = Value
End Property

'##ModelId=3E04A7970394
Public Property Get AcceptSheetNumber() As Integer
   Let AcceptSheetNumber = mnAcceptSheetNumber
End Property

'##ModelId=3E04A797020E
Public Property Get OperateTime() As Date
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim TempNow As Date
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT getdate() as NowTime"
    Set rsTemp = oDb.Execute(szSql)
    
    TempNow = FormatDbValue(rsTemp!NowTime)
    
'    If mdtOperateTime < TempNow Then
'        mdtOperateTime = TempNow
'    End If
        
    Let OperateTime = mdtOperateTime
End Property

'##ModelId=3E04A7970091
Public Property Get OperatorName() As String
'    If mszProtocolName = "" Then
'        mszOperatorName = m_oActiveUser.UserName
'    End If
    
    Let OperatorName = mszOperatorName
End Property
'##ModelId=3E04A79602F3
Public Property Get MoveWorker() As String
'    If mszOperatorID = "" Then
'        mszOperatorID = m_oActiveUser.UserID
'    End If
    Let MoveWorker = mszMoveWorker
End Property
'##ModelId=3E04A795016B
Public Property Let MoveWorker(ByVal Value As String)
    Let mszMoveWorker = Value
End Property

'##ModelId=3E04A79602F3
Public Property Get OperatorID() As String
'    If mszOperatorID = "" Then
'        mszOperatorID = m_oActiveUser.UserID
'    End If
    Let OperatorID = mszOperatorID
End Property

'##ModelId=3E04A7960180
Public Property Get ProtocolName() As String
    
    Let ProtocolName = mszProtocolName
End Property

'##ModelId=3E04A79503D7
Public Property Get ProtocolID() As String
   Let ProtocolID = mszProtocolID
End Property

'##ModelId=3E04A795016B
Public Property Let ProtocolID(ByVal Value As String)
    Let mszProtocolID = Value
End Property

'##ModelId=3E04A79403E0
Public Property Get Valid() As Boolean
   Let Valid = mbValid
End Property

'##ModelId=3E04A7940278
Public Property Get BusOwnerName() As String
   Let BusOwnerName = mszBusOwnerName
End Property

Public Property Let BusOwnerName(ByVal Value As String)
    Let mszBusOwnerName = Value
End Property


'##ModelId=3E04A7940105
Public Property Get BusOwnerID() As String
   Let BusOwnerID = mszBusOwnerID
End Property

'##ModelId=3E04A7930294
Public Property Let BusOwnerID(ByVal Value As String)
    Let mszBusOwnerID = Value
End Property

'##ModelId=3E04A7930136
Public Property Get VehicleTypeName() As String
   Let VehicleTypeName = mszVehicleTypeName
End Property

'##ModelId=3E04A79203BF
Public Property Get VehicleTypeID() As String
   Let VehicleTypeID = mszVehicleTypeID
End Property

'##ModelId=3E04A7920170
Public Property Let VehicleTypeID(ByVal Value As String)
    Let mszVehicleTypeID = Value
End Property

'##ModelId=3E04A792001C
Public Property Get RouteName() As String
   Let RouteName = mszRouteName
End Property

'##ModelId=3E04A79102AF
Public Property Get RouteID() As String
   Let RouteID = mszRouteID
End Property

'##ModelId=3E04A7910074
Public Property Let RouteID(ByVal Value As String)
    Let mszRouteID = Value
End Property

'##ModelId=3E04A7900312
Public Property Get SplitCompanyName() As String
   Let SplitCompanyName = mszSplitCompanyName
End Property

Public Property Let SplitCompanyName(ByVal Value As String)
    Let mszSplitCompanyName = Value
End Property
'##ModelId=3E04A79001C8
Public Property Get SplitCompanyID() As String
   Let SplitCompanyID = mszSplitCompanyID
End Property

'##ModelId=3E04A78F0389
Public Property Let SplitCompanyID(ByVal Value As String)
    Let mszSplitCompanyID = Value
End Property

'##ModelId=3E04A78F0248
Public Property Get CompanyName() As String
   Let CompanyName = mszCompanyName
End Property

Public Property Let CompanyName(ByVal Value As String)
    Let mszCompanyName = Value
End Property

'##ModelId=3E04A78F0108
Public Property Get CompanyID() As String
   Let CompanyID = mszCompanyID
End Property

'##ModelId=3E04A78E02DD
Public Property Let CompanyID(ByVal Value As String)
    Let mszCompanyID = Value
End Property

'##ModelId=3E04A78E019D
Public Property Get VehicleLicense() As String
   Let VehicleLicense = mszVehicleLicense
End Property
Public Property Let VehicleLicense(ByVal Value As String)
    Let mszVehicleLicense = Value
End Property
'##ModelId=3E04A78E0070
Public Property Get VehicleID() As String
   Let VehicleID = mszVehicleID
End Property

'##ModelId=3E04A78D024F
Public Property Let VehicleID(ByVal Value As String)
    Let mszVehicleID = Value
End Property

'##ModelId=3E04A78D0123
Public Property Get BusStartOffTime() As Date
   Let BusStartOffTime = mdtBusStartOffTime
End Property

'##ModelId=3E04A78C0316
Public Property Let BusStartOffTime(ByVal Value As Date)
    Let mdtBusStartOffTime = Value
End Property

'##ModelId=3E04A78C01EA
Public Property Get BusDate() As Date
   Let BusDate = mdtBusDate
End Property

'##ModelId=3E04A78B03E7
Public Property Let BusDate(ByVal Value As Date)
    Let mdtBusDate = Value
End Property

'##ModelId=3E04A78B02C5
Public Property Get BusID() As String
   Let BusID = mszBusID
End Property

'##ModelId=3E04A78B00DA
Public Property Let BusID(ByVal Value As String)
    Let mszBusID = Value
End Property

'##ModelId=3E04A78A03B4
Public Property Get SheetID() As String
   Let SheetID = mszSheetID
End Property


'##ModelId=3E04A78A01DD
Public Property Let SheetID(ByVal Value As String)
    Let mszSheetID = Value
End Property
'
'Public Property Let SellStationID(ByVal Value As String)
'    Let mszSellStationID = Value
'End Property

Public Property Get SellStationID() As String
    Let SellStationID = mszSellStationID
End Property
Public Property Get AcceptType() As String
   Let AcceptType = mszAcceptType
End Property

'##ModelId=3E04A7A7006C
Public Property Let AcceptType(ByVal Value As String)
    Let mszAcceptType = Value
End Property


Public Property Get EndStation() As String
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr
    szSql = " SELECT s.station_name FROM " & cszTableRouteInfo & " r , " & cszTableStationInfo & " s " _
        & " WHERE r.route_id = " & TransFieldValueToString(mszRouteID) _
        & " AND r.end_station_id = s.station_id "
    Set rsTemp = oDb.Execute(szSql)
    
    If rsTemp.RecordCount > 0 Then
        EndStation = FormatDbValue(rsTemp!station_name)
    End If
    
End Property

'打印的应结费率


Public Property Get PrintSettleRatio() As Double
   Let PrintSettleRatio = m_dbPrintSettleRatio
End Property

Public Property Let PrintSettleRatio(ByVal Value As Double)
    Let m_dbPrintSettleRatio = Value
End Property

'打印的应结运费
Public Property Get PrintSettlePrice() As Double
   Let PrintSettlePrice = m_dbPrintSettlePrice
End Property

Public Property Let PrintSettlePrice(ByVal Value As Double)
    Let m_dbPrintSettlePrice = Value
End Property


