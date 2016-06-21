VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FinanceSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'行包财务结算单
Option Base 0

Option Explicit


'出错信息
Public Enum EErrLuggageFinanceSheet
                                                                '起始号33600
    ERR_PriceItemNotExistsdf = ERR_LuggageFinanceSheet + 20   '33620票价项不存在
    ERR_FinSheetNotExist = ERR_LuggageFinanceSheet + 21 '33621结算单不存在
    ERR_FinSheetIDNullOrNotAdd = ERR_LuggageFinanceSheet + 22 '33622结算单ID为空或非新增状态
    ERR_FinSheetExist = ERR_LuggageFinanceSheet + 23 '33623结算单已存在
    ERR_FinSheetNoEditStatus = ERR_LuggageFinanceSheet + 24 '33624不是编辑状态
    ERR_FinSheetIsNull = ERR_LuggageFinanceSheet + 25 '33625签发单不存在
    ERR_FinSheetIDExist = ERR_LuggageFinanceSheet + 26 '33626签发单已存在
    ERR_FinSheetProtocolNoExist = ERR_LuggageFinanceSheet + 27 '33627车辆协议不存在
    ERR_FinSheetPriceItemNoExist = ERR_LuggageFinanceSheet + 28 '33628票价项不存在
    ERR_SheetIDExist = ERR_LuggageFinanceSheet + 29 '33629 签发单已经存在
    '出错信息
End Enum
'权限
Public Enum RIGHT_LuggageFinanceSheet
    RIGHT_FinancesheetAdd = ERR_LuggageFinanceSheet + cnMidRightBegin + cnMidRightStep * 1  ' 33706  结算新增权限
    RIGHT_FinancesheetDel = ERR_LuggageFinanceSheet + cnMidRightBegin + cnMidRightStep * 2 ' 33712结算单删除
    RIGHT_FinancesheetEdit = ERR_LuggageFinanceSheet + cnMidRightBegin + cnMidRightStep * 3   '33718结算音编辑
    RIGHT_FinanceSheetCancel = ERR_LuggageFinanceSheet + cnMidRightBegin + cnMidRightStep * 4 ' 33724 废结算单权限
    RIGHT_FinanceSheetSetPayed = ERR_LuggageFinanceSheet + cnMidRightBegin + cnMidRightStep * 5  '33730 结结算单权限
End Enum

Public Enum ELuggageSettleStatus
    '签发单是否结算状态
    LuggageNotSettle = 0
    LuggageHasSettle = 1
End Enum


Public Enum ELuggageSettleValidMark
    '结算单是否有效
    LuggageValid = 1
    LuggageNotValid = 0
End Enum
    
Public Enum ELuggageProtocolDefault
    '是否默认协议
    LuggageProtocolDefaultGeneral = 0 '默认快件
    LuggageProtocolDefaultMan = 1 '默认随行协议
    LuggageProtocolNotDefault = 2 '未默认协议
    
End Enum




Const MyUnhandledError = 234
'结算单号
'##ModelId=3E1B4B3C0311
Private mszSheetID As String

'售票站代码
'##ModelId=3E29E28402BC
Private mszSellStationID As String

'售票站名称（只读）
'##ModelId=3E29E2AF011A
Private mszSellStationName As String

'结算月份
'##ModelId=3E1F41640212j
Private mdtSettleMonth As Date

'开始结算时间
'##ModelId=3E1B4C5303E4
Private mdtStartSettleDate As Date

'结束结算时间
'##ModelId=3E1B4C8E0258
Private mdtStopSettleDate As Date

'总结算款，（只读）
'##ModelId=3E1B4B3C03BB
Private mdbTotalPrice As Double

'应拆款，（只读）
'##ModelId=3E1B4D7B03D6
Private mdbNeedSplitPrice As Double

'实拆款
'##ModelId=3E1B4D7C000C
Private mdbActSplitPrice As Double

'承运车辆代码（只读）
'##ModelId=3E1B4B3C0325
Private mszVehicleID As String

'承运车辆牌号（只读）
'##ModelId=3E1B4B3C032F
Private mszVehicleLicense As String

'参运公司代码（只读）
'##ModelId=3E1B4B3C0339
Private mszCompanyID As String

'参运公司名称（只读）
'##ModelId=3E1B4B3C034D
Private mszCompanyName As String

'拆帐公司代码（只读）
'##ModelId=3E1B4B3C0357
Private mszSplitCompanyID As String

'拆帐公司名称（只读）
'##ModelId=3E1B4B3C036B
Private mszSplitCompanyName As String

'线路代码（只读）
'##ModelId=3E1B4B3C0375
Private mszRouteID As String

'线路名称（只读）
'##ModelId=3E1B4B3C0389
Private mszRouteName As String

'车主代码（只读）
'##ModelId=3E1B4B3C039D
Private mszBusOwnerID As String

'车主名称（只读）
'##ModelId=3E1B4B3C03B1
Private mszBusOwnerName As String

'结算时间
'##ModelId=3E1B4CCF01F8
Private mdtOperateTime As Date

'操作员姓名
'##ModelId=3E1B4CF20389
Private mszOperatorName As String

'结算单状态(1-已结,0-作废)
'##ModelId=3E1B4F8802CC
Private mnStatus As Integer

'（只读）结算单状态(未结,已结,作废)
'##ModelId=3E1B4F8802E0
Private mszStatusString As String

'备注
'##ModelId=3E1B68AF018B
Private mszRemark As String
Private m_oActiveUser As ActiveUser
Private m_nObjectStatus As EObjectStatus
Private mszAcceptType As String
Private mszProtocolID As String
Private mszProtocolName As String
Private mnSplitObjectType As Integer
Private mszSplitObjectID As String
Private mszSplitObjectName As String
Private maszSplitLuggageItem() As String
Private moCalc As New Calculater

'初始化活动对象
'##ModelId=3E0B593301D1
Public Sub Init(poAUser As ActiveUser)
    On Error GoTo InitErr
    
    Set m_oActiveUser = poAUser
    Exit Sub
InitErr:
    Call RaiseError(MyUnhandledError, "FinanceSheet:Init Method")
End Sub

'指定结算单
'##ModelId=3E1B705A009D
Public Sub Identify(SheetID As String)
    AssertActiveUserValid m_oActiveUser, ERR_LuggageFinanceSheet
    On Error GoTo IdentifyErr
    mszSheetID = SheetID
    RefreshInfo
    m_nObjectStatus = ST_NormalObj
    Exit Sub
IdentifyErr:
    Call RaiseError(MyUnhandledError, "FinanceSheet:Identify Method")
End Sub
Public Function GetVehicleInfo() As ADODB.Recordset
    AssertActiveUserValid m_oActiveUser, ERR_LuggageFinanceSheet
On Error GoTo here:
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim aszTempVehicleInfo() As String
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    
    szSql = "SELECT max(f.license_tag_no) as license_tag_no,max(f.owner_name) as owner_name,f.vehicle_id," _
           & "max(f.transport_company_short_name) as transport_company_short_name,max(f.split_company_name) as split_company_name, " _
           & "max(f.protocol_name) as protocol_name,max(s.charge_code) as charge_code,max(s.charge_name) as charge_name ," _
           & "SUM (s.split_out_money) as split_out_money " _
           & "FROM fin_luggage_sheet_lst f,fin_luggage_split_detail_lst s WHERE f.fin_sheet_id=s.fin_sheet_id AND f.vehicle_id=s.vehicle_id AND f.fin_sheet_id = " & TransFieldValueToString(mszSheetID) _
           & " GROUP BY f.vehicle_id,charge_code"
    Set rsTemp = oDb.Execute(szSql)
    Set GetVehicleInfo = rsTemp
Exit Function
here:
   Call RaiseError(MyUnhandledError, "FinanceSheet:Identify Method")
End Function


Private Sub RefreshInfo()
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT * FROM fin_luggage_sheet_info " _
        & " WHERE fin_sheet_id = " & TransFieldValueToString(mszSheetID)
         
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then
        RaiseError ERR_FinSheetNotExist '结算单不存在
    End If
    mszSheetID = FormatDbValue(rsTemp!fin_sheet_id)
    mszSellStationID = FormatDbValue(rsTemp!sell_station_id)
    mdtSettleMonth = FormatDbValue(rsTemp!settle_month)
    mszAcceptType = FormatDbValue(rsTemp!accept_type)
    mdtStartSettleDate = FormatDbValue(rsTemp!settlement_start_time)
    mdtStopSettleDate = FormatDbValue(rsTemp!settlement_end_time)
    mdbTotalPrice = FormatDbValue(rsTemp!total_price)
    mnSplitObjectType = FormatDbValue(rsTemp!split_object_type)
    mszSplitObjectID = FormatDbValue(rsTemp!split_object_id)
    mszSplitObjectName = FormatDbValue(rsTemp!split_object_name)
    mdbNeedSplitPrice = FormatDbValue(rsTemp!need_split_out)
    mdbActSplitPrice = FormatDbValue(rsTemp!act_split_out)
    mnStatus = FormatDbValue(rsTemp!Status)
    mszOperatorName = FormatDbValue(rsTemp!Operator)
    mdtOperateTime = FormatDbValue(rsTemp!operate_date)
    mszRemark = FormatDbValue(rsTemp!Annotation)
    
'    mszProtocolID = FormatDbValue(rsTemp!protocol_id)
'    mszProtocolName = FormatDbValue(rsTemp!protocol_name)
'    mszVehicleLicense = FormatDbValue(rsTemp!license_tag_no)
'    mszCompanyID = FormatDbValue(rsTemp!transport_company_id)
'    mszCompanyName = FormatDbValue(rsTemp!transport_company_short_name)
'    mszBusOwnerID = FormatDbValue(rsTemp!owner_id)
'    mszBusOwnerName = FormatDbValue(rsTemp!owner_name)
'    mszSplitCompanyID = FormatDbValue(rsTemp!split_company_id)
'    mszSplitCompanyName = FormatDbValue(rsTemp!split_company_name)
'    mszRouteID = FormatDbValue(rsTemp!route_id)
'    mszRouteName = FormatDbValue(rsTemp!route_name)
    
End Sub

'新增结算单
'##ModelId=3E1B703C014E
Public Sub AddNew()
    AssertActiveUserValid m_oActiveUser, ERR_LuggageFinanceSheet
    On Error GoTo AddNewErr
    m_nObjectStatus = ST_AddObj
    mszBusOwnerID = ""
    mszBusOwnerName = ""
    mszCompanyID = ""
    mszCompanyName = ""
    mszOperatorName = ""
    mszRemark = ""
    mszRouteID = ""
    mszRouteName = ""
    mszSellStationID = ""
    mszSellStationName = ""
    mszSheetID = ""
    mszSplitCompanyID = ""
    mszSplitCompanyName = ""
    mszStatusString = ""
    mszVehicleID = ""
    mszVehicleLicense = ""
    mdtOperateTime = Now
    mdtSettleMonth = Now
    mdtStartSettleDate = Now
    mdtStopSettleDate = Now
    mdbActSplitPrice = 0
    mdbNeedSplitPrice = 0
    mdbTotalPrice = 0
    mnStatus = ELuggageSettleValidMark.LuggageValid
    
    Exit Sub
AddNewErr:
    Call RaiseError(MyUnhandledError, "FinanceSheet:AddNew Method")
End Sub

'更新结算单
'##ModelId=3E1B703C0162
Public Sub Update()
    AssertActiveUserValid m_oActiveUser, ERR_LuggageFinanceSheet
    On Error GoTo UpdateErr
    If m_nObjectStatus = ST_AddObj Then
        AddObject
    ElseIf m_nObjectStatus = ST_EditObj Then
        AssertHaveRight m_oActiveUser, RIGHT_FinancesheetEdit
        UpdateToDB
    End If
    m_nObjectStatus = ST_NormalObj
    Exit Sub
UpdateErr:
    Call RaiseError(MyUnhandledError, "FinanceSheet:Update Method")
End Sub

Private Sub AddObject()
On Error GoTo FinErr
    Dim oDb As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    If mszSheetID = "" Or mszSheetID = Null Or m_nObjectStatus <> ST_AddObj Then
        RaiseError ERR_FinSheetIDNullOrNotAdd '结算单ID为空或非新增状态
    End If
    szSql = "INSERT fin_luggage_sheet_info ( fin_sheet_id,sell_station_id,settle_month,accept_type,settlement_start_time, " _
        & " settlement_end_time,total_price,split_object_type,split_object_id,split_object_name,need_split_out,act_split_out, " _
        & " status,operator, operate_date,annotation ) " _
        & " VALUES ( " _
        & TransFieldValueToString(mszSheetID) & "," & TransFieldValueToString(mszSellStationID) & "," _
        & TransFieldValueToString(mdtSettleMonth) & "," & TransFieldValueToString(GetLuggageTypeInt(mszAcceptType)) & "," _
        & TransFieldValueToString(mdtStartSettleDate) & "," & TransFieldValueToString(mdtStopSettleDate) & "," _
        & TransFieldValueToString(mdbTotalPrice) & "," & TransFieldValueToString(mnSplitObjectType) & "," _
        & TransFieldValueToString(mszSplitObjectID) & "," & TransFieldValueToString(mszSplitObjectName) & "," _
        & TransFieldValueToString(mdbNeedSplitPrice) & "," _
        & TransFieldValueToString(mdbActSplitPrice) & "," & TransFieldValueToString(ELuggageSettleValidMark.LuggageValid) & "," _
        & TransFieldValueToString(mszOperatorName) & "," _
        & TransFieldValueToString(mdtOperateTime) & "," & TransFieldValueToString(mszRemark) & ")"
    oDb.Execute szSql
    Exit Sub
FinErr:
    RaiseError ERR_FinSheetExist '结算单已存在
End Sub
Private Sub UpdateToDB()
On Error GoTo FinErr
    Dim oDb As New RTConnection
    Dim szSql As String
    
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "UPDATE fin_luggage_sheet_info " _
        & " SET sell_station_id = " & TransFieldValueToString(mszSellStationID) & ", settle_month = " & TransFieldValueToString(mdtSettleMonth) & "," _
        & " accept_type = " & TransFieldValueToString(mszAcceptType) & ", settlement_start_time =" & TransFieldValueToString(mdtStartSettleDate) & "," _
        & " settlement_end_time = " & TransFieldValueToString(mdtStopSettleDate) & ", total_price =" & TransFieldValueToString(mdbTotalPrice) & "," _
        & " split_object_type = " & TransFieldValueToString(mnSplitObjectType) & ", split_object_id = " & TransFieldValueToString(mszSplitObjectID) & "," _
        & " split_object_name = " & TransFieldValueToString(mszSplitObjectName) & "," _
        & " need_split_out = " & TransFieldValueToString(mdbNeedSplitPrice) & ", act_split_out = " & TransFieldValueToString(mdbActSplitPrice) & "," _
        & " status = " & TransFieldValueToString(mnStatus) & "," _
        & " operator = " & TransFieldValueToString(mszOperatorName) & "," _
        & " operate_date = " & TransFieldValueToString(mdtOperateTime) & ", annotation = " & TransFieldValueToString(mszRemark) _
        & " WHERE fin_sheet_id = " & TransFieldValueToString(mszSheetID)
    oDb.Execute szSql
    Exit Sub
FinErr:
    RaiseError ERR_FinSheetExist
End Sub
'删除结算单
'##ModelId=3E1B703C0176
Public Sub Delete()
    AssertActiveUserValid m_oActiveUser, ERR_LuggageFinanceSheet
    AssertHaveRight m_oActiveUser, RIGHT_FinancesheetDel
    On Error GoTo DeleteErr
    If m_nObjectStatus <> ST_AddObj Then
        DeleteObject
    End If
    Exit Sub
DeleteErr:
    Call RaiseError(MyUnhandledError, "FinanceSheet:Delete Method")
End Sub
Private Sub DeleteObject()
On Error GoTo FinErr
    Dim oDb As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = " DELETE fin_luggage_sheet_lst WHERE fin_sheet_id = " & TransFieldValueToString(mszSheetID)
    oDb.Execute szSql
    szSql = "DELETE fin_luggage_sheet_info WHERE fin_sheet_id = " & TransFieldValueToString(mszSheetID)
    oDb.Execute szSql
    Set oDb = Nothing
    Exit Sub
FinErr:
    RaiseError ERR_FinSheetExist '结算单已存在
End Sub
'作废
'
'##ModelId=3E1B7097018B
Public Sub Cancel()
    AssertActiveUserValid m_oActiveUser, ERR_LuggageFinanceSheet
    AssertHaveRight m_oActiveUser, RIGHT_FinanceSheetCancel
    On Error GoTo CancelErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT * FROM fin_luggage_sheet_info WHERE fin_sheet_id = " & TransFieldValueToString(mszSheetID)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then
        RaiseError ERR_FinSheetNotExist '结算单不存在
    End If
    oDb.BeginTrans
    '行包车辆结算明细表 状态
    szSql = "UPDATE fin_luggage_sheet_lst SET status = " & TransFieldValueToString(ELuggageSettleValidMark.LuggageNotValid) _
        & " WHERE fin_sheet_id = " & TransFieldValueToString(mszSheetID)
'        & " AND status <> 1 "
    oDb.Execute szSql
    szSql = "UPDATE fin_luggage_sheet_info SET status = " & TransFieldValueToString(ELuggageSettleValidMark.LuggageNotValid) _
        & " WHERE fin_sheet_id = " & TransFieldValueToString(mszSheetID)
 '       & " AND status <> 1 "
    oDb.Execute szSql
    'luggage_sheet_info's settlement_status 0 not结, 1 结
    szSql = "UPDATE luggage_sheet_info SET settlement_status=" & TransFieldValueToString(ELuggageSettleStatus.LuggageNotSettle) & " WHERE sheet_id in ( SELECT luggage_sheet_id FROM fin_sheet_to_luggage_lst " _
            & " WHERE fin_sheet_id =" & TransFieldValueToString(mszSheetID) & ")"
    oDb.Execute szSql
    oDb.CommitTrans
    Exit Sub
CancelErr:
    oDb.RollbackTrans
    Call RaiseError(MyUnhandledError, "FinanceSheet:Cancel Method")
End Sub

'设置已结
'##ModelId=3E28D0ED020F
Public Sub SetPayed()
    AssertActiveUserValid m_oActiveUser, ERR_LuggageFinanceSheet
    AssertHaveRight m_oActiveUser, RIGHT_FinanceSheetSetPayed
    On Error GoTo SetPayedErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT * FROM fin_luggage_sheet_info WHERE fin_sheet_id = " & TransFieldValueToString(mszSheetID)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then
        RaiseError ERR_FinSheetNotExist '结算单不存在
    End If
    szSql = "UPDATE fin_luggage_sheet_info SET status=1 WHERE fin_sheet_id = " & TransFieldValueToString(mszSheetID)
    oDb.Execute (szSql)
    szSql = "UPDATE fin_luggage_sheet_lst SET status=1 WHERE fin_sheet_id = " & TransFieldValueToString(mszSheetID)
    oDb.Execute szSql
    
    Exit Sub
SetPayedErr:
    Call RaiseError(MyUnhandledError, "FinanceSheet:SetPayed Method")
End Sub
'得到当前结算单的拆算项记录集
Public Function GetSplitItemRS(Optional VehicleID As String) As ADODB.Recordset
    AssertActiveUserValid m_oActiveUser, ERR_LuggageFinanceSheet
    On Error GoTo GetSplitDetailRSErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim szTempWheSql As String
    If VehicleID <> "" Then
        szTempWheSql = " AND vehicle_ID = " & TransFieldValueToString(VehicleID)
    Else
        szTempWheSql = " "
    End If
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT * FROM fin_luggage_split_detail_lst " _
        & " WHERE fin_sheet_id = " & TransFieldValueToString(mszSheetID) & szTempWheSql
    Set rsTemp = oDb.Execute(szSql)
    Set GetSplitItemRS = rsTemp
    Exit Function
GetSplitDetailRSErr:
    Call RaiseError(MyUnhandledError, "FinanceSheet:GetSplitDetailRS Method")

End Function

'得到当前结算单的结算明细记录集
'##ModelId=3E1B7B3F0011
Public Function GetFinSheetDetailRS() As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_LuggageFinanceSheet
    On Error GoTo GetSplitDetailRSErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT * FROM fin_luggage_sheet_lst " _
        & " WHERE fin_sheet_id = " & TransFieldValueToString(mszSheetID)
    Set rsTemp = oDb.Execute(szSql)
    Set GetFinSheetDetailRS = rsTemp
    Exit Function
GetSplitDetailRSErr:
    Call RaiseError(MyUnhandledError, "FinanceSheet:GetSplitDetailRS Method")
End Function

'查询结算单相关的行包签发单信息,返回签发单记录集
'##ModelId=3E28CF4F0137
Public Function GetLuggageSheetRS() As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_LuggageFinanceSheet
    On Error GoTo GetLuggageSheetRSErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT * FROM fin_sheet_to_luggage_lst S,luggage_sheet_info l " _
        & " WHERE l.sheet_id = s.luggage_sheet_id AND s.fin_sheet_id = " & TransFieldValueToString(mszSheetID)
    Set rsTemp = oDb.Execute(szSql)
    Set GetLuggageSheetRS = rsTemp
    Exit Function
GetLuggageSheetRSErr:
    Call RaiseError(MyUnhandledError, "FinanceSheet:GetLuggageSheetRS Method")
End Function

'拆算权限验证
Public Sub SplitMan()
    AssertActiveUserValid m_oActiveUser, ERR_LuggageFinanceSheet
    AssertHaveRight m_oActiveUser, RIGHT_FinancesheetAdd
End Sub

Private Function CaleLugValue(Formula As String, atProtocolChargeItem() As String) As Double
    Dim szTemp As String
    Dim dbValue As Double
    szTemp = FullStr(Trim(Formula), atProtocolChargeItem)
    'szTemp = "(9*10+12*10)*0.8"
    If Trim(szTemp) <> "" Then
        dbValue = moCalc.CalculateIt(szTemp)
    Else
        dbValue = 0
    End If
    CaleLugValue = dbValue
End Function
Private Function FullStr(pszFormula As String, paszKeyValue() As String) As String
    '将公式字符串转换为值
    
    Dim nCount As Integer
    Dim i As Integer
    Dim szTemp As String, nLastSeat As Integer
    Dim nFormulaLen As Integer
    Dim nTemp As Integer, j As Integer
    Dim bContinue As Boolean
    Dim nTag As Integer
    pszFormula = pszFormula
    nCount = UBound(paszKeyValue, 1)
    nFormulaLen = Len(pszFormula)
    nLastSeat = 1
    bContinue = True
    nTag = 1
    Do While bContinue
        bContinue = False

        nTemp = GetLowestSeat(nLastSeat, pszFormula, paszKeyValue, j)

        If nTemp > 0 Then
            szTemp = szTemp & Mid(pszFormula, nLastSeat, nTemp - nLastSeat)
            If paszKeyValue(j, 2) < 0 Then
'               nTag = -1
'               paszKeyValue(j, 2) = Abs(paszKeyValue(j, 2))
                szTemp = szTemp & "(0-" & Abs(paszKeyValue(j, 2)) & ")"
            Else
                szTemp = szTemp & paszKeyValue(j, 2)
            End If
'            szTemp = szTemp & paszKeyValue(j, 2)
            nTemp = Len(paszKeyValue(j, 1)) + nTemp
            nLastSeat = nTemp
            bContinue = True
        End If

    Loop
    If nLastSeat <= nFormulaLen Then
        szTemp = szTemp & Right(pszFormula, nFormulaLen - nLastSeat + 1)
    End If
    FullStr = szTemp
End Function

'##ModelId=3DDD20600241
Private Function GetLowestSeat(pnStart As Integer, pszFormula As String, paszKeyValue() As String, pnIndex As Integer) As Integer '得到最小字符串
    Dim nCount As Integer
    Dim nResult As Integer, i As Integer
    Dim nLen As Integer
    Dim nIndex As Integer
    nCount = UBound(paszKeyValue, 1)
    nIndex = 0
    nResult = Len(pszFormula)
    nLen = Len(pszFormula)
    For i = 1 To nCount
        nLen = InStr(pnStart, pszFormula, paszKeyValue(i, 1), vbTextCompare)
        If nLen < nResult And nLen <> 0 And IsLastCharLegalSeperator(pszFormula, Len(paszKeyValue(i, 1)) + nLen) Then
            nResult = nLen
            nIndex = i
        End If
    Next
    If nIndex = 0 Then
        GetLowestSeat = 0
    Else
        GetLowestSeat = nResult
    End If
    pnIndex = nIndex
End Function

'##ModelId=3DDD2060030C
Private Function IsLastCharLegalSeperator(ByVal pszString As String, ByVal pnIndex As Integer) As Boolean   '判断是否是最后公式项
    Dim szChar As String
    On Error GoTo Error_HANDle
    szChar = Mid(pszString, pnIndex, 1)
    If szChar = " " Or szChar = "+" Or szChar = "-" Or szChar = "*" Or szChar = "/" Or szChar = ")" Or szChar = "" Then
        IsLastCharLegalSeperator = True
    Else
        IsLastCharLegalSeperator = False
    End If
    Exit Function
Error_HANDle:
    IsLastCharLegalSeperator = True
End Function



Private Function GetSplitAllItem(AcceptType As Integer) As String()
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim aszTempItemp() As String
    Dim szAcceptType As String
    Dim i As Integer
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT chinese_name,accept_type FROM luggage_price_item_info WHERE accept_type = " & TransFieldValueToString(AcceptType)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then
        RaiseError ERR_FinSheetPriceItemNoExist '票价项不存在
    End If
    ReDim maszSplitLuggageItem(1 To 12, 1 To 2)
    maszSplitLuggageItem(1, 1) = "总价格"
    maszSplitLuggageItem(2, 1) = "总件数"
    For i = 3 To rsTemp.RecordCount + 2
        szAcceptType = GetLuggageTypeString(FormatDbValue(rsTemp!accept_type))
        maszSplitLuggageItem(i, 1) = "[" & szAcceptType & "]" & FormatDbValue(rsTemp!chinese_name)
'        maszSplitLuggageItem(i, 1) = FormatDbValue(rsTemp!chinese_name)
        rsTemp.MoveNext
    Next i
    GetSplitAllItem = maszSplitLuggageItem
    Set rsTemp = Nothing
End Function

'##ModelId=3E29E90D00E1
Public Property Get Remark() As String
   Let Remark = mszRemark
End Property

'##ModelId=3E29E90C034D
Public Property Let Remark(ByVal Value As String)
    Let mszRemark = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E29E90C025D
Public Property Get StatusString() As String
    If mnStatus <> -1 Then
        Select Case mnStatus
'            Case 0
'                mszStatusString = "未结"
            Case 1
                mszStatusString = "已结"
            Case 0
                mszStatusString = "作废"
        End Select
    End If
   Let StatusString = mszStatusString
End Property

'##ModelId=3E29E90C016C
Public Property Get Status() As Integer
   Let Status = mnStatus
End Property

'##ModelId=3E29E90B03E2
Public Property Let Status(ByVal Value As Integer)
    Let mnStatus = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj

End Property

'##ModelId=3E29E90B02FB
Public Property Get OperatorName() As String
   Let OperatorName = mszOperatorName
End Property

'##ModelId=3E29E90B019D
Public Property Let OperatorName(ByVal Value As String)
    Let mszOperatorName = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E29E90B00B7
Public Property Get OperateTime() As Date
   Let OperateTime = mdtOperateTime
End Property

'##ModelId=3E29E90A0340
Public Property Let OperateTime(ByVal Value As Date)
    Let mdtOperateTime = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E29E90A026E
Public Property Get BusOwnerName() As String
   Let BusOwnerName = mszBusOwnerName
   
End Property

'##ModelId=3E29E90A0191
Public Property Get BusOwnerID() As String
   Let BusOwnerID = mszBusOwnerID
End Property

'##ModelId=3E29E90A00BF
Public Property Get RouteName() As String
   Let RouteName = mszRouteName
End Property

'##ModelId=3E29E90903CB
Public Property Get RouteID() As String
   Let RouteID = mszRouteID
End Property

'##ModelId=3E29E90902F8
Public Property Get SplitCompanyName() As String
   Let SplitCompanyName = mszSplitCompanyName
End Property

'##ModelId=3E29E9090226
Public Property Get SplitCompanyID() As String
   Let SplitCompanyID = mszSplitCompanyID
End Property

'##ModelId=3E29E909015E
Public Property Get CompanyName() As String
   Let CompanyName = mszCompanyName
End Property

'##ModelId=3E29E9090096
Public Property Get CompanyID() As String
   Let CompanyID = mszCompanyID
End Property

'##ModelId=3E29E90803BF
Public Property Get VehicleLicense() As String
   Let VehicleLicense = mszVehicleLicense
End Property

'##ModelId=3E29E90802F7
Public Property Get VehicleID() As String
   Let VehicleID = mszVehicleID
End Property

'##ModelId=3E29E9080239
Public Property Get ActSplitPrice() As Double
   Let ActSplitPrice = mdbActSplitPrice
End Property

'##ModelId=3E29E908010C
Public Property Let ActSplitPrice(ByVal Value As Double)
    Let mdbActSplitPrice = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E29E9080058
Public Property Get NeedSplitPrice() As Double
   Let NeedSplitPrice = mdbNeedSplitPrice
End Property

'##ModelId=3E29E9070382
Public Property Get TotalPrice() As Double
   Let TotalPrice = mdbTotalPrice
End Property

'##ModelId=3E29E90702CE
Public Property Get StopSettleDate() As Date
   Let StopSettleDate = mdtStopSettleDate
End Property

'##ModelId=3E29E90701B5
Public Property Let StopSettleDate(ByVal Value As Date)
    Let mdtStopSettleDate = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E29E9070101
Public Property Get StartSettleDate() As Date
   Let StartSettleDate = mdtStartSettleDate
End Property

'##ModelId=3E29E90603D0
Public Property Let StartSettleDate(ByVal Value As Date)
    Let mdtStartSettleDate = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E29E9060330
Public Property Get SettleMonth() As Date
   Let SettleMonth = mdtSettleMonth
End Property

'##ModelId=3E29E906022C
Public Property Let SettleMonth(ByVal Value As Date)
    Let mdtSettleMonth = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E29E906018C
Public Property Get SellStationName() As String
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    
    If mszSellStationID <> "" Then
        szSql = "SELECT * FROM sell_station_info WHERE sell_station_id = " & TransFieldValueToString(mszSellStationID)
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount <> 0 Then
            mszSellStationName = FormatDbValue(rsTemp!sell_station_name)
            Let SellStationName = mszSellStationName
        End If
    End If
End Property

'##ModelId=3E29E90600E1
Public Property Get SellStationID() As String
    Let SellStationID = mszSellStationID
End Property

'##ModelId=3E29E90503CF
Public Property Let SellStationID(ByVal Value As String)
    Let mszSellStationID = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E29E905032F
Public Property Get SheetID() As String
   Let SheetID = mszSheetID
End Property


'##ModelId=3E29E905023E
Public Property Let SheetID(ByVal Value As String)
    Let mszSheetID = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

Public Property Get AcceptType() As String
    Let AcceptType = mszAcceptType
End Property
Public Property Let AcceptType(ByVal Value As String)
    Let mszAcceptType = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

Public Property Get ProtocolID() As String
    Let ProtocolID = mszProtocolID
End Property
Public Property Get ProtocolName() As String
    Let ProtocolName = mszProtocolName
End Property

Public Property Get SplitObjectType() As Integer
    Let SplitObjectType = mnSplitObjectType
End Property

Public Property Let SplitObjectType(ByVal Value As Integer)
    Let mnSplitObjectType = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

Public Property Get SplitObjectID() As String
    Let SplitObjectID = mszSplitObjectID
End Property
Public Property Let SplitObjectID(ByVal Value As String)
    Let mszSplitObjectID = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

Public Property Get SplitObjectName() As String
    Let SplitObjectName = mszSplitObjectName
End Property
Public Property Let SplitObjectName(ByVal Value As String)
    Let mszSplitObjectName = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property



'行包签发单预览
Public Function PreviewSplitCarrySheets(SheetID() As String) As String()
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim szaTemp() As String
    Dim szaTemp1() As String
    Dim i As Integer, j As Integer
    Dim nCount As Integer
    Dim rsChargeItem As Recordset
    Dim dbChargeMoney As Double
    Dim dbNeedSplitPrice As Double
    Dim nChargeType As Integer
    Dim szformuarName As String
    Dim nTempAcceptType  As Integer
    Dim szProtocoID As String, szProtocolName As String, szAcceptType As String
    Dim szSheetid As String
    szaTemp = SheetID
    nCount = ArrayLength(szaTemp)
    If nCount = 0 Then RaiseError ERR_FinSheetNotExist
    
    
    szSheetid = "("
    For i = 1 To ArrayLength(szaTemp) - 1
        szSheetid = szSheetid & TransFieldValueToString(szaTemp(i)) & ","
    Next i
    szSheetid = szSheetid & TransFieldValueToString(szaTemp(i)) & ")"
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT vehicle_id,MAX(accept_type) accept_type, MAX(protocol_id) protocol_id,SUM(total_price) total_price , " _
        & " SUM(baggage_number) baggage_number, SUM(price_item_1) price_item_1,SUM(price_item_2) price_item_2 ," _
        & " SUM(price_item_3) price_item_3,SUM(price_item_4) price_item_4,SUM(price_item_5) price_item_5,SUM(price_item_6) price_item_6 , " _
        & " SUM(price_item_7) price_item_7,SUM(price_item_8) price_item_8,SUM(price_item_9) price_item_9,SUM(price_item_10) price_item_10 , " _
        & " MAX(protocol_name) protocol_name,MAX(license_tag_no) license_tag_no,MAX(transport_company_id) transport_company_id , " _
        & " MAX(transport_company_short_name) transport_company_short_name,MAX(owner_id) owner_id,MAX(owner_name) owner_name , " _
        & " MAX(split_company_id) split_company_id,MAX(splict_company_short_name) split_company_name,MAX(route_id) route_id , " _
        & " MAX(route_name) route_name " _
        & " FROM luggage_sheet_info " _
        & " WHERE 1=1 " & IIf(szSheetid <> "", " AND sheet_id IN " & szSheetid, "") _
        & " AND settlement_status = 0 AND valid_mark =1 " _
        & " GROUP BY vehicle_id ,accept_type "
    Set rsTemp = oDb.Execute(szSql)
    ReDim szaTemp1(1 To rsTemp.RecordCount, 1 To 17)
    Dim rsProtocol As Recordset
    Dim rsChargeItemTemp As Recordset
    For i = 1 To rsTemp.RecordCount

        szSql = " SELECT c.protocol_id,c.accept_type,c.charge_code,c.charge_type,c.charge_name,c.formular_name,c.charge_money, " _
            & " c.formula_comment ,p.protocol_name  " _
            & " FROM vehicle_lug_protocol_info v,luggage_settle_protocol_info p,fin_luggage_chargeitem_lst c " _
            & " WHERE v.protocol_id  = p.protocol_id AND p.protocol_id = c.protocol_id  " _
            & " AND v.vehicle_id = " & TransFieldValueToString(FormatDbValue(rsTemp!vehicle_id)) _
            & " AND v.accept_type = " & TransFieldValueToString(FormatDbValue(rsTemp!accept_type))
        Set rsChargeItem = oDb.Execute(szSql)
        If rsChargeItem.RecordCount > 0 Then
        
        '注释的为默认协议代替的部分,现不起用此方式
'            szSql = "SELECT * FROM luggage_settle_protocol_info WHERE default_mark!=2"
'            Set rsProtocol = oDb.Execute(szSql)
'            If rsProtocol.RecordCount > 0 Then '启用默认
'                szSql = "SELECT c.protocol_id,c.accept_type,c.charge_code,c.charge_type,c.charge_name,c.formular_name,c.charge_money,c.formula_comment " _
'                    & " FROM fin_luggage_chargeitem_lst c " _
'                    & " WHERE c.protocol_id = " & TransFieldValueToString(FormatDbValue(rsProtocol!protocol_id))
'                Set rsChargeItemTemp = oDb.Execute(szSql)
                    
'        dbChargeMoney = 0
'        dbNeedSplitPrice = 0
'        For j = 1 To rsChargeItemTemp.RecordCount
'             nChargeType = FormatDbValue(rsChargeItemTemp!charge_type)
'             szProtocoID = FormatDbValue(rsProtocol!protocol_id)
'             szProtocolName = FormatDbValue(rsProtocol!protocol_name)
'            If nChargeType <> 0 Then
'                szformuarName = FormatDbValue(rsChargeItemTemp!formula_comment)
'                '得到每辆车的费用项
'                '中文名称
'                nTempAcceptType = GetLuggageTypeInt(mszAcceptType)
'                maszSplitLuggageItem = GetSplitAllItem(rsChargeItemTemp!accept_type)
'                '各名称的值
'                maszSplitLuggageItem(1, 2) = FormatDbValue(rsTemp!total_price)
'                maszSplitLuggageItem(2, 2) = FormatDbValue(rsTemp!baggage_number)
'                maszSplitLuggageItem(3, 2) = FormatDbValue(rsTemp!price_item_1)
'                maszSplitLuggageItem(4, 2) = FormatDbValue(rsTemp!price_item_2)
'                maszSplitLuggageItem(5, 2) = FormatDbValue(rsTemp!price_item_3)
'                maszSplitLuggageItem(6, 2) = FormatDbValue(rsTemp!price_item_4)
'                maszSplitLuggageItem(7, 2) = FormatDbValue(rsTemp!price_item_5)
'                maszSplitLuggageItem(8, 2) = FormatDbValue(rsTemp!price_item_6)
'                maszSplitLuggageItem(9, 2) = FormatDbValue(rsTemp!price_item_7)
'                maszSplitLuggageItem(10, 2) = FormatDbValue(rsTemp!price_item_8)
'                maszSplitLuggageItem(11, 2) = FormatDbValue(rsTemp!price_item_9)
'                maszSplitLuggageItem(12, 2) = FormatDbValue(rsTemp!price_item_10)
'                dbChargeMoney = CaleLugValue(szformuarName, maszSplitLuggageItem)  '跟据协议得出折算款
'                dbNeedSplitPrice = dbNeedSplitPrice + dbChargeMoney
'            Else
'                dbChargeMoney = FormatDbValue(rsChargeItemTemp!charge_money)  '为固定费用项
'                dbNeedSplitPrice = dbNeedSplitPrice + dbChargeMoney
'            End If
'            rsChargeItemTemp.MoveNext
'        Next j
'        szaTemp1(i, 1) = FormatDbValue(rsTemp!vehicle_id)
'        szaTemp1(i, 2) = FormatDbValue(rsTemp!license_tag_no)
'        szaTemp1(i, 3) = dbNeedSplitPrice
'        szaTemp1(i, 4) = szProtocoID
'        szaTemp1(i, 5) = szProtocolName
'        szaTemp1(i, 6) = FormatDbValue(rsTemp!total_price)
'        szaTemp1(i, 7) = FormatDbValue(rsTemp!price_item_1)
'        szaTemp1(i, 8) = FormatDbValue(rsTemp!price_item_2)
'        szaTemp1(i, 9) = FormatDbValue(rsTemp!price_item_3)
'        szaTemp1(i, 10) = FormatDbValue(rsTemp!price_item_4)
'        szaTemp1(i, 11) = FormatDbValue(rsTemp!price_item_5)
'        szaTemp1(i, 12) = FormatDbValue(rsTemp!price_item_6)
'        szaTemp1(i, 13) = FormatDbValue(rsTemp!price_item_7)
'        szaTemp1(i, 14) = FormatDbValue(rsTemp!price_item_8)
'        szaTemp1(i, 15) = FormatDbValue(rsTemp!price_item_9)
'        szaTemp1(i, 16) = FormatDbValue(rsTemp!price_item_10)
'        szaTemp1(i, 17) = FormatDbValue(rsTemp!accept_type)
'        GoTo here:
'            Else
'                RaiseError ERR_FinSheetProtocolNoExist '车辆协议不存在
'            End If
'        End If
        '车辆协议存在
        dbChargeMoney = 0
        dbNeedSplitPrice = 0
        For j = 1 To rsChargeItem.RecordCount
             nChargeType = FormatDbValue(rsChargeItem!charge_type)
             szProtocoID = FormatDbValue(rsChargeItem!protocol_id)
             szProtocolName = FormatDbValue(rsChargeItem!protocol_name)
            If nChargeType <> 0 Then
                szformuarName = FormatDbValue(rsChargeItem!formula_comment)
                '得到每辆车的费用项
                '中文名称
                nTempAcceptType = GetLuggageTypeInt(mszAcceptType)
                maszSplitLuggageItem = GetSplitAllItem(rsChargeItem!accept_type)
                '各名称的值
                maszSplitLuggageItem(1, 2) = FormatDbValue(rsTemp!total_price)
                maszSplitLuggageItem(2, 2) = FormatDbValue(rsTemp!baggage_number)
                maszSplitLuggageItem(3, 2) = FormatDbValue(rsTemp!price_item_1)
                maszSplitLuggageItem(4, 2) = FormatDbValue(rsTemp!price_item_2)
                maszSplitLuggageItem(5, 2) = FormatDbValue(rsTemp!price_item_3)
                maszSplitLuggageItem(6, 2) = FormatDbValue(rsTemp!price_item_4)
                maszSplitLuggageItem(7, 2) = FormatDbValue(rsTemp!price_item_5)
                maszSplitLuggageItem(8, 2) = FormatDbValue(rsTemp!price_item_6)
                maszSplitLuggageItem(9, 2) = FormatDbValue(rsTemp!price_item_7)
                maszSplitLuggageItem(10, 2) = FormatDbValue(rsTemp!price_item_8)
                maszSplitLuggageItem(11, 2) = FormatDbValue(rsTemp!price_item_9)
                maszSplitLuggageItem(12, 2) = FormatDbValue(rsTemp!price_item_10)
                dbChargeMoney = CaleLugValue(szformuarName, maszSplitLuggageItem)  '跟据协议得出折算款
                dbNeedSplitPrice = dbNeedSplitPrice + dbChargeMoney
            Else
                dbChargeMoney = FormatDbValue(rsChargeItem!charge_money)  '为固定费用项
                dbNeedSplitPrice = dbNeedSplitPrice + dbChargeMoney
            End If
            rsChargeItem.MoveNext
        Next j
        szaTemp1(i, 1) = FormatDbValue(rsTemp!vehicle_id)
        szaTemp1(i, 2) = FormatDbValue(rsTemp!license_tag_no)
        szaTemp1(i, 3) = dbNeedSplitPrice
        szaTemp1(i, 4) = szProtocoID
        szaTemp1(i, 5) = szProtocolName
        szaTemp1(i, 6) = FormatDbValue(rsTemp!total_price)
        szaTemp1(i, 7) = FormatDbValue(rsTemp!price_item_1)
        szaTemp1(i, 8) = FormatDbValue(rsTemp!price_item_2)
        szaTemp1(i, 9) = FormatDbValue(rsTemp!price_item_3)
        szaTemp1(i, 10) = FormatDbValue(rsTemp!price_item_4)
        szaTemp1(i, 11) = FormatDbValue(rsTemp!price_item_5)
        szaTemp1(i, 12) = FormatDbValue(rsTemp!price_item_6)
        szaTemp1(i, 13) = FormatDbValue(rsTemp!price_item_7)
        szaTemp1(i, 14) = FormatDbValue(rsTemp!price_item_8)
        szaTemp1(i, 15) = FormatDbValue(rsTemp!price_item_9)
        szaTemp1(i, 16) = FormatDbValue(rsTemp!price_item_10)
        szaTemp1(i, 17) = FormatDbValue(rsTemp!accept_type)
        rsTemp.MoveNext
     End If
    Next i
    
    PreviewSplitCarrySheets = szaTemp1
End Function



'行包签发单拆算
'
'##ModelId=3E1B7AF80113
Public Sub SplitCarrySheets(SheetID() As String, Optional Operator As String, Optional OperateDate As Date = cszEmptyDateStr)

'    On Error GoTo SplitCarrySheetsErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim aszTempSheetID() As String
    Dim nCount As Integer, i As Integer
    Dim szSheetid As String
    Dim rsTemp As Recordset
    Dim rsTemp1 As Recordset
    Dim rsTempProtocol As Recordset
    Dim szTempSql As String
    Dim szformuarName As String  '公式名称
    Dim nChargeType As Integer  ' 费用类型
    Dim dbChargeMoney As Double
    Dim dbTotalMoney As Double
    Dim dbCharge As Double
    Dim nTempAcceptType As Integer
    Dim rsTempVehicl As Recordset
    Dim rsTempNeedSplitOut As Recordset
    Dim rsTempCharge As Recordset
    Dim j As Integer
    Dim szSqlTotalPrice As String
    Dim rsTempTotalPrice As Recordset

    oDb.ConnectionString = GetConnectionStr(cszSplit)
    If m_nObjectStatus = ST_AddObj Then RaiseError ERR_FinSheetNoEditStatus '不是编辑状态
    aszTempSheetID = SheetID
    nCount = ArrayLength(aszTempSheetID)
    If nCount = 0 Then
        RaiseError ERR_FinSheetIsNull '没有签发单号
    End If
    szSheetid = "("
    For i = 1 To ArrayLength(aszTempSheetID) - 1
        szSheetid = szSheetid & TransFieldValueToString(aszTempSheetID(i)) & ","
    Next i
    szSheetid = szSheetid & TransFieldValueToString(aszTempSheetID(i)) & ")"


    '按VEHICLE＿ID　分组，统计出总价格
    '签发单状态为”有效“并且”结算标记“有0
    szSql = "SELECT l.vehicle_id,MAX(l.accept_type) accept_type, MAX(p.protocol_id) protocol_id,SUM(l.total_price) total_price , " _
        & " SUM(baggage_number) baggage_number, SUM(price_item_1) price_item_1,SUM(price_item_2) price_item_2 ," _
        & " SUM(price_item_3) price_item_3,SUM(price_item_4) price_item_4,SUM(price_item_5) price_item_5,SUM(price_item_6) price_item_6 , " _
        & " SUM(price_item_7) price_item_7,SUM(price_item_8) price_item_8,SUM(price_item_9) price_item_9,SUM(price_item_10) price_item_10 , " _
        & " MAX(p.protocol_name) protocol_name,MAX(l.license_tag_no) license_tag_no,MAX(transport_company_id) transport_company_id , " _
        & " MAX(transport_company_short_name) transport_company_short_name,MAX(owner_id) owner_id,MAX(owner_name) owner_name , " _
        & " MAX(split_company_id) split_company_id,MAX(splict_company_short_name) split_company_name,MAX(route_id) route_id , " _
        & " MAX(route_name) route_name " _
        & " FROM luggage_sheet_info l,vehicle_lug_protocol_info v,luggage_settle_protocol_info p" _
        & " WHERE l.accept_type=v.accept_type  " & IIf(szSheetid <> "", " AND sheet_id IN " & szSheetid, "") _
        & " AND settlement_status = 0 AND valid_mark =1 " _
        & " AND p.protocol_id=v.protocol_id AND v.vehicle_id=l.vehicle_id " _
        & " GROUP BY l.vehicle_id  "
    Set rsTemp = oDb.Execute(szSql)

    If rsTemp.RecordCount = 0 Then
        RaiseError ERR_FinSheetIsNull '签发单不存在
    End If

    '查询“行包车辆结算明细表”如车辆以存在，则将这张表中的费用折算并累加拆出款，其它信息不做考虑
    szSql = "SELECT * FROM fin_luggage_sheet_lst " _
        & " WHERE vehicle_id = " & TransFieldValueToString(FormatDbValue(rsTemp!vehicle_id)) _
        & " AND fin_sheet_id = " & TransFieldValueToString(mszSheetID)
    Set rsTempVehicl = oDb.Execute(szSql)

    '更新fin_luggage_sheet_info 表中的总价
    Dim mdbTotalPrice1 As Double
    szSqlTotalPrice = "SELECT SUM(total_price) total_price FROM fin_luggage_sheet_info WHERE fin_sheet_id=" & TransFieldValueToString(mszSheetID)
    Set rsTempTotalPrice = oDb.Execute(szSqlTotalPrice)
    If rsTempTotalPrice!total_price <> 0 Then '如果当前结算单已有不同车辆,结算额累加
        mdbTotalPrice1 = FormatDbValue(rsTemp!price_item_1) + FormatDbValue(rsTempTotalPrice!total_price)
    Else
        mdbTotalPrice1 = FormatDbValue(rsTemp!price_item_1)
    End If
    
    
    oDb.BeginTrans
    '开始事务
    
    szSql = "UPDATE fin_luggage_sheet_info SET total_price = " & TransFieldValueToString(mdbTotalPrice1) _
        & " WHERE fin_sheet_id = " & TransFieldValueToString(mszSheetID)
    oDb.Execute (szSql)


If rsTempVehicl.RecordCount <> 0 Then  '同一辆车
        '得出折算款
        '==============================================================================================
        '跟据协议号查询出公式内容，如为1“公式计算费用”者有公式内容，否则取出“款项”
        If FormatDbValue(rsTemp!protocol_id) = "" Then RaiseError ERR_FinSheetProtocolNoExist '车辆协议不存在
'        szSql = "SELECT * FROM fin_luggage_chargeitem_lst " _
'            & " WHERE protocol_id = " & TransFieldValueToString(FormatDbValue(rsTemp!protocol_id)) _
'            & " AND accept_type = " & TransFieldValueToString(FormatDbValue(rsTemp!accept_type))
        szSql = " SELECT c.protocol_id,c.accept_type,c.charge_code,c.charge_type,c.charge_name,c.formular_name,c.charge_money,c.formula_comment " _
            & " FROM vehicle_lug_protocol_info v,luggage_settle_protocol_info p,fin_luggage_chargeitem_lst c " _
            & " WHERE v.protocol_id  = p.protocol_id AND p.protocol_id = c.protocol_id " _
            & " AND v.vehicle_id = " & TransFieldValueToString(FormatDbValue(rsTemp!vehicle_id)) _
            & " AND v.accept_type = " & TransFieldValueToString(FormatDbValue(rsTemp!accept_type))

        Set rsTempProtocol = oDb.Execute(szSql)
        If rsTempProtocol.RecordCount = 0 Then
            RaiseError ERR_FinSheetProtocolNoExist '车辆协议不存在
        End If
        szSql = "SELECT SUM(total_price) total_price,SUM(need_split_out) NeedSplitOut FROM fin_luggage_sheet_lst " _
                & " WHERE fin_sheet_id = " & TransFieldValueToString(mszSheetID) _
                & " AND vehicle_id = " & TransFieldValueToString(FormatDbValue(rsTemp!vehicle_id))
        Set rsTempNeedSplitOut = oDb.Execute(szSql)
        mdbTotalPrice = FormatDbValue(rsTempNeedSplitOut!total_price) + FormatDbValue(rsTemp!price_item_1)
        mdbNeedSplitPrice = FormatDbValue(rsTempNeedSplitOut!NeedSplitOut)

        '跟据不同的费用项将拆出款累加起来
        For j = 1 To rsTempProtocol.RecordCount
            szSql = "SELECT * FROM fin_luggage_split_detail_lst " _
                & " WHERE fin_sheet_id = " & TransFieldValueToString(mszSheetID) _
                & " AND vehicle_id = " & TransFieldValueToString(FormatDbValue(rsTemp!vehicle_id)) _
                & " AND charge_code = " & TransFieldValueToString(FormatDbValue(rsTempProtocol!charge_code))
            Set rsTempCharge = oDb.Execute(szSql)


            nChargeType = FormatDbValue(rsTempProtocol!charge_type)  '费用类型 0固定费用 1公式计算费用
            If nChargeType <> 0 Then
                szformuarName = FormatDbValue(rsTempProtocol!formula_comment)
                '得到每辆车的费用项
                '中文名称
                nTempAcceptType = FormatDbValue(rsTempProtocol!accept_type)
                maszSplitLuggageItem = GetSplitAllItem(nTempAcceptType)
                '各名称的值
                maszSplitLuggageItem(1, 2) = FormatDbValue(rsTemp!total_price)
                maszSplitLuggageItem(2, 2) = FormatDbValue(rsTemp!baggage_number)
                maszSplitLuggageItem(3, 2) = FormatDbValue(rsTemp!price_item_1)
                maszSplitLuggageItem(4, 2) = FormatDbValue(rsTemp!price_item_2)
                maszSplitLuggageItem(5, 2) = FormatDbValue(rsTemp!price_item_3)
                maszSplitLuggageItem(6, 2) = FormatDbValue(rsTemp!price_item_4)
                maszSplitLuggageItem(7, 2) = FormatDbValue(rsTemp!price_item_5)
                maszSplitLuggageItem(8, 2) = FormatDbValue(rsTemp!price_item_6)
                maszSplitLuggageItem(9, 2) = FormatDbValue(rsTemp!price_item_7)
                maszSplitLuggageItem(10, 2) = FormatDbValue(rsTemp!price_item_8)
                maszSplitLuggageItem(11, 2) = FormatDbValue(rsTemp!price_item_9)
                maszSplitLuggageItem(12, 2) = FormatDbValue(rsTemp!price_item_10)
                dbChargeMoney = CaleLugValue(szformuarName, maszSplitLuggageItem)  '跟据协议得出折算款
                If rsTempCharge.RecordCount = 0 Then
                    dbCharge = dbChargeMoney
                Else
                    dbCharge = FormatDbValue(rsTempCharge!split_out_money) + dbChargeMoney
                End If
                mdbNeedSplitPrice = mdbNeedSplitPrice + dbChargeMoney
            Else
                dbChargeMoney = FormatDbValue(rsTempProtocol!charge_money)  '为固定费用项
                If rsTempCharge.RecordCount = 0 Then
                    dbCharge = dbChargeMoney
                Else
                    dbCharge = FormatDbValue(rsTempCharge!split_out_money) + dbChargeMoney
                End If
                mdbNeedSplitPrice = mdbNeedSplitPrice + dbChargeMoney
            End If
            '更新”行包车辆结算明细表“应拆金额项

            szSql = "UPDATE fin_luggage_sheet_lst SET total_price = " & TransFieldValueToString(mdbTotalPrice) _
                & " ,need_split_out = " & TransFieldValueToString(mdbNeedSplitPrice) _
                & " WHERE fin_sheet_id = " & TransFieldValueToString(mszSheetID) _
                & " AND vehicle_id = " & TransFieldValueToString(rsTemp!vehicle_id)
            oDb.Execute szSql

            Dim rsTempNeedSplitOut1 As Recordset
            Dim mdbNeedSplitPrice1 As Double
            szSql = "SELECT SUM(need_split_out) need_split_out FROM fin_luggage_sheet_info WHERE fin_sheet_id=" & TransFieldValueToString(mszSheetID)
            Set rsTempNeedSplitOut1 = oDb.Execute(szSql)
            If FormatDbValue(rsTempNeedSplitOut1!need_split_out) <> 0 Then
                mdbNeedSplitPrice1 = dbChargeMoney + FormatDbValue(rsTempNeedSplitOut1!need_split_out)
            End If
            szSql = "UPDATE fin_luggage_sheet_info SET " _
                & " need_split_out = " & TransFieldValueToString(mdbNeedSplitPrice1) _
                & " WHERE fin_sheet_id = " & TransFieldValueToString(mszSheetID)
            oDb.Execute szSql
            
            If rsTempCharge.RecordCount = 0 Then
                '同一车辆不同费用类型
                szSql = "INSERT fin_luggage_split_detail_lst ( fin_sheet_id , vehicle_id , charge_code , charge_name , split_out_money ) " _
                    & " VALUES ( " _
                    & TransFieldValueToString(mszSheetID) & "," _
                    & TransFieldValueToString(mszVehicleID) & "," _
                    & TransFieldValueToString(FormatDbValue(rsTempProtocol!charge_code)) & "," _
                    & TransFieldValueToString(FormatDbValue(rsTempProtocol!charge_name)) & "," _
                    & TransFieldValueToString(dbChargeMoney) & ")"
                oDb.Execute szSql
             End If
            
                '更新结算单拆帐票项表
                szSql = "UPDATE fin_luggage_split_detail_lst SET split_out_money = " & TransFieldValueToString(dbCharge) _
                    & " WHERE fin_sheet_id = " & TransFieldValueToString(mszSheetID) _
                    & " AND vehicle_id = " & TransFieldValueToString(rsTemp!vehicle_id) _
                    & " AND charge_code = " & TransFieldValueToString(rsTempProtocol!charge_code)
                    
                oDb.Execute szSql
            rsTempProtocol.MoveNext
        Next j
Else    '不同车辆

        For i = 1 To rsTemp.RecordCount
            '得到各字段
            mszVehicleID = FormatDbValue(rsTemp!vehicle_id)
    '        mszAcceptType = FormatDbValue(rsTemp!accept_type)
    '        mszSellStationID = FormatDbValue(rsTemp!sell_station_id)
            mszProtocolID = FormatDbValue(rsTemp!protocol_id)
            mszProtocolName = FormatDbValue(rsTemp!protocol_name)
            mszVehicleLicense = FormatDbValue(rsTemp!license_tag_no)
            mszCompanyID = FormatDbValue(rsTemp!transport_company_id)
            mszCompanyName = FormatDbValue(rsTemp!transport_company_short_name)
            mszBusOwnerID = FormatDbValue(rsTemp!owner_id)
            mszBusOwnerName = FormatDbValue(rsTemp!owner_name)
            mszSplitCompanyID = FormatDbValue(rsTemp!split_company_id)
            mszSplitCompanyName = FormatDbValue(rsTemp!split_company_name)
            mszRouteID = FormatDbValue(rsTemp!route_id)
            mszRouteName = FormatDbValue(rsTemp!route_name)
    '        mszSellStationID = FormatDbValue(rsTemp!sell_station_id)
            mdbTotalPrice = FormatDbValue(rsTemp!price_item_1)

            '插入到fin_luggage_sheet_lst表中
            szSql = "INSERT fin_luggage_sheet_lst (fin_sheet_id,vehicle_id,sell_station_id,settle_month,accept_type,status,total_price, " _
                & " need_split_out,protocol_id,protocol_name,license_tag_no,transport_company_id,transport_company_short_name , " _
                & " owner_id,owner_name,split_company_id,split_company_name,route_id,route_name) " _
                & " VALUES ( " _
                & TransFieldValueToString(mszSheetID) & "," & TransFieldValueToString(mszVehicleID) & "," _
                & TransFieldValueToString(mszSellStationID) & "," & TransFieldValueToString(ToDBDate(mdtSettleMonth)) & "," _
                & TransFieldValueToString(GetLuggageTypeInt(mszAcceptType)) & "," & TransFieldValueToString(ELuggageSettleValidMark.LuggageValid) & "," _
                & TransFieldValueToString(mdbTotalPrice) & "," & TransFieldValueToString(mdbNeedSplitPrice) & "," _
                & TransFieldValueToString(mszProtocolID) & "," & TransFieldValueToString(mszProtocolName) & "," _
                & TransFieldValueToString(mszVehicleLicense) & "," & TransFieldValueToString(mszCompanyID) & "," _
                & TransFieldValueToString(mszCompanyName) & "," & TransFieldValueToString(mszBusOwnerID) & "," _
                & TransFieldValueToString(mszBusOwnerName) & "," & TransFieldValueToString(mszSplitCompanyID) & "," _
                & TransFieldValueToString(mszSplitCompanyName) & "," & TransFieldValueToString(mszRouteID) & "," _
                & TransFieldValueToString(mszRouteName) & ")"
            oDb.Execute (szSql)
'            '更新fin_luggage_sheet_info 表中的总价
'            szSqlTotalPrice = "SELECT SUM(total_price) total_price FROM fin_luggage_sheet_info WHERE fin_sheet_id=" & TransFieldValueToString(mszSheetID)
'            Set rsTempTotalPrice = oDb.Execute(szSqlTotalPrice)
'            If rsTempTotalPrice!total_price <> 0 Then '如果当前结算单已有不同车辆,结算额累加
'                mdbTotalPrice = mdbTotalPrice + FormatDbValue(rsTempTotalPrice!total_price)
'            End If
'            szSql = "UPDATE fin_luggage_sheet_info SET total_price = " & TransFieldValueToString(mdbTotalPrice) _
'                & " WHERE fin_sheet_id = " & TransFieldValueToString(mszSheetID)
'            oDb.Execute (szSql)
            '得出折算款

            '==============================================================================================
            '跟据协议号查询出公式内容，如为1“公式计算费用”者有公式内容，否则取出“款项”
'            szSql = "SELECT * FROM fin_luggage_chargeitem_lst " _
'                & " WHERE protocol_id = " & TransFieldValueToString(mszProtocolID) _
'                & " AND accept_type = " & TransFieldValueToString(GetLuggageTypeInt(mszAcceptType))
            szSql = " SELECT c.protocol_id,c.accept_type,c.charge_code,c.charge_type,c.charge_name,c.formular_name,c.charge_money,c.formula_comment " _
                & " FROM vehicle_lug_protocol_info v,luggage_settle_protocol_info p,fin_luggage_chargeitem_lst c " _
                & " WHERE v.protocol_id  = p.protocol_id AND p.protocol_id = c.protocol_id" _
                & " AND v.vehicle_id = " & TransFieldValueToString(FormatDbValue(rsTemp!vehicle_id)) _
                & " AND v.accept_type = " & TransFieldValueToString(FormatDbValue(rsTemp!accept_type))
            Set rsTempProtocol = oDb.Execute(szSql)
            If rsTempProtocol.RecordCount = 0 Then
                RaiseError ERR_FinSheetProtocolNoExist '   协议未设定费用项
            End If
            '跟据不同的费用项将拆出款累加起来
            For j = 1 To rsTempProtocol.RecordCount
                nChargeType = FormatDbValue(rsTempProtocol!charge_type)
                If nChargeType <> 0 Then
                    szformuarName = FormatDbValue(rsTempProtocol!formula_comment)

                    '得到每辆车的费用项
                    '中文名称
                    nTempAcceptType = GetLuggageTypeInt(mszAcceptType)
                    maszSplitLuggageItem = GetSplitAllItem(rsTempProtocol!accept_type)
                    '各名称的值
                    maszSplitLuggageItem(1, 2) = FormatDbValue(rsTemp!total_price)
                    maszSplitLuggageItem(2, 2) = FormatDbValue(rsTemp!baggage_number)
                    maszSplitLuggageItem(3, 2) = FormatDbValue(rsTemp!price_item_1)
                    maszSplitLuggageItem(4, 2) = FormatDbValue(rsTemp!price_item_2)
                    maszSplitLuggageItem(5, 2) = FormatDbValue(rsTemp!price_item_3)
                    maszSplitLuggageItem(6, 2) = FormatDbValue(rsTemp!price_item_4)
                    maszSplitLuggageItem(7, 2) = FormatDbValue(rsTemp!price_item_5)
                    maszSplitLuggageItem(8, 2) = FormatDbValue(rsTemp!price_item_6)
                    maszSplitLuggageItem(9, 2) = FormatDbValue(rsTemp!price_item_7)
                    maszSplitLuggageItem(10, 2) = FormatDbValue(rsTemp!price_item_8)
                    maszSplitLuggageItem(11, 2) = FormatDbValue(rsTemp!price_item_9)
                    maszSplitLuggageItem(12, 2) = FormatDbValue(rsTemp!price_item_10)
                    dbChargeMoney = CaleLugValue(szformuarName, maszSplitLuggageItem)  '跟据协议得出折算款
                    mdbNeedSplitPrice = dbChargeMoney
                Else
                    dbChargeMoney = FormatDbValue(rsTempProtocol!charge_money)  '为固定费用项
                    mdbNeedSplitPrice = dbChargeMoney
                End If


                '更新结算单拆帐票项表
                szSql = "INSERT fin_luggage_split_detail_lst ( fin_sheet_id , vehicle_id , charge_code , charge_name , split_out_money ) " _
                    & " VALUES ( " _
                    & TransFieldValueToString(mszSheetID) & "," _
                    & TransFieldValueToString(mszVehicleID) & "," _
                    & TransFieldValueToString(FormatDbValue(rsTempProtocol!charge_code)) & "," _
                    & TransFieldValueToString(FormatDbValue(rsTempProtocol!charge_name)) & "," _
                    & TransFieldValueToString(dbChargeMoney) & ")"
                oDb.Execute szSql
                '更新行包车辆结算明细表
                szSql = "UPDATE fin_luggage_sheet_lst SET " _
                    & " need_split_out= " & TransFieldValueToString(mdbNeedSplitPrice) _
                    & " Where fin_sheet_id = " & TransFieldValueToString(mszSheetID) _
                    & " AND vehicle_id = " & TransFieldValueToString(FormatDbValue(rsTemp!vehicle_id))
                oDb.Execute szSql
                '拆出额累加
                Dim rsTempNeedSplitOut2 As Recordset

                szSql = "SELECT SUM(need_split_out) need_split_out FROM fin_luggage_sheet_info WHERE fin_sheet_id=" & TransFieldValueToString(mszSheetID)
                Set rsTempNeedSplitOut2 = oDb.Execute(szSql)
                If FormatDbValue(rsTempNeedSplitOut2!need_split_out) <> 0 Then
                    mdbNeedSplitPrice = mdbNeedSplitPrice + FormatDbValue(rsTempNeedSplitOut2!need_split_out)
                End If
                szSql = "UPDATE fin_luggage_sheet_info SET " _
                    & " need_split_out = " & TransFieldValueToString(mdbNeedSplitPrice) _
                    & " WHERE fin_sheet_id = " & TransFieldValueToString(mszSheetID)
                oDb.Execute szSql
                rsTempProtocol.MoveNext
            Next j
            rsTemp.MoveNext
        Next
    End If
    Set rsTemp = Nothing

    '插入fin_sheet_to_luggage_lst表中
    For i = 1 To ArrayLength(aszTempSheetID)
        szSheetid = aszTempSheetID(i)
        szSql = "INSERT fin_sheet_to_luggage_lst (fin_sheet_id,luggage_sheet_id ) " _
            & " VALUES ( " _
            & TransFieldValueToString(mszSheetID) & "," _
            & TransFieldValueToString(szSheetid) & ")"
        oDb.Execute szSql
        szSql = "UPDATE luggage_sheet_info SET settlement_status =1 WHERE sheet_id = " & TransFieldValueToString(szSheetid)
        oDb.Execute szSql
    Next i
    oDb.CommitTrans
    Set oDb = Nothing

    Exit Sub
SplitCarrySheetsErr:
    oDb.RollbackTrans
    Call RaiseError(MyUnhandledError, "LugFinSplitSvr:SplitCarrySheets Method")
End Sub

