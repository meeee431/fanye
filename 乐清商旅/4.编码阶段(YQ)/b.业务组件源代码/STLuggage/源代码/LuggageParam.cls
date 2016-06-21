VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LuggageParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'系统参数
Option Base 0
Option Explicit

Private m_oActiveUser As ActiveUser


'参数名称常量
Const cszLuggageIDNumberLen = "LuggageIDNumberLen"          '行包单号数字部分长
Const cszCarrySheetIDNumberLen = "CheckLuggageIDNumberLen"  '行包签发单号数字部分长
Const cszLuggageIDPrefixLen = "LuggageIDPrefixLen"          '行包单号前缀部分长
Const cszLuggageReturnRatio1 = "LuggageReturnRatio1"        '行包受理单退运费率1
Const cszEnabledPrintReturnSheet = "EnabledPrintLuggageReturnSheet"        '是否打印退运凭证
Const cszOurCompany = "OurCompany" '本公司
Const cszIsOneFormulaEachStation = "IsOneFormulaEachStation" '是否每站点一公式,0为所有站点相同公司,1为每站点一公式


Private moSystemParam As New SystemParam    '系统参数对象

'本模块私有定义部分
'==========================================================================
'错误号
Private Enum ELuggageParam
    ERR_SystemParamWithoutPrice = ERR_LuggageParam + 21   '没有行包种类信息 '30821开始
End Enum

'权限号
Private Enum ERightLuggageParam
    RIGHT_LuggageParam = ERR_LuggageParam + cnMidRightBegin             '31300开始
    RIGHT_LuggageParamSet = RIGHT_LuggageParam + cnMidRightStep * 1    '参数选项设置
End Enum
Public Sub Init(poAUser As ActiveUser)
    Set m_oActiveUser = poAUser
    moSystemParam.Init poAUser
    
End Sub
'得到包装形式
'暂时定义
Public Function GetPackageType() As String()
    Dim szaPackageType(1 To 2) As String
    szaPackageType(1) = "固件"
    szaPackageType(2) = "轻泡沬"
    GetPackageType = szaPackageType
End Function

'得到行包种类
Public Function GetLuggageKinds() As ADODB.Recordset
    On Error GoTo GetLuggageKindsErr
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    
    AssertActiveUserValid m_oActiveUser, ERR_SystemParam
    AssertHaveRight m_oActiveUser, RIGHT_LuggageParamSet
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    szSql = "SELECT * FROM luggage_kinds_code ORDER BY kinds_code "
    Set rsTemp = oDb.Execute(szSql)
    Set GetLuggageKinds = rsTemp
    'your code goes here...

    Exit Function
GetLuggageKindsErr:
    Call RaiseError(err.Number, "LuggageSvr:GetLuggageKinds Method", err.Description)
End Function

'得到运费项信息
Public Function GetPriceItemInfo(mAcceptType As Integer) As TLuggagePriceItem()
    AssertActiveUserValid m_oActiveUser, ERR_LuggageParam
    On Error GoTo GetPriceItemInfoErr
    Dim oCaculator As Object
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim atsTemp() As TLuggagePriceItem
    Dim i As Integer
    Dim price(1 To 5) As Double
    Dim szItemFormula As String
'    AssertHaveRight m_oActiveUser, RIGHT_SystemParam
    
'    If oCaculator Is Nothing Then
'        Set oCaculator = CreateObject("STLugPriLib.PriceItemFunLib")
'        oCaculator.Init m_oActiveUser
'    End If
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT  * FROM luggage_price_item_info WHERE use_mark = 0 AND accept_type =" & TransFieldValueToString(mAcceptType)
    Set rsTemp = oDb.Execute(szSql)
    
    If rsTemp.RecordCount = 0 Then ShowError ERR_SystemParamWithoutPrice
    ReDim atsTemp(1 To rsTemp.RecordCount)
    
    For i = 1 To rsTemp.RecordCount
'        szItemFormula = FormatDbValue(rsTemp!item_formula)
        atsTemp(i).PriceID = FormatDbValue(rsTemp!charge_item)
        atsTemp(i).PriceName = FormatDbValue(rsTemp!chinese_name)
        
        
        
        rsTemp.MoveNext
        
        'TLuggageItemInfo类型无luggage_price_item_info表中的字段
    Next
    
    GetPriceItemInfo = atsTemp
    Exit Function
GetPriceItemInfoErr:
    Call RaiseError(err.Number, "LuggageSvr:GetPriceItemInfo Method", err.Description)
End Function



Public Property Get LuggageReturnRatio1() As Double
    Let LuggageReturnRatio1 = moSystemParam.GetParam(cszLuggageReturnRatio1).szParamValue
End Property

Public Property Let LuggageReturnRatio1(ByVal Value As Double)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageParam
    AssertHaveRight m_oActiveUser, RIGHT_LuggageParamSet
    
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszLuggageReturnRatio1
    piTemp.szParamValue = Value
    moSystemParam.ModifyParam piTemp, True
End Property

Public Property Get LuggageIDNumberLen() As Double
    Let LuggageIDNumberLen = moSystemParam.GetParam(cszLuggageIDNumberLen).szParamValue
End Property

Public Property Let LuggageIDNumberLen(ByVal Value As Double)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageParam
    AssertHaveRight m_oActiveUser, RIGHT_LuggageParamSet
    
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszLuggageIDNumberLen
    piTemp.szParamValue = Value
    moSystemParam.ModifyParam piTemp, True
End Property
Public Property Get LuggageIDPrefixLen() As Double
    Let LuggageIDPrefixLen = moSystemParam.GetParam(cszLuggageIDPrefixLen).szParamValue
End Property

Public Property Let LuggageIDPrefixLen(ByVal Value As Double)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageParam
    AssertHaveRight m_oActiveUser, RIGHT_LuggageParamSet
    
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszLuggageIDPrefixLen
    piTemp.szParamValue = Value
    moSystemParam.ModifyParam piTemp, True
End Property

Public Property Get CarrySheetIDNumberLen() As Double
    Let CarrySheetIDNumberLen = moSystemParam.GetParam(cszCarrySheetIDNumberLen).szParamValue
End Property

Public Property Let CarrySheetIDNumberLen(ByVal Value As Double)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageParam
    AssertHaveRight m_oActiveUser, RIGHT_LuggageParamSet
    
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszCarrySheetIDNumberLen
    piTemp.szParamValue = Value
    moSystemParam.ModifyParam piTemp, True
End Property

Public Function GetPriceItemRS(Optional UseMark As Integer = -1, Optional AcceptType As Integer = -1) As ADODB.Recordset
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim i As Integer
    AssertActiveUserValid m_oActiveUser, ERR_SystemParam
    oDb.ConnectionString = GetConnectionStr
    
    szSql = "SELECT charge_item,chinese_name ,accept_type ,use_mark  FROM luggage_price_item_info WHERE 1=1 "
    If UseMark <> -1 Then szSql = szSql & " AND use_mark = " & UseMark
    If AcceptType <> -1 Then szSql = szSql & " AND accept_type = " & TransFieldValueToString(AcceptType)
    Set rsTemp = oDb.Execute(szSql)
    Set GetPriceItemRS = rsTemp
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

Public Function GetSellStationID() As String()
    Dim aszTemp() As String
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim i As Integer
    
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT * FROM sell_station_info "
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then
        '出错
    End If
    ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 2)
    For i = 1 To rsTemp.RecordCount
        aszTemp(i, 1) = FormatDbValue(rsTemp!sell_station_id)
        aszTemp(i, 2) = FormatDbValue(rsTemp!sell_station_name)
        rsTemp.MoveNext
    Next
    
    GetSellStationID = aszTemp
    
End Function
'是否打印退运凭证
Public Property Get EnabledPrintReturnSheet() As Boolean
    Let EnabledPrintReturnSheet = IIf(moSystemParam.GetParam(cszEnabledPrintReturnSheet).szParamValue = "1", True, False)
End Property

Public Property Let EnabledPrintReturnSheet(ByVal Value As Boolean)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageParam
    AssertHaveRight m_oActiveUser, RIGHT_LuggageParamSet
    
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszEnabledPrintReturnSheet
    If Value = True Then
        piTemp.szParamValue = "1"
    Else
        piTemp.szParamValue = "0"
    End If
    moSystemParam.ModifyParam piTemp, True
End Property
Public Function GetPriceItem(ChargeItem As String, AcceptType As Integer) As ADODB.Recordset
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT * FROM luggage_price_item_info WHERE charge_item = " & TransFieldValueToString(ChargeItem) _
          & " AND accept_type =" & TransFieldValueToString(AcceptType)
    Set rsTemp = oDb.Execute(szSql)
    Set GetPriceItem = rsTemp
    Set oDb = Nothing
End Function


'设置托运收费项


Public Sub SetPriceItem(PriceItem As TLuggagePriceItemFormulaEx)
    Dim szSql As String
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "UPDATE luggage_price_item_info" _
        & " SET  " _
        & " chinese_name = " & TransFieldValueToString(PriceItem.PriceItemName) & "," _
        & " use_mark = " & TransFieldValueToString(PriceItem.UsedMark) & "," _
        & " item_formula = " & TransFieldValueToString(PriceItem.Formula) & "," _
        & " parameter_1 = " & Val(PriceItem.szParam1) & "," _
        & " parameter_2 = " & Val(PriceItem.szParam2) & "," _
        & " parameter_3 = " & Val(PriceItem.szParam3) & "," _
        & " parameter_4 = " & Val(PriceItem.szParam4) & "," _
        & " parameter_5 = " & Val(PriceItem.szParam5) & "," _
        & " parameter_6 = " & Val(PriceItem.szParam6) & "," _
        & " parameter_7 = " & Val(PriceItem.szParam7) & "," _
        & " parameter_8 = " & Val(PriceItem.szParam8) & "," _
        & " parameter_9 = " & Val(PriceItem.szParam9) & "," _
        & " parameter_10 = " & Val(PriceItem.szParam10) _
        & " WHERE charge_item = " & TransFieldValueToString(PriceItem.PriceItem) _
        & " AND  accept_type = " & TransFieldValueToString(PriceItem.AcceptType)
    oDb.Execute szSql
End Sub

Public Function ListBaseDefine(ByVal DefineType As Integer) As String()
    On Error GoTo GetBaseCodeErr
    Const cszProcName = "ListBaseDefine"

    Dim szSql As String
    Dim aszTemp() As String
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim oDb As New RTConnection
    
    oDb.ConnectionString = GetConnectionStr
    '如果参数缺省,则不加该条件,不需要模糊查询
    szSql = " SELECT * FROM base_code WHERE 1=1  "
    If DefineType <> -1 Then
        szSql = szSql & " AND type = " & TransFieldValueToString(DefineType)
    End If
    
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 4)
    For i = 1 To rsTemp.RecordCount
        aszTemp(i, 1) = FormatDbValue(rsTemp!no)
        aszTemp(i, 2) = FormatDbValue(rsTemp!Type)
        aszTemp(i, 3) = FormatDbValue(rsTemp!Name)
        aszTemp(i, 4) = FormatDbValue(rsTemp!Memo)
        rsTemp.MoveNext
    Next i
    ListBaseDefine = aszTemp
    Set oDb = Nothing
    Set rsTemp = Nothing
    Exit Function
GetBaseCodeErr:
    Set oDb = Nothing
    Set rsTemp = Nothing
    Call RaiseError(err.Number, "PackageParam:ListBaseDefine Method", err.Description)
End Function





Public Property Get OurCompany() As String
    Let OurCompany = moSystemParam.GetParam(cszOurCompany).szParamValue
End Property

Public Property Let OurCompany(ByVal Value As String)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageParam
    AssertHaveRight m_oActiveUser, RIGHT_LuggageParamSet
    
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszOurCompany
    piTemp.szParamValue = Value
    moSystemParam.ModifyParam piTemp, True
    
End Property


'是否每站点一公式,0为所有站点相同公司1为每站点一公式)
Public Property Get IsOneFormulaEachStation() As Boolean
    Let IsOneFormulaEachStation = IIf(moSystemParam.GetParam(cszIsOneFormulaEachStation).szParamValue = "1", True, False)
End Property

Public Property Let IsOneFormulaEachStation(ByVal Value As Boolean)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageParam
    AssertHaveRight m_oActiveUser, RIGHT_LuggageParamSet
    
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszIsOneFormulaEachStation
    If Value = True Then
        piTemp.szParamValue = "1"
    Else
        piTemp.szParamValue = "0"
    End If
    moSystemParam.ModifyParam piTemp, True
    
End Property

