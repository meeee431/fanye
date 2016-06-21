VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PackageParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
' *  Source File Name  :  BaseInfo.cls                            *
' *  Project Name: FNBase.prj                                *
' *  Engineer: 陆勇庆                                               *
'*  Brief Description   : 基本定义及码表对象                     *
' *******************************************************************'

Option Base 0
Option Explicit


'本模块私有定义部分
'==========================================================================
'错误号
Public Enum EErrPackageParam
    ERR_BaseCodeNotExist = ERR_PackageParam + 21    '基本代码不存在
    ERR_SystemParam_InvalidParam = ERR_PackageParam + 22    '非法参数
End Enum
'权限号

Public Enum ERightPackageParam
    RIGHT_PackageParamBaseCodeEdit = ERR_PackageParam + cnMidRightBegin + cnMidRightStep         '30500开始
    RIGHT_PackageParamBaseCodeDelete = ERR_PackageParam + cnMidRightBegin + cnMidRightStep * 2
    RIGHT_PackageParamSet = ERR_PackageParam + cnMidRightBegin + cnMidRightStep * 3
End Enum

'以下固定枚举定义
'===================================================
'公用枚举部分:
'---------------------------------------------------------
Public Enum ECommStatus           '通用状态
    ENormal = 0
    EOther = 9
End Enum

Public Enum EDefineType     '常用定义类别
    EDT_SavePosition = 61            '存放位置
    EDT_LoadWorker = 62            '装卸工
    EDT_CalWeight = 63          '计重
    EDT_PickType = 64       '交付方式
    EDT_AreaType = 65          '地区方式
    EDT_PackageName = 66        '行包名称
    EDT_PackType = 67        '包装类型
    EDT_Operator = 68            '受理人
    
    EDT_Other1 = 91         '补充字典1
    EDT_Other2 = 92         '补充字典2
    EDT_Other3 = 93         '补充字典3
    EDT_Other4 = 94         '补充字典4
End Enum




'参数名称常量
Const cszSheetIDNumberLen = "PackageIDNumberLen"          '单据号数字部分长
Const cszTransitChargeRatio = "PackageTransitRatio"        '代收运费管理费
Const cszNormalKeepCharge = "NormalKeepCharge"              '每件每天货物保管费
Const cszKeepFeeDays = "KeepFeeDays"

Private moSystemParam As New SystemParam    '系统参数对象



'##ModelId=3FFD972E0349
Const cszModName = "STPackage.PackageParam"


'活动用户
Private m_oActiveUser As ActiveUser


'##ModelId=3FFD972E0353
Public Function ListBaseDefine(ByVal DefineType As Integer) As String()
    On Error GoTo GetBaseCodeErr
    Const cszProcName = "ListBaseDefine"
   '验证用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_PackageParam

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

'##ModelId=3FFD972E037B
Public Function AddBaseDefine(ByVal DefineType As EDefineType, ByVal pszBaseName As String, ByVal pszRemark As String) As Long
    On Error GoTo AddBaseCodeErr
    Const cszLocalFunction = "EditBaseCode"
    Const cszProcName = "AddBase"
   '验证用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_PackageParam
    AssertHaveRight m_oActiveUser, RIGHT_PackageParamBaseCodeEdit
    
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr
    '如果该类型的基本代码不存在,则出错
    Dim lMaxID As Long
    oDb.BeginTrans
    
    
'    szSql = " SELECT * FROM base_code WHERE type = " & TransFieldValueToString(DefineType)
    '新增该基本代码
    szSql = " INSERT base_code (" _
        & " type , " _
        & " name , " _
        & " memo ) VALUES ( " _
        & TransFieldValueToString(DefineType) & "," _
        & TransFieldValueToString(pszBaseName) & "," _
        & TransFieldValueToString(pszRemark) & ") "
    oDb.Execute szSql
    
    szSql = "SELECT * FROM base_code "
    Set rsTemp = oDb.Execute(szSql)
    lMaxID = rsTemp.RecordCount
    
'    WriteDataUpdate "base_code", cszInsert, "id", lMaxID, False, True, oDb, "AddBaseDefine"
    oDb.CommitTrans
    
    Set oDb = Nothing
    Set rsTemp = Nothing
    AddBaseDefine = lMaxID
    Exit Function
AddBaseCodeErr:
    oDb.RollbackTrans
    Set oDb = Nothing
    Set rsTemp = Nothing
    Call RaiseError(err.Number, "PackageParam:AddBaseDefine Method", err.Description)
End Function

'##ModelId=3FFD972E03A4
Public Sub DelBaseDefine(Optional plID As Long = -1, Optional ByVal DefineType As EDefineType = -1, Optional ByVal pszBaseName As String = "")
    On Error GoTo DeletEDefineTypeErr
    Const cszLocalFunction = "EditBaseCode"
    Const cszProcName = "DelBaseDefine"
   '验证用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_PackageParam
    AssertHaveRight m_oActiveUser, RIGHT_PackageParamBaseCodeDelete
    
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim lEffect As Long
    oDb.ConnectionString = GetConnectionStr
'    oDb.BeginTrans
    
    szSql = " DELETE base_code WHERE 1 =1 "
    If plID <> -1 Then
        szSql = szSql & " AND no=" & TransFieldValueToString(plID)
    End If
    If DefineType <> -1 Then
        szSql = szSql & " AND type=" & TransFieldValueToString(DefineType)
    End If
    If pszBaseName <> "" Then
        szSql = szSql & " AND  name=" & TransFieldValueToString(pszBaseName)
    End If
    oDb.Execute szSql, lEffect
   '如果删除的影响记录数为零则出错.
    If lEffect = 0 Then RaiseError ERR_BaseCodeNotExist
    
'    WriteDataUpdate "base_code", cszDelete, "id, type ,name ", CStr(plID) & "," & CStr(DefineType) & "," & CStr(pszBaseName), False, True, oDb, "DelBaseDefine"
'
'    oDb.CommitTrans
    Set oDb = Nothing
    Set rsTemp = Nothing
    Exit Sub
DeletEDefineTypeErr:
    oDb.RollbackTrans
    Set oDb = Nothing
    Set rsTemp = Nothing
    Call RaiseError(err.Number, "PackageParam:DelBaseDefine Method", err.Description)
End Sub

'##ModelId=3FFD972E03C2
Public Sub UpdBaseDefine(ByVal plID As Long, ByVal DefineType As EDefineType, ByVal pszBaseName As String, ByVal pszRemark As String)
    On Error GoTo EditBaseCodeErr
    Const cszLocalFunction = "EditBaseCode"
    Const cszProcName = "UpdBaseDefine"
   '验证用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_PackageParam
    AssertHaveRight m_oActiveUser, RIGHT_PackageParamBaseCodeEdit
    
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr
    szSql = " SELECT * FROM base_code WHERE no = " & TransFieldValueToString(plID)
    Set rsTemp = oDb.Execute(szSql)
    '如果该基本名称不存在
    If rsTemp.RecordCount = 0 Then RaiseError ERR_BaseCodeNotExist
    '修改该基本代码信息
'    oDb.BeginTrans
    szSql = " UPDATE base_code SET " _
        & " type = " & TransFieldValueToString(DefineType) & "," _
        & " name = " & TransFieldValueToString(pszBaseName) & "," _
        & " memo= " & TransFieldValueToString(pszRemark) _
        & " WHERE no = " & TransFieldValueToString(plID)
    oDb.Execute szSql
    
'    WriteDataUpdate "base_code", cszUpdate, "id", plID, False, True, oDb, "UpdBaseDefine"
'    oDb.CommitTrans
    
    Set oDb = Nothing
    Set rsTemp = Nothing
    Exit Sub
EditBaseCodeErr:
    oDb.RollbackTrans
    Set oDb = Nothing
    Set rsTemp = Nothing
    Call RaiseError(err.Number, "PackageParam:UpdBaseDefine Method", err.Description)
End Sub

'初始化活动对象和当前售票站点
'##ModelId=3E03A6A700B8
Public Sub Init(poAUser As ActiveUser)
    On Error GoTo InitErr
    Set m_oActiveUser = poAUser
    moSystemParam.Init m_oActiveUser
    
    Exit Sub
InitErr:
    Call RaiseError(err.Number, "LuggageSvr:Init Method", err.Description)
End Sub
Public Property Get SheetIDNumberLen() As Double
    Let SheetIDNumberLen = moSystemParam.GetParam(cszSheetIDNumberLen).szParamValue
End Property

Public Property Let SheetIDNumberLen(ByVal Value As Double)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_PackageParam
    AssertHaveRight m_oActiveUser, RIGHT_PackageParamSet
    
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszSheetIDNumberLen
    piTemp.szParamValue = Value
    moSystemParam.ModifyParam piTemp, True
End Property

Public Property Get TransitChargeRatio() As Double
    Let TransitChargeRatio = moSystemParam.GetParam(cszTransitChargeRatio).szParamValue
End Property

Public Property Let TransitChargeRatio(ByVal Value As Double)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_PackageParam
    AssertHaveRight m_oActiveUser, RIGHT_PackageParamSet
    
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszTransitChargeRatio
    piTemp.szParamValue = Value
    moSystemParam.ModifyParam piTemp, True
End Property

Public Property Get KeepFeeDays() As Integer
    Let KeepFeeDays = moSystemParam.GetParam(cszKeepFeeDays).szParamValue
End Property
Public Property Let KeepFeeDays(ByVal Value As Integer)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_PackageParam
    AssertHaveRight m_oActiveUser, RIGHT_PackageParamSet
    
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszKeepFeeDays
    piTemp.szParamValue = Value
    moSystemParam.ModifyParam piTemp, True
End Property
Public Property Get NormalKeepCharge() As Double
    Let NormalKeepCharge = moSystemParam.GetParam(cszNormalKeepCharge).szParamValue
End Property

Public Property Let NormalKeepCharge(ByVal Value As Double)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_PackageParam
    AssertHaveRight m_oActiveUser, RIGHT_PackageParamSet
    
    Dim piTemp As TParamInfo
    piTemp.szParamName = cszNormalKeepCharge
    piTemp.szParamValue = Value
    moSystemParam.ModifyParam piTemp, True
End Property

'##ModelId=3FFD972E0353
Public Function ListLoadChargeCode(Optional pszWeightID As String) As String()
    On Error GoTo GetBaseCodeErr
    Const cszProcName = "ListLoadChargeCode"
   '验证用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_PackageParam
'    AssertHaveRight m_oActiveUser, RIGHT_LuggageSvrNormalAccept

    Dim szSql As String
    Dim aszTemp() As String
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim oDb As New RTConnection
    
    oDb.ConnectionString = GetConnectionStr
    '如果参数缺省,则不加该条件,不需要模糊查询
    
    szSql = " SELECT * FROM weigth_charge_code WHERE 1=1  "
    If pszWeightID <> "" Then
        szSql = szSql & " AND weight_name= " & TransFieldValueToString(pszWeightID)
    End If
    
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 4)
    For i = 1 To rsTemp.RecordCount
        aszTemp(i, 1) = FormatDbValue(rsTemp!weight_id)
        aszTemp(i, 2) = FormatDbValue(rsTemp!weight_name)
        aszTemp(i, 3) = FormatDbValue(rsTemp!load_charge)
        aszTemp(i, 4) = FormatDbValue(rsTemp!Memo)
        rsTemp.MoveNext
    Next i
    ListLoadChargeCode = aszTemp
    Set oDb = Nothing
    Set rsTemp = Nothing
    Exit Function
GetBaseCodeErr:
    Set oDb = Nothing
    Set rsTemp = Nothing
    Call RaiseError(err.Number, "PackageParam:ListLoadChargeCode Method", err.Description)
End Function
'##ModelId=3FFD972E037B
Public Sub AddLoadChargeCode(pszWeightID As String, ByVal pszWeightName As String, ByVal pdbLoadCharge As Double, ByVal pszRemark As String)
    On Error GoTo AddBaseCodeErr
    Const cszLocalFunction = "EditBaseCode"
    Const cszProcName = "LoadChargeCode"
   '验证用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_PackageParam
    AssertHaveRight m_oActiveUser, RIGHT_PackageParamBaseCodeEdit
    
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr
    '如果该类型的基本代码不存在,则出错
    
    
'    szSql = " SELECT * FROM base_code WHERE type = " & TransFieldValueToString(DefineType)
    '新增该基本代码
    szSql = " INSERT weigth_charge_code (" _
        & " weight_id , weight_name ,load_charge, memo ) VALUES ( " _
        & TransFieldValueToString(pszWeightID) & "," _
        & TransFieldValueToString(pszWeightName) & "," _
        & TransFieldValueToString(pdbLoadCharge) & "," _
        & TransFieldValueToString(pszRemark) & ") "
    oDb.Execute szSql
    
    
    Set oDb = Nothing
    Set rsTemp = Nothing
    Exit Sub
AddBaseCodeErr:
    Set oDb = Nothing
    Set rsTemp = Nothing
    Call RaiseError(err.Number, "PackageParam:AddLoadChargeCode Method", err.Description)
End Sub

'##ModelId=3FFD972E03A4
Public Sub DelLoadChargeCode(Optional pszWeightID As String)
    On Error GoTo DeletEDefineTypeErr
    Const cszLocalFunction = "EditBaseCode"
    Const cszProcName = "DelLoadChargeCode"
   '验证用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_PackageParam
    AssertHaveRight m_oActiveUser, RIGHT_PackageParamBaseCodeDelete
    
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim lEffect As Long
    oDb.ConnectionString = GetConnectionStr
'    oDb.BeginTrans
    
    szSql = " DELETE weigth_charge_code WHERE 1 =1 "
    If pszWeightID <> "" Then
        szSql = szSql & " AND weight_id=" & TransFieldValueToString(pszWeightID)
    End If
    oDb.Execute szSql, lEffect
   '如果删除的影响记录数为零则出错.
    If lEffect = 0 Then RaiseError ERR_BaseCodeNotExist
    
'    WriteDataUpdate "base_code", cszDelete, "id, type ,name ", CStr(plID) & "," & CStr(DefineType) & "," & CStr(pszBaseName), False, True, oDb, "DelLoadChargeCode"
'
'    oDb.CommitTrans
    Set oDb = Nothing
    Set rsTemp = Nothing
    Exit Sub
DeletEDefineTypeErr:
    Set oDb = Nothing
    Set rsTemp = Nothing
    Call RaiseError(err.Number, "PackageParam:DelLoadChargeCode Method", err.Description)
End Sub

'##ModelId=3FFD972E03C2
Public Sub UpdLoadChargeCode(pszWeightID As String, ByVal pszWeightName As String, ByVal pdbLoadCharge As Double, ByVal pszRemark As String)
    On Error GoTo EditBaseCodeErr
    Const cszLocalFunction = "EditBaseCode"
    Const cszProcName = "UpdLoadChargeCode"
   '验证用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_PackageParam
    AssertHaveRight m_oActiveUser, RIGHT_PackageParamBaseCodeEdit
    
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr
    '修改该基本代码信息
'    oDb.BeginTrans
    szSql = " UPDATE weigth_charge_code SET " _
        & " weight_name = " & TransFieldValueToString(pszWeightName) & "," _
        & " load_charge = " & TransFieldValueToString(pdbLoadCharge) & "," _
        & " memo= " & TransFieldValueToString(pszRemark) _
        & " WHERE weight_id = " & TransFieldValueToString(pszWeightID)
    oDb.Execute szSql
    
'    WriteDataUpdate "base_code", cszUpdate, "id", plID, False, True, oDb, "UpdLoadChargeCode"
'    oDb.CommitTrans
    
    Set oDb = Nothing
    Set rsTemp = Nothing
    Exit Sub
EditBaseCodeErr:
    Set oDb = Nothing
    Set rsTemp = Nothing
    Call RaiseError(err.Number, "PackageParam:UpdLoadChargeCode Method", err.Description)
End Sub

