VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LugProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3DDD2055009B"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'行包结算协议
Option Base 0
Option Explicit
Const MyUnhandledError = 2134
Public Type TLugProtocol    '行包拆算协议结构
    ProtocolID As String    '协议号
    ProtocolName As String '协议名称
    Annotation As String    '备注
    IsDefault As Boolean    '是否是缺省协议
    
End Type
'**********************下面的类型已不用改用TProtocolChargeItemEX********************
Public Type TProtocolChargeItem     '协议拆算项目结构
    ProtocolID As String    '协议号
    ProtocolName As String  '协议名称
'    AcceptType As String    '托运方式
    ChargeID As String      '项目代码
    ChargeType As Integer   '项目类型 0-固定费用 1-公式计算
    ChargeName As String    '项目名称
    FormulaName As String   '公式名称
    FixCharge As Double     '固定费用时的费用值
    FormulaText As String   '公式内容
End Type
'***********************************************************************************
Public Type TProtocolChargeItemEX     '协议拆算项目结构
    ProtocolID As String    '协议号
    ProtocolName As String  '协议名称
    AcceptType As String    '托运方式
    ChargeID As String      '项目代码
    ChargeType As Integer   '项目类型 0-固定费用 1-公式计算
    ChargeName As String    '项目名称
    FormulaName As String   '公式名称
    FixCharge As Double     '固定费用时的费用值
    FormulaText As String   '公式内容
End Type


Public Type TVehicleProtocol    '车辆协议
    VehicleID As String     '车辆代码
    VehicleLicense As String    '车牌号
    AcceptType As String    '托运方式
    ProtocolID As String    '协议号
    ProtocolName As String  '协议名称
End Type
'
''出错
Public Enum EErrLuggageProtocol
                                                            ' 起始号为32000
        ERR_ProtocolNotExist = ERR_LuggageProtocol + 21  '32021 协议不存在
        ERR_ProtocolIDNotNull = ERR_LuggageProtocol + 22 '32022协议代码不能为空
        ERR_SplitNotAdd = ERR_LuggageProtocol + 23 '32023不能新增,拆算代码已存在
        ERR_ProtocolExistNotAdd = ERR_LuggageProtocol + 24  '32024不能新增 , 协议已存在
        ERR_SplitNotAddStatus = ERR_LuggageProtocol + 25 '32025拆算项代码对象不处于新增状态
        ERR_SplitIDNotNull = ERR_LuggageProtocol + 26   '32026拆算项代码不能为空
        ERR_VehicleProtocolExist = ERR_LuggageProtocol + 27 '32027车辆已设定了协议
        ERR_VehiclNoExistInfo = ERR_LuggageProtocol + 28 '32028车辆信息不存在
        ERR_ChargeItemExist = ERR_LuggageProtocol + 29 '32029行包费用名表中已有记录
        ERR_ChargeItemNoExist = ERR_LuggageProtocol + 30 '32030 没有行包费用表信息
        ERR_ProtocolNoDelete = ERR_LuggageProtocol + 31 '32031不能删除协议,因为协议已由被使用过
        ERR_ProtocolNotAddStatus = ERR_LuggageProtocol + 32 '32032协议对象不处于新增状态
End Enum
'权限
Public Enum ERightProtocol
    
                                                                                '32000
    RIGHT_LuggageProtocolAdd = ERR_LuggageProtocol + cnMidRightBegin + cnMidRightStep * 1 '32106 协议新增权限
    RIGHT_LuggageProtocolEdit = ERR_LuggageProtocol + cnMidRightBegin + cnMidRightStep * 2 '32112　　　协议编辑权限
    RIGHT_LuggageProtocolDel = ERR_LuggageProtocol + cnMidRightBegin + cnMidRightStep * 3 ' 32118协议删除权限
'    RIGHT_FinSplitSvr = ERR_LuggageFinSplitSvr + cnMidRightBegin + cnMidRightStep * 5  ' 27330 结算统计管理权限
End Enum
'set this to 0 to disable debug code in this class
#Const DebugMode = 1
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3E287B650203
    Private mlClassDebugID As Long
#End If

'协议代码
'##ModelId=3E23108001DD
Private mszProtocolID As String

'协议名称
'##ModelId=3E23108001E7
Private mszProtocolName As String

'公式内容
'##ModelId=3E23108001F1
Private mnDefault As Integer

'备注
'##ModelId=3E2352EA02CB
Private mszAnnotation As String
Private m_oActiveUser As ActiveUser
Private m_nObjectStatus As EObjectStatus '对象状态
Private mszAcceptType As String

'##ModelId=3E287B65032C
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3E287B65030D
'Private Sub Class_Initialize()
''    #If DebugMode Then
''        'get the next available class ID, and print out
''        'that the class was created successfully
'''        mlClassDebugID = GetNextClassDebugID()
''        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
''    #End If
'End Sub

'初始化活动对象
'##ModelId=3E23107100AF
Public Sub Init(poAUser As ActiveUser)
    On Error GoTo InitErr
    Set m_oActiveUser = poAUser
    Exit Sub
InitErr:
    Call RaiseError(MyUnhandledError, "LugProtocol:Init Method")
End Sub

'指定公式
'##ModelId=3E23107100B9
'刷新出协议信息
Public Sub Identify(FormulaID As String)
    On Error GoTo IdentifyErr
    AssertActiveUserValid m_oActiveUser, ERR_Protocol
    mszProtocolID = FormulaID
     
    RefreshInfo
    m_nObjectStatus = ST_NormalObj
    Exit Sub
IdentifyErr:
    Call RaiseError(MyUnhandledError, "LugProtocol:Identify Method")
End Sub

'刷新出协议信息
Private Sub RefreshInfo()
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT * FROM luggage_settle_protocol_info " _
        & " WHERE protocol_id = " & TransFieldValueToString(mszProtocolID)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then
        Exit Sub   '该协议不存在
    End If
    mszProtocolID = FormatDbValue(rsTemp!protocol_id)
    mszProtocolName = FormatDbValue(rsTemp!protocol_name)
    mszAnnotation = FormatDbValue(rsTemp!Annotation)
    mnDefault = FormatDbValue(rsTemp!default_mark)
  
    Set oDb = Nothing
End Sub
'新增公式
'##ModelId=3E23107100C3
Public Sub AddNew()
    On Error GoTo AddNewErr
    AssertActiveUserValid m_oActiveUser, ERR_Protocol
    AssertHaveRight m_oActiveUser, RIGHT_LuggageProtocolAdd
    mszAcceptType = ""
    m_nObjectStatus = ST_AddObj   '新增状态
    mszAnnotation = ""
    mnDefault = -1
    mszProtocolID = ""
    mszProtocolName = ""
    Exit Sub
AddNewErr:
    Call RaiseError(MyUnhandledError, "LugProtocol:AddNew Method")
End Sub

'更新公式
'##ModelId=3E23107100D7
Public Sub Update()
    On Error GoTo UpdateErr
    AssertActiveUserValid m_oActiveUser, ERR_Protocol
    
    If m_nObjectStatus = ST_AddObj Then
        AddObject
    ElseIf m_nObjectStatus = ST_EditObj Then
        AssertHaveRight m_oActiveUser, RIGHT_LuggageProtocolEdit
        UpdateToDB
    End If
    m_nObjectStatus = ST_NormalObj
    Exit Sub
UpdateErr:
    Call RaiseError(MyUnhandledError, "LugProtocol:Update Method")
End Sub

Private Sub AddObject()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    If mszProtocolID = "" Or mszProtocolID = Null Then
        RaiseError ERR_ProtocolIDNotNull
    End If
    If m_nObjectStatus <> ST_AddObj Then RaiseError ERR_ProtocolNotAddStatus
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
'    If mnDefault <> 0 Then
'        szSql = "SELECT * FROM luggage_settle_protocol_info WHERE default_mark = " & TransFieldValueToString(mnDefault)
'        Set rsTemp = oDb.Execute(szSql)
'        If rsTemp.RecordCount <> 0 Then
'            szSql = "UPDATE luggage_settle_protocol_info SET default_mark = 0 WHERE protocol_id = " & TransFieldValueToString(FormatDbValue(rsTemp!protocol_id))
'            oDb.Execute (szSql)
'        End If
'    End If
    szSql = "INSERT luggage_settle_protocol_info " _
        & " (protocol_id,protocol_name,annotation,default_mark) " _
        & " VALUES ( " _
        & TransFieldValueToString(mszProtocolID) & "," _
        & TransFieldValueToString(mszProtocolName) & "," _
        & TransFieldValueToString(mszAnnotation) & "," _
        & TransFieldValueToString(mnDefault) & ")"
    oDb.Execute (szSql)
    WriteOperateLog m_oActiveUser, RIGHT_LuggageProtocolAdd, "新增协议" & mszProtocolID & "[" & mszProtocolName & "]"
    Set oDb = Nothing
    Exit Sub
ErrHandle:
    RaiseError ERR_ProtocolExistNotAdd
End Sub
Private Sub UpdateToDB()
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim mszOldID As String  '原协议
    Dim mszOldName As String
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT * FROM luggage_settle_protocol_info WHERE protocol_id = " & TransFieldValueToString(mszProtocolID)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then RaiseError ERR_ProtocolNotExist
        mszOldID = FormatDbValue(rsTemp!protocol_id)
        mszOldName = FormatDbValue(rsTemp!protocol_name)
        '把原先的默认协议设为非默认协议
        szSql = "SELECT * FROM luggage_settle_protocol_info WHERE default_mark = " & TransFieldValueToString(mnDefault)
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount <> 0 Then
            szSql = "UPDATE luggage_settle_protocol_info SET default_mark = 2 WHERE protocol_id = " & TransFieldValueToString(FormatDbValue(rsTemp!protocol_id))
            oDb.Execute (szSql)
        End If

    szSql = "UPDATE luggage_settle_protocol_info " _
        & " SET protocol_name = " & TransFieldValueToString(mszProtocolName) & "," _
        & " annotation = " & TransFieldValueToString(mszAnnotation) & "," _
        & " default_mark = " & TransFieldValueToString(mnDefault) _
        & " WHERE protocol_id = " & TransFieldValueToString(mszProtocolID)
    oDb.Execute szSql
    WriteOperateLog m_oActiveUser, RIGHT_LuggageProtocolEdit, "将原协议" & mszOldID & "[" & mszOldName & "]更改为" & mszProtocolID & "[ " & mszProtocolName & "]"
    Set oDb = Nothing
    Exit Sub
ErrHandle:
'    RaiseError ERR_Protocol
End Sub
'删除公式
'##ModelId=3E23107100E1
Public Sub Delete()
    On Error GoTo DeleteErr
    AssertActiveUserValid m_oActiveUser, ERR_Protocol
    AssertHaveRight m_oActiveUser, RIGHT_LuggageProtocolDel
    
    If m_nObjectStatus <> ST_AddObj Then
        DeleteObject
    End If

    Exit Sub
DeleteErr:
    Call RaiseError(ERR_ProtocolNoDelete, "LugProtocol:Delete Method")
End Sub

Private Sub DeleteObject()
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim mszOldName As String   '原协议
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
'    szSql = "SELECT 1 FROM fin_luggage_chargeitem_lst WHERE protocol_id = " & TransFieldValueToString(mszProtocolID)
'    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount <> 0 Then RaiseError ERR_ProtocolNoDelete  '协议已被使用，不能删除
    szSql = "SELECT 1 FROM vehicle_lug_protocol_info WHERE protocol_id = " & TransFieldValueToString(mszProtocolID)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 0 Then RaiseError ERR_ProtocolNoDelete  '协议已被使用，不能删除
    
    szSql = "SELECT * FROM luggage_settle_protocol_info WHERE protocol_id = " & TransFieldValueToString(mszProtocolID)
    Set rsTemp = oDb.Execute(szSql)
    mszOldName = FormatDbValue(rsTemp!protocol_name)
    
    szSql = "DELETE luggage_settle_protocol_info " _
        & " WHERE protocol_id = " & TransFieldValueToString(mszProtocolID)
    oDb.Execute szSql
    WriteOperateLog m_oActiveUser, RIGHT_LuggageProtocolDel, "删除协议" & mszProtocolID & "[" & mszOldName & "]"
    '删除该协议指定的所有费用项
    szSql = "SELECT 1 FROM fin_luggage_chargeitem_lst WHERE protocol_id = " & TransFieldValueToString(mszProtocolID)
    Set rsTemp = oDb.Execute(szSql)

    If rsTemp.RecordCount > 0 Then
        szSql = "DELETE  fin_luggage_chargeitem_lst WHERE protocol_id =" & TransFieldValueToString(mszProtocolID)
        oDb.Execute szSql
    End If
    
    Set oDb = Nothing
    Exit Sub
End Sub
'查询结算协议
'##ModelId=3E2310D20163
Public Function GetProtocol(Optional AcceptType As String = "", Optional OnlyDefault As Boolean = False) As TLugProtocol()
    AssertActiveUserValid m_oActiveUser, ERR_Protocol
    On Error GoTo GetProtocolErr
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim i As Integer
    Dim szTemp() As TLugProtocol
    Dim szTempSql As String
    
    oDb.ConnectionString = GetConnectionStr(cszSplit)
  
    If OnlyDefault = True Then
      szSql = "SELECT * FROM luggage_settle_protocol_info WHERE default_mark!=" & TransFieldValueToString(LuggageProtocolNotDefault)
    Else
      szSql = "SELECT * FROM luggage_settle_protocol_info WHERE 1 = 1 "
    End If
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then
       Exit Function '协议不存在
    End If
    
    If AcceptType <> "" Then
        szTempSql = " SELECT * FROM luggage_settle_protocol_info WHERE protocol_id IN ( " _
            & " SELECT protocol_id FROM fin_luggage_chargeitem_lst " _
            & " WHERE accept_type = " & TransFieldValueToString(GetLuggageTypeInt(AcceptType)) _
            & " ) "
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount = 0 Then
            Exit Function '协议不存在
        End If
    Else
        szTempSql = ""
    End If
    ReDim szTemp(1 To rsTemp.RecordCount)
    For i = 1 To rsTemp.RecordCount
        szTemp(i).ProtocolID = FormatDbValue(rsTemp!protocol_id)
        szTemp(i).ProtocolName = FormatDbValue(rsTemp!protocol_name)
        szTemp(i).Annotation = FormatDbValue(rsTemp!Annotation)
        If rsTemp!default_mark <> LuggageProtocolNotDefault Then
         szTemp(i).IsDefault = True
        Else
         szTemp(i).IsDefault = False
        End If
    rsTemp.MoveNext
    Next i
    GetProtocol = szTemp
    Set rsTemp = Nothing
    Set oDb = Nothing
    Exit Function
GetProtocolErr:
    Call RaiseError(MyUnhandledError, "LugProtocol:GetProtocol Method")
End Function

'列出当前协议的拆算项目信息，返回TProtocolChargeItemex类型的数组
'##ModelId=3E2312ED0050
Public Function ListChargeItem() As TProtocolChargeItemEX()
    AssertActiveUserValid m_oActiveUser, ERR_Protocol
    On Error GoTo ListChargeItemErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim aTempChargeItem() As TProtocolChargeItemEX
    Dim rsTemp As Recordset
    Dim nCount As Integer
    Dim nChargeType  As Integer, j As Integer, i As Integer
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT f.protocol_id,p.protocol_name ,f.accept_type,f.charge_code,f.charge_type,f.charge_name,f.charge_money,f.formular_name,f.formula_comment " _
        & " FROM fin_luggage_chargeitem_lst f,luggage_settle_protocol_info p " _
        & " WHERE f.protocol_id = " & TransFieldValueToString(mszProtocolID) _
        & " AND f.protocol_id = p.protocol_id "
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 0 Then
        ReDim aTempChargeItem(1 To rsTemp.RecordCount)
        For i = 1 To rsTemp.RecordCount
            aTempChargeItem(i).ProtocolID = FormatDbValue(rsTemp!protocol_id)
            aTempChargeItem(i).AcceptType = FormatDbValue(rsTemp!accept_type)
            aTempChargeItem(i).ProtocolName = FormatDbValue(rsTemp!protocol_name)
            aTempChargeItem(i).ChargeID = FormatDbValue(rsTemp!charge_code)
            aTempChargeItem(i).ChargeType = FormatDbValue(rsTemp!charge_type)
            aTempChargeItem(i).ChargeName = FormatDbValue(rsTemp!charge_name)
            aTempChargeItem(i).FormulaName = FormatDbValue(rsTemp!formular_name)
            aTempChargeItem(i).FixCharge = FormatDbValue(rsTemp!charge_money)
            aTempChargeItem(i).FormulaText = FormatDbValue(rsTemp!formula_comment)
            rsTemp.MoveNext
        Next i
    End If
    ListChargeItem = aTempChargeItem
    Exit Function
ListChargeItemErr:
    Call RaiseError(MyUnhandledError, "LugProtocol:ListChargeItem Method")
End Function

'对当前的拆算协议指定拆算项目计算方法
'##ModelId=3E235CBC019A
Public Sub SetChargeItem(ChargeItem() As TProtocolChargeItemEX)
    AssertActiveUserValid m_oActiveUser, ERR_Protocol
    AssertHaveRight m_oActiveUser, RIGHT_LuggageProtocolAdd
    
    On Error GoTo SetChargeItemErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim nCount As Integer
    Dim aTempChargeItem() As TProtocolChargeItemEX
    Dim szMaxFormulaID As String
    aTempChargeItem = ChargeItem
    If ArrayLength(aTempChargeItem) = 0 Then RaiseError ERR_ChargeItemNoExist
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT * FROM fin_luggage_chargeitem_lst WHERE protocol_id=" & TransFieldValueToString(aTempChargeItem(1).ProtocolID)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        szSql = "DELETE fin_luggage_chargeitem_lst " _
            & " WHERE protocol_id = " & TransFieldValueToString(aTempChargeItem(1).ProtocolID)
        oDb.Execute szSql
    End If
'        & " AND charge_code = " & TransFieldValueToString(aTempChargeItem(i).ChargeID)
        nCount = ArrayLength(aTempChargeItem)
    For i = 1 To nCount
        '行包费用名表
        szSql = "INSERT fin_luggage_chargeitem_lst (protocol_id,accept_type,charge_code,charge_type,charge_name,formular_name,charge_money,formula_comment) " _
            & " VALUES  ( " _
            & TransFieldValueToString(aTempChargeItem(i).ProtocolID) & "," _
            & TransFieldValueToString(aTempChargeItem(i).AcceptType) & "," _
            & TransFieldValueToString(aTempChargeItem(i).ChargeID) & "," _
            & TransFieldValueToString(aTempChargeItem(i).ChargeType) & "," _
            & TransFieldValueToString(aTempChargeItem(i).ChargeName) & "," _
            & TransFieldValueToString(aTempChargeItem(i).FormulaName) & "," _
            & TransFieldValueToString(aTempChargeItem(i).FixCharge) & "," _
            & TransFieldValueToString(aTempChargeItem(i).FormulaText) & ")"
        oDb.Execute szSql
    Next
    Set oDb = Nothing
    Set rsTemp = Nothing
    Exit Sub
SetChargeItemErr:
    RaiseError ERR_ChargeItemExist  '行包费用名表中已有记录
End Sub

'查询指定车辆的协议
'##ModelId=3E232FF000B3
Public Function GetVehicleProtocol(Optional VehicleID As String = "", Optional AcceptType As Integer = -1, Optional ProtocolID As String, _
                                   Optional LikeLicense As String = "", Optional CompanyID As String = "", _
                                   Optional SplitCompanyID As String = "", Optional OwnerID As String = "") As TVehicleProtocol()
    On Error GoTo GetVehicleProtocolErr
    AssertActiveUserValid m_oActiveUser, ERR_Protocol
    Dim szSql As String
    Dim szFroSql As String
    Dim szWheSql As String
    Dim szFroSql1 As String
    Dim szWheSql1 As String
    Dim szFroSql2 As String
    Dim szWheSql2 As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    Dim aTempVehiclpro() As TVehicleProtocol
    
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    If VehicleID <> "" Then   ' 车辆代码
        szFroSql = ""
        szWheSql = " AND v.vehicle_id = " & TransFieldValueToString(VehicleID)
    Else
        szFroSql = ""
        szWheSql = ""
    End If
    
    If LikeLicense <> "" Then    '车牌代码
        szFroSql = szFroSql + ""
        szWheSql = szWheSql + " AND v.license_tag_no LIKE '%" & LikeLicense & "%'"
    Else
        szFroSql = szFroSql + ""
        szWheSql = szWheSql + ""
    End If
    
    If AcceptType <> -1 Then  '托运方式
        szFroSql = szFroSql + ""
        szWheSql = szWheSql + " AND v.accept_type = " & TransFieldValueToString(AcceptType)
    Else
        szFroSql = szFroSql + ""
       szWheSql = szWheSql + ""
    End If
    
    If ProtocolID <> "" Then
        szFroSql = szFroSql + ""
        szWheSql = szWheSql + " AND v.protocol_id = " & TransFieldValueToString(ProtocolID)
    Else
        szFroSql = szFroSql + ""
        szWheSql = szWheSql + ""
    End If
    
    
    If CompanyID <> "" Then    '公司代码
        szFroSql2 = " ," & cszTableCompany & " c "
        szWheSql2 = " AND c.transport_company_id = vi.transport_company_id AND c.transport_company_id = " & TransFieldValueToString(CompanyID)
    Else
        szFroSql2 = ""
        szWheSql2 = ""
     End If
    
    If SplitCompanyID <> "" Then    '折分公司代码
        szFroSql2 = szFroSql2 + " ," & cszTableCompany & " c1"
        szWheSql2 = szWheSql2 + " AND c1.transport_company_id = vi.transport_company_id AND vi.split_company_id = " & TransFieldValueToString(SplitCompanyID)
        
    Else
        szFroSql2 = szFroSql2 + ""
        szWheSql2 = szWheSql2 + ""
    End If
    
    If OwnerID <> "" Then    '车主代码
        szFroSql2 = szFroSql2 + " ," & cszTableOwnerInfo & " o"
        szWheSql2 = szWheSql2 + " AND o.owner_id = vi.owner_id AND o.owner_id = " & TransFieldValueToString(OwnerID)
    Else
        szFroSql2 = szFroSql2 + ""
        szWheSql2 = szWheSql2 + ""
    End If
    If szFroSql2 <> "" Then
        szFroSql2 = " ," & cszTableVehicle & " vi" & szFroSql2
        szWheSql2 = szWheSql2 + " AND v.vehicle_id=vi.vehicle_id "
    End If
    
    '处理临时新增的车辆
        szSql = "INSERT INTO vehicle_lug_protocol_info" _
            & " SELECT t.vehicle_id,0,'',t.license_tag_no FROM " & cszTableVehicle & " t WHERE t.vehicle_id NOT IN " _
            & " (SELECT vehicle_id FROM vehicle_lug_protocol_info WHERE accept_type=0) "
        oDb.Execute szSql
        szSql = "INSERT INTO vehicle_lug_protocol_info" _
            & " SELECT t.vehicle_id,1,'',t.license_tag_no FROM " & cszTableVehicle & " t WHERE t.vehicle_id NOT IN " _
            & " (SELECT vehicle_id FROM vehicle_lug_protocol_info WHERE accept_type=1) "
        oDb.Execute szSql

    
    szSql = "SELECT  v.vehicle_id,v.license_tag_no,v.accept_type,v.protocol_id ,p.protocol_name " _
        & " FROM vehicle_lug_protocol_info v,luggage_settle_protocol_info p " & szFroSql & szFroSql2 _
        & " WHERE v.protocol_id = p.protocol_id " & szWheSql & szWheSql2
    Set rsTemp = oDb.Execute(szSql)
    '显示没有协议的记录
    Dim rsTemp1 As Recordset
    szSql = "SELECT v.vehicle_id,v.license_tag_no,v.accept_type,v.protocol_id " _
        & " FROM vehicle_lug_protocol_info v " & szFroSql & szFroSql2 _
        & " WHERE v.protocol_id='' " & szWheSql & szWheSql2
    Set rsTemp1 = oDb.Execute(szSql)
    Dim nCount As Integer, nCount1 As Integer, i As Integer, j As Integer
    nCount = rsTemp.RecordCount
    nCount1 = rsTemp1.RecordCount
    If nCount = 0 And nCount1 = 0 Then Exit Function
    ReDim aTempVehiclpro(1 To nCount1 + nCount)
    If nCount1 <> 0 Then
        
        For j = 1 To nCount1
            aTempVehiclpro(j).AcceptType = GetLuggageTypeString(FormatDbValue(rsTemp1!accept_type))
            aTempVehiclpro(j).ProtocolID = ""
            aTempVehiclpro(j).ProtocolName = ""
            aTempVehiclpro(j).VehicleID = FormatDbValue(rsTemp1!vehicle_id)
            aTempVehiclpro(j).VehicleLicense = FormatDbValue(rsTemp1!license_tag_no)
            rsTemp1.MoveNext
        Next j
    End If
    
        If nCount = 0 And nCount1 = 0 Then
            Exit Function   '没有指定车辆的信息
        End If
        
        For i = nCount1 + 1 To nCount + nCount1
            aTempVehiclpro(i).AcceptType = GetLuggageTypeString(FormatDbValue(rsTemp!accept_type))
            aTempVehiclpro(i).ProtocolID = FormatDbValue(rsTemp!protocol_id)
            aTempVehiclpro(i).ProtocolName = FormatDbValue(rsTemp!protocol_name)
            aTempVehiclpro(i).VehicleID = FormatDbValue(rsTemp!vehicle_id)
            aTempVehiclpro(i).VehicleLicense = FormatDbValue(rsTemp!license_tag_no)
            rsTemp.MoveNext
        Next i
        GetVehicleProtocol = aTempVehiclpro
    Set oDb = Nothing
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    Exit Function
GetVehicleProtocolErr:
    Call RaiseError(MyUnhandledError, "LugProtocol:GetVehicleProtocol Method")
 End Function

'设置车辆对应的拆算协议
'##ModelId=3E245ED001B7
Public Sub SetVehicleProtocol(VehicleProtocol() As TVehicleProtocol)
    AssertActiveUserValid m_oActiveUser, ERR_Protocol
    AssertHaveRight m_oActiveUser, RIGHT_LuggageProtocolAdd
    On Error GoTo SetVehicleProtocolErr
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim i As Integer
    Dim nLen As Integer
    Dim aTempVehicle() As TVehicleProtocol
    aTempVehicle = VehicleProtocol
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    
   
    '处理临时新增的车辆
        szSql = "INSERT INTO vehicle_lug_protocol_info" _
            & " SELECT t.vehicle_id,0,'',t.license_tag_no FROM " & cszTableVehicle & " t WHERE t.vehicle_id NOT IN " _
            & " (SELECT vehicle_id FROM vehicle_lug_protocol_info WHERE accept_type=0) "
        oDb.Execute szSql
        szSql = "INSERT INTO vehicle_lug_protocol_info" _
            & " SELECT t.vehicle_id,1,'',t.license_tag_no FROM " & cszTableVehicle & " t WHERE t.vehicle_id NOT IN " _
            & " (SELECT vehicle_id FROM vehicle_lug_protocol_info WHERE accept_type=1) "
        oDb.Execute szSql
    
    nLen = ArrayLength(aTempVehicle)
    For i = 1 To nLen
'        szSql = "DELETE vehicle_lug_protocol_info WHERE accept_type =" & TransFieldValueToString(VehicleProtocol(i).AcceptType) _
'            & " AND vehicle_id = " & TransFieldValueToString(VehicleProtocol(i).VehicleID) _
'            & " AND accept_type = " & TransFieldValueToString(VehicleProtocol(i).AcceptType)
'        oDb.Execute (szSql)
    
    
       szSql = "UPDATE vehicle_lug_protocol_info SET protocol_id=" & TransFieldValueToString(VehicleProtocol(i).ProtocolID) _
          & " ,license_tag_no=" & TransFieldValueToString(VehicleProtocol(i).VehicleLicense) _
          & " WHERE vehicle_id=" & TransFieldValueToString(VehicleProtocol(i).VehicleID) _
          & " AND accept_type=" & TransFieldValueToString(VehicleProtocol(i).AcceptType)
'        szSql = " INSERT vehicle_lug_protocol_info (vehicle_id,accept_type,protocol_id,license_tag_no) " _
'            & " VALUES ( " _
'            & TransFieldValueToString(aTempVehicle(i).VehicleID) & "," _
'            & TransFieldValueToString(aTempVehicle(i).AcceptType) & "," _
'            & TransFieldValueToString(aTempVehicle(i).ProtocolID) & "," _
'            & TransFieldValueToString(aTempVehicle(i).VehicleLicense) & ")"
        oDb.Execute szSql
    Next i

    Exit Sub
SetVehicleProtocolErr:
    RaiseError ERR_VehicleProtocolExist
End Sub

'删除车辆对应的协议
'##ModelId=3E246349020D
Public Sub DelVehicleProtocol(VehicleID() As String, Optional AcceptType As ELuggageAcceptType = -1)
    AssertActiveUserValid m_oActiveUser, ERR_Protocol
    AssertHaveRight m_oActiveUser, RIGHT_LuggageProtocolDel
    
    On Error GoTo DelVehicleProtocolErr
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim rsTemp1 As Recordset
    Dim szSql As String
    Dim szTempProtocolID As String
    Dim szTemp As String
    Dim szTempWheSql As String
    Dim aszVehicleID() As String, i As Integer
    Dim nCount As Integer
        oDb.ConnectionString = GetConnectionStr(cszSplit)
    nCount = ArrayLength(VehicleID)
    If nCount = 0 Then Exit Sub
        For i = 1 To nCount
            szSql = "UPDATE vehicle_lug_protocol_info SET protocol_id='' WHERE  vehicle_id =" & TransFieldValueToString(VehicleID(i, 1)) _
                & " AND accept_type=" & TransFieldValueToString(VehicleID(i, 2))
            oDb.Execute szSql
        Next i
        Set oDb = Nothing
    
    Exit Sub
DelVehicleProtocolErr:
    Call RaiseError(MyUnhandledError, "LugProtocol:DelVehicleProtocol Method")
End Sub
'默认协议设置 2为不默认,0为快件默认,1 为随行默认
Public Sub SetProtocol(ProtocolID() As String)
    AssertActiveUserValid m_oActiveUser, ERR_Protocol
    AssertHaveRight m_oActiveUser, RIGHT_LuggageProtocolAdd
    On Error GoTo err
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim i As Integer
    Dim nCount As Integer
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    nCount = ArrayLength(ProtocolID)
    oDb.BeginTrans
    If nCount <> 0 Then
        For i = 1 To nCount
            If ProtocolID(i, 1) <> "" Then
                szSql = "UPDATE luggage_settle_protocol_info SET default_mark = " & TransFieldValueToString(ProtocolID(i, 1)) _
                    & " WHERE protocol_id=" & TransFieldValueToString(ProtocolID(i, 2))
            Else
                szSql = "UPDATE luggage_settle_protocol_info SET default_mark = 2" _
                    & " WHERE protocol_id=" & TransFieldValueToString(ProtocolID(i, 2))
            End If
            oDb.Execute szSql
        Next i
    End If
    '将所有协议为空的车辆,设置为默认协议
    
    
    oDb.CommitTrans
    Exit Sub
err:
    oDb.RollbackTrans
    Call RaiseError(MyUnhandledError, "LugProtocol:DelVehicleProtocol Method")
End Sub

'查询所有协议
Public Function GetAllProtocol(Optional ProtocolID As String = "") As ADODB.Recordset

    On Error GoTo err
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT * FROM luggage_settle_protocol_info "
    
    If ProtocolID <> "" Then
        szSql = szSql + " WHERE protocol_id= " & TransFieldValueToString(ProtocolID)
    End If

    Set rsTemp = oDb.Execute(szSql)
    Set GetAllProtocol = rsTemp
    
    Set rsTemp = Nothing
    Exit Function
err:
    Set rsTemp = Nothing
    Call RaiseError(MyUnhandledError, "LugProtocol:DelVehicleProtocol Method")
End Function

'将所有协议为空的车辆,设置为默认协议
Public Sub SetAllNoProtocolVehicle(VehicleID() As String)
    AssertActiveUserValid m_oActiveUser, ERR_Protocol
    AssertHaveRight m_oActiveUser, RIGHT_LuggageProtocolAdd
    On Error GoTo err
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim nCount As Integer
    Dim i As Integer
    Dim mProtocol1 As String
    Dim mProtocol2 As String
    nCount = ArrayLength(VehicleID)
    If nCount = 0 Then Exit Sub
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    '查询默认协议
    Dim rsTempProtocol As Recordset
    szSql = "SELECT * FROM luggage_settle_protocol_info WHERE default_mark=0"
    Set rsTempProtocol = oDb.Execute(szSql)
    If rsTempProtocol.RecordCount > 0 Then
        mProtocol1 = rsTempProtocol!protocol_id
    End If
    szSql = "SELECT * FROM luggage_settle_protocol_info WHERE default_mark=1"
    Set rsTempProtocol = oDb.Execute(szSql)
    If rsTempProtocol.RecordCount > 0 Then
        mProtocol2 = rsTempProtocol!protocol_id
    End If
    
    oDb.BeginTrans
    For i = 1 To nCount
        '处理临时新增车辆
        szSql = "SELECT t.vehicle_id,0,'',t.license_tag_no FROM " & cszTableVehicle & " t WHERE t.vehicle_id NOT IN (SELECT vehicle_id FROM vehicle_lug_protocol_info)"
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount > 0 Then '新增协议车辆
            szSql = "INSERT INTO vehicle_lug_protocol_info" _
                & " SELECT t.vehicle_id,0,'',t.license_tag_no FROM " & cszTableVehicle & " t WHERE t.vehicle_id NOT IN " _
                & " (SELECT vehicle_id FROM vehicle_lug_protocol_info WHERE accept_type=0) "
            oDb.Execute szSql
            szSql = "INSERT INTO vehicle_lug_protocol_info" _
                & " SELECT t.vehicle_id,1,'',t.license_tag_no FROM " & cszTableVehicle & " t WHERE t.vehicle_id NOT IN " _
                & " (SELECT vehicle_id FROM vehicle_lug_protocol_info WHERE accept_type=1) "
            oDb.Execute szSql
        End If
        '设置默认协议
        If CInt(VehicleID(i, 2)) = 0 Then
            szSql = "UPDATE vehicle_lug_protocol_info SET protocol_id=" & TransFieldValueToString(mProtocol1) _
                & " WHERE vehicle_id=" & TransFieldValueToString(VehicleID(i, 1))
        ElseIf CInt(VehicleID(i, 2)) = 1 Then
            szSql = "UPDATE vehicle_lug_protocol_info SET protocol_id=" & TransFieldValueToString(mProtocol2) _
                & " WHERE vehicle_id=" & TransFieldValueToString(VehicleID(i, 1))
        End If
        oDb.Execute szSql
    Next i
    
    oDb.CommitTrans
    Exit Sub
err:
    oDb.RollbackTrans
    Call RaiseError(MyUnhandledError, "LugProtocol:DelVehicleProtocol Method")
End Sub

#If DebugMode Then
    '##ModelId=3E287B650242
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

'##ModelId=3E288D83038A
Public Property Get Annotation() As String
   Let Annotation = mszAnnotation
   If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E288D8302CE
Public Property Let Annotation(ByVal Value As String)
    Let mszAnnotation = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E288D830261
Public Property Get Default() As Integer
   Let Default = mnDefault
   If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E288D8301C5
Public Property Let Default(ByVal Value As Integer)
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    
    Let mnDefault = Value
    If mnDefault = 1 And mszAcceptType <> "" Then
        szSql = "SELECT * FROM luggage_settle_protocol_info WHERE accept_type = " & TransFieldValueToString(GetLuggageTypeInt(mszAcceptType))
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount <> 0 Then
            szSql = "UPDATE luggage_settle_protocol_info SET default_mark = 0  " _
                & " WHERE protocol_id = " & TransFieldValueToString(FormatDbValue(rsTemp!protocol_id)) _
                & " AND accept_type = " & TransFieldValueToString(FormatDbValue(rsTemp!accept_type))
            oDb.Execute szSql
        End If
    End If
    
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property


'##ModelId=3E288D830157
Public Property Get ProtocolName() As String
   Let ProtocolName = mszProtocolName
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E288D8300BB
Public Property Let ProtocolName(ByVal Value As String)
    Let mszProtocolName = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E288D83004E
Public Property Get ProtocolID() As String
   Let ProtocolID = mszProtocolID
   If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property
'##ModelId=3E288D8203A9
Public Property Let ProtocolID(ByVal Value As String)
    Let mszProtocolID = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property
Public Property Let AcceptType(ByVal Value As String)
    Let mszAcceptType = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property
Public Property Get AcceptType() As String
    Let AcceptType = mszAcceptType
End Property

