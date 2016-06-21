VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Protocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E80E573029F"
'此接口为设置协议和费用名表使用(详情请见<详细设计文档>)
Option Base 0
Option Explicit

'set this to 0 to disable debug code in this class
#Const DebugMode = 1
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3E8219BD0222
    Private mlClassDebugID As Long
#End If

'协议号
'##ModelId=3E80E5AE01A5
Private m_szProtocolID As String

'协议名称
'##ModelId=3E80E5D002BF
Private m_szProtocolName As String

'注释
'##ModelId=3E80E5EB02BF
Private m_szAnnotation As String

'是否默认协议
'##ModelId=3E80E5F200CB
Private m_EDefaultMark As EDefault

'折算项代码
'##ModelId=3E80E5F80399
Private m_szSplitItemID As String

'公式名称
'##ModelId=3E80E60002BF
Private m_szFormularName As String

'底限费用
'##ModelId=3E80E6050177
Private m_dbLimitCharge As Double
Private m_dbUpCharge As Double
'公式描述
'##ModelId=3E80E609030D
Private m_szFormulaComment As String

'折算项名称
'##ModelId=3E80E61302CE
Private m_szSplitItemName As String

'使用状态
'##ModelId=3E80E619009C
Private m_ESplitStatus As ESplitItemUseStatus

'折算类型
'##ModelId=3E80E61D03C8
Private m_ESplitType As ESplitItemType

'是否允许修改
'##ModelId=3E80E6220196
Private m_EAllowModify As ESplitItemAllowModify

Private m_oActiveUser As ActiveUser
Private m_nObjectStatus As EObjectStatus '对象状态

'初始化活动对象
'##ModelId=3E80E64603C8
Public Sub Init(poAUser As ActiveUser)
    On Error GoTo InitErr
    Set m_oActiveUser = poAUser
    m_nObjectStatus = ST_NormalObj
    Exit Sub
InitErr:
    Call RaiseError("", "Protocol:Init Method")
End Sub

'刷新对象
'##ModelId=3E80E6590157
Public Sub Identify(pszProtocolID As String)
    On Error GoTo IdentifyErr
    AssertActiveUserValid m_oActiveUser, ERR_Project
    m_szProtocolID = pszProtocolID
    RefreshObject
    m_nObjectStatus = ST_NormalObj
    Exit Sub
IdentifyErr:
    Call RaiseError(MyUnhandledError, "Protocol:Identify Method")
End Sub
Private Sub RefreshObject()
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "SELECT * FROM fin_settle_protocol_info " _
        & " WHERE protocol_id = " & TransFieldValueToString(m_szProtocolID)
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then RaiseError ERR_ProtocolNotExist  '此协议不存在出错
    m_szProtocolID = FormatDbValue(rsTemp!protocol_id)
    m_szProtocolName = FormatDbValue(rsTemp!protocol_name)
    m_szAnnotation = FormatDbValue(rsTemp!Annotation)
    m_EDefaultMark = FormatDbValue(rsTemp!default_mark)
    Set oDB = Nothing
    Set rsTemp = Nothing
End Sub
'新增协议表(状态改为新增)
'##ModelId=3E80E66A0399
Public Sub AddNew()
    AssertActiveUserValid m_oActiveUser, ERR_Protocol
    AssertHaveRight m_oActiveUser, RIGHT_Protocol
    m_nObjectStatus = ST_AddObj   '新增状态
    m_szAnnotation = ""
    m_szProtocolID = ""
    m_szProtocolName = ""
    m_EDefaultMark = Default
End Sub

'新增协议表
'##ModelId=3E80E6780119
Private Sub AddObject()
    On Error GoTo AddObjectErr
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If m_EDefaultMark = Default Then
        szSql = "UPDATE fin_settle_protocol_info SET default_mark = 1 "
        oDB.Execute (szSql)
    End If
    szSql = "INSERT fin_settle_protocol_info (protocol_id,protocol_name,annotation,default_mark) " _
        & " VALUES( " _
        & TransFieldValueToString(m_szProtocolID) & "," _
        & TransFieldValueToString(m_szProtocolName) & "," _
        & TransFieldValueToString(m_szAnnotation) & "," _
        & TransFieldValueToString(m_EDefaultMark) & ")"
    oDB.Execute (szSql)
    
'    szSql = "INSERT fin_charge_item_info (protocol_id,split_item_id,formular_name,limit_charge,formula_comment) " _
'        & " VALUES( " _
'        & TransFieldValueToString(m_szProtocolID) & "," _
'        & TransFieldValueToString("") & "," _
'        & TransFieldValueToString("") & "," _
'        & TransFieldValueToString("") & "," _
'        & TransFieldValueToString("") & ")"
'    oDb.Execute (szSql)
    
    Set oDB = Nothing
    Set rsTemp = Nothing
    WriteOperateLog m_oActiveUser, RIGHT_Protocol, "新增协议" & m_szProtocolID    '写日志
    Exit Sub
AddObjectErr:

    err.Raise err.Number, err.Source, err.Description
End Sub

'##ModelId=3E80E67D0399
Public Sub UpDate()

    AssertActiveUserValid m_oActiveUser, ERR_Protocol
    AssertHaveRight m_oActiveUser, RIGHT_Protocol
    AssertObjIsValid
    Select Case m_nObjectStatus
        Case ST_AddObj
            AddObject
        Case ST_EditObj
            UpDateToDB
        Case ST_NormalObj
        Case Else
        RaiseError ERR_NotAvailable
    End Select
    m_nObjectStatus = ST_NormalObj
    Exit Sub

End Sub

'##ModelId=3E80E6840177
Private Sub UpDateToDB()
    On Error GoTo UpDateDBErr
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If m_EDefaultMark = Default Then
        szSql = "UPDATE fin_settle_protocol_info SET default_mark = 1 "
        oDB.Execute (szSql)
    End If
    
    szSql = "UPDATE fin_settle_protocol_info SET " _
        & " protocol_name = " & TransFieldValueToString(m_szProtocolName) & "," _
        & " annotation = " & TransFieldValueToString(m_szAnnotation) & "," _
        & " default_mark = " & TransFieldValueToString(m_EDefaultMark) _
        & " WHERE protocol_id = " & TransFieldValueToString(m_szProtocolID)
    oDB.Execute (szSql)
    'WriteOperateLog 写日志
    Exit Sub
UpDateDBErr:
    Call RaiseError(ERR_Protocol)
End Sub

'##ModelId=3E80E689038A
Public Sub Delete()
    On Error GoTo DeleteErr
    AssertHaveRight m_oActiveUser, RIGHT_Protocol
    AssertObjIsValid
    If m_nObjectStatus <> ST_AddObj Then
        DeleteObject
    End If
    Exit Sub
DeleteErr:
    Call RaiseError(ERR_ProtocolNotExist)
End Sub

'##ModelId=3E80E68E01A5
Private Sub DeleteObject()
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "DELETE FROM fin_charge_item_info " _
        & " WHERE protocol_id = " & TransFieldValueToString(m_szProtocolID)
    oDB.Execute (szSql)
'    If rsTemp.RecordCount <> 0 Then RaiseError ERR_ProtocolUse  '协议已经使用，不能使用
    szSql = "DELETE fin_settle_protocol_info WHERE protocol_id = " & TransFieldValueToString(m_szProtocolID)
    oDB.Execute (szSql)
    'WriteOperateLog '写日志
    Exit Sub
End Sub

'设置费用名表和拆算费用项表
'##ModelId=3E80E692000F
Public Sub AddChargeItemInfo(pszSplitItemID As String, pszFormularName As String, pdbLimitCharge As Double, pszFormulaComment As String, pszUpCharge As Double)
    On Error GoTo AddChargeitemInfoErr
    AssertActiveUserValid m_oActiveUser, ERR_Protocol
    AssertHaveRight m_oActiveUser, RIGHT_Protocol
    AssertObjIsValid
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "INSERT fin_charge_item_info (protocol_id,split_item_id,formular_name,limit_charge,up_charge,formula_comment) " _
        & " VALUES( " _
        & TransFieldValueToString(m_szProtocolID) & "," _
        & TransFieldValueToString(pszSplitItemID) & "," _
        & TransFieldValueToString(pszFormularName) & "," _
        & TransFieldValueToString(pdbLimitCharge) & "," _
        & TransFieldValueToString(pszUpCharge) & " , " _
        & TransFieldValueToString(pszFormulaComment) & ")"
    oDB.Execute (szSql)
    'WriteOperateLog 写日志
    Exit Sub
AddChargeitemInfoErr:
    err.Raise err.Number, err.Source, err.Description
End Sub

'修改数据费用名表
'##ModelId=3E80E697035B
Public Sub UpDateChargeItemInfo(pszSplitItemID As String, pszFormularName As String, pdbLimitCharge As Double, pszFormulaComment As String, pdbUpCharge As Double)
    On Error GoTo UpDateChargeitemInfoErr
    AssertActiveUserValid m_oActiveUser, ERR_Protocol
    AssertHaveRight m_oActiveUser, RIGHT_Protocol
    AssertObjIsValid
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "UPDATE fin_charge_item_info SET " _
        & " split_item_id = " & TransFieldValueToString(pszSplitItemID) & "," _
        & " formular_name = " & TransFieldValueToString(pszFormularName) & "," _
        & " limit_charge = " & TransFieldValueToString(pdbLimitCharge) & "," _
        & " up_charge= " & TransFieldValueToString(pdbUpCharge) & "," _
        & " formula_comment = " & TransFieldValueToString(pszFormulaComment) _
        & " WHERE split_item_id = " & TransFieldValueToString(pszSplitItemID) & "AND" _
        & " protocol_id = " & TransFieldValueToString(m_szProtocolID)
    oDB.Execute (szSql)
    'WriteOperateLog 写日志
    Exit Sub
UpDateChargeitemInfoErr:
    err.Raise err.Number, err.Source, err.Description
End Sub

'得到费用具细信息
'##ModelId=3E80E69C034B
Public Function GetChargeitemInfo(Optional pszSplitItemID As String = "") As TFinChargeItemInfo()
    AssertActiveUserValid m_oActiveUser, ERR_Protocol
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim szSqlTemp As String
    Dim rsTemp As Recordset
    Dim TTemp() As TFinChargeItemInfo
    Dim i As Integer
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    If pszSplitItemID <> "" Then
        szSqlTemp = " AND f.split_item_id = " & TransFieldValueToString(pszSplitItemID)
    End If
    szSql = "SELECT f.split_item_id,s.split_item_name,f.formular_name,f.limit_charge,f.formula_comment,s.split_status,s.split_type,s.allow_modify , f.up_charge " _
        & " FROM fin_charge_item_info f,split_item_info s  " _
        & " WHERE f.protocol_id = " & TransFieldValueToString(m_szProtocolID) _
        & " AND f.split_item_id =s.split_item_id " & szSqlTemp
    Set rsTemp = oDB.Execute(szSql)
    
    If rsTemp.RecordCount = 0 Then Exit Function
    ReDim TTemp(1 To rsTemp.RecordCount)
    For i = 1 To rsTemp.RecordCount
        TTemp(i).ProtocolID = m_szProtocolID
        TTemp(i).ProtocolName = m_szProtocolName
        TTemp(i).SplitItemID = FormatDbValue(rsTemp!split_item_id)
        TTemp(i).SplitItemName = FormatDbValue(rsTemp!split_item_name)
        TTemp(i).AllowModify = FormatDbValue(rsTemp!allow_modify)
        TTemp(i).FormulaComment = FormatDbValue(rsTemp!formula_comment)
        TTemp(i).FormularName = FormatDbValue(rsTemp!formular_name)
        TTemp(i).SplitStatus = FormatDbValue(rsTemp!split_status)
        TTemp(i).SplitType = FormatDbValue(rsTemp!split_type)
        TTemp(i).LimitCharge = FormatDbValue(rsTemp!limit_charge)
        TTemp(i).UpCharge = FormatDbValue(rsTemp!up_charge)
        
        rsTemp.MoveNext
    Next i
    GetChargeitemInfo = TTemp
    Exit Function
End Function

'删除费用名表
'##ModelId=3E80E6A501B5
Public Sub DeleteChargeitemInfo(pszSplitItemID As String)
    On Error GoTo DeleteChargeitemInfoErr
    AssertActiveUserValid m_oActiveUser, ERR_Protocol
    AssertHaveRight m_oActiveUser, RIGHT_Protocol
    AssertObjIsValid
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDB.ConnectionString = GetConnectionStr(cszSettle)
    szSql = "DELETE fin_charge_item_info " _
        & " WHERE protocol_id = " & TransFieldValueToString(m_szProtocolID) _
        & " AND split_item_id = " & TransFieldValueToString(m_szSplitItemID)
    oDB.Execute (szSql)
    'WriteOperateLog 写日志    Exit Sub
DeleteChargeitemInfoErr:
    err.Raise err.Number, err.Source, err.Description
End Sub

'##ModelId=3E8219BE02EE
'Private Sub Class_Initialize()
'    #If DebugMode Then
'        'get the next available class ID, and print out
'        'that the class was created successfully
'        mlClassDebugID = GetNextClassDebugID()
'        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
'    #End If
'End Sub

'##ModelId=3E8219BE036B
'Private Sub Class_Terminate()
'    #If DebugMode Then
'    'the class is being destroyed
'    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
'    #End If
'End Sub
'
'#If DebugMode Then
'    '##ModelId=3E8225D70242
'    Public Property Get ClassDebugID() As Long
'        'if we are in debug mode, surface this property that consumers can query
'        ClassDebugID = mlClassDebugID
'    End Property
'#End If

'##ModelId=3E82291102CE
Public Property Get AllowModify() As ESplitItemAllowModify
    AllowModify = m_EAllowModify
End Property

'##ModelId=3E8229110196
Public Property Get SplitType() As ESplitItemType
    SplitType = m_ESplitType
End Property

'##ModelId=3E822911006D
Public Property Get SplitStatus() As ESplitItemUseStatus
    SplitStatus = m_ESplitStatus
End Property

'##ModelId=3E822910031C
Public Property Get SplitItemName() As String
    SplitItemName = m_szSplitItemName
End Property

'##ModelId=3E82291001F4
Public Property Get FormulaComment() As String
    FormulaComment = m_szFormulaComment
End Property

'##ModelId=3E82291000CB
Public Property Get LimitCharge() As Double
    LimitCharge = m_dbLimitCharge
End Property

Public Property Get UpCharge() As Double
    UpCharge = m_dbUpCharge
End Property

'##ModelId=3E82290F0399
Public Property Get FormularName() As String
    FormularName = m_szFormularName
End Property

'##ModelId=3E82290F0271
Public Property Get SplitItemID() As String
    SplitItemID = m_szSplitItemID
End Property

'##ModelId=3E82290F0157
Public Property Get DefaultMark() As EDefault
    DefaultMark = m_EDefaultMark
End Property

'##ModelId=3E82290E033C
Public Property Let DefaultMark(ByVal Value As EDefault)
    m_EDefaultMark = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E82290E0232
Public Property Get Annotation() As String
    Annotation = m_szAnnotation
End Property

'##ModelId=3E82290E002E
Public Property Let Annotation(ByVal Value As String)
    m_szAnnotation = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E82290D030D
Public Property Get ProtocolName() As String
    ProtocolName = m_szProtocolName
End Property

'##ModelId=3E82290D0128
Public Property Let ProtocolName(ByVal Value As String)
    m_szProtocolName = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E82290D000F
Public Property Get ProtocolID() As String
    ProtocolID = m_szProtocolID
End Property


'##ModelId=3E82290C0213
Public Property Let ProtocolID(ByVal Value As String)
    m_szProtocolID = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'**************************************************
'Member Code:S3
'Brief Description:测试对象是否有效
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then RaiseError ERR_NotAvailable
End Sub

'**************************************************
'Member Code:S4
'Brief Description:测试对象是否有效
'Engineer:
'Date Generated:1999/10/11
'Last Revision Date:1999/10/12
'**************************************************
Private Sub AssertObjIsValid()
    AssertActiveUserValid m_oActiveUser, ERR_Protocol
    AssertStatusAvailable
End Sub

