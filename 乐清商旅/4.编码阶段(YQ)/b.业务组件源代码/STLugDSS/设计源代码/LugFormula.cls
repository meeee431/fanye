VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LugFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'行包拆算公式
Option Base 0
Option Explicit


    '##ModelId=3E287D8F029F
Private mlClassDebugID As Long
Const MyUnhandledError = 123

'公式代码
'##ModelId=3E230FB200A1
Private mszFormulaID As String

'公式名称
'##ModelId=3E230FB200AB
Private mszFormulaName As String

'公式内容
'##ModelId=3E230FB200B5
Private mszFormulaText As String
Private m_oActiveUser As ActiveUser
Private m_nObjectStatus As EObjectStatus
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝

'出错信息
Public Enum ELuggageFormulaErr
                                                    ' 起始号为32400
    ERR_FormulaIDNullOrNoAdd = ERR_LuggageFromula + 20  '32420 公式名称为空或非新增状态
    ERR_FormulaExist = ERR_LuggageFromula + 21  '32421公式已经存在或被使用
    ERR_FormulaNotExist = ERR_LuggageFromula + 22 '32422公式不存在
    ERR_PriceItemNotExist = ERR_LuggageFromula + 23 '32423票价项不存在
    ERR_FormulaTextDiffer = ERR_LuggageFromula + 24 '32424公式不能有两种以上托运方式
    '出错信息
End Enum

Public Enum RIGHTLuggageFormula
                                                                                            '32500
    RIGHT_LuggageFormulaAdd = ERR_LuggageFromula + cnMidRightBegin + cnMidRightStep * 1   '32506 公式新增权限
    RIGHT_LuggageFormulaEdit = ERR_LuggageFromula + cnMidRightBegin + cnMidRightStep * 2   '32512 公式编辑权限
    RIGHT_LuggageFormulaDel = ERR_LuggageFromula + cnMidRightBegin + cnMidRightStep * 3   ' 32518 公式删除权限
End Enum
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝

'初始化活动对象
'##ModelId=3E230F6D030E
Public Sub Init(poAUser As ActiveUser)
    On Error GoTo InitErr
    
    Set m_oActiveUser = poAUser
        

    Exit Sub
InitErr:
'    Call RaiseError(MyUnhandledError, "LugFormula:Init Method")
End Sub

'指定公式
'##ModelId=3E230F6D0318
Public Sub Identify(FormulaID As String)
    AssertActiveUserValid m_oActiveUser, ERR_LuggageFromula
    On Error GoTo IdentifyErr
    mszFormulaID = FormulaID
    RefreshInfo
    m_nObjectStatus = ST_NormalObj
    Exit Sub
IdentifyErr:
    Call RaiseError(MyUnhandledError, "LugFormula:Identify Method")
End Sub

Private Sub RefreshInfo()
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    szSql = "SELECT * FROM lug_settle_formula_info WHERE formuar_id = " & TransFieldValueToString(mszFormulaID)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then
        RaiseError ERR_FormulaNotExist '公式不存在
    End If
    mszFormulaID = FormatDbValue(rsTemp!formuar_id)
    mszFormulaName = FormatDbValue(rsTemp!formular_name)
    mszFormulaText = FormatDbValue(rsTemp!formular_content)
    
    
End Sub
'新增公式
'##ModelId=3E230F6D0322
Public Sub AddNew()
    On Error GoTo AddNewErr
    AssertActiveUserValid m_oActiveUser, ERR_LuggageFromula
    m_nObjectStatus = ST_AddObj
    mszFormulaID = ""
    mszFormulaName = ""
    mszFormulaText = ""
    mlClassDebugID = -1

    Exit Sub
AddNewErr:
    Call RaiseError(MyUnhandledError, "LugFormula:AddNew Method")
End Sub

'更新公式
'##ModelId=3E230F6D032C
Public Sub Update()
    AssertActiveUserValid m_oActiveUser, ERR_LuggageFromula
    
    On Error GoTo UpdateErr
    If m_nObjectStatus = ST_AddObj Then
        AddObject
    ElseIf m_nObjectStatus = ST_EditObj Then
        AssertHaveRight m_oActiveUser, RIGHT_LuggageFormulaEdit
        UpdateToDB
    End If
    m_nObjectStatus = ST_NormalObj
    Exit Sub
UpdateErr:
    Call RaiseError(MyUnhandledError, "LugFormula:Update Method")
End Sub

Private Sub AddObject()
On Error GoTo AddErr
    Dim oDb As New RTConnection
    Dim szSql As String
    If mszFormulaID = "" Or mszFormulaID = Null Or m_nObjectStatus <> ST_AddObj Then
        RaiseError ERR_FormulaIDNullOrNoAdd  '出错 '公式名称为空或者非新增状态
    End If
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "INSERT lug_settle_formula_info " _
        & "(formuar_id,formular_name,formular_content) " _
        & "VALUES ( " _
        & TransFieldValueToString(mszFormulaID) & "," _
        & TransFieldValueToString(mszFormulaName) & "," _
        & TransFieldValueToString(mszFormulaText) & ")"
    oDb.Execute (szSql)
    WriteOperateLog m_oActiveUser, RIGHT_LuggageFormulaAdd, "新增公式" & mszFormulaID & "[" & mszFormulaName & "]" & "(" & mszFormulaText & ")"
    Exit Sub
AddErr:
    RaiseError ERR_FormulaExist '该公式ID已经存在
End Sub

Private Sub UpdateToDB()
On Error GoTo FormularErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim mszOldFormularName As String
    Dim mszOldFormularText As String
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT * FROM lug_settle_formula_info WHERE formuar_id=" & TransFieldValueToString(mszFormulaID)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        mszOldFormularName = FormatDbValue(rsTemp!formular_name)
        mszOldFormularText = FormatDbValue(rsTemp!formular_content)
    End If
    szSql = "UPDATE lug_settle_formula_info " _
        & " SET " _
        & " formular_name = " & TransFieldValueToString(mszFormulaName) & "," _
        & " formular_content = " & TransFieldValueToString(mszFormulaText) _
        & " WHERE formuar_id = " & TransFieldValueToString(mszFormulaID)
    oDb.Execute szSql
    WriteOperateLog m_oActiveUser, RIGHT_LuggageFormulaEdit, "将公式" & mszFormulaID & "[" & mszOldFormularName & "](" & mszOldFormularText & ") 更改为" & mszFormulaID & "[" & mszFormulaName & "](" & mszOldFormularText & ")"
    
    
    Exit Sub
FormularErr:
    RaiseError ERR_FormulaExist
End Sub
'删除公式
'##ModelId=3E230F6D0340
Public Sub Delete()
    AssertActiveUserValid m_oActiveUser, ERR_LuggageFromula
    AssertHaveRight m_oActiveUser, RIGHT_LuggageFormulaDel
    On Error GoTo DeleteErr
    If m_nObjectStatus <> ST_AddObj Then
        DeleteObject
    End If
    Exit Sub
DeleteErr:
    Call RaiseError(MyUnhandledError, "LugFormula:Delete Method")
End Sub

Private Sub DeleteObject()
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim mszOldName As String
    
    oDb.ConnectionString = GetConnectionStr(cszSplit)
    
    szSql = "SELECT * FROM fin_luggage_chargeitem_lst l, lug_settle_formula_info i " _
        & " WHERE formuar_id = " & TransFieldValueToString(mszFormulaID) _
        & " AND l.formular_name = i.formular_name "
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 0 Then RaiseError ERR_FormulaExist '该公式被使用
'    mszOldName = FormatDbValue(rsTemp!formular_name)
    szSql = "DELETE lug_settle_formula_info " _
        & " WHERE formuar_id = " & TransFieldValueToString(mszFormulaID)
    oDb.Execute szSql
    WriteOperateLog m_oActiveUser, RIGHT_LuggageFormulaDel, "删除公式" & mszFormulaID & "[" & mszOldName & "]"
    Set oDb = Nothing
End Sub
'得到所有拆算公式，返回由[公式代码]、[公式名称]、[公式内容]组成的二维数组
'##ModelId=3E234A5500CE
Public Function GetAllFormulas() As String()
    On Error GoTo GetAllFormulasErr
    
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim aTempFormulas() As String
    Dim i As Integer
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT * FROM lug_settle_formula_info"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then
'        RaiseError ERR_FormulaNotExist '公式不存在
        Exit Function
    End If
    If rsTemp.RecordCount > 0 Then
        ReDim aTempFormulas(1 To rsTemp.RecordCount, 1 To 3)
        For i = 1 To rsTemp.RecordCount
            aTempFormulas(i, 1) = FormatDbValue(rsTemp!formuar_id)
            aTempFormulas(i, 2) = FormatDbValue(rsTemp!formular_name)
            aTempFormulas(i, 3) = FormatDbValue(rsTemp!formular_content)
            rsTemp.MoveNext
        Next
        
    End If
    
    GetAllFormulas = aTempFormulas
    Exit Function
GetAllFormulasErr:
    Call RaiseError(MyUnhandledError, "LugFormula:GetAllFormulas Method")
End Function

'得到所有允许被用于组成拆算公式的可计算项目，返回由名称和有效性组成的二维数组
'逻辑参见STSplit.GetAllLuggageItem方法，(省略托运方式)
'##ModelId=3E25A59400FE
Public Function GetAllFormulaItems() As String()
    On Error GoTo GetAllFormulaItemsErr

    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim szaTempFormulas() As String
    Dim szAcceptType As String
    Dim i As Integer
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT * FROM luggage_price_item_info WHERE use_mark <> 1"
'    szSql = "SELECT DISTINCT  chinese_name,use_mark FROM luggage_price_item_info WHERE use_mark <> 1"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then
        RaiseError ERR_PriceItemNotExist '票价项不存在
    End If
    ReDim szaTempFormulas(1 To 2 + rsTemp.RecordCount, 1 To 2)
    szaTempFormulas(1, 1) = "总运价"
    szaTempFormulas(1, 2) = "0"
      
    szaTempFormulas(2, 1) = "总件数"
    szaTempFormulas(2, 2) = "0"
    If rsTemp.RecordCount > 0 Then
        For i = 1 To rsTemp.RecordCount
            szAcceptType = GetLuggageTypeString(FormatDbValue(rsTemp!accept_type))  '得行包类型项
            szaTempFormulas(2 + i, 1) = "[" & szAcceptType & "]" & FormatDbValue(rsTemp!chinese_name)
'            szaTempFormulas =”[快件]运费“
'            szaTempFormulas(2 + i, 1) = FormatDbValue(rsTemp!chinese_name)
            szaTempFormulas(2 + i, 2) = FormatDbValue(rsTemp!use_mark)
            rsTemp.MoveNext
        Next
    End If
    GetAllFormulaItems = szaTempFormulas

    Exit Function
GetAllFormulaItemsErr:
    Call RaiseError(MyUnhandledError, "LugFormula:GetAllFormulaItems Method")
End Function

'##ModelId=3E287D8F032C
Private Sub Class_Initialize()
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'##ModelId=3E287D8F034B
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

#If DebugMode Then
    '##ModelId=3E287D8F02DE
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
'##ModelId=3E28A808033C
Public Property Get FormulaID() As String
   Let FormulaID = mszFormulaID
   If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property


'##ModelId=3E28A80802AF
Public Property Let FormulaID(ByVal Value As String)
    Let mszFormulaID = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E28A8090138
Public Property Get FormulaText() As String
   Let FormulaText = mszFormulaText
   If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E28A809009C
Public Property Let FormulaText(ByVal Value As String)
    Dim szTemp As String
    Dim szValue As String
    Dim nCount As Integer
    Dim nLen As Integer, i As Integer
    nCount = Len(Value)
    nLen = 1
    For i = 1 To nCount
        nLen = InStr(i, Value, "[")
        If nLen <> 0 Then
            szTemp = szTemp + Mid(Value, nLen + 1, 2)
            i = nLen + 1
        Else
            Exit For
        End If
    Next i
    If Len(szTemp) <> 0 Then
        szValue = Mid(szTemp, 1, 2)
        For i = 1 To Len(szTemp) Step 2
            If szValue <> Mid(szTemp, i, 2) Then
                RaiseError ERR_FormulaTextDiffer
            End If
        Next i
    End If
    Let mszFormulaText = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E28A809003E
Public Property Get FormulaName() As String
   Let FormulaName = mszFormulaName
   If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

'##ModelId=3E28A808038A
Public Property Let FormulaName(ByVal Value As String)
    Let mszFormulaName = Value
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property


