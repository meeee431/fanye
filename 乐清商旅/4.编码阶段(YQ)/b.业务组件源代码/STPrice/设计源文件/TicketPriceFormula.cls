VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TicketPriceFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**********************************************************
'* Source File Name:TicketPriceFormula.cls
'* Project Name:STPrice.vbp
'* Engineer:
'* Data Generated:
'* Last Revision Date:
'* Brief Description:票价计算公式
'* Relational Document:
'**********************************************************
Option Explicit

Public Enum EFormulaIsDefault
    IsDefaultFormula = 1
    NotDefaultFormula = 0
End Enum

Public Enum EErrTicketPriceFormula
    ERR_TPFNotAddObj = ERR_TicketPriceFormula + ERR_NotAddObj  '对象不处于新增状态
    ERR_TPFAddObj = ERR_TicketPriceFormula + ERR_AddObj '对象新增状态
    ERR_TPFEditObj = ERR_TicketPriceFormula + ERR_EditObj '对象编辑状态
    ERR_TPFNormalObj = ERR_TicketPriceFormula + ERR_NormalObj '对象正常状态
    ERR_TPFNotAvailable = ERR_TicketPriceFormula + ERR_NotAvailable '对象无效状态


    ERR_TPFNoFormulaName = ERR_TicketPriceFormula + 21 '未设置票价公式名
    ERR_TPFFormulaNoExist = ERR_TicketPriceFormula + 22 '无票价公式
    ERR_TPFCantInsertFormulaList = ERR_TicketPriceFormula + 23 '插入公式明细表错误
    ERR_TPFCantUpdateFormulaList = ERR_TicketPriceFormula + 24 'UPDATE公式明细表错误
    ERR_TPFFormulaInfoNoExist = ERR_TicketPriceFormula + 25 '公式明细信息丢失
    ERR_TPFItemFandPModifyIllegleParam = ERR_TicketPriceFormula + 26 '修改票价公式某票价项参数信息时参数不合法
End Enum

Public Enum ERightRoutePriceFormula
    RIGHT_TPFAddFormula = ERR_TicketPriceFormula + cnMidRightBegin + cnMidRightStep * 1 '添加公式权限
    RIGHT_TPFModifyDB = ERR_TicketPriceFormula + cnMidRightBegin + cnMidRightStep * 2 '票价公式修改权限
    RIGHT_TPFDeleteFormula = ERR_TicketPriceFormula + cnMidRightBegin + cnMidRightStep * 3 '删除公式权限
    RIGHT_TPFSetAsDefault = ERR_TicketPriceFormula + cnMidRightBegin + cnMidRightStep * 4 '设定为默认票价表公式
End Enum

'----------------------
'各类特有的变量
Private m_szFormulaName As String
Private m_IsDefault As EFormulaIsDefault
Private m_szAnnotation As String

'---------------------
'所有的实体类都有的私有类变量<<
Private m_nObjectStatus As EObjectStatus

'所有的类都有的私有类变量
Private m_oActiveUser As ActiveUser

Public Sub Init(poAUser As ActiveUser)
    Set SelfUser = poAUser
End Sub

Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
    m_oActiveUser.IAmLive
End Property

Public Property Set SelfUser(vNewValue As ActiveUser)
    If m_nObjectStatus = ST_AddObj Then
        ShowError ERR_TPFAddObj
    ElseIf m_nObjectStatus = ST_EditObj Then
        ShowError ERR_TPFEditObj
    ElseIf m_nObjectStatus = ERR_NormalObj Then
        ShowError ERR_TPFNormalObj
    Else
        Set m_oActiveUser = vNewValue
    End If
    m_oActiveUser.IAmLive
End Property


Private Sub Class_Initialize()
    m_nObjectStatus = ST_NotAvailable '类刚建立时处于不可用状态
End Sub
Private Sub AssertStatusAvailable() '测试对象的状态是否有效（不为无效ECNotAvailable状态）
    If m_nObjectStatus = ST_NotAvailable Then ShowError ERR_TPFNotAvailable
End Sub
Private Sub AssertObjIsValid() '测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
    AssertActiveUserValid m_oActiveUser, ERR_TicketPriceFormula
    AssertStatusAvailable
End Sub
Private Sub AssertCanCallSub()
    AssertActiveUserValid m_oActiveUser, ERR_TicketPriceFormula
    If m_nObjectStatus = ST_NotAvailable Then ShowError ERR_TPFNotAvailable
    If m_nObjectStatus = ST_AddObj Then ShowError ERR_TPFAddObj
End Sub

Public Sub Identify(pszObjID As String)
    AssertActiveUserValid m_oActiveUser, ERR_TicketPriceFormula
    m_szFormulaName = pszObjID
    RefreshMemoryInfo
    m_nObjectStatus = ST_NormalObj
End Sub


Private Sub RefreshMemoryInfo() '按照对象现在的ID将数据库中相应的信息读至对象内
    AssertActiveUserValid m_oActiveUser, ERR_TicketPriceFormula
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim odb As New RTConnection
    szSql = "SELECT * FROM Price_formula_info WHERE formula_name='" & m_szFormulaName & "'"
    odb.ConnectionString = GetConnectionStr(cszPriceMan)
    Set rsTemp = odb.Execute(szSql)
    
    If rsTemp.RecordCount = 0 Then ShowError ERR_TPFFormulaNoExist
    If rsTemp.RecordCount > 0 Then
        If FormatDbValue(rsTemp("default_execute_mark")) = "1" Then
            m_IsDefault = EFormulaIsDefault.IsDefaultFormula
        Else
            m_IsDefault = EFormulaIsDefault.NotDefaultFormula
        End If
        m_szAnnotation = FormatDbValue(rsTemp("annotation"))
    End If
    Set rsTemp = Nothing
    Set odb = Nothing
End Sub

Public Sub AddNew() '调用新增方法后，对象处于新增状态。
    AssertActiveUserValid m_oActiveUser, ERR_TicketPriceFormula
    AssertHaveRight m_oActiveUser, RIGHT_TPFAddFormula
    m_nObjectStatus = ST_AddObj '对象状态转换

End Sub

Private Sub AddObject()
    Dim isDBPrimTableSQLExecute As Boolean '标注是主表操作错或子表错
    Dim szSql As String
    Dim odb As New RTConnection
    
    If m_szFormulaName = "" Or m_szFormulaName = Null Then
        ShowError ERR_TPFNoFormulaName '主键空
    Else
        odb.ConnectionString = GetConnectionStr(cszPriceMan)
    odb.BeginTrans  'START事务
On Error GoTo Here
        isDBPrimTableSQLExecute = True
        szSql = "INSERT  Price_formula_info(formula_name,default_execute_mark,annotation) " _
        & " VALUES('" _
        & m_szFormulaName & " '," _
        & 0 & ",'" _
        & m_szAnnotation & "')"
        odb.Execute szSql
    End If

    odb.CommitTrans '提交
    WriteOperateLog m_oActiveUser, RIGHT_TPFAddFormula, "增加公式" & m_szFormulaName
     
Exit Sub
Here:
    odb.RollbackTrans '回滚
    If isDBPrimTableSQLExecute = True Then
        AssertAddObjectError ERR_TicketPriceFormula, odb
    Else
        ShowError ERR_TPFCantInsertFormulaList
    End If
End Sub

Public Sub Update()
    AssertObjIsValid
    
    If m_nObjectStatus = ST_AddObj Then '如对象处于新增状态，则新增一对象
        AssertHaveRight m_oActiveUser, RIGHT_TPFAddFormula '权限判断在AddNew
        AddObject
    ElseIf m_nObjectStatus = ST_EditObj Then '如对象处于修改状态，则更新它
        AssertHaveRight m_oActiveUser, RIGHT_TPFModifyDB '权限判断
        UpdateToDB
    End If
    m_nObjectStatus = ST_NormalObj
End Sub

Private Sub UpdateToDB()
    Dim szSql As String
    Dim odb As New RTConnection
    Dim szTemp As String
'        If m_IsDefault = IsDefaultFormula Then
'            szTemp = "1"
'        Else
'            szTemp = "0"
'        End If
        szTemp = "0"
On Error GoTo Here
    odb.ConnectionString = GetConnectionStr(cszPriceMan)
    szSql = "UPDATE Price_formula_info  SET " _
    & " default_execute_mark=" & szTemp _
    & ",annotation='" & m_szAnnotation _
    & "' WHERE formula_name='" & m_szFormulaName & "'"

    odb.Execute szSql
    WriteOperateLog m_oActiveUser, RIGHT_TPFModifyDB, "修改票价计算公式" & m_szFormulaName
On Error GoTo 0
    Set odb = Nothing
Exit Sub
Here:
    AssertUpdateObjectError ERR_TicketPriceFormula, odb
    Set odb = Nothing
End Sub


Public Property Get FormulaName() As String
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_QueryPrice
    FormulaName = m_szFormulaName
End Property


Public Property Let FormulaName(ByVal vNewValue As String)
    AssertObjIsValid
    If m_nObjectStatus = ST_AddObj Then '新增时可修改
        m_szFormulaName = vNewValue
    Else
        ShowError ERR_TPFNotAddObj
    End If
End Property

Public Property Get Annotation() As String
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_QueryPrice
    Annotation = m_szAnnotation
End Property


Public Property Let Annotation(ByVal vNewValue As String)
    AssertObjIsValid
    m_szAnnotation = vNewValue
    If m_nObjectStatus <> ST_AddObj Then m_nObjectStatus = ST_EditObj
End Property

Public Property Get IsDefault() As EFormulaIsDefault
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_QueryPrice
    IsDefault = m_IsDefault
End Property

Public Sub Delete() '调用删除方法后，对象处于不可用状态
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_TPFDeleteFormula
    If m_nObjectStatus <> ST_AddObj Then
        DeleteObject
    End If
    m_nObjectStatus = ST_NotAvailable '对象状态转换
End Sub
Private Sub DeleteObject()
    Dim szSql As String
    Dim odb As New RTConnection
    
On Error GoTo Here
        odb.ConnectionString = GetConnectionStr(cszPriceMan)
        szSql = "DELETE  FROM Price_formula_info WHERE formula_name='" & m_szFormulaName & "'"
        
        odb.Execute szSql
        WriteOperateLog m_oActiveUser, RIGHT_TPFDeleteFormula, "删除票价计算公式" & m_szFormulaName
        Set odb = Nothing
    Exit Sub
Here:
    AssertDeleteObjectError ERR_TicketPriceFormula, odb
    Set odb = Nothing
End Sub


Public Function GetAllFAndP(PriceTableID As String) As TItemFAndP()
    AssertCanCallSub
On Error GoTo Here
    Dim aTinfoTemp() As TItemFAndP, lRowCount As Long
    Dim rsTemp As Recordset
    Dim szSql As String, odb As New RTConnection
    Dim i As Integer

    szSql = " SELECT * FROM price_formula_detail_lst WHERE formula_name='" & m_szFormulaName & "' AND price_table_id='" & PriceTableID & "' ORDER BY price_item "
    odb.ConnectionString = GetConnectionStr(cszPriceMan)
    Set rsTemp = odb.Execute(szSql, lRowCount)
    If lRowCount = 0 Then ShowError ERR_TPFFormulaInfoNoExist     'SHOWERROR
    If rsTemp.RecordCount > 0 Then
        ReDim aTinfoTemp(1 To rsTemp.RecordCount)
        For i = 1 To rsTemp.RecordCount
            aTinfoTemp(i).szFormula = FormatDbValue(rsTemp!item_formula)
            aTinfoTemp(i).szPriceItem = FormatDbValue(rsTemp!price_item)
            aTinfoTemp(i).sgParam1 = FormatDbValue(rsTemp!parameter_1)
            aTinfoTemp(i).sgParam2 = FormatDbValue(rsTemp!parameter_2)
            aTinfoTemp(i).sgParam3 = FormatDbValue(rsTemp!parameter_3)
            aTinfoTemp(i).sgParam4 = FormatDbValue(rsTemp!parameter_4)
            aTinfoTemp(i).sgParam5 = FormatDbValue(rsTemp!parameter_5)
            rsTemp.MoveNext
         Next i
    End If
    GetAllFAndP = aTinfoTemp
On Error GoTo 0
Set rsTemp = Nothing
Set odb = Nothing
Exit Function
Here:
    err.Raise err.Number, , err.Description
    Set rsTemp = Nothing
    Set odb = Nothing
End Function

Public Function ModifyItemFAndP(FormulaInfo As TItemFAndP, PriceTableID As String, PriceFormulaName As String)
    AssertCanCallSub
    AssertHaveRight m_oActiveUser, RIGHT_TPFModifyDB

    Dim szSql As String, odb As New RTConnection, ItemNum As Integer
    Dim sgPar1 As String, sgPar2 As String, sgPar3 As String, sgPar4 As String, sgPar5 As String, szPriceItem As String
    Dim szItemNum As String, szItemFormu As String
On Error GoTo there
    ItemNum = CInt(FormulaInfo.szPriceItem)
On Error GoTo 0
    
    szItemFormu = StrConv(FormulaInfo.szFormula, vbFromUnicode)
    Dim nLenTemp As Integer
    nLenTemp = LenB(szItemFormu)
    If nLenTemp > 99 Or nLenTemp < 0 Then ShowError ERR_TPFItemFandPModifyIllegleParam
    szPriceItem = FormulaInfo.szPriceItem
    szItemFormu = FormulaInfo.szFormula
    sgPar1 = FormulaInfo.sgParam1
    sgPar2 = FormulaInfo.sgParam2
    sgPar3 = FormulaInfo.sgParam3
    sgPar4 = FormulaInfo.sgParam4
    sgPar5 = FormulaInfo.sgParam5
'    If IsNumeric(sgPar1) Then
'       If Len(sgPar1) > 1 And Left(sgPar1, 1) <> 0 Then
'          If (sgPar1 < -9999.999 Or sgPar1 > 9999.999) Then ShowError m_oActiveUser, ERR_TPFItemFandPModifyIllegleParam
'       End If
'    End If
'    If IsNumeric(sgPar2) Then
'        If Len(sgPar2) > 1 And Left(sgPar2, 1) <> 0 Then
'           If (sgPar2 < -9999.999 Or sgPar2 > 9999.999) Then ShowError m_oActiveUser, ERR_TPFItemFandPModifyIllegleParam
'        End If
'    End If
'    If IsNumeric(sgPar3) Then
'        If Len(sgPar3) > 1 And Left(sgPar3, 1) <> 0 Then
'           If (sgPar3 < -9999.999 Or sgPar3 > 9999.999) Then ShowError m_oActiveUser, ERR_TPFItemFandPModifyIllegleParam
'        End If
'    End If
'    If IsNumeric(sgPar4) Then
'        If Len(sgPar4) > 1 And Left(sgPar4, 1) <> 0 Then
'           If (sgPar4 < -9999.999 Or sgPar4 > 9999.999) Then ShowError m_oActiveUser, ERR_TPFItemFandPModifyIllegleParam
'        End If
'    End If
'    If IsNumeric(sgPar5) Then
'        If Len(sgPar5) > 1 And Left(sgPar5, 1) <> 0 Then
'           If (sgPar5 < -9999.999 Or sgPar5 > 9999.999) Then ShowError m_oActiveUser, ERR_TPFItemFandPModifyIllegleParam
'        End If
'    End If
'    If (sgPar1 < -9999.999 Or sgPar1 > 9999.999) Or (sgPar2 < -9999.999 Or sgPar2 > 9999.999) Or (sgPar3 < -9999.999 Or sgPar3 > 9999.999) Or (sgPar4 < -9999.999 Or sgPar4 > 9999.999) Or (sgPar5 < -9999.999 Or sgPar5 > 9999.999) Then
'        ShowError m_oActiveUser, ERR_TPFItemFandPModifyIllegleParam
'    End If
    
    If (ItemNum > 0 And ItemNum < 10) Or ItemNum = 0 Then
        szItemNum = "000" & CStr(ItemNum)
    ElseIf ItemNum > 9 And ItemNum < 16 Then
    
        szItemNum = "00" & CStr(ItemNum)
    Else
        ShowError ERR_TPFItemFandPModifyIllegleParam
    End If
    
On Error GoTo Here
    szSql = " UPDATE price_formula_detail_lst SET " _
    & " item_formula='" & szItemFormu _
    & " ',parameter_1='" & sgPar1 & "'" _
    & " ,parameter_2='" & sgPar2 & "'" _
    & " ,parameter_3='" & sgPar3 & "'" _
    & " ,parameter_4='" & sgPar4 & "'" _
    & " ,parameter_5='" & sgPar5 & "'" _
    & "  WHERE price_Table_ID='" & PriceTableID & "' AND formula_name='" _
    & m_szFormulaName _
    & "' AND price_item='" _
    & szItemNum & "'"
    Dim lEffected As Long
    
    odb.ConnectionString = GetConnectionStr(cszPriceMan)
    odb.Execute szSql, lEffected
    If lEffected = 0 Then
        szSql = " INSERT price_formula_detail_lst (price_table_id,formula_name,price_item,item_formula,parameter_1,parameter_2,parameter_3,parameter_4,parameter_5) " _
        & " VALUES ( " _
        & TransFieldValueToString(PriceTableID) & "," _
        & TransFieldValueToString(PriceFormulaName) & "," _
        & TransFieldValueToString(szPriceItem) & "," _
        & TransFieldValueToString(szItemFormu) & "," _
        & TransFieldValueToString(sgPar1) & "," _
        & TransFieldValueToString(sgPar2) & "," _
        & TransFieldValueToString(sgPar3) & "," _
        & TransFieldValueToString(sgPar4) & "," _
        & TransFieldValueToString(sgPar5) & ")"
     odb.Execute szSql, lEffected
   
    End If
        
    WriteOperateLog m_oActiveUser, RIGHT_TPFModifyDB, "修改票价计算公式" & m_szFormulaName & " 的票价项" & szItemNum & " 的参数"
    
On Error GoTo 0
    Set odb = Nothing
Exit Function
Here:
    ShowError ERR_TPFCantUpdateFormulaList
    Set odb = Nothing
Exit Function
there:
    ShowError ERR_TPFItemFandPModifyIllegleParam
    Set odb = Nothing
End Function

Public Sub SetAsDefault()
    AssertCanCallSub
    AssertHaveRight m_oActiveUser, RIGHT_TPFSetAsDefault
On Error GoTo Here
    Dim szSql As String, odb As New RTConnection, lRowCount As Long
    szSql = "UPDATE Price_formula_info SET default_execute_mark=CASE WHEN formula_name <> '" _
    & m_szFormulaName & "' THEN 0 WHEN formula_name='" & m_szFormulaName & "'THEN 1 END"

    odb.ConnectionString = GetConnectionStr(cszPriceMan)
    odb.Execute szSql
    
    WriteOperateLog m_oActiveUser, RIGHT_TPFSetAsDefault, "把非缺省公式设为缺省或把缺省公式设为非缺省"
On Error GoTo 0
Exit Sub
Here:
    err.Raise err.Number, , err.Description
End Sub
