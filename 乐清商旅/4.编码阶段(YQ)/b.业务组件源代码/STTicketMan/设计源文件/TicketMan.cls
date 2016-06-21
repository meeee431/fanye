VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TicketMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private m_oActiveUser As ActiveUser

'Public Enum EErrSellerFinance
''    ERR_SellerFinanceNoActiveUser = ERR_SellerFinance + ERR_NoActiveUser
''    ERR_CompanyNotExist = ERR_SellerFinance + 12 '29212
'End Enum
'权限
'Public Enum ERightSellerFinance
''    RIGHT_QuerySellerEveryDaySellDetail = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 1 '查询售票员每日结算
''    RIGHT_ModifyCompanyName = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 2 '修改参运公司
''    RIGHT_GetBusStatByBusDate = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 3 '车次统计(按车次日期进行汇总)
''    RIGHT_GetBusStatBySaleTime = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 4 '按售票日期汇总,得到车次统计
''    RIGHT_GetBusStationTickets = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 5 '按售票车站,得到车次统计
''    RIGHT_GetBusStationTicketsCount = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 6 '按售票车站汇总,得到车次统计
''    RIGHT_GetBusTransStat = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 7 '得到指定条件的车次的运量
''    RIGHT_GetCombineBusSimply = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 8 '按照车次组合查询公司简报(按车次列出)
''    RIGHT_GetCombineBusSimplyByBusDate = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 9 '根据车次日期统计,指定车次组合内的车次的统计信息
''    RIGHT_GetStationTickets = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 10 '按照车次站点票统计信息
''    RIGHT_GetStationTicketsCount = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 11 '按照车次站点售票汇兑统计信息
''    'RIGHT_GetTotalItem = ERR_SellerFinance + cnMidRightBegin + cnMidRightStep * 12
'End Enum
Private m_nObjectStatus As EObjectStatus 'P2所有的实体类都有的私有类变量

Public Enum EErrTicketMan
    ERR_TicketNoDraw = 1001     '票号没有领用记录
    ERR_AddTicketMan = 1002     '票证不能新增
    ERR_ModifyTicketMan = 1003  '票证不能修改
    ERR_DeleteTicketMan = 1004  '票证不能删除
End Enum

'此处定义权限号：
Public Enum ERightTicketMan
    RIGHT_AddTicketMan = 1011       '票证新增
    RIGHT_ModifyTicketMan = 1017    '票证修改
    RIGHT_DeleteTicketMan = 1023    '票证删除
End Enum

'常规操作
Public Sub Init(poActiveUser As ActiveUser)
    Set SelfUser = poActiveUser
End Sub

Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
End Property

Public Property Set SelfUser(ByVal vNewValue As ActiveUser)
    Set m_oActiveUser = vNewValue
End Property
'**************************************************
'Member Code:S4
'Brief Description:测试对象是否有效
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Private Sub AssertObjIsValid()
    AssertActiveUserValid m_oActiveUser, ERR_Bus
'    AssertStatusAvailable
End Sub

'得到领票记录
Public Function GetTicketInfo(Optional ByVal pdtGetTicketStartDate As Date = cszEmptyDateStr, Optional pdtGetTicketEndDate As Date = cszForeverDateStr, Optional ByVal pszSellerID As String = "", Optional ByVal pszTicketNo As String = "") As Recordset

    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim nCount As Integer
    Dim szWhere As String
    Dim szOrder As String
    Dim i As Integer
    AssertObjIsValid
    oDB.ConnectionString = GetConnectionStr
    
    szWhere = ""
    If pszSellerID <> "" Then
        szWhere = szWhere & " AND  g.user_id Like '%" & pszSellerID & "%'"
        
    End If
    If m_oActiveUser.SellStationID <> "" Then
        szWhere = szWhere & " AND u.sell_station_id = " & TransFieldValueToString(m_oActiveUser.SellStationID)
    End If
    szSql = "SELECT * From getticket g , user_info u " _
            & " WHERE g.getticket_date  >= " & TransFieldValueToString(pdtGetTicketStartDate) _
            & " AND g.getticket_date< = " & TransFieldValueToString(pdtGetTicketEndDate) _
            & " AND g.user_id = u.user_id " _
            & szWhere & " ORDER BY getticket_date DESC,g.user_id"
     
    Set rsTemp = oDB.Execute(szSql)
    Set GetTicketInfo = rsTemp
    Set rsTemp = Nothing
    Set oDB = Nothing
End Function

'得到空白票记录
Public Function GetNullTicketInfo(Optional ByVal pdtGetNullTicketStartDate As Date = cszEmptyDateStr, Optional pdtGetNullTicketEndDate As Date = cszForeverDateStr, Optional ByVal pszSellerID As String = "", Optional ByVal pszTicketNo As String = "", Optional ByVal pszEndTicketNo As String = "") As Recordset

    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDB As New RTConnection
    Dim nTemp As Integer
    Dim nCount As Integer
    Dim szWhere As String
    Dim szOrder As String
    Dim i As Integer
    AssertObjIsValid
    oDB.ConnectionString = GetConnectionStr
    
    szWhere = ""
    If pszSellerID <> "" Then
        szWhere = szWhere & " AND  g.user_id Like '%" & pszSellerID & "%'"
    End If
    
    If m_oActiveUser.SellStationID <> "" Then
        szWhere = szWhere & " AND u.sell_station_id = " & TransFieldValueToString(m_oActiveUser.SellStationID)
    End If
    
    If pszTicketNo <> "" Then
        szWhere = szWhere & " AND  g.firstnullticket_no = " & TransFieldValueToString(pszTicketNo)
    End If
    
    If pszEndTicketNo <> "" Then
        szWhere = szWhere & " AND  g.lastnullticket_no = " & TransFieldValueToString(pszEndTicketNo)
    End If
    
    szSql = "SELECT * From getNullticket g , user_info u " _
            & " WHERE g.getnullticket_date  >= " & TransFieldValueToString(pdtGetNullTicketStartDate) _
            & " AND g.getnullticket_date< = " & TransFieldValueToString(pdtGetNullTicketEndDate) _
            & " AND g.user_id = u.user_id " _
            & szWhere & " ORDER BY getnullticket_date DESC,g.user_id"
     
    Set rsTemp = oDB.Execute(szSql)
    Set GetNullTicketInfo = rsTemp
    Set rsTemp = Nothing
    Set oDB = Nothing
End Function

'售票员领票
Public Function SellerGetTicket(prsSellGetTicketInfo As Recordset)
   On Error GoTo SellerGetTicketErr
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    Dim lEffect As Long
    AssertObjIsValid
    oDB.ConnectionString = GetConnectionStr
    
    AssertHaveRight m_oActiveUser, RIGHT_AddTicketMan

    If prsSellGetTicketInfo.RecordCount = 0 Then Exit Function
    prsSellGetTicketInfo.MoveFirst
    Dim oField As Field
    Dim nCount As Integer
    For Each oField In prsSellGetTicketInfo.Fields
        Select Case LCase(oField.Name)
            Case "user_id", "getticket_date", "user_name", "firstticket_no", "lastticket_no", "memo", "isuseing"
                nCount = nCount + 1
            Case Else
'                RaiseError ERR_ColumnName
        End Select
    Next
    If nCount <> 7 Then
'        RaiseError ERR_ColumnNum, cszProcName
    End If
    oDB.BeginTrans
        Do While Not prsSellGetTicketInfo.EOF

            szSql = " INSERT getticket ( " _
                & " user_id , " _
                & " getticket_date , " _
                & " user_name," _
                & " firstticket_no , " _
                & " lastticket_no , " _
                & " isuseing ," _
                & " memo ) "
                szSql = szSql & " VALUES ( " & TransFieldValueToString(prsSellGetTicketInfo!user_id) & "," _
                & TransFieldValueToString(prsSellGetTicketInfo!getticket_date) & "," _
                & TransFieldValueToString(prsSellGetTicketInfo!user_name) & "," _
                & TransFieldValueToString(prsSellGetTicketInfo!firstticket_no) & "," _
                & TransFieldValueToString(prsSellGetTicketInfo!lastticket_no) & "," _
                & TransFieldValueToString(FormatDbValue(prsSellGetTicketInfo!isuseing)) & "," _
                & TransFieldValueToString(prsSellGetTicketInfo!Memo) & ")"
                oDB.Execute szSql
            prsSellGetTicketInfo.MoveNext
        Loop
    oDB.CommitTrans
    Exit Function
SellerGetTicketErr:
    oDB.RollbackTrans
    Set oDB = Nothing
    Set rsTemp = Nothing
    err.Raise err.Number, err.Source, err.Description
End Function

'售票员空白票登记
Public Function SellerGetNullTicket(prsSellGetNullTicketInfo As Recordset)
   On Error GoTo SellerGetNullTicketErr
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    Dim lEffect As Long
    AssertObjIsValid
    oDB.ConnectionString = GetConnectionStr
    
    AssertHaveRight m_oActiveUser, RIGHT_AddTicketMan

    If prsSellGetNullTicketInfo.RecordCount = 0 Then Exit Function
    prsSellGetNullTicketInfo.MoveFirst
    Dim oField As Field
    Dim nCount As Integer
    For Each oField In prsSellGetNullTicketInfo.Fields
        Select Case LCase(oField.Name)
            Case "user_id", "getnullticket_date", "user_name", "firstnullticket_no", "lastnullticket_no", "operate_id", "operate_name", "memo"
                nCount = nCount + 1
            Case Else
'                RaiseError ERR_ColumnName
        End Select
    Next
    If nCount <> 7 Then
'        RaiseError ERR_ColumnNum, cszProcName
    End If
    oDB.BeginTrans
        Do While Not prsSellGetNullTicketInfo.EOF

            szSql = " INSERT getNullticket ( " _
                & " user_id , " _
                & " getnullticket_date , " _
                & " user_name," _
                & " firstnullticket_no , " _
                & " lastnullticket_no , " _
                & " operate_id , " _
                & " operate_name," _
                & " memo ) "
            szSql = szSql & " VALUES ( " & TransFieldValueToString(prsSellGetNullTicketInfo!user_id) & "," _
                & TransFieldValueToString(prsSellGetNullTicketInfo!getnullticket_date) & "," _
                & TransFieldValueToString(prsSellGetNullTicketInfo!user_name) & "," _
                & TransFieldValueToString(prsSellGetNullTicketInfo!firstnullticket_no) & "," _
                & TransFieldValueToString(prsSellGetNullTicketInfo!lastnullticket_no) & "," _
                & TransFieldValueToString(prsSellGetNullTicketInfo!operate_id) & "," _
                & TransFieldValueToString(prsSellGetNullTicketInfo!operate_name) & "," _
                & TransFieldValueToString(prsSellGetNullTicketInfo!Memo) & ")"
                oDB.Execute szSql
            prsSellGetNullTicketInfo.MoveNext
        Loop
    oDB.CommitTrans
    
    Exit Function
SellerGetNullTicketErr:
    oDB.RollbackTrans
    Set oDB = Nothing
    Set rsTemp = Nothing
    err.Raise err.Number, err.Source, err.Description
End Function

'返回最近的已经启用的领票记录(起始票号)
Public Function GetFirstTicketasMaxDate(pszSellerID As String, Optional bIsStartNewTicket As Boolean = False) As String

Dim szFirstTicketasMaxDate As String
Dim szSql As String
Dim rsTemp As New Recordset
Dim rsTmp As New Recordset
Dim oDB As New RTConnection
AssertObjIsValid
oDB.ConnectionString = GetConnectionStr
'判断该用户是否存在领票记录，如果没有否退出
szSql = "SELECT * FROM getticket WHERE user_id=" & Trim(pszSellerID)
Set rsTmp = oDB.Execute(szSql)
If rsTmp.RecordCount = 0 Then Exit Function
'判断该用户是否存在已经使用的领票记录,如果没有则返加空值( 可能有问题)
szSql = "SELECT * FROM getticket WHERE user_id =" & TransFieldValueToString(pszSellerID) _
       & " AND isuseing=1 " _
       & " ORDER BY getticket_date DESC"
Set rsTemp = oDB.Execute(szSql)
If rsTemp.RecordCount = 0 Then GetFirstTicketasMaxDate = "": Exit Function

'返回最近已经启用的最近的领票记录
If rsTemp.RecordCount > 0 Then
    szSql = " SELECT top 1 firstticket_no FROM getticket " _
       & " WHERE user_id =" & TransFieldValueToString(pszSellerID) _
       & " AND isuseing=1 " _
       & " ORDER BY getticket_date DESC"
    szFirstTicketasMaxDate = oDB.Execute(szSql).Fields("firstticket_no")
    If szFirstTicketasMaxDate <> "" Then
        GetFirstTicketasMaxDate = szFirstTicketasMaxDate
    Else
        GetFirstTicketasMaxDate = ""
    End If
End If
End Function


'返回最近的领票记录(结束票号)
Public Function GetLastTicketasMaxDate(pszSellerID As String) As String

Dim szLastTicketasMaxDate As String
Dim szSql As String
Dim oDB As New RTConnection
Dim rsTmp As New Recordset
Dim rsTemp As New Recordset
AssertObjIsValid
oDB.ConnectionString = GetConnectionStr
'判断该用户是否存在领票记录，如果没有否退出
szSql = "SELECT * FROM getticket WHERE user_id=" & Trim(pszSellerID)
Set rsTmp = oDB.Execute(szSql)
If rsTmp.RecordCount = 0 Then Exit Function
'判断该用户是否存在已经使用的领票记录,如果没有则返加空值( 可能有问题)
szSql = "SELECT * FROM getticket WHERE user_id =" & TransFieldValueToString(pszSellerID) _
       & " AND isuseing=1 " _
       & " ORDER BY getticket_date DESC"
Set rsTemp = oDB.Execute(szSql)
If rsTemp.RecordCount = 0 Then GetLastTicketasMaxDate = "": Exit Function
'返回最近已经启用的最近的领票记录
szSql = " SELECT top 1 lastticket_no FROM getticket WHERE user_id =" & TransFieldValueToString(pszSellerID) _
        & " AND isuseing=1 " _
        & " ORDER BY getticket_date DESC"
szLastTicketasMaxDate = oDB.Execute(szSql).Fields("lastticket_no")
If szLastTicketasMaxDate <> "" Then
    GetLastTicketasMaxDate = szLastTicketasMaxDate
Else
    GetLastTicketasMaxDate = ""
End If
End Function


'售票员领票
Public Function SellerUpdateGetTicket(prsSellGetTicketInfo As Recordset) As Boolean
   On Error GoTo SellerGetTicketErr

    Dim szSql As String
    Dim rsTemp As Recordset
    Dim rsTmp As New Recordset
    Dim oDB As New RTConnection
    Dim lEffect As Long
    AssertObjIsValid
    oDB.ConnectionString = GetConnectionStr
    
    AssertHaveRight m_oActiveUser, RIGHT_ModifyTicketMan
    
    If prsSellGetTicketInfo.RecordCount = 0 Then Exit Function
    szSql = "SELECT * FROM getticket WHERE user_id= " & TransFieldValueToString(FormatDbValue(prsSellGetTicketInfo!user_id)) _
            & "AND  getticket_date =" & TransFieldValueToString(FormatDbValue(prsSellGetTicketInfo!getticket_date))
    Set rsTmp = oDB.Execute(szSql)
    If rsTmp.RecordCount = 0 Then SellerUpdateGetTicket = False: Exit Function
    prsSellGetTicketInfo.MoveFirst
    Dim oField As Field
    Dim nCount As Integer
    For Each oField In prsSellGetTicketInfo.Fields
        Select Case LCase(oField.Name)
            Case "user_id", "getticket_date", "user_name", "firstticket_no", "lastticket_no", "memo", "isuseing"
                nCount = nCount + 1
            Case Else
'                RaiseError ERR_ColumnName
        End Select
    Next
    oDB.BeginTrans
        Do While Not prsSellGetTicketInfo.EOF

            szSql = " UPDATE getticket SET user_id= " & TransFieldValueToString(FormatDbValue(prsSellGetTicketInfo!user_id)) & "," _
                & " getticket_date = " & TransFieldValueToString(FormatDbValue(prsSellGetTicketInfo!getticket_date)) & "," _
                & " user_name=" & TransFieldValueToString(FormatDbValue(prsSellGetTicketInfo!user_name)) & "," _
                & " firstticket_no = " & TransFieldValueToString(FormatDbValue(prsSellGetTicketInfo!firstticket_no)) & "," _
                & " lastticket_no = " & TransFieldValueToString(FormatDbValue(prsSellGetTicketInfo!lastticket_no)) & "," _
                & " isuseing = " & TransFieldValueToString(FormatDbValue(prsSellGetTicketInfo!isuseing)) & "," _
                & " memo =" & TransFieldValueToString(FormatDbValue(prsSellGetTicketInfo!Memo)) _
                & " WHERE user_id =" & TransFieldValueToString(FormatDbValue(prsSellGetTicketInfo!user_id)) _
                & " AND getticket_date= " & TransFieldValueToString(FormatDbValue(prsSellGetTicketInfo!getticket_date))
            oDB.Execute szSql
            prsSellGetTicketInfo.MoveNext
        Loop
    oDB.CommitTrans
    SellerUpdateGetTicket = True
    Exit Function
SellerGetTicketErr:
    oDB.RollbackTrans
    Set oDB = Nothing
    Set rsTemp = Nothing
    SellerUpdateGetTicket = False
    err.Raise err.Number, err.Source, err.Description
End Function

'售票员空白票
Public Function SellerUpdateGetNullTicket(prsSellGetNullTicketInfo As Recordset, szStartNo As String, szEndNo As String) As Boolean
   On Error GoTo SellerUpdateGetNullTicketErr

    Dim szSql As String
    Dim rsTemp As Recordset
    Dim rsTmp As New Recordset
    Dim oDB As New RTConnection
    Dim lEffect As Long
    AssertObjIsValid
    oDB.ConnectionString = GetConnectionStr
    
    AssertHaveRight m_oActiveUser, RIGHT_ModifyTicketMan
    
    If prsSellGetNullTicketInfo.RecordCount = 0 Then Exit Function
    szSql = "SELECT * FROM getNullticket WHERE user_id= " & TransFieldValueToString(FormatDbValue(prsSellGetNullTicketInfo!user_id)) _
            & " AND getnullticket_date =" & TransFieldValueToString(FormatDbValue(prsSellGetNullTicketInfo!getnullticket_date)) _
            & " AND firstnullticket_no =" & TransFieldValueToString(szStartNo) _
            & " AND lastnullticket_no =" & TransFieldValueToString(szEndNo)
    Set rsTmp = oDB.Execute(szSql)
    If rsTmp.RecordCount = 0 Then SellerUpdateGetNullTicket = False: Exit Function
    prsSellGetNullTicketInfo.MoveFirst
    Dim oField As Field
    Dim nCount As Integer
    For Each oField In prsSellGetNullTicketInfo.Fields
        Select Case LCase(oField.Name)
            Case "user_id", "getnullticket_date", "user_name", "firstnullticket_no", "lastnullticket_no", "operate_id", "operate_name", "memo"
                nCount = nCount + 1
            Case Else
'                RaiseError ERR_ColumnName
        End Select
    Next
    oDB.BeginTrans
        Do While Not prsSellGetNullTicketInfo.EOF

            szSql = " UPDATE getNullticket SET user_id= " & TransFieldValueToString(FormatDbValue(prsSellGetNullTicketInfo!user_id)) & "," _
                & " getnullticket_date = " & TransFieldValueToString(FormatDbValue(prsSellGetNullTicketInfo!getnullticket_date)) & "," _
                & " user_name=" & TransFieldValueToString(FormatDbValue(prsSellGetNullTicketInfo!user_name)) & "," _
                & " firstnullticket_no = " & TransFieldValueToString(FormatDbValue(prsSellGetNullTicketInfo!firstnullticket_no)) & "," _
                & " lastnullticket_no = " & TransFieldValueToString(FormatDbValue(prsSellGetNullTicketInfo!lastnullticket_no)) & "," _
                & " operate_id=" & TransFieldValueToString(FormatDbValue(prsSellGetNullTicketInfo!operate_id)) & "," _
                & " operate_name=" & TransFieldValueToString(FormatDbValue(prsSellGetNullTicketInfo!operate_name)) & "," _
                & " memo =" & TransFieldValueToString(FormatDbValue(prsSellGetNullTicketInfo!Memo)) _
                & " WHERE user_id =" & TransFieldValueToString(FormatDbValue(prsSellGetNullTicketInfo!user_id)) _
                & " AND getnullticket_date= " & TransFieldValueToString(FormatDbValue(prsSellGetNullTicketInfo!getnullticket_date)) _
                & " AND firstnullticket_no= " & TransFieldValueToString(szStartNo) _
                & " AND lastnullticket_no= " & TransFieldValueToString(szEndNo)
            oDB.Execute szSql
            prsSellGetNullTicketInfo.MoveNext
        Loop
    oDB.CommitTrans
    SellerUpdateGetNullTicket = True
    
    Exit Function
SellerUpdateGetNullTicketErr:
    oDB.RollbackTrans
    Set oDB = Nothing
    Set rsTemp = Nothing
    SellerUpdateGetNullTicket = False
    err.Raise err.Number, err.Source, err.Description
End Function

'删除领票记录
Public Function DeleteTicketMan(pszSeller As String, pdtGetTicketDate As Date) As Boolean
   On Error GoTo DeleteTicketManerr
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim rsTmp As New Recordset
    Dim oDB As New RTConnection
    Dim szMaxGetTicketDate As String
    Dim lEffect As Long
    Dim bIsUseing As Boolean
    AssertObjIsValid
    oDB.ConnectionString = GetConnectionStr
    
    AssertHaveRight m_oActiveUser, RIGHT_DeleteTicketMan
    
    '如果记录不存在
    szSql = "SELECT * FROM getticket WHERE user_id= " & TransFieldValueToString(pszSeller) _
            & "AND  getticket_date =" & TransFieldValueToString(pdtGetTicketDate)
    Set rsTmp = oDB.Execute(szSql)
    If rsTmp.RecordCount = 0 Then DeleteTicketMan = False: Exit Function

    szSql = " SELECT  MAX(getticket_date) getticket_date FROM getticket WHERE user_id= " & TransFieldValueToString(pszSeller) _
            & "AND  isuseing=1"
    szMaxGetTicketDate = IIf(IsNull(oDB.Execute(szSql).Fields("getticket_date")), cszEmptyDateStr, oDB.Execute(szSql).Fields("getticket_date"))
    If CDate(szMaxGetTicketDate) = pdtGetTicketDate Then
        MsgBox "当前领票记录不能删除，因为当前领票记录正在使用中！", vbOKOnly + vbExclamation, "错误"
        DeleteTicketMan = False
        Exit Function
    End If


    szSql = "DELETE getticket WHERE user_id=" & TransFieldValueToString(pszSeller) _
            & "AND getticket_date = " & TransFieldValueToString(pdtGetTicketDate)
    oDB.Execute szSql
    DeleteTicketMan = True
    
    Exit Function
DeleteTicketManerr:
    Set rsTemp = Nothing
    Set oDB = Nothing
    DeleteTicketMan = False
    err.Raise err.Number, err.Source, err.Description
End Function

'删除空白票记录
Public Function DeleteNullTicketMan(pszSeller As String, pdtGetNullTicketDate As Date, szStartNo As String, szEndNo As String) As Boolean
   On Error GoTo DeleteNullTicketManerr
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim rsTmp As New Recordset
    Dim oDB As New RTConnection
    Dim szMaxGetNullTicketDate As String
    Dim lEffect As Long
    Dim bIsUseing As Boolean
    AssertObjIsValid
    oDB.ConnectionString = GetConnectionStr
    
    AssertHaveRight m_oActiveUser, RIGHT_DeleteTicketMan
    
    '如果记录不存在
    szSql = "SELECT * FROM getNullticket WHERE user_id= " & TransFieldValueToString(pszSeller) _
            & " AND getnullticket_date =" & TransFieldValueToString(pdtGetNullTicketDate) _
            & " AND firstnullticket_no =" & TransFieldValueToString(szStartNo) _
            & " AND lastnullticket_no =" & TransFieldValueToString(szEndNo)
    Set rsTmp = oDB.Execute(szSql)
    If rsTmp.RecordCount = 0 Then DeleteNullTicketMan = False: Exit Function

    szSql = "DELETE FROM getNullticket WHERE user_id=" & TransFieldValueToString(pszSeller) _
            & "AND getnullticket_date = " & TransFieldValueToString(pdtGetNullTicketDate) _
            & " AND firstnullticket_no =" & TransFieldValueToString(szStartNo) _
            & " AND lastnullticket_no =" & TransFieldValueToString(szEndNo)
    oDB.Execute szSql
    DeleteNullTicketMan = True
    
    Exit Function
DeleteNullTicketManerr:
    Set rsTemp = Nothing
    Set oDB = Nothing
    DeleteNullTicketMan = False
    err.Raise err.Number, err.Source, err.Description
End Function

'判断该售票员是否有新的未被使用的领票记录，如果有的设计isuseing状态为1，否则返回False
Public Function IsNewTicketManRecord(pszSellerID As String) As Boolean
On Error GoTo IsNewTicketManRecordErr
Dim oDB As New RTConnection
Dim szSql As String
Dim rsTmp As New Recordset
AssertObjIsValid
oDB.ConnectionString = GetConnectionStr


szSql = "SELECT * FROM getticket WHERE getticket_date =" _
        & "(SELECT Max(getticket_date) FROM  getticket WHERE user_id=" & TransFieldValueToString(pszSellerID) & ")" _
        & " AND isuseing =0 "
    
Set rsTmp = oDB.Execute(szSql)
'如果没有新的未被作用的记录，则返回False
If rsTmp.RecordCount = 0 Then IsNewTicketManRecord = False: Exit Function
If rsTmp.RecordCount > 1 Then IsNewTicketManRecord = False: Exit Function

szSql = " UPDATE getticket SET isuseing =1 " _
        & "WHERE getticket_date= " _
        & " (SELECT Max(getticket_date) FROM  getticket WHERE user_id=" & TransFieldValueToString(pszSellerID) & " )AND isuseing =0 "
oDB.Execute szSql
IsNewTicketManRecord = True
Exit Function
IsNewTicketManRecordErr:
    IsNewTicketManRecord = False
End Function


'修改提示张数
Public Function upDateTicketParm(szParmEnglisName As String, szParmChinseName As String, iParmValues As Integer, szMemo As String) As Boolean
On Error GoTo upDateTicketParmErr
Dim oDB As New RTConnection
Dim szSql As String
AssertObjIsValid
oDB.ConnectionString = GetConnectionStr
szSql = "UPDATE System_param_info " _
        & "SET parameter_chinese_name='" & szParmChinseName _
        & "',parameter_value='" & iParmValues _
        & "',annotation='" & szMemo _
        & "' WHERE parameter_name='" & szParmEnglisName & "'"
oDB.Execute szSql
upDateTicketParm = True
Exit Function
upDateTicketParmErr:
upDateTicketParm = False
End Function

'得到系统参数
Public Function GetParam(ByVal pszParamName As String) As Integer
    
    Dim piTemp As TParamInfo
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    AssertObjIsValid
    oDB.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "SELECT * FROM System_param_info WHERE parameter_name='" & FormatStr(pszParamName) & "'"
    Set rsTemp = oDB.Execute(szSql)
    
    If rsTemp.RecordCount = 0 Then GetParam = 0:  Exit Function

    GetParam = FormatDbValue(rsTemp!parameter_value)
    Set rsTemp = Nothing
    Set oDB = Nothing
End Function

''启用票据
''参数为起始票号,结束票号
''如果票据有领用记录，则将该记录设为已启用。否则提示票据未领用
Public Function StartUserSheet(pszStartSheetID As String, pszEndSheetID As String) As Boolean

   On Error GoTo err

    Dim szSql As String
    Dim rsTemp As New Recordset
    Dim poDb As New RTConnection
    
    poDb.ConnectionString = GetConnectionStr

    poDb.BeginTrans

    Set rsTemp = poDb.Execute("SELECT * from getticket " _
                & " WHERE user_id = '" & m_oActiveUser.UserID & "' " _
                & " AND firstticket_no <='" & pszStartSheetID & "' " _
                & " AND lastticket_no >= '" & pszEndSheetID & "' ")

    
    If rsTemp.RecordCount = 0 Then StartUserSheet = False: Exit Function 'ShowError ERR_TicketNoDraw
    
    '将该用户领用的记录全部标记为未启用
    poDb.Execute ("update getticket set isuseing = 0 " _
                & " WHERE user_id = '" & m_oActiveUser.UserID & "'")
                
    '将指定的票号段标记为已启用
    poDb.Execute ("update getticket set isuseing = 1 " _
                & " WHERE user_id = '" & m_oActiveUser.UserID & "' " _
                & " AND firstticket_no <='" & pszStartSheetID & "' " _
                & " AND lastticket_no >= '" & pszEndSheetID & "' ")

    poDb.CommitTrans
    StartUserSheet = True
    
    
    Exit Function
err:
    poDb.RollbackTrans
    Set poDb = Nothing
    StartUserSheet = False

    err.Raise err.Number

End Function
