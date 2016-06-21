VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TicketPriceMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As String, ByVal Source As String, ByVal Length As Long)

Public Event SetProgressRange(ByVal lRange)
Public Event SetProgressValue(ByVal lValue)
Public Event SetCopyBusPriceStatus(ByVal lStatus As String)
Private m_CancelCopyBusPrice As Boolean '取消复制车次票价

'保持属性值的局部变量
Private m_oActiveUser As ActiveUser '局部复制
'---------------------
'所有的实体类都有的私有类变量<<
Private m_nObjectStatus As EObjectStatus


'尾数处理方法
Public Type TTailDealMethod
    szMethodNo As String
    szSerialNo As String
    nStartNumber As Integer
    nEndNumber As Integer
    nDealedNumber As Integer
    szAnnotation As String
End Type

'区域尾数处理采用
Public Type TAreaTailDealMethod
    szAreaCode As String
    szAreaName As String
    szBusType As String
    szPriceItemCode As String
    szPriceItemName As String
    nTailBitNo As Integer
    szTailBitNoName As String
    szMethodNo As String
End Type

Const szCentMili = "厘"
Const szCent = "分"
Const szTenCent = "角"
Const szTotalPrice = "票价总和"
Const nTotalPrice = 0
Const nDetailPrice = 1

Const nAllBusType = 100

Const szAreaCode = "szAreaCode"
Const szBusType = "szBusType"
Const szTailBitNo = "nTailBitNo"
Const szPriceItem = "szPriceItem"
'区域
'****此类型可以删除
Public Type TArea
    szAreaCode As String
    szAreaName As String
    szProvinceInOut As Integer
End Type

'处理值
Public Type TDealValue
    pbTotalCol As Boolean '是否总票价列
    sgValue As Single     '处理值
    sgLastColValue As Single '舍入项值
End Type

Public Enum EErrTailDeal
    ERR_TailDealMethodDelete = ERR_RouteTicketPrice + 37 '该尾数处理方法已经在是使用，不能删除！
    ERR_TailDealMethodSave = ERR_RouteTicketPrice + 38 '保存尾数处理错误
    ERR_GetDetailKeepBit = ERR_RouteTicketPrice + 39 '系统参数“PriceDetailKeepBit”不存在
    ERR_DeleteAreaMethod = ERR_RouteTicketPrice + 40 '删除区域尾数处理错误
    ERR_AddAreaMethod = ERR_RouteTicketPrice + 41 '增加区域尾数处理错误
    ERR_UpdateAreaMethod = ERR_RouteTicketPrice + 42 '更新区域尾数处理错误
    ERR_TailDeal = ERR_RouteTicketPrice + 43 '尾数处理错误
    ERR_TPFNotAvailable = ERR_TicketPriceFormula + ERR_NotAvailable '对象无效状态
End Enum
    
Public Enum ERightTailDeal
    RIGHT_TailDealMethodDelete = ERR_RouteTicketPrice + cnMidRightBegin + cnMidRightStep * 6 '删除尾数处理方法权限
    RIGHT_TailDealMethodSave = ERR_RouteTicketPrice + cnMidRightBegin + cnMidRightStep * 7 '保存尾数处理方法权限
    RIGHT_DeleteAreaMethod = ERR_RouteTicketPrice + cnMidRightBegin + cnMidRightStep * 8 '删除区域尾数处理方法权限
    RIGHT_AddAreaMethod = ERR_RouteTicketPrice + cnMidRightBegin + cnMidRightStep * 9 '增加区域尾数处理方法权限
    RIGHT_UpdateAreaMethod = ERR_RouteTicketPrice + cnMidRightBegin + cnMidRightStep * 10 '更新区域尾数处理方法权限
End Enum

Public Property Get CancelCopyBusPrice() As Boolean
    CancelCopyBusPrice = m_CancelCopyBusPrice
End Property

Public Property Let CancelCopyBusPrice(ByVal vNewValue As Boolean)
   m_CancelCopyBusPrice = vNewValue
End Property


Private Sub AssertObjIsValid() '测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
    AssertActiveUserValid m_oActiveUser, ERR_TicketPriceFormula
'    AssertStatusAvailable
End Sub

Private Sub AssertStatusAvailable() '测试对象的状态是否有效（不为无效ECNotAvailable状态）
    If m_nObjectStatus = ST_NotAvailable Then ShowError ERR_TPFNotAvailable
End Sub

Public Property Get ObjStatus() As EObjectStatus
    ObjStatus = m_nObjectStatus
End Property
Public Property Let ObjStatus(vNewValue As EObjectStatus)
    m_nObjectStatus = vNewValue
End Property

Public Property Set SelfUser(ByVal vData As ActiveUser)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.SelfUser = Form1
    Set m_oActiveUser = vData
End Property

'GetAllTicketPriceFormula
Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
End Property

Public Sub Init(AUser As ActiveUser)
    Set SelfUser = AUser
End Sub

Public Function GetAllTicketPriceFormula() As String()
    Dim odb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim aszTemp() As String
    
    AssertActiveUserValid m_oActiveUser, ERR_TicketPriceMan
    
    odb.ConnectionString = GetConnectionStr(cszSNPrice)
    szSql = "SELECT * FROM Price_formula_info"
    Set rsTemp = odb.Execute(szSql)
    
    If rsTemp.RecordCount > 0 Then
        ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 3)
        For i = 1 To rsTemp.RecordCount
            aszTemp(i, 1) = FormatDbValue(rsTemp!formula_name)
            aszTemp(i, 2) = FormatDbValue(rsTemp!default_execute_mark)
            aszTemp(i, 3) = FormatDbValue(rsTemp!Annotation)
            rsTemp.MoveNext
        Next
  
    
    End If
    GetAllTicketPriceFormula = aszTemp
    Set rsTemp = Nothing
    Set odb = Nothing
End Function

'得到指定车次计划的所有线路票价表和票价表名
Public Function GetAllRoutePriceTable() As String()
    Dim odb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim aszTemp() As String

    AssertActiveUserValid m_oActiveUser, ERR_TicketPriceMan

    odb.ConnectionString = GetConnectionStr(cszSNPrice)

''    If Trim(pszProjectID) <> "" Then
'       szSql = "SELECT * FROM price_table_info WHERE project_id='" & pszProjectID & "'"
'    Else
       szSql = "SELECT * FROM price_table_info "
'    End If

    Set rsTemp = odb.Execute(szSql)

    If rsTemp.RecordCount > 0 Then
        ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 6)
        For i = 1 To rsTemp.RecordCount
            aszTemp(i, 1) = FormatDbValue(rsTemp!price_table_id)
            aszTemp(i, 2) = FormatDbValue(rsTemp!price_table_name)
            aszTemp(i, 3) = FormatDbValue(rsTemp!start_run_time)
            aszTemp(i, 4) = FormatDbValue(rsTemp!create_time)
            aszTemp(i, 5) = FormatDbValue(rsTemp!last_modify_time)
'            aszTemp(i, 6) = FormatDbValue(rsTemp!project_id)
            rsTemp.MoveNext
        Next
    End If
    GetAllRoutePriceTable = aszTemp
    Set rsTemp = Nothing
    Set odb = Nothing
End Function


'复制票价表
'此处需修改一下,改为一句SQL
'2002-11-19陈峰

Public Sub CopyPriceTable(SouTable As String, DesTable As String, BusID() As String)

    Dim odb As New RTConnection
    Dim szSql As String
    Dim BusCount As Integer
    Dim i As Integer
    
    On Error GoTo ErrHandle
    AssertActiveUserValid m_oActiveUser, ERR_TicketPriceMan
    odb.ConnectionString = GetConnectionStr(cszSNPrice)
    
    BusCount = ArrayLength(BusID)
    RaiseEvent SetProgressRange(BusCount)
    
    Dim szBusIDs As String
    For i = 1 To BusCount
        If i = 1 Then szBusIDs = " AND bus_id IN ("
        If i = BusCount Then
            szBusIDs = szBusIDs & TransFieldValueToString(BusID(i)) & ")"
        Else
            szBusIDs = szBusIDs & TransFieldValueToString(BusID(i)) & ","
        End If
    Next i
        
    odb.BeginTrans
    szSql = "DELETE bus_price_lst WHERE price_table_id=" & TransFieldValueToString(DesTable) & szBusIDs
    odb.Execute szSql
    szSql = "INSERT bus_price_lst (sell_station_id,bus_id,station_id, price_table_id,ticket_type,vehicle_type_code,seat_type_id,station_serial_no," _
            & " mileage,base_carriage,price_item_1,price_item_2,price_item_3,price_item_4,price_item_5,price_item_6, " _
            & " price_item_7,price_item_8,price_item_9,price_item_10,price_item_11,price_item_12,price_item_13, " _
            & "price_item_14,price_item_15 )" _
            & " SELECT sell_station_id,bus_id,station_id," & TransFieldValueToString(DesTable) & ",ticket_type,vehicle_type_code,seat_type_id,station_serial_no," _
            & " mileage,base_carriage,price_item_1,price_item_2,price_item_3,price_item_4,price_item_5,price_item_6, " _
            & " price_item_7,price_item_8,price_item_9,price_item_10,price_item_11,price_item_12,price_item_13, " _
            & "price_item_14,price_item_15 " _
            & " FROM bus_price_lst WHERE price_table_id=" & TransFieldValueToString(SouTable) & szBusIDs
    odb.Execute szSql
    odb.CommitTrans
    Set odb = Nothing
    Exit Sub
ErrHandle:
    odb.RollbackTrans
    Set odb = Nothing
    RaiseError err.Number, , err.Description
    


'    Dim odb As New RTConnection
'    Dim szSql As String
'    Dim rsTemp As Recordset
'    Dim j As Long
'    Dim szTemp As String
'    Dim oRoutePriceTable As New RoutePriceTable
'    Dim BusCount As Integer
'    Dim szBusID As String
'    Dim szSqlUpdate As String
'    Dim i As Integer
'    Dim szCopyBusPriceStatus  As String
'
'    On Error GoTo err
'    AssertActiveUserValid m_oActiveUser, ERR_TicketPriceMan
'    odb.ConnectionString = GetConnectionStr(cszSNPrice)
'
'    BusCount = ArrayLength(BusID)
'    RaiseEvent SetProgressRange(BusCount)
'
'
'    oRoutePriceTable.Init m_oActiveUser
'    oRoutePriceTable.Identify DesTable
'
'    For i = 1 To BusCount
'nextBus:
'        On Error GoTo errorCopy
'        szSql = "DELETE bus_price_lst WHERE price_table_id='" & DesTable & "' AND  bus_id=" & "'" & BusID(i) & "'"
'        odb.BeginTrans
'        odb.Execute szSql 'odb.Execute szSql, , , 0
'        szSql = "SELECT bus_id,station_id" & ", '" & DesTable & "',ticket_type,vehicle_type_code,seat_type_id,station_serial_no,mileage,base_carriage"
'        For j = 1 To 15
'            szSql = szSql & ",price_item_" & j
'        Next
'        szSql = szSql & " FROM bus_price_lst WHERE price_table_id='" & SouTable & "' AND  bus_id=" & "'" & BusID(i) & "'"
'
'        szSql = "INSERT  into bus_price_lst(sell_station_id,bus_id,station_id, price_table_id,ticket_type,vehicle_type_code,seat_type_id,station_serial_no," _
'            & " mileage,base_carriage,price_item_1,price_item_2,price_item_3,price_item_4,price_item_5,price_item_6, " _
'            & " price_item_7,price_item_8,price_item_9,price_item_10,price_item_11,price_item_12,price_item_13, " _
'            & "price_item_14,price_item_15 ) " & szSql
'        odb.Execute szSql 'odb.Execute szSql, , , 0
''        odb.CommitTrans
'        '*******尾数处理
'        'oRoutePriceTable.DealTableTail DesTable, BusID(i)
'        szCopyBusPriceStatus = "[" & BusID(i) & "]" & "成功"
'        RaiseEvent SetCopyBusPriceStatus(szCopyBusPriceStatus)
'        RaiseEvent SetProgressValue(i)
'        DoEvents
'    Next
'
'    Set oRoutePriceTable = Nothing
'    Set odb = Nothing
'    Exit Sub
'
'errorCopy:
'    szCopyBusPriceStatus = "[" & BusID(i) & "]" & "失败"
'    szSql = "DElETE FROM bus_price_lst WHERE bus_id='" & BusID(i) & "' AND price_table_id='" & DesTable & "'"
'    RaiseEvent SetCopyBusPriceStatus(szCopyBusPriceStatus)
'    RaiseEvent SetProgressValue(i)
'    i = i + 1
'    If i <= BusCount Then GoTo nextBus
'err:
''    odb.RollbackTrans
'    ShowError ERR_CopyRoutePriceTable
'    Set rsTemp = Nothing
'    Set odb = Nothing
End Sub

'得到正在用的票价计算公式库中的公式信息
Public Function GetPriceItemFormulaInfo() As TPriceFormulaInfo()
    Dim oCaculator As Object
    Set oCaculator = CreateObject("STPriLib.PriceItemFunLib")
    oCaculator.Init m_oActiveUser
    GetPriceItemFormulaInfo = oCaculator.FormulaInfo
    Set oCaculator = Nothing
End Function

'验证参数的有效性
Public Function AssertPriceItemParamIsValid(FormulaName As String, Param1 As String, Param2 As String, Param3 As String, Param4 As String, Param5 As String) As Long
    Dim oCaculator As Object
    Set oCaculator = CreateObject("STPriLib.PriceItemFunLib")
    oCaculator.Init m_oActiveUser
    On Error GoTo Error_Handle
    CallByName oCaculator, FormulaName, VbMethod, Param1, Param2, Param3, Param4, Param5
    Set oCaculator = Nothing
    Exit Function
Error_Handle:
    Dim szTemp As String
    szTemp = oCaculator.GetLastError()
    err.Raise err.Number, , szTemp
    Set oCaculator = Nothing
End Function

Private Sub CopyStr(pszDes As String, pszSource As String)
    Dim nCount As Integer
    nCount = LenB(StrConv(pszSource, vbFromUnicode))
    If nCount > 0 Then
        pszDes = Space(nCount)
        CopyMemory pszDes, pszSource, nCount
    End If
End Sub


Public Function GetDetailKeepBit() As Integer
Dim odb As New RTConnection
Dim orsTemp As Recordset
On Error GoTo errorHander

odb.ConnectionString = GetConnectionStr(cszPriceMan)
Set orsTemp = odb.Execute("SELECT * FROM System_param_info WHERE parameter_name='PriceDetailKeepBit'")

If orsTemp.RecordCount = 1 Then
   GetDetailKeepBit = orsTemp!parameter_value
Else
   GoTo errorHander
End If
Set orsTemp = Nothing
Set odb = Nothing
Exit Function

errorHander:
    ShowError ERR_GetDetailKeepBit
End Function

'尾数处理过程
Public Function DetailTailDeal(vNewValue As Single, szAreaCode As String, szBusType As String, szPriceItem As String, ttTailDeal() As TAreaTailDealMethod, ttTailDealMethod() As TTailDealMethod) As Single
    Dim TempTailDeal(1 To 3) As TAreaTailDealMethod
    Dim TempTailDealMethod() As TTailDealMethod

    Dim szTemp As String
    Dim nTailMethod As Integer
    Dim nStation As Long
    Dim m, l As Long

    Dim nIntPart As Long '整数部分
    Dim nDecimalPart As Integer '小数部分
    Dim nDecimalDigit1 As Long '小数后第一位
    Dim nDecimalDigit2 As Integer '小数后第二位
    Dim nDecimalDigit3 As Integer '小数后第三位
    Dim nFirstCol As Long
    
    Dim nHalf As Long
    Dim nStart As Long
    Dim nEnd As Long
    Dim FindNo As Long
    Dim n As Long
    
    On Error GoTo errorHander
    
    vNewValue = Round(vNewValue, 3)
    nStation = ArrayLength(ttTailDeal)
    nTailMethod = ArrayLength(ttTailDealMethod)
    
    For l = 1 To 3
        n = FindTailDeal(ttTailDeal, szAreaCode, szBusType, Str(l), szPriceItem) '查找该区域的处理方法

         If n > 0 Then TempTailDeal(l) = ttTailDeal(n)
    Next
    
    nIntPart = Int(vNewValue)
    nDecimalPart = Round((vNewValue - Int(vNewValue)) * 1000, 0)
    nDecimalDigit1 = Int(nDecimalPart / 100)
    nDecimalDigit2 = Int((nDecimalPart - nDecimalDigit1 * 100) / 10)
    nDecimalDigit3 = nDecimalPart Mod 10
    
    For l = 1 To nTailMethod
        If TempTailDeal(3).szMethodNo <> "" Then
            If TempTailDeal(3).szMethodNo = ttTailDealMethod(l).szMethodNo Then
                If nDecimalDigit3 >= ttTailDealMethod(l).nStartNumber And nDecimalDigit3 <= ttTailDealMethod(l).nEndNumber Then
                    If ttTailDealMethod(l).nDealedNumber <> 100 Then
                        nDecimalDigit3 = ttTailDealMethod(l).nDealedNumber
                        If nDecimalDigit3 = 10 Then
                            nDecimalDigit3 = 0
                            nDecimalDigit2 = nDecimalDigit2 + 1
                        End If
                        If nDecimalDigit2 = 10 Then
                            nDecimalDigit2 = 0
                            nDecimalDigit1 = nDecimalDigit1 + 1
                        End If
                        If nDecimalDigit1 = 10 Then
                            nDecimalDigit1 = 0
                            nIntPart = nIntPart + 1
                        End If
                        Exit For
                    End If
                End If
            End If
        Else '直接四舍五入，对小数点后一、二位不处理
            If nDecimalDigit3 < 5 Then
                nDecimalDigit3 = 0
            Else
               nDecimalDigit3 = 0
               nDecimalDigit2 = nDecimalDigit2 + 1
            End If
            If nDecimalDigit2 >= 10 Then
               nDecimalDigit2 = 0
               nDecimalDigit1 = nDecimalDigit1 + 1
            End If
            If nDecimalDigit1 >= 10 Then
               nDecimalDigit1 = 0
               nIntPart = nIntPart + 1
            End If
            Exit For
        End If
    Next
    
    For l = 1 To nTailMethod
        If TempTailDeal(2).szMethodNo <> "" Then
            If TempTailDeal(2).szMethodNo = ttTailDealMethod(l).szMethodNo Then
                If nDecimalDigit2 >= ttTailDealMethod(l).nStartNumber And nDecimalDigit2 <= ttTailDealMethod(l).nEndNumber Then
                    If ttTailDealMethod(l).nDealedNumber <> 100 Then
                        nDecimalDigit2 = ttTailDealMethod(l).nDealedNumber
                        If nDecimalDigit2 = 10 Then
                           nDecimalDigit2 = 0
                           nDecimalDigit1 = nDecimalDigit1 + 1
                        End If
                        If nDecimalDigit1 = 10 Then
                           nDecimalDigit1 = 0
                           nIntPart = nIntPart + 1
                        End If
                        Exit For
                    End If
                End If
            End If
        Else '对小数点后一、二位不处理
'            If nDecimalDigit2 < 5 Then
'               nDecimalDigit2 = 0
'            Else
'               nDecimalDigit2 = 0
'               nDecimalDigit1 = nDecimalDigit1 + 1
'            End If
'            If nDecimalDigit1 >= 10 Then
'               nDecimalDigit1 = 0
'               nIntPart = nIntPart + 1
'            End If
            Exit For
        End If
    Next
         
    For l = 1 To nTailMethod
        If TempTailDeal(1).szMethodNo <> "" Then
            If TempTailDeal(1).szMethodNo = ttTailDealMethod(l).szMethodNo Then
                If nDecimalDigit1 >= ttTailDealMethod(l).nStartNumber And nDecimalDigit1 <= ttTailDealMethod(l).nEndNumber Then
                    If ttTailDealMethod(l).nDealedNumber <> 100 Then
                        nDecimalDigit1 = ttTailDealMethod(l).nDealedNumber
                        If nDecimalDigit1 = 10 Then
                           nDecimalDigit1 = 0
                           nIntPart = nIntPart + 1
                        End If
                        Exit For
                    End If
                End If
            End If
       End If
     Next
    
    If nTailMethod = 0 Then
        If nDecimalDigit3 < 5 Then
           nDecimalDigit3 = 0
        Else
           nDecimalDigit3 = 0
           nDecimalDigit2 = nDecimalDigit2 + 1
        End If
        If nDecimalDigit2 >= 10 Then
           nDecimalDigit2 = 0
           nDecimalDigit1 = nDecimalDigit1 + 1
        End If
'        If nDecimalDigit2 < 5 Then
'           nDecimalDigit2 = 0
'        Else
'           nDecimalDigit2 = 0
'           nDecimalDigit1 = nDecimalDigit1 + 1
'        End If
        If nDecimalDigit1 >= 10 Then
           nDecimalDigit1 = 0
           nIntPart = nIntPart + 1
        End If
    End If
    nDecimalPart = nDecimalDigit2 * 10 + nDecimalDigit1 * 100
    vNewValue = nIntPart + CDbl(nDecimalPart) / 1000
    DetailTailDeal = vNewValue
    Exit Function
    
errorHander:
    ShowError ERR_TailDeal
End Function

Public Function TotalTailDeal(sgTotalColValue As Single, szAreaCode As String, szBusType As String, szPriceItem As String, ttTailDeal() As TAreaTailDealMethod, ttTailDealMethod() As TTailDealMethod, sgLastColValue As Single) As TDealValue
    Dim TempTailDeal As TAreaTailDealMethod
    Dim nTailMethod As Integer
    Dim nStation As Long
    
    Dim nTotalColInt As Long
    Dim sgTotalColDecimal As Single
    Dim sgTemp As Single
    Dim sgStartValue As Single
    Dim sgEndValue As Single
    Dim TTempDealValue As TDealValue
    Dim m As Long
    Dim nHalf As Long
    Dim nStart As Long
    Dim nEnd As Long
    Dim FindNo As Long
    
    On Error GoTo errorHander
    '先在总价里减去原先的舍入项
    sgTotalColValue = sgTotalColValue - sgLastColValue
    '把舍入项清零
'    sgLastColValue = 0 + Format(sgTotalColValue, "0.0") - sgTotalColValue
    sgLastColValue = 0 + Int(sgTotalColValue) + Int((sgTotalColValue - Int(sgTotalColValue)) * 10) / 10 - sgTotalColValue
    
'    sgTotalColValue = Format(sgTotalColValue, "0.0")
    sgTotalColValue = Int(sgTotalColValue) + Int((sgTotalColValue - Int(sgTotalColValue)) * 10) / 10
    sgLastColValue = Round(sgLastColValue, 2)
    sgTotalColValue = Round(sgTotalColValue, 2)
    nStation = ArrayLength(ttTailDeal)
    nTailMethod = ArrayLength(ttTailDealMethod)
    nTotalColInt = Int(sgTotalColValue)
    sgTotalColDecimal = Format((sgTotalColValue - nTotalColInt) * 10, "0.00")
    
    m = FindTailDeal(ttTailDeal, szAreaCode, szBusType, "0", szPriceItem) '查找该区域的处理方法
    If m > 0 Then TempTailDeal = ttTailDeal(m)
    
    For m = 1 To nTailMethod
       If TempTailDeal.szMethodNo <> "" Then
           If TempTailDeal.szMethodNo = ttTailDealMethod(m).szMethodNo Then
              '取结束值
              If sgTotalColDecimal > 9 Then
                 sgEndValue = ttTailDealMethod(m).nEndNumber + 1
              Else
                 sgEndValue = ttTailDealMethod(m).nEndNumber
              End If
              '取起始值
              If sgTotalColDecimal > Int(sgTotalColDecimal) Then
                 sgStartValue = ttTailDealMethod(m).nStartNumber - 1
              Else
                 sgStartValue = ttTailDealMethod(m).nStartNumber
              End If
              '进行进位处理
               If sgTotalColDecimal >= sgStartValue And sgTotalColDecimal <= sgEndValue Then
                   If ttTailDealMethod(m).nDealedNumber <> 100 Then
                       sgTemp = ttTailDealMethod(m).nDealedNumber
                       If Format(sgTotalColDecimal, "0.00") > 0 Then
                            sgLastColValue = Format(sgLastColValue + (sgTemp - sgTotalColDecimal) / 10, "0.00")
                            sgTotalColValue = Format(sgTotalColValue + (sgTemp - sgTotalColDecimal) / 10, "0.00")
                       End If
                       Exit For
                   End If
               End If
           End If
      Else
           sgLastColValue = sgLastColValue + Format(sgTotalColValue, "0.0") - sgTotalColValue
           sgTotalColValue = Format(sgTotalColValue, "0.0")
           Exit For
      End If
    Next
    If nTailMethod = 0 Then
       sgLastColValue = sgLastColValue + Format(sgTotalColValue, "0.0") - sgTotalColValue
       sgTotalColValue = Format(sgTotalColValue, "0.0")
    End If
    TTempDealValue.pbTotalCol = True
    TTempDealValue.sgValue = Format(sgTotalColValue, "0.00")
    TTempDealValue.sgLastColValue = Format(sgLastColValue, "0.00")
                                           
    TotalTailDeal = TTempDealValue
    Exit Function
    
errorHander:
    ShowError ERR_TailDeal
End Function

'得到所有尾数处理方法
Public Function GetAllTailDealMethod() As TTailDealMethod()
    Dim odb As New RTConnection
    Dim ttTailDealMethod() As TTailDealMethod
    Dim i As Integer
    Dim szSql As String
    Dim rsTemp As Recordset
    
    On Error GoTo errorHander

    odb.ConnectionString = GetConnectionStr(cszPriceMan)

    szSql = "SELECT * FROM Price_taildeal_info ORDER BY Method_no"

    Set rsTemp = odb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then ReDim ttTailDealMethod(1 To rsTemp.RecordCount)
    i = 1
    Do While Not rsTemp.EOF
       ttTailDealMethod(i).szMethodNo = rsTemp!method_no
       ttTailDealMethod(i).szSerialNo = rsTemp!serial_no
       ttTailDealMethod(i).nStartNumber = rsTemp!start_number
       ttTailDealMethod(i).nEndNumber = rsTemp!end_number
       ttTailDealMethod(i).nDealedNumber = rsTemp!dealed_number
       If Not IsNull(rsTemp!Annotation) Then
          ttTailDealMethod(i).szAnnotation = rsTemp!Annotation
       Else
          ttTailDealMethod(i).szAnnotation = ""
       End If
       i = i + 1
       rsTemp.MoveNext
    Loop
    
    GetAllTailDealMethod = ttTailDealMethod
    
    Set rsTemp = Nothing
    Set odb = Nothing
    Exit Function
errorHander:
End Function

'得到所有区域的尾数处理方法
Public Function GetAreaTailMethod(PriceTableID As String) As TAreaTailDealMethod()
    Dim odb As New RTConnection
    Dim ttTailDeal() As TAreaTailDealMethod
    Dim i As Integer
    Dim szSql As String
    Dim rsTemp As Recordset

    On Error GoTo errorHander

    odb.ConnectionString = GetConnectionStr(cszPriceMan)

    szSql = "SELECT * FROM area_tail_deal_info WHERE price_table_id='" & PriceTableID & "' ORDER BY area_code,Tail_bit_no ,price_item,bus_type"
    Set rsTemp = odb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then ReDim ttTailDeal(1 To rsTemp.RecordCount)
    i = 1
     
    Do While Not rsTemp.EOF
       ttTailDeal(i).szAreaCode = rsTemp!Area_Code
       ttTailDeal(i).szBusType = rsTemp!bus_type
       ttTailDeal(i).szPriceItemCode = rsTemp!price_item
       ttTailDeal(i).nTailBitNo = rsTemp!tail_bit_no
       ttTailDeal(i).szMethodNo = rsTemp!method_no
       rsTemp.MoveNext
       i = i + 1
    Loop
    GetAreaTailMethod = ttTailDeal
    Set rsTemp = Nothing
    Set odb = Nothing
    Exit Function

errorHander:
End Function

Public Sub SaveTailDealMethod(ttTailDealMethod() As TTailDealMethod)
 Dim odb As New RTConnection
 Dim szTemp As String
 Dim i As Integer
 Dim count As Integer
    
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_TailDealMethodSave
        
    On Error GoTo errorHander
    count = ArrayLength(ttTailDealMethod)
    If count > 0 Then
        szTemp = ttTailDealMethod(1).szMethodNo
        odb.ConnectionString = GetConnectionStr(cszPriceMan)
        odb.BeginTrans
        odb.Execute ("DELETE FROM Price_taildeal_info WHERE method_no='" & szTemp & "'")
        For i = 1 To count
            odb.Execute ("INSERT INTO Price_taildeal_info (method_no,serial_no,start_number,end_number,dealed_number,annotation) VALUES ('" & ttTailDealMethod(i).szMethodNo & "','" & ttTailDealMethod(i).szSerialNo & "','" & ttTailDealMethod(i).nStartNumber & "','" & ttTailDealMethod(i).nEndNumber & "','" & ttTailDealMethod(i).nDealedNumber & "','" & ttTailDealMethod(i).szAnnotation & "' ) ")
        Next
        odb.CommitTrans
        WriteOperateLog m_oActiveUser, RIGHT_TailDealMethodSave, "保存尾数处理方法" & szTemp
    End If
    
    Set odb = Nothing
    Exit Sub
    
errorHander:
    ShowError ERR_TailDealMethodSave
End Sub

Public Function DeleteTailDealMethod(ttTailDealMethod() As TTailDealMethod) As Boolean
 Dim odb As New RTConnection
 Dim szTemp As String
 Dim i As Integer
 Dim count As Integer
 Dim rsTemp As Recordset
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_TailDealMethodDelete
        
    On Error GoTo errorHander
    count = ArrayLength(ttTailDealMethod)
    If count > 0 Then
        szTemp = ttTailDealMethod(1).szMethodNo
        odb.ConnectionString = GetConnectionStr(cszPriceMan)
        Set rsTemp = odb.Execute("SELECT * FROM area_tail_deal_info WHERE method_no='" & szTemp & "'")
        If rsTemp.RecordCount = 0 Then
           odb.BeginTrans
           odb.Execute ("DELETE FROM Price_taildeal_info WHERE method_no='" & szTemp & "'")
           WriteOperateLog m_oActiveUser, RIGHT_TailDealMethodDelete, "删除尾数处理方法" & szTemp
           odb.CommitTrans
           DeleteTailDealMethod = True
        Else
           DeleteTailDealMethod = False
           GoTo errorHander
        End If
    End If
    Set rsTemp = Nothing
    Set odb = Nothing
    
    Exit Function
    
errorHander:
    Set rsTemp = Nothing
    Set odb = Nothing
    ShowError ERR_TailDealMethodDelete
End Function

Public Function GetAllAreaTailMethod(PriceTableID As String, ttArea() As TArea, BusType() As String, PriceItem() As String, TailBit() As Integer, MethodNo As String, Optional DealItem As Integer = 1) As TAreaTailDealMethod()
 Dim odb As New RTConnection
 Dim rsTemp As Recordset
 Dim szSql As String
 Dim nBitCount As Integer
 Dim nPriceItemCount As Integer
 Dim nAreaCount As Integer
 Dim i As Integer
 Dim szTailBit As String
 Dim szPriceItem As String
 Dim szArea As String
 Dim ttAreaMethod() As TAreaTailDealMethod
 Dim nBusTypeCount As Integer
 Dim szBusType As String
 
    On Error GoTo errorHander
    odb.ConnectionString = GetConnectionStr(cszPriceMan)
    nBitCount = ArrayLength(TailBit)
    nAreaCount = ArrayLength(ttArea)
    nPriceItemCount = ArrayLength(PriceItem)
    nBusTypeCount = ArrayLength(BusType)
    
    If nPriceItemCount > 0 Then
        For i = 1 To nPriceItemCount
            If i = 1 Then
               szPriceItem = " AND area_tail_deal_info.price_item in ("
            End If
            If i < nPriceItemCount Then
               szPriceItem = szPriceItem & "'" & PriceItem(i) & " ',"
            Else
               szPriceItem = szPriceItem & "'" & PriceItem(i) & " ')"
            End If
        Next
    Else
        szPriceItem = " "
    End If
    
    If nBusTypeCount > 0 Then
        For i = 1 To nBusTypeCount
            If i = 1 Then
               szBusType = " AND area_tail_deal_info.bus_type in ("
            End If
            If i < nBusTypeCount Then
               szBusType = szBusType & "'" & BusType(i) & " ',"
            Else
               szBusType = szBusType & "'" & BusType(i) & " ')"
            End If
        Next
    Else
        szBusType = " "
    End If
    
    If nBitCount > 0 Then
        For i = 1 To nBitCount
            If i < nBitCount Then
               szTailBit = szTailBit & "'" & TailBit(i) & " ',"
            Else
               szTailBit = szTailBit & "'" & TailBit(i) & " '"
            End If
        Next
    Else
        szTailBit = "'0','1','2','3'"
    End If
    If nAreaCount > 0 Then
        For i = 1 To nAreaCount
            If i < nAreaCount Then
               szArea = szArea & "'" & ttArea(i).szAreaCode & " ',"
            Else
               szArea = szArea & "'" & ttArea(i).szAreaCode & " '"
            End If
        Next
        szSql = "SELECT area_tail_deal_info.*,area_code.*,price_item_info.chinese_name FROM area_tail_deal_info,area_code,price_item_info WHERE area_tail_deal_info.price_table_id='" & PriceTableID & "' AND area_code.area_code=area_tail_deal_info.area_code and price_item_info.price_item=area_tail_deal_info.price_item " & szBusType & szPriceItem & " AND tail_bit_no IN (" & szTailBit & ")  AND area_tail_deal_info.area_code in (" & szArea & ")"
    Else
        szSql = "SELECT area_tail_deal_info.*,area_code.*,price_item_info.chinese_name FROM area_tail_deal_info,area_code,price_item_info WHERE area_tail_deal_info.price_table_id='" & PriceTableID & "' AND  area_code.area_code=area_tail_deal_info.area_code and price_item_info.price_item=area_tail_deal_info.price_item " & szBusType & szPriceItem & " AND tail_bit_no IN (" & szTailBit & ")"
    End If
    If MethodNo <> "" Then szSql = szSql & " AND method_no='" & MethodNo & "'"
    
    Set rsTemp = odb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then ReDim ttAreaMethod(1 To rsTemp.RecordCount)
    i = 1

    Do While Not rsTemp.EOF
       ttAreaMethod(i).szMethodNo = rsTemp!method_no
       ttAreaMethod(i).szAreaCode = rsTemp!Area_Code
       ttAreaMethod(i).szAreaName = rsTemp!area_name
       ttAreaMethod(i).szBusType = rsTemp!bus_type
       ttAreaMethod(i).szPriceItemCode = rsTemp!price_item
       ttAreaMethod(i).szPriceItemName = rsTemp!chinese_name
       ttAreaMethod(i).nTailBitNo = rsTemp!tail_bit_no
       If rsTemp!tail_bit_no = 1 Then
          ttAreaMethod(i).szTailBitNoName = szTenCent
       ElseIf rsTemp!tail_bit_no = 2 Then
          ttAreaMethod(i).szTailBitNoName = szCent
       ElseIf rsTemp!tail_bit_no = 3 Then
          ttAreaMethod(i).szTailBitNoName = szCentMili
       Else
          ttAreaMethod(i).szTailBitNoName = szTotalPrice
       End If
       i = i + 1
       rsTemp.MoveNext
    Loop
    GetAllAreaTailMethod = ttAreaMethod
    Set rsTemp = Nothing
    Set odb = Nothing
    Exit Function
    
errorHander:

End Function

'Public Function GetAllArea() As TArea()
''******此处,此接口可以删除
' Dim odb As New RTConnection
' Dim rsTemp As Recordset
' Dim szSql As String
' Dim ttArea() As TArea
' Dim i As Integer
'
'    On Error GoTo errorHander
'    odb.ConnectionString = GetConnectionStr(cszPriceMan)
'    szSql = "SELECT * FROM area_code"
'    Set rsTemp = odb.Execute(szSql)
'    ReDim ttArea(1 To rsTemp.RecordCount)
'    i = 1
'    Do While Not rsTemp.EOF
'       ttArea(i).szAreaCode = rsTemp!Area_Code
'       ttArea(i).szAreaName = rsTemp!area_name
'       ttArea(i).szProvinceInOut = rsTemp!province_in_out
'       i = i + 1
'       rsTemp.MoveNext
'    Loop
'    GetAllArea = ttArea
'    Set rsTemp = Nothing
'    Set odb = Nothing
'    Exit Function
'
'errorHander:
'
'End Function

Public Sub UpdateAreaMethod(PriceTableID As String, AreaCode As String, BusType As String, PriceItem As String, TailBit As Integer, MethodNo As String, Optional DealItem As Integer = 1)
Dim odb As New RTConnection
Dim plAffectedRows As Long

    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_UpdateAreaMethod
   
    On Error GoTo errorHander
    odb.ConnectionString = GetConnectionStr(cszPriceMan)
    odb.BeginTrans
    odb.Execute ("Update area_tail_deal_info SET method_no='" & MethodNo & "' WHERE Price_table_ID='" & PriceTableID & "' AND " & " area_code='" & AreaCode & "' AND bus_type='" & BusType & "' AND price_Item='" & PriceItem & "' AND tail_bit_no='" & TailBit & "'"), plAffectedRows
    If plAffectedRows = 0 Then odb.Execute ("INSERT INTO area_tail_deal_info (price_table_ID,area_code,bus_type,price_item,tail_bit_no,method_no) VALUES ('" & PriceTableID & "','" & AreaCode & "' ,'" & BusType & "','" & PriceItem & "' ,'" & TailBit & "','" & MethodNo & "')")
    odb.CommitTrans
    WriteOperateLog m_oActiveUser, RIGHT_UpdateAreaMethod, "更新或增加票价表[" & PriceTableID & "]的区域[" & AreaCode & "]的尾数处理方法为" & MethodNo
    Set odb = Nothing
    Exit Sub
errorHander:
    ShowError ERR_UpdateAreaMethod
End Sub

Public Sub AddAreaMethod(PriceTableID As String, AreaCode() As String, BusType() As String, PriceItem() As String, TailBit() As Integer, MethodNo As String)
Dim odb As New RTConnection
Dim rsTemp As Recordset
Dim szSql As String
Dim nAreaCount As Integer
Dim nPriceItem As Integer
Dim nBusType As Integer
Dim nTailCount As Integer
Dim szArea As String
Dim szPriceItem As String
Dim szBusType As String
Dim szTail As String
Dim i As Integer
Dim n As Integer
Dim j As Integer
Dim m As Integer

    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_AddAreaMethod
       
    On Error GoTo errorHander
    odb.ConnectionString = GetConnectionStr(cszPriceMan)
    odb.BeginTrans
    nAreaCount = ArrayLength(AreaCode)
    If nAreaCount > 0 Then
       For i = 1 To nAreaCount
            If i < nAreaCount Then
               szArea = szArea & "'" & AreaCode(i) & "',"
            Else
               szArea = szArea & "'" & AreaCode(i) & "'"
            End If
       Next
    End If
    
    nPriceItem = ArrayLength(PriceItem)
    If nPriceItem > 0 Then
       For i = 1 To nPriceItem
            If i < nPriceItem Then
               szPriceItem = szPriceItem & "'" & PriceItem(i) & "',"
            Else
               szPriceItem = szPriceItem & "'" & PriceItem(i) & "'"
            End If
       Next
    End If
    
    nBusType = ArrayLength(BusType)
    If nBusType > 0 Then
       For i = 1 To nBusType
            If i < nBusType Then
               szBusType = szBusType & "'" & BusType(i) & "',"
            Else
               szBusType = szBusType & "'" & BusType(i) & "'"
            End If
       Next
    End If
    
    nTailCount = ArrayLength(TailBit)
    If nTailCount > 0 Then
       For i = 1 To nTailCount
            If i < nTailCount Then
               szTail = szTail & "'" & TailBit(i) & "',"
            Else
               szTail = szTail & "'" & TailBit(i) & "'"
            End If
       Next
    End If
    szSql = "SELECT * FROM area_tail_deal_info WHERE price_table_id='" & PriceTableID & "' AND area_code in (" & szArea & ") AND bus_type in (" & szBusType & ") AND price_item in (" & szPriceItem & ") AND tail_bit_no in (" & szTail & ") "
    
    Set rsTemp = odb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
       GoTo errorHander
    End If
    For i = 1 To nAreaCount
        For m = 1 To nBusType
            For n = 1 To nPriceItem
                For j = 1 To nTailCount
                    odb.Execute ("INSERT INTO area_tail_deal_info (price_table_id,area_code,bus_type,price_item,tail_bit_no,method_no) VALUES ('" & PriceTableID & "' ,'" & AreaCode(i) & "' ,'" & BusType(m) & "' ,'" & PriceItem(n) & "' ,'" & TailBit(j) & "','" & MethodNo & "')")
                Next
            Next
        Next
    Next
    odb.CommitTrans
    Set odb = Nothing
    Exit Sub
errorHander:
    ShowError ERR_AddAreaMethod
End Sub

Public Sub DeleteAreaMethod(PriceTableID As String, AreaCode() As String, BusType() As String, PriceItem() As String, TailBit() As Integer, MethodNo As String)
Dim odb As New RTConnection
Dim rsTemp As Recordset
Dim szSql As String
Dim nAreaCount As Integer
Dim nTailCount As Integer
Dim szArea As String
Dim nBusTypeCount As Integer
Dim szBusType As String
Dim szTail As String
Dim i As Integer
Dim j As Integer
Dim n As Integer
Dim szMethod As String
Dim szPriceItem As String
Dim nPriceItemCount As Integer

    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_DeleteAreaMethod
        
    On Error GoTo errorHander
    odb.ConnectionString = GetConnectionStr(cszPriceMan)
    odb.BeginTrans
    nAreaCount = ArrayLength(AreaCode)
    
    szArea = " AND area_code in ("
    If nAreaCount > 0 Then
       For i = 1 To nAreaCount
            If i < nAreaCount Then
               szArea = szArea & "'" & AreaCode(i) & "',"
            Else
               szArea = szArea & "'" & AreaCode(i) & "') "
            End If
       Next
    Else
       szArea = ""
    End If
    
    nPriceItemCount = ArrayLength(PriceItem)
    szPriceItem = " AND price_item in ("
    If nPriceItemCount > 0 Then
       For i = 1 To nPriceItemCount
            If i < nPriceItemCount Then
               szPriceItem = szPriceItem & "'" & PriceItem(i) & "',"
            Else
               szPriceItem = szPriceItem & "'" & PriceItem(i) & "')"
            End If
       Next
    Else
       szPriceItem = ""
    End If
    
    nBusTypeCount = ArrayLength(BusType)
    szBusType = " AND bus_type in ("
    If nBusTypeCount > 0 Then
       For i = 1 To nBusTypeCount
            If i < nBusTypeCount Then
               szBusType = szBusType & "'" & BusType(i) & "',"
            Else
               szBusType = szBusType & "'" & BusType(i) & "')"
            End If
       Next
    Else
       szBusType = ""
    End If
    
    nTailCount = ArrayLength(TailBit)
    szTail = " AND tail_bit_no in ("
    If nTailCount > 0 Then
       For i = 1 To nTailCount
            If i < nTailCount Then
               szTail = szTail & "'" & TailBit(i) & "',"
            Else
               szTail = szTail & "'" & TailBit(i) & "')"
            End If
       Next
    Else
       szTail = ""
    End If
    
    If MethodNo = "" Then
       szMethod = ""
    Else
       szMethod = " AND method_no='" & MethodNo & "'"
    End If
    szSql = "DELETE FROM area_tail_deal_info WHERE price_table_id='" & PriceTableID & "'" & szArea & szBusType & szPriceItem & szTail & szMethod
    odb.Execute szSql
    odb.CommitTrans
    Set odb = Nothing
    WriteOperateLog m_oActiveUser, RIGHT_DeleteAreaMethod, "删除区域尾数处理方法" & MethodNo
    Exit Sub
errorHander:
    ShowError ERR_DeleteAreaMethod
End Sub

'得到车次的座位类型
Public Function GetAllBusVehicleTypeSeatType(BusID() As String) As TBusVehicleSeatType()
Dim odb As New RTConnection
Dim szSql As String
Dim szSql1 As String
Dim szBusCondition As String
Dim nTemp, i, n, j As Integer
Dim oBaseInfo As BaseInfo
Dim szSeatType() As String
Dim rsTemp As Recordset
Dim rsTemp1 As Recordset
Dim ttBusVehicleSeat() As TBusVehicleSeatType
Dim nTotalSeatNumber() As Integer
Dim szOldCondition As String

    AssertActiveUserValid m_oActiveUser, ERR_TicketPriceMan

    
    Set oBaseInfo = New BaseInfo
    oBaseInfo.Init m_oActiveUser

    szSeatType = oBaseInfo.GetAllSeatType
    Set oBaseInfo = Nothing
    
    On Error GoTo ExitPosition
    odb.ConnectionString = GetConnectionStr(cszPriceMan)
    szSql = "SELECT DISTINCT bva.bus_id,tvi.vehicle_id,vst.seat_type_id, CONVERT(smallint,end_seat_no)-CONVERT(smallint,vst.start_seat_no)+1 as This_Seat_Type_No," & _
        " tvi.vehicle_type_code , vtc.vehicle_type_short_name, vtc.vehicle_type_name " & _
        " FROM bus_vehicle_code bva,vehicle_seat_type_info  vst, " & _
        " Vehicle_info tvi,vehicle_type_code vtc " & _
        " WHERE tvi.vehicle_id=bva.vehicle_id  AND vst.vehicle_id=tvi.vehicle_id " & _
        " AND vtc.vehicle_type_code=tvi.vehicle_type_code "

  
  szSql1 = "SELECT DISTINCT bva.bus_id,tvi.vehicle_id,vst.seat_type_id, " & _
        " tvi.vehicle_type_code , vtc.vehicle_type_short_name, vtc.vehicle_type_name " & _
        " FROM bus_vehicle_code bva,vehicle_seat_type_info  vst, " & _
        " Vehicle_info tvi,vehicle_type_code vtc " & _
        " WHERE tvi.vehicle_id=bva.vehicle_id  AND vst.vehicle_id=tvi.vehicle_id " & _
        " AND vtc.vehicle_type_code=tvi.vehicle_type_code "
       
        
  nTemp = ArrayLength(BusID)
  If nTemp > 0 Then
     For i = 1 To nTemp
         If i = 1 Then szBusCondition = " AND bva.bus_id IN ("
         szBusCondition = szBusCondition & "'" & BusID(i) & "'"
         If i < nTemp Then
            szBusCondition = szBusCondition & ","
         Else
            szBusCondition = szBusCondition & ")"
         End If
     Next
  End If
  szSql = szSql & szBusCondition
  szSql1 = szSql1 & szBusCondition
  
  Set rsTemp = odb.Execute(szSql)
  
  szSql = "SELECT DISTINCT bva.bus_id,tvi.* FROM bus_vehicle_code bva,Vehicle_info tvi WHERE bva.vehicle_id=tvi.vehicle_id" & szBusCondition & " Order by bva.bus_id,tvi.vehicle_id "
  Set rsTemp1 = odb.Execute(szSql)
  If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
  If rsTemp1.RecordCount > 0 Then ReDim nTotalSeatNumber(1 To rsTemp1.RecordCount)
  
  i = 0
  Do While Not rsTemp1.EOF
     i = i + 1
     nTotalSeatNumber(i) = rsTemp1!seat_quantity
    Do While Not rsTemp.EOF
       If rsTemp!Bus_id = rsTemp1!Bus_id And rsTemp!vehicle_id = rsTemp1!vehicle_id Then
          nTotalSeatNumber(i) = nTotalSeatNumber(i) - rsTemp!This_Seat_Type_No
       Else
          Exit Do
       End If
       rsTemp.MoveNext
    Loop
    rsTemp1.MoveNext
 Loop
 
 i = 0
 n = 0
 If rsTemp1.RecordCount > 0 Then rsTemp1.MoveFirst
 Do While Not rsTemp1.EOF
    i = i + 1
    If nTotalSeatNumber(i) > 0 Then n = n + 1
    rsTemp1.MoveNext
 Loop
 
 j = 0
 i = 0
 szOldCondition = szBusCondition
 If rsTemp1.RecordCount > 0 Then rsTemp1.MoveFirst
 Do While Not rsTemp1.EOF
    i = i + 1
    If nTotalSeatNumber(i) > 0 Then
       j = j + 1
       If j = 1 Then szBusCondition = szBusCondition & " AND tvi.vehicle_id IN ("
       szBusCondition = szBusCondition & "'" & rsTemp1!vehicle_id & "'"
       
       If j < n Then
          szBusCondition = szBusCondition & ","
       Else
         szBusCondition = szBusCondition & ")"
         Exit Do
       End If
    End If
    rsTemp1.MoveNext
 Loop
 If szOldCondition = szBusCondition Then szBusCondition = szBusCondition & "AND tvi.vehicle_id IN ( '')"
 
  szSql = " SELECT bva.bus_id,bva.vehicle_id,'" & cszSeatTypeIsNormal & "' as seat_type_id,tvi.vehicle_type_code ," & _
    " vtc.vehicle_type_short_name, vtc.vehicle_type_name " & _
    " FROM bus_vehicle_code bva,vehicle_type_code vtc,Vehicle_info tvi " & _
    " WHERE bva.vehicle_id = tvi.vehicle_id And tvi.vehicle_type_code = vtc.vehicle_type_code " & szBusCondition
    
  szSql = "SELECT DISTINCT a.bus_id,a.seat_type_id,a.vehicle_type_code,a.vehicle_type_short_name,a.vehicle_type_name FROM  (" & szSql1 & " UNION " & szSql & ") a "

  Set rsTemp1 = odb.Execute(szSql)
  
  
  ReDim ttBusVehicleSeat(1 To rsTemp1.RecordCount)
'
  i = 1
  If rsTemp1.RecordCount > 0 Then rsTemp1.MoveFirst
  nTemp = ArrayLength(szSeatType) '所有座位种类
  Do While Not rsTemp1.EOF
     ttBusVehicleSeat(i).szBusID = rsTemp1!Bus_id
     ttBusVehicleSeat(i).szVehicleTypeCode = rsTemp1!vehicle_type_code
     ttBusVehicleSeat(i).szVehicleTypeShortName = rsTemp1!vehicle_type_short_name
     ttBusVehicleSeat(i).szVehicleTypeName = rsTemp1!vehicle_type_name
     ttBusVehicleSeat(i).szSeatTypeID = rsTemp1!seat_type_id
     i = i + 1
     rsTemp1.MoveNext
  Loop

  For i = 1 To ArrayLength(ttBusVehicleSeat)
      For n = 1 To nTemp
        If Trim(ttBusVehicleSeat(i).szSeatTypeID) = Trim(szSeatType(n, 1)) Then
           ttBusVehicleSeat(i).szSeatTypeName = szSeatType(n, 2)
           Exit For
        End If
     Next
  Next
  Set rsTemp1 = Nothing
  Set rsTemp = Nothing
  Set odb = Nothing
  GetAllBusVehicleTypeSeatType = ttBusVehicleSeat
  Exit Function
  
ExitPosition:
  
End Function

''获得所有的票价项的代码，名称，使用标记
'Public Function GetAllTicketItem(PriceTableID As String) As String()
'    Dim oDb As New rtConnection
'    Dim szSql As String
'    Dim rsTemp As Recordset
'    Dim i As Integer
'    Dim aszTemp() As String
'    AssertActiveUserValid m_oActiveUser, ERR_TicketPriceMan
'    oDb.ConnectionString = GetConnectionStr(cszSNPrice)
'    szSql = "SELECT * FROM price_item_info WHERE price_table_id='" & PriceTableID & "' ORDER BY price_item "
'    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount > 0 Then
'        ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 4) As String
'        For i = 1 To rsTemp.RecordCount
'            aszTemp(i, 1) = FormatDbValue(rsTemp!price_item)
'            aszTemp(i, 2) = FormatDbValue(rsTemp!chinese_name)
'            aszTemp(i, 3) = FormatDbValue(rsTemp!use_mark)
'            aszTemp(i, 4) = FormatDbValue(rsTemp!price_table_id)
'            rsTemp.MoveNext
'        Next
'    End If
'    GetAllTicketItem = aszTemp
'    Set rsTemp = Nothing
'    Set oDb = Nothing
'End Function

'获得所有的票价项的代码，名称，使用标记
'PriceTableID 参数不用，因为会给拆帐带来很多改动
Public Function GetAllTicketItem() As String()
    Dim odb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim aszTemp() As String
    AssertActiveUserValid m_oActiveUser, ERR_TicketPriceMan
    odb.ConnectionString = GetConnectionStr(cszSNPrice)
    szSql = "SELECT * FROM price_item_info  ORDER BY price_item "
    Set rsTemp = odb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 3) As String
        For i = 1 To rsTemp.RecordCount
            aszTemp(i, 1) = FormatDbValue(rsTemp!price_item)
            aszTemp(i, 2) = FormatDbValue(rsTemp!chinese_name)
            aszTemp(i, 3) = FormatDbValue(rsTemp!use_mark)
            rsTemp.MoveNext
        Next
    End If
    GetAllTicketItem = aszTemp
    Set rsTemp = Nothing
    Set odb = Nothing
End Function

Public Function GetAllTicketItemRS(Optional pnValid As Integer = -1) As Recordset
    Dim odb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim i As Integer
    AssertActiveUserValid m_oActiveUser, ERR_TicketPriceMan
    odb.ConnectionString = GetConnectionStr(cszSNPrice)
    
    szSql = "SELECT * FROM price_item_info "
    If pnValid <> -1 Then szSql = szSql & " WHERE use_mark = " & pnValid
    Set rsTemp = odb.Execute(szSql)
    Set GetAllTicketItemRS = rsTemp
    Set rsTemp = Nothing
    Set odb = Nothing
End Function


'查找某值某个处理方法
Public Function FindTailDeal(ttTailDeal() As TAreaTailDealMethod, AreaCode As String, BusType As String, TailBitNo As Integer, PriceItem As String) As Long
On Error GoTo errorHander
Dim nStart As Long
Dim nEnd As Long
Dim nFROMAndTo() As Long
Dim nArea As Long
Dim nMid As Long

     nArea = ArrayLength(ttTailDeal)
     nStart = 1
     nEnd = nArea
     
     nFROMAndTo = FindSomeTailDealMethod(ttTailDeal, nStart, nEnd, AreaCode, szAreaCode)
     If nFROMAndTo(1) = 0 Then
        nMid = 0 '没有该区域的对应尾数处理方法
        GoTo exitHander
     Else
           nStart = nFROMAndTo(1)
           nEnd = nFROMAndTo(2)
           nFROMAndTo = FindSomeTailDealMethod(ttTailDeal, nStart, nEnd, TailBitNo, szTailBitNo)
            If nFROMAndTo(1) = 0 Then
               nMid = 0 '没有该位的对应尾数处理方法
               GoTo exitHander
            Else
               nStart = nFROMAndTo(1)
               nEnd = nFROMAndTo(2)
               nFROMAndTo = FindSomeTailDealMethod(ttTailDeal, nStart, nEnd, PriceItem, szPriceItem)
               If nFROMAndTo(1) = 0 Then
                  nMid = 0 '没有该票价项的对应尾数处理方法
                  GoTo exitHander
               Else
                  nStart = nFROMAndTo(1)
                  nEnd = nFROMAndTo(2)
                  nFROMAndTo = FindSomeTailDealMethod(ttTailDeal, nStart, nEnd, BusType, szBusType)
                  If nFROMAndTo(1) = 0 Then
                     nMid = 0 '没有该车次类型的对应尾数处理方法
                     nFROMAndTo = FindSomeTailDealMethod(ttTailDeal, nStart, nEnd, nAllBusType, szBusType)
                     If nFROMAndTo(1) = 0 Then
                        nMid = 0
                        GoTo exitHander
                     Else
                        nMid = nFROMAndTo(1)
                     End If
                  Else
                     nMid = nFROMAndTo(1)
                  End If
            End If
        End If
     End If
     
exitHander:
    FindTailDeal = nMid
    Exit Function
    
errorHander:
End Function

'根据某项查找所需尾数处理方法所在区域，得到开始和结束位置 szString 为 "szAreaCode"、"szBusType" 、"nTailBitNo"、 "szPriceItem"
Public Function FindSomeTailDealMethod(ttTailDeal() As TAreaTailDealMethod, nFirst As Long, nLast As Long, szTemp As Variant, szString As String) As Long()
On Error GoTo errorHander
Dim nStart As Long
Dim nEnd As Long
Dim nHalf As Long
Dim bnFind As Boolean
Dim nFirstAndLast(1 To 2) As Long '记录最终找到的开始、结束值
Dim nStartHalfEnd(1 To 3) As Long '记录第一轮查找的开始、中间、结束值
Dim szAnyTemp As String
Dim szAnyTemp1 As String

     nEnd = nLast
     nStart = nFirst
     If Int((nEnd - nStart) / 2) * 2 = nEnd - nStart Then
        nHalf = nStart + (nEnd - nStart) / 2
     Else
        nHalf = nStart + Int((nEnd - nStart) / 2) + 1
     End If
     If szString = szAreaCode Then
        szAnyTemp = ttTailDeal(nStart).szAreaCode
     ElseIf szString = szBusType Then
        szAnyTemp = ttTailDeal(nStart).szBusType
     ElseIf szString = szTailBitNo Then
        szAnyTemp = ttTailDeal(nStart).nTailBitNo
     Else
        szAnyTemp = ttTailDeal(nStart).szPriceItemCode
     End If
     If Trim(szTemp) = Trim(szAnyTemp) Then
        nFirstAndLast(1) = nStart
        nStartHalfEnd(1) = nStart
        nStartHalfEnd(2) = nStart
        nStartHalfEnd(3) = nEnd
        bnFind = True
     End If
     If szString = szAreaCode Then
        szAnyTemp = ttTailDeal(nEnd).szAreaCode
     ElseIf szString = szBusType Then
        szAnyTemp = ttTailDeal(nEnd).szBusType
     ElseIf szString = szTailBitNo Then
        szAnyTemp = ttTailDeal(nEnd).nTailBitNo
     Else
        szAnyTemp = ttTailDeal(nEnd).szPriceItemCode
     End If
     If Trim(szTemp) = Trim(szAnyTemp) Then
        nFirstAndLast(2) = nEnd
        nStartHalfEnd(1) = nStart
        nStartHalfEnd(2) = nEnd
        nStartHalfEnd(3) = nEnd
        bnFind = True
     End If
     
     If bnFind = False Then
         Do While nStart < nEnd - 1 '查找区域相同的处理方法中某个
            If szString = szAreaCode Then
               szAnyTemp = ttTailDeal(nHalf).szAreaCode
            ElseIf szString = szBusType Then
               szAnyTemp = ttTailDeal(nHalf).szBusType
            ElseIf szString = szTailBitNo Then
               szAnyTemp = ttTailDeal(nHalf).nTailBitNo
            Else
               szAnyTemp = ttTailDeal(nHalf).szPriceItemCode
            End If
           If Trim(szTemp) = Trim(szAnyTemp) Then
              nStartHalfEnd(1) = nStart
              nStartHalfEnd(2) = nHalf
              nStartHalfEnd(3) = nEnd
              bnFind = True
              Exit Do
           ElseIf Trim(szTemp) < Trim(szAnyTemp) Then
              nEnd = nHalf
           ElseIf Trim(szTemp) > Trim(szAnyTemp) Then
              nStart = nHalf
           End If
           If Int((nEnd - nStart) / 2) * 2 = nEnd - nStart Then
              nHalf = nStart + (nEnd - nStart) / 2
           Else
              nHalf = nStart + Int((nEnd - nStart) / 2) + 1
           End If
        Loop
     End If
    If bnFind = True Then
        '前一部分
        nStart = nStartHalfEnd(1)
        nEnd = nStartHalfEnd(2)
        If Int((nEnd - nStart) / 2) * 2 = nEnd - nStart Then
           nHalf = nStart + (nEnd - nStart) / 2
        Else
           nHalf = nStart + Int((nEnd - nStart) / 2) + 1
        End If
        Do While nStart < nEnd - 1
            If szString = szAreaCode Then
               szAnyTemp = ttTailDeal(nHalf).szAreaCode
            ElseIf szString = szBusType Then
               szAnyTemp = ttTailDeal(nHalf).szBusType
            ElseIf szString = szTailBitNo Then
               szAnyTemp = ttTailDeal(nHalf).nTailBitNo
            Else
               szAnyTemp = ttTailDeal(nHalf).szPriceItemCode
            End If
            If Trim(szTemp) <= Trim(szAnyTemp) Then
                nEnd = nHalf
            ElseIf Trim(szTemp) > Trim(szAnyTemp) Then
                 nStart = nHalf
            End If
            If Int((nEnd - nStart) / 2) * 2 = nEnd - nStart Then
               nHalf = nStart + (nEnd - nStart) / 2
            Else
               nHalf = nStart + Int((nEnd - nStart) / 2) + 1
            End If
        Loop
        If nStart < nEnd Then
            If szString = szAreaCode Then
               szAnyTemp = ttTailDeal(nStart).szAreaCode
            ElseIf szString = szBusType Then
               szAnyTemp = ttTailDeal(nStart).szBusType
            ElseIf szString = szTailBitNo Then
               szAnyTemp = ttTailDeal(nStart).nTailBitNo
            Else
               szAnyTemp = ttTailDeal(nStart).szPriceItemCode
            End If
            
            If szString = szAreaCode Then
               szAnyTemp1 = ttTailDeal(nEnd).szAreaCode
            ElseIf szString = szBusType Then
               szAnyTemp1 = ttTailDeal(nEnd).szBusType
            ElseIf szString = szTailBitNo Then
               szAnyTemp1 = ttTailDeal(nEnd).nTailBitNo
            Else
               szAnyTemp1 = ttTailDeal(nEnd).szPriceItemCode
            End If
            If szAnyTemp = szAnyTemp1 Then
               nFirstAndLast(1) = nStart
               If nFirstAndLast(2) < nEnd Then nFirstAndLast(2) = nEnd
            Else
               If Trim(szTemp) = Trim(szAnyTemp1) Then
                  nFirstAndLast(1) = nEnd
               Else
                  nFirstAndLast(1) = nStart
                  nFirstAndLast(2) = 0
                  GoTo exitHander
               End If
            End If
        End If
        '后一部分
        nStart = nStartHalfEnd(2)
        nEnd = nStartHalfEnd(3)
        If Int((nEnd - nStart) / 2) * 2 = nEnd - nStart Then
           nHalf = nStart + (nEnd - nStart) / 2
        Else
           nHalf = nStart + Int((nEnd - nStart) / 2) + 1
        End If
        Do While nStart < nEnd - 1
            If szString = szAreaCode Then
               szAnyTemp = ttTailDeal(nHalf).szAreaCode
            ElseIf szString = szBusType Then
               szAnyTemp = ttTailDeal(nHalf).szBusType
            ElseIf szString = szTailBitNo Then
               szAnyTemp = ttTailDeal(nHalf).nTailBitNo
            Else
               szAnyTemp = ttTailDeal(nHalf).szPriceItemCode
            End If
            If Trim(szTemp) < Trim(szAnyTemp) Then
                 nEnd = nHalf
            ElseIf Trim(szTemp) >= Trim(szAnyTemp) Then
                 nStart = nHalf
            End If
            If Int((nEnd - nStart) / 2) * 2 = nEnd - nStart Then
                nHalf = nStart + (nEnd - nStart) / 2
            Else
               nHalf = nStart + Int((nEnd - nStart) / 2) + 1
            End If
        Loop
        If nStart < nEnd Then
            If szString = szAreaCode Then
               szAnyTemp = ttTailDeal(nStart).szAreaCode
            ElseIf szString = szBusType Then
               szAnyTemp = ttTailDeal(nStart).szBusType
            ElseIf szString = szTailBitNo Then
               szAnyTemp = ttTailDeal(nStart).nTailBitNo
            Else
               szAnyTemp = ttTailDeal(nStart).szPriceItemCode
            End If
            
            If szString = szAreaCode Then
               szAnyTemp1 = ttTailDeal(nEnd).szAreaCode
            ElseIf szString = szBusType Then
               szAnyTemp1 = ttTailDeal(nEnd).szBusType
            ElseIf szString = szTailBitNo Then
               szAnyTemp1 = ttTailDeal(nEnd).nTailBitNo
            Else
               szAnyTemp1 = ttTailDeal(nEnd).szPriceItemCode
            End If
            If szAnyTemp = szAnyTemp1 Then
               If nFirstAndLast(1) = 0 Then nFirstAndLast(1) = nStart
               nFirstAndLast(2) = nEnd
            Else
               If Trim(szTemp) = Trim(szAnyTemp) Then nFirstAndLast(2) = nStart
            End If
        End If
    End If
    
exitHander:

    If nFirstAndLast(1) = 0 Then nFirstAndLast(1) = nFirstAndLast(2)
    If nFirstAndLast(2) = 0 Then nFirstAndLast(2) = nFirstAndLast(1)
    
    FindSomeTailDealMethod = nFirstAndLast
    Exit Function
    
errorHander:
End Function


'获得使用的所有票价项
Public Function GetUseTicketItem() As String()
    Dim odb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim aszTemp() As String
    AssertActiveUserValid m_oActiveUser, ERR_TicketPriceMan
    odb.ConnectionString = GetConnectionStr(cszSNPrice)
    szSql = "SELECT * FROM price_item_info WHERE use_mark=1"
    Set rsTemp = odb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 2) As String
        For i = 1 To rsTemp.RecordCount
            aszTemp(i, 1) = FormatDbValue(rsTemp!price_item)
            aszTemp(i, 2) = FormatDbValue(rsTemp!chinese_name)
            rsTemp.MoveNext
        Next
    End If
    GetUseTicketItem = aszTemp
    Set rsTemp = Nothing
    Set odb = Nothing
End Function


