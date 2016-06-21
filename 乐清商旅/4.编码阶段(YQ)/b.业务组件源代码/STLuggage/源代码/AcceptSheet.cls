VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AcceptSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E039C23002B"
'行包受理单
Option Base 0

Option Explicit



'受理单号
'##ModelId=3E03B0D702F3
Private m_szSheetID As String

'托运类型
'##ModelId=3E03B10C0000
Private m_szAcceptType As String

'起点站代码
'##ModelId=3E03B12E0059
Private m_szStartStationID As String

'起点站名称
'##ModelId=3E03B14F01DD
Private m_szStartStationName As String

'终点站代码
'##ModelId=3E03B1550344
Private m_szDesStationID As String

'终点站
'##ModelId=3E03B15C00D7
Private m_szDesStationName As String

'里程
'##ModelId=3E03B1AC01CD
Private m_dbMileage As Double

'承运车次
'##ModelId=3E03B1F00134
Private m_szBusID As String

'车次日期
'##ModelId=3E03B1F60332
Private m_dtBusDate As Date

'计重
'##ModelId=3E03B29D0314
Private m_dbCalWeight As Double

'实重
'##ModelId=3E03B2CD01FA
Private m_dbActWeight As Double

'起始标签号
'##ModelId=3E03B2E90359
Private m_szStartLabelID As String

'结束标签号
'##ModelId=3E03B2F90230
Private m_szEndLabelID As String

'件数
'##ModelId=3E03B3120286
Private m_nNumber As Integer

'超重件数
'##ModelId=3E03B3250305
Private m_nOverNumber As Integer

'受理单状态
'##ModelId=3E03B37B01DD
Private m_nStatus As Integer

'受理单状态
'##ModelId=3E03B3940251
Private m_szStatusString As String

'托运人
'##ModelId=3E03B3A40268
Private m_szShipper As String

'收件人
'##ModelId=3E03B3BB0365
Private m_szPicker As String

'交货方式
'##ModelId=3E03B3E300B0
Private m_szPickType As String

'收件人联系电话
'##ModelId=3E03B40301CE
Private m_szPickerPhone As String

'收件人联系地址
'##ModelId=3E03B5870302
Private m_szPickerAddress As String

'受理人代码
'##ModelId=3E03B5AF018D
Private m_szOperatorID As String

'受理人姓名
'##ModelId=3E03B5FD019A
Private m_szOperator As String

'受理时间
'##ModelId=3E03B5C40314
Private m_dtOperateTime As Date

'总运价
'##ModelId=3E03B62101CD

Private m_dbTotalPrice As Double

'运费项数组
'##ModelId=3E03B68E0008
Private m_atPriceItems() As TLuggagePriceItem


Private m_szLicenseTagNo As String '车牌
Private m_szPack As String '包装

Private m_dbCalBasePrice As String '计算的行包的运费价



Private m_szTransTicketID  As String '运输单号
Private m_szInsuranceID As String '保险单号
Private m_dbSettleRatio As Double '应结运费比率
Private m_dbSettlePrice As Double '应结运费
Private m_szAnnotation1 As String '备注1
Private m_szAnnotation2 As String '备注2
Private m_szSellStationID As String '受理站


Private m_bIsOneFormulaEachStation As Boolean '是否每个站点一个公式?



'活动用户
Private m_oActiveUser As ActiveUser
Private m_nObjectStatus As EObjectStatus

Private m_szLuggageName As String
Private m_dbOldTotalPrice As Double
Private m_szLuggageShipperPhone As String

'本模块私有定义部分
'==========================================================================
'错误号
Private Enum EErrAcceptSheet
    ERR_AcceptSheetWithout = ERR_LuggageAcceptSheet + 21  '没有些受理单     30021开始
    ERR_AcceptSheetHaveLuggageDetail = ERR_LuggageAcceptSheet + 22 '行包明细已存在
    ERR_AcceptSheetHaveNo = ERR_LuggageAcceptSheet + 23 '行包明细不存在
    ERR_AcceptSheetNoPriceIiem = ERR_LuggageAcceptSheet + 24 '没有票价项
    ERR_AcceptSheetNoStationName = ERR_LuggageAcceptSheet + 25 '没有站点
    
    ERR_StationItemFormulaNotSet = ERR_LuggageAcceptSheet + 26 '该站点的行包计算公式未设置
    
    
End Enum
'权限号

Private Enum ERightAcceptSheet
    RIGHT_AcceptSheet = ERR_LuggageAcceptSheet + cnMidRightBegin             '30100开始
    RIGHT_AcceptSheetEdit = RIGHT_AcceptSheet + cnMidRightStep * 1    '更改行包单
    RIGHT_AcceptSheetDelete = RIGHT_AcceptSheet + cnMidRightStep * 2    '删除行包单
End Enum


'##ModelId=3E04A7B90144
Public Property Let PriceItems(Value() As TLuggagePriceItem)
m_atPriceItems = Value
End Property
'##ModelId=3E04A7BA002D
Public Property Get PriceItems() As TLuggagePriceItem()

PriceItems = m_atPriceItems
End Property
'初始化活动对象
'##ModelId=3E03A0940076
'**************************************************
'Member Code:S1
'Brief Description:初始化对象
'Engineer:
'Date Generated:2002/12/21
'Last Revision Date:
'**************************************************
Public Sub Init(poAUser As ActiveUser)
    On Error GoTo InitErr
    Dim oParam As New SystemParam
    Dim rsTemp As New Recordset
    Dim oDb As New RTConnection
    Dim szSql As String
    
    Set m_oActiveUser = poAUser
    m_szOperator = poAUser.UserName
    m_szOperatorID = poAUser.UserID
    oParam.Init poAUser
    
    m_szStartStationID = poAUser.SellStationID
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    szSql = " SELECT si.station_id, si.station_name FROM station_info si, sell_station_info ssi WHERE  ssi.station_id = si.station_id " _
            & " AND sell_station_id = " & TransFieldValueToString(m_szStartStationID)
    Set rsTemp = oDb.Execute(szSql)
    
    If rsTemp.RecordCount > 0 Then
        m_szStartStationID = FormatDbValue(rsTemp!station_id)
        m_szStartStationName = FormatDbValue(rsTemp!station_name)
    Else
        m_szStartStationID = ""
        m_szStartStationName = ""
    End If
    
    m_bIsOneFormulaEachStation = True 'oParam.IsOneFormulaEachStation
    
    Exit Sub
InitErr:
    Call RaiseError(err.Number, "AcceptSheet:Init Method", err.Description)
End Sub

'指定受理单
'##ModelId=3E03A11D0023
Public Sub Identify(SheetID As String)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageAcceptSheet
    
    On Error GoTo IdentifyErr

       m_szSheetID = SheetID
       RefreshMemoryInfo

    Exit Sub
IdentifyErr:
    Call RaiseError(err.Number, "AcceptSheet:Identify Method", err.Description)
End Sub

Private Sub RefreshMemoryInfo() '按照对象现在的ID将数据库中相应的信息读至对象内
    Dim rsTemp As Recordset, rstemp1 As Recordset
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim i As Integer, j As Integer
    
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT * " _
        & " FROM luggage_accept_info  " _
        & " WHERE luggage_id='" & m_szSheetID & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_AcceptSheetWithout
    
    m_szAcceptType = GetLuggageTypeString(FormatDbValue(rsTemp!Accept_type))
    m_szBusID = FormatDbValue(rsTemp!bus_id)
    m_szDesStationID = FormatDbValue(rsTemp!des_station_id)
    m_szDesStationName = FormatDbValue(rsTemp!des_station_name)
    m_szLuggageName = FormatDbValue(rsTemp!luggage_name)
    m_szStartStationID = FormatDbValue(rsTemp!start_station_id)
    m_szStartStationName = FormatDbValue(rsTemp!start_station_name)
    m_dbMileage = FormatDbValue(rsTemp!Mileage)
    m_dtBusDate = FormatDbValue(rsTemp!bus_date)
    m_dbCalWeight = FormatDbValue(rsTemp!cal_weight)
    m_dbActWeight = FormatDbValue(rsTemp!fact_weight)
    m_szStartLabelID = FormatDbValue(rsTemp!start_label_id)
    m_szEndLabelID = FormatDbValue(rsTemp!end_label_id)
    m_nNumber = FormatDbValue(rsTemp!baggage_number)
    m_nOverNumber = FormatDbValue(rsTemp!over_weight_number)
    m_nStatus = FormatDbValue(rsTemp!Status)
    m_szShipper = FormatDbValue(rsTemp!Shipper)
    m_szPickType = FormatDbValue(rsTemp!pick_type)
    m_szPicker = FormatDbValue(rsTemp!Picker)
    m_szPickerAddress = FormatDbValue(rsTemp!picker_address)
    m_szPickerPhone = FormatDbValue(rsTemp!picker_phone)
    m_szOperatorID = FormatDbValue(rsTemp!user_id)
    m_dtOperateTime = FormatDbValue(rsTemp!operation_time)
    m_szLuggageShipperPhone = FormatDbValue(rsTemp!shipper_phone)
    
    m_szPack = FormatDbValue(rsTemp!Pack)
    m_szLicenseTagNo = FormatDbValue(rsTemp!license_tag_no)
    
    
    
    m_szTransTicketID = FormatDbValue(rsTemp!trans_ticket_id)
    m_szInsuranceID = FormatDbValue(rsTemp!insurance_id)
    m_dbSettleRatio = FormatDbValue(rsTemp!settle_ratio)
    m_dbSettlePrice = FormatDbValue(rsTemp!settle_price)
    m_szAnnotation1 = FormatDbValue(rsTemp!Annotation1)
    m_szAnnotation2 = FormatDbValue(rsTemp!Annotation2)
    
    m_szSellStationID = FormatDbValue(rsTemp!sell_station_id)
    
    
    
    If m_szOperatorID = m_oActiveUser.UserID Then
        m_szOperator = m_oActiveUser.UserName
    Else
        szSql = "SELECT user_name FROM user_info WHERE user_id =" & TransFieldValueToString(m_szOperatorID)
        Set rsTemp = oDb.Execute(szSql)
        m_szOperator = FormatDbValue(rsTemp!user_name)
    End If
    
    szSql = "SELECT * FROM luggage_price_item_info WHERE accept_type= " & TransFieldValueToString(GetLuggageTypeInt(m_szAcceptType))
    Set rstemp1 = oDb.Execute(szSql)
    
    szSql = "SELECT * FROM luggage_price_lst WHERE luggage_id=" & TransFieldValueToString(m_szSheetID)
    Set rsTemp = oDb.Execute(szSql)
    j = 1
    For i = 1 To rstemp1.RecordCount
        If FormatDbValue(rstemp1!use_mark) = 0 Then
            ReDim Preserve m_atPriceItems(1 To j) As TLuggagePriceItem
            m_atPriceItems(j).PriceID = FormatDbValue(rstemp1!charge_item)
            m_atPriceItems(j).PriceName = FormatDbValue(rstemp1!chinese_name)
            m_atPriceItems(j).PriceValue = FormatDbValue(rsTemp.Fields(Trim("price_item_" + Trim(Str(j)))))
            j = j + 1
        End If
        rstemp1.MoveNext
    Next
    m_dbTotalPrice = FormatDbValue(rsTemp!price_total)
    
    Set rsTemp = Nothing
    
End Sub

'新增受理单
'##ModelId=3E03A16D00E6
Public Sub AddNew()
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageAcceptSheet
    
    On Error GoTo AddNewErr
    Dim atTemp() As TLuggagePriceItem
    m_nObjectStatus = ST_AddObj
    m_nStatus = 0
    '把属性值清空
    m_szAcceptType = ""
    m_szBusID = ""
    m_szDesStationID = ""
    m_szDesStationName = ""
    m_szEndLabelID = ""
    m_szDesStationName = ""
    m_szEndLabelID = ""
    m_szLuggageName = ""
    m_szLuggageShipperPhone = ""
'    m_szOperator = ""
'    m_szOperatorID = ""
    m_szPicker = ""
    m_szPickerAddress = ""
    m_szPickerPhone = ""
    m_szPickType = ""
    m_szSheetID = ""
    m_szShipper = ""
    m_szStartLabelID = ""
'    m_szStartStationID = ""
'    m_szStartStationName = ""
    m_szStatusString = ""
    
    m_nNumber = 0
    m_nOverNumber = 0
    m_dbActWeight = 0
    m_dbCalWeight = 0
    m_dbMileage = 0
    m_dbOldTotalPrice = 0
    m_dbTotalPrice = 0
    m_atPriceItems = atTemp
    m_dtBusDate = Date
    m_dtOperateTime = Now
    
    m_szPack = ""
    m_szLicenseTagNo = ""
    
    
    m_szTransTicketID = ""
    m_szInsuranceID = ""
    m_dbSettleRatio = 0
    m_dbSettlePrice = 0
    m_szAnnotation1 = ""
    m_szAnnotation2 = ""
    
    m_szSellStationID = m_oActiveUser.SellStationID
    
    Exit Sub
AddNewErr:
    Call RaiseError(err.Number, "AcceptSheet:AddNew Method", err.Description)
End Sub



'更新受理单
'##ModelId=3E03A179024C
Public Sub Update()
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageAcceptSheet
    
    On Error GoTo UpdateErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szSqlField As String
    
    Select Case m_nObjectStatus
    Case ST_AddObj
        AddObject
    Case ST_EditObj
        AssertHaveRight m_oActiveUser, RIGHT_AcceptSheetEdit
        
        UpdateToDB
    Case ST_NormalObj
    Case Else
        ShowError ERR_NotAvailable
    End Select

    Exit Sub
UpdateErr:
    Call RaiseError(err.Number, "AcceptSheet:Update Method", err.Description)
End Sub

Private Sub AddObject()

End Sub

Private Sub UpdateToDB()
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szSqlField As String
    On Error GoTo Here
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT * FROM luggage_accept_info WHERE luggage_id =" & TransFieldValueToString(m_szSheetID)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_AcceptSheetWithout
    
    szSql = "UPDATE luggage_accept_info SET" _
        & "accept_type='" & m_szAcceptType & "'," _
        & "start_station_id='" & m_szStartStationID & "'," _
        & "start_station_name='" & m_szStartStationName & "'," _
        & "des_station_id='" & m_szDesStationID & "'," _
        & "des_station_name='" & m_szDesStationName & "'," _
        & "mileage=" & m_dbMileage & "," _
        & "bus_id='" & m_szBusID & "'," _
        & "bus_date='" & m_dtBusDate & "'," _
        & "cal_weight=" & m_dbActWeight & "," _
        & "fact_weight=" & m_dbCalWeight & "," _
        & "start_label_id='" & m_szStartLabelID & "'," _
        & "end_label_id='" & m_szEndLabelID & "'," _
        & "baggage_number=" & m_nNumber & "," _
        & "over_weight_number=" & m_nOverNumber & "," _
        & "status=" & m_nStatus & "," _
        & "shipper='" & m_szShipper & "'," _
        & "pick_type='" & m_szPickType & "'," _
        & "picker='" & m_szPicker & "'," _
        & "picker_address='" & m_szPickerAddress & "'," _
        & "picker_phone='" & m_szPickerPhone & "'," _
        & "user_id='" & m_szOperatorID & "'," _
        & "operation_time='" & m_dtOperateTime & "'," _
        & "pack='" & m_szPack & "'," _
        & "license_tag_no='" & m_szLicenseTagNo & "' ,  "
        
        
        
    szSql = szSql & " trans_ticket_id = " & TransFieldValueToString(m_szTransTicketID) & " , " _
        & " insurance_id = " & TransFieldValueToString(m_szInsuranceID) & " , " _
        & " settle_ratio = " & TransFieldValueToString(m_dbSettleRatio) & " , " _
        & " settle_price = " & TransFieldValueToString(m_dbSettlePrice) & " , " _
        & " annotation1 = " & TransFieldValueToString(m_szAnnotation1) & " , " _
        & " annotation2= " & TransFieldValueToString(m_szAnnotation2) & " " _
        & " WHERE luggage_id =" & TransFieldValueToString(m_szSheetID)
        
    oDb.Execute szSql
    Set oDb = Nothing
    Exit Sub
Here:
    Call RaiseError(err.Number, "AcceptSheet:GetLugItemDetail Method", err.Description)
End Sub


Private Sub DeleteObject()
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    oDb.Execute ("DELETE luggage_accept_info WHERE luggage_id='" & m_szSheetID & "'")
    Set oDb = Nothing
End Sub
'删除受理单
'##ModelId=3E03A52A0345
Public Sub Delete()
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageAcceptSheet
    AssertHaveRight m_oActiveUser, RIGHT_AcceptSheetDelete
    
    On Error GoTo DeleteErr
    
    DeleteObject
    m_nObjectStatus = ST_NotAvailable

    Exit Sub
DeleteErr:
    Call RaiseError(err.Number, "AcceptSheet:Delete Method", err.Description)
End Sub

'得到行包明细信息
'##ModelId=3E03BB180149
Public Function GetLugItemDetail() As TLuggageItemInfo()
    On Error GoTo GetLugItemDetailErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim nCount As Integer
    Dim szaTemp() As TLuggageItemInfo
    Dim i As Integer
    szSql = "SELECT ldl.*,lkc.kinds_name  " _
        & " FROM luggage_detail_lst ldl,luggage_kinds_code lkc " _
        & " WHERE ldl.kinds_code = lkc.kinds_code AND luggage_id=" & TransFieldValueToString(m_szSheetID)
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nCount <> 0 Then
      ReDim szaTemp(1 To nCount)
      For i = 1 To nCount
          szaTemp(i).LabelID = FormatDbValue(rsTemp!label_id)
          szaTemp(i).LuggageID = FormatDbValue(rsTemp!luggage_id)
          szaTemp(i).LuggageName = FormatDbValue(rsTemp!luggage_name)
          szaTemp(i).LuggageType = FormatDbValue(rsTemp!Kinds_Code)
          szaTemp(i).LuggageTypeName = FormatDbValue(rsTemp!Kinds_Name)
          szaTemp(i).Number = FormatDbValue(rsTemp!baggage_number)
          szaTemp(i).ActWeight = FormatDbValue(rsTemp!fact_weight)
          szaTemp(i).CalWeight = FormatDbValue(rsTemp!cal_weight)
          szaTemp(i).luggage_bulk = FormatDbValue(rsTemp!luggage_bulk)
          szaTemp(i).PackType = FormatDbValue(rsTemp!package_type)
       rsTemp.MoveNext
      Next i
    GetLugItemDetail = szaTemp
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
    Exit Function
GetLugItemDetailErr:
    Call RaiseError(err.Number, "AcceptSheet:GetLugItemDetail Method", err.Description)
End Function

'添加行包明细
'##ModelId=3E03BDBC030E
Public Sub AddLugItem(patItemInfo() As TLuggageItemInfo)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageAcceptSheet
    AssertHaveRight m_oActiveUser, RIGHT_AcceptSheetEdit

On Error GoTo AddLugItemErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim ArrLen As Integer
    Dim i As Integer
    Dim rsTemp As Recordset
    
'    AssertActiveUserValid m_oActiveUser, ERR_LugInfo
'    AssertHaveRight m_oActiveUser, RIGHT_AcceptSheet
    ArrLen = ArrayLength(patItemInfo)
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    If ArrLen > 0 Then
        For i = 1 To ArrLen
            szSql = "SELECT * " _
                & " FROM luggage_detail_lst " _
                & " WHERE label_id =" & TransFieldValueToString(patItemInfo(i).LabelID)
            Set rsTemp = oDb.Execute(szSql)
            If rsTemp.RecordCount <> 0 Then ShowError ERR_AcceptSheetHaveLuggageDetail
            oDb.BeginTrans
            szSql = "INSERT luggage_detail_lst (" _
                  & "label_id," _
                  & "luggage_id," _
                  & "luggage_name," _
                  & "kinds_code," _
                  & "baggage_number," _
                  & "fact_weight," _
                  & "cal_weight," _
                  & "luggage_bulk," _
                  & "package_type) VALUES ('" _
                  & patItemInfo(i).LabelID & "','" _
                  & patItemInfo(i).LuggageID & "','" _
                  & patItemInfo(i).LuggageName & "','" _
                  & patItemInfo(i).LuggageType & "'," _
                  & patItemInfo(i).Number & "," _
                  & patItemInfo(i).ActWeight & "," _
                  & patItemInfo(i).CalWeight & ",'" _
                  & patItemInfo(i).luggage_bulk & "','" _
                  & patItemInfo(i).PackType & "')"
           oDb.Execute (szSql)
           oDb.CommitTrans
        Next i
    End If
    Set oDb = Nothing
    Exit Sub
AddLugItemErr:
    Call RaiseError(err.Number, "AcceptSheet:AddLugItem Method", err.Description)
End Sub

'更改行包明细
'##ModelId=3E03BF8403BD
Public Sub UpdateLugItem(patItemInfo() As TLuggageItemInfo)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageAcceptSheet
    AssertHaveRight m_oActiveUser, RIGHT_AcceptSheetEdit
    
    On Error GoTo UpdateLugItemErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim ArrLen As Integer
    Dim i As Integer
    Dim rsTemp As Recordset
    
    
    ArrLen = ArrayLength(patItemInfo)
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    If ArrLen > 0 Then
    For i = 1 To ArrLen
    szSql = "SELECT * " _
        & " FROM luggage_detail_lst " _
        & " WHERE label_id =" & TransFieldValueToString(patItemInfo(i).LabelID)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_AcceptSheetHaveNo
    
    oDb.BeginTrans
        
     '& " label_id='" & patItemInfo(i).LabelID & "',"
        szSql = "UPDATE luggage_detail_lst SET " _
              & " luggage_id='" & patItemInfo(i).LuggageID & "'," _
              & " luggage_name='" & patItemInfo(i).LuggageName & "'," _
              & " kinds_code='" & patItemInfo(i).LuggageType & "'," _
              & " baggage_number=" & patItemInfo(i).Number & "," _
              & " fact_weight=" & patItemInfo(i).ActWeight & "," _
              & " cal_weight=" & patItemInfo(i).CalWeight & "," _
              & " luggage_bulk='" & patItemInfo(i).luggage_bulk & "'," _
              & " package_type='" & patItemInfo(i).PackType & "'" _
              & " WHERE luggage_id =" & TransFieldValueToString(m_szSheetID)
              
              
        oDb.Execute (szSql)
    oDb.CommitTrans
    Next i
    
    End If
    Set oDb = Nothing
    Exit Sub
UpdateLugItemErr:
    Call RaiseError(err.Number, "AcceptSheet:UpdateLugItem Method", err.Description)
End Sub

'删除指定行包明细
'##ModelId=3E03C06C01E0
Public Sub DeleteLugItem(pszLabelID As String)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageAcceptSheet
    AssertHaveRight m_oActiveUser, RIGHT_AcceptSheetEdit
    
    On Error GoTo DeleteLugItemErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT * " _
        & " FROM luggage_detail_lst " _
        & " WHERE label_id =" & TransFieldValueToString(pszLabelID)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_AcceptSheetHaveNo
    oDb.BeginTrans
    oDb.Execute ("DELETE luggage_detail_lst WHERE label_id='" & pszLabelID & "'")
    oDb.CommitTrans
    Set oDb = Nothing
    Exit Sub
DeleteLugItemErr:
    Call RaiseError(err.Number, "AcceptSheet:DeleteLugItem Method", err.Description)
End Sub

'统计行包明细信息
''汇总当前受理单的所有行包明细的件数、计重、实重、标签号放入属性中
'##ModelId=3E03C0F40359
Public Sub TotalLugItem()
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageAcceptSheet
    
    On Error GoTo TotalLugItemErr
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset

'    AssertActiveUserValid m_oActiveUser, ERR_LugInfo
'    AssertHaveRight m_oActiveUser, RIGHT_AcceptSheet

    Dim i As Integer
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    szSql = "SELECT Max(label_id) AS label_id,luggage_id,SUM(baggage_number) AS baggage_number,SUM(fact_weight) AS fact_weight,SUM(cal_weight) AS cal_weight " _
        & " FROM luggage_detail_lst  " _
        & " WHERE luggage_id='" & m_szSheetID & "'" _
        & " GROUP BY luggage_id "
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 0 Then
        '返回的标签号有问题，看界面的需要再说
        m_szStartLabelID = FormatDbValue(rsTemp!label_id)
        m_nNumber = FormatDbValue(rsTemp!baggage_number)
        m_dbActWeight = FormatDbValue(rsTemp!fact_weight)
        m_dbCalWeight = FormatDbValue(rsTemp!cal_weight)
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
    Exit Sub
TotalLugItemErr:
    Call RaiseError(err.Number, "AcceptSheet:TotalLugItem Method", err.Description)
End Sub

'根据件数、计重、超重件数属性计算行包托运价
'##ModelId=3E03E61101D4
Public Sub CalculatePrice(AcceptType As Integer, StationID As String, VehicleType As String)

    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_LuggageAcceptSheet
    
    On Error GoTo CalculatePriceErr
    Static oCaculator As Object
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim oParam As New SystemParam
    Dim szItemFormula As String
    Dim i As Integer
    Dim szSql As String
'    AssertActiveUserValid m_oActiveUser, ERR_LugInfo
'    AssertHaveRight m_oActiveUser, RIGHT_AcceptSheet
    m_nStatus = 0
    '初始化对象
    If oCaculator Is Nothing Then
        Set oCaculator = CreateObject("STLugPriLib.LugFunLib")
        oCaculator.Init m_oActiveUser
    End If
    m_dbTotalPrice = 0
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    
    '重新赋值m_bIsOneFormulaEachStation，解决温州行包联网时的不会刷新的错误
    oParam.Init m_oActiveUser
    m_bIsOneFormulaEachStation = True 'oParam.IsOneFormulaEachStation
    
    If m_bIsOneFormulaEachStation Then
        '如果为每个站点一个公式。
        
        szSql = "SELECT f.* , l.charge_item , l.accept_type , l.station_id , l.formula_id , p.chinese_name " _
            & " FROM luggage_price_item_info p , luggage_item_formula_lst l , luggage_formula_info f " _
            & " WHERE l.charge_item = p.charge_item AND l.accept_type = p.accept_type AND l.formula_id=f.formula_id " _
            & " AND p.use_mark=0 AND l.accept_type =" & TransFieldValueToString(AcceptType) _
            & " AND l.station_id = " & TransFieldValueToString(StationID)
        Set rsTemp = oDb.Execute(szSql)
    Else
        '如果统一一个公式
        szSql = "SELECT * FROM luggage_price_item_info WHERE use_mark=0 AND accept_type =" & TransFieldValueToString(AcceptType)
        Set rsTemp = oDb.Execute(szSql)
        
    End If
    If rsTemp.RecordCount <> 0 Then
        ReDim m_atPriceItems(1 To rsTemp.RecordCount)
        For i = 1 To rsTemp.RecordCount
            szItemFormula = FormatDbValue(rsTemp!item_formula)
            '分项价格小数点后保留两位
            m_atPriceItems(i).PriceValue = Round(CallByName(oCaculator, szItemFormula, VbMethod, VehicleType, m_dbMileage, m_dbCalWeight, m_nNumber, m_nOverNumber, _
                          rsTemp!parameter_1, rsTemp!parameter_2, rsTemp!parameter_3, rsTemp!parameter_4, rsTemp!parameter_5, rsTemp!parameter_6, rsTemp!parameter_7, rsTemp!parameter_8, rsTemp!parameter_9, rsTemp!parameter_10))
            m_atPriceItems(i).PriceID = FormatDbValue(rsTemp!charge_item)
            m_atPriceItems(i).PriceName = FormatDbValue(rsTemp!chinese_name)
            rsTemp.MoveNext
        Next i
    Else
        If m_bIsOneFormulaEachStation Then
            '出行包的该站点的计算公式未设置的错
            ShowError ERR_StationItemFormulaNotSet
        Else
            '无行包票价项
            ShowError ERR_AcceptSheetNoPriceIiem
        End If
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
    Exit Sub
CalculatePriceErr:
    Call err.Raise(err.Number, "AcceptSheet:CalculatePrice Method", err.Description)
End Sub



'##ModelId=3E04A7B8037E
Public Property Get TotalPrice() As Double
    Dim Count As Integer
    Dim i As Integer
    m_dbTotalPrice = 0
    Count = ArrayLength(m_atPriceItems)
    For i = 1 To Count
        m_dbTotalPrice = m_dbTotalPrice + m_atPriceItems(i).PriceValue
    Next
   Let TotalPrice = m_dbTotalPrice
End Property

'##ModelId=3E04A7B801D9
Public Property Get OperateTime() As Date

    
OperateTime = m_dtOperateTime
End Property

'##ModelId=3E04A7B80034
Public Property Get Operator() As String
'    If m_szOperator = "" Then
'        m_szOperator = m_oActiveUser.UserName
'    End If
Operator = m_szOperator
End Property

'##ModelId=3E04A7B70282
Public Property Get OperatorID() As String
'    If m_szOperatorID = "" Then
'        m_szOperatorID = m_oActiveUser.UserID
'    End If
    
OperatorID = m_szOperatorID
End Property

'##ModelId=3E04A7B700DD
Public Property Get PickerAddress() As String
   Let PickerAddress = m_szPickerAddress
End Property

'##ModelId=3E04A7B6021C
Public Property Let PickerAddress(ByVal Value As String)
m_szPickerAddress = Value
End Property

'##ModelId=3E04A7B60082
Public Property Get PickerPhone() As String
   Let PickerPhone = m_szPickerPhone
End Property

'##ModelId=3E04A7B501CB
Public Property Let PickerPhone(ByVal Value As String)
m_szPickerPhone = Value
End Property

'##ModelId=3E04A7B5003A
Public Property Get PickType() As String
   Let PickType = m_szPickType
End Property

'##ModelId=3E04A7B4018D
Public Property Let PickType(ByVal Value As String)
m_szPickType = Value
End Property

'##ModelId=3E04A7B303E5
Public Property Get Picker() As String
Picker = m_szPicker
End Property

'##ModelId=3E04A7B3015A
Public Property Let Picker(ByVal Value As String)
m_szPicker = Value
End Property

'##ModelId=3E04A7B203C5
Public Property Get Shipper() As String
Shipper = m_szShipper
End Property

'##ModelId=3E04A7B20144
Public Property Let Shipper(ByVal Value As String)
m_szShipper = Value
End Property

'##ModelId=3E04A7B103BA
Public Property Get StatusString() As String

    Select Case m_nStatus
        Case 0
            m_szStatusString = "正常待签发"
        Case 1
            m_szStatusString = "废票"
        Case 2
            m_szStatusString = "退票"
        Case 3
            m_szStatusString = "已签发"
    End Select
        
StatusString = m_szStatusString
End Property

'##ModelId=3E04A7B1025B
Public Property Get Status() As Integer

        
Status = m_nStatus
End Property

'##ModelId=3E04A7B100E9
Public Property Get OverNumber() As Integer
   Let OverNumber = m_nOverNumber
End Property

'##ModelId=3E04A7B0026E
Public Property Let OverNumber(ByVal Value As Integer)
m_nOverNumber = Value
End Property

'##ModelId=3E04A7B00105
Public Property Get Number() As Integer
   Let Number = m_nNumber
End Property

'##ModelId=3E04A7AF0294
Public Property Let Number(ByVal Value As Integer)
m_nNumber = Value
End Property

'##ModelId=3E04A7AF0136
Public Property Get EndLabelID() As String
   Let EndLabelID = m_szEndLabelID
End Property

'##ModelId=3E04A7AE02CF
Public Property Let EndLabelID(ByVal Value As String)
m_szEndLabelID = Value
End Property

'##ModelId=3E04A7AE0171
Public Property Get StartLabelID() As String
   Let StartLabelID = m_szStartLabelID
End Property

'##ModelId=3E04A7AD0314
Public Property Let StartLabelID(ByVal Value As String)
m_szStartLabelID = Value
End Property

'##ModelId=3E04A7AD01BF
Public Property Get ActWeight() As Double
   Let ActWeight = m_dbActWeight
End Property

'##ModelId=3E04A7AC0376
Public Property Let ActWeight(ByVal Value As Double)
m_dbActWeight = Value
End Property

'##ModelId=3E04A7AC022C
Public Property Get CalWeight() As Double
   Let CalWeight = m_dbCalWeight
End Property

'##ModelId=3E04A7AB03E3
Public Property Let CalWeight(ByVal Value As Double)
m_dbCalWeight = Value
End Property

'##ModelId=3E04A7AB02A3
Public Property Get BusDate() As Date

BusDate = m_dtBusDate
End Property

'##ModelId=3E04A7AB0086
Public Property Let BusDate(ByVal Value As Date)
m_dtBusDate = Value
End Property

'##ModelId=3E04A7AA032D
Public Property Get BusID() As String
   Let BusID = m_szBusID
End Property

'##ModelId=3E04A7AA0125
Public Property Let BusID(ByVal Value As String)

m_szBusID = Value

End Property

'##ModelId=3E04A7AA000C
Public Property Get Mileage() As Double
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset

    
Mileage = m_dbMileage
End Property

'##ModelId=3E04A7A902D2
Public Property Get DesStationName() As String
    
DesStationName = m_szDesStationName
End Property

'##ModelId=3E04A7A901A5
Public Property Get DesStationID() As String
   Let DesStationID = m_szDesStationID
End Property

'##ModelId=3E04A7A8038F
Public Property Let DesStationID(ByVal Value As String)
m_szDesStationID = Value
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    If m_szDesStationID <> "" Then
        szSql = "SELECT station_name FROM  station_info WHERE station_id =" & TransFieldValueToString(m_szDesStationID)
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount = 0 Then ShowError ERR_AcceptSheetNoStationName
        m_szDesStationName = FormatDbValue(rsTemp!station_name)
        
        szSql = " SELECT DISTINCT  end_station_mileage FROM " & cszTableRouteSection & " " _
            & " WHERE end_station_id= " & TransFieldValueToString(m_szDesStationID) _
            & " AND sell_station_id = " & TransFieldValueToString(m_oActiveUser.SellStationID)
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount = 0 Then
            ShowError ERR_AcceptSheetNoStationName
            Exit Property
        End If
        m_dbMileage = FormatDbValue(rsTemp!end_station_mileage)
    End If
    
End Property

'##ModelId=3E04A7A80280
Public Property Get StartStationName() As String

    
StartStationName = m_szStartStationName
End Property

'##ModelId=3E04A7A8015E
Public Property Get StartStationID() As String
   Let StartStationID = m_szStartStationID
End Property

'##ModelId=3E04A7A7035B
Public Property Let StartStationID(ByVal Value As String)
    
End Property

'##ModelId=3E04A7A70243
Public Property Get AcceptType() As String
   Let AcceptType = m_szAcceptType
End Property

'##ModelId=3E04A7A7006C
Public Property Let AcceptType(ByVal Value As String)
m_szAcceptType = Value
End Property

'##ModelId=3E04A7A60346
Public Property Get SheetID() As String
   Let SheetID = m_szSheetID
End Property


'##ModelId=3E04A7A6016F
Public Property Let SheetID(ByVal Value As String)
m_szSheetID = Value
End Property
'行包名称
Public Property Let LuggageName(ByVal Value As String)
m_szLuggageName = Value
End Property

Public Property Get LuggageName() As String
LuggageName = m_szLuggageName
End Property
'行包托运人名称
Public Property Let LuggageShipperPhone(ByVal Value As String)
m_szLuggageShipperPhone = Value
End Property

Public Property Get LuggageShipperPhone() As String
LuggageShipperPhone = m_szLuggageShipperPhone
End Property

'车牌
Public Property Let LicenseTagNo(ByVal Value As String)
    m_szLicenseTagNo = Value
End Property

Public Property Get LicenseTagNo() As String
    LicenseTagNo = m_szLicenseTagNo
End Property

'包装
Public Property Let Pack(ByVal Value As String)
    m_szPack = Value
End Property

Public Property Get Pack() As String
    Pack = m_szPack
End Property



'计算的行包的运费价
Public Property Let CalBasePrice(ByVal Value As Double)
    m_dbCalBasePrice = Value
End Property

Public Property Get CalBasePrice() As Double
    CalBasePrice = m_dbCalBasePrice
End Property


'运输单号
Public Property Let TransTicketID(ByVal Value As String)
    m_szTransTicketID = Value
End Property

Public Property Get TransTicketID() As String
    TransTicketID = m_szTransTicketID
End Property


'保险单号
Public Property Let InsuranceID(ByVal Value As String)
    m_szInsuranceID = Value
End Property

Public Property Get InsuranceID() As String
    InsuranceID = m_szInsuranceID
End Property

'应结运费比率
Public Property Let SettleRatio(ByVal Value As Double)
    m_dbSettleRatio = Value
End Property

Public Property Get SettleRatio() As Double
    SettleRatio = m_dbSettleRatio
End Property

'应结运费
Public Property Let SettlePrice(ByVal Value As Double)
    m_dbSettlePrice = Value
End Property

Public Property Get SettlePrice() As Double
    SettlePrice = m_dbSettlePrice
End Property

'备注1
Public Property Let Annotation1(ByVal Value As String)
    m_szAnnotation1 = Value
End Property

Public Property Get Annotation1() As String
    Annotation1 = m_szAnnotation1
End Property

'备注2
Public Property Let Annotation2(ByVal Value As String)
    m_szAnnotation2 = Value
End Property

Public Property Get Annotation2() As String
    Annotation2 = m_szAnnotation2
End Property

'受理站
Public Property Get SellStationID() As String
    SellStationID = m_szSellStationID
End Property

