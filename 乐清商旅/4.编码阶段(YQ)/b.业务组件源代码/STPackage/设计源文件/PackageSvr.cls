VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PackageSvr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'行包到达服务
Option Base 0
Option Explicit


'以下定义枚举类型,公用部分
'======================================================
'中间层行包类的类号
Public Enum ECOMID_Package
    CID_Package = 60
    CID_PackageParam = 61
    CID_PackageSvr = 62
End Enum

'行包类的错误的起始号(实际也是其所有资源的起始号)
Public Enum EClassErrBegin_Package
    ERR_Package = CID_Package * cnMidErrStep + cnMidErrBegin '34000
    ERR_PackageParam = CID_PackageParam * cnMidErrStep + cnMidErrBegin  '34400
    ERR_PackageSvr = CID_PackageSvr * cnMidErrStep + cnMidErrBegin  '34800
End Enum





Public Enum EPackageStatus
    EPS_Normal = 0          '正常待提
    EPS_Picked = 1          '已提
    EPS_Cancel = 2          '已作废
End Enum
    

'本模块私有定义部分
'==========================================================================
'错误号
'Private Enum EErrPackageSvr    '30800开始
'End Enum
'权限号

Private Enum ERightPackageSvr
    RIGHT_PackageSvr = ERR_PackageSvr + cnMidRightBegin              '30900开始
    RIGHT_PackageSvrAccept = RIGHT_PackageSvr + cnMidRightStep * 1    '行包单受理
    RIGHT_PackageSvrPick = RIGHT_PackageSvr + cnMidRightStep * 2    '行包单提取
    RIGHT_PackageSvrCancel = RIGHT_PackageSvr + cnMidRightStep * 3    '行包单作废
    RIGHT_PackageSvrStat = RIGHT_PackageSvr + cnMidRightStep * 4    '行包到达统计
End Enum
'以下定义本模块变量
'------------------------------------------
Private m_oActiveUser As ActiveUser



'初始化活动对象和当前售票站点
'##ModelId=3E03A6A700B8
Public Sub Init(poAUser As ActiveUser)
    On Error GoTo InitErr
    Set m_oActiveUser = poAUser
    

    Exit Sub
InitErr:
    Call RaiseError(err.Number, "LuggageSvr:Init Method", err.Description)
End Sub

'行包受理
'##ModelId=3E03AB1D0088
Public Sub AcceptPackage(poPackage As Package)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_PackageSvr
    AssertHaveRight m_oActiveUser, RIGHT_PackageSvrAccept
'
    On Error GoTo AcceptLuggageErr
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    
    
    oDb.ConnectionString = GetConnectionStr(cszPackage)
    
    '自动生成自编号，衢州改
    szSql = " SELECT ISNULL(MAX(package_id),'') package_id FROM package_arrived_info WHERE year(operation_time) = " & Year(Date) & " AND month(operation_time) = " & Month(Date)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 1 And FormatDbValue(rsTemp!package_id) <> "" Then
        poPackage.PackageID = Format(Date, "yy") & Format(Date, "MM") & Format(Val(FormatDbValue(rsTemp!package_id)) Mod 10 ^ 6 + 1, "000000")
    Else
        poPackage.PackageID = Format(Date, "yy") & Format(Date, "MM") & Format("1", "000000")
    End If
        
    
    szSql = " SELECT * FROM package_arrived_info WHERE package_id = " & TransFieldValueToString(poPackage.PackageID)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        err.Raise ERR_PackageSvr, "", "自编号已存在,不能新增"
    End If
    oDb.BeginTrans
    
'    szSql = "SELECT ISNULL(MAX(package_id),0) package_id FROM package_arrived_info"
'    Set rsTmp = oDb.Execute(szSql)
'    Dim lTmpID As Long
'    lTmpID = FormatDbValue(rsTmp!package_id) + 1
    
        'INSERT 受理信息表
    szSql = "INSERT package_arrived_info (package_id,sell_station_id,package_name,area_type,pack,arrive_time,weight,start_station_name,license_tag_no,package_number,save_position," _
                                        & "status,send_name,send_unit,send_phone,pick_type,picker,picker_unit,picker_address,picker_phone,operator,user_id," _
                                        & "operation_time,loader,transit_charge,memo) " & _
            " VALUES(" & TransFieldValueToString(poPackage.PackageID) & _
            "," & TransFieldValueToString(poPackage.SellStationID) & _
            "," & TransFieldValueToString(poPackage.PackageName) & _
            "," & TransFieldValueToString(poPackage.AreaType) & _
            "," & TransFieldValueToString(poPackage.PackType) & _
            "," & TransFieldValueToString(poPackage.ArrivedTime) & _
            "," & TransFieldValueToString(poPackage.CalWeight) & _
            "," & TransFieldValueToString(poPackage.StartStationName) & _
            "," & TransFieldValueToString(poPackage.LicenseTagNo) & _
            "," & TransFieldValueToString(poPackage.PackageNumber) & _
            "," & TransFieldValueToString(poPackage.SavePosition) & _
            "," & TransFieldValueToString(EPS_Normal) & _
            "," & TransFieldValueToString(poPackage.Shipper) & _
            "," & TransFieldValueToString(poPackage.ShipperUnit) & _
            "," & TransFieldValueToString(poPackage.ShipperPhone) & _
            "," & TransFieldValueToString(poPackage.PickType) & _
            "," & TransFieldValueToString(poPackage.Picker) & _
            "," & TransFieldValueToString(poPackage.PickerUnit) & _
            "," & TransFieldValueToString(poPackage.PickerAddress)
        szSql = szSql & _
            "," & TransFieldValueToString(poPackage.PickerPhone) & _
            "," & TransFieldValueToString(poPackage.Operator) & _
            "," & TransFieldValueToString(m_oActiveUser.UserID) & _
            "," & TransFieldValueToString(Now) & _
            "," & TransFieldValueToString(poPackage.Loader) & _
            "," & TransFieldValueToString(poPackage.TransitCharge) & _
            "," & TransFieldValueToString(poPackage.Remark) & ")"
        
    oDb.Execute szSql
    
        
'    poPackage.PackageID = lTmpID
    oDb.CommitTrans
    
    Set oDb = Nothing
    Exit Sub
AcceptLuggageErr:
    oDb.RollbackTrans
    Call RaiseError(err.Number, "PackageSvr:AcceptPackage Method", err.Description)
End Sub

Public Sub PickPackage(poPackage As Package)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_PackageSvr
    AssertHaveRight m_oActiveUser, RIGHT_PackageSvrPick
'
    On Error GoTo ErrHandle
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTmp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszPackage)
    'sheet_id为唯一
    szSql = "SELECT * FROM package_arrived_info WHERE package_id=" & TransFieldValueToString(poPackage.PackageID) & " OR sheet_id=" & TransFieldValueToString(poPackage.SheetID)
    Set rsTmp = oDb.Execute(szSql)
    If rsTmp.RecordCount <> 1 Then RaiseError ERR_AcceptPackageWithout
    
    oDb.BeginTrans
    
    szSql = "UPDATE package_arrived_info SET picker_credit=" & TransFieldValueToString(poPackage.PickerCreditID) & _
                                        ",user_id=" & TransFieldValueToString(m_oActiveUser.UserID) & _
                                        ",status=" & TransFieldValueToString(EPS_Picked) & _
                                        ",sheet_id=" & TransFieldValueToString(poPackage.SheetID) & _
                                        ",pick_time=" & TransFieldValueToString(poPackage.PickTime) & _
                                        ",load_charge=" & TransFieldValueToString(poPackage.LoadCharge) & _
                                        ",keep_charge=" & TransFieldValueToString(poPackage.KeepCharge) & _
                                        ",move_charge=" & TransFieldValueToString(poPackage.MoveCharge) & _
                                        ",send_charge=" & TransFieldValueToString(poPackage.SendCharge) & _
                                        ",other_charge=" & TransFieldValueToString(poPackage.OtherCharge) & _
                                        ",drawer=" & TransFieldValueToString(poPackage.Drawer) & _
                                        ",drawer_phone=" & TransFieldValueToString(poPackage.DrawerPhone) & _
                                        ",memo=" & TransFieldValueToString(poPackage.Remark) & _
            " WHERE package_id=" & TransFieldValueToString(poPackage.PackageID)
    oDb.Execute szSql
    
    oDb.CommitTrans
    
    poPackage.Status = EPS_Picked

    Set oDb = Nothing
    Exit Sub
ErrHandle:
    oDb.RollbackTrans
    Call RaiseError(err.Number, "PackageSvr:PickPackage Method", err.Description)
End Sub


Public Sub CancelPackage(poPackage As Package)
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_PackageSvr
    AssertHaveRight m_oActiveUser, RIGHT_PackageSvrCancel
'
    On Error GoTo ErrHandle
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTmp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszPackage)
    szSql = "SELECT status FROM package_arrived_info WHERE package_id=" & TransFieldValueToString(poPackage.PackageID)
    Set rsTmp = oDb.Execute(szSql)
    If rsTmp.RecordCount = 0 Then RaiseError ERR_AcceptPackageWithout
    Dim eStatus As EPackageStatus
    eStatus = FormatDbValue(rsTmp!Status)
    Select Case eStatus
        Case EPS_Normal
            eStatus = EPS_Cancel
        Case EPS_Picked
            eStatus = EPS_Normal
    End Select
    
    oDb.BeginTrans
            
    szSql = "UPDATE package_arrived_info SET status=" & TransFieldValueToString(eStatus) & _
                                        ",memo=" & TransFieldValueToString(poPackage.Remark) & _
                                        ",cancel_time = getdate() " & _
                                        ",cancel_user_id= " & TransFieldValueToString(m_oActiveUser.UserID) & _
            " WHERE package_id=" & TransFieldValueToString(poPackage.PackageID)
    oDb.Execute szSql
    
    oDb.CommitTrans
    
    poPackage.Status = EPS_Cancel

    Set oDb = Nothing
    Exit Sub
ErrHandle:
    oDb.RollbackTrans
    Call RaiseError(err.Number, "PackageSvr:CancelPackage Method", err.Description)
End Sub

'得到最新的行包签发单号
Public Function GetLastSheetID(UserID As String) As String
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszPackage)
    
    szSql = "SELECT MAX(sheet_id) AS sheet_id FROM package_arrived_info " _
    & " WHERE pick_time=(" _
    & " SELECT MAX(pick_time) FROM package_arrived_info WHERE " _
    & " user_id='" & UserID & "') " _
    & " AND user_id='" & UserID & "' "
    
    Set rsTemp = oDb.Execute(szSql)
    Dim oParam As PackageParam
    If rsTemp.RecordCount = 1 Then
        GetLastSheetID = FormatDbValue(rsTemp!sheet_id)
        If Trim(GetLastSheetID) = "" Then
            Set oParam = New PackageParam
            oParam.Init m_oActiveUser
            GetLastSheetID = String(oParam.SheetIDNumberLen, "0")
        End If
    Else
        Set oParam = New PackageParam
        oParam.Init m_oActiveUser
        GetLastSheetID = String(oParam.SheetIDNumberLen, "0")
    End If
    Set oDb = Nothing
End Function
'查询行包
Public Function ListPackageRS(pszSearch As String) As Recordset
On Error GoTo ErrHandle
    AssertActiveUserValid m_oActiveUser, ERR_Package
    
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszPackage)
    Dim szSql As String
    
    szSql = "SELECT * FROM package_arrived_info WHERE 1=1 "
    
    If m_oActiveUser.SellStationID <> "" Then
        szSql = szSql & " AND sell_station_id = " & TransFieldValueToString(m_oActiveUser.SellStationID)
    End If
    
    If Trim(pszSearch) <> "" Then
        szSql = szSql & " AND " & pszSearch
    End If
    
    Set rsTemp = oDb.Execute(szSql)
    
    Set ListPackageRS = rsTemp
    Exit Function
ErrHandle:
    Call RaiseError(err.Number, "AcceptSheet:Identify Method", err.Description)
End Function
'根据到达时间进行统计
Public Function StatPackageByArrivedRS(pdtStartDate As Date, pdtEndDate As Date, pszTotalBy As String, pszOtherSearch As String)
On Error GoTo ErrHandle
    AssertActiveUserValid m_oActiveUser, ERR_PackageSvr
    AssertHaveRight m_oActiveUser, RIGHT_PackageSvrStat
    
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szSellStationCondition As String
    oDb.ConnectionString = GetConnectionStr(cszPackage)
    
    Dim dbTransParam As Double
    szSql = "SELECT parameter_value FROM system_param_info WHERE parameter_name=" & TransFieldValueToString("PackageTransitRatio")
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        dbTransParam = Val(FormatDbValue(rsTemp!parameter_value))
    End If
    
    If m_oActiveUser.SellStationID <> "" Then
        szSellStationCondition = " AND p.sell_station_id = " & TransFieldValueToString(m_oActiveUser.SellStationID)
    End If
    
    If pszTotalBy = "t.user_name" Then
        szSql = "SELECT " & pszTotalBy & ",'' sheet_rang, SUM(package_number) package_nums,SUM(transit_charge) transit_charge,SUM(transit_charge)*" & dbTransParam & " transit_man_charge,SUM(load_charge) load_charge,SUM(keep_charge) keep_charge,SUM(move_charge) move_charge," & _
                        "SUM(send_charge) send_charge,SUM(other_charge) other_charge " & _
                " ,SUM(load_charge)+SUM(keep_charge)+SUM(move_charge)+SUM(send_charge)+SUM(other_charge)+SUM(transit_charge) total_charge " & _
                " FROM package_arrived_info p, user_info t " & _
                " WHERE p.user_id = t.user_id AND arrive_time>=" & TransFieldValueToString(pdtStartDate) & _
                " AND arrive_time<" & TransFieldValueToString(DateAdd("d", 1, pdtEndDate)) & szSellStationCondition
    Else
        szSql = "SELECT " & pszTotalBy & ", '' sheet_rang, SUM(package_number) package_nums,SUM(transit_charge) transit_charge,SUM(transit_charge)*" & dbTransParam & " transit_man_charge,SUM(load_charge) load_charge,SUM(keep_charge) keep_charge,SUM(move_charge) move_charge," & _
                        "SUM(send_charge) send_charge,SUM(other_charge) other_charge " & _
                " ,SUM(load_charge)+SUM(keep_charge)+SUM(move_charge)+SUM(send_charge)+SUM(other_charge)+SUM(transit_charge) total_charge " & _
                " FROM package_arrived_info p " & _
                " WHERE arrive_time>=" & TransFieldValueToString(pdtStartDate) & _
                " AND arrive_time<" & TransFieldValueToString(DateAdd("d", 1, pdtEndDate)) & szSellStationCondition
    End If
    If pszOtherSearch <> "" Then
        szSql = szSql & " AND " & pszOtherSearch
    End If
    szSql = szSql & " GROUP BY " & pszTotalBy
    Set rsTemp = oDb.Execute(szSql)
    
    Set StatPackageByArrivedRS = rsTemp
    Exit Function
ErrHandle:
    Call RaiseError(err.Number, "AcceptSheet:StatPackageByArrivedRS Method", err.Description)
End Function

'根据提取时间进行统计
Public Function StatPackageByPickedRS(pdtStartDate As Date, pdtEndDate As Date, pszTotalBy As String, pszOtherSearch As String)
On Error GoTo ErrHandle
    AssertActiveUserValid m_oActiveUser, ERR_Package
    AssertHaveRight m_oActiveUser, RIGHT_PackageSvrStat
    
    Dim rsTemp As Recordset
    Dim szSellStationCondition As String
    Dim oDb As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszPackage)
    
    '得到代收运费手续费
    Dim dbTransParam As Double
    szSql = "SELECT parameter_value FROM system_param_info WHERE parameter_name=" & TransFieldValueToString("PackageTransitRatio")
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        dbTransParam = Val(FormatDbValue(rsTemp!parameter_value))
    End If
    
    If m_oActiveUser.SellStationID <> "" Then
        szSellStationCondition = " AND p.sell_station_id = " & TransFieldValueToString(m_oActiveUser.SellStationID)
    End If
    
    If pszTotalBy = "t.user_name" Then
        szSql = "SELECT " & pszTotalBy & ", MIN(sheet_id) +'--'+MAX(sheet_id) sheet_rang, SUM(package_number) package_nums,SUM(transit_charge) transit_charge,SUM(transit_charge)*" & dbTransParam & " transit_man_charge, SUM(load_charge) load_charge,SUM(keep_charge) keep_charge,SUM(move_charge) move_charge," & _
                        "SUM(send_charge) send_charge,SUM(other_charge) other_charge" & _
                " ,SUM(load_charge)+SUM(keep_charge)+SUM(move_charge)+SUM(send_charge)+SUM(other_charge)+SUM(transit_charge) total_charge " & _
                " FROM package_arrived_info p, user_info t" & _
                " WHERE p.user_id = t.user_id AND pick_time>= " & TransFieldValueToString(pdtStartDate) & _
                " AND pick_time<" & TransFieldValueToString(DateAdd("d", 1, pdtEndDate)) & szSellStationCondition
    Else
        szSql = "SELECT " & pszTotalBy & ",'' sheet_rang, SUM(package_number) package_nums,SUM(transit_charge) transit_charge,SUM(transit_charge)*" & dbTransParam & " transit_man_charge, SUM(load_charge) load_charge,SUM(keep_charge) keep_charge,SUM(move_charge) move_charge," & _
                    "SUM(send_charge) send_charge,SUM(other_charge) other_charge" & _
            " ,SUM(load_charge)+SUM(keep_charge)+SUM(move_charge)+SUM(send_charge)+SUM(other_charge)+SUM(transit_charge) total_charge " & _
            " FROM package_arrived_info p" & _
            " WHERE pick_time>=" & TransFieldValueToString(pdtStartDate) & _
            " AND pick_time<" & TransFieldValueToString(DateAdd("d", 1, pdtEndDate)) & szSellStationCondition
    End If
    If pszOtherSearch <> "" Then
        szSql = szSql & " AND " & pszOtherSearch
    End If
    szSql = szSql & " GROUP BY " & pszTotalBy
    Set rsTemp = oDb.Execute(szSql)
    
    Set StatPackageByPickedRS = rsTemp
    Exit Function
ErrHandle:
    Call RaiseError(err.Number, "AcceptSheet:StatPackageByArrivedRS Method", err.Description)
End Function

'明细查询
Public Function GetArrivedPackageRS(pdtStartDate As Date, pdtEndDate As Date, pszOrderBy As String, pszOtherSearch As String)
On Error GoTo ErrHandle
    AssertActiveUserValid m_oActiveUser, ERR_Package
    AssertHaveRight m_oActiveUser, RIGHT_PackageSvrStat
    
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszPackage)
    
    
    szSql = "SELECT * , CASE p.status WHEN 0 THEN '未提' WHEN 1 THEN '已提' WHEN 2 THEN '已废' ELSE '其他' END status_name " & _
            " FROM package_arrived_info p, user_info t " & _
            " WHERE pick_time>=" & TransFieldValueToString(pdtStartDate) & _
            " AND pick_time<" & TransFieldValueToString(DateAdd("d", 1, pdtEndDate)) & _
            " AND t.user_id = p.user_id "
            
    If m_oActiveUser.SellStationID <> "" Then
        szSql = szSql & " AND p.sell_station_id = " & TransFieldValueToString(m_oActiveUser.SellStationID)
    End If
    
    If pszOtherSearch <> "" Then
        szSql = szSql & " AND " & pszOtherSearch
    End If
    If pszOrderBy <> "" Then
        szSql = szSql & " order BY " & pszOrderBy
    End If
    Set rsTemp = oDb.Execute(szSql)
    
    Set GetArrivedPackageRS = rsTemp
    Exit Function
ErrHandle:
    Call RaiseError(err.Number, "AcceptSheet:GetArrivedPackageRS Method", err.Description)
End Function
'重打单据,将原记录生成一条新记录
'##ModelId=3E03AB1D0088
Public Function ChangeSheetID(pszOldSheetID As String, pszNewSheetID As String, Optional plPackageID As Long) As Long
    '验证活动用户及权限
    AssertActiveUserValid m_oActiveUser, ERR_PackageSvr
    AssertHaveRight m_oActiveUser, RIGHT_PackageSvrCancel
'
    On Error GoTo AcceptLuggageErr
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTmp As Recordset
    
    
    oDb.ConnectionString = GetConnectionStr(cszPackage)
    
    oDb.BeginTrans
    szSql = "SELECT status FROM package_arrived_info WHERE sheet_id=" & TransFieldValueToString(pszOldSheetID)
    Set rsTmp = oDb.Execute(szSql)
    If rsTmp.RecordCount = 0 Then RaiseError ERR_AcceptPackageWithout
    If FormatDbValue(rsTmp!Status) = EPS_Cancel Then RaiseError ERR_AcceptPackageCanced
    
    szSql = "SELECT status FROM package_arrived_info WHERE sheet_id=" & TransFieldValueToString(pszNewSheetID)
    Set rsTmp = oDb.Execute(szSql)
    If rsTmp.RecordCount > 0 Then RaiseError ERR_PackageSheetExist
    
    Dim lTmpID As Long
    
'    If plPackageID = 0 Then
'        szSql = "SELECT ISNULL(MAX(package_id),0) package_id FROM package_arrived_info"
'        Set rsTmp = oDb.Execute(szSql)
'        lTmpID = FormatDbValue(rsTmp!package_id) + 1
'    Else
'        lTmpID = plPackageID
'    End If

    '以上是温州的重打方式，以下是衢州（自动生成自编号）的方式
    szSql = " SELECT ISNULL(MAX(package_id),'') package_id FROM package_arrived_info WHERE year(operation_time) = " & Year(Date) & " AND month(operation_time) = " & Month(Date)
    Set rsTmp = oDb.Execute(szSql)
    If rsTmp.RecordCount = 1 And FormatDbValue(rsTmp!package_id) <> "" Then
        lTmpID = Format(Date, "yy") & Format(Date, "MM") & Format(Val(FormatDbValue(rsTmp!package_id)) Mod 10 ^ 6 + 1, "000000")
    Else
        lTmpID = Format(Date, "yy") & Format(Date, "MM") & Format("1", "000000")
    End If
    
        'INSERT 受理信息表
    szSql = "INSERT package_arrived_info (package_id,sell_station_id,package_name,area_type,pack,arrive_time,weight,start_station_name,license_tag_no,package_number,save_position,sheet_id," _
                                        & "status,send_name,send_unit,send_phone,pick_type,picker,picker_unit,picker_address,picker_phone,picker_credit,pick_time,operator,user_id," _
                                        & "operation_time,loader,transit_charge,load_charge,keep_charge,move_charge,send_charge,other_charge,drawer,drawer_phone,memo) " & _
            " SELECT " & TransFieldValueToString(lTmpID) & ",sell_station_id,package_name,area_type,pack,arrive_time,weight,start_station_name,license_tag_no,package_number,save_position," & TransFieldValueToString(pszNewSheetID) & "," _
                                        & "status,send_name,send_unit,send_phone,pick_type,picker,picker_unit,picker_address,picker_phone,picker_credit,pick_time,operator,user_id," _
                                        & "getdate(),loader,transit_charge,load_charge,keep_charge,move_charge,send_charge,other_charge,drawer,drawer_phone,memo " & _
            " FROM package_arrived_info WHERE sheet_id=" & TransFieldValueToString(pszOldSheetID)
    oDb.Execute szSql
    szSql = "UPDATE  package_arrived_info SET status=" & TransFieldValueToString(EPS_Cancel) & _
            " ,cancel_user_id=" & TransFieldValueToString(m_oActiveUser.UserID) & _
            " ,cancel_time=getdate()" & _
            " WHERE sheet_id=" & TransFieldValueToString(pszOldSheetID)
    oDb.Execute szSql
        
    ChangeSheetID = lTmpID
    oDb.CommitTrans
    
    Set oDb = Nothing
    Exit Function
AcceptLuggageErr:
    oDb.RollbackTrans
    Call RaiseError(err.Number, "PackageSvr:ChangeSheetID Method", err.Description)
End Function




Public Function AcceptEveryDaySellDetail(UserID As String, StartDate As Date, EndDate As Date) As Recordset
'行包员受理的明细
    AssertActiveUserValid m_oActiveUser, ERR_PackageSvr
    AssertHaveRight m_oActiveUser, RIGHT_PackageSvrStat
    
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr("")
    
    '状态为0的也要统计，因为有提货错，取消提货（退提）的
    szSql = "SELECT * ,transit_charge+load_charge+keep_charge+move_charge+send_charge+other_charge price_total FROM package_arrived_info   WHERE  " _
    & " user_id='" & UserID & "' AND " _
    & " pick_time>='" & ToDBDateTime(StartDate) & "' AND " _
    & " pick_time<'" & ToDBDateTime(EndDate) & "'" _
    & " ORDER BY sheet_id "
'    & " operation_time>='" & ToDBDateTime(StartDate) & "' AND " _
'    & " operation_time<'" & ToDBDateTime(EndDate) & "'" _
'2005-7-13 lyq changed
'    & " AND status <> " & TransFieldValueToString(ST_LuggageAcceptCancel) & "  ORDER BY a.luggage_id"
      
    
    Set rsTemp = oDb.Execute(szSql)
    Set AcceptEveryDaySellDetail = rsTemp
    Set oDb = Nothing
    
End Function


Public Function AcceptEveryDaySellTotal(UserID As String, StartDate As Date, EndDate As Date) As Double
'行包员受理的明细
    AssertActiveUserValid m_oActiveUser, ERR_PackageSvr
    AssertHaveRight m_oActiveUser, RIGHT_PackageSvrStat
    
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr("")
    
    szSql = "SELECT sum(transit_charge+load_charge+keep_charge+move_charge+send_charge+other_charge) price_total FROM package_arrived_info a  WHERE " _
    & " user_id='" & UserID & "' AND " _
    & " pick_time>='" & ToDBDateTime(StartDate) & "' AND " _
    & " pick_time<'" & ToDBDateTime(EndDate) & "'"

'& " status <>" & TransFieldValueToString(EPS_Normal) & " and user_id='" & UserID & "' AND " _
'    & " AND status <> " & TransFieldValueToString(ST_LuggageAcceptCancel)

      
    
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        AcceptEveryDaySellTotal = FormatDbValue(rsTemp!price_total)
    End If
    Set oDb = Nothing
    
End Function


'数组相应为废
'行包员的退废签信息
Public Function AcceptEveryDayAnotherThing(UserID As String, StartDate As Date, EndDate As Date) As Double()
    AssertActiveUserValid m_oActiveUser, ERR_PackageSvr
    AssertHaveRight m_oActiveUser, RIGHT_PackageSvrStat
    
    Dim adbResult(1 To 5, 1 To 3) As Double
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    
    oDb.ConnectionString = GetConnectionStr("")
    szSql = "SELECT COUNT(*) AS countx ,SUM(transit_charge+load_charge+keep_charge+move_charge+send_charge+other_charge) AS total_total_price FROM " _
    & " package_arrived_info a  WHERE " _
    & " status = 2 AND " _
    & " a.cancel_user_id ='" & UserID & "' AND " _
    & " cancel_time>='" & ToDBDateTime(StartDate) & "' AND " _
    & " cancel_time<'" & ToDBDateTime(EndDate) & "'"
    
    Set rsTemp = oDb.Execute(szSql)
    adbResult(1, 1) = rsTemp!countx
    adbResult(1, 2) = FormatDbValue(rsTemp!total_total_price)
    
    '取消提货统计(退提，当日)
    szSql = "SELECT COUNT(*) AS countx ,SUM(transit_charge+load_charge+keep_charge+move_charge+send_charge+other_charge) AS total_total_price FROM " _
    & " package_arrived_info a  WHERE " _
    & " status = 0 AND " _
    & " pick_time>='" & ToDBDateTime(StartDate) & "' AND " _
    & " pick_time<'" & ToDBDateTime(EndDate) & "' AND " _
    & " a.cancel_user_id ='" & UserID & "' AND " _
    & " cancel_time>='" & ToDBDateTime(StartDate) & "' AND " _
    & " cancel_time<'" & ToDBDateTime(EndDate) & "'"
    
    Set rsTemp = oDb.Execute(szSql)
    adbResult(2, 1) = rsTemp!countx
    adbResult(2, 2) = FormatDbValue(rsTemp!total_total_price)
    
    '隔日取消提货统计(退提，当日以后)
    szSql = "SELECT COUNT(*) AS countx ,SUM(transit_charge+load_charge+keep_charge+move_charge+send_charge+other_charge) AS total_total_price FROM " _
    & " package_arrived_info a  WHERE " _
    & " status = 0 AND " _
    & " pick_time>='" & ToDBDateTime(StartDate) & "' AND " _
    & " pick_time<'" & ToDBDateTime(EndDate) & "' AND " _
    & " a.cancel_user_id ='" & UserID & "' AND " _
    & " cancel_time>'" & ToDBDateTime(EndDate) & "'"
    
    Set rsTemp = oDb.Execute(szSql)
    adbResult(3, 1) = rsTemp!countx
    adbResult(3, 2) = FormatDbValue(rsTemp!total_total_price)
    
    '隔日取消提货统计(退提，当日以前)
    szSql = "SELECT COUNT(*) AS countx ,SUM(transit_charge+load_charge+keep_charge+move_charge+send_charge+other_charge) AS total_total_price FROM " _
    & " package_arrived_info a  WHERE " _
    & " status = 0 AND " _
    & " pick_time<'" & ToDBDateTime(StartDate) & "' AND " _
    & " pick_time>'" & ToDBDateTime(cdtEmptyDate) & "' AND " _
    & " a.cancel_user_id ='" & UserID & "' AND " _
    & " cancel_time>='" & ToDBDateTime(StartDate) & "' AND " _
    & " cancel_time<'" & ToDBDateTime(EndDate) & "'"
    
    
    Set rsTemp = oDb.Execute(szSql)
    adbResult(4, 1) = rsTemp!countx
    adbResult(4, 2) = FormatDbValue(rsTemp!total_total_price)
    
    '未提过货的作废（登记错的)
    szSql = "SELECT COUNT(*) AS countx ,SUM(transit_charge+load_charge+keep_charge+move_charge+send_charge+other_charge) AS total_total_price FROM " _
    & " package_arrived_info a  WHERE " _
    & " status = 2 AND sheet_id = '' AND " _
    & " a.cancel_user_id ='" & UserID & "' AND " _
    & " cancel_time>='" & ToDBDateTime(StartDate) & "' AND " _
    & " cancel_time<'" & ToDBDateTime(EndDate) & "'"
    
    Set rsTemp = oDb.Execute(szSql)
    adbResult(5, 1) = rsTemp!countx
    adbResult(5, 2) = FormatDbValue(rsTemp!total_total_price)

    AcceptEveryDayAnotherThing = adbResult
    
    Set oDb = Nothing
    
End Function

'车站到达行包营收简报
Public Function StationStat(FromDate As Date, ToDate As Date) As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_PackageSvr
    AssertHaveRight m_oActiveUser, RIGHT_PackageSvrStat
    
    Dim szSql As String
    Dim szSellStationCondition As String
    Dim oDb As New RTConnection
    Dim rsTemp1 As Recordset
    Dim rsTemp2 As Recordset
    Dim rsTemp3 As Recordset
    Dim rsData As New Recordset
    Dim dbNormalAmount, dbPickedAmount, dbCancelAmount As Double
    
    
    If m_oActiveUser.SellStationID <> "" Then
        szSellStationCondition = " AND sell_station_id = " & TransFieldValueToString(m_oActiveUser.SellStationID)
    End If
    
    '提取行包,按提取时间来
    szSql = "select sum(package_number) as normal_number,sum(transit_charge)  as transit_charge,sum(load_charge) as load_charge,sum(keep_charge) as keep_charge,sum(move_charge) as move_charge,sum(send_charge) as send_charge,sum(other_charge) as other_charge from package_arrived_info WHERE " _
    & " status=" & TransFieldValueToString(EPackageStatus.EPS_Picked) & " AND " _
    & " pick_time>='" & ToDBDateTime(FromDate) & "' AND " _
    & " pick_time<='" & ToDBDate(ToDate) & " 23:59:59" & "' " & szSellStationCondition
    oDb.ConnectionString = GetConnectionStr()
    Set rsTemp1 = oDb.Execute(szSql)

    '未提行包,按到达时间来
    szSql = "select sum(package_number) as picked_number,sum(transit_charge)  as transit_charge,sum(load_charge) as load_charge,sum(keep_charge) as keep_charge,sum(move_charge) as move_charge,sum(send_charge) as send_charge,sum(other_charge) as other_charge from package_arrived_info WHERE " _
    & " status=" & TransFieldValueToString(EPackageStatus.EPS_Normal) & " AND " _
    & " arrive_time>='" & ToDBDateTime(FromDate) & "' AND " _
    & " arrive_time<='" & ToDBDate(ToDate) & " 23:59:59" & "' " & szSellStationCondition
    oDb.ConnectionString = GetConnectionStr()
    Set rsTemp2 = oDb.Execute(szSql)
    
    '已废行包,按操作时间来
    szSql = "select sum(package_number) as cancel_number,sum(transit_charge)  as transit_charge,sum(load_charge) as load_charge,sum(keep_charge) as keep_charge,sum(move_charge) as move_charge,sum(send_charge) as send_charge,sum(other_charge) as other_charge from package_arrived_info WHERE " _
    & " status=" & TransFieldValueToString(EPackageStatus.EPS_Cancel) & " AND " _
    & " cancel_time>='" & ToDBDateTime(FromDate) & "' AND " _
    & " cancel_time<='" & ToDBDate(ToDate) & " 23:59:59" & "' " & szSellStationCondition
    oDb.ConnectionString = GetConnectionStr()
    Set rsTemp3 = oDb.Execute(szSql)
    
    
'做出Recordset
    '添加字段
        rsData.Fields.Append "normal_number", adInteger
        rsData.Fields.Append "transit_charge", adCurrency
        rsData.Fields.Append "load_charge", adCurrency
        rsData.Fields.Append "keep_charge", adCurrency
        rsData.Fields.Append "move_charge", adCurrency
        rsData.Fields.Append "send_charge", adCurrency
        rsData.Fields.Append "other_charge", adCurrency
        rsData.Fields.Append "normal_amount", adCurrency

        rsData.Fields.Append "picked_number", adInteger
        rsData.Fields.Append "picked_amount", adCurrency
    
        rsData.Fields.Append "cancel_number", adInteger
        rsData.Fields.Append "cancel_amount", adCurrency
    
        rsData.Fields.Append "total_number", adInteger
        rsData.Fields.Append "total_amount", adCurrency
    
    rsData.Open

     rsData.AddNew
    Do While Not rsTemp1.EOF
        dbNormalAmount = IIf(IsNull(rsTemp1!transit_charge), 0, rsTemp1!transit_charge) + IIf(IsNull(rsTemp1!load_charge), 0, rsTemp1!load_charge) + IIf(IsNull(rsTemp1!keep_charge), 0, rsTemp1!keep_charge) + IIf(IsNull(rsTemp1!move_charge), 0, rsTemp1!move_charge) + IIf(IsNull(rsTemp1!send_charge), 0, rsTemp1!send_charge) + IIf(IsNull(rsTemp1!other_charge), 0, rsTemp1!other_charge)
        rsData("normal_number") = IIf(IsNull(rsTemp1!normal_number), 0, rsTemp1!normal_number)
        rsData("transit_charge") = IIf(IsNull(rsTemp1!transit_charge), 0, rsTemp1!transit_charge)
        rsData("load_charge") = IIf(IsNull(rsTemp1!load_charge), 0, rsTemp1!load_charge)
        rsData("keep_charge") = IIf(IsNull(rsTemp1!keep_charge), 0, rsTemp1!keep_charge)
        rsData("move_charge") = IIf(IsNull(rsTemp1!move_charge), 0, rsTemp1!move_charge)
        rsData("send_charge") = IIf(IsNull(rsTemp1!send_charge), 0, rsTemp1!send_charge)
        rsData("other_charge") = IIf(IsNull(rsTemp1!other_charge), 0, rsTemp1!other_charge)
        rsData("normal_amount") = dbNormalAmount
        rsTemp1.MoveNext
    Loop
    
    Do While Not rsTemp2.EOF
        dbPickedAmount = IIf(IsNull(rsTemp2!transit_charge), 0, rsTemp2!transit_charge) + IIf(IsNull(rsTemp2!load_charge), 0, rsTemp2!load_charge) + IIf(IsNull(rsTemp2!keep_charge), 0, rsTemp2!keep_charge) + IIf(IsNull(rsTemp2!move_charge), 0, rsTemp2!move_charge) + IIf(IsNull(rsTemp2!send_charge), 0, rsTemp2!send_charge) + IIf(IsNull(rsTemp2!other_charge), 0, rsTemp2!other_charge)
        rsData("picked_number") = IIf(IsNull(rsTemp2!picked_number), 0, rsTemp2!picked_number)
        rsData("picked_amount") = dbPickedAmount
        rsTemp2.MoveNext
    Loop
    
    Do While Not rsTemp3.EOF
        dbCancelAmount = IIf(IsNull(rsTemp3!transit_charge), 0, rsTemp3!transit_charge) + IIf(IsNull(rsTemp3!load_charge), 0, rsTemp3!load_charge) + IIf(IsNull(rsTemp3!keep_charge), 0, rsTemp3!keep_charge) + IIf(IsNull(rsTemp3!move_charge), 0, rsTemp3!move_charge) + IIf(IsNull(rsTemp3!send_charge), 0, rsTemp3!send_charge) + IIf(IsNull(rsTemp3!other_charge), 0, rsTemp3!other_charge)
        rsData("cancel_number") = IIf(IsNull(rsTemp3!cancel_number), 0, rsTemp3!cancel_number)
        rsData("cancel_amount") = dbCancelAmount
        rsTemp3.MoveNext
    Loop
    
    rsData("total_number") = rsData("normal_number")
    rsData("total_amount") = dbNormalAmount

        rsData.Update
    
    Set StationStat = rsData
End Function

'车站到达行包营收月报
Public Function StationStatMonth(FromDate As Date, ToDate As Date) As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_PackageSvr
    AssertHaveRight m_oActiveUser, RIGHT_PackageSvrStat
    
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szSellStationCondition As String
    Dim rsTemp As Recordset
    Dim rsData As New Recordset
    Dim j As Integer
    Dim lNumber As Long
    Dim dbAmount As Double
    
    If m_oActiveUser.SellStationID <> "" Then
        szSellStationCondition = " AND sell_station_id = " & TransFieldValueToString(m_oActiveUser.SellStationID)
    End If
    
    
    szSql = "select status,convert(char(8),arrive_time,112) as arrive_date,sum(ISNULL(package_number,0)) as normal_number,sum(ISNULL(transit_charge,0))  as transit_charge,sum(ISNULL(load_charge,0)) as load_charge,sum(ISNULL(keep_charge,0)) as keep_charge,sum(ISNULL(move_charge,0)) as move_charge,sum(ISNULL(send_charge,0)) as send_charge,sum(ISNULL(other_charge,0)) as other_charge,sum(ISNULL(transit_charge,0))+sum(ISNULL(load_charge,0))+sum(ISNULL(keep_charge,0))+sum(ISNULL(move_charge,0))+sum(ISNULL(send_charge,0))+sum(ISNULL(other_charge,0)) as total_amount from package_arrived_info WHERE " _
    & " arrive_time>='" & ToDBDateTime(FromDate) & "' AND " _
    & " arrive_time<='" & ToDBDate(ToDate) & " 23:59:59" & "' " & szSellStationCondition _
    & " GROUP BY convert(char(8),arrive_time,112),status " _
    & " ORDER BY convert(char(8),arrive_time,112),status "
    oDb.ConnectionString = GetConnectionStr()
    Set rsTemp = oDb.Execute(szSql)

'做出Recordset
    '添加字段
    rsData.Fields.Append "arrive_date", rsTemp("arrive_date").Type, rsTemp("arrive_date").DefinedSize
    '添加数量与金额
    For j = 0 To 2
        rsData.Fields.Append "number_ticket_status" & j, adInteger
        rsData.Fields.Append "amount_ticket_status" & j, adCurrency
    Next j
    
    rsData.Fields.Append "total_number", adInteger
    rsData.Fields.Append "total_amount", adCurrency

    
    rsData.Open
    Dim dtLastDate As Date
    Dim szFieldPrefix As String
    
    Do While Not rsTemp.EOF
        If dtLastDate <> rsTemp!arrive_date Then
        
            If rsData.RecordCount > 0 Then
                '求总数及总票款
                lNumber = 0
                dbAmount = 0
                For j = 0 To 2
                    If j <> 2 Then
                        lNumber = lNumber + rsData("number_ticket_status" & j)
                        dbAmount = dbAmount + rsData("amount_ticket_status" & j)
                    End If
                Next j
                rsData!total_number = lNumber
                rsData!total_amount = dbAmount
                rsData.Update
            End If
            
            rsData.AddNew
            
            dtLastDate = CDate(GetDateString(FormatDbValue(rsTemp!arrive_date)))
            rsData!arrive_date = Format(dtLastDate, "yyyyMMdd")
        End If
        
        rsData("number_ticket_status" & rsTemp!Status) = IIf(IsNull(rsTemp!normal_number), 0, rsTemp!normal_number)
        rsData("amount_ticket_status" & rsTemp!Status) = IIf(IsNull(rsTemp!total_amount), 0, rsTemp!total_amount)
        
        rsTemp.MoveNext
    Loop
    
    If rsData.RecordCount > 0 Then
        '求总数及总票款
        lNumber = 0
        dbAmount = 0
        For j = 1 To 2
            If j <> 2 Then
                lNumber = lNumber + rsData("number_ticket_status" & j)
                dbAmount = dbAmount + rsData("amount_ticket_status" & j)
            End If
        Next j
        rsData!total_number = lNumber
        rsData!total_amount = dbAmount
        
        rsData.Update
    End If

        
    Set StationStatMonth = rsData
End Function


Private Function GetDateString(ByVal pszString As String) As String
    Dim pszReturn

    pszReturn = Left(pszString, 4) & "-"
    pszReturn = pszReturn & Mid(pszString, 5, 2) & "-"
    pszReturn = pszReturn & Mid(pszString, 7, 2)
    
    GetDateString = pszReturn
    
End Function
