VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LugDataStatSvr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit

'权限
Public Enum ERightLuggageDataStatSvr
    '33200
    RIGHT_luggagerDataStatSvr = ERR_LuggageDataStatSvr + cnMidRightBegin + cnMidRightStep * 1   ' 33306 日，月，年
    RIGHT_SellStationDataStatSvr = ERR_LuggageDataStatSvr + cnMidRightBegin + cnMidRightStep * 2   '33312售票站日常统计　日，月，年
    RIGHT_StationDataStatSvr = ERR_LuggageDataStatSvr + cnMidRightBegin + cnMidRightStep * 3   '33318站点日常统计　日，月，年
    RIGHT_SellStationStat = ERR_LuggageDataStatSvr + cnMidRightBegin + cnMidRightStep * 4 '33324售票站统计
    RIGHT_SplitCompanyCheckStat = ERR_LuggageDataStatSvr + cnMidRightBegin + cnMidRightStep * 5 ' 33330 拆帐公司签发简报权限
    RIGHT_VehicleCheckStat = ERR_LuggageDataStatSvr + cnMidRightBegin + cnMidRightStep * 6 ' 33336 车辆签发简报权限
    RIGHT_RouteCheckStat = ERR_LuggageDataStatSvr + cnMidRightBegin + cnMidRightStep * 8 ' 33348 线路签发简报权限
    
End Enum
Private m_oActiveUser As ActiveUser


'初始化活动对象和当前售票站点
'##ModelId=3E03A6A700B8
Public Sub Init(poAUser As ActiveUser)
    On Error GoTo InitErr
    Set m_oActiveUser = poAUser
    Exit Sub
InitErr:
    Call RaiseError(err.Number, "LuggageSheet:Init Method", err.Description)
End Sub



'行包统计分析服务


'*********************************************
'站点受理年报
'SellStationId  售票站代码
'StartDate 开始日期
'EndDate  结束日期
'StationId 到达站代码
'AcceptType 托运方式
'
'
'按照售票站、年、到达站、托运方式进行分组，统计指定年的受退废情况
'**********************************************
'##ModelId=3E2363DB018A
'Public Function StationYearDataStat(StartDate As Date, EndDate As Date, Optional SellStationID As String = "", Optional StationID As String = "", Optional AcceptType As String = "") As Recordset
Public Function StationYearDataStat(SellStationID As String, StartDate As Date, EndDate As Date, StationID As String, Optional AcceptType As String = "") As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_LuggageDataStatSvr
    AssertHaveRight m_oActiveUser, RIGHT_StationDataStatSvr
    
    On Error GoTo StationYearDataStatErr
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szTempSelSql As String 'SELECT SQL
    Dim szTempWheSql As String 'WHERE SQL
    Dim szTempGroSql As String 'GROUP BY SQL
    Dim szTempFroSql As String 'FROM SQL
    Dim oDb As New RTConnection
    Dim rsTempStationName As Recordset
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    If SellStationID <> "" Then
        szTempSelSql = " ,s.sell_station_id , MAX(i1.sell_station_name) sell_station_name  "
        szTempFroSql = " ,sell_station_info i1 "
        szTempWheSql = " AND s.sell_station_id =" & TransFieldValueToString(SellStationID) & " AND s.sell_station_id =i1.sell_station_id "
        szTempGroSql = " ,s.sell_station_id  "
    Else
        szTempSelSql = " ,s.sell_station_id , MAX(i1.sell_station_name) sell_station_name "
        szTempFroSql = " ,sell_station_info i1 "
        szTempWheSql = " AND s.sell_station_id =i1.sell_station_id "
        szTempGroSql = " ,s.sell_station_id "
    End If
    
    If StationID <> "" Then
        szTempSelSql = szTempSelSql + " ,s.station_id,MAX(station_name) station_name "
        szTempWheSql = szTempWheSql + " AND s.station_id = " & TransFieldValueToString(StationID)
        szTempGroSql = szTempGroSql + " ,s.station_id "
                
    Else
        szTempSelSql = szTempSelSql + " ,s.station_id,MAX(station_name) station_name"
        szTempWheSql = szTempWheSql + ""
        szTempGroSql = szTempGroSql + " ,s.station_id "
    End If
    
    If AcceptType <> "" Then
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type   "
        szTempWheSql = szTempWheSql + " AND accept_type = " & TransFieldValueToString(GetLuggageTypeInt(AcceptType))
        szTempGroSql = szTempGroSql + " ,accept_type"
    Else
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type  "
        szTempWheSql = szTempWheSql + ""
        szTempGroSql = szTempGroSql + " ,accept_type"
    End If
    
    szSql = "SELECT CONVERT(CHAR(4),bus_date,120) bus_date " & szTempSelSql & "," _
        & " SUM(accept_number) accept_number ,SUM(accept_luggage_number) accept_luggage_number ,SUM(accept_calculate_weight) accept_calculate_weight ,SUM(accept_fact_weight) accept_fact_weight ,SUM(accept_over_number) accept_over_number ,SUM(accept_amount) accept_amount , " _
        & " SUM(cancel_number) cancel_number ,SUM(cancel_luggage_number) cancel_luggage_number ,SUM(cancel_calculate_weight) cancel_calculate_weight ,SUM(cancel_fact_weight) cancel_fact_weight ,SUM(cancel_over_number) cancel_over_number ,SUM(cancel_amount) cancel_amount , " _
        & " SUM(return_number) return_number ,SUM(return_luggage_number) return_luggage_number ,SUM(return_calculate_weight) return_calculate_weight ,SUM(return_fact_weight) return_fact_weight ,SUM(return_over_number) return_over_number ,SUM(return_amount) return_amount , " _
        & " SUM(return_charge) return_charge ,SUM(price_item_1) price_item_1 ,SUM(price_item_2) price_item_2 ,SUM(price_item_3) price_item_3 ,SUM(price_item_4) price_item_4 ,SUM(price_item_5) price_item_5 ,SUM(price_item_6) price_item_6,SUM(price_item_7) price_item_7, " _
        & " SUM(price_item_8) price_item_8 ,SUM(price_item_9) price_item_9 ,SUM(price_item_10) price_item_10 ,SUM(accept_amount-cancel_amount-return_amount+return_charge) total " _
        & " FROM stat_station_luggage_lst s " & szTempFroSql _
        & " WHERE bus_date >=" & TransFieldValueToString(StartDate) & " AND bus_date < " & TransFieldValueToString(EndDate) _
        & szTempWheSql _
        & " GROUP BY CONVERT(CHAR(4),bus_date,120) " & szTempGroSql _
        & " ORDER BY bus_date "
    
    Set rsTemp = oDb.Execute(szSql)
    
    Set StationYearDataStat = rsTemp
    

    Exit Function
StationYearDataStatErr:
'    Call RaiseError(MyUnhandledError, "LugDataStatSvr:StationYearDataStat Method")
End Function

'*********************************************
'站点受理月报
'SellStationId  售票站代码
'StartDate 开始日期
'EndDate  结束日期
'StationId 到达站代码
'AcceptType 托运方式
'
'
'按照售票站、月、到达站、托运方式进行分组，统计指定月的受退废情况
'**********************************************
'##ModelId=3E2363A4019F
Public Function StationMonthDataStat(SellStationID As String, StartDate As Date, EndDate As Date, StationID As String, Optional AcceptType As String = "") As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_LuggageDataStatSvr
    AssertHaveRight m_oActiveUser, RIGHT_StationDataStatSvr
    On Error GoTo StationMonthDataStatErr
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szTempSelSql As String 'SELECT SQL
    Dim szTempWheSql As String 'WHERE SQL
    Dim szTempGroSql As String 'GROUP BY SQL
    Dim szTempFroSql As String 'FROM SQL
    Dim oDb As New RTConnection
    Dim TempType As Integer
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    If SellStationID <> "" Then
        szTempSelSql = " ,s.sell_station_id ,MAX(i1.sell_station_name) sell_station_name  "
        szTempFroSql = " ,sell_station_info i1 "
        szTempWheSql = " AND s.sell_station_id =" & TransFieldValueToString(SellStationID) & " AND s.sell_station_id = i1.sell_station_id "
        szTempGroSql = " ,s.sell_station_id  "
    Else
        szTempSelSql = " ,s.sell_station_id ,MAX(i1.sell_station_name) sell_station_name "
        szTempFroSql = " ,sell_station_info i1 "
        szTempWheSql = " AND s.sell_station_id = i1.sell_station_id "
        szTempGroSql = " ,s.sell_station_id "
    End If
    
    If StationID <> "" Then
        szTempSelSql = szTempSelSql + " ,s.station_id,MAX(station_name)  station_name "
        szTempWheSql = szTempWheSql + " AND s.station_id = " & TransFieldValueToString(StationID)
        szTempGroSql = szTempGroSql + " ,s.station_id "
    Else
        szTempSelSql = szTempSelSql + " ,s.station_id,MAX(station_name) station_name "
        szTempWheSql = szTempWheSql + ""
        szTempGroSql = szTempGroSql + " ,s.station_id "
    End If
    
    If AcceptType <> "" Then
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type  "
        szTempWheSql = szTempWheSql + " AND accept_type = " & TransFieldValueToString(GetLuggageTypeInt(AcceptType))
        szTempGroSql = szTempGroSql + " ,accept_type "
    Else
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type"
        szTempWheSql = szTempWheSql + ""
        szTempGroSql = szTempGroSql + " ,accept_type "
        
    End If
    szSql = "SELECT CONVERT(CHAR(7),bus_date,120) bus_date " & szTempSelSql & "," _
        & " SUM(accept_number) accept_number ,SUM(accept_luggage_number) accept_luggage_number ,SUM(accept_calculate_weight) accept_calculate_weight ,SUM(accept_fact_weight) accept_fact_weight ,SUM(accept_over_number) accept_over_number ,SUM(accept_amount) accept_amount , " _
        & " SUM(cancel_number) cancel_number ,SUM(cancel_luggage_number) cancel_luggage_number ,SUM(cancel_calculate_weight) cancel_calculate_weight ,SUM(cancel_fact_weight) cancel_fact_weight ,SUM(cancel_over_number) cancel_over_number ,SUM(cancel_amount) cancel_amount , " _
        & " SUM(return_number) return_number ,SUM(return_luggage_number) return_luggage_number ,SUM(return_calculate_weight) return_calculate_weight ,SUM(return_fact_weight) return_fact_weight ,SUM(return_over_number) return_over_number ,SUM(return_amount) return_amount , " _
        & " SUM(return_charge) return_charge ,SUM(price_item_1) price_item_1 ,SUM(price_item_2) price_item_2 ,SUM(price_item_3) price_item_3 ,SUM(price_item_4) price_item_4 ,SUM(price_item_5) price_item_5 ,SUM(price_item_6) price_item_6,SUM(price_item_7) price_item_7, " _
        & " SUM(price_item_8) price_item_8 ,SUM(price_item_9) price_item_9 ,SUM(price_item_10) price_item_10 ,SUM(accept_amount-cancel_amount-return_amount+return_charge) total " _
        & " FROM stat_station_luggage_lst s " & szTempFroSql _
        & " WHERE bus_date >=" & TransFieldValueToString(StartDate) & " AND bus_date < " & TransFieldValueToString(EndDate) _
        & szTempWheSql _
        & " GROUP BY CONVERT(CHAR(7),bus_date,120) " & szTempGroSql _
        & " ORDER BY bus_date "
    Set rsTemp = oDb.Execute(szSql)
    Set StationMonthDataStat = rsTemp
    Exit Function
StationMonthDataStatErr:
'    Call RaiseError(MyUnhandledError, "LugDataStatSvr:StationMonthDataStat Method")
End Function

'*********************************************
'站点受理日报
'SellStationId  售票站代码
'StartDate 开始日期
'EndDate  结束日期
'StationId 到达站代码
'AcceptType 托运方式
'
'
'按照售票站、日、到达站、托运方式进行分组，统计指定日的受退废情况
'**********************************************
'##ModelId=3E232CD1009E
Public Function StationDayDataStat(SellStationID As String, StartDate As Date, EndDate As Date, StationID As String, Optional AcceptType As String = "") As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_LuggageDataStatSvr
    AssertHaveRight m_oActiveUser, RIGHT_StationDataStatSvr
    On Error GoTo StationDayDataStatErr
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szTempSelSql As String      'SELECT SQL
    Dim szTempWheSql As String      'WHERE SQL
    Dim szTempGroSql As String      'GROUP BY SQL
    Dim szTempFroSql As String      'FROM SQL
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    If SellStationID <> "" Then
        szTempSelSql = " ,s.sell_station_id, MAX(i1.sell_station_name) as sell_station_name  "
        szTempFroSql = " ,sell_station_info i1 "
        szTempWheSql = " AND s.sell_station_id =" & TransFieldValueToString(SellStationID) & " AND s.sell_station_id=i1.sell_station_id "
        szTempGroSql = " ,s.sell_station_id  "
    Else
        szTempSelSql = " ,s.sell_station_id ,MAX(i1.sell_station_name) as sell_station_name "
        szTempFroSql = " ,sell_station_info i1 "
        szTempWheSql = " AND s.sell_station_id=i1.sell_station_id "
        szTempGroSql = " ,s.sell_station_id"
    End If
    
    If StationID <> "" Then
        szTempSelSql = szTempSelSql + " ,s.station_id,MAX(station_name) station_name"
        szTempWheSql = szTempWheSql + " AND  s.station_id = " & TransFieldValueToString(StationID)
        szTempGroSql = szTempGroSql + " ,s.station_id "
    Else
        szTempSelSql = szTempSelSql + " ,s.station_id,MAX(station_name) station_name "
        szTempWheSql = szTempWheSql + ""
        szTempGroSql = szTempGroSql + " ,s.station_id "
    End If
    
    If AcceptType <> "" Then
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type   "
        szTempWheSql = szTempWheSql + " AND accept_type = " & TransFieldValueToString(GetLuggageTypeInt(AcceptType))
        szTempGroSql = szTempGroSql + " ,accept_type"
    Else
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type  "
        szTempWheSql = szTempWheSql + ""
        szTempGroSql = szTempGroSql + " ,accept_type "
    End If
    szSql = "SELECT CONVERT(CHAR(10),bus_date,120) bus_date " & szTempSelSql & "," _
        & " SUM(accept_number) accept_number ,SUM(accept_luggage_number) accept_luggage_number ,SUM(accept_calculate_weight) accept_calculate_weight ,SUM(accept_fact_weight) accept_fact_weight ,SUM(accept_over_number) accept_over_number ,SUM(accept_amount) accept_amount , " _
        & " SUM(cancel_number) cancel_number ,SUM(cancel_luggage_number) cancel_luggage_number ,SUM(cancel_calculate_weight) cancel_calculate_weight ,SUM(cancel_fact_weight) cancel_fact_weight ,SUM(cancel_over_number) cancel_over_number ,SUM(cancel_amount) cancel_amount , " _
        & " SUM(return_number) return_number ,SUM(return_luggage_number) return_luggage_number ,SUM(return_calculate_weight) return_calculate_weight ,SUM(return_fact_weight) return_fact_weight ,SUM(return_over_number) return_over_number ,SUM(return_amount) return_amount , " _
        & " SUM(return_charge) return_charge ,SUM(price_item_1) price_item_1 ,SUM(price_item_2) price_item_2 ,SUM(price_item_3) price_item_3 ,SUM(price_item_4) price_item_4 ,SUM(price_item_5) price_item_5 ,SUM(price_item_6) price_item_6,SUM(price_item_7) price_item_7, " _
        & " SUM(price_item_8) price_item_8 ,SUM(price_item_9) price_item_9 ,SUM(price_item_10) price_item_10 ,SUM(accept_amount-cancel_amount-return_amount+return_charge) total " _
        & " FROM stat_station_luggage_lst s " & szTempFroSql _
        & " WHERE bus_date >=" & TransFieldValueToString(StartDate) & " AND bus_date < " & TransFieldValueToString(EndDate) _
        & szTempWheSql _
        & " GROUP BY CONVERT(CHAR(10),bus_date,120) " & szTempGroSql _
        & " ORDER BY bus_date "
    Set rsTemp = oDb.Execute(szSql)
    Set StationDayDataStat = rsTemp
    Exit Function
StationDayDataStatErr:
'    Call RaiseError(MyUnhandledError, "LugDataStatSvr:StationDayDataStat Method")
End Function

'*********************************************
'行包员受理年报
'SellStationId  售票站代码
'StartDate 开始日期
'EndDate  结束日期
'UserID() 用户代码数组
'AcceptType 托运方式
'
'按照售票站、年、行包员、托运方式进行分组，统计指定年的受退废情况
'**********************************************

'##ModelId=3E23620F017F
Public Function LuggagerDataYearStat(SellStationID As String, StartDate As Date, EndDate As Date, UserID() As String, Optional AcceptType As String = "") As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_LuggageDataStatSvr
    AssertHaveRight m_oActiveUser, RIGHT_luggagerDataStatSvr
    On Error GoTo LuggagerDataYearStatErr
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szTempSelSql As String
    Dim szTempFroSql As String
    Dim szTempWheSql As String
    Dim szTempGroSql As String
    Dim oDb As New RTConnection
    Dim aszUserID() As String
    Dim szUserID As String
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    If ArrayLength(UserID) Then
        szUserID = "("
        Dim i As Integer
        For i = 1 To ArrayLength(UserID) - 1
            szUserID = szUserID & TransFieldValueToString(UserID(i)) & ","
        Next i
        szUserID = szUserID & TransFieldValueToString(UserID(i)) & ")"
        szTempFroSql = " ,user_info u "
        szTempWheSql = " u.user_id = s.user_id "
        szTempSelSql = " ,s.user_id , MAX(u.user_name) user_name "
        szTempGroSql = " ,s.user_id "
    Else
        szTempSelSql = " ,s.user_id , MAX(u.user_name) user_name "
        szTempFroSql = " ,user_info u "
        szTempWheSql = "  u.user_id = s.user_id "
        szTempGroSql = " ,s.user_id "
    End If
    If SellStationID <> "" Then
        szTempSelSql = szTempSelSql + " ,s.sell_station_id ,MAX(i1.sell_station_name) sell_station_name  "
        szTempFroSql = szTempFroSql + " ,sell_station_info i1 "
        szTempWheSql = szTempWheSql + " AND s.sell_station_id =" & TransFieldValueToString(SellStationID) & " AND s.sell_station_id = i1.sell_station_id "
        szTempGroSql = szTempGroSql + " ,s.sell_station_id  "
    Else
        szTempSelSql = szTempSelSql + " ,s.sell_station_id ,MAX(i1.sell_station_name) sell_station_name  "
        szTempFroSql = szTempFroSql + " ,sell_station_info i1 "
        szTempWheSql = szTempWheSql + " AND s.sell_station_id = i1.sell_station_id "
        szTempGroSql = szTempGroSql + " ,s.sell_station_id "
    End If
    
    
    If AcceptType <> "" Then
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type "
        szTempWheSql = szTempWheSql + " AND accept_type = " & TransFieldValueToString(GetLuggageTypeInt(AcceptType))
        szTempGroSql = szTempGroSql + " ,accept_type "
    Else
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type"
        szTempWheSql = szTempWheSql + " "
        szTempGroSql = szTempGroSql + " ,accept_type "
    End If
    szSql = "SELECT CONVERT(CHAR(4),operation_time,120) operation_time " & szTempSelSql & "," _
        & " SUM(accept_number) accept_number ,SUM(accept_luggage_number) accept_luggage_number ,SUM(accept_calculate_weight) accept_calculate_weight ,SUM(accept_fact_weight) accept_fact_weight ,SUM(accept_over_number) accept_over_number ,SUM(accept_amount) accept_amount , " _
        & " SUM(cancel_number) cancel_number ,SUM(cancel_luggage_number) cancel_luggage_number ,SUM(cancel_calculate_weight) cancel_calculate_weight ,SUM(cancel_fact_weight) cancel_fact_weight ,SUM(cancel_over_number) cancel_over_number ,SUM(cancel_amount) cancel_amount , " _
        & " SUM(return_number) return_number ,SUM(return_luggage_number) return_luggage_number ,SUM(return_calculate_weight) return_calculate_weight ,SUM(return_fact_weight) return_fact_weight ,SUM(return_over_number) return_over_number ,SUM(return_amount) return_amount , " _
        & " SUM(return_charge) return_charge ,SUM(price_item_1) price_item_1 ,SUM(price_item_2) price_item_2 ,SUM(price_item_3) price_item_3 ,SUM(price_item_4) price_item_4 ,SUM(price_item_5) price_item_5 ,SUM(price_item_6) price_item_6 ,SUM(price_item_7) price_item_7 , " _
        & " SUM(price_item_8) price_item_8 ,SUM(price_item_9) price_item_9 ,SUM(price_item_10) price_item_10 ,SUM(accept_amount-cancel_amount-return_amount+return_charge) total " _
        & " FROM stat_saler_luggage_lst s " & szTempFroSql _
        & " WHERE operation_time >=" & TransFieldValueToString(StartDate) & " AND operation_time < " & TransFieldValueToString(EndDate) _
        & IIf(szUserID <> "", " AND S.user_id IN " & szUserID, "") & " AND " & szTempWheSql _
        & " GROUP BY CONVERT(CHAR(4),operation_time,120) " & szTempGroSql _
        & " ORDER BY operation_time "
    Set rsTemp = oDb.Execute(szSql)
    Set LuggagerDataYearStat = rsTemp
    Exit Function
LuggagerDataYearStatErr:
'    Call RaiseError(MyUnhandledError, "LugDataStatSvr:LuggagerDataYearStat Method")
End Function

'*********************************************
'行包员受理月报
'SellStationId  售票站代码
'StartDate 开始日期
'EndDate  结束日期
'UserID() 用户代码数组
'AcceptType 托运方式
'按照售票站、月、行包员、托运方式进行分组，统计指定月的受退废情况
'**********************************************
'##ModelId=3E2361B800B1
Public Function LuggagerDataMonthStat(SellStationID As String, StartDate As Date, EndDate As Date, UserID() As String, Optional AcceptType As String = "") As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_LuggageDataStatSvr
    AssertHaveRight m_oActiveUser, RIGHT_luggagerDataStatSvr
    On Error GoTo LuggagerDataMonthStatErr
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szTempSelSql As String
    Dim szTempFroSql As String
    Dim szTempWheSql As String
    Dim szTempGroSql As String
    Dim oDb As New RTConnection
    Dim szUserID As String
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    If ArrayLength(UserID) Then
        szUserID = "("
        Dim i As Integer
        For i = 1 To ArrayLength(UserID) - 1
            szUserID = szUserID & TransFieldValueToString(UserID(i)) & ","
        Next i
        szUserID = szUserID & TransFieldValueToString(UserID(i)) & ")"
        szTempSelSql = " ,s.user_id , MAX(u.user_name) user_name "
        szTempFroSql = " ,user_info u "
        szTempWheSql = " AND u.user_id = s.user_id "
        szTempGroSql = " ,s.user_id"
    Else
        szTempSelSql = " ,s.user_id , MAX(u.user_name) user_name "
        szTempFroSql = " ,user_info u "
        szTempWheSql = " AND u.user_id = s.user_id "
        szTempGroSql = " ,s.user_id"
    End If
    
    If SellStationID <> "" Then
        szTempSelSql = szTempSelSql + " ,s.sell_station_id ,MAX(i1.sell_station_name) sell_station_name "
        szTempFroSql = szTempFroSql + " ,sell_station_info i1 "
        szTempWheSql = szTempWheSql + " AND s.sell_station_id =" & TransFieldValueToString(SellStationID) & " AND s.sell_station_id = i1.sell_station_id "
        szTempGroSql = szTempGroSql + " ,s.sell_station_id  "
    Else
        szTempSelSql = szTempSelSql + " ,s.sell_station_id ,MAX(i1.sell_station_name) sell_station_name "
        szTempFroSql = szTempFroSql + " ,sell_station_info i1 "
        szTempWheSql = szTempWheSql + "  AND s.sell_station_id = i1.sell_station_id "
        szTempGroSql = szTempGroSql + " ,s.sell_station_id "
    End If
    
    If AcceptType <> "" Then
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type "
        szTempWheSql = szTempWheSql + " AND accept_type = " & TransFieldValueToString(GetLuggageTypeInt(AcceptType))
        szTempGroSql = szTempGroSql + " ,accept_type"
    Else
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type"
        szTempWheSql = szTempWheSql + ""
        szTempGroSql = szTempGroSql + " ,accept_type"
    End If
    szSql = "SELECT CONVERT(CHAR(7),operation_time,120) operation_time " & szTempSelSql & "," _
        & " SUM(accept_number) accept_number ,SUM(accept_luggage_number) accept_luggage_number ,SUM(accept_calculate_weight) accept_calculate_weight ,SUM(accept_fact_weight) accept_fact_weight ,SUM(accept_over_number) accept_over_number ,SUM(accept_amount) accept_amount , " _
        & " SUM(cancel_number) cancel_number ,SUM(cancel_luggage_number) cancel_luggage_number ,SUM(cancel_calculate_weight) cancel_calculate_weight ,SUM(cancel_fact_weight) cancel_fact_weight ,SUM(cancel_over_number) cancel_over_number ,SUM(cancel_amount) cancel_amount , " _
        & " SUM(return_number) return_number ,SUM(return_luggage_number) return_luggage_number ,SUM(return_calculate_weight) return_calculate_weight ,SUM(return_fact_weight) return_fact_weight ,SUM(return_over_number) return_over_number ,SUM(return_amount) return_amount , " _
        & " SUM(return_charge) return_charge ,SUM(price_item_1) price_item_1 ,SUM(price_item_2) price_item_2 ,SUM(price_item_3) price_item_3 ,SUM(price_item_4) price_item_4 ,SUM(price_item_5) price_item_5 ,SUM(price_item_6) price_item_6 ,SUM(price_item_7) price_item_7 , " _
        & " SUM(price_item_8) price_item_8 ,SUM(price_item_9) price_item_9 ,SUM(price_item_10) price_item_10 ,SUM(accept_amount-cancel_amount-return_amount+return_charge) total " _
        & " FROM stat_saler_luggage_lst s " & szTempFroSql _
        & " WHERE operation_time >=" & TransFieldValueToString(StartDate) & " AND operation_time < " & TransFieldValueToString(EndDate) _
        & IIf(szUserID <> "", " AND s.user_id IN " & szUserID, "") & szTempWheSql _
        & " GROUP BY CONVERT(CHAR(7),operation_time,120) " & szTempGroSql _
        & " ORDER BY operation_time "
    Set rsTemp = oDb.Execute(szSql)
    Set LuggagerDataMonthStat = rsTemp
    Exit Function
LuggagerDataMonthStatErr:
'    Call RaiseError(MyUnhandledError, "LugDataStatSvr:LuggagerDataMonthStat Method")
End Function
'*********************************************
'行包员受理日报
'SellStationId  售票站代码
'StartDate 开始日期
'EndDate  结束日期
'UserID() 用户代码数组
'AcceptType 托运方式
'按照售票站、日、行包员、托运方式进行分组，统计指定日的受退废情况
'**********************************************
'##ModelId=3E2311E1032D
Public Function LuggagerDataDayStat(SellStationID As String, StartDate As Date, EndDate As Date, UserID() As String, Optional AcceptType As String = "") As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_LuggageDataStatSvr
    AssertHaveRight m_oActiveUser, RIGHT_luggagerDataStatSvr
    On Error GoTo LuggagerDataDayStatErr
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szTempSelSql As String
    Dim szTempFroSql As String
    Dim szTempWheSql As String
    Dim szTempGroSql As String
    Dim oDb As New RTConnection
    Dim szUserID As String
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    If ArrayLength(UserID) Then
        szUserID = "("
        Dim i As Integer
        For i = 1 To ArrayLength(UserID) - 1
            szUserID = szUserID & TransFieldValueToString(UserID(i)) & ","
        Next i
        szUserID = szUserID & TransFieldValueToString(UserID(i)) & ")"
        szTempFroSql = " ,user_info u"
        szTempSelSql = " ,s.user_id ,MAX(u.user_name) user_name  "
        szTempWheSql = " AND u.user_id = s.user_id "
        szTempGroSql = ",s.user_id"
    Else
        szTempFroSql = " ,user_info u"
        szTempSelSql = " ,s.user_id ,MAX(u.user_name) user_name "
        szTempWheSql = " AND u.user_id = s.user_id "
        szTempGroSql = ",s.user_id"
    End If
    
    If SellStationID <> "" Then
        szTempSelSql = szTempSelSql + " ,s.sell_station_id , MAX(i1.sell_station_name) sell_station_name  "
        szTempFroSql = szTempFroSql + " ,sell_station_info i1 "
        szTempWheSql = szTempWheSql + " AND s.sell_station_id =" & TransFieldValueToString(SellStationID) & " AND s.sell_station_id = i1.sell_station_id "
        szTempGroSql = szTempGroSql + " ,s.sell_station_id  "
    Else
        szTempSelSql = szTempSelSql + " ,s.sell_station_id , MAX(i1.sell_station_name) sell_station_name "
        szTempFroSql = szTempFroSql + " ,sell_station_info i1 "
        szTempWheSql = szTempWheSql + " AND s.sell_station_id = i1.sell_station_id "
        szTempGroSql = szTempGroSql + " ,s.sell_station_id"
    End If

    If AcceptType <> "" Then
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type "
        szTempWheSql = szTempWheSql + " AND accept_type = " & TransFieldValueToString(GetLuggageTypeInt(AcceptType))
        szTempGroSql = szTempGroSql + " ,accept_type"
    Else
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type "
        szTempWheSql = szTempWheSql + ""
        szTempGroSql = szTempGroSql + " ,accept_type"
    End If
    szSql = "SELECT CONVERT(CHAR(10),operation_time,120) operation_time " & szTempSelSql & "," _
        & " SUM(accept_number) accept_number ,SUM(accept_luggage_number) accept_luggage_number ,SUM(accept_calculate_weight) accept_calculate_weight ,SUM(accept_fact_weight) accept_fact_weight ,SUM(accept_over_number) accept_over_number ,SUM(accept_amount) accept_amount , " _
        & " SUM(cancel_number) cancel_number ,SUM(cancel_luggage_number) cancel_luggage_number ,SUM(cancel_calculate_weight) cancel_calculate_weight ,SUM(cancel_fact_weight) cancel_fact_weight ,SUM(cancel_over_number) cancel_over_number ,SUM(cancel_amount) cancel_amount , " _
        & " SUM(return_number) return_number ,SUM(return_luggage_number) return_luggage_number ,SUM(return_calculate_weight) return_calculate_weight ,SUM(return_fact_weight) return_fact_weight ,SUM(return_over_number) return_over_number ,SUM(return_amount) return_amount , " _
        & " SUM(return_charge) return_charge ,SUM(price_item_1) price_item_1 ,SUM(price_item_2) price_item_2 ,SUM(price_item_3) price_item_3 ,SUM(price_item_4) price_item_4 ,SUM(price_item_5) price_item_5 ,SUM(price_item_6) price_item_6 ,SUM(price_item_7) price_item_7 , " _
        & " SUM(price_item_8) price_item_8 ,SUM(price_item_9) price_item_9 ,SUM(price_item_10) price_item_10 ,SUM(accept_amount-cancel_amount-return_amount+return_charge) total " _
        & " FROM stat_saler_luggage_lst s" & szTempFroSql _
        & " WHERE operation_time >=" & TransFieldValueToString(StartDate) & " AND operation_time < " & TransFieldValueToString(EndDate) _
        & IIf(szUserID <> "", " AND s.user_id IN " & szUserID, " ") & szTempWheSql _
        & " GROUP BY CONVERT(CHAR(10),operation_time,120) " & szTempGroSql _
        & " ORDER BY operation_time "
    Set rsTemp = oDb.Execute(szSql)
    Set LuggagerDataDayStat = rsTemp
    Exit Function
LuggagerDataDayStatErr:
'    Call RaiseError(MyUnhandledError, "LugDataStatSvr:LuggagerDataDayStat Method")
End Function
'指定受理单
'##ModelId=3E233BD7019E
Public Sub Identify(SheetID As String)
    On Error GoTo IdentifyErr

    'your code goes here...

    Exit Sub
IdentifyErr:
'    Call RaiseError(MyUnhandledError, "LugDataStatSvr:Identify Method")
End Sub

'##ModelId=3E233E470001
Private Sub Class_Initialize()
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'##ModelId=3E233E470083
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

#If DebugMode Then
    '##ModelId=3E233E46024E
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

'售票站点按年统计
Public Function SellStationYearDataStat(SellStationID As String, StartDate As Date, EndDate As Date, Optional AcceptType As String = "") As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_LuggageDataStatSvr
    AssertHaveRight m_oActiveUser, RIGHT_SellStationDataStatSvr
    On Error GoTo StationYearDataStatErr
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szTempSelSql As String 'SELECT SQL
    Dim szTempWheSql As String 'WHERE SQL
    Dim szTempGroSql As String 'GROUP BY SQL
    Dim szTempFroSql As String 'FROM SQL
    Dim oDb As New RTConnection
    Dim rsTempStationName As Recordset
    
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    If SellStationID <> "" Then
        szTempSelSql = " ,s.sell_station_id ,MAX(i.sell_station_name) as sell_station_name "
        szTempFroSql = " ,sell_station_info i"
        szTempWheSql = " AND s.sell_station_id =" & TransFieldValueToString(SellStationID) & " AND s.sell_station_id = i.sell_station_id "
        szTempGroSql = " ,s.sell_station_id  "
    Else
        szTempSelSql = " ,s.sell_station_id ,MAX(i.sell_station_name) as sell_station_name "
        szTempFroSql = " ,sell_station_info i "
        szTempWheSql = " AND s.sell_station_id = i.sell_station_id "
        szTempGroSql = " ,s.sell_station_id "
    End If
    
    If AcceptType <> "" Then
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type   "
        szTempWheSql = szTempWheSql + " AND accept_type = " & TransFieldValueToString(GetLuggageTypeInt(AcceptType))
        szTempGroSql = szTempGroSql + " ,accept_type"
    Else
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type  "
        szTempWheSql = szTempWheSql + ""
        szTempGroSql = szTempGroSql + " ,accept_type"
    End If
    
    szSql = "SELECT CONVERT(CHAR(4),bus_date,120) bus_date " & szTempSelSql & "," _
        & " SUM(accept_number) accept_number ,SUM(accept_luggage_number) accept_luggage_number ,SUM(accept_calculate_weight) accept_calculate_weight ,SUM(accept_fact_weight) accept_fact_weight ,SUM(accept_over_number) accept_over_number ,SUM(accept_amount) accept_amount , " _
        & " SUM(cancel_number) cancel_number ,SUM(cancel_luggage_number) cancel_luggage_number ,SUM(cancel_calculate_weight) cancel_calculate_weight ,SUM(cancel_fact_weight) cancel_fact_weight ,SUM(cancel_over_number) cancel_over_number ,SUM(cancel_amount) cancel_amount , " _
        & " SUM(return_number) return_number ,SUM(return_luggage_number) return_luggage_number ,SUM(return_calculate_weight) return_calculate_weight ,SUM(return_fact_weight) return_fact_weight ,SUM(return_over_number) return_over_number ,SUM(return_amount) return_amount , " _
        & " SUM(return_charge) return_charge ,SUM(price_item_1) price_item_1 ,SUM(price_item_2) price_item_2 ,SUM(price_item_3) price_item_3 ,SUM(price_item_4) price_item_4 ,SUM(price_item_5) price_item_5 ,SUM(price_item_6) price_item_6,SUM(price_item_7) price_item_7, " _
        & " SUM(price_item_8) price_item_8 ,SUM(price_item_9) price_item_9 ,SUM(price_item_10) price_item_10 ,SUM(accept_amount-cancel_amount-return_amount+return_charge) total " _
        & " FROM stat_station_luggage_lst s " & szTempFroSql _
        & " WHERE bus_date >=" & TransFieldValueToString(StartDate) & " AND bus_date < " & TransFieldValueToString(EndDate) _
        & szTempWheSql _
        & " GROUP BY CONVERT(CHAR(4),bus_date,120) " & szTempGroSql _
        & " ORDER BY bus_date "
    Set rsTemp = oDb.Execute(szSql)
    Set SellStationYearDataStat = rsTemp
    Exit Function
StationYearDataStatErr:
'    Call RaiseError(MyUnhandledError, "LugDataStatSvr:StationYearDataStat Method")
End Function
'售票站点按月统计
Public Function SellStationMonthDataStat(SellStationID As String, StartDate As Date, EndDate As Date, Optional AcceptType As String = "") As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_LuggageDataStatSvr
    AssertHaveRight m_oActiveUser, RIGHT_luggagerDataStatSvr
    On Error GoTo StationMonthDataStatErr
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szTempSelSql As String 'SELECT SQL
    Dim szTempWheSql As String 'WHERE SQL
    Dim szTempGroSql As String 'GROUP BY SQL
    Dim szTempFroSql As String 'FROM SQL
    Dim oDb As New RTConnection
    Dim TempType As Integer
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    If SellStationID <> "" Then
        szTempSelSql = " ,s.sell_station_id ,MAX(i.sell_station_name) as sell_station_name  "
        szTempFroSql = " ,sell_station_info i "
        szTempWheSql = " AND s.sell_station_id =" & TransFieldValueToString(SellStationID) & " AND s.sell_station_id = i.sell_station_id "
        szTempGroSql = " ,s.sell_station_id  "
    Else
        szTempSelSql = " ,s.sell_station_id ,MAX(i.sell_station_name) as sell_station_name "
        szTempFroSql = " ,sell_station_info i "
        szTempWheSql = " AND s.sell_station_id = i.sell_station_id "
        szTempGroSql = " ,s.sell_station_id "
    End If
    
    If AcceptType <> "" Then
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type  "
        szTempWheSql = szTempWheSql + " AND accept_type = " & TransFieldValueToString(GetLuggageTypeInt(AcceptType))
        szTempGroSql = szTempGroSql + " ,accept_type "
    Else
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type"
        szTempWheSql = szTempWheSql + ""
        szTempGroSql = szTempGroSql + " ,accept_type "
        
    End If
    szSql = "SELECT CONVERT(CHAR(7),bus_date,120) bus_date " & szTempSelSql & "," _
        & " SUM(accept_number) accept_number ,SUM(accept_luggage_number) accept_luggage_number ,SUM(accept_calculate_weight) accept_calculate_weight ,SUM(accept_fact_weight) accept_fact_weight ,SUM(accept_over_number) accept_over_number ,SUM(accept_amount) accept_amount , " _
        & " SUM(cancel_number) cancel_number ,SUM(cancel_luggage_number) cancel_luggage_number ,SUM(cancel_calculate_weight) cancel_calculate_weight ,SUM(cancel_fact_weight) cancel_fact_weight ,SUM(cancel_over_number) cancel_over_number ,SUM(cancel_amount) cancel_amount , " _
        & " SUM(return_number) return_number ,SUM(return_luggage_number) return_luggage_number ,SUM(return_calculate_weight) return_calculate_weight ,SUM(return_fact_weight) return_fact_weight ,SUM(return_over_number) return_over_number ,SUM(return_amount) return_amount , " _
        & " SUM(return_charge) return_charge ,SUM(price_item_1) price_item_1 ,SUM(price_item_2) price_item_2 ,SUM(price_item_3) price_item_3 ,SUM(price_item_4) price_item_4 ,SUM(price_item_5) price_item_5 ,SUM(price_item_6) price_item_6,SUM(price_item_7) price_item_7, " _
        & " SUM(price_item_8) price_item_8 ,SUM(price_item_9) price_item_9 ,SUM(price_item_10) price_item_10 ,SUM(accept_amount-cancel_amount-return_amount+return_charge) total " _
        & " FROM stat_station_luggage_lst s " & szTempFroSql _
        & " WHERE bus_date >=" & TransFieldValueToString(StartDate) & " AND bus_date < " & TransFieldValueToString(EndDate) _
        & szTempWheSql _
        & " GROUP BY CONVERT(CHAR(7),bus_date,120) " & szTempGroSql _
        & " ORDER BY bus_date "
    Set rsTemp = oDb.Execute(szSql)
    Set SellStationMonthDataStat = rsTemp
    Exit Function
StationMonthDataStatErr:
'    Call RaiseError(MyUnhandledError, "LugDataStatSvr:StationMonthDataStat Method")
End Function

'售票站点按日统计
Public Function SellStationDayDataStat(SellStationID As String, StartDate As Date, EndDate As Date, Optional AcceptType As String = "") As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_LuggageDataStatSvr
    AssertHaveRight m_oActiveUser, RIGHT_SellStationDataStatSvr
    On Error GoTo StationDayDataStatErr
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szTempSelSql As String      'SELECT SQL
    Dim szTempWheSql As String      'WHERE SQL
    Dim szTempGroSql As String      'GROUP BY SQL
    Dim szTempFroSql As String      'FROM SQL
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    If SellStationID <> "" Then
        szTempSelSql = " ,s.sell_station_id ,MAX(i.sell_station_name) as sell_station_name "
        szTempFroSql = " ,sell_station_info i "
        szTempWheSql = " AND s.sell_station_id =" & TransFieldValueToString(SellStationID) & " AND s.sell_station_id = i.sell_station_id "
        szTempGroSql = " ,s.sell_station_id  "
    Else
        szTempSelSql = " ,s.sell_station_id ,MAX(i.sell_station_name) as sell_station_name "
        szTempFroSql = " ,sell_station_info i "
        szTempWheSql = " AND s.sell_station_id = i.sell_station_id "
        szTempGroSql = " ,s.sell_station_id"
    End If
    If AcceptType <> "" Then
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type   "
        szTempWheSql = szTempWheSql + " AND accept_type = " & TransFieldValueToString(GetLuggageTypeInt(AcceptType))
        szTempGroSql = szTempGroSql + " ,accept_type"
    Else
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type  "
        szTempWheSql = szTempWheSql + ""
        szTempGroSql = szTempGroSql + " ,accept_type "
    End If
    szSql = "SELECT CONVERT(CHAR(10),bus_date,120) bus_date " & szTempSelSql & "," _
        & " SUM(accept_number) accept_number ,SUM(accept_luggage_number) accept_luggage_number ,SUM(accept_calculate_weight) accept_calculate_weight ,SUM(accept_fact_weight) accept_fact_weight ,SUM(accept_over_number) accept_over_number ,SUM(accept_amount) accept_amount , " _
        & " SUM(cancel_number) cancel_number ,SUM(cancel_luggage_number) cancel_luggage_number ,SUM(cancel_calculate_weight) cancel_calculate_weight ,SUM(cancel_fact_weight) cancel_fact_weight ,SUM(cancel_over_number) cancel_over_number ,SUM(cancel_amount) cancel_amount , " _
        & " SUM(return_number) return_number ,SUM(return_luggage_number) return_luggage_number ,SUM(return_calculate_weight) return_calculate_weight ,SUM(return_fact_weight) return_fact_weight ,SUM(return_over_number) return_over_number ,SUM(return_amount) return_amount , " _
        & " SUM(return_charge) return_charge ,SUM(price_item_1) price_item_1 ,SUM(price_item_2) price_item_2 ,SUM(price_item_3) price_item_3 ,SUM(price_item_4) price_item_4 ,SUM(price_item_5) price_item_5 ,SUM(price_item_6) price_item_6,SUM(price_item_7) price_item_7, " _
        & " SUM(price_item_8) price_item_8 ,SUM(price_item_9) price_item_9 ,SUM(price_item_10) price_item_10 ,SUM(accept_amount-cancel_amount-return_amount+return_charge) total " _
        & " FROM stat_station_luggage_lst s " & szTempFroSql _
        & " WHERE bus_date >=" & TransFieldValueToString(StartDate) & " AND bus_date < " & TransFieldValueToString(EndDate) _
        & szTempWheSql _
        & " GROUP BY CONVERT(CHAR(10),bus_date,120) " & szTempGroSql _
        & " ORDER BY bus_date "
    Set rsTemp = oDb.Execute(szSql)
    Set SellStationDayDataStat = rsTemp
    Exit Function
StationDayDataStatErr:
'    Call RaiseError(MyUnhandledError, "LugDataStatSvr:StationDayDataStat Method")
End Function
'售票站点简报
Public Function SellStationStat(SellStationID As String, StartDate As Date, EndDate As Date, Optional AcceptType As String = "") As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_LuggageDataStatSvr
    AssertHaveRight m_oActiveUser, RIGHT_SellStationStat
    On Error GoTo StationMonthDataStatErr
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szTempSelSql As String 'SELECT SQL
    Dim szTempWheSql As String 'WHERE SQL
    Dim szTempGroSql As String 'GROUP BY SQL
    Dim szTempFroSql As String 'FROM SQL
    Dim oDb As New RTConnection
    Dim TempType As Integer
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
'    If SellStationID <> "" Then
'        szTempWheSql = " AND start_station_id = " & TransFieldValueToString(SellStationID)
'    Else
'        szTempWheSql = ""
'    End If
'    If AcceptType <> "" Then
'        szTempWheSql = szTempWheSql + " AND accept_type = " & TransFieldValueToString(GetLuggageTypeInt(AcceptType))
'    Else
'        szTempWheSql = szTempWheSql + ""
'    End If
'    szSql = "SELECT start_station_id,accept_type,MAX(start_station_name) start_station_name ,MAX(des_station_id) des_station_id,MAX(des_station_name) des_station_name, " _
'        & " MAX(mileage) mileage,MAX(bus_id) bus_id,MAX(bus_date) bus_date,MAX(luggage_name) luggage_name,SUM(cal_weight),SUM(fact_weight) fact_weight " _
'        & " ,MAX(start_label_id) start_label_id,MAX(end_label_id) end_label_id,COUNT(baggage_number) baggage_number,COUNT(over_weight_number) over_weight_number " _
'        & " ,SUM(p.price_total) price_total,SUM(p.price_item_1) price_item_1,SUM(p.price_item_2) price_item_2,SUM(p.price_item_3) price_item_3,SUM(p.price_item_4) price_item_4 " _
'        & " ,SUM(p.price_item_5) price_item_5,SUM(p.price_item_6) price_item_6,SUM(p.price_item_7) price_item_7,SUM(p.price_item_8) price_item_8,SUM(p.price_item_9) price_item_9 " _
'        & " ,SUM(p.price_item_10) price_item_10 " _
'        & " FROM luggage_accept_info l,luggage_price_lst p WHERE l.luggage_id = p.luggage_id " & szTempWheSql _
'        & " GROUP BY start_station_id ,accept_type "
    
    If SellStationID <> "" Then
        szTempSelSql = " s.sell_station_id ,MAX(i.sell_station_name) as sell_station_name  "
        szTempFroSql = " ,sell_station_info i "
        szTempWheSql = " AND s.sell_station_id =" & TransFieldValueToString(SellStationID) & " AND s.sell_station_id = i.sell_station_id "
        szTempGroSql = " s.sell_station_id  "
    Else
        szTempSelSql = " s.sell_station_id ,MAX(i.sell_station_name) as sell_station_name "
        szTempFroSql = " ,sell_station_info i "
        szTempWheSql = " AND s.sell_station_id = i.sell_station_id "
        szTempGroSql = " s.sell_station_id "
    End If
    If AcceptType <> "" Then
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type  "
        szTempWheSql = szTempWheSql + " AND accept_type = " & TransFieldValueToString(GetLuggageTypeInt(AcceptType))
        szTempGroSql = szTempGroSql + " ,accept_type "
    Else
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type"
        szTempWheSql = szTempWheSql + ""
        szTempGroSql = szTempGroSql + " ,accept_type "

    End If
    szSql = "SELECT " & szTempSelSql & "," _
        & " SUM(accept_number) accept_number ,SUM(accept_luggage_number) accept_luggage_number ,SUM(accept_calculate_weight) accept_calculate_weight ,SUM(accept_fact_weight) accept_fact_weight ,SUM(accept_over_number) accept_over_number ,SUM(accept_amount) accept_amount , " _
        & " SUM(cancel_number) cancel_number ,SUM(cancel_luggage_number) cancel_luggage_number ,SUM(cancel_calculate_weight) cancel_calculate_weight ,SUM(cancel_fact_weight) cancel_fact_weight ,SUM(cancel_over_number) cancel_over_number ,SUM(cancel_amount) cancel_amount , " _
        & " SUM(return_number) return_number ,SUM(return_luggage_number) return_luggage_number ,SUM(return_calculate_weight) return_calculate_weight ,SUM(return_fact_weight) return_fact_weight ,SUM(return_over_number) return_over_number ,SUM(return_amount) return_amount , " _
        & " SUM(return_charge) return_charge ,SUM(price_item_1) price_item_1 ,SUM(price_item_2) price_item_2 ,SUM(price_item_3) price_item_3 ,SUM(price_item_4) price_item_4 ,SUM(price_item_5) price_item_5 ,SUM(price_item_6) price_item_6,SUM(price_item_7) price_item_7, " _
        & " SUM(price_item_8) price_item_8 ,SUM(price_item_9) price_item_9 ,SUM(price_item_10) price_item_10 ,SUM(accept_amount-cancel_amount-return_amount+return_charge) total " _
        & " FROM stat_station_luggage_lst s " & szTempFroSql _
        & " WHERE bus_date >=" & TransFieldValueToString(StartDate) & " AND bus_date < " & TransFieldValueToString(EndDate) _
        & szTempWheSql _
        & " GROUP BY   " & szTempGroSql
    Set rsTemp = oDb.Execute(szSql)
    Set SellStationStat = rsTemp
    Exit Function
StationMonthDataStatErr:
'    Call RaiseError(MyUnhandledError, "LugDataStatSvr:StationMonthDataStat Method")iseError(MyUnhandledError, "LugDataStatSvr:LuggagerDataDayStat Method")
End Function

'行包员简报
Public Function LuggagerStat(StartDate As Date, EndDate As Date, UserID() As String, Optional SellStationID As String = "", Optional AcceptType As String = "") As ADODB.Recordset
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As ADODB.Recordset
    Dim szUserID As String
    Dim szTempFroSql As String
    Dim szTempSelSql As String
    Dim szTempWheSql As String
    Dim szTempGroSql As String
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    If ArrayLength(UserID) Then
        szUserID = "("
        Dim i As Integer
        For i = 1 To ArrayLength(UserID) - 1
            szUserID = szUserID & TransFieldValueToString(UserID(i)) & ","
        Next i
        szUserID = szUserID & TransFieldValueToString(UserID(i)) & ")"
        szTempFroSql = " ,user_info u"
        szTempSelSql = "s.user_id ,MAX(u.user_name) user_name  "
        szTempWheSql = " AND u.user_id = s.user_id "
        szTempGroSql = " s.user_id"
    Else
        szTempFroSql = " ,user_info u"
        szTempSelSql = " s.user_id ,MAX(u.user_name) user_name "
        szTempWheSql = " AND u.user_id = s.user_id "
        szTempGroSql = " s.user_id"
    End If
    If SellStationID <> "" Then
        szTempSelSql = szTempSelSql + " ,s.sell_station_id , MAX(i1.sell_station_name) sell_station_name  "
        szTempFroSql = szTempFroSql + " ,sell_station_info i1 "
        szTempWheSql = szTempWheSql + " AND s.sell_station_id =" & TransFieldValueToString(SellStationID) & " AND s.sell_station_id = i1.sell_station_id "
        szTempGroSql = szTempGroSql + " ,s.sell_station_id  "
    Else
        szTempSelSql = szTempSelSql + " ,s.sell_station_id , MAX(i1.sell_station_name) sell_station_name "
        szTempFroSql = szTempFroSql + " ,sell_station_info i1 "
        szTempWheSql = szTempWheSql + " AND s.sell_station_id = i1.sell_station_id "
        szTempGroSql = szTempGroSql + " ,s.sell_station_id"
    End If
    
     If AcceptType <> "" Then
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type "
        szTempWheSql = szTempWheSql + " AND accept_type = " & TransFieldValueToString(GetLuggageTypeInt(AcceptType))
        szTempGroSql = szTempGroSql + " ,accept_type"
    Else
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type "
        szTempWheSql = szTempWheSql + ""
        szTempGroSql = szTempGroSql + " ,accept_type"
    End If
    szSql = "SELECT " & szTempSelSql & "," _
        & " SUM(accept_number) accept_number ,SUM(accept_luggage_number) accept_luggage_number ,SUM(accept_calculate_weight) accept_calculate_weight ,SUM(accept_fact_weight) accept_fact_weight ,SUM(accept_over_number) accept_over_number ,SUM(accept_amount) accept_amount , " _
        & " SUM(cancel_number) cancel_number ,SUM(cancel_luggage_number) cancel_luggage_number ,SUM(cancel_calculate_weight) cancel_calculate_weight ,SUM(cancel_fact_weight) cancel_fact_weight ,SUM(cancel_over_number) cancel_over_number ,SUM(cancel_amount) cancel_amount , " _
        & " SUM(return_number) return_number ,SUM(return_luggage_number) return_luggage_number ,SUM(return_calculate_weight) return_calculate_weight ,SUM(return_fact_weight) return_fact_weight ,SUM(return_over_number) return_over_number ,SUM(return_amount) return_amount , " _
        & " SUM(return_charge) return_charge ,SUM(price_item_1) price_item_1 ,SUM(price_item_2) price_item_2 ,SUM(price_item_3) price_item_3 ,SUM(price_item_4) price_item_4 ,SUM(price_item_5) price_item_5 ,SUM(price_item_6) price_item_6 ,SUM(price_item_7) price_item_7 , " _
        & " SUM(price_item_8) price_item_8 ,SUM(price_item_9) price_item_9 ,SUM(price_item_10) price_item_10 ,SUM(accept_amount-cancel_amount-return_amount+return_charge) total " _
        & " FROM stat_saler_luggage_lst s" & szTempFroSql _
        & " WHERE operation_time >=" & TransFieldValueToString(StartDate) & " AND operation_time <= " & TransFieldValueToString(EndDate) _
        & IIf(szUserID <> "", " AND s.user_id IN " & szUserID, " ") & szTempWheSql _
        & " GROUP BY " & szTempGroSql
    Set rsTemp = oDb.Execute(szSql)
    Set LuggagerStat = rsTemp
    Exit Function
End Function

'车站简报
Public Function StationStat(SellStationID As String, StartDate As Date, EndDate As Date, StationID As String, Optional AcceptType As String = "") As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_LuggageDataStatSvr
    AssertHaveRight m_oActiveUser, RIGHT_StationDataStatSvr
    On Error GoTo StationDayDataStatErr
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szTempSelSql As String      'SELECT SQL
    Dim szTempWheSql As String      'WHERE SQL
    Dim szTempGroSql As String      'GROUP BY SQL
    Dim szTempFroSql As String      'FROM SQL
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr(cszLuggage)
    
    If SellStationID <> "" Then
        szTempSelSql = " s.sell_station_id, MAX(i1.sell_station_name) as sell_station_name  "
        szTempFroSql = " ,sell_station_info i1 "
        szTempWheSql = " AND s.sell_station_id =" & TransFieldValueToString(SellStationID) & " AND s.sell_station_id=i1.sell_station_id "
        szTempGroSql = " s.sell_station_id  "
    Else
        szTempSelSql = " s.sell_station_id ,MAX(i1.sell_station_name) as sell_station_name "
        szTempFroSql = " ,sell_station_info i1 "
        szTempWheSql = " AND s.sell_station_id=i1.sell_station_id "
        szTempGroSql = " s.sell_station_id"
    End If
    
    If StationID <> "" Then
        szTempSelSql = szTempSelSql + " ,s.station_id,MAX(s.station_name) station_name"
        szTempWheSql = szTempWheSql + " AND  s.station_id = " & TransFieldValueToString(StationID)
        szTempGroSql = szTempGroSql + " ,s.station_id "
    Else
        szTempSelSql = szTempSelSql + " ,s.station_id,MAX(s.station_name) station_name "
        szTempWheSql = szTempWheSql + ""
        szTempGroSql = szTempGroSql + " ,s.station_id "
    End If
    
    If AcceptType <> "" Then
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type   "
        szTempWheSql = szTempWheSql + " AND accept_type = " & TransFieldValueToString(GetLuggageTypeInt(AcceptType))
        szTempGroSql = szTempGroSql + " ,accept_type"
    Else
        szTempSelSql = szTempSelSql + " ,case accept_type when 0 then " & "'" & GetLuggageTypeString(0) & "'" & " when 1 then " & "'" & GetLuggageTypeString(1) & "'" & " end accept_type  "
        szTempWheSql = szTempWheSql + ""
        szTempGroSql = szTempGroSql + " ,accept_type "
    End If
    szSql = "SELECT " & szTempSelSql & "," _
        & " SUM(accept_number) accept_number ,SUM(accept_luggage_number) accept_luggage_number ,SUM(accept_calculate_weight) accept_calculate_weight ,SUM(accept_fact_weight) accept_fact_weight ,SUM(accept_over_number) accept_over_number ,SUM(accept_amount) accept_amount , " _
        & " SUM(cancel_number) cancel_number ,SUM(cancel_luggage_number) cancel_luggage_number ,SUM(cancel_calculate_weight) cancel_calculate_weight ,SUM(cancel_fact_weight) cancel_fact_weight ,SUM(cancel_over_number) cancel_over_number ,SUM(cancel_amount) cancel_amount , " _
        & " SUM(return_number) return_number ,SUM(return_luggage_number) return_luggage_number ,SUM(return_calculate_weight) return_calculate_weight ,SUM(return_fact_weight) return_fact_weight ,SUM(return_over_number) return_over_number ,SUM(return_amount) return_amount , " _
        & " SUM(return_charge) return_charge ,SUM(price_item_1) price_item_1 ,SUM(price_item_2) price_item_2 ,SUM(price_item_3) price_item_3 ,SUM(price_item_4) price_item_4 ,SUM(price_item_5) price_item_5 ,SUM(price_item_6) price_item_6,SUM(price_item_7) price_item_7, " _
        & " SUM(price_item_8) price_item_8 ,SUM(price_item_9) price_item_9 ,SUM(price_item_10) price_item_10 ,SUM(accept_amount-cancel_amount-return_amount+return_charge) total " _
        & " FROM stat_station_luggage_lst s " & szTempFroSql _
        & " WHERE bus_date >=" & TransFieldValueToString(StartDate) & " AND bus_date < " & TransFieldValueToString(EndDate) _
        & szTempWheSql _
        & " GROUP BY " & szTempGroSql
    Set rsTemp = oDb.Execute(szSql)
    Set StationStat = rsTemp
    Exit Function
StationDayDataStatErr:
'    Call RaiseError(MyUnhandledError, "LugDataStatSvr:StationDayDataStat Method")
End Function





Public Function AcceptEveryDaySellDetail(UserID As String, StartDate As Date, EndDate As Date) As Recordset
'行包员受理的明细
    AssertActiveUserValid m_oActiveUser, ERR_LuggageDataStatSvr
    AssertHaveRight m_oActiveUser, RIGHT_LuggageSvrTotalAccept
    
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr("")
    
    szSql = "SELECT * FROM luggage_accept_info a, luggage_price_lst p  WHERE a.luggage_id = p.luggage_id AND " _
    & " user_id='" & UserID & "' AND " _
    & " operation_time>='" & ToDBDateTime(StartDate) & "' AND " _
    & " operation_time<'" & ToDBDateTime(EndDate) & "'" _
    & " AND status <> " & TransFieldValueToString(ST_LuggageAcceptCancel) _
    & " AND status <> " & TransFieldValueToString(ST_LuggageAcceptReturn) & "  ORDER BY a.luggage_id"
      
    
    Set rsTemp = oDb.Execute(szSql)
    Set AcceptEveryDaySellDetail = rsTemp
    Set oDb = Nothing
    
End Function


Public Function AcceptEveryDaySellTotal(UserID As String, StartDate As Date, EndDate As Date) As Double
'行包员受理的明细
    AssertActiveUserValid m_oActiveUser, ERR_LuggageDataStatSvr
    AssertHaveRight m_oActiveUser, RIGHT_LuggageSvrTotalAccept
    
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr("")
    
    szSql = "SELECT sum(p.price_total) price_total FROM luggage_accept_info a, luggage_price_lst p  WHERE a.luggage_id = p.luggage_id AND " _
    & " user_id='" & UserID & "' AND " _
    & " operation_time>='" & ToDBDateTime(StartDate) & "' AND " _
    & " operation_time<'" & ToDBDateTime(EndDate) & "'" _

      
    
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        AcceptEveryDaySellTotal = FormatDbValue(rsTemp!price_total)
    End If
    Set oDb = Nothing
    
End Function


'数组相应为废、退、改
'行包员的退废签信息
Public Function AcceptEveryDayAnotherThing(UserID As String, StartDate As Date, EndDate As Date) As Double()
    AssertActiveUserValid m_oActiveUser, ERR_LuggageDataStatSvr
    AssertHaveRight m_oActiveUser, RIGHT_LuggageSvrTotalAccept
    
    Dim adbResult(1 To 2, 1 To 3) As Double
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    
    oDb.ConnectionString = GetConnectionStr("")
    szSql = "SELECT COUNT(*) AS countx ,SUM(p.price_total) AS total_total_price FROM " _
    & " luggage_accept_info a ,luggage_cancel_lst , luggage_price_lst p WHERE " _
    & " a.luggage_Id=luggage_cancel_lst.luggage_Id AND a.luggage_id = p.luggage_id AND " _
    & " a.user_id ='" & UserID & "' AND " _
    & " luggage_cancel_lst.cancel_time>='" & ToDBDateTime(StartDate) & "' AND " _
    & " luggage_cancel_lst.cancel_time<'" & ToDBDateTime(EndDate) & "'"
    
    
    Set rsTemp = oDb.Execute(szSql)
    adbResult(1, 1) = rsTemp!countx
    adbResult(1, 2) = FormatDbValue(rsTemp!total_total_price)
    
    szSql = "SELECT COUNT(*) AS countx ,SUM(p.price_total) AS total_total_price, " _
    & " SUM(return_charge) AS total_return_chare   FROM " _
    & " luggage_return_lst a, luggage_price_lst p WHERE " _
    & " user_id='" & UserID & "' AND " _
    & " operation_time>='" & ToDBDateTime(StartDate) & "' AND  a.luggage_id = p.luggage_id AND " _
    & " operation_time<'" & ToDBDateTime(EndDate) & "'"
    
    Set rsTemp = oDb.Execute(szSql)
    adbResult(2, 1) = rsTemp!countx
    adbResult(2, 2) = FormatDbValue(rsTemp!total_total_price)
    adbResult(2, 3) = FormatDbValue(rsTemp!total_return_chare)
    
'
'    '以下是改签票处理
'    szSql = "SELECT COUNT(*) AS countx ,SUM(former_ticket_price) AS total_total_price, " _
'    & " SUM(change_charge) AS total_change_chare FROM " _
'    & " change_luggage_lst WHERE " _
'    & " user_id='" & UserID & "' AND " _
'    & " change_time>='" & ToDBDateTime(StartDate) & "' AND " _
'    & " change_time<'" & ToDBDateTime(EndDate) & "'"
'
'    Set rsTemp = oDb.Execute(szSql)
'    adbResult(3, 1) = rsTemp!countx
'    adbResult(3, 2) = FormatDbValue(rsTemp!total_total_price)
'    adbResult(3, 3) = FormatDbValue(rsTemp!total_change_chare)
'    '改签票处理到这里完
    
    
    AcceptEveryDayAnotherThing = adbResult
    
    Set oDb = Nothing
    
End Function

'得到受理员的收费项价格
Public Function GetAccepterPriceDetail(UserID As String, StartDate As Date, EndDate As Date) As Double()
    AssertActiveUserValid m_oActiveUser, ERR_LuggageDataStatSvr
    AssertHaveRight m_oActiveUser, RIGHT_LuggageSvrTotalAccept
    
    Dim adbResult(1 To 11) As Double
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim szPrice As String
    Dim i As Integer
    
    Dim rsReturn As Recordset
    Dim rsCancel As Recordset
    Dim rsChange As Recordset
    
    oDb.ConnectionString = GetConnectionStr("")
    
    
    For i = 1 To 10
        szPrice = szPrice & "SUM(price_item_" & i & ") price_item_" & i & ","
    Next i
    
    szSql = " SELECT " _
        & szPrice & " SUM(price_total) ticket_price_total " _
        & " FROM luggage_accept_info a , luggage_price_lst p WHERE a.luggage_id = p.luggage_id " _
        & " AND user_id = '" & UserID & "' AND operation_time >= '" & ToDBDateTime(StartDate) & "'" _
        & " AND operation_time < '" & ToDBDateTime(EndDate) & "'"
        
    Set rsTemp = oDb.Execute(szSql)
    
    szSql = " SELECT " _
        & szPrice & " SUM(price_total) ticket_price_total " _
        & " FROM luggage_return_lst r  , luggage_accept_info a , luggage_price_lst p  WHERE r.luggage_id = a.luggage_id and r.luggage_id = p.luggage_id  " _
        & " AND r.user_id = '" & UserID & "' AND r.operation_time >= '" & ToDBDateTime(StartDate) & "'" _
        & " AND r.operation_time < '" & ToDBDateTime(EndDate) & "'"
        
    Set rsReturn = oDb.Execute(szSql)
    
    szSql = " SELECT " _
        & szPrice & " SUM(price_total) ticket_price_total " _
        & " FROM luggage_cancel_lst c  , luggage_accept_info a , luggage_price_lst p  " _
        & " WHERE c.luggage_id = a.luggage_id and c.luggage_id = p.luggage_id  " _
        & " AND a.user_id = '" & UserID & "' AND cancel_time >= '" & ToDBDateTime(StartDate) & "'" _
        & " AND cancel_time < '" & ToDBDateTime(EndDate) & "'"
        
    Set rsCancel = oDb.Execute(szSql)
'
'    szSql = " SELECT " _
'        & szPrice & " SUM(ticket_price_total) ticket_price_total " _
'        & " FROM luggage_accept_info a , luggage_price_lst p WHERE a.luggage_id = p.luggage_id " _
'        & " AND user_id = '" & UserID & "' AND operation_time >= '" & ToDBDateTime(StartDate) & "'" _
'        & " AND operation_time < '" & ToDBDateTime(EndDate) & "'"
'
'    Set rsChange = odb.Execute(szSql)
'
    
    For i = 1 To 10
        adbResult(i) = FormatDbValue(rsTemp("price_item_" & i)) - FormatDbValue(rsReturn("price_item_" & i)) - FormatDbValue(rsCancel("price_item_" & i))
    Next i
    adbResult(i) = FormatDbValue(rsTemp!ticket_price_total) - FormatDbValue(rsReturn!ticket_price_total) - FormatDbValue(rsCancel!ticket_price_total)
    
    GetAccepterPriceDetail = adbResult
    
    Set oDb = Nothing
    
End Function

'拆帐公司签发简报
Public Function SplitCompanyCheckStat(pszSellStationID As String, pdtStartDate As Date, pdtEndDate As Date, pszSplitCompanyID As String, Optional pnAcceptType As Integer = -1) As Recordset
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    
    On Error GoTo ErrorHandle
    
    AssertActiveUserValid m_oActiveUser, ERR_LuggageDataStatSvr
    AssertHaveRight m_oActiveUser, RIGHT_SplitCompanyCheckStat '拆帐公司签发简报权限
    
    oDb.ConnectionString = GetConnectionStr
    szSql = "  SELECT transport_company_id split_company_id , MAX(transport_company_short_name) split_company_short_name , COUNT(*) check_number" _
        & " , SUM(luggage_number) luggage_number , SUM(baggage_number) baggage_number , SUM(cal_weight) cal_weight , SUM(fact_weight) fact_weight " _
        & " , SUM(over_number) over_number , SUM(total_price) total_price , SUM(print_settle_price) print_settle_price " & GetSumPriceItemStr _
        & " FROM luggage_sheet_info c " _
        & " WHERE c.sheet_make_time >= " & TransFieldValueToString(pdtStartDate) _
        & " AND c.sheet_make_time < " & TransFieldValueToString(pdtEndDate) _
        & " AND c.valid_mark = " & TransFieldValueToString(ELuggageSheetValidStatus.ST_LuggageSheetValidStatusNormal)
    
    
    If pszSellStationID <> "" Then
        szSql = szSql & " AND c.sell_station_id =  " & TransFieldValueToString(pszSellStationID)
    End If
    
    If pszSplitCompanyID <> "" Then
        szSql = szSql & " AND c.transport_company_id =  " & TransFieldValueToString(pszSplitCompanyID)
    End If
    If pnAcceptType <> -1 Then
        szSql = szSql & " AND c.accept_type = " & TransFieldValueToString(pnAcceptType)
    End If
    szSql = szSql & " GROUP BY c.transport_company_id "
    
    Set rsTemp = oDb.Execute(szSql)
    
    Set SplitCompanyCheckStat = rsTemp
    Set rsTemp = Nothing
    Set oDb = Nothing

    Exit Function
ErrorHandle:
    Set rsTemp = Nothing
    Set oDb = Nothing
    err.Raise err.Number, err.Source, err.Description
    
End Function

'车辆签发简报权限
Public Function VehicleCheckStat(pszSellStationID As String, pdtStartDate As Date, pdtEndDate As Date, pszVehicle As String, Optional pnAcceptType As Integer = -1) As Recordset
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    
    On Error GoTo ErrorHandle
    
    AssertActiveUserValid m_oActiveUser, ERR_LuggageDataStatSvr
    AssertHaveRight m_oActiveUser, RIGHT_VehicleCheckStat  '车辆签发简报权限
    
    oDb.ConnectionString = GetConnectionStr
    szSql = "  SELECT vehicle_id , MAX(license_tag_no) license_tag_no , COUNT(*) check_number" _
        & " , SUM(luggage_number) luggage_number , SUM(baggage_number) baggage_number , SUM(cal_weight) cal_weight , SUM(fact_weight) fact_weight " _
        & " , SUM(over_number) over_number , SUM(total_price) total_price  , SUM(print_settle_price) print_settle_price " & GetSumPriceItemStr _
        & " FROM luggage_sheet_info c " _
        & " WHERE c.sheet_make_time >= " & TransFieldValueToString(pdtStartDate) _
        & " AND c.sheet_make_time < " & TransFieldValueToString(pdtEndDate) _
        & " AND c.valid_mark = " & TransFieldValueToString(ELuggageSheetValidStatus.ST_LuggageSheetValidStatusNormal)
    
    
    If pszSellStationID <> "" Then
        szSql = szSql & " AND c.sell_station_id =  " & TransFieldValueToString(pszSellStationID)
    End If
    
    If pszVehicle <> "" Then
        szSql = szSql & " AND c.vehicle_id =  " & TransFieldValueToString(pszVehicle)
    End If
    If pnAcceptType <> -1 Then
        szSql = szSql & " AND c.accept_type = " & TransFieldValueToString(pnAcceptType)
    End If
    szSql = szSql & " GROUP BY c.vehicle_id "
    
    Set rsTemp = oDb.Execute(szSql)
    
    Set VehicleCheckStat = rsTemp
    Set rsTemp = Nothing
    Set oDb = Nothing

    Exit Function
ErrorHandle:
    Set rsTemp = Nothing
    Set oDb = Nothing
    err.Raise err.Number, err.Source, err.Description
    
End Function


'线路签发简报权限
Public Function RouteCheckStat(pszSellStationID As String, pdtStartDate As Date, pdtEndDate As Date, pszRoute As String, Optional pnAcceptType As Integer = -1) As Recordset
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    
    On Error GoTo ErrorHandle
    
    AssertActiveUserValid m_oActiveUser, ERR_LuggageDataStatSvr
    AssertHaveRight m_oActiveUser, RIGHT_VehicleCheckStat
    
    oDb.ConnectionString = GetConnectionStr
    szSql = "  SELECT route_id , MAX(route_name) route_name , COUNT(*) check_number" _
        & " , SUM(luggage_number) luggage_number , SUM(baggage_number) baggage_number , SUM(cal_weight) cal_weight , SUM(fact_weight) fact_weight " _
        & " , SUM(over_number) over_number , SUM(total_price) total_price , SUM(print_settle_price) print_settle_price " & GetSumPriceItemStr _
        & " FROM luggage_sheet_info c " _
        & " WHERE c.sheet_make_time >= " & TransFieldValueToString(pdtStartDate) _
        & " AND c.sheet_make_time < " & TransFieldValueToString(pdtEndDate) _
        & " AND c.valid_mark = " & TransFieldValueToString(ELuggageSheetValidStatus.ST_LuggageSheetValidStatusNormal)
    
    
    If pszSellStationID <> "" Then
        szSql = szSql & " AND c.sell_station_id =  " & TransFieldValueToString(pszSellStationID)
    End If
    
    If pszRoute <> "" Then
        szSql = szSql & " AND c.route_id =  " & TransFieldValueToString(pszRoute)
    End If
    If pnAcceptType <> -1 Then
        szSql = szSql & " AND c.accept_type = " & TransFieldValueToString(pnAcceptType)
    End If
    szSql = szSql & " GROUP BY c.route_id "
    
    Set rsTemp = oDb.Execute(szSql)
    
    Set RouteCheckStat = rsTemp
    Set rsTemp = Nothing
    Set oDb = Nothing

    Exit Function
ErrorHandle:
    Set rsTemp = Nothing
    Set oDb = Nothing
    err.Raise err.Number, err.Source, err.Description
    
End Function

Private Function GetSumPriceItemStr()
    Dim szSql As String
    Dim i As Integer
    For i = 1 To 10
        szSql = szSql & " , SUM(price_item_" & i & ") price_item_" & i & " "
    Next i
    GetSumPriceItemStr = szSql
End Function





