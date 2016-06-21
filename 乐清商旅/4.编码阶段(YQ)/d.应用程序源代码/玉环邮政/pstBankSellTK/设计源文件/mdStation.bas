Attribute VB_Name = "mdStation"
Option Explicit
'客运系统所共有的模块



'可互联售票
Public Const CnInternetCanSell = 0
Public Const CnInternetNotCanSell = 1

'默认座位类型
Public Const cszSeatTypeIsNormal = "01"
Public Const cszSeatType = "01"
Public Const cszBedType = "02"
Public Const cszAdditionalType = "03"
Public Const cszOtherType1 = "04"
Public Const cszOtherType2 = "05"


Public Const CnSlitpStatus = 32

Public Const TP_TicketTypeCount = 6



'Public cdtEmptyDateTime As Date


Public Enum ECheckedTicketStatus
    NormalTicket = 1
    ChangedTicket = 2
    MergedTicket = 3
End Enum

Public Enum eQueryMode             '查询模式
    BlurQuery = 1         '模糊查询
    CustomQuery = 2       '指定查询
End Enum


Public Function GetTicketStatusStr(nStatus As Integer) As String
    Dim szTemp As String
    If (nStatus And ST_TicketNormal) <> 0 Then
        szTemp = "正常售出"
    Else
        szTemp = "改签售出"
    End If
    
    If (nStatus And ST_TicketCanceled) <> 0 Then
        szTemp = szTemp & "/已废"
    ElseIf (nStatus And ST_TicketChanged) <> 0 Then
        szTemp = szTemp & "/已被改签"
    ElseIf (nStatus And ST_TicketChecked) <> 0 Then
        szTemp = szTemp & "/已检"
    ElseIf (nStatus And ST_TicketReturned) <> 0 Then
        szTemp = szTemp & "/已退"
    End If
    GetTicketStatusStr = szTemp
End Function


Public Function GetTicketTypeName(nTicketType As Integer) As String
    Dim szTemp As String
    '只返回全票与半票,其他的都返回空
    Select Case nTicketType
    Case TP_FullPrice
        szTemp = "全票"
    Case TP_HalfPrice
        szTemp = "半票"
    Case Else
        szTemp = ""
    End Select
    GetTicketTypeName = szTemp
End Function

'得到车次站点限售张数的信息字符串
Public Function GetStationLimitedCountStr(pnCount As Integer) As String
    If pnCount < 0 Then
        GetStationLimitedCountStr = "不限"
    ElseIf pnCount = 0 Then
        GetStationLimitedCountStr = "不可售"
    Else
        GetStationLimitedCountStr = "限售" & pnCount & "张"
    End If
End Function

'得到车次站点限售时间的信息字符串
Public Function GetStationLimitedTimeStr(pnTime As Integer, pdtDate As Date, pdtOffTime As Date) As String
    Dim dtFullDateTime As Date
    If pnTime <= 0 Then
        GetStationLimitedTimeStr = "不限"
    Else
        dtFullDateTime = DateAdd("h", -pnTime, CDate(Format(pdtDate, "YYYY-MM-DD") & " " & Format(pdtOffTime, "HH:mm")))
        GetStationLimitedTimeStr = "在" & Format(dtFullDateTime, "DD日HH:mm") & "后才可售"
    End If
End Function



Public Function getCheckStatusStr(nStatus As Integer) As String
    Select Case nStatus
        Case EREBusStatus.ST_BusChecking
            getCheckStatusStr = "正在检票"
        Case EREBusStatus.ST_BusExtraChecking
            getCheckStatusStr = "正在补检"
        Case EREBusStatus.ST_BusMergeStopped
            getCheckStatusStr = "并班停检"
        Case EREBusStatus.ST_BusNormal
            getCheckStatusStr = "未检"
        Case EREBusStatus.ST_BusStopCheck
            getCheckStatusStr = "停检"
        Case EREBusStatus.ST_BusExtraChecking
            getCheckStatusStr = "正在补检"
        Case EREBusStatus.ST_BusStopped
            getCheckStatusStr = "车次停班"
    End Select
End Function
Public Function getCheckedTicketStatus(nStatus As Integer) As String
    Select Case nStatus
        Case ECheckedTicketStatus.NormalTicket
            getCheckedTicketStatus = "正常检入"
        Case ECheckedTicketStatus.ChangedTicket
            getCheckedTicketStatus = "改乘检入"
        Case ECheckedTicketStatus.MergedTicket
            getCheckedTicketStatus = "并班检入"
    End Select
End Function
