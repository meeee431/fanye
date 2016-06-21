Attribute VB_Name = "mdlMain"
Option Explicit
Public Const cnMaxLoginErrCount = 3

'对话框调用状态
Public Enum eFormStatus
    EFS_AddNew = 0
    EFS_Modify = 1
    EFS_Show = 2
End Enum

Public Type Sheet_ChkInfo
    SheetNo As String
    BusID As String
    dtTime As Date
    Checker As String
    StartUpTime As Date
    Route As String
    Company As String
    VehicleTag As String
    Owner As String
    SerialNo As Integer
End Type
Public Type SheetContent           '路单具体信息
    StationId As String
    StationName As String
    TicketTypeID As Integer
    FullTk_Numer As Integer
    FullTk_Price As Single
    HalfTk_Numer As Integer
    HalfTk_Price As Single
    PreferentialTk1_Numer As Integer
    PreferentialTk1_Price As Single
    PreferentialTk2_Numer As Integer
    PreferentialTk2_Price As Single
    PreferentialTk3_Numer As Integer
    PreferentialTk3_Price As Single
End Type
Public Enum ECheckedTicketStatus        '车票检票状态
    NormalTicket = 1
    ChangedTicket = 2
    MergedTicket = 3
End Enum

'结算状态
Public Const mStatusNo = "未结"     '0
Public Const mStatusReal = "已结"   '1
Public Const mStatusCancel = "作废" '2


Public Const szAcceptTypeGeneral = "快件" '托运方式
Public Const szAcceptTypeMan = "随行"



Public Function GetCheckedTicketStatus(nStatus As Integer) As String
    Select Case nStatus
        Case ECheckedTicketStatus.NormalTicket
            GetCheckedTicketStatus = "正常检入"
        Case ECheckedTicketStatus.ChangedTicket
            GetCheckedTicketStatus = "改乘检入"
        Case ECheckedTicketStatus.MergedTicket
            GetCheckedTicketStatus = "并班检入"
    End Select
End Function

 '结算状态转换
Public Function GetFinTypeString(pnType As Integer) As String
    Select Case pnType
        Case 0
            GetFinTypeString = mStatusCancel
        Case 1
            GetFinTypeString = mStatusReal
'        Case 2
'            GetFinTypeString = mStatusCancel
    End Select
End Function



Public Function GetLuggageTypeString(pnType As Integer) As String
    Select Case pnType
        Case 0
            GetLuggageTypeString = szAcceptTypeGeneral
        Case 1
            GetLuggageTypeString = szAcceptTypeMan
    End Select
End Function

