Attribute VB_Name = "mdStation"
Option Explicit
'����ϵͳ�����е�ģ��



'�ɻ�����Ʊ
Public Const CnInternetCanSell = 0
Public Const CnInternetNotCanSell = 1

'Ĭ����λ����
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

Public Enum eQueryMode             '��ѯģʽ
    BlurQuery = 1         'ģ����ѯ
    CustomQuery = 2       'ָ����ѯ
End Enum


Public Function GetTicketStatusStr(nStatus As Integer) As String
    Dim szTemp As String
    If (nStatus And ST_TicketNormal) <> 0 Then
        szTemp = "�����۳�"
    Else
        szTemp = "��ǩ�۳�"
    End If
    
    If (nStatus And ST_TicketCanceled) <> 0 Then
        szTemp = szTemp & "/�ѷ�"
    ElseIf (nStatus And ST_TicketChanged) <> 0 Then
        szTemp = szTemp & "/�ѱ���ǩ"
    ElseIf (nStatus And ST_TicketChecked) <> 0 Then
        szTemp = szTemp & "/�Ѽ�"
    ElseIf (nStatus And ST_TicketReturned) <> 0 Then
        szTemp = szTemp & "/����"
    End If
    GetTicketStatusStr = szTemp
End Function


Public Function GetTicketTypeName(nTicketType As Integer) As String
    Dim szTemp As String
    'ֻ����ȫƱ���Ʊ,�����Ķ����ؿ�
    Select Case nTicketType
    Case TP_FullPrice
        szTemp = "ȫƱ"
    Case TP_HalfPrice
        szTemp = "��Ʊ"
    Case Else
        szTemp = ""
    End Select
    GetTicketTypeName = szTemp
End Function

'�õ�����վ��������������Ϣ�ַ���
Public Function GetStationLimitedCountStr(pnCount As Integer) As String
    If pnCount < 0 Then
        GetStationLimitedCountStr = "����"
    ElseIf pnCount = 0 Then
        GetStationLimitedCountStr = "������"
    Else
        GetStationLimitedCountStr = "����" & pnCount & "��"
    End If
End Function

'�õ�����վ������ʱ�����Ϣ�ַ���
Public Function GetStationLimitedTimeStr(pnTime As Integer, pdtDate As Date, pdtOffTime As Date) As String
    Dim dtFullDateTime As Date
    If pnTime <= 0 Then
        GetStationLimitedTimeStr = "����"
    Else
        dtFullDateTime = DateAdd("h", -pnTime, CDate(Format(pdtDate, "YYYY-MM-DD") & " " & Format(pdtOffTime, "HH:mm")))
        GetStationLimitedTimeStr = "��" & Format(dtFullDateTime, "DD��HH:mm") & "��ſ���"
    End If
End Function



Public Function getCheckStatusStr(nStatus As Integer) As String
    Select Case nStatus
        Case EREBusStatus.ST_BusChecking
            getCheckStatusStr = "���ڼ�Ʊ"
        Case EREBusStatus.ST_BusExtraChecking
            getCheckStatusStr = "���ڲ���"
        Case EREBusStatus.ST_BusMergeStopped
            getCheckStatusStr = "����ͣ��"
        Case EREBusStatus.ST_BusNormal
            getCheckStatusStr = "δ��"
        Case EREBusStatus.ST_BusStopCheck
            getCheckStatusStr = "ͣ��"
        Case EREBusStatus.ST_BusExtraChecking
            getCheckStatusStr = "���ڲ���"
        Case EREBusStatus.ST_BusStopped
            getCheckStatusStr = "����ͣ��"
    End Select
End Function
Public Function getCheckedTicketStatus(nStatus As Integer) As String
    Select Case nStatus
        Case ECheckedTicketStatus.NormalTicket
            getCheckedTicketStatus = "��������"
        Case ECheckedTicketStatus.ChangedTicket
            getCheckedTicketStatus = "�ĳ˼���"
        Case ECheckedTicketStatus.MergedTicket
            getCheckedTicketStatus = "�������"
    End Select
End Function
