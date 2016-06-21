Attribute VB_Name = "mdlDefine"
Option Explicit

'ģ�鶨��
'����Ķ����Ǵ�ԭ�ȵ��м�㶨�帴�ƹ�����


'��Ʊ����
Public Enum ETicketType
    TP_FullPrice = 1 'ȫƱ
    TP_HalfPrice = 2 ' ��Ʊ
    TP_FreeTicket = 3 '��Ʊ
    TP_PreferentialTicket1 = 4 '�Ż�Ʊ1
    TP_PreferentialTicket2 = 5 '�Ż�Ʊ2
    TP_PreferentialTicket3 = 6 '�Ż�Ʊ3
End Enum


Public Enum ETicketTypeValid
    TP_TicketTypeValid = 1 '����
    TP_TicketTypeNotValid = 0 '������
    TP_TicketTypeAll = 2 '����Ʊ��
End Enum

'��Ʊ״̬
Public Enum ETicketStatus
    ST_TicketNormal = 1 '��Ʊ�����۳�
    ST_TicketSellChange = 2 ' ��Ʊ��ǩ�۳�
    ST_TicketCanceled = 4 '��Ʊ�ѷ�
    ST_TicketChanged = 8 '����ǩ
    ST_TicketReturned = 16 '��Ʊ����
    ST_TicketChecked = 32 '��Ʊ�Ѽ�
End Enum

'��������״̬
Public Enum EREBusStatus
    ST_BusNormal = 0 '��������
    ST_BusStopped = 1 ' �����ѱ�ͣ��
    ST_BusMergeStopped = 2 '�����ѱ�����ͣ��
    ST_BusStopCheck = 3 '�����Ѿ�ͣ��
    ST_BusChecking = 4 '�������ڼ�Ʊ
    ST_BusExtraChecking = 5 '�������ڲ���
    ST_BusSlitpStop = 8 '���β��ͣ��
    ST_BusReplace = 16 '���ζ���
    'ST_BusLock = 32+ x   other +32
End Enum


'������λ״̬
Public Enum ERESeatStatus
    ST_SeatCanSell = 0 '����λ����
    ST_SeatReserved = 1 ' ����λ�ѱ�Ԥ��
    ST_SeatBooked = 2 '����λ�ѱ�Ԥ��
    ST_SeatSold = 3 '����λ�ѱ��۳�
    ST_SeatSlitp = 4 '����λ�ѱ��۳�,��ֵõ�
    ST_SeatReplace = 5 '����λ�ѱ��۳�������õ�
    ST_SeatMerge = 6 '����λ�ѱ��۳�������õ�
    ST_SeatProjectBooked = 64 '�ƻ�Ԥ��
End Enum
'��������
Public Enum EBusType
    TP_RegularBus = 0 '�̶�����
    TP_ScrollBus = 1 ' ��ˮ���Σ��������Σ�
End Enum

Public Type TBuyTicketInfo
    nTicketType As ETicketType
    szTicketNo As String
    szSeatNo As String '�ձ�ʾϵͳ���Ѹ���λ�ţ�'ST'��ʾվƱ
    szReserved As String
    szSeatTypeID As String  '��λ���ʹ���
    szSeatTypeName As String '��λ��������
    
End Type



'�˶�����Ʊ���ݲ���
Public Type TSellTicketParam
    BuyTicketInfo() As TBuyTicketInfo
    pasgSellTicketPrice() As Single
    aszOrgTicket() As String    '��ǩ��
    aszChangeSheetID() As String   '��ǩ��
End Type

'��Ʊ���
Public Type TSellTicketResult
    asgTicketPrice() As Double
    aszSeatNo() As String
    aszTicketType() As ETicketType
    aszSeatType() As String
    
    szDesStationName As String
    
    szBeginStationID As String
    szBeginStationName As String
    
    szReserved As String
End Type


Public Type TTicketType
    nTicketTypeID As Integer
    szTicketTypeName As String
    nTicketTypeValid As Integer
    szAnnotation As String
End Type

'�����ǹ��õ���������
Public Type TBusOrderCount
    szStatioinID As String
    dbCount As Double
End Type

'�õ�������Ϣ
Public Function GetBusRs(pdyBusDate As Date, pszStationID As String) As Recordset
    
End Function

''�õ�����վ��
'Public Function GetAllStationRs() As Recordset
'
'End Function

'�õ�����Ʊ��
Public Function GetAllTicketType(Optional pnTicketTypeID As Integer) As TTicketType()
    Dim atTemp() As TTicketType
    ReDim atTemp(1 To 6)
    
    
    atTemp(1).nTicketTypeID = 1
    atTemp(1).szTicketTypeName = "ȫƱ"
    atTemp(1).nTicketTypeValid = 1
    atTemp(1).szAnnotation = ""

    atTemp(2).nTicketTypeID = 2
    atTemp(2).szTicketTypeName = "��Ʊ"
    atTemp(2).nTicketTypeValid = 1
    atTemp(2).szAnnotation = ""
    
    atTemp(3).nTicketTypeID = 3
    atTemp(3).szTicketTypeName = "��Ʊ"
    atTemp(3).nTicketTypeValid = 0
    atTemp(3).szAnnotation = ""
    
    
    atTemp(4).nTicketTypeID = 4
    atTemp(4).szTicketTypeName = "�Ż�Ʊ1"
    atTemp(4).nTicketTypeValid = 0
    atTemp(4).szAnnotation = ""
    
    
    atTemp(5).nTicketTypeID = 5
    atTemp(5).szTicketTypeName = "�Ż�Ʊ2"
    atTemp(5).nTicketTypeValid = 0
    atTemp(5).szAnnotation = ""
    
    
    atTemp(6).nTicketTypeID = 6
    atTemp(6).szTicketTypeName = "�Ż�Ʊ3"
    atTemp(6).nTicketTypeValid = 0
    atTemp(6).szAnnotation = ""
    
    
    
    GetAllTicketType = atTemp
    
    
End Function

'�õ�������λ����
Public Function GetAllSeatType() As String()
    Dim aszSeatType() As String
    ReDim aszSeatType(1 To 3, 1 To 3)
    
    aszSeatType(1, 1) = "01"
    aszSeatType(1, 2) = "��ͨ"
    aszSeatType(1, 3) = ""
    
    aszSeatType(2, 1) = "02"
    aszSeatType(2, 2) = "����"
    aszSeatType(2, 3) = ""
    
    aszSeatType(3, 1) = "03"
    aszSeatType(3, 2) = "����"
    aszSeatType(3, 3) = ""
    
    GetAllSeatType = aszSeatType
    
    
End Function



 '�õ�Ʊ�ָ���
Public Function GetTicketTypeCount() As Integer
    GetTicketTypeCount = 6
End Function

'�õ���λ������Ŀ
Public Function GetSeatTypeCount() As Integer
    GetSeatTypeCount = 3
End Function

Public Function GetActiveUserID() As String
    GetActiveUserID = m_cszOperatorID
End Function

'Public Function GetAllStartStation() As String()
'    Dim aszTemp() As String
'    ReDim aszTemp(1 To 5, 1 To 2) As String
'    aszTemp(1, 1) = "0"
'    aszTemp(1, 2) = "��������"
'    aszTemp(2, 1) = "1"
'    aszTemp(2, 2) = "�³�վ"
'    aszTemp(3, 1) = "2"
'    aszTemp(3, 2) = "����վ"
'    aszTemp(4, 1) = "3"
'    aszTemp(4, 2) = "��վ"
'    aszTemp(5, 1) = "4"
'    aszTemp(5, 2) = "��վ"
'    GetAllStartStation = aszTemp
'
'End Function



