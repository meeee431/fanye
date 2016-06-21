Attribute VB_Name = "mdlBankTrans"
Option Explicit
'���Ĵ����ͨ�ñ������弰ͨ��ת��ȡ�ַ�����



'=============================================================
'���״��붨��
'=============================================================
Public Const BUYTICKETSID = "8001" '��Ʊ
Public Const GETSTATIONSID = "8002" 'ȡվ��
Public Const GETSCHEDULESID = "8003" 'ȡ����
Public Const GETSEATSID = "8004" 'ȡĳ���ε����е���λ��Ϣ
Public Const GETACCOUNTLISTID = "8005" 'ȡ���ʵ�
Public Const GETTKINFO = "8006" 'ȡ��Ʊ����ϸ��Ϣ
Public Const GETTKPRICEID = "8007" 'ȡƱ��
Public Const CANCELTICKETSID = "8011" '��Ʊ
Public Const GETCHECKGATE = "8015" 'ȡ��Ʊ��

Public Const AUTOCANCELTKID = "9999" '�Զ���Ʊ
'=============================================================
'������������
'=============================================================
Public Const LIMITTIME = 60 '����ʱ��
Public Const LIMITTICKETS = 40 '����Ʊ����
Public Const FIXPKGLEN = 194 '���ĳ���
Public Const CONNECTEDMAX = 80 '�������������

Public Const cszPackageBegin = "B"
Public Const cszPackageEnd = "@"



Public Const cszUnitName = "����"
Public Const cnTimeOut = 8000 '�Ժ���Ϊ��λ ���ʱ�䳬��8����Ʊδ�ɹ�,���Զ���Ʊ����
Public Const cszCorrectRetCode = "0000" '��ȷ����Ӧ��
Public Const cszAutoSeat = "00" '����Ϊ�Զ���ȡ��λ

Public m_szRemoteHost As String  ' "chenf"
Public m_szRemotePort As String '  "8000"
Public m_nSellStationCount '��Ʊ��վ�ĸ���

Public Const m_lTicketNoNumLen = 8 'Ʊ�ŵ����ֲ��ֵĳ���
'=============================================================
'���͵ĸ�λ����ʼλ��
'=============================================================
Public Const cnPosBegin = 1
Public Const cnPosLen = 2
Public Const cnPosTradeID = 10
Public Const cnPosOperatorID = 14
Public Const cnPosOperatorBankID = 19
Public Const cnPosTicketID = 24
Public Const cnPosTicketType = 34
Public Const cnPosStartStationID = 35
Public Const cnPosDestStationID = 44
Public Const cnPosBusID = 53
Public Const cnPosBusType = 58
Public Const cnPosTicketNum = 60
Public Const cnPosTicketPrice = 62
Public Const cnPosBusOffDate = 68
Public Const cnPosBusOffTime = 76
Public Const cnPosSeatID = 80
Public Const cnPosRetCode = 160
Public Const cnPosReserved = 164
Public Const cnPosEnd = 194

'=============================================================
'���͵ĸ�λ�ĳ���
'=============================================================
Public Const cnLenBegin = 1
Public Const cnLenLen = 8
Public Const cnLenTradeID = 4
Public Const cnLenOperatorID = 5
Public Const cnLenOperatorBankID = 5
Public Const cnLenTicketID = 10
Public Const cnLenTicketType = 1
Public Const cnLenStartStationID = 9
Public Const cnLenDestStationID = 9
Public Const cnLenBusID = 5
Public Const cnLenBusType = 2
Public Const cnLenTicketNum = 2
Public Const cnLenTicketPrice = 6
Public Const cnLenBusOffDate = 8
Public Const cnLenBusOffTime = 4
Public Const cnLenSeatID = 80
Public Const cnLenRetCode = 4
Public Const cnLenReserved = 30
Public Const cnLenEnd = 1


'=============================================================
'��վ���ѯ������Ϣ�ķ����ֶγ��ȶ���
'=============================================================
Const cnBusRsLen_bus_date = 8
Const cnBusRsLen_bus_id = 5
'    Const cnBusRsLen_check_gate_id = 2
Const cnBusRsLen_vehicle_type_name = 16
Const cnBusRsLen_vehicle_type_code = 3
Const cnBusRsLen_total_seat = 2
Const cnBusRsLen_sale_seat_quantity = 2
'    Const cnBusRsLen_total_stand_quantity = 2
'    Const cnBusRsLen_sale_stand_seat_quantity = 2
'    Const cnBusRsLen_bus_start_time = 4
Const cnBusRsLen_bus_type = 2
'    Const cnBusRsLen_register_status = 2
'    Const cnBusRsLen_is_all_refundment = 2
Const cnBusRsLen_status = 2
Const cnBusRsLen_route_id = 4
Const cnBusRsLen_end_station_id = 9
Const cnBusRsLen_end_station_name = 10
Const cnBusRsLen_owner_id = 4
Const cnBusRsLen_transport_company_id = 12
Const cnBusRsLen_vehicle_id = 5
Const cnBusRsLen_split_company_id = 12
'    Const cnBusRsLen_baggage_number = 10
'    Const cnBusRsLen_fact_weight = 10
'    Const cnBusRsLen_calculate_weight = 10
'    Const cnBusRsLen_over_weight_number = 10
'    Const cnBusRsLen_luggage_total_price = 10
'    Const cnBusRsLen_internet_status = 10
'    Const cnBusRsLen_scrollbus_check_time = 10
Const cnBusRsLen_seat_remain = 2
Const cnBusRsLen_bed_remain = 2
Const cnBusRsLen_additional_remain = 2
Const cnBusRsLen_other_remain_1 = 2
Const cnBusRsLen_other_remain_2 = 2
Const cnBusRsLen_sell_check_gate_id = 2
Const cnBusRsLen_seat_type_id = 3
Const cnBusRsLen_route_name = 16
'    Const cnBusRsLen_end_station_mileage = 4
Const cnBusRsLen_sell_station_id = 9
Const cnBusRsLen_sell_station_name = 10
Const cnBusRsLen_busstarttime = 4
Const cnBusRsLen_full_price = cnLenTicketPrice
Const cnBusRsLen_half_price = cnLenTicketPrice
Const cnBusRsLen_preferential_ticket1 = cnLenTicketPrice
Const cnBusRsLen_preferential_ticket2 = cnLenTicketPrice
Const cnBusRsLen_preferential_ticket3 = cnLenTicketPrice
Const cnBusRsLen_sale_ticket_quantity = 2
Const cnBusRsLen_stop_sale_time = 4 '2.50
Const cnBusRsLen_book_count = 2
    

'=============================================================
'��վ���ѯ������Ϣ���صĸ��ֶε���ʼλ�õĶ���
'=============================================================

Const cnBusRsPos_bus_date = 1
Const cnBusRsPos_bus_id = 9
'    Const cnBusRsPos_check_gate_id = 2
Const cnBusRsPos_vehicle_type_name = 14
Const cnBusRsPos_vehicle_type_code = 30
Const cnBusRsPos_total_seat = 33
Const cnBusRsPos_sale_seat_quantity = 35
'    Const cnBusRsPos_total_stand_quantity = 2
'    Const cnBusRsPos_sale_stand_seat_quantity = 2
'    Const cnBusRsPos_bus_start_time = 4
Const cnBusRsPos_bus_type = 37
'    Const cnBusRsPos_register_status = 2
'    Const cnBusRsPos_is_all_refundment = 2
Const cnBusRsPos_status = 39
Const cnBusRsPos_route_id = 41
Const cnBusRsPos_end_station_id = 45
Const cnBusRsPos_end_station_name = 54
Const cnBusRsPos_owner_id = 64
Const cnBusRsPos_transport_company_id = 68
Const cnBusRsPos_vehicle_id = 80
Const cnBusRsPos_split_company_id = 85
'    Const cnBusRsPos_baggage_number = 10
'    Const cnBusRsPos_fact_weight = 10
'    Const cnBusRsPos_calculate_weight = 10
'    Const cnBusRsPos_over_weight_number = 10
'    Const cnBusRsPos_luggage_total_price = 10
'    Const cnBusRsPos_internet_status = 10
'    Const cnBusRsPos_scrollbus_check_time = 10
Const cnBusRsPos_seat_remain = 97
Const cnBusRsPos_bed_remain = 99
Const cnBusRsPos_additional_remain = 101
Const cnBusRsPos_other_remain_1 = 103
Const cnBusRsPos_other_remain_2 = 105
Const cnBusRsPos_sell_check_gate_id = 107
Const cnBusRsPos_seat_type_id = 109
Const cnBusRsPos_route_name = 112
'    Const cnBusRsPos_end_station_mileage = 4
Const cnBusRsPos_sell_station_id = 128
Const cnBusRsPos_sell_station_name = 137
Const cnBusRsPos_busstarttime = 147
Const cnBusRsPos_full_price = 151
Const cnBusRsPos_half_price = 157
Const cnBusRsPos_preferential_ticket1 = 163
Const cnBusRsPos_preferential_ticket2 = 169
Const cnBusRsPos_preferential_ticket3 = 175
Const cnBusRsPos_sale_ticket_quantity = 181
Const cnBusRsPos_stop_sale_time = 183 '2.50
Const cnBusRsPos_book_count = 187



'=============================================================
'��ѯվ����Ϣ�ķ����ֶγ��ȶ���
'=============================================================
Public Const cnStationLenID = 9 'վ�����ĳ���
Public Const cnStationLenInputID = 6 'վ��������ĳ���
Public Const cnStationLenName = 10 'վ�����Ƶĳ���

Public Const cnStationPosID = 1 'վ������λ��
Public Const cnStationPosName = 10 'վ�����Ƶ�λ��
Public Const cnStationPosInputID = 16 'վ���������λ��


'=============================================================
'��ѯ��λ��Ϣ�ķ��ظ��ֶ���ʼλ�ö���
'=============================================================
Const cnSeatRsPos_SeatNo = 1
Const cnSeatRsPos_Status = 3
Const cnSeatRsPos_SticketNO = 5

Const cnSeatRsLen_SeatNo = 2
Const cnSeatRsLen_Status = 2
Const cnSeatRsLen_TicketNO = 10

'=============================================================
'��ѯԤ����Ϣ�ķ��ظ��ֶ���ʼλ�ö���
'=============================================================
Const cnBookRsPos_SeatNo = 1
Const cnBookRsPos_Telephone = 3


Const cnBookRsLen_SeatNo = 2
Const cnBookRsLen_Telephone = 20

    
'=============================================================
'���͵ĸ�λ���ַ����Ķ���
'=============================================================
Private cszBegin As String  '1-1   �̶�ΪB[1]
Private cszLen As String '2-9    �̶�Ϊ  ����ʱ00000194[8]
Private cszTradeID As String '10-13  ���״���[4]

'��������������Ϊ��ȫ�ֱ���
'Ϊ��д���򷽱�,��д��һ��
Public m_cszOperatorID As String '14-18  ���в���Ա[5]
Public m_cszOperatorBankID As String '19-23  ���з�����[5]
Public m_cszIsAmin As Integer '�û��ǳ�������ͨ���ǹ���Ա
Public m_cszOperatorBankName As String ' ���з�������


Private cszTicketID As String '24-33  Ʊ��[10]
Private cszTicketType As String '34-34  Ʊ���� 0 =ȫƱ 1 =��Ʊ[1]
Private cszStartStationID As String '35-43   ��ʼ����վ[9]
Private cszDestStationID As String '44-52  Ŀ������վ[9]
Private cszBusID As String '53-57 ���δ���[5]
Private cszBusType As String '58-59  ����[2]
Private cszTicketNum As String '60-61  ��Ʊ����[2]
Private cszTicketPrice As String '62-67  Ʊ��[8]
Private cszBusOffDate As String '68-75  ���η�������YYYYMMDD[8]
Private cszBusOffTime As String '76-79  ���η���ʱ��HHMM[4]
Private cszSeatID As String '80-159 ��λ��[80]
Private cszRetCode As String '160-163 ��Ӧ��[4]
Private cszReserved As String '164-193[30]
Private cszEnd As String '194-194 �̶�ΪE[1]






'��ʼ���̶�ֵ
Public Sub InitValue()
    
    cszBegin = cszPackageBegin
    cszLen = Format(FIXPKGLEN, String(cnLenLen, "0"))
    If Trim(m_cszOperatorID) = "" Then m_cszOperatorID = Space(cnLenOperatorID)
    If Trim(m_cszOperatorBankID) = "" Then m_cszOperatorBankID = Space(cnLenOperatorBankID)
    cszReserved = Space(cnLenReserved)
    cszEnd = cszPackageEnd
    
    '������ֵ����Ϊ�մ�
    cszTradeID = FormatLen("", cnLenTradeID)
    cszTicketID = FormatLen("", cnLenTicketID)
    cszTicketType = FormatLen("", cnLenTicketType)
    cszStartStationID = FormatLen("", cnLenStartStationID)
    cszDestStationID = FormatLen("", cnLenDestStationID)
    cszBusID = FormatLen("", cnLenBusID)
    cszBusType = FormatLen("", cnLenBusType)
    cszTicketNum = FormatLen("", cnLenTicketNum)
    cszTicketPrice = FormatLen("", cnLenTicketPrice)
    cszBusOffDate = FormatLen("", cnLenBusOffDate)
    cszBusOffTime = FormatLen("", cnLenBusOffTime)
    cszSeatID = FormatLen("", cnLenSeatID)
    cszRetCode = FormatLen("", cnLenRetCode)
    
    
End Sub

Private Sub SetStationValue(pszStartStationID As String)

    cszTradeID = GETSTATIONSID
    cszTicketID = FormatLen("", cnLenTicketID)
    cszTicketType = FormatLen("", cnLenTicketType)
    cszStartStationID = FormatLen(pszStartStationID, cnLenStartStationID)
    cszDestStationID = FormatLen("", cnLenDestStationID)
    cszBusID = FormatLen("", cnLenBusID)
    cszBusType = FormatLen("", cnLenBusType)
    cszTicketNum = FormatLen("", cnLenTicketNum)
    cszTicketPrice = FormatLen("", cnLenTicketPrice)
    cszBusOffDate = FormatLen("", cnLenBusOffDate)
    cszBusOffTime = FormatLen("", cnLenBusOffTime)
    cszSeatID = FormatLen("", cnLenSeatID)
    cszRetCode = FormatLen("", cnLenRetCode)
End Sub

Private Sub SetCheckGateValue(pszStartStationID As String)

    cszTradeID = GETCHECKGATE
    cszTicketID = FormatLen("", cnLenTicketID)
    cszTicketType = FormatLen("", cnLenTicketType)
    cszStartStationID = FormatLen(pszStartStationID, cnLenStartStationID)
    cszDestStationID = FormatLen("", cnLenDestStationID)
    cszBusID = FormatLen("", cnLenBusID)
    cszBusType = FormatLen("", cnLenBusType)
    cszTicketNum = FormatLen("", cnLenTicketNum)
    cszTicketPrice = FormatLen("", cnLenTicketPrice)
    cszBusOffDate = FormatLen("", cnLenBusOffDate)
    cszBusOffTime = FormatLen("", cnLenBusOffTime)
    cszSeatID = FormatLen("", cnLenSeatID)
    cszRetCode = FormatLen("", cnLenRetCode)
End Sub

Private Sub SetSeatValue(pszStartStationID As String, pdyBusDate As Date, pszBusID As String)
    InitValue
    cszTradeID = GETSEATSID
    cszBusOffDate = FormatLen(DateToPackage(pdyBusDate), cnLenBusOffDate)
    cszBusID = FormatLen(pszBusID, cnLenBusID)
    cszStartStationID = FormatLen(pszStartStationID, cnLenStartStationID)
End Sub


Private Sub SetBusValue(pszStartStationID As String)

    cszTradeID = GETSCHEDULESID
    cszTicketID = FormatLen("", cnLenTicketID)
    cszTicketType = FormatLen("", cnLenTicketType)
    cszStartStationID = FormatLen(pszStartStationID, cnLenStartStationID)
    cszDestStationID = FormatLen("", cnLenDestStationID)
    cszBusID = FormatLen("", cnLenBusID)
    cszBusType = FormatLen("", cnLenBusType)
    cszTicketNum = FormatLen("", cnLenTicketNum)
    cszTicketPrice = FormatLen("", cnLenTicketPrice)
    cszBusOffDate = FormatLen("", cnLenBusOffDate)
    cszBusOffTime = FormatLen("", cnLenBusOffTime)
    cszSeatID = FormatLen("", cnLenSeatID)
    cszRetCode = FormatLen("", cnLenRetCode)
End Sub



Private Sub SetSellTicketValue(pszTicketID As String, pnTicketType As Integer, pszStartStationID As String, pszEndStationID As String, pszBusID As String _
    , pnTicketNum As Integer, pdyBusDate As Date, pszSeatID As String, Optional pszOffTime As String)
    
    
    cszTradeID = BUYTICKETSID
    cszTicketID = FormatLen(pszTicketID, cnLenTicketID)
    cszTicketType = FormatLen(pnTicketType, cnLenTicketType)
    cszStartStationID = FormatLen(pszStartStationID, cnLenStartStationID)
    cszDestStationID = FormatLen(pszEndStationID, cnLenDestStationID)
    cszBusID = FormatLen(pszBusID, cnLenBusID)
    cszTicketNum = FormatLen(pnTicketNum, cnLenTicketNum)
    cszBusOffDate = FormatLen(DateToPackage(pdyBusDate), cnLenBusOffDate)
    cszBusType = FormatLen("", cnLenBusType)
    cszTicketPrice = FormatLen("", cnLenTicketPrice)
    cszBusOffTime = FormatLen(pszOffTime, cnLenBusOffTime + 1)
    cszSeatID = FormatLen(pszSeatID, cnLenSeatID)
    cszRetCode = FormatLen("", cnLenRetCode)
End Sub

Private Sub SetQueryTicketInfoValue(pszTicketID As String, pnTicketType As Integer, pszStartStationID As String, pszEndStationID As String, pszBusID As String _
    , pdyBusDate As Date, pszStartStationName As String, pszEndStationName As String, pdtBusStartTime As Date, pszUserID As String, pszSeatID As String, psgTicketPrice As Double, pnTicketStatus As Integer)
    
    
    InitValue
    cszTradeID = GETTKINFO
    cszTicketID = FormatLen(pszTicketID, cnLenTicketID)
    cszTicketType = FormatLen(pnTicketType, cnLenTicketType)
    cszStartStationID = FormatLen(pszStartStationID, cnLenStartStationID)
    cszDestStationID = FormatLen(pszEndStationID, cnLenDestStationID)
    cszBusID = FormatLen(pszBusID, cnLenBusID)
    cszTicketNum = FormatLen(1, cnLenTicketNum)
    cszBusOffDate = FormatLen(DateToPackage(pdyBusDate), cnLenBusOffDate)
    cszBusType = FormatLen(0, cnLenBusType)
    cszTicketPrice = FormatLen(MoneyToPackage(psgTicketPrice), cnLenTicketPrice)
    cszBusOffTime = FormatLen(TimeToPackage(pdtBusStartTime), cnLenBusOffTime)
    cszSeatID = FormatLen(pszSeatID, cnLenSeatID)
    cszRetCode = FormatLen(cszCorrectRetCode, cnLenRetCode)
'    m_cszOperatorID = FormatLen(pszUserID, cnLenOperatorID)
    '�����վ��\�յ�վ��\Ʊ��״̬�ϲ���ŵ�Ԥ����
    cszReserved = FormatLen(FormatLen(pszStartStationName, 10) & FormatLen(pszEndStationName, 10) & FormatLen(pnTicketStatus, 2), cnLenReserved)
    
    
End Sub

Private Sub SetCancelTicketValue(pszStartStationID As String, pszTicketID As String, pdbTicketPrice As Double)

    cszTradeID = CANCELTICKETSID
    cszTicketID = FormatLen(pszTicketID, cnLenTicketID)
    cszTicketType = FormatLen("", cnLenTicketType)
    cszStartStationID = FormatLen(pszStartStationID, cnLenStartStationID)
    cszDestStationID = FormatLen("", cnLenDestStationID)
    cszBusID = FormatLen("", cnLenBusID)
    cszBusType = FormatLen("", cnLenBusType)
    cszTicketNum = FormatLen(1, cnLenTicketNum)
    cszTicketPrice = FormatLen(MoneyToPackage(pdbTicketPrice), cnLenTicketPrice)
    cszBusOffDate = FormatLen("", cnLenBusOffDate)
    cszBusOffTime = FormatLen("", cnLenBusOffTime)
    cszSeatID = FormatLen("", cnLenSeatID)
    cszRetCode = FormatLen("", cnLenRetCode)
End Sub


Public Function GetStationRequestStr(pszStartStationID As String) As String
    SetStationValue pszStartStationID
    GetStationRequestStr = cszBegin & cszLen & cszTradeID & m_cszOperatorID & m_cszOperatorBankID & cszTicketID & cszTicketType & cszStartStationID & cszDestStationID & cszBusID _
        & cszBusType & cszTicketNum & cszTicketPrice & cszBusOffDate & cszBusOffTime & cszSeatID & cszRetCode & cszReserved & cszEnd
End Function

Public Function GetCheckGateRequestStr(pszStartStationID As String) As String
    SetCheckGateValue pszStartStationID
    GetCheckGateRequestStr = cszBegin & cszLen & cszTradeID & m_cszOperatorID & m_cszOperatorBankID & cszTicketID & cszTicketType & cszStartStationID & cszDestStationID & cszBusID _
        & cszBusType & cszTicketNum & cszTicketPrice & cszBusOffDate & cszBusOffTime & cszSeatID & cszRetCode & cszReserved & cszEnd
End Function

Public Function GetCancelTicketRequestStr(pszStartStationID As String, pszTicketID As String, pdbTicketPrice As Double) As String
    SetCancelTicketValue pszStartStationID, pszTicketID, pdbTicketPrice
    GetCancelTicketRequestStr = cszBegin & cszLen & cszTradeID & m_cszOperatorID & m_cszOperatorBankID & cszTicketID & cszTicketType & cszStartStationID & cszDestStationID & cszBusID _
        & cszBusType & cszTicketNum & cszTicketPrice & cszBusOffDate & cszBusOffTime & cszSeatID & cszRetCode & cszReserved & cszEnd
End Function


Public Function GetBusRequestStr(pszStartStationID As String, pdyDate As Date, pszEndStationID As String) As String
    SetBusValue pszStartStationID
    
    cszBusOffDate = FormatLen(DateToPackage(pdyDate), cnLenBusOffDate)
    cszDestStationID = FormatLen(pszEndStationID, cnLenDestStationID)
    GetBusRequestStr = cszBegin & cszLen & cszTradeID & m_cszOperatorID & m_cszOperatorBankID & cszTicketID & cszTicketType & cszStartStationID & cszDestStationID & cszBusID _
        & cszBusType & cszTicketNum & cszTicketPrice & cszBusOffDate & cszBusOffTime & cszSeatID & cszRetCode & cszReserved & cszEnd

End Function

Public Function GetSellTicketRequestStr(pszTicketID As String, pnTicketType As Integer, pszStartStationID As String, pszEndStationID As String, pszBusID As String _
    , pnTicketNum As Integer, pdyBusDate As Date, pszSeatID As String, Optional pszOffTime As String) As String
    
    SetSellTicketValue pszTicketID, pnTicketType, pszStartStationID, pszEndStationID, pszBusID, pnTicketNum, pdyBusDate, pszSeatID, pszOffTime
    
    GetSellTicketRequestStr = cszBegin & cszLen & cszTradeID & m_cszOperatorID & m_cszOperatorBankID & cszTicketID & cszTicketType & cszStartStationID & cszDestStationID & cszBusID _
        & cszBusType & cszTicketNum & cszTicketPrice & cszBusOffDate & cszBusOffTime & cszSeatID & cszRetCode & cszReserved & cszEnd
    
End Function

Public Function GetQueryTicketInfoRequestStr(pszTicketID As String, pnTicketType As Integer, pszStartStationID As String, pszEndStationID As String, pszBusID As String _
    , pdyBusDate As Date, pszStartStationName As String, pszEndStationName As String, pdtBusStartTime As Date, pszUserID As String, pszSeatID As String, psgTicketPrice As Double, pnTicketStatus As Integer) As String
    
    
    SetQueryTicketInfoValue pszTicketID, pnTicketType, pszStartStationID, pszEndStationID, pszBusID _
        , pdyBusDate, pszStartStationName, pszEndStationName, pdtBusStartTime, pszUserID, pszSeatID, psgTicketPrice, pnTicketStatus
    
    GetQueryTicketInfoRequestStr = cszBegin & cszLen & cszTradeID & m_cszOperatorID & m_cszOperatorBankID & cszTicketID & cszTicketType & cszStartStationID & cszDestStationID & cszBusID _
        & cszBusType & cszTicketNum & cszTicketPrice & cszBusOffDate & cszBusOffTime & cszSeatID & cszRetCode & cszReserved & cszEnd
    
End Function


Public Function GetSeatRequestStr(pszStartStationID As String, pdyBusDate As Date, pszBusID As String) As String
    SetSeatValue pszStartStationID, pdyBusDate, pszBusID
    
    GetSeatRequestStr = cszBegin & cszLen & cszTradeID & m_cszOperatorID & m_cszOperatorBankID & cszTicketID & cszTicketType & cszStartStationID & cszDestStationID & cszBusID _
        & cszBusType & cszTicketNum & cszTicketPrice & cszBusOffDate & cszBusOffTime & cszSeatID & cszRetCode & cszReserved & cszEnd
End Function


Public Function GetTradeID(ByVal pszStr As String) As String
    '�õ������ַ����Ľ��״���
    
    GetTradeID = MidA(pszStr, cnPosTradeID, cnLenTradeID)
    
End Function

Public Function GetBegin(ByVal pszStr As String) As String
    '�õ���ʼ��
    GetBegin = MidA(pszStr, cnPosBegin, cnLenBegin)
    
End Function

Public Function GetLen(ByVal pszStr As String) As Integer
    '�õ������ַ����ĳ���
    GetLen = Val(MidA(pszStr, cnPosLen, cnLenLen))
    
End Function

Public Function GetOperatorID(ByVal pszStr As String) As String
    '�õ��������в���Ա
    GetOperatorID = Trim(MidA(pszStr, cnPosOperatorID, cnLenOperatorID))
End Function

Public Function GetOperatorBankID(ByVal pszStr As String) As String
    '�õ��������з�����
    GetOperatorBankID = Trim(MidA(pszStr, cnPosOperatorBankID, cnLenOperatorBankID))
End Function

Public Function GetTicketID(ByVal pszStr As String) As String
    'Ʊ��
    GetTicketID = Trim(MidA(pszStr, cnPosTicketID, cnLenTicketID))
End Function

Public Function GetTicketType(ByVal pszStr As String) As String
    'Ʊ��
    GetTicketType = Val(MidA(pszStr, cnPosTicketType, cnLenTicketType))
End Function

Public Function GetStartStationID(ByVal pszStr As String) As String
    '���վ
    GetStartStationID = Trim(MidA(pszStr, cnPosStartStationID, cnLenStartStationID))
End Function

Public Function GetDestStationID(ByVal pszStr As String) As String
    '�յ�վ
    GetDestStationID = Trim(MidA(pszStr, cnPosDestStationID, cnLenDestStationID))
End Function

Public Function GetPackageBusID(ByVal pszStr As String) As String
    '���δ���
    GetPackageBusID = Trim(MidA(pszStr, cnPosBusID, cnLenBusID))
End Function

Public Function GetBusType(ByVal pszStr As String) As String
    '����
    GetBusType = Trim(MidA(pszStr, cnPosBusType, cnLenBusType))
End Function

Public Function GetTicketNum(ByVal pszStr As String) As String
    '��Ʊ����
    GetTicketNum = Val(MidA(pszStr, cnPosTicketNum, cnLenTicketNum))
End Function

Public Function GetTicketPrice(ByVal pszStr As String) As String
    'Ʊ��
    GetTicketPrice = Val(MidA(pszStr, cnPosTicketPrice, cnLenTicketPrice))
End Function

Public Function GetBusOffDate(ByVal pszStr As String) As String
    '���η�������YYYYMMDD
    GetBusOffDate = MidA(pszStr, cnPosBusOffDate, cnLenBusOffDate)
End Function

Public Function GetBusOffTime(ByVal pszStr As String) As String
    '���η���ʱ��HHMM
    GetBusOffTime = MidA(pszStr, cnPosBusOffTime, cnLenBusOffTime)
End Function

Public Function GetBusID2(ByVal pszStr As String) As String
    '���δ���
    GetBusID2 = MidA(pszStr, cnPosBusID, cnLenBusID)
End Function

Public Function GetSeatID(ByVal pszStr As String) As String
    '��λ
    GetSeatID = Trim(MidA(pszStr, cnPosSeatID, cnLenSeatID))
End Function

Public Function GetRetCode(ByVal pszStr As String) As String
    '��Ӧ��
    GetRetCode = Trim(MidA(pszStr, cnPosRetCode, cnLenRetCode))
End Function

Public Function GetReserved(pszStr As String) As String
    'Ԥ��
    GetReserved = Trim(MidA(pszStr, cnPosReserved, cnLenReserved))
End Function

Public Function GetEnd(ByVal pszStr As String) As String
    '�õ���ֹ��
    GetEnd = MidA(pszStr, cnPosEnd, cnLenEnd)
    
End Function

'
'Public Function GetAllStartStationID() As String()
'    '������վ���վ�����
'    Dim aszTemp() As String
'    If m_nSellStationCount = 0 Then
'        m_nSellStationCount = 5
'    End If
'    ReDim aszTemp(0 To m_nSellStationCount - 1)
'    aszTemp(0) = "wzkyz"
'    aszTemp(1) = "wzxcz"
'    aszTemp(2) = "wzxnz"
'    aszTemp(3) = "wzxz"
'    aszTemp(4) = "wzdz"
'    GetAllStartStationID = aszTemp
'
'End Function
'
'Public Function GetAllSellStationID() As String()
'    '�õ����е��ϳ�վ
'    Dim aszTemp() As String
'    If m_nSellStationCount = 0 Then
'        m_nSellStationCount = 5
'    End If
'    ReDim aszTemp(0 To m_nSellStationCount - 1)
'    aszTemp(0) = "000003"
'    aszTemp(1) = "xcz"
'    aszTemp(2) = "720577002"
'    aszTemp(3) = "xz"
'    aszTemp(4) = "dz"
'    GetAllSellStationID = aszTemp
'
'End Function

'��վ���ѯ���÷��صļ�¼��ת�����ַ���
Public Function ConvertStationRSToString(ByVal prsStationRS As ADODB.Recordset) As String
    '��ConvertStringToStationRS���������
    Dim szStr As String
    Dim i As Integer
    szStr = ""
    With prsStationRS
'        .MoveFirst
        For i = 1 To .RecordCount
            szStr = szStr & FormatLen(!station_id, cnStationLenID)
            szStr = szStr & FormatLen(!station_name, cnStationLenName)
            szStr = szStr & FormatLen(!station_input_code, cnStationLenInputID)
            szStr = szStr & "|"
            .MoveNext
        Next i
    End With
    ConvertStationRSToString = szStr
End Function
    


'��վ���ѯ���÷��ص��ַ���ת���ɼ�¼��
Public Function ConvertStringToStationRS(ByVal pszStr As String) As ADODB.Recordset
    '��ConvertStationRSToString���������
    Dim szTemp As String
    Dim nLen As Integer '��Ŵ���ǰ׺�ĳ���
    Dim aszStation() As String
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    '������ı��ĵ�ǰ�沿��ȥ��
    szTemp = pszStr
    
    
'    Debug.Print sztemp
    
    
'    nLen = GetLen(szTemp)
'    szTemp = Right(szTemp, LenA(szTemp) - nLen)
    aszStation = Split(szTemp, "|")
    '������¼��
    With rsTemp
        .Fields.Append "station_id", adChar, 9
        .Fields.Append "station_name", adChar, 10
        .Fields.Append "station_input_code", adChar, 6
    End With
    rsTemp.Open
    With rsTemp
        '��Ҫע��һ�����һ��"|"�����
        For i = 1 To ArrayLength(aszStation) - 1 - 1 '��1 ��ʼ����Ϊ��һ��"|"ǰ������������      ��Ϊ�����һ��"|",�ʻ���Ҫ��1
            .AddNew
            .Fields("station_id") = MidA(aszStation(i), cnStationPosID, cnStationLenID)
            .Fields("station_name") = MidA(aszStation(i), cnStationPosName, cnStationLenName)
            .Fields("station_input_code") = MidA(aszStation(i), cnStationPosInputID, cnStationLenInputID)
            .Update
        Next i
    End With
    Set ConvertStringToStationRS = rsTemp
    
    
End Function

    
'��վ���ѯ���÷��ص��ַ���ת���ɼ�¼��
Public Function ConvertStringToCheckGateRS(ByVal pszStr As String) As ADODB.Recordset
    '��ConvertStationRSToString���������
    Dim szTemp As String
    Dim nLen As Integer '��Ŵ���ǰ׺�ĳ���
    Dim aszStation() As String
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    '������ı��ĵ�ǰ�沿��ȥ��
    szTemp = pszStr
    
    
'    Debug.Print sztemp
    
    
'    nLen = GetLen(szTemp)
'    szTemp = Right(szTemp, LenA(szTemp) - nLen)
    aszStation = Split(szTemp, "|")
    '������¼��
    With rsTemp
        .Fields.Append "check_gate_id", adChar, 9
        .Fields.Append "check_gate_name", adChar, 10
    End With
    rsTemp.Open
    With rsTemp
        '��Ҫע��һ�����һ��"|"�����
        For i = 1 To ArrayLength(aszStation) - 1 - 1 '��1 ��ʼ����Ϊ��һ��"|"ǰ������������      ��Ϊ�����һ��"|",�ʻ���Ҫ��1
            .AddNew
            .Fields("check_gate_id") = MidA(aszStation(i), cnStationPosID, cnStationLenID)
            .Fields("check_gate_name") = MidA(aszStation(i), cnStationPosName, cnStationLenName)
            .Update
        Next i
    End With
    Set ConvertStringToCheckGateRS = rsTemp
    
    
End Function

'��������λ��ѯ���÷��ص��ַ���ת���ɼ�¼��
Public Function ConvertStringToSeatRS(ByVal pszStr As String) As ADODB.Recordset
    Dim szTemp As String
    Dim nLen As Integer '��Ŵ���ǰ׺�ĳ���
    Dim aszTemp() As String
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    '������ı��ĵ�ǰ�沿��ȥ��
    szTemp = pszStr
    
'    nLen = GetLen(szTemp)
'    szTemp = Right(szTemp, LenA(szTemp) - nLen)
    aszTemp = Split(szTemp, "|")
    '������¼��
    With rsTemp
        .Fields.Append "seat_no", adChar, cnSeatRsLen_SeatNo
        .Fields.Append "status", adInteger
        .Fields.Append "ticket_no", adChar, cnSeatRsLen_TicketNO
    End With
    
    rsTemp.Open
    With rsTemp
        '��Ҫע��һ�����һ��"|"�����
        For i = 1 To ArrayLength(aszTemp) - 1 - 1 '��1 ��ʼ����Ϊ��һ��"|"ǰ������������      ��Ϊ�����һ��"|",�ʻ���Ҫ��1
            .AddNew
            .Fields("seat_no") = MidA(aszTemp(i), cnSeatRsPos_SeatNo, cnSeatRsLen_SeatNo)
            .Fields("status") = Val(MidA(aszTemp(i), cnSeatRsPos_Status, cnSeatRsLen_Status))
            .Fields("ticket_no") = MidA(aszTemp(i), cnSeatRsPos_SticketNO, cnSeatRsLen_TicketNO)
            .Update
        Next i
    End With
    Set ConvertStringToSeatRS = rsTemp
    
    
End Function


'����λ��ѯ���÷��صļ�¼��ת�����ַ���
Public Function ConvertSeatRSToString(ByVal prsInfo As ADODB.Recordset) As String
    '��ConvertStringToSeatRS���������
    Dim szStr As String
    Dim i As Integer
    szStr = ""
    With prsInfo
        For i = 1 To .RecordCount
            szStr = szStr & FormatLen(!seat_no, cnSeatRsLen_SeatNo)
            szStr = szStr & FormatLen(!status, cnSeatRsLen_Status)
            szStr = szStr & FormatLen(!ticket_no, cnSeatRsLen_TicketNO)
            szStr = szStr & "|"
            .MoveNext
        Next i
    End With
    ConvertSeatRSToString = szStr
End Function





'�����β�ѯ���÷��صļ�¼��ת�����ַ���
Public Function ConvertBusRSToString(ByVal prsBusRS As ADODB.Recordset) As String
    '��ConvertStringToBusRS���������
    Dim szStr As String
    Dim i As Integer
    szStr = ""
    
    With prsBusRS
'        .MoveFirst
        For i = 1 To .RecordCount
            '���ֶι̶����������,ÿһ����¼��"|"��β
            szStr = szStr & FormatLen(DateToPackage(!bus_date), cnBusRsLen_bus_date)
            szStr = szStr & FormatLen(!bus_id, cnBusRsLen_bus_id)
        '    szStr = szStr & FormatLen(!check_gate_id, cnBusRsLen_check_gate_id)
            szStr = szStr & FormatLen(!vehicle_type_name, cnBusRsLen_vehicle_type_name)
            szStr = szStr & FormatLen(!vehicle_type_code, cnBusRsLen_vehicle_type_code)
            szStr = szStr & FormatLen(!total_seat, cnBusRsLen_total_seat)
            szStr = szStr & FormatLen(!sale_seat_quantity, cnBusRsLen_sale_seat_quantity)
        '    szStr = szStr & FormatLen(!total_stand_quantity, cnBusRsLen_total_stand_quantity)
        '    szStr = szStr & FormatLen(!sale_stand_seat_quantity, cnBusRsLen_sale_stand_seat_quantity)
    '        szStr = szStr & FormatLen(TimeToPackage(!bus_start_time), cnBusRsLen_bus_start_time)
            szStr = szStr & FormatLen(!bus_type, cnBusRsLen_bus_type)
        '    szStr = szStr & FormatLen(!register_status, cnBusRsLen_register_status)
        '    szStr = szStr & FormatLen(!is_all_refundment, cnBusRsLen_is_all_refundment)
            szStr = szStr & FormatLen(!status, cnBusRsLen_status)
            szStr = szStr & FormatLen(!route_id, cnBusRsLen_route_id)
            szStr = szStr & FormatLen(!end_station_id, cnBusRsLen_end_station_id)
            szStr = szStr & FormatLen(!end_station_name, cnBusRsLen_end_station_name)
            szStr = szStr & FormatLen(!owner_id, cnBusRsLen_owner_id)
            szStr = szStr & FormatLen(!transport_company_id, cnBusRsLen_transport_company_id)
            szStr = szStr & FormatLen(!vehicle_id, cnBusRsLen_vehicle_id)
            szStr = szStr & FormatLen(!split_company_id, cnBusRsLen_split_company_id)
        '    szStr = szStr & FormatLen(!baggage_number, cnBusRsLen_baggage_number)
        '    szStr = szStr & FormatLen(!fact_weight, cnBusRsLen_fact_weight)
        '    szStr = szStr & FormatLen(!calculate_weight, cnBusRsLen_calculate_weight)
        '    szStr = szStr & FormatLen(!over_weight_number, cnBusRsLen_over_weight_number)
        '    szStr = szStr & FormatLen(!luggage_total_price, cnBusRsLen_luggage_total_price)
        '    szStr = szStr & FormatLen(!internet_status, cnBusRsLen_internet_status)
        '    szStr = szStr & FormatLen(!scrollbus_check_time, cnBusRsLen_scrollbus_check_time)
            szStr = szStr & FormatLen(!seat_remain, cnBusRsLen_seat_remain)
            szStr = szStr & FormatLen(!bed_remain, cnBusRsLen_bed_remain)
            szStr = szStr & FormatLen(!additional_remain, cnBusRsLen_additional_remain)
            szStr = szStr & FormatLen(!other_remain_1, cnBusRsLen_other_remain_1)
            szStr = szStr & FormatLen(!other_remain_2, cnBusRsLen_other_remain_2)
            szStr = szStr & FormatLen(!sell_check_gate_id, cnBusRsLen_sell_check_gate_id)
            szStr = szStr & FormatLen(!seat_type_id, cnBusRsLen_seat_type_id)
            szStr = szStr & FormatLen(!route_name, cnBusRsLen_route_name)
        '    szStr = szStr & FormatLen(!end_station_mileage, cnBusRsLen_end_station_mileage)
            szStr = szStr & FormatLen(!sell_station_id, cnBusRsLen_sell_station_id)
            szStr = szStr & FormatLen(!sell_station_name, cnBusRsLen_sell_station_name)
            szStr = szStr & FormatLen(TimeToPackage(!busstarttime), cnBusRsLen_busstarttime)
            szStr = szStr & FormatLen(MoneyToPackage(!full_price), cnBusRsLen_full_price)
            szStr = szStr & FormatLen(MoneyToPackage(!half_price), cnBusRsLen_half_price)
            szStr = szStr & FormatLen(MoneyToPackage(!preferential_ticket1), cnBusRsLen_preferential_ticket1)
            szStr = szStr & FormatLen(MoneyToPackage(!preferential_ticket2), cnBusRsLen_preferential_ticket2)
            szStr = szStr & FormatLen(MoneyToPackage(!preferential_ticket3), cnBusRsLen_preferential_ticket3)
            szStr = szStr & FormatLen(!sale_ticket_quantity, cnBusRsLen_sale_ticket_quantity)
            szStr = szStr & FormatLen(!stop_sale_time, cnBusRsLen_stop_sale_time)
            szStr = szStr & FormatLen(!book_count, cnBusRsLen_book_count)
            szStr = szStr & "|"
            
            
            .MoveNext
        Next i
    End With
    ConvertBusRSToString = szStr
End Function

'�����β�ѯ���÷��ص��ַ���ת���ɼ�¼��
Public Function ConvertStringToBusRS(ByVal pszStr As String) As ADODB.Recordset
    Dim szTemp As String
    Dim nLen As Integer '��Ŵ���ǰ׺�ĳ���
    Dim aszBus() As String
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    '������ı��ĵ�ǰ�沿��ȥ��
    szTemp = pszStr
'    nLen = GetLen(szTemp)
'    szTemp = Right(szTemp, LenA(szTemp) - nLen)
    aszBus = Split(szTemp, "|")
    Dim nBusLen As Integer
    nBusLen = ArrayLength(aszBus)
    If nBusLen <= 2 Then Exit Function
    
    '������¼��
    With rsTemp
        .Fields.Append "bus_date", adDate
        .Fields.Append "bus_id", adChar, cnBusRsLen_bus_id
        .Fields.Append "vehicle_type_name", adChar, cnBusRsLen_vehicle_type_name
        .Fields.Append "vehicle_type_code", adChar, cnBusRsLen_vehicle_type_code
        .Fields.Append "total_seat", adInteger
        .Fields.Append "sale_seat_quantity", adInteger
        .Fields.Append "bus_type", adInteger
        
        
        .Fields.Append "Status", adInteger
        .Fields.Append "route_id", adChar, cnBusRsLen_route_id
        .Fields.Append "end_station_id", adChar, cnBusRsLen_end_station_id
        .Fields.Append "end_station_name", adChar, cnBusRsLen_end_station_name
        .Fields.Append "owner_id", adChar, cnBusRsLen_owner_id
        .Fields.Append "transport_company_id", adChar, cnBusRsLen_transport_company_id
        
        
        .Fields.Append "vehicle_id", adChar, cnBusRsLen_vehicle_id
        .Fields.Append "split_company_id", adChar, cnBusRsLen_split_company_id
        .Fields.Append "seat_remain", adInteger
        .Fields.Append "bed_remain", adInteger
        .Fields.Append "additional_remain", adInteger
        .Fields.Append "other_remain_1", adInteger
        .Fields.Append "other_remain_2", adInteger
        .Fields.Append "sell_check_gate_id", adChar, cnBusRsLen_sell_check_gate_id
        
        
        
        .Fields.Append "seat_type_id", adChar, cnBusRsLen_seat_type_id
        .Fields.Append "route_name", adChar, cnBusRsLen_route_name
        .Fields.Append "sell_station_id", adChar, cnBusRsLen_sell_station_id
        .Fields.Append "sell_station_name", adChar, cnBusRsLen_sell_station_name
        .Fields.Append "busstarttime", adDate
        .Fields.Append "full_price", adDouble
        .Fields.Append "half_price", adDouble
        .Fields.Append "preferential_ticket1", adDouble
        .Fields.Append "preferential_ticket2", adDouble
        .Fields.Append "preferential_ticket3", adDouble
        .Fields.Append "sale_ticket_quantity", adInteger
        .Fields.Append "stop_sale_time", adDouble
        .Fields.Append "book_count", adInteger
    End With
    
    rsTemp.Open
    With rsTemp
        '��Ҫע��һ�����һ��"|"�����
        For i = 1 To ArrayLength(aszBus) - 1 - 1 '��Ϊ�����һ��"|",�ʻ���Ҫ��1
            .AddNew
            .Fields("bus_date") = PackageToDate(MidA(aszBus(i), cnBusRsPos_bus_date, cnBusRsLen_bus_date))
            .Fields("bus_id") = MidA(aszBus(i), cnBusRsPos_bus_id, cnBusRsLen_bus_id)
            .Fields("vehicle_type_name") = MidA(aszBus(i), cnBusRsPos_vehicle_type_name, cnBusRsLen_vehicle_type_name)
            .Fields("vehicle_type_code") = MidA(aszBus(i), cnBusRsPos_vehicle_type_code, cnBusRsLen_vehicle_type_code)
            .Fields("total_seat") = MidA(aszBus(i), cnBusRsPos_total_seat, cnBusRsLen_total_seat)
            .Fields("sale_seat_quantity") = MidA(aszBus(i), cnBusRsPos_sale_seat_quantity, cnBusRsLen_sale_seat_quantity)
            .Fields("bus_type") = MidA(aszBus(i), cnBusRsPos_bus_type, cnBusRsLen_bus_type)
            
            
            .Fields("Status") = MidA(aszBus(i), cnBusRsPos_status, cnBusRsLen_status)
            .Fields("route_id") = MidA(aszBus(i), cnBusRsPos_route_id, cnBusRsLen_route_id)
            .Fields("end_station_id") = MidA(aszBus(i), cnBusRsPos_end_station_id, cnBusRsLen_end_station_id)
            .Fields("end_station_name") = MidA(aszBus(i), cnBusRsPos_end_station_name, cnBusRsLen_end_station_name)
            .Fields("owner_id") = MidA(aszBus(i), cnBusRsPos_owner_id, cnBusRsLen_owner_id)
            .Fields("vehicle_id") = MidA(aszBus(i), cnBusRsPos_vehicle_id, cnBusRsLen_vehicle_id)
            .Fields("split_company_id") = MidA(aszBus(i), cnBusRsPos_transport_company_id, cnBusRsLen_transport_company_id)
            
            
            .Fields("seat_remain") = MidA(aszBus(i), cnBusRsPos_seat_remain, cnBusRsLen_seat_remain)
            .Fields("bed_remain") = MidA(aszBus(i), cnBusRsPos_bed_remain, cnBusRsLen_bed_remain)
            .Fields("additional_remain") = MidA(aszBus(i), cnBusRsPos_additional_remain, cnBusRsLen_additional_remain)
            .Fields("other_remain_1") = MidA(aszBus(i), cnBusRsPos_other_remain_1, cnBusRsLen_other_remain_1)
            .Fields("other_remain_2") = MidA(aszBus(i), cnBusRsPos_other_remain_2, cnBusRsLen_other_remain_2)
            .Fields("sell_check_gate_id") = MidA(aszBus(i), cnBusRsPos_sell_check_gate_id, cnBusRsLen_sell_check_gate_id)
            
            
            .Fields("seat_type_id") = MidA(aszBus(i), cnBusRsPos_seat_type_id, cnBusRsLen_seat_type_id)
            .Fields("route_name") = MidA(aszBus(i), cnBusRsPos_route_name, cnBusRsLen_route_name)
            
            .Fields("sell_station_id") = MidA(aszBus(i), cnBusRsPos_sell_station_id, cnBusRsLen_sell_station_id)
            .Fields("sell_station_name") = MidA(aszBus(i), cnBusRsPos_sell_station_name, cnBusRsLen_sell_station_name)
            .Fields("busstarttime") = PackageToTime(MidA(aszBus(i), cnBusRsPos_busstarttime, cnBusRsLen_busstarttime))
            .Fields("full_price") = PackageToMoney(MidA(aszBus(i), cnBusRsPos_full_price, cnBusRsLen_full_price))
            .Fields("half_price") = PackageToMoney(MidA(aszBus(i), cnBusRsPos_half_price, cnBusRsLen_half_price))
            .Fields("preferential_ticket1") = PackageToMoney(MidA(aszBus(i), cnBusRsPos_preferential_ticket1, cnBusRsLen_preferential_ticket1))
            
            
            .Fields("preferential_ticket2") = PackageToMoney(MidA(aszBus(i), cnBusRsPos_preferential_ticket2, cnBusRsLen_preferential_ticket2))
            .Fields("preferential_ticket3") = PackageToMoney(MidA(aszBus(i), cnBusRsPos_preferential_ticket3, cnBusRsLen_preferential_ticket3))
            .Fields("sale_ticket_quantity") = MidA(aszBus(i), cnBusRsPos_sale_ticket_quantity, cnBusRsLen_sale_ticket_quantity)
            .Fields("stop_sale_time") = MidA(aszBus(i), cnBusRsPos_stop_sale_time, cnBusRsLen_stop_sale_time)
            .Fields("book_count") = Val(MidA(aszBus(i), cnBusRsPos_book_count, cnBusRsLen_book_count))
            
            .Update
        Next i
    End With
    Set ConvertStringToBusRS = rsTemp
End Function


Public Function FormatLen(ByVal pszStr As String, ByVal pnLen As Integer) As String
    '����ָ�����ȵ��ַ���
    Dim szTemp As String
    If pnLen > 0 Then
        If LenA(pszStr) >= pnLen Then
            FormatLen = Left(pszStr, pnLen)
        Else
            FormatLen = pszStr & Space(pnLen - LenA(pszStr))
        End If
    Else
        FormatLen = ""
    End If
    
End Function

'�õ����ݿ�󶨵����ݿؼ��������ַ���
Public Function GetAdodcConnectionStr()
    GetAdodcConnectionStr = "PROVIDER=MSDASQL;dsn=sx;uid=sa;pwd=;"
End Function

'������ת���ɴ����õİ�������
Public Function DateToPackage(ByVal pdyDate As Date) As String
    DateToPackage = Format(pdyDate, "YYYYMMDD")
End Function

'��ʱ��ת���ɴ����õİ���ʱ��
Public Function TimeToPackage(ByVal pdyTime As Date) As String
    TimeToPackage = Format(pdyTime, "hhmm")
    
End Function

'�����ת���ɴ����õİ��Ľ��
Public Function MoneyToPackage(ByVal pdbMoney As Double) As String
    '�˴����еĽ���ʱ������*100   ��ʱ���ʱ,��Ҫ����/100
    MoneyToPackage = Trim(Str(pdbMoney * 100))
    
End Function


'�������õİ��Ľ��ת����ʵ�ʽ��
Public Function PackageToMoney(ByVal pszString As String) As Double
    '�˴����еĽ���ʱ������*100   ��ʱ���ʱ,��Ҫ����/100
    PackageToMoney = Val(pszString) / 100
    
End Function



'�������õİ�������ת����ʵ������
Public Function PackageToDate(ByVal pszString As String) As Date
    Dim szTemp As String
    szTemp = Left(pszString, 4) & "-" & MidA(pszString, 5, 2) & "-" & Right(pszString, 2)
    If IsDate(szTemp) Then
        
        PackageToDate = szTemp
    End If
End Function

'�������õİ���ʱ��ת����ʵ��ʱ��
Public Function PackageToTime(ByVal pszString As String) As Date
'    TimeToPackage = Format(pdyTime, "hhmm")
    Dim szTemp As String
    szTemp = Left(pszString, 2) & ":" & Right(pszString, 2)
    If IsDate(szTemp) Then
        PackageToTime = szTemp
    End If
    
End Function

'���ֶ�ֵת����TSQL���õ��ֶ��ַ���
Public Function TransFieldValueToString(pvFieldValue As Variant) As String
    TransFieldValueToString = ""
    If Not IsNull(pvFieldValue) Then
        Select Case VarType(pvFieldValue)
            Case vbSingle, vbDouble, vbInteger, vbLong, vbCurrency, vbDecimal, vbByte
                TransFieldValueToString = pvFieldValue
            Case vbBoolean
                TransFieldValueToString = IIf(pvFieldValue, 1, 0)
            Case vbDate
                TransFieldValueToString = "'" & ToDBDateTime(CDate(pvFieldValue)) & "'"
            Case vbString
                Dim aszSplitString() As String
                Dim i As Integer
                aszSplitString = Split(pvFieldValue, "'")
                If ArrayLength(aszSplitString) > 0 Then     '�ڲ����������ź�˫���ţ��������
                    TransFieldValueToString = "'" & aszSplitString(0) & "'"
                    For i = 1 To ArrayLength(aszSplitString) - 1
                        TransFieldValueToString = TransFieldValueToString & "+" & Chr(34) & "'" & Chr(34) & "+'" & aszSplitString(i) & "'"
                    Next i
                Else
                    TransFieldValueToString = "'" & pvFieldValue & "'"  '�����Ž�������
                End If
        End Select
    End If
End Function







