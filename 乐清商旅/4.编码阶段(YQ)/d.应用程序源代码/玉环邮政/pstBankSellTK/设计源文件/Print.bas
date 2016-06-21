Attribute VB_Name = "Print"
Option Explicit


Const cnStartStation1 = 1
Const cnEndStation1 = 2
Const cnTicketType1 = 3
Const cnTicketPrice1 = 4
Const cnStartStation2 = 5
Const cnEndStation2 = 6
Const cnTicketPrice2 = 7
Const cnTicketType2 = 8
Const cnUserID = 9
Const cnBusDate1 = 10
Const cnBusID1 = 11
Const cnBusDate2 = 12
Const cnBusID2 = 13
Const cnSeat = 14
Const cnStartupTime = 15
Const cnVehicleType = 16
Const cnCheckGate = 17
Const cnTicketNO = 18
Const cnUserID2 = 19
Const cnTicketNO1 = 20
Const cnSeat2 = 21
Const cnCompany = 22
Const cnInsurance = 23

'���ش�ӡͷ
'����ԭ����22�������ļ���Ҫ��
Const cnReturnPoint = 24

Const clRTicketID = 1
Const clRTicketPrice = 2
Const clRReturnCharge = 3
Const clRSheetID = 4
Const clRStartUpTime = 5
Const clRUserID = 6
Const clRBusID = 7
Const clREndStation = 8
Const clRNowDate = 9
Const clRTicketType = 10
Const clREndStation2 = 11
Const clRTicketType2 = 12
Const clRStartUpTime2 = 13
Const clRTicketCount = 14
Const clRTicketCount2 = 15
Const clRReturnCharge2 = 16
Const clRReturnCountStr = 17
Const clRReturnChargeStr = 18
Const clRTicketID2 = 19
Const clRBusID2 = 20
Const clRVehicleType = 21
Const clRSeatNo = 22
Const clRNowDate2 = 23
Const clRStartStation1 = 24
Const clRStartStation2 = 25
Const clRReturnChargeStr2 = 26
Const clRBusDate1 = 27
Const clRBusDate2 = 28
Const clRSheetIDStr = 29
Const clRReturnTimeStr = 32
Const clRUserID1 = 30
Const clRSeatNo1 = 31

#If PRINT_SHEET <> 0 Then
    Public m_oFastReturn As New FastPrint
    Public m_oFastPrint As New FastPrint
    
    Public m_oPrintTicket As New BPrint
    Public m_oPrintReturn As New BPrint
#End If
''
''#If PRINT_SHEET <> 0 Then
'''    Public m_oPrintTicket As BPrint
''    Public m_oPrintReturn As New FastPrint
''    Public m_oFastPrint As New FastPrint
''#End If

'��ӡ��Ʊƾ֤
Public Function PrintReturnSheet(pszTicketNo As String, pszSheetID As String, psgReturnCharge As Double, psgTicketPrice As Double, pszOpTime As String, pszBusID As String, pszEndStation As String, pnReturnCount As String) As Long
 #If PRINT_SHEET <> 0 Then
On Error GoTo Error_Handle
    If m_bUseFastPrint Then     '��������ٴ�ӡ������billPrint
        PrintReturnSheetFast pszTicketNo, pszSheetID, psgReturnCharge, psgTicketPrice, pszOpTime, pszBusID, pszEndStation, pnReturnCount
    Else
        PrintReturnSheetSlow pszTicketNo, pszSheetID, psgReturnCharge, psgTicketPrice, pszOpTime, pszBusID, pszEndStation, pnReturnCount
    End If
    Exit Function
Error_Handle:
    ShowErrorMsg
#End If
End Function
Private Function PrintReturnSheetFast(pszTicketNo As String, pszSheetID As String, psgReturnCharge As Double, psgTicketPrice As Double, pszOpTime As String, pszBusID As String, pszEndStation As String, pnReturnCount As String) As Long
    m_oFastReturn.ClosePort
    m_oFastReturn.OpenPort
    m_oFastReturn.SetObject clRTicketID
    m_oFastReturn.SetCaption pszTicketNo   '��ӡ��ǰƱ��

    m_oFastReturn.SetObject clRTicketPrice  '��ӡ��ǰƱ��
    m_oFastReturn.SetCaption pszTicketNo

    m_oFastReturn.SetObject clRReturnCharge   '��Ʊ������
    m_oFastReturn.SetCaption Format(psgReturnCharge, "0.00") & "Ԫ"

    m_oFastReturn.SetObject clRSheetID
    m_oFastReturn.SetCaption "��Ʊ�վ�"

    m_oFastReturn.SetObject clRNowDate    '��ӡ����
    m_oFastReturn.SetCaption Format(Date, "YYYY-MM-DD")

    m_oFastReturn.SetObject clRUserID  '��ӡ����
    m_oFastReturn.SetCaption GetActiveUserID

    m_oFastReturn.SetObject clRNowDate2
    m_oFastReturn.SetCaption Format(Date, "YYYY-MM-DD")

    m_oFastReturn.SetObject clRTicketCount  '��ӡƱ������
    m_oFastReturn.SetCaption pnReturnCount & "��"

    m_oFastReturn.SetObject clRTicketCount2  '��ӡƱ������
    m_oFastReturn.SetCaption pnReturnCount & "��"

    m_oFastReturn.SetObject clRReturnCharge2   '��Ʊ������
    m_oFastReturn.SetCaption Format(psgReturnCharge, "0.00") & "Ԫ"

    m_oFastReturn.SetObject clRReturnCountStr  '��ӡ��Ʊ�����ַ���
    m_oFastReturn.SetCaption "��Ʊ����"

    m_oFastReturn.SetObject clRReturnChargeStr  '��ӡ�������ַ���
    m_oFastReturn.SetCaption "������"

    m_oFastReturn.SetObject clREndStation '��ӡ�յ�վ
    m_oFastReturn.SetCaption pszEndStation

    m_oFastReturn.PrintFile

    m_oFastReturn.ClosePort

End Function
Private Function PrintReturnSheetSlow(pszTicketNo As String, pszSheetID As String, psgReturnCharge As Double, psgTicketPrice As Double, pszOpTime As String, pszBusID As String, pszEndStation As String, pnReturnCount As String) As Long
    m_oPrintReturn.SetCurrentObject clRTicketID
    m_oPrintReturn.LabelSetCaption ""   '��ӡ��ǰƱ��

    m_oPrintReturn.SetCurrentObject clRStartStation1  '��ӡ���վ
    m_oPrintReturn.LabelSetCaption ""
    
    m_oPrintReturn.SetCurrentObject clREndStation '��ӡ�յ�վ
    m_oPrintReturn.LabelSetCaption ""   'pszEndStation
    
    m_oPrintReturn.SetCurrentObject clRTicketPrice  '��ӡ��ǰƱ��
    m_oPrintReturn.LabelSetCaption Format(psgTicketPrice, "0.00")

    m_oPrintReturn.SetCurrentObject clRReturnCharge   '��Ʊ������
    m_oPrintReturn.LabelSetCaption Format(psgReturnCharge, "0.00") & "Ԫ"

    m_oPrintReturn.SetCurrentObject clRSheetID
    m_oPrintReturn.LabelSetCaption pszSheetID

    m_oPrintReturn.SetCurrentObject clRNowDate    '��ӡ����
    m_oPrintReturn.LabelSetCaption Format(Now, "YYYY-MM-DD HH:mm")

    m_oPrintReturn.SetCurrentObject clRUserID  '��ӡ����
    m_oPrintReturn.LabelSetCaption GetActiveUserID

'    m_oPrintReturn.SetCurrentObject clRNowDate2
'    m_oPrintReturn.LabelSetCaption Format(Date, "YYYY-MM-DD")

'    m_oPrintReturn.SetCurrentObject clRReturnCharge2   '��Ʊ������
'    m_oPrintReturn.LabelSetCaption Format(psgReturnCharge, "0.00") & "Ԫ"

    '����ʾ��Ʊ���Ρ�����ʱ���
    m_oPrintReturn.SetCurrentObject clRBusID  '
    m_oPrintReturn.LabelSetCaption ""
    m_oPrintReturn.SetCurrentObject clRStartUpTime
    m_oPrintReturn.LabelSetCaption ""
    
    
    '��ӡ��Ʊ����
    m_oPrintReturn.SetCurrentObject clRReturnCountStr
    m_oPrintReturn.LabelSetCaption "��Ʊ����:"
    m_oPrintReturn.SetCurrentObject clRTicketCount  '��ӡƱ������
    m_oPrintReturn.LabelSetCaption pnReturnCount & "��"
    m_oPrintReturn.SetCurrentObject clRTicketCount2  '��ӡƱ������
    m_oPrintReturn.LabelSetCaption pnReturnCount & "��"

'    m_oPrintReturn.SetCurrentObject clRReturnChargeStr  '��ӡ�������ַ���
'    m_oPrintReturn.LabelSetCaption "������"


    m_oPrintReturn.PrintA 100

End Function
'��ӡ��Ʊƾ֤
Public Function PrintSingleReturnSheet(pszTicketNo As String, pszSheetID As String, psgReturnCharge As Double, psgTicketPrice As Double, pdtStartUpTime As String _
    , pszBusID As String, pszEndStation As String, pszTicketType As String, pszVehicleType As String, pszSeatNo As String, pszStartStation As String _
    , pdtBusDate As Date) As Long
 #If PRINT_SHEET <> 0 Then
On Error GoTo Error_Handle
    If m_bUseFastPrint Then     '��������ٴ�ӡ������billPrint
        PrintSingleReturnSheetFast pszTicketNo, pszSheetID, psgReturnCharge, psgTicketPrice, pdtStartUpTime, pszBusID, pszEndStation, pszTicketType, pszVehicleType, pszSeatNo, pszStartStation, pdtBusDate
    Else
        PrintSingleReturnSheetSlow pszTicketNo, pszSheetID, psgReturnCharge, psgTicketPrice, pdtStartUpTime, pszBusID, pszEndStation, pszTicketType, pszVehicleType, pszSeatNo, pszStartStation, pdtBusDate
    End If
    Exit Function
Error_Handle:
    ShowErrorMsg
    #End If
End Function
'''��ӡ��Ʊƾ֤
''Public Function PrintSingleReturnSheet(pszTicketNo As String, psgReturnCharge As Double, psgTicketPrice As Double, pdtStartUpTime As String _
''    , pszBusID As String, pszEndStation As String, pszTicketType As String, pszVehicleType As String, pszSeatNo As String, pszStartStation As String _
''    , pdtBusDate As Date) As Long
''#If PRINT_SHEET <> 0 Then
''
''    m_oPrintReturn.ClosePort
''    m_oPrintReturn.OpenPort
''
''    m_oPrintReturn.SetObject clRTicketID
''    m_oPrintReturn.SetCaption pszTicketNo   '��ӡ��ǰƱ��
''
''    m_oPrintReturn.SetObject clRTicketPrice   'Ʊ��
''    m_oPrintReturn.SetCaption Format(psgTicketPrice, "0.00") & "Ԫ"
''
''    m_oPrintReturn.SetObject clRReturnCharge   '��Ʊ������
''    m_oPrintReturn.SetCaption Format(psgReturnCharge, "0.00") & "Ԫ"
''
''    m_oPrintReturn.SetObject clRSheetID
''    m_oPrintReturn.SetCaption "��Ʊ"
''
''    m_oPrintReturn.SetObject clRStartUpTime   '����ʱ��
''    m_oPrintReturn.SetCaption Format(pdtStartUpTime, "hh:mm")
''
''    m_oPrintReturn.SetObject clRUserID  '��ӡ����
''    m_oPrintReturn.SetCaption GetActiveUserID
''
''    m_oPrintReturn.SetObject clRBusID  '
''    m_oPrintReturn.SetCaption pszBusID
''
''    m_oPrintReturn.SetObject clREndStation  '
''    m_oPrintReturn.SetCaption pszEndStation
''
''    m_oPrintReturn.SetObject clRNowDate   '��ӡ����
''    m_oPrintReturn.SetCaption Format(Date, "YYYY-MM-DD")
''
''    m_oPrintReturn.SetObject clRTicketType  '
''    m_oPrintReturn.SetCaption pszTicketType
''
''    m_oPrintReturn.SetObject clREndStation2  '
''    m_oPrintReturn.SetCaption pszEndStation
''
''    m_oPrintReturn.SetObject clRTicketType2  '
''    m_oPrintReturn.SetCaption pszTicketType
''
''    m_oPrintReturn.SetObject clRStartUpTime2  '����ʱ��
''    m_oPrintReturn.SetCaption Format(pdtStartUpTime, "hh:mm")
''
''    m_oPrintReturn.SetObject clRTicketCount  '
''    m_oPrintReturn.SetCaption "1��"
''
''    m_oPrintReturn.SetObject clRTicketCount2  '
''    m_oPrintReturn.SetCaption "1��"
''
''    m_oPrintReturn.SetObject clRReturnCharge2  '
''    m_oPrintReturn.SetCaption Format(psgReturnCharge, "0.00") & "Ԫ"
''
''    m_oPrintReturn.SetObject clRReturnCountStr  '
''    m_oPrintReturn.SetCaption "��Ʊ����"
''
''    m_oPrintReturn.SetObject clRReturnChargeStr  '
''    m_oPrintReturn.SetCaption "������"
''
''    m_oPrintReturn.SetObject clRTicketID2  '
''    m_oPrintReturn.SetCaption pszTicketNo
''
''    m_oPrintReturn.SetObject clRBusID2  '
''    m_oPrintReturn.SetCaption pszBusID
''
''    m_oPrintReturn.SetObject clRVehicleType  '
''    m_oPrintReturn.SetCaption pszVehicleType
''
''
''    m_oPrintReturn.SetObject clRSeatNo '
''    m_oPrintReturn.SetCaption pszSeatNo
''
''    m_oPrintReturn.SetObject clRStartStation1
''    m_oPrintReturn.SetCaption pszStartStation
''
''    m_oPrintReturn.SetObject clRStartStation2
''    m_oPrintReturn.SetCaption pszStartStation
''
''    m_oPrintReturn.SetObject clRReturnChargeStr2
''    m_oPrintReturn.SetCaption "������"
''
''    m_oPrintReturn.SetObject clRBusDate1
''    m_oPrintReturn.SetCaption Format(pdtBusDate, "yyyy-mm-dd")
''
''    m_oPrintReturn.SetObject clRBusDate2
''    m_oPrintReturn.SetCaption Format(pdtBusDate, "yyyy-mm-dd")
''
''    m_oPrintReturn.PrintFile
''
''    m_oPrintReturn.ClosePort
''
''#End If
''
''End Function

Public Sub GetIniFile()

#If PRINT_SHEET <> 0 Then
'            Set m_oPrintTicket = New BPrint
    
On Error GoTo Error_Handle
    If Not m_bUseFastPrint Then     '��������ٴ�ӡ������billPrint
        If FileIsExist(App.Path & "\SellTk.bpf") Then
            m_oPrintTicket.ReadFormatFileA App.Path & "\SellTk.bpf"
        Else
            MsgBox "��ӡ�����ļ�""SellTk.bpf""δ�ҵ�,�޷�������Ʊ����", vbCritical
            End
        End If
        
        If FileIsExist(App.Path & "\ReturnTk.bpf") Then
            m_oPrintReturn.ReadFormatFileA App.Path & "\ReturnTk.bpf"
        Else
            MsgBox "��ӡ�����ļ�""ReturnTk.bpf""δ�ҵ�,�޷�������Ʊ����", vbCritical
            End
        End If
    Else
        If FileIsExist(App.Path & "\SellTk.ini") Then
            m_oFastPrint.ReadFormatFile App.Path & "\SellTk.ini"
        Else
            MsgBox "��ӡ�����ļ�""SellTk.ini""δ�ҵ�,�޷�������Ʊ����", vbCritical
            End
        End If
        
        If FileIsExist(App.Path & "\ReturnTk.ini") Then
            m_oFastReturn.ReadFormatFile App.Path & "\ReturnTk.ini"
        Else
            MsgBox "��ӡ�����ļ�""ReturnTk.ini""δ�ҵ�,�޷�������Ʊ����", vbCritical
            End
        End If
    End If
    
    Exit Sub
Error_Handle:
    ShowErrorMsg
#End If
        
End Sub


'��ӡƱ����
'szEndStation As String
'dtOffTime As Date
'szBusID As String
'szVehicleType As String
'szCheckGate As String
'bSaleChange As Boolean

Public Function PrintTicket(aptTicketInfo() As TPrintTicketParam, szBusDate() As String, nTicketCount() As Integer, szEndStation() As String _
, szOffTime() As String, szBusID() As String, szVehicleType() As String, szCheckGate() As String, bSaleChange() As Boolean, aszTerminateName() As String _
, szSellStationName() As String, panInsurance() As Integer) As Long

 #If PRINT_SHEET <> 0 Then
'    Close #1
'    Open "lpt1:" For Output As #1
    Dim i As Integer
    Dim szDateTemp As String
    Dim dtTemp As Date
    Dim bPrintAid As Boolean
    Dim bPrintBarCode As Boolean
    Dim iLen As Integer
    Dim nCount As Integer
    
'    Dim oParam As SystemParam
    
    Dim dScrollTime As Date
    Dim szTicketType As String
    Dim szRouteName As String
    
    iLen = 0
    On Error GoTo Error_Handle
    

    '��ӡ��������
    Dim szStationName As String
    Dim szOffTime1 As String
    Dim szCheckGate1 As String
    Dim szTicketPrice As String
    Dim szTicketType1 As String
    Dim szBusID1 As String
    Dim szSeatNo As String
    Dim szDate As String
    Dim szUserID As String
    Dim aszRouteAndTime() As String
    Dim szScrollTime As String
    Dim szStartStationName As String
    Dim szTicketNo As String
    
    Dim szTermination As String
    Dim szVehicleModel As String
    
    Dim szStartShortName As String
    
    
    Dim szCompanyID As String
    Dim szCompanyName As String
    
    
    Dim szInsurance As String
    
    If m_bUseFastPrint Then
        m_oFastPrint.ClosePort
        m_oFastPrint.OpenPort
    End If
    
    iLen = ArrayLength(szBusID)
'    If oParam Is Nothing Then Set oParam = m_oParam
    For nCount = 1 To iLen
        dtTemp = CDate(szBusDate(nCount))
        szDateTemp = Format(dtTemp, "YYYY") & "-" & Format(dtTemp, "MM") & "-" & Format(dtTemp, "DD")
'        aszRouteAndTime = m_oSell.GetRouteAndTime(CDate(szBusDate(nCount)), szBusID(nCount))
'
'        If ArrayLength(aszRouteAndTime) <> 0 Then
'            szRouteName = aszRouteAndTime(1)
'            szScrollTime = aszRouteAndTime(2)
'            szCompanyID = aszRouteAndTime(3)
'            szCompanyName = aszRouteAndTime(4)
'        Else
'            szRouteName = ""
'            szScrollTime = ""
'            szCompanyID = ""
'            szCompanyName = ""
'
'        End If
'
        For i = 1 To nTicketCount(nCount)
        
            szTicketNo = aptTicketInfo(nCount).aptPrintTicketInfo(i).szTicketNo  '�õ�Ʊ��
'            #If PRINT_SELLSTATION = 1 Then
                szStartStationName = szSellStationName(nCount)
'            #Else
'                '�����ӡ�����վ���ϳ�վ,����
'                szStartStationName = m_oSell.SellUnitShortName  '�õ����վ��
'                szStartStationName = Left(szStartStationName, Len(szStartStationName) - 1)
'                szStartShortName = Right(szStartStationName, 1)
'            #End If
            
            szStationName = szEndStation(nCount)  '�õ�վ��
    
            szTicketPrice = CStr(Format(aptTicketInfo(nCount).aptPrintTicketInfo(i).sgTicketPrice, "0.00")) '�õ�Ʊ��
            
            szTicketType = Trim(GetTicketTypeStr2(aptTicketInfo(nCount).aptPrintTicketInfo(i).nTicketType))

            szTicketType1 = Left(szTicketType, Len(szTicketType) - 1) '�õ�Ʊ��
            
            szUserID = GetActiveUserID   '�õ�����
            If bSaleChange(nCount) Then szUserID = szUserID & "[��]"
            
            
            szDate = szDateTemp   '�õ���������
            If Format(szOffTime(nCount), "hh:mm") = cszScrollBus Then
                szBusID1 = GetBusID(szBusID(nCount)) 'szRouteName  '�õ�����
            Else
                szBusID1 = GetBusID(szBusID(nCount))   '�õ�����
            End If
            
            szVehicleModel = szVehicleType(nCount) '�õ�����
            
            szTermination = Trim(aszTerminateName(nCount)) '�õ������յ�վ����

            szSeatNo = aptTicketInfo(nCount).aptPrintTicketInfo(i).szSeatNo '�õ���λ��
    
'            m_oPrintTicket.Setobject clStartupTime
            If GetPrintScrollMode Then
                If Format(szOffTime(nCount), "hh:mm") = cszScrollBus Then
                    '���Ϊ�������,
                    '��11:30��֮ǰ��ʱ,��ӡ������Ʊ��Ϊ11:30��֮ǰ
'                    �����ӡ����ʱ���֮ǰ
                    
                    If DateDiff("s", Time, cszMiddleTime) > 0 Then
                        szOffTime1 = cszMiddleTime & cszScrollBusTime
                    Else
                        szOffTime1 = CStr(Format(szScrollTime, "hh:mm")) & cszScrollBusTime  '�õ�����ʱ��
                    End If
                Else
                    szOffTime1 = CStr(Format(szOffTime(nCount), "hh:mm"))   '�õ�����ʱ��
                End If
            Else
                If Format(szOffTime(nCount), "hh:mm") = cszScrollBus Then
                
                    szOffTime1 = cszScrollBus  '�õ�����ʱ��
                Else
                    szOffTime1 = CStr(Format(szOffTime(nCount), "hh:mm"))
                End If
            End If

            szCheckGate1 = szCheckGate(nCount)       '�õ���Ʊ��
            Select Case Len(Trim(szCheckGate1))
                Case 1
                    szCheckGate1 = szCheckGate1 & Space(6)
                Case 2
                    szCheckGate1 = szCheckGate1 & Space(5)
                Case 3
                    szCheckGate1 = szCheckGate1 & Space(1)
            End Select
            
            
            
            szInsurance = IIf(panInsurance(nCount) = 0, "", "[��]")
            
            
            ''��ʼ��ӡ--------------------------------------------------------
            If m_bUseFastPrint Then
                
                m_oFastPrint.SetObject cnStartStation1
                m_oFastPrint.SetCaption szStartStationName
                
                m_oFastPrint.SetObject cnEndStation1
                m_oFastPrint.SetCaption szStationName
                
                m_oFastPrint.SetObject cnTicketType1
                m_oFastPrint.SetCaption szTicketType1
                
                m_oFastPrint.SetObject cnTicketPrice1
                m_oFastPrint.SetCaption szTicketPrice
                
                m_oFastPrint.SetObject cnStartStation2
                m_oFastPrint.SetCaption szStartStationName
                
                m_oFastPrint.SetObject cnEndStation2
                m_oFastPrint.SetCaption szStationName
                
                m_oFastPrint.SetObject cnTicketPrice2
                m_oFastPrint.SetCaption szTicketPrice
                
                m_oFastPrint.SetObject cnTicketType2
                m_oFastPrint.SetCaption szTicketType1
                
                m_oFastPrint.SetObject cnUserID
                m_oFastPrint.SetCaption szUserID
                
                
                m_oFastPrint.SetObject cnUserID2
                m_oFastPrint.SetCaption szUserID
                
                
                
                m_oFastPrint.SetObject cnBusDate1
                m_oFastPrint.SetCaption szDate
                
                m_oFastPrint.SetObject cnBusID1
                m_oFastPrint.SetCaption szBusID1
                
                m_oFastPrint.SetObject cnBusDate2
                m_oFastPrint.SetCaption szDate
                
                m_oFastPrint.SetObject cnBusID2
                m_oFastPrint.SetCaption szBusID1
                
                m_oFastPrint.SetObject cnSeat
                m_oFastPrint.SetCaption szSeatNo
                
                m_oFastPrint.SetObject cnSeat2
                m_oFastPrint.SetCaption szSeatNo
                
                
                
                m_oFastPrint.SetObject cnStartupTime
                m_oFastPrint.SetCaption szOffTime1
                
                m_oFastPrint.SetObject cnVehicleType
                m_oFastPrint.SetCaption szVehicleModel
                
                m_oFastPrint.SetObject cnCheckGate
                m_oFastPrint.SetCaption szCheckGate1
                
                m_oFastPrint.SetObject cnTicketNO
                m_oFastPrint.SetCaption Right(szTicketNo, 4)
                
                m_oFastPrint.SetObject cnTicketNO1
                m_oFastPrint.SetCaption szTicketNo
                
                m_oFastPrint.SetObject cnCompany
                m_oFastPrint.SetCaption szCompanyName
                
                '����
                m_oFastPrint.SetObject cnInsurance
                m_oFastPrint.SetCaption szInsurance
                
                '��ӡ�㣬Ϊ���ش�ӡͷ
                m_oFastPrint.SetObject cnReturnPoint
                m_oFastPrint.SetCaption "..."
                
                m_oFastPrint.PrintFile

            Else
                m_oPrintTicket.SetCurrentObject cnStartStation1
                m_oPrintTicket.LabelSetCaption szStartStationName
                
                m_oPrintTicket.SetCurrentObject cnEndStation1
                m_oPrintTicket.LabelSetCaption szStationName
                
                m_oPrintTicket.SetCurrentObject cnTicketType1
                m_oPrintTicket.LabelSetCaption szTicketType1
                
                m_oPrintTicket.SetCurrentObject cnTicketPrice1
                m_oPrintTicket.LabelSetCaption szTicketPrice
                
                m_oPrintTicket.SetCurrentObject cnStartStation2
                m_oPrintTicket.LabelSetCaption szStartStationName
                
                m_oPrintTicket.SetCurrentObject cnEndStation2
                m_oPrintTicket.LabelSetCaption szStationName
                
                m_oPrintTicket.SetCurrentObject cnTicketPrice2
                m_oPrintTicket.LabelSetCaption szTicketPrice
                
                m_oPrintTicket.SetCurrentObject cnTicketType2
                m_oPrintTicket.LabelSetCaption szTicketType1
                
                m_oPrintTicket.SetCurrentObject cnUserID
                m_oPrintTicket.LabelSetCaption szUserID
                
                
                m_oPrintTicket.SetCurrentObject cnUserID2
                m_oPrintTicket.LabelSetCaption szUserID
                
                
                m_oPrintTicket.SetCurrentObject cnBusDate1
                m_oPrintTicket.LabelSetCaption szDate
                
                m_oPrintTicket.SetCurrentObject cnBusID1
                m_oPrintTicket.LabelSetCaption szBusID1
                
                m_oPrintTicket.SetCurrentObject cnBusDate2
                m_oPrintTicket.LabelSetCaption szDate
                
                m_oPrintTicket.SetCurrentObject cnBusID2
                m_oPrintTicket.LabelSetCaption szBusID1
                
                
                m_oPrintTicket.SetCurrentObject cnSeat
                m_oPrintTicket.LabelSetCaption szSeatNo
                
                m_oPrintTicket.SetCurrentObject cnSeat2
                m_oPrintTicket.LabelSetCaption szSeatNo
                
                
                m_oPrintTicket.SetCurrentObject cnStartupTime
                m_oPrintTicket.LabelSetCaption szOffTime1
                
                m_oPrintTicket.SetCurrentObject cnVehicleType
                m_oPrintTicket.LabelSetCaption szVehicleModel
                
                m_oPrintTicket.SetCurrentObject cnCheckGate
                m_oPrintTicket.LabelSetCaption szCheckGate1
                
                m_oPrintTicket.SetCurrentObject cnTicketNO
                m_oPrintTicket.LabelSetCaption Right(szTicketNo, 4)
                
                m_oPrintTicket.SetCurrentObject cnTicketNO1
                m_oPrintTicket.LabelSetCaption szTicketNo
                
                m_oPrintTicket.SetCurrentObject cnInsurance
                m_oPrintTicket.LabelSetCaption szInsurance
                
                '��ӡ�㣬Ϊ���ش�ӡͷ
                m_oPrintTicket.SetCurrentObject cnReturnPoint
                m_oPrintTicket.LabelSetCaption "..."
                
                m_oPrintTicket.PrintA 100
            End If
            
            
        Next
    Next nCount
    If m_bUseFastPrint Then
        m_oFastPrint.ClosePort
    End If
'    Set oParam = Nothing
    Exit Function
Error_Handle:
    ShowErrorMsg
'    Set oParam = Nothing
    #End If
    
End Function
'��ӡ��Ʊƾ֤
Public Function PrintSingleReturnSheetSlow(pszTicketNo As String, pszSheetID As String, psgReturnCharge As Double, psgTicketPrice As Double, pdtStartUpTime As String _
    , pszBusID As String, pszEndStation As String, pszTicketType As String, pszVehicleType As String, pszSeatNo As String, pszStartStation As String _
    , pdtBusDate As Date) As Long
#If PRINT_SHEET <> 0 Then
    m_oPrintReturn.SetCurrentObject clRStartStation1
    m_oPrintReturn.LabelSetCaption pszStartStation
    
    m_oPrintReturn.SetCurrentObject clREndStation  '
    m_oPrintReturn.LabelSetCaption pszEndStation
    
    m_oPrintReturn.SetCurrentObject clRUserID  '��ӡ����
    m_oPrintReturn.LabelSetCaption GetActiveUserID
    
'    m_oPrintReturn.SetCurrentObject clRReturnChargeStr  '
'    m_oPrintReturn.LabelSetCaption "��Ʊ������:"
    
    m_oPrintReturn.SetCurrentObject clRReturnCharge   '��Ʊ������
    m_oPrintReturn.LabelSetCaption Format(psgReturnCharge, "0.00") & "Ԫ"

'    m_oPrintReturn.SetCurrentObject clRReturnTimeStr   '
'    m_oPrintReturn.LabelSetCaption "��Ʊʱ��:"

    m_oPrintReturn.SetCurrentObject clRNowDate   '��ӡ����
    m_oPrintReturn.LabelSetCaption Format(Now, "YYYY-MM-DD HH:mm")
    
    m_oPrintReturn.SetCurrentObject clRTicketID
    m_oPrintReturn.LabelSetCaption pszTicketNo   '��ӡ��ǰƱ��
     
'    m_oPrintReturn.SetCurrentObject clRSheetIDStr
'    m_oPrintReturn.LabelSetCaption "ƾ֤��:"
    
    m_oPrintReturn.SetCurrentObject clRSheetID
    m_oPrintReturn.LabelSetCaption pszSheetID

    '��ʾ��Ʊ���Ρ�����ʱ���
    m_oPrintReturn.SetCurrentObject clRBusID  '
    m_oPrintReturn.LabelSetCaption "(" & pszBusID & ")"
    m_oPrintReturn.SetCurrentObject clRStartUpTime   '����ʱ��
    m_oPrintReturn.LabelSetCaption pdtStartUpTime      'Format(pdtStartUpTime, "YYYY-MM-DD HH:mm")

    m_oPrintReturn.SetCurrentObject clRTicketPrice   'Ʊ��
    m_oPrintReturn.LabelSetCaption Format(psgTicketPrice, "0.00")
    
    '����ʾ��Ʊ����
    m_oPrintReturn.SetCurrentObject clRReturnCountStr
    m_oPrintReturn.LabelSetCaption ""
    m_oPrintReturn.SetCurrentObject clRTicketCount
    m_oPrintReturn.LabelSetCaption ""
'    m_oPrintReturn.SetCurrentObject clRTicketCount2
'    m_oPrintReturn.LabelSetCaption ""
    

    '����Ϊ�˴��
    m_oPrintReturn.SetCurrentObject cnReturnPoint
    m_oPrintReturn.LabelSetCaption "..."
    
    m_oPrintReturn.PrintA 100
    
#End If

End Function

'��ӡ��Ʊƾ֤
Private Function PrintSingleReturnSheetFast(pszTicketNo As String, pszSheetID As String, psgReturnCharge As Double, psgTicketPrice As Double, pdtStartUpTime As String _
    , pszBusID As String, pszEndStation As String, pszTicketType As String, pszVehicleType As String, pszSeatNo As String, pszStartStation As String _
    , pdtBusDate As Date) As Long
#If PRINT_SHEET <> 0 Then
    m_oFastReturn.ClosePort
    m_oFastReturn.OpenPort
    
    m_oFastReturn.SetObject clRTicketID
    m_oFastReturn.SetCaption pszTicketNo   '��ӡ��ǰƱ��
    
    m_oFastReturn.SetObject clRTicketPrice   'Ʊ��
    m_oFastReturn.SetCaption Format(psgTicketPrice, "0.00") & "Ԫ"

    m_oFastReturn.SetObject clRReturnCharge   '��Ʊ������
    m_oFastReturn.SetCaption Format(psgReturnCharge, "0.00") & "Ԫ"

    m_oFastReturn.SetObject clRSheetIDStr
'    m_oFastReturn.SetCaption "��Ʊƾ֤��:"
    
    m_oFastReturn.SetObject clRSheetID
    m_oFastReturn.SetCaption pszSheetID

    m_oFastReturn.SetObject clRStartUpTime   '����ʱ��
    m_oFastReturn.SetCaption pdtStartUpTime      'Format(pdtStartUpTime, "YYYY-MM-DD HH:mm")

    m_oFastReturn.SetObject clRUserID  '��ӡ����
    m_oFastReturn.SetCaption GetActiveUserID
    
    m_oFastReturn.SetObject clRBusID  '
    m_oFastReturn.SetCaption "(" & pszBusID & ")"
    
    m_oFastReturn.SetObject clREndStation  '
    m_oFastReturn.SetCaption pszEndStation
    
    m_oFastReturn.SetObject clRNowDate   '��ӡ����
    m_oFastReturn.SetCaption Format(Now, "YYYY-MM-DD HH:mm")
    
    m_oFastReturn.SetObject clRTicketType  '
    m_oFastReturn.SetCaption pszTicketType

    m_oFastReturn.SetObject clREndStation2  '
    m_oFastReturn.SetCaption pszEndStation

    m_oFastReturn.SetObject clRTicketType2  '
    m_oFastReturn.SetCaption pszTicketType

    m_oFastReturn.SetObject clRStartUpTime2  '����ʱ��
    m_oFastReturn.SetCaption Format(pdtStartUpTime, "HH:mm")

    m_oFastReturn.SetObject clRTicketCount  '
    m_oFastReturn.SetCaption "1��"

    m_oFastReturn.SetObject clRTicketCount2  '
    m_oFastReturn.SetCaption "1��"

    m_oFastReturn.SetObject clRReturnCharge2  '
    m_oFastReturn.SetCaption Format(psgReturnCharge, "0.00") & "Ԫ"

    m_oFastReturn.SetObject clRReturnCountStr  '
'    m_oFastReturn.SetCaption "��Ʊ����:"

    m_oFastReturn.SetObject clRReturnChargeStr  '
'    m_oFastReturn.SetCaption "��Ʊ������:"
    
    m_oFastReturn.SetObject clRReturnTimeStr   '
'    m_oFastReturn.SetCaption "��Ʊʱ��:"

    m_oFastReturn.SetObject clRTicketID2  '
    m_oFastReturn.SetCaption pszTicketNo

    m_oFastReturn.SetObject clRBusID2  '
    m_oFastReturn.SetCaption pszBusID

    m_oFastReturn.SetObject clRVehicleType  '
    m_oFastReturn.SetCaption pszVehicleType
    
    
    m_oFastReturn.SetObject clRSeatNo '
    m_oFastReturn.SetCaption pszSeatNo
    
    m_oFastReturn.SetObject clRStartStation1
    m_oFastReturn.SetCaption pszStartStation
    
    m_oFastReturn.SetObject clRStartStation2
    m_oFastReturn.SetCaption pszStartStation
    
    m_oFastReturn.SetObject clRReturnChargeStr2
'    m_oFastReturn.SetCaption "��Ʊ������:"
    
    m_oFastReturn.SetObject clRBusDate1
    m_oFastReturn.SetCaption Format(pdtBusDate, "yyyy-mm-dd")
    
    m_oFastReturn.SetObject clRBusDate2
    m_oFastReturn.SetCaption Format(pdtBusDate, "yyyy-mm-dd")
    
    '��������

    m_oFastReturn.SetObject clRUserID1  '��ӡ����
    m_oFastReturn.SetCaption GetActiveUserID
    
    m_oFastReturn.SetObject clRSeatNo1 '��λ��
    m_oFastReturn.SetCaption pszSeatNo
        
    m_oFastReturn.PrintFile
    
    m_oFastReturn.ClosePort
    
#End If

End Function




