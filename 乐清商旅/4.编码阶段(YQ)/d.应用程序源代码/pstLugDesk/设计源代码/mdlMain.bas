Attribute VB_Name = "mdlMain"
Option Explicit

Public Const m_cRegParamKey = "DataBaseSet"

Public Const cszLuggageAccount = "LugAcc"
'====================================================================
'���¶���ö��


'������״̬���ַ�������
Public Enum EStatusBarArea
    ESB_WorkingInfo = 1
    ESB_ResultCountInfo = 2
    ESB_UserInfo = 3
    ESB_LoginTime = 4
End Enum
'�Ի������״̬
Public Enum EFormStatus
    EFS_AddNew = 0
    EFS_Modify = 1
    EFS_Show = 2
    EFS_Delete = 3
End Enum
Private Declare Function GetComputerNameAPI Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'====================================================================
'����ȫ�ֱ�������
'--------------------------------------------------------------------
Public m_szLuggageNo As String '������
Public m_szLuggagePrefix As String '������ǰ׺
Public m_szTicketNoFromatStr As String
Public m_nLuggageNoNumLen As Integer '���������ֳ���

Public g_szAcceptSheetID   As String      '��ǰ������
Public g_szCarrySheetID  As String        '��ǰǩ������

Public moAcceptSheet As New AcceptSheet
Public moCarrySheet As New CarrySheet
Public moSysParam As New STLuggage.LuggageParam
Public moLugSvr As New LuggageSvr
Public m_oParam As New SystemParam
Public m_oBase As New BaseInfo
Public m_bIsRelationWithVehicleType As Boolean '�а��Ĺ�ʽ�Ƿ��복���й�ϵ

Public m_bIsDispSettlePriceInAccept As Boolean '�Ƿ�������ʱ��ʾӦ�����
Public m_bIsDispSettlePriceInCheck As Boolean '�Ƿ���ǩ��ʱ��ʾӦ�����
Public m_bIsSettlePriceFromAcceptInCheck As Boolean '�����˷��ǲ��Ǵ�����Ľ����˷��л��ܵõ�
Public m_bIsPrintCheckSheet As Boolean '�Ƿ��ӡǩ����


Public m_szCustom As String '�а���ӡ���Զ�����Ϣ


Public Const clActiveColor = &HFF0000
Public m_nCanSellDay  As Integer
Public m_oShell As New STShell.CommShell
Public m_oCmdDlg As New STShell.CommDialog
Public m_oAUser As ActiveUser

Public Const szAcceptTypeGeneral = "���"
Public Const szAcceptTypeMan = "��ͨ"
Public Const szPickTypeGeneral = "����"
Public Const szPickTypeEms = "�ͻ�"
Public Const szLuggageQucke = "���"

Public Const szAcceptStatus = 0 '"����"

Public m_oPrintTicket As FastPrint ' BPrint

Public m_oPrintCarrySheet As BPrint

Public g_rsPriceItem As Recordset   '�˷����¼


Public g_szOurCompany As String
'���¶����״�ö��
'���������״�

Public Enum PrintAcceptObjectIndexEnum
    PAI_LabelNo = 1 '��ǩ��
    PAI_Shipper = 3 '������
    PAI_CalWeight = 4 '����
    PAI_Picker = 5 '�ռ���
    PAI_Pack = 6 '��װ
    
    PAI_StartStation = 7 'ʼ��վ
    PAI_TransType = 8  '���˷�ʽ
    
    PAI_LongLuggageID = 10 '��Ʊ��
    
    PAI_UserName = 12 '��Ʊ��
    PAI_LuggageName = 13 '����
    
    PAI_EndStation = 15 '��վ��
    PAI_LuggageNumber = 16 '����
    PAI_TotalPriceBig = 17 '�ϼƴ�д
    PAI_TotalPrice = 18 '�ϼƣ�Сд��
    PAI_AcceptDate = 19 ' ����ʱ��
    PAI_Vehicle = 20 '���ƺ�
    PAI_OperationDate = 21 '��������
    PAI_ShipperPhone = 22 '�����˵绰
    PAI_PickerPhone = 23 '�ռ��˵绰
    PAI_PickerAddress = 24 '�ռ��˵�ַ
    PAI_ActWeight = 25 'ʵ��
    PAI_OverNumber = 26    '���ؼ���
    PAI_BusID = 27  '����
    PAI_StartTime = 28 '����ʱ��
    PAI_LicenseTagNo = 29 '���ƺ�
    PAI_BusDate = 30  '��������
    PAI_TotalPrice2 = 31 '�ϼ�Сд2
    PAI_TotalPriceName = 32 '�ϼ�Сд����
    
    PAI_Year = 33  '��
    PAI_Month = 34 '��
    PAI_Day = 35 '��
    
    PAI_BusID2 = 36  '����2
    PAI_Vehicle2 = 37 '���ƺ�2
    PAI_StartTime2 = 38  '����ʱ��2
    
    '��չ֧�ֲ���
    PAI_PriceItem1 = 40  'Ʊ��1���˷ѣ�
    PAI_PriceItem2 = 41  'Ʊ��2������ѣ�
    PAI_PriceItem3 = 42  'Ʊ��3�����Ž��ͷѣ�
    PAI_PriceItem4 = 43  '
    PAI_PriceItem5 = 44
    PAI_PriceItem6 = 45
    PAI_PriceItem7 = 46
    PAI_PriceItem8 = 47
    PAI_PriceItem9 = 48
    PAI_PriceItem10 = 49
    
    
    '��д���λ��
    PAI_Cent = 51
    PAI_Jiao = 52
    PAI_Yuan = 53
    PAI_Ten = 54
    PAI_Hundred = 55
    PAI_Thousand = 56
    
    
    
    PAI_StartStation2 = 60 'ʼ��վ2
    PAI_EndStation2 = 61 '��վ��2
    PAI_StartStation3 = 63 'ʼ��վ3
    PAI_EndStation3 = 64 '��վ��3
    PAI_StartStation4 = 65 'ʼ��վ4
    PAI_EndStation4 = 66 '��վ��4
    
    PAI_TransTicketID1 = 71    '���䵥��1
    PAI_TransTicketID2 = 72    '���䵥��2
    PAI_TransTicketID3 = 73    '���䵥��3
    PAI_TransTicketID4 = 74    '���䵥��4
    
    
    PAI_InsuranceID1 = 75    '���յ���1
    PAI_Annotation1 = 76    '��ע1
    PAI_Annotation2 = 77    '��ע2
    
    PAI_LuggageNumber2 = 80 '����2
    PAI_LuggageNumber3 = 81 '����3
    PAI_LuggageNumber4 = 82 '����4
    PAI_LuggageName2 = 83 '����2
    PAI_LuggageName3 = 84 '����3
    PAI_LuggageName4 = 85 '����4
    
    
    PAI_Year2 = 90  '��
    PAI_Month2 = 91 '��
    PAI_Day2 = 92 '��
    PAI_Year3 = 93  '��
    PAI_Month3 = 94 '��
    PAI_Day3 = 95 '��
    PAI_Year4 = 96  '��
    PAI_Month4 = 97 '��
    PAI_Day4 = 98 '��
    
    
    PAI_SettlePrice = 99 'Ӧ���˷�
    
    PAI_BasePrice = 100 '�а��˷�2
    PAI_BasePriceName = 101 '�а��˷�����
    PAI_SettlePriceName = 102 '�а�Ӧ���˷�����
    
    
    PAI_Custom1 = 110    '�Զ�����Ϣ1
    PAI_Custom2 = 111    '�Զ�����Ϣ2
    
    
    PAI_LongLuggageID2 = 112 '��Ʊ��
    PAI_LongLuggageID3 = 113 '��Ʊ��
    PAI_LongLuggageID4 = 114 '��Ʊ��
    
    PAI_ActWeight2 = 115 'ʵ��2
    PAI_CalWeight2 = 116 '����2
    PAI_OperationDate2 = 117 '��Ʊ����2
    PAI_TotalPriceBig2 = 118 '�ϼƴ�д2
    PAI_UserName2 = 119 '��Ʊ��2
    PAI_LabelNo2 = 120 '��ǩ2
    PAI_TransType2 = 121 '���˷�ʽ
    
    PAI_Mark = 122 '������
    PAI_Mark2 = 123 '������2
    
    
    
End Enum
'���������״�
Public Enum PrintReturnAcceptObjectIndexEnum
    PRI_LongLuggageID = 10 '��Ʊ��
    PRI_CredenceID = 9 '��Ʊƾ֤��
    PRI_TransType = 8  '���˷�ʽ
    PRI_StartStation = 7 'ʼ��վ
    PRI_EndStation = 15 '��վ��
    PRI_LuggageNumber = 16 '����
    PRI_Shipper = 3 '������
    PRI_CalWeight = 4 '����
    PRI_Picker = 5 '�ռ���
    PRI_LuggageName = 13 '����
    PRI_LabelNo = 1 '��ǩ��
    PRI_TotalPrice = 17 '�ϼƣ�Сд��
    PRI_TotalPriceBig = 18 '�ϼƴ�д
    PRI_ReturnCharge = 19 '���˷�
    PRI_ReturnChargeBig = 20 '���˷Ѵ�д
    PRI_UserName = 12 '��Ʊ��
    
    PRI_OperationDate = 21 '��������
    PRI_OverNumber = 26    '���ؼ���
    
    PRI_ReturnDate = 27 '����ʱ��
    PRI_ReturnChargeName = 28 '��������������
End Enum
'ǩ�����״�


Const cnDetailItemCount = 4    '��ϸ��Ŀ����

Public Enum PrintCarrySheetObjectIndexEnum
    '�а������嵥��
    PCI_SheetID = 1         'ǩ������
    PCI_TransType = 2  '���˷�ʽ
    PCI_StartStation = 3 'ʼ��վ
    PCI_Year = 4
    PCI_Month = 5
    PCI_Day = 6
    PCI_AcceptDetailItem = 7     '�����嵥
    PCI_UserID = 8
    
    
    '�а�������
    PCI_Year2 = 9
    PCI_Month2 = 10
    PCI_Day2 = 11
    PCI_SheetID2 = 12 'ǩ������
    PCI_EndStation = 13 '�յ�վ
    PCI_StartTime = 14 '����ʱ��
    PCI_LicenseTagNo = 15 '���ƺ�
    PCI_TotalPrice = 16 '���˷� �ϼƣ�Сд��
    PCI_TotalPriceBig = 17 '�ϼƴ�д
    PCI_UserID2 = 18
    PCI_Number = 19 '����
    
    
    'װ���嵥��
    PCI_Year3 = 20
    PCI_Month3 = 21
    PCI_Day3 = 22
    PCI_SheetID3 = 23         'ǩ������
    PCI_UserID3 = 24
    PCI_LicenseTagNo2 = 25 '���ƺ�
    PCI_StartTime2 = 26 '����ʱ��
    PCI_CarryDetailItem = 27     'װ���嵥
    
    
    PCI_ShipperPhone = 28     '�����˵绰�嵥
    
    PCI_CarryTime = 29       'ǩ��ʱ��
    PCI_CarryTime2 = 30     'ǩ��ʱ��2
    PCI_CarryTime3 = 31     'ǩ��ʱ��3
    PCI_TotalPrice2 = 32    '���˷�2 �ϼƣ�Сд)
    
    PCI_MoveWorker = 33 '�¼Ӵ�ӡ��װж����,fpd
End Enum

Public Function ToStandardDateStr(pdtDate As Date) As String
    ToStandardDateStr = Format(pdtDate, "YYYY-MM-DD")
End Function
' *******************************************************************
' *   Member Name: ShowSBInfo                                      *
' *   Brief Description: дϵͳ״̬����Ϣ                           *
' *   Engineer: ½����                                              *
' *******************************************************************
Public Sub ShowSBInfo(Optional pszInfo As String = "", Optional peArea As EStatusBarArea = ESB_WorkingInfo)
'����ע��
'*************************************
'pnArea(״̬������,Ĭ��ΪӦ�ó���״̬��)
'pszInfo(��Ϣ����)
'*************************************
    With mdiMain
        Select Case peArea
        Case EStatusBarArea.ESB_WorkingInfo
            .abMenu.Bands("statusBar").Tools("pnWorkingInfo").Caption = pszInfo
        Case EStatusBarArea.ESB_ResultCountInfo
            .abMenu.Bands("statusBar").Tools("pnResultCountInfo").Caption = pszInfo
        Case EStatusBarArea.ESB_UserInfo
            .abMenu.Bands("statusBar").Tools("progressBar").Visible = False
            .abMenu.Bands("statusBar").Tools("pnUserInfo").Caption = pszInfo
        Case EStatusBarArea.ESB_LoginTime
            If pszInfo <> "" Then pszInfo = "��¼ʱ��: " & pszInfo
            .abMenu.Bands("statusBar").Tools("pnLoginTime").Caption = pszInfo
        End Select
    End With
End Sub
' *******************************************************************
' *   Member Name: WriteProcessBar                                  *
' *   Brief Description: дϵͳ������״̬                           *
' *   Engineer: ½����                                              *
' *******************************************************************
Public Sub WriteProcessBar(Optional pbVisual As Boolean = True, Optional ByVal plCurrValue As Variant = 0, Optional ByVal plMaxValue As Variant = 0, Optional pszShowInfo As String = cszUnrepeatString)
'����ע��
'*************************************
'plCurrValue(��ǰ����ֵ)
'plMaxValue(������ֵ)
'*************************************
    If pszShowInfo <> cszUnrepeatString Then ShowSBInfo pszShowInfo, ESB_WorkingInfo
    If plMaxValue = 0 And pbVisual = True Then Exit Sub
    Dim nCurrProcess As Integer
    With mdiMain.abMenu.Bands("statusBar")
        If pbVisual Then
            If Not .Tools("progressBar").Visible Then
                .Tools("progressBar").Visible = True
                .Tools("pnResultCountInfo").Caption = ""
                .Tools("pnResultCountInfo").Visible = False
                mdiMain.pbLoad.Max = 100
                mdiMain.abMenu.RecalcLayout
            End If
            nCurrProcess = Int(plCurrValue / plMaxValue * 100)
            mdiMain.pbLoad.Value = nCurrProcess
        Else
            .Tools("progressBar").Visible = False
            .Tools("pnResultCountInfo").Visible = True
        End If
    End With
End Sub

Public Function GetComputerName() As String
    ' Set or retrieve the name of the computer.
    Dim strBuffer As String
    Dim lngLen As Long
        
    strBuffer = Space(255 + 1)
    lngLen = Len(strBuffer)
    If CBool(GetComputerNameAPI(strBuffer, lngLen)) Then
        GetComputerName = MidA(strBuffer, 1, lngLen)
    Else
        GetComputerName = ""
    End If
End Function

Sub Main()
 On Error GoTo Error_Handle
    Dim szLoginCommandLine As String
    Dim oLugSysParam As New STLuggage.LuggageParam
'     Dim oSysMan As New User
'    �жϴ�ӡ�Ƿ���ڣ������������
'    If App.PrevInstance Then
'        End
'    End If
    If Not IsPrinterValid Then
        MsgBox "��ӡ��δ���ã�", vbInformation, "��ӡ������:"
        End
        Exit Sub
    End If
    '��¼
    szLoginCommandLine = TransferLoginParam(Trim(Command()))
    If szLoginCommandLine = "" Then
        Set m_oAUser = m_oShell.ShowLogin()
        
    Else
        Set m_oAUser = New ActiveUser
        m_oAUser.Login GetLoginParam(szLoginCommandLine, cszUserName), GetLoginParam(szLoginCommandLine, cszUserPassword), GetComputerName()
        m_oCmdDlg.Init m_oAUser
    End If
    If Not m_oAUser Is Nothing Then
'        App.HelpFile = SetHTMLHelpStrings("SNSellTK.chm") '�趨App.HelpFile
        
        m_oParam.Init m_oAUser
        Date = m_oParam.NowDate
        Time = m_oParam.NowDateTime
        m_bIsRelationWithVehicleType = False 'm_oParam.IsRelationWithVehicleType

        m_bIsDispSettlePriceInAccept = False 'm_oParam.IsDispSettlePriceInAccept
        m_bIsDispSettlePriceInCheck = False 'm_oParam.IsDispSettlePriceInCheck
        m_bIsSettlePriceFromAcceptInCheck = False 'm_oParam.IsSettlePriceFromAcceptInCheck
        m_bIsPrintCheckSheet = False 'm_oParam.IsPrintCheckSheet
        
        SetHTMLHelpStrings "pstLugDesk.chm"
        
        
        
        '����ע������ע������
        '�а�ע������
        Dim oFreeReg As CFreeReg
        Set oFreeReg = New CFreeReg
        oFreeReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
        m_szCustom = oFreeReg.GetSetting(cszLuggageAccount, "CareContent") '�Զ�����Ϣ
        
        
        
        
        oLugSysParam.Init m_oAUser
        Set g_rsPriceItem = oLugSysParam.GetPriceItemRS(0)
        

        moAcceptSheet.Init m_oAUser
        moCarrySheet.Init m_oAUser
        moSysParam.Init m_oAUser
        moLugSvr.Init m_oAUser
        'frmSplash.Show

        Set m_oPrintTicket = New FastPrint
        Set m_oPrintCarrySheet = New BPrint
        '��ʼ����ӡǩ����
    
        m_oPrintTicket.ReadFormatFile App.Path & "\AcceptSheet.ini"
'        m_oPrintTicket.ClearAll
        m_oPrintCarrySheet.ReadFormatFileA App.Path & "\CarrySheet.bpf"
        
        g_szOurCompany = oLugSysParam.OurCompany

    '��frmChgSheetNo���壬������ʼ���а����ź�ǩ�����ţ� ����g_szAcceptSheetID��g_szCarrySheetID
    
    '�õ���ʼ���а����ź�ǩ������
        GetAppSetting2
        
        frmChgSheetNo.Show vbModal
        If Not frmChgSheetNo.m_bOk Then
            End
        End If
        Load mdiMain
        DoEvents
    '    ��ʾsplash����
        
        mdiMain.Show
        '�ر�splash����
        

    End If
  Exit Sub
Error_Handle:
ShowErrorMsg
End Sub

'�����������ϵı�ǩ��
Public Sub SetSheetNoLabel(pbIsAccept As Boolean, pszSheetNo As String)
    'pbIsAccept �Ƿ���������
    If pbIsAccept Then
        mdiMain.lblSheetNoName = "��ǰ������:"
    Else
        mdiMain.lblSheetNoName = "��ǰǩ������:"
    End If
    mdiMain.lblSheetNoName.Visible = True
    mdiMain.lblSheetNo.Visible = True
    mdiMain.lblSheetNo.Caption = FormatSheetID(pszSheetNo)
End Sub
'�����������ϵı�ǩ��
Public Sub HideSheetNoLabel()
    mdiMain.lblSheetNoName.Visible = False
    mdiMain.lblSheetNo.Visible = False
End Sub







'��ʼ������ǩ�����Ĵ���

Public Function TicketNoNumLen() As Integer
    If m_nLuggageNoNumLen = 0 Then
        m_nLuggageNoNumLen = moSysParam.LuggageIDNumberLen
    End If
    TicketNoNumLen = m_nLuggageNoNumLen
End Function


Public Function GetTicketNo(Optional pnOffset As Integer = 0) As String
    GetTicketNo = MakeTicketNo(m_szLuggageNo + pnOffset, m_szLuggagePrefix)
End Function

'ǩ����������
Public Sub IncSheetID(Optional pnOffset As Integer = 1, Optional pbNoShow As Boolean = False)
    g_szCarrySheetID = g_szCarrySheetID + pnOffset
    If Not pbNoShow Then
        mdiMain.lblSheetNoName.Caption = "��ǰǩ������:"
        mdiMain.lblSheetNo.Caption = FormatSheetID(g_szCarrySheetID)
    End If
End Sub


'����������
Public Sub IncTicketNo(Optional pnOffset As Integer = 1, Optional pbNoShow As Boolean = False)
    Dim oReg As New CFreeReg
    Dim szConnectName As String
    
    szConnectName = "Luggage"
    
    On Error GoTo ErrHandle
    
    oReg.Init "RTStation", HKEY_LOCAL_MACHINE, "Software\RTSoft"
    m_szLuggageNo = m_szLuggageNo + pnOffset
    If Not pbNoShow Then
        mdiMain.lblSheetNoName.Caption = "��ǰ������:"
        mdiMain.lblSheetNo.Caption = GetTicketNo()
'        g_szAcceptSheetID = GetTicketNo
    End If
    
    g_szAcceptSheetID = GetTicketNo
    oReg.SaveSetting szConnectName, "CurrentSheetID", g_szAcceptSheetID
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'����Ʊ��
Public Function MakeTicketNo(plTicketNo As Long, Optional pszPrefix As String = "") As String
   
    MakeTicketNo = pszPrefix & Format(plTicketNo, TicketNoFormatStr())
End Function

'Ʊ���ַ���
Private Function TicketNoFormatStr() As String
    Dim i As Integer
    If m_szTicketNoFromatStr = "" Then
        m_szTicketNoFromatStr = String(TicketNoNumLen(), "0")
    End If
    TicketNoFormatStr = m_szTicketNoFromatStr
End Function

'��Ʊ�ŷֽ��ǰ׺���������ֲ���
Public Function ResolveTicketNo(pszFullTicketNo, ByRef pszTicketPrefix As String) As Long
    Dim i As Integer, j As Integer
    Dim nCount As Integer, nTemp As Integer, nTicketPrefixLen As Integer
    
    pszFullTicketNo = Trim(pszFullTicketNo)
    nCount = Len(pszFullTicketNo)
    
    For i = 1 To nCount
        nTemp = Asc(Mid(pszFullTicketNo, nCount - i + 1, 1))
        If nTemp < vbKey0 Or nTemp > vbKey9 Then
            Exit For
        End If
    Next
    i = i - 1
    If i > 0 Then
        nTemp = TicketNoNumLen()
        nTemp = IIf(nTemp > i, i, nTemp)
        ResolveTicketNo = CLng(Right(pszFullTicketNo, nTemp))
        
        nTicketPrefixLen = m_oParam.LuggageIDPrefixLen
        If nTicketPrefixLen <= Len(pszFullTicketNo) Then
            pszTicketPrefix = Left(pszFullTicketNo, nTicketPrefixLen)
        Else
            pszTicketPrefix = pszFullTicketNo
        End If
        
    Else
        pszTicketPrefix = ""
        ResolveTicketNo = 0
    End If
    
End Function

Public Function FormatSheetID(pszCheckID As String)
    FormatSheetID = Format(IIf(pszCheckID <> "", pszCheckID, 0), String(moSysParam.CarrySheetIDNumberLen, "0"))
End Function

Public Sub GetAppSetting()
    Dim szLastTicketNo As String
    
    Dim szLastSheetID As String
    On Error GoTo here
    szLastSheetID = moLugSvr.GetLastSheetID(m_oAUser.UserID)
    szLastTicketNo = moLugSvr.GetLastLuggageID(m_oAUser.UserID)
    
    g_szCarrySheetID = FormatSheetID(szLastSheetID)
    
    m_szLuggageNo = ResolveTicketNo(szLastTicketNo, m_szLuggagePrefix)
    
    IncTicketNo , True
    IncSheetID , True
    
    Exit Sub
here:
    ShowErrorMsg
    
    
End Sub

'���ŸĵĶ�ע�����ĵ���
Public Sub GetAppSetting2()
    Dim szLastTicketNo As String
    Dim szLastSheetID As String
    Dim oReg As New CFreeReg
    Dim szConnectName As String

    
    oReg.Init "RTStation", HKEY_LOCAL_MACHINE, "Software\RTSoft"
    
    szConnectName = "Luggage"
    
    On Error GoTo here
    szLastSheetID = moLugSvr.GetLastSheetID(m_oAUser.UserID)
    szLastTicketNo = oReg.GetSetting(szConnectName, "CurrentSheetID", "")
    
    g_szCarrySheetID = FormatSheetID(szLastSheetID)
    
    m_szLuggageNo = ResolveTicketNo(IIf(Val(szLastTicketNo) - 1 >= 0, Val(szLastTicketNo) - 1, 0), m_szLuggagePrefix)
    m_szLuggageNo = GetTicketNo
    
    IncTicketNo , True
    IncSheetID , True
    
    Exit Sub
here:
    ShowErrorMsg
End Sub


'ˢ�½����ϵĵ��ݺ�
Public Sub RefreshCurrentSheetID()
    Dim szLastTicketNo As String
    Dim oReg As New CFreeReg
    Dim szConnectName As String
    
    
    oReg.Init "RTStation", HKEY_LOCAL_MACHINE, "Software\RTSoft"
    
    szConnectName = "Luggage"
    
    On Error GoTo here
    szLastTicketNo = oReg.GetSetting(szConnectName, "CurrentSheetID", "")
    
    m_szLuggageNo = szLastTicketNo
    
    mdiMain.lblSheetNoName.Caption = "��ǰ������:"
    mdiMain.lblSheetNo.Caption = GetTicketNo()
    g_szAcceptSheetID = GetTicketNo
    
    Exit Sub
here:
    ShowErrorMsg
End Sub

'��ӡ�а�����
Public Sub PrintAcceptSheet(poAcceptSheet As AcceptSheet, pdtBusStartTime As Date)
           
#If PRINT_SHEET <> 0 Then

''    m_oPrintTicket.ClearAll
''    m_oPrintTicket.ReadFormatFileA App.Path & "\AcceptSheet.bpf"
'
'    '����ʱ��
'    m_oPrintTicket.SetCurrentObject PAI_AcceptDate
'    m_oPrintTicket.LabelSetCaption Format(Date, "yyyy��MM��dd��")
'
'    '���ƺ�
'    m_oPrintTicket.SetCurrentObject PAI_Vehicle
'    m_oPrintTicket.LabelSetCaption Trim(frmAccept.cboVehicle.Text)
'
'    '���˷�ʽ
'    m_oPrintTicket.SetCurrentObject PAI_TransType
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.AcceptType
'
'    '��Ʊ��
'    m_oPrintTicket.SetCurrentObject PAI_LongLuggageID
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.SheetID
'
'    'ʼ��վ
'    m_oPrintTicket.SetCurrentObject PAI_StartStation
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.StartStationName
'
'    '��վ��
'    m_oPrintTicket.SetCurrentObject PAI_EndStation
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.DesStationName
'
'    '������
'    m_oPrintTicket.SetCurrentObject PAI_Shipper
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.Shipper
'
'    '�ռ���
'    m_oPrintTicket.SetCurrentObject PAI_Picker
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.Picker
'
'    '��װ
'    m_oPrintTicket.SetCurrentObject PAI_Pack
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.Pack
'
'
'    '����
'    m_oPrintTicket.SetCurrentObject PAI_LuggageName
'    m_oPrintTicket.LabelSetCaption Trim(poAcceptSheet.LuggageName)
'
'    '��ǩ��
'    m_oPrintTicket.SetCurrentObject PAI_LabelNo
'    m_oPrintTicket.LabelSetCaption Trim(IIf(poAcceptSheet.StartLabelID = poAcceptSheet.EndLabelID, poAcceptSheet.StartLabelID, poAcceptSheet.StartLabelID & "-" & poAcceptSheet.EndLabelID))
'
'    '����
'    m_oPrintTicket.SetCurrentObject PAI_LuggageNumber
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.Number
'
'    '����
'    m_oPrintTicket.SetCurrentObject PAI_CalWeight
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.CalWeight
'
'
'    '�շ���
'    Dim i As Integer
'    Dim atTmpPriceItem() As TLuggagePriceItem
'    atTmpPriceItem = poAcceptSheet.PriceItems
'    For i = 1 To ArrayLength(atTmpPriceItem)
'        m_oPrintTicket.SetCurrentObject PAI_PriceItem1 + Val(atTmpPriceItem(i).PriceID)        '�����Ч��
'        m_oPrintTicket.LabelSetCaption FormatMoney(atTmpPriceItem(i).PriceValue)
'    Next i
'
'    '����д
'    m_oPrintTicket.SetCurrentObject PAI_TotalPriceBig
'    m_oPrintTicket.LabelSetCaption GetNumber(poAcceptSheet.TotalPrice)
'
'    '���Сд
'    m_oPrintTicket.SetCurrentObject PAI_TotalPrice
'    m_oPrintTicket.LabelSetCaption FormatMoney(poAcceptSheet.TotalPrice)
'
'    '��Ʊ��
'    m_oPrintTicket.SetCurrentObject PAI_UserName
'    m_oPrintTicket.LabelSetCaption m_oAUser.UserID
'
'
'    '���Сд
'    m_oPrintTicket.SetCurrentObject PAI_TotalPrice2
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.TotalPrice
'
'    '��Ʊ��
'    m_oPrintTicket.SetCurrentObject PAI_TotalPriceName
'    m_oPrintTicket.LabelSetCaption "�˷�+�����"
'
'
''
'    '�ռ��˵绰
'    m_oPrintTicket.SetCurrentObject PAI_PickerPhone
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.PickerPhone
'
'    '�����˵绰
'    m_oPrintTicket.SetCurrentObject PAI_ShipperPhone
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.LuggageShipperPhone
'
'    '�ռ��˵�ַ
'    m_oPrintTicket.SetCurrentObject PAI_PickerAddress
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.PickerAddress
'
'    '����ʱ��
'    m_oPrintTicket.SetCurrentObject PAI_StartTime
'    m_oPrintTicket.LabelSetCaption Format(pdtBusStartTime, "MM-dd hh:mm")
'
'    '��
'    m_oPrintTicket.SetCurrentObject PAI_Year
'    m_oPrintTicket.LabelSetCaption Format(Date, "yyyy")
'
'    '��
'    m_oPrintTicket.SetCurrentObject PAI_Month
'    m_oPrintTicket.LabelSetCaption Format(Date, "MM")
'
'    '��
'    m_oPrintTicket.SetCurrentObject PAI_Day
'    m_oPrintTicket.LabelSetCaption Format(Date, "dd")
'
'
''    '��������
''    m_oPrintTicket.SetCurrentObject PAI_OperationDate
''    m_oPrintTicket.LabelSetCaption Format(poAcceptSheet.OperateTime, cszDateStr)
''
''    'ʵ��
''    m_oPrintTicket.SetCurrentObject PAI_ActWeight
''    m_oPrintTicket.LabelSetCaption poAcceptSheet.ActWeight
''
''    '���ؼ���
''    m_oPrintTicket.SetCurrentObject PAI_OverNumber
''    m_oPrintTicket.LabelSetCaption poAcceptSheet.OverNumber
''
'
'    '���˳���
'    m_oPrintTicket.SetCurrentObject PAI_BusID
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.BusID
''
'    '��������
'    m_oPrintTicket.SetCurrentObject PAI_BusDate
'    m_oPrintTicket.LabelSetCaption Format(poAcceptSheet.BusDate, cszDateStr)
'
'    m_oPrintTicket.PrintA 100

''    m_oPrintTicket.ReadFormatFile App.Path & "\AcceptSheet.bpf"
    
'
    m_oPrintTicket.ClosePort
    m_oPrintTicket.OpenPort
    
    '����ʱ��
    m_oPrintTicket.SetObject PAI_AcceptDate
    m_oPrintTicket.SetCaption Format(Date, "yyyy��MM��dd��")
    
    '���ƺ�
    m_oPrintTicket.SetObject PAI_Vehicle
    m_oPrintTicket.SetCaption Trim(frmAccept.cboVehicle.Text)
    
    '���ƺ�2
    m_oPrintTicket.SetObject PAI_Vehicle2
    m_oPrintTicket.SetCaption Trim(frmAccept.cboVehicle.Text)
     
    '���˷�ʽ
    m_oPrintTicket.SetObject PAI_TransType
    m_oPrintTicket.SetCaption poAcceptSheet.AcceptType
    
    '��Ʊ��
    m_oPrintTicket.SetObject PAI_LongLuggageID
    m_oPrintTicket.SetCaption poAcceptSheet.SheetID
    
    'ʼ��վ
    m_oPrintTicket.SetObject PAI_StartStation
    m_oPrintTicket.SetCaption poAcceptSheet.StartStationName
    
    '��վ��
    m_oPrintTicket.SetObject PAI_EndStation
    m_oPrintTicket.SetCaption poAcceptSheet.DesStationName
    
    '������
    m_oPrintTicket.SetObject PAI_Shipper
    m_oPrintTicket.SetCaption poAcceptSheet.Shipper
    
    '�ռ���
    m_oPrintTicket.SetObject PAI_Picker
    m_oPrintTicket.SetCaption poAcceptSheet.Picker
    
    '��װ
    m_oPrintTicket.SetObject PAI_Pack
    m_oPrintTicket.SetCaption poAcceptSheet.Pack
    
    
    '����
    m_oPrintTicket.SetObject PAI_LuggageName
    m_oPrintTicket.SetCaption Trim(poAcceptSheet.LuggageName)
    
    '��ǩ��
    m_oPrintTicket.SetObject PAI_LabelNo
    m_oPrintTicket.SetCaption Trim(IIf(poAcceptSheet.StartLabelID = poAcceptSheet.EndLabelID, poAcceptSheet.StartLabelID, poAcceptSheet.StartLabelID & "-" & poAcceptSheet.EndLabelID))
    
    '����
    m_oPrintTicket.SetObject PAI_LuggageNumber
    m_oPrintTicket.SetCaption poAcceptSheet.Number
    
    '����
    m_oPrintTicket.SetObject PAI_CalWeight
    m_oPrintTicket.SetCaption poAcceptSheet.CalWeight
    
    
    '�շ���
    Dim i As Integer
    Dim atTmpPriceItem() As TLuggagePriceItem
    atTmpPriceItem = poAcceptSheet.PriceItems
    For i = 1 To ArrayLength(atTmpPriceItem)
        m_oPrintTicket.SetObject PAI_PriceItem1 + Val(atTmpPriceItem(i).PriceID)        '�����Ч��
        m_oPrintTicket.SetCaption FormatMoney(atTmpPriceItem(i).PriceValue)
    Next i
    
    '����д
    m_oPrintTicket.SetObject PAI_TotalPriceBig
    m_oPrintTicket.SetCaption GetNumber(poAcceptSheet.TotalPrice)
    
    
    Dim aszFig() As String
    aszFig = ApartFig(poAcceptSheet.TotalPrice)
    
    
    '��д���λ��
    
    m_oPrintTicket.SetObject PAI_Cent
    m_oPrintTicket.SetCaption GetSplitNum(aszFig, 1)
    
    m_oPrintTicket.SetObject PAI_Jiao
    m_oPrintTicket.SetCaption GetSplitNum(aszFig, 2)
    
    m_oPrintTicket.SetObject PAI_Yuan
    m_oPrintTicket.SetCaption GetSplitNum(aszFig, 3)
    
    m_oPrintTicket.SetObject PAI_Ten
    m_oPrintTicket.SetCaption GetSplitNum(aszFig, 4)
    
    m_oPrintTicket.SetObject PAI_Hundred
    m_oPrintTicket.SetCaption GetSplitNum(aszFig, 5)
    
    m_oPrintTicket.SetObject PAI_Thousand
    m_oPrintTicket.SetCaption GetSplitNum(aszFig, 6)
    
    
    '���Сд
    m_oPrintTicket.SetObject PAI_TotalPrice
    m_oPrintTicket.SetCaption FormatMoney(poAcceptSheet.TotalPrice)
    
    '��Ʊ��
    m_oPrintTicket.SetObject PAI_UserName
    m_oPrintTicket.SetCaption m_oAUser.UserID
    
    
    '���Сд
    m_oPrintTicket.SetObject PAI_TotalPrice2
    m_oPrintTicket.SetCaption poAcceptSheet.TotalPrice
    
    '��Ʊ��
    m_oPrintTicket.SetObject PAI_TotalPriceName
    m_oPrintTicket.SetCaption "�˷�+�����"
    
    

    '�ռ��˵绰
    m_oPrintTicket.SetObject PAI_PickerPhone
    m_oPrintTicket.SetCaption poAcceptSheet.PickerPhone
    
    '�����˵绰
    m_oPrintTicket.SetObject PAI_ShipperPhone
    m_oPrintTicket.SetCaption poAcceptSheet.LuggageShipperPhone
    
    '�ռ��˵�ַ
    m_oPrintTicket.SetObject PAI_PickerAddress
    m_oPrintTicket.SetCaption poAcceptSheet.PickerAddress
    
    '����ʱ��
    m_oPrintTicket.SetObject PAI_StartTime
    m_oPrintTicket.SetCaption Format(pdtBusStartTime, "MM-dd hh:mm")

    '��
    m_oPrintTicket.SetObject PAI_Year
    m_oPrintTicket.SetCaption Format(Date, "yyyy")
    
    '��
    m_oPrintTicket.SetObject PAI_Month
    m_oPrintTicket.SetCaption Format(Date, "MM")
    
    '��
    m_oPrintTicket.SetObject PAI_Day
    m_oPrintTicket.SetCaption Format(Date, "dd")
    

    '��������
    m_oPrintTicket.SetObject PAI_OperationDate
    m_oPrintTicket.SetCaption Format(poAcceptSheet.OperateTime, cszDateStr)
'
    'ʵ��
    m_oPrintTicket.SetObject PAI_ActWeight
    m_oPrintTicket.SetCaption poAcceptSheet.ActWeight
'
'    '���ؼ���
'    m_oPrintTicket.SetObject PAI_OverNumber
'    m_oPrintTicket.SetCaption poAcceptSheet.OverNumber
'

    '���˳���
    m_oPrintTicket.SetObject PAI_BusID
    m_oPrintTicket.SetCaption poAcceptSheet.BusID
    
    '���˳���
    m_oPrintTicket.SetObject PAI_BusID2
    m_oPrintTicket.SetCaption poAcceptSheet.BusID
'
    '��������
    m_oPrintTicket.SetObject PAI_BusDate
    m_oPrintTicket.SetCaption Format(poAcceptSheet.BusDate, cszDateStr)
    
    
    'ʼ��վ
    m_oPrintTicket.SetObject PAI_StartStation2
    m_oPrintTicket.SetCaption poAcceptSheet.StartStationName
    
    '��վ��
    m_oPrintTicket.SetObject PAI_EndStation2
    m_oPrintTicket.SetCaption poAcceptSheet.DesStationName
    
    'ʼ��վ
    m_oPrintTicket.SetObject PAI_StartStation3
    m_oPrintTicket.SetCaption poAcceptSheet.StartStationName
    
    '��վ��
    m_oPrintTicket.SetObject PAI_EndStation3
    m_oPrintTicket.SetCaption poAcceptSheet.DesStationName
    
    'ʼ��վ
    m_oPrintTicket.SetObject PAI_StartStation4
    m_oPrintTicket.SetCaption poAcceptSheet.StartStationName
    
    '��վ��
    m_oPrintTicket.SetObject PAI_EndStation4
    m_oPrintTicket.SetCaption poAcceptSheet.DesStationName
    
    
    '���䵥��1
    m_oPrintTicket.SetObject PAI_TransTicketID1
    m_oPrintTicket.SetCaption poAcceptSheet.TransTicketID
    
    '���䵥��2
    m_oPrintTicket.SetObject PAI_TransTicketID2
    m_oPrintTicket.SetCaption poAcceptSheet.TransTicketID
    '���䵥��3
    m_oPrintTicket.SetObject PAI_TransTicketID3
    m_oPrintTicket.SetCaption poAcceptSheet.TransTicketID
    '���䵥��4
    m_oPrintTicket.SetObject PAI_TransTicketID4
    m_oPrintTicket.SetCaption poAcceptSheet.TransTicketID

    '���յ���1
    m_oPrintTicket.SetObject PAI_InsuranceID1
    m_oPrintTicket.SetCaption poAcceptSheet.InsuranceID

    '��ע1
    m_oPrintTicket.SetObject PAI_Annotation1
    m_oPrintTicket.SetCaption poAcceptSheet.Annotation1

    '��ע2
    m_oPrintTicket.SetObject PAI_Annotation2
    m_oPrintTicket.SetCaption poAcceptSheet.Annotation2

    
    '����
    m_oPrintTicket.SetObject PAI_LuggageNumber2
    m_oPrintTicket.SetCaption poAcceptSheet.Number

    '����
    m_oPrintTicket.SetObject PAI_LuggageNumber3
    m_oPrintTicket.SetCaption poAcceptSheet.Number

    '����
    m_oPrintTicket.SetObject PAI_LuggageNumber4
    m_oPrintTicket.SetCaption poAcceptSheet.Number

    


    '����
    m_oPrintTicket.SetObject PAI_LuggageName2
    m_oPrintTicket.SetCaption Trim(poAcceptSheet.LuggageName)
    '����
    m_oPrintTicket.SetObject PAI_LuggageName3
    m_oPrintTicket.SetCaption Trim(poAcceptSheet.LuggageName)
    '����
    m_oPrintTicket.SetObject PAI_LuggageName4
    m_oPrintTicket.SetCaption Trim(poAcceptSheet.LuggageName)
    
    
    '��
    m_oPrintTicket.SetObject PAI_Year2
    m_oPrintTicket.SetCaption Format(Date, "yyyy")
    '��
    m_oPrintTicket.SetObject PAI_Month2
    m_oPrintTicket.SetCaption Format(Date, "MM")
    '��
    m_oPrintTicket.SetObject PAI_Day2
    m_oPrintTicket.SetCaption Format(Date, "dd")
    '��
    m_oPrintTicket.SetObject PAI_Year3
    m_oPrintTicket.SetCaption Format(Date, "yyyy")
    '��
    m_oPrintTicket.SetObject PAI_Month3
    m_oPrintTicket.SetCaption Format(Date, "MM")
    '��
    m_oPrintTicket.SetObject PAI_Day3
    m_oPrintTicket.SetCaption Format(Date, "dd")
    '��
    m_oPrintTicket.SetObject PAI_Year4
    m_oPrintTicket.SetCaption Format(Date, "yyyy")
    '��
    m_oPrintTicket.SetObject PAI_Month4
    m_oPrintTicket.SetCaption Format(Date, "MM")
    '��
    m_oPrintTicket.SetObject PAI_Day4
    m_oPrintTicket.SetCaption Format(Date, "dd")
    

    'Ӧ���˷�
    m_oPrintTicket.SetObject PAI_SettlePrice
    m_oPrintTicket.SetCaption FormatMoney(poAcceptSheet.SettlePrice)
    
    '�а��˷�
    m_oPrintTicket.SetObject PAI_BasePrice
    m_oPrintTicket.SetCaption FormatMoney(atTmpPriceItem(1).PriceValue)

    '�а��˷�����
    m_oPrintTicket.SetObject PAI_BasePriceName
    m_oPrintTicket.SetCaption "�˷ѣ�"
    '�а�Ӧ���˷�����
    m_oPrintTicket.SetObject PAI_SettlePriceName
    m_oPrintTicket.SetCaption "Ӧ���˷ѣ�"
    
    
    '�Զ�����Ϣ1
    m_oPrintTicket.SetObject PAI_Custom1
    m_oPrintTicket.SetCaption m_szCustom
    
    
'    '�Զ�����Ϣ2
'    m_oPrintTicket.SetObject PAI_Custom2
'    m_oPrintTicket.SetCaption ""
'
'    PAI_Custom1 = 110    '�Զ�����Ϣ1
'    PAI_Custom2 = 111    '�Զ�����Ϣ2
'
'
'
'
'
'
'
'
    
    '��Ʊ��
    m_oPrintTicket.SetObject PAI_LongLuggageID2
    m_oPrintTicket.SetCaption poAcceptSheet.SheetID
    
    '��Ʊ��
    m_oPrintTicket.SetObject PAI_LongLuggageID3
    m_oPrintTicket.SetCaption poAcceptSheet.SheetID
    
    '��Ʊ��
    m_oPrintTicket.SetObject PAI_LongLuggageID4
    m_oPrintTicket.SetCaption poAcceptSheet.SheetID
    
    '����ʱ��2
    m_oPrintTicket.SetObject PAI_StartTime2
    m_oPrintTicket.SetCaption Format(pdtBusStartTime, "MM-dd hh:mm")
    
    '��ǩ��2
    m_oPrintTicket.SetObject PAI_LabelNo2
    m_oPrintTicket.SetCaption Trim(IIf(poAcceptSheet.StartLabelID = poAcceptSheet.EndLabelID, poAcceptSheet.StartLabelID, poAcceptSheet.StartLabelID & "-" & poAcceptSheet.EndLabelID))
    
    'ʵ��2
    m_oPrintTicket.SetObject PAI_ActWeight2
    m_oPrintTicket.SetCaption poAcceptSheet.ActWeight
    
    '����2
    m_oPrintTicket.SetObject PAI_CalWeight2
    m_oPrintTicket.SetCaption poAcceptSheet.CalWeight
    
    '��Ʊ����2
    m_oPrintTicket.SetObject PAI_OperationDate2
    m_oPrintTicket.SetCaption Format(poAcceptSheet.OperateTime, cszDateStr)
    
    '��Ʊ��2
    m_oPrintTicket.SetObject PAI_UserName2
    m_oPrintTicket.SetCaption poAcceptSheet.OperatorID
    
    '����д2
    m_oPrintTicket.SetObject PAI_TotalPriceBig2
    m_oPrintTicket.SetCaption GetNumber(poAcceptSheet.TotalPrice)
    
    '���˷�ʽ2
    m_oPrintTicket.SetObject PAI_TransType2
    m_oPrintTicket.SetCaption poAcceptSheet.AcceptType
    
    '������
    m_oPrintTicket.SetObject PAI_Mark
    m_oPrintTicket.SetCaption "��"
    
    '������2
    m_oPrintTicket.SetObject PAI_Mark2
    m_oPrintTicket.SetCaption "��"
    
    
    
    m_oPrintTicket.PrintFile

    m_oPrintTicket.ClosePort
    
#End If
End Sub

'ǩ������ӡ
Public Sub PrintCarrySheet(poCarrySheet As CarrySheet)
#If PRINT_SHEET <> 0 Then

    Dim i As Integer
    Dim j As Integer
    Dim oLugSvr As New LuggageSvr
    Dim rsAccept As Recordset
    Dim szDetail() As String
    Dim nNumber As Integer '�ܼ���
    Dim dbPrice As Double
    Dim szAcceptType As String
    Dim szStartStation As String
    Dim szEndStation As String '��վ
    Dim szLuggageName As String 'Ʒ��
    Dim nBaggageNumber As Integer '����
    Dim szPack As String '��װ3
    Dim dbCalWeight As Double '����
    Dim szLuggageID As String '������
    Dim szPicker As String '�ջ���
    Dim szPickerPhone As String '�ջ��˵绰

    Dim szBusEndStation As String

'    m_oPrintCarrySheet.ClearAll
'    m_oPrintCarrySheet.ReadFormatFileA App.Path & "\CarrySheet.bpf"


    On Error GoTo ErrorHandle
    Dim szYear As String
    Dim szMonth As String
    Dim szDay As String
    Dim szSenderTel As String '�����˵绰

    szYear = Year(poCarrySheet.OperateTime)
    szMonth = Month(poCarrySheet.OperateTime)
    szDay = Day(poCarrySheet.OperateTime)

    oLugSvr.Init m_oAUser
    Set rsAccept = oLugSvr.GetAcceptSheetRS(cszEmptyDateStr, cszForeverDateStr, , , , poCarrySheet.SheetID)

    '=====================================
    '�����嵥��
    '=====================================
    
    
    
    
    'ǩ������
    m_oPrintCarrySheet.SetCurrentObject PCI_SheetID
    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.SheetID

'    m_oPrintCarrySheet.SetCurrentObject PCI_TransType
'    m_oPrintCarrySheet.LabelSetCaption ""


'    m_oPrintCarrySheet.SetCurrentObject PCI_Year
'    m_oPrintCarrySheet.LabelSetCaption szYear
'
'
'    m_oPrintCarrySheet.SetCurrentObject PCI_Month
'    m_oPrintCarrySheet.LabelSetCaption szMonth
'
'    m_oPrintCarrySheet.SetCurrentObject PCI_Day
'    m_oPrintCarrySheet.LabelSetCaption szDay

'    m_oPrintCarrySheet.SetCurrentObject PCI_UserID
'    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.OperatorID

'    'ʼ��վ
'    m_oPrintCarrySheet.SetCurrentObject PCI_StartStation
'    m_oPrintCarrySheet.LabelSetCaption szStartStation

    rsAccept.MoveFirst
    m_oPrintCarrySheet.SetCurrentObject PCI_AcceptDetailItem '�����嵥���
    szSenderTel = ""
    For i = 1 To cnDetailItemCount
        If Not rsAccept.EOF Then
            szSenderTel = szSenderTel & FormatDbValue(rsAccept!shipper_phone) & " "

            szAcceptType = GetLuggageTypeString(FormatDbValue(rsAccept!accept_type))    '�õ���һ����¼����������
'            szStartStation = FormatDbValue(rsAccept!start_station_name) '���վ
            szLuggageName = FormatDbValue(rsAccept!luggage_name)
            szEndStation = FormatDbValue(rsAccept!des_station_name)
            nBaggageNumber = FormatDbValue(rsAccept!baggage_number)
            szLuggageID = FormatDbValue(rsAccept!luggage_id)
            szPack = FormatDbValue(rsAccept!Pack)
            szPicker = FormatDbValue(rsAccept!Picker)
            szPickerPhone = FormatDbValue(rsAccept!picker_phone)
            dbCalWeight = FormatDbValue(rsAccept!cal_weight)
            ReDim szDetail(1 To 9)
            szDetail(1) = szEndStation
            szDetail(2) = szLuggageID
            szDetail(3) = szLuggageName
            szDetail(4) = nBaggageNumber
            szDetail(5) = szPack
            szDetail(6) = ""
            szDetail(7) = dbCalWeight
            szDetail(8) = szPicker
            szDetail(9) = szPickerPhone




            For j = 1 To 9

                m_oPrintCarrySheet.GridGetCertainCell i, j
                m_oPrintCarrySheet.GridLabelSetCaption szDetail(j)
            Next j


            nNumber = nNumber + FormatDbValue(rsAccept!baggage_number)
            dbPrice = dbPrice + FormatDbValue(rsAccept!price_item_1)        '�˷Ѻϼ�

            rsAccept.MoveNext

        Else

            '���ʣ���,�ÿմ�

            For j = 1 To 9
                m_oPrintCarrySheet.GridGetCertainCell i, j
                m_oPrintCarrySheet.GridLabelSetCaption ""
            Next j
        End If
    Next i
    
    m_oPrintCarrySheet.SetCurrentObject PCI_CarryTime
    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.OperateTime


'    '�а�������
    '=====================================
    '�а�������
    '=====================================

    'ǩ������
    m_oPrintCarrySheet.SetCurrentObject PCI_SheetID2
    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.SheetID

'    m_oPrintCarrySheet.SetCurrentObject PCI_Year2
'    m_oPrintCarrySheet.LabelSetCaption szYear
'
'
'    m_oPrintCarrySheet.SetCurrentObject PCI_Month2
'    m_oPrintCarrySheet.LabelSetCaption szMonth
'
'    m_oPrintCarrySheet.SetCurrentObject PCI_Day2
'    m_oPrintCarrySheet.LabelSetCaption szDay

'    m_oPrintCarrySheet.SetCurrentObject PCI_UserID2
'    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.OperatorID

    szBusEndStation = poCarrySheet.EndStation

    m_oPrintCarrySheet.SetCurrentObject PCI_EndStation
    m_oPrintCarrySheet.LabelSetCaption szBusEndStation


    '������(���Ÿ�Ϊ����)
    m_oPrintCarrySheet.SetCurrentObject PCI_LicenseTagNo
    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.BusID

'    '����ʱ��
'    m_oPrintCarrySheet.SetCurrentObject PCI_StartTime
'    m_oPrintCarrySheet.LabelSetCaption Format(poCarrySheet.BusStartOffTime, "HH:mm")

    '����
    m_oPrintCarrySheet.SetCurrentObject PCI_Number
    m_oPrintCarrySheet.LabelSetCaption nNumber

    '����д
    m_oPrintCarrySheet.SetCurrentObject PCI_TotalPriceBig
    m_oPrintCarrySheet.LabelSetCaption GetNumber(poCarrySheet.PrintSettlePrice)   'GetNumber(dbPrice)

    '���Сд
    m_oPrintCarrySheet.SetCurrentObject PCI_TotalPrice
    m_oPrintCarrySheet.LabelSetCaption FormatMoney(poCarrySheet.PrintSettlePrice) 'FormatMoney(dbPrice)
    
    m_oPrintCarrySheet.SetCurrentObject PCI_CarryTime2
    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.OperateTime

'
'
'End Enum

    '=====================================
    'װ���嵥��
    '=====================================

    'ǩ������
    m_oPrintCarrySheet.SetCurrentObject PCI_SheetID3
    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.SheetID

'    m_oPrintCarrySheet.SetCurrentObject PCI_Year3
'    m_oPrintCarrySheet.LabelSetCaption szYear
'
'
'    m_oPrintCarrySheet.SetCurrentObject PCI_Month3
'    m_oPrintCarrySheet.LabelSetCaption szMonth
'
'    m_oPrintCarrySheet.SetCurrentObject PCI_Day3
'    m_oPrintCarrySheet.LabelSetCaption szDay

'    m_oPrintCarrySheet.SetCurrentObject PCI_UserID3
'    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.OperatorID

    '������
    m_oPrintCarrySheet.SetCurrentObject PCI_LicenseTagNo2
    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.VehicleLicense

'    '����ʱ��
'    m_oPrintCarrySheet.SetCurrentObject PCI_StartTime2
'    m_oPrintCarrySheet.LabelSetCaption Format(poCarrySheet.BusStartOffTime, "HH:mm")

    m_oPrintCarrySheet.SetCurrentObject PCI_CarryDetailItem  'װ���嵥

    rsAccept.MoveFirst
    For i = 1 To cnDetailItemCount
        If Not rsAccept.EOF Then
'            szAcceptType = GetLuggageTypeString(FormatDbValue(rsAccept!accept_type))    '�õ���һ����¼����������
            szEndStation = FormatDbValue(rsAccept!des_station_name)
            szLuggageID = FormatDbValue(rsAccept!luggage_id)
            szLuggageName = FormatDbValue(rsAccept!luggage_name)
            nBaggageNumber = FormatDbValue(rsAccept!baggage_number)
            szPack = FormatDbValue(rsAccept!Pack)

            ReDim szDetail(1 To 5)
            szDetail(1) = szEndStation
            szDetail(2) = szLuggageID
            szDetail(3) = szLuggageName
            szDetail(4) = nBaggageNumber
            szDetail(5) = szPack


            For j = 1 To 5

                m_oPrintCarrySheet.GridGetCertainCell i, j
                m_oPrintCarrySheet.GridLabelSetCaption szDetail(j)
            Next j

            rsAccept.MoveNext

        Else

            '���ʣ���,�ÿմ�

            For j = 1 To 5
                m_oPrintCarrySheet.GridGetCertainCell i, j
                m_oPrintCarrySheet.GridLabelSetCaption ""
            Next j
        End If
    Next i

    m_oPrintCarrySheet.SetCurrentObject PCI_ShipperPhone
    m_oPrintCarrySheet.LabelSetCaption Trim(szSenderTel)
    
    m_oPrintCarrySheet.SetCurrentObject PCI_CarryTime3
    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.OperateTime
    
    m_oPrintCarrySheet.SetCurrentObject PCI_TotalPrice2
    m_oPrintCarrySheet.LabelSetCaption FormatMoney(poCarrySheet.PrintSettlePrice)
    
    '�¼ӡ�װж������ӡ,fpd
    m_oPrintCarrySheet.SetCurrentObject PCI_MoveWorker
    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.MoveWorker
    
    m_oPrintCarrySheet.PrintA 100

#End If


    Exit Sub
ErrorHandle:
    ShowErrorMsg


End Sub

'�а���Ʊ��ӡ
Public Sub PrintReturnAccept(poAcceptSheet As AcceptSheet, pszCredenceID As String, pdbReturnCharge As Double, pszOperator As String, pdtReturnTime As Date)
           
#If PRINT_SHEET <> 0 Then

'    m_oPrintTicket.ClearAll
'    m_oPrintTicket.ReadFormatFileA App.Path & "\ReturnAccept.bpf"
'
'    m_oPrintTicket.ClosePort
'    m_oPrintTicket.OpenPort
'        '����ʱ��
'    m_oPrintTicket.SetCurrentObject PRI_ReturnDate
'    m_oPrintTicket.LabelSetCaption Format(Date, "yyyy��MM��dd��")
'
'    '����ƾ֤��
'    m_oPrintTicket.SetCurrentObject PRI_CredenceID
'    m_oPrintTicket.LabelSetCaption pszCredenceID
'
'    '���˷�ʽ
'    m_oPrintTicket.SetCurrentObject PRI_TransType
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.AcceptType
'
'    '��Ʊ��
'    m_oPrintTicket.SetCurrentObject PRI_LongLuggageID
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.SheetID
'
'    'ʼ��վ
'    m_oPrintTicket.SetCurrentObject PRI_StartStation
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.StartStationName
'
'    '��վ��
'    m_oPrintTicket.SetCurrentObject PRI_EndStation
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.DesStationName
'
'    '������
'    m_oPrintTicket.SetCurrentObject PRI_Shipper
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.Shipper
'
'    '�ռ���
'    m_oPrintTicket.SetCurrentObject PRI_Picker
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.Picker
'
'    '����
'    m_oPrintTicket.SetCurrentObject PRI_LuggageName
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.LuggageName
'
'    '��ǩ��
'    m_oPrintTicket.SetCurrentObject PRI_LabelNo
'    m_oPrintTicket.LabelSetCaption IIf(poAcceptSheet.StartLabelID = poAcceptSheet.EndLabelID, poAcceptSheet.StartLabelID, poAcceptSheet.StartLabelID & "-" & poAcceptSheet.EndLabelID)
'
'    '����
'    m_oPrintTicket.SetCurrentObject PRI_LuggageNumber
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.Number
'
'    '����������
'    m_oPrintTicket.SetCurrentObject PRI_ReturnCharge
'    m_oPrintTicket.LabelSetCaption pdbReturnCharge
'
'    '��������������
'    m_oPrintTicket.SetCurrentObject PRI_ReturnChargeName
'    m_oPrintTicket.LabelSetCaption "����������"
'
'    m_oPrintTicket.SetCurrentObject PRI_CalWeight
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.CalWeight
'
'    '����д
'    m_oPrintTicket.SetCurrentObject PRI_TotalPriceBig
'    m_oPrintTicket.LabelSetCaption GetNumber(poAcceptSheet.TotalPrice)
'
'    '���Сд
'    m_oPrintTicket.SetCurrentObject PRI_TotalPrice
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.TotalPrice
'
'    '������
'    m_oPrintTicket.SetCurrentObject PRI_UserName
'    m_oPrintTicket.LabelSetCaption m_oAUser.UserName
'
'    '��������
'    m_oPrintTicket.SetCurrentObject PRI_OperationDate
'    m_oPrintTicket.LabelSetCaption Format(pdtReturnTime, cszDateStr)
'
'
'    '���ؼ���
'    m_oPrintTicket.SetCurrentObject PRI_OverNumber
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.OverNumber
'
'    m_oPrintTicket.PrintA 100

'    m_oPrintTicket.ClosePort
#End If

End Sub

'���˷�ʽ״̬ת��
Public Function GetLuggageTypeString(pnType As Integer) As String
    Select Case pnType
        Case 0
            GetLuggageTypeString = "���"
        Case 1
            GetLuggageTypeString = "����"
    End Select
End Function
Public Function GetLuggageTypeInt(szType As String) As Integer
    Select Case szType
        Case "���"
            GetLuggageTypeInt = 0
        Case "����"
            GetLuggageTypeInt = 1
    End Select
    
End Function


Private Function AlignTextToCell(pszString As String, pnFixWidth As Integer) As String
    Dim nActLen As Integer
    nActLen = LenA(pszString)
    If nActLen > pnFixWidth Then
        AlignTextToCell = MidA(pszString, 1, pnFixWidth)
    Else
        AlignTextToCell = Space(pnFixWidth - nActLen) & pszString
    End If
End Function

Private Function GetSplitNum(paszNum() As String, pnNum As Integer) As String
    
    Dim nLen As Integer
    Const cszO = "��"
    
    nLen = ArrayLength(paszNum)
    If pnNum > nLen Then
        GetSplitNum = cszO
    Else
        If paszNum(pnNum) = "" Then
            GetSplitNum = cszO
        
        Else
            GetSplitNum = paszNum(pnNum)
        End If
    End If
    
End Function

'�õ����س��εķ���ʱ��
Public Function GetAllotStationBusStartTime(ByVal pszBusID As String, ByVal pdtBusDate As Date) As Date
On Error GoTo ErrHandle

    Dim oLugSvr As New LuggageSvr
    Dim rsTemp As Recordset
    oLugSvr.Init m_oAUser
    Set rsTemp = oLugSvr.GetAllotStationBusStartTime(pszBusID, pdtBusDate)
    If rsTemp.RecordCount = 1 Then
        GetAllotStationBusStartTime = FormatDbValue(rsTemp!bus_date) & " " & FormatDbValue(rsTemp!bus_start_time)
    Else
        GetAllotStationBusStartTime = cdtEmptyDate
    End If
    
    Exit Function
ErrHandle:
    ShowErrorMsg
End Function
