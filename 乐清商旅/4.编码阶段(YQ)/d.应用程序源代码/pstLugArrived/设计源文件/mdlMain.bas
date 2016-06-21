Attribute VB_Name = "mdlMain"
Option Explicit



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

Const cszRecentSeller = "RecentSeller"
'====================================================================
'����ȫ�ֳ�������
'--------------------------------------------------------------------
Public Const CSZNoneString = "(ȫ��)"
Public Const CPick_Normal = "δ��"
Public Const CPick_Picked = "����"
Public Const CPick_Canceled = "�ѷ�"

Public Const cnColor_Active = &HFF0000
Public Const cnColor_Normal = vbBlack
Public Const cnColor_Edited = vbRed
Public Const cszLongDateFormat = "yyyy��MM��dd��"

Private Declare Function GetComputerNameAPI Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'====================================================================
'����ȫ�ֱ�������
'--------------------------------------------------------------------
Public g_szSheetID   As String      '��ǰ������
Public g_oActUser As ActiveUser
Public g_oParam As New SystemParam
Public g_oPackageParam As New PackageParam
Public g_oPackageSvr As New PackageSvr
Public g_oShell As New STShell.CommShell
    '�����Ǽ�����ź�������ȫ�ֱ���
Public g_bSMSValid As Boolean
Public g_szSMSServer As String
Public g_szSMSUser As String
Public g_szSMSPassword As String
Public g_szSMSApiPort As String

Public Enum PrintPackageObjectIndexEnum
    PPI_PackageID = 1 '�а���ˮ��
    PPI_SheetID = 2 '��ƱƱ�ݺ�
    PPI_PackageName = 3 '����
    PPI_PackType = 5 '��װ
    PPI_ArrivedDate = 6 ' ����ʱ��
    PPI_CalWeight = 7 '����
    PPI_AreaType = 8 '�������
    PPI_StartStation = 9 'ʼ��վ
    PPI_Vehicle = 10 '���ƺ�
    PPI_PackageNumber = 11 '����
    PPI_SavePosition = 12 '���λ��
    
    PPI_Shipper = 13 '������
    PPI_ShipperPhone = 14 '�����˵绰
    PPI_ShipperUnit = 15 '�����˵�λ
    PPI_PickType = 16  '������ʽ
    
    PPI_Picker = 17 '�ռ���
    PPI_PickerPhone = 18 '�ռ��˵绰
    PPI_PickerUnit = 19 '�ռ��˵�λ
    PPI_PickerAddress = 20 '�ռ��˵�ַ
    PPI_PickerCredit = 21 '��ȡ�����֤��
    PPI_PickTime = 22 '���ʱ��
    
    PPI_Operator = 23 '������
    PPI_OperationDate = 24 '��������
    PPI_UserName = 25 '�����û�
    PPI_Loader = 26 'װж��
    
    PPI_TransCharge = 27 '�����˷�
    
    PPI_TotalPriceBig = 28 '�ϼƴ�д
    PPI_TotalPrice = 29 '�ϼƣ�Сд��
    PPI_PriceItem1 = 40  'Ʊ��1��װж�ѣ�
    PPI_PriceItem2 = 41  'Ʊ��2�����ܷѣ�
    PPI_PriceItem3 = 42  'Ʊ��3���ͻ��ѣ�
    PPI_PriceItem4 = 43  'Ʊ��4�����˷ѣ�
    PPI_PriceItem5 = 44  'Ʊ��5�������ѣ�
    
    PPI_Year = 51  '��
    PPI_Month = 52 '��
    PPI_Day = 53 '��
    
    '��������
    PPI_AreaType2 = 30 '�������
    PPI_StartStation2 = 31 'ʼ��վ
    PPI_PackageName2 = 32 '����
    PPI_PackageNumber2 = 33 '����
    PPI_PackType2 = 34 '��װ
    
    PPI_TotalPrice2 = 35 '�ϼ�Сд2
    
    
    PPI_Year2 = 54  '��
    PPI_Month2 = 55 '��
    PPI_Day2 = 56 '��
    
    
    '��д���λ��
    PAI_Cent = 61
    PAI_Jiao = 62
    PAI_Yuan = 63
    PAI_Ten = 64
    PAI_Hundred = 65
    PAI_Thousand = 66
    
    PPI_PackageID2 = 70 '������ˮ��2
    PPI_Drawer = 71   '�����
    PPI_DrawerPhone = 72   '����˵绰
    PPI_SheetID2 = 73   'Ʊ�ݺ�2
    
    PPI_ArrivedDate2 = 84 ' ����ʱ��2
    PPI_CalWeight2 = 85 '����2
    PPI_CalWeight3 = 86 '����3
    PPI_CalWeight4 = 87 '����4
    PPI_Vehicle2 = 88 '���ƺ�2
    PPI_TransCharge2 = 89 '�����˷�2
    PPI_TotalPriceBig2 = 90 '�ϼƴ�д2
    PPI_Picker2 = 91 '�ռ���2
    PPI_Drawer2 = 92   '�����2
    PPI_PickTime2 = 93 '���ʱ��2
    PPI_UserName2 = 94 '�����û�2
    
    PPI_Mark = 95 '������
    PPI_Mark2 = 96 '������2
    
End Enum

Dim m_oPrintTicket As FastPrint


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
        Set g_oActUser = g_oShell.ShowLogin()

    Else
        Set g_oActUser = New ActiveUser
        g_oActUser.Login GetLoginParam(szLoginCommandLine, cszUserName), GetLoginParam(szLoginCommandLine, cszUserPassword), GetComputerName()
    End If
    If Not g_oActUser Is Nothing Then
'        App.HelpFile = SetHTMLHelpStrings("SNSellTK.chm") '�趨App.HelpFile
        
        g_oParam.init g_oActUser
        Date = g_oParam.NowDate
        Time = g_oParam.NowDateTime

        
        g_oPackageParam.init g_oActUser
        
        g_oPackageSvr.init g_oActUser
'        SetHTMLHelpStrings "pstLugDesk.chm"
        
        'frmSplash.Show

        Set m_oPrintTicket = New FastPrint
        '��ʼ����ӡǩ����
        FileIsExist (App.Path & "\PackageSheet.ini")
        m_oPrintTicket.ReadFormatFile App.Path & "\PackageSheet.ini"
        
        
    '��frmChgSheetNo���壬������ʼ�ĵ��ݺţ� ����g_szSheetID
    
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


'ǩ����������
Public Sub IncSheetID(Optional pnOffset As Integer = 1, Optional pbNoShow As Boolean = False)
 On Error GoTo ErrHandle
 
    Dim oReg As New CFreeReg
    Dim szConnectName As String
    
    szConnectName = "Luggage"
    
    oReg.init "RTStation", HKEY_LOCAL_MACHINE, "Software\RTSoft"
    g_szSheetID = FormatSheetID(g_szSheetID + pnOffset)
    If Not pbNoShow Then
        mdiMain.lblSheetNoName.Caption = "��ǰ���ݺ�:"
        mdiMain.lblSheetNo.Caption = g_szSheetID
    End If
    
    oReg.SaveSetting szConnectName, "CurrentSheetID", g_szSheetID
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub


Public Function FormatSheetID(pszCheckID As String)
    FormatSheetID = Format(IIf(pszCheckID <> "", pszCheckID, 0), String(g_oPackageParam.SheetIDNumberLen, "0"))
End Function

Public Sub GetAppSetting()
    Dim szLastTicketNo As String
    
    Dim szLastSheetID As String
    On Error GoTo Here
    szLastSheetID = g_oPackageSvr.GetLastSheetID(g_oActUser.UserID)
    
    g_szSheetID = FormatSheetID(szLastSheetID)
    
    IncSheetID , True
    
    Exit Sub
Here:
    ShowErrorMsg
    
    
End Sub

'���ŸĵĶ�ע�����ĵ���
Public Sub GetAppSetting2()

    Dim szLastSheetID As String
    Dim oReg As New CFreeReg
    Dim szConnectName As String

    
    oReg.init "RTStation", HKEY_LOCAL_MACHINE, "Software\RTSoft"
    
    szConnectName = "Luggage"
    
    On Error GoTo Here

    szLastSheetID = oReg.GetSetting(szConnectName, "CurrentSheetID", "")
    
    g_szSheetID = FormatSheetID(szLastSheetID) - 1

    IncSheetID , True
    
    Exit Sub
Here:
    ShowErrorMsg
End Sub

'��ӡ�а�����
Public Sub PrintAcceptSheet(poAcceptSheet As Package)
           
#If PRINT_SHEET <> 0 Then
    
'    m_oPrintTicket.ClearAll
'    m_oPrintTicket.ReadFormatFileA App.Path & "\AcceptSheet.bpf"
    
    m_oPrintTicket.ClosePort
    m_oPrintTicket.OpenPort
    
    
    '��Ʊ����|2
    m_oPrintTicket.SetObject PPI_SheetID
    m_oPrintTicket.SetCaption poAcceptSheet.SheetID
    
    '����վ|9
    m_oPrintTicket.SetObject PPI_StartStation
    m_oPrintTicket.SetCaption IIf(poAcceptSheet.StartStationName <> "", poAcceptSheet.StartStationName, poAcceptSheet.AreaType)
    
    'Ʒ��|3
    m_oPrintTicket.SetObject PPI_PackageName
    m_oPrintTicket.SetCaption poAcceptSheet.PackageName
    
    '����|11
    m_oPrintTicket.SetObject PPI_PackageNumber
    m_oPrintTicket.SetCaption poAcceptSheet.PackageNumber
    
    '����|7
    m_oPrintTicket.SetObject PPI_CalWeight
    m_oPrintTicket.SetCaption poAcceptSheet.CalWeight
  
    '�ռ���|17
    m_oPrintTicket.SetObject PPI_Picker
    m_oPrintTicket.SetCaption poAcceptSheet.Picker
    
    '����|1
    m_oPrintTicket.SetObject PPI_PackageID
    m_oPrintTicket.SetCaption poAcceptSheet.PackageID
    
    'װж��|40
    m_oPrintTicket.SetObject PPI_PriceItem1
    m_oPrintTicket.SetCaption poAcceptSheet.LoadCharge
    
    '���ܷ�|41
    m_oPrintTicket.SetObject PPI_PriceItem4
    m_oPrintTicket.SetCaption poAcceptSheet.KeepCharge
    
    '�����(���˷�)|43
    m_oPrintTicket.SetObject PPI_PriceItem2
    m_oPrintTicket.SetCaption poAcceptSheet.MoveCharge
    
    '�����˷�|27
    m_oPrintTicket.SetObject PPI_TransCharge
    m_oPrintTicket.SetCaption poAcceptSheet.TransitCharge
    
    '�ϼ�Сд|29
    m_oPrintTicket.SetObject PPI_TotalPrice
    m_oPrintTicket.SetCaption poAcceptSheet.LoadCharge + poAcceptSheet.KeepCharge + poAcceptSheet.MoveCharge + poAcceptSheet.TransitCharge
    
    '�ϼƴ�д|28
    m_oPrintTicket.SetObject PPI_TotalPriceBig
    m_oPrintTicket.SetCaption GetNumber(poAcceptSheet.LoadCharge + poAcceptSheet.KeepCharge + poAcceptSheet.MoveCharge + poAcceptSheet.TransitCharge)
    
    '����|25
    m_oPrintTicket.SetObject PPI_UserName
    m_oPrintTicket.SetCaption poAcceptSheet.UserID
    
    '�������|22
    m_oPrintTicket.SetObject PPI_PickTime
    m_oPrintTicket.SetCaption Format(poAcceptSheet.PickTime, "YYYY-MM-DD HH:mm")
    
    '�����Ǹ�������
    
    '����2|70
    m_oPrintTicket.SetObject PPI_PackageID2
    m_oPrintTicket.SetCaption poAcceptSheet.PackageID
    
    '����2|33
    m_oPrintTicket.SetObject PPI_PackageNumber2
    m_oPrintTicket.SetCaption poAcceptSheet.PackageNumber
    
    '���֤��|21
    m_oPrintTicket.SetObject PPI_PickerCredit
    If poAcceptSheet.PickerCreditID <> "" Then
        m_oPrintTicket.SetCaption Left(poAcceptSheet.PickerCreditID, Len(poAcceptSheet.PickerCreditID) - 4) & "****"
    Else
        m_oPrintTicket.SetCaption ""
    End If
    
    '�ϼ�Сд2|35
    m_oPrintTicket.SetObject PPI_TotalPrice2
    m_oPrintTicket.SetCaption poAcceptSheet.LoadCharge + poAcceptSheet.KeepCharge + poAcceptSheet.MoveCharge + poAcceptSheet.TransitCharge
    
    '����2|94
    m_oPrintTicket.SetObject PPI_UserName2
    m_oPrintTicket.SetCaption poAcceptSheet.UserID
    
    '�������2|93
    m_oPrintTicket.SetObject PPI_PickTime2
    m_oPrintTicket.SetCaption Format(poAcceptSheet.PickTime, "YYYY-MM-DD HH:mm")
  
    
    DoEvents
    m_oPrintTicket.PrintFile

    m_oPrintTicket.ClosePort
    
#End If
End Sub



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
    
    Dim nlen As Integer
    Const cszO = "��"
    
    nlen = ArrayLength(paszNum)
    If pnNum > nlen Then
        GetSplitNum = cszO
    Else
        If paszNum(pnNum) = "" Then
            GetSplitNum = cszO
        
        Else
            GetSplitNum = paszNum(pnNum)
        End If
    End If
    
End Function

Public Sub SetFlex(vsItem As VSFlexGrid, Optional pnRows As Integer = -1, Optional pnCols As Integer = -1)
    If pnRows <> -1 Then
        vsItem.Rows = pnRows
        If pnRows > 0 Then vsItem.FixedRows = 1
    End If
    If pnCols <> -1 Then
        vsItem.Cols = pnCols
        If pnCols > 0 Then vsItem.FixedCols = 1
    End If
 
End Sub

Public Sub SaveRecentSeller(pvaUser As Variant)
    Dim oReg As New CFreeReg
    oReg.init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    Dim nSellerCount As Integer
    Dim i As Integer
    Dim szRecentSeller As String
    nSellerCount = ArrayLength(pvaUser)
    If nSellerCount > 0 Then
        szRecentSeller = pvaUser(1)
        For i = 2 To nSellerCount
            szRecentSeller = szRecentSeller & "," & pvaUser(i)
        Next
        oReg.SaveSetting cszLuggageAccount, cszRecentSeller, szRecentSeller
    End If
End Sub

Public Function GetRecentSeller() As String
    Dim oReg As New CFreeReg
    oReg.init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    GetRecentSeller = oReg.GetSetting(cszLuggageAccount, cszRecentSeller)
End Function


Public Sub FillSellStation(cboSellStation As ComboBox)
    Dim oSystemMan As New SystemMan
    Dim atTemp() As TDepartmentInfo
    Dim i As Integer
    On Error GoTo Here
    '�ж��û������ĸ��ϳ�վ,���Ϊ�������һ������,��������е��ϳ�վ
    oSystemMan.init g_oActUser
    atTemp = oSystemMan.GetAllSellStation(g_oActUser.UserUnitID)
    If g_oActUser.SellStationID = "" Then
        cboSellStation.AddItem ""
        For i = 1 To ArrayLength(atTemp)
            cboSellStation.AddItem MakeDisplayString(atTemp(i).szSellStationID, atTemp(i).szSellStationName)
        Next i
    '����ֻ����û��������ϳ�վ
    Else
        For i = 1 To ArrayLength(atTemp)
            If g_oActUser.SellStationID = atTemp(i).szSellStationID Then
               cboSellStation.AddItem MakeDisplayString(atTemp(i).szSellStationID, atTemp(i).szSellStationName)
               Exit For
            End If
        Next i
        cboSellStation.ListIndex = 0
    End If
    Exit Sub
Here:
    ShowErrorMsg
End Sub
Public Function BuildPacketID(plInputID As Long, Optional pszYear As String, Optional pszMonth As String) As Long
    If pszYear = "" Then
        pszYear = Format(Date, "yy")
    End If
    If pszMonth = "" Then
        pszMonth = Format(Date, "MM")
    End If
    BuildPacketID = Val(pszYear & pszMonth & Format(plInputID, "000000"))
End Function
Public Function UnBuildPacketID(plPacketID As Long) As Long
    UnBuildPacketID = Val(plPacketID Mod 10 ^ 6)
    
End Function

Public Sub InitSMS()
'    Dim oReg As New CFreeReg
'    oReg.init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany       'HKEY_LOCAL_MACHINE
'    '1�Ƚ�Ĭ��ֵ����
'
'
'    g_szSMSServer = oReg.GetSetting("Package", "SMSServer")
'    g_szSMSUser = oReg.GetSetting("Package", "SMSUser")
'    g_szSMSPassword = oReg.GetSetting("Package", "SMSPwd")
'    g_szSMSApiPort = Val(oReg.GetSetting("Package", "SMSApiPort"))
'
'    If init(g_szSMSServer, g_szSMSUser, g_szSMSPassword, g_szSMSApiPort) = 0 Then
'        g_bSMSValid = True
'    Else
'        g_bSMSValid = False
'    End If

    If init("10.20.20.20", "xb", "xb123", 11) = 0 Then
        g_bSMSValid = True
    Else
        g_bSMSValid = False
    End If
End Sub
Public Sub SendSMS(pszPhone As String, pszMessage As String)
'    If sendSM(pszPhone, pszMessage, Val(g_szSMSApiPort)) = 0 Then
'        MsgBox "���ͳɹ�!", vbInformation, "��ʾ"
'    Else
'        MsgBox "����ʧ��!", vbExclamation, "����"
'    End If
        
    If sendSM(pszPhone, pszMessage, 11) = 0 Then
        MsgBox "���ͳɹ�!", vbInformation, "��ʾ"
    Else
        MsgBox "����ʧ��!", vbExclamation, "����"
    End If
End Sub
Public Sub ReleaseSMS()
    If g_bSMSValid Then
        release
    End If
End Sub

'ˢ�½����ϵĵ��ݺ�
Public Sub RefreshCurrentSheetID()
    Dim szLastSheetID As String
    Dim oReg As New CFreeReg
    Dim szConnectName As String
    
    
    oReg.init "RTStation", HKEY_LOCAL_MACHINE, "Software\RTSoft"
    
    szConnectName = "Luggage"
    
    On Error GoTo Here
    szLastSheetID = oReg.GetSetting(szConnectName, "CurrentSheetID", "")
    
    g_szSheetID = FormatSheetID(szLastSheetID)
    
    mdiMain.lblSheetNoName.Caption = "��ǰ���ݺ�:"
    mdiMain.lblSheetNo.Caption = g_szSheetID
 
    oReg.SaveSetting szConnectName, "CurrentSheetID", g_szSheetID
       
    Exit Sub
Here:
    ShowErrorMsg
End Sub
