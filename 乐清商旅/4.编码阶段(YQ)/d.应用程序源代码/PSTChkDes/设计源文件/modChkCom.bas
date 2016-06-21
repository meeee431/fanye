Attribute VB_Name = "mdlMain"

Option Explicit
'====================================================================
'���¶��峣��
'--------------------------------------------------------------------
Public Const MaxLine = 5                       'ͬʱ���е�����Ʊ������
Public Const g_cszTitle_Info = "��ʾ"
Public Const g_cszTitle_Warning = "����"
Public Const g_cszTitle_Error = "����"
Public Const g_cszTitle_Question = "����"
Public Const g_cszTitleScollBus = "��ˮ����"

Public m_rsTicketType As Recordset
Public m_oActiveUser As ActiveUser
'--------------------------------------------------------------------
'���¶���ö��
'--------------------------------------------------------------------
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
End Enum

Public Enum ECheckStatus
    ECS_CanotCheck = 1      '���ܼ�
    ECS_CanCheck = 2        '�ܼ�Ʊ
    ECS_CanExtraCheck = 3   '�ܲ���
    ECS_BeChecking = 4      '���ڼ�Ʊ
    ECS_BeExtraChecking = 5 '���ڲ���
    ECS_Checked = 6         '�Ѽ�
End Enum

Public Enum ETabAct             'TabStrip����Ϊ
    Addone = 1
    CloseOne = 2
End Enum

Public Enum eEventId
    AddBus = 1              '��ӳ���
    AdjustTime = 2          '����ʱ��
    ChangeBusCheckGate = 3  '���ļ�Ʊ��
    ChangeBusSeat = 4       '������λ
    ChangeBusStandCount = 5 'ĳ���ε�վƱ���ı�
    ChangeBusTime = 6       '���ĳ��η���ʱ��
    ChangeParam = 7         '���Ĳ���
    ExStartCheckBus = 8     '���쳵��
    MergeBus = 9            '���β���
    RemoveBus = 10          'ɾ������
    ResumeBus = 11          '���θ���
    StartCheckBus = 12      '���쳵��
    StopBus = 13            '����ͣ��
    StopCheckBus = 14       'ͣ�쳵��
End Enum

Type TEventSound                        '�¼���Ч�ļ���·����Ϣ�ṹ
    InvalidTicket As String             '��ЧƱ
    CanceledTicket As String            'Ʊ�ѷ�
    ReturnedTicket As String            'Ʊ����
    NoMatchedBus As String              '�ǵ��쳵��
    CheckedTicket As String             '�Ѽ�Ʊ
    CheckSucess As String               '��Ʊ�ɹ�
    CheckTimeOn As String               '��Ʊʱ���ѵ�
    StartupCheckTimeOn As String        '����ʱ���ѵ�
    FreeTicket As String                '��Ʊ��ʾ
    HalfTicket As String                '��Ʊ��ʾ
    PreferentialTicket1 As String       '�Ż�Ʊ1��ʾ
    PreferentialTicket2 As String       '�Ż�Ʊ2��ʾ
    PreferentialTicket3 As String       '�Ż�Ʊ3��ʾ
'    sndChanged As String                '������ǩ
'    sndNormal As String                 '��������
'    sndHasBeChanged As String           '����ǩ
End Type


Type tCheckInfo
    CheckDate  As Date                              '��Ʊ����
    '��Ʊ����Ϣ
    CheckGateNo As String                           '��ǰ��Ʊ��
    SellStationID As String                         '��ǰ�ϳ�վ����
    SellStationName As String                       '��ǰ�ϳ�վ����
    AutoPrint As Boolean                            '�Ƿ�ͣ���ֱ�Ӵ�ӡ·��
    CheckGateName As String                         '��ǰ��Ʊ��
    CheckerId As String                             '��Ʊ��Id
    Checker As String                               '��Ʊ��
    CurrSheetNo As String                           '��ǰ·��
    
    '��Ʊ������Ϣ
    BusID As String                                 '���κ�
    EndStationName As String                        '�յ�վ
    StartUpTime As Date                             '����ʱ��
    StartCheckTime As Date                          '����ʱ��
    StopCheckTime As Date                           'ͣ��ʱ��
    BusMode As EBusType                             '����״̬
    SellTickets As Integer                          '��Ʊ��
    SelfSellStationTickets As Integer            '�û������ϳ�վ��Ʊ��
    SeatCount As Integer                            '��λ��
    Owner As String                                 '����
    Company As String                               '��Ӫ��˾
    MergedBus As String                             '���복��
    MergeType As Integer
    SplitSeat As Integer
    MergeInSells As Integer                         '���복����Ʊ��
    VehicleId As String
    Vehicle As String                               '��������
    VehicleMode As String                           '��������
    SerialNo As Integer                             '�������
    CheckSheet As String                            '����·����
    '���г�����Ϣ
    RunVehicle As M_TRunVehicle                     '��ǰ���г�����Ϣ
End Type
Type TTicketInfo
    TicketID As String                              '��Ʊ��
    EndStation As String                            '�յ�վ
    TicketStatus As ETicketStatus                   '��Ʊ״̬
    g_tTicketType As ETicketType                       '��Ʊ����
    TicketDate As Date                              '��Ʊ����
End Type

Type TCheckLineFormInfo
'��Ʊ���̱��Ļ�����Ϣ�����ڶԵ�ǰͬʱִ�еĶ����Ʊ���̽��м��
    BusID As String
    ExCheck As Boolean
    SerialNo As Integer
End Type

Type TWillStopBusStack
    'ͣ�쳵�ζ�ջ
    Top As Integer
    MsgStyle(1 To MaxLine) As Integer
    ChkLine(1 To MaxLine) As Integer
End Type



'--------------------------------------------------------------------
'���¶���ȫ�ֱ���
'--------------------------------------------------------------------
Public g_oActiveUser As ActiveUser      '��ǰ��û�
Public g_oChkTicket As CheckTicket   '��ǰ��Ʊ����
Public g_oEnvBus As REBus         '��ǰ��������
Public g_tCheckInfo As tCheckInfo   '��ǰϵͳ�ļ�Ʊ�����
Public g_tEventSoundPath As TEventSound                 '�¼���Ч�ļ�·��
Public g_cWillCheckBusList As BusCollection            '����ĳ����б���
Public g_cCheckedBusList As BusCollection           '����ĳ����б���
Public g_atCheckLine(1 To MaxLine) As TCheckLineFormInfo
Public g_aofrmCheckForm(1 To MaxLine) As frmCheckTicket

    'ϵͳ����ȫ�ֱ���
Public g_nLatestExtraCheckTime As Integer
Public g_nBeginCheckTime As Integer
Public g_nExtraCheckTime As Integer
Public g_nCheckTicketTime As Integer
Public g_bAllowChangeRide As Boolean
Public g_szUnitID As String
Public g_szUnitName As String
Public g_tTicketType() As TTicketType
Public g_nCheckSheetLen As Integer
Public g_nCurrLineIndex As Integer                      '��ǰ������һ����Ʊ���ν���

Public g_bAllowStartChectNotRearchTime As Boolean '�Ƿ�����δ������ʱ�俪��


Public g_szSellStationName As String
'--------------------------------------------------------------------
'���¶��屾ģ�����
'--------------------------------------------------------------------


'��Ʊ����ģ��


'���³�������
'*************************************************************
'*      �˴�����һЩ��ʱ����
'Public Const AheadTime = 0.0069                'Ԥ�������ǰ����ʱ��
'Public Const cntCheckTime = 10                 'Ԥ����ļ�Ʊʱ��
Public Const m_cRegSystemKey = "ChkDes\CheckEnviroment"         '��Ʊϵͳ�����ַ���
Public Const m_cRegSoundKey = "ChkDes\EventSound"      '��Ʊ��Ч�����ַ���
Public Const m_cnTimeWindage = 1            'ʱ��ƫ����������Э����һ�೵�ε�ʱ�䣨�Է���Ϊ��λ��

Public szSeatBusID As String

'***********************************************************

'����ö�ٶ���


'Public m_sgAheadTime As Single                 'Ԥ�������ǰ����ʱ��(��СʱΪ��λ)
'Public m_sgCheckTime As Single                  'Ԥ����ļ�Ʊʱ��(�Է���Ϊ��λ)
Public m_szPrnFmtFile As String                  '·����ӡ��ʽ�ļ���·��

Public m_dtAheadTime As Date                            '��Ʊ����ǰʱ��
Public g_oNextEnvBus As REBus



'Public m_nPrevLineIndex As Integer                      'ǰһ��ʹ�õļ�Ʊ���ν���
Public g_szTitle As String
Public m_bIsFormActive As Boolean                       '�Ƿ����ȼ����˴���
Public m_bCloseOne As Boolean                           '�Ƿ�ر��˴���
Public m_lErrorCode As Long                             '�����

'��������һЩ��ʹ�õ�API����
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'��������API�����õ��ĳ�������
Public Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Public Const SND_FILENAME = &H20000     '  name is a file name
Public Const SND_ASYNC = &H1         '  play asynchronously

Public Const HWND_TOPMOST = -1

'�õ���ǰ�ļ�Ʊ������
Public Function CheckLineCount() As Integer
    If MDIMain.tbsBusList.Visible Then
        CheckLineCount = MDIMain.tbsBusList.Tabs.Count
    Else
        CheckLineCount = 0
    End If
End Function
'����µļ�Ʊ����
Public Sub AddNewCheckLine(BusID As String, Optional ExCheck As Boolean = False, Optional IsChecking As Boolean = False, Optional szVehicleId As String, Optional oREBus As REBus)
'    ' ����ע��
'    ' *************************************
'    ' BusID:���δ���
'    ' ExCheck:��ѡ���Ƿ񲹼�
'    ' IsChecking:��ѡ���Ƿ����ڼ�Ʊ������Ҫ���ڴ���ϵͳ�쳣�ж�ʱ��ָ����ã�
'    ' szVehicleId:��ѡ���������ε����г�����
'    ' *************************************
On Error GoTo ErrHandle
    Dim nPrevLineIndex As Integer   'm_nCurrLineIdex�ĳ�ֵ
    Dim nCheckLineCount As Integer
    Dim nErrorNum As Long
    
    nCheckLineCount = CheckLineCount
    If nCheckLineCount + 1 > MaxLine Then
        MsgboxEx "��ϵͳ���֧��ͬʱ����" & Str(MaxLine) & "����Ʊ���̣�����", vbInformation, g_cszTitle_Info
        Exit Sub
    End If
    nPrevLineIndex = g_nCurrLineIndex
    g_nCurrLineIndex = nCheckLineCount + 1
    
    
    ShowSBInfo "���ڳ�ʼ����Ʊ����..."
    
    ResetEnvBusInfo BusID, nErrorNum, oREBus
    If nErrorNum <> 0 Then GoTo ErrHandle
    g_atCheckLine(g_nCurrLineIndex).BusID = g_tCheckInfo.BusID
    g_atCheckLine(g_nCurrLineIndex).ExCheck = ExCheck
    g_atCheckLine(g_nCurrLineIndex).SerialNo = g_tCheckInfo.SerialNo
    
    If Not IsChecking Then
        '���µ����м�㷽�����죬�����ĵ���ó��εĳ���״̬
        If ExCheck Then
            g_oChkTicket.ExtraStartCheckBus g_tCheckInfo.BusID, g_tCheckInfo.SerialNo
        Else
            If g_tCheckInfo.BusMode = TP_ScrollBus Then
                g_oChkTicket.StartCheckScrollBus g_tCheckInfo.BusID, g_tCheckInfo.SerialNo, szVehicleId
            Else
                If szVehicleId = "" Then
                    g_oChkTicket.StartCheckRegularBus g_tCheckInfo.BusID
                Else
                    g_oChkTicket.StartCheckRegularBus g_tCheckInfo.BusID, szVehicleId
                End If
            End If
        End If
        
        Dim nIndex As Integer           '���Ļ���������״̬
        Dim tTmpBusInfo As tCheckBusLstInfo
        nIndex = g_cWillCheckBusList.FindItem(g_tCheckInfo.BusID)
        If nIndex > 0 Then
            tTmpBusInfo = g_cWillCheckBusList.Item(nIndex)
            tTmpBusInfo.Status = IIf(ExCheck, EREBusStatus.ST_BusExtraChecking, EREBusStatus.ST_BusChecking)
            g_cWillCheckBusList.UpdateOne tTmpBusInfo
            If frmBusList.IsShow Then frmBusList.UpdateWillCheckBusItem 2, tTmpBusInfo      '����һ��
        End If
    End If
    
    nCheckLineCount = nCheckLineCount + 1
    Dim szTabsString As String
    szTabsString = g_tCheckInfo.BusID & IIf(g_tCheckInfo.BusMode = TP_ScrollBus, "-" & g_tCheckInfo.SerialNo, "") & _
                                             g_tCheckInfo.EndStationName & "(&" & nCheckLineCount & ")"
    With MDIMain                   '���ü�Ʊ���ν��̱�ǩ
    If .tbsBusList.Visible = False Then
        .tbsBusList.Tabs(1).Caption = szTabsString
        .tbsBusList.Visible = True
    Else
        .tbsBusList.Tabs.Add g_nCurrLineIndex, , szTabsString
    End If
    End With
    
    '��ʾ��Ʊ����
    Dim ofrmCheckTicket As New frmCheckTicket
    Set g_aofrmCheckForm(g_nCurrLineIndex) = ofrmCheckTicket
    Set g_aofrmCheckForm(g_nCurrLineIndex).m_oREBus = g_oEnvBus
    g_aofrmCheckForm(g_nCurrLineIndex).Show
    
    MDIMain.tbsBusList.Tabs(g_nCurrLineIndex).Selected = True
    ShowSBInfo ""
    Exit Sub
ErrHandle:
    If err.Number = ERR_ChkTkBusAlreadyExist Then
         '��Ʊ�����Ѵ���ʱ����ִ�У���Ҫ���ڷ������˳���Ʊʱ�ָ�
        Resume Next
    Else
        ShowErrorMsg
    End If
    g_nCurrLineIndex = nPrevLineIndex   '���س�ʼֵ
    ShowSBInfo ""
End Sub
'�ر�ĳһ���ļ�Ʊ����
Public Sub CloseOneCheckLine(nWhichOne As Integer)
    Dim i As Integer
    Dim nCheckLineCount As Integer
    nCheckLineCount = CheckLineCount
'    For i = g_nCurrLineIndex To nCheckLineCount - 1
    For i = nWhichOne To nCheckLineCount - 1
        g_atCheckLine(i).BusID = g_atCheckLine(i + 1).BusID
        g_atCheckLine(i).ExCheck = g_atCheckLine(i + 1).ExCheck
        Set g_aofrmCheckForm(i) = g_aofrmCheckForm(i + 1)
        MDIMain.tbsBusList.Tabs(i).Caption = _
            Left(MDIMain.tbsBusList.Tabs(i + 1).Caption, _
            Len(MDIMain.tbsBusList.Tabs(i + 1).Caption) - 4) _
            & "(&" & Trim(Str(i)) & ")"
        g_aofrmCheckForm(i).Tag = Str(i)
    Next i
    Set g_aofrmCheckForm(nCheckLineCount) = Nothing
    If nCheckLineCount = 1 Then
        g_nCurrLineIndex = 0
        MDIMain.tbsBusList.Visible = False
'        MDIMain.mnu_Query_Ticket.Enabled = False
'        MDIMain.mnu_Check_Bus.Enabled = False
    Else
        If g_nCurrLineIndex >= nWhichOne And g_nCurrLineIndex <> 1 Then g_nCurrLineIndex = g_nCurrLineIndex - 1
        MDIMain.tbsBusList.Tabs.Remove nCheckLineCount
        MDIMain.tbsBusList.Tabs.Item(g_nCurrLineIndex).Selected = True
    End If
End Sub
'���ݳ��κŻ�ȡ���³�����Ϣ������ϵͳ�����g_tCheckInfo�У����д��󽫴��󷵻�
Public Sub ResetEnvBusInfo(szBusid As String, Optional ByRef ErrorCode As Long, Optional oREBus As REBus)
'    ' ����ע��
'    ' *************************************
'    ' szBusID:����Id
'    ' ErrorCode:�����,��ѡ
'    ' *************************************
'    ' ****************************************************************
'    ' �ѽ������g_tCheckInfo������У����д��󽫴��󷵻�
'    ' ****************************************************************

On Error GoTo ErrHandle
    If oREBus Is Nothing Then
        g_oEnvBus.Identify szBusid, Date, g_tCheckInfo.CheckGateNo
    Else
        Set g_oEnvBus = oREBus
    End If
    g_tCheckInfo.BusID = UCase(Trim(g_oEnvBus.BusID))
    g_tCheckInfo.BusMode = g_oEnvBus.BusType
    g_tCheckInfo.Company = g_oEnvBus.CompanyName
    g_tCheckInfo.Owner = g_oEnvBus.OwnerName
    g_tCheckInfo.EndStationName = g_oEnvBus.EndStationName
    If g_oEnvBus.BusType <> TP_ScrollBus Then
            g_tCheckInfo.MergedBus = Trim(g_oEnvBus.BeMergedBus.szBusid)
            g_tCheckInfo.MergeType = g_oEnvBus.BeMergedBus.nMergeType
    End If
'    If g_oEnvBus.BusType = TP_ScrollBus Then
        
'    Else
    g_tCheckInfo.SeatCount = g_oEnvBus.TotalSeat
'    End If
    g_tCheckInfo.StartUpTime = g_oEnvBus.StartUpTime
    g_tCheckInfo.StartCheckTime = g_oEnvBus.StartCheckTime
    g_tCheckInfo.StopCheckTime = g_oEnvBus.StopCheckTime
    g_tCheckInfo.CheckSheet = Trim(g_oEnvBus.CheckSheet)
    g_tCheckInfo.VehicleId = g_oEnvBus.Vehicle
    g_tCheckInfo.Vehicle = g_oEnvBus.VehicleTag
    g_tCheckInfo.VehicleMode = g_oEnvBus.VehicleModelName
    ErrorCode = 0
    Exit Sub
ErrHandle:
    ShowErrorMsg
    ErrorCode = err.Number
End Sub
Public Sub PlayEventSound(szFileName As String)
    '������Ч
    PlaySound szFileName, 0, SND_FILENAME + SND_ASYNC
    
End Sub
'Public Sub ShowErrorMsg()
'    MsgboxEx err.Description, vbExclamation, "����-" & err.Number
'End Sub
Public Sub Main()
    Dim oCommShell As CommShell
    
    If App.PrevInstance Then
        MsgBox "ϵͳ������!", vbExclamation, "����"
        End
    End If
On Error GoTo ErrHandle
    Set oCommShell = New CommShell

TryLogin:
    Set g_oActiveUser = oCommShell.ShowLogin()
    If g_oActiveUser Is Nothing Then Exit Sub
    
    
'    m_szPrnFmtFile = App.Path & "\ChkSheet.bpf"
'    App.HelpFile = SetHTMLHelpStrings("SNChkSys.chm") '�趨App.HelpFile
    
    InitSystemParam
    GetIniFile
    
    If g_tCheckInfo.CheckGateNo = "" Then
        MsgBox "��һ��ʹ�ü�Ʊ̨,��ָ����ǰ��Ʊ��!", vbInformation, g_cszTitle_Info
        frmSetOption.Show vbModal
    End If
    If g_tCheckInfo.CheckGateNo <> "" Then
        frmChangeSheetNo.FirstLoad = True
        frmChangeSheetNo.Show vbModal
    End If
    
    
    '����������
'    oCommShell.ShowSplash "��Ʊϵͳ", "Check Ticket Desktop", LoadResPicture(101, 0), App.Major, App.Minor, App.Revision
    DoEvents
    
    '�õ��ϳ�վ����
    Dim oCheckGate As New CheckGate
    oCheckGate.Init g_oActiveUser
    oCheckGate.Identify g_tCheckInfo.CheckGateNo
    g_tCheckInfo.SellStationID = oCheckGate.SellStationID
    g_tCheckInfo.SellStationName = oCheckGate.SellStationName
    
    g_szSellStationName = oCheckGate.SellStationName
    
    '��ʼ��ȫ�ֱ���
    g_oChkTicket.CheckGateNo = g_tCheckInfo.CheckGateNo
    Set g_oEnvBus = New REBus
    g_oEnvBus.Init g_oActiveUser
    SetHTMLHelpStrings "STChkDes.chm"
    
'    Dim oPriceMan As New STPrice.TicketPriceMan
'    oPriceMan.Init m_oActiveUser
    
    
    
    Load MDIMain
    EvisibleCloseButton MDIMain
    MDIMain.Show
    oCommShell.CloseSplash
    Exit Sub
ErrHandle:
    If err.Number = 500 Then        '��ע�����
        SaveRegInitInfo
        Resume
    Else
        ShowErrorMsg
        Resume TryLogin
    End If
End Sub
'�رյ�ǰ������ʾ��Modal����,����Tag��ʶΪModal��
Public Sub CloseModalForm()
    '����Modal���ڵ�tag����
On Error GoTo ErrHandle
    Do
        If Screen.ActiveForm.Tag <> "Modal" Then
            Exit Do
        End If
        If Screen.ActiveForm Is frmCheckSheet Then
            Exit Do
        Else
            Unload Screen.ActiveForm    '·����ӡ���漤��ʱ����Ҫ�رգ�����·���Ż�����
        End If
    Loop
    Exit Sub
ErrHandle:
    On Error Resume Next
End Sub
Public Sub WriteCheckGateInfo()
'�õ���Ʊ��״̬,��д�����
    With MDIMain
        .lblChecker.Caption = g_tCheckInfo.Checker
        .lblCheckGate.Caption = g_tCheckInfo.CheckGateName
        .lblCurrentSheetNo.Caption = g_tCheckInfo.CurrSheetNo
        .moMessage.SellStation = g_tCheckInfo.SellStationID
    End With
End Sub
Public Sub WriteNextBus()
'�õ����쳵��,��д�����
    Dim lHaveTime As Double
    Dim dtTmp As Date
    Dim dtStartUpTime As Date
    Dim dtStopCheckTime As Date
    
    On Error Resume Next
    
    With MDIMain
        Set g_oNextEnvBus = g_oChkTicket.GetNextCheckBus
        
        If Not (g_oNextEnvBus Is Nothing) Then
            dtStartUpTime = g_oNextEnvBus.StartUpTime
            dtStopCheckTime = g_oNextEnvBus.StopCheckTime
            .lblBusID.Caption = g_oNextEnvBus.BusID
            .lblStartupTime.Caption = Format(dtStartUpTime, "HH:MM:SS")
            .lblEndStation.Caption = g_oNextEnvBus.EndStationName
            .lblCompany.Caption = g_oNextEnvBus.CompanyName
            .lblOwner.Caption = g_oNextEnvBus.OwnerName
            .lblLicense.Caption = g_oNextEnvBus.Vehicle
            '********Need change******
            dtTmp = Now
            lHaveTime = DateDiff("s", dtTmp, DateAdd("n", -g_nBeginCheckTime, dtStartUpTime))
            
            lHaveTime = IIf(lHaveTime > 0, lHaveTime, 0)
            .rvtTime.Second = lHaveTime
            EnabledMDITimer True
            
            '��RevTimer1����Ч������ʵʱˢ����һ���쳵�Σ�����û�п���ĳ��Σ�
            '����ʱ��Ϊ��һ���쳵�η���ʱ����m_cnTimeWindage����
            lHaveTime = DateDiff("s", dtTmp, _
                dtStartUpTime)
            lHaveTime = lHaveTime + 60 * m_cnTimeWindage
            lHaveTime = IIf(lHaveTime > 0, lHaveTime, 0)
            .RevTimer1.Second = lHaveTime
            .RevTimer1.Enabled = True
        Else
            .lblBusID.Caption = ""
            .lblStartupTime.Caption = ""
            .lblEndStation.Caption = ""
            .lblCompany.Caption = ""
            .lblOwner.Caption = ""
            .lblLicense.Caption = ""
            .rvtTime.Second = 0
'            .rvtTime.Enabled = False
            EnabledMDITimer False
        End If
    End With
    Exit Sub
End Sub
Public Sub WriteInitReg()
'���浱ǰ·������ע���
    Dim oReg As New CFreeReg
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    oReg.SaveSetting m_cRegSystemKey, "SheetNo", g_tCheckInfo.CurrSheetNo
    Set oReg = Nothing
End Sub


Public Sub EnabledMDITimer(bEnabled As Boolean)
'����������Ĵ��쳵�ε���ʱ��
    If bEnabled Then
        MDIMain.flblrevTime.Visible = False
        MDIMain.rvtTime.Enabled = True
        MDIMain.rvtTime.Visible = True
    Else
        MDIMain.flblrevTime.Visible = True
        MDIMain.rvtTime.Enabled = False
        MDIMain.rvtTime.Visible = False
    End If
End Sub
Public Function MsgboxEx(Optional Prompt As String = "", Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String = "", Optional showMode As FormShowConstants = vbModal) As VbMsgBoxResult
'�Զ���msgbox����
    Dim ofrm As New frmMsgbox
    ofrm.Prompt = Prompt
    ofrm.Buttons = Buttons
    ofrm.Title = Title
    ofrm.Show showMode
    MsgboxEx = ofrm.Result
'    MsgboxEx = meMsgboxEx
End Function

'��Ϣ�������                              *

Public Sub RunMsgEvent(EventMode As eEventId, EventParam() As String)
'    ' ����ע��
'    ' *************************************
'    ' EventMode:��Ϣ���ͣ���ϢId��
'    ' EventParam:��������
'    ' *************************************

    Dim nTmp As Integer
    Dim tTmpBusInfo As tCheckBusLstInfo
    Select Case EventMode
        Case eEventId.AddBus
            If Trim(EventParam(3)) = Trim(g_tCheckInfo.CheckGateNo) Then
                BuildBusCollection
                If frmBusList.IsShow Then
                    frmBusList.RefreshBus
                End If
            End If
        Case eEventId.AdjustTime
        Case eEventId.ChangeBusCheckGate
            If Trim(EventParam(3)) = Trim(g_tCheckInfo.CheckGateNo) Then
                BuildBusCollection
            Else
                nTmp = g_cWillCheckBusList.FindItem(EventParam(1))
                If nTmp > 0 Then
                    tTmpBusInfo = g_cWillCheckBusList.Item(nTmp)
                    g_cWillCheckBusList.RemoveOne nTmp
                Else
                    Exit Sub
                End If
            End If
            'ˢ�¼�Ʊ�����б���
            If frmBusList.IsShow Then
                frmBusList.RefreshBus
            End If
            
            'ˢ�¿��촰��
            '******************
            
            'ˢ����һ����
            If Not g_oNextEnvBus Is Nothing Then
                If g_oNextEnvBus.BusID = EventParam(1) Then
                    WriteNextBus
                End If
            End If
        Case eEventId.ChangeBusSeat
                                                        
        Case eEventId.ChangeBusStandCount
        Case eEventId.ChangeBusTime
            nTmp = g_cWillCheckBusList.FindItem(EventParam(1))
            If nTmp > 0 Then
                tTmpBusInfo = g_cWillCheckBusList.Item(nTmp)
                tTmpBusInfo.StartUpTime = Format(EventParam(3), cszDateTimeStr)
                g_cWillCheckBusList.UpdateOne tTmpBusInfo
            Else
                Exit Sub
            End If
            
            'ˢ�¼�Ʊ�����б���
            If frmBusList.IsShow Then
                frmBusList.RefreshBus
            End If
            
            'ˢ�¿��촰��
            '******************
            
            'ˢ����һ����
            If Not g_oNextEnvBus Is Nothing Then
                If g_oNextEnvBus.BusID = EventParam(1) Then
                    WriteNextBus
                End If
            End If
        Case eEventId.ChangeParam
        Case eEventId.MergeBus
            nTmp = g_cWillCheckBusList.FindItem(EventParam(1))
            If nTmp > 0 Then
                tTmpBusInfo = g_cWillCheckBusList.Item(nTmp)
                tTmpBusInfo.Status = EREBusStatus.ST_BusSlitpStop
                g_cWillCheckBusList.UpdateOne tTmpBusInfo
            Else
                Exit Sub
            End If
            
            'ˢ�¼�Ʊ�����б���
            If frmBusList.IsShow Then
                frmBusList.RefreshBus
            End If
            
            'ˢ�¿��촰��
            '******************
            
            'ˢ����һ����
            If Not g_oNextEnvBus Is Nothing Then
                If g_oNextEnvBus.BusID = EventParam(1) Then
                    WriteNextBus
                End If
            End If
        Case eEventId.RemoveBus
            nTmp = g_cWillCheckBusList.FindItem(EventParam(1))
            If nTmp > 0 Then
                tTmpBusInfo = g_cWillCheckBusList.Item(nTmp)
                g_cWillCheckBusList.RemoveOne nTmp
            Else
                Exit Sub
            End If
            
            'ˢ�¼�Ʊ�����б���
            If frmBusList.IsShow Then
                frmBusList.RefreshBus
            End If
            
            'ˢ�¿��촰��
            '******************
            
            'ˢ����һ����
            If Not g_oNextEnvBus Is Nothing Then
                If g_oNextEnvBus.BusID = EventParam(1) Then
                    WriteNextBus
                End If
            End If
        Case eEventId.ResumeBus
            nTmp = g_cWillCheckBusList.FindItem(EventParam(1))
            If nTmp > 0 Then
                tTmpBusInfo = g_cWillCheckBusList.Item(nTmp)
                tTmpBusInfo.Status = EREBusStatus.ST_BusNormal
                g_cWillCheckBusList.UpdateOne tTmpBusInfo
            Else
                Exit Sub
            End If
            
            'ˢ�¼�Ʊ�����б���
            If frmBusList.IsShow Then
                frmBusList.RefreshBus
            End If
            
            'ˢ�¿��촰��
            '******************
            
            'ˢ����һ����
            If Not g_oNextEnvBus Is Nothing Then
                If g_oNextEnvBus.BusID = EventParam(1) Then
                   WriteNextBus
                Else
                   If tTmpBusInfo.StartUpTime < g_oNextEnvBus.StartUpTime Then
                     WriteNextBus
                   End If
                End If
            End If
        Case eEventId.StopBus
            nTmp = g_cWillCheckBusList.FindItem(EventParam(1))
            If nTmp > 0 Then
                tTmpBusInfo = g_cWillCheckBusList.Item(nTmp)
                tTmpBusInfo.Status = EREBusStatus.ST_BusStopped
                g_cWillCheckBusList.UpdateOne tTmpBusInfo
            Else
                Exit Sub
            End If
            
            'ˢ�¼�Ʊ�����б���
            If frmBusList.IsShow Then
                frmBusList.RefreshBus
            End If
            
            'ˢ�¿��촰��
            '******************
            
            'ˢ����һ����
            If Not g_oNextEnvBus Is Nothing Then
                If g_oNextEnvBus.BusID = EventParam(1) Then
                    WriteNextBus
                End If
            End If
    End Select
End Sub
Public Sub BuildBusCollection()
    'ȡ�õ���ļ�Ʊ�����б���Ϣ������g_cWillCheckBusList
    Dim atCheckSheetInfo() As tCheckBusLstInfo
    Dim szBusid As String, szCheckedBusID As String
    Dim nTmpIndex As Integer
    Dim i As Integer, nCount As Integer
    Dim j As Integer, l As Integer, n As Integer
    
    If g_cWillCheckBusList Is Nothing Then Set g_cWillCheckBusList = New BusCollection
    If g_cCheckedBusList Is Nothing Then Set g_cCheckedBusList = New BusCollection
    g_cWillCheckBusList.RemoveAll
    g_cCheckedBusList.RemoveAll
    
    '�õ��������г�����Ϣ
    Dim rsBusInfo As Recordset
    Set rsBusInfo = g_oChkTicket.GetBusInfoRs(Date, g_tCheckInfo.CheckGateNo)
    
    '�õ��Ѽ쳵�μ�·����Ϣ
    Dim rsCheckedBus As Recordset
    Set rsCheckedBus = g_oChkTicket.GetBusCheckSheetRs(Date, g_tCheckInfo.CheckGateNo)
    
    Dim tTmpBusInfo As tCheckBusLstInfo '�����б�Ԫ����ʱ����
    Dim tTmpBusInfo2 As tCheckBusLstInfo '�����б�Ԫ����ʱ����
'    rsBusInfo.MoveFirst: rsCheckedBus.MoveFirst
'    Dim bStack As Boolean       '��True��ʾrsBusInfoδ�ƶ�
    Do While Not rsBusInfo.EOF Or Not rsCheckedBus.EOF
        '�����ж��Ƿ����ظ����Σ���ͬһ�������β�Ϊһ����¼
        If Not rsBusInfo.EOF Then
            szBusid = UCase(Trim(rsBusInfo("bus_id")))
            tTmpBusInfo.BusID = szBusid
            tTmpBusInfo.BusMode = rsBusInfo("bus_type")
            tTmpBusInfo.Company = rsBusInfo("transport_company_short_name")
            tTmpBusInfo.Vehicle = rsBusInfo("license_tag_no")
            tTmpBusInfo.StartUpTime = rsBusInfo("bus_start_time")
            tTmpBusInfo.EndStationName = rsBusInfo("end_station_name")
            tTmpBusInfo.Owner = rsBusInfo("owner_name")
            tTmpBusInfo.Status = rsBusInfo("status")
            tTmpBusInfo.BusSerial = 0
        Else
            szBusid = ""
        End If
        If Not rsCheckedBus.EOF Then
            szCheckedBusID = UCase(Trim(rsCheckedBus("bus_id")))
            tTmpBusInfo2.BusID = szCheckedBusID
            tTmpBusInfo2.BusSerial = rsCheckedBus("bus_serial_no")
            tTmpBusInfo2.BusMode = rsCheckedBus("bus_type")
            tTmpBusInfo2.Company = rsCheckedBus("transport_company_short_name")
            tTmpBusInfo2.Vehicle = rsCheckedBus("license_tag_no")
            tTmpBusInfo2.StartUpTime = rsCheckedBus("bus_start_time")
            tTmpBusInfo2.EndStationName = rsCheckedBus("end_station_name")
            tTmpBusInfo2.StartChkTime = rsCheckedBus("check_start_time")
            tTmpBusInfo2.StopChkTime = rsCheckedBus("check_end_time")
            tTmpBusInfo2.Owner = rsCheckedBus("owner_name")
            tTmpBusInfo2.Status = EREBusStatus.ST_BusStopCheck
            tTmpBusInfo2.CheckSheet = Trim(rsCheckedBus("check_sheet_id"))
        Else
            szCheckedBusID = ""
        End If
        '�������μȷ��ڴ��쳵�μ����У��ַ����Ѽ쳵�μ�����
        '�̶�����Ҫô�ڴ��쳵���У�Ҫô���Ѽ쳵����
        
        If szBusid = szCheckedBusID Then
            If szCheckedBusID <> "" Then
                g_cCheckedBusList.Addone tTmpBusInfo2
            End If
            If Not rsCheckedBus.EOF Then rsCheckedBus.MoveNext
            If tTmpBusInfo.BusMode = TP_ScrollBus Then
                g_cWillCheckBusList.Addone tTmpBusInfo
            End If
            rsBusInfo.MoveNext
        Else
            If szBusid <> "" Then
                g_cWillCheckBusList.Addone tTmpBusInfo
                rsBusInfo.MoveNext
            Else
                g_cCheckedBusList.Addone tTmpBusInfo2
                rsCheckedBus.MoveNext
            End If
        End If
    Loop
End Sub


Public Function GetCodeStr(szSource As String, nlen As Integer) As String
'��Ҫ�󳤶��������ִ�
    Dim szNum As String
    Dim nZeroNum As Integer
    
    szNum = Trim(Str(Val(szSource)))
    nZeroNum = nlen - Len(szNum)
    If nZeroNum < 0 Then nZeroNum = 0
    GetCodeStr = String(nZeroNum, "0") & Left(szNum, nlen - nZeroNum)
End Function
'����ע�����ʼ��Ϣ
Private Sub SaveRegInitInfo()
On Error GoTo ErrHandle
    Dim oReg As New CFreeReg
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    oReg.SaveSetting m_cRegSystemKey, "CheckGate", ""
    oReg.SaveSetting m_cRegSystemKey, "SheetNo", ""
    oReg.SaveSetting m_cRegSystemKey, "AutoPrint", ""
    
    oReg.SaveSetting m_cRegSoundKey, "CanceledTicket", ""
    oReg.SaveSetting m_cRegSoundKey, "CheckedTicket", ""
    oReg.SaveSetting m_cRegSoundKey, "CheckSucess", ""
    oReg.SaveSetting m_cRegSoundKey, "CheckTimeOn", ""
    oReg.SaveSetting m_cRegSoundKey, "InvalidTicket", ""
    oReg.SaveSetting m_cRegSoundKey, "NoMatchedBus", ""
    oReg.SaveSetting m_cRegSoundKey, "ReturnedTicket", ""
    oReg.SaveSetting m_cRegSoundKey, "StartupCheckTimeOn", ""
    oReg.SaveSetting m_cRegSoundKey, "FreeTicket", ""
    oReg.SaveSetting m_cRegSoundKey, "HalfTicket", ""
    oReg.SaveSetting m_cRegSoundKey, "PreferentialTicket1", ""
    oReg.SaveSetting m_cRegSoundKey, "PreferentialTicket2", ""
    oReg.SaveSetting m_cRegSoundKey, "PreferentialTicket3", ""
    Set oReg = Nothing
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

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
    With MDIMain
        Select Case peArea
            Case EStatusBarArea.ESB_WorkingInfo
                .abMenu.Bands("statusBar").Tools("pnWorkingInfo").Caption = pszInfo
            Case EStatusBarArea.ESB_ResultCountInfo
                .abMenu.Bands("statusBar").Tools("pnResultCountInfo").Caption = pszInfo
            Case EStatusBarArea.ESB_UserInfo
                .abMenu.Bands("statusBar").Tools("pnUserInfo").Caption = pszInfo
            Case EStatusBarArea.ESB_LoginTime
                If pszInfo <> "" Then pszInfo = "��¼ʱ��: " & pszInfo
                .abMenu.Bands("statusBar").Tools("pnLoginTime").Caption = pszInfo
        End Select
    End With
End Sub
'' *******************************************************************
'' *   Member Name: WriteProcessBar                                  *
'' *   Brief Description: дϵͳ������״̬                           *
'' *   Engineer: ½����                                              *
'' *******************************************************************
'Public Sub WriteProcessBar(Optional pbVisual As Boolean = True, Optional ByVal plCurrValue As Variant = 0, Optional ByVal plMaxValue As Variant = 0, Optional pszShowInfo As String = cszUnrepeatString)
''����ע��
''*************************************
''plCurrValue(��ǰ����ֵ)
''plMaxValue(������ֵ)
''*************************************
'    If plMaxValue = 0 And pbVisual = True Then Exit Sub
'    Dim nCurrProcess As Integer
'    With MDIMain.abMenu.Bands("statusBar").Tools("progressBar")
'        If pbVisual Then
'            If Not .Visible Then
'                .Visible = True
'                MDIMain.pbLoad.Max = 100
'            End If
'            nCurrProcess = Int(plCurrValue / plMaxValue * 100)
'            MDIMain.pbLoad.Value = nCurrProcess
'        Else
'            .Visible = False
'        End If
'    End With
'    If pszShowInfo <> cszUnrepeatString Then ShowSBInfo pszShowInfo, ESB_WorkingInfo
'End Sub
'���ó�ʼ������

Private Sub InitSystemParam()
    Dim oSystemParam As New SystemParam
    oSystemParam.Init g_oActiveUser
    'У������ʱ��
    Date = oSystemParam.NowDate
    Time = oSystemParam.NowTime
    
    '��ȡ��ʼ��ϵͳ����
    g_nBeginCheckTime = oSystemParam.BeginCheckTime
    g_nLatestExtraCheckTime = oSystemParam.LatestExtraCheckTime
    g_nExtraCheckTime = oSystemParam.ExtraCheckTime
    g_nCheckTicketTime = oSystemParam.CheckTicketTime
    g_szTitle = oSystemParam.RoadSheetTitle
    g_bAllowChangeRide = oSystemParam.AllowChangeRide
    g_szUnitID = oSystemParam.UnitID
    g_nCheckSheetLen = oSystemParam.CheckSheetLen
    g_tTicketType = oSystemParam.GetAllTicketType(1, True)
    
    g_bAllowStartChectNotRearchTime = oSystemParam.AllowStartChectNotRearchTime '�Ƿ�����δ������ʱ�俪��
    
    
    
    Set m_rsTicketType = oSystemParam.GetAllTicketTypeRS(TP_TicketTypeValid)

    Set oSystemParam = Nothing


    '���³�ʼע�������
    Dim oReg As CFreeReg
    Set oReg = New CFreeReg
    
    Set g_oChkTicket = New CheckTicket
    g_oChkTicket.Init g_oActiveUser
    
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    g_tCheckInfo.CheckGateNo = Trim(oReg.GetSetting(m_cRegSystemKey, "CheckGate"))
    g_tCheckInfo.CurrSheetNo = Format(Val(g_oChkTicket.GetLastCheckSheetID(g_oActiveUser.UserID)) + 1, String(g_nCheckSheetLen, "0"))  'Trim(oReg.GetSetting(m_cRegSystemKey, "SheetNo"))
    g_tCheckInfo.AutoPrint = IIf(Val(oReg.GetSetting(m_cRegSystemKey, "AutoPrint")) <> 0, True, False)
    g_tCheckInfo.CheckerId = g_oActiveUser.UserID
    g_tCheckInfo.Checker = g_oActiveUser.UserName
    g_tCheckInfo.CheckDate = Date


    '��ȡ��Ʊʱ������Ч�ļ�·��
    g_tEventSoundPath.CanceledTicket = oReg.GetSetting(m_cRegSoundKey, "CanceledTicket")
    g_tEventSoundPath.CheckedTicket = oReg.GetSetting(m_cRegSoundKey, "CheckedTicket")
    g_tEventSoundPath.CheckSucess = oReg.GetSetting(m_cRegSoundKey, "CheckSucess")
    g_tEventSoundPath.CheckTimeOn = oReg.GetSetting(m_cRegSoundKey, "CheckTimeOn")
    g_tEventSoundPath.InvalidTicket = oReg.GetSetting(m_cRegSoundKey, "InvalidTicket")
    g_tEventSoundPath.NoMatchedBus = oReg.GetSetting(m_cRegSoundKey, "NoMatchedBus")
    g_tEventSoundPath.ReturnedTicket = oReg.GetSetting(m_cRegSoundKey, "ReturnedTicket")
    g_tEventSoundPath.StartupCheckTimeOn = oReg.GetSetting(m_cRegSoundKey, "StartupCheckTimeOn")
    g_tEventSoundPath.FreeTicket = oReg.GetSetting(m_cRegSoundKey, "FreeTicket")
    g_tEventSoundPath.HalfTicket = oReg.GetSetting(m_cRegSoundKey, "HalfTicket")
    g_tEventSoundPath.PreferentialTicket1 = oReg.GetSetting(m_cRegSoundKey, " PreferentialTicket1")
    g_tEventSoundPath.PreferentialTicket2 = oReg.GetSetting(m_cRegSoundKey, " PreferentialTicket2")
    g_tEventSoundPath.PreferentialTicket3 = oReg.GetSetting(m_cRegSoundKey, " PreferentialTicket3")
    
    Set oReg = Nothing
End Sub
Public Function GetStatusString(nStatus As Integer) As String
    Select Case nStatus
        Case EREBusStatus.ST_BusChecking
            GetStatusString = "���ڼ�Ʊ"
        Case EREBusStatus.ST_BusExtraChecking
            GetStatusString = "���ڲ���"
        Case EREBusStatus.ST_BusMergeStopped, EREBusStatus.ST_BusSlitpStop
            GetStatusString = "����ͣ��"
        Case EREBusStatus.ST_BusNormal, EREBusStatus.ST_BusReplace
            GetStatusString = "δ��"
        Case EREBusStatus.ST_BusStopCheck
            GetStatusString = "ͣ��"
        Case EREBusStatus.ST_BusExtraChecking
            GetStatusString = "���ڲ���"
        Case EREBusStatus.ST_BusStopped
            GetStatusString = "����ͣ��"
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
