Attribute VB_Name = "mdlCard"
Option Explicit

Public Const lCardPortAuto = 0      '�Զ�ѡ��

Public Const lCardPortSerial1 = 1   '����1
Public Const lCardPortSerial2 = 2   '����2
Public Const lCardPortSerial3 = 3   '����3
Public Const lCardPortSerial4 = 4   '����4
Public Const lCardPortSerial5 = 5   '����5

Public Const lCardPortUSB1 = 1001   'USB1
Public Const lCardPortUSB2 = 1002   'USB2
Public Const lCardPortUSB3 = 1003   'USB3
Public Const lCardPortUSB4 = 1004   'USB4
Public Const lCardPortUSB5 = 1005   'USB5

'���֤����������ֵ����
Public Const lSuccess = 0           '�ɹ�
Public Const lOperateSuccess = 144  '�����ɹ� 0x90
Public Const lHaveNotContent = 145  'û�и������� 0x91
Public Const lFindCardSuccess = 159 '�����ҿ��ɹ���Ϣ 0x9F
Public Const lPortMistake = 1       '�˿ڴ�ʧ��/�˿���δ��/�˿ںŲ��Ϸ� 0x01
Public Const lPCReciveTimeOut = 2   'PC���ճ�ʱ���ڹ涨��ʱ����δ���յ��涨���ȵ����� 0x02
Public Const lPCJudgeVerifyMistake = 3     'PC�ж�У��ʹ� 0x03
Public Const lUSBNotConfigure = 4   'USB�豸δ���� 0x04
Public Const lSAMNotUse = 5         '��SAM���ڲ����ã�ֻ��SDT_GetCOMBaudʱ���п��ܷ��� 0x05
Public Const lUSBHasForbidden = 6   'USB�豸������ 0x06
Public Const lSAMJudgeVerifyMistake = 16   'SAM�ж�У��ʹ� 0x10
Public Const lSAMReciveTimeOut = 17 'SAM���ճ�ʱ���ڹ涨��ʱ����δ���յ��涨���ȵ����ݡ�0x11
Public Const lReciveOrderMistake = 33   '����ҵ���ն˵�������󣬰��������еĸ�����ֵ���߼�������� 0x21
Public Const lBeyondRigt = 35       'ԽȨ�Ĳ������� 0x23
Public Const lFindCardFail = 128    '�ҿ����ɹ� 0x80
Public Const lSelectCardFail = 129  'ѡ�����ɹ� 0x81
Public Const lCardIdentfyMachineFail = 49   '����֤����ʧ�� 0x31
Public Const lMachineIdentfyCardFail = 50   '������֤��ʧ�� 0x32
Public Const lInfoValidateMistake = 51      '��Ϣ��֤���� 0x33
Public Const lCanNotOperateCard = 52        '��δ�ҿ������ܽ��жԿ��Ĳ��� 0x34
Public Const lCantNotIndentfyCardType = 64  '�޷�ʶ��Ŀ����� 0x40
Public Const lReadCartFail = 65     '��������ʧ�� 0x41
Public Const lWriteCardFail = 80    'д������ʧ�� 0x50
Public Const lUserLoginFail = 97    '�û���¼ʧ�� 0x61
Public Const lSelfCheckFail = 96    '�Լ�ʧ�ܣ����ܽ������� 0x60
Public Const lKDCHaveNotDownloadKey = 102  'KDCû��������ʽ��Կ 0x66

'���֤������������Ϣ����
Public Const cszSuccess = "�ɹ�"             '�ɹ�
Public Const cszOperateSuccess = "�����ɹ�"  '�����ɹ� 0x90
Public Const cszHaveNotContent = "û�и�������"  'û�и������� 0x91
Public Const cszFindCardSuccess = "�����ҿ��ɹ���Ϣ" '�����ҿ��ɹ���Ϣ 0x9F
Public Const cszPortMistake = "�˿ڴ�ʧ��/�˿���δ��/�˿ںŲ��Ϸ�"       '�˿ڴ�ʧ��/�˿���δ��/�˿ںŲ��Ϸ� 0x01
Public Const cszPCReciveTimeOut = "PC���ճ�ʱ���ڹ涨��ʱ����δ���յ��涨���ȵ�����"   'PC���ճ�ʱ���ڹ涨��ʱ����δ���յ��涨���ȵ����� 0x02
Public Const cszPCJudgeVerifyMistake = "PC�ж�У��ʹ�"     'PC�ж�У��ʹ� 0x03
Public Const cszUSBNotConfigure = "USB�豸δ����"   'USB�豸δ���� 0x04
Public Const cszSAMNotUse = "��SAM���ڲ����ã�ֻ��SDT_GetCOMBaudʱ���п��ܷ���"   '��SAM���ڲ����ã�ֻ��SDT_GetCOMBaudʱ���п��ܷ��� 0x05
Public Const cszUSBHasForbidden = "USB�豸������"   'USB�豸������ 0x06
Public Const cszSAMJudgeVerifyMistake = "SAM�ж�У��ʹ�"  'SAM�ж�У��ʹ� 0x10
Public Const cszSAMReciveTimeOut = "SAM���ճ�ʱ���ڹ涨��ʱ����δ���յ��涨���ȵ�����" 'SAM���ճ�ʱ���ڹ涨��ʱ����δ���յ��涨���ȵ����ݡ�0x11
Public Const cszReciveOrderMistake = "����ҵ���ն˵�������󣬰��������еĸ�����ֵ���߼������"   '����ҵ���ն˵�������󣬰��������еĸ�����ֵ���߼�������� 0x21
Public Const cszBeyondRigt = "ԽȨ�Ĳ�������"       'ԽȨ�Ĳ������� 0x23
Public Const cszFindCardFail = "�ҿ����ɹ�"    '�ҿ����ɹ� 0x80
Public Const cszSelectCardFail = "ѡ�����ɹ�"  'ѡ�����ɹ� 0x81
Public Const cszCardIdentfyMachineFail = "����֤����ʧ��"   '����֤����ʧ�� 0x31
Public Const cszMachineIdentfyCardFail = "������֤��ʧ��"   '������֤��ʧ�� 0x32
Public Const cszInfoValidateMistake = "��Ϣ��֤����"      '��Ϣ��֤���� 0x33
Public Const cszCanNotOperateCard = "��δ�ҿ������ܽ��жԿ��Ĳ���"        '��δ�ҿ������ܽ��жԿ��Ĳ��� 0x34
Public Const cszCantNotIndentfyCardType = "�޷�ʶ��Ŀ�����"  '�޷�ʶ��Ŀ����� 0x40
Public Const cszReadCartFail = "��������ʧ��"     '��������ʧ�� 0x41
Public Const cszWriteCardFail = "д������ʧ��"    'д������ʧ�� 0x50
Public Const cszUserLoginFail = "�û���¼ʧ��"    '�û���¼ʧ�� 0x61
Public Const cszSelfCheckFail = "�Լ�ʧ�ܣ����ܽ�������"    '�Լ�ʧ�ܣ����ܽ������� 0x60
Public Const cszKDCHaveNotDownloadKey = "KDCû��������ʽ��Կ"  'KDCû��������ʽ��Կ 0x66


Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As Any) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As Any) As Long
Public Const TH32CS_SNAPHEAPLIST = &H1
Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPTHREAD = &H4
Public Const TH32CS_SNAPMODULE = &H8
Public Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST + TH32CS_SNAPPROCESS + TH32CS_SNAPTHREAD + TH32CS_SNAPMODULE)
Public Const TH32CS_INHERIT = &H80000000

'���ҽ��̺���
Public Const MAX_PATH = 260
Public Type PROCESSENTRY32
    dwSize   As Long
    cntUsage   As Long
    th32ProcessID   As Long
    th32DefaultHeapID   As Long
    th32ModuleID   As Long
    cntThreads   As Long
    th32ParentProcessID   As Long
    pcPriClassBase   As Long
    dwFlags   As Long
    szExeFile   As String * MAX_PATH
End Type

'�������̺���
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Const PROCESS_ALL_ACCESS = 1

''��ʼ���������˿�
'Public Function SetCardPortPotevio(Optional faxCard As FirstActivex) As String
'On Error GoTo ErrHandle
'    Dim lReturn As Long
'    Dim szResolve As String
'
'    If OpenPort <> cszSuccess Then
'        SetCardPortPotevio = "��ʼ���豸ʧ�ܣ������ԣ�"
'        Exit Function
'    End If
'
'    lReturn = faxCard.setPortNum(lCardPortAuto)
'    szResolve = CardReturnedValueResolve(lReturn)
'
'    SetCardPortPotevio = szResolve
'
'    Exit Function
'ErrHandle:
'    WriteErrorLog "SetCardPortPotevio", "ERROR:" & Err.Source & "-->[" & Err.Number & "]" & Err.Description
'    Err.Raise Err.Number, Err.Source, Err.Description
'End Function

'��ʼ������(��������)
Public Function SetCardPortVision() As String
On Error GoTo ErrHandle
    Dim szResolve As String
    
    If Val(CVR_InitComm(lCardPortUSB1)) = 0 Then
        szResolve = "��ʼ��ʧ��"
    Else
        szResolve = cszSuccess
    End If
    
    SetCardPortVision = szResolve
    
    Exit Function
ErrHandle:
    WriteErrorLog "SetCardPortVision", "ERROR:" & err.Source & "-->[" & err.Number & "]" & err.Description
    err.Raise err.Number, err.Source, err.Description
End Function

''����
'Public Function GetReadCardPotevio(Optional faxCard As FirstActivex) As String
'On Error GoTo ErrHandle
'    Dim lReturn As Long
'    Dim szResolve As String
'
'    szResolve = SetCardPortPotevio(faxCard)
'    If szResolve <> cszSuccess Then
'        GetReadCardPotevio = szResolve
'        Exit Function
'    End If
'
'    lReturn = faxCard.ReadCard()
'    szResolve = CardReturnedValueResolve(lReturn)
'
'    GetReadCardPotevio = szResolve
'
'    Exit Function
'ErrHandle:
'    WriteErrorLog "GetReadCardPotevio", "ERROR:" & Err.Source & "-->[" & Err.Number & "]" & Err.Description
'    Err.Raise Err.Number, Err.Source, Err.Description
'End Function

'����
Public Function GetReadCardVision() As String
On Error GoTo ErrHandle
    Dim lReturn As Long
    Dim szResolve As String
    
    CVR_Authenticate
    If CVR_Read_Content(4) = 0 Then
        szResolve = "ʧ��"
    Else
        szResolve = cszSuccess
    End If
    
    GetReadCardVision = szResolve
    
    Exit Function
ErrHandle:
    WriteErrorLog "GetReadCardVision", "ERROR:" & err.Source & "-->[" & err.Number & "]" & err.Description
    err.Raise err.Number, err.Source, err.Description
End Function

'��ȡ����
'Public Function GetCardID(Optional faxCard As FirstActivex) As String
'On Error GoTo ErrHandle
'
'    faxCard.setPortNum lCardPortAuto
'    GetCardID = faxCard.CardNo
'
'    Exit Function
'ErrHandle:
'    WriteErrorLog "GetCardID", "ERROR:" & Err.Source & "-->[" & Err.Number & "]" & Err.Description
'    Err.Raise Err.Number, Err.Source, Err.Description
'End Function

''��ȡ���֤����
'Public Function GetCardName(Optional faxCard As FirstActivex) As String
'On Error GoTo ErrHandle
'
'    GetCardName = faxCard.NameL
'
'    Exit Function
'ErrHandle:
'    WriteErrorLog "GetCardName", "ERROR:" & Err.Source & "-->[" & Err.Number & "]" & Err.Description
'    Err.Raise Err.Number, Err.Source, Err.Description
'End Function

'����������ֵ����
Public Function CardReturnedValueResolve(Optional ReturnedValue As Long) As String
On Error GoTo ErrHandle
    Dim szResolveValue As String
    
    Select Case ReturnedValue
        Case lSuccess, lOperateSuccess, lFindCardSuccess
            szResolveValue = cszSuccess
        Case lHaveNotContent
            szResolveValue = cszHaveNotContent
        Case lPortMistake
            szResolveValue = cszPortMistake
        Case lPCReciveTimeOut
            szResolveValue = cszPCReciveTimeOut
        Case lPCJudgeVerifyMistake
            szResolveValue = cszPCJudgeVerifyMistake
        Case lUSBNotConfigure
            szResolveValue = cszUSBNotConfigure
        Case lSAMNotUse
            szResolveValue = cszSAMNotUse
        Case lUSBHasForbidden
            szResolveValue = cszUSBHasForbidden
        Case lSAMJudgeVerifyMistake
            szResolveValue = cszSAMJudgeVerifyMistake
        Case lSAMReciveTimeOut
            szResolveValue = cszSAMReciveTimeOut
        Case lReciveOrderMistake
            szResolveValue = cszReciveOrderMistake
        Case lBeyondRigt
            szResolveValue = cszBeyondRigt
        Case lFindCardFail
            szResolveValue = cszFindCardFail
        Case lSelectCardFail
            szResolveValue = cszSelectCardFail
        Case lCardIdentfyMachineFail
            szResolveValue = cszCardIdentfyMachineFail
        Case lMachineIdentfyCardFail
            szResolveValue = cszMachineIdentfyCardFail
        Case lInfoValidateMistake
            szResolveValue = cszInfoValidateMistake
        Case lCanNotOperateCard
            szResolveValue = cszCanNotOperateCard
        Case lCantNotIndentfyCardType
            szResolveValue = cszCantNotIndentfyCardType
        Case lReadCartFail
            szResolveValue = cszReadCartFail
        Case lWriteCardFail
            szResolveValue = cszWriteCardFail
        Case lUserLoginFail
            szResolveValue = cszUserLoginFail
        Case lSelfCheckFail
            szResolveValue = cszSelfCheckFail
        Case lKDCHaveNotDownloadKey
            szResolveValue = cszKDCHaveNotDownloadKey
    End Select
    
    CardReturnedValueResolve = szResolveValue
    
    Exit Function
ErrHandle:
    WriteErrorLog "CardReturnedValueResolve", "ERROR:" & err.Source & "-->[" & err.Number & "]" & err.Description
End Function

'�򿪴���
Private Function OpenPort() As String
On Error GoTo ErrHandle

    Dim szSuccess As String
    Dim nSuccess As Integer
    nSuccess = SDT_OpenPort(lCardPortUSB1)

    If nSuccess = 144 Then
        szSuccess = cszSuccess
    ElseIf nSuccess = 1 Then
        szSuccess = "���֤�Ķ���,���ڴ�ʧ��"
    ElseIf nSuccess = 2 Then
        szSuccess = "PC���ճ�ʱ���ڹ涨��ʱ����δ���յ��涨���ȵ�����"
    ElseIf nSuccess = 3 Then
        szSuccess = "PC�ж�У��ʹ�"
    ElseIf nSuccess = 4 Then
        szSuccess = "USB�豸δ����"
    ElseIf nSuccess = 6 Then
        szSuccess = "USB�豸������"
    Else
        szSuccess = "���֤�Ķ���,���ڴ�ʧ��"
    End If

    OpenPort = szSuccess


    
    Exit Function
ErrHandle:
    WriteErrorLog "OpenPort", "ERROR:" & err.Source & "-->[" & err.Number & "]" & err.Description
    err.Raise err.Number, err.Source, err.Descriptio
End Function

'�رմ���
Public Sub ClosePort()
On Error GoTo ErrHandle

'    SDT_ClosePort lCardPortUSB1
    CVR_CloseComm
    Exit Sub
ErrHandle:
    WriteErrorLog "ClosePort", "ERROR:" & err.Source & "-->[" & err.Number & "]" & err.Description
    err.Raise err.Number, err.Source, err.Description
End Sub

Public Function KillProcess(ProcessName As String) As Boolean

    '���ҽ��̺��ս����
    Dim hSnapshot     As Long, lRet       As Long, P       As PROCESSENTRY32
    Dim exitCode     As Long
    Dim myProcess     As Long
    Dim AppKill     As Boolean
    P.dwSize = Len(P)
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, ByVal 0)
    If hSnapshot Then
        lRet = Process32First(hSnapshot, P)
        Do While lRet
            If InStr(P.szExeFile, ProcessName) <> 0 Then
                myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, P.th32ProcessID)
                AppKill = TerminateProcess(myProcess, exitCode)
                Call CloseHandle(myProcess)
            End If
            lRet = Process32Next(hSnapshot, P)
        Loop
        lRet = CloseHandle(hSnapshot)
    End If
    KillProcess = AppKill
End Function

'��ȡ����
Public Function GetName() As String
On Error GoTo ErrHandle
    Dim strTemp As String
    Dim nReturnLen As Integer
    Dim nReturn As Integer
    
    strTemp = Space(255)
    nReturn = GetPeopleName(strTemp, nReturnLen)
    GetName = Trim(strTemp)
    Exit Function
ErrHandle:
    err.Raise err.Number, err.Source, err.Description
End Function

'��ȡ���֤����
Public Function GetCardID() As String
On Error GoTo ErrHandle
    Dim strTemp As String
    Dim nReturnLen As Integer
    Dim nReturn As Integer
    
    strTemp = Space(255)
    nReturn = GetPeopleIDCode(strTemp, nReturnLen)
    GetCardID = Trim(strTemp)
    Exit Function
ErrHandle:
    err.Raise err.Number, err.Source, err.Description
End Function

'��ȡ�Ա�
Public Function GetSex() As String
On Error GoTo ErrHandle
    Dim strTemp As String
    Dim nReturnLen As Integer
    Dim nReturn As Integer
    
    strTemp = Space(255)
    nReturn = GetPeopleSex(strTemp, nReturnLen)
    GetSex = Trim(strTemp)
    Exit Function
ErrHandle:
    err.Raise err.Number, err.Source, err.Description
End Function

'��ȡ����
Public Function GetNation() As String
On Error GoTo ErrHandle
    Dim strTemp As String
    Dim nReturnLen As Integer
    Dim nReturn As Integer
    
    strTemp = Space(255)
    nReturn = GetPeopleNation(strTemp, nReturnLen)
    GetNation = Trim(strTemp)
    Exit Function
ErrHandle:
    err.Raise err.Number, err.Source, err.Description
End Function

'��ȡ��������
Public Function GetBirthday() As String
On Error GoTo ErrHandle
    Dim strTemp As String
    Dim nReturnLen As Integer
    Dim nReturn As Integer
    
    strTemp = Space(255)
    nReturn = GetPeopleBirthday(strTemp, nReturnLen)
    GetBirthday = Trim(strTemp)
    Exit Function
ErrHandle:
    err.Raise err.Number, err.Source, err.Description
End Function

'��ȡ��ַ
Public Function GetAddress() As String
On Error GoTo ErrHandle
    Dim strTemp As String
    Dim nReturnLen As Integer
    Dim nReturn As Integer
    
    strTemp = Space(255)
    nReturn = GetPeopleAddress(strTemp, nReturnLen)
    GetAddress = Trim(strTemp)
    Exit Function
ErrHandle:
    err.Raise err.Number, err.Source, err.Description
End Function

'��ȡ��֤����
Public Function GetDepartmentEx() As String
On Error GoTo ErrHandle
    Dim strTemp As String
    Dim nReturnLen As Integer
    Dim nReturn As Integer
    
    strTemp = Space(255)
    nReturn = GetDepartment(strTemp, nReturnLen)
    GetDepartmentEx = Trim(strTemp)
    Exit Function
ErrHandle:
    err.Raise err.Number, err.Source, err.Description
End Function

'��ȡ��Ч��ʼ����
Public Function GetStartDateEx() As String
On Error GoTo ErrHandle
    Dim strTemp As String
    Dim nReturnLen As Integer
    Dim nReturn As Integer
    
    strTemp = Space(255)
    nReturn = GetStartDate(strTemp, nReturnLen)
    GetStartDateEx = Trim(strTemp)
    Exit Function
ErrHandle:
    err.Raise err.Number, err.Source, err.Description
End Function

'��ȡ��Ч��ֹ����
Public Function GetEndDateEx() As String
On Error GoTo ErrHandle
    Dim strTemp As String
    Dim nReturnLen As Integer
    Dim nReturn As Integer
    
    strTemp = Space(255)
    nReturn = GetEndDate(strTemp, nReturnLen)
    GetEndDateEx = Trim(strTemp)
    Exit Function
ErrHandle:
    err.Raise err.Number, err.Source, err.Description
End Function
