Attribute VB_Name = "mdlCard"
Option Explicit

Public Const lCardPortAuto = 0      '自动选择

Public Const lCardPortSerial1 = 1   '串口1
Public Const lCardPortSerial2 = 2   '串口2
Public Const lCardPortSerial3 = 3   '串口3
Public Const lCardPortSerial4 = 4   '串口4
Public Const lCardPortSerial5 = 5   '串口5

Public Const lCardPortUSB1 = 1001   'USB1
Public Const lCardPortUSB2 = 1002   'USB2
Public Const lCardPortUSB3 = 1003   'USB3
Public Const lCardPortUSB4 = 1004   'USB4
Public Const lCardPortUSB5 = 1005   'USB5

'身份证读卡器返回值定义
Public Const lSuccess = 0           '成功
Public Const lOperateSuccess = 144  '操作成功 0x90
Public Const lHaveNotContent = 145  '没有该项内容 0x91
Public Const lFindCardSuccess = 159 '返回找卡成功信息 0x9F
Public Const lPortMistake = 1       '端口打开失败/端口尚未打开/端口号不合法 0x01
Public Const lPCReciveTimeOut = 2   'PC接收超时，在规定的时间内未接收到规定长度的数据 0x02
Public Const lPCJudgeVerifyMistake = 3     'PC判断校验和错 0x03
Public Const lUSBNotConfigure = 4   'USB设备未配置 0x04
Public Const lSAMNotUse = 5         '该SAM串口不可用，只在SDT_GetCOMBaud时才有可能返回 0x05
Public Const lUSBHasForbidden = 6   'USB设备被禁用 0x06
Public Const lSAMJudgeVerifyMistake = 16   'SAM判断校验和错 0x10
Public Const lSAMReciveTimeOut = 17 'SAM接收超时，在规定的时间内未接收到规定长度的数据。0x11
Public Const lReciveOrderMistake = 33   '接收业务终端的命令错误，包括命令中的各种数值或逻辑搭配错误 0x21
Public Const lBeyondRigt = 35       '越权的操作申请 0x23
Public Const lFindCardFail = 128    '找卡不成功 0x80
Public Const lSelectCardFail = 129  '选卡不成功 0x81
Public Const lCardIdentfyMachineFail = 49   '卡认证机具失败 0x31
Public Const lMachineIdentfyCardFail = 50   '机具认证卡失败 0x32
Public Const lInfoValidateMistake = 51      '信息验证错误 0x33
Public Const lCanNotOperateCard = 52        '尚未找卡，不能进行对卡的操作 0x34
Public Const lCantNotIndentfyCardType = 64  '无法识别的卡类型 0x40
Public Const lReadCartFail = 65     '读卡操作失败 0x41
Public Const lWriteCardFail = 80    '写卡操作失败 0x50
Public Const lUserLoginFail = 97    '用户登录失败 0x61
Public Const lSelfCheckFail = 96    '自检失败，不能接收命令 0x60
Public Const lKDCHaveNotDownloadKey = 102  'KDC没有下载正式密钥 0x66

'身份证读卡器返回信息定义
Public Const cszSuccess = "成功"             '成功
Public Const cszOperateSuccess = "操作成功"  '操作成功 0x90
Public Const cszHaveNotContent = "没有该项内容"  '没有该项内容 0x91
Public Const cszFindCardSuccess = "返回找卡成功信息" '返回找卡成功信息 0x9F
Public Const cszPortMistake = "端口打开失败/端口尚未打开/端口号不合法"       '端口打开失败/端口尚未打开/端口号不合法 0x01
Public Const cszPCReciveTimeOut = "PC接收超时，在规定的时间内未接收到规定长度的数据"   'PC接收超时，在规定的时间内未接收到规定长度的数据 0x02
Public Const cszPCJudgeVerifyMistake = "PC判断校验和错"     'PC判断校验和错 0x03
Public Const cszUSBNotConfigure = "USB设备未配置"   'USB设备未配置 0x04
Public Const cszSAMNotUse = "该SAM串口不可用，只在SDT_GetCOMBaud时才有可能返回"   '该SAM串口不可用，只在SDT_GetCOMBaud时才有可能返回 0x05
Public Const cszUSBHasForbidden = "USB设备被禁用"   'USB设备被禁用 0x06
Public Const cszSAMJudgeVerifyMistake = "SAM判断校验和错"  'SAM判断校验和错 0x10
Public Const cszSAMReciveTimeOut = "SAM接收超时，在规定的时间内未接收到规定长度的数据" 'SAM接收超时，在规定的时间内未接收到规定长度的数据。0x11
Public Const cszReciveOrderMistake = "接收业务终端的命令错误，包括命令中的各种数值或逻辑搭配错"   '接收业务终端的命令错误，包括命令中的各种数值或逻辑搭配错误 0x21
Public Const cszBeyondRigt = "越权的操作申请"       '越权的操作申请 0x23
Public Const cszFindCardFail = "找卡不成功"    '找卡不成功 0x80
Public Const cszSelectCardFail = "选卡不成功"  '选卡不成功 0x81
Public Const cszCardIdentfyMachineFail = "卡认证机具失败"   '卡认证机具失败 0x31
Public Const cszMachineIdentfyCardFail = "机具认证卡失败"   '机具认证卡失败 0x32
Public Const cszInfoValidateMistake = "信息验证错误"      '信息验证错误 0x33
Public Const cszCanNotOperateCard = "尚未找卡，不能进行对卡的操作"        '尚未找卡，不能进行对卡的操作 0x34
Public Const cszCantNotIndentfyCardType = "无法识别的卡类型"  '无法识别的卡类型 0x40
Public Const cszReadCartFail = "读卡操作失败"     '读卡操作失败 0x41
Public Const cszWriteCardFail = "写卡操作失败"    '写卡操作失败 0x50
Public Const cszUserLoginFail = "用户登录失败"    '用户登录失败 0x61
Public Const cszSelfCheckFail = "自检失败，不能接收命令"    '自检失败，不能接收命令 0x60
Public Const cszKDCHaveNotDownloadKey = "KDC没有下载正式密钥"  'KDC没有下载正式密钥 0x66


Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As Any) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As Any) As Long
Public Const TH32CS_SNAPHEAPLIST = &H1
Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPTHREAD = &H4
Public Const TH32CS_SNAPMODULE = &H8
Public Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST + TH32CS_SNAPPROCESS + TH32CS_SNAPTHREAD + TH32CS_SNAPMODULE)
Public Const TH32CS_INHERIT = &H80000000

'查找进程函数
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

'结束进程函数
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Const PROCESS_ALL_ACCESS = 1

''初始化读卡器端口
'Public Function SetCardPortPotevio(Optional faxCard As FirstActivex) As String
'On Error GoTo ErrHandle
'    Dim lReturn As Long
'    Dim szResolve As String
'
'    If OpenPort <> cszSuccess Then
'        SetCardPortPotevio = "初始化设备失败，请重试！"
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

'初始化连接(机具连接)
Public Function SetCardPortVision() As String
On Error GoTo ErrHandle
    Dim szResolve As String
    
    If Val(CVR_InitComm(lCardPortUSB1)) = 0 Then
        szResolve = "初始化失败"
    Else
        szResolve = cszSuccess
    End If
    
    SetCardPortVision = szResolve
    
    Exit Function
ErrHandle:
    WriteErrorLog "SetCardPortVision", "ERROR:" & err.Source & "-->[" & err.Number & "]" & err.Description
    err.Raise err.Number, err.Source, err.Description
End Function

''读卡
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

'读卡
Public Function GetReadCardVision() As String
On Error GoTo ErrHandle
    Dim lReturn As Long
    Dim szResolve As String
    
    CVR_Authenticate
    If CVR_Read_Content(4) = 0 Then
        szResolve = "失败"
    Else
        szResolve = cszSuccess
    End If
    
    GetReadCardVision = szResolve
    
    Exit Function
ErrHandle:
    WriteErrorLog "GetReadCardVision", "ERROR:" & err.Source & "-->[" & err.Number & "]" & err.Description
    err.Raise err.Number, err.Source, err.Description
End Function

'读取姓名
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

''读取身份证号码
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

'读卡器返回值解析
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

'打开串口
Private Function OpenPort() As String
On Error GoTo ErrHandle

    Dim szSuccess As String
    Dim nSuccess As Integer
    nSuccess = SDT_OpenPort(lCardPortUSB1)

    If nSuccess = 144 Then
        szSuccess = cszSuccess
    ElseIf nSuccess = 1 Then
        szSuccess = "身份证阅读器,串口打开失败"
    ElseIf nSuccess = 2 Then
        szSuccess = "PC接收超时，在规定的时间内未接收到规定长度的数据"
    ElseIf nSuccess = 3 Then
        szSuccess = "PC判断校验和错"
    ElseIf nSuccess = 4 Then
        szSuccess = "USB设备未配置"
    ElseIf nSuccess = 6 Then
        szSuccess = "USB设备被禁用"
    Else
        szSuccess = "身份证阅读器,串口打开失败"
    End If

    OpenPort = szSuccess


    
    Exit Function
ErrHandle:
    WriteErrorLog "OpenPort", "ERROR:" & err.Source & "-->[" & err.Number & "]" & err.Description
    err.Raise err.Number, err.Source, err.Descriptio
End Function

'关闭串口
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

    '查找进程和终结进程
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

'读取姓名
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

'读取身份证号码
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

'读取性别
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

'读取民族
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

'读取出生日期
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

'读取地址
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

'读取发证机关
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

'读取有效开始日期
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

'读取有效截止日期
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
