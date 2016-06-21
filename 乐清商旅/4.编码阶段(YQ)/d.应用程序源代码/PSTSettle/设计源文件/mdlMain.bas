Attribute VB_Name = "mdlMain"
Option Explicit

Const cszDocumentDir = "DocumentDir"
Const cszDefDocumentDir = "C:\"
Const cszLuggageAccount = "LugAcc"
Const cszRecentSeller = "RecentSeller"

Public Const cszCompanyName = "���˹�˾"
Public Const cszVehicleName = "����"
Public Const cszBusName = "����"

Public Const m_cRegSystemKey = "STSettle"         '����ϵͳ�����ַ���
Public g_oActiveUser As ActiveUser
Public g_oParam As New SystemParam
Public g_szUnitID As String
Public m_adbSplitItem() As Double   '�����޸�ʱ����
'Public m_IsSave As Boolean
Public g_szFixFeeItem As String '�������еĹ̶�������.
Public g_bAllowSellteTotalNegative As Boolean '�Ƿ����½���ĸ����㵽����
Public g_szSettleNegativeSplitItem As String '�����½���ĸ�������ڽ������λ��
Public Const g_cnSplitItemCount = 20
Public g_szIsFixFeeUpdateEachMonth As Boolean '�Ƿ�̶�������ÿ���¸��µ�


'Public g_tEventSoundPath as TEventSound                 '�¼���Ч�ļ�·��

'��������һЩ��ʹ�õ�API����
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'��������API�����õ��ĳ�������
Public Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Public Const SND_FILENAME = &H20000     '  name is a file name
Public Const SND_ASYNC = &H1         '  play asynchronously

Public Const HWND_TOPMOST = -1

'====================================================================
'���¶���ö��
'--------------------------------------------------------------------
'������״̬���ַ�������
Public Enum EStatusBarArea
    ESB_WorkingInfo = 1
    ESB_ResultCountInfo = 2
    ESB_UserInfo = 3
    ESB_LoginTime = 4
End Enum

Public Enum EFormStatus
    AddStatus = 0
    ModifyStatus = 1
    ShowStatus = 2
    NotStatus = 3
End Enum


Public Type TEventSound                        '�¼���Ч�ļ���·����Ϣ�ṹ
    CheckSheetNotExist As String           '·��������
    CheckSheetCanceled As String           '������·��
    CheckSheetSettled As String            '·���ѽ���
    CheckSheetSelected As String           '·����ѡ��
    CheckSheetValid As String              '·����Ч
    ObjectNotSame As String                '·����Ч,��������Ҫ����ʱ������Χ֮��
End Type

Public g_tEventSoundPath As TEventSound


Public Const m_cRegSoundKey = "Settle\EventSound"      '��Ч�����ַ���
Public g_atAllSellStation() As TDepartmentInfo  '���е���Ʊվ��
Public g_szStationID As String 'ϵͳ�����еı�վ����


Public Const cnPriceItemNum = 15


Public Sub Main()
    Dim oShell As New CommShell
    
    On Error GoTo ErrorHandle
    
     
    
    Set g_oActiveUser = oShell.ShowLogin()
    If g_oActiveUser Is Nothing Then Exit Sub
    
    InitSound
    
'    oShell.ShowSplash "����������", "Station Settle Management", LoadResPicture(101, 0), App.Major, App.Minor, App.Revision
    '������ʱ
    DoEvents
    mdiMain.Show
    
    
    g_oParam.Init g_oActiveUser
    g_szUnitID = g_oParam.UnitID
    Date = g_oParam.NowDate
    Time = g_oParam.NowTime

    g_szFixFeeItem = g_oParam.FixFeeItem '�������еĹ̶�������.
    g_bAllowSellteTotalNegative = g_oParam.AllowSettleTotalNegative '�Ƿ����½���ĸ����㵽����
    g_szSettleNegativeSplitItem = g_oParam.SettleNegativeSplitItem '�����½���ĸ�������ڽ������λ��
    g_szIsFixFeeUpdateEachMonth = g_oParam.IsFixFeeUpdateEachMonth  '�Ƿ�̶�������ÿ���¸��µ�
    
    
'    oShell.CloseSplash
    DoEvents
    Exit Sub
ErrorHandle:
    ShowErrorMsg
'    Resume GoOn
End Sub

Private Sub InitSound()

    '���³�ʼע�������
    Dim oReg As CFreeReg
    Set oReg = New CFreeReg
    
    
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    
    '��ȡ������Ч�ļ�·��
    g_tEventSoundPath.CheckSheetCanceled = oReg.GetSetting(m_cRegSoundKey, "CheckSheetCanceled")
    g_tEventSoundPath.CheckSheetNotExist = oReg.GetSetting(m_cRegSoundKey, "CheckSheetNotExist")
    g_tEventSoundPath.CheckSheetSelected = oReg.GetSetting(m_cRegSoundKey, "CheckSheetSelected")
    g_tEventSoundPath.CheckSheetSettled = oReg.GetSetting(m_cRegSoundKey, "CheckSheetSettled")
    g_tEventSoundPath.CheckSheetValid = oReg.GetSetting(m_cRegSoundKey, "CheckSheetValid")
    g_tEventSoundPath.ObjectNotSame = oReg.GetSetting(m_cRegSoundKey, "ObjectNotSame")
    
    
    Set oReg = Nothing
End Sub







Public Sub PlayEventSound(szFileName As String)
    '������Ч
    PlaySound szFileName, 0, SND_FILENAME + SND_ASYNC
    
End Sub

Public Function GetDocumentDir() As String
    Dim oReg As New CFreeReg
    Dim szFileDir As String
    On Error GoTo Error_Handle
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    szFileDir = oReg.GetSetting(cszLuggageAccount, cszDocumentDir, cszDefDocumentDir)
    szFileDir = IIf(szFileDir = "", cszDefDocumentDir, szFileDir)
    
    GetDocumentDir = szFileDir
    Exit Function
Error_Handle:
    GetDocumentDir = cszDefDocumentDir
End Function

Public Sub SaveDocumentDir(pszFullFileName As String)
    Dim oReg As New CFreeReg
    Dim szPath As String
    On Error Resume Next
    szPath = Left(pszFullFileName, InStrRev(pszFullFileName, "\") - 1)
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    If szPath <> "" Then
        oReg.SaveSetting cszLuggageAccount, cszDocumentDir, szPath
    Else
        oReg.SaveSetting cszLuggageAccount, cszDocumentDir, cszDocumentDir
    End If
End Sub

Public Sub SaveRecentSeller(pvaUser As Variant)
    Dim oReg As New CFreeReg
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
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
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    GetRecentSeller = oReg.GetSetting(cszLuggageAccount, cszRecentSeller)
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
        .abMenu.Refresh
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


Public Sub FillSellStation(cboSellStation As ComboBox)
    Dim oSystemMan As New SystemMan
    Dim atTemp() As TDepartmentInfo
    Dim i As Integer
    On Error GoTo here
    '�ж��û������ĸ��ϳ�վ,���Ϊ�������һ������,��������е��ϳ�վ
    oSystemMan.Init g_oActiveUser
    atTemp = oSystemMan.GetAllSellStation(g_szUnitID)
    If g_oActiveUser.SellStationID = "" Then
        cboSellStation.AddItem ""
        For i = 1 To ArrayLength(atTemp)
            cboSellStation.AddItem MakeDisplayString(atTemp(i).szSellStationID, atTemp(i).szSellStationName)
        Next i
    '����ֻ����û��������ϳ�վ
    Else
        For i = 1 To ArrayLength(atTemp)
            If g_oActiveUser.SellStationID = atTemp(i).szSellStationID Then
               cboSellStation.AddItem MakeDisplayString(atTemp(i).szSellStationID, atTemp(i).szSellStationName)
               Exit For
            End If
        Next i
        cboSellStation.ListIndex = 0
    End If
    Exit Sub
here:
    ShowErrorMsg
End Sub


'д������ı�����
Public Sub WriteTitleBar(Optional pszFormName As String = "", Optional poIcon As StdPicture)
'    'pszFormName��ʱ�����
'    With mdiMain
'    If pszFormName = "" Then
'        .lblInfoBar = ""
'        Set .imgInfoBar.Picture = Nothing
'    Else
'        .lblInfoBar = pszFormName
'        Set .imgInfoBar.Picture = poIcon
'    End If
'    End With
End Sub






Public Function GetDefaultMark(ByVal nValue As Integer) As String
    Select Case nValue
        Case 0
            GetDefaultMark = "Ĭ��"
        Case 1
            GetDefaultMark = "��Ĭ��"
    End Select
End Function


Public Function GetSplitStatus(ByVal Value As Integer) As String '1-ʹ��,0-δ��
    Select Case Value
        Case 0
            GetSplitStatus = "δ��"
        Case 1
            GetSplitStatus = "ʹ��"
    End Select
End Function

Public Function GetSplitType(ByVal Value As Integer) As String '0-����Է���˾,1-���վ��,2-��������˾
    Select Case Value
        Case 0
            GetSplitType = "����Է���˾"
        Case 1
            GetSplitType = "���վ��"
        Case 2
            GetSplitType = "��������˾"
    End Select
End Function
Public Function GetAllowModify(ByVal Value As Integer) As String '0-�������޸�,1-�����޸�
    Select Case Value
        Case 0
            GetAllowModify = "�������޸�"
        Case 1
            GetAllowModify = "�����޸�"
        
    End Select
End Function

  '0-���ʹ�˾ 1-���� 2-���˹�˾ 3-���� 4-����
Public Function GetObjectTypeInt(szObject As String) As Integer
    Select Case szObject
           Case "����"
            GetObjectTypeInt = CS_SettleByBus
           Case "����"
            GetObjectTypeInt = CS_SettleByVehicle
           Case "���˹�˾"
            GetObjectTypeInt = CS_SettleByTransportCompany
           Case "����"
            GetObjectTypeInt = CS_SettleByOwner
           Case "���˹�˾"
            GetObjectTypeInt = CS_SettleBySplitCompany
           Case "ȫ��"
            GetObjectTypeInt = -1
    End Select
End Function

Public Function GetObjectTypeString(nObject As Integer) As String
    Select Case nObject
           Case CS_SettleByBus
            GetObjectTypeString = "����"
           Case CS_SettleByVehicle
            GetObjectTypeString = "����"
           Case CS_SettleByTransportCompany
            GetObjectTypeString = "���˹�˾"
           Case CS_SettleByOwner
            GetObjectTypeString = "����"
           Case CS_SettleBySplitCompany
            GetObjectTypeString = "���˹�˾"
    End Select
End Function

''�õ����㵥״̬����
'Public Function GetSettleSheetStatusInt(szStatus As String) As Integer
'    Select Case szStatus
'        Case "δ��"
'            GetSettleSheetStatusInt = CS_SettleSheetValid
'        Case "����"
'            GetSettleSheetStatusInt = CS_SettleSheetInvalid
'        Case "�ѽ�"
'            GetSettleSheetStatusInt = CS_SettleSheetSettled
'    End Select
'End Function
'
''�õ����㵥״̬
'Public Function GetSettleSheetStatusString(szStatus As Integer) As String
'    Select Case szStatus
'        Case CS_SettleSheetValid
'            GetSettleSheetStatusString = "δ��"
'        Case CS_SettleSheetInvalid
'            GetSettleSheetStatusString = "����"
'        Case CS_SettleSheetSettled
'            GetSettleSheetStatusString = "�ѽ�"
'    End Select
'End Function





Public Function GetSettleSheetStatusInt(szStatus As String) As Integer
    Select Case szStatus
        Case "δ��"
            GetSettleSheetStatusInt = CS_SettleSheetValid
        Case "����"
            GetSettleSheetStatusInt = CS_SettleSheetInvalid
        Case "�ѻ�"
            GetSettleSheetStatusInt = CS_SettleSheetSettled
'        Case "Ӧ�ۿ�δ����"
'            GetSettleSheetStatusInt = CS_SettleSheetNegativeNotPay
'        Case "Ӧ�ۿ��ѽ���"
'            GetSettleSheetStatusInt = CS_SettleSheetNegativeHasPayed
        Case "ȫ��"
            GetSettleSheetStatusInt = -1
    End Select
End Function

Public Function GetSettleSheetStatusString(szStatus As Integer) As String
    Select Case szStatus
        Case CS_SettleSheetValid
            GetSettleSheetStatusString = "δ��"
        Case CS_SettleSheetInvalid
            GetSettleSheetStatusString = "����"
        Case CS_SettleSheetSettled
            GetSettleSheetStatusString = "�ѻ�"
        Case CS_SettleSheetNotInvalid
            GetSettleSheetStatusString = "������"
'        Case CS_SettleSheetNegativeHasPayed
'            GetSettleSheetStatusString = "Ӧ�ۿ��ѽ���"
        Case -1
            GetSettleSheetStatusString = "ȫ��"
    End Select
End Function




'�õ�·��վ��ĸĲ�״̬����
Public Function GetSheetStationStatusName(pszStatus As Integer)
    Select Case pszStatus
    '0-������1-��ǩ��2-����
    Case 0
        GetSheetStationStatusName = "����"
    Case 1
        GetSheetStationStatusName = "��ǩ"
    Case 2
        GetSheetStationStatusName = "����"
    End Select
    
    
End Function


Public Function GetQueryNegativeStatusString(nStatus As EQueryNegativeType) As String
    Select Case nStatus
        Case CS_QueryAll
            GetQueryNegativeStatusString = "ȫ��"
        Case CS_QueryNegative
            GetQueryNegativeStatusString = "Ӧ���Ϊ��"
        Case CS_QueryNotNegative
            GetQueryNegativeStatusString = "Ӧ���Ϊ��"
    End Select
End Function


Public Function GetQueryNegativeStatusInt(szStatus As String) As Integer
    Select Case szStatus
        Case "ȫ��"
            GetQueryNegativeStatusInt = CS_QueryAll
        Case "Ӧ���Ϊ��"
            GetQueryNegativeStatusInt = CS_QueryNegative
        Case "Ӧ���Ϊ��"
            GetQueryNegativeStatusInt = CS_QueryNotNegative
            
    End Select
End Function




'�õ�·��վ��ĸĲ�״̬����
Public Function GetFixFeeStatusName(pnStatus As Integer)
    Select Case pnStatus
    Case -1
        GetFixFeeStatusName = "ȫ��"
    Case 0
        GetFixFeeStatusName = "δ��"
    Case Else
        GetFixFeeStatusName = "�ѿ�"
    End Select
    
    
End Function
