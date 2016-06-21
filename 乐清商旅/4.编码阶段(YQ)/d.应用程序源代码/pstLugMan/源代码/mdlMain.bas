Attribute VB_Name = "mdlMain"
Option Explicit

Public Const cszLuggageAccount = "LugAcc"


Const cszDocumentDir = "DocumentDir"
Const cszDefDocumentDir = "C:\"
Const cszRecentSeller = "RecentSeller"

Public m_oAUser As ActiveUser
Public g_oParam As New SystemParam
Public g_szUnitID As String
Public g_rsPriceItem As Recordset   '�˷����¼
Public m_oLuggageKinds As New LuggageKinds
Public m_oLugParam As New LuggageParam
Public moLugSvr As New LuggageSvr
Public moAcceptSheet As New AcceptSheet
Public moCarrySheet As New CarrySheet

Public m_oProtocol As New LugProtocol
Public m_oFinanceSheet As New FinanceSheet
Public m_oLugFormula As New LugFormula
Public m_oLugFinSvr As New LugFinSplitSvr
Public m_obase As New BaseInfo
Public m_oLugSplitSvr As New LugFinSplitSvr
Public m_oluggageSvr As New LuggageSvr


Public m_oPriceItemFunLib As New LugFunLib


Public m_bIsOneFormulaEachStation As Boolean

Public g_szCarrySheetID As String

Public Const szAcceptTypeGeneral = "���" '���˷�ʽ
Public Const szAcceptTypeMan = "��ͨ"

'Public Const szGeneralProtocol = 0 '���Ĭ��Э��
'Public Const szManProtocol = 1 '����Ĭ��Э��
'Public Const szNotProtocol = 2 '������Ĭ��Э��

Public Const szConstType = "�̶�����"  '0  ��������
Public Const szCalType = "��ʽ�������" '1

'����״̬
Public Const mStatusNo = "δ��"     '0
Public Const mStatusReal = "�ѽ�"   '1
Public Const mStatusCancel = "����" '2

'Public Const mStatusNoInt = 0
'Public Const mStatusRealInt = 1
'Public Const mStatusCancelInt = 2


Public mSplitCompanyID() As String  '��ѡ���ʹ�˾�ܼ�
Public mSplitVehicleID() As String  '�����ܼ�
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

Public Enum ECheckStat
    UI_SplitCompanyCheckStat = 1
    UI_VehicleCheckStat = 2
    UI_RouteCheckStat = 3
    
End Enum
'�Ի������״̬
'Public Enum EFormStatus
'    EFS_AddNew = 0
'    EFS_Modify = 1
'    EFS_Show = 2
'    EFS_Delete = 3
'End Enum

'''����ĵ�ǰ״̬
Public Enum eFormStatus
    AddStatus = 0
    ModifyStatus = 1
    ShowStatus = 2
    NotStatus = 3
End Enum
   
Public Sub Main()
    Dim oShell As New CommShell
    
    On Error GoTo HelpFileErr
'    App.HelpFile = SetHTMLHelpStrings("SNTKAcc.chm")
    
GoOn:
    On Error GoTo 0
    
    
    
    
    Set m_oAUser = oShell.ShowLogin()
    If m_oAUser Is Nothing Then Exit Sub
'    oShell.ShowSplash "�а�����", "RTStation Luggage Management", LoadResPicture(101, 0), App.Major, App.Minor, App.Revision
    DoEvents
    mdiMain.Show
        
    
    g_oParam.Init m_oAUser
    g_szUnitID = g_oParam.UnitID

    m_bIsOneFormulaEachStation = True 'g_oParam.IsOneFormulaEachStation '�Ƿ�ÿ��վ��һ����ʽ.
    
    
    Dim oLugSysParam As New STLuggage.LuggageParam
    oLugSysParam.Init m_oAUser
    Set g_rsPriceItem = oLugSysParam.GetPriceItemRS(0)
    
    moLugSvr.Init m_oAUser
    moAcceptSheet.Init m_oAUser
    moCarrySheet.Init m_oAUser
    m_obase.Init m_oAUser
    m_oLugFinSvr.Init m_oAUser
    m_oFinanceSheet.Init m_oAUser
    m_oLugFormula.Init m_oAUser
    m_oLuggageKinds.Init m_oAUser
    m_oluggageSvr.Init m_oAUser
    m_oLugParam.Init m_oAUser
    m_oPriceItemFunLib.Init m_oAUser
    
    '������ʱ
    
    oShell.CloseSplash
    DoEvents
    Exit Sub
HelpFileErr:
    ShowMsg "�����ҵ������ļ�"
    Resume GoOn
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
    oSystemMan.Init m_oAUser
    atTemp = oSystemMan.GetAllSellStation(g_szUnitID)
    If m_oAUser.SellStationID = "" Then
        cboSellStation.AddItem ""
        For i = 1 To ArrayLength(atTemp)
            cboSellStation.AddItem MakeDisplayString(atTemp(i).szSellStationID, atTemp(i).szSellStationName)
        Next i
    '����ֻ����û��������ϳ�վ
    Else
        For i = 1 To ArrayLength(atTemp)
            If m_oAUser.SellStationID = atTemp(i).szSellStationID Then
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

'�Ի������״̬

Public Function GetLuggageTypeString(pnType As Integer) As String
    Select Case pnType
        Case 0
            GetLuggageTypeString = szAcceptTypeGeneral
        Case 1
            GetLuggageTypeString = szAcceptTypeMan
    End Select
End Function

Public Function GetLuggageTypeInt(szType As String) As Integer
    Select Case szType
        Case szAcceptTypeGeneral
            GetLuggageTypeInt = 0
        Case szAcceptTypeMan
            GetLuggageTypeInt = 1
        Case "ȫ��", ""
            GetLuggageTypeInt = -1
            
    End Select
    
End Function
 
 '����״̬ת��
Public Function GetFinTypeString(pnType As Integer) As String
    Select Case pnType
        Case 0
            GetFinTypeString = mStatusCancel
        Case 1
            GetFinTypeString = mStatusReal
'        Case 2
'            GetFinTypeString = mStatusCancel
    End Select
End Function
Public Function GetFinTypeInt(szType As String) As Integer
    Select Case szType
        Case mStatusCancel
            GetFinTypeInt = 0
        Case mStatusReal
            GetFinTypeInt = 1
'        Case mStatusCancel
'            GetFinTypeInt = 2
        Case "ȫ��"
            GetFinTypeInt = -1
    End Select
End Function
  '0-���ʹ�˾ 1-���� 2-���˹�˾ 3-���� 4-����
Public Function GetObjectTypeInt(szType As String) As Integer
     Select Case szType
            Case "���ʹ�˾"
                  GetObjectTypeInt = 0
            Case "����"
                  GetObjectTypeInt = 1
            Case "���˹�˾"
                  GetObjectTypeInt = 2
            Case "����"
                  GetObjectTypeInt = 3
            Case "����"
                  GetObjectTypeInt = 4
     End Select
End Function
Public Sub HideSheetNoLabel()

End Sub
Public Sub SetSheetNoLabel(pbIsAccept As Boolean, pszSheetNo As String)

End Sub

