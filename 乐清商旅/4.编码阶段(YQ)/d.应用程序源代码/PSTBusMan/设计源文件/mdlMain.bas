Attribute VB_Name = "mdlMain"
Option Explicit

' *******************************************************************
' *  Source File Name: mdlMain                                      *
' *  Brief Description: ϵͳ��ģ��                                  *
' *******************************************************************
'====================================================================
'���³�������
'--------------------------------------------------------------------
'Public Const CnInternetCanSell = 0 '�ɻ�����Ʊ
'Public Const CnInternetNotCanSell = 1   '���ɻ�����Ʊ
Public Const cszRegKeySystem = "RTBusMan" '��ϵͳ��ע����
Public Const cvChangeColor = vbBlue
Public Const cszKeyPopMenu = 93
Public Const cnPreViewMaxDays = 30 '���ɻ���Ԥ�����������

'Ʊ���еĳ�������
Public Const cnNotRunTable = 0 'δ���е�Ʊ�۱�
Public Const cnRunTable = 1 '����ִ�е�Ʊ�۱�
Public Const cszItemBaseCarriage = "0000" '�����˼�
Public Const cnAllBusType = 100 '���г������ͣ�����β������
Public Const cszAllBusType = "��������" '���г������ͣ�����β������


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
'�Ի������״̬
Public Enum EFormStatus
    EFS_AddNew = 0
    EFS_Modify = 1
    EFS_Show = 2
    EFS_Delete = 3
End Enum

Public Enum ECheckStatus
    NormalTicket = 1
    ChangeTicket = 2
    MergeTicket = 3
End Enum

'====================================================================
'����ȫ�ֱ���
'--------------------------------------------------------------------
Public g_oActiveUser As ActiveUser
'Public g_szExePlanID As String 'ִ�еĳ��μƻ�
Public g_nPreSell As Byte
Public g_szExePriceTable As String 'ִ�е�Ʊ�۱�
Public g_szLocalUnit As String      '��ǰ��λ
Public g_bStopAllRefundment As Boolean '�Ƿ�ȫ����Ʊ
Public g_szLicenseForce As String '����ǰ׺

Public g_szUserPassword As String '�û�����
Public g_szStationID As String 'ϵͳ�����еı�վ����

'==================================================================
'���±���Ʊ�۲����õ�
Public g_atTicketTypeValid() As TTicketType '���õ�Ʊ����ϸ
Public g_nTicketCountValid As Integer   '���õ�Ʊ����Ŀ
Public g_atAllSellStation() As TDepartmentInfo  '���е���Ʊվ��

Public Sub Main()
'===================================================
'Modify Date��2002-11-19
'Author:½����
'Reamrk:�����ȫ�ֱ���g_atAllSellStation���ڴ��������Ʊվ��
'===================================================
    
    On Error GoTo ErrHandle
    Dim oShell As New CommShell
    Dim dtTemp As Date
    Dim oScheme As New RegularScheme
    Dim oBus As New BusProject
    Dim oSys As New SystemParam
'    If App.PrevInstance Then
'    MsgBox "Ӧ�ó����Ѵ�!", vbExclamation, "����"
'    Exit Sub
'    End If
    
    
    
    
    Set g_oActiveUser = oShell.ShowLogin()
    
    If g_oActiveUser Is Nothing Then Exit Sub
'    oShell.ShowSplash "�ۺϵ���", "Bus Scheme", LoadResPicture(101, 0), App.Major, App.Minor, App.Revision
    DoEvents
    dtTemp = Now
    g_szUserPassword = oShell.UserPassword
    '��ʼ���ã��õ�����ȫ�ֲ���
    oScheme.Init g_oActiveUser
    oBus.Init g_oActiveUser
    oSys.Init g_oActiveUser
    g_bStopAllRefundment = True
    
    Time = oSys.NowDateTime
    Date = oSys.NowDate
    g_szLocalUnit = oSys.UnitID
    g_nPreSell = oSys.PreSellDate
    g_szStationID = oSys.StationID
    '�õ����õ�Ʊ����Ϣ
    g_atTicketTypeValid = oSys.GetAllTicketType(TP_TicketTypeValid, False)
    g_nTicketCountValid = ArrayLength(g_atTicketTypeValid)
    
'    g_szExePlanID = oScheme.GetExecuteBusProject(Now).szProjectID
    '����ϵͳϵͳ
    '    g_bStopFullReturn = True
    
    
    
    '���ð���
    SetHTMLHelpStrings "stBusMan.chm"
    
    
    
    Dim oBase As New SystemMan
    oBase.Init g_oActiveUser
    g_atAllSellStation = oBase.GetAllSellStation '(g_szLocalUnit)
    
    oBus.Identify
    g_szExePriceTable = oBus.ExecutePriceTable
    Load MDIScheme
    MDIScheme.Show
'    DoEvents
'    Do
'    Loop While Second(Now - dtTemp) <= 3
    oShell.CloseSplash
    DoEvents
    

    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'���������վ������
Public Function AddCboStation(cboStationTemp As Object) As Boolean
    Dim oBusInfo As New BaseInfo
    Dim i As Integer
    Dim szaData() As String
    Dim nCount As Integer
    Dim cboStation As ComboBox
    Set cboStation = cboStationTemp
    oBusInfo.Init g_oActiveUser
    szaData = oBusInfo.GetStation(, cboStation.Text, cboStation.Text, cboStation.Text)
    Set oBusInfo = Nothing
    nCount = ArrayLength(szaData)
    If nCount > 0 Then
        cboStation.Clear
        For i = 1 To nCount
            cboStation.AddItem Trim(szaData(i, 1)) & "[" & Trim(szaData(i, 2)) & "]"
        Next
        cboStation.ListIndex = 0
    Else
        AddCboStation = False
        Beep
    End If
    AddCboStation = True
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
    With MDIScheme
        Select Case peArea
        Case EStatusBarArea.ESB_WorkingInfo
            .abMenuTool.Bands("statusBar").Tools("pnWorkingInfo").Caption = pszInfo
        Case EStatusBarArea.ESB_ResultCountInfo
            .abMenuTool.Bands("statusBar").Tools("pnResultCountInfo").Caption = pszInfo
        Case EStatusBarArea.ESB_UserInfo
            .abMenuTool.Bands("statusBar").Tools("pnUserInfo").Caption = pszInfo
        Case EStatusBarArea.ESB_LoginTime
            If pszInfo <> "" Then pszInfo = "��¼ʱ��: " & pszInfo
            .abMenuTool.Bands("statusBar").Tools("pnLoginTime").Caption = pszInfo
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
    With MDIScheme.abMenuTool.Bands("statusBar")
        If pbVisual Then
            If Not .Tools("progressBar").Visible Then
                .Tools("progressBar").Visible = True
                MDIScheme.pbLoad.Max = 100
                MDIScheme.abMenuTool.RecalcLayout
            End If
            nCurrProcess = Int(plCurrValue / plMaxValue * 100)
            MDIScheme.pbLoad.Value = nCurrProcess
        Else
            .Tools("progressBar").Visible = False
        End If
    End With
End Sub

'д������ı�����
Public Sub WriteTitleBar(Optional pszFormName As String = "", Optional poIcon As StdPicture)
'    'pszFormName��ʱ�����
'    With MDIScheme
'    If pszFormName = "" Then
'        .lblInfoBar = ""
'        Set .imgInfoBar.Picture = Nothing
'    Else
'        .lblInfoBar = pszFormName
'        Set .imgInfoBar.Picture = poIcon
'    End If
'    End With
End Sub
'�Ƿ񼤻�ϵͳ������
Public Sub ActiveSystemToolBar(pbTrue As Boolean)
    With MDIScheme
        .abMenuTool.Bands("mnu_System").Tools("mnu_ExportFile").Enabled = pbTrue
        .abMenuTool.Bands("mnu_System").Tools("mnu_ExportFileOpen").Enabled = pbTrue
        .abMenuTool.Bands("mnu_System").Tools("mnu_system_print").Enabled = pbTrue
        .abMenuTool.Bands("mnu_System").Tools("mnu_system_printview").Enabled = pbTrue
        .abMenuTool.Bands("mnu_System").Tools("mnu_PageOption").Enabled = pbTrue
        .abMenuTool.Bands("mnu_System").Tools("mnu_PrintOption").Enabled = pbTrue
        .abMenuTool.Bands("tbn_system").Tools("tbn_system_export").Enabled = pbTrue
        .abMenuTool.Bands("tbn_system").Tools("tbn_system_exportopen").Enabled = pbTrue
        .abMenuTool.Bands("tbn_system").Tools("tbn_system_print").Enabled = pbTrue
        .abMenuTool.Bands("tbn_system").Tools("tbn_system_printview").Enabled = pbTrue
    End With
End Sub


Public Function MakeArray(ByRef tvBuArray() As String, szOther As String)
    Dim nCount As Integer
    Dim j As Integer
    Dim bflgTemp As Boolean
    Dim nCountTemp As Integer
    nCountTemp = ArrayLength(tvBuArray)
    If nCountTemp = 0 Then
        ReDim tvBuArray(1 To 1) As String
        nCountTemp = 1
    End If
    j = 1
    If nCountTemp = 1 Then
        If tvBuArray(1) = "" Then
            tvBuArray(1) = szOther
        Else
            For j = 1 To nCountTemp
                If Trim(tvBuArray(j)) = Trim(szOther) Then
                    bflgTemp = False
                    Exit For
                Else
                    bflgTemp = True
                End If
            Next
        End If
    Else
        For j = 1 To nCountTemp
            If Trim(tvBuArray(j)) = Trim(szOther) Then
                bflgTemp = False: Exit For
            Else
                bflgTemp = True
            End If
        Next
    End If
    If bflgTemp = True Then
        ReDim Preserve tvBuArray(1 To nCountTemp + 1)
        tvBuArray(nCountTemp + 1) = Trim(szOther)
    End If

End Function

Public Function IdentifyBusStatus(eStatuts As EREBusStatus) As Boolean
    Dim szMsgBusStatus As String
    IdentifyBusStatus = False
    If eStatuts = ST_BusNormal Then IdentifyBusStatus = True: Exit Function
    If eStatuts <> ST_BusSlitpStop And eStatuts <> ST_BusStopped And eStatuts <> ST_BusMergeStopped Then
        Select Case eStatuts
        Case 3
            szMsgBusStatus = "�����Ѿ�ͣ��"
        Case 4
            szMsgBusStatus = "�������ڼ�Ʊ"
        Case 5
            szMsgBusStatus = "�������ڲ���"
        Case 16
            szMsgBusStatus = "�����Ѷ���"
        Case Else
            If eStatuts >= 32 Then
                szMsgBusStatus = "���ο������ڶ�����ֲ���������ͣ��"
                MsgBox szMsgBusStatus, vbInformation + vbOKOnly, "����ͣ��"
                Exit Function
            Else
                IdentifyBusStatus = True
                Exit Function
            End If
        End Select
        If szMsgBusStatus = "" Or MsgBox(szMsgBusStatus, vbInformation + vbYesNo, "����ͣ��") = vbYes Then
            IdentifyBusStatus = True
        End If
    Else
        Select Case eStatuts
        Case ST_BusSlitpStop
            szMsgBusStatus = "�����Ѳ��ͣ��"
        Case ST_BusStopped
            szMsgBusStatus = "������ͣ��,��������ͣ��"
        Case ST_BusMergeStopped
            szMsgBusStatus = "���β���ͣ��"
        End Select
        MsgBox szMsgBusStatus, vbInformation + vbOKOnly, "����ͣ��"
    End If
End Function

'
'Public Sub ShowTBInfo(Optional pszMsgInfo As String, Optional pnCount As Integer, Optional pnIndex As Integer, Optional pbVisibled As Boolean)
'
'End Sub

Public Function ConvertTypeFromArray(paszBusID() As String, paszVehicleModel() As String, paszSeatType() As String) As TBusVehicleSeatType()
    '�����δ��롢���͡���λ��������ת��Ϊ����
    Dim nBus As Integer
    Dim nSeatType As Integer
    Dim nVehicleType As Integer
    Dim lTemp As Long
    Dim atBusVehicleSeat() As TBusVehicleSeatType
    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    
    nBus = ArrayLength(paszBusID)
    nSeatType = ArrayLength(paszSeatType)
    nVehicleType = ArrayLength(paszVehicleModel)
'    If nSeatType > 30 And nBus = nSeatType And nBus = nVehicleType Then
        '�����λ���͸�������30 �ҳ�����������λ����������������
        lTemp = nBus
        ReDim atBusVehicleSeat(1 To lTemp)
        For i = 1 To lTemp
            atBusVehicleSeat(i).szbusID = paszBusID(i)
            atBusVehicleSeat(i).szVehicleTypeCode = paszVehicleModel(i)
            atBusVehicleSeat(i).szSeatTypeID = paszSeatType(i)
        Next i
'    Else
'        lTemp = nBus * nSeatType * nVehicleType
'        ReDim atBusVehicleSeat(1 To lTemp)
'        lTemp = 0
'        For i = 1 To nBus
'            For j = 1 To nVehicleType
'                For n = 1 To nSeatType
'                    lTemp = lTemp + 1
'                    atBusVehicleSeat(lTemp).szBusID = paszBusID(i)
'                    atBusVehicleSeat(lTemp).szVehicleTypeCode = paszVehicleModel(j)
'                    atBusVehicleSeat(lTemp).szSeatTypeID = paszSeatType(n)
'                Next n
'            Next j
'        Next i
'    End If
    ConvertTypeFromArray = atBusVehicleSeat
End Function

'Public Function GetProjectExcutePriceTable(ProjectID As String) As String
'    '�õ�ĳ�ƻ��ĵ�ǰ��ִ��Ʊ�۱�
'    On Error GoTo ErrorHandle
'    Dim oRegularScheme As New RegularScheme
'    Dim aszTable() As String
'    Dim i As Integer, nCount As Integer
'    Dim szTemp As String
'
'    oRegularScheme.Init g_oActiveUser
'    aszTable = oRegularScheme.ProjectExistTable(ProjectID)
'    nCount = ArrayLength(aszTable)
'    If nCount > 0 Then
'        For i = 1 To nCount
'             If aszTable(i, 6) <= Now Then
'               szTemp = aszTable(i, 2)
'               Exit For
'            End If
'        Next
'    End If
'    Set oRegularScheme = Nothing
'    GetProjectExcutePriceTable = szTemp
'    Exit Function
'
'ErrorHandle:
'    MsgBox "�˼ƻ�����ӦƱ�۱�"
'End Function

Public Function GetPriceTable(ThisDate As Date) As String()
    '�õ�Ʊ�۱�,�����䰴һ��˳�����к�,����
    
    Dim aszRoutePriceTable() As String
    Dim i, j As Integer, nCount As Integer
    Dim szPriceTable As String
    Dim oRegularScheme As New RegularScheme
    Dim tTemp As TSchemeArrangement
    Dim szRunProject As String

    Dim szPriceTableTemp() As String
    Dim dtMaxDate As Date '��ִ��Ʊ�۱�����ִ������]
    Dim oTicketPriceMan As New TicketPriceMan

On Error GoTo ErrorHandle
    oRegularScheme.Init g_oActiveUser
    tTemp = oRegularScheme.GetExecuteBusProject(Now)
    szRunProject = tTemp.szProjectID
    oTicketPriceMan.Init g_oActiveUser
    aszRoutePriceTable = oTicketPriceMan.GetAllRoutePriceTable()
    nCount = ArrayLength(aszRoutePriceTable)
    dtMaxDate = Format("1900-01-01", cszDateStr)

    If nCount > 0 Then
       ReDim szPriceTableTemp(1 To nCount, 7)
       '����ʼִ����������������ǰ��Ʊ�۱���Ϊִ�б�ǣ�������Ϊ��ִ�б��
       '���������մ������
        For i = 1 To nCount
            For j = 1 To 6
                szPriceTableTemp(i, j) = aszRoutePriceTable(i, j)
            Next
            If aszRoutePriceTable(i, 6) = szRunProject And Format(aszRoutePriceTable(i, 3), cszDateStr) <= Format(ThisDate, cszDateStr) Then
               szPriceTableTemp(i, 7) = cnRunTable
               If dtMaxDate < Format(aszRoutePriceTable(i, 3), cszDateStr) Then
                  dtMaxDate = Format(aszRoutePriceTable(i, 3), cszDateStr)
               End If
            Else
               szPriceTableTemp(i, 7) = cnNotRunTable
            End If
        Next
        '����ʼִ����������������ǰ�����������������Ʊ�۱���Ϊִ�У�������Ϊ��ִ�У�
        '����ΨһƱ�۱���Ϊִ��Ʊ�۱�
        For i = 1 To nCount
            If szPriceTableTemp(i, 7) = cnRunTable Then
               If dtMaxDate > Format(aszRoutePriceTable(i, 3), cszDateStr) Then
                  szPriceTableTemp(i, 7) = cnNotRunTable
               End If
            End If
        Next
    End If

    GetPriceTable = szPriceTableTemp
    Exit Function
ErrorHandle:
    ShowErrorMsg
End Function

'�ı����ı����ȼ��
Public Function TextLongValidate(nCharLong As Integer, szText As String) As Boolean
    Dim szTemp As String, szTemp1 As String, szTemp2 As String
    szTemp1 = CStr(nCharLong)
    If nCharLong Mod 2 = 0 Then
        szTemp2 = CStr(Int(nCharLong / 2))
    Else
        szTemp2 = CStr(Int(nCharLong / 2) + 0.5)
    End If
    szTemp = szText
    szTemp = StrConv(szTemp, vbFromUnicode)
    If LenB(szTemp) > nCharLong Then
        MsgBox "������" & szTemp1 & "������<Ӣ����ĸ>��" & szTemp2 & "������<����>,����ʹ��<Ӣ����ĸ>.", vbOKOnly + vbInformation, "ϵͳ����"
        TextLongValidate = True
    End If

End Function
'
