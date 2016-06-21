Attribute VB_Name = "Print"
Option Explicit
Private mrsSheetData As Recordset   '·����¼��

Const cnStartStation = 1
Const cnCheckGateID = 2
Const cnChecker = 3
Const cnSheetID = 4

Const cnBusID = 5
Const cnRountName = 6
Const cnStartTime = 7

Const cnBusSerialNo = 8
Const cnLicenseTag = 9
Const cnCompanyName = 10

Const cnStationAndTicketType = 11
Const cnNumberAndPriceName = 12
Const cnNumberAndPrice1 = 13
Const cnNumberAndPrice2 = 14
Const cnNumberAndPrice3 = 15
Const cnNumberAndPrice4 = 16
Const cnNumberAndPrice5 = 17
Const cnNumberAndPrice6 = 18
Const cnNumberAndPrice7 = 19
Const cnNumberAndPrice8 = 20
Const cnNumberAndPrice9 = 21
Const cnNumberAndPrice10 = 22

Const cnNumberAndPrice11 = 23
Const cnNumberAndPrice12 = 24
Const cnNumberAndPrice13 = 25
Const cnNumberAndPrice14 = 26
Const cnNumberAndPrice15 = 27
Const cnNumberAndPrice16 = 28

Const cnNumberAndPrice17 = 29
Const cnNumberAndPrice18 = 30

Public m_oFastPrint As Object


'��ӡ·��
Public Sub PrintSheet(mszSheetID As String)
    Dim moChkTicket As New STChkTk.CheckTicket
    Dim atSheetResult()  As TCheckSheetStationInfoEx
    Dim tSheetInfo As TCheckSheetInfo
    Dim nCount As Integer
    Dim szStation As String
    Dim i As Integer, j As Integer, k As Integer
    Dim szChecker As String
    Dim aszSheetInfo() As String
    Dim dbTotalMan As Double
    Dim dbTotalPrice As Double
    Dim dbTotalMileage As Double
'    Dim aszTemp() As String
    Dim szTemp As String
    Dim nTicketTypeCount As Integer
    Dim dtNow As Date
    
    
    Const cnStationLen = 8
    Const cnManLen = 5
    Const cnPriceLen = 8
    Const cszSplit = " "
    
    
    
    m_oFastPrint.ClosePort
    m_oFastPrint.OpenPort
    
    
    On Error GoTo Error_Handle
    moChkTicket.Init g_oActiveUser
    tSheetInfo = moChkTicket.GetCheckSheetInfo(mszSheetID)
    '�����Զ�����Ŀ
    ReDim maszSheetCustom(1 To 17, 1 To 2)
    
    '���ó�����Ϣ
    Dim oVehicle As Vehicle
    Set oVehicle = New Vehicle
    oVehicle.Init g_oActiveUser
    oVehicle.Identify tSheetInfo.szVehicleId

    Dim oRoute As Route
    Set oRoute = New Route
    oRoute.Init g_oActiveUser
    oRoute.Identify Trim(tSheetInfo.szRouteID)
    
    Dim oVehicleType As New VehicleModel
    oVehicleType.Init g_oActiveUser
    oVehicleType.Identify tSheetInfo.szVehicleModelID
    
    '�õ�·��վ����ϸ��Ϣ
    atSheetResult = moChkTicket.GetCheckSheetStationInfo(mszSheetID)
    nCount = ArrayLength(atSheetResult)
    
    
    dbTotalMan = 0
    dbTotalPrice = 0
    dbTotalMileage = 0
    If nCount > 0 Then
        ReDim aszSheetInfo(1 To nCount, 1 To 13)
    End If
    j = 0
    For i = 1 To nCount
        If j = 0 Then
            aszSheetInfo(1, 1) = atSheetResult(1).szStationID
            aszSheetInfo(1, 13) = atSheetResult(i).sgMileage
            j = 1
        End If
        If atSheetResult(i).szStationID <> aszSheetInfo(j, 1) Then
                j = j + 1
                aszSheetInfo(j, 1) = atSheetResult(i).szStationID
                aszSheetInfo(j, 13) = atSheetResult(i).sgMileage
        End If
        If atSheetResult(i).nCheckStatus <> ECheckedTicketStatus.NormalTicket Then
            aszSheetInfo(j, 2) = LeftAndRight(LeftAndRight(atSheetResult(i).szCheckSheet, False, "["), True, "]") & "(�Ĳ�)"
        Else
            aszSheetInfo(j, 2) = Trim(LeftAndRight(LeftAndRight(atSheetResult(i).szCheckSheet, False, "["), True, "]"))
        End If
        If atSheetResult(i).nTicketType = TP_FullPrice Then
            aszSheetInfo(j, 3) = atSheetResult(i).nManCount
            aszSheetInfo(j, 4) = atSheetResult(i).sgTicketPrice
            dbTotalMan = dbTotalMan + atSheetResult(i).nManCount
            dbTotalPrice = dbTotalPrice + atSheetResult(i).sgTicketPrice
            dbTotalMileage = dbTotalMileage + atSheetResult(i).sgMileage * atSheetResult(i).nManCount
        End If
        If atSheetResult(i).nTicketType = TP_HalfPrice Then
            aszSheetInfo(j, 5) = atSheetResult(i).nManCount
            aszSheetInfo(j, 6) = atSheetResult(i).sgTicketPrice
            dbTotalMan = dbTotalMan + atSheetResult(i).nManCount
            dbTotalPrice = dbTotalPrice + atSheetResult(i).sgTicketPrice
            dbTotalMileage = dbTotalMileage + atSheetResult(i).sgMileage * atSheetResult(i).nManCount
'            aszSheetInfo(j, 7) = atSheetResult(i).nManCount
'            aszSheetInfo(j, 8) = atSheetResult(i).sgTicketPrice
        End If
        If atSheetResult(i).nTicketType = TP_PreferentialTicket1 Then
            aszSheetInfo(j, 7) = atSheetResult(i).nManCount
            aszSheetInfo(j, 8) = atSheetResult(i).sgTicketPrice
            dbTotalMan = dbTotalMan + atSheetResult(i).nManCount
            dbTotalPrice = dbTotalPrice + atSheetResult(i).sgTicketPrice
            dbTotalMileage = dbTotalMileage + atSheetResult(i).sgMileage * atSheetResult(i).nManCount
        End If
        If atSheetResult(i).nTicketType = TP_PreferentialTicket2 Then
            aszSheetInfo(j, 9) = atSheetResult(i).nManCount
            aszSheetInfo(j, 10) = atSheetResult(i).sgTicketPrice
            dbTotalMan = dbTotalMan + atSheetResult(i).nManCount
            dbTotalPrice = dbTotalPrice + atSheetResult(i).sgTicketPrice
            dbTotalMileage = dbTotalMileage + atSheetResult(i).sgMileage * atSheetResult(i).nManCount
        End If
        If atSheetResult(i).nTicketType = TP_PreferentialTicket3 Then
            aszSheetInfo(j, 11) = atSheetResult(i).nManCount
            aszSheetInfo(j, 12) = atSheetResult(i).sgTicketPrice
            dbTotalMan = dbTotalMan + atSheetResult(i).nManCount
            dbTotalPrice = dbTotalPrice + atSheetResult(i).sgTicketPrice
            dbTotalMileage = dbTotalMileage + atSheetResult(i).sgMileage * atSheetResult(i).nManCount
        End If
        If atSheetResult(i).nTicketType = TP_FreeTicket Then        '��Ʊ����ȫƱ
            aszSheetInfo(j, 3) = Val(aszSheetInfo(j, 3)) + atSheetResult(i).nManCount
            aszSheetInfo(j, 4) = Val(aszSheetInfo(j, 4)) + atSheetResult(i).sgTicketPrice
            dbTotalMan = dbTotalMan + atSheetResult(i).nManCount
            dbTotalPrice = dbTotalPrice + atSheetResult(i).sgTicketPrice
            dbTotalMileage = dbTotalMileage + atSheetResult(i).sgMileage * atSheetResult(i).nManCount
        End If
        
    Next i
    
    '����һ����¼��
    Set mrsSheetData = New Recordset
    mrsSheetData.CursorLocation = adUseClient
    '����������֧�ֵ��ֶ�
    mrsSheetData.Fields.Append "station_name", adVarChar, 30       'վ������
    mrsSheetData.Fields.Append "mileage", adVarChar, 30             '���
    mrsSheetData.Fields.Append "full_number", adVarChar, 30        'ȫƱ��
    mrsSheetData.Fields.Append "full_price", adVarChar, 30        'ȫƱ���
    mrsSheetData.Fields.Append "half_number", adVarChar, 30        '��Ʊ��
    mrsSheetData.Fields.Append "half_price", adVarChar, 30        '��Ʊ���
    mrsSheetData.Fields.Append "pre1_number", adVarChar, 30        '�Ż�Ʊ1��
    mrsSheetData.Fields.Append "pre1_price", adVarChar, 30        '�Ż�Ʊ1���
    mrsSheetData.Fields.Append "pre2_number", adVarChar, 30        '�Ż�Ʊ2��
    mrsSheetData.Fields.Append "pre2_price", adVarChar, 30        '�Ż�Ʊ2���
    mrsSheetData.Fields.Append "pre3_number", adVarChar, 30        '�Ż�Ʊ3��
    mrsSheetData.Fields.Append "pre3_price", adVarChar, 30        '�Ż�Ʊ3���
    mrsSheetData.Fields.Append "total_price", adVarChar, 30        '�ܼƽ��
    mrsSheetData.Fields.Append "total_number", adVarChar, 30
    
    
    mrsSheetData.Open
    
    Dim nTemp As Integer
    nTemp = 0
    Dim aszTemp(1 To 14) As String
    For i = 1 To nCount       '�����յļ�¼��
        mrsSheetData.AddNew

            For j = 1 To mrsSheetData.Fields.Count
                mrsSheetData.Fields(j - 1) = ""
            Next j

            mrsSheetData.Fields("station_name") = aszSheetInfo(i, 2)
            mrsSheetData.Fields("full_number") = aszSheetInfo(i, 3)
            mrsSheetData.Fields("full_price") = Format(aszSheetInfo(i, 4), "0.00")
            mrsSheetData.Fields("half_number") = aszSheetInfo(i, 5)
            mrsSheetData.Fields("half_price") = Format(aszSheetInfo(i, 6), "0.00")
            mrsSheetData.Fields("pre1_number") = aszSheetInfo(i, 7)
            mrsSheetData.Fields("pre1_price") = Format(aszSheetInfo(i, 8), "0.00")
            mrsSheetData.Fields("pre2_number") = aszSheetInfo(i, 9)
            mrsSheetData.Fields("pre2_price") = Format(aszSheetInfo(i, 10), "0.00")
            mrsSheetData.Fields("pre3_number") = aszSheetInfo(i, 11)
            mrsSheetData.Fields("pre3_price") = Format(aszSheetInfo(i, 12), "0.00")
            mrsSheetData.Fields("total_price") = Format(Val(aszSheetInfo(i, 4)) + Val(aszSheetInfo(i, 6)) + Val(aszSheetInfo(i, 8)) + Val(aszSheetInfo(i, 10)) + Val(aszSheetInfo(i, 12)), "0.00")
            
            mrsSheetData.Fields("mileage") = aszSheetInfo(i, 13)
            
            mrsSheetData.Fields("total_number") = Val(aszSheetInfo(i, 3)) + Val(aszSheetInfo(i, 5)) + Val(aszSheetInfo(i, 7)) + Val(aszSheetInfo(i, 9)) + Val(aszSheetInfo(i, 11))

            '����
            For j = 3 To 12
                aszTemp(j) = Val(aszTemp(j)) + Val(aszSheetInfo(i, j))
            Next j

            aszTemp(13) = Val(aszTemp(13)) + Val(mrsSheetData!total_number)
            aszTemp(14) = Val(aszTemp(14)) + Val(mrsSheetData!total_price)
        mrsSheetData.Update
    Next i
    
    
    '---------------------------------------
    '����Ϊ��ӡ·���Ĵ���
    '���³������Ʊ�����ò�����,��ӡλ�û᲻��
    '---------------------------------------
    
    m_oFastPrint.SetObject cnStartStation
    m_oFastPrint.SetCaption Trim(g_szSellStationName)
    
    m_oFastPrint.SetObject cnCheckGateID
    m_oFastPrint.SetCaption Format(tSheetInfo.szCheckGateID)
    
    m_oFastPrint.SetObject cnChecker
    m_oFastPrint.SetCaption Trim(tSheetInfo.szMakeSheetUser)

    m_oFastPrint.SetObject cnSheetID
    m_oFastPrint.SetCaption "[" & mszSheetID & "]"

    m_oFastPrint.SetObject cnBusID
    m_oFastPrint.SetCaption Trim(tSheetInfo.szBusid) & IIf(tSheetInfo.nBusSerialNo > 0, "-" & tSheetInfo.nBusSerialNo, "")

    m_oFastPrint.SetObject cnRountName
    m_oFastPrint.SetCaption Trim(oRoute.RouteName)
    
    dtNow = Now
    m_oFastPrint.SetObject cnStartTime   '
    m_oFastPrint.SetCaption Format(dtNow, "HH:mm")
    
    m_oFastPrint.SetObject cnBusSerialNo
    m_oFastPrint.SetCaption Format(tSheetInfo.nBusSerialNo)
    
    m_oFastPrint.SetObject cnLicenseTag
    m_oFastPrint.SetCaption Trim(oVehicle.LicenseTag)

    m_oFastPrint.SetObject cnCompanyName
    m_oFastPrint.SetCaption Trim(oVehicle.CompanyName)
    
    szTemp = ""
    Dim l As Integer
    nTicketTypeCount = ArrayLength(g_tTicketType)
    For l = 1 To nTicketTypeCount
        szTemp = szTemp & FormatSize(Trim(g_tTicketType(l).szTicketTypeName), cnManLen + cnPriceLen + 1, 2) & cszSplit
    Next l
    m_oFastPrint.SetObject cnStationAndTicketType
    szTemp = FormatSize("վ��", cnStationLen, 2) & cszSplit & szTemp & FormatSize("�ϼ�", cnManLen + cnPriceLen, 2)
    m_oFastPrint.SetCaption szTemp

    szTemp = FormatSize("", cnStationLen) & cszSplit
    For j = 1 To m_rsTicketType.RecordCount + 1
        szTemp = szTemp & FormatSize("����", cnManLen, 2) & cszSplit & FormatSize("Ʊ��", cnPriceLen, 2) & cszSplit
    Next j
    m_oFastPrint.SetObject cnNumberAndPriceName
    m_oFastPrint.SetCaption szTemp
    
    If nCount > 16 Then
        'վ��������16 ��,����
        MsgBox "վ������������16��", vbExclamation, "��Ʊ��ӡ"
        Exit Sub
    End If
    '�������վ�������
    For i = 1 To nCount
        m_oFastPrint.SetObject cnNumberAndPriceName + i
        
        
        szTemp = FormatSize(aszSheetInfo(i, 2), cnStationLen) & cszSplit
        For j = 1 To ArrayLength(g_tTicketType)
            
            szTemp = szTemp & FormatSize(aszSheetInfo(i, j * 2 + 1), cnManLen) & cszSplit '����
            If Val(aszSheetInfo(i, j * 2 + 1)) > 0 Then
                szTemp = szTemp & FormatSize(Format(Val(aszSheetInfo(i, j * 2 + 2)) / aszSheetInfo(i, j * 2 + 1), "0.00"), cnPriceLen) & cszSplit '���
            Else
                szTemp = szTemp & FormatSize("0", cnPriceLen) & cszSplit '���
            End If
        Next j
        dbTotalMan = Val(aszSheetInfo(i, 3)) + Val(aszSheetInfo(i, 5)) + Val(aszSheetInfo(i, 7)) + Val(aszSheetInfo(i, 9)) + Val(aszSheetInfo(i, 11)) '�ϼ�����
        dbTotalPrice = FormatMoney(Val(aszSheetInfo(i, 4)) + Val(aszSheetInfo(i, 6)) + Val(aszSheetInfo(i, 8)) + Val(aszSheetInfo(i, 10)) + Val(aszSheetInfo(i, 12))) '�ϼƽ��
        szTemp = szTemp & FormatSize(Str(dbTotalMan), cnManLen) & cszSplit
        szTemp = szTemp & FormatSize(FormatMoney(dbTotalPrice), cnPriceLen)
        
        m_oFastPrint.SetCaption szTemp
'        Debug.Print szTemp
        
    Next i
    
    m_oFastPrint.SetObject cnNumberAndPriceName + nCount + 1
    szTemp = ""
    For j = 3 To 2 * (nTicketTypeCount + 1)
        If j Mod 2 = 1 Then
            szTemp = szTemp & FormatSize(aszTemp(j), cnManLen) & cszSplit
        Else
            szTemp = szTemp & FormatSize("0", cnPriceLen) & cszSplit 'FormatMoney(aszTemp(j))
        End If
        
    Next j
    '���Ӻϼ���
    m_oFastPrint.SetCaption FormatSize("�ϼ�", cnStationLen) & cszSplit & szTemp & FormatSize(aszTemp(13), cnManLen) & cszSplit & FormatSize(FormatMoney(aszTemp(14)), cnPriceLen)
    
    m_oFastPrint.SetObject cnNumberAndPriceName + nCount + 2
    m_oFastPrint.SetCaption "�Ʊ�ʱ�䣺" & ToDBDateTime(dtNow)
    
    For k = cnNumberAndPriceName + nCount + 3 To cnNumberAndPrice18
        
        'ʣ�µ�������Ϊ����
        m_oFastPrint.SetObject k
        m_oFastPrint.SetCaption ""
    Next k

    m_oFastPrint.PrintFile
    m_oFastPrint.ClosePort
    Exit Sub
Error_Handle:
    ShowErrorMsg

End Sub


Private Function FormatSize(pszSource As String, pnLen As Integer, Optional pnFlag As Integer = 0) As String
    'pnFlag Ϊ���뷽ʽ 0Ϊ��Ч,1�����,2����,3�Ҷ���    ,���Ϊ����,�������Ч
    '
    
    
    'ת�ַ���������ת��Ϊ�̶���С���ַ���
    Dim nlen As Integer
    Dim szTemp As String
    Dim nTemp As Integer
    nlen = LenString(pszSource)
    If nlen >= pnLen Then
        '������ȴ��ڹ涨����,��ԭ�������
        FormatSize = pszSource
        Exit Function
    End If
    If IsNumeric(pszSource) Then
        '���Ϊ����,���Ҷ���
        szTemp = IIf(Val(pszSource) = 0, "", pszSource)
        If szTemp = "" Then nlen = 0
        FormatSize = Space(pnLen - nlen) & szTemp
    Else
        If pnFlag = 1 Or pnFlag = 0 Then
            '���Ϊ�ַ�,�������
            FormatSize = pszSource & Space(pnLen - nlen)
        ElseIf pnFlag = 2 Then
            nTemp = Int((pnLen - nlen) / 2)
            If nTemp > 0 And pnLen - nTemp - nlen > 0 Then
                '�����ո������ҿո���������0
                FormatSize = Space(nTemp) & pszSource & Space(pnLen - nTemp - nlen)
            Else
                '�����
                FormatSize = pszSource & Space(pnLen - nlen)
            End If
        ElseIf pnFlag = 3 Then
            
            FormatSize = Space(pnLen - nlen) & szTemp
        End If
    End If
End Function



Public Sub GetIniFile()

'    On Error GoTo ErrorHandle
'
'
'    If FileIsExist(App.Path & "\chksheet.bpf") Then
'        m_oFastPrint.ReadFormatFile App.Path & "\chksheet.bpf"
'    Else
'        MsgBox "��ӡ�����ļ�""chksheet.bpf""δ�ҵ�,�޷����м�Ʊ����", vbCritical
'        End
'    End If
'
'
'    Exit Sub
'ErrorHandle:
'    ShowErrorMsg
End Sub


