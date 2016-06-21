Attribute VB_Name = "mdMiddleLib"
'=========================================================================================
'Author:
'Detail:�м�㹲��ģ��
'=========================================================================================

Option Explicit
'------------------------------------------------------------------------------------------
'���³�������
'---------------------------------------------------------

'---------------------------------------------------------
'***********��������*************
Public Const cnMidErrBegin = 10000 '�м����ʼ�����
Public Const cnMidErrStep = 400 '�м����󲽳�
Public Const cnMidRightBegin = 100 '�м��Ȩ����Դ��ʼ
Public Const cnMidRightStep = 6 '��Դ����



Public Function GetConnectionStr(Optional ByVal pszWhich As String) As String
    Dim oReg As New CFreeReg
    Dim szDatabaseType As String
    Dim szServer As String, szUser As String, szPassword As String, szDatabase As String, szTimeout As String
    Dim szDriverType As String
    Dim szIntegrated As String '�Ƿ񼯳��ʻ�
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany       'HKEY_LOCAL_MACHINE
    '1�Ƚ�Ĭ��ֵ����

        
    Dim szDBSetSection As String
    szDBSetSection = IIf(Trim(pszWhich) = "", "DataBaseSet", "DataBaseSet\" & pszWhich)
    
    szDatabaseType = UCase(oReg.GetSetting(szDBSetSection, "DBType"))
    szServer = oReg.GetSetting(szDBSetSection, "DBServer")
    szUser = oReg.GetSetting(szDBSetSection, "User")
    szPassword = UnEncryptPassword(oReg.GetSetting(szDBSetSection, "Password"))
    szDatabase = oReg.GetSetting(szDBSetSection, "DataBase")
    szTimeout = oReg.GetSetting(szDBSetSection, "Timeout")
    Select Case szDatabaseType
        Case "SQLOLEDB.1"   'SQL Server
'SQLServer��֤��ʽ
'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=foricq;Data Source=jhxu
'NT���ɷ�ʽ
'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=RTArchDB;Data Source=LENGEND
            '�Ƿ񼯳ɷ�ʽ
            szIntegrated = oReg.GetSetting(szDBSetSection, "Integrated")
            GetConnectionStr = "Provider=" & szDatabaseType _
            & ";Persist Security Info=False" _
            & IIf(szIntegrated <> "", ";Integrated Security=" & szIntegrated, ";User ID=" & szUser & ";Password=" & szPassword) _
            & ";Initial Catalog=" & szDatabase _
            & ";Data Source=" & szServer _
            & IIf(szTimeout = "", "", ";Timeout=" & Val(szTimeout))
        Case "MSDASQL.1"
            'ODBC��������
            szDriverType = oReg.GetSetting(szDBSetSection, "DBDriverType")
            Select Case szDriverType
                Case "Sybase System 11", "Sybase ASE ODBC Driver"    'Sybase 11.xϵ��
'Sybase 11.x�����ַ���
    'Provider=MSDASQL.1;Persist Security Info=False
    ';Extended Properties="DRIVER={Sybase System 11};UID=lyq;DB=RTArchDB;SRVR=CHENF;PWD=activex"
                    szUser = oReg.GetSetting(szDBSetSection, "User")
                    szPassword = UnEncryptPassword(oReg.GetSetting(szDBSetSection, "Password"))
                    szDatabase = oReg.GetSetting(szDBSetSection, "DataBase")
                    GetConnectionStr = "Provider=" & szDatabaseType & ";Persist Security Info=False" _
                    & ";Extended Properties=""DRIVER={" & szDriverType & "}" & ";UID=" & szUser & ";DB=" & szDatabase & ";SRVR=" & szServer & ";PWD=" & szPassword & """" _
                    & IIf(szTimeout = "", "", ";Timeout=" & Val(szTimeout))
                Case Else       '����ODBC��������
                    GetConnectionStr = ""
            End Select
        Case "SYBASE.ASEOLEDBPROVIDER.2"   'Sybase OLE DB���İ�
'Provider=Sybase.ASEOLEDBProvider.2;������ȫ����Ϣ=False;�û� ID=sa;����Դ=sybase11
            GetConnectionStr = "Provider=" & szDatabaseType _
            & ";������ȫ����Ϣ=False" _
            & ";�û� ID=" & szUser & ";����=" & szPassword _
            & ";����Դ=" & szServer _
            & IIf(szTimeout = "", "", ";��ʱ����=" & Val(szTimeout))
        Case Else
            GetConnectionStr = ""
    End Select
End Function


'���ڱȽϺ���,���رȽϵ��ַ�������SQL
Public Function DBDateCompare(pdtDate As Date, pszField As String, Optional pszOperator As String = "=") As String
'    DBDateCompare = "  '" & ToDBDate(pdtDate) & "'" & pszOperator & "CONVERT(CHAR(10)," & pszField & ",120) "
    DBDateCompare = " '" & ToDBDate(pdtDate) & "'" & pszOperator & pszField & " "
End Function

'ʱ��ȽϺ���,���رȽϵ��ַ�������SQL
Public Function DBTimeCompare(pdtDate As Date, pszField As String, Optional pszOperator As String = "=") As String
    DBTimeCompare = "  '" & ToDBTime(pdtDate) & "'" & pszOperator & "CONVERT (CHAR(8)," & pszField & ",108) "
End Function

''��ָ���Ļ�����ת��Ϊ��IP��ַ
'Public Function HostToIP(pszHost As String) As String
'     HostToIP = pszHost
'End Function
'
''��ָ���Ļ�����ת��Ϊ���������
'Public Function HostToName(pszHost As String) As String
'    HostToName = pszHost
'End Function


'ʱ��ȽϺ�����ֻ�Ƚ���ʱ�䲿��
Public Function TimeDiff(pszInterval As String, pdtTime1 As Date, pdtTime2 As Date) As Long
    Dim dtFirst As Date, dtSecond As Date
    dtFirst = CDate(ToDBTime(pdtTime1))
    dtSecond = CDate(ToDBTime(pdtTime2))
    TimeDiff = DateDiff(pszInterval, dtFirst, dtSecond)
End Function
'�õ���ǰ����ʱ��
Public Function SelfNowDateTime() As Date
    Dim oDb As New RTConnection
    Dim szSql As String, rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr("")
    szSql = "SELECT GETDATE() AS now_time"
    Set rsTemp = oDb.Execute(szSql)
    
    SelfNowDateTime = ToDBDateTime(rsTemp!now_time)
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

'�õ���ǰ����
Public Function SelfNowDate() As Date
    Dim oDb As New RTConnection
    Dim szSql As String, rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr("")
    szSql = "SELECT GETDATE() AS now_time"
    Set rsTemp = oDb.Execute(szSql)

    SelfNowDate = ToDBDate(rsTemp!now_time)
    Set rsTemp = Nothing
    Set oDb = Nothing
    'SelfNowDate = Date
End Function

'�õ���ǰʱ��
Public Function SelfNowTime() As Date
    Dim oDb As New RTConnection
    Dim szSql As String, rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr("")
    szSql = "SELECT GETDATE() AS now_time"
    Set rsTemp = oDb.Execute(szSql)

    SelfNowTime = ToDBTime(rsTemp!now_time)
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

'��֯WHERE�Ӿ�
Private Function GetWhereStr(pszSql As String, pszAddition As String) As String
    Dim szTemp As String
    
    If InStr(1, pszSql, "WHERE") <> 0 Then
        szTemp = " AND " & pszAddition
    Else
        szTemp = " WHERE " & pszAddition
    End If
End Function

'' *******************************************************************
'' *   Brief Description: ����MTS����                                *
'' *   Engineer: ½����                                              *
'' *   Date Generated: 2001/02/16                                    *
'' *   Last Revision Date:                                           *
'' *******************************************************************
'Public Function CreateMTSObject(pszObjectName As String) As Object
'    If Not GetObjectContext Is Nothing Then
'        Set CreateMTSObject = GetObjectContext.CreateInstance(pszObjectName)
'    Else
'        Set CreateMTSObject = CreateObject(pszObjectName)
'    End If
'End Function
' *******************************************************************
' *   Brief Description:�ύMTS����                                 *
' *   Engineer: ½����                                              *
' *   Date Generated: 2001/02/16                                    *
' *   Last Revision Date:                                           *
' *******************************************************************
'Public Sub CommitMTSTransaction()
'    If Not GetObjectContext Is Nothing Then
'        GetObjectContext.SetComplete
'    End If
'End Sub
'' *******************************************************************
'' *   Brief Description: �ع�MTS����                                *
'' *   Engineer: ½����                                              *
'' *   Date Generated: 2001/02/16                                    *
'' *   Last Revision Date:                                           *
'' *******************************************************************
'Public Sub RollbackMTSTransaction()
'    If Not GetObjectContext Is Nothing Then
'        GetObjectContext.SetAbort
'    End If
'End Sub






'�ж�ָ�������ݿ������Ƿ��ǿյ�
Public Function DBDateIsEmpty(pdtIn As Date) As Boolean
    'Dim dtTemp As Date
    DBDateIsEmpty = IIf(Format(pdtIn, cszDateStr) = Format(cdtEmptyDate, cszDateStr), True, False)
End Function

'�ж�ָ�������ݿ�ʱ���Ƿ��ǿյ�
Public Function DBTimeIsEmpty(pdtIn As Date) As Boolean
    'Dim dtTemp As Date
    DBTimeIsEmpty = IIf(Format(pdtIn, cszTimeStr) = Format(cdtEmptyTime, cszTimeStr), True, False)
End Function

'�ж�ָ�������ݿ�����ʱ���Ƿ��ǿյ�
Public Function DBDateTimeIsEmpty(pdtIn As Date) As Boolean
    'Dim dtTemp As Date
    DBDateTimeIsEmpty = IIf(Format(pdtIn, cszDateTimeStr) = Format(cdtEmptyDateTime, cszDateTimeStr), True, False)
End Function

'���SQL��䣬�����������Ч��װ
Public Sub CheckSQL(ByRef pszNeedCheckSQL As String)
    pszNeedCheckSQL = pszNeedCheckSQL
End Sub
'���ֶ�ֵת����TSQL���õ��ֶ��ַ���
Public Function TransFieldValueToString(pvFieldValue As Variant) As String
    Dim aszSplitString() As String
    Dim i As Integer
    TransFieldValueToString = ""
    Select Case VarType(pvFieldValue)
    Case vbSingle, vbDouble, vbInteger, vbLong, vbCurrency, vbDecimal, vbByte
        TransFieldValueToString = 0
        If Not IsNull(pvFieldValue) Then
            TransFieldValueToString = pvFieldValue
        End If
    Case vbBoolean
        
        TransFieldValueToString = False
        If Not IsNull(pvFieldValue) Then
            TransFieldValueToString = IIf(pvFieldValue, 1, 0)
        End If
    Case vbDate
        
        TransFieldValueToString = "'" & cszEmptyDateStr & "'"
        If Not IsNull(pvFieldValue) Then
            TransFieldValueToString = "'" & ToDBDateTime(CDate(pvFieldValue)) & "'"
        End If
    Case vbString
        TransFieldValueToString = "''"
        
        If Not IsNull(pvFieldValue) Then
            aszSplitString = Split(pvFieldValue, "'")
            If ArrayLength(aszSplitString) > 0 Then     '�ڲ����������ź�˫���ţ��������
                TransFieldValueToString = "'" & aszSplitString(0) & "'"
                For i = 1 To ArrayLength(aszSplitString) - 1
                    TransFieldValueToString = TransFieldValueToString & "+" & Chr(34) & "'" & Chr(34) & "+'" & aszSplitString(i) & "'"
                Next i
            Else
                TransFieldValueToString = "'" & pvFieldValue & "'"  '�����Ž�������
            End If
        End If
    End Select
End Function

' *******************************************************************
' *   Brief Description: �ж��Ƿ�Ϸ��û�֤                         *
' *   Engineer: ½����                                              *
' *   Date Generated: 2001/02/16                                    *
' *   Last Revision Date:                                           *
' *******************************************************************
Public Sub AssertCreditCard(pszCreditCard As String)
'pszCreditCard �û�֤����
'    Dim oCreditCard As New CreditCard
'    oCreditCard.CreditString = pszCreditCard
'    If Not oCreditCard.SeekIt Then
'        RaiseError ERR_AssertCreditCard, cszCustomErrSource
'    End If
'    Set oCreditCard = Nothing
    Dim sTemp As String
    sTemp = Left(CStr(Cos(1)), 8)
    sTemp = "RT" & sTemp
    If pszCreditCard <> sTemp Then
        RaiseError ERR_AssertCreditCard, cszCustomErrSource
    End If
End Sub



Public Function EncodePassword(pszPassword As String) As String
    If pszPassword = "" Then
        EncodePassword = pszPassword
    Else
        Dim i As Integer
        Dim nLen As Integer
        Dim lResult As Long
        nLen = Len(pszPassword)
        lResult = 0
        For i = 1 To nLen
            lResult = lResult + i * Asc(Mid(pszPassword, i, 1))
        Next i
        EncodePassword = lResult
    End If
End Function

'��ʽ���ַ������ж��Ƿ��зǷ��ַ���ȱʡ�ķǷ��ַ���',Ҳ���Լ�ָ������Ȼ���ַ����ĺ󵼿ո�ȥ��
Public Function FormatStr(pszInStr As String, Optional pszInValidChars As String = "';") As String
    Dim nStrLen As Integer
    Dim i As Integer
    nStrLen = Len(pszInValidChars)
    If nStrLen > 0 Then
        For i = 1 To nStrLen
            If InStr(1, pszInStr, Mid(pszInValidChars, i, 1), vbTextCompare) > 0 Then ShowError 501 'ERR_StrIllegal
        Next i
    End If
    FormatStr = Trim(pszInStr)
End Function

'��ָ���Ļ�����ת��Ϊ��IP��ַ
Public Function HostToIP(pszHost As String) As String
    HostToIP = pszHost
End Function

'��ָ���Ļ�����ת��Ϊ���������
Public Function HostToName(pszHost As String) As String
    HostToName = pszHost
End Function

