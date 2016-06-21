Attribute VB_Name = "mdMiddleLib"
'=========================================================================================
'Author:
'Detail:中间层共用模块
'=========================================================================================

Option Explicit
'------------------------------------------------------------------------------------------
'以下常量声明
'---------------------------------------------------------

'---------------------------------------------------------
'***********错误处理常量*************
Public Const cnMidErrBegin = 10000 '中间层起始错误号
Public Const cnMidErrStep = 400 '中间层错误步长
Public Const cnMidRightBegin = 100 '中间层权限资源开始
Public Const cnMidRightStep = 6 '资源步长



Public Function GetConnectionStr(Optional ByVal pszWhich As String) As String
    Dim oReg As New CFreeReg
    Dim szDatabaseType As String
    Dim szServer As String, szUser As String, szPassword As String, szDatabase As String, szTimeout As String
    Dim szDriverType As String
    Dim szIntegrated As String '是否集成帐户
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany       'HKEY_LOCAL_MACHINE
    '1先将默认值读出

        
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
'SQLServer认证方式
'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=foricq;Data Source=jhxu
'NT集成方式
'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=RTArchDB;Data Source=LENGEND
            '是否集成方式
            szIntegrated = oReg.GetSetting(szDBSetSection, "Integrated")
            GetConnectionStr = "Provider=" & szDatabaseType _
            & ";Persist Security Info=False" _
            & IIf(szIntegrated <> "", ";Integrated Security=" & szIntegrated, ";User ID=" & szUser & ";Password=" & szPassword) _
            & ";Initial Catalog=" & szDatabase _
            & ";Data Source=" & szServer _
            & IIf(szTimeout = "", "", ";Timeout=" & Val(szTimeout))
        Case "MSDASQL.1"
            'ODBC驱动类型
            szDriverType = oReg.GetSetting(szDBSetSection, "DBDriverType")
            Select Case szDriverType
                Case "Sybase System 11", "Sybase ASE ODBC Driver"    'Sybase 11.x系列
'Sybase 11.x连接字符串
    'Provider=MSDASQL.1;Persist Security Info=False
    ';Extended Properties="DRIVER={Sybase System 11};UID=lyq;DB=RTArchDB;SRVR=CHENF;PWD=activex"
                    szUser = oReg.GetSetting(szDBSetSection, "User")
                    szPassword = UnEncryptPassword(oReg.GetSetting(szDBSetSection, "Password"))
                    szDatabase = oReg.GetSetting(szDBSetSection, "DataBase")
                    GetConnectionStr = "Provider=" & szDatabaseType & ";Persist Security Info=False" _
                    & ";Extended Properties=""DRIVER={" & szDriverType & "}" & ";UID=" & szUser & ";DB=" & szDatabase & ";SRVR=" & szServer & ";PWD=" & szPassword & """" _
                    & IIf(szTimeout = "", "", ";Timeout=" & Val(szTimeout))
                Case Else       '其他ODBC驱动程序
                    GetConnectionStr = ""
            End Select
        Case "SYBASE.ASEOLEDBPROVIDER.2"   'Sybase OLE DB中文版
'Provider=Sybase.ASEOLEDBProvider.2;持续安全性信息=False;用户 ID=sa;数据源=sybase11
            GetConnectionStr = "Provider=" & szDatabaseType _
            & ";持续安全性信息=False" _
            & ";用户 ID=" & szUser & ";口令=" & szPassword _
            & ";数据源=" & szServer _
            & IIf(szTimeout = "", "", ";超时连接=" & Val(szTimeout))
        Case Else
            GetConnectionStr = ""
    End Select
End Function


'日期比较函数,返回比较的字符串用于SQL
Public Function DBDateCompare(pdtDate As Date, pszField As String, Optional pszOperator As String = "=") As String
'    DBDateCompare = "  '" & ToDBDate(pdtDate) & "'" & pszOperator & "CONVERT(CHAR(10)," & pszField & ",120) "
    DBDateCompare = " '" & ToDBDate(pdtDate) & "'" & pszOperator & pszField & " "
End Function

'时间比较函数,返回比较的字符串用于SQL
Public Function DBTimeCompare(pdtDate As Date, pszField As String, Optional pszOperator As String = "=") As String
    DBTimeCompare = "  '" & ToDBTime(pdtDate) & "'" & pszOperator & "CONVERT (CHAR(8)," & pszField & ",108) "
End Function

''将指定的机子名转变为其IP地址
'Public Function HostToIP(pszHost As String) As String
'     HostToIP = pszHost
'End Function
'
''将指定的机子名转变为其机子名字
'Public Function HostToName(pszHost As String) As String
'    HostToName = pszHost
'End Function


'时间比较函数，只比较其时间部分
Public Function TimeDiff(pszInterval As String, pdtTime1 As Date, pdtTime2 As Date) As Long
    Dim dtFirst As Date, dtSecond As Date
    dtFirst = CDate(ToDBTime(pdtTime1))
    dtSecond = CDate(ToDBTime(pdtTime2))
    TimeDiff = DateDiff(pszInterval, dtFirst, dtSecond)
End Function
'得到当前日期时间
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

'得到当前日期
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

'得到当前时间
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

'组织WHERE子句
Private Function GetWhereStr(pszSql As String, pszAddition As String) As String
    Dim szTemp As String
    
    If InStr(1, pszSql, "WHERE") <> 0 Then
        szTemp = " AND " & pszAddition
    Else
        szTemp = " WHERE " & pszAddition
    End If
End Function

'' *******************************************************************
'' *   Brief Description: 创建MTS对象                                *
'' *   Engineer: 陆勇庆                                              *
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
' *   Brief Description:提交MTS事务                                 *
' *   Engineer: 陆勇庆                                              *
' *   Date Generated: 2001/02/16                                    *
' *   Last Revision Date:                                           *
' *******************************************************************
'Public Sub CommitMTSTransaction()
'    If Not GetObjectContext Is Nothing Then
'        GetObjectContext.SetComplete
'    End If
'End Sub
'' *******************************************************************
'' *   Brief Description: 回滚MTS事务                                *
'' *   Engineer: 陆勇庆                                              *
'' *   Date Generated: 2001/02/16                                    *
'' *   Last Revision Date:                                           *
'' *******************************************************************
'Public Sub RollbackMTSTransaction()
'    If Not GetObjectContext Is Nothing Then
'        GetObjectContext.SetAbort
'    End If
'End Sub






'判断指定的数据库日期是否是空的
Public Function DBDateIsEmpty(pdtIn As Date) As Boolean
    'Dim dtTemp As Date
    DBDateIsEmpty = IIf(Format(pdtIn, cszDateStr) = Format(cdtEmptyDate, cszDateStr), True, False)
End Function

'判断指定的数据库时间是否是空的
Public Function DBTimeIsEmpty(pdtIn As Date) As Boolean
    'Dim dtTemp As Date
    DBTimeIsEmpty = IIf(Format(pdtIn, cszTimeStr) = Format(cdtEmptyTime, cszTimeStr), True, False)
End Function

'判断指定的数据库日期时间是否是空的
Public Function DBDateTimeIsEmpty(pdtIn As Date) As Boolean
    'Dim dtTemp As Date
    DBDateTimeIsEmpty = IIf(Format(pdtIn, cszDateTimeStr) = Format(cdtEmptyDateTime, cszDateTimeStr), True, False)
End Function

'检查SQL语句，并对其进行有效包装
Public Sub CheckSQL(ByRef pszNeedCheckSQL As String)
    pszNeedCheckSQL = pszNeedCheckSQL
End Sub
'将字段值转换成TSQL可用的字段字符串
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
            If ArrayLength(aszSplitString) > 0 Then     '内部包括单引号和双引号，则将其解释
                TransFieldValueToString = "'" & aszSplitString(0) & "'"
                For i = 1 To ArrayLength(aszSplitString) - 1
                    TransFieldValueToString = TransFieldValueToString & "+" & Chr(34) & "'" & Chr(34) & "+'" & aszSplitString(i) & "'"
                Next i
            Else
                TransFieldValueToString = "'" & pvFieldValue & "'"  '单引号将其括起
            End If
        End If
    End Select
End Function

' *******************************************************************
' *   Brief Description: 判断是否合法用户证                         *
' *   Engineer: 陆勇庆                                              *
' *   Date Generated: 2001/02/16                                    *
' *   Last Revision Date:                                           *
' *******************************************************************
Public Sub AssertCreditCard(pszCreditCard As String)
'pszCreditCard 用户证号码
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

'格式化字符串，判断是否有非法字符（缺省的非法字符是',也可自己指定），然后将字符串的后导空格去掉
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

'将指定的机子名转变为其IP地址
Public Function HostToIP(pszHost As String) As String
    HostToIP = pszHost
End Function

'将指定的机子名转变为其机子名字
Public Function HostToName(pszHost As String) As String
    HostToName = pszHost
End Function

