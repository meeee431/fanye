Attribute VB_Name = "mdChkTk"
Option Explicit

Public Const cszStartCheckTime = "BeginCheckTime"
Public Const cszExtraStartCheckTime = "LatestExtraCheckTime"

Public Const cnErroeStartNo = EClassErrBegin.ERR_CheckTicket
Public Const cnRightStartNo = EClassErrBegin.ERR_CheckTicket

Enum ECheckStatus
    NormalTicket = 1
    ChangeTicket = 2
    MergeTicket = 3
End Enum
'ȡϵͳ��������ʱ��
Public Function GetParameterValue(ParameterName As String) As Double
    Dim odb As New RTConnection
    Dim rsTemp As New Recordset
    Dim szSql As String
    
    odb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    szSql = "SELECT parameter_value from System_param_info WHERE parameter_name='" & ParameterName & "'"
    Set rsTemp = odb.Execute(szSql, , -1)
'    GetParameterValue = Val(rsTemp.Fields(0)) / (24 * 60)
     GetParameterValue = Val(rsTemp.Fields(0))
End Function

'ʱ��ȽϺ���,���رȽϵ��ַ�������SQL
Public Function DBTimeCompareEX(pdtDate As Date, pszField As String, Optional pszOperator As String = "=") As String

    DBTimeCompareEX = "  CONVERT(CHAR(10),CONVERT(DATETIME,'" & ToDBDateTime(pdtDate) & "'),108)" & pszOperator & "CONVERT (CHAR(10)," & pszField & ",108) "

End Function


