Attribute VB_Name = "mdSystemBase"
Option Explicit


Public Function GetSpecialActiveUserCount(pszUserID As String) As Integer
    Dim oTemp As ActiveUserData
    Dim nCount As Integer
    nCount = 0
'#If ACTIVEUSER_USE_MEMORY Then
'    For Each oTemp In m_clActiveUserData
'        If oTemp.m_szUserID = pszUserID Then nCount = nCount + 1
'    Next
'#Else
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "SELECT * FROM active_user_info"
    Set rsTemp = oDb.Execute(szSql)
    Do While Not rsTemp.EOF
        If FormatDbValue(rsTemp!user_id) = pszUserID Then nCount = nCount + 1
        rsTemp.MoveNext
    Loop
'#End If
    GetSpecialActiveUserCount = nCount
End Function

   
