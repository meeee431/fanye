Attribute VB_Name = "mdDss"
Option Explicit

Public Const cszDss = ""



Public Function GetUniqueTeam(prsTemp As Recordset, paszTemp() As String) As String()
'�õ�Ψһ������
Dim nCount As Integer
Dim i As Integer
Dim j As Integer
Dim nCount2 As Integer
    nCount = ArrayLength(paszTemp)
    For i = 1 To prsTemp.RecordCount
        nCount2 = nCount
        For j = 1 To nCount2
            If UCase(Trim(prsTemp!user_id)) = UCase(Trim(paszTemp(j))) Then
                Exit For
            End If
        Next j
        If j > nCount2 Then
        '�����û�������ʱ������ӵ�������ȥ
            nCount = nCount + 1
            ReDim Preserve paszTemp(1 To nCount)
            paszTemp(nCount) = Trim(prsTemp!user_id)
        End If
        prsTemp.MoveNext
    Next i
    GetUniqueTeam = paszTemp
End Function

