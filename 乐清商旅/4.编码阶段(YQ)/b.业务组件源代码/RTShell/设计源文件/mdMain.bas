Attribute VB_Name = "mdMain"
' *******************************************************************
' *   Member Name: RealLocateLVW                                    *
' *   Brief Description: ����ƥ�䴮ʵʱListView��λ��               *
' *   Engineer: ½����                                              *
' *******************************************************************
Public Sub RealLocateLVW(pszMatchSZ As String, poLocateObject As Object, Optional pnColumnIndex As Integer = 1, Optional plStartIndex As Long = 1)
'����ע��
'*************************************
'pszMatchSZ(ƥ�䴮)
'poLocateObject(��λ�ؼ�)
'pnColumnIndex(��λ�ֶ�)
'plStartIndex(��ʼ����λ��)
'************************************
    Dim i As Long
    Dim oFoundItem As Object
    Dim lFoundIndex As Long
        
        
    If pnColumnIndex < 1 Then Exit Sub
    If pnColumnIndex > poLocateObject.ColumnHeaders.Count Then Exit Sub
    If plStartIndex < 1 Then plStartIndex = 1
    
    pszMatchSZ = UCase(pszMatchSZ)
    Dim nTmpLen As Integer
    nTmpLen = Len(pszMatchSZ)
    If nTmpLen = 0 Then Exit Sub
    
    lFoundIndex = -1
    If pnColumnIndex = 1 Then       '���ı�����
        Set oFoundItem = poLocateObject.FindItem(pszMatchSZ, , , 1)
        If Not (oFoundItem Is Nothing) Then
            lFoundIndex = oFoundItem.Index
        End If
    Else        '��������
        Dim szTmp As String
        For i = plStartIndex To poLocateObject.ListItems.Count
            szTmp = UCase(poLocateObject.ListItems(i).SubItems(pnColumnIndex - 1))
            If pszMatchSZ = Left(szTmp, nTmpLen) Then
                lFoundIndex = i
                Exit For
            End If
        Next i
    End If
    If lFoundIndex = -1 Then Exit Sub
    '��λ
    If oFoundItem Is Nothing Then
        Set oFoundItem = poLocateObject.ListItems(lFoundIndex)
    End If
    oFoundItem.EnsureVisible
    oFoundItem.Selected = True
End Sub

