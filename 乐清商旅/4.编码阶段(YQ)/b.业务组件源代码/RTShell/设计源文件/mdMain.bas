Attribute VB_Name = "mdMain"
' *******************************************************************
' *   Member Name: RealLocateLVW                                    *
' *   Brief Description: 根据匹配串实时ListView框位置               *
' *   Engineer: 陆勇庆                                              *
' *******************************************************************
Public Sub RealLocateLVW(pszMatchSZ As String, poLocateObject As Object, Optional pnColumnIndex As Integer = 1, Optional plStartIndex As Long = 1)
'参数注释
'*************************************
'pszMatchSZ(匹配串)
'poLocateObject(定位控件)
'pnColumnIndex(定位字段)
'plStartIndex(开始搜索位置)
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
    If pnColumnIndex = 1 Then       '主文本区域
        Set oFoundItem = poLocateObject.FindItem(pszMatchSZ, , , 1)
        If Not (oFoundItem Is Nothing) Then
            lFoundIndex = oFoundItem.Index
        End If
    Else        '其他区域
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
    '定位
    If oFoundItem Is Nothing Then
        Set oFoundItem = poLocateObject.ListItems(lFoundIndex)
    End If
    oFoundItem.EnsureVisible
    oFoundItem.Selected = True
End Sub

