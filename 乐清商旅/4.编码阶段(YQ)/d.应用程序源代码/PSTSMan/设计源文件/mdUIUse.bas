Attribute VB_Name = "mdUIUse"
' *******************************************************************
' *  Source File Name  : mdSystemMan.bas                            *
' *  Project Name: PSTSMan                                          *
' *  Engineer:                                                      *
' *  Date Generated: 2002/08/19                                     *
' *  Last Revision Date : 2002/08/20                                *
' *  Brief Description   : ����ͨ�ú���                             *
' *******************************************************************

Option Explicit

'���������������ӿؼ�Enable����
Public Sub EnableContainer(poContainer As Object, Optional pbEnabled As Boolean = True, Optional poReserved As Object = Nothing)
    Dim oContainer As Object
    Dim oControl As Control
    On Error Resume Next
    For Each oControl In poContainer.Parent.Controls
        If Not oControl Is poReserved Then
            Set oContainer = ControlContainer(oControl)
            If Not oContainer Is Nothing Then
                If oContainer.name = poContainer.name Then
                    oControl.Enabled = pbEnabled
                End If
            End If
        End If
    Next
End Sub
'���ؿؼ�������(���������ؿ�)
Public Function ControlContainer(poControl As Control) As Object
    On Error GoTo here
    Set ControlContainer = poControl.Container
    Exit Function
here:
    Set ControlContainer = Nothing
End Function
'����������е�TextBox
Public Sub ClearTextBox(poContainer As Form)
    Dim oControl As Control
    For Each oControl In poContainer.Controls
        If TypeName(oControl) = "TextBox" Then
                 oControl.Text = ""
        End If
    Next

End Sub
'�ж������Ƿ�Ϊ��
Public Function IsArrayEmpty(pvaTemp As Variant) As Boolean
    Dim nTemp  As Integer
    On Error GoTo here
    nTemp = UBound(pvaTemp)
    IsArrayEmpty = False
    Exit Function
here:
    IsArrayEmpty = True
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

'�����ַ����
Public Function SpacialStrValid(szText As String, SpacialStr As String) As Boolean
    Dim nTemp As Integer
    nTemp = InStr(1, szText, SpacialStr)
    If nTemp = 0 Then
        SpacialStrValid = False
    Else
        MsgBox "�˴�����ʹ�������ַ���" & SpacialStr, vbInformation, cszMsg
        SpacialStrValid = True
    End If
    
End Function



Public Function GetIPString(szIPs As String) As String()
    Dim aszTemp() As String
    Dim szTemp As String
    Dim nComma As Integer
    Dim nPosition As Integer
    szTemp = szIPs
    nComma = 0
    nPosition = 1
    
    If (szTemp = "") Or (szTemp = Null) Then
        ''''
    Else
        Do While nPosition <> 0
            nPosition = InStr(1, szTemp, ",", vbBinaryCompare)
                If nPosition <> 0 Then
                    nComma = nComma + 1
                    ReDim Preserve aszTemp(1 To nComma)
                    aszTemp(nComma) = Left(szTemp, nPosition - 1)
                    szTemp = Right(szTemp, Len(szTemp) - nPosition)
                End If
        Loop
        If szTemp <> "" Then
            nComma = nComma + 1
            ReDim Preserve aszTemp(1 To nComma)
            aszTemp(nComma) = szTemp
        End If
    End If
    GetIPString = aszTemp
End Function

Public Function GetIPParts(szIP As String) As String()
    Dim aszTemp() As String
    Dim szTemp As String
    Dim nDot As Integer
    Dim nPosition As Integer
    szTemp = szIP
    nDot = 0
    nPosition = 1
    
    If (szTemp = "") Or (szTemp = Null) Then
        ''''
    Else
        Do While nPosition <> 0
            nPosition = InStr(1, szTemp, ".", vbBinaryCompare)
                If nPosition <> 0 Then
                    nDot = nDot + 1
                    ReDim Preserve aszTemp(1 To nDot)
                    aszTemp(nDot) = Left(szTemp, nPosition - 1)
                    szTemp = Right(szTemp, Len(szTemp) - nPosition)
                End If
        Loop
        If szTemp <> "" Then
            nDot = nDot + 1
            ReDim Preserve aszTemp(1 To nDot)
            aszTemp(nDot) = szTemp
        End If
    End If
    GetIPParts = aszTemp
    
    
End Function

'ȡ��Code[Name]��Code
Public Function PartCode(CodeName As String, Optional bBackCode As Boolean = True) As String
    
    Dim nPosition As Integer
    Dim nTemp As Integer
    Dim szTemp As String
    
    nPosition = InStr(1, CodeName, "[")
    If bBackCode = True Then
        If nPosition = 0 Then
            PartCode = CodeName
        Else
            PartCode = Left(CodeName, nPosition - 1)
        End If
    Else
        If nPosition = 0 Then
            PartCode = ""
        Else
            nTemp = Len(CodeName)
            szTemp = Left(CodeName, nTemp - 1)
            PartCode = Right(szTemp, (nTemp - 1 - nPosition))
        End If
    End If
End Function




Public Function NumberText(KeyAscii As Integer, AllText As String, Seltext As String, SelStart As Integer, Optional AllowDot As Boolean = True, Optional AllowNegative As Boolean = True) As Integer
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Or KeyAscii = 45 Then
        If KeyAscii = 46 Then 'Ϊ���(.)
            If AllowDot Then
                If InStr(1, AllText, ".") > 0 And InStr(1, Seltext, ".") = 0 Then KeyAscii = 0
            Else
                KeyAscii = 0
            End If
        ElseIf KeyAscii = 45 Then 'Ϊ����(-)
            If AllowNegative Then
                If InStr(1, AllText, "-") > 0 And InStr(1, Seltext, "-") = 0 Then KeyAscii = 0
                If SelStart <> 0 Then KeyAscii = 0
            Else
                KeyAscii = 0
            End If
        End If
    Else
        KeyAscii = 0
    End If
    NumberText = KeyAscii
End Function
