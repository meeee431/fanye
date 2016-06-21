Attribute VB_Name = "mdlCheckIDCard"
' *******************************************************************
' *  Source File Name  :                                            *
' *  Project Name: StationNet V6                                    *
' *  Engineer: 范鹏东                                               *
' *  Date Generated: 2015/11/24                                     *
' *  Last Revision Date :                                           *
' *  Brief Description  : 身份证号码检验模块                        *
' *******************************************************************

Option Explicit

Public gcTmp As Collection
Dim Wi(1 To 18) As Integer     '身-份-证校验码

'身-份-证检验函数,如果有错误,返回的是错误的字符信息,正确,则返回空
Public Function CIDCheck(strId As String) As String
    strId = Trim(strId)
    If Len(strId) = 15 Then
         CIDCheck = CheckCIDC15(strId)
    ElseIf Len(strId) = 18 Then
         CIDCheck = CheckCIDC18(strId)
    Else
         CIDCheck = "位数不对"
    End If
End Function

'15位身-份-证校验
Public Function CheckCIDC15(ByVal StrID15 As String) As String
    Dim sA, sDate
    Dim iA
    
    sA = Mid(StrID15, 7, 2)
    If CInt(sA) > Year(Now) - 2000 Then
       sA = "19" & sA
    Else
       sA = "20" & sA
    End If
    
    sDate = sA & "-" & Mid(StrID15, 9, 2) & "-" & Mid(StrID15, 11, 2)
    sA = Mid(StrID15, 9, 2)
    Set gcTmp = New Collection
    gcTmp.Add sDate
    
    If Not IsNumeric(StrID15) Then
        CheckCIDC15 = "号码输入有误！有非数字出现！"
        Exit Function
    End If
    
    If Val(sA) < 1 Or Val(sA) > 12 Then
        CheckCIDC15 = "号码输入有误！月份不正确！"
        Exit Function
    End If
    
    If Val(Mid(StrID15, 11, 2)) < 1 Or Val(Mid(StrID15, 11, 2)) > 31 Then
        CheckCIDC15 = "号码输入有误！日期不正确！"
        Exit Function
    Else
        If (Val(Mid(StrID15, 9, 2)) = 4 Or Val(Mid(StrID15, 9, 2)) = 6 Or Val(Mid(StrID15, 9, 2)) = 9 Or Val(Mid(StrID15, 9, 2)) = 11) And Val(Mid(StrID15, 11, 2)) = 31 Then
            CheckCIDC15 = "号码输入有误！月份和日期不匹配"
            Exit Function
        ElseIf Val(Mid(StrID15, 9, 2)) = 2 And (Val(Mid(StrID15, 11, 2)) = 30 Or Val(Mid(StrID15, 11, 2)) = 31) Then
            CheckCIDC15 = "号码输入有误！2月份没有" & Val(Mid(StrID15, 11, 2)) & "天"
            Exit Function
        End If
    End If
End Function

'18位身-份-证校验
Public Function CheckCIDC18(ByVal StrID18 As String) As String
    Dim StrID17 As String, AiWi As Integer, Num As Integer, A18 As String
    Dim sA, sDate
    Dim iA
    
    sA = Mid(StrID18, 7, 4)
    sDate = sA & "-" & Mid(StrID18, 11, 2) & "-" & Mid(StrID18, 13, 2)
    Set gcTmp = New Collection
    gcTmp.Add sDate
    s_SetWi
    
    If Not IsNumeric(Left(StrID18, 17)) Then
        CheckCIDC18 = "号码输入有误！"
        Exit Function
    End If
    
    If Val(Mid(StrID18, 11, 2)) < 1 Or Val(Mid(StrID18, 11, 2)) > 12 Then
        CheckCIDC18 = "号码输入有误！月份不正确！"
        Exit Function
    End If
    
    If Val(Mid(StrID18, 13, 2)) < 1 Or Val(Mid(StrID18, 13, 2)) > 31 Then
        CheckCIDC18 = "号码输入有误！" & vbCrLf & "日期不正确！"
        Exit Function
    Else
        If (Val(Mid(StrID18, 11, 2)) = 4 Or Val(Mid(StrID18, 11, 2)) = 6 Or Val(Mid(StrID18, 11, 2)) = 9 Or Val(Mid(StrID18, 11, 2)) = 11) And Val(Mid(StrID18, 13, 2)) = 31 Then
            CheckCIDC18 = "号码输入有误！月份和日期不匹配"
            Exit Function
        ElseIf Val(Mid(StrID18, 11, 2)) = 2 And (Val(Mid(StrID18, 13, 2)) = 30 Or Val(Mid(StrID18, 13, 2)) = 31) Then
            CheckCIDC18 = "号码输入有误！2月份没有" & Val(Mid(StrID18, 13, 2)) & "天"
            Exit Function
        End If
    End If
    
    StrID17 = Left(StrID18, 17)
    AiWi = 0
    
    For Num = 1 To 17
           AiWi = AiWi + Val(Mid(StrID17, Num, 1)) * Wi(Num)
    Next Num
    
    Select Case AiWi Mod 11
        Case 0
            A18 = "1"
        Case 1
            A18 = "0"
        Case 2
            A18 = "X"
        Case 3
            A18 = "9"
        Case 4
            A18 = "8"
        Case 5
            A18 = "7"
        Case 6
            A18 = "6"
        Case 7
            A18 = "5"
        Case 8
            A18 = "4"
        Case 9
            A18 = "3"
        Case 10
            A18 = "2"
    End Select
    
    If A18 <> Right(StrID18, 1) Then
        CheckCIDC18 = "号码输入有误！" '尾数校验码不正确"
        Exit Function
    End If
End Function

'15 升18位算法
Public Function CIDC15To18(ByVal StrID15 As String) As String
    s_SetWi
    Dim StrID17 As String, StrID18 As String, Num As Integer, AiWi As Integer
    
    If Not IsNumeric(StrID15) Then
        CIDC15To18 = "15位号码输入有误！" & vbCrLf & "有非数字出现！"
        Exit Function
    End If
    
    If Val(Mid(StrID15, 9, 2)) < 1 Or Val(Mid(StrID15, 9, 2)) > 12 Then
        CIDC15To18 = "号码输入有误！" & vbCrLf & "月份不正确！"
        Exit Function
    End If
    
    If Val(Mid(StrID15, 11, 2)) < 1 Or Val(Mid(StrID15, 11, 2)) > 31 Then
        CIDC15To18 = "号码输入有误！" & vbCrLf & "日期不正确！"
        Exit Function
    Else
        If (Val(Mid(StrID15, 9, 2)) = 4 Or Val(Mid(StrID15, 9, 2)) = 6 Or Val(Mid(StrID15, 9, 2)) = 9 Or Val(Mid(StrID15, 9, 2)) = 11) And Val(Mid(StrID15, 11, 2)) = 31 Then
            CIDC15To18 = "号码输入有误！" & vbCrLf & "月份和日期不匹配"
            Exit Function
        ElseIf Val(Mid(StrID15, 9, 2)) = 2 And (Val(Mid(StrID15, 11, 2)) = 30 Or Val(Mid(StrID15, 11, 2)) = 31) Then
            CIDC15To18 = "号码输入有误！" & vbCrLf & "2月份没有" & Val(Mid(StrID15, 11, 2)) & "天"
            Exit Function
        End If
    End If
    
    StrID17 = Left(StrID15, 6) & "19" & Right(StrID15, 9)
    AiWi = 0
    
    For Num = 1 To 17
           AiWi = AiWi + Val(Mid(StrID17, Num, 1)) * Wi(Num)
    Next Num
    
    Select Case AiWi Mod 11
        Case 0
            StrID18 = StrID17 & "1"
        Case 1
            StrID18 = StrID17 & "0"
        Case 2
            StrID18 = StrID17 & "X"
        Case 3
            StrID18 = StrID17 & "9"
        Case 4
            StrID18 = StrID17 & "8"
        Case 5
            StrID18 = StrID17 & "7"
        Case 6
            StrID18 = StrID17 & "6"
        Case 7
            StrID18 = StrID17 & "5"
        Case 8
            StrID18 = StrID17 & "4"
        Case 9
            StrID18 = StrID17 & "3"
        Case 10
            StrID18 = StrID17 & "2"
    End Select
    
    CIDC15To18 = StrID18
End Function

'*******************************
'名称     fun_checkIDCard
'功能     检查-身-份-证的合法性
'返回     true,false
'*******************************
Private Sub s_SetWi()
    Wi(1) = 7
    Wi(2) = 9
    Wi(3) = 10
    Wi(4) = 5
    Wi(5) = 8
    Wi(6) = 4
    Wi(7) = 2
    Wi(8) = 1
    Wi(9) = 6
    Wi(10) = 3
    Wi(11) = 7
    Wi(12) = 9
    Wi(13) = 10
    Wi(14) = 5
    Wi(15) = 8
    Wi(16) = 4
    Wi(17) = 2
    Wi(18) = 1
End Sub
