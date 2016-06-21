Attribute VB_Name = "mdSellCtl"
Option Explicit

Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_KEYSCANCODE_TAB = &HF0001
Public Const VK_TAB = &H9

Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Public Sub SendTab()
    PressKey GetFocus, WM_KEYSCANCODE_TAB, VK_TAB
End Sub
Public Function SepByChar(ByRef ss As String, CompareChar As String) As String
    SepByChar = GetLeft(ss, CompareChar)
    ss = Mid(ss, Len(SepByChar) + 2)
End Function

Public Function GetLeft(ss As String, CompareChar As String) As String
Dim nPosition As Integer
    nPosition = InStr(ss, Left(CompareChar, 1)) - 1
    If nPosition <= 0 Then
        nPosition = Len(ss)
    End If
    GetLeft = Left(ss, nPosition)
End Function
Public Sub PressKey(hwnd As Long, KeyScanCode As Long, VirtualKey As Integer)
    If hwnd <> 0 Then
        PostMessage hwnd, WM_KEYDOWN, VirtualKey, KeyScanCode
'        PostMessage hwnd, WM_CHAR, VirtualKey, &HF0001
            PostMessage hwnd, WM_KEYUP, VirtualKey, &HC0000000 Or KeyScanCode
    End If
End Sub
