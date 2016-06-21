Attribute VB_Name = "mdLuggage"
Option Explicit

Public Const cszLuggage = ""
Public Const cszSplit = ""
Public Const StartFormuarID = "10000"

Public Const cszLuggageCarryAcceptName = "���"
Public Const cszLuggageNormalAcceptName = "��ͨ"


Public Function GetLuggageTypeString(pnType As Integer) As String
    Select Case pnType
        Case 0
            GetLuggageTypeString = cszLuggageCarryAcceptName
        Case 1
            GetLuggageTypeString = cszLuggageNormalAcceptName
    End Select
End Function

Public Function GetLuggageTypeInt(szType As String) As Integer
    Select Case szType
        Case cszLuggageCarryAcceptName
            GetLuggageTypeInt = 0
        Case cszLuggageNormalAcceptName
            GetLuggageTypeInt = 1
    End Select
    
End Function

'
'Public Function GetLuggagePickTypeString(pnType As Integer) As String
'    Select Case pnType
'        Case 0
'            GetLuggagePickTypeString = "�����а�"
'        Case 1
'            GetLuggagePickTypeString = "�ͻ��а�"
'    End Select
'End Function
'
'Public Function GetLuggagePickTypeInt(szType As String) As Integer
'    Select Case szType
'        Case "�����а�"
'            GetLuggagePickTypeInt = 0
'        Case "�ͻ��а�"
'            GetLuggagePickTypeInt = 1
'    End Select
'
'End Function

