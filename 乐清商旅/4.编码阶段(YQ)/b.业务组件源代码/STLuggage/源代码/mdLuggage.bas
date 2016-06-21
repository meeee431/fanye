Attribute VB_Name = "mdLuggage"
Option Explicit

Public Const cszLuggage = ""
Public Const cszSplit = ""
Public Const StartFormuarID = "10000"

Public Const cszLuggageCarryAcceptName = "快件"
Public Const cszLuggageNormalAcceptName = "普通"


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
'            GetLuggagePickTypeString = "自提行包"
'        Case 1
'            GetLuggagePickTypeString = "送货行包"
'    End Select
'End Function
'
'Public Function GetLuggagePickTypeInt(szType As String) As Integer
'    Select Case szType
'        Case "自提行包"
'            GetLuggagePickTypeInt = 0
'        Case "送货行包"
'            GetLuggagePickTypeInt = 1
'    End Select
'
'End Function

