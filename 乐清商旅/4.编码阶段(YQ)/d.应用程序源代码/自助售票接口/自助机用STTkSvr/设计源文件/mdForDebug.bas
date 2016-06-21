Attribute VB_Name = "mdForDebug"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long


#If IN_DEBUG Then
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Dim oForDebug As Object
Public Sub OutDebugInfo(pszMsg As String, Optional pszAddMsg As String = "")
    On Error Resume Next
    If oForDebug Is Nothing Then
        Set oForDebug = CreateObject("prjDebug.clsDebug")
    End If
    oForDebug.DebugMsg pszMsg, pszAddMsg, GetCurrentProcessId(), App.ThreadID
End Sub

#Else

Public Sub OutDebugInfo(pszMsg As String, Optional pszAddMsg As String = "")
End Sub

#End If
