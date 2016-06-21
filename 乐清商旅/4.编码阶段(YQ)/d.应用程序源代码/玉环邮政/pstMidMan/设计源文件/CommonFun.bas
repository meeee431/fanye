Attribute VB_Name = "CommonFun"
Option Explicit
Option Base 0

Public oBusStation As Object
Public StartName(5) As String
Public BusNetName(5) As String
Public BusNetIP(5) As String
Public UserName(5) As String
Public UserPWD(5) As String
Public StationsStr As String
Public SchedulesStr As String
Public CheckGatesStr As String
Public ConnectedNum As Long
Public tcpok(CONNECTEDMAX) As Integer
Public szReceiveStr(CONNECTEDMAX) As String
Public szSendStr(CONNECTEDMAX) As String
Public tcpok2(CONNECTEDMAX) As Integer
Public Enum E_FileType
    E_BUYTICKETSID = 1
    E_CANCELTICKETSID = 2
    E_INTERNETSELL = 3
    E_INTERNETCANCEL = 4
    E_GetNetTKCOUNT = 5
    E_GetTKCOUNT = 6
    E_GetNetTK = 7
    E_CancelPreNetTK = 8
    E_Error = 9
End Enum
Public ReConnStation() As Integer
Public Function Old2New(OldRet As Long) As Long
Dim TmpRet As Long
   If OldRet > 9999 Then
        TmpRet = OldRet - 10000
    Else
        TmpRet = OldRet
    End If
Select Case OldRet
        Case 12411
        Case 12422
        Case 12423
        Case 12424
        Case 12425
            TmpRet = 3003
        Case 12426
            TmpRet = 3004
        Case 12427
            TmpRet = 3007
        'Case 12428
            
        'Case 12429
        'Case 12430
        Case 12431
            TmpRet = 2017
        'Case 12432
        'Case 12433
         Case 12447
              TmpRet = 2111
             
     End Select
     Old2New = TmpRet


End Function



Public Function ShowErrMsg()
    MsgBox err.Description, vbOKOnly, err.Number
    
End Function
