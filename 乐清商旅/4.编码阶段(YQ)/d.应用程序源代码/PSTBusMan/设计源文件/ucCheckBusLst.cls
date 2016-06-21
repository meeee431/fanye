VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ucCheckBusLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const cnGrowStep = 30        '数组动态增长值

Dim mnMaxArrayLen As Integer       '初始化的数组最大长度
Dim mnCount As Integer
Dim mszBusid As String
Dim matCheckBusLstInfo() As tCheckBusLstInfo

Public Sub Addone(tCheckBusInfo As tCheckBusLstInfo)
    Dim i As Integer
    i = FindItem(tCheckBusInfo.BusID)
    If i = 0 Then
        mnCount = mnCount + 1
        If mnCount > mnMaxArrayLen Then
            mnMaxArrayLen = mnMaxArrayLen + cnGrowStep
            ReDim Preserve matCheckBusLstInfo(1 To mnMaxArrayLen)
        End If
        matCheckBusLstInfo(mnCount) = tCheckBusInfo
    End If
End Sub

Public Sub RemoveAll()
    ReDim matCheckBusLstInfo(1 To mnMaxArrayLen)
    mnCount = 0
End Sub

Public Sub RemoveOne(Index As Integer)
    Dim j As Integer
    For j = Index To mnCount
        matCheckBusLstInfo(j) = matCheckBusLstInfo(j + 1)
    Next j
    mnCount = mnCount - 1
End Sub

Public Property Get Count() As Integer
    Count = mnCount
End Property


Public Sub UpdateOne(tCheckBusInfo As tCheckBusLstInfo)
    Dim i As Integer
    i = FindItem(tCheckBusInfo.BusID)
    If i > 0 Then
        matCheckBusLstInfo(i).BusMode = tCheckBusInfo.BusMode
        matCheckBusLstInfo(i).CheckSheet = tCheckBusInfo.CheckSheet
        matCheckBusLstInfo(i).Company = tCheckBusInfo.Company
        matCheckBusLstInfo(i).EndStationName = tCheckBusInfo.EndStationName
        matCheckBusLstInfo(i).Owner = tCheckBusInfo.Owner
        matCheckBusLstInfo(i).StartupTime = tCheckBusInfo.StartupTime
        matCheckBusLstInfo(i).Status = tCheckBusInfo.Status
        matCheckBusLstInfo(i).Vehicle = tCheckBusInfo.Vehicle
        matCheckBusLstInfo(i).CheckGate = tCheckBusInfo.CheckGate
        matCheckBusLstInfo(i).StopChkTime = tCheckBusInfo.StopChkTime
        matCheckBusLstInfo(i).StartChkTime = tCheckBusInfo.StartChkTime
        matCheckBusLstInfo(i).BusSerial = tCheckBusInfo.BusSerial
    End If
End Sub
Public Property Get Item(Index As Integer) As tCheckBusLstInfo
    Item = matCheckBusLstInfo(Index)
End Property

Private Sub Class_Initialize()
    mnMaxArrayLen = 20
    mnCount = 0
    ReDim matCheckBusLstInfo(1 To mnMaxArrayLen)
End Sub

Public Function FindItem(szBusId As String) As Integer
    Dim i As Integer
    For i = 1 To mnCount
        If matCheckBusLstInfo(i).BusID = szBusId Then
            Exit For
        End If
    Next i
    If i > mnCount Then
        FindItem = 0
    Else
        FindItem = i
    End If
End Function

Public Sub RemoveByGate(szGateId As String)
'删除指定检票口的所有记录
    Dim i As Integer, j As Integer
    Dim nStep As Integer    '步进数
    i = 1
    While i <= mnCount
        If matCheckBusLstInfo(i).CheckGate = szGateId Then
            For j = i + 1 To mnCount
                If matCheckBusLstInfo(j).CheckGate <> szGateId Then Exit For
            Next j
            nStep = j - i
            mnCount = mnCount - nStep
            For j = i To mnCount
                matCheckBusLstInfo(j) = matCheckBusLstInfo(j + nStep)
            Next j
        End If
        i = i + 1
    Wend
End Sub
Public Function FindItemByGate(szGateId As String) As Integer
'按指定检票口查询第一条记录
    Dim i As Integer
    For i = 1 To mnCount
        If matCheckBusLstInfo(i).CheckGate = szGateId Then
            Exit For
        End If
    Next i
    If i > mnCount Then
        FindItemByGate = 0
    Else
        FindItemByGate = i
    End If

End Function
