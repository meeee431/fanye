VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BusCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const cnGrowStep = 5        '数组动态增长值

Dim mnMaxArrayLen      '初始化的数组最大长度
Dim mnCount As Integer
Dim mszBusID As String
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
    Dim j As Integer, i As Integer
    Dim atCheckInfo() As tCheckBusLstInfo
    ReDim atCheckInfo(mnCount - 1) As tCheckBusLstInfo
    j = 0
    For i = 1 To mnCount
        If i <> Index Then
            j = j + 1
            atCheckInfo(j) = matCheckBusLstInfo(i)
        End If
    Next
    mnCount = mnCount - 1
    matCheckBusLstInfo = atCheckInfo
End Sub

Public Property Get Count() As Integer
    Count = mnCount
End Property


Public Sub UpdateOne(tCheckBusInfo As tCheckBusLstInfo)
    Dim i As Integer
    i = FindItem(tCheckBusInfo.BusID, tCheckBusInfo.BusSerial)
    If i > 0 Then
        matCheckBusLstInfo(i).BusMode = tCheckBusInfo.BusMode
        matCheckBusLstInfo(i).CheckSheet = tCheckBusInfo.CheckSheet
        matCheckBusLstInfo(i).Company = tCheckBusInfo.Company
        matCheckBusLstInfo(i).EndStationName = tCheckBusInfo.EndStationName
        matCheckBusLstInfo(i).Owner = tCheckBusInfo.Owner
        matCheckBusLstInfo(i).StartUpTime = tCheckBusInfo.StartUpTime
        matCheckBusLstInfo(i).Status = tCheckBusInfo.Status
        matCheckBusLstInfo(i).Vehicle = tCheckBusInfo.Vehicle
        matCheckBusLstInfo(i).StartChkTime = tCheckBusInfo.StartChkTime
        matCheckBusLstInfo(i).StopChkTime = tCheckBusInfo.StopChkTime
        matCheckBusLstInfo(i).SplitCompany = tCheckBusInfo.SplitCompany
        matCheckBusLstInfo(i).CheckGate = tCheckBusInfo.CheckGate
    End If
End Sub
Public Property Get Item(Index As Integer) As tCheckBusLstInfo
    Item = matCheckBusLstInfo(Index)
End Property

Private Sub Class_Initialize()
    mnMaxArrayLen = 15
    mnCount = 0
    ReDim matCheckBusLstInfo(1 To mnMaxArrayLen)
End Sub

Public Function FindItem(szBusid As String, Optional pnSerial As Integer) As Integer
    Dim i As Integer
    For i = 1 To mnCount
        If LCase(matCheckBusLstInfo(i).BusID) = LCase(szBusid) And matCheckBusLstInfo(i).BusSerial = pnSerial Then
            Exit For
        End If
    Next i
    If i > mnCount Then
        FindItem = 0
    Else
        FindItem = i
    End If
End Function
