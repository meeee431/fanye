Attribute VB_Name = "Print"
Option Explicit


Const cnStartStation1 = 1
Const cnEndStation1 = 2
Const cnTicketType1 = 3
Const cnTicketPrice1 = 4
Const cnStartStation2 = 5
Const cnEndStation2 = 6
Const cnTicketPrice2 = 7
Const cnTicketType2 = 8
Const cnUserID = 9
Const cnBusDate1 = 10
Const cnBusID1 = 11
Const cnBusDate2 = 12
Const cnBusID2 = 13
Const cnSeat = 14
Const cnStartupTime = 15
Const cnVehicleType = 16
Const cnCheckGate = 17
Const cnTicketNO = 18
Const cnUserID2 = 19
Const cnTicketNO1 = 20
Const cnSeat2 = 21
Const cnCompany = 22
Const cnInsurance = 23

'返回打印头
'嘉兴原来是22，配置文件里要改
Const cnReturnPoint = 24

Const clRTicketID = 1
Const clRTicketPrice = 2
Const clRReturnCharge = 3
Const clRSheetID = 4
Const clRStartUpTime = 5
Const clRUserID = 6
Const clRBusID = 7
Const clREndStation = 8
Const clRNowDate = 9
Const clRTicketType = 10
Const clREndStation2 = 11
Const clRTicketType2 = 12
Const clRStartUpTime2 = 13
Const clRTicketCount = 14
Const clRTicketCount2 = 15
Const clRReturnCharge2 = 16
Const clRReturnCountStr = 17
Const clRReturnChargeStr = 18
Const clRTicketID2 = 19
Const clRBusID2 = 20
Const clRVehicleType = 21
Const clRSeatNo = 22
Const clRNowDate2 = 23
Const clRStartStation1 = 24
Const clRStartStation2 = 25
Const clRReturnChargeStr2 = 26
Const clRBusDate1 = 27
Const clRBusDate2 = 28
Const clRSheetIDStr = 29
Const clRReturnTimeStr = 32
Const clRUserID1 = 30
Const clRSeatNo1 = 31

#If PRINT_SHEET <> 0 Then
    Public m_oFastReturn As New FastPrint
    Public m_oFastPrint As New FastPrint
    
    Public m_oPrintTicket As New BPrint
    Public m_oPrintReturn As New BPrint
#End If
''
''#If PRINT_SHEET <> 0 Then
'''    Public m_oPrintTicket As BPrint
''    Public m_oPrintReturn As New FastPrint
''    Public m_oFastPrint As New FastPrint
''#End If

'打印退票凭证
Public Function PrintReturnSheet(pszTicketNo As String, pszSheetID As String, psgReturnCharge As Double, psgTicketPrice As Double, pszOpTime As String, pszBusID As String, pszEndStation As String, pnReturnCount As String) As Long
 #If PRINT_SHEET <> 0 Then
On Error GoTo Error_Handle
    If m_bUseFastPrint Then     '如果是慢速打印，调用billPrint
        PrintReturnSheetFast pszTicketNo, pszSheetID, psgReturnCharge, psgTicketPrice, pszOpTime, pszBusID, pszEndStation, pnReturnCount
    Else
        PrintReturnSheetSlow pszTicketNo, pszSheetID, psgReturnCharge, psgTicketPrice, pszOpTime, pszBusID, pszEndStation, pnReturnCount
    End If
    Exit Function
Error_Handle:
    ShowErrorMsg
#End If
End Function
Private Function PrintReturnSheetFast(pszTicketNo As String, pszSheetID As String, psgReturnCharge As Double, psgTicketPrice As Double, pszOpTime As String, pszBusID As String, pszEndStation As String, pnReturnCount As String) As Long
    m_oFastReturn.ClosePort
    m_oFastReturn.OpenPort
    m_oFastReturn.SetObject clRTicketID
    m_oFastReturn.SetCaption pszTicketNo   '打印当前票号

    m_oFastReturn.SetObject clRTicketPrice  '打印当前票价
    m_oFastReturn.SetCaption pszTicketNo

    m_oFastReturn.SetObject clRReturnCharge   '退票手续费
    m_oFastReturn.SetCaption Format(psgReturnCharge, "0.00") & "元"

    m_oFastReturn.SetObject clRSheetID
    m_oFastReturn.SetCaption "退票收据"

    m_oFastReturn.SetObject clRNowDate    '打印日期
    m_oFastReturn.SetCaption Format(Date, "YYYY-MM-DD")

    m_oFastReturn.SetObject clRUserID  '打印工号
    m_oFastReturn.SetCaption GetActiveUserID

    m_oFastReturn.SetObject clRNowDate2
    m_oFastReturn.SetCaption Format(Date, "YYYY-MM-DD")

    m_oFastReturn.SetObject clRTicketCount  '打印票的张数
    m_oFastReturn.SetCaption pnReturnCount & "张"

    m_oFastReturn.SetObject clRTicketCount2  '打印票的张数
    m_oFastReturn.SetCaption pnReturnCount & "张"

    m_oFastReturn.SetObject clRReturnCharge2   '退票手续费
    m_oFastReturn.SetCaption Format(psgReturnCharge, "0.00") & "元"

    m_oFastReturn.SetObject clRReturnCountStr  '打印退票张数字符串
    m_oFastReturn.SetCaption "退票张数"

    m_oFastReturn.SetObject clRReturnChargeStr  '打印手续费字符串
    m_oFastReturn.SetCaption "手续费"

    m_oFastReturn.SetObject clREndStation '打印终点站
    m_oFastReturn.SetCaption pszEndStation

    m_oFastReturn.PrintFile

    m_oFastReturn.ClosePort

End Function
Private Function PrintReturnSheetSlow(pszTicketNo As String, pszSheetID As String, psgReturnCharge As Double, psgTicketPrice As Double, pszOpTime As String, pszBusID As String, pszEndStation As String, pnReturnCount As String) As Long
    m_oPrintReturn.SetCurrentObject clRTicketID
    m_oPrintReturn.LabelSetCaption ""   '打印当前票号

    m_oPrintReturn.SetCurrentObject clRStartStation1  '打印起点站
    m_oPrintReturn.LabelSetCaption ""
    
    m_oPrintReturn.SetCurrentObject clREndStation '打印终点站
    m_oPrintReturn.LabelSetCaption ""   'pszEndStation
    
    m_oPrintReturn.SetCurrentObject clRTicketPrice  '打印当前票价
    m_oPrintReturn.LabelSetCaption Format(psgTicketPrice, "0.00")

    m_oPrintReturn.SetCurrentObject clRReturnCharge   '退票手续费
    m_oPrintReturn.LabelSetCaption Format(psgReturnCharge, "0.00") & "元"

    m_oPrintReturn.SetCurrentObject clRSheetID
    m_oPrintReturn.LabelSetCaption pszSheetID

    m_oPrintReturn.SetCurrentObject clRNowDate    '打印日期
    m_oPrintReturn.LabelSetCaption Format(Now, "YYYY-MM-DD HH:mm")

    m_oPrintReturn.SetCurrentObject clRUserID  '打印工号
    m_oPrintReturn.LabelSetCaption GetActiveUserID

'    m_oPrintReturn.SetCurrentObject clRNowDate2
'    m_oPrintReturn.LabelSetCaption Format(Date, "YYYY-MM-DD")

'    m_oPrintReturn.SetCurrentObject clRReturnCharge2   '退票手续费
'    m_oPrintReturn.LabelSetCaption Format(psgReturnCharge, "0.00") & "元"

    '不显示单票车次、发车时间等
    m_oPrintReturn.SetCurrentObject clRBusID  '
    m_oPrintReturn.LabelSetCaption ""
    m_oPrintReturn.SetCurrentObject clRStartUpTime
    m_oPrintReturn.LabelSetCaption ""
    
    
    '打印退票张数
    m_oPrintReturn.SetCurrentObject clRReturnCountStr
    m_oPrintReturn.LabelSetCaption "退票张数:"
    m_oPrintReturn.SetCurrentObject clRTicketCount  '打印票的张数
    m_oPrintReturn.LabelSetCaption pnReturnCount & "张"
    m_oPrintReturn.SetCurrentObject clRTicketCount2  '打印票的张数
    m_oPrintReturn.LabelSetCaption pnReturnCount & "张"

'    m_oPrintReturn.SetCurrentObject clRReturnChargeStr  '打印手续费字符串
'    m_oPrintReturn.LabelSetCaption "手续费"


    m_oPrintReturn.PrintA 100

End Function
'打印退票凭证
Public Function PrintSingleReturnSheet(pszTicketNo As String, pszSheetID As String, psgReturnCharge As Double, psgTicketPrice As Double, pdtStartUpTime As String _
    , pszBusID As String, pszEndStation As String, pszTicketType As String, pszVehicleType As String, pszSeatNo As String, pszStartStation As String _
    , pdtBusDate As Date) As Long
 #If PRINT_SHEET <> 0 Then
On Error GoTo Error_Handle
    If m_bUseFastPrint Then     '如果是慢速打印，调用billPrint
        PrintSingleReturnSheetFast pszTicketNo, pszSheetID, psgReturnCharge, psgTicketPrice, pdtStartUpTime, pszBusID, pszEndStation, pszTicketType, pszVehicleType, pszSeatNo, pszStartStation, pdtBusDate
    Else
        PrintSingleReturnSheetSlow pszTicketNo, pszSheetID, psgReturnCharge, psgTicketPrice, pdtStartUpTime, pszBusID, pszEndStation, pszTicketType, pszVehicleType, pszSeatNo, pszStartStation, pdtBusDate
    End If
    Exit Function
Error_Handle:
    ShowErrorMsg
    #End If
End Function
'''打印退票凭证
''Public Function PrintSingleReturnSheet(pszTicketNo As String, psgReturnCharge As Double, psgTicketPrice As Double, pdtStartUpTime As String _
''    , pszBusID As String, pszEndStation As String, pszTicketType As String, pszVehicleType As String, pszSeatNo As String, pszStartStation As String _
''    , pdtBusDate As Date) As Long
''#If PRINT_SHEET <> 0 Then
''
''    m_oPrintReturn.ClosePort
''    m_oPrintReturn.OpenPort
''
''    m_oPrintReturn.SetObject clRTicketID
''    m_oPrintReturn.SetCaption pszTicketNo   '打印当前票号
''
''    m_oPrintReturn.SetObject clRTicketPrice   '票价
''    m_oPrintReturn.SetCaption Format(psgTicketPrice, "0.00") & "元"
''
''    m_oPrintReturn.SetObject clRReturnCharge   '退票手续费
''    m_oPrintReturn.SetCaption Format(psgReturnCharge, "0.00") & "元"
''
''    m_oPrintReturn.SetObject clRSheetID
''    m_oPrintReturn.SetCaption "退票"
''
''    m_oPrintReturn.SetObject clRStartUpTime   '发车时间
''    m_oPrintReturn.SetCaption Format(pdtStartUpTime, "hh:mm")
''
''    m_oPrintReturn.SetObject clRUserID  '打印工号
''    m_oPrintReturn.SetCaption GetActiveUserID
''
''    m_oPrintReturn.SetObject clRBusID  '
''    m_oPrintReturn.SetCaption pszBusID
''
''    m_oPrintReturn.SetObject clREndStation  '
''    m_oPrintReturn.SetCaption pszEndStation
''
''    m_oPrintReturn.SetObject clRNowDate   '打印日期
''    m_oPrintReturn.SetCaption Format(Date, "YYYY-MM-DD")
''
''    m_oPrintReturn.SetObject clRTicketType  '
''    m_oPrintReturn.SetCaption pszTicketType
''
''    m_oPrintReturn.SetObject clREndStation2  '
''    m_oPrintReturn.SetCaption pszEndStation
''
''    m_oPrintReturn.SetObject clRTicketType2  '
''    m_oPrintReturn.SetCaption pszTicketType
''
''    m_oPrintReturn.SetObject clRStartUpTime2  '发车时间
''    m_oPrintReturn.SetCaption Format(pdtStartUpTime, "hh:mm")
''
''    m_oPrintReturn.SetObject clRTicketCount  '
''    m_oPrintReturn.SetCaption "1张"
''
''    m_oPrintReturn.SetObject clRTicketCount2  '
''    m_oPrintReturn.SetCaption "1张"
''
''    m_oPrintReturn.SetObject clRReturnCharge2  '
''    m_oPrintReturn.SetCaption Format(psgReturnCharge, "0.00") & "元"
''
''    m_oPrintReturn.SetObject clRReturnCountStr  '
''    m_oPrintReturn.SetCaption "退票张数"
''
''    m_oPrintReturn.SetObject clRReturnChargeStr  '
''    m_oPrintReturn.SetCaption "手续费"
''
''    m_oPrintReturn.SetObject clRTicketID2  '
''    m_oPrintReturn.SetCaption pszTicketNo
''
''    m_oPrintReturn.SetObject clRBusID2  '
''    m_oPrintReturn.SetCaption pszBusID
''
''    m_oPrintReturn.SetObject clRVehicleType  '
''    m_oPrintReturn.SetCaption pszVehicleType
''
''
''    m_oPrintReturn.SetObject clRSeatNo '
''    m_oPrintReturn.SetCaption pszSeatNo
''
''    m_oPrintReturn.SetObject clRStartStation1
''    m_oPrintReturn.SetCaption pszStartStation
''
''    m_oPrintReturn.SetObject clRStartStation2
''    m_oPrintReturn.SetCaption pszStartStation
''
''    m_oPrintReturn.SetObject clRReturnChargeStr2
''    m_oPrintReturn.SetCaption "手续费"
''
''    m_oPrintReturn.SetObject clRBusDate1
''    m_oPrintReturn.SetCaption Format(pdtBusDate, "yyyy-mm-dd")
''
''    m_oPrintReturn.SetObject clRBusDate2
''    m_oPrintReturn.SetCaption Format(pdtBusDate, "yyyy-mm-dd")
''
''    m_oPrintReturn.PrintFile
''
''    m_oPrintReturn.ClosePort
''
''#End If
''
''End Function

Public Sub GetIniFile()

#If PRINT_SHEET <> 0 Then
'            Set m_oPrintTicket = New BPrint
    
On Error GoTo Error_Handle
    If Not m_bUseFastPrint Then     '如果是慢速打印，调用billPrint
        If FileIsExist(App.Path & "\SellTk.bpf") Then
            m_oPrintTicket.ReadFormatFileA App.Path & "\SellTk.bpf"
        Else
            MsgBox "打印配置文件""SellTk.bpf""未找到,无法运行售票程序", vbCritical
            End
        End If
        
        If FileIsExist(App.Path & "\ReturnTk.bpf") Then
            m_oPrintReturn.ReadFormatFileA App.Path & "\ReturnTk.bpf"
        Else
            MsgBox "打印配置文件""ReturnTk.bpf""未找到,无法运行售票程序", vbCritical
            End
        End If
    Else
        If FileIsExist(App.Path & "\SellTk.ini") Then
            m_oFastPrint.ReadFormatFile App.Path & "\SellTk.ini"
        Else
            MsgBox "打印配置文件""SellTk.ini""未找到,无法运行售票程序", vbCritical
            End
        End If
        
        If FileIsExist(App.Path & "\ReturnTk.ini") Then
            m_oFastReturn.ReadFormatFile App.Path & "\ReturnTk.ini"
        Else
            MsgBox "打印配置文件""ReturnTk.ini""未找到,无法运行售票程序", vbCritical
            End
        End If
    End If
    
    Exit Sub
Error_Handle:
    ShowErrorMsg
#End If
        
End Sub


'打印票函数
'szEndStation As String
'dtOffTime As Date
'szBusID As String
'szVehicleType As String
'szCheckGate As String
'bSaleChange As Boolean

Public Function PrintTicket(aptTicketInfo() As TPrintTicketParam, szBusDate() As String, nTicketCount() As Integer, szEndStation() As String _
, szOffTime() As String, szBusID() As String, szVehicleType() As String, szCheckGate() As String, bSaleChange() As Boolean, aszTerminateName() As String _
, szSellStationName() As String, panInsurance() As Integer) As Long

 #If PRINT_SHEET <> 0 Then
'    Close #1
'    Open "lpt1:" For Output As #1
    Dim i As Integer
    Dim szDateTemp As String
    Dim dtTemp As Date
    Dim bPrintAid As Boolean
    Dim bPrintBarCode As Boolean
    Dim iLen As Integer
    Dim nCount As Integer
    
'    Dim oParam As SystemParam
    
    Dim dScrollTime As Date
    Dim szTicketType As String
    Dim szRouteName As String
    
    iLen = 0
    On Error GoTo Error_Handle
    

    '打印变量声明
    Dim szStationName As String
    Dim szOffTime1 As String
    Dim szCheckGate1 As String
    Dim szTicketPrice As String
    Dim szTicketType1 As String
    Dim szBusID1 As String
    Dim szSeatNo As String
    Dim szDate As String
    Dim szUserID As String
    Dim aszRouteAndTime() As String
    Dim szScrollTime As String
    Dim szStartStationName As String
    Dim szTicketNo As String
    
    Dim szTermination As String
    Dim szVehicleModel As String
    
    Dim szStartShortName As String
    
    
    Dim szCompanyID As String
    Dim szCompanyName As String
    
    
    Dim szInsurance As String
    
    If m_bUseFastPrint Then
        m_oFastPrint.ClosePort
        m_oFastPrint.OpenPort
    End If
    
    iLen = ArrayLength(szBusID)
'    If oParam Is Nothing Then Set oParam = m_oParam
    For nCount = 1 To iLen
        dtTemp = CDate(szBusDate(nCount))
        szDateTemp = Format(dtTemp, "YYYY") & "-" & Format(dtTemp, "MM") & "-" & Format(dtTemp, "DD")
'        aszRouteAndTime = m_oSell.GetRouteAndTime(CDate(szBusDate(nCount)), szBusID(nCount))
'
'        If ArrayLength(aszRouteAndTime) <> 0 Then
'            szRouteName = aszRouteAndTime(1)
'            szScrollTime = aszRouteAndTime(2)
'            szCompanyID = aszRouteAndTime(3)
'            szCompanyName = aszRouteAndTime(4)
'        Else
'            szRouteName = ""
'            szScrollTime = ""
'            szCompanyID = ""
'            szCompanyName = ""
'
'        End If
'
        For i = 1 To nTicketCount(nCount)
        
            szTicketNo = aptTicketInfo(nCount).aptPrintTicketInfo(i).szTicketNo  '得到票号
'            #If PRINT_SELLSTATION = 1 Then
                szStartStationName = szSellStationName(nCount)
'            #Else
'                '如果打印的起点站是上车站,否则
'                szStartStationName = m_oSell.SellUnitShortName  '得到起点站名
'                szStartStationName = Left(szStartStationName, Len(szStartStationName) - 1)
'                szStartShortName = Right(szStartStationName, 1)
'            #End If
            
            szStationName = szEndStation(nCount)  '得到站名
    
            szTicketPrice = CStr(Format(aptTicketInfo(nCount).aptPrintTicketInfo(i).sgTicketPrice, "0.00")) '得到票价
            
            szTicketType = Trim(GetTicketTypeStr2(aptTicketInfo(nCount).aptPrintTicketInfo(i).nTicketType))

            szTicketType1 = Left(szTicketType, Len(szTicketType) - 1) '得到票种
            
            szUserID = GetActiveUserID   '得到工号
            If bSaleChange(nCount) Then szUserID = szUserID & "[改]"
            
            
            szDate = szDateTemp   '得到发车日期
            If Format(szOffTime(nCount), "hh:mm") = cszScrollBus Then
                szBusID1 = GetBusID(szBusID(nCount)) 'szRouteName  '得到车次
            Else
                szBusID1 = GetBusID(szBusID(nCount))   '得到车次
            End If
            
            szVehicleModel = szVehicleType(nCount) '得到车型
            
            szTermination = Trim(aszTerminateName(nCount)) '得到车次终点站名称

            szSeatNo = aptTicketInfo(nCount).aptPrintTicketInfo(i).szSeatNo '得到座位号
    
'            m_oPrintTicket.Setobject clStartupTime
            If GetPrintScrollMode Then
                If Format(szOffTime(nCount), "hh:mm") = cszScrollBus Then
                    '如果为滚动班次,
                    '当11:30分之前买时,打印出来的票上为11:30分之前
'                    否则打印车次时间加之前
                    
                    If DateDiff("s", Time, cszMiddleTime) > 0 Then
                        szOffTime1 = cszMiddleTime & cszScrollBusTime
                    Else
                        szOffTime1 = CStr(Format(szScrollTime, "hh:mm")) & cszScrollBusTime  '得到发车时间
                    End If
                Else
                    szOffTime1 = CStr(Format(szOffTime(nCount), "hh:mm"))   '得到发车时间
                End If
            Else
                If Format(szOffTime(nCount), "hh:mm") = cszScrollBus Then
                
                    szOffTime1 = cszScrollBus  '得到发车时间
                Else
                    szOffTime1 = CStr(Format(szOffTime(nCount), "hh:mm"))
                End If
            End If

            szCheckGate1 = szCheckGate(nCount)       '得到检票口
            Select Case Len(Trim(szCheckGate1))
                Case 1
                    szCheckGate1 = szCheckGate1 & Space(6)
                Case 2
                    szCheckGate1 = szCheckGate1 & Space(5)
                Case 3
                    szCheckGate1 = szCheckGate1 & Space(1)
            End Select
            
            
            
            szInsurance = IIf(panInsurance(nCount) = 0, "", "[保]")
            
            
            ''开始打印--------------------------------------------------------
            If m_bUseFastPrint Then
                
                m_oFastPrint.SetObject cnStartStation1
                m_oFastPrint.SetCaption szStartStationName
                
                m_oFastPrint.SetObject cnEndStation1
                m_oFastPrint.SetCaption szStationName
                
                m_oFastPrint.SetObject cnTicketType1
                m_oFastPrint.SetCaption szTicketType1
                
                m_oFastPrint.SetObject cnTicketPrice1
                m_oFastPrint.SetCaption szTicketPrice
                
                m_oFastPrint.SetObject cnStartStation2
                m_oFastPrint.SetCaption szStartStationName
                
                m_oFastPrint.SetObject cnEndStation2
                m_oFastPrint.SetCaption szStationName
                
                m_oFastPrint.SetObject cnTicketPrice2
                m_oFastPrint.SetCaption szTicketPrice
                
                m_oFastPrint.SetObject cnTicketType2
                m_oFastPrint.SetCaption szTicketType1
                
                m_oFastPrint.SetObject cnUserID
                m_oFastPrint.SetCaption szUserID
                
                
                m_oFastPrint.SetObject cnUserID2
                m_oFastPrint.SetCaption szUserID
                
                
                
                m_oFastPrint.SetObject cnBusDate1
                m_oFastPrint.SetCaption szDate
                
                m_oFastPrint.SetObject cnBusID1
                m_oFastPrint.SetCaption szBusID1
                
                m_oFastPrint.SetObject cnBusDate2
                m_oFastPrint.SetCaption szDate
                
                m_oFastPrint.SetObject cnBusID2
                m_oFastPrint.SetCaption szBusID1
                
                m_oFastPrint.SetObject cnSeat
                m_oFastPrint.SetCaption szSeatNo
                
                m_oFastPrint.SetObject cnSeat2
                m_oFastPrint.SetCaption szSeatNo
                
                
                
                m_oFastPrint.SetObject cnStartupTime
                m_oFastPrint.SetCaption szOffTime1
                
                m_oFastPrint.SetObject cnVehicleType
                m_oFastPrint.SetCaption szVehicleModel
                
                m_oFastPrint.SetObject cnCheckGate
                m_oFastPrint.SetCaption szCheckGate1
                
                m_oFastPrint.SetObject cnTicketNO
                m_oFastPrint.SetCaption Right(szTicketNo, 4)
                
                m_oFastPrint.SetObject cnTicketNO1
                m_oFastPrint.SetCaption szTicketNo
                
                m_oFastPrint.SetObject cnCompany
                m_oFastPrint.SetCaption szCompanyName
                
                '保险
                m_oFastPrint.SetObject cnInsurance
                m_oFastPrint.SetCaption szInsurance
                
                '打印点，为返回打印头
                m_oFastPrint.SetObject cnReturnPoint
                m_oFastPrint.SetCaption "..."
                
                m_oFastPrint.PrintFile

            Else
                m_oPrintTicket.SetCurrentObject cnStartStation1
                m_oPrintTicket.LabelSetCaption szStartStationName
                
                m_oPrintTicket.SetCurrentObject cnEndStation1
                m_oPrintTicket.LabelSetCaption szStationName
                
                m_oPrintTicket.SetCurrentObject cnTicketType1
                m_oPrintTicket.LabelSetCaption szTicketType1
                
                m_oPrintTicket.SetCurrentObject cnTicketPrice1
                m_oPrintTicket.LabelSetCaption szTicketPrice
                
                m_oPrintTicket.SetCurrentObject cnStartStation2
                m_oPrintTicket.LabelSetCaption szStartStationName
                
                m_oPrintTicket.SetCurrentObject cnEndStation2
                m_oPrintTicket.LabelSetCaption szStationName
                
                m_oPrintTicket.SetCurrentObject cnTicketPrice2
                m_oPrintTicket.LabelSetCaption szTicketPrice
                
                m_oPrintTicket.SetCurrentObject cnTicketType2
                m_oPrintTicket.LabelSetCaption szTicketType1
                
                m_oPrintTicket.SetCurrentObject cnUserID
                m_oPrintTicket.LabelSetCaption szUserID
                
                
                m_oPrintTicket.SetCurrentObject cnUserID2
                m_oPrintTicket.LabelSetCaption szUserID
                
                
                m_oPrintTicket.SetCurrentObject cnBusDate1
                m_oPrintTicket.LabelSetCaption szDate
                
                m_oPrintTicket.SetCurrentObject cnBusID1
                m_oPrintTicket.LabelSetCaption szBusID1
                
                m_oPrintTicket.SetCurrentObject cnBusDate2
                m_oPrintTicket.LabelSetCaption szDate
                
                m_oPrintTicket.SetCurrentObject cnBusID2
                m_oPrintTicket.LabelSetCaption szBusID1
                
                
                m_oPrintTicket.SetCurrentObject cnSeat
                m_oPrintTicket.LabelSetCaption szSeatNo
                
                m_oPrintTicket.SetCurrentObject cnSeat2
                m_oPrintTicket.LabelSetCaption szSeatNo
                
                
                m_oPrintTicket.SetCurrentObject cnStartupTime
                m_oPrintTicket.LabelSetCaption szOffTime1
                
                m_oPrintTicket.SetCurrentObject cnVehicleType
                m_oPrintTicket.LabelSetCaption szVehicleModel
                
                m_oPrintTicket.SetCurrentObject cnCheckGate
                m_oPrintTicket.LabelSetCaption szCheckGate1
                
                m_oPrintTicket.SetCurrentObject cnTicketNO
                m_oPrintTicket.LabelSetCaption Right(szTicketNo, 4)
                
                m_oPrintTicket.SetCurrentObject cnTicketNO1
                m_oPrintTicket.LabelSetCaption szTicketNo
                
                m_oPrintTicket.SetCurrentObject cnInsurance
                m_oPrintTicket.LabelSetCaption szInsurance
                
                '打印点，为返回打印头
                m_oPrintTicket.SetCurrentObject cnReturnPoint
                m_oPrintTicket.LabelSetCaption "..."
                
                m_oPrintTicket.PrintA 100
            End If
            
            
        Next
    Next nCount
    If m_bUseFastPrint Then
        m_oFastPrint.ClosePort
    End If
'    Set oParam = Nothing
    Exit Function
Error_Handle:
    ShowErrorMsg
'    Set oParam = Nothing
    #End If
    
End Function
'打印退票凭证
Public Function PrintSingleReturnSheetSlow(pszTicketNo As String, pszSheetID As String, psgReturnCharge As Double, psgTicketPrice As Double, pdtStartUpTime As String _
    , pszBusID As String, pszEndStation As String, pszTicketType As String, pszVehicleType As String, pszSeatNo As String, pszStartStation As String _
    , pdtBusDate As Date) As Long
#If PRINT_SHEET <> 0 Then
    m_oPrintReturn.SetCurrentObject clRStartStation1
    m_oPrintReturn.LabelSetCaption pszStartStation
    
    m_oPrintReturn.SetCurrentObject clREndStation  '
    m_oPrintReturn.LabelSetCaption pszEndStation
    
    m_oPrintReturn.SetCurrentObject clRUserID  '打印工号
    m_oPrintReturn.LabelSetCaption GetActiveUserID
    
'    m_oPrintReturn.SetCurrentObject clRReturnChargeStr  '
'    m_oPrintReturn.LabelSetCaption "退票手续费:"
    
    m_oPrintReturn.SetCurrentObject clRReturnCharge   '退票手续费
    m_oPrintReturn.LabelSetCaption Format(psgReturnCharge, "0.00") & "元"

'    m_oPrintReturn.SetCurrentObject clRReturnTimeStr   '
'    m_oPrintReturn.LabelSetCaption "退票时间:"

    m_oPrintReturn.SetCurrentObject clRNowDate   '打印日期
    m_oPrintReturn.LabelSetCaption Format(Now, "YYYY-MM-DD HH:mm")
    
    m_oPrintReturn.SetCurrentObject clRTicketID
    m_oPrintReturn.LabelSetCaption pszTicketNo   '打印当前票号
     
'    m_oPrintReturn.SetCurrentObject clRSheetIDStr
'    m_oPrintReturn.LabelSetCaption "凭证号:"
    
    m_oPrintReturn.SetCurrentObject clRSheetID
    m_oPrintReturn.LabelSetCaption pszSheetID

    '显示单票车次、发车时间等
    m_oPrintReturn.SetCurrentObject clRBusID  '
    m_oPrintReturn.LabelSetCaption "(" & pszBusID & ")"
    m_oPrintReturn.SetCurrentObject clRStartUpTime   '发车时间
    m_oPrintReturn.LabelSetCaption pdtStartUpTime      'Format(pdtStartUpTime, "YYYY-MM-DD HH:mm")

    m_oPrintReturn.SetCurrentObject clRTicketPrice   '票价
    m_oPrintReturn.LabelSetCaption Format(psgTicketPrice, "0.00")
    
    '不显示退票张数
    m_oPrintReturn.SetCurrentObject clRReturnCountStr
    m_oPrintReturn.LabelSetCaption ""
    m_oPrintReturn.SetCurrentObject clRTicketCount
    m_oPrintReturn.LabelSetCaption ""
'    m_oPrintReturn.SetCurrentObject clRTicketCount2
'    m_oPrintReturn.LabelSetCaption ""
    

    '嘉兴为了打点
    m_oPrintReturn.SetCurrentObject cnReturnPoint
    m_oPrintReturn.LabelSetCaption "..."
    
    m_oPrintReturn.PrintA 100
    
#End If

End Function

'打印退票凭证
Private Function PrintSingleReturnSheetFast(pszTicketNo As String, pszSheetID As String, psgReturnCharge As Double, psgTicketPrice As Double, pdtStartUpTime As String _
    , pszBusID As String, pszEndStation As String, pszTicketType As String, pszVehicleType As String, pszSeatNo As String, pszStartStation As String _
    , pdtBusDate As Date) As Long
#If PRINT_SHEET <> 0 Then
    m_oFastReturn.ClosePort
    m_oFastReturn.OpenPort
    
    m_oFastReturn.SetObject clRTicketID
    m_oFastReturn.SetCaption pszTicketNo   '打印当前票号
    
    m_oFastReturn.SetObject clRTicketPrice   '票价
    m_oFastReturn.SetCaption Format(psgTicketPrice, "0.00") & "元"

    m_oFastReturn.SetObject clRReturnCharge   '退票手续费
    m_oFastReturn.SetCaption Format(psgReturnCharge, "0.00") & "元"

    m_oFastReturn.SetObject clRSheetIDStr
'    m_oFastReturn.SetCaption "退票凭证号:"
    
    m_oFastReturn.SetObject clRSheetID
    m_oFastReturn.SetCaption pszSheetID

    m_oFastReturn.SetObject clRStartUpTime   '发车时间
    m_oFastReturn.SetCaption pdtStartUpTime      'Format(pdtStartUpTime, "YYYY-MM-DD HH:mm")

    m_oFastReturn.SetObject clRUserID  '打印工号
    m_oFastReturn.SetCaption GetActiveUserID
    
    m_oFastReturn.SetObject clRBusID  '
    m_oFastReturn.SetCaption "(" & pszBusID & ")"
    
    m_oFastReturn.SetObject clREndStation  '
    m_oFastReturn.SetCaption pszEndStation
    
    m_oFastReturn.SetObject clRNowDate   '打印日期
    m_oFastReturn.SetCaption Format(Now, "YYYY-MM-DD HH:mm")
    
    m_oFastReturn.SetObject clRTicketType  '
    m_oFastReturn.SetCaption pszTicketType

    m_oFastReturn.SetObject clREndStation2  '
    m_oFastReturn.SetCaption pszEndStation

    m_oFastReturn.SetObject clRTicketType2  '
    m_oFastReturn.SetCaption pszTicketType

    m_oFastReturn.SetObject clRStartUpTime2  '发车时间
    m_oFastReturn.SetCaption Format(pdtStartUpTime, "HH:mm")

    m_oFastReturn.SetObject clRTicketCount  '
    m_oFastReturn.SetCaption "1张"

    m_oFastReturn.SetObject clRTicketCount2  '
    m_oFastReturn.SetCaption "1张"

    m_oFastReturn.SetObject clRReturnCharge2  '
    m_oFastReturn.SetCaption Format(psgReturnCharge, "0.00") & "元"

    m_oFastReturn.SetObject clRReturnCountStr  '
'    m_oFastReturn.SetCaption "退票张数:"

    m_oFastReturn.SetObject clRReturnChargeStr  '
'    m_oFastReturn.SetCaption "退票手续费:"
    
    m_oFastReturn.SetObject clRReturnTimeStr   '
'    m_oFastReturn.SetCaption "退票时间:"

    m_oFastReturn.SetObject clRTicketID2  '
    m_oFastReturn.SetCaption pszTicketNo

    m_oFastReturn.SetObject clRBusID2  '
    m_oFastReturn.SetCaption pszBusID

    m_oFastReturn.SetObject clRVehicleType  '
    m_oFastReturn.SetCaption pszVehicleType
    
    
    m_oFastReturn.SetObject clRSeatNo '
    m_oFastReturn.SetCaption pszSeatNo
    
    m_oFastReturn.SetObject clRStartStation1
    m_oFastReturn.SetCaption pszStartStation
    
    m_oFastReturn.SetObject clRStartStation2
    m_oFastReturn.SetCaption pszStartStation
    
    m_oFastReturn.SetObject clRReturnChargeStr2
'    m_oFastReturn.SetCaption "退票手续费:"
    
    m_oFastReturn.SetObject clRBusDate1
    m_oFastReturn.SetCaption Format(pdtBusDate, "yyyy-mm-dd")
    
    m_oFastReturn.SetObject clRBusDate2
    m_oFastReturn.SetCaption Format(pdtBusDate, "yyyy-mm-dd")
    
    '东阳补充

    m_oFastReturn.SetObject clRUserID1  '打印工号
    m_oFastReturn.SetCaption GetActiveUserID
    
    m_oFastReturn.SetObject clRSeatNo1 '座位号
    m_oFastReturn.SetCaption pszSeatNo
        
    m_oFastReturn.PrintFile
    
    m_oFastReturn.ClosePort
    
#End If

End Function




