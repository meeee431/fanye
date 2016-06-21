Attribute VB_Name = "Insurance"
Option Explicit

Public Function CombInsurance(papiTicketInfo() As TPrintTicketParam, szBusID() As String, pnTicketCount() As Integer, pszBusDate() As String, pszEndStationID() As String, pszEndStationName() As String, pszOffTime() As String, pszSellStationID() As String, szSellStationName() As String, aszRealNameInfo() As TCardInfo, Optional pbIsUseInsurance As Boolean) As String()
    Dim sazTemp() As String
    Dim i As Integer
    Dim szTmp As String
    Dim nCount As Integer
    Dim j As Integer
    Dim X As Integer
    Dim Num As Long
    Dim bIsChild As Boolean '老客运系统是否有携童票种，根据各个车站的实际情况而定
    bIsChild = True
    
    nCount = ArrayLength(szBusID)
    For X = 1 To ArrayLength(pnTicketCount)
        Num = Num + pnTicketCount(X)
    Next X
    X = 0
    ReDim sazTemp(1 To Num)
    
    Dim nStart As Integer
    If g_bIsUseInsurance Then
        For i = 1 To nCount
            For j = 1 To pnTicketCount(i)
                If bIsChild = True And papiTicketInfo(i).aptPrintTicketInfo(j).nTicketType = m_szSpecialTicketTypePosition Then
                    szTmp = "" & "|" & papiTicketInfo(i).aptPrintTicketInfo(j).szTicketNo & "|" & Format(pszBusDate(i), "yyyy-mm-dd") & "|" _
                            & Trim(szBusID(i)) & "|" & Format(pszBusDate(i), "yyyy-mm-dd") & " " & pszOffTime(i) & "|" & m_szCurrentUnitID & "|" _
                            & "" & "|" & pszSellStationID(i) & "|" & szSellStationName(i) & "|" _
                            & "" & "|" & "" & "|" & Trim(pszEndStationID(i)) & "|" & pszEndStationName(i) & "|" _
                            & Trim(papiTicketInfo(i).aptPrintTicketInfo(j).nTicketType) & "|" _
                            & Trim(GetTicketTypeStr2(Trim(papiTicketInfo(i).aptPrintTicketInfo(j).nTicketType))) & "|" _
                            & m_oAUser.UserID & "|" & m_oAUser.UserName & "|" & papiTicketInfo(i).aptPrintTicketInfo(j).szSeatNo & "|" _
                            & FormatMoney(papiTicketInfo(i).aptPrintTicketInfo(j).sgTicketPrice) & "|" _
                            & "0.00" & "|" & 1 _
                            & "|" & ReplaceEnterKey(aszRealNameInfo(nStart + j).szIDCardNo) _
                            & "|" & ReplaceEnterKey(aszRealNameInfo(nStart + j).szPersonName) & "|"
                Else
                    szTmp = "" & "|" & papiTicketInfo(i).aptPrintTicketInfo(j).szTicketNo & "|" & Format(pszBusDate(i), "yyyy-mm-dd") & "|" _
                            & Trim(szBusID(i)) & "|" & Format(pszBusDate(i), "yyyy-mm-dd") & " " & pszOffTime(i) & "|" & m_szCurrentUnitID & "|" _
                            & "" & "|" & pszSellStationID(i) & "|" & szSellStationName(i) & "|" _
                            & "" & "|" & "" & "|" & Trim(pszEndStationID(i)) & "|" & pszEndStationName(i) & "|" _
                            & Trim(papiTicketInfo(i).aptPrintTicketInfo(j).nTicketType) & "|" _
                            & Trim(GetTicketTypeStr2(Trim(papiTicketInfo(i).aptPrintTicketInfo(j).nTicketType))) & "|" _
                            & m_oAUser.UserID & "|" & m_oAUser.UserName & "|" & papiTicketInfo(i).aptPrintTicketInfo(j).szSeatNo & "|" _
                            & FormatMoney(papiTicketInfo(i).aptPrintTicketInfo(j).sgTicketPrice) & "|" _
                            & "0.00" & "|" & 0 _
                            & "|" & ReplaceEnterKey(aszRealNameInfo(nStart + j).szIDCardNo) _
                            & "|" & ReplaceEnterKey(aszRealNameInfo(nStart + j).szPersonName) & "|"
                End If
                X = X + 1
                sazTemp(X) = szTmp
            Next j
            nStart = nStart + pnTicketCount(i)
        Next i
    Else
        For i = 1 To nCount
            For j = 1 To pnTicketCount(i)
                szTmp = IIf(bIsChild = True And papiTicketInfo(i).aptPrintTicketInfo(j).nTicketType = m_szSpecialTicketTypePosition, 1, 0) & "|" & papiTicketInfo(i).aptPrintTicketInfo(j).szTicketNo & "|" & papiTicketInfo(i).aptPrintTicketInfo(j).szSeatNo & "|" _
                        & Format(m_oParam.NowDateTime, "yyyy-mm-dd hh:mm:ss") & "|" & szSellStationName(i) & "|" & pszSellStationID(i) & "|" _
                        & pszEndStationName(i) & "|" & Trim(pszEndStationID(i)) & "|" & FormatMoney(papiTicketInfo(i).aptPrintTicketInfo(j).sgTicketPrice) & "|" _
                        & Format(pszBusDate(i), "yyyy-mm-dd") & " " & pszOffTime(i) & "|" & m_oAUser.UserID & "|" & 1 _
                        & "|" & ReplaceEnterKey(aszRealNameInfo(nStart + j).szIDCardNo) _
                        & "|" & ReplaceEnterKey(aszRealNameInfo(nStart + j).szPersonName) & "|"
                X = X + 1
                sazTemp(X) = szTmp
            Next j
            nStart = nStart + pnTicketCount(i)
        Next i
    End If
    
    CombInsurance = sazTemp
    
End Function

Public Sub SaveInsurance(pszStr() As String)  '嘉兴保险打印存放
    Dim oFile As Object  '文件对象
    Dim FileNo As Integer
    Dim i As Integer
    Const cszInsurance = "c:\cyx\ticket" '保险打印文本存放位置
    
On Error GoTo here
    Set oFile = CreateObject("Scripting.FileSystemObject")
    FileNo = FreeFile
        If oFile.FolderExists(Left(cszInsurance, InStrRev(cszInsurance, "\", , vbTextCompare) - 1)) = False Then
            oFile.CreateFolder Left(cszInsurance, InStrRev(cszInsurance, "\", , vbTextCompare) - 1)
        End If
    
    Open cszInsurance & "_" & Trim(m_oAUser.UserID) & ".txt" For Output As #FileNo
        For i = 1 To ArrayLength(pszStr)
            Print #FileNo, pszStr(i)
        Next i
    Close #FileNo
    
    Exit Sub
here:
    frmNotify.m_szErrorDescription = err.Description
    frmNotify.Show vbModal
End Sub


