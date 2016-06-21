Attribute VB_Name = "mdTSvr"
Option Explicit

Public Const cszSellTicket = ""

Public m_szTicketPrefix As String
Public m_lTicketNo As Long
Public m_lTicketNoNumLen As Long
Private m_szTicketNoFromatStr As String
'可互联售票
'Public Const CnInternetCanSell = 0
Public Const CnInternetNotCanSell = 1
Public Function FormatTail(pdbValue As Double) As Double
    '进行尾数处理
    '0-4算0,5-9算10
    Dim dbTemp As Double
    dbTemp = pdbValue - Int(pdbValue)
    If dbTemp >= 0 And dbTemp < 0.5 Then
        '0-4算0
        FormatTail = Int(pdbValue)
    Else
        '5-9算10
        FormatTail = FormatMoney(Int(pdbValue) + 1)
    End If
End Function

Public Function GetTicketNo(Optional pnOffset As Integer = 0) As String
    GetTicketNo = MakeTicketNo(m_lTicketNo + pnOffset, m_szTicketPrefix)
End Function

Public Function MakeTicketNo(plTicketNo As Long, Optional pszPrefix As String = "") As String
    MakeTicketNo = pszPrefix & Format(plTicketNo, TicketNoFormatStr())
End Function
Private Function TicketNoFormatStr() As String
    Dim i As Integer
    If m_szTicketNoFromatStr = "" Then
        m_szTicketNoFromatStr = String(m_lTicketNoNumLen, "0")
    End If
    TicketNoFormatStr = m_szTicketNoFromatStr
End Function




