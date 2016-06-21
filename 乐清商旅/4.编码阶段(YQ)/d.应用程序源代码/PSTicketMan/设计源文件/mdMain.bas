Attribute VB_Name = "mdMain"
Public m_oActiveUser As ActiveUser
Public m_oParam As New SystemParam
Public m_oShell As New CommDialog

Private m_lTicketNoNumLen As Long
Public m_szTicketPrefix As String
Public m_lTicketNo As Long
Private m_szTicketNoFromatStr As String
Public m_oSell As New SellTicketClient

'对话框调用状态
Public Enum EFormStatus
    EFS_AddNew = 0
    EFS_Modify = 1
    EFS_Show = 2
    EFS_Delete = 3
End Enum
Public Sub Main()
    Dim szLoginCommandLine As String
    Dim oShell As New CommShell
    On Error GoTo HelpFileErr
GoOn:
    On Error GoTo 0
    
    Set m_oActiveUser = oShell.ShowLogin()
    If m_oActiveUser Is Nothing Then Exit Sub
'    oShell.ShowSplash "票证管理", "Station Business TicketMan", LoadResPicture(101, 0), App.Major, App.Minor, App.Revision
    DoEvents
    m_oShell.Init m_oActiveUser
    m_oParam.Init m_oActiveUser

    g_szUnitID = m_oParam.UnitID
    MDITicketMan.Show
    oShell.CloseSplash
    DoEvents
    Exit Sub
HelpFileErr:
    ShowMsg "不能找到帮助文件"
    Resume GoOn
End Sub

Public Function TicketNoNumLen() As Integer
    If m_lTicketNoNumLen = 0 Then
        m_lTicketNoNumLen = m_oParam.TicketNumberLen
    End If
    TicketNoNumLen = m_lTicketNoNumLen
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
        m_szTicketNoFromatStr = String(TicketNoNumLen(), "0")
    End If
    TicketNoFormatStr = m_szTicketNoFromatStr
End Function

Public Sub GetAppSetting(Optional szUserID As String)
'    Dim oReg As New CFreeReg
    Dim szLastTicketNo As String
    m_oSell.Init m_oActiveUser
    szLastTicketNo = m_oSell.GetLastTicketNo(szUserID)
    m_lTicketNo = ResolveTicketNo(szLastTicketNo, m_szTicketPrefix)
'    Set oReg = Nothing
End Sub
Public Function ResolveTicketNo(pszFullTicketNo, ByRef pszTicketPrefix As String) As Long
    Dim i As Integer, j As Integer
    Dim nCount As Integer, nTemp As Integer, nTicketPrefixLen As Integer
    'On Error Resume Next
    pszFullTicketNo = Trim(pszFullTicketNo)
    nCount = Len(pszFullTicketNo)
    
    For i = 1 To nCount
        nTemp = Asc(Mid(pszFullTicketNo, nCount - i + 1, 1))
        If nTemp < vbKey0 Or nTemp > vbKey9 Then
            Exit For
        End If
    Next
    i = i - 1
    If i > 0 Then
        nTemp = TicketNoNumLen()
        nTemp = IIf(nTemp > i, i, nTemp)
        ResolveTicketNo = CLng(Right(pszFullTicketNo, nTemp))
        
        nTicketPrefixLen = m_oParam.TicketPrefixLen
        If nTicketPrefixLen <= Len(pszFullTicketNo) Then
            pszTicketPrefix = Left(pszFullTicketNo, nTicketPrefixLen)
        Else
            pszTicketPrefix = pszFullTicketNo
        End If
        
    Else
        pszTicketPrefix = ""
        ResolveTicketNo = 0
    End If
    
    
    
End Function



'得到提示张数
Function GetTicketNum() As Integer
    Dim oTicketMan As New TicketMan
    Dim iHowTikcetts As Integer
    oTicketMan.Init m_oActiveUser
    iHowTikcetts = oTicketMan.GetParam("HowTicketts")
    If iHowTikcetts <> 0 Then
        GetTicketNum = iHowTikcetts
    Else
        GetTicketNum = 0
    End If
End Function
