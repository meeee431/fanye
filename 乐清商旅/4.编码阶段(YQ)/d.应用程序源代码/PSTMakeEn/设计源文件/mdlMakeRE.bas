Attribute VB_Name = "mdlMakeRe"
Global g_oActiveUser As ActiveUser
Public Const cszMakeEn = "MakeEn"
Public m_szPassword As String
Public m_szUser As String
Public m_szExecute As String
Public Declare Function GetVersion Lib "kernel32" () As Long 'ÅÐ¶ÏWindows °æ±¾


Public Sub InitReg(oReg As Object)
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
End Sub

Public Sub Main()
    If App.PrevInstance Then
        End
    End If
    Dim SnShell As New CommShell
    Dim szLoginCommandLine As String
    szLoginCommandLine = TransferLoginParam(Trim(Command()))
    
    m_szPassword = "PASS"
    If szLoginCommandLine = "" Then
        Set g_oActiveUser = SnShell.ShowLogin(m_szPassword)
    Else
        Set g_oActiveUser = New ActiveUser
        g_oActiveUser.Login GetLoginParam(szLoginCommandLine, cszUserName), GetLoginParam(szLoginCommandLine, cszUserPassword), GetComputerName()
        m_szPassword = GetLoginParam(szLoginCommandLine, cszUserPassword)
        m_szUser = GetLoginParam(szLoginCommandLine, cszUserName)
'        SnShell.Init g_oActiveUser
        
    End If
    If Not g_oActiveUser Is Nothing Then
        frmMain.Show
        m_szUser = g_oActiveUser.UserID
'        App.HelpFile = SetHTMLHelpStrings("SNScheme.CHM")
        frmMain.txtExePassword.Text = m_szPassword
        frmMain.txtExeUser.Text = m_szUser
        SetHTMLHelpStrings "STMakeEn.chm"
    Else
        End
    End If
End Sub

Public Sub ShowErrorU(ErrNumber As Long)
    MsgBox err.Description, vbExclamation + vbOKOnly, "´íÎó" & err.Number
End Sub
