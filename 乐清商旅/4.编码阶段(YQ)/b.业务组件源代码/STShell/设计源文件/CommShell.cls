VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
'Public Enum ESNCommDlgShowType
'    TP_SNSelArea = 1
'    TP_SNSelBus = 2
'    TP_SNSelCheckGate = 3
'    TP_SNSelCompany = 4
'    TP_SNInput = 5
'    TP_SNSelOwner = 6
'    TP_SNSelProject = 7
'    TP_SNSelRoute = 8
'    TP_SNSelSection = 9
'    TP_SNSelStation = 10
'    TP_SNUser = 11
'    TP_SNSelVehicle = 12
'    TP_SNSelVehicelModel = 13
'    TP_SNProtocol = 14 '拆算协议
'    TP_SNSelEnvBus = 15 '拆算协议
'    TP_SNSelCard = 16 '报班卡
'    TP_SNSelChunnul = 17 '通道
'    TP_SNSelRoom = 18 '候车厅
'    TP_SNSelFile = 19 '文件
'    TP_SNSelRule = 20 '规则
'    TP_SNSelTask = 21 '任务
'
'
'End Enum

'
'Public Enum BusTypeConstants
'    AllBusType = 0  '所有车次
'    RegularBus = 1  '固定车次
'    ScrollBus = 2   '流水车次
'End Enum


'Public m_nShowType As ESNCommDlgShowType


'关闭Splash窗口
Public Sub CloseSplash()
    On Error Resume Next
    Unload frmSplash
End Sub
'显示Splash窗口
Public Sub ShowSplash(AppName As String, _
    AppEnglishName As String, _
    imgProductLogo As StdPicture, _
    VerMajor As String, _
    VerMinor As String, _
    VerBuild As String)
    
    frmSplash.AppEnglishName = AppEnglishName
    frmSplash.AppName = AppName
    frmSplash.VerMajor = VerMajor
    frmSplash.VerMinor = VerMinor
    frmSplash.VerBuild = VerBuild
    Set frmSplash.imgProductLogo = imgProductLogo
    frmSplash.Show
End Sub
'显示登录对话框
Public Function ShowLogin(Optional szPassword As String = "") As ActiveUser
    Dim m_nLogCount As Integer
    Dim oAUser As New ActiveUser
    
    m_nLogCount = 0
    On Error GoTo here

RetryInput:
    frmLogin.Show vbModal
    If frmLogin.m_bLoginOk Then
        oAUser.Login frmLogin.m_szUserID, frmLogin.m_szPasword, GetComputerName()
''        Set m_oActiveUser = oAUser
''        Set ShowLogin = m_oActiveUser
        Set ShowLogin = oAUser
    End If
    If szPassword <> "" Then
        szPassword = frmLogin.m_szPasword
    End If
    Exit Function
here:
    If err.Number = ERR_ActiveUserUserNotExistOrPasswordNotOk Then
        m_nLogCount = m_nLogCount + 1
        If m_nLogCount >= cnMaxLoginErrCount Then
            MsgBox "错误输入密码次数太多！", , "注意"
        Else
            MsgBox err.Description & "，你还有" & cnMaxLoginErrCount - m_nLogCount & "次机会！", , "注意"
            Resume RetryInput
        End If
    Else
        MsgBox err.Description, , "错误"
    End If
    
End Function


'显示输入对话框
Public Function ShowInput(Optional DlgTitle As String = "", Optional AddMsg As String = "", Optional NumberOnly As Boolean = False, Optional pszDefault As String) As String
    Load frmInput
    If DlgTitle <> "" Then
        frmInput.Caption = DlgTitle
    End If
    If AddMsg <> "" Then
        frmInput.lblMsg.Caption = AddMsg
    End If
    frmInput.mbNumberOnly = NumberOnly
    frmInput.txtInput.Text = pszDefault
    frmInput.txtInput.SelStart = 0
    frmInput.txtInput.SelLength = Len(pszDefault)
    frmInput.Show vbModal
    
    ShowInput = frmInput.mszResult
End Function

'显示About窗口
Public Sub ShowAbout(AppName As String, _
    AppEnglishName As String, _
    AppDescription As String, _
    imgProductLogo As StdPicture, _
    VerMajor As String, _
    VerMinor As String, _
    VerBuild As String)
    
    frmAbout.AppName = AppName
    frmAbout.AppEnglishName = AppEnglishName
    frmAbout.AppDescription = AppDescription
    frmAbout.VerMajor = VerMajor
    frmAbout.VerMinor = VerMinor
    frmAbout.VerBuild = VerBuild
    Set frmAbout.imgProductLogo = imgProductLogo
    frmAbout.Show vbModal
End Sub

Public Property Get UserPassword() As String
    UserPassword = frmLogin.m_szPasword
End Property
