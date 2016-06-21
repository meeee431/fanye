Attribute VB_Name = "mdlRunRev"
Option Explicit

Public szCmdLine As String
Public g_oActiveUser As New ActiveUser
Public g_szPath As String

'登录系统初始化g_oActiveUser
Public Sub Main()

    If App.PrevInstance Then
        MsgBox "正在生成环境，等生成完成后，再重试！", vbExclamation, "警告"
        End
    End If

    '用户名、密码
    Dim szUser As String, szPassword As String
    Dim szTemp As String
    '取命令行参数
    '第一个参数和第二个参数分别为用户名和口令
    szCmdLine = Command
    
    If szCmdLine = "" Then GoTo NoParaPrompt
    szUser = ""
    szPassword = ""
    
    szUser = LeftAndRight(szCmdLine, True, ",")
    szTemp = LeftAndRight(szCmdLine, False, ",")
    szPassword = LeftAndRight(szTemp, True, ",")
    g_oActiveUser.Login szUser, szPassword, GetComputerName
    Load frmMake
    frmMake.Show
    App.HelpFile = SetHTMLHelpStrings("SNScheme.CHM")
    Exit Sub
ErrorDo:
NoParaPrompt:
    MsgBox "执行程序需指定命令行参数:" & vbCrLf & _
        "使用: PSTRunRev username,password,[rundate],[[bus1]，[bus2],...],[PromptWhenError]," & _
            "[AppEndExit],[CreateLogFile],[StopMake],[IsTray]" & vbCrLf & _
            "参数说明:(用逗号隔开,车次列表外加中括号)" & vbCrLf & _
            "  UserName:有效的用户名" & vbCrLf & _
            "  Password:用户口令" & vbCrLf & _
            "  [RunDate]:生成计划日期（缺省为预售天数的最后一天）(可选)" & vbCrLf & _
            "  [[bus1],[bus2],...]:生成车次数组，缺省为空（所有车次）(可选)" & vbCrLf & _
            "  [PromptWhenError]:错误提示标志('F'不提示,'T'提示[缺省]）(可选)" & vbCrLf & _
            "  [AppEndExit]程序运行完成后是否退出('F'退出,'T'不退出[缺省])(可选)" & vbCrLf & _
            "  [CreateLogFile]生成信息文件名(缺省自动创建)(可选)" & vbCrLf & _
            "  [StopMake]停班车次生成('T'生成[缺省]，'F'不生成)(可选)" & vbCrLf & _
            "  [ISTray])是否以托盘方式运行程序，'F'不[缺省],'T'是)(可选)", _
            vbExclamation + vbOKOnly
    End
End Sub
