Attribute VB_Name = "mdlRunRev"
Option Explicit

Public szCmdLine As String
Public g_oActiveUser As New ActiveUser
Public g_szPath As String

'��¼ϵͳ��ʼ��g_oActiveUser
Public Sub Main()

    If App.PrevInstance Then
        MsgBox "�������ɻ�������������ɺ������ԣ�", vbExclamation, "����"
        End
    End If

    '�û���������
    Dim szUser As String, szPassword As String
    Dim szTemp As String
    'ȡ�����в���
    '��һ�������͵ڶ��������ֱ�Ϊ�û����Ϳ���
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
    MsgBox "ִ�г�����ָ�������в���:" & vbCrLf & _
        "ʹ��: PSTRunRev username,password,[rundate],[[bus1]��[bus2],...],[PromptWhenError]," & _
            "[AppEndExit],[CreateLogFile],[StopMake],[IsTray]" & vbCrLf & _
            "����˵��:(�ö��Ÿ���,�����б����������)" & vbCrLf & _
            "  UserName:��Ч���û���" & vbCrLf & _
            "  Password:�û�����" & vbCrLf & _
            "  [RunDate]:���ɼƻ����ڣ�ȱʡΪԤ�����������һ�죩(��ѡ)" & vbCrLf & _
            "  [[bus1],[bus2],...]:���ɳ������飬ȱʡΪ�գ����г��Σ�(��ѡ)" & vbCrLf & _
            "  [PromptWhenError]:������ʾ��־('F'����ʾ,'T'��ʾ[ȱʡ]��(��ѡ)" & vbCrLf & _
            "  [AppEndExit]����������ɺ��Ƿ��˳�('F'�˳�,'T'���˳�[ȱʡ])(��ѡ)" & vbCrLf & _
            "  [CreateLogFile]������Ϣ�ļ���(ȱʡ�Զ�����)(��ѡ)" & vbCrLf & _
            "  [StopMake]ͣ�೵������('T'����[ȱʡ]��'F'������)(��ѡ)" & vbCrLf & _
            "  [ISTray])�Ƿ������̷�ʽ���г���'F'��[ȱʡ],'T'��)(��ѡ)", _
            vbExclamation + vbOKOnly
    End
End Sub
