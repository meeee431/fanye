VERSION 5.00
Object = "{A0123751-4698-48C1-A06C-A2482B5ED508}#2.0#0"; "RTComctl2.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ɻ�������"
   ClientHeight    =   4155
   ClientLeft      =   2265
   ClientTop       =   2325
   ClientWidth     =   7770
   HelpContextID   =   3000001
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   7770
   Begin VB.Timer tmStart 
      Interval        =   1
      Left            =   6330
      Top             =   2835
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "�趨(&O)>>"
      Height          =   315
      HelpContextID   =   6000015
      Left            =   6465
      TabIndex        =   14
      Top             =   975
      Width           =   1185
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   315
      Left            =   6465
      TabIndex        =   13
      Top             =   585
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "�ر�(&C)"
      Height          =   315
      Left            =   6465
      TabIndex        =   12
      Top             =   180
      Width           =   1185
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "��־��ѯ"
      Height          =   1020
      Left            =   90
      TabIndex        =   17
      Top             =   3030
      Width           =   6180
      Begin RTComctl2.FileBrowser txtLogFile 
         Height          =   300
         HelpContextID   =   6000019
         Left            =   1005
         TabIndex        =   6
         Top             =   585
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Filter          =   "��־�ļ�(*.cei) | *.cei|�����ļ�(*.*)|*.*"
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "�鿴(&V)"
         Height          =   315
         HelpContextID   =   6000019
         Left            =   4890
         TabIndex        =   4
         Top             =   578
         Width           =   1185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "��־�ļ�(&D):"
         Height          =   180
         Left            =   1005
         TabIndex        =   5
         Top             =   315
         Width           =   1080
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   135
         Picture         =   "frmMain.frx":16AC2
         Top             =   390
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "�������л���"
      Height          =   2790
      Left            =   105
      TabIndex        =   16
      Top             =   105
      Width           =   6180
      Begin VB.TextBox txtPassword 
         Enabled         =   0   'False
         Height          =   300
         HelpContextID   =   6000011
         IMEMode         =   3  'DISABLE
         Left            =   4440
         PasswordChar    =   "*"
         TabIndex        =   26
         Top             =   2310
         Width           =   1155
      End
      Begin VB.TextBox txtUser 
         Enabled         =   0   'False
         Height          =   300
         HelpContextID   =   6000011
         Left            =   2430
         TabIndex        =   24
         Top             =   2310
         Width           =   1125
      End
      Begin VB.CommandButton cmdExecute 
         Caption         =   "��������(&K)"
         Enabled         =   0   'False
         Height          =   315
         HelpContextID   =   6000011
         Left            =   4860
         TabIndex        =   3
         Top             =   1275
         Width           =   1185
      End
      Begin VB.CheckBox chkTask 
         BackColor       =   &H00FFFFC0&
         Caption         =   "(&A)ÿ�հ������Զ�����"
         Height          =   255
         HelpContextID   =   6000011
         Left            =   960
         TabIndex        =   1
         Top             =   1560
         Width           =   2265
      End
      Begin VB.CommandButton cmdMakeBus 
         Caption         =   "����(&M)"
         Height          =   315
         Left            =   4875
         TabIndex        =   0
         Top             =   645
         Width           =   1185
      End
      Begin MSComCtl2.DTPicker dtpTaskTime 
         Height          =   300
         HelpContextID   =   6000011
         Left            =   1740
         TabIndex        =   2
         Top             =   1875
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   22872066
         CurrentDate     =   .875
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&Q):"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3615
         TabIndex        =   25
         Top             =   2370
         Width           =   720
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NT�û���(&N):"
         Enabled         =   0   'False
         Height          =   180
         Left            =   1320
         TabIndex        =   23
         Top             =   2370
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Զ��������л���"
         Enabled         =   0   'False
         Height          =   180
         Left            =   2955
         TabIndex        =   22
         Top             =   1935
         Width           =   1440
      End
      Begin VB.Label lblDay 
         BackStyle       =   0  'Transparent
         Caption         =   "ÿ��"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1320
         TabIndex        =   21
         Top             =   1935
         Width           =   375
      End
      Begin VB.Label lblTask 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ÿ���Զ��������л�����ʱ��"
         Height          =   180
         Left            =   930
         TabIndex        =   20
         Top             =   1275
         Width           =   2700
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ƻ�"
         Height          =   180
         Left            =   150
         TabIndex        =   19
         Top             =   1020
         Width           =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   960
         X2              =   6050
         Y1              =   1125
         Y2              =   1125
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   960
         X2              =   6050
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "�������л���������ѡ�񳵴����ɻ�ȫ�����ɡ�ϵͳ������Ԥ�������ڵ��κ�һ������г��Ρ�"
         Height          =   525
         Left            =   960
         TabIndex        =   18
         Top             =   285
         Width           =   3810
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   135
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ϵͳ����"
      Height          =   2130
      Left            =   90
      TabIndex        =   15
      Top             =   4200
      Width           =   6180
      Begin VB.TextBox txtExeUser 
         Height          =   300
         HelpContextID   =   6000015
         Left            =   2055
         TabIndex        =   28
         Top             =   1635
         Width           =   1125
      End
      Begin VB.TextBox txtExePassword 
         Height          =   300
         HelpContextID   =   6000015
         IMEMode         =   3  'DISABLE
         Left            =   4155
         PasswordChar    =   "*"
         TabIndex        =   31
         Top             =   1620
         Width           =   1155
      End
      Begin RTComctl2.FileBrowser txtExecuteFile 
         Height          =   300
         HelpContextID   =   6000015
         Left            =   2070
         TabIndex        =   10
         Top             =   1230
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   529
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Filter          =   "���л������ɳ���(snrunrev.exe) | snrunrev.exe"
         InitDir         =   ""
      End
      Begin RTComctl2.FolderBrowser txtLogDirect 
         Height          =   300
         HelpContextID   =   6000015
         Left            =   2070
         TabIndex        =   8
         Top             =   855
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   529
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "����(&S)"
         Enabled         =   0   'False
         Height          =   315
         HelpContextID   =   6000015
         Left            =   4875
         TabIndex        =   11
         Top             =   285
         Width           =   1185
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "�趨��־�ļ����Ŀ¼�����л��������ŵ�λ�á� �����������л������û��������롣"
         Height          =   465
         Left            =   720
         TabIndex        =   30
         Top             =   285
         Width           =   4170
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�û���(&G):"
         Height          =   180
         Left            =   195
         TabIndex        =   27
         Top             =   1680
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&F):"
         Height          =   180
         Left            =   3315
         TabIndex        =   29
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���л������ɳ���(&E):"
         Height          =   180
         Left            =   195
         TabIndex        =   9
         Top             =   1305
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��־���Ŀ¼(&P):"
         Height          =   180
         Left            =   195
         TabIndex        =   7
         Top             =   915
         Width           =   1440
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   150
         Picture         =   "frmMain.frx":1738C
         Top             =   270
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_oReg As New CFreeReg
Private m_oTask As Task2
Private m_oTaskMan As New TaskScheduler2

Private Sub chkTask_Click()
    If chkTask.Value = vbChecked Then
        dtpTaskTime.Enabled = True
        lblDay.Enabled = True
        lblInfo.Enabled = True
        If IsOsNt Then
            SetUserPassword True
        End If
        cmdExecute.Caption = "��������(&K)"
    Else
        dtpTaskTime.Enabled = False
        lblDay.Enabled = False
        lblInfo.Enabled = False
        SetUserPassword False
        cmdExecute.Caption = "ɾ������(&K)"
    End If
    cmdExecute.Enabled = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
    End
End Sub

Private Sub cmdExecute_Click()
    If chkTask.Value Then
        If MsgBox("��������[MakeRunEnv]ÿ��" & Format(dtpTaskTime.Value, "HH:MM:SS") & "����?", vbQuestion + vbYesNo, "�ƻ�") = vbYes Then
            SaveTask
            MsgBox "���񱣴����", vbInformation, "���ɻ���"
            cmdExecute.Enabled = False
        End If
    Else
        If MsgBox("ɾ������[MakeRunEnv],�Ժ󶼲�������?", vbQuestion + vbYesNo + vbDefaultButton2, "�ƻ�") = vbYes Then
            DeleteTask
            MsgBox "����ɾ�����", vbInformation, "���ɻ���"
            cmdExecute.Enabled = False
        End If
    End If
End Sub

Private Sub cmdHelp_Click()
DisplayHelp Me
End Sub

Private Sub cmdMakeBus_Click()
    m_szExecute = txtExecuteFile.Text
    frmMakeRE.Show , Me
End Sub

Private Sub cmdSave_Click()
    txtLogFile.InitDir = Trim(txtLogDirect.Text)
On Error GoTo here
    m_oReg.SaveSetting cszMakeEn & "\MakeEnvDesktop", "LogDirect", txtLogDirect.Text
    m_oReg.SaveSetting cszMakeEn & "\MakeEnvDesktop", "ExecuteFile", txtExecuteFile.Text
    cmdSave.Enabled = False
Exit Sub
here:
    ShowErrorU err.Number
End Sub

Private Sub cmdSetup_Click()
    Me.Height = 6765
    cmdSetup.Enabled = False
    txtLogDirect.Enabled = True
    txtExecuteFile.Enabled = True
End Sub

Private Sub cmdView_Click()
    Dim szExecuteApp As String
    Dim tempFile As String
    'Dim frmTemp As New frmViewLog
    'frmTemp.m_szLogFile = txtLogFile.Text
    'frmTemp.Show , Me
    tempFile = Trim(txtLogFile.Text)
    On Error GoTo here
    If Not FileIsExist(tempFile) Then
        MsgBox "�ļ�" & tempFile & "������!", vbOKOnly + vbInformation
        Exit Sub
    End If
    szExecuteApp = "NOTEPAD.EXE " & tempFile
    Shell szExecuteApp, vbNormalFocus
Exit Sub
here:
    ShowErrorU err.Number
End Sub


Private Sub dtpTaskTime_Change()
    cmdExecute.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
End Sub

Private Sub tmStart_Timer()
    Dim tempdir As String
    Dim szTemp As String
On Error GoTo here
    tmStart.Enabled = False
    
    InitReg m_oReg
    tempdir = m_oReg.GetSetting(cszMakeEn & "\MakeEnvDesktop", "LogDirect", "") '"C:\My Documents")
    If tempdir = "" Then
        tempdir = Environ("temp")
        m_oReg.SaveSetting cszMakeEn & "\MakeEnvDesktop", "LogDirect", tempdir
    End If
    
    If Not Right(tempdir, 1) = "\" Then
        tempdir = tempdir + "\"
    End If
    txtLogFile.InitDir = tempdir
    txtLogDirect.Text = tempdir
    m_szExecute = m_oReg.GetSetting(cszMakeEn & "\MakeEnvDesktop", "ExecuteFile", App.Path & "\PSTRunRev.exe")
    If m_szExecute = "" Then
        m_szExecute = App.Path & "\PSTRunRev.exe"
    End If
    txtLogFile.Filter = "��־�ļ�(*.REN) | *.REN|�����ļ�(*.*)|*.*"
    txtLogFile.Text = tempdir & Format(Date, "YYYYMMDD") & ".REN"
    txtExecuteFile.Text = m_szExecute
    cmdSave.Enabled = False
    If IsOsNt Then
        
        chkTask.Value = Unchecked
        lblDay.Enabled = False
        lblInfo.Enabled = False
        SetUserPassword False
        
    Else
        GetSchemeRunTime
    End If
    cmdExecute.Enabled = False
Exit Sub
here:
    ShowErrorU err.Number
End Sub

Private Sub txtExecuteFile_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtExecuteFile_LostFocus()
    Dim tempFile As String
    txtExecuteFile.Text = Trim(txtExecuteFile.Text)
    tempFile = txtExecuteFile.Text
    If Not FileIsExist(tempFile) Then
        MsgBox "�ļ�" & tempFile & "������!", vbOKOnly + vbInformation
        txtExecuteFile.Text = m_oReg.GetSetting(cszMakeEn & "\MakeEnvDesktop", "ExecuteFile", App.Path & "\PSTRunRev.exe")
    End If
End Sub

Private Sub txtExePassword_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtExeUser_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtLogDirect_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtLogDirect_LostFocus()
    Dim tempdir As String
    txtLogDirect.Text = Trim(txtLogDirect.Text)
    tempdir = txtLogDirect.Text
    If Not Right(tempdir, 1) = "\" Then
        txtLogDirect.Text = tempdir + "\"
        tempdir = tempdir + "\"
    End If
    'If Not FileIsExist(tempdir) Then
    '    MsgBox "Ŀ¼" & tempFile & "������!", vbOKOnly + vbInformation
    '    txtLogDirect.Text = m_oReg.GetSetting("MakeEnvDesktop", "LogDirect", "C:\My Documents")
    'End If
End Sub

Private Sub GetSchemeRunTime()
    '�ж������Ƿ�����
On Error GoTo here
    Dim oTrigger As TaskTrigger2
    Dim tiTriggerInfo As tagTaskTriggerInfo
    Set m_oTask = m_oTaskMan.Activate("MakeRunEnv", 0)
    If m_oTask Is Nothing Then
        lblTask.Caption = "ϵͳ��δ�趨����MakeRunEnv"
        cmdExecute.Caption = "��������(&K)"
    Else
        lblTask.Caption = "����ÿ���Զ��������л�����ʱ��"
        cmdExecute.Caption = "��������(&K)"
    End If
    Set oTrigger = m_oTask.GetTrigger(0)
    tiTriggerInfo = oTrigger.GetTrigger
    With tiTriggerInfo
        dtpTaskTime.Value = CDate(.wStartHour & ":" & .wStartMinute)
    End With
    chkTask.Value = vbChecked
    lblDay.Enabled = True
    lblInfo.Enabled = True
    If IsOsNt Then
        SetUserPassword True
    Else
        SetUserPassword False
    End If
Exit Sub
here:
    chkTask.Value = vbUnchecked
    lblDay.Enabled = False
    lblInfo.Enabled = False
    SetUserPassword False
End Sub

Private Sub SaveTask()
On Error GoTo Error_Handle
    Dim tiTriggerInfo As tagTaskTriggerInfo
    Dim bAddModify As Boolean
    Dim oTask As Task2
    Dim oTrigger As TaskTrigger2
    Dim lTrigerIndex As Long
    bAddModify = False
    With tiTriggerInfo
        .TriggerType = TP_TIME_TRIGGER_DAILY
        .wBeginDay = Day(CDate(cszEmptyDateStr))
        .wBeginMonth = Month(CDate(cszEmptyDateStr))
        .wBeginYear = Year(CDate(cszEmptyDateStr))
        .wStartHour = Hour(dtpTaskTime.Value)
        .wStartMinute = Minute(dtpTaskTime.Value)
        .Type.lValue1 = 1 'ÿ��
    End With
    Set oTask = m_oTaskMan.Activate("MakeRunEnv", 0)
AddTask:
    If bAddModify Then
        Set oTask = m_oTaskMan.NewTask("MakeRunEnv", 0, 0)
    End If
    oTask.SetApplicationName txtExecuteFile.Text
    If Trim(txtExeUser.Text) = "" Then MsgBox "����дÿ��ִ�����ɻ������û���������", vbExclamation, "���ɻ���": Exit Sub
    oTask.SetParameters Trim(txtExeUser.Text) & "," & txtExePassword.Text
    If IsOsNt Then
        oTask.SetAccountInformation txtUser.Text, txtPassword.Text
    End If
    If bAddModify Then
        Set oTrigger = oTask.CreateTrigger(lTrigerIndex)
    Else
        Set oTrigger = oTask.GetTrigger(0)
    End If
    oTrigger.SetTrigger tiTriggerInfo
    oTask.Save '����
Exit Sub
Error_Handle:
    bAddModify = True
    Resume AddTask
End Sub

Private Sub DeleteTask()
On Error GoTo Error_Handle
    m_oTaskMan.Delete "MakeRunEnv"
Exit Sub
Error_Handle:
End Sub

Private Sub SetUserPassword(bEnabled As Boolean)
    lblPassword.Enabled = bEnabled
    lblUser.Enabled = bEnabled
    txtPassword.Enabled = bEnabled
    txtUser.Enabled = bEnabled
End Sub


