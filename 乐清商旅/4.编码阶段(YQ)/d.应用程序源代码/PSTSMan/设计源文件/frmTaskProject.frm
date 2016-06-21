VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTaskProject 
   BackColor       =   &H00FFFFFF&
   Caption         =   "设定自动删除日志任务"
   ClientHeight    =   5025
   ClientLeft      =   2355
   ClientTop       =   2115
   ClientWidth     =   5775
   HelpContextID   =   5003601
   Icon            =   "frmTaskProject.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5775
   Begin VB.Frame fraNt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Window NT 操作系统用户账号:"
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   135
      TabIndex        =   22
      Top             =   3525
      Width           =   4335
      Begin VB.TextBox txtNTPassword 
         Appearance      =   0  'Flat
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   2175
         PasswordChar    =   "*"
         TabIndex        =   24
         Top             =   570
         Width           =   1875
      End
      Begin VB.TextBox txtNTUser 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   2175
         TabIndex        =   23
         Top             =   270
         Width           =   1875
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "用户密码:"
         Height          =   270
         Left            =   1275
         TabIndex        =   26
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "用户账号:"
         Height          =   240
         Left            =   1275
         TabIndex        =   25
         Top             =   315
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H8000000A&
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4590
      TabIndex        =   20
      Top             =   1005
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "关闭(&C)"
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4590
      TabIndex        =   19
      Top             =   615
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   315
      Left            =   4590
      TabIndex        =   18
      Top             =   225
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "执行时间:"
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   120
      TabIndex        =   16
      Top             =   1335
      Width           =   4335
      Begin MSComCtl2.DTPicker dtpTime 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   4
         EndProperty
         Height          =   330
         Left            =   2190
         TabIndex        =   17
         ToolTipText     =   "推荐使用默认值,应设定系统空闲时执行,且不与其他任务执行时间冲突"
         Top             =   210
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   582
         _Version        =   393216
         Format          =   71958530
         CurrentDate     =   36530.9583333333
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "选择时间:"
         Height          =   180
         Left            =   585
         TabIndex        =   21
         Top             =   270
         Width           =   810
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "StationNet系统登录用户账号:"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   2055
      Width           =   4335
      Begin VB.TextBox txtPass 
         Appearance      =   0  'Flat
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   2190
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   960
         Width           =   1875
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   2190
         TabIndex        =   9
         Top             =   660
         Width           =   1875
      End
      Begin VB.OptionButton optGarnt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "指定用户:"
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   525
         TabIndex        =   6
         Top             =   420
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton optNow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "当前用户"
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   525
         TabIndex        =   5
         Top             =   225
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用户密码:"
         Height          =   180
         Left            =   1275
         TabIndex        =   8
         Top             =   1020
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用户账号:"
         Height          =   180
         Left            =   1275
         TabIndex        =   7
         Top             =   720
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "执行周期:"
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.ComboBox cboMonth 
         Height          =   300
         Left            =   2205
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   735
         Width           =   720
      End
      Begin VB.ComboBox cboWeek 
         Height          =   300
         Left            =   2190
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   375
         Width           =   735
      End
      Begin VB.OptionButton optDay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "每天"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   570
         TabIndex        =   3
         Top             =   165
         Width           =   795
      End
      Begin VB.OptionButton optWeek 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "每周:"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   570
         TabIndex        =   2
         Top             =   435
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optMonth 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "每月:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   570
         TabIndex        =   1
         Top             =   735
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "日"
         Height          =   180
         Left            =   3150
         TabIndex        =   15
         Top             =   780
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "第"
         Height          =   180
         Left            =   1695
         TabIndex        =   13
         Top             =   780
         Width           =   180
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "星期"
         Height          =   225
         Left            =   1680
         TabIndex        =   11
         Top             =   495
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmTaskProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim TriggerInfo As tagTaskTriggerInfo
'Dim TaskTriType As tagTaskTriggerType
'Dim TaskType As tagTriggerTypeUnion
'Dim lStartHour As Long
'Dim lStartMin As Long
'Dim szRunName As String
'Dim lBeginDay As Long
'Dim lBeginMouth As Long
'Dim lBeginYear As Long
'Dim szNTUser As String
'Dim szNTPassWord As String
'Dim szParam As String
'
'
''lvalue1=2^(日期-1)
'Const cAllMouth2 = 4095 'lvalue2 and for TP_TIME_TRIGGER_MONTHLYDATE,表示所有月有效
'Const cAllMouth3 = 0 '对TP_TIME_TRIGGER_MONTHLYDATE来说无意义
'
''lvalue1=2^星期数(星期日=0)
'Const cPerWeek2 = 1 'lvalue2
'Const cPerWeek3 = 0 'lvalue3
'
'Const cDaily1 = 1 'lvalue1
'Const cDaily2 = 0 'lvalue2
'Const cDaily3 = 0 'lvalue3
'Const cszTask = "RTStation日志清理程序"
'Const cszRunProgramShort = "snlogdel" '用小写,方便比较
'
'
'Private Sub cmdClose_Click()
'    Unload Me
'
'End Sub
'
'Private Sub cmdHelp_Click()
'    DisplayHelp Me, content
'End Sub
'
'Private Sub cmdOK_Click()
'    Dim lTriCount As Long
'    Dim lTaskCount As Long
'    Dim oTask As Task2
'    Dim oTaskSched As New TaskScheduler2
'    Dim oTri As TaskTrigger2
'    Dim i As Long
'    Dim szTaskName As String
'    Dim szTemp As String
'    Dim bIsReSet As Boolean
'
'    bIsReSet = False
'
'    GetTaskInfo
'
'    On Error GoTo TASKERR
''Back:
'    lTaskCount = oTaskSched.GetTaskCount
'    If lTaskCount > 0 Then
'        For i = 1 To lTaskCount
'            szTaskName = oTaskSched.GetTask(i)
'            Set oTask = oTaskSched.Activate(szTaskName, 0)
'            szTemp = LCase(oTask.GetApplicationName)
'            If (InStr(1, szTemp, cszRunProgramShort) <> 0) Or (szTaskName = cszTask & ".job") Then
'                bIsReSet = True
'                Call oTask.SetApplicationName(szRunName)
'                Call oTask.SetParameters(szParam)
'
'                Set oTri = oTask.GetTrigger(0)
'                Call oTri.SetTrigger(TriggerInfo)
'                oTask.Save
'                If IsOsNt() = True Then
'                    Call oTask.SetAccountInformation(szNTUser, szNTPassWord)
'                    oTask.Save
'                End If
''                oTaskSched.Delete szTaskName
''                GoTo Back
'            End If
'        Next i
'    End If
'    If bIsReSet = False Then
'
'        Set oTask = oTaskSched.NewTask(cszTask, 0, 0)
'        Call oTask.SetApplicationName(szRunName)
'        Call oTask.SetParameters(szParam)
'        Set oTri = oTask.CreateTrigger(0)
'        Call oTri.SetTrigger(TriggerInfo)
'        oTask.Save
'
'        If IsOsNt() = True Then
'            Call oTask.SetAccountInformation(szNTUser, szNTPassWord)
'            oTask.Save
'        End If
'    End If
'
'
'
'    i = MsgBox("设置成功!是否退出?", vbInformation + vbYesNo, cszMsg)
'    If i = vbYes Then Unload Me
'
'Exit Sub
'TASKERR:
'    ShowErrorMsg
'End Sub
'
'Private Sub Form_Activate()
'    txtUser.SetFocus
'    txtUser.SelStart = 0
'    txtUser.SelLength = Len(txtUser.Text)
'End Sub
'
'Private Sub Form_Load()
'    If IsOsNt() = True Then
'        Me.Height = 4965
'        fraNt.Enabled = True
'    Else
'        Me.Height = 3885
'        fraNt.Enabled = False
'    End If
'
'    lBeginDay = 1
'    lBeginMouth = 1
'    lBeginYear = 2000
'
'
'    txtUser = g_oActUser.UserID
'    dtpTime.Value = "21:30:00"
'    Dim i As Integer
'    Me.Top = (Screen.Height - Me.ScaleHeight) / 2
'    Me.Left = (Screen.Width - Me.ScaleWidth) / 2
'    cboWeek.AddItem "日"
'    cboWeek.AddItem "一"
'    cboWeek.AddItem "二"
'    cboWeek.AddItem "三"
'    cboWeek.AddItem "四"
'    cboWeek.AddItem "五"
'    cboWeek.AddItem "六"
'    cboWeek.ListIndex = 0
'    For i = 1 To 31
'        cboMonth.AddItem CStr(i)
'    Next i
'    cboMonth.ListIndex = 0
'    optDay_Click
'
'End Sub
'
'Private Sub optDay_Click()
'    If optDay.Value = True Then
'        cboWeek.Enabled = False
'        cboMonth.Enabled = False
'    ElseIf optWeek.Value = True Then
'        cboWeek.Enabled = True
'        cboMonth.Enabled = False
'    Else
'        cboWeek.Enabled = False
'        cboMonth.Enabled = True
'    End If
'End Sub
'
'Private Sub optGarnt_Click()
'    If optGarnt.Value = True Then
'        txtUser.Enabled = True
'        txtPass.Enabled = True
'        txtUser.BackColor = vbWhite
'        txtPass.BackColor = vbWhite
'    Else
'        txtUser.Enabled = False
'        txtPass.Enabled = False
'        txtPass.BackColor = &HE0E0E0
'        txtUser.BackColor = &HE0E0E0
'    End If
'End Sub
'
'Private Sub optMonth_Click()
'    optDay_Click
'End Sub
'
'Private Sub optNow_Click()
'    optGarnt_Click
'End Sub
'
'Private Sub optWeek_Click()
'    optDay_Click
'End Sub
'
'Private Sub GetTaskInfo()
'    Dim szNowUser As String
'    Dim szNowPass As String
'    Dim szTemp As String
'
'    If IsOsNt() = True Then
'        szNTUser = txtNTUser.Text
'        szNTPassWord = txtNTPassword.Text
'    Else
'        szNTUser = Empty
'        szNTPassWord = Empty
'    End If
'
'
'    If optNow.Value = True Then
'        On Error GoTo NormalErr
'        szNowUser = g_oActUser.UserID
'
'        szRunName = frmAutoDelectLog.szRunProgram
'        szParam = szNowUser & "," & g_szPassword & ",t"
'    Else
'        szNowUser = txtUser.Text
'        szNowPass = txtPass.Text
'        szRunName = frmAutoDelectLog.szRunProgram
'        szParam = szNowUser & "," & szNowPass & ",t"
'    End If
'
'    lStartHour = CLng(Format(dtpTime.Value, "hh"))
'    szTemp = Format(dtpTime.Value, "hh:mm")
'    lStartMin = CLng(Right(szTemp, 2))
'
'    If optDay.Value = True Then
'        TaskTriType = TP_TIME_TRIGGER_DAILY
'        TaskType.lValue1 = cDaily1
'        TaskType.lValue2 = cDaily2
'        TaskType.lValue3 = cDaily3
'    ElseIf optWeek.Value = True Then
'        TaskTriType = TP_TIME_TRIGGER_WEEKLY
'        TaskType.lValue1 = 2 ^ CLng(cboWeek.ListIndex)
'        TaskType.lValue2 = cPerWeek2
'        TaskType.lValue3 = cPerWeek3
'
'    Else
'        TaskTriType = TP_TIME_TRIGGER_MONTHLYDATE
'        TaskType.lValue1 = 2 ^ CLng(cboMonth.ListIndex)
'        TaskType.lValue2 = cAllMouth2
'        TaskType.lValue3 = cAllMouth3
'    End If
'
'    TriggerInfo.TriggerType = TaskTriType
'    TriggerInfo.Type = TaskType
'    TriggerInfo.wBeginDay = lBeginDay
'    TriggerInfo.wBeginMonth = lBeginMouth
'    TriggerInfo.wBeginYear = lBeginYear
'    TriggerInfo.wStartHour = lStartHour
'    TriggerInfo.wStartMinute = lStartMin
'
'
'
'Exit Sub
'NormalErr:
'    ShowErrorMsg
'End Sub
'
'Private Sub txtUser_Validate(Cancel As Boolean)
'    Dim i As Integer
'    Dim nLen As Integer
'    Dim bIsExist As Boolean
'
'    If optGarnt.Value = True Then
'        bIsExist = False
'
'        nLen = ArrayLength(g_atUserInfo)
'
'        If nLen > 0 Then
'            For i = 1 To nLen
'                If UCase(txtUser.Text) = UCase(g_atUserInfo(i).UserID) Then
'                    bIsExist = True
'                    Exit For
'                End If
'            Next i
'        End If
'    End If
'    If bIsExist = False Then MsgBox "无此用户,请设定一有效的用户ID,且确保此用户有清理日志的权限!" _
'            & vbCrLf & "从数据安全性考虑,可为此特设一用户,此用户有且仅有清理日志的权限.", vbExclamation, cszMsg
'    Cancel = IIf(bIsExist, False, True)
'
'End Sub
