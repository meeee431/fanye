VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMake 
   BackColor       =   &H00FFFFC0&
   Caption         =   "生成环境 - 1999年12月12日"
   ClientHeight    =   4500
   ClientLeft      =   3360
   ClientTop       =   3300
   ClientWidth     =   7380
   Icon            =   "frmMake.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7380
   Begin PSTRunRev.cSysTray cSysTray1 
      Left            =   330
      Top             =   3705
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "frmMake.frx":16AC2
      TrayTip         =   "VB 5 - SysTray Control."
   End
   Begin VB.PictureBox ptMake 
      Height          =   465
      Index           =   1
      Left            =   5595
      Picture         =   "frmMake.frx":187CC
      ScaleHeight     =   405
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   -15
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox ptMake 
      Height          =   495
      Index           =   0
      Left            =   6330
      Picture         =   "frmMake.frx":2F28E
      ScaleHeight     =   435
      ScaleWidth      =   540
      TabIndex        =   4
      Top             =   30
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   3045
      TabIndex        =   1
      Top             =   3735
      Width           =   1200
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5235
      Top             =   3015
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5955
      Top             =   2955
   End
   Begin VB.ListBox lstCreateInfo 
      Appearance      =   0  'Flat
      Height          =   2910
      ItemData        =   "frmMake.frx":45D50
      Left            =   105
      List            =   "frmMake.frx":45D52
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   345
      Width           =   7155
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   2
      Top             =   4185
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6350
            Text            =   "生成车次运行环境"
            TextSave        =   "生成车次运行环境"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "2016-3-12"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "13:57"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar CreateProgressBar 
      Height          =   300
      Left            =   645
      Negotiate       =   -1  'True
      TabIndex        =   3
      Top             =   3330
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   60
      TabIndex        =   8
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "环境生成日志文件:"
      Height          =   180
      Left            =   60
      TabIndex        =   7
      Top             =   90
      Width           =   1530
   End
   Begin VB.Label lblLogFileName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   1710
      TabIndex        =   6
      Top             =   90
      Width           =   90
   End
   Begin VB.Menu mnu_show 
      Caption         =   "显示"
      Visible         =   0   'False
      Begin VB.Menu mnu_normalshow 
         Caption         =   "显示生成窗体(&M)"
      End
   End
End
Attribute VB_Name = "frmMake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Base 1

'--------------------------------------
'以托盘程序显示图标
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Dim m_szCreateLogFile As String  '1生成日志文件的目录
Dim m_bPromptWhenError As Boolean '2是否提示错误
Dim m_bEndExit As Boolean '3程序运行完成后自动退出,FALSE 退出，TRUE不退出
Dim m_bLogFileValid As Boolean '4日志文件
Dim m_dtRunDate As Date '5运行日期
Dim m_bStopMake As Boolean '6停班车次是否生成
Public m_bShowIcon As Boolean  '9是否已托盘程序显示
Dim m_szaScheduledBuses() As String '车次代码
Dim m_szMsgTitle As String

Dim CancelHasPress As Boolean
Private m_oScheme As New REScheme
Private m_oProject As New BusProject
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Const cszSystemParam = " system_param_info"
Private Const cszLocalUnitID = "LocalUnitID" '本单位代码
Private Sub Command1_Click()
Dim aa As New REScheme
  aa.Init g_oActiveUser
  aa.MakeRunEvironment "2001-04-21", "k6", True
End Sub

Private Sub cSysTray1_MouseDblClick(Button As Integer, Id As Long)
    Me.WindowState = vbNormal
    Me.Show
'    cSysTray1.InTray = False
End Sub

Private Sub cSysTray1_MouseUp(Button As Integer, Id As Long)
    If Button = vbRightButton Then
        Me.PopupMenu mnu_show
    End If
End Sub

Private Sub Form_Load()
    Dim szTemp As String
    'Dim szCmdLine As String
    Dim nPreSellDate As Integer '预售日期
    Dim nCurrentDate As Date
    Dim oSystem As New SystemParam
    Dim nErrorClass As Integer

    On Error GoTo ErrorDo
    '----------
    oSystem.Init g_oActiveUser
    nPreSellDate = oSystem.PreSellDate
    Time = oSystem.NowDateTime
    Date = oSystem.NowDate
    Set oSystem = Nothing
    m_oScheme.Init g_oActiveUser
    m_oProject.Init g_oActiveUser
    '读取命令行参数
    '参数(用逗号隔开,车次列表外加中括号):
    '   1.UserID:用户名
    '   2.Password:用户口令
    '   3.([RunDate]):生成计划日期（缺省为预售天数的最后一天）
    '   4.([[车次1],[车次2],...])：缺省为空（所有车次）
    '   5.([PromptWhenError])错误提示标志("F"不提示,"T"提示[缺省])
    '   6.([AppEndExit])程序运行完成后是否退出("F"退出,"T"不退出[缺省])
    '   7.([CreateLogFile])生成信息文件名(缺省自动创建)
    '   8.([StopMake])停班车次生成("F"不生成，"T"生成，缺省不生成)
    '   9.([ISTray])是否以托盘上的小图显示程序在运行，此时窗口hide并且mininize
    '                               (缺省为不显示为小图标"F"不显示,"T"显示)
        
    m_szMsgTitle = "生成车次运行环境"

    '取用户名
    szTemp = LeftAndRight(szCmdLine, True, ",")
    szCmdLine = LeftAndRight(szCmdLine, False, ",")
    
    '取用户口令
    szTemp = LeftAndRight(szCmdLine, True, ",")
    szCmdLine = LeftAndRight(szCmdLine, False, ",")
    
    '取生成计划日期
    szTemp = LeftAndRight(szCmdLine, True, ",")
    szCmdLine = LeftAndRight(szCmdLine, False, ",")
    If LTrim(szTemp) = "" Then
        If DateDiff("h", Time, CDate("23:59:59")) <= 12 Then '如果是第二天早晨生成则生成日期不加一
              nPreSellDate = nPreSellDate + 1
        End If
        m_dtRunDate = DateAdd("d", nPreSellDate, Date)
    Else
        m_dtRunDate = CDate(szTemp)
    End If
    
    Me.Caption = "生成环境——" & Format(m_dtRunDate, "YYYY年MM月DD日")
    '取生成车次
    If Left(szCmdLine, 1) = "[" Then
        '车次由[]包括
        On Error GoTo BusErr
        szCmdLine = LeftAndRight(szCmdLine, False, "[")
        szTemp = LeftAndRight(szCmdLine, True, "]")
        szCmdLine = LeftAndRight(szCmdLine, False, "]")
        '从[1020,1232,12312,2343]的字符串中获得车次
        m_szaScheduledBuses = GetScheduledBuses(szTemp)
    End If
    
    '取错误提示标志
    szCmdLine = LeftAndRight(szCmdLine, False, ",")
    '是否提示错误
    szTemp = LeftAndRight(szCmdLine, True, ",")
    szCmdLine = LeftAndRight(szCmdLine, False, ",")
    szTemp = UCase(RTrim(szTemp))
    If szTemp = "T" Then
        m_bPromptWhenError = True
    Else
        m_bPromptWhenError = False
    End If
    
    '取是否退出程序
    szTemp = LeftAndRight(szCmdLine, True, ",")
    szCmdLine = LeftAndRight(szCmdLine, False, ",")
    szTemp = UCase(LTrim(RTrim(szTemp)))
    If szTemp = "T" Then
        m_bEndExit = True
    Else
        m_bEndExit = False
    End If

    '取日志文件
    szTemp = LeftAndRight(szCmdLine, True, ",")
    If szTemp = "" Then
        m_szCreateLogFile = CreateOnlyLogFile
    Else
        m_szCreateLogFile = szTemp
    End If
    
    '取是否停班生成
    szCmdLine = LeftAndRight(szCmdLine, True, ",")
    szTemp = LeftAndRight(szCmdLine, True, ",")
    If szTemp = "F" Then
        m_bStopMake = False
    Else
        m_bStopMake = True
    End If
    
    '取是以托盘程序生成
    szCmdLine = LeftAndRight(szCmdLine, False, ",")
    szTemp = LeftAndRight(szCmdLine, True, ",")
    If szTemp = "T" Then
        Me.WindowState = vbMinimized
    Else
        m_bShowIcon = False
    End If

Exit Sub
DateErr:
    MsgBox "参数3错误:指定的生成运行环境日期格式不正确(YYYY-MM-DD[年-月-日])", _
        vbOKOnly + vbCritical
    GoTo ErrorDo
BusErr:
    MsgBox "参数4错误:指定生成计划车次格式错误.格式" _
        & "[ [车次代码1],[车次代码2],.. ],例:[110,111]", _
        vbOKOnly + vbCritical

ErrorDo:
    Timer1.Enabled = False
    Timer2.Enabled = False
    End
End Sub

Private Sub CreateRunEnvirment()
    Dim szaAllBus() As String
    Dim ErrSource As String
    Dim oRegularScheme As New RegularScheme
    Dim oProject As New BusProject
    Dim bCreateOk As Integer
    Dim szPlanID As String
    Dim nCount  As Long, i As Long
    Dim nCreateCount As Integer
    Dim tmStart, tmEnd As Date
    Dim szTableID As String
    StatusBar1.Panels(2).Text = "开始时间:" & Format(Time, "HH:MM")
    ErrSource = "生成车次运行环境"
    On Error GoTo ErrorDo
    oRegularScheme.Init g_oActiveUser
    oProject.Init g_oActiveUser
    szPlanID = oRegularScheme.GetExecuteBusProject(m_dtRunDate).szProjectID
    szTableID = oRegularScheme.GetRunPriceTableEx(m_dtRunDate)
    
    
    
    oProject.Identify ' szPlanID
    If ArrayLength(m_szaScheduledBuses) = 0 Then
        m_oProject.Identify 'szPlanID
        szaAllBus = m_oProject.GetAllBus
        nCount = ArrayLength(szaAllBus)
        ReDim m_szaScheduledBuses(1 To nCount) As String
        For i = 1 To nCount
            m_szaScheduledBuses(i) = szaAllBus(i, 1)
        Next
    End If
    nCount = ArrayLength(m_szaScheduledBuses)
    CloseAll
    OpenLogFile m_szCreateLogFile
    
    If m_bLogFileValid Then
        '如果打开文件成功，显示日志文件名
        lblLogFileName.Caption = m_szCreateLogFile
    End If
    
    tmStart = Now
    RecordLog "================================================================="
    RecordLog "=  生成环境日志记录"
    RecordLog "= ----------------------------------"
    RecordLog "=  使用者：" & g_oActiveUser.UserID & "/" & g_oActiveUser.UserName
    RecordLog "=  生成车次日期:" & Format(m_dtRunDate, "YYYY-MM-DD")
    RecordLog "=  当前时间:" & Format(tmStart, "YYYY-MM-DD HH:MM:SS")
    RecordLog "=  运行计划:" & Format(szPlanID) & "/" & oProject.ProjectName
    RecordLog "=  执行票价表:" & szTableID
    RecordLog "================================================================="
    CreateProgressBar.Min = 0
    CreateProgressBar.Max = nCount
        
    '开始显示生成动画
    Timer2.Enabled = True
    
    For i = 1 To nCount
        If CancelHasPress Then
            CancelHasPress = False
            If MsgBox("生成运行环境还未结束，停止生成吗?", _
                vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then
                
                RecordLog vbCrLf
                RecordLog "**生成运行环境被用户中断**"
                RecordLog vbCrLf
                
                Timer2.Enabled = False
                GoTo Report
            End If
            Timer2.Enabled = True
        End If
        bCreateOk = CreateRunEnvirmentBus(i)
        lblProgress.Caption = Str(Int(100 * i / nCount)) & "%"
        lblProgress.Refresh
        If bCreateOk = 1 Then
            nCreateCount = nCreateCount + 1
        End If
        If bCreateOk = 3 Then
            i = i - 1
        End If
        CreateProgressBar.Value = i
    Next
    lblProgress.Caption = "100%"
    lblProgress.Refresh

Report:
          Dim m_oSndMsg As New CSendMsg
            Dim oParam As New SystemParam
            Dim aszValues(1 To 2) As String
            aszValues(1) = ""
            aszValues(2) = ToDBDate(m_dtRunDate)
            m_oSndMsg.Unit = GetUnitID()
            m_oSndMsg.MsgSource = "SNRunEnv"
            
            m_oSndMsg.SendMsg oParam.NowDateTime, "MakeEnv", aszValues
    RecordLog "================================================================="
    RecordLog "运行环境生成结束"
    RecordLog "总共生成车次:" & nCreateCount & "个"
    RecordLog "未生成车次:" & nCount - nCreateCount & "个"
    tmEnd = Now
    RecordLog "结束时间:" & Format(tmEnd, "HH:MM:SS")
    RecordLog "共使用时间:" & Format(tmEnd - tmStart, "HH小时MM分SS秒")
    
    
    CloseLogFile
    Timer2.Enabled = False
    cmdCancel.Caption = "确定"
    If Not cSysTray1.InTray And Not m_bEndExit Then
        Exit Sub
    End If
Over:
    cmdCancel_Click
    Exit Sub
ErrorDo:
    MsgBox err.Description, vbOKOnly + vbCritical
    CloseLogFile
    GoTo Over
End Sub

'生成某车次并记录生成1成功，2错误，3重试
Public Function CreateRunEnvirmentBus(Index As Long) As Integer
    Dim vbMsg As VbMsgBoxResult
    Dim ErrString As String
    Dim bCreateOk As Integer
    Dim nErrNumber As Long
    Dim szErrDescription As String
    Static nHasPrompt, nPromptTime As Integer
    
    On Error GoTo here
    '初始化设置开始提醒时的错误次数
    If nPromptTime < 2 Then nPromptTime = 2
    
    bCreateOk = 1
    ErrString = "成功    "
    m_oScheme.MakeRunEvironment m_dtRunDate, m_szaScheduledBuses(Index), m_bStopMake
    RecordLog m_szaScheduledBuses(Index) & "[" & ErrString & "]"
ErrContinue:
    DoEvents
    CreateRunEnvirmentBus = bCreateOk
    Exit Function
here:
    bCreateOk = 2
    nErrNumber = err.Number
    szErrDescription = err.Description
    ErrString = m_szaScheduledBuses(Index) & "[  未生成]" & _
        " * 错误描述:(" & Trim(Str(nErrNumber)) & ")" & Trim(szErrDescription) & " *"
    RecordLog ErrString
    If m_bPromptWhenError Then
        ErrString = "车次" & m_szaScheduledBuses(Index) & "未生成！" & vbCrLf & _
            Trim(szErrDescription) & "(" & Trim(Str(nErrNumber)) & ")"
        vbMsg = MsgBox(ErrString, vbExclamation + vbAbortRetryIgnore + vbDefaultButton3)
        Select Case vbMsg
               Case vbAbort
                   CancelHasPress = True
               Case vbRetry
                   CreateRunEnvirmentBus = 3
                   Exit Function
               Case vbIgnore
                   If nHasPrompt >= nPromptTime - 1 Then
                        If MsgBox("以后不再提示生成错误？", vbQuestion + vbYesNo) = vbYes Then
                            m_bPromptWhenError = False
                        End If
                        nHasPrompt = 0
                        nPromptTime = nPromptTime + 1
                   End If
                   nHasPrompt = nHasPrompt + 1
               Exit Function
        End Select
    Else
        GoTo ErrContinue
    End If
End Function

Private Sub cmdCancel_Click()
    CancelHasPress = True
    Timer2.Enabled = False
    If cmdCancel.Caption = "确定" Then
        If cSysTray1.InTray Then cSysTray1.InTray = False
        End
    End If
End Sub

Private Sub Form_Resize()
    If WindowState = vbNormal Then
        Me.Height = 4890
        Me.Width = 7395
    End If
    If WindowState = vbMinimized Then
        m_bShowIcon = True
        Me.Hide
        Set cSysTray1.TrayIcon = ptMake(0).Picture
        If Not cSysTray1.InTray Then
            cSysTray1.TrayTip = "正在生成" & Format(m_dtRunDate, "YYYY年MM月DD日") & "运行环境..." & Chr$(0)
            cSysTray1.InTray = True
        End If
    Else
        cSysTray1.InTray = False
        m_bShowIcon = False
    End If
End Sub

Private Sub lstCreateInfo_DblClick()
MsgBox lstCreateInfo.Text, vbInformation + vbOKOnly, "生成信息"
End Sub

Private Sub mnu_normalshow_Click()
    cSysTray1_MouseDblClick 0, 0
End Sub

Private Sub timer1_Timer()
    Timer1.Enabled = False
    CreateRunEnvirment
End Sub

Private Sub Timer2_Timer()
 Static i As Long, img As Long
   DoEvents
   If cSysTray1.InTray Then
        cSysTray1.TrayTip = "正在生成" & _
            Format(m_dtRunDate, "YYYY年MM月DD日") & _
            "运行环境..." & Trim(lblProgress.Caption) & Chr$(0)
        Set cSysTray1.TrayIcon = ptMake(i).Picture
    End If
   Me.Icon = ptMake(i).Picture
   i = i + 1
   If i = 2 Then i = 0
End Sub

Private Function GetScheduledBuses(strBus As String) As String()
'将车次字符串转换成车次数组
    Dim strTemp As String
    Dim ScheduledBuses() As String
    strTemp = LTrim(RTrim(strBus))
    strTemp = LTrim(RTrim(LeftAndRight(strBus, True, ",")))
    If Not strTemp = "" Then
        ReDim ScheduledBuses(1)
        ScheduledBuses(1) = strTemp
    Else
        Exit Function
    End If
    If InStr(1, strBus, ",") = 0 Then
        strBus = ""
    Else
        strBus = LTrim(RTrim(LeftAndRight(strBus, False, ",")))
    End If
    If strBus = "" Then
        GetScheduledBuses = ScheduledBuses
        Exit Function
    End If
    Do While True
        strTemp = LTrim(RTrim(LeftAndRight(strBus, True, ",")))
        If Not strTemp = "" Then
            ReDim Preserve ScheduledBuses(UBound(ScheduledBuses) + 1)
            ScheduledBuses(UBound(ScheduledBuses)) = strTemp
        Else
            GetScheduledBuses = ScheduledBuses
            Exit Function
        End If
        
        strBus = LTrim(RTrim(LeftAndRight(strBus, False, ",")))
        If strBus = "" Then
            GetScheduledBuses = ScheduledBuses
            Exit Function
        End If
    Loop
    GetScheduledBuses = ScheduledBuses
End Function

Private Sub RecordLog(log As String)
    With lstCreateInfo
        .AddItem log
        .ListIndex = .ListCount - 1
        .Refresh
    End With
    If m_bLogFileValid Then
        AddLogToFile log
    End If
End Sub

Private Sub CloseAll()
'如果打开文件,则关闭文件
    If m_bLogFileValid Then
        CloseLogFile
    End If
End Sub
Private Sub AddLogToFile(log As String)
    On Error Resume Next
    If m_bLogFileValid Then
        Print #1, log
    End If
End Sub
Private Sub CloseLogFile()
    On Error Resume Next
    If m_bLogFileValid Then
        Close #1
    End If
End Sub

Private Sub OpenLogFile(LogFile As String)
    On Error GoTo ErrorDo
    m_bLogFileValid = False
    If LogFile = "" Then
        Exit Sub
    End If
    Open LogFile For Append As #1
    m_bLogFileValid = True
    Exit Sub
ErrorDo:
    m_bLogFileValid = False
    If m_bPromptWhenError Then
        If MsgBox("不能生成日志文件" & LogFile & "!" & vbCrLf _
            & "错误:" & err.Description _
            & vbCr & err.Description & vbCr & "继续吗?", _
            vbYesNo + vbInformation, "生成车次运行环境") = vbNo Then
            End
        End If
    End If
End Sub

Private Function CreateOnlyLogFile() As String
'创建一个唯一的日志文件名
    Dim LogFileDir As String
    Dim FileName As String
    Dim m_oReg As New CFreeReg

    '取日志文件存放目录
    LogFileDir = m_szCreateLogFile
    m_oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    LogFileDir = m_oReg.GetSetting("MakeEn\MakeEnvDesktop", "LogDirect", "")
    If LTrim(RTrim(LogFileDir)) = "" Then
        LogFileDir = Environ("temp")
    End If
    LogFileDir = Trim(LogFileDir)
    If Not Right(LogFileDir, 1) = "\" Then
        LogFileDir = LogFileDir + "\"
    End If
    FileName = LogFileDir + Format(Date, "YYYYMMDD") & ".REN"
    If FileIsExist(FileName) Then
        Dim n As Integer
        Do While FileIsExist(FileName)
            n = n + 1
            FileName = LogFileDir + Format(Date, "YYYYMMDD") & "-" & _
                RTrim(CStr(n)) & ".REN"
        Loop
    End If
    CreateOnlyLogFile = FileName
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = True
    cmdCancel_Click
End Sub

Private Sub ShowIcon()
    Set cSysTray1.TrayIcon = ptMake(0).Picture
    cSysTray1.InTray = True
    cSysTray1.TrayTip = "正在生成" & Format(m_dtRunDate, "YYYY年MM月DD日") & "运行环境..." & Chr$(0)
End Sub



'内部用得到本单位的代码
Public Function GetUnitID() As String
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection

    oDb.ConnectionString = GetConnectionStr
'    '=========================================================================
'    '数据库
'    '-------------------------------------------------------------------------
'    szSql = "SELECT * FROM System_param_info WHERE parameter_name='" & cszLocalUnitID & "'"
'    '=========================================================================
    '=========================================================================
    '嘉兴数据库
    '-------------------------------------------------------------------------
    szSql = "SELECT * FROM " & cszSystemParam & " WHERE parameter_name='" & cszLocalUnitID & "'"
    '=========================================================================

    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 1 Then
        GetUnitID = FormatDbValue(rsTemp!parameter_value)
    End If
End Function
