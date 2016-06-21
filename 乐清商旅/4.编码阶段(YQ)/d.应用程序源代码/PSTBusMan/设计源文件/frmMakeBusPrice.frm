VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmMakeBusPrice 
   BackColor       =   &H00E0E0E0&
   Caption         =   "生成车次票价"
   ClientHeight    =   4620
   ClientLeft      =   2745
   ClientTop       =   2925
   ClientWidth     =   7380
   ControlBox      =   0   'False
   Icon            =   "frmMakeBusPrice.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7380
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   345
      Left            =   3045
      TabIndex        =   1
      Top             =   4170
      Width           =   1320
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "取消(&O)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMakeBusPrice.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5235
      Top             =   2880
   End
   Begin VB.ListBox lstCreateInfo 
      Height          =   2760
      IMEMode         =   1  'ON
      ItemData        =   "frmMakeBusPrice.frx":0166
      Left            =   60
      List            =   "frmMakeBusPrice.frx":0168
      MultiSelect     =   2  'Extended
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   345
      Width           =   7245
   End
   Begin MSComctlLib.ProgressBar CreateProgressBar 
      Height          =   270
      Left            =   1725
      Negotiate       =   -1  'True
      TabIndex        =   2
      Top             =   3720
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   476
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.ProgressBar CreateProgressBarTotal 
      Height          =   270
      Left            =   1725
      Negotiate       =   -1  'True
      TabIndex        =   6
      Top             =   3330
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   476
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次进度:"
      Height          =   180
      Left            =   75
      TabIndex        =   9
      Top             =   3765
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "总进度:"
      Height          =   180
      Left            =   75
      TabIndex        =   8
      Top             =   3375
      Width           =   630
   End
   Begin VB.Label lblProgressTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Left            =   1275
      TabIndex        =   7
      Top             =   3345
      Width           =   315
   End
   Begin VB.Label lblProgress 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Left            =   1275
      TabIndex        =   5
      Top             =   3735
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次票价生成日志文件:"
      Height          =   180
      Left            =   60
      TabIndex        =   4
      Top             =   90
      Width           =   1890
   End
   Begin VB.Label lblLogFileName 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1710
      TabIndex        =   3
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
Attribute VB_Name = "frmMakeBusPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const cnBusVehicleSeat = 30 '生成车次票价时每次生成的数量

Private m_szPriceTableID As String '选择的票价表
Private m_atBusVehicleSeat() As TBusVehicleSeatType   '选择的车次、车型、座位类型
Private m_bnMakeStopStation  As Boolean '是否只生成停靠的站点

Private m_aszBusID() As String '选择的车次
Private m_aszVehicleType() As String '选择的车型
Private m_aszSeatType() As String '选择的座位类型

Private WithEvents m_oRoutePriceTable As RoutePriceTable
Attribute m_oRoutePriceTable.VB_VarHelpID = -1

Private m_szProjectID As String
Private m_bCancelMakePrice As Boolean
Private m_szCreateLogFile As String
Private m_bLogFileValid As Boolean '日志文件
Private m_lCount  As Long
Private m_lCreateCount As Integer



Public Property Let GetBusID(vNewValue() As String)
    m_aszBusID = vNewValue
End Property

Public Property Let GetVehicleType(vNewValue() As String)
    m_aszVehicleType = vNewValue
End Property

Public Property Let GetSeatType(vNewValue() As String)
    m_aszSeatType = vNewValue
End Property

Public Property Let GetPriceTableID(vNewValue As String)
    m_szPriceTableID = vNewValue
End Property

Public Property Let GetIsOnlyStop(vNewValue As Boolean)
    m_bnMakeStopStation = vNewValue
End Property


Private Sub CreatePrice()
    '写入初始日志
    Dim dyStart As Date, dyEnd As Date

    On Error GoTo ErrorHandle
    m_szCreateLogFile = CreateOnlyLogFile
    CloseLogFile
    OpenLogFile m_szCreateLogFile
    If m_bLogFileValid Then
        '如果打开文件成功，显示日志文件名
        lblLogFileName.Caption = m_szCreateLogFile
    End If
    
    dyStart = Now
    RecordLog "================================================================="
    RecordLog "=  车次票价生成日志"
    RecordLog "= ----------------------------------"
    RecordLog "=  使用者：" & g_oActiveUser.UserID & "/" & g_oActiveUser.UserName
    RecordLog "=  开始时间:" & Format(dyStart, cszDateTimeStr)
    RecordLog "=  生成的票价表:" & m_szPriceTableID
    RecordLog "================================================================="
    CloseLogFile
    CreateProgressBar.Min = 0
    MakeBusPrice
    lblProgress.Refresh
    '写入结果信息
    OpenLogFile m_szCreateLogFile
    RecordLog "================================================================="
    RecordLog "生成车次票价结束"
    RecordLog "总共生成车次:" & m_lCreateCount & "个"
    RecordLog "未生成车次:" & m_lCount - m_lCreateCount & "个"
    dyEnd = Now
    RecordLog "结束时间:" & Format(dyEnd, cszTimeStr)
    RecordLog "共使用时间:" & Format(dyEnd - dyStart, "HH小时mm分SS秒")


    CloseLogFile
    cmdCancel.Caption = "确定(&O)"
Over:
    SetNormal
    Exit Sub
ErrorHandle:
    MsgBox err.Description, vbOKOnly + vbCritical
    CloseLogFile
    SetNormal
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo ErrorHandle
    If m_bCancelMakePrice = False Then
        m_bCancelMakePrice = True
        cmdCancel.Caption = "确定(&O)"
        m_oRoutePriceTable.CancelMakeBusPrice = True
    End If
    If cmdCancel.Caption = "确定(&O)" Then
        Unload Me
    End If
ErrorHandle:

End Sub


Private Sub Form_Load()
    m_szProjectID = g_szExePriceTable  '设为执行计划
    m_lCount = 0
    m_lCreateCount = 0
    
    '将数组转换为类型
    m_atBusVehicleSeat = ConvertTypeFromArray(m_aszBusID, m_aszVehicleType, m_aszSeatType)
    
End Sub


Private Sub RecordLog(log As String)
    '动态新增行,并写入日志
    With lstCreateInfo
        .AddItem log
        .ListIndex = .ListCount - 1
        .Refresh
    End With
    If m_bLogFileValid Then
        AddLogToFile log
    End If
End Sub

Private Sub AddLogToFile(log As String)
    '添加日志到文件当中
    On Error Resume Next
    If m_bLogFileValid Then
        Print #1, log
    End If
End Sub

Private Sub CloseLogFile()
    '关闭日志文件
    On Error Resume Next
    '如果打开文件,则关闭文件
    If m_bLogFileValid Then
        Close #1
    End If
End Sub

Private Sub OpenLogFile(LogFile As String)
    '打开日志文件
    On Error GoTo ErrorDo
    m_bLogFileValid = False
    If LogFile = "" Then
        Exit Sub
    End If
    Open LogFile For Append As #1
    m_bLogFileValid = True
    Exit Sub
ErrorDo:
End Sub

Private Function CreateOnlyLogFile() As String
    '创建一个唯一的日志文件名
    Dim LogFileDir As String
    Dim FileName As String
    Dim oReg As New CFreeReg
    
    '取日志文件存放目录
    oReg.Init cszRegKeyProduct & "\Scheme", HKEY_LOCAL_MACHINE, cszRegKeyCompany
    LogFileDir = oReg.GetSetting("MakeBusPrice", "MakeBusPrice", "")
    If LTrim(RTrim(LogFileDir)) = "" Then
        LogFileDir = Environ("temp")
    End If
    LogFileDir = Trim(LogFileDir)
    If Not Right(LogFileDir, 1) = "\" Then
        LogFileDir = LogFileDir + "\"
    End If
    
    FileName = LogFileDir + Format(Date, "YYYYMMDD" & "_" & Format(Now, "HHMMSS")) & ".log"
    If FileIsExist(FileName) Then
        FileName = LogFileDir + Format(Date, "YYYYMMDD" & "_" & Format(Now, "HHMMSS")) & ".log"
    End If
    CreateOnlyLogFile = FileName
End Function


Private Sub lstCreateInfo_DblClick()
    If Left(lstCreateInfo.Text, 1) = "[" And Right(lstCreateInfo.Text, 2) <> "成功" Then
       MsgBox lstCreateInfo.Text, vbOKOnly + vbInformation, Me.Caption
    End If
End Sub

Private Sub m_oRoutePriceTable_SetMakeBusPriceStatus(ByVal lStatus As String)
    '设置生成状态
    OpenLogFile m_szCreateLogFile
    If Left(lStatus, 1) = "[" And Right(lStatus, 2) = "成功" Then
        m_lCreateCount = m_lCreateCount + 1
        m_lCount = m_lCount + 1
        RecordLog lStatus
    ElseIf Left(lStatus, 1) = "[" And Right(lStatus, 2) <> "成功" Then
        m_lCount = m_lCount + 1
        RecordLog lStatus
    Else
        MsgBox lStatus, vbInformation + vbOKOnly, Me.Caption
        RecordLog ""
        RecordLog "********注意：********" & lStatus & "！"
    End If
  CloseLogFile
End Sub

Private Sub m_oRoutePriceTable_SetProgressRange(ByVal lRange As Variant)
    '设置进度条总长
    CreateProgressBar.Max = lRange
End Sub

Private Sub m_oRoutePriceTable_SetProgressValue(ByVal lValue As Variant)
    '设置进度条
    With CreateProgressBar
        If lValue > .Min Then
            .Value = lValue
            lblProgress.Caption = str(Int((lValue / .Max) * 100)) & "%"
        End If
    End With
End Sub



Private Sub MakeBusPrice()
    '真正的生成票价
    'On Error GoTo ErrorHandle

'    Dim atBusVehicleSeat() As TBusVehicleSeatType
    Dim atTempTBusVehicleSeat() As TBusVehicleSeatType
    Dim aszBusID() As String
    Dim nTemp As Long
    Dim i As Long, j As Long, n As Long
    Dim nBusVehicleSeat As Long
    Dim oTicketPriceMan As New TicketPriceMan
    
    SetBusy
    Set m_oRoutePriceTable = New RoutePriceTable
    oTicketPriceMan.Init g_oActiveUser
    m_oRoutePriceTable.Init g_oActiveUser

'    atBusVehicleSeat = m_atBusVehicleSeat
    
    nBusVehicleSeat = ArrayLength(m_atBusVehicleSeat)
    
    m_oRoutePriceTable.Identify m_szPriceTableID
    nTemp = 0
    '设置进度条
    lblProgressTotal.Visible = True
    CreateProgressBarTotal.Max = nBusVehicleSeat
    If nBusVehicleSeat > cnBusVehicleSeat Then
        For i = 1 To nBusVehicleSeat
            If i >= nTemp + cnBusVehicleSeat Then
                If m_atBusVehicleSeat(i).szbusID <> m_atBusVehicleSeat(nTemp + cnBusVehicleSeat).szbusID Then
                    ReDim atTempTBusVehicleSeat(1 To i - nTemp - 1)
                    For j = 1 To i - nTemp - 1
                        For n = nTemp + j To i - 1
                            atTempTBusVehicleSeat(j) = m_atBusVehicleSeat(n)
                            Exit For
                        Next n
                    Next j
                    m_oRoutePriceTable.MakeBusPrice atTempTBusVehicleSeat, True, m_bnMakeStopStation
                    nTemp = i - 1
                    lblProgressTotal.Caption = str(Int((nTemp / CreateProgressBarTotal.Max) * 100)) & "%"
                    CreateProgressBarTotal.Value = nTemp
                End If
            ElseIf nTemp + cnBusVehicleSeat > nBusVehicleSeat Then
                i = nBusVehicleSeat
                ReDim atTempTBusVehicleSeat(1 To i - nTemp)
                For j = 1 To i - nTemp
                    For n = nTemp + j To i
                        atTempTBusVehicleSeat(j) = m_atBusVehicleSeat(n)
                        Exit For
                    Next n
                Next j
                m_oRoutePriceTable.MakeBusPrice atTempTBusVehicleSeat, True, m_bnMakeStopStation
                lblProgressTotal.Caption = str(Int((nBusVehicleSeat / CreateProgressBarTotal.Max) * 100)) & "%"
                CreateProgressBarTotal.Value = nBusVehicleSeat
            End If
        Next i
    Else
        m_oRoutePriceTable.MakeBusPrice m_atBusVehicleSeat, True, m_bnMakeStopStation
        lblProgressTotal.Caption = str(Int((nBusVehicleSeat / CreateProgressBarTotal.Max) * 100)) & "%"
        CreateProgressBarTotal.Value = nBusVehicleSeat
    End If
    CreateProgressBarTotal.Value = CreateProgressBarTotal.Max
    lblProgressTotal.Caption = "100%"
    SetNormal
    Set oTicketPriceMan = Nothing
    Exit Sub
ErrorHandle:
    Set oTicketPriceMan = Nothing
    SetNormal
    cmdCancel.Caption = "确定(&O)"
    MsgBox err.Description
End Sub

Private Sub Timer1_Timer()
    SetBusy
    Timer1.Enabled = False
    CreatePrice
End Sub

