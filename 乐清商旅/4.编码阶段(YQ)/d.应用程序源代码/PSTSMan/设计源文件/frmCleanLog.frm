VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCleanLog 
   BackColor       =   &H00FFFFFF&
   Caption         =   "导出日志"
   ClientHeight    =   3915
   ClientLeft      =   1530
   ClientTop       =   2700
   ClientWidth     =   6870
   HelpContextID   =   5002601
   Icon            =   "frmCleanLog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6870
   Begin MSComctlLib.ListView lvExport 
      Height          =   1875
      Left            =   105
      TabIndex        =   15
      Top             =   1935
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3307
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   300
      Left            =   5655
      TabIndex        =   14
      Top             =   945
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "关闭(&C)"
      Height          =   300
      Left            =   5655
      TabIndex        =   13
      Top             =   570
      Width           =   1095
   End
   Begin VB.CommandButton cmdOperate 
      Caption         =   "导出(&E)"
      Height          =   300
      Left            =   5655
      TabIndex        =   12
      Top             =   195
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   105
      TabIndex        =   2
      Top             =   1005
      Width           =   5415
      Begin VB.CheckBox chkCleanOpe 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "操作日志导出"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   150
         TabIndex        =   5
         Top             =   0
         Value           =   1  'Checked
         Width           =   1410
      End
      Begin MSComCtl2.DTPicker dtpOpeFirst 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dddddd aaaa"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   270
         Left            =   1635
         TabIndex        =   9
         Top             =   345
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   476
         _Version        =   393216
         Format          =   67698688
         CurrentDate     =   36530
      End
      Begin MSComCtl2.DTPicker dtpOpeSecond 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dddddd aaaa"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   270
         Left            =   3585
         TabIndex        =   10
         Top             =   345
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   476
         _Version        =   393216
         Format          =   67698688
         CurrentDate     =   36530
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   3300
         TabIndex        =   11
         Top             =   405
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "选择日期(&O):"
         Height          =   180
         Left            =   495
         TabIndex        =   4
         Top             =   390
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   5415
      Begin MSComCtl2.DTPicker dtpLoginFirst 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   270
         Left            =   1620
         TabIndex        =   6
         Top             =   345
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   476
         _Version        =   393216
         Format          =   67698688
         CurrentDate     =   36530
      End
      Begin VB.CheckBox chkCleanLogin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "登录日志导出"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   165
         TabIndex        =   1
         Top             =   0
         Value           =   1  'Checked
         Width           =   1395
      End
      Begin MSComCtl2.DTPicker dtpLoginSecond 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dddddd aaaa"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   270
         Left            =   3585
         TabIndex        =   7
         Top             =   345
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   476
         _Version        =   393216
         Format          =   67698688
         CurrentDate     =   36530
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   3300
         TabIndex        =   8
         Top             =   390
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "选择日期(&L):"
         Height          =   180
         Left            =   480
         TabIndex        =   3
         Top             =   390
         Width           =   1080
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5985
      Top             =   2010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCleanLog.frx":014A
            Key             =   "success"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCleanLog.frx":0472
            Key             =   "Doing"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCleanLog.frx":079A
            Key             =   "Error"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCleanLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim szDir As String
    
    
Private Sub chkCleanLogin_Click()
    If chkCleanLogin.Value = vbChecked Then
        dtpLoginFirst.Enabled = True
        dtpLoginSecond.Enabled = True
    Else
        dtpLoginFirst.Enabled = False
        dtpLoginSecond.Enabled = False
    End If
    cmdOperate.Enabled = chkCleanLogin.Value = vbChecked Or chkCleanOpe.Value = vbChecked

End Sub

Private Sub chkCleanOpe_Click()
    If chkCleanOpe.Value = vbChecked Then
        dtpOpeFirst.Enabled = True
        dtpOpeSecond.Enabled = True
    Else
        dtpOpeFirst.Enabled = False
        dtpOpeSecond.Enabled = False
    End If
    cmdOperate.Enabled = chkCleanLogin.Value = vbChecked Or chkCleanOpe.Value = vbChecked


End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me, content
End Sub

Private Sub cmdOperate_Click()
'    Me.Height = 3870
    Dim aszTemp() As String
    Dim nLenDays As Integer
    Dim dtStart As Date, dtEnd As Date
    Dim i As Integer
    Dim bTempLog As Boolean '标志导出文件成功
    Dim bTempOpe As Boolean '标志导出文件成功
    Dim lTemp As Long
    Dim n As Integer

    lvExport.ListItems.Clear

    '登录日志
    If chkCleanLogin.Value = vbChecked Then

        dtStart = dtpLoginFirst.Value
        dtEnd = dtpLoginSecond.Value
        If dtStart > dtEnd Or dtEnd > Now Or dtStart > Now Then
            MsgBox "请选择正确日期[登录日志]导出!", vbInformation, cszMsg
            Exit Sub
        End If
        
        nLenDays = dtEnd - dtStart + 1
        For i = 0 To nLenDays - 1
            lvExport.ListItems.Add , , "导出" & Format((dtStart + i), "YYYY-MM-DD") & "的登录日志...", "Doing", "Doing"
            aszTemp = TransToStrs(dtStart + i, False)
            n = lvExport.ListItems.Count
            If ArrayLength(aszTemp) > 0 Then
                bTempLog = CreatLogFile(szDir, dtStart + i, aszTemp, False)
                If bTempLog = True Then
                    lvExport.ListItems.Item(n).SmallIcon = "success"
                    lvExport.ListItems.Item(n).SubItems(1) = "成功"
                Else
                    lvExport.ListItems.Item(n).SmallIcon = "Error"
                    lvExport.ListItems.Item(n).SubItems(1) = "失败"
                End If
            Else
                lvExport.ListItems.Item(n).SmallIcon = "success"
                lvExport.ListItems.Item(n).SubItems(1) = "无当天登录日志"
            End If
        Next i

    End If

    
    '操作日志
    If chkCleanOpe.Value = vbChecked Then
        dtStart = dtpOpeFirst.Value
        dtEnd = dtpOpeSecond.Value
        If dtStart > dtEnd Or dtEnd > Now Or dtStart > Now Then
            MsgBox "请选择正确的日期[操作日志]导出!", vbInformation, cszMsg
            Exit Sub
        End If
        nLenDays = dtEnd - dtStart + 1
        For i = 0 To nLenDays - 1
            lvExport.ListItems.Add , , _
            "导出" & Format((dtStart + i), "YYYY-MM-DD") & "的操作日志...", _
            "Doing", "Doing"
            aszTemp = TransToStrs(dtStart + i)
            n = lvExport.ListItems.Count
            If ArrayLength(aszTemp) > 0 Then
                bTempOpe = CreatLogFile(szDir, dtStart + i, aszTemp)
                If bTempOpe = True Then
                    lvExport.ListItems.Item(n).SmallIcon = "success"
                    lvExport.ListItems.Item(n).SubItems(1) = "成功"
                Else
                    lvExport.ListItems.Item(n).SmallIcon = "Error"
                    lvExport.ListItems.Item(n).SubItems(1) = "失败"
                End If
            Else
                lvExport.ListItems.Item(n).SmallIcon = "success"
                lvExport.ListItems.Item(n).SubItems(1) = "无当天操作日志"
            End If
        Next i
    End If
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Load()
    dtpLoginFirst.Value = Now - 7
    dtpLoginSecond.Value = Now
    dtpOpeFirst.Value = Now - 7
    dtpOpeSecond = Now
    
    Me.Top = (Screen.Height - Me.ScaleHeight) / 2
    Me.Left = (Screen.Width - Me.ScaleWidth) / 2
    
    
    With frmAutoDelectLog
        szDir = .szDir
    End With


End Sub




Private Function CreatLogFile(LogDir As String, DayForName As Date, Texts() As String, Optional IsOpeLog As Boolean = True) As Boolean
    Dim oTextStream As TextStream
    Dim oFileSystem As FileSystemObject
    Dim nLenTexts As Integer, i As Integer
    Dim szFileName As String
    Dim szTemp As String
    
    CreatLogFile = True
    If LogDir = Empty Then
        CreatLogFile = False
    Else
        szTemp = Right(LogDir, 1)
        If szTemp <> "\" Then
            LogDir = LogDir & "\"
        End If
        If IsOpeLog = True Then
            szFileName = "Op" & CStr(Format(DayForName, "YYYYMMDD")) & ".log"
        Else
            szFileName = "Lo" & CStr(Format(DayForName, "YYYYMMDD")) & ".log"
        End If
        szFileName = LogDir & szFileName
        nLenTexts = ArrayLength(Texts)
        If nLenTexts > 0 Then
            Set oFileSystem = CreateObject("Scripting.FileSystemObject")
            On Error GoTo ErrorHandle
            Set oTextStream = oFileSystem.CreateTextFile(szFileName)
            For i = 1 To nLenTexts
                oTextStream.WriteLine (Texts(i))
            Next i
            oTextStream.Close
            
        End If
    End If
Exit Function
ErrorHandle:
    CreatLogFile = False
End Function

Private Function TransToStrs(OneDay As Date, Optional bIsOpe As Boolean = True) As String()
    Dim rsLoginInfo As New Recordset
    Dim dtStart As Date, dtEnd As Date
    Dim tmStart As Date, tmEnd As Date
    Dim g_aszUser() As String '空
    Dim aszWS_FG() As String '空
    Dim aszTemp() As String
    Dim i As Integer
    
    dtStart = CDate(Format(OneDay, "yyyy-mm-dd"))
    dtEnd = dtStart
    tmStart = "00:00:00"
    tmEnd = "23:59:59"
    If bIsOpe = False Then
        On Error GoTo ErrorHandle
        Set rsLoginInfo = g_oSysMan.GetLoginLogRs(g_aszUser(), dtStart, dtStart, tmStart, tmEnd, aszWS_FG)
        
        If rsLoginInfo.RecordCount > 0 Then
            ReDim aszTemp(1 To rsLoginInfo.RecordCount + 1)
            '标题
            aszTemp(1) = "登录事件代码" & vbTab & "登录时间" & vbTab & "用户代码" & vbTab & "登录工作站" & vbTab & "IP地址" & vbTab & "注销时间" & vbTab & "注销类型"
            '内容
            For i = 2 To rsLoginInfo.RecordCount + 1
                aszTemp(i) = FormatDbValue(rsLoginInfo!login_event_id) & vbTab _
                        & FormatDbValue(rsLoginInfo!login_start_time) & vbTab _
                        & FormatDbValue(rsLoginInfo!user_id) & vbTab _
                        & FormatDbValue(rsLoginInfo!computer_name) & vbTab _
                        & FormatDbValue(rsLoginInfo!ip_address) & vbTab _
                        & FormatDbValue(rsLoginInfo!login_off_time) & vbTab _
                        & FormatDbValue(rsLoginInfo!login_off_type)
                rsLoginInfo.MoveNext
            Next i
        End If
    Else
        On Error GoTo ErrorHandle
        Set rsLoginInfo = g_oSysMan.GetOperateLogRs(g_aszUser(), dtStart, dtStart, tmStart, tmEnd, aszWS_FG)
        
        If rsLoginInfo.RecordCount > 0 Then
            ReDim aszTemp(1 To rsLoginInfo.RecordCount + 1)
            '标题
            aszTemp(1) = "操作事件代码" & vbTab & "用户代码" & vbTab & "功能组" & vbTab & "功能代码" & vbTab & "操作时间" & vbTab & "操作注释"
            '内容
            For i = 2 To rsLoginInfo.RecordCount + 1
                aszTemp(i) = FormatDbValue(rsLoginInfo!operation_event_id) & vbTab _
                        & FormatDbValue(rsLoginInfo!user_id) & vbTab _
                        & FormatDbValue(rsLoginInfo!function_group_id) & vbTab _
                        & FormatDbValue(rsLoginInfo!function_id) & vbTab _
                        & FormatDbValue(rsLoginInfo!operation_time) & vbTab _
                        & FormatDbValue(rsLoginInfo!Annotation) & vbTab
                rsLoginInfo.MoveNext
            Next i
        End If
    End If
    TransToStrs = aszTemp
Exit Function
ErrorHandle:
    ShowErrorMsg
End Function

