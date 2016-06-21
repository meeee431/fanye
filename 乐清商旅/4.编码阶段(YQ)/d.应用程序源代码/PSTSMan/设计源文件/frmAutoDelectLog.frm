VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A0123751-4698-48C1-A06C-A2482B5ED508}#2.0#0"; "RTComctl2.ocx"
Begin VB.Form frmAutoDelectLog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "日志整理"
   ClientHeight    =   4950
   ClientLeft      =   1020
   ClientTop       =   2205
   ClientWidth     =   7755
   HelpContextID   =   5000801
   Icon            =   "frmAutoDelectLog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存设置(&S)"
      Height          =   315
      Left            =   6225
      TabIndex        =   18
      Top             =   585
      Width           =   1400
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "日志导出(&E)"
      Height          =   315
      Left            =   6225
      TabIndex        =   21
      Top             =   1755
      Width           =   1400
   End
   Begin VB.CommandButton cmdProject 
      Caption         =   "计划整理(&P)"
      Height          =   315
      Left            =   6225
      TabIndex        =   20
      Top             =   1365
      Width           =   1400
   End
   Begin VB.CommandButton cmdDo 
      Caption         =   "立即整理(&I)"
      Height          =   315
      Left            =   6225
      TabIndex        =   19
      Top             =   975
      Width           =   1400
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "关闭(&C)"
      Height          =   315
      Left            =   6225
      TabIndex        =   17
      Top             =   195
      Width           =   1400
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   315
      HelpContextID   =   5000801
      Left            =   6225
      TabIndex        =   22
      Top             =   2145
      Width           =   1400
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "日志自动整理设置:"
      Height          =   4740
      Left            =   90
      TabIndex        =   23
      Top             =   105
      Width           =   6000
      Begin VB.CommandButton cmdView 
         Caption         =   "查看(&V)"
         Height          =   315
         Left            =   4635
         TabIndex        =   16
         Top             =   4290
         Width           =   1095
      End
      Begin VB.CheckBox chkOpeExport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "自动整理前导出操作日志(&B)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1470
         TabIndex        =   9
         Top             =   2235
         Width           =   2760
      End
      Begin VB.CheckBox chkAutoDelOpeLog 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "自动整理操作日志(&O)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   975
         TabIndex        =   5
         Top             =   1665
         Value           =   1  'Checked
         Width           =   2205
      End
      Begin VB.CheckBox chkLoginExport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "自动整理前导出登录日志(&A)"
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   1470
         TabIndex        =   4
         Top             =   975
         Width           =   2910
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   4215
         TabIndex        =   3
         Top             =   660
         Width           =   270
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         BuddyControl    =   "txtReserveTime1"
         BuddyDispid     =   196620
         OrigLeft        =   2790
         OrigTop         =   330
         OrigRight       =   3060
         OrigBottom      =   645
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtReserveTime1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3840
         TabIndex        =   2
         Text            =   "7"
         Top             =   660
         Width           =   390
      End
      Begin VB.CheckBox chkAutoDelLoginLog 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "自动整理登录日志(&L)"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   990
         TabIndex        =   0
         Top             =   465
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   300
         Left            =   4200
         TabIndex        =   8
         Top             =   1830
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtReseverTime2"
         BuddyDispid     =   196622
         OrigLeft        =   2790
         OrigTop         =   330
         OrigRight       =   3060
         OrigBottom      =   645
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtReseverTime2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3825
         TabIndex        =   7
         Text            =   "7"
         Top             =   1830
         Width           =   390
      End
      Begin RTComctl2.FolderBrowser txtDir 
         Height          =   300
         Left            =   2445
         TabIndex        =   11
         Top             =   2730
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   529
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
      Begin RTComctl2.FileBrowser txtRunProgram 
         Height          =   300
         Left            =   2775
         TabIndex        =   13
         Top             =   3150
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Filter          =   "整理运行程序(*.exe) | *.exe"
      End
      Begin RTComctl2.FileBrowser txtExpFile 
         Height          =   300
         Left            =   2445
         TabIndex        =   15
         Top             =   3870
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Filter          =   "日志导出文件(*.log) | *.log|所有文件(*.*)|*.*"
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
         X1              =   1455
         X2              =   5800
         Y1              =   3690
         Y2              =   3690
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   6
         X1              =   1455
         X2              =   5800
         Y1              =   3675
         Y2              =   3675
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "导出日志查看"
         Height          =   180
         Left            =   180
         TabIndex        =   27
         Top             =   3585
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "日志导出文件(&W):"
         Height          =   180
         Left            =   945
         TabIndex        =   14
         Top             =   3930
         Width           =   1440
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "操作日志"
         Height          =   210
         Left            =   195
         TabIndex        =   26
         Top             =   1380
         Width           =   795
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   1020
         X2              =   5800
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   1020
         X2              =   5800
         Y1              =   1425
         Y2              =   1425
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "登录日志"
         Height          =   210
         Left            =   195
         TabIndex        =   25
         Top             =   240
         Width           =   795
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   1020
         X2              =   5800
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   1020
         X2              =   5800
         Y1              =   285
         Y2              =   285
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "其他设置"
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   2505
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   1005
         X2              =   5800
         Y1              =   2565
         Y2              =   2565
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "日志整理运行程序(&F):"
         Height          =   180
         Left            =   945
         TabIndex        =   12
         Top             =   3180
         Width           =   1800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   1005
         X2              =   5800
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   150
         Picture         =   "frmAutoDelectLog.frx":014A
         Top             =   555
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "选择导出目录(&R):"
         Height          =   180
         Left            =   960
         TabIndex        =   10
         Top             =   2820
         Width           =   1440
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "自动整理后保留日志天数(&T):"
         Height          =   180
         Left            =   1455
         TabIndex        =   6
         Top             =   1965
         Width           =   2340
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "自动整理后保留日志天数(&D):"
         Height          =   180
         Left            =   1455
         TabIndex        =   1
         Top             =   780
         Width           =   2340
      End
   End
End
Attribute VB_Name = "frmAutoDelectLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ******************************************************************
' *  Source File Name  : frmAutoDelectLog                           *
' *  Project Name: PSTSMan                                          *
' *  Engineer:                                                      *
' *  Date Generated: 2002/08/19                                     *
' *  Last Revision Date : 2002/08/19                                *
' *  Brief Description   : 设定系统自动删除日志的条件               *
' *******************************************************************


Option Explicit

Const PKeyParent = "SysMan"
Public szLoginLogLastCleanTime As String
Public szOperLogLastCleanTime As String
Public szOperLogLastCleanedTime As String
Public szLoginLogLastCleanedTime  As String
Public szDir As String
Public szRunProgram As String

Dim szEnable1 As String
Dim szReserveDay1 As String
Dim szIsExport1 As String
Dim szEnable2 As String
Dim szReserveDay2 As String
Dim szIsExport2 As String
Dim szLastView As String
Dim oReg As New CFreeReg


Private Sub chkAutoDelLoginLog_Click()
    AutoDelLoginLog
End Sub

Private Sub chkAutoDelOpeLog_Click()
    AutoDelOpeLog
End Sub


Private Sub cmdClose_Click()
    '立即整理结果设置

    Unload Me
End Sub

Private Sub cmdDo_Click()
    Dim i As Integer
    Dim bTempChg As Boolean
    Dim szUser As String
    
    
    bTempChg = False
    i = IIf(szEnable1 = "True", vbChecked, vbUnchecked)
    bTempChg = IIf(chkAutoDelLoginLog.Value <> i, True, bTempChg)
    i = IIf(szIsExport1 = "True", vbChecked, vbUnchecked)
    bTempChg = IIf(chkLoginExport.Value <> i, True, bTempChg)
    bTempChg = IIf(txtReserveTime1.Text <> szReserveDay1, True, bTempChg)
    i = IIf(szEnable2 = "True", vbChecked, vbUnchecked)
    bTempChg = IIf(chkAutoDelOpeLog.Value <> i, True, bTempChg)
    i = IIf(szIsExport2 = "True", vbChecked, vbUnchecked)
    bTempChg = IIf(chkOpeExport.Value <> i, True, bTempChg)
    bTempChg = IIf(txtReseverTime2.Text <> szReserveDay2, True, bTempChg)
    bTempChg = IIf(txtDir.Text <> szDir, True, bTempChg)
    bTempChg = IIf(txtRunProgram.Text <> szRunProgram, True, bTempChg)
    
    If bTempChg = True Then
        i = MsgBox("设置参数已改动,是否确认改变?", vbQuestion + vbYesNo, cszMsg)
        If i = vbYes Then
            cmdSave_Click
        Else
            DisplayInfo
        End If
    End If

    '调用PSTLogDel.exe
    Dim szAllRunProgram As String
    szUser = g_oActUser.UserID
    szAllRunProgram = szRunProgram & " " & szUser & "," & g_szPassword & ","
    On Error GoTo ErrorHandle
    Shell szAllRunProgram, vbNormalFocus
Exit Sub
ErrorHandle:
    MsgBox "找不到日志整理运行程序,请重设!", vbExclamation, cszMsg
End Sub

Private Sub cmdExport_Click()
    frmCleanLog.Show vbModal, Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me, content
End Sub

Private Sub cmdProject_Click()
    Dim i As Integer
    Dim bTempChg As Boolean
    bTempChg = False
    i = IIf(szEnable1 = "True", vbChecked, vbUnchecked)
    bTempChg = IIf(chkAutoDelLoginLog.Value <> i, True, bTempChg)
    i = IIf(szIsExport1 = "True", vbChecked, vbUnchecked)
    bTempChg = IIf(chkLoginExport.Value <> i, True, bTempChg)
    bTempChg = IIf(txtReserveTime1.Text <> szReserveDay1, True, bTempChg)
    i = IIf(szEnable2 = "True", vbChecked, vbUnchecked)
    bTempChg = IIf(chkAutoDelOpeLog.Value <> i, True, bTempChg)
    i = IIf(szIsExport2 = "True", vbChecked, vbUnchecked)
    bTempChg = IIf(chkOpeExport.Value <> i, True, bTempChg)
    bTempChg = IIf(txtReseverTime2.Text <> szReserveDay2, True, bTempChg)
    bTempChg = IIf(txtDir.Text <> szDir, True, bTempChg)
    bTempChg = IIf(txtRunProgram.Text <> szRunProgram, True, bTempChg)
    
    If bTempChg = True Then
        i = MsgBox("设置参数已改动,是否确认改变?", vbQuestion + vbYesNo, cszMsg)
        If i = vbYes Then
            cmdSave_Click
        Else
            DisplayInfo
        End If
    End If


    frmTaskProject.Show vbModal, Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorHandle
    GetInfoFromUI
    SetValue
    MsgBox "保存成功!", vbInformation, cszMsg
Exit Sub
ErrorHandle:
    MsgBox err.Number & err.Description
End Sub

Private Sub cmdView_Click()
    '察看日志
    Dim szExecuteApp As String
    Dim tempFile As String
    tempFile = Trim(txtExpFile.Text)
    On Error GoTo ErrorHandle
    If Not FileIsExist(tempFile) Then
        MsgBox "文件" & tempFile & "不存在!", vbOKOnly + vbInformation
        Exit Sub
    End If
    szExecuteApp = "NOTEPAD.EXE " & tempFile
    Shell szExecuteApp, vbNormalFocus
    szLastView = tempFile
    Call oReg.SaveSetting(PKeyParent, "LastViewLogFile", tempFile)
    
'    Init Dir
    On Error GoTo there
    Dim szTemp As String
    Dim i As Integer, j As Integer
    i = 1
    szTemp = tempFile
    Do While i > 0
        i = InStr(1, szTemp, "\")
        j = Len(szTemp)
        szTemp = Right(szTemp, j - i)
    Loop
    szTemp = Left(tempFile, Len(tempFile) - j)
    txtExpFile.InitDir = szTemp
 
    
Exit Sub
ErrorHandle:
    MsgBox err.Number & err.Description, vbExclamation, cszMsg
Exit Sub
there:
End Sub

Private Sub Form_Load()
    
    Me.Top = (Screen.Height - Me.ScaleHeight) / 2
    Me.Left = (Screen.Width - Me.ScaleWidth) / 2
    
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    
    szEnable1 = oReg.GetSetting(PKeyParent, "IsAutoDelLoginLog", "True")
    szReserveDay1 = oReg.GetSetting(PKeyParent, "LoginLogSaveDays", "15")
    szIsExport1 = oReg.GetSetting(PKeyParent, "IsLoginLogExport", "True")
    szDir = oReg.GetSetting(PKeyParent, "LogExportDir", "C:\My Documents")
    szEnable2 = oReg.GetSetting(PKeyParent, "IsAutoDelOperLog", "True")
    szReserveDay2 = oReg.GetSetting(PKeyParent, "OperLogSaveDays", "7")
    szIsExport2 = oReg.GetSetting(PKeyParent, "IsOperLogExport", "True")
    szRunProgram = oReg.GetSetting(PKeyParent, "LogDelProgramName", "")
    szLastView = oReg.GetSetting(PKeyParent, "LastViewLogFile", "")
    
    szLoginLogLastCleanTime = oReg.GetSetting(PKeyParent, "LoginLogLastCleanTime", "")
    szOperLogLastCleanTime = oReg.GetSetting(PKeyParent, "OperLogLastCleanTime", "")
    szLoginLogLastCleanedTime = oReg.GetSetting(PKeyParent, "LoginLogLastCleanedTime", "")
    szOperLogLastCleanedTime = oReg.GetSetting(PKeyParent, "OperLogLastCleanedTime", "")

    DisplayInfo
    
    
End Sub



Private Sub AutoDelLoginLog()
    If chkAutoDelLoginLog.Value Then
        txtReserveTime1.Enabled = True
        txtReserveTime1.BackColor = vbWhite
        chkLoginExport.Enabled = True
    Else
        txtReserveTime1.Enabled = False
        txtReserveTime1.BackColor = cGreyColor
        chkLoginExport.Enabled = False
    End If
    
End Sub

Private Sub AutoDelOpeLog()
    If chkAutoDelOpeLog.Value Then
        txtReseverTime2.Enabled = True
        txtReseverTime2.BackColor = vbWhite
        chkOpeExport.Enabled = True
    Else
        txtReseverTime2.Enabled = False
        txtReseverTime2.BackColor = cGreyColor
        chkOpeExport.Enabled = False
    End If

End Sub

Private Sub DisplayInfo()
    If szEnable1 = "True" Then
        chkAutoDelLoginLog.Value = vbChecked
    Else
        chkAutoDelLoginLog.Value = Unchecked
    End If
    
    If szEnable2 = "True" Then
        chkAutoDelOpeLog.Value = vbChecked
    Else
        chkAutoDelOpeLog.Value = Unchecked
    End If

    If szIsExport1 = "True" Then
        chkLoginExport.Value = vbChecked
    Else
        chkLoginExport.Value = Unchecked
    End If
    
    If szIsExport2 = "True" Then
        chkOpeExport.Value = vbChecked
    Else
        chkOpeExport.Value = Unchecked
    End If
    
    txtDir.Text = szDir
    txtRunProgram.Text = szRunProgram
    
    Dim tempFile As String
    txtRunProgram.Text = Trim(txtRunProgram.Text)
    tempFile = txtRunProgram.Text
    If Not FileIsExist(tempFile) Then
        MsgBox "文件 " & tempFile & " 不存在!" & vbCrLf & "日志整理程序指向错误, 请重设.", vbOKOnly + vbInformation, cszMsg
        txtRunProgram.Text = szRunProgram
    End If
    
    
    txtReserveTime1.Text = szReserveDay1
    txtReseverTime2.Text = szReserveDay2
    
    txtExpFile.Text = szLastView

    AutoDelLoginLog
    AutoDelOpeLog

'    Call InitDir(txtExpFile.Name, szLastView)
'    InitDir txtRunProgram.Name, szRunProgram
    'txtExpFile.InitDir
    On Error GoTo there
    Dim szTemp As String
    Dim i As Integer, j As Integer
    i = 1
    szTemp = szLastView
    Do While i > 0
        i = InStr(1, szTemp, "\")
        j = Len(szTemp)
        szTemp = Right(szTemp, j - i)
    Loop
    szTemp = Left(szLastView, Len(szLastView) - j)
    txtExpFile.InitDir = szTemp
    
    'txtRunProgram.InitDir
    i = 1
    szTemp = szRunProgram
    Do While i > 0
        i = InStr(1, szTemp, "\")
        j = Len(szTemp)
        szTemp = Right(szTemp, j - i)
    Loop
    szTemp = Left(szRunProgram, Len(szRunProgram) - j)
    txtRunProgram.InitDir = szTemp

there:
End Sub

Private Sub GetInfoFromUI()
    If chkAutoDelLoginLog.Value = vbChecked Then
        szEnable1 = "True"
    Else
        szEnable1 = "False"
    End If
    
    If chkAutoDelOpeLog.Value = vbChecked Then
        szEnable2 = "True"
    Else
        szEnable2 = "False"
    End If
    
    If chkLoginExport.Value = vbChecked Then
        szIsExport1 = "True"
    Else
        szIsExport1 = "False"
    End If
    
    If chkOpeExport.Value = vbChecked Then
        szIsExport2 = "True"
    Else
        szIsExport2 = "False"
    End If
    
    szReserveDay1 = txtReserveTime1.Text
    szReserveDay2 = txtReseverTime2.Text
    szDir = txtDir.Text
    szRunProgram = txtRunProgram.Text
End Sub

Private Sub SetValue()
    Call oReg.SaveSetting(PKeyParent, "IsAutoDelLoginLog", szEnable1)
    Call oReg.SaveSetting(PKeyParent, "LoginLogSaveDays", szReserveDay1)
    Call oReg.SaveSetting(PKeyParent, "IsLoginLogExport", szIsExport1)
    Call oReg.SaveSetting(PKeyParent, "LogExportDir", szDir)
    Call oReg.SaveSetting(PKeyParent, "IsAutoDelOperLog", szEnable2)
    Call oReg.SaveSetting(PKeyParent, "OperLogSaveDays", szReserveDay2)
    Call oReg.SaveSetting(PKeyParent, "IsOperLogExport", szIsExport2)
    Call oReg.SaveSetting(PKeyParent, "LogDelProgramName", szRunProgram)
    Call oReg.SaveSetting(PKeyParent, "LastViewLogFile", szLastView)

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call oReg.SaveSetting(PKeyParent, "LoginLogLastCleanTime", szLoginLogLastCleanTime)
    Call oReg.SaveSetting(PKeyParent, "OperLogLastCleanTime", szOperLogLastCleanTime)
    Call oReg.SaveSetting(PKeyParent, "LoginLogLastCleanedTime", szLoginLogLastCleanedTime)
    Call oReg.SaveSetting(PKeyParent, "OperLogLastCleanedTime", szOperLogLastCleanedTime)
    
End Sub

Private Sub txtDir_LostFocus()
    Dim tempdir As String
    txtDir.Text = Trim(txtDir.Text)
    tempdir = txtDir.Text
    If Not Right(tempdir, 1) = "\" Then
        txtDir.Text = tempdir + "\"
        tempdir = tempdir + "\"
    End If
    If Not FileIsExist(tempdir) Then
        MsgBox "目录" & tempdir & "不存在!", vbOKOnly + vbInformation, cszMsg
        txtDir.Text = szDir
    End If

End Sub

Private Sub txtReserveTime1_KeyPress(KeyAscii As Integer)
    KeyAscii = NumberText(KeyAscii, txtReserveTime1.Text, txtReserveTime1.Seltext, txtReserveTime1.SelStart, False, False)
End Sub

Private Sub txtReserveTime1_Validate(Cancel As Boolean)
    If txtReserveTime1.Text = "" Then txtReserveTime1.Text = "15"
End Sub

Private Sub txtReseverTime2_KeyPress(KeyAscii As Integer)
    KeyAscii = NumberText(KeyAscii, txtReseverTime2.Text, txtReseverTime2.Seltext, txtReseverTime2.SelStart, False, False)
End Sub

Private Sub txtReseverTime2_Validate(Cancel As Boolean)
    If txtReseverTime2.Text = "" Then txtReseverTime2.Text = "7"
End Sub

Private Sub txtRunProgram_LostFocus()
    Dim tempFile As String
    txtRunProgram.Text = Trim(txtRunProgram.Text)
    tempFile = txtRunProgram.Text
    If Not FileIsExist(tempFile) Then
        MsgBox "文件 " & tempFile & " 不存在!" & vbCrLf & "请重设日志整理运行程序!", vbOKOnly + vbInformation, cszMsg
        txtRunProgram.Text = szRunProgram
    End If
    
End Sub

