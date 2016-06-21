VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmSel_DelLoginLog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择登录日志"
   ClientHeight    =   4425
   ClientLeft      =   285
   ClientTop       =   495
   ClientWidth     =   7290
   HelpContextID   =   50000220
   Icon            =   "frmSel_DelLoginLog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   315
      Left            =   6045
      TabIndex        =   17
      Top             =   4005
      Width           =   1095
   End
   Begin VB.Frame fraTime 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   21
      Top             =   2055
      Width           =   7035
      Begin VB.CheckBox chkTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "(&T)指定操作时间"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   -15
         Width           =   1650
      End
      Begin VB.CheckBox chkEndTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "指定终止时间(&B):"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3465
         TabIndex        =   10
         Top             =   420
         Value           =   1  'Checked
         Width           =   1740
      End
      Begin MSComCtl2.DTPicker dtpEndTime 
         Height          =   315
         Left            =   5250
         TabIndex        =   11
         Top             =   360
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   64487426
         CurrentDate     =   36410.9999884259
      End
      Begin MSComCtl2.DTPicker dtpBeginTime 
         Height          =   315
         Left            =   1665
         TabIndex        =   9
         Top             =   360
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   64487426
         CurrentDate     =   36410.0000115741
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起始操作时间(&A):"
         Height          =   180
         Left            =   180
         TabIndex        =   22
         Top             =   435
         Width           =   1440
      End
   End
   Begin VB.Frame fraWorkStation 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   20
      Top             =   3015
      Width           =   7035
      Begin VB.TextBox txtWorkStation 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1665
         TabIndex        =   14
         Top             =   360
         Width           =   5175
      End
      Begin VB.CheckBox chkWorkStation 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "(&S)指定登录的工作站"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   15
         Width           =   2025
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "登录工作站(&H):"
         Height          =   180
         Left            =   180
         TabIndex        =   13
         Top             =   420
         Width           =   1260
      End
   End
   Begin VB.Frame fraOperater 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   135
      Width           =   7035
      Begin RTComctl3.TextButtonBox txtOperater 
         Height          =   315
         Left            =   1665
         TabIndex        =   2
         Top             =   345
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   556
         BackColor       =   -2147483637
         Enabled         =   0   'False
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
      Begin VB.CheckBox chkOperater 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "(&U)指定操作员"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   0
         Top             =   0
         Width           =   1485
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "操作员(&N):"
         Height          =   180
         Left            =   180
         TabIndex        =   1
         Top             =   412
         Width           =   900
      End
   End
   Begin VB.Frame fraDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   18
      Top             =   1095
      Width           =   7035
      Begin VB.CheckBox chkDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "(&D)指定操作日期"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   0
         Value           =   1  'Checked
         Width           =   1650
      End
      Begin VB.CheckBox chkEndDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "终止操作日期(&T):"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3465
         TabIndex        =   6
         Top             =   420
         Value           =   1  'Checked
         Width           =   1785
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Left            =   5250
         TabIndex        =   7
         Top             =   360
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Format          =   64487424
         CurrentDate     =   36410
      End
      Begin MSComCtl2.DTPicker dtpBeginDate 
         Height          =   315
         Left            =   1665
         TabIndex        =   5
         Top             =   360
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Format          =   64487424
         CurrentDate     =   36410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起始操作日期(&F):"
         Height          =   180
         Left            =   180
         TabIndex        =   4
         Top             =   420
         Width           =   1440
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   315
      Left            =   4815
      TabIndex        =   16
      Top             =   4005
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   315
      Left            =   3600
      TabIndex        =   15
      Top             =   4005
      Width           =   1095
   End
End
Attribute VB_Name = "frmSel_DelLoginLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *******************************************************************
' *  Source File Name  : frmSel_DelLoginLog                         *
' *  Project Name: PSTSMan                                          *
' *  Engineer:                                                      *
' *  Date Generated: 2002/08/19                                     *
' *  Last Revision Date : 2002/08/19                                *
' *  Brief Description   : 打开部分登录日志/删除部分登录日志        *
' *******************************************************************


Option Explicit
Option Base 1
Public m_bDelLog As Boolean
Const cGrayColor = &HC0C0C0

Dim g_aszUser() As String
Dim aszWorkStation() As String
Dim dtStart As Date
Dim dtEnd As Date
Dim tmStart As Date
Dim tmEnd As Date



Private Sub chkDate_Click()
    If chkDate.Value = Unchecked Then
        dtpBeginDate.Enabled = False
        dtpEndDate.Enabled = False
        chkEndDate.Enabled = False
    Else
        dtpBeginDate.Enabled = True
        chkEndDate.Enabled = True
        chkEndDate_Click
    End If
End Sub

Private Sub chkEndDate_Click()
    If chkEndDate.Value = vbChecked Then
        dtpEndDate.Enabled = True
    Else
        dtpEndDate.Enabled = False
    End If
End Sub

Private Sub chkEndTime_Click()
    If chkEndTime.Value = vbChecked Then
        dtpEndTime.Enabled = True
    Else
        dtpEndTime.Enabled = False
    End If
End Sub

Private Sub chkOperater_Click()
    If chkOperater.Value = vbChecked Then
        txtOperater.Enabled = True
        txtOperater.BackColor = vbWhite
    Else
        txtOperater.Enabled = False
        txtOperater.BackColor = cGrayColor
    End If
End Sub

Private Sub chkTime_Click()
    If chkTime.Value = Unchecked Then
        dtpBeginTime.Enabled = False
        dtpEndTime.Enabled = False
        chkEndTime.Enabled = False
    Else
        dtpBeginTime.Enabled = True
        chkEndTime.Enabled = True
        chkEndTime_Click
    End If
End Sub

Private Sub chkWorkStation_Click()
    If chkWorkStation.Value = Unchecked Then
        txtWorkStation.Enabled = False
        txtWorkStation.BackColor = cGrayColor
    Else
        txtWorkStation.Enabled = True
        txtWorkStation.BackColor = vbWhite
    End If
End Sub

Private Sub cmdCancel_Click()
    
'    ReDim g_aszSelect(1)
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdOK_Click()
    GetInfoFromUI
    If m_bDelLog = False Then
        Call frmStoreMenu.OpenDefLoginLog(g_aszUser, dtStart, dtEnd, tmStart, tmEnd, aszWorkStation)
    Else
    End If
'    ReDim g_aszSelect(1)
    Unload Me
End Sub

Private Sub cmdSelOperater_Click()
    
    frmSelect.m_szCaption = "选择工作人员"
    frmSelect.Show vbModal, Me
End Sub


Private Sub Form_Activate()
    If frmSelect.m_bOk Then
        txtOperater.Text = GetString
    End If

End Sub

Private Sub Form_Load()

    Me.Top = (Screen.Height - Me.ScaleHeight) / 2
    Me.Left = (Screen.Width - Me.ScaleWidth) / 2
    
    dtpBeginDate.Value = Date
    dtpEndDate.Value = Date
    
    frmSelect.m_bOk = False

    If m_bDelLog Then
        Me.Caption = "自定义删除登录日志"
        cmdOk.Caption = "执行"
        cmdCancel.Caption = "关闭"
    Else
        Me.Caption = "选择要显示的登录日志"
        cmdOk.Caption = "确定"
        cmdCancel.Caption = "取消"
    End If
    ClearTextBox Me
End Sub


Private Function GetString() As String
    Dim szTemp As String
    Dim nLen As Integer, i As Integer
    Dim aszTemp() As Variant
    szTemp = ""
    aszTemp = frmSelect.m_aszSelect
    nLen = ArrayLength(aszTemp)
    If nLen > 0 Then
        If aszTemp(1) <> "" Then
            For i = 1 To nLen
                If i = 1 Then
                    szTemp = szTemp & aszTemp(i)
                Else
                    szTemp = szTemp & "," & aszTemp(i)
                End If
            Next i
        End If
    End If
    GetString = szTemp
End Function

Private Sub GetInfoFromUI()
    If chkOperater.Value = Unchecked Then
        ReDim g_aszUser(1)
    Else
        g_aszUser = GetIPString(txtOperater.Text)
    End If
    
    If chkDate.Value = Unchecked Then
        dtStart = cszEmptyDateStr
        dtEnd = cszEmptyDateStr
    Else
        dtStart = dtpBeginDate.Value
        If chkEndDate.Value = vbChecked Then
            dtEnd = dtpEndDate.Value
        Else
            dtEnd = cszEmptyDateStr
        End If
    End If
    
    If chkTime.Value = Unchecked Then
        tmStart = "00:00:00"
        tmEnd = "00:00:00"
    Else
        tmStart = dtpBeginTime.Value
        If chkEndTime.Value = vbChecked Then
            tmEnd = dtpEndTime.Value
        Else
            tmEnd = "00:00:00"
        End If
    End If
    
    If chkWorkStation.Value = Unchecked Then
        ReDim aszWorkStation(1)
    Else
        aszWorkStation = GetIPString(txtWorkStation.Text)
    End If
    
End Sub


Private Sub txtOperater_Click()
    frmSelect.m_szCaption = "选择工作人员"
    frmSelect.Show vbModal
End Sub
