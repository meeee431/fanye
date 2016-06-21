VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmSel_DelOperateLog 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择操作日志"
   ClientHeight    =   2640
   ClientLeft      =   1785
   ClientTop       =   3015
   ClientWidth     =   6990
   HelpContextID   =   50000300
   Icon            =   "frmSel_DelOperateLog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   2160
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "帮助(&H)"
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
      MICON           =   "frmSel_DelOperateLog.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Height          =   375
      Left            =   4380
      TabIndex        =   9
      Top             =   2160
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "取消(&C)"
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
      MICON           =   "frmSel_DelOperateLog.frx":0166
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdOK 
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   2160
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "确定(&O)"
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
      MICON           =   "frmSel_DelOperateLog.frx":0182
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtLog 
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Top             =   1530
      Width           =   5205
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   315
      Left            =   4740
      TabIndex        =   2
      Top             =   450
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   63307779
      UpDown          =   -1  'True
      CurrentDate     =   36410
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   315
      Left            =   1500
      TabIndex        =   3
      Top             =   450
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   63307779
      UpDown          =   -1  'True
      CurrentDate     =   36410
   End
   Begin RTComctl3.TextButtonBox txtOperater 
      Height          =   315
      Left            =   1500
      TabIndex        =   5
      Top             =   960
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   556
      BackColor       =   16777215
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束时间(&E)"
      Height          =   180
      Left            =   3630
      TabIndex        =   7
      Top             =   510
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "操作员(&N):"
      Height          =   180
      Left            =   330
      TabIndex        =   6
      Top             =   1050
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "起始时间(&S):"
      Height          =   180
      Left            =   330
      TabIndex        =   4
      Top             =   510
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "模糊条件(&L):"
      Height          =   180
      Left            =   330
      TabIndex        =   0
      Top             =   1590
      Width           =   1080
   End
End
Attribute VB_Name = "frmSel_DelOperateLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *******************************************************************
' *  Source File Name  : frmSel_DelOperateLog                       *
' *  Project Name: PSTSMan                                          *
' *  Engineer:                                                      *
' *  Date Generated: 2002/08/19                                     *
' *  Last Revision Date : 2002/08/19                                *
' *  Brief Description   : 打开部分操作日志/删除部分操作日志        *
' *******************************************************************


Option Explicit
Option Base 1
Public m_bDelLog As Boolean
Const cGrayColor = &HC0C0C0

Dim aszOperate() As String
Dim aszFunAndGroup() As String
Dim dtStart As Date
Dim dtEnd As Date
Dim tmStart As Date
Dim tmEnd As Date
Dim bIsFun As Boolean
Dim szLog As String

Enum EText
    nIsUser = 0
    nIsFun = 1
    nIsFunGroup = 2
End Enum
Dim nText As EText





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
        Call frmStoreMenu.OpenDefOpeLog(aszOperate, dtStart, dtEnd, tmStart, tmEnd, aszFunAndGroup, bIsFun, szLog)
    Else
    End If
'    ReDim g_aszSelect(1)
    Unload Me
End Sub

Private Sub Form_Load()

    Me.Top = (Screen.Height - Me.ScaleHeight) / 2
    Me.Left = (Screen.Width - Me.ScaleWidth) / 2
    
    dtpBeginDate.Value = Date
    dtpEndDate.Value = Date
    
    If m_bDelLog Then
        Me.Caption = "删除操作日志"
        cmdOk.Caption = "执行"
        cmdCancel.Caption = "关闭"
    Else
        Me.Caption = "选择要显示的操作日志"
        cmdOk.Caption = "确定"
        cmdCancel.Caption = "取消"
    End If
    ClearTextBox Me
End Sub
Private Function GetString() As String
    Dim szTemp As String
    Dim nLen As Integer, i As Integer
    Dim aszTemp As Variant
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
    
    ReDim aszOperate(1)
    If txtOperater.Text <> "" Then
        aszOperate = GetIPString(txtOperater.Text)
    End If
    
    ReDim aszFunAndGroup(1)
    
    dtStart = dtpBeginDate.Value
    dtEnd = dtpEndDate.Value
    szLog = txtLog.Text
    
End Sub

Private Sub GetSelectResult()
    If frmSelect.m_bOk = True Then
        Select Case nText
            Case nIsUser
                txtOperater.Text = GetString
'            Case nIsFun
'                txtFun.Text = GetString
'            Case nIsFunGroup
'                txtFunGroup.Text = GetString
        End Select
    End If

End Sub
'
'Private Sub txtFun_Click()
'    frmSelect.m_szCaption = "选择功能"
'    frmSelect.Show vbModal
'    nText = nIsFun
'    GetSelectResult
'End Sub
'
'Private Sub txtFunGroup_Click()
''    frmSelect.m_bOk = False
'    frmSelect.m_szCaption = "选择功能组"
'    frmSelect.Show vbModal, Me
'    nText = nIsFunGroup
'    GetSelectResult
'
'End Sub

Private Sub txtOperater_Click()
'    g_bSelectOK = False
    frmSelect.m_szCaption = "选择操作人员"
    frmSelect.Show vbModal
    nText = nIsUser
    GetSelectResult
End Sub
