VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmTicketIssueScheme 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "车票发售计划"
   ClientHeight    =   3150
   ClientLeft      =   3735
   ClientTop       =   3165
   ClientWidth     =   5055
   HelpContextID   =   60000290
   Icon            =   "frmTicketIssueScheme.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   345
      Left            =   990
      TabIndex        =   10
      Top             =   2670
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "帮助"
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
      MICON           =   "frmTicketIssueScheme.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   6615
      TabIndex        =   7
      Top             =   0
      Width           =   6615
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择查询条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   8
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -60
      TabIndex        =   6
      Top             =   690
      Width           =   6885
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   3720
      TabIndex        =   1
      Top             =   2670
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "取消"
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
      MICON           =   "frmTicketIssueScheme.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   345
      Left            =   2370
      TabIndex        =   0
      Top             =   2670
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "确定"
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
      MICON           =   "frmTicketIssueScheme.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   315
      Left            =   2250
      TabIndex        =   2
      Top             =   1080
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   315
      Left            =   2250
      TabIndex        =   3
      Top             =   1680
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   3120
      Left            =   -120
      TabIndex        =   9
      Top             =   2430
      Width           =   8745
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   990
      TabIndex        =   5
      Top             =   1740
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   990
      TabIndex        =   4
      Top             =   1140
      Width           =   1080
   End
End
Attribute VB_Name = "frmTicketIssueScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IConditionForm
Const cszFileName = "车票发售计划.xls"
Public m_bOk As Boolean
Private m_rsData As Recordset
Private m_vaCustomData As Variant
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdok_Click()
    
    On Error GoTo Error_Handle
    '生成记录集

    Dim oTicketBusDim As New TicketUnitDim

    oTicketBusDim.Init m_oActiveUser
    Me.MousePointer = vbHourglass
    
    Set m_rsData = oTicketBusDim.BusSellScheme(dtpBeginDate.Value, dtpEndDate.Value)
    
    ReDim m_vaCustomData(1 To 2, 1 To 2)
        
        
    m_vaCustomData(1, 1) = "日期"
    m_vaCustomData(1, 2) = m_oParam.NowDate
        
  
    m_vaCustomData(2, 1) = "制表人"
    m_vaCustomData(2, 2) = m_oActiveUser.UserID
    
    m_bOk = True
    Me.MousePointer = vbDefault
    Unload Me
    Exit Sub
    
Error_Handle:
    Me.MousePointer = vbDefault
    ShowErrorMsg
End Sub

Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    m_bOk = False
    On Error GoTo Here
    

    dtpBeginDate.Value = m_oParam.NowDate
    dtpEndDate.Value = DateAdd("d", 2, m_oParam.NowDate)
      
    Exit Sub
Here:
    ShowMsg err.Description
End Sub

Private Property Get IConditionForm_CustomData() As Variant
    IConditionForm_CustomData = m_vaCustomData
End Property

Private Property Get IConditionForm_FileName() As String
    IConditionForm_FileName = cszFileName
End Property

Private Property Get IConditionForm_RecordsetData() As Recordset
    Set IConditionForm_RecordsetData = m_rsData
End Property
