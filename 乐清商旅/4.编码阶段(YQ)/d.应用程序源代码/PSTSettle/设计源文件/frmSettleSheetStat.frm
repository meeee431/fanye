VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmSettleSheetStat 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "路单结算明细表"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5805
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -30
      TabIndex        =   9
      Top             =   630
      Width           =   10755
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   705
      Left            =   -60
      TabIndex        =   7
      Top             =   1920
      Width           =   6285
      Begin RTComctl3.CoolButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   345
         Left            =   4320
         TabIndex        =   5
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
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
         MICON           =   "frmSettleSheetStat.frx":0000
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
         Left            =   3180
         TabIndex        =   4
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "确定(&E)"
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
         MICON           =   "frmSettleSheetStat.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      TabIndex        =   6
      Top             =   690
      Width           =   5625
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Left            =   3870
         TabIndex        =   3
         Top             =   480
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61800448
         CurrentDate     =   37725
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Left            =   1140
         TabIndex        =   1
         Top             =   480
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61800448
         CurrentDate     =   37725
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期:"
         Height          =   180
         Left            =   2940
         TabIndex        =   2
         Top             =   540
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期:"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   540
         Width           =   810
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "路单结算明细:"
      Height          =   180
      Left            =   330
      TabIndex        =   8
      Top             =   270
      Width           =   1170
   End
End
Attribute VB_Name = "frmSettleSheetStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_bOk As Boolean
Public m_dtStartDate As Date
Public m_dtEndDate As Date

Private Sub cmdCancel_Click()
    Unload Me
    m_bOk = False
End Sub

Private Sub cmdok_Click()
On Error GoTo ErrHandle
    m_dtStartDate = dtpStartDate.Value
    m_dtEndDate = dtpEndDate.Value
    m_bOk = True
    Unload Me
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Load()
    AlignFormPos Me
    dtpStartDate.Value = GetFirstMonthDay(Date)
    dtpEndDate.Value = GetLastMonthDay(Date)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    Unload Me
End Sub
