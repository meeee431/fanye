VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmSellStationDayStat 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "售票站行包营收日报"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6690
   StartUpPosition =   3  '窗口缺省
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   315
      Left            =   3960
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmSellStationDayStat.frx":0000
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
      Cancel          =   -1  'True
      Height          =   315
      Left            =   5280
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmSellStationDayStat.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   960
      Left            =   -120
      TabIndex        =   7
      Top             =   2280
      Width           =   6945
   End
   Begin VB.ComboBox cboAcceptType 
      Height          =   300
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1680
      Width           =   1755
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   -60
      ScaleHeight     =   825
      ScaleWidth      =   6825
      TabIndex        =   4
      Top             =   0
      Width           =   6825
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择查询条件:"
         Height          =   180
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -360
      TabIndex        =   3
      Top             =   840
      Width           =   7125
   End
   Begin VB.ComboBox cboSellStation 
      Height          =   300
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1680
      Width           =   1755
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   300
      Left            =   4560
      TabIndex        =   8
      Top             =   1200
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   529
      _Version        =   393216
      Format          =   61669376
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   300
      Left            =   1500
      TabIndex        =   9
      Top             =   1200
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   529
      _Version        =   393216
      Format          =   61669376
      CurrentDate     =   36572
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "托运方式(&A)"
      Height          =   180
      Left            =   3480
      TabIndex        =   13
      Top             =   1740
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "结束日期(&E)"
      Height          =   180
      Left            =   3480
      TabIndex        =   12
      Top             =   1260
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "开始日期(&S)"
      Height          =   180
      Left            =   420
      TabIndex        =   11
      Top             =   1260
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "售 票 站(&T)"
      Height          =   180
      Left            =   420
      TabIndex        =   10
      Top             =   1740
      Width           =   990
   End
End
Attribute VB_Name = "frmSellStationDayStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Implements IConditionForm

Public m_bOk As Boolean

Public m_dtWorkDate As Date
Public m_dtEndDate As Date
Public m_szStation As String
Public m_szAcceptType As String
Public m_SellStation As String

Private Sub cmdOk_Click()
    m_dtWorkDate = dtpBeginDate.Value
    m_dtEndDate = dtpEndDate.Value
'    m_szStation = ResolveDisplay(txtEndStation.Text)
    m_szAcceptType = cboAcceptType.Text
    m_SellStation = cboSellStation.Text
    m_bOk = True
    Unload Me
End Sub

Private Sub Form_Load()
   m_bOk = False
   
  AlignFormPos Me
  dtpBeginDate.Value = DateAdd("d", -1, g_oParam.NowDate)
  dtpEndDate.Value = Format(dtpBeginDate.Value, "yyyy-mm-dd") & " 23:59:59"
    
    
    FillSellStation cboSellStation
    FillAcceptType

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
'Private Property Get IConditionForm_FileName() As String
'    IConditionForm_FileName = cszFileName
'End Property

Private Sub FillAcceptType()
With cboAcceptType
   .AddItem ""
   .AddItem GetLuggageTypeString(0)
   .AddItem GetLuggageTypeString(1)
   .ListIndex = 0
End With
End Sub





Private Sub Form_Unload(Cancel As Integer)
   SaveFormPos Me
End Sub
