VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmCopyEvnBus 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "复制环境车次"
   ClientHeight    =   2100
   ClientLeft      =   2955
   ClientTop       =   3525
   ClientWidth     =   5685
   Icon            =   "frmCopyEvnBus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   5685
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtOldBusID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2055
      TabIndex        =   0
      Top             =   300
      Width           =   1605
   End
   Begin VB.TextBox txtBusID 
      Height          =   285
      Left            =   2055
      TabIndex        =   1
      Top             =   720
      Width           =   1605
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   285
      Left            =   2055
      TabIndex        =   2
      Top             =   1155
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   503
      _Version        =   393216
      Format          =   68091904
      CurrentDate     =   38677
   End
   Begin RTComctl3.CoolButton cmdOk 
      Height          =   345
      Left            =   4065
      TabIndex        =   6
      Top             =   315
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "确定(&S)"
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
      MICON           =   "frmCopyEvnBus.frx":000C
      PICN            =   "frmCopyEvnBus.frx":0028
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
      Height          =   345
      Left            =   4065
      TabIndex        =   7
      Top             =   810
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "关闭(&C)"
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
      MICON           =   "frmCopyEvnBus.frx":03C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   285
      Left            =   2055
      TabIndex        =   8
      Top             =   1590
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   503
      _Version        =   393216
      Format          =   68091904
      CurrentDate     =   38677
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "目标结束日期(&D):"
      Height          =   180
      Left            =   540
      TabIndex        =   9
      Top             =   1635
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "原车次代码(&B):"
      Height          =   180
      Left            =   525
      TabIndex        =   5
      Top             =   330
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "目标车次代码(&B):"
      Height          =   180
      Left            =   525
      TabIndex        =   4
      Top             =   795
      Width           =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "目标起始日期(&D):"
      Height          =   180
      Left            =   525
      TabIndex        =   3
      Top             =   1260
      Width           =   1440
   End
End
Attribute VB_Name = "frmCopyEvnBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oReBus As New REBus
Private moScheme As New REScheme

Public m_szOldBusID As String
Public m_szNewBusID As String
Public m_dtStartDate As Date
Public m_dtEndDate As Date
Public m_bOk As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    m_bOk = True
    m_szOldBusID = txtOldBusID.Text
    m_szNewBusID = txtBusID.Text
    m_dtStartDate = dtpStartDate.Value
    m_dtEndDate = dtpEndDate.Value
    Unload Me
    
End Sub

Private Sub Form_Load()
    m_bOk = False
    moScheme.Init g_oActiveUser
    m_oReBus.Init g_oActiveUser
    dtpStartDate.Value = Date
    dtpEndDate.Value = Date
    If frmEnvBus.lvBus.SelectedItem Is Nothing Then
        MsgBox "请选择车次!", vbExclamation, "提示"
        Exit Sub
    End If
    txtOldBusID.Text = frmEnvBus.lvBus.SelectedItem.Text
    txtBusID.Text = txtOldBusID.Text
    
End Sub

Private Sub txtBusID_GotFocus()
    txtBusID.SelStart = 0
    txtBusID.SelLength = 100
End Sub
