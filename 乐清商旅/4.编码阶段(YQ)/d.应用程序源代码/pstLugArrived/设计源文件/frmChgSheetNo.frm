VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmChgSheetNo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "更改起始单号"
   ClientHeight    =   1800
   ClientLeft      =   5220
   ClientTop       =   5490
   ClientWidth     =   4950
   HelpContextID   =   7000001
   Icon            =   "frmChgSheetNo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   3660
      TabIndex        =   2
      Top             =   630
      Width           =   1125
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
      MICON           =   "frmChgSheetNo.frx":000C
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
      Left            =   3660
      TabIndex        =   1
      Top             =   180
      Width           =   1125
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
      MICON           =   "frmChgSheetNo.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   90
      Width           =   3345
      Begin VB.TextBox txtSheetID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1050
         TabIndex        =   0
         Top             =   420
         Width           =   2145
      End
      Begin VB.Line Line1 
         X1              =   90
         X2              =   3210
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "单据号(&S):"
         Height          =   180
         Left            =   90
         TabIndex        =   4
         Top             =   480
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmChgSheetNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public m_bOk As Boolean
Public m_bNoCancel As Boolean

Private Sub cmdCancel_Click()
    m_bOk = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo ErrHandle

    Dim oReg As New CFreeReg
    Dim szConnectName As String
    
    szConnectName = "Luggage"
    
    oReg.init "RTStation", HKEY_LOCAL_MACHINE, "Software\RTSoft"

    If Len(txtSheetID.Text) > g_oPackageParam.SheetIDNumberLen Then
        ShowMsg "签发单号长度应小于系统参数限定值,无法更改"
        Exit Sub
    End If
    g_szSheetID = FormatSheetID(txtSheetID.Text)
    oReg.SaveSetting szConnectName, "CurrentSheetID", g_szSheetID
    m_bOk = True
    Unload Me
    
    Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
    If KeyAscii = vbKeyF1 Then
        DisplayHelp Me
    End If
End Sub

Private Sub Form_Load()
'  AlignFormPos Me

    cmdCancel.Enabled = Not m_bNoCancel
    txtSheetID.Text = g_szSheetID
    txtSheetID.SelStart = Len(txtSheetID.Text)
    txtSheetID.MaxLength = g_oPackageParam.SheetIDNumberLen

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = m_bNoCancel And (Not m_bOk)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'  SaveFormPos Me
End Sub



Private Sub txtSheetID_Validate(Cancel As Boolean)
    If Not IsNumeric(txtSheetID.Text) Then

        MsgBox "签发单号必须为数字", vbExclamation, Me.Caption
        Cancel = True
    End If
End Sub

