VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmStationDayStat 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��վ�а�Ӫ���ձ�"
   ClientHeight    =   3570
   ClientLeft      =   3150
   ClientTop       =   2775
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6765
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   315
      Left            =   3960
      TabIndex        =   13
      Top             =   3000
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "ȷ��"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "frmStationDayStat.frx":0000
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
      TabIndex        =   14
      Top             =   3000
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "ȡ��"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "frmStationDayStat.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cboSellStation 
      Height          =   300
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1680
      Width           =   1755
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -360
      TabIndex        =   4
      Top             =   840
      Width           =   7125
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   -60
      ScaleHeight     =   825
      ScaleWidth      =   6825
      TabIndex        =   2
      Top             =   0
      Width           =   6825
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѡ���ѯ����:"
         Height          =   180
         Left            =   480
         TabIndex        =   3
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.ComboBox cboAcceptType 
      Height          =   300
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2160
      Width           =   1755
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   960
      Left            =   -120
      TabIndex        =   0
      Top             =   2640
      Width           =   6945
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   300
      Left            =   4560
      TabIndex        =   5
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
      TabIndex        =   6
      Top             =   1200
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   529
      _Version        =   393216
      Format          =   61669376
      CurrentDate     =   36572
   End
   Begin FText.asFlatTextBox txtEndStation 
      Height          =   300
      Left            =   4560
      TabIndex        =   15
      Top             =   1680
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      ButtonVisible   =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "�� Ʊ վ(&T)"
      Height          =   180
      Left            =   420
      TabIndex        =   12
      Top             =   1740
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "��ʼ����(&S)"
      Height          =   180
      Left            =   420
      TabIndex        =   10
      Top             =   1260
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "��������(&E)"
      Height          =   180
      Left            =   3480
      TabIndex        =   9
      Top             =   1260
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "�� �� վ(&T)"
      Height          =   180
      Left            =   3480
      TabIndex        =   8
      Top             =   1740
      Width           =   990
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "���˷�ʽ(&A)"
      Height          =   180
      Left            =   420
      TabIndex        =   7
      Top             =   2220
      Width           =   990
   End
End
Attribute VB_Name = "frmStationDayStat"
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
    m_szStation = ResolveDisplay(txtEndStation.Text)
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

Private Sub txtEndStation_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init m_oAUser
    aszTemp = oShell.SelectStation()
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtEndStation.Text = aszTemp(1, 1) & "[" & Trim(aszTemp(1, 2)) & "]"
'    cmdQuery.Enabled = True
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

