VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmUnitSellSimple 
   BackColor       =   &H00E0E0E0&
   Caption         =   "车站售票营收简报"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   5235
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtSellStationID 
      Height          =   300
      Left            =   2325
      TabIndex        =   13
      Top             =   1440
      Width           =   1785
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3840
      TabIndex        =   5
      Top             =   3300
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
      MICON           =   "frmUnitSellSimple.frx":0000
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
      Height          =   315
      Left            =   2490
      TabIndex        =   4
      Top             =   3300
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
      MICON           =   "frmUnitSellSimple.frx":001C
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
      Height          =   3120
      Left            =   -150
      TabIndex        =   9
      Top             =   3060
      Width           =   8745
   End
   Begin VB.ComboBox cboSellStation 
      Height          =   300
      Left            =   2325
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1020
      Width           =   1785
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -90
      TabIndex        =   3
      Top             =   690
      Width           =   6885
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   -30
      ScaleHeight     =   705
      ScaleWidth      =   6615
      TabIndex        =   1
      Top             =   0
      Width           =   6615
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择查询条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   2
         Top             =   240
         Width           =   1350
      End
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   1170
      TabIndex        =   0
      Top             =   3150
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
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
      MICON           =   "frmUnitSellSimple.frx":0038
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
      Height          =   300
      Left            =   2325
      TabIndex        =   7
      Top             =   2400
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   62652419
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   300
      Left            =   2325
      TabIndex        =   8
      Top             =   1860
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   62652419
      CurrentDate     =   36572
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&T):"
      Height          =   180
      Left            =   1050
      TabIndex        =   12
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   1050
      TabIndex        =   11
      Top             =   1920
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   1050
      TabIndex        =   10
      Top             =   2460
      Width           =   1080
   End
End
Attribute VB_Name = "frmUnitSellSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IConditionForm

Const cszFileName = "车站售票营收简报模板.xls"

Private m_rsData As Recordset
Public m_bOk As Boolean
Private m_vaCustomData As Variant

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error GoTo Error_Handle
    Dim aszSellStation(1 To 1) As String
    Dim oDss As New TicketUnitDim
    oDss.Init m_oActiveUser
    aszSellStation(1) = ResolveDisplay(cboSellStation.Text)
    If txtSellStationID.Text <> "" Then
        aszSellStation(1) = txtSellStationID.Text
    End If
    Set m_rsData = oDss.UnitSellSimple(aszSellStation, dtpBeginDate.Value, dtpEndDate.Value)
    ReDim m_vaCustomData(1 To 4, 1 To 2)

    m_vaCustomData(1, 1) = "统计日期"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "yyyy-MM-dd HH:mm:ss") & "——" & Format(dtpEndDate.Value, "yyyy-MM-dd HH:mm:ss")
    
    m_vaCustomData(2, 1) = "制表人"
    m_vaCustomData(2, 2) = m_oActiveUser.UserID
    
    m_vaCustomData(3, 1) = "制表日期"
    m_vaCustomData(3, 2) = Format(Now, "yyyy-MM-dd HH:mm:ss")
    
    m_vaCustomData(4, 1) = "统计单位"
    m_vaCustomData(4, 2) = ResolveDisplayEx(cboSellStation.Text)
    
    m_bOk = True
    Unload Me
    Exit Sub
Error_Handle:
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

    dtpBeginDate.Value = m_oParam.NowDate
    dtpEndDate.Value = Format(dtpBeginDate.Value, "yyyy-mm-dd") & " 23:59:59"
    
    FillSellStation cboSellStation
    
    If Trim(m_oActiveUser.SellStationID) <> "" Then
        txtSellStationID.Enabled = False
    End If
End Sub

'Private Sub FillUnit()
'    Dim oSysMan As New SystemMan
'    Dim auniUnitInfo() As TUnit
'    Dim i As Integer, nUnitCount As Integer
'    oSysMan.Init m_oActiveUser
'    auniUnitInfo = oSysMan.GetAllUnit()
'    nUnitCount = ArrayLength(auniUnitInfo)
'    cboUnit.Clear
'    cboUnit.AddItem ""
'    For i = 1 To nUnitCount
'        cboUnit.AddItem MakeDisplayString(auniUnitInfo(i).szUnitID, auniUnitInfo(i).szUnitShortName)
'    Next
'    If cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
'End Sub

Private Property Get IConditionForm_CustomData() As Variant
    IConditionForm_CustomData = m_vaCustomData
End Property

Private Property Get IConditionForm_FileName() As String
    IConditionForm_FileName = cszFileName
End Property

Private Property Get IConditionForm_RecordsetData() As Recordset
    Set IConditionForm_RecordsetData = m_rsData
End Property

