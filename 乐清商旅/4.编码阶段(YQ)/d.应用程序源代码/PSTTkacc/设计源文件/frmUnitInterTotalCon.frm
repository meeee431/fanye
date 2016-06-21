VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmUnitInterTotalCon 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "车站互连售票统计报表"
   ClientHeight    =   4365
   ClientLeft      =   3105
   ClientTop       =   1470
   ClientWidth     =   6705
   HelpContextID   =   6000601
   Icon            =   "frmUnitInterTotalCon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   2520
      TabIndex        =   11
      Top             =   3930
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
      MICON           =   "frmUnitInterTotalCon.frx":000C
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
      Left            =   3930
      TabIndex        =   4
      Top             =   3930
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
      MICON           =   "frmUnitInterTotalCon.frx":0028
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
      Left            =   5370
      TabIndex        =   3
      Top             =   3930
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
      MICON           =   "frmUnitInterTotalCon.frx":0044
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
      ScaleWidth      =   7665
      TabIndex        =   1
      Top             =   0
      Width           =   7665
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -60
      TabIndex        =   0
      Top             =   690
      Width           =   7725
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   315
      Left            =   4680
      TabIndex        =   5
      Top             =   900
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   315
      Left            =   1530
      TabIndex        =   6
      Top             =   900
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin PSTTKAcc.AddDel adUnit 
      Height          =   2415
      Left            =   300
      TabIndex        =   7
      Top             =   1260
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   4260
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonWidth     =   1215
      ButtonHeight    =   315
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   3120
      Left            =   -120
      TabIndex        =   8
      Top             =   3690
      Width           =   8745
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   3480
      TabIndex        =   10
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   330
      TabIndex        =   9
      Top             =   960
      Width           =   1080
   End
End
Attribute VB_Name = "frmUnitInterTotalCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Implements IConditionForm
Const cszFileName = "车站互连售票统计报表模板.xls"

Public m_bOk As Boolean
Private m_rsData As Recordset
Private m_vaCustomData As Variant

Private Sub adUnit_DataChange()
    EnableOK
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error GoTo Error_Handle
    '生成记录集
    Dim aszUnit() As String
    Dim i As Integer, nUnitCount As Integer
    Dim oDss As New TicketUnitDim
    nUnitCount = ArrayLength(adUnit.RightData)
    If nUnitCount > 0 Then
        ReDim aszUnit(1 To nUnitCount)
        For i = 1 To nUnitCount
            aszUnit(i) = ResolveDisplay(adUnit.RightData(i))
        Next
    End If
    
    oDss.Init m_oActiveUser
    Set m_rsData = oDss.DateUnitStat(aszUnit, dtpBeginDate.Value, dtpEndDate.Value)
    ReDim m_vaCustomData(1 To 3, 1 To 2)
    m_vaCustomData(1, 1) = "统计开始日期"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY年MM月DD日")
    
    m_vaCustomData(2, 1) = "统计结束日期"
    m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")
    
    m_vaCustomData(3, 1) = "制表人"
    m_vaCustomData(3, 2) = m_oActiveUser.UserID
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

    dtpBeginDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    dtpEndDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    
    
    FillUnit
    EnableOK
End Sub
Private Sub FillUnit()
    Dim oSysMan As New SystemMan
    Dim auniUnitInfo() As TUnit
    Dim i As Integer, nUnitCount As Integer
    Dim aszTemp() As String
    oSysMan.Init m_oActiveUser
    auniUnitInfo = oSysMan.GetAllUnit()
    nUnitCount = ArrayLength(auniUnitInfo)
    If nUnitCount > 0 Then
        ReDim aszTemp(1 To nUnitCount)
        For i = 1 To nUnitCount
            aszTemp(i) = MakeDisplayString(auniUnitInfo(i).szUnitID, auniUnitInfo(i).szUnitShortName)
        Next
    End If
    adUnit.LeftData = aszTemp
End Sub

Private Sub EnableOK()
    Dim nSelUnitCount As Integer
    nSelUnitCount = ArrayLength(adUnit.RightData)
    cmdOk.Enabled = IIf(nSelUnitCount > 0, True, False)
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
