VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmSellerInterSimpleCon 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "售票员互连售票简报"
   ClientHeight    =   4890
   ClientLeft      =   2970
   ClientTop       =   3255
   ClientWidth     =   7200
   HelpContextID   =   60000230
   Icon            =   "frmSellerInterSimpleCon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   345
      Left            =   2970
      TabIndex        =   15
      Top             =   4380
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
      MICON           =   "frmSellerInterSimpleCon.frx":000C
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
      Left            =   5820
      TabIndex        =   7
      Top             =   4380
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
      MICON           =   "frmSellerInterSimpleCon.frx":0028
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
      Left            =   4410
      TabIndex        =   6
      Top             =   4380
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
      MICON           =   "frmSellerInterSimpleCon.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -90
      TabIndex        =   13
      Top             =   690
      Width           =   7725
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   -30
      ScaleHeight     =   705
      ScaleWidth      =   7665
      TabIndex        =   11
      Top             =   0
      Width           =   7665
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择查询条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   12
         Top             =   240
         Width           =   1350
      End
   End
   Begin PSTTKAcc.AddDel adUnit 
      Height          =   2445
      Left            =   480
      TabIndex        =   10
      Top             =   1650
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   4313
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
   Begin VB.ComboBox cboSeller 
      Height          =   300
      Left            =   510
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1185
      Width           =   1635
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   315
      Left            =   5010
      TabIndex        =   5
      Top             =   1170
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   315
      Left            =   2790
      TabIndex        =   3
      Top             =   1170
      Width           =   1635
      _ExtentX        =   2884
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
      Left            =   -150
      TabIndex        =   14
      Top             =   4140
      Width           =   8745
   End
   Begin VB.Frame Frame1 
      Caption         =   "报表说明"
      Height          =   555
      Left            =   270
      TabIndex        =   8
      Top             =   5700
      Width           =   7635
      Begin VB.Label Label4 
         Caption         =   "统计售票员售出各车站的票的情况。"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   7455
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   2790
      TabIndex        =   2
      Top             =   930
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   5010
      TabIndex        =   4
      Top             =   930
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "售票员(&S):"
      Height          =   180
      Left            =   510
      TabIndex        =   0
      Top             =   930
      Width           =   900
   End
End
Attribute VB_Name = "frmSellerInterSimpleCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Implements IConditionForm

Const cszFileName = "售票员互连售票简报模板.xls"
Public m_bOk As Boolean
Private m_rsData As Recordset
Private m_vaCustomData As Variant

Private Sub adUnit_DataChange()
    EnableOK
End Sub

Private Sub cboSeller_Click()
    EnableOK
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error GoTo Error_Handle
    
    Dim nUnitCount As Integer, i As Integer
    Dim aszUnit() As String
    Dim oDss As New TicketSellerDim
    nUnitCount = ArrayLength(adUnit.RightData)
    If nUnitCount > 0 Then
        ReDim aszUnit(1 To nUnitCount)
        For i = 1 To nUnitCount
            aszUnit(i) = ResolveDisplay(adUnit.RightData(i))
        Next
        oDss.Init m_oActiveUser
        Set m_rsData = oDss.SellerUnitStat(ResolveDisplay(cboSeller.Text), aszUnit, dtpBeginDate.Value, dtpEndDate.Value)
    End If
    
    
    ReDim m_vaCustomData(1 To 4, 1 To 2)
    m_vaCustomData(1, 1) = "售票员"
    m_vaCustomData(1, 2) = cboSeller.Text
    
    m_vaCustomData(2, 1) = "统计开始日期"
    m_vaCustomData(2, 2) = Format(dtpBeginDate.Value, "YYYY年MM月DD日")
    m_vaCustomData(3, 1) = "统计结束日期"
    m_vaCustomData(3, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")
    
    m_vaCustomData(4, 1) = "制表人"
    m_vaCustomData(4, 2) = m_oActiveUser.UserID
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
    
    FillSeller
    FillUnit

    dtpBeginDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    dtpEndDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    
    EnableOK
End Sub

Private Sub FillSeller()
    Dim oSysMan As New SystemMan
    Dim auiUserInfo() As TUserInfo
    Dim i As Integer, nUserCount As Integer

    oSysMan.Init m_oActiveUser
    auiUserInfo = oSysMan.GetAllUser()
    nUserCount = ArrayLength(auiUserInfo)
    cboSeller.Clear
    For i = 1 To nUserCount
        cboSeller.AddItem MakeDisplayString(auiUserInfo(i).UserID, auiUserInfo(i).UserName)
    Next
    If cboSeller.ListCount > 0 Then cboSeller.ListIndex = 0
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
    cmdOk.Enabled = IIf(nSelUnitCount > 0 And cboSeller.Text <> "", True, False)
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
