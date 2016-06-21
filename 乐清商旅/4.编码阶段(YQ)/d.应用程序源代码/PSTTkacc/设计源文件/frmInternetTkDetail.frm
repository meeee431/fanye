VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmInternetTkDetail 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "网上售票明细"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox cboSellStation 
      Height          =   300
      Left            =   4680
      TabIndex        =   14
      Top             =   1320
      Width           =   1755
   End
   Begin VB.ComboBox cboStatus 
      Height          =   300
      Left            =   1590
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   1320
      Width           =   1755
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   60
      ScaleHeight     =   705
      ScaleWidth      =   7665
      TabIndex        =   5
      Top             =   0
      Width           =   7665
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择查询条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   6
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   0
      TabIndex        =   4
      Top             =   690
      Width           =   7725
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   2610
      TabIndex        =   0
      Top             =   4020
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
      MICON           =   "frmInternetTkDetail.frx":0000
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
      Left            =   3990
      TabIndex        =   1
      Top             =   4020
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
      MICON           =   "frmInternetTkDetail.frx":001C
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
      Left            =   5430
      TabIndex        =   2
      Top             =   4020
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
      MICON           =   "frmInternetTkDetail.frx":0038
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
      Height          =   315
      Left            =   4740
      TabIndex        =   7
      Top             =   930
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   315
      Left            =   1590
      TabIndex        =   8
      Top             =   930
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin PSTTKAcc.AddDel adUnit 
      Height          =   2175
      Left            =   360
      TabIndex        =   9
      Top             =   1590
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   3836
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
      Left            =   60
      TabIndex        =   3
      Top             =   3780
      Width           =   8745
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站:"
      Height          =   180
      Left            =   3600
      TabIndex        =   15
      Top             =   1320
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "状态(&S):"
      Height          =   180
      Left            =   390
      TabIndex        =   12
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   3540
      TabIndex        =   11
      Top             =   990
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   390
      TabIndex        =   10
      Top             =   990
      Width           =   1080
   End
End
Attribute VB_Name = "frmInternetTkDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Implements IConditionForm



Const cszFileName1 = "网上售票明细报表.xls"

Private m_rsData As Recordset
Public m_bOk As Boolean
Private m_vaCustomData As Variant
Public bSellCount As Boolean
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
    Dim nStatus As Integer
    Dim i As Integer, nUnitCount As Integer
    Dim oDss As New TicketUnitDim
    nUnitCount = ArrayLength(adUnit.RightData)
    If nUnitCount > 0 Then
        ReDim aszUnit(1 To nUnitCount)
        For i = 1 To nUnitCount
            aszUnit(i) = ResolveDisplay(adUnit.RightData(i))
        Next
    End If
    
    If cboStatus.Text = "已购" Then
        nStatus = 0
    ElseIf cboStatus.Text = "全部" Then
        nStatus = -1
    ElseIf cboStatus.Text = "已取票" Then
        nStatus = 1
    ElseIf cboStatus.Text = "已购+已取票" Then
        nStatus = 4
    ElseIf cboStatus.Text = "已取消" Then
        nStatus = 2
    ElseIf cboStatus.Text = "已废" Then
        nStatus = 3
    ElseIf cboStatus.Text = "已退" Then
        nStatus = 5
    End If
    oDss.Init m_oActiveUser
    Set m_rsData = oDss.InternetTkDetail(aszUnit, dtpBeginDate.Value, dtpEndDate.Value, nStatus, bSellCount, IIf(cboSellStation.Text <> "", ResolveDisplay(cboSellStation.Text), ""))
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
    
    cboStatus.Clear
    cboStatus.AddItem "全部"

    cboStatus.AddItem "已购"
    cboStatus.AddItem "已取票"
    cboStatus.AddItem "已购+已取票"
    cboStatus.AddItem "已取消"
    cboStatus.AddItem "已废"
    cboStatus.AddItem "已退"
    cboStatus.ListIndex = 2
    
    FillSellStation cboSellStation
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
    IConditionForm_FileName = cszFileName1
End Property

Private Property Get IConditionForm_RecordsetData() As Recordset
    Set IConditionForm_RecordsetData = m_rsData
End Property

Public Sub FillSellStation(cboSellStation As ComboBox)
    Dim oSystemMan As New SystemMan
    Dim atTemp() As TDepartmentInfo
    Dim i As Integer
    On Error GoTo Here
    '判断用户属于哪个上车站,如果为空则填充一个空行,再填充所有的上车站
    oSystemMan.Init m_oActiveUser
    atTemp = oSystemMan.GetAllSellStation(g_szUnitID)
    If m_oActiveUser.SellStationID = "" Then
        cboSellStation.AddItem ""
        For i = 1 To ArrayLength(atTemp)
            cboSellStation.AddItem MakeDisplayString(atTemp(i).szSellStationID, atTemp(i).szSellStationName)
        Next i
    '否则只填充用户所属的上车站
    Else
        For i = 1 To ArrayLength(atTemp)
            If m_oActiveUser.SellStationID = atTemp(i).szSellStationID Then
               cboSellStation.AddItem MakeDisplayString(atTemp(i).szSellStationID, atTemp(i).szSellStationName)
               Exit For
            End If
        Next i
        cboSellStation.ListIndex = 0
    End If
    Exit Sub
Here:
    ShowMsg err.Description
End Sub

