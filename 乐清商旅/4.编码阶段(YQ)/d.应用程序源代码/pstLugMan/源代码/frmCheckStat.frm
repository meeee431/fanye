VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmCheckStat 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   3375
   ClientTop       =   1935
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   7035
   Begin RTComctl3.TextButtonBox txtCondition 
      Height          =   285
      Left            =   4620
      TabIndex        =   7
      Top             =   1905
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -60
      TabIndex        =   12
      Top             =   690
      Width           =   7125
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   7005
      TabIndex        =   10
      Top             =   0
      Width           =   7005
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择查询条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   11
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.ComboBox cboSellStation 
      Height          =   300
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1935
      Width           =   1785
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   5370
      TabIndex        =   9
      Top             =   3375
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
      MICON           =   "frmCheckStat.frx":0000
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
      TabIndex        =   8
      Top             =   3375
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
      MICON           =   "frmCheckStat.frx":001C
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
      Left            =   1560
      TabIndex        =   1
      Top             =   1350
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   61669379
      UpDown          =   -1  'True
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   315
      Left            =   4620
      TabIndex        =   3
      Top             =   1350
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   61669379
      UpDown          =   -1  'True
      CurrentDate     =   36572
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   3120
      Left            =   -120
      TabIndex        =   13
      Top             =   3105
      Width           =   8745
   End
   Begin VB.Label lblCondition 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   180
      Left            =   3525
      TabIndex        =   6
      Top             =   1980
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结算日期(&B):"
      Height          =   180
      Left            =   450
      TabIndex        =   0
      Top             =   1410
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   3495
      TabIndex        =   2
      Top             =   1410
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&T):"
      Height          =   180
      Left            =   450
      TabIndex        =   4
      Top             =   1995
      Width           =   900
   End
End
Attribute VB_Name = "frmCheckStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IConditionForm

Public m_nStatType As ECheckStat
Public m_bOk As Boolean

Private m_oDataStat As New LugDataStatSvr

Private m_rsData As Recordset
Private m_vaCustomData As Variant
Private m_szFileName As String



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    '查询
    Dim rsTemp As Recordset
    Dim szSellstation As String
    On Error GoTo Error_Handle
    Select Case m_nStatType
    Case UI_SplitCompanyCheckStat
        Set rsTemp = m_oDataStat.SplitCompanyCheckStat(ResolveDisplay(cboSellStation.Text), dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(txtCondition.Text))
    Case UI_VehicleCheckStat
        Set rsTemp = m_oDataStat.VehicleCheckStat(ResolveDisplay(cboSellStation.Text), dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(txtCondition.Text))
    Case UI_RouteCheckStat
        Set rsTemp = m_oDataStat.RouteCheckStat(ResolveDisplay(cboSellStation.Text), dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(txtCondition.Text))
    End Select
    
    
    Set m_rsData = rsTemp
    
     
    ReDim m_vaCustomData(1 To 4, 1 To 2)
    
    m_vaCustomData(1, 1) = "统计开始日期"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY年MM月DD日")
    m_vaCustomData(2, 1) = "统计结束日期"
    m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")
    
    ResolveDisplay cboSellStation, szSellstation
    m_vaCustomData(3, 1) = "上车站"
    m_vaCustomData(3, 2) = szSellstation
    
    Select Case m_nStatType
    Case UI_SplitCompanyCheckStat
        m_vaCustomData(4, 1) = "拆帐公司"
    Case UI_VehicleCheckStat
        m_vaCustomData(4, 1) = "车辆"
    Case UI_RouteCheckStat
        m_vaCustomData(4, 1) = "线路"
    End Select
    
    m_vaCustomData(4, 2) = IIf(txtCondition.Text = "", "全部", txtCondition.Text)
    m_bOk = True
    Unload Me
    Exit Sub
Error_Handle:
    ShowErrorMsg
End Sub

Private Sub Form_Activate()
    mdiMain.SetPrintEnabled True
End Sub

Private Sub Form_Deactivate()
    mdiMain.SetPrintEnabled False
End Sub

Private Sub Form_Load()
    Dim dyNow As Date
    
    AlignFormPos Me
    m_oDataStat.Init m_oAUser
    
    FillSellStation cboSellStation
    
    dyNow = g_oParam.NowDate
    dtpBeginDate.Value = Format(DateAdd("m", -1, dyNow), "yyyy-mm-01")
    dtpEndDate.Value = DateAdd("d", -1, Format(dyNow, "yyyy-mm-01"))
    
    Select Case m_nStatType
    Case UI_SplitCompanyCheckStat
        lblCondition.Caption = "拆帐公司:"
        Me.Caption = "拆帐公司签发简报"
    Case UI_VehicleCheckStat
        lblCondition.Caption = "车辆:"
        Me.Caption = "车辆签发简报"
    Case UI_RouteCheckStat
        lblCondition.Caption = "线路:"
        Me.Caption = "线路签发简报"
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Sub txtCondition_Click()
    Dim oCommDialog As New CommDialog
    Dim aszTemp() As String
    
    oCommDialog.Init m_oAUser
    Select Case m_nStatType
    Case UI_SplitCompanyCheckStat
        aszTemp = oCommDialog.SelectCompany()
    Case UI_VehicleCheckStat
        aszTemp = oCommDialog.SelectVehicleEX()
    Case UI_RouteCheckStat
        aszTemp = oCommDialog.SelectRoute()
    End Select
    
    If ArrayLength(aszTemp) > 0 Then
        txtCondition.Text = MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))
    End If
End Sub


Private Property Get IConditionForm_CustomData() As Variant
    IConditionForm_CustomData = m_vaCustomData
End Property

Private Property Get IConditionForm_FileName() As String

    Select Case m_nStatType
    Case UI_SplitCompanyCheckStat
        m_szFileName = "拆帐公司签发简报.xls"
    Case UI_VehicleCheckStat
        m_szFileName = "车辆签发简报.xls"
    Case UI_RouteCheckStat
        m_szFileName = "线路签发简报.xls"
    End Select
    IConditionForm_FileName = m_szFileName
End Property

Private Property Get IConditionForm_RecordsetData() As Recordset
    Set IConditionForm_RecordsetData = m_rsData
End Property

'Private Sub FillSellStation()
'    '判断用户属于哪个上车站,如果为空则填充一个空行,再填充所有的上车站
'
'    '否则只填充用户所属的上车站
'End Sub

