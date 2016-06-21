VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmSellStationSimply 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "上车站各类款额类别核算简报"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   Icon            =   "frmSellStationdSimply.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6045
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox cboSellStation 
      Height          =   300
      Left            =   2145
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2625
      Width           =   2475
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -30
      TabIndex        =   1
      Top             =   690
      Width           =   6885
   End
   Begin VB.ComboBox cboExtraStatus 
      Height          =   300
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2115
      Width           =   2475
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   300
      Left            =   2160
      TabIndex        =   3
      Top             =   1545
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   529
      _Version        =   393216
      Format          =   19791872
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   300
      Left            =   2160
      TabIndex        =   4
      Top             =   990
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   529
      _Version        =   393216
      Format          =   19791872
      CurrentDate     =   36572
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   4200
      TabIndex        =   7
      Top             =   3360
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
      MICON           =   "frmSellStationdSimply.frx":000C
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
      Left            =   2880
      TabIndex        =   8
      Top             =   3360
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
      MICON           =   "frmSellStationdSimply.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   345
      Left            =   1560
      TabIndex        =   9
      Top             =   3360
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
      MICON           =   "frmSellStationdSimply.frx":0044
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
      Caption         =   " RTStation"
      Enabled         =   0   'False
      Height          =   3120
      Left            =   -120
      TabIndex        =   10
      Top             =   3120
      Width           =   8745
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   6615
      TabIndex        =   5
      Top             =   0
      Width           =   6615
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   6
         Top             =   270
         Width           =   990
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&T):"
      Height          =   180
      Left            =   960
      TabIndex        =   14
      Top             =   2685
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "状态(&S):"
      Height          =   180
      Left            =   960
      TabIndex        =   13
      Top             =   2145
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   960
      TabIndex        =   12
      Top             =   1605
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   960
      TabIndex        =   11
      Top             =   1050
      Width           =   1080
   End
End
Attribute VB_Name = "frmSellStationSimply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IConditionForm
Const cszFileName = "车次各上车站各类款额售票简报模板.xls"


Public m_bOk As Boolean
Public m_bBySaleTime As Boolean
Private m_rsData As Recordset
Private m_vaCustomData As Variant

Dim m_aszTemp() As String
Dim oDss As New TicketBusDim

Dim m_szCode As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error GoTo Error_Handle
    '生成记录集
    Dim rsTemp As Recordset

    Dim rsData As New Recordset
    Dim i As Integer
    If m_bBySaleTime Then
     
            Set rsTemp = oDss.GetBusSellStationSellInfoBySaleTimeSimply(m_szCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), , ResolveDisplay(cboSellStation))
    Else
    
            Set rsTemp = oDss.GetBusSellStationSellInfoByBusDateSimply(m_szCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), , ResolveDisplay(cboSellStation))
    End If
    Set m_rsData = rsTemp
     
    ReDim m_vaCustomData(1 To 5, 1 To 2)
    m_vaCustomData(1, 1) = "上车站"
    If cboSellStation <> "" Then
        m_vaCustomData(1, 2) = ResolveDisplayEx(cboSellStation)
    Else
        m_vaCustomData(1, 2) = ""
    End If
    m_vaCustomData(2, 1) = "统计开始日期"
    m_vaCustomData(2, 2) = Format(dtpBeginDate.Value, "YYYY年MM月DD日")
    m_vaCustomData(3, 1) = "统计结束日期"
    m_vaCustomData(3, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")
    m_vaCustomData(4, 1) = "补票状态"
    m_vaCustomData(4, 2) = cboExtraStatus.Text
    m_vaCustomData(5, 1) = "统计方式"
    m_vaCustomData(5, 2) = IIf(m_bBySaleTime, cszByOperationTime, cszByBusDate)
    
    m_bOk = True
    Unload Me
    Exit Sub
Error_Handle:
    ShowErrorMsg
End Sub


Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_Load()
    AlignFormPos Me
    oDss.Init m_oActiveUser
    m_szCode = ""
    m_bOk = False
    Dim dyNow As Date
    dyNow = m_oParam.NowDate
    dtpBeginDate.Value = Format(DateAdd("m", -1, dyNow), "yyyy-mm-01")
    dtpEndDate.Value = DateAdd("d", -1, Format(dyNow, "yyyy-mm-01"))
    cboExtraStatus.AddItem "1[售票]"
    cboExtraStatus.AddItem "2[补票]"
    cboExtraStatus.AddItem "3[售票+补票]"
    
    cboExtraStatus.ListIndex = 2

    FillSellStation cboSellStation

    
    If m_bBySaleTime Then
        Me.Caption = "款额类别核算[按售票时间汇总]"
'        lblCaption = "请输入售票的起止日期:"
    Else
        Me.Caption = "款额类别核算[按车次日期汇总]"
'        lblCaption = "请输入车次的起止日期:"
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
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




