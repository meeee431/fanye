VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmSelectVehicleCompany 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "车辆营收简报"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6660
   StartUpPosition =   2  '屏幕中心
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   5250
      TabIndex        =   16
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
      MICON           =   "frmSelectVehicleCompany.frx":0000
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
      Left            =   3810
      TabIndex        =   17
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
      MICON           =   "frmSelectVehicleCompany.frx":001C
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
      Height          =   1470
      Left            =   -120
      TabIndex        =   9
      Top             =   3630
      Width           =   8745
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   6615
      TabIndex        =   7
      Top             =   -30
      Width           =   6615
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   8
         Top             =   270
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   0
      TabIndex        =   6
      Top             =   690
      Width           =   6885
   End
   Begin VB.ComboBox cboExtraStatus 
      Height          =   300
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1530
      Width           =   1905
   End
   Begin FText.asFlatTextBox txtSellStation 
      Height          =   300
      Left            =   1290
      TabIndex        =   1
      ToolTipText     =   "请点...进行选择"
      Top             =   2010
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonHotBackColor=   -2147483633
      Locked          =   -1  'True
      ButtonPressedBackColor=   -2147483627
      Text            =   ""
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
   End
   Begin FText.asFlatTextBox txtVehicle 
      Height          =   300
      Left            =   1290
      TabIndex        =   11
      ToolTipText     =   "请点...进行选择"
      Top             =   2460
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonHotBackColor=   -2147483633
      Locked          =   -1  'True
      ButtonPressedBackColor=   -2147483627
      Text            =   ""
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   300
      Left            =   4680
      TabIndex        =   13
      Top             =   1080
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   300
      Left            =   1260
      TabIndex        =   14
      Top             =   1050
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin FText.asFlatTextBox txtTransportCompanyID 
      Height          =   300
      Left            =   4680
      TabIndex        =   15
      Top             =   1560
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
   Begin VB.Label lblCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参运公司(T):"
      Height          =   180
      Left            =   3390
      TabIndex        =   12
      Top             =   1620
      Width           =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车辆(&V):"
      Height          =   180
      Left            =   210
      TabIndex        =   10
      Top             =   2520
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   210
      TabIndex        =   5
      Top             =   1110
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   3450
      TabIndex        =   4
      Top             =   1110
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "补票状态(&S):"
      Height          =   180
      Left            =   210
      TabIndex        =   3
      Top             =   1590
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&T):"
      Height          =   180
      Left            =   210
      TabIndex        =   2
      Top             =   2070
      Width           =   900
   End
End
Attribute VB_Name = "frmSelectVehicleCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IConditionForm


Private m_rsData As Recordset
Public m_bOk As Boolean
Private m_vaCustomData As Variant
Dim oDss As New TicketVehicleDim

Dim m_szCode As String          '公司代码
Dim m_szSellStationID As String '上车站代码
Dim m_szVehicleID As String     '车辆代码
Dim m_szFileName As String

Private Sub cmdok_Click()
On Error GoTo Error_Handle
    '生成记录集
    Dim rsTemp As Recordset
    
    If MDIMain.m_szMethod = False Then    '车辆营收统计[按售票统计]
       Set rsTemp = oDss.GetVehicleStat(m_szCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), m_szVehicleID, m_szSellStationID, False)
    Else                                  '车辆营收统计[按检票统计]
       Set rsTemp = oDss.GetVehicleStat(m_szCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), m_szVehicleID, m_szSellStationID, True)
    End If

    ReDim m_vaCustomData(1 To 7, 1 To 2)
    
    m_vaCustomData(1, 1) = "统计开始日期"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY年MM月DD日")
    m_vaCustomData(2, 1) = "统计结束日期"
    m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")
    m_vaCustomData(3, 1) = "补票状态"
    m_vaCustomData(3, 2) = ResolveDisplayEx(cboExtraStatus.Text)
    m_vaCustomData(4, 1) = "统计方式"

    If MDIMain.m_szMethod = False Then
        m_vaCustomData(4, 2) = cszBySale
        m_szFileName = "车辆售票营收简报模板.xls"
    Else
        m_vaCustomData(4, 2) = cszByCheck
        m_szFileName = "车辆售票营收简报模板按检票.xls"
    End If
    
    m_vaCustomData(5, 1) = "制表人"
    m_vaCustomData(5, 2) = m_oActiveUser.UserName
    
    m_vaCustomData(6, 1) = "上车站"
    m_vaCustomData(6, 2) = txtSellStation.Text
    
    m_vaCustomData(7, 1) = "拆帐公司"
    m_vaCustomData(7, 2) = IIf((txtTransportCompanyID.Text <> ""), txtTransportCompanyID.Text, "全部公司")
    
    Set m_rsData = rsTemp
    
    m_bOk = True
    Unload Me
    Exit Sub
Error_Handle:
    ShowErrorMsg
End Sub

Private Sub Form_Load()
    oDss.Init m_oActiveUser
    
    m_szCode = ""
    m_szSellStationID = ""
    m_szVehicleID = ""
    
    m_bOk = False

'    dtpBeginDate.Value = DateAdd("d", -1, m_oParam.NowDate)
'    dtpEndDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    '设置为上个月的一号到31
    Dim dyNow As Date
    dyNow = m_oParam.NowDate
    dtpBeginDate.Value = Format(DateAdd("m", -1, dyNow), "yyyy-mm-01")
    dtpEndDate.Value = DateAdd("d", -1, Format(dyNow, "yyyy-mm-01"))
    cboExtraStatus.AddItem "1[售票]"
    cboExtraStatus.AddItem "2[补票]"
    cboExtraStatus.AddItem "3[售票+补票]"
    
    cboExtraStatus.ListIndex = 2
    
    If Trim(m_oActiveUser.SellStationID) <> "" Then
        txtSellStation.Enabled = False
        m_szSellStationID = "'" & m_oActiveUser.SellStationID & "'"
    End If
    
    
End Sub

Private Property Get IConditionForm_CustomData() As Variant
    IConditionForm_CustomData = m_vaCustomData
End Property

Private Property Get IConditionForm_FileName() As String
    IConditionForm_FileName = m_szFileName
End Property

Private Property Get IConditionForm_RecordsetData() As Recordset
    Set IConditionForm_RecordsetData = m_rsData
End Property

Private Sub txtSellStation_ButtonClick()
    Dim aszTemp() As String
    Dim nCount As Integer
    On Error GoTo ErrorHandle
    aszTemp = m_oShell.SelectSellStation(m_oActiveUser.SellStationID, , True)
    txtSellStation.Text = TeamToString(aszTemp, 2, False)
    
    m_szSellStationID = TeamToString(aszTemp, 1, True)
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub txtTransportCompanyID_ButtonClick()
    Dim aszTransportCompanyID() As String
    aszTransportCompanyID = m_oShell.SelectCompany
    txtTransportCompanyID.Text = TeamToString(aszTransportCompanyID, 2)
    
    m_szCode = TeamToString(aszTransportCompanyID, 1)
End Sub

Private Sub txtVehicle_ButtonClick()
    Dim aszTemp() As String
    aszTemp = m_oShell.SelectVehicle(, , , , , True)
    txtVehicle.Text = TeamToString(aszTemp, 2)
    m_szVehicleID = TeamToString(aszTemp, 1)
End Sub

