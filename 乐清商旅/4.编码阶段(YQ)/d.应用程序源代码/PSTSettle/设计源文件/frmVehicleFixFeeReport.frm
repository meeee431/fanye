VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmVehicleFixFeeReport 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "车辆固定费用报表"
   ClientHeight    =   4380
   ClientLeft      =   4140
   ClientTop       =   3495
   ClientWidth     =   6660
   Icon            =   "frmVehicleFixFeeReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6660
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboKinds 
      Height          =   300
      Left            =   2085
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2859
      Width           =   3270
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   -15
      ScaleHeight     =   795
      ScaleWidth      =   7185
      TabIndex        =   1
      Top             =   0
      Width           =   7185
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请输入查询条件:"
         Height          =   180
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -75
      TabIndex        =   0
      Top             =   780
      Width           =   7215
   End
   Begin FText.asFlatTextBox txtVehicle 
      Height          =   315
      Left            =   2085
      TabIndex        =   3
      Top             =   1950
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
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
      ButtonPressedBackColor=   -2147483627
      Text            =   ""
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   315
      Left            =   2085
      TabIndex        =   4
      Top             =   1485
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      _Version        =   393216
      Format          =   61800448
      CurrentDate     =   37725
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   315
      Left            =   2085
      TabIndex        =   5
      Top             =   1035
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      _Version        =   393216
      Format          =   61800448
      CurrentDate     =   37725
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   4980
      TabIndex        =   6
      Top             =   3930
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "取消(&C)"
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
      MICON           =   "frmVehicleFixFeeReport.frx":000C
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
      Left            =   3495
      TabIndex        =   7
      Top             =   3930
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "确定(&E)"
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
      MICON           =   "frmVehicleFixFeeReport.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FText.asFlatTextBox txtCompany 
      Height          =   315
      Left            =   2085
      TabIndex        =   9
      Top             =   2400
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
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
      ButtonPressedBackColor=   -2147483627
      Text            =   ""
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   975
      Left            =   -60
      TabIndex        =   8
      Top             =   3645
      Width           =   6960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期:"
      Height          =   180
      Left            =   1125
      TabIndex        =   15
      Top             =   1095
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期:"
      Height          =   180
      Left            =   1125
      TabIndex        =   14
      Top             =   1560
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车辆:"
      Height          =   180
      Left            =   1125
      TabIndex        =   13
      Top             =   2010
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公司:"
      Height          =   180
      Left            =   1125
      TabIndex        =   12
      Top             =   2475
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "类别:"
      Height          =   180
      Left            =   1125
      TabIndex        =   11
      Top             =   2925
      Width           =   450
   End
End
Attribute VB_Name = "frmVehicleFixFeeReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_bOk As Boolean
Public m_dtStartDate As Date
Public m_dtEndDate As Date
Public m_szVehicleID As String
Public m_szCompanyID As String

Public m_szVehicleTagNo As String
Public m_szCompanyName As String

Public m_nStatus As Integer



Private Sub cmdCancel_Click()
    Unload Me
    m_bOk = False
End Sub

Private Sub cmdok_Click()
On Error GoTo ErrHandle
    m_szVehicleTagNo = txtVehicle.Text
    m_szCompanyName = txtCompany.Text
    
    If txtVehicle.Text = "" Then
        m_szVehicleID = ""
    End If
    If txtCompany.Text = "" Then
        m_szCompanyID = ""
    End If
    
    m_dtStartDate = dtpStartDate.Value
    m_dtEndDate = dtpEndDate.Value
    m_bOk = True
    m_nStatus = ResolveDisplay(cboKinds.Text)
    Unload Me
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Load()
    AlignFormPos Me
    dtpStartDate.Value = GetFirstMonthDay(Date)
    dtpEndDate.Value = GetLastMonthDay(Date)
    txtVehicle.Text = ""
    
    m_szVehicleID = ""
    m_szVehicleTagNo = ""
    m_szCompanyID = ""
    m_szCompanyName = ""
    
    FillKinds
    
End Sub

Private Sub FillKinds()
    cboKinds.AddItem MakeDisplayString(-1, GetFixFeeStatusName(-1))
    cboKinds.AddItem MakeDisplayString(0, GetFixFeeStatusName(0))
    cboKinds.AddItem MakeDisplayString(1, GetFixFeeStatusName(1))
    
    cboKinds.ListIndex = 1
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    Unload Me
End Sub




Private Sub txtCompany_ButtonClick()
    On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectCompany(True)
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    
    txtCompany.Text = ""
    m_szCompanyID = ""
    txtCompany.Text = TeamToString(aszTemp, 2, False)
    m_szCompanyID = TeamToString(aszTemp, 1)
'    txtCompany.Text = aszTemp(1, 1) & "[" & Trim(aszTemp(1, 2)) & "]"
    
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub



Private Sub txtVehicle_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    txtVehicle.Text = ""
    m_szVehicleID = ""
    aszTemp = oShell.SelectVehicleEX(True)
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
'    txtVehicle.Text = aszTemp(1, 1) & "[" & Trim(aszTemp(1, 2)) & "]"
    txtVehicle.Text = TeamToString(aszTemp, 2, False)
    m_szVehicleID = aszTemp(1, 1) 'TeamToString(aszTemp, 1)
    
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub



