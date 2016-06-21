VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmRoute 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "线路--线路信息"
   ClientHeight    =   4380
   ClientLeft      =   3390
   ClientTop       =   3075
   ClientWidth     =   6675
   HelpContextID   =   10000390
   Icon            =   "frmRoute.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6675
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtRoute 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1755
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1110
      Width           =   4095
   End
   Begin RTComctl3.CoolButton cmdSection 
      Height          =   315
      Left            =   2520
      TabIndex        =   5
      Top             =   3900
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "路段管理(&L)"
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
      MICON           =   "frmRoute.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   7185
      TabIndex        =   19
      Top             =   0
      Width           =   7185
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请修改或新增线路信息:"
         Height          =   180
         Left            =   180
         TabIndex        =   20
         Top             =   300
         Width           =   1890
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -60
      TabIndex        =   18
      Top             =   780
      Width           =   7215
   End
   Begin RTComctl3.CoolButton lblStartStation 
      Height          =   270
      Left            =   1320
      TabIndex        =   12
      Top             =   3180
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   476
      BTYPE           =   8
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   14737632
      MPTR            =   1
      MICON           =   "frmRoute.frx":0166
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtRouteName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1755
      TabIndex        =   3
      Top             =   1515
      Width           =   4095
   End
   Begin RTComctl3.CoolButton lblEndStation 
      Height          =   270
      Left            =   3330
      TabIndex        =   14
      Top             =   3180
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   476
      BTYPE           =   8
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRoute.frx":0182
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   315
      Left            =   240
      TabIndex        =   15
      Top             =   3900
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "帮助(&H)"
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
      MICON           =   "frmRoute.frx":019E
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
      Height          =   315
      Left            =   3810
      TabIndex        =   4
      Top             =   3900
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "保存(&S)"
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
      MICON           =   "frmRoute.frx":01BA
      PICN            =   "frmRoute.frx":01D6
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
      Left            =   5130
      TabIndex        =   6
      Top             =   3900
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "关闭(&C)"
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
      MICON           =   "frmRoute.frx":0570
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
      Height          =   1560
      Left            =   -150
      TabIndex        =   21
      Top             =   3600
      Width           =   8745
   End
   Begin FText.asFlatTextBox txtFormula 
      Height          =   315
      Left            =   1755
      TabIndex        =   8
      Top             =   1920
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   556
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483631
      ButtonHotBackColor=   -2147483633
      ButtonPressedBackColor=   -2147483627
      Text            =   ""
      ButtonBackColor =   -2147483633
   End
   Begin FText.asFlatMemo txtAnnotation 
      Height          =   720
      Left            =   1755
      TabIndex        =   10
      Top             =   2340
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1270
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonHotForeColor=   -2147483628
      ButtonHotBackColor=   -2147483632
   End
   Begin VB.Label lblMileage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0公里"
      Height          =   180
      Left            =   5310
      TabIndex        =   17
      Top             =   3240
      Width           =   450
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "里程数:"
      Height          =   180
      Left            =   4620
      TabIndex        =   16
      Top             =   3240
      Width           =   630
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "计算公式(&D):"
      Height          =   180
      Left            =   630
      TabIndex        =   7
      Top             =   1980
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "备注(&R):"
      Height          =   180
      Left            =   630
      TabIndex        =   9
      Top             =   2415
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "终点站:"
      Height          =   180
      Left            =   2685
      TabIndex        =   13
      Top             =   3240
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "起点站:"
      Height          =   180
      Left            =   630
      TabIndex        =   11
      Top             =   3240
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "线路名称(&N):"
      Height          =   180
      Left            =   630
      TabIndex        =   2
      Top             =   1575
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "线路代码(&C):"
      Height          =   180
      Left            =   630
      TabIndex        =   0
      Top             =   1170
      Width           =   1080
   End
End
Attribute VB_Name = "frmRoute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************
'* Source File Name:frmRoute.frm
'* Project Name:RTBusMan
'* Engineer:
'* Data Generated:2002/08/27
'* Last Revision Date:2002/08/31
'* Brief Description:
'* Relational Document:UI_BS_SM_31.DOC
'**********************************************************
Public Status As EFormStatus
Public m_szRouteID As String
Public m_bIsParent As Boolean
Private m_oRoute As New BackRoute

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdok_Click()
    On Error GoTo ErrorHandle
    Select Case Status
    Case AddStatus
        m_oRoute.AddNew
        m_oRoute.Annotation = txtAnnotation.Text
        m_oRoute.RouteID = txtRoute.Text
        m_oRoute.RouteName = txtRouteName.Text
'        m_oRoute.TicketPriceFormula = txtFormula.Text
        m_oRoute.Update
        m_szRouteID = txtRoute.Text
        cmdSection.Enabled = True
        cmdOk.Caption = "保存(&S)"
        If m_bIsParent Then
            frmAllRoute.AddList m_oRoute.RouteID
        End If
        Status = ModifyStatus
    Case ModifyStatus
        m_oRoute.Identify txtRoute.Text
        m_oRoute.Annotation = txtAnnotation.Text
        m_oRoute.RouteName = txtRouteName.Text
'        m_oRoute.TicketPriceFormula = txtFormula.Text
        m_oRoute.Update
        If m_bIsParent Then
            frmAllRoute.UpdateList txtRoute.Text
        End If
    '    If lblStartStation.Caption <> "起点站" Then
        ''              frmWizardAddBus.txtRouteID.Text = Trim(txtRoute.Text) & "[" & Trim(txtRouteName.Text) & "]"
        ''              '当frmaddBus窗体被打开,并调用了新增路线窗体时
        
        '              Select Case m_szfromstatus
        '                     Case "环境新增车次-新增线路"
        '                      frmAddBus.txtRouteId.Text = Trim(txtRoute.Text) & "[" & Trim(txtRouteName.Text) & "]"
        '                     Case "新增车次向导-新增线路"
        '                     frmWizardAddBus.txtRouteId.Text = Trim(txtRoute.Text) & "[" & Trim(txtRouteName.Text) & "]"
        '              End Select
    '    Else
    '    MsgBox "线路无路段", vbOKOnly, "提示"
    '    Exit Sub
    '    End If
        Unload Me
    End Select
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub cmdSection_Click()
'    If g_szStationID = "" Then
'        MsgBox "用户还未在系统管理中设定本站站点码,无法安排线路路段!", vbExclamation, "线路管理"
'        Exit Sub
'    End If
    frmArrangeSection.m_szRouteID = txtRoute.Text
'    Set frmArrangeSection.m_oRoute = m_oRoute
    frmArrangeSection.m_bIsParent = True
    frmArrangeSection.Show vbModal
End Sub

Private Sub Form_Load()
    AlignFormPos Me
    m_oRoute.Init g_oActiveUser
    Select Case Status
    Case EFormStatus.AddStatus
        txtRoute.Text = ""
        cmdOk.Caption = "新增(&A)"
        cmdSection.Enabled = False
        frmRoute.HelpContextID = 10000620
    Case EFormStatus.ModifyStatus
        RefreshRoute
        txtRoute.Enabled = False
        frmRoute.HelpContextID = 10000660
    Case EFormStatus.ShowStatus
        RefreshRoute
        txtRoute.Enabled = False
    End Select
    cmdOk.Enabled = False
    
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Public Sub RefreshRoute()
On Error GoTo ErrorHandle
    m_oRoute.Identify m_szRouteID
    txtRoute.Text = m_szRouteID
    txtRouteName.Text = m_oRoute.RouteName
'    txtFormula.Text = m_oRoute.TicketPriceFormula
    lblMileage.Caption = m_oRoute.Mileage & "公里"
    txtAnnotation.Text = m_oRoute.Annotation
    If m_oRoute.StartStationName = "" Then
        lblStartStation.Caption = "起点站"
        lblEndStation.Caption = "终点站"
        lblStartStation.Enabled = False
        lblEndStation.Enabled = False
    End If
    lblStartStation.Caption = m_oRoute.StartStationName
    lblStartStation.Tag = m_oRoute.StartStation
    lblEndStation.Caption = m_oRoute.EndStationName
    lblEndStation.Tag = m_oRoute.EndStation
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
'    If m_bIsParent Then
'        frmAllRoute.UpdateList m_oRoute.RouteID
'    End If
End Sub

Private Sub lblEndStation_Click()
    If Trim(lblEndStation.Caption) = "起点站" Then Exit Sub
    frmStation.szStationID = lblEndStation.Tag
    frmStation.Status = ModifyStatus
    frmStation.Show vbModal
End Sub

Private Sub lblStartStation_Click()
    If Trim(lblStartStation.Caption) = "起点站" Then Exit Sub
    frmStation.szStationID = lblStartStation.Tag
    frmStation.Status = ModifyStatus
    frmStation.Show vbModal
End Sub

Private Sub txtAnnotation_Change()
    IsSave
End Sub

Private Sub txtAnnotation_GotFocus()
    cmdOk.Default = False
End Sub

Private Sub txtAnnotation_LostFocus()
    cmdOk.Default = True
End Sub

Private Sub txtFormula_Change()
    IsSave
End Sub

'Private Sub txtRoute_ButtonClick()
'    Dim oShell As New STShell.CommDialog
'    Dim aszTemp() As String
'    If Status = AddStatus Then
'        MsgBox "请输入新增线路代码", vbInformation, "线路"
'        Exit Sub
'    End If
'    oShell.Init g_oActiveUser
'    aszTemp = oShell.SelectRoute(False)
'    Set oShell = Nothing
'    If ArrayLength(aszTemp) = 0 Then Exit Sub
'    txtRoute.Text = aszTemp(1, 1)
'    m_szRouteID = Trim(txtRoute.Text)
'    RefreshRoute
'End Sub

Private Sub txtRoute_Change()
    IsSave
End Sub

Private Sub IsSave()
    If txtRoute.Text = "" Or txtRouteName.Text = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub

Private Sub txtRouteName_Change()
    txtRouteName.Text = GetUnicodeBySize(txtRouteName.Text, 16)
    IsSave
End Sub
