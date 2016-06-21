VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmSellerTicketDetail 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "售票员售票明细查询"
   ClientHeight    =   4935
   ClientLeft      =   2775
   ClientTop       =   4140
   ClientWidth     =   7035
   HelpContextID   =   60000240
   Icon            =   "frmSellerTicketDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame FraTicketNo 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   960
      Left            =   3780
      TabIndex        =   18
      Top             =   450
      Width           =   3000
      Begin VB.TextBox TxtToTicketNo 
         Height          =   315
         Left            =   1260
         TabIndex        =   20
         Top             =   555
         Width           =   1545
      End
      Begin VB.TextBox TxtFromTicketNo 
         Height          =   315
         Left            =   1260
         TabIndex        =   19
         Top             =   210
         Width           =   1545
      End
      Begin VB.Label LblToTicketNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "终止票号:"
         Height          =   180
         Left            =   300
         TabIndex        =   22
         Top             =   615
         Width           =   810
      End
      Begin VB.Label LblFromTicketNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起始票号:"
         Height          =   180
         Left            =   300
         TabIndex        =   21
         Top             =   270
         Width           =   810
      End
   End
   Begin VB.Frame FraDateTime 
      BackColor       =   &H00E0E0E0&
      Height          =   960
      Left            =   150
      TabIndex        =   13
      Top             =   450
      Width           =   3495
      Begin MSComCtl2.DTPicker dtpEnddate 
         Height          =   315
         Left            =   1485
         TabIndex        =   14
         Top             =   585
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   97910787
         CurrentDate     =   36819
      End
      Begin MSComCtl2.DTPicker dtpBeginDate 
         Height          =   315
         Left            =   1485
         TabIndex        =   15
         Top             =   210
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   97910787
         CurrentDate     =   36572
      End
      Begin VB.Label LblBeginTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间(&B):"
         Height          =   180
         Left            =   285
         TabIndex        =   17
         Top             =   270
         Width           =   1080
      End
      Begin VB.Label LblEndTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间(&E):"
         Height          =   180
         Left            =   285
         TabIndex        =   16
         Top             =   645
         Width           =   1080
      End
   End
   Begin VB.TextBox txtBusId 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1590
      TabIndex        =   12
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CheckBox ChkBusId 
      BackColor       =   &H00E0E0E0&
      Caption         =   "指定车次(&U):"
      Height          =   225
      Left            =   150
      TabIndex        =   11
      Top             =   1860
      Width           =   1395
   End
   Begin VB.CheckBox ChkTicketNo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "指定票号段"
      Height          =   255
      Left            =   3780
      TabIndex        =   10
      Top             =   450
      Width           =   1200
   End
   Begin VB.CheckBox ChkTime 
      BackColor       =   &H00E0E0E0&
      Caption         =   "指定时间段"
      Height          =   255
      Left            =   150
      TabIndex        =   9
      Top             =   510
      Value           =   1  'Checked
      Width           =   1200
   End
   Begin VB.ComboBox cboStatus 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   120
      Width           =   2595
   End
   Begin VB.CheckBox ChkUser 
      BackColor       =   &H00E0E0E0&
      Caption         =   "指定用户"
      Height          =   255
      Left            =   150
      TabIndex        =   4
      Top             =   2160
      Value           =   1  'Checked
      Width           =   1200
   End
   Begin VB.ComboBox cboSellStation 
      Height          =   300
      Left            =   4785
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1800
      Width           =   1995
   End
   Begin VB.TextBox txtPersonName 
      Height          =   270
      Left            =   720
      TabIndex        =   1
      Top             =   1470
      Width           =   2175
   End
   Begin VB.TextBox txtIDCardNo 
      Height          =   270
      Left            =   4440
      TabIndex        =   0
      Top             =   1470
      Width           =   2295
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   5970
      TabIndex        =   2
      Top             =   90
      Width           =   915
      _ExtentX        =   1614
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
      MICON           =   "frmSellerTicketDetail.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PSTTKAcc.AddDel adUser 
      Height          =   2805
      Left            =   150
      TabIndex        =   5
      Top             =   2160
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   4948
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
      LeftLabel       =   ""
      RightLabel      =   "已选用户(&R)"
      ButtonWidth     =   1215
      ButtonHeight    =   315
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   315
      Left            =   3780
      TabIndex        =   6
      Top             =   90
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
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
      MICON           =   "frmSellerTicketDetail.frx":0028
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
      Left            =   4920
      TabIndex        =   7
      Top             =   90
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
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
      MICON           =   "frmSellerTicketDetail.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "票状态(&S):"
      Height          =   180
      Left            =   150
      TabIndex        =   26
      Top             =   165
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&T):"
      Height          =   180
      Left            =   3780
      TabIndex        =   25
      Top             =   1860
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "姓名:"
      Height          =   180
      Left            =   240
      TabIndex        =   24
      Top             =   1530
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "证件号:"
      Height          =   180
      Left            =   3810
      TabIndex        =   23
      Top             =   1530
      Width           =   630
   End
End
Attribute VB_Name = "frmSellerTicketDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IConditionForm

Const cszFileName = "售票员售票明细查询.xls"

Public m_bOk As Boolean
Public m_vaSeller As Variant
Public m_dtBeginDateTime As Date
Public m_dtEndDateTime As Date
Public m_nStatus As Integer
'Private m_anTemp(1 To 7) As Integer
Private m_anTemp(1 To 7) As Integer
Public m_szFromTicketNo As String
Public m_szToTicketNo As String
Public m_szBusId As String
Private m_aszDepartment() As String
Public m_nStatus2 As Integer '0为售票员明细查询 1为本单位车票明细查询

Public m_szIDCardNo As String   '证件号
Public m_szPersonName As String '证件姓名

Private Sub adUser_DataChange()
    EnableOK
End Sub


Private Sub cboDepartment_Click()
    '更改部门刷新对应的售票员
    FillSeller
End Sub

Private Sub cboSellStation_Change()
   FillSellerEx
End Sub
Public Sub FillSellerEx()
    Dim oUnit As New Unit
    Dim aszUser() As String
    Dim aszUser2() As String
    Dim aszUser22() As String
    
    Dim oUser As New User
    Dim i As Integer, nUserCount As Integer
    Dim szRecentSeller As String
    Dim szTemp As String
    
    Dim nNoSelected As Integer, nSelected As Integer
    
    If m_nStatus2 = 0 Then
        oUnit.Init m_oActiveUser
        oUnit.Identify m_oParam.UnitID
    End If
    aszUser = oUnit.GetAllUserEX(, ResolveDisplay(cboSellStation))
    nUserCount = ArrayLength(aszUser)
    If nUserCount > 0 Then
        
'        oUser.Init m_oActiveUser
        szRecentSeller = GetRecentSeller()
        
        nNoSelected = 0
        nSelected = 0
        
        For i = 1 To nUserCount
'            oUser.Identify aszUser(i)
            szTemp = MakeDisplayString(aszUser(i, 1), aszUser(i, 2))
            If InStr(1, szRecentSeller, szTemp, vbTextCompare) = 0 Then
                nNoSelected = nNoSelected + 1
                ReDim Preserve aszUser2(1 To nNoSelected)
                aszUser2(nNoSelected) = szTemp
            Else
                nSelected = nSelected + 1
                ReDim Preserve aszUser22(1 To nSelected)
                aszUser22(nSelected) = szTemp
            End If
        Next
    End If
    adUser.LeftData = aszUser2
    adUser.RightData = aszUser22
End Sub
Private Sub cboSellStation_Click()
    cboSellStation_Change
End Sub

Private Sub ChkBusId_Click()
    txtBusId.Enabled = Not txtBusId.Enabled
    EnableBusID
    
End Sub



Private Sub ChkTicketNo_Click()
    FraTicketNo.Enabled = Not FraTicketNo.Enabled
    EnableTicketNoFra
End Sub

Private Sub ChkTime_Click()
    FraDateTime.Enabled = Not FraDateTime.Enabled
    EnableDateTime
End Sub

Private Sub ChkUser_Click()
    If ChkUser.Value = vbChecked Then
        adUser.Enabled = True
    Else
        adUser.Enabled = False
    End If
    EnableOK
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    
    If ChkUser.Value = vbChecked Then
       m_vaSeller = adUser.RightData
    Else
       m_vaSeller = ""
    End If
    If ChkTime.Value = vbChecked Then
        m_dtBeginDateTime = dtpBeginDate.Value
        m_dtEndDateTime = dtpEnddate.Value
    Else
        m_dtBeginDateTime = cszEmptyDateStr
        m_dtEndDateTime = cszForeverDateStr
    End If
    If ChkTicketNo.Value = vbChecked Then
       m_szFromTicketNo = TxtFromTicketNo.Text
       m_szToTicketNo = TxtToTicketNo.Text
    Else
        m_szFromTicketNo = ""
        m_szToTicketNo = ""
    End If
    
    m_szIDCardNo = Trim(txtIDCardNo.Text)
    m_szPersonName = Trim(txtPersonName.Text)
    
    If ChkBusId.Value = vbChecked Then
       m_szBusId = txtBusId.Text
    Else
       m_szBusId = ""
    End If
    
    SaveRecentSeller m_vaSeller
    
    m_bOk = True
    m_nStatus = m_anTemp(cboStatus.ListIndex + 1)
    Unload Me
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
    dtpEnddate.Value = ToDBDate(dtpBeginDate.Value) & " 23:59:59"
    
    FillSellStation cboSellStation
'    FillDepartment
    FillSellerEx
    FillTicketStatus
    EnableOK
    EnableTicketNoFra
    EnableDateTime
    EnableBusID
    
    m_szIDCardNo = ""
    m_szPersonName = ""
    
    adUser.Enabled = True
End Sub

''填充单位
'Private Sub FillDepartment()
'    Dim i As Integer
'    Dim nCount As Integer
'    '得到所有部门
'    m_aszDepartment = GetAllDepartment
'    '填充到cboDepartment中
'    cboDepartment.Clear
'    nCount = ArrayLength(m_aszDepartment)
'    For i = 1 To nCount
'        cboDepartment.AddItem m_aszDepartment(i)
'    Next i
'    If cboDepartment.ListCount > 0 Then
'        cboDepartment.ListIndex = 0
'    End If
'
'End Sub

Private Sub FillSeller()
'    Dim oUnit As New Unit
'    Dim aszUser() As String
'    Dim aszUser2() As String
'    Dim aszUser22() As String
'
'    Dim oUser As New User
'    Dim i As Integer, nUserCount As Integer
'    Dim szRecentSeller As String
'    Dim szTemp As String
'
'    Dim nNoSelected As Integer, nSelected As Integer
'
'    oUnit.Init m_oActiveUser
'    oUnit.Identify m_oParam.UnitID
'    aszUser = oUnit.GetAllUser()
'    nUserCount = ArrayLength(aszUser)
'    If nUserCount > 0 Then
'        oUser.Init m_oActiveUser
'        szRecentSeller = GetRecentSeller()
'
'        nNoSelected = 0
'        nSelected = 0
'        For i = 1 To nUserCount
'            oUser.Identify aszUser(i)
'            szTemp = MakeDisplayString(oUser.UserID, oUser.FullName)
'            If InStr(1, szRecentSeller, szTemp, vbTextCompare) = 0 Then
'                nNoSelected = nNoSelected + 1
'                ReDim Preserve aszUser2(1 To nNoSelected)
'                aszUser2(nNoSelected) = szTemp
'            Else
'                nSelected = nSelected + 1
'                ReDim Preserve aszUser22(1 To nSelected)
'                aszUser22(nSelected) = szTemp
'            End If
'        Next
'    End If
'    adUser.LeftData = aszUser2
'    adUser.RightData = aszUser22


    Dim oUnit As New Unit
    Dim aszUser() As String
    Dim aszUser2() As String
    Dim aszUser22() As String
    
    Dim oUser As New User
    Dim i As Integer, nUserCount As Integer
    Dim szRecentSeller As String
    Dim szTemp As String
    
    Dim nNoSelected As Integer, nSelected As Integer
    
    oUnit.Init m_oActiveUser
    oUnit.Identify m_oParam.UnitID
    aszUser = oUnit.GetAllUserEX()
    nUserCount = ArrayLength(aszUser)
    If nUserCount > 0 Then
        
'        oUser.Init m_oActiveUser
        szRecentSeller = GetRecentSeller()
        
        nNoSelected = 0
        nSelected = 0
        
        For i = 1 To nUserCount
'            oUser.Identify aszUser(i)
            szTemp = MakeDisplayString(aszUser(i, 1), aszUser(i, 2))
            If InStr(1, szRecentSeller, szTemp, vbTextCompare) = 0 Then
                nNoSelected = nNoSelected + 1
                ReDim Preserve aszUser2(1 To nNoSelected)
                aszUser2(nNoSelected) = szTemp
            Else
                nSelected = nSelected + 1
                ReDim Preserve aszUser22(1 To nSelected)
                aszUser22(nSelected) = szTemp
            End If
        Next
    End If
    adUser.LeftData = aszUser2
    adUser.RightData = aszUser22
    
End Sub

Private Sub EnableOK()
    Dim nCount As Integer
    If Not (ChkUser.Value = vbChecked) Then
        cmdOk.Enabled = True
    Else
        nCount = ArrayLength(adUser.RightData)
        cmdOk.Enabled = IIf(nCount > 0, True, False)
    End If
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Property Get IConditionForm_CustomData() As Variant

End Property

Private Property Get IConditionForm_FileName() As String
    IConditionForm_FileName = cszFileName
End Property

Private Property Get IConditionForm_RecordsetData() As Recordset

End Property

Private Sub FillTicketStatus()
    m_anTemp(1) = 0
    m_anTemp(2) = ST_TicketNormal
    m_anTemp(3) = ST_TicketSellChange   'ST_TicketNormal ST_TicketSellChange ST_TicketCanceled ST_TicketReturned ST_TicketChecked
    m_anTemp(4) = ST_TicketCanceled
    'm_anTemp(5) = ST_TicketChanged
    m_anTemp(5) = ST_TicketReturned
    m_anTemp(6) = ST_TicketChecked
    cboStatus.AddItem "所有状态"
    cboStatus.AddItem "正常售出(未检)"
    cboStatus.AddItem "改签"
    cboStatus.AddItem "废票"
    'cboStatus.AddItem "被改签"
    cboStatus.AddItem "退票"
    cboStatus.AddItem "已检"
    cboStatus.ListIndex = 0
End Sub

Public Sub EnableTicketNoFra()
    If FraTicketNo.Enabled = False Then
       TxtFromTicketNo.Enabled = False
       TxtToTicketNo.Enabled = False
       LblFromTicketNo.Enabled = False
       LblToTicketNo.Enabled = False
       TxtFromTicketNo.BackColor = &H8000000F
       TxtToTicketNo.BackColor = &H8000000F
    Else
       TxtFromTicketNo.Enabled = True
       TxtToTicketNo.Enabled = True
       LblFromTicketNo.Enabled = True
       LblToTicketNo.Enabled = True
       TxtFromTicketNo.BackColor = &H80000005
       TxtToTicketNo.BackColor = &H80000005
    End If
End Sub

Private Sub EnableDateTime()
    If FraDateTime.Enabled = False Then
       dtpBeginDate.Enabled = False
       dtpEnddate.Enabled = False
       LblBeginTime.Enabled = False
       LblEndTime.Enabled = False
    Else
       dtpBeginDate.Enabled = True
       dtpEnddate.Enabled = True
       LblBeginTime.Enabled = True
       LblEndTime.Enabled = True
    End If
End Sub


Private Sub EnableBusID()
    If Not txtBusId.Enabled Then
        txtBusId.BackColor = &H8000000F
        
    Else
        txtBusId.BackColor = &H80000005
    End If
End Sub

'Private Sub FillSellStation()
'    '判断用户属于哪个上车站,如果为空则填充一个空行,再填充所有的上车站
'
'    '否则只填充用户所属的上车站
'End Sub

Private Sub txtIDCardNo_GotFocus()
    With txtIDCardNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPersonName_GotFocus()
    With txtPersonName
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

