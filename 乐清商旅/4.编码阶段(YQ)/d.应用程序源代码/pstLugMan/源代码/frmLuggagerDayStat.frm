VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmLuggagerDayStat 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "行包员受理日报"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6945
   StartUpPosition =   3  '窗口缺省
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   315
      Left            =   3960
      TabIndex        =   13
      Top             =   5160
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
      MICON           =   "frmLuggagerDayStat.frx":0000
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
      Left            =   5280
      TabIndex        =   14
      Top             =   5160
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
      MICON           =   "frmLuggagerDayStat.frx":001C
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
      Height          =   960
      Left            =   -120
      TabIndex        =   12
      Top             =   4800
      Width           =   8745
   End
   Begin pstLugMan.AddDel adUser 
      Height          =   2295
      Left            =   480
      TabIndex        =   11
      Top             =   2280
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4048
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
      LeftLabel       =   "待选列表(&L)"
      RightLabel      =   "已选列表(&R)"
      ButtonWidth     =   1215
      ButtonHeight    =   315
   End
   Begin VB.ComboBox cboAcceptType 
      Height          =   300
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1680
      Width           =   1755
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   6945
      TabIndex        =   2
      Top             =   0
      Width           =   6945
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择查询条件:"
         Height          =   180
         Left            =   480
         TabIndex        =   3
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -300
      TabIndex        =   1
      Top             =   840
      Width           =   7245
   End
   Begin VB.ComboBox cboSellStation 
      Height          =   300
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1680
      Width           =   1755
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   300
      Left            =   4680
      TabIndex        =   4
      Top             =   1200
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd"
      Format          =   61669376
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   300
      Left            =   1560
      TabIndex        =   5
      Top             =   1200
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd"
      Format          =   61669376
      CurrentDate     =   36572
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "托运方式(&A)"
      Height          =   180
      Left            =   3480
      TabIndex        =   9
      Top             =   1740
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "售 票 站(&T)"
      Height          =   180
      Left            =   480
      TabIndex        =   8
      Top             =   1740
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "结束日期(&E)"
      Height          =   180
      Left            =   3480
      TabIndex        =   7
      Top             =   1260
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "开始日期(&S)"
      Height          =   180
      Left            =   480
      TabIndex        =   6
      Top             =   1260
      Width           =   990
   End
End
Attribute VB_Name = "frmLuggagerDayStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Implements IConditionForm

Const cszFileName = "行包员结算日报.xls"

Public m_bOk As Boolean
Public m_vaSeller As Variant
Public m_dtWorkDate As Date
Public m_dtEndDate As Date
Public m_SellStation As String
Public m_AcceptType As String


Private Sub cmdOk_Click()
    m_vaSeller = adUser.RightData
    
    m_dtWorkDate = dtpBeginDate.Value
    m_dtEndDate = dtpEndDate.Value
    m_AcceptType = cboAcceptType.Text
    m_SellStation = cboSellStation.Text
    
    SaveRecentSeller m_vaSeller
    m_bOk = True
    Unload Me
End Sub

Private Sub Form_Load()
   m_bOk = False
    AlignFormPos Me
    dtpBeginDate.Value = DateAdd("d", -1, g_oParam.NowDate)
    dtpEndDate.Value = dtpBeginDate.Value
    
    FillSellStation cboSellStation
    FillAcceptType
    FillSellerEx
    EnableOK
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
    
    oUnit.Init m_oAUser
    oUnit.Identify g_oParam.UnitID
    aszUser = oUnit.GetAllUserEX(, ResolveDisplay(cboSellStation))
    nUserCount = ArrayLength(aszUser)
    If nUserCount > 0 Then
        
'        oUser.Init m_oAUser
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
    nCount = ArrayLength(adUser.RightData)
    cmdOk.Enabled = IIf(nCount > 0, True, False)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub cboSellStation_Change()
    FillSellerEx
End Sub
Private Sub cboSellStation_Click()
    cboSellStation_Change
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub adUser_DataChange()
    EnableOK
End Sub
'Private Property Get IConditionForm_FileName() As String
'    IConditionForm_FileName = cszFileName
'End Property

Private Sub FillAcceptType()
With cboAcceptType
   
    .AddItem GetLuggageTypeString(0)
   .AddItem GetLuggageTypeString(1)

    
End With
End Sub



Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub
