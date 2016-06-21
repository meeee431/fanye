VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmCheckerEveryMonth 
   BackColor       =   &H00E0E0E0&
   Caption         =   "检票员检票统计月报"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   6945
   StartUpPosition =   2  '屏幕中心
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   5340
      TabIndex        =   2
      Top             =   4650
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
      MICON           =   "frmCheckerEveryMonth.frx":0000
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
      Left            =   3945
      TabIndex        =   3
      Top             =   4650
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
      MICON           =   "frmCheckerEveryMonth.frx":001C
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
      Height          =   315
      Left            =   2550
      TabIndex        =   0
      Top             =   4650
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
      MICON           =   "frmCheckerEveryMonth.frx":0038
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
      Height          =   3120
      Left            =   -150
      TabIndex        =   7
      Top             =   4380
      Width           =   8745
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -90
      TabIndex        =   6
      Top             =   690
      Width           =   7125
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   -30
      ScaleHeight     =   705
      ScaleWidth      =   7005
      TabIndex        =   4
      Top             =   0
      Width           =   7005
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择查询条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   5
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.ComboBox cboSellStation 
      Height          =   300
      Left            =   1575
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1380
      Width           =   4725
   End
   Begin PSTTKAcc.AddDel adUser 
      Height          =   2535
      Left            =   480
      TabIndex        =   8
      Top             =   1800
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4471
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
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   315
      Left            =   1575
      TabIndex        =   9
      Top             =   900
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62652416
      UpDown          =   -1  'True
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   315
      Left            =   4650
      TabIndex        =   10
      Top             =   900
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62652416
      UpDown          =   -1  'True
      CurrentDate     =   36572
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   510
      TabIndex        =   13
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   3570
      TabIndex        =   12
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&T):"
      Height          =   180
      Left            =   510
      TabIndex        =   11
      Top             =   1440
      Width           =   900
   End
End
Attribute VB_Name = "frmCheckerEveryMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IConditionForm

Const cszCheckerFileName = "检票员检票统计月报模板.xls"

Public m_bCheck As Boolean


Public m_bOk As Boolean
Public m_vaSeller As Variant
Public m_dtWorkDate As Date
Public m_dtEndDate As Date

Private Sub adUser_DataChange()
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
    
    oUnit.Init m_oActiveUser
    oUnit.Identify m_oParam.UnitID
    aszUser = oUnit.GetAllUserEX(, ResolveDisplay(cboSellStation))
    nUserCount = ArrayLength(aszUser)
    If nUserCount > 0 Then
        
'        oUser.Init m_oActiveUser
        szRecentSeller = GetRecentSeller()
        
        nNoSelected = 0
        nSelected = 0
        
        For i = 1 To nUserCount
'            oUser.Identify aszUser(i)
'            szTemp = aszUser(i, 1)
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
Private Sub cboSellStation_Change()
    FillSellerEx
End Sub
Private Sub cboSellStation_Click()
    cboSellStation_Change
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    
    m_vaSeller = adUser.RightData
    
    m_dtWorkDate = dtpBeginDate.Value
    m_dtEndDate = dtpEndDate.Value
    
    SaveRecentSeller m_vaSeller
    
    
    
    
    m_bOk = True
    
    
    
    
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
    Dim dyNow As Date
    dyNow = m_oParam.NowDate
    
    dtpBeginDate.Value = dyNow
    dtpEndDate.Value = dyNow
    
    FillSellStation cboSellStation
    
    FillSellerEx
    EnableOK
End Sub

Private Sub FillSeller()
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
    nCount = ArrayLength(adUser.RightData)
    cmdOk.Enabled = IIf(nCount > 0, True, False)
End Sub


Private Property Get IConditionForm_CustomData() As Variant

End Property

Private Property Get IConditionForm_FileName() As String
        IConditionForm_FileName = cszCheckerFileName
End Property

Private Property Get IConditionForm_RecordsetData() As Recordset

End Property

'Private Sub FillSellStation()
'    '判断用户属于哪个上车站,如果为空则填充一个空行,再填充所有的上车站
'
'    '否则只填充用户所属的上车站
'End Sub


