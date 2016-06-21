VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmSelectSaler 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择售票员"
   ClientHeight    =   4590
   ClientLeft      =   4380
   ClientTop       =   2235
   ClientWidth     =   6525
   Icon            =   "frmSelectSaler.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   5040
      TabIndex        =   2
      Top             =   4110
      Width           =   1245
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
      MICON           =   "frmSelectSaler.frx":000C
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
      Left            =   3570
      TabIndex        =   1
      Top             =   4110
      Width           =   1245
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
      MICON           =   "frmSelectSaler.frx":0028
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
      TabIndex        =   8
      Top             =   3840
      Width           =   8745
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -90
      TabIndex        =   7
      Top             =   690
      Width           =   6885
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   -30
      ScaleHeight     =   705
      ScaleWidth      =   6615
      TabIndex        =   5
      Top             =   0
      Width           =   6615
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
   Begin PSTTKAcc.AddDel adUser 
      Height          =   2535
      Left            =   390
      TabIndex        =   0
      Top             =   1290
      Width           =   5805
      _ExtentX        =   10239
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
      Left            =   1290
      TabIndex        =   3
      Top             =   930
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   62652419
      UpDown          =   -1  'True
      CurrentDate     =   36572
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "时间(&B):"
      Height          =   180
      Left            =   420
      TabIndex        =   4
      Top             =   990
      Width           =   720
   End
End
Attribute VB_Name = "frmSelectSaler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_bOk As Boolean
Public m_vaSeller As Variant
Public m_dyTime As Date

Private Sub cmdCancel_Click()
    m_bOk = False
    Unload Me
End Sub

Private Sub cmdok_Click()
    
    m_vaSeller = adUser.RightData
    SaveRecentSeller m_vaSeller
    m_dyTime = dtpBeginDate.Value
    
    m_bOk = True
    Unload Me
End Sub
Private Sub adUser_DataChange()
    EnableOK
End Sub
Private Sub Form_Load()
    m_bOk = False
    FillSeller
    EnableOK
    
    dtpBeginDate.Value = m_dyTime
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


