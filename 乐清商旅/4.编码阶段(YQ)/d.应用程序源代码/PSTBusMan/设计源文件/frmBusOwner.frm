VERSION 5.00
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmBusOwner 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "车主"
   ClientHeight    =   4950
   ClientLeft      =   4650
   ClientTop       =   3090
   ClientWidth     =   6015
   HelpContextID   =   10000110
   Icon            =   "frmBusOwner.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6015
   Begin VB.TextBox txtAccount 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1680
      TabIndex        =   11
      Top             =   2370
      Width           =   3540
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   210
      TabIndex        =   25
      Top             =   4545
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "frmBusOwner.frx":038A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtOwnerID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1680
      MaxLength       =   5
      TabIndex        =   1
      Top             =   840
      Width           =   930
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   -30
      ScaleHeight     =   705
      ScaleWidth      =   6615
      TabIndex        =   21
      Top             =   0
      Width           =   6615
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   -30
         TabIndex        =   22
         Top             =   660
         Width           =   6885
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请修改或新增车主信息:"
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   270
         Width           =   1890
      End
   End
   Begin RTComctl3.CoolButton cmdOk 
      Height          =   315
      Left            =   3060
      TabIndex        =   19
      Top             =   4545
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
      MICON           =   "frmBusOwner.frx":03A6
      PICN            =   "frmBusOwner.frx":03C2
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
      Left            =   4290
      TabIndex        =   20
      Top             =   4545
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
      MICON           =   "frmBusOwner.frx":075C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdQueryVehicle 
      Height          =   315
      Left            =   1650
      TabIndex        =   18
      Top             =   4545
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "车辆(&V)..."
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
      MICON           =   "frmBusOwner.frx":0778
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1680
      TabIndex        =   9
      Top             =   1980
      Width           =   3540
   End
   Begin VB.TextBox txtContact 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1680
      TabIndex        =   7
      Top             =   1590
      Width           =   3540
   End
   Begin VB.TextBox txtIDCard 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1680
      TabIndex        =   5
      Top             =   1215
      Width           =   3540
   End
   Begin VB.TextBox txtOwnerName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3840
      MaxLength       =   5
      TabIndex        =   3
      Top             =   840
      Width           =   1380
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   780
      Left            =   -150
      TabIndex        =   24
      Top             =   4305
      Width           =   8745
   End
   Begin FText.asFlatTextBox txtCompany 
      Height          =   300
      Left            =   1680
      TabIndex        =   13
      Top             =   2760
      Width           =   3540
      _ExtentX        =   6244
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
      ButtonPressedBackColor=   -2147483627
      Text            =   ""
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
   End
   Begin FText.asFlatTextBox txtSplitCompany 
      Height          =   300
      Left            =   1680
      TabIndex        =   15
      Top             =   3165
      Width           =   3540
      _ExtentX        =   6244
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
      ButtonPressedBackColor=   -2147483627
      Text            =   ""
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
   End
   Begin FText.asFlatMemo txtAnnotation 
      Height          =   630
      Left            =   1680
      TabIndex        =   17
      Top             =   3570
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   1111
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
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "帐号(&O):"
      Height          =   180
      Left            =   555
      TabIndex        =   10
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "地址(&E):"
      Height          =   180
      Left            =   555
      TabIndex        =   8
      Top             =   2010
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "拆帐公司(&T):"
      Height          =   180
      Left            =   555
      TabIndex        =   14
      Top             =   3225
      Width           =   1080
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "备注(&R):"
      Height          =   255
      Left            =   540
      TabIndex        =   16
      Top             =   3600
      Width           =   1020
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "联系方法(&L):"
      Height          =   180
      Left            =   555
      TabIndex        =   6
      Top             =   1665
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "身份证号(&D):"
      Height          =   180
      Left            =   555
      TabIndex        =   4
      Top             =   1275
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车主名称(&N):"
      Height          =   180
      Left            =   2715
      TabIndex        =   2
      Top             =   915
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车主代码(&I):"
      Height          =   180
      Left            =   555
      TabIndex        =   0
      Top             =   915
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参运公司(&C):"
      Height          =   180
      Left            =   555
      TabIndex        =   12
      Top             =   2805
      Width           =   1080
   End
End
Attribute VB_Name = "frmBusOwner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status As EFormStatus

Public m_szOwnerID As String
Private m_oOwner As Owner
Public m_bIsParent As Boolean '是否是父窗体调用


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo ErrHandle
    Select Case Status
        Case EFormStatus.EFS_AddNew
            m_oOwner.AddNew
            m_oOwner.OwnerID = txtOwnerID.Text
            m_oOwner.OwnerName = txtOwnerName.Text
            m_oOwner.IDCard = txtIDCard.Text
            m_oOwner.Contact = txtContact.Text
            m_oOwner.Address = txtAddress.Text
            m_oOwner.Company = ResolveDisplay(txtCompany.Text)
            m_oOwner.Annotation = txtAnnotation.Text
            m_oOwner.SplitCompanyID = ResolveDisplay(txtSplitCompany.Text)
            m_oOwner.AccountID = txtAccount.Text
            m_oOwner.Update
      Case EFormStatus.EFS_Modify
            m_oOwner.OwnerName = txtOwnerName.Text
            m_oOwner.IDCard = txtIDCard.Text
            m_oOwner.Contact = txtContact.Text
            m_oOwner.Address = txtAddress.Text
            m_oOwner.Annotation = txtAnnotation.Text
            m_oOwner.Company = ResolveDisplay(txtCompany.Text)
            m_oOwner.SplitCompanyID = ResolveDisplay(txtSplitCompany.Text)
            m_oOwner.AccountID = txtAccount.Text
            m_oOwner.Update
    End Select
        
    Dim aszInfo(1 To 3) As String
    aszInfo(1) = Trim(txtOwnerID.Text)
    aszInfo(2) = Trim(txtOwnerName.Text)
    aszInfo(3) = Trim(txtAccount.Text)
    
    '刷新基本信息窗体
    Dim oListItem As ListItem
    If Status = EFormStatus.EFS_Modify Then
        frmBaseInfo.UpdateList aszInfo
        Unload Me
        Exit Sub
    End If
    If Status = EFormStatus.EFS_AddNew Then
        frmBaseInfo.AddList aszInfo
        RefreshOwner
        txtOwnerID.SetFocus
        Exit Sub
    End If
    Exit Sub
ErrHandle:
    Status = EFS_AddNew
    ShowErrorMsg

End Sub

Private Sub cmdQueryVehicle_Click()
    Dim oShell As New CommDialog
    oShell.Init g_oActiveUser
    oShell.SelectVehicle , , txtOwnerID.Text
End Sub


Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys "{TAB}"
    End Select
End Sub


Private Sub Form_Load()
    On Error GoTo ErrHandle
    '布置窗体
    AlignFormPos Me
    
    Set m_oOwner = CreateObject("STBase.Owner")
    m_oOwner.Init g_oActiveUser
    Select Case Status
        Case EFormStatus.EFS_AddNew
           cmdOk.Caption = "新增(&A)"
            RefreshOwner
        Case EFormStatus.EFS_Modify
           txtOwnerID.Enabled = False
           RefreshOwner
        Case EFormStatus.EFS_Show
           txtOwnerID.Enabled = False
           RefreshOwner
    End Select
    cmdOk.Enabled = False
    
    Exit Sub
ErrHandle:
    Status = EFS_AddNew
    ShowErrorMsg
End Sub
Private Sub RefreshOwner()
    If Status = EFS_AddNew Then
        m_szOwnerID = ""
        txtOwnerID.Text = ""
        txtAnnotation.Text = ""
        txtOwnerName.Text = ""
        txtIDCard.Text = ""
        txtAddress.Text = ""
        txtCompany.Text = ""
        txtContact.Text = ""
        txtSplitCompany.Text = ""
        txtAccount.Text = ""
    Else
        txtOwnerID.Text = m_szOwnerID
        m_oOwner.Identify m_szOwnerID
        txtAnnotation.Text = m_oOwner.Annotation
        txtOwnerName.Text = m_oOwner.OwnerName
        txtIDCard.Text = m_oOwner.IDCard
        txtAddress.Text = m_oOwner.Address
        txtCompany.Text = MakeDisplayString(m_oOwner.Company, m_oOwner.CompanyName)
        txtContact.Text = m_oOwner.Contact
        txtSplitCompany.Text = MakeDisplayString(m_oOwner.SplitCompanyID, m_oOwner.SplitCompanyName)
        txtAccount.Text = m_oOwner.AccountID
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_oOwner = Nothing
    SaveFormPos Me
    m_bIsParent = False
End Sub

Private Sub txtAddress_Change()
    IsSave
End Sub

Private Sub txtAnnotation_Change()
    IsSave
End Sub

Private Sub txtAccount_Change()
    IsSave
End Sub

Private Sub txtAnnotation_GotFocus()
    cmdOk.Default = False
End Sub

Private Sub txtAnnotation_LostFocus()
    cmdOk.Default = True
End Sub

Private Sub txtContact_Change()
    IsSave
End Sub

Private Sub txtContact_GotFocus()
'    cmdOk.Default = False
End Sub

Private Sub txtContact_LostFocus()
'    cmdOk.Default = True
End Sub

Private Sub txtCompany_Change()
    IsSave
End Sub

Private Sub txtCompany_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init g_oActiveUser
    aszTmp = oShell.SelectCompany
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtCompany.Text = Trim(aszTmp(1, 1)) & "[" & Trim(aszTmp(1, 2)) & "]"
    txtSplitCompany.Text = txtCompany.Text
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub txtIDCard_Change()
    IsSave
    FormatTextBoxBySize txtIDCard, 18
End Sub

Private Sub txtOwnerID_Change()
    IsSave
    FormatTextBoxBySize txtOwnerID, 4
End Sub

'Private Sub txtOwnerID_Click()
'    Dim oShell As New STShell.CommDialog
'    Dim aszTmp() As String
'    If Status = EFS_AddNew Then
'        MsgBox "请输入新增车主代码", vbInformation, "车主"
'        Exit Sub
'    End If
'    oShell.Init g_oActiveUser
'    aszTmp = oShell.SelectOwner(False)
'    Set oShell = Nothing
'    If ArrayLength(aszTmp) = 0 Then Exit Sub
'    txtOwnerID.Text = aszTmp(1, 1)
'End Sub

Private Sub IsSave()
    If txtCompany.Text = "" Or txtOwnerID.Text = "" Or txtOwnerName.Text = "" Or txtSplitCompany.Text = "" Then
            
            
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub

Private Sub txtOwnerName_Change()
    IsSave
    FormatTextBoxBySize txtOwnerName, 10

End Sub
Private Sub UpDateAll()
    If m_bIsParent Then
        If Status = EFS_AddNew Then
            frmBaseInfo.AddList Trim(txtOwnerID.Text)
        Else
            frmBaseInfo.UpdateList Trim(txtOwnerID.Text)
        End If
    End If
End Sub

Private Sub txtSplitCompany_Change()
    IsSave
End Sub

Private Sub txtSplitCompany_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init g_oActiveUser
    aszTmp = oShell.SelectCompany
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtSplitCompany.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
