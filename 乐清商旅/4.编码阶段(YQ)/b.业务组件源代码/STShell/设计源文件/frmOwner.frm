VERSION 5.00
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmOwner 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "车主"
   ClientHeight    =   4545
   ClientLeft      =   4110
   ClientTop       =   4215
   ClientWidth     =   5865
   HelpContextID   =   2002001
   Icon            =   "frmOwner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5865
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   315
      Left            =   3240
      TabIndex        =   20
      Top             =   4080
      Width           =   1065
      _ExtentX        =   1879
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
      MICON           =   "frmOwner.frx":038A
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
      Left            =   4470
      TabIndex        =   21
      Top             =   4080
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "关闭"
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
      MICON           =   "frmOwner.frx":03A6
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
      TabIndex        =   16
      Top             =   0
      Width           =   6615
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   -30
         TabIndex        =   17
         Top             =   660
         Width           =   6885
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请修改或新增车主信息:"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   270
         Width           =   1890
      End
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
      MultiLine       =   -1  'True
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
      TabIndex        =   19
      Top             =   3810
      Width           =   8745
   End
   Begin FText.asFlatTextBox txtCompany 
      Height          =   300
      Left            =   1680
      TabIndex        =   11
      Top             =   2370
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
      TabIndex        =   13
      Top             =   2775
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
      TabIndex        =   15
      Top             =   3180
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
      TabIndex        =   12
      Top             =   2835
      Width           =   1080
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "备注(&R):"
      Height          =   255
      Left            =   540
      TabIndex        =   14
      Top             =   3210
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
      TabIndex        =   10
      Top             =   2415
      Width           =   1080
   End
End
Attribute VB_Name = "frmOwner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status As eFormStatus

Private moOwner As Owner
Public mszOwnerID As String
Public g_oActiveUser As ActiveUser
Public maszReturnItem As Variant    '返回值

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOk_Click()
On Error GoTo ErrHandle
    Select Case Status
        Case eFormStatus.EFS_AddNew
            moOwner.AddNew
            moOwner.OwnerID = txtOwnerID.Text
            moOwner.OwnerName = txtOwnerName.Text
            moOwner.IDCard = txtIDCard.Text
            moOwner.Contact = txtContact.Text
            moOwner.Address = txtAddress.Text
            moOwner.Company = ResolveDisplay(txtCompany.Text)
            moOwner.Annotation = txtAnnotation.Text
            moOwner.SplitCompanyID = ResolveDisplay(txtSplitCompany.Text)
            moOwner.Update
      Case eFormStatus.EFS_Modify
            moOwner.OwnerName = txtOwnerName.Text
            moOwner.IDCard = txtIDCard.Text
            moOwner.Contact = txtContact.Text
            moOwner.Address = txtAddress.Text
            moOwner.Annotation = txtAnnotation.Text
            moOwner.Company = ResolveDisplay(txtCompany.Text)
            moOwner.SplitCompanyID = ResolveDisplay(txtSplitCompany.Text)
            moOwner.Update
    End Select
        
    Dim aszInfo(1 To 3) As String
    aszInfo(1) = Trim(txtOwnerID.Text)
    aszInfo(2) = Trim(txtOwnerName.Text)
    aszInfo(3) = Trim(txtAnnotation.Text)
    maszReturnItem = aszInfo
'
'    '刷新基本信息窗体
'    Dim oListItem As ListItem
'    If Status = EFormStatus.EFS_Modify Then
'        frmBaseInfo.UpdateItemToList aszInfo
'        Unload Me
'        Exit Sub
'    End If
'    If Status = EFormStatus.EFS_AddNew Then
'        frmBaseInfo.AddItemToList aszInfo
'        RefreshOwner
'        txtOwnerID.SetFocus
'        Exit Sub
'    End If
    Exit Sub
ErrHandle:
    Status = EFS_AddNew
    ShowErrorMsg

End Sub

Private Sub Form_Activate()
    Dim vTmp As Variant
    maszReturnItem = vTmp
'    If Status = EFS_Show Then cmdCancel.SetFocus
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys "{TAB}"
    End Select
End Sub


Private Sub Form_Load()
    On Error GoTo ErrHandle
    Set moOwner = CreateObject("STBase.Owner")
    moOwner.Init g_oActiveUser
    Select Case Status
        Case eFormStatus.EFS_AddNew
           cmdOk.Caption = "新增(&A)"
            RefreshOwner
        Case eFormStatus.EFS_Modify
            txtOwnerID.Enabled = False
            RefreshOwner
        Case eFormStatus.EFS_Show
            cmdCancel.TabIndex = 0
            cmdCancel.Default = True
            lblCaption.Caption = "车主信息:"
            cmdOk.Visible = False
            txtOwnerID.Locked = True
            txtOwnerName.Locked = True
            txtAddress.Locked = True
            txtAnnotation.Locked = True
            txtCompany.Locked = True
            txtContact.Locked = True
            txtIDCard.Locked = True
            txtSplitCompany.Locked = True
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
        mszOwnerID = ""
        txtOwnerID.Text = ""
        txtAnnotation.Text = ""
        txtOwnerName.Text = ""
        txtIDCard.Text = ""
        txtAddress.Text = ""
        txtCompany.Text = ""
        txtContact.Text = ""
        txtSplitCompany.Text = ""
    Else
        txtOwnerID.Text = mszOwnerID
        moOwner.Identify mszOwnerID
        txtAnnotation.Text = moOwner.Annotation
        txtOwnerName.Text = moOwner.OwnerName
        txtIDCard.Text = moOwner.IDCard
        txtAddress.Text = moOwner.Address
        txtCompany.Text = MakeDisplayString(moOwner.Company, moOwner.CompanyName)
        txtContact.Text = moOwner.Contact
        txtSplitCompany.Text = MakeDisplayString(moOwner.SplitCompanyID, moOwner.SplitCompanyName)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moOwner = Nothing
End Sub

Private Sub txtAddress_Change()
    IsSave
End Sub

Private Sub txtAnnotation_Change()
    IsSave
End Sub

'Private Sub txtAnnotation_GotFocus()
'    cmdOk.Default = False
'End Sub
'
'Private Sub txtAnnotation_LostFocus()
'    cmdOk.Default = True
'End Sub

Private Sub txtContact_Change()
    IsSave
End Sub

'Private Sub txtContact_GotFocus()
'    cmdOk.Default = False
'End Sub
'
'Private Sub txtContact_LostFocus()
'    cmdOk.Default = True
'End Sub

Private Sub txtCompany_Change()
    IsSave
End Sub

Private Sub txtCompany_ButtonClick()
On Error GoTo ErrHandle
    If txtCompany.Locked Then Exit Sub
    Dim oShell As New STShell.CommDialog
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


Private Sub txtSplitCompany_Change()
    IsSave
End Sub

Private Sub txtSplitCompany_ButtonClick()
On Error GoTo ErrHandle
    If txtCompany.Locked Then Exit Sub
    Dim oShell As New STShell.CommDialog
    Dim aszTmp() As String
    oShell.Init g_oActiveUser
    aszTmp = oShell.SelectCompany
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtSplitCompany.Text = Trim(aszTmp(1, 1)) & "[" & Trim(aszTmp(1, 2)) & "]"
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
