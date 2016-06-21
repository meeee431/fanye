VERSION 5.00
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmCompany 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参运公司"
   ClientHeight    =   4515
   ClientLeft      =   1830
   ClientTop       =   3045
   ClientWidth     =   6330
   HelpContextID   =   2003601
   Icon            =   "frmCompany.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   315
      Left            =   3840
      TabIndex        =   12
      Top             =   4050
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
      MICON           =   "frmCompany.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtCompanyID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1635
      TabIndex        =   1
      Top             =   990
      Width           =   3990
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   -30
      ScaleHeight     =   795
      ScaleWidth      =   6315
      TabIndex        =   14
      Top             =   0
      Width           =   6315
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   0
         TabIndex        =   15
         Top             =   750
         Width           =   6885
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请修改或新增参运公司信息:"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   2250
      End
   End
   Begin VB.TextBox txtContact 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1635
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2505
      Width           =   3990
   End
   Begin VB.TextBox txtPrincipal 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1635
      TabIndex        =   7
      Top             =   2130
      Width           =   3990
   End
   Begin VB.TextBox txtSimpleCompany 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1635
      TabIndex        =   3
      Top             =   1380
      Width           =   3990
   End
   Begin VB.TextBox txtCompanyName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1635
      TabIndex        =   5
      Top             =   1740
      Width           =   3990
   End
   Begin FText.asFlatMemo txtAnnotation 
      Height          =   810
      Left            =   1635
      TabIndex        =   11
      Top             =   2880
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   1429
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
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   5070
      TabIndex        =   13
      Top             =   4050
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
      MICON           =   "frmCompany.frx":0028
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
      Height          =   750
      Left            =   -150
      TabIndex        =   17
      Top             =   3810
      Width           =   8745
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "备注(&R):"
      Height          =   180
      Left            =   510
      TabIndex        =   10
      Top             =   2940
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "联系方法(&L):"
      Height          =   180
      Left            =   510
      TabIndex        =   8
      Top             =   2565
      Width           =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公司简称(&K):"
      Height          =   180
      Left            =   510
      TabIndex        =   2
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "负责人(&D):"
      Height          =   180
      Left            =   510
      TabIndex        =   6
      Top             =   2190
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公司名称(&N):"
      Height          =   180
      Left            =   510
      TabIndex        =   4
      Top             =   1815
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公司代码(&I):"
      Height          =   180
      Left            =   510
      TabIndex        =   0
      Top             =   1035
      Width           =   1080
   End
End
Attribute VB_Name = "frmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status As EFormStatus

Private moCompany As Company
Public mszCompanyID As String
Public g_oActiveUser As ActiveUser
Public maszReturnItem As Variant    '返回值
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOk_Click()
On Error GoTo ErrHandle
    Select Case Status
        Case EFormStatus.EFS_AddNew
            moCompany.AddNew
            moCompany.CompanyId = txtCompanyID.Text
            moCompany.CompanyName = txtCompanyName.Text
            moCompany.CompanyShortName = txtSimpleCompany.Text
            moCompany.Annotation = txtAnnotation.Text
            moCompany.Contact = txtContact.Text
            moCompany.Principal = txtPrincipal.Text
            moCompany.Update
      Case EFormStatus.EFS_Modify
            moCompany.Identify txtCompanyID.Text
            moCompany.CompanyName = txtCompanyName.Text
            moCompany.CompanyShortName = txtSimpleCompany.Text
            moCompany.Annotation = txtAnnotation.Text
            moCompany.Contact = txtContact.Text
            moCompany.Principal = txtPrincipal.Text
            moCompany.Update
    End Select
        
    Dim aszInfo(1 To 3) As String
    aszInfo(1) = Trim(txtCompanyID.Text)
    aszInfo(2) = Trim(txtSimpleCompany.Text)
    aszInfo(3) = Trim(txtAnnotation.Text)
    maszReturnItem = aszInfo
''    '刷新基本信息窗体
''    Dim oListItem As ListItem
''    If Status = EFormStatus.EFS_Modify Then
''        frmBaseInfo.UpdateItemToList aszInfo
''        Unload Me
''        Exit Sub
''    End If
''    If Status = EFormStatus.EFS_AddNew Then
''        frmBaseInfo.AddItemToList aszInfo
''        RefreshCompany
''        txtCompanyID.SetFocus
''        Exit Sub
''    End If
    Unload Me
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Activate()
    Dim vTmp As Variant
    maszReturnItem = vTmp
End Sub

''Private Sub cmdQueryVehicle_Click()
''    frmQueryVehicleA.txtCompany.Text = Trim(txtCompanyID.Text) & "[" & Trim(txtCompanyName.Text) & "]"
''    frmQueryVehicleA.Show vbModal, Me
''End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    
           Case vbKeyReturn
                SendKeys "{TAB}"
    End Select
End Sub
Private Sub Form_Load()
    On Error GoTo ErrHandle
    '布置窗体
'    AlignFormPos Me
    
    Set moCompany = CreateObject("SNBase.Company")
    moCompany.Init g_oActiveUser
    Select Case Status
        Case EFormStatus.EFS_AddNew
           cmdOK.Caption = "新增(&A)"
            RefreshCompany
        Case EFormStatus.EFS_Modify
           txtCompanyID.Enabled = False
           RefreshCompany
        Case EFormStatus.EFS_Show
            cmdCancel.TabIndex = 0
            cmdCancel.Default = True
            lblCaption.Caption = "参运公司信息:"
            cmdOK.Visible = False
            txtAnnotation.Locked = True
            txtCompanyID.Locked = True
            txtCompanyName.Locked = True
            txtContact.Locked = True
            txtPrincipal.Locked = True
            txtSimpleCompany.Locked = True
            RefreshCompany
    End Select
    cmdOK.Enabled = False
    
    Exit Sub
ErrHandle:
    Status = EFS_AddNew
    ShowErrorMsg
End Sub

Public Sub RefreshCompany()
    If Status = EFS_AddNew Then
        txtCompanyID.Text = ""
        txtCompanyName.Text = ""
        txtSimpleCompany.Text = ""
        txtPrincipal.Text = ""
        txtContact.Text = ""
        txtAnnotation.Text = ""
    Else
        moCompany.Identify Trim(mszCompanyID)
        txtCompanyID.Text = mszCompanyID
        moCompany.Identify mszCompanyID
        txtCompanyName.Text = moCompany.CompanyName
        txtSimpleCompany.Text = moCompany.CompanyShortName
        txtPrincipal.Text = moCompany.Principal
        txtContact.Text = moCompany.Contact
        txtAnnotation.Text = moCompany.Annotation
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moCompany = Nothing
'    SaveFormPos Me
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
    FormatTextBoxBySize txtContact, 20
End Sub

'Private Sub txtContact_GotFocus()
'    cmdOk.Default = False
'End Sub
'
'Private Sub txtContact_LostFocus()
'    cmdOk.Default = True
'End Sub

Private Sub txtCompanyId_Change()
    IsSave
    FormatTextBoxBySize txtCompanyID, 12
End Sub

Private Sub IsSave()
    If txtCompanyID.Text = "" Or txtCompanyName.Text = "" Then
        cmdOK.Enabled = False
''        cmdQueryVehicle.Enabled = False
    Else
        cmdOK.Enabled = True
''        cmdQueryVehicle.Enabled = True
    End If
End Sub

Private Sub txtCompanyName_Change()
    IsSave
    FormatTextBoxBySize txtCompanyName, 30
End Sub

Private Sub txtPrincipal_Change()
    IsSave
    FormatTextBoxBySize txtPrincipal, 10
End Sub

Private Sub txtSimpleCompany_Change()
    IsSave
    FormatTextBoxBySize txtSimpleCompany, 10
End Sub
