VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmHalveEdit 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "加总平分属性"
   ClientHeight    =   4185
   ClientLeft      =   3675
   ClientTop       =   4305
   ClientWidth     =   4860
   Icon            =   "frmHalveEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   4860
   StartUpPosition =   1  '所有者中心
   Begin FText.asFlatTextBox txtRatio 
      Height          =   300
      Left            =   2085
      TabIndex        =   12
      Top             =   2475
      Width           =   1635
      _ExtentX        =   2884
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
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   -30
      TabIndex        =   1
      Top             =   3420
      Width           =   6045
      Begin RTComctl3.CoolButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   345
         Left            =   3450
         TabIndex        =   4
         Top             =   270
         Width           =   1110
         _ExtentX        =   1958
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
         MICON           =   "frmHalveEdit.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdok 
         Height          =   345
         Left            =   2085
         TabIndex        =   3
         Top             =   270
         Width           =   1110
         _ExtentX        =   1958
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
         MICON           =   "frmHalveEdit.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   30
      Left            =   -390
      TabIndex        =   0
      Top             =   795
      Width           =   7815
   End
   Begin FText.asFlatTextBox txtCompanyEx 
      Height          =   300
      Left            =   2085
      TabIndex        =   5
      Top             =   1995
      Width           =   1635
      _ExtentX        =   2884
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
   Begin FText.asFlatTextBox txtCompany 
      Height          =   300
      Left            =   2085
      TabIndex        =   6
      Top             =   1500
      Width           =   1635
      _ExtentX        =   2884
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
   Begin FText.asFlatTextBox txtRoute 
      Height          =   300
      Left            =   2085
      TabIndex        =   7
      Top             =   1020
      Width           =   1635
      _ExtentX        =   2884
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "说明:"
      Height          =   180
      Left            =   135
      TabIndex        =   14
      Top             =   2970
      Width           =   450
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "分成比率指的是参运公司所占的比率,如果为3:7开,则设置为0.3"
      Height          =   435
      Left            =   870
      TabIndex        =   13
      Top             =   2985
      Width           =   3345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "分成比率:"
      Height          =   180
      Left            =   1185
      TabIndex        =   11
      Top             =   2535
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "参运公司:"
      Height          =   180
      Left            =   1185
      TabIndex        =   10
      Top             =   1560
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "线路:"
      Height          =   180
      Left            =   1185
      TabIndex        =   9
      Top             =   1080
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "对方公司:"
      Height          =   180
      Left            =   1185
      TabIndex        =   8
      Top             =   2055
      Width           =   780
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "公司加总平分:"
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   225
      Width           =   1395
   End
End
Attribute VB_Name = "frmHalveEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_status As EFormStatus
Dim m_oHalve As New HalveCompany

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
On Error GoTo here
    If Trim(txtCompany.Text) = Trim(txtCompanyEx.Text) Then
        MsgBox "参动公司和对方公司，不能相同，请重新选择!", vbInformation, "加总平公"
        Exit Sub
    End If
    If m_status = ModifyStatus Then
        m_oHalve.CompanyID = ResolveDisplay(Trim(txtCompany.Text))
        m_oHalve.OtherCompanyId = ResolveDisplay(Trim(txtCompanyEx.Text))
        m_oHalve.RouteID = ResolveDisplay(Trim(txtRoute.Text))
        m_oHalve.Ratio = txtRatio.Text
        m_oHalve.Update
    ElseIf m_status = AddStatus Then
        m_oHalve.AddNew
        m_oHalve.CompanyID = ResolveDisplay(Trim(txtCompany.Text))
        m_oHalve.OtherCompanyId = ResolveDisplay(Trim(txtCompanyEx.Text))
        m_oHalve.RouteID = ResolveDisplay(Trim(txtRoute.Text))
        m_oHalve.Ratio = txtRatio.Text
        m_oHalve.Update
    End If
    '刷新主列表
    frmHalve.FilllvHavle
    Unload Me
Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub Form_Load()
    On Error GoTo err
    AlignFormPos Me
    m_oHalve.Init g_oActiveUser
    If m_status = ModifyStatus Then
        txtRoute.Text = frmHalve.lvHavle.SelectedItem.Text
        txtCompany.Text = frmHalve.lvHavle.SelectedItem.SubItems(1)
        txtCompanyEx.Text = frmHalve.lvHavle.SelectedItem.SubItems(2)
        txtRatio.Text = frmHalve.lvHavle.SelectedItem.SubItems(3)
        txtRoute.Enabled = False
        txtCompany.Enabled = False
    ElseIf m_status = AddStatus Then
        txtRoute.Text = ""
        txtCompany.Text = ""
        txtCompanyEx.Text = ""
        txtRatio.Text = 0.5
        
'        m_oHalve.AddNew
    End If
    Exit Sub
err:
ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    Unload Me
    '刷新列表
End Sub

Private Sub txtCompany_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectCompany
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtCompany.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))

Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

Private Sub txtCompanyEx_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectCompany
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtCompanyEx.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))

Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

Private Sub txtRatio_Change()
     txtRatio.Text = GetTextToNumeric(txtRatio.Text, False, True)
     
End Sub

Private Sub txtRatio_Validate(Cancel As Boolean)
    If txtRatio.Text > 1 Or txtRatio.Text < 0 Then
        MsgBox "分成比率必须大于0 且小于1", vbExclamation, Me.Caption
        Cancel = True
    End If
End Sub

Private Sub txtRoute_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectRoute
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtRoute.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))

Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

