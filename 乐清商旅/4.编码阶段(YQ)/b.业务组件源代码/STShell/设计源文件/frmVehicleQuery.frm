VERSION 5.00
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmVehicleQuery 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "车辆查询"
   ClientHeight    =   4740
   ClientLeft      =   2745
   ClientTop       =   3720
   ClientWidth     =   7080
   Icon            =   "frmVehicleQuery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7080
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton cmdFind 
      Height          =   315
      Left            =   5460
      TabIndex        =   10
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "查询(&Q)"
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
      MICON           =   "frmVehicleQuery.frx":0C42
      PICN            =   "frmVehicleQuery.frx":0C5E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FText.asFlatTextBox txtVehicleType 
      Height          =   300
      Left            =   930
      TabIndex        =   9
      Top             =   900
      Width           =   1680
      _ExtentX        =   2963
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
      Text            =   "(全部)"
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
      Registered      =   -1  'True
   End
   Begin FText.asFlatTextBox txtBusOwner 
      Height          =   300
      Left            =   3510
      TabIndex        =   7
      Top             =   510
      Width           =   1680
      _ExtentX        =   2963
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
      Text            =   "(全部)"
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
      Registered      =   -1  'True
   End
   Begin FText.asFlatTextBox txtCompany 
      Height          =   300
      Left            =   3510
      TabIndex        =   3
      Top             =   120
      Width           =   1680
      _ExtentX        =   2963
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
      Text            =   "(全部)"
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
      Registered      =   -1  'True
   End
   Begin VB.TextBox txtVehicle 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   930
      TabIndex        =   5
      Top             =   510
      Width           =   1680
   End
   Begin VB.TextBox txtLicense 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   930
      TabIndex        =   1
      Top             =   120
      Width           =   1680
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2925
      Top             =   2445
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVehicleQuery.frx":0FF8
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVehicleQuery.frx":1154
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvVehicle 
      Height          =   2670
      Left            =   180
      TabIndex        =   13
      Top             =   1560
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   4710
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "车辆代码"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "车辆车牌"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "参运公司"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "车主"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "车型"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "座位数"
         Object.Width           =   1764
      EndProperty
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Height          =   315
      Left            =   5460
      TabIndex        =   11
      Top             =   510
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "frmVehicleQuery.frx":12B0
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
      Cancel          =   -1  'True
      Height          =   315
      Left            =   5610
      TabIndex        =   14
      Top             =   4320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "选择(&O)"
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
      MICON           =   "frmVehicleQuery.frx":12CC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   1290
      X2              =   6930
      Y1              =   1425
      Y2              =   1425
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车辆列表(&L):"
      Height          =   180
      Left            =   180
      TabIndex        =   12
      Top             =   1305
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车型(&T):"
      Height          =   180
      Left            =   180
      TabIndex        =   8
      Top             =   960
      Width           =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   1290
      X2              =   6930
      Y1              =   1410
      Y2              =   1410
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公司(&Z):"
      Height          =   180
      Left            =   2760
      TabIndex        =   2
      Top             =   180
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车主(&W):"
      Height          =   180
      Left            =   2760
      TabIndex        =   6
      Top             =   570
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "代码(&N):"
      Height          =   180
      Left            =   180
      TabIndex        =   4
      Top             =   570
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车牌(&P):"
      Height          =   180
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   720
   End
End
Attribute VB_Name = "frmVehicleQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'车辆查询
Option Explicit
'ListView的列位置
Const cnVehicleID = 0
Const cnLicenseTagNO = 1
Const cnTransportCompany = 2
Const cnOwner = 3
Const cnVehicleType = 4
Const cnTotalSeat = 5
Const cnCols = 6


Const cszChinaAll = "(全部)"


Public m_bOk As Boolean
Public m_oActiveUser As ActiveUser


Private m_aszVehicle() As String
Private m_oBaseInfo As New BaseInfo '基本信息BaseInfo


Private Sub cmdCancel_Click()
    Unload Me
End Sub

'获得选中的车次代码
Private Function GetAllVehicle() As String()
    On Error GoTo ErrorHandle
    Dim aszVehicle() As String
    Dim nSelectedVehicle As Integer
    Dim i As Integer
    '得到选择的个数
    nSelectedVehicle = 0
    For i = 1 To lvVehicle.ListItems.Count
        If lvVehicle.ListItems(i).Selected Then nSelectedVehicle = nSelectedVehicle + 1
    Next i
    If nSelectedVehicle > 0 Then
        ReDim aszVehicle(1 To nSelectedVehicle, 1 To cnCols) As String
    Else
        Exit Function
    End If
    nSelectedVehicle = 0
    For i = 1 To lvVehicle.ListItems.Count
        If lvVehicle.ListItems(i).Selected Then
            nSelectedVehicle = nSelectedVehicle + 1
            With lvVehicle.ListItems(i)
                aszVehicle(nSelectedVehicle, 1) = .Text
                aszVehicle(nSelectedVehicle, 2) = .SubItems(cnLicenseTagNO)
                aszVehicle(nSelectedVehicle, 3) = .SubItems(cnTransportCompany)
                aszVehicle(nSelectedVehicle, 4) = .SubItems(cnOwner)
                aszVehicle(nSelectedVehicle, 5) = .SubItems(cnVehicleType)
                aszVehicle(nSelectedVehicle, 6) = .SubItems(cnTotalSeat)
            End With
        End If
    Next
    GetAllVehicle = aszVehicle
    Exit Function
ErrorHandle:
    ShowErrorMsg
End Function


Private Sub cmdFind_Click()
    Dim szCompany As String
    Dim szOwner As String
    Dim szBusType As String
    Dim szLicense As String
    Dim i As Integer, nCount As Integer
    Dim ltTemp As ListItem
    Dim aszVehicle() As String
On Error GoTo ErrorHandle
    lvVehicle.ListItems.Clear
    szCompany = IIf(txtCompany.Text = cszChinaAll, "", ResolveDisplay(txtCompany.Text))
    szOwner = IIf(txtBusOwner.Text = cszChinaAll, "", ResolveDisplay(txtBusOwner.Text))
    szLicense = IIf(txtLicense.Text = cszChinaAll, "", txtLicense.Text)
    szBusType = IIf(txtVehicleType.Text = cszChinaAll, "", ResolveDisplay(txtVehicleType.Text))
    aszVehicle = m_oBaseInfo.GetVehicle(txtVehicle.Text, szCompany, szOwner, szBusType, szLicense)
    nCount = ArrayLength(aszVehicle)
    For i = 1 To nCount
        Set ltTemp = lvVehicle.ListItems.Add(, , Trim(aszVehicle(i, 1)), , "Run")
        If val(aszVehicle(i, 6)) <> ST_VehicleRun Then
            ltTemp.SmallIcon = "Stop"
        End If
        ltTemp.SubItems(cnLicenseTagNO) = aszVehicle(i, 2)
        ltTemp.SubItems(cnTransportCompany) = aszVehicle(i, 4)
        ltTemp.SubItems(cnOwner) = aszVehicle(i, 5)
        ltTemp.SubItems(cnVehicleType) = aszVehicle(i, 8)
        ltTemp.SubItems(cnTotalSeat) = aszVehicle(i, 3)
    Next
    If lvVehicle.ListItems.Count > 0 Then
        lvVehicle.ListItems(1).Selected = True
        If lvVehicle.Enabled Then lvVehicle.SetFocus
    End If
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub



Private Sub cmdOk_Click()
    '如果该窗口的调用者是车次车辆安排则新增车次车辆
On Error GoTo ErrorHandle
    m_aszVehicle = GetAllVehicle
    m_bOk = True
    Unload Me
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Not (Me.ActiveControl Is lvVehicle) Then
        cmdFind_Click
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    m_oBaseInfo.Init m_oActiveUser
    m_bOk = False
End Sub

Private Sub lvVehicle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvVehicle, ColumnHeader.Index
End Sub


Private Sub lvVehicle_DblClick()
    cmdOk_Click
End Sub


Private Sub lvVehicle_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdOk.Enabled = True
End Sub


Private Sub lvVehicle_KeyPress(KeyAscii As Integer)
    SendKeys "{TAB}"
End Sub

Private Sub txtBusOwner_ButtonClick()
    Dim oShell As New STShell.CommDialog
    Dim szaTemp() As String
    oShell.Init m_oActiveUser
    szaTemp = oShell.SelectOwner
    Set oShell = Nothing
    If ArrayLength(szaTemp) = 0 Then Exit Sub
    txtBusOwner.Text = MakeDisplayString(szaTemp(1, 1), Trim(szaTemp(1, 2)))

End Sub

Private Sub txtCompany_ButtonClick()
    Dim oShell As New STShell.CommDialog
    Dim szaTemp() As String
    oShell.Init m_oActiveUser
    szaTemp = oShell.SelectCompany(False)
    Set oShell = Nothing
    If ArrayLength(szaTemp) = 0 Then Exit Sub
    txtCompany.Text = MakeDisplayString(szaTemp(1, 1), Trim(szaTemp(1, 2)))

End Sub

Private Sub txtVehicleType_ButtonClick()
    Dim oShell As New STShell.CommDialog
    Dim szaTemp() As String
    oShell.Init m_oActiveUser
    szaTemp = oShell.SelectVehicleType(False)
    Set oShell = Nothing
    If ArrayLength(szaTemp) = 0 Then Exit Sub
    txtVehicleType.Text = MakeDisplayString(szaTemp(1, 1), Trim(szaTemp(1, 2)))

End Sub


Public Property Get GetSelectVehicle() As String()
    GetSelectVehicle = m_aszVehicle
End Property


Public Property Let MultiSelect(ByVal bNewValue As Boolean)
    lvVehicle.MultiSelect = bNewValue
End Property



