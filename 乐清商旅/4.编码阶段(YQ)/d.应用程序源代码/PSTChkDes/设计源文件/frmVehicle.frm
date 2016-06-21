VERSION 5.00
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmVehicle 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "车辆--车辆信息"
   ClientHeight    =   5175
   ClientLeft      =   3210
   ClientTop       =   2925
   ClientWidth     =   6915
   HelpContextID   =   2008201
   Icon            =   "frmVehicle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin FText.asFlatTextBox txtVehicleModel 
      Height          =   300
      Left            =   4350
      TabIndex        =   7
      Top             =   1365
      Width           =   1845
      _ExtentX        =   3254
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
   Begin VB.TextBox txtVehicleID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1770
      TabIndex        =   1
      Top             =   990
      Width           =   1335
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   7185
      TabIndex        =   24
      Top             =   0
      Width           =   7185
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   0
         TabIndex        =   25
         Top             =   750
         Width           =   7215
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请修改或新增车辆信息:"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   300
         Width           =   1890
      End
   End
   Begin VB.TextBox txtCardId 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1770
      TabIndex        =   5
      Top             =   1365
      Width           =   1335
   End
   Begin VB.OptionButton optRun 
      BackColor       =   &H00E0E0E0&
      Caption         =   "运行车辆(&R)"
      Height          =   210
      Left            =   1770
      TabIndex        =   20
      Top             =   4170
      Value           =   -1  'True
      Width           =   1425
   End
   Begin VB.OptionButton optStop 
      BackColor       =   &H00E0E0E0&
      Caption         =   "停班车辆(&U)"
      Height          =   240
      Left            =   3210
      TabIndex        =   21
      Top             =   4155
      Width           =   1290
   End
   Begin VB.TextBox txtLicense 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4350
      TabIndex        =   3
      Top             =   990
      Width           =   1845
   End
   Begin FText.asFlatTextBox txtCompany 
      Height          =   300
      Left            =   1770
      TabIndex        =   9
      Top             =   1740
      Width           =   4425
      _ExtentX        =   7805
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
   Begin FText.asFlatTextBox txtOwner 
      Height          =   300
      Left            =   1770
      TabIndex        =   13
      Top             =   2520
      Width           =   4425
      _ExtentX        =   7805
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
   Begin FText.asFlatTextBox txtSplitCompanyID 
      Height          =   300
      Left            =   1770
      TabIndex        =   11
      Top             =   2130
      Width           =   4425
      _ExtentX        =   7805
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
   Begin FText.asFlatSpinEdit txtStartSeatNo 
      Height          =   300
      Left            =   1965
      TabIndex        =   15
      Top             =   2910
      Width           =   1140
      _ExtentX        =   2011
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
      Text            =   "1"
      ButtonBackColor =   -2147483633
      Value           =   1
   End
   Begin FText.asFlatSpinEdit txtEndSeatNo 
      Height          =   300
      Left            =   4515
      TabIndex        =   17
      Top             =   2910
      Width           =   1140
      _ExtentX        =   2011
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
      Text            =   "40"
      ButtonBackColor =   -2147483633
      Value           =   40
   End
   Begin FText.asFlatMemo txtAnnotation 
      Height          =   750
      Left            =   1770
      TabIndex        =   19
      Top             =   3300
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   1323
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
   Begin RTComctl3.CoolButton cmdOk 
      Height          =   315
      Left            =   4200
      TabIndex        =   22
      Top             =   4710
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
      MICON           =   "frmVehicle.frx":0C42
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
      Height          =   315
      Left            =   5430
      TabIndex        =   23
      Top             =   4710
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
      MICON           =   "frmVehicle.frx":0C5E
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
      Left            =   -120
      TabIndex        =   27
      Top             =   4470
      Width           =   8745
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "拆帐公司(&L):"
      Height          =   180
      Left            =   660
      TabIndex        =   10
      Top             =   2190
      Width           =   1080
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "备注(&R):"
      Height          =   180
      Left            =   660
      TabIndex        =   18
      Top             =   3390
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "报班卡号(&C):"
      Height          =   180
      Left            =   660
      TabIndex        =   4
      Top             =   1410
      Width           =   1080
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "终止座位号(&Q):"
      Height          =   180
      Left            =   3210
      TabIndex        =   16
      Top             =   2985
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车辆代码(&I):"
      Height          =   180
      Left            =   660
      TabIndex        =   0
      Top             =   1035
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车辆车牌(&N):"
      Height          =   180
      Left            =   3210
      TabIndex        =   2
      Top             =   1035
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参运公司(&D):"
      Height          =   180
      Left            =   660
      TabIndex        =   8
      Top             =   1785
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车辆车主(&E):"
      Height          =   180
      Left            =   660
      TabIndex        =   12
      Top             =   2580
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车辆车型(&S):"
      Height          =   180
      Left            =   3210
      TabIndex        =   6
      Top             =   1410
      Width           =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "起始座位号(&Q):"
      Height          =   180
      Left            =   660
      TabIndex        =   14
      Top             =   2985
      Width           =   1260
   End
End
Attribute VB_Name = "frmVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mszBusID As String
Public Status As EFormStatus
Private moVehicle As Vehicle  '车辆对象 Vehicle
Public mszVehicleId As String
Dim mszOldStartSeat As String, mszOldEndSeat As String
Public g_oActiveUser As ActiveUser
Public maszReturnItem As Variant    '返回值

''Private Sub cmdBus_Click()
''    frmVehicleBus.Init moVehicle
''    frmVehicleBus.Show vbModal
''End Sub

''Public Sub cmdSetSeatType_Click()
''   frmSetVehicleSeatType.m_nEndSeatNo = CInt(val(txtEndSeatNo.Text))
''   frmSetVehicleSeatType.m_nStartSeatNo = CInt(val(txtStartSeatNo.Text))
''   frmSetVehicleSeatType.m_szVehicleId = txtVehicleID.Text
''   frmSetVehicleSeatType.Show vbModal
''End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys "{TAB}"
    End Select
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Dim vTmp As Variant
    maszReturnItem = vTmp
End Sub

Private Sub cmdOk_Click()
    Dim szMsg As String
    Select Case Status
        Case EFormStatus.EFS_AddNew
            moVehicle.AddNew
            moVehicle.VehicleId = txtVehicleID.Text
            moVehicle.VehicleModel = ResolveDisplay(txtVehicleModel.Text)
            moVehicle.SplitCompanyID = ResolveDisplay(txtSplitCompanyID.Text)
            moVehicle.CardID = txtCardId.Text
            moVehicle.Company = ResolveDisplay(txtCompany.Text)
            moVehicle.LicenseTag = txtLicense.Text
            moVehicle.Owner = ResolveDisplay(txtOwner.Text)
            moVehicle.SeatCount = Val(txtEndSeatNo.Text) + 1 - Val(txtStartSeatNo.Text)
            moVehicle.StartSeatNo = Val(txtStartSeatNo.Text)
            If optRun.Value = True Then
                moVehicle.Status = ST_VehicleRun
            Else
                moVehicle.Status = ST_VehicleStop
            End If
            moVehicle.Annotation = txtAnnotation.Text
            
            If mszBusID <> "" Then
             moVehicle.ProjectBusID = mszBusID
            End If
            moVehicle.Update
      Case EFormStatus.EFS_Modify
        moVehicle.Identify txtVehicleID.Text
        moVehicle.CardID = txtCardId.Text
        moVehicle.Company = ResolveDisplay(txtCompany.Text)
        moVehicle.LicenseTag = txtLicense.Text
        moVehicle.Company = ResolveDisplay(txtVehicleModel.Text)
        moVehicle.SplitCompanyID = ResolveDisplay(txtSplitCompanyID.Text)
        moVehicle.Owner = ResolveDisplay(txtOwner.Text)
        moVehicle.SeatCount = Val(txtEndSeatNo.Text) + 1 - Val(txtStartSeatNo.Text)
        moVehicle.StartSeatNo = Val(txtStartSeatNo.Text)
        moVehicle.Annotation = txtAnnotation.Text
    
        If Val(mszOldEndSeat) <> Val(txtEndSeatNo.Text) Then
            szMsg = szMsg & "结束座位号改变!"
        End If
        If Val(mszOldStartSeat) <> Val(txtStartSeatNo.Text) Then
          szMsg = szMsg & "起始座位号改变!"
        End If
    End Select
        
        
        '在非基本信息的其它窗体调用时,忽略基本信息窗体的处理
''    If frmBaseInfo Is Nothing Then Exit Sub
''    If frmBaseInfo.tvBaseItem.SelectedItem.Key <> "KVehicle" Then Exit Sub
    
    Dim aszInfo(0 To 3) As String
    If optStop Then aszInfo(0) = "STOP"
    aszInfo(1) = Trim(txtVehicleID.Text)
    aszInfo(2) = Trim(txtLicense.Text)
    aszInfo(3) = EncodeString("所属公司:" & Trim(moVehicle.Company)) & _
                        EncodeString("车主:" & Trim(moVehicle.Owner)) & _
                        EncodeString("车型:" & Trim(moVehicle.Company)) & _
                        EncodeString("座位数:" & moVehicle.SeatCount)
    maszReturnItem = aszInfo
    '刷新基本信息窗体
''    Dim oListItem As ListItem
''    If Status = EFormStatus.EFS_Modify Then
''        '如果座位有更改,则
''        If szMsg <> "" Then
''            MsgBox szMsg & vbCrLf & "您必须进行座位设置!", vbInformation, Me.Caption
''            Dim nLen As Integer
''            Dim tv() As TVehcileSeatType
''            Dim oBase As New BaseInfo
''            oBase.Init g_oActiveUser
''
''            tv = oBase.GetAllVehicleSeatTypeInfo(mszVehicleId)
''            nLen = ArrayLength(tv)
''
''            Set oBase = Nothing
''
''            If nLen <> 0 Then
''               cmdSetSeatType_Click
''            End If
''
''        End If
''
''        frmBaseInfo.UpdateItemToList aszInfo
''        Unload Me
''        Exit Sub
''    End If
''    If Status = EFormStatus.EFS_AddNew Then
''        frmBaseInfo.AddItemToList aszInfo, True
''        RefresheVehicle
''        txtVehicleID.SetFocus
''    End If
    Unload Me
    Exit Sub
ErrHandle:
    Status = EFS_AddNew
    ShowErrorMsg
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandle
'    '布置窗体
'    AlignFormPos Me
    Set moVehicle = CreateObject("STBase.Vehicle")
    moVehicle.Init g_oActiveUser
    
    Select Case Status
        Case EFormStatus.EFS_AddNew
           cmdOK.Caption = "新增(&A)"
            RefresheVehicle
        Case EFormStatus.EFS_Modify
           txtVehicleID.Enabled = False
           RefresheVehicle
        Case EFormStatus.EFS_Show
            cmdCancel.TabIndex = 0
            cmdCancel.Default = True
            lblCaption.Caption = "车辆信息:"
            cmdOK.Visible = False
            txtAnnotation.Locked = True
            txtCardId.Locked = True
            txtCompany.Locked = True
            txtEndSeatNo.Locked = True
            txtLicense.Locked = True
            txtOwner.Locked = True
            txtSplitCompanyID.Locked = True
            txtStartSeatNo.Locked = True
            txtVehicleID.Locked = True
            txtVehicleModel.Locked = True
            RefresheVehicle
    End Select
    cmdOK.Enabled = False
    
    Exit Sub
ErrHandle:
    Status = EFS_AddNew
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moVehicle = Nothing
'    SaveFormPos Me

End Sub

Private Sub RefresheVehicle()
    If Status = EFS_AddNew Then
        txtVehicleID.Text = ""
        txtLicense.Text = ""
'        txtCompany.Text = ""
'        txtSplitCompanyID.Text = ""
        txtOwner.Text = ""
        txtCardId.Text = ""
'        txtVehicleModel.Text = ""
        txtAnnotation.Text = ""
        
        optRun.Value = True
    Else
        txtVehicleID.Text = mszVehicleId
        moVehicle.Identify mszVehicleId
        txtEndSeatNo.Text = Val(moVehicle.StartSeatNo) + Val(moVehicle.SeatCount) - 1
        txtStartSeatNo.Text = Format(moVehicle.StartSeatNo)
        txtLicense.Text = moVehicle.LicenseTag
        txtCompany.Text = MakeDisplayString(moVehicle.Company, moVehicle.CompanyName)
        txtSplitCompanyID.Text = MakeDisplayString(moVehicle.SplitCompanyID, moVehicle.SplitCompanyName)
        
        txtOwner.Text = MakeDisplayString(moVehicle.Owner, moVehicle.OwnerName)
        txtCardId.Text = moVehicle.CardID
        txtVehicleModel.Text = MakeDisplayString(moVehicle.VehicleModel, moVehicle.VehicleModelName)
        txtAnnotation.Text = moVehicle.Annotation
        
        mszOldStartSeat = Trim(txtStartSeatNo.Text)
        mszOldEndSeat = Trim(txtEndSeatNo.Text)
        
        If moVehicle.Status = ST_VehicleRun Then
            optRun.Value = True
        Else
            optStop.Value = True
        End If
    End If
End Sub

Private Sub optRun_Click()
    IsSave
End Sub

Private Sub optStop_Click()
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

Private Sub txtCardId_Change()
    IsSave
    FormatTextBoxBySize txtCardId, 10
End Sub

Private Sub txtCompany_Change()
    IsSave
End Sub

Private Sub txtCompany_ButtonClick()
On Error GoTo ErrHandle
    If txtCompany.Locked Then Exit Sub
    Dim aszTmp() As String
    Dim oTmp As New CommDialog
    oTmp.Init g_oActiveUser
    aszTmp = oTmp.SelectCompany
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtCompany.Text = MakeDisplayString(Trim(aszTmp(1, 1)), Trim(aszTmp(1, 2)))
    txtSplitCompanyID.Text = txtCompany.Text
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub


Private Sub txtCompany_GotFocus()
    With txtCompany
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtEndSeatNo_Change()
    IsSave
End Sub

Private Sub txtLicense_Change()
    IsSave
    FormatTextBoxBySize txtLicense, 10
End Sub

Private Sub txtOwner_Change()
    IsSave
End Sub
Private Sub txtOwner_ButtonClick()
On Error GoTo ErrHandle
    If txtOwner.Locked Then Exit Sub
    Dim aszTmp() As String
    Dim oTmp As New CommDialog
    oTmp.Init g_oActiveUser
    aszTmp = oTmp.SelectOwner
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtOwner.Text = MakeDisplayString(Trim(aszTmp(1, 1)), Trim(aszTmp(1, 2)))
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
Private Sub txtSplitCompanyID_Change()
    IsSave
End Sub

Private Sub txtSplitCompanyID_ButtonClick()
On Error GoTo ErrHandle
    If txtSplitCompanyID.Locked Then Exit Sub
    Dim aszTmp() As String
    Dim oTmp As New CommDialog
    oTmp.Init g_oActiveUser
    aszTmp = oTmp.SelectCompany
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtSplitCompanyID.Text = MakeDisplayString(Trim(aszTmp(1, 1)), Trim(aszTmp(1, 2)))
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub txtSplitCompanyID_GotFocus()
    With txtSplitCompanyID
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtStartSeatNo_Change()
    IsSave
End Sub

Private Sub txtVehicleID_Change()
    IsSave
    FormatTextBoxBySize txtVehicleID, 5
End Sub


Private Sub txtVehicleModel_Change()
    IsSave
End Sub

Private Sub txtVehicleModel_ButtonClick()
On Error GoTo ErrHandle
    If txtVehicleModel.Locked Then Exit Sub
    Dim aszTmp() As String
    Dim oTmp As New CommDialog
    oTmp.Init g_oActiveUser
    aszTmp = oTmp.SelectVehicleType
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtVehicleModel.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))
    SetSeatInfo
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub txtVehicleModel_GotFocus()
    txtVehicleModel.SelStart = 0
    txtVehicleModel.SelLength = Len(txtVehicleModel.Text)
End Sub

Private Sub txtVehicleModel_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
           SetSeatInfo
    End Select
End Sub

Private Sub SetSeatInfo()
On Error GoTo ErrHandle
    Dim oVehicleModel As New VehicleModel
    oVehicleModel.Init g_oActiveUser
    oVehicleModel.Identify ResolveDisplay(txtVehicleModel.Text)
    txtStartSeatNo.Text = oVehicleModel.StartSeatNumber
    txtEndSeatNo.Text = oVehicleModel.StartSeatNumber + oVehicleModel.SeatCount - 1
    Set oVehicleModel = Nothing
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub IsSave()
    If txtVehicleID.Text = "" Or txtCompany.Text = "" Or txtOwner.Text = "" _
        Or txtVehicleModel.Text = "" Or txtLicense.Text = "" Or Val(txtStartSeatNo.Text) = 0 _
        Or Val(txtEndSeatNo.Text) = 0 Or txtSplitCompanyID.Text = "" Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
End Sub

