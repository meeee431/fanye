VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmExPrintSheet 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "补打路单"
   ClientHeight    =   4995
   ClientLeft      =   3105
   ClientTop       =   3405
   ClientWidth     =   6555
   HelpContextID   =   4000601
   Icon            =   "frmExPrintSheet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Tag             =   "Modal"
   Begin RTComctl3.FlatLabel lblCurSheetNo 
      Height          =   285
      Left            =   4530
      TabIndex        =   29
      Top             =   1950
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OutnerStyle     =   2
      Caption         =   "01234556"
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   -30
      ScaleHeight     =   795
      ScaleWidth      =   7185
      TabIndex        =   25
      Top             =   0
      Width           =   7185
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   0
         TabIndex        =   26
         Top             =   750
         Width           =   7215
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请输入需要补打的路单编号:"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   300
         Width           =   2250
      End
   End
   Begin VB.TextBox txtOriSheetNo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1785
      TabIndex        =   1
      Text            =   "0000001"
      Top             =   1950
      Width           =   1410
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "路单摘要信息"
      Height          =   1845
      Left            =   480
      TabIndex        =   8
      Top             =   2280
      Width           =   5535
      Begin RTComctl3.CoolButton lblBusCheckInfo 
         Height          =   285
         Left            =   2880
         TabIndex        =   2
         Top             =   1470
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   503
         BTYPE           =   7
         TX              =   "车次检票信息(&C)"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   -2147483640
         FCOLO           =   -2147483640
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmExPrintSheet.frx":0A02
         PICN            =   "frmExPrintSheet.frx":0A1E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblMakeTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1999-9-8 12:30:23"
         Height          =   180
         Left            =   3720
         TabIndex        =   24
         Top             =   1200
         Width           =   1530
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "制单时间:"
         Height          =   180
         Left            =   2880
         TabIndex        =   23
         Top             =   1200
         Width           =   810
      End
      Begin VB.Label lblCheckor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "001/啊啊啊啊"
         Height          =   180
         Left            =   1140
         TabIndex        =   22
         Top             =   1200
         Width           =   1080
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票员:"
         Height          =   180
         Left            =   240
         TabIndex        =   21
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label lblRoute 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "重庆成都线"
         Height          =   180
         Left            =   3720
         TabIndex        =   20
         Top             =   660
         Width           =   900
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "线路:"
         Height          =   180
         Left            =   2880
         TabIndex        =   19
         Top             =   660
         Width           =   450
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "绍兴汽运集团"
         Height          =   180
         Left            =   3720
         TabIndex        =   18
         Top             =   930
         Width           =   1080
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参营公司:"
         Height          =   180
         Left            =   2880
         TabIndex        =   17
         Top             =   930
         Width           =   810
      End
      Begin VB.Label lblLicense 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "川A 1234567"
         Height          =   180
         Left            =   1140
         TabIndex        =   16
         Top             =   930
         Width           =   990
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "运行车辆:"
         Height          =   180
         Left            =   240
         TabIndex        =   15
         Top             =   930
         Width           =   810
      End
      Begin VB.Label lblBusSerial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         Height          =   180
         Left            =   3720
         TabIndex        =   14
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次序号:"
         Height          =   180
         Left            =   2880
         TabIndex        =   13
         Top             =   360
         Width           =   810
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1999-06-08"
         Height          =   180
         Left            =   1140
         TabIndex        =   12
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "日期:"
         Height          =   180
         Left            =   240
         TabIndex        =   11
         Top             =   630
         Width           =   450
      End
      Begin VB.Label lblBusID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1234"
         Height          =   180
         Left            =   1140
         TabIndex        =   10
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次代码:"
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "注意"
      Height          =   990
      Left            =   480
      TabIndex        =   6
      Top             =   870
      Width           =   5535
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "  补打路单不生成新的路单号，不作废原路单！"
         Height          =   555
         Left            =   960
         TabIndex        =   30
         Top             =   270
         Width           =   4470
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   225
         Picture         =   "frmExPrintSheet.frx":0B78
         Top             =   285
         Width           =   480
      End
   End
   Begin RTComctl3.CoolButton cmdPreView 
      Height          =   345
      Left            =   1770
      TabIndex        =   3
      Top             =   4500
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "路单预览(&V)"
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
      MICON           =   "frmExPrintSheet.frx":1442
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdExit 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   4830
      TabIndex        =   5
      Top             =   4500
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
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
      MICON           =   "frmExPrintSheet.frx":145E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdExPrintSheet 
      Height          =   345
      Left            =   3300
      TabIndex        =   4
      Top             =   4500
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "补打路单(&P)"
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
      MICON           =   "frmExPrintSheet.frx":147A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   990
      Left            =   -90
      TabIndex        =   28
      Top             =   4230
      Width           =   8745
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "当前路单编号:"
      Height          =   180
      Left            =   3315
      TabIndex        =   7
      Top             =   1995
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "原路单编号(&N):"
      Height          =   180
      Left            =   495
      TabIndex        =   0
      Top             =   1995
      Width           =   1260
   End
End
Attribute VB_Name = "frmExPrintSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const msgTitle = "补打路单"
Private mtOriSheetInfo As TCheckSheetInfo       '原始路单信息


Private Function ValidateSheetId() As Boolean
'    If txtOriSheetNo.Text <> "" Then
        If txtOriSheetNo.Text <> lblCurSheetNo.Caption Then
            mtOriSheetInfo = g_oChkTicket.GetCheckSheetInfo(txtOriSheetNo.Text)
            If mtOriSheetInfo.szCheckSheet = "" Then
                MsgboxEx "该路单不存在！", vbCritical, "错误"
                ValidateSheetId = False
            Else
                writeSheetInfo
                ValidateSheetId = True
            End If
        End If
 '   End If
End Function

Private Sub writeSheetInfo()
    Dim oVehicle As New Vehicle
    Dim oRoute As New Route
    On Error GoTo ErrorHandle
    oVehicle.Init g_oActiveUser
    oVehicle.Identify mtOriSheetInfo.szVehicleId
    oRoute.Init g_oActiveUser
    oRoute.Identify mtOriSheetInfo.szRouteID
    
    lblBusID.Caption = mtOriSheetInfo.szBusid
    lblBusSerial.Caption = mtOriSheetInfo.nBusSerialNo
    lblCheckor.Caption = mtOriSheetInfo.szMakeSheetUser
    lblCompany.Caption = oVehicle.Company
    lblDate.Caption = Format(mtOriSheetInfo.dtDate, "YYYY-MM-DD")
    lblLicense.Caption = oVehicle.LicenseTag
    lblMakeTime.Caption = Format(mtOriSheetInfo.dtMakeSheetDateTime, "YYYY-MM-DD HH:MM:SS")
    lblRoute.Caption = oRoute.RouteName
    Set oVehicle = Nothing
    Set oRoute = Nothing
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub InitForm()
    lblBusID.Caption = ""
    lblBusSerial.Caption = ""
    lblCheckor.Caption = ""
    lblCompany.Caption = ""
    lblDate.Caption = ""
    lblLicense.Caption = ""
    lblMakeTime.Caption = ""
    lblRoute.Caption = ""
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    Dim szSheetID As String
    szSheetID = Trim(txtOriSheetNo.Text)
    If szSheetID = "" Then Exit Sub
    Dim ofrmTmp As frmCheckSheet
    Set ofrmTmp = New frmCheckSheet
    Set ofrmTmp.g_oActiveUser = g_oActiveUser
    Set ofrmTmp.moChkTicket = g_oChkTicket
    ofrmTmp.mbViewMode = True
    ofrmTmp.mbNoPrintPrompt = True
    ofrmTmp.mbExitAfterPrint = False
    ofrmTmp.mszSheetID = szSheetID
    ofrmTmp.Show vbModal
End Sub

Private Sub cmdExPrintSheet_Click()
    On Error GoTo ErrHandle
    
    
    If Not ValidateSheetId Then
        txtOriSheetNo.SetFocus
        Exit Sub
    End If
    If txtOriSheetNo.Text = "" Then
        MsgboxEx "请输入需补打的路单号", vbInformation, "补打路单"
        txtOriSheetNo.SetFocus
        Exit Sub
    End If
    If lblCurSheetNo.Caption <> txtOriSheetNo.Text Then
        If MsgboxEx("请确定需要补打路单吗？", vbExclamation + vbYesNo, msgTitle) = vbYes Then
        
            If Trim(mtOriSheetInfo.szSettleSheetID) <> "" Then
                If MsgboxEx("路单已经结算，要继续补打吗？", vbExclamation + vbYesNo + vbDefaultButton2, msgTitle) = vbNo Then Exit Sub
            End If
            
            '显示路单并打印
            Dim ofrmTmp As frmCheckSheet
            Set ofrmTmp = New frmCheckSheet
            Set ofrmTmp.g_oActiveUser = g_oActiveUser
            Set ofrmTmp.moChkTicket = g_oChkTicket
            ofrmTmp.mbViewMode = False
            ofrmTmp.mbNoPrintPrompt = True
            ofrmTmp.mbExitAfterPrint = False
            ofrmTmp.mszSheetID = txtOriSheetNo.Text
            ofrmTmp.Show vbModal
            
            g_tCheckInfo.CurrSheetNo = NumAdd(lblCurSheetNo.Caption, 1)
            MDIMain.lblCurrentSheetNo.Caption = g_tCheckInfo.CurrSheetNo
            Unload Me
        End If
        '如果路单号一致,则不重新生成路单
        '不一致，则首先作废路单，然后重新生成路单
        '调用路单显示窗体
        '
    End If
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
    txtOriSheetNo.SetFocus
End Sub

Private Sub Form_Activate()
    lblCurSheetNo.Caption = g_tCheckInfo.CurrSheetNo
End Sub

Private Sub Form_Load()
    lblCurSheetNo.Caption = g_tCheckInfo.CurrSheetNo
    lblBusID.Caption = ""
    lblBusSerial.Caption = ""
    lblCheckor.Caption = ""
    lblCompany.Caption = ""
    lblDate.Caption = ""
    lblLicense.Caption = ""
    lblMakeTime.Caption = ""
    lblRoute.Caption = ""
    txtOriSheetNo.Text = ""
End Sub

Private Sub Form_Resize()
    txtOriSheetNo.SetFocus
End Sub

Private Sub lblBusCheckInfo_Click()
    If lblBusID.Caption <> "" Then
        Dim oFrmCheckInfo As New frmCheckBusInfo
        Set oFrmCheckInfo.g_oActiveUser = g_oActiveUser
        oFrmCheckInfo.mszBusID = lblBusID
        oFrmCheckInfo.mnBusSerialNo = Val(lblBusSerial.Caption)
        oFrmCheckInfo.mdtBusDate = lblDate.Caption
        oFrmCheckInfo.Show vbModal
        Set oFrmCheckInfo = Nothing
    End If
End Sub

Private Sub txtOriSheetNo_Change()
    If txtOriSheetNo.Text <> "" Then
        If Len(txtOriSheetNo.Text) > 10 Then
            txtOriSheetNo.Text = Left(txtOriSheetNo.Text, 10)
        End If
    End If
End Sub

Private Sub txtOriSheetNo_GotFocus()
    txtOriSheetNo.SelStart = 0
    txtOriSheetNo.SelLength = Len(txtOriSheetNo.Text)
End Sub

Private Sub txtOriSheetNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not ValidateSheetId Then
            txtOriSheetNo.SetFocus
        Else
            cmdExPrintSheet.SetFocus
        End If
    End If
    If KeyAscii = 32 Then
        KeyAscii = 0
    End If
End Sub

Private Sub DisplayHelp(Optional HelpType As EHelpType = content)
    Dim lActiveControl As Long
    
    Select Case HelpType
        Case content
            lActiveControl = Me.ActiveControl.HelpContextID
            If lActiveControl = 0 Then
                TopicID = Me.HelpContextID
                CallHTMLShowTopicID
            Else
                TopicID = lActiveControl
                CallHTMLShowTopicID
            End If
        Case Index
            CallHTMLHelpIndex
        Case Support
            TopicID = clSupportID
            CallHTMLShowTopicID
    End Select

End Sub

