VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmRePrintSheet 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "重打签发单"
   ClientHeight    =   5430
   ClientLeft      =   2565
   ClientTop       =   3510
   ClientWidth     =   7140
   HelpContextID   =   7000070
   Icon            =   "frmRePrintSheet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtOldSheetNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   1830
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "注意"
      Height          =   870
      Left            =   390
      TabIndex        =   47
      Top             =   870
      Width           =   6405
      Begin VB.Label label2 
         BackStyle       =   0  'Transparent
         Caption         =   "  重打签发单将生成新的行包签发单编号，以便与打印机的当前路单编号一致，请在签发单打印错误时才使用此功能，正常时请勿使用。"
         Height          =   405
         Index           =   1
         Left            =   840
         TabIndex        =   0
         Top             =   330
         Width           =   5430
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   165
         Picture         =   "frmRePrintSheet.frx":038A
         Top             =   255
         Width           =   480
      End
   End
   Begin RTComctl3.FlatLabel lblCurSheetNo 
      Height          =   285
      Left            =   4920
      TabIndex        =   46
      Top             =   1800
      Width           =   1395
      _ExtentX        =   2461
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
   Begin RTComctl3.CoolButton cmdExit 
      Height          =   345
      Left            =   5400
      TabIndex        =   7
      Top             =   4890
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
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
      MICON           =   "frmRePrintSheet.frx":0C54
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdRePrintSheet 
      Height          =   345
      Left            =   3780
      TabIndex        =   8
      Top             =   4890
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "重新打印(&P)"
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
      MICON           =   "frmRePrintSheet.frx":0C70
      PICN            =   "frmRePrintSheet.frx":0C8C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   -30
      ScaleHeight     =   795
      ScaleWidth      =   7185
      TabIndex        =   4
      Top             =   0
      Width           =   7185
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   0
         TabIndex        =   5
         Top             =   750
         Width           =   7215
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请输入要重打的签发单号:"
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   2070
      End
      Begin VB.Image Image2 
         Height          =   855
         Left            =   1650
         Picture         =   "frmRePrintSheet.frx":1026
         Top             =   0
         Width           =   5925
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   " RTStation"
      Height          =   990
      Left            =   -120
      TabIndex        =   9
      Top             =   4650
      Width           =   8745
      Begin RTComctl3.CoolButton cmdHelp 
         Cancel          =   -1  'True
         Height          =   345
         Left            =   480
         TabIndex        =   48
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "帮助(&H)"
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
         MICON           =   "frmRePrintSheet.frx":2510
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   2430
      Left            =   390
      TabIndex        =   10
      Top             =   2130
      Width           =   6405
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "220.3"
         Height          =   180
         Left            =   5310
         TabIndex        =   45
         Top             =   2070
         Width           =   450
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总运价:"
         Height          =   180
         Left            =   4560
         TabIndex        =   44
         Top             =   2070
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "超重件数:"
         Height          =   180
         Left            =   4560
         TabIndex        =   43
         Top             =   1770
         Width           =   810
      End
      Begin VB.Label lblOverNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         Height          =   180
         Left            =   5385
         TabIndex        =   42
         Tag             =   "提货人"
         Top             =   1770
         Width           =   180
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总计重:"
         Height          =   180
         Left            =   840
         TabIndex        =   41
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label lblTotalCatWeight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "20.5公斤"
         Height          =   180
         Left            =   1500
         TabIndex        =   40
         Tag             =   "提货人"
         Top             =   2040
         Width           =   720
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000005&
         Index           =   0
         X1              =   840
         X2              =   6120
         Y1              =   1635
         Y2              =   1635
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000003&
         Index           =   0
         X1              =   840
         X2              =   6120
         Y1              =   1620
         Y2              =   1620
      End
      Begin VB.Label lblBusDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2002-12-11"
         Height          =   180
         Left            =   3540
         TabIndex        =   39
         Top             =   540
         Width           =   900
      End
      Begin VB.Label lblVehicle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "浙D00001"
         Height          =   180
         Left            =   1680
         TabIndex        =   38
         Top             =   800
         Width           =   720
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "浙江快客"
         Height          =   180
         Left            =   3540
         TabIndex        =   37
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label lblStateChange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次日期:"
         Height          =   180
         Left            =   2700
         TabIndex        =   36
         Top             =   540
         Width           =   810
      End
      Begin VB.Label label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参运公司:"
         Height          =   180
         Index           =   0
         Left            =   2700
         TabIndex        =   35
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "承运车辆:"
         Height          =   180
         Index           =   2
         Left            =   840
         TabIndex        =   34
         Top             =   810
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "运行线路:"
         Height          =   180
         Left            =   840
         TabIndex        =   33
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label lblRoute 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "杭州温州线"
         Height          =   180
         Left            =   1680
         TabIndex        =   32
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "托运单数:"
         Height          =   180
         Left            =   840
         TabIndex        =   31
         Top             =   1770
         Width           =   810
      End
      Begin VB.Label lblShippNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         Height          =   180
         Left            =   1680
         TabIndex        =   30
         Top             =   1800
         Width           =   180
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总件数:"
         Height          =   180
         Left            =   2700
         TabIndex        =   29
         Top             =   1770
         Width           =   630
      End
      Begin VB.Label lblNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         Height          =   180
         Left            =   3345
         TabIndex        =   28
         Tag             =   "提货人"
         Top             =   1770
         Width           =   180
      End
      Begin VB.Label lblOffTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "流水班次"
         Height          =   180
         Left            =   5400
         TabIndex        =   27
         Top             =   540
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发车时间:"
         Height          =   180
         Left            =   4560
         TabIndex        =   26
         Top             =   540
         Width           =   810
      End
      Begin VB.Label lblDesCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "浙江快客公司"
         Height          =   180
         Left            =   3540
         TabIndex        =   25
         Top             =   1350
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆帐公司:"
         Height          =   180
         Left            =   2700
         TabIndex        =   24
         Top             =   1350
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "签发车次:"
         Height          =   180
         Left            =   840
         TabIndex        =   23
         Top             =   540
         Width           =   810
      End
      Begin VB.Label lblCarryVehicle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2002"
         Height          =   180
         Left            =   1680
         TabIndex        =   22
         Top             =   540
         Width           =   360
      End
      Begin VB.Label lblSheetID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "U000001"
         Height          =   195
         Left            =   1845
         TabIndex        =   21
         Top             =   270
         Width           =   630
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "签发单代码:"
         Height          =   180
         Index           =   1
         Left            =   840
         TabIndex        =   20
         Top             =   270
         Width           =   990
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   195
         Picture         =   "frmRePrintSheet.frx":252C
         Top             =   210
         Width           =   480
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "承运车型:"
         Height          =   180
         Index           =   3
         Left            =   2700
         TabIndex        =   19
         Top             =   810
         Width           =   810
      End
      Begin VB.Label lblVehicleType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "大客"
         Height          =   180
         Left            =   3540
         TabIndex        =   18
         Top             =   810
         Width           =   360
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车主:"
         Height          =   180
         Index           =   4
         Left            =   4560
         TabIndex        =   17
         Top             =   810
         Width           =   450
      End
      Begin VB.Label lblOwner 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "陆永庆"
         Height          =   180
         Left            =   5040
         TabIndex        =   16
         Top             =   810
         Width           =   540
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "作废"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   4560
         TabIndex        =   15
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆算协议:"
         Height          =   180
         Left            =   840
         TabIndex        =   14
         Top             =   1350
         Width           =   810
      End
      Begin VB.Label lblProtocol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5：5"
         Height          =   180
         Left            =   1680
         TabIndex        =   13
         Top             =   1350
         Width           =   360
      End
      Begin VB.Label lblTotalActWeight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "20.5公斤"
         Height          =   180
         Left            =   3360
         TabIndex        =   12
         Tag             =   "提货人"
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总实重:"
         Height          =   180
         Left            =   2700
         TabIndex        =   11
         Top             =   2040
         Width           =   630
      End
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "原签发单号(&N):"
      Height          =   180
      Index           =   0
      Left            =   855
      TabIndex        =   3
      Top             =   1875
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "当前签发单号:"
      Height          =   180
      Left            =   3720
      TabIndex        =   2
      Top             =   1875
      Width           =   1170
   End
End
Attribute VB_Name = "frmRePrintSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub


Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdRePrintSheet_Click()
  On Error GoTo ErrHandle
     Dim nAnswer
     
     nAnswer = MsgBox("您是否确认打印?", vbInformation + vbYesNo, Me.Caption)
     If nAnswer = vbYes Then
        moLugSvr.ReprintCarrySheet Trim(txtOldSheetNo.Text), Trim(lblCurSheetNo.Caption)
        moCarrySheet.SheetID = lblCurSheetNo.Caption  '更新对象中的签发单号
        'BillPrint
    ShowSBInfo "正在重打签发单..."
    PrintCarrySheet moCarrySheet
    ShowSBInfo ""

'        g_szCarrySheetID = g_szCarrySheetID + 1
        IncSheetID
        lblCurSheetNo.Caption = Trim(mdiMain.lblSheetNo)
        FormClear
        txtOldSheetNo.SetFocus
     End If
     
  Exit Sub
ErrHandle:
  ShowErrorMsg
End Sub

Private Sub Form_Activate()
    SetSheetNoLabel False, g_szCarrySheetID
End Sub

Private Sub Form_Deactivate()
    HideSheetNoLabel
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
    If KeyAscii = vbKeyF1 Then
        DisplayHelp Me
    End If
End Sub

Private Sub Form_Load()
 AlignFormPos Me
 
 FormClear
 lblCurSheetNo.Caption = FormatSheetID(g_szCarrySheetID)
 cmdRePrintSheet.Enabled = False
End Sub

Private Sub FormClear()
  txtOldSheetNo.Text = ""
  lblVehicle.Caption = ""
  lblSheetID.Caption = ""
  lblCarryVehicle.Caption = ""
  lblBusDate.Caption = ""
  lblOffTime.Caption = ""
  lblCarryVehicle.Caption = ""
  lblVehicleType.Caption = ""
  lblOwner.Caption = ""
  lblRoute.Caption = ""
  lblCompany.Caption = ""
  lblProtocol.Caption = ""
  lblDesCompany.Caption = ""
  lblShippNumber.Caption = ""
  lblNumber.Caption = ""
  lblOverNumber.Caption = ""
  lblTotalCatWeight.Caption = ""
  lblTotalActWeight.Caption = ""
  lblTotalPrice.Caption = ""
  lblStatus.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    HideSheetNoLabel
    SaveFormPos Me
End Sub

Private Sub lblTotalWeight_Click()

End Sub



Private Sub txtOldSheetNo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHandle
  If KeyCode = 13 Then
     If txtOldSheetNo.Text = "" Then
        MsgBox "签发单号不能为空!", vbExclamation, Me.Caption
        Exit Sub
     Else
        FillForm
        cmdRePrintSheet.Enabled = True
        cmdRePrintSheet.SetFocus
     End If
  End If
  
Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub
Private Sub FillForm()
    On Error GoTo ErrHandle
    moCarrySheet.Identify Trim(txtOldSheetNo.Text)
    lblSheetID.Caption = Trim(moCarrySheet.SheetID)
    lblCarryVehicle.Caption = Trim(moCarrySheet.VehicleID)
    lblBusDate.Caption = CStr(moCarrySheet.BusDate)
    lblOffTime.Caption = CStr(Format(moCarrySheet.BusStartOffTime, "hh:mm"))
    lblVehicle.Caption = Trim(moCarrySheet.VehicleLicense)
    lblVehicleType.Caption = Trim(moCarrySheet.VehicleTypeName)
    lblOwner.Caption = Trim(moCarrySheet.BusOwnerName)
    lblRoute.Caption = Trim(moCarrySheet.RouteName)
    lblCompany.Caption = Trim(moCarrySheet.CompanyName)
    lblProtocol.Caption = Trim(moCarrySheet.ProtocolName)
    lblDesCompany.Caption = Trim(moCarrySheet.SplitCompanyName)
    lblShippNumber.Caption = CStr(moCarrySheet.AcceptSheetNumber)  '托运单数
    lblNumber.Caption = CStr(moCarrySheet.Number)
    lblOverNumber.Caption = CStr(moCarrySheet.OverNumber)
    lblTotalCatWeight.Caption = CStr(moCarrySheet.CalWeight)
    lblTotalActWeight.Caption = CStr(moCarrySheet.ActWeight)
    lblTotalPrice.Caption = CStr(moCarrySheet.TotalPrice)
'    If moCarrySheet.Status <> 0 Then
      lblStatus.ForeColor = vbRed
'    Else
'      lblStatus.ForeColor = 0
'    End If
    lblStatus.Caption = CStr(moCarrySheet.StatusString)
    Exit Sub
ErrHandle:
ShowErrorMsg
End Sub
