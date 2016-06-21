VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmCancelSheet 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "作废签发单"
   ClientHeight    =   4515
   ClientLeft      =   2580
   ClientTop       =   3045
   ClientWidth     =   7110
   HelpContextID   =   7000080
   Icon            =   "frmCancelSheet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   7110
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   2430
      Left            =   390
      TabIndex        =   7
      Top             =   1260
      Width           =   6405
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总实重:"
         Height          =   180
         Left            =   2700
         TabIndex        =   42
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label lblTotalActWeight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3360
         TabIndex        =   41
         Tag             =   "提货人"
         Top             =   2040
         Width           =   90
      End
      Begin VB.Label lblProtocol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1680
         TabIndex        =   40
         Top             =   1350
         Width           =   90
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆算协议:"
         Height          =   180
         Left            =   840
         TabIndex        =   39
         Top             =   1350
         Width           =   810
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   4560
         TabIndex        =   38
         Top             =   270
         Width           =   90
      End
      Begin VB.Label lblOwner 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   5040
         TabIndex        =   37
         Top             =   810
         Width           =   90
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车主:"
         Height          =   180
         Index           =   4
         Left            =   4560
         TabIndex        =   36
         Top             =   810
         Width           =   450
      End
      Begin VB.Label lblVehicleType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3540
         TabIndex        =   35
         Top             =   810
         Width           =   90
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "承运车型:"
         Height          =   180
         Index           =   3
         Left            =   2700
         TabIndex        =   34
         Top             =   810
         Width           =   810
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   195
         Picture         =   "frmCancelSheet.frx":030A
         Top             =   210
         Width           =   480
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "签发单代码:"
         Height          =   180
         Index           =   1
         Left            =   840
         TabIndex        =   33
         Top             =   270
         Width           =   990
      End
      Begin VB.Label lblSheetID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1845
         TabIndex        =   32
         Top             =   270
         Width           =   90
      End
      Begin VB.Label lblCarryVehicle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1680
         TabIndex        =   31
         Top             =   540
         Width           =   90
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "签发车次:"
         Height          =   180
         Left            =   840
         TabIndex        =   30
         Top             =   540
         Width           =   810
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆帐公司:"
         Height          =   180
         Left            =   2700
         TabIndex        =   29
         Top             =   1350
         Width           =   810
      End
      Begin VB.Label lblDesCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3540
         TabIndex        =   28
         Top             =   1350
         Width           =   90
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发车时间:"
         Height          =   180
         Left            =   4560
         TabIndex        =   27
         Top             =   540
         Width           =   810
      End
      Begin VB.Label lblOffTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   5400
         TabIndex        =   26
         Top             =   540
         Width           =   90
      End
      Begin VB.Label lblNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3345
         TabIndex        =   25
         Tag             =   "提货人"
         Top             =   1770
         Width           =   90
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总件数:"
         Height          =   180
         Left            =   2700
         TabIndex        =   24
         Top             =   1770
         Width           =   630
      End
      Begin VB.Label lblShippNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1680
         TabIndex        =   23
         Top             =   1800
         Width           =   90
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "托运单数:"
         Height          =   180
         Left            =   840
         TabIndex        =   22
         Top             =   1770
         Width           =   810
      End
      Begin VB.Label lblRoute 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1680
         TabIndex        =   21
         Top             =   1080
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "运行线路:"
         Height          =   180
         Left            =   840
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   810
         Width           =   810
      End
      Begin VB.Label label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参运公司:"
         Height          =   180
         Index           =   0
         Left            =   2700
         TabIndex        =   18
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label lblStateChange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次日期:"
         Height          =   180
         Left            =   2700
         TabIndex        =   17
         Top             =   540
         Width           =   810
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3540
         TabIndex        =   16
         Top             =   1080
         Width           =   90
      End
      Begin VB.Label lblVehicle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1680
         TabIndex        =   15
         Top             =   795
         Width           =   90
      End
      Begin VB.Label lblBusDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3540
         TabIndex        =   14
         Top             =   540
         Width           =   90
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000003&
         Index           =   0
         X1              =   840
         X2              =   6120
         Y1              =   1620
         Y2              =   1620
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000005&
         Index           =   0
         X1              =   840
         X2              =   6120
         Y1              =   1635
         Y2              =   1635
      End
      Begin VB.Label lblTotalCatWeight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1500
         TabIndex        =   13
         Tag             =   "提货人"
         Top             =   2040
         Width           =   90
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总计重:"
         Height          =   180
         Left            =   840
         TabIndex        =   12
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label lblOverNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   5385
         TabIndex        =   11
         Tag             =   "提货人"
         Top             =   1770
         Width           =   90
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "超重件数:"
         Height          =   180
         Left            =   4560
         TabIndex        =   10
         Top             =   1770
         Width           =   810
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总运价:"
         Height          =   180
         Left            =   4560
         TabIndex        =   9
         Top             =   2070
         Width           =   630
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   5670
         TabIndex        =   8
         Top             =   2070
         Width           =   90
      End
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   -30
      ScaleHeight     =   795
      ScaleWidth      =   7185
      TabIndex        =   3
      Top             =   0
      Width           =   7185
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   0
         TabIndex        =   4
         Top             =   750
         Width           =   7215
      End
      Begin VB.Image Image2 
         Height          =   855
         Left            =   2310
         Picture         =   "frmCancelSheet.frx":0BD4
         Top             =   0
         Width           =   5925
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请输入要作废的签发单号:"
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   2070
      End
   End
   Begin VB.TextBox txtOldSheetNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   960
      Width           =   1395
   End
   Begin RTComctl3.CoolButton cmdExit 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   5640
      TabIndex        =   1
      Top             =   4020
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "frmCancelSheet.frx":20BE
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
      Left            =   4350
      TabIndex        =   2
      Top             =   4020
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "作废(&P)"
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
      MICON           =   "frmCancelSheet.frx":20DA
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
      Caption         =   " RTStation"
      Height          =   990
      Left            =   -120
      TabIndex        =   6
      Top             =   3750
      Width           =   8745
      Begin RTComctl3.CoolButton cmdHelp 
         Height          =   345
         Left            =   540
         TabIndex        =   44
         Top             =   300
         Width           =   1155
         _ExtentX        =   2037
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
         MICON           =   "frmCancelSheet.frx":20F6
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
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "原签发单号(&N):"
      Height          =   180
      Index           =   0
      Left            =   855
      TabIndex        =   43
      Top             =   1005
      Width           =   1260
   End
End
Attribute VB_Name = "frmCancelSheet"
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
     
     nAnswer = MsgBox("您是否确认作废?", vbInformation + vbYesNo, Me.Caption)
     If nAnswer = vbYes Then
        moLugSvr.CancelSheet Trim(txtOldSheetNo.Text)
        FillForm
     MsgBox "  作废成功!", vbInformation, Me.Caption
     moCarrySheet.AddNew
     frmCarryLuggage.RefreshLuggage
        'BillPrint
'    ShowSBInfo "正在重打签发单..."
'    PrintCarrySheet moCarrySheet
'    ShowSBInfo ""

'        g_szCarrySheetID = g_szCarrySheetID + 1
'        IncSheetID
'        lblCurSheetNo.Caption = Trim(mdiMain.lblSheetNo)
'        FormClear
'        txtOldSheetNo.SetFocus
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
' lblCurSheetNo.Caption = g_szCarrySheetID
 cmdRePrintSheet.Enabled = False
End Sub

Private Sub FormClear()
'  txtOldSheetNo.Text = ""
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



Private Sub txtOldSheetNo_Change()
    If txtOldSheetNo.Text = "" Then
        cmdRePrintSheet.Enabled = False
    Else
        cmdRePrintSheet.Enabled = True
    End If
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
    FormClear
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

