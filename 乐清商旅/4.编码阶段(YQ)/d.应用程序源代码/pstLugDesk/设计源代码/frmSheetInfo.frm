VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmSheetInfo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "签发单"
   ClientHeight    =   6225
   ClientLeft      =   2460
   ClientTop       =   2850
   ClientWidth     =   7395
   Icon            =   "frmSheetInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2790
      Left            =   495
      TabIndex        =   7
      Top             =   810
      Width           =   6405
      Begin VB.Label lblAcceptType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "快件"
         Height          =   180
         Left            =   3570
         TabIndex        =   51
         Top             =   280
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "托运方式:"
         Height          =   180
         Left            =   2700
         TabIndex        =   50
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总实重:"
         Height          =   180
         Left            =   2700
         TabIndex        =   48
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label lblTotalActWeight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "20.5公斤"
         Height          =   180
         Left            =   3360
         TabIndex        =   47
         Tag             =   "提货人"
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label lblProtocol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5：5"
         Height          =   180
         Left            =   1680
         TabIndex        =   46
         Top             =   1350
         Width           =   360
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆算协议:"
         Height          =   180
         Left            =   840
         TabIndex        =   45
         Top             =   1350
         Width           =   810
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "作废"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   4920
         TabIndex        =   44
         Top             =   270
         Width           =   360
      End
      Begin VB.Label lblOwner 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "陆永庆"
         Height          =   180
         Left            =   5040
         TabIndex        =   43
         Top             =   810
         Width           =   540
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车主:"
         Height          =   180
         Index           =   4
         Left            =   4560
         TabIndex        =   42
         Top             =   810
         Width           =   450
      End
      Begin VB.Label lblVehicleType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "大客"
         Height          =   180
         Left            =   3540
         TabIndex        =   41
         Top             =   810
         Width           =   360
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "承运车型:"
         Height          =   180
         Index           =   3
         Left            =   2700
         TabIndex        =   40
         Top             =   810
         Width           =   810
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   195
         Picture         =   "frmSheetInfo.frx":038A
         Top             =   210
         Width           =   480
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "签发单代码:"
         Height          =   180
         Index           =   0
         Left            =   840
         TabIndex        =   35
         Top             =   270
         Width           =   990
      End
      Begin VB.Label lblSheetID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "U000001"
         Height          =   195
         Left            =   1845
         TabIndex        =   34
         Top             =   270
         Width           =   630
      End
      Begin VB.Label lblCarryBus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2002"
         Height          =   180
         Left            =   1680
         TabIndex        =   33
         Top             =   540
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "签发车次:"
         Height          =   180
         Left            =   840
         TabIndex        =   32
         Top             =   540
         Width           =   810
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆帐公司:"
         Height          =   180
         Left            =   2700
         TabIndex        =   31
         Top             =   1350
         Width           =   810
      End
      Begin VB.Label lblDesCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "浙江快客公司"
         Height          =   180
         Left            =   3540
         TabIndex        =   30
         Top             =   1350
         Width           =   1080
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发车时间:"
         Height          =   180
         Left            =   4560
         TabIndex        =   29
         Top             =   540
         Width           =   810
      End
      Begin VB.Label lblOffBusTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "流水班次"
         Height          =   180
         Left            =   5400
         TabIndex        =   28
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lblTotalNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         Height          =   180
         Left            =   3345
         TabIndex        =   27
         Tag             =   "提货人"
         Top             =   1770
         Width           =   180
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总件数:"
         Height          =   180
         Left            =   2700
         TabIndex        =   26
         Top             =   1770
         Width           =   630
      End
      Begin VB.Label lblShipperNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         Height          =   180
         Left            =   1680
         TabIndex        =   25
         Top             =   1770
         Width           =   180
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "托运单数:"
         Height          =   180
         Left            =   840
         TabIndex        =   24
         Top             =   1770
         Width           =   810
      End
      Begin VB.Label lblRoute 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "杭州温州线"
         Height          =   180
         Left            =   1680
         TabIndex        =   23
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "运行线路:"
         Height          =   180
         Left            =   840
         TabIndex        =   22
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label lblTimeChange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "签发时间:"
         Height          =   180
         Left            =   2700
         TabIndex        =   21
         Top             =   2460
         Width           =   810
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "承运车辆:"
         Height          =   180
         Index           =   1
         Left            =   840
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label lblStateChange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次日期:"
         Height          =   180
         Left            =   2700
         TabIndex        =   18
         Top             =   540
         Width           =   810
      End
      Begin VB.Label lblOperatorChange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "操作员:"
         Height          =   180
         Left            =   840
         TabIndex        =   17
         Top             =   2460
         Width           =   630
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "浙江快客"
         Height          =   180
         Left            =   3540
         TabIndex        =   16
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label lblVehicle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "浙D00001"
         Height          =   180
         Left            =   1680
         TabIndex        =   15
         Top             =   810
         Width           =   720
      End
      Begin VB.Label lblOperater 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "陆永庆"
         Height          =   180
         Left            =   1500
         TabIndex        =   14
         Top             =   2460
         Width           =   540
      End
      Begin VB.Label lblOperationTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2002-12-18 9:00"
         Height          =   180
         Left            =   3540
         TabIndex        =   13
         Top             =   2460
         Width           =   1350
      End
      Begin VB.Label lblBusDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2002-12-11"
         Height          =   180
         Left            =   3540
         TabIndex        =   12
         Top             =   540
         Width           =   900
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
         Caption         =   "20.5公斤"
         Height          =   180
         Left            =   1500
         TabIndex        =   11
         Tag             =   "提货人"
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总计重:"
         Height          =   180
         Left            =   840
         TabIndex        =   10
         Top             =   2040
         Width           =   630
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000003&
         Index           =   1
         X1              =   840
         X2              =   6120
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000005&
         Index           =   1
         X1              =   840
         X2              =   6120
         Y1              =   2355
         Y2              =   2355
      End
      Begin VB.Label lblOverNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         Height          =   180
         Left            =   5385
         TabIndex        =   9
         Tag             =   "提货人"
         Top             =   1770
         Width           =   180
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "超重件数:"
         Height          =   180
         Left            =   4560
         TabIndex        =   8
         Top             =   1770
         Width           =   810
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1800
      Left            =   495
      TabIndex        =   3
      Top             =   3600
      Width           =   6405
      Begin VSFlex7LCtl.VSFlexGrid VSFGTotalPrice 
         Height          =   1080
         Left            =   870
         TabIndex        =   4
         Top             =   540
         Width           =   5295
         _cx             =   9340
         _cy             =   1905
         _ConvInfo       =   -1
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16777215
         ForeColorSel    =   -2147483640
         BackColorBkg    =   14737632
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   210
         Picture         =   "frmSheetInfo.frx":0C54
         Top             =   285
         Width           =   480
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总运价:"
         Height          =   180
         Left            =   870
         TabIndex        =   6
         Top             =   270
         Width           =   630
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "220.3"
         Height          =   180
         Left            =   1620
         TabIndex        =   5
         Top             =   270
         Width           =   450
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -30
      TabIndex        =   0
      Top             =   690
      Width           =   7815
   End
   Begin RTComctl3.CoolButton cmdOK 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   5385
      TabIndex        =   36
      Top             =   5760
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   609
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
      MICON           =   "frmSheetInfo.frx":0F5E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdPrint 
      Default         =   -1  'True
      Height          =   345
      Left            =   3765
      TabIndex        =   49
      Top             =   5760
      Visible         =   0   'False
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "作废重打(&P)"
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
      MICON           =   "frmSheetInfo.frx":0F7A
      PICN            =   "frmSheetInfo.frx":0F96
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
      Enabled         =   0   'False
      Height          =   3405
      Left            =   -105
      TabIndex        =   37
      Top             =   5445
      Width           =   8745
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   -30
      ScaleHeight     =   735
      ScaleWidth      =   7785
      TabIndex        =   1
      Top             =   0
      Width           =   7785
      Begin VB.Image Image3 
         Height          =   855
         Left            =   1950
         Picture         =   "frmSheetInfo.frx":1330
         Top             =   0
         Width           =   5925
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包签发单信息:"
         Height          =   180
         Left            =   270
         TabIndex        =   2
         Top             =   270
         Width           =   1350
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "浙D00001"
      Height          =   180
      Left            =   810
      TabIndex        =   39
      Top             =   0
      Width           =   720
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "承运车辆:"
      Height          =   180
      Index           =   2
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   810
   End
End
Attribute VB_Name = "frmSheetInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SheetID As String


Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

  frmRePrintSheet.Show vbModal
 
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
   AlignFormPos Me
   RefreshFill

End Sub

Private Sub RefreshFill()
On Error GoTo ErrHandle
 Dim tPriceItem() As TLuggagePriceItem
 Dim i As Integer
 Dim nlen As Integer
   moCarrySheet.Identify SheetID
   lblSheetID.Caption = Trim(moCarrySheet.SheetID)
   lblCarryBus.Caption = Trim(moCarrySheet.BusID)
   If moCarrySheet.BusDate = "1900-1-1" Then
      lblBusDate.Caption = ""
   Else
      lblBusDate.Caption = CStr(moCarrySheet.BusDate)
   End If
   lblAcceptType.Caption = Trim(moCarrySheet.AcceptType)
   lblOffBusTime.Caption = CStr(Format(moCarrySheet.BusStartOffTime, "hh:mm"))
   lblVehicle.Caption = Trim(moCarrySheet.VehicleLicense)
   lblVehicleType.Caption = Trim(moCarrySheet.VehicleTypeName)
   lblOwner.Caption = Trim(moCarrySheet.BusOwnerName)
   lblRoute.Caption = Trim(moCarrySheet.RouteName)
   lblCompany.Caption = Trim(moCarrySheet.CompanyName)
   lblProtocol.Caption = Trim(moCarrySheet.ProtocolName)
   lblDesCompany.Caption = Trim(moCarrySheet.SplitCompanyName)
   lblShipperNumber.Caption = CStr(moCarrySheet.AcceptSheetNumber)
   lblTotalNum.Caption = CStr(moCarrySheet.Number)
   lblOverNumber.Caption = CStr(moCarrySheet.OverNumber)
   lblTotalCatWeight.Caption = CStr(moCarrySheet.CalWeight)
   lblTotalActWeight.Caption = CStr(moCarrySheet.ActWeight)
   lblOperater.Caption = Trim(moCarrySheet.OperatorName)
   lblOperationTime.Caption = CStr(moCarrySheet.OperateTime)
   lblTotalPrice.Caption = CStr(moCarrySheet.TotalPrice)
   If moCarrySheet.Status = 0 Then
     lblStatus.ForeColor = vbRed
   Else
     lblStatus.ForeColor = 0
   End If
   lblStatus.Caption = Trim(moCarrySheet.StatusString)
   
    'VSFGTotalPrice 填
   nlen = ArrayLength(moCarrySheet.PriceItems)
   If nlen > 0 Then
      ReDim tPriceItem(1 To nlen)
      tPriceItem = moCarrySheet.PriceItems
      VSFGTotalPrice.Cols = nlen
      For i = 0 To nlen - 1
          VSFGTotalPrice.ColWidth(i) = VSFGTotalPrice.Width * 0.26
          VSFGTotalPrice.TextMatrix(0, i) = tPriceItem(i + 1).PriceName
          VSFGTotalPrice.TextMatrix(1, i) = tPriceItem(i + 1).PriceValue
      Next i
   End If
   
 Exit Sub
ErrHandle:
 ShowErrorMsg
 
End Sub
Private Sub Form_Unload(Cancel As Integer)
   SaveFormPos Me
End Sub


Private Sub Label13_Click()

End Sub

