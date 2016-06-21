VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmAcceptInfo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "行包单"
   ClientHeight    =   6270
   ClientLeft      =   2835
   ClientTop       =   2325
   ClientWidth     =   7455
   Icon            =   "frmAcceptInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -30
      TabIndex        =   4
      Top             =   690
      Width           =   7815
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   -30
      ScaleHeight     =   735
      ScaleWidth      =   7785
      TabIndex        =   5
      Top             =   0
      Width           =   7785
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包受理单信息:"
         Height          =   180
         Left            =   270
         TabIndex        =   6
         Top             =   270
         Width           =   1350
      End
      Begin VB.Image Image3 
         Height          =   855
         Left            =   1950
         Picture         =   "frmAcceptInfo.frx":038A
         Top             =   0
         Width           =   5925
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   540
      TabIndex        =   3
      Top             =   3900
      Width           =   6405
      Begin VSFlex7LCtl.VSFlexGrid VSPriceItme 
         Height          =   660
         Left            =   870
         TabIndex        =   49
         Top             =   540
         Width           =   5295
         _cx             =   9340
         _cy             =   1164
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
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次日期:"
         Height          =   180
         Left            =   2640
         TabIndex        =   48
         Top             =   1305
         Width           =   810
      End
      Begin VB.Label lblBusDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3480
         TabIndex        =   47
         Top             =   1305
         Width           =   90
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "承运车次:"
         Height          =   180
         Left            =   840
         TabIndex        =   46
         Top             =   1305
         Width           =   810
      End
      Begin VB.Label lblBus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1650
         TabIndex        =   45
         Top             =   1305
         Width           =   90
      End
      Begin VB.Label lblTicketPrice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1860
         TabIndex        =   44
         Top             =   270
         Width           =   90
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总运价:"
         Height          =   180
         Left            =   840
         TabIndex        =   43
         Top             =   270
         Width           =   630
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   210
         Picture         =   "frmAcceptInfo.frx":1874
         Top             =   285
         Width           =   480
      End
   End
   Begin RTComctl3.CoolButton cmdOK 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   5340
      TabIndex        =   7
      Top             =   5775
      Width           =   1620
      _ExtentX        =   2858
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
      MICON           =   "frmAcceptInfo.frx":1B7E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdDetail 
      Default         =   -1  'True
      Height          =   345
      Left            =   3630
      TabIndex        =   50
      Top             =   5775
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "行包明细(&L)"
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
      MICON           =   "frmAcceptInfo.frx":1B9A
      PICN            =   "frmAcceptInfo.frx":1BB6
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
      Height          =   3120
      Left            =   -120
      TabIndex        =   8
      Top             =   5550
      Width           =   8745
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3180
      Left            =   540
      TabIndex        =   0
      Top             =   720
      Width           =   6405
      Begin VB.Label lblShippePhone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   2220
         TabIndex        =   52
         Top             =   1800
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "托运人联系电话:"
         Height          =   210
         Left            =   840
         TabIndex        =   51
         Top             =   1800
         Width           =   1350
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "交付方式:"
         Height          =   180
         Left            =   4560
         TabIndex        =   42
         Top             =   1530
         Width           =   810
      End
      Begin VB.Label lblPickType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   5370
         TabIndex        =   41
         Tag             =   "提货人"
         Top             =   1530
         Width           =   90
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000005&
         Index           =   1
         X1              =   840
         X2              =   6120
         Y1              =   2625
         Y2              =   2625
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000003&
         Index           =   1
         X1              =   840
         X2              =   6120
         Y1              =   2610
         Y2              =   2610
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收件人地址:"
         Height          =   180
         Left            =   840
         TabIndex        =   40
         Top             =   2340
         Width           =   990
      End
      Begin VB.Label lblPickAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1860
         TabIndex        =   39
         Tag             =   "提货人"
         Top             =   2340
         Width           =   90
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收件人联系电话:"
         Height          =   180
         Left            =   840
         TabIndex        =   38
         Top             =   2070
         Width           =   1350
      End
      Begin VB.Label lblPhone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   2220
         TabIndex        =   37
         Tag             =   "提货人"
         Top             =   2070
         Width           =   90
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000005&
         Index           =   0
         X1              =   840
         X2              =   6120
         Y1              =   1395
         Y2              =   1395
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000003&
         Index           =   0
         X1              =   840
         X2              =   6120
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   5040
         TabIndex        =   36
         Top             =   270
         Width           =   90
      End
      Begin VB.Label lblOperationTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3600
         TabIndex        =   35
         Top             =   2730
         Width           =   90
      End
      Begin VB.Label lblOperater 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1500
         TabIndex        =   34
         Top             =   2730
         Width           =   90
      End
      Begin VB.Label lblStartStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1500
         TabIndex        =   33
         Top             =   540
         Width           =   90
      End
      Begin VB.Label lblEndStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3270
         TabIndex        =   32
         Top             =   540
         Width           =   90
      End
      Begin VB.Label lblOperatorChange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "操作员:"
         Height          =   180
         Left            =   870
         TabIndex        =   31
         Top             =   2730
         Width           =   630
      End
      Begin VB.Label lblStateChange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "状态:"
         Height          =   180
         Left            =   4560
         TabIndex        =   30
         Top             =   270
         Width           =   450
      End
      Begin VB.Label label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "到站:"
         Height          =   180
         Index           =   0
         Left            =   2760
         TabIndex        =   29
         Top             =   540
         Width           =   450
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起点站:"
         Height          =   180
         Index           =   1
         Left            =   840
         TabIndex        =   28
         Top             =   540
         Width           =   630
      End
      Begin VB.Label lblTimeChange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "受理时间:"
         Height          =   180
         Left            =   2760
         TabIndex        =   27
         Top             =   2730
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "计重:"
         Height          =   180
         Left            =   840
         TabIndex        =   26
         Top             =   1110
         Width           =   450
      End
      Begin VB.Label lblCalWeight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1320
         TabIndex        =   25
         Top             =   1110
         Width           =   90
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "托运人:"
         Height          =   180
         Left            =   840
         TabIndex        =   24
         Top             =   1530
         Width           =   630
      End
      Begin VB.Label lblShipper 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1500
         TabIndex        =   23
         Top             =   1530
         Width           =   90
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收件人:"
         Height          =   180
         Left            =   2760
         TabIndex        =   22
         Top             =   1530
         Width           =   630
      End
      Begin VB.Label lblPicker 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3420
         TabIndex        =   21
         Tag             =   "提货人"
         Top             =   1530
         Width           =   90
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "里程:"
         Height          =   180
         Left            =   4560
         TabIndex        =   20
         Top             =   540
         Width           =   450
      End
      Begin VB.Label lblMileage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   5040
         TabIndex        =   19
         Top             =   540
         Width           =   90
      End
      Begin VB.Label lblActWeight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3240
         TabIndex        =   18
         Top             =   1110
         Width           =   90
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "实重:"
         Height          =   180
         Left            =   2760
         TabIndex        =   17
         Top             =   1110
         Width           =   450
      End
      Begin VB.Label lblLabelID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1500
         TabIndex        =   16
         Top             =   825
         Width           =   90
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标签号:"
         Height          =   180
         Left            =   840
         TabIndex        =   15
         Top             =   825
         Width           =   630
      End
      Begin VB.Label lblBagNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   5040
         TabIndex        =   14
         Top             =   825
         Width           =   90
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "件数:"
         Height          =   180
         Left            =   4560
         TabIndex        =   13
         Top             =   825
         Width           =   450
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "托运方式:"
         Height          =   180
         Left            =   2760
         TabIndex        =   12
         Top             =   270
         Width           =   810
      End
      Begin VB.Label lblAcceptType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3600
         TabIndex        =   11
         Top             =   270
         Width           =   90
      End
      Begin VB.Label lblOverNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   5400
         TabIndex        =   10
         Top             =   1110
         Width           =   90
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "超重件数:"
         Height          =   180
         Left            =   4560
         TabIndex        =   9
         Top             =   1110
         Width           =   810
      End
      Begin VB.Label lblLuggageUnitID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1845
         TabIndex        =   2
         Top             =   270
         Width           =   90
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包单代码:"
         Height          =   180
         Index           =   0
         Left            =   840
         TabIndex        =   1
         Top             =   270
         Width           =   990
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   195
         Picture         =   "frmAcceptInfo.frx":2150
         Top             =   210
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmAcceptInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LuggageID As String
Private Sub cmdDetail_Click()
'需要更改
    frmLugDetail.LuggageID = LuggageID
    frmLugDetail.Show vbModal
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
   AlignFormPos Me
'   LuggageID = g_szAcceptSheetID
    RefreshFill
'   HideSheetNoLabel
   
End Sub
Private Sub RefreshFill()
On Error GoTo ErrHandle
 Dim tPriceItem() As TLuggagePriceItem
 Dim i As Integer
 Dim nlen As Integer
     moAcceptSheet.Identify LuggageID
   lblLuggageUnitID.Caption = LuggageID
   lblAcceptType.Caption = moAcceptSheet.AcceptType
   If moAcceptSheet.Status <> 0 Then
   lblStatus.ForeColor = vbRed
   Else
   lblStatus.ForeColor = 0
   End If
   lblStatus.Caption = moAcceptSheet.StatusString
   lblStartStation.Caption = moAcceptSheet.StartStationName
   lblEndStation.Caption = moAcceptSheet.DesStationName
   lblMileage.Caption = moAcceptSheet.Mileage
   lblLabelID.Caption = CStr(moAcceptSheet.StartLabelID) & "-" & CStr(moAcceptSheet.EndLabelID)
   lblBagNumber.Caption = moAcceptSheet.Number
   lblCalWeight.Caption = moAcceptSheet.CalWeight
   lblActWeight.Caption = moAcceptSheet.ActWeight
   lblOverNumber.Caption = moAcceptSheet.OverNumber
   lblShipper.Caption = moAcceptSheet.Shipper
   lblPicker.Caption = moAcceptSheet.Picker
   lblPickType.Caption = moAcceptSheet.PickType
   lblPhone.Caption = moAcceptSheet.PickerPhone
   lblShippePhone.Caption = moAcceptSheet.LuggageShipperPhone
   lblPickAddress.Caption = moAcceptSheet.PickerAddress
   lblOperater.Caption = moAcceptSheet.Operator
   lblOperationTime.Caption = moAcceptSheet.OperateTime
   lblTicketPrice.Caption = moAcceptSheet.TotalPrice
   lblBus.Caption = moAcceptSheet.BusID
   If moAcceptSheet.BusDate = "1900-1-1" Then
   lblBusDate = ""
   Else
   lblBusDate.Caption = moAcceptSheet.BusDate
   End If
   nlen = ArrayLength(moAcceptSheet.PriceItems)
   If nlen > 0 Then
      ReDim tPriceItem(1 To nlen)
      tPriceItem = moAcceptSheet.PriceItems
      '显示列头
      VSPriceItme.Cols = nlen
      For i = 0 To nlen - 1
          VSPriceItme.ColWidth(i) = VSPriceItme.Width * 0.26
          VSPriceItme.TextMatrix(0, i) = tPriceItem(i + 1).PriceName
          VSPriceItme.TextMatrix(1, i) = tPriceItem(i + 1).PriceValue
      Next i
   End If
   
 Exit Sub
ErrHandle:
 ShowErrorMsg
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveFormPos Me
End Sub

