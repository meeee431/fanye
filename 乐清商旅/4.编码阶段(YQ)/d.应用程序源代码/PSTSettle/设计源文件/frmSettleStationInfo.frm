VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmSettleStationInfo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "结算单站点信息"
   ClientHeight    =   5700
   ClientLeft      =   3360
   ClientTop       =   2130
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8070
   StartUpPosition =   2  '屏幕中心
   Begin VSFlex7LCtl.VSFlexGrid vsStation 
      Height          =   3765
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   7755
      _cx             =   13679
      _cy             =   6641
      _ConvInfo       =   -1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1320
      Top             =   4725
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   -75
      ScaleHeight     =   825
      ScaleWidth      =   8685
      TabIndex        =   1
      Top             =   0
      Width           =   8685
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算单站点信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Width           =   1365
      End
      Begin VB.Label lblContent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   450
         TabIndex        =   2
         Top             =   420
         Width           =   90
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -135
      TabIndex        =   0
      Top             =   810
      Width           =   8775
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   5595
      TabIndex        =   5
      Top             =   5160
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   714
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
      MICON           =   "frmSettleStationInfo.frx":0000
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
      Height          =   2880
      Left            =   -75
      TabIndex        =   4
      Top             =   4905
      Width           =   9465
   End
End
Attribute VB_Name = "frmSettleStationInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_szSheetID As String

'网格的列首
'Const VI_SellStation = 1
'Const VI_Route = 2
''Const VI_Bus = 3
'Const VI_VehicleType = 3
'Const VI_Station = 4
'Const VI_TicketType = 5
'Const VI_Quantity = 6
'Const VI_PassCharge = 7
'Const VI_SettlePrice = 8
'Const VI_HalvePrice = 9
'Const VI_ServicePrice = 10
'Const VI_SpringPrice = 11

Const VI_Route = 1
Const VI_Station = 2
Const VI_TicketType = 3
Const VI_Quantity = 4


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    vsStation.Cols = 5 '12
    vsStation.FixedCols = 1
    vsStation.Rows = 2
    vsStation.FixedRows = 1
    
    
    FillColHead
    AlignHeadWidth Me.name, vsStation
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveHeadWidth Me.name, vsStation
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    FillSettleStation

End Sub


Private Sub FillColHead()
    '填充列首信息
    

    With vsStation
'        .TextMatrix(0, VI_SellStation) = "上车站"
        .TextMatrix(0, VI_Route) = "线路"
'        .TextMatrix(0, VI_Bus) = "车次"
        .TextMatrix(0, VI_Station) = "站点"
'        .TextMatrix(0, VI_VehicleType) = "车型"
        .TextMatrix(0, VI_TicketType) = "票种"
        .TextMatrix(0, VI_Quantity) = "人数"
'        .TextMatrix(0, VI_PassCharge) = "通行费"
'        .TextMatrix(0, VI_SettlePrice) = "结算价"
'        .TextMatrix(0, VI_HalvePrice) = "平分价"
'        .TextMatrix(0, VI_ServicePrice) = "劳务费"
'        .TextMatrix(0, VI_SpringPrice) = "春运费"
        
        
'
'        .ColWidth(0) = 100
'        .ColWidth(VI_SellStation) = 0
'        .ColWidth(VI_Route) = 915
'        .ColWidth(VI_VehicleType) = 510
'        .ColWidth(VI_Station) = 480
'        .ColWidth(VI_TicketType) = 495
'        .ColWidth(VI_Quantity) = 510
'        .ColWidth(VI_PassCharge) = 600
'        .ColWidth(VI_SettlePrice) = 1155
'        .ColWidth(VI_HalvePrice) = 1155
'        .ColWidth(VI_ServicePrice) = 1110
'        .ColWidth(VI_SpringPrice) = 1110
'
    
    End With

End Sub


Private Sub FillSettleStation()
    '填充结算单的站点信息
    
    Dim rsTemp As Recordset
    Dim oReport As New Report
    Dim i As Integer
    On Error GoTo ErrorHandle
    oReport.Init g_oActiveUser
    
    Set rsTemp = oReport.GetSettleRouteQuantity(m_szSheetID)
    If rsTemp.RecordCount = 0 Then Exit Sub
    With vsStation
        .Rows = rsTemp.RecordCount + 1
        For i = 1 To rsTemp.RecordCount
'            .TextMatrix(i, VI_SellStation) = FormatDbValue(rsTemp!sell_station_name)
            .TextMatrix(i, VI_Route) = FormatDbValue(rsTemp!route_name)
'            .TextMatrix(i, VI_VehicleType) = FormatDbValue(rsTemp!vehicle_type_name)
            .TextMatrix(i, VI_Station) = FormatDbValue(rsTemp!station_name)
'            .TextMatrix(i, vi_bus) = FormatDbValue(rsTemp!bus_id)
            .TextMatrix(i, VI_TicketType) = FormatDbValue(rsTemp!ticket_type_name)
            .TextMatrix(i, VI_Quantity) = FormatDbValue(rsTemp!Quantity)
'            .TextMatrix(i, VI_PassCharge) = FormatDbValue(rsTemp!pass_charge)
'            .TextMatrix(i, VI_SettlePrice) = FormatDbValue(rsTemp!settle_price)
'            .TextMatrix(i, VI_HalvePrice) = FormatDbValue(rsTemp!halve_price)
'            .TextMatrix(i, VI_ServicePrice) = FormatDbValue(rsTemp!service_price)
'            .TextMatrix(i, VI_SpringPrice) = FormatDbValue(rsTemp!spring_price)
            
            rsTemp.MoveNext
            
        Next i
    End With
    vsStation.MergeCells = flexMergeRestrictColumns
'    vsStation.MergeCol(VI_SellStation) = True
    vsStation.MergeCol(VI_Route) = True
'    vsStation.MergeCol(VI_VehicleType) = True
    vsStation.MergeCol(VI_Station) = True
    vsStation.MergeCol(VI_TicketType) = True

    vsStation.AllowUserResizing = flexResizeColumns
    Exit Sub
ErrorHandle:
    ShowErrorMsg

End Sub



