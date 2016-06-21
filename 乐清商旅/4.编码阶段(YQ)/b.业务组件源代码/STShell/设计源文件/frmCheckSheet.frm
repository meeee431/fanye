VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmCheckSheet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "路单"
   ClientHeight    =   6645
   ClientLeft      =   225
   ClientTop       =   2070
   ClientWidth     =   9195
   Icon            =   "frmCheckSheet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Tag             =   "Modal"
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7125
      Left            =   0
      ScaleHeight     =   7125
      ScaleWidth      =   9285
      TabIndex        =   1
      Top             =   0
      Width           =   9285
      Begin VB.Timer tmStart 
         Interval        =   10
         Left            =   4020
         Top             =   4200
      End
      Begin VSFlex7LCtl.VSFlexGrid VSCheckSheet 
         Height          =   4485
         Left            =   150
         TabIndex        =   2
         Top             =   1470
         Width           =   8895
         _cx             =   15690
         _cy             =   7911
         _ConvInfo       =   -1
         Appearance      =   0
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
         BackColorFixed  =   -2147483634
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   10
         FixedRows       =   2
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
         BackColorFrozen =   4194304
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin RTComctl3.CoolButton cmdOk 
         Cancel          =   -1  'True
         Default         =   -1  'True
         Height          =   360
         Left            =   7620
         TabIndex        =   0
         Top             =   6120
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   635
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
         MICON           =   "frmCheckSheet.frx":014A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblBusSerialNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   1725
         TabIndex        =   28
         Top             =   855
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   180
         Left            =   1500
         TabIndex        =   27
         Top             =   870
         Width           =   90
      End
      Begin VB.Label lblSheetTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "客运凭单"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3900
         TabIndex        =   26
         Top             =   120
         Width           =   1020
      End
      Begin VB.Label lblGenerateTiem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1999-09-09 09:33"
         Height          =   165
         Left            =   5550
         TabIndex        =   25
         Top             =   6330
         Width           =   1440
      End
      Begin VB.Label lblSheetNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "749312574382"
         Height          =   180
         Left            =   7320
         TabIndex        =   24
         Top             =   450
         Width           =   1080
      End
      Begin VB.Label lblCheckor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "******"
         Height          =   180
         Left            =   960
         TabIndex        =   23
         Top             =   450
         Width           =   540
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "制单时间:"
         Height          =   180
         Left            =   4680
         TabIndex        =   22
         Top             =   6330
         Width           =   810
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "盖章有效:"
         Height          =   180
         Left            =   210
         TabIndex        =   21
         Top             =   6330
         Width           =   810
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO."
         Height          =   180
         Left            =   7065
         TabIndex        =   20
         Top             =   450
         Width           =   270
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票员:"
         Height          =   180
         Left            =   210
         TabIndex        =   19
         Top             =   450
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "托运包件数(件):"
         Height          =   180
         Left            =   180
         TabIndex        =   18
         Top             =   6000
         Width           =   1350
      End
      Begin VB.Label LblPiece 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   1620
         TabIndex        =   17
         Top             =   6000
         Width           =   90
      End
      Begin VB.Label Lables2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "托运金额(元):"
         Height          =   180
         Left            =   4650
         TabIndex        =   16
         Top             =   6000
         Width           =   1170
      End
      Begin VB.Label LblCarriage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   5880
         TabIndex        =   15
         Top             =   6000
         Width           =   90
      End
      Begin VB.Label lblOwner 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "啊啊啊"
         Height          =   180
         Left            =   6480
         TabIndex        =   14
         Top             =   1230
         Width           =   540
      End
      Begin VB.Label lblLicense 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "******"
         Height          =   180
         Left            =   3810
         TabIndex        =   13
         Top             =   1230
         Width           =   540
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "******"
         Height          =   180
         Left            =   1200
         TabIndex        =   12
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label lblRoute 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "******"
         Height          =   180
         Left            =   6420
         TabIndex        =   11
         Top             =   840
         Width           =   540
      End
      Begin VB.Label lblStartupTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1999-09-09 09:23"
         Height          =   180
         Left            =   3420
         TabIndex        =   10
         Top             =   840
         Width           =   1440
      End
      Begin VB.Label lblBusID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "******"
         Height          =   180
         Left            =   960
         TabIndex        =   9
         Top             =   870
         Width           =   540
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车主:"
         Height          =   180
         Left            =   5790
         TabIndex        =   8
         Top             =   1230
         Width           =   450
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车牌:"
         Height          =   180
         Left            =   3210
         TabIndex        =   7
         Top             =   1230
         Width           =   450
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参营公司:"
         Height          =   180
         Left            =   300
         TabIndex        =   6
         Top             =   1200
         Width           =   810
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "线路:"
         Height          =   180
         Left            =   5760
         TabIndex        =   5
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发车时间:"
         Height          =   180
         Left            =   2550
         TabIndex        =   4
         Top             =   840
         Width           =   810
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次:"
         Height          =   180
         Left            =   300
         TabIndex        =   3
         Top             =   870
         Width           =   450
      End
      Begin VB.Shape Shape1 
         Height          =   735
         Left            =   150
         Top             =   750
         Width           =   8895
      End
      Begin VB.Line Line2 
         X1              =   150
         X2              =   9030
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Line Line3 
         X1              =   2460
         X2              =   2460
         Y1              =   750
         Y2              =   1110
      End
      Begin VB.Line Line4 
         X1              =   5670
         X2              =   5670
         Y1              =   750
         Y2              =   1470
      End
      Begin VB.Line Line5 
         X1              =   3120
         X2              =   3120
         Y1              =   1110
         Y2              =   1470
      End
   End
End
Attribute VB_Name = "frmCheckSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public g_oActiveUser As ActiveUser
Public moChkTicket As CheckTicket
Public mszSheetID As String

Dim mtSheetInfo As Sheet_ChkInfo
Dim matSheetContent() As SheetContent  '路单详细信息
Dim mnCountSheetDetail As Integer '路单详细信息的记录数
Dim TicketType() As TTicketType

'补充无用的，为了使用frmCheckSheet能与检票中的frmCheckSheet共享添加的附加参数，无实际用途
Public mbViewMode As Boolean
Public mbNoPrintPrompt As Boolean




Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub WriteChkInfo()
    Dim rsTemp As Recordset
'    If g_szTitle <> "" Then
'      lblSheetTitle.Caption = g_szTitle
'    End If
    lblBusID.Caption = mtSheetInfo.BusID
    lblBusSerialNO.Caption = mtSheetInfo.SerialNo
    lblCheckor.Caption = mtSheetInfo.Checker
    lblSheetNo.Caption = mtSheetInfo.SheetNo
    lblRoute.Caption = mtSheetInfo.Route
    lblCompany.Caption = mtSheetInfo.Company
    lblLicense.Caption = mtSheetInfo.VehicleTag
    lblOwner.Caption = mtSheetInfo.Owner
    lblStartupTime.Caption = IIf(DBDateTimeIsEmpty(mtSheetInfo.StartUpTime), "", Format(mtSheetInfo.StartUpTime, "YYYY-MM-DD HH:mm"))
    lblGenerateTiem.Caption = IIf(DBDateTimeIsEmpty(mtSheetInfo.dtTime), "", Format(mtSheetInfo.dtTime, "YYYY-MM-DD HH:mm"))
'    Set rsTemp = moChkTicket.GetLsInfo(mtSheetInfo.BusID, mtSheetInfo.SerialNo, Format(mtSheetInfo.StartUpTime, "YYYY-MM-DD"))
'    If rsTemp.RecordCount = 0 Then Exit Sub
'    LblPiece.Caption = rsTemp!baggage_number
'    LblCarriage.Caption = rsTemp!total_price
    
End Sub

Private Sub WriteSheetDetail()
    Dim i As Integer
   
    VSCheckSheet.Rows = IIf(mnCountSheetDetail > 17, mnCountSheetDetail, 17)
    
    For i = 1 To mnCountSheetDetail
            VSCheckSheet.TextMatrix(i + 1, 0) = i
            VSCheckSheet.TextMatrix(i + 1, 1) = matSheetContent(i).StationName
            Select Case VSCheckSheet.Cols
                 Case 6
                      FillTicketPrice i, 2
                 Case 8
                      FillTicketPrice i, 4
                 Case 10
                      FillTicketPrice i, 6
                 Case 12
                      FillTicketPrice i, 8
                 Case 14
                      FillTicketPrice i, 10
            End Select
    Next i
   
    VSCheckSheet.SubtotalPosition = flexSTBelow
    Dim sformat As String
    For i = 2 To VSCheckSheet.Cols - 1
        If i Mod 2 = 0 Then
             sformat = "###0"
        Else
             sformat = "###0.00"
        End If
        VSCheckSheet.Subtotal flexSTSum, -1, i, sformat, , vbBlue, , "合计", , True
    Next
End Sub
Public Sub RefreshForm()
    GetCheckSheetInfo
    WriteChkInfo
    WriteSheetDetail
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If moChkTicket Is Nothing Then
        Set moChkTicket = New CheckTicket
        moChkTicket.Init g_oActiveUser
    End If
    SetCheckSheetCaption VSCheckSheet
End Sub




Private Sub tmStart_Timer()
On Error GoTo ErrHandle
    tmStart.Enabled = False

    RefreshForm
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
Private Function GetTicketType(sChar As String) As Integer
    Dim cCount As Integer
    Dim i As Integer
    cCount = ArrayLength(TicketType)
    For i = 1 To cCount
        If Trim(TicketType(i).szTicketTypeName) = Trim(sChar) Then
              GetTicketType = TicketType(i).nTicketTypeID
              Exit For
        End If
    Next i
End Function
 '填充对应表种的票价
Private Function FillTicketPrice(cRow As Integer, cCols As Integer)
       Dim i As Integer
         For i = 2 To cCols Step 2
            Select Case GetTicketType(VSCheckSheet.TextMatrix(0, i))
                Case TP_FullPrice
                          VSCheckSheet.TextMatrix(cRow + 1, i) = matSheetContent(cRow).FullTk_Numer
                          VSCheckSheet.TextMatrix(cRow + 1, i + 1) = Format(matSheetContent(cRow).FullTk_Price, "#,##0.00")
                Case TP_HalfPrice
                          VSCheckSheet.TextMatrix(cRow + 1, i) = matSheetContent(cRow).HalfTk_Numer
                          VSCheckSheet.TextMatrix(cRow + 1, i + 1) = Format(matSheetContent(cRow).HalfTk_Price, "#,##0.00")
                Case TP_PreferentialTicket1
                          VSCheckSheet.TextMatrix(cRow + 1, i) = matSheetContent(cRow).PreferentialTk1_Numer
                          VSCheckSheet.TextMatrix(cRow + 1, i + 1) = Format(matSheetContent(cRow).PreferentialTk1_Price, "#,##0.00")
                Case TP_PreferentialTicket2
                          VSCheckSheet.TextMatrix(cRow + 1, i) = matSheetContent(cRow).PreferentialTk2_Numer
                          VSCheckSheet.TextMatrix(cRow + 1, i + 1) = Format(matSheetContent(cRow).PreferentialTk2_Price, "#,##0.00")
                Case TP_PreferentialTicket3
                          VSCheckSheet.TextMatrix(cRow + 1, i) = matSheetContent(cRow).PreferentialTk3_Numer
                          VSCheckSheet.TextMatrix(cRow + 1, i + 1) = Format(matSheetContent(cRow).PreferentialTk3_Price, "#,##0.00")
             End Select
             VSCheckSheet.TextMatrix(cRow + 1, VSCheckSheet.Cols - 1) = Format(matSheetContent(cRow).FullTk_Price + matSheetContent(cRow).HalfTk_Price + matSheetContent(cRow).PreferentialTk1_Price + matSheetContent(cRow).PreferentialTk2_Price + matSheetContent(cRow).PreferentialTk3_Price, "#,##0.00")
             VSCheckSheet.TextMatrix(cRow + 1, VSCheckSheet.Cols - 2) = matSheetContent(cRow).FullTk_Numer + matSheetContent(cRow).HalfTk_Numer + matSheetContent(cRow).PreferentialTk1_Numer + matSheetContent(cRow).PreferentialTk2_Numer + matSheetContent(cRow).PreferentialTk3_Numer
             Next i
End Function

Private Sub SetCheckSheetCaption(Vsobject As VSFlexGrid) '填写路单各种票价标题
    Dim i As Integer
    Dim Count As Integer
    Dim sysParam As New SystemParam
    Dim j As Integer
    Dim nCount As Integer
    TicketType = sysParam.GetAllTicketType(1)
    Set sysParam = Nothing
    nCount = ArrayLength(TicketType)
    For i = 1 To nCount
        If TicketType(i).nTicketTypeID <> TP_FreeTicket Then Count = Count + 1
    Next i
  With Vsobject
                .Rows = 17
                .Cols = Count * 2 + 4
                .ColWidth(0) = 500
                .ColWidth(1) = 2000
                .TextMatrix(0, 0) = "序号"
                .TextMatrix(0, 1) = "站点"
                .TextMatrix(0, .Cols - 2) = "合计"
                .TextMatrix(0, .Cols - 1) = "合计"
                .TextMatrix(1, 0) = "序号"
                .TextMatrix(1, 1) = "站点"
                .TextMatrix(1, .Cols - 2) = "人数"
                .TextMatrix(1, .Cols - 1) = "金额"
                 For i = 2 To nCount * 2 Step 2
                   If TicketType(i / 2).nTicketTypeID <> TP_FreeTicket Then
                         j = j + 2
                        .TextMatrix(0, j) = Trim(TicketType(i / 2).szTicketTypeName)
                        .TextMatrix(0, j + 1) = Trim(TicketType(i / 2).szTicketTypeName)
                        .ColWidth(j) = 1000
                        .ColWidth(j + 1) = 800
                        .TextMatrix(1, j) = "人数"
                        .TextMatrix(1, j + 1) = "金额"
                   End If
                 Next i
                .MergeCells = 5
                .MergeRow(0) = True
                .MergeCol(0) = True
                .MergeCol(1) = True
                For i = 0 To .Cols - 1
                    .FixedAlignment(i) = flexAlignCenterCenter
                    .ColAlignment(i) = flexAlignGeneral
                Next i
  End With
End Sub

Private Sub GetCheckSheetInfo()
'********************************************************************
'取得指定路单窗体中的检票信息和详细路单信息
'********************************************************************
    Dim aSheetResult()  As TCheckSheetStationInfoEx
    Dim tSheetInfo As TCheckSheetInfo
    Dim oVehicle As Vehicle
    Dim oRoute As Route
    Dim nCount As Integer
    Dim szStation As String
    Dim i As Integer, j As Integer
    Dim moChkTicket As New CheckTicket
    moChkTicket.Init g_oActiveUser
    
    Set oVehicle = New Vehicle
    Set oRoute = New Route
    
    tSheetInfo = moChkTicket.GetCheckSheetInfo(mszSheetID)
    mtSheetInfo.SheetNo = Trim(tSheetInfo.szCheckSheet)
    mtSheetInfo.BusID = Trim(tSheetInfo.szBusID)
    mtSheetInfo.SerialNo = Trim(tSheetInfo.nBusSerialNo)
    mtSheetInfo.dtTime = tSheetInfo.dtMakeSheetDateTime
    mtSheetInfo.Checker = Trim(tSheetInfo.szMakeSheetUser)
    If mtSheetInfo.Checker = g_oActiveUser.UserID Then
        mtSheetInfo.Checker = "[" & mtSheetInfo.Checker & "]" & g_oActiveUser.UserName
    Else
        Dim atUsers() As String
        atUsers = moChkTicket.GetAllUser
        For i = 1 To ArrayLength(atUsers)
            If Trim(atUsers(i, 1)) = mtSheetInfo.Checker Then
                mtSheetInfo.Checker = "[" & mtSheetInfo.Checker & "]" & Trim(atUsers(i, 2))
                Exit For
            End If
        Next i
            
    End If
    mtSheetInfo.StartUpTime = tSheetInfo.dtStartupTime
    oVehicle.Init g_oActiveUser
    oVehicle.Identify tSheetInfo.szVehicleID
    mtSheetInfo.Company = Trim(oVehicle.CompanyName)
    mtSheetInfo.Owner = Trim(oVehicle.OwnerName)
    mtSheetInfo.VehicleTag = Trim(oVehicle.LicenseTag)
    oRoute.Init g_oActiveUser
    oRoute.Identify Trim(tSheetInfo.szRouteID)
    mtSheetInfo.Route = oRoute.RouteName
        
    aSheetResult = moChkTicket.GetCheckSheetStationInfo(mszSheetID)
    nCount = ArrayLength(aSheetResult)
    If nCount > 0 Then
        ReDim matSheetContent(1 To nCount)
    End If
    j = 0
    For i = 1 To nCount
        If j = 0 Then
           matSheetContent(1).StationId = aSheetResult(1).szStationID
           j = 1
        End If
        If aSheetResult(i).szStationID <> matSheetContent(j).StationId Then
                j = j + 1
                matSheetContent(j).StationId = aSheetResult(i).szStationID
        End If
'        If aSheetResult(i).nCheckStatus <> ECheckedTicketStatus.NormalTicket Then
'            matSheetContent(j).StationName = Trim(LeftAndRight(LeftAndRight(aSheetResult(i).szCheckSheet, False, "["), True, "]")) & "(改并)"
'        Else
            matSheetContent(j).StationName = Trim(LeftAndRight(LeftAndRight(aSheetResult(i).szCheckSheet, False, "["), True, "]"))
'        End If
        If aSheetResult(i).nTicketType = TP_FullPrice Then
            matSheetContent(j).FullTk_Numer = aSheetResult(i).nManCount
            matSheetContent(j).FullTk_Price = aSheetResult(i).sgTicketPrice
        End If
        If aSheetResult(i).nTicketType = TP_HalfPrice Then
            matSheetContent(j).HalfTk_Numer = aSheetResult(i).nManCount
            matSheetContent(j).HalfTk_Price = aSheetResult(i).sgTicketPrice
        End If
        If aSheetResult(i).nTicketType = TP_PreferentialTicket1 Then
            matSheetContent(j).PreferentialTk1_Numer = aSheetResult(i).nManCount
            matSheetContent(j).PreferentialTk1_Price = aSheetResult(i).sgTicketPrice
        End If
        If aSheetResult(i).nTicketType = TP_PreferentialTicket2 Then
            matSheetContent(j).PreferentialTk2_Numer = aSheetResult(i).nManCount
            matSheetContent(j).PreferentialTk2_Price = aSheetResult(i).sgTicketPrice
        End If
        If aSheetResult(i).nTicketType = TP_PreferentialTicket3 Then
            matSheetContent(j).PreferentialTk3_Numer = aSheetResult(i).nManCount
            matSheetContent(j).PreferentialTk3_Price = aSheetResult(i).sgTicketPrice
        End If
        If aSheetResult(i).nTicketType = TP_FreeTicket Then
            matSheetContent(j).FullTk_Numer = matSheetContent(j).FullTk_Numer + aSheetResult(i).nManCount
            matSheetContent(j).FullTk_Price = matSheetContent(j).FullTk_Price + aSheetResult(i).sgTicketPrice
        End If
    Next i
    mnCountSheetDetail = j
    Exit Sub
End Sub
'根据车次设置路单号
Public Sub SetSheetID(BusDate As Date, BusID As String, BusSerialNo As Integer)
'    Dim tCheckSheet As TCheckSheetInfo
'    tCheckSheet = moChkTicket.GetBusCheckSheet(BusDate, BusID, BusSerialNo)
'    mszSheetID = tCheckSheet.szCheckSheet
End Sub
'判断指定的数据库日期时间是否是空的
Public Function DBDateTimeIsEmpty(pdtIn As Date) As Boolean
    'Dim dtTemp As Date
    DBDateTimeIsEmpty = IIf(Format(pdtIn, cszDateTimeStr) = Format(cdtEmptyDateTime, cszDateTimeStr), True, False)
End Function
