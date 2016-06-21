VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmEditSettlePrice 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "结算价设置"
   ClientHeight    =   6150
   ClientLeft      =   3195
   ClientTop       =   2235
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   10515
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboCompany 
      Height          =   300
      Left            =   5130
      TabIndex        =   21
      Top             =   1260
      Width           =   1905
   End
   Begin VB.CheckBox chkBack 
      BackColor       =   &H00E0E0E0&
      Caption         =   "回程线路"
      Height          =   225
      Left            =   2970
      TabIndex        =   20
      Top             =   1290
      Width           =   1950
   End
   Begin VB.ComboBox cboVehicleType 
      Height          =   300
      Left            =   6270
      TabIndex        =   19
      Top             =   930
      Width           =   1515
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -150
      TabIndex        =   16
      Top             =   720
      Width           =   10755
   End
   Begin FText.asFlatTextBox txtCompany 
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      Top             =   930
      Width           =   1305
      _ExtentX        =   2302
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
      ButtonHotBackColor=   -2147483633
      ButtonPressedBackColor=   -2147483627
      Text            =   ""
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin VSFlex7LCtl.VSFlexGrid VsSettlePrice 
      Height          =   3585
      Left            =   150
      TabIndex        =   6
      Top             =   1650
      Width           =   10185
      _cx             =   17965
      _cy             =   6324
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   14737632
      GridColorFixed  =   14737632
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
      Rows            =   10
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSettlePrice.frx":0000
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
      Editable        =   2
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      CellButtonPicture=   "frmSettlePrice.frx":00FD
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   930
      Left            =   -60
      TabIndex        =   2
      Top             =   5370
      Width           =   10725
      Begin RTComctl3.CoolButton cmdRowADD 
         Height          =   345
         Left            =   6420
         TabIndex        =   10
         Top             =   990
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "新增行(&A)"
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
         MICON           =   "frmSettlePrice.frx":0944
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdRowDel 
         Height          =   345
         Left            =   5130
         TabIndex        =   9
         Top             =   990
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "删除行(&D)"
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
         MICON           =   "frmSettlePrice.frx":0960
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdHelp 
         Height          =   315
         Left            =   1260
         TabIndex        =   3
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
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
         MICON           =   "frmSettlePrice.frx":097C
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
         Cancel          =   -1  'True
         Height          =   345
         Left            =   8505
         TabIndex        =   4
         Top             =   285
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
         MICON           =   "frmSettlePrice.frx":0998
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
         Height          =   345
         Left            =   7050
         TabIndex        =   5
         ToolTipText     =   "保存协议"
         Top             =   285
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
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
         MICON           =   "frmSettlePrice.frx":09B4
         PICN            =   "frmSettlePrice.frx":09D0
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   10605
      TabIndex        =   0
      Top             =   0
      Width           =   10605
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "**结算价"
         Height          =   180
         Left            =   810
         TabIndex        =   8
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lbl232 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   330
         TabIndex        =   1
         Top             =   300
         Width           =   90
      End
   End
   Begin FText.asFlatTextBox txtRouteID 
      Height          =   285
      Left            =   3630
      TabIndex        =   15
      Top             =   930
      Width           =   1305
      _ExtentX        =   2302
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
      ButtonHotBackColor=   -2147483633
      ButtonPressedBackColor=   -2147483627
      Text            =   ""
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin FText.asFlatTextBox txtVehilce 
      Height          =   285
      Left            =   6270
      TabIndex        =   17
      Top             =   930
      Width           =   3705
      _ExtentX        =   6535
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
      ButtonHotBackColor=   -2147483633
      ButtonPressedBackColor=   -2147483627
      Text            =   ""
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin VB.Label lblVehicleType 
      BackStyle       =   0  'Transparent
      Caption         =   "车型代码："
      Height          =   225
      Left            =   5370
      TabIndex        =   18
      Top             =   990
      Width           =   915
   End
   Begin VB.Label lbl4 
      BackStyle       =   0  'Transparent
      Caption         =   "线路"
      Height          =   225
      Left            =   2940
      TabIndex        =   14
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "车型代码"
      Height          =   195
      Left            =   5370
      TabIndex        =   13
      Top             =   990
      Width           =   825
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "公司代码"
      Height          =   255
      Left            =   180
      TabIndex        =   11
      Top             =   960
      Width           =   915
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结算价列表(&L):"
      Height          =   180
      Left            =   180
      TabIndex        =   7
      Top             =   1275
      Width           =   1260
   End
End
Attribute VB_Name = "frmEditSettlePrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_oCompanySettlePrice As New CompanySettlePrice
Private m_oVehicleSettlePrice As New VehicleSettlePrice
Private m_oBusSettlePrice As New BusSettlePrice
Private m_oReport As New Report
Private m_oRoute As New BaseInfo
Private nVehicleCount As Integer
Private nComapnyCount As Integer

Private m_aszTemp() As String

Public szTitle As String

Public m_eFormStatus As EFormStatus

Public m_szCompany As String
Public m_szVehicle As String
Public m_szVehicleType As String
Public m_szRoute As String
Public m_szBus As String
Public m_szTransportCompany As String

Private Sub cboVehicleType_Click()
    If cboVehicleType.Text <> "" And txtCompany.Text <> "" And txtRouteID.Text <> "" Then
        cmdOk.Enabled = True
    Else
        cmdOk.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error GoTo here
    Dim i As Integer
'    Dim szSettlePrice() As String
    '判断是否有重复项

    Dim j As Integer
    Dim szRouteName As String
    If vsSettlePrice.Rows >= 1 Then
        If lblTitle = "公司结算价" Then
            For i = 1 To vsSettlePrice.Rows - 1

                m_oCompanySettlePrice.Init g_oActiveUser
                If m_eFormStatus = AddStatus Then

                    m_oCompanySettlePrice.AddNew
                Else
                    m_oCompanySettlePrice.Identify ResolveDisplay(txtCompany.Text), ResolveDisplay(txtRouteID.Text, szRouteName), ResolveDisplay(cboVehicleType.Text), ResolveDisplay(vsSettlePrice.TextMatrix(i, 1)), ResolveDisplay(vsSettlePrice.TextMatrix(i, 2))
                End If
                m_oCompanySettlePrice.CompanyID = ResolveDisplay(txtCompany.Text)
                m_oCompanySettlePrice.VehicleTypeCode = ResolveDisplay(cboVehicleType.Text)
                m_oCompanySettlePrice.SellStationID = ResolveDisplay(vsSettlePrice.TextMatrix(i, 1))
                m_oCompanySettlePrice.RouteID = ResolveDisplay(txtRouteID.Text, szRouteName)
                m_oCompanySettlePrice.StationID = ResolveDisplay(vsSettlePrice.TextMatrix(i, 2))
                m_oCompanySettlePrice.Mileage = vsSettlePrice.TextMatrix(i, 3)
                m_oCompanySettlePrice.PassCharge = vsSettlePrice.TextMatrix(i, 4)
                m_oCompanySettlePrice.SettlefullPrice = vsSettlePrice.TextMatrix(i, 5)
                m_oCompanySettlePrice.SettleHalfPrice = vsSettlePrice.TextMatrix(i, 6)
                m_oCompanySettlePrice.HalveFullPrice = vsSettlePrice.TextMatrix(i, 7)
                m_oCompanySettlePrice.HalveHalfPrice = vsSettlePrice.TextMatrix(i, 8)
                m_oCompanySettlePrice.ServiceFullPrice = vsSettlePrice.TextMatrix(i, 9)
                m_oCompanySettlePrice.ServiceHalfPrice = vsSettlePrice.TextMatrix(i, 10)
                m_oCompanySettlePrice.SpringFullPrice = vsSettlePrice.TextMatrix(i, 11)
                m_oCompanySettlePrice.SpringHalfPrice = vsSettlePrice.TextMatrix(i, 12)
                m_oCompanySettlePrice.Annotation = vsSettlePrice.TextMatrix(i, 13)
                m_oCompanySettlePrice.RouteName = szRouteName
                m_oCompanySettlePrice.Update
            Next i
            frmCompanySettlePrice.FillvsSettlePrice ResolveDisplay(txtCompany.Text), ResolveDisplay(txtRouteID.Text), ResolveDisplay(cboVehicleType.Text)
            Unload Me

            '调用接口
        ElseIf lblTitle = "车辆结算价" Then

            For j = 1 To ArrayLength(m_aszTemp)
                For i = 1 To vsSettlePrice.Rows - 1
                    m_oVehicleSettlePrice.Init g_oActiveUser
                    If m_eFormStatus = AddStatus Then
                        m_oVehicleSettlePrice.AddNew
                    Else
                        m_oVehicleSettlePrice.Identify m_aszTemp(j, 1), ResolveDisplay(txtRouteID.Text), ResolveDisplay(vsSettlePrice.TextMatrix(i, 1)), ResolveDisplay(vsSettlePrice.TextMatrix(i, 2))
                    End If
                    m_oVehicleSettlePrice.VehicleID = ResolveDisplay(m_aszTemp(j, 1))
                    m_oVehicleSettlePrice.SellStationID = ResolveDisplay(vsSettlePrice.TextMatrix(i, 1))
                    m_oVehicleSettlePrice.RouteID = ResolveDisplay(txtRouteID.Text, szRouteName)
                    m_oVehicleSettlePrice.StationID = ResolveDisplay(vsSettlePrice.TextMatrix(i, 2))
                    m_oVehicleSettlePrice.Mileage = vsSettlePrice.TextMatrix(i, 3)
                    m_oVehicleSettlePrice.PassCharge = vsSettlePrice.TextMatrix(i, 4)
                    m_oVehicleSettlePrice.SettlefullPrice = vsSettlePrice.TextMatrix(i, 5)
                    m_oVehicleSettlePrice.SettleHalfPrice = vsSettlePrice.TextMatrix(i, 6)
                    m_oVehicleSettlePrice.HalveFullPrice = vsSettlePrice.TextMatrix(i, 7)
                    m_oVehicleSettlePrice.HalveHalfPrice = vsSettlePrice.TextMatrix(i, 8)
                    m_oVehicleSettlePrice.ServiceFullPrice = vsSettlePrice.TextMatrix(i, 9)
                    m_oVehicleSettlePrice.ServiceHalfPrice = vsSettlePrice.TextMatrix(i, 10)
                    m_oVehicleSettlePrice.SpringFullPrice = vsSettlePrice.TextMatrix(i, 11)
                    m_oVehicleSettlePrice.SpringHalfPrice = vsSettlePrice.TextMatrix(i, 12)
                    m_oVehicleSettlePrice.Annotation = vsSettlePrice.TextMatrix(i, 13)
                    m_oVehicleSettlePrice.RouteName = szRouteName
                    m_oVehicleSettlePrice.Update
                Next i
            Next j

            frmVehicleSettlePrice.QueryVehicleSettlePrice , ResolveDisplay(txtRouteID.Text)
            Unload Me
            '调用接口

        Else
            For j = 1 To ArrayLength(m_aszTemp)
                For i = 1 To vsSettlePrice.Rows - 1
                    m_oBusSettlePrice.Init g_oActiveUser
                    If m_eFormStatus = AddStatus Then
                        m_oBusSettlePrice.AddNew
                    Else
                        m_oBusSettlePrice.Identify m_aszTemp(j, 1), ResolveDisplay(cboCompany.Text), ResolveDisplay(vsSettlePrice.TextMatrix(i, 1)), ResolveDisplay(vsSettlePrice.TextMatrix(i, 2))
                    End If
                    m_oBusSettlePrice.BusID = ResolveDisplay(m_aszTemp(j, 1))
                    m_oBusSettlePrice.SellStationID = ResolveDisplay(vsSettlePrice.TextMatrix(i, 1))
                    m_oBusSettlePrice.TransportCompanyID = ResolveDisplay(cboCompany.Text)
                    m_oBusSettlePrice.StationID = ResolveDisplay(vsSettlePrice.TextMatrix(i, 2))
                    m_oBusSettlePrice.Mileage = vsSettlePrice.TextMatrix(i, 3)
                    m_oBusSettlePrice.PassCharge = vsSettlePrice.TextMatrix(i, 4)
                    m_oBusSettlePrice.SettlefullPrice = vsSettlePrice.TextMatrix(i, 5)
                    m_oBusSettlePrice.SettleHalfPrice = vsSettlePrice.TextMatrix(i, 6)
                    m_oBusSettlePrice.HalveFullPrice = vsSettlePrice.TextMatrix(i, 7)
                    m_oBusSettlePrice.HalveHalfPrice = vsSettlePrice.TextMatrix(i, 8)
                    m_oBusSettlePrice.ServiceFullPrice = vsSettlePrice.TextMatrix(i, 9)
                    m_oBusSettlePrice.ServiceHalfPrice = vsSettlePrice.TextMatrix(i, 10)
                    m_oBusSettlePrice.SpringFullPrice = vsSettlePrice.TextMatrix(i, 11)
                    m_oBusSettlePrice.SpringHalfPrice = vsSettlePrice.TextMatrix(i, 12)
                    m_oBusSettlePrice.Annotation = vsSettlePrice.TextMatrix(i, 13)
                    m_oBusSettlePrice.Update
                Next i
            Next j

            frmBusSettlePrice.QueryBusSettlePrice txtVehilce.Text, ResolveDisplay(cboCompany.Text)
            Unload Me
            '调用接口

        End If
    End If

    Exit Sub
here:
    ShowErrorMsg
End Sub
Private Sub cmdRowADD_Click()
    vsSettlePrice.Rows = vsSettlePrice.Rows + 1
    cmdRowDel.Enabled = True
End Sub

Private Sub cmdRowDel_Click()
    If vsSettlePrice.Rows <> 1 Then
        vsSettlePrice.RemoveItem (vsSettlePrice.Row)
    Else
        cmdRowDel.Enabled = False
    End If

End Sub

Public Sub FillVehicleType(CompanyID As String, RouteID As String)
    Dim aszTemp() As String
    Dim i As Integer
    If chkBack.Value = vbUnchecked Then
        aszTemp = m_oReport.GetVehileType(ResolveDisplay(CompanyID), ResolveDisplay(RouteID))
        cboVehicleType.Clear
        If ArrayLength(aszTemp) <> 0 Then
    '        cboVehicleType.AddItem ""
            For i = 1 To ArrayLength(aszTemp)
                cboVehicleType.AddItem MakeDisplayString(aszTemp(i, 1), aszTemp(i, 2))
            Next i
            cboVehicleType.ListIndex = 0
        End If
    Else
        FillAllVehicleType

    End If
End Sub

Private Sub FillAllVehicleType()
    '填充所有的车型
    Dim oBaseInfo As New BaseInfo
    Dim aszTemp() As String
    Dim i As Integer
    oBaseInfo.Init g_oActiveUser
    aszTemp = oBaseInfo.GetAllVehicleModel()

    cboVehicleType.Clear
    If ArrayLength(aszTemp) <> 0 Then
        For i = 1 To ArrayLength(aszTemp)
            cboVehicleType.AddItem MakeDisplayString(aszTemp(i, 1), aszTemp(i, 2))
        Next i
        cboVehicleType.ListIndex = 0
    End If


End Sub


Private Sub Form_Load()


    AlignFormPos Me
    m_oReport.Init g_oActiveUser
    lblTitle.Caption = szTitle
    cmdOk.Enabled = False
    If szTitle = "公司结算价" Then
        cboVehicleType.Visible = True
        lblVehicleType.Visible = True
        lbl2.Visible = False
        txtVehilce.Visible = False
        FillVehicleType "", ""

        FillCompanyVSHead
        lbl1.Visible = True
        lbl2.Caption = "车型代码"
        txtCompany.Text = ""
        txtCompany.Visible = True
        cboCompany.Visible = False

        If m_eFormStatus = ModifyStatus Then
            txtCompany.Text = m_szCompany
            txtRouteID.Text = m_szRoute
            cboVehicleType.Text = m_szVehicleType
            txtCompany.Enabled = False
            txtRouteID.Enabled = False
            cboVehicleType.Enabled = False
            chkBack.Enabled = False
            RefreshCompanySettlePriceInfo
        ElseIf m_eFormStatus = AddStatus Then
            '列表新增一列
            vsSettlePrice.Rows = 2
        End If
    ElseIf szTitle = "车辆结算价" Then
        cboVehicleType.Visible = False
        lblVehicleType.Visible = False
        lbl2.Visible = True
        txtVehilce.Visible = True
        cboCompany.Visible = False

        FillVehicleVSHead
        lbl1.Visible = False
        txtCompany.Visible = False
        txtCompany.Text = "Temp"
        lbl2.Caption = "车辆代码"

        If m_eFormStatus = ModifyStatus Then
            txtVehilce.Text = m_szVehicle
            txtRouteID.Text = m_szRoute
            RefreshVehicleSettlePriceInfo
            txtVehilce.Enabled = False
            txtRouteID.Enabled = False
            chkBack.Enabled = False
            ReDim m_aszTemp(1 To 1, 1 To 2)
            m_aszTemp(1, 1) = m_szVehicle
            m_aszTemp(1, 2) = m_szVehicle

        ElseIf m_eFormStatus = AddStatus Then
'            cmdCheck.Value = False
            '列表新增一列
            vsSettlePrice.Rows = 2
        End If
    Else
        cboVehicleType.Visible = False
        lblVehicleType.Visible = False
        lbl2.Visible = True
        cboCompany.Visible = True
        txtRouteID.Enabled = False

        FillBusVSHead
        lbl1.Visible = True
        txtCompany.Visible = False
        lbl2.Caption = "车次代码"
        lbl2.Left = 180
        lbl2.Top = 960
        txtVehilce.Left = 1000
        txtVehilce.Top = 930
        lbl1.Left = 5370
        lbl1.Top = 990
        cboCompany.Left = 6270
        cboCompany.Top = 930
        txtVehilce.Width = 1515

        If m_eFormStatus = ModifyStatus Then
            txtVehilce.Text = m_szBus
            txtRouteID.Text = m_szRoute
            cboCompany.Text = m_szTransportCompany
            RefreshBusSettlePriceInfo
            txtVehilce.Enabled = False
            cboCompany.Enabled = False
            txtRouteID.Enabled = False
            chkBack.Enabled = False
            cmdOk.Enabled = True
            ReDim m_aszTemp(1 To 1, 1 To 2)
            m_aszTemp(1, 1) = m_szBus
            m_aszTemp(1, 2) = m_szBus

        ElseIf m_eFormStatus = AddStatus Then
'            cmdCheck.Value = False
            '列表新增一列
            vsSettlePrice.Rows = 2
        End If
    End If
    AlignHeadWidth Me.name, vsSettlePrice

End Sub


'公司结算价列表填充
Private Sub FillCompanyVSHead()
    vsSettlePrice.Cols = 14
'    VsSettlePrice.TextMatrix(0, 1) = "参运公司代码"
'    VsSettlePrice.TextMatrix(0, 2) = "车型代码"
    vsSettlePrice.TextMatrix(0, 1) = "上车站代码"
'    VsSettlePrice.TextMatrix(0, 4) = "线路代码"
    vsSettlePrice.TextMatrix(0, 2) = "站点代码"
    vsSettlePrice.TextMatrix(0, 3) = "里程"
    vsSettlePrice.TextMatrix(0, 4) = "通行费"
    vsSettlePrice.TextMatrix(0, 5) = "结算全价"
    vsSettlePrice.TextMatrix(0, 6) = "结算半价"
   vsSettlePrice.TextMatrix(0, 7) = "平分全价"
    vsSettlePrice.TextMatrix(0, 8) = "平分半价"
    vsSettlePrice.TextMatrix(0, 9) = "劳务费全价"
    vsSettlePrice.TextMatrix(0, 10) = "劳务费半价"
    vsSettlePrice.TextMatrix(0, 11) = "春运费全价"
    vsSettlePrice.TextMatrix(0, 12) = "春运费半价"

    vsSettlePrice.TextMatrix(0, 13) = "计算说明"

    '填充列选项
End Sub

'车辆结算价列表填充
Private Sub FillVehicleVSHead()
    vsSettlePrice.Cols = 14
'    VsSettlePrice.TextMatrix(0, 1) = "参运车辆代码"
    vsSettlePrice.TextMatrix(0, 1) = "上车站代码"
'    VsSettlePrice.TextMatrix(0, 3) = "线路代码"
    vsSettlePrice.TextMatrix(0, 2) = "站点代码"
    vsSettlePrice.TextMatrix(0, 3) = "里程"
    vsSettlePrice.TextMatrix(0, 4) = "通行费"
    vsSettlePrice.TextMatrix(0, 5) = "结算全价"
    vsSettlePrice.TextMatrix(0, 6) = "结算半价"
    vsSettlePrice.TextMatrix(0, 7) = "平分全价"
    vsSettlePrice.TextMatrix(0, 8) = "平分半价"
    vsSettlePrice.TextMatrix(0, 9) = "劳务费全价"
    vsSettlePrice.TextMatrix(0, 10) = "劳务费半价"
    vsSettlePrice.TextMatrix(0, 11) = "春运费全价"
    vsSettlePrice.TextMatrix(0, 12) = "春运费半价"
    vsSettlePrice.TextMatrix(0, 13) = "计算说明"

    '填充列选项

End Sub

'车次结算价列表填充
Private Sub FillBusVSHead()
    vsSettlePrice.Cols = 14
'    VsSettlePrice.TextMatrix(0, 1) = "参运车辆代码"
    vsSettlePrice.TextMatrix(0, 1) = "上车站代码"
'    VsSettlePrice.TextMatrix(0, 3) = "线路代码"
    vsSettlePrice.TextMatrix(0, 2) = "站点代码"
    vsSettlePrice.TextMatrix(0, 3) = "里程"
    vsSettlePrice.TextMatrix(0, 4) = "通行费"
    vsSettlePrice.TextMatrix(0, 5) = "结算全价"
    vsSettlePrice.TextMatrix(0, 6) = "结算半价"
    vsSettlePrice.TextMatrix(0, 7) = "平分全价"
    vsSettlePrice.TextMatrix(0, 8) = "平分半价"
    vsSettlePrice.TextMatrix(0, 9) = "劳务费全价"
    vsSettlePrice.TextMatrix(0, 10) = "劳务费半价"
    vsSettlePrice.TextMatrix(0, 11) = "春运费全价"
    vsSettlePrice.TextMatrix(0, 12) = "春运费半价"
    vsSettlePrice.TextMatrix(0, 13) = "计算说明"

    '填充列选项

End Sub

Private Sub txtObject_LostFocus()
'    如果为车次 , 则填充所有的该车次的公司
    Dim oBus As New Bus
    Dim nCount As Integer
    Dim aszCompany() As String
    Dim i As Integer
    On Error GoTo ErrorHandle
    If szTitle = "车次结算价" Then
        cboCompany.Clear
        cboCompany.AddItem ""
        oBus.Init g_oActiveUser
        oBus.Identify Trim(cboCompany.Text)
        aszCompany = oBus.GetAllCompany
        nCount = ArrayLength(aszCompany)
        For i = 1 To nCount
            cboCompany.AddItem MakeDisplayString(aszCompany(i, 1), aszCompany(i, 2))
        Next i

    End If
    Exit Sub
ErrorHandle:

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    SaveHeadWidth Me.name, vsSettlePrice
    Unload Me
    '刷新记录
End Sub



Private Sub tbSelect_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "ADD"
            vsSettlePrice.Rows = vsSettlePrice.Rows + 1
        Case "Delete"
            If vsSettlePrice.Rows > 1 And vsSettlePrice.Row <> 0 Then
               vsSettlePrice.RemoveItem vsSettlePrice.Row
            End If
    End Select
End Sub

Private Sub txtCompany_ButtonClick()
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectCompany()
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtCompany.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))
End Sub



Private Sub txtCompany_Change()
    If txtCompany.Text = "" Or txtVehilce.Text = "" Or txtRouteID.Text = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
    If txtCompany.Text <> "" And txtRouteID.Text <> "" Then
        FillVehicleType txtCompany.Text, txtRouteID.Text
    End If
End Sub

Private Sub txtRouteID_ButtonClick()
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    If chkBack.Value = vbUnchecked Then
        aszTemp = oShell.SelectRoute
    Else
        aszTemp = oShell.SelectBackRoute
    End If
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtRouteID.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))
    Fill

End Sub
Private Sub Fill()
    Dim atCompnayPrice() As TCompanySettlePrice
    Dim atVehiclePrice() As TVehcileSettlePrice
    Dim atBusPrice() As TBusSettlePrice
    Dim i As Integer
    vsSettlePrice.MergeCol(1) = True
    vsSettlePrice.MergeCells = flexMergeRestrictColumns
    If lblTitle = "公司结算价" Then
            FillVS
    Else
            FillVS

    End If
End Sub
Private Sub FillVS()
    On Error GoTo err
    Dim oRoute As Object
    Dim szRouteID As String, i As Integer
    Dim aszTemp() As String
    Dim szOldSellStation As String
    Dim szNewSellStation As String
    vsSettlePrice.MergeCol(1) = True
    vsSettlePrice.MergeCells = flexMergeRestrictColumns

    If txtRouteID.Text <> "" Then
        szRouteID = ResolveDisplay(txtRouteID.Text)
        If chkBack.Value = vbUnchecked Then
            Set oRoute = CreateObject("STSettle.Report")
        Else
            Set oRoute = CreateObject("STSettle.BackRoute")


        End If

        oRoute.Init g_oActiveUser
        If chkBack.Value = vbUnchecked Then

            aszTemp = oRoute.GetAllSectionInfo(szRouteID)
        Else
            oRoute.Identify szRouteID
            aszTemp = oRoute.GetAllSectionInfo
        End If
        vsSettlePrice.Rows = ArrayLength(aszTemp) + 1
        For i = 1 To ArrayLength(aszTemp)
            vsSettlePrice.Cell(flexcpText, i, 1) = MakeDisplayString(aszTemp(i, 3), aszTemp(i, 4))
            vsSettlePrice.TextMatrix(i, 2) = MakeDisplayString(aszTemp(i, 5), aszTemp(i, 6))
            vsSettlePrice.TextMatrix(i, 3) = aszTemp(i, 9)
        Next i
    End If

    For i = 1 To ArrayLength(aszTemp)
        vsSettlePrice.TextMatrix(i, 4) = 0
        vsSettlePrice.TextMatrix(i, 5) = 0
        vsSettlePrice.TextMatrix(i, 6) = 0
        vsSettlePrice.TextMatrix(i, 7) = 0
        vsSettlePrice.TextMatrix(i, 8) = 0
        vsSettlePrice.TextMatrix(i, 9) = 0
        vsSettlePrice.TextMatrix(i, 10) = 0
        vsSettlePrice.TextMatrix(i, 11) = 0
        vsSettlePrice.TextMatrix(i, 12) = 0
        vsSettlePrice.TextMatrix(i, 13) = ""
    Next i
    Exit Sub
err:
ShowErrorMsg
End Sub


Private Sub txtRouteID_Change()
    If txtCompany.Text = "" Or txtVehilce.Text = "" Or txtRouteID.Text = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
    If txtCompany.Text <> "" And txtRouteID.Text <> "" Then
        FillVehicleType txtCompany.Text, txtRouteID.Text
    End If
End Sub

Private Sub txtRouteID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Fill
    End If
End Sub

Private Sub txtVehilce_ButtonClick()
    Dim oShell As New STShell.CommDialog
    Dim rsTemp As Recordset
    Dim oBus As New Bus

    oShell.Init g_oActiveUser
    If lbl2.Caption = "车辆代码" Then
        m_aszTemp = oShell.SelectVehicle(, , , , , True)
        If ArrayLength(m_aszTemp) = 0 Then Exit Sub
        txtVehilce.Text = TeamToString(m_aszTemp, 2)
    ElseIf lbl2.Caption = "车次代码" Then
        m_aszTemp = oShell.SelectBus()
        If ArrayLength(m_aszTemp) = 0 Then Exit Sub
        txtVehilce.Text = Trim(m_aszTemp(1, 1))
        oBus.Init g_oActiveUser
        oBus.Identify m_aszTemp(1, 1)
        
        
'        Set rsTemp = oRoute.GetRouteID(Trim(m_aszTemp(1, 1)))
        txtRouteID.Text = MakeDisplayString(oBus.Route, oBus.RouteName)    ' MakeDisplayString(FormatDbValue(rsTemp!route_id), Trim(m_aszTemp(1, 3)))
    cmdOk.Enabled = True
    Fill

    End If

End Sub

Private Sub txtVehilce_Change()
    If txtCompany.Text = "" Or txtVehilce.Text = "" Or txtRouteID.Text = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If

End Sub

Private Sub txtVehilce_LostFocus()
    '如果为车次,则填充所有的该车次的公司
    Dim oBus As New Bus
    Dim nCount As Integer
    Dim aszCompany() As String
    Dim i As Integer
    On Error GoTo ErrorHandle
    If lbl2.Caption = "车次代码" Then
        cboCompany.Clear
        cboCompany.AddItem ""
        oBus.Init g_oActiveUser
        oBus.Identify txtVehilce.Text
        aszCompany = oBus.GetAllCompany
        nCount = ArrayLength(aszCompany)
        For i = 1 To nCount
            cboCompany.AddItem MakeDisplayString(aszCompany(i, 1), aszCompany(i, 2))
        Next i

    End If
    Exit Sub
ErrorHandle:
End Sub

Private Sub VsSettlePrice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '判断是否有重复项
    If IsNumeric(Trim(vsSettlePrice.TextMatrix(vsSettlePrice.Row, 4))) = False Then
        vsSettlePrice.TextMatrix(vsSettlePrice.Row, 4) = 0
    End If
    If IsNumeric(Trim(vsSettlePrice.TextMatrix(vsSettlePrice.Row, 6))) = False Then
        vsSettlePrice.TextMatrix(vsSettlePrice.Row, 6) = 0
    End If
    If IsNumeric(Trim(vsSettlePrice.TextMatrix(vsSettlePrice.Row, 8))) = False Then
        vsSettlePrice.TextMatrix(vsSettlePrice.Row, 8) = 0
    End If
    If IsNumeric(Trim(vsSettlePrice.TextMatrix(vsSettlePrice.Row, 10))) = False Then
        vsSettlePrice.TextMatrix(vsSettlePrice.Row, 10) = 0
    End If
    If IsNumeric(Trim(vsSettlePrice.TextMatrix(vsSettlePrice.Row, 5))) = False Then
        vsSettlePrice.TextMatrix(vsSettlePrice.Row, 5) = 0
    ElseIf IsNumeric(Trim(vsSettlePrice.TextMatrix(vsSettlePrice.Row, 5))) = True And Col = 5 Then
        vsSettlePrice.TextMatrix(vsSettlePrice.Row, 6) = (vsSettlePrice.TextMatrix(vsSettlePrice.Row, 5) / 2)
    End If

    If IsNumeric(Trim(vsSettlePrice.TextMatrix(vsSettlePrice.Row, 7))) = False Then
        vsSettlePrice.TextMatrix(vsSettlePrice.Row, 7) = 0
    ElseIf IsNumeric(Trim(vsSettlePrice.TextMatrix(vsSettlePrice.Row, 7))) = True And Col = 7 Then
        vsSettlePrice.TextMatrix(vsSettlePrice.Row, 8) = (vsSettlePrice.TextMatrix(vsSettlePrice.Row, 7) / 2)
    End If

    If IsNumeric(Trim(vsSettlePrice.TextMatrix(vsSettlePrice.Row, 9))) = False Then
        vsSettlePrice.TextMatrix(vsSettlePrice.Row, 9) = 0
    ElseIf IsNumeric(Trim(vsSettlePrice.TextMatrix(vsSettlePrice.Row, 9))) = True And Col = 9 Then
        vsSettlePrice.TextMatrix(vsSettlePrice.Row, 10) = (vsSettlePrice.TextMatrix(vsSettlePrice.Row, 9) / 2)
    End If

    If IsNumeric(Trim(vsSettlePrice.TextMatrix(vsSettlePrice.Row, 11))) = False Then
        vsSettlePrice.TextMatrix(vsSettlePrice.Row, 11) = 0
    ElseIf IsNumeric(Trim(vsSettlePrice.TextMatrix(vsSettlePrice.Row, 11))) = True And Col = 11 Then
        vsSettlePrice.TextMatrix(vsSettlePrice.Row, 12) = (vsSettlePrice.TextMatrix(vsSettlePrice.Row, 11) / 2)
    End If

End Sub

Private Sub RefreshCompanySettlePriceInfo()
    On Error GoTo err
    Dim i As Integer
    Dim rsTemp As Recordset

    Dim atCompanySettlePrice() As TCompanySettlePrice

    vsSettlePrice.MergeCol(1) = True
    vsSettlePrice.MergeCells = flexMergeRestrictColumns

'    szRouteID = ResolveDisplay(txtRouteID.Text)


    m_oReport.Init g_oActiveUser
    atCompanySettlePrice = m_oReport.GetCompanySettlePriceLst(ResolveDisplay(txtCompany.Text), ResolveDisplay(cboVehicleType.Text), ResolveDisplay(txtRouteID.Text))

    vsSettlePrice.Rows = ArrayLength(atCompanySettlePrice) + 1
    For i = 1 To ArrayLength(atCompanySettlePrice)
        vsSettlePrice.TextMatrix(i, 1) = MakeDisplayString(atCompanySettlePrice(i).SellStationID, atCompanySettlePrice(i).SellStationName)
        vsSettlePrice.TextMatrix(i, 2) = MakeDisplayString(atCompanySettlePrice(i).StationID, atCompanySettlePrice(i).StationName)
        vsSettlePrice.TextMatrix(i, 3) = atCompanySettlePrice(i).Mileage
        vsSettlePrice.TextMatrix(i, 4) = atCompanySettlePrice(i).PassCharge
        vsSettlePrice.TextMatrix(i, 5) = atCompanySettlePrice(i).SettlefullPrice
        vsSettlePrice.TextMatrix(i, 6) = atCompanySettlePrice(i).SettleHalfPrice
        vsSettlePrice.TextMatrix(i, 7) = atCompanySettlePrice(i).HalveFullPrice
        vsSettlePrice.TextMatrix(i, 8) = atCompanySettlePrice(i).HalveHalfPrice
        vsSettlePrice.TextMatrix(i, 9) = atCompanySettlePrice(i).ServiceFullPrice
        vsSettlePrice.TextMatrix(i, 10) = atCompanySettlePrice(i).ServiceHalfPrice
        vsSettlePrice.TextMatrix(i, 11) = atCompanySettlePrice(i).SpringFullPrice
        vsSettlePrice.TextMatrix(i, 12) = atCompanySettlePrice(i).SpringHalfPrice
        vsSettlePrice.TextMatrix(i, 13) = atCompanySettlePrice(i).Annotation

    Next i
    Exit Sub
err:
    ShowErrorMsg
End Sub


Public Sub RefreshVehicleSettlePriceInfo()

    On Error GoTo err
    Dim i As Integer
    Dim rsTemp As Recordset

    Dim atVehicleSettlePrice() As TVehcileSettlePrice

    vsSettlePrice.MergeCol(1) = True
    vsSettlePrice.MergeCells = flexMergeRestrictColumns

'    szRouteID = ResolveDisplay(txtRouteID.Text)


    m_oReport.Init g_oActiveUser
    atVehicleSettlePrice = m_oReport.GetVehicleSettlePriceLst(ResolveDisplay(txtVehilce.Text), , ResolveDisplay(txtRouteID.Text))

    vsSettlePrice.Rows = ArrayLength(atVehicleSettlePrice) + 1
    For i = 1 To ArrayLength(atVehicleSettlePrice)
        vsSettlePrice.Cell(flexcpText, i, 1) = MakeDisplayString(atVehicleSettlePrice(i).SellStationID, atVehicleSettlePrice(i).SellStationName)
        vsSettlePrice.TextMatrix(i, 2) = MakeDisplayString(atVehicleSettlePrice(i).StationID, atVehicleSettlePrice(i).StationName)
        vsSettlePrice.TextMatrix(i, 3) = atVehicleSettlePrice(i).Mileage
        vsSettlePrice.TextMatrix(i, 4) = atVehicleSettlePrice(i).PassCharge
        vsSettlePrice.TextMatrix(i, 5) = atVehicleSettlePrice(i).SettlefullPrice
        vsSettlePrice.TextMatrix(i, 6) = atVehicleSettlePrice(i).SettleHalfPrice
        vsSettlePrice.TextMatrix(i, 7) = atVehicleSettlePrice(i).HalveFullPrice
        vsSettlePrice.TextMatrix(i, 8) = atVehicleSettlePrice(i).HalveHalfPrice
        vsSettlePrice.TextMatrix(i, 9) = atVehicleSettlePrice(i).ServiceFullPrice
        vsSettlePrice.TextMatrix(i, 10) = atVehicleSettlePrice(i).ServiceHalfPrice
        vsSettlePrice.TextMatrix(i, 11) = atVehicleSettlePrice(i).SpringFullPrice
        vsSettlePrice.TextMatrix(i, 12) = atVehicleSettlePrice(i).SpringHalfPrice
        vsSettlePrice.TextMatrix(i, 13) = atVehicleSettlePrice(i).Annotation
    Next i

    Exit Sub
err:
    ShowErrorMsg
End Sub

Public Sub RefreshBusSettlePriceInfo()

    On Error GoTo err
    Dim i As Integer
    Dim rsTemp As Recordset

    Dim atBusSettlePrice() As TBusSettlePrice

    vsSettlePrice.MergeCol(1) = True
    vsSettlePrice.MergeCells = flexMergeRestrictColumns

    m_oReport.Init g_oActiveUser
    atBusSettlePrice = m_oReport.GetBusSettlePriceLst(ResolveDisplay(txtVehilce.Text), ResolveDisplay(txtCompany.Text))

    vsSettlePrice.Rows = ArrayLength(atBusSettlePrice) + 1
    For i = 1 To ArrayLength(atBusSettlePrice)
        vsSettlePrice.Cell(flexcpText, i, 1) = MakeDisplayString(atBusSettlePrice(i).SellStationID, atBusSettlePrice(i).SellStationName)
        vsSettlePrice.TextMatrix(i, 2) = MakeDisplayString(atBusSettlePrice(i).StationID, atBusSettlePrice(i).StationName)
        vsSettlePrice.TextMatrix(i, 3) = atBusSettlePrice(i).Mileage
        vsSettlePrice.TextMatrix(i, 4) = atBusSettlePrice(i).PassCharge
        vsSettlePrice.TextMatrix(i, 5) = atBusSettlePrice(i).SettlefullPrice
        vsSettlePrice.TextMatrix(i, 6) = atBusSettlePrice(i).SettleHalfPrice
        vsSettlePrice.TextMatrix(i, 7) = atBusSettlePrice(i).HalveFullPrice
        vsSettlePrice.TextMatrix(i, 8) = atBusSettlePrice(i).HalveHalfPrice
        vsSettlePrice.TextMatrix(i, 9) = atBusSettlePrice(i).ServiceFullPrice
        vsSettlePrice.TextMatrix(i, 10) = atBusSettlePrice(i).ServiceHalfPrice
        vsSettlePrice.TextMatrix(i, 11) = atBusSettlePrice(i).SpringFullPrice
        vsSettlePrice.TextMatrix(i, 12) = atBusSettlePrice(i).SpringHalfPrice
        vsSettlePrice.TextMatrix(i, 13) = atBusSettlePrice(i).Annotation
    Next i

    Exit Sub
err:
    ShowErrorMsg
End Sub

Private Sub VsSettlePrice_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    If NewColSel < 3 Then
        vsSettlePrice.Editable = flexEDNone
    Else
        vsSettlePrice.Editable = flexEDKbdMouse
    End If
End Sub
