VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VsFlex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form FrmProtocolFormula 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设定车次拆算公式"
   ClientHeight    =   4875
   ClientLeft      =   2940
   ClientTop       =   3105
   ClientWidth     =   7860
   Icon            =   "FrmProtocolFormula.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   Begin RTComctl3.CoolButton CmdOk 
      Height          =   345
      Left            =   5190
      TabIndex        =   2
      Top             =   4410
      Width           =   1155
      _ExtentX        =   2037
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
      MICON           =   "FrmProtocolFormula.frx":0442
      PICN            =   "FrmProtocolFormula.frx":045E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "批量设置(&M)"
      TabPicture(0)   =   "FrmProtocolFormula.frx":07F8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbProtocolName"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "单项设置(&S)"
      TabPicture(1)   =   "FrmProtocolFormula.frx":0814
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   30
         TabIndex        =   12
         Top             =   330
         Width           =   7515
         Begin VB.ComboBox cboSplitObject 
            Height          =   315
            ItemData        =   "FrmProtocolFormula.frx":0830
            Left            =   1200
            List            =   "FrmProtocolFormula.frx":0840
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   150
            Width           =   2265
         End
         Begin MSComctlLib.ListView LvInfo 
            Height          =   2685
            Left            =   1170
            TabIndex        =   13
            Top             =   930
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   4736
            View            =   3
            Arrange         =   1
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            Icons           =   "ImageList1"
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin FText.asFlatTextBox asFlatTextBox1 
            Height          =   315
            Left            =   5100
            TabIndex        =   19
            Top             =   540
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   556
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
            Text            =   "asFlatTextBox1"
            ButtonBackColor =   -2147483633
         End
         Begin RTComctl3.TextButtonBox txtObjectID 
            Height          =   315
            Left            =   5100
            TabIndex        =   21
            Top             =   150
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RTComctl3.TextButtonBox txtProtocol 
            Height          =   315
            Left            =   1200
            TabIndex        =   22
            ToolTipText     =   "协议代码"
            Top             =   540
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   4
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "对应列表(&L):"
            Height          =   180
            Left            =   90
            TabIndex        =   18
            Top             =   1050
            Width           =   1080
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "对象代码(&I):"
            Height          =   240
            Left            =   3690
            TabIndex        =   17
            Top             =   240
            Width           =   990
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "对象类型(&T):"
            Height          =   255
            Left            =   90
            TabIndex        =   16
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "协议名称(&N):"
            Height          =   270
            Left            =   3720
            TabIndex        =   15
            Top             =   570
            Width           =   1080
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "协议号(&P):"
            Height          =   180
            Left            =   90
            TabIndex        =   14
            Top             =   630
            Width           =   900
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   3705
         Left            =   -74970
         TabIndex        =   4
         Top             =   330
         Width           =   7515
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Height          =   855
            Left            =   150
            TabIndex        =   5
            Top             =   90
            Width           =   7275
            Begin VB.ComboBox cboType 
               Height          =   300
               ItemData        =   "FrmProtocolFormula.frx":0860
               Left            =   1200
               List            =   "FrmProtocolFormula.frx":0870
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   330
               Width           =   1305
            End
            Begin RTComctl3.TextButtonBox txtID 
               Height          =   330
               Left            =   3690
               TabIndex        =   8
               Top             =   330
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   582
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin RTComctl3.CoolButton cmdFind 
               Height          =   345
               Left            =   5910
               TabIndex        =   6
               Top             =   330
               Width           =   1155
               _ExtentX        =   0
               _ExtentY        =   0
               BTYPE           =   3
               TX              =   "查询(&F)"
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
               MICON           =   "FrmProtocolFormula.frx":0890
               PICN            =   "FrmProtocolFormula.frx":08AC
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "对象类型(&T):"
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   390
               Width           =   1125
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "对象代码(&I):"
               Height          =   180
               Left            =   2610
               TabIndex        =   9
               Top             =   390
               Width           =   1080
            End
         End
         Begin VSFlex7LCtl.VSFlexGrid VSSelbus 
            Height          =   2565
            Left            =   150
            TabIndex        =   11
            Top             =   1020
            Width           =   7275
            _cx             =   12832
            _cy             =   4524
            _ConvInfo       =   -1
            Appearance      =   2
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
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   14737632
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmProtocolFormula.frx":0C46
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
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5040
         Top             =   -150
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProtocolFormula.frx":0CBE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lbProtocolName 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   4440
         TabIndex        =   3
         Top             =   930
         Width           =   90
      End
   End
   Begin RTComctl3.CoolButton CmdCancel 
      Height          =   345
      Left            =   6480
      TabIndex        =   0
      Top             =   4410
      Width           =   1155
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "FrmProtocolFormula.frx":1058
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
      BorderColor     =   &H8000000C&
      Index           =   0
      X1              =   150
      X2              =   7650
      Y1              =   4290
      Y2              =   4290
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      Index           =   2
      X1              =   150
      X2              =   7650
      Y1              =   4305
      Y2              =   4305
   End
   Begin VB.Menu Pmnu_FormualSet 
      Caption         =   "弹出设定"
      Visible         =   0   'False
      Begin VB.Menu Pmnu_Add 
         Caption         =   "新增对象(&A)"
      End
      Begin VB.Menu Pmnu_Delte 
         Caption         =   "删除对象(&D)"
      End
   End
End
Attribute VB_Name = "FrmProtocolFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim m_oProtocol As New Protocol
'Dim m_szProtocol As String
'Private m_vaExtSelInfo As Variant
'
'
'Private Sub cboSplitObject_Click()
'   LvInfo.ColumnHeaders.Clear
'   LvInfo.ColumnHeaders.Add , , cboSplitObject.Text & "代码", LvInfo.Width / 2 - 100
'   LvInfo.ColumnHeaders.Add , , cboSplitObject.Text & "名称", LvInfo.Width / 2, lvwColumnCenter
'   LvInfo.ListItems.Clear
'End Sub
'
'Private Sub cmdCancel_Click()
'   Unload Me
'End Sub
'
'Private Sub cmdFind_Click()
'    Dim szStemp() As String
'    Dim nCount As Integer, i As Integer, j As Integer
'    Dim m_szObjectID As String
'    m_szObjectID = ResolveDisplay(txtID.Text)
'    With m_oProtocol
'        ShowTBInfo "查询车辆协议..."
'        SetBusy
'            Select Case cboType.ListIndex
'                Case 0                      'Vehicle
'                    szStemp = .GetAccordVehicleProtocol(, , , , , m_szObjectID)
'                Case 1                      'Vehicle Type
'                    szStemp = .GetAccordVehicleProtocol(, m_szObjectID)
'                Case 2                      'company
'                    szStemp = .GetAccordVehicleProtocol(, , , m_szObjectID)
'                Case 3                      'Owner
'                    szStemp = .GetAccordVehicleProtocol(, , m_szObjectID)
'            End Select
'            nCount = ArrayLength(szStemp)
'            VSSelbus.Rows = nCount + 1
'            For j = 1 To nCount
'                VSSelbus.TextMatrix(j, 0) = j
'                VSSelbus.TextMatrix(j, 1) = szStemp(j, 1)
'                VSSelbus.TextMatrix(j, 2) = szStemp(j, 2)
'                VSSelbus.TextMatrix(j, 3) = szStemp(j, 3)
'                VSSelbus.TextMatrix(j, 4) = szStemp(j, 4)
'            Next j
'        SetNormal
'        ShowTBInfo
'    End With
'End Sub
'
'Private Sub cmdOk_Click()
'Dim szStemp() As String
'Dim nCount As Integer, i As Integer, j As Integer
'On Error GoTo ErrorHandle
'With m_oProtocol
'   Select Case SSTab1.Tab
'       Case 0
'        If LvInfo.ListItems.Count <> 0 Then
'           ShowTBInfo "保存车辆协议..."
'           SetBusy
'           For i = 1 To LvInfo.ListItems.Count
'            Select Case cboSplitObject.ListIndex
'                Case 0                      'Vehicle
'                    szStemp = .GetVehicleInfo(, , , , , LvInfo.ListItems.Item(i).Text)
'                Case 1                      'Vehicle Type
'                    szStemp = .GetVehicleInfo(, LvInfo.ListItems.Item(i).Text)
'                Case 2                      'company
'                    szStemp = .GetVehicleInfo(, , , LvInfo.ListItems.Item(i).Text)
'                Case 3                      'Owner
'                    szStemp = .GetVehicleInfo(, , LvInfo.ListItems.Item(i).Text)
'            End Select
'            nCount = ArrayLength(szStemp)
'            For j = 1 To nCount
'                .SaveVehicleProtocol szStemp(j, 1), szStemp(j, 2), txtProtocol.Text
'                ShowTBInfo , LvInfo.ListItems.Count, i, True
'            Next j
'            Next i
'            SetNormal
'            ShowTBInfo
'          End If
'      Case 1
'         SetBusy
'         ShowTBInfo "保存车次协议..."
'         nCount = VSSelbus.Rows - 1
'         For i = 1 To nCount
'            If VSSelbus.Cell(flexcpBackColor, i, 1, i, VSSelbus.Cols - 1) = vbBlue Then
'                .SaveVehicleProtocol VSSelbus.TextMatrix(i, 1), VSSelbus.TextMatrix(i, 3), VSSelbus.TextMatrix(i, 4)
'                VSSelbus.Cell(flexcpBackColor, i, 1, i, VSSelbus.Cols - 1) = vbWhite
'            End If
'         Next i
'         SetNormal
'         ShowTBInfo
'        End Select
'    End With
'    Exit Sub
'ErrorHandle:
'    SetNormal
'    ShowErrorU
'End Sub
'
'
'
'Private Sub Form_Load()
'  cboSplitObject.ListIndex = 0
'  SelBusSet
'End Sub
'
'Private Sub LvInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'  Static s_nUpColumn As Integer
'  LvInfo.SortKey = ColumnHeader.Index - 1
'  If s_nUpColumn = ColumnHeader.Index - 1 Then
'    LvInfo.SortOrder = lvwDescending
'    s_nUpColumn = ColumnHeader.Index
'  Else
'    LvInfo.SortOrder = lvwAscending
'    s_nUpColumn = ColumnHeader.Index - 1
'  End If
'  LvInfo.Sorted = True
'End Sub
'
'Private Sub LvInfo_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    txtObjectID.Text = Item.Text & "[" & Item.SubItems(1) & "]"
'End Sub
'
'Private Sub LvInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'      If Button = 2 Then
'          PopupMenu Pmnu_FormualSet
'       End If
'End Sub
'
'Private Sub pmnu_Add_Click()
'    Dim aszTemp() As String
'    Dim oOpen As New STShell.CommDialog
'    Dim Item As ListItem
'    Dim bExist  As Boolean
'    Dim i As Integer
'    Dim nCount As Integer
'    Dim j As Integer
'    On Error Resume Next
'    bExist = False
'    oOpen.Init g_oActiveUser
'    Select Case cboSplitObject.ListIndex
'           Case 0 '车辆
'                aszTemp = oOpen.SelectVehicle(, True)
'           Case 1 '车型
'                aszTemp = oOpen.SelectVehicleType(True)
'           Case 2 '参运公司
'                aszTemp = oOpen.SelectCompany(True)
'           Case 3 '车主
'                aszTemp = oOpen.SelectOwner(, True)
'    End Select
'    nCount = LvInfo.ListItems.Count
'    If ArrayLength(aszTemp) < 1 Then Exit Sub
'    For j = 1 To ArrayLength(aszTemp)
'       For i = 1 To nCount
'          If aszTemp(j, 1) = LvInfo.ListItems.Item(i).Text Then
'              bExist = True
'              Exit For
'          Else
'              bExist = False
'          End If
'       Next i
'       If bExist = False Then
'         LvInfo.ListItems.Add(, , aszTemp(j, 1), , 1).ListSubItems.Add
'         If cboSplitObject.ListIndex = 0 Then
'            LvInfo.ListItems(LvInfo.ListItems.Count).ListSubItems(1).Text = "车次" & aszTemp(j, 2)
'         Else
'            LvInfo.ListItems(LvInfo.ListItems.Count).ListSubItems(1).Text = aszTemp(j, 2)
'         End If
'        End If
'    Next j
'End Sub
'
'Private Sub pmnu_Delte_Click()
'  If LvInfo.ListItems.Count > 0 Then
'       LvInfo.ListItems.Remove (LvInfo.SelectedItem.Index)
'  End If
'End Sub
'
'Private Sub txtID_Click()
'    Dim aszTemp() As String
'    Dim oOpen As New STShell.CommDialog
'    Dim i As Integer
'    On Error Resume Next
'    oOpen.Init g_oActiveUser
'    Select Case cboType.ListIndex
'           Case 0 '车辆
'                aszTemp = oOpen.SelectVehicle(, False)
'           Case 1 '车型
'                aszTemp = oOpen.SelectVehicleType(False)
'           Case 2 '参运公司
'                aszTemp = oOpen.SelectCompany(False)
'           Case 3 '车主
'                aszTemp = oOpen.SelectOwner(, False)
'    End Select
'    If ArrayLength(aszTemp) < 1 Then Exit Sub
'    txtID.Text = Trim(aszTemp(1, 1)) & "[" & Trim(aszTemp(1, 2)) & "]"
'
'End Sub
'
'Private Sub txtObjectID_Click()
'    Dim aszTemp() As String
'    Dim oOpen As New STShell.CommDialog
'    Dim Item As ListItem
'    Dim i As Integer
'    On Error Resume Next
'    oOpen.Init g_oActiveUser
'    Select Case cboSplitObject.ListIndex
'           Case 0 '车辆
'                aszTemp = oOpen.SelectVehicle(, , , , , True)
'           Case 1 '车型
'                aszTemp = oOpen.SelectVehicleType(True)
'           Case 2 '参运公司
'                aszTemp = oOpen.SelectCompany(True)
'           Case 3 '车主
'                aszTemp = oOpen.SelectOwner(, True)
'    End Select
'    If ArrayLength(aszTemp) < 1 Then Exit Sub
'    LvInfo.ListItems.Clear
'    If cboSplitObject.ListIndex = 0 Then
'        txtObjectID.Text = Trim(aszTemp(1, 1)) & "[车次" & Trim(aszTemp(1, 2)) & "]"
'        For i = 1 To ArrayLength(aszTemp)
'            Set Item = LvInfo.ListItems.Add(, , Trim(aszTemp(i, 1)), 1, 1)
'            Item.SubItems(1) = "车次" & Trim(aszTemp(i, 2))
'        Next i
'    Else
'        txtObjectID.Text = Trim(aszTemp(1, 1)) & "[" & Trim(aszTemp(1, 2)) & "]"
'         For i = 1 To ArrayLength(aszTemp)
'            Set Item = LvInfo.ListItems.Add(, , Trim(aszTemp(i, 1)), 1, 1)
'            Item.SubItems(1) = Trim(aszTemp(i, 2))
'        Next i
'    End If
'
'
'End Sub
'Private Sub txtProtocol_Click()
'    Dim oShell As New STShell.CommDialog
'    Dim aszTemp() As String
'
'    oShell.Init g_oActiveUser
'    aszTemp = oShell.SelectProtocol
'    If ArrayLength(aszTemp) > 0 Then
'        txtProtocol.Text = aszTemp(1, 1)
'        lbProtocolName.Caption = aszTemp(1, 2)
'    End If
'    Set oShell = Nothing
''
''    frmSelectProtocol.Show vbModal
''    txtProtocol.Text = frmSelectProtocol.m_szProtocolID
''    lbProtocolName.Caption = frmSelectProtocol.m_szProtocolName
'End Sub
'Public Function SelBusSet(Optional ProjectID As String = "") As String()
'  Dim i As Integer
'     VSSelbus.Cols = 5
'     VSSelbus.ColWidth(0) = 450
'     VSSelbus.ColWidth(1) = 800
'     VSSelbus.ColWidth(2) = 800
'     VSSelbus.ColWidth(3) = 1600
''     VSSelbus.ColWidth(4) = 1600
'     VSSelbus.ColWidth(4) = 1000
'     VSSelbus.TextMatrix(0, 0) = "序号"
'     VSSelbus.TextMatrix(0, 1) = "车辆代码"
'     VSSelbus.TextMatrix(0, 2) = "车主姓名"
'     VSSelbus.TextMatrix(0, 3) = "车辆牌号"
''     VSSelbus.TextMatrix(0, 4) = "车次线路"
'     VSSelbus.TextMatrix(0, 4) = "拆算协议"
'     VSSelbus.ColAlignment(0) = flexAlignLeftCenter
'     VSSelbus.ColAlignment(1) = flexAlignLeftCenter
'     For i = 2 To 4
'        VSSelbus.FixedAlignment(i) = flexAlignCenterCenter
'        VSSelbus.ColAlignment(i) = flexAlignGeneral
'     Next i
'End Function
'
'Private Sub AssertActiveUserIsNotThing()
'    'If m_oActiveUser Is Nothing Then ShowError ERR_NoActiveUser
'    If g_oActiveUser Is Nothing Then err.Raise ERR_ActiveUser, , "还未设置活动用户对象"
'End Sub
'
'Private Function RefreshBus()        '得到所有车次的协议
'    Dim oProject As New Protocol
'    Dim nDataCount As Integer, i As Integer
'    Dim liTemp As ListItem
'    Dim szTemp() As String
'    ShowTBInfo "读取车次协议..."
'    SetBusy
'    szTemp = oProject.GetALLVehicleProtocol
'    nDataCount = ArrayLength(szTemp)
'    VSSelbus.Rows = 1 + nDataCount
'    For i = 1 To nDataCount
'        VSSelbus.TextMatrix(i, 0) = i
'        VSSelbus.TextMatrix(i, 1) = RTrim(szTemp(i, 1))
'        VSSelbus.TextMatrix(i, 2) = RTrim(szTemp(i, 2))
'        VSSelbus.TextMatrix(i, 3) = RTrim(szTemp(i, 4))
'        ShowTBInfo , nDataCount, i, True
'    Next
'    SetNormal
'    ShowTBInfo
'End Function
'
'Private Sub VSSelbus_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    If Col = 1 Or Col = 2 Or Col = 3 Then
'        Cancel = True
'    ElseIf Col = 4 Then
'        Cancel = False
'    End If
'End Sub
'
'Private Sub VSSelbus_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'    Dim oShell As New STShell.CommDialog
'    Dim aszTemp() As String
'
'    oShell.Init g_oActiveUser
'    aszTemp = oShell.SelectProtocol
'    If ArrayLength(aszTemp) > 0 Then
'        If VSSelbus.TextMatrix(Row, 4) <> aszTemp(1, 1) Then
'            VSSelbus.Cell(flexcpBackColor, Row, 1, Row, VSSelbus.Cols - 1) = vbBlue
'            VSSelbus.TextMatrix(Row, 4) = aszTemp(1, 1)
'        End If
'     End If
'     Set oShell = Nothing
'
'End Sub
'Private Sub GetReport()
''    MDIFinance.CellExport.ExportVSFlexGrid VSSelbus
'End Sub
'
'
'
'Private Sub VSSelbus_DblClick()
'    If VSSelbus.TextMatrix(VSSelbus.Row, 1) <> "" Then
'        frmEspecialProtocol.OldFormularID = VSSelbus.TextMatrix(VSSelbus.Row, 1)
'        frmEspecialProtocol.OldProtocolID = VSSelbus.TextMatrix(VSSelbus.Row, 4)
'        frmEspecialProtocol.Show vbModal
'    End If
'End Sub
'
'
'Private Sub VSSelbus_KeyDown(KeyCode As Integer, Shift As Integer)
'If Shift = 2 And KeyCode = vbKeyV Then
'   PrintView
'ElseIf Shift = 2 And KeyCode = vbKeyP Then
'   Printa
'ElseIf Shift = 2 And KeyCode = vbKeyE Then
'   ExportFile
'ElseIf Shift = 2 And KeyCode = vbKeyB Then
'   ExportAndOpen
'End If
'
'End Sub
'
'Private Sub VSSelbus_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'       KeyAscii = 0
'End Sub
'
'Private Sub VSSelbus_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    Dim i As Integer
'    If Button = 1 And VSSelbus.MouseRow = 0 Then
'        VSSelbus.Sort = flexSortGenericAscending
'    End If
'End Sub
'Private Sub Printa()
''    GetReport
''     MDIFinance.CellExport.PrintEx False
'End Sub
'Private Sub PrintView()
''    GetReport
''    MDIFinance.CellExport.PrintPreview True
'End Sub
'Private Sub ExportFile()
''    GetReport
''    MDIFinance.CellExport.ExportFile
'End Sub
'Private Sub ExportAndOpen()
''    GetReport
''    MDIFinance.CellExport.ExportFile True
'End Sub
Private Sub Label9_Click()

End Sub
