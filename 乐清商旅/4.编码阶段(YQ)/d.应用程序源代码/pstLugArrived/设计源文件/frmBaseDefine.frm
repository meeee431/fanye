VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmBaseDefine 
   Caption         =   "常用定义"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   Icon            =   "frmBaseDefine.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11475
   ScaleWidth      =   23880
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3120
      Top             =   1110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseDefine.frx":038A
            Key             =   "cf"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseDefine.frx":0924
            Key             =   "close"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseDefine.frx":0EBE
            Key             =   "open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseDefine.frx":1458
            Key             =   "cl"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox ptMain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5715
      Left            =   0
      ScaleHeight     =   5715
      ScaleWidth      =   9915
      TabIndex        =   4
      Top             =   1170
      Width           =   9915
      Begin VB.PictureBox ptLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4725
         Left            =   0
         ScaleHeight     =   4725
         ScaleWidth      =   2430
         TabIndex        =   7
         Top             =   0
         Width           =   2430
         Begin MSComctlLib.TreeView tvItemTree 
            Height          =   4635
            Left            =   60
            TabIndex        =   8
            Top             =   420
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   8176
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   353
            LabelEdit       =   1
            Style           =   7
            FullRowSelect   =   -1  'True
            ImageList       =   "ImageList1"
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.Label lblLeftTitle 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "基本信息列表"
            Height          =   180
            Left            =   600
            TabIndex        =   10
            Top             =   60
            Width           =   1080
         End
         Begin VB.Image imgLeftBar 
            Height          =   300
            Left            =   0
            Picture         =   "frmBaseDefine.frx":17F2
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2430
         End
      End
      Begin VB.PictureBox ptRight 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4695
         Left            =   3210
         ScaleHeight     =   4695
         ScaleWidth      =   6735
         TabIndex        =   6
         Top             =   1950
         Width           =   6735
         Begin VSFlex7LCtl.VSFlexGrid vsBaseInfo 
            Height          =   1755
            Left            =   -390
            TabIndex        =   9
            Top             =   360
            Width           =   6105
            _cx             =   10769
            _cy             =   3096
            _ConvInfo       =   -1
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   14737632
            ForeColorFixed  =   0
            BackColorSel    =   -2147483635
            ForeColorSel    =   16744576
            BackColorBkg    =   16777215
            BackColorAlternate=   16777215
            GridColor       =   12632256
            GridColorFixed  =   12632256
            TreeColor       =   16777215
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   0
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   5
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   5
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
            ExplorerBar     =   1
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   0   'False
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            ComboSearch     =   2
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   16777215
            ForeColorFrozen =   0
            WallPaperAlignment=   9
         End
      End
      Begin RTComctl3.Spliter Spliter1 
         Height          =   375
         Left            =   2550
         TabIndex        =   5
         Top             =   2820
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   661
         BackColor       =   12677235
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SelectColor     =   16744576
      End
   End
   Begin VB.PictureBox ptTop 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   0
      ScaleHeight     =   885
      ScaleWidth      =   13755
      TabIndex        =   0
      Top             =   30
      Width           =   13755
      Begin RTComctl3.CoolButton cmdAddNew 
         Height          =   315
         Left            =   7440
         TabIndex        =   1
         Top             =   480
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "新增(&A)"
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
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBaseDefine.frx":288E
         PICN            =   "frmBaseDefine.frx":28AA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdDel 
         Height          =   315
         Left            =   8790
         TabIndex        =   2
         Top             =   480
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "删除(&D)"
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
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBaseDefine.frx":2E44
         PICN            =   "frmBaseDefine.frx":2E60
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdSave 
         Height          =   315
         Left            =   10140
         TabIndex        =   3
         Top             =   480
         Width           =   1275
         _ExtentX        =   2249
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
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBaseDefine.frx":33FA
         PICN            =   "frmBaseDefine.frx":3416
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image Image2 
         Height          =   330
         Left            =   240
         Picture         =   "frmBaseDefine.frx":37B0
         Top             =   300
         Width           =   1155
      End
      Begin VB.Image Image1 
         Height          =   1275
         Left            =   0
         Picture         =   "frmBaseDefine.frx":3A56
         Top             =   0
         Width           =   2010
      End
   End
End
Attribute VB_Name = "frmBaseDefine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mnItemIndex As Integer
Private maszDeleteID() As String     '删除的ID号
Private maszDeleteUnitID() As String
Private malNewRows() As Long     '新增的项目
Private malEditRows() As Long    '修改的项目
Private mbModified As Boolean
Private mszUnitComboString As String

Private Sub AddTreeItem()
    '"cl"所指是码表

    Dim oNode As Node
    Set oNode = tvItemTree.Nodes.Add(, , "comm_define", "常用定义", "close", "open")
    oNode.Expanded = True

    Set oNode = tvItemTree.Nodes.Add("comm_define", tvwChild, "code_packname", "行包名称", "cf")
    oNode.Tag = EDT_PackageName
    Set oNode = tvItemTree.Nodes.Add("comm_define", tvwChild, "code_packtype", "包装类型", "cf")
    oNode.Tag = EDT_PackType
    Set oNode = tvItemTree.Nodes.Add("comm_define", tvwChild, "code_areatype", "地区", "cf")
    oNode.Tag = EDT_AreaType
    Set oNode = tvItemTree.Nodes.Add("comm_define", tvwChild, "code_loader", "装卸工", "cf")
    oNode.Tag = EDT_LoadWorker
    Set oNode = tvItemTree.Nodes.Add("comm_define", tvwChild, "code_operator", "受理人", "cf")
    oNode.Tag = EDT_Operator
    Set oNode = tvItemTree.Nodes.Add("comm_define", tvwChild, "code_saveposition", "保存位置", "cf")
    oNode.Tag = EDT_SavePosition
    Set oNode = tvItemTree.Nodes.Add("comm_define", tvwChild, "code_message", "常用短信", "cf")
    oNode.Tag = EDT_Other1

    Set oNode = tvItemTree.Nodes.Add(, tvwChild, "load_charge", "装卸费设置", "cl")

End Sub
Private Sub AddLvHeader(pszKey As String)
    '添加ListView的列首

    If Left(pszKey, 4) = "code" Then    '基本码表
        SetFlex vsBaseInfo, 1, 3
        vsBaseInfo.TextMatrix(0, 0) = "序"
        vsBaseInfo.TextMatrix(0, 1) = "名称"
        vsBaseInfo.TextMatrix(0, 2) = "备注"
    Else
        Select Case pszKey
            Case "load_charge"
                SetFlex vsBaseInfo, 1, 5
                vsBaseInfo.TextMatrix(0, 0) = "序"
                vsBaseInfo.TextMatrix(0, 1) = "代码"
                vsBaseInfo.TextMatrix(0, 2) = "计重分类"
                vsBaseInfo.TextMatrix(0, 3) = "装卸费"
                vsBaseInfo.TextMatrix(0, 4) = "备注"

        End Select
    End If
    AlignHeadWidth Me.name, vsBaseInfo
End Sub
Private Sub ListBaseInfo(pszKey As String)
    WriteProcessBar True, , , "正在查询..."

    cmdSave.Enabled = False

    '添加ListView的列首
    Dim i As Integer, j As Integer, nRows As Integer
    nRows = 1
    If Left(pszKey, 4) = "code" Then
        Dim nType As Integer
        Dim aszBaseInfo() As String
        nType = tvItemTree.Nodes(pszKey).Tag
        aszBaseInfo = g_oPackageParam.ListBaseDefine(nType)
        For i = 1 To ArrayLength(aszBaseInfo)
            WriteProcessBar , i, ArrayLength(aszBaseInfo)
            vsBaseInfo.AddItem nRows
            vsBaseInfo.RowData(nRows) = aszBaseInfo(i, 1)
            vsBaseInfo.TextMatrix(nRows, 1) = aszBaseInfo(i, 3)
            vsBaseInfo.TextMatrix(nRows, 2) = aszBaseInfo(i, 4)
            nRows = nRows + 1
        Next i
    Else
        Dim avBaseInfo
        Select Case pszKey
            Case "load_charge"
                avBaseInfo = g_oPackageParam.ListLoadChargeCode
                For i = 1 To ArrayLength(avBaseInfo)
                    WriteProcessBar , i, ArrayLength(avBaseInfo)
                    vsBaseInfo.AddItem nRows
                    vsBaseInfo.TextMatrix(nRows, 1) = avBaseInfo(i, 1)
                    vsBaseInfo.TextMatrix(nRows, 2) = avBaseInfo(i, 2)
                    vsBaseInfo.TextMatrix(nRows, 3) = avBaseInfo(i, 3)
                    vsBaseInfo.TextMatrix(nRows, 4) = avBaseInfo(i, 4)
                    nRows = nRows + 1
                Next i


        End Select

    End If
    WriteProcessBar False, , , ""

End Sub
Private Sub cmdAddNew_Click()
    If vsBaseInfo.Row < 0 Or vsBaseInfo.Col < 0 Then Exit Sub

    vsBaseInfo.AddItem vsBaseInfo.Rows
    ReDim Preserve malNewRows(1 To ArrayLength(malNewRows) + 1)
    malNewRows(ArrayLength(malNewRows)) = vsBaseInfo.Rows - 1

    vsBaseInfo.Row = vsBaseInfo.Rows - 1
    Select Case tvItemTree.Nodes(mnItemIndex).Key
        Case "load_charge"
            vsBaseInfo.Col = 1
       Case Else
            vsBaseInfo.Col = 1
    End Select

    'vsBaseInfo.Cell(flexcpBackColor, vsBaseInfo.Row, 1, vsBaseInfo.Row, vsBaseInfo.Cols - 1) = cnColor_Edited
    vsBaseInfo.EditCell
End Sub

Private Sub cmdDel_Click()
    If vsBaseInfo.Row <= 0 Or vsBaseInfo.Col <= 0 Then Exit Sub

    Dim i As Integer, j As Integer
    For i = 1 To ArrayLength(malNewRows)    '删除新增队列的指定项目
        If malNewRows(i) = vsBaseInfo.Row Then
            RemoveArrayItem malNewRows, i
            GoTo RemoveItem
        End If
    Next i
    For i = 1 To ArrayLength(malEditRows)
        If malEditRows(i) = vsBaseInfo.Row Then
            RemoveArrayItem malEditRows, i
            Exit For
        End If
    Next i
    '添加删除项
    ReDim Preserve maszDeleteID(1 To ArrayLength(maszDeleteID) + 1)
    If Left(tvItemTree.Nodes(mnItemIndex).Key, 4) = "code" Then     '常用定义的话
        maszDeleteID(ArrayLength(maszDeleteID)) = vsBaseInfo.RowData(vsBaseInfo.Row)
    Else
        ReDim Preserve maszDeleteUnitID(1 To ArrayLength(maszDeleteUnitID) + 1)
        maszDeleteID(ArrayLength(maszDeleteID)) = vsBaseInfo.TextMatrix(vsBaseInfo.Row, 2)
        maszDeleteUnitID(ArrayLength(maszDeleteUnitID)) = ResolveDisplay(vsBaseInfo.TextMatrix(vsBaseInfo.Row, 1))
    End If


    cmdSave.Enabled = True
RemoveItem:
    For i = vsBaseInfo.Row + 1 To vsBaseInfo.Rows - 1
        vsBaseInfo.TextMatrix(i, 0) = Val(vsBaseInfo.TextMatrix(i, 0)) - 1
    Next i
    vsBaseInfo.RemoveItem vsBaseInfo.Row
End Sub
Private Sub RemoveArrayItem(paItems() As Long, ItemIndex As Integer)
    Dim aTmpItems() As Long
    If ArrayLength(paItems) = 1 Then
        paItems = aTmpItems
        Exit Sub
    End If
    Dim i As Integer
    For i = ItemIndex + 1 To ArrayLength(paItems)
        paItems(i) = paItems(i) - 1 '下调一行
        paItems(i - 1) = paItems(i)
    Next i
    ReDim Preserve paItems(1 To ArrayLength(paItems) - 1)
End Sub
Private Sub RemoveArrayItem2(paItems() As Long, ItemIndex As Integer)
    Dim aTmpItems() As Long
    If ArrayLength(paItems) = 1 Then
        paItems = aTmpItems
        Exit Sub
    End If
    Dim i As Integer
    For i = ItemIndex + 1 To ArrayLength(paItems)
        paItems(i - 1) = paItems(i)
    Next i
    ReDim Preserve paItems(1 To ArrayLength(paItems) - 1)
End Sub
Private Sub cmdSave_Click()
On Error GoTo ErrHandle
    Dim lProgess As Integer, lStep As Integer
    lProgess = ArrayLength(malNewRows) + ArrayLength(malEditRows) + ArrayLength(maszDeleteID)
    WriteProcessBar True, , lProgess



    Dim i As Integer
    '以下新增处理
    Dim aszValues() As String
    Dim avValues()
    '----------------------------------------------
    ShowSBInfo "正在保存新增的项目，请等待..."
    If ArrayLength(malNewRows) > 0 Then
        If Left(tvItemTree.Nodes(mnItemIndex).Key, 4) = "code" Then
            '以下处理码表
            For i = 1 To ArrayLength(malNewRows)
                lStep = lStep + 1
                WriteProcessBar , lStep, lProgess

                vsBaseInfo.RowData(malNewRows(1)) = g_oPackageParam.AddBaseDefine(tvItemTree.Nodes(mnItemIndex).Tag, Trim(vsBaseInfo.TextMatrix(malNewRows(1), 1)), vsBaseInfo.TextMatrix(malNewRows(1), 2))
                '设置返回回来的ID号并恢复颜色
                'vsBaseInfo.Cell(flexcpBackColor, malNewRows(1), 1, malNewRows(1), vsBaseInfo.Cols - 1) = cnColor_Normal
                RemoveArrayItem2 malNewRows, 1

            Next i
        Else
            Select Case tvItemTree.Nodes(mnItemIndex).Key

                Case "load_charge"
                    For i = 1 To ArrayLength(malNewRows)
                        lStep = lStep + 1
                        WriteProcessBar , lStep, lProgess

                        g_oPackageParam.AddLoadChargeCode vsBaseInfo.TextMatrix(malNewRows(1), 1), _
                                                            vsBaseInfo.TextMatrix(malNewRows(1), 2), _
                                                            Val(vsBaseInfo.TextMatrix(malNewRows(1), 3)), _
                                                            vsBaseInfo.TextMatrix(malNewRows(1), 4)
                        vsBaseInfo.Cell(flexcpBackColor, malNewRows(1), 1, malNewRows(1), vsBaseInfo.Cols - 1) = cnColor_Normal
                        RemoveArrayItem2 malNewRows, 1
                    Next i
            End Select
        End If
    End If


    '以下更改处理
    '----------------------------------------------
    ShowSBInfo "正在保存修改的项目，请等待..."
    If ArrayLength(malEditRows) > 0 Then
        If Left(tvItemTree.Nodes(mnItemIndex).Key, 4) = "code" Then
            '以下处理码表
            For i = ArrayLength(malEditRows) To 1 Step -1
                g_oPackageParam.UpdBaseDefine vsBaseInfo.RowData(malEditRows(i)), tvItemTree.Nodes(mnItemIndex).Tag, vsBaseInfo.TextMatrix(malEditRows(i), 1), vsBaseInfo.TextMatrix(malEditRows(i), 2)
                vsBaseInfo.Cell(flexcpBackColor, malEditRows(i), 1, malEditRows(i), vsBaseInfo.Cols - 1) = cnColor_Normal
                RemoveArrayItem malEditRows, i
            Next i
        Else
            Select Case tvItemTree.Nodes(mnItemIndex).Key

                Case "load_charge"
                    For i = ArrayLength(malEditRows) To 1 Step -1
                        lStep = lStep + 1
                        WriteProcessBar , lStep, lProgess

                        g_oPackageParam.UpdLoadChargeCode vsBaseInfo.TextMatrix(malEditRows(i), 1), _
                                                           vsBaseInfo.TextMatrix(malEditRows(i), 2), _
                                                            Val(vsBaseInfo.TextMatrix(malEditRows(i), 3)), _
                                                            vsBaseInfo.TextMatrix(malEditRows(i), 4)
                        vsBaseInfo.Cell(flexcpBackColor, malEditRows(i), 1, malEditRows(i), vsBaseInfo.Cols - 1) = cnColor_Normal
                        RemoveArrayItem malEditRows, i
                    Next i

            End Select

        End If
    End If

    '以下删除处理
    '----------------------------------------------
    Dim aszTmp() As String

    ShowSBInfo "正在删除项目，请等待..."
    If ArrayLength(maszDeleteID) > 0 Then
        If Left(tvItemTree.Nodes(mnItemIndex).Key, 4) = "code" Then
            '以下处理码表
            For i = ArrayLength(maszDeleteID) To 1 Step -1
                lStep = lStep + 1
                WriteProcessBar , lStep, lProgess

                g_oPackageParam.DelBaseDefine CLng(maszDeleteID(i))
                If ArrayLength(maszDeleteID) = 1 Then
                    maszDeleteID = aszTmp
                Else
                    ReDim Preserve maszDeleteID(1 To ArrayLength(maszDeleteID))
                End If
            Next i
        Else
            Select Case tvItemTree.Nodes(mnItemIndex).Key

                Case "load_charge"
                    For i = ArrayLength(maszDeleteID) To 1 Step -1
                        lStep = lStep + 1
                        WriteProcessBar , lStep, lProgess

                        g_oPackageParam.DelLoadChargeCode maszDeleteID(i)
                        If i = 1 Then
                            maszDeleteID = aszTmp
                            maszDeleteUnitID = aszTmp
                        Else
                            ReDim Preserve maszDeleteID(1 To i - 1)
                            ReDim Preserve maszDeleteUnitID(1 To i - 1)
                        End If
                    Next i

            End Select
        End If

    End If


    WriteProcessBar False, , , ""
    '设置回初始状态
    cmdSave.Enabled = False

    Exit Sub
ErrHandle:
    WriteProcessBar False, , , ""
    ShowErrorMsg
End Sub



Private Sub Form_Load()
    '添加列表头
    mnItemIndex = -1
    AddTreeItem


    '设置初始值

    SetFlex vsBaseInfo, 0, 0

    Spliter1.InitSpliter ptLeft, ptRight
    AlignHeadWidth Me.name, vsBaseInfo
End Sub
Private Sub Form_Resize()
On Error Resume Next
    ptTop.Move 0, 0, Me.ScaleWidth
    ptMain.Move 0, ptTop.ScaleHeight, Me.ScaleWidth, Me.ScaleHeight - ptTop.Height

    Spliter1.LayoutIt
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveHeadWidth Me.name, vsBaseInfo
End Sub

Private Sub ptLeft_Resize()
On Error Resume Next
    imgLeftBar.Move 0, 0, ptLeft.ScaleWidth, imgLeftBar.Height
    lblLeftTitle.Move 0, lblLeftTitle.Top, ptLeft.ScaleWidth, lblLeftTitle.Height
    tvItemTree.Move 0, imgLeftBar.Height, ptLeft.ScaleWidth, ptLeft.ScaleHeight - imgLeftBar.Height
End Sub

Private Sub ptRight_Resize()
On Error Resume Next
    vsBaseInfo.Move 0, 0, ptRight.ScaleWidth, ptRight.ScaleHeight
End Sub


Private Sub tvItemTree_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Index = mnItemIndex Then
        Exit Sub
    Else
        Dim alTmp() As Long, aszTmp() As String, nYes As Integer
        If ArrayLength(malNewRows) > 0 Or ArrayLength(malEditRows) > 0 Or ArrayLength(maszDeleteID) > 0 Then
            nYes = MsgBox("项目已变动，是否保存变动？", vbYesNo + vbQuestion, "问题")
            If nYes = vbYes Then
                Call cmdSave_Click
            Else
                malNewRows = alTmp: malEditRows = alTmp
                maszDeleteID = aszTmp: maszDeleteUnitID = aszTmp
            End If
        End If
        mnItemIndex = Node.Index
        SaveHeadWidth Me.name, vsBaseInfo
    End If
    AddLvHeader Node.Key
    AlignHeadWidth Me.name, vsBaseInfo
    ListBaseInfo Node.Key
End Sub



Private Sub LayoutButton(pnStatus As EFormStatus)
    Select Case pnStatus
        Case EFS_AddNew
            cmdAddNew.Enabled = False
'            cmdEdit.Enabled = False
            cmdDel.Enabled = True
            cmdSave.Enabled = False
        Case EFS_Modify
            cmdAddNew.Enabled = False
'            cmdEdit.Enabled = False
            cmdDel.Enabled = True
            cmdSave.Enabled = False
        Case EFS_Show
            cmdAddNew.Enabled = True
'            cmdEdit.Enabled = True
            cmdDel.Enabled = True
            cmdSave.Enabled = False
    End Select
End Sub



Private Sub vsBaseInfo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Row <= 0 Or Col <= 0 Then Exit Sub
On Error Resume Next
    '完成编辑后，截去多余的字段长度
    Select Case tvItemTree.Nodes(mnItemIndex).Key
        Case "load_charge"
            Select Case Col
                Case 1
                    vsBaseInfo.ColEditMask(1) = ""
                Case 2
                    vsBaseInfo.Text = GetUnicodeBySize(vsBaseInfo.Text, 30)
                Case 3
                    vsBaseInfo.ColEditMask(3) = ""
                Case 4
            End Select
        Case Else

    End Select
End Sub

Private Sub vsBaseInfo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row <= 0 Or Col <= 0 Then Exit Sub
On Error Resume Next
    Select Case tvItemTree.Nodes(mnItemIndex).Key

        Case "load_charge"
            Select Case Col
                Case 1      '编号
                    vsBaseInfo.ColEditMask(1) = "00"
                Case 3      '装卸费
                    'vsBaseInfo.ColEditMask(3) = "0"
            End Select
    End Select
End Sub



Private Sub vsBaseInfo_ChangeEdit()
    If vsBaseInfo.Row = 0 Or vsBaseInfo.Col = 0 Then Exit Sub
    Dim i As Integer
    vsBaseInfo.CellBackColor = cnColor_Edited

    cmdSave.Enabled = True
    For i = 1 To ArrayLength(malNewRows)
        If malNewRows(i) = vsBaseInfo.Row Then
            Exit Sub
        End If
    Next i
    For i = 1 To ArrayLength(malEditRows)
        If malEditRows(i) = vsBaseInfo.Row Then
            Exit Sub
        End If
    Next i
    ReDim Preserve malEditRows(1 To ArrayLength(malEditRows) + 1)
    malEditRows(ArrayLength(malEditRows)) = vsBaseInfo.Row
End Sub

Private Sub vsBaseInfo_EnterCell()
    If vsBaseInfo.Row <= 0 Or vsBaseInfo.Col <= 0 Then Exit Sub
    vsBaseInfo.EditCell
End Sub

Private Sub vsBaseInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If vsBaseInfo.Col < vsBaseInfo.Cols - 1 Then
            vsBaseInfo.Col = vsBaseInfo.Col + 1
            vsBaseInfo.EditCell
        Else
            cmdAddNew_Click
        End If
    End If
End Sub

Private Sub vsBaseInfo_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row = 0 Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        If Col < vsBaseInfo.Cols - 1 Then
            vsBaseInfo.Col = vsBaseInfo.Col + 1
            vsBaseInfo.EditCell
        Else
            cmdAddNew_Click
        End If
    End If
End Sub
'初始化控件
Private Sub InitVsFlex(vsFlex As VSFlexGrid)

End Sub


