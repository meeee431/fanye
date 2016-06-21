VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmItemFormula 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "站点行包公式设置"
   ClientHeight    =   6630
   ClientLeft      =   4170
   ClientTop       =   2130
   ClientWidth     =   8220
   Icon            =   "frmItemFormula.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   8220
   StartUpPosition =   2  '屏幕中心
   Begin RTComctl3.CoolButton cmdDelRow 
      Height          =   345
      Left            =   3810
      TabIndex        =   16
      ToolTipText     =   "保存协议"
      Top             =   6150
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   609
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItemFormula.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdAddRow 
      Height          =   345
      Left            =   2790
      TabIndex        =   15
      ToolTipText     =   "保存协议"
      Top             =   6150
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   609
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItemFormula.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdQuery 
      Default         =   -1  'True
      Height          =   345
      Left            =   6810
      TabIndex        =   14
      Top             =   900
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "查询(&Q)"
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
      MICON           =   "frmItemFormula.frx":0044
      PICN            =   "frmItemFormula.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtStationName 
      Height          =   300
      Left            =   3570
      TabIndex        =   13
      ToolTipText     =   "可以查询以该字开头的所有的站点"
      Top             =   915
      Width           =   885
   End
   Begin VB.ComboBox cboAcceptType 
      Height          =   300
      Left            =   5655
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   915
      Width           =   945
   End
   Begin RTComctl3.TextButtonBox txtStationID 
      Height          =   300
      Left            =   945
      TabIndex        =   9
      Top             =   915
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   529
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
   Begin VSFlex7LCtl.VSFlexGrid vsDetail 
      Height          =   4305
      Left            =   60
      TabIndex        =   3
      Top             =   1380
      Width           =   8070
      _cx             =   14235
      _cy             =   7594
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   8430
      TabIndex        =   1
      Top             =   -30
      Width           =   8430
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包公式设置:"
         Height          =   180
         Left            =   270
         TabIndex        =   2
         Top             =   270
         Width           =   1170
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   75
      Left            =   0
      TabIndex        =   0
      Top             =   690
      Width           =   8430
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   345
      Left            =   255
      TabIndex        =   4
      Top             =   6150
      Width           =   1140
      _ExtentX        =   2011
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
      MICON           =   "frmItemFormula.frx":03FA
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
      Left            =   6795
      TabIndex        =   5
      Top             =   6150
      Width           =   1140
      _ExtentX        =   2011
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
      MICON           =   "frmItemFormula.frx":0416
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
      Left            =   5385
      TabIndex        =   6
      ToolTipText     =   "保存协议"
      Top             =   6150
      Width           =   1140
      _ExtentX        =   2011
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
      MICON           =   "frmItemFormula.frx":0432
      PICN            =   "frmItemFormula.frx":044E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   930
      Left            =   -165
      TabIndex        =   7
      Top             =   5865
      Width           =   8745
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "站点名称(&N):"
      Height          =   180
      Left            =   2445
      TabIndex        =   12
      Top             =   975
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "托运方式(&M):"
      Height          =   180
      Left            =   4500
      TabIndex        =   10
      Top             =   975
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "站点(&T):"
      Height          =   180
      Left            =   135
      TabIndex        =   8
      Top             =   975
      Width           =   720
   End
End
Attribute VB_Name = "frmItemFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const cnCols = 5

Const cnStation = 1
Const cnAcceptType = 2
Const cnPriceItem = 3
Const cnFormula = 4


Const cnNormalColor = vbBlack
Const cnChangedColor = vbRed

Private m_rsItems As Recordset


Private m_bQuery As Boolean


Private Sub cmdAddRow_Click()
    '新增一个站点的空白行
    Dim nRow As Integer
    Dim i As Integer
    With vsDetail
        nRow = .Rows
        .Rows = nRow + m_rsItems.RecordCount
        
        .MergeCol(cnStation) = True
        .MergeCells = flexMergeRestrictColumns
        m_rsItems.MoveFirst
        For i = 1 To m_rsItems.RecordCount
            .Row = i + nRow - 1
            .Col = cnPriceItem
            .CellForeColor = cnChangedColor
            .TextMatrix(i + nRow - 1, cnPriceItem) = MakeDisplayString(FormatDbValue(m_rsItems!charge_item), FormatDbValue(m_rsItems!chinese_name))
            m_rsItems.MoveNext
        Next i
        For i = nRow - 1 To 1 Step -1
            If .TextMatrix(i, cnStation) = "" Then .RemoveItem i
        Next i
        
    End With
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelRow_Click()
    '删除一个站点的公式设置
    Dim i As Integer
    Dim szStation As String
    With vsDetail
        szStation = .TextMatrix(.Row, cnStation)
        If .Row = 0 Then Exit Sub
        If MsgBox("是否删除站点为""" & szStation & """的所有行的公式信息", vbYesNo, Me.Caption) = vbYes Then
            For i = .Rows - 1 To 1 Step -1
                If .TextMatrix(i, cnStation) = szStation Then
                    .RemoveItem i
                End If
            Next i
        
        End If
    End With
End Sub

Private Sub cmdOk_Click()
    '保存新增或修改到数据库
    Dim i As Integer
    Dim j As Integer
    Dim atItemFormulaInfo() As TLuggageItemFormulaInfo
    Dim nCount As Integer
    '计算有多少条需要保存
    With vsDetail
        For i = 1 To .Rows - 1
            .Row = i
            For j = 1 To .Cols - 1
                .Col = j
                If .CellForeColor = cnChangedColor Then Exit For
                
            Next j
            If j < .Cols Then
                '说明已经新增或修改
                nCount = nCount + 1
            End If
        Next i
        If nCount = 0 Then
            MsgBox "无数据需要保存", vbInformation, Me.Caption
            Exit Sub
        
        Else
            ReDim atItemFormulaInfo(1 To nCount)
            nCount = 0
            For i = 1 To .Rows - 1
                .Row = i
                For j = 1 To .Cols - 1
                    .Col = j
                    If .CellForeColor = cnChangedColor Then Exit For
                    
                Next j
                If j < .Cols Then
                    '说明已经新增或修改
                    nCount = nCount + 1
                    atItemFormulaInfo(nCount).PriceItemID = ResolveDisplay(.TextMatrix(i, cnPriceItem))
                    atItemFormulaInfo(nCount).AcceptType = ResolveDisplay(.TextMatrix(i, cnAcceptType))
                    atItemFormulaInfo(nCount).StationID = ResolveDisplay(.TextMatrix(i, cnStation))
                    atItemFormulaInfo(nCount).FormulaID = ResolveDisplay(.TextMatrix(i, cnFormula))
                End If
            Next i
            m_oluggageSvr.SetLugItemFormulaInfo atItemFormulaInfo
            '设置为已修改
            SetColor cnNormalColor
            MsgBox "公式设置已保存", vbInformation, Me.Caption
            
        End If
        
    End With
    
End Sub

Private Sub SetColor(pnColor As ColorConstants)
    Dim i As Integer
    Dim j As Integer
    With vsDetail
        For i = 1 To .Rows - 1
            .Row = i
            For j = 1 To .Cols - 1
                .Col = j
                .CellForeColor = pnColor
            Next j
        Next i
    End With
    
        
End Sub


Private Sub cmdQuery_Click()
    '查询出符合条件的行包项的公式
    
    Dim atItemFormulaInfo() As TLuggageItemFormulaInfo
    Dim nCount As Integer
    Dim i As Integer
    With vsDetail
        .Rows = 1
        
        atItemFormulaInfo = m_oluggageSvr.GetLugItemFormulaInfo(ResolveDisplay(txtStationID.Text), txtStationName.Text, ResolveDisplay(cboAcceptType.Text))
        nCount = ArrayLength(atItemFormulaInfo)
        If nCount > 0 Then
            .Rows = 1 + nCount
            m_bQuery = True
            
            For i = 1 To nCount
                
                .TextMatrix(i, cnStation) = MakeDisplayString(atItemFormulaInfo(i).StationID, atItemFormulaInfo(i).StationName)
                .TextMatrix(i, cnAcceptType) = MakeDisplayString(atItemFormulaInfo(i).AcceptType, atItemFormulaInfo(i).AcceptTypeName)
                .TextMatrix(i, cnPriceItem) = MakeDisplayString(atItemFormulaInfo(i).PriceItemID, atItemFormulaInfo(i).PriceItemName)
                .TextMatrix(i, cnFormula) = MakeDisplayString(atItemFormulaInfo(i).FormulaID, atItemFormulaInfo(i).FormulaName)
            Next i
            SetColor cnNormalColor
            m_bQuery = False
        Else
            
            cmdAddRow_Click
        End If
        
    End With
    
    
End Sub


Private Sub Form_Load()
    AlignFormPos Me
    Set m_rsItems = m_oLugParam.GetPriceItemRS(LuggagePriceItemUsed, ELuggageAcceptType.NormalAccept)
    FillAcceptType
    InitFlex
    
End Sub


Private Sub InitFlex()
    '初始化网格
    With vsDetail
        .Cols = cnCols
        .Rows = 1
        
        .TextMatrix(0, cnStation) = "站点"
        .TextMatrix(0, cnAcceptType) = "受理类型"
        .TextMatrix(0, cnPriceItem) = "行包收费项"
        .TextMatrix(0, cnFormula) = "计算公式"
        '预先加入一条空白记录
        
        cmdAddRow_Click
        
        .FixedRows = 1
        .Row = 1
        '设置合并
        .MergeCol(cnStation) = True
        .MergeCells = flexMergeRestrictColumns
        '设置列宽
        .ColWidth(0) = 200
        .ColWidth(cnStation) = 1500
        .ColWidth(cnAcceptType) = 1500
        .ColWidth(cnPriceItem) = 1500
        .ColWidth(cnFormula) = 2000
        
        MakeFormulaStr
        MakeAcceptTypeStr
        
        .ColComboList(cnStation) = "..."
        
    End With
End Sub

Private Sub FillAcceptType()
    cboAcceptType.clear
    With cboAcceptType
        .AddItem MakeDisplayString("-1", "全部")
        .AddItem MakeDisplayString(ELuggageAcceptType.NormalAccept, GetLuggageTypeString(ELuggageAcceptType.NormalAccept))
        .AddItem MakeDisplayString(ELuggageAcceptType.CarryAccept, GetLuggageTypeString(ELuggageAcceptType.CarryAccept))
        .ListIndex = 0
    End With
    
End Sub

Private Sub MakeFormulaStr()
    Dim atAllFormulaInfo() As TLuggageFormulaInfo
    Dim nCount As Integer
    Dim i As Integer
    Dim szTemp As String
    
    atAllFormulaInfo = m_oluggageSvr.GetLugFormulaInfo()
    nCount = ArrayLength(atAllFormulaInfo)
    szTemp = " "
    For i = 1 To nCount
        szTemp = szTemp & "|" & MakeDisplayString(atAllFormulaInfo(i).FormulaID, atAllFormulaInfo(i).FormulaName)
    Next i
    
    vsDetail.ColComboList(cnFormula) = szTemp
End Sub


Private Sub MakeAcceptTypeStr()
    vsDetail.ColComboList(cnAcceptType) = MakeDisplayString(ELuggageAcceptType.NormalAccept, GetLuggageTypeString(ELuggageAcceptType.NormalAccept)) _
            & "|" & MakeDisplayString(ELuggageAcceptType.CarryAccept, GetLuggageTypeString(ELuggageAcceptType.CarryAccept))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Sub txtStationID_Click()
    '选择站点
    Dim oCommDialog As New CommDialog
    Dim aszTemp() As String
    oCommDialog.Init m_oAUser
    aszTemp = oCommDialog.SelectStation
    If ArrayLength(aszTemp) > 0 Then
        txtStationID.Text = MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))
    End If
    
End Sub

Private Sub vsDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer
    Dim szAcceptType As String
    Dim szStation As String
    If Col = cnAcceptType Then
        szAcceptType = vsDetail.TextMatrix(Row, Col)
        szStation = vsDetail.TextMatrix(Row, cnStation)
        If szAcceptType = "" Then Exit Sub
        With vsDetail
            For i = Row - m_rsItems.RecordCount To Row + m_rsItems.RecordCount
                If i > 0 And i < vsDetail.Rows Then
                    .Row = i
                    .Col = cnAcceptType
                    .CellForeColor = cnChangedColor
                    If vsDetail.TextMatrix(i, cnStation) = szStation Then vsDetail.TextMatrix(i, Col) = szAcceptType
                End If
            Next i
        End With
    End If
    
End Sub

Private Sub vsDetail_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer
    Dim oCommDialog As New CommDialog
    Dim aszTemp() As String
    Dim szStation As String
    If Col = cnStation Then
        '选择站点
        oCommDialog.Init m_oAUser
        aszTemp = oCommDialog.SelectStation
         szStation = vsDetail.TextMatrix(Row, cnStation)
        If ArrayLength(aszTemp) > 0 Then
            With vsDetail
                For i = Row - m_rsItems.RecordCount To Row + m_rsItems.RecordCount
                    If i > 0 And i < vsDetail.Rows Then
                        
                        .Row = i
                        .Col = cnStation 'szStation
                        .CellForeColor = cnChangedColor
                        If vsDetail.TextMatrix(i, Col) = szStation Then vsDetail.TextMatrix(i, Col) = MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))
                    End If
                Next i
            End With
        End If
    
    End If
End Sub

Private Sub vsDetail_ChangeEdit()
    If Not m_bQuery Then
        vsDetail.CellForeColor = cnChangedColor
    End If
    
End Sub
