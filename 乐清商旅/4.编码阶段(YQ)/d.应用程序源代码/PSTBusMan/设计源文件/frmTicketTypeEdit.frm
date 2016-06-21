VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmTicketTypeEdit 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "特殊票价计算参数编辑"
   ClientHeight    =   4920
   ClientLeft      =   2760
   ClientTop       =   2715
   ClientWidth     =   7050
   HelpContextID   =   1003801
   Icon            =   "frmTicketTypeEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox CboPriceTable 
      Height          =   300
      Left            =   1125
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   75
      Width           =   1995
   End
   Begin RTComctl3.CoolButton cmdHelp 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   5670
      TabIndex        =   5
      Top             =   4500
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmTicketTypeEdit.frx":0442
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
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4455
      TabIndex        =   4
      Top             =   4500
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "取消"
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
      MICON           =   "frmTicketTypeEdit.frx":045E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdOK 
      Default         =   -1  'True
      Height          =   315
      Left            =   3225
      TabIndex        =   3
      Top             =   4500
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmTicketTypeEdit.frx":047A
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
      Caption         =   "特殊票价类票价=(系数)×(票价项)＋(附加值)"
      Height          =   3900
      Left            =   60
      TabIndex        =   6
      Top             =   510
      Width           =   6885
      Begin VSFlex7LCtl.VSFlexGrid vsParam 
         Height          =   3500
         Left            =   90
         TabIndex        =   2
         Top             =   240
         Width           =   6645
         _cx             =   5080
         _cy             =   5080
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
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   65
         Cols            =   6
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
         FillStyle       =   1
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   -1  'True
         WordWrap        =   -1  'True
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
   Begin VB.Label lblExcuteTable 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "票价表(&I):"
      Height          =   180
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "frmTicketTypeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *******************************************************************
' *  Source File Name  : frmTicketTypeEdit.frm                      *
' *  Project Name: PSTBusMan                                        *
' *  Engineer:                                                      *
' *  Date Generated: 2002/09/03                                     *
' *  Last Revision Date : 2002/09/03                                *
' *  Brief Description   : 半价编辑                                 *
' *  DocNo:                                                         *
' *******************************************************************


Option Explicit
Option Base 1
Const clColorChanged = vbRed
Const cnMax = 100
Const cnPriceItemCount = 16
Dim tHalfTickInfo() As THalfTicketItemParam
Dim aszPriceItem() As String

Private Sub GetInfoFromUI()
    Dim i As Integer, j As Integer
    Dim szitem As String
    Dim nCount As Integer
    On Error GoTo there
    With vsParam
    For j = 1 To .Rows - 1
        .Col = 2
        .Row = j
        If .CellForeColor = clColorChanged Then
            nCount = nCount + 1
        End If
    Next j
    If nCount = 0 Then Exit Sub
    ReDim tHalfTickInfo(1 To nCount)
    nCount = 0
    For j = 1 To .Rows - 1
        .Col = 2
        .Row = j
        If .CellForeColor = clColorChanged Then
            nCount = nCount + 1
        
        
            tHalfTickInfo(nCount).szTicketType = ResolveDisplay(.TextMatrix(j, 0))
            If j Mod cnPriceItemCount > 0 Then
               szitem = (j Mod cnPriceItemCount) - 1
            Else
               szitem = 15
            End If
            szitem = Format(szitem, "0000")
            tHalfTickInfo(nCount).szTicketItem = szitem
            tHalfTickInfo(nCount).sgParam1 = CSng(.TextMatrix(j, 2))
            tHalfTickInfo(nCount).sgParam2 = CSng(.TextMatrix(j, 3))
            tHalfTickInfo(nCount).szAnnotation = .TextMatrix(j, 4)
        End If
    Next
    End With
Exit Sub
there:

    ShowErrorMsg
End Sub

Private Sub CboPriceTable_Click()
    GetParamInfo
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    '显示帮助
End Sub

Private Sub cmdOk_Click()
    Update
End Sub

Private Sub Form_Load()
    With Me
        .Top = Screen.Height / 2 - Me.Height / 2
        .Left = Screen.Width / 2 - Me.Width / 2
    End With
    InitGrid
    FillPriceTable
    GetParamInfo
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub FillPriceTable()
    '填充票价表
    Dim aszRoutePriceTable() As String
    Dim i As Integer, nCount As Integer
    Dim szPriceTable As String
    
    On Error GoTo ErrorHandle
    aszRoutePriceTable = GetPriceTable(Now)
    nCount = ArrayLength(aszRoutePriceTable)
    cboPriceTable.Clear
    If nCount > 0 Then
        For i = 1 To nCount
            szPriceTable = MakeDisplayString(aszRoutePriceTable(i, 1), aszRoutePriceTable(i, 2))
            cboPriceTable.AddItem szPriceTable
            If aszRoutePriceTable(i, 7) = cnRunTable Then cboPriceTable.Text = MakeDisplayString(aszRoutePriceTable(i, 1), aszRoutePriceTable(i, 2))
        Next i
    End If
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub Update()
'    保存更改
    Dim oHalfPrice As New HalfTicketPrice
    Dim nResult As Integer
    Dim i As Integer
    Dim szPriceTable As String
    szPriceTable = ResolveDisplay(cboPriceTable.Text)
    nResult = MsgBox("确实要修改特殊票参数吗？", vbYesNo + vbQuestion, Me.Caption)
    SetBusy
    If nResult = vbYes Then
        GetInfoFromUI
        On Error GoTo Here
        oHalfPrice.Init g_oActiveUser
        For i = 1 To ArrayLength(tHalfTickInfo)
            Call oHalfPrice.ModifyItemParam(CInt(tHalfTickInfo(i).szTicketType), tHalfTickInfo(i), szPriceTable)
        Next i
    End If
    SetNormal
    Unload Me
Exit Sub
Here:
    ShowErrorMsg
    SetNormal
End Sub


Private Sub vsParam_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim szValid As String
    Dim lTemp As Long
    Dim i As Integer
    Const szMsgNumberContext = "请入0到1之间的数字"
    Const szMsgTooBigContext = "数值应该在-100到100之间,请重输入."
    Const szMsgInputContext = "请输入数字信息"
    '验证有效性
    '正在编辑的数据
    szValid = vsParam.EditText
    If Not IsNumeric(szValid) Then
        MsgBox szMsgInputContext, vbInformation, Me.Caption
        Cancel = True
        Exit Sub
    End If
    '验证系数是否正确
    lTemp = Val(szValid)
    If Col = 2 Then
        '如果为数字
        If lTemp > 1 Or lTemp < 0 Then
            MsgBox szMsgNumberContext, vbInformation, Me.Caption
            Cancel = True
            Exit Sub
        End If
    ElseIf Col = 3 Then
        '验证附加值
        If lTemp < 0 Then
            '如果小于零
            If Abs(lTemp) > cnMax Then
                MsgBox szMsgTooBigContext, vbInformation, Me.Caption
                Cancel = True
                Exit Sub
            End If
        ElseIf lTemp > cnMax Then
            '如果大于规定值
            MsgBox szMsgTooBigContext, vbInformation, Me.Caption
            Cancel = True
            Exit Sub
        End If
    End If
    If vsParam.Text <> vsParam.EditText Then
        '如果修改过了
        For i = 2 To 4
            vsParam.Col = i
            vsParam.CellForeColor = clColorChanged
        Next i
        vsParam.Col = Col
        cmdOK.Enabled = True
    End If
End Sub

Private Sub InitGrid()
    '初始化网格
    With vsParam
        '初始化
        .Rows = (g_nTicketCountValid - 1) * cnPriceItemCount + 1 '''由Database 定
        .Cols = 5
        '设定标题
        .TextMatrix(0, 0) = "票种"
        .TextMatrix(0, 1) = "票价项名"
        .TextMatrix(0, 2) = "系数"
        .TextMatrix(0, 3) = "附加值"
        .TextMatrix(0, 4) = "注释"
        .ColWidth(0) = 1500
        .ColWidth(1) = 1350
        .ColWidth(2) = 600
        .ColWidth(3) = 800
        .ColWidth(4) = 2000
        '票价项号和票价项名栏置灰
        .FixedCols = 2
        .Col = 2
        .Row = 1
    End With
End Sub

Private Sub GetParamInfo()
    On Error GoTo Here
    
    Dim oHalfTicket As New HalfTicketPrice
    Dim oTicketPriceMan As New TicketPriceMan
    oTicketPriceMan.Init g_oActiveUser
    aszPriceItem = oTicketPriceMan.GetAllTicketItem

    oHalfTicket.Init g_oActiveUser
    tHalfTickInfo = oHalfTicket.GetItemParam(0, ResolveDisplay(cboPriceTable.Text)) '所有参数
    Set oTicketPriceMan = Nothing
    Set oHalfTicket = Nothing
    FillGrid
    Exit Sub
Here:
    ShowErrorMsg
    Set oTicketPriceMan = Nothing
    Set oHalfTicket = Nothing

End Sub

Private Sub FillGrid()
    '填充真正的信息
    Dim i As Integer, j As Integer
    Dim nTemp As Integer
    Dim n As Integer
    Dim TicketType() As TTicketType
    Dim nTicketTypeCount As Integer
    Dim oParam As New SystemParam
    On Error GoTo there
    '得到所有的票种
    oParam.Init g_oActiveUser
    TicketType = oParam.GetAllTicketType(TP_TicketTypeValid)
    nTicketTypeCount = ArrayLength(TicketType)
    With vsParam
        .Visible = False
        For i = 1 To .Rows - 1
            .Row = i Mod cnPriceItemCount
            If .Row = 0 Then .Row = cnPriceItemCount
            If aszPriceItem(.Row, 3) = TP_PriceItemNotUse Then
                '如果票价项未使用
               .RowHeight(i) = 0
            Else
               .RowHeight(i) = 255
            End If
            nTemp = CInt(tHalfTickInfo(i).szTicketItem)
            n = Int((i - 1) / cnPriceItemCount) * cnPriceItemCount
            .Row = n + nTemp + 1
            .Col = 0
            For j = 2 To nTicketTypeCount
                If tHalfTickInfo(i).szTicketType = TicketType(j).nTicketTypeID Then
                    .Text = MakeDisplayString(TicketType(j).nTicketTypeID, Trim(TicketType(j).szTicketTypeName))
                    Exit For
                End If
            Next
            .Col = 1
            For j = 1 To cnPriceItemCount
                If nTemp = CInt(aszPriceItem(j, 1)) Then
                     .Text = aszPriceItem(j, 2)
                     Exit For
                End If
            Next j
            .Col = 2
            .Text = tHalfTickInfo(i).sgParam1
            .Col = 3
            .Text = tHalfTickInfo(i).sgParam2
            .Col = 4
            .Text = tHalfTickInfo(i).szAnnotation
        Next i
        .Visible = True
        .MergeCol(0) = True
        .MergeCells = flexMergeRestrictRows
        .Col = 2
        .Row = 1
    End With
Exit Sub
there:

    ShowErrorMsg
End Sub

