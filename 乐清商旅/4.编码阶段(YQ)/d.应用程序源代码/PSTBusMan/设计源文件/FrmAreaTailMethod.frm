VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmAreaTailMethod 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "区域尾数处理方法设置"
   ClientHeight    =   5940
   ClientLeft      =   1155
   ClientTop       =   2025
   ClientWidth     =   8580
   HelpContextID   =   10000460
   Icon            =   "FrmAreaTailMethod.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   8580
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   330
      Left            =   7320
      TabIndex        =   21
      Top             =   5460
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "帮助"
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
      MICON           =   "FrmAreaTailMethod.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdFind 
      Height          =   330
      Left            =   7275
      TabIndex        =   16
      Top             =   705
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   582
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
      MICON           =   "FrmAreaTailMethod.frx":0166
      PICN            =   "FrmAreaTailMethod.frx":0182
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox CboPriceTable 
      Height          =   300
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1995
   End
   Begin VB.ListBox lstBusType 
      Appearance      =   0  'Flat
      Height          =   1080
      Left            =   1740
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   720
      Width           =   1590
   End
   Begin VB.ListBox lstPriceItem 
      Appearance      =   0  'Flat
      Height          =   1080
      Left            =   3480
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   720
      Width           =   1995
   End
   Begin RTComctl3.CoolButton CmdDelete 
      Height          =   330
      Left            =   7275
      TabIndex        =   19
      Top             =   2160
      Width           =   1140
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "FrmAreaTailMethod.frx":051C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton CmdAdd 
      Height          =   330
      Left            =   7275
      TabIndex        =   17
      Top             =   1185
      Width           =   1140
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "FrmAreaTailMethod.frx":0538
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdClose 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   7275
      TabIndex        =   20
      Top             =   2610
      Width           =   1140
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
      MICON           =   "FrmAreaTailMethod.frx":0554
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton CmdUpdate 
      Height          =   330
      Left            =   7275
      TabIndex        =   18
      Top             =   1680
      Width           =   1140
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "修改(&U)"
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
      MICON           =   "FrmAreaTailMethod.frx":0570
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cboMethod 
      Height          =   300
      Left            =   5130
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   1485
   End
   Begin VB.ListBox lstArea 
      Appearance      =   0  'Flat
      Height          =   1080
      Left            =   120
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   720
      Width           =   1500
   End
   Begin VB.ListBox lstTail 
      Appearance      =   0  'Flat
      Height          =   1080
      Left            =   5610
      Style           =   1  'Checkbox
      TabIndex        =   11
      Top             =   720
      Width           =   1425
   End
   Begin VSFlex7LCtl.VSFlexGrid VSAllTailMethod 
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   6945
      _cx             =   12250
      _cy             =   2355
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
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
   Begin VSFlex7LCtl.VSFlexGrid VSArea 
      Height          =   1995
      Left            =   120
      TabIndex        =   15
      Top             =   3825
      Width           =   6960
      _cx             =   12277
      _cy             =   3519
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   9
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
   Begin VB.Label lblExcuteTable 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "票价表:"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   195
      Width           =   630
   End
   Begin VB.Label lblBusType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次类型(&B):"
      Height          =   180
      Left            =   1740
      TabIndex        =   6
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label lblPriceItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "票价项(&P):"
      Height          =   180
      Left            =   3480
      TabIndex        =   8
      Top             =   480
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "区域尾数处理:"
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   3585
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "尾数处理方式具体情况:"
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1890
   End
   Begin VB.Label lblTail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "尾数位(&T):"
      Height          =   180
      Left            =   5670
      TabIndex        =   10
      Top             =   480
      Width           =   900
   End
   Begin VB.Label lblMethod 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "尾数处理方式(&M):"
      Height          =   180
      Left            =   3540
      TabIndex        =   2
      Top             =   180
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "区域(&A):"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   720
   End
End
Attribute VB_Name = "frmAreaTailMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const szPriceTotal = "总票价"
Private m_oTicketPriceMan As New TicketPriceMan

Private Sub LoadAllTailDealMethod()
Dim szTemp As String
Dim tTTailDealMethod() As TTailDealMethod
Dim nCount As Integer
Dim i As Integer

On Error GoTo ErrorHandle

   tTTailDealMethod = m_oTicketPriceMan.GetAllTailDealMethod()
   nCount = ArrayLength(tTTailDealMethod)
   VSAllTailMethod.Rows = nCount + 1
   For i = 1 To nCount
       VSAllTailMethod.TextMatrix(i, 1) = tTTailDealMethod(i).szMethodNo
       VSAllTailMethod.TextMatrix(i, 2) = tTTailDealMethod(i).szSerialNo
       VSAllTailMethod.TextMatrix(i, 3) = tTTailDealMethod(i).nStartNumber
       VSAllTailMethod.TextMatrix(i, 4) = tTTailDealMethod(i).nEndNumber
       If tTTailDealMethod(i).nDealedNumber <> 100 Then
          VSAllTailMethod.TextMatrix(i, 5) = tTTailDealMethod(i).nDealedNumber
       Else
          VSAllTailMethod.TextMatrix(i, 5) = "原值"
       End If
       VSAllTailMethod.TextMatrix(i, 6) = tTTailDealMethod(i).szAnnotation
   Next
   Exit Sub

ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub CboPriceTable_Click()
    FillPriceItem
End Sub

Private Sub cmdAdd_Click()
Dim tAreaMethod() As TAreaTailDealMethod
Dim nMethodCount As Integer
Dim i As Integer
Dim ttArea() As TArea
Dim szPriceItem() As String
Dim tTailBit() As Integer
Dim szTemp As String
Dim nSelectCount As Integer
Dim k As Integer
Dim nTail() As Integer
Dim szArea() As String
Dim szBusType() As String

On Error GoTo ErrorHandle

    If cboMethod.Text = "" Then
       MsgBox "请选择尾数处理方式中其中一种！", vbOKOnly + vbInformation, Me.Caption
       Exit Sub
    End If
    For i = 1 To lstArea.ListCount
        If lstArea.Selected(i - 1) = True Then
            nSelectCount = nSelectCount + 1
        End If
    Next
    If nSelectCount = 0 Then
       MsgBox "请选择一个或几个地区！", vbOKOnly + vbInformation, Me.Caption
       Exit Sub
    End If
    ReDim ttArea(1 To nSelectCount)
    ReDim szArea(1 To nSelectCount)
    k = 1
    For i = 1 To lstArea.ListCount
        If lstArea.Selected(i - 1) = True Then
           ttArea(k).szAreaCode = ResolveDisplay(lstArea.List(i - 1))
           ttArea(k).szAreaName = GetAreaName(lstArea.List(i - 1))
           szArea(k) = ttArea(k).szAreaCode
           k = k + 1
        End If
    Next

    nSelectCount = 0
    For i = 1 To lstBusType.ListCount
        If lstBusType.Selected(i - 1) = True Then
            nSelectCount = nSelectCount + 1
        End If
    Next
    If nSelectCount = 0 Then
       MsgBox "请选择一个或几个车次类型！", vbOKOnly + vbInformation, Me.Caption
       Exit Sub
    End If
    ReDim szBusType(1 To nSelectCount)
    k = 1
    For i = 1 To lstBusType.ListCount
        If lstBusType.Selected(i - 1) = True Then
           szBusType(k) = ResolveDisplay(lstBusType.List(i - 1))
           k = k + 1
        End If
    Next

    nSelectCount = 0
    If lstTail.Selected(lstTail.ListCount - 1) = False Then
        nSelectCount = 0
        For i = 1 To lstPriceItem.ListCount
            If lstPriceItem.Selected(i - 1) = True Then
                nSelectCount = nSelectCount + 1
            End If
        Next
        If nSelectCount = 0 Then
           MsgBox "请选择一个或几个票价项！", vbOKOnly + vbInformation, Me.Caption
           Exit Sub
        End If
        ReDim szPriceItem(1 To nSelectCount)
        k = 1
        For i = 1 To lstPriceItem.ListCount
            If lstPriceItem.Selected(i - 1) = True Then
               szPriceItem(k) = ResolveDisplay(lstPriceItem.List(i - 1))
               k = k + 1
            End If
        Next
    Else
       ReDim szPriceItem(1 To 1)
       szPriceItem(1) = cszItemBaseCarriage
    End If

    nSelectCount = 0
    For i = 1 To lstTail.ListCount
        If lstTail.Selected(i - 1) = True Then
            nSelectCount = nSelectCount + 1
        End If
    Next
    If nSelectCount = 0 Then
       MsgBox "请选择一个或几个尾数！", vbOKOnly + vbInformation, Me.Caption
       Exit Sub
    End If
    ReDim nTail(1 To nSelectCount)
    k = 1
    For i = 0 To lstTail.ListCount - 1
        If lstTail.Selected(i) = True Then
           If lstTail.List(i) = "第一位（角）" Then
              nTail(k) = 1
           ElseIf lstTail.List(i) = "第二位（分）" Then
              nTail(k) = 2
           ElseIf lstTail.List(i) = "第三位（厘）" Then
              nTail(k) = 3
           Else
              nTail(k) = 0
           End If
           k = k + 1
        End If
    Next

    SetBusy
    m_oTicketPriceMan.AddAreaMethod ResolveDisplay(cboPriceTable.Text), szArea, szBusType, szPriceItem, nTail, cboMethod.Text
    cmdFind_Click
    SetNormal
    MsgBox "成功增加区域尾数处理！", vbOKOnly + vbInformation, "提示！"
    Exit Sub

ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim tAreaMethod() As TAreaTailDealMethod
Dim nMethodCount As Integer
Dim i As Integer
Dim ttArea() As TArea
Dim tTailBit() As Integer
Dim szTemp As String
Dim nSelectCount As Integer
Dim k As Integer
Dim nTail() As Integer
Dim szArea() As String
Dim nNullCount As Integer
Dim szPriceItem() As String
Dim szBusType() As String

On Error GoTo ErrorHandle
    If lstArea.SelCount = 0 Then
       MsgBox "请选择一个或几个地区！", vbOKOnly + vbInformation, Me.Caption

       Exit Sub
    End If

    SetBusy
    nNullCount = 0
    For i = 1 To lstArea.ListCount
        If lstArea.Selected(i - 1) = True Then
            nSelectCount = nSelectCount + 1
        End If
    Next

    If nSelectCount > 0 Then
        ReDim ttArea(1 To nSelectCount)
        ReDim szArea(1 To nSelectCount)
    Else
        nNullCount = nNullCount + 1
    End If
    k = 1
    For i = 1 To lstArea.ListCount
        If lstArea.Selected(i - 1) = True Then
           ttArea(k).szAreaCode = ResolveDisplay(lstArea.List(i - 1))
           ttArea(k).szAreaName = GetAreaName(lstArea.List(i - 1))
           szArea(k) = ttArea(k).szAreaCode
           k = k + 1
        End If
    Next

    nSelectCount = 0
    For i = 1 To lstBusType.ListCount
        If lstBusType.Selected(i - 1) = True Then
            nSelectCount = nSelectCount + 1
        End If
    Next

    If nSelectCount > 0 Then
        ReDim szBusType(1 To nSelectCount)
    Else
        nNullCount = nNullCount + 1
    End If

    k = 1
    For i = 1 To lstBusType.ListCount
        If lstBusType.Selected(i - 1) = True Then
           szBusType(k) = ResolveDisplay(lstBusType.List(i - 1))
           k = k + 1
        End If
    Next

    'nNullCount = 0
    nSelectCount = 0
    For i = 1 To lstPriceItem.ListCount
        If lstPriceItem.Selected(i - 1) = True Then
            nSelectCount = nSelectCount + 1
        End If
    Next

    If nSelectCount > 0 Then
        ReDim szPriceItem(1 To nSelectCount)
    Else
        nNullCount = nNullCount + 1
    End If
    k = 1
    For i = 1 To lstPriceItem.ListCount
        If lstPriceItem.Selected(i - 1) = True Then
           szPriceItem(k) = ResolveDisplay(lstPriceItem.List(i - 1))
           k = k + 1
        End If
    Next

    nSelectCount = 0
    For i = 1 To lstTail.ListCount
        If lstTail.Selected(i - 1) = True Then
            nSelectCount = nSelectCount + 1
        End If
    Next
    If nSelectCount > 0 Then
       ReDim nTail(1 To nSelectCount)
    Else
       nNullCount = nNullCount + 1
    End If
    k = 1
    For i = 0 To lstTail.ListCount - 1
        If lstTail.Selected(i) = True Then
           If lstTail.List(i) = "第一位（角）" Then
              nTail(k) = 1
           ElseIf lstTail.List(i) = "第二位（分）" Then
              nTail(k) = 2
           ElseIf lstTail.List(i) = "第三位（厘）" Then
              nTail(k) = 3
           Else
              nTail(k) = 0
           End If
           k = k + 1
        End If
    Next
    If cboMethod.Text = "" Then
       nNullCount = nNullCount + 1
    End If
    If nNullCount < 4 Then
       m_oTicketPriceMan.DeleteAreaMethod ResolveDisplay(cboPriceTable.Text), szArea, szBusType, szPriceItem, nTail, cboMethod.Text
    Else
       SetNormal
       MsgBox "请选择区域、车次类型、票价项、尾数位、尾数处理方式五者至少其一进行区域尾数处理方式删除！", vbOKOnly + vbInformation, Me.Caption
       SetBusy
    End If
    cmdFind_Click
    MsgBox "成功删除区域尾数处理！", vbOKOnly + vbInformation, "提示！"
    SetNormal
    Exit Sub

ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub cmdFind_Click()
Dim tAreaMethod() As TAreaTailDealMethod
Dim nMethodCount As Integer
Dim i As Integer
Dim ttArea() As TArea
Dim tTailBit() As Integer
Dim szTemp As String
Dim nSelectCount As Integer
Dim k As Integer
Dim nTail() As Integer
Dim szPriceItem() As String
Dim szBusType() As String

On Error GoTo ErrorHandle
    For i = 1 To lstArea.ListCount
        If lstArea.Selected(i - 1) = True Then
            nSelectCount = nSelectCount + 1
        End If
    Next
    If nSelectCount > 0 Then ReDim ttArea(1 To nSelectCount)
    k = 1
    For i = 1 To lstArea.ListCount
        If lstArea.Selected(i - 1) = True Then
           ttArea(k).szAreaCode = ResolveDisplay(lstArea.List(i - 1))
           ttArea(k).szAreaName = GetAreaName(lstArea.List(i - 1))
           k = k + 1
        End If
    Next

    For i = 1 To lstBusType.ListCount
        If lstBusType.Selected(i - 1) = True Then
            nSelectCount = nSelectCount + 1
        End If
    Next
    If nSelectCount > 0 Then ReDim szBusType(1 To nSelectCount)
    k = 1
    For i = 1 To lstBusType.ListCount
        If lstBusType.Selected(i - 1) = True Then
           szBusType(k) = ResolveDisplay(lstBusType.List(i - 1))
           k = k + 1
        End If
    Next

    If lstTail.Selected(lstTail.ListCount - 1) = False Then
        nSelectCount = 0
        For i = 1 To lstPriceItem.ListCount
            If lstPriceItem.Selected(i - 1) = True Then
                nSelectCount = nSelectCount + 1
            End If
        Next
        If nSelectCount > 0 Then ReDim szPriceItem(1 To nSelectCount)
        k = 1
        For i = 1 To lstPriceItem.ListCount
            If lstPriceItem.Selected(i - 1) = True Then
               szPriceItem(k) = ResolveDisplay(lstPriceItem.List(i - 1))
               k = k + 1
            End If
        Next
    Else
        ReDim szPriceItem(1 To 1)
        szPriceItem(1) = cszItemBaseCarriage
    End If

    nSelectCount = 0
    For i = 1 To lstTail.ListCount
        If lstTail.Selected(i - 1) = True Then
            nSelectCount = nSelectCount + 1
        End If
    Next
    If nSelectCount > 0 Then ReDim nTail(1 To nSelectCount)
    k = 1
    For i = 0 To lstTail.ListCount - 1
        If lstTail.Selected(i) = True Then
           If lstTail.List(i) = "第一位（角）" Then
              nTail(k) = 1
           ElseIf lstTail.List(i) = "第二位（分）" Then
              nTail(k) = 2
           ElseIf lstTail.List(i) = "第三位（厘）" Then
              nTail(k) = 3
           Else
              nTail(k) = 0
           End If
           k = k + 1
        End If
    Next

    tAreaMethod = m_oTicketPriceMan.GetAllAreaTailMethod(ResolveDisplay(cboPriceTable.Text), ttArea, szBusType, szPriceItem, nTail, cboMethod.Text)
    nMethodCount = ArrayLength(tAreaMethod)
    If nMethodCount > 0 Then
       VSArea.Rows = nMethodCount + 1
       For i = 1 To VSArea.Rows - 1
           VSArea.TextMatrix(i, 1) = tAreaMethod(i).szAreaCode
           VSArea.TextMatrix(i, 2) = tAreaMethod(i).szAreaName
           VSArea.TextMatrix(i, 3) = tAreaMethod(i).szBusType
           If tAreaMethod(i).nTailBitNo = 0 Then
                VSArea.TextMatrix(i, 4) = szPriceTotal
                VSArea.TextMatrix(i, 5) = szPriceTotal
           Else
                VSArea.TextMatrix(i, 4) = tAreaMethod(i).szPriceItemCode
                VSArea.TextMatrix(i, 5) = tAreaMethod(i).szPriceItemName
           End If
           VSArea.TextMatrix(i, 6) = tAreaMethod(i).nTailBitNo

           VSArea.TextMatrix(i, 7) = tAreaMethod(i).szTailBitNoName
           VSArea.TextMatrix(i, 8) = tAreaMethod(i).szMethodNo
       Next
    Else
       VSArea.Rows = 1
    End If
    Exit Sub
ErrorHandle:
    SetNormal
End Sub

Private Sub CmdUpdate_Click()
Dim tAreaMethod() As TAreaTailDealMethod
Dim nMethodCount As Integer
Dim i As Integer
Dim ttArea() As TArea
Dim szPriceItem() As String
Dim tTailBit() As Integer
Dim szTemp As String
Dim nSelectCount As Integer
Dim k As Integer
Dim nTail() As Integer
Dim n As Integer
Dim szBusType() As String
Dim m As Integer
Dim j As Integer
On Error GoTo ErrorHandle

    If cboMethod.Text = "" Then
       MsgBox "请选择尾数处理方式中其中一种！", vbOKOnly + vbInformation, Me.Caption
       Exit Sub
    End If
    For i = 1 To lstArea.ListCount
        If lstArea.Selected(i - 1) = True Then
            nSelectCount = nSelectCount + 1
        End If
    Next
    If nSelectCount = 0 Then
       MsgBox "请选择一个或几个地区！", vbOKOnly + vbInformation, Me.Caption
       Exit Sub
    End If

    ReDim ttArea(1 To nSelectCount)
    k = 1
    For i = 1 To lstArea.ListCount
        If lstArea.Selected(i - 1) = True Then
           ttArea(k).szAreaCode = ResolveDisplay(lstArea.List(i - 1))
           ttArea(k).szAreaName = GetAreaName(lstArea.List(i - 1))
           k = k + 1
        End If
    Next

    nSelectCount = 0
    For i = 1 To lstBusType.ListCount
        If lstBusType.Selected(i - 1) = True Then
            nSelectCount = nSelectCount + 1
        End If
    Next
    If nSelectCount = 0 Then
       MsgBox "请选择一个或几个车次类型！", vbOKOnly + vbInformation, Me.Caption
       Exit Sub
    End If
    ReDim szBusType(1 To nSelectCount)
    k = 1
    For i = 1 To lstBusType.ListCount
        If lstBusType.Selected(i - 1) = True Then
           szBusType(k) = ResolveDisplay(lstBusType.List(i - 1))
           k = k + 1
        End If
    Next

    If lstTail.Selected(lstTail.ListCount - 1) = False Then
        nSelectCount = 0
        For i = 1 To lstPriceItem.ListCount
            If lstPriceItem.Selected(i - 1) = True Then
                nSelectCount = nSelectCount + 1
            End If
        Next
        If nSelectCount = 0 Then
           MsgBox "请选择一个或几个票价项！", vbOKOnly + vbInformation, Me.Caption
           Exit Sub
        End If
        ReDim szPriceItem(1 To nSelectCount)
        k = 1
        For i = 1 To lstPriceItem.ListCount
            If lstPriceItem.Selected(i - 1) = True Then
               szPriceItem(k) = ResolveDisplay(lstPriceItem.List(i - 1))
               k = k + 1
            End If
        Next
    Else
        ReDim szPriceItem(1 To 1)
        szPriceItem(1) = cszItemBaseCarriage
    End If

    nSelectCount = 0
    For i = 1 To lstTail.ListCount
        If lstTail.Selected(i - 1) = True Then
            nSelectCount = nSelectCount + 1
        End If
    Next
    If nSelectCount = 0 Then
       MsgBox "请选择一个或几个尾数！", vbOKOnly + vbInformation, Me.Caption
       Exit Sub
    End If
    ReDim nTail(1 To nSelectCount)
    k = 1
    SetBusy
    For i = 0 To lstTail.ListCount - 1
        If lstTail.Selected(i) = True Then
           If lstTail.List(i) = "第一位（角）" Then
              nTail(k) = 1
           ElseIf lstTail.List(i) = "第二位（分）" Then
              nTail(k) = 2
           ElseIf lstTail.List(i) = "第三位（厘）" Then
              nTail(k) = 3
           Else
              nTail(k) = 0
           End If
           k = k + 1
        End If
    Next

    szTemp = ResolveDisplay(cboPriceTable.Text)
    For i = 1 To ArrayLength(ttArea)
        For m = 1 To ArrayLength(szBusType)
            For n = 1 To ArrayLength(szPriceItem)
                For j = 1 To ArrayLength(nTail)
                    m_oTicketPriceMan.UpdateAreaMethod szTemp, ttArea(i).szAreaCode, szBusType(m), szPriceItem(n), nTail(j), cboMethod.Text
                Next
            Next
        Next
    Next

    cmdFind_Click
    SetNormal
    MsgBox "成功更新区域尾数处理！", vbOKOnly + vbInformation, "提示！"
    Exit Sub

ErrorHandle:
    ShowErrorMsg
    SetNormal
End Sub

Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_Load()
        
    
    m_oTicketPriceMan.Init g_oActiveUser
'    m_oTicketPriceMan.ObjStatus = ST_NormalObj

    With VSAllTailMethod
         .ColWidth(0) = 100
         .ColWidth(1) = 1000
         .ColWidth(2) = 500
         .ColWidth(3) = 1000
         .ColWidth(4) = 1000
         .ColWidth(5) = 1000
         .ColWidth(6) = 2000

         .TextMatrix(0, 1) = "方式编号"
         .TextMatrix(0, 2) = "序号"
         .TextMatrix(0, 3) = "起始数字"
         .TextMatrix(0, 4) = "结束数字"
         .TextMatrix(0, 5) = "处理值"
         .TextMatrix(0, 6) = "备   注"

         .FixedAlignment(1) = flexAlignCenterCenter
         .FixedAlignment(2) = flexAlignCenterCenter
         .FixedAlignment(3) = flexAlignCenterCenter
         .FixedAlignment(4) = flexAlignCenterCenter
         .FixedAlignment(5) = flexAlignCenterCenter
         .FixedAlignment(6) = flexAlignCenterCenter

         .MergeCol(1) = True
         .MergeCol(6) = True
         .MergeCells = flexMergeFree
    End With

     LoadAllTailDealMethod

    With VSArea
        .ColWidth(0) = 100
        .ColWidth(1) = 0 '1000
        .ColWidth(2) = 1500
        .ColWidth(3) = 1000
        .ColWidth(4) = 0 '1000
        .ColWidth(5) = 1500
        .ColWidth(6) = 0 '1000
        .ColWidth(7) = 1000
        .ColWidth(8) = 1000
        .TextMatrix(0, 1) = "区域编号"
        .TextMatrix(0, 2) = "区域名称"
        .TextMatrix(0, 3) = "类型代码"
        .TextMatrix(0, 4) = "票价项代码"
        .TextMatrix(0, 5) = "票价项名称"
        .TextMatrix(0, 6) = "尾数位"
        .TextMatrix(0, 7) = "尾数名称"
        .TextMatrix(0, 8) = "处理方法"
        .FixedAlignment(1) = flexAlignCenterCenter
        .FixedAlignment(2) = flexAlignCenterCenter
        .FixedAlignment(3) = flexAlignCenterCenter
        .FixedAlignment(4) = flexAlignCenterCenter
        .FixedAlignment(5) = flexAlignCenterCenter
        .FixedAlignment(6) = flexAlignCenterCenter
        .FixedAlignment(7) = flexAlignCenterCenter
        .FixedAlignment(8) = flexAlignCenterCenter
    End With

    lstTail.AddItem "第一位（角）"
    lstTail.AddItem "第二位（分）"
    lstTail.AddItem "第三位（厘）"
    lstTail.AddItem "票 价 总 和 "
    FillPriceTable
    FillMethod
    FillArea
    FillBusType
End Sub

Public Sub FillArea()
    Dim i As Integer
    Dim oBase As New BaseInfo
    Dim aszTemp() As String
    Dim nCount As Integer
    oBase.Init g_oActiveUser
    aszTemp = oBase.GetAllArea()
    nCount = ArrayLength(aszTemp)
    For i = 1 To nCount
        lstArea.AddItem MakeDisplayString(aszTemp(i, 1), aszTemp(i, 2))
        
    Next
End Sub

Public Sub FillMethod()
Dim i As Integer
Dim nMethodCount As Integer
Dim ttMethod() As TTailDealMethod

    ttMethod = m_oTicketPriceMan.GetAllTailDealMethod()
    cboMethod.AddItem ""
    nMethodCount = ArrayLength(ttMethod)
    If nMethodCount > 0 Then
        For i = 1 To nMethodCount
            If i = 1 Then
               cboMethod.AddItem ttMethod(i).szMethodNo
            ElseIf i > 1 Then
               If ttMethod(i - 1).szMethodNo <> ttMethod(i).szMethodNo Then cboMethod.AddItem ttMethod(i).szMethodNo
            End If
        Next
        cboMethod.ListIndex = 0
    End If
End Sub

Public Function GetAreaName(szValue As String) As String
Dim i As Integer
Dim szTemp As String
Dim nTemp As Integer

On Error GoTo ErrorHandle

    For i = 1 To Len(szValue)
        If Mid(szValue, i, 1) = "[" Then
           nTemp = i
           Exit For
        End If
    Next
    GetAreaName = Mid(szValue, nTemp + 1, Len(szValue) - i - 1)
ErrorHandle:
End Function


Public Sub FillPriceItem()
    '填充使用的票价项
    Dim szPriceItem() As String
    Dim i As Integer
    Dim Count As Integer

    lstPriceItem.Clear
    szPriceItem = m_oTicketPriceMan.GetAllTicketItem
    Count = ArrayLength(szPriceItem)
    For i = 1 To Count
        If szPriceItem(i, 3) = 1 Then '使用票价项
           lstPriceItem.AddItem MakeDisplayString(szPriceItem(i, 1), szPriceItem(i, 2))
        End If
    Next
    lstPriceItem.ListIndex = 0
End Sub

Private Sub lstPriceItem_Click()
   lstTail.Selected(lstTail.ListCount - 1) = False
End Sub

Private Sub lstTail_Click()
Dim i As Integer

    If lstTail.ListIndex = lstTail.ListCount - 1 Then
       For i = 0 To lstTail.ListCount - 2
           lstTail.Selected(i) = False
       Next
       For i = 0 To lstPriceItem.ListCount - 1
           lstPriceItem.Selected(i) = False
       Next
    ElseIf lstTail.ListIndex < lstTail.ListCount - 1 Then
       lstTail.Selected(lstTail.ListCount - 1) = False
    End If
End Sub

Public Sub FillBusType()
On Error GoTo ErrorHandle

    Dim oBase As New BaseInfo
    Dim szTemp() As String
    Dim i As Integer
    Dim nTemp As Integer

    oBase.Init g_oActiveUser
    szTemp = oBase.GetAllBusType

    nTemp = ArrayLength(szTemp)
    For i = 1 To nTemp
        lstBusType.AddItem MakeDisplayString(szTemp(i, 1), szTemp(i, 2))
    Next i
    lstBusType.AddItem MakeDisplayString(cnAllBusType, cszAllBusType)
    Exit Sub

ErrorHandle:
    MsgBox err.Description, vbOKOnly
End Sub

Private Sub FillPriceTable()
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
        Next
    End If
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

