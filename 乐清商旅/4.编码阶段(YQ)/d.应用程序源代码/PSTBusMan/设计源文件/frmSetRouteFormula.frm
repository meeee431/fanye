VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmSetRouteFormula 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置线路的票价公式"
   ClientHeight    =   5115
   ClientLeft      =   2655
   ClientTop       =   2445
   ClientWidth     =   7800
   HelpContextID   =   10000580
   Icon            =   "frmSetRouteFormula.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VSFlex7LCtl.VSFlexGrid vsRoute 
      Height          =   4875
      Left            =   150
      TabIndex        =   3
      Top             =   120
      Width           =   6195
      _cx             =   10936
      _cy             =   8599
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14737632
      BackColorAlternate=   16777215
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
      Rows            =   1
      Cols            =   3
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
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   345
      Left            =   6495
      TabIndex        =   2
      Top             =   1095
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmSetRouteFormula.frx":014A
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
      Left            =   6495
      TabIndex        =   1
      Top             =   615
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
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
      MICON           =   "frmSetRouteFormula.frx":0166
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
      Default         =   -1  'True
      Height          =   345
      Left            =   6495
      TabIndex        =   0
      Top             =   165
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmSetRouteFormula.frx":0182
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
Attribute VB_Name = "frmSetRouteFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'* Source File Name:frmSetRouteFormula.frm
'* Project Name:PSTBusMan.vbp
'* Engineer:陈峰
'* Date Generated:2002/09/03
'* Last Revision Date:2002/09/03
'* Brief Description:设置线路的票价公式
'* Relational Document:
'**********************************************************

Option Explicit
Const cszDefaultFormula = "缺省公式"
Const cszDisCountFormula = "不计算"
Const cszModifyColor = &HFFC0C0
Const cszNormalColor = &HFFFFFF

Private m_aszRouteInfo() As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdOk_Click()
    On Error GoTo ErrorHandle
    Update
    Unload Me
    Exit Sub
ErrorHandle:
    vsRoute.Redraw = True
    ShowErrorMsg
End Sub

Private Sub vsRoute_EnterCell()
    If vsRoute.Col < 2 Then
        vsRoute.Editable = flexEDNone
    Else
        vsRoute.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub vsRoute_LeaveCell()
Dim i As Integer
On Error GoTo ErrorHandle

    vsRoute.Editable = flexEDNone
    If vsRoute.Col = 2 Then
       If vsRoute.TextMatrix(vsRoute.Row, 0) = RTrim(m_aszRouteInfo(vsRoute.Row, 1)) Then
          If RTrim(m_aszRouteInfo(vsRoute.Row, 7)) <> "" Then
             If vsRoute.TextMatrix(vsRoute.Row, vsRoute.Col) <> RTrim(m_aszRouteInfo(vsRoute.Row, 7)) Then
                 vsRoute.CellBackColor = cszModifyColor
             Else
                 vsRoute.CellBackColor = cszNormalColor
             End If
          Else
             If vsRoute.TextMatrix(vsRoute.Row, vsRoute.Col) <> cszDefaultFormula Then
                vsRoute.CellBackColor = cszModifyColor
             Else
                vsRoute.CellBackColor = cszNormalColor
             End If
          End If
        End If
    End If
ErrorHandle:
End Sub


Private Sub Form_Load()
    ShowRouteInfo
    FillFormula
End Sub

Private Sub ShowRouteInfo()
    '显示所有线路的公式信息
    Dim oBase As New BaseInfo
    Dim szTemp As String

    Dim i As Integer, nCount As Integer
    On Error GoTo ErrorHandle

    vsRoute.Rows = 1

    vsRoute.TextMatrix(0, 0) = "线路代码"
    vsRoute.TextMatrix(0, 1) = "线路名称"
    vsRoute.TextMatrix(0, 2) = "线路票价公式"
    vsRoute.ColWidth(1) = vsRoute.ColWidth(0) * 1.5
    vsRoute.ColWidth(2) = vsRoute.ColWidth(0) * 4
    oBase.Init g_oActiveUser

    m_aszRouteInfo = oBase.GetRoute()
    nCount = ArrayLength(m_aszRouteInfo)
    vsRoute.Rows = 1 + nCount
    For i = 1 To nCount
        vsRoute.TextMatrix(i, 0) = RTrim(m_aszRouteInfo(i, 1))
        vsRoute.TextMatrix(i, 1) = RTrim(m_aszRouteInfo(i, 2))
        szTemp = RTrim(m_aszRouteInfo(i, 7))
        If szTemp = "" Then
            vsRoute.TextMatrix(i, 2) = cszDefaultFormula
        Else
            vsRoute.TextMatrix(i, 2) = szTemp
        End If
    Next
    Set oBase = Nothing
    Exit Sub
ErrorHandle:
    ShowErrorMsg
    Set oBase = Nothing
End Sub

Private Sub FillFormula()
    '填充所有公式
    Dim i As Integer, nCount As Integer
    Dim aszAllFormula() As String
    Dim szTemp As String
    Dim oTicketPriceMan As New TicketPriceMan
    
    On Error GoTo ErrorHandle
    oTicketPriceMan.Init g_oActiveUser
    aszAllFormula = oTicketPriceMan.GetAllTicketPriceFormula()
    nCount = ArrayLength(aszAllFormula)
     szTemp = "|" & szTemp & cszDefaultFormula & "|"
    For i = 1 To nCount
         szTemp = szTemp & RTrim(aszAllFormula(i, 1)) & "|"
    Next
    szTemp = "|" & szTemp & cszDisCountFormula
    vsRoute.ColComboList(2) = szTemp
    Set oTicketPriceMan = Nothing
    Exit Sub

ErrorHandle:
    Set oTicketPriceMan = Nothing
    ShowErrorMsg
End Sub


Private Sub Update()
    Dim i As Integer
    Dim oRoute As New Route
    Dim lOrgRow As Long, lOrgCol As Long
    Dim oBase As New BaseInfo

    On Error GoTo ErrorHandle


    oRoute.Init g_oActiveUser

    With vsRoute
        lOrgRow = .Row
        lOrgCol = .Col

        .Redraw = False
        .Col = 2
        For i = 1 To .Rows - 1
            .Row = i
            If .CellBackColor = cszModifyColor Then
                oRoute.Identify .TextMatrix(i, 0)
                If .Text = cszDefaultFormula Then
                    oRoute.TicketPriceFormula = ""
                Else
                    oRoute.TicketPriceFormula = .Text
                End If
                oRoute.Update
            End If
        Next
        .Row = lOrgRow
        .Col = lOrgCol
        .Redraw = True

        Set oRoute = Nothing
        oBase.Init g_oActiveUser
        m_aszRouteInfo = oBase.GetRoute()
        Set oBase = Nothing
        For i = 1 To .Rows - 1
            .Row = i
        Next
    End With
    Set oRoute = Nothing
    Set oBase = Nothing
    Exit Sub
ErrorHandle:
    Set oRoute = Nothing
    Set oBase = Nothing
    err.Raise err.Number
End Sub
