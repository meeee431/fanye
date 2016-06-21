VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmPriceItem 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "定义票价项"
   ClientHeight    =   3915
   ClientLeft      =   4380
   ClientTop       =   2955
   ClientWidth     =   5070
   HelpContextID   =   10000450
   Icon            =   "frmPriceItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   3450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceItem.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceItem.frx":025E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   315
      Left            =   3690
      TabIndex        =   4
      Top             =   3495
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
      MICON           =   "frmPriceItem.frx":07B2
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
      Height          =   315
      Left            =   2475
      TabIndex        =   3
      Top             =   3495
      Width           =   1095
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
      MICON           =   "frmPriceItem.frx":07CE
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
      Height          =   315
      Left            =   1260
      TabIndex        =   2
      Top             =   3495
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "确定(&O)"
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
      MICON           =   "frmPriceItem.frx":07EA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid vsItem 
      Height          =   3000
      Left            =   135
      TabIndex        =   1
      Top             =   300
      Width           =   4800
      _cx             =   5080
      _cy             =   5080
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
      Rows            =   17
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
   Begin VB.Image Image1 
      Height          =   240
      Left            =   3945
      Top             =   1815
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次票价项定义(&B):"
      Height          =   180
      Left            =   150
      TabIndex        =   0
      Top             =   45
      Width           =   1620
   End
End
Attribute VB_Name = "frmPriceItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *******************************************************************
' *  Source File Name  : frmPriceItem.frm                           *
' *  Project Name: PSTBusMan                                        *
' *  Engineer:                                                      *
' *  Date Generated: 2002/09/03                                     *
' *  Last Revision Date : 2002/09/03                                *
' *  Brief Description   :定义票价项                                *
' *******************************************************************
Option Explicit
Option Base 1

Dim aszItemInfo() As String
Dim bModifyBusTag(1 To 16) As Boolean '标注源数据库中是否使用,即是否可修改(True,可修改)
Dim bBusItemTags(1 To 16) As Boolean '标注cmdOk_Click前的设定
Dim nLBoundNum As Integer '票价项下限
Dim nModifyCount As Integer 'cmdOK前计算修改量

Const flexAlignCenterCenter = 4
Const cWhite = &H80000005
Const cGray = &H80000000
Const cBlue = &HFFC0C
Const cszFatalInfo = "对票价项的修改会影响整个数据库,"
Const cszFatalInfo1 = "新增 "
Const cszFatalInfo2 = " 项,确定?"

Private Sub CboPriceTable_Click()
    LoadInfo
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdOk_Click()
    Dim nMsg As Integer
    Dim i As Integer
    Dim oItem As New TicketItem
    Dim szTemp() As String
    Dim szPriceTableID As String
    
    SetBusy
    '验证有效性
    With vsItem
        For i = 1 To .Rows - 1
            .Col = 2
            .Row = i
            If .CellPicture = ImageList1.ListImages(2).Picture Then
                If Trim(.TextMatrix(i, 1)) = "" Then
                    .Col = 1
                    MsgBox "请输入票价项名称", vbInformation + vbOKOnly
                    SetNormal
                    Exit Sub
                End If
            End If
        Next
    End With
    '******此处应修改一下,只保存修改的东东
    '得到信息
    GetInfoFormUI
    nMsg = MsgBox(cszFatalInfo & vbCrLf & cszFatalInfo1 & CStr(nModifyCount) & cszFatalInfo2, vbYesNo + vbQuestion)
    If nMsg = vbYes Then
        On Error GoTo here
        oItem.Init g_oActiveUser
        For i = 1 To 16
            oItem.Identify szPriceTableID, aszItemInfo(i, 1)
            oItem.ItemName = aszItemInfo(i, 2)
            If aszItemInfo(i, 3) = 1 Then
                oItem.ItemUseMark = True
            Else
                oItem.ItemUseMark = False
            End If
            oItem.Update
        Next i
    End If
    SetNormal
    Set oItem = Nothing
    Unload Me
    
Exit Sub
here:
    SetNormal
    Set oItem = Nothing
    ShowErrorMsg
End Sub



Private Sub Form_Load()
    
    Dim i As Integer
    Dim szTemp As String
    With Me
        .Top = Screen.Height / 2 - Me.Height / 2
        .Left = Screen.Width / 2 - Me.Width / 2
    End With
    With vsItem
         .FormatString = "<票价项号|<票价项名"
         .Row = 0
         .Col = 0
         .Text = "票价项号"
         .Col = 1
         .Text = "票价项名"
         .Col = 2
         .Text = "使用标记"
         .ColWidth(2) = 820
         .ColWidth(1) = 1520
         .ColWidth(0) = 1500
         .Col = 0
         .Row = 1
         .Text = "基本票价项"
         For i = 2 To 16
            .Row = i
            szTemp = "车次票价项" & CStr(i - 1)
            .Text = szTemp
         Next i
         .ColAlignment(2) = flexAlignCenterCenter
    End With
    LoadInfo
    'FillPriceTable
End Sub


Private Sub vsItem_Click()
    If cmdOk.Enabled = False Then
        cmdOk.Enabled = True
        LoadInfo
    Else
        With vsItem
            If .Col = 2 Then
                If .Row < .Rows Or .Row > 0 Then
                    If bModifyBusTag(.Row) = True Then
                        bBusItemTags(.Row) = Not bBusItemTags(.Row)
                        If bBusItemTags(.Row) = True Then
                            Set .CellPicture = ImageList1.ListImages(2).Picture '标注为使用
                            .CellBackColor = cBlue
                        Else
                            Set .CellPicture = ImageList1.ListImages(1).Picture '标注为不使用
                            .CellBackColor = cWhite
                        End If
                    End If
                End If
            End If
        End With
    End If


End Sub



Private Sub LoadInfo()
    Dim i As Integer
    Dim oTicketPriceMan As New TicketPriceMan
    On Error GoTo here
    oTicketPriceMan.Init g_oActiveUser
    aszItemInfo = oTicketPriceMan.GetAllTicketItem()
    nLBoundNum = LBound(aszItemInfo)
    
    DisplayBus
Exit Sub
here:
    ShowErrorMsg
End Sub


Public Sub DisplayBus()
    Dim i As Integer
    Dim nTemp As Integer
    
    With vsItem
            .Col = 1
        For i = 1 To 16
            nTemp = CInt(aszItemInfo(i, 1)) '票价项编号
            If aszItemInfo(i, 3) = 1 Then
                bModifyBusTag(nTemp + 1) = False
            Else
                bModifyBusTag(nTemp + 1) = True
            End If
            bBusItemTags(nTemp + 1) = Not bModifyBusTag(nTemp + 1)
            .Row = nTemp + 1
            .Text = aszItemInfo(i, 2)
        Next i
        .Col = 2
        For i = 1 To 16
            .Row = i
            If bModifyBusTag(i) = False Then
                Set .CellPicture = ImageList1.ListImages(2).Picture
                .CellBackColor = cGray
            Else
                Set .CellPicture = ImageList1.ListImages(1).Picture
                .CellBackColor = cWhite
            End If
        Next i
    End With
        
End Sub

Private Sub GetInfoFormUI()
    Dim i As Integer, j As Integer
    nModifyCount = 0

    
    With vsItem
    For i = 1 To 16
        .Row = i
        .Col = 1
        
        For j = 1 To 16
            If CInt(aszItemInfo(j, 1)) + 1 = i Then
                aszItemInfo(j, 2) = .Text
                If bBusItemTags(i) = True Then
                    aszItemInfo(j, 3) = 1
                Else
                    aszItemInfo(j, 3) = 0
                End If
            End If
        Next j

        If bModifyBusTag(i) = bBusItemTags(i) Then
            nModifyCount = nModifyCount + 1
        End If
    Next i
    End With
    
End Sub

'Private Sub FillPriceTable()
'    Dim aszRoutePriceTable() As String
'    Dim i As Integer, nCount As Integer
'    Dim szPriceTable As String
'
'On Error GoTo ErrorHandle
'    aszRoutePriceTable = GetPriceTable(Now)
'    nCount = ArrayLength(aszRoutePriceTable)
'
'    CboPriceTable.Clear
'    If nCount > 0 Then
'        For i = 1 To nCount
'            szPriceTable = MakeDisplayString(aszRoutePriceTable(i, 1), aszRoutePriceTable(i, 2))
'            CboPriceTable.AddItem szPriceTable
'            If aszRoutePriceTable(i, 7) = cnRunTable Then CboPriceTable.Text = MakeDisplayString(aszRoutePriceTable(i, 1), aszRoutePriceTable(i, 2))
'        Next
'    End If
'
'
'    Exit Sub
'ErrorHandle:
'    showerrormsg
'End Sub
Private Sub vsItem_EnterCell()
    If vsItem.Col = 1 Then
        vsItem.Editable = flexEDKbdMouse
    Else
        vsItem.Editable = flexEDNone
    End If
End Sub
