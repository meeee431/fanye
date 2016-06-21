VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmProtocolItem 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "拆算项设置"
   ClientHeight    =   5355
   ClientLeft      =   2610
   ClientTop       =   2655
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8355
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   930
      Left            =   -60
      TabIndex        =   1
      Top             =   4545
      Width           =   9495
      Begin RTComctl3.CoolButton CoolButton2 
         Height          =   345
         Left            =   2340
         TabIndex        =   7
         Top             =   300
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
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
         MICON           =   "frmProtocolItem.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton CoolButton1 
         Height          =   345
         Left            =   3570
         TabIndex        =   6
         Top             =   300
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
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
         MICON           =   "frmProtocolItem.frx":001C
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
         Left            =   210
         TabIndex        =   2
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
         MICON           =   "frmProtocolItem.frx":0038
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
         Left            =   6015
         TabIndex        =   3
         Top             =   300
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
         MICON           =   "frmProtocolItem.frx":0054
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
         Left            =   4800
         TabIndex        =   4
         ToolTipText     =   "保存协议"
         Top             =   300
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
         MICON           =   "frmProtocolItem.frx":0070
         PICN            =   "frmProtocolItem.frx":008C
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
   Begin RTComctl3.FlatLabel lblProtocol 
      Height          =   405
      Left            =   1080
      TabIndex        =   0
      Top             =   135
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OutnerStyle     =   2
      BorderWidth     =   0
      BevelWidth      =   0
      HorizontalAlignment=   3
      NormTextColor   =   16711680
      Caption         =   "0001[常州公司协议]"
   End
   Begin VSFlex7LCtl.VSFlexGrid VsProtocolItem 
      Height          =   3855
      Left            =   105
      TabIndex        =   5
      Top             =   645
      Width           =   8175
      _cx             =   14420
      _cy             =   6800
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
      AllowUserResizing=   3
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
      FormatString    =   $"frmProtocolItem.frx":0426
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
   Begin RTComctl3.FlatLabel lblISProtocol 
      Height          =   315
      Left            =   6060
      TabIndex        =   8
      Top             =   225
      Width           =   405
      _ExtentX        =   714
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
      BorderWidth     =   0
      BevelWidth      =   0
      NormTextColor   =   16711680
      Caption         =   "否"
   End
   Begin RTComctl3.FlatLabel FlatLabel1 
      Height          =   345
      Left            =   4680
      TabIndex        =   9
      Top             =   195
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      Caption         =   "是否默认协议:"
   End
   Begin RTComctl3.FlatLabel FlatLabel 
      Height          =   285
      Left            =   180
      TabIndex        =   10
      Top             =   195
      Width           =   795
      _ExtentX        =   1402
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
      BevelWidth      =   0
      Caption         =   "协议:"
   End
End
Attribute VB_Name = "frmProtocolItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_oReport As New Report
Private m_oProtocol As New Protocol
Private atSplitItem() As TSplitItemInfo
Private atFormula() As String

Public m_bIsParent As Boolean '是否父窗体调用

Public m_eStatus As EFormStatus
Public m_szProtocolID As String
Public m_szProtocolName As String
Public m_bIsDefault As Boolean

Private m_atFormula() As String
Private bIsNull As Boolean

Const cnSplitItem = 1
Const cnFormulaName = 2
Const cnLimitCharge = 3
Const cnFormulaComment = 4
Const cnUpCharge = 5


Private m_aszFixFeeItem() As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
On Error GoTo ErrHandle
    Dim i As Integer
    If m_eStatus = AddStatus Then
        AddSplistItem
    Else
        '修改
        EditSplitItem
    End If
    Unload Me
    '将数组传入接口中
    '判断有无重复记录
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
Private Sub EditSplitItem()
    On Error GoTo err
    Dim i As Integer
    For i = 1 To VsProtocolItem.Rows - 1
        m_oProtocol.Init g_oActiveUser
        m_oProtocol.Identify ResolveDisplay(lblProtocol.Caption)
        If VsProtocolItem.TextMatrix(i, cnLimitCharge) <= 0 Then
            VsProtocolItem.TextMatrix(i, cnLimitCharge) = 0
        End If
        m_oProtocol.UpDateChargeItemInfo ResolveDisplay(VsProtocolItem.TextMatrix(i, cnSplitItem)), VsProtocolItem.TextMatrix(i, cnFormulaName), _
            VsProtocolItem.TextMatrix(i, cnLimitCharge), VsProtocolItem.TextMatrix(i, cnFormulaComment), VsProtocolItem.TextMatrix(i, cnUpCharge)
    Next i
    Exit Sub
err:
ShowErrorMsg

End Sub

Private Sub AddSplistItem()
    On Error GoTo err
    Dim i As Integer, j As Integer
    Dim atTemp() As TSplitItemInfo
    Dim aszTemp() As String
    Dim x As Integer
    m_oReport.Init g_oActiveUser
    atTemp = m_oReport.GetSplitItemInfo
    ReDim aszTemp(1 To VsProtocolItem.Rows - 1)
    For i = 1 To VsProtocolItem.Rows - 1
        m_oProtocol.Init g_oActiveUser
        m_oProtocol.Identify ResolveDisplay(lblProtocol.Caption)
        If VsProtocolItem.TextMatrix(i, 3) <= 0 Then
            VsProtocolItem.TextMatrix(i, 3) = 0
        End If
        m_oProtocol.AddChargeItemInfo ResolveDisplay(VsProtocolItem.TextMatrix(i, cnSplitItem)), VsProtocolItem.TextMatrix(i, cnFormulaName), _
                VsProtocolItem.TextMatrix(i, cnLimitCharge), VsProtocolItem.TextMatrix(i, cnFormulaComment), VsProtocolItem.TextMatrix(i, cnUpCharge)
        aszTemp(i) = ResolveDisplay(VsProtocolItem.TextMatrix(i, cnSplitItem))
        
    Next i
    '为了后面新增的拆算项用的
    For i = 1 To ArrayLength(atTemp)
        x = 0
        For j = 1 To VsProtocolItem.Rows - 1
            If ResolveDisplay(VsProtocolItem.TextMatrix(j, cnSplitItem)) = atTemp(i).SplitItemID Then
                x = 1
                Exit For
            End If
        Next j
        If x = 0 Then
            m_oProtocol.AddChargeItemInfo atTemp(i).SplitItemID, "", 0, "", 0
        End If
    Next i
    Exit Sub
err:
ShowErrorMsg
End Sub

Private Sub CoolButton1_Click()
    VsProtocolItem.Rows = VsProtocolItem.Rows + 1
End Sub

Private Sub CoolButton2_Click()
'    VsProtocolItem.Rows = VsProtocolItem.Rows - 1
    VsProtocolItem.RemoveItem (VsProtocolItem.Row)
End Sub

Private Sub FillCmdOK()
    If m_eStatus = AddStatus Then
        cmdOk.Caption = "新增(&A)"
    Else
        cmdOk.Caption = "修改(&E)"
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo err
    Dim atTemp() As TSplitItemInfo, atSplitItem() As TSplitItemInfo
'    Dim atFormulaTemp() As String, atFormula() As String
    Dim szTemp As String
    Dim i As Integer
    m_oReport.Init g_oActiveUser
    m_oProtocol.Init g_oActiveUser
    m_aszFixFeeItem = Split(g_szFixFeeItem, ",")
    
    
    FillVSHead
    
    FillCmdOK
    AlignFormPos Me
    
    m_oProtocol.Identify m_szProtocolID
    
    
    If m_oProtocol.DefaultMark = Default Then
        lblISProtocol.Caption = "是"
    Else
        lblISProtocol.Caption = "否"
    End If
    m_atFormula = m_oReport.GetAllFormula
    
    FillVsProtocolItem
    '填列表
    If m_eStatus = ModifyStatus Then
        lblProtocol.Caption = MakeDisplayString(m_oProtocol.ProtocolID, m_oProtocol.ProtocolName)
        GetProtocolItem
    Else
        lblProtocol.Caption = m_szProtocolID
        FillItem
    End If
    AlignHeadWidth Me.name, VsProtocolItem
    Exit Sub
err:
ShowErrorMsg
End Sub
Private Sub FillItem()
    On Error GoTo err
    Dim atTemp() As TSplitItemInfo
    Dim i As Integer, j As Integer
    Dim k As Integer
    Dim nCount As Integer
    
    nCount = ArrayLength(m_aszFixFeeItem)
    atTemp = m_oReport.GetSplitItemInfo
    j = 1
    If ArrayLength(atTemp) <> 0 Then
        For i = 1 To ArrayLength(atTemp)
            If atTemp(i).SplitStatus = CS_SplitItemUse Then
            
                '剔除固定费用的项
                For k = 0 To nCount - 1
                    If Val(atTemp(i).SplitItemID) = Val(m_aszFixFeeItem(k)) Then
                        Exit For
                    End If
                Next k
                '如果不是固定费用项,则显示
                If k > nCount - 1 Then
                    VsProtocolItem.Rows = VsProtocolItem.Rows + 1
                    VsProtocolItem.TextMatrix(j, cnSplitItem) = MakeDisplayString(atTemp(i).SplitItemID, atTemp(i).SplitItemName)
                    VsProtocolItem.TextMatrix(j, cnLimitCharge) = 0
                    VsProtocolItem.TextMatrix(j, cnUpCharge) = 0
                    j = j + 1
                End If
            End If
        Next i
        VsProtocolItem.RemoveItem (j)
'        bIsNull = True
    End If
    Exit Sub
err:
ShowErrorMsg
End Sub
Private Sub GetProtocolItem()
    On Error GoTo err
    Dim atTemp() As TFinChargeItemInfo
    Dim i As Integer
    Dim j As Integer
    Dim nCount As Integer
    Dim k As Integer
    nCount = ArrayLength(m_aszFixFeeItem)
    
    m_oProtocol.Init g_oActiveUser
    m_oProtocol.Identify (ResolveDisplay(lblProtocol.Caption))
    atTemp = m_oProtocol.GetChargeitemInfo
    If ArrayLength(atTemp) <> 0 Then
        j = 0
        For i = 1 To ArrayLength(atTemp)
            
            If atTemp(i).SplitStatus = CS_SplitItemUse Then
                '剔除固定费用的项
                For k = 0 To nCount - 1
                    If Val(atTemp(i).SplitItemID) = Val(m_aszFixFeeItem(k)) Then
                        Exit For
                    End If
                Next k
                '如果不是固定费用项,则显示
                If k > nCount - 1 Then
                    j = j + 1
                    VsProtocolItem.Rows = VsProtocolItem.Rows + 1
                    VsProtocolItem.TextMatrix(j, cnSplitItem) = MakeDisplayString(atTemp(i).SplitItemID, atTemp(i).SplitItemName)
                    VsProtocolItem.TextMatrix(j, cnFormulaName) = atTemp(i).FormularName
                    VsProtocolItem.TextMatrix(j, cnLimitCharge) = atTemp(i).LimitCharge
                    VsProtocolItem.TextMatrix(j, cnFormulaComment) = atTemp(i).FormulaComment
                    VsProtocolItem.TextMatrix(j, cnUpCharge) = atTemp(i).UpCharge
                End If
            End If
        Next i
        bIsNull = False  '存在收费项信息
        VsProtocolItem.Rows = VsProtocolItem.Rows - 1
    Else
        FillItem
        bIsNull = True '没有收费项信息
    End If
'    VsProtocolItem.Rows = VsProtocolItem.Rows - 1
    Exit Sub
err:
ShowErrorMsg
End Sub
Private Sub FillVsProtocolItem()
    On Error GoTo err
    Dim atTemp() As TSplitItemInfo
    
    Dim szTemp As String
    Dim i As Integer
    AlignFormPos Me
    szTemp = ""
    
    If ArrayLength(m_atFormula) <> 0 Then
        For i = 1 To ArrayLength(m_atFormula)
            szTemp = szTemp & MakeDisplayString(Trim(m_atFormula(i, 1)), Trim(m_atFormula(i, 2))) & "|"
        Next i
        szTemp = szTemp & " " & "|"
        VsProtocolItem.ColComboList(cnFormulaName) = szTemp
    End If
    
    Exit Sub
err:
ShowErrorMsg
End Sub
Private Sub Form_Unload(Cancel As Integer)
    SaveHeadWidth Me.name, VsProtocolItem
    SaveFormPos Me
End Sub

'
Private Sub VsProtocolItem_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    On Error GoTo err
    Dim aszTemp() As String
    Select Case Col
        Case cnSplitItem

        Case cnFormulaName
            If Trim(VsProtocolItem.TextMatrix(Row, Col)) = "" Then
                VsProtocolItem.TextMatrix(Row, cnFormulaComment) = ""
'                VsProtocolItem.ColComboList(cnFormulaComment) = ""
                Exit Sub
            End If
            aszTemp = m_oReport.GetAllFormula(ResolveDisplay(VsProtocolItem.TextMatrix(Row, Col)))
            If ArrayLength(aszTemp) <> 0 Then
                VsProtocolItem.Col = cnFormulaComment
                VsProtocolItem.Text = aszTemp(1, 3)
                VsProtocolItem.Col = cnFormulaName
            End If

    End Select
    Exit Sub
err:
ShowErrorMsg
End Sub


'填充VS的列头
Private Sub FillVSHead()
    With VsProtocolItem
        .Clear
        .Rows = 2
        .Cols = 6

        .TextMatrix(0, cnSplitItem) = "费用项"
        .TextMatrix(0, cnFormulaName) = "公式名称"
        .TextMatrix(0, cnLimitCharge) = "底限"
        .TextMatrix(0, cnFormulaComment) = "公式描述"
        .TextMatrix(0, cnUpCharge) = "上限"
    
    End With
End Sub

