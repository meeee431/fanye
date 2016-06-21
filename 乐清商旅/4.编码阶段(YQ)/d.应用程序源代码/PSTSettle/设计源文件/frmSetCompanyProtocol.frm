VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmSetCompanyProtocol 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "公司协议设置"
   ClientHeight    =   4470
   ClientLeft      =   5160
   ClientTop       =   3675
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -780
      TabIndex        =   6
      Top             =   780
      Width           =   7815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   915
      Left            =   -30
      TabIndex        =   2
      Top             =   3780
      Width           =   6015
      Begin RTComctl3.CoolButton cmdOk 
         Height          =   345
         Left            =   3120
         TabIndex        =   3
         Top             =   270
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "确定(&E)"
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
         MICON           =   "frmSetCompanyProtocol.frx":0000
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
         Height          =   345
         Left            =   240
         TabIndex        =   4
         Top             =   270
         Width           =   945
         _ExtentX        =   1667
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
         MICON           =   "frmSetCompanyProtocol.frx":001C
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
         Left            =   4500
         TabIndex        =   5
         Top             =   270
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "取消(&C)"
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
         MICON           =   "frmSetCompanyProtocol.frx":0038
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
   Begin VSFlex7LCtl.VSFlexGrid vsCompanyRoute 
      Height          =   2265
      Left            =   60
      TabIndex        =   1
      Top             =   1380
      Width           =   5595
      _cx             =   9869
      _cy             =   3995
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSetCompanyProtocol.frx":0054
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
   Begin FText.asFlatTextBox txtCompany 
      Height          =   285
      Left            =   1290
      TabIndex        =   7
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
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
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "公司代码："
      Height          =   225
      Left            =   180
      TabIndex        =   8
      Top             =   990
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "公司协议设置"
      Height          =   255
      Left            =   300
      TabIndex        =   0
      Top             =   240
      Width           =   1275
   End
End
Attribute VB_Name = "frmSetCompanyProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private aszCompanyID() As String
Public m_oReport As New Report
Public m_oSplit As New Split


Public m_eFormStatus As EFormStatus
Public m_szCompanyID As String
Public m_bIsBack As Boolean

Private Sub cmdok_Click()
    Dim i As Integer
    Dim aszTemp() As String
    Dim szRouteID As String
    Dim szRouteName As String
    
    ReDim aszTemp(1 To vsCompanyRoute.Rows)
    
    For i = 1 To vsCompanyRoute.Rows - 1
        szRouteID = ResolveDisplay(vsCompanyRoute.TextMatrix(i, 2), szRouteName)
        m_oSplit.SetCompanyProtocol ResolveDisplay(vsCompanyRoute.TextMatrix(i, 1)), szRouteID, ResolveDisplay(vsCompanyRoute.TextMatrix(i, 3)), szRouteName
        frmCompanyProtocol.FilllvCompany ResolveDisplay(vsCompanyRoute.TextMatrix(i, 1)), szRouteID
    Next i
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    cmdOk.Enabled = False
    m_oReport.Init g_oActiveUser
    m_oSplit.Init g_oActiveUser
    FillHead
    AlignHeadWidth Me.Caption, vsCompanyRoute
    
    vsCompanyRoute.Cols = 4
    vsCompanyRoute.Rows = 2
    
    If m_eFormStatus = AddStatus Then
        
    ElseIf m_eFormStatus = ModifyStatus Then
        '刷新该公司的协议信息
        txtCompany.Text = m_szCompanyID
        ReDim aszCompanyID(1 To 1)
        aszCompanyID(1) = m_szCompanyID
        FillVSCompanyRoute aszCompanyID
    End If
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveHeadWidth Me.Caption, vsCompanyRoute
End Sub

Private Sub txtCompany_ButtonClick()
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectCompany
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtCompany.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))
    ReDim aszCompanyID(1 To 1)
    aszCompanyID(1) = txtCompany.Text
    FillVSCompanyRoute aszCompanyID
End Sub

Public Sub GetCompanyID(CompnayID() As String)
    aszCompanyID = CompnayID
End Sub


Private Sub FillVSCompanyRoute(CompanyID() As String)
    Dim aszTemp() As String
    Dim i As Integer
    Dim oBackBaseInfo As New BackBaseInfo
'    vsCompanyRoute.MergeCol(1) = True
'    vsCompanyRoute.MergeCells = flexMergeRestrictColumns
'    vsCompanyRoute.Clear
    vsCompanyRoute.Rows = 2
    If ArrayLength(CompanyID) > 0 Then
        If m_bIsBack And m_eFormStatus = AddStatus Then
            oBackBaseInfo.Init g_oActiveUser
            aszTemp = oBackBaseInfo.GetRoute
        Else
            aszTemp = m_oReport.GetCopmanyRoute(CompanyID(1))
        End If
        If ArrayLength(aszTemp) > 0 Then
            vsCompanyRoute.Rows = ArrayLength(aszTemp) + 1
            For i = 1 To ArrayLength(aszTemp)
                vsCompanyRoute.Cell(flexcpText, i, 1) = CompanyID(1)
                vsCompanyRoute.TextMatrix(i, 2) = MakeDisplayString(aszTemp(i, 1), aszTemp(i, 2))
            Next i
            cmdOk.Enabled = True
        Else
            cmdOk.Enabled = False
        End If
           
    End If
    
    '填充已有的协议
    FillCompanyProtocol
    
    
End Sub



Private Sub FillHead()
    Dim aszTemp() As String, i As Integer, aszTempF() As String
    Dim szTemp As String
    vsCompanyRoute.TextMatrix(0, 1) = "公司"
    vsCompanyRoute.TextMatrix(0, 2) = "线路"
    vsCompanyRoute.TextMatrix(0, 3) = "协议"
    aszTemp = m_oReport.GetAllProtocol
    
    If ArrayLength(aszTemp) > 0 Then
        
        For i = 1 To ArrayLength(aszTemp)
            szTemp = szTemp & MakeDisplayString(aszTemp(i, 1), aszTemp(i, 2)) & "|"
        Next i
        vsCompanyRoute.ColComboList(3) = szTemp
    End If
End Sub

Private Sub FillCompanyProtocol()
    On Error GoTo err
    Dim nCount As Integer
    Dim aszTemp() As String
    Dim i As Integer
    Dim lvTemp As ListItem
    Dim j As Integer
    
    aszTemp = m_oReport.GetAllCompanyProtocol(ResolveDisplay(txtCompany.Text))
    nCount = ArrayLength(aszTemp)
    For i = 1 To nCount
        With vsCompanyRoute
            For j = 1 To .Rows - 1
                If aszTemp(i, 1) = ResolveDisplay(.TextMatrix(j, 1)) And aszTemp(i, 5) = ResolveDisplay(.TextMatrix(j, 2)) Then
                    .TextMatrix(j, 3) = MakeDisplayString(aszTemp(i, 3), aszTemp(i, 4))
                    Exit For
                End If
            Next j
            '未找到的话,新增一行,加入此设置
            If j = .Rows Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = MakeDisplayString(aszTemp(i, 1), aszTemp(i, 2))
                .TextMatrix(.Rows - 1, 2) = MakeDisplayString(aszTemp(i, 5), aszTemp(i, 6))
                .TextMatrix(.Rows - 1, 3) = MakeDisplayString(aszTemp(i, 3), aszTemp(i, 4))
            End If
        End With
    Next i
    Exit Sub
err:
    ShowErrorMsg
End Sub






Private Sub FillNewCompanyProtocol()
    On Error GoTo err
    Dim nCount As Integer
    Dim aszTemp() As String
    Dim i As Integer
    Dim lvTemp As ListItem
    Dim j As Integer
    
    aszTemp = m_oReport.GetAllCompanyProtocol(ResolveDisplay(txtCompany.Text))
    nCount = ArrayLength(aszTemp)
    With vsCompanyRoute
        .MergeCells = flexMergeRestrictColumns
        
        .Rows = nCount + 1
        For i = 1 To nCount
            .TextMatrix(i, 1) = MakeDisplayString(aszTemp(i, 1), aszTemp(i, 2))
            .TextMatrix(i, 2) = MakeDisplayString(aszTemp(i, 5), aszTemp(i, 6))
            .TextMatrix(i, 3) = MakeDisplayString(aszTemp(i, 3), aszTemp(i, 4))
'            .TextMatrix(i, 4) = MakeDisplayString(aszTemp(i, 7), GetCheckStr(aszTemp(i, 7)))
        Next i
        .MergeCol(1) = True
    End With
    Exit Sub
err:
    ShowErrorMsg
End Sub






