VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmMantissa 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "β��������"
   ClientHeight    =   4875
   ClientLeft      =   2745
   ClientTop       =   3225
   ClientWidth     =   7020
   HelpContextID   =   10000550
   Icon            =   "frmMantissa.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7020
   StartUpPosition =   1  '����������
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   330
      Left            =   5670
      TabIndex        =   8
      Top             =   2610
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "����"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "frmMantissa.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdDelete 
      Height          =   330
      Left            =   5640
      TabIndex        =   7
      Top             =   1155
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "ɾ��(&D)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "frmMantissa.frx":0166
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
      Height          =   330
      Left            =   5640
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "�ر�(&C)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "frmMantissa.frx":0182
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdAdd 
      Height          =   330
      Left            =   5640
      TabIndex        =   5
      Top             =   345
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "����(A)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "frmMantissa.frx":019E
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
      Height          =   330
      Left            =   5640
      TabIndex        =   4
      Top             =   750
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "����(&S)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "frmMantissa.frx":01BA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid vsTail 
      Height          =   1365
      Left            =   75
      TabIndex        =   1
      Top             =   360
      Width           =   5430
      _cx             =   9578
      _cy             =   2408
      _ConvInfo       =   -1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Rows            =   2
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
   Begin VSFlex7LCtl.VSFlexGrid vsAllTail 
      Height          =   2745
      Left            =   75
      TabIndex        =   3
      Top             =   2040
      Width           =   5430
      _cx             =   9578
      _cy             =   4842
      _ConvInfo       =   -1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ǰ�༭β��������(&I):"
      Height          =   180
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   2160
   End
   Begin VB.Label lblAllDealMethod 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����β��������(&L):"
      Height          =   180
      Left            =   75
      TabIndex        =   2
      Top             =   1800
      Width           =   1800
   End
End
Attribute VB_Name = "frmMantissa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'* Source File Name:frmMantissa.frm
'* Project Name:RTBusMan
'* Engineer:�·�
'* Data Generated:2002/09/04
'* Last Revision Date:2002/09/04
'* Brief Description:
'* Relational Document:
'**********************************************************
Option Explicit
'vsAllTail���׶���
Const cnMethodNOAll = 1
Const cnSerialNOAll = 2
Const cnStartNumberAll = 3
Const cnEndNumberAll = 4
Const cnDealNumberAll = 5
Const cnAnnoAll = 6
'vsTail���׶���
Const cnMethodNO = 1
Const cnSerialNO = 2
Const cnStartNumber = 3
Const cnEndNumber = 4
Const cnDealNumber = 5
Const cnAnno = 6

Private Const cnTotal = "Ʊ����ֵ"
Private Const cnDetail = "Ʊ�۷���"
Private m_oTicketPriceMan As New TicketPriceMan

Private Sub cmdAdd_Click()
    Dim i As Integer
    Dim szTemp As String
    Dim nTemp As Integer
    On Error GoTo ErrorHandle
    For i = 1 To vsAllTail.Rows - 1
        If nTemp < Val(vsAllTail.TextMatrix(i, cnMethodNOAll)) Then nTemp = Val(vsAllTail.TextMatrix(i, cnMethodNOAll))
    Next
    nTemp = nTemp + 1
    szTemp = Trim(str(nTemp))
    For i = 1 To 3
        If Len(szTemp) < 3 Then szTemp = "0" & szTemp
    Next
    vsTail.Enabled = True
    vsTail.Rows = 2
    For i = 1 To vsTail.Rows - 1
        vsTail.TextMatrix(i, cnMethodNO) = szTemp
        vsTail.TextMatrix(i, cnSerialNO) = i
        vsTail.TextMatrix(i, cnAnno) = ""
    Next
    vsTail.TextMatrix(1, cnStartNumber) = "0"
    cmdAdd.Enabled = False
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim i, j As Integer
    Dim nFromRow As Integer
    Dim nSameRowCount As Integer
    Dim szTemp As String
    Dim tTTailDealMethod() As TTailDealMethod

    On Error GoTo ErrorHandle
    
    If vsAllTail.Row > 0 Then szTemp = vsAllTail.TextMatrix(vsAllTail.Row, cnMethodNOAll)
    For i = 1 To vsAllTail.Rows - 1
        If szTemp = vsAllTail.TextMatrix(i, cnMethodNOAll) Then
           If nSameRowCount = 0 Then nFromRow = i
           nSameRowCount = nSameRowCount + 1
        End If
    Next i
    ReDim tTTailDealMethod(1 To nSameRowCount)
    For i = nFromRow To nFromRow + nSameRowCount - 1
        tTTailDealMethod(i - nFromRow + 1).szMethodNo = Trim(vsAllTail.TextMatrix(i, cnMethodNOAll))
        tTTailDealMethod(i - nFromRow + 1).szSerialNo = Trim(vsAllTail.TextMatrix(i, cnSerialNOAll))
        tTTailDealMethod(i - nFromRow + 1).nStartNumber = vsAllTail.TextMatrix(i, cnStartNumberAll)
        tTTailDealMethod(i - nFromRow + 1).nEndNumber = vsAllTail.TextMatrix(i, cnEndNumberAll)
        If Trim(vsAllTail.TextMatrix(i, cnDealNumberAll)) = "ԭֵ" Then
           tTTailDealMethod(i - nFromRow + 1).nDealedNumber = "100"
        Else
           tTTailDealMethod(i - nFromRow + 1).nDealedNumber = vsAllTail.TextMatrix(i, cnDealNumberAll)
        End If
        tTTailDealMethod(i - nFromRow + 1).szAnnotation = Trim(vsAllTail.TextMatrix(i, cnAnnoAll))
    Next i
    If m_oTicketPriceMan.DeleteTailDealMethod(tTTailDealMethod) = False Then Exit Sub
    For i = nFromRow To vsAllTail.Rows - 1 - nSameRowCount
        For j = 1 To vsAllTail.Cols - 1
            vsAllTail.TextMatrix(i, j) = vsAllTail.TextMatrix(i + nSameRowCount, j)
        Next j
    Next i
    vsAllTail.Rows = vsAllTail.Rows - nSameRowCount
    vsTail.Rows = 2
    Exit Sub

ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub cmdSave_Click()
Dim i As Integer, j As Integer, n As Integer
Dim oldRows As Integer
Dim newI As Integer
Dim nSameRowCount As Integer
Dim nFromRow As Integer
Dim tTTailDealMethod() As TTailDealMethod

On Error GoTo ErrorHandle
    If vsTail.TextMatrix(1, cnMethodNO) = "" Then Exit Sub
    If Trim(vsTail.TextMatrix(vsTail.Rows - 1, cnEndNumber)) <> "" Then
       If vsTail.TextMatrix(vsTail.Rows - 1, cnEndNumber) < 9 Then
           Exit Sub
       End If
    Else
       vsTail.TextMatrix(vsTail.Rows - 1, cnEndNumber) = 9
    End If
    For i = 1 To vsTail.Rows - 1
        If Trim(vsTail.TextMatrix(i, cnDealNumber)) = "" Then vsTail.TextMatrix(i, cnDealNumber) = "ԭֵ"
    Next
    For i = 1 To vsAllTail.Rows - 1
        If vsTail.TextMatrix(1, cnMethodNO) = vsAllTail.TextMatrix(i, cnMethodNOAll) Then
           If nSameRowCount = 0 Then nFromRow = i
           nSameRowCount = nSameRowCount + 1
        End If
    Next
    For i = nFromRow + nSameRowCount To vsAllTail.Rows - 1
        For j = 1 To vsAllTail.Cols - 1
            vsAllTail.TextMatrix(i - nSameRowCount, j) = vsAllTail.TextMatrix(i, j)
        Next
    Next

    ReDim tTTailDealMethod(1 To vsTail.Rows - 1)
    For i = 1 To vsTail.Rows - 1
        tTTailDealMethod(i).szMethodNo = Trim(vsTail.TextMatrix(i, cnMethodNO))
        tTTailDealMethod(i).szSerialNo = Trim(vsTail.TextMatrix(i, cnSerialNO))
        tTTailDealMethod(i).nStartNumber = vsTail.TextMatrix(i, cnStartNumber)
        tTTailDealMethod(i).nEndNumber = vsTail.TextMatrix(i, cnEndNumber)
        If Trim(vsTail.TextMatrix(i, cnDealNumber)) = "ԭֵ" Then
           tTTailDealMethod(i).nDealedNumber = "100"
        Else
           tTTailDealMethod(i).nDealedNumber = vsTail.TextMatrix(i, cnDealNumber)
        End If
        tTTailDealMethod(i).szAnnotation = Trim(vsTail.TextMatrix(i, cnAnno))
    Next

    m_oTicketPriceMan.SaveTailDealMethod tTTailDealMethod
    vsAllTail.Rows = vsAllTail.Rows - nSameRowCount
    oldRows = vsAllTail.Rows
    vsAllTail.Rows = vsAllTail.Rows + vsTail.Rows - 1
    newI = 1
    For n = oldRows To vsAllTail.Rows - 1
        For j = 1 To vsTail.Cols - 1
            vsAllTail.TextMatrix(n, j) = vsTail.TextMatrix(newI, j)
        Next
        newI = newI + 1
    Next
    cmdAdd.Enabled = True
'    vsTail.Rows = 1
    vsTail.Rows = 2
    MsgBox "β�����������ӻ���±���ɹ���", vbOKOnly + vbInformation, Me.Caption
    Exit Sub

ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_Load()
Dim i As Integer
    m_oTicketPriceMan.Init g_oActiveUser
    m_oTicketPriceMan.ObjStatus = ST_NormalObj
'    vsTail.Cols = 6
'    vsAllTail.Cols = 6
    InitTailGrid
    InitAllTailGrid
    
    vsTail.Enabled = False
    cmdDelete.Enabled = False
    LoadAllTailDealMethod
End Sub

Private Sub vsAllTail_EnterCell()
    Dim szTemp As String
    Dim i, j As Integer
    Dim nFromRow As Integer
    Dim nSameRowCount As Integer
    Dim str() As String
    
    On Error GoTo ErrorHandle
    
    szTemp = vsAllTail.TextMatrix(vsAllTail.Row, cnMethodNOAll)
    For i = 1 To vsAllTail.Rows - 1
        If vsAllTail.TextMatrix(i, cnMethodNOAll) = szTemp Then
           If nSameRowCount = 0 Then nFromRow = i
           nSameRowCount = nSameRowCount + 1
        End If
    Next
    ReDim str(1 To nSameRowCount, 1 To vsAllTail.Cols - 1)
    For i = nFromRow To nFromRow + nSameRowCount - 1
        For j = 1 To vsAllTail.Cols - 1
            str(i - nFromRow + 1, j) = vsAllTail.TextMatrix(i, j)
        Next
    Next
    vsTail.Rows = nSameRowCount + 1
    For i = 1 To vsTail.Rows - 1
        For j = 1 To vsTail.Cols - 1
            vsTail.TextMatrix(i, j) = str(i, j)
        Next
    Next
    cmdDelete.Enabled = True
'    cmdCancel.Enabled = False
    vsTail.Enabled = True
    cmdAdd.Enabled = True
    Exit Sub
    
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub vsTail_EnterCell()
On Error GoTo ErrorHandle
    If vsTail.Col = cnMethodNO Or vsTail.Col = cnSerialNO Or vsTail.Col = cnStartNumber Then
       vsTail.Editable = flexEDNone
    ElseIf vsTail.Col = vsTail.Cols - 1 And vsTail.Row <> 1 Then
       vsTail.Editable = flexEDNone
    Else
       vsTail.Editable = flexEDKbdMouse
    End If
    If vsTail.Col = vsTail.Cols - 1 Then vsTail.Row = 1
    If vsTail.Col = vsTail.Cols - 1 And vsTail.Row = 1 Then
        If Trim(vsTail.TextMatrix(1, vsTail.Cols - 1)) = "" Then
           vsTail.TextMatrix(1, vsTail.Cols - 1) = ""
        End If
    End If
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub vsTail_GotFocus()
    cmdDelete.Enabled = False
End Sub

Private Sub vsTail_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       If vsTail.Col < vsTail.Cols - 1 Then
          vsTail.Col = vsTail.Col + 1
       ElseIf vsTail.Col = vsTail.Cols - 1 And vsTail.Row = vsTail.Rows - 1 Then
          vsTail.Col = cnEndNumber
          vsTail.Row = vsTail.Row + 1
       End If
    End If
End Sub


Private Sub vsTail_LeaveCell()
Dim i As Integer
    
    On Error GoTo ErrorHandle
    
    If Trim(vsTail.TextMatrix(vsTail.Rows - 1, cnEndNumber)) <> "" Then
        If vsTail.TextMatrix(vsTail.Rows - 1, cnEndNumber) < 9 Then
           vsTail.Rows = vsTail.Rows + 1
           vsTail.TextMatrix(vsTail.Rows - 1, cnMethodNO) = vsTail.TextMatrix(vsTail.Rows - 2, cnMethodNO)
           vsTail.TextMatrix(vsTail.Rows - 1, cnSerialNO) = vsTail.TextMatrix(vsTail.Rows - 2, cnSerialNO) + 1
           vsTail.TextMatrix(vsTail.Rows - 1, cnAnno) = vsTail.TextMatrix(vsTail.Rows - 2, cnAnno)
        End If
    End If
    If vsTail.Col = cnEndNumber Then
       If vsTail.TextMatrix(vsTail.Row, cnEndNumber) <= vsTail.TextMatrix(vsTail.Row, cnStartNumber) Then vsTail.TextMatrix(vsTail.Row, cnEndNumber) = vsTail.TextMatrix(vsTail.Row, cnStartNumber)
    End If
    For i = vsTail.Row To vsTail.Rows - 1
        If Trim(vsTail.TextMatrix(i, cnEndNumber)) <> "" Then
           If vsTail.Row > 0 Then
              If vsTail.TextMatrix(vsTail.Row, cnEndNumber) <= vsTail.TextMatrix(vsTail.Row, cnStartNumber) Then vsTail.TextMatrix(vsTail.Row, cnEndNumber) = vsTail.TextMatrix(vsTail.Row, cnStartNumber)
           End If
        End If
    Next
    If vsTail.Row <> 0 Then
        If Trim(vsTail.TextMatrix(vsTail.Row, cnEndNumber)) <> "" Then
            If vsTail.TextMatrix(vsTail.Row, cnEndNumber) < 9 Then
               If vsTail.Row >= vsTail.Rows - 1 Then
                  vsTail.Rows = vsTail.Rows + 1
                  vsTail.TextMatrix(vsTail.Rows - 1, cnMethodNO) = vsTail.TextMatrix(vsTail.Rows - 2, cnMethodNO)
                  vsTail.TextMatrix(vsTail.Rows - 1, cnSerialNO) = vsTail.TextMatrix(vsTail.Rows - 2, cnSerialNO) + 1
               End If
               vsTail.TextMatrix(vsTail.Row + 1, cnStartNumber) = vsTail.TextMatrix(vsTail.Row, cnEndNumber) + 1
            End If
        End If
    End If
    For i = 1 To vsTail.Rows - 1
        If Trim(vsTail.TextMatrix(i, cnEndNumber)) <> "" Then
           If vsTail.TextMatrix(i, cnStartNumber) = 9 Then vsTail.TextMatrix(i, cnEndNumber) = 9
           If vsTail.TextMatrix(i, cnEndNumber) = 9 Then vsTail.Rows = vsTail.Rows - (vsTail.Rows - (i + 1))
        End If
    Next
    
    If vsTail.Col = vsTail.Cols - 1 And vsTail.Row = 1 Then
       If Trim(vsTail.TextMatrix(1, vsTail.Cols - 1)) = "" Then
          vsTail.TextMatrix(1, vsTail.Cols - 1) = " "
       End If
        For i = 2 To vsTail.Rows - 1
            vsTail.TextMatrix(i, vsTail.Cols - 1) = vsTail.TextMatrix(1, vsTail.Cols - 1)
        Next
    End If
    
    vsTail.MergeCells = flexMergeFree
    
    Exit Sub
    
ErrorHandle:
End Sub

Private Sub LoadAllTailDealMethod()
Dim szTemp As String
Dim tTTailDealMethod() As TTailDealMethod
Dim nCount As Integer
Dim i As Integer

On Error GoTo ErrorHandle

   tTTailDealMethod = m_oTicketPriceMan.GetAllTailDealMethod()
   nCount = ArrayLength(tTTailDealMethod)
   vsAllTail.Rows = nCount + 1
   For i = 1 To nCount

       vsAllTail.TextMatrix(i, cnMethodNOAll) = tTTailDealMethod(i).szMethodNo
       vsAllTail.TextMatrix(i, cnSerialNOAll) = tTTailDealMethod(i).szSerialNo
       vsAllTail.TextMatrix(i, cnStartNumberAll) = tTTailDealMethod(i).nStartNumber
       vsAllTail.TextMatrix(i, cnEndNumberAll) = tTTailDealMethod(i).nEndNumber
       If tTTailDealMethod(i).nDealedNumber <> 100 Then
          vsAllTail.TextMatrix(i, cnDealNumberAll) = tTTailDealMethod(i).nDealedNumber
       Else
          vsAllTail.TextMatrix(i, cnDealNumberAll) = "ԭֵ"
       End If
       vsAllTail.TextMatrix(i, cnAnnoAll) = tTTailDealMethod(i).szAnnotation
   Next
   Exit Sub

ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub InitTailGrid()
    '�����п�
    vsTail.ColWidth(0) = 100
    vsTail.ColWidth(cnMethodNO) = 500
    vsTail.ColWidth(cnSerialNO) = 500
    vsTail.ColWidth(cnStartNumber) = 1000
    vsTail.ColWidth(cnEndNumber) = 1000
    vsTail.ColWidth(cnDealNumber) = 1000
    vsTail.ColWidth(cnAnno) = 1000
    '���ö���
    vsTail.FixedAlignment(cnMethodNO) = flexAlignCenterCenter
    vsTail.FixedAlignment(cnSerialNO) = flexAlignCenterCenter
    vsTail.FixedAlignment(cnStartNumber) = flexAlignCenterCenter
    vsTail.FixedAlignment(cnEndNumber) = flexAlignCenterCenter
    vsTail.FixedAlignment(cnDealNumber) = flexAlignCenterCenter
    vsTail.FixedAlignment(cnAnno) = flexAlignCenterCenter
    '���úϲ�
    vsTail.MergeCol(cnMethodNO) = True
    vsTail.MergeCol(cnAnno) = True
    '�����ı�
    vsTail.TextMatrix(0, cnMethodNO) = "���"
    vsTail.TextMatrix(0, cnSerialNO) = "���"
    vsTail.TextMatrix(0, cnStartNumber) = "��ʼ����"
    vsTail.TextMatrix(0, cnEndNumber) = "��������"
    vsTail.TextMatrix(0, cnDealNumber) = "����ֵ"
    vsTail.TextMatrix(0, cnAnno) = "��ע"
    '������Ͽ������
    vsTail.ColComboList(cnDealNumber) = "0|1|2|3|4|5|6|7|8|9|10|ԭֵ"
    vsTail.ColComboList(cnEndNumber) = "0|1|2|3|4|5|6|7|8|9"
    
    vsTail.MergeCells = flexMergeFree
End Sub

Private Sub InitAllTailGrid()
    '��������
    vsAllTail.ColSort(cnMethodNOAll) = flexSortStringAscending
    '���ÿ��
    vsAllTail.ColWidth(0) = 100
    vsAllTail.ColWidth(cnMethodNOAll) = 500
    vsAllTail.ColWidth(cnSerialNOAll) = 500
    vsAllTail.ColWidth(cnStartNumberAll) = 1000
    vsAllTail.ColWidth(cnEndNumberAll) = 1000
    vsAllTail.ColWidth(cnDealNumberAll) = 1000
    vsAllTail.ColWidth(cnAnnoAll) = 1000
    '�����ı�
    vsAllTail.TextMatrix(0, cnMethodNOAll) = "���"
    vsAllTail.TextMatrix(0, cnSerialNOAll) = "���"
    vsAllTail.TextMatrix(0, cnStartNumberAll) = "��ʼ����"
    vsAllTail.TextMatrix(0, cnEndNumberAll) = "��������"
    vsAllTail.TextMatrix(0, cnDealNumberAll) = "����ֵ"
    vsAllTail.TextMatrix(0, cnAnnoAll) = "��ע"
    '���ö���
    vsAllTail.FixedAlignment(cnMethodNOAll) = flexAlignCenterCenter
    vsAllTail.FixedAlignment(cnSerialNOAll) = flexAlignCenterCenter
    vsAllTail.FixedAlignment(cnStartNumberAll) = flexAlignCenterCenter
    vsAllTail.FixedAlignment(cnEndNumberAll) = flexAlignCenterCenter
    vsAllTail.FixedAlignment(cnDealNumberAll) = flexAlignCenterCenter
    vsAllTail.FixedAlignment(cnAnnoAll) = flexAlignCenterCenter
    '���úϲ�
    vsAllTail.MergeCol(cnMethodNOAll) = True
    vsAllTail.MergeCol(cnAnnoAll) = True
    vsAllTail.MergeCells = flexMergeFree
End Sub

