VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmRealNameReg 
   BackColor       =   &H00FF80FF&
   BorderStyle     =   0  'None
   Caption         =   "ʵ���ƵǼ�"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1200
      Left            =   8400
      Top             =   30
   End
   Begin VB.ComboBox cboCardType 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1650
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   1695
   End
   Begin VSFlex7LCtl.VSFlexGrid vsCardInfo 
      Height          =   2505
      Left            =   120
      TabIndex        =   0
      Top             =   570
      Width           =   8580
      _cx             =   15134
      _cy             =   4419
      _ConvInfo       =   -1
      Appearance      =   1
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
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
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
   Begin MSComctlLib.Toolbar tbAddDel 
      Height          =   360
      Left            =   3540
      TabIndex        =   3
      Top             =   165
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ilsToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "AddOneLine"
            Object.ToolTipText     =   "����һ��"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DelOneLine"
            Object.ToolTipText     =   "ɾ��һ��"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ilsToolBar 
      Left            =   3690
      Top             =   30
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
            Picture         =   "frmRealNameReg.frx":0000
            Key             =   "add"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRealNameReg.frx":015A
            Key             =   "del"
         EndProperty
      EndProperty
   End
   Begin RTComctl3.CoolButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   5580
      TabIndex        =   6
      Top             =   3270
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "ȷ��(&O)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   11.25
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
      MICON           =   "frmRealNameReg.frx":04AD
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
      Height          =   375
      Left            =   7290
      TabIndex        =   7
      Top             =   3270
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "�ر�(&C)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   11.25
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
      MICON           =   "frmRealNameReg.frx":04C9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblIDCardNoTemp 
      Height          =   135
      Left            =   5220
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   30
      TabIndex        =   4
      Top             =   3300
      Width           =   5535
   End
   Begin VB.Label lblPrevDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "֤������(&T):"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   1260
   End
End
Attribute VB_Name = "frmRealNameReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' *******************************************************************
' *  Source File Name  :                                            *
' *  Project Name: StationNet V6                                    *
' *  Engineer: ������                                               *
' *  Date Generated: 2015/11/20                                     *
' *  Last Revision Date :                                           *
' *  Brief Description  : ʵ���ƵǼ�                                *
' *******************************************************************

Option Explicit

Public m_bOk As Boolean
Public m_nSellCount As Integer          '��Ʊ��
Private m_aszCardInfo() As TCardInfo    'ʵ����Ϣ
Public m_vaCardInfo As Variant     'ʵ����Ϣ

Const cnCardType = 0        '֤������
Const cnIDCardNo = 1        '֤������
Const cnPersonName = 2      '����
Const cnSex = 3             '�Ա�
Const cnNation = 4          '����(�й�)/����(���)
Const cnAddress = 5         '��ַ
Const cnPersonPicture = 6   '֤����Ƭ

Const cnTotalCols = 7       '����

Private Sub cboCardType_Click()
On Error GoTo err

    If cboCardType.Text = "���֤" Then
        Timer1.Enabled = True
    Else
        Timer1.Enabled = False
    End If

    Exit Sub
err:
    WriteErrorLog "cboCardType_Click", "ERROR:" & err.Source & "-->[" & err.Number & "]" & err.Description
End Sub

Private Sub cmdClose_Click()
On Error GoTo here
    Dim nCount As Integer

    nCount = vsCardInfo.Rows

    If MsgBox("ʵ������Ϣ�Ǽǻ�δ��ɣ�ȷ��Ҫ�˳��Ǽ���", vbInformation + vbYesNo + vbDefaultButton2, "ע��") = vbNo Then Exit Sub

    m_bOk = False
    Unload Me

    Exit Sub
here:
    WriteErrorLog "cmdClose_Click", "ERROR:" & err.Source & "-->[" & err.Number & "]" & err.Description
    ShowErrorMsg
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo here
    Dim i As Integer
    Dim nCount As Integer
    Dim szCheckIDReturn As String

    nCount = vsCardInfo.Rows
    If nCount <= 1 Then Exit Sub

    If nCount - 1 <> m_nSellCount Then
        MsgBox "֤����[" & nCount - 1 & "]������Ʊ��[" & m_nSellCount & "]�Ų�����", vbExclamation, App.Title
        Exit Sub
    End If

    ReDim m_aszCardInfo(1 To m_nSellCount)

    With vsCardInfo
        For i = 1 To nCount - 1
            szCheckIDReturn = ""
            If Trim(.TextMatrix(i, cnCardType)) = "���֤" Then
                If Trim(.TextMatrix(i, cnIDCardNo)) = "" And Trim(.TextMatrix(i, cnPersonName)) = "" Then
                    MsgBox "���֤������֤���Ų���Ϊ�գ�", vbExclamation, App.Title
                    Exit Sub
                End If
                If Trim(.TextMatrix(i, cnIDCardNo)) <> "" Then
                    szCheckIDReturn = CIDCheck(Trim(.TextMatrix(i, cnIDCardNo)))
                    If szCheckIDReturn <> "" Then
                        MsgBox "���֤[" & Trim(.TextMatrix(i, cnIDCardNo)) & "]������Ҫ��" & vbCrLf & szCheckIDReturn, vbExclamation, App.Title
                        Exit Sub
                    End If
                End If
            ElseIf Trim(.TextMatrix(i, cnIDCardNo)) <> "���֤" And Trim(.TextMatrix(i, cnIDCardNo)) = "" Then
                MsgBox "֤���Ų���Ϊ�գ�", vbExclamation, App.Title
                Exit Sub
            End If
            m_aszCardInfo(i).szCardType = Trim(.TextMatrix(i, cnCardType))
            m_aszCardInfo(i).szIDCardNo = Trim(.TextMatrix(i, cnIDCardNo))
            m_aszCardInfo(i).szPersonName = Trim(.TextMatrix(i, cnPersonName))
            m_aszCardInfo(i).szSex = Trim(.TextMatrix(i, cnSex))
            m_aszCardInfo(i).szNation = Trim(.TextMatrix(i, cnNation))
            m_aszCardInfo(i).szAddress = Trim(.TextMatrix(i, cnAddress))
            m_aszCardInfo(i).szPersonPicture = Trim(.TextMatrix(i, cnPersonPicture))
        Next i
    End With

    m_vaCardInfo = m_aszCardInfo

    m_bOk = True

    Erase m_aszCardInfo
    Timer1.Enabled = False
    Unload Me

    Exit Sub
here:
    WriteErrorLog "cmdOK_Click", "ERROR:" & err.Source & "-->[" & err.Number & "]" & err.Description
    ShowErrorMsg
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        cmdClose_Click
    End If
End Sub

Private Sub Form_Load()
On Error GoTo here
    Timer1.Enabled = False
    m_bOk = False

    FillCardType
    InitVsLst

    Dim szReturnMsg As String
    szReturnMsg = SetCardPortVision
    If szReturnMsg = cszSuccess Then
        ReadCard True
    Else
        lblMsg.Caption = szReturnMsg
    End If
    Timer1.Enabled = True

    Exit Sub
here:
    WriteErrorLog "Form_Load", "ERROR:" & err.Source & "-->[" & err.Number & "]" & err.Description
    Timer1.Enabled = False
    ShowErrorMsg
End Sub

Private Sub ReadCard(Optional bFormLoadRead As Boolean = False)
'bFormLoadRead �Ƿ��Ǵ������ʱ����
On Error GoTo ErrHandle
    Dim i As Integer
    Dim nCount As Integer
    Dim szReturnMsg As String
    Dim bFind As Boolean

    Dim szCardID As String
    Dim szPersonName As String
    Dim szSex As String
    Dim szAddress As String
    Dim szNation As String

    szReturnMsg = ""
    bFind = False
    szCardID = ""
    szPersonName = ""
    szSex = ""
    szAddress = ""
    szNation = ""
    lblIDCardNoTemp.Caption = ""

    nCount = vsCardInfo.Rows
    If nCount <= 0 Then Exit Sub

    szReturnMsg = GetReadCardVision
    If bFormLoadRead = True And szReturnMsg = cszReadCartFail Then Exit Sub

    If szReturnMsg <> cszSuccess Then
        lblMsg.Caption = szReturnMsg & "�������ԣ�"
        Exit Sub
    Else
        lblMsg.Caption = "�����ɹ��������һ�ţ�"
    End If

    szCardID = GetCardID
    szPersonName = GetName
    szSex = GetSex
    szNation = GetNation
    szAddress = GetAddress

    '��ʱ�洢,��Ϊ����������Ϣ������������ţ�ת����û����
    lblIDCardNoTemp.Caption = szCardID

    '�����б����Ƿ��Ѵ��ڴ�֤�������Ѵ��ڣ����˳����粻���ڣ�������һ��
    For i = 1 To nCount - 1
        If Trim(vsCardInfo.TextMatrix(i, cnIDCardNo)) = lblIDCardNoTemp.Caption Then bFind = True: Exit Sub
    Next i

    '����һ��
    If bFind = False Then
        With vsCardInfo
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, cnCardType) = "���֤"
            .TextMatrix(.Rows - 1, cnIDCardNo) = szCardID
            .TextMatrix(.Rows - 1, cnPersonName) = szPersonName
            .TextMatrix(.Rows - 1, cnSex) = szSex
            .TextMatrix(.Rows - 1, cnNation) = szNation
            .TextMatrix(.Rows - 1, cnAddress) = szAddress
            .TextMatrix(.Rows - 1, cnPersonPicture) = ""
        End With
    End If

    Exit Sub
ErrHandle:
    WriteErrorLog "ReadCard", "ERROR:" & err.Source & "-->[" & err.Number & "]" & err.Description
    Timer1.Enabled = False
End Sub

'��ʼ��֤������
Private Sub FillCardType()
    cboCardType.Clear
    cboCardType.AddItem "���֤"
    cboCardType.AddItem "����"
    cboCardType.AddItem "�۰�ͨ��֤"
    cboCardType.AddItem "̨��ͨ��֤"
    cboCardType.AddItem "����"
    cboCardType.ListIndex = 0
End Sub

'��ʼ���б�ؼ�
Private Sub InitVsLst()
    With vsCardInfo
        .Rows = 1
        .RowHeight(0) = 350
        .RowHeightMin = 300
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightNever
        .ScrollBars = flexScrollBarBoth
        .AllowUserResizing = flexResizeColumns
        .FontSize = 10.5
        .ShowComboButton = True
        .BackColorBkg = RGB(255, 255, 255)
        .SheetBorder = RGB(255, 255, 255)
        .GridColorFixed = RGB(163, 208, 217)
        .BackColorFixed = RGB(212, 221, 226)
        .GridColor = RGB(163, 208, 217)
        .BackColorAlternate = RGB(245, 245, 245)
        
        .Cols = cnTotalCols
        .Cell(flexcpText, 0, cnCardType) = "֤������"
        .Cell(flexcpText, 0, cnIDCardNo) = "֤������"
        .Cell(flexcpText, 0, cnPersonName) = "����"
        .Cell(flexcpText, 0, cnSex) = "�Ա�"
        .Cell(flexcpText, 0, cnNation) = "����/����"
        .Cell(flexcpText, 0, cnAddress) = "��ַ"
        .Cell(flexcpText, 0, cnPersonPicture) = "֤����Ƭ"
        
        .ColWidth(cnCardType) = 1200
        .ColWidth(cnIDCardNo) = 2200
        .ColWidth(cnPersonName) = 1200
        .ColWidth(cnSex) = 600
        .ColWidth(cnNation) = 1020
        .ColWidth(cnAddress) = 7000
        .ColHidden(cnPersonPicture) = True
    End With
End Sub

Private Sub tbAddDel_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrHandle
    Select Case Button.Key
        Case "AddOneLine"
            AddOneLine
        Case "DelOneLine"
            DelOneLine
    End Select
    Exit Sub
ErrHandle:
    WriteErrorLog "tbAddDel_ButtonClick", "ERROR:" & err.Source & "-->[" & err.Number & "]" & err.Description
    ShowErrorMsg
End Sub

'����һ��
Private Sub AddOneLine()
On Error GoTo ErrHandle

    With vsCardInfo
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, cnCardType) = cboCardType.Text
        .TextMatrix(.Rows - 1, cnIDCardNo) = ""
        .TextMatrix(.Rows - 1, cnPersonName) = ""
        .TextMatrix(.Rows - 1, cnSex) = "��"
        .TextMatrix(.Rows - 1, cnNation) = "��"
        .TextMatrix(.Rows - 1, cnAddress) = ""
        .TextMatrix(.Rows - 1, cnPersonPicture) = ""
    End With
    
    Exit Sub
ErrHandle:
    WriteErrorLog "AddOneLine", "ERROR:" & err.Source & "-->[" & err.Number & "]" & err.Description
    ShowErrorMsg
End Sub

'ɾ��һ��
Private Sub DelOneLine()
On Error GoTo ErrHandle
    If vsCardInfo.Rows <= 1 Then Exit Sub
    vsCardInfo.Rows = vsCardInfo.Rows - 1
    
    Exit Sub
ErrHandle:
    WriteErrorLog "DelOneLine", "ERROR:" & err.Source & "-->[" & err.Number & "]" & err.Description
    ShowErrorMsg
End Sub

Private Sub Timer1_Timer()
On Error GoTo ErrHandle

    Timer1.Enabled = False
    ReadCard
    Timer1.Enabled = True

    Exit Sub
ErrHandle:
    WriteErrorLog "Timer1_Timer", "ERROR:" & err.Source & "-->[" & err.Number & "]" & err.Description
    Timer1.Enabled = False
    ShowErrorMsg
End Sub

Private Sub vsCardInfo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrHandle
    If vsCardInfo.Rows <= 1 Then Exit Sub
    Select Case Col
        Case cnSex
          vsCardInfo.ColComboList(Col) = "��|Ů"
    End Select

    Exit Sub
ErrHandle:
    WriteErrorLog "vsCardInfo_BeforeEdit", "ERROR:" & err.Source & "-->[" & err.Number & "]" & err.Description
End Sub

Private Sub vsCardInfo_EnterCell()
On Error GoTo err

    If vsCardInfo.Row <= 0 Or vsCardInfo.Col <= 0 Then Exit Sub
    '������Щ��Ԫ����Ա༭����Щֻ�ܲ鿴
    Select Case vsCardInfo.Col
        Case cnCardType   'ֻ�������ɱ༭
            vsCardInfo.Editable = flexEDNone
        Case Else '�ɱ༭
            vsCardInfo.Editable = flexEDKbdMouse
    End Select
    Exit Sub
err:
    WriteErrorLog "vsCardInfo_EnterCell", "ERROR:" & err.Source & "-->[" & err.Number & "]" & err.Description
End Sub
