VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmSellTicketQuery 
   Caption         =   "��Ʊ��ʱ��ѯ"
   ClientHeight    =   4185
   ClientLeft      =   2610
   ClientTop       =   2550
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4185
   ScaleWidth      =   7545
   Begin MSComDlg.CommonDialog SaveDialogue 
      Left            =   4650
      Top             =   1020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboQueryType 
      Height          =   300
      Left            =   1470
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   600
      Width           =   2115
   End
   Begin FText.asFlatTextBox tbbTitileID 
      Height          =   285
      Left            =   5220
      TabIndex        =   10
      Top             =   600
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpStartTime 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy-M-d"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   1470
      TabIndex        =   8
      Top             =   240
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   23658499
      CurrentDate     =   37022
   End
   Begin RTComctl3.CoolButton cmdSave 
      TX         =   "����(&S)"
      Height          =   345
      Left            =   6210
      TabIndex        =   7
      Top             =   1710
      Width           =   1245
   End
   Begin RTComctl3.CoolButton cmdPrint 
      TX         =   "��ӡ(&P)"
      Height          =   345
      Left            =   6210
      TabIndex        =   6
      Top             =   2610
      Width           =   1245
   End
   Begin RTComctl3.CoolButton cmdQuery 
      TX         =   "��ѯ(&Q)"
      Default         =   -1  'True
      Height          =   345
      Left            =   6210
      TabIndex        =   5
      Top             =   1320
      Width           =   1245
   End
   Begin RTComctl3.CoolButton cmdPreView 
      TX         =   "Ԥ��(&V)"
      Height          =   345
      Left            =   6210
      TabIndex        =   4
      Top             =   2220
      Width           =   1245
   End
   Begin VSFlex7LCtl.VSFlexGrid vsQuery 
      Height          =   2565
      Left            =   90
      TabIndex        =   0
      Top             =   1290
      Width           =   5955
      _cx             =   10504
      _cy             =   4524
      _ConvInfo       =   -1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
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
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   8
      Cols            =   8
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
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSComCtl2.DTPicker dtpEndTime 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy-M-d"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   5220
      TabIndex        =   9
      Top             =   270
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   23658499
      CurrentDate     =   37022
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "��ѯ���ͣ�"
      Height          =   180
      Left            =   90
      TabIndex        =   12
      Top             =   630
      Width           =   900
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "��Ʊ�ڣ�"
      Height          =   180
      Left            =   3780
      TabIndex        =   11
      Top             =   630
      Width           =   720
   End
   Begin VB.Label Label3 
      Caption         =   "��ѯ����б�"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   990
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "��ѯ��ֹʱ�䣺"
      Height          =   195
      Left            =   3780
      TabIndex        =   2
      Top             =   300
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "��ѯ��ʼʱ�䣺"
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   270
      Width           =   1305
   End
End
Attribute VB_Name = "frmSellTicketQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_bDefault As Boolean
Public m_nQueryType As Integer
Public m_dtStartTime As Date
Public m_dtEndTime As Date
Public m_szTitleID As String

Private m_rsTitle As New Recordset

Private Sub cboQueryType_Click()
    m_nQueryType = cboQueryType.ListIndex + 1
    InitVsQuery
End Sub

Private Sub cmdPreView_Click()
Dim cellTemp As New cCellTemplate
Dim aszFormatData() As String
cellTemp.TemplateFileName = m_szTemplatePathName

If Not m_rsTitle Is Nothing Then
    ReDim aszFormatData(1 To 2, 1 To 2)
    aszFormatData(1, 1) = "ͳ�ƿ�ʼ����"
    aszFormatData(1, 2) = Format(dtpStartTime.Value, "YYYY��MM��DD�� HHʱMM��SS��")
    aszFormatData(2, 1) = "ͳ�ƽ�������"
    aszFormatData(2, 2) = Format(dtpEndTime.Value, "YYYY��MM��DD�� HHʱMM��SS��")
    m_rsTitle.MoveFirst
    cellTemp.FillCellWithRecordset m_rsTitle
    cellTemp.DoPrintSheetPreview
End If

End Sub

Private Sub cmdPrint_Click()
Dim cellTemp As New cCellTemplate
Dim aszFormatData() As String
cellTemp.TemplateFileName = m_szTemplatePathName

If Not m_rsTitle Is Nothing Then
    ReDim aszFormatData(1 To 2, 1 To 2)
    aszFormatData(1, 1) = "ͳ�ƿ�ʼ����"
    aszFormatData(1, 2) = Format(dtpStartTime.Value, "YYYY��MM��DD�� HHʱMM��SS��")
    aszFormatData(2, 1) = "ͳ�ƽ�������"
    aszFormatData(2, 2) = Format(dtpEndTime.Value, "YYYY��MM��DD�� HHʱMM��SS��")
    m_rsTitle.MoveFirst
    cellTemp.FillCellWithRecordset m_rsTitle
    cellTemp.DoPrintSheet
End If

End Sub

Private Sub cmdQuery_Click()
    ResultShow
End Sub

Private Sub cmdSave_Click()
Dim cellTemp As New cCellTemplate
Dim szPathName As String
Dim aszFormatData() As String
cellTemp.TemplateFileName = m_szTemplatePathName

SaveDialogue.Filter = "*.cll"
SaveDialogue.ShowSave
szPathName = SaveDialogue.FileName
If Not m_rsTitle Is Nothing Then
    ReDim aszFormatData(1 To 2, 1 To 2)
    aszFormatData(1, 1) = "ͳ�ƿ�ʼ����"
    aszFormatData(1, 2) = Format(dtpStartTime.Value, "YYYY��MM��DD�� HHʱMM��SS��")
    aszFormatData(2, 1) = "ͳ�ƽ�������"
    aszFormatData(2, 2) = Format(dtpEndTime.Value, "YYYY��MM��DD�� HHʱMM��SS��")
    
    m_rsTitle.MoveFirst
    cellTemp.FillCellWithRecordset m_rsTitle, aszFormatData
    cellTemp.DoSave szPathName
End If

End Sub

Private Sub Form_Load()
    InitVsQuery
    dtpStartTime.Value = m_dtStartTime
    dtpEndTime.Value = m_dtEndTime
    tbbTitileID.Text = m_szTitleID
    If m_bDefault = True Then ResultShow
    InitCboQueryType
End Sub

Private Sub ResultShow()
Dim rsTemp As Recordset
Dim oTemp As New SellerFinance
Dim i As Integer
Select Case m_nQueryType
    Case 1  '����·��ѯ
       Set rsTemp = oTemp.GetSellRouteID(dtpStartTime.Value, dtpEndTime.Value, tbbTitileID.Text)
       vsQuery.Rows = 1
       vsQuery.Cols = 5
       If rsTemp.RecordCount <> 0 Then
                
            For i = 1 To rsTemp.RecordCount
                vsQuery.AddItem ""
                vsQuery.TextMatrix(i, 0) = i
                vsQuery.TextMatrix(i, 1) = rsTemp!route_id
                vsQuery.TextMatrix(i, 2) = rsTemp!route_name
                vsQuery.TextMatrix(i, 3) = rsTemp!ticketcount
                vsQuery.TextMatrix(i, 4) = Format(rsTemp!totalprice, "0.00")
                rsTemp.MoveNext
            Next i
       End If
    Case 2 '�����β�ѯ
        Set rsTemp = oTemp.GetSellBusID(dtpStartTime.Value, dtpEndTime.Value, tbbTitileID.Text)
        vsQuery.Rows = 1
        vsQuery.Cols = 5
        
        If rsTemp.RecordCount <> 0 Then
            For i = 1 To rsTemp.RecordCount
                vsQuery.AddItem ""
                vsQuery.TextMatrix(i, 0) = i
                vsQuery.TextMatrix(i, 1) = rsTemp!bus_id
                vsQuery.TextMatrix(i, 2) = rsTemp!vehicle_type_name
                vsQuery.TextMatrix(i, 3) = rsTemp!ticketcount
                vsQuery.TextMatrix(i, 4) = Format(rsTemp!totalprice, "0.00")
                rsTemp.MoveNext
            Next i
       End If
    Case 3  '����Ʊ�ڲ�ѯ
        Set rsTemp = oTemp.GetSellCheckGateID(dtpStartTime.Value, dtpEndTime.Value, tbbTitileID.Text)
        vsQuery.Rows = 1
        vsQuery.Cols = 4
        If rsTemp.RecordCount <> 0 Then
            For i = 1 To rsTemp.RecordCount
                vsQuery.AddItem ""
                vsQuery.TextMatrix(i, 0) = i
                vsQuery.TextMatrix(i, 1) = rsTemp!check_gate_name
                vsQuery.TextMatrix(i, 2) = rsTemp!ticketcount
                vsQuery.TextMatrix(i, 3) = Format(rsTemp!totalprice, "0.00")
                rsTemp.MoveNext
            Next i
       End If
    Case 4  '��վ���ѯ
        Set rsTemp = oTemp.GetSellStationID(dtpStartTime.Value, dtpEndTime.Value, tbbTitileID.Text)
        vsQuery.Rows = 1
        vsQuery.Cols = 5
        If rsTemp.RecordCount <> 0 Then
            For i = 1 To rsTemp.RecordCount
                vsQuery.AddItem ""
                vsQuery.TextMatrix(i, 0) = i
                vsQuery.TextMatrix(i, 1) = rsTemp!des_station_id
                vsQuery.TextMatrix(i, 2) = rsTemp!station_name
                vsQuery.TextMatrix(i, 3) = rsTemp!ticketcount
                vsQuery.TextMatrix(i, 4) = Format(rsTemp!totalprice, "0.00")
                rsTemp.MoveNext
            Next i
       End If
End Select
Set m_rsTitle = rsTemp
Set rsTemp = Nothing
End Sub

'��ʼ��vsQuery
Private Sub InitVsQuery()
    vsQuery.Editable = flexEDNone
    vsQuery.MergeCells = flexMergeRestrictRows
    vsQuery.MergeCol(1) = True
    vsQuery.FixedAlignment(-1) = flexAlignCenterCenter
    vsQuery.Rows = 1
    vsQuery.ColWidth(-1) = 1300
    vsQuery.ColWidth(0) = 600
    vsQuery.RowHeight(-1) = 300
    Select Case m_nQueryType
        Case 1  '����·��ѯ
            Me.Caption = "��Ʊ��ʱ��ѯ��������·����"
            vsQuery.Cols = 5
            lblTitle.Caption = "��·���룺"
            vsQuery.TextMatrix(0, 0) = "���"
            vsQuery.TextMatrix(0, 1) = "��·����"
            vsQuery.TextMatrix(0, 2) = "��·����"
            vsQuery.TextMatrix(0, 3) = "��Ʊ��"
            vsQuery.TextMatrix(0, 4) = "�ܽ��"

        Case 2 '�����β�ѯ
            Me.Caption = "��Ʊ��ʱ��ѯ���������δ���"
            vsQuery.Cols = 5
            lblTitle.Caption = "���δ��룺"
            vsQuery.TextMatrix(0, 0) = "���"
            vsQuery.TextMatrix(0, 1) = "���δ���"
            vsQuery.TextMatrix(0, 2) = "����"
            vsQuery.TextMatrix(0, 3) = "��Ʊ��"
            vsQuery.TextMatrix(0, 4) = "�ܽ��"
        Case 3  '����Ʊ�ڲ�ѯ
            Me.Caption = "��Ʊ��ʱ��ѯ��������Ʊ��"
            vsQuery.Cols = 4
            lblTitle.Caption = "��Ʊ�ڴ��룺"
            vsQuery.TextMatrix(0, 0) = "���"
            vsQuery.TextMatrix(0, 1) = "��Ʊ��"
            vsQuery.TextMatrix(0, 2) = "��Ʊ��"
            vsQuery.TextMatrix(0, 3) = "�ܽ��"
        Case 4  '��վ���ѯ
            Me.Caption = "��Ʊ��ʱ��ѯ������վ�����"
            vsQuery.Cols = 5
            lblTitle.Caption = "վ����룺"
            vsQuery.TextMatrix(0, 0) = "���"
            vsQuery.TextMatrix(0, 1) = "վ�����"
            vsQuery.TextMatrix(0, 2) = "վ������"
            vsQuery.TextMatrix(0, 3) = "��Ʊ��"
            vsQuery.TextMatrix(0, 4) = "�ܽ��"
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set m_rsTitle = Nothing

End Sub

Private Sub tbbTitileID_Click()
    Dim frmTemp As New frmAllTitleShow
    frmTemp.m_nTitleType = m_nQueryType
    frmTemp.m_dtStartTime = dtpStartTime.Value
    frmTemp.m_dtEndTime = dtpEndTime
    Set frmTemp.frmTitle = Me
    frmTemp.Show vbModal
    Set frmTemp = Nothing
End Sub

Private Sub InitCboQueryType()
    cboQueryType.AddItem "����·��ѯ"
    cboQueryType.AddItem "�����β�ѯ"
    cboQueryType.AddItem "����Ʊ�ڲ�ѯ"
    cboQueryType.AddItem "��վ���ѯ"
    cboQueryType.ListIndex = m_nQueryType - 1
End Sub

