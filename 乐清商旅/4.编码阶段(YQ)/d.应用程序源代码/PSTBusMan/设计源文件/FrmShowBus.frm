VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmShowBus 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ѡ�񳵴�"
   ClientHeight    =   5985
   ClientLeft      =   1950
   ClientTop       =   2265
   ClientWidth     =   10140
   ControlBox      =   0   'False
   Icon            =   "FrmShowBus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   10140
   Begin VB.ListBox lstSeatType 
      Appearance      =   0  'Flat
      Columns         =   2
      Height          =   1290
      Left            =   6645
      MultiSelect     =   2  'Extended
      TabIndex        =   14
      Top             =   2940
      Width           =   3375
   End
   Begin VB.ListBox lstType 
      Appearance      =   0  'Flat
      Columns         =   2
      Height          =   2190
      Left            =   6645
      MultiSelect     =   2  'Extended
      TabIndex        =   13
      Top             =   390
      Width           =   3375
   End
   Begin VB.ComboBox cboSellStation 
      Height          =   300
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   150
      Width           =   2535
   End
   Begin VB.CheckBox ChkStop 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "ֻ����ͣ��վ��(&S)"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6645
      TabIndex        =   8
      Top             =   5010
      Value           =   1  'Checked
      Width           =   1830
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1620
      Top             =   510
   End
   Begin VB.CheckBox chkEmpty 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "���ɿյ�Ʊ����(&K)"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6645
      TabIndex        =   7
      Top             =   4635
      Width           =   2000
   End
   Begin VB.CheckBox ChkExist 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "�Ѵ��ڵĳ�������λ���͵ĳ���Ʊ��(E)"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6645
      TabIndex        =   2
      Top             =   4305
      Width           =   3645
   End
   Begin VB.ComboBox cboPriceTable 
      Height          =   300
      Left            =   1035
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   150
      Width           =   1755
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   345
      Left            =   7350
      TabIndex        =   3
      Top             =   5520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "ȷ��"
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
      MICON           =   "FrmShowBus.frx":000C
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
      Left            =   8805
      TabIndex        =   4
      Top             =   5520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "ȡ��"
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
      MICON           =   "FrmShowBus.frx":0028
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
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "��λ����ѡ��(&Y):"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6645
      TabIndex        =   6
      Top             =   2685
      Width           =   3375
   End
   Begin VB.Frame framSelVehicleType 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "����ѡ��(&T):"
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   6645
      TabIndex        =   5
      Top             =   165
      Width           =   3375
   End
   Begin MSComctlLib.ListView lvBus 
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   780
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����ʱ��"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "��������"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "������·"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "�յ�վ"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ϳ�վ(&D):"
      Height          =   180
      Left            =   3030
      TabIndex        =   11
      Top             =   210
      Width           =   900
   End
   Begin VB.Label lblExcuteTable 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ʊ�۱�(&P):"
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   210
      Width           =   900
   End
   Begin VB.Label lblBus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����(&B):"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   555
      Width           =   720
   End
End
Attribute VB_Name = "frmShowBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'* Source File Name:frmShowBus.frm
'* Project Name:PSTBusMan.vbp
'* Engineer:�·�
'* Date Generated:2002/09/03
'* Last Revision Date:2002/09/09
'* Brief Description:�򿪡�������ɾ������Ʊ��ʱ������ѡ����������
'* Relational Document:
'****************************************************************

Option Explicit
Public m_bOk As Boolean '�Ƿ���Ok
Public m_eFormStatus As EFormStatus
Public m_bEnabledStop As Boolean '�Ƿ���ʾ "ֻ����ͣ��վ��"��CheckBox


Private m_szPriceTable As String 'ѡ���Ʊ�۱�
Private m_aszBusID() As String 'ѡ��ĳ���
Private m_aszVehicleType() As String 'ѡ��ĳ���
Private m_aszSeatType() As String 'ѡ�����λ����
'Private m_bExist As Boolean '�Ƿ�ֻȡ���ڵĳ�����λ��
Private m_bEmpty As Boolean '�Ƿ�ֻ���ɿյ�Ʊ����

Private m_oTicketPriceMan As New TicketPriceMan


Private m_aszSellStationID() As String 'ѡ����ϳ�վ
Const cszAllSellStation = "(�����ϳ�վ)"


Public Property Get GetBusID() As String()
    GetBusID = m_aszBusID
End Property

Public Property Get GetVehicleType() As String()
    GetVehicleType = m_aszVehicleType
End Property

Public Property Get GetSeatType() As String()
    GetSeatType = m_aszSeatType
End Property

Public Property Get GetPriceTableID() As String
    GetPriceTableID = m_szPriceTable
End Property

Public Property Get GetSellStation() As String()
    GetSellStation = m_aszSellStationID
End Property


'Public Property Get IsExist() As Boolean
'    IsExist = m_bExist
'End Property

Public Property Get IsEmpty() As Boolean
    IsEmpty = m_bEmpty
End Property

Public Property Get IsOnlyStop() As Boolean
    IsOnlyStop = m_bEnabledStop
End Property

Private Sub ChkExist_Click()
    '�����Ƿ�����ѡ��������λ����
    If ChkExist.Value = vbChecked Then
        lstType.Enabled = False
        lstSeatType.Enabled = False
    Else
        lstType.Enabled = True
        lstSeatType.Enabled = True
    End If
    EnableOk
End Sub

Private Sub cmdCancel_Click()
    m_bOk = False
    Unload Me
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    m_bOk = False
    If m_bEnabledStop Then ChkStop.Enabled = True
    
    Select Case m_eFormStatus
    Case EFS_AddNew
        '��������Ʊ��
        Me.Caption = "��ѡ�������ĳ���"
        cmdOk.Caption = "����(&O)"
'        If Not m_bEnabledStop Then
'            '��������Զ�����
            chkEmpty.Enabled = True
            cboSellStation.Enabled = False
'        End If
    Case EFS_Show
        '�򿪳���Ʊ��
        Me.Caption = "��ѡ��򿪵ĳ���"
        cmdOk.Caption = "��(&O)"
        chkEmpty.Enabled = False
        cboSellStation.Enabled = True
    Case EFS_Delete
        'ɾ������Ʊ��
        Me.Caption = "��ѡ��Ҫɾ���ĳ���"
        cmdOk.Caption = "ɾ��(&O)"
        chkEmpty.Enabled = False
        cboSellStation.Enabled = False
        '*****�˴����貹��
    End Select
    
    FillSellStation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_oTicketPriceMan = Nothing
    m_bEnabledStop = False
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    SetBusy
    m_oTicketPriceMan.Init g_oActiveUser
    lvBus.MultiSelect = True
    FillPriceTable
    FillVehicleType
    FillSeatType
    FillBus g_szExePriceTable
    EnableOk
    ChkExist.Value = vbChecked
    SetNormal
    
End Sub

Private Sub FillPriceTable()
'Private Sub FillPriceTable(ProjectID As String)
    '���Ʊ�۱�
    On Error GoTo ErrorHandle
    Dim oRegularScheme As New RegularScheme
    Dim aszTable() As String
    Dim i As Integer, nCount As Integer

    oRegularScheme.Init g_oActiveUser
    aszTable = oRegularScheme.ProjectExistTable
    nCount = ArrayLength(aszTable)
    cboPriceTable.Clear
'    If nCount > 0 Then
        For i = 1 To nCount
            cboPriceTable.AddItem MakeDisplayString(aszTable(i, 2), aszTable(i, 3))
            If Trim(aszTable(i, 2)) = Trim(g_szExePriceTable) Then
                cboPriceTable.ListIndex = i - 1
            End If
        Next
'    End If
    Set oRegularScheme = Nothing
    Exit Sub

ErrorHandle:
    MsgBox "�˼ƻ�����ӦƱ�۱�"
End Sub

Private Sub EnableOk()
    '�ж�cmdOk�Ƿ����
    If ChkExist.Value = vbChecked Then
        If lvBus.ListItems.Count > 0 And cboPriceTable.ListIndex >= 0 Then
            cmdOk.Enabled = True
        Else
            cmdOk.Enabled = False
        End If
    Else
        If lstType.SelCount = 0 Or lstSeatType.SelCount = 0 Then
            cmdOk.Enabled = False
        Else
            cmdOk.Enabled = True
        End If
    End If
End Sub

Private Sub FillVehicleType()
    '��䳵��
    Dim oBase As New BaseInfo
    Dim aszVehicleType() As String
    Dim i As Integer, nCount As Integer
    lstType.Clear
    oBase.Init g_oActiveUser
    aszVehicleType = oBase.GetAllVehicleModel()
    nCount = ArrayLength(aszVehicleType)
    For i = 1 To nCount
        lstType.AddItem MakeDisplayString(RTrim(aszVehicleType(i, 1)), RTrim(aszVehicleType(i, 2)))
    Next
End Sub

Private Sub FillSeatType()
    'ˢ����λ������Ϣ
    Dim oBase As New BaseInfo
    Dim aszAllSeatType() As String
    Dim i As Integer
    Dim nCount As Integer
    lstSeatType.Clear
    oBase.Init g_oActiveUser
    aszAllSeatType = oBase.GetAllSeatType
    
    nCount = ArrayLength(aszAllSeatType)
    For i = 1 To nCount
        lstSeatType.AddItem MakeDisplayString(RTrim(aszAllSeatType(i, 1)), RTrim(aszAllSeatType(i, 2)))
    Next i
End Sub

Private Sub FillBus(ByVal pszProjectID As String)
    'ˢ�³�����Ϣ
    On Error GoTo ErrorHandle
    Dim oProject As BusProject
    Dim nDataCount As Integer, i As Integer
    Dim liTemp As ListItem
    Dim aszTemp() As String
    Set oProject = New BusProject
    oProject.Init g_oActiveUser
    oProject.Identify
    aszTemp = oProject.GetAllBus()
    nDataCount = ArrayLength(aszTemp)
    lvBus.ListItems.Clear
    For i = 1 To nDataCount
        WriteProcessBar True, i, nDataCount, "����ˢ�³���"
        Set liTemp = lvBus.ListItems.Add(, , RTrim(aszTemp(i, 1)))
        liTemp.ListSubItems.Add , , Format(aszTemp(i, 2), "HH:mm")
        liTemp.ListSubItems.Add , , Trim(aszTemp(i, 8))
        liTemp.ListSubItems.Add , , RTrim(RTrim(aszTemp(i, 4)))
        liTemp.ListSubItems.Add , , RTrim(RTrim(aszTemp(i, 12)))
    Next
    WriteProcessBar False
    If lvBus.ListItems.Count > 0 Then lvBus.ListItems(1).Selected = True
    
ErrorHandle:
End Sub

Private Sub lvbus_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvBus, ColumnHeader.Index
End Sub

Private Sub lstSeatType_Click()
    EnableOk
End Sub

Private Sub lstType_Click()
    EnableOk
End Sub

Private Sub cmdOk_Click()
    '�õ���ѡ��ĳ��ͺ���λ�������ȡ�üƻ��ó��εĳ��ͺ���λ���͵Ľ���
    Dim i As Integer, j As Integer ', n As Integer
    Dim nBus As Integer
    Dim nVehicleType As Integer
    Dim nSeatType As Integer
    Dim ttBusVehicleSeat() As TBusVehicleSeatType
    Dim nTemp As Long
    Dim Count As Long
    On Error GoTo ErrorHandle
    SetBusy
    m_bOk = True
    m_szPriceTable = ResolveDisplay(cboPriceTable.Text)
    '�õ�ѡ��ĳ���
'    For i = 1 To lvBus.ListItems.Count
'        If lvBus.ListItems(i).Selected = True Then nBus = nBus + 1
'    Next i
'    ReDim m_aszBusId(1 To nBus)
    ReDim m_aszSellStationID(1 To 1)
    '�õ�ѡ����ϳ�վ
    If cboSellStation.Text = cszAllSellStation Then
        m_aszSellStationID(1) = ""
    Else
        m_aszSellStationID(1) = ResolveDisplay(cboSellStation.Text)
    End If
    
    For i = 1 To lvBus.ListItems.Count
        If lvBus.ListItems(i).Selected = True Then
           j = j + 1
           ReDim Preserve m_aszBusID(1 To j)
           m_aszBusID(j) = lvBus.ListItems(i).Text
        End If
    Next i
    '�Ƿ�ѡ�������ɿ�Ʊ��
    If chkEmpty.Enabled = True Then
        If chkEmpty.Value = vbChecked Then
            m_bEmpty = True
        Else
            m_bEmpty = False
        End If
    End If
    If ChkExist.Value = vbChecked Then
'        m_bExist = True
        'ѡ����ֻ�򿪴��ڵĳ�������λ����
        ttBusVehicleSeat = m_oTicketPriceMan.GetAllBusVehicleTypeSeatType(m_aszBusID)
        nTemp = ArrayLength(ttBusVehicleSeat)
        If nTemp > 0 Then
           ReDim m_aszBusID(1 To nTemp)
           ReDim m_aszSeatType(1 To nTemp)
           ReDim m_aszVehicleType(1 To nTemp)
        End If
        For i = 1 To nTemp
            m_aszBusID(i) = ttBusVehicleSeat(i).szbusID
            m_aszSeatType(i) = ttBusVehicleSeat(i).szSeatTypeID
            m_aszVehicleType(i) = ttBusVehicleSeat(i).szVehicleTypeCode
        Next i
    Else
'        m_bExist = False
        '�õ���ѡ��ĳ�������λ����
        For i = 1 To lstType.ListCount
            If lstType.Selected(i - 1) = True Then
                nVehicleType = nVehicleType + 1
                ReDim Preserve m_aszVehicleType(1 To nVehicleType)
                m_aszVehicleType(j) = ResolveDisplay(lstType.List(i - 1))
            End If
        Next i
        For i = 1 To lstSeatType.ListCount
            If lstSeatType.Selected(i - 1) = True Then
                nSeatType = nSeatType + 1
                ReDim Preserve m_aszSeatType(1 To nSeatType)
                m_aszSeatType(j) = ResolveDisplay(lstSeatType.List(i - 1))
            End If
        Next i
    End If
    SetNormal
    Unload Me
    Exit Sub
ErrorHandle:
    SetNormal
End Sub


'===================================================
'Modify Date��2002-11-19
'Author:fl
'Reamrk:������е��ϳ�վ��
'===================================================b

Private Sub FillSellStation()

    '����ϳ�վ
    Dim nCount As Integer
    Dim i As Integer
    cboSellStation.Clear
    nCount = ArrayLength(g_atAllSellStation)
    cboSellStation.AddItem cszAllSellStation
    For i = 1 To nCount
        cboSellStation.AddItem MakeDisplayString(g_atAllSellStation(i).szSellStationID, g_atAllSellStation(i).szSellStationName)
        'cszAllSellStation
    Next i
    
    '������е��ϳ�վ
    If nCount > 0 Then cboSellStation.ListIndex = 0
End Sub

