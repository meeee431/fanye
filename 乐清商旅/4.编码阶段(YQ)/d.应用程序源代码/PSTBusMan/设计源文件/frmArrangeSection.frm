VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmArrangeSection 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��·--·�ι���"
   ClientHeight    =   5445
   ClientLeft      =   3465
   ClientTop       =   2775
   ClientWidth     =   9390
   HelpContextID   =   2001401
   Icon            =   "frmArrangeSection.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboSellStation 
      Height          =   300
      Left            =   4140
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   390
      Width           =   2265
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   8040
      TabIndex        =   12
      Top             =   5100
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
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
      MICON           =   "frmArrangeSection.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer tmStart 
      Interval        =   500
      Left            =   3570
      Top             =   2130
   End
   Begin MSComctlLib.ListView lvSection 
      Height          =   2415
      Left            =   6510
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "·�δ���"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "·����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "���վ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�յ�վ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "�����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "��·�ȼ�"
         Object.Width           =   2540
      EndProperty
   End
   Begin RTBusMan.ucSuperCombo cboStation 
      Height          =   4350
      Left            =   6495
      TabIndex        =   1
      Top             =   720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   7673
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1995
      Top             =   1545
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrangeSection.frx":0166
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrangeSection.frx":02C2
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrangeSection.frx":041E
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrangeSection.frx":057A
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrangeSection.frx":06D6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrangeSection.frx":0832
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrangeSection.frx":098E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrangeSection.frx":0AEA
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrangeSection.frx":13C6
            Key             =   "SellTicket"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrangeSection.frx":1522
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrangeSection.frx":167E
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrangeSection.frx":17DA
            Key             =   "NoSellTicket"
         EndProperty
      EndProperty
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   300
      Left            =   8025
      TabIndex        =   4
      Top             =   5595
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   2
      TX              =   ""
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
      MICON           =   "frmArrangeSection.frx":20B6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvRouteSection 
      Height          =   4665
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   8229
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "��"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "·�δ���"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "·����"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "���վ"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "�յ�վ"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "�����"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "վ������"
         Object.Width           =   1411
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbSection 
      Height          =   360
      Left            =   7335
      TabIndex        =   14
      Top             =   90
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "AddSection"
            Object.ToolTipText     =   "��������·��"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ModifySection"
            Object.ToolTipText     =   "�༭·��"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DeleteSection"
            Object.ToolTipText     =   "ɾ������·��"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "InsertSection"
            Object.ToolTipText     =   "����·��"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "InsertRout"
            Object.ToolTipText     =   "����·��"
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "SectionStation"
            Object.ToolTipText     =   "վ����Ʊ�趨"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Station"
            Object.ToolTipText     =   "վ������"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "ˢ��"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���վ:"
      Height          =   180
      Left            =   3465
      TabIndex        =   13
      Top             =   465
      Width           =   630
   End
   Begin VB.Label lblRouteSection 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��··���б�(&R):"
      Height          =   180
      Left            =   135
      TabIndex        =   2
      Top             =   495
      Width           =   1440
   End
   Begin VB.Label lblStation 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "վ��(&S):"
      Height          =   180
      Left            =   6510
      TabIndex        =   0
      Top             =   465
      Width           =   720
   End
   Begin VB.Label lblRoute 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   930
      TabIndex        =   11
      Top             =   180
      Width           =   90
   End
   Begin VB.Label lblMileage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   5070
      TabIndex        =   10
      Top             =   180
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����:"
      Height          =   180
      Left            =   4410
      TabIndex        =   9
      Top             =   180
      Width           =   630
   End
   Begin VB.Label lblRouteName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   2490
      TabIndex        =   8
      Top             =   180
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��·����:"
      Height          =   180
      Left            =   1665
      TabIndex        =   7
      Top             =   180
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��·����:"
      Height          =   180
      Left            =   135
      TabIndex        =   6
      Top             =   180
      Width           =   810
   End
   Begin VB.Menu mnu_Route 
      Caption         =   "��·����"
      Visible         =   0   'False
      Begin VB.Menu mnu_Section 
         Caption         =   "·����Ϣ(&R)"
      End
      Begin VB.Menu mnu_break 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Delete 
         Caption         =   "ɾ�����·��(&D)"
      End
   End
End
Attribute VB_Name = "frmArrangeSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************
'* Source File Name:frmArrangeSection.frm
'* Project Name:RTBusMan
'* Engineer:
'* Data Generated:2002/08/27
'* Last Revision Date:2002/09/02
'* Brief Description:��·����·��
'* Relational Document:
'**********************************************************
Public m_szRouteID As String
Public m_bIsParent As Boolean '�Ƿ��Ǹ�����ֱ�ӵ���

Private m_oBaseInfo As New BaseInfo
Private m_oRoute As New Route
Private m_oSection As New Section
Private m_szStation As String




Private Sub Form_Unload(Cancel As Integer)
    Dim szStationID As String
    Dim szStationName As String
    
    If m_bIsParent Then
        frmRoute.lblMileage.Caption = lblMileage.Caption
        If m_szStation <> "" Then
            szStationID = ResolveDisplay(m_szStation, szStationName)
            frmRoute.lblEndStation.Caption = szStationName
            frmRoute.lblEndStation.Tag = szStationID
        End If
    End If
    m_bIsParent = False
    frmAllRoute.UpdateList lblRoute.Caption
End Sub

Private Sub lvSection_DblClick()
    If lvSection.SelectedItem Is Nothing Then Exit Sub
    AppendSectionToLv lvSection.SelectedItem.Text
End Sub

Private Sub lvSection_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If lvSection.SelectedItem Is Nothing Then Exit Sub
        
        AppendSectionToLv lvSection.SelectedItem.Text
    End If
End Sub

Private Sub lvSection_LostFocus()
    lvSection.Visible = False
End Sub


'===================================================
'Modify Date��2002-11-19
'Author:fl
'Reamrk:;��վ����
'===================================================b
'
Private Sub cboStation_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
'        If lvRouteSection.ListItems.Count = 0 Then
''            lblStartStation.Caption = Trim(cboStation.BoundText)
'
'        Else
        AppendSection
'        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()
    '��ʼ������
    m_oBaseInfo.Init g_oActiveUser
    m_oRoute.Init g_oActiveUser
    m_oSection.Init g_oActiveUser
    
End Sub

Private Sub lvRouteSection_DblClick()
    
    EditSection
End Sub

Private Sub tbSection_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "AddSection"
        AddAnotherSectionToDB
    Case "ModifySection"
        EditSection
    Case "DeleteSection"
        RemoveSection
    Case "SectionStation"
        
    Case "Station"
        
    Case "Refresh"
        FillStation
    Case "InsertSection"
        Load frmInsertSection
        frmInsertSection.txtRouteID.Text = m_szRouteID
        frmInsertSection.txtDelSectionID.Text = MakeDisplayString(lvRouteSection.SelectedItem.ListSubItems(1).Text, Trim(lvRouteSection.SelectedItem.ListSubItems(2).Text))
'        m_szCheckText = ""
'        If m_nCount = 1 Then
'        frmInsertSection.txtDelSectionID.Text = Trim(lvRouteSection.ListItems(1).ListSubItems(1).Text) + "[" + Trim(lvRouteSection.ListItems(1).ListSubItems(2).Text) + "]"
'        End If
        frmInsertSection.Show vbModal
    End Select
End Sub

'===================================================
'Modify Date��2002-11-19
'Author:fl
'Reamrk:����ѯ���������еı���λ����Ʊվ"����[����]"��䵽cboSellStation��
'===================================================b

Private Sub FillStartStation()
    Dim aszTmp() As String
    aszTmp = m_oRoute.GetStartStation
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    Dim i As Integer
    cboSellStation.Clear
    For i = 1 To ArrayLength(aszTmp)
        cboSellStation.AddItem MakeDisplayString(aszTmp(i, 1), aszTmp(i, 2))
    Next i
    cboSellStation.ListIndex = 0
    '�����Ʊվ
  
    '������е���Ʊվ
End Sub

Private Sub tmStart_Timer()
    SetBusy
    tmStart.Enabled = False
    FillLvRouteSection
    FillStartStation
    FillStation
    RefreshRouteInfo
    SetNormal
End Sub

Private Sub AppendSection()
    '������·׷��վ��
    Dim aszSection() As String
    Dim liTemp As ListItem
    Dim nResult As VbMsgBoxResult
    Dim szStartStation As String, szEndStation As String
    Dim nCount As Integer
    Dim i As Integer
    
    On Error GoTo ErrorHandle
    szEndStation = Trim(cboStation.BoundText)
    If lvRouteSection.ListItems.Count = 0 Then
        szStartStation = ResolveDisplay(cboSellStation.Text)    'ȡ�����վ
    Else
        szStartStation = ResolveDisplay(lvRouteSection.ListItems.Item(lvRouteSection.ListItems.Count).ListSubItems(4).Text)
    End If
    '�õ�·����Ϣ
    aszSection = m_oSection.GetSESection(szStartStation, szEndStation)
    nCount = ArrayLength(aszSection)
    If nCount > 1 Then
        '�����㵽�յ���ж��·��,��ѡ��·��
        lvSection.Visible = True
        lvSection.ListItems.Clear
        For i = 1 To nCount
            m_oSection.Identify aszSection(i)
            Set liTemp = lvSection.ListItems.Add(, , aszSection(i))
            liTemp.SubItems(1) = m_oSection.SectionName
            liTemp.SubItems(2) = MakeDisplayString(m_oSection.BeginStationCode, m_oSection.BeginStationName)
            liTemp.SubItems(3) = MakeDisplayString(m_oSection.EndStationCode, m_oSection.EndStationName)
            liTemp.SubItems(4) = m_oSection.Mileage
            liTemp.SubItems(5) = m_oSection.RoadLevelName
            
        Next
    ElseIf nCount = 1 Then
        AppendSectionToLv aszSection(1)
    ElseIf nCount = 0 Then
        AddSectionToDB szStartStation, szEndStation
    End If
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub
'
Public Sub AppendSectionToLv(pszSection As String)
    '����Ϣ׷�ӵ�·�ε�ListView��
On Error GoTo ErrorHandle
    m_oSection.Identify pszSection
    m_oRoute.AddLastSection pszSection
    AddList pszSection
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub AddSectionToDB(pszStartStation As String, pszEndStation As String)
    '�������վ���յ�վ��·�μӵ����ݿ���,������·�μӵ�����·��
    Dim oStation As New Station
    Dim szSectionName As String
    Dim szSectionID As String
    Dim nResult  As VbMsgBoxResult
    nResult = MsgBox("������վ��·�β�����,�Ƿ�����·��", vbQuestion + vbYesNo, "��·")
    If nResult = vbYes Then
        On Error GoTo ErrorHandle
        oStation.Init g_oActiveUser
        oStation.Identify pszStartStation
        frmSection.m_bRouteArrange = True
        frmSection.m_eStatus = EFS_AddNew
        frmSection.m_szStartStation = MakeDisplayString(oStation.StationID, oStation.StationName)
        szSectionID = Left(oStation.StationInputCode, 2)
        szSectionName = Left(oStation.StationName, 2)
        oStation.Identify pszEndStation
        frmSection.m_szEndStation = MakeDisplayString(oStation.StationID, oStation.StationName)
        frmSection.m_szSectionID = szSectionID & Left(oStation.StationInputCode, 2)
        frmSection.m_szSectionName = szSectionName & Left(oStation.StationName, 2)
        frmSection.m_szAreaID = MakeDisplayString(oStation.AreaCode, oStation.AreaName)
        frmSection.Show vbModal
    End If
    Set oStation = Nothing
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub AddAnotherSectionToDB()
    '��Ӷ����·�ε����ݿ���,����ԭ����㵽�յ�����һ������,������Ҫ�ټ�һ��,�������վ���յ�վ��·�μӵ����ݿ���,������·�μӵ�����·��
    Dim oStation As New Station
    Dim szSectionName As String
    Dim szSectionID As String
    Dim nResult  As VbMsgBoxResult
    Dim szStartStation As String
    Dim szEndStation As String
    
        On Error Resume Next
        
        szEndStation = Trim(cboStation.BoundText)
        If lvRouteSection.ListItems.Count = 0 Then
            szStartStation = g_szStationID 'ȡ�����վ
        Else
            szStartStation = ResolveDisplay(lvRouteSection.ListItems.Item(lvRouteSection.ListItems.Count).ListSubItems(4).Text)
        End If
        
        oStation.Init g_oActiveUser
        oStation.Identify szStartStation
        frmSection.m_bRouteArrange = True
        frmSection.m_eStatus = EFS_AddNew
        frmSection.m_szStartStation = MakeDisplayString(oStation.StationID, oStation.StationName)
        szSectionID = Left(oStation.StationInputCode, 2)
        szSectionName = Left(oStation.StationName, 2)
        oStation.Identify szEndStation
        frmSection.m_szEndStation = MakeDisplayString(oStation.StationID, oStation.StationName)
        frmSection.m_szSectionID = szSectionID & Left(oStation.StationInputCode, 2)
        frmSection.m_szSectionName = szSectionName & Left(oStation.StationName, 2)
        frmSection.m_szAreaID = MakeDisplayString(oStation.AreaCode, oStation.AreaName)
        frmSection.Show vbModal
    Set oStation = Nothing
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub FillStation()
    '������е�վ��
    Dim rsStation As New Recordset
    Dim szaStation() As String
    Dim nCount As Integer, i As Integer
    On Error GoTo ErrorHandle
    ShowSBInfo "���վ����Ϣ..."
    cboStation.Clear
    szaStation = m_oBaseInfo.GetStation
    nCount = ArrayLength(szaStation)
    rsStation.Fields.Append "Code", adChar, 10
    rsStation.Fields.Append "Name", adChar, 10
    rsStation.Fields.Append "Input", adChar, 10
    rsStation.Open
    For i = 1 To nCount
        rsStation.AddNew
        rsStation!Code = szaStation(i, 1)
        rsStation!name = szaStation(i, 2)
        rsStation!Input = szaStation(i, 3)
        rsStation.Update
    Next
    Set cboStation.RowSource = rsStation
    cboStation.BoundField = "Code"
    cboStation.ListFields = "Input:6,Name:6,Code:9"
    cboStation.AppendWithFields "Code:9,Name"
    ShowSBInfo
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Public Sub RefreshSection()
    'ˢ�¸���·��·��'
    Dim atSectionInfo() As TRouteSectionInfoEx
    Dim nCount As Integer
    On Error GoTo ErrorHandle
    lvRouteSection.ListItems.Clear
    ShowSBInfo "���·����Ϣ..."
    atSectionInfo = m_oRoute.GetSectionInfoEx
    FillItem atSectionInfo
    RefreshMileage
    ShowSBInfo
    Exit Sub
ErrorHandle:
    ShowErrorMsg
    
End Sub

Private Sub RefreshMileage(Optional pbIsUpdate As Boolean = False)
    'ˢ��������յ�վ��Ϣ
    Dim nCount As Integer
    Dim atSectionInfo() As TRouteSectionInfoEx
    Dim nSectionCount As Integer '·����
    Dim i As Integer
    nCount = lvRouteSection.ListItems.Count
'    For i = 1 To nCount
'        nMileage = nMileage + lvRouteSection.ListItems(i).SubItems(5)
'    Next i
    If nCount > 0 Then
        If Not pbIsUpdate Then
            '������Ǹ���
            m_szStation = lvRouteSection.ListItems(nCount).SubItems(4)
            lblMileage.Caption = lvRouteSection.ListItems(nCount).SubItems(5)
        Else
            If lvRouteSection.SelectedItem Is Nothing Then Exit Sub
            If lvRouteSection.SelectedItem.Index < nCount Then
                '����Ǹ���,��ͬʱ���º�������
                atSectionInfo = m_oRoute.GetSectionInfoEx
                nSectionCount = ArrayLength(atSectionInfo)
                For i = lvRouteSection.SelectedItem.Index To nCount
                    lvRouteSection.ListItems(i).SubItems(5) = atSectionInfo(i).sgEndStationMileage
                Next i
                lblMileage.Caption = lvRouteSection.ListItems(nCount).SubItems(5)
            End If
        End If
    End If
    
End Sub

Public Sub AddList(pszID As String)
    '��������·��ˢ�³���
    Dim atSectionInfo() As TRouteSectionInfoEx
    Dim nCount As Integer
    On Error GoTo ErrorHandle
    
    atSectionInfo = m_oRoute.GetSectionInfoEx(pszID)
    FillItem atSectionInfo
    RefreshMileage
    '*******����Ҫˢ�������Ϣ
    Exit Sub
ErrorHandle:
    ShowErrorMsg
    
End Sub

Public Sub UpdateList(pszID As String)
    '���޸ĵ�·��ˢ�³���
    Dim atSectionInfo() As TRouteSectionInfoEx
    Dim nCount As Integer
    On Error GoTo ErrorHandle
    
    atSectionInfo = m_oRoute.GetSectionInfoEx(pszID)
    FillItem atSectionInfo, True
    '*******����Ҫˢ�������Ϣ
    RefreshMileage True
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub FillItem(patSectionInfo() As TRouteSectionInfoEx, Optional pbIsUpdate As Boolean = False)
    Dim liTemp As ListItem
    Dim i As Integer
    Dim nCount As Integer
    nCount = ArrayLength(patSectionInfo)
    If nCount = 0 Then Exit Sub
    For i = 1 To nCount
        With patSectionInfo(i)
            If Not pbIsUpdate Then
                Set liTemp = lvRouteSection.ListItems.Add(, , .nSectionSerial)
            Else
                Set liTemp = lvRouteSection.SelectedItem
            End If
            liTemp.SubItems(1) = .szSectionID
            liTemp.SubItems(2) = .szSectionName
            liTemp.SubItems(3) = MakeDisplayString(.szStartStationID, .szStartStationName)
            liTemp.SubItems(4) = MakeDisplayString(.szEndStationID, .szEndStationName)
            liTemp.SubItems(5) = .sgEndStationMileage
            If .nEndStationType = TP_CanSellTicket Then
                liTemp.SubItems(6) = "����"
                liTemp.SmallIcon = "SellTicket"
            Else
                liTemp.SubItems(6) = "������"
                liTemp.SmallIcon = "NoSellTicket"
            End If
        End With
    Next
    If nCount > 1 Then
        lvRouteSection.ListItems(1).Selected = True
        lvRouteSection.ListItems(1).EnsureVisible
    Else
        liTemp.Selected = True
        liTemp.EnsureVisible
    End If
End Sub

Private Sub SetToolEnabled()
    '���ù������Ŀ�����
End Sub

Private Sub EditSection()
    '�༭·��
'    Dim oSection As Object
    On Error GoTo ErrorHandle
    If lvRouteSection.SelectedItem Is Nothing Then Exit Sub
'    Set oSection = New frmSection
    frmSection.m_eStatus = EFS_Modify
    frmSection.m_bRouteArrange = True
    frmSection.m_szSectionID = lvRouteSection.SelectedItem.ListSubItems(1).Text
    frmSection.Show vbModal
'    Set oSection = Nothing
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub
Private Sub RefreshRouteInfo()
    'ˢ����·��Ϣ
    m_oRoute.Identify m_szRouteID
    lblRoute.Caption = m_oRoute.RouteID
    lblRouteName.Caption = m_oRoute.RouteName
    lblMileage.Caption = m_oRoute.Mileage
    Dim i As Integer
    For i = 1 To cboSellStation.ListCount
        If ResolveDisplay(cboSellStation.List(i - 1)) = m_oRoute.StartStation Then
            cboSellStation.ListIndex = i - 1
            cboSellStation.Enabled = False
            Exit For
        End If
    Next i
    RefreshSection
End Sub

Private Sub RemoveSection()
    '�Ƴ�(������ɾ��)·��,ֻ�ǽ�·�δӱ���·���Ƴ�
    On Error GoTo ErrorHandle
    Dim nResult As VbMsgBoxResult
    nResult = MsgBox("�Ƿ��Ƴ���·�����·�Σ�", vbQuestion + vbYesNo, Me.Caption)
    If nResult = vbNo Then Exit Sub
    m_oRoute.DeleteLastSection
    lvRouteSection.ListItems.Remove lvRouteSection.ListItems(lvRouteSection.ListItems.Count).Index
'    If lvRouteSection.ListItems.Count > 0 Then
'        lblMileage.Caption = lvRouteSection.ListItems(lvRouteSection.ListItems.Count).ListSubItems(5).Text
'    End If
    RefreshMileage False
    Exit Sub
ErrorHandle:
    ShowErrorMsg

End Sub

Private Sub InsertSection()
    '����·��
End Sub

Private Sub FillLvRouteSection()
    '��:540.2835   ·�δ���:945.0709           ·����:989.8583             ���վ:1440   �յ�վ:1679.811             �����:750.0473             վ������:675.2126
    '�������
    lvRouteSection.ColumnHeaders.Clear
    lvRouteSection.ColumnHeaders.Add , , "��", 540
    lvRouteSection.ColumnHeaders.Add , , "����", 945
    lvRouteSection.ColumnHeaders.Add , , "·����", 989
    lvRouteSection.ColumnHeaders.Add , , "���վ", 1440
    lvRouteSection.ColumnHeaders.Add , , "�յ�վ", 1679
    lvRouteSection.ColumnHeaders.Add , , "���", 750
    lvRouteSection.ColumnHeaders.Add , , "վ������", 675
End Sub



'Private Sub SectionStationSell()
'    Dim liTemp As ListItem
'    Dim nMsg As Integer
'    Dim tSection As TRouteSectionInfo
'    Dim szEndStation As String
'On Error GoTo ErrorHandle
'    Set liTemp = lvRouteSection.SelectedItem
'    szEndStation = LeftAndRight(liTemp.ListSubItems(4).Text, False, "[")
'    nMsg = MsgBox("��··��:[" & liTemp.ListSubItems(2).Text & "]���յ�վ[" & szEndStation & "�Ƿ����Ʊ" & vbCrLf & "* ѡ��[��(Y)]���趨վ��Ϊ����վ��" & vbCrLf & "* ѡ��[��(N)]���趨վ��Ϊ������վ��", vbQuestion + vbYesNoCancel, "վ����Ʊ�趨")
'    If nMsg = vbCancel Then Exit Sub
'    If nMsg = vbYes Then
'        tSection.nEndStationType = TP_CanSellTicket
'    Else
'        tSection.nEndStationType = TP_CanNotSellTicket
'    End If
'    tSection.szSectionID = liTemp.ListSubItems(1).Text
'    tSection.szSectionName = liTemp.ListSubItems(2).Text
'    tSection.sgEndStationMileage = Val(liTemp.ListSubItems(5).Text)
'    m_oRoute.ModifySection tSection
'    If tSection.nEndStationType = TP_CanNotSellTicket Then
'        liTemp.ListSubItems(6).Text = "������"
'        liTemp.SmallIcon = "NoSellTicket"
'    Else
'        liTemp.ListSubItems(6).Text = "����"
'        liTemp.SmallIcon = "SellTicket"
'    End If
'    lvRouteSection_ItemClick liTemp
'Exit Sub
'ErrorHandle:
'    ShowErrorMsg
'End Sub
'
