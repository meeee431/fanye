VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmBaseInfo 
   BackColor       =   &H00E0E0E0&
   Caption         =   "������Ϣ����"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   10020
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ptLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5985
      Left            =   150
      ScaleHeight     =   5985
      ScaleWidth      =   2475
      TabIndex        =   6
      Top             =   270
      Width           =   2475
      Begin MSComctlLib.TreeView tvBaseItem 
         Height          =   3915
         Left            =   30
         TabIndex        =   7
         Top             =   2490
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   6906
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   0
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "bigImgLists"
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image imgTreeTitle 
         Height          =   2700
         Left            =   -270
         Picture         =   "frmBaseInfo.frx":0000
         Top             =   60
         Width           =   3300
      End
   End
   Begin VB.PictureBox ptRight 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6405
      Left            =   2790
      ScaleHeight     =   6405
      ScaleWidth      =   6765
      TabIndex        =   0
      Top             =   -60
      Width           =   6765
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1830
         Top             =   5220
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":4CE3
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":55BD
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":5E97
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":6771
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox ptShowInfo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   990
         Left            =   -60
         ScaleHeight     =   990
         ScaleWidth      =   6615
         TabIndex        =   1
         Top             =   0
         Width           =   6615
         Begin VB.Image Image1 
            Height          =   1275
            Left            =   60
            Picture         =   "frmBaseInfo.frx":704B
            Top             =   30
            Width           =   2010
         End
         Begin VB.Label lblTitlePrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ʊ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   2430
            TabIndex        =   2
            Top             =   510
            Width           =   765
         End
         Begin VB.Image imgObject 
            Height          =   480
            Left            =   1800
            Top             =   300
            Width           =   480
         End
      End
      Begin MSComctlLib.ImageList smallImgLists 
         Left            =   2310
         Top             =   3540
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":851E
               Key             =   "luggagetype"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":88B8
               Key             =   "protocol"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":8C52
               Key             =   "vehicle"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":8FEC
               Key             =   "formula"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfo.frx":9146
               Key             =   "priceitem"
            EndProperty
         EndProperty
      End
      Begin ActiveBar2LibraryCtl.ActiveBar2 abAction 
         Height          =   4875
         Left            =   5220
         TabIndex        =   3
         Top             =   1260
         Width           =   1485
         _LayoutVersion  =   1
         _ExtentX        =   2619
         _ExtentY        =   8599
         _DataPath       =   ""
         Bands           =   "frmBaseInfo.frx":94E0
      End
      Begin MSComctlLib.ListView lvObject 
         Height          =   4515
         Left            =   570
         TabIndex        =   4
         Top             =   1470
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   7964
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "bigImgLists"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin MSComctlLib.ImageList bigImgLists 
      Left            =   1890
      Top             =   6330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfo.frx":CAC2
            Key             =   "Protocol"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfo.frx":D79C
            Key             =   "CompanySettlePrice"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfo.frx":F4A6
            Key             =   "VehicleSettlePrice"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfo.frx":111B0
            Key             =   "Formula"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfo.frx":11A8A
            Key             =   "SplitItem"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfo.frx":128DC
            Key             =   "VehicleProtocol"
         EndProperty
      EndProperty
   End
   Begin RTComctl3.Spliter spMove 
      Height          =   915
      Left            =   2370
      TabIndex        =   5
      Top             =   4230
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   1614
      BackColor       =   14737632
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
   Begin VB.Menu pmnu_Action 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu pmnu_Add 
         Caption         =   "����(&A)"
      End
      Begin VB.Menu pmnu_BaseInfo 
         Caption         =   "����(&R)"
      End
      Begin VB.Menu pmnu_Del 
         Caption         =   "ɾ��(&D)"
      End
   End
End
Attribute VB_Name = "frmBaseInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const cszProtocolID = 1
Private m_oReport As New Report
Private m_oProtocol As New Protocol
Private m_oFormula As New Formular
Dim aszTemp() As String


Private Sub abAction_BandClose(ByVal Band As ActiveBar2LibraryCtl.Band)
    If Band.name = "bndActionTabs" Then
        abAction.Visible = False
        Call ptRight_Resize
    End If
End Sub

Private Sub lvObject_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu pmnu_Action
    End If
End Sub

Private Sub pmnu_Add_Click()
    AddObject
End Sub

Private Sub pmnu_BaseInfo_Click()
    EditObject
End Sub

Private Sub pmnu_Del_Click()
    DeleteObject
End Sub

Private Sub ptRight_Resize()
On Error Resume Next
    Const cnMargin = 50
    ptShowInfo.Left = 0
    ptShowInfo.Top = 0
    ptShowInfo.Width = ptRight.ScaleWidth
    lvObject.Left = cnMargin
    lvObject.Top = ptShowInfo.Height + cnMargin
    lvObject.Width = ptRight.ScaleWidth - IIf(abAction.Visible, abAction.Width, 0) - 2 * cnMargin
    lvObject.Height = ptRight.ScaleHeight - ptShowInfo.Height - 2 * cnMargin
    '���������ر�ʱ�䴦��
    If abAction.Visible Then
        abAction.Move lvObject.Width + cnMargin, lvObject.Top
        abAction.Height = lvObject.Height
    End If
End Sub

Private Sub abAction_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "mnu_Add", "act_BaseMan_Add"
            AddObject
        Case "mnu_BaseInfo", "act_BaseMan_BaseInfo"
            EditObject
        Case "mnu_Del", "act_BaseMan_Del"
            DeleteObject
    End Select
End Sub

Public Sub AddObject()
On Error GoTo ErrHandle
    Dim szSelectKey As String
    szSelectKey = tvBaseItem.SelectedItem.Key
    Select Case szSelectKey
         Case "KProtocol"       'Э��
            DoProtocol 0
        Case "KFormula"         '��ʽ
            DoKFormula 0
        Case "KSplitItem"
            DoKSplitItem 0
    End Select
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'���ĵ�ǰѡ�е���Ŀ
Public Sub EditObject()
    On Error GoTo ErrHandle
    If lvObject.SelectedItem Is Nothing Then Exit Sub
    Select Case tvBaseItem.SelectedItem.Key
         Case "KProtocol"
            DoProtocol 1
         Case "KFormula"
            DoKFormula 1
         Case "KSplitItem"
            DoSplitItem
    End Select
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'ɾ������
Public Sub DeleteObject()
    On Error GoTo ErrHandle
    Dim oBus As Object
    Dim szSelectKey As String
    szSelectKey = tvBaseItem.SelectedItem.Key
    Select Case szSelectKey
         Case "KProtocol"
            DoProtocol 2
         Case "KFormula"
            DoKFormula 2
    End Select
    SetMenuEnabled
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'�������Э��
Private Sub DoProtocol(pnOper As Integer)
    'pnOper=0       '����
    'pnOper=1       '�޸�
    'pnOper=2       'ɾ��
    'pnOper=3       'ֻ��
    On Error GoTo err
    Select Case pnOper
    Case 0
        frmAddProtocol.Status = ST_AddObj
        frmAddProtocol.Caption = "����Э��"
        frmAddProtocol.cmdOk.Caption = "����(&A)"
        frmAddProtocol.Show vbModal
    Case 1
        frmAddProtocol.Status = ST_EditObj
        frmAddProtocol.mszProtocolID = lvObject.SelectedItem.Text
        frmAddProtocol.Caption = "�޸�Э��"
        frmAddProtocol.cmdOk.Caption = "�޸�(&E)"
        frmAddProtocol.Show vbModal
    Case 2
       ' ɾ������Э��
        Dim vbYesOrNo As Integer

        vbYesOrNo = MsgBox("�Ƿ����ɾ��" & lvObject.SelectedItem & "[" & lvObject.SelectedItem.SubItems(1) & "]", vbQuestion + vbYesNo + vbDefaultButton2, "Э�����")
        If vbYesOrNo = vbYes Then
            m_oProtocol.Init g_oActiveUser
            m_oProtocol.Identify lvObject.SelectedItem.Text
            m_oProtocol.Delete
            lvObject.ListItems.Remove lvObject.SelectedItem.Index
        End If
    End Select
    Exit Sub
err:
    ShowErrorMsg
End Sub

'�������
Private Sub DoSplitItem()
    frmSplitItem.Status = ST_EditObj
    frmSplitItem.szSplitItemID = lvObject.SelectedItem.Text
    frmSplitItem.Show vbModal
    

End Sub

Private Sub DoKSplitItem(pnOper As Integer)
    'pnOper=0       '����
    'pnOper=1       '�޸�
    'pnOper=2       'ɾ��

    Select Case pnOper
    Case 0
        frmSplitItem.Status = ST_AddObj
        frmSplitItem.Caption = "����������"
        frmSplitItem.Show vbModal
    Case 1
        frmSplitItem.Status = ST_EditObj
        frmSplitItem.Caption = "�޸ķ�����"
        frmSplitItem.Show vbModal
    Case 2
    End Select

End Sub
'����Э����㹫ʽ
Private Sub DoKFormula(pnOper As Integer)
    'pnOper=0       '����
    'pnOper=1       '�޸�
    'pnOper=2       'ɾ��

    Select Case pnOper
    Case 0
'        FrmSaveFormular.m_state = ST_AddObj

        frmFormula.m_state = ST_AddObj
        frmFormula.Show vbModal
    Case 1
'        FrmSaveFormular.m_state = ST_EditObj
        frmFormula.m_state = ST_EditObj
        frmFormula.m_szFormula = lvObject.SelectedItem.Text
        frmFormula.Show vbModal
    Case 2
        Dim vbYesOrNo As Integer
        vbYesOrNo = MsgBox("�Ƿ����ɾ��" & "[" & lvObject.SelectedItem & "]", vbQuestion + vbYesNo + vbDefaultButton2, "ɾ��������Ϣ")
        If vbYesOrNo = vbYes Then
            m_oFormula.Init g_oActiveUser
            m_oFormula.Identify lvObject.SelectedItem.Text
            m_oFormula.Delete
            lvObject.ListItems.Remove lvObject.SelectedItem.Index
        End If
    End Select

End Sub
Private Sub Form_Activate()
    spMove.LayoutIt
    WriteTitleBar "������Ϣ"
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandle
    m_oFormula.Init g_oActiveUser
    m_oProtocol.Init g_oActiveUser
    m_oReport.Init g_oActiveUser
    ptRight.Left = imgTreeTitle.Width + spMove.Width
    ptShowInfo.Left = ptRight.Left
    spMove.InitSpliter ptLeft, ptRight
    FillBaseItemTree
'    FillItemLists
 
    m_oFormula.Init g_oActiveUser
    m_oProtocol.Init g_oActiveUser
    m_oReport.Init g_oActiveUser
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'���û�����Ϣ��
Private Sub FillBaseItemTree()
    With tvBaseItem.Nodes
        .Add , , "KProtocol", "����Э�����", "Protocol"
        .Add , , "KSplitItem", "������������", "SplitItem"
        .Add , , "KFormula", "Э����㹫ʽ", "Formula"
        tvBaseItem.Nodes(1).Selected = True
        FillHead
        FillItemLists
    End With
End Sub

'����lvObject�б�ͷ
Private Sub FillHead()
'    SaveHeadWidth Me.name, tvBaseItem.SelectedItem.Key
    lvObject.ColumnHeaders.Clear
    If tvBaseItem.SelectedItem.Key = "KProtocol" Then
        With lvObject.ColumnHeaders
            
            .Add , , "Э�����", 1200
            .Add , , "Э������", 2000
            .Add , , "�Ƿ�Ĭ��Э��", 1600
            .Add , , "��ע", 1200
        End With
        
    ElseIf tvBaseItem.SelectedItem.Key = "KSplitItem" Then
        With lvObject.ColumnHeaders
            .Add , , "���������", 1200
            .Add , , "����������", 1200
            .Add , , "ʹ��״̬", 1200
            .Add , , "��������", 1600
            .Add , , "�Ƿ������޸�", 1600
        End With
    Else
        With lvObject.ColumnHeaders
            .Add , , "��ʽ����", 1200
            .Add , , "��ʽ����", 1600
            .Add , , "��ʽ����", 3000
        End With

    End If
    lvObject.SelectedItem = Nothing
'    AlignHeadWidth Me.name, tvBaseItem.SelectedItem.Key
End Sub

Private Sub tvBaseItem_NodeClick(ByVal Node As MSComctlLib.Node)
    
    FillHead
    FillItemLists
    SetMenuEnabled
End Sub

'����activebar�Ƿ�����
Private Sub SetMenuEnabled()
    Dim szSelectKey As String
    Dim bEnabled As Boolean
    szSelectKey = tvBaseItem.SelectedItem.Key
    
    If lvObject.ListItems.Count > 0 Then
        bEnabled = True
    Else
        bEnabled = False
    End If
    Select Case szSelectKey
    Case "KProtocol"     'Э�����ú͹�ʽ����
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Add").Caption = "����"
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_BaseInfo").Caption = "����"
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Del").Caption = "ɾ��"
        
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Add").Enabled = True
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_BaseInfo").Enabled = bEnabled
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Del").Enabled = bEnabled
    Case "KFormula"
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Add").Caption = "����"
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_BaseInfo").Caption = "����"
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Del").Caption = "ɾ��"
        
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Add").Enabled = True
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_BaseInfo").Enabled = bEnabled
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Del").Enabled = bEnabled
    Case "KSplitItem" '�������������,ɾ��
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Add").Caption = "����"
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_BaseInfo").Caption = "����"
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Del").Caption = "ɾ��"
        
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Add").Enabled = False
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_BaseInfo").Enabled = bEnabled
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Del").Enabled = False
    End Select
End Sub


'�г�������Ϣ
Public Sub FillItemLists(Optional szValue As String = "")
    
    On Error GoTo ErrHandle
    Dim nCount As Integer
    Dim lvItem As ListItem
    Dim atTemp() As TSplitItemInfo
    Dim i As Integer
    Dim szSelectKey As String
    Dim szSelectItem As String
    '�ֱ���ʾ��ͬ����Ϣ
    ShowSBInfo "���ڲ�ѯ�����Ե�..."
    SetBusy
    lblTitlePrompt.Caption = tvBaseItem.SelectedItem.Text
    szSelectKey = tvBaseItem.SelectedItem.Key
    If lvObject.SelectedItem Is Nothing Then
        szSelectItem = ""
    Else
        szSelectItem = lvObject.SelectedItem.Text
    End If
    If szValue = "" Then
        lvObject.ListItems.Clear
    End If
    '�õ�������Ϣ
    Select Case szSelectKey
    Case "KProtocol" '·������Э��
        aszTemp = m_oReport.GetAllProtocol(szValue)
        nCount = ArrayLength(aszTemp)
        If nCount <> 0 Then
'            ReDim Preserve aszTemp(1 To nCount, 1 To 4)
            For i = 1 To nCount
                If szSelectItem <> aszTemp(i, 1) Or lvObject.SelectedItem Is Nothing Then
                    Set lvItem = lvObject.ListItems.Add(, , aszTemp(i, 1))
                Else
                    Set lvItem = lvObject.SelectedItem
                End If
                lvItem.SmallIcon = 1
                lvItem.SubItems(1) = aszTemp(i, 2)
                lvItem.SubItems(2) = IIf(aszTemp(i, 4) = 0, GetDefaultMark(0), GetDefaultMark(1))
                lvItem.SubItems(3) = aszTemp(i, 3)
                lvItem.Selected = True
                lvItem.EnsureVisible
            Next i
        End If
    
    Case "KFormula"
        aszTemp = m_oReport.GetAllFormula(szValue)
        nCount = ArrayLength(aszTemp)
        If nCount <> 0 Then
'            ReDim Preserve aszTemp(1 To nCount)
            For i = 1 To nCount
                If szSelectItem <> aszTemp(i, 1) Or lvObject.SelectedItem Is Nothing Then
                    Set lvItem = lvObject.ListItems.Add(, , aszTemp(i, 1))
                Else
                    Set lvItem = lvObject.SelectedItem
                End If
                lvItem.SmallIcon = 2
                lvItem.SubItems(1) = aszTemp(i, 2)
                lvItem.SubItems(2) = aszTemp(i, 3)
                lvItem.Selected = True
                lvItem.EnsureVisible
            Next i
        End If
    Case "KSplitItem"
        atTemp = m_oReport.GetSplitItemInfo(szValue)
        nCount = ArrayLength(atTemp)
        If nCount <> 0 Then
'            ReDim Preserve atTemp(1 To nCount)
            For i = 1 To nCount
                If szSelectItem <> atTemp(i).SplitItemID Or lvObject.SelectedItem Is Nothing Then
                    Set lvItem = lvObject.ListItems.Add(, , atTemp(i).SplitItemID)
                Else
                    Set lvItem = lvObject.SelectedItem
                End If
                lvItem.SmallIcon = 3
                lvItem.SubItems(1) = atTemp(i).SplitItemName
                lvItem.SubItems(2) = GetSplitStatus(atTemp(i).SplitStatus)
                lvItem.SubItems(3) = GetSplitType(atTemp(i).SplitType)
                lvItem.SubItems(4) = GetAllowModify(atTemp(i).AllowModify)
                lvItem.Selected = True
                lvItem.EnsureVisible
            Next i
        End If
    
    End Select
    SetNormal
    lvObject.Refresh
    If lvObject.ListItems.Count > 0 And szValue = "" Then
        lvObject.ListItems(1).Selected = True
    
    End If
    WriteProcessBar False
    ShowSBInfo "��" & nCount & "������", ESB_ResultCountInfo
    ShowSBInfo ""
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Resize()
    spMove.LayoutIt
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '������ͷ
'    SaveHeadWidth Me.name, tvBaseItem.SelectedItem.Key
    Unload Me
End Sub

Private Sub lvobject_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvObject, ColumnHeader.Index
End Sub

Private Sub lvObject_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            If Not lvObject.SelectedItem Is Nothing Then
                DeleteObject
            End If
    End Select
End Sub

Private Sub lvObject_DblClick()
    Dim i As Integer
    Dim szSelectKey As Variant
    szSelectKey = tvBaseItem.SelectedItem.Key
    If szSelectKey = "KProtocol" Then
        DoProtocol 1
    ElseIf szSelectKey = "KSplitItem" Then
        DoSplitItem
    Else
        DoKFormula 1
    End If
  
End Sub


Private Sub ptLeft_Resize()
On Error Resume Next
    Const cnMargin = 50
    imgTreeTitle.Left = 0
    imgTreeTitle.Top = 0
    tvBaseItem.Left = imgTreeTitle.Left + cnMargin
    tvBaseItem.Top = imgTreeTitle.Top + imgTreeTitle.Height
    tvBaseItem.Width = imgTreeTitle.Width
    tvBaseItem.Height = ptLeft.ScaleHeight - imgTreeTitle.Height - cnMargin
End Sub

'�ı�LVOBJECT��Ӧ����ͷ��
Public Sub FillItemHead()
    On Error GoTo ErrorHandle
    Dim oListItem As ListItem
    Dim szSelectKey As String
    szSelectKey = tvBaseItem.SelectedItem.Key
    
    Select Case szSelectKey
    Case "KProtocol"
        lvObject.Checkboxes = False
        lvObject.MultiSelect = False
        lvObject.ColumnHeaders(cszProtocolID).Text = "Э�����"
        lvObject.ColumnHeaders(2).Text = "Э������"
        lvObject.ColumnHeaders(3).Text = "��ע"
        lvObject.ColumnHeaders(4).Text = "Ĭ��Э��"
        lvObject.ColumnHeaders(5).Width = 0
    Case "KFormula"
        lvObject.Checkboxes = False
        lvObject.MultiSelect = False
        lvObject.ColumnHeaders(1).Text = "��ʽ����"
        lvObject.ColumnHeaders(2).Text = "��ʽ����"
        lvObject.ColumnHeaders(3).Text = "��ʽ����"
        lvObject.ColumnHeaders(4).Width = 0
        lvObject.ColumnHeaders(5).Width = 0
    Case "KSplitItem"
        lvObject.Checkboxes = False
        lvObject.MultiSelect = False
        lvObject.ColumnHeaders(1).Text = "���������"
        lvObject.ColumnHeaders(2).Text = "����������"
        lvObject.ColumnHeaders(3).Text = "ʹ��״̬"
        lvObject.ColumnHeaders(4).Text = "��������"
        lvObject.ColumnHeaders(5).Text = "�Ƿ������޸�"
    End Select
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub
