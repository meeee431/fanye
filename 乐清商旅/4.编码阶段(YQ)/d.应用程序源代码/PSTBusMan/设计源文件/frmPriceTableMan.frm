VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmPriceTableMan 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��·Ʊ�۹���"
   ClientHeight    =   3405
   ClientLeft      =   2355
   ClientTop       =   2685
   ClientWidth     =   7860
   HelpContextID   =   10000430
   Icon            =   "frmPriceTableMan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin RTComctl3.CoolButton cmdReCal 
      Height          =   315
      Left            =   6435
      TabIndex        =   5
      Top             =   1290
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "�ؼ���(&R)"
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
      MICON           =   "frmPriceTableMan.frx":030A
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
      Left            =   6435
      TabIndex        =   6
      Top             =   2160
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
      MICON           =   "frmPriceTableMan.frx":0326
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
      Left            =   6435
      TabIndex        =   7
      Top             =   2565
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "����(&H)"
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
      MICON           =   "frmPriceTableMan.frx":0342
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
      Height          =   315
      Left            =   6435
      TabIndex        =   3
      Top             =   525
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
      MICON           =   "frmPriceTableMan.frx":035E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdProperty 
      Height          =   315
      Left            =   6435
      TabIndex        =   4
      Top             =   900
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "�༭(&E)"
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
      MICON           =   "frmPriceTableMan.frx":037A
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
      Height          =   315
      Left            =   6435
      TabIndex        =   2
      Top             =   150
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "����(&A)"
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
      MICON           =   "frmPriceTableMan.frx":0396
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvTable 
      Height          =   2895
      Left            =   135
      TabIndex        =   1
      Top             =   375
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Ʊ�۱����"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ʊ�۱�����"
         Object.Width           =   2294
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "��ʼִ������"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "��������"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "����޸�����"
         Object.Width           =   2469
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ƻ���������·Ʊ�۱�(&T):"
      Height          =   180
      Left            =   135
      TabIndex        =   0
      Top             =   75
      Width           =   2160
   End
End
Attribute VB_Name = "frmPriceTableMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'* Source File Name:frmPricetableman.frm
'* Project Name:PSTBusMan.vbp
'* Engineer:�·�
'* Date Generated:2002/09/03
'* Last Revision Date:2002/09/03
'* Brief Description:��·Ʊ�۱����
'* Relational Document:
'**********************************************************
Option Explicit
Private WithEvents m_oRouteTable As RoutePriceTable '
Attribute m_oRouteTable.VB_VarHelpID = -1

Private m_lRange As Long

Private Sub cmdAdd_Click()
    AddTable
End Sub


Private Sub cmdDelete_Click()
    DeleteTable
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdProperty_Click()
    EditTable
End Sub

Private Sub cmdReCal_Click()
    ReCalculate
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
'    m_bOk = False
    Set m_oRouteTable = New RoutePriceTable
    m_oRouteTable.Init g_oActiveUser
    FillLv
    FillTable
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_oRouteTable = Nothing
End Sub

Private Sub lvTable_DblClick()
    EditTable
End Sub

'Private Function ExcuteProjectTable() As String
'    '�õ�ִ�е�Ʊ�۱�
'    ExcuteProjectTable = GetProjectExcutePriceTable(g_szExePlanID)
'End Function

Private Sub m_oRouteTable_SetProgressRange(ByVal lRange As Variant)
    WriteProcessBar True, , lRange, "�����ؼ�������Ʊ�����Ժ�......"
    m_lRange = lRange
End Sub

Private Sub m_oRouteTable_SetProgressValue(ByVal lValue As Variant)
    WriteProcessBar , lValue, m_lRange
    If lValue = m_lRange Then
        WriteProcessBar False
        
    End If
End Sub

Private Sub FillLv()
    'Ʊ�۱����:585.0709         Ʊ�۱�����:1769.953         ��ʼִ������:1170.142       ��������:1080 ����޸�����:1349.858
    lvTable.ColumnHeaders.Clear
    lvTable.ColumnHeaders.Add , , "����", 585
    lvTable.ColumnHeaders.Add , , "����", 1769
    lvTable.ColumnHeaders.Add , , "ִ������", 1170
    lvTable.ColumnHeaders.Add , , "��������", 1080
    lvTable.ColumnHeaders.Add , , "�޸�����", 1349
End Sub

Private Sub AddTable()
    '����Ʊ�۱�
'    frmAddPriceTable.m_bAdd = True
    frmAddPriceTable.m_bIsParent = True
    frmAddPriceTable.m_eStatus = EFS_AddNew
    frmAddPriceTable.Show vbModal
End Sub



Private Sub DeleteTable()
    'ɾ��Ʊ�۱�
'    MsgBox "Щ�����ѱ�ע��"
    Dim oTemp As New RegularScheme
    
    On Error GoTo ErrorHandle
    oTemp.Init g_oActiveUser
    If MsgBox("�����Ҫɾ��ѡ�е���·Ʊ�۱���", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
        If Trim(lvTable.SelectedItem.Text) = Trim(oTemp.GetRunPriceTableEx(Date)) Then
            MsgBox "����ɾ��ִ�мƻ���ִ��Ʊ�ۣ�", vbInformation + vbOKOnly
            Exit Sub
        End If
        SetBusy
        m_oRouteTable.Identify lvTable.SelectedItem.Text
        m_oRouteTable.Delete
        lvTable.ListItems.Remove lvTable.SelectedItem.Index
        SetEnabled
        SetNormal
    End If
    Set oTemp = Nothing
    Exit Sub
ErrorHandle:
    ShowErrorMsg
    Set oTemp = Nothing
    SetNormal
End Sub

Private Sub EditTable()
    '�༭Ʊ�۱�
    frmAddPriceTable.m_eStatus = EFS_Modify
    frmAddPriceTable.m_bIsParent = True
    frmAddPriceTable.m_szTableID = lvTable.SelectedItem.Text
    frmAddPriceTable.Show vbModal
    
End Sub

Private Sub ReCalculate()
'    '�ؼ�������Ʊ
'    '********�˴��ؼ�������Ʊ���м��Ӧ�ÿ����Ż���.
    
    Dim szTableID As String
    Dim oParam As New SystemParam
    
    On Error GoTo ErrorHandle
    If MsgBox("���¼�������Ʊ�ǰ������ڵ�����Ʊ��������������������Ʊ��������Ҫ�����ӣ����¼�����", vbYesNo Or vbQuestion Or vbDefaultButton2) = vbYes Then
        szTableID = lvTable.SelectedItem.Text
        SetBusy
        m_oRouteTable.Identify szTableID
        WriteProcessBar True
        m_oRouteTable.ReMakeHalfPrice
        SetNormal
        ShowMsg "������������Ʊ�۳ɹ���"
        oParam.Init g_oActiveUser
        g_atTicketTypeValid = oParam.GetAllTicketType(TP_TicketTypeValid)
        g_nTicketCountValid = ArrayLength(g_atTicketTypeValid)
    End If
    Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub FillTable()
    '���Ʊ�۱�
    Dim oTicketPriceMan As New TicketPriceMan
    Dim aszRoutePriceTable() As String
    Dim i As Integer, nCount As Integer
    Dim liTemp As ListItem
    On Error GoTo ErrorHandle
    oTicketPriceMan.Init g_oActiveUser
    aszRoutePriceTable = oTicketPriceMan.GetAllRoutePriceTable()
    nCount = ArrayLength(aszRoutePriceTable)
    lvTable.ListItems.Clear
    If nCount > 0 Then
        For i = 1 To nCount
            Set liTemp = lvTable.ListItems.Add(, GetEncodedKey(aszRoutePriceTable(i, 1)), aszRoutePriceTable(i, 1))
            liTemp.ListSubItems.Add , , aszRoutePriceTable(i, 2)
            liTemp.ListSubItems.Add , , Format(aszRoutePriceTable(i, 3), "YYYY-MM-DD") '��ʼִ��ʱ��
            liTemp.ListSubItems.Add , , Format(aszRoutePriceTable(i, 4), "YYYY-MM-DD")
            liTemp.ListSubItems.Add , , Format(aszRoutePriceTable(i, 5), "YYYY-MM-DD")
        Next i
        lvTable.ListItems(1).Selected = True
    End If
    SetEnabled
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Public Sub AddList(pszID As String)
    '��������Ʊ�۱�ˢ�³���
    Dim liTemp As ListItem
    m_oRouteTable.Identify pszID
    Set liTemp = lvTable.ListItems.Add(, GetEncodedKey(m_oRouteTable.RoutePriceTableID), m_oRouteTable.RoutePriceTableID)
    liTemp.SubItems(1) = m_oRouteTable.RoutePriceTableName
    liTemp.SubItems(2) = Format(m_oRouteTable.StartRunTime, "YYYY-MM-DD")
    liTemp.SubItems(3) = Format(m_oRouteTable.CreateTime, "YYYY-MM-DD")
    liTemp.SubItems(4) = Format(m_oRouteTable.LastModifyTime, "YYYY-MM-DD")
    liTemp.EnsureVisible
End Sub

Public Sub UpdateList(pszID As String)
    'ˢ���޸ĵ�Ʊ�۱�
    Dim liTemp As ListItem
    If lvTable.SelectedItem Is Nothing Then Exit Sub
    m_oRouteTable.Identify pszID
    Set liTemp = lvTable.SelectedItem
    liTemp.SubItems(1) = m_oRouteTable.RoutePriceTableName
    liTemp.SubItems(2) = Format(m_oRouteTable.StartRunTime, "YYYY-MM-DD")
    liTemp.SubItems(3) = Format(m_oRouteTable.CreateTime, "YYYY-MM-DD")
    liTemp.SubItems(4) = Format(m_oRouteTable.LastModifyTime, "YYYY-MM-DD")
    liTemp.EnsureVisible
End Sub

Private Sub SetEnabled()
    Dim bEnabled As Boolean
    If lvTable.ListItems.Count > 0 Then
        bEnabled = True
    Else
        bEnabled = False
    End If
    cmdDelete.Enabled = bEnabled
    cmdProperty.Enabled = bEnabled
    cmdReCal.Enabled = bEnabled
    
End Sub
