VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStoreMenu 
   Caption         =   "Form1"
   ClientHeight    =   3945
   ClientLeft      =   2760
   ClientTop       =   3405
   ClientWidth     =   7860
   HelpContextID   =   5000001
   Icon            =   "frmStoreMenu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3945
   ScaleWidth      =   7860
   Visible         =   0   'False
   Begin MSComctlLib.ImageList ilSysMan 
      Left            =   4380
      Top             =   1770
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreMenu.frx":000C
            Key             =   "sysman"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreMenu.frx":03A6
            Key             =   "localunit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreMenu.frx":0502
            Key             =   "usergroup"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreMenu.frx":089C
            Key             =   "user"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreMenu.frx":0C36
            Key             =   "actuser"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreMenu.frx":0FD0
            Key             =   "logman"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreMenu.frx":1138
            Key             =   "unitman"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreMenu.frx":1294
            Key             =   "component"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreMenu.frx":13F0
            Key             =   "funright"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreMenu.frx":154C
            Key             =   "stationman"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_System 
      Caption         =   "ϵͳ(&S)"
      Begin VB.Menu mnuCPassWord 
         Caption         =   "��ǰ�û�����(&U)"
      End
      Begin VB.Menu mnuModiPara 
         Caption         =   "����ϵͳ����(&R)"
      End
      Begin VB.Menu mnuSetTicketType 
         Caption         =   "Ʊ������(&I)"
      End
      Begin VB.Menu mnuWriteOffTicket 
         Caption         =   "ע����Ʊ(&W)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_Line1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_ExprotFile 
         Caption         =   "�����ļ�(&F)..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_ExpOpen 
         Caption         =   "�����ļ�����(&T)..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_Line2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_PrintEX 
         Caption         =   "��ӡ(&P)"
         Enabled         =   0   'False
         Shortcut        =   ^P
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_PrintView 
         Caption         =   "��ӡԤ��(&V)"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_Line3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_PageSet 
         Caption         =   "ҳ������(&G)"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_PrintSet 
         Caption         =   "��ӡ����(&S)"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Tree 
         Caption         =   "������(&K)"
         Checked         =   -1  'True
         Shortcut        =   ^T
      End
      Begin VB.Menu mnu_StatusBar 
         Caption         =   "״̬��(&Z)"
         Checked         =   -1  'True
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_View 
         Caption         =   "��Ϣ��ͼ(&V)"
         Begin VB.Menu mnu_Icon 
            Caption         =   "��ͼ��(&I)"
         End
         Begin VB.Menu mnu_SmallIcon 
            Caption         =   "Сͼ��(&M)"
         End
         Begin VB.Menu mnu_List 
            Caption         =   "�б�(&L)"
         End
         Begin VB.Menu mnu_Detail 
            Caption         =   "��ϸ���(&D)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�˳�(&X)"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnu_Action 
      Caption         =   "����(&T)"
      Begin VB.Menu mnu_SubAction 
         Caption         =   "����(&P)"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "����(&H)"
      Begin VB.Menu mnu_HelpContent 
         Caption         =   "����(&C)"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnu_HelpIndex 
         Caption         =   "����(&I)"
      End
      Begin VB.Menu mnuBreak12 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_About 
         Caption         =   "���� ϵͳ����(&A)..."
      End
   End
   Begin VB.Menu pmnuUserGroup 
      Caption         =   "pmnu�û���"
      Visible         =   0   'False
      Begin VB.Menu pmnuGroupProperty 
         Caption         =   "�û�������"
      End
      Begin VB.Menu pmnuDelectGroup 
         Caption         =   "ɾ��ѡ�е��û���"
      End
      Begin VB.Menu pmnuAddGroup 
         Caption         =   "�����û���"
      End
   End
   Begin VB.Menu pmnuUser 
      Caption         =   "pmnu�û�"
      Visible         =   0   'False
      Begin VB.Menu pmnuUserProperty 
         Caption         =   "�û�����"
      End
      Begin VB.Menu pmnuDelectUser 
         Caption         =   "ɾ��ѡ�е��û�"
      End
      Begin VB.Menu pmnuAddUser 
         Caption         =   "�����û�"
      End
   End
   Begin VB.Menu pmnuLogMan 
      Caption         =   "pmnu��־����"
      Visible         =   0   'False
      Begin VB.Menu pmnuSetAutoDelectLog 
         Caption         =   "��־����"
      End
   End
   Begin VB.Menu pmnuLoginLog 
      Caption         =   "pmnu��¼��־"
      Visible         =   0   'False
      Begin VB.Menu pmnuSelectLoginLog 
         Caption         =   "��ѯ"
      End
      Begin VB.Menu pmnuDeleteLoginLog 
         Caption         =   "ɾ��"
      End
   End
   Begin VB.Menu pmnuOpeLog 
      Caption         =   "pmnu������־"
      Visible         =   0   'False
      Begin VB.Menu pmnuSelectOpeLog 
         Caption         =   "��ѯ"
      End
      Begin VB.Menu pmnuDeleteOpeLog 
         Caption         =   "ɾ��"
      End
   End
   Begin VB.Menu pmnuStation 
      Caption         =   "pmnu��վ����"
      Visible         =   0   'False
      Begin VB.Menu pmnuStionProperty 
         Caption         =   "����"
      End
      Begin VB.Menu pmnuAddStation 
         Caption         =   "������վ"
      End
   End
   Begin VB.Menu pmnuUnit 
      Caption         =   "pmnu��λ����"
      Visible         =   0   'False
      Begin VB.Menu pmnuUnitProperty 
         Caption         =   "����"
      End
      Begin VB.Menu pmnuAddUnit 
         Caption         =   "������λ"
      End
      Begin VB.Menu pmnuDeleteUnit 
         Caption         =   "ɾ����λ"
      End
      Begin VB.Menu pmnuRecoverUnit 
         Caption         =   "�ָ���ɾ��λ"
      End
   End
   Begin VB.Menu pmnuActUser 
      Caption         =   "pmnu��û�����"
      Visible         =   0   'False
      Begin VB.Menu pmnuRefresh 
         Caption         =   "ˢ��"
      End
      Begin VB.Menu pmnuLogout 
         Caption         =   "ǿ��ע��"
      End
   End
   Begin VB.Menu pmnuFunction 
      Caption         =   "pmnu���ܹ���"
      Visible         =   0   'False
      Begin VB.Menu pmnuGrantFun 
         Caption         =   "��������Ȩ"
      End
   End
   Begin VB.Menu pmnuFunGroup 
      Caption         =   "pmnu���������"
      Visible         =   0   'False
      Begin VB.Menu pmnuGrantFunGroup 
         Caption         =   "����������Ȩ"
      End
   End
End
Attribute VB_Name = "frmStoreMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *******************************************************************
' *  Source File Name  : frmStoreMenu                               *
' *  Project Name: PSTSMan                                          *
' *  Engineer:                                                      *
' *  Date Generated: 2002/08/19                                     *
' *  Last Revision Date : 2002/08/19                                *
' *  Brief Description   : ������ʵ�ʿ��ƴ���(����ʾ)               *
' *******************************************************************
'===================================================
'Modify Date��2002-11-13
'Author:½����
'Reamrk:��ӶԳ�վ�Ĵ���
'===================================================

Option Explicit
Option Base 1
Const cnStartHID = 5000000
Const cnlvDetailHID = 11
Const cnlvDetail2HID = 15
Private Enum ECurTask
    ERootHID = 0
    EUserAndGroupHID = 10
    ELogHID = 20
    EUnitHID = 30
    EFunctionAndComHID = 40
    EActUserHID = 50
End Enum

Public m_frmMain As frmSMCMain

Public WithEvents tvAll As TreeView
Attribute tvAll.VB_VarHelpID = -1
Public WithEvents lvDetail As ListView
Attribute lvDetail.VB_VarHelpID = -1
Public WithEvents lvDetail2 As ListView
Attribute lvDetail2.VB_VarHelpID = -1

'����frmSMCMain
Dim lHwnd1 As Long, lhwnd2 As Long
Dim lx As Long, ly As Long, lw As Long, lh As Long
'��������
'Dim oUser As New User
'Dim oUserGroup As New UserGroup
'Dim oUnit As New Unit
'�������¶���
Dim aTactUserInfo() As TActiveUserInfo
Dim aTAllCOMInfo() As TCOMSelfInfo1
Dim rsLoginInfo As Recordset
Dim rsOpeInfo As Recordset
Dim CellExpSourceName As Object

Public Function LoadMenuForm(pfrmMain As Object) As Long
    Set m_frmMain = pfrmMain
    Set tvAll = m_frmMain.tvAll
    Set lvDetail = m_frmMain.lvDetail
    Set lvDetail2 = m_frmMain.lvDetail2
    Set lvDetail.SmallIcons = ilSysMan
    Set lvDetail2.SmallIcons = ilSysMan
    Set lvDetail.Icons = ilSysMan
    Set lvDetail2.Icons = ilSysMan
    lvDetail.Arrange = lvwAutoLeft
    lvDetail2.Arrange = lvwAutoLeft
    Set m_frmActiveSMC = Me
End Function

Private Sub Form_Load()
'===================================================
'Modify Date��2002-11-13
'Author:½����
'Reamrk:���StationMan
'===================================================    Dim tvTemp As TreeView
    Dim i As Integer, j As Integer
    Dim tvTemp As TreeView
    lHwnd1 = Me.hwnd
    lhwnd2 = frmSMCMain.hwnd
    Me.Top = (Screen.Height - Me.ScaleHeight) / 2
    Me.Left = (Screen.Width - Me.ScaleWidth) / 2
    Set tvTemp = frmSMCMain.tvAll
    Set tvTemp.ImageList = ilSysMan
    tvTemp.Nodes.Clear
    tvTemp.Nodes.Add , , cszConRoot, "ϵͳ����", "sysman"
    tvTemp.Nodes.Add cszConRoot, tvwChild, cszUserGroupMan, "�û����û������", "user"
    tvTemp.Nodes.Add cszConRoot, tvwChild, cszActiveUserMan, "��û�����", "actuser"
    tvTemp.Nodes.Add cszConRoot, tvwChild, cszLogMan, "��־����", "logman"
    tvTemp.Nodes.Add cszLogMan, tvwChild, cszLoginLogMan, "��¼��־", "logman"
    tvTemp.Nodes.Add cszLogMan, tvwChild, cszOperateLogMan, "������־", "logman"
'    tvTemp.Nodes.Add cszConRoot, tvwChild, cszFunctionMan, "���ܼ��������", "funright"
'    tvTemp.Nodes.Add cszFunctionMan, tvwChild, cszFun_GroupMan, "���ܼ����������", "funright"
'    tvTemp.Nodes.Add cszFunctionMan, tvwChild, cszComponent, "�������", "component"
    tvTemp.Nodes.Add cszConRoot, tvwChild, cszUnitMan, "��λ����", "unitman"
    tvTemp.Nodes.Add cszConRoot, tvwChild, cszStationMan, "��վ����", "stationman"
    tvTemp.Nodes.Item(1).Expanded = True
    LoadCommonData
    
End Sub




Private Sub lvDetail_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If g_szCurrentTask = cszActiveUserMan Then
        LoadActUserInfo
    End If
End Sub

Private Sub lvDetail_DblClick()
    Select Case g_szCurrentTask
    Case cszConRoot
    Case cszUserGroupMan
        Call pmnuUserProperty_Click
    Case cszLogMan
        Call pmnuSetAutoDelectLog_Click
    Case cszLoginLogMan
    Case cszOperateLogMan
    Case cszUnitMan
        Call pmnuUnitProperty_Click
    Case cszFunctionMan
    Case cszActiveUserMan
        pmnuRefresh_Click
    Case cszStationMan
        Call pmnuStionProperty_Click
    Case Else
    End Select

End Sub

Private Sub lvDetail_GotFocus()
    Select Case g_szCurrentTask
        Case cszConRoot
            EnableMnunOfCellExport (False)
        Case cszLogMan
            EnableMnunOfCellExport (False)
        Case cszFunctionMan
            EnableMnunOfCellExport (False)
        Case Else
            EnableMnunOfCellExport (True)
            Set CellExpSourceName = lvDetail
    End Select


End Sub

Private Sub lvDetail_ItemClick(ByVal Item As MSComctlLib.ListItem) '�õ�ѡ�е�Item��Text(��"����ID")
    Dim oTemp As ListItems, i As Integer, nTemp As Integer, nLen As Integer
    Set oTemp = lvDetail.ListItems
    ReDim g_alvItemText(1)
    nTemp = oTemp.Count
    If nTemp > 0 Then
        nLen = 0
        For i = 1 To nTemp
            If oTemp(i).Selected Then
                nLen = nLen + 1
                ReDim Preserve g_alvItemText(1 To nLen)
                g_alvItemText(nLen) = oTemp(i).Text
            End If
        Next i
    End If
    
    If g_szCurrentTask = cszActiveUserMan Then
        mnu_SubAction(2).Enabled = True
    End If
    
End Sub


Private Sub lvDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton Then
        Select Case g_szCurrentTask
            Case cszConRoot
                '
            Case cszUserGroupMan
                PopupMenu pmnuUser ', , , , pmnuUserProperty
            Case cszLogMan
                PopupMenu pmnuLogMan ', , , , pmnuSetAutoDelectLog
            Case cszLoginLogMan
                PopupMenu pmnuLoginLog ', , , , pmnuSelectLoginLog
            Case cszOperateLogMan
                PopupMenu pmnuOpeLog ', , , , pmnuSelectOpeLog
            Case cszUnitMan
                PopupMenu pmnuUnit ', , , , pmnuUnitProperty
            Case cszFunctionMan
            Case cszFun_GroupMan
                PopupMenu pmnuFunction
            Case cszActiveUserMan
                PopupMenu pmnuActUser ', , , , pmnuActUser
            Case cszStationMan
                PopupMenu pmnuStation  ', , , , pmnuStation
            Case Else
            ''''
        End Select
    End If
End Sub


Private Sub lvDetail2_DblClick()
        Select Case g_szCurrentTask
            Case cszUserGroupMan
                Call pmnuGroupProperty_Click
            Case cszFunctionMan
'                PopupMenu pmnuCOM_Function, , , , pmnuLoadCOM
            Case Else
            ''''
        End Select

End Sub

Private Sub lvDetail2_GotFocus()
    Select Case g_szCurrentTask
        Case cszUserGroupMan
            EnableMnunOfCellExport (True)
            Set CellExpSourceName = lvDetail2
        Case cszFun_GroupMan
            EnableMnunOfCellExport (True)
            Set CellExpSourceName = lvDetail2
        Case Else
            EnableMnunOfCellExport (False)
    End Select
End Sub

Private Sub lvDetail2_ItemClick(ByVal Item As MSComctlLib.ListItem) '�õ�ѡ�е�Item��Text(��"����ID")
    Dim oTemp As ListItems, i As Integer, nTemp As Integer, nLen As Integer
    Set oTemp = lvDetail2.ListItems
    ReDim g_alvItemText2(1)
    nTemp = oTemp.Count
    If nTemp > 0 Then
        nLen = 0
        For i = 1 To nTemp
            If oTemp(i).Selected Then
                nLen = nLen + 1
                ReDim Preserve g_alvItemText2(1 To nLen)
                g_alvItemText2(nLen) = oTemp(i).Text
            End If
        Next i
    End If

End Sub


Private Sub lvDetail2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Select Case g_szCurrentTask
            Case cszUserGroupMan
                PopupMenu pmnuUserGroup ', , , , pmnuGroupProperty
            Case cszFunctionMan
'                PopupMenu pmnuCOM_Function, , , , pmnuLoadCOM
            Case cszFun_GroupMan
                PopupMenu pmnuFunGroup
            Case Else
            ''''
        End Select
    End If
End Sub


Private Sub mnu_About_Click()
    Dim picTemp As StdPicture
    Set picTemp = ilSysMan.ListImages(1).Picture
    Call g_oLogin.ShowAbout("ϵͳ���� ", "System Manage", "ϵͳ����", frmSMCMain.Icon, App.Major, App.Minor, App.Revision)
End Sub

Private Sub mnu_Detail_Click()
    ListViewShow lvwReport
End Sub



Public Sub QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
End Sub

Private Sub mnu_ExpOpen_Click()
    InitCellExport CellExpSourceName, 3
'    CellExport.ExportFile (True)

End Sub

Private Sub mnu_ExprotFile_Click()
    InitCellExport CellExpSourceName, 3
'    CellExport.ExportFile
End Sub

Private Sub mnu_HelpContent_Click()
    DisplayHelp Me, content
End Sub

Private Sub mnu_HelpIndex_Click()
    DisplayHelp Me, Index
End Sub

Private Sub mnu_Icon_Click()
    ListViewShow lvwIcon
End Sub

Private Sub mnu_List_Click()
    ListViewShow lvwList
End Sub


Public Sub ActiveUserMan()
    Dim lvtemp As ListView
    
    Dim ltTemp As ListItem
    Dim vaTemp As Variant, nCount As Integer, i As Integer
    
    Set lvtemp = m_frmMain.lvDetail
    lvtemp.ColumnHeaders.Clear


    lvtemp.ColumnHeaders.Add , "��¼����", "��¼����"
    lvtemp.ColumnHeaders.Add , "��¼�û�����", "��¼�û�����"
    lvtemp.ColumnHeaders.Add , "��¼�û���", "��¼�û���"
    lvtemp.ColumnHeaders.Add , "��¼����վ", "��¼����վ"
    lvtemp.ColumnHeaders.Add , "��¼ʱ��", "��¼ʱ��"
    lvtemp.ColumnHeaders.Add , "���ʱ��", "���ʱ��"
    lvtemp.ListItems.Clear
    ShowHowDetail
    
    ClearActionMenu
    mnu_SubAction(MProperty).Enabled = False
    AddSubAction "ˢ��(&R)"
    AddSubAction "ǿ��ע��(&L)"
    mnu_SubAction(1).Enabled = True
    mnu_Action.Visible = True
    LoadActUserInfo
End Sub

Public Sub UserGroupMan()
    Dim lvtemp As ListView
    Set lvtemp = m_frmMain.lvDetail
    lvtemp.ListItems.Clear
    lvtemp.ColumnHeaders.Clear

    lvtemp.ColumnHeaders.Add , "�û�����", "�û�����"
    lvtemp.ColumnHeaders.Add , "�û���", "�û���"
    lvtemp.ColumnHeaders.Add , "�Ƿ������û�", "�Ƿ������û�"
    lvtemp.ColumnHeaders.Add , "�û�������λ", "�û�������λ"
    
    Set lvtemp = m_frmMain.lvDetail2
    lvtemp.ListItems.Clear
    lvtemp.ColumnHeaders.Clear
    lvtemp.ColumnHeaders.Add , "�û������", "�û������"
    lvtemp.ColumnHeaders.Add , "�û�����", "�û�����"
    lvtemp.ColumnHeaders.Add , "�Ƿ�������", "�Ƿ�������"
    lvtemp.ColumnHeaders.Add , "ע��", "ע��"
    
    ShowHowDetail False
    
    ClearActionMenu
    mnu_SubAction(MProperty).Caption = "�û�����(&U)"
    AddSubAction "�û�������(&G)"

    AddSubAction "ɾ��ѡ�е��û�(&D)"
    AddSubAction "ɾ��ѡ�е��û���(&S)"
    AddSubAction "�����û�(&A)"
    AddSubAction "�����û���(&N)"
    mnu_Action.Visible = True
    LoadUser_Group '�����û����û������Ϣ
    
End Sub

Public Sub LoginLogMan()
    Dim dtStart As Date
    Dim dtEnd As Date
    Dim tmStart As Date
    Dim tmEnd As Date
    Dim g_aszUser() As String
    Dim aszWorkStation() As String
    
    
    Dim lvtemp As ListView
    Set lvtemp = m_frmMain.lvDetail
    lvtemp.ColumnHeaders.Clear

    lvtemp.ColumnHeaders.Add , "�¼�����", "�¼�����"
    lvtemp.ColumnHeaders.Add , "��¼ʱ��", "��¼ʱ��"
    lvtemp.ColumnHeaders.Add , "�û�����", "�û�����"
    lvtemp.ColumnHeaders.Add , "�������", "�������"
    lvtemp.ColumnHeaders.Add , "IP��ַ", "IP��ַ"
    lvtemp.ColumnHeaders.Add , "�˳�ʱ��", "�˳�ʱ��"
    lvtemp.ListItems.Clear
    
    ShowHowDetail
    
    ClearActionMenu
    
    AddSubAction "��ѯ(&Q)"
    AddSubAction "ɾ��(&D)"
    mnu_SubAction(MProperty).Enabled = False
    mnu_Action.Visible = True
    
    dtStart = Date
    dtEnd = Date
    tmStart = "00:00:00"
    tmEnd = "23:59:59"
    
    OpenDefLoginLog g_aszUser, dtStart, dtEnd, tmStart, tmEnd, aszWorkStation
    
    
End Sub

Public Sub OperateLogMan()
    Dim dtStart As Date
    Dim dtEnd As Date
    Dim tmStart As Date
    Dim tmEnd As Date
    Dim g_aszUser() As String
    Dim aszFunOrGroup() As String


    Dim lvtemp As ListView
    Set lvtemp = m_frmMain.lvDetail
    lvtemp.ListItems.Clear
    lvtemp.ColumnHeaders.Clear

    lvtemp.ColumnHeaders.Add , "�¼�����", "�¼�����"
    lvtemp.ColumnHeaders.Add , "�û�����", "�û�����"
    lvtemp.ColumnHeaders.Add , "ע��", "ע��"
    lvtemp.ColumnHeaders.Add , "����ʱ��", "����ʱ��"
    lvtemp.ColumnHeaders.Add , "���������", "���������"
    lvtemp.ColumnHeaders.Add , "���ܴ���", "���ܴ���"
    ShowHowDetail
    ClearActionMenu
    AddSubAction "��ѯ(&Q)"
    AddSubAction "ɾ��(&D)"
    mnu_SubAction(MProperty).Enabled = False
    mnu_Action.Visible = True
    
    dtStart = Date
    dtEnd = Date
    tmStart = "00:00:01"
    tmEnd = "23:59:59"
    OpenDefOpeLog g_aszUser, dtStart, dtEnd, tmStart, tmEnd, aszFunOrGroup
End Sub

Public Sub UnitMan()
    Dim lvtemp As ListView
    Set lvtemp = m_frmMain.lvDetail
    lvtemp.ListItems.Clear
    lvtemp.ColumnHeaders.Clear
    
    lvtemp.ColumnHeaders.Add , "��λ����", "��λ����"
    lvtemp.ColumnHeaders.Add , "��λ���", "��λ���"
    lvtemp.ColumnHeaders.Add , "��λȫ��", "��λȫ��"
    lvtemp.ColumnHeaders.Add , "��������", "��������"
    lvtemp.ColumnHeaders.Add , "IP��ַ", "IP��ַ"
    lvtemp.ColumnHeaders.Add , "ע��", "ע��"
    
    ShowHowDetail
    ClearActionMenu
    AddSubAction "������λ(&A)"
    AddSubAction "ɾ����λ(&D)"
    AddSubAction "�ָ���ɾ���ĵ�λ(&R)"
    
    mnu_Action.Visible = True
    LoadUnitInfo
    
End Sub
Public Sub StationMan()
    Dim lvtemp As ListView
    Set lvtemp = m_frmMain.lvDetail
    lvtemp.ListItems.Clear
    lvtemp.ColumnHeaders.Clear
    
    lvtemp.ColumnHeaders.Add , "��վ����", "��վ����"
    lvtemp.ColumnHeaders.Add , "��վ���", "��վ���"
    lvtemp.ColumnHeaders.Add , "��վȫ��", "��վȫ��"
    lvtemp.ColumnHeaders.Add , "������λ", "������λ"
    lvtemp.ColumnHeaders.Add , "ע��", "ע��"
    
    ShowHowDetail
    ClearActionMenu
    AddSubAction "������λ(&A)"
    
    mnu_Action.Visible = True
    LoadStationInfo
    
End Sub

Public Sub Fun_GroupMan()
    Dim lvtemp As ListView
    Set lvtemp = m_frmMain.lvDetail
    lvtemp.ColumnHeaders.Clear
    lvtemp.ListItems.Clear
    
    lvtemp.ColumnHeaders.Add , "���ܴ���", "���ܴ���"
    lvtemp.ColumnHeaders.Add , "������", "������"
    lvtemp.ColumnHeaders.Add , "�������", "�������"
    lvtemp.ColumnHeaders.Add , "������", "������"
    lvtemp.ColumnHeaders.Add , "ע��", "ע��"
    Set lvtemp = m_frmMain.lvDetail2
    lvtemp.ListItems.Clear
    lvtemp.ColumnHeaders.Clear
    lvtemp.ColumnHeaders.Add , "��������", "��������"
    lvtemp.ColumnHeaders.Add , "�������", "�������"
    lvtemp.ColumnHeaders.Add , "�����Ĺ���", "�����Ĺ���"
    ShowHowDetail False   '��ʾlvDetail2
    ClearActionMenu 'reset Action�˵���submnun
    mnu_SubAction(MProperty).Caption = "����"
    AddSubAction "��������Ȩ(&F)"
    AddSubAction "����������Ȩ(&G)"
    mnu_Action.Visible = True
    mnu_SubAction(MProperty).Enabled = False
    LoadCommonData
'    LoadFun_GroupInfo
End Sub
Public Sub ShowRoot()
    Dim lvtemp As ListView
    Set lvtemp = m_frmMain.lvDetail
    lvtemp.ListItems.Clear
    lvtemp.ColumnHeaders.Clear
    ShowHowDetail
    ClearActionMenu
    mnu_SubAction(MProperty).Enabled = False

End Sub

Public Sub ShowHowDetail(Optional pbOne As Boolean = True)
    m_frmMain.lvDetail2.Visible = Not pbOne
    
    m_frmMain.splDetail.LayoutIt
End Sub

Private Sub mnu_PageSet_Click()
    InitCellExport CellExpSourceName, 3
'    CellExport.PageSetup
End Sub

Private Sub mnu_PrintEX_Click()
    InitCellExport CellExpSourceName, 3
'    CellExport.PrintEx (True)
End Sub

Private Sub mnu_PrintView_Click()
    InitCellExport CellExpSourceName, 3
'    CellExport.PrintPreview
End Sub

Private Sub mnu_SmallIcon_Click()
    ListViewShow lvwSmallIcon
End Sub

Private Sub mnu_StatusBar_Click()
    mnu_StatusBar.Checked = Not mnu_StatusBar.Checked
    m_frmMain.sbStatus.Visible = mnu_StatusBar.Checked
    m_frmMain.LayoutForm
End Sub


Private Sub mnu_Tree_Click()
    mnu_Tree.Checked = Not mnu_Tree.Checked
    m_frmMain.ptLeft.Visible = mnu_Tree.Checked
    m_frmMain.Spliter1.LayoutIt
    
End Sub

Private Sub ListViewShow(pnViewStyle As ListViewConstants)
    If m_frmMain.lvDetail.Visible Then m_frmMain.lvDetail.View = pnViewStyle
    If m_frmMain.lvDetail2.Visible Then m_frmMain.lvDetail2.View = pnViewStyle
    mnu_Icon.Checked = False
    mnu_SmallIcon.Checked = False
    mnu_Detail.Checked = False
    mnu_List.Checked = False
    Select Case pnViewStyle
        Case lvwIcon
        mnu_Icon.Checked = True
        
        Case lvwSmallIcon
        mnu_SmallIcon.Checked = True
        
        Case lvwList
        mnu_List.Checked = True
        
        Case lvwReport
        mnu_Detail.Checked = True
    End Select
End Sub

Private Sub ClearActionMenu()  'reset Action�˵���submnun
    Dim i As Integer, nCount As Integer
    nCount = mnu_SubAction.Count
    For i = 1 To nCount - 1
        Unload mnu_SubAction(i)
    Next
    mnu_SubAction(MProperty).Caption = "����"
    mnu_SubAction(MProperty).Enabled = True
    mnu_Action.Visible = False
End Sub

Private Function AddSubAction(pszCaption As String) As Integer
    Dim nTemp As Integer
    nTemp = mnu_SubAction.Count
    
    Load mnu_SubAction(nTemp)
    mnu_SubAction(nTemp).Caption = pszCaption
    mnu_SubAction(nTemp).Visible = True
    AddSubAction = nTemp
End Function

Private Sub mnu_SubAction_Click(Index As Integer)
    Select Case g_szCurrentTask
        Case cszUserGroupMan
        DoActUserGroup Index
        
        Case cszLogMan
        DoActLogMan Index
        
        Case cszUnitMan
        DoActUnit Index
        
        Case cszSystemMan
        DoActStation Index
        
        Case cszFunctionMan
'        DoActFunction Index
    
        Case cszLoginLogMan
        DoActLoginLog Index
        
        Case cszOperateLogMan
        DoActOperateLog Index
        
        Case cszActiveUserMan
        DoActActiveUser Index
        
        Case cszComponent
'        DoActComponent Index
        
        Case cszFun_GroupMan
        DoActFun_Group Index
        
    End Select
    
End Sub


Public Sub DoActUserGroup(pnIndex As Integer)
    Select Case pnIndex
        Case MProperty
            pmnuUserProperty_Click
        Case mPropertyGroup
            pmnuGroupProperty_Click
        Case MDelectUse
            pmnuDelectUser_Click
        Case mDelectGroup
            pmnuDelectGroup_Click
        Case MAddUser
            pmnuAddUser_Click
        Case MAddGroup
            pmnuAddGroup_Click
'        Case MRecoverUser
'            pmnuRecoverUser_Click
    
    End Select
End Sub

Public Sub DoActUnit(pnIndex As Integer)
    Select Case pnIndex
        Case MProperty
        pmnuUnitProperty_Click
    
        Case MAddUnit
        pmnuAddUnit_Click
        Case MRecoverUnit
        pmnuRecoverUnit_Click
        Case Else
        pmnuDeleteUnit_Click
    
    End Select
    
End Sub
Public Sub DoActStation(pnIndex As Integer)
    Select Case pnIndex
        Case MProperty
        pmnuStionProperty_Click
    
        Case MAddUnit
        pmnuAddStation_Click
    
    End Select
    
End Sub
'Public Sub DoActFunction(pnIndex As Integer)
''    Dim nTemp As Integer
''    Select Case pnIndex
''        Case Mproperty
''            frmEditFunction.Show vbModal, m_frmMain
''
''        Case MAddCOM
'            frmAddCOM.Show vbModal, m_frmMain
''        Case MDeleteCOM
''            nTemp = MsgBox("ȷ��ɾ�����", vbYesNo + vbExclamation, cszMsg)
''            If nTemp = vbYes Then
''                LV2_Delect
''            End If
''        Case Else
''            nTemp = MsgBox("ȷ��ɾ������", vbYesNo + vbExclamation, cszMsg)
''            If nTemp = vbYes Then
''                LV_Delect
''            End If
''        ''''
''    End Select
'End Sub

Public Sub DoActOperateLog(pnIndex As Integer)
Dim ndTemp As Node
    Select Case pnIndex
        Case MProperty
        
        Case MDeleteSel
            pmnuDeleteOpeLog_Click
        Case MSelect
            pmnuSelectOpeLog_Click
        Case Else
        ''''
    End Select
End Sub

Public Sub DoActLoginLog(pnIndex As Integer)
Dim ndTemp As Node
    Select Case pnIndex
        Case MProperty
        
        Case MDeleteSel
            pmnuDeleteLoginLog_Click
        Case MSelect
            pmnuSelectLoginLog_Click
        Case Else
        'Do Nothing
    End Select
End Sub



Private Sub LV_Delect()
    Dim oUnit As New Unit
    Dim oUser As New User
    
    Dim oFunction As New COMFunction
    Dim oTemp As ListItems, nLen As Integer, i As Integer, j As Integer
    Set oTemp = lvDetail.ListItems
    nLen = ArrayLength(g_alvItemText)
    If nLen > 0 Then
    If g_alvItemText(1) <> "" Then
        
            
            
                
                
                Select Case g_szCurrentTask
                    Case cszUserGroupMan
                        '***************����User�����DElete����
                        For i = 1 To nLen
                        For j = 1 To ArrayLength(g_atUserInfo)
                            If g_atUserInfo(j).UserID = g_alvItemText(i) Then
                                If g_atUserInfo(j).InnerUser = False And g_atUserInfo(j).UserID <> g_oActUser.UserID Then
                                    On Error GoTo ErrorHandle
                                    oUser.Init g_oActUser
                                    oUser.Identify g_alvItemText(i)
                                    oUser.Delete
                                    oTemp.Remove ("A" & g_alvItemText(i))
                                Else
                                    MsgBox "����ɾ�������û����Լ�!", vbInformation, cszMsg
                                End If
                            End If
                        Next j
                        Next i

                    Case cszActiveUserMan
                        '*********************
                        For i = 1 To nLen
                        On Error GoTo ErrorHandle
                        g_oSysMan.ForceLogout g_alvItemText(i)
                        
                        
                        oTemp.Remove ("A" & g_alvItemText(i))
                        
                        Next i

'                        bNewOfActUser = False
                    Case cszFunctionMan
                        '***************
                        For i = 1 To nLen
                        On Error GoTo ErrorHandle
                        oFunction.Init g_oActUser
                        oFunction.Identify g_alvItemText(i)
                        oFunction.Delete
                        
                        
                        oTemp.Remove ("A" & g_alvItemText(i))
                        
                        Next i
                        
                    Case cszLoginLogMan
                        '****************
                        For i = 1 To nLen
                        On Error GoTo ErrorHandle
                        g_oSysMan.DeleteLoginLog CLng(g_alvItemText(i))
                        
                        
                        oTemp.Remove ("A" & g_alvItemText(i))
                        
                        Next i

                    Case cszOperateLogMan
                        '**************
                        For i = 1 To nLen
                        On Error GoTo ErrorHandle
                        g_oSysMan.DeleteOperateLog CLng(g_alvItemText(i))
                        
                        
                        oTemp.Remove ("A" & g_alvItemText(i))
                        
                        Next i


                    Case cszUnitMan 'ɾ����λ
                        For i = 1 To nLen
                        If g_alvItemText(i) = g_szLocalUnit Then
                            MsgBox "����ɾ������λ.", vbInformation, cszMsg
                        Else
                            oUnit.Init g_oActUser
                            For j = 1 To ArrayLength(g_atAllUnit)
                                If g_atAllUnit(j).szUnitID = g_alvItemText(i) Then
                                    On Error GoTo ErrorHandle
                                    oUnit.Identify g_alvItemText(i)
                                    oUnit.Delete
                                    
                                    
                                    oTemp.Remove ("A" & g_alvItemText(i))
                                    
                                End If
                            Next j
                        End If
                        Next i
                    Case cszStationMan
                            MsgBox "����ɾ����վ��Ϣ.", vbInformation, cszMsg
                    Case Else
                    '''''
                End Select
                LoadCommonData
                
                
            ReDim g_alvItemText(1)
            lvDetail.Refresh
            
            
    
    End If
    End If
Exit Sub
ErrorHandle:
    ShowErrorMsg
    ReDim g_alvItemText(1)
End Sub

Private Sub LV2_Delect()
    Dim j As Integer
'    Dim oCOM As New Component
    Dim oUserGroup As New UserGroup
    Dim oTemp As ListItems, nLen As Integer, i As Integer
    Set oTemp = lvDetail2.ListItems
    nLen = ArrayLength(g_alvItemText2)
    If nLen > 0 Then
        If g_alvItemText2(1) <> "" Then
            For i = 1 To nLen
                If g_szCurrentTask = cszUserGroupMan Then
                    '***************����User�����DElete����
                    For j = 1 To ArrayLength(g_atUserGroupInfo)
                        If g_atUserGroupInfo(j).UserGroupID = g_alvItemText2(i) Then
                            If g_atUserGroupInfo(j).InnerGroup = True Then
                                On Error GoTo ErrorHandle
                                oUserGroup.Init g_oActUser
                                oUserGroup.Identify g_alvItemText2(i)
                                oUserGroup.Delete
                                oTemp.Remove ("A" & g_alvItemText2(i))
                            Else
                                MsgBox "����ɾ�������û���.", vbInformation, cszMsg
                            End If
                        End If
                    Next j
                    LoadCommonData
                End If
            Next i
            ReDim g_alvItemText(1)
            lvDetail.Refresh
        End If
    End If
Exit Sub
ErrorHandle:
    ShowErrorMsg
    ReDim g_alvItemText(1)
End Sub

Private Sub mnuCPassWord_Click()
    Dim oShell As New CommDialog
    oShell.Init g_oActUser
    oShell.ShowUserInfo
    g_bShowUserInfo = True
End Sub

Private Sub mnuExit_Click()
    Unload Me
    Unload frmSMCMain
End Sub

Private Sub mnuModiPara_Click()
    frmSystemParam.Show vbModal, m_frmMain
End Sub

Private Sub mnuSetTicketType_Click()
    
    frmModifyTicketType.Show vbModal
End Sub

Private Sub mnuWriteOffTicket_Click()
    frmWriteOffCheck.Show vbModal, m_frmMain
End Sub

Private Sub pmnuAddGroup_Click()
    frmAEGroup.bEdit = False
    frmAEGroup.Show vbModal, m_frmMain
End Sub

Private Sub pmnuAddStation_Click()
    frmAEStation.bEdit = False
    frmAEStation.Show vbModal, m_frmMain
End Sub

Private Sub pmnuAddUnit_Click()
    frmAEUnit.bEdit = False
    frmAEUnit.Show vbModal, m_frmMain
End Sub

Private Sub pmnuAddUser_Click()
    frmAEUser.bEdit = False
    frmAEUser.Show vbModal, m_frmMain
End Sub

Private Sub pmnuDelectGroup_Click()
    Dim nTem As Integer
    nTem = MsgBox("ȷ��ɾ���û���", vbYesNo + vbExclamation, cszMsg)
    If nTem = vbYes Then
        LV2_Delect
    End If

End Sub

Private Sub pmnuDelectUser_Click()
    Dim nTem As Integer
    nTem = MsgBox("ȷ��ɾ���û�", vbYesNo + vbExclamation, cszMsg)
    If nTem = vbYes Then
    
        LV_Delect
    End If

End Sub

Private Sub pmnuDeleteLoginLog_Click()
    Dim nTem As Integer
    nTem = MsgBox("ȷ��ɾ����¼��־.", vbYesNo + vbExclamation, cszMsg)
    If nTem = vbYes Then
        LV_Delect
    End If

End Sub

Private Sub pmnuDeleteOpeLog_Click()
    Dim nTem As Integer
    nTem = MsgBox("ȷ��ɾ��������־.", vbYesNo + vbExclamation, cszMsg)
    If nTem = vbYes Then
        LV_Delect
    End If

End Sub

Private Sub pmnuDeleteUnit_Click()
    Dim nTem As Integer
    nTem = MsgBox("ȷ��ɾ����λ", vbYesNo + vbExclamation, cszMsg)
    If nTem = vbYes Then
        LV_Delect
    End If
End Sub

Private Sub pmnuGrantFun_Click()
    
    Dim i As Integer
    i = ArrayLength(g_alvItemText)
    If i > 0 Then
        If g_alvItemText(1) = Empty Then
            MsgBox "����ѡ��ĳһ����,����!", vbInformation, cszMsg
        Else
            frmGrant.bFun = True
            frmGrant.szFunCode = g_alvItemText(1)
            frmGrant.Show vbModal
        End If
    Else
        MsgBox "����ѡ��ĳһ����,����!", vbInformation, cszMsg
    End If
End Sub

Private Sub pmnuGrantFunGroup_Click()
    Dim i As Integer
    i = ArrayLength(g_alvItemText2)
    If i > 0 Then
        If g_alvItemText2(1) = Empty Then
            MsgBox "����ѡ��ĳһ������,����!", vbInformation, cszMsg
        Else
            frmGrant.bFun = False
            frmGrant.szFunGroup = g_alvItemText2(1)
            frmGrant.Show vbModal
        End If
    Else
        MsgBox "����ѡ��ĳһ������,����!", vbInformation, cszMsg
    End If
End Sub

Private Sub pmnuGroupProperty_Click()
    If g_szCurrentTask = cszUserGroupMan Then
        On Error GoTo there
        If ArrayLength(g_alvItemText2) = 1 Then
            If g_alvItemText2(1) <> "" Then
                frmAEGroup.bEdit = True
                frmAEGroup.Show vbModal, Me
            End If
        End If
    End If
Exit Sub
there:
    MsgBox "��ѡ��ĳһ�û���", vbInformation, "RTStationϵͳ����"

End Sub

Private Sub pmnuLogout_Click()
    Dim nTem As Integer
    nTem = MsgBox("ȷ��ǿ��ע��.", vbYesNo + vbExclamation, cszMsg)
    If nTem = vbYes Then
        LV_Delect
    End If
End Sub

Private Sub pmnuRecoverUnit_Click()
'    frmRecover.bRecoverUser = False
    frmRecover.Show vbModal, Me
End Sub

Private Sub pmnuRefresh_Click()
    LoadActUserInfo
End Sub

Private Sub pmnuSelectLoginLog_Click()
    
    frmSel_DelLoginLog.m_bDelLog = False
    frmSel_DelLoginLog.Show vbModal, m_frmMain

End Sub

Private Sub pmnuSelectOpeLog_Click()

    frmSel_DelOperateLog.m_bDelLog = False
    frmSel_DelOperateLog.Show vbModal, m_frmMain

End Sub


Private Sub pmnuSetAutoDelectLog_Click()
    frmAutoDelectLog.Show vbModal, m_frmMain

End Sub

Private Sub pmnuStionProperty_Click()
        If g_szCurrentTask = cszStationMan Then

        On Error GoTo there
        If ArrayLength(g_alvItemText) = 1 Then
        

            If g_alvItemText(1) <> "" Then
                frmAEStation.bEdit = True
                frmAEStation.Show vbModal, m_frmMain
            End If
        End If
    End If
Exit Sub
there:
    MsgBox "��ѡ��ĳһ����", vbInformation, cszMsg

End Sub

Private Sub pmnuUnitProperty_Click()
    If g_szCurrentTask = cszUnitMan Then

        On Error GoTo there
        If ArrayLength(g_alvItemText) = 1 Then
        

            If g_alvItemText(1) <> "" Then
                frmAEUnit.bEdit = True
                frmAEUnit.Show vbModal, m_frmMain
            End If
        End If
    End If
Exit Sub
there:
    MsgBox "��ѡ��ĳһ��λ", vbInformation, cszMsg


End Sub


Private Sub pmnuUserProperty_Click()
    If g_szCurrentTask = cszUserGroupMan Then

        On Error GoTo there
        If ArrayLength(g_alvItemText) = 1 Then
        

            If g_alvItemText(1) <> "" Then
                frmAEUser.bEdit = True
                frmAEUser.Show vbModal, m_frmMain
            End If
        End If
    End If
Exit Sub
there:
    MsgBox "��ѡ��ĳһ�û�", vbInformation, cszMsg

End Sub


Private Sub tvAll_GotFocus()
    EnableMnunOfCellExport (False)
End Sub

'----------------------------------------------------------
Private Sub tvAll_NodeClick(ByVal Node As MSComctlLib.Node)
    
   
    If Node.Key = g_szCurrentTask Then Exit Sub
    Select Case Node.Key
        Case cszActiveUserMan
            ActiveUserMan
            mnu_SubAction(2).Enabled = False
            
        Case cszUserGroupMan
            UserGroupMan
        
        Case cszLogMan
            Node.Expanded = True
            LogMan
        
        Case cszLoginLogMan
           LoginLogMan 'Ĭ�ϵ���
        
        Case cszOperateLogMan
            OperateLogMan 'Ĭ�ϵ���
        
        Case cszFunctionMan
            Node.Expanded = True
            ClearActionMenu
            mnu_SubAction(MProperty).Enabled = False
            Dim lvtemp As ListView
            Set lvtemp = m_frmMain.lvDetail
            lvtemp.ColumnHeaders.Clear
            ShowHowDetail

        Case cszFun_GroupMan
            Fun_GroupMan
        Case cszComponent
            ComponentMan
        
        Case cszUnitMan
            UnitMan
        Case cszStationMan
            StationMan
        Case cszConRoot
            ShowRoot
        
    End Select

    g_szCurrentTask = Node.Key
    
    With frmSMCMain
        Select Case g_szCurrentTask
            Case cszActiveUserMan
                .lvDetail.HelpContextID = cnStartHID + EActUserHID + cnlvDetailHID
                .tvAll.HelpContextID = .lvDetail.HelpContextID
            Case cszUserGroupMan
                .lvDetail.HelpContextID = cnStartHID + EUserAndGroupHID + cnlvDetailHID
                .tvAll.HelpContextID = .lvDetail.HelpContextID
                .lvDetail2.HelpContextID = cnStartHID + EUserAndGroupHID + cnlvDetail2HID
            Case cszLogMan
                .lvDetail.HelpContextID = cnStartHID + ELogHID + cnlvDetailHID
                .tvAll.HelpContextID = .lvDetail.HelpContextID
            Case cszLoginLogMan
                .lvDetail.HelpContextID = cnStartHID + ELogHID + cnlvDetailHID
                .tvAll.HelpContextID = .lvDetail.HelpContextID
            Case cszOperateLogMan
                .lvDetail.HelpContextID = cnStartHID + ELogHID + cnlvDetailHID
                .tvAll.HelpContextID = .lvDetail.HelpContextID
            Case cszFunctionMan
                .lvDetail.HelpContextID = cnStartHID + EFunctionAndComHID + cnlvDetailHID
                .tvAll.HelpContextID = .lvDetail.HelpContextID
                .lvDetail2.HelpContextID = cnStartHID + EFunctionAndComHID + cnlvDetail2HID
            Case cszUnitMan
                .lvDetail.HelpContextID = cnStartHID + EUnitHID + cnlvDetailHID
                .tvAll.HelpContextID = .lvDetail.HelpContextID
            Case cszConRoot
                .lvDetail.HelpContextID = cnStartHID + ERootHID + cnlvDetailHID
                .tvAll.HelpContextID = .lvDetail.HelpContextID
            
        End Select
    End With
    
End Sub



Public Sub LogMan()

    mnu_Action.Visible = True
    ShowHowDetail True
    ClearActionMenu
    mnu_SubAction(MProperty).Enabled = False
    AddSubAction "��־����(&T)"
    mnu_SubAction(1).Enabled = True
    mnu_Action.Visible = True
    
    Dim lvtemp As ListView
    Set lvtemp = m_frmMain.lvDetail
    lvtemp.ColumnHeaders.Clear

End Sub
Public Sub DoActLogMan(pnIndex As Integer)
   pmnuSetAutoDelectLog_Click
End Sub



Public Sub OpenDefLoginLog(paszUserID() As String, ByVal pdtBeginDate As Date, ByVal pdtEndDate As Date, ByVal pdtBeginTime As Date, ByVal pdtEndTime As Date, paszWorkstation() As String)
    Dim aszTemp1() As String, bTemp1 As Boolean
    Dim aszTemp2() As String, bTemp2 As Boolean
    Dim nRsCount As Integer, i As Integer
    Dim liTemp As ListItem
    bTemp1 = False
    bTemp2 = False

    SetBusy
    If ArrayLength(paszUserID) = 1 Then
        If paszUserID(1) = "" Then
                bTemp1 = True
        End If
    End If
    If ArrayLength(paszWorkstation) = 1 Then
        If paszWorkstation(1) = "" Then
                bTemp2 = True
        End If
    End If

    On Error GoTo ErrorHandle
    If bTemp1 = True And bTemp2 = False Then
        Set rsLoginInfo = g_oSysMan.GetLoginLogRs(aszTemp1, pdtBeginDate, pdtEndDate, pdtBeginTime, pdtEndTime, paszWorkstation)
    ElseIf bTemp1 = True And bTemp2 = True Then
        Set rsLoginInfo = g_oSysMan.GetLoginLogRs(aszTemp1, pdtBeginDate, pdtEndDate, pdtBeginTime, pdtEndTime, aszTemp2)
    ElseIf bTemp1 = False And bTemp2 = False Then
        Set rsLoginInfo = g_oSysMan.GetLoginLogRs(paszUserID, pdtBeginDate, pdtEndDate, pdtBeginTime, pdtEndTime, paszWorkstation)
    Else
        Set rsLoginInfo = g_oSysMan.GetLoginLogRs(paszUserID, pdtBeginDate, pdtEndDate, pdtBeginTime, pdtEndTime, aszTemp2)
    End If
    

    frmSMCMain.lvDetail.ListItems.Clear
    nRsCount = rsLoginInfo.RecordCount
    frmSMCMain.lvDetail.ColumnHeaders(1).Width = 1000
    frmSMCMain.lvDetail.ColumnHeaders(2).Width = 2000
    frmSMCMain.lvDetail.ColumnHeaders(3).Width = 1000
    frmSMCMain.lvDetail.ColumnHeaders(6).Width = 2000
    frmSMCMain.lvDetail.ColumnHeaders(2).Alignment = lvwColumnCenter
    frmSMCMain.lvDetail.ColumnHeaders(6).Alignment = lvwColumnCenter
    If nRsCount <> 0 Then
        For i = 1 To nRsCount
            Set liTemp = frmSMCMain.lvDetail.ListItems.Add(, "A" & FormatDbValue(rsLoginInfo!login_event_id), FormatDbValue(rsLoginInfo!login_event_id))
            liTemp.SubItems(1) = FormatDbValue(rsLoginInfo!login_start_time)
            liTemp.SubItems(2) = FormatDbValue(rsLoginInfo!user_id)
            liTemp.SubItems(3) = FormatDbValue(rsLoginInfo!computer_name)
            liTemp.SubItems(4) = FormatDbValue(rsLoginInfo!ip_address)
            liTemp.SubItems(5) = FormatDbValue(rsLoginInfo!login_off_time)
            rsLoginInfo.MoveNext
        Next i
    End If
    SetNormal
Exit Sub
ErrorHandle:
    ShowErrorMsg
    SetNormal
End Sub

Public Sub OpenDefOpeLog(paszUserID() As String, ByVal pdtBeginDate As Date, ByVal pdtEndDate As Date, ByVal pdtBeginTime As Date, ByVal pdtEndTime As Date, paszFunOrGroup() As String, Optional pbIsFunction As Boolean = True, Optional pszLogLike As String = "")
    Dim aszTemp1() As String, bTemp1 As Boolean
    Dim aszTemp2() As String, bTemp2 As Boolean
    Dim nRsCount As Integer, i As Integer
    Dim liTemp As ListItem
    bTemp1 = False
    bTemp2 = False
    
    SetBusy
    If ArrayLength(paszUserID) = 1 Then
        If paszUserID(1) = "" Then
                bTemp1 = True
        End If
    End If
    If ArrayLength(paszFunOrGroup) = 1 Then
        If paszFunOrGroup(1) = "" Then
                bTemp2 = True
        End If
    End If
    
    On Error GoTo ErrorHandle
    If bTemp1 = True And bTemp2 = True Then
        Set rsOpeInfo = g_oSysMan.GetOperateLogRs(aszTemp1, pdtBeginDate, pdtEndDate, pdtBeginTime, pdtEndTime, aszTemp2, , pszLogLike)
    ElseIf bTemp1 = True And bTemp2 = False Then
        Set rsOpeInfo = g_oSysMan.GetOperateLogRs(aszTemp1, pdtBeginDate, pdtEndDate, pdtBeginTime, pdtEndTime, paszFunOrGroup, pbIsFunction, pszLogLike)
    ElseIf bTemp1 = False And bTemp2 = True Then
        Set rsOpeInfo = g_oSysMan.GetOperateLogRs(paszUserID, pdtBeginDate, pdtEndDate, pdtBeginTime, pdtEndTime, aszTemp2, , pszLogLike)
    Else
        Set rsOpeInfo = g_oSysMan.GetOperateLogRs(paszUserID, pdtBeginDate, pdtEndDate, pdtBeginTime, pdtEndTime, paszFunOrGroup, pbIsFunction, pszLogLike)
    End If
    
    
    frmSMCMain.lvDetail.ListItems.Clear
    nRsCount = rsOpeInfo.RecordCount
    frmSMCMain.lvDetail.ColumnHeaders(1).Width = 1000
    frmSMCMain.lvDetail.ColumnHeaders(2).Width = 1000
    frmSMCMain.lvDetail.ColumnHeaders(3).Width = 3000
    frmSMCMain.lvDetail.ColumnHeaders(4).Width = 2000
    frmSMCMain.lvDetail.ColumnHeaders(4).Alignment = lvwColumnCenter

    If nRsCount > 0 Then
        For i = 1 To nRsCount
        Set liTemp = frmSMCMain.lvDetail.ListItems.Add(, "A" & FormatDbValue(rsOpeInfo!operation_event_id), FormatDbValue(rsOpeInfo!operation_event_id))
            liTemp.SubItems(1) = FormatDbValue(rsOpeInfo!user_id)
            liTemp.SubItems(4) = FormatDbValue(rsOpeInfo!function_group_id)
            liTemp.SubItems(5) = FormatDbValue(rsOpeInfo!function_id)
            liTemp.SubItems(3) = FormatDbValue(rsOpeInfo!operation_time)
            liTemp.SubItems(2) = FormatDbValue(rsOpeInfo!Annotation)
            rsOpeInfo.MoveNext
        Next i
    End If
    SetNormal
Exit Sub
ErrorHandle:
    ShowErrorMsg
    SetNormal
End Sub

Private Sub LoadUser_Group()
    Dim nLen1 As Integer, nLen2 As Integer
    Dim i As Integer
    
'    If bNewOfUser_Group = False Then
    SetBusy
    
        On Error GoTo ErrorHandle
'        oUser.Init g_oActUser
'        oUserGroup.Init g_oActUser
        
        nLen1 = ArrayLength(g_atUserInfo) 'resume next nLen1=0
        
        
        If nLen1 <> 0 Then
            DisplayUserInfo (nLen1)
        End If
        nLen2 = ArrayLength(g_atUserGroupInfo) '''''
        If nLen2 <> 0 Then
            DisplayGroupInfo (nLen2)
        End If


    SetNormal
Exit Sub
ErrorHandle:
    ShowErrorMsg
    SetNormal
End Sub

Public Sub DisplayUserInfo(arrLen As Integer)
    Dim i As Integer, j As Integer
    Dim szTemp As String
    
    frmSMCMain.lvDetail.ListItems.Clear
    frmSMCMain.lvDetail.ColumnHeaders(4).Width = 2500
    If arrLen > 0 Then
        For i = 1 To arrLen
            Dim xListItem As ListItem
            Set xListItem = frmSMCMain.lvDetail.ListItems.Add(, "A" & g_atUserInfo(i).UserID, g_atUserInfo(i).UserID, "user", "user")
                xListItem.SubItems(1) = g_atUserInfo(i).UserName
                If g_atUserInfo(i).InnerUser = True Then
                    xListItem.SubItems(2) = "��"
                Else
                    xListItem.SubItems(2) = "��"
                End If
                szTemp = g_atUserInfo(i).UnitID
                For j = 1 To ArrayLength(g_atAllUnit)
                    If g_atUserInfo(i).UnitID = g_atAllUnit(j).szUnitID Then
                        szTemp = szTemp & "[" & g_atAllUnit(j).szUnitFullName & "]"
                        xListItem.SubItems(3) = szTemp
                    End If
                Next j
        Next i
    End If
End Sub

Public Sub DisplayGroupInfo(arrLen As Integer)
    Dim i As Integer
    frmSMCMain.lvDetail2.ListItems.Clear
    frmSMCMain.lvDetail2.ColumnHeaders(4).Width = 3000
    For i = 1 To arrLen
        Dim xListItem As ListItem
        Set xListItem = frmSMCMain.lvDetail2.ListItems.Add(, "A" & g_atUserGroupInfo(i).UserGroupID, g_atUserGroupInfo(i).UserGroupID, "usergroup", "usergroup")
            xListItem.SubItems(1) = g_atUserGroupInfo(i).GroupName
            If g_atUserGroupInfo(i).InnerGroup = True Then
                xListItem.SubItems(2) = "��"
            Else
                xListItem.SubItems(2) = "��"
            End If
            xListItem.SubItems(3) = g_atUserGroupInfo(i).Annotation
    Next i

End Sub
Public Function GetAllFunGroup() As String()
    Dim nLen1 As Integer, nLen2 As Integer
    Dim i As Integer, j As Integer
    Dim aszTemp() As String
    Dim bSame As Boolean

    nLen1 = ArrayLength(g_atAllFun)


    If nLen1 <> 0 Then
        bSame = False
        nLen2 = 1
        ReDim aszTemp(1 To nLen2)
        For i = 1 To nLen1
            For j = 1 To nLen2
                If g_atAllFun(i).szFunctionGroup = aszTemp(j) Then
                    bSame = True
                End If
            Next j
            If bSame = False Then


                aszTemp(nLen2) = g_atAllFun(i).szFunctionGroup
                nLen2 = nLen2 + 1
                ReDim Preserve aszTemp(1 To nLen2)

            End If
            bSame = False
        Next i
    End If

    GetAllFunGroup = aszTemp

End Function


Public Sub LoadCommonData()
    SetBusy
    On Error GoTo ErrorHandle
    g_oSysMan.Init g_oActUser
    g_atAllUnit = g_oSysMan.GetAllUnit
    g_atAllUnitDelTag = g_oSysMan.GetAllUnit(False)
    g_atAllSellStation = g_oSysMan.GetAllSellStation
    g_atAllFun = g_oSysMan.GetAllFunction
    g_atUserGroupInfo = g_oSysMan.GetAllUserGroup
    
    g_atUserInfo = g_oSysMan.GetAllUser(True, False)
    g_atUserInfoDelTag = g_oSysMan.GetAllUser(False, False)
    g_atAllUserInfo = g_oSysMan.GetAllUser(True)

    g_oSysParam.Init g_oActUser
    g_szLocalUnit = g_oSysParam.UnitID
    
    SetNormal
Exit Sub
ErrorHandle:
    ShowErrorMsg
    SetNormal
End Sub

Private Sub LoadActUserInfo()
    Dim liTemp As ListItem, nLen As Integer, i As Integer, j As Integer, nLenUser As Integer

    SetBusy '    Me.MousePointer = vbHourglass
    On Error GoTo ErrorHandle
    aTactUserInfo = g_oSysMan.GetAllActiveUser
    nLen = ArrayLength(aTactUserInfo)
    nLenUser = ArrayLength(g_atUserInfo)
    If nLenUser > 0 Then
        lvDetail.ListItems.Clear
        lvDetail.ColumnHeaders(5).Width = 2000
        lvDetail.ColumnHeaders(6).Width = 2000
        lvDetail.ColumnHeaders(5).Alignment = lvwColumnCenter
        lvDetail.ColumnHeaders(6).Alignment = lvwColumnCenter
        If nLen <> 0 Then
        For i = 1 To nLen
            Set liTemp = lvDetail.ListItems.Add(, "A" & aTactUserInfo(i).lLoginEventID, aTactUserInfo(i).lLoginEventID, "actuser", "actuser")
            liTemp.SubItems(1) = aTactUserInfo(i).szUserID
            For j = 1 To nLenUser
                If UCase(g_atUserInfo(j).UserID) = UCase(aTactUserInfo(i).szUserID) Then
                    liTemp.SubItems(2) = g_atUserInfo(j).UserName
                End If
            Next j
            liTemp.SubItems(3) = aTactUserInfo(i).szLoginHost
            liTemp.SubItems(4) = Format(aTactUserInfo(i).dtLoginTime, "yyyy-mm-dd hh:mm:ss")
            liTemp.SubItems(5) = Format(aTactUserInfo(i).dtLastTime, "yyyy-mm-dd hh:mm:ss")
        Next i
        End If
    End If
    SetNormal
Exit Sub
ErrorHandle:
    ShowErrorMsg
    SetNormal
End Sub

'Public Sub LoadFun_GroupInfo()
'    Dim liTemp As ListItem, nLen As Integer, i As Integer
'
'    nLen = ArrayLength(g_atAllFun)
'    '���ܲ���
'    frmSMCMain.lvDetail.ListItems.Clear
'    frmSMCMain.lvDetail.ColumnHeaders(5).Width = 3000
'    If nLen <> 0 Then
'        For i = 1 To nLen
'            Set liTemp = frmSMCMain.lvDetail.ListItems.Add(, "A" & g_atAllFun(i).szFunctionCode, g_atAllFun(i).szFunctionCode, "funright", "funright")
'            liTemp.SubItems(1) = g_atAllFun(i).szFunctionName
'            liTemp.SubItems(2) = g_atAllFun(i).szcomponentID
'            liTemp.SubItems(3) = g_atAllFun(i).szFunctionGroup
'            liTemp.SubItems(4) = g_atAllFun(i).szAnnotation
'        Next i
'    End If
'
'
'
'    '�����鲿��
'    Dim aszFunGroup() As String, nLen1 As Integer, j As Integer
'    Dim bFirst As Boolean, szTemp As String, bTooLong As Boolean
'
'    aszFunGroup = GetAllFunGroup
'
'    nLen = ArrayLength(aszFunGroup)
'    nLen1 = ArrayLength(g_atAllFun)
'
'
'    frmSMCMain.lvDetail2.ListItems.Clear
'    frmSMCMain.lvDetail2.ColumnHeaders(3).Width = 20000
'    If nLen <> 0 Then
'        For i = 1 To nLen
'            bFirst = True
'            bTooLong = False
'            If aszFunGroup(i) <> "" Then
'            Set liTemp = frmSMCMain.lvDetail2.ListItems.Add(, "A" & aszFunGroup(i), aszFunGroup(i), "funright", "funright")
'            For j = 1 To nLen1
'                If g_atAllFun(j).szFunctionGroup = aszFunGroup(i) Then
'                    If bFirst = True Then
'                        szTemp = g_atAllFun(j).szFunctionName
'                        liTemp.SubItems(1) = g_atAllFun(j).szcomponentID
'                        bFirst = False
'                    Else
'                        If LenB(StrConv(szTemp, vbUnicode)) > 240 Then
'                            If bTooLong = False Then
'                                szTemp = szTemp & "......"
'                                bTooLong = True
'                            End If
'                        Else
'                            szTemp = szTemp & "," & g_atAllFun(j).szFunctionName
'                        End If
'                    End If
'                End If
'            Next j
'            End If
'            liTemp.SubItems(2) = szTemp
'
'        Next i
'    End If
''    setnormal
'Exit Sub
'ErrorHandle:
'    ShowErrorMsg
'    SetNormal
'End Sub

Public Sub LoadStationInfo()
    Dim i As Integer, nLen As Integer
    Dim liTemp As ListItem
    
'    setbusy
    
    nLen = 0
    
    nLen = ArrayLength(g_atAllSellStation)
    
    frmSMCMain.lvDetail.ListItems.Clear
    frmSMCMain.lvDetail.ColumnHeaders(5).Width = 3000
    If nLen <> 0 Then
        For i = 1 To nLen
            Set liTemp = frmSMCMain.lvDetail.ListItems.Add(, "A" & g_atAllSellStation(i).szSellStationID, g_atAllSellStation(i).szSellStationID, "stationman", "stationman")
            liTemp.SubItems(1) = g_atAllSellStation(i).szSellStationName
            liTemp.SubItems(2) = g_atAllSellStation(i).szSellStationFullName
            liTemp.SubItems(3) = g_atAllSellStation(i).szUnitFullName
            liTemp.SubItems(4) = g_atAllSellStation(i).szAnnotation
        Next i
    End If
    'setnormal
            
End Sub

Public Sub LoadUnitInfo()
    Dim i As Integer, nLen As Integer
    Dim liTemp As ListItem
    
'    setbusy
    
    nLen = 0
    
    nLen = ArrayLength(g_atAllUnit)
    
    frmSMCMain.lvDetail.ListItems.Clear
    frmSMCMain.lvDetail.ColumnHeaders(6).Width = 3000
    If nLen <> 0 Then
        For i = 1 To nLen
            If g_atAllUnit(i).szUnitID = g_szLocalUnit Then
                Set liTemp = frmSMCMain.lvDetail.ListItems.Add(, "A" & g_atAllUnit(i).szUnitID, g_atAllUnit(i).szUnitID, "localunit", "localunit")
            Else
                Set liTemp = frmSMCMain.lvDetail.ListItems.Add(, "A" & g_atAllUnit(i).szUnitID, g_atAllUnit(i).szUnitID, "unitman", "unitman")
            End If
            liTemp.SubItems(1) = g_atAllUnit(i).szUnitShortName
            liTemp.SubItems(2) = g_atAllUnit(i).szUnitFullName
            If g_atAllUnit(i).szUnitID = g_szLocalUnit Then
                liTemp.SubItems(3) = "����λ"
            Else
                Select Case g_atAllUnit(i).nUnitType
                    Case TP_UnitSC
                        liTemp.SubItems(3) = "���۵�λ"
                    Case TP_UnitClient
                        liTemp.SubItems(3) = "���۵�λ"
                    Case TP_UnitServer
                        liTemp.SubItems(3) = "��Ʊ�����ṩ��λ"
                End Select
            End If
            liTemp.SubItems(4) = g_atAllUnit(i).szIPAddress
            liTemp.SubItems(5) = g_atAllUnit(i).szAnnotation
'            liTemp.SubItems(6) = g_atAllUnit(i).dbSellCharge
        Next i
    End If
    'setnormal
            
End Sub

Private Sub ComponentMan()
    Dim lvtemp As Object
    Set lvtemp = m_frmMain.lvDetail
    lvtemp.ListItems.Clear
    lvtemp.ColumnHeaders.Clear
    lvtemp.ColumnHeaders.Add , "�������", "�������"
    lvtemp.ColumnHeaders.Add , "�������", "�������"
    lvtemp.ColumnHeaders.Add , "����汾", "����汾"
    lvtemp.ColumnHeaders.Add , "װ��ʱ��", "װ��ʱ��"
    ShowHowDetail
    
    ClearActionMenu
    mnu_SubAction(MProperty).Enabled = False
'    AddSubAction "�������(&L)"
'    mnu_SubAction(1).Enabled = True
'    mnu_Action.Visible = True
    
    LoadComponentInfo

End Sub

Public Sub LoadComponentInfo()
    Dim liTemp As ListItem, nLen As Integer, i As Integer
    
    
    SetBusy
    On Error GoTo ErrorHandle
    aTAllCOMInfo = g_oSysMan.GetAllCOM
    

    nLen = 0
    
    nLen = ArrayLength(aTAllCOMInfo)
    
    
    frmSMCMain.lvDetail.ListItems.Clear
    frmSMCMain.lvDetail.ColumnHeaders(4).Width = 3000
    If nLen <> 0 Then
        For i = 1 To nLen
            Set liTemp = frmSMCMain.lvDetail.ListItems.Add(, "A" & aTAllCOMInfo(i).szCOMCode, aTAllCOMInfo(i).szCOMCode, "component", "component")
            liTemp.SubItems(1) = aTAllCOMInfo(i).szCOMName
            liTemp.SubItems(2) = aTAllCOMInfo(i).szCOMVersion
            liTemp.SubItems(3) = aTAllCOMInfo(i).dtLoadTime
        Next i
    End If
    SetNormal
Exit Sub
ErrorHandle:
    ShowErrorMsg
    SetNormal
    
End Sub

Public Sub DoActActiveUser(pnIndex As Integer)
    Select Case pnIndex
        Case 0
         '
        Case MActUserRefresh
            pmnuRefresh_Click
        Case MFroceLogout
            pmnuLogout_Click
    End Select

End Sub

Public Sub DoActFun_Group(pnIndex As Integer)
    Select Case pnIndex
        Case 0
         '
        Case 1
            pmnuGrantFun_Click
        Case 2
            pmnuGrantFunGroup_Click
    End Select
End Sub

Private Sub EnableMnunOfCellExport(Optional IsEnable As Boolean = True)
    If IsEnable = True Then
        mnu_ExprotFile.Enabled = True
        mnu_ExpOpen.Enabled = True
        mnu_PrintEX.Enabled = True
        mnu_PrintSet.Enabled = True
        mnu_PageSet.Enabled = True
        mnu_PrintView.Enabled = True
    Else
        mnu_ExprotFile.Enabled = False
        mnu_ExpOpen.Enabled = False
        mnu_PrintEX.Enabled = False
        mnu_PrintSet.Enabled = False
        mnu_PageSet.Enabled = False
        mnu_PrintView.Enabled = False
    End If
End Sub

Private Sub InitCellExport(CellExpSourceName As Object, SelectSource As Integer)
'    Set CellExport.ListViewSource = Nothing
'    Set CellExport.ListViewSource = CellExpSourceName
'    CellExport.SourceSelect = SelectSource
End Sub

