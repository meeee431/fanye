VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUser_GroupRight 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmUser_GroupRight"
   ClientHeight    =   4185
   ClientLeft      =   1380
   ClientTop       =   2760
   ClientWidth     =   9390
   HelpContextID   =   50000310
   Icon            =   "frmUser_GroupRight.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkInherit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "(&D)显示包含从组继承的功能"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6720
      TabIndex        =   16
      Top             =   735
      Width           =   2580
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   ">>"
      Height          =   285
      Index           =   0
      Left            =   4215
      TabIndex        =   10
      Top             =   1455
      Width           =   600
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   ">"
      Height          =   285
      Index           =   1
      Left            =   4215
      TabIndex        =   9
      Top             =   1920
      Width           =   600
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "<"
      Height          =   285
      Index           =   2
      Left            =   4215
      TabIndex        =   8
      Top             =   2400
      Width           =   600
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "<<"
      Height          =   285
      Index           =   3
      Left            =   4215
      TabIndex        =   7
      Top             =   2895
      Width           =   600
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   315
      Left            =   8190
      TabIndex        =   6
      Top             =   3735
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   315
      Left            =   6915
      TabIndex        =   1
      Top             =   3735
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   315
      Left            =   5625
      TabIndex        =   0
      Top             =   3735
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvBrowse 
      Height          =   2625
      Left            =   4890
      TabIndex        =   11
      Top             =   1020
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   4630
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvAuthor 
      Height          =   2625
      Left            =   120
      TabIndex        =   12
      Top             =   1020
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   4630
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   945
      X2              =   9315
      Y1              =   615
      Y2              =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   945
      X2              =   9315
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "功能权限"
      Height          =   180
      Left            =   150
      TabIndex        =   15
      Top             =   495
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已授权功能"
      Height          =   180
      Left            =   4890
      TabIndex        =   14
      Top             =   780
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "所有功能"
      Height          =   180
      Left            =   150
      TabIndex        =   13
      Top             =   780
      Width           =   720
   End
   Begin VB.Label lblName 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4515
      TabIndex        =   5
      Top             =   195
      Width           =   90
   End
   Begin VB.Label lblID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1200
      TabIndex        =   4
      Top             =   195
      Width           =   90
   End
   Begin VB.Label lblName1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户组名:"
      Height          =   180
      Left            =   3585
      TabIndex        =   3
      Top             =   195
      Width           =   810
   End
   Begin VB.Label lblID1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户组代码:"
      Height          =   180
      Left            =   150
      TabIndex        =   2
      Top             =   195
      Width           =   990
   End
End
Attribute VB_Name = "frmUser_GroupRight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *******************************************************************
' *  Source File Name  : frmUser_GroupRight                         *
' *  Project Name: PSTSMan                                          *
' *  Engineer:                                                      *
' *  Date Generated: 2002/08/19                                     *
' *  Last Revision Date : 2002/08/19                                *
' *  Brief Description   : 给用户或用户组授权                       *
' *******************************************************************

Option Explicit
Option Base 1
Public m_bUser As Boolean
Dim nPar, nStart, nEnd As Integer
Dim aTAllFunTemp() As TCOMFunctionInfoShort
Dim nAllFunCount As Integer
Dim bIsParant As Boolean

Private Sub chkInherit_Click()
    GetRightInfoForLV
    InitLV_TV
End Sub

Private Sub cmdCancel_Click()
    If m_bUser = True Then
        If frmAEUser.bEdit = False Then
            ReDim g_atInBrowse(1)
        End If
    Else
        If frmAEGroup.bEdit = False Then
            ReDim g_atInBrowse(1)
        End If
    End If
    Unload Me

End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me, content
End Sub

Private Sub cmdOK_Click()
    Dim lmsg As Long
    SetBusy
    If m_bUser = True Then
        frmAEUser.bRightRead = True
        If frmAEUser.bEdit = True Then
            If chkInherit.Value = Unchecked Then
                ModifyUserRight
            Else
                lmsg = MsgBox("所有显示的功能将设定为直接功能,确认吗?", vbQuestion + vbYesNo, cszMsg)
                If lmsg = vbYes Then
                    ModifyUserRight
                Else
                    Exit Sub
                End If
            End If
        End If
    Else
        frmAEGroup.bRightRead = True
        If frmAEGroup.bEdit = True Then
            ModifyGroupRight
        End If

    End If
    SetNormal
    Unload Me
End Sub

Private Sub lvBrowse_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static bAsc As Boolean
    bAsc = Not bAsc
    lvBrowse.SortKey = ColumnHeader.Index - 1
    If bAsc Then
        lvBrowse.SortOrder = lvwAscending
    Else
        lvBrowse.SortOrder = lvwDescending
    End If
    lvBrowse.Sorted = True
End Sub

Private Sub lvBrowse_DblClick()
    cmdSelect_Click (2)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 0
        lvBrowse.View = lvwIcon
    Case 1
        lvBrowse.View = lvwSmallIcon
    Case 2
        lvBrowse.View = lvwList
    Case 3
        lvBrowse.View = lvwReport
    End Select
End Sub

Private Sub cmdSelect_Click(Index As Integer)
On Error GoTo errHandler
Dim oTemp As ListItem, i As Integer, aszTemp() As String, j As Integer
    Select Case Index
    Case 0  'Add All
            ReDim g_atAddBrowse(1 To ArrayLength(g_atAllFun))
            lvBrowse.ListItems.Clear
            For i = 1 To UBound(g_atAddBrowse)
                g_atAddBrowse(i).FunID = g_atAllFun(i).szFunctionCode
                g_atAddBrowse(i).FunName = g_atAllFun(i).szFunctionName
                g_atAddBrowse(i).FunGroup = g_atAllFun(i).szFunctionGroup
                Set oTemp = lvBrowse.ListItems.Add(, "A" & g_atAddBrowse(i).FunID, g_atAddBrowse(i).FunID & "[" & g_atAddBrowse(i).FunName & "]")
                oTemp.SubItems(1) = g_atAddBrowse(i).FunGroup
            Next
            GetListviewData  '重置aszInBrowse()
            lvBrowse.Refresh '刷新
            ReDim g_atAddBrowse(1) '置空

        Case 1 'Add Selected
            If g_atAddBrowse(1).FunID <> "" Then
            For i = 1 To UBound(g_atAddBrowse)
                If CompareID(g_atAddBrowse(i).FunID) = False Then '数据是否重复
                    Set oTemp = lvBrowse.ListItems.Add(, "A" & g_atAddBrowse(i).FunID, g_atAddBrowse(i).FunID & "[" & g_atAddBrowse(i).FunName & "]")
                    oTemp.SubItems(1) = g_atAddBrowse(i).FunGroup
                End If
            Next
            End If
            lvBrowse.Refresh '刷新
            GetListviewData  '重置aszInBrowse()
            ReDim g_atAddBrowse(1) '置空

        Case 2 'Delect selected
            Dim oliTemp As ListItem
            i = 0
            For Each oliTemp In lvBrowse.ListItems
                If oliTemp.Selected = True Then
                    i = i + 1
                    ReDim Preserve aszTemp(1 To i)
                    aszTemp(i) = oliTemp.Key
                End If
            Next
            For j = 1 To ArrayLength(aszTemp)
                lvBrowse.ListItems.Remove (aszTemp(j))
            Next j
            lvBrowse.Refresh '刷新
            GetListviewData  '重置aszInBrowse()
            ReDim g_atAddBrowse(1) '置空
        Case 3 'Delect All
            i = lvBrowse.ListItems.Count
            For j = 1 To i
                lvBrowse.ListItems.Remove (1)
            Next j
            lvBrowse.Refresh '刷新
            GetListviewData '重置aszInBrowse()
            ReDim g_atAddBrowse(1) '置空
    End Select
Exit Sub
errHandler:
'    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim oTemp As ColumnHeader

    Me.Top = (Screen.Height - Me.ScaleHeight) / 2
    Me.Left = (Screen.Width - Me.ScaleWidth) / 2
    If m_bUser Then
        Me.Caption = "用户权限"
        lblID1.Caption = "用户代码:"
        lblName1.Caption = "用户名:"
        If frmAEUser.bEdit = True Then
            lblID.Caption = frmAEUser.lblUserID
        Else
            lblID.Caption = frmAEUser.txtUserID
            chkInherit.Enabled = False
            chkInherit.Visible = False
            chkInherit.Value = Unchecked
        End If
        lblName.Caption = frmAEUser.txtUserName

    Else
        Me.Caption = "用户组权限"
        lblID1.Caption = "用户组代码:"
        lblName1.Caption = "用户组名:"

        chkInherit.Enabled = False
        chkInherit.Visible = False

        If frmAEGroup.bEdit = True Then
            lblID.Caption = frmAEGroup.lblGroupID
        Else
            lblID.Caption = frmAEGroup.txtGroupID
        End If
        lblName.Caption = frmAEGroup.txtGroupName
    End If
    Set oTemp = lvBrowse.ColumnHeaders.Add(, , "功能", 2800)
    Set oTemp = lvBrowse.ColumnHeaders.Add(, , "功能组", 1800)
    nAllFunCount = ArrayLength(g_atAllFun) ' g_atAllFun 在frmStortmneu.GetCommonDate中得到
    '所有功能
    If nAllFunCount = 0 Then
        MsgBox "没有组件配置文件的安装信息或组件不可自识别!重装组件配置文件.", vbExclamation, cszMsg
        Unload Me
    Else
        ReDim aTAllFunTemp(1 To nAllFunCount)
        For i = 1 To nAllFunCount
            aTAllFunTemp(i).FunID = g_atAllFun(i).szFunctionCode
            aTAllFunTemp(i).FunName = g_atAllFun(i).szFunctionName
            aTAllFunTemp(i).FunGroup = g_atAllFun(i).szFunctionGroup
        Next i
    End If
'    此用户(组)已授权的功能
    If m_bUser Then
        If frmAEUser.bEdit = False Then
        Else
                ReDim g_atAddBrowse(1)
                GetRightInfoForLV '设定aTAuthored(), g_atInBrowse(),g_aszFunOld() 的初始值(g_atAllFun()由父窗体LOAD时得到)
        End If
    Else
        If frmAEGroup.bEdit = False Then
        Else
                ReDim g_atAddBrowse(1)
                GetRightInfoForLV
        End If
    End If
    ''初始化lvBrowse & tvAuthor
    InitLV_TV
End Sub



Private Sub tvAuthor_DblClick()
    If bIsParant = False Then
        cmdSelect_Click (1)
    Else
        '是否展开
    End If
End Sub

Private Sub tvAuthor_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim ndTempPar As Node
    Dim ndTempChi As Node

    bIsParant = False
    Set ndTempPar = Node.Parent
    If ndTempPar Is Nothing Then
        Set ndTempChi = Node.Child
        nStart = ndTempChi.Index
        nEnd = ndTempChi.LastSibling.Index
        nPar = Node.Index
        bIsParant = True
    Else
        nPar = ndTempPar.Index
        nStart = Node.Index
        nEnd = nStart
    End If
    GetDataForlvBrowse '设置aTAddBrowse()
End Sub

Private Sub GetDataForlvBrowse()
On Error GoTo errHandler
    Dim szTemp As String, szTemp1 As String, szTemp2 As String
    Dim i As Integer, j As Integer, nTemp As Integer, nLen As Integer
    nLen = 0
    ReDim g_atAddBrowse(1)
    For i = 1 To nEnd - nStart + 1
        If (tvAuthor.Nodes(nPar).Text <> "") And (tvAuthor.Nodes(nStart + i - 1).Text <> "") Then
            nLen = nLen + 1
            ReDim Preserve g_atAddBrowse(1 To nLen)
            g_atAddBrowse(nLen).FunGroup = tvAuthor.Nodes(nPar).Text

            szTemp = StrConv(tvAuthor.Nodes(nStart + i - 1).Key, vbFromUnicode)
            nTemp = LenB(szTemp) - 1 '去掉"A"
            szTemp = RightB(szTemp, nTemp)
            szTemp2 = StrConv(szTemp, vbUnicode)
            szTemp1 = StrConv(tvAuthor.Nodes(nStart + i - 1).Text, vbFromUnicode)
            j = LenB(szTemp1) - LenB(szTemp) - 1
            szTemp = StrConv(RightB(szTemp1, j), vbUnicode)

            g_atAddBrowse(nLen).FunID = szTemp2
            g_atAddBrowse(nLen).FunName = szTemp
        End If
    Next i
Exit Sub
errHandler:

End Sub

Private Function CompareID(szTemp As String) As Boolean
'On Error GoTo errHandler
    CompareID = False
    GetListviewData
    Dim i As Integer
    Dim bTemp As Boolean
    bTemp = False
    On Error Resume Next
    i = UBound(g_atInBrowse)
    On Error GoTo 0
    Dim j As Integer
    For j = 1 To i
        If g_atInBrowse(j).FunID <> szTemp Then
            bTemp = False
        Else
            bTemp = True
        End If
        CompareID = CompareID Or bTemp
    Next j

Exit Function
errHandler:
    CompareID = False
End Function
'读lvBrowse数据
Private Sub GetListviewData()
Dim szTemp As String
Dim nTemp As Integer
Dim nTemp1 As Integer
Dim oTemp As ListItem
Dim i As Integer, j As Integer

On Error GoTo errHandler
    j = lvBrowse.ListItems.Count
    If j = 0 Then
        g_bBrowseNull = True
        ReDim g_atInBrowse(1)
    Else
        g_bBrowseNull = False
        ReDim g_atInBrowse(1 To j)
        i = 0
        For Each oTemp In lvBrowse.ListItems
            i = i + 1
            szTemp = oTemp.Key
'            nTemp = LenB(StrConv(szTemp, vbUnicode))
            nTemp = Len(szTemp)
            szTemp = Right(szTemp, nTemp - 1)
            g_atInBrowse(i).FunID = szTemp
            szTemp = oTemp.Text
'            nTemp1 = LenB(StrConv(szTemp, vbUnicode))
            nTemp1 = Len(szTemp)
            nTemp = nTemp1 - nTemp
            szTemp = Right(szTemp, nTemp)
            szTemp = Left(szTemp, nTemp - 1)
            g_atInBrowse(i).FunName = szTemp
            g_atInBrowse(i).FunGroup = oTemp.SubItems(1)
        Next
    End If
Exit Sub
errHandler:

End Sub

Private Sub GetRightInfoForLV() '编辑时使用
    Dim oUserTemp As New User
    Dim aszFunTemp() As String
    Dim i As Integer, j As Integer, k As Integer
    Dim oGroupTemp As New UserGroup
    If m_bUser = True Then
        On Error GoTo ErrorHandle
        oUserTemp.Init g_oActUser
        oUserTemp.Identify lblID
        If chkInherit.Value = Unchecked Then
            aszFunTemp = oUserTemp.GetDirectFunction
            g_aszFunOld = aszFunTemp
        Else
            aszFunTemp = oUserTemp.GetAllFunction
            g_aszFunOld = oUserTemp.GetDirectFunction '直接权限
        End If
    Else
        On Error GoTo ErrorHandle
        oGroupTemp.Init g_oActUser
        oGroupTemp.Identify lblID
        aszFunTemp = oGroupTemp.GetAllFunction
        g_aszFunOld = aszFunTemp '用户组不存在直接权限
    End If
    ReDim g_atAuthored(1)
    i = ArrayLength(aszFunTemp)
    If i > 0 Then
        ReDim g_atAuthored(1 To i)
        For j = 1 To i
            For k = 1 To nAllFunCount
                If g_atAllFun(k).szFunctionCode = aszFunTemp(j) Then
                    g_atAuthored(j).FunID = g_atAllFun(k).szFunctionCode
                    g_atAuthored(j).FunName = g_atAllFun(k).szFunctionName
                    g_atAuthored(j).FunGroup = g_atAllFun(k).szFunctionGroup
                End If
            Next k
        Next j
    End If
    g_atInBrowse() = g_atAuthored()
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub InitLV_TV()

'
    Dim oListItem As ListItem
    Dim oNodeTemp1 As Node, oNodeTemp2 As Node
    Dim nLen As Integer, nLen1 As Integer
    Dim i As Integer, j As Integer
    Dim aszAllFunGroup() As String
    On Error Resume Next
    '初始化listview
    lvBrowse.ListItems.Clear
    If m_bUser = True And frmAEUser.bRightRead = False Then '给用户授权,且内存中无数据
        nLen = 0
        nLen = UBound(g_atAuthored)
        If nLen <> 0 Then
            If g_atAuthored(1).FunID <> "" Then
            For i = 1 To nLen
                Set oListItem = lvBrowse.ListItems.Add(, "A" & g_atAuthored(i).FunID, g_atAuthored(i).FunID & "[" & g_atAuthored(i).FunName & "]")
'                oListItem.SubItems(1) = g_atAuthored(i).FunName
                oListItem.SubItems(1) = g_atAuthored(i).FunGroup
            Next i
            End If
        End If
    ElseIf m_bUser = True And frmAEUser.bRightRead = True Then '给用户授权,且内存中有数据
        nLen = 0
        nLen = UBound(g_atInBrowse)
        If nLen <> 0 Then
            If g_atInBrowse(1).FunID <> "" Then
            For i = 1 To nLen
                Set oListItem = lvBrowse.ListItems.Add(, "A" & g_atInBrowse(i).FunID, g_atInBrowse(i).FunID & "[" & g_atInBrowse(i).FunName & "]")
'                oListItem.SubItems(1) = g_atInBrowse(i).FunName
                oListItem.SubItems(1) = g_atInBrowse(i).FunGroup

            Next i
            End If
        End If
    ElseIf (m_bUser = False) And (frmAEGroup.bRightRead = False) Then '给用户组授权,且内存中无数据
        nLen = 0
        nLen = UBound(g_atAuthored)
        If nLen <> 0 Then
            If g_atAuthored(1).FunID <> "" Then
            For i = 1 To nLen
                Set oListItem = lvBrowse.ListItems.Add(, "A" & g_atAuthored(i).FunID, g_atAuthored(i).FunID & "[" & g_atAuthored(i).FunName & "]")
'                oListItem.SubItems(1) = g_atAuthored(i).FunName
                oListItem.SubItems(1) = g_atAuthored(i).FunGroup
            Next i
            End If
        End If
    ElseIf (m_bUser = False) And (frmAEGroup.bRightRead = True) Then '给用户组授权,且内存中有数据
        nLen = 0
        nLen = UBound(g_atInBrowse)
        If nLen <> 0 Then
            If g_atInBrowse(1).FunID <> "" Then
            For i = 1 To nLen
                Set oListItem = lvBrowse.ListItems.Add(, "A" & g_atInBrowse(i).FunID, g_atInBrowse(i).FunID & "[" & g_atInBrowse(i).FunName & "]")
'                oListItem.SubItems(1) = g_atInBrowse(i).FunName
                oListItem.SubItems(1) = g_atInBrowse(i).FunGroup
            Next i
            End If
        End If
    Else
    End If

    aszAllFunGroup = frmStoreMenu.GetAllFunGroup '得到所有的功能组
    nLen = 0
    nLen = UBound(aTAllFunTemp)
    '初始化Treeview
    tvAuthor.Nodes.Clear
    nLen1 = ArrayLength(aszAllFunGroup)
    If nLen1 = 0 Then
        MsgBox "组件功能信息出错!重装组件配置文件.", vbExclamation, cszMsg
        Unload Me
    Else
        If nLen <> 0 Then
            For i = 1 To nLen1
                If aszAllFunGroup(i) <> "" Then
                Set oNodeTemp1 = tvAuthor.Nodes.Add(, , "A" & aszAllFunGroup(i), aszAllFunGroup(i))
                For j = 1 To nLen
                    If aTAllFunTemp(j).FunGroup = aszAllFunGroup(i) Then
                        If aTAllFunTemp(j).FunID <> "" Then
                        Set oNodeTemp2 = tvAuthor.Nodes.Add(oNodeTemp1, tvwChild, "A" & aTAllFunTemp(j).FunID, aTAllFunTemp(j).FunID & "[" & aTAllFunTemp(j).FunName & "]")
                        End If
                    End If
                Next j
                End If
            Next i
        End If
    End If
    Exit Sub
ErrorHandle:

End Sub


Private Sub ModifyUserRight()
    Dim oUserTemp As New User
    Dim narrLenOld As Integer
    Dim narrLen As Integer
    Dim i As Integer, j As Integer, bShouldDel As Boolean, bShouldAdd As Boolean
    Dim nAddCount As Integer, nDelCount As Integer
    Dim aszDel() As String
    Dim aszAdd() As String
    Dim szUserID As String
    szUserID = lblID
    On Error Resume Next
'    If bRightChange = True Then
        '''修改权限
        oUserTemp.Init g_oActUser
        narrLenOld = 0
        narrLenOld = ArrayLength(g_aszFunOld)
        narrLen = 0
        narrLen = UBound(g_atInBrowse)
        On Error GoTo 0
        On Error GoTo there
        If narrLenOld = 0 And narrLen = 0 Then
            'do nothing
        ElseIf narrLenOld = 0 Then
            If g_atInBrowse(1).FunID <> "" Then
            oUserTemp.Identify szUserID

            For i = 1 To narrLen
                oUserTemp.AddFunction (g_atInBrowse(i).FunID)
            Next i
            End If
        ElseIf narrLen = 0 Then
            If g_aszFunOld(1) <> "" Then
            oUserTemp.Identify szUserID
            For i = 1 To narrLenOld
                oUserTemp.DeleteFunction (g_aszFunOld(i))
            Next i
            End If
        Else
            '""""""""""""""""""增删权力
            bShouldDel = True
            bShouldAdd = True
            nAddCount = 0
            nDelCount = 0
            '删除权力
            For i = 1 To narrLenOld
                For j = 1 To narrLen
                    If g_aszFunOld(i) = g_atInBrowse(j).FunID Then
                        bShouldDel = False
                    End If
                Next j
                If bShouldDel = True Then
                    nDelCount = nDelCount + 1
                    ReDim Preserve aszDel(1 To nDelCount)
                    aszDel(nDelCount) = g_aszFunOld(i)
                End If
                bShouldDel = True
            Next i
            If ArrayLength(aszDel) <> 0 Then
                oUserTemp.Identify szUserID

                For i = 1 To ArrayLength(aszDel)
                    If aszDel(i) <> "" Then
                    On Error GoTo there '修改
                    oUserTemp.DeleteFunction aszDel(i)

                    End If
                Next i
            End If

            '增加权限
            For i = 1 To narrLen
                For j = 1 To narrLenOld
                    If g_atInBrowse(i).FunID = g_aszFunOld(j) Then
                        bShouldAdd = False
                    End If
                Next j
                If bShouldAdd = True Then
                    nAddCount = nAddCount + 1
                    ReDim Preserve aszAdd(1 To nAddCount)
                    aszAdd(nAddCount) = g_atInBrowse(i).FunID
                End If
                bShouldAdd = True
            Next i

            If ArrayLength(aszAdd) <> 0 Then
                oUserTemp.Identify szUserID
                For i = 1 To ArrayLength(aszAdd)
                    If aszAdd(i) <> "" Then
                    On Error GoTo there '修改
                    oUserTemp.AddFunction (aszAdd(i))

                    End If
                Next i
            End If
        End If
'    End If

    '重设aszFunOld()数组
    ReDim g_aszFunOld(1 To narrLen)
    For i = 1 To narrLen
        g_aszFunOld(i) = g_atInBrowse(i).FunID
    Next i

Exit Sub
there:
    ShowErrorMsg

End Sub

Private Sub ModifyGroupRight()
    Dim oGroup As New UserGroup
    Dim nLen As Integer, i As Integer, j As Integer
    Dim nLenOld As Integer
    Dim bShouldDel As Boolean, bShouldAdd As Boolean
    Dim nAddCount As Integer, nDelCount As Integer
    Dim aszDel() As String
    Dim aszAdd() As String
    Dim szGroupID As String
    szGroupID = lblID

    oGroup.Init g_oActUser
    nLenOld = 0
    nLenOld = ArrayLength(g_aszFunOld)
    nLen = 0
    On Error Resume Next
    nLen = UBound(g_atInBrowse)
    On Error GoTo 0
    On Error GoTo ErrorHandle
    If nLenOld = 0 And nLen = 0 Then
        'do nothing
    ElseIf nLenOld = 0 Then
        If g_atInBrowse(1).FunID <> "" Then
        oGroup.Identify szGroupID

        For i = 1 To nLen
            oGroup.AddFunction (g_atInBrowse(i).FunID)
        Next i
        End If
    ElseIf nLen = 0 Then
        If g_aszFunOld(1) <> "" Then
        oGroup.Identify szGroupID
        For i = 1 To nLenOld
            oGroup.DeleteFunction (g_aszFunOld(i))
        Next i
        End If
    Else
        '""""""""""""""""""增删权限
        bShouldDel = True
        bShouldAdd = True
        nAddCount = 0
        nDelCount = 0
        '删除权限
        For i = 1 To nLenOld
            For j = 1 To nLen
                If g_aszFunOld(i) = g_atInBrowse(j).FunID Then
                    bShouldDel = False
                End If
            Next j
            If bShouldDel = True Then
                nDelCount = nDelCount + 1
                ReDim Preserve aszDel(1 To nDelCount)
                aszDel(nDelCount) = g_aszFunOld(i)
            End If
            bShouldDel = True
        Next i
        If ArrayLength(aszDel) <> 0 Then
            oGroup.Identify szGroupID

            For i = 1 To ArrayLength(aszDel)
                If aszDel(i) <> "" Then
                oGroup.DeleteFunction aszDel(i)

                End If
            Next i
        End If

        '增加权限
        For i = 1 To nLen
            For j = 1 To nLenOld
                If g_atInBrowse(i).FunID = g_aszFunOld(j) Then
                    bShouldAdd = False
                End If
            Next j
            If bShouldAdd = True Then
                nAddCount = nAddCount + 1
                ReDim Preserve aszAdd(1 To nAddCount)
                aszAdd(nAddCount) = g_atInBrowse(i).FunID
            End If
            bShouldAdd = True
        Next i
        If ArrayLength(aszAdd) <> 0 Then
            oGroup.Identify szGroupID
            For i = 1 To ArrayLength(aszAdd)
                If aszAdd(i) <> "" Then
                oGroup.AddFunction (aszAdd(i))
                End If
            Next i
        End If
    End If
    ReDim g_aszFunOld(1 To nLen)
    For i = 1 To nLen
        g_aszFunOld(i) = g_atInBrowse(i).FunID
    Next i
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub
