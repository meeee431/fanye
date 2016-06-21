VERSION 5.00
Begin VB.Form frmGrant 
   BackColor       =   &H00FFFFFF&
   Caption         =   "按功能或功能组授权"
   ClientHeight    =   4005
   ClientLeft      =   1335
   ClientTop       =   2190
   ClientWidth     =   6570
   HelpContextID   =   5003201
   Icon            =   "frmGrant.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6570
   Begin VB.ListBox lstU 
      Appearance      =   0  'Flat
      Height          =   3270
      Left            =   105
      MultiSelect     =   2  'Extended
      TabIndex        =   6
      Top             =   600
      Width           =   2505
   End
   Begin VB.ListBox lstG 
      Appearance      =   0  'Flat
      Height          =   3270
      Left            =   2730
      MultiSelect     =   2  'Extended
      TabIndex        =   5
      Top             =   600
      Width           =   2505
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   300
      Left            =   5370
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelect 
      Caption         =   "删除(&D)"
      Height          =   300
      Left            =   5370
      TabIndex        =   1
      Top             =   990
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "添加(&A)"
      Height          =   300
      Left            =   5370
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "关闭(&C)"
      Height          =   300
      Left            =   5370
      TabIndex        =   2
      Top             =   1395
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户组:"
      Height          =   180
      Left            =   2730
      TabIndex        =   8
      Top             =   345
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户:"
      Height          =   180
      Left            =   105
      TabIndex        =   7
      Top             =   345
      Width           =   450
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "已授权的用户或用户组"
      Height          =   180
      Left            =   105
      TabIndex        =   4
      Top             =   90
      Width           =   1890
   End
End
Attribute VB_Name = "frmGrant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bFun As Boolean

Public szFunCode As String
Dim szFunName As String

Public szFunGroup As String
Dim aszFunCode() As String

Dim aszUserDel() As String
Dim aszUserGroupDel() As String


Private Sub cmdAdd_Click()
    Dim oUser As New User
    Dim oGroup As New UserGroup
    Dim nlenU As Integer, nlenUOld As Integer
    Dim nLenG As Integer, nLenGOld As Integer
    Dim i As Integer, j As Integer
    Dim bIsNull As Boolean
    Dim nFunsCount As Integer
    frmAddForFun.Show vbModal
    If bFun = False Then
        nFunsCount = ArrayLength(aszFunCode)
    End If
    nlenU = ArrayLength(g_aszUserAdd, 2)
    nLenG = ArrayLength(g_aszUserGroupAdd, 2)
    nLenGOld = ArrayLength(g_aszUserGroup)
    nlenUOld = ArrayLength(g_aszUser)
    bIsNull = False '标识原用户非空
    If nlenU > 0 Then
    If g_aszUserAdd(1, 1) <> "" Then
        If g_aszUser(1) = "" Then
            ReDim g_aszUser(1 To nlenU)
            bIsNull = True
        Else
            ReDim Preserve g_aszUser(1 To (nlenU + nlenUOld))
        End If
        
        On Error GoTo ErrorHandle
        oUser.Init g_oActUser
        For i = 1 To nlenU
            If bFun = True Then '按功能授权给用户授权
                oUser.Identify g_aszUserAdd(1, i)
                oUser.AddFunction szFunCode
                
            Else '按功能组给用户组授权
                If nFunsCount > 0 Then
                    If aszFunCode(1) <> "" Then
                        oUser.Identify g_aszUserAdd(1, i)
                        For j = 1 To nFunsCount
                             '授权可能重*************
                            oUser.AddFunction aszFunCode(j)
                        Next j
                    End If
                End If
            End If
            lstU.AddItem g_aszUserAdd(1, i) & "[" & g_aszUserAdd(2, i) & "]"
            If bIsNull = True Then '标识原用户非空
                g_aszUser(i) = g_aszUserAdd(1, i)
            Else
                g_aszUser(i + nlenUOld) = g_aszUserAdd(1, i)
            End If
        Next i
    End If
    End If
    
    bIsNull = False
    If nLenG > 0 Then
    If g_aszUserGroupAdd(1, 1) <> "" Then
        
        If g_aszUserGroup(1) = "" Then
            ReDim g_aszUserGroup(1 To nLenG)
            bIsNull = True
        Else
            ReDim Preserve g_aszUserGroup(1 To (nLenG + nLenGOld))
        End If
        oGroup.Init g_oActUser
        For i = 1 To nLenG
            If bFun = True Then '按功能给用户组授权
                oGroup.Identify g_aszUserGroupAdd(1, i)
                oGroup.AddFunction szFunCode
                
            Else '按功能组给用户组授权
                If nFunsCount > 0 Then
                    If aszFunCode(1) <> "" Then
                        oGroup.Identify g_aszUserGroupAdd(1, i)
                        For j = 1 To nFunsCount
                             '授权可能重
                            oGroup.AddFunction aszFunCode(j)
                        Next j
                    End If
                End If
            End If
            lstG.AddItem g_aszUserGroupAdd(1, i) & "[" & g_aszUserGroupAdd(2, i) & "]"
            If bIsNull = True Then
                g_aszUserGroup(i) = g_aszUserGroupAdd(1, i)
            Else
                g_aszUserGroup(i + nLenGOld) = g_aszUserGroupAdd(1, i)
            End If
        Next i
    End If
    End If
Exit Sub
ErrorHandle:
    ShowErrorMsg

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelect_Click()
    Dim i As Integer, j As Integer
    Dim nLenUsers As Integer, k As Integer
    Dim oUser As New User, oGroup As New UserGroup
    If bFun = True Then
        i = MsgBox("确认收回选中的用户和用户组的" & szFunCode & "[" & szFunName & "]" & "权限!", vbYesNo + vbQuestion, cszMsg)
    Else
        i = MsgBox("确认收回选中的用户和用户组的" & "[" & szFunGroup & "]" & "权限组!", vbYesNo + vbQuestion, cszMsg)
    End If
    If i = vbNo Then Exit Sub
ReGetSelectedInfo:
    If lstU.SelCount > 0 Then
        ReDim aszUserDel(1 To lstU.SelCount)
    End If
    If lstG.SelCount > 0 Then
        ReDim aszUserGroupDel(1 To lstG.SelCount)
    End If
    '得到应删除此功能的用户和组
    nLenUsers = ArrayLength(g_atUserInfo)
    j = 0
    For i = 0 To lstU.ListCount - 1
        If lstU.Selected(i) = True Then
            lstU.ListIndex = i
            For k = 1 To nLenUsers
                If g_atUserInfo(k).UserID = PartCode(lstU.Text) Then
                    If g_atUserInfo(k).InnerUser = False Then
                        j = j + 1
                        aszUserDel(j) = PartCode(lstU.Text)
                    Else
                        MsgBox lstU.Text & "为内置用户,不能修改其权限!", vbInformation, cszMsg
                        lstU.Selected(i) = False
                        GoTo ReGetSelectedInfo
                    End If
                    Exit For
                End If
            Next k
        End If
    Next i
    
    j = 0
    For i = 0 To lstG.ListCount - 1
        If lstG.Selected(i) = True Then
            lstG.ListIndex = i
            j = j + 1
            aszUserGroupDel(j) = PartCode(lstG.Text)
        End If
    Next i
    
    '删除数据库中的对应信息
    SetBusy
    DoDelUser_Group
    SetNormal
    
    
    'lstU刷新
    j = 0
    For i = 0 To lstU.ListCount - 1
        If lstU.Selected(i - j) = True Then
            lstU.RemoveItem (i - j)
            j = j + 1
        End If
    Next i
    
    'lstG刷新
    j = 0
    For i = 0 To lstG.ListCount - 1
        If lstG.Selected(i - j) = True Then
            lstG.RemoveItem (i - j)
            j = j + 1
        End If
    Next i
    
    ReFreshArray
    
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me, content
End Sub

Private Sub Form_Load()
    Dim nLen As Integer, i As Integer
    Dim nLen1 As Integer, j As Integer
    Dim szTemp As String

'FL 2002-12-16

'    If bFun = False Then
'
'        nLen = ArrayLength(g_atAllFun)
'
'        If nLen > 0 Then
'            For i = 1 To nLen
'                If g_atAllFun(i).szFunctionGroup = g_alvItemText2(1) Then
'                    j = j + 1
'                    ReDim Preserve aszFunCode(1 To j)
'                    aszFunCode(j) = g_atAllFun(i).szFunctionCode
'                End If
'            Next i
'        End If
'    End If
'FL

    On Error GoTo ErrorHandle

     '得到已授权的用户和用户组
    If bFun = True Then '按功能
        Dim oFun As New COMFunction
        oFun.Init g_oActUser
        oFun.Identify szFunCode
        szFunName = oFun.FunctionName
        'g_aszUser = oFun.GetAllUser
        g_aszUser = oFun.GetDirectUser
        g_aszUserGroup = oFun.GetAllUserGroup
        Me.Caption = "按功能:" & szFunCode & "[" & szFunName & "]授权"
    Else '按功能组
        g_aszUser = g_oSysMan.GetFunGroupGranted(aszFunCode)
        g_aszUserGroup = g_oSysMan.GetFunGroupGranted(aszFunCode, False)
        
        Me.Caption = "按功能组:[" & szFunGroup & "]授权"
    End If
    '得到用户名并显示
    nLen = ArrayLength(g_aszUser)
    nLen1 = 0
    
    nLen1 = ArrayLength(g_atUserInfo)
    
    If nLen > 0 And nLen1 > 0 Then
        For i = 1 To nLen
            szTemp = g_aszUser(i)
            For j = 1 To nLen1
                If g_atUserInfo(j).UserID = g_aszUser(i) Then
                    szTemp = g_aszUser(i) & "[" & g_atUserInfo(j).UserName & "]"
                    Exit For
                End If
            Next j
           lstU.AddItem (szTemp)

        Next i
    End If
    
    '得到用户组名并显示
    nLen = ArrayLength(g_aszUserGroup)
    nLen1 = ArrayLength(g_atUserGroupInfo)
    

    If nLen > 0 And nLen1 > 0 Then
        For i = 1 To nLen
            szTemp = g_aszUserGroup(i)
            For j = 1 To nLen1
                If g_atUserGroupInfo(j).UserGroupID = g_aszUserGroup(i) Then
                    szTemp = g_aszUserGroup(i) & "[" & g_atUserGroupInfo(j).GroupName & "]"
                    Exit For
                End If
            Next j
            lstG.AddItem (szTemp)
        Next i
    End If


Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub



Private Sub DoDelUser_Group()
    Dim i As Integer, j As Integer
    Dim nLen As Integer
    Dim oUser As New User
    Dim oGroup As New UserGroup
    On Error GoTo ErrorHandle
    If bFun = True Then
        oUser.Init g_oActUser
        For i = 1 To lstU.SelCount
            oUser.Identify aszUserDel(i)
            oUser.DeleteFunction szFunCode
        
            
        Next i
        oGroup.Init g_oActUser
        For i = 1 To lstG.SelCount
            oGroup.Identify aszUserGroupDel(i)
            oGroup.DeleteFunction szFunCode
            
        Next i
            
    Else
        nLen = ArrayLength(aszFunCode)
        
        oUser.Init g_oActUser
        
        For i = 1 To lstU.SelCount
            oUser.Identify aszUserDel(i)
            
            If nLen > 0 Then
                If aszFunCode(1) <> "" Then
                    For j = 1 To nLen
                        oUser.DeleteFunction aszFunCode(j)
                    Next j
                End If
            End If
        Next i
        oGroup.Init g_oActUser
        For i = 1 To lstG.SelCount
            oGroup.Identify aszUserGroupDel(i)
            If nLen > 0 Then
                If aszFunCode(1) <> "" Then
                    For j = 1 To nLen
                        oGroup.DeleteFunction aszFunCode(j)
                    Next j
                End If
            End If
        Next i
    End If
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub


Private Sub ReFreshArray()
    Dim nlenU As Integer, nLenG As Integer
    Dim nlenUOld As Integer, nLenGOld As Integer
    Dim i As Integer
    Dim j As Integer
    Dim aszTemp() As String
    Dim bExist As Boolean
    Dim nTemp As Integer
    nlenU = ArrayLength(aszUserDel)
    nLenG = ArrayLength(aszUserGroupDel)
    nlenUOld = ArrayLength(g_aszUser)
    nLenGOld = ArrayLength(g_aszUserGroup)
    If nlenU > 0 Then
        If aszUserDel(1) <> "" Then
            For i = 1 To nlenUOld
                For j = 1 To nlenU
                    If g_aszUser(i) = aszUserDel(j) Then
                        bExist = True
                        Exit For
                    End If
                Next j
                If bExist = True Then
                    bExist = False
                Else
                    nTemp = nTemp + 1
                    ReDim Preserve aszTemp(1 To nTemp)
                    aszTemp(nTemp) = g_aszUser(i)
                End If
            Next i
            ReDim g_aszUser(1 To nTemp)
            g_aszUser = aszTemp
        End If
    End If
    If nLenG > 0 Then
        If aszUserGroupDel(1) <> "" Then
            nTemp = 0
            bExist = False
            ReDim aszTemp(1 To 1)
            For i = 1 To nLenGOld
                For j = 1 To nLenG
                    If g_aszUserGroup(i) = aszUserGroupDel(j) Then
                        bExist = True
                        Exit For
                    End If
                Next j
                If bExist = True Then
                    bExist = False
                Else
                    nTemp = nTemp + 1
                    ReDim Preserve aszTemp(1 To nTemp)
                    aszTemp(nTemp) = g_aszUserGroup(i)
                End If
            Next i
            ReDim g_aszUserGroup(1 To nTemp)
            g_aszUserGroup = aszTemp
        End If
    End If
    ReDim aszUserDel(1)
    ReDim aszUserGroupDel(1)
End Sub

Private Sub lstG_DblClick()
    lstG.Selected(lstG.ListIndex) = Not lstG.Selected(lstG.ListIndex)
End Sub

Private Sub lstU_dblClick()
    lstU.Selected(lstU.ListIndex) = Not lstU.Selected(lstU.ListIndex)
End Sub


