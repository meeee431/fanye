VERSION 5.00
Begin VB.Form frmAEGroup 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmAEGroup"
   ClientHeight    =   3720
   ClientLeft      =   2085
   ClientTop       =   3495
   ClientWidth     =   5805
   HelpContextID   =   5000201
   Icon            =   "frmAEGroup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   Begin PSTSMan.AddDel adUser 
      Height          =   2505
      Left            =   60
      TabIndex        =   10
      Top             =   1290
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   4419
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LeftLabel       =   "未选用户(&L)"
      RightLabel      =   "已选用户(&R)"
      ButtonWidth     =   1215
      ButtonHeight    =   315
   End
   Begin VB.CommandButton cmdGroupRight 
      Caption         =   "权限(&R)"
      Height          =   315
      Left            =   4485
      TabIndex        =   6
      Top             =   915
      Width           =   1215
   End
   Begin VB.TextBox txtGroupID 
      Height          =   330
      Left            =   1470
      TabIndex        =   0
      Top             =   135
      Width           =   2445
   End
   Begin VB.TextBox txtGroupAnnotation 
      Height          =   315
      Left            =   1470
      TabIndex        =   2
      Top             =   930
      Width           =   2445
   End
   Begin VB.TextBox txtGroupName 
      Height          =   315
      Left            =   1470
      TabIndex        =   1
      Top             =   540
      Width           =   2445
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   315
      Left            =   4485
      TabIndex        =   5
      Top             =   510
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   315
      Left            =   4485
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "注释(&A):"
      Height          =   180
      Left            =   105
      TabIndex        =   3
      Top             =   975
      Width           =   720
   End
   Begin VB.Label lblGroupID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1470
      TabIndex        =   9
      Top             =   210
      Width           =   2445
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户组名(&F):"
      Height          =   180
      Left            =   90
      TabIndex        =   8
      Top             =   615
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户组代码(&U):"
      Height          =   180
      Left            =   105
      TabIndex        =   7
      Top             =   225
      Width           =   1260
   End
End
Attribute VB_Name = "frmAEGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  *******************************************************************
' *  Source File Name  : frmAEGroup                                    *
' *  Project Name: PSTSMan                                    *
' *  Engineer:                                              *
' *  Date Generated: 2002/08/19                      *
' *  Last Revision Date : 2002/08/19             *
' *  Brief Description   : 添加用户组或编辑用户组属性                            *
' *******************************************************************


Option Explicit
Option Base 1
Public bEdit As Boolean
Public bRightRead As Boolean
Dim aszIncUser() As String
Dim szGroupID As String
Dim szGroupName As String
Dim szGroupAnn As String




Private Sub cmdCancel_Click()
    Unload Me

    
    
    '*****清空下列值
    ReDim g_atAuthored(1)
    ReDim g_atAddBrowse(1)
    ReDim g_atInBrowse(1)
    ReDim g_aszFunOld(1)
    
    ReDim g_atExcludUser(1)
    ReDim g_atIncludUser(1)
    ReDim g_atIncludUserOld(1)
    
    g_bBrowseNull = False

End Sub

Private Sub cmdGroupRight_Click()
    If bEdit = False Then
        If txtGroupID.Text = "" Then
            MsgBox "请输入用户组代码,重试.", vbInformation, cszMsg
        Else
            frmUser_GroupRight.m_bUser = False
            frmUser_GroupRight.Show vbModal, Me
        End If
    Else
            frmUser_GroupRight.m_bUser = False
            frmUser_GroupRight.Show vbModal, Me
    End If
End Sub

Private Sub cmdOK_Click()
    Dim oGroup As New UserGroup
    Dim nLen As Integer, i As Integer, j As Integer
    Dim nLenOld As Integer
    Dim bShouldDel As Boolean, bShouldAdd As Boolean
    Dim nAddCount As Integer, nDelCount As Integer
    Dim aszDel() As String
    Dim aszAdd() As String
    '得到信息
    szGroupAnn = txtGroupAnnotation.Text
    szGroupName = txtGroupName
    
    GetInfoFromUI

    If bEdit = True Then
        szGroupID = lblGroupID.Caption
        '更新
        On Error GoTo ErrorHandle
        oGroup.Init g_oActUser
            oGroup.Identify szGroupID
            oGroup.Annotation = szGroupAnn
            oGroup.UserGroupName = szGroupName
            oGroup.Update
        '所属用户
        On Error GoTo 0
        nLen = 0
        On Error Resume Next
        nLen = UBound(g_atIncludUser)
        
        
        nLenOld = 0
        
        nLenOld = UBound(g_atIncludUserOld)
        On Error GoTo 0
        On Error GoTo ErrorHandle
        If nLen = 0 And nLenOld = 0 Then
            'do nothing
        ElseIf nLen = 0 Then
            '只删除
            If g_atIncludUserOld(1).UserID <> "" Then
                For i = 1 To nLenOld
                    oGroup.DeleteUser g_atIncludUserOld(i).UserID
                Next i
            End If
        ElseIf nLenOld = 0 Then
            '只新增
            If g_atIncludUser(1).UserID <> "" Then
                For i = 1 To nLen
                    oGroup.AddUser g_atIncludUser(i).UserID
                Next i
                
            End If
        Else
            '有新增,删除
            bShouldDel = True
            bShouldAdd = True
            nAddCount = 0
            nDelCount = 0
            '删除用户
            For i = 1 To nLenOld
                For j = 1 To nLen
                    If g_atIncludUserOld(i).UserID = g_atIncludUser(j).UserID Then
                        bShouldDel = False
                    End If
                Next j
                If bShouldDel = True Then
                    nDelCount = nDelCount + 1
                    ReDim Preserve aszDel(1 To nDelCount)
                    aszDel(nDelCount) = g_atIncludUserOld(i).UserID
                End If
                bShouldDel = True
            Next i
            If ArrayLength(aszDel) <> 0 Then
                For i = 1 To ArrayLength(aszDel)
                    If aszDel(i) <> "" Then
                    oGroup.Identify szGroupID
                    oGroup.DeleteUser aszDel(i)
                    
                    End If
                Next i
            End If
            '增加组
            For i = 1 To nLen
                For j = 1 To nLenOld
                    If g_atIncludUser(i).UserID = g_atIncludUserOld(j).UserID Then
                        bShouldAdd = False
                    End If
                Next j
                If bShouldAdd = True Then
                    nAddCount = nAddCount + 1
                    ReDim Preserve aszAdd(1 To nAddCount)
                    aszAdd(nAddCount) = g_atIncludUser(i).UserID
                End If
                bShouldAdd = True
            Next i
            '修改到数据库
            If ArrayLength(aszAdd) <> 0 Then
                For i = 1 To ArrayLength(aszAdd)
                    If aszAdd(i) <> "" Then
                    oGroup.Identify szGroupID
                    oGroup.AddUser aszAdd(i)
                    
                    End If
                Next i
            End If
        End If

        '修改后数据刷新
        nLen = ArrayLength(g_atUserGroupInfo)
'        If bInnerGroup = False Then
            For i = 1 To nLen
                If g_atUserGroupInfo(i).UserGroupID = szGroupID Then
                    g_atUserGroupInfo(i).Annotation = szGroupAnn
                    g_atUserGroupInfo(i).GroupName = szGroupName
                End If
            Next i
'        End If
        frmStoreMenu.DisplayGroupInfo (nLen)

        
        
    Else '************新增
        szGroupID = txtGroupID.Text
        If szGroupID = "" Then
            MsgBox "请输入组名,重试.", vbInformation, cszMsg
            Exit Sub
        End If
        On Error GoTo ErrorHandle
        oGroup.Init g_oActUser
        oGroup.AddNew
        oGroup.Annotation = szGroupAnn
        oGroup.UserGroupID = szGroupID
        oGroup.UserGroupName = szGroupName
        oGroup.Update
        
        On Error GoTo 0
        nLen = 0
        On Error Resume Next
        nLen = UBound(g_atIncludUser)
        On Error GoTo 0
        On Error GoTo ErrorHandle
        If nLen <> 0 Then
            If g_atIncludUser(1).UserID <> "" Then
                For i = 1 To nLen
                    oGroup.AddUser g_atIncludUser(i).UserID
                Next i
                
            
            End If
        End If
        ''权限***********
        nLen = 0
        On Error GoTo 0
        On Error Resume Next
        nLen = UBound(g_atInBrowse)
        On Error GoTo 0
        On Error GoTo ErrorHandle
        If (nLen <> 0) And (g_bBrowseNull = False) Then
            If g_atInBrowse(1).FunID <> "" Then
                For i = 1 To nLen
                
                    
                    oGroup.Identify szGroupID
                    oGroup.AddFunction (g_atInBrowse(i).FunID)
                    
                    
                Next i
            End If
        End If
        '新增用户组后主窗体显示的刷新
        i = ArrayLength(g_atUserGroupInfo) + 1
        ReDim Preserve g_atUserGroupInfo(1 To i)
        g_atUserGroupInfo(i).Annotation = szGroupAnn
        g_atUserGroupInfo(i).GroupName = szGroupName
        g_atUserGroupInfo(i).InnerGroup = True
        g_atUserGroupInfo(i).UserGroupID = szGroupID
        frmStoreMenu.DisplayGroupInfo (i)
    End If
there:
'****************善后
    For i = 1 To frmSMCMain.lvDetail2.ListItems.Count
        If frmSMCMain.lvDetail2.ListItems.Item(i).Key = "A" & oGroup.UserGroupID Then
            frmSMCMain.lvDetail2.ListItems.Item(i).Selected = True
        Else
            frmSMCMain.lvDetail2.ListItems.Item(i).Selected = False
        End If
    Next
    '*****清空下列值
    ReDim g_atAuthored(1)
    ReDim g_atAddBrowse(1)
    ReDim g_atInBrowse(1)
    ReDim g_aszFunOld(1)
    
    ReDim g_atExcludUser(1)
    ReDim g_atIncludUser(1)
    ReDim g_atIncludUserOld(1)
    
    g_bBrowseNull = False
        

    Unload Me
        
        
    Exit Sub
ErrorHandle:
    ShowErrorMsg
    GoTo there
End Sub

Private Sub Form_Activate()
    adUser.SetFocus
End Sub

Private Sub Form_Load()
    bRightRead = False

    Me.Top = (Screen.Height - Me.ScaleHeight) / 2
    Me.Left = (Screen.Width - Me.ScaleWidth) / 2
    
    If bEdit Then
        Me.Caption = "修改用户组属性"
        cmdOk.Caption = "确定"
        cmdCancel.Caption = "取消"
        txtGroupID.Visible = False
        lblGroupID.Visible = True
        LoadData
        DisplayadUser
    Else
        Me.Caption = "新增用户组"
        cmdOk.Caption = "新增"
        cmdCancel.Caption = "关闭"
        lblGroupID.Visible = False
        txtGroupID.Visible = True
        LoadDateNew
        DisplayadUser
    End If
    
End Sub

Private Sub txtGroupAnnotation_Validate(Cancel As Boolean)
    If TextLongValidate(255, txtGroupAnnotation.Text) Then Cancel = True
End Sub

Private Sub txtGroupID_Validate(Cancel As Boolean)
    If TextLongValidate(20, txtGroupID.Text) Then Cancel = True
    If SpacialStrValid(txtGroupID.Text, "[") Then Cancel = True
    If SpacialStrValid(txtGroupID.Text, ",") Then Cancel = True
    If SpacialStrValid(txtGroupID.Text, "]") Then Cancel = True

End Sub

Private Sub txtGroupName_Validate(Cancel As Boolean)
    If TextLongValidate(50, txtGroupName.Text) Then Cancel = True
    If SpacialStrValid(txtGroupName.Text, "[") Then Cancel = True
    If SpacialStrValid(txtGroupName.Text, ",") Then Cancel = True
    If SpacialStrValid(txtGroupName.Text, "]") Then Cancel = True

End Sub

'读入数据
Private Sub LoadData()
    Dim oGroup As New UserGroup
    Dim i As Integer, nLen3 As Integer
    Dim nLen As Integer, bInc As Boolean
    Dim nLen1 As Integer, nLen2 As Integer
    Dim j As Integer
    
    
    lblGroupID.Caption = g_alvItemText2(1)
    For i = 1 To ArrayLength(g_atUserGroupInfo)
        If g_atUserGroupInfo(i).UserGroupID = g_alvItemText2(1) Then '取得内存中对应用户组信息
            txtGroupName.Text = g_atUserGroupInfo(i).GroupName '用户组名
            txtGroupAnnotation.Text = g_atUserGroupInfo(i).Annotation '用户组注释
        End If
    Next i

    On Error GoTo ErrorHandle
    oGroup.Init g_oActUser
    oGroup.Identify lblGroupID
    aszIncUser = oGroup.GetAllUser
    nLen = ArrayLength(g_atUserInfo)
    nLen1 = ArrayLength(aszIncUser)
    nLen2 = 0
    nLen3 = 0
    bInc = False
    If nLen <> 0 Then
        For i = 1 To nLen
            For j = 1 To nLen1
            If g_atUserInfo(i).UserID = aszIncUser(j) Then
                bInc = True
                nLen2 = nLen2 + 1
                ReDim Preserve g_atIncludUser(1 To nLen2)
                g_atIncludUser(nLen2).UserID = aszIncUser(j)
                g_atIncludUser(nLen2).UserName = g_atUserInfo(i).UserName
            End If
            Next j
            If bInc = False Then
                nLen3 = nLen3 + 1
'                If nLen3 <> 1 Then
                ReDim Preserve g_atExcludUser(1 To nLen3)
'                End If
                g_atExcludUser(nLen3).UserID = g_atUserInfo(i).UserID
                g_atExcludUser(nLen3).UserName = g_atUserInfo(i).UserName
            End If
            bInc = False
        Next i
    End If
    g_atIncludUserOld = g_atIncludUser
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub DisplayadUser()
    Dim nLeft As Integer, nRight As Integer
    Dim szTemp As String, i As Integer
    
    nLeft = 0
    On Error Resume Next
    nLeft = UBound(g_atExcludUser)
    nRight = 0
    nRight = UBound(g_atIncludUser)
    On Error GoTo 0
    If nLeft <> 0 Then
        For i = 1 To nLeft
            If g_atExcludUser(i).UserID <> "" Then
                szTemp = g_atExcludUser(i).UserID & "[" & g_atExcludUser(i).UserName & "]"
                adUser.AddData szTemp
            End If
        Next i
    End If
    If nRight <> 0 Then
        For i = 1 To nRight
            If g_atIncludUser(i).UserID <> "" Then
                szTemp = g_atIncludUser(i).UserID & "[" & g_atIncludUser(i).UserName & "]"
                adUser.AddData szTemp, False
            End If
        Next i
        
    End If
    
End Sub

Private Sub GetInfoFromUI()
    Dim aszTempLeft As Variant
    Dim asztempRight As Variant
    Dim i As Integer, j As Integer
    Dim nLeft As Integer, nRight As Integer
    
    aszTempLeft = adUser.LeftData
    asztempRight = adUser.RightData
    nLeft = ArrayLength(aszTempLeft)
    If nLeft = 0 Then
        ReDim g_atExcludUser(1 To 1)
    Else
        ReDim g_atExcludUser(1 To nLeft)
        For i = 1 To nLeft
            g_atExcludUser(i).UserID = PartCode(CStr(aszTempLeft(i)), True)
            g_atExcludUser(i).UserName = PartCode(CStr(aszTempLeft(i)), False)
        Next i
    End If
    nRight = ArrayLength(asztempRight)
    If nRight = 0 Then
        ReDim g_atIncludUser(1 To 1)
    Else
        ReDim g_atIncludUser(1 To nRight)
        For i = 1 To nRight
            g_atIncludUser(i).UserID = PartCode(CStr(asztempRight(i)))
            g_atIncludUser(i).UserName = PartCode(CStr(asztempRight(i)), False)
        Next i
    End If
    
End Sub

Private Sub LoadDateNew()
    Dim nLen As Integer
    Dim i As Integer
    
    ReDim g_atIncludUser(1)
    ReDim g_atIncludUserOld(1)
    
    nLen = 0
    
    nLen = ArrayLength(g_atUserInfo)
    
    
    
    If nLen <> 0 Then
        ReDim g_atExcludUser(1 To nLen)
        For i = 1 To nLen
            g_atExcludUser(i).UserID = g_atUserInfo(i).UserID
            g_atExcludUser(i).UserName = g_atUserInfo(i).UserName
        Next i
    End If
    
End Sub
