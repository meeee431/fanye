VERSION 5.00
Begin VB.Form frmUserBeGroup 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "用户所属组"
   ClientHeight    =   3495
   ClientLeft      =   1215
   ClientTop       =   2190
   ClientWidth     =   6315
   HelpContextID   =   50000320
   Icon            =   "frmUserBeGroup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   315
      Left            =   5160
      TabIndex        =   6
      Top             =   3120
      Width           =   1015
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   315
      Left            =   3990
      TabIndex        =   1
      Top             =   3120
      Width           =   1015
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   315
      Left            =   2805
      TabIndex        =   0
      Top             =   3105
      Width           =   1015
   End
   Begin PSTSMan.AddDel adGroup 
      Height          =   2430
      Left            =   45
      TabIndex        =   7
      Top             =   675
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   4286
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
      ButtonWidth     =   1215
      ButtonHeight    =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户所属组"
      Height          =   180
      Left            =   105
      TabIndex        =   8
      Top             =   420
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   1095
      X2              =   6245
      Y1              =   525
      Y2              =   525
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   1095
      X2              =   6245
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label lblUserName 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3885
      TabIndex        =   2
      Top             =   135
      Width           =   90
   End
   Begin VB.Label lblUserID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1290
      TabIndex        =   5
      Top             =   135
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户名(&N):"
      Height          =   180
      Left            =   2835
      TabIndex        =   4
      Top             =   135
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户代码(&U):"
      Height          =   180
      Left            =   105
      TabIndex        =   3
      Top             =   135
      Width           =   1080
   End
End
Attribute VB_Name = "frmUserBeGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *******************************************************************
' *  Source File Name  : frmUserBeGroup                             *
' *  Project Name: PSTSMan                                    *
' *  Engineer:                               *
' *  Date Generated: 2002/08/19                      *
' *  Last Revision Date : 2002/08/19             *
' *  Brief Description   : 给用户分配组                             *
' *******************************************************************

Option Explicit
Option Base 1



Private Sub cmdCancel_Click()
'    frmAEUser.bGroupRead = True
'    frmAEUser.bGroupChange = False
    '''读数据存于内存
    '.......
'    GetInfofromForm
    
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me, content
End Sub

Private Sub cmdOK_Click()
    SetBusy
    frmAEUser.bGroupRead = True
'    frmAEUser.bGroupChange = True
    '''''
    '读数据存于内存
    GetInfofromForm

    If frmAEUser.bEdit = True Then
        ModifyGroup
        
    End If

    SetNormal
    Unload Me

End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.ScaleHeight) / 2
    Me.Left = (Screen.Width - Me.ScaleWidth) / 2

    
    
    If frmAEUser.bEdit = True Then
        lblUserID.Caption = frmAEUser.lblUserID.Caption
    Else
        lblUserID.Caption = frmAEUser.txtUserID.Text
    End If
    lblUserName.Caption = frmAEUser.txtUserName.Text
    
    
    Dim LenTemp As Integer
    LenTemp = ArrayLength(g_atUserGroupInfo)
    Dim i As Integer
    ReDim g_atAllGroup(1 To LenTemp)
    For i = 1 To LenTemp
        g_atAllGroup(i).GroupID = g_atUserGroupInfo(i).UserGroupID
        g_atAllGroup(i).GroupName = g_atUserGroupInfo(i).GroupName
    Next i
    
    
    If frmAEUser.bEdit = False Then
        If frmAEUser.bGroupRead = False Then
            g_atUnBelongGroup = g_atAllGroup
            ReDim g_atBelongGroup(1)
            g_bRightNull = True
        End If
    Else
'        If frmAEUser.bGroupRead = False Then
            GetGroupInfoForADGroup
'        End If
    End If
    
'    Dim aszTemp(1 To 2) As String
'    aszTemp(1) = "用户组代码"
'    aszTemp(2) = "用户组名"
'    adGroup.ColumnHeaders = aszTemp
    ShowAdGroup
    
    

    
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub
'显示组信息
Private Sub ShowAdGroup()
    Dim nNum1 As Integer '标注左边纪录数
    Dim nNum2 As Integer '标注右边纪录数
'    Dim aszTemp(1 To 2) As String, i As Integer  'adGroup.AddData方法调用
    Dim szTemp As String, i As Integer
    
    On Error Resume Next
    If g_bRightNull = False Then
    
        nNum2 = UBound(g_atBelongGroup)
    Else
        nNum2 = 0
    End If
    If g_bLeftNull = False Then
        nNum1 = UBound(g_atUnBelongGroup)
    Else
        nNum1 = 0
    End If
    On Error GoTo 0
    If nNum1 <> 0 Then '加载左边数据
        For i = 1 To nNum1
            If g_atUnBelongGroup(i).GroupID <> "" Then
'            aszTemp(1) = g_atUnBelongGroup(i).GroupID
'            aszTemp(2) = g_atUnBelongGroup(i).GroupName
                szTemp = g_atUnBelongGroup(i).GroupID & "[" & g_atUnBelongGroup(i).GroupName & "]"
'            Call adGroup.AddData(aszTemp)
                adGroup.AddData szTemp
            End If
        Next i
    End If
    
    If nNum2 <> 0 Then '加在右边数据
        For i = 1 To nNum2
            If g_atBelongGroup(i).GroupID <> "" Then
'            aszTemp(1) = g_atBelongGroup(i).GroupID
'            aszTemp(2) = g_atBelongGroup(i).GroupName
'            Call adGroup.AddData(aszTemp, False)
                szTemp = g_atBelongGroup(i).GroupID & "[" & g_atBelongGroup(i).GroupName & "]"
                adGroup.AddData szTemp, False
            End If
        Next i
    End If
    
End Sub



Private Sub GetGroupInfoForADGroup()
    Dim oUserTemp As New User
    Dim nNum1 As Integer, nNum2 As Integer, nNum3 As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim bTemp As Boolean
    Dim aszBelongGroup() As String
    
    On Error GoTo ErrorHandle
    oUserTemp.Init g_oActUser
    oUserTemp.Identify frmAEUser.lblUserID.Caption
    aszBelongGroup = oUserTemp.GetAllGroup '得到数据库的所有组
    
    
    k = ArrayLength(aszBelongGroup)
    If k <> 0 Then
    ReDim g_atBelongGroup(1 To k) '转换数组元数据类型
    '信息显示******************
    On Error Resume Next
    nNum1 = 0
    nNum1 = UBound(g_atBelongGroup)
    nNum2 = 0
    nNum2 = UBound(g_atAllGroup)
    On Error GoTo 0
    For i = 1 To k
        For j = 1 To nNum2
            If g_atAllGroup(j).GroupID = aszBelongGroup(i) Then
                g_atBelongGroup(i).GroupID = aszBelongGroup(i)
                g_atBelongGroup(i).GroupName = g_atAllGroup(j).GroupName
            End If
        Next j
    Next i
    End If
    g_atBelongGroupOld = g_atBelongGroup '保存旧信息
    

    If nNum1 = 0 Then
        g_atUnBelongGroup = g_atAllGroup
        g_bRightNull = True 'adGroup右边为空
    ElseIf nNum1 = nNum2 Then
        g_bLeftNull = True 'adGroup左边为空
    Else
        bTemp = False
        nNum3 = 0
        For i = 1 To nNum2
            
            For j = 1 To nNum1
                If g_atAllGroup(i).GroupID = g_atBelongGroup(j).GroupID Then
                    bTemp = True
                End If
            Next j
            If bTemp = False Then
                nNum3 = nNum3 + 1
                ReDim Preserve g_atUnBelongGroup(1 To nNum3)
                g_atUnBelongGroup(nNum3).GroupID = g_atAllGroup(i).GroupID
                g_atUnBelongGroup(nNum3).GroupName = g_atAllGroup(i).GroupName
            End If
            bTemp = False
        Next i
    End If
Exit Sub
ErrorHandle:
    ShowErrorMsg
    
End Sub

Private Sub GetInfofromForm()
    Dim aszTempLeft As Variant
    Dim asztempRight As Variant
    Dim i As Integer, j As Integer
    
    aszTempLeft = adGroup.LeftData
    asztempRight = adGroup.RightData
    
    On Error GoTo ErrorHandle
    i = ArrayLength(aszTempLeft)
    ReDim g_atUnBelongGroup(1 To i)
    For j = 1 To i
'        g_atUnBelongGroup(j).GroupID = aszTempLeft(j, 1)
'        g_atUnBelongGroup(j).GroupName = aszTempLeft(j, 2)
        g_atUnBelongGroup(j).GroupID = PartCode(CStr(aszTempLeft(j)))
        g_atUnBelongGroup(j).GroupName = PartCode(CStr(aszTempLeft(j)), False)

    Next j
    g_bLeftNull = False
    
    
backhere:
    On Error GoTo there
    i = ArrayLength(asztempRight)
    ReDim g_atBelongGroup(1 To i)
    For j = 1 To i
        g_atBelongGroup(j).GroupID = PartCode(CStr(asztempRight(j)))
        g_atBelongGroup(j).GroupName = PartCode(CStr(asztempRight(j)), False)
    Next j
    g_bRightNull = False
    

Exit Sub
there:
    ReDim g_atBelongGroup(1)
    g_bRightNull = True
Exit Sub
ErrorHandle:
    ReDim g_atUnBelongGroup(1)
    g_bLeftNull = True
    GoTo backhere
End Sub


Private Sub ModifyGroup()
    Dim oUserTemp As New User
    Dim narrLenOld As Integer
    Dim narrLen As Integer
    Dim i As Integer, j As Integer, bShouldDel As Boolean, bShouldAdd As Boolean
    Dim nAddCount As Integer, nDelCount As Integer
    Dim aszDel() As String
    Dim aszAdd() As String
    Dim szUserID As String
    Dim oGroupTemp As New UserGroup
    
    
    szUserID = lblUserID
    On Error Resume Next
    narrLenOld = 0
    narrLenOld = UBound(g_atBelongGroupOld)
    narrLen = 0
    narrLen = UBound(g_atBelongGroup)
    On Error GoTo 0
    On Error GoTo ErrorHandle '修改
    oGroupTemp.Init g_oActUser
    If narrLenOld = 0 And narrLen = 0 Then
    ElseIf narrLen = 0 Then
        If g_atBelongGroupOld(1).GroupID <> "" Then
            For i = 1 To narrLenOld
                oGroupTemp.Identify g_atBelongGroupOld(i).GroupID
                oGroupTemp.DeleteUser szUserID
            
            Next i
        End If
    ElseIf narrLenOld = 0 Then
        If g_atBelongGroup(1).GroupID <> "" Then
            For i = 1 To narrLen
                oGroupTemp.Identify g_atBelongGroup(i).GroupID
                oGroupTemp.AddUser szUserID
            Next i
        End If
    Else
        bShouldDel = True
        bShouldAdd = True
        nAddCount = 0
        nDelCount = 0
        '删除组
        For i = 1 To narrLenOld
            For j = 1 To narrLen
                If g_atBelongGroupOld(i).GroupID = g_atBelongGroup(j).GroupID Then
                bShouldDel = False
                End If
            Next j
            If bShouldDel = True Then
            nDelCount = nDelCount + 1
            ReDim Preserve aszDel(1 To nDelCount)
            aszDel(nDelCount) = g_atBelongGroupOld(i).GroupID
            End If
            bShouldDel = True
        Next i
        If ArrayLength(aszDel) <> 0 Then
            For i = 1 To ArrayLength(aszDel)
                If aszDel(i) <> "" Then
                    oGroupTemp.Identify aszDel(i)
                    oGroupTemp.DeleteUser szUserID
                End If
            Next i
        End If
        '增加组
        For i = 1 To narrLen
            For j = 1 To narrLenOld
                If g_atBelongGroup(i).GroupID = g_atBelongGroupOld(j).GroupID Then
                    bShouldAdd = False
                End If
            Next j
            If bShouldAdd = True Then
                nAddCount = nAddCount + 1
                ReDim Preserve aszAdd(1 To nAddCount)
                aszAdd(nAddCount) = g_atBelongGroup(i).GroupID
            End If
            bShouldAdd = True
        Next i
        If ArrayLength(aszAdd) <> 0 Then
            For i = 1 To ArrayLength(aszAdd)
                If aszAdd(i) <> "" Then
                    '修改
                    oGroupTemp.Identify aszAdd(i)
                    oGroupTemp.AddUser szUserID
                End If
            Next i
        End If
    End If

Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

