VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAddRemoteUser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "添加远程登录用户"
   ClientHeight    =   2850
   ClientLeft      =   1485
   ClientTop       =   3075
   ClientWidth     =   6300
   HelpContextID   =   50000330
   Icon            =   "frmAddRemoteUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAnno 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1740
      TabIndex        =   6
      Top             =   450
      Width           =   3225
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   300
      Left            =   5115
      TabIndex        =   4
      Top             =   795
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   300
      Left            =   5115
      TabIndex        =   3
      Top             =   435
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   300
      Left            =   5115
      TabIndex        =   2
      Top             =   75
      Width           =   1095
   End
   Begin VB.TextBox txtRemoteUser 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      TabIndex        =   0
      Top             =   75
      Width           =   1485
   End
   Begin MSComctlLib.ListView lvLocalUser 
      Height          =   1680
      Left            =   90
      TabIndex        =   8
      Top             =   1065
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   2963
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "用户ID "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "用户名"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "选择本单位关联用户(&L):"
      Height          =   225
      Left            =   90
      TabIndex        =   7
      Top             =   825
      Width           =   3615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "远程用户注释(&A):"
      Height          =   180
      Left            =   90
      TabIndex        =   5
      Top             =   480
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "远程登录用户代码(对方单位提供)(&R):"
      Height          =   180
      Left            =   90
      TabIndex        =   1
      Top             =   120
      Width           =   3060
   End
End
Attribute VB_Name = "frmAddRemoteUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'添加远程登录用户

Option Explicit
Option Base 1


Dim aszAlllocalUser() As String
Dim aszUnionLocalUser() As String
Dim aszUnionLocalUserOld() As String

Dim szAnno As String
Dim szRemoteID As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me, content
End Sub

Private Sub cmdOK_Click()
    Dim oUnitTem As New Unit
    Dim nLen As Integer, nLen1 As Integer
    Dim i As Integer
    Dim nTemp As Integer
    
    GetInfoFromUI
    
    On Error GoTo ErrorHandle
    oUnitTem.Init g_oActUser
    oUnitTem.Identify g_alvItemText(1)
    nLen = ArrayLength(aszUnionLocalUser)
    If nLen = 0 Or aszUnionLocalUser(1) = "" Then
        nTemp = MsgBox("此远程账号,未与本地用户关联,重试?", vbYesNo + vbInformation, cszMsg)
        If nTemp = vbYes Then
            Exit Sub
        Else
        End If
    End If
    If frmUnitBeUser.bEditRemote = True Then '修改
        nLen1 = ArrayLength(aszUnionLocalUserOld)
        '修改代码及注释
        Call oUnitTem.ModifyRemoteUserAnnotation(szRemoteID, szAnno)
        If nLen1 <> 0 Then
            If aszUnionLocalUserOld(1) <> "" Then
                For i = 1 To nLen1
                Call oUnitTem.DetachUserToUnit(aszUnionLocalUserOld(i), szRemoteID)
                Next i
            End If
        End If
        '将用户加入进去
        If aszUnionLocalUser(1) <> "" Then
            For i = 1 To nLen
                Call oUnitTem.AttachUserToUnit(aszUnionLocalUser(i), szRemoteID)
            Next i
        End If
    Else '新增
        If szRemoteID = "" Then
            MsgBox "请输入远程用户代码", vbInformation, cszMsg
            Exit Sub
        Else
            '新增用户到该远程帐户中
            oUnitTem.Identify g_alvItemText(1)
            oUnitTem.AddRemoteUser szRemoteID, "", szAnno
            If nLen <> 0 Then
            If aszUnionLocalUser(1) <> "" Then
                For i = 1 To nLen
                    Call oUnitTem.AttachUserToUnit(aszUnionLocalUser(i), szRemoteID)
                Next i
            End If
            End If
        End If
    End If
    
    frmUnitBeUser.GetAndDisPlayRemote
    Unload Me
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.ScaleHeight) / 2
    Me.Left = (Screen.Width - Me.ScaleWidth) / 2
    Dim bTemp As Boolean
    Dim liTemp As ListItem
    Dim nLen As Integer, nLen1 As Integer, nLen2 As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim oUnitTem As New Unit
        
    '得到所有的本地用户信息
    On Error GoTo ErrorHandle
    oUnitTem.Init g_oActUser
    oUnitTem.Identify g_szLocalUnit
    aszAlllocalUser = oUnitTem.GetAllUser
'    Set oUnitTem = Nothing
    nLen = ArrayLength(aszAlllocalUser)
    nLen1 = ArrayLength(g_atUserInfo) '所有用户
    If frmUnitBeUser.bEditRemote = True Then
        Me.Caption = "编辑远程登录用户"
        txtRemoteUser.Locked = True
        txtRemoteUser.Text = frmUnitBeUser.szRemoteUserID
        txtAnno.Text = frmUnitBeUser.szAnno
        '得到关联的本地用户
        bTemp = GetAttachLocUser
        nLen2 = 0
        If bTemp = True Then
            ReDim aszUnionLocalUserOld(1)
        Else
            aszUnionLocalUserOld = aszUnionLocalUser
            nLen2 = ArrayLength(aszUnionLocalUser)
        End If
        If nLen2 = 0 Then
            For i = 1 To nLen1
                For j = 1 To nLen
                    If g_atUserInfo(i).UserID = aszAlllocalUser(j) Then
                        Set liTemp = lvLocalUser.ListItems.Add(, , g_atUserInfo(i).UserID)
                        liTemp.SubItems(1) = g_atUserInfo(i).UserName
                        liTemp.Checked = False
                    End If
                Next j
            Next i
        Else
            For i = 1 To nLen1
                 For j = 1 To nLen
                     If g_atUserInfo(i).UserID = aszAlllocalUser(j) Then
                         Set liTemp = lvLocalUser.ListItems.Add(, , g_atUserInfo(i).UserID)
                         liTemp.SubItems(1) = g_atUserInfo(i).UserName
                         liTemp.Checked = False
                         '将东东选中
                         For k = 1 To nLen2
                             If g_atUserInfo(i).UserID = aszUnionLocalUser(k) Then
                                 liTemp.Checked = True
                             End If
                         Next k
                     End If
                 Next j
            Next i
        End If
    Else '新增
        For i = 1 To nLen1
            For j = 1 To nLen
                If g_atUserInfo(i).UserID = aszAlllocalUser(j) Then
                    Set liTemp = lvLocalUser.ListItems.Add(, , g_atUserInfo(i).UserID)
                    liTemp.SubItems(1) = g_atUserInfo(i).UserName
                    liTemp.Checked = False
                End If
            Next j
        Next i
    End If
    Set oUnitTem = Nothing
Exit Sub
ErrorHandle:
    ShowErrorMsg

End Sub
Private Function GetAttachLocUser() As Boolean
    '得到关联的本地用户
    Dim oUnitTem As New Unit
    On Error GoTo ErrorHandle
    oUnitTem.Init g_oActUser
    oUnitTem.Identify g_alvItemText(1)
    aszUnionLocalUser = oUnitTem.GetAllAttachUser(frmUnitBeUser.szRemoteUserID)
    Set oUnitTem = Nothing
        
Exit Function
ErrorHandle:
    ShowErrorMsg
    GetAttachLocUser = True
End Function


Private Sub GetInfoFromUI()
    '得到所选中的信息及输入的东东
    
    Dim nCount  As Integer
    Dim nLen As Integer
    Dim i As Integer
    Dim liTemp As ListItem
    
    szRemoteID = txtRemoteUser.Text
    szAnno = txtAnno
    nLen = 0
    ReDim aszUnionLocalUser(1)
    For Each liTemp In lvLocalUser.ListItems
        If liTemp.Checked = True Then
            nLen = nLen + 1
            ReDim Preserve aszUnionLocalUser(1 To nLen)
            aszUnionLocalUser(nLen) = liTemp.Text
        End If
    Next
    
    
End Sub

Private Sub lvLocalUser_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim nLen As Integer, i As Integer
    Dim szTemp As String
    nLen = ArrayLength(g_aszUsedLocUser)
    If nLen > 0 And Item.Checked = True Then
        szTemp = Item.Text
        For i = 1 To nLen
            If szTemp = g_aszUsedLocUser(i) Then
                MsgBox "此用户已关联此外单位的某一远程用户", vbInformation, cszMsg
                Item.Checked = False
            End If
        Next i
    End If
End Sub


Private Sub txtAnno_Validate(Cancel As Boolean)
    If TextLongValidate(255, txtAnno.Text) Then Cancel = True
End Sub

Private Sub txtRemoteUser_Validate(Cancel As Boolean)
    If TextLongValidate(40, txtRemoteUser.Text) Then Cancel = True
    If SpacialStrValid(txtRemoteUser.Text, "[") Then Cancel = True
    If SpacialStrValid(txtRemoteUser.Text, ",") Then Cancel = True
    If SpacialStrValid(txtRemoteUser.Text, "]") Then Cancel = True
End Sub

