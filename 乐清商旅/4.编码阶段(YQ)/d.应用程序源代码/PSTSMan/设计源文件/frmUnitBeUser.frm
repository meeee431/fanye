VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUnitBeUser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "远程单位用户管理"
   ClientHeight    =   3870
   ClientLeft      =   2880
   ClientTop       =   1905
   ClientWidth     =   5775
   HelpContextID   =   50000330
   Icon            =   "frmUnitBeUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1845
      Left            =   135
      TabIndex        =   4
      Top             =   1920
      Width           =   5595
      Begin VB.CommandButton cmdPass 
         Caption         =   "修改密码..."
         Enabled         =   0   'False
         Height          =   300
         Left            =   4365
         TabIndex        =   8
         Top             =   1320
         Width           =   1200
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "编辑(&E)..."
         Enabled         =   0   'False
         Height          =   300
         Left            =   4365
         TabIndex        =   7
         Top             =   610
         Width           =   1200
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "删除(&D)"
         Enabled         =   0   'False
         Height          =   300
         Left            =   4365
         TabIndex        =   6
         Top             =   255
         Width           =   1200
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "添加(&A)..."
         Enabled         =   0   'False
         Height          =   300
         Left            =   4365
         TabIndex        =   5
         Top             =   965
         Width           =   1200
      End
      Begin MSComctlLib.ListView lvServer 
         Height          =   1590
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   2805
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
         BorderStyle     =   1
         Appearance      =   0
         Enabled         =   0   'False
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "远程用户ID(账号)"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "注释"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Label lblServer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "本单位登录互连单位的用户(&G):"
         Height          =   180
         Left            =   15
         TabIndex        =   10
         Top             =   0
         Width           =   2520
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4965
      Top             =   1110
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   300
      Left            =   4560
      TabIndex        =   3
      Top             =   750
      Width           =   1095
   End
   Begin VB.CommandButton cmdclose 
      Cancel          =   -1  'True
      Caption         =   "关闭(&C)"
      Height          =   300
      Left            =   4560
      TabIndex        =   2
      Top             =   390
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvAgent 
      Height          =   1455
      Left            =   105
      TabIndex        =   1
      Top             =   375
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   0
      Enabled         =   0   'False
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "用户ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "用户全名"
         Object.Width           =   5115
      EndProperty
   End
   Begin VB.Label lblAgent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "互连单位登录本单位的用户(&W):"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   2520
   End
End
Attribute VB_Name = "frmUnitBeUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1


Dim aszAllUser() As String
Public bEditRemote As Boolean
Public szRemoteUserID As String '选中的远程用户代码
Public szAnno As String '选中的远程用户注释


Private Sub cmdAdd_Click()
    
    frmUnitBeUser.bEditRemote = False
    GetUsedLocalUser (bEditRemote)
    frmAddRemoteUser.Show vbModal, Me
    szRemoteUserID = ""
End Sub

Private Sub cmdClose_Click()
    Unload Me
    ReDim g_atRemoteUser(1)
    szRemoteUserID = ""
End Sub

Private Sub cmdDelete_Click()
    Dim oUnitTemp As New Unit
    Dim nTemp As Integer
    
    If szRemoteUserID <> "" Then
        nTemp = MsgBox("确认删除此远程账号:[" & szRemoteUserID & "]?", vbExclamation + vbYesNo, cszMsg)
        If nTemp = vbYes Then
            On Error GoTo ErrorHandle
            oUnitTemp.Init g_oActUser
            oUnitTemp.Identify g_alvItemText(1)
            oUnitTemp.DeleteRemoteUser szRemoteUserID
            GetAndDisPlayRemote
        Else
            Exit Sub
        End If
    Else
        MsgBox "请选择某一远程账号.", vbInformation, cszMsg
        Exit Sub
    End If
    szRemoteUserID = ""
Exit Sub
ErrorHandle:
    ShowErrorMsg
    szRemoteUserID = ""
End Sub

Private Sub cmdEdit_Click()
    If szRemoteUserID <> "" Then
        frmUnitBeUser.bEditRemote = True
        GetUsedLocalUser (bEditRemote)
        frmAddRemoteUser.Show vbModal, Me
        szRemoteUserID = ""
    Else
        MsgBox "请选择某一远程账号.", vbInformation, cszMsg
        Exit Sub
    End If
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me, content
End Sub

Private Sub cmdPass_Click()
    If szRemoteUserID <> "" Then
        frmChangeRemotePassWord.Show vbModal, Me
        szRemoteUserID = ""
    Else
        MsgBox "请选择某一远程账号.", vbInformation, cszMsg
        Exit Sub
    End If
    szRemoteUserID = ""
End Sub


Private Sub Form_Load()
    Dim oUnitTemp As New Unit
    Dim nUnitType As EUnitType
    Dim i As Integer
    Dim nLen As Integer
    Dim nLen1 As Integer
    Dim j As Integer
    Dim liTemp As ListItem
    
    Me.Top = (Screen.Height - Me.ScaleHeight) / 2
    Me.Left = (Screen.Width - Me.ScaleWidth) / 2

    
    bEditRemote = False
    
    If frmAEUnit.bEdit = True Then
        On Error GoTo ErrorHandle
        oUnitTemp.Init g_oActUser
        oUnitTemp.Identify g_alvItemText(1)
        nUnitType = oUnitTemp.UnitType
        
        
        '登录本单位的用户
        If nUnitType = TP_UnitSC Or nUnitType = TP_UnitClient Then
            lvAgent.Enabled = True
            lvAgent.BackColor = &H80000005
            On Error GoTo 0
            On Error Resume Next
            aszAllUser = oUnitTemp.GetAllUser
            
            
            nLen = 0
            
            nLen = ArrayLength(aszAllUser)
            
            
            
            nLen1 = ArrayLength(g_atUserInfo)
            On Error GoTo 0
            On Error GoTo ErrorHandle
            If nLen <> 0 Then
                For i = 1 To nLen1
                    For j = 1 To nLen
                        If g_atUserInfo(i).UserID = aszAllUser(j) Then
                            Set liTemp = lvAgent.ListItems.Add(, , g_atUserInfo(i).UserID)
                            liTemp.SubItems(1) = g_atUserInfo(i).UserName
                        End If
                    Next j
                Next i
            End If
        End If
        
        '远程用户
        If nUnitType = TP_UnitSC Or nUnitType = TP_UnitServer Then

            lvServer.Enabled = True
            lvServer.BackColor = &H80000005
            If lvServer.ListItems.Count > 0 Then lvServer.ListItems(1).Selected = True
            
            
            GetAndDisPlayRemote
            
            Frame1.Enabled = True
            cmdDelete.Enabled = True
            cmdEdit.Enabled = True
            cmdAdd.Enabled = True
            cmdPass.Enabled = True

        End If
    End If
        
    
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub



Private Sub lvServer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    szRemoteUserID = Item.Text
    szAnno = Item.SubItems(1)
End Sub

Public Sub GetAndDisPlayRemote()
    Dim oUnitTemp As New Unit
    Dim nLen As Integer, i As Integer
    Dim liTemp As ListItem
    
    On Error GoTo ErrorHandle
    oUnitTemp.Init g_oActUser
    oUnitTemp.Identify g_alvItemText(1)
    g_atRemoteUser = oUnitTemp.GetAllRemouteUser  '得到所有的远程用户信息
    On Error GoTo 0
    On Error Resume Next
    nLen = 0
    nLen = ArrayLength(g_atRemoteUser)
    lvServer.ListItems.Clear
    If nLen > 0 Then
        For i = 1 To nLen
            Set liTemp = lvServer.ListItems.Add(, , g_atRemoteUser(i).szRemoteUserID)
            liTemp.SubItems(1) = g_atRemoteUser(i).szAnnotation
        Next i
    End If
    On Error GoTo 0
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub GetUsedLocalUser(bTemp As Boolean)
    Dim i As Integer, nLen As Integer, nListCount As Integer, nLen1 As Integer, j As Integer
    Dim aszTemp() As String
    Dim oUnit As New Unit
    nLen1 = 0
    nListCount = lvServer.ListItems.Count
    
    On Error GoTo ErrorHandle
    oUnit.Init g_oActUser
    oUnit.Identify g_alvItemText(1)
    
    
    If nListCount > 0 Then
        For i = 1 To nListCount
            If bTemp = True Then
                If lvServer.ListItems(i).Text <> szRemoteUserID Then
                    aszTemp = oUnit.GetAllAttachUser(lvServer.ListItems(i).Text)
                    nLen = ArrayLength(aszTemp)
                    If nLen > 0 Then
                        nLen1 = nLen1 + nLen
                        ReDim Preserve g_aszUsedLocUser(1 To nLen1)
                        For j = 1 To nLen
                            g_aszUsedLocUser(nLen1 - nLen + j) = aszTemp(j)
                        Next j
                    End If
                End If
            Else
                ReDim aszTemp(1)
                aszTemp = oUnit.GetAllAttachUser(lvServer.ListItems(i).Text)
                nLen = ArrayLength(aszTemp)
                If nLen > 0 And aszTemp(1) <> "" Then
                    nLen1 = nLen1 + nLen
                    ReDim Preserve g_aszUsedLocUser(1 To nLen1)
                    For j = 1 To nLen
                        g_aszUsedLocUser(nLen1 - nLen + j) = aszTemp(j)
                    Next j
                End If
            End If
        Next i
    End If
    
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub


Private Sub Timer1_Timer()

    If szRemoteUserID = "" Then
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
        cmdPass.Enabled = False
    Else
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True
        cmdPass.Enabled = True
    End If
    
    
End Sub
