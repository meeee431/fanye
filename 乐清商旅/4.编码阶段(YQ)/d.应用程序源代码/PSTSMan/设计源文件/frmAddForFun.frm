VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmAddForFun 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "添加"
   ClientHeight    =   3375
   ClientLeft      =   1680
   ClientTop       =   2880
   ClientWidth     =   4650
   HelpContextID   =   5003001
   Icon            =   "frmAddForFun.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab stAdd 
      Height          =   2865
      Left            =   90
      TabIndex        =   2
      Top             =   45
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   5054
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "用户组(&G)"
      TabPicture(0)   =   "frmAddForFun.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "adGroup"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "用户(&U)"
      TabPicture(1)   =   "frmAddForFun.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "adUser"
      Tab(1).ControlCount=   1
      Begin PSTSMan.AddDel adUser 
         Height          =   2505
         Left            =   -74940
         TabIndex        =   4
         Top             =   315
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   4419
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
      Begin PSTSMan.AddDel adGroup 
         Height          =   2490
         Left            =   45
         TabIndex        =   3
         Top             =   345
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   4392
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
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   315
      Left            =   3450
      TabIndex        =   1
      Top             =   2985
      Width           =   1015
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   315
      Left            =   2250
      TabIndex        =   0
      Top             =   2985
      Width           =   1015
   End
End
Attribute VB_Name = "frmAddForFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'选择要赋指定功能的用户或用户组

Option Explicit
Option Base 1


Private Sub cmdCancel_Click()
    Unload Me
    ReDim g_aszUserAdd(1 To 2, 1 To 1)
    ReDim g_aszUserGroupAdd(1 To 2, 1 To 1)
End Sub

Private Sub cmdOK_Click()
    '得到要赋权的用户及用户组
    Dim aszUtemp() As String
    Dim aszGtemp() As String
    Dim nlenU As Integer
    Dim nLenG As Integer
    Dim i As Integer
    
    i = MsgBox("确认给选择的用户和用户组添加指定的权限或权限组?", vbQuestion + vbYesNo, cszMsg)
    If i = vbNo Then Exit Sub
    
    aszUtemp = adUser.RightData
    aszGtemp = adGroup.RightData
    nlenU = ArrayLength(aszUtemp)
    nLenG = ArrayLength(aszGtemp)
    
    
    If nlenU > 0 Then
        ReDim g_aszUserAdd(1 To 2, 1 To nlenU)
        For i = 1 To nlenU
            g_aszUserAdd(1, i) = PartCode(aszUtemp(i))
            g_aszUserAdd(2, i) = PartCode(aszUtemp(i), False)
        Next i
    Else
        ReDim g_aszUserAdd(1 To 2, 1 To 1)
    End If
    
    If nLenG > 0 Then
        ReDim g_aszUserGroupAdd(1 To 2, 1 To nLenG)
        For i = 1 To nLenG
            g_aszUserGroupAdd(1, i) = PartCode(aszGtemp(i))
            g_aszUserGroupAdd(2, i) = PartCode(aszGtemp(i), False)
        Next i
    Else
        ReDim g_aszUserGroupAdd(1 To 2, 1 To 1)
    End If
    
    Unload Me
    
    
End Sub

Private Sub Form_Load()
    
    LoadUserInfo
    LoadGroupInfo
'    stAdd_Click (0)
End Sub




Private Sub LoadUserInfo()
    '显示用户信息
    Dim i As Integer, nLen1 As Integer
    Dim j As Integer, nLen2 As Integer
    Dim bExist As Boolean

    
    nLen1 = ArrayLength(g_atUserInfo)
    nLen2 = ArrayLength(g_aszUser)
'
    
    If nLen2 > 0 Then
        If g_aszUser(1) <> "" Then
            For i = 1 To nLen1
                For j = 1 To nLen2
                    If g_atUserInfo(i).UserID = g_aszUser(j) Then
                        bExist = True
                        Exit For
                    End If
                Next j
                If bExist = True Then
                    bExist = False
                Else
                    adUser.AddData MakeDisplayString(g_atUserInfo(i).UserID, g_atUserInfo(i).UserName)
                End If
            Next i
        Else
            For i = 1 To nLen1
                adUser.AddData MakeDisplayString(g_atUserInfo(i).UserID, g_atUserInfo(i).UserName)
            Next i
        End If
    Else
        For i = 1 To nLen1
            adUser.AddData MakeDisplayString(g_atUserInfo(i).UserID, g_atUserInfo(i).UserName)
        Next i
    End If
    
    
End Sub

Private Sub LoadGroupInfo()
    '显示功能组信息
    Dim i As Integer, nLen1 As Integer
    Dim j As Integer, nLen2 As Integer
    Dim bExist As Boolean

    
    nLen1 = ArrayLength(g_atUserGroupInfo)
    nLen2 = ArrayLength(g_aszUserGroup)
    
    
    If nLen2 > 0 Then
        If g_aszUserGroup(1) <> "" Then
            For i = 1 To nLen1
                For j = 1 To nLen2
                    If g_atUserGroupInfo(i).UserGroupID = g_aszUserGroup(j) Then
                        bExist = True
                        Exit For
                    End If
                Next j
                If bExist = True Then
                    bExist = False
                Else
                    adGroup.AddData MakeDisplayString(g_atUserGroupInfo(i).UserGroupID, g_atUserGroupInfo(i).GroupName)
                End If
            Next i
        Else
            For i = 1 To nLen1
                adGroup.AddData MakeDisplayString(g_atUserGroupInfo(i).UserGroupID, g_atUserGroupInfo(i).GroupName)
            Next i
        End If
    Else
        For i = 1 To nLen1
            adGroup.AddData MakeDisplayString(g_atUserGroupInfo(i).UserGroupID, g_atUserGroupInfo(i).GroupName)
        Next i
    
    End If

End Sub
