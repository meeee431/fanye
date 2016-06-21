VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmCheckerMan 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "检票员管理"
   ClientHeight    =   4245
   ClientLeft      =   2070
   ClientTop       =   2370
   ClientWidth     =   7680
   HelpContextID   =   10000380
   Icon            =   "frmCheckerMan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "检票员设置"
      Height          =   2460
      Left            =   1950
      TabIndex        =   14
      Top             =   1230
      Width           =   5655
      Begin VB.ListBox lstCheckerNotSelect 
         Appearance      =   0  'Flat
         Height          =   1650
         Left            =   90
         MultiSelect     =   2  'Extended
         TabIndex        =   5
         Top             =   570
         Width           =   2325
      End
      Begin VB.ListBox lstCheckerSelect 
         Appearance      =   0  'Flat
         Height          =   1650
         Left            =   3210
         MultiSelect     =   2  'Extended
         TabIndex        =   3
         Top             =   570
         Width           =   2325
      End
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H00E0E0E0&
         Caption         =   "<"
         Height          =   315
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1455
         Width           =   500
      End
      Begin VB.CommandButton cmdRemoveAll 
         BackColor       =   &H00E0E0E0&
         Caption         =   "<<"
         Height          =   315
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1845
         Width           =   500
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00E0E0E0&
         Caption         =   ">"
         Height          =   315
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   660
         Width           =   500
      End
      Begin VB.CommandButton cmdAddAll 
         BackColor       =   &H00E0E0E0&
         Caption         =   ">>"
         Height          =   315
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1050
         Width           =   500
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "已选检票员(&Y):"
         Height          =   180
         Left            =   3210
         TabIndex        =   2
         Top             =   270
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "未选检票员(&N):"
         Height          =   180
         Left            =   105
         TabIndex        =   4
         Top             =   270
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "检票口属性"
      Height          =   825
      Left            =   1950
      TabIndex        =   13
      Top             =   300
      Width           =   5655
      Begin VB.Label lblCheckGateAnn 
         BackStyle       =   0  'Transparent
         Caption         =   "这个检票口属于钢筋混凝土结构，十分坚固。"
         Height          =   180
         Left            =   735
         TabIndex        =   20
         Top             =   525
         Width           =   4680
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "说明:"
         Height          =   180
         Left            =   180
         TabIndex        =   19
         Top             =   525
         Width           =   450
      End
      Begin VB.Label lblCheckGateName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "高速"
         Height          =   180
         Left            =   3240
         TabIndex        =   18
         Top             =   255
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票口名称:"
         Height          =   180
         Left            =   2190
         TabIndex        =   17
         Top             =   240
         Width           =   990
      End
      Begin VB.Label lblCheckGate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "01"
         Height          =   180
         Left            =   1200
         TabIndex        =   16
         Top             =   270
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票口代码:"
         Height          =   180
         Left            =   180
         TabIndex        =   15
         Top             =   270
         Width           =   990
      End
   End
   Begin VB.ListBox lstCheckGate 
      Appearance      =   0  'Flat
      Height          =   3270
      Left            =   135
      TabIndex        =   1
      Top             =   390
      Width           =   1695
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   315
      Left            =   6405
      TabIndex        =   12
      Top             =   3810
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "帮助(&H)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      MICON           =   "frmCheckerMan.frx":038A
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
      Left            =   4980
      TabIndex        =   11
      Top             =   3810
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "取消"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      MICON           =   "frmCheckerMan.frx":03A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   315
      Left            =   3570
      TabIndex        =   10
      Top             =   3810
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "确定"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      MICON           =   "frmCheckerMan.frx":03C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "检票口(E):"
      Height          =   225
      Left            =   150
      TabIndex        =   0
      Top             =   135
      Width           =   1005
   End
End
Attribute VB_Name = "frmCheckerMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim g_oActiveUser As New ActiveUser
Dim oBase As New BaseInfo
Dim oCheckGate As New CheckGate
Dim oChkTk As New CheckTicket

Dim m_aszCheckGates() As String  '检票口信息
Dim m_abChangeSet() As Boolean    '监测检票口是否经过检票员更改
Dim m_anSelectedPos() As Integer    '选中的检票员列表位置映象 二维（检票口，位置）
Dim m_anNotSelectedPos() As Integer '未选中的检票员列位置映象
Dim m_atAllUser() As TOperatorInfo     '所有的检票员信息
Dim m_nCountAllUsers As Integer

Public Sub SetCmdStatus()
    If lstCheckerSelect.ListCount = 0 Then
        cmdRemove.Enabled = False
        cmdRemoveAll.Enabled = False
    Else
        If lstCheckerSelect.SelCount = 0 Then
            cmdRemove.Enabled = False
        Else
            cmdRemove.Enabled = True
        End If
        cmdRemoveAll.Enabled = True
    End If
    If lstCheckerNotSelect.ListCount = 0 Then
        cmdAdd.Enabled = False
        cmdAddAll.Enabled = False
    Else
        If lstCheckerNotSelect.SelCount = 0 Then
            cmdAdd.Enabled = False
        Else
            cmdAdd.Enabled = True
        End If
        cmdAddAll.Enabled = True
    End If
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
DisplayHelp Me
End Sub

Private Sub cmdOk_Click()
    
    MousePointer = vbHourglass
On Error GoTo ErrorPos:

    Dim aszCheckGate() As String
    Dim aszCheckers() As String
    Dim nArrayLength As Integer, nArrayLength2 As Integer, nArrayLength3 As Integer
    Dim nLoop As Integer
    Dim i As Integer, j As Integer
    aszCheckGate = oBase.GetAllCheckGate
    nArrayLength = ArrayLength(aszCheckGate)
    For nLoop = 1 To nArrayLength
        If m_abChangeSet(nLoop) Then
            nArrayLength2 = 0
            For i = 1 To m_nCountAllUsers + 1
                If m_anSelectedPos(nLoop, i) = 0 Then
                    nArrayLength2 = i - 1
                    Exit For
                End If
            Next i
            
            oCheckGate.Identify Trim(aszCheckGate(nLoop, 1))
            aszCheckers = oCheckGate.GetAllChecker
            nArrayLength3 = ArrayLength(aszCheckers)
            For i = 1 To nArrayLength2                        '新增检票员
                For j = 1 To nArrayLength3
                    If m_atAllUser(m_anSelectedPos(nLoop, i)).OperatorId = Trim(aszCheckers(j)) Then
                        Exit For
                    End If
                Next j
                If j > nArrayLength3 Then
                    oCheckGate.AddChecker m_atAllUser(m_anSelectedPos(nLoop, i)).OperatorId
                End If
            Next i
            For i = 1 To nArrayLength3                           '删除检票员
                For j = 1 To nArrayLength2
                    If Trim(aszCheckers(i)) = m_atAllUser(m_anSelectedPos(nLoop, j)).OperatorId Then
                        Exit For
                    End If
                Next j
                If j > nArrayLength2 Then
                    oCheckGate.DeleteChecker Trim(aszCheckers(i))
                End If
            Next i
        End If
    Next nLoop
                    
    
    
    MousePointer = vbNormal
    '检票员设置成功
    MsgBox "检票员设置成功", vbInformation, Me.Caption
'    Unload Me
    Exit Sub
ErrorPos:
    ShowErrorMsg
    MousePointer = vbNormal
'    Unload Me
End Sub

Private Sub cmdRemove_Click()
'    Dim i As Integer, j As Integer
'    For i = 1 To lstCheckerNotSelect.SelCount
'        For j = 0 To lstCheckerNotSelect.ListCount - 1
'            If lstCheckerNotSelect.Selected(j) Then
'                lstcheckerselect.AddItem lstCheckerNotSelect.List(j)
'                lstCheckerNotSelect.RemoveItem j
'                Exit For
'            End If
'        Next j
'    Next i
'    SetCmdStatus
'以上原代码

    Dim i As Integer, j As Integer
    Dim nCurrGateIndex As Integer
    nCurrGateIndex = lstCheckGate.ListIndex + 1
    m_abChangeSet(nCurrGateIndex) = True
    
    Dim szTmpItem As String
    i = 1
    While i <= lstCheckerSelect.ListCount
        If lstCheckerSelect.Selected(i - 1) Then
            Dim nTmpLocationValue As Integer
            nTmpLocationValue = m_anSelectedPos(nCurrGateIndex, i)
            
            szTmpItem = "[" & m_atAllUser(nTmpLocationValue).OperatorId & "]"
            szTmpItem = szTmpItem & m_atAllUser(nTmpLocationValue).OperatorName
            lstCheckerNotSelect.AddItem szTmpItem, nTmpLocationValue - i
            lstCheckerSelect.RemoveItem i - 1
                        
            For j = i To m_nCountAllUsers - 1    '未选中检票员位置映象上移
                If m_anSelectedPos(nCurrGateIndex, j) = 0 Then
                    Exit For
                Else
                    m_anSelectedPos(nCurrGateIndex, j) = m_anSelectedPos(nCurrGateIndex, j + 1)
                End If
            Next j
            If j = m_nCountAllUsers Then
                m_anSelectedPos(nCurrGateIndex, j) = 0
            End If
            
            '选中检票员位置映象后移
            For j = m_nCountAllUsers - 1 To nTmpLocationValue - i + 1 Step -1
                m_anNotSelectedPos(nCurrGateIndex, j + 1) = m_anNotSelectedPos(nCurrGateIndex, j)
            Next j
            m_anNotSelectedPos(nCurrGateIndex, nTmpLocationValue - i + 1) = nTmpLocationValue
        Else
            i = i + 1
        End If
    Wend
    SetCmdStatus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Sub lstCheckerSelect_DblClick()
'    Dim i As Integer, j As Integer
'    For i = 1 To lstCheckerSelect.SelCount
'        For j = 0 To lstCheckerSelect.ListCount - 1
'            If lstCheckerSelect.Selected(j) Then
'                lstCheckerNotSelect.AddItem lstCheckerSelect.List(j)
'                lstCheckerSelect.RemoveItem j
'                Exit For
'            End If
'        Next j
'    Next i
'    SetCmdStatus

    Dim i As Integer, j As Integer
    Dim nCurrGateIndex As Integer
    nCurrGateIndex = lstCheckGate.ListIndex + 1
    m_abChangeSet(nCurrGateIndex) = True
    
    Dim szTmpItem As String
    i = 1
    While i <= lstCheckerSelect.ListCount
        If lstCheckerSelect.Selected(i - 1) Then
            Dim nTmpLocationValue As Integer
            nTmpLocationValue = m_anSelectedPos(nCurrGateIndex, i)
            
            szTmpItem = "[" & m_atAllUser(nTmpLocationValue).OperatorId & "]"
            szTmpItem = szTmpItem & m_atAllUser(nTmpLocationValue).OperatorName
            lstCheckerNotSelect.AddItem szTmpItem, nTmpLocationValue - i
            lstCheckerSelect.RemoveItem i - 1
                        
            For j = i To m_nCountAllUsers - 1    '未选中检票员位置映象上移
                If m_anSelectedPos(nCurrGateIndex, j) = 0 Then
                    Exit For
                Else
                    m_anSelectedPos(nCurrGateIndex, j) = m_anSelectedPos(nCurrGateIndex, j + 1)
                End If
            Next j
            If j = m_nCountAllUsers Then
                m_anSelectedPos(nCurrGateIndex, j) = 0
            End If
            
            '选中检票员位置映象后移
            For j = m_nCountAllUsers - 1 To nTmpLocationValue - i + 1 Step -1
                m_anNotSelectedPos(nCurrGateIndex, j + 1) = m_anNotSelectedPos(nCurrGateIndex, j)
            Next j
            m_anNotSelectedPos(nCurrGateIndex, nTmpLocationValue - i + 1) = nTmpLocationValue
        Else
            i = i + 1
        End If
    Wend
    SetCmdStatus
    
End Sub

Private Sub cmdRemoveAll_Click()
'    Dim i As Integer
'    For i = 0 To lstCheckerSelect.ListCount - 1
'        lstCheckerNotSelect.AddItem lstCheckerSelect.List(i)
'    Next i
'    For i = 1 To lstCheckerSelect.ListCount
'        lstCheckerSelect.RemoveItem 0
'    Next i
'    SetCmdStatus
'以上原代码
    Dim i As Integer, j As Integer
    Dim nCurrGateIndex As Integer
    nCurrGateIndex = lstCheckGate.ListIndex + 1
    m_abChangeSet(nCurrGateIndex) = True
    
    For i = 1 To m_nCountAllUsers
        m_anNotSelectedPos(nCurrGateIndex, i) = i
        m_anSelectedPos(nCurrGateIndex, i) = 0
    Next i
    lstCheckerNotSelect.Clear
    lstCheckerSelect.Clear
    
    Dim szTmpItem As String
    For i = 1 To m_nCountAllUsers        '全部加入
        szTmpItem = "[" & m_atAllUser(i).OperatorId & "]"
        szTmpItem = szTmpItem & m_atAllUser(i).OperatorName
        lstCheckerNotSelect.AddItem szTmpItem
    Next i
    SetCmdStatus

End Sub

Private Sub cmdAdd_Click()
'    Dim i As Integer, j As Integer
'    For i = 1 To lstCheckerNotSelect.SelCount
'        For j = 0 To lstCheckerNotSelect.ListCount - 1
'            If lstCheckerNotSelect.Selected(j) Then
'                lstCheckerSelect.AddItem lstCheckerNotSelect.List(j)
'                lstCheckerNotSelect.RemoveItem j
'                Exit For
'            End If
'        Next j
'    Next i
'    SetCmdStatus
'以上原代码
    Dim i As Integer, j As Integer
    Dim nCurrGateIndex As Integer
    nCurrGateIndex = lstCheckGate.ListIndex + 1
    m_abChangeSet(nCurrGateIndex) = True
    
    Dim szTmpItem As String
    i = 1
    While i <= lstCheckerNotSelect.ListCount
        If lstCheckerNotSelect.Selected(i - 1) Then
            lstCheckerNotSelect.Selected(i - 1) = False
            Dim nTmpLocationValue As Integer
            nTmpLocationValue = m_anNotSelectedPos(nCurrGateIndex, i)
            
            szTmpItem = "[" & m_atAllUser(nTmpLocationValue).OperatorId & "]"
            szTmpItem = szTmpItem & m_atAllUser(nTmpLocationValue).OperatorName
            lstCheckerSelect.AddItem szTmpItem, nTmpLocationValue - i
            lstCheckerNotSelect.RemoveItem i - 1
                        
            For j = i To m_nCountAllUsers - 1    '未选中检票员位置映象上移
                If m_anNotSelectedPos(nCurrGateIndex, j) = 0 Then
                    Exit For
                Else
                    m_anNotSelectedPos(nCurrGateIndex, j) = m_anNotSelectedPos(nCurrGateIndex, j + 1)
                End If
            Next j
            If j = m_nCountAllUsers Then
                m_anNotSelectedPos(nCurrGateIndex, j) = 0
            End If
            
            '选中检票员位置映象后移
            For j = m_nCountAllUsers - 1 To nTmpLocationValue - i + 1 Step -1
                m_anSelectedPos(nCurrGateIndex, j + 1) = m_anSelectedPos(nCurrGateIndex, j)
            Next j
            m_anSelectedPos(nCurrGateIndex, nTmpLocationValue - i + 1) = nTmpLocationValue
        Else
            i = i + 1
        End If
    Wend
    SetCmdStatus

End Sub

Private Sub lstCheckerNotSelect_DblClick()
'    Dim i As Integer, j As Integer
'    For i = 1 To lstCheckerNotSelect.SelCount
'        For j = 0 To lstCheckerNotSelect.ListCount - 1
'            If lstCheckerNotSelect.Selected(j) Then
'                lstCheckerSelect.AddItem lstCheckerNotSelect.List(j)
'                lstCheckerNotSelect.RemoveItem j
'                Exit For
'            End If
'        Next j
'    Next i
'    SetCmdStatus
    
    Dim i As Integer, j As Integer
    Dim nCurrGateIndex As Integer
    nCurrGateIndex = lstCheckGate.ListIndex + 1
    m_abChangeSet(nCurrGateIndex) = True
    
    Dim szTmpItem As String
    i = 1
    While i <= lstCheckerNotSelect.ListCount
        If lstCheckerNotSelect.Selected(i - 1) Then
            lstCheckerNotSelect.Selected(i - 1) = False
            Dim nTmpLocationValue As Integer
            nTmpLocationValue = m_anNotSelectedPos(nCurrGateIndex, i)
            
            szTmpItem = "[" & m_atAllUser(nTmpLocationValue).OperatorId & "]"
            szTmpItem = szTmpItem & m_atAllUser(nTmpLocationValue).OperatorName
            lstCheckerSelect.AddItem szTmpItem, nTmpLocationValue - i
            lstCheckerNotSelect.RemoveItem i - 1
                        
            For j = i To m_nCountAllUsers - 1    '未选中检票员位置映象上移
                If m_anNotSelectedPos(nCurrGateIndex, j) = 0 Then
                    Exit For
                Else
                    m_anNotSelectedPos(nCurrGateIndex, j) = m_anNotSelectedPos(nCurrGateIndex, j + 1)
                End If
            Next j
            If j = m_nCountAllUsers Then
                m_anNotSelectedPos(nCurrGateIndex, j) = 0
            End If
            
            '选中检票员位置映象后移
            For j = m_nCountAllUsers - 1 To nTmpLocationValue - i + 1 Step -1
                m_anSelectedPos(nCurrGateIndex, j + 1) = m_anSelectedPos(nCurrGateIndex, j)
            Next j
            m_anSelectedPos(nCurrGateIndex, nTmpLocationValue - i + 1) = nTmpLocationValue
        Else
            i = i + 1
        End If
    Wend
    SetCmdStatus

End Sub

Private Sub cmdAddAll_Click()
'    Dim i As Integer
'    For i = 0 To lstCheckerNotSelect.ListCount - 1
'        lstCheckerSelect.AddItem lstCheckerNotSelect.List(i)
'    Next i
'    For i = 1 To lstCheckerNotSelect.ListCount
'        lstCheckerNotSelect.RemoveItem 0
'    Next i
'    SetCmdStatus
'以上原代码
    Dim i As Integer, j As Integer
    Dim nCurrGateIndex As Integer
    nCurrGateIndex = lstCheckGate.ListIndex + 1
    m_abChangeSet(nCurrGateIndex) = True
    
    For i = 1 To m_nCountAllUsers
        m_anNotSelectedPos(nCurrGateIndex, i) = 0
        m_anSelectedPos(nCurrGateIndex, i) = i
    Next i
    lstCheckerNotSelect.Clear
    lstCheckerSelect.Clear
    
    Dim szTmpItem As String
    For i = 1 To m_nCountAllUsers        '全部加入
        szTmpItem = "[" & m_atAllUser(i).OperatorId & "]"
        szTmpItem = szTmpItem & m_atAllUser(i).OperatorName
        lstCheckerSelect.AddItem szTmpItem
    Next i
    SetCmdStatus
End Sub

'Private Sub Form_Load()
''    g_oActiveUser.Login "0000", "password", "saa"
''    MousePointer = MousePointerConstants.vbHourglass
'    oBase.Init g_oActiveUser
'    oCheckGate.Init g_oActiveUser
'    oChkTk.Init g_oActiveUser
'
'    Dim szCheckGateID()  As String
'    Dim nArrayLength As Integer
'    Dim nArrayLength2 As Integer
'
'    szCheckGateID = oBase.GetAllCheckGate
'    nArrayLength = ArrayLength(szCheckGateID, 1) '检票口数量
'
'    Dim nLoop As Integer
'    For nLoop = 1 To nArrayLength
'        lstCheckGate.AddItem szCheckGateID(nLoop, 2) '初始化检票口
'    Next nLoop
'
'    Label5.Caption = szCheckGateID(1, 1)
'    Label7.Caption = szCheckGateID(1, 2)
'    Label9.Caption = szCheckGateID(1, 3)
'
'    Dim szCheckerID() As String
'    Dim aszCheckerIDSelect() As String
'    Dim nArrayLengthSelect As Integer
'    oCheckGate.Identify szCheckGateID(1, 1)
'    aszCheckerIDSelect = oCheckGate.GetAllChecker  '取某一个检票口的检票员
'    nArrayLengthSelect = ArrayLength(aszCheckerIDSelect)
'    For nLoop = 1 To nArrayLengthSelect '已选检票员
'        lstCheckerSelect.AddItem aszCheckerIDSelect(nLoop)
'    Next nLoop
'
'    Dim nLoop2 As Integer
'    Dim nLoop3 As Integer
'    Dim bExist As Boolean
'    For nLoop = 2 To nArrayLength '未选检票员
'        oCheckGate.Identify szCheckGateID(nLoop, 1)
'        szCheckerID = oCheckGate.GetAllChecker '取某一个检票口的检票员
'        nArrayLength2 = ArrayLength(szCheckerID)
'        For nLoop2 = 1 To nArrayLength2
'            bExist = False
'            For nLoop3 = 1 To nArrayLengthSelect
'                If Trim(aszCheckerIDSelect(nLoop3)) = Trim(szCheckerID(nLoop2)) Then
'                    bExist = True
'                    Exit For
'                End If
'            Next nLoop3
'            For nLoop3 = 1 To lstCheckerNotSelect.ListCount '已有该检票员
'                If Trim(lstCheckerNotSelect.List(nLoop3 - 1)) = Trim(szCheckerID(nLoop2)) Then
'                    bExist = True
'                    Exit For
'                End If
'            Next nLoop3
'            If bExist = False Then
'                lstCheckerNotSelect.AddItem szCheckerID(nLoop2)
'            End If
'        Next nLoop2
'    Next nLoop
'
'    szCheckerID = oChkTk.GetNewChecker '新增检票员
'    nArrayLength2 = ArrayLength(szCheckerID)
'    For nLoop = 1 To nArrayLength2
'        lstCheckerNotSelect.AddItem szCheckerID(nLoop)
'    Next nLoop
'
'    SetCmdStatus
''    MousePointer = MousePointerConstants.vbDefault
''    lstCheckGate.Selected(0) = True
''    lstCheckerSelect.Selected(0) = True
''    lstCheckerNotSelect.Selected(0) = True
'End Sub
'以上原代码
Private Sub Form_Load()
    AlignFormPos Me
    
    ShowSBInfo "正在读取检票口列表..."
    
    oBase.Init g_oActiveUser
    oCheckGate.Init g_oActiveUser
    oChkTk.Init g_oActiveUser

    Dim nArrayLength As Integer
    Dim nArrayLength2 As Integer

    
    m_aszCheckGates = oBase.GetAllCheckGate
    nArrayLength = ArrayLength(m_aszCheckGates, 1) '检票口数量
    If nArrayLength = 0 Then
        Exit Sub
    End If
    
    Dim nLoop As Integer
    For nLoop = 1 To nArrayLength
        lstCheckGate.AddItem m_aszCheckGates(nLoop, 2) '初始化检票口
    Next nLoop

    If nArrayLength = 0 Then
        Exit Sub
    End If
    
    lblCheckGate.Caption = m_aszCheckGates(1, 1)
    lblCheckGateName.Caption = m_aszCheckGates(1, 2)
    lblCheckGateAnn.Caption = m_aszCheckGates(1, 3)

    ShowSBInfo "取得检票口操作员列表..."
    
    Dim aszUsers() As TUserInfo
    Dim oUser As New SystemMan
    oUser.Init g_oActiveUser
    aszUsers = oUser.GetAllUser
    nArrayLength2 = ArrayLength(aszUsers)
    m_nCountAllUsers = nArrayLength2
    If nArrayLength2 = 0 Then
        Exit Sub
    End If
    ReDim m_atAllUser(1 To m_nCountAllUsers)
    
    For nLoop = 1 To nArrayLength2
        m_atAllUser(nLoop).OperatorId = aszUsers(nLoop).UserID
        m_atAllUser(nLoop).OperatorName = aszUsers(nLoop).UserName
    Next
    
    ReDim m_anSelectedPos(1 To nArrayLength, 1 To nArrayLength2 + 1)
    ReDim m_anNotSelectedPos(1 To nArrayLength, 1 To nArrayLength2 + 1)
    ReDim m_abChangeSet(1 To nArrayLength)
    
    For nLoop = 1 To nArrayLength
        m_anSelectedPos(nLoop, 1) = -1   '未操作标志
        m_anNotSelectedPos(nLoop, 1) = -1
        m_abChangeSet(nLoop) = False
    Next nLoop
    
    lstCheckGate.ListIndex = 0
    
    ShowSBInfo ""
'    lstCheckGate_Click
'    SetCmdStatus
End Sub

Private Sub lstCheckerSelect_Click()
    SetCmdStatus
End Sub

Private Sub lstCheckerNotSelect_Click()
    SetCmdStatus
End Sub

'Private Sub lstCheckGate_Click()
'    MousePointer = MousePointerConstants.vbHourglass
'    Dim szCheckGateID As String
'    szCheckGateID = Trim(lstCheckGate.Text)
'    Dim aszCheckGate() As String
'    Dim nArrayLength As Integer
'    Dim nArrayLength2 As Integer
'    Dim nLoop As Integer
'
''    ClearBoxItem lstCheckerSelect
''    ClearBoxItem lstCheckerNotSelect
'    lstCheckerSelect.Clear
'    lstCheckerNotSelect.Clear
'
'    aszCheckGate = oBase.GetAllCheckGate
'    nArrayLength = ArrayLength(aszCheckGate)
'    For nLoop = 1 To nArrayLength
'        If Trim(szCheckGateID) = Trim(aszCheckGate(nLoop, 2)) Then
'            szCheckGateID = Trim(aszCheckGate(nLoop, 1))
'            Label5.Caption = aszCheckGate(nLoop, 1)
'            Label7.Caption = aszCheckGate(nLoop, 2)
'            Label9.Caption = aszCheckGate(nLoop, 3)
'        End If
'    Next nLoop
'
'    Dim szCheckerID() As String
'    Dim aszCheckerIDSelect() As String
'    Dim nArrayLengthSelect As Integer
'    oCheckGate.Identify szCheckGateID
'    aszCheckerIDSelect = oCheckGate.GetAllChecker  '取某一个检票口的检票员
'    nArrayLengthSelect = ArrayLength(aszCheckerIDSelect)
'    For nLoop = 1 To nArrayLengthSelect '已选检票员
'        lstCheckerSelect.AddItem aszCheckerIDSelect(nLoop)
'    Next nLoop
'
'    Dim nLoop2 As Integer
'    Dim nLoop3 As Integer
'    Dim bExist As Boolean
'    For nLoop = 1 To nArrayLength '未选检票员
'        If Trim(szCheckGateID) <> Trim(aszCheckGate(nLoop, 1)) Then
'            oCheckGate.Identify aszCheckGate(nLoop, 1)
'            szCheckerID = oCheckGate.GetAllChecker '取某一个检票口的检票员
'            nArrayLength2 = ArrayLength(szCheckerID)
'            For nLoop2 = 1 To nArrayLength2
'                bExist = False
'                For nLoop3 = 1 To nArrayLengthSelect '该检票口是否已有该检票员
'                    If Trim(aszCheckerIDSelect(nLoop3)) = Trim(szCheckerID(nLoop2)) Then
'                        bExist = True
'                        Exit For
'                    End If
'                Next nLoop3
'                For nLoop3 = 1 To lstCheckerNotSelect.ListCount '已有该检票员
'                    If Trim(lstCheckerNotSelect.List(nLoop3 - 1)) = Trim(szCheckerID(nLoop2)) Then
'                        bExist = True
'                        Exit For
'                    End If
'                Next nLoop3
'                If bExist = False Then
'                    lstCheckerNotSelect.AddItem szCheckerID(nLoop2)
'                End If
'            Next nLoop2
'        End If
'    Next nLoop
'
'    szCheckerID = oChkTk.GetNewChecker '新增检票员
'    nArrayLength2 = ArrayLength(szCheckerID)
'    For nLoop = 1 To nArrayLength2
'        lstCheckerNotSelect.AddItem szCheckerID(nLoop)
'    Next nLoop
'    MousePointer = MousePointerConstants.vbDefault
'    SetCmdStatus
'End Sub

Private Sub lstCheckGate_Click()
    MousePointer = MousePointerConstants.vbHourglass
    Dim szCheckGateID As String
    szCheckGateID = Trim(m_aszCheckGates(lstCheckGate.ListIndex + 1, 1))
    Dim nCurrSelectIndex As Integer
    nCurrSelectIndex = lstCheckGate.ListIndex + 1
    
    Dim nCountAllCheckers As Integer
    Dim nLoop As Integer, nLoop2 As Integer

'    ClearBoxItem lstCheckerSelect
'    ClearBoxItem lstCheckerNotSelect

    ShowSBInfo "取得操作员列表..."

    lblCheckGate.Caption = szCheckGateID
    lblCheckGateName.Caption = m_aszCheckGates(lstCheckGate.ListIndex + 1, 2)
    lblCheckGateAnn.Caption = m_aszCheckGates(lstCheckGate.ListIndex + 1, 3)
        
    If m_anSelectedPos(nCurrSelectIndex, 1) = -1 Then        '未经过操作
        m_anSelectedPos(nCurrSelectIndex, 1) = 0
        Dim oChkGate As New CheckGate
        Dim aszCheckers() As String
        Dim szTmpChecker As String
        oChkGate.Init g_oActiveUser
        oChkGate.Identify szCheckGateID
        aszCheckers = oChkGate.GetAllChecker
        nCountAllCheckers = ArrayLength(aszCheckers)
        
        Dim nCountCheckers As Integer
        nCountCheckers = 0
        
        '以下代码将已选检票员和未选检票员在matAllUser中的位置分别放入manSelectedPos和manNotSelectedPos
        For nLoop = 1 To nCountAllCheckers
            szTmpChecker = Trim(aszCheckers(nLoop))
            For nLoop2 = 1 To m_nCountAllUsers
                If szTmpChecker = m_atAllUser(nLoop2).OperatorId Then
                    nCountCheckers = nCountCheckers + 1
                    m_anSelectedPos(nCurrSelectIndex, nCountCheckers) = nLoop2
                    Exit For
                End If
            Next nLoop2
        Next nLoop
        For nLoop = nCountCheckers + 1 To m_nCountAllUsers    '余下置0
            m_anSelectedPos(nCurrSelectIndex, nLoop) = 0
        Next nLoop
        nCountCheckers = 0
        nLoop2 = 1
        For nLoop = 1 To m_nCountAllUsers
            If nLoop = m_anSelectedPos(nCurrSelectIndex, nLoop2) Then
                nLoop2 = nLoop2 + 1
            Else
                nCountCheckers = nCountCheckers + 1
                m_anNotSelectedPos(nCurrSelectIndex, nCountCheckers) = nLoop
            End If
        Next nLoop
        For nLoop = nCountCheckers + 1 To m_nCountAllUsers    '余下置0
            m_anNotSelectedPos(nCurrSelectIndex, nLoop) = 0
        Next nLoop
    Else
        For nLoop = 1 To m_nCountAllUsers
            If m_anSelectedPos(nCurrSelectIndex, nLoop) = 0 Then
                nCountAllCheckers = nLoop - 1
                Exit For
            End If
        Next nLoop
        If nLoop > m_nCountAllUsers Then
            nCountAllCheckers = m_nCountAllUsers
        End If
    End If
            
    '以下代码将在本检票口的已选检票员和未选检票员分别放入列表框
    Dim szTmpItem As String
    lstCheckerSelect.Clear
    lstCheckerNotSelect.Clear
    For nLoop = 1 To nCountAllCheckers
        If m_anSelectedPos(nCurrSelectIndex, nLoop) > 0 Then
            szTmpItem = "[" & m_atAllUser(m_anSelectedPos(nCurrSelectIndex, nLoop)).OperatorId & "]"
            szTmpItem = szTmpItem & m_atAllUser(m_anSelectedPos(nCurrSelectIndex, nLoop)).OperatorName
            lstCheckerSelect.AddItem szTmpItem
        End If
    Next nLoop
    For nLoop = 1 To m_nCountAllUsers - nCountAllCheckers
        If m_anNotSelectedPos(nCurrSelectIndex, nLoop) > 0 Then
            szTmpItem = "[" & m_atAllUser(m_anNotSelectedPos(nCurrSelectIndex, nLoop)).OperatorId & "]"
            szTmpItem = szTmpItem & m_atAllUser(m_anNotSelectedPos(nCurrSelectIndex, nLoop)).OperatorName
            lstCheckerNotSelect.AddItem szTmpItem
        Else
            Exit For
        End If
    Next nLoop
    
    MousePointer = MousePointerConstants.vbDefault
    SetCmdStatus
    
    ShowSBInfo ""
    
End Sub



