VERSION 5.00
Object = "{BBF95DAA-F9CB-4CA9-A673-E0E9E0193752}#1.0#0"; "STSellCtl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmLugFormula 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "行包计算公式"
   ClientHeight    =   6825
   ClientLeft      =   3210
   ClientTop       =   2295
   ClientWidth     =   6810
   Icon            =   "frmLugFormula.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   6810
   Begin VB.TextBox txtComment 
      Height          =   4020
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   33
      Top             =   1965
      Width           =   4350
   End
   Begin VB.TextBox txtID 
      Height          =   300
      Left            =   1395
      TabIndex        =   1
      Top             =   900
      Width           =   885
   End
   Begin VB.PictureBox ptCaption 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   7965
      TabIndex        =   31
      Top             =   0
      Width           =   7965
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请设定行包计算公式:"
         Height          =   180
         Left            =   270
         TabIndex        =   32
         Top             =   240
         Width           =   1710
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   0
      TabIndex        =   30
      Top             =   690
      Width           =   8115
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   345
      Left            =   300
      TabIndex        =   27
      Top             =   6375
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   609
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
      MICON           =   "frmLugFormula.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtName 
      Height          =   465
      Left            =   1395
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1350
      Width           =   5235
   End
   Begin STSellCtl.ucNumTextBox txtParam 
      Height          =   315
      Index           =   1
      Left            =   5220
      TabIndex        =   7
      Top             =   1980
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin STSellCtl.ucNumTextBox txtParam 
      Height          =   315
      Index           =   5
      Left            =   5220
      TabIndex        =   14
      Top             =   3570
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin STSellCtl.ucNumTextBox txtParam 
      Height          =   315
      Index           =   4
      Left            =   5220
      TabIndex        =   28
      Top             =   3180
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin STSellCtl.ucNumTextBox txtParam 
      Height          =   315
      Index           =   3
      Left            =   5220
      TabIndex        =   11
      Top             =   2775
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin STSellCtl.ucNumTextBox txtParam 
      Height          =   315
      Index           =   2
      Left            =   5220
      TabIndex        =   9
      Top             =   2385
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin STSellCtl.ucNumTextBox txtParam 
      Height          =   315
      Index           =   6
      Left            =   5220
      TabIndex        =   16
      Top             =   3975
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin STSellCtl.ucNumTextBox txtParam 
      Height          =   315
      Index           =   10
      Left            =   5220
      TabIndex        =   24
      Top             =   5565
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin STSellCtl.ucNumTextBox txtParam 
      Height          =   315
      Index           =   9
      Left            =   5220
      TabIndex        =   22
      Top             =   5160
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin STSellCtl.ucNumTextBox txtParam 
      Height          =   315
      Index           =   8
      Left            =   5220
      TabIndex        =   20
      Top             =   4770
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin STSellCtl.ucNumTextBox txtParam 
      Height          =   315
      Index           =   7
      Left            =   5220
      TabIndex        =   18
      Top             =   4365
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   5415
      TabIndex        =   26
      Top             =   6375
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "关闭(&C)"
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
      MICON           =   "frmLugFormula.frx":0028
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
      Height          =   345
      Left            =   4005
      TabIndex        =   25
      ToolTipText     =   "保存协议"
      Top             =   6375
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "保存(&S)"
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
      MICON           =   "frmLugFormula.frx":0044
      PICN            =   "frmLugFormula.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   930
      Left            =   -120
      TabIndex        =   29
      Top             =   6090
      Width           =   8745
   End
   Begin VB.ComboBox cboFormula 
      Height          =   300
      Left            =   3540
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   900
      Width           =   3090
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "计算公式(&F):"
      Height          =   180
      Left            =   2385
      TabIndex        =   4
      Top             =   945
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公式名称(&N):"
      Height          =   180
      Left            =   150
      TabIndex        =   2
      Top             =   1395
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公式代码(&I):"
      Height          =   180
      Left            =   150
      TabIndex        =   0
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数1&0"
      Height          =   180
      Index           =   10
      Left            =   4650
      TabIndex        =   23
      Top             =   5625
      Width           =   540
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数&9"
      DataField       =   "9"
      Height          =   180
      Index           =   9
      Left            =   4650
      TabIndex        =   21
      Top             =   5235
      Width           =   450
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数&8"
      DataField       =   "8"
      Height          =   180
      Index           =   8
      Left            =   4650
      TabIndex        =   19
      Top             =   4815
      Width           =   450
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数&7"
      DataField       =   "7"
      Height          =   180
      Index           =   7
      Left            =   4650
      TabIndex        =   17
      Top             =   4425
      Width           =   450
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数&6"
      Height          =   180
      Index           =   6
      Left            =   4650
      TabIndex        =   15
      Top             =   4035
      Width           =   450
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数&1"
      Height          =   180
      Index           =   1
      Left            =   4650
      TabIndex        =   6
      Top             =   2040
      Width           =   450
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数&2"
      Height          =   180
      Index           =   2
      Left            =   4650
      TabIndex        =   8
      Top             =   2445
      Width           =   450
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数&3"
      Height          =   180
      Index           =   3
      Left            =   4650
      TabIndex        =   10
      Top             =   2850
      Width           =   450
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数&4"
      Height          =   180
      Index           =   4
      Left            =   4650
      TabIndex        =   12
      Top             =   3240
      Width           =   450
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数&5"
      Height          =   180
      Index           =   5
      Left            =   4650
      TabIndex        =   13
      Top             =   3645
      Width           =   450
   End
End
Attribute VB_Name = "frmLugFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'* Source File Name:frmLugFormula.frm
'* Project Name:PSTLugMan.vbp
'* Engineer:
'* Date Generated:2004/08/2
'* Last Revision Date:2004/08/02
'* Brief Description:修改行包票价公式
'* Relational Document:
'**********************************************************

Option Explicit

Public m_bIsParent As Boolean '是否父窗体调用
Public m_eStatus As eFormStatus
Public m_szFormulaId As String


Private m_aItemFormulaInfo() As TPriceFormulaInfo '所有的公式信息(包括公式的说明,公式的中文名,参数个数,参数的说明等)
Private m_nItemFormulaInfoCount As Integer

Private m_oPriceItemFunLib As New LugFunLib

Private Sub cboFormula_Change()
    ShowFormulaInfo
    txtParamClear
End Sub

Private Sub cboFormula_Click()
    ShowFormulaInfo
    txtParamClear
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdOk_Click()
    '保存公式
    Dim atFormulaInfo() As TLuggageFormulaInfo
    Dim aszInfo(0 To 3) As String
    Dim i As Integer
    
    
    ReDim atFormulaInfo(1 To 1) As TLuggageFormulaInfo
    On Error GoTo ErrorHandle
    With atFormulaInfo(1)
        .FormulaID = txtID.Text
        .FormulaName = txtName.Text
        .ItemFormula = m_aItemFormulaInfo(cboFormula.ListIndex + 1).szFunName
        For i = 1 To 10
            .Param(i) = txtParam(i).Text
        Next i
        
    End With
    If m_eStatus = AddStatus Then
        '新增
        m_oluggageSvr.AddLugFormulaInfo atFormulaInfo()
        If m_bIsParent Then
            aszInfo(0) = Trim(txtID.Text)
            aszInfo(1) = Trim(txtName.Text)

            frmBaseInfo.AddList aszInfo, True
             
        End If
        txtParamClear
        txtID.Text = ""
        txtName.Text = ""
        If cboFormula.ListCount > 0 Then cboFormula.ListIndex = 0
        
        
    Else
        '修改
        m_oluggageSvr.EditLugFormulaInfo atFormulaInfo()
        If m_bIsParent Then
            
            aszInfo(0) = Trim(txtID.Text)
            aszInfo(1) = Trim(txtName.Text)

            frmBaseInfo.UpdateList aszInfo
        End If
        Unload Me
    End If
    
  Exit Sub
ErrorHandle:
    ShowErrorMsg
    If txtParam(1).Visible Then txtParam(1).SetFocus
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandle
    InitItemFormulaInfo '初如化
    FillCbo
    If m_eStatus = ModifyStatus Then
        RefreshFormulaInfo
        ShowFormulaInfo '显示公式信息
    Else
        ShowFormulaInfo
        txtParamClear
    End If
    '填充票价项信息
    
     AlignFormPos Me
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub InitItemFormulaInfo()
    '得到所有支持的公式信息 , 并放到模块变量中
    On Error GoTo ErrorHandle
    m_oPriceItemFunLib.Init m_oAUser
    m_aItemFormulaInfo = m_oPriceItemFunLib.LuggageFormulaInfo()
    m_nItemFormulaInfoCount = ArrayLength(m_aItemFormulaInfo)
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveFormPos Me
End Sub



Private Sub ShowFormulaInfo()
    '将对应的票价项的公式信息显示出来
    Dim i As Integer, j As Integer
    Dim szTemp As String
    Dim nTemp As Integer
    If cboFormula.Text <> "" Then
        '查找该公式在模块变量中的位置
        For i = 1 To m_nItemFormulaInfoCount
            If cboFormula.Text = m_aItemFormulaInfo(i).szFunChineseName Then Exit For
        Next
'        '设置该公式应显示的参数。
        If i <= m_nItemFormulaInfoCount Then
            txtComment.Text = m_aItemFormulaInfo(i).szFunIntroduce & vbCrLf
            For j = 1 To 10
                '设置公式的各参数的可见性
                If j > m_aItemFormulaInfo(i).nFunParamCount Then
                    txtParam(j).Visible = False
                    lblParam(j).Visible = False
                Else
                    txtParam(j).Visible = True
                    lblParam(j).Visible = True
                    txtComment.Text = txtComment.Text & vbCrLf & "参数" & j & "--" & m_aItemFormulaInfo(i).aszParamIntroduce(j)
                End If
            Next j
        End If
    End If
End Sub


Private Function GetParamErrorMsg() As String
    '判断当前的参数设置是否有效
    '无错返回空串，不过后来又改了
    Dim szFunction As String
    Dim i As Integer
    For i = 1 To m_nItemFormulaInfoCount
        If m_aItemFormulaInfo(i).szFunChineseName = cboFormula.Text Then
            szFunction = RTrim(m_aItemFormulaInfo(i).szCheckParamValidFunName)
        End If
    Next i
    On Error GoTo ErrorHandle
    If szFunction <> "" Then
'        m_oTicketPriceMan.AssertPriceItemParamIsValid szFunction, txtParam(1).Text, txtParam(2).Text, txtParam(3).Text, txtParam(4).Text, txtParam(5).Text
    End If
    GetParamErrorMsg = ""
    Exit Function
ErrorHandle:
    GetParamErrorMsg = err.Description
'    err.Raise err.Number
End Function

Private Function GetFunctionChineseName(pszFunction As String) As String
    Dim i As Integer
    For i = 1 To m_nItemFormulaInfoCount
        If pszFunction = m_aItemFormulaInfo(i).szFunName Then
            GetFunctionChineseName = m_aItemFormulaInfo(i).szFunChineseName
            Exit For
        End If
    Next
End Function

Private Function GetFunctionName(pszFunctionChineseName As String) As String
    Dim i As Integer
    For i = 1 To m_nItemFormulaInfoCount
        If pszFunctionChineseName = m_aItemFormulaInfo(i).szFunChineseName Then
            GetFunctionName = m_aItemFormulaInfo(i).szFunName
            Exit For
        End If
    Next

End Function

Private Sub FillCbo()
    Dim i As Integer
    cboFormula.clear
    For i = 1 To m_nItemFormulaInfoCount
        cboFormula.AddItem m_aItemFormulaInfo(i).szFunChineseName
    Next i
    If cboFormula.ListCount > 0 Then
        cboFormula.ListIndex = 0
    End If
End Sub

'从数据库中取出相应的票价项信息

Private Sub RefreshFormulaInfo()
    Dim szFormulaTemp As String
    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim atFormulaInfo() As TLuggageFormulaInfo
    atFormulaInfo = m_oluggageSvr.GetLugFormulaInfo(m_szFormulaId)
    Dim nCount As Integer
    nCount = ArrayLength(atFormulaInfo)
    
    If nCount > 0 Then
        txtID.Text = atFormulaInfo(1).FormulaID
        txtName.Text = atFormulaInfo(1).FormulaName
        
        '如果原先有公式设置,则公式设置为该公式
        If atFormulaInfo(1).ItemFormula = "" Then
            If cboFormula.ListCount > 0 Then
                cboFormula.ListIndex = 0
            End If
        Else
            For i = 1 To cboFormula.ListCount
                If cboFormula.List(i - 1) = GetFunctionChineseName(atFormulaInfo(1).ItemFormula) Then Exit For
            Next
            If cboFormula.ListCount > 0 Then cboFormula.Text = GetFunctionChineseName(atFormulaInfo(1).ItemFormula)
             '否则则设为第一个公式
            For i = 1 To 10
                txtParam(i).Text = atFormulaInfo(1).Param(i)
            Next i
        End If
    End If
End Sub

Private Sub txtParamClear()
    Dim i As Integer
    For i = 1 To 10
        txtParam(i).Text = 0
    Next i
    
End Sub

