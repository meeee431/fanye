VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmPriceItemAdvance 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "行包收费项设置"
   ClientHeight    =   5430
   ClientLeft      =   330
   ClientTop       =   2085
   ClientWidth     =   8295
   HelpContextID   =   7000260
   Icon            =   "frmPriceItemAdvance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtParam 
      Height          =   315
      Index           =   10
      Left            =   5355
      TabIndex        =   23
      Top             =   4995
      Width           =   1410
   End
   Begin VB.TextBox txtParam 
      Height          =   315
      Index           =   9
      Left            =   5355
      TabIndex        =   21
      Top             =   4593
      Width           =   1410
   End
   Begin VB.TextBox txtParam 
      Height          =   315
      Index           =   8
      Left            =   5355
      TabIndex        =   19
      Top             =   4198
      Width           =   1410
   End
   Begin VB.TextBox txtParam 
      Height          =   315
      Index           =   7
      Left            =   5355
      TabIndex        =   17
      Top             =   3803
      Width           =   1410
   End
   Begin VB.TextBox txtParam 
      Height          =   315
      Index           =   6
      Left            =   5355
      TabIndex        =   15
      Top             =   3408
      Width           =   1410
   End
   Begin VB.TextBox txtParam 
      Height          =   315
      Index           =   5
      Left            =   5355
      TabIndex        =   13
      Top             =   3013
      Width           =   1410
   End
   Begin VB.TextBox txtParam 
      Height          =   315
      Index           =   4
      Left            =   5355
      TabIndex        =   11
      Top             =   2618
      Width           =   1410
   End
   Begin VB.TextBox txtParam 
      Height          =   315
      Index           =   3
      Left            =   5355
      TabIndex        =   9
      Top             =   2223
      Width           =   1410
   End
   Begin VB.TextBox txtParam 
      Height          =   315
      Index           =   2
      Left            =   5355
      TabIndex        =   7
      Top             =   1828
      Width           =   1410
   End
   Begin VB.TextBox txtParam 
      Height          =   315
      Index           =   1
      Left            =   5355
      TabIndex        =   5
      Top             =   1433
      Width           =   1410
   End
   Begin VB.CheckBox chkUsed 
      BackColor       =   &H00E0E0E0&
      Caption         =   "是否使用"
      Height          =   270
      Left            =   4755
      TabIndex        =   27
      Top             =   75
      Width           =   1365
   End
   Begin VB.TextBox txtFormulaName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1515
      TabIndex        =   1
      Top             =   405
      Width           =   4305
   End
   Begin VB.ComboBox cboItemFormula 
      Height          =   300
      ItemData        =   "frmPriceItemAdvance.frx":014A
      Left            =   1845
      List            =   "frmPriceItemAdvance.frx":014C
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1095
      Width           =   2775
   End
   Begin FText.asFlatMemo txtComment 
      Height          =   3825
      Left            =   195
      TabIndex        =   28
      Top             =   1470
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6747
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonHotForeColor=   -2147483628
      ButtonHotBackColor=   -2147483632
      Registered      =   -1  'True
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   7020
      TabIndex        =   25
      Top             =   1545
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
      MICON           =   "frmPriceItemAdvance.frx":014E
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
      Left            =   7020
      TabIndex        =   24
      ToolTipText     =   "保存协议"
      Top             =   1125
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
      MICON           =   "frmPriceItemAdvance.frx":016A
      PICN            =   "frmPriceItemAdvance.frx":0186
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   315
      Left            =   7020
      TabIndex        =   26
      Top             =   2580
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "frmPriceItemAdvance.frx":0520
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数&6"
      Height          =   180
      Index           =   6
      Left            =   4755
      TabIndex        =   14
      Top             =   3480
      Width           =   450
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数&7"
      DataField       =   "7"
      Height          =   180
      Index           =   7
      Left            =   4755
      TabIndex        =   16
      Top             =   3870
      Width           =   450
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数&8"
      DataField       =   "8"
      Height          =   180
      Index           =   8
      Left            =   4755
      TabIndex        =   18
      Top             =   4260
      Width           =   450
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数&9"
      DataField       =   "9"
      Height          =   180
      Index           =   9
      Left            =   4755
      TabIndex        =   20
      Top             =   4680
      Width           =   450
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数1&0"
      Height          =   180
      Index           =   10
      Left            =   4755
      TabIndex        =   22
      Top             =   5070
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "收费项名称(N):"
      Height          =   180
      Left            =   195
      TabIndex        =   0
      Top             =   465
      Width           =   1260
   End
   Begin VB.Label lblExcuteTable 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "收费项代码:"
      Height          =   180
      Left            =   165
      TabIndex        =   32
      Top             =   150
      Width           =   990
   End
   Begin VB.Label lblPriteItemID 
      BackStyle       =   0  'Transparent
      Caption         =   "0001"
      Height          =   225
      Left            =   1245
      TabIndex        =   31
      Top             =   135
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "托运方式:"
      Height          =   180
      Left            =   2055
      TabIndex        =   30
      Top             =   105
      Width           =   810
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   30
      X2              =   8795
      Y1              =   915
      Y2              =   915
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   30
      X2              =   8795
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数&1"
      Height          =   180
      Index           =   1
      Left            =   4755
      TabIndex        =   4
      Top             =   1500
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "选择票价项公式(&F):"
      Height          =   180
      Left            =   180
      TabIndex        =   2
      Top             =   1155
      Width           =   1620
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数&2"
      Height          =   180
      Index           =   2
      Left            =   4755
      TabIndex        =   6
      Top             =   1890
      Width           =   450
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数&3"
      Height          =   180
      Index           =   3
      Left            =   4755
      TabIndex        =   8
      Top             =   2295
      Width           =   450
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数&4"
      Height          =   180
      Index           =   4
      Left            =   4755
      TabIndex        =   10
      Top             =   2700
      Width           =   450
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数&5"
      Height          =   180
      Index           =   5
      Left            =   4755
      TabIndex        =   12
      Top             =   3090
      Width           =   450
   End
   Begin VB.Label lblAcceptType 
      BackStyle       =   0  'Transparent
      Caption         =   "快件"
      Height          =   225
      Left            =   2955
      TabIndex        =   29
      Top             =   105
      Width           =   615
   End
End
Attribute VB_Name = "frmPriceItemAdvance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''**********************************************************
''* Source File Name:frmpriceItem.frm
''* Project Name:PSTLugMan.vbp
''* Engineer:
''* Date Generated:2003/01/25
''* Last Revision Date:2005/05/16
''* Brief Description:修改行包票价公式
''* Relational Document:
''**********************************************************
'
'Option Explicit
'
'Public m_bIsParent As Boolean '是否父窗体调用
'Public m_szPriceItemId As String '公式代码
'Public m_szAcceptType As Integer '托运方式
'
'Private m_bIsOneFormulaEachStation As Boolean
'
'
'
'Private Sub cmdCancel_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdHelp_Click()
'    DisplayHelp Me
'End Sub
'
'
'
'Private Sub cmdOk_Click()
'    '保存设置
'On Error GoTo ErrorHandle
'    Dim tTemp As TLuggagePriceItemFormula
'
'    tTemp.PriceItem = Trim(lblPriteItemID.Caption)
'    If lblAcceptType.Caption = szAcceptTypeGeneral Then
'        tTemp.AcceptType = 0
'    Else
'        tTemp.AcceptType = 1
'    End If
'    tTemp.PriceItemName = txtFormulaName.Text
'    If chkUsed.Value = 1 Then
'        tTemp.UsedMark = 0
'    Else
'        tTemp.UsedMark = 1
'    End If
'
'    m_oLugParam.SetPriceItem tTemp
'
'    Dim aszInfo(0 To 3) As String
'    aszInfo(0) = Trim(lblPriteItemID.Caption)
'    aszInfo(1) = Trim(txtFormulaName.Text)
'    aszInfo(2) = Trim(lblAcceptType.Caption)
'    If chkUsed.Value = 1 Then
'       aszInfo(3) = "是"
'    Else
'        aszInfo(3) = "否"
'    End If
'    frmBaseInfo.UpdateList aszInfo
'    Unload Me
'
'  Exit Sub
'ErrorHandle:
'    ShowErrorMsg
'End Sub
'
'
'Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        SendKeys "{TAB}"
'    End If
'End Sub
'
'Private Sub Form_Load()
'On Error GoTo ErrorHandle
'    FillPriceItem '填充票价项信息
'
'    AlignFormPos Me
'
'    Dim oParam As New SystemParam
'    Dim i As Integer
'    oParam.Init m_oAUser
'
'    '设置不同方式的界面布置
'    m_bIsOneFormulaEachStation = oParam.IsOneFormulaEachStation '是否每个站点一个公式.
'    If m_bIsOneFormulaEachStation Then
'        txtComment.Height = 1950
'        For i = 6 To 10
'            txtParam(i).Visible = False
'        Next i
'        Me.Height = 3990
'    Else
'        txtComment.Height = 3795
'        For i = 6 To 10
'            txtParam(i).Visible = True
'        Next i
'        Me.Height = 5880
'    End If
'    Set oParam = Nothing
'
'
'    Exit Sub
'ErrorHandle:
'    ShowErrorMsg
'End Sub
'
'
'Private Sub Form_Unload(Cancel As Integer)
'  SaveFormPos Me
'End Sub
'
'
''从数据库中取出相应的票价项信息
'
'Public Sub FillPriceItem()
'    Dim szFormulaTemp As String
'    Dim rsTemp As New Recordset
'    Dim i As Integer
'    lblPriteItemID.Caption = m_szPriceItemId
'    Set rsTemp = m_oLugParam.GetPriceItem(m_szPriceItemId, m_szAcceptType)
'    If rsTemp!accept_type = 0 Then
'        lblAcceptType.Caption = szAcceptTypeGeneral
'    Else
'        lblAcceptType.Caption = szAcceptTypeMan
'    End If
'    If rsTemp!use_mark = 0 Then
'        chkUsed.Value = 1
'    Else
'        chkUsed.Value = 0
'    End If
'    txtFormulaName.Text = rsTemp!chinese_name
'
'
'End Sub
'
'

Option Explicit
 
Public m_bIsParent As Boolean '是否父窗体调用
Public m_szPriceItemId As String '公式代码
Public m_szAcceptType As Integer '托运方式

Private m_bIsOneFormulaEachStation As Boolean


Private m_aItemFormulaInfo() As TPriceFormulaInfo '所有的公式信息(包括公式的说明,公式的中文名,参数个数,参数的说明等)
Private m_nItemFormulaInfoCount As Integer

Private Sub cboItemFormula_Click()
    ShowItemFormulaInfo
    txtParamClear
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdOk_Click()
    '保存设置]
On Error GoTo ErrorHandle
    Dim rsTemp As TLuggagePriceItemFormulaEx
    
    rsTemp.PriceItem = Trim(lblPriteItemID.Caption)
    If lblAcceptType.Caption = szAcceptTypeGeneral Then
        rsTemp.AcceptType = 0
    Else
        rsTemp.AcceptType = 1
    End If
    rsTemp.PriceItemName = txtFormulaName.Text
    rsTemp.Formula = GetFunctionName(Trim(cboItemFormula.Text))
    If chkUsed.Value = 1 Then
        rsTemp.UsedMark = 0
    Else
        rsTemp.UsedMark = 1
    End If
    rsTemp.szParam1 = txtParam(1).Text
    rsTemp.szParam2 = txtParam(2).Text
    rsTemp.szParam3 = txtParam(3).Text
    rsTemp.szParam4 = txtParam(4).Text
    rsTemp.szParam5 = txtParam(5).Text
    rsTemp.szParam6 = txtParam(6).Text
    rsTemp.szParam7 = txtParam(7).Text
    rsTemp.szParam8 = txtParam(8).Text
    rsTemp.szParam9 = txtParam(9).Text
    rsTemp.szParam10 = txtParam(10).Text
    
    m_oLugParam.SetPriceItem rsTemp
    
    Dim aszInfo(0 To 3) As String
    aszInfo(0) = Trim(lblPriteItemID.Caption)
    aszInfo(1) = Trim(txtFormulaName.Text)
    aszInfo(2) = Trim(lblAcceptType.Caption)
    If chkUsed.Value = 1 Then
       aszInfo(3) = "是"
    Else
        aszInfo(3) = "否"
    End If
    frmBaseInfo.UpdateList aszInfo
    Unload Me
    
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
    FillPriceItem '填充票价项信息
    
    ShowItemFormulaInfo '显示公式信息
    AlignFormPos Me
     
    SetFormLayout
    
     
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub SetFormLayout()


End Sub


Private Sub InitItemFormulaInfo()
    '得到所有支持的公式信息 , 并放到模块变量中
    On Error GoTo ErrorHandle
    m_aItemFormulaInfo = m_oPriceItemFunLib.LuggageFormulaInfo()
    m_nItemFormulaInfoCount = ArrayLength(m_aItemFormulaInfo)
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveFormPos Me
End Sub



Private Sub ShowItemFormulaInfo()
    '将对应的票价项的公式信息显示出来
    Dim i As Integer, j As Integer
    Dim szTemp As String
    Dim nTemp As Integer
    If cboItemFormula.Text <> "" Then
        '查找该公式在模块变量中的位置
        For i = 1 To m_nItemFormulaInfoCount
            If cboItemFormula.Text = m_aItemFormulaInfo(i).szFunChineseName Then Exit For
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
        If m_aItemFormulaInfo(i).szFunChineseName = cboItemFormula.Text Then
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



'从数据库中取出相应的票价项信息

Public Sub FillPriceItem()
    Dim szFormulaTemp As String
    Dim rsTemp As New Recordset
    Dim i As Integer
    lblPriteItemID.Caption = m_szPriceItemId
    Set rsTemp = m_oLugParam.GetPriceItem(m_szPriceItemId, m_szAcceptType)
    If rsTemp!accept_type = 0 Then
        lblAcceptType.Caption = szAcceptTypeGeneral
    Else
        lblAcceptType.Caption = szAcceptTypeMan
    End If
    If rsTemp!use_mark = 0 Then
        chkUsed.Value = 1
    Else
        chkUsed.Value = 0
    End If
    txtFormulaName.Text = rsTemp!chinese_name
    For i = 1 To m_nItemFormulaInfoCount
        cboItemFormula.AddItem m_aItemFormulaInfo(i).szFunChineseName
    Next

  
'            '如果原先有公式设置,则公式设置为该公式
         If rsTemp!item_formula = "" Then
            If cboItemFormula.ListCount > 0 Then
                cboItemFormula.ListIndex = 0
             End If
        Else
        For i = 1 To cboItemFormula.ListCount
            If cboItemFormula.List(i - 1) = GetFunctionChineseName(rsTemp!item_formula) Then Exit For
        Next
        If cboItemFormula.ListCount > 0 Then cboItemFormula.Text = GetFunctionChineseName(rsTemp!item_formula)
         '否则则设为第一个公式
        End If
   
   txtParam(1).Text = rsTemp!parameter_1
   txtParam(2).Text = rsTemp!parameter_2
   txtParam(3).Text = rsTemp!parameter_3
   txtParam(4).Text = rsTemp!parameter_4
   txtParam(5).Text = rsTemp!parameter_5
   txtParam(6).Text = rsTemp!parameter_6
   txtParam(7).Text = rsTemp!parameter_7
   txtParam(8).Text = rsTemp!parameter_8
   txtParam(9).Text = rsTemp!parameter_9
   txtParam(10).Text = rsTemp!parameter_10
    
    
End Sub

Public Sub txtParamClear()
    txtParam(1).Text = 0
    txtParam(2).Text = 0
    txtParam(3).Text = 0
    txtParam(4).Text = 0
    txtParam(5).Text = 0
    txtParam(6).Text = 0
    txtParam(7).Text = 0
    txtParam(8).Text = 0
    txtParam(9).Text = 0
    txtParam(10).Text = 0
End Sub


Private Sub txtParam_Change(Index As Integer)
    txtParam(Index) = GetTextToNumeric(txtParam(Index), True, True)
End Sub
