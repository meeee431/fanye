VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmFormula 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "拆算计算公式"
   ClientHeight    =   5790
   ClientLeft      =   4290
   ClientTop       =   2685
   ClientWidth     =   8130
   HelpContextID   =   7000250
   Icon            =   "frmFormula.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ListView lvItem 
      Height          =   4170
      Left            =   150
      TabIndex        =   20
      Top             =   1365
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   7355
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "请选择"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   3705
      TabIndex        =   17
      Top             =   915
      Width           =   2895
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1140
      TabIndex        =   16
      Top             =   915
      Width           =   975
   End
   Begin VB.PictureBox ptCaption 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   8205
      TabIndex        =   13
      Top             =   0
      Width           =   8205
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请定义折算计算公式:"
         Height          =   180
         Left            =   270
         TabIndex        =   14
         Top             =   240
         Width           =   1710
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -30
      TabIndex        =   12
      Top             =   690
      Width           =   8235
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "清除"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5220
      TabIndex        =   11
      Top             =   2910
      Width           =   705
   End
   Begin VB.TextBox txtRegFormula 
      Appearance      =   0  'Flat
      Height          =   2190
      Left            =   4635
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3330
      Width           =   3180
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1695
      Top             =   4515
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormula.frx":014A
            Key             =   "Price"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormula.frx":02A6
            Key             =   "Self"
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstFormula 
      Appearance      =   0  'Flat
      Height          =   2190
      Left            =   2490
      TabIndex        =   1
      Top             =   3330
      Width           =   2085
   End
   Begin VB.CommandButton cmdBBracket 
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4755
      TabIndex        =   10
      Top             =   2910
      Width           =   420
   End
   Begin VB.CommandButton cmdFBracket 
      Caption         =   "("
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4290
      TabIndex        =   9
      Top             =   2910
      Width           =   420
   End
   Begin VB.CommandButton cmdDivide 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3825
      TabIndex        =   8
      Top             =   2910
      Width           =   420
   End
   Begin VB.CommandButton cmdRide 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3360
      TabIndex        =   7
      Top             =   2910
      Width           =   420
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2430
      TabIndex        =   5
      Top             =   2910
      Width           =   420
   End
   Begin VB.TextBox txtFormula 
      Appearance      =   0  'Flat
      Height          =   1440
      Left            =   2430
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1380
      Width           =   4185
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   6750
      TabIndex        =   3
      Top             =   1980
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "取消(&C)"
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
      MICON           =   "frmFormula.frx":0402
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdSave 
      Height          =   315
      Left            =   6750
      TabIndex        =   4
      Top             =   1560
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
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
      MICON           =   "frmFormula.frx":041E
      PICN            =   "frmFormula.frx":043A
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
      Left            =   6750
      TabIndex        =   15
      Top             =   2370
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
      MICON           =   "frmFormula.frx":07D4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CommandButton cmdDecrease 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2895
      TabIndex        =   6
      Top             =   2910
      Width           =   420
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "公式名称:"
      Height          =   225
      Left            =   2595
      TabIndex        =   19
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "公式代码："
      Height          =   225
      Left            =   210
      TabIndex        =   18
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "frmFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"38852FDC0154"
Option Explicit
Private m_oFormula As New Formular
Private m_oReport As New Report
Private m_oSplit As New Split
Public m_szFormula As String
Public m_szFormulaName As String

Public m_state As EFormStatus
Public OldFormula As String
Public OldFormulaName As String
Dim m_aszItem() As String
Dim m_aszLsItem() As String
Dim szFormulaID As String
Dim aszTemp() As String



'Dim m_oLugFormula As New LugFormula

Private Sub cmdAdd_Click()
    txtFormula.SetFocus
    SendKeys "{+}"
End Sub

Private Sub cmdBBracket_Click()
    txtFormula.SetFocus
    SendKeys "{)}"
    
End Sub

Private Sub cmdCancel_Click()
    m_szFormula = OldFormula
    m_szFormulaName = OldFormulaName
'    frmBaseInfo.FillItemLists
    Unload Me
End Sub

Private Sub cmdClear_Click()
    txtFormula.Text = ""
    
End Sub

Private Sub cmdDecrease_Click()
    txtFormula.SetFocus
    SendKeys "{-}"
    
End Sub

'Private Sub CmdDelete_Click()
'    On Error Resume Next
'
'    If szFormulaID <> "" Then
'        oSplit.DeleteFormular szFormulaID
'        lstFormula.RemoveItem lstFormula.ListIndex
'    End If
'    If lstFormula.ListCount = 0 Then CmdDelete.Enabled = False
'End Sub

Private Sub cmdDivide_Click()
    txtFormula.SetFocus
    SendKeys "{/}"
    

End Sub

Private Sub cmdFBracket_Click()
    txtFormula.SetFocus
    SendKeys "{(}"

End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

'Private Sub cmdOk_Click()
'    m_szFormula = txtFormula.Text
'    m_szFormulaName = lstFormula.Text
'    Unload Me
'End Sub

Private Sub cmdRide_Click()
    txtFormula.SetFocus
    SendKeys "{*}"
    
End Sub

Private Sub cmdSave_Click()
    On Error GoTo err
    Dim i As Integer
    If m_state = ST_AddObj Then
        m_oFormula.AddNew
        m_oFormula.FormularID = txtID.Text
        m_oFormula.FormularName = txtName.Text
        m_oFormula.FormularContent = txtFormula.Text
        m_oFormula.Update
        frmBaseInfo.FillItemLists txtID.Text
        frmFormula.txtRegFormula = m_oFormula.FormularContent
        frmFormula.lstFormula.AddItem MakeDisplayString(txtID.Text, txtName.Text)
        txtID.Text = ""
        txtName.Text = ""
        txtFormula.Text = ""
    Else
        m_oFormula.Identify m_szFormula
        m_oFormula.FormularContent = txtFormula.Text
        m_oFormula.FormularName = txtName.Text
        m_oFormula.Update
        frmBaseInfo.FillItemLists txtID.Text
        Unload Me
        
    End If
    
    Exit Sub
err:
ShowErrorMsg
End Sub

Private Sub Form_Load()
    On Error GoTo err
    Dim nCount As Integer
    Dim i As Integer
    Dim nLen As Integer
    m_oReport.Init g_oActiveUser
    ShowSBInfo "读取公式组合条件..."
    ShowSBInfo
    FillLvItem
    m_oFormula.Init g_oActiveUser
    cmdSave.Enabled = False
    If m_state = ST_AddObj Then
        txtID.Text = ""
        txtName.Text = ""
        txtFormula.Text = ""
        txtID.Enabled = True
        txtName.Enabled = True
        frmFormula.Caption = "新增拆算计算公式"
        frmFormula.cmdSave.Caption = "新增(&A)"
    ElseIf m_state = ST_EditObj Then
        frmFormula.cmdSave.Caption = "修改(&E)"
        txtID.Enabled = False
        m_oFormula.Init g_oActiveUser
        m_oFormula.Identify m_szFormula
        txtID.Text = m_oFormula.FormularID
        txtName.Text = m_oFormula.FormularName
        txtFormula.Text = m_oFormula.FormularContent
''        FrmSaveFormular.szFormularID = m_oFormula.FormularID
''        FrmSaveFormular.szFormularName = m_oFormula.FormularName
        frmFormula.Caption = "修改拆算计算公式" & MakeDisplayString(m_oFormula.FormularID, m_oFormula.FormularName)
    End If
    FillLstFormula
    AlignFormPos Me
    Exit Sub
err:
ShowErrorMsg

End Sub
Private Sub FillLstFormula()
    On Error GoTo err
    Dim aszTemp() As String, i As Integer
    aszTemp = m_oSplit.GetFormulaItem
    If ArrayLength(aszTemp) Then
        ReDim aszTemp(1 To ArrayLength(aszTemp))
        aszTemp = m_oReport.GetAllFormula
        For i = 1 To ArrayLength(aszTemp)
            If aszTemp(i, 2) <> "" Then
                lstFormula.AddItem MakeDisplayString(aszTemp(i, 1), aszTemp(i, 2))
            End If
        Next i
    End If
    Exit Sub
err:
ShowErrorMsg
End Sub
Private Sub FillLvItem()
    On Error GoTo err
    Dim aszFormulaItemp() As String, i As Integer
    aszFormulaItemp = m_oSplit.GetFormulaItem
    If ArrayLength(aszFormulaItemp) <> 0 Then
        ReDim aszFormulaItemp(1 To ArrayLength(aszFormulaItemp), 1 To 2)
        aszFormulaItemp = m_oSplit.GetFormulaItem
        For i = 1 To ArrayLength(aszFormulaItemp)
            If aszFormulaItemp(i, 1) <> "[票价]" Then
                lvItem.ListItems.Add , , aszFormulaItemp(i, 1)
            End If
        Next i
    End If
    Exit Sub
err:
ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Sub lstFormula_Click()
    Dim szTemp() As String
    szFormulaID = ResolveDisplay(lstFormula.Text)
    m_oFormula.Init g_oActiveUser
    m_oFormula.Identify szFormulaID
    txtRegFormula.Text = m_oFormula.FormularContent
End Sub

Private Sub lstFormula_DblClick()
    txtFormula.Text = txtRegFormula.Text
'    FrmSaveFormular.szFormularContent = txtRegFormula.Text
'    FrmSaveFormular.szFormularID = ResolveDisplay(lstFormula.Text)
'    FrmSaveFormular.szFormularName = ResolveDisplayEx(lstFormula.Text)
'    Dim i As Integer
'    Dim szFirstPart As String, szLastPart As String
'    i = txtFormula.SelStart
'    If i = 0 Then txtFormula.SelStart = Len(txtFormula.Text)
'    szFirstPart = Left(txtFormula.Text, i)
'    szLastPart = Mid(txtFormula.Text, i + 1)
'    txtFormula.SetFocus
'    txtFormula.Text = szFirstPart & txtRegFormula.Text & szLastPart
'    txtFormula.SelStart = i + Len(txtRegFormula.Text)
End Sub

Private Sub lvItem_DblClick()
    Dim i As Integer
    Dim szFirstPart As String, szLastPart As String
    i = txtFormula.SelStart
    If i = 0 Then txtFormula.SelStart = Len(txtFormula.Text)
    szFirstPart = Left(txtFormula.Text, i)
    szLastPart = Mid(txtFormula.Text, i + 1)
    txtFormula.SetFocus
    txtFormula.Text = szFirstPart & lvItem.SelectedItem.Text & szLastPart
    txtFormula.SelStart = i + Len(lvItem.SelectedItem.Text)
    
End Sub

'从注册表中将公式读出
Private Sub LoadRegFormula(Optional bLstEnable As Boolean = False)

     cmdSave.Enabled = False
End Sub

'从注册表中写入公式
Private Sub SaveRegFormula()
'    FrmSaveFormular.Show vbModal

End Sub

Private Sub txtID_Change()
    FormatTextBoxBySize txtID, 4
    If txtID.Text = "" Or txtName.Text = "" Or txtFormula.Text = "" Then
        cmdSave.Enabled = False
    Else
        cmdSave.Enabled = True
    End If
    
End Sub

Private Sub txtName_Change()
    FormatTextBoxBySize txtName, 50
    If txtID.Text = "" Or txtName.Text = "" Or txtFormula.Text = "" Then
        cmdSave.Enabled = False
    Else
        cmdSave.Enabled = True
    End If
End Sub

Private Sub txtFormula_Change()
    If txtID.Text = "" Or txtName.Text = "" Or txtFormula.Text = "" Then
        cmdSave.Enabled = False
    Else
        cmdSave.Enabled = True
    End If
End Sub
