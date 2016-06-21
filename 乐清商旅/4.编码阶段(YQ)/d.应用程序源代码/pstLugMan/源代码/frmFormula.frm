VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmFormula 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "拆算计算公式"
   ClientHeight    =   5100
   ClientLeft      =   2730
   ClientTop       =   2760
   ClientWidth     =   7965
   HelpContextID   =   7000250
   Icon            =   "frmFormula.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox ptCaption 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   7965
      TabIndex        =   15
      Top             =   0
      Width           =   7965
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请定义折算计算公式:"
         Height          =   180
         Left            =   270
         TabIndex        =   16
         Top             =   240
         Width           =   1710
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   0
      TabIndex        =   14
      Top             =   690
      Width           =   8115
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   315
      Left            =   6630
      TabIndex        =   4
      Top             =   840
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "确定(&0)"
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
      MICON           =   "frmFormula.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      Left            =   5100
      TabIndex        =   13
      Top             =   2370
      Width           =   705
   End
   Begin VB.TextBox txtRegFormula 
      Appearance      =   0  'Flat
      Height          =   2190
      Left            =   4515
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2790
      Width           =   3180
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1575
      Top             =   3975
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
            Picture         =   "frmFormula.frx":0166
            Key             =   "Price"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormula.frx":02C2
            Key             =   "Self"
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstFormula 
      Appearance      =   0  'Flat
      Height          =   2190
      Left            =   2325
      TabIndex        =   2
      Top             =   2790
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
      Left            =   4635
      TabIndex        =   12
      Top             =   2370
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
      Left            =   4170
      TabIndex        =   11
      Top             =   2370
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
      Left            =   3705
      TabIndex        =   10
      Top             =   2370
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
      Left            =   3240
      TabIndex        =   9
      Top             =   2370
      Width           =   420
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
      Left            =   2775
      TabIndex        =   8
      Top             =   2370
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
      Left            =   2310
      TabIndex        =   7
      Top             =   2370
      Width           =   420
   End
   Begin VB.TextBox txtFormula 
      Appearance      =   0  'Flat
      Height          =   1440
      Left            =   2310
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   840
      Width           =   4185
   End
   Begin MSComctlLib.ListView lvItem 
      Height          =   4140
      Left            =   90
      TabIndex        =   0
      Top             =   840
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   7303
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "公式项"
         Object.Width           =   2910
      EndProperty
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   6630
      TabIndex        =   5
      Top             =   1620
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
      MICON           =   "frmFormula.frx":041E
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
      Left            =   6630
      TabIndex        =   6
      Top             =   1230
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
      MICON           =   "frmFormula.frx":043A
      PICN            =   "frmFormula.frx":0456
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
      Left            =   6630
      TabIndex        =   17
      Top             =   2010
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
      MICON           =   "frmFormula.frx":07F0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"38852FDC0154"
Option Explicit
Public m_szFormula As String
Public m_szFormulaName As String

Public m_eStatus As eFormStatus
Public OldFormula As String
Public OldFormulaName As String
Dim m_aszItem() As String
Dim m_aszLsItem() As String
Dim szFormulaID As String
Dim aszTemp() As String

'Dim m_oLugFormula As New LugFormula

Private Sub CmdAdd_Click()
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

Private Sub cmdOk_Click()
    m_szFormula = txtFormula.Text
    m_szFormulaName = lstFormula.Text
    Unload Me
End Sub

Private Sub cmdRide_Click()
    txtFormula.SetFocus
    SendKeys "{*}"
End Sub

Private Sub cmdSave_Click()
'    LoadRegFormula
    SaveRegFormula
End Sub

Private Sub Form_Load()
    Dim nCount As Integer
    Dim i As Integer
    Dim nlen As Integer
  
        ShowSBInfo "读取公式组合条件..."
        ShowSBInfo
        nlen = ArrayLength(m_oLugFormula.GetAllFormulaItems)
        If nlen > 0 Then
           ReDim m_aszItem(1 To nlen, 1 To 2)
           m_aszItem = m_oLugFormula.GetAllFormulaItems
           For i = 1 To nlen
                If m_aszItem(i, 1) <> "" Then
                    lvItem.ListItems.Add , , m_aszItem(i, 1)
                End If
            Next i
        End If
        LoadRegFormula
'    txtFormula.Text = OldFormula
End Sub

Private Sub lstFormula_Click()
    Dim szTemp() As String
    szFormulaID = LeftAndRight(lstFormula.Text, True, "[")
    m_oLugFormula.Identify szFormulaID
'    szTemp = m_oLugFormula.GetAllFormulas()
'    If ArrayLength(szTemp) = 0 Then Exit Sub
    txtRegFormula.Text = m_oLugFormula.FormulaText
End Sub

Private Sub lstFormula_DblClick()
    Dim i As Integer
    Dim szFirstPart As String, szLastPart As String
    i = txtFormula.SelStart
    If i = 0 Then txtFormula.SelStart = Len(txtFormula.Text)
    szFirstPart = Left(txtFormula.Text, i)
    szLastPart = Mid(txtFormula.Text, i + 1)
    txtFormula.SetFocus
    txtFormula.Text = szFirstPart & txtRegFormula.Text & szLastPart
    txtFormula.SelStart = i + Len(txtRegFormula.Text)
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
   Select Case m_eStatus
          Dim i As Integer, nCount As Integer
        Case ST_AddObj
          
            
            aszTemp = m_oLugFormula.GetAllFormulas
            nCount = ArrayLength(aszTemp)
            For i = 1 To nCount
              lstFormula.AddItem aszTemp(i, 1) & "[" & aszTemp(i, 2) & "]"
            Next
             lstFormula.Enabled = bLstEnable
              cmdOk.Visible = False
        Case ST_EditObj
             m_oLugFormula.Identify Trim(frmBaseInfo.lvObject.SelectedItem.Text)
              lstFormula.clear
              lstFormula.AddItem m_oLugFormula.FormulaID & "[" & m_oLugFormula.FormulaName & "]"
              lstFormula.Enabled = bLstEnable
              txtRegFormula.Enabled = bLstEnable
              txtFormula.Text = m_oLugFormula.FormulaText
              txtRegFormula.Text = m_oLugFormula.FormulaText
              FrmSaveFormular.TxtFormularID.Enabled = bLstEnable
              FrmSaveFormular.TxtFormularID = m_oLugFormula.FormulaID
              FrmSaveFormular.TxtFormularName = m_oLugFormula.FormulaName
               cmdOk.Visible = False
        Case ST_NormalObj
      
            aszTemp = m_oLugFormula.GetAllFormulas
            nCount = ArrayLength(aszTemp)
            For i = 1 To nCount
              lstFormula.AddItem aszTemp(i, 1) & "[" & aszTemp(i, 2) & "]"
            Next
            
 
    End Select
   
     cmdSave.Enabled = False
End Sub

'从注册表中写入公式
Private Sub SaveRegFormula()
 FrmSaveFormular.FormularContent = txtFormula.Text
  
  FrmSaveFormular.Show vbModal

End Sub


Private Sub txtFormula_Change()
   cmdSave.Enabled = True
End Sub
