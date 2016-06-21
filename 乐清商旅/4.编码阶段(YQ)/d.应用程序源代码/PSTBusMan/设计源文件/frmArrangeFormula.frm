VERSION 5.00
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{BBF95DAA-F9CB-4CA9-A673-E0E9E0193752}#1.0#0"; "STSellCtl.ocx"
Begin VB.Form frmArrangeFormula 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "票价公式属性"
   ClientHeight    =   4725
   ClientLeft      =   2565
   ClientTop       =   2535
   ClientWidth     =   8940
   Icon            =   "frmArrangeFormula.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboPriceTable 
      Height          =   300
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   2595
   End
   Begin VB.TextBox txtFormulaName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1260
      TabIndex        =   3
      Top             =   450
      Width           =   2595
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2595
      Left            =   1920
      TabIndex        =   17
      Top             =   1560
      Width           =   6860
      Begin FText.asFlatMemo txtComment 
         Height          =   1905
         Left            =   120
         TabIndex        =   25
         Top             =   570
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   3360
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
      Begin VB.ListBox cboParam 
         Appearance      =   0  'Flat
         Height          =   660
         Left            =   5265
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   240
         Width           =   1415
      End
      Begin VB.ComboBox cboItemFormula 
         Height          =   315
         ItemData        =   "frmArrangeFormula.frx":014A
         Left            =   1800
         List            =   "frmArrangeFormula.frx":014C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   180
         Width           =   2775
      End
      Begin STSellCtl.ucNumTextBox txtParam 
         Height          =   315
         Index           =   1
         Left            =   5265
         TabIndex        =   24
         Top             =   585
         Visible         =   0   'False
         Width           =   1415
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
         Left            =   5265
         TabIndex        =   20
         Top             =   2160
         Visible         =   0   'False
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
         Left            =   5265
         TabIndex        =   18
         Top             =   1785
         Visible         =   0   'False
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
         Left            =   5265
         TabIndex        =   16
         Top             =   1380
         Visible         =   0   'False
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
         Left            =   5265
         TabIndex        =   14
         Top             =   990
         Visible         =   0   'False
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
      Begin VB.Label lblParam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参数&1"
         Height          =   180
         Index           =   1
         Left            =   4755
         TabIndex        =   11
         Top             =   645
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "选择票价项公式(&F):"
         Height          =   180
         Left            =   135
         TabIndex        =   9
         Top             =   240
         Width           =   1620
      End
      Begin VB.Label lblParam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参数&2"
         Height          =   180
         Index           =   2
         Left            =   4755
         TabIndex        =   13
         Top             =   1035
         Width           =   450
      End
      Begin VB.Label lblParam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参数&3"
         Height          =   180
         Index           =   3
         Left            =   4755
         TabIndex        =   15
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label lblParam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参数&4"
         Height          =   180
         Index           =   4
         Left            =   4755
         TabIndex        =   26
         Top             =   1845
         Width           =   450
      End
      Begin VB.Label lblParam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参数&5"
         Height          =   180
         Index           =   5
         Left            =   4755
         TabIndex        =   19
         Top             =   2235
         Width           =   450
      End
   End
   Begin RTComctl3.CoolButton cmdOK 
      Height          =   315
      Left            =   4680
      TabIndex        =   21
      Top             =   4290
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmArrangeFormula.frx":014E
      PICN            =   "frmArrangeFormula.frx":016A
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
      CausesValidation=   0   'False
      Height          =   315
      Left            =   5925
      TabIndex        =   22
      Top             =   4290
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmArrangeFormula.frx":0504
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
      CausesValidation=   0   'False
      Height          =   315
      Left            =   7185
      TabIndex        =   23
      Top             =   4290
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmArrangeFormula.frx":0520
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lstItem 
      Appearance      =   0  'Flat
      Height          =   2490
      IntegralHeight  =   0   'False
      ItemData        =   "frmArrangeFormula.frx":053C
      Left            =   120
      List            =   "frmArrangeFormula.frx":053E
      TabIndex        =   8
      Top             =   1650
      Width           =   1695
   End
   Begin VB.CheckBox chkDefault 
      BackColor       =   &H00E0E0E0&
      Caption         =   "缺省票价计算公式(&D)"
      Height          =   270
      Left            =   4140
      TabIndex        =   4
      Top             =   480
      Width           =   2025
   End
   Begin VB.TextBox txtAnno 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1260
      TabIndex        =   6
      Top             =   850
      Width           =   7520
   End
   Begin VB.Label lblExcuteTable 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "票价表(&T):"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   75
      X2              =   8840
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公式名称(N):"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   510
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "选择票价项(&I):"
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   1400
      Width           =   1260
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   75
      X2              =   8840
      Y1              =   1250
      Y2              =   1250
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "注释(A):"
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   910
      Width           =   720
   End
End
Attribute VB_Name = "frmArrangeFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'* Source File Name:frmArrangeFormual.frm
'* Project Name:PSTBusMan.vbp
'* Engineer:陈峰
'* Date Generated:2002/09/03
'* Last Revision Date:2002/09/03
'* Brief Description:修改票价公式
'* Relational Document:
'**********************************************************

Option Explicit
'********此处的cmdOk_click 与FillVehicleType 及FillCheckGate均可以修改 ,而且算法很差,现在由于时间原因,所以未改

Public m_eStatus As EFormStatus
Public m_szFormulaID As String '公式代码
Public m_bIsParent As Boolean '是否父窗体调用

Private m_oTicketPriceMan As New TicketPriceMan
Private m_atFormulaParam(1 To 16) As TItemFAndP '该公式的信息及16个票价项公式参数
Private m_aItemFormulaInfo() As TPriceFormulaInfo '所有的公式信息(包括公式的说明,公式的中文名,参数个数,参数的说明等)
Private m_nItemFormulaInfoCount As Integer

Private m_bLastItemIsBaseCarriage As Boolean '最近选择的票价项是否是基本运价
Private m_nLastItemIndex As Integer '最近选择的票价项
Private m_bPriceTableChanged As Boolean '票价表变化

Private Sub cboItemFormula_Click()
    ShowItemFormulaInfo
End Sub


Private Sub CboPriceTable_Click()
'    m_bPriceTableChanged = True
    m_nLastItemIndex = -1
    GetFormula
    FillPriceItem
    ShowItemFormulaInfo
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdOk_Click()
    '保存设置
    On Error GoTo ErrorHandle
    Dim nTemp As Integer
'    Dim n As Integer
    Dim szTemp As String
    Dim i As Integer, j As Integer
    Dim szFuntion As String
    Dim szErrMsg As String
    Dim bIsDefault As Boolean '是否是缺省公式
    Dim oFormula As New TicketPriceFormula
    Dim afpFAndP(1 To 16) As TItemFAndP '为了保存用
    '将当前设置的值,赋予模块变量
    nTemp = CInt(ResolveDisplay(lstItem.List(m_nLastItemIndex))) + 1
    m_atFormulaParam(nTemp).szFormula = cboItemFormula.Text
    m_atFormulaParam(nTemp).szPriceItem = GetFunctionName(cboItemFormula.Text)
    If m_atFormulaParam(nTemp).szPriceItem = "VehicleModelCharge" Or m_atFormulaParam(nTemp).szPriceItem = "BaseCarriagePerKm" Or m_atFormulaParam(nTemp).szPriceItem = "CheckGateCharge" Or m_atFormulaParam(nTemp).szPriceItem = "SpringBaseDistVTypeAddCharge" Or m_atFormulaParam(nTemp).szPriceItem = "VehicleModelFarDistanceAddCharge" Then
        szTemp = ""
        For i = 0 To cboParam.ListCount - 1
            If cboParam.Selected(i) = True Then
               If szTemp <> "" Then szTemp = szTemp & ","
               szTemp = szTemp & ResolveDisplay(Trim(cboParam.List(i)))
            End If
        Next i
        If szTemp = "" Then szTemp = "1000"
        m_atFormulaParam(nTemp).sgParam1 = szTemp
    Else
        m_atFormulaParam(nTemp).sgParam1 = txtParam(1).Text
    End If
    m_atFormulaParam(nTemp).sgParam2 = txtParam(2).Text
    
    m_atFormulaParam(nTemp).sgParam3 = txtParam(3).Text
    m_atFormulaParam(nTemp).sgParam4 = txtParam(4).Text
    m_atFormulaParam(nTemp).sgParam5 = txtParam(5).Text
    
    If m_eStatus = EFS_AddNew Then
        For i = 1 To lstItem.ListCount
            nTemp = Int(ResolveDisplay(lstItem.List(i - 1))) + 1
            szTemp = Trim(m_atFormulaParam(nTemp).szFormula)
            For j = 1 To m_nItemFormulaInfoCount
                If szTemp = m_aItemFormulaInfo(j).szFunChineseName Then
                    szFuntion = m_aItemFormulaInfo(j).szCheckParamValidFunName
                    Exit For
                End If
            Next
            If szFuntion <> "" Then
                '检查参数的有效性
                m_oTicketPriceMan.AssertPriceItemParamIsValid szFuntion, m_atFormulaParam(nTemp).sgParam1, _
                m_atFormulaParam(nTemp).sgParam2, m_atFormulaParam(nTemp).sgParam3, m_atFormulaParam(nTemp).sgParam4, _
                m_atFormulaParam(nTemp).sgParam5
            End If
        Next i
    Else
        szErrMsg = GetParamErrorMsg()
        If szErrMsg <> "" Then
            ShowMsg szErrMsg
            Exit Sub
        End If
    End If
    '修改公式,保存到数据库
    oFormula.Init g_oActiveUser
    If m_eStatus = EFS_AddNew Then
        '如果为新增
        oFormula.AddNew
        oFormula.FormulaName = txtFormulaName.Text
        oFormula.Annotation = txtAnno.Text
        oFormula.Update
        If m_bIsParent Then frmFormulaMan.AddList txtFormulaName.Text
    ElseIf m_eStatus = EFS_Modify Then
        oFormula.Identify txtFormulaName.Text
        oFormula.Annotation = txtAnno.Text
        bIsDefault = IIf(chkDefault.Value = vbChecked, True, False)
        oFormula.Update
        '如果为缺省公式,则 设置该公式为缺省公式
        If bIsDefault Then oFormula.SetAsDefault
        If m_bIsParent Then frmFormulaMan.UpdateList txtFormulaName.Text
    End If
    
    
On Error Resume Next
    szErrMsg = ""
    For i = 1 To 16
        
        afpFAndP(i).szPriceItem = Format(i - 1, "0000")
        afpFAndP(i).szFormula = m_atFormulaParam(i).szPriceItem '
        afpFAndP(i).sgParam1 = IIf(m_atFormulaParam(i).sgParam1 = "", 0, m_atFormulaParam(i).sgParam1)
        afpFAndP(i).sgParam2 = IIf(m_atFormulaParam(i).sgParam2 = "", 0, m_atFormulaParam(i).sgParam2)
        afpFAndP(i).sgParam3 = IIf(m_atFormulaParam(i).sgParam3 = "", 0, m_atFormulaParam(i).sgParam3)
        afpFAndP(i).sgParam4 = IIf(m_atFormulaParam(i).sgParam4 = "", 0, m_atFormulaParam(i).sgParam4)
        afpFAndP(i).sgParam5 = IIf(m_atFormulaParam(i).sgParam5 = "", 0, m_atFormulaParam(i).sgParam5)
        oFormula.ModifyItemFAndP afpFAndP(i), ResolveDisplay(cboPriceTable.Text), Trim(txtFormulaName.Text)
        If err Then szErrMsg = EncodeString(afpFAndP(i).szPriceItem) & ":" & err.Description & vbCrLf
    Next
    If szErrMsg <> "" Then MsgBox szErrMsg, vbExclamation, "警告"
    
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
    m_nLastItemIndex = -1
    m_bPriceTableChanged = False
    m_oTicketPriceMan.Init g_oActiveUser
    If m_eStatus = EFS_AddNew Then
        txtFormulaName.Enabled = True
    Else
        txtFormulaName.Enabled = False
        txtFormulaName.Text = m_szFormulaID
    End If
    InitItemFormulaInfo
    FillPriceTable
    GetFormula
    FillItemFormual
    ShowItemFormulaInfo
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub FillPriceItem()
    '填充所有的可用的票价项
    Dim i As Integer
    Dim aszPriceItem() As String
    Dim Count As Integer
    On Error GoTo ErrorHandle
    lstItem.Clear
    '得到票价项
    aszPriceItem = m_oTicketPriceMan.GetAllTicketItem
    Count = ArrayLength(aszPriceItem)
    For i = 1 To Count
        If aszPriceItem(i, 3) = TP_PriceItemUse Then lstItem.AddItem MakeDisplayString(aszPriceItem(i, 1), aszPriceItem(i, 2))
    Next
    If lstItem.ListCount > 0 Then
        lstItem.ListIndex = 0
    End If
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub FillItemFormual()
    Dim i As Integer ', j As Integer
    Dim bBaseCarriage As Boolean '是否是基本运价项
    Dim szitem As String
    Dim szTemp As String
    Dim nTemp As Integer
    If lstItem.ListIndex >= 0 Then

        szitem = ResolveDisplay(lstItem.List(lstItem.ListIndex))
        nTemp = CInt(szitem) + 1
        If szitem = cszItemBaseCarriage Then
            bBaseCarriage = True
        Else
            bBaseCarriage = False
        End If
        If bBaseCarriage <> m_bLastItemIsBaseCarriage Then
            '现在的是否基本运价项与原先票价项的是否基本运价项不一致,则需要重新填充公式
            cboItemFormula.Clear
            If bBaseCarriage Then
                '如果是基本运价项,则填充基本运价项的公式
                For i = 1 To m_nItemFormulaInfoCount
                    If m_aItemFormulaInfo(i).bBaseCarriage Then
                        cboItemFormula.AddItem m_aItemFormulaInfo(i).szFunChineseName
                    End If
                Next
            Else
                '填充非基本运价项的公式
                For i = 1 To m_nItemFormulaInfoCount
                    If Not m_aItemFormulaInfo(i).bBaseCarriage Then
                        cboItemFormula.AddItem m_aItemFormulaInfo(i).szFunChineseName
                    End If
                Next
            End If
            m_bLastItemIsBaseCarriage = bBaseCarriage
        End If

        szTemp = m_atFormulaParam(CInt(szitem) + 1).szFormula
        If szTemp <> "" Then
            '如果原先有公式设置,则公式设置为该公式
            For i = 1 To cboItemFormula.ListCount
                If cboItemFormula.List(i - 1) = szTemp Then Exit For
            Next
            If cboItemFormula.ListCount > 0 Then cboItemFormula.ListIndex = i - 1
        Else
            '否则则设为第一个公式
            If cboItemFormula.ListCount > 0 Then
                cboItemFormula.ListIndex = 0
            End If
        End If
    End If

End Sub

Private Sub InitItemFormulaInfo()
    '得到所有支持的公式信息,并放到模块变量中
    On Error GoTo ErrorHandle
    m_aItemFormulaInfo = m_oTicketPriceMan.GetPriceItemFormulaInfo()
    m_nItemFormulaInfoCount = ArrayLength(m_aItemFormulaInfo)
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim m_atTemp() As TPriceFormulaInfo '清空内存数据用
    Dim m_atTemp2 As TItemFAndP '清空内存数据用
    Dim i As Integer
    
    Set m_oTicketPriceMan = Nothing
    m_bIsParent = False
    m_bLastItemIsBaseCarriage = False
    m_bPriceTableChanged = False
    
    '清空内存数据
    m_aItemFormulaInfo = m_atTemp
    For i = 1 To 16
        m_atFormulaParam(i) = m_atTemp2
    Next i
End Sub

Private Sub lstItem_Click()
    ItemChanged
End Sub

Private Sub ShowItemFormulaInfo()
    '将对应的票价项的公式信息显示出来
    Dim i As Integer, j As Integer
    Dim szTemp As String
    Dim szVehicleGate() As String
    Dim nTemp As Integer
    If cboItemFormula.Text <> "" Then
        '查找该公式在模块变量中的位置
        For i = 1 To m_nItemFormulaInfoCount
            If cboItemFormula.Text = m_aItemFormulaInfo(i).szFunChineseName Then Exit For
        Next
        '设置该公式应显示的参数。
        If i <= m_nItemFormulaInfoCount Then
            txtComment.Text = m_aItemFormulaInfo(i).szFunIntroduce & vbCrLf
            For j = 1 To 5
                '设置公式的各参数的可见性
                If j > m_aItemFormulaInfo(i).nFunParamCount Then
                    txtParam(j).Visible = False
                    lblParam(j).Visible = False
                Else
                    txtParam(j).Visible = True
                    lblParam(j).Visible = True
                    txtComment.Text = txtComment.Text & vbCrLf & "参数" & j & "--" & m_aItemFormulaInfo(i).aszParamIntroduce(j)
                End If
                If j = 1 Then
                    szTemp = GetFunctionName(cboItemFormula.Text)
                    If lstItem.ListIndex >= 0 Then nTemp = CInt(ResolveDisplay(lstItem.List(lstItem.ListIndex))) + 1
                    If szTemp = "VehicleModelCharge" Or szTemp = "BaseCarriagePerKm" Or szTemp = "SpringBaseDistVTypeAddCharge" Or szTemp = "VehicleModelFarDistanceAddCharge" Then
                        '如果公式为[根据车型计算费用]或[每公里基本运价]或[基本运价和加成费按车型加价春运费]
                        cboParam.Visible = True
                        If Trim(m_atFormulaParam(nTemp).sgParam1) <> "" Then szVehicleGate = StringToTeam(m_atFormulaParam(nTemp).sgParam1)
                        FillVehicleType szVehicleGate
                        txtParam(1).Visible = False
                    ElseIf szTemp = "CheckGateCharge" Then
                        '如果公式为[根据检票口计算费用]
                        cboParam.Visible = True
                        If Trim(m_atFormulaParam(nTemp).sgParam1) <> "" Then szVehicleGate = StringToTeam(m_atFormulaParam(nTemp).sgParam1)
                        FillCheckGate szVehicleGate
                        txtParam(1).Visible = False
                    Else
                        '否则,则不可见
                        cboParam.Visible = False
                    End If
                End If
            Next j
        End If
    End If
End Sub

'用户改变了选中的票价项
Private Function ItemChanged()
    Dim szErrMsg As String
    Dim i As Integer
    Dim nTemp As Integer
    Dim szTemp As String
    Dim szVehicleGate() As String
    On Error GoTo ErrorHandle
    
    If (lstItem.ListCount <> m_nLastItemIndex And lstItem.ListIndex <> m_nLastItemIndex) Or m_bPriceTableChanged Then
        '如果票价项选择改变了,或者票价表选择改变了
        If m_nLastItemIndex >= 0 And m_bPriceTableChanged = False Then
            '如果不是第一次显示
            szErrMsg = GetParamErrorMsg()
            If szErrMsg <> "" Then
                '先前的设置不合法,则返回到原来的票价项重新设置
                ShowMsg szErrMsg
                lstItem.ListIndex = m_nLastItemIndex
                Exit Function
            Else
                '如果设置合法，则将信息保存到内存变量
                nTemp = CInt(ResolveDisplay(lstItem.List(m_nLastItemIndex))) + 1
                If m_bPriceTableChanged = False Then
                    '如果不是票价表选择改变
                    m_atFormulaParam(nTemp).szFormula = cboItemFormula.Text
                    m_atFormulaParam(nTemp).szPriceItem = GetFunctionName(cboItemFormula.Text)
                    If txtParam(1).Visible = False Then
                        '如果为检票口或车型
                        szTemp = ""
                        For i = 0 To cboParam.ListCount - 1
                            If cboParam.Selected(i) Then
                                If szTemp <> "" Then szTemp = szTemp & ","
                                szTemp = szTemp & ResolveDisplay(cboParam.List(i))
                            End If
                        Next i
                        If szTemp = "" Then szTemp = "1000"
                        m_atFormulaParam(nTemp).sgParam1 = szTemp
                    Else
                        m_atFormulaParam(nTemp).sgParam1 = Val(txtParam(1).Text)
                    End If
                    m_atFormulaParam(nTemp).sgParam2 = Val(txtParam(2).Text)
                    m_atFormulaParam(nTemp).sgParam3 = Val(txtParam(3).Text)
                    m_atFormulaParam(nTemp).sgParam4 = Val(txtParam(4).Text)
                    m_atFormulaParam(nTemp).sgParam5 = Val(txtParam(5).Text)
                Else
                    m_bPriceTableChanged = False
                End If
            End If
        End If
        '将内存变量中的信息显示在对应的控件中
        FillItemFormual
        nTemp = CInt(ResolveDisplay(lstItem.List(lstItem.ListIndex))) + 1
        szTemp = GetFunctionName(m_atFormulaParam(nTemp).szFormula)
        If szTemp = "BaseCarriagePerKm" Or szTemp = "VehicleModelCharge" Or szTemp = "CheckGateCharge" Or szTemp = "SpringBaseDistVTypeAddCharge" Or szTemp = "VehicleModelFarDistanceAddCharge" Then
            '如果公式为[根据车型计算费用]或[每公里基本运价]或[基本运价和加成费按车型加价春运费]
            If Trim(m_atFormulaParam(nTemp).sgParam1) <> "" Then
                szVehicleGate = StringToTeam(m_atFormulaParam(nTemp).sgParam1)
            Else
                txtParam(1).Text = 0
            End If
        Else
            If Trim(m_atFormulaParam(nTemp).sgParam1) <> "" Then
                txtParam(1).Text = m_atFormulaParam(nTemp).sgParam1
            Else
                txtParam(1).Text = 0
            End If
        End If
    
        If szTemp = "BaseCarriagePerKm" Or szTemp = "VehicleModelCharge" Or szTemp = "SpringBaseDistVTypeAddCharge" Or szTemp = "VehicleModelFarDistanceAddCharge" Then
            '如果公式为[根据车型计算费用]或[每公里基本运价]或[基本运价和加成费按车型加价春运费]
            '显示车型信息
            FillVehicleType szVehicleGate
            txtParam(1).Visible = False
        ElseIf szTemp = "CheckGateCharge" Then
            '如果公式为[根据检票口计算费用]
            '显示检票口信息
            FillCheckGate szVehicleGate
            txtParam(1).Visible = False
        End If
        '设置其余各参数的值
        If Trim(m_atFormulaParam(nTemp).sgParam2) <> "" Then
            txtParam(2).Text = m_atFormulaParam(nTemp).sgParam2
        Else
            txtParam(2).Text = 0
        End If
        If Trim(m_atFormulaParam(nTemp).sgParam3) <> "" Then
            txtParam(3).Text = m_atFormulaParam(nTemp).sgParam3
        Else
            txtParam(3).Text = 0
        End If
        If Trim(m_atFormulaParam(nTemp).sgParam4) <> "" Then
            txtParam(4).Text = m_atFormulaParam(nTemp).sgParam4
        Else
            txtParam(4).Text = 0
        End If
        If Trim(m_atFormulaParam(nTemp).sgParam5) <> "" Then
            txtParam(5).Text = m_atFormulaParam(nTemp).sgParam5
        Else
            txtParam(5).Text = 0
        End If
        m_nLastItemIndex = lstItem.ListIndex
    End If
    Exit Function
ErrorHandle:
    ShowErrorMsg
    If txtParam(1).Visible Then txtParam(1).SetFocus
    lstItem.ListIndex = m_nLastItemIndex
End Function

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
        m_oTicketPriceMan.AssertPriceItemParamIsValid szFunction, txtParam(1).Text, txtParam(2).Text, txtParam(3).Text, txtParam(4).Text, txtParam(5).Text
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

Private Sub FillPriceTable()
    '填充票价表
    
    Dim aszRoutePriceTable() As String
    Dim i As Integer, nCount As Integer
    Dim szPriceTable As String
    On Error GoTo ErrorHandle
    aszRoutePriceTable = GetPriceTable(Now) 'GetProjectExcutePriceTable(g_szExePlanID)
    nCount = ArrayLength(aszRoutePriceTable)
    cboPriceTable.Clear
    If nCount > 0 Then
        For i = 1 To nCount
            szPriceTable = MakeDisplayString(aszRoutePriceTable(i, 1), aszRoutePriceTable(i, 2))
            cboPriceTable.AddItem szPriceTable
            If aszRoutePriceTable(i, 7) = cnRunTable Then cboPriceTable.Text = MakeDisplayString(aszRoutePriceTable(i, 1), aszRoutePriceTable(i, 2))
        Next
    End If
    

    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub GetFormula()
''从数据库中得到公式信息,放在模块变量中
On Error GoTo ErrorHandle

    Dim oFormula As New TicketPriceFormula
    Dim aifpTemp() As TItemFAndP, i As Integer
    If txtFormulaName.Text = "" Then Exit Sub
    oFormula.Init g_oActiveUser
    oFormula.Identify txtFormulaName.Text
    aifpTemp = oFormula.GetAllFAndP(ResolveDisplay(cboPriceTable.Text))
    If ArrayLength(aifpTemp) > 0 Then
        For i = 1 To 16
            m_atFormulaParam(i).sgParam1 = aifpTemp(i).sgParam1
            m_atFormulaParam(i).sgParam2 = aifpTemp(i).sgParam2
            m_atFormulaParam(i).sgParam3 = aifpTemp(i).sgParam3
            m_atFormulaParam(i).sgParam4 = aifpTemp(i).sgParam4
            m_atFormulaParam(i).sgParam5 = aifpTemp(i).sgParam5
            '特别注意以下两句
            m_atFormulaParam(i).szFormula = GetFunctionChineseName(aifpTemp(i).szFormula)
            m_atFormulaParam(i).szPriceItem = aifpTemp(i).szFormula
        Next
    End If
    chkDefault.Value = IIf(oFormula.IsDefault = IsDefaultFormula, vbChecked, vbUnchecked)
    txtAnno.Text = oFormula.Annotation
    Set oFormula = Nothing
    
    Exit Sub

ErrorHandle:
    ShowErrorMsg
    Set oFormula = Nothing
End Sub


Private Sub FillVehicleType(VehicleType() As String)
    '填充车型
    Dim oBase As New BaseInfo
    Dim aszVehicleType() As String
    Dim nVehicleTypeCount As Integer '车型个数
    Dim i As Integer
    Dim nCount As Integer
    Dim j As Integer
    Dim m As Integer
    Dim n As Integer
    Dim bFind As Boolean
    Dim szTemp() As String
On Error GoTo ErrorHandle:
    nVehicleTypeCount = ArrayLength(VehicleType)
    '得到所有的车型
    oBase.Init g_oActiveUser
    aszVehicleType = oBase.GetAllVehicleModel()
    nCount = ArrayLength(aszVehicleType)
    cboParam.Clear
    For i = 1 To nCount
        '填充所有传入的车型
        bFind = False
        cboParam.AddItem MakeDisplayString(RTrim(aszVehicleType(i, 1)), RTrim(aszVehicleType(i, 2)))
        If cboParam.SelCount < nVehicleTypeCount Then
            For j = 1 To 16
                If m_atFormulaParam(j).szFormula = Trim(cboItemFormula.Text) Then
                    '将字符串转换为数组
                    szTemp = StringToTeam(m_atFormulaParam(j).sgParam1)
                    For n = 1 To ArrayLength(szTemp)
                        If szTemp(n) = RTrim(aszVehicleType(i, 1)) Then
                            For m = 1 To nVehicleTypeCount
                                If Trim(aszVehicleType(i, 1)) = Trim(VehicleType(m)) Then
                                    cboParam.Selected(i - 1) = True
                                    bFind = True
                                    Exit For
                                End If
                            Next m
                        End If
                        If bFind = True Then Exit For
                    Next n
                If bFind = True Then Exit For
                End If
            Next j
        End If
    Next i
    Set oBase = Nothing
    Exit Sub
ErrorHandle:
    Set oBase = Nothing
    ShowErrorMsg
End Sub

'填充检票口
Private Sub FillCheckGate(CheckGate() As String)
    Dim oBase As New BaseInfo
    Dim aszCheckGate() As String
    Dim i As Integer, nCount As Integer
    Dim j As Integer
    Dim n, m As Integer
    Dim length1, length2 As Integer
    Dim bFind As Boolean
    Dim nCheckGateCount As Integer
    Dim szCheckGate() As String

On Error GoTo ErrorHandle:

    oBase.Init g_oActiveUser
    aszCheckGate = oBase.GetAllCheckGate()
    nCount = ArrayLength(aszCheckGate)
    nCheckGateCount = ArrayLength(CheckGate)

    cboParam.Clear
    For i = 1 To nCount
        bFind = False
        cboParam.AddItem MakeDisplayString(RTrim(aszCheckGate(i, 1)), RTrim(aszCheckGate(i, 2)))
        For j = 1 To 16
            If m_atFormulaParam(j).szFormula = Trim(cboItemFormula.Text) Then
               szCheckGate = StringToTeam(m_atFormulaParam(j).sgParam1)
               For n = 1 To ArrayLength(szCheckGate)
                   If szCheckGate(n) = RTrim(aszCheckGate(i, 1)) Then
                     For m = 1 To nCheckGateCount
                        If RTrim(aszCheckGate(i, 1)) = Trim(CheckGate(m)) Then
                           cboParam.Selected(i - 1) = True
                           bFind = True
                           Exit For
                        End If
                    Next m
                  End If
                  If bFind = True Then Exit For
              Next n
              If bFind = True Then Exit For
            End If
        Next j
    Next i
    'If bFind = False Then CboParam.Selected(0) = True
    Exit Sub

ErrorHandle:
    ShowErrorMsg
End Sub
