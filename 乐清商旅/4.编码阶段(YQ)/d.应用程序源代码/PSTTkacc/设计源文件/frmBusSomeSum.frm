VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmBusSomeSum 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "车次售票统计简报"
   ClientHeight    =   4395
   ClientLeft      =   4350
   ClientTop       =   3645
   ClientWidth     =   6555
   Icon            =   "frmBusSomeSum.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtSellStationID 
      Height          =   300
      Left            =   1305
      TabIndex        =   23
      Top             =   2970
      Width           =   1635
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   345
      Left            =   2250
      TabIndex        =   22
      Top             =   3900
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "帮助"
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
      MICON           =   "frmBusSomeSum.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cboSellStation 
      Height          =   300
      Left            =   1305
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   2520
      Width           =   1650
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   345
      Left            =   3660
      TabIndex        =   14
      Top             =   3915
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
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
      MICON           =   "frmBusSomeSum.frx":0028
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
      Height          =   345
      Left            =   5100
      TabIndex        =   15
      Top             =   3900
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmBusSomeSum.frx":0044
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
      Enabled         =   0   'False
      Height          =   3120
      HelpContextID   =   6000040
      Left            =   -120
      TabIndex        =   19
      Top             =   3660
      Width           =   8745
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   6555
      TabIndex        =   17
      Top             =   -60
      Width           =   6555
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   18
         Top             =   270
         Width           =   990
      End
   End
   Begin VB.TextBox txtLike 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4440
      TabIndex        =   13
      Top             =   2520
      Width           =   1605
   End
   Begin VB.ComboBox cboExtraStatus 
      Height          =   300
      Left            =   1305
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2025
      Width           =   1635
   End
   Begin VB.OptionButton optCompany 
      BackColor       =   &H00E0E0E0&
      Caption         =   "按拆帐公司汇总"
      Height          =   285
      Left            =   3150
      TabIndex        =   7
      Top             =   1470
      Width           =   1605
   End
   Begin VB.OptionButton optCombine 
      BackColor       =   &H00E0E0E0&
      Caption         =   "按车次段汇总"
      Height          =   255
      Left            =   3150
      TabIndex        =   6
      Top             =   930
      Width           =   1455
   End
   Begin VB.ComboBox cboBusSection 
      Height          =   300
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2010
      Width           =   1905
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -30
      TabIndex        =   16
      Top             =   660
      Width           =   6885
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   300
      Left            =   1305
      TabIndex        =   3
      Top             =   1455
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   300
      Left            =   1305
      TabIndex        =   1
      Top             =   900
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin FText.asFlatTextBox txtTransportCompanyID 
      Height          =   300
      Left            =   4440
      TabIndex        =   11
      Top             =   2010
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      ButtonVisible   =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&T):"
      Height          =   180
      Left            =   240
      TabIndex        =   21
      Top             =   2580
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "模糊车次(&A):"
      Height          =   180
      Left            =   3150
      TabIndex        =   12
      Top             =   2580
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   1515
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "补票状态(&S):"
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   2055
      Width           =   1080
   End
   Begin VB.Label lblCombine 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次段序号(&R):"
      Height          =   180
      Left            =   3150
      TabIndex        =   8
      Top             =   2055
      Width           =   1260
   End
   Begin VB.Label lblCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参运公司(T):"
      Height          =   180
      Left            =   3150
      TabIndex        =   10
      Top             =   2070
      Width           =   1080
   End
End
Attribute VB_Name = "frmBusSomeSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IConditionForm
Const cszFileName = "车次人数及拆入小计模板.xls"
Const cnColNumber = 5 '列数
Public m_bOk As Boolean
Public m_bBySaleTime As Boolean
Private m_rsData As Recordset
Private m_vaCustomData As Variant

Dim m_aszTemp() As String
Dim oDss As New TicketBusDim
Dim m_szCode As String

Dim oDssCompany As New TicketCompanyDim

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error GoTo Error_Handle
    '生成记录集
    Dim rsTemp As Recordset
    Dim nTemp As Integer
    Dim rsData As New Recordset
    Dim i As Integer
    Dim j As Integer
    If m_bBySaleTime Then
        If optCompany.Value Then
            Set rsTemp = oDss.GetBusStatBySaleTime(m_szCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), txtLike.Text, IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
        Else
            Set rsTemp = oDss.GetCombineBusSimply(dtpBeginDate.Value, dtpEndDate.Value, Val(cboBusSection.Text), ResolveDisplay(cboExtraStatus.Text), txtLike.Text, IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
        End If
    Else
        If optCompany.Value Then
            Set rsTemp = oDss.GetBusStatByBusDate(m_szCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), txtLike.Text, IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
        Else
            Set rsTemp = oDss.GetCombineBusSimplyByBusDate(dtpBeginDate.Value, dtpEndDate.Value, Val(cboBusSection.Text), ResolveDisplay(cboExtraStatus.Text), txtLike.Text, IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
        End If
    End If
    
    '新建记录集
    With rsData.Fields
        For j = 1 To cnColNumber
            .Append "bus_id_" & j, adChar, 5
            .Append "passenger_number_" & j, adInteger
            .Append "some_sum_" & j, adCurrency
        Next j
           
    End With
    '赋记录集
    
    rsData.Open
    
    For i = 1 To rsTemp.RecordCount
        If i Mod cnColNumber = 1 Then
            '新增一条
            rsData.AddNew
            rsData!bus_id_1 = rsTemp!bus_id
            rsData!passenger_number_1 = rsTemp!passenger_number
            rsData!some_sum_1 = rsTemp!some_sum
            
            For j = 2 To cnColNumber
                rsData("bus_id_" & j) = ""
                rsData("passenger_number_" & j) = 0
                rsData("some_sum_" & j) = 0
            Next j
        Else
            nTemp = (i Mod cnColNumber)
            If nTemp = 0 Then nTemp = cnColNumber
            rsData("bus_id_" & nTemp) = rsTemp!bus_id
            rsData("passenger_number_" & nTemp) = rsTemp!passenger_number
            rsData("some_sum_" & nTemp) = rsTemp!some_sum
        End If
        rsTemp.MoveNext
    Next i
    
    
    
    Set m_rsData = rsData
    
     
    ReDim m_vaCustomData(1 To 6, 1 To 2)
    
    m_vaCustomData(1, 1) = "统计开始日期"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY年MM月DD日")
    m_vaCustomData(2, 1) = "统计结束日期"
    m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")
    m_vaCustomData(3, 1) = "补票状态"
    m_vaCustomData(3, 2) = cboExtraStatus.Text
    m_vaCustomData(4, 1) = "统计方式"
    m_vaCustomData(4, 2) = IIf(m_bBySaleTime, cszByOperationTime, cszByBusDate)
    
    m_vaCustomData(5, 1) = "制表人"
    m_vaCustomData(5, 2) = m_oActiveUser.UserID
    Dim szSellStation As String
    ResolveDisplay cboSellStation, szSellStation
    m_vaCustomData(6, 1) = "上车站"
    m_vaCustomData(6, 2) = szSellStation
    
    m_bOk = True
    Unload Me
    Exit Sub
Error_Handle:
    ShowErrorMsg
End Sub


Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_Load()
    oDss.Init m_oActiveUser
    oDssCompany.Init m_oActiveUser
    m_szCode = ""
    m_bOk = False
    FillCombine
'    dtpBeginDate.Value = DateAdd("d", -1, m_oParam.NowDate)
'    dtpEndDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    '设置为上个月的一号到31
    Dim dyNow As Date
    dyNow = m_oParam.NowDate
    dtpBeginDate.Value = Format(DateAdd("m", -1, dyNow), "yyyy-mm-01")
    dtpEndDate.Value = DateAdd("d", -1, Format(dyNow, "yyyy-mm-01"))
    cboExtraStatus.AddItem "1[售票]"
    cboExtraStatus.AddItem "2[补票]"
    cboExtraStatus.AddItem "3[售票+补票]"
    
    cboExtraStatus.ListIndex = 2
    
    optCompany.Value = True
    SetVisible False
    FillSellStation cboSellStation
    
    If Trim(m_oActiveUser.SellStationID) <> "" Then
        txtSellStationID.Enabled = False
    End If

    
    If m_bBySaleTime Then
        Me.Caption = "车次售票统计简报[按售票时间汇总]"
        lblCaption = "请输入售票的起止日期:"
    Else
        Me.Caption = "车次售票统计简报[按车次日期汇总]"
        lblCaption = "请输入车次的起止日期:"
    End If
    
End Sub



Private Property Get IConditionForm_CustomData() As Variant
    IConditionForm_CustomData = m_vaCustomData
End Property

Private Property Get IConditionForm_FileName() As String
    IConditionForm_FileName = cszFileName
End Property

Private Property Get IConditionForm_RecordsetData() As Recordset
    Set IConditionForm_RecordsetData = m_rsData
End Property



Private Sub optCombine_Click()
    SetVisible True
End Sub

Private Sub optCompany_Click()
    SetVisible False
    
End Sub

Private Sub txtTransportCompanyID_ButtonClick()
    Dim aszTransportCompanyID() As String
    aszTransportCompanyID = m_oShell.SelectCompany
    txtTransportCompanyID.Text = TeamToString(aszTransportCompanyID, 2)
    
    m_szCode = TeamToString(aszTransportCompanyID, 1)
    
End Sub

Private Sub FillCombine()
    '填充唯一的车次序号
    Dim aszTemp() As String
    Dim i As Integer
    Dim nCount As Integer
    
    aszTemp = oDssCompany.GetUniqueCombine
    nCount = ArrayLength(aszTemp)
    For i = 1 To nCount
        cboBusSection.AddItem aszTemp(i)
    Next i
    If cboBusSection.ListCount > 0 Then cboBusSection.ListIndex = 0
    
End Sub


Private Sub SetVisible(pbVisible As Boolean)
    lblCombine.Visible = pbVisible
    cboBusSection.Visible = pbVisible
    lblCompany.Visible = Not pbVisible
    txtTransportCompanyID.Visible = Not pbVisible
    
End Sub




'Private Sub FillSellStation()
'    '判断用户属于哪个上车站,如果为空则填充一个空行,再填充所有的上车站
'
'    '否则只填充用户所属的上车站
'End Sub

