VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmSellerEveryMonth 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "售票员结算简报"
   ClientHeight    =   3750
   ClientLeft      =   2415
   ClientTop       =   2385
   ClientWidth     =   6765
   Icon            =   "frmSellerEveryMonth.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ListBox lstSinceSaler 
      Height          =   1320
      Left            =   1290
      MultiSelect     =   2  'Extended
      TabIndex        =   11
      Top             =   2340
      Width           =   3255
   End
   Begin VB.ListBox lstPreSaler 
      Height          =   1320
      Left            =   1290
      MultiSelect     =   2  'Extended
      TabIndex        =   10
      Top             =   930
      Width           =   3255
   End
   Begin RTComctl3.CoolButton cmdSinceSaler 
      Height          =   405
      Left            =   4830
      TabIndex        =   9
      Top             =   1560
      Width           =   1755
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "添加结束售票员(&R)"
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
      MICON           =   "frmSellerEveryMonth.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdPreSaler 
      Height          =   405
      Left            =   4830
      TabIndex        =   8
      Top             =   1080
      Width           =   1755
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "添加起始售票员(&S)"
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
      MICON           =   "frmSellerEveryMonth.frx":0028
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
      Height          =   405
      Left            =   4830
      TabIndex        =   1
      Top             =   600
      Width           =   1755
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmSellerEveryMonth.frx":0044
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
      Height          =   405
      Left            =   4830
      TabIndex        =   0
      Top             =   120
      Width           =   1755
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "确定(&O)"
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
      MICON           =   "frmSellerEveryMonth.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   315
      Left            =   1290
      TabIndex        =   2
      Top             =   90
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   67305475
      UpDown          =   -1  'True
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   315
      Left            =   1290
      TabIndex        =   3
      Top             =   480
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   67305475
      UpDown          =   -1  'True
      CurrentDate     =   36572
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "结束售票员:"
      Height          =   180
      Left            =   90
      TabIndex        =   7
      Top             =   2790
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "起始售票员:"
      Height          =   180
      Left            =   90
      TabIndex        =   6
      Top             =   1080
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   90
      TabIndex        =   5
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   90
      TabIndex        =   4
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmSellerEveryMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IConditionForm

Const cszFileName = "售票员每月结算模板.xls"
Public m_bOk As Boolean
'Private m_szPreSaler As String
'Private m_szSinceSaler As String


Private m_rsData As Recordset
Private m_vaCustomData As Variant


Private m_vaPreSaler As Variant
Private m_vaSinceSaler As Variant

Private m_rsPreSaler As Recordset '前面的售票员售票的记录集
Private m_rsSinceSaler As Recordset '后面的售票员售票的记录集
Private m_rsMidSaler As Recordset '中间的售票员售票的记录集
Private m_auiUserInfo() As TUserInfo
'Private m_rsTemp As New Recordset


'Dim oSellFinance As New TicketSellerDim
Dim oSysMan As New SystemMan
Dim oSellDim As New TicketSellerDim

Private Sub cmdCancel_Click()
    m_bOk = False
    Unload Me
End Sub

Private Sub cmdok_Click()
    Dim i As Integer
    Dim aszUserInfo() As String
    Dim dyStartDate As Date
    Dim dyEndDate As Date
    
    On Error GoTo ErrorHandle
    SetMouseBusy True
    
    oSysMan.Init m_oActiveUser
    m_auiUserInfo = oSysMan.GetAllUser()
    ReDim aszUserInfo(1 To ArrayLength(m_auiUserInfo))
    For i = 1 To ArrayLength(m_auiUserInfo)
        aszUserInfo(i) = m_auiUserInfo(i).UserID
    Next
    
    ReDim m_vaPreSaler(1 To lstPreSaler.ListCount)
    For i = 0 To lstPreSaler.ListCount - 1
        m_vaPreSaler(i + 1) = lstPreSaler.List(i)
    Next i
    
    ReDim m_vaSinceSaler(1 To lstSinceSaler.ListCount)
    For i = 0 To lstSinceSaler.ListCount - 1
        m_vaSinceSaler(i + 1) = lstSinceSaler.List(i)
    Next i
'    oSellFinance.Init m_oActiveUser
    oSellDim.Init m_oActiveUser
    '初始化记录集
    InitRecordset
    '得到起始售票员的明细记录
    If dtpBeginDate.Value > CDate(Format(dtpBeginDate.Value, "yyyy-mm-dd")) Then
    
        dyStartDate = Format(DateAdd("d", 1, dtpBeginDate.Value), "yyyy-mm-dd 00:00:00")
        Set m_rsPreSaler = GetPreRs(m_vaPreSaler, dtpBeginDate.Value, dyStartDate, True)
    Else
        dyStartDate = dtpBeginDate.Value
    End If
    '得到结束售票员的明细记录
    If dtpEndDate.Value > CDate(Format(dtpEndDate.Value, "yyyy-mm-dd")) Then
        dyEndDate = Format(dtpEndDate.Value, "yyyy-mm-dd 00:00:00")
        Set m_rsSinceSaler = GetPreRs(m_vaSinceSaler, dyEndDate, dtpEndDate.Value, False)
    Else
        dyEndDate = dtpEndDate.Value
    End If
    '得到中间段内的售票员的统计记录
    
    Set m_rsMidSaler = oSellDim.SellerDateStat(aszUserInfo, dyStartDate, DateAdd("d", -1, dyEndDate))
    
    '记录集合并
    Set m_rsData = m_rsMidSaler
    MergeRecordset m_rsData, m_rsPreSaler
    MergeRecordset m_rsData, m_rsSinceSaler
    m_bOk = True
    Unload Me
    SetMouseBusy False
    ReDim m_vaCustomData(1 To 3, 1 To 2)
    m_vaCustomData(1, 1) = "统计开始日期"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY年MM月DD日")
    
    m_vaCustomData(2, 1) = "统计结束日期"
    m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")
    
    
    m_vaCustomData(3, 1) = "制表人"
    m_vaCustomData(3, 2) = m_oActiveUser.UserID
    
    
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg
    SetMouseBusy False
    

End Sub


Private Sub InitRecordset()
    '初始化记录集
    Dim j As Integer
    Set m_rsData = New Recordset
    With m_rsData.Fields
        .Append "user_id", adChar, 40
        
        '添加各种票种的数量与金额
        For j = 1 To TP_TicketTypeCount
            .Append "number_ticket_type" & j, adInteger
            .Append "amount_ticket_type" & j, adCurrency
        Next j
        
        .Append "cancel_number", adInteger
        .Append "cancel_amount", adCurrency
        
        .Append "return_number", adInteger
        .Append "return_amount", adCurrency
        .Append "return_charge", adCurrency
        
        .Append "change_number", adInteger
        .Append "change_amount", adCurrency
        .Append "change_charge", adCurrency
        
        .Append "total_number", adInteger
        .Append "total_amount", adCurrency
    End With
    m_rsData.Open
End Sub


Private Function GetPreRs(paszUser As Variant, pdyStartDate As Date, pdyEndDate As Date, pbStart As Boolean) As Recordset
    '得到起始售票员的明细记录
    Dim rsSellDetail As Recordset
    Dim adbOther() As Double
    Dim i As Integer, nUserCount As Integer
    'Dim rsData As Recordset, vaCostumData As Variant
    Dim j As Integer
    Dim rsTemp As New Recordset
    Dim szUser As String
    Dim bNeedAdd As Boolean
    
    '初始化记录集
    With rsTemp.Fields
        .Append "user_id", adChar, 40
        For j = 1 To TP_TicketTypeCount
            .Append "number_ticket_type" & j, adInteger
            .Append "amount_ticket_type" & j, adCurrency
        Next j
        .Append "cancel_number", adInteger
        .Append "cancel_amount", adCurrency
        .Append "return_number", adInteger
        .Append "return_amount", adCurrency
        .Append "return_charge", adCurrency
        .Append "change_number", adInteger
        .Append "change_amount", adCurrency
        .Append "change_charge", adCurrency
        .Append "total_number", adInteger
        .Append "total_amount", adCurrency
    End With
    rsTemp.Open
    
    nUserCount = ArrayLength(paszUser)
    
    If nUserCount > 0 Then
        WriteProcessBar True, , nUserCount, "正在形成售票员记录集..."
        For i = 1 To nUserCount
            '查找售票员是否在记录集中已存在
            szUser = Trim(ResolveDisplay(paszUser(i)))
            bNeedAdd = False
            If Not (rsTemp Is Nothing) Then
                If rsTemp.RecordCount > 0 Then
                    rsTemp.MoveFirst
                    For j = 1 To rsTemp.RecordCount
                        If szUser = rsTemp!user_id Then
                            Exit For
                        End If
                        rsTemp.MoveNext
                    Next j
                    If j > rsTemp.RecordCount Then
                        bNeedAdd = True
                    End If
                Else
                    bNeedAdd = True
                End If
            Else
                bNeedAdd = True
            End If
            If bNeedAdd Then
                '该用户不存在,则新增一条记录
                    rsTemp.AddNew
                    rsTemp!user_id = Trim(szUser)
                    For j = 1 To TP_TicketTypeCount
                        rsTemp("number_ticket_type" & j) = 0
                        rsTemp("amount_ticket_type" & j) = 0
                    Next j
                    rsTemp!cancel_number = 0
                    rsTemp!cancel_amount = 0
                    rsTemp!return_number = 0
                    rsTemp!return_amount = 0
                    rsTemp!return_charge = 0
                    rsTemp!change_number = 0
                    rsTemp!change_amount = 0
                    rsTemp!change_charge = 0
                    rsTemp!total_number = 0
                    rsTemp!total_amount = 0
            End If
            If pbStart Then
                Set rsSellDetail = oSellDim.SellerEveryDaySellDetail(ResolveDisplay(paszUser(i)), GetCombineTime(paszUser(i)), pdyEndDate)
            Else
                
                Set rsSellDetail = oSellDim.SellerEveryDaySellDetail(ResolveDisplay(paszUser(i)), pdyStartDate, GetCombineTime(paszUser(i)))
            End If
            Do While Not rsSellDetail.EOF
                rsTemp("number_ticket_type" & rsSellDetail!ticket_type) = rsTemp("number_ticket_type" & rsSellDetail!ticket_type) + 1
                If rsSellDetail!ticket_type <> TP_FreeTicket Then
                    rsTemp("amount_ticket_type" & rsSellDetail!ticket_type) = rsTemp("amount_ticket_type" & rsSellDetail!ticket_type) + rsSellDetail!ticket_price
                End If
                rsSellDetail.MoveNext
            Loop
            If pbStart Then
                adbOther = oSellDim.SellerEveryDayAnotherThing(ResolveDisplay(paszUser(i)), GetCombineTime(paszUser(i)), pdyEndDate)
            Else
                adbOther = oSellDim.SellerEveryDayAnotherThing(ResolveDisplay(paszUser(i)), pdyStartDate, GetCombineTime(paszUser(i)))
                
            
            
            End If
            rsTemp!cancel_number = adbOther(1, 1)
            rsTemp!cancel_amount = adbOther(1, 2)
            rsTemp!return_number = adbOther(2, 1)
            rsTemp!return_amount = adbOther(2, 2)
            rsTemp!return_charge = adbOther(2, 3)
            rsTemp!change_number = adbOther(3, 1)
            rsTemp!change_amount = adbOther(3, 2)
            rsTemp!change_charge = adbOther(3, 3)
            rsTemp!total_number = rsTemp("number_ticket_type1") + rsTemp("number_ticket_type2") + rsTemp("number_ticket_type3") + rsTemp("number_ticket_type4") + rsTemp("number_ticket_type5") + rsTemp("number_ticket_type6") - adbOther(2, 1)
            rsTemp!total_amount = rsTemp("amount_ticket_type1") + rsTemp("amount_ticket_type2") + rsTemp("amount_ticket_type3") + rsTemp("amount_ticket_type4") + rsTemp("amount_ticket_type5") + rsTemp("amount_ticket_type6") - adbOther(1, 2) - adbOther(2, 2) + adbOther(2, 3) - adbOther(3, 2) + adbOther(3, 3)
            WriteProcessBar , i, nUserCount
            
        Next i
    End If
    WriteProcessBar False, , ""
    Set GetPreRs = rsTemp
End Function

Private Sub MergeRecordset(prsFirst As Recordset, prsSecond As Recordset)
    Dim bNeedAdd As Boolean
    Dim i As Integer
    Dim j As Integer
    '合并记录
    If (prsFirst Is Nothing) And (prsSecond Is Nothing) Then
        Exit Sub
    ElseIf prsFirst Is Nothing Then
        Set prsFirst = prsSecond
        Exit Sub
    ElseIf prsSecond Is Nothing Then
        Exit Sub
    End If
    prsSecond.MoveFirst
    For i = 1 To prsSecond.RecordCount
    
        bNeedAdd = False
        If prsFirst.RecordCount > 0 Then prsFirst.MoveFirst
        For j = 1 To prsFirst.RecordCount
            If Trim(prsSecond!user_id) = Trim(ResolveDisplay(prsFirst!user_id)) Then
                Exit For
            End If
            prsFirst.MoveNext
        Next j
        If j > prsFirst.RecordCount Then
            '该用户不存在,则新增一条记录
            prsFirst.AddNew
            prsFirst!user_id = prsSecond!user_id
            For j = 1 To TP_TicketTypeCount
                prsFirst("number_ticket_type" & j) = 0
                prsFirst("amount_ticket_type" & j) = 0
            Next j
            prsFirst!cancel_number = 0
            prsFirst!cancel_amount = 0
            prsFirst!return_number = 0
            prsFirst!return_amount = 0
            prsFirst!return_charge = 0
            prsFirst!change_number = 0
            prsFirst!change_amount = 0
            prsFirst!change_charge = 0
            prsFirst!total_number = 0
            prsFirst!total_amount = 0
        Else
            'prsFirst!user_id = prsSecond!user_id
        End If
        
        For j = 1 To TP_TicketTypeCount
            prsFirst("number_ticket_type" & j) = prsFirst("number_ticket_type" & j) + prsSecond("number_ticket_type" & j)
            prsFirst("amount_ticket_type" & j) = prsFirst("amount_ticket_type" & j) + prsSecond("amount_ticket_type" & j)
        Next j
        prsFirst!cancel_number = prsFirst!cancel_number + prsSecond!cancel_number
        prsFirst!cancel_amount = prsFirst!cancel_amount + prsSecond!cancel_amount
        prsFirst!return_number = prsFirst!return_number + prsSecond!return_number
        prsFirst!return_amount = prsFirst!return_amount + prsSecond!return_amount
        prsFirst!return_charge = prsFirst!return_charge + prsSecond!return_charge
        prsFirst!change_number = prsFirst!change_number + prsSecond!change_number
        prsFirst!change_amount = prsFirst!change_amount + prsSecond!change_amount
        prsFirst!change_charge = prsFirst!change_charge + prsSecond!change_charge
        prsFirst!total_number = prsFirst!total_number + prsSecond!total_number
        prsFirst!total_amount = prsFirst!total_amount + prsSecond!total_amount
        
        prsSecond.MoveNext
    Next i
    
End Sub



Private Sub cmdPreSaler_Click()
    AddPreSaler
End Sub



Private Sub cmdSinceSaler_Click()
    AddSinceSaler
End Sub


Private Sub Form_Load()
    Dim dyNow As Date
    dyNow = m_oParam.NowDate
    dtpBeginDate.Value = Format(DateAdd("d", -1, Format(DateAdd("m", -1, dyNow), "yyyy-mm-01")), "YYYY-mm-dd") & " 11:30"
    dtpEndDate.Value = Format(Format(dyNow, "yyyy-mm-01"), "YYYY-mm-dd") & " 11:30"
End Sub

Private Sub EnableOK()
    If lstPreSaler.ListCount > 0 And lstSinceSaler.ListCount > 0 Then
        cmdOk.Enabled = True
    Else
        cmdOk.Enabled = False
    End If
End Sub



Private Function VariantTOString(paszString As Variant)

    Dim i As Integer
    Dim szTemp As String
    Dim nCount As Integer
    nCount = ArrayLength(paszString)
    For i = 1 To nCount - 1
        szTemp = szTemp & "'" & Trim(paszString(i)) & "',"
    Next i
    If nCount > 0 Then
        szTemp = szTemp & "'" & Trim(paszString(i)) & "'"
    End If
    
    VariantTOString = szTemp
    
End Function

Private Property Get IConditionForm_CustomData() As Variant
    IConditionForm_CustomData = m_vaCustomData
End Property

Private Property Get IConditionForm_FileName() As String
    IConditionForm_FileName = cszFileName
End Property

Private Property Get IConditionForm_RecordsetData() As Recordset
    Set IConditionForm_RecordsetData = m_rsData
End Property

Private Sub lstPreSaler_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu MDIMain.pmnu_SelectSaler
        
    End If
End Sub

Private Sub lstSinceSaler_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu MDIMain.pmnu_SelectSaler2
        
    End If
End Sub

Public Sub AddSinceSaler()
    Dim nCount As Integer
    Dim i As Integer
    
    frmSelectSaler.m_bOk = False
    frmSelectSaler.m_dyTime = dtpEndDate.Value
    frmSelectSaler.Show vbModal
    If frmSelectSaler.m_bOk Then
        '更新文本框
        m_vaSinceSaler = frmSelectSaler.m_vaSeller
        
        nCount = ArrayLength(m_vaSinceSaler)
        For i = 1 To nCount
            lstSinceSaler.AddItem m_vaSinceSaler(i) & " " & FormatDateTime(frmSelectSaler.m_dyTime)
            
        Next i
        'txtPreSaler.Text = VariantTOString(frmSelectSaler.m_vaSeller)
        
        EnableOK
    End If
End Sub



Public Sub AddPreSaler()
    Dim nCount As Integer
    Dim i As Integer
    
    frmSelectSaler.m_bOk = False
    frmSelectSaler.m_dyTime = dtpBeginDate.Value
    frmSelectSaler.Show vbModal
    If frmSelectSaler.m_bOk Then
        '更新文本框
        m_vaPreSaler = frmSelectSaler.m_vaSeller
        
        nCount = ArrayLength(m_vaPreSaler)
        For i = 1 To nCount
            lstPreSaler.AddItem m_vaPreSaler(i) & " " & FormatDateTime(frmSelectSaler.m_dyTime)
            
        Next i
        'txtPreSaler.Text = VariantTOString(frmSelectSaler.m_vaSeller)
        
        EnableOK
    End If
End Sub

Public Sub RemovePreSaler()
    Dim i As Integer
    For i = lstPreSaler.ListCount - 1 To 0 Step -1
        If lstPreSaler.Selected(i) Then
            lstPreSaler.RemoveItem i
        End If
    Next i
End Sub


Public Sub RemoveSinceSaler()
    Dim i As Integer
    For i = lstSinceSaler.ListCount - 1 To 0 Step -1
        If lstSinceSaler.Selected(i) Then
            lstSinceSaler.RemoveItem i
        End If
    Next i
End Sub


Public Sub RemovePreAll()
    lstPreSaler.Clear
    
End Sub

Public Sub RemoveSinceAll()
    lstSinceSaler.Clear
End Sub




Public Function GetCombineTime(pszTemp As Variant) As Date
    '得到选择用户的时间
    Dim i As Integer
    i = InStr(1, pszTemp, "]")
    GetCombineTime = CDate(Mid(pszTemp, i + 1, Len(pszTemp) - i))
End Function
