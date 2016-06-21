VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmNetPrint 
   BackColor       =   &H8000000C&
   Caption         =   "打印网上售票"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.PictureBox picParent 
      BackColor       =   &H00FCF0EC&
      Height          =   1935
      Left            =   1440
      ScaleHeight     =   1875
      ScaleWidth      =   8865
      TabIndex        =   0
      Top             =   960
      Width           =   8925
      Begin VB.TextBox txtValiDate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   4920
         TabIndex        =   8
         Top             =   1230
         Width           =   1695
      End
      Begin VB.TextBox txtGetTicketID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   7
         Top             =   1230
         Width           =   2655
      End
      Begin VB.Frame fraOther 
         BackColor       =   &H00FCF0EC&
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6495
         Begin VB.TextBox txtCardID 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3840
            TabIndex        =   3
            Top             =   330
            Width           =   2415
         End
         Begin VB.CheckBox chkOther 
            BackColor       =   &H00FCF0EC&
            Height          =   255
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   255
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   255
            Left            =   1200
            TabIndex        =   4
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   96600065
            CurrentDate     =   39055
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "身份证号:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   6
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "上车日期:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   1215
         End
      End
      Begin RTComctl3.CoolButton cmdRePrint 
         Height          =   465
         Left            =   7080
         TabIndex        =   11
         Top             =   1200
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   820
         BTYPE           =   3
         TX              =   "重打(&R)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         MICON           =   "frmNetPrint.frx":0000
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
         Height          =   465
         Left            =   7080
         TabIndex        =   12
         Top             =   360
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   820
         BTYPE           =   3
         TX              =   "打印(&P)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         MICON           =   "frmNetPrint.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "密码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4200
         TabIndex        =   10
         Top             =   1320
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "取票号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmNetPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkOther_Click()
    If chkOther.Value = 1 Then
        dtpDate.Enabled = True
        txtCardID.Enabled = True
    Else
        dtpDate.Enabled = False
        txtCardID.Enabled = False
    End If
End Sub

Private Sub cmdOK_Click()
    cmdFin_Click
End Sub

Private Sub Form_Activate()
On Error GoTo here
    txtGetTicketID.SetFocus
    ShowSBInfo Me.Caption, ESB_WorkingInfo
    
    Exit Sub
here:
    ShowErrorMsg
    
End Sub
Private Sub cmdRePrint_Click()

    frmReprint.Show vbModal
    frmReprint.ZOrder 0

End Sub

Private Sub cmdFin_Click()
On Error GoTo here
    Dim rs As Recordset
    Dim aszResult() As TSellTicketResult
    Dim atSellParam() As TSellTicketParam
    If Trim(txtGetTicketID.Text) <> "" Or Trim(txtValiDate.Text) <> "" Or Trim(txtCardID.Text) <> "" Then
        If chkOther.Value = 1 Then
            Set rs = m_oNetSell.InterNetValiDate(txtGetTicketID.Text, txtValiDate.Text, txtCardID.Text, dtpDate.Value)
        Else
            Set rs = m_oNetSell.InterNetValiDate(txtGetTicketID.Text, txtValiDate.Text)
        End If
            
        If rs.RecordCount <= 0 Then
            frmNotify.m_szErrorDescription = "无效订票号!请正确输入!"
            frmNotify.Show vbModal
            Unload Me
        ElseIf rs.RecordCount > Val(m_lEndTicketNo) - Val(m_lTicketNo) + 1 Then
            frmNotify.m_szErrorDescription = "机器所剩票少取您所订的票,请与管理员联系!"
            frmNotify.Show vbModal
            Unload Me
        Else

            IncTicketNo -1, True
            aszResult = m_oNetSell.GetInterNetTicket(rs, GetTicketNo, atSellParam)
            PrintTicketEx aszResult, rs, atSellParam
            Unload Me
            

            
            
        End If
    Else
        MsgBox "请填写身份证号、取票号或密码!"
    End If
    Set rs = Nothing
    Exit Sub
here:

    frmNotify.m_szErrorDescription = err.Description
    frmNotify.Show vbModal
'    Unload Me
End Sub
Private Sub Form_Load()

    txtGetTicketID.Text = ""
    txtValiDate.Text = ""
    dtpDate.Value = Date
End Sub


Private Sub Form_Deactivate()
    ShowSBInfo "", ESB_WorkingInfo
End Sub
Private Sub Form_Resize()
    If MDISellTicket.ActiveForm Is Me Then
        If Not Me.WindowState = vbMaximized Then Me.WindowState = vbMaximized
        picParent.Left = (Me.ScaleWidth - picParent.Width) / 2
        picParent.Top = (Me.ScaleHeight - picParent.Height) / 2
    End If
End Sub


Private Sub txtValiDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdFin_Click
    End If
End Sub
Private Sub PrintTicketEx(paszResult() As TSellTicketResult, pszRsTemp As Recordset, atSellParam() As TSellTicketParam)
    Dim i As Integer
    Dim aSellTicket() As TSellTicketParam
    Dim nCount As Integer
    Dim dyBusDate() As Date
    Dim szBusID() As String
    Dim szDesStationID() As String
    Dim szDesStationName() As String
    Dim szSellStationID() As String
    Dim szSellStationName() As String
    Dim szStartStationName As String
    Dim rs As Recordset
    Dim psgDiscount() As Single
    Dim apiTicketInfo() As TPrintTicketParam
    Dim pszBusDate() As String
    Dim pnTicketCount() As Integer
    Dim pszEndStation() As String
    Dim pszOffTime() As String
    Dim pszBusID() As String
    Dim pszVehicleType() As String
    Dim pszCheckGate() As String

    Dim anInsurance() As Integer '售票用
    Dim Netparam() As String '返回网上买票信息
    Dim nCount2 As Integer
    Dim szTemp As String
    Dim aszTerminateName() As String
    Dim bSaleChange() As Boolean
    Dim nBusType() As EBusType
    Dim pbIsTakeChild() As Boolean
    
    Dim aszRealNameInfo() As TCardInfo
    Dim rsTempOther As New Recordset
    Set rsTempOther = pszRsTemp
    rsTempOther.MoveFirst
    
    nCount2 = 1
    ReDim apiTicketInfo(1 To nCount2)
    ReDim pszBusDate(1 To nCount2)
    ReDim pnTicketCount(1 To nCount2)
    ReDim pszEndStation(1 To nCount2)
    ReDim pszOffTime(1 To nCount2)
    ReDim pszBusID(1 To nCount2)
    ReDim pszVehicleType(1 To nCount2)
    ReDim pszCheckGate(1 To nCount2)
    ReDim szSellStationName(1 To nCount2)
    ReDim anInsurance(1 To nCount2)
    ReDim Netparam(1 To nCount2, 1 To 3)
    
    ReDim aszTerminateName(1 To nCount2)
    ReDim bSaleChange(1 To nCount2)
    ReDim nBusType(1 To nCount2)
    ReDim pbIsTakeChild(1 To pszRsTemp.RecordCount)
    
    ReDim aszRealNameInfo(1 To ArrayLength(paszResult(1).aszSeatNo))
    
            pszRsTemp.MoveFirst
            For nCount = 1 To nCount2
                ReDim apiTicketInfo(nCount2).aptPrintTicketInfo(1 To pszRsTemp.RecordCount)
                Set rs = m_oNetSell.GetBusExRs(Trim(pszRsTemp!bus_date), Trim(pszRsTemp!bus_id), Trim(pszRsTemp!sell_station_id), Trim(pszRsTemp!des_station_id))
                pnTicketCount(nCount) = pszRsTemp.RecordCount
                pszEndStation(nCount) = Trim(rs!station_name)
                pszVehicleType(nCount) = Trim(rs!vehicle_type_name)
                pszCheckGate(nCount) = GetCheckName(Trim(rs!sell_check_gate_id))
                pszBusDate(nCount) = CDate(Trim(rs!bus_date))
                pszOffTime(nCount) = Format(rs!sell_bus_start_time, "hh:mm")
                szSellStationName(nCount) = Trim(rs!sell_station_name)
                pszBusID(nCount) = Trim(rs!bus_id)
                anInsurance(nCount) = 0
                Netparam(nCount, 1) = Trim(pszRsTemp!bank_card_id)
                Netparam(nCount, 2) = Trim(pszRsTemp!center_scroll_id)
                Netparam(nCount, 3) = Trim(pszRsTemp!pay_count)
                For i = 1 To ArrayLength(paszResult(1).aszSeatNo)
                    Dim szTmp As String
'                    szTmp = GetTicketNo + i
                    apiTicketInfo(1).aptPrintTicketInfo(i).nTicketType = paszResult(nCount2).aszTicketType(i) 'Trim(pszRsTemp!ticket_type)
                    apiTicketInfo(1).aptPrintTicketInfo(i).sgTicketPrice = FormatDbValue(rsTempOther!pay_count) 'paszResult(nCount2).asgTicketPrice(i)
                    apiTicketInfo(1).aptPrintTicketInfo(i).szSeatNo = paszResult(nCount2).aszSeatNo(i)
                    apiTicketInfo(1).aptPrintTicketInfo(i).szTicketNo = atSellParam(nCount2).BuyTicketInfo(i).szTicketNo
                    
                    aszRealNameInfo(i).szIDCardNo = Trim(FormatDbValue(rsTempOther!card_id))
                    aszRealNameInfo(i).szPersonName = Trim(FormatDbValue(rsTempOther!passenger))
                    rsTempOther.MoveNext
                    
'                    pbIsTakeChild(i) = atSellParam(nCount2).BuyTicketInfo(i).szInvoiceID
'                    IncTicketNo 1, True
                Next
                aszTerminateName(nCount) = ""
                bSaleChange(nCount) = False
                nBusType(nCount) = TP_NormalBus
                
'                pszRsTemp.MoveNext
            Next nCount

'                lTkNum = lTkNum - ArrayLength(paszResult(nCount2).aszSeatNo)
                PrintNetTicket apiTicketInfo, pszBusDate, pnTicketCount, pszEndStation, pszOffTime, pszBusID, pszVehicleType, pszCheckGate, bSaleChange, aszTerminateName, szSellStationName, anInsurance, nBusType, pbIsTakeChild, aszRealNameInfo
            IncTicketNo pszRsTemp.RecordCount + 1
            frmNotify.m_szErrorDescription = "   购票成功,    请拿好车票    祝旅途愉快! 欢迎再次光临！"
            frmNotify.Show vbModal
    Exit Sub
here:
    frmNotify.m_szErrorDescription = err.Description
    frmNotify.Show vbModal
End Sub

