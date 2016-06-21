VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmReprint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "重打网上票"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6675
   Icon            =   "frmReprint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmReprint.frx":000C
   ScaleHeight     =   3165
   ScaleWidth      =   6675
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtNewTicketNo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2220
      MaxLength       =   10
      TabIndex        =   12
      Top             =   720
      Width           =   1950
   End
   Begin VB.Frame fraOther 
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   1200
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
         TabIndex        =   8
         Top             =   330
         Width           =   2415
      End
      Begin VB.CheckBox chkOther 
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   255
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   255
         Left            =   1200
         TabIndex        =   9
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
   End
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
      Left            =   4800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2400
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
      Left            =   1200
      TabIndex        =   2
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox txtTicketNo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2220
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin RTComctl3.CoolButton cmdCancelTicket 
      Height          =   435
      Left            =   4680
      TabIndex        =   14
      Top             =   120
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   767
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
      MICON           =   "frmReprint.frx":9BEAE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdExit 
      Height          =   435
      Left            =   4680
      TabIndex        =   15
      Top             =   720
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "退出(&E)"
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
      MICON           =   "frmReprint.frx":9BECA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblNewTicketNo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "新票号(&E):"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   13
      Top             =   780
      Width           =   1200
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
      Left            =   4080
      TabIndex        =   5
      Top             =   2490
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
      Left            =   0
      TabIndex        =   4
      Top             =   2490
      Width           =   945
   End
   Begin VB.Label lblOldTktNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "需重打票号(&Z):"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   1680
   End
End
Attribute VB_Name = "frmReprint"
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

Private Sub cmdCancelTicket_Click()

    Dim oCTicket As Object
    Dim rs As Recordset
    Dim aszResult() As TSellTicketResult
    Dim paszCancelTKID(1 To 1) As String

On Error GoTo here
    paszCancelTKID(1) = Trim(txtTicketNo)
'    Set oCTicket = m_oNetSell.GetTicketClient(paszCancelTKID(1))
    If MsgBox("是否要重新打印票号为" & txtTicketNo & "的网上卖票?", vbYesNo) = vbYes Then
        If Trim(txtGetTicketID.Text) <> "" Or Trim(txtValiDate.Text) <> "" Or Trim(txtCardID.Text) <> "" Then
            
            If Val(m_lEndTicketNo) - Val(m_lTicketNo) + 1 < 1 Then
                frmNotify.m_szErrorDescription = "机器所剩票少取您所订的票,请与管理员联系!"
                frmNotify.Show vbModal
                Unload Me
            Else
                m_lTicketNo = txtNewTicketNo
                If chkOther.Value = 1 Then
                    Set rs = m_oNetSell.NetTK(txtGetTicketID.Text, txtValiDate.Text, paszCancelTKID, txtCardID.Text, dtpDate.Value)
                Else
                    Set rs = m_oNetSell.NetTK(txtGetTicketID.Text, txtValiDate.Text, paszCancelTKID)
                End If
                If rs.RecordCount > 0 Then
                    aszResult = m_oNetSell.RePrintNetTK(paszCancelTKID, FormatDbValue(rs!getticket_id), FormatDbValue(rs!validate_id), txtNewTicketNo)
                    PrintTicketEx aszResult, rs
                    Unload Me
                Else
                    frmNotify.m_szErrorDescription = "无效订票号!请重新输入!"
                    frmNotify.Show vbModal
                    Unload Me
                End If
                
            End If
        Else
            MsgBox "请填写身份证号、取票号或密码!"
        End If
    End If
    Set rs = Nothing
    Set oCTicket = Nothing
    Exit Sub
here:
    Set oCTicket = Nothing
    ShowErrorMsg
End Sub

Private Sub PrintTicketEx(paszResult() As TSellTicketResult, pszRsTemp As Recordset)
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
    
    ReDim aszRealNameInfo(1 To nCount2)
    
            pszRsTemp.MoveFirst
            For nCount = 1 To nCount2
                ReDim apiTicketInfo(nCount2).aptPrintTicketInfo(1 To pszRsTemp.RecordCount)
                Set rs = m_oNetSell.GetBusExRs(Trim(pszRsTemp!bus_date), Trim(pszRsTemp!bus_id), Trim(pszRsTemp!sell_station_id), Trim(pszRsTemp!des_station_id))
                pnTicketCount(nCount) = pszRsTemp.RecordCount
                pszEndStation(nCount) = Trim(rs!station_name)
                pszVehicleType(nCount) = Trim(rs!vehicle_type_name)
                pszCheckGate(nCount) = GetCheckName(Trim(rs!sell_check_gate_id))
                pszBusDate(nCount) = CDate(Trim(rs!bus_date))
                pszOffTime(nCount) = Format(rs!bus_start_time, "hh:mm")
                szSellStationName(nCount) = Trim(rs!sell_station_name)
                pszBusID(nCount) = Trim(rs!bus_id)
                anInsurance(nCount) = 0
                Netparam(nCount, 1) = Trim(pszRsTemp!bank_card_id)
                Netparam(nCount, 2) = Trim(pszRsTemp!center_scroll_id)
                Netparam(nCount, 3) = Trim(pszRsTemp!pay_count)
                For i = 1 To ArrayLength(paszResult(1).aszSeatNo)
                    Dim szTmp As String

                    apiTicketInfo(1).aptPrintTicketInfo(i).nTicketType = paszResult(nCount2).aszTicketType(i) 'Trim(pszRsTemp!ticket_type)
                    apiTicketInfo(1).aptPrintTicketInfo(i).sgTicketPrice = FormatDbValue(rsTempOther!pay_count) 'paszResult(nCount2).asgTicketPrice(i)
                    apiTicketInfo(1).aptPrintTicketInfo(i).szSeatNo = paszResult(nCount2).aszSeatNo(i)
                    apiTicketInfo(1).aptPrintTicketInfo(i).szTicketNo = GetTicketNo
                    pbIsTakeChild(i) = FormatDbValue(pszRsTemp!has_child)
                    
                    aszRealNameInfo(i).szIDCardNo = Trim(FormatDbValue(rsTempOther!card_id))
                    aszRealNameInfo(i).szPersonName = Trim(FormatDbValue(rsTempOther!passenger))
                    rsTempOther.MoveNext
                    
                    IncTicketNo 1, False
                Next
                aszTerminateName(nCount) = ""
                bSaleChange(nCount) = False
                nBusType(nCount) = TP_NormalBus
                
'                pszRsTemp.MoveNext
            Next nCount

'                lTkNum = lTkNum - ArrayLength(paszResult(nCount2).aszSeatNo)
                PrintNetTicket apiTicketInfo, pszBusDate, pnTicketCount, pszEndStation, pszOffTime, pszBusID, pszVehicleType, pszCheckGate, bSaleChange, aszTerminateName, szSellStationName, anInsurance, nBusType, pbIsTakeChild, aszRealNameInfo
'            IncTicketNo pszRsTemp.RecordCount + 1
            frmNotify.m_szErrorDescription = "   购票成功,    请拿好车票    祝旅途愉快! 欢迎再次光临！"
            frmNotify.Show vbModal
    Exit Sub
here:
    frmNotify.m_szErrorDescription = err.Description
    frmNotify.Show vbModal
End Sub
'
''得到检票口名称和代码
'Private Function GetCheckName(pszCheckGateID As String) As String
'    Dim i As Integer
'    Dim szResult As String
'    Dim nLen As Integer
'    nLen = 0
'    nLen = ArrayLength(m_aszCheckGateInfo)
'    szResult = ""
'    For i = 1 To nLen
'        If Trim(m_aszCheckGateInfo(i, 1)) = Trim(pszCheckGateID) Then
'            szResult = Trim(m_aszCheckGateInfo(i, 2))
'            Exit For
'        End If
'    Next i
'    GetCheckName = szResult
'
'End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtNewTicketNo.Text = GetTicketNo
    txtTicketNo.Text = GetTicketNo(-1)
    fraOther.BackColor = RGB(241, 246, 250)
    dtpDate.Value = Date
    txtCardID = ""
End Sub

Private Sub txtGetTicketID_Change()

'    If Len(txtGetTicketID.Text) = 13 Then
'        txtValiDate.SetFocus
'    ElseIf Len(txtGetTicketID.Text) > 13 Then
'        txtGetTicketID.Text = Left(txtGetTicketID.Text, 13)
'    End If
End Sub



Private Sub txtGetTicketID_GotFocus()
    eNetTicket = GetNetID
End Sub

Private Sub txtValiDate_Change()
''    eNetTicket = Validate
'    If Len(txtValiDate.Text) > 5 Then
'        txtValiDate.Text = Left(txtValiDate.Text, 5)
'    End If
End Sub
