VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmChangeTicketType 
   Caption         =   "全特票转换"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6525
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton cmdChange 
      Height          =   345
      Left            =   2730
      TabIndex        =   2
      Top             =   4320
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "转换(&P)"
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
      MICON           =   "frmChangeTicketType.frx":0000
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
      Cancel          =   -1  'True
      Height          =   345
      Left            =   4380
      TabIndex        =   3
      Top             =   4320
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "取消(&E)"
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
      MICON           =   "frmChangeTicketType.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   840
      Left            =   0
      TabIndex        =   25
      Top             =   4080
      Width           =   8745
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "注意"
      Height          =   720
      Left            =   570
      TabIndex        =   23
      Top             =   870
      Width           =   5535
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   " 此功能用于将全票转换成特票，或将特票转换成全票。"
         Height          =   225
         Left            =   960
         TabIndex        =   24
         Top             =   270
         Width           =   4470
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   225
         Picture         =   "frmChangeTicketType.frx":0038
         Top             =   195
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "车票摘要信息"
      Height          =   1845
      Left            =   570
      TabIndex        =   6
      Top             =   2130
      Width           =   5535
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "200.00"
         Height          =   180
         Left            =   3720
         TabIndex        =   30
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票价:"
         Height          =   180
         Left            =   2880
         TabIndex        =   29
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "售出"
         Height          =   180
         Left            =   1140
         TabIndex        =   28
         Top             =   1470
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "状态:"
         Height          =   180
         Left            =   240
         TabIndex        =   27
         Top             =   1470
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次代码:"
         Height          =   180
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   810
      End
      Begin VB.Label lblBusID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1234"
         Height          =   180
         Left            =   1140
         TabIndex        =   21
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发车时间:"
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   630
         Width           =   810
      End
      Begin VB.Label lblBusStartTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2006-07-02"
         Height          =   180
         Left            =   1140
         TabIndex        =   19
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "站点:"
         Height          =   180
         Left            =   2880
         TabIndex        =   18
         Top             =   360
         Width           =   450
      End
      Begin VB.Label lblStationName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "绍兴"
         Height          =   180
         Left            =   3720
         TabIndex        =   17
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上车站:"
         Height          =   180
         Left            =   240
         TabIndex        =   16
         Top             =   930
         Width           =   630
      End
      Begin VB.Label lblSellStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "玉环"
         Height          =   180
         Left            =   1140
         TabIndex        =   15
         Top             =   930
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票种:"
         Height          =   180
         Left            =   2880
         TabIndex        =   14
         Top             =   930
         Width           =   450
      End
      Begin VB.Label lblTicketTypeName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "全票"
         Height          =   180
         Left            =   3720
         TabIndex        =   13
         Top             =   930
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "座位:"
         Height          =   180
         Left            =   2880
         TabIndex        =   12
         Top             =   660
         Width           =   450
      End
      Begin VB.Label lblSeatNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "01"
         Height          =   180
         Left            =   3720
         TabIndex        =   11
         Top             =   630
         Width           =   180
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "售票员:"
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label lblOperator 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "范鹏东"
         Height          =   180
         Left            =   1140
         TabIndex        =   9
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "售票时间:"
         Height          =   180
         Left            =   2880
         TabIndex        =   8
         Top             =   1470
         Width           =   810
      End
      Begin VB.Label lblSellTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2006-07-02 10:00:00"
         Height          =   180
         Left            =   3720
         TabIndex        =   7
         Top             =   1470
         Width           =   1710
      End
   End
   Begin VB.TextBox txtTicketID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2235
      TabIndex        =   1
      Text            =   "0000001"
      Top             =   1710
      Width           =   1410
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   60
      ScaleHeight     =   795
      ScaleWidth      =   7185
      TabIndex        =   0
      Top             =   0
      Width           =   7185
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   0
         TabIndex        =   4
         Top             =   750
         Width           =   7215
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请输入需要转换的票号:"
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   1890
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "需要转换的票号(&N):"
      Height          =   180
      Left            =   585
      TabIndex        =   26
      Top             =   1755
      Width           =   1620
   End
End
Attribute VB_Name = "frmChangeTicketType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_oSell As New SellTicketClient
Dim m_oParam As New SystemParam
Const cszScrollBus = "滚动"

Private Sub cmdChange_Click()
    If txtTicketID.Text = "" Then
        MsgBox "请输入需要转换的票号", vbInformation, "全特票转换"
        txtTicketID.SetFocus
        Exit Sub
    End If
    
    Dim szTicketType As String
    If ResolveDisplay(lblTicketTypeName.Caption) = 1 Then
        szTicketType = "是否确认要将[全票]转换成[特票]？"
    ElseIf ResolveDisplay(lblTicketTypeName.Caption) = m_oParam.SpecialTicketTypePosition Then
        szTicketType = "是否确认要将[特票]转换成[全票]？"
    Else
        MsgBox "此票不是全票或特票，所以不能转换票种！", vbInformation, "全特票转换"
        txtTicketID.SetFocus
        Exit Sub
    End If
    
    If MsgBox(szTicketType, vbYesNo, "提示") = vbYes Then
        Dim bChangeStatus As Boolean
        bChangeStatus = m_oSell.ChangeTicketType(txtTicketID.Text, IIf(ResolveDisplay(lblTicketTypeName.Caption) = 1, m_oParam.SpecialTicketTypePosition, 1))
        If bChangeStatus = True Then
            lblTicketTypeName.Caption = IIf(ResolveDisplay(lblTicketTypeName.Caption) = 1, MakeDisplayString(m_oParam.SpecialTicketTypePosition, Trim(GetTicketTypeStr2(m_oParam.SpecialTicketTypePosition))), MakeDisplayString(1, Trim(GetTicketTypeStr2(1))))
            ShowMsg "票种转换成功！"
        Else
            ShowMsg "特票已售完，票种转换失败！"
        End If
    End If
    
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtTicketID.Text = ""
    lblTicketTypeName.ForeColor = vbRed
    InitForm
    m_oSell.Init g_oActiveUser
    m_oParam.Init g_oActiveUser
End Sub

Private Sub InitForm()
    lblBusStartTime.Caption = ""
    lblSellStation.Caption = ""
    lblBusID.Caption = ""
    lblStationName.Caption = ""
    lblSeatNo.Caption = ""
    lblTicketTypeName.Caption = ""
    lblSellTime.Caption = ""
    lblOperator.Caption = ""
    lblStatus.Caption = ""
    lblPrice.Caption = ""
End Sub

Private Sub Form_Resize()
    txtTicketID.SetFocus
End Sub

Private Sub txtTicketID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtTicketID.Text <> "" Then
        InitForm
        RefreshTicketInfo
        cmdChange.SetFocus
    End If
End Sub

Private Sub RefreshTicketInfo()
'得到车票信息显示在界面上
    Dim oTicket As ServiceTicket
    Dim oCTicket As ClientTicket
    Dim oREBus As REBus
    Dim szTemp As String
On Error GoTo Here
    
    Set oTicket = m_oSell.GetTicket(txtTicketID.Text)
    Set oCTicket = m_oSell.GetTicketClient(txtTicketID.Text)
    
        If Not oCTicket Is Nothing Then
            If Trim(oCTicket.UnitID) = Trim(g_oActiveUser.UserUnitID) Then
                Set oREBus = m_oSell.CreateServiceObject("STReSch.REBus")
                oREBus.Init g_oActiveUser
                oREBus.Identify oTicket.REBusID, oTicket.REBusDate
                If oREBus.BusType <> TP_ScrollBus Then
                    lblBusStartTime.Caption = oREBus.StartUpTime
                Else
                    lblBusStartTime.Caption = cszScrollBus
                End If
            Else
                lblBusStartTime = "远程车票..."
            End If
            lblSellStation.Caption = oCTicket.StartStaionName
            lblBusID.Caption = oCTicket.REBusID
            lblStationName.Caption = oCTicket.ToStationName
            lblSeatNo.Caption = oCTicket.SeatNo
            lblTicketTypeName.Caption = MakeDisplayString(oCTicket.TicketType, Trim(GetTicketTypeStr2(oCTicket.TicketType)))
            lblSellTime.Caption = oCTicket.SellTime
            lblOperator.Caption = oCTicket.Operator
            lblPrice.Caption = FormatMoney(oCTicket.TicketPrice)
            
            If (oCTicket.TicketStatus And ST_TicketNormal) <> 0 Then
                szTemp = "正常售出"
            Else
                szTemp = "改签售出"
            End If
    
            If (oCTicket.TicketStatus And ST_TicketCanceled) <> 0 Then
                szTemp = szTemp & "/已废"
            ElseIf (oCTicket.TicketStatus And ST_TicketChanged) <> 0 Then
                szTemp = szTemp & "/已被改签"
            ElseIf (oCTicket.TicketStatus And ST_TicketChecked) <> 0 Then
                szTemp = szTemp & "/已检"
            ElseIf (oCTicket.TicketStatus And ST_TicketReturned) <> 0 Then
                szTemp = szTemp & "/已退"
            End If
            lblStatus.Caption = szTemp
        Else
            Set oREBus = m_oSell.CreateServiceObject("STReSch.REBus")
            oREBus.Init g_oActiveUser
            oREBus.Identify oTicket.REBusID, oTicket.REBusDate
            If oREBus.BusType <> TP_ScrollBus Then
               lblBusStartTime.Caption = oREBus.StartUpTime
            Else
               lblBusStartTime.Caption = cszScrollBus
            End If
            lblSellStation.Caption = m_oSell.SellUnitShortName
            lblBusID.Caption = oTicket.REBusID
            lblStationName.Caption = oTicket.ToStationName
            lblSeatNo.Caption = oTicket.SeatNo
            lblTicketTypeName.Caption = MakeDisplayString(oTicket.TicketType, Trim(GetTicketTypeStr2(oTicket.TicketType)))
            lblSellTime.Caption = oTicket.SellTime
            lblOperator.Caption = oTicket.Operator
            lblPrice.Caption = FormatMoney(oTicket.TicketPrice)
            
            If (oCTicket.TicketStatus And ST_TicketNormal) <> 0 Then
                szTemp = "正常售出"
            Else
                szTemp = "改签售出"
            End If
    
            If (oCTicket.TicketStatus And ST_TicketCanceled) <> 0 Then
                szTemp = szTemp & "/已废"
            ElseIf (oCTicket.TicketStatus And ST_TicketChanged) <> 0 Then
                szTemp = szTemp & "/已被改签"
            ElseIf (oCTicket.TicketStatus And ST_TicketChecked) <> 0 Then
                szTemp = szTemp & "/已检"
            ElseIf (oCTicket.TicketStatus And ST_TicketReturned) <> 0 Then
                szTemp = szTemp & "/已退"
            End If
            lblStatus.Caption = oTicket.TicketStatus
            
        End If
    
    Exit Sub
Here:
    Set oCTicket = Nothing
    Set oTicket = Nothing
    ShowErrorMsg
End Sub

Private Sub txtTicketID_GotFocus()
    txtTicketID.SelStart = 0
    txtTicketID.SelLength = Len(txtTicketID.Text)
End Sub

Public Function GetTicketTypeStr2(ByVal pnTicketType As Integer) As String
Dim j As Integer
Dim TicketType() As TTicketType
Dim intEnableTicketNo As Integer

   TicketType = m_oSell.GetAllTicketType(1)
   intEnableTicketNo = UBound(TicketType) - LBound(TicketType) + 1
    For j = 1 To intEnableTicketNo
        If TicketType(j).nTicketTypeID = pnTicketType And TicketType(j).nTicketTypeValid = TP_TicketTypeValid Then
           GetTicketTypeStr2 = TicketType(j).szTicketTypeName
           Exit For
        End If
    Next j
End Function

