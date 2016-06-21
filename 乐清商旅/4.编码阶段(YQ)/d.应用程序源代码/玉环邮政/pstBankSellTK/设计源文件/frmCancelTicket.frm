VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmCancelTicket 
   BackColor       =   &H8000000C&
   Caption         =   "废票"
   ClientHeight    =   7935
   ClientLeft      =   1965
   ClientTop       =   2010
   ClientWidth     =   10860
   HelpContextID   =   4000220
   Icon            =   "frmCancelTicket.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   10860
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrConnected 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   465
   End
   Begin VB.Frame fraOutLine 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7005
      Left            =   360
      TabIndex        =   3
      Top             =   180
      Width           =   9900
      Begin VB.ComboBox cboStartStation 
         Height          =   300
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   225
         Width           =   1890
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
         Height          =   400
         Left            =   5355
         MaxLength       =   10
         TabIndex        =   1
         Top             =   125
         Width           =   1950
      End
      Begin RTComctl3.CoolButton cmdCancelTicket 
         Height          =   435
         Left            =   7920
         TabIndex        =   2
         Top             =   930
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   767
         BTYPE           =   3
         TX              =   "废票(&T)"
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
         MICON           =   "frmCancelTicket.frx":014A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame fraTktInfoChange 
         BackColor       =   &H00E0E0E0&
         Caption         =   "车票信息"
         Height          =   2775
         Left            =   150
         TabIndex        =   4
         Top             =   765
         Width           =   7245
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "票号:"
            Height          =   180
            Left            =   135
            TabIndex        =   28
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发车时间:"
            Height          =   180
            Left            =   135
            TabIndex        =   27
            Top             =   1335
            Width           =   810
         End
         Begin VB.Label lblOperatorChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "售票员:"
            Height          =   180
            Left            =   135
            TabIndex        =   26
            Top             =   1710
            Width           =   630
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "起站:"
            Height          =   180
            Index           =   1
            Left            =   135
            TabIndex        =   25
            Top             =   570
            Width           =   450
         End
         Begin VB.Label lblScheduleChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "车次:"
            Height          =   180
            Left            =   135
            TabIndex        =   24
            Top             =   945
            Width           =   450
         End
         Begin VB.Label lblTimeChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "售票时间:"
            Height          =   180
            Left            =   135
            TabIndex        =   23
            Top             =   2460
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "票价:"
            Height          =   180
            Left            =   135
            TabIndex        =   22
            Top             =   2085
            Width           =   450
         End
         Begin VB.Label lblTicketID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A0000134590"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1170
            TabIndex        =   21
            Tag             =   "lblCurrentTktNum"
            Top             =   210
            Width           =   1320
         End
         Begin VB.Label lblSeatNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "01"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   4890
            TabIndex        =   20
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "座位号:"
            Height          =   180
            Left            =   3975
            TabIndex        =   19
            Top             =   1710
            Width           =   630
         End
         Begin VB.Label lblTypeChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "票种:"
            Height          =   180
            Left            =   3975
            TabIndex        =   18
            Top             =   1335
            Width           =   450
         End
         Begin VB.Label label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "到站:"
            Height          =   180
            Left            =   3975
            TabIndex        =   17
            Top             =   570
            Width           =   450
         End
         Begin VB.Label lblStateChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "状态:"
            Height          =   180
            Left            =   3975
            TabIndex        =   16
            Top             =   2085
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "日期:"
            Height          =   180
            Left            =   3975
            TabIndex        =   15
            Top             =   945
            Width           =   450
         End
         Begin VB.Label lblBusID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "25101"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1170
            TabIndex        =   14
            Top             =   915
            Width           =   600
         End
         Begin VB.Label lblEndStation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "杭州"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   4890
            TabIndex        =   13
            Top             =   540
            Width           =   480
         End
         Begin VB.Label lblStartStation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "宁波南站"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1170
            TabIndex        =   12
            Top             =   540
            Width           =   960
         End
         Begin VB.Label lblSeller 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "张三"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1170
            TabIndex        =   11
            Top             =   1680
            Width           =   480
         End
         Begin VB.Label lblTicketType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "全票"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   4890
            TabIndex        =   10
            Top             =   1305
            Width           =   480
         End
         Begin VB.Label lblSellTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2002-07-15 07:00:00"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1170
            TabIndex        =   9
            Top             =   2445
            Visible         =   0   'False
            Width           =   2280
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2002-07-15"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   4890
            TabIndex        =   8
            Top             =   915
            Width           =   1200
         End
         Begin VB.Label lblTicketPrice 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "37.50"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1170
            TabIndex        =   7
            Top             =   2055
            Width           =   600
         End
         Begin VB.Label lblStatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "正常售出"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   4890
            TabIndex        =   6
            Top             =   2055
            Width           =   960
         End
         Begin VB.Label lblOffTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "10:00"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1170
            TabIndex        =   5
            Top             =   1305
            Width           =   600
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起点站:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   30
         Top             =   285
         Width           =   840
      End
      Begin VB.Label lblOldTktNum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "废票票号(&Z):"
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
         Left            =   3810
         TabIndex        =   0
         Top             =   255
         Width           =   1440
      End
   End
   Begin MSWinsockLib.Winsock wsClient 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCancelTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'废票中用的枚举

Private Enum CancelTicketInfo
    CT_BusID = 1
    CT_StartStation = 2
    CT_EndStation = 3
    CT_Date = 4
    CT_OffTime = 5
    CT_SeatNo = 6
    CT_Status = 7
    CT_SellTime = 8
    CT_TicketPrice = 9
    CT_TicketType = 10
    CT_Seller = 11
End Enum

Private m_szAllSend As String



Private Sub cmdCancelTicket_Click()
'    Dim aszCancelTicket() As String
    On Error GoTo Here
        If txtTicketNo.Text = "" Then Exit Sub

        If MsgBox("是否确认废除这些票？", vbYesNo, "提示") = vbYes Then
            If Right(lblStatus.Caption, 2) = "已废" Then
                MsgBox "此票已废！" & vbCrLf, vbInformation, "废票"
                Exit Sub
            Else
                SendCancelTicketRequest txtTicketNo.Text
            End If
        End If
    Exit Sub
Here:
    ShowErrorMsg
End Sub



Private Sub cmdRefresh_Click()
On Error GoTo Here
    SendQueryTicketRequest
'    SerialCancelTkt
On Error GoTo 0
Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub Form_Activate()
On Error GoTo Here

    txtTicketNo.Text = GetTicketNo(-1)
    MDISellTicket.SetFunAndUnit
Exit Sub
Here:
    ShowErrorMsg
'-------------------
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo Error_Handle
    If KeyAscii = vbKeyReturn And (Not Me.ActiveControl Is txtTicketNo) Then
        SendKeys "{TAB}"
    ElseIf KeyAscii = vbKeyEscape Then
    
    
        txtTicketNo.SetFocus
        EnableCancelButton
    ElseIf KeyAscii = Asc("+") Then
        '如果按了加号
        '则继续可以废下一张

    End If
    Exit Sub
Error_Handle:
    ShowErrorMsg

End Sub

'初始化winsock
Private Sub InitSock()
    
    wsClient.Close
    wsClient.RemoteHost = m_szRemoteHost
    wsClient.RemotePort = m_szRemotePort
    wsClient.Connect
    
End Sub

Private Sub Form_Load()


On Error GoTo Here

    '===============================
    '初始化winsock控件
    '===============================
    InitValue
    InitSock
    '===============================
    
    FillStartStation
    
    
    txtTicketNo.MaxLength = 10
'    FillColumnHeader
    EnableCancelButton
    SetDefaultValue

    On Error GoTo 0
Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub Form_Resize()
    If MDISellTicket.ActiveForm Is Me Then
        If Not Me.WindowState = vbMaximized Then Me.WindowState = vbMaximized
        fraOutLine.Left = (Me.ScaleWidth - fraOutLine.Width) / 2
        fraOutLine.Top = (Me.ScaleHeight - fraOutLine.Height) / 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    MDISellTicket.lblCancel.Value = vbUnchecked
    MDISellTicket.abMenuTool.Bands("mnuFunction").Tools("mnuCancelTkt").Checked = False
'    MDISellTicket.mnuCancelTkt.Checked = False
End Sub

Private Sub SendQueryTicketRequest()
    '发送查询票信息的请求
    Dim szSend As String
    
    szSend = GetQueryTicketInfoRequestStr(txtTicketNo.Text, 0, g_aszAllStartStation(cboStartStation.ListIndex + 1, 1), "", "", Date, "", "", Date, "", "", 0, 0) '除票号外,其他的参数均为随便传入的.
    
    
    wsClient.SendData szSend
    
    
End Sub


Private Sub RefreshTicketInfo(pszReceive As String)
    '显示车票信息

    On Error GoTo Here
    Dim szReserve As String
    Dim szStartStationName As String
    Dim szEndStationName As String
    Dim szStatus As String
    
    If txtTicketNo.Text <> "" Then
        
        If Val(GetBusType(pszReceive)) <> TP_ScrollBus Then
            lblOffTime.Caption = PackageToTime(GetBusOffTime(pszReceive))
        Else
            lblOffTime.Caption = cszScrollBus
        End If

        lblBusID.Caption = GetPackageBusID(pszReceive)
        lblDate.Caption = PackageToDate(GetBusOffDate(pszReceive))
        lblTicketType.Caption = MakeDisplayString(GetTicketType(pszReceive), GetTicketTypeName(GetTicketType(pszReceive)))
        
        lblTicketPrice.Caption = FormatMoney(PackageToMoney(GetTicketPrice(pszReceive)))
        lblSeatNo.Caption = GetSeatID(pszReceive)
        lblTicketID.Caption = txtTicketNo.Text
        lblSeller.Caption = GetOperatorID(pszReceive)
        
        
        szReserve = GetReserved(pszReceive)
        '起点站名\到站名\票状态都放在预留信息中.    需把它分解一下
        szStartStationName = MidA(szReserve, 1, 10)
        szEndStationName = MidA(szReserve, 11, 10)
        szStatus = MidA(szReserve, 21, 2)
        
        
        lblEndStation.Caption = szEndStationName
        lblStartStation.Caption = szStartStationName
        lblStatus.Caption = GetTicketStatusStr(Val(szStatus))
                
    End If

    On Error GoTo 0
    Exit Sub
Here:
    SetDefaultValue
    ShowErrorMsg
End Sub
'显示HTMLHELP,直接拷贝
Private Sub DisplayHelp(Optional HelpType As EHelpType = content)
    Dim lActiveControl As Long

    Select Case HelpType
        Case content
            lActiveControl = Me.ActiveControl.HelpContextID
            If lActiveControl = 0 Then
                TopicID = Me.HelpContextID
                CallHTMLShowTopicID
            Else
                TopicID = lActiveControl
                CallHTMLShowTopicID
            End If
        Case Index
            CallHTMLHelpIndex
        Case Support
            TopicID = clSupportID
            CallHTMLShowTopicID
    End Select
End Sub

Private Sub TicketNumberAddOne()
    Dim count As Integer
    Dim TxtLenth As Integer
    Dim TicketNumber As String
    Dim ZeroNumber As Integer

    TxtLenth = Len(txtTicketNo.Text)
    For count = 1 To TxtLenth
       If Asc(Mid(txtTicketNo.Text, count, 1)) >= 48 And Asc(Mid(txtTicketNo.Text, count, 1)) <= 57 Then
          TicketNumber = Right(txtTicketNo.Text, TxtLenth - count + 1) + 1
          Do While Len(Right(txtTicketNo.Text, TxtLenth - count + 1)) > Len(TicketNumber)
             TicketNumber = "0" & TicketNumber
          Loop
          txtTicketNo.Text = Left(txtTicketNo.Text, count - 1) & TicketNumber
          Exit For
       End If
    Next count
End Sub

Private Sub tmrConnected_Timer()
    '说明连接成功
    tmrConnected.Enabled = False
    
End Sub


Private Sub txtTicketNo_Change()
    EnableCancelButton
End Sub

Private Sub txtTicketNo_GotFocus()
        txtTicketNo.SelStart = 0
        txtTicketNo.SelLength = 100 'Len(txtTicketNo.Text)
        cmdCancelTicket.Default = False
End Sub

Private Sub txtTicketNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim szTemp As String
    Dim lTemp As Long
    On Error GoTo Error_Handel
    If Len(txtTicketNo.Text) >= TicketNoNumLen() Then
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
            lTemp = Right(txtTicketNo.Text, TicketNoNumLen())
            szTemp = Left(txtTicketNo.Text, Len(txtTicketNo.Text) - TicketNoNumLen())

            If KeyCode = vbKeyUp Then
                lTemp = lTemp + 1
            Else
                lTemp = lTemp - 1
                lTemp = IIf(lTemp > 0, lTemp, 0)
            End If
            txtTicketNo.Text = MakeTicketNo(lTemp, szTemp)
            KeyCode = 0
        End If
    End If
    Exit Sub
Error_Handel:
End Sub

'设置退票按钮状态
Private Sub EnableCancelButton()
'    设置废票按钮
    If txtTicketNo.Text <> "" Then
        cmdCancelTicket.Enabled = True
    Else
        cmdCancelTicket.Enabled = False
    End If

End Sub

Private Sub txtTicketNo_KeyPress(KeyAscii As Integer)
On Error GoTo Here
    If KeyAscii = 13 And txtTicketNo.Text <> "" Then
        SendQueryTicketRequest
        If cmdCancelTicket.Enabled Then cmdCancelTicket.SetFocus
    End If
On Error GoTo 0
Exit Sub
Here:

  ShowErrorMsg
End Sub

Private Sub SetDefaultValue()
    '设置默认的控件值
    lblStartStation.Caption = ""
    lblEndStation.Caption = ""
    lblBusID.Caption = ""
    lblDate.Caption = ""
    lblOffTime.Caption = ""
    lblTicketType.Caption = ""
    lblSeller.Caption = ""
    lblSeatNo.Caption = ""
    lblTicketPrice.Caption = ""
    lblStatus.Caption = ""
    lblSellTime.Caption = ""
    lblTicketID.Caption = ""

End Sub






Private Sub wsClient_Connect()
    tmrConnected.Enabled = True
End Sub

Private Sub wsClient_DataArrival(ByVal bytesTotal As Long)
    '数据到达处理
    Dim szReceive As String
'    Debug.Print "DataArrival" & "," & bytesTotal
'    If bytesTotal >= FIXPKGLEN Then
        wsClient.GetData szReceive
        
        '下面这个条件处理是针对有时包会转成两个来发送的.处理情况
        If Right(szReceive, 1) = cszPackageEnd And Left(szReceive, 1) = cszPackageBegin Then
            '如果数据包以"B"开头,以"@"结尾
'            Debug.Print szReceive, bytesTotal
            RxPkgProcess szReceive
        Else
            
            If Left(szReceive, 1) = cszPackageBegin Then
                '为第一串包
                m_szAllSend = szReceive
            Else
                '为后续包,则进行合并
                m_szAllSend = m_szAllSend & szReceive
            End If
            If Right(szReceive, 1) = cszPackageEnd Then
                
'                Debug.Print m_szAllSend, bytesTotal
                RxPkgProcess m_szAllSend
            End If
        End If
'    End If
End Sub

Private Sub wsClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Error(" & "," & Number & "," & Description & "," & Scode & "," & Source & "," & HelpFile & "," & HelpContext & "," & CancelDisplay

End Sub


Private Sub RxPkgProcess(pszReceive As String)
    '对返回的数据进行处理
    
    Dim szTradeID As String
    Dim szTemp As String
    Dim szStationInfo As String
    Dim szBusInfo As String
    Dim nLen As Integer
    Dim rsTemp As Recordset
    
    
    szTradeID = GetTradeID(pszReceive)
    nLen = GetLen(pszReceive)
    
    Select Case szTradeID
    
    Case GETTKINFO
        '得到车票信息
        Debug.Print pszReceive
        On Error GoTo jerr
            If GetRetCode(pszReceive) <> cszCorrectRetCode Then
                MsgBox MidA(pszReceive, cnPosReserved, LenA(pszReceive) - cnPosReserved), vbOKOnly, "错误" & GetRetCode(pszReceive)
            Else
                RefreshTicketInfo pszReceive
            End If
        
        On Error GoTo 0
    Case CANCELTICKETSID
        '废票
        Debug.Print pszReceive
        On Error GoTo jerr
            If GetRetCode(pszReceive) <> cszCorrectRetCode Then
                MsgBox MidA(pszReceive, cnPosReserved, LenA(pszReceive) - cnPosReserved), vbOKOnly, "错误" & GetRetCode(pszReceive)
            Else
                CancelTicket pszReceive
            End If
        
        On Error GoTo 0
    Case Else
    End Select
    Exit Sub
    
jerr:
    ShowErrorMsg
End Sub


Private Sub SendCancelTicketRequest(pszTicketID As String)
    '进行废票请求操作
    Dim szTemp As String
    szTemp = GetCancelTicketRequestStr(ResolveDisplay(cboStartStation.Text), pszTicketID, lblTicketPrice.Caption)
    wsClient.SendData szTemp

End Sub

Private Sub CancelTicket(pszReceive)

    ShowMsg "正常废票成功！"
    EnableCancelButton
    '刷新该票的最新情况
    SendQueryTicketRequest
    txtTicketNo.SetFocus
    
    Exit Sub
    '不进行强行废票处理
'forcecancel:
'    On Error GoTo Here
''        m_oSell.ForceCancelTicket aszCancelTicket
'        SendQueryTicketRequest
'
'        ShowMsg "强行废票成功！"
'        SetDefaultValue
'        txtTicketNo.SetFocus
    Exit Sub
Here:
'    If err.Number = ERR_CancelTicketTimeOut Then
'        If MsgBox("已过正常废票时间,要否强行废票？", vbYesNo Or vbQuestion Or vbDefaultButton2) = vbYes Then
'            Resume forcecancel
'        End If
'    ElseIf err.Number = ERR_UserIsNotTicketSeller Then
'        If MsgBox("您不是此票的售票员,要否强行废票？", vbYesNo Or vbQuestion Or vbDefaultButton2) = vbYes Then
'            Resume forcecancel
'        End If
'    Else
        ShowErrorMsg
'    End If
End Sub

'填充起点站
Private Sub FillStartStation()
    Dim aszTemp() As String
    Dim i As Integer
    
    aszTemp = g_aszAllStartStation
    cboStartStation.Clear
    With cboStartStation
        For i = 1 To ArrayLength(aszTemp)
            .AddItem MakeDisplayString(aszTemp(i, 1), aszTemp(i, 3))
        Next i
        If .ListCount > 0 Then
            .ListIndex = 0
        End If
    End With
    
    
End Sub


