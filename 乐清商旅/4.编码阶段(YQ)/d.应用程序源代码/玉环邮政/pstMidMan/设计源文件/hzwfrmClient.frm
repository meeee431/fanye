VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmclient 
   Caption         =   "售票"
   ClientHeight    =   3408
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   3408
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetSeats 
      Caption         =   "座位数"
      Height          =   252
      Left            =   3512
      TabIndex        =   27
      Top             =   2880
      Width           =   612
   End
   Begin VB.CommandButton cmdAccount 
      Caption         =   "帐单"
      Height          =   252
      Left            =   1936
      TabIndex        =   26
      Top             =   2880
      Width           =   612
   End
   Begin VB.TextBox amount 
      Height          =   492
      Left            =   3960
      TabIndex        =   24
      Text            =   " "
      Top             =   1920
      Width           =   1092
   End
   Begin VB.TextBox OperatorID 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#####"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      Height          =   372
      Left            =   3960
      MaxLength       =   6
      TabIndex        =   22
      Text            =   " "
      Top             =   600
      Width           =   972
   End
   Begin VB.TextBox ID 
      Height          =   372
      Left            =   1080
      TabIndex        =   21
      Text            =   " "
      Top             =   600
      Width           =   972
   End
   Begin VB.TextBox BankID 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#####"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      Height          =   372
      Left            =   3960
      MaxLength       =   6
      TabIndex        =   19
      Text            =   " "
      Top             =   0
      Width           =   972
   End
   Begin VB.CommandButton cmdCancelTicket 
      Caption         =   "废票"
      Height          =   252
      Left            =   5088
      TabIndex        =   18
      Top             =   2880
      Width           =   612
   End
   Begin VB.CommandButton cmdGetSchedules 
      Caption         =   "取车次"
      Height          =   252
      Left            =   4300
      TabIndex        =   17
      Top             =   2880
      Width           =   612
   End
   Begin VB.CommandButton cmdGetStations 
      Caption         =   "取站点"
      Height          =   252
      Left            =   2724
      TabIndex        =   16
      Top             =   2880
      Width           =   612
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "连接"
      Height          =   252
      Left            =   360
      TabIndex        =   15
      Top             =   2880
      Width           =   612
   End
   Begin VB.CommandButton cmdBuyTicket 
      Caption         =   "购票"
      Height          =   252
      Left            =   1148
      TabIndex        =   14
      Top             =   2880
      Width           =   612
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "清除"
      Height          =   252
      Left            =   5880
      TabIndex        =   13
      Top             =   2880
      Width           =   612
   End
   Begin VB.TextBox TradeID 
      Height          =   372
      Left            =   3960
      TabIndex        =   11
      Text            =   " "
      Top             =   1200
      Width           =   972
   End
   Begin VB.OptionButton Repaired 
      Caption         =   " "
      Height          =   252
      Left            =   6000
      TabIndex        =   10
      Top             =   2040
      Width           =   612
   End
   Begin VB.OptionButton SeatNeed 
      Caption         =   " "
      Height          =   252
      Left            =   6000
      TabIndex        =   9
      Top             =   1440
      Width           =   252
   End
   Begin VB.TextBox LineID 
      Height          =   372
      Left            =   1080
      TabIndex        =   2
      Text            =   " "
      Top             =   1200
      Width           =   972
   End
   Begin VB.TextBox TType 
      Height          =   372
      Left            =   1080
      TabIndex        =   1
      Text            =   " "
      Top             =   0
      Width           =   972
   End
   Begin VB.TextBox StartDate 
      Height          =   372
      Left            =   1080
      TabIndex        =   0
      Text            =   " "
      Top             =   1800
      Width           =   972
   End
   Begin MSWinsockLib.Winsock TcpClient 
      Left            =   1440
      Top             =   600
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
      RemoteHost      =   "Fdcmain"
      RemotePort      =   4660
   End
   Begin VB.Label Label10 
      Caption         =   "数量"
      Height          =   372
      Left            =   2760
      TabIndex        =   25
      Top             =   1920
      Width           =   852
   End
   Begin VB.Label Label9 
      Caption         =   "操作员号"
      Height          =   492
      Left            =   2640
      TabIndex        =   23
      Top             =   480
      Width           =   852
   End
   Begin VB.Label Label8 
      Caption         =   "所号"
      Height          =   252
      Left            =   2640
      TabIndex        =   20
      Top             =   120
      Width           =   852
   End
   Begin VB.Label Label7 
      Caption         =   "交易码"
      Height          =   372
      Left            =   2640
      TabIndex        =   12
      Top             =   1200
      Width           =   1092
   End
   Begin VB.Label Label6 
      Caption         =   "补票"
      Height          =   372
      Left            =   5640
      TabIndex        =   8
      Top             =   1680
      Width           =   972
   End
   Begin VB.Label label5 
      Caption         =   "需要座位"
      Height          =   372
      Left            =   5640
      TabIndex        =   7
      Top             =   720
      Width           =   852
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6600
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label4 
      Caption         =   "日期"
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   612
   End
   Begin VB.Label Label3 
      Caption         =   "车次"
      Height          =   372
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   492
   End
   Begin VB.Label Label2 
      Caption         =   "票号"
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   612
   End
   Begin VB.Label Label1 
      Caption         =   "票类型"
      Height          =   252
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   612
   End
End
Attribute VB_Name = "frmclient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type TicketPkgType
  StartByte As String * 1
  PkgLen As String * 8
  TradeID As String * 4
  TicketID As String * 7
  TicketType As String * 1
  StartID As String * 9
  DestID As String * 9
  ScheduleID As String * 4
  BusType As String * 2
  ticketnum As String * 2
  TicketPrice As String * 6
  StartDate As String * 8
  StartTime As String * 4
  SeatNO(40) As String * 2
  OperatorID As String * 5
  BankID    As String * 5
  ReturnID As String * 4
  Reserverd As String * 30
  EndByte As String * 1
End Type
Private Enum TicketType
  FullTicket = 0
  HalfTicket = 1
  FreeTicket = 2
End Enum
Dim TicketPkg As TicketPkgType
Dim dddd


Private Sub cmdAccount_Click()
   Dim sendstr As String
   sendstr = "B0000019080051234598765" & Right(ID, 7) & Right(TType, 1) & _
    String(9, &H30) & "zj       " & Right(LineID, 4) & "00" & Right(amount, 2) _
    & "000000" & Right(StartDate, 8) & String(155, &H30)
   TcpClient.SendData sendstr
End Sub

Private Sub cmdBuyTicket_Click()
  Dim sendstr As String
  sendstr = "B0000019080011234598765" & Right(ID, 7) & Right(TType, 1) & _
    "SX       " & "CPL      " & Right(LineID, 4) & "00" & Right(amount, 2) _
    & "000000" & Right(StartDate, 8) & String(155, &H30)
   TcpClient.SendData sendstr
 
End Sub

Private Sub cmdCancelTicket_Click()
  Dim sendstr As String
  sendstr = "B0000019080111234598765" & Right(ID, 7) & String(160, &H30)
   TcpClient.SendData sendstr
End Sub

Private Sub cmdGetSchedules_Click()
  Dim sendstr As String
  TicketPkg.StartByte = "S"
  TicketPkg.PkgLen = "00000190"
  TicketPkg.TradeID = GETSCHEDULESID
  TicketPkg.BankID = BankID
  TicketPkg.OperatorID = OperatorID
  TradeID = TicketPkg.TradeID
  TicketPkg.EndByte = "E"
  With TicketPkg
    sendstr = .StartByte & .PkgLen & .TradeID & .OperatorID & .BankID
    sendstr = sendstr & String(FIXPKGLEN - Len(sendstr) - 1, Chr$(&H30)) & .EndByte
  End With
  TcpClient.SendData sendstr
End Sub

Private Sub cmdGetStations_Click()
  Dim sendstr As String
  TicketPkg.StartByte = "S"
  TicketPkg.PkgLen = "00000190"
  TicketPkg.TradeID = GETSTATIONSID
  TicketPkg.BankID = Right(BankID, Len(BankID) - 1)
  TicketPkg.OperatorID = Right(OperatorID, Len(OperatorID) - 1)
  TradeID = TicketPkg.TradeID
  TicketPkg.EndByte = "E"
  With TicketPkg
    sendstr = .StartByte & .PkgLen & .TradeID & .OperatorID & .BankID
    sendstr = sendstr & String(FIXPKGLEN - Len(sendstr) - 1, Chr$(&H30)) & .EndByte
  End With
  MsgBox sendstr
  TcpClient.SendData sendstr
End Sub

Private Sub Form_Load()
 TcpClient.RemoteHost = "192.168.1.118"
 TcpClient.RemotePort = 8000
 dddd = Format$(Now, "yyyy-mm-dd")
 TType = "0"
 TradeID = "8001"
 StartDate = Format$(Now, "yyyymmdd")
 amount = "01"
 ID = Str(CLng(10000000 * Rnd(10 * Rnd())))
 SeatNeed = True
 Repaired = False
 End Sub
Private Sub cmdConnect_Click()
  
 If TcpClient.State = sckClosed Then
   TcpClient.Connect
 End If
 

   
End Sub
 
 
 
Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
 Dim strData As String
 TcpClient.GetData strData
 MsgBox "len=[" & Str(Len(strData)) & "]" & vbCr & strData, vbOKOnly
 
 End Sub
 
 
