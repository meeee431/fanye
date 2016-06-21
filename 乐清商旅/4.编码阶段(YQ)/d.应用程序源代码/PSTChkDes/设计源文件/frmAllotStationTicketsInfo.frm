VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAllotStationTicketsInfo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "配载站检/售票信息"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   7200
   StartUpPosition =   1  '所有者中心
   Tag             =   "Modal"
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   345
      Left            =   5760
      TabIndex        =   3
      Tag             =   "Modal"
      Top             =   3210
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   609
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
      MICON           =   "frmAllotStationTicketsInfo.frx":0000
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
      Height          =   3120
      Left            =   -60
      TabIndex        =   4
      Top             =   3000
      Width           =   8745
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid vsBusAllotInfo 
      Height          =   1755
      Left            =   120
      TabIndex        =   0
      Top             =   1140
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   3096
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid vsSellStationInfo 
      Height          =   1755
      Left            =   2460
      TabIndex        =   1
      Top             =   1140
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   3096
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid vsTicketsChekInfo 
      Height          =   1755
      Left            =   4830
      TabIndex        =   2
      Top             =   1140
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   3096
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblChangeCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      Height          =   180
      Left            =   1170
      TabIndex        =   19
      Top             =   360
      Width           =   180
   End
   Begin VB.Label lblMergeCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      Height          =   180
      Left            =   2700
      TabIndex        =   18
      Top             =   360
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "改乘张数:"
      Height          =   180
      Left            =   330
      TabIndex        =   17
      Top             =   360
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "并入票数:"
      Height          =   180
      Left            =   1830
      TabIndex        =   16
      Top             =   360
      Width           =   810
   End
   Begin VB.Label lblBusDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2007-02-09"
      Height          =   180
      Left            =   2670
      TabIndex        =   15
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次日期:"
      Height          =   180
      Left            =   1830
      TabIndex        =   14
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "配载站检票信息:"
      Height          =   180
      Left            =   4860
      TabIndex        =   13
      Top             =   870
      Width           =   1350
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "售票点售票信息:"
      Height          =   180
      Left            =   2490
      TabIndex        =   12
      Top             =   870
      Width           =   1350
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站车票信息:"
      Height          =   180
      Left            =   150
      TabIndex        =   11
      Top             =   870
      Width           =   1350
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已检票数:"
      Height          =   180
      Left            =   5280
      TabIndex        =   10
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "售票张数:"
      Height          =   180
      Left            =   3900
      TabIndex        =   9
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "检票车次:"
      Height          =   180
      Left            =   330
      TabIndex        =   8
      Top             =   120
      Width           =   810
   End
   Begin VB.Label lblCheckCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      Height          =   180
      Left            =   6150
      TabIndex        =   7
      Top             =   120
      Width           =   180
   End
   Begin VB.Label lblSellCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      Height          =   180
      Left            =   4740
      TabIndex        =   6
      Top             =   120
      Width           =   180
   End
   Begin VB.Label lblBusID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2051"
      Height          =   180
      Left            =   1170
      TabIndex        =   5
      Top             =   120
      Width           =   360
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808080&
      X1              =   150
      X2              =   6930
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   165
      X2              =   6960
      Y1              =   615
      Y2              =   615
   End
End
Attribute VB_Name = "frmAllotStationTicketsInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private nTicketsSellCount As Integer
Private nTicketsCheckCount As Integer

Public m_szBusID As String
Public m_dtBusDate As Date
Public m_nChangeCount As Integer
Public m_nMergeCount As Integer


Private Sub InitGridHead()
    vsBusAllotInfo.TextMatrix(0, 0) = "上车站"
    vsBusAllotInfo.TextMatrix(0, 1) = "票数"
    
    vsSellStationInfo.TextMatrix(0, 0) = "售票点"
    vsSellStationInfo.TextMatrix(0, 1) = "票数"
    
    vsTicketsChekInfo.TextMatrix(0, 0) = "上车站"
    vsTicketsChekInfo.TextMatrix(0, 1) = "票数"
End Sub

Private Sub cmdCancel_Click()
    nTicketsCheckCount = 0
    nTicketsSellCount = 0
    m_nChangeCount = 0
    m_nMergeCount = 0
    m_szBusID = ""
    m_dtBusDate = cdtEmptyDate
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    lblBusID.Caption = m_szBusID
    lblBusDate.Caption = m_dtBusDate
    InitGridHead
    RefreshAllotInfo
    RefreshSellStationTicketsInfo
    RefreshSellStationTicketsCheckInfo
    lblSellCount.Caption = nTicketsSellCount
    lblCheckCount.Caption = nTicketsCheckCount
    lblChangeCount.Caption = m_nChangeCount
    lblMergeCount.Caption = m_nMergeCount
End Sub

Private Sub RefreshAllotInfo()
On Error GoTo ErrorHandle
    
    Dim rsTemp As Recordset
    Dim i As Integer
    Set rsTemp = g_oChkTicket.GetAllotStationTicketsInfo(lblBusID.Caption, CDate(lblBusDate.Caption))
    
    vsBusAllotInfo.Rows = rsTemp.RecordCount + 1
    
    For i = 1 To rsTemp.RecordCount
        vsBusAllotInfo.TextMatrix(i, 0) = MakeDisplayString(FormatDbValue(rsTemp!sell_station_id), FormatDbValue(rsTemp!sell_station_name))
        vsBusAllotInfo.TextMatrix(i, 1) = FormatDbValue(rsTemp!ticket_num)
        
        rsTemp.MoveNext
    Next i
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub


Private Sub RefreshSellStationTicketsInfo()
On Error GoTo ErrorHandle
    
    Dim rsTemp As Recordset
    Dim oChkTicket As New CheckTicket
    Dim i As Integer
    
    Set rsTemp = oChkTicket.GetSellStationTicketsInfo(lblBusID.Caption, CDate(lblBusDate.Caption))
    
    vsSellStationInfo.Rows = rsTemp.RecordCount + 1
    
    For i = 1 To rsTemp.RecordCount
        vsSellStationInfo.TextMatrix(i, 0) = MakeDisplayString(FormatDbValue(rsTemp!sell_station_id), FormatDbValue(rsTemp!sell_station_name))
        vsSellStationInfo.TextMatrix(i, 1) = FormatDbValue(rsTemp!ticket_num)
        nTicketsSellCount = nTicketsSellCount + Val(rsTemp!ticket_num)
        
        rsTemp.MoveNext
    Next i
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub RefreshSellStationTicketsCheckInfo()
On Error GoTo ErrorHandle
    
    Dim rsTemp As Recordset
    Dim i As Integer
    
    Set rsTemp = g_oChkTicket.GetAllotTicketsCheckInfo(lblBusID.Caption, CDate(lblBusDate.Caption))
    
    vsTicketsChekInfo.Rows = rsTemp.RecordCount + 1
    
    For i = 1 To rsTemp.RecordCount
        vsTicketsChekInfo.TextMatrix(i, 0) = MakeDisplayString(FormatDbValue(rsTemp!sell_station_id), FormatDbValue(rsTemp!sell_station_name))
        vsTicketsChekInfo.TextMatrix(i, 1) = FormatDbValue(rsTemp!ticket_num)
        nTicketsCheckCount = nTicketsCheckCount + Val(rsTemp!ticket_num)
    
        rsTemp.MoveNext
    Next i
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    nTicketsCheckCount = 0
    nTicketsSellCount = 0
    m_nChangeCount = 0
    m_nMergeCount = 0
    m_szBusID = ""
    m_dtBusDate = cdtEmptyDate
End Sub
