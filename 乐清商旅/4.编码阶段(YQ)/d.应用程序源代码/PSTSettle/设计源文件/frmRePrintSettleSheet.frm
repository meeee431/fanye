VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmRePrintSettleSheet 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "重打结算单"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   Icon            =   "frmRePrintSettleSheet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   5880
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -30
      TabIndex        =   25
      Top             =   810
      Width           =   8775
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   990
      Left            =   -60
      TabIndex        =   21
      Top             =   4110
      Width           =   8745
      Begin RTComctl3.CoolButton cmdExit 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   4350
         TabIndex        =   3
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "关闭(&E)"
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
         MICON           =   "frmRePrintSettleSheet.frx":030A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdRePrint 
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   210
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "重打结算单(&R)"
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
         MICON           =   "frmRePrintSettleSheet.frx":0326
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "结算单摘要信息"
      Height          =   1605
      Left            =   150
      TabIndex        =   8
      Top             =   2400
      Width           =   5535
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "状态:"
         Height          =   180
         Left            =   270
         TabIndex        =   24
         Top             =   300
         Width           =   450
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已结"
         Height          =   180
         Left            =   1290
         TabIndex        =   23
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总路单数:"
         Height          =   180
         Left            =   270
         TabIndex        =   20
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "456"
         Height          =   180
         Left            =   1290
         TabIndex        =   19
         Top             =   600
         Width           =   270
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总人数:"
         Height          =   180
         Left            =   270
         TabIndex        =   18
         Top             =   900
         Width           =   630
      End
      Begin VB.Label lblNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1234"
         Height          =   180
         Left            =   1260
         TabIndex        =   17
         Top             =   900
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应结票款:"
         Height          =   180
         Left            =   2820
         TabIndex        =   16
         Top             =   570
         Width           =   810
      End
      Begin VB.Label lblNeedSplitOut 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1255"
         Height          =   210
         Left            =   3900
         TabIndex        =   15
         Top             =   540
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总票款:"
         Height          =   180
         Left            =   2820
         TabIndex        =   14
         Top             =   270
         Width           =   630
      End
      Begin VB.Label lblTotalPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4568"
         Height          =   210
         Left            =   3930
         TabIndex        =   13
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算期限:"
         Height          =   180
         Left            =   270
         TabIndex        =   12
         Top             =   1200
         Width           =   810
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2003年3月17日--2003年03月31日"
         Height          =   180
         Left            =   1200
         TabIndex        =   11
         Top             =   1200
         Width           =   2610
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算人:"
         Height          =   180
         Left            =   2970
         TabIndex        =   10
         Top             =   900
         Width           =   630
      End
      Begin VB.Label lblOperator 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "陈峰"
         Height          =   180
         Left            =   3840
         TabIndex        =   9
         Top             =   870
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "注意"
      Height          =   990
      Left            =   150
      TabIndex        =   5
      Top             =   930
      Width           =   5535
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "  重打结算单将生成新的结算单编号，以便与打印机的当前结算单编号一致，请在结算单打印错误时才使用此功能，正常时请勿使用。"
         Height          =   555
         Left            =   900
         TabIndex        =   6
         Top             =   270
         Width           =   4470
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   225
         Picture         =   "frmRePrintSettleSheet.frx":0342
         Top             =   285
         Width           =   480
      End
   End
   Begin RTComctl3.FlatLabel lblCurFinSheetID 
      Height          =   285
      Left            =   4230
      TabIndex        =   7
      Top             =   2010
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OutnerStyle     =   2
      Caption         =   "01234556"
   End
   Begin VB.TextBox txtFinSheetID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1455
      TabIndex        =   1
      Text            =   "0000001"
      Top             =   2010
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "原结算单号(&N):"
      Height          =   180
      Left            =   165
      TabIndex        =   0
      Top             =   2085
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "当前结算单号:"
      Height          =   180
      Left            =   3015
      TabIndex        =   22
      Top             =   2055
      Width           =   1170
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入需要重打的结算单编号:"
      Height          =   180
      Left            =   270
      TabIndex        =   4
      Top             =   360
      Width           =   2430
   End
End
Attribute VB_Name = "frmRePrintSettleSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub RefreshClear()
    txtFinSheetID.Text = ""
'    lblCurFinSheetID.Caption = ""
    
    lblStatus.Caption = ""
    lblCount.Caption = ""
    lblNum.Caption = ""
    lblDate.Caption = ""
    
    lblTotalPrice.Caption = ""
    lblNeedSplitOut.Caption = ""
    lblOperator.Caption = ""
    
    cmdRePrint.Enabled = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRePrint_Click()
On Error GoTo here
    Dim szSettleSheetID As String
    Dim m_oSplit As New STSettle.Split
    
'    '先处理行包结算
'    frmRePrintLugSheet.ZOrder 0
'    frmRePrintLugSheet.Show vbModal
    
    Dim oSettleSheet As New SettleSheet
    oSettleSheet.Init g_oActiveUser
    oSettleSheet.Identify txtFinSheetID.Text
    frmPrintFinSheet.m_szLugSettleSheetID = oSettleSheet.LuggageSettleIDs
'    frmPrintFinSheet.m_szProtocol = oSettleSheet.LuggageProtocolName
'    frmPrintFinSheet.m_dbTotalPrice = oSettleSheet.LuggageTotalBaseCarriage
'    frmPrintFinSheet.m_dbNeedSplitPrice = oSettleSheet.LuggageTotalSettlePrice
'
    
    
'    frmPrintFinSheet.m_bNeedPrint = True
    
    frmPrintFinSheet.m_SheetID = Trim(lblCurFinSheetID.Caption)
    frmPrintFinSheet.m_OldSheetID = Trim(txtFinSheetID.Text)
    frmPrintFinSheet.m_bRePrint = True
    frmPrintFinSheet.ZOrder 0
    frmPrintFinSheet.Show vbModal

    '自动生成结算单号 YYYYMM0001格式
    m_oSplit.Init g_oActiveUser
    szSettleSheetID = m_oSplit.GetLastSettleSheetID
    If szSettleSheetID = "0" Then
        lblCurFinSheetID.Caption = CStr(Year(Now)) + CStr(Month(Now)) + "0001"
    Else
        lblCurFinSheetID.Caption = szSettleSheetID
    End If

    RefreshClear

    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub Form_Load()
    Dim szSettleSheetID As String
    Dim m_oSplit As New STSettle.Split
    AlignFormPos Me
    RefreshClear

    
    '当前结算单号
    '自动生成结算单号 YYYYMM0001格式
     m_oSplit.Init g_oActiveUser
     szSettleSheetID = m_oSplit.GetLastSettleSheetID
     If szSettleSheetID = "0" Then
      lblCurFinSheetID.Caption = CStr(Year(Now)) + CStr(Month(Now)) + "0001"
     Else
      lblCurFinSheetID.Caption = szSettleSheetID
     End If
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    
End Sub

Private Sub txtFinSheetID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If RefreshInfo Then
            SendKeys "{TAB}"
        End If
    End If
End Sub

Private Function RefreshInfo() As Boolean
    On Error GoTo err
    Dim m_oSettleSheet As New SettleSheet
    
    m_oSettleSheet.Init g_oActiveUser
    m_oSettleSheet.Identify Trim(txtFinSheetID.Text)
    
    
    If m_oSettleSheet.Status = ESettleSheetStatus.CS_SettleSheetInvalid Then
        lblStatus.ForeColor = vbRed
        lblStatus.Caption = "作废"    '转换
        cmdRePrint.Enabled = False
    Else
        lblStatus.BackColor = 0
        lblStatus.Caption = "已结"
        cmdRePrint.Enabled = True
    End If
    lblCount.Caption = m_oSettleSheet.CheckSheetCount
    lblNum.Caption = m_oSettleSheet.TotalQuantity
    lblDate.Caption = Format(m_oSettleSheet.SettleStartDate, "YYYY年MM月DD日") & "—" & Format(m_oSettleSheet.SettleEndDate, "YYYY年MM月DD日")
    
    lblTotalPrice.Caption = m_oSettleSheet.TotalTicketPrice
    lblNeedSplitOut.Caption = m_oSettleSheet.SettleLocalCompanyPrice ' m_oSettleSheet.SettleOtherCompanyPrice - m_oSettleSheet.SettleStationPrice
    lblOperator.Caption = m_oSettleSheet.Settler
    
    RefreshInfo = True
    Exit Function
err:
    ShowErrorMsg
    RefreshInfo = False
End Function
