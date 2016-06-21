VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmRePrintFinSheet 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "重打结算单"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "frmRePrintFinSheet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   5910
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   990
      Left            =   -60
      TabIndex        =   18
      Top             =   4110
      Width           =   8745
      Begin RTComctl3.CoolButton cmdExit 
         Height          =   375
         Left            =   4350
         TabIndex        =   22
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
         MICON           =   "frmRePrintFinSheet.frx":030A
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
         Left            =   2730
         TabIndex        =   21
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
         MICON           =   "frmRePrintFinSheet.frx":0326
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
      Left            =   210
      TabIndex        =   5
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
         Caption         =   "结算对象:"
         Height          =   210
         Left            =   270
         TabIndex        =   17
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblObjectType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车辆"
         Height          =   180
         Left            =   1290
         TabIndex        =   16
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算月份:"
         Height          =   180
         Left            =   270
         TabIndex        =   15
         Top             =   900
         Width           =   810
      End
      Begin VB.Label lblMonth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2003-3"
         Height          =   180
         Left            =   1260
         TabIndex        =   14
         Top             =   900
         Width           =   540
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆出金额:"
         Height          =   210
         Left            =   2820
         TabIndex        =   13
         Top             =   570
         Width           =   810
      End
      Begin VB.Label lblNeedSplitOut 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1255"
         Height          =   210
         Left            =   3900
         TabIndex        =   12
         Top             =   540
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包运费:"
         Height          =   210
         Left            =   2820
         TabIndex        =   11
         Top             =   270
         Width           =   810
      End
      Begin VB.Label lblTotalPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4568"
         Height          =   210
         Left            =   3930
         TabIndex        =   10
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算期限:"
         Height          =   180
         Left            =   270
         TabIndex        =   9
         Top             =   1200
         Width           =   810
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2003年-3月-17日--2003年-03月-31日"
         Height          =   180
         Left            =   1200
         TabIndex        =   8
         Top             =   1200
         Width           =   2970
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算人:"
         Height          =   180
         Left            =   2970
         TabIndex        =   7
         Top             =   900
         Width           =   630
      End
      Begin VB.Label lblOperator 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "王建波"
         Height          =   180
         Left            =   3840
         TabIndex        =   6
         Top             =   870
         Width           =   540
      End
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
      Left            =   1515
      TabIndex        =   4
      Text            =   "0000001"
      Top             =   2010
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "注意"
      Height          =   990
      Left            =   210
      TabIndex        =   1
      Top             =   930
      Width           =   5535
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "  重打结算单将生成新的结算单编号，以便与打印机的当前结算单编号一致，请在结算单打印错误时才使用此功能，正常时请勿使用。"
         Height          =   555
         Left            =   900
         TabIndex        =   2
         Top             =   270
         Width           =   4470
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   225
         Picture         =   "frmRePrintFinSheet.frx":0342
         Top             =   285
         Width           =   480
      End
   End
   Begin RTComctl3.FlatLabel lblCurFinSheetID 
      Height          =   285
      Left            =   4290
      TabIndex        =   3
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
      BackColor       =   -2147483633
      OutnerStyle     =   2
      Caption         =   "01234556"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "原结算单号(&N):"
      Height          =   180
      Left            =   225
      TabIndex        =   20
      Top             =   2085
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "当前结算单号:"
      Height          =   180
      Left            =   3075
      TabIndex        =   19
      Top             =   2055
      Width           =   1170
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入需要重打的结算单编号:"
      Height          =   180
      Left            =   270
      TabIndex        =   0
      Top             =   360
      Width           =   2430
   End
End
Attribute VB_Name = "frmRePrintFinSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub RefreshClear()
    txtFinSheetID.Text = ""
'    lblCurFinSheetID.Caption = ""
    
    lblStatus.Caption = ""
    lblObjectType.Caption = ""
    lblMonth.Caption = ""
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
    Dim rsTemp As Recordset
    frmPrintFinSheet.SheetID = Trim(lblCurFinSheetID.Caption)
    frmPrintFinSheet.OldSheetID = Trim(txtFinSheetID.Text)
    frmPrintFinSheet.mRePrint = True
    frmPrintFinSheet.ZOrder 0
    frmPrintFinSheet.Show vbModal
    
    '自动生成结算单号 YYYYMM0001格式
    m_oLugFinSvr.Init m_oAUser
    Set rsTemp = m_oLugFinSvr.GetFinSheetID
    If rsTemp.RecordCount = 0 Then
        lblCurFinSheetID.Caption = CStr(Year(Now)) + CStr(Month(Now)) + "0001"
    Else
        lblCurFinSheetID.Caption = CStr(rsTemp!fin_sheet_id + 1)
    End If
    
    RefreshClear
    
    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub Form_Load()
    Dim rsTemp As Recordset
    AlignFormPos Me
    RefreshClear

    
    '当前结算单号
    '自动生成结算单号 YYYYMM0001格式
     m_oLugFinSvr.Init m_oAUser
     Set rsTemp = m_oLugFinSvr.GetFinSheetID
     If rsTemp.RecordCount = 0 Then
      lblCurFinSheetID.Caption = CStr(Year(Now)) + CStr(Month(Now)) + "0001"
     Else
      lblCurFinSheetID.Caption = CStr(rsTemp!fin_sheet_id + 1)
     End If
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    
End Sub

Private Sub txtFinSheetID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RefreshInfo
    End If
End Sub

Private Sub RefreshInfo()
On Error GoTo err
    Dim rsTemp As Recordset
    
    Set rsTemp = m_oLugFinSvr.GetFinSheetInfo(Trim(txtFinSheetID.Text))
    If rsTemp.RecordCount = 0 Then
        MsgBox "此结算单不存在!", vbInformation, Me.Caption
        RefreshClear
    Else
        
        
        If FormatDbValue(rsTemp!Status) = ELuggageSettleValidMark.LuggageNotValid Then
            lblStatus.ForeColor = vbRed
            lblStatus.Caption = "作废"
        Else
            lblStatus.BackColor = 0
            lblStatus.Caption = "已结"
        End If
        lblObjectType.Caption = FormatDbValue(rsTemp!split_object_name)
        lblMonth.Caption = Format(FormatDbValue(rsTemp!settle_month), "YYYY年MM月")
        lblDate.Caption = Format(FormatDbValue(rsTemp!settlement_start_time), "YYYY年MM月DD日") & "--" & Format(FormatDbValue(rsTemp!settlement_end_time), "YYYY年MM月DD日")
        
        lblTotalPrice.Caption = FormatDbValue(rsTemp!total_price)
        lblNeedSplitOut.Caption = FormatDbValue(rsTemp!need_split_out)
        lblOperator.Caption = FormatDbValue(rsTemp!Operator)
        
        cmdRePrint.Enabled = True
    End If
    
Exit Sub
err:
    ShowErrorMsg
End Sub
