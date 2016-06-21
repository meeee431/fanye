VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmRePrintSheet 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "重打单据"
   ClientHeight    =   5520
   ClientLeft      =   3105
   ClientTop       =   3405
   ClientWidth     =   6630
   HelpContextID   =   4000601
   Icon            =   "frmRePrintSheet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Tag             =   "Modal"
   Begin RTComctl3.FlatLabel lblCurSheetNo 
      Height          =   285
      Left            =   4530
      TabIndex        =   30
      Top             =   1950
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OutnerStyle     =   2
      Caption         =   "01234556"
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   -30
      ScaleHeight     =   795
      ScaleWidth      =   7185
      TabIndex        =   26
      Top             =   0
      Width           =   7185
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   0
         TabIndex        =   27
         Top             =   750
         Width           =   7215
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请输入需要重打的单据编号:"
         Height          =   180
         Left            =   120
         TabIndex        =   28
         Top             =   300
         Width           =   2250
      End
   End
   Begin VB.TextBox txtOriSheetNo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1785
      TabIndex        =   1
      Text            =   "0000001"
      Top             =   1950
      Width           =   1410
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "行包摘要信息"
      Height          =   2235
      Left            =   480
      TabIndex        =   9
      Top             =   2280
      Width           =   5535
      Begin VB.TextBox txtPacketID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1350
         TabIndex        =   3
         Text            =   "0000001"
         Top             =   210
         Width           =   1350
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "新自编号(&I):"
         Height          =   180
         Left            =   240
         TabIndex        =   2
         Top             =   270
         Width           =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   210
         X2              =   5250
         Y1              =   1695
         Y2              =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   210
         X2              =   5250
         Y1              =   1710
         Y2              =   1710
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "托运费总计:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2880
         TabIndex        =   34
         Top             =   1860
         Width           =   1155
      End
      Begin VB.Label lblCharge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4080
         TabIndex        =   33
         Top             =   1860
         Width           =   315
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "代收运费:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   32
         Top             =   1860
         Width           =   945
      End
      Begin VB.Label lblTransitCharge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1230
         TabIndex        =   31
         Top             =   1530
         Width           =   105
      End
      Begin VB.Label lblOperator 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "李地"
         Height          =   180
         Left            =   3720
         TabIndex        =   25
         Top             =   1410
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "受理人:"
         Height          =   180
         Left            =   2880
         TabIndex        =   24
         Top             =   1410
         Width           =   630
      End
      Begin VB.Label lblLoader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "001/啊啊啊啊"
         Height          =   180
         Left            =   1140
         TabIndex        =   23
         Top             =   1410
         Width           =   1080
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "装卸工:"
         Height          =   180
         Left            =   240
         TabIndex        =   22
         Top             =   1410
         Width           =   630
      End
      Begin VB.Label lblPackageNums 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "重庆成都线"
         Height          =   180
         Left            =   3720
         TabIndex        =   21
         Top             =   870
         Width           =   900
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包件数:"
         Height          =   180
         Left            =   2880
         TabIndex        =   20
         Top             =   870
         Width           =   810
      End
      Begin VB.Label lblPicker 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "绍兴汽运集团"
         Height          =   180
         Left            =   3720
         TabIndex        =   19
         Top             =   1140
         Width           =   1080
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收件人:"
         Height          =   180
         Left            =   2880
         TabIndex        =   18
         Top             =   1140
         Width           =   630
      End
      Begin VB.Label lblShipper 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "川A 1234567"
         Height          =   180
         Left            =   1140
         TabIndex        =   17
         Top             =   1140
         Width           =   990
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发件人:"
         Height          =   180
         Left            =   240
         TabIndex        =   16
         Top             =   1140
         Width           =   630
      End
      Begin VB.Label lblStartStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         Height          =   180
         Left            =   3720
         TabIndex        =   15
         Top             =   570
         Width           =   180
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起运站:"
         Height          =   180
         Left            =   2880
         TabIndex        =   14
         Top             =   570
         Width           =   630
      End
      Begin VB.Label lblArrivedDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1999-06-08"
         Height          =   180
         Left            =   1140
         TabIndex        =   13
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "到达时间:"
         Height          =   180
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   810
      End
      Begin VB.Label lblPackageName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1234"
         Height          =   180
         Left            =   1140
         TabIndex        =   11
         Top             =   570
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包名称:"
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   570
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "注意"
      Height          =   990
      Left            =   480
      TabIndex        =   6
      Top             =   870
      Width           =   5535
      Begin VB.Image Image1 
         Height          =   480
         Left            =   225
         Picture         =   "frmRePrintSheet.frx":0A02
         Top             =   285
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "  重打单据将生成新的路单编号，以便与打印机的当前路单编号一致，请在路单打印错误时才使用此功能，正常时请勿使用。"
         Height          =   555
         Left            =   960
         TabIndex        =   7
         Top             =   270
         Width           =   4470
      End
   End
   Begin RTComctl3.CoolButton cmdExit 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   4830
      TabIndex        =   5
      Top             =   4950
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
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
      MICON           =   "frmRePrintSheet.frx":12CC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdRePrintSheet 
      Height          =   345
      Left            =   3180
      TabIndex        =   4
      Top             =   4950
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "重新打印(&P)"
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
      MICON           =   "frmRePrintSheet.frx":12E8
      PICN            =   "frmRePrintSheet.frx":1304
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
      Height          =   990
      Left            =   -90
      TabIndex        =   29
      Top             =   4680
      Width           =   8745
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "当前单据编号:"
      Height          =   180
      Left            =   3315
      TabIndex        =   8
      Top             =   1995
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "原单据编号(&N):"
      Height          =   180
      Left            =   495
      TabIndex        =   0
      Top             =   1995
      Width           =   1260
   End
End
Attribute VB_Name = "frmRePrintSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const msgTitle = "重打单据"
Dim m_rsSheetInfo As Recordset

Private Function ValidateSheetId() As Boolean
    If Trim(txtOriSheetNo.Text) <> "" Then
        If Trim(txtOriSheetNo.Text) <> lblCurSheetNo.Caption Then
            Set m_rsSheetInfo = g_oPackageSvr.ListPackageRS("sheet_id='" & Trim(txtOriSheetNo.Text) & "'")
            If m_rsSheetInfo.RecordCount = 0 Then
                MsgBox "该单据不存在！", vbCritical, "错误"
                ValidateSheetId = False
            Else
                writeSheetInfo
                ValidateSheetId = True
            End If
        End If
    End If
    If Val(Trim(txtPacketID.Text)) = 0 Then
        MsgBox "请输入新的自编号！", vbExclamation, "错误"
        ValidateSheetId = False
    Else
        ValidateSheetId = True
    End If
End Function

Private Sub writeSheetInfo()
    lblPackageName.Caption = FormatDbValue(m_rsSheetInfo!package_name)
    lblStartStation.Caption = FormatDbValue(m_rsSheetInfo!area_type) & " " & FormatDbValue(m_rsSheetInfo!start_station_name)
    lblArrivedDate.Caption = Format(FormatDbValue(m_rsSheetInfo!arrive_time), "YYYY-MM-DD hh:mm")
    lblPackageNums.Caption = FormatDbValue(m_rsSheetInfo!package_number)
    lblShipper.Caption = FormatDbValue(m_rsSheetInfo!send_name)
    lblPicker.Caption = FormatDbValue(m_rsSheetInfo!Picker)
    lblLoader.Caption = FormatDbValue(m_rsSheetInfo!Loader)
    lblOperator.Caption = FormatDbValue(m_rsSheetInfo!Operator)
    lblTransitCharge.Caption = FormatDbValue(m_rsSheetInfo!transit_charge)
    lblCharge.Caption = FormatDbValue(m_rsSheetInfo!load_charge) + FormatDbValue(m_rsSheetInfo!keep_charge) + FormatDbValue(m_rsSheetInfo!move_charge) + FormatDbValue(m_rsSheetInfo!send_charge) + FormatDbValue(m_rsSheetInfo!other_charge)
'    txtPacketID.Text = ""
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub
Private Sub InitForm()
    lblPackageName.Caption = ""
    lblStartStation.Caption = ""
    lblArrivedDate.Caption = ""
    lblPackageNums.Caption = ""
    lblShipper.Caption = ""
    lblPicker.Caption = ""
    lblLoader.Caption = ""
    lblOperator.Caption = ""
    lblTransitCharge.Caption = ""
    lblCharge.Caption = ""

End Sub




Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub cmdRePrintSheet_Click()
    On Error GoTo ErrHandle
    
    
    If Not ValidateSheetId Then
        txtOriSheetNo.SetFocus
        Exit Sub
    End If
    If txtOriSheetNo.Text = "" Then
        MsgBox "请输入需重打的单据号", vbInformation, "重打单据"
        txtOriSheetNo.SetFocus
        Exit Sub
    End If
    If lblCurSheetNo.Caption <> txtOriSheetNo.Text Then
        If MsgBox("* 重打单据将作废原单据!" & vbCrLf & "* 请确定重打的单据编号同打印机上的单据号一致，继续吗?", vbExclamation + vbYesNo, msgTitle) = vbYes Then
            SetBusy
            Dim lPackageID As Long
            lPackageID = g_oPackageSvr.ChangeSheetID(txtOriSheetNo.Text, lblCurSheetNo.Caption, BuildPacketID(Val(txtPacketID.Text)))
            
            Dim oPackage As New Package
            oPackage.init g_oActUser
            oPackage.Identify lPackageID
                                    
            'Print
            'PrintAcceptSheet oPackage
            frmSheet.PrintSheetReport oPackage
                        
            SetNormal
            IncSheetID 1, True
            MsgBox "重打单据成功！", vbInformation, msgTitle
            
        End If
        '如果单据号一致,则不重新生成单据
        '不一致，则首先作废单据，然后重新生成单据
        '调用单据显示窗体
        Unload Me
    End If
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
    SetNormal
    txtOriSheetNo.SetFocus
End Sub

Private Sub Form_Activate()
    lblCurSheetNo.Caption = g_szSheetID
End Sub

Private Sub Form_Load()
    lblCurSheetNo.Caption = g_szSheetID
    lblArrivedDate.Caption = ""
    lblCharge.Caption = ""
    lblLoader.Caption = ""
    lblOperator.Caption = ""
    lblPackageName.Caption = ""
    lblPackageNums.Caption = ""
    lblShipper.Caption = ""
    lblStartStation.Caption = ""
    txtOriSheetNo.Text = ""
    lblPicker.Caption = ""
    
    
End Sub




Private Sub Form_Resize()
    txtOriSheetNo.SetFocus
End Sub




Private Sub txtOriSheetNo_Change()
    If txtOriSheetNo.Text <> "" Then
        If Len(txtOriSheetNo.Text) > 10 Then
            txtOriSheetNo.Text = Left(txtOriSheetNo.Text, 10)
        End If
    End If
End Sub

Private Sub txtOriSheetNo_GotFocus()
    txtOriSheetNo.SelStart = 0
    txtOriSheetNo.SelLength = Len(txtOriSheetNo.Text)
End Sub

Private Sub txtOriSheetNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not ValidateSheetId Then
            txtOriSheetNo.SetFocus
        Else
            cmdRePrintSheet.SetFocus
        End If
    End If
    If KeyAscii = 32 Then
        KeyAscii = 0
    End If
End Sub

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

Private Sub txtPacketID_Change()
    FormatTextToNumeric txtPacketID, False, False

End Sub

Private Sub txtPacketID_GotFocus()
With txtPacketID
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub
