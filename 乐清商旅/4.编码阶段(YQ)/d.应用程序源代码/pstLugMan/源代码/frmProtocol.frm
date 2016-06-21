VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmProtocol 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "拆算协议"
   ClientHeight    =   6870
   ClientLeft      =   2295
   ClientTop       =   1530
   ClientWidth     =   9165
   HelpContextID   =   7000230
   Icon            =   "frmProtocol.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   0
      TabIndex        =   22
      Top             =   690
      Width           =   9240
   End
   Begin RTComctl3.CoolButton cmdVehicleSet 
      Height          =   345
      Left            =   4155
      TabIndex        =   19
      Top             =   6360
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "车辆设定(&V)"
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
      MICON           =   "frmProtocol.frx":014A
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
      Height          =   345
      Left            =   7830
      TabIndex        =   20
      Top             =   6360
      Width           =   1140
      _ExtentX        =   2011
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
      MICON           =   "frmProtocol.frx":0166
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
      Height          =   345
      Left            =   6615
      TabIndex        =   21
      ToolTipText     =   "保存协议"
      Top             =   6360
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "保存(&S)"
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
      MICON           =   "frmProtocol.frx":0182
      PICN            =   "frmProtocol.frx":019E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer tmStart 
      Interval        =   50
      Left            =   3600
      Top             =   240
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "协议公式"
      ForeColor       =   &H00FF0000&
      Height          =   4215
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   8700
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   60
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   8190
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   1695
         Left            =   240
         ScaleHeight     =   1635
         ScaleWidth      =   8025
         TabIndex        =   18
         Top             =   270
         Width           =   8085
         Begin VSFlex7LCtl.VSFlexGrid VSLuggageItem 
            Height          =   1335
            Left            =   120
            TabIndex        =   0
            Top             =   120
            Width           =   7710
            _cx             =   13600
            _cy             =   2355
            _ConvInfo       =   -1
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   16777215
            BackColorAlternate=   -2147483643
            GridColor       =   14737632
            GridColorFixed  =   -2147483639
            TreeColor       =   -2147483639
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   5
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   -1  'True
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
         End
      End
      Begin VB.Frame fraItem 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1935
         Left            =   120
         TabIndex        =   10
         Top             =   2145
         Width           =   8055
         Begin VB.TextBox txtChargeNumber 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1920
            TabIndex        =   2
            ToolTipText     =   "协议名称"
            Top             =   120
            Width           =   2340
         End
         Begin VB.ComboBox cboAcceptType 
            Height          =   300
            ItemData        =   "frmProtocol.frx":0538
            Left            =   5640
            List            =   "frmProtocol.frx":053A
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   570
            Width           =   2340
         End
         Begin VB.TextBox txtChargeName 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   5640
            TabIndex        =   3
            ToolTipText     =   "协议名称"
            Top             =   120
            Width           =   2340
         End
         Begin FText.asFlatMemo txtFormulaText 
            Height          =   495
            Left            =   1935
            TabIndex        =   7
            Top             =   1440
            Width           =   6075
            _ExtentX        =   10716
            _ExtentY        =   873
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
            ButtonHotForeColor=   -2147483628
            ButtonHotBackColor=   -2147483632
         End
         Begin FText.asFlatTextBox txtFormulaName 
            Height          =   300
            Left            =   1935
            TabIndex        =   5
            Top             =   990
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonHotBackColor=   -2147483633
            ButtonPressedBackColor=   -2147483627
            Text            =   ""
            ButtonBackColor =   -2147483633
            ButtonVisible   =   -1  'True
         End
         Begin VB.TextBox txtChargeMoney 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   5640
            TabIndex        =   6
            Top             =   990
            Width           =   2340
         End
         Begin VB.ComboBox cboChargeType 
            Height          =   300
            ItemData        =   "frmProtocol.frx":053C
            Left            =   1935
            List            =   "frmProtocol.frx":053E
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   600
            Width           =   2340
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "费用代号(&I):"
            Height          =   180
            Left            =   840
            TabIndex        =   36
            Top             =   240
            Width           =   1080
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   90
            Picture         =   "frmProtocol.frx":0540
            Top             =   135
            Width           =   480
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "托运方式(&T):"
            Height          =   180
            Left            =   4440
            TabIndex        =   34
            Top             =   660
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "固定款项(G):"
            Height          =   180
            Left            =   4440
            TabIndex        =   15
            Top             =   1050
            Width           =   1080
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "公式描述(&A):"
            Height          =   180
            Left            =   825
            TabIndex        =   14
            Top             =   1440
            Width           =   1080
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "公式(&F):"
            Height          =   180
            Left            =   855
            TabIndex        =   13
            Top             =   1050
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "费用名称(&I):"
            Height          =   180
            Left            =   4440
            TabIndex        =   12
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "费用类型(&T):"
            Height          =   180
            Left            =   825
            TabIndex        =   11
            Top             =   660
            Width           =   1080
         End
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   9165
      TabIndex        =   23
      Top             =   0
      Width           =   9165
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆算协议信息:"
         Height          =   180
         Left            =   480
         TabIndex        =   24
         Top             =   270
         Width           =   1170
      End
   End
   Begin RTComctl3.CoolButton cmdAdd 
      Height          =   345
      Left            =   5400
      TabIndex        =   35
      ToolTipText     =   "保存协议"
      Top             =   6360
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "新增(&A)"
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
      MICON           =   "frmProtocol.frx":084A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   930
      Left            =   -120
      TabIndex        =   17
      Top             =   6060
      Width           =   9495
      Begin RTComctl3.CoolButton cmdHelp 
         Height          =   315
         Left            =   390
         TabIndex        =   37
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "帮助(&H)"
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
         MICON           =   "frmProtocol.frx":0866
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
   Begin VB.Label lbRemark 
      BackColor       =   &H00E0E0E0&
      Caption         =   "该协议是常年有效"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   32
      Top             =   1320
      Width           =   6735
   End
   Begin VB.Label lbDefault 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "是"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   7440
      TabIndex        =   31
      Top             =   960
      Width           =   225
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "是否默认(&D):"
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
      Left            =   6120
      TabIndex        =   30
      Top             =   960
      Width           =   1260
   End
   Begin VB.Label lbProtocolName 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "55分折算协议"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3360
      TabIndex        =   29
      Top             =   960
      Width           =   2565
   End
   Begin VB.Label lbProtocol 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "0001"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1320
      TabIndex        =   28
      Top             =   960
      Width           =   480
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "承运车辆:"
      Height          =   180
      Index           =   2
      Left            =   30
      TabIndex        =   26
      Top             =   0
      Width           =   810
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "浙D00001"
      Height          =   180
      Left            =   840
      TabIndex        =   25
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "备  注(&R):"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   16
      Top             =   1320
      Width           =   1080
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "协议号(&P):"
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
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "协议名称(&N):"
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
      Left            =   2040
      TabIndex        =   8
      Top             =   960
      Width           =   1260
   End
End
Attribute VB_Name = "frmProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3885ABE70122"
Option Explicit
Public m_eStatus As eFormStatus
Public m_protocolID As String
Dim mChargeItem() As TProtocolChargeItemEX

Private Sub cboAcceptType_Click()
   VSLuggageItem.TextMatrix(VSLuggageItem.Row, 4) = cboAcceptType.Text
End Sub

Private Sub cboChargeType_Click()
  If Trim(cboChargeType.Text) = szConstType Then
     txtFormulaName.Enabled = False
     txtFormulaText.Enabled = False
     txtChargeMoney.Enabled = True
  Else
     txtFormulaName.Enabled = True
     txtFormulaText.Enabled = True
     txtChargeMoney.Enabled = False
  End If
    VSLuggageItem.TextMatrix(VSLuggageItem.Row, 3) = cboChargeType.Text
End Sub
Private Sub CmdAdd_Click()
    txtChargeNumber.Enabled = True
    txtChargeNumber.SetFocus
    VSLuggageItem.Select VSLuggageItem.Rows - 1, 1
    VSLuggageItem.AddItem ""
     If SaveErr = True Then
       Exit Sub
     End If
    FillVSLuggageItem
    VSLuggageItem.Select VSLuggageItem.Rows - 1, 1
    clear

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo ErrHandle
Dim i As Integer
Dim j As Integer
Dim nCount As Integer
Dim nType As String
If SaveErr = True Then
       Exit Sub
End If
'不充许有相同的托运类型
nCount = VSLuggageItem.Rows
If nCount = 0 Then Exit Sub
nType = VSLuggageItem.TextMatrix(1, 4)
j = 1
For i = 1 To nCount
    If j < nCount Then
        If VSLuggageItem.TextMatrix(j, 4) <> nType Then
            MsgBox "设置协议时,托运类型应唯一.", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    j = j + 1
Next i

If MsgBox("是否保存该协议的费用项吗?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub
       GetChargeItem
       m_oProtocol.Init m_oAUser
       m_oProtocol.SetChargeItem mChargeItem
Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub
'得到协议信息
Private Sub GetChargeItem()
  Dim i As Integer
  Dim vsRow As Integer
  If VSLuggageItem.Rows <= 1 Then Exit Sub
  
    ReDim mChargeItem(1 To VSLuggageItem.Rows - 1)
      For i = 1 To VSLuggageItem.Rows - 1
      mChargeItem(i).ProtocolID = Trim(lbProtocol.Caption)
      mChargeItem(i).ProtocolName = Trim(lbProtocolName.Caption)
      mChargeItem(i).ChargeID = Trim(VSLuggageItem.TextMatrix(i, 1))
      mChargeItem(i).ChargeName = Trim(VSLuggageItem.TextMatrix(i, 2))
      mChargeItem(i).AcceptType = IIf(Trim(VSLuggageItem.TextMatrix(i, 4)) = szAcceptTypeGeneral, 0, 1)
      If VSLuggageItem.TextMatrix(i, 3) = szConstType Then
         mChargeItem(i).ChargeType = 0
         mChargeItem(i).FixCharge = VSLuggageItem.TextMatrix(i, 5)
         mChargeItem(i).FormulaName = ""
         mChargeItem(i).FormulaText = ""
      Else
         mChargeItem(i).ChargeType = 1
         mChargeItem(i).FormulaName = Trim(VSLuggageItem.TextMatrix(i, 6))
         mChargeItem(i).FixCharge = 0
         mChargeItem(i).FormulaText = Trim(VSLuggageItem.TextMatrix(i, 7))
      End If
 Next
End Sub
Private Sub cmdVehicleSet_Click()
 frmSetVehicleProtocol.txtProtocol.Text = Trim(lbProtocol.Caption)
 frmSetVehicleProtocol.lblProtocolName.Caption = Trim(lbProtocolName.Caption)
 frmSetVehicleProtocol.Show vbModal
End Sub


Private Sub Form_Load()
    AlignFormPos Me
    '费用类型
    With cboChargeType
       .AddItem szConstType
       .AddItem szCalType
    End With
    FillAcceptType
    FillProtocol
    FillProtocolHead
    
    VSLuggageItem.Select VSLuggageItem.Rows - 1, 1
    GetVsFlexGrid VSLuggageItem
    '显示是否为默认协议
'    Dim rsTemp As Recordset
'    Set rsTemp = m_oProtocol.GetAllProtocol
'    If rsTemp.RecordCount = 0 Then Exit Sub
'    if rsTemp!default_mark
Exit Sub
ErrHandle:
    m_eStatus = ST_AddObj
    ShowErrorMsg
End Sub
'填充VSLuggageItem
Private Sub FillProtocol()
On Error GoTo ErrHandle
    Dim i As Integer, j As Integer
    Dim nlen As Integer
    Dim szTemp() As TProtocolChargeItemEX
    m_oProtocol.Identify Trim(m_protocolID)
    szTemp = m_oProtocol.ListChargeItem
    nlen = ArrayLength(szTemp)
    If nlen > 0 Then
        ReDim szTemp(1 To nlen)
        szTemp = m_oProtocol.ListChargeItem
'        j = 1
'        For i = 1 To nlen
'            If GetLuggageTypeInt(cboAcceptType.Text) = szTemp(i).AcceptType Then
'                VSLuggageItem.TextMatrix(j, 0) = szTemp(i).ProtocolID
'                VSLuggageItem.TextMatrix(j, 1) = szTemp(i).ChargeID
'                VSLuggageItem.TextMatrix(j, 2) = szTemp(i).ChargeName
'                VSLuggageItem.TextMatrix(j, 4) = cboAcceptType.Text
'                If szTemp(i).ChargeType = 0 Then
'                    VSLuggageItem.TextMatrix(j, 3) = szConstType
'                    VSLuggageItem.TextMatrix(j, 5) = szTemp(i).FixCharge
'                    VSLuggageItem.TextMatrix(j, 6) = ""
'                    VSLuggageItem.TextMatrix(j, 7) = ""
'                Else
'                    VSLuggageItem.TextMatrix(j, 3) = szCalType
'                    VSLuggageItem.TextMatrix(j, 5) = 0
'                    VSLuggageItem.TextMatrix(j, 6) = szTemp(i).FormulaName
'                    VSLuggageItem.TextMatrix(j, 7) = szTemp(i).FormulaText
'                End If
'                VSLuggageItem.AddItem ""
'                j = j + 1
'            ElseIf GetLuggageTypeInt(cboAcceptType.Text) = szTemp(i).AcceptType Then
'                VSLuggageItem.TextMatrix(j, 0) = szTemp(i).ProtocolID
'                VSLuggageItem.TextMatrix(j, 1) = szTemp(i).ChargeID
'                VSLuggageItem.TextMatrix(j, 2) = szTemp(i).ChargeName
'                VSLuggageItem.TextMatrix(j, 4) = cboAcceptType.Text
'                If szTemp(i).ChargeType = 0 Then
'                    VSLuggageItem.TextMatrix(j, 3) = szConstType
'                    VSLuggageItem.TextMatrix(j, 5) = szTemp(i).FixCharge
'                    VSLuggageItem.TextMatrix(j, 6) = ""
'                    VSLuggageItem.TextMatrix(j, 7) = ""
'                Else
'                    VSLuggageItem.TextMatrix(j, 3) = szCalType
'                    VSLuggageItem.TextMatrix(j, 5) = 0
'                    VSLuggageItem.TextMatrix(j, 6) = szTemp(i).FormulaName
'                    VSLuggageItem.TextMatrix(j, 7) = szTemp(i).FormulaText
'                End If
'                VSLuggageItem.AddItem ""
'                j = j + 1
'            End If
'        Next i
'    End If
    For i = 1 To nlen
         VSLuggageItem.TextMatrix(i, 0) = szTemp(i).ProtocolID
        VSLuggageItem.TextMatrix(i, 1) = szTemp(i).ChargeID
        VSLuggageItem.TextMatrix(i, 2) = szTemp(i).ChargeName
        If szTemp(i).AcceptType = 0 Then
            VSLuggageItem.TextMatrix(i, 4) = szAcceptTypeGeneral
        Else
            VSLuggageItem.TextMatrix(i, 4) = szAcceptTypeMan
        End If
        If szTemp(i).ChargeType = 0 Then
            VSLuggageItem.TextMatrix(i, 3) = szConstType

            VSLuggageItem.TextMatrix(i, 5) = szTemp(i).FixCharge
            VSLuggageItem.TextMatrix(i, 6) = ""
            VSLuggageItem.TextMatrix(i, 7) = ""
        Else
            VSLuggageItem.TextMatrix(i, 3) = szCalType

            VSLuggageItem.TextMatrix(i, 5) = 0
            VSLuggageItem.TextMatrix(i, 6) = szTemp(i).FormulaName
            VSLuggageItem.TextMatrix(i, 7) = szTemp(i).FormulaText
        End If
        If i < nlen Then
            VSLuggageItem.AddItem ""
        End If
    Next i
    Else
    Exit Sub
    End If
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub
Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    Unload Me
End Sub

Private Sub txtChargeMoney_Change()
  FormatTextToNumeric txtChargeMoney, False, False
    VSLuggageItem.TextMatrix(VSLuggageItem.Row, 5) = txtChargeMoney
End Sub
Private Sub txtChargeName_Change()
   VSLuggageItem.TextMatrix(VSLuggageItem.Row, 2) = txtChargeName
End Sub

Private Sub txtChargeNumber_Change()
    VSLuggageItem.TextMatrix(VSLuggageItem.Row, 1) = txtChargeNumber.Text
    FormatTextToNumeric txtChargeMoney, False, False
    FormatTextBoxBySize txtChargeNumber, 4
End Sub

'打开公式列表
Private Sub txtFormulaName_ButtonClick()
    frmFormula.m_eStatus = ST_NormalObj
    frmFormula.Show vbModal
    txtFormulaName.Text = Trim(frmFormula.m_szFormulaName)
    txtFormulaText.Text = frmFormula.m_szFormula
End Sub
'填充vsluggageitem的值。
Private Sub FillVSLuggageItem()
   With VSLuggageItem
        
        .TextMatrix(.Rows - 1, 1) = Trim(txtChargeNumber.Text)
        .TextMatrix(.Rows - 1, 2) = Trim(txtChargeName.Text)
        .TextMatrix(.Rows - 1, 3) = Trim(cboChargeType.Text)
        .TextMatrix(.Rows - 1, 4) = Trim(cboAcceptType.Text)
       If Trim(cboChargeType.Text) = szConstType Then
          .TextMatrix(.Rows - 1, 5) = Trim(txtChargeMoney.Text)
          .TextMatrix(.Rows - 1, 6) = ""
          .TextMatrix(.Rows - 1, 7) = ""
       Else
          .TextMatrix(.Rows - 1, 5) = 0
          .TextMatrix(.Rows - 1, 6) = Trim(txtFormulaName.Text)
          .TextMatrix(.Rows - 1, 7) = Trim(txtFormulaText.Text)
       End If
    End With
End Sub


Public Sub clear()
  txtChargeMoney.Text = ""
  txtFormulaName.Text = ""
  txtFormulaText.Text = ""
  txtChargeName.Text = ""
  txtChargeNumber.Text = ""
  cboAcceptType.ListIndex = 0
  cboChargeType.ListIndex = 0
End Sub
'填充列头
Public Sub FillProtocolHead()
  VSLuggageItem.TextMatrix(0, 0) = "协议代号"
  VSLuggageItem.TextMatrix(0, 1) = "费用代号"
  VSLuggageItem.TextMatrix(0, 2) = "费用名称"
  VSLuggageItem.TextMatrix(0, 3) = "费用类型"
  VSLuggageItem.TextMatrix(0, 4) = "托运方式"
  VSLuggageItem.TextMatrix(0, 5) = "固定款项"
  VSLuggageItem.TextMatrix(0, 6) = "公式名称"
  VSLuggageItem.TextMatrix(0, 7) = "公式内容"
End Sub

Private Sub txtFormulaName_Change()
    VSLuggageItem.TextMatrix(VSLuggageItem.Row, 6) = txtFormulaName.Text
End Sub

Private Sub txtFormulaText_Change()
     VSLuggageItem.TextMatrix(VSLuggageItem.Row, 7) = txtFormulaText.Text
End Sub

Private Sub VSLuggageItem_Click()
  GetVsFlexGrid VSLuggageItem
End Sub
  '点击某行刷新信息
Private Function GetVsFlexGrid(VsObject As VSFlexGrid)
 Dim i As Integer
 With VsObject
     If .Text <> "" Then
     txtChargeNumber.Text = .TextMatrix(.Row, 1)
     txtChargeName.Text = .TextMatrix(.Row, 2)
'     cboChargeType.Text = .TextMatrix(.Row, 3)
     For i = 0 To cboChargeType.ListCount
        If cboChargeType.List(i) = .TextMatrix(.Row, 3) And .TextMatrix(.Row, 3) <> "" Then
            cboChargeType.ListIndex = i
        End If
     Next i
     txtFormulaName.Text = .TextMatrix(.Row, 6)
     txtChargeMoney.Text = .TextMatrix(.Row, 5)
     txtFormulaText.Text = .TextMatrix(.Row, 7)
     cboAcceptType.Text = .TextMatrix(.Row, 4)
     End If
 End With
 End Function
Private Sub FillAcceptType()
With cboAcceptType

   .AddItem GetLuggageTypeString(0)
   .AddItem GetLuggageTypeString(1)
   .ListIndex = 0
End With
End Sub
Public Function SaveErr(Optional bError As Boolean = False) As Boolean
  If txtChargeNumber.Text = "" Then
        MsgBox "费用代码必须填写，请重新输入费用代码！", vbError, "错误"
        txtChargeNumber.SetFocus
        bError = True
    
  Else
        If cboChargeType.Text = szCalType Then
            If txtFormulaName.Text = "" Then
               MsgBox "你选择用了公式计算，必须输入公式名称！", vbExclamation, "错误"
               txtFormulaName.SetFocus
               bError = True
             End If
        Else
             If txtChargeMoney.Text = "" Then
             MsgBox "你选择用了固定费用，必须输入固定款项！", vbExclamation, "错误"
             bError = True
             txtChargeMoney.SetFocus
             End If
        End If
    End If
   SaveErr = bError
End Function

Private Sub VSLuggageItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    If VSLuggageItem.Row = 1 Then
       MsgBox "不能删除固定行，你只能修改该费用项！", vbExclamation, "错误"
       Exit Sub
    End If
    VSLuggageItem.RemoveItem (VSLuggageItem.Row)
    GetVsFlexGrid VSLuggageItem
    
End If
End Sub
