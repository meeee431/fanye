VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmCancelAccept 
   BackColor       =   &H8000000C&
   Caption         =   "��������"
   ClientHeight    =   5985
   ClientLeft      =   3630
   ClientTop       =   2400
   ClientWidth     =   7785
   HelpContextID   =   7000040
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   7785
   WindowState     =   2  'Maximized
   Begin VB.Frame fraOutLine 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5445
      Left            =   330
      TabIndex        =   0
      Top             =   300
      Width           =   6945
      Begin VB.Frame fraTktInfoChange 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�а�Ʊ��Ϣ"
         Height          =   4545
         Left            =   210
         TabIndex        =   2
         Top             =   780
         Width           =   6570
         Begin VB.Label lblLuggageName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1410
            TabIndex        =   36
            Top             =   360
            Width           =   120
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�а�����:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   210
            TabIndex        =   35
            Top             =   330
            Width           =   1080
         End
         Begin VB.Line Line2 
            X1              =   180
            X2              =   3030
            Y1              =   4320
            Y2              =   4320
         End
         Begin VB.Line Line1 
            X1              =   180
            X2              =   3030
            Y1              =   3840
            Y2              =   3840
         End
         Begin VB.Label lblStatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4410
            TabIndex        =   32
            Top             =   1605
            Width           =   120
         End
         Begin VB.Label lblTicketPrice 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1080
            TabIndex        =   31
            Top             =   3900
            Width           =   1425
         End
         Begin VB.Label lblOperationTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4410
            TabIndex        =   30
            Top             =   2895
            Width           =   120
         End
         Begin VB.Label lblOperater 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1080
            TabIndex        =   29
            Top             =   2895
            Width           =   120
         End
         Begin VB.Label lblStartStation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1080
            TabIndex        =   28
            Top             =   750
            Width           =   120
         End
         Begin VB.Label lblEndStation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4410
            TabIndex        =   27
            Top             =   750
            Width           =   120
         End
         Begin VB.Label lblOperatorChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����Ա:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   210
            TabIndex        =   26
            Top             =   2895
            Width           =   840
         End
         Begin VB.Label lblStateChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "״̬:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3300
            TabIndex        =   25
            Top             =   1605
            Width           =   600
         End
         Begin VB.Label label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��վ:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3300
            TabIndex        =   24
            Top             =   750
            Width           =   600
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���վ:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   210
            TabIndex        =   23
            Top             =   750
            Width           =   840
         End
         Begin VB.Label lblTimeChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3300
            TabIndex        =   22
            Top             =   2895
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ʊ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   210
            TabIndex        =   21
            Top             =   3960
            Width           =   600
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   210
            TabIndex        =   20
            Top             =   2040
            Width           =   600
         End
         Begin VB.Label lblCalWeight 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1080
            TabIndex        =   19
            Top             =   2040
            Width           =   120
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   210
            TabIndex        =   18
            Top             =   3330
            Width           =   840
         End
         Begin VB.Label lblShipper 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1080
            TabIndex        =   17
            Top             =   3330
            Width           =   120
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ȡ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3300
            TabIndex        =   16
            Top             =   3330
            Width           =   840
         End
         Begin VB.Label lblPicker 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4410
            TabIndex        =   15
            Tag             =   "�����"
            Top             =   3330
            Width           =   120
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   210
            TabIndex        =   14
            Top             =   1185
            Width           =   600
         End
         Begin VB.Label lblMileage 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1080
            TabIndex        =   13
            Top             =   1185
            Width           =   120
         End
         Begin VB.Label lblActWeight 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4410
            TabIndex        =   12
            Top             =   2040
            Width           =   120
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ʵ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3300
            TabIndex        =   11
            Top             =   2040
            Width           =   600
         End
         Begin VB.Label lblLabelID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1080
            TabIndex        =   10
            Top             =   1605
            Width           =   120
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ǩ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   210
            TabIndex        =   9
            Top             =   1605
            Width           =   840
         End
         Begin VB.Label lblBagNumber 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1080
            TabIndex        =   8
            Top             =   2475
            Width           =   120
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   210
            TabIndex        =   7
            Top             =   2475
            Width           =   600
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���˷�ʽ:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3300
            TabIndex        =   6
            Top             =   1185
            Width           =   1080
         End
         Begin VB.Label lblAcceptType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4410
            TabIndex        =   5
            Top             =   1185
            Width           =   120
         End
         Begin VB.Label lblOverNumber 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4410
            TabIndex        =   4
            Top             =   2475
            Width           =   120
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ؼ���:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3300
            TabIndex        =   3
            Top             =   2475
            Width           =   1080
         End
      End
      Begin VB.TextBox txtLuggageID 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1830
         MaxLength       =   10
         TabIndex        =   1
         Top             =   225
         Width           =   2490
      End
      Begin RTComctl3.CoolButton cmdCancelAccept 
         Height          =   525
         Left            =   4650
         TabIndex        =   34
         Top             =   210
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "��Ʊ(&C)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         MICON           =   "frmCancelAccept.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblOldTktNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���а�����(&N):"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   33
         Top             =   300
         Width           =   1470
      End
   End
End
Attribute VB_Name = "frmCancelAccept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelAccept_Click()
On Error GoTo ErrHandle
Dim sAnswer
  sAnswer = MsgBox("  ��ȷʵҪ����Щ����?", vbInformation + vbYesNo, "��������")
  If sAnswer = vbYes Then
     moLugSvr.CancelAcceptSheet (Trim(txtLuggageID.Text))
  End If
  lblStatus.ForeColor = vbRed
  lblStatus.Caption = "�ѷ�"
  cmdCancelAccept.Enabled = False
Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
       txtLuggageID.Text = ""
       txtLuggageID.SetFocus
       FormClear
    End If
    If KeyAscii = vbKeyF1 Then
        DisplayHelp Me
    End If
End Sub
Private Sub FormClear()
        lblLuggageName.Caption = ""
        lblStartStation.Caption = ""
        lblEndStation.Caption = ""
        lblMileage.Caption = ""
        lblAcceptType.Caption = ""
        lblLabelID.Caption = ""
        lblStatus.Caption = ""
        lblCalWeight.Caption = ""
        lblActWeight.Caption = ""
        lblBagNumber.Caption = ""
        lblOverNumber.Caption = ""
        lblOperater.Caption = ""
        lblOperationTime.Caption = ""
        lblShipper.Caption = ""
        lblPicker.Caption = ""
        lblTicketPrice.Caption = ""
        cmdCancelAccept.Enabled = False
End Sub
Private Sub Form_Load()
 AlignFormPos Me
 FormClear
End Sub

Private Sub Form_Resize()
    If mdiMain.ActiveForm Is Me Then
        If Not Me.WindowState = vbMaximized Then Me.WindowState = vbMaximized
        fraOutLine.Left = (Me.ScaleWidth - fraOutLine.Width) / 2
        fraOutLine.Top = (Me.ScaleHeight - fraOutLine.Height) / 2
    End If
End Sub

Private Sub Form_Activate()
    SetSheetNoLabel True, g_szAcceptSheetID
    txtLuggageID.Text = ""
    FormClear
End Sub

Private Sub Form_Deactivate()
    HideSheetNoLabel
End Sub
Private Sub Form_Unload(Cancel As Integer)
    HideSheetNoLabel
End Sub

Private Sub txtLuggageID_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrHandle
    If KeyCode = vbKeyReturn And txtLuggageID.Text <> "" Then
        moAcceptSheet.Identify Trim(txtLuggageID.Text)
        
        lblLuggageName.Caption = moAcceptSheet.LuggageName
        lblStartStation.Caption = moAcceptSheet.StartStationName
        lblEndStation.Caption = moAcceptSheet.DesStationName
        lblMileage.Caption = moAcceptSheet.Mileage
        lblAcceptType.Caption = moAcceptSheet.AcceptType
        lblLabelID.Caption = moAcceptSheet.StartLabelID
        If moAcceptSheet.Status <> 0 Then
        lblStatus.ForeColor = vbRed
        End If
        lblStatus.Caption = moAcceptSheet.StatusString
        lblCalWeight.Caption = moAcceptSheet.CalWeight
        lblActWeight.Caption = moAcceptSheet.ActWeight
        lblBagNumber.Caption = moAcceptSheet.Number
        lblOverNumber.Caption = moAcceptSheet.OverNumber
        lblOperater.Caption = moAcceptSheet.Operator
        lblOperationTime.Caption = CStr(moAcceptSheet.OperateTime)
        lblShipper.Caption = moAcceptSheet.Shipper
        lblPicker.Caption = moAcceptSheet.Picker
        lblTicketPrice.Caption = moAcceptSheet.TotalPrice
        If moAcceptSheet.Status = 0 Then
            cmdCancelAccept.Enabled = True
            cmdCancelAccept.SetFocus
        Else
            If moAcceptSheet.Status = 1 Then
                
                'MsgBox "�а����ѷϣ����ܽ��д˲�����"
                cmdCancelAccept.Enabled = False
            ElseIf moAcceptSheet.Status = 2 Then
                'MsgBox "�а������ˣ����ܽ��д˲���"
                cmdCancelAccept.Enabled = False
            ElseIf moAcceptSheet.Status = 3 Then
                'MsgBox "�а�����ǩ�������ܽ��д˲���"
                cmdCancelAccept.Enabled = False
            End If
        End If
        
    End If
    
    Exit Sub
ErrHandle:
  ShowErrorMsg
 
End Sub
