VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmWizSplitLuggage 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�а�������"
   ClientHeight    =   5460
   ClientLeft      =   3195
   ClientTop       =   1980
   ClientWidth     =   7425
   HelpContextID   =   7000280
   Icon            =   "frmWizSplitLuggage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -120
      TabIndex        =   22
      Top             =   -300
      Width           =   8775
   End
   Begin RTComctl3.CoolButton cmdFinish 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4845
      TabIndex        =   50
      Top             =   4950
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "���(&F)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "frmWizSplitLuggage.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraWizard 
      BorderStyle     =   0  'None
      Caption         =   "��һ��"
      Height          =   4590
      Index           =   0
      Left            =   10620
      TabIndex        =   16
      Top             =   10440
      Width           =   8700
      Begin VB.OptionButton optList 
         Caption         =   "�а��б����(&G)"
         Enabled         =   0   'False
         Height          =   360
         Left            =   2340
         TabIndex        =   18
         Top             =   1965
         Width           =   1905
      End
      Begin VB.OptionButton optNew 
         Caption         =   "ѡ����Χ����(&F)"
         Height          =   285
         Left            =   2340
         TabIndex        =   17
         Top             =   3180
         Value           =   -1  'True
         Width           =   3720
      End
      Begin MSComctlLib.ImageList imglv 
         Left            =   3180
         Top             =   3720
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizSplitLuggage.frx":0166
               Key             =   "company"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizSplitLuggage.frx":02C0
               Key             =   "splitcompany"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizSplitLuggage.frx":041A
               Key             =   "vehicle"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizSplitLuggage.frx":0574
               Key             =   "bus"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizSplitLuggage.frx":06CE
               Key             =   "busowner"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgObject 
         Left            =   4290
         Top             =   3660
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizSplitLuggage.frx":0828
               Key             =   "Company"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizSplitLuggage.frx":0984
               Key             =   "Owner"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizSplitLuggage.frx":0AE0
               Key             =   "Route"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizSplitLuggage.frx":0C3C
               Key             =   "Bus"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizSplitLuggage.frx":0D98
               Key             =   "Vehicle"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizSplitLuggage.frx":0EF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizSplitLuggage.frx":1210
               Key             =   "NoAvailability"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizSplitLuggage.frx":136C
               Key             =   "RepetitiousSettle"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizSplitLuggage.frx":1688
               Key             =   "SettleFinished"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label10 
         Caption         =   "�а��б����:�����ڼ����д򿪵��а�."
         Height          =   345
         Left            =   2310
         TabIndex        =   21
         Top             =   1560
         Width           =   5250
      End
      Begin VB.Label Label11 
         Caption         =   "ѡ����Χ����:�����趨���㷶Χ(���Ρ����������˹�˾)"
         Height          =   360
         Left            =   2340
         TabIndex        =   20
         Top             =   2805
         Width           =   5040
      End
      Begin VB.Label Label9 
         Caption         =   "ѡ������а���Χ:�����а��������������������֡�"
         Height          =   450
         Left            =   2340
         TabIndex        =   19
         Top             =   540
         Width           =   4230
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         X1              =   2310
         X2              =   8505
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         X1              =   2310
         X2              =   8535
         Y1              =   1120
         Y2              =   1120
      End
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Height          =   315
      Left            =   6090
      TabIndex        =   9
      Top             =   4950
      Width           =   1125
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "�ر�(&C)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "frmWizSplitLuggage.frx":17E4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdNext 
      Default         =   -1  'True
      Height          =   315
      Left            =   4845
      TabIndex        =   7
      Top             =   4950
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "��һ��(&N)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "frmWizSplitLuggage.frx":1800
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdPrevious 
      Height          =   315
      Left            =   3600
      TabIndex        =   8
      Top             =   4950
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "��һ��(&P)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "frmWizSplitLuggage.frx":181C
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
      Caption         =   " "
      Height          =   3120
      Left            =   -60
      TabIndex        =   23
      Top             =   4680
      Width           =   9465
      Begin RTComctl3.CoolButton cmdHelp 
         Height          =   315
         Left            =   600
         TabIndex        =   66
         Top             =   270
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "����(&H)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         MICON           =   "frmWizSplitLuggage.frx":1838
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
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   7395
      TabIndex        =   46
      Top             =   0
      Width           =   7395
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   -30
         TabIndex        =   47
         Top             =   810
         Width           =   7485
      End
      Begin VB.Label lblContent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѡ��·������ķ�ʽ��"
         Height          =   180
         Left            =   360
         TabIndex        =   49
         Top             =   450
         Width           =   1980
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���㷽ʽ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   48
         Top             =   150
         Width           =   780
      End
   End
   Begin VB.Frame fraWizard 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "������"
      Height          =   3780
      Index           =   3
      Left            =   0
      TabIndex        =   26
      Top             =   840
      Width           =   7440
      Begin VB.TextBox txtSettleSheetID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1410
         TabIndex        =   69
         Text            =   "0000001"
         Top             =   180
         Width           =   1245
      End
      Begin RTComctl3.CoolButton cmdInfo 
         Height          =   345
         Left            =   6090
         TabIndex        =   29
         Top             =   3420
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "����(&S)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         MICON           =   "frmWizSplitLuggage.frx":1854
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdCancelSelect 
         Height          =   345
         Left            =   4860
         TabIndex        =   28
         Top             =   3420
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "��ѡ(&R)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         MICON           =   "frmWizSplitLuggage.frx":1870
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdSelectAll 
         Height          =   345
         Left            =   3600
         TabIndex        =   27
         Top             =   3420
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "ȫѡ(&A)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         MICON           =   "frmWizSplitLuggage.frx":188C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView lvObject 
         Height          =   2775
         Left            =   240
         TabIndex        =   10
         Top             =   540
         Width           =   7020
         _ExtentX        =   12383
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imglv"
         ColHdrIcons     =   "imgObject"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin RTComctl3.FlatLabel FlatLabel1 
         Height          =   375
         Left            =   1320
         TabIndex        =   70
         Top             =   90
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         InnerStyle      =   2
         Caption         =   ""
      End
      Begin VB.Label lblValidCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   5220
         TabIndex        =   73
         Top             =   180
         Width           =   180
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѡ����:"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   4260
         TabIndex        =   72
         Top             =   180
         Width           =   810
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ч"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   2940
         TabIndex        =   71
         Top             =   180
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image imgEnabled 
         Height          =   405
         Left            =   2940
         Picture         =   "frmWizSplitLuggage.frx":18A8
         Stretch         =   -1  'True
         Top             =   90
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lblHaveProtocol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ĭ��Э��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   240
         TabIndex        =   68
         Top             =   3510
         Width           =   1365
      End
      Begin VB.Label lblNoProtocol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Э��ĳ�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   240
         TabIndex        =   67
         Top             =   3480
         Width           =   2025
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "�а�ǩ����:"
         Height          =   180
         Left            =   240
         TabIndex        =   31
         Top             =   210
         Width           =   1290
      End
      Begin VB.Label lblLuggageSheetCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ǩ��������:123"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   5775
         TabIndex        =   30
         Top             =   180
         Width           =   1260
      End
   End
   Begin VB.Frame fraWizard 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "������"
      Height          =   3900
      Index           =   5
      Left            =   0
      TabIndex        =   54
      Top             =   960
      Width           =   7440
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�а�������ϸ��"
         Height          =   2775
         Left            =   360
         TabIndex        =   60
         Top             =   840
         Width           =   6615
         Begin MSComctlLib.ListView vsDetailPrice 
            Height          =   2355
            Left            =   720
            TabIndex        =   61
            Top             =   240
            Width           =   5640
            _ExtentX        =   9948
            _ExtentY        =   4154
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin VB.Image Image3 
            Height          =   480
            Left            =   120
            Picture         =   "frmWizSplitLuggage.frx":2172
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�а�������ܵ�"
         Height          =   765
         Left            =   360
         TabIndex        =   55
         Top             =   0
         Width           =   6645
         Begin VB.Label lblprice_1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "100"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   1500
            TabIndex        =   65
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "���˼�:"
            Height          =   180
            Left            =   750
            TabIndex        =   64
            Top             =   360
            Width           =   630
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   120
            Picture         =   "frmWizSplitLuggage.frx":247C
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "ǩ��������:"
            Height          =   180
            Left            =   5160
            TabIndex        =   63
            Top             =   360
            Width           =   990
         End
         Begin VB.Label lblSheetId 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "2��"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   6240
            TabIndex        =   62
            Top             =   360
            Width           =   270
         End
         Begin VB.Label lblNeedSplitMoney 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "220.3"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   4590
            TabIndex        =   59
            Top             =   360
            Width           =   450
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ӧ������:"
            Height          =   180
            Left            =   3450
            TabIndex        =   58
            Top             =   360
            Width           =   990
         End
         Begin VB.Label lblTotalPrice 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "220.3"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   2760
            TabIndex        =   57
            Top             =   360
            Width           =   450
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ܶ�:"
            Height          =   180
            Left            =   2130
            TabIndex        =   56
            Top             =   360
            Width           =   450
         End
      End
   End
   Begin VB.Frame fraWizard 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "���һ��"
      Height          =   3780
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   7410
      Begin MSComctlLib.ProgressBar CreateProgressBar 
         Height          =   300
         Left            =   690
         Negotiate       =   -1  'True
         TabIndex        =   43
         Top             =   3435
         Width           =   6570
         _ExtentX        =   11589
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.ListBox lstCreateInfo 
         Appearance      =   0  'Flat
         Height          =   2910
         ItemData        =   "frmWizSplitLuggage.frx":2786
         Left            =   210
         List            =   "frmWizSplitLuggage.frx":2788
         MultiSelect     =   2  'Extended
         TabIndex        =   42
         Top             =   450
         Width           =   7005
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "ִ�����:"
         Height          =   255
         Left            =   210
         TabIndex        =   45
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   44
         Top             =   3450
         Width           =   615
      End
   End
   Begin VB.Frame fraWizard 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "�ڶ���"
      Height          =   3780
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Top             =   840
      Width           =   7440
      Begin VB.ComboBox cboAcceptType 
         Height          =   300
         ItemData        =   "frmWizSplitLuggage.frx":278A
         Left            =   4800
         List            =   "frmWizSplitLuggage.frx":278C
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   1410
         Width           =   1575
      End
      Begin VB.ComboBox cboSellStation 
         Height          =   300
         Left            =   1605
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   1410
         Width           =   1710
      End
      Begin FText.asFlatTextBox txtObject 
         Height          =   315
         Left            =   4800
         TabIndex        =   32
         Top             =   1860
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ButtonVisible   =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Left            =   4800
         TabIndex        =   6
         Top             =   2820
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         Format          =   61669376
         CurrentDate     =   37642
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   1575
         TabIndex        =   3
         Top             =   2820
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   529
         _Version        =   393216
         Format          =   61669376
         CurrentDate     =   37622
      End
      Begin MSComctlLib.ImageCombo imgcbo 
         Height          =   330
         Left            =   1590
         TabIndex        =   1
         Top             =   1860
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "imglv"
      End
      Begin MSComCtl2.DTPicker dtpMonth 
         Height          =   300
         Left            =   1590
         TabIndex        =   33
         Top             =   2370
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��"
         Format          =   61669379
         UpDown          =   -1  'True
         CurrentDate     =   37642
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ϳ�վ(&T):"
         Height          =   180
         Left            =   510
         TabIndex        =   52
         Top             =   1470
         Width           =   900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���˷�ʽ(&T):"
         Height          =   180
         Left            =   3630
         TabIndex        =   35
         Top             =   1530
         Width           =   1080
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ���·�(&M):"
         Height          =   225
         Left            =   450
         TabIndex        =   34
         Top             =   2430
         Width           =   1380
      End
      Begin VB.Line Line6 
         BorderColor     =   &H8000000E&
         X1              =   450
         X2              =   6360
         Y1              =   1245
         Y2              =   1245
      End
      Begin VB.Line Line5 
         BorderColor     =   &H8000000C&
         X1              =   435
         X2              =   6360
         Y1              =   1230
         Y2              =   1230
      End
      Begin VB.Label label1 
         BackStyle       =   0  'Transparent
         Caption         =   "����(&S):"
         Height          =   180
         Index           =   1
         Left            =   510
         TabIndex        =   0
         Top             =   1920
         Width           =   1080
      End
      Begin VB.Label label2 
         BackStyle       =   0  'Transparent
         Caption         =   "����(&O):"
         Height          =   180
         Index           =   1
         Left            =   3660
         TabIndex        =   4
         Top             =   1950
         Width           =   810
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ����(&S):"
         Height          =   225
         Left            =   450
         TabIndex        =   2
         Top             =   2865
         Width           =   1380
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&E):"
         Height          =   180
         Left            =   3660
         TabIndex        =   5
         Top             =   2880
         Width           =   1080
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWizSplitLuggage.frx":278E
         Height          =   840
         Left            =   450
         TabIndex        =   25
         Top             =   390
         Width           =   6165
      End
   End
   Begin VB.Frame fraWizard 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "���Ĳ�"
      Height          =   3780
      Index           =   2
      Left            =   0
      TabIndex        =   12
      Top             =   840
      Width           =   7410
      Begin VB.TextBox txtOperator 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1980
         TabIndex        =   36
         Top             =   1290
         Width           =   1785
      End
      Begin VB.TextBox txtSheetID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1980
         TabIndex        =   13
         Top             =   840
         Width           =   1785
      End
      Begin MSComCtl2.DTPicker dtpSplitDate 
         Height          =   300
         Left            =   1980
         TabIndex        =   39
         Top             =   1740
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         Format          =   61669376
         CurrentDate     =   37642
      End
      Begin FText.asFlatMemo txtAnnotation 
         Height          =   1020
         Left            =   1980
         TabIndex        =   40
         Top             =   2280
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   1799
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin VB.Label lblMemo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ע(&R):"
         Height          =   180
         Left            =   600
         TabIndex        =   41
         Top             =   2325
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&D):"
         Height          =   180
         Left            =   600
         TabIndex        =   38
         Top             =   1890
         Width           =   1080
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������(&P):"
         Height          =   180
         Left            =   600
         TabIndex        =   37
         Top             =   1350
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���㵥���(&S):"
         Height          =   180
         Left            =   600
         TabIndex        =   14
         Top             =   900
         Width           =   1260
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "��д���㵥��ź��������㵥��Ϣ�����ݽ����������ͬ����ĳЩ���㵥��Ϣ���Բ�����д��"
         Height          =   405
         Left            =   570
         TabIndex        =   15
         Top             =   240
         Width           =   5445
      End
   End
End
Attribute VB_Name = "frmWizSplitLuggage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3888FDA40140"
Option Explicit
Dim mSheetCount As Integer  ' ͳ��ǩ��������
Dim mNum As Integer  'ѡ������ǩ������
Dim mSheetID() As String 'ǩ��������
Dim mVehicleID() As String
Dim mNotProtocol As Boolean '�Ƿ���ûЭ��ĳ���
Dim m_bLogFileValid As Boolean '4��־�ļ�
Dim m_bPromptWhenError As Boolean '2�Ƿ���ʾ����
Dim CancelHasPress As Boolean
Dim mProtocolID As String ' Ĭ��Э��
Dim nValidCount As Integer '��ѡ����

Private Sub cmdFinish_Click()
 On Error GoTo ErrHandle
    Dim i As Integer
    'Ȩ����֤
    m_oFinanceSheet.SplitMan
    cmdNext_Click
    m_oFinanceSheet.AddNew
    m_oFinanceSheet.SheetID = Trim(txtSheetID.Text)
    m_oFinanceSheet.SellStationID = ResolveDisplay(Trim(cboSellStation.Text))
    m_oFinanceSheet.AcceptType = GetLuggageTypeInt(Trim(cboAcceptType.Text))
    m_oFinanceSheet.SplitObjectID = ResolveDisplay(Trim(txtObject.Text))
    m_oFinanceSheet.SplitObjectName = ResolveDisplayEx(Trim(txtObject.Text))
    m_oFinanceSheet.SplitObjectType = GetObjectTypeInt(Trim(imgcbo.Text))
    m_oFinanceSheet.SettleMonth = dtpMonth.Value
    m_oFinanceSheet.StartSettleDate = dtpStartDate.Value
    m_oFinanceSheet.StopSettleDate = dtpEndDate.Value
    m_oFinanceSheet.OperatorName = Trim(txtOperator.Text)
    m_oFinanceSheet.OperateTime = Trim(dtpSplitDate.Value)
    m_oFinanceSheet.Remark = Trim(txtAnnotation.Text)
    m_oFinanceSheet.Update
    
       '��ʼ���ɽ��㵥  ��дlstCreateInfo��Ϣ
       CreateFinanceInfo
'       cmdFinish.Enabled = False
        cmdPrevious.Enabled = False
        cmdNext.Enabled = False
        cmdFinish.Enabled = False
        '��ӡ���㵥
        Dim mAnswer
        mAnswer = MsgBox("�Ƿ��ӡ���㵥", vbQuestion + vbYesNo, Me.Caption)
        If mAnswer = vbYes Then
            frmPrintFinSheet.SheetID = Trim(txtSheetID.Text)
            frmPrintFinSheet.mRePrint = False
            frmPrintFinSheet.ZOrder 0
            frmPrintFinSheet.Show vbModal
        End If
 Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub

'//***********************************    ������־��  **************************************
'���ɽ��㵥��־
Private Sub CreateFinanceInfo()
  On Error GoTo ErrHandle
    Dim mStart, mEnd As Date '��ʼʱ��,����ʱ��
    Dim bCreateOk As Integer
    Dim nCreateCount As Integer
    Dim i As Long
    mStart = Now
    RecordLog "================================================================="
    RecordLog "=  RTStation ���ɽ��㵥��¼��־"
    RecordLog "= ----------------------------------"
    RecordLog "=  ʹ���ߣ�" & m_oAUser.UserID & "/" & m_oAUser.UserName
    RecordLog "=  ���ɽ��㵥����:" & Format(mStart, "YYYY-MM-DD")
    RecordLog "=  ��ǰʱ��:" & Format(mStart, "YYYY-MM-DD HH:MM:SS")
    RecordLog "================================================================="
    CreateProgressBar.Min = 0
    CreateProgressBar.Max = mNum
  
     For i = 1 To mNum
        bCreateOk = CreateFinanceSheetRs(i)
        lblProgress.Caption = Str(Int(100 * i / mNum)) & "%"
        lblProgress.Refresh
        If bCreateOk = 1 Then
            nCreateCount = nCreateCount + 1
        End If
        If bCreateOk = 3 Then
            i = i - 1
        End If
        CreateProgressBar.Value = i
    Next
    lblProgress.Caption = "100%"
    lblProgress.Refresh

'Report:
    RecordLog "================================================================="
    RecordLog "���ɽ��㵥����"
    RecordLog "�ܹ����ɽ��㵥:" & nCreateCount & "��"
    RecordLog "δ���ɽ��㵥:" & mNum - nCreateCount & "��"
    mEnd = Now
    RecordLog "����ʱ��:" & Format(mEnd, "HH:MM:SS")
    RecordLog "��ʹ��ʱ��:" & Format(mEnd - mStart, "HHСʱMM��SS��")
    
    
    CloseLogFile
  Exit Sub
ErrHandle:
  ShowErrorMsg
End Sub

'����ĳ���㵥����¼���� 1�ɹ���2����3����
Private Function CreateFinanceSheetRs(Index As Long) As Integer
    Dim vbMsg As VbMsgBoxResult
    Dim ErrString As String
    Dim bCreateOk As Integer
    Dim nErrNumber As Long
    Dim szErrDescription As String
    Dim mszSheetID() As String   'Ϊ����Ӧ�ӿ���������
    Static nHasPrompt, nPromptTime As Integer
    
    On Error GoTo here
    '��ʼ�����ÿ�ʼ����ʱ�Ĵ������
    If nPromptTime < 2 Then nPromptTime = 2
    ReDim mszSheetID(1 To 1)
    mszSheetID(1) = mSheetID(Index)
    bCreateOk = 1
    ErrString = " ����ɹ�    "
    m_oFinanceSheet.SplitCarrySheets mszSheetID, Trim(txtOperator.Text), dtpSplitDate.Value
    RecordLog mSheetID(Index) & "[" & ErrString & "]"
ErrContinue:
    DoEvents
    CreateFinanceSheetRs = bCreateOk
    Exit Function
here:
    bCreateOk = 2
    nErrNumber = err.Number
    szErrDescription = err.Description
    ErrString = mSheetID(Index) & "[  δ����]" & _
        " * ��������:(" & Trim(Str(nErrNumber)) & ")" & Trim(szErrDescription) & " *"
    RecordLog ErrString
    If m_bPromptWhenError Then
        ErrString = "���㵥" & mSheetID(Index) & "δ���ɣ�" & vbCrLf & _
            Trim(szErrDescription) & "(" & Trim(Str(nErrNumber)) & ")"
        vbMsg = MsgBox(ErrString, vbExclamation + vbAbortRetryIgnore + vbDefaultButton3)
        Select Case vbMsg
               Case vbAbort
                   CancelHasPress = True
               Case vbRetry
                   CreateFinanceSheetRs = 3
                   Exit Function
               Case vbIgnore
                   If nHasPrompt >= nPromptTime - 1 Then
                        If MsgBox("�Ժ�����ʾ���ɴ���", vbQuestion + vbYesNo) = vbYes Then
                            m_bPromptWhenError = False
                        End If
                        nHasPrompt = 0
                        nPromptTime = nPromptTime + 1
                   End If
                   nHasPrompt = nHasPrompt + 1
               Exit Function
        End Select
    Else
        GoTo ErrContinue
    End If
End Function

Private Sub CloseLogFile()
    On Error Resume Next
    If m_bLogFileValid Then
        Close #1
    End If
End Sub

Private Sub RecordLog(log As String)
    With lstCreateInfo
        .AddItem log
        .ListIndex = .ListCount - 1
        .Refresh
    End With
    If m_bLogFileValid Then
        AddLogToFile log
    End If
End Sub

Private Sub AddLogToFile(log As String)
    On Error Resume Next
    If m_bLogFileValid Then
        Print #1, log
    End If
End Sub



Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

'//**************************************************************************************




Private Sub cmdInfo_Click()
 If lvObject.ListItems.Count = 0 Then Exit Sub
 frmSheetInfo.SheetID = Trim(lvObject.SelectedItem.Text)
 frmSheetInfo.Show vbModal
End Sub

Private Sub cmdNext_Click()
On Error GoTo ErrHandle
  Dim rsTemp As Recordset
  Dim i As Integer
   cmdPrevious.Enabled = True
   If fraWizard(1).Visible = True Then
      '�жϲ��������û��ѡ
      If txtObject.Text = "" Then
         MsgBox " ��ѡ��������!", vbInformation, Me.Caption
         cmdPrevious.Enabled = False
         Exit Sub
      End If
      cmdNext.Enabled = True
      fraWizard(1).Visible = False
      fraWizard(2).Visible = True
      fraWizard(3).Visible = False
      fraWizard(4).Visible = False
      fraWizard(5).Visible = False
       '��յڶ�ҳ����
'      txtSheetID.Text = ""
      txtOperator.Text = m_oAUser.UserID
      txtAnnotation.Text = ""
      '�Զ����ɽ��㵥�� YYYYMM0001��ʽ
     Set rsTemp = m_oLugFinSvr.GetFinSheetID
     If rsTemp.RecordCount = 0 Then
      txtSheetID.Text = CStr(Year(Now)) + CStr(Month(Now)) + "0001"
     Else
      txtSheetID.Text = CStr(rsTemp!fin_sheet_id + 1)
     End If
      txtOperator.SetFocus
      Set rsTemp = Nothing
   ElseIf fraWizard(2).Visible = True Then

      '�жϽ��㵥���Ƿ�Ϊ��
      If txtSheetID.Text = "" Then
       MsgBox " ���㵥��Ų���Ϊ��!", vbInformation, Me.Caption
       Exit Sub
      End If
      '���Ҫ�����ǩ������¼
      FilllvObject
       '��ʾǩ��������
      mSheetCount = lvObject.ListItems.Count
      lblLuggageSheetCount.Caption = "ǩ��������: " & CStr(mSheetCount)
 
      cmdFinish.Visible = False
      cmdFinish.Enabled = True
      cmdNext.Visible = True
      fraWizard(1).Visible = False
      fraWizard(2).Visible = False
      fraWizard(3).Visible = True
      fraWizard(4).Visible = False
      fraWizard(5).Visible = False
      
      txtSettleSheetID.Text = ""
      txtSettleSheetID.SetFocus
      cmdNext.Default = False
      nValidCount = 0
      lblValidCount.Caption = nValidCount
   ElseIf fraWizard(3).Visible = True Then

      'ͳ����ѡ��ǩ������
      CountmNum
      If mSheetCount > 0 Then
        If mNum = 0 Then
           MsgBox "���ڲ����ǩ������ǰ�򹴣�", vbExclamation, Me.Caption
           Exit Sub
        End If
      Else
        Exit Sub
      End If
        
      '��ûЭ��ĳ�������ΪĬ��Э��
      If mNotProtocol = True Then
        Dim mAnswer
        mAnswer = MsgBox("�б��д���û��Э��ĳ���,�Ƿ���Щ��������ΪĬ��Э��?", vbInformation + vbYesNo, Me.Caption)
        If mAnswer = vbYes Then '����ʾ��û��Э��ĳ���,����Ĭ��Э��
            m_oProtocol.Init m_oAUser
            m_oProtocol.SetAllNoProtocolVehicle mVehicleID
        Else
            Exit Sub
        End If
      End If
        
      '���ǩ��������
      FillSheetID
      FIllCarryInfo
      
      '���lstCreateInfo
      lstCreateInfo.clear
      cmdFinish.Enabled = True
      cmdFinish.Visible = True
      cmdNext.Visible = False
      fraWizard(1).Visible = False
      fraWizard(2).Visible = False
      fraWizard(3).Visible = False
      fraWizard(4).Visible = False
      fraWizard(5).Visible = True
   ElseIf fraWizard(5).Visible = True Then
      
      cmdFinish.Visible = True
      cmdFinish.Enabled = False
      cmdNext.Visible = False
      fraWizard(1).Visible = False
      fraWizard(2).Visible = False
      fraWizard(3).Visible = False
      fraWizard(4).Visible = True
      fraWizard(5).Visible = False
   End If
   
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub
'ͳ����ѡ��ǩ������
Private Sub CountmNum()
  Dim i As Integer
   mNum = 0
   If mSheetCount > 0 Then
      For i = 1 To mSheetCount
        If lvObject.ListItems.Item(i).Checked = True Then
            mNum = mNum + 1
        End If
      Next i
   End If
End Sub
'���Ҫ�����ǩ������¼
Private Sub FilllvObject()
  On Error GoTo ErrHandle
    Dim i As Integer
    Dim rsTemp As Recordset
    Dim rsTempRs As Recordset
    Dim lvItem As ListItem
    Dim Accept As Integer

 '�ڲ�ѯǩ������¼ʱ������һ���жϣ�������˷�ʽΪ�յĻ�����ѯ���е����˷�ʽ��ǩ������Ϣ��
 If cboAcceptType.Text = "" Then
    Accept = -1
 Else
   Accept = GetLuggageTypeInt(Trim(cboAcceptType.Text))
 End If
 
        '����г���Э���ǩ����
        If Trim(imgcbo.Text) = "���ʹ�˾" Then
           Set rsTemp = m_oLugFinSvr.GetWillSplitSheetRS(dtpStartDate.Value, dtpEndDate.Value, ResolveDisplay(Trim(cboSellStation.Text)), Accept, , , , ResolveDisplay(Trim(txtObject.Text)))
        ElseIf Trim(imgcbo.Text) = "���˹�˾" Then
                Set rsTemp = m_oLugFinSvr.GetWillSplitSheetRS(dtpStartDate.Value, dtpEndDate.Value, ResolveDisplay(Trim(cboSellStation.Text)), Accept, , , ResolveDisplay(Trim(txtObject.Text)))
        ElseIf Trim(imgcbo.Text) = "����" Then
                Set rsTemp = m_oLugFinSvr.GetWillSplitSheetRS(dtpStartDate.Value, dtpEndDate.Value, ResolveDisplay(Trim(cboSellStation.Text)), Accept, , ResolveDisplay(Trim(txtObject.Text)))
        ElseIf Trim(imgcbo.Text) = "����" Then
                Set rsTemp = m_oLugFinSvr.GetWillSplitSheetRS(dtpStartDate.Value, dtpEndDate.Value, ResolveDisplay(Trim(cboSellStation.Text)), Accept, , , , , ResolveDisplay(Trim(txtObject.Text)))
         
        ElseIf Trim(imgcbo.Text) = "����" Then
                 Set rsTemp = m_oLugFinSvr.GetWillSplitSheetRS(dtpStartDate.Value, dtpEndDate.Value, ResolveDisplay(Trim(cboSellStation.Text)), Accept, ResolveDisplay(Trim(txtObject.Text)))
        End If
    
    FilllvObjectHead '�������
    If rsTemp.RecordCount = 0 Then GoTo here
    lvObject.ListItems.clear
    For i = 1 To rsTemp.RecordCount
      Set lvItem = lvObject.ListItems.Add(, , FormatDbValue(rsTemp!sheet_id))
          lvItem.SubItems(1) = GetLuggageTypeString(FormatDbValue(rsTemp!accept_type))
          lvItem.SubItems(2) = FormatDbValue(rsTemp!bus_id)
          lvItem.SubItems(3) = FormatDbValue(rsTemp!price_item_1)
          lvItem.SubItems(4) = FormatDbValue(rsTemp!total_price)
          lvItem.SubItems(5) = FormatDbValue(rsTemp!protocol_name)
          lvItem.SubItems(6) = FormatDbValue(rsTemp!transport_company_short_name)
          lvItem.SubItems(7) = FormatDbValue(rsTemp!splict_company_short_name)
          lvItem.SubItems(8) = FormatDbValue(rsTemp!bus_date)
          lvItem.SubItems(9) = FormatDbValue(rsTemp!license_tag_no)
      rsTemp.MoveNext
    Next i
    mNotProtocol = False
here:
    '����޳���Э���ǩ����
        If Trim(imgcbo.Text) = "���ʹ�˾" Then
           Set rsTempRs = m_oLugFinSvr.GetWillSplitSheetRSTemp(dtpStartDate.Value, dtpEndDate.Value, ResolveDisplay(Trim(cboSellStation.Text)), Accept, , , , ResolveDisplay(Trim(txtObject.Text)))
        ElseIf Trim(imgcbo.Text) = "���˹�˾" Then
                Set rsTempRs = m_oLugFinSvr.GetWillSplitSheetRSTemp(dtpStartDate.Value, dtpEndDate.Value, ResolveDisplay(Trim(cboSellStation.Text)), Accept, , , ResolveDisplay(Trim(txtObject.Text)))
        ElseIf Trim(imgcbo.Text) = "����" Then
                Set rsTempRs = m_oLugFinSvr.GetWillSplitSheetRSTemp(dtpStartDate.Value, dtpEndDate.Value, ResolveDisplay(Trim(cboSellStation.Text)), Accept, , ResolveDisplay(Trim(txtObject.Text)))
        ElseIf Trim(imgcbo.Text) = "����" Then
                Set rsTempRs = m_oLugFinSvr.GetWillSplitSheetRSTemp(dtpStartDate.Value, dtpEndDate.Value, ResolveDisplay(Trim(cboSellStation.Text)), Accept, , , , , ResolveDisplay(Trim(txtObject.Text)))
         
        ElseIf Trim(imgcbo.Text) = "����" Then
                 Set rsTempRs = m_oLugFinSvr.GetWillSplitSheetRSTemp(dtpStartDate.Value, dtpEndDate.Value, ResolveDisplay(Trim(cboSellStation.Text)), Accept, ResolveDisplay(Trim(txtObject.Text)))
        End If
        
    lblNoProtocol.Visible = False
    lblHaveProtocol.Visible = False
    If rsTempRs.RecordCount = 0 Then Exit Sub
    ReDim mVehicleID(1 To rsTempRs.RecordCount, 1 To 2)
    mNotProtocol = True
    For i = 1 To rsTempRs.RecordCount
      Set lvItem = lvObject.ListItems.Add(, , FormatDbValue(rsTempRs!sheet_id))
          lvItem.SubItems(1) = GetLuggageTypeString(FormatDbValue(rsTempRs!accept_type))
          lvItem.SubItems(2) = FormatDbValue(rsTempRs!bus_id)
          lvItem.SubItems(3) = FormatDbValue(rsTempRs!price_item_1)
          lvItem.SubItems(4) = FormatDbValue(rsTempRs!total_price)
          lvItem.SubItems(5) = ""
          lvItem.SubItems(6) = FormatDbValue(rsTempRs!transport_company_short_name)
          lvItem.SubItems(7) = FormatDbValue(rsTempRs!splict_company_short_name)
          lvItem.SubItems(8) = FormatDbValue(rsTempRs!bus_date)
          lvItem.SubItems(9) = FormatDbValue(rsTempRs!license_tag_no)
          mVehicleID(i, 1) = FormatDbValue(rsTempRs!vehicle_id)
          mVehicleID(i, 2) = FormatDbValue(rsTempRs!accept_type)
      rsTempRs.MoveNext
    Next i
    
    '�ж���û������Ĭ��Э��
    Dim rsTempProtocol As Recordset
    Set rsTempProtocol = m_oLugFinSvr.GetHaveProtocol
    If rsTempProtocol.RecordCount = 0 Then
        lblNoProtocol.Visible = True
    Else
        

        lblHaveProtocol.Visible = True
        lblHaveProtocol.Caption = "Ĭ��Э��������"
'        lblHaveProtocol.Caption = FormatDbValue(rsTempProtocol!protocol_name) & " " & lblHaveProtocol.Caption
'        mProtocolID = FormatDbValue(rsTempProtocol!protocol_id)
    End If
    

    
    Set rsTemp = Nothing
    Set rsTempRs = Nothing
   Exit Sub
ErrHandle:
  ShowErrorMsg
  Set rsTemp = Nothing
  Set rsTempRs = Nothing
End Sub

 '�������
Private Sub FilllvObjectHead()
   With lvObject.ColumnHeaders
     .clear
     .Add , , "ǩ������", 1200
     .Add , , "���˷�ʽ", 900
     .Add , , "���δ���", 900
     .Add , , "���˷�", 900
     .Add , , "�ܼ�", 700
     .Add , , "Э������", 1200
     .Add , , "��Ӫ��˾", 900
     .Add , , "���ʹ�˾", 900
     .Add , , "��������", 900
     .Add , , "���ƺ�", 900
   End With
End Sub
'���ǩ��������
Private Sub FillSheetID()
 On Error GoTo ErrHandle
   Dim i As Integer
   Dim j As Integer
'   Dim oFinanceSheet As New FinanceSheet
'   Dim asztemp() As String
   If mNum = 0 Then Exit Sub
   ReDim mSheetID(1 To mNum)
   j = 1
   For i = 1 To mSheetCount
      If lvObject.ListItems.Item(i).Checked = True Then
         mSheetID(j) = Trim(lvObject.ListItems(i).Text)
         j = j + 1
      End If
   Next i
 Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub
Private Sub cmdPrevious_Click()
On Error GoTo ErrHandle

   If fraWizard(4).Visible = True Then
      cmdFinish.Visible = True
      cmdFinish.Enabled = True
      fraWizard(1).Visible = False
      fraWizard(2).Visible = False
      fraWizard(5).Visible = True
      fraWizard(4).Visible = False
      fraWizard(3).Visible = False
   ElseIf fraWizard(5).Visible = True Then
      cmdNext.Visible = True
      cmdNext.Enabled = True
      cmdFinish.Visible = False
      fraWizard(1).Visible = False
      fraWizard(2).Visible = False
      fraWizard(3).Visible = True
      fraWizard(4).Visible = False
      fraWizard(5).Visible = False
   ElseIf fraWizard(3).Visible = True Then
      cmdNext.Visible = True
      cmdNext.Enabled = True
      cmdFinish.Visible = False
      fraWizard(1).Visible = False
      fraWizard(2).Visible = True
      fraWizard(3).Visible = False
      fraWizard(4).Visible = False
   ElseIf fraWizard(2).Visible = True Then
      cmdNext.Visible = True
      cmdNext.Enabled = True
      cmdFinish.Visible = False
      cmdPrevious.Enabled = False
      lvObject.ListItems.clear
      fraWizard(1).Visible = True
      fraWizard(2).Visible = False
      fraWizard(3).Visible = False
      fraWizard(4).Visible = False
   End If
   
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

Private Sub cmdSelectAll_Click()
  On Error GoTo ErrHandle
    Dim i As Integer
    If lvObject.ListItems.Count = 0 Then Exit Sub
    
    For i = 1 To lvObject.ListItems.Count
        lvObject.ListItems.Item(i).Checked = True
    Next i
    nValidCount = lvObject.ListItems.Count
    lblValidCount.Caption = nValidCount
  Exit Sub
ErrHandle:
  ShowErrorMsg
End Sub

Private Sub Form_Load()
 Dim nlen As Integer
 Dim szTemp() As String
 Dim i As Integer
 
 AlignFormPos Me
 cmdPrevious.Enabled = False
 cmdFinish.Visible = False
 cmdNext.Visible = True
    
 fraWizard(1).Visible = True
 fraWizard(2).Visible = False
 fraWizard(3).Visible = False
 fraWizard(4).Visible = False
 fraWizard(5).Visible = False
      
  '��ʼ��cbo����
    dtpEndDate.Value = Now
    dtpSplitDate.Value = Now
    lblNoProtocol.Visible = False
    lblHaveProtocol.Visible = False
  '����ϳ�վ
    '��վ
    cboSellStation.clear
    FillSellStation cboSellStation
    cboSellStation.ListIndex = 0
  ' ������˷�ʽ
  With cboAcceptType
    .clear
    .AddItem ""
    .AddItem szAcceptTypeGeneral
    .AddItem szAcceptTypeMan
    .ListIndex = 0
  End With
   
  '����
  '0-���ʹ�˾ 1-���� 2-���˹�˾ 3-���� 4-����
  With imgcbo
    
    .ComboItems.clear
    .ComboItems.Add , , "���ʹ�˾", 1
    .ComboItems.Add , , "����", 3
    .ComboItems.Add , , "���˹�˾", 2
    .ComboItems.Add , , "����", 5
    .ComboItems.Add , , "����", 4
    .Locked = True
'    .Text = "����"
    .ComboItems(2).Selected = True
   End With
  
   '��ʼ����������
   dtpMonth.Value = Date
   dtpStartDate.Value = CDate(Format(dtpMonth.Value, "yyyy-mm") & "-1" & " 00:00:01")
   Select Case Month(dtpMonth.Value)
          Case 1, 3, 5, 7, 8, 10, 12
           dtpEndDate.Value = CDate(Format(dtpMonth.Value, "yyyy-mm") & "-31" & " 23:59:59")
          Case 4, 6, 9, 11
           dtpEndDate.Value = CDate(Format(dtpMonth.Value, "yyyy-mm") & "-30" & " 23:59:59")
          Case 2
           dtpEndDate.Value = CDate(Format(dtpMonth.Value, "yyyy-mm") & "-28" & " 23:59:59")
   End Select
End Sub



Private Sub lstCreateInfo_DblClick()
 MsgBox lstCreateInfo.Text, vbInformation + vbOKOnly, "������Ϣ"
End Sub

Private Sub lvObject_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvObject, ColumnHeader.Index
End Sub

Private Sub lvObject_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked = True Then
        nValidCount = nValidCount + 1
        lblValidCount.Caption = nValidCount
    Else
        nValidCount = nValidCount - 1
        lblValidCount.Caption = nValidCount
    End If
End Sub

Private Sub txtObject_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init m_oAUser
    If Trim(imgcbo.Text) = "���ʹ�˾" Or Trim(imgcbo.Text) = "���˹�˾" Then
       aszTemp = oShell.SelectCompany()
    ElseIf Trim(imgcbo.Text) = "����" Then
       aszTemp = oShell.SelectVehicle()
    ElseIf Trim(imgcbo.Text) = "����" Then
       aszTemp = oShell.SelectOwner()
    ElseIf Trim(imgcbo.Text) = "����" Then
       aszTemp = oShell.SelectBus()
    End If
    
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtObject.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))
    
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub
Private Sub Form_Unload(Cancel As Integer)
 SaveFormPos Me
 Unload Me

End Sub

Private Sub lvObject_DblClick()
cmdInfo_Click
End Sub

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdCancelSelect_Click()
  On Error GoTo ErrHandle
    Dim i As Integer
    If lvObject.ListItems.Count = 0 Then Exit Sub
    For i = 1 To lvObject.ListItems.Count
       lvObject.ListItems.Item(i).Checked = False
    Next i
    nValidCount = 0
    lblValidCount.Caption = nValidCount
  Exit Sub
ErrHandle:
  ShowErrorMsg
End Sub


Private Sub txtSettleSheetID_Change()
    imgEnabled.Visible = False
End Sub

Private Sub txtSettleSheetID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FillTOlvObject
    End If
End Sub

Private Sub FillTOlvObject()
On Error GoTo here
    Dim i As Integer
    Dim rsTemp As Recordset
    
    If lvObject.ListItems.Count = 0 Then Exit Sub
    For i = 1 To lvObject.ListItems.Count
        If Trim(txtSettleSheetID.Text) = Trim(lvObject.ListItems.Item(i).Text) Then
            GoTo ok
        End If
    Next i
    imgEnabled.Visible = True
    MsgBox " �б���û�д�ǩ����!", vbExclamation, Me.Caption
    Beep
    '�б��д���Ч·��
ok: imgEnabled.Visible = False
    For i = 1 To lvObject.ListItems.Count
        If Trim(txtSettleSheetID.Text) = Trim(lvObject.ListItems.Item(i).Text) Then
            lvObject.ListItems.Item(i).Checked = True
            nValidCount = nValidCount + 1
            lblValidCount.Caption = nValidCount
            lvObject.ListItems.Item(i).EnsureVisible
        End If
    Next i
    txtSettleSheetID.Text = ""
    txtSettleSheetID.SetFocus
    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub txtSheetID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      txtOperator.SetFocus
   End If
End Sub
'�����Ϣ
Public Sub FIllCarryInfo()
  Dim oFinanceSheet As New FinanceSheet
  Dim oLuggageParam As New LuggageParam
  Dim rsTemp As New Recordset
  Dim aszTemp() As String
  Dim aszSheetID() As String
  Dim i As Integer, j As Integer, k As Integer
  k = 1
  Dim nCount As Integer
  Dim nItem As Integer
  Dim nlen As Integer

  Dim lvItem As ListItem
'�������
With vsDetailPrice.ColumnHeaders
    .clear
    .Add , , "���δ���", 900
    .Add , , "���ƺ�", 900
    .Add , , "Э������", 1200
    .Add , , "���˷�", 900
    .Add , , "�ܶ�", 800
    .Add , , "Ӧ���", 900
End With
vsDetailPrice.ListItems.clear

  Dim iTotalPrice As Double, iNeedSplitPrice As Double, iPrice_1 As Double
  nCount = lvObject.ListItems.Count
  For i = 1 To nCount
     If lvObject.ListItems.Item(i).Checked = True Then
        nItem = nItem + 1
     End If
  Next
  If nItem = 0 Then Exit Sub
  ReDim aszSheetID(1 To nItem) As String
  For i = 1 To nCount
    If lvObject.ListItems.Item(i).Checked = True Then
       aszSheetID(k) = lvObject.ListItems(i).Text
       k = k + 1
    End If
  Next i
'��ʾ�а�������ܺ���ϸ��Ϣ
  ReDim aszTemp(1 To nItem, 1 To 16)
  aszTemp = oFinanceSheet.PreviewSplitCarrySheets(aszSheetID)
  nlen = ArrayLength(aszTemp)
  For i = 1 To nlen
       '������Ϣ
     iPrice_1 = iPrice_1 + aszTemp(i, 7)
     iTotalPrice = iTotalPrice + aszTemp(i, 6)
     iNeedSplitPrice = iNeedSplitPrice + aszTemp(i, 3)
           '��ϸ��Ϣ
     Set lvItem = vsDetailPrice.ListItems.Add(, , aszTemp(i, 1))
     lvItem.SubItems(1) = aszTemp(i, 2)
     lvItem.SubItems(2) = aszTemp(i, 5)
     lvItem.SubItems(3) = aszTemp(i, 7)
     lvItem.SubItems(4) = aszTemp(i, 6)
     lvItem.SubItems(5) = aszTemp(i, 3)
   Next i

   lblprice_1.Caption = iPrice_1
   lblTotalPrice.Caption = iTotalPrice
   lblNeedSplitMoney.Caption = iNeedSplitPrice
   lblSheetId.Caption = nItem & "��"
End Sub


