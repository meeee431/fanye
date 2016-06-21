VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10605
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   ScaleHeight     =   10605
   ScaleWidth      =   14760
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame11 
      Caption         =   "把预定信息转为预留"
      Height          =   1935
      Left            =   9960
      TabIndex        =   91
      Top             =   8160
      Width           =   4935
      Begin VB.TextBox Text38 
         Height          =   375
         Left            =   3360
         TabIndex        =   95
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text37 
         Height          =   375
         Left            =   1080
         TabIndex        =   94
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command12 
         Caption         =   "转换"
         Height          =   375
         Left            =   3000
         TabIndex        =   93
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text36 
         Height          =   855
         Left            =   120
         TabIndex        =   92
         Top             =   1080
         Width           =   4695
      End
      Begin VB.Label Label38 
         Caption         =   "密码"
         Height          =   255
         Left            =   2640
         TabIndex        =   98
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label37 
         Caption         =   "手机号"
         Height          =   255
         Left            =   360
         TabIndex        =   97
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label36 
         Caption         =   "返回"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "列预定信息"
      Height          =   1935
      Left            =   9960
      TabIndex        =   83
      Top             =   6000
      Width           =   4935
      Begin VB.TextBox Text35 
         Height          =   855
         Left            =   120
         TabIndex        =   87
         Top             =   1080
         Width           =   4695
      End
      Begin VB.CommandButton Command11 
         Caption         =   "查询"
         Height          =   375
         Left            =   3000
         TabIndex        =   86
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text34 
         Height          =   375
         Left            =   1080
         TabIndex        =   85
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text33 
         Height          =   375
         Left            =   3360
         TabIndex        =   84
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label35 
         Caption         =   "返回"
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label34 
         Caption         =   "手机号"
         Height          =   255
         Left            =   360
         TabIndex        =   89
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label33 
         Caption         =   "密码"
         Height          =   255
         Left            =   2640
         TabIndex        =   88
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "列网上售票信息"
      Height          =   2895
      Left            =   9960
      TabIndex        =   75
      Top             =   0
      Width           =   4935
      Begin VB.TextBox Text32 
         Height          =   375
         Left            =   3360
         TabIndex        =   79
         Text            =   "1"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text31 
         Height          =   375
         Left            =   1080
         TabIndex        =   78
         Text            =   "1234"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command10 
         Caption         =   "查询"
         Height          =   375
         Left            =   360
         TabIndex        =   77
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text30 
         Height          =   855
         Left            =   120
         TabIndex        =   76
         Top             =   1800
         Width           =   4695
      End
      Begin VB.Label Label32 
         Caption         =   "密码"
         Height          =   255
         Left            =   2640
         TabIndex        =   82
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label31 
         Caption         =   "取票号"
         Height          =   255
         Left            =   360
         TabIndex        =   81
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label30 
         Caption         =   "返回"
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "取网上售票"
      Height          =   2895
      Left            =   9960
      TabIndex        =   67
      Top             =   3000
      Width           =   4935
      Begin VB.TextBox Text41 
         Height          =   375
         Left            =   1080
         TabIndex        =   103
         Text            =   "1111"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text40 
         Height          =   375
         Left            =   3600
         TabIndex        =   101
         Text            =   "1111"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text29 
         Height          =   855
         Left            =   120
         TabIndex        =   73
         Top             =   1800
         Width           =   4695
      End
      Begin VB.CommandButton Command9 
         Caption         =   "取票"
         Height          =   375
         Left            =   1200
         TabIndex        =   72
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox Text28 
         Height          =   375
         Left            =   1080
         TabIndex        =   70
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text27 
         Height          =   375
         Left            =   3360
         TabIndex        =   68
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label41 
         Caption         =   "身份正号"
         Height          =   255
         Left            =   360
         TabIndex        =   104
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label40 
         Caption         =   "票号"
         Height          =   255
         Left            =   2880
         TabIndex        =   102
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label29 
         Caption         =   "返回"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label28 
         Caption         =   "取票号"
         Height          =   255
         Left            =   360
         TabIndex        =   71
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label27 
         Caption         =   "密码"
         Height          =   255
         Left            =   2640
         TabIndex        =   69
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "0.注册"
      Height          =   615
      Left            =   120
      TabIndex        =   61
      Top             =   0
      Width           =   9615
      Begin VB.TextBox Text26 
         Height          =   375
         Left            =   840
         TabIndex        =   64
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "注册"
         Height          =   375
         Left            =   2880
         TabIndex        =   63
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox Text25 
         Height          =   390
         Left            =   5760
         TabIndex        =   62
         Top             =   120
         Width           =   3495
      End
      Begin VB.Label Label26 
         Caption         =   "用户名"
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label25 
         Caption         =   "返回"
         Height          =   255
         Left            =   5040
         TabIndex        =   65
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "6.取消锁定"
      Height          =   1695
      Left            =   120
      TabIndex        =   50
      Top             =   8760
      Width           =   9495
      Begin VB.TextBox Text24 
         Height          =   375
         Left            =   6120
         TabIndex        =   58
         Text            =   "01|02"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text23 
         Height          =   855
         Left            =   960
         TabIndex        =   54
         Top             =   720
         Width           =   8055
      End
      Begin VB.CommandButton Command6 
         Caption         =   "取消锁定座位"
         Height          =   375
         Left            =   7920
         TabIndex        =   53
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text22 
         Height          =   375
         Left            =   3480
         TabIndex        =   52
         Text            =   "30204"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text21 
         Height          =   375
         Left            =   960
         TabIndex        =   51
         Text            =   "2011-09-23"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label24 
         Caption         =   "座位号"
         Height          =   255
         Left            =   5160
         TabIndex        =   59
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "返回"
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label22 
         Caption         =   "车次号"
         Height          =   255
         Left            =   2640
         TabIndex        =   56
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "发车日期"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "5.售出"
      Height          =   1695
      Left            =   120
      TabIndex        =   34
      Top             =   7080
      Width           =   9495
      Begin VB.TextBox Text39 
         Height          =   375
         Left            =   4560
         TabIndex        =   99
         Text            =   "1111"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text20 
         Height          =   375
         Left            =   6120
         TabIndex        =   48
         Text            =   "01|02"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text19 
         Height          =   375
         Left            =   960
         TabIndex        =   41
         Text            =   "2011-09-23"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text18 
         Height          =   375
         Left            =   3480
         TabIndex        =   40
         Text            =   "720000347"
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "售出锁定座位"
         Height          =   375
         Left            =   7800
         TabIndex        =   39
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text17 
         Height          =   375
         Left            =   960
         TabIndex        =   38
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox Text16 
         Height          =   375
         Left            =   6120
         TabIndex        =   37
         Text            =   "bz"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text15 
         Height          =   375
         Left            =   960
         TabIndex        =   36
         Text            =   "2"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   3480
         TabIndex        =   35
         Text            =   "30204"
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label39 
         Caption         =   "票号"
         Height          =   255
         Left            =   3840
         TabIndex        =   100
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "座位号"
         Height          =   255
         Left            =   5160
         TabIndex        =   49
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "发车日期"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "站点代码"
         Height          =   255
         Left            =   2640
         TabIndex        =   46
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "返回"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "起站点代码"
         Height          =   255
         Left            =   5160
         TabIndex        =   44
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "售票张数"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "车次号"
         Height          =   255
         Left            =   2520
         TabIndex        =   42
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "4.锁定座位"
      Height          =   1695
      Left            =   120
      TabIndex        =   20
      Top             =   5400
      Width           =   9495
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   3480
         TabIndex        =   32
         Text            =   "30204"
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   960
         TabIndex        =   30
         Text            =   "2"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   6120
         TabIndex        =   28
         Text            =   "bz"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   960
         TabIndex        =   24
         Top             =   1200
         Width           =   8055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "锁定座位"
         Height          =   375
         Left            =   7800
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   3480
         TabIndex        =   22
         Text            =   "720000347"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   960
         TabIndex        =   21
         Text            =   "2011-09-23"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "车次号"
         Height          =   255
         Left            =   2520
         TabIndex        =   33
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "售票张数"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "起站点代码"
         Height          =   255
         Left            =   5160
         TabIndex        =   29
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "返回"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "站点代码"
         Height          =   255
         Left            =   2640
         TabIndex        =   26
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "发车日期"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "3.取车次"
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   9495
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   960
         TabIndex        =   16
         Text            =   "2011-09-23"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   3480
         TabIndex        =   15
         Text            =   "720000347"
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "取车次"
         Height          =   375
         Left            =   5280
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   855
         Left            =   960
         TabIndex        =   13
         Top             =   720
         Width           =   8055
      End
      Begin VB.Label Label7 
         Caption         =   "发车日期"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "站点代码"
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "返回"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "2.取站点"
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   9495
      Begin VB.CommandButton Command7 
         Caption         =   "取检票口"
         Height          =   375
         Left            =   6840
         TabIndex        =   60
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "取站点"
         Height          =   375
         Left            =   5040
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Height          =   855
         Left            =   960
         TabIndex        =   9
         Top             =   720
         Width           =   8055
      End
      Begin VB.Label Label4 
         Caption         =   "返回"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "1.登陆"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   9495
      Begin VB.TextBox Text3 
         Height          =   390
         Left            =   960
         TabIndex        =   7
         Top             =   720
         Width           =   8055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "登陆"
         Height          =   375
         Left            =   5040
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3360
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Text            =   "BZZZ"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "返回"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "密码"
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "用户名"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oSell As New SelfSell

Private Sub Command1_Click()
Dim szreturn As String
szreturn = oSell.Login(Text1.Text, Text2.Text)
Text3.Text = szreturn
End Sub

Private Sub Command10_Click()
Dim szreturn As String
Dim pszReturn As String
szreturn = oSell.ListNetTicket(Text31.Text, Text32.Text, "", pszReturn)

Text30.Text = "返回状态:" & szreturn & "返回数据:" & pszReturn
End Sub

Private Sub Command11_Click()
Dim szreturn As String
Dim pszReturn As String
szreturn = oSell.ListBook(Text34.Text, Text33.Text, pszReturn)

Text35.Text = "返回状态:" & szreturn & "返回数据:" & pszReturn
End Sub

Private Sub Command12_Click()
Dim szreturn As String
Dim pszReturn As String
szreturn = oSell.SetBookChange(Text37.Text, Text38.Text)

Text36.Text = "返回状态:" & szreturn & "返回数据:" & pszReturn
End Sub

Private Sub Command2_Click()
Dim szreturn As String
Dim pszReturn As String
szreturn = oSell.GetStation(pszReturn)
Text4.Text = "返回状态:" & szreturn & "返回数据:" & pszReturn

End Sub

Private Sub Command3_Click()
Dim szreturn As String
Dim pszReturn As String
szreturn = oSell.GetBus(Text7.Text, Text6.Text, pszReturn)

Text5.Text = "返回状态:" & szreturn & "返回数据:" & pszReturn

End Sub

Private Sub Command4_Click()
Dim szreturn As String
Dim pszReturn As String
szreturn = oSell.BookTicket(Text14.Text, Text8.Text, Text11.Text, Text9.Text, Text12.Text, pszReturn)

Text10.Text = "返回状态:" & szreturn & "返回数据:" & pszReturn

End Sub

Private Sub Command5_Click()
Dim szreturn As String
Dim pszReturn As String
szreturn = oSell.SetSell(Text13.Text, Text19.Text, Text16.Text, Text18.Text, Text15.Text, Text20.Text, Text39.Text, pszReturn)

Text17.Text = "返回状态:" & szreturn & "返回数据:" & pszReturn

End Sub

Private Sub Command6_Click()
Dim szreturn As String
Dim pszReturn As String
szreturn = oSell.UnBookTicket(Text22.Text, Text21.Text, Text24.Text)

Text23.Text = "返回状态:" & szreturn

End Sub

Private Sub Command7_Click()
Dim szreturn As String
Dim pszReturn As String
szreturn = oSell.GetCheckGate(pszReturn)
Text4.Text = "返回状态:" & szreturn & "返回数据:" & pszReturn

End Sub

Private Sub Command8_Click()
Dim szreturn As String
szreturn = oSell.RegisterCode(Text26.Text)
Text25.Text = szreturn
End Sub

Private Sub Command9_Click()
Dim szreturn As String
Dim pszReturn As String
szreturn = oSell.GetNetPrint(Text28.Text, Text27.Text, Text41.Text, Text40.Text, pszReturn)

Text29.Text = "返回状态:" & szreturn & "返回数据:" & pszReturn

End Sub

