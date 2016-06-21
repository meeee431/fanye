VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmModifyTicketType 
   BackColor       =   &H00FFFFFF&
   Caption         =   "修改票种属性"
   ClientHeight    =   2715
   ClientLeft      =   2460
   ClientTop       =   2640
   ClientWidth     =   6795
   Icon            =   "frmModifyTicketType.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2715
   ScaleWidth      =   6795
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   " 关闭(&C)"
      Height          =   345
      Left            =   5580
      TabIndex        =   10
      Top             =   630
      Width           =   1065
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   " 保存(&S)"
      Default         =   -1  'True
      Height          =   345
      Left            =   5580
      TabIndex        =   9
      Top             =   120
      Width           =   1065
   End
   Begin VB.TextBox txtAnnotation 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3240
      TabIndex        =   8
      Top             =   2340
      Width           =   2205
   End
   Begin VB.TextBox txtTicketTypeName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3240
      TabIndex        =   7
      Top             =   1965
      Width           =   2205
   End
   Begin VB.TextBox txtTicketTypeID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   990
      TabIndex        =   6
      Top             =   1965
      Width           =   1005
   End
   Begin VB.OptionButton optNotUsed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "不使用"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1020
      TabIndex        =   5
      Top             =   2355
      Width           =   885
   End
   Begin VB.OptionButton optUsed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "使用"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2355
      Width           =   705
   End
   Begin MSComctlLib.ListView lvTicketType 
      Height          =   1755
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   3096
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "票种代码"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "票种名称"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "使用标记"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "注释"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "注意：全票不允许修改，免票不可以改名，但可以设置是否可用属性，其余票种属性均可以任意修改。"
      Height          =   1455
      Left            =   5580
      TabIndex        =   11
      Top             =   1170
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "票种代码:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   2010
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "注释:"
      Height          =   180
      Index           =   1
      Left            =   2310
      TabIndex        =   2
      Top             =   2385
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "票种名称:"
      Height          =   180
      Left            =   2310
      TabIndex        =   1
      Top             =   2010
      Width           =   810
   End
End
Attribute VB_Name = "frmModifyTicketType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim atTicketType() As TTicketType
Const ID_TicketTypeName = 1
Const ID_IsValid = 2
Const ID_Annotation = 3
Const SM_CanUsed = "可用"
Const SM_CannotUsed = "不可用"

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Private Sub cmdSave_Click()

Dim oTicketType As New TicketType
    On Error GoTo ErrorHandle
    oTicketType.Init g_oActUser
    oTicketType.Identify txtTicketTypeID.Text
    oTicketType.TicketTypeName = GetUnicodeBySize(Trim(txtTicketTypeName.Text), 12)
    oTicketType.TicketTypeValid = IIf(optUsed.Value, TP_TicketTypeValid, TP_TicketTypeNotValid)
    oTicketType.Annotation = txtAnnotation.Text
    oTicketType.Update
    With lvTicketType.SelectedItem
        .ListSubItems(ID_TicketTypeName) = oTicketType.TicketTypeName
        .ListSubItems(ID_IsValid) = IIf(oTicketType.TicketTypeValid = TP_TicketTypeValid, SM_CanUsed, SM_CannotUsed)
        .ListSubItems(ID_Annotation) = oTicketType.Annotation
    End With
    MsgBox "修改保存成功", vbInformation, Me.Caption
    Set oTicketType = Nothing
    
    Exit Sub
ErrorHandle:
    MsgBox err.Description, vbCritical, err.Number
    
End Sub

Private Sub Form_Load()

Dim liTemp As ListItem
Dim nCount As Integer
Dim i As Integer
    On Error GoTo ErrorHandle
    '填充所有的票种
    atTicketType = g_oSysParam.GetAllTicketType
    nCount = ArrayLength(atTicketType)
    For i = 1 To nCount
        
        Set liTemp = lvTicketType.ListItems.Add(, , atTicketType(i).nTicketTypeID)
        liTemp.ListSubItems.Add , , Trim(atTicketType(i).szTicketTypeName)
        liTemp.ListSubItems.Add , , IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, SM_CanUsed, SM_CannotUsed)
        liTemp.ListSubItems.Add , , atTicketType(i).szAnnotation
        
    Next i
    cmdSave.Enabled = False
    Exit Sub
ErrorHandle:
    MsgBox err.Description, vbCritical, err.Number
    
End Sub

Private Sub lvTicketType_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '显示可修改的明细
    txtTicketTypeID.Text = Item.Text
    txtTicketTypeName.Text = Item.ListSubItems(ID_TicketTypeName).Text
    If Item.ListSubItems(ID_IsValid).Text = SM_CanUsed Then

        optNotUsed.Enabled = False
        optUsed.Enabled = False
        optUsed.Value = True
    Else
        optNotUsed.Enabled = True
        optUsed.Enabled = True
        optNotUsed.Value = True
    End If
    txtAnnotation.Text = Item.ListSubItems(ID_Annotation).Text
    SetEnabled
    
End Sub

Private Sub SetEnabled()
    '设置是否可修改
    txtTicketTypeName.Enabled = True
    Select Case txtTicketTypeID.Text
    Case TP_FullPrice
        txtTicketTypeName.Enabled = False
        optUsed.Enabled = False
        optNotUsed.Enabled = False
        
    Case TP_FreeTicket
        txtTicketTypeName.Enabled = False
        
    End Select
    If txtTicketTypeID.Text <> "" Then cmdSave.Enabled = True
End Sub

