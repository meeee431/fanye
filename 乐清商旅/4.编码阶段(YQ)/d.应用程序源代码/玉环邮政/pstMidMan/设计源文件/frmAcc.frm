VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmAcc 
   Caption         =   "������Ʊ���ʵ�"
   ClientHeight    =   5355
   ClientLeft      =   1110
   ClientTop       =   330
   ClientWidth     =   7620
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   7620
   Begin VB.TextBox txtOperatorID 
      Height          =   288
      Left            =   5160
      TabIndex        =   22
      Text            =   " "
      Top             =   3204
      Width           =   1212
   End
   Begin VB.TextBox txtBankID 
      Height          =   288
      Left            =   2976
      TabIndex        =   17
      Text            =   " "
      Top             =   3204
      Width           =   1212
   End
   Begin VB.TextBox txtOpDate 
      Height          =   288
      Left            =   792
      TabIndex        =   16
      Text            =   " "
      Top             =   3204
      Width           =   1212
   End
   Begin VB.TextBox CancelMoney 
      Height          =   288
      Left            =   5160
      TabIndex        =   15
      Text            =   " "
      Top             =   4200
      Width           =   1212
   End
   Begin VB.TextBox ValidMoney 
      Height          =   288
      Left            =   2976
      TabIndex        =   13
      Top             =   4182
      Width           =   1212
   End
   Begin VB.TextBox SumMoney 
      Height          =   288
      Left            =   792
      TabIndex        =   11
      Text            =   " "
      Top             =   4182
      Width           =   1212
   End
   Begin VB.TextBox CancelNos 
      Height          =   288
      Left            =   5160
      TabIndex        =   9
      Top             =   3720
      Width           =   1212
   End
   Begin VB.TextBox ValidNos 
      Height          =   288
      Left            =   2976
      TabIndex        =   7
      Text            =   " "
      Top             =   3702
      Width           =   1212
   End
   Begin VB.TextBox SumNos 
      DataSource      =   """select count(amount) from dailybook where tradeid='8001'"
      Height          =   288
      Left            =   792
      TabIndex        =   5
      Text            =   " "
      Top             =   3720
      Width           =   1212
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   7620
      TabIndex        =   1
      Top             =   4725
      Width           =   7620
      Begin VB.CommandButton cmdDailyList 
         Caption         =   "������ˮ�ļ�"
         Height          =   252
         Left            =   2520
         TabIndex        =   20
         Top             =   0
         Width           =   1212
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "��ӡ"
         Height          =   252
         Left            =   720
         TabIndex        =   3
         Top             =   0
         Width           =   1092
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "�ر�"
         Height          =   300
         Left            =   4675
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Align           =   1  'Align Top
      Bindings        =   "frmAcc.frx":0000
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   12
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "BankID"
         Caption         =   "����"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "H:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "OperatorID"
         Caption         =   "����"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "opDate"
         Caption         =   "����ʱ��"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "mm.dd hh:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "TradeID"
         Caption         =   "������"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "TicketID"
         Caption         =   "Ʊ��"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Amount"
         Caption         =   "����"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "sumMoney"
         Caption         =   "���(Ԫ)"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   929.764
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   5025
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "PROVIDER=MSDASQL;dsn=sx;uid=sa;pwd=;"
      OLEDBString     =   "PROVIDER=MSDASQL;dsn=sx;uid=sa;pwd=;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select opDate,BankID,OperatorID,TradeID,TicketID,Amount,sumMoney  from DailyBook Order by BankID"
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label9 
      Caption         =   "����"
      Height          =   252
      Left            =   4488
      TabIndex        =   21
      Top             =   3240
      Width           =   612
   End
   Begin VB.Label Label8 
      Caption         =   "����"
      Height          =   252
      Left            =   2280
      TabIndex        =   19
      Top             =   3240
      Width           =   612
   End
   Begin VB.Label Label7 
      Caption         =   "����"
      Height          =   252
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   612
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7680
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label6 
      Caption         =   "���(Ԫ)"
      Height          =   252
      Left            =   4368
      TabIndex        =   14
      Top             =   4200
      Width           =   732
   End
   Begin VB.Label Label5 
      Caption         =   "���(Ԫ)"
      Height          =   252
      Left            =   2160
      TabIndex        =   12
      Top             =   4200
      Width           =   732
   End
   Begin VB.Label Label4 
      Caption         =   "���(Ԫ)"
      Height          =   252
      Left            =   0
      TabIndex        =   10
      Top             =   4200
      Width           =   732
   End
   Begin VB.Label Label3 
      Caption         =   "��Ʊ"
      Height          =   252
      Left            =   4488
      TabIndex        =   8
      Top             =   3720
      Width           =   612
   End
   Begin VB.Label Label2 
      Caption         =   "�۳�Ʊ"
      Height          =   252
      Left            =   2160
      TabIndex        =   6
      Top             =   3720
      Width           =   612
   End
   Begin VB.Label Label1 
      Caption         =   "Ʊ����"
      Height          =   252
      Left            =   0
      TabIndex        =   4
      Top             =   3720
      Width           =   612
   End
End
Attribute VB_Name = "frmAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Private Const ACCFILE = "c:\sxicbcbus\account.txt"
Private Const FILEPROMPT = "��ˮ�ļ�������" & ACCFILE & "."
Private Const sqlAccheader = "select  opDate,BankID,operatorID,tradeid,ticketID,amount,summoney from dailybook  where (tradeid='8001' or tradeid='8011' or tradeid='9999') and tradeok='OK'"
Private Const sqlAcctail = " order by bankiD,operatorid"
 
Private Sub cmdDailyList_Click()
  Dim FileNo As Integer
  Dim tmpBankID1 As String
  Dim tmpBankID2 As String
  Dim tmpSumMoney As Currency
  Dim tmpValidMoney As Currency
  Dim tmpCancelMoney As Currency
  Dim tmpSumNos As Integer
  Dim tmpValidNos As Integer
  Dim tmpCancelNos As Integer
  tmpBankID1 = "": tmpBankID2 = ""
  tmpSumMoney = 0: tmpCancelMoney = 0: tmpSumNos = 0: tmpCancelNos = 0
  FileNo = FreeFile
  Open App.Path + "\Account.txt" For Output As #FileNo
  txtBankID = "": txtOperatorID = ""
  AccQuery
  With datPrimaryRS.Recordset
    Do While Not .EOF
      tmpBankID1 = !bankId
      If tmpBankID1 <> tmpBankID2 And tmpBankID2 <> "" Then
        tmpValidMoney = tmpSumMoney - tmpCancelMoney
        tmpValidNos = tmpSumNos - tmpCancelNos
        Print #FileNo, String(80, "-")
        Print #FileNo, "Ʊ����(��): "; Str(tmpSumNos); vbTab; "�۳�Ʊ(��): "; _
          Str(tmpValidNos); vbTab; "�ܽ��(Ԫ): "; Str(tmpSumMoney); vbTab; _
          "ʵ�ս��(Ԫ): "; Str(tmpValidMoney)
        Print #FileNo, String(80, "-")
        tmpSumMoney = 0: tmpCancelMoney = 0: tmpSumNos = 0: tmpCancelNos = 0
      End If
      tmpBankID2 = tmpBankID1
      If !TradeID = "8001" Then
        tmpSumMoney = tmpSumMoney + !SumMoney
        tmpSumNos = tmpSumNos + !Amount
      Else
        tmpCancelMoney = tmpCancelMoney + !SumMoney
        tmpCancelNos = tmpCancelNos + 1
      End If
      Print #FileNo, !bankId; " "; !operatorid; " "; !TradeID; _
        " "; Format(Str(!SumMoney), "####0.00"); " "; "0"; " "; "0"; " ";
      If !TradeID = "8011" Then
        Print #FileNo, "��Ʊ";
      Else
        Print #FileNo, "��Ʊ";
      End If
      Print #FileNo, " "; !ticketID; " "; "��������Ʊ"; " "; _
        "0"; " "; "0"; " "; txtOpDate; " "; txtOpDate
      .MoveNext
    Loop
  End With
  If tmpSumMoney <> 0 Then
    tmpValidMoney = tmpSumMoney - tmpCancelMoney
    tmpValidNos = tmpSumNos - tmpCancelNos
    Print #FileNo, String(80, "-")
    Print #FileNo, "Ʊ����(��): "; Str(tmpSumNos); vbTab; "�۳�Ʊ(��): "; _
    Str(tmpValidNos); vbTab; "�ܽ��(Ԫ): "; Str(tmpSumMoney); vbTab; _
      "ʵ�ս��(Ԫ): "; Str(tmpValidMoney)
        Print #FileNo, String(80, "-")
        tmpSumMoney = 0: tmpCancelMoney = 0: tmpSumNos = 0: tmpCancelNos = 0
  End If
  Close #FileNo
  MsgBox FILEPROMPT
End Sub

Private Sub cmdPrint_Click()
  Dim i As Integer
  i = Printer.FontSize
  Printer.FontSize = i * 2
  Printer.Print "���˹�������������Ʊ���ʵ�" & vbTab & Format(CDate(txtOpDate), "yyyy.mm.dd")
  Printer.FontSize = 10
  Printer.Print String(85, "-")
  Printer.Print "����" & vbTab & "����Ա��" & vbTab & vbTab & "����ʱ��" & vbTab & _
    vbTab & "������" & vbTab & "Ʊ��" & vbTab & "����" & vbTab & "���(Ԫ)"
  AccQuery
  With datPrimaryRS.Recordset
    Do While Not .EOF
      Printer.Print !bankId & vbTab & !operatorid & vbTab & _
        vbTab & Format(!opDate, "hh:mm") & vbTab & vbTab & !TradeID & vbTab & _
        !ticketID & vbTab & Str(!Amount) & vbTab & Format(Str(!SumMoney), "####0.00")
       .MoveNext
    Loop
  End With
  Printer.Print String(85, "-")
  Printer.FontSize = 12
  Printer.Print "Ʊ����:" & SumNos & "(��)" & vbTab & vbTab & "�۳�Ʊ:" & _
    ValidNos & "(��)" & vbTab & vbTab & "��Ʊ:" & CancelNos & vbTab & "(��)"
  Printer.Print "�ܽ��:" & SumMoney & "(Ԫ)" & vbTab & "ʵ�ս��:" & _
    ValidMoney & "(Ԫ)" & vbTab & "��Ʊ���:" & CancelMoney & vbTab & "(Ԫ)"
  Printer.EndDoc
  Printer.FontSize = i
End Sub

Private Sub Form_Load()
  Dim i As Integer
  txtOpDate = Format(Now, "yyyy/mm/dd")
 End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'This will resize the grid when the form is resized
  grdDataGrid.Height = grdDataGrid.RowHeight * 17
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Description
End Sub
'
'Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'  datPrimaryRS.Caption = "���ʼ�¼: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
'End Sub

'Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'  'This is where you put validation code
'  'This event gets called when the following actions occur
'  Dim bCancel As Boolean
'  Select Case adReason
'  Case adRsnAddNew
'  Case adRsnClose
'  Case adRsnDelete
'  Case adRsnFirstChange
'  Case adRsnMove
'  Case adRsnRequery
'  Case adRsnResynch
'  Case adRsnUndoAddNew
'  Case adRsnUndoDelete
'  Case adRsnUndoUpdate
'  Case adRsnUpdate
'  End Select
'
'  If bCancel Then adStatus = adStatusCancel
'End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  datPrimaryRS.Recordset.MoveLast
  grdDataGrid.SetFocus
  SendKeys "{down}"

  Exit Sub
AddErr:
  MsgBox err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
  Exit Sub
RefreshErr:
  MsgBox err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  MsgBox err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Option1_Click(Index As Integer)

End Sub

 
Private Sub StationSel_Click(Index As Integer)
'  AccQuery
'  AccStatistic datPrimaryRS.Recordset
'  datPrimaryRS.Refresh
End Sub
Private Sub txtOpDate_Change()
  If Len(Trim(txtOpDate.Text)) < 6 Then
    Exit Sub
  End If
  AccQuery
  AccStatistic datPrimaryRS.Recordset
  datPrimaryRS.Refresh
End Sub
Private Sub txtBankID_Change()
  If Len(Trim(txtBankID)) < 5 Then
    Exit Sub
  End If
  AccQuery
  AccStatistic datPrimaryRS.Recordset
  datPrimaryRS.Refresh
End Sub
Private Sub txtOperatorID_Change()
  If Len(Trim(txtOperatorID)) < 5 Then
    Exit Sub
  End If
  AccQuery
  AccStatistic datPrimaryRS.Recordset
  datPrimaryRS.Refresh
End Sub
Private Sub AccQuery()
  Dim tmpDate As Date
  Dim szStr As String
  Dim tmpsqlStr As String
  Dim i As Integer
  tmpsqlStr = sqlAccheader
  If Len(Trim(txtOpDate)) > 5 Then
    On Error Resume Next
    tmpDate = CDate(Trim(txtOpDate.Text))
    If err.Number <> 0 Then
      tmpDate = Now
      On Error GoTo txtBankIDErrHandle
    End If
    szStr = Format(tmpDate, "yyyymmdd")
    tmpsqlStr = tmpsqlStr & " and convert(char(10),opdate,112)='" & szStr & "'"
  End If
  If Len(Trim(txtBankID.Text)) = 5 Then
    tmpsqlStr = tmpsqlStr & " and BankId='" & Trim(txtBankID.Text) & "'"
  End If
  If Len(Trim(txtOperatorID.Text)) = 5 Then
    tmpsqlStr = tmpsqlStr & " and OperatorId='" & Trim(txtOperatorID.Text) & "'"
  End If
  tmpsqlStr = tmpsqlStr & sqlAcctail
  datPrimaryRS.RecordSource = tmpsqlStr
  datPrimaryRS.Refresh
  If datPrimaryRS.Recordset.RecordCount <> 0 Then
    datPrimaryRS.Recordset.MoveFirst
  End If
txtBankIDErrHandle:
End Sub
Public Sub AccStatistic(tmpRec As ADODb.Recordset)
  Dim iSumNos As Integer
  Dim iCancelNos As Integer
  Dim iValidNos As Integer
  Dim cSumMoney As Currency
  Dim cCancelMoney As Currency
  Dim cValidMoney As Currency
  iSumNos = 0: iCancelNos = 0: iValidNos = 0
  cSumMoney = 0: cCancelMoney = 0: cValidMoney = 0
  If tmpRec Is Nothing Then
     Exit Sub
  End If
  With tmpRec
    If Not .BOF And .RecordCount > 0 Then
      .MoveFirst
    End If
    Do While Not .EOF
      If !TradeID = "8001" Then
        iSumNos = iSumNos + !Amount
        cSumMoney = cSumMoney + !SumMoney
      Else
        iCancelNos = iCancelNos + !Amount
        cCancelMoney = cCancelMoney + !SumMoney
      End If
      .MoveNext
    Loop
  End With
  iValidNos = iSumNos - iCancelNos
  cValidMoney = cSumMoney - cCancelMoney
  SumNos = Trim(Str(iSumNos)): ValidNos = Trim(Str(iValidNos)): CancelNos = Trim(Str(iCancelNos))
  SumMoney = Trim(Str(cSumMoney)): ValidMoney = Trim(Str(cValidMoney)): CancelMoney = Trim(Str(cCancelMoney))
End Sub


