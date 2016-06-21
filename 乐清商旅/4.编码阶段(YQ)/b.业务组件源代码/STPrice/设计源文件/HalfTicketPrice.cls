VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HalfTicketPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**********************************************************
'* Source File Name:HalfTicketPrice.cls
'* Project Name:STPrice.vbp
'* Engineer:
'* Data Generated:
'* Last Revision Date:
'* Brief Description:半票价
'* Relational Document:
'**********************************************************

Option Explicit

Public Enum EErrHalfTicketPrice
    ERR_HalfTicketNoItemParam = ERR_HalfTicketPrice + 21 '半价票价项参数不存在
    ERR_HalfTicketUPDATEParamError = ERR_HalfTicketPrice + 22 '半价参数UPDATEToDB错误
    ERR_HalfTicketUPDATEParamIllegal = ERR_HalfTicketPrice + 23 '半价参数UPDATE参数不合法
    ERR_HalfTickerGetParamErrDB = ERR_HalfTicketPrice + 24 '从数据库得到半价参数出错
End Enum

Public Enum ERightHalfTicketPrice
    
    RIGHT_HalfTicketGetModifyParam = ERR_HalfTicketPrice + cnMidRightBegin + cnMidRightStep * 1 'GetModify半价参数
End Enum

'所有的实体类都有的私有类变量
'Private m_nObjectStatus As EObjectStatus
Private m_oActiveUser As ActiveUser

Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
    m_oActiveUser.IAmLive
End Property

Public Property Set SelfUser(vNewValue As ActiveUser)
    Set m_oActiveUser = vNewValue
    m_oActiveUser.IAmLive
End Property

Public Sub Init(poAUser As ActiveUser)
    Set SelfUser = poAUser
End Sub


Public Function GetItemParam(Ticket_Type As Integer, TableId As String, Optional UseMark As EPriceItemUse = TP_PriceItemAll) As THalfTicketItemParam()
    AssertActiveUserValid m_oActiveUser, ERR_HalfTicketPrice
    AssertHaveRight m_oActiveUser, RIGHT_HalfTicketGetModifyParam
    
    Dim szSql As String, lRowCount As Long
    Dim odb As New RTConnection
    Dim i As Integer, rsTemp As Recordset, j As Integer
    Dim aTHTemp() As THalfTicketItemParam
    Dim eTemp As ETicketType
    
    If Ticket_Type = 0 Then
       szSql = "SELECT tcf.* FROM Ticket_type_formula_info tcf,Ticket_type_code tti, price_item_info pii WHERE table_id=" & "'" & TableId & "'" & " AND ticket_type_id=ticket_type AND ticket_type_id<>'" & TP_FreeTicket & "' AND ticket_type_valid='" & TP_TicketTypeValid & "' AND pii.price_item=tcf.price_item "  'ORDER BY ticket_type , tcf.price_item"
    Else
       szSql = "SELECT tcf.* FROM Ticket_type_formula_info tcf,Ticket_type_code tti, price_item_info pii WHERE ticket_Type=" & Ticket_Type & " AND table_id=" & "'" & TableId & "'" & " AND ticket_type_id=ticket_type AND ticket_type_id<>'" & TP_FreeTicket & "' AND ticket_type_valid='" & TP_TicketTypeValid & "' AND pii.price_item=tcf.price_item "  'ORDER BY ticket_type , tcf.price_item"
    End If
    If UseMark <> TP_PriceItemAll Then
       szSql = szSql & " AND use_mark='" & UseMark & "'"
    End If
    szSql = szSql & "ORDER BY ticket_type , tcf.price_item"
    
On Error GoTo Here
    odb.ConnectionString = GetConnectionStr(cszPriceMan)
    Set rsTemp = odb.Execute(szSql, lRowCount)
On Error GoTo 0
    If lRowCount = 0 Then
        
        ShowError ERR_HalfTicketNoItemParam

    Else
        j = rsTemp.RecordCount
        ReDim aTHTemp(1 To j)
        For i = 1 To j
            eTemp = FormatDbValue(rsTemp!Ticket_Type)
            If Ticket_Type = 0 Then
                aTHTemp(i).szTicketType = FormatDbValue(rsTemp!Ticket_Type)
                aTHTemp(i).szTicketItem = FormatDbValue(rsTemp!price_item)
                aTHTemp(i).sgParam1 = FormatDbValue(rsTemp!parameter_1)
                aTHTemp(i).sgParam2 = FormatDbValue(rsTemp!parameter_2)
                aTHTemp(i).szAnnotation = FormatDbValue(rsTemp!Annotation)
            ElseIf eTemp = Ticket_Type Then
                aTHTemp(i).szTicketType = FormatDbValue(rsTemp!Ticket_Type)
                aTHTemp(i).szTicketItem = FormatDbValue(rsTemp!price_item)
                aTHTemp(i).sgParam1 = FormatDbValue(rsTemp!parameter_1)
                aTHTemp(i).sgParam2 = FormatDbValue(rsTemp!parameter_2)
                aTHTemp(i).szAnnotation = FormatDbValue(rsTemp!Annotation)
            End If
            rsTemp.MoveNext
        Next i
     End If
     GetItemParam = aTHTemp
     Set rsTemp = Nothing
     Set odb = Nothing
Exit Function
Here:
    ShowError ERR_HalfTickerGetParamErrDB
    Set rsTemp = Nothing
    Set odb = Nothing
End Function

Public Function ModifyItemParam(Ticket_Type As Integer, HalfTicketParam As THalfTicketItemParam, TableId As String) As Long
    AssertActiveUserValid m_oActiveUser, ERR_HalfTicketPrice
    AssertHaveRight m_oActiveUser, RIGHT_HalfTicketGetModifyParam
    
    Dim szSql As String, odb As New RTConnection, lRowCount As Long
    Dim szTemp1 As String, sgTemp1 As Double, sgTemp2 As Double, szTemp2 As String
    Dim nTemp As Integer

    nTemp = Ticket_Type
    szTemp1 = HalfTicketParam.szTicketItem
    sgTemp1 = HalfTicketParam.sgParam1
    sgTemp2 = HalfTicketParam.sgParam2
    szTemp2 = HalfTicketParam.szAnnotation
    
    If (sgTemp1 > 9999.999 And sgTemp1 < -9999.999) Or (sgTemp2 > 9999.999 And sgTemp2 < -9999.999) Then
        ShowError ERR_HalfTicketUPDATEParamIllegal
    End If
    szSql = "UPDATE Ticket_type_formula_info SET " _
    & "parameter_1=" & Str(sgTemp1) _
    & ",parameter_2=" & Str(sgTemp2) _
    & ",annotation='" & szTemp2 _
    & "' WHERE price_item='" & szTemp1 _
    & "' AND ticket_type=" & Str(nTemp) & " AND table_id='" & TableId & "'"
On Error GoTo Here
    
    odb.ConnectionString = GetConnectionStr(cszPriceMan)
    odb.Execute szSql, lRowCount
    ModifyItemParam = lRowCount
    WriteOperateLog m_oActiveUser, RIGHT_HalfTicketGetModifyParam, "修改特殊票参数（票种代码为" & Str(nTemp) & "票价项为" & szTemp1 & "）的参数"
    Set odb = Nothing
Exit Function

Here:
     ShowError ERR_HalfTicketUPDATEParamError
    Set odb = Nothing
    
End Function



