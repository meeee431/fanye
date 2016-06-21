VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EveryObjClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrEveryObject
    ERR_ERROR1 = ERR_EveryObject + 10
End Enum

Public Enum ERightEveryObject
    RIGHT_Fun1 = ERR_EveryObject + cnMidErrBegin + cnMidErrStep * 1
End Enum



'----------------------
'各类特有的变量
Private m_szObjID As String
Private m_szObjName As String


'---------------------
'所有的实体类都有的私有类变量<<
Private m_nObjectStatus As EObjectStatus

'所有的类都有的私有类变量
Private m_oActiveUser As ActiveUser
'--------------------------->>
Public Sub Init(poAUser As ActiveUser)
    Set SelfUser = poAUser
End Sub

Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
    m_oActiveUser.IAmLive
End Property

Public Property Set SelfUser(vNewValue As ActiveUser)
    If m_nObjectStatus = ST_AddObj Then
        ShowError ERR_AddObj
    ElseIf m_nObjectStatus = ST_EditObj Then
        ShowError ERR_EditObj
    ElseIf m_nObjectStatus = ERR_NormalObj Then
        ShowError ERR_NormalObj
    Else
        Set m_oActiveUser = vNewValue
    End If
    m_oActiveUser.IAmLive
End Property


Private Sub Class_Initialize()
    m_nObjectStatus = ST_NotAvailable '类刚建立时处于不可用状态
End Sub
Private Sub AssertStatusAvailable() '测试对象的状态是否有效（不为无效ECNotAvailable状态）
    If m_nObjectStatus = ST_NotAvailable Then ShowError ERR_NotAvailable
End Sub
Public Sub AssertObjIsValid() '测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
    AssertActiveUserValid m_oActiveUser
    AssertStatusAvailable
End Sub


Public Sub Identify(pszObjID As String)
    AssertActiveUserValid m_oActiveUser
    
    m_szObjID = pszObjID
    RefreshMemoryInfo
    
    m_nObjectStatus = ST_NormalObj
End Sub

Public Sub AddNew() '调用新增方法后，对象处于新增状态。
    AssertActiveUserValid m_oActiveUser
    m_nObjectStatus = ST_AddObj
End Sub

Public Sub Delete() '调用删除方法后，对象处于不可用状态
    AssertObjIsValid
    
    If m_nObjectStatus <> ST_AddObj Then
        DeleteObject
    End If
    
    m_nObjectStatus = ST_NotAvailable
End Sub

Public Sub Update()
    AssertObjIsValid
    
    Select Case m_nObjectStatus
        Case ST_AddObj '如对象处于新增状态，则新增一对象
            AddObject
        Case ST_EditObj '如对象处于修改状态，则更新它
            UpdateToDB
        Case ST_NotAvailable '如果对象不可用则出错
            ShowError ERR_NotAvailable
    End Select
    
    m_nObjectStatus = ST_NormalObj
End Sub


Public Property Get ObjStatus() As EObjectStatus
    ObjStatus = m_nObjectStatus
End Property

Private Sub RefreshMemoryInfo() '按照对象现在的ID将数据库中相应的信息读至对象内
    Dim rsTemp As Recordset
    AssertActiveUserValid m_oActiveUser
    
    '-------------
    MsgBox "你还未写刷新代码"
    If rsTemp.RecordCount <> 1 Then ShowError 5 '指定的东西不存在,5是乱写的应为你自己的代码
End Sub

Private Sub AddObject()
    On Error GoTo here
    
    Exit Sub
here:
    TreatAddError ERR_EveryObject '开始错误代码
End Sub

Private Sub DeleteObject()
    On Error GoTo here
    
    Exit Sub
here:
    TreatDeleteError ERR_EveryObject  '开始错误代码

End Sub

Private Sub UpdateToDB()
    On Error GoTo here
    
    Exit Sub
here:
    TreatUpdateError ERR_EveryObject  '开始错误代码
End Sub


'------------------------
Public Function Function1() As String
    AssertObjIsValid
    AssertHaveRight m_oActiveUser, RIGHT_Fun1
    'do thing
    
    WriteOperateLog m_oActiveUser, RIGHT_Fun1
End Function


Public Property Get Property1() As String
    AssertObjIsValid
    
End Property

Public Property Let Property1(ByVal vNewValue As String)
    AssertObjIsValid
    
    
    If m_nObjectStatus = ST_NormalObj Then
        m_nObjectStatus = ST_EditObj
    End If
End Property

Public Property Get ObjID() As String
    AssertObjIsValid
End Property

Public Property Let ObjID(ByVal vNewValue As String)
    AssertObjIsValid
    If m_nObjectStatus <> ST_AddObj Then ShowError ERR_NotAddObj
    m_szObjID = vNewValue

End Property
