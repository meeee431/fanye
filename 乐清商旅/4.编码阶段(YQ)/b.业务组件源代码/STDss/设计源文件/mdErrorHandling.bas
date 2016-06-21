Attribute VB_Name = "mdErrorLib"
'=========================================================================================
'Author:
'Detail:中间层共用模块
'=========================================================================================


'=========================================================================================
'以下常量声明
'=========================================================================================
'系统常量定义
'---------------------------------------------------------
Public Const cszDefErrMsg = "由于在资源文件中找不到相应的资源，所以此错误的描述为空。"
Public Const cszOrgErrMsgNote = "[源错误描述]:"
Public Const cszCustomErrSource = "自定义错误"

'对象错误公共声明
'---------------------------------------------------------
Public Const ERR_CoreErrorStart = 10000      '核心错误号起点
Public Const ERR_DBObjErrorStart = 10500     '数据对象错误号起点
Public Const ERR_BusObjErrorStart = 11000    '业务组件错误号起点
Public Const ERR_ClientToolStart = 12000    '客户端工具错误号起点
Public Const ERR_DLLStep = 50               '每个对象可分配的错误号范围

'各实体类的新增、修改、删除的错误号位移（其真正的错误号要加上它的错误起始号）
Public Enum ENormalOffset
    ERR_AddDuplicate = 1  '新增对象时指定的对象已经存在
    ERR_AddNoParent = 2  '新增对象时对象所要求的父不存在
    
    ERR_EditNoParent = 4  '修改对象时对象所要求的父不存在
    ERR_EditChildExist = 5  '修改对象时有依赖它的子存在
    
    ERR_DeleteChildExist = 7  '删除对象时它有子存在
    ERR_DeleteNotExist = 8  '删除对象时无指定的对象
    
    '---------------------------------------------
    ERR_NoActiveUser = 11     '对象还未设置活动用户对象
    ERR_NotAvailable = 12     '对象现在处于不可用状态
    ERR_EditObj = 13    '对象现在处于修改状态
    ERR_AddObj = 14     '对象现在处于新增状态
    ERR_NormalObj = 15     '对象现在处于正常状态
    ERR_NotAddObj = 16    '对象现在不处于新增状态
End Enum

Public Enum EErrDatabase_Trigger    '触发器错误
    ERR_DBAddNoParent = 30002 '插入时父不存在
    
    ERR_DBEditNoParent = 30003 '父表不存在,子表不能修改
    ERR_DBEditChildExist = 30005 '子表存在,父表不能修改
    
    ERR_DBDeleteChildExist = 30006 '子表存在,不能删除父表
End Enum

'---------------------------------------------------------
Public Const ERR_TransContext_MTSNotSupport = ERR_CoreErrorStart + 0 'MTS不支持!
Public Const ERR_InvalidProgramer = ERR_CoreErrorStart + 1 '未经授权的开发者!
Public Const ERR_AssertCreditCard = ERR_CoreErrorStart + 2 '用户身份无法识别!
Public Const ERR_INIFileNotValid = ERR_CoreErrorStart + 3 '非法系统配置文件!
Public Const ERR_DatabaseConnectErr = ERR_CoreErrorStart + 4   '通用打开数据库错误!
Public Const ERR_DuplicatePrimaryKeyErr = ERR_CoreErrorStart + 5 '该记录已经存在，无法添加!
Public Const ERR_ForeignKeyConflict = ERR_CoreErrorStart + 6    '该记录在其它地方已经被引用，无法修改!
Public Const ERR_ColumnName = ERR_CoreErrorStart + 7  '记录集字段名与数据库不符!
Public Const ERR_IDName = ERR_CoreErrorStart + 8  '更新时，记录集字段中不能给主键附值!
Public Const ERR_ColumnNum = ERR_CoreErrorStart + 9 '记录集字段数与数据库不符!
Public Const ERR_SameRecordExist = ERR_CoreErrorStart + 10  '同名记录已存在!
Public Const ERR_BuiltinRecordNotDelete = ERR_CoreErrorStart + 11  '系统内置记录不允许删除!
Public Const ERR_InterfaceNotUsable = ERR_CoreErrorStart + 12 '接口暂时不可用

    


' *******************************************************************
' *   Brief Description: 通用错误触发程序                           *
' *   Engineer: 陆勇庆                                              *
' *   Date Generated: 2002/01/20                                    *
' *   Last Revision Date:                                           *
' *******************************************************************
Public Sub RaiseError(ByVal plErrNum As Long, Optional ByVal pszSource As String = "", Optional ByVal pszOrgMsg As String = "", Optional ByVal pbWantOrgMsg As Boolean = False)
    Dim szErrMsg As String, szSource As String
'    On Error GoTo Here
On Error Resume Next
    If pszSource <> "" Then
        szSource = pszSource
    Else
        szSource = err.Source
    End If
    
    '得到预定义的错误号
    szErrMsg = LoadResString(plErrNum)
    On Error GoTo 0
    If szErrMsg = "" Then
        If pszOrgMsg <> "" Then
            If Left(pszOrgMsg, Len(cszOrgErrMsgNote)) <> cszOrgErrMsgNote Then
                szErrMsg = cszOrgErrMsgNote & pszOrgMsg
            Else
                szErrMsg = pszOrgMsg
            End If
        Else
            szErrMsg = cszDefErrMsg
            If pbWantOrgMsg Then szErrMsg = szErrMsg & vbCrLf & cszOrgErrMsgNote & pszOrgMsg
        End If
    End If
     
'    '得到系统设置是否显示原始错误信息，用于NOTES等无法得到组件错误的语言中
'    If Left(szErrMsg, 2) <> "错误" Then
'        szErrMsg = "错误 " & plErrNum & vbCrLf & szErrMsg
'    End If
    err.Raise plErrNum, szSource, szErrMsg
End Sub

' *******************************************************************
' *   Brief Description: 通用错误触发程序(兼容旧版系统)             *
' *   Engineer: 陆勇庆                                              *
' *   Date Generated: 2002/01/20                                    *
' *   Last Revision Date:                                           *
' *******************************************************************
Public Sub ShowError(ByVal plErrNum As Long, Optional ByVal pszSource As String = "", Optional ByVal pszOrgMsg As String = "", Optional ByVal pbWantOrgMsg As Boolean = False)
    Dim szErrMsg As String, szSource As String
'    On Error GoTo Here
On Error Resume Next
    If pszSource <> "" Then
        szSource = pszSource
    Else
        szSource = err.Source
    End If
    
    '得到预定义的错误号
    szErrMsg = LoadResString(plErrNum)
    On Error GoTo 0
    If szErrMsg = "" Then
        If pszOrgMsg <> "" Then
            If Left(pszOrgMsg, Len(cszOrgErrMsgNote)) <> cszOrgErrMsgNote Then
                szErrMsg = cszOrgErrMsgNote & pszOrgMsg
            Else
                szErrMsg = pszOrgMsg
            End If
        Else
            szErrMsg = cszDefErrMsg
            If pbWantOrgMsg Then szErrMsg = szErrMsg & vbCrLf & cszOrgErrMsgNote & pszOrgMsg
        End If
    End If
     
'    '得到系统设置是否显示原始错误信息，用于NOTES等无法得到组件错误的语言中
'    If Left(szErrMsg, 2) <> "错误" Then
'        szErrMsg = "错误 " & plErrNum & vbCrLf & szErrMsg
'    End If
    err.Raise plErrNum, szSource, szErrMsg
End Sub
'验证新增对象错误
Public Sub AssertAddObjectError(ByVal plErrBegin As Long, poDb As RTConnection)
    If poDb.HaveThisNativeErr(ERR_DBAddDuplicate) Then   '如果是主键重错
        RaiseError plErrBegin + ERR_AddDuplicate
    ElseIf poDb.HaveThisNativeErr(ERR_DBAddNoParent) Then '如果是新增无父错误
        RaiseError plErrBegin + ERR_AddNoParent
    Else
        RaiseError err.Number, , err.Description
    End If
End Sub
'验证更改对象错误
Public Sub AssertUpdateObjectError(plErrBegin As Long, poDb As RTConnection)
    If poDb.HaveThisNativeErr(ERR_DBEditChildExist) Then '如果是更新子存在错
        RaiseError plErrBegin + ERR_EditChildExist
    ElseIf poDb.HaveThisNativeErr(ERR_DBEditNoParent) Then '如果是更新父不存在错
        RaiseError plErrBegin + ERR_EditNoParent
    Else
        RaiseError err.Number, , err.Description
    End If
End Sub

'验证删除对象错误
Public Sub AssertDeleteObjectError(plErrBegin As Long, poDb As RTConnection)
    If poDb.HaveThisNativeErr(ERR_DBDeleteChildExist) Then '如果是删除子存在错
        RaiseError plErrBegin + ERR_DeleteChildExist
    Else
        RaiseError err.Number, , err.Description
    End If
End Sub
'恢复删除对象错误
Public Sub AssertReCoverObjectError(plErrBegin As Long, poDb As RTConnection)
    If poDb.HaveThisNativeErr(ERR_DBDeleteChildExist) Then '如果是恢复子存在错
        RaiseError plErrBegin + ERR_DeleteChildExist
    Else
        RaiseError err.Number, , err.Description
    End If
End Sub
