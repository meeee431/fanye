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
Public Const cszOrgErrMsgNote = "源错误描述:"
Public Const cszCustomErrSource = "自定义错误"

'对象错误公共声明
'---------------------------------------------------------
Public Const ERR_CoreErrorStart = 10000      '核心错误号起点
Public Const ERR_DBObjErrorStart = 10500     '数据对象错误号起点
Public Const ERR_BusObjErrorStart = 11000    '业务组件错误号起点
Public Const ERR_ClientToolStart = 12000    '客户端工具错误号起点
Public Const ERR_DLLStep = 50               '每个对象可分配的错误号范围
Public Const ERR_DLLIndex = 1               '组件索引号

'---------------------------------------------------------
Public Const ERR_InvalidProgramer = ERR_CoreErrorStart + 1 '未经授权的开发者!
Public Const ERR_AssertCreditCard = ERR_CoreErrorStart + 2 '用户身份无法识别!
Public Const ERR_INIFileNotValid = ERR_ClientToolStart + 3 '非法系统配置文件!

'---------------------------------------------------------
Public Const ERR_FileNotExist = ERR_ClientToolStart + ERR_DLLIndex * ERR_DLLStep + 1    '文件不存在
Public Const ERR_TemplateFileNotFound = ERR_ClientToolStart + ERR_DLLIndex * ERR_DLLStep + 2    '模板文件找不到!
Public Const ERR_FileExportError = ERR_ClientToolStart + ERR_DLLIndex * ERR_DLLStep + 3 '文件导出出错
Public Const ERR_FileSaveError = ERR_ClientToolStart + ERR_DLLIndex * ERR_DLLStep + 4 '文件保存出错
Public Const ERR_FileTypeNotSupport = ERR_ClientToolStart + ERR_DLLIndex * ERR_DLLStep + 5 '文件类型错误!
Public Const ERR_BodyDataInvalid = ERR_ClientToolStart + ERR_DLLIndex * ERR_DLLStep + 6 '报表数据参数不正确!


' *******************************************************************
' *   Brief Description: 通用错误触发程序                           *
' *   Engineer: 陆勇庆                                              *
' *   Date Generated: 2002/01/20                                    *
' *   Last Revision Date:                                           *
' *******************************************************************
Public Sub RaiseError(ByVal plErrNum As Long, Optional pszSource As String = "", Optional pszOrgMsg As String = "")
    Dim szErrMsg As String, szSource As String
'    On Error GoTo Here
On Error Resume Next
    If pszSource <> "" Then
        szSource = pszSource
    Else
        szSource = Err.Source
    End If
    
    '得到预定义的错误号
    szErrMsg = LoadResString(plErrNum)
    On Error GoTo 0
    If szErrMsg = "" Then szErrMsg = cszDefErrMsg
     
    '得到系统设置是否显示原始错误信息
    Dim bShowOrgError As Boolean
    bShowOrgError = True '值需读取注册表
    If bShowOrgError Then
        szErrMsg = szErrMsg & vbCrLf & vbCrLf & cszOrgErrMsgNote & vbCrLf & pszOrgMsg
    End If
    
    If szErrMsg = cszOrgErrMsgNote Then
        szErrMsg = pszOrgMsg
    End If
    
    Err.Raise plErrNum, szSource, szErrMsg
End Sub
