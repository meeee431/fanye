Attribute VB_Name = "mdlMessage"
Option Explicit

'Const SM_ID_LEN = 23
'Const SM_ORG_NO_LEN = 21
'Const SM_DST_NO_LEN = 21
'Const SM_FEE_NO_LEN = 21
'Const SM_SEND_TIME_LEN = 19
'Const SM_CONTEXT_LEN = 140
'Const BUSINESS_ID_LEN = 10
'Const SM_MAX_DST_NO_LEN = (SM_DST_NO_LEN + 1) * 50
'Const FEE_TYPE_LEN = 10
'Const FEE_SUM_LEN = 6
'Const LINK_ID_LEN = 20
'Const VER_LEN = 3
'Const CMD_WORD_LEN = 5 '命令字的长度(含结尾符'\0')
'
Public Const IMAPI_SUCC = 0               ' 操作成功
Public Const IMAPI_CONN_ERR = -1          ' 连接数据库出错
Public Const IMAPI_CONN_CLOSE_ERR = -2    ' 数据库关闭失败
Public Const IMAPI_INS_ERR = -3           ' 数据库插入错误
Public Const IMAPI_DEL_ERR = -4           ' 数据库删除错误
Public Const IMAPI_QUERY_ERR = -5         ' 数据库查询错误
Public Const IMAPI_DATA_ERR = -6          ' 数据错误
Public Const IMAPI_API_ERR = -7           ' API编码不存在
Public Const IMAPI_DATA_TOOLONG = -8      ' 内容太长
Public Const IMAPI_INIT_ERR = -9          ' 没有初始化或初始化失败

Public Const SM_ID_LEN = 8                ' 短信ID的最大长度(0-99999999)
Public Const SM_MOBILE_LEN = 16           ' 手机号码最大长度
Public Const SM_CONTEXT_LEN = 260         ' 短信内容最大长度
Public Const SM_RPT_LEN = 100             ' 短信回执描述的最大长度

Public Type MOItem
    smMobile(SM_MOBILE_LEN - 1) As Byte
    smContent(SM_CONTEXT_LEN - 1) As Byte
    smID As Long
End Type

Public Type RptItem
    rptMobile(SM_MOBILE_LEN - 1) As Byte
    smID As Long
    rptId As Long
    rptDesc(SM_RPT_LEN - 1) As Byte
End Type


Public Declare Function init Lib "ImApi.dll" (ByVal ip As String, ByVal userName As String, ByVal password As String, ByVal apiCode As String) As Long
'释放
Public Declare Function release Lib "ImApi.dll" () As Long
'发送消息
Public Declare Function sendSM Lib "ImApi.dll" (ByVal mobile As String, ByVal content As String, ByVal smID As Long) As Long
'发送Wap Push消息
Public Declare Function sendWapPushSM Lib "ImApi.dll" (ByVal mobile As String, ByVal content As String, ByVal smID As Long, ByVal mobile As String) As Long
'接收短信，返回查询到的短信数，并删除这些短信
Public Declare Function receiveSM Lib "ImApi.dll" (ByRef MOItems As MOItem, ByVal retsize As Long) As Long
'接收回执，返回查询到的回执数，并删除这些回执
Public Declare Function receiveRPT Lib "ImApi.dll" (ByRef RptItems As RptItem, ByVal retsize As Long) As Long
                         
