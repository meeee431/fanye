Attribute VB_Name = "mdApiVision"
Option Explicit

'华视身份证阅读器相关API

'注意：若采用查询方式自动判断卡片是否放置，则间隔时间建议大于300ms

'初始化连接
'参数：Port：连接串口（COM1~COM16）或USB口(1001~1016)
Public Declare Function CVR_InitComm Lib "termb.dll" (ByVal Port As Long) As Integer

'关闭连接
Public Declare Function CVR_CloseComm Lib "termb.dll" () As Integer

'射频操作
Public Declare Function CVR_Authenticate Lib "termb.dll" () As Integer

'读卡操作
'参数：Active
'1：读基本信息 生成文字WZ.TXT?相片数据XP.WLT和相片ZP.BMP(解码)
'2： 读基本信息 生成文字WZ.TXT和相片数据XP.WLT
'3：读最新住址信息 生成最新住址NEWADD.TXT(卡无最新地址则生成空文件)
'4：读基本信息  生成WZ.TXT(解码)，相片ZP.BMP(解码)
'5：读芯片管理号 芯片管理号IINSNDN.bin
'6：读基本信息  以设备唯一标志号，生成文字WZ.TXT(解码)，相片XP.BMP(解码)（用于终端网络环境）
Public Declare Function CVR_Read_Content Lib "termb.dll" (ByVal Active As Long) As Integer

'卡认证
Public Declare Function CVR_Ant Lib "termb.dll" (ByVal mode As Long) As Integer

'读卡操作(读入内存)
'参数：
'pucCHMsg：身份文字信息内存缓冲指针/方向：Out
'puiCHMsgLen：身份文字信息长度/默认 256 Byte
'pucPHMsg：身份照片信息内存缓冲指针/方向：Out
'puiPHMsgLen：身份照片信息长度/默认 1024 Byte
'nMode：传入参数1：文字编码为默认UCS-2格式，照片未解压成bmp文件  传入参数2：文字编码已转换成GBK国标码格式，照片未解压成bmp文件  传入参数3：文字编码为默认UCS-2格式，照片已解压成zp.bmp文件  传入参数4：文字编码已转换成GBK国标码格式，照片已解压成zp.bmp文件
Public Declare Function CVR_ReadBaseMsg Lib "termb.dll" (ByVal pucCHMsg As String, ByRef puiCHMsgLen As Integer, ByVal pucPHMsg As String, ByRef puiPHMsgLen As Integer, ByRef nMode As Integer) As Integer


'说 明：以下函数调用流程为：调用CVR_Read_Content 或者CVR_ReadBaseMsg函数，成功后再分别调用以上函数。CVR_Read_Content或者CVR_ReadBaseMsg函数自动在应用程序当前目录产生BMP照片文件。

'得到姓名信息
Public Declare Function GetPeopleName Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer

'得到性别信息
Public Declare Function GetPeopleSex Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer

'得到民族信息
Public Declare Function GetPeopleNation Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer

'得到出生日期
Public Declare Function GetPeopleBirthday Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer

'得到地址信息
Public Declare Function GetPeopleAddress Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer

'得到身份证号信息
Public Declare Function GetPeopleIDCode Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer

'得到发证机关信息
Public Declare Function GetDepartment Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer

'得到有效开始日期
Public Declare Function GetStartDate Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer

'得到有效截止日期
Public Declare Function GetEndDate Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer
