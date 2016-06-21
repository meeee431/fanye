Attribute VB_Name = "mdlMain"
Option Explicit
Public Const cszPackage = ""


'ÉèÖÃ´íÎóÀ´Ô´
Function SetErrSource(modName As String, procName As String) As String
    SetErrSource = modName & ":" & procName
'    SetErrSource = err.Source & _
                   IIf(Left$(modName, (InStr(1, modName, ".") - 1)) = err.Source, _
                   Mid$(modName, InStr(1, modName, ".")), "->" & modName) & _
                   "." & procName & "@" & GetComputerName

End Function

