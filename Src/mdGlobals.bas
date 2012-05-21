Attribute VB_Name = "mdGlobals"
Option Explicit

Public Const LIB_NAME           As String = "VszLib"

#If VszLogging Then
    Public Sub DebugOutput(sText As String, sSource As String)
        Call APIOutputDebugString(GetCurrentThreadId() & ": " & sSource & " " & sText & vbCrLf)
    End Sub
#End If

Public Property Get InIde() As Boolean
    Debug.Assert pvSetTrue(InIde)
End Property

Private Function pvSetTrue(bValue As Boolean) As Boolean
    bValue = True
    pvSetTrue = True
End Function

Public Function FileExists(sFile As String) As Boolean
    On Error Resume Next
    If GetAttr(sFile) = -1 Then
    Else
        FileExists = True
    End If
    On Error GoTo 0
End Function
