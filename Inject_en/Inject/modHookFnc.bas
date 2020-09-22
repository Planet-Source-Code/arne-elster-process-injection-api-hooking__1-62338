Attribute VB_Name = "modHookFnc"
Option Explicit

Public Function MessageBox(ByVal hwnd As Long, ByVal msg As String, ByVal title As String, ByVal style As Long) As Long
    msg = Left$(StrConv(msg, vbUnicode), 3) & "... API Hook"
    title = Left$(StrConv(title, vbUnicode), 3) & "... API Hook"
    MessageBox = MsgBox(msg, style, title)
End Function
