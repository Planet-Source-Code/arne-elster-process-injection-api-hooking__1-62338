Attribute VB_Name = "modEnd"
Option Explicit

Const BeProper = 1

' __vbaEnd replacement
Public Sub EndEx()
    Dim frm As Form

    MsgBox "End called.", vbInformation

    If BeProper Then

        ' clean up
        For Each frm In Forms
            Unload frm
        Next

    Else

        ' just call __vbaEnd
        RealEnd

    End If
End Sub

' this function will point to __vbaEnd
Public Sub RealEnd()
    DoEvents
End Sub
