Attribute VB_Name = "modShellAndWaitMsgBox"
' Copyright (c) 2015-2022 Jeffrey J. Long. All rights reserved

'@Folder("Open Source")

Option Explicit

Public Sub ShellAndWaitMessage(ByVal ret As Long)

    Select Case ret
    '@Ignore EmptyCaseBlock
    Case ShellAndWaitResult.success
        ' No action

    Case ShellAndWaitResult.Failure
        MsgBox GetMessage("msgboxShellAndWaitFailure"), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)

    Case ShellAndWaitResult.timeout
        MsgBox GetMessage("msgboxShellAndWaitTimeout"), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)

    Case ShellAndWaitResult.InvalidParameter
        MsgBox GetMessage("msgboxShellAndWaitInvalidParameter"), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)

    Case ShellAndWaitResult.SysWaitAbandoned
        MsgBox GetMessage("msgboxShellAndWaitSysWait"), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)

    Case ShellAndWaitResult.UserWaitAbandoned
        MsgBox GetMessage("msgboxShellAndWaitUserWait"), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)

    Case ShellAndWaitResult.UserBreak
        MsgBox GetMessage("msgboxShellAndWaitUserBreak"), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)

    Case Else
        MsgBox GetMessage("msgboxShellAndWaitUnknownError") & ret, vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)

    End Select

End Sub

