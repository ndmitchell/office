Option Explicit

Public Sub ReallyReplyAll()
Dim o As MailItem
Set o = Application.ActiveExplorer.Selection.Item(1)
o.ReplyAll.Display
End Sub
