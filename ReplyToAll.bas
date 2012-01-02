Attribute VB_Name = "ReplyToAll"
Option Explicit
Private Const ReplyToAll_Id As Integer = 355
Private Const ReplyToAll_MyTag As String = "ReplyToAll_MyTag"

Public Sub ReplyToAll_Run()
If Application.ActiveExplorer.Selection.Count >= 1 Then
    Dim o As Object
    Set o = Application.ActiveExplorer.Selection.Item(1)
    If TypeName(o) = "MailItem" Then
        Dim mi As MailItem
        Set mi = o
        mi.ReplyAll.Display
        Exit Sub
    End If
End If
MsgBox "Cannot Reply to All when no mail items are selected"
End Sub

Public Sub EditMessage_Run()
Dim o As Inspector
Set o = Application.ActiveInspector()
If Not o Is Nothing Then
    o.WordEditor.UnProtect
    Exit Sub
End If
MsgBox "Cannot Edit Message unless run from the a message window"
End Sub

Public Sub DeleteImages_Run()
Dim o As Inspector
Set o = Application.ActiveInspector()
If Not o Is Nothing Then
    o.WordEditor.UnProtect
    Dim i As Object
    For Each i In o.WordEditor.InlineShapes
        i.Delete
    Next
    Exit Sub
End If
MsgBox "Cannot Delete Images unless run from the a message window"
End Sub

Public Sub ReplyToAll_AddButtons()
' For help on command buttons, see: http://support.microsoft.com/kb/201095
'ReplyToAll_UnregisterButtons
Dim bar As CommandBar
Dim added As Integer: added = 0
For Each bar In Application.ActiveExplorer.CommandBars
    ReplyToAll_AddButtonsTo added, bar.Controls
Next
If added = 0 Then
    MsgBox "Failed to find any existing Reply To All buttons to replace, please add buttons manually", vbExclamation
Else
    Dim s As String
    If added = 1 Then s = "" Else s = "s"
    MsgBox "Replaced " & added & " ReplyToAll button" & s, vbInformation
End If
End Sub

Private Sub ReplyToAll_AddButtonsTo(ByRef added As Integer, cs As CommandBarControls)
Dim i As Integer
For i = 1 To cs.Count
    Dim c As CommandBarControl
    Set c = cs(i)

    If c.Type = msoControlPopup Then
        Dim cpop As CommandBarPopup: Set cpop = c
        ReplyToAll_AddButtonsTo added, cpop.Controls
    ElseIf c.Type = msoControlButton Then
        Dim cbtn As CommandBarButton: Set cbtn = c
        If cbtn.ID = ReplyToAll_Id Then
            added = added + 1
            cbtn.Visible = False
            
            If cbtn.Parent.Controls.Count > cbtn.Index Then
                If cbtn.Parent.Controls(cbtn.Index + 1).Tag = ReplyToAll_MyTag Then
                    cbtn.Parent.Controls(cbtn.Index + 1).Delete False
                End If
            End If
    
            Dim cnew As CommandBarButton
            Set cnew = cbtn.Parent.Controls.Add(, , , cbtn.Index + 1, False)
            cnew.FaceId = cbtn.FaceId
            cnew.Caption = cbtn.Caption
            cnew.DescriptionText = cbtn.DescriptionText
            cnew.Style = cbtn.Style
            cnew.Tag = ReplyToAll_MyTag
            cnew.OnAction = "ReplyToAll_Run"
        End If
    End If
Next
End Sub
