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

Public Sub ReplyToAll_AddButtons()
' For help on command buttons, see: http://support.microsoft.com/kb/201095
'ReplyToAll_UnregisterButtons
Dim bar As CommandBar
For Each bar In Application.ActiveExplorer.CommandBars
    ReplyToAll_AddButtonsTo bar.Controls
Next
End Sub

Private Sub ReplyToAll_AddButtonsTo(cs As CommandBarControls)
Dim i As Integer
For i = 1 To cs.Count
    Dim c As CommandBarControl
    Set c = cs(i)

    If c.Type = msoControlPopup Then
        Dim cpop As CommandBarPopup: Set cpop = c
        ReplyToAll_AddButtonsTo cpop.Controls
    ElseIf c.Type = msoControlButton Then
        Dim cbtn As CommandBarButton: Set cbtn = c
        If cbtn.ID = ReplyToAll_Id Then
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
