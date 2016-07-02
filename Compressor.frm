VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Compressor 
   Caption         =   "Email Compressor"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6090
   OleObjectBlob   =   "Compressor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Compressor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fold As Folder
Dim CountPre As Integer
Dim CountPost As Integer
Dim SizePre As Long
Dim SizePost As Long
Dim Pile As Collection


Private Sub btnPick_Click()
Set fold = Outlook.Session.Accounts.Session.PickFolder()
If fold Is Nothing Then
    lblFolder = ""
    btnRun.Enabled = False
Else
    lblFolder = fold.FolderPath
    btnRun.Enabled = True
End If
End Sub

Private Sub ApplyMail(m As MailItem)
CountPre = CountPre + 1
CountPost = CountPost + 1
SizePre = SizePre + m.Size
If chkPlainText.Value And m.BodyFormat <> olFormatPlain Then
    m.BodyFormat = olFormatPlain
    m.Save
End If
SizePost = SizePost + m.Size
If chkDeleteOverlaps.Value Then
    Dim old As MailItem: Set old = Lookup(m.Subject)
    If old Is Nothing Then
        Pile.Add m, m.Subject
    Else
        If old.Size > m.Size Then
            If InStr(1, old.Body, m.Body) > 0 Then
                SizePost = SizePost - m.Size
                m.Delete
                CountPost = CountPost - 1
            End If
        Else
            ' don't delete an item if it occurs twice
            If old.Size < m.Size Then
                If InStr(1, m.Body, old.Body) > 0 Then
                    SizePost = SizePost - old.Size
                    old.Delete
                    CountPost = CountPost - 1
                End If
                Pile.Remove m.Subject
                Pile.Add m, m.Subject
            End If
        End If
    End If
End If
End Sub

Private Function Lookup(s As String) As MailItem
On Error GoTo err
Set Lookup = Pile.Item(s)
err:
End Function

Private Sub ApplyFolder(f As Folder)
Dim d As Date
d = DateAdd("d", -spnDays.Value, Date)

Dim o As Object
Dim m As MailItem
For Each o In f.Items
    If o.Class = olMail Then
        Set m = o
        If m.CreationTime < d Then
            ApplyMail m
            SetMessage True
        End If
    End If
Next

If chkRecursive.Value Then
    Dim y As Folder
    For Each y In f.Folders
        ApplyFolder y
    Next
End If
End Sub

Private Sub btnRun_Click()
CountPre = 0
CountPost = 0
SizePre = 0
SizePost = 0
Set Pile = New Collection
SetMessage True
ApplyFolder fold
SetMessage False
End Sub

Private Sub SetMessage(Running As Boolean)
lblStatus.Caption = IIf(Running, "Running, ", "Finished, ") & " items, from " & CountPre & " @ " & Mb(SizePre) & " to " & CountPost & " @ " & Mb(SizePost) & IIf(Running, "...", "")
DoEvents
End Sub

Private Function Mb(x As Long) As String
Mb = Round(x / 1024 / 1024, 2)
Dim i As Integer: i = InStr(1, Mb, ".")
If i = 0 Then
    Mb = Mb & ".00"
ElseIf i + 1 = Len(Mb) Then
    Mb = Mb & "0"
End If
Mb = Mb & "Mb"
End Function

Private Sub spnDays_Change()
lblDays.Caption = "Skip the last " & spnDays.Value & " days"
End Sub

Private Sub UserForm_Initialize()
spnDays_Change
End Sub
