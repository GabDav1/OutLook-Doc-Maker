Attribute VB_Name = "Z_Mail_Extract"
Sub browseInbox()

Dim ouT As New Outlook.Application
Dim nS As Outlook.NameSpace
Dim foL As Outlook.Folder
Dim mI As Outlook.MailItem
Dim jCount As Integer
Dim nCount As Integer
Dim fCount As Integer

Set nS = ouT.GetNamespace("MAPI")
Set foL = nS.GetDefaultFolder(olFolderInbox)

jCount = 0
nCount = 0
fCount = 0

'This is the initial non-recursive attempt("the kernel").
'Check the module "M_E_Recursive" for actual implementation.

For Each why In nS.Folders
    For Each whyf In why.Folders
    
        For Each witem In whyf.Items
            If witem.Class = olFolder Then fCount = fCount + 1
            If witem.Class = olMail Then
                Set mI = witem
                jCount = jCount + 1
        
                 If InStr(1, mI.Subject, "Automatizare - Macro ZEM GAP") <> 0 Then
                     ActiveDocument.Content = ActiveDocument.Content & mI.Subject & " trimis de " & mI.SenderName & vbNewLine & mI.ReceivedTime
                 'Else: Exit Sub
                 End If
        
            Else: nCount = nCount + 1
            End If
    Next witem
    'Debug.Print why.Subject
Next whyf
Next why

MsgBox "actual mails " & jCount & " and non-mails " & nCount & vbNewLine & fCount & " extra folders"

End Sub
