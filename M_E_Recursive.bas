Attribute VB_Name = "M_E_Recursive"
'This is the recursive final implementation.
'You can find a non-recursive "kernel" of this macro inside the module "Z_Mail_Extract".

Public ouT As New Outlook.Application
Public nS As Outlook.NameSpace
Public jCount As Integer
Public nCount As Integer
Public fCount As Integer
Public cntr As Integer
Dim objWord
Public doC As Word.Document
Public tdoC As Word.Document
Public tdoCR As Word.Range
Public fsTime As Boolean
Public Traversed
Public subjTxt As String
Public nomTxt As String

Sub justcall(cs As Integer)

Dim mI As Outlook.MailItem
Dim rangeS() As Range
Dim rangeTxt() As String

'The MAPI namespace is where all the mail content "lives"
Set nS = ouT.GetNamespace("MAPI")

'Left-over code from trying to add content in a different Word document
    'Set objWord = CreateObject("word.application")
    'Set tdoC = objWord.Documents.Add
    'objWord.Visible = True

'Setting up some variables
Set tdoC = ActiveDocument
Set Traversed = CreateObject("Scripting.Dictionary")
fsTime = True
jCount = 0
nCount = 0
fCount = 0
cntr = 0

'The case parameter gets passed from the calling userform.
'The cases represent what fields or combination of fields have been used in the userform.
            Select Case cs:
                    Case 1:
                        'Case: only first field(no replies to threads)
                        For Each yfolder In nS.Folders
                            If InStr(1, yfolder.Name, "Public Folders") = 0 Then Call browseInbox1(yfolder)
                        Next yfolder
                    Case 2:
                        'Case: only second field(threads with replies)
                        For Each yfolder In nS.Folders
                            If InStr(1, yfolder.Name, "Public Folders") = 0 Then Call browseInbox2(yfolder)
                        Next yfolder
                    Case 3:
                        'Case: both fields
                        For Each yfolder In nS.Folders
                            If InStr(1, yfolder.Name, "Public Folders") = 0 Then Call browseInbox3(yfolder)
                        Next yfolder
                End Select

'Collection not populated
If Traversed.Count = 0 Then
    MsgBox "No mails found"
    Exit Sub
End If

'Arrays that need the number of mails:
jCount = Traversed.Count
ReDim rangeS(1 To jCount)
ReDim rangeTxt(1 To jCount)

'In the first traversal we populate the table of contents:
For Each mail In Traversed.Items
    cntr = cntr + 1
    Set mI = mail
    Set tdoCR = tdoC.Range
    
    With tdoCR
            .InsertAfter cntr & ". " & mI.Subject & " - " & mI.ReceivedTime
             rangeTxt(cntr) = mI.Subject & " - " & mI.ReceivedTime & vbNewLine
            .InsertParagraphAfter
            .ParagraphFormat.SpaceAfter = 20
    End With
    
    'The ranges array will be used as a table of contents
    Set tdoCR = tdoC.Paragraphs(cntr).Range
    Set rangeS(cntr) = tdoCR
Next mail

'In the second traversal we'll add the actual content:
cntr = 0
For Each mail In Traversed.Items

    cntr = cntr + 1
    Set mI = mail

    If fsTime Then
         mI.GetInspector.Activate
         'fsTime = False
    End If

    Set doC = mI.GetInspector.WordEditor
    
    With doC.Range
        '.InsertBefore mI.Subject
        .Copy
    End With
    
    'Selection of whole range from target document
    Set tdoCR = tdoC.Range
    tdoC.Activate
    
    'Placement of cursor at the end of target document
    With tdoCR
        .Collapse direction:=wdCollapseEnd
        .Select
    End With
    
    'Placement of separator(between table of contents and contents) symbol and first time traversal flag:
    If fsTime Then
            With tdoCR
                .InsertParagraphAfter
                .InsertAfter "=================================================================================" & vbNewLine
                .Collapse direction:=wdCollapseEnd
                .Select
                Selection.HomeKey Unit:=wdLine
                Selection.EndKey Unit:=wdLine, Extend:=wdExtend
                Selection.Font.Underline = wdUnderlineSingle
                Selection.Font.ColorIndex = wdTeal
                Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            End With
        fsTime = False
    End If
    
    'Bookmarks the spot where the cursor is at:
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="bkm" & cntr
        .ShowHidden = False
    End With
    
    'Pastes the titles and mail contents inside document:
    With tdoCR
        .InsertParagraphAfter
        .InsertAfter rangeTxt(cntr)
        .Collapse direction:=wdCollapseEnd
        .Select
        .PasteSpecial
    End With
          
Next mail

'Adding hyperlinks, inverse traversal is used (this corrected an earlier bug):
For k = jCount To 1 Step -1
 ActiveDocument.Hyperlinks.Add Anchor:=rangeS(k), SubAddress:="bkm" & k, _
        ScreenTip:="bkm" & k, TextToDisplay:=rangeS(k).Text
Next k

UserForm1.Hide
MsgBox "Done !"

End Sub

Sub browseInbox1(ByVal thisFolder As Outlook.Folder)

Dim czeC As Boolean

        For Each witem In thisFolder.Items
            If witem.Class = olMail Then
                jCount = jCount + 1
                 'No replies and no auto replies
                 If InStr(1, UCase(witem.Subject), subjTxt) <> 0 And InStr(1, witem.Subject, "RE: ") = 0 _
                 And InStr(1, witem.Subject, "Automatic reply: ") = 0 Then

                     'Identifies each mail by its subject and time of receival
                     If Not Traversed.Exists(witem.Subject + Format(witem.ReceivedTime, "yyyy/mm/dd hh")) Then Traversed.Add witem.Subject + Format(witem.ReceivedTime, "yyyy/mm/dd hh"), witem
                     
                 End If
        
            Else: nCount = nCount + 1
            End If
    Next witem

If (thisFolder.Folders.Count > 0) Then
    For Each zfolder In thisFolder.Folders
        browseInbox1 zfolder
    Next zfolder
End If

End Sub

Sub browseInbox2(ByVal thisFolder As Outlook.Folder)

Dim czeC As Boolean

        For Each witem In thisFolder.Items
            If witem.Class = olMail Then
                jCount = jCount + 1
                 'No auto replies
                 If InStr(1, UCase(witem.Subject), nomTxt) <> 0 And InStr(1, witem.Subject, "Automatic reply: ") = 0 Then

                     'Identifies each mail by its subject and time of receival
                     If Not Traversed.Exists(witem.Subject + Format(witem.ReceivedTime, "yyyy/mm/dd hh")) Then Traversed.Add witem.Subject + Format(witem.ReceivedTime, "yyyy/mm/dd hh"), witem
                     
                 End If
        
            Else: nCount = nCount + 1
            End If
    Next witem

If (thisFolder.Folders.Count > 0) Then
    For Each zfolder In thisFolder.Folders
        browseInbox2 zfolder
    Next zfolder
End If

End Sub

Sub browseInbox3(ByVal thisFolder As Outlook.Folder)

Dim czeC As Boolean

        For Each witem In thisFolder.Items
            If witem.Class = olMail Then
                jCount = jCount + 1
                 'No auto replies but works with both fields
                 If ((InStr(1, UCase(witem.Subject), subjTxt) And InStr(1, witem.Subject, "RE: ") = 0) <> 0 Or InStr(1, UCase(witem.Subject), nomTxt) <> 0) _
                 And InStr(1, witem.Subject, "Automatic reply: ") = 0 Then

                     'Identifies each mail by its subject and time of receival
                     If Not Traversed.Exists(witem.Subject + Format(witem.ReceivedTime, "yyyy/mm/dd hh")) Then Traversed.Add witem.Subject + Format(witem.ReceivedTime, "yyyy/mm/dd hh"), witem
                     
                 End If
        
            Else: nCount = nCount + 1
            End If
    Next witem

If (thisFolder.Folders.Count > 0) Then
    For Each zfolder In thisFolder.Folders
        browseInbox3 zfolder
    Next zfolder
End If

End Sub

