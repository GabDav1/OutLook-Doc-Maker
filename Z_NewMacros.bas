Attribute VB_Name = "Z_NewMacros"
Sub test()
MsgBox ActiveDocument.TextBox1.Value

End Sub

Sub Macro1()
'
' Recorded code for bookmarks
'
'
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="tmacro1"
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
    
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.Macro2"
'
' Recorded code for hyperlinks
'
'
    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="", _
        SubAddress:="bkm1", ScreenTip:="", TextToDisplay:="bkm1"
End Sub


Sub Macrox()
'
' Recorded code for formatting text
'
    Selection.HomeKey Unit:=wdLine
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.Font.Underline = wdUnderlineSingle
    Selection.Font.ColorIndex = wdTeal
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter

End Sub

