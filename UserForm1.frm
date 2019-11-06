VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "MailExtract"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Dim caz As Integer

subjTxt = UCase(UserForm1.TextBox1.Value)
nomTxt = UCase(UserForm1.TextBox2.Value)

If subjTxt <> "" And nomTxt <> "" Then
    caz = 3
ElseIf subjTxt <> "" And nomTxt = "" Then
    caz = 1
ElseIf subjTxt = "" And nomTxt <> "" Then
    caz = 2
Else: caz = 4
End If

If caz = 4 Then
    MsgBox "You should type something, mate"
    'UserForm1.Hide
Else:
    Call justcall(caz)
    'MsgBox ActiveDocument.TextBox1.Value
End If

End Sub
