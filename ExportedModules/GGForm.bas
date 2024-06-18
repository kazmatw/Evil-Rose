VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GGForm 
   Caption         =   "UserForm1"
   ClientHeight    =   9435.001
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   16980
   OleObjectBlob   =   "GGForm.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "GGForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim regEx As New RegExp 'using regular expression

Private Sub TextBox1_Change()
    userName = TextBox1.Text
End Sub
Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        regEx.Pattern = "[\s+]" 'set regEx pattern
        If regEx.Test(userName) Or userName = "" Then
            response = MsgBox("YOU HAVEN'T ENTER YOUR NAME!!!", vbOKOnly, "GGFORM WARNING")
        Else
            Unload GGForm 'Close GGForm
        End If
        KeyCode = 0
    End If
End Sub

Private Sub UserForm_Initialize()
    userName = ""
    TextBox1.Text = ""
End Sub
