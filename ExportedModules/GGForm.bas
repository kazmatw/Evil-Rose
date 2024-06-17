VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GGForm 
   Caption         =   "UserForm1"
   ClientHeight    =   9432.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   16980
   OleObjectBlob   =   "GGForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GGForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TextBox1_Change()
    userName = TextBox1.Text
End Sub

