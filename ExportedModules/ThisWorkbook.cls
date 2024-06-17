VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    ' Initialize game by setting initial values, preparing the game, creating the sheet

    ' Disable all alerts to prevent Excel from showing default messages
    Application.DisplayAlerts = False
    
    ' Protect the "Game" worksheet but allow scripts to interact with it
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        With ws
            .Protect UserInterfaceOnly:=True
            .EnableSelection = xlNoSelection
            .Activate
            .Cells(22, 2).Select
        End With
    Next ws

    Call HideAllSheetsExceptCover
    Call LoadListBoxData
    
    Call BGM_Start
    
    Dim ws_rd As Worksheet
    Set ws_rd = Worksheets("Record")
    Dim historyList As MSForms.ListBox

    Set historyList = ws_rd.OLEObjects("ListBox1").Object

    With historyList
        .Font.Name = "Mont"
        .Font.Size = 15
        .Font.Bold = True
        .BackColor = RGB(0, 0, 0)
        .ForeColor = &HFDEDB0
        .Enabled = False
    End With
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Terminate the game timer and remove key bindings
    Call EndTimer
    Call RemoveKeyAssignations
    Call SaveListBoxData
    
    ' Re-enable alerts for future Excel interactions
    Application.DisplayAlerts = True
    
    ' Unprotect all worksheets and allow unrestricted selection
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        With ws
            .Unprotect
            .EnableSelection = xlNoRestrictions
        End With
    Next ws
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    ' Cancel any double-click actions on sheets to prevent unintended changes or interactions
    Cancel = True
End Sub

Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    ' Cancel any right-click actions on sheets to maintain the game's interface integrity
    Cancel = True
End Sub


