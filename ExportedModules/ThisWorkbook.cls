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
    With Worksheets("Game")
        '.Protect userinterfaceonly:=True
        .EnableSelection = xlNoSelection  ' Prevent any cell selection on this sheet
    End With
    
    With Worksheets("2p")
        '.Protect userinterfaceonly:=True
        .EnableSelection = xlNoSelection  ' Prevent any cell selection on this sheet
    End With
    
    With Worksheets("Game Records")
        .Protect userinterfaceonly:=True
        .EnableSelection = xlNoSelection  ' Prevent any cell selection on this sheet
    End With
    
    'Call CreateGameSheet
    
    Dim ws_rd As Worksheet
    Set ws_rd = Worksheets("Game Records")
    Dim historyList As MSForms.ListBox

    Set historyList = ws_rd.OLEObjects("ListBox1").Object

    With historyList
        .Font.Name = "Arial"
        .Font.Size = 10
        .BackColor = RGB(0, 0, 0)
        .ForeColor = RGB(255, 255, 255)
    End With
    
    Dim currentSheet As Worksheet
    Set currentSheet = ThisWorkbook.ActiveSheet
    
    ws_rd.Activate
    ActiveWindow.DisplayGridlines = False
    currentSheet.Activate
    ActiveWindow.DisplayGridlines = False
    
    
    Call SelectMode
    
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Terminate the game timer and remove key bindings
    Call EndTimer
    Call RemoveKeyAssignations
    
    ' Re-enable alerts for future Excel interactions
    Application.DisplayAlerts = True
    
    ' Unprotect the "Game" worksheet and allow unrestricted selection
    With Worksheets("Game")
        .Unprotect
        .EnableSelection = xlNoRestrictions
    End With
    
    With Worksheets("2p")
        .Unprotect
        .EnableSelection = xlNoRestrictions
    End With
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    ' Cancel any double-click actions on sheets to prevent unintended changes or interactions
    Cancel = True
End Sub

Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    ' Cancel any right-click actions on sheets to maintain the game's interface integrity
    Cancel = True
End Sub


