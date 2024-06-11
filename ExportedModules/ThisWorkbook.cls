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
    Call SetInitialValues
    Call InitializeGame
    
    ' Disable all alerts to prevent Excel from showing default messages
    Application.DisplayAlerts = False
    
    ' Protect the "Game" worksheet but allow scripts to interact with it
    'With Worksheets("Game")
    '    .Protect userinterfaceonly:=True
    '    .EnableSelection = xlNoSelection  ' Prevent any cell selection on this sheet
    'End With
    
    Call CreateGameSheet
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
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    ' Cancel any double-click actions on sheets to prevent unintended changes or interactions
    'Cancel = True
End Sub

Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    ' Cancel any right-click actions on sheets to maintain the game's interface integrity
    'Cancel = True
End Sub

