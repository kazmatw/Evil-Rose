Attribute VB_Name = "ModTimerEvents2p"
Sub StartTimer_2p()
    ' Set the timer interval to 50 milliseconds
    MilSec_2p = 50
    ' Create a timer that calls the TimerProcedure subroutine every 50 milliseconds
    ' 0& indicates the use of the default timer and window
    ' AddressOf TimerProcedure gets the address of the TimerProcedure subroutine
    TimID_2p = SetTimer(0&, 0&, MilSec_2p, AddressOf TimerProcedure_2p)
End Sub

Sub EndTimer_2p()
    On Error Resume Next  ' Continue if there's an error
    KillTimer_2p 0&, TimID_2p  ' Kill the timer with the ID TimID
End Sub

Sub TimerProcedure_2p(ByVal HWnd_2p As LongPtr, ByVal uMsg_2p As LongPtr, ByVal nIDEvent_2p As LongPtr, ByVal dwTimer_2p As LongPtr)
    ' Check if the current pass equals the execution threshold
    If Tim_2p.CurPas = Tim_2p.ExeThr Then
        Tim_2p.CurPas = 0  ' Reset the current pass counter
        ' Execute actions based on the game state
        Select Case GamSta_2p
            Case 1  ' Game is running
                Call MoveBlockDown_2p(0)  ' Move the block down
            Case 2  ' Rows have been deleted
                Call DrawDeletedRows_2p  ' Draw the deleted rows
            Case 3  ' Rows need to be dropped
                Call DropRows_2p  ' Drop the rows
                Call GenerateBlocks_2p(1)  ' Generate a new block
                Tim_2p.ExeThr = Tim_2p.LevTim  ' Set the execution threshold to the level timer
                GamSta_2p = 1  ' Set game state back to running
        End Select
    Else
        Tim_2p.CurPas = Tim_2p.CurPas + 1  ' Increment the current pass counter
    End If
End Sub



