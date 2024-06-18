Attribute VB_Name = "ModTimerEvents"
Sub StartTimer()
    ' Set the timer interval to 50 milliseconds
    MilSec = 50
    ' Create a timer that calls the TimerProcedure subroutine every 50 milliseconds
    ' 0& indicates the use of the default timer and window
    ' AddressOf TimerProcedure gets the address of the TimerProcedure subroutine
    TimID = SetTimer(0&, 0&, MilSec, AddressOf TimerProcedure)
End Sub

Sub EndTimer()
    On Error Resume Next
    KillTimer 0&, TimID  ' Kill the timer with the ID TimID
End Sub

Sub TimerProcedure()
    ' Check if the current pass equals the execution threshold
    ' Tim.CurPas will be init (= 0) in GenerateBlocks(4)
    ' Tim.ExeThr = 16 at first
    If Tim.CurPas = Tim.ExeThr Then
        Tim.CurPas = 0  ' Reset the current pass
        ' Do gameLogic based on Game status
        Select Case GamSta
            Case 1  ' Game is running
                Call MoveBlockDown(0)  ' Move the block down
            Case 2  ' Rows have been deleted
                Call DrawDeletedRows  ' Draw the deleted rows
            Case 3  ' Rows need to be dropped
                Call DropRows
                Call GenerateBlocks(1)  ' Generate one new block
                Tim.ExeThr = Tim.LevTim  ' Set the execution threshold to the level timer
                GamSta = 1  ' Set game state back to running
        End Select
    Else
        Tim.CurPas = Tim.CurPas + 1  ' Increment the current pass
    End If
End Sub


