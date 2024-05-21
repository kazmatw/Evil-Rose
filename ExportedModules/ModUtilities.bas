Attribute VB_Name = "ModUtilities"
Function ChangeBrightness(Col As Long, Per As Integer) As Long
    ' Declare variables for the RGB components of a color
    Dim R, G, B As Single
    
    ' Extract the Red, Green, and Blue components from the composite long integer color value
    R = Col Mod 256  ' Get the red component
    G = Col \ 256 Mod 256  ' Get the green component
    B = Col \ 65536 Mod 256  ' Get the blue component
    
    ' If brightness percentage is positive and any color component is at its maximum, adjust zero components slightly to prevent visual issues
    If Per > 0 Then
        If R = 255 Or G = 255 Or B = 255 Then
            If R = 0 Then R = 32  ' Increase red if initially zero
            If G = 0 Then G = 32  ' Increase green if initially zero
            If B = 0 Then B = 32  ' Increase blue if initially zero
        End If
    End If
    
    ' Adjust the RGB components based on the percentage (Per)
    R = R + Per * (R / 100)  ' Calculate new red value
    G = G + Per * (G / 100)  ' Calculate new green value
    B = B + Per * (B / 100)  ' Calculate new blue value
    
    ' Ensure RGB values remain within the 0-255 range
    If R < 0 Then R = 0
    If G < 0 Then G = 0
    If B < 0 Then B = 0
    If R > 255 Then R = 255
    If G > 255 Then G = 255
    If B > 255 Then B = 255
    
    ' Combine the adjusted RGB values into a single long integer and return it
    ChangeBrightness = RGB(Int(R), Int(G), Int(B))
End Function

Function CopyBloLibArrToCurBloArr(Blo As Byte)
    ' Loop through each element in the block's array (assuming 6x6 grid)
    For i = 1 To 6
        For j = 1 To 6
            ' Copy each block definition from the block library to the current block's array
            CurBlo.Arr(i, j) = BloLib(BloSet(CurBloSet).Blo(Blo)).Arr(i, j)
        Next j
    Next i
End Function
