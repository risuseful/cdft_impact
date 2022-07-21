Private Sub CommandButton1_Click()

Dim wcol As Integer, mat_row As Integer
Dim cdft_row As Integer
Dim t_i As Date, t_f As Date, t_cdft As Date
Dim wname As String


'clear values in cells
Range("H4:EQ27").ClearContents


For wcol = 8 To 147

    wname = Cells(3, wcol)
    
    For mat_row = 4 To 27
    
        t_i = Cells(mat_row, 7)
        t_f = Cells(mat_row + 1, 7)
        
        For cdft_row = 4 To 161
            t_cdft = Cells(cdft_row, 3)
            w_cdft = Cells(cdft_row, 4)
            If t_cdft > t_i And t_cdft <= t_f Then
                If w_cdft = wname Then
                    Cells(mat_row + 1, wcol) = t_cdft
                    Exit For
                End If
            End If
            
        Next cdft_row
        
        
        
    Next mat_row
Next wcol



End Sub
