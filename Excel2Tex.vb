'Option Explicit

Sub excel2tex()
    Dim start_cell, end_cell As String
    set_row = range("B2").Row
    set_col = range("B2").Column
    
    start_cell = Cells(set_row, set_col).Value
    end_cell = Cells(set_row + 1, set_col).Value
    start_row = range(start_cell).Row
    start_col = range(start_cell).Column
    end_row = range(end_cell).Row
    end_col = range(end_cell).Column
    
    out = Cells(set_row + 2, set_col).Value
    cap = Cells(set_row + 3, set_col).Text
    lab = Cells(set_row + 4, set_col).Text
'    ara = Cells(set_row + 5, set_col).Text
    ara = String(end_col - start_col + 1, "c")
    
    tex_table = "¥begin{table}[tbph]" & vbLf & "¥centering" & vbLf
    tex_table = tex_table & "¥caption{" & cap & "}" & vbLf & "¥label{tab:" & lab & "}" & vbLf
    tex_table = tex_table & "¥begin{tabular}{" & ara & "}" & vbLf & "¥hline ¥hline" & vbLf
    
    For i = start_row To end_row
        For j = start_col To end_col
            str_temp = Cells(i, j).Text
            If j = start_col Then
                tex_table = tex_table & str_temp
            Else
                If j = end_col Then
                    If i = start_row Then
                        tex_table = tex_table & " & " & str_temp & " ¥¥ ¥hline" & vbLf
                    Else
                        tex_table = tex_table & " & " & str_temp & " ¥¥ " & vbLf
                    End If
                Else
                    tex_table = tex_table & " & " & str_temp
                End If
            End If
        Next j
    Next i
    tex_table = tex_table & "¥hline ¥hline" & vbLf & "¥end{tabular}" & vbLf & "¥end{table}"
    
    range(out).Value = tex_table
     
End Sub


Sub excel2texMac()
    Dim start_cell, end_cell As String
    set_row = range("B2").Row
    set_col = range("B2").Column
    
    start_cell = Cells(set_row, set_col).Value
    end_cell = Cells(set_row + 1, set_col).Value
    start_row = range(start_cell).Row
    start_col = range(start_cell).Column
    end_row = range(end_cell).Row
    end_col = range(end_cell).Column
    
    out = Cells(set_row + 2, set_col).Value
    cap = Cells(set_row + 3, set_col).Text
    lab = Cells(set_row + 4, set_col).Text
'    ara = Cells(set_row + 5, set_col).Text
    ara = String(end_col - start_col + 1, "c")
    
    tex_table = "\begin{table}[tbph]" & vbLf & "\centering" & vbLf
    tex_table = tex_table & "\caption{" & cap & "}" & vbLf & "\label{tab:" & lab & "}" & vbLf
    tex_table = tex_table & "\begin{tabular}{" & ara & "}" & vbLf & "\hline \hline" & vbLf
    
    For i = start_row To end_row
        For j = start_col To end_col
            str_temp = Cells(i, j).Text
            If j = start_col Then
                tex_table = tex_table & str_temp
            Else
                If j = end_col Then
                    If i = start_row Then
                        tex_table = tex_table & " & " & str_temp & " \\ \hline" & vbLf
                    Else
                        tex_table = tex_table & " & " & str_temp & " \\ " & vbLf
                    End If
                Else
                    tex_table = tex_table & " & " & str_temp
                End If
            End If
        Next j
    Next i
    tex_table = tex_table & "\hline \hline" & vbLf & "\end{tabular}" & vbLf & "\end{table}"
    
    range(out).Value = tex_table
     
End Sub


