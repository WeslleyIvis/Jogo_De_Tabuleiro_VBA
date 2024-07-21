Sub Cria_Tabuleiro()

    Sheets("Corrida_Maluca").Activate
    Range("B:Z").ColumnWidth = 10
    Range("B2:AF17").RowHeight = 50
    
    
    Dim Tabuleiro As Range
    Dim cor_casas_tabuleiro As Double
    Dim cell As Range
    Dim i As Integer, j As Integer
    Dim cellAdjacente As Range
    Dim count_cases As Integer
    
    
    Set Tabuleiro = Range("B2:AF17")
    count_cases = 1

    cor_casas_tabuleiro = 6569237

    For i = 1 To Tabuleiro.Rows.Count
        For j = 1 To Tabuleiro.Columns.Count
        
            Set cell = Tabuleiro.Cells(i, j)
            
            If cell.Interior.Color = cor_casas_tabuleiro Then
    
                If cell.Offset(-1, 0).Value Or cell.Offset(0, -1).Value Or cell.Offset(0, 1).Value Or cell.Offset(1, 0).Value Then
                                              
                    If cell.Value = "" Then
                        count_cases = count_cases + 1
                        cell.Value = count_cases
                    End If
                    
                    If cell.Offset(0, -1).Interior.Color = cor_casas_tabuleiro And cell.Offset(0, -1).Value = "" And cell.Value Then
                        j = j - 2
                    End If
                        
                    If cell.Offset(-1, 0).Interior.Color = cor_casas_tabuleiro And IsEmpty(cell.Offset(-1, 0).Value) And cell.Value <> "" Then
                       i = i - 1
                       j = j - 2
                    End If
                    
                    ' Otimiza o processo do algoritmo!
                    'If IsEmpty(cell.Offset(0, 1).Value) Then
                        'i = i + 1
                        
                        'If cell.Offset(0, -1).Interior.Color = cor_casas_tabuleiro
                        'j = j - j + 1
                    'End If
                    
                End If
            
            End If
        Next j
    Next i


    
    'casas_tabuleiro = Array("C3", "D3", "E3")
    
  
    
    
    'For i = LBound(casas_tabuleiro) To UBound(casas_tabuleiro)
        'Debug.Print Range(casas_tabuleiro(i)).Value
        
        'Set cell = Range(casas_tabuleiro(i))
        'Item = cell.Value
        
   ' Next
    
    
End Sub