
'Bonus 

Sub Greatest_Function():

    Dim maximum, minimum, max_total As Double
    
    'create five variables' title
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volumn"
    
    Set my_range = Worksheets("2018").Range("K1:K3001")
    Set my_range2 = Worksheets("2018").Range("L1:L3001")
    
    maximum = Application.WorksheetFunction.Max(my_range)
    Range("Q2").Value = FormatPercent(maximum)

    minimum = Application.WorksheetFunction.Min(my_range)
    Range("Q3").Value = FormatPercent(minimum)
    
    max_total = Application.WorksheetFunction.Max(my_range2)
    Range("Q4").Value = max_total

    for i = 1 To 3001

    	if Cells(i,11) = maximum then
    	Range("P2").value = Cells(i,9) 
        end if 

    	if Cells(i,11) = minimum then
    	Range("P3").value = Cells(i,9) 
        end if 
        
    	if Cells(i,12) = max_total then
    	Range("P4").value = Cells(i,9) 
   		end if

    Next i 
  
End Sub
