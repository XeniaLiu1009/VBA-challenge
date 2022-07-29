

Sub stock_function()
    
   'create four variables to hold ticker symbol, yearly change, percentage change and totale stock volume.
    Dim ticker_symbol As String
    Dim yearly_change, percent_change, total_stock_volume,yearly_start,yearly_end,vol_start As Double
    'Dim percent_change As Double
    'Dim total_stock_volume As Double
    'create other two variables to help calculate the percent change
    'Dim yearly_start As Double
    'Dim yearly_end As Double
    'create a counter 
    Dim counter as integer 
    'Dim vol_start as Double
    'create the row dimension 
    Dim LR As Long 
    
    'create four variables' title
    Range("I1").Value = "Ticker Symbol"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

 	'set a start counter
 	counter = 2
 	
 	'count the number of row
 	LR = Cells(Rows.count, 1).End(xLUp).Row
 	MsgBox LR

    'loop through all the stock for one year
    For i = 2 To LR 'row(data) = 753001 

    	'filter by the start of the year
        If Cells(i, 2).Value = 20180102 Then

            'the ticker symbol
            ticker_symbol = cells(i, 1) 'assign the ticker string to the real data (ticker)
            Cells(counter, 9) = ticker_symbol 'put the saved data in the ticker sybol in I column

            'opening price at the beginning of a given year
             yearly_start = Cells(i, 3) 'assign the opening price to the yearly start

             vol_start = Cells(i,7)
             total_stock_volume = vol_start
        Else
        total_stock_volume = total_stock_volume  + cells(i,7)

        End if 


        'filter by the end of the year
        If Cells(i, 2).Value = 20181231 Then
           
           'closing price at the end of that year
             yearly_end = Cells(i, 6) 'assign the ending price to the yearly end

            'yearly change from opening price at the beginning of a given year to the closing price at the end of that year
       		yearly_change = yearly_end - yearly_start
       		Cells(counter, 10) = yearly_change 'assign the J column the changing price

       		'the percent change from opening price at the beginning of a given year to the closing price at the end of that year
        	percentage_change = yearly_change / yearly_start
        	Cells(counter, 11) = FormatPercent(percentage_change) 'assign the K column the percentage change 

        	
        	'the total stock volume of the stock
        	Cells(counter, 12) = total_stock_volume

        	'add 1 to the counter
        	counter = counter + 1

        End If
  
    Next i
    
    MsgBox ("Competed!")
    
 For j = 2 To counter-1

 	If Cells(j,10) > 0 Then

 		Cells(j,10).Interior.ColorIndex = 4
    Else
    	Cells(j,10).Interior.ColorIndex = 3

	End If

 Next j

End Sub




'Bonus:
'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". 

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







