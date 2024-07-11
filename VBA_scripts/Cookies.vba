Sub Cookies()

Dim rn As Integer
Dim S()
Dim C()

'Count Number of rows containing cooking sells
rn = WorksheetFunction.CountA(Columns("A:A")) - 2

'Finding the distinct sellers (S) and cookies(C)
S = WorksheetFunction.Unique(Range("A6:A" & rn + 5))
C = WorksheetFunction.Unique(Range("C6:C" & rn + 3))

'Assign values to Range(I11) and Range(I13)
Range("H11").Value = "TOTAL SALES"
Range("I13").Value = "Total Boxes"

'Bold the value in Range("I13")
Range("I13").Font.Bold = True

'Assign Monthly Sales to Range(I5)
Range("I5").Value = "Monthly Sales"

'Format Range("I5")
With Range("I5")
    .HorizontalAlignment = xlCenter
    .Font.Bold = True
    .WrapText = True
End With

'Populate Range("H6:H10") with the unique sellers
Range("H6:H10") = S

'Populate Range("H14:H16") with the unique cookies
Range("H14:H16") = C

'Create a border around Range("H5:I11")
With Range("H5:I11").Borders(xlEdgeBottom)
         .Weight = xlThin
End With
With Range("H5:I11").Borders(xlEdgeLeft)
         .Weight = xlThin
End With
With Range("H5:I11").Borders(xlEdgeRight)
         .Weight = xlThin
End With
With Range("H5:I11").Borders(xlEdgeTop)
         .Weight = xlThin
End With

'Create a border around Range("H13:I16")
With Range("H13:I16").Borders(xlEdgeBottom)
         .Weight = xlThin
End With
With Range("H13:I16").Borders(xlEdgeLeft)
         .Weight = xlThin
End With
With Range("H13:I16").Borders(xlEdgeRight)
         .Weight = xlThin
End With
With Range("H13:I16").Borders(xlEdgeTop)
         .Weight = xlThin
End With

'Calculate Monthly sales for each sellers using R1C1 formula style
Range("H6").Offset(0, 1).Resize(5, 1).FormulaR1C1 = "=SUMIF(R6C1:R" & rn + 5 & "C1,RC[-1],R6C6:R" & rn + 5 & "C6)"
Range("H6").Offset(5, 1).FormulaR1C1 = "=SUM(R[-5]C:R[-1]C)"
Range("H6").Offset(8, 1).Resize(3, 1).FormulaR1C1 = "=SUMIFS(R6C4:R" & rn + 5 & "C4,R6C3:R" & rn + 5 & "C3,RC[-1])"
End Sub
