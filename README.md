# AUTOMATING MONTHLY SALES REPORTING PROCESSES
This project seeks to provide a comprehensive approach to automating sales report using VBA programming. 

## PROMBLEM STATEMENT
A sales sheet for cookie orders for the month of Febuary has been provided for as shown below

![Screenshot 2024-07-10 180727](https://github.com/dannieRope/Automating-Monthly-sales-Reporting-processes-VBA-Project/assets/132214828/640d9337-3e7a-4bd9-bb27-5a4f650dbe73)

Write a VBA script to calculate Monthly Sales automatically for each of the 5 Sellers  in cells I6:I10.  Furthermore, calculate the TOTAL sales for all sellers in cell I11. 

Your approach should adapt to the size of the data in this sales sheet.  In other words, there are 11 rows shown in the data, yet your sub(VBA script) should work the same if we had 20 rows, for example, or 50 rows.  

Next, you need to have VBA automatically calculate the Total Boxes of each of the 3 different cookies in cells I14:I16.

## SOLUTION
To solve the above problem, two subroutines were created.
1. cookies: This contains codes that automate the entire processes
2. Reset: This contains code that clears all contents in cell I6:I11 and cell I14:I16

Also, at the end of the project, two buttons are created and assigned to the subroutines, and runs the subroutines when clicked. 

### Cookies() subroutine 
Below is the code written for this subroutine 

```vba

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
```
### Reset() subroutine 
Below is the code script for this subroutine 

```vba
Sub Reset()
'coopies the original data from sheet2 and paste to replace the one in sheet2
Sheets("Original Data").Columns("A:F").Copy Sheets("February").Columns("A:F")

'Clears the contents of Range("I6:I11") and Range("I14:I16") in sheet1
With Sheets("February")
    .Range("I6:I11").ClearContents
    .Range("I14:I16").ClearContents
End With
End Sub
```

### Buttons
Two button are created and assigned to the subroutines to runs the subroutines when clicked
As shown below. 
![Screenshot 2024-07-11 062742](https://github.com/dannieRope/Automating-Monthly-sales-Reporting-processes-VBA-Project/assets/132214828/e1520d0b-e34f-45f6-8a4d-68e064ffd02d)

The image below shows what happens when the run button is clicked 

![Screenshot 2024-07-11 063337](https://github.com/dannieRope/Automating-Monthly-sales-Reporting-processes-VBA-Project/assets/132214828/134dafda-6929-488d-8aea-a8ae063d80cd)

Also the below shows what happens when the reset button is clicked

![Screenshot 2024-07-11 063808](https://github.com/dannieRope/Automating-Monthly-sales-Reporting-processes-VBA-Project/assets/132214828/9d3b5239-9d5a-4b4d-994e-ed33ea47bbc5)


Thanks for reading.
Thanks For reading and Feel free to comment, share and correct the codes in case of an error. I would also love your feedbacks.
Find the the workbook [here]() and the VBA script [here]()


## License
[MIT License](LICENSE)





