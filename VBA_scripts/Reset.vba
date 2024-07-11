Sub Reset()
'coopies the original data from sheet2 and paste to replace the one in sheet2
Sheets("Original Data").Columns("A:F").Copy Sheets("February").Columns("A:F")
With Sheets("February")
    .Range("I6:I11").ClearContents
    .Range("I14:I16").ClearContents
End With
End Sub
