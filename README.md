# FieldTripCalculations
This macro takes raw data from a system and manipulates it so that a finance employee is quickly able to input how much certain bus drivers for certain trips need to be paid.

```vbscript
Public Sub main()

    Call setup_worksheets
    Call header_format

End Sub



Public Sub setup_worksheets()

    'create "final" and "original" sheet
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "original"
    ActiveSheet.Move After:=Worksheets(Worksheets.Count)
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "final"
    ActiveSheet.Move After:=Worksheets(Worksheets.Count)

End Sub



Public Sub header_format()

    'declarations
    Dim fnl As Worksheet: Set fnl = ThisWorkbook.Worksheets("final")
    Dim ogl As Worksheet: Set ogl = ThisWorkbook.Worksheets("original")

    'column headers
    With fnl
        With .Cells(2, 1)
            .Value = "Trip Date"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(2, 2)
            .Value = "Trip ID"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(2, 3)
            .Value = "Driver"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(2, 4)
            .Value = "Bill To"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(2, 5)
            .Value = ""
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(2, 6)
            .Value = "Requester"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(2, 7)
            .Value = "Total Hours"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(2, 8)
            .Value = "Driver Cost"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(2, 9)
            .Value = "Recalculated"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(2, 10)
            .Value = "Difference"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(2, 11)
            .Value = "Total Miles"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(2, 12)
            .Value = "Mileage Cost"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(2, 13)
            .Value = "Recalculated"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(2, 14)
            .Value = "Difference"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(2, 15)
            .Value = "Budget #"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
        With .Cells(2, 16)
            .Value = "Total Due"
            With .Font
            .Bold = True
            .Italic = True
            End With
        End With
    End With


End Sub


Public Sub main2()

    Call column_calculations
    Call totals

End Sub


Public Sub totals()
    
    Dim fnl As Worksheet: Set fnl = ThisWorkbook.Worksheets("final")
    lastrow1 = fnl.Range("A" & fnl.Rows.Count).End(xlUp).Row
    
    For j = 1 To 10
    
        fnl.Cells(lastrow1 + 5, j + 6) = Application.Sum(Range(fnl.Cells(3, j + 6), fnl.Cells(lastrow1, j + 6)))
    
    Next j
    
    fnl.Cells(lastrow1 + 5, j - 4).NumberFormat = "$#,##0.00"
    
    With fnl

        .Columns(8).NumberFormat = "$#,##0.00"
        .Columns(9).NumberFormat = "$#,##0.00"
        .Columns(10).NumberFormat = "$#,##0.00"
        .Columns(12).NumberFormat = "$#,##0.00"
        .Columns(13).NumberFormat = "$#,##0.00"
        .Columns(14).NumberFormat = "$#,##0.00"
        .Columns(16).NumberFormat = "$#,##0.00"

    End With
    
    fnl.UsedRange.Columns.AutoFit

End Sub

Public Sub column_calculations()

    Dim fnl As Worksheet: Set fnl = ThisWorkbook.Worksheets("final")
    Dim ogl As Worksheet: Set ogl = ThisWorkbook.Worksheets("original")
    Dim str1 As String
    lastrow = ogl.Range("B" & ogl.Rows.Count).End(xlUp).Row
    lastrow1 = fnl.Range("A" & fnl.Rows.Count).End(xlUp).Row

    'format columns
    With fnl

        .Columns(1).NumberFormat = "mm/dd/yyyy"
        .Columns(7).NumberFormat = "0.00"
        .Columns(8).NumberFormat = "$#,##0.00"
        .Columns(9).NumberFormat = "$#,##0.00"
        .Columns(10).NumberFormat = "$#,##0.00"
        .Columns(12).NumberFormat = "$#,##0.00"
        .Columns(13).NumberFormat = "$#,##0.00"
        .Columns(14).NumberFormat = "$#,##0.00"
        .Columns(16).NumberFormat = "$#,##0.00"

    End With


    'input data
    For i = 3 To lastrow

        With fnl

            .Cells(i, 1) = ogl.Cells(i + 3, 2)
            .Cells(i, 2) = ogl.Cells(i + 3, 4)
            .Cells(i, 4) = ogl.Cells(i + 3, 5)
            .Cells(i, 6) = ogl.Cells(i + 3, 6)
            .Cells(i, 7) = ogl.Cells(i + 3, 8)
            .Cells(i, 8) = ogl.Cells(i + 3, 9)
            .Cells(i, 9) = (fnl.Cells(i, 7) * 30)
            .Cells(i, 10) = fnl.Cells(i, 8) - fnl.Cells(i, 9)
            .Cells(i, 11) = ogl.Cells(i + 3, 11)
            .Cells(i, 12) = ogl.Cells(i + 3, 13)
            .Cells(i, 13) = (fnl.Cells(i, 11) * 1.25)
            .Cells(i, 14) = (fnl.Cells(i, 12) - fnl.Cells(i, 13))
            .Cells(i, 15) = ogl.Cells(i + 3, 16)
            .Cells(i, 16) = fnl.Cells(i, 8) + fnl.Cells(i, 12)
           
        End With


    Next i

    
End Sub



Public Sub highlighting()


    Dim fnl As Worksheet: Set fnl = ThisWorkbook.Worksheets("final")

    lastrow1 = fnl.Range("A" & fnl.Rows.Count).End(xlUp).Row
    'highlight sports cells (green)
    For i = 3 To lastrow1

        str1 = Left(fnl.Cells(i, 15).Value, 2)
        If str1 = "79" Then

            fnl.Range(fnl.Cells(i, 11), fnl.Cells(i, 15)).Interior.ColorIndex = 43

        End If

        If IsEmpty(fnl.Cells(i, 3)) Then
        Else

            fnl.Range(fnl.Cells(i, 7), fnl.Cells(i, 16)).Interior.ColorIndex = 33

        End If

    Next i


End Sub


Public Sub reset()

    Dim fnl As Worksheet: Set fnl = ThisWorkbook.Worksheets("final")
    Dim ogl As Worksheet: Set ogl = ThisWorkbook.Worksheets("original")
    Dim cp As Worksheet: Set cp = ThisWorkbook.Worksheets("Control_Panel")

    fnl.Delete
    ogl.Delete
    cp.Activate

End Sub



```
