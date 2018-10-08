Attribute VB_Name = "a_FileProcessor"
Function ExxonProcessor()
    Dim fueldata

    With Sheet2
        failnames = ""
        failnums = ""
        
        'read data into array
        lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
        fueldata = .Range("A3:BJ" & lastrow).Value
        ReDim Preserve fueldata(1 To UBound(fueldata, 1), 1 To UBound(fueldata, 2) + 1) 'overflow column if necessary
        
        'check for errors
        For i = 1 To UBound(fueldata, 1)
            If IsError(Application.VLookup(fueldata(i, 5), Sheet5.Range("J:K"), 2, 0)) Then
                If InStr(failnames, fueldata(i, 5)) = 0 Then failnames = failnames + fueldata(i, 5) + ", "
            ElseIf IsError(Application.VLookup(Application.WorksheetFunction.VLookup(fueldata(i, 5), Sheet5.Range("J:K"), 2, 0), Sheet5.Range("A:B"), 2, 0)) Then
                If InStr(failnums, Application.VLookup(fueldata(i, 5), Sheet5.Range("J:K"), 2, 0)) = 0 Then failnums = failnums + Application.VLookup(fueldata(i, 5), Sheet5.Range("J:K"), 2, 0) + ", "
            End If
        Next i
        
        If Len(failnums) > 0 Or Len(failnames) > 0 Then
            MsgBox ("ERROR - Process Aborted" + vbNewLine + vbNewLine + "These store names are missing from Lookup Column J: " + vbNewLine + "  " + failnames + vbNewLine + "These stores are missing from Lookup Column A: " + vbNewLine + "  " + failnums)
            ExxonProcessor = 0
            Exit Function
        End If
        
        'move columns correctly to different columns of the array
        For i = 1 To UBound(fueldata, 1)
            fueldata(i, 2) = fueldata(i, 5)
            fueldata(i, 3) = fueldata(i, 9)
            fueldata(i, 4) = fueldata(i, 11)
            fueldata(i, 5) = fueldata(i, 12)
            fueldata(i, 6) = fueldata(i, 35)
            fueldata(i, 7) = fueldata(i, 37)
            fueldata(i, 8) = fueldata(i, 38)
            fueldata(i, 9) = "EXXON"
            fueldata(i, 10) = fueldata(i, 52)
            fueldata(i, 11) = Application.WorksheetFunction.VLookup(fueldata(i, 2), Sheet5.Range("J:K"), 2, 0)
            fueldata(i, 2) = Application.WorksheetFunction.VLookup(fueldata(i, 11), Sheet5.Range("A:B"), 2, 0)
            fueldata(i, 12) = StrConv(fueldata(i, 53) & " " & fueldata(i, 10), vbProperCase)
            fueldata(i, 13) = Month(fueldata(i, 1))
            fueldata(i, 14) = Day(fueldata(i, 1))
        Next i
        
        'shave extra columns off array
        ReDim Preserve fueldata(1 To UBound(fueldata, 1), 1 To 14)
        
        'push array back to sheet
        .Cells.Clear
        .Range("A1:" & Split(Cells(1, UBound(fueldata, 2)).Address, "$")(1) & UBound(fueldata, 1)).Value = fueldata
        .Range("A:Z").Sort key1:=.Range("K:K"), Header:=xlNo
        
    End With
    
    
    Call Dangerzone
    MsgBox ("All Done")
    ExxonProcessor = 1
End Function

Function ChaseProcessor()
    Dim fueldata
    
        failnames = ""
        failnums = ""

    With Sheet2
        'read data into array
        lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
        fueldata = .Range("A3:M" & lastrow).Value
        ReDim Preserve fueldata(1 To UBound(fueldata, 1), 1 To UBound(fueldata, 2) + 1) 'overflow column if necessary
        
        'check for errors
        For i = 1 To UBound(fueldata, 1)
            If IsError(Application.VLookup(left(fueldata(i, 2), 4), Sheet5.Range("A:B"), 2, 0)) Then
                If InStr(failnums, left(fueldata(i, 2), 4)) = 0 Then failnums = failnums + left(fueldata(i, 2), 4) + ", "
            End If
        Next i
        
        If Len(failnums) > 0 Or Len(failnames) > 0 Then
            MsgBox ("ERROR - Process Aborted" + vbNewLine + vbNewLine + "These stores are missing from Lookup Column A: " + vbNewLine + "  " + failnums)
            ChaseProcessor = 0
            Exit Function
        End If
        
        'move columns correctly to different columns of the array
        For i = 1 To UBound(fueldata, 1)
            'check for pointless data
            If left(fueldata(i, 2), 1) <> "L" Then
                For n = 1 To 14
                    fueldata(i, n) = ""
                Next n
            Else
            fueldata(i, 13) = fueldata(i, 12)
            fueldata(i, 12) = StrConv(fueldata(i, 1), vbProperCase)
            fueldata(i, 1) = fueldata(i, 4)
            fueldata(i, 5) = fueldata(i, 10)
            fueldata(i, 3) = fueldata(i, 9)
            fueldata(i, 4) = fueldata(i, 13)
            fueldata(i, 9) = fueldata(i, 6)
            fueldata(i, 6) = fueldata(i, 8)
            fueldata(i, 8) = fueldata(i, 7)
            fueldata(i, 7) = fueldata(i, 9)
            fueldata(i, 9) = "CHASE"
            fueldata(i, 10) = "CHASE"
            fueldata(i, 11) = left(fueldata(i, 2), 4)
            fueldata(i, 13) = Month(fueldata(i, 1))
            fueldata(i, 14) = Day(fueldata(i, 1))
            fueldata(i, 2) = Application.WorksheetFunction.VLookup(fueldata(i, 11), Sheet5.Range("A:B"), 2, 0)
            End If
        Next i
        
        'push array back to sheet
        .Cells.Clear
        .Range("A1:" & Split(Cells(1, UBound(fueldata, 2)).Address, "$")(1) & UBound(fueldata, 1)).Value = fueldata
        .Range("A:Z").Sort key1:=.Range("K:K"), Header:=xlNo
        
    End With
    
    Call Dangerzone
    MsgBox ("All Done")
    ChaseProcessor = 1
End Function

Function FuelmanProcessor()
    Dim fueldata

    With Sheet2
        'read data into array
        lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
        fueldata = .Range("A16:AJ" & lastrow).Value
        ReDim Preserve fueldata(1 To UBound(fueldata, 1), 1 To UBound(fueldata, 2) + 1) 'overflow column if necessary
        
        'move columns correctly to different columns of the array
         For i = 1 To UBound(fueldata, 1)
            fueldata(i, 1) = fueldata(i, 5)
            fueldata(i, 3) = fueldata(i, 32)
            fueldata(i, 4) = fueldata(i, 33)
            fueldata(i, 5) = fueldata(i, 36)
            fueldata(i, 6) = fueldata(i, 10)
            fueldata(i, 7) = fueldata(i, 11)
            fueldata(i, 8) = fueldata(i, 12)
            fueldata(i, 9) = "FUELMAN"
            fueldata(i, 10) = fueldata(i, 16)
            fueldata(i, 11) = Application.WorksheetFunction.VLookup(fueldata(i, 2), Sheet5.Range("AY:AZ"), 2, 0)
            fueldata(i, 2) = Application.WorksheetFunction.VLookup(fueldata(i, 11), Sheet5.Range("A:B"), 2, 0)
            fueldata(i, 12) = StrConv(fueldata(i, 15) & " " & fueldata(i, 16), vbProperCase)
            fueldata(i, 13) = Month(fueldata(i, 1))
            fueldata(i, 14) = Day(fueldata(i, 1))
        Next i
        
        'shave extra columns off array
        ReDim Preserve fueldata(1 To UBound(fueldata, 1), 1 To 14)
        
        'push array back to sheet
        .Cells.Clear
        .Range("A1:" & Split(Cells(1, UBound(fueldata, 2)).Address, "$")(1) & UBound(fueldata, 1)).Value = fueldata
        .Range("A:Z").Sort key1:=.Range("K:K"), Header:=xlNo
        
    End With
    
    Call Dangerzone
    MsgBox ("All Done")
End Function

Sub Dangerzone()
    Sheet2.Range("A1:N1").Insert shift:=xlShiftDown
    
    Sheet2.Range("a1:n1") = Split("Transaction Date|Account Name|Units|Unit Cost|Total Fuel Cost|Merchant Name|Merchant City|Merchant State / Province|Driver First Name|Driver Last Name|Store#|Card Name|Month|Day", "|")

    'filter entries that have an unusual number of gallons purchased
    Sheet2.Range("A:N").AdvancedFilter _
        Action:=xlFilterInPlace, _
        criteriarange:=Sheet5.Range("BB1:BB2")
    
    lastrow = Sheet2.Cells(Sheet2.Rows.Count, "A").End(xlUp).Row
    varstore = ""
    
    'build a string of stores with strange transactions
    If lastrow > 1 Then
        For Each cell In Sheet2.Range("K2:K" & lastrow).SpecialCells(xlVisible)
            If Not InStr(varstore, cell.Value) > 0 Then
                varstore = cell.Value & ", " & varstore
            End If
        Next cell
    End If
    
    Sheet2.ShowAllData
    
    If Len(varstore) > 0 Then
        MsgBox ("The following stores have an unusual transaction:" & vbNewLine & vbNewLine & left(varstore, Len(varstore) - 2))
    Else
        MsgBox ("No unusual purchases.")
    End If
    
    Sheet2.Range("A1:N1").Delete shift:=xlUp
    
End Sub

Function holding()
    'figure out first open row in the holding sheet and push data to it
    lastrowdst = Sheet3.Cells(Sheet3.Rows.Count, "A").End(xlUp).Row
    If lastrowdst = 1 And Sheet3.Range("A1") = "" Then lastrowdst = 0
    
    lastrowsrc = Sheet2.Cells(Sheet2.Rows.Count, "A").End(xlUp).Row
    
    Sheet2.Rows("1:" & lastrowsrc).Cut
    Sheet3.Rows(lastrowdst + 1).Insert shift:=xlDown
    
End Function

Function InventoryProcessor()
    'find column that starts with "Total"
    'delete all columns after
        Sheet2.Columns("K:M").Delete
        Sheet2.Columns("B:I").Delete
    'delete columns after A but before TOTAL
    'Delete top row
        Sheet2.Rows(1).EntireRow.Delete
        lastrow = Sheet2.Cells(Sheet2.Rows.Count, "A").End(xlUp).Row
        Sheet2.Rows(lastrow).EntireRow.Delete
        lastrow = Sheet2.Cells(Sheet2.Rows.Count, "A").End(xlUp).Row
    'insert column before B
        Sheet2.Columns("B").EntireColumn.Insert
            
    'use A to lookup storenumber and put into B
        Sheet2.Range("B1").Value = "Store#"
        For Each cell In Sheet2.Range("A2:a" & lastrow)
            If cell.Value = "Total" Then cell.Row.Delete
            Sheet2.Range("B" & cell.Row).Value = Application.WorksheetFunction.VLookup(cell.Value, Sheet5.Range("AV:AW"), 2, 0)
        Next cell
        Sheet2.Columns("A").Delete
End Function

Function invholding()
    lastrowdst = Sheet3.Cells(Sheet3.Rows.Count, "A").End(xlUp).Row
    If lastrowdst = 1 And Sheet3.Range("A1") = "" Then lastrowdst = 0
    
    lastrowsrc = Sheet2.Cells(Sheet2.Rows.Count, "A").End(xlUp).Row
    
    If Sheet2.Range("B1").Value = Sheet3.Range("B1").Value Then
        Sheet2.Rows("2:" & lastrowsrc).Cut
        Sheet3.Rows(lastrowdst + 1).Insert shift:=xlDown
    Else
        If Sheet3.Range("B1") = "" Then
            invconflict = 6
        Else
            invconflict = MsgBox("The month and/or category are different than the data in the holding sheet. Overwrite the data?", vbYesNo)
        End If
        If invconflict = 6 Then
            Sheet3.Cells.Clear
            Sheet2.Rows("1:" & lastrowsrc).Cut
            Sheet3.Rows(1).Insert shift:=xlDown
        Else
            abort = MsgBox("Operation aborted.", vbOKOnly)
            Exit Function
        End If
    End If
    
End Function
