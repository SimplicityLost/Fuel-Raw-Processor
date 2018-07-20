Attribute VB_Name = "a_FileProcessor"
Function holding()
    lastrowdst = Sheet3.Cells(Sheet3.Rows.Count, "A").End(xlUp).Row
    If lastrowdst = 1 And Sheet3.Range("A1") = "" Then lastrowdst = 0
    
    lastrowsrc = Sheet2.Cells(Sheet2.Rows.Count, "A").End(xlUp).Row
    
    Sheet2.Rows("1:" & lastrowsrc).Cut
    Sheet3.Rows(lastrowdst + 1).Insert shift:=xlDown
    
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

Function fuelmanprocessor()
    fnd = Application.WorksheetFunction.Match("Account Code", Sheet2.Range("A:A"), 0)
    Sheet2.Range("A1:A" & fnd + 1).EntireRow.Delete
    'ak-ay,ah-ai, q-ae, m-n,f-i,d,a
    Sheet2.Columns("AK:AY").Delete
    Sheet2.Columns("Ah:Ai").Delete
    Sheet2.Columns("q:Ae").Delete
    Sheet2.Columns("m:n").Delete
    Sheet2.Columns("f:i").Delete
    Sheet2.Columns("d").Delete
    Sheet2.Columns("a").Delete
    Sheet2.Columns("I:K").Cut
    Sheet2.Columns("C").Insert shift:=xlRight
    Sheet2.Columns("F").Cut
    Sheet2.Columns("A").Insert shift:=xlRight
    
    lastrow = Sheet2.Cells(Sheet2.Rows.Count, "A").End(xlUp).Row

    For Each cell In Sheet2.Range("A1:A" & lastrow)
        Sheet2.Range("L" & cell.Row).Value = Application.WorksheetFunction.VLookup(Sheet2.Range("B" & cell.Row).Value, Sheet5.Range("Ay:az"), 2, 0)
        Sheet2.Range("C" & cell.Row).Value = Application.WorksheetFunction.VLookup(Sheet2.Range("L" & cell.Row).Value, Sheet5.Range("A:B"), 2, 0)
        Sheet2.Range("M" & cell.Row).Value = Application.WorksheetFunction.Proper(Sheet2.Range("j" & cell.Row).Value & " " & Sheet2.Range("k" & cell.Row).Value)
        Sheet2.Range("j" & cell.Row).Value = "FUELMAN"
        Sheet2.Range("N" & cell.Row) = Month(Sheet2.Range("A" & cell.Row))
        Sheet2.Range("O" & cell.Row) = Day(Sheet2.Range("A" & cell.Row))
    Next cell
    Sheet2.Cells.ClearFormats
    Sheet2.Range("A:A").NumberFormat = "mm/d/yyyy;@"
    Sheet2.Rows(1).Insert
    Sheet2.Range("D1").Value = "Units"
    Sheet2.Columns("b").Delete
    Call Dangerzone
    Sheet2.Rows(1).Delete
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

Function GLProcessor()
    'delete everything from journals 10,10L,3,11,16,20,20L,30,31,32,33,50,51,6,7,9
    'delete src jrnls 10,10L,3,6,7,9
    'delete everything with exxon or fuel in Q,M,N,I,H
    'delete everything with acq in N
        If Sheet2.ListObjects.Count = 1 Then Sheet2.ListObjects(1).Unlist
        Sheet2.Cells.ClearFormats
        lastrow = Sheet2.Cells(Sheet2.Rows.Count, "A").End(xlUp).Row
        Sheet2.Range("A:Z").AdvancedFilter _
            Action:=xlFilterInPlace, _
            criteriarange:=Sheet5.Range("M1:AT35")
        Sheet2.Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        Sheet2.ShowAllData
    'use search strings to add vendors to things
        lastrow = Sheet2.Cells(Sheet2.Rows.Count, "A").End(xlUp).Row
        lastcrit = Sheet5.Cells(Sheet5.Rows.Count, "N").End(xlUp).Row
        Sheet2.Range("AA1").Value = "Vendor"
        For Each cell In Sheet5.Range("N37:N" & lastcrit)
            Sheet5.Range("O37").Value = cell.Value
            Sheet2.Range("A:AA").AdvancedFilter _
                Action:=xlFilterInPlace, _
                criteriarange:=Sheet5.Range("O36:O37")
            If Not Sheet2.Range("A" & Sheet2.Rows.Count).End(xlUp).Row = 1 Then
                For Each cell2 In Sheet2.Range("AA2:AA" & lastrow).SpecialCells(xlCellTypeVisible)
                    cell2.Value = Sheet5.Range("M" & cell.Row).Value
                Next cell2
            End If
            Sheet2.ShowAllData
        Next cell
        
        lastcrit = Sheet5.Cells(Sheet5.Rows.Count, "Q").End(xlUp).Row
        For Each cell In Sheet5.Range("Q37:Q" & lastcrit)
            Sheet5.Range("R37").Value = cell.Value
            Sheet2.Range("A:AA").AdvancedFilter _
                Action:=xlFilterInPlace, _
                criteriarange:=Sheet5.Range("R36:R37")
            If Not Sheet2.Range("A" & Sheet2.Rows.Count).End(xlUp).Row = 1 Then
                For Each cell2 In Sheet2.Range("AA2:AA" & lastrow).SpecialCells(xlCellTypeVisible)
                    cell2.Value = Sheet5.Range("P" & cell.Row).Value
                Next cell2
            End If
            Sheet2.ShowAllData
        Next cell

    'delete non-vendored entries
        lastrow = Sheet2.Cells(Sheet2.Rows.Count, "A").End(xlUp).Row
        Sheet2.Range("A:AA").AdvancedFilter _
            Action:=xlFilterInPlace, _
            criteriarange:=Sheet5.Range("S36:S37")
        Sheet2.Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        Sheet2.ShowAllData
    'convert the entries to fuel data lines
        'delete columns s-z,m-q,c-k
        Sheet2.Columns("S:Z").Delete
        Sheet2.Columns("m:q").Delete
        Sheet2.Columns("c:k").Delete
        'use A to build store number in K
        lastrow = Sheet2.Cells(Sheet2.Rows.Count, "A").End(xlUp).Row
        For Each cell In Sheet2.Range("A2:A" & lastrow)
            strnum = cell.Value
            For i = 0 To 2 - Len(cell.Value)
                strnum = "0" & strnum
            Next i
                strnum = "L" & strnum
            Sheet2.Range("k" & cell.Row) = strnum
        Next cell
        
        'use K to build store name back into B
        For Each cell In Sheet2.Range("b2:b" & lastrow)
            strname = Application.WorksheetFunction.VLookup(Sheet2.Range("k" & cell.Row).Value, Sheet5.Columns("A:B"), 2, False)
            cell.Value = strname
        Next cell
        'copy posted date into A
        Sheet2.Columns("A").Value = Sheet2.Columns("D").Value
        Sheet2.Columns("a").NumberFormat = "M/dd/yy"
        'make sure amt is in D
        'put vendor name into f, unknown in g
        Sheet2.Columns("F").Value = Sheet2.Columns("E").Value
        Sheet2.Columns("E").Value = Sheet2.Columns("C").Value
        For Each cell In Sheet2.Range("b2:b" & lastrow)
            Sheet2.Range("G" & cell.Row) = "Unknown"
            Sheet2.Range("I" & cell.Row) = "Unknown"
            Sheet2.Range("J" & cell.Row) = "Unknown"
            Sheet2.Range("L" & cell.Row) = "Unknown"
            Sheet2.Range("M" & cell.Row) = Month(Sheet2.Range("A" & cell.Row))
            Sheet2.Range("N" & cell.Row) = Day(Sheet2.Range("A" & cell.Row))
            Sheet2.Range("H" & cell.Row) = Application.WorksheetFunction.VLookup(Sheet2.Range("k" & cell.Row).Value, Sheet5.Columns("c:d"), 2, False)
            Sheet2.Range("D" & cell.Row) = Application.WorksheetFunction.VLookup(Sheet2.Range("h" & cell.Row).Value, Sheet5.Columns("f:g"), 2, False)
            Sheet2.Range("C" & cell.Row) = Round(Sheet2.Range("E" & cell.Row).Value / Sheet2.Range("D" & cell.Row), 3)
        Next cell
        'use lookup to get state
            
        'use state and amt to get unit cost and units
        'put vendor name into f, unknown in g
        'build m out of posted date
        'set driver names to Unknown
        
        Call Dangerzone
        
        Sheet2.Rows(1).Delete
        MsgBox ("All Done")

End Function

Function ExxonProcessor()
    'delete columns bb - bi
    'delete columns am-ay
    'delete aj, m-ah, j, f-h, b-d
        Sheet2.Columns("bb:bi").Delete
        Sheet2.Columns("am:ay").Delete
        Sheet2.Columns("aj").Delete
        Sheet2.Columns("m:ah").Delete
        Sheet2.Columns("j").Delete
        Sheet2.Columns("f:h").Delete
        Sheet2.Columns("d").Delete
        Sheet2.Columns("b:c").Delete

    'move j to i
        Sheet2.Columns("j").Cut
        Sheet2.Columns("i").Insert shift:=xlRight
        
        lastrow = Sheet2.Cells(Sheet2.Rows.Count, "A").End(xlUp).Row

        Sheet2.Range("A:N").AdvancedFilter _
            Action:=xlFilterInPlace, _
            criteriarange:=Sheet5.Range("BD1:BF4")
        If Sheet2.Range("A1:A" & lastrow).SpecialCells(xlVisible).Count > 1 Then
            Sheet2.Range("A2:A" & lastrow).SpecialCells(xlVisible).EntireRow.Delete
        End If
        Sheet2.ShowAllData
        
    'use b to get store nums in k
    'use k to standardize names in b
    lastrow = Sheet2.Cells(Sheet2.Rows.Count, "A").End(xlUp).Row
        For Each cell In Sheet2.Range("B2:B" & lastrow)
            strnum = Application.WorksheetFunction.VLookup(cell.Value, Sheet5.Columns("J:K"), 2, False)
            Sheet2.Range("K" & cell.Row).Value = strnum
            
            strname = Application.WorksheetFunction.VLookup(strnum, Sheet5.Columns("A:B"), 2, False)
            cell.Value = strname
        Next cell

    'concatenate i and j into l
        For Each cell In Sheet2.Range("l2:l" & lastrow)
            cell.Value = Sheet2.Range("i" & cell.Row).Value & " " & Sheet2.Range("j" & cell.Row).Value
            cell.Value = WorksheetFunction.Proper(cell.Value)
            Sheet2.Range("i" & cell.Row).Value = "EXXON"
        Next cell
    'use a to put month in m
        For Each cell In Sheet2.Range("m2:m" & lastrow)
            cell.Value = Month(Sheet2.Range("a" & cell.Row).Value)
            Sheet2.Range("N" & cell.Row) = Day(Sheet2.Range("A" & cell.Row))
        Next cell
    'delete top row
        
        Call Dangerzone
        
        Sheet2.Rows(1).Delete
        MsgBox ("All Done")
End Function

Function ChaseProcessor()
    Dim fueldata

    With Sheet2
        'read data into array
        lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
        fueldata = .Range("A3:M" & lastrow).Value
        ReDim Preserve fueldata(1 To UBound(fueldata, 1), 1 To UBound(fueldata, 2) + 1)
        'move columns correctly to different columns of the array
        For i = 1 To UBound(fueldata, 1)
            'check for pointless data
            If Left(fueldata(i, 2), 1) <> "L" Then
                For n = 1 To 14
                    fueldata(i, n) = ""
                Next n
            Else
            fueldata(i, 13) = fueldata(i, 12)
            fueldata(i, 12) = fueldata(i, 1)
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
            fueldata(i, 11) = Left(fueldata(i, 2), 4)
            fueldata(i, 13) = Month(fueldata(i, 1))
            fueldata(i, 14) = Day(fueldata(i, 1))
            fueldata(i, 2) = Application.WorksheetFunction.VLookup(fueldata(i, 11), Sheet5.Range("A:B"), 2, 0)
            End If
        Next i
        
        'push array back to sheet
        .Cells.Clear
        ' .Range("A1:n1").Value = Split("Transaction Date|Account Name|Units|Unit Cost|Total Fuel Cost|Merchant Name|Merchant City|Merchant State / Province|Driver First Name|Driver Last Name|Store#|Card Name|Month|Day", "|")
        .Range("A1:" & Split(Cells(1, UBound(fueldata, 2)).Address, "$")(1) & UBound(fueldata, 1)).Value = fueldata
        .Range("A:Z").Sort key1:=.Range("K:K"), Header:=xlNo
        
        
    End With
    
    
    Call Dangerzone
    MsgBox ("All Done")
End Function

Function oldChaseProcessor()
    With Sheet2
    'delete top 2 rows
    .Rows("1:2").Delete
    
    lastrow = Sheet2.Cells(Sheet2.Rows.Count, "A").End(xlUp).Row
    
    'delete K
    .Columns("k").Delete
    'delete C
    .Columns("c").Delete
    'cut C->A
    .Columns("c").Cut
    .Columns("a").Insert shift:=xlRight
    'insert B
    .Columns("b").Insert shift:=xlRight
    'cut I->C
    .Columns("I").Cut
    .Columns("C").Insert shift:=xlRight
    'cut K->D
    .Columns("K").Cut
    .Columns("D").Insert shift:=xlRight
    'cut K->E
    .Columns("K").Cut
    .Columns("E").Insert shift:=xlRight
    'cut K->F
    .Columns("K").Cut
    .Columns("F").Insert shift:=xlRight
    'cut J->G [might be able to do
    'cut K->H  this in one step]
    .Columns("J:K").Cut
    .Columns("G").Insert shift:=xlRight
    
    'insert I-N
    .Columns("I:N").Insert shift:=xlRight
    'walk through each row
    For i = 1 To lastrow
        If Left(.Range("P" & i), 1) <> "L" Then
            .Rows(i).Delete
            GoTo oops:
        End If
        
        strnum = Left(.Range("P" & i), 4)
        
        'fill B with vlookup of left(P,4)
        .Range("B" & i).Value = Application.WorksheetFunction.VLookup(strnum, Sheet5.Range("A:B"), 2, 0)
        
        'fill I with CHASE
        .Range("I" & i).Value = "CHASE"
        
        'fill J with CHASE
        .Range("J" & i).Value = "CHASE"
        
        'fill K with left(P,4)
        .Range("K" & i).Value = strnum
        
        'fill L with STRCONV(O,vbProperCase)
        .Range("L" & i).Value = StrConv(.Range("O" & i).Value, vbProperCase)
        
        'fill M with Month(A)
        .Range("M" & i).Value = Month(.Range("A" & i).Value)
        
        'fill N with Day(A)
        .Range("N" & i).Value = Day(.Range("A" & i).Value)
oops:
    Next i
    'delete O-Z
    .Columns("O:Z").Delete
    
    End With
    Call Dangerzone
    MsgBox ("All Done")
End Function

Function Dangerzone()
    Sheet2.Range("A:N").AdvancedFilter _
        Action:=xlFilterInPlace, _
        criteriarange:=Sheet5.Range("BB1:BB2")
    
    lastrow = Sheet2.Cells(Sheet2.Rows.Count, "A").End(xlUp).Row
    varstore = ""
    If lastrow > 1 Then
        For Each cell In Sheet2.Range("K2:K" & lastrow).SpecialCells(xlVisible)
            If Not InStr(varstore, cell.Value) > 0 Then
                varstore = cell.Value & ", " & varstore
            End If
        Next cell
    End If
    
    Sheet2.ShowAllData
    
    Dangerzone = MsgBox("All Done!" & vbNewLine & "The following stores have an unusual transaction:" & vbNewLine & varstore)
End Function
