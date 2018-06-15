Attribute VB_Name = "b_InventoryWriter"
Function FuelWriter()
    Dim fuelform As FuelPusher
    Dim destws As Worksheet
    Dim srcws As Worksheet
    Dim rawwb As Workbook
    Dim analyzerwb As Workbook
    
    
    Set fuelform = New FuelPusher

    fNameAndPath = Application.GetOpenFilename(FileFilter:="All Files, *", Title:="Where is the Fuel Analyzer Workbook?")
    Set rawwb = ActiveWorkbook
    If fNameAndPath = False Then
        Set analyzerwb = Nothing
        FuelWriter = True
        Exit Function
    Else
        Set analyzerwb = Workbooks.Open(fNameAndPath)
    End If
    
    Set srcws = rawwb.Sheets("Holding Data")
    Set destws = analyzerwb.Worksheets("Compiled Fuel Data")
    
    searchstring = srcws.Range("M1").Value
    found = Application.Match(searchstring, destws.Range("M:M"), 0)
    
    lastrowsrc = srcws.Cells(srcws.Rows.Count, "A").End(xlUp).Row
    lastrowdest = destws.Cells(destws.Rows.Count, "A").End(xlUp).Row
    
    If Not IsError(found) Then
        fuelform.Show
        Select Case fuelform.Tag
            Case 0
                fuelform.Hide
                GoTo funexit
            Case 1
                'replace
                'filter for the month in question
                    srcws.Range("AA1").Value = "Month"
                    srcws.Range("AA2").Value = searchstring
                    
                    destws.Range("A:N").AdvancedFilter _
                        Action:=xlFilterInPlace, _
                        criteriarange:=srcws.Range("AA1:AA2")
                    
                    destws.Range("A2:A" & lastrowdest).SpecialCells(xlCellTypeVisible).EntireRow.Delete
                    
                    destws.ShowAllData
                    
                    lastrowdest = destws.Cells(destws.Rows.Count, "A").End(xlUp).Row
                    
                    srcws.Range("A1:N" & lastrowsrc).Cut
                    destws.Range("a" & lastrowdest + 1).Insert shift:=xlDown
                    
                
            Case 2
                'append
                srcws.Range("A1:N" & lastrowsrc).Cut
                destws.Range("a" & lastrowdest + 1).Insert shift:=xlDown
        End Select
    
    Else
                srcws.Range("A1:N" & lastrowsrc).Cut
                destws.Range("a" & lastrowdest + 1).Insert shift:=xlDown
    End If
    
    
    
funexit:
        analyzerwb.Close
        Set analyzerwb = Nothing
End Function
Function inventorywriter(srcsheetname As String)
'push finished data to final resting place
    fNameAndPath = Application.GetOpenFilename(FileFilter:="All Files, *", Title:="Where is the Fuel Analyzer Workbook?")
    Set rawwb = ActiveWorkbook
    If fNameAndPath = False Then
        Set analyzerwb = Nothing
        inventorywriter = True
        Exit Function
    Else
        Set analyzerwb = Workbooks.Open(fNameAndPath)
    End If
    
    Set srcws = rawwb.Sheets(srcsheetname)
    Set destws = analyzerwb.Worksheets("Inventory Data")

'search to see if relevant column already exists
    searchstring = srcws.Range("B1").Value
    datacol = Application.Match(searchstring, destws.Range("A1:AA1"), 0)

'if it does, throw a message and ask if it's an update
    If Not IsError(datacol) Then
        yesno = MsgBox("This data already exists. Would you like to overwrite it?", vbYesNo)
        If yesno = 7 Then
            analyzerwb.Close
            Exit Function
        End If
    Else
'if data being added then find first open column
        datacol = destws.Cells(1, 26).End(xlToLeft).Column + 1
    End If


'anything that throws exception to lookup is probably a new store, create an entry
    
    destws.Cells(1, datacol).Value = srcws.Cells(1, 2).Value

lastrow = srcws.Cells(srcws.Rows.Count, "A").End(xlUp).Row
    For Each store2 In srcws.Range("A2:A" & lastrow)
    Set found = destws.Range("A:A").Find(store2.Value)
        If found Is Nothing Then
            lastrow = destws.Cells(destws.Rows.Count, "A").End(xlUp).Row + 1
            destws.Cells(lastrow, 1).Value = store2.Value
            destws.Cells(lastrow, datacol).Value = srcws.Cells(store2.Row, 2)
        End If
    Next store2

lastrow = destws.Cells(destws.Rows.Count, "A").End(xlUp).Row
Dim storeinv As Variant
    For Each store In destws.Range("A2:A" & lastrow)
        
        storeinv = Application.VLookup(store, srcws.Range("A:B"), 2, 0)
        If Not IsError(storeinv) Then
            destws.Cells(store.Row, datacol).Value = storeinv
        Else
            destws.Cells(store.Row, datacol).Value = 0
        End If
    Next store
    
'then go fix the inventory part in the other spreadsheet
    analyzerwb.Close
    
End Function



