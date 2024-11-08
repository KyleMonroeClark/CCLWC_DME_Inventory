Sub PrintAndTransferData()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim wsInventory As Worksheet
    Dim lastRow As Long
    Dim NextRow As Long
    Dim partID As String
    Dim quantity As Variant

    ' Define worksheets
    On Error GoTo ErrorHandler
    Set wsSource = ThisWorkbook.Sheets("MEDICARE POD")
    Set wsTarget = ThisWorkbook.Sheets("ORDER LIST")
    Set wsInventory = ThisWorkbook.Sheets("INVENTORY")
    
    ' Print the current sheet
    wsSource.PrintOut

    ' Find the last row with data in ORDER LIST
    lastRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row
    
    ' Determine the next empty row in ORDER LIST
    NextRow = lastRow + 1
    
    ' Transfer data from MEDICARE POD to ORDER LIST
    With wsSource
        wsTarget.Cells(NextRow, 1).Value = .Range("B7").MergeArea.Cells(1, 1).Value ' Name
        wsTarget.Cells(NextRow, 2).Value = .Range("B8").MergeArea.Cells(1, 1).Value ' Address
        wsTarget.Cells(NextRow, 3).Value = .Range("B9").MergeArea.Cells(1, 1).Value ' Phone Number
        wsTarget.Cells(NextRow, 4).Value = .Range("B10").MergeArea.Cells(1, 1).Value ' Date of Birth
        wsTarget.Cells(NextRow, 5).Value = "Medicare" ' Insurance
        wsTarget.Cells(NextRow, 6).Value = Date ' Date of Service

        ' Part 1
        partID = Trim(.Range("A13").MergeArea.Cells(1, 1).Value)
        quantity = .Range("B13").MergeArea.Cells(1, 1).Value
        If partID <> "" Then
            If IsEmpty(quantity) Or quantity = "" Then
                MsgBox "Please add a quantity for part #1.", vbExclamation
                Exit Sub
            End If
            wsTarget.Cells(NextRow, 7).Value = partID
            wsTarget.Cells(NextRow, 8).Value = quantity
            wsTarget.Cells(NextRow, 9).Value = .Range("C13").MergeArea.Cells(1, 1).Value ' Side 1
            AdjustInventory wsInventory, partID, CDbl(quantity)
        End If

        ' Part 2
        partID = Trim(.Range("A19").MergeArea.Cells(1, 1).Value)
        quantity = .Range("B19").MergeArea.Cells(1, 1).Value
        If partID <> "" Then
            If IsEmpty(quantity) Or quantity = "" Then
                MsgBox "Please add a quantity for part #2.", vbExclamation
                Exit Sub
            End If
            wsTarget.Cells(NextRow, 10).Value = partID
            wsTarget.Cells(NextRow, 11).Value = quantity
            wsTarget.Cells(NextRow, 12).Value = .Range("C19").MergeArea.Cells(1, 1).Value ' Side 2
            AdjustInventory wsInventory, partID, CDbl(quantity)
        End If

        ' Part 3
        partID = Trim(.Range("A25").MergeArea.Cells(1, 1).Value)
        quantity = .Range("B25").MergeArea.Cells(1, 1).Value
        If partID <> "" Then
            If IsEmpty(quantity) Or quantity = "" Then
                MsgBox "Please add a quantity for part #3.", vbExclamation
                Exit Sub
            End If
            wsTarget.Cells(NextRow, 13).Value = partID
            wsTarget.Cells(NextRow, 14).Value = quantity
            wsTarget.Cells(NextRow, 15).Value = .Range("C25").MergeArea.Cells(1, 1).Value ' Side 3
            AdjustInventory wsInventory, partID, CDbl(quantity)
        End If

        ' Part 4
        partID = Trim(.Range("A31").MergeArea.Cells(1, 1).Value)
        quantity = .Range("B31").MergeArea.Cells(1, 1).Value
        If partID <> "" Then
            If IsEmpty(quantity) Or quantity = "" Then
                MsgBox "Please add a quantity for part #4.", vbExclamation
                Exit Sub
            End If
            wsTarget.Cells(NextRow, 16).Value = partID
            wsTarget.Cells(NextRow, 17).Value = quantity
            wsTarget.Cells(NextRow, 18).Value = .Range("C31").MergeArea.Cells(1, 1).Value ' Side 4
            AdjustInventory wsInventory, partID, CDbl(quantity)
        End If
    End With
    
    ' Inform the user
    MsgBox "Data transferred successfully and sheet printed.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

Sub AdjustInventory(wsInventory As Worksheet, partID As String, quantity As Double)
    Dim foundRow As Long
    Dim garmentsPerPackage As Double
    Dim currentInventory As Double
    Dim lastRow As Long
    Dim i As Long

    ' Find the last row in the inventory sheet
    lastRow = wsInventory.Cells(wsInventory.Rows.Count, "A").End(xlUp).Row

    ' Loop through the rows in the inventory sheet to find the part ID
    foundRow = 0
    For i = 1 To lastRow
        If Trim(wsInventory.Cells(i, "A").Value) = partID Then
            foundRow = i
            Exit For
        End If
    Next i

    If foundRow > 0 Then
        ' Retrieve the value in the S column (garments per package)
        garmentsPerPackage = wsInventory.Cells(foundRow, "S").Value

        ' Adjust the quantity based on garmentsPerPackage
        If garmentsPerPackage = 2 Then
            quantity = quantity / 2
        End If

        ' Get the current inventory count from the N column
        currentInventory = wsInventory.Cells(foundRow, "N").Value

        ' If the current inventory is empty, treat it as 0
        If IsEmpty(currentInventory) Then
            currentInventory = 0
        End If

        ' Update the inventory count
        wsInventory.Cells(foundRow, "N").Value = currentInventory - quantity
    Else
        MsgBox "Part ID '" & partID & "' not found in the inventory list.", vbExclamation
    End If
End Sub

