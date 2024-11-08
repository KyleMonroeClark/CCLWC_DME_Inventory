Sub RecordInventory()
    Dim invSheet As Worksheet
    Dim recordSheet As Worksheet
    Dim partID As String
    Dim quantity As Long
    Dim foundCell As Range
    Dim i As Long
    Dim notFoundIDs As String
    Dim invCountCol As Long
    Dim partIDCol As Long
    Dim quantityCol As Long
    
    ' Set references to the sheets
    Set invSheet = ThisWorkbook.Sheets("INVENTORY")
    Set recordSheet = ThisWorkbook.Sheets("INVENTORY RECORDING")

    ' Define column numbers for ease of reading
    partIDCol = 1 ' Part ID column in INVENTORY RECORDING
    quantityCol = 2 ' Quantity column in INVENTORY RECORDING
    invCountCol = 14 ' "Inv. Count" column in INVENTORY (12 columns after Part ID)
    
    notFoundIDs = "" ' Initialize empty string for not found IDs

    ' Loop through each row in the INVENTORY RECORDING sheet
    For i = 2 To 25 ' From row 2 to 25
        partID = Trim(recordSheet.Cells(i, partIDCol).Value)
        quantity = recordSheet.Cells(i, quantityCol).Value
        
        ' Skip if part ID is blank or quantity is missing
        If partID = "" Or IsEmpty(quantity) Then
            If partID <> "" Then
                MsgBox "Quantity missing for Part ID: " & partID, vbExclamation
            End If
            GoTo NextRow
        End If

        ' Find the part ID in the INVENTORY sheet
        Set foundCell = invSheet.Columns(partIDCol).Find(What:=partID, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not foundCell Is Nothing Then
            ' Increase the inventory count in the INVENTORY sheet
            invSheet.Cells(foundCell.Row, invCountCol).Value = invSheet.Cells(foundCell.Row, invCountCol).Value + quantity
        Else
            ' If part ID is not found, add to the notFoundIDs string
            notFoundIDs = notFoundIDs & partID & vbCrLf
        End If

NextRow:
    Next i

    ' Display message based on the outcome
    If notFoundIDs = "" Then
        MsgBox "Inventory successfully recorded."
    Else
        MsgBox "The following Part IDs were not found in the inventory:" & vbCrLf & notFoundIDs, vbExclamation
    End If
End Sub

