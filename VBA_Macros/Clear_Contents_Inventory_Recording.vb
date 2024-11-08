Sub ClearContents()
    ' Clear contents in the INVENTORY RECORDING sheet from A2 to B25
    With ThisWorkbook.Sheets("INVENTORY RECORDING")
        .Range("A2:B25").ClearContents
    End With
End Sub
