Sub PopulateSheet2()
    Dim sheet1 As Worksheet
    Dim sheet2 As Worksheet
    Dim lastRowSheet1 As Long
    Dim lastRowSheet2 As Long
    Dim i As Long
    
    ' Set the sheet variables
    Set sheet1 = ThisWorkbook.Sheets("Sheet1")
    Set sheet2 = ThisWorkbook.Sheets("Sheet2")
    
    ' Get the last row of data in Sheet1
    lastRowSheet1 = sheet1.Cells(sheet1.Rows.Count, 1).End(xlUp).Row
    
    ' Get the last row of data in Sheet2
    lastRowSheet2 = sheet2.Cells(sheet2.Rows.Count, 1).End(xlUp).Row
    
    ' Loop through each row in Sheet2
    For i = 2 To lastRowSheet2
        ' Get the billing_code from Sheet2
        Dim billingCode As String
        billingCode = sheet2.Range("A" & i).Value
        
        ' Find the corresponding type_flag in Sheet1
        Dim typeFlag As String
        
        ' Initialize a variable to store the combined type_flags for a billing_code
        Dim combinedTypeFlag As String
        combinedTypeFlag = ""
        
        ' Loop through each row in Sheet1 to find all type_flags for the billing_code
        For j = 2 To lastRowSheet1
            If sheet1.Range("A" & j).Value = billingCode Then
                typeFlag = sheet1.Range("B" & j).Value
                
                ' Check the type_flag against the rules and populate the corresponding column in Sheet2
                Select Case typeFlag
                    Case "1"
                        sheet2.Range("B" & i).Value = typeFlag
                    Case "2"
                        sheet2.Range("C" & i).Value = typeFlag
                    Case "3"
                        sheet2.Range("D" & i).Value = typeFlag
                    Case "4"
                        sheet2.Range("E" & i).Value = typeFlag
                    Case "5"
                        sheet2.Range("F" & i).Value = typeFlag
                    Case "BP"
                        sheet2.Range("G" & i).Value = typeFlag
                    Case "D"
                        sheet2.Range("H" & i).Value = typeFlag
                    Case "L"
                        sheet2.Range("I" & i).Value = typeFlag
                    Case "R"
                        sheet2.Range("J" & i).Value = typeFlag
                    Case "RA"
                        sheet2.Range("K" & i).Value = typeFlag
                    Case "RR"
                        sheet2.Range("L" & i).Value = typeFlag
                    Case "SC"
                        sheet2.Range("M" & i).Value = typeFlag
                    Case "SS"
                        sheet2.Range("N" & i).Value = typeFlag
                    Case "SU"
                        sheet2.Range("O" & i).Value = typeFlag
                    Case "TC"
                        sheet2.Range("P" & i).Value = typeFlag
                    Case "TN"
                        sheet2.Range("Q" & i).Value = typeFlag
                    Case "WA"
                        sheet2.Range("R" & i).Value = typeFlag
                    Case "WB"
                        sheet2.Range("S" & i).Value = typeFlag
                    Case "WD"
                        sheet2.Range("T" & i).Value = typeFlag
                    Case "WG"
                        sheet2.Range("U" & i).Value = typeFlag
                    Case "WM"
                        sheet2.Range("V" & i).Value = typeFlag
                    Case "WR"
                        sheet2.Range("W" & i).Value = typeFlag
                    ' Add more cases for other rules
                End Select
            End If
        Next j
    Next i
End Sub

