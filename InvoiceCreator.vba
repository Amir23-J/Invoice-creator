' ==================================================================================
' AUTO SHOP INVOICE CREATOR - EXCEL VBA VERSION
' ==================================================================================
'
' SETUP INSTRUCTIONS:
' 1. Open Excel and press ALT + F11 to open VBA Editor
' 2. Insert > Module
' 3. Copy and paste this entire code into the module
' 4. Close VBA Editor and return to Excel
' 5. Run the "SetupInvoiceWorkbook" macro once to create the worksheets
' 6. Use the buttons on the Data Entry sheet to manage invoices
'
' FEATURES:
' - Live calculations as you enter data
' - Add/remove line items dynamically
' - Professional invoice layout
' - Export to PDF with industry-standard filename
' - Clear form for new invoices
' ==================================================================================

Option Explicit

' ==================================================================================
' MAIN SETUP - RUN THIS ONCE TO CREATE THE WORKBOOK STRUCTURE
' ==================================================================================
Sub SetupInvoiceWorkbook()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Delete existing sheets if they exist
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Data Entry").Delete
    Sheets("Invoice").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Create worksheets
    Call CreateDataEntrySheet
    Call CreateInvoiceSheet

    ' Activate Data Entry sheet
    Sheets("Data Entry").Activate
    Range("C4").Select

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Invoice workbook setup complete!" & vbCrLf & vbCrLf & _
           "Enter your invoice data in the 'Data Entry' sheet." & vbCrLf & _
           "The 'Invoice' sheet will update automatically.", vbInformation, "Setup Complete"
End Sub

' ==================================================================================
' CREATE DATA ENTRY SHEET
' ==================================================================================
Sub CreateDataEntrySheet()
    Dim ws As Worksheet
    Set ws = Worksheets.Add(Before:=Sheets(1))
    ws.Name = "Data Entry"

    With ws
        ' Set column widths
        .Columns("A:A").ColumnWidth = 2
        .Columns("B:B").ColumnWidth = 20
        .Columns("C:C").ColumnWidth = 30
        .Columns("D:D").ColumnWidth = 15
        .Columns("E:E").ColumnWidth = 15
        .Columns("F:F").ColumnWidth = 15
        .Columns("G:G").ColumnWidth = 15

        ' Header
        .Range("B2:G2").Merge
        .Range("B2").Value = "AUTO SHOP INVOICE BUILDER - DATA ENTRY"
        .Range("B2").Font.Size = 16
        .Range("B2").Font.Bold = True
        .Range("B2").Interior.Color = RGB(102, 126, 234)
        .Range("B2").Font.Color = RGB(255, 255, 255)
        .Range("B2").HorizontalAlignment = xlCenter

        ' Shop Information
        .Range("B4").Value = "SHOP INFORMATION"
        .Range("B4").Font.Bold = True
        .Range("B4").Font.Size = 12
        .Range("B4").Interior.Color = RGB(230, 230, 250)

        .Range("B5").Value = "Shop Name *"
        .Range("B6").Value = "Phone *"
        .Range("B7").Value = "Email"
        .Range("B8").Value = "Address"

        .Range("C5").Name = "ShopName"
        .Range("C6").Name = "ShopPhone"
        .Range("C7").Name = "ShopEmail"
        .Range("C8").Name = "ShopAddress"

        ' Invoice Details
        .Range("E4").Value = "INVOICE DETAILS"
        .Range("E4").Font.Bold = True
        .Range("E4").Font.Size = 12
        .Range("E4").Interior.Color = RGB(230, 230, 250)

        .Range("E5").Value = "Invoice # *"
        .Range("E6").Value = "Date *"

        .Range("F5").Name = "InvoiceNumber"
        .Range("F6").Name = "InvoiceDate"
        .Range("F5").Value = GenerateInvoiceNumber()
        .Range("F6").Value = Date
        .Range("F6").NumberFormat = "mm/dd/yyyy"

        ' Customer Information
        .Range("B10").Value = "CUSTOMER INFORMATION"
        .Range("B10").Font.Bold = True
        .Range("B10").Font.Size = 12
        .Range("B10").Interior.Color = RGB(230, 230, 250)

        .Range("B11").Value = "Customer Name *"
        .Range("B12").Value = "Phone *"
        .Range("B13").Value = "Email"
        .Range("B14").Value = "Address"

        .Range("C11").Name = "CustomerName"
        .Range("C12").Name = "CustomerPhone"
        .Range("C13").Name = "CustomerEmail"
        .Range("C14").Name = "CustomerAddress"

        ' Vehicle Information
        .Range("E10").Value = "VEHICLE INFORMATION"
        .Range("E10").Font.Bold = True
        .Range("E10").Font.Size = 12
        .Range("E10").Interior.Color = RGB(230, 230, 250)

        .Range("E11").Value = "Year *"
        .Range("E12").Value = "Make *"
        .Range("E13").Value = "Model *"
        .Range("E14").Value = "VIN"
        .Range("E15").Value = "License Plate"
        .Range("E16").Value = "Mileage"

        .Range("F11").Name = "VehicleYear"
        .Range("F12").Name = "VehicleMake"
        .Range("F13").Name = "VehicleModel"
        .Range("F14").Name = "VehicleVIN"
        .Range("F15").Name = "VehicleLicense"
        .Range("F16").Name = "VehicleMileage"

        ' Line Items Header
        .Range("B18").Value = "SERVICES & PARTS"
        .Range("B18").Font.Bold = True
        .Range("B18").Font.Size = 12
        .Range("B18").Interior.Color = RGB(230, 230, 250)

        .Range("B19").Value = "Description"
        .Range("D19").Value = "Qty"
        .Range("E19").Value = "Price"
        .Range("F19").Value = "Total"
        .Range("B19:F19").Font.Bold = True
        .Range("B19:F19").Interior.Color = RGB(200, 200, 220)

        ' Initial line items
        Dim i As Integer
        For i = 20 To 21
            .Range("B" & i).Value = ""
            .Range("D" & i).Value = 1
            .Range("E" & i).Value = 0
            .Range("F" & i).Formula = "=IF(B" & i & "<>"""",D" & i & "*E" & i & ","""")"
            .Range("F" & i).NumberFormat = "$#,##0.00"
        Next i

        ' Totals Section
        .Range("E30").Value = "Tax Rate (%)"
        .Range("F30").Name = "TaxRate"
        .Range("F30").Value = 8.5

        .Range("E31").Value = "Subtotal:"
        .Range("F31").Name = "Subtotal"
        .Range("F31").Formula = "=SUM(F20:F29)"
        .Range("F31").NumberFormat = "$#,##0.00"

        .Range("E32").Value = "Tax:"
        .Range("F32").Name = "TaxAmount"
        .Range("F32").Formula = "=Subtotal*(TaxRate/100)"
        .Range("F32").NumberFormat = "$#,##0.00"

        .Range("E33").Value = "TOTAL:"
        .Range("E33").Font.Bold = True
        .Range("E33").Font.Size = 14
        .Range("F33").Name = "GrandTotal"
        .Range("F33").Formula = "=Subtotal+TaxAmount"
        .Range("F33").NumberFormat = "$#,##0.00"
        .Range("F33").Font.Bold = True
        .Range("F33").Font.Size = 14
        .Range("F33").Interior.Color = RGB(102, 126, 234)
        .Range("F33").Font.Color = RGB(255, 255, 255)

        ' Notes
        .Range("B35").Value = "NOTES / TERMS"
        .Range("B35").Font.Bold = True
        .Range("B35").Font.Size = 12
        .Range("B35").Interior.Color = RGB(230, 230, 250)

        .Range("B36:G40").Merge
        .Range("B36").Name = "Notes"
        .Range("B36").WrapText = True
        .Range("B36").VerticalAlignment = xlTop
        With .Range("B36").Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With

        ' Add Buttons
        Call AddButton(ws, "B42", "Add Line Item", "AddLineItem")
        Call AddButton(ws, "C42", "Remove Last Line", "RemoveLineItem")
        Call AddButton(ws, "D42", "Export PDF", "ExportInvoiceToPDF")
        Call AddButton(ws, "E42", "Clear Form", "ClearForm")
        Call AddButton(ws, "F42", "Preview Invoice", "PreviewInvoice")

        ' Format all input cells
        .Range("C5:C8,F5:F6,C11:C14,F11:F16,B20:E29,F30,B36").Interior.Color = RGB(255, 255, 220)

        ' Add borders to sections
        .Range("B4:C8").BorderAround xlContinuous, xlMedium
        .Range("E4:F6").BorderAround xlContinuous, xlMedium
        .Range("B10:C14").BorderAround xlContinuous, xlMedium
        .Range("E10:F16").BorderAround xlContinuous, xlMedium
        .Range("B18:F29").BorderAround xlContinuous, xlMedium
        .Range("E30:F33").BorderAround xlContinuous, xlMedium
        .Range("B35:G40").BorderAround xlContinuous, xlMedium
    End With
End Sub

' ==================================================================================
' CREATE INVOICE SHEET
' ==================================================================================
Sub CreateInvoiceSheet()
    Dim ws As Worksheet
    Set ws = Worksheets.Add(After:=Sheets(Sheets.Count))
    ws.Name = "Invoice"

    With ws
        ' Set column widths
        .Columns("A:A").ColumnWidth = 2
        .Columns("B:B").ColumnWidth = 40
        .Columns("C:C").ColumnWidth = 10
        .Columns("D:D").ColumnWidth = 15
        .Columns("E:E").ColumnWidth = 15

        ' Shop Header
        .Range("B2:C5").Merge
        .Range("B2").Formula = "='Data Entry'!ShopName"
        .Range("B2").Font.Size = 18
        .Range("B2").Font.Bold = True
        .Range("B2").Font.Color = RGB(102, 126, 234)

        .Range("B6").Formula = "='Data Entry'!ShopAddress"
        .Range("B7").Formula = "=""Phone: "" & 'Data Entry'!ShopPhone"
        .Range("B8").Formula = "=IF('Data Entry'!ShopEmail<>"""",""Email: "" & 'Data Entry'!ShopEmail,"""")"

        ' Invoice Details (right side)
        .Range("D2").Value = "INVOICE"
        .Range("D2").Font.Size = 20
        .Range("D2").Font.Bold = True
        .Range("D2:E2").Merge
        .Range("D2").HorizontalAlignment = xlRight

        .Range("D3").Value = "Invoice #:"
        .Range("D3").Font.Bold = True
        .Range("E3").Formula = "='Data Entry'!InvoiceNumber"

        .Range("D4").Value = "Date:"
        .Range("D4").Font.Bold = True
        .Range("E4").Formula = "='Data Entry'!InvoiceDate"
        .Range("E4").NumberFormat = "mmmm dd, yyyy"

        ' Separator line
        .Range("B10:E10").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("B10:E10").Borders(xlEdgeBottom).Weight = xlThick
        .Range("B10:E10").Borders(xlEdgeBottom).Color = RGB(102, 126, 234)

        ' Customer Information
        .Range("B12").Value = "CUSTOMER INFORMATION"
        .Range("B12").Font.Bold = True
        .Range("B12").Font.Color = RGB(102, 126, 234)

        .Range("B13").Formula = "=""Name: "" & 'Data Entry'!CustomerName"
        .Range("B14").Formula = "=""Phone: "" & 'Data Entry'!CustomerPhone"
        .Range("B15").Formula = "=IF('Data Entry'!CustomerEmail<>"""",""Email: "" & 'Data Entry'!CustomerEmail,"""")"
        .Range("B16").Formula = "=IF('Data Entry'!CustomerAddress<>"""",""Address: "" & 'Data Entry'!CustomerAddress,"""")"

        ' Vehicle Information
        .Range("D12").Value = "VEHICLE INFORMATION"
        .Range("D12").Font.Bold = True
        .Range("D12").Font.Color = RGB(102, 126, 234)

        .Range("D13").Formula = "=""Vehicle: "" & 'Data Entry'!VehicleYear & "" "" & 'Data Entry'!VehicleMake & "" "" & 'Data Entry'!VehicleModel"
        .Range("D14").Formula = "=IF('Data Entry'!VehicleVIN<>"""",""VIN: "" & 'Data Entry'!VehicleVIN,"""")"
        .Range("D15").Formula = "=IF('Data Entry'!VehicleLicense<>"""",""License: "" & 'Data Entry'!VehicleLicense,"""")"
        .Range("D16").Formula = "=IF('Data Entry'!VehicleMileage<>"""",""Mileage: "" & 'Data Entry'!VehicleMileage,"""")"

        ' Services & Parts Table
        .Range("B18").Value = "SERVICES & PARTS"
        .Range("B18").Font.Bold = True
        .Range("B18").Font.Color = RGB(102, 126, 234)
        .Range("B18").Font.Size = 12

        ' Table Header
        .Range("B19").Value = "Description"
        .Range("C19").Value = "Qty"
        .Range("D19").Value = "Price"
        .Range("E19").Value = "Total"
        .Range("B19:E19").Font.Bold = True
        .Range("B19:E19").Interior.Color = RGB(102, 126, 234)
        .Range("B19:E19").Font.Color = RGB(255, 255, 255)
        .Range("B19:E19").HorizontalAlignment = xlCenter

        ' Line items with formulas
        Dim i As Integer
        For i = 20 To 29
            .Range("B" & i).Formula = "='Data Entry'!B" & i
            .Range("C" & i).Formula = "=IF('Data Entry'!B" & i & "<>"""",'Data Entry'!D" & i & ","""")"
            .Range("D" & i).Formula = "=IF('Data Entry'!B" & i & "<>"""",'Data Entry'!E" & i & ","""")"
            .Range("E" & i).Formula = "=IF('Data Entry'!B" & i & "<>"""",'Data Entry'!F" & i & ","""")"

            .Range("C" & i).HorizontalAlignment = xlCenter
            .Range("D" & i).NumberFormat = "$#,##0.00"
            .Range("E" & i).NumberFormat = "$#,##0.00"
            .Range("D" & i).HorizontalAlignment = xlRight
            .Range("E" & i).HorizontalAlignment = xlRight
        Next i

        ' Table border
        .Range("B19:E29").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("B19:E29").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("B19:E29").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("B19:E29").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("B19:E29").Borders(xlInsideVertical).LineStyle = xlContinuous
        .Range("B19:E29").Borders(xlInsideHorizontal).LineStyle = xlDot

        ' Totals
        .Range("D31").Value = "Subtotal:"
        .Range("E31").Formula = "='Data Entry'!Subtotal"
        .Range("E31").NumberFormat = "$#,##0.00"

        .Range("D32").Formula = "=""Tax ("" & 'Data Entry'!TaxRate & "%):"""
        .Range("E32").Formula = "='Data Entry'!TaxAmount"
        .Range("E32").NumberFormat = "$#,##0.00"

        .Range("D33").Value = "TOTAL:"
        .Range("D33").Font.Bold = True
        .Range("D33").Font.Size = 14
        .Range("E33").Formula = "='Data Entry'!GrandTotal"
        .Range("E33").NumberFormat = "$#,##0.00"
        .Range("E33").Font.Bold = True
        .Range("E33").Font.Size = 14
        .Range("D33:E33").Interior.Color = RGB(102, 126, 234)
        .Range("D33:E33").Font.Color = RGB(255, 255, 255)

        ' Notes
        .Range("B35").Value = "NOTES / TERMS"
        .Range("B35").Font.Bold = True
        .Range("B35").Font.Color = RGB(102, 126, 234)

        .Range("B36:E40").Merge
        .Range("B36").Formula = "='Data Entry'!Notes"
        .Range("B36").WrapText = True
        .Range("B36").VerticalAlignment = xlTop
        .Range("B36").Interior.Color = RGB(248, 249, 250)
        With .Range("B36:E40").Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With

        ' Hide gridlines
        ActiveWindow.DisplayGridlines = False
    End With
End Sub

' ==================================================================================
' HELPER FUNCTION TO ADD BUTTONS
' ==================================================================================
Sub AddButton(ws As Worksheet, cellAddress As String, buttonText As String, macroName As String)
    Dim btn As Button
    Dim rng As Range
    Set rng = ws.Range(cellAddress)

    Set btn = ws.Buttons.Add(rng.Left, rng.Top, rng.Width, rng.Height)
    With btn
        .Caption = buttonText
        .OnAction = macroName
    End With
End Sub

' ==================================================================================
' GENERATE INVOICE NUMBER
' ==================================================================================
Function GenerateInvoiceNumber() As String
    GenerateInvoiceNumber = "INV-" & Year(Date) & "-" & Format(Int(Rnd() * 10000), "0000")
End Function

' ==================================================================================
' ADD LINE ITEM
' ==================================================================================
Sub AddLineItem()
    Dim ws As Worksheet
    Set ws = Sheets("Data Entry")

    Dim lastRow As Long
    lastRow = 20

    ' Find the last used row
    Do While ws.Range("B" & lastRow).Value <> "" And lastRow < 29
        lastRow = lastRow + 1
    Loop

    If lastRow >= 29 Then
        MsgBox "Maximum number of line items (10) reached.", vbExclamation, "Limit Reached"
        Exit Sub
    End If

    ' Add new line item
    ws.Range("B" & lastRow + 1).Value = ""
    ws.Range("D" & lastRow + 1).Value = 1
    ws.Range("E" & lastRow + 1).Value = 0
    ws.Range("F" & lastRow + 1).Formula = "=IF(B" & lastRow + 1 & "<>"""",D" & lastRow + 1 & "*E" & lastRow + 1 & ","""")"
    ws.Range("F" & lastRow + 1).NumberFormat = "$#,##0.00"

    ' Select the description cell
    ws.Range("B" & lastRow + 1).Select
End Sub

' ==================================================================================
' REMOVE LINE ITEM
' ==================================================================================
Sub RemoveLineItem()
    Dim ws As Worksheet
    Set ws = Sheets("Data Entry")

    Dim lastRow As Long
    lastRow = 29

    ' Find the last used row
    Do While (ws.Range("B" & lastRow).Value = "" Or IsEmpty(ws.Range("B" & lastRow))) And lastRow > 20
        lastRow = lastRow - 1
    Loop

    If lastRow < 20 Then
        MsgBox "No line items to remove.", vbExclamation, "Nothing to Remove"
        Exit Sub
    End If

    ' Clear the last line item
    ws.Range("B" & lastRow & ":F" & lastRow).ClearContents
End Sub

' ==================================================================================
' PREVIEW INVOICE
' ==================================================================================
Sub PreviewInvoice()
    Sheets("Invoice").Activate
    Range("B2").Select
End Sub

' ==================================================================================
' VALIDATE FORM
' ==================================================================================
Function ValidateForm() As Boolean
    Dim ws As Worksheet
    Set ws = Sheets("Data Entry")

    ValidateForm = False

    ' Check required fields
    If Trim(ws.Range("ShopName").Value) = "" Then
        MsgBox "Please enter Shop Name.", vbExclamation, "Required Field"
        ws.Range("ShopName").Select
        Exit Function
    End If

    If Trim(ws.Range("ShopPhone").Value) = "" Then
        MsgBox "Please enter Shop Phone.", vbExclamation, "Required Field"
        ws.Range("ShopPhone").Select
        Exit Function
    End If

    If Trim(ws.Range("InvoiceNumber").Value) = "" Then
        MsgBox "Please enter Invoice Number.", vbExclamation, "Required Field"
        ws.Range("InvoiceNumber").Select
        Exit Function
    End If

    If Trim(ws.Range("CustomerName").Value) = "" Then
        MsgBox "Please enter Customer Name.", vbExclamation, "Required Field"
        ws.Range("CustomerName").Select
        Exit Function
    End If

    If Trim(ws.Range("CustomerPhone").Value) = "" Then
        MsgBox "Please enter Customer Phone.", vbExclamation, "Required Field"
        ws.Range("CustomerPhone").Select
        Exit Function
    End If

    If Trim(ws.Range("VehicleYear").Value) = "" Then
        MsgBox "Please enter Vehicle Year.", vbExclamation, "Required Field"
        ws.Range("VehicleYear").Select
        Exit Function
    End If

    If Trim(ws.Range("VehicleMake").Value) = "" Then
        MsgBox "Please enter Vehicle Make.", vbExclamation, "Required Field"
        ws.Range("VehicleMake").Select
        Exit Function
    End If

    If Trim(ws.Range("VehicleModel").Value) = "" Then
        MsgBox "Please enter Vehicle Model.", vbExclamation, "Required Field"
        ws.Range("VehicleModel").Select
        Exit Function
    End If

    ' Check for at least one line item
    Dim hasLineItem As Boolean
    hasLineItem = False
    Dim i As Integer
    For i = 20 To 29
        If Trim(ws.Range("B" & i).Value) <> "" And ws.Range("E" & i).Value > 0 Then
            hasLineItem = True
            Exit For
        End If
    Next i

    If Not hasLineItem Then
        MsgBox "Please add at least one service or part with a description and price.", vbExclamation, "No Line Items"
        ws.Range("B20").Select
        Exit Function
    End If

    ValidateForm = True
End Function

' ==================================================================================
' EXPORT TO PDF
' ==================================================================================
Sub ExportInvoiceToPDF()
    ' Validate form first
    If Not ValidateForm() Then
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = Sheets("Invoice")

    ' Generate filename
    Dim fileName As String
    Dim invoiceNum As String
    Dim customerName As String

    invoiceNum = Sheets("Data Entry").Range("InvoiceNumber").Value
    customerName = Sheets("Data Entry").Range("CustomerName").Value

    ' Sanitize filename
    invoiceNum = Replace(invoiceNum, "/", "-")
    invoiceNum = Replace(invoiceNum, "\", "-")
    invoiceNum = Replace(invoiceNum, ":", "-")
    invoiceNum = Replace(invoiceNum, "*", "-")
    invoiceNum = Replace(invoiceNum, "?", "-")
    invoiceNum = Replace(invoiceNum, """", "-")
    invoiceNum = Replace(invoiceNum, "<", "-")
    invoiceNum = Replace(invoiceNum, ">", "-")
    invoiceNum = Replace(invoiceNum, "|", "-")

    customerName = Replace(customerName, " ", "_")
    customerName = Replace(customerName, "/", "-")
    customerName = Replace(customerName, "\", "-")
    customerName = Replace(customerName, ":", "-")
    customerName = Replace(customerName, "*", "-")
    customerName = Replace(customerName, "?", "-")
    customerName = Replace(customerName, """", "-")
    customerName = Replace(customerName, "<", "-")
    customerName = Replace(customerName, ">", "-")
    customerName = Replace(customerName, "|", "-")

    If Len(customerName) > 30 Then
        customerName = Left(customerName, 30)
    End If

    fileName = invoiceNum & "_" & customerName & ".pdf"

    ' Get save location
    Dim savePath As String
    savePath = Application.GetSaveAsFilename(InitialFileName:=fileName, _
                                             FileFilter:="PDF Files (*.pdf), *.pdf", _
                                             Title:="Save Invoice as PDF")

    If savePath = "False" Then Exit Sub ' User cancelled

    ' Export to PDF
    On Error GoTo ErrorHandler
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
                           fileName:=savePath, _
                           Quality:=xlQualityStandard, _
                           IncludeDocProperties:=True, _
                           IgnorePrintAreas:=False, _
                           OpenAfterPublish:=True

    MsgBox "Invoice PDF exported successfully!" & vbCrLf & vbCrLf & savePath, vbInformation, "Export Complete"
    Exit Sub

ErrorHandler:
    MsgBox "Error exporting PDF: " & Err.Description, vbCritical, "Export Error"
End Sub

' ==================================================================================
' CLEAR FORM
' ==================================================================================
Sub ClearForm()
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to clear all fields? This cannot be undone.", _
                     vbYesNo + vbQuestion, "Clear Form")

    If response = vbNo Then Exit Sub

    Dim ws As Worksheet
    Set ws = Sheets("Data Entry")

    Application.ScreenUpdating = False

    ' Clear all fields except tax rate
    ws.Range("ShopName").ClearContents
    ws.Range("ShopPhone").ClearContents
    ws.Range("ShopEmail").ClearContents
    ws.Range("ShopAddress").ClearContents

    ws.Range("CustomerName").ClearContents
    ws.Range("CustomerPhone").ClearContents
    ws.Range("CustomerEmail").ClearContents
    ws.Range("CustomerAddress").ClearContents

    ws.Range("VehicleYear").ClearContents
    ws.Range("VehicleMake").ClearContents
    ws.Range("VehicleModel").ClearContents
    ws.Range("VehicleVIN").ClearContents
    ws.Range("VehicleLicense").ClearContents
    ws.Range("VehicleMileage").ClearContents

    ws.Range("Notes").ClearContents

    ' Clear line items
    ws.Range("B20:F29").ClearContents

    ' Reset first two line items
    ws.Range("D20:D21").Value = 1
    ws.Range("E20:E21").Value = 0
    ws.Range("F20").Formula = "=IF(B20<>"""",D20*E20,"""")"
    ws.Range("F21").Formula = "=IF(B21<>"""",D21*E21,"""")"
    ws.Range("F20:F21").NumberFormat = "$#,##0.00"

    ' Generate new invoice number and date
    ws.Range("InvoiceNumber").Value = GenerateInvoiceNumber()
    ws.Range("InvoiceDate").Value = Date

    ' Return to top
    ws.Range("C5").Select

    Application.ScreenUpdating = True

    MsgBox "Form cleared successfully!", vbInformation, "Clear Complete"
End Sub
