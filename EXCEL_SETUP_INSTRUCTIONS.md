# Excel Invoice Creator - Setup Instructions

## Overview
This VBA code converts the web-based invoice creator into a fully functional Excel workbook with:
- **Data Entry Sheet**: Easy-to-use form for entering invoice information
- **Invoice Sheet**: Professional, formatted invoice that updates automatically
- **PDF Export**: One-click export with industry-standard filename format
- **Dynamic Line Items**: Add/remove services and parts as needed
- **Live Calculations**: Totals update automatically as you type

---

## Installation Instructions

### Step 1: Open Excel
1. Open Microsoft Excel
2. Create a new blank workbook

### Step 2: Open VBA Editor
1. Press **ALT + F11** (Windows) or **Option + F11** (Mac)
2. This opens the Visual Basic for Applications (VBA) Editor

### Step 3: Insert a New Module
1. In the VBA Editor, click **Insert** → **Module**
2. A new code window will appear

### Step 4: Copy the VBA Code
1. Open the file `InvoiceCreator.vba`
2. Select all the code (CTRL + A or CMD + A)
3. Copy it (CTRL + C or CMD + C)

### Step 5: Paste the Code
1. Return to the VBA Editor
2. Click inside the blank module window
3. Paste the code (CTRL + V or CMD + V)

### Step 6: Close VBA Editor
1. Click the X to close the VBA Editor
2. Return to Excel

### Step 7: Run the Setup Macro
1. Press **ALT + F8** (Windows) or **Option + F8** (Mac) to open the Macro dialog
2. Select **SetupInvoiceWorkbook** from the list
3. Click **Run**

**That's it!** Your invoice workbook is now ready to use.

---

## How to Use

### Data Entry Sheet
The "Data Entry" sheet is where you enter all your invoice information:

#### 1. **Shop Information** (Left Column, Top)
- Shop Name *
- Phone *
- Email
- Address

#### 2. **Invoice Details** (Right Column, Top)
- Invoice # (auto-generated)
- Date (defaults to today)

#### 3. **Customer Information** (Left Column, Middle)
- Customer Name *
- Phone *
- Email
- Address

#### 4. **Vehicle Information** (Right Column, Middle)
- Year *
- Make *
- Model *
- VIN
- License Plate
- Mileage

#### 5. **Services & Parts** (Bottom Section)
- Enter description, quantity, and price for each line item
- Total calculates automatically
- Use buttons to add/remove line items

#### 6. **Totals**
- Tax Rate: Adjust as needed (default 8.5%)
- Subtotal, Tax, and Grand Total calculate automatically

#### 7. **Notes/Terms**
- Add payment terms, warranty info, or any additional notes

**Required fields are marked with * asterisk**

---

## Using the Buttons

### 📝 Add Line Item
- Adds a new row for services/parts
- Maximum of 10 line items
- Cursor automatically moves to the new description field

### ❌ Remove Last Line
- Removes the last entered line item
- Useful for correcting mistakes

### 👁️ Preview Invoice
- Switches to the formatted Invoice sheet
- Shows exactly what will be exported to PDF
- All data updates automatically from Data Entry sheet

### 📄 Export PDF
- Validates all required fields are filled
- Opens save dialog with auto-generated filename format: `INV-2025-1234_John_Doe.pdf`
- Saves and opens the PDF automatically
- Professional format ready for printing or emailing

### 🗑️ Clear Form
- Clears all fields to start a new invoice
- Asks for confirmation before clearing
- Generates new invoice number and resets date
- Keeps tax rate setting

---

## Automatic Features

### ✅ Live Calculations
- Line item totals update as you type
- Subtotal recalculates automatically
- Tax amount adjusts with tax rate changes
- Grand total always accurate

### ✅ Professional Formatting
- Color-coded sections for easy navigation
- Yellow-highlighted input fields
- Clean, organized layout
- Print-ready invoice sheet

### ✅ Smart Filename Generation
- Format: `InvoiceNumber_CustomerName.pdf`
- Example: `INV-2025-1234_John_Doe.pdf`
- Removes special characters automatically
- Perfect for file organization

---

## Tips & Best Practices

### 💡 Workflow Recommendation
1. Fill out Shop Information once (it will stay for all invoices)
2. Enter Customer and Vehicle information
3. Add line items for services/parts
4. Adjust tax rate if needed
5. Add notes/terms
6. Click "Preview Invoice" to review
7. Click "Export PDF" to save

### 💡 Saving Your Workbook
1. Save as Excel Macro-Enabled Workbook (.xlsm)
2. File → Save As → Select "Excel Macro-Enabled Workbook (*.xlsm)"
3. This preserves all your VBA code

### 💡 Invoice Numbering
- Invoice numbers auto-generate with format: INV-YYYY-####
- You can manually edit them if needed
- Consider using sequential numbers for better tracking

### 💡 Tax Rates
- Default is 8.5%
- Change in cell F30 on Data Entry sheet
- Tax rate persists even after clearing form
- Can be changed for individual invoices as needed

---

## Troubleshooting

### ❓ Macros are Disabled
**Problem**: Buttons don't work when clicking them

**Solution**:
1. Go to File → Options → Trust Center → Trust Center Settings
2. Click "Macro Settings"
3. Select "Enable all macros" (or "Disable with notification")
4. Click OK and restart Excel

### ❓ PDF Export Not Working
**Problem**: Error when exporting to PDF

**Solution**:
- Ensure you have Excel 2007 or later (PDF export feature required)
- Check that you filled all required fields (marked with *)
- Try saving to a different location
- Ensure you have write permissions to the save location

### ❓ Invoice Sheet Not Updating
**Problem**: Changes in Data Entry don't show in Invoice

**Solution**:
- Press F9 to force recalculation
- Check that formulas weren't accidentally deleted
- Re-run the SetupInvoiceWorkbook macro to rebuild

### ❓ Lost a Line Item
**Problem**: Accidentally cleared a line item

**Solution**:
- Just re-type the information
- Or click "Add Line Item" to add a new one
- Line items can be in any row between 20-29

---

## Customization

### Changing Colors
To modify the color scheme:
1. Open VBA Editor (ALT + F11)
2. Find the CreateDataEntrySheet or CreateInvoiceSheet sub
3. Look for `RGB(102, 126, 234)` - this is the main purple color
4. Change RGB values (Red, Green, Blue) to your preference
5. Re-run SetupInvoiceWorkbook macro

### Changing Tax Rate Default
1. Open VBA Editor (ALT + F11)
2. Find line: `.Range("F30").Value = 8.5`
3. Change 8.5 to your default tax rate
4. Re-run SetupInvoiceWorkbook macro

### Adding More Line Items
To increase from 10 to more line items:
1. Open VBA Editor (ALT + F11)
2. Find all instances of "29" (last row number)
3. Change to your desired last row (e.g., 39 for 20 line items)
4. Re-run SetupInvoiceWorkbook macro

---

## Features Comparison with Web Version

| Feature | Web Version | Excel VBA Version |
|---------|-------------|-------------------|
| Live Preview | ✅ Side-by-side | ✅ Separate sheet |
| Editable After Creation | ✅ Always | ✅ Always |
| PDF Export | ✅ Yes | ✅ Yes |
| Industry-Standard Filename | ✅ Yes | ✅ Yes |
| Add/Remove Line Items | ✅ Yes | ✅ Yes |
| Auto Calculations | ✅ Yes | ✅ Yes |
| Clear Form | ✅ Yes | ✅ Yes |
| Save for Later | ❌ No | ✅ Save as .xlsm |
| Works Offline | ❌ Browser only | ✅ Yes |
| Print Directly | ❌ Print PDF | ✅ Yes |

---

## Support

For issues or questions:
1. Check the Troubleshooting section above
2. Review the code comments in InvoiceCreator.vba
3. Ensure you're using Excel 2007 or later

---

## License

This code is provided as-is for use in your auto shop business. Feel free to customize it to your needs.

---

**Enjoy your new Excel Invoice Creator!** 🚗🔧
