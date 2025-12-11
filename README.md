# 🔧 Auto Shop Invoice Creator

A professional invoice creation system for auto shops, available in **two versions**: Web and Excel.

---

## 📦 Two Versions Available

### 🌐 Web Version
A modern, browser-based invoice creator with live preview.

**Files:**
- `index.html` - Main application
- `script.js` - Application logic
- `styles.css` - Styling

**Features:**
- ✅ Live preview that updates as you type
- ✅ Side-by-side editing and preview (on desktop)
- ✅ Always editable - no mode switching needed
- ✅ Add/remove line items dynamically
- ✅ Export to PDF with smart filename
- ✅ Works in any modern browser
- ✅ No installation required

**How to Use:**
1. Open `index.html` in any web browser
2. Fill out the form - preview updates automatically
3. Click "Export as PDF" when ready
4. PDF saves as: `InvoiceNumber_CustomerName.pdf`

---

### 📊 Excel VBA Version
A fully-featured Excel workbook with automated invoice generation.

**Files:**
- `InvoiceCreator.vba` - Complete VBA code
- `EXCEL_SETUP_INSTRUCTIONS.md` - Detailed setup guide
- `QUICK_START.txt` - Quick reference

**Features:**
- ✅ Two-sheet design (Data Entry + Invoice)
- ✅ All calculations automatic
- ✅ Add/remove line items with buttons
- ✅ Export to PDF with smart filename
- ✅ Save invoices for later editing
- ✅ Works completely offline
- ✅ Print directly from Excel
- ✅ Professional formatted invoice

**How to Use:**
1. Open Excel, press ALT+F11
2. Insert → Module, paste code from `InvoiceCreator.vba`
3. Run "SetupInvoiceWorkbook" macro
4. Fill out Data Entry sheet
5. Export PDF or print directly

**Quick Setup:**
```
ALT+F11 → Insert Module → Paste Code → ALT+F8 → Run SetupInvoiceWorkbook
```

See `EXCEL_SETUP_INSTRUCTIONS.md` for full details.

---

## 🎯 Common Features (Both Versions)

### Invoice Information Captured:
- **Shop Details**: Name, phone, email, address
- **Customer Info**: Name, phone, email, address
- **Vehicle Info**: Year, make, model, VIN, license, mileage
- **Line Items**: Description, quantity, price (up to 10 items)
- **Calculations**: Subtotal, tax (adjustable %), grand total
- **Notes**: Terms, warranty, payment info

### PDF Export:
Both versions create professional PDFs with industry-standard filenames:
```
Format: InvoiceNumber_CustomerName.pdf
Example: INV-2025-1234_John_Doe.pdf
```

---

## 🆚 Which Version Should You Use?

| Feature | Web Version | Excel Version |
|---------|-------------|---------------|
| **Best For** | Quick invoices on any device | Office environment with Excel |
| **Installation** | None - open and use | One-time VBA setup |
| **Portability** | Share HTML file or host online | Share .xlsm workbook file |
| **Editing** | Always editable, live preview | Always editable, separate sheets |
| **Save Progress** | Must print/export to save | Save workbook anytime |
| **Offline Use** | Yes, if file is local | Yes |
| **Calculations** | Real-time JavaScript | Real-time Excel formulas |
| **Customization** | Edit HTML/CSS/JS | Edit VBA code |
| **Database/History** | Not included | Save multiple in workbook |

---

## 🚀 Getting Started

### For Web Version:
```bash
# Just open the file
open index.html
# or double-click index.html in your file browser
```

### For Excel Version:
1. Read `QUICK_START.txt` for ultra-fast setup
2. Read `EXCEL_SETUP_INSTRUCTIONS.md` for detailed guide
3. Open `InvoiceCreator.vba` to see the code

---

## 📝 Invoice Workflow

Both versions follow the same logical workflow:

1. **Enter Shop Information** (one-time setup, reusable)
2. **Enter Customer Details** (name, contact, etc.)
3. **Enter Vehicle Information** (make, model, year, etc.)
4. **Add Services & Parts** (description, qty, price)
5. **Adjust Tax Rate** (if needed)
6. **Add Notes/Terms** (payment terms, warranties)
7. **Preview** (web: automatic, Excel: click Preview button)
8. **Export PDF** (saves with smart filename)
9. **Clear for Next Invoice** (both versions have clear function)

---

## 💡 Key Features Explained

### Always Editable
Unlike traditional invoice systems that "lock" after creation, both versions let you edit everything at any time. Made a typo? Just fix it. Forgot a line item? Add it. The invoice updates instantly.

### Smart PDF Filenames
Instead of generic "Invoice.pdf" names, both versions create descriptive filenames:
- Includes invoice number for easy sorting
- Includes customer name for quick identification
- Removes special characters automatically
- Professional format used by accounting systems

Example: `INV-2025-1234_ABC_Auto_Repair.pdf`

### Dynamic Line Items
Start with 2 line items, add up to 10 as needed:
- Web: Click "Add Line Item" button
- Excel: Click "Add Line Item" button or manually add rows

Remove line items you don't need:
- Web: Click X button on each line
- Excel: Click "Remove Last Line" button

### Automatic Calculations
All math is handled automatically:
- Line totals = Quantity × Price
- Subtotal = Sum of all line totals
- Tax = Subtotal × Tax Rate %
- Grand Total = Subtotal + Tax

Change any number and totals update instantly.

---

## 🎨 Customization

### Web Version (HTML/CSS/JS)
- **Colors**: Edit `styles.css` - search for `#667eea` (purple theme)
- **Fields**: Edit `index.html` to add/remove fields
- **Calculations**: Edit `script.js` functions
- **Layout**: Modify CSS grid and flexbox in `styles.css`

### Excel Version (VBA)
- **Colors**: Edit RGB values in CreateDataEntrySheet sub
- **Tax Rate Default**: Change `.Range("F30").Value = 8.5`
- **More Line Items**: Increase row 29 to higher number (e.g., 39)
- **Company Logo**: Add image to Invoice sheet manually

---

## 📋 Requirements

### Web Version:
- Any modern web browser (Chrome, Firefox, Safari, Edge)
- JavaScript enabled
- For PDF export: Internet connection (uses CDN libraries)

### Excel Version:
- Microsoft Excel 2007 or later (Windows or Mac)
- Macros enabled
- PDF export feature (included in Excel 2007+)

---

## 🔧 Troubleshooting

### Web Version:
**PDF not downloading?**
- Check browser popup blocker
- Ensure JavaScript is enabled
- Try a different browser

**Preview not updating?**
- Refresh the page
- Check browser console for errors (F12)

### Excel Version:
**Buttons not working?**
- Enable macros: File → Options → Trust Center → Macro Settings
- Re-run SetupInvoiceWorkbook macro

**Invoice sheet blank?**
- Press F9 to force recalculation
- Check formulas weren't accidentally deleted

See `EXCEL_SETUP_INSTRUCTIONS.md` for more troubleshooting.

---

## 📄 Files in This Repository

```
Invoice-creator/
├── index.html                      # Web version - main file
├── script.js                       # Web version - logic
├── styles.css                      # Web version - styling
├── InvoiceCreator.vba             # Excel version - VBA code
├── EXCEL_SETUP_INSTRUCTIONS.md    # Excel version - full guide
├── QUICK_START.txt                # Excel version - quick reference
└── README.md                      # This file
```

---

## 🎓 Tips for Best Results

1. **Fill shop info once** - It stays for all invoices (especially in Excel)
2. **Use consistent invoice numbering** - Both versions auto-generate, but you can edit
3. **Be specific in descriptions** - Helps customers understand charges
4. **Include VIN when possible** - Professional touch for record-keeping
5. **Add payment terms in notes** - Clarify expectations upfront
6. **Save your work** - Excel version: save as .xlsm; Web version: export PDF
7. **Keep tax rate current** - Update when rates change in your area

---

## 🔄 Recent Updates

**Latest Changes:**
- ✅ Made everything always editable (removed "complete invoice" workflow)
- ✅ Added live preview that updates as you type (web version)
- ✅ Implemented industry-standard PDF filename format
- ✅ Created full Excel VBA version with same features
- ✅ Side-by-side layout on desktop (web version)
- ✅ Professional formatting and color scheme

---

## 📞 Support

Choose the right documentation for your version:

- **Web Version**: Review `index.html`, `script.js`, `styles.css`
- **Excel Version**: See `EXCEL_SETUP_INSTRUCTIONS.md` and `QUICK_START.txt`
- **General Questions**: See this README

---

## 📜 License

This project is provided as-is for use in auto shop businesses. Feel free to customize it to your specific needs.

---

## 🙏 Acknowledgments

Built with:
- **Web Version**: HTML5, CSS3, JavaScript, jsPDF, html2canvas
- **Excel Version**: Excel VBA, native Excel features

---

**Made with ❤️ for auto shop owners who need simple, professional invoicing** 🚗🔧
