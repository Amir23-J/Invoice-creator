# 🔧 Auto Shop Invoice Builder

A simple, user-friendly web application for auto shops to create and export professional invoices. Built with vanilla HTML, CSS, and JavaScript - no frameworks required!

## Features

- ✨ **Easy-to-Use Interface** - Clean, intuitive design that anyone can use
- 🚗 **Auto Shop Specific** - Fields for vehicle information (make, model, year, VIN, mileage, etc.)
- 📝 **Dynamic Line Items** - Add unlimited services and parts
- 🧮 **Automatic Calculations** - Real-time calculation of subtotals, tax, and totals
- 👁️ **Invoice Preview** - Review the invoice before exporting
- 📄 **PDF Export** - Download professional invoices as PDF files
- 📱 **Responsive Design** - Works on desktop, tablet, and mobile devices
- 💾 **No Database Required** - Everything runs in the browser

## How to Use

### Getting Started

1. Simply open `index.html` in any modern web browser
2. No installation or setup required!

### Creating an Invoice

#### 1. Shop Information
Fill in your auto shop details:
- Shop Name (required)
- Phone Number (required)
- Email (optional)
- Address (optional)

#### 2. Invoice Details
- Invoice Number (auto-generated, but you can change it)
- Date (defaults to today)

#### 3. Customer Information
Enter the customer's details:
- Name (required)
- Phone (required)
- Email (optional)
- Address (optional)

#### 4. Vehicle Information
Add details about the customer's vehicle:
- Year (required)
- Make (required)
- Model (required)
- VIN (optional)
- License Plate (optional)
- Mileage (optional)

#### 5. Services & Parts
- Click "+ Add Line Item" to add services or parts
- Enter description, quantity, and price
- The total is calculated automatically
- Click the "×" button to remove a line item
- Add as many line items as needed

#### 6. Totals
- Adjust the tax rate if needed (default is 8.5%)
- Subtotal, tax, and total are calculated automatically

#### 7. Notes / Terms
Add any additional information:
- Payment terms
- Warranty information
- Special instructions
- Thank you message

#### 8. Complete & Export
1. Click "✓ Complete Invoice" to preview your invoice
2. Review the invoice preview
3. Click "📄 Export as PDF" to download
4. Use "Edit Invoice" if you need to make changes

### Additional Features

- **Clear Form** - Start over with a fresh invoice (generates a new invoice number)
- **Edit Invoice** - Go back and modify the invoice after completing it

## Technical Details

### Technologies Used

- **HTML5** - Structure and content
- **CSS3** - Styling with modern features (Grid, Flexbox, gradients)
- **Vanilla JavaScript** - All functionality without frameworks
- **jsPDF** - PDF generation library
- **html2canvas** - HTML to canvas conversion for PDF export

### Browser Compatibility

Works in all modern browsers:
- Chrome/Edge (recommended)
- Firefox
- Safari
- Opera

### File Structure

```
Invoice-creator/
├── index.html      # Main HTML file
├── styles.css      # All styling
├── script.js       # Application logic
└── README.md       # Documentation
```

## Features Breakdown

### User-Friendly Design
- Purple gradient background for a modern look
- Clear section headers with color coding
- Large, easy-to-click buttons
- Helpful placeholder text in all fields
- Required fields marked with asterisks

### Smart Calculations
- Line item totals update as you type
- Subtotal calculates from all line items
- Tax calculated based on customizable rate
- Total updates automatically

### Professional Invoice Preview
- Clean, printable layout
- All information organized clearly
- Professional typography and spacing
- Color-coded sections for easy reading

### PDF Export
- High-quality PDF generation
- Preserves all formatting and styling
- Automatic filename based on invoice number
- One-click download

## Tips for Best Results

1. **Fill Required Fields** - The app will remind you if you miss required fields
2. **Double-Check Numbers** - Review quantities and prices before completing
3. **Save Your Work** - Export to PDF to keep a record
4. **Customize Tax Rate** - Adjust to match your local tax rate
5. **Use Notes Section** - Add payment terms, warranties, or thank you messages

## Future Enhancement Ideas

- Save invoices to browser storage
- Customer database
- Auto-fill returning customers
- Email invoice directly
- Print invoice
- Multiple currency support
- Invoice templates
- Custom branding/logo upload

## Support

This is a standalone web application that runs entirely in your browser. No server or internet connection required after the initial page load.

## License

Free to use for any auto shop or business.

---

Made with ❤️ for auto shops everywhere
