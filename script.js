// Invoice Builder Application
let lineItemCounter = 0;

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
});

function initializeApp() {
    // Set default values
    document.getElementById('invoiceDate').valueAsDate = new Date();
    document.getElementById('invoiceNumber').value = generateInvoiceNumber();

    // Add initial line items
    addLineItem();
    addLineItem();

    // Event listeners
    document.getElementById('addLineItemBtn').addEventListener('click', addLineItem);
    document.getElementById('completeBtn').addEventListener('click', completeInvoice);
    document.getElementById('exportPdfBtn').addEventListener('click', exportToPDF);
    document.getElementById('clearBtn').addEventListener('click', clearForm);
    document.getElementById('editBtn').addEventListener('click', editInvoice);
    document.getElementById('taxRate').addEventListener('input', calculateTotals);
}

// Generate a random invoice number
function generateInvoiceNumber() {
    const date = new Date();
    const year = date.getFullYear();
    const random = Math.floor(Math.random() * 10000).toString().padStart(4, '0');
    return `INV-${year}-${random}`;
}

// Add a new line item
function addLineItem() {
    lineItemCounter++;
    const container = document.getElementById('lineItemsContainer');

    const lineItem = document.createElement('div');
    lineItem.className = 'line-item';
    lineItem.id = `lineItem${lineItemCounter}`;

    lineItem.innerHTML = `
        <input type="text" class="line-description" placeholder="Description of service or part" />
        <input type="number" class="line-quantity" placeholder="Qty" min="1" value="1" />
        <input type="number" class="line-price" placeholder="Price" min="0" step="0.01" />
        <div class="line-item-total">$0.00</div>
        <button type="button" class="remove-line-item" onclick="removeLineItem(${lineItemCounter})">×</button>
    `;

    container.appendChild(lineItem);

    // Add event listeners for calculation
    const inputs = lineItem.querySelectorAll('input');
    inputs.forEach(input => {
        input.addEventListener('input', function() {
            updateLineItemTotal(lineItemCounter);
            calculateTotals();
        });
    });
}

// Remove a line item
function removeLineItem(id) {
    const lineItem = document.getElementById(`lineItem${id}`);
    if (lineItem) {
        lineItem.remove();
        calculateTotals();
    }
}

// Update individual line item total
function updateLineItemTotal(id) {
    const lineItem = document.getElementById(`lineItem${id}`);
    if (!lineItem) return;

    const quantity = parseFloat(lineItem.querySelector('.line-quantity').value) || 0;
    const price = parseFloat(lineItem.querySelector('.line-price').value) || 0;
    const total = quantity * price;

    lineItem.querySelector('.line-item-total').textContent = formatCurrency(total);
}

// Calculate all totals
function calculateTotals() {
    let subtotal = 0;

    const lineItems = document.querySelectorAll('.line-item');
    lineItems.forEach(item => {
        const quantity = parseFloat(item.querySelector('.line-quantity').value) || 0;
        const price = parseFloat(item.querySelector('.line-price').value) || 0;
        subtotal += quantity * price;
    });

    const taxRate = parseFloat(document.getElementById('taxRate').value) || 0;
    const tax = subtotal * (taxRate / 100);
    const total = subtotal + tax;

    document.getElementById('subtotal').textContent = formatCurrency(subtotal);
    document.getElementById('tax').textContent = formatCurrency(tax);
    document.getElementById('total').textContent = formatCurrency(total);
}

// Format number as currency
function formatCurrency(amount) {
    return '$' + amount.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
}

// Validate form
function validateForm() {
    const requiredFields = [
        'shopName',
        'shopPhone',
        'invoiceNumber',
        'invoiceDate',
        'customerName',
        'customerPhone',
        'vehicleYear',
        'vehicleMake',
        'vehicleModel'
    ];

    for (let field of requiredFields) {
        const element = document.getElementById(field);
        if (!element.value.trim()) {
            alert(`Please fill in the required field: ${element.previousElementSibling.textContent}`);
            element.focus();
            return false;
        }
    }

    // Check if there's at least one line item with values
    const lineItems = document.querySelectorAll('.line-item');
    let hasValidItem = false;

    lineItems.forEach(item => {
        const description = item.querySelector('.line-description').value.trim();
        const price = parseFloat(item.querySelector('.line-price').value) || 0;
        if (description && price > 0) {
            hasValidItem = true;
        }
    });

    if (!hasValidItem) {
        alert('Please add at least one service or part with a description and price.');
        return false;
    }

    return true;
}

// Complete invoice and show preview
function completeInvoice() {
    if (!validateForm()) {
        return;
    }

    // Generate invoice preview
    const invoiceContent = generateInvoiceHTML();
    document.getElementById('invoiceContent').innerHTML = invoiceContent;

    // Show preview and hide form
    document.querySelector('.invoice-form').style.display = 'none';
    document.getElementById('invoicePreview').style.display = 'block';

    // Enable PDF export button
    document.getElementById('exportPdfBtn').disabled = false;

    // Scroll to top
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

// Generate invoice HTML for preview
function generateInvoiceHTML() {
    // Get all form values
    const shopName = document.getElementById('shopName').value;
    const shopPhone = document.getElementById('shopPhone').value;
    const shopEmail = document.getElementById('shopEmail').value;
    const shopAddress = document.getElementById('shopAddress').value;

    const invoiceNumber = document.getElementById('invoiceNumber').value;
    const invoiceDate = document.getElementById('invoiceDate').value;

    const customerName = document.getElementById('customerName').value;
    const customerPhone = document.getElementById('customerPhone').value;
    const customerEmail = document.getElementById('customerEmail').value;
    const customerAddress = document.getElementById('customerAddress').value;

    const vehicleYear = document.getElementById('vehicleYear').value;
    const vehicleMake = document.getElementById('vehicleMake').value;
    const vehicleModel = document.getElementById('vehicleModel').value;
    const vehicleVIN = document.getElementById('vehicleVIN').value;
    const vehicleLicense = document.getElementById('vehicleLicense').value;
    const vehicleMileage = document.getElementById('vehicleMileage').value;

    const notes = document.getElementById('notes').value;

    // Generate line items table
    let lineItemsHTML = '';
    const lineItems = document.querySelectorAll('.line-item');
    lineItems.forEach(item => {
        const description = item.querySelector('.line-description').value;
        const quantity = item.querySelector('.line-quantity').value;
        const price = parseFloat(item.querySelector('.line-price').value) || 0;
        const total = quantity * price;

        if (description && price > 0) {
            lineItemsHTML += `
                <tr>
                    <td>${description}</td>
                    <td class="text-right">${quantity}</td>
                    <td class="text-right">${formatCurrency(price)}</td>
                    <td class="text-right"><strong>${formatCurrency(total)}</strong></td>
                </tr>
            `;
        }
    });

    const subtotal = document.getElementById('subtotal').textContent;
    const tax = document.getElementById('tax').textContent;
    const total = document.getElementById('total').textContent;
    const taxRate = document.getElementById('taxRate').value;

    // Build invoice HTML
    return `
        <div class="invoice-header">
            <div class="invoice-shop-info">
                <h2>${shopName}</h2>
                ${shopAddress ? `<p>${shopAddress}</p>` : ''}
                <p>Phone: ${shopPhone}</p>
                ${shopEmail ? `<p>Email: ${shopEmail}</p>` : ''}
            </div>
            <div class="invoice-details">
                <p><strong>Invoice #:</strong> ${invoiceNumber}</p>
                <p><strong>Date:</strong> ${formatDate(invoiceDate)}</p>
            </div>
        </div>

        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 30px; margin-bottom: 30px;">
            <div class="invoice-section">
                <h3>Customer Information</h3>
                <p><strong>Name:</strong> ${customerName}</p>
                <p><strong>Phone:</strong> ${customerPhone}</p>
                ${customerEmail ? `<p><strong>Email:</strong> ${customerEmail}</p>` : ''}
                ${customerAddress ? `<p><strong>Address:</strong> ${customerAddress}</p>` : ''}
            </div>

            <div class="invoice-section">
                <h3>Vehicle Information</h3>
                <p><strong>Vehicle:</strong> ${vehicleYear} ${vehicleMake} ${vehicleModel}</p>
                ${vehicleVIN ? `<p><strong>VIN:</strong> ${vehicleVIN}</p>` : ''}
                ${vehicleLicense ? `<p><strong>License:</strong> ${vehicleLicense}</p>` : ''}
                ${vehicleMileage ? `<p><strong>Mileage:</strong> ${vehicleMileage}</p>` : ''}
            </div>
        </div>

        <div class="invoice-section">
            <h3>Services & Parts</h3>
            <table class="invoice-table">
                <thead>
                    <tr>
                        <th>Description</th>
                        <th class="text-right">Quantity</th>
                        <th class="text-right">Price</th>
                        <th class="text-right">Total</th>
                    </tr>
                </thead>
                <tbody>
                    ${lineItemsHTML}
                </tbody>
            </table>
        </div>

        <div class="invoice-totals">
            <table class="invoice-totals-table">
                <tr>
                    <td>Subtotal:</td>
                    <td class="text-right"><strong>${subtotal}</strong></td>
                </tr>
                <tr>
                    <td>Tax (${taxRate}%):</td>
                    <td class="text-right"><strong>${tax}</strong></td>
                </tr>
                <tr class="total-final-row">
                    <td>Total:</td>
                    <td class="text-right"><strong>${total}</strong></td>
                </tr>
            </table>
        </div>

        ${notes ? `
            <div class="invoice-notes">
                <h3>Notes / Terms</h3>
                <p>${notes.replace(/\n/g, '<br>')}</p>
            </div>
        ` : ''}
    `;
}

// Format date
function formatDate(dateString) {
    const date = new Date(dateString);
    return date.toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    });
}

// Export to PDF
async function exportToPDF() {
    const { jsPDF } = window.jspdf;
    const invoiceContent = document.getElementById('invoiceContent');

    try {
        // Show loading state
        const exportBtn = document.getElementById('exportPdfBtn');
        const originalText = exportBtn.textContent;
        exportBtn.textContent = '⏳ Generating PDF...';
        exportBtn.disabled = true;

        // Use html2canvas to capture the invoice
        const canvas = await html2canvas(invoiceContent, {
            scale: 2,
            useCORS: true,
            logging: false,
            backgroundColor: '#ffffff'
        });

        const imgData = canvas.toDataURL('image/png');

        // Create PDF
        const pdf = new jsPDF({
            orientation: 'portrait',
            unit: 'mm',
            format: 'a4'
        });

        const imgWidth = 210; // A4 width in mm
        const imgHeight = (canvas.height * imgWidth) / canvas.width;

        pdf.addImage(imgData, 'PNG', 0, 0, imgWidth, imgHeight);

        // Save PDF
        const invoiceNumber = document.getElementById('invoiceNumber').value;
        pdf.save(`Invoice_${invoiceNumber}.pdf`);

        // Restore button
        exportBtn.textContent = originalText;
        exportBtn.disabled = false;

        alert('Invoice PDF downloaded successfully!');
    } catch (error) {
        console.error('Error generating PDF:', error);
        alert('Error generating PDF. Please try again.');

        const exportBtn = document.getElementById('exportPdfBtn');
        exportBtn.textContent = '📄 Export as PDF';
        exportBtn.disabled = false;
    }
}

// Edit invoice (go back to form)
function editInvoice() {
    document.querySelector('.invoice-form').style.display = 'block';
    document.getElementById('invoicePreview').style.display = 'none';
    document.getElementById('exportPdfBtn').disabled = true;
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

// Clear form
function clearForm() {
    if (!confirm('Are you sure you want to clear all fields? This cannot be undone.')) {
        return;
    }

    // Clear all input fields
    document.querySelectorAll('input[type="text"], input[type="tel"], input[type="email"], input[type="number"], textarea').forEach(input => {
        if (input.id !== 'taxRate' && input.id !== 'invoiceDate' && input.id !== 'invoiceNumber') {
            input.value = '';
        }
    });

    // Reset line items
    document.getElementById('lineItemsContainer').innerHTML = '';
    lineItemCounter = 0;
    addLineItem();
    addLineItem();

    // Reset totals
    calculateTotals();

    // Generate new invoice number
    document.getElementById('invoiceNumber').value = generateInvoiceNumber();

    // Scroll to top
    window.scrollTo({ top: 0, behavior: 'smooth' });
}
