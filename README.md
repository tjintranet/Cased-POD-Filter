# Cased POD Filter

A web-based application for filtering Print-On-Demand (POD) orders and generating PDF work lists.

## Overview

Cased POD Filter is a web application designed to help efficiently filter, sort, and process Print-On-Demand order data from Excel files. It provides an easy-to-use interface for:

- Uploading and processing Excel files containing order data
- Filtering orders by multiple criteria (Customer, Bind Method, Spine size, etc.)
- Sorting orders by spine size for efficient production
- Calculating total quantities of filtered orders
- Generating PDF work lists with applied filters
- Exporting filtered results as customized PDF documents

## Features

### File Handling
- Import Excel (.xlsx) files via file browser or drag-and-drop
- Client-side processing (no data is sent to any server)
- Support for Excel files with up to 3,000 rows

### Data Display
- Interactive data table with sorting, filtering, and pagination
- Configurable column widths optimized for book manufacturing data
- Automatic total quantity calculation for filtered results

### Filtering Capabilities
- Dynamic filter dropdowns based on Excel data
- Multiple simultaneous filters (Customer, Bind Method, Spine, etc.)
- Visual indication of active filters
- One-click filter reset

### PDF Generation
- Customized PDF output formatted for work orders
- Filename automatically includes applied filters
- Date and time stamping
- Optimized column widths for different data types

### Integration
- Direct link to PACE Enquiry system
- Consistent styling and user experience

## Using the Application

### Importing Excel Files

1. When you first open the application, you'll see the file upload screen
2. Click the "Browse Files" button to select an Excel file, or drag and drop a file onto the upload area
3. The application will process the file and display the data in a table

### Filtering Data

1. Once data is loaded, filter dropdowns will appear above the table
2. Select values in one or more dropdowns to filter the data
3. Click "Apply Filters" to execute the filtering
4. Active filters will be displayed as badges
5. To remove filters, click "Reset All Filters"

### Sorting and Viewing Data

1. Use the "Sort by Spine Size" button to sort the table by spine width
2. The total quantity of the displayed orders is shown above the table
3. Use the search box to find specific data within the filtered results
4. Adjust the number of entries shown per page using the dropdown

### PDF and Excel Generation

1. After filtering the data as needed, you have two export options:
   - Click "Generate PDF Work List" to create and download a PDF
   - Click "Download Excel" to export the filtered data as an Excel file
2. Both file types will include only the filtered data
3. The filenames will include the active filters (e.g., `Cased_POD_Customer-ABC_Spine-12.pdf`)
4. The PDF includes a date/time stamp and formatted table of filtered orders
5. The Excel file preserves all data in a format that can be further processed

### Starting Over

1. To process a new file, click "Select New File"
2. This will return you to the upload screen where you can select another file

## Excel File Format

The application works best with Excel files that contain the following columns:
- Customer
- Customer Order No.
- Bind Method
- H (Height)
- W (Width)
- Spine
- Text Paper
- Title
- ISBN (or Code)
- Quantity (or Qty)

Additional columns may be present but will not be used for filtering.

## Technical Details

### Technologies Used

- HTML5, CSS3, JavaScript (ES6+)
- Bootstrap 5 for responsive UI
- jQuery for DOM manipulation
- DataTables for interactive table functionality
- SheetJS (xlsx) for client-side Excel processing
- jsPDF and jspdf-autotable for PDF generation

### Browser Compatibility

The application is compatible with modern browsers:
- Chrome (recommended)
- Firefox
- Edge
- Safari

### Performance Considerations

- Excel processing happens entirely in the browser
- Large files (>3,000 rows) may cause performance issues in older browsers
- PDF generation may take longer for large datasets

## Installation

### Hosting

This is a purely client-side application and can be hosted on any web server or static site hosting service:

1. Download all files from the repository
2. Upload to your web server or hosting service
3. No server-side processing or database is required

### Local Testing

To test locally:

1. Clone the repository or download the files
2. Open `index.html` in a web browser
   - Note: For security reasons, some browsers may block local file access. In this case, use a local development server:

```bash
# Using Python's built-in server
python -m http.server

# Or using Node.js with http-server (npm install -g http-server)
http-server
```

## Customization

### Styling

The application uses Bootstrap 5 and custom CSS. To modify the appearance:

1. Edit the CSS styles in the `<style>` section of `index.html`
2. For major styling changes, consider extracting the styles to a separate CSS file

### Application Logic

To modify the application behavior:

1. Edit `app.js` to change the application logic
2. Key functions that may require customization:
   - `processExcelFile()` - For changing how Excel files are processed
   - `displayData()` - For modifying how data is displayed
   - `generatePDF()` - For customizing PDF output

### PACE Integration

The "PACE Enquiry" button links to a specific URL. To change this:

1. Find the following line in `index.html`:
   ```html
   <a href="http://192.168.10.251/epace/company:c001/inquiry/UserDefinedInquiry/view/5294?" target="_blank" class="btn btn-light">PACE Enquiry</a>
   ```
2. Update the `href` attribute to point to the correct URL

## Troubleshooting

### Common Issues

1. **File not loading**
   - Ensure the file is in .xlsx format (not .xls)
   - Check that the file is not corrupted
   - Try with a smaller file if browser performance is an issue

2. **Filters not working**
   - Verify the Excel file has the expected column names
   - Check for inconsistent formatting in the Excel data
   - Ensure column names match what the application expects

3. **PDF generation errors**
   - Verify there is data in the table to export
   - Check if the filtered dataset is not too large
   - Try with fewer filters if the PDF is too complex

4. **Display issues**
   - Ensure you're using a modern, updated browser
   - Try clearing your browser cache
   - Check browser console for JavaScript errors