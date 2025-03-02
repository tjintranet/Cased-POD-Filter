// Global variables
let dataTable;
let excelData = {};
let originalData = {};
let currentFileName = null;

// DOM elements
const fileInput = document.getElementById('fileInput');
const browseButton = document.getElementById('browseButton');
const uploadArea = document.getElementById('uploadArea');
const loadingIndicator = document.getElementById('loadingIndicator');
const tableContainer = document.getElementById('tableContainer');
const filterControls = document.getElementById('filterControls');
const tableHeader = document.getElementById('tableHeader');
const tableBody = document.getElementById('tableBody');
const filterButtonsContainer = document.getElementById('filterButtonsContainer');
const generatePdfButton = document.getElementById('generatePdfButton');
const downloadExcelButton = document.getElementById('downloadExcelButton');
const resetFiltersButton = document.getElementById('resetFiltersButton');
const sortSpineButton = document.getElementById('sortSpineButton');
const backToUploadButton = document.getElementById('backToUploadButton');
const applyFiltersButton = document.getElementById('applyFiltersButton');
const filterSummary = document.getElementById('filterSummary');
const activeFilters = document.getElementById('activeFilters');
const qtySummary = document.getElementById('qtySummary');

// Set up event listeners
document.addEventListener('DOMContentLoaded', function() {
    console.log('DOM loaded, setting up event listeners');
    
    // File input change event
    fileInput.addEventListener('change', handleFileSelection);
    
    // Browse button click event
    browseButton.addEventListener('click', function() {
        fileInput.click();
    });
    
    // Drag and drop events
    setupDragAndDrop();
    
    // Generate PDF button click event
    generatePdfButton.addEventListener('click', generatePDF);
    
    // Download Excel button click event
    downloadExcelButton.addEventListener('click', downloadExcel);
    
    // Reset filters button click event
    resetFiltersButton.addEventListener('click', resetFilters);
    
    // Sort by spine button click event
    sortSpineButton.addEventListener('click', sortBySpine);
    
    // Apply filters button click event
    applyFiltersButton.addEventListener('click', applyFilters);
    
    // Back to upload button click event
    backToUploadButton.addEventListener('click', function() {
        // Reset the UI to show the upload area
        tableContainer.classList.add('hidden');
        filterControls.classList.add('hidden');
        uploadArea.classList.remove('hidden');
        
        // Clear any existing data
        if (dataTable) {
            try {
                dataTable.destroy();
                console.log("DataTable destroyed");
            } catch(e) {
                console.error("Error destroying DataTable:", e);
            }
            dataTable = null;
        }
        
        tableHeader.innerHTML = '';
        tableBody.innerHTML = '';
        filterButtonsContainer.innerHTML = '';
        activeFilters.innerHTML = '';
        filterSummary.classList.add('hidden');
        
        // Reset the file input so the same file can be selected again
        fileInput.value = '';
        currentFileName = null;
        
        // Add this to help with memory cleanup
        window.setTimeout(function() {
            console.log("Running garbage collection");
            if (window.gc) window.gc();
        }, 100);
    });
});

// Setup drag and drop functionality
function setupDragAndDrop() {
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        uploadArea.addEventListener(eventName, preventDefaults, false);
    });
    
    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }
    
    ['dragenter', 'dragover'].forEach(eventName => {
        uploadArea.addEventListener(eventName, function() {
            uploadArea.classList.add('dragover');
        }, false);
    });
    
    ['dragleave', 'drop'].forEach(eventName => {
        uploadArea.addEventListener(eventName, function() {
            uploadArea.classList.remove('dragover');
        }, false);
    });
    
    uploadArea.addEventListener('drop', function(e) {
        const dt = e.dataTransfer;
        const files = dt.files;
        
        if (files.length > 0 && files[0].name.endsWith('.xlsx')) {
            fileInput.files = files;
            handleFileSelection();
        } else {
            alert('Please select an Excel (.xlsx) file');
        }
    }, false);
}

// Handle file selection
function handleFileSelection() {
    try {
        if (fileInput.files.length === 0) {
            return;
        }
        
        const file = fileInput.files[0];
        console.log('Selected file:', file.name);
        
        if (!file.name.endsWith('.xlsx')) {
            alert('Please select an Excel (.xlsx) file');
            return;
        }
        
        currentFileName = file.name;
        
        // Show loading indicator
        uploadArea.classList.add('hidden');
        loadingIndicator.classList.remove('hidden');
        tableContainer.classList.add('hidden'); // Hide table until ready
        
        // Process the Excel file
        console.log('Processing Excel file...');
        processExcelFile(file)
            .then(data => {
                console.log('Excel file processed:', data);
                
                // Store the original data
                originalData = data;
                
                // Store the data for later use
                excelData = data;
                
                // Display the data
                displayData(data);
            })
            .catch(error => {
                console.error('Error processing file:', error);
                alert('Error processing the Excel file: ' + error.message);
                resetUI();
            });
    } catch (error) {
        console.error('Error selecting file:', error);
        alert('Error processing the Excel file: ' + error.message);
        resetUI();
    }
}

// Process Excel file
function processExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                
                // Get the first worksheet
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Convert to array of objects
                const rows = XLSX.utils.sheet_to_json(worksheet, {header: 1, raw: false});
                
                if (rows.length < 2) {
                    throw new Error('No data found in the Excel file');
                }
                
                // Extract headers (first row)
                let headers = rows[0].map(header => header ? header.trim() : '');
                
                // Find the "Code" column and rename it to "ISBN"
                const codeColumnIndex = headers.findIndex(header => 
                    header && header.trim() === 'Code'
                );

                if (codeColumnIndex !== -1) {
                    // Replace "Code" with "ISBN" in the headers array
                    headers[codeColumnIndex] = 'ISBN';
                }
                
                // Find Title and ISBN column indices
                const titleColumnIndex = headers.findIndex(header => 
                    header && header.trim() === 'Title'
                );
                const isbnColumnIndex = headers.findIndex(header => 
                    header && header.trim() === 'ISBN'
                );

                // Swap the Title and ISBN columns if both exist
                if (titleColumnIndex !== -1 && isbnColumnIndex !== -1 && titleColumnIndex < isbnColumnIndex) {
                    // Swap positions in headers array
                    [headers[titleColumnIndex], headers[isbnColumnIndex]] = 
                    [headers[isbnColumnIndex], headers[titleColumnIndex]];
                    
                    // Also swap data in each row
                    for (let i = 1; i < rows.length; i++) {
                        if (rows[i][titleColumnIndex] !== undefined && rows[i][isbnColumnIndex] !== undefined) {
                            [rows[i][titleColumnIndex], rows[i][isbnColumnIndex]] = 
                            [rows[i][isbnColumnIndex], rows[i][titleColumnIndex]];
                        }
                    }
                }
                
                // Find the indices of columns we want to filter
                const columnsToFilter = ['Customer', 'Customer Order No.', 'Bind Method', 'H', 'W', 'Spine', 'Text Paper'];
                const filteredColumnIndices = [];
                const filteredHeaders = [];
                
                headers.forEach((header, index) => {
                    if (!header) return;
                    
                    const isFilterColumn = columnsToFilter.some(col => 
                        header.includes(col) || 
                        (col === 'H' && (header === 'H' || header === 'h')) ||
                        (col === 'W' && (header === 'W' || header === 'w'))
                    );
                    
                    if (isFilterColumn) {
                        filteredColumnIndices.push(index);
                        filteredHeaders.push(header);
                    }
                });
                
                // Make sure Title, ISBN, and Quantity columns are included even if not filtered
                headers.forEach((header, index) => {
                    if (!header) return;
                    
                    // Don't add duplicates
                    if (filteredColumnIndices.includes(index)) return;
                    
                    // Check for special columns we want to display but not filter
                    const isSpecialColumn = 
                        header === 'Title' || 
                        header === 'ISBN' || 
                        header === 'Quantity' || 
                        header === 'Qty';
                    
                    if (isSpecialColumn) {
                        filteredColumnIndices.push(index);
                        filteredHeaders.push(header);
                    }
                });
                
                // Extract unique values for each column (for dropdowns)
                const uniqueValues = {};
                filteredHeaders.forEach(header => {
                    // Only add unique values for filterable columns
                    if (columnsToFilter.some(col => header.includes(col) || 
                        (col === 'H' && (header === 'H' || header === 'h')) ||
                        (col === 'W' && (header === 'W' || header === 'w')))) {
                        uniqueValues[header] = new Set();
                    }
                });
                
                // Extract data from rows
                const dataRows = [];
                let rowCount = 0;

                for (let i = 1; i < rows.length; i++) {
                    const row = rows[i];
                    const rowData = {};
                    let hasData = false;
                    
                    filteredColumnIndices.forEach((colIndex, j) => {
                        let value = '';
                        
                        if (row[colIndex] !== undefined) {
                            value = String(row[colIndex]).trim();
                            hasData = true;
                        }
                        
                        rowData[filteredHeaders[j]] = value;
                        
                        // Add to unique values for dropdowns (only for filterable columns)
                        const headerName = filteredHeaders[j];
                        if (value && uniqueValues[headerName]) {
                            uniqueValues[headerName].add(value);
                        }
                    });
                    
                    if (hasData) {
                        dataRows.push(rowData);
                        rowCount++;
                    }
                }
                
                // Convert uniqueValues from Sets to sorted Arrays
                Object.keys(uniqueValues).forEach(key => {
                    if (key.includes('Spine')) {
                        // Special handling for spine values - sort numerically
                        uniqueValues[key] = [...uniqueValues[key]].sort((a, b) => {
                            const numA = parseFloat(a);
                            const numB = parseFloat(b);
                            if (isNaN(numA) || isNaN(numB)) return a.localeCompare(b);
                            return numA - numB;
                        });
                    } else {
                        uniqueValues[key] = [...uniqueValues[key]].sort((a, b) => a.localeCompare(b));
                    }
                });
                
                resolve({
                    headers: filteredHeaders,
                    data: dataRows,
                    uniqueValues,
                    totalRows: rowCount,
                    filteredRowCount: dataRows.length
                });
            } catch (error) {
                console.error('Error processing Excel data:', error);
                reject(error);
            }
        };
        
        reader.onerror = function(e) {
            reject(new Error('Error reading file'));
        };
        
        reader.readAsArrayBuffer(file);
    });
}

// Display data in table and create filters
function displayData(data) {
    console.log("Displaying data:", data);
    
    if (!data || !data.headers || !data.data) {
        console.error("Invalid data format:", data);
        alert("Invalid data format received");
        return;
    }
    
    // Destroy existing DataTable if it exists
    if (dataTable) {
        console.log("Destroying existing DataTable");
        try {
            dataTable.destroy();
        } catch(e) {
            console.error("Error destroying DataTable:", e);
        }
        dataTable = null;
    }
    
    // Clear the table completely
    tableHeader.innerHTML = '';
    tableBody.innerHTML = '';
    
    // Create table headers with proper spacing
    data.headers.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        
        // Set width for specific columns
        if (header === 'H' || header === 'W' || header === 'Spine') {
            th.style.minWidth = '60px';
            th.style.maxWidth = '60px';
            th.style.width = '60px';
        } else if (header === 'Quantity' || header === 'Qty') {
            th.style.minWidth = '70px';
            th.style.maxWidth = '70px';
            th.style.width = '70px';
        } else {
            // Ensure headers don't get merged visually
            th.style.minWidth = '80px';
        }
        
        th.style.paddingRight = '10px';
        tableHeader.appendChild(th);
    });
    
    // Hide loading indicator and show table
    loadingIndicator.classList.add('hidden');
    tableContainer.classList.remove('hidden');
    filterControls.classList.remove('hidden');
    
    // Create filter dropdowns
    createFilterDropdowns(data);
    
    // Format data for DataTables
    const tableData = data.data.map(row => {
        return data.headers.map(header => row[header] || '');
    });
    
    console.log("Initializing DataTable with", tableData.length, "rows");
    
    // Initialize DataTable
    dataTable = $('#excelDataTable').DataTable({
        data: tableData,
        columns: data.headers.map(header => ({ 
            title: header,
            // Ensure proper display for each column type
            className: ['H', 'W', 'Spine', 'Quantity', 'Qty'].includes(header) ? 'text-center' : ''
        })),
        paging: true,
        searching: true,
        ordering: true,
        info: true,
        responsive: true,
        pageLength: 25,
        lengthMenu: [10, 25, 50, 100],
        columnDefs: [
            {
                // Special sorting for spine column
                targets: data.headers.findIndex(h => h.includes('Spine')),
                type: 'num-fmt',
                render: function(data, type) {
                    // For sorting, convert to number
                    if (type === 'sort') {
                        return parseFloat(data) || 0;
                    }
                    // For display, keep original
                    return data;
                }
            },
            {
                // Special sorting for quantity column (if present)
                targets: data.headers.findIndex(h => h === 'Quantity' || h === 'Qty'),
                type: 'num-fmt',
                render: function(data, type) {
                    // For sorting, convert to number
                    if (type === 'sort') {
                        return parseInt(data) || 0;
                    }
                    // For display, keep original
                    return data;
                }
            }
        ].filter(def => def.targets !== -1), // Only include if column exists
        drawCallback: function() {
            // Calculate and display quantity sum after table draw
            updateQuantitySum();
        }
    });
    
    // Calculate and display quantity sum
    updateQuantitySum();
    
    console.log("DataTable initialized successfully");
}

// Create filter dropdowns
function createFilterDropdowns(data) {
    filterButtonsContainer.innerHTML = '';
    
    data.headers.forEach((header, index) => {
        // Get unique values for this column
        const uniqueValues = data.uniqueValues[header] || [];
        
        // Skip if no unique values (non-filterable column)
        if (uniqueValues.length === 0) return;
        
        // Create the filter dropdown
        const colDiv = document.createElement('div');
        colDiv.className = 'col-md-4 col-lg-3 mb-3';
        
        const formGroup = document.createElement('div');
        formGroup.className = 'form-group';
        
        const label = document.createElement('label');
        label.textContent = header;
        label.className = 'form-label';
        
        const select = document.createElement('select');
        select.className = 'form-select filter-select';
        select.dataset.column = header;
        
        // Add default option
        const defaultOption = document.createElement('option');
        defaultOption.value = '';
        defaultOption.textContent = 'All';
        select.appendChild(defaultOption);
        
        // Add options for each unique value
        uniqueValues.forEach(value => {
            const option = document.createElement('option');
            option.value = value;
            option.textContent = value;
            select.appendChild(option);
        });
        
        formGroup.appendChild(label);
        formGroup.appendChild(select);
        colDiv.appendChild(formGroup);
        filterButtonsContainer.appendChild(colDiv);
    });
}

// Calculate and update quantity sum
function updateQuantitySum() {
    try {
        const table = $('#excelDataTable').DataTable();
        const headers = Array.from(document.querySelectorAll('#tableHeader th'));
        const qtyColumnIndex = headers.findIndex(th => 
            th.textContent === 'Quantity' || th.textContent === 'Qty'
        );
        
        if (qtyColumnIndex !== -1) {
            let sum = 0;
            table.rows({ search: 'applied' }).every(function(rowIdx) {
                try {
                    const qtyValue = table.cell(rowIdx, qtyColumnIndex).data();
                    const qtyNumber = parseInt(qtyValue, 10);
                    if (!isNaN(qtyNumber)) {
                        sum += qtyNumber;
                    }
                } catch (err) {
                    console.warn("Error processing row", rowIdx, err);
                }
            });
            
            qtySummary.textContent = `Total Quantity: ${sum.toLocaleString()}`;
        } else {
            qtySummary.textContent = 'Total Quantity: N/A';
        }
    } catch (error) {
        console.error("Error updating quantity sum:", error);
        qtySummary.textContent = 'Total Quantity: Error';
    }
}

// Apply filters
function applyFilters() {
    try {
        const filterSelects = document.querySelectorAll('.filter-select');
        const filters = {};
        const activeFilterValues = [];
        
        // Collect filter values
        filterSelects.forEach(select => {
            const column = select.dataset.column;
            const value = select.value;
            
            if (value) {
                filters[column] = value;
                activeFilterValues.push({
                    column: column,
                    value: value
                });
            }
        });
        
        // Update active filters display
        updateActiveFiltersDisplay(activeFilterValues);
        
        // If no filters active, just reset
        if (Object.keys(filters).length === 0) {
            // Reset to original data
            excelData = { ...originalData };
            displayData(excelData);
            return;
        }
        
        // Show loading indicator
        tableContainer.classList.add('hidden');
        loadingIndicator.classList.remove('hidden');
        
        // Apply filters to the data
        const filteredData = applyFiltersToData(originalData, filters);
        
        // Store filtered data
        excelData = filteredData;
        
        // Display filtered data
        displayData(excelData);
        
        // Ensure quantity sum is updated
        updateQuantitySum();
    } catch (error) {
        console.error('Error applying filters:', error);
        alert('Error applying filters: ' + error.message);
        
        // Restore UI
        loadingIndicator.classList.add('hidden');
        tableContainer.classList.remove('hidden');
    }
}

// Apply filters to data
function applyFiltersToData(data, filters) {
    // If no filters, return original data
    if (!filters || Object.keys(filters).length === 0) {
        return data;
    }
    
    // Filter the rows
    const filteredRows = data.data.filter(row => {
        for (const [key, value] of Object.entries(filters)) {
            if (!value) continue; // Skip empty filters
            
            const rowValue = String(row[key] || '').toLowerCase();
            const filterValue = String(value).toLowerCase();
            
            if (rowValue !== filterValue) {
                return false;
            }
        }
        return true;
    });
    
    // Return filtered data with the same structure
    return {
        ...data,
        data: filteredRows,
        filteredRowCount: filteredRows.length
    };
}

// Update the active filters display
function updateActiveFiltersDisplay(activeFilterValues) {
    activeFilters.innerHTML = '';
    
    if (activeFilterValues.length > 0) {
        filterSummary.classList.remove('hidden');
        
        activeFilterValues.forEach(filter => {
            const badge = document.createElement('span');
            badge.className = 'badge bg-primary filter-badge';
            badge.textContent = `${filter.column}: ${filter.value}`;
            activeFilters.appendChild(badge);
        });
    } else {
        filterSummary.classList.add('hidden');
    }
}

// Reset filters
function resetFilters() {
    const filterSelects = document.querySelectorAll('.filter-select');
    filterSelects.forEach(select => {
        select.value = '';
    });
    
    // Clear the active filters display
    activeFilters.innerHTML = '';
    filterSummary.classList.add('hidden');
    
    // Reset to original data
    excelData = { ...originalData };
    displayData(excelData);
    
    // Ensure quantity sum is updated
    updateQuantitySum();
}

// Sort by spine size
function sortBySpine() {
    // Find the spine column
    const spineColumnIndex = Array.from(document.querySelectorAll('#tableHeader th'))
        .findIndex(th => th.textContent.includes('Spine'));
    
    if (spineColumnIndex !== -1) {
        // Apply the sorting to the DataTable
        dataTable.order([spineColumnIndex, 'asc']).draw();
    } else {
        alert('Spine column not found in the table.');
    }
}

// Generate PDF function
function generatePDF() {
    try {
        console.log("Starting PDF generation");
        
        // Get DataTable instance
        const table = $('#excelDataTable').DataTable();
        
        // Get visible data (after filtering)
        const tableRows = [];
        table.rows({ search: 'applied' }).every(function(rowIdx) {
            const rowData = [];
            table.columns().every(function(colIdx) {
                rowData.push(table.cell(rowIdx, colIdx).data());
            });
            tableRows.push(rowData);
        });
        
        if (tableRows.length === 0) {
            alert('No data to export. Please adjust your filters.');
            return;
        }
        
        // Get column headers
        const headers = [];
        table.columns().every(function(index) {
            const headerText = $(table.column(index).header()).text().trim();
            headers.push(headerText);
        });
        
        // Show loading indicator
        loadingIndicator.classList.remove('hidden');
        
        // Collect filters for filename - use active filters from display
        const activeFilterBadges = document.querySelectorAll('.filter-badge');
        console.log("Active filter badges found:", activeFilterBadges.length);
        
        const filenameFilters = [];
        
        activeFilterBadges.forEach(badge => {
            const filterText = badge.textContent.trim();
            console.log("Filter badge text:", filterText);
            
            // Extract column and value (format is "Column: Value")
            const parts = filterText.split(':');
            if (parts.length === 2) {
                const columnName = parts[0].trim().replace(/\s+/g, '');
                const filterValue = parts[1].trim().replace(/\s+/g, '_');
                filenameFilters.push(`${columnName}-${filterValue}`);
                console.log("Added to filename:", `${columnName}-${filterValue}`);
            }
        });
        
        // Create filename with filter information
        let filename = 'Cased_POD';
        if (filenameFilters.length > 0) {
            // Add filters to filename
            filename += '_' + filenameFilters.join('_');
            console.log("Filename with filters:", filename);
        } else {
            filename += '_AllOrders';
            console.log("No filters applied, using default filename:", filename);
        }
        filename += '.pdf';
        
        // Use jsPDF to create PDF
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF({
            orientation: 'landscape',
            unit: 'mm',
            format: 'a4'
        });
        
        // Add title
        doc.setFontSize(18);
        doc.text('Cased POD Work List', 14, 15);
        
        // Add generation date
        doc.setFontSize(10);
        const now = new Date();
        const dateOptions = { day: '2-digit', month: '2-digit', year: 'numeric' };
        const formattedDate = now.toLocaleDateString('en-GB', dateOptions);
        const formattedTime = now.toLocaleTimeString('en-GB');
        doc.text(`Generated on: ${formattedDate} ${formattedTime}`, 14, 22);
        
        // Create column definitions for better layout
        const columnDefinitions = headers.map((header, index) => {
            // Define width based on content type
            let width;
            if (header.includes('Customer')) {
                width = 35; // Customer name
            } else if (header.includes('Order')) {
                width = 25; // Order numbers
            } else if (header === 'H' || header === 'W' || header === 'Spine') {
                width = 15; // Dimensions
            } else if (header.includes('Paper')) {
                width = 32; // Paper descriptions
            } else if (header.includes('Bind')) {
                width = 20; // Bind method
            } else if (header === 'ISBN') {
                width = 25; // Title is usually longer
            } else if (header === 'Title') {
                width = 45; // ISBN
            } else if (header === 'Quantity' || header === 'Qty') {
                width = 15; // Quantity is usually a small number
            } else {
                width = 25; // Default
            }
            
            return {
                header: header,
                dataKey: index.toString(),
                width: width
            };
        });
        
        // Format rows as objects for easier processing
        const tableData = tableRows.map(row => {
            const rowObj = {};
            row.forEach((cell, i) => {
                rowObj[i.toString()] = cell;
            });
            return rowObj;
        });
        
        // Find indices of columns that should be centered
        const centerColumns = {};
        headers.forEach((header, index) => {
            if (header === 'H' || header === 'W' || header === 'Spine' || 
                header === 'Quantity' || header === 'Qty' || 
                header.includes('Bind')) {
                centerColumns[index.toString()] = { halign: 'center' };
            }
        });
        
        // Add the table with fixed columns
        doc.autoTable({
            columns: columnDefinitions,
            body: tableData,
            startY: 29 + 5,
            margin: { top: 10, right: 10, bottom: 10, left: 10 },
            styles: {
                fontSize: 9,
                cellPadding: 3,
                lineColor: [0, 0, 0],
                lineWidth: 0.1
            },
            headStyles: {
                fillColor: [41, 128, 185],
                textColor: 255,
                fontStyle: 'bold',
                halign: 'center',
                valign: 'middle'
            },
            bodyStyles: {
                halign: 'left'
            },
            // Special styling for specific columns
            columnStyles: centerColumns,
            // Ensure table spans across full page width
            tableWidth: 'auto',
        });
        
        // Save the PDF file
        doc.save(filename);
        
        // Hide loading indicator
        loadingIndicator.classList.add('hidden');
        
        console.log('PDF generated and downloaded:', filename);
    } catch (error) {
        // Hide loading indicator
        loadingIndicator.classList.add('hidden');
        
        console.error('Error generating PDF:', error);
        alert('Error generating PDF: ' + error.message);
    }
}

// Reset UI
function resetUI() {
    loadingIndicator.classList.add('hidden');
    uploadArea.classList.remove('hidden');
    tableContainer.classList.add('hidden');
    filterControls.classList.add('hidden');
}

// Download Excel function
function downloadExcel() {
    try {
        console.log("Starting Excel download");
        
        // Show loading indicator
        loadingIndicator.classList.remove('hidden');
        
        // Get DataTable instance
        const table = $('#excelDataTable').DataTable();
        
        // Get visible data (after filtering)
        const rows = [];
        const headers = [];
        
        // Get headers
        table.columns().every(function(index) {
            const headerText = $(table.column(index).header()).text().trim();
            headers.push(headerText);
        });
        
        // Add header row
        rows.push(headers);
        
        // Get data rows
        table.rows({ search: 'applied' }).every(function(rowIdx) {
            const rowData = [];
            table.columns().every(function(colIdx) {
                rowData.push(table.cell(rowIdx, colIdx).data());
            });
            rows.push(rowData);
        });
        
        if (rows.length <= 1) { // Only header row
            alert('No data to export. Please adjust your filters.');
            loadingIndicator.classList.add('hidden');
            return;
        }
        
        // Create workbook and worksheet
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(rows);
        
        // Generate filename with filter information
        let filename = 'Cased_POD';
        
        // Collect filters for filename - use active filters from display
        const activeFilterBadges = document.querySelectorAll('.filter-badge');
        console.log("Active filter badges found:", activeFilterBadges.length);
        
        const filenameFilters = [];
        
        activeFilterBadges.forEach(badge => {
            const filterText = badge.textContent.trim();
            console.log("Filter badge text:", filterText);
            
            // Extract column and value (format is "Column: Value")
            const parts = filterText.split(':');
            if (parts.length === 2) {
                const columnName = parts[0].trim().replace(/\s+/g, '');
                const filterValue = parts[1].trim().replace(/\s+/g, '_');
                filenameFilters.push(`${columnName}-${filterValue}`);
                console.log("Added to filename:", `${columnName}-${filterValue}`);
            }
        });
        
        // Create filename with filter information
        if (filenameFilters.length > 0) {
            // Add filters to filename
            filename += '_' + filenameFilters.join('_');
            console.log("Filename with filters:", filename);
        } else {
            filename += '_AllOrders';
            console.log("No filters applied, using default filename:", filename);
        }
        filename += '.xlsx';
        
        // Add worksheet to workbook
        XLSX.utils.book_append_sheet(wb, ws, 'Filtered Orders');
        
        // Save workbook and trigger download
        XLSX.writeFile(wb, filename);
        
        // Hide loading indicator
        loadingIndicator.classList.add('hidden');
        
        console.log('Excel file generated and downloaded:', filename);
    } catch (error) {
        // Hide loading indicator
        loadingIndicator.classList.add('hidden');
        
        console.error('Error generating Excel file:', error);
        alert('Error generating Excel file: ' + error.message);
    }
}