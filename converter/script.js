document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const previewSection = document.getElementById('preview-section');
    const tableHead = document.getElementById('table-head');
    const tableBody = document.getElementById('table-body');
    const convertBtn = document.getElementById('convert-btn');
    const resetBtn = document.getElementById('reset-btn');
    const loader = document.getElementById('loader');

    let currentData = null;

    // Trigger file input on click
    dropZone.addEventListener('click', () => fileInput.click());

    // Handle drag and drop
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-over');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleFile(files[0]);
        }
    });

    // Handle file selection
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    });

    function handleFile(file) {
        loader.style.display = 'block';
        dropZone.style.display = 'none';

        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Convert to JSON
            const json = XLSX.utils.sheet_to_json(worksheet);
            
            if (json.length === 0) {
                alert('The file is empty or invalid.');
                reset();
                return;
            }

            // Check for required fields
            const firstRow = json[0];
            const headers = Object.keys(firstRow);
            const required = ['Name', 'Sales Price', 'Barcode'];
            const missing = required.filter(h => !headers.includes(h));

            if (missing.length > 0) {
                alert(`Missing required columns: ${missing.join(', ')}`);
                reset();
                return;
            }

            currentData = json;
            displayPreview(json);
            loader.style.display = 'none';
            previewSection.style.display = 'block';
        };
        reader.readAsArrayBuffer(file);
    }

    function displayPreview(data) {
        // Limit preview to 10 rows
        const previewData = data.slice(0, 10);
        const headers = ['Name', 'Sales Price', 'Barcode'];

        tableHead.innerHTML = headers.map(h => `<th>${h}</th>`).join('');
        tableBody.innerHTML = previewData.map(row => {
            return `<tr>${headers.map(h => `<td>${row[h] || ''}</td>`).join('')}</tr>`;
        }).join('');
    }

    function reset() {
        currentData = null;
        fileInput.value = '';
        dropZone.style.display = 'block';
        previewSection.style.display = 'none';
        loader.style.display = 'none';
    }

    resetBtn.addEventListener('click', reset);

    convertBtn.addEventListener('click', () => {
        if (!currentData) return;

        const outputData = currentData.map((row, index) => {
            const lfCode = index + 1;
            const barcode = String(row['Barcode'] || '');
            
            // Logic for Code: Stripped version of barcode (210012300005 -> 123)
            // Extraction: take characters from index 2 to 7 and parse as int to remove leading zeros
            let extractedCode = '';
            if (barcode.length >= 7) {
                extractedCode = parseInt(barcode.substring(2, 7), 10);
            } else {
                extractedCode = lfCode; // Fallback
            }

            return {
                'LF Code': lfCode,
                'PLU name': row['Name'] || '',
                'Code': extractedCode,
                'Hotkey': lfCode,
                'Unit Price': parseFloat(row['Sales Price']) || 0,
                'Weight unit': 4,
                'Quantity type': 0,
                'Barcode type': 103,
                'Shelf time': 15,
                'Department': 21,
                'Message 1': 0,
                'Message 2': 0,
                'Label': 'D0',
                'Discount/Schedule': 0,
                'Package type': 0,
                'Package Weight': '0.000',
                'Tare': '0.000',
                'Nutrition': 0,
                'Tolerance (%)': 0,
                'MFG date': '',
                'Price model': 3,
                'Sale mode': 0,
                'Min Weight': '0.000',
                'Max Weight': '0.000',
                'MRP': '0.00',
                'Tax type': 0
            };
        });

        const worksheet = XLSX.utils.json_to_sheet(outputData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "PLU Data");

        // Download
        XLSX.writeFile(workbook, "plu_import_ready.xlsx");
    });
});
