let preloadSheets = [];
let postloadSheets = [];
let sheetColumns = {};

async function loadAllSheets() {
    const preloadFile = document.getElementById('preloadFile').files[0];
    const postloadFile = document.getElementById('postloadFile').files[0];

    if (!preloadFile || !postloadFile) {
        showStatus('Please select both pre-load and post-load files', 'error');
        return;
    }

    showStatus('Loading sheets and columns...', 'info');

    try {
        // Load pre-load file sheets and columns
        await loadSheets('preload');
        document.getElementById('preloadStatus').textContent = '✓ Sheets loaded';
        document.getElementById('preloadStatus').className = 'file-status success';
        
        // Load post-load file sheets and columns
        await loadSheets('postload');
        document.getElementById('postloadStatus').textContent = '✓ Sheets loaded';
        document.getElementById('postloadStatus').className = 'file-status success';

        showStatus('Sheets and columns loaded successfully', 'success');
    } catch (error) {
        console.error('Error loading sheets and columns:', error);
        showStatus('Error loading sheets and columns', 'error');
    }
}

async function loadSheets(fileType) {
    const fileInput = document.getElementById(`${fileType}File`);
    const file = fileInput.files[0];
    
    if (file) {
        const formData = new FormData();
        formData.append('file', file);
        
        try {
            const response = await fetch(`/get_sheets/${fileType}`, {
                method: 'POST',
                body: formData
            });
            
            const sheets = await response.json();
            console.log(`${fileType} sheets:`, sheets);
            
            if (fileType === 'preload') {
                preloadSheets = sheets;
                // Load columns for each preload sheet
                for (const sheet of sheets) {
                    const columns = await loadColumnsForSheet('preload', sheet);
                    sheetColumns[`preload_${sheet}`] = columns;
                    console.log(`Preload columns for ${sheet}:`, columns);
                }
            } else {
                postloadSheets = sheets;
                // Load columns for each postload sheet
                for (const sheet of sheets) {
                    const columns = await loadColumnsForSheet('postload', sheet);
                    sheetColumns[`postload_${sheet}`] = columns;
                    console.log(`Postload columns for ${sheet}:`, columns);
                }
            }
            
            updateSheetMappings();
        } catch (error) {
            console.error('Error loading sheets:', error);
            throw error;
        }
    }
}

async function loadColumnsForSheet(fileType, sheetName) {
    try {
        const response = await fetch(`/get_columns/${fileType}/${sheetName}`);
        return await response.json();
    } catch (error) {
        console.error(`Error loading columns for ${fileType} ${sheetName}:`, error);
        return [];
    }
}

function updateSheetMappings() {
    const mappingContainer = document.getElementById('sheetMappings');
    mappingContainer.innerHTML = '';
    
    preloadSheets.forEach((preSheet) => {
        const mappingRow = document.createElement('div');
        mappingRow.className = 'mapping-row';
        
        // Preload sheet name (static)
        const preSheetLabel = document.createElement('div');
        preSheetLabel.className = 'sheet-label';
        preSheetLabel.textContent = preSheet;
        
        // Postload sheet dropdown
        const postSelect = document.createElement('select');
        postSelect.className = 'postload-sheet';
        postSelect.innerHTML = `
            <option value="">Select Post-load Sheet</option>
            <option value="none">None (Skip Comparison)</option>
        `;
        postloadSheets.forEach(sheet => {
            const option = document.createElement('option');
            option.value = sheet;
            option.textContent = sheet;
            postSelect.appendChild(option);
        });
        
        // Key column selection
        const keySelect = document.createElement('select');
        keySelect.className = 'key-column';
        keySelect.innerHTML = '<option value="">Select Key Column</option>';
        keySelect.disabled = true;
        
        // Handle file name display
        document.getElementById('preloadFile').addEventListener('change', function(e) {
            document.getElementById('preloadFileName').textContent = e.target.files[0].name;
        });
        
        document.getElementById('postloadFile').addEventListener('change', function(e) {
            document.getElementById('postloadFileName').textContent = e.target.files[0].name;
        });
        
        // Update key column options when postload sheet is selected
        postSelect.onchange = () => {
            const selectedPostSheet = postSelect.value;
            keySelect.innerHTML = '<option value="">Select Key Column</option>';
            
            if (selectedPostSheet && selectedPostSheet !== 'none') {
                keySelect.disabled = false;
                const preColumns = sheetColumns[`preload_${preSheet}`] || [];
                const postColumns = sheetColumns[`postload_${selectedPostSheet}`] || [];
                const commonColumns = preColumns.filter(col => postColumns.includes(col));
                
                commonColumns.forEach(column => {
                    const option = document.createElement('option');
                    option.value = column;
                    option.textContent = column;
                    keySelect.appendChild(option);
                });
            } else {
                keySelect.disabled = true;
            }
        };
        
        mappingRow.appendChild(preSheetLabel);
        mappingRow.appendChild(postSelect);
        mappingRow.appendChild(keySelect);
        mappingContainer.appendChild(mappingRow);
    });
}

async function compareFiles() {
    const mappings = [];
    document.querySelectorAll('.mapping-row').forEach(row => {
        const preSheet = row.querySelector('.sheet-label').textContent;
        const postSheet = row.querySelector('.postload-sheet').value;
        const keyColumn = row.querySelector('.key-column').value;
        
        if (preSheet && postSheet && (postSheet === 'none' || keyColumn)) {
            mappings.push({
                preloadSheet: preSheet,
                postloadSheet: postSheet,
                keyColumn: keyColumn
            });
        }
    });
    
    if (mappings.length === 0) {
        showStatus('Please complete at least one sheet mapping with key column', 'error');
        return;
    }
    
    try {
        showStatus('Comparing files...', 'info');
        
        const response = await fetch('/compare', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ sheetMappings: mappings })
        });
        
        if (response.ok) {
            const result = await response.json();
            if (result.downloadReady) {
                // Create download button
                const downloadBtn = document.createElement('a');
                downloadBtn.href = '/download_result';
                downloadBtn.className = 'action-button download-button';
                downloadBtn.innerHTML = '⬇️ Download Comparison Result';
                downloadBtn.download = 'comparison_result.xlsx';
                
                // Add or replace download button
                const existingBtn = document.querySelector('.download-button');
                if (existingBtn) {
                    existingBtn.remove();
                }
                document.querySelector('.container').appendChild(downloadBtn);
                
                showStatus('Comparison completed successfully. Click the download button to get the result.', 'success');
            }
        } else {
            const error = await response.json();
            showStatus(`Error during comparison: ${error.error}`, 'error');
        }
    } catch (error) {
        console.error('Error during comparison:', error);
        showStatus('Error during comparison', 'error');
    }
}

function showStatus(message, type) {
    const statusDiv = document.getElementById('status');
    statusDiv.textContent = message;
    statusDiv.className = `status-message ${type}`;
}

function updateFileName(fileInput, fileNameId, statusId) {
    const fileName = fileInput.files[0]?.name;
    const fileNameElement = document.getElementById(fileNameId);
    const statusElement = document.getElementById(statusId);
    
    if (fileName) {
        fileNameElement.textContent = fileName;
        statusElement.textContent = '✓ File selected';
        statusElement.className = 'file-status success';
    } else {
        fileNameElement.textContent = 'No file selected';
        statusElement.textContent = '';
        statusElement.className = 'file-status';
    }
}

// Add these event listeners
document.getElementById('preloadFile').addEventListener('change', function(e) {
    updateFileName(this, 'preloadFileName', 'preloadStatus');
});

document.getElementById('postloadFile').addEventListener('change', function(e) {
    updateFileName(this, 'postloadFileName', 'postloadStatus');
});

function updateColumnSelectors(preSheet, postSheet) {
    fetch('/get_column_suggestions', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({
            preSheet: preSheet,
            postSheet: postSheet
        })
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            const suggestions = data.suggestions;
            
            // Update the column selectors with suggestions
            const mappingRows = document.querySelectorAll('.mapping-row');
            mappingRows.forEach(row => {
                const postSheet = row.querySelector('.sheet-select').value;
                if (postSheet !== 'none') {
                    const columnSelect = row.querySelector('.column-select');
                    updateColumnSelectWithSuggestions(columnSelect, suggestions);
                }
            });
        } else {
            console.error('Error getting column suggestions:', data.error);
        }
    })
    .catch(error => console.error('Error:', error));
}

function updateColumnSelectWithSuggestions(select, suggestions) {
    // Keep the default option
    select.innerHTML = '<option value="">Select Key Column</option>';
    
    // Add all suggested columns with their matches
    Object.entries(suggestions).forEach(([postCol, preColumns]) => {
        const optgroup = document.createElement('optgroup');
        optgroup.label = `Matches for: ${postCol}`;
        
        preColumns.forEach(preCol => {
            const option = document.createElement('option');
            option.value = preCol;
            option.textContent = `${preCol} (matches ${postCol})`;
            optgroup.appendChild(option);
        });
        
        select.appendChild(optgroup);
    });
} 