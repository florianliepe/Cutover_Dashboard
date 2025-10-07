document.addEventListener('DOMContentLoaded', () => {
    const uploadArea = document.getElementById('upload-area');
    const fileInput = document.getElementById('file-input');
    const uploadStatus = document.getElementById('upload-status');

    // --- Event Listeners ---
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        uploadArea.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
        }, false);
    });

    ['dragenter', 'dragover'].forEach(eventName => {
        uploadArea.addEventListener(eventName, () => uploadArea.classList.add('dragover'), false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        uploadArea.addEventListener(eventName, () => uploadArea.classList.remove('dragover'), false);
    });

    uploadArea.addEventListener('drop', (e) => {
        const files = e.dataTransfer.files;
        if (files.length > 0) handleFile(files[0]);
    }, false);

    uploadArea.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) handleFile(e.target.files[0]);
    });

    // --- Core File Handling Function ---
    function handleFile(file) {
        const validExtensions = ['.xlsx', '.xls', '.csv'];
        const fileExtension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();

        if (!validExtensions.includes(fileExtension)) {
            showStatus('Error: Please upload a valid Excel file (.xlsx, .xls, or .csv).', true);
            return;
        }

        showStatus(`Processing "${file.name}"...`, false);

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = e.target.result;
                const workbook = XLSX.read(data, { type: 'binary' });

                const requiredSheet = 'Overview';
                
                // Clear old data
                localStorage.clear();

                if (workbook.Sheets[requiredSheet]) {
                    const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[requiredSheet]);
                    localStorage.setItem('overviewData', JSON.stringify(jsonData));
                    console.log(`Successfully processed and stored "${requiredSheet}".`);
                    
                    showStatus('Success! Redirecting to dashboard...', false);
                    setTimeout(() => {
                        window.location.href = 'dashboard.html';
                    }, 1000);

                } else {
                    showStatus(`Error: Required sheet "${requiredSheet}" was not found in the Excel file.`, true);
                }

            } catch (error) {
                console.error(error);
                showStatus('An unexpected error occurred while processing the file.', true);
            }
        };

        reader.onerror = (error) => {
            console.error(error);
            showStatus('Error reading the file.', true);
        };

        reader.readAsBinaryString(file);
    }

    function showStatus(message, isError) {
        uploadStatus.textContent = message;
        uploadStatus.className = 'status-message';
        uploadStatus.classList.add(isError ? 'status-error' : 'status-success');
    }
});
