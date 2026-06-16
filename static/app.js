document.addEventListener('DOMContentLoaded', () => {
    // DOM Elements
    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('file-input');
    const modePdfToWord = document.getElementById('mode-pdf-to-word');
    const modeWordToPdf = document.getElementById('mode-word-to-pdf');
    const statusPanel = document.getElementById('status-panel');
    const fileNameEl = document.getElementById('file-name');
    const fileSizeEl = document.getElementById('file-size');
    const fileIconContainer = document.getElementById('file-icon-container');
    const progressBar = document.getElementById('progress-bar');
    const progressText = document.getElementById('progress-text');
    const progressPercentage = document.getElementById('progress-percentage');
    const actionButtons = document.getElementById('action-buttons');
    const downloadBtn = document.getElementById('download-btn');
    const resetBtn = document.getElementById('reset-btn');
    const errorBanner = document.getElementById('error-banner');
    const errorMessage = document.getElementById('error-message');
    const errorCloseBtn = document.getElementById('error-close-btn');
    const themeToggle = document.getElementById('theme-toggle');

    // ============================
    //  Theme (Dark/Light) Logic
    // ============================
    function getPreferredTheme() {
        const stored = localStorage.getItem('docconvert-theme');
        if (stored) return stored;
        // Respect OS-level preference
        return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
    }

    function applyTheme(theme) {
        if (theme === 'dark') {
            document.documentElement.setAttribute('data-theme', 'dark');
        } else {
            document.documentElement.removeAttribute('data-theme');
        }
    }

    // Apply immediately on load
    applyTheme(getPreferredTheme());

    themeToggle.addEventListener('click', () => {
        const currentTheme = document.documentElement.getAttribute('data-theme') === 'dark' ? 'dark' : 'light';
        const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
        applyTheme(newTheme);
        localStorage.setItem('docconvert-theme', newTheme);
    });

    // Listen for OS preference changes (e.g., user toggles system dark mode)
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', (e) => {
        // Only auto-switch if user hasn't manually set a preference
        if (!localStorage.getItem('docconvert-theme')) {
            applyTheme(e.matches ? 'dark' : 'light');
        }
    });

    // State
    let currentMode = 'pdf-to-word'; // default
    let activeXhr = null;

    // SVG Icons
    const pdfIcon = `
        <svg class="file-icon pdf" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
            <polyline points="14 2 14 8 20 8"></polyline>
            <line x1="9" y1="15" x2="12" y2="15"></line>
            <line x1="9" y1="11" x2="15" y2="11"></line>
            <line x1="9" y1="19" x2="13" y2="19"></line>
        </svg>
    `;

    const wordIcon = `
        <svg class="file-icon word" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
            <polyline points="14 2 14 8 20 8"></polyline>
            <line x1="16" y1="13" x2="8" y2="13"></line>
            <line x1="16" y1="17" x2="8" y2="17"></line>
            <polyline points="10 9 9 9 8 9"></polyline>
        </svg>
    `;

    // Handle Mode Toggles
    modePdfToWord.addEventListener('click', () => setMode('pdf-to-word'));
    modeWordToPdf.addEventListener('click', () => setMode('word-to-pdf'));

    function setMode(mode) {
        currentMode = mode;
        if (mode === 'pdf-to-word') {
            modePdfToWord.classList.add('active');
            modeWordToPdf.classList.remove('active');
            fileInput.setAttribute('accept', '.pdf');
        } else {
            modeWordToPdf.classList.add('active');
            modePdfToWord.classList.remove('active');
            fileInput.setAttribute('accept', '.docx,.doc');
        }
    }

    // Dropzone Event Listeners
    dropzone.addEventListener('click', () => fileInput.click());

    ['dragenter', 'dragover'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.add('dragover');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.remove('dragover');
        }, false);
    });

    dropzone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        if (files.length > 0) {
            handleFileSelection(files[0]);
        }
    });

    fileInput.addEventListener('change', () => {
        if (fileInput.files.length > 0) {
            handleFileSelection(fileInput.files[0]);
        }
    });

    // Error Close Action
    errorCloseBtn.addEventListener('click', hideError);

    // Reset Action
    resetBtn.addEventListener('click', resetUI);

    // Main File Handler
    function handleFileSelection(file) {
        hideError();
        
        const filename = file.name;
        const fileSizeInBytes = file.size;
        const extension = filename.slice(((filename.lastIndexOf(".") - 1) >>> 0) + 2).toLowerCase();

        // 1. Client-Side Size Validation (25MB Limit)
        const sizeLimit = 25 * 1024 * 1024; // 25MB
        if (fileSizeInBytes > sizeLimit) {
            showError('File is too large. Sane limit is 25MB.');
            fileInput.value = '';
            return;
        }

        // 2. Client-Side Format Validation & Auto Mode Toggle
        if (extension === 'pdf') {
            setMode('pdf-to-word');
            fileIconContainer.innerHTML = pdfIcon;
        } else if (extension === 'docx' || extension === 'doc') {
            setMode('word-to-pdf');
            fileIconContainer.innerHTML = wordIcon;
        } else {
            showError('Unsupported file type. Please upload a PDF, DOCX, or DOC file.');
            fileInput.value = '';
            return;
        }

        // 3. Update File metadata in UI
        fileNameEl.textContent = filename;
        fileSizeEl.textContent = formatBytes(fileSizeInBytes);

        // 4. Update UI states
        dropzone.classList.add('hidden');
        statusPanel.classList.remove('hidden');
        actionButtons.classList.add('hidden');
        progressBar.style.width = '0%';
        progressBar.style.backgroundColor = 'var(--primary-color)';
        progressText.textContent = 'Uploading file...';
        progressPercentage.textContent = '0%';

        // 5. Begin Upload and Conversion
        uploadAndConvert(file);
    }

    // Ajax upload and convert handling
    function uploadAndConvert(file) {
        const formData = new FormData();
        formData.append('file', file);

        activeXhr = new XMLHttpRequest();
        activeXhr.open('POST', '/api/convert', true);

        // Upload progress listener
        activeXhr.upload.addEventListener('progress', (e) => {
            if (e.lengthComputable) {
                const percentComplete = Math.round((e.loaded / e.total) * 100);
                // Reserve 100% upload completion text change
                if (percentComplete < 100) {
                    progressBar.style.width = `${percentComplete}%`;
                    progressText.textContent = `Uploading file...`;
                    progressPercentage.textContent = `${percentComplete}%`;
                } else {
                    progressBar.style.width = '100%';
                    progressText.textContent = 'Converting file server-side... This may take a moment.';
                    progressPercentage.textContent = 'Processing';
                }
            }
        });

        // Request complete listener
        activeXhr.onload = function() {
            if (activeXhr.status === 200) {
                try {
                    const response = JSON.parse(activeXhr.responseText);
                    if (response.success) {
                        // Success State
                        progressBar.style.width = '100%';
                        progressBar.style.backgroundColor = 'var(--success-color)';
                        progressText.textContent = 'Conversion complete!';
                        progressPercentage.textContent = 'Done';
                        
                        downloadBtn.href = response.download_url;
                        downloadBtn.setAttribute('download', response.filename);
                        actionButtons.classList.remove('hidden');
                    } else {
                        handleFailure(response.error || 'Conversion failed.');
                    }
                } catch (err) {
                    handleFailure('Failed to parse server response.');
                }
            } else {
                let errorMsg = 'An error occurred during conversion.';
                try {
                    const response = JSON.parse(activeXhr.responseText);
                    errorMsg = response.error || errorMsg;
                } catch (e) {}
                handleFailure(errorMsg);
            }
        };

        activeXhr.onerror = function() {
            handleFailure('Network error. Could not connect to conversion server.');
        };

        activeXhr.send(formData);
    }

    // Handling Failures
    function handleFailure(message) {
        showError(message);
        resetUI();
    }

    // UI helper functions
    function formatBytes(bytes, decimals = 1) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const dm = decimals < 0 ? 0 : decimals;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
    }

    function showError(message) {
        errorMessage.textContent = message;
        errorBanner.classList.remove('hidden');
    }

    function hideError() {
        errorBanner.classList.add('hidden');
    }

    function resetUI() {
        if (activeXhr) {
            activeXhr.abort();
            activeXhr = null;
        }
        fileInput.value = '';
        dropzone.classList.remove('hidden');
        statusPanel.classList.add('hidden');
        actionButtons.classList.add('hidden');
    }
});
