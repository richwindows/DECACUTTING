// Global variables
let uploadedFile = null;
let processedData = null;

// DOM elements
const fileInput = document.getElementById('fileInput');
const fileUploadArea = document.getElementById('fileUploadArea');
const processBtn = document.getElementById('processBtn');
const statusCard = document.getElementById('statusCard');
const fileInfo = document.getElementById('fileInfo');
const progressBar = document.getElementById('progressBar');
const resultsSection = document.getElementById('resultsSection');
const loadingOverlay = document.getElementById('loadingOverlay');
const processType = document.getElementById('processType');

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
    initializeEventListeners();
});

function initializeEventListeners() {
    // File upload area click
    fileUploadArea.addEventListener('click', () => {
        fileInput.click();
    });

    // File input change
    fileInput.addEventListener('change', handleFileSelect);

    // Drag and drop events
    fileUploadArea.addEventListener('dragover', handleDragOver);
    fileUploadArea.addEventListener('dragleave', handleDragLeave);
    fileUploadArea.addEventListener('drop', handleFileDrop);

    // Process button click
    processBtn.addEventListener('click', processFile);

    // Download buttons
    document.getElementById('downloadExcel').addEventListener('click', () => downloadFile('excel'));
    document.getElementById('downloadCSV').addEventListener('click', () => downloadFile('csv'));
}

function handleFileSelect(event) {
    const file = event.target.files[0];
    if (file) {
        validateAndSetFile(file);
    }
}

function handleDragOver(event) {
    event.preventDefault();
    fileUploadArea.classList.add('dragover');
}

function handleDragLeave(event) {
    event.preventDefault();
    fileUploadArea.classList.remove('dragover');
}

function handleFileDrop(event) {
    event.preventDefault();
    fileUploadArea.classList.remove('dragover');
    
    const files = event.dataTransfer.files;
    if (files.length > 0) {
        validateAndSetFile(files[0]);
    }
}

function validateAndSetFile(file) {
    // Validate file type
    const allowedTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
        'application/vnd.ms-excel', // .xls
        'application/vnd.ms-excel.sheet.macroEnabled.12' // .xlsm
    ];
    
    const fileExtension = file.name.toLowerCase().split('.').pop();
    const allowedExtensions = ['xlsx', 'xls', 'xlsm'];
    
    if (!allowedExtensions.includes(fileExtension)) {
        showError('请选择有效的Excel文件 (.xlsx, .xls, .xlsm)');
        return;
    }

    // Validate file size (max 10MB)
    const maxSize = 10 * 1024 * 1024; // 10MB
    if (file.size > maxSize) {
        showError('文件大小不能超过10MB');
        return;
    }

    uploadedFile = file;
    updateFileStatus();
    enableProcessButton();
}

function updateFileStatus() {
    if (!uploadedFile) return;

    // Update status card
    statusCard.className = 'status-card uploaded';
    statusCard.innerHTML = `
        <div class="status-icon">✅</div>
        <div class="status-title">文件已上传</div>
        <div class="status-message">准备开始处理</div>
    `;

    // Show file info
    document.getElementById('fileName').textContent = uploadedFile.name;
    document.getElementById('fileType').textContent = processType.value;
    document.getElementById('fileSize').textContent = formatFileSize(uploadedFile.size);
    fileInfo.style.display = 'block';
}

function enableProcessButton() {
    processBtn.disabled = false;
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

async function processFile() {
    if (!uploadedFile) {
        showError('请先选择文件');
        return;
    }

    try {
        // Show loading state
        showLoading();
        updateProcessingStatus();

        // Create FormData
        const formData = new FormData();
        formData.append('file', uploadedFile);
        formData.append('processType', processType.value);

        // Send request to API
        const response = await fetch('/api/process', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const result = await response.json();

        if (result.success) {
            processedData = result.data;
            showProcessingSuccess();
            displayResults(result.data);
        } else {
            throw new Error(result.message || '处理失败');
        }

    } catch (error) {
        console.error('Processing error:', error);
        showError(`处理文件时出错: ${error.message}`);
    } finally {
        hideLoading();
    }
}

function showLoading() {
    loadingOverlay.style.display = 'flex';
}

function hideLoading() {
    loadingOverlay.style.display = 'none';
}

function updateProcessingStatus() {
    statusCard.className = 'status-card processing';
    statusCard.innerHTML = `
        <div class="status-icon">🔄</div>
        <div class="status-title">正在处理</div>
        <div class="status-message">智能切割优化进行中...</div>
    `;
    
    progressBar.style.display = 'block';
    progressBar.querySelector('.progress-fill').style.width = '100%';
}

function showProcessingSuccess() {
    statusCard.className = 'status-card success';
    statusCard.innerHTML = `
        <div class="status-icon">🎉</div>
        <div class="status-title">处理完成</div>
        <div class="status-message">智能切割优化成功完成</div>
    `;
    
    progressBar.style.display = 'none';
}

function showError(message) {
    statusCard.className = 'status-card error';
    statusCard.innerHTML = `
        <div class="status-icon">❌</div>
        <div class="status-title">处理失败</div>
        <div class="status-message">${message}</div>
    `;
    
    progressBar.style.display = 'none';
    hideLoading();
}

function displayResults(data) {
    if (!data || !data.rows) {
        showError('无效的处理结果');
        return;
    }

    // Update metrics
    document.getElementById('totalRows').textContent = data.rows.length;
    document.getElementById('materialTypes').textContent = data.materialTypes || 0;
    document.getElementById('cuttingGroups').textContent = data.maxCuttingId || 0;
    document.getElementById('totalLength').textContent = `${(data.totalLength || 0).toFixed(1)}mm`;

    // Display data table
    displayDataTable(data.rows.slice(0, 10)); // Show first 10 rows

    // Show results section
    resultsSection.style.display = 'block';
    
    // Scroll to results
    resultsSection.scrollIntoView({ behavior: 'smooth' });
}

function displayDataTable(rows) {
    if (!rows || rows.length === 0) {
        document.getElementById('dataTable').innerHTML = '<p>没有数据可显示</p>';
        return;
    }

    // Get column headers from first row
    const headers = Object.keys(rows[0]);
    
    let tableHTML = '<table><thead><tr>';
    headers.forEach(header => {
        tableHTML += `<th>${header}</th>`;
    });
    tableHTML += '</tr></thead><tbody>';

    // Add data rows
    rows.forEach(row => {
        tableHTML += '<tr>';
        headers.forEach(header => {
            const value = row[header] !== null && row[header] !== undefined ? row[header] : '';
            tableHTML += `<td>${value}</td>`;
        });
        tableHTML += '</tr>';
    });

    tableHTML += '</tbody></table>';
    document.getElementById('dataTable').innerHTML = tableHTML;
}

async function downloadFile(format) {
    if (!processedData) {
        showError('没有可下载的数据');
        return;
    }

    try {
        const response = await fetch('/api/download', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                data: processedData,
                format: format,
                filename: uploadedFile.name
            })
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        // Get the blob from response
        const blob = await response.blob();
        
        // Create download link
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        
        // Set filename
        const baseFilename = uploadedFile.name.replace(/\.[^/.]+$/, "");
        const extension = format === 'excel' ? 'xlsx' : 'csv';
        a.download = `${baseFilename}_CutFrame.${extension}`;
        
        document.body.appendChild(a);
        a.click();
        
        // Cleanup
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

    } catch (error) {
        console.error('Download error:', error);
        showError(`下载文件时出错: ${error.message}`);
    }
}

// Utility functions
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

// Add some visual feedback for better UX
document.addEventListener('DOMContentLoaded', function() {
    // Add smooth scrolling for anchor links
    document.querySelectorAll('a[href^="#"]').forEach(anchor => {
        anchor.addEventListener('click', function (e) {
            e.preventDefault();
            const target = document.querySelector(this.getAttribute('href'));
            if (target) {
                target.scrollIntoView({
                    behavior: 'smooth',
                    block: 'start'
                });
            }
        });
    });

    // Add intersection observer for animations
    const observerOptions = {
        threshold: 0.1,
        rootMargin: '0px 0px -50px 0px'
    };

    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.style.opacity = '1';
                entry.target.style.transform = 'translateY(0)';
            }
        });
    }, observerOptions);

    // Observe elements for animation
    document.querySelectorAll('.upload-section, .status-section, .results-section').forEach(el => {
        el.style.opacity = '0';
        el.style.transform = 'translateY(20px)';
        el.style.transition = 'opacity 0.6s ease, transform 0.6s ease';
        observer.observe(el);
    });
});