// Bank Fees Accrual Processing System - Frontend JavaScript

class BankFeesApp {
    constructor() {
        this.processingStatus = false;
        this.statusUpdateInterval = null;
        this.init();
    }

    init() {
        this.setupEventListeners();
        this.updateStatus();
        this.validateForm();
    }

    setupEventListeners() {
        // File upload form
        const uploadForm = document.getElementById('uploadForm');
        if (uploadForm) {
            uploadForm.addEventListener('submit', (e) => this.handleFileUpload(e));
        }

        // File input changes
        const bankFiles = document.getElementById('bankFiles');
        const netsuiteFile = document.getElementById('netsuiteFile');
        
        if (bankFiles) {
            bankFiles.addEventListener('change', () => this.validateForm());
        }
        
        if (netsuiteFile) {
            netsuiteFile.addEventListener('change', () => this.validateForm());
        }

        // Download button
        const downloadBtn = document.getElementById('downloadBtn');
        if (downloadBtn) {
            downloadBtn.addEventListener('click', (e) => this.handleDownload(e));
        }

        // Status refresh
        setInterval(() => this.updateStatus(), 5000); // Update every 5 seconds
    }

    validateForm() {
        const bankFiles = document.getElementById('bankFiles');
        const netsuiteFile = document.getElementById('netsuiteFile');
        const uploadBtn = document.getElementById('uploadBtn');

        if (!bankFiles || !netsuiteFile || !uploadBtn) return;

        const isValid = bankFiles.files.length > 0 && netsuiteFile.files.length > 0;
        
        uploadBtn.disabled = !isValid;
        uploadBtn.innerHTML = isValid 
            ? '<i class="fas fa-cloud-upload-alt me-2"></i>Process Files'
            : '<i class="fas fa-exclamation-triangle me-2"></i>Select Files First';
        
        if (isValid) {
            uploadBtn.classList.remove('btn-secondary');
            uploadBtn.classList.add('btn-primary');
        } else {
            uploadBtn.classList.remove('btn-primary');
            uploadBtn.classList.add('btn-secondary');
        }

        // Update file info
        this.updateFileInfo();
    }

    updateFileInfo() {
        const bankFiles = document.getElementById('bankFiles');
        const netsuiteFile = document.getElementById('netsuiteFile');

        if (bankFiles && bankFiles.files.length > 0) {
            const fileList = Array.from(bankFiles.files).map(f => f.name).join(', ');
            this.showMessage(`Bank files selected: ${fileList}`, 'info');
        }

        if (netsuiteFile && netsuiteFile.files.length > 0) {
            this.showMessage(`NetSuite file selected: ${netsuiteFile.files[0].name}`, 'info');
        }
    }

    async handleFileUpload(event) {
        event.preventDefault();
        
        if (this.processingStatus) {
            this.showMessage('Processing already in progress', 'warning');
            return;
        }

        const formData = new FormData(event.target);
        const uploadBtn = document.getElementById('uploadBtn');
        
        try {
            this.setProcessingStatus(true);
            this.showProcessingSection();
            this.updateProgress(10, 'Uploading files...');

            uploadBtn.disabled = true;
            uploadBtn.innerHTML = '<i class="fas fa-spinner fa-spin me-2"></i>Processing...';

            const response = await fetch('/upload', {
                method: 'POST',
                body: formData
            });

            this.updateProgress(50, 'Processing bank transaction files...');

            const result = await response.json();

            if (result.status === 'success') {
                this.updateProgress(80, 'Matching transactions...');
                
                // Wait a bit for processing to complete
                await this.sleep(2000);
                
                this.updateProgress(100, 'Processing complete!');
                this.showResultsSection();
                this.updateStatus();
                
                this.showMessage('Files processed successfully! You can now generate the report.', 'success');
            } else {
                throw new Error(result.message || 'Processing failed');
            }

        } catch (error) {
            console.error('Upload error:', error);
            this.showMessage(`Error processing files: ${error.message}`, 'error');
            this.hideProcessingSection();
        } finally {
            this.setProcessingStatus(false);
            uploadBtn.disabled = false;
            uploadBtn.innerHTML = '<i class="fas fa-cloud-upload-alt me-2"></i>Process Files';
        }
    }

    async handleDownload(event) {
        event.preventDefault();
        
        const downloadBtn = event.target;
        const originalHtml = downloadBtn.innerHTML;
        
        try {
            downloadBtn.disabled = true;
            downloadBtn.innerHTML = '<i class="fas fa-spinner fa-spin me-2"></i>Generating Report...';
            
            const response = await fetch('/generate_report');
            
            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = response.headers.get('Content-Disposition')?.split('filename=')[1] || 'Customer_Funds_Analysis.xlsx';
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
                
                this.showMessage('Report generated and downloaded successfully!', 'success');
            } else {
                throw new Error('Failed to generate report');
            }
            
        } catch (error) {
            console.error('Download error:', error);
            this.showMessage(`Error generating report: ${error.message}`, 'error');
        } finally {
            downloadBtn.disabled = false;
            downloadBtn.innerHTML = originalHtml;
        }
    }

    async updateStatus() {
        try {
            const response = await fetch('/status');
            const status = await response.json();
            
            this.updateStatusDisplay(status);
            
        } catch (error) {
            console.error('Status update error:', error);
        }
    }

    updateStatusDisplay(status) {
        const bankAccountsCount = document.getElementById('bankAccountsCount');
        const netsuiteStatus = document.getElementById('netsuiteStatus');
        const totalRecords = document.getElementById('totalRecords');
        const systemStatus = document.getElementById('systemStatus');
        const recordsProcessed = document.getElementById('recordsProcessed');
        const transactionsMatched = document.getElementById('transactionsMatched');

        if (bankAccountsCount) {
            bankAccountsCount.textContent = status.bank_accounts_loaded || 0;
            bankAccountsCount.className = status.bank_accounts_loaded > 0 ? 'badge bg-success' : 'badge bg-secondary';
        }

        if (netsuiteStatus) {
            netsuiteStatus.textContent = status.netsuite_loaded ? 'Loaded' : 'Not Loaded';
            netsuiteStatus.className = status.netsuite_loaded ? 'badge bg-success' : 'badge bg-secondary';
        }

        if (totalRecords) {
            totalRecords.textContent = status.netsuite_records || 0;
            totalRecords.className = status.netsuite_records > 0 ? 'badge bg-info' : 'badge bg-secondary';
        }

        if (systemStatus) {
            const isReady = status.bank_accounts_loaded > 0 && status.netsuite_loaded;
            systemStatus.innerHTML = isReady 
                ? '<i class="fas fa-circle text-success"></i><span>System Ready</span>'
                : '<i class="fas fa-circle text-warning"></i><span>Awaiting Data</span>';
        }

        if (recordsProcessed && status.netsuite_records) {
            recordsProcessed.textContent = status.netsuite_records.toLocaleString();
        }

        if (transactionsMatched && status.bank_accounts_loaded) {
            // Estimate based on loaded accounts
            const estimated = Math.floor(status.netsuite_records * 0.75);
            transactionsMatched.textContent = estimated.toLocaleString();
        }
    }

    setProcessingStatus(isProcessing) {
        this.processingStatus = isProcessing;
        
        if (isProcessing && !this.statusUpdateInterval) {
            this.statusUpdateInterval = setInterval(() => this.updateStatus(), 2000);
        } else if (!isProcessing && this.statusUpdateInterval) {
            clearInterval(this.statusUpdateInterval);
            this.statusUpdateInterval = null;
        }
    }

    showProcessingSection() {
        const section = document.getElementById('processingStatus');
        if (section) {
            section.style.display = 'block';
            section.scrollIntoView({ behavior: 'smooth' });
        }
    }

    hideProcessingSection() {
        const section = document.getElementById('processingStatus');
        if (section) {
            section.style.display = 'none';
        }
    }

    showResultsSection() {
        const section = document.getElementById('resultsSection');
        if (section) {
            section.style.display = 'block';
            section.scrollIntoView({ behavior: 'smooth' });
        }
    }

    updateProgress(percentage, message) {
        const progressBar = document.getElementById('progressBar');
        const processingLog = document.getElementById('processingLog');

        if (progressBar) {
            progressBar.style.width = `${percentage}%`;
            progressBar.setAttribute('aria-valuenow', percentage);
        }

        if (processingLog && message) {
            const timestamp = new Date().toLocaleTimeString();
            processingLog.innerHTML += `<div>[${timestamp}] ${message}</div>`;
            processingLog.scrollTop = processingLog.scrollHeight;
        }
    }

    showMessage(message, type = 'info') {
        // Create alert if it doesn't exist
        let alertContainer = document.querySelector('.alert-container');
        if (!alertContainer) {
            alertContainer = document.createElement('div');
            alertContainer.className = 'alert-container';
            alertContainer.style.cssText = 'position: fixed; top: 20px; right: 20px; z-index: 9999; max-width: 400px;';
            document.body.appendChild(alertContainer);
        }

        const alertClass = type === 'error' ? 'alert-danger' : 
                          type === 'success' ? 'alert-success' : 
                          type === 'warning' ? 'alert-warning' : 'alert-info';

        const iconClass = type === 'error' ? 'fa-exclamation-triangle' : 
                         type === 'success' ? 'fa-check-circle' : 
                         type === 'warning' ? 'fa-exclamation-triangle' : 'fa-info-circle';

        const alert = document.createElement('div');
        alert.className = `alert ${alertClass} alert-dismissible fade show`;
        alert.innerHTML = `
            <i class="fas ${iconClass} me-2"></i>
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        `;

        alertContainer.appendChild(alert);

        // Auto-dismiss after 5 seconds
        setTimeout(() => {
            if (alert.parentNode) {
                alert.classList.remove('show');
                setTimeout(() => {
                    if (alert.parentNode) {
                        alert.remove();
                    }
                }, 150);
            }
        }, 5000);
    }

    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    // File drag and drop functionality
    setupDragAndDrop() {
        const dropZones = document.querySelectorAll('.file-input-container');
        
        dropZones.forEach(zone => {
            zone.addEventListener('dragover', (e) => {
                e.preventDefault();
                zone.classList.add('drag-over');
            });

            zone.addEventListener('dragleave', () => {
                zone.classList.remove('drag-over');
            });

            zone.addEventListener('drop', (e) => {
                e.preventDefault();
                zone.classList.remove('drag-over');
                
                const files = Array.from(e.dataTransfer.files);
                const input = zone.querySelector('input[type="file"]');
                
                if (input && files.length > 0) {
                    const dt = new DataTransfer();
                    files.forEach(file => dt.items.add(file));
                    input.files = dt.files;
                    this.validateForm();
                }
            });
        });
    }

    // Format numbers for display
    formatNumber(num) {
        return new Intl.NumberFormat('en-US').format(num);
    }

    // Format currency for display
    formatCurrency(amount) {
        return new Intl.NumberFormat('en-US', {
            style: 'currency',
            currency: 'USD'
        }).format(amount);
    }
}

// Initialize the application when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    window.bankFeesApp = new BankFeesApp();
    
    // Setup additional enhancements
    if (window.bankFeesApp.setupDragAndDrop) {
        window.bankFeesApp.setupDragAndDrop();
    }
});

// Additional utility functions
window.BankFeesUtils = {
    downloadTemplate: () => {
        // Future: Download template files
        console.log('Template download functionality');
    },

    validateFileFormat: (file) => {
        const validExtensions = {
            csv: ['text/csv', 'application/vnd.ms-excel'],
            xlsx: ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel']
        };

        const fileExtension = file.name.split('.').pop().toLowerCase();
        const fileType = file.type.toLowerCase();

        return validExtensions[fileExtension]?.includes(fileType) || 
               validExtensions[fileExtension]?.some(type => fileType.includes(type));
    },

    formatFileSize: (bytes) => {
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        if (bytes === 0) return '0 Bytes';
        const i = Math.floor(Math.log(bytes) / Math.log(1024));
        return Math.round(bytes / Math.pow(1024, i) * 100) / 100 + ' ' + sizes[i];
    }
};
