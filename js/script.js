 // Main application object
        const app = {
            // Initialize data with localStorage persistence
            data: {
                headers: ['Name', 'Email', 'Phone'],
                rows: [],
                fileName: 'Untitled.xlsx',
                lastModified: null,
                settings: {
                    theme: 'default',
                    autoSave: 5,
                    rowsPerPage: 10
                },
                fileHistory: []
            },

            // DOM elements
            elements: {
                formFields: document.getElementById('formFields'),
                dataForm: document.getElementById('dataForm'),
                tableHeader: document.getElementById('tableHeader'),
                tableBody: document.getElementById('tableBody'),
                editIndexInput: document.getElementById('editIndex'),
                newColumnName: document.getElementById('newColumnName'),
                addColumnBtn: document.getElementById('addColumnBtn'),
                removeColumnBtn: document.getElementById('removeColumnBtn'),
                resetColumnsBtn: document.getElementById('resetColumnsBtn'),
                clearFormBtn: document.getElementById('clearForm'),
                downloadExcelBtn: document.getElementById('downloadExcel'),
                printDataBtn: document.getElementById('printData'),
                loadSampleBtn: document.getElementById('loadSample'),
                clearAllDataBtn: document.getElementById('clearAllData'),
                fileInfo: document.getElementById('fileInfo'),
                entryFileName: document.getElementById('entryFileName'),
                saveModal: document.getElementById('saveModal'),
                closeModal: document.querySelectorAll('.close-modal'),
                cancelSave: document.getElementById('cancelSave'),
                confirmSave: document.getElementById('confirmSave'),
                saveFileName: document.getElementById('saveFileName'),
                saveFileFormat: document.getElementById('saveFileFormat'),
                navLinks: document.querySelectorAll('.nav-link'),
                tabContents: document.querySelectorAll('.tab-content'),
                noDataMessage: document.getElementById('noDataMessage'),
                dataTable: document.getElementById('dataTable'),
                analyticsColumn: document.getElementById('analyticsColumn'),
                reportColumnsCheckboxes: document.getElementById('reportColumnsCheckboxes'),
                generatePdf: document.getElementById('generatePdf'),
                previewReport: document.getElementById('previewReport'),
                appTheme: document.getElementById('appTheme'),
                autoSave: document.getElementById('autoSave'),
                rowsPerPage: document.getElementById('rowsPerPage'),
                saveSettings: document.getElementById('saveSettings'),
                resetSettings: document.getElementById('resetSettings'),
                exportSettings: document.getElementById('exportSettings'),
                importSettings: document.getElementById('importSettings'),
                importSettingsModal: document.getElementById('importSettingsModal'),
                settingsFile: document.getElementById('settingsFile'),
                confirmImport: document.getElementById('confirmImport'),
                cancelImport: document.getElementById('cancelImport'),
                importSpinner: document.getElementById('importSpinner'),
                toast: document.getElementById('toast'),
                toastMessage: document.getElementById('toastMessage'),
                mobileMenuBtn: document.querySelector('.mobile-menu-btn'),
                sidebar: document.querySelector('.sidebar'),
                userProfile: document.querySelector('.user-profile'),
                userMenu: document.querySelector('.user-menu'),
                refreshHistory: document.getElementById('refreshHistory'),
                clearHistory: document.getElementById('clearHistory'),
                fileHistory: document.getElementById('fileHistory'),
                noHistoryMessage: document.getElementById('noHistoryMessage'),
                mobileDotsMenu: document.querySelector('.mobile-dots-menu'),
                mobileDotsMenuContent: document.querySelector('.mobile-dots-menu-content'),
                mobileDotsMenuItems: document.querySelectorAll('.mobile-dots-menu-item')
            },

            // Chart instances
            charts: {
                dataChart: null,
                columnChart: null,
                valueDistributionChart: null
            },

            // Initialize the application
            init() {
                this.loadData();
                this.setupEventListeners();
                this.initializeForm();
                this.updateUI();
                this.renderCharts();
                this.updateTheme();
                this.setupUserMenu();
                this.setupMobileMenu();
                this.updateFileHistory();
            },

            // Load data from localStorage if available
            loadData() {
                const savedData = localStorage.getItem('excelData');
                if (savedData) {
                    this.data = JSON.parse(savedData);
                    this.updateUI();
                    this.showToast('Data loaded successfully', 'success');
                }
            },

            // Save data to localStorage
            saveData() {
                this.data.lastModified = new Date().toLocaleString();
                localStorage.setItem('excelData', JSON.stringify(this.data));
                this.updateUI();
                this.showToast('Data saved successfully', 'success');
            },

            // Initialize the form with current columns
            initializeForm() {
                this.elements.formFields.innerHTML = '';
                this.data.headers.forEach(header => {
                    this.elements.formFields.innerHTML += `
                        <div class="form-group">
                            <label for="${header.toLowerCase().replace(/\s+/g, '_')}">${header}:</label>
                            <input type="text" id="${header.toLowerCase().replace(/\s+/g, '_')}" name="${header.toLowerCase().replace(/\s+/g, '_')}" required>
                        </div>
                    `;
                });
            },

            // Update the entire UI
            updateUI() {
                this.updateTable();
                this.updateFileInfo();
                this.updateDashboardCards();
                this.updateAnalyticsColumnSelect();
                this.updateReportColumnsCheckboxes();
                this.renderCharts();
            },

            // Update the data table with SNo column
            updateTable() {
                // Show/hide no data message
                if (this.data.rows.length === 0) {
                    this.elements.noDataMessage.style.display = 'block';
                    this.elements.dataTable.style.display = 'none';
                    return;
                } else {
                    this.elements.noDataMessage.style.display = 'none';
                    this.elements.dataTable.style.display = 'table';
                }

                // Update table header with SNo column
                this.elements.tableHeader.innerHTML = '<th class="sno-column">SNo</th>';
                this.data.headers.forEach(header => {
                    this.elements.tableHeader.innerHTML += `<th>${header}</th>`;
                });
                this.elements.tableHeader.innerHTML += '<th class="no-print">Actions</th>';
                
                // Update table body with SNo column
                this.elements.tableBody.innerHTML = '';
                this.data.rows.forEach((row, index) => {
                    const tr = document.createElement('tr');
                    
                    // Add SNo column
                    const snoTd = document.createElement('td');
                    snoTd.className = 'sno-column';
                    snoTd.textContent = index + 1;
                    tr.appendChild(snoTd);
                    
                    // Add data columns
                    this.data.headers.forEach(header => {
                        const td = document.createElement('td');
                        td.textContent = row[header] || '';
                        tr.appendChild(td);
                    });
                    
                    // Add action buttons
                    const actionTd = document.createElement('td');
                    actionTd.className = 'action-buttons no-print';
                    actionTd.innerHTML = `
                        <button class="btn btn-primary action-btn" onclick="app.editRow(${index})">
                            <i class="fas fa-edit"></i> Edit
                        </button>
                        <button class="btn btn-danger action-btn" onclick="app.deleteRow(${index})">
                            <i class="fas fa-trash"></i> Delete
                        </button>
                    `;
                    tr.appendChild(actionTd);
                    
                    this.elements.tableBody.appendChild(tr);
                });
            },

            // Update file information display
            updateFileInfo() {
                this.elements.fileInfo.textContent = `Current File: ${this.data.fileName} | Records: ${this.data.rows.length} | Last Modified: ${this.data.lastModified || 'Not saved yet'}`;
            },

            // Update dashboard cards
            updateDashboardCards() {
                document.getElementById('totalRecordsCard').textContent = this.data.rows.length;
                document.getElementById('totalColumnsCard').textContent = this.data.headers.length;
                document.getElementById('lastModifiedCard').textContent = this.data.lastModified || 'Never';
                document.getElementById('currentFileCard').textContent = this.data.fileName;
            },

            // Update analytics column select
            updateAnalyticsColumnSelect() {
                this.elements.analyticsColumn.innerHTML = '';
                this.data.headers.forEach(header => {
                    this.elements.analyticsColumn.innerHTML += `<option value="${header}">${header}</option>`;
                });
            },

            // Update report columns checkboxes
            updateReportColumnsCheckboxes() {
                this.elements.reportColumnsCheckboxes.innerHTML = '';
                this.data.headers.forEach(header => {
                    this.elements.reportColumnsCheckboxes.innerHTML += `
                        <div style="margin-bottom: 5px;">
                            <input type="checkbox" id="report-col-${header.toLowerCase().replace(/\s+/g, '-')}" 
                                   name="reportColumns" value="${header}" checked>
                            <label for="report-col-${header.toLowerCase().replace(/\s+/g, '-')}">${header}</label>
                        </div>
                    `;
                });
            },

            // Update file history display
            updateFileHistory() {
                if (this.data.fileHistory.length === 0) {
                    this.elements.noHistoryMessage.style.display = 'block';
                    this.elements.fileHistory.style.display = 'none';
                    return;
                }
                
                this.elements.noHistoryMessage.style.display = 'none';
                this.elements.fileHistory.style.display = 'grid';
                this.elements.fileHistory.innerHTML = '';
                
                // Sort history by date (newest first)
                const sortedHistory = [...this.data.fileHistory].sort((a, b) => 
                    new Date(b.lastModified) - new Date(a.lastModified)
                );
                
                sortedHistory.forEach((file, index) => {
                    const fileCard = document.createElement('div');
                    fileCard.className = 'file-card';
                    fileCard.innerHTML = `
                        <div class="file-card-header">
                            <div class="file-card-title">${file.fileName}</div>
                            <div class="file-card-date">${file.lastModified}</div>
                        </div>
                        <div class="file-card-stats">
                            <div class="file-card-stat">
                                <i class="fas fa-database"></i> ${file.rows.length} records
                            </div>
                            <div class="file-card-stat">
                                <i class="fas fa-columns"></i> ${file.headers.length} columns
                            </div>
                        </div>
                        <div class="file-card-actions">
                            <button class="btn btn-primary file-card-btn" onclick="app.loadFromHistory(${index})">
                                <i class="fas fa-file-import"></i> Load
                            </button>
                            <button class="btn btn-success file-card-btn" onclick="app.downloadFromHistory(${index})">
                                <i class="fas fa-download"></i> Download
                            </button>
                            <button class="btn btn-danger file-card-btn" onclick="app.removeFromHistory(${index})">
                                <i class="fas fa-trash"></i> Remove
                            </button>
                        </div>
                    `;
                    this.elements.fileHistory.appendChild(fileCard);
                });
            },

            // Render charts
            renderCharts() {
                // Data distribution chart
                const ctx1 = document.getElementById('dataChart').getContext('2d');
                if (this.charts.dataChart) this.charts.dataChart.destroy();
                this.charts.dataChart = new Chart(ctx1, {
                    type: 'bar',
                    data: {
                        labels: this.data.headers,
                        datasets: [{
                            label: 'Data Distribution',
                            data: this.data.headers.map(header => {
                                return this.data.rows.filter(row => row[header] && row[header].trim() !== '').length;
                            }),
                            backgroundColor: 'rgba(67, 97, 238, 0.7)',
                            borderColor: 'rgba(67, 97, 238, 1)',
                            borderWidth: 1
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        scales: {
                            y: {
                                beginAtZero: true
                            }
                        },
                        plugins: {
                            legend: {
                                position: 'top',
                            },
                            tooltip: {
                                callbacks: {
                                    label: function(context) {
                                        return `${context.dataset.label}: ${context.raw}`;
                                    }
                                }
                            }
                        }
                    }
                });

                // Column chart
                const ctx2 = document.getElementById('columnChart').getContext('2d');
                if (this.charts.columnChart) this.charts.columnChart.destroy();
                if (this.data.headers.length > 0 && this.data.rows.length > 0) {
                    const selectedColumn = this.elements.analyticsColumn.value || this.data.headers[0];
                    const columnData = this.data.rows.map(row => row[selectedColumn]);
                    
                    // Try to convert to numbers if possible
                    const numericData = columnData.map(item => {
                        const num = parseFloat(item);
                        return isNaN(num) ? 0 : num;
                    });
                    
                    this.charts.columnChart = new Chart(ctx2, {
                        type: 'line',
                        data: {
                            labels: this.data.rows.map((_, i) => `Record ${i + 1}`),
                            datasets: [{
                                label: selectedColumn,
                                data: numericData,
                                borderColor: 'rgba(67, 97, 238, 1)',
                                backgroundColor: 'rgba(67, 97, 238, 0.1)',
                                borderWidth: 2,
                                fill: true,
                                tension: 0.4
                            }]
                        },
                        options: {
                            responsive: true,
                            maintainAspectRatio: false,
                            interaction: {
                                mode: 'index',
                                intersect: false
                            },
                            scales: {
                                y: {
                                    beginAtZero: false
                                }
                            },
                            plugins: {
                                tooltip: {
                                    callbacks: {
                                        label: function(context) {
                                            return `${selectedColumn}: ${columnData[context.dataIndex]}`;
                                        }
                                    }
                                }
                            }
                        }
                    });
                }

                // Value distribution chart
                const ctx3 = document.getElementById('valueDistributionChart').getContext('2d');
                if (this.charts.valueDistributionChart) this.charts.valueDistributionChart.destroy();
                if (this.data.headers.length > 0 && this.data.rows.length > 0) {
                    this.charts.valueDistributionChart = new Chart(ctx3, {
                        type: 'pie',
                        data: {
                            labels: this.data.headers,
                            datasets: [{
                                data: this.data.headers.map(header => {
                                    return this.data.rows.filter(row => row[header] && row[header].trim() !== '').length;
                                }),
                                backgroundColor: this.data.headers.map((_, i) => 
                                    `hsl(${i * 360 / this.data.headers.length}, 70%, 50%)`
                                ),
                                borderWidth: 1
                            }]
                        },
                        options: {
                            responsive: true,
                            maintainAspectRatio: false,
                            plugins: {
                                legend: {
                                    position: 'right',
                                },
                                tooltip: {
                                    callbacks: {
                                        label: function(context) {
                                            const label = context.label || '';
                                            const value = context.raw || 0;
                                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                            const percentage = Math.round((value / total) * 100);
                                            return `${label}: ${value} (${percentage}%)`;
                                        }
                                    }
                                }
                            }
                        }
                    });
                }
            },

            // Update theme based on settings
            updateTheme() {
                const theme = this.data.settings.theme;
                document.documentElement.style.setProperty('--primary-color', this.getThemeColor(theme, 'primary'));
                document.documentElement.style.setProperty('--secondary-color', this.getThemeColor(theme, 'secondary'));
                document.documentElement.style.setProperty('--accent-color', this.getThemeColor(theme, 'accent'));
                document.documentElement.style.setProperty('--danger-color', this.getThemeColor(theme, 'danger'));
                document.documentElement.style.setProperty('--success-color', this.getThemeColor(theme, 'success'));
                document.documentElement.style.setProperty('--warning-color', this.getThemeColor(theme, 'warning'));
                
                // Update select element
                this.elements.appTheme.value = theme;
                this.elements.autoSave.value = this.data.settings.autoSave;
                this.elements.rowsPerPage.value = this.data.settings.rowsPerPage;
            },

            // Get theme colors
            getThemeColor(theme, type) {
                const themes = {
                    default: {
                        primary: '#4361ee',
                        secondary: '#3f37c9',
                        accent: '#4895ef',
                        danger: '#f72585',
                        success: '#4cc9f0',
                        warning: '#f8961e'
                    },
                    dark: {
                        primary: '#212529',
                        secondary: '#343a40',
                        accent: '#495057',
                        danger: '#dc3545',
                        success: '#28a745',
                        warning: '#ffc107'
                    },
                    green: {
                        primary: '#2e7d32',
                        secondary: '#1b5e20',
                        accent: '#4caf50',
                        danger: '#c62828',
                        success: '#2e7d32',
                        warning: '#f9a825'
                    },
                    purple: {
                        primary: '#6a1b9a',
                        secondary: '#4a148c',
                        accent: '#9c27b0',
                        danger: '#ad1457',
                        success: '#7b1fa2',
                        warning: '#ab47bc'
                    },
                    red: {
                        primary: '#c62828',
                        secondary: '#b71c1c',
                        accent: '#d32f2f',
                        danger: '#c62828',
                        success: '#2e7d32',
                        warning: '#f9a825'
                    }
                };
                
                return themes[theme][type];
            },

            // Show toast notification
            showToast(message, type = 'success') {
                this.elements.toastMessage.textContent = message;
                this.elements.toast.className = `toast ${type} show`;
                
                // Set icon based on type
                const icon = this.elements.toast.querySelector('i');
                if (type === 'success') {
                    icon.className = 'fas fa-check-circle';
                } else if (type === 'error') {
                    icon.className = 'fas fa-exclamation-circle';
                } else if (type === 'warning') {
                    icon.className = 'fas fa-exclamation-triangle';
                }
                
                // Hide after 3 seconds
                setTimeout(() => {
                    this.elements.toast.className = 'toast';
                }, 3000);
            },

            // Setup user menu toggle
            setupUserMenu() {
                this.elements.userProfile.addEventListener('click', (e) => {
                    e.stopPropagation();
                    this.elements.userMenu.classList.toggle('show');
                });
                
                // Close menu when clicking outside
                document.addEventListener('click', () => {
                    this.elements.userMenu.classList.remove('show');
                });
            },

            // Setup mobile menu toggle
            setupMobileMenu() {
                this.elements.mobileMenuBtn.addEventListener('click', () => {
                    this.elements.sidebar.classList.toggle('mobile-show');
                });
                
                // Close menu when clicking a link
                document.querySelectorAll('.nav-link').forEach(link => {
                    link.addEventListener('click', () => {
                        this.elements.sidebar.classList.remove('mobile-show');
                    });
                });
                
                // Close menu when clicking outside
                document.addEventListener('click', (e) => {
                    if (!this.elements.sidebar.contains(e.target) && !this.elements.mobileMenuBtn.contains(e.target)) {
                        this.elements.sidebar.classList.remove('mobile-show');
                    }
                });

                // Mobile dots menu
                this.elements.mobileDotsMenu.addEventListener('click', (e) => {
                    e.stopPropagation();
                    this.elements.mobileDotsMenu.classList.toggle('active');
                    this.elements.mobileDotsMenuContent.classList.toggle('show');
                });

                // Close dots menu when clicking outside
                document.addEventListener('click', () => {
                    this.elements.mobileDotsMenu.classList.remove('active');
                    this.elements.mobileDotsMenuContent.classList.remove('show');
                });

                // Handle dots menu item clicks
                this.elements.mobileDotsMenuItems.forEach(item => {
                    item.addEventListener('click', (e) => {
                        e.preventDefault();
                        const tabId = item.getAttribute('data-tab');
                        
                        // Update active tab
                        this.elements.navLinks.forEach(nav => nav.classList.remove('active'));
                        document.querySelector(`[data-tab="${tabId}"]`).classList.add('active');
                        
                        // Show corresponding content
                        this.elements.tabContents.forEach(content => content.classList.remove('active'));
                        document.getElementById(tabId).classList.add('active');
                        
                        // Close menu
                        this.elements.mobileDotsMenu.classList.remove('active');
                        this.elements.mobileDotsMenuContent.classList.remove('show');
                        
                        // Refresh charts if needed
                        if (tabId === 'analytics') {
                            this.renderCharts();
                        }
                    });
                });
            },

            // Add current file to history
            addToHistory() {
                // Create a copy of the current data to store in history
                const historyEntry = {
                    fileName: this.data.fileName,
                    headers: [...this.data.headers],
                    rows: JSON.parse(JSON.stringify(this.data.rows)),
                    lastModified: new Date().toLocaleString()
                };
                
                // Check if this file already exists in history
                const existingIndex = this.data.fileHistory.findIndex(
                    file => file.fileName === historyEntry.fileName
                );
                
                if (existingIndex >= 0) {
                    // Update existing entry
                    this.data.fileHistory[existingIndex] = historyEntry;
                } else {
                    // Add new entry
                    this.data.fileHistory.push(historyEntry);
                }
                
                this.saveData();
                this.updateFileHistory();
            },

            // Setup event listeners
            setupEventListeners() {
                // Navigation tabs
                this.elements.navLinks.forEach(link => {
                    link.addEventListener('click', (e) => {
                        e.preventDefault();
                        const tabId = link.getAttribute('data-tab');
                        
                        // Update active tab
                        this.elements.navLinks.forEach(nav => nav.classList.remove('active'));
                        link.classList.add('active');
                        
                        // Show corresponding content
                        this.elements.tabContents.forEach(content => content.classList.remove('active'));
                        document.getElementById(tabId).classList.add('active');
                        
                        // Refresh charts when analytics tab is opened
                        if (tabId === 'analytics') {
                            this.renderCharts();
                        }
                        
                        // Refresh file history when that tab is opened
                        if (tabId === 'file-history') {
                            this.updateFileHistory();
                        }
                    });
                });

                // Add a new column
                this.elements.addColumnBtn.addEventListener('click', () => {
                    const columnName = this.elements.newColumnName.value.trim();
                    if (columnName && !this.data.headers.includes(columnName)) {
                        this.data.headers.push(columnName);
                        this.elements.newColumnName.value = '';
                        this.initializeForm();
                        this.saveData();
                        this.showToast(`Column "${columnName}" added successfully`, 'success');
                    } else {
                        this.showToast('Please enter a valid and unique column name', 'error');
                    }
                });

                // Remove last column
                this.elements.removeColumnBtn.addEventListener('click', () => {
                    if (this.data.headers.length > 1) {
                        const removedColumn = this.data.headers.pop();
                        // Also remove this column data from all rows
                        this.data.rows.forEach(row => {
                            delete row[removedColumn];
                        });
                        this.initializeForm();
                        this.saveData();
                        this.showToast(`Column "${removedColumn}" removed`, 'warning');
                    } else {
                        this.showToast('You must have at least one column', 'error');
                    }
                });

                // Reset columns to default
                this.elements.resetColumnsBtn.addEventListener('click', () => {
                    if (confirm('This will reset all columns to default (Name, Email, Phone). Continue?')) {
                        this.data.headers = ['Name', 'Email', 'Phone'];
                        // Filter rows to only keep default columns
                        this.data.rows = this.data.rows.map(row => {
                            const newRow = {};
                            this.data.headers.forEach(header => {
                                newRow[header] = row[header] || '';
                            });
                            return newRow;
                        });
                        this.initializeForm();
                        this.saveData();
                        this.showToast('Columns reset to default', 'success');
                    }
                });

                // Form submission
                this.elements.dataForm.addEventListener('submit', (e) => {
                    e.preventDefault();
                    
                    // Update filename if changed
                    const newFileName = this.elements.entryFileName.value.trim() || 'Untitled';
                    this.data.fileName = newFileName.endsWith('.xlsx') ? newFileName : newFileName + '.xlsx';
                    
                    const formData = new FormData(this.elements.dataForm);
                    const rowData = {};
                    this.data.headers.forEach(header => {
                        rowData[header] = formData.get(header.toLowerCase().replace(/\s+/g, '_')) || '';
                    });
                    
                    const editIndex = parseInt(this.elements.editIndexInput.value);
                    if (editIndex >= 0) {
                        // Update existing row
                        this.data.rows[editIndex] = rowData;
                        this.elements.editIndexInput.value = "-1";
                        this.showToast('Record updated successfully', 'success');
                    } else {
                        // Add new row
                        this.data.rows.push(rowData);
                        this.showToast('New record added successfully', 'success');
                    }
                    
                    this.saveData();
                    this.elements.dataForm.reset();
                    
                    // Switch to data view tab
                    document.querySelector('[data-tab="data-view"]').click();
                });

                // Clear form
                this.elements.clearFormBtn.addEventListener('click', () => {
                    this.elements.dataForm.reset();
                    this.elements.editIndexInput.value = "-1";
                    this.showToast('Form cleared', 'info');
                });

                // Clear all data
                this.elements.clearAllDataBtn.addEventListener('click', () => {
                    if (confirm('Are you sure you want to clear ALL data? This cannot be undone.')) {
                        this.data.rows = [];
                        this.saveData();
                        this.showToast('All data cleared', 'warning');
                    }
                });

                // Download Excel
                this.elements.downloadExcelBtn.addEventListener('click', () => {
                    if (this.data.rows.length === 0) {
                        this.showToast('No data to download', 'error');
                        return;
                    }
                    
                    // Show save modal
                    this.elements.saveFileName.value = this.data.fileName.replace('.xlsx', '');
                    this.elements.saveModal.style.display = 'flex';
                });

                // Print data
                this.elements.printDataBtn.addEventListener('click', () => {
                    if (this.data.rows.length === 0) {
                        this.showToast('No data to print', 'error');
                        return;
                    }
                    
                    const reportTitle = document.getElementById('reportTitle').value || 'Data Report';
                    const reportDescription = document.getElementById('reportDescription').value || '';
                    
                    // Get selected columns for printing
                    const selectedColumns = [];
                    document.querySelectorAll('input[name="reportColumns"]:checked').forEach(checkbox => {
                        selectedColumns.push(checkbox.value);
                    });
                    
                    if (selectedColumns.length === 0) {
                        this.showToast('Please select at least one column to print', 'error');
                        return;
                    }
                    
                    // Create print content with SNo column
                    const printContent = `
                        <style>
                            @page {
                                size: A4;
                                margin: 1cm;
                            }
                            body {
                                font-family: Arial, sans-serif;
                                padding: 20px;
                            }
                            h1 {
                                text-align: center;
                                margin-bottom: 20px;
                            }
                            table {
                                width: 100%;
                                border-collapse: collapse;
                                margin-top: 20px;
                            }
                            th, td {
                                border: 1px solid #ddd;
                                padding: 8px;
                                text-align: left;
                            }
                            th {
                                background-color: #f2f2f2;
                            }
                            .sno-column {
                                width: 50px;
                                text-align: center;
                            }
                            .print-footer {
                                margin-top: 50px;
                                padding-top: 20px;
                                border-top: 1px solid #ddd;
                            }
                            .signature-line {
                                display: flex;
                                justify-content: space-between;
                                margin-top: 50px;
                            }
                            .signature {
                                width: 200px;
                                border-top: 1px solid #000;
                                text-align: center;
                                padding-top: 10px;
                            }
                        </style>
                        <h1>${reportTitle}</h1>
                        ${reportDescription ? `<p style="text-align: center; margin-bottom: 30px;">${reportDescription}</p>` : ''}
                        <table border="1">
                            <thead>
                                <tr>
                                    <th class="sno-column">SNo</th>
                                    ${selectedColumns.map(col => `<th>${col}</th>`).join('')}
                                </tr>
                            </thead>
                            <tbody>
                                ${this.data.rows.map((row, index) => `
                                    <tr>
                                        <td class="sno-column">${index + 1}</td>
                                        ${selectedColumns.map(col => `<td>${row[col] || ''}</td>`).join('')}
                                    </tr>
                                `).join('')}
                            </tbody>
                        </table>
                        <div class="print-footer">
                            <p style="text-align: right; font-size: 12px;">
                                Generated on ${new Date().toLocaleString()}
                            </p>
                            <div class="signature-line">
                                <div class="signature">Prepared By: ___________________</div>
                                <div class="signature">Approved By: ___________________</div>
                            </div>
                        </div>
                    `;
                    
                    const printWindow = window.open('', '_blank');
                    printWindow.document.write(printContent);
                    printWindow.document.close();
                    printWindow.focus();
                    
                    // Wait for content to load before printing
                    setTimeout(() => {
                        printWindow.print();
                        printWindow.close();
                    }, 500);
                });

                // Load sample data
                this.elements.loadSampleBtn.addEventListener('click', () => {
                    if (confirm('This will replace your current data with sample data. Continue?')) {
                        this.data.headers = ['Name', 'Email', 'Phone', 'Department', 'Salary'];
                        this.data.rows = [
                            { Name: 'John Doe', Email: 'john@example.com', Phone: '1234567890', Department: 'Sales', Salary: '75000' },
                            { Name: 'Jane Smith', Email: 'jane@example.com', Phone: '9876543210', Department: 'Marketing', Salary: '82000' },
                            { Name: 'Bob Johnson', Email: 'bob@example.com', Phone: '5551234567', Department: 'IT', Salary: '95000' },
                            { Name: 'Alice Williams', Email: 'alice@example.com', Phone: '4445678901', Department: 'HR', Salary: '68000' },
                            { Name: 'Charlie Brown', Email: 'charlie@example.com', Phone: '3337890123', Department: 'Finance', Salary: '105000' }
                        ];
                        this.data.fileName = 'Sample_Data.xlsx';
                        this.saveData();
                        this.showToast('Sample data loaded successfully', 'success');
                    }
                });

                // Modal controls
                this.elements.closeModal.forEach(btn => {
                    btn.addEventListener('click', () => {
                        this.elements.saveModal.style.display = 'none';
                        this.elements.importSettingsModal.style.display = 'none';
                    });
                });

                this.elements.cancelSave.addEventListener('click', () => {
                    this.elements.saveModal.style.display = 'none';
                });

                this.elements.confirmSave.addEventListener('click', () => {
                    const fileName = this.elements.saveFileName.value.trim() || 'Untitled';
                    const fileFormat = this.elements.saveFileFormat.value;
                    let finalFileName = fileName;
                    
                    if (fileFormat === 'xlsx' && !fileName.endsWith('.xlsx')) {
                        finalFileName = fileName + '.xlsx';
                    } else if (fileFormat === 'csv' && !fileName.endsWith('.csv')) {
                        finalFileName = fileName + '.csv';
                    } else if (fileFormat === 'json' && !fileName.endsWith('.json')) {
                        finalFileName = fileName + '.json';
                    }
                    
                    this.data.fileName = finalFileName;
                    this.elements.saveModal.style.display = 'none';
                    
                    // Add to history before downloading
                    this.addToHistory();
                    
                    // Create a new workbook
                    const wb = XLSX.utils.book_new();
                    
                    // Prepare data for Excel (including SNo column)
                    const excelRows = this.data.rows.map((row, index) => {
                        const newRow = { 'SNo': index + 1 };
                        this.data.headers.forEach(header => {
                            newRow[header] = row[header] || '';
                        });
                        return newRow;
                    });
                    
                    if (fileFormat === 'xlsx' || fileFormat === 'csv') {
                        // Add headers as first row (including SNo)
                        const wsData = [
                            ['SNo', ...this.data.headers],
                            ...excelRows.map(row => {
                                return [row['SNo'], ...this.data.headers.map(header => row[header])];
                            })
                        ];
                        
                        const ws = XLSX.utils.aoa_to_sheet(wsData);
                        XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
                        
                        // Generate and download the file
                        XLSX.writeFile(wb, finalFileName);
                    } else if (fileFormat === 'json') {
                        // Create JSON data (including SNo in each row)
                        const jsonData = {
                            headers: ['SNo', ...this.data.headers],
                            rows: excelRows
                        };
                        
                        // Create download link
                        const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(jsonData, null, 2));
                        const downloadAnchorNode = document.createElement('a');
                        downloadAnchorNode.setAttribute("href", dataStr);
                        downloadAnchorNode.setAttribute("download", finalFileName);
                        document.body.appendChild(downloadAnchorNode);
                        downloadAnchorNode.click();
                        downloadAnchorNode.remove();
                    }
                    
                    this.saveData();
                    this.showToast(`File "${finalFileName}" downloaded successfully`, 'success');
                });

                // Analytics column change
                this.elements.analyticsColumn.addEventListener('change', () => {
                    this.renderCharts();
                });

                // Generate PDF report
                this.elements.generatePdf.addEventListener('click', () => {
                    if (this.data.rows.length === 0) {
                        this.showToast('No data to generate report', 'error');
                        return;
                    }
                    
                    this.showToast('PDF generation would be implemented in a real application', 'info');
                    // In a real app, you would use a library like jsPDF or pdfmake here
                });

                // Preview report
                this.elements.previewReport.addEventListener('click', () => {
                    if (this.data.rows.length === 0) {
                        this.showToast('No data to preview', 'error');
                        return;
                    }
                    
                    const reportTitle = document.getElementById('reportTitle').value || 'Data Report';
                    const reportDescription = document.getElementById('reportDescription').value || '';
                    
                    // Get selected columns
                    const selectedColumns = [];
                    document.querySelectorAll('input[name="reportColumns"]:checked').forEach(checkbox => {
                        selectedColumns.push(checkbox.value);
                    });
                    
                    if (selectedColumns.length === 0) {
                        this.showToast('Please select at least one column', 'error');
                        return;
                    }
                    
                    // Create preview content with SNo column
                    let previewContent = `
                        <h1 style="text-align: center; margin-bottom: 20px;">${reportTitle}</h1>
                        ${reportDescription ? `<p style="margin-bottom: 30px; text-align: center;">${reportDescription}</p>` : ''}
                        <table border="1" cellpadding="5" cellspacing="0" style="width: 100%; border-collapse: collapse; margin-bottom: 30px;">
                            <thead>
                                <tr>
                                    <th class="sno-column" style="width: 50px; text-align: center;">SNo</th>
                                    ${selectedColumns.map(col => `<th>${col}</th>`).join('')}
                                </tr>
                            </thead>
                            <tbody>
                                ${this.data.rows.map((row, index) => `
                                    <tr>
                                        <td class="sno-column" style="text-align: center;">${index + 1}</td>
                                        ${selectedColumns.map(col => `<td>${row[col] || ''}</td>`).join('')}
                                    </tr>
                                `).join('')}
                            </tbody>
                        </table>
                        <div class="print-footer">
                            <p style="text-align: right; margin-top: 30px; font-size: 12px;">
                                Generated on ${new Date().toLocaleString()}
                            </p>
                            <div class="signature-line">
                                <div class="signature">Prepared By: ___________________</div>
                                <div class="signature">Approved By: ___________________</div>
                            </div>
                        </div>
                    `;
                    
                    // Open in new window for preview
                    const previewWindow = window.open('', '_blank');
                    previewWindow.document.write(`
                        <!DOCTYPE html>
                        <html>
                        <head>
                            <title>Report Preview: ${reportTitle}</title>
                            <style>
                                body { font-family: Arial, sans-serif; padding: 20px; }
                                table { width: 100%; border-collapse: collapse; }
                                th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
                                th { background-color: #f2f2f2; }
                                .sno-column { width: 50px; text-align: center; }
                                .signature { width: 200px; border-top: 1px solid #000; text-align: center; padding-top: 10px; }
                                .signature-line { display: flex; justify-content: space-between; margin-top: 50px; }
                                .print-footer { margin-top: 50px; }
                            </style>
                        </head>
                        <body>
                            ${previewContent}
                        </body>
                        </html>
                    `);
                    previewWindow.document.close();
                });

                // Refresh history
                this.elements.refreshHistory.addEventListener('click', () => {
                    this.updateFileHistory();
                    this.showToast('File history refreshed', 'success');
                });

                // Clear history
                this.elements.clearHistory.addEventListener('click', () => {
                    if (confirm('Are you sure you want to clear all file history? This cannot be undone.')) {
                        this.data.fileHistory = [];
                        this.saveData();
                        this.updateFileHistory();
                        this.showToast('File history cleared', 'warning');
                    }
                });

                // Save settings
                this.elements.saveSettings.addEventListener('click', () => {
                    this.data.settings.theme = this.elements.appTheme.value;
                    this.data.settings.autoSave = parseInt(this.elements.autoSave.value) || 5;
                    this.data.settings.rowsPerPage = parseInt(this.elements.rowsPerPage.value) || 10;
                    this.saveData();
                    this.updateTheme();
                    this.showToast('Settings saved successfully', 'success');
                });

                // Reset settings
                this.elements.resetSettings.addEventListener('click', () => {
                    if (confirm('Reset all settings to default values?')) {
                        this.data.settings = {
                            theme: 'default',
                            autoSave: 5,
                            rowsPerPage: 10
                        };
                        this.saveData();
                        this.updateTheme();
                        this.showToast('Settings reset to defaults', 'success');
                    }
                });

                // Export settings
                this.elements.exportSettings.addEventListener('click', () => {
                    const settingsStr = JSON.stringify(this.data.settings, null, 2);
                    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(settingsStr);
                    const downloadAnchorNode = document.createElement('a');
                    downloadAnchorNode.setAttribute("href", dataStr);
                    downloadAnchorNode.setAttribute("download", "excel_manager_settings.json");
                    document.body.appendChild(downloadAnchorNode);
                    downloadAnchorNode.click();
                    downloadAnchorNode.remove();
                    this.showToast('Settings exported successfully', 'success');
                });

                // Import settings
                this.elements.importSettings.addEventListener('click', () => {
                    this.elements.importSettingsModal.style.display = 'flex';
                });

                this.elements.confirmImport.addEventListener('click', () => {
                    const file = this.elements.settingsFile.files[0];
                    if (!file) {
                        this.showToast('Please select a settings file', 'error');
                        return;
                    }
                    
                    this.elements.importSpinner.style.display = 'block';
                    
                    const reader = new FileReader();
                    reader.onload = (e) => {
                        try {
                            const settings = JSON.parse(e.target.result);
                            this.data.settings = settings;
                            this.saveData();
                            this.updateTheme();
                            this.elements.importSpinner.style.display = 'none';
                            this.elements.importSettingsModal.style.display = 'none';
                            this.showToast('Settings imported successfully', 'success');
                        } catch (error) {
                            this.elements.importSpinner.style.display = 'none';
                            this.showToast('Invalid settings file', 'error');
                        }
                    };
                    reader.readAsText(file);
                });

                // Close modal when clicking outside
                window.addEventListener('click', (e) => {
                    if (e.target === this.elements.saveModal) {
                        this.elements.saveModal.style.display = 'none';
                    }
                    if (e.target === this.elements.importSettingsModal) {
                        this.elements.importSettingsModal.style.display = 'none';
                    }
                });

                // Auto-save functionality
                setInterval(() => {
                    if (this.data.rows.length > 0) {
                        this.saveData();
                    }
                }, this.data.settings.autoSave * 60 * 1000);
            },

            // Edit row
            editRow(index) {
                const row = this.data.rows[index];
                this.data.headers.forEach(header => {
                    const inputId = header.toLowerCase().replace(/\s+/g, '_');
                    const input = document.getElementById(inputId);
                    if (input) {
                        input.value = row[header] || '';
                    }
                });
                this.elements.editIndexInput.value = index;
                // Switch to data entry tab
                document.querySelector('[data-tab="data-entry"]').click();
                window.scrollTo({ top: 0, behavior: 'smooth' });
                this.showToast('Editing record', 'info');
            },

            // Delete row
            deleteRow(index) {
                if (confirm('Are you sure you want to delete this row?')) {
                    const deletedRow = this.data.rows.splice(index, 1);
                    this.saveData();
                    this.showToast('Record deleted', 'warning');
                }
            },

            // Load data from history
            loadFromHistory(index) {
                if (confirm('Load this file? Current data will be replaced.')) {
                    const historyEntry = this.data.fileHistory[index];
                    this.data.fileName = historyEntry.fileName;
                    this.data.headers = [...historyEntry.headers];
                    this.data.rows = JSON.parse(JSON.stringify(historyEntry.rows));
                    this.data.lastModified = new Date().toLocaleString();
                    
                    this.saveData();
                    this.initializeForm();
                    this.showToast('File loaded from history', 'success');
                    
                    // Switch to data view tab
                    document.querySelector('[data-tab="data-view"]').click();
                }
            },

            // Download file from history
            downloadFromHistory(index) {
                const historyEntry = this.data.fileHistory[index];
                
                // Create a new workbook
                const wb = XLSX.utils.book_new();
                
                // Prepare data for Excel (including SNo column)
                const excelRows = historyEntry.rows.map((row, index) => {
                    const newRow = { 'SNo': index + 1 };
                    historyEntry.headers.forEach(header => {
                        newRow[header] = row[header] || '';
                    });
                    return newRow;
                });
                
                // Add headers as first row (including SNo)
                const wsData = [
                    ['SNo', ...historyEntry.headers],
                    ...excelRows.map(row => {
                        return [row['SNo'], ...historyEntry.headers.map(header => row[header])];
                    })
                ];
                
                const ws = XLSX.utils.aoa_to_sheet(wsData);
                XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
                
                // Generate and download the file
                XLSX.writeFile(wb, historyEntry.fileName);
                
                this.showToast(`File "${historyEntry.fileName}" downloaded from history`, 'success');
            },

            // Remove file from history
            removeFromHistory(index) {
                if (confirm('Remove this file from history?')) {
                    const removedFile = this.data.fileHistory.splice(index, 1);
                    this.saveData();
                    this.updateFileHistory();
                    this.showToast(`File "${removedFile[0].fileName}" removed from history`, 'warning');
                }
            }
        };

        // Initialize the application when DOM is loaded
        document.addEventListener('DOMContentLoaded', () => app.init());