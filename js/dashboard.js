/* ============================================
   WAM Audit Dashboard - Main JavaScript
   Real-time Google Sheets Integration
   Version: 2.0
   ============================================ */

// ============================================
// CONFIGURATION
// ============================================

const CONFIG = {
    // Google Sheets API Configuration
    // Get your API key from: https://console.cloud.google.com/apis/credentials
    API_KEY: 'AIzaSyD_zwQym7_DzSmfWS-qhj5UEwtGdhi47uY',
    // Get your Sheet ID from the URL: https://docs.google.com/spreadsheets/d/SHEET_ID/edit
    SHEET_ID: '1_Vd6SZiNs4a734dU5PHOyFeIMmLtJEfwQJb3oxqE7eY',
    SHEET_RANGE: 'Weekly Audit for Offices',  // Change to your sheet name
    
    // Auto-refresh settings (in milliseconds)
    AUTO_REFRESH_INTERVAL: 2 * 60 * 1000, // 2 minutes
    
    // Chart colors
    COLORS: {
        primary: '#1e3c72',
        secondary: '#2a5298',
        success: '#28a745',
        warning: '#ffc107',
        danger: '#dc3545',
        info: '#17a2b8',
        chartPalette: [
            '#1e3c72', '#2a5298', '#3498db', '#16a085', 
            '#27ae60', '#f39c12', '#e74c3c', '#9b59b6'
        ]
    }
};

// ============================================
// GLOBAL STATE
// ============================================

let appState = {
    sheetData: [],
    filteredData: [],
    charts: {},
    autoRefreshTimer: null,
    isLoading: false,
    headerRowIndex: 0,  // Which row contains the headers (0 or 1)
    filters: {
        timePeriod: 'all', // all, week, month, year
        personnel: 'all',  // all or specific personnel name
        office: 'all'      // all or specific office
    },
    // Table pagination state
    pagination: {
        currentPage: 1,
        rowsPerPage: 25,
        totalRows: 0,
        totalPages: 1
    },
    // Column visibility state
    visibleColumns: [],
    allTableData: [], // Store all filtered table data for pagination
    columnIndices: {
        personnel: -1,
        office: -1,
        date: -1,
        rockReview: 4  // Column E (Rock Review) - index 4 (0-based: A=0, B=1, C=2, D=3, E=4)
    }
};

// ============================================
// INITIALIZATION
// ============================================

/**
 * Initialize the dashboard when DOM is ready
 */
window.addEventListener('DOMContentLoaded', () => {
    console.log('🚀 WAM Audit Dashboard initialized');
    initializeDashboard();
});

/**
 * Main initialization function
 */
function initializeDashboard() {
    // Fetch data immediately on load
    fetchSheetData();
    
    // Add smooth scroll for navigation links
    setupSmoothScroll();
    
    // Log initialization
    console.log('✅ Dashboard ready');
}

// ============================================
// DATA FETCHING
// ============================================

/**
 * Fetch data from Google Sheets API
 */
async function fetchSheetData() {
    if (appState.isLoading) {
        console.log('⏳ Already loading data...');
        return;
    }

    appState.isLoading = true;
    showLoading(true);
    updateSyncStatus('syncing');

    try {
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.SHEET_ID}/values/${CONFIG.SHEET_RANGE}?key=${CONFIG.API_KEY}`;
        
        console.log('📡 Fetching data from Google Sheets...');
        const response = await fetch(url);
        const data = await response.json();

        // Handle API errors
        if (data.error) {
            throw new Error(data.error.message || 'Failed to fetch data from Google Sheets');
        }

        // Validate data
        if (!data.values || data.values.length === 0) {
            throw new Error('No data found in the specified sheet range');
        }

        console.log(`✅ Loaded ${data.values.length} rows from Google Sheets`);
        
        // Store data in global state
        appState.sheetData = data.values;
        
        // Detect column indices automatically (also sets headerRowIndex)
        detectColumnIndices();
        
        // Initialize filtered data with all data (skip title and header rows)
        const dataStartRow = appState.headerRowIndex + 1;
        appState.filteredData = appState.sheetData.slice(dataStartRow);
        
        console.log(`📊 Data rows: ${appState.filteredData.length} (starting from row ${dataStartRow + 1})`);
        
        // Initialize advanced filters
        initializeFilters();
        
        // Render filters first
        renderFilters();
        
        // Render all dashboard components
        renderDashboard();
        
        // Update UI status
        updateSyncStatus('success');
        updateLastSync();
        showNotification('Data synced successfully!', 'success');
        
        // Setup auto-refresh timer
        setupAutoRefresh();

    } catch (error) {
        console.error('❌ Error fetching data:', error);
        updateSyncStatus('error');
        showNotification(`Error: ${error.message}`, 'danger');
        
        // Show helpful error message
        showErrorHelp(error.message);
    } finally {
        appState.isLoading = false;
        showLoading(false);
    }
}

/**
 * Refresh data manually
 */
function refreshData() {
    console.log('🔄 Manual refresh triggered');
    fetchSheetData();
}

// ============================================
// DASHBOARD RENDERING
// ============================================

/**
 * Render all dashboard components
 */
function renderDashboard() {
    console.log('🎨 Rendering dashboard components...');
    
    try {
        renderSummaryCards();
        createRockReviewChart();  // New Rock Review analysis
        renderKPIMetrics();
        renderDataTable();
        renderAllCharts();
        renderProgressBars();
        
        console.log('✅ Dashboard rendered successfully');
    } catch (error) {
        console.error('❌ Error rendering dashboard:', error);
        showNotification('Error rendering dashboard components', 'danger');
    }
}

/**
 * Render summary statistics cards
 */
function renderSummaryCards() {
    if (appState.sheetData.length === 0) return;

    const headers = appState.sheetData[0];
    const rows = appState.filteredData.length > 0 ? appState.filteredData : appState.sheetData.slice(appState.headerRowIndex + 1);
    
    // Filter out header rows
    const dataRows = rows.filter(row => {
        const officeName = (row[appState.columnIndices.office] || '').trim().toUpperCase();
        return officeName !== 'OFFICE NAME' && 
               !officeName.includes('WEEKLY ACCOUNTABILITY') &&
               !officeName.includes('WAM') && 
               officeName !== '';
    });
    
    const totalRecords = dataRows.length;

    // Count unique offices
    const uniqueOffices = new Set();
    dataRows.forEach(row => {
        const officeName = (row[appState.columnIndices.office] || '').trim();
        if (officeName) uniqueOffices.add(officeName);
    });
    const totalOffices = uniqueOffices.size;

    // Calculate On-Track rate from Column E (Rock Review)
    let onTrackCount = 0;
    dataRows.forEach(row => {
        const rockReview = (row[appState.columnIndices.rockReview] || '').toLowerCase().trim();
        if (rockReview.includes('on') && rockReview.includes('track') && !rockReview.includes('off')) {
            onTrackCount++;
        }
    });
    const onTrackRate = totalRecords > 0 ? ((onTrackCount / totalRecords) * 100).toFixed(1) : 0;

    // Calculate average meeting rating from Column H (CONCLUDE)
    const concludeColIndex = 7; // Column H
    let totalRating = 0;
    let ratingCount = 0;
    dataRows.forEach(row => {
        const rating = parseFloat(row[concludeColIndex]);
        if (!isNaN(rating) && rating >= 1 && rating <= 10) {
            totalRating += rating;
            ratingCount++;
        }
    });
    const avgRating = ratingCount > 0 ? (totalRating / ratingCount).toFixed(1) : 0;

    const cardsHTML = `
        <div class="col-md-6 col-lg-3">
            <div class="card stat-card animate__animated animate__bounceIn">
                <div class="card-body position-relative">
                    <i class="bi bi-file-earmark-text stat-icon"></i>
                    <p class="stat-label">Total Audits</p>
                    <h2 class="stat-value">${totalRecords}</h2>
                    <span class="badge bg-info badge-custom">WAM Records</span>
                </div>
            </div>
        </div>
        <div class="col-md-6 col-lg-3">
            <div class="card stat-card animate__animated animate__bounceIn" style="animation-delay: 0.1s">
                <div class="card-body position-relative">
                    <i class="bi bi-building stat-icon text-primary"></i>
                    <p class="stat-label">Total Offices</p>
                    <h2 class="stat-value text-primary">${totalOffices}</h2>
                    <span class="badge bg-primary badge-custom">Unique Locations</span>
                </div>
            </div>
        </div>
        <div class="col-md-6 col-lg-3">
            <div class="card stat-card animate__animated animate__bounceIn" style="animation-delay: 0.2s">
                <div class="card-body position-relative">
                    <i class="bi bi-bullseye stat-icon ${onTrackRate >= 80 ? 'text-success' : onTrackRate >= 60 ? 'text-warning' : 'text-danger'}"></i>
                    <p class="stat-label">On-Track Rate</p>
                    <h2 class="stat-value ${onTrackRate >= 80 ? 'text-success' : onTrackRate >= 60 ? 'text-warning' : 'text-danger'}">${onTrackRate}%</h2>
                    <span class="badge ${onTrackRate >= 80 ? 'bg-success' : onTrackRate >= 60 ? 'bg-warning' : 'bg-danger'} badge-custom">${onTrackCount} of ${totalRecords}</span>
                </div>
            </div>
        </div>
        <div class="col-md-6 col-lg-3">
            <div class="card stat-card animate__animated animate__bounceIn" style="animation-delay: 0.3s">
                <div class="card-body position-relative">
                    <i class="bi bi-star-fill stat-icon ${avgRating >= 8 ? 'text-success' : avgRating >= 6 ? 'text-warning' : 'text-danger'}"></i>
                    <p class="stat-label">Avg Meeting Rating</p>
                    <h2 class="stat-value ${avgRating >= 8 ? 'text-success' : avgRating >= 6 ? 'text-warning' : 'text-danger'}">${avgRating}/10</h2>
                    <span class="badge ${avgRating >= 8 ? 'bg-success' : avgRating >= 6 ? 'bg-warning' : 'bg-danger'} badge-custom">${ratingCount} ratings</span>
                </div>
            </div>
        </div>
    `;

    document.getElementById('summaryCards').innerHTML = cardsHTML;
}

/**
 * Render KPI metrics for numeric columns
 */
function renderKPIMetrics() {
    if (appState.sheetData.length === 0) return;

    const headers = appState.sheetData[0];
    const rows = appState.sheetData.slice(1);
    let metricsHTML = '';

    headers.forEach((header, index) => {
        const numericValues = rows
            .map(row => parseFloat(row[index]))
            .filter(val => !isNaN(val));

        if (numericValues.length > 0) {
            const sum = numericValues.reduce((a, b) => a + b, 0);
            const avg = (sum / numericValues.length).toFixed(2);
            const max = Math.max(...numericValues);
            const min = Math.min(...numericValues);

            metricsHTML += `
                <div class="col-md-6 col-lg-4">
                    <div class="metric-card animate__animated animate__fadeInUp">
                        <h6 class="text-uppercase mb-3">
                            <i class="bi bi-currency-dollar"></i> ${header}
                        </h6>
                        <div class="d-flex justify-content-between align-items-center mb-2">
                            <span>Total:</span>
                            <strong class="fs-4">${sum.toLocaleString()}</strong>
                        </div>
                        <div class="d-flex justify-content-between align-items-center mb-2">
                            <span>Average:</span>
                            <strong>${parseFloat(avg).toLocaleString()}</strong>
                        </div>
                        <div class="d-flex justify-content-between align-items-center">
                            <span>Range:</span>
                            <strong>${min.toLocaleString()} - ${max.toLocaleString()}</strong>
                        </div>
                    </div>
                </div>
            `;
        }
    });

    if (metricsHTML) {
        document.getElementById('kpiMetrics').innerHTML = metricsHTML;
    } else {
        document.getElementById('kpiMetrics').innerHTML = '';
    }
}

/**
 * Render data table
 */
function renderDataTable() {
    console.log('📊 renderDataTable called');
    if (appState.sheetData.length === 0) {
        console.log('❌ No sheet data available');
        return;
    }

    // Use the correct header row (row index 2 = 3rd row)
    const headers = appState.sheetData[2]; // Row 3 contains the actual headers
    let rows = appState.sheetData.slice(3); // Data starts from row 4
    
    console.log('📋 Headers:', headers);
    console.log('📊 Total rows before filtering:', rows.length);
    
    // Filter out header/title rows
    const officeColIndex = appState.columnIndices.office; // Column A = 0
    const rockReviewColIndex = appState.columnIndices.rockReview; // Column E = 4
    
    rows = rows.filter(row => {
        const officeName = (row[officeColIndex] || '').trim().toUpperCase();
        return officeName !== 'OFFICE NAME' && 
               !officeName.includes('WEEKLY ACCOUNTABILITY') &&
               !officeName.includes('WAM') && 
               officeName !== '';
    });
    
    console.log('📊 Rows after filtering:', rows.length);

    // Store all data for pagination and filtering
    appState.allTableData = rows;
    appState.pagination.totalRows = rows.length;
    
    // Initialize visible columns if not set
    if (appState.visibleColumns.length === 0) {
        appState.visibleColumns = headers.map((_, index) => index);
    }
    
    // Create column checkboxes
    renderColumnCheckboxes(headers);
    
    // Render paginated data
    renderTablePage();
}

function renderColumnCheckboxes(headers) {
    const checkboxContainer = document.getElementById('columnCheckboxes');
    if (!checkboxContainer) return;
    
    const checkboxHTML = headers.map((header, index) => `
        <div class="col-md-4 col-lg-3 mb-2">
            <div class="form-check">
                <input class="form-check-input" type="checkbox" value="${index}" 
                       id="col${index}" ${appState.visibleColumns.includes(index) ? 'checked' : ''}
                       onchange="toggleColumn(${index})">
                <label class="form-check-label small" for="col${index}">
                    ${escapeHtml(header.substring(0, 30))}${header.length > 30 ? '...' : ''}
                </label>
            </div>
        </div>
    `).join('');
    
    checkboxContainer.innerHTML = checkboxHTML;
}

function renderTablePage() {
    const headers = appState.sheetData[2]; // Use correct header row
    const rowsPerPage = appState.pagination.rowsPerPage;
    const currentPage = appState.pagination.currentPage;
    
    console.log('📄 renderTablePage - Total data rows:', appState.allTableData.length);
    
    // Calculate pagination
    const startIndex = rowsPerPage === 'all' ? 0 : (currentPage - 1) * rowsPerPage;
    const endIndex = rowsPerPage === 'all' ? appState.allTableData.length : startIndex + parseInt(rowsPerPage);
    const pageRows = appState.allTableData.slice(startIndex, endIndex);
    
    console.log('📄 Showing rows:', startIndex, 'to', endIndex, '- Page rows:', pageRows.length);
    
    const totalPages = rowsPerPage === 'all' ? 1 : Math.ceil(appState.allTableData.length / rowsPerPage);
    appState.pagination.totalPages = totalPages;
    
    // Create table headers (only visible columns)
    const headerHTML = '<tr>' + headers.map((header, index) => {
        if (!appState.visibleColumns.includes(index)) return '';
        return `<th onclick="sortTable(${index})" style="cursor: pointer;">
                    ${escapeHtml(header)} <i class="bi bi-arrow-down-up"></i>
                </th>`;
    }).join('') + '</tr>';

    // Create table rows with highlighting for Off-Track
    const rockReviewColIndex = appState.columnIndices.rockReview;
    const rowsHTML = pageRows.map(row => {
        const rockReview = (row[rockReviewColIndex] || '').toLowerCase();
        const isOffTrack = rockReview.includes('off') && rockReview.includes('track');
        const rowClass = isOffTrack ? 'table-danger' : '';
        
        return `<tr class="${rowClass}">` + headers.map((_, index) => {
            if (!appState.visibleColumns.includes(index)) return '';
            return `<td>${escapeHtml(row[index] || '')}</td>`;
        }).join('') + '</tr>';
    }).join('');

    document.getElementById('tableHead').innerHTML = headerHTML;
    document.getElementById('tableBody').innerHTML = rowsHTML;
    document.getElementById('rowCount').textContent = appState.allTableData.length;
    
    // Update pagination info
    const showingStart = rowsPerPage === 'all' ? 1 : startIndex + 1;
    const showingEnd = rowsPerPage === 'all' ? appState.allTableData.length : Math.min(endIndex, appState.allTableData.length);
    document.getElementById('showingRange').textContent = `${showingStart}-${showingEnd}`;
    document.getElementById('totalRows').textContent = appState.allTableData.length;
    
    // Render pagination controls
    renderPaginationControls(totalPages);
}

// ============================================
// CHART RENDERING
// ============================================

/**
 * Render all charts
 */
function renderAllCharts() {
    if (appState.sheetData.length === 0) return;

    // Destroy existing charts
    Object.values(appState.charts).forEach(chart => {
        if (chart) chart.destroy();
    });
    appState.charts = {};

    const headers = appState.sheetData[appState.headerRowIndex];
    const dataStartRow = appState.headerRowIndex + 1;
    let rows = appState.filteredData.length > 0 ? appState.filteredData : appState.sheetData.slice(dataStartRow);

    // Filter out header/title rows from data
    const officeColIndex = appState.columnIndices.office;
    rows = rows.filter(row => {
        const officeName = (row[officeColIndex] || '').trim().toUpperCase();
        // Exclude header rows and title rows
        return officeName !== 'OFFICE NAME' && 
               !officeName.includes('WEEKLY ACCOUNTABILITY') &&
               !officeName.includes('WAM') && 
               officeName !== '';
    });
    
    console.log(`📊 Filtered out header rows. Rows remaining: ${rows.length}`);

    const rockReviewColIndex = appState.columnIndices.rockReview;
    const dateColIndex = appState.columnIndices.date;
    
    if (officeColIndex >= 0) {
        createPieChart('Office Distribution', rows, officeColIndex);
    } else {
        // Fallback to first column if office column not detected
        createPieChart('Distribution', rows, 0);
    }

    // Create bar chart for On-Track vs Off-Track by Office
    console.log('📊 Bar Chart Check:', { officeColIndex, rockReviewColIndex, rows: rows.length });
    
    // Show diagnostic message on page
    const statusDiv = document.getElementById('barChartStatus');
    const statusText = document.getElementById('barChartStatusText');
    if (statusDiv && statusText) {
        statusDiv.style.display = 'block';
        statusText.textContent = `Checking: officeColIndex=${officeColIndex}, rockReviewColIndex=${rockReviewColIndex}`;
    }
    
    if (officeColIndex >= 0 && rockReviewColIndex >= 0) {
        createOfficePerformanceBarChart(rows, officeColIndex, rockReviewColIndex);
    } else {
        console.warn('⚠️ Bar chart not created. Office:', officeColIndex, 'RockReview:', rockReviewColIndex);
        
        // Show why it's not created
        if (statusDiv && statusText) {
            statusDiv.style.background = '#f8d7da';
            statusDiv.style.borderColor = '#dc3545';
            statusText.textContent = `❌ Bar chart not created! Office column: ${officeColIndex}, Rock Review column: ${rockReviewColIndex}. Need both >= 0.`;
        }
    }
    
    // Create line chart for trend analysis
    console.log('📈 Line Chart Check:', { dateColIndex, rockReviewColIndex, rows: rows.length });
    
    const lineStatusDiv = document.getElementById('lineChartStatus');
    const lineStatusText = document.getElementById('lineChartStatusText');
    if (lineStatusDiv && lineStatusText) {
        lineStatusDiv.style.display = 'block';
        lineStatusText.textContent = `Checking: dateColIndex=${dateColIndex}, rockReviewColIndex=${rockReviewColIndex}`;
    }
    
    if (dateColIndex >= 0 && rockReviewColIndex >= 0) {
        createTrendLineChart(rows, dateColIndex, rockReviewColIndex);
    } else {
        console.warn('⚠️ Line chart not created. Date:', dateColIndex, 'RockReview:', rockReviewColIndex);
        
        // Show why it's not created
        if (lineStatusDiv && lineStatusText) {
            lineStatusDiv.style.background = '#f8d7da';
            lineStatusDiv.style.borderColor = '#dc3545';
            lineStatusText.textContent = `❌ Line chart not created! Date column: ${dateColIndex}, Rock Review column: ${rockReviewColIndex}. Need both >= 0.`;
        }
    }
}

/**
 * Create pie/doughnut chart for Office Distribution
 */
function createPieChart(headerName, rows, colIndex) {
    const officeCounts = {};
    rows.forEach(row => {
        const office = row[colIndex] || 'Unknown';
        const officeName = office.trim();
        officeCounts[officeName] = (officeCounts[officeName] || 0) + 1;
    });

    // Sort by count (descending)
    const sortedOffices = Object.entries(officeCounts)
        .sort((a, b) => b[1] - a[1]);
    
    const labels = sortedOffices.map(([name, _]) => name);
    const data = sortedOffices.map(([_, count]) => count);
    const total = data.reduce((a, b) => a + b, 0);

    // Destroy existing chart if it exists
    if (appState.charts.pieChart) {
        appState.charts.pieChart.destroy();
    }

    appState.charts.pieChart = new Chart(document.getElementById('pieChart'), {
        type: 'doughnut',
        data: {
            labels: labels.map((label, i) => {
                const percent = ((data[i] / total) * 100).toFixed(1);
                return `${label} (${percent}%)`;
            }),
            datasets: [{
                data: data,
                backgroundColor: [
                    '#2563eb', '#10b981', '#f59e0b', '#ef4444', 
                    '#8b5cf6', '#06b6d4', '#ec4899', '#14b8a6',
                    '#f97316', '#6366f1', '#84cc16', '#a855f7'
                ],
                borderWidth: 3,
                borderColor: '#fff',
                hoverOffset: 10
            }]
        },
        options: {
                responsive: true,
            maintainAspectRatio: false,
            cutout: '50%',
            layout: {
                padding: {
                    top: 10,
                    bottom: 10,
                    left: 10,
                    right: 10
                }
            },
            plugins: {
                legend: {
                    position: 'bottom',
                    align: 'center',
                    labels: { 
                        padding: 12, 
                        font: { size: 11 },
                        usePointStyle: true,
                        pointStyle: 'circle',
                        boxWidth: 12,
                        boxHeight: 12
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    padding: 12,
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            const office = labels[context.dataIndex];
                            const count = context.parsed;
                            const percentage = ((count / total) * 100).toFixed(1);
                            return `${office}: ${count} records (${percentage}%) - Click to view details`;
                        }
                    }
                },
                title: {
                    display: false
                }
            },
            onClick: (event, activeElements) => {
                if (activeElements.length > 0) {
                    const index = activeElements[0].index;
                    const officeName = labels[index];
                    showOfficeDetails(officeName, colIndex);
                }
            },
            onHover: (event, activeElements) => {
                event.native.target.style.cursor = activeElements.length > 0 ? 'pointer' : 'default';
            }
        }
    });
    
    console.log('📊 Office Distribution:', officeCounts);
}

/**
 * Show details for a specific office when clicked
 */
function showOfficeDetails(officeName, officeColIndex) {
    console.log(`🔍 Showing details for office: ${officeName}`);
    
    // Filter data to show only this office
    const headers = appState.sheetData[appState.headerRowIndex];
    const dataStartRow = appState.headerRowIndex + 1;
    const allRows = appState.sheetData.slice(dataStartRow);
    
    const officeRows = allRows.filter(row => {
        const office = (row[officeColIndex] || '').trim();
        return office === officeName;
    });
    
    // Create a modal/detail view
    const detailsHTML = `
        <div class="office-details-overlay" id="officeDetailsOverlay" onclick="closeOfficeDetails()">
            <div class="office-details-modal" onclick="event.stopPropagation()">
                <div class="d-flex justify-content-between align-items-center mb-4">
                    <h4 class="mb-0">
                        <i class="bi bi-building-fill"></i> ${officeName} - Detailed Records
                    </h4>
                    <button class="btn btn-sm btn-outline-secondary" onclick="closeOfficeDetails()">
                        <i class="bi bi-x-lg"></i> Close
                    </button>
                </div>
                
                <div class="mb-3">
                    <span class="badge bg-primary fs-6">Total Records: ${officeRows.length}</span>
                </div>
                
                <div class="table-responsive" style="max-height: 500px; overflow-y: auto;">
                    <table class="table table-hover table-striped">
                        <thead class="sticky-top bg-white">
                            <tr>
                                ${headers.slice(0, 10).map(h => `<th>${h}</th>`).join('')}
                            </tr>
                        </thead>
                        <tbody>
                            ${officeRows.map(row => `
                                <tr>
                                    ${row.slice(0, 10).map(cell => `<td>${cell || '-'}</td>`).join('')}
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>
                
                <div class="mt-4 text-center">
                    <button class="btn btn-success" onclick="exportOfficeData('${officeName}')">
                        <i class="bi bi-download"></i> Export ${officeName} Data to CSV
                    </button>
                    <button class="btn btn-primary ms-2" onclick="filterByOffice('${officeName}')">
                        <i class="bi bi-filter"></i> Apply as Filter
                    </button>
                </div>
            </div>
        </div>
    `;
    
    // Remove existing overlay if any
    const existingOverlay = document.getElementById('officeDetailsOverlay');
    if (existingOverlay) {
        existingOverlay.remove();
    }
    
    // Add to body
    document.body.insertAdjacentHTML('beforeend', detailsHTML);
    
    // Show notification
    showNotification(`Showing ${officeRows.length} records for ${officeName}`, 'info');
}

/**
 * Close office details modal
 */
function closeOfficeDetails() {
    const overlay = document.getElementById('officeDetailsOverlay');
    if (overlay) {
        overlay.style.opacity = '0';
        setTimeout(() => overlay.remove(), 300);
    }
}

/**
 * Export specific office data to CSV
 */
function exportOfficeData(officeName) {
    const headers = appState.sheetData[appState.headerRowIndex];
    const dataStartRow = appState.headerRowIndex + 1;
    const allRows = appState.sheetData.slice(dataStartRow);
    const officeColIndex = appState.columnIndices.office;
    
    const officeRows = allRows.filter(row => {
        const office = (row[officeColIndex] || '').trim();
        return office === officeName;
    });
    
    // Create CSV content
    let csv = headers.join(',') + '\n';
    officeRows.forEach(row => {
        csv += row.map(cell => `"${cell || ''}"`).join(',') + '\n';
    });
    
    // Download
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${officeName}_data_${new Date().toISOString().split('T')[0]}.csv`;
    a.click();
    window.URL.revokeObjectURL(url);
    
    showNotification(`Exported ${officeRows.length} records for ${officeName}`, 'success');
}

/**
 * Filter dashboard by specific office
 */
function filterByOffice(officeName) {
    // Set the office filter dropdown
    const officeFilter = document.getElementById('officeFilter');
    if (officeFilter) {
        officeFilter.value = officeName;
        applyFilters();
        closeOfficeDetails();
        
        // Scroll to top
        window.scrollTo({ top: 0, behavior: 'smooth' });
        
        showNotification(`Filtered dashboard to show only ${officeName}`, 'success');
    }
}

/**
 * Create stacked bar chart for On-Track vs Off-Track by Office
 */
function createOfficePerformanceBarChart(rows, officeColIndex, rockReviewColIndex) {
    // Show status on page
    const statusDiv = document.getElementById('barChartStatus');
    const statusText = document.getElementById('barChartStatusText');
    
    if (statusDiv) {
        statusDiv.style.display = 'block';
        statusText.textContent = 'Function called with ' + rows.length + ' rows, office col: ' + officeColIndex + ', review col: ' + rockReviewColIndex;
    }
    
    console.log('%c ========== BAR CHART FUNCTION CALLED ========== ', 'background: red; color: white; font-size: 20px; padding: 10px;');
    console.log('Params:', { rows: rows.length, officeColIndex, rockReviewColIndex });
    
    const canvas = document.getElementById('barChart');
    console.log('Canvas element:', canvas);
    
    if (!canvas) {
        console.error('%c ❌ CANVAS NOT FOUND! ', 'background: red; color: white; font-size: 16px;');
        if (statusText) statusText.textContent = '❌ ERROR: Canvas element not found!';
        return;
    }
    
    if (rows.length === 0) {
        console.warn('⚠️ No rows');
        if (statusText) statusText.textContent = '⚠️ No data rows to display';
        return;
    }
    
    if (statusText) statusText.textContent = 'Processing ' + rows.length + ' rows...';
    
    // Group data by office
    const officeData = {};
    rows.forEach(row => {
        const office = (row[officeColIndex] || 'Unknown').trim();
        const rockReview = (row[rockReviewColIndex] || '').toLowerCase().trim();
        
        if (!officeData[office]) {
            officeData[office] = { onTrack: 0, offTrack: 0, other: 0 };
        }
        
        if (rockReview.includes('on') && rockReview.includes('track') && !rockReview.includes('off')) {
            officeData[office].onTrack++;
        } else if (rockReview.includes('off') && rockReview.includes('track')) {
            officeData[office].offTrack++;
        } else {
            officeData[office].other++;
        }
    });
    
    // Get top 10 offices
    const sortedOffices = Object.entries(officeData)
        .sort((a, b) => {
            const totalA = a[1].onTrack + a[1].offTrack + a[1].other;
            const totalB = b[1].onTrack + b[1].offTrack + b[1].other;
            return totalB - totalA;
        })
        .slice(0, 10);
    
    if (sortedOffices.length === 0) {
        canvas.parentElement.innerHTML = '<p style="text-align: center; padding: 40px;">No data</p>';
        return;
    }
    
    const labels = sortedOffices.map(([office]) => office);
    const onTrackData = sortedOffices.map(([_, data]) => data.onTrack);
    const offTrackData = sortedOffices.map(([_, data]) => data.offTrack);
    const otherData = sortedOffices.map(([_, data]) => data.other);
    
    console.log('📊 Chart data ready:', { labels, onTrackData, offTrackData, otherData });
    
    if (statusText) statusText.textContent = 'Data processed. Found ' + labels.length + ' offices. Creating chart...';
    
    // Destroy existing
    if (appState.charts.barChart) {
        try {
            appState.charts.barChart.destroy();
        } catch (e) {}
    }
    
    if (statusText) statusText.textContent = 'Calling Chart.js to create bar chart...';
    
    // Create chart
    try {
        appState.charts.barChart = new Chart(canvas, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [{
                    label: 'On-Track',
                    data: onTrackData,
                    backgroundColor: '#10b981'
                }, {
                    label: 'Off-Track',
                    data: offTrackData,
                    backgroundColor: '#ef4444'
                }, {
                    label: 'Other',
                    data: otherData,
                    backgroundColor: '#9ca3af'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                indexAxis: 'y',
                plugins: {
                    legend: {
                        position: 'top'
                    }
                },
                scales: {
                    x: {
                        stacked: true,
                        beginAtZero: true
                    },
                    y: {
                        stacked: true
                    }
                }
            }
        });
        
        console.log('%c ✅ BAR CHART CREATED SUCCESSFULLY! ', 'background: green; color: white; font-size: 16px; padding: 5px;');
        console.log('Chart instance:', appState.charts.barChart);
        console.log('Canvas after creation:', canvas);
        console.log('Canvas dimensions:', canvas.width, 'x', canvas.height);
        
        // Update status to success
        if (statusDiv && statusText) {
            statusDiv.style.background = '#d4edda';
            statusDiv.style.borderColor = '#28a745';
            statusText.textContent = '✅ Chart created successfully! Showing ' + labels.length + ' offices.';
            
            // Hide success message after 3 seconds
            setTimeout(() => {
                statusDiv.style.display = 'none';
            }, 3000);
        }
        
    } catch (error) {
        console.error('%c ❌ ERROR CREATING CHART! ', 'background: red; color: white; font-size: 16px;');
        console.error('Error:', error);
        console.error('Message:', error.message);
        console.error('Stack:', error.stack);
        
        // Show error on page
        if (statusDiv && statusText) {
            statusDiv.style.background = '#f8d7da';
            statusDiv.style.borderColor = '#dc3545';
            statusText.textContent = '❌ ERROR: ' + error.message;
        }
        
        // Show error in the container
        canvas.parentElement.innerHTML = '<p style="color: red; padding: 20px; text-align: center;"><strong>Error creating chart:</strong><br>' + error.message + '</p>';
    }
}

/**
 * Create line chart for trend analysis over time
 */
function createTrendLineChart(rows, dateColIndex, rockReviewColIndex) {
    console.log('📈 Creating Trend Line Chart...');
    console.log('📊 Input:', { rows: rows.length, dateColIndex, rockReviewColIndex });
    
    const lineStatusDiv = document.getElementById('lineChartStatus');
    const lineStatusText = document.getElementById('lineChartStatusText');
    
    if (lineStatusDiv && lineStatusText) {
        lineStatusDiv.style.display = 'block';
        lineStatusText.textContent = 'Processing ' + rows.length + ' rows for trend analysis...';
    }
    
    const canvas = document.getElementById('lineChart');
    if (!canvas) {
        console.error('❌ Line chart canvas not found!');
        if (lineStatusText) lineStatusText.textContent = '❌ Canvas element not found!';
        return;
    }
    
    // Force canvas container size (let Chart.js handle canvas dimensions for crisp rendering)
    const parentWidth = canvas.parentElement.offsetWidth;
    canvas.style.height = '450px';
    canvas.style.width = '100%';
    canvas.parentElement.style.height = '450px';
    canvas.parentElement.style.width = '100%';
    
    // Update visible debug info
    const debugDiv = document.getElementById('canvasDimensions');
    if (debugDiv) {
        debugDiv.textContent = `Initial: Container ${parentWidth}px x 450px (Chart.js handles pixel ratio)`;
        debugDiv.style.color = '#1976d2';
    }
    console.log('📐 Initial canvas container set to:', parentWidth, 'x 450');
    
    if (rows.length === 0) {
        console.warn('⚠️ No rows to display in line chart');
        if (lineStatusText) lineStatusText.textContent = '⚠️ No data to display';
        return;
    }
    
    // Group data by month
    const monthlyData = {};
    
    rows.forEach(row => {
        const dateStr = row[dateColIndex];
        if (!dateStr) return;
        
        const date = parseDate(dateStr);
        if (!date) return;
        
        const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
        const rockReview = (row[rockReviewColIndex] || '').toLowerCase().trim();
        
        if (!monthlyData[monthKey]) {
            monthlyData[monthKey] = { onTrack: 0, offTrack: 0, total: 0 };
        }
        
        if (rockReview.includes('on') && rockReview.includes('track') && !rockReview.includes('off')) {
            monthlyData[monthKey].onTrack++;
        } else if (rockReview.includes('off') && rockReview.includes('track')) {
            monthlyData[monthKey].offTrack++;
        }
        monthlyData[monthKey].total++;
    });
    
    // Sort by date
    const sortedMonths = Object.entries(monthlyData)
        .sort((a, b) => a[0].localeCompare(b[0]));
    
    const labels = sortedMonths.map(([month, _]) => {
        const [year, m] = month.split('-');
        const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        return `${monthNames[parseInt(m) - 1]} ${year}`;
    });
    
    const onTrackPercentages = sortedMonths.map(([_, data]) => 
        data.total > 0 ? ((data.onTrack / data.total) * 100).toFixed(1) : 0
    );
    
    const totalRecords = sortedMonths.map(([_, data]) => data.total);
    
    // Destroy existing chart
    if (appState.charts.lineChart) {
        appState.charts.lineChart.destroy();
    }

    // Re-apply canvas container size after destruction
    const currentParentWidth = canvas.parentElement.offsetWidth;
    canvas.style.height = '450px';
    canvas.style.width = '100%';
    canvas.parentElement.style.height = '450px';
    canvas.parentElement.style.width = '100%';
    
    // Update visible debug info
    const debugDiv2 = document.getElementById('canvasDimensions');
    if (debugDiv2) {
        debugDiv2.textContent = `After Destroy: Container ${currentParentWidth}px x 450px`;
        debugDiv2.style.color = '#f57c00';
    }
    console.log('📐 Canvas container set after destroy:', currentParentWidth, 'x 450');

    appState.charts.lineChart = new Chart(document.getElementById('lineChart'), {
        type: 'line',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'On-Track Rate (%)',
                    data: onTrackPercentages,
                    borderColor: '#10b981',
                    backgroundColor: 'rgba(16, 185, 129, 0.1)',
                    borderWidth: 6,
                    tension: 0.4,
                    fill: true,
                    pointRadius: 12,
                    pointHoverRadius: 18,
                    pointBackgroundColor: '#10b981',
                    pointBorderColor: '#fff',
                    pointBorderWidth: 5
                },
                {
                    label: 'Total Records',
                    data: totalRecords,
                    borderColor: '#2563eb',
                    backgroundColor: 'rgba(37, 99, 235, 0.1)',
                    borderWidth: 5,
                    tension: 0.4,
                    fill: false,
                    pointRadius: 10,
                    pointHoverRadius: 15,
                    pointBackgroundColor: '#2563eb',
                    pointBorderColor: '#fff',
                    pointBorderWidth: 4,
                    yAxisID: 'y2'
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            onResize: (chart, size) => {
                // Update debug info when Chart.js resizes
                const debugDiv = document.getElementById('canvasDimensions');
                if (debugDiv) {
                    debugDiv.textContent = `⚠️ Chart.js auto-resized to: ${size.width}px x ${size.height}px`;
                    debugDiv.style.color = '#d32f2f';
                }
                console.log('📐 Chart.js resized to:', size.width, 'x', size.height);
            },
            interaction: {
                mode: 'index',
                intersect: false
            },
            plugins: {
                legend: {
                    display: true,
                    position: 'top',
                    labels: {
                        padding: 30,
                        font: { size: 18, weight: 'bold' },
                        usePointStyle: true,
                        pointStyle: 'circle',
                        boxWidth: 22,
                        boxHeight: 22
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.9)',
                    padding: 16,
                    cornerRadius: 8,
                    titleFont: { size: 14, weight: 'bold' },
                    bodyFont: { size: 13 },
                    callbacks: {
                        label: function(context) {
                            const label = context.dataset.label;
                            const value = context.parsed.y;
                            if (label === 'On-Track Rate (%)') {
                                return `  ${label}: ${value}%`;
                            } else {
                                return `  ${label}: ${value} audit records`;
                            }
                        },
                        afterLabel: function(context) {
                            const label = context.dataset.label;
                            if (label === 'Total Records') {
                                return '  (Number of audits completed this month)';
                            }
                            return '';
                        }
                    }
                }
            },
            scales: {
                y: {
                    type: 'linear',
                    display: true,
                    position: 'left',
                    beginAtZero: true,
                    max: 100,
                    ticks: {
                        font: { size: 16 },
                        callback: value => value + '%'
                    },
                    title: {
                        display: true,
                        text: 'On-Track Rate (%)',
                        font: { size: 16, weight: 'bold' }
                    },
                    grid: {
                        color: 'rgba(0, 0, 0, 0.05)'
                    }
                },
                y2: {
                    type: 'linear',
                    display: true,
                    position: 'right',
                    beginAtZero: true,
                    ticks: {
                        font: { size: 16 },
                        stepSize: 1,
                        callback: value => Math.floor(value)
                    },
                    title: {
                        display: true,
                        text: 'Total Records',
                        font: { size: 16, weight: 'bold' }
                    },
                    grid: {
                        drawOnChartArea: false
                    }
                },
                x: {
                    ticks: {
                        font: { size: 16 }
                    },
                    grid: {
                        display: false
                    }
                }
            }
        }
    });
    
    // Check final size after chart is fully rendered
    setTimeout(() => {
        const canvas = document.getElementById('lineChart');
        if (canvas && canvas.parentElement) {
            const parentWidth = canvas.parentElement.offsetWidth;
            
            // Update visible debug info - final confirmation
            const debugDiv3 = document.getElementById('canvasDimensions');
            if (debugDiv3) {
                debugDiv3.textContent = `✅ Chart Ready: Container ${parentWidth}px x 450px | Canvas ${canvas.width}x${canvas.height} (HiDPI)`;
                debugDiv3.style.color = '#388e3c';
            }
            console.log('📐 Final - Container:', parentWidth, 'x 450 | Canvas bitmap:', canvas.width, 'x', canvas.height);
        }
    }, 250);
    
    console.log('✅ Line chart created successfully');
    
    // Update status to success
    if (lineStatusDiv && lineStatusText) {
        lineStatusDiv.style.background = '#d4edda';
        lineStatusDiv.style.borderColor = '#28a745';
        lineStatusText.textContent = '✅ Trend chart created! Showing ' + labels.length + ' months.';
        
        // Hide success message after 3 seconds
        setTimeout(() => {
            lineStatusDiv.style.display = 'none';
        }, 3000);
    }
}

/**
 * Render progress bars
 */
function renderProgressBars() {
    if (appState.sheetData.length === 0) return;

    const headers = appState.sheetData[0];
    const rows = appState.sheetData.slice(1);
    let progressHTML = '';

    const statusColIndex = findColumnIndex(headers, ['status']);

    if (statusColIndex !== -1) {
        const statusCounts = {};
        rows.forEach(row => {
            const status = row[statusColIndex] || 'Unknown';
            statusCounts[status] = (statusCounts[status] || 0) + 1;
        });

        Object.entries(statusCounts).forEach(([status, count]) => {
            const percentage = ((count / rows.length) * 100).toFixed(1);
            progressHTML += `
                <div class="col-md-6">
                    <div class="card stat-card mb-3">
                        <div class="card-body">
                            <div class="d-flex justify-content-between mb-2">
                                <span class="fw-bold">${status}</span>
                                <span class="text-muted">${count} / ${rows.length}</span>
                            </div>
                            <div class="progress progress-modern">
                                <div class="progress-bar progress-bar-modern" 
                                     style="width: ${percentage}%">
                                    ${percentage}%
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            `;
        });

        document.getElementById('progressBars').innerHTML = progressHTML;
    }
}

// ============================================
// TABLE UTILITIES
// ============================================

/**
 * Filter table based on search input
 */
function filterTable() {
    const searchTerm = document.getElementById('searchBox').value.toLowerCase();
    const rows = document.getElementById('tableBody').getElementsByTagName('tr');
    let visibleCount = 0;

    Array.from(rows).forEach(row => {
        const text = row.textContent.toLowerCase();
        if (text.includes(searchTerm)) {
            row.style.display = '';
            visibleCount++;
        } else {
            row.style.display = 'none';
        }
    });

    document.getElementById('rowCount').textContent = visibleCount;
}

/**
 * Sort table by column
 */
function sortTable(columnIndex) {
    const tbody = document.getElementById('tableBody');
    const rows = Array.from(tbody.getElementsByTagName('tr'));
    
    rows.sort((a, b) => {
        const aValue = a.cells[columnIndex].textContent;
        const bValue = b.cells[columnIndex].textContent;
        
        if (!isNaN(aValue) && !isNaN(bValue)) {
            return parseFloat(aValue) - parseFloat(bValue);
        }
        
        return aValue.localeCompare(bValue);
    });

    rows.forEach(row => tbody.appendChild(row));
}

// ============================================
// EXPORT FUNCTIONALITY
// ============================================

/**
 * Export data to CSV file
 */
function exportToCSV() {
    if (appState.sheetData.length === 0) {
        showNotification('No data to export', 'warning');
        return;
    }

    let csv = appState.sheetData.map(row => 
        row.map(cell => `"${cell}"`).join(',')
    ).join('\n');

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `wam_audit_data_${new Date().toISOString().split('T')[0]}.csv`;
    link.click();
    
    window.URL.revokeObjectURL(url);
    showNotification('Data exported successfully!', 'success');
}

// ============================================
// AUTO-REFRESH
// ============================================

/**
 * Setup automatic data refresh
 */
function setupAutoRefresh() {
    if (appState.autoRefreshTimer) {
        clearInterval(appState.autoRefreshTimer);
    }
    
    appState.autoRefreshTimer = setInterval(() => {
        console.log('🔄 Auto-refresh triggered');
        fetchSheetData();
    }, CONFIG.AUTO_REFRESH_INTERVAL);
    
    console.log(`✅ Auto-refresh enabled (every ${CONFIG.AUTO_REFRESH_INTERVAL / 60000} minutes)`);
}

// ============================================
// UI UTILITIES
// ============================================

/**
 * Show/hide loading overlay
 */
function showLoading(show) {
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) {
        overlay.style.display = show ? 'flex' : 'none';
    }
}

/**
 * Update sync status indicator
 */
function updateSyncStatus(status) {
    const text = document.getElementById('syncText');
    if (!text) return;
    
    const statusConfig = {
        syncing: { html: '<span class="text-warning">Syncing...</span>', icon: 'arrow-repeat' },
        success: { html: '<span class="text-success">✓ Live</span>', icon: 'check-circle' },
        error: { html: '<span class="text-danger">✗ Error</span>', icon: 'x-circle' }
    };
    
    const config = statusConfig[status];
    if (config) {
        text.innerHTML = config.html;
    }
}

/**
 * Update last sync timestamp
 */
function updateLastSync() {
    const element = document.getElementById('lastSync');
    if (element) {
        const now = new Date();
        element.textContent = now.toLocaleTimeString();
    }
}

/**
 * Show notification message
 */
function showNotification(message, type = 'info') {
    console.log(`[${type.toUpperCase()}] ${message}`);
    // You can implement Bootstrap Toast here
}

/**
 * Show error help message
 */
function showErrorHelp(errorMessage) {
    const messageArea = document.getElementById('messageArea');
    if (!messageArea) return;
    
    let helpText = '';
    if (errorMessage.includes('API')) {
        helpText = 'Check your API key and make sure Google Sheets API is enabled.';
    } else if (errorMessage.includes('No data')) {
        helpText = 'Make sure your sheet has data and the range is correct.';
    } else {
        helpText = 'Make sure your Google Sheet is set to "Anyone with the link can view".';
    }
    
    messageArea.innerHTML = `
        <div class="alert alert-danger alert-dismissible fade show" role="alert">
            <strong>Error:</strong> ${errorMessage}
            <br><small>${helpText}</small>
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        </div>
    `;
}

// ============================================
// HELPER FUNCTIONS
// ============================================

/**
 * Find column index by keyword match
 */
function findColumnIndex(headers, keywords) {
    return headers.findIndex(h => 
        keywords.some(keyword => h.toLowerCase().includes(keyword))
    );
}

/**
 * Find first numeric column
 */
function findNumericColumn(headers, rows) {
    for (let i = 1; i < headers.length; i++) {
        if (rows.some(row => !isNaN(parseFloat(row[i])))) {
            return i;
        }
    }
    return -1;
}

/**
 * Get badge class based on percentage
 */
function getBadgeClass(percentage) {
    if (percentage >= 75) return 'bg-success';
    if (percentage >= 50) return 'bg-warning';
    return 'bg-danger';
}

/**
 * Get success label based on percentage
 */
function getSuccessLabel(percentage) {
    if (percentage >= 75) return 'Excellent';
    if (percentage >= 50) return 'Good';
    return 'Needs Work';
}

/**
 * Escape HTML to prevent XSS
 */
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

/**
 * Setup smooth scrolling for navigation
 */
function setupSmoothScroll() {
    document.querySelectorAll('a[href^="#"]').forEach(anchor => {
        anchor.addEventListener('click', function (e) {
            e.preventDefault();
            const target = document.querySelector(this.getAttribute('href'));
            if (target) {
                target.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
        });
    });
}

// ============================================
// EXPOSE FUNCTIONS GLOBALLY
// ============================================

window.fetchSheetData = fetchSheetData;
window.refreshData = refreshData;
window.filterTable = filterTable;
window.sortTable = sortTable;
window.exportToCSV = exportToCSV;
window.closeOfficeDetails = closeOfficeDetails;
window.exportOfficeData = exportOfficeData;
window.filterByOffice = filterByOffice;

// ============================================
// COLUMN DETECTION & FILTERING
// ============================================

/**
 * Auto-detect column indices from headers
 */
function detectColumnIndices() {
    if (appState.sheetData.length < 2) return;
    
    // Check if first row is empty (title row), use second row as headers
    const firstRow = appState.sheetData[0];
    const secondRow = appState.sheetData[1];
    
    let headers = firstRow;
    let headerRowIndex = 0;
    
    // If first row is mostly empty, use second row
    const emptyCount = firstRow.filter(cell => !cell || cell.trim() === '').length;
    if (emptyCount > firstRow.length / 2 && secondRow && secondRow.length > 0) {
        headers = secondRow;
        headerRowIndex = 1;
        console.log('📋 Using row 2 as headers (row 1 is title/empty)');
    }
    
    console.log('📋 Detected columns:', headers);
    console.log('📋 Headers array:', headers.map((h, i) => `Col ${i}: "${h}"`));
    
    // FORCE: Column A (index 0) is Office Name
    appState.columnIndices.office = 0;
    console.log('✅ Force set office to Column A (index 0)');
    
    // Detect Personnel/Name column (Column B)
    appState.columnIndices.personnel = headers.findIndex(h => 
        h && (h.toLowerCase().includes('name') || 
        h.toLowerCase().includes('personnel') ||
        h.toLowerCase().includes('employee'))
    );
    
    // If personnel not found, use Column B
    if (appState.columnIndices.personnel === -1) {
        appState.columnIndices.personnel = 1;
        console.log('✅ Force set personnel to Column B (index 1)');
    }
    
    // Detect Date column
    appState.columnIndices.date = headers.findIndex(h => 
        h && (h.toLowerCase().includes('date') || 
        h.toLowerCase().includes('audit'))
    );
    
    // If date not found, use Column J (index 9)
    if (appState.columnIndices.date === -1) {
        appState.columnIndices.date = 9; // Column J (Audit Date)
        console.log('✅ Force set date to Column J (index 9)');
    }
    
    // Column E is index 4 (0-indexed: A=0, B=1, C=2, D=3, E=4)
    appState.columnIndices.rockReview = 4;
    
    // Store header row index for later use
    appState.headerRowIndex = headerRowIndex;
    
    console.log('✅ Column Detection:', {
        headerRow: headerRowIndex + 1,
        personnel: appState.columnIndices.personnel >= 0 ? `Column ${String.fromCharCode(65 + appState.columnIndices.personnel)}: ${headers[appState.columnIndices.personnel]}` : 'Not found',
        office: appState.columnIndices.office >= 0 ? `Column ${String.fromCharCode(65 + appState.columnIndices.office)}: ${headers[appState.columnIndices.office]}` : 'Not found',
        date: appState.columnIndices.date >= 0 ? `Column ${String.fromCharCode(65 + appState.columnIndices.date)}: ${headers[appState.columnIndices.date]}` : 'Not found',
        rockReview: `Column E (index 4): ${headers[4] || 'ROCK REVIEW'}`
    });
}

/**
 * Render filter controls
 */
function renderFilters() {
    if (appState.sheetData.length === 0) return;
    
    const headers = appState.sheetData[appState.headerRowIndex];
    const dataStartRow = appState.headerRowIndex + 1;
    const rows = appState.sheetData.slice(dataStartRow);
    
    // Get unique personnel names
    const personnelSet = new Set();
    if (appState.columnIndices.personnel >= 0) {
        rows.forEach(row => {
            const name = row[appState.columnIndices.personnel];
            if (name && name.trim()) personnelSet.add(name.trim());
        });
    }
    const personnelList = Array.from(personnelSet).sort();
    
    // Get unique offices
    const officeSet = new Set();
    if (appState.columnIndices.office >= 0) {
        rows.forEach(row => {
            const office = row[appState.columnIndices.office];
            if (office && office.trim()) officeSet.add(office.trim());
        });
    }
    const officeList = Array.from(officeSet).sort();
    
    const filterHTML = `
        <div class="filter-section">
            <div class="row g-3">
                <div class="col-md-4">
                    <label class="filter-label">
                        <i class="bi bi-calendar-range"></i> Time Period
                    </label>
                    <select class="filter-select" id="timePeriodFilter" onchange="applyFilters()">
                        <option value="all">All Time</option>
                        <option value="week">This Week</option>
                        <option value="month">This Month</option>
                        <option value="year">This Year</option>
                    </select>
                </div>
                ${appState.columnIndices.personnel >= 0 ? `
                <div class="col-md-4">
                    <label class="filter-label">
                        <i class="bi bi-person"></i> Personnel
                    </label>
                    <select class="filter-select" id="personnelFilter" onchange="applyFilters()">
                        <option value="all">All Personnel</option>
                        ${personnelList.map(name => `<option value="${name}">${name}</option>`).join('')}
                    </select>
                </div>
                ` : ''}
                ${appState.columnIndices.office >= 0 ? `
                <div class="col-md-4">
                    <label class="filter-label">
                        <i class="bi bi-building"></i> Office
                    </label>
                    <select class="filter-select" id="officeFilter" onchange="applyFilters()">
                        <option value="all">All Offices</option>
                        ${officeList.map(office => `<option value="${office}">${office}</option>`).join('')}
                    </select>
                </div>
                ` : ''}
            </div>
        </div>
    `;
    
    // Insert filters before overview section
    const overviewSection = document.getElementById('overview');
    if (overviewSection && !document.getElementById('filtersContainer')) {
        const filtersDiv = document.createElement('div');
        filtersDiv.id = 'filtersContainer';
        filtersDiv.innerHTML = filterHTML;
        overviewSection.parentNode.insertBefore(filtersDiv, overviewSection);
    }
}

/**
 * Apply filters to data
 */
function applyFilters() {
    const dataStartRow = appState.headerRowIndex + 1;
    const rows = appState.sheetData.slice(dataStartRow);
    let filtered = [...rows];
    
    // Get filter values
    const timePeriod = document.getElementById('timePeriodFilter')?.value || 'all';
    const personnel = document.getElementById('personnelFilter')?.value || 'all';
    const office = document.getElementById('officeFilter')?.value || 'all';
    
    // Apply personnel filter
    if (personnel !== 'all' && appState.columnIndices.personnel >= 0) {
        filtered = filtered.filter(row => 
            row[appState.columnIndices.personnel] === personnel
        );
    }
    
    // Apply office filter
    if (office !== 'all' && appState.columnIndices.office >= 0) {
        filtered = filtered.filter(row => 
            row[appState.columnIndices.office] === office
        );
    }
    
    // Apply date filter
    if (timePeriod !== 'all' && appState.columnIndices.date >= 0) {
        const now = new Date();
        filtered = filtered.filter(row => {
            const dateStr = row[appState.columnIndices.date];
            if (!dateStr) return false;
            
            const rowDate = parseDate(dateStr);
            if (!rowDate) return false;
            
            switch(timePeriod) {
                case 'week':
                    const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
                    return rowDate >= weekAgo;
                case 'month':
                    return rowDate.getMonth() === now.getMonth() && 
                           rowDate.getFullYear() === now.getFullYear();
                case 'year':
                    return rowDate.getFullYear() === now.getFullYear();
                default:
                    return true;
            }
        });
    }
    
    console.log(`🔍 Filtered: ${filtered.length} of ${rows.length} records`);
    appState.filteredData = filtered;
    
    // Re-render dashboard with filtered data
    renderDashboard();
}

/**
 * Parse date from various formats
 */
function parseDate(dateStr) {
    if (!dateStr) return null;
    
    // Try ISO format
    let date = new Date(dateStr);
    if (!isNaN(date.getTime())) return date;
    
    // Try MM/DD/YYYY or DD/MM/YYYY
    const parts = dateStr.split('/');
    if (parts.length === 3) {
        date = new Date(parts[2], parts[0] - 1, parts[1]);
        if (!isNaN(date.getTime())) return date;
    }
    
    return null;
}

/**
 * Create Circular Progress Ring
 */
function createCircularProgress(elementId, percentage, color, label, count, total) {
    const element = document.getElementById(elementId);
    if (!element) return;
    
    const radius = 90;
    const circumference = 2 * Math.PI * radius;
    const offset = circumference - (percentage / 100) * circumference;
    
    const icon = color === '#10b981' ? 'bi-check-circle-fill' : 'bi-exclamation-circle-fill';
    
    element.innerHTML = `
        <svg viewBox="0 0 200 200">
            <circle class="circle-bg" cx="100" cy="100" r="${radius}"></circle>
            <circle class="circle-progress ${color === '#10b981' ? 'on-track' : 'off-track'}" 
                    cx="100" cy="100" r="${radius}"
                    style="stroke-dasharray: ${circumference}; stroke-dashoffset: ${circumference};"></circle>
        </svg>
        <div class="percentage-text" style="color: ${color}">
            <span class="number">${percentage}%</span>
            <span class="label">
                <i class="bi ${icon}"></i> ${label}
            </span>
        </div>
        <div class="text-center mt-3" style="color: #6b7280; font-size: 0.95rem; font-weight: 500;">
            ${count} of ${total} records
        </div>
    `;
    
    // Animate the circle
    setTimeout(() => {
        const circle = element.querySelector('.circle-progress');
        if (circle) {
            circle.style.strokeDashoffset = offset;
        }
    }, 100);
}

/**
 * Create Rock Review Analysis Chart
 */
function createRockReviewChart() {
    console.log('🎯 Creating Rock Review Chart...');
    
    // Use filteredData if available, otherwise use all data (minus header)
    let rows = appState.filteredData.length > 0 ? appState.filteredData : appState.sheetData.slice(1);
    
    // Filter out header/title rows
    const officeColIndex = appState.columnIndices.office;
    rows = rows.filter(row => {
        const officeName = (row[officeColIndex] || '').trim().toUpperCase();
        return officeName !== 'OFFICE NAME' && 
               !officeName.includes('WEEKLY ACCOUNTABILITY') &&
               !officeName.includes('WAM') && 
               officeName !== '';
    });
    
    if (rows.length === 0) {
        console.log('⚠️ No data available for Rock Review chart');
        return;
    }
    
    const rockReviewIndex = appState.columnIndices.rockReview;
    console.log('📍 Rock Review Column Index:', rockReviewIndex);
    
    let onTrack = 0;
    let offTrack = 0;
    let other = 0;
    
    rows.forEach((row, idx) => {
        const value = (row[rockReviewIndex] || '').toLowerCase().trim();
        if (idx < 5) {
            console.log(`Row ${idx} Column E value: "${row[rockReviewIndex]}"`);
        }
        
        if (value.includes('on') && value.includes('track') && !value.includes('off')) {
            onTrack++;
        } else if (value.includes('off') && value.includes('track')) {
            offTrack++;
        } else if (value) {
            other++;
        }
    });
    
    const total = onTrack + offTrack + other;
    const onTrackPercent = total > 0 ? ((onTrack / total) * 100).toFixed(1) : 0;
    const offTrackPercent = total > 0 ? ((offTrack / total) * 100).toFixed(1) : 0;
    const otherPercent = total > 0 ? ((other / total) * 100).toFixed(1) : 0;
    
    console.log('📊 Rock Review Analysis:', { onTrack, offTrack, other, total });
    
    // Create circular progress rings with record counts
    createCircularProgress('onTrackCircle', onTrackPercent, '#10b981', 'On-Track', onTrack, total);
    createCircularProgress('offTrackCircle', offTrackPercent, '#ef4444', 'Off-Track', offTrack, total);
    
    // Optional: Create doughnut chart if canvas exists
    const canvas = document.getElementById('rockReviewChart');
    if (!canvas) {
        console.log('✅ Circular progress rings created. Doughnut chart canvas not found (optional).');
        return;
    }
    
    if (appState.charts.rockReviewChart) {
        appState.charts.rockReviewChart.destroy();
    }
    
    appState.charts.rockReviewChart = new Chart(canvas, {
        type: 'doughnut',
        data: {
            labels: [`On-Track (${onTrackPercent}%)`, `Off-Track (${offTrackPercent}%)`, other > 0 ? `Other (${otherPercent}%)` : ''],
            datasets: [{
                data: [onTrack, offTrack, other],
                backgroundColor: [
                    '#10b981',   // Solid Green for On-Track
                    '#ef4444',   // Solid Red for Off-Track
                    '#9ca3af'    // Gray for Other
                ],
                borderColor: '#ffffff',
                borderWidth: 4,
                hoverOffset: 20
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            cutout: '70%',  // Makes the doughnut thicker
            plugins: {
                legend: {
                    display: true,
                    position: 'bottom',
                    labels: {
                        padding: 20,
                        font: { 
                            size: 14,
                            weight: 'bold'
                        },
                        usePointStyle: true,
                        pointStyle: 'circle',
                        generateLabels: function(chart) {
                            const data = chart.data;
                            return data.labels.map((label, i) => {
                                if (data.datasets[0].data[i] === 0) return null;
                                return {
                                    text: label,
                                    fillStyle: data.datasets[0].backgroundColor[i],
                                    hidden: false,
                                    index: i
                                };
                            }).filter(item => item !== null);
                        }
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.9)',
                    titleColor: '#ffffff',
                    bodyColor: '#ffffff',
                    padding: 15,
                    cornerRadius: 8,
                    displayColors: true,
                    bodyFont: {
                        size: 14
                    },
                    callbacks: {
                        label: function(context) {
                            const label = context.label || '';
                            const value = context.parsed;
                            return `${label}: ${value} records`;
                        }
                    }
                }
            },
            layout: {
                padding: 20
            }
        },
        plugins: [{
            id: 'centerText',
            afterDatasetsDraw: function(chart) {
                const { ctx, chartArea: { width, height } } = chart;
                
                ctx.save();
                
                // Calculate center position
                const centerX = width / 2;
                const centerY = height / 2;
                
                // Draw On-Track percentage (large, centered)
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                ctx.fillStyle = '#10b981';
                ctx.font = 'bold 48px sans-serif';
                ctx.fillText(`${onTrackPercent}%`, centerX, centerY - 15);
                
                // Draw "On-Track" label
                ctx.fillStyle = '#6b7280';
                ctx.font = '16px sans-serif';
                ctx.fillText('On-Track', centerX, centerY + 15);
                
                // Draw Off-Track percentage (smaller, below)
                ctx.fillStyle = '#ef4444';
                ctx.font = 'bold 20px sans-serif';
                ctx.fillText(`${offTrackPercent}% Off-Track`, centerX, centerY + 40);
                
                ctx.restore();
            }
        }]
    });
    
    // Summary card is now replaced by circular progress rings above
}

// ============================================
// EXPOSE NEW FUNCTIONS GLOBALLY
// ============================================

window.applyFilters = applyFilters;

// ============================================
// TEST FUNCTION FOR BAR CHART
// ============================================

window.testBarChart = function() {
    console.log('🧪 Manual test chart creation triggered');
    const canvas = document.getElementById('barChart');
    
    if (!canvas) {
        alert('❌ Canvas not found!');
        return;
    }
    
    // Destroy existing
    if (appState.charts.barChart) {
        appState.charts.barChart.destroy();
    }
    
    // Create simple test chart
    try {
        appState.charts.barChart = new Chart(canvas, {
            type: 'bar',
            data: {
                labels: ['GSO', 'CEO', 'SWMO', 'Finance', 'HR'],
                datasets: [{
                    label: 'Test Data',
                    data: [12, 8, 6, 5, 3],
                    backgroundColor: '#10b981'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                indexAxis: 'y'
            }
        });
        console.log('✅ Test chart created!', appState.charts.barChart);
        alert('✅ Test chart created! Check the green box.');
    } catch (error) {
        console.error('❌ Test chart error:', error);
        alert('❌ Error: ' + error.message);
    }
};

// ============================================
// TABLE ENHANCEMENT FUNCTIONS
// ============================================

function renderPaginationControls(totalPages) {
    const paginationContainer = document.getElementById('paginationControls');
    if (!paginationContainer) return;
    
    const currentPage = appState.pagination.currentPage;
    let html = '';
    
    // Previous button
    html += `<li class="page-item ${currentPage === 1 ? 'disabled' : ''}">
                <a class="page-link" href="#" onclick="changePage(${currentPage - 1}); return false;">
                    <i class="bi bi-chevron-left"></i>
                </a>
             </li>`;
    
    // Page numbers (show max 5 pages)
    const maxPagesToShow = 5;
    let startPage = Math.max(1, currentPage - Math.floor(maxPagesToShow / 2));
    let endPage = Math.min(totalPages, startPage + maxPagesToShow - 1);
    
    if (endPage - startPage + 1 < maxPagesToShow) {
        startPage = Math.max(1, endPage - maxPagesToShow + 1);
    }
    
    if (startPage > 1) {
        html += `<li class="page-item"><a class="page-link" href="#" onclick="changePage(1); return false;">1</a></li>`;
        if (startPage > 2) html += `<li class="page-item disabled"><span class="page-link">...</span></li>`;
    }
    
    for (let i = startPage; i <= endPage; i++) {
        html += `<li class="page-item ${i === currentPage ? 'active' : ''}">
                    <a class="page-link" href="#" onclick="changePage(${i}); return false;">${i}</a>
                 </li>`;
    }
    
    if (endPage < totalPages) {
        if (endPage < totalPages - 1) html += `<li class="page-item disabled"><span class="page-link">...</span></li>`;
        html += `<li class="page-item"><a class="page-link" href="#" onclick="changePage(${totalPages}); return false;">${totalPages}</a></li>`;
    }
    
    // Next button
    html += `<li class="page-item ${currentPage === totalPages ? 'disabled' : ''}">
                <a class="page-link" href="#" onclick="changePage(${currentPage + 1}); return false;">
                    <i class="bi bi-chevron-right"></i>
                </a>
             </li>`;
    
    paginationContainer.innerHTML = html;
}

function changePage(page) {
    const totalPages = appState.pagination.totalPages;
    if (page < 1 || page > totalPages) return;
    
    appState.pagination.currentPage = page;
    renderTablePage();
    
    // Scroll to top of table
    document.getElementById('data').scrollIntoView({ behavior: 'smooth' });
}

function changeRowsPerPage() {
    const select = document.getElementById('rowsPerPage');
    const value = select.value;
    appState.pagination.rowsPerPage = value === 'all' ? 'all' : parseInt(value);
    appState.pagination.currentPage = 1; // Reset to first page
    renderTablePage();
}

function toggleColumnSelector() {
    const selector = document.getElementById('columnSelector');
    if (selector.style.display === 'none') {
        selector.style.display = 'block';
    } else {
        selector.style.display = 'none';
    }
}

function toggleColumn(columnIndex) {
    const index = appState.visibleColumns.indexOf(columnIndex);
    if (index > -1) {
        // Remove column
        appState.visibleColumns.splice(index, 1);
    } else {
        // Add column
        appState.visibleColumns.push(columnIndex);
        appState.visibleColumns.sort((a, b) => a - b); // Keep in order
    }
    renderTablePage();
}

function exportFilteredData() {
    const headers = appState.sheetData[2]; // Use correct header row
    const rows = appState.allTableData;
    
    // Get visible headers
    const visibleHeaders = headers.filter((_, index) => appState.visibleColumns.includes(index));
    
    // Create CSV content with only visible columns
    let csvContent = visibleHeaders.map(h => `"${h}"`).join(',') + '\n';
    
    rows.forEach(row => {
        const visibleCells = row.filter((_, index) => appState.visibleColumns.includes(index));
        csvContent += visibleCells.map(cell => `"${(cell || '').toString().replace(/"/g, '""')}"`).join(',') + '\n';
    });
    
    // Create download link
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', `wam_filtered_data_${new Date().toISOString().split('T')[0]}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    console.log(`✅ Exported ${rows.length} filtered records to CSV`);
}

// Make functions globally accessible
window.renderPaginationControls = renderPaginationControls;
window.changePage = changePage;
window.changeRowsPerPage = changeRowsPerPage;
window.toggleColumnSelector = toggleColumnSelector;
window.toggleColumn = toggleColumn;
window.exportFilteredData = exportFilteredData;

// ============================================
// LOG INFO
// ============================================

console.log('%c WAM Audit Dashboard ', 'background: #2563eb; color: white; padding: 5px; border-radius: 3px;');
console.log('Version: 3.1 - Enhanced Table Features');
console.log('Auto-refresh:', CONFIG.AUTO_REFRESH_INTERVAL / 60000, 'minutes');
console.log('Sheet ID:', CONFIG.SHEET_ID);

// ============================================
// ADVANCED FILTERING FUNCTIONS
// ============================================

/**
 * Initialize filter dropdowns with unique values
 */
function initializeFilters() {
    if (!appState.sheetData || appState.sheetData.length < 4) return;
    
    const headers = appState.sheetData[2]; // Row 3 contains headers
    const dataRows = appState.sheetData.slice(3); // Data starts from row 4
    
    // Get unique offices
    const officeIndex = 0; // Column A
    const uniqueOffices = [...new Set(dataRows.map(row => row[officeIndex]).filter(Boolean))];
    
    // Get unique personnel
    const personnelIndex = 1; // Column B
    const uniquePersonnel = [...new Set(dataRows.map(row => row[personnelIndex]).filter(Boolean))];
    
    // Populate Office dropdown
    const officeSelect = document.getElementById('officeFilter');
    if (officeSelect) {
        officeSelect.innerHTML = '<option value="">All Offices</option>';
        uniqueOffices.sort().forEach(office => {
            officeSelect.innerHTML += `<option value="${office}">${office}</option>`;
        });
    }
    
    // Populate Personnel dropdown
    const personnelSelect = document.getElementById('personnelFilter');
    if (personnelSelect) {
        personnelSelect.innerHTML = '<option value="">All Personnel</option>';
        uniquePersonnel.sort().forEach(person => {
            personnelSelect.innerHTML += `<option value="${person}">${person}</option>`;
        });
    }
    
    console.log('Filters initialized:', uniqueOffices.length, 'offices,', uniquePersonnel.length, 'personnel');
}

/**
 * Apply filters to the data
 */
function applyFilters() {
    if (!appState.sheetData || appState.sheetData.length < 4) {
        showMessage('No data to filter', 'warning');
        return;
    }
    
    const dateFrom = document.getElementById('dateFrom')?.value;
    const dateTo = document.getElementById('dateTo')?.value;
    const selectedOffices = Array.from(document.getElementById('officeFilter')?.selectedOptions || [])
        .map(opt => opt.value).filter(Boolean);
    const selectedPersonnel = Array.from(document.getElementById('personnelFilter')?.selectedOptions || [])
        .map(opt => opt.value).filter(Boolean);
    const selectedStatus = document.getElementById('statusFilter')?.value;
    
    const headers = appState.sheetData[2];
    let dataRows = appState.sheetData.slice(3);
    
    // Filter by date range
    if (dateFrom || dateTo) {
        const dateIndex = 9; // Column J - Audit Date
        dataRows = dataRows.filter(row => {
            const rowDate = parseDate(row[dateIndex]);
            if (!rowDate) return false;
            
            if (dateFrom && dateTo) {
                return rowDate >= new Date(dateFrom) && rowDate <= new Date(dateTo);
            } else if (dateFrom) {
                return rowDate >= new Date(dateFrom);
            } else if (dateTo) {
                return rowDate <= new Date(dateTo);
            }
            return true;
        });
    }
    
    // Filter by office
    if (selectedOffices.length > 0) {
        const officeIndex = 0;
        dataRows = dataRows.filter(row => selectedOffices.includes(row[officeIndex]));
    }
    
    // Filter by personnel
    if (selectedPersonnel.length > 0) {
        const personnelIndex = 1;
        dataRows = dataRows.filter(row => selectedPersonnel.includes(row[personnelIndex]));
    }
    
    // Filter by status
    if (selectedStatus) {
        const statusIndex = 4; // Column E - Rock Review
        dataRows = dataRows.filter(row => {
            const status = row[statusIndex]?.toString().trim();
            return status === selectedStatus;
        });
    }
    
    // Update filtered data
    appState.filteredData = [appState.sheetData[0], appState.sheetData[1], headers, ...dataRows];
    
    // Update UI
    updateActiveFiltersDisplay();
    renderDashboardWithFilteredData();
    
    showMessage(`Filters applied: ${dataRows.length} records found`, 'success');
}

/**
 * Quick filter presets
 */
function quickFilter(type) {
    const today = new Date();
    const dateFrom = document.getElementById('dateFrom');
    const dateTo = document.getElementById('dateTo');
    const statusFilter = document.getElementById('statusFilter');
    
    // Reset filters first
    resetFilters(false);
    
    switch(type) {
        case 'today':
            const todayStr = today.toISOString().split('T')[0];
            dateFrom.value = todayStr;
            dateTo.value = todayStr;
            break;
            
        case 'week':
            const weekAgo = new Date(today);
            weekAgo.setDate(today.getDate() - 7);
            dateFrom.value = weekAgo.toISOString().split('T')[0];
            dateTo.value = today.toISOString().split('T')[0];
            break;
            
        case 'month':
            const monthAgo = new Date(today);
            monthAgo.setMonth(today.getMonth() - 1);
            dateFrom.value = monthAgo.toISOString().split('T')[0];
            dateTo.value = today.toISOString().split('T')[0];
            break;
            
        case 'offtrack':
            statusFilter.value = 'Off-Track';
            break;
    }
    
    applyFilters();
}

/**
 * Reset all filters
 */
function resetFilters(refresh = true) {
    document.getElementById('dateFrom').value = '';
    document.getElementById('dateTo').value = '';
    document.getElementById('officeFilter').selectedIndex = 0;
    document.getElementById('personnelFilter').selectedIndex = 0;
    document.getElementById('statusFilter').selectedIndex = 0;
    
    appState.filteredData = [];
    document.getElementById('activeFilters').style.display = 'none';
    
    if (refresh) {
        renderDashboardWithFilteredData();
        showMessage('Filters reset', 'info');
    }
}

/**
 * Update active filters display
 */
function updateActiveFiltersDisplay() {
    const activeFiltersDiv = document.getElementById('activeFilters');
    const filterTagsDiv = document.getElementById('filterTags');
    const tags = [];
    
    const dateFrom = document.getElementById('dateFrom')?.value;
    const dateTo = document.getElementById('dateTo')?.value;
    const selectedOffices = Array.from(document.getElementById('officeFilter')?.selectedOptions || [])
        .map(opt => opt.text).filter(t => t !== 'All Offices');
    const selectedPersonnel = Array.from(document.getElementById('personnelFilter')?.selectedOptions || [])
        .map(opt => opt.text).filter(t => t !== 'All Personnel');
    const selectedStatus = document.getElementById('statusFilter')?.value;
    
    if (dateFrom) tags.push(`From: ${dateFrom}`);
    if (dateTo) tags.push(`To: ${dateTo}`);
    if (selectedOffices.length > 0) tags.push(`Office: ${selectedOffices.join(', ')}`);
    if (selectedPersonnel.length > 0) tags.push(`Personnel: ${selectedPersonnel.join(', ')}`);
    if (selectedStatus) tags.push(`Status: ${selectedStatus}`);
    
    if (tags.length > 0) {
        filterTagsDiv.innerHTML = tags.map(tag => 
            `<span class="badge bg-primary me-1">${tag}</span>`
        ).join('');
        activeFiltersDiv.style.display = 'block';
    } else {
        activeFiltersDiv.style.display = 'none';
    }
}

/**
 * Render dashboard with filtered data
 */
function renderDashboardWithFilteredData() {
    const dataToUse = appState.filteredData.length > 0 ? appState.filteredData : appState.sheetData;
    
    // Temporarily swap data
    const originalData = appState.sheetData;
    appState.sheetData = dataToUse;
    
    // Re-render all components
    renderSummaryCards();
    renderRockReviewCharts();
    createOfficeDistributionChart();
    createOfficePerformanceChart();
    createTrendLineChart();
    renderDataTable();
    
    // Restore original data
    appState.sheetData = originalData;
}

/**
 * Helper function to parse dates
 */
function parseDate(dateStr) {
    if (!dateStr) return null;
    
    // Try MM/DD/YYYY format
    const parts = dateStr.toString().split('/');
    if (parts.length === 3) {
        const month = parseInt(parts[0]) - 1;
        const day = parseInt(parts[1]);
        const year = parseInt(parts[2]);
        return new Date(year, month, day);
    }
    
    // Try standard date parse
    const date = new Date(dateStr);
    return isNaN(date.getTime()) ? null : date;
}

// ============================================
// PDF EXPORT FUNCTIONS
// ============================================

/**
 * Export dashboard to PDF with charts
 */
async function exportToPDF() {
    showMessage('Generating PDF... This may take a moment', 'info');
    
    try {
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF('p', 'mm', 'a4');
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        let yPosition = 20;
        
        // Add MyNaga App Logo and Title
        pdf.setFontSize(20);
        pdf.setTextColor(30, 58, 138);
        pdf.text('MyNaga App WAM Dashboard', pageWidth / 2, yPosition, { align: 'center' });
        
        yPosition += 10;
        pdf.setFontSize(12);
        pdf.setTextColor(100, 100, 100);
        pdf.text('Weekly Accountability Meeting Report', pageWidth / 2, yPosition, { align: 'center' });
        
        yPosition += 5;
        pdf.setFontSize(10);
        const exportDate = new Date().toLocaleString();
        pdf.text(`Generated: ${exportDate}`, pageWidth / 2, yPosition, { align: 'center' });
        
        yPosition += 15;
        
        // Add Summary Cards
        pdf.setFontSize(14);
        pdf.setTextColor(30, 58, 138);
        pdf.text('Dashboard Overview', 15, yPosition);
        yPosition += 10;
        
        const summaryCards = document.querySelectorAll('#summaryCards .card');
        pdf.setFontSize(10);
        pdf.setTextColor(0, 0, 0);
        
        let cardX = 15;
        summaryCards.forEach((card, index) => {
            const title = card.querySelector('h6')?.textContent || '';
            const value = card.querySelector('h3')?.textContent || '';
            
            pdf.setFillColor(240, 249, 255);
            pdf.rect(cardX, yPosition - 5, 45, 15, 'F');
            pdf.setFontSize(8);
            pdf.setTextColor(100, 100, 100);
            pdf.text(title, cardX + 2, yPosition);
            pdf.setFontSize(14);
            pdf.setTextColor(30, 58, 138);
            pdf.text(value, cardX + 2, yPosition + 8);
            
            cardX += 47;
            if ((index + 1) % 4 === 0) {
                cardX = 15;
                yPosition += 20;
            }
        });
        
        if (summaryCards.length % 4 !== 0) {
            yPosition += 20;
        }
        
        // Capture and add charts
        yPosition += 10;
        
        // Pie Chart
        const pieChartElement = document.getElementById('pieChart');
        if (pieChartElement) {
            pdf.setFontSize(12);
            pdf.setTextColor(30, 58, 138);
            pdf.text('Office Distribution', 15, yPosition);
            yPosition += 10;
            
            const pieCanvas = await html2canvas(pieChartElement.parentElement, { 
                scale: 2, 
                backgroundColor: '#ffffff'
            });
            const pieImgData = pieCanvas.toDataURL('image/png');
            pdf.addImage(pieImgData, 'PNG', 15, yPosition, 85, 60);
            yPosition += 70;
        }
        
        // Add new page for more charts
        if (yPosition > pageHeight - 80) {
            pdf.addPage();
            yPosition = 20;
        }
        
        // Bar Chart
        const barChartElement = document.getElementById('barChart');
        if (barChartElement) {
            pdf.setFontSize(12);
            pdf.setTextColor(30, 58, 138);
            pdf.text('Office Performance Comparison', 15, yPosition);
            yPosition += 10;
            
            const barCanvas = await html2canvas(barChartElement.parentElement, { 
                scale: 2,
                backgroundColor: '#ffffff'
            });
            const barImgData = barCanvas.toDataURL('image/png');
            pdf.addImage(barImgData, 'PNG', 15, yPosition, 180, 70);
            yPosition += 80;
        }
        
        // Add new page for line chart
        pdf.addPage();
        yPosition = 20;
        
        // Line Chart
        const lineChartElement = document.getElementById('lineChart');
        if (lineChartElement) {
            pdf.setFontSize(12);
            pdf.setTextColor(30, 58, 138);
            pdf.text('Performance Trend Over Time', 15, yPosition);
            yPosition += 10;
            
            const lineCanvas = await html2canvas(lineChartElement.parentElement.parentElement, { 
                scale: 2,
                backgroundColor: '#ffffff'
            });
            const lineImgData = lineCanvas.toDataURL('image/png');
            pdf.addImage(lineImgData, 'PNG', 15, yPosition, 180, 100);
            yPosition += 110;
        }
        
        // Add filtered data summary if filters are active
        if (appState.filteredData.length > 0) {
            pdf.addPage();
            yPosition = 20;
            pdf.setFontSize(14);
            pdf.setTextColor(30, 58, 138);
            pdf.text('Active Filters', 15, yPosition);
            yPosition += 10;
            
            pdf.setFontSize(10);
            pdf.setTextColor(0, 0, 0);
            
            const filterTags = document.querySelectorAll('#filterTags .badge');
            filterTags.forEach(tag => {
                pdf.text('• ' + tag.textContent, 20, yPosition);
                yPosition += 7;
            });
        }
        
        // Footer on last page
        pdf.setFontSize(8);
        pdf.setTextColor(150, 150, 150);
        pdf.text('MyNaga App WAM Dashboard - City Government of Naga', pageWidth / 2, pageHeight - 10, { align: 'center' });
        pdf.text('Powered by MyNaga App', pageWidth / 2, pageHeight - 6, { align: 'center' });
        
        // Save PDF
        const filename = `MyNaga_WAM_Report_${new Date().toISOString().split('T')[0]}.pdf`;
        pdf.save(filename);
        
        showMessage('PDF exported successfully!', 'success');
        
    } catch (error) {
        console.error('PDF export error:', error);
        showMessage('Error generating PDF: ' + error.message, 'danger');
    }
}
