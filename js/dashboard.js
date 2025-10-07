document.addEventListener('DOMContentLoaded', () => {
    // --- GLOBAL VARIABLES ---
    let overviewData = [];
    let activityModal;
    let statusChartInstance = null;
    let topicChartInstance = null;
    const storageKey = 'overviewData';

    // --- INITIALIZATION ---
    function initialize() {
        const modalElement = document.getElementById('activityModal');
        if (modalElement) {
            activityModal = new bootstrap.Modal(modalElement);
        }
        
        loadDataAndRender();
        addEventListeners();
    }

    // --- DATA HANDLING ---
    function loadDataAndRender() {
        const storedData = localStorage.getItem(storageKey);
        if (!storedData) {
            console.warn('No overview data found. Redirecting to upload page.');
            window.location.href = 'index.html';
            return;
        }

        try {
            overviewData = JSON.parse(storedData);
        } catch (e) {
            console.error("Failed to parse overview data from localStorage.", e);
            overviewData = [];
        }
        
        renderDashboard();
    }
    
    function saveDataAndReRender() {
        localStorage.setItem(storageKey, JSON.stringify(overviewData));
        renderDashboard();
    }

    // --- RENDERING ---
    function renderDashboard() {
        renderSummaryCards(overviewData);
        renderCharts(overviewData);
        renderTable(overviewData);
    }

    function renderSummaryCards(data) {
        const statusCounts = data.reduce((acc, item) => {
            const status = (item.Status || 'open').toLowerCase();
            acc[status] = (acc[status] || 0) + 1;
            return acc;
        }, { open: 0, done: 0 });

        const uniqueTopics = new Set(data.map(item => item.Topics)).size;

        document.getElementById('total-activities').textContent = data.length;
        document.getElementById('status-done').textContent = statusCounts.done;
        document.getElementById('status-open').textContent = statusCounts.open;
        document.getElementById('unique-topics').textContent = uniqueTopics;
    }

    function renderCharts(data) {
        const statusCtx = document.getElementById('status-chart')?.getContext('2d');
        const topicCtx = document.getElementById('topic-chart')?.getContext('2d');
        if (!statusCtx || !topicCtx) return;
        
        // Status Chart
        const statusCounts = data.reduce((acc, item) => {
            const status = (item.Status || 'open').toLowerCase();
            acc[status] = (acc[status] || 0) + 1;
            return acc;
        }, {});
        if (statusChartInstance) statusChartInstance.destroy();
        statusChartInstance = new Chart(statusCtx, {
            type: 'doughnut',
            data: {
                labels: Object.keys(statusCounts),
                datasets: [{ 
                    data: Object.values(statusCounts),
                    backgroundColor: ['#dc3545', '#198754', '#ffc107']
                }]
            },
            options: { responsive: true, maintainAspectRatio: true }
        });

        // Topic Chart
        const topicCounts = data.reduce((acc, item) => {
            const topic = item.Topics || 'Unassigned';
            acc[topic] = (acc[topic] || 0) + 1;
            return acc;
        }, {});
        if (topicChartInstance) topicChartInstance.destroy();
        topicChartInstance = new Chart(topicCtx, {
            type: 'bar',
            data: {
                labels: Object.keys(topicCounts),
                datasets: [{
                    label: '# of Activities',
                    data: Object.values(topicCounts),
                    backgroundColor: '#0d6efd'
                }]
            },
            options: { responsive: true, maintainAspectRatio: true, plugins: { legend: { display: false } } }
        });
    }

    function renderTable(data) {
        const tableBody = document.getElementById('overview-table-body');
        if (!tableBody) return;
        tableBody.innerHTML = '';

        if (data.length === 0) {
            tableBody.innerHTML = '<tr><td colspan="9" class="text-center">No data available.</td></tr>';
            return;
        }

        data.forEach((item, index) => {
            const status = (item.Status || 'open').toLowerCase();
            const statusClass = status === 'done' ? 'text-success' : 'text-danger';
            const row = `
                <tr>
                    <td>${item.Topics || ''}</td>
                    <td>${item.Activity || ''}</td>
                    <td>${item.Responsible || ''}</td>
                    <td>${item['Target Date'] || ''}</td>
                    <td><strong class="${statusClass}">${status}</strong></td>
                    <td>${item.Comment || ''}</td>
                    <td>${item.Environment || ''}</td>
                    <td>${item.Details || ''}</td>
                    <td>
                        <button class="btn btn-sm btn-outline-primary edit-btn" data-index="${index}"><i class="bi bi-pencil"></i></button>
                        <button class="btn btn-sm btn-outline-danger delete-btn" data-index="${index}"><i class="bi bi-trash"></i></button>
                    </td>
                </tr>`;
            tableBody.innerHTML += row;
        });
    }

    // --- EVENT HANDLING ---
    function addEventListeners() {
        document.getElementById('add-activity-btn')?.addEventListener('click', showModalForAdd);
        document.getElementById('save-activity-btn')?.addEventListener('click', handleFormSave);
        document.getElementById('export-excel-btn')?.addEventListener('click', exportDataToExcel);

        document.getElementById('overview-table-body')?.addEventListener('click', function(e) {
            const target = e.target.closest('button');
            if (!target) return;

            const index = target.dataset.index;
            if (target.classList.contains('edit-btn')) {
                showModalForEdit(index);
            } else if (target.classList.contains('delete-btn')) {
                handleDelete(index);
            }
        });
    }
    
    // --- MODAL & CRUD ---
    function showModalForAdd() {
        document.getElementById('activity-form').reset();
        document.getElementById('modal-activity-index').value = '';
        document.getElementById('activityModalLabel').textContent = 'Add New Activity';
        activityModal.show();
    }

    function showModalForEdit(index) {
        const item = overviewData[index];
        if (!item) return;

        document.getElementById('activity-form').reset();
        document.getElementById('modal-activity-index').value = index;
        document.getElementById('activityModalLabel').textContent = 'Edit Activity';

        document.getElementById('modal-topics').value = item.Topics || '';
        document.getElementById('modal-activity').value = item.Activity || '';
        document.getElementById('modal-responsible').value = item.Responsible || '';
        document.getElementById('modal-target-date').value = item['Target Date'] || '';
        document.getElementById('modal-status').value = (item.Status || 'open').toLowerCase();
        document.getElementById('modal-environment').value = item.Environment || '';
        document.getElementById('modal-details').value = item.Details || '';
        document.getElementById('modal-comment').value = item.Comment || '';

        activityModal.show();
    }

    function handleFormSave() {
        const index = document.getElementById('modal-activity-index').value;
        const record = {
            'Topics': document.getElementById('modal-topics').value,
            'Activity': document.getElementById('modal-activity').value,
            'Responsible': document.getElementById('modal-responsible').value,
            'Target Date': document.getElementById('modal-target-date').value,
            'Status': document.getElementById('modal-status').value,
            'Environment': document.getElementById('modal-environment').value,
            'Details': document.getElementById('modal-details').value,
            'Comment': document.getElementById('modal-comment').value,
        };

        if (index === '') {
            overviewData.push(record);
        } else {
            overviewData[parseInt(index)] = record;
        }

        saveDataAndReRender();
        activityModal.hide();
    }

    function handleDelete(index) {
        const item = overviewData[index];
        if (confirm(`Are you sure you want to delete the activity "${item.Activity}"?`)) {
            overviewData.splice(index, 1);
            saveDataAndReRender();
        }
    }

    // --- START THE APP ---
    initialize();
});
