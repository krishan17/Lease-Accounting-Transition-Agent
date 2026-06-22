document.addEventListener('DOMContentLoaded', () => {
    // Current active calculation inputs (single lease mode)
    let currentInputs = {
        pmt: 10000,
        annual_rate: 0.05,
        term: 24
    };

    // State tracking
    let parsedLeases = [];
    let isPortfolioActive = false;

    // UI Elements
    const navDashboard = document.getElementById('btn-nav-dashboard');
    const navSchedules = document.getElementById('btn-nav-schedules');
    const navEntries = document.getElementById('btn-nav-entries');
    
    const viewDashboard = document.getElementById('view-dashboard');
    const viewSchedules = document.getElementById('view-schedules');
    const viewEntries = document.getElementById('view-entries');

    const tabManual = document.getElementById('tab-manual');
    const tabUpload = document.getElementById('tab-upload');
    const panelManual = document.getElementById('panel-manual');
    const panelUpload = document.getElementById('panel-upload');

    const leaseForm = document.getElementById('lease-form');
    const rentInput = document.getElementById('rent-input');
    const termInput = document.getElementById('term-input');
    const rateInput = document.getElementById('rate-input');

    const dragArea = document.getElementById('drag-area');
    const fileInput = document.getElementById('file-input');
    const uploadProgress = document.getElementById('upload-progress');
    const fileParsedSection = document.getElementById('file-parsed-section');
    const leaseSelector = document.getElementById('lease-selector');

    const btnRunPortfolio = document.getElementById('btn-run-portfolio');
    const btnExportPortfolio = document.getElementById('btn-export-portfolio');

    const btnExportTop = document.getElementById('btn-export-excel-top');
    const toast = document.getElementById('toast');

    // Amortization Schedule Tabs
    const tabSchedComparison = document.getElementById('tab-sched-comparison');
    const tabSchedIfrs = document.getElementById('tab-sched-ifrs');
    const tabSchedUs = document.getElementById('tab-sched-us');

    const panelSchedComparison = document.getElementById('panel-sched-comparison');
    const panelSchedIfrs = document.getElementById('panel-sched-ifrs');
    const panelSchedUs = document.getElementById('panel-sched-us');

    // ----------------------------------------------------
    // NAVIGATION CONTROLLER
    // ----------------------------------------------------
    const sections = [
        { btn: navDashboard, view: viewDashboard },
        { btn: navSchedules, view: viewSchedules },
        { btn: navEntries, view: viewEntries }
    ];

    sections.forEach(sec => {
        sec.btn.addEventListener('click', (e) => {
            e.preventDefault();
            sections.forEach(s => {
                s.btn.classList.remove('active');
                s.view.classList.remove('active-view');
            });
            sec.btn.classList.add('active');
            sec.view.classList.add('active-view');
        });
    });

    // ----------------------------------------------------
    // INPUT TAB CONTROLLER (Manual vs Upload)
    // ----------------------------------------------------
    tabManual.addEventListener('click', () => {
        tabManual.classList.add('active');
        tabUpload.classList.remove('active');
        panelManual.classList.remove('hidden');
        panelUpload.classList.add('hidden');
    });

    tabUpload.addEventListener('click', () => {
        tabUpload.classList.add('active');
        tabManual.classList.remove('active');
        panelUpload.classList.remove('hidden');
        panelManual.classList.add('hidden');
    });

    // ----------------------------------------------------
    // SCHEDULE DETAILS TAB CONTROLLER
    // ----------------------------------------------------
    const schedTabs = [
        { tab: tabSchedComparison, panel: panelSchedComparison },
        { tab: tabSchedIfrs, panel: panelSchedIfrs },
        { tab: tabSchedUs, panel: panelSchedUs }
    ];

    schedTabs.forEach(item => {
        item.tab.addEventListener('click', () => {
            schedTabs.forEach(i => {
                i.tab.classList.remove('active');
                i.panel.classList.add('hidden');
            });
            item.tab.classList.add('active');
            item.panel.classList.remove('hidden');
        });
    });

    // ----------------------------------------------------
    // HELPERS & TOASTS
    // ----------------------------------------------------
    function showToast(message, type = 'success') {
        toast.className = `toast ${type}`;
        toast.innerHTML = type === 'success' 
            ? `<i class="fa-solid fa-circle-check"></i> <span>${message}</span>`
            : `<i class="fa-solid fa-circle-exclamation"></i> <span>${message}</span>`;
        toast.classList.remove('hidden');
        
        setTimeout(() => {
            toast.classList.add('hidden');
        }, 4000);
    }

    function formatCurrency(val) {
        if (val === '-' || val === null || val === undefined) return '-';
        return Number(val).toLocaleString('en-US', { style: 'currency', currency: 'USD' });
    }

    function formatPercent(val) {
        return (Number(val) * 100).toFixed(4) + '%';
    }

    // ----------------------------------------------------
    // DATA COMPUTATION & INVOCATIONS (Single Mode)
    // ----------------------------------------------------
    async function calculateSchedules(pmt, rate, term) {
        currentInputs = { pmt, annual_rate: rate, term };
        isPortfolioActive = false;
        
        try {
            const response = await fetch('/api/calculate', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ pmt, annual_rate: rate, term })
            });
            
            const data = await response.json();
            if (data.success) {
                renderAllResults(data.results);
                showToast('Lease schedules computed successfully!');
            } else {
                showToast(data.error || 'Failed to compute schedules', 'error');
            }
        } catch (err) {
            showToast('Server connection error during calculation', 'error');
            console.error(err);
        }
    }

    // Handle Form Submit
    leaseForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const pmt = parseFloat(rentInput.value);
        const rate = parseFloat(rateInput.value) / 100.0;
        const term = parseInt(termInput.value);
        
        calculateSchedules(pmt, rate, term);
    });

    // ----------------------------------------------------
    // DATA COMPUTATION & INVOCATIONS (Portfolio Mode)
    // ----------------------------------------------------
    async function calculatePortfolio() {
        if (parsedLeases.length === 0) {
            showToast('No parsed leases available. Please upload a file first.', 'error');
            return;
        }
        
        try {
            const response = await fetch('/api/calculate_portfolio', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ leases: parsedLeases })
            });
            
            const data = await response.json();
            if (data.success) {
                isPortfolioActive = true;
                renderPortfolioResults(data.results);
                showToast('Portfolio calculations consolidated successfully!');
            } else {
                showToast(data.error || 'Failed to compute portfolio schedules', 'error');
            }
        } catch (err) {
            showToast('Server connection error during portfolio calculation', 'error');
            console.error(err);
        }
    }

    btnRunPortfolio.addEventListener('click', calculatePortfolio);

    // ----------------------------------------------------
    // FILE UPLOAD AND PARSING
    // ----------------------------------------------------
    // Drag-over styling
    ['dragenter', 'dragover'].forEach(eventName => {
        dragArea.addEventListener(eventName, (e) => {
            e.preventDefault();
            dragArea.classList.add('dragover');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dragArea.addEventListener(eventName, (e) => {
            e.preventDefault();
            dragArea.classList.remove('dragover');
        }, false);
    });

    dragArea.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        if (files.length) {
            handleFileUpload(files[0]);
        }
    });

    dragArea.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
        if (fileInput.files.length) {
            handleFileUpload(fileInput.files[0]);
        }
    });

    async function handleFileUpload(file) {
        const formData = new FormData();
        formData.append('file', file);

        dragArea.classList.add('hidden');
        uploadProgress.classList.remove('hidden');
        fileParsedSection.classList.add('hidden');

        try {
            const response = await fetch('/api/parse', {
                method: 'POST',
                body: formData
            });

            const data = await response.json();
            uploadProgress.classList.add('hidden');
            
            if (data.success) {
                parsedLeases = data.leases;
                populateLeaseSelector(parsedLeases);
                
                fileParsedSection.classList.remove('hidden');
                showToast(`Loaded ${parsedLeases.length} leases from spreadsheet!`);
                
                // Automatically run the consolidated portfolio view
                calculatePortfolio();
            } else {
                dragArea.classList.remove('hidden');
                showToast(data.error || 'Failed to parse file', 'error');
            }
        } catch (err) {
            uploadProgress.classList.add('hidden');
            dragArea.classList.remove('hidden');
            showToast('Server connection error during upload', 'error');
            console.error(err);
        }
    }

    function populateLeaseSelector(leases) {
        leaseSelector.innerHTML = '';
        
        // Add default option
        const defOpt = document.createElement('option');
        defOpt.value = "-1";
        defOpt.textContent = "-- Choose lease to analyze single --";
        leaseSelector.appendChild(defOpt);

        leases.forEach((lease, idx) => {
            const option = document.createElement('option');
            option.value = idx;
            option.textContent = `${lease.name} (Rent: $${lease.rent.toLocaleString()}, Term: ${lease.term}m, Rate: ${(lease.rate * 100).toFixed(2)}%)`;
            leaseSelector.appendChild(option);
        });
    }

    function selectLease(idx) {
        if (idx === -1) {
            // Re-run portfolio consolidated
            calculatePortfolio();
            return;
        }
        
        const lease = parsedLeases[idx];
        if (!lease) return;

        // Auto-populate form
        rentInput.value = lease.rent;
        termInput.value = lease.term;
        rateInput.value = (lease.rate * 100).toFixed(4);

        // Run calculation
        calculateSchedules(lease.rent, lease.rate, lease.term);
    }

    leaseSelector.addEventListener('change', () => {
        selectLease(parseInt(leaseSelector.value));
    });

    // ----------------------------------------------------
    // EXPORT TO EXCEL (Single & Portfolio)
    // ----------------------------------------------------
    async function triggerExcelExport() {
        try {
            const response = await fetch('/api/export', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(currentInputs)
            });

            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = `DilipFin_Lease_Report_${currentInputs.pmt}_${currentInputs.term}m.xlsx`;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                showToast('Hulu Green Excel spreadsheet exported successfully!');
            } else {
                showToast('Excel generation failed on server', 'error');
            }
        } catch (err) {
            showToast('Connection error during Excel export', 'error');
            console.error(err);
        }
    }

    async function triggerPortfolioExport() {
        if (parsedLeases.length === 0) {
            showToast('No leases to export. Please upload a file first.', 'error');
            return;
        }
        
        try {
            const response = await fetch('/api/export_portfolio', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ leases: parsedLeases })
            });

            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                a.download = `DilipFin_Portfolio_Report.xlsx`;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                showToast('Hulu Green Portfolio Excel exported successfully!');
            } else {
                showToast('Portfolio Excel generation failed on server', 'error');
            }
        } catch (err) {
            showToast('Connection error during portfolio Excel export', 'error');
            console.error(err);
        }
    }

    btnExportTop.addEventListener('click', () => {
        if (isPortfolioActive) {
            triggerPortfolioExport();
        } else {
            triggerExcelExport();
        }
    });

    btnExportPortfolio.addEventListener('click', triggerPortfolioExport);

    // ----------------------------------------------------
    // RENDERING FUNCTIONS (Single Mode)
    // ----------------------------------------------------
    function renderAllResults(results) {
        // 1. Present Value details
        document.getElementById('lbl-pmt').textContent = `${formatCurrency(currentInputs.pmt)} per month`;
        document.getElementById('lbl-rate').textContent = `${(currentInputs.annual_rate * 100).toFixed(2)}% &divide; 12 = ${formatPercent(currentInputs.annual_rate / 12.0)} per month`;
        document.getElementById('lbl-term').textContent = `${currentInputs.term} months`;
        document.getElementById('val-pv').textContent = formatCurrency(results.pv);

        // 2. Yearly summary table
        const tbodyYearly = document.getElementById('tbody-yearly-summary');
        tbodyYearly.innerHTML = '';
        results.yearly_comparison.forEach(y => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>Year ${y.year}</td>
                <td>Months ${y.months}</td>
                <td class="text-right">${formatCurrency(y.ifrs_expense)}</td>
                <td class="text-right">${formatCurrency(y.us_expense)}</td>
                <td class="text-right ${y.expense_diff > 0 ? 'highlight-row' : ''}">${formatCurrency(y.expense_diff)}</td>
                <td class="text-right">${formatCurrency(y.ifrs_rou)}</td>
                <td class="text-right">${formatCurrency(y.us_rou)}</td>
                <td class="text-right">${formatCurrency(y.rou_diff)}</td>
                <td class="text-right">${formatCurrency(y.ifrs_liab)}</td>
            `;
            tbodyYearly.appendChild(tr);
        });

        // 3. Side-by-side comparison schedule
        const tbodyComparison = document.getElementById('tbody-sched-comparison');
        tbodyComparison.innerHTML = '';
        results.monthly_comparison.forEach(m => {
            const tr = document.createElement('tr');
            if (m.month === 0) tr.className = 'highlight-row';
            tr.innerHTML = `
                <td>${m.month}</td>
                <td class="text-right">${m.month === 0 ? '-' : formatCurrency(m.ifrs_expense)}</td>
                <td class="text-right">${m.month === 0 ? '-' : formatCurrency(m.us_expense)}</td>
                <td class="text-right">${m.month === 0 ? '-' : formatCurrency(m.expense_diff)}</td>
                <td class="text-right">${formatCurrency(m.ifrs_rou)}</td>
                <td class="text-right">${formatCurrency(m.us_rou)}</td>
                <td class="text-right">${formatCurrency(m.rou_diff)}</td>
                <td class="text-right">${formatCurrency(m.ifrs_liab)}</td>
                <td class="text-right">${formatCurrency(m.us_liab)}</td>
                <td class="text-right">${formatCurrency(m.liab_diff)}</td>
            `;
            tbodyComparison.appendChild(tr);
        });

        // 4. Detailed IFRS Schedule
        const tbodyIfrs = document.getElementById('tbody-sched-ifrs');
        tbodyIfrs.innerHTML = '';
        results.ifrs_schedule.forEach(r => {
            const tr = document.createElement('tr');
            if (r.month === 0) tr.className = 'highlight-row';
            tr.innerHTML = `
                <td>${r.month}</td>
                <td class="text-right">${r.month === 0 ? '-' : formatCurrency(r.cash)}</td>
                <td class="text-right">${r.month === 0 ? '-' : formatCurrency(r.interest)}</td>
                <td class="text-right">${r.month === 0 ? '-' : formatCurrency(r.depreciation)}</td>
                <td class="text-right">${r.month === 0 ? '-' : formatCurrency(r.total_expense)}</td>
                <td class="text-right">${formatCurrency(r.liability_closing)}</td>
                <td class="text-right">${formatCurrency(r.rou_closing)}</td>
            `;
            tbodyIfrs.appendChild(tr);
        });

        // 5. Detailed US GAAP Schedule
        const tbodyUs = document.getElementById('tbody-sched-us');
        tbodyUs.innerHTML = '';
        results.us_schedule.forEach(r => {
            const tr = document.createElement('tr');
            if (r.month === 0) tr.className = 'highlight-row';
            tr.innerHTML = `
                <td>${r.month}</td>
                <td class="text-right">${r.month === 0 ? '-' : formatCurrency(r.cash)}</td>
                <td class="text-right">${r.month === 0 ? '-' : formatCurrency(r.lease_expense)}</td>
                <td class="text-right">${r.month === 0 ? '-' : formatCurrency(r.interest)}</td>
                <td class="text-right">${r.month === 0 ? '-' : formatCurrency(r.liab_reduction)}</td>
                <td class="text-right">${r.month === 0 ? '-' : formatCurrency(r.rou_amortization)}</td>
                <td class="text-right">${formatCurrency(r.liability_balance)}</td>
                <td class="text-right">${formatCurrency(r.rou_balance)}</td>
            `;
            tbodyUs.appendChild(tr);
        });

        // 6. Journal Entries
        renderJournalEntries(results.journal_entries);
    }

    // ----------------------------------------------------
    // RENDERING FUNCTIONS (Portfolio Mode)
    // ----------------------------------------------------
    function renderPortfolioResults(results) {
        // 1. Present Value details (Aggregate representation)
        document.getElementById('lbl-pmt').textContent = "Portfolio Combined";
        document.getElementById('lbl-rate').textContent = `${results.leases.length} Active Leases Consolidated`;
        document.getElementById('lbl-term').textContent = "Varying lease tenures";
        document.getElementById('val-pv').textContent = formatCurrency(results.pv);

        // 2. Yearly summary table
        const tbodyYearly = document.getElementById('tbody-yearly-summary');
        tbodyYearly.innerHTML = '';
        results.yearly_comparison.forEach(y => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>Year ${y.year}</td>
                <td>Months ${y.months}</td>
                <td class="text-right">${formatCurrency(y.ifrs_expense)}</td>
                <td class="text-right">${formatCurrency(y.us_expense)}</td>
                <td class="text-right ${y.expense_diff > 0 ? 'highlight-row' : ''}">${formatCurrency(y.expense_diff)}</td>
                <td class="text-right">${formatCurrency(y.ifrs_rou)}</td>
                <td class="text-right">${formatCurrency(y.us_rou)}</td>
                <td class="text-right">${formatCurrency(y.rou_diff)}</td>
                <td class="text-right">${formatCurrency(y.ifrs_liab)}</td>
            `;
            tbodyYearly.appendChild(tr);
        });

        // 3. Informational tables in schedules view for bulk mode
        const msgHtml = `
            <tr>
                <td colspan="10" class="text-center" style="padding: 40px; color: var(--text-muted);">
                    <i class="fa-solid fa-circle-info" style="font-size: 24px; margin-bottom: 12px; color: var(--neon-green); display: block;"></i>
                    Portfolio Consolidated Summary is active.
                    <br><span style="font-size: 12px; margin-top: 4px; display: inline-block;">Select a specific lease from the dropdown on the left to view detailed monthly schedules.</span>
                </td>
            </tr>
        `;
        document.getElementById('tbody-sched-comparison').innerHTML = msgHtml;
        document.getElementById('tbody-sched-ifrs').innerHTML = msgHtml.replace('colspan="10"', 'colspan="7"');
        document.getElementById('tbody-sched-us').innerHTML = msgHtml.replace('colspan="10"', 'colspan="8"');

        // 4. Journal Entries
        renderJournalEntries(results.journal_entries);
    }

    function renderJournalEntries(journalEntries) {
        const jeContainer = document.getElementById('journal-entries-container');
        jeContainer.innerHTML = '';
        
        journalEntries.forEach(jeGroup => {
            const card = document.createElement('div');
            card.className = 'je-card';
            
            let tableRows = '';
            jeGroup.entries.forEach(entry => {
                const isCredit = entry.credit > 0;
                tableRows += `
                    <tr>
                        <td class="${isCredit ? 'credit-account' : ''}">${entry.account}</td>
                        <td class="text-right">${entry.debit > 0 ? formatCurrency(entry.debit) : ''}</td>
                        <td class="text-right">${entry.credit > 0 ? formatCurrency(entry.credit) : ''}</td>
                    </tr>
                `;
            });

            card.innerHTML = `
                <div class="je-card-header">
                    <h4>${jeGroup.type}</h4>
                    <span class="badge ifrs">IFRS 16</span>
                </div>
                <div class="je-body">
                    <table class="je-table">
                        <thead>
                            <tr>
                                <th>GL Account / Description</th>
                                <th class="text-right" style="width: 150px;">Debit</th>
                                <th class="text-right" style="width: 150px;">Credit</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${tableRows}
                        </tbody>
                    </table>
                </div>
            `;
            jeContainer.appendChild(card);
        });
    }

    // Trigger initial calculation on load
    calculateSchedules(10000, 0.05, 24);
});
