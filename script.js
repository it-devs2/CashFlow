const API_URL = "https://script.google.com/macros/s/AKfycbydwExXEbI-ZkhxTPx9FCc_PYbHlYlTYMS1WPsPumHHqXAQ7GNBQNY8QisMu7oCBjkQ/exec";

// Utility: Format Number as Currency safely
function checkValue(val) {
    if (val === null || val === undefined || val === '') return '-';
    return Number(val).toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Utility: Robust Row Detection & Value Extraction
function getRowType(row) {
    const t = (row['Type'] || row.type || '').toString().trim().toLowerCase();
    if (t === 'income' || t === 'expense') return t;

    // Fallback: Detection by column presence/values
    const cashIn = Number(row['Cash In'] || row.cashIn) || 0;
    const cashOut = Number(row['Cash Out'] || row.cashOut || row['# Amount']) || 0;

    if (cashIn !== 0 && cashOut === 0) return 'income';
    if (cashOut !== 0 && cashIn === 0) return 'expense';

    // If only Generic Amount exists, assume income as fallback
    const generic = Number(row['Amount (THB)'] || row['Amount'] || row.amount) || 0;
    if (generic !== 0) return 'income';

    return t; // might be empty
}

function getRowAmount(row, targetType) {
    const cashIn = Number(row['Cash In'] || row.cashIn) || 0;
    const cashOut = Number(row['Cash Out'] || row.cashOut || row['# Amount']) || 0;
    const generic = Number(row['Amount (THB)'] || row['Amount'] || row.amount) || 0;

    const maxVal = Math.max(cashIn, cashOut, Math.abs(generic));

    if (targetType === 'income') {
        const valIn = cashIn || generic;
        const valOut = cashOut;
        
        // Fix: If Amount exactly equals CashOut but it's assigned 'Income', or user swapped columns
        if (valIn === valOut && valIn > 0) return valIn;
        if (cashIn === 0 && generic === 0 && cashOut > 0) return cashOut;

        const net = valIn - valOut;
        return net >= 0 ? net : (valIn || maxVal);
    }
    if (targetType === 'expense') {
        const valOut = cashOut || generic;
        const valIn = cashIn;
        
        // Fix: If Amount exactly equals CashIn but it's assigned 'Expense' (e.g. 1900000 - 1900000 = 0 bug)
        if (valOut === valIn && valOut > 0) return valOut;
        if (cashOut === 0 && generic === 0 && cashIn > 0) return cashIn;

        const net = valOut - valIn;
        return net >= 0 ? net : (valOut || maxVal);
    }
    return generic || cashIn || cashOut || 0;
}

// Global State
let totalIncomeActual = 0;
let totalExpenseActual = 0;
let totalIncomePlan = 0;
let totalExpensePlan = 0;

let allTransactions = []; // All Transactions (Actual)
let allPlans = [];        // All Plans
let _lastFilteredTransactions = [];
let _lastFilteredPlans = [];
let allParties = [];      // All party names from All_Party sheet
let selectedCreditors = new Set(); // Multi-select Set for creditors

let comparisonChart; // ApexCharts instance for Overview

// Bank Balances Array (Will be populated from API)
let bankBalances = [];

// Bank Detail Modal state
let _bankModalRows = [];
let _currentBankName = '';

// ✅ เพิ่มตัวแปรสำหรับรับค่าจาก Cell G1 และ H2 โดยตรง
let _availableBalanceH2 = 0;
let _dateG1 = '-';

// -------------------------------------------------
// INIT: Fetch data from Google Apps Script API
// -------------------------------------------------
async function initDashboard() {
    try {
        document.getElementById('table-body').innerHTML = '<tr><td colspan="15" style="text-align: center; padding: 30px;">Loading data from Google Sheets...</td></tr>';

        const response = await fetch(API_URL);
        const dataStatus = await response.json();

        if (dataStatus && dataStatus.status === 'success') {
            // Filter out completely empty rows from spreadsheet before storing
            const isValidRow = row => Object.values(row).some(v => v !== null && v !== undefined && v.toString().trim() !== '');
            allTransactions = (dataStatus.transactions || []).filter(isValidRow);
            allPlans = (dataStatus.plans || []).filter(isValidRow);

            // Sort all data by Date Ascending (oldest to newest, e.g. 1-31). Push invalid/empty dates to the bottom.
            const sortByDateAsc = (a, b) => {
                const dA = a['Date'] || a.date;
                const dB = b['Date'] || b.date;
                
                if (!dA && !dB) return 0;
                if (!dA) return 1;
                if (!dB) return -1;

                const dateA = new Date(dA);
                const dateB = new Date(dB);
                
                if (isNaN(dateA) && isNaN(dateB)) return 0;
                if (isNaN(dateA)) return 1;
                if (isNaN(dateB)) return -1;
                
                return dateA - dateB;
            };
            allTransactions.sort(sortByDateAsc);
            allPlans.sort(sortByDateAsc);

            if (dataStatus.bankBalances && dataStatus.bankBalances.length > 0) {
                bankBalances = dataStatus.bankBalances;
            }

            // ✅ รับค่าจาก Cell H2 และ G1 หากมีส่งมาจาก API (รองรับหลายชื่อเผื่อไว้)
            _availableBalanceH2 = (
                dataStatus.availableBalanceH2 ||
                dataStatus.balanceH2 ||
                dataStatus.selectedBalance ||
                dataStatus.totalAvailable ||
                dataStatus.h2Value ||
                0
            );
            _dateG1 = (
                dataStatus.dateG1 ||
                dataStatus.asOfDate ||
                dataStatus.lastUpdate ||
                dataStatus.sheetDate ||
                dataStatus.g1Value ||
                '-'
            );

            // Log ข้อมูลเบื้องต้นเพื่อช่วยตรวจสอบกรณีไม่ตรง
            console.log("Bank Summary Data Received:", { availableBalanceH2: _availableBalanceH2, dateG1: _dateG1 });

            // Use All_Party from API if available, else extract from transaction data
            const apiParties = (dataStatus.parties || []).filter(p => p && p.trim() !== '');
            if (apiParties.length > 0) {
                allParties = apiParties;
            } else {
                // Fallback: extract unique names from all data columns
                const nameSet = new Set();
                const allData = [...allTransactions, ...allPlans];
                allData.forEach(row => {
                    ['Customer', 'customer', 'Vendor', 'vendor', 'Party', 'party', 'Name', 'name'].forEach(k => {
                        const val = row[k];
                        if (val && val.toString().trim()) nameSet.add(val.toString().trim());
                    });
                });
                allParties = [...nameSet].sort((a, b) => a.localeCompare(b, 'th'));
            }

            // ✅ เก็บยอดรวมที่คำนวณจาก Code.gs โดยตรง (แม่นยำ 100% ตรงกับ Sheet)
            if (dataStatus.summaryIncomeActual !== undefined) window._serverSummary = {
                incomeActual: dataStatus.summaryIncomeActual,
                expenseActual: dataStatus.summaryExpenseActual,
                incomePlan: dataStatus.summaryIncomePlan,
                expensePlan: dataStatus.summaryExpensePlan
            };

            populateFilterDropdowns(allTransactions, allPlans);
            populateBankDateFilters();
            initCreditorAutocomplete();
            applyFilters();
            renderBankBalances();
        } else {
            console.error("API Error or Invalid format:", dataStatus);
            document.getElementById('table-body').innerHTML = '<tr><td colspan="15" style="text-align: center; color: var(--expense); padding: 30px;">Error: API returned invalid data format.</td></tr>';
        }
    } catch (error) {
        console.error("Fetch failed:", error);
        document.getElementById('table-body').innerHTML = '<tr><td colspan="15" style="text-align: center; color: var(--expense); padding: 30px;">Failed to fetch data. Check API URL or CORS policy.</td></tr>';
    }
}

// -------------------------------------------------
// POPULATE: Fill dropdown options from real data
// -------------------------------------------------
function populateFilterDropdowns(transactions, plans) {
    const allData = [...transactions, ...plans];

    const banks = new Set();
    const days = new Set();
    const months = new Set();
    const years = new Set();

    const monthNames = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
        'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];

    allData.forEach(row => {
        const bank = row['Bank'] || row.bank;
        if (bank) banks.add(bank.toString().trim());

        const rawDate = row['Date'] || row.date;
        if (rawDate) {
            const d = new Date(rawDate);
            if (!isNaN(d)) {
                days.add(d.getDate());
                months.add(d.getMonth() + 1); // 1-12
                years.add(d.getFullYear());
            }
        }
    });

    // Populate Bank
    const bankSel = document.getElementById('filter-bank');
    bankSel.innerHTML = '<option value="All">ทั้งหมด</option>';
    [...banks].sort().forEach(b => {
        const opt = document.createElement('option');
        opt.value = b; opt.textContent = b;
        bankSel.appendChild(opt);
    });

    // Populate Day
    const daySel = document.getElementById('filter-day');
    daySel.innerHTML = '<option value="All">ทั้งหมด</option>';
    [...days].sort((a, b) => a - b).forEach(d => {
        const opt = document.createElement('option');
        opt.value = d;
        opt.textContent = String(d).padStart(2, '0');
        daySel.appendChild(opt);
    });

    // Populate Month
    const monthSel = document.getElementById('filter-month');
    monthSel.innerHTML = '<option value="All">ทั้งหมด</option>';
    [...months].sort((a, b) => a - b).forEach(m => {
        const opt = document.createElement('option');
        opt.value = m;
        opt.textContent = `${String(m).padStart(2, '0')} - ${monthNames[m - 1]}`;
        monthSel.appendChild(opt);
    });

    // Populate Year
    const yearSel = document.getElementById('filter-year');
    yearSel.innerHTML = '<option value="All">ทั้งหมด</option>';
    [...years].sort((a, b) => b - a).forEach(y => {
        const opt = document.createElement('option');
        opt.value = y; opt.textContent = y;
        yearSel.appendChild(opt);
    });
}

// -------------------------------------------------
// FILTER: Apply all active filters and re-render
// -------------------------------------------------
function applyFilters() {
    const creditor = (document.getElementById('filter-creditor')?.value || '').toLowerCase().trim();
    const bank = document.getElementById('filter-bank')?.value || 'All';
    const type = document.getElementById('filter-type')?.value || 'All';
    const day = document.getElementById('filter-day')?.value || 'All';
    const month = document.getElementById('filter-month')?.value || 'All';
    const year = document.getElementById('filter-year')?.value || 'All';

    function matchRow(row) {
        // Creditor filter: search across Customer, Vendor, Name, Party
        if (selectedCreditors.size > 0) {
            const fields = [
                row['Customer'], row['Vendor'], row['Name'], row['Party'], row['Description'], row.customer, row.vendor, row.name, row.party, row.description
            ].map(v => (v || '').toString().trim());

            // Check if any row field matches exactly any of the selected creditors
            let matched = false;
            for (let f of fields) {
                if (!f) continue;
                if (selectedCreditors.has(f)) {
                    matched = true;
                    break;
                }
            }
            if (!matched) return false;
        }

        // Bank filter
        if (bank !== 'All') {
            const b = (row['Bank'] || row.bank || '').toString().trim();
            if (b !== bank) return false;
        }

        // Type filter
        if (type !== 'All') {
            const t = (row['Type'] || row.type || '').toString().trim().toLowerCase();
            if (t !== type.toLowerCase()) return false;
        }

        // Date filters
        const rawDate = row['Date'] || row.date;
        if (rawDate && (day !== 'All' || month !== 'All' || year !== 'All')) {
            const d = new Date(rawDate);
            if (!isNaN(d)) {
                if (day !== 'All' && d.getDate() !== Number(day)) return false;
                if (month !== 'All' && (d.getMonth() + 1) !== Number(month)) return false;
                if (year !== 'All' && d.getFullYear() !== Number(year)) return false;
            }
        }

        return true;
    }

    const filteredTransactions = allTransactions.filter(matchRow);
    const filteredPlans = allPlans.filter(matchRow);

    // Store for modal access
    _lastFilteredTransactions = filteredTransactions;
    _lastFilteredPlans = filteredPlans;

    const isFiltered = selectedCreditors.size > 0 || bank !== 'All' || type !== 'All' || day !== 'All' || month !== 'All' || year !== 'All';

    window.tableRenderLimit = 150; // Reset load limit on filter change

    renderTable(filteredTransactions, filteredPlans);
    updateSummary(isFiltered);

    // กรอง Bank Balances ตาม filter-bank ที่เลือก
    if (bank !== 'All') {
        const filteredBankBalances = bankBalances.filter(b =>
            (b.bank || '').trim().toLowerCase() === bank.trim().toLowerCase()
        );
        renderBankBalances(filteredBankBalances);
    } else {
        renderBankBalances(bankBalances);
    }
}

// -------------------------------------------------
// LOAD MORE TRANSACTIONS (PAGINATION)
// -------------------------------------------------
window.tableRenderLimit = 150;
function loadMoreTransactions() {
    window.tableRenderLimit += 200;
    renderTable(typeof _lastFilteredTransactions !== 'undefined' ? _lastFilteredTransactions : allTransactions, typeof _lastFilteredPlans !== 'undefined' ? _lastFilteredPlans : allPlans);
}

// -------------------------------------------------
// RENDER TABLE + CALCULATE TOTALS
// -------------------------------------------------
function renderTable(transactionsData, plansData = []) {
    const tbody = document.getElementById('table-body');

    totalIncomeActual = 0;
    totalExpenseActual = 0;
    totalIncomePlan = 0;
    totalExpensePlan = 0;

    // Process Plans (Incoming_Plan + Payment_Plan sheets)
    plansData.forEach(row => {
        const rowStatus = (row.status || row['Status'] || '').toString().trim().toLowerCase();
        if (rowStatus !== 'plan') return;

        const rowType = getRowType(row);
        const amt = getRowAmount(row, rowType);

        if (rowType === 'income') totalIncomePlan += amt;
        if (rowType === 'expense') totalExpensePlan += amt;
    });

    let htmlFragments = [];
    let renderCount = 0;
    const limit = window.tableRenderLimit || 150;

    // Process Transactions (Actual)
    transactionsData.forEach(row => {
        const rowStatus = (row.status || row['Status'] || '').toString().trim().toLowerCase();
        const rowType = getRowType(row);

        // ✅ นับเฉพาะแถวที่ไม่ใช่ Status = "Plan" เข้า Actual
        // ✅ คำนวณแบบ (In - Out) เพื่อความแม่นยำสูงสุดตาม Sheet
        if (rowStatus !== 'plan') {
            const amt = getRowAmount(row, rowType);
            if (rowType === 'income') totalIncomeActual += amt;
            if (rowType === 'expense') totalExpenseActual += amt;
        }

        // Render limited rows to DOM to prevent lag
        if (renderCount < limit) {
            const cashInVal = row.cashIn || row['Cash In'];
            const cashOutVal = row.cashOut || row['Cash Out'];

            const rawDate = row['Date'] || row.date || '';
            let displayDate = rawDate;
            try {
                const d = new Date(rawDate);
                if (!isNaN(d)) displayDate = d.toLocaleDateString('th-TH', { day: '2-digit', month: '2-digit', year: 'numeric' });
            } catch (e) { }

            htmlFragments.push(`
                <tr>
                    <td>${displayDate}</td>
                    <td>${row.docNo || row['Doc No'] || ''}</td>
                    <td>${row.description || row['Description'] || ''}</td>
                    <td>${row.bank || row['Bank'] || ''}</td>
                    <td class="type-cell"><span class="type-${rowType}">${rowType || ''}</span></td>
                    <td>${row.category || row['Category'] || ''}</td>
                    <td>${row.group || row['Group'] || ''}</td>
                    <td>${row.name || row['Name'] || ''}</td>
                    <td>${row.party || row['Party'] || row.customer || row['Customer'] || ''}</td>
                    <td>${row.project || row['Project'] || ''}</td>
                    <td>${row.status || row['Status'] || ''}</td>
                    <td class="numeric">${checkValue(row.amount || row['Amount (THB)'] || row['Amount'])}</td>
                    <td class="numeric">${checkValue(cashInVal)}</td>
                    <td class="numeric">${checkValue(cashOutVal)}</td>
                    <td>${row.transferTo || row['Transfer To'] || ''}</td>
                </tr>
            `);
            renderCount++;
        }
    });

    if (transactionsData.length > limit) {
        const remaining = transactionsData.length - limit;
        htmlFragments.push(`
            <tr><td colspan="15" style="text-align: center; padding: 25px;">
                <button onclick="loadMoreTransactions()" style="padding: 10px 24px; background: rgba(56, 189, 248, 0.1); border: 1px solid rgba(56, 189, 248, 0.4); color: var(--primary); border-radius: 8px; cursor: pointer; font-family: inherit; font-size: 14px; font-weight: 600; transition: all 0.2s;">
                    👇 โหลดข้อมูลเพิ่มเติม (${remaining.toLocaleString()} รายการที่เหลือ)
                </button>
            </td></tr>
        `);
    }

    // Apply entire HTML at once instead of appendChild in loop (Massive performance boost)
    tbody.innerHTML = htmlFragments.join('');
}

// -------------------------------------------------
// UPDATE SUMMARY CARDS
// -------------------------------------------------
function updateSummary(isFiltered = false) {
    // ✅ ถ้าไม่ได้กรองข้อมูล และมีค่าจาก Server (Sheet) ให้ใช้ค่านั้นเพื่อให้ตรงกับ Sheet 100%
    if (!isFiltered && window._serverSummary) {
        totalIncomeActual = window._serverSummary.incomeActual;
        totalExpenseActual = window._serverSummary.expenseActual;
        totalIncomePlan = window._serverSummary.incomePlan;
        totalExpensePlan = window._serverSummary.expensePlan;
    }

    document.getElementById('income-actual').innerText = checkValue(totalIncomeActual);
    document.getElementById('expense-actual').innerText = checkValue(totalExpenseActual);
    document.getElementById('income-plan').innerText = checkValue(totalIncomePlan);
    document.getElementById('expense-plan').innerText = checkValue(totalExpensePlan);

    const totalIncomeGroup = totalIncomeActual + totalIncomePlan;
    const totalExpenseGroup = totalExpenseActual + totalExpensePlan;
    const netBalance = totalIncomeGroup - totalExpenseGroup;

    const netAmountEl = document.getElementById('header-net-amount');
    netAmountEl.innerText = checkValue(netBalance);
    netAmountEl.style.color = netBalance >= 0 ? 'var(--income)' : 'var(--expense)';

    // Update Overview Chart
    if (typeof updateOverviewChart === 'function') {
        updateOverviewChart();
    }
}

// -------------------------------------------------
// UPDATE OVERVIEW CHART
// -------------------------------------------------
function updateOverviewChart() {
    const monthlyIncome = new Array(12).fill(0);
    const monthlyExpense = new Array(12).fill(0);
    const thaiMonthCategories = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];

    let dataToProcess = allTransactions;
    if (typeof _lastFilteredTransactions !== 'undefined' && _lastFilteredTransactions.length > 0) {
        // If there are filters applied and there are results, use them.
        // However, if lengths don't match, we assume filters are active.
        if (allTransactions.length !== _lastFilteredTransactions.length) {
            dataToProcess = _lastFilteredTransactions;
        }
        // If user resets filters, _lastFilteredTransactions equals allTransactions
    } else if (typeof _lastFilteredTransactions !== 'undefined' && _lastFilteredTransactions.length === 0) {
        // No results after filter
        dataToProcess = [];
    }

    dataToProcess.forEach(row => {
        const rawDate = row['Date'] || row.date;
        if (!rawDate) return;
        const d = new Date(rawDate);
        if (isNaN(d)) return;

        const m = d.getMonth();
        const rowStatus = (row.status || row['Status'] || '').toString().trim().toLowerCase();

        if (rowStatus !== 'plan') {
            const rowType = getRowType(row);
            const amt = getRowAmount(row, rowType);
            if (rowType === 'income') monthlyIncome[m] += amt;
            if (rowType === 'expense') monthlyExpense[m] += amt;
        }
    });

    const seriesData = [
        { name: 'Income', data: monthlyIncome },
        { name: 'Expense', data: monthlyExpense }
    ];

    const chartData = {
        series: seriesData,
        chart: {
            type: 'line',
            height: 350,
            background: 'transparent',
            toolbar: { show: false },
            fontFamily: 'Outfit, sans-serif',
            zoom: { enabled: false },
            selection: { enabled: false },
            animations: {
                enabled: true,
                easing: 'easeinout',
                speed: 800,
                animateGradually: { enabled: true, delay: 150 },
                dynamicAnimation: { enabled: true, speed: 350 }
            },
            events: {
                mounted: function(ctx) {
                    const el = document.querySelector('#comparison-chart');
                    if (el) {
                        el.addEventListener('wheel', function(e) {
                            e.stopPropagation();
                        }, { passive: true });
                    }
                }
            }
        },
        colors: ['#10b981', '#ef4444'], // Income: Green, Expense: Red
        dataLabels: {
            enabled: true,
            formatter: function (val) {
                if (val === 0) return '';
                if (val >= 1000000) return (val / 1000000).toFixed(2) + 'M';
                if (val >= 1000) return (val / 1000).toFixed(1) + 'K';
                return val.toLocaleString('th-TH');
            },
            offsetY: -20,
            style: {
                fontSize: '12px',
                colors: ['#cbd5e1']
            }
        },
        stroke: {
            show: true,
            curve: 'smooth',
            width: 4
        },
        markers: {
            size: 6,
            colors: ['#10b981', '#ef4444'],
            strokeColors: '#1e293b',
            strokeWidth: 2,
            hover: {
                size: 8
            }
        },
        xaxis: {
            categories: thaiMonthCategories,
            labels: {
                style: {
                    colors: '#94a3b8',
                    fontSize: '14px',
                    fontWeight: 600
                }
            },
            axisBorder: { show: false },
            axisTicks: { show: false }
        },
        yaxis: {
            labels: {
                style: {
                    colors: '#94a3b8',
                    fontSize: '13px'
                },
                formatter: function (val) {
                    if (val === 0) return "฿0";
                    if (val >= 1000000) return "฿" + (val / 1000000).toFixed(1) + 'M';
                    if (val >= 1000) return "฿" + (val / 1000).toFixed(1) + 'K';
                    return "฿" + val.toLocaleString('th-TH');
                }
            }
        },
        fill: {
            type: 'gradient',
            gradient: {
                shade: 'dark',
                type: 'vertical',
                shadeIntensity: 0.5,
                opacityFrom: 1,
                opacityTo: 0.85,
                stops: [0, 100]
            }
        },
        grid: {
            borderColor: 'rgba(255, 255, 255, 0.05)',
            strokeDashArray: 4,
            yaxis: {
                lines: { show: true }
            }
        },
        legend: {
            position: 'top',
            horizontalAlign: 'right',
            labels: { colors: '#f8fafc' },
            fontFamily: 'Outfit, sans-serif',
            markers: {
                radius: 12
            }
        },
        tooltip: {
            theme: 'dark',
            y: {
                formatter: function (val) {
                    return "฿ " + val.toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
                }
            },
            style: {
                fontSize: '14px',
                fontFamily: 'Outfit, sans-serif'
            }
        }
    };

    if (comparisonChart) {
        comparisonChart.updateOptions({ xaxis: { categories: thaiMonthCategories } });
        comparisonChart.updateSeries(seriesData);
    } else {
        const chartEl = document.querySelector("#comparison-chart");
        if (chartEl) {
            comparisonChart = new ApexCharts(chartEl, chartData);
            comparisonChart.render();
        }
    }
}

// -------------------------------------------------
// RENDER BANK BALANCES
// -------------------------------------------------
function getBankLogoUrl(bankName) {
    const b = (bankName || '').toUpperCase();

    // ✅ ใช้รูปโลคอลที่อยู่ในโฟลเดอร์โปรเจกต์ก่อน
    if (b.includes('KBANK') || b.includes('K-BANK') || b.includes('KASIKORN')) return 'KBank.png';
    if (b.includes('KKP') || b.includes('KIATNAKIN')) return 'KKP.png';
    if (b.includes('SCB') || b.includes('SIAM COMMERCIAL')) return 'SCB.jpg';
    if (b.includes('TTB') || b.includes('TMB')) return 'TTB.png';

    if (b.includes('BBL') || b.includes('BANGKOK')) return 'BBL.png';
    if (b.includes('KTB') || b.includes('KRUNGTHAI')) return 'KTB.png';
    if (b.includes('BAY') || b.includes('KRUNGSRI')) return 'https://upload.wikimedia.org/wikipedia/commons/thumb/1/18/Krungsri_Bank_logo.svg/150px-Krungsri_Bank_logo.svg.png';
    if (b.includes('GSB') || b.includes('ออมสิน')) return 'https://upload.wikimedia.org/wikipedia/en/thumb/8/87/Government_Savings_Bank_%28Thailand%29_logo.svg/150px-Government_Savings_Bank_%28Thailand%29_logo.svg.png';

    // ❓ Default fallback icon (generic bank icon)
    return 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0ibm9uZSIgc3Ryb2tlPSIjMzhiZGY4IiBzdHJva2Utd2lkdGg9IjIiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCI+PHJlY3QgeD0iMiIgeT0iNyIgd2lkdGg9IjIwIiBoZWlnaHQ9IjE0IiByeD0iMiIgcnk9IjIiPjwvcmVjdD48cGF0aCBkPSJNMTYgMjFWNWEyIDIgMCAwIDAtMi0yaC00YTIgMiAwIDAgMC0yIDJ2MTYiPjwvcGF0aD48L3N2Zz4=';
}

function renderBankBalances(filteredBalances) {
    const container = document.getElementById('bank-balances-container');
    container.innerHTML = '';

    const displayBalances = filteredBalances !== undefined ? filteredBalances : bankBalances;

    if (!displayBalances || displayBalances.length === 0) {
        container.innerHTML = '<span style="color: var(--text-muted); font-size: 14px;">ไม่พบข้อมูลยอดคงเหลือ</span>';
        return;
    }

    // เรียงตามชื่อธนาคาร (KBANK, KKP, SCB ...) แล้วตามเลขบัญชี
    const sorted = [...displayBalances].sort((a, b) => {
        const nameA = (a.bank || '').toUpperCase();
        const nameB = (b.bank || '').toUpperCase();
        return nameA.localeCompare(nameB);
    });

    sorted.forEach(bank => {
        const card = document.createElement('div');
        card.className = 'bank-card';
        const logoUrl = getBankLogoUrl(bank.bank);

        // แยกชื่อธนาคาร vs เลขที่บัญชี
        const dashIdx = bank.bank.indexOf('-');
        const bankType = dashIdx !== -1 ? bank.bank.substring(0, dashIdx).trim() : bank.bank.trim();
        const accountNum = dashIdx !== -1 ? bank.bank.substring(dashIdx + 1).trim() : '';
        const safeBank = bank.bank.replace(/'/g, "\\'");

        card.innerHTML = `
            <div class="bank-card-header">
                <img src="${logoUrl}" alt="${bankType} Logo" class="bank-logo" onerror="this.onerror=null; this.src='data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0ibm9uZSIgc3Ryb2tlPSIjMzhiZGY4IiBzdHJva2Utd2lkdGg9IjIiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCI+PHJlY3QgeD0iMiIgeT0iNyIgd2lkdGg9IjIwIiBoZWlnaHQ9IjE0IiByeD0iMiIgcnk9IjIiPjwvcmVjdD48cGF0aCBkPSJNMTYgMjFWNWEyIDIgMCAwIDAtMi0yaC00YTIgMiAwIDAgMC0yIDJ2MTYiPjwvcGF0aD48L3N2Zz4='">
                <div class="bank-info">
                    <span class="bank-name">${bankType}</span>
                    ${accountNum ? `<span class="bank-account">เลขที่บัญชี: ${accountNum}</span>` : ''}
                </div>
            </div>
            <div class="bank-balance-wrapper">
                <span class="bank-balance">${checkValue(bank.balance)}</span>
            </div>
            <button class="btn-bank-detail" onclick="openBankDetailModal('${safeBank}', '${bankType}', '${accountNum}')">📋 ดูรายละเอียด</button>
        `;
        container.appendChild(card);
    });

    // --- Update Bank Summary (Total + Date) --- Unified Box ---
    const totalAvailableEl = document.getElementById('bank-total-available');
    const selectedDateEl = document.getElementById('bank-selected-date');

    if (totalAvailableEl) {
        // ให้ความสำคัญกับค่าจาก H2 หากมีการส่งมาจาก Server (ไม่เป็น 0)
        // หากไม่มีข้อมูลจาก Server จริงๆ ถึงจะใช้ยอดรวมจากบัตรธนาคารทั้งหมด
        const sumOfAllCards = sorted.reduce((sum, b) => sum + (Number(b.balance) || 0), 0);
        const finalTotal = (_availableBalanceH2 && _availableBalanceH2 !== 0) ? _availableBalanceH2 : sumOfAllCards;

        totalAvailableEl.textContent = checkValue(finalTotal);
    }

    if (selectedDateEl) {
        // ให้ความสำคัญกับค่าจาก G1 (ข้อมูล ณ วันที่) หากมีการส่งมาจาก Server
        if (_dateG1 && _dateG1 !== '-') {
            selectedDateEl.textContent = _dateG1;
        } else {
            // Fallback: แสดงวันที่ตามฟิลเตอร์
            const d = document.getElementById('bb-filter-day').value;
            const m = document.getElementById('bb-filter-month').value;
            const y = document.getElementById('bb-filter-year').value;

            if (d === 'All' && m === 'All' && y === 'All') {
                selectedDateEl.textContent = 'ยอดล่าสุดทั้งหมด';
            } else {
                const cleanM = m !== 'All' ? (m.split('-')[1] || m).trim() : '';
                selectedDateEl.textContent = `${d !== 'All' ? d : ''} ${cleanM} ${y !== 'All' ? y : ''}`.trim() || '-';
            }
        }
    }
}

// -------------------------------------------------
// BANK DATE FILTER POPULATE
// -------------------------------------------------
function populateBankDateFilters() {
    const monthNames = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
        'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];

    const days = new Set();
    const months = new Set();
    const years = new Set();

    allTransactions.forEach(row => {
        const rawDate = row['Date'] || row.date;
        if (rawDate) {
            const d = new Date(rawDate);
            if (!isNaN(d)) {
                days.add(d.getDate());
                months.add(d.getMonth() + 1);
                years.add(d.getFullYear());
            }
        }
    });

    const daySel = document.getElementById('bb-filter-day');
    const monthSel = document.getElementById('bb-filter-month');
    const yearSel = document.getElementById('bb-filter-year');

    daySel.innerHTML = '<option value="All">ทั้งหมด</option>';
    [...days].sort((a, b) => a - b).forEach(d => {
        const opt = document.createElement('option');
        opt.value = d; opt.textContent = String(d).padStart(2, '0');
        daySel.appendChild(opt);
    });

    monthSel.innerHTML = '<option value="All">ทั้งหมด</option>';
    [...months].sort((a, b) => a - b).forEach(m => {
        const opt = document.createElement('option');
        opt.value = m; opt.textContent = `${String(m).padStart(2, '0')} - ${monthNames[m - 1]}`;
        monthSel.appendChild(opt);
    });

    yearSel.innerHTML = '<option value="All">ทั้งหมด</option>';
    [...years].sort((a, b) => b - a).forEach(y => {
        const opt = document.createElement('option');
        opt.value = y; opt.textContent = y;
        yearSel.appendChild(opt);
    });
}

// -------------------------------------------------
// BANK DETAIL MODAL
// -------------------------------------------------
function openBankDetailModal(bankFullName, bankType, accountNum) {
    _currentBankName = bankFullName;

    const day = document.getElementById('bb-filter-day')?.value || 'All';
    const month = document.getElementById('bb-filter-month')?.value || 'All';
    const year = document.getElementById('bb-filter-year')?.value || 'All';

    // Helper: แยกชื่อธนาคาร (ส่วนก่อน "-" แรก)
    function extractBankType(str) {
        return (str || '').split('-')[0].trim().toUpperCase();
    }
    // Helper: เอาแค่ตัวเลข 4 หลักสุดท้าย
    function last4digits(str) {
        return (str || '').replace(/\D/g, '').slice(-4);
    }

    const bankTypeUpper = bankType.toUpperCase();
    const acctLast4 = last4digits(accountNum);

    const rows = allTransactions.filter(row => {
        const b = (row['Bank'] || row.bank || '').trim();

        // ① Exact match
        if (b === bankFullName) {
            // pass → ไปเช็ค date ด้านล่าง
        }
        // ② ชื่อธนาคารตรง + 4 หลักสุดท้ายของบัญชีตรงกัน
        else if (acctLast4 && extractBankType(b) === bankTypeUpper && last4digits(b) === acctLast4) {
            // pass
        }
        // ③ Fallback: match แค่ชื่อธนาคาร เช่น "KBANK" ทุกบัญชี
        else if (extractBankType(b) === bankTypeUpper) {
            // pass
        }
        else {
            return false;
        }

        const rawDate = row['Date'] || row.date;
        if (rawDate && (day !== 'All' || month !== 'All' || year !== 'All')) {
            const d = new Date(rawDate);
            if (!isNaN(d)) {
                if (day !== 'All' && d.getDate() !== Number(day)) return false;
                if (month !== 'All' && (d.getMonth() + 1) !== Number(month)) return false;
                if (year !== 'All' && d.getFullYear() !== Number(year)) return false;
            }
        }
        return true;
    });

    _bankModalRows = rows;

    const filterLabel = [day !== 'All' ? `วัน ${day}` : '', month !== 'All' ? `เดือน ${month}` : '', year !== 'All' ? `ปี ${year}` : ''].filter(Boolean).join(' / ');
    const title = accountNum ? `🏦 ${bankType}  (เลขที่บัญชี: ${accountNum})` : `🏦 ${bankType}`;
    document.getElementById('bank-modal-title').textContent = title;
    document.getElementById('bank-modal-subtitle').textContent = filterLabel ? `กรอง: ${filterLabel}` : 'แสดงทุกรายการ';
    document.getElementById('bank-modal-search').value = '';

    renderBankDetailRows(rows);
    document.getElementById('bank-detail-modal').classList.add('active');
    document.body.style.overflow = 'hidden';
}

let _bankModalViewMode = 'all';

function updateBankModalView(mode) {
    _bankModalViewMode = mode;
    document.getElementById('bank-btn-view-all').classList.toggle('active', mode === 'all');
    document.getElementById('bank-btn-view-group').classList.toggle('active', mode === 'group');
    filterBankModalTable();
}

function renderBankDetailRows(rows) {
    const tbody = document.getElementById('bank-modal-table-body');
    const thead = document.getElementById('bank-modal-table-head');
    tbody.innerHTML = '';
    let totalIn = 0, totalOut = 0;

    if (_bankModalViewMode === 'group') {
        thead.innerHTML = `<tr><th>#</th><th>Category</th><th>จำนวนรายการ</th><th class="numeric">Cash In (฿)</th><th class="numeric">Cash Out (฿)</th></tr>`;
        
        const grouped = {};
        rows.forEach(row => {
            const cat = row['Category'] || row.category || 'ไม่ระบุหมวดหมู่';
            if (!grouped[cat]) grouped[cat] = { count: 0, in: 0, out: 0 };
            const cashIn = Number(row['Cash In'] || row.cashIn || row['Amount'] || row['Amount (THB)'] || row.amount) || 0;
            const cashOut = Number(row['Cash Out'] || row.cashOut || row['# Amount']) || 0;
            
            grouped[cat].count++;
            grouped[cat].in += cashIn;
            grouped[cat].out += cashOut;
        });

        const sortedKeys = Object.keys(grouped).sort((a,b) => (grouped[b].in + grouped[b].out) - (grouped[a].in + grouped[a].out));
        let totalCount = 0;
        
        sortedKeys.forEach((cat, i) => {
            const item = grouped[cat];
            totalIn += item.in;
            totalOut += item.out;
            totalCount += item.count;
            
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${i + 1}</td>
                <td>${cat}</td>
                <td>${item.count} รายการ</td>
                <td class="numeric modal-amount-income">${item.in > 0 ? '฿' + checkValue(item.in) : '-'}</td>
                <td class="numeric modal-amount-expense">${item.out > 0 ? '฿' + checkValue(item.out) : '-'}</td>
            `;
            tbody.appendChild(tr);
        });
        document.getElementById('bank-modal-row-count').textContent = `รวม ${totalCount} รายการ (${sortedKeys.length} หมวดหมู่)`;
    } else {
        thead.innerHTML = `<tr><th>#</th><th>วันที่</th><th>คำอธิบาย</th><th>ประเภท</th><th>Category</th><th>Status</th><th class="numeric">Cash In (฿)</th><th class="numeric">Cash Out (฿)</th></tr>`;
        
        rows.forEach((row, i) => {
            const rawDate = row['Date'] || row.date || '';
            let displayDate = rawDate;
            try {
                const d = new Date(rawDate);
                if (!isNaN(d)) displayDate = d.toLocaleDateString('th-TH', { day: '2-digit', month: '2-digit', year: 'numeric' });
            } catch (e) { }

            const desc = row['Description'] || row.description || '-';
            const type = row['Type'] || row.type || '-';
            const category = row['Category'] || row.category || '-';
            const status = row['Status'] || row.status || '-';
            const cashIn = Number(row['Cash In'] || row.cashIn || row['Amount'] || row['Amount (THB)'] || row.amount) || 0;
            const cashOut = Number(row['Cash Out'] || row.cashOut || row['# Amount']) || 0;

            totalIn += cashIn;
            totalOut += cashOut;

            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${i + 1}</td>
                <td>${displayDate}</td>
                <td title="${desc}">${desc}</td>
                <td><span class="type-${type.toLowerCase()}">${type}</span></td>
                <td>${category}</td>
                <td>${status}</td>
                <td class="numeric modal-amount-income">${cashIn > 0 ? '฿' + checkValue(cashIn) : '-'}</td>
                <td class="numeric modal-amount-expense">${cashOut > 0 ? '฿' + checkValue(cashOut) : '-'}</td>
            `;
            tbody.appendChild(tr);
        });
        document.getElementById('bank-modal-row-count').textContent = `${rows.length} รายการ`;
    }

    document.getElementById('bank-modal-totals').innerHTML =
        `รับเข้า: <span class="modal-amount-income">฿${checkValue(totalIn)}</span> &nbsp;|&nbsp; จ่ายออก: <span class="modal-amount-expense">฿${checkValue(totalOut)}</span>`;
}

function filterBankModalTable() {
    const q = (document.getElementById('bank-modal-search')?.value || '').toLowerCase();
    if (!q) { renderBankDetailRows(_bankModalRows); return; }
    const filtered = _bankModalRows.filter(row => {
        return [
            row['Description'], row.description,
            row['Type'], row.type,
            row['Category'], row.category,
            row['Status'], row.status
        ].some(v => (v || '').toString().toLowerCase().includes(q));
    });
    renderBankDetailRows(filtered);
}

function closeBankDetailModal(event, force = false) {
    if (force || (event && event.target === document.getElementById('bank-detail-modal'))) {
        document.getElementById('bank-detail-modal').classList.remove('active');
        document.body.style.overflow = '';
    }
}


// -------------------------------------------------
// EVENT LISTENERS for Filters
// -------------------------------------------------
document.addEventListener('DOMContentLoaded', () => {
    initDashboard();

    // Listen to filter inputs (exclude creditor - handled by autocomplete)
    ['filter-bank', 'filter-type', 'filter-day', 'filter-month', 'filter-year'].forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('input', applyFilters);
            el.addEventListener('change', applyFilters);
        }
    });

    // filter-creditor handled by initCreditorMultiSelect

    // Reset button
    document.getElementById('btn-reset-filters')?.addEventListener('click', () => {
        selectedCreditors.clear();
        updateCreditorSelectText();
        document.getElementById('filter-creditor-search').value = '';
        document.getElementById('filter-bank').value = 'All';
        document.getElementById('filter-type').value = 'All';
        document.getElementById('filter-day').value = 'All';
        document.getElementById('filter-month').value = 'All';
        document.getElementById('filter-year').value = 'All';
        applyFilters();
    });

    // Bank Balance date filters
    ['bb-filter-day', 'bb-filter-month', 'bb-filter-year'].forEach(id => {
        document.getElementById(id)?.addEventListener('change', () => {
            // Re-render cards with same bank filter that's already applied
            const bank = document.getElementById('filter-bank')?.value || 'All';
            if (bank !== 'All') {
                renderBankBalances(bankBalances.filter(b =>
                    (b.bank || '').trim().toLowerCase() === bank.trim().toLowerCase()
                ));
            } else {
                renderBankBalances(bankBalances);
            }
        });
    });

    document.getElementById('bb-btn-reset')?.addEventListener('click', () => {
        document.getElementById('bb-filter-day').value = 'All';
        document.getElementById('bb-filter-month').value = 'All';
        document.getElementById('bb-filter-year').value = 'All';
        renderBankBalances(bankBalances);
    });
});

// -------------------------------------------------
// DETAIL MODAL
// -------------------------------------------------

// Store currently displayed modal rows for search filtering
let _modalRows = [];
let _modalType = '';  // 'income' | 'expense'

const MODAL_LABELS = {
    'income-actual': { title: '📥 Income (Actual)', color: 'income' },
    'income-plan': { title: '📋 Income (Plan)', color: 'income' },
    'expense-actual': { title: '📤 Expense (Actual)', color: 'expense' },
    'expense-plan': { title: '📋 Expense (Plan)', color: 'expense' },
};

function openDetailModal(cardId) {
    const meta = MODAL_LABELS[cardId];
    if (!meta) return;
    _modalType = meta.color;

    let rows = [];
    if (cardId === 'income-actual') {
        rows = _lastFilteredTransactions.filter(row => {
            const s = (row['Status'] || row.status || '').toLowerCase();
            return getRowType(row) === 'income' && s !== 'plan';
        });
    } else if (cardId === 'income-plan') {
        const fromPlans = _lastFilteredPlans.filter(row => {
            const s = (row['Status'] || row.status || '').toLowerCase();
            return getRowType(row) === 'income' && s === 'plan';
        });
        const fromTx = _lastFilteredTransactions.filter(row => {
            const s = (row['Status'] || row.status || '').toLowerCase();
            return getRowType(row) === 'income' && s === 'plan';
        });
        rows = [...fromPlans, ...fromTx];
    } else if (cardId === 'expense-actual') {
        rows = _lastFilteredTransactions.filter(row => {
            const s = (row['Status'] || row.status || '').toLowerCase();
            return getRowType(row) === 'expense' && s !== 'plan';
        });
    } else if (cardId === 'expense-plan') {
        const fromPlans = _lastFilteredPlans.filter(row => {
            const s = (row['Status'] || row.status || '').toLowerCase();
            return getRowType(row) === 'expense' && s === 'plan';
        });
        const fromTx = _lastFilteredTransactions.filter(row => {
            const s = (row['Status'] || row.status || '').toLowerCase();
            return getRowType(row) === 'expense' && s === 'plan';
        });
        rows = [...fromPlans, ...fromTx];
    }

    _modalRows = rows;

    // Set header info
    document.getElementById('modal-title').textContent = meta.title;
    document.getElementById('modal-search').value = '';

    renderModalRows(rows);

    // Open modal
    document.getElementById('detail-modal').classList.add('active');
    document.body.style.overflow = 'hidden';
}

let _detailModalViewMode = 'all';

function updateModalView(mode) {
    _detailModalViewMode = mode;
    document.getElementById('btn-view-all').classList.toggle('active', mode === 'all');
    document.getElementById('btn-view-group').classList.toggle('active', mode === 'group');
    filterModalTable();
}

function exportModalPdf(type) {
    const isBank = type === 'bank';
    const sourceTable = document.getElementById(isBank ? 'bank-modal-table' : 'modal-table');
    const title = document.getElementById(isBank ? 'bank-modal-title' : 'modal-title').textContent;
    const mode = isBank ? _bankModalViewMode : _detailModalViewMode;

    const footerCount = document.getElementById(isBank ? 'bank-modal-row-count' : 'modal-row-count').textContent;
    const footerTotal = document.getElementById(isBank ? 'bank-modal-totals' : 'modal-total-amount').innerText;

    // Clone the visible table including all rendered rows
    const tableClone = sourceTable.cloneNode(true);
    tableClone.removeAttribute('id');

    // Inject colgroup for precise column widths to avoid wrapping (A4 portrait ~190mm usable)
    // Different widths for group vs detail view
    let colgroupHtml = '';
    if (mode === 'group') {
        if (isBank) {
            // # | Category | Count | CashIn | CashOut
            colgroupHtml = `<colgroup>
                <col style="width:6%">
                <col style="width:36%">
                <col style="width:18%">
                <col style="width:20%">
                <col style="width:20%">
            </colgroup>`;
        } else {
            // # | Category | Count | Total
            colgroupHtml = `<colgroup>
                <col style="width:6%">
                <col style="width:52%">
                <col style="width:18%">
                <col style="width:24%">
            </colgroup>`;
        }
    } else {
        if (isBank) {
            // # | Date | Desc | Type | Category | Status | CashIn | CashOut
            colgroupHtml = `<colgroup>
                <col style="width:4%">
                <col style="width:9%">
                <col style="width:24%">
                <col style="width:9%">
                <col style="width:15%">
                <col style="width:9%">
                <col style="width:15%">
                <col style="width:15%">
            </colgroup>`;
        } else {
            // # | Date | Desc | Party | Bank | Category | Status | Amount
            colgroupHtml = `<colgroup>
                <col style="width:4%">
                <col style="width:9%">
                <col style="width:22%">
                <col style="width:14%">
                <col style="width:10%">
                <col style="width:14%">
                <col style="width:9%">
                <col style="width:18%">
            </colgroup>`;
        }
    }

    // Insert colgroup right after <table tag
    let tableHtml = tableClone.outerHTML.replace(/^<table[^>]*>/, match => match + colgroupHtml);

    const printWindow = window.open('', '_blank', 'width=900,height=700');
    if (!printWindow) { alert('กรุณาอนุญาต Pop-up สำหรับเว็บนี้ก่อนครับ'); return; }

    printWindow.document.write(`<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="UTF-8">
<title>${title}</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap');
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Sarabun', sans-serif; font-size: 7.5pt; color: #111; background: #fff; padding: 18px 20px; }
  .hdr { text-align: center; border-bottom: 2px solid #1e3a5f; padding-bottom: 10px; margin-bottom: 14px; }
  .hdr h1 { font-size: 15pt; color: #1e3a5f; font-weight: 700; margin-bottom: 4px; }
  .hdr h2 { font-size: 11pt; color: #334155; font-weight: 600; margin-bottom: 4px; }
  .hdr p  { font-size: 8pt; color: #555; }
  table { width: 100%; border-collapse: collapse; margin-bottom: 14px; font-size: 7pt; table-layout: fixed; }
  thead th { background: #1e3a5f; color: #fff; padding: 7px 5px; text-align: left; font-weight: 700; border: 1px solid #1e3a5f; white-space: nowrap; overflow: hidden; font-size: 7pt; }
  tbody tr:nth-child(even) { background: #f0f4f8; }
  tbody tr:nth-child(odd)  { background: #fff; }
  tbody td { padding: 5px; border: 1px solid #d1d5db; color: #111; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .numeric { text-align: right; font-family: monospace; white-space: nowrap; }
  thead th.numeric { text-align: right; }
  .modal-amount-income { color: #16a34a !important; font-weight: 700; }
  .modal-amount-expense { color: #dc2626 !important; font-weight: 700; }
  .ftr { border-top: 2px solid #1e3a5f; padding-top: 8px; display: flex; justify-content: space-between; font-weight: 700; font-size: 9pt; color: #1e3a5f; }
  @media print { @page { size: A4 portrait; margin: 1cm; } body { padding: 0; } }
</style>
</head>
<body>
<div class="hdr">
  <h1>รายงานสรุปข้อมูลทางการเงิน</h1>
  <h2>${title}</h2>
  <p>รูปแบบ: ${mode === 'group' ? 'สรุปตามหมวดหมู่' : 'รายการละเอียด'} &nbsp;|&nbsp; วันที่เรียกดู: ${new Date().toLocaleString('th-TH')}</p>
</div>
${tableHtml}
<div class="ftr"><span>${footerCount}</span><span>${footerTotal}</span></div>
<script>window.onload=function(){window.print();}<\/script>
</body></html>`);
    printWindow.document.close();
}


function renderModalRows(rows) {
    const tbody = document.getElementById('modal-table-body');
    const thead = document.getElementById('modal-table-head');
    tbody.innerHTML = '';
    let total = 0;

    if (_detailModalViewMode === 'group') {
        thead.innerHTML = `<tr><th>#</th><th>Category</th><th>จำนวนรายการ</th><th class="numeric">ยอดรวม (฿)</th></tr>`;
        
        const grouped = {};
        rows.forEach(row => {
            const cat = row['Category'] || row.category || 'ไม่ระบุหมวดหมู่';
            if (!grouped[cat]) grouped[cat] = { count: 0, sum: 0 };
            grouped[cat].count++;
            grouped[cat].sum += getRowAmount(row, _modalType);
        });
        
        const sortedKeys = Object.keys(grouped).sort((a, b) => grouped[b].sum - grouped[a].sum);
        let totalCount = 0;
        
        sortedKeys.forEach((cat, i) => {
            const item = grouped[cat];
            total += item.sum;
            totalCount += item.count;
            const amtClass = _modalType === 'income' ? 'modal-amount-income' : 'modal-amount-expense';
            
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${i + 1}</td>
                <td>${cat}</td>
                <td>${item.count} รายการ</td>
                <td class="numeric ${amtClass}">฿${checkValue(item.sum)}</td>
            `;
            tbody.appendChild(tr);
        });
        
        document.getElementById('modal-row-count').textContent = `รวม ${totalCount} รายการ (${sortedKeys.length} หมวดหมู่)`;
    } else {
        thead.innerHTML = `<tr><th>#</th><th>วันที่</th><th>คำอธิบาย</th><th>เจ้าหนี้ / ลูกหนี้</th><th>Bank</th><th>Category</th><th>Status</th><th class="numeric">จำนวนเงิน (฿)</th></tr>`;
        
        rows.forEach((row, i) => {
            const rawDate = row['Date'] || row.date || '';
            let displayDate = rawDate;
            try {
                const d = new Date(rawDate);
                if (!isNaN(d)) displayDate = d.toLocaleDateString('th-TH', { day: '2-digit', month: '2-digit', year: 'numeric' });
            } catch (e) { }

            const desc = row['Description'] || row.description || '-';
            const creditor = row['Customer'] || row['Vendor'] || row['Party'] || row['Name'] || row.customer || row.party || row.name || '-';
            const bank = row['Bank'] || row.bank || '-';
            const category = row['Category'] || row.category || '-';
            const status = row['Status'] || row.status || '-';

            const numAmt = getRowAmount(row, _modalType);
            total += numAmt;

            const amtClass = _modalType === 'income' ? 'modal-amount-income' : 'modal-amount-expense';

            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${i + 1}</td>
                <td>${displayDate}</td>
                <td title="${desc}">${desc}</td>
                <td title="${creditor}">${creditor}</td>
                <td>${bank}</td>
                <td>${category}</td>
                <td>${status}</td>
                <td class="numeric ${amtClass}">฿${checkValue(numAmt)}</td>
            `;
            tbody.appendChild(tr);
        });
        document.getElementById('modal-row-count').textContent = `${rows.length} รายการ`;
    }

    const amtClass = _modalType === 'income' ? 'modal-amount-income' : 'modal-amount-expense';
    document.getElementById('modal-total-amount').innerHTML = `ยอดรวม: <span class="${amtClass}">฿${checkValue(total)}</span>`;
}

function filterModalTable() {
    const q = (document.getElementById('modal-search')?.value || '').toLowerCase();
    if (!q) {
        renderModalRows(_modalRows);
        return;
    }
    const filtered = _modalRows.filter(row => {
        const fields = [
            row['Description'], row.description,
            row['Customer'], row.customer,
            row['Vendor'], row.vendor,
            row['Party'], row.party,
            row['Name'], row.name,
            row['Bank'], row.bank,
            row['Category'], row.category,
            row['Status'], row.status,
        ].map(v => (v || '').toString().toLowerCase());
        return fields.some(f => f.includes(q));
    });
    renderModalRows(filtered);
}

function closeDetailModal(event, force = false) {
    if (force || (event && event.target === document.getElementById('detail-modal'))) {
        document.getElementById('detail-modal').classList.remove('active');
        document.body.style.overflow = '';
    }
}

// ESC key to close any open modal
document.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
        closeDetailModal(null, true);
        closeBankDetailModal(null, true);
    }
});

// -------------------------------------------------
// CREDITOR MULTI-SELECT (All_Party sheet)
// -------------------------------------------------
function updateCreditorSelectText() {
    const textEl = document.getElementById('creditor-selected-text');
    if (!textEl) return;
    if (selectedCreditors.size === 0) {
        textEl.textContent = 'ทั้งหมด';
        textEl.style.color = '';
    } else if (selectedCreditors.size === 1) {
        textEl.textContent = Array.from(selectedCreditors)[0];
        textEl.style.color = 'var(--primary)';
    } else {
        textEl.textContent = `เลือกแล้ว (${selectedCreditors.size})`;
        textEl.style.color = 'var(--primary)';
    }
}

function initCreditorAutocomplete() { // renamed internally but keeping the original name for initDashboard call
    const toggleBox = document.getElementById('btn-creditor-dropdown');
    const dropdown = document.getElementById('creditor-dropdown');
    const searchInput = document.getElementById('filter-creditor-search');
    const suggestionsList = document.getElementById('creditor-suggestions');
    const btnSelectAll = document.getElementById('btn-ms-select-all');
    const btnClear = document.getElementById('btn-ms-clear');

    if (!toggleBox || !dropdown) return;

    let currentMatches = [];

    // Toggle Dropdown
    toggleBox.addEventListener('click', (e) => {
        e.stopPropagation();
        const isOpen = dropdown.classList.contains('open');
        dropdown.classList.toggle('open', !isOpen);
        if (!isOpen) {
            searchInput.focus();
            renderMsList(searchInput.value.trim());
        }
    });

    // Close when clicking outside
    document.addEventListener('click', e => {
        if (!toggleBox.contains(e.target) && !dropdown.contains(e.target)) {
            dropdown.classList.remove('open');
        }
    });

    // Allow clicking inside dropdown without closing it
    dropdown.addEventListener('click', e => {
        e.stopPropagation();
    });

    function renderMsList(query) {
        suggestionsList.innerHTML = '';
        const q = query.toLowerCase();

        let matches = allParties;
        if (q) {
            matches = allParties.filter(p => p.toLowerCase().includes(q));
            matches.sort((a, b) => {
                const aStart = a.toLowerCase().startsWith(q);
                const bStart = b.toLowerCase().startsWith(q);
                if (aStart && !bStart) return -1;
                if (!aStart && bStart) return 1;
                return a.localeCompare(b, 'th');
            });
        }

        currentMatches = matches;

        if (matches.length === 0) {
            suggestionsList.innerHTML = '<div class="autocomplete-empty">ไม่พบชื่อที่ตรงกัน</div>';
            return;
        }

        // Display all matches (or cap at 100 to prevent lag)
        matches.slice(0, 100).forEach(name => {
            const item = document.createElement('label');
            item.className = 'ms-item';

            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.className = 'ms-checkbox';
            checkbox.value = name;
            checkbox.checked = selectedCreditors.has(name);

            checkbox.addEventListener('change', (e) => {
                if (e.target.checked) {
                    selectedCreditors.add(name);
                } else {
                    selectedCreditors.delete(name);
                }
                updateCreditorSelectText();
                applyFilters();
            });

            const span = document.createElement('span');
            span.className = 'ms-item-name';

            // Highlight text
            if (q) {
                const idx = name.toLowerCase().indexOf(q);
                if (idx !== -1) {
                    span.innerHTML = name.slice(0, idx) +
                        '<mark style="background:transparent; color:var(--primary); font-weight:bold;">' +
                        name.slice(idx, idx + q.length) + '</mark>' +
                        name.slice(idx + q.length);
                } else {
                    span.textContent = name;
                }
            } else {
                span.textContent = name;
            }

            item.appendChild(checkbox);
            item.appendChild(span);
            suggestionsList.appendChild(item);
        });
    }

    searchInput.addEventListener('input', () => {
        renderMsList(searchInput.value.trim());
    });

    btnSelectAll.addEventListener('click', () => {
        currentMatches.forEach(name => {
            selectedCreditors.add(name);
        });
        renderMsList(searchInput.value.trim());
        updateCreditorSelectText();
        applyFilters();
    });

    btnClear.addEventListener('click', () => {
        selectedCreditors.clear();
        renderMsList(searchInput.value.trim());
        updateCreditorSelectText();
        applyFilters();
    });
}

// -------------------------------------------------
// TOGGLE BALANCES (SHOW/HIDE)
// -------------------------------------------------
function toggleBalances() {
    const body = document.body;
    const btn = document.getElementById('btn-toggle-balance');
    
    if (body.classList.contains('hide-balances')) {
        body.classList.remove('hide-balances');
        btn.innerHTML = '🔒 ซ่อนยอดเงิน';
    } else {
        body.classList.add('hide-balances');
        btn.innerHTML = '👁️ แสดงทั้งหมด';
    }
}
