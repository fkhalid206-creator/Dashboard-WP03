// Initialize Chart variables to destroy them on reload
let weeklyQtyChart = null;
let weeklyValueChart = null;
let dailyQtyChart = null;
let dailyValueChart = null;
let deptQtyChart = null;
let deptValueChart = null;
let materialQtyChart = null;
let materialValueChart = null;

// Register DataLabels plugin globally
Chart.register(ChartDataLabels);

// DOM Elements
const csvFileInput = document.getElementById('csvFileInput');
const fileStatus = document.getElementById('fileStatus');
const emptyState = document.getElementById('emptyState');
const dashboardDataArea = document.getElementById('dashboardData');
const dataLoadedBadge = document.getElementById('dataLoadedBadge');

// Set Current Date
document.getElementById('currentDate').textContent = new Date().toLocaleDateString('en-US', {
    weekday: 'long', year: 'numeric', month: 'long', day: 'numeric'
});

// Event Listener for File Upload
csvFileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
        fileStatus.textContent = file.name;

        Papa.parse(file, {
            header: true,
            skipEmptyLines: true,
            dynamicTyping: true,
            complete: function (results) {
                if (results.data && results.data.length > 0) {
                    processData(results.data);
                } else {
                    alert('No data found in the CSV file.');
                }
            },
            error: function (err) {
                console.error("PapaParse Error:", err);
                alert("Error parsing CSV file.");
            }
        });
    }
});

// Main Data Processing Function
function processData(data) {
    // Reveal Dashboard
    emptyState.classList.add('hidden');
    dashboardDataArea.classList.remove('hidden');
    dataLoadedBadge.classList.remove('hidden');

    let totalValue = 0;
    let totalQty = 0;
    let totalTransactions = data.length;
    let uniqueItems = new Set();

    // Grouping objects for 8 charts
    let deptStats = {};
    let materialStats = {};
    let storekeeperStats = {};
    let dailyStats = {};
    let weeklyStats = {};

    data.forEach(row => {
        // Safe value extraction
        const value = parseFloat(row["Issued Value"]) || parseFloat(row["Value"]) || 0;
        const qty = parseFloat(row["Issued Qty"]) || parseFloat(row["Quantity"]) || 0;

        let rawDate = row["Issue Date"] || row["Posting Date"] || row["Transaction Date"] || row["Date"] || row["Date "] || null;
        let dateObj = null;
        if (rawDate) {
            // Check for DD/MM/YYYY or DD-MM-YY or DD-MM-YYYY format
            const dmyMatch = String(rawDate).match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
            if (dmyMatch) {
                // Parse as YYYY-MM-DD
                let year = parseInt(dmyMatch[3], 10);
                if (year < 100) year += 2000; // handle 2-digit year (e.g. 26 -> 2026)
                dateObj = new Date(year, parseInt(dmyMatch[2], 10) - 1, parseInt(dmyMatch[1], 10));
            } else {
                dateObj = new Date(rawDate);
            }
            if (isNaN(dateObj)) dateObj = null;
        }

        const dept = row["DEPARTMENT"] || row["Department"] || "Unknown Dept";
        const desc = row["Description"] || row["Material Description"] || row["Material"] || row["Item Name"] || "Unknown Material";
        const itemCode = row["Item Code"] || desc;

        // Accumulate KPIs
        totalValue += value;
        totalQty += qty;
        if (itemCode) uniqueItems.add(itemCode);

        // Grouping logic helper
        const addStat = (obj, key, v, q) => {
            if (!obj[key]) obj[key] = { value: 0, qty: 0 };
            obj[key].value += v;
            obj[key].qty += q;
        };

        // 1. Department
        addStat(deptStats, dept, value, qty);

        // 2. Material
        addStat(materialStats, desc, value, qty);

        // 5. Storekeeper
        const storekeeper = row["Issued By"] || row["User"] || row["Storekeeper"] || "Unknown";
        addStat(storekeeperStats, storekeeper, value, qty);

        // Dates for Trends
        if (dateObj) {
            // 3. Daily
            const dayStr = dateObj.getFullYear() + '-' + String(dateObj.getMonth() + 1).padStart(2, '0') + '-' + String(dateObj.getDate()).padStart(2, '0');
            addStat(dailyStats, dayStr, value, qty);
        }

        // 4. Weekly (Prefer explicit WEEK column from CSV)
        if (row["WEEK"]) {
            addStat(weeklyStats, row["WEEK"], value, qty);
        } else if (dateObj) {
            // Fallback weekly calculation
            const startOfYear = new Date(dateObj.getFullYear(), 0, 1);
            const days = Math.floor((dateObj - startOfYear) / (24 * 60 * 60 * 1000));
            const weekNum = Math.ceil((dateObj.getDay() + 1 + days) / 7);
            const weekStr = `${dateObj.getFullYear()}-W${weekNum}`;
            addStat(weeklyStats, weekStr, value, qty);
        }
    });

    // Calculate Additional Audit KPIs
    let movingMaterialsCount = uniqueItems.size;
    let nonMovingMaterialsCount = 0; // Assuming the CSV only contains issuance data, all items listed moved.

    let dailyValues = Object.values(dailyStats).map(day => day.value);
    let dailyQtys = Object.values(dailyStats).map(day => day.qty);

    let avgDailyVal = dailyValues.length > 0 ? totalValue / dailyValues.length : 0;
    let avgDailyQty = dailyQtys.length > 0 ? totalQty / dailyQtys.length : 0;

    let highDailyVal = dailyValues.length > 0 ? Math.max(...dailyValues) : 0;
    let lowDailyVal = dailyValues.length > 0 ? Math.min(...dailyValues) : 0;

    // Update KPIs on DOM
    document.getElementById('kpiUnique').textContent = formatShortNumber(uniqueItems.size, false).replace(' Units', '');
    document.getElementById('kpiQty').textContent = formatShortNumber(totalQty, false).replace(' Units', '');
    document.getElementById('kpiValue').textContent = formatShortNumber(totalValue, true).replace('SAR ', ''); // keep span class handling SAR
    document.getElementById('kpiTrans').textContent = formatShortNumber(totalTransactions, false).replace(' Units', '');

    // Additional Audit KPIs
    if (document.getElementById('kpiMoving')) {
        document.getElementById('kpiMoving').textContent = formatShortNumber(movingMaterialsCount, false).replace(' Units', '');
        document.getElementById('kpiNonMoving').textContent = formatShortNumber(nonMovingMaterialsCount, false).replace(' Units', '');
        document.getElementById('kpiAvgQty').textContent = formatShortNumber(avgDailyQty, false).replace(' Units', '');
        document.getElementById('kpiAvgVal').textContent = formatShortNumber(avgDailyVal, true);
        document.getElementById('kpiHighVal').textContent = formatShortNumber(highDailyVal, true);
        document.getElementById('kpiLowVal').textContent = formatShortNumber(lowDailyVal, true);
    }

    // Render 9 Charts
    renderTop10BarChart('deptQtyChart', deptStats, 'qty', 'Top 10 Departments by Units', 'Units', deptQtyChart, inst => deptQtyChart = inst, '#3b82f6');
    renderTop10BarChart('deptValueChart', deptStats, 'value', 'Top 10 Departments by Currency', 'Currency', deptValueChart, inst => deptValueChart = inst, '#8b5cf6');
    renderStorekeeperBarChart('storekeeperChart', storekeeperStats, 'qty', 'Material Issuance Distribution by Storekeeper', 'Units', window.storekeeperChartInst, inst => window.storekeeperChartInst = inst, '#f59e0b');

    renderTop10BarChart('materialQtyChart', materialStats, 'qty', 'Fast Moving Materials (Units)', 'Units', materialQtyChart, inst => materialQtyChart = inst, '#14b8a6');
    renderTop10BarChart('materialValueChart', materialStats, 'value', 'Fast Moving Materials (Currency)', 'Currency', materialValueChart, inst => materialValueChart = inst, '#f43f5e');

    renderTrendLineChart('dailyQtyChart', dailyStats, 'qty', 'Daily Trend (Units)', 'Units', dailyQtyChart, inst => dailyQtyChart = inst, '#f59e0b');
    renderTrendLineChart('dailyValueChart', dailyStats, 'value', 'Daily Trend (Currency)', 'Currency', dailyValueChart, inst => dailyValueChart = inst, '#2563eb');

    renderTrendLineChart('weeklyQtyChart', weeklyStats, 'qty', 'Weekly Trend (Units)', 'Units', weeklyQtyChart, inst => weeklyQtyChart = inst, '#4f46e5');
    renderTrendLineChart('weeklyValueChart', weeklyStats, 'value', 'Weekly Trend (Currency)', 'Currency', weeklyValueChart, inst => weeklyValueChart = inst, '#059669');
}
// Chart Renderers
function renderTop10BarChart(canvasId, dataStats, metric, label, axisLabel, chartInstance, setInstance, color) {
    const ctx = document.getElementById(canvasId).getContext('2d');

    // Convert to Array, Sort Descending, Take Top 10
    const sortedData = Object.entries(dataStats)
        .sort((a, b) => b[1][metric] - a[1][metric])
        .slice(0, 10);

    const labels = sortedData.map(item => {
        let labelStr = String(item[0]);
        if (canvasId.includes('material')) {
            labelStr = shortenMaterialName(labelStr);
        }
        if (labelStr.length <= 30) return labelStr;

        // Wrap text to multiple lines if longer than 30 chars
        const words = labelStr.split(' ');
        let lines = [''];
        let currentLine = 0;
        for (const word of words) {
            if ((lines[currentLine] + word).length > 30 && lines[currentLine].length > 0) {
                currentLine++;
                lines[currentLine] = '';
            }
            lines[currentLine] += (lines[currentLine].length > 0 ? ' ' : '') + word;
        }
        return lines;
    });
    // Ensure proper rounding before chart rendering
    const data = sortedData.map(item => {
        let val = item[1][metric];
        return metric === 'qty' ? Math.round(val) : Number(val.toFixed(2));
    });

    if (chartInstance) chartInstance.destroy();

    const newInst = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: label,
                data: data,
                fullLabels: sortedData.map(item => String(item[0])),
                backgroundColor: color,
                borderRadius: 4,
                hoverBackgroundColor: color + 'cc' // slight transparency on hover
            }]
        },
        options: {
            ...getChartOptions(axisLabel),
            indexAxis: 'y', // horizontal bar
            font: { family: 'Inter' },
            plugins: {
                ...getChartOptions(axisLabel).plugins,
                datalabels: {
                    anchor: 'end',
                    align: 'end',
                    offset: 6, // Distance from bar end
                    clamp: true, // Keep within chart boundaries
                    color: '#334155',
                    font: { weight: 'bold', size: 12 },
                    formatter: function (val) {
                        return formatShortNumber(val, axisLabel === 'Currency').replace(' Units', '');
                    }
                }
            }
        }
    });
    setInstance(newInst);
}

function renderStorekeeperBarChart(canvasId, dataStats, metric, label, axisLabel, chartInstance, setInstance, color) {
    const ctx = document.getElementById(canvasId).getContext('2d');

    // Convert to Array, Sort Descending
    const sortedData = Object.entries(dataStats)
        .sort((a, b) => b[1][metric] - a[1][metric]);

    const labels = sortedData.map(item => String(item[0]));
    const data = sortedData.map(item => {
        let val = item[1][metric];
        return metric === 'qty' ? Math.round(val) : Number(val.toFixed(2));
    });

    if (chartInstance) chartInstance.destroy();

    const newInst = new Chart(ctx, {
        type: 'bar', // vertical bar for storekeepers
        data: {
            labels: labels,
            datasets: [{
                label: label,
                data: data,
                backgroundColor: color,
                borderRadius: 4,
                hoverBackgroundColor: color + 'cc'
            }]
        },
        options: {
            ...getChartOptions(axisLabel),
            font: { family: 'Inter' },
            plugins: {
                ...getChartOptions(axisLabel).plugins,
                datalabels: {
                    anchor: 'end',
                    align: 'top',
                    offset: 6,
                    clamp: true,
                    color: '#334155',
                    font: { weight: 'bold', size: 12 },
                    formatter: function (val) {
                        return formatShortNumber(val, axisLabel === 'Currency').replace(' Units', '');
                    }
                }
            }
        }
    });
    setInstance(newInst);
}

function renderTrendLineChart(canvasId, dataStats, metric, label, axisLabel, chartInstance, setInstance, color) {
    const ctx = document.getElementById(canvasId).getContext('2d');

    // Sort chronologically by keys
    const sortedKeys = Object.keys(dataStats).sort();

    // Convert to Date and format as DD-MMM if daily
    let labels = sortedKeys;
    if (canvasId.includes('daily')) {
        labels = sortedKeys.map(k => {
            const d = new Date(k);
            return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short' });
        });
    }

    // Ensure proper rounding before chart rendering
    const data = sortedKeys.map(k => {
        let val = dataStats[k][metric];
        return metric === 'qty' ? Math.round(val) : Number(val.toFixed(2));
    });

    if (chartInstance) chartInstance.destroy();

    const newInst = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: label,
                data: data,
                borderColor: color,
                backgroundColor: color + '1a', // 10% opacity
                borderWidth: 2,
                tension: 0.3,
                fill: true,
                pointBackgroundColor: '#ffffff',
                pointBorderColor: color,
                pointBorderWidth: 2,
                pointRadius: 3,
                pointHoverRadius: 5
            }]
        },
        options: {
            ...getChartOptions(axisLabel),
            layout: {
                padding: { top: 20, right: 20, left: 10, bottom: 10 }
            },
            plugins: {
                ...getChartOptions(axisLabel).plugins,
                datalabels: {
                    align: 'top',
                    offset: 8, // Distance above point
                    color: color,
                    font: { weight: 'bold', size: 12 },
                    display: function (context) {
                        // Only show label every 3rd point to avoid clutter
                        const len = context.dataset.data.length;
                        return context.dataIndex === 0 || context.dataIndex === len - 1 || context.dataIndex % 3 === 0;
                    },
                    formatter: function (val) {
                        return formatShortNumber(val, axisLabel === 'Currency').replace(' Units', '');
                    }
                },
                tooltip: {
                    mode: 'index',
                    intersect: false,
                    callbacks: {
                        label: function (context) {
                            return context.dataset.label + ': ' + formatShortNumber(context.raw, axisLabel === 'Currency', true);
                        }
                    }
                }
            },
            scales: {
                ...getChartOptions(axisLabel).scales,
                x: {
                    ...getChartOptions(axisLabel).scales.x,
                    ticks: {
                        color: '#475569',
                        font: { family: 'Inter', size: 11 },
                        maxRotation: 45,
                        minRotation: 45
                    }
                }
            }
        }
    });
    setInstance(newInst);
}

// Helpers
// Helpers
function formatShortNumber(val, isCurrency = false, isFull = false) {
    if (!isFull) {
        let formatted = "";
        let roundedVal = isCurrency ? Number(val.toFixed(2)) : Math.round(val);

        if (roundedVal >= 1000000) {
            formatted = (roundedVal / 1000000).toFixed(1).replace(/\.0$/, '') + ' M';
        } else if (roundedVal >= 1000) {
            formatted = (roundedVal / 1000).toFixed(1).replace(/\.0$/, '') + ' K';
        } else {
            formatted = roundedVal.toLocaleString('en-US');
        }

        return isCurrency ? 'SAR ' + formatted : formatted + ' Units';
    } else {
        // Full representation for tooltips
        let formatted = isCurrency ?
            Number(val.toFixed(2)).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) :
            Math.round(val).toLocaleString('en-US');
        return isCurrency ? 'SAR ' + formatted : formatted + ' Units';
    }
}

function getChartOptions(yAxisLabel) {
    return {
        responsive: true,
        maintainAspectRatio: false,
        layout: {
            padding: {
                left: 10,
                right: 60, // Ensure value labels have space on right
                top: 20,
                bottom: 10
            }
        },
        plugins: {
            legend: { display: false },
            tooltip: {
                mode: 'index',
                intersect: false,
                callbacks: {
                    title: function (context) {
                        const dataset = context[0].dataset;
                        if (dataset.fullLabels && dataset.fullLabels[context[0].dataIndex]) {
                            return dataset.fullLabels[context[0].dataIndex];
                        }
                        // Default fallback
                        return Array.isArray(context[0].label) ? context[0].label.join(' ') : context[0].label;
                    },
                    label: function (context) {
                        return context.dataset.label + ': ' + formatShortNumber(context.raw, yAxisLabel === 'Currency', true);
                    }
                }
            }
        },
        scales: {
            y: {
                beginAtZero: true,
                grid: { color: '#f1f5f9', lineWidth: 1 },
                ticks: {
                    color: '#475569',
                    font: { family: 'Inter', size: 12 },
                    crossAlign: 'far', // Aligns multiline text cleanly
                    callback: function (value, index, values) {
                        // Keep Y axis mapping for values, intercept for numbers
                        if (typeof value === 'number' && this.getLabelForValue) {
                            let label = this.getLabelForValue(value);
                            if (label === undefined || label === null || isNaN(label) === false) {
                                return formatShortNumber(value, yAxisLabel === 'Currency', false).replace(' Units', '');
                            }
                        }
                        return formatShortNumber(value, yAxisLabel === 'Currency', false).replace(' Units', '');
                    }
                }
            },
            x: {
                grid: { display: false },
                ticks: {
                    color: '#475569',
                    font: { family: 'Inter', size: 12 },
                    callback: function (value, index, values) {
                        // Fallback so it doesn't break horizontal vs vertical
                        if (typeof value === 'number' && this.getLabelForValue) {
                            let label = this.getLabelForValue(value);
                            // If it's a category chart, it uses category names, else values
                            if (label === undefined) return value;
                        }
                        return value;
                    }
                }
            }
        }
    };
}

function shortenMaterialName(labelStr) {
    if (!labelStr) return "Unknown";

    let parts = labelStr.split(';');
    let text = parts.length > 1 ? parts[1].trim() : parts[0].trim();

    // Remove variations of sizes, units, packaging 
    text = text.replace(/\b[0-9.]+\s*(PCS|PC|PKT|MTR|ML|LTR|GAL|KG|MM|CM|INCH|Z|W|V|A|PLY|BOX|BTL|ROLL|CTN|\%|X)\b/gi, '');
    text = text.replace(/\b(SIZE\s+[A-Z]+|BALE OF \d+|PART NO\.?\s*[A-Z0-9]+)\b/gi, '');
    text = text.replace(/\([^\)]*\)/g, ''); // remove parentheses
    text = text.replace(/-\s*[a-zA-Z0-9\s]+$/i, ''); // remove trailing hyphen details like "- FINE"
    // Keep letters, space, and maybe hyphens to prevent complete destruction
    text = text.replace(/[^a-zA-Z\s-]/g, ' ');

    // Clean up spaces and title case
    text = text.trim().replace(/\s+/g, ' ');
    let words = text.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase());
    text = words.join(' ');

    if (text.length > 30) {
        text = text.substring(0, 30);
        let lastSpace = text.lastIndexOf(' ');
        if (lastSpace > 10) {
            text = text.substring(0, lastSpace);
        }
    }

    return text.trim() || parts[0].trim().split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ').substring(0, 30);
}
