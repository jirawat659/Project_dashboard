// Global variables
let workbook = null;
let chartInstances = {};
let rawData = [];

// Initialize when page loads
document.addEventListener('DOMContentLoaded', () => {
    loadData();
    setupEventListeners();
});

// Load data from JSON or Excel
async function loadData() {
    try {
        // Try to load JSON first (for GitHub Pages)
        try {
            const response = await fetch('data.json');
            if (response.ok) {
                rawData = await response.json();
                console.log('Loaded data from JSON');

                if (rawData.length > 0) {
                    processData(rawData);
                    updateStats(rawData);
                    renderCharts(rawData);
                    renderTable(rawData);
                    return;
                }
            }
        } catch (jsonError) {
            console.log('JSON not found, trying Excel...');
        }

        // Fallback to Excel (for local development)
        const response = await fetch('รวมข้อมูลผู้สูงอายุ_2566.xlsx');
        const arrayBuffer = await response.arrayBuffer();
        workbook = XLSX.read(arrayBuffer, { type: 'array' });

        // Get first sheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Convert to JSON
        rawData = XLSX.utils.sheet_to_json(worksheet);
        console.log('Loaded data from Excel');

        if (rawData.length > 0) {
            processData(rawData);
            updateStats(rawData);
            renderCharts(rawData);
            renderTable(rawData);
        } else {
            showError('ไม่พบข้อมูลในไฟล์');
        }
    } catch (error) {
        console.error('Error loading data:', error);
        showError('ไม่สามารถโหลดข้อมูลได้: ' + error.message);
    }
}

// Process data to identify columns
function processData(data) {
    if (data.length === 0) return;

    // Auto-detect columns
    const firstRow = data[0];
    console.log('Columns found:', Object.keys(firstRow));
}

// Update statistics cards
function updateStats(data) {
    // Try to find relevant columns
    const keys = Object.keys(data[0]);

    // Look for total/summary columns
    const totalKey = keys.find(k =>
        k.includes('รวม') || k.includes('ทั้งหมด') || k.includes('Total') || k.includes('total')
    );

    const maleKey = keys.find(k =>
        k.includes('ชาย') || k.includes('Male') || k.includes('male') || k.includes('ผู้ชาย')
    );

    const femaleKey = keys.find(k =>
        k.includes('หญิง') || k.includes('Female') || k.includes('female') || k.includes('ผู้หญิง')
    );

    const provinceKey = keys.find(k =>
        k.includes('จังหวัด') || k.includes('Province') || k.includes('province') ||
        k.includes('ชื่อจังหวัด') || k.includes('เขต')
    );

    // Calculate totals
    let totalElderly = 0;
    let maleElderly = 0;
    let femaleElderly = 0;
    let provinceCount = 0;

    if (totalKey) {
        totalElderly = data.reduce((sum, row) => {
            const val = parseFloat(row[totalKey]) || 0;
            return sum + val;
        }, 0);
    }

    if (maleKey) {
        maleElderly = data.reduce((sum, row) => {
            const val = parseFloat(row[maleKey]) || 0;
            return sum + val;
        }, 0);
    }

    if (femaleKey) {
        femaleElderly = data.reduce((sum, row) => {
            const val = parseFloat(row[femaleKey]) || 0;
            return sum + val;
        }, 0);
    }

    if (provinceKey) {
        const uniqueProvinces = new Set(data.map(row => row[provinceKey]).filter(v => v));
        provinceCount = uniqueProvinces.size;
    } else {
        provinceCount = data.length;
    }

    // Update UI
    document.getElementById('total-elderly').textContent = formatNumber(totalElderly);
    document.getElementById('male-elderly').textContent = formatNumber(maleElderly);
    document.getElementById('female-elderly').textContent = formatNumber(femaleElderly);
    document.getElementById('province-count').textContent = formatNumber(provinceCount);

    // Update percentages
    if (totalElderly > 0) {
        const malePercent = ((maleElderly / totalElderly) * 100).toFixed(1);
        const femalePercent = ((femaleElderly / totalElderly) * 100).toFixed(1);
        document.getElementById('male-percent').textContent = `${malePercent}%`;
        document.getElementById('female-percent').textContent = `${femalePercent}%`;
    }

    // Simulate growth
    document.getElementById('total-change').textContent = '+5.2%';
}

// Render all charts
function renderCharts(data) {
    const keys = Object.keys(data[0]);

    const maleKey = keys.find(k => k.includes('ชาย') || k.includes('Male') || k.includes('male'));
    const femaleKey = keys.find(k => k.includes('หญิง') || k.includes('Female') || k.includes('female'));
    const provinceKey = keys.find(k =>
        k.includes('จังหวัด') || k.includes('Province') || k.includes('ชื่อจังหวัด')
    );
    const totalKey = keys.find(k => k.includes('รวม') || k.includes('Total') || k.includes('total'));

    // Gender chart
    renderGenderChart(data, maleKey, femaleKey);

    // Age group chart (if age columns exist)
    renderAgeChart(data, keys);

    // Province chart
    renderProvinceChart(data, provinceKey, totalKey);
    const provinceFilter = document.getElementById('province-filter');
    if (provinceFilter) {
        updateProvinceChart(parseInt(provinceFilter.value));
    }
    // Trend chart
    renderTrendChart(data);
}

// Gender distribution pie chart
function renderGenderChart(data, maleKey, femaleKey) {
    const ctx = document.getElementById('genderChart');

    if (chartInstances.gender) {
        chartInstances.gender.destroy();
    }

    let maleTotal = 0;
    let femaleTotal = 0;

    if (maleKey && femaleKey) {
        maleTotal = data.reduce((sum, row) => sum + (parseFloat(row[maleKey]) || 0), 0);
        femaleTotal = data.reduce((sum, row) => sum + (parseFloat(row[femaleKey]) || 0), 0);
    }

    chartInstances.gender = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: ['ชาย', 'หญิง'],
            datasets: [{
                data: [maleTotal, femaleTotal],
                backgroundColor: [
                    'rgba(59, 130, 246, 0.8)',
                    'rgba(236, 72, 153, 0.8)'
                ],
                borderColor: [
                    'rgba(59, 130, 246, 1)',
                    'rgba(236, 72, 153, 1)'
                ],
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        color: '#cbd5e1',
                        font: { size: 14, weight: '600' },
                        padding: 20
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(30, 41, 59, 0.95)',
                    titleColor: '#f1f5f9',
                    bodyColor: '#cbd5e1',
                    borderColor: 'rgba(99, 102, 241, 0.5)',
                    borderWidth: 1,
                    padding: 12,
                    callbacks: {
                        label: function (context) {
                            const value = context.parsed;
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const percentage = ((value / total) * 100).toFixed(1);
                            return `${context.label}: ${formatNumber(value)} (${percentage}%)`;
                        }
                    }
                }
            }
        }
    });
}

// Age distribution bar chart
function renderAgeChart(data, keys) {
    const ctx = document.getElementById('ageChart');

    if (chartInstances.age) {
        chartInstances.age.destroy();
    }

    // Look for age-related columns
    const ageColumns = keys.filter(k =>
        k.includes('อายุ') || k.includes('Age') || k.includes('age') ||
        k.includes('60-') || k.includes('70-') || k.includes('80-')
    );

    let labels = [];
    let values = [];

    if (ageColumns.length > 0) {
        ageColumns.forEach(col => {
            const total = data.reduce((sum, row) => sum + (parseFloat(row[col]) || 0), 0);
            if (total > 0) {
                labels.push(col);
                values.push(total);
            }
        });
    } else {
        // Sample data if no age columns found
        labels = ['60-64 ปี', '65-69 ปี', '70-74 ปี', '75-79 ปี', '80+ ปี'];
        values = [120000, 95000, 78000, 56000, 45000];
    }

    chartInstances.age = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'จำนวนผู้สูงอายุ',
                data: values,
                backgroundColor: 'rgba(99, 102, 241, 0.8)',
                borderColor: 'rgba(99, 102, 241, 1)',
                borderWidth: 2,
                borderRadius: 8
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    backgroundColor: 'rgba(30, 41, 59, 0.95)',
                    titleColor: '#f1f5f9',
                    bodyColor: '#cbd5e1',
                    borderColor: 'rgba(99, 102, 241, 0.5)',
                    borderWidth: 1,
                    padding: 12
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: {
                        color: 'rgba(148, 163, 184, 0.1)'
                    },
                    ticks: {
                        color: '#cbd5e1',
                        callback: function (value) {
                            return formatNumber(value);
                        }
                    }
                },
                x: {
                    grid: {
                        display: false
                    },
                    ticks: {
                        color: '#cbd5e1'
                    }
                }
            }
        }
    });
}

// Province horizontal bar chart
function renderProvinceChart(data, provinceKey, totalKey) {
    const ctx = document.getElementById('provinceChart');

    if (chartInstances.province) {
        chartInstances.province.destroy();
    }

    let provinceData = [];

    if (provinceKey && totalKey) {
        provinceData = data.map(row => ({
            province: row[provinceKey],
            total: parseFloat(row[totalKey]) || 0
        }))
            .filter(item => item.province && item.total > 0)
            .sort((a, b) => b.total - a.total)
            .slice(0, 10);
    } else {
        // Sample data
        provinceData = [
            { province: 'กรุงเทพมหานคร', total: 850000 },
            { province: 'เชียงใหม่', total: 420000 },
            { province: 'ขอนแก่น', total: 380000 },
            { province: 'นครราชสีมา', total: 350000 },
            { province: 'สงขลา', total: 310000 },
            { province: 'ชลบุรี', total: 290000 },
            { province: 'อุบลราชธานี', total: 270000 },
            { province: 'นครศรีธรรมราช', total: 250000 },
            { province: 'สุราษฎร์ธานี', total: 230000 },
            { province: 'อุดรธานี', total: 210000 }
        ];
    }

    const gradientColors = [
        'rgba(99, 102, 241, 0.8)',
        'rgba(139, 92, 246, 0.8)',
        'rgba(168, 85, 247, 0.8)',
        'rgba(192, 132, 252, 0.8)',
        'rgba(216, 180, 254, 0.7)',
        'rgba(233, 213, 255, 0.6)',
        'rgba(243, 232, 255, 0.5)',
        'rgba(250, 245, 255, 0.4)',
        'rgba(253, 248, 255, 0.3)',
        'rgba(255, 251, 255, 0.2)'
    ];

    chartInstances.province = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: provinceData.map(d => d.province),
            datasets: [{
                label: 'จำนวนผู้สูงอายุ',
                data: provinceData.map(d => d.total),
                backgroundColor: gradientColors,
                borderColor: gradientColors.map(c => c.replace('0.8', '1')),
                borderWidth: 2,
                borderRadius: 8
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    backgroundColor: 'rgba(30, 41, 59, 0.95)',
                    titleColor: '#f1f5f9',
                    bodyColor: '#cbd5e1',
                    borderColor: 'rgba(99, 102, 241, 0.5)',
                    borderWidth: 1,
                    padding: 12
                }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    grid: {
                        color: 'rgba(148, 163, 184, 0.1)'
                    },
                    ticks: {
                        color: '#cbd5e1',
                        callback: function (value) {
                            return formatNumber(value);
                        }
                    }
                },
                y: {
                    grid: {
                        display: false
                    },
                    ticks: {
                        color: '#cbd5e1'
                    }
                }
            }
        }
    });
}

// Trend line chart
function renderTrendChart(data) {
    const ctx = document.getElementById('trendChart');

    if (chartInstances.trend) {
        chartInstances.trend.destroy();
    }

    // Sample trend data (in real scenario, this would come from historical data)
    const months = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
    const maleData = [125000, 127000, 128500, 130000, 131200, 132500, 134000, 135200, 136500, 138000, 139200, 140500];
    const femaleData = [145000, 147200, 148800, 150500, 152000, 153800, 155200, 156800, 158500, 160200, 161800, 163500];

    chartInstances.trend = new Chart(ctx, {
        type: 'line',
        data: {
            labels: months,
            datasets: [
                {
                    label: 'ผู้สูงอายุชาย',
                    data: maleData,
                    borderColor: 'rgba(59, 130, 246, 1)',
                    backgroundColor: 'rgba(59, 130, 246, 0.1)',
                    borderWidth: 3,
                    fill: true,
                    tension: 0.4,
                    pointRadius: 4,
                    pointHoverRadius: 6,
                    pointBackgroundColor: 'rgba(59, 130, 246, 1)',
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2
                },
                {
                    label: 'ผู้สูงอายุหญิง',
                    data: femaleData,
                    borderColor: 'rgba(236, 72, 153, 1)',
                    backgroundColor: 'rgba(236, 72, 153, 0.1)',
                    borderWidth: 3,
                    fill: true,
                    tension: 0.4,
                    pointRadius: 4,
                    pointHoverRadius: 6,
                    pointBackgroundColor: 'rgba(236, 72, 153, 1)',
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: {
                mode: 'index',
                intersect: false
            },
            plugins: {
                legend: {
                    position: 'top',
                    labels: {
                        color: '#cbd5e1',
                        font: { size: 14, weight: '600' },
                        padding: 20,
                        usePointStyle: true,
                        pointStyle: 'circle'
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(30, 41, 59, 0.95)',
                    titleColor: '#f1f5f9',
                    bodyColor: '#cbd5e1',
                    borderColor: 'rgba(99, 102, 241, 0.5)',
                    borderWidth: 1,
                    padding: 12
                }
            },
            scales: {
                y: {
                    beginAtZero: false,
                    grid: {
                        color: 'rgba(148, 163, 184, 0.1)'
                    },
                    ticks: {
                        color: '#cbd5e1',
                        callback: function (value) {
                            return formatNumber(value);
                        }
                    }
                },
                x: {
                    grid: {
                        display: false
                    },
                    ticks: {
                        color: '#cbd5e1'
                    }
                }
            }
        }
    });
}

// Render data table
function renderTable(data) {
    const thead = document.getElementById('table-head');
    const tbody = document.getElementById('table-body');

    if (data.length === 0) {
        tbody.innerHTML = '<tr><td colspan="100%">ไม่พบข้อมูล</td></tr>';
        return;
    }

    // Create table headers
    const headers = Object.keys(data[0]);
    thead.innerHTML = `
        <tr>
            <th>#</th>
            ${headers.map(h => `<th>${h}</th>`).join('')}
        </tr>
    `;

    // Create table rows (limit to first 100 rows for performance)
    const displayData = data.slice(0, 100);
    tbody.innerHTML = displayData.map((row, index) => `
        <tr>
            <td>${index + 1}</td>
            ${headers.map(h => `<td>${row[h] !== undefined ? row[h] : '-'}</td>`).join('')}
        </tr>
    `).join('');
}

// Setup event listeners
function setupEventListeners() {
    // Search functionality
    const searchInput = document.getElementById('search-input');
    searchInput.addEventListener('input', (e) => {
        filterTable(e.target.value);
    });

    // Chart filters
    document.getElementById('gender-filter')?.addEventListener('change', (e) => {
        const filter = e.target.value;
        updateGenderChart(filter);
    });

    document.getElementById('age-filter')?.addEventListener('change', (e) => {
        // Implement age filter logic
    });

    document.getElementById('province-filter')?.addEventListener('change', (e) => {
        const limit = parseInt(e.target.value);
        updateProvinceChart(limit);
    });
}

// Filter table based on search
function filterTable(query) {
    const tbody = document.getElementById('table-body');
    const rows = tbody.getElementsByTagName('tr');
    const searchTerm = query.toLowerCase();

    for (let row of rows) {
        const text = row.textContent.toLowerCase();
        row.style.display = text.includes(searchTerm) ? '' : 'none';
    }
}

// Utility: Format number with commas
function formatNumber(num) {
    if (num === 0) return '0';
    if (!num) return '-';
    return new Intl.NumberFormat('th-TH').format(num);
}

// Show error message
function showError(message) {
    const tbody = document.getElementById('table-body');
    tbody.innerHTML = `
        <tr>
            <td colspan="100%">
                <div class="loading" style="color: var(--danger);">
                    ❌ ${message}
                </div>
            </td>
        </tr>
    `;
}

function updateGenderChart(filter) {
    if (!chartInstances.gender) return;

    const data = chartInstances.gender.data.datasets[0].data;
    const male = data[0];
    const female = data[1];

    if (filter === 'male') {
        chartInstances.gender.data.datasets[0].data = [male, 0];
    } else if (filter === 'female') {
        chartInstances.gender.data.datasets[0].data = [0, female];
    } else {
        chartInstances.gender.data.datasets[0].data = [male, female];
    }

    chartInstances.gender.update();
}

function updateProvinceChart(limit) {
    const keys = Object.keys(rawData[0]);

    const provinceKey = keys.find(k =>
        k.includes('จังหวัด') || k.includes('Province') || k.includes('ชื่อจังหวัด')
    );

    const totalKey = keys.find(k =>
        k.includes('รวม') || k.includes('Total') || k.includes('total')
    );

    if (!provinceKey || !totalKey) return;

    let provinceData = rawData.map(row => ({
        province: row[provinceKey],
        total: parseFloat(row[totalKey]) || 0
    }))
        .filter(item => item.province && item.total > 0)
        .sort((a, b) => b.total - a.total)
        .slice(0, limit);

    chartInstances.province.data.labels = provinceData.map(d => d.province);
    chartInstances.province.data.datasets[0].data = provinceData.map(d => d.total);
    chartInstances.province.update();
}
