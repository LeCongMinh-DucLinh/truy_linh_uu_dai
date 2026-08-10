<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Công Cụ Tính Truy Lĩnh Phụ Cấp Ưu Đãi</title>
    <!-- Tailwind CSS for styling -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Font Awesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- SheetJS for Excel export -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        @media print {
            .no-print { display: none !important; }
            body { font-size: 12pt; background: white; }
            .container { max-width: 100% !important; margin: 0 !important; padding: 0 !important; }
            table { width: 100% !important; border-collapse: collapse; }
            th, td { border: 1px solid #000 !important; padding: 6px !important; }
        }
    </style>
</head>
<body class="bg-slate-50 text-slate-800 min-h-screen font-sans">

    <!-- Header Banner -->
    <header class="bg-indigo-700 text-white shadow-lg no-print">
        <div class="max-w-7xl mx-auto px-4 py-6 flex flex-col md:flex-row justify-between items-center gap-4">
            <div>
                <h1 class="text-2xl font-bold flex items-center gap-2">
                    <i class="fa-solid font-calculator"></i> Công Cụ Tính Truy Lĩnh Phụ Cấp Ưu Đãi
                </h1>
                <p class="text-indigo-200 text-sm mt-1">Tính toán nhanh chóng, chính xác khoản chênh lệch phụ cấp ưu đãi từ ngày 01/01/2026</p>
            </div>
            <div class="flex gap-2">
                <button onclick="exportToExcel()" class="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg font-medium shadow transition flex items-center gap-2">
                    <i class="fa-solid fa-file-excel"></i> Xuất Excel
                </button>
                <button onclick="window.print()" class="bg-slate-700 hover:bg-slate-800 text-white px-4 py-2 rounded-lg font-medium shadow transition flex items-center gap-2">
                    <i class="fa-solid fa-print"></i> In / Xuất PDF
                </button>
            </div>
        </div>
    </header>

    <!-- Main Container -->
    <main class="max-w-7xl mx-auto px-4 py-8">

        <!-- User Info & Presets -->
        <div class="bg-white p-6 rounded-xl shadow-sm border border-slate-200 mb-6 no-print">
            <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                    <label class="block text-xs font-semibold uppercase tracking-wider text-slate-500 mb-1">Họ và tên cán bộ/giáo viên</label>
                    <input type="text" id="employeeName" value="Lê Công Minh" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:outline-none font-medium text-slate-800" oninput="updateNameHeader()">
                </div>
                <div>
                    <label class="block text-xs font-semibold uppercase tracking-wider text-slate-500 mb-1">Mức phụ cấp ưu đãi cũ (%)</label>
                    <input type="number" id="defaultOldPct" value="30" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:outline-none">
                </div>
                <div>
                    <label class="block text-xs font-semibold uppercase tracking-wider text-slate-500 mb-1">Mức phụ cấp ưu đãi mới (%)</label>
                    <input type="number" id="defaultNewPct" value="60" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:outline-none">
                </div>
            </div>
        </div>

        <!-- Print Header Title -->
        <div class="hidden print:block text-center mb-6">
            <h1 class="text-xl font-bold uppercase">BẢNG TÍNH TRUY LĨNH PHỤ CẤP ƯU ĐÃI</h1>
            <p class="text-sm italic">Áp dụng chênh lệch mức phụ cấp từ ngày 01/01/2026</p>
            <p class="text-sm font-semibold mt-2">Họ và tên: <span id="printName">Lê Công Minh</span></p>
        </div>

        <!-- Interactive Data Table -->
        <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-x-auto">
            <table class="w-full text-left border-collapse text-sm" id="calcTable">
                <thead>
                    <tr class="bg-slate-100 text-slate-700 font-semibold border-b border-slate-200">
                        <th class="p-3 text-center w-12">#</th>
                        <th class="p-3">Từ tháng/năm</th>
                        <th class="p-3">Đến tháng/năm</th>
                        <th class="p-3 text-right">HS Lương</th>
                        <th class="p-3 text-right">HS Chức vụ</th>
                        <th class="p-3 text-right">Mức cũ (%)</th>
                        <th class="p-3 text-right">Mức mới (%)</th>
                        <th class="p-3 text-right">Chênh lệch (%)</th>
                        <th class="p-3 text-center">Số tháng</th>
                        <th class="p-3 text-right">Lương cơ sở (đ)</th>
                        <th class="p-3 text-right">Số tiền chênh lệch (đ)</th>
                        <th class="p-3 text-center no-print w-16">Thao tác</th>
                    </tr>
                </thead>
                <tbody id="tableBody" class="divide-y divide-slate-200">
                    <!-- Dynamic Rows Rendered Here -->
                </tbody>
                <tfoot>
                    <tr class="bg-indigo-50 font-bold text-slate-900 border-t-2 border-indigo-200">
                        <td colspan="8" class="p-4 text-right uppercase tracking-wider text-xs text-indigo-900">Tổng cộng tiền truy lĩnh nhận được:</td>
                        <td id="totalMonths" class="p-4 text-center text-indigo-900">0</td>
                        <td></td>
                        <td id="grandTotal" class="p-4 text-right text-indigo-700 text-base">0 đ</td>
                        <td class="no-print"></td>
                    </tr>
                </tfoot>
            </table>
        </div>

        <!-- Add Row Button -->
        <div class="mt-4 flex justify-between items-center no-print">
            <button onclick="addRow()" class="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg font-medium shadow transition flex items-center gap-2">
                <i class="fa-solid fa-plus"></i> Thêm giai đoạn tính
            </button>
            <button onclick="resetDefault()" class="text-slate-500 hover:text-slate-700 text-sm font-medium underline">
                Đặt lại mẫu ban đầu
            </button>
        </div>

        <!-- Formula Explanation & Formula Cards -->
        <div class="mt-8 grid grid-cols-1 md:grid-cols-2 gap-6 no-print">
            <div class="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
                <h3 class="font-bold text-slate-800 mb-2 flex items-center gap-2">
                    <i class="fa-solid fa-circle-info text-indigo-600"></i> Công thức tính chi tiết
                </h3>
                <ul class="text-xs text-slate-600 space-y-2 list-disc pl-4">
                    <li><b>% Chênh lệch:</b> = Mức mới (%) - Mức cũ (%)</li>
                    <li><b>Số tháng hưởng:</b> Tính tổng số tháng bao gồm cả tháng bắt đầu và tháng kết thúc.</li>
                    <li><b>Số tiền chênh lệch:</b> = % Chênh lệch × (Hệ số lương + HS Chức vụ) × Lương cơ sở × Số tháng.</li>
                </ul>
            </div>
            <div class="bg-amber-50 p-5 rounded-xl border border-amber-200 text-amber-900 text-xs">
                <h3 class="font-bold mb-2 flex items-center gap-2 text-amber-800">
                    <i class="fa-solid fa-lightbulb"></i> Hướng dẫn chia sẻ nhanh
                </h3>
                <p class="mb-2">Bạn có thể dễ dàng tải file HTML này lên <b>GitHub Pages</b> hoặc <b>Firebase Hosting</b> hoàn toàn miễn phí để tạo đường dẫn link chia sẻ cho đồng nghiệp.</p>
                <p>Xem file hướng dẫn triển khai chi tiết được đính kèm để xem từng bước thực hiện.</p>
            </div>
        </div>
    </main>

    <script>
        // Sample data loaded from Excel
        let rowsData = [
            { startMonth: "2026-01", endMonth: "2026-05", hsl: 5.7, hscv: 0.45, oldPct: 30, newPct: 60, baseSalary: 2340000 },
            { startMonth: "2026-06", endMonth: "2026-06", hsl: 6.04, hscv: 0.45, oldPct: 30, newPct: 60, baseSalary: 2340000 },
            { startMonth: "2026-07", endMonth: "2026-08", hsl: 6.04, hscv: 0.45, oldPct: 30, newPct: 60, baseSalary: 2530000 }
        ];

        function formatCurrency(val) {
            return new Intl.NumberFormat('vi-VN').format(Math.round(val));
        }

        function calculateMonths(startStr, endStr) {
            if (!startStr || !endStr) return 0;
            const [startY, startM] = startStr.split('-').map(Number);
            const [endY, endM] = endStr.split('-').map(Number);
            
            const totalStart = startY * 12 + startM;
            const totalEnd = endY * 12 + endM;
            const diff = totalEnd - totalStart + 1;
            return diff > 0 ? diff : 0;
        }

        function updateNameHeader() {
            const name = document.getElementById('employeeName').value;
            document.getElementById('printName').innerText = name || '...';
        }

        function renderTable() {
            const tbody = document.getElementById('tableBody');
            tbody.innerHTML = '';
            
            let grandTotal = 0;
            let totalMonthsSum = 0;

            rowsData.forEach((row, index) => {
                const diffPct = (row.newPct || 0) - (row.oldPct || 0);
                const months = calculateMonths(row.startMonth, row.endMonth);
                const coeffSum = (parseFloat(row.hsl) || 0) + (parseFloat(row.hscv) || 0);
                const rowAmount = (diffPct / 100) * coeffSum * (parseFloat(row.baseSalary) || 0) * months;

                grandTotal += rowAmount;
                totalMonthsSum += months;

                const tr = document.createElement('tr');
                tr.className = "hover:bg-slate-50 transition";

                tr.innerHTML = `
                    <td class="p-3 text-center text-slate-400 font-medium">${index + 1}</td>
                    <td class="p-3">
                        <input type="month" value="${row.startMonth}" onchange="updateRow(${index}, 'startMonth', this.value)" class="w-full p-1 border border-slate-300 rounded focus:ring-1 focus:ring-indigo-500 no-print">
                        <span class="hidden print:inline">${row.startMonth}</span>
                    </td>
                    <td class="p-3">
                        <input type="month" value="${row.endMonth}" onchange="updateRow(${index}, 'endMonth', this.value)" class="w-full p-1 border border-slate-300 rounded focus:ring-1 focus:ring-indigo-500 no-print">
                        <span class="hidden print:inline">${row.endMonth}</span>
                    </td>
                    <td class="p-3 text-right">
                        <input type="number" step="0.01" value="${row.hsl}" onchange="updateRow(${index}, 'hsl', this.value)" class="w-20 p-1 border border-slate-300 rounded text-right focus:ring-1 focus:ring-indigo-500 no-print">
                        <span class="hidden print:inline">${row.hsl}</span>
                    </td>
                    <td class="p-3 text-right">
                        <input type="number" step="0.01" value="${row.hscv}" onchange="updateRow(${index}, 'hscv', this.value)" class="w-20 p-1 border border-slate-300 rounded text-right focus:ring-1 focus:ring-indigo-500 no-print">
                        <span class="hidden print:inline">${row.hscv}</span>
                    </td>
                    <td class="p-3 text-right">
                        <input type="number" value="${row.oldPct}" onchange="updateRow(${index}, 'oldPct', this.value)" class="w-16 p-1 border border-slate-300 rounded text-right focus:ring-1 focus:ring-indigo-500 no-print">
                        <span class="hidden print:inline">${row.oldPct}%</span>
                    </td>
                    <td class="p-3 text-right">
                        <input type="number" value="${row.newPct}" onchange="updateRow(${index}, 'newPct', this.value)" class="w-16 p-1 border border-slate-300 rounded text-right focus:ring-1 focus:ring-indigo-500 no-print">
                        <span class="hidden print:inline">${row.newPct}%</span>
                    </td>
                    <td class="p-3 text-right font-medium text-slate-700">${diffPct}%</td>
                    <td class="p-3 text-center font-semibold text-indigo-600">${months}</td>
                    <td class="p-3 text-right">
                        <input type="number" value="${row.baseSalary}" onchange="updateRow(${index}, 'baseSalary', this.value)" class="w-28 p-1 border border-slate-300 rounded text-right focus:ring-1 focus:ring-indigo-500 no-print">
                        <span class="hidden print:inline">${formatCurrency(row.baseSalary)} đ</span>
                    </td>
                    <td class="p-3 text-right font-bold text-indigo-700">${formatCurrency(rowAmount)} đ</td>
                    <td class="p-3 text-center no-print">
                        <button onclick="deleteRow(${index})" class="text-rose-500 hover:text-rose-700 px-2 py-1 transition" title="Xóa hàng">
                            <i class="fa-solid fa-trash-can"></i>
                        </button>
                    </td>
                `;
                tbody.appendChild(tr);
            });

            document.getElementById('totalMonths').innerText = totalMonthsSum;
            document.getElementById('grandTotal').innerText = formatCurrency(grandTotal) + " đ";
        }

        function updateRow(index, field, value) {
            if (field === 'hsl' || field === 'hscv' || field === 'oldPct' || field === 'newPct' || field === 'baseSalary') {
                rowsData[index][field] = parseFloat(value) || 0;
            } else {
                rowsData[index][field] = value;
            }
            renderTable();
        }

        function addRow() {
            const oldPct = parseFloat(document.getElementById('defaultOldPct').value) || 30;
            const newPct = parseFloat(document.getElementById('defaultNewPct').value) || 60;
            const lastRow = rowsData[rowsData.length - 1];
            
            let startMonth = "2026-09";
            let endMonth = "2026-12";
            let hsl = 6.04;
            let hscv = 0.45;
            let baseSalary = 2530000;

            if (lastRow) {
                hsl = lastRow.hsl;
                hscv = lastRow.hscv;
                baseSalary = lastRow.baseSalary;
            }

            rowsData.push({ startMonth, endMonth, hsl, hscv, oldPct, newPct, baseSalary });
            renderTable();
        }

        function deleteRow(index) {
            if (rowsData.length <= 1) {
                alert("Bảng phải có ít nhất 1 dòng dữ liệu.");
                return;
            }
            rowsData.splice(index, 1);
            renderTable();
        }

        function resetDefault() {
            rowsData = [
                { startMonth: "2026-01", endMonth: "2026-05", hsl: 5.7, hscv: 0.45, oldPct: 30, newPct: 60, baseSalary: 2340000 },
                { startMonth: "2026-06", endMonth: "2026-06", hsl: 6.04, hscv: 0.45, oldPct: 30, newPct: 60, baseSalary: 2340000 },
                { startMonth: "2026-07", endMonth: "2026-08", hsl: 6.04, hscv: 0.45, oldPct: 30, newPct: 60, baseSalary: 2530000 }
            ];
            document.getElementById('employeeName').value = "Lê Công Minh";
            updateNameHeader();
            renderTable();
        }

        function exportToExcel() {
            const name = document.getElementById('employeeName').value || "CanBo";
            const exportData = rowsData.map((r, i) => {
                const diffPct = r.newPct - r.oldPct;
                const months = calculateMonths(r.startMonth, r.endMonth);
                const rowAmount = (diffPct / 100) * (r.hsl + r.hscv) * r.baseSalary * months;
                return {
                    "STT": i + 1,
                    "Từ tháng": r.startMonth,
                    "Đến tháng": r.endMonth,
                    "Hệ số lương": r.hsl,
                    "HS Chức vụ": r.hscv,
                    "Mức cũ (%)": r.oldPct,
                    "Mức mới (%)": r.newPct,
                    "Chênh lệch (%)": diffPct,
                    "Số tháng": months,
                    "Lương cơ sở (đ)": r.baseSalary,
                    "Số tiền chênh lệch (đ)": rowAmount
                };
            });

            const ws = XLSX.utils.json_to_sheet(exportData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "TruyLinhPhuCap");
            XLSX.writeFile(wb, `TruyLinhPhuCap_${name}.xlsx`);
        }

        // Initial rendering
        renderTable();
    </script>
</body>
</html>
