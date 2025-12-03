// 1. 讀取 Excel 檔案
function readExcelFile(file, mustInclude = ['員工工號']) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                let targetSheetData = null;

                for (const sheetName of workbook.SheetNames) {
                    const worksheet = workbook.Sheets[sheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    const isValidSheet = jsonData.slice(0, 20).some(row => 
                        row && mustInclude.every(keyword => row.some(cell => cell && cell.toString().includes(keyword)))
                    );

                    if (isValidSheet) {
                        targetSheetData = jsonData;
                        console.log(`Bingo! 在分頁 "${sheetName}" 找到資料！`);
                        break; 
                    }
                }

                if (!targetSheetData) { resolve(null); return; }
                resolve(targetSheetData);
            } catch (error) { reject(error); }
        };
        reader.readAsArrayBuffer(file);
    });
}

// 2. 處理請假資料
function processLeaveData(leaveData) {
    if (!leaveData || leaveData.length === 0) return {};

    let headerRowIndex = -1;
    let headers = [];

    for (let i = 0; i < Math.min(leaveData.length, 20); i++) {
        const row = leaveData[i];
        if (row && row.some(c => c && c.toString().includes('工號')) && row.some(c => c && c.toString().includes('合計'))) {
            headerRowIndex = i; headers = row; break;
        }
    }

    if (headerRowIndex === -1) { alert('請假表錯誤：找不到「員工工號」與「合計」欄位。'); return {}; }

    const dataRows = leaveData.slice(headerRowIndex + 1);
    const empIdIndex = headers.findIndex(c => c && c.toString().includes('工號'));
    const totalLeaveIndex = headers.findIndex(c => c && c.toString().includes('合計'));

    const leaveMap = {};
    dataRows.forEach(row => {
        if (row && row.length > empIdIndex) {
            const empId = row[empIdIndex];
            const totalHours = parseFloat(row[totalLeaveIndex]) || 0; 
            if (empId) { leaveMap[empId] = { total: totalHours }; }
        }
    });
    return leaveMap;
}

// 3. 處理員工名單 (判定考核類別 & 讀取項目類別)
function processEmployeeData(employeeData, leaveMap) {
    const processedData = [];
    if (!employeeData || employeeData.length === 0) return processedData;

    let headerRowIndex = -1;
    let headers = [];

    for (let i = 0; i < Math.min(employeeData.length, 20); i++) {
        const row = employeeData[i];
        if (row && row.some(c => c && c.toString().includes('工號')) && row.some(c => c && c.toString().includes('姓名'))) {
            headerRowIndex = i; headers = row; break;
        }
    }

    if (headerRowIndex === -1) { alert('員工名單錯誤：找不到標題列。'); return processedData; }

    const dataRows = employeeData.slice(headerRowIndex + 1);
    const findIdx = (keywords) => headers.findIndex(h => h && keywords.some(k => h.toString().includes(k)));

    const idx = {
        id: findIdx(['工號']),
        name: findIdx(['中文姓名', '員工姓名']),
        dept: findIdx(['部門']),
        hire: findIdx(['到職']),
        title: findIdx(['職務', '職稱']),
        level: findIdx(['類組', '分類', '職等']), 
        leaveStart: findIdx(['留職停薪']),
        leaveEnd: findIdx(['留停復職']),
        itemCategory: findIdx(['項目考核類別', '項目', '考核項目'])
    };

    if (idx.name === -1) idx.name = findIdx(['姓名']);

    // 強制指定 M 欄邏輯
    if (idx.itemCategory === -1 && headers.length > 12) {
        idx.itemCategory = 12; 
        console.log("未找到項目考核類別標題，自動指定為 M 欄 (Index 12)");
    }

    const endOfYear = new Date(Date.UTC(2024, 11, 31)); 
    const startOfYear = new Date(Date.UTC(2024, 0, 1));   

    dataRows.forEach(row => {
        if (!row || idx.id === -1 || !row[idx.id]) return;

        const empId = row[idx.id];
        const hireDate = idx.hire !== -1 ? parseExcelDate(row[idx.hire]) : null;
        const formattedHireDate = hireDate ? hireDate.toISOString().split('T')[0] : '';
        
        const leaveInfo = leaveMap[empId] || { total: 0 };
        const totalLeaveHours = leaveInfo.total;

        let level = (idx.level !== -1 && row[idx.level]) ? row[idx.level].toString().trim() : '';
        let jobTitle = (idx.title !== -1 && row[idx.title]) ? row[idx.title].toString().trim() : '';
        
        let itemCategory = '';
        if (idx.itemCategory !== -1 && row[idx.itemCategory] !== undefined) {
            itemCategory = row[idx.itemCategory]; 
        }

        let appraisalType = '特考'; 
        
        if (hireDate) {
            let lwopDays = 0;
            if (idx.leaveStart !== -1 && row[idx.leaveStart]) {
                const lwopStart = parseExcelDate(row[idx.leaveStart]);
                const lwopEnd = (idx.leaveEnd !== -1 && row[idx.leaveEnd]) ? parseExcelDate(row[idx.leaveEnd]) : endOfYear;

                if (lwopStart && lwopEnd) {
                    const effStart = lwopStart < startOfYear ? startOfYear : lwopStart;
                    const effEnd = lwopEnd > endOfYear ? endOfYear : lwopEnd;
                    if (effEnd >= effStart) {
                        lwopDays = (effEnd - effStart) / (1000 * 60 * 60 * 24) + 1;
                    }
                }
            }

            if (hireDate < startOfYear && lwopDays === 0) {
                appraisalType = '年考';
            } else {
                const baseDate = hireDate < startOfYear ? startOfYear : hireDate;
                const potentialDays = (endOfYear - baseDate) / (1000 * 60 * 60 * 24) + 1;
                const actualServiceDays = potentialDays - lwopDays;

                if (actualServiceDays >= 183) {
                    appraisalType = '另考';
                } else {
                    appraisalType = '特考';
                }
            }
        }

        processedData.push({
            employeeId: empId,
            name: idx.name !== -1 ? row[idx.name] : '',
            department: idx.dept !== -1 ? row[idx.dept] : '',
            jobTitle: jobTitle,
            hireDate: formattedHireDate,
            totalLeave: totalLeaveHours,
            appraisalType: appraisalType,
            level: level,
            itemCategory: itemCategory
        });
    });
    return processedData;
}

function parseExcelDate(value) {
    if (!value) return null;
    if (typeof value === 'number') { return new Date(Date.UTC(1899, 11, 30 + value)); }
    if (typeof value === 'string') { const date = new Date(value); if (!isNaN(date.getTime())) return date; }
    return null;
}

// 4. 項目考核類別員額統計
function calculateAndRenderAllocation(data) {
    const categoryStats = {};

    // 1~5 類別
    const TARGET_CATEGORIES = ['1', '2', '3', '4', '5'];
    
    // 初始化
    TARGET_CATEGORIES.forEach(cat => {
        categoryStats[cat] = { 
            A: 0, a: 0, 
            B: 0, b: 0, 
            C: 0, c: 0, 
            total: 0 
        };
    });

    // 統計人數
    data.forEach(emp => {
        const type = emp.appraisalType;
        const leave = parseFloat(emp.totalLeave) || 0;
        const cat = String(emp.itemCategory).trim();

        if (TARGET_CATEGORIES.includes(cat)) {
            const stats = categoryStats[cat];
            
            if (type === '年考') {
                if (leave === 0) stats.A++;
                else stats.a++;
            } else if (type === '另考') {
                if (leave === 0) stats.B++;
                else stats.b++;
            } else if (type === '特考') {
                if (leave === 0) stats.C++;
                else stats.c++;
            }
            stats.total++;
        }
    });

    // 計算配額 (F, H, I, J)
    const tableData = Object.keys(categoryStats).map(catName => {
        const s = categoryStats[catName];
        
        // D: 受考人數
        const D = s.total;
        
        // d: 事病假
        const d = s.a + s.b + s.c;
        
        // E: 2等基數
        const E = D - d;
        
        // F: 2等員額
        const F = Math.floor(E * 0.25);
        
        // G: 3等基數
        const G = E - F;
        
        // H: 3等員額
        const H = Math.floor(G * 0.60);
        
        // I: 4等員額
        const I = Math.ceil((G - H) + (d / 2.0));
        
        // J: 5等員額
        let J = D - F - H - I;
        if (J < 0) J = 0;

        return { catName, D, d, E, F, G, H, I, J };
    });

    // 渲染新表格 (只保留 2, 3, 4, 5 等員額)
    if ($.fn.DataTable.isDataTable('#allocationTable')) {
        $('#allocationTable').DataTable().destroy();
    }

    $('#allocationTable').DataTable({
        data: tableData,
        columns: [
            { title: "類別", data: "catName", className: "fw-bold text-center" },
            { 
                title: "2等員額", data: "F", className: "text-center fw-bold",
                render: data => data + ' 人'
            },
            { 
                title: "3等員額", data: "H", className: "text-center fw-bold",
                render: data => data + ' 人'
            },
            { 
                title: "4等員額", data: "I", className: "text-center fw-bold",
                render: data => data + ' 人'
            },
            { 
                title: "5等員額", data: "J", className: "text-center fw-bold",
                render: data => data + ' 人'
            }
        ],
        language: { "url": "https://cdn.datatables.net/plug-ins/1.11.5/i18n/zh-HANT.json" },
        dom: 'Bfrtip',
        order: [[0, 'asc']],
        paging: false,
        searching: false,
        info: false 
    });
}