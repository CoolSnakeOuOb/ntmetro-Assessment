// 1. 讀取 Excel 檔案 (聰明搜尋版：自動遍歷所有分頁，直到找到關鍵字)
function readExcelFile(file, mustInclude = ['員工工號']) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                let targetSheetData = null;

                // 迴圈：一個一個分頁檢查
                for (const sheetName of workbook.SheetNames) {
                    const worksheet = workbook.Sheets[sheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    const isValidSheet = jsonData.slice(0, 20).some(row => 
                        row && mustInclude.every(keyword => row.some(cell => cell && cell.toString().includes(keyword)))
                    );

                    if (isValidSheet) {
                        targetSheetData = jsonData;
                        console.log(`Bingo! 在分頁 "${sheetName}" 找到目標資料！`);
                        break; 
                    }
                }

                if (!targetSheetData) {
                    console.warn(`找不到包含 ${mustInclude} 的資料分頁。`);
                    resolve(null); 
                    return;
                }
                resolve(targetSheetData);
            } catch (error) { reject(error); }
        };
        reader.readAsArrayBuffer(file);
    });
}

// 2. 處理請假資料 (只負責讀取時數，不做判斷)
function processLeaveData(leaveData) {
    if (!leaveData || leaveData.length === 0) return {};

    let headerRowIndex = -1;
    let headers = [];

    // 自動搜尋標題：找 "工號" 和 "合計"
    for (let i = 0; i < Math.min(leaveData.length, 20); i++) {
        const row = leaveData[i];
        if (row && 
            row.some(c => c && c.toString().includes('工號')) && 
            row.some(c => c && c.toString().includes('合計'))) {
            headerRowIndex = i;
            headers = row;
            break;
        }
    }

    if (headerRowIndex === -1) {
        alert('請假表錯誤：找不到「員工工號」與「合計」欄位。');
        return {};
    }

    const dataRows = leaveData.slice(headerRowIndex + 1);
    const empIdIndex = headers.findIndex(c => c && c.toString().includes('工號'));
    const totalLeaveIndex = headers.findIndex(c => c && c.toString().includes('合計'));

    const leaveMap = {};
    dataRows.forEach(row => {
        if (row && row.length > empIdIndex) {
            const empId = row[empIdIndex];
            const totalHours = parseFloat(row[totalLeaveIndex]) || 0; 
            if (empId) {
                leaveMap[empId] = { total: totalHours };
            }
        }
    });
    return leaveMap;
}

// 3. 處理員工名單 (核心修正：只扣除留職停薪，不扣事病假)
function processEmployeeData(employeeData, leaveMap) {
    const processedData = [];
    if (!employeeData || employeeData.length === 0) return processedData;

    let headerRowIndex = -1;
    let headers = [];

    // 自動搜尋標題
    for (let i = 0; i < Math.min(employeeData.length, 20); i++) {
        const row = employeeData[i];
        if (row && 
            row.some(c => c && c.toString().includes('工號')) && 
            row.some(c => c && c.toString().includes('姓名'))) {
            headerRowIndex = i;
            headers = row;
            break;
        }
    }

    if (headerRowIndex === -1) {
        alert('員工名單錯誤：找不到標題列 (需包含「工號」與「姓名」)。');
        return processedData;
    }

    const dataRows = employeeData.slice(headerRowIndex + 1);
    
    // 欄位定位
    const findIdx = (keywords) => headers.findIndex(h => h && keywords.some(k => h.toString().includes(k)));

    const idx = {
        id: findIdx(['工號']),
        name: findIdx(['姓名']),
        dept: findIdx(['部門']),
        hire: findIdx(['到職']),
        title: findIdx(['職務', '職稱']),
        level: findIdx(['類組', '分類', '職等']), 
        leaveStart: findIdx(['留職停薪']), // 留停開始
        leaveEnd: findIdx(['留停復職'])   // 留停結束
    };

    // 設定考核年度 (假設為 2024)
    const startOfYear = new Date(Date.UTC(2024, 0, 1)); // 2024-01-01
    const endOfYear = new Date(Date.UTC(2024, 11, 31)); // 2024-12-31
    const daysInYear = (endOfYear - startOfYear) / (1000 * 60 * 60 * 24) + 1; // 366天

    dataRows.forEach(row => {
        if (!row || idx.id === -1 || !row[idx.id]) return;

        const empId = row[idx.id];
        const hireDate = idx.hire !== -1 ? parseExcelDate(row[idx.hire]) : null;
        const formattedHireDate = hireDate ? hireDate.toISOString().split('T')[0] : '';
        
        // 請假時數 (僅顯示用，不影響考核)
        const leaveInfo = leaveMap[empId] || { total: 0 };
        const totalLeaveHours = leaveInfo.total;

        // 讀取其他資訊
        let level = (idx.level !== -1 && row[idx.level]) ? row[idx.level].toString().trim() : '';
        let jobTitle = (idx.title !== -1 && row[idx.title]) ? row[idx.title].toString().trim() : '';

        // --- 核心邏輯修正區 ---
        let appraisalType = '不予考核';
        
        if (hireDate) {
            // 1. 計算留職停薪天數 (LWOP)
            let lwopDays = 0;
            if (idx.leaveStart !== -1 && row[idx.leaveStart]) {
                const lwopStart = parseExcelDate(row[idx.leaveStart]);
                // 如果沒有復職日，假設請到年底 (還在留停中)
                const lwopEnd = (idx.leaveEnd !== -1 && row[idx.leaveEnd]) ? parseExcelDate(row[idx.leaveEnd]) : endOfYear;

                if (lwopStart && lwopEnd) {
                    // 計算留停期間與「今年」的重疊天數
                    const effStart = lwopStart < startOfYear ? startOfYear : lwopStart;
                    const effEnd = lwopEnd > endOfYear ? endOfYear : lwopEnd;
                    
                    if (effEnd >= effStart) {
                        lwopDays = (effEnd - effStart) / (1000 * 60 * 60 * 24) + 1;
                    }
                }
            }

            // 2. 判斷考核類別
            // 規則：
            // - 到職日在今年 1/1 以前 且 沒有留職停薪 => 年考
            // - 否則計算「實際在職天數 (扣除留停)」
            //   - 滿 6 個月 (約 183 天) => 另考
            //   - 不滿 6 個月 => 不予考核

            if (hireDate < startOfYear && lwopDays === 0) {
                appraisalType = '年考';
            } else {
                // 計算基礎：如果是舊員工，從年初算；新員工從到職日算
                const baseDate = hireDate < startOfYear ? startOfYear : hireDate;
                // 潛在最大在職天數
                const potentialDays = (endOfYear - baseDate) / (1000 * 60 * 60 * 24) + 1;
                // 實際在職 = 潛在 - 留停
                const actualServiceDays = potentialDays - lwopDays;

                if (actualServiceDays >= 183) { // 滿半年
                    appraisalType = '另考';
                } else {
                    appraisalType = '不予考核';
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
            level: level
        });
    });
    return processedData;
}

// 日期解析輔助函式
function parseExcelDate(value) {
    if (!value) return null;
    if (typeof value === 'number') { return new Date(Date.UTC(1899, 11, 30 + value)); }
    if (typeof value === 'string') { const date = new Date(value); if (!isNaN(date.getTime())) return date; }
    return null;
}