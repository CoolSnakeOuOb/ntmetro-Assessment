// 1. 讀取 Excel 檔案 (超級聰明版：支援指定必要關鍵字)
// mustInclude: 一個陣列，指定該分頁必須包含哪些字，例如 ['員工工號', '合計']
function readExcelFile(file, mustInclude = ['員工工號']) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                let targetSheetData = null;
                let foundSheetName = "";

                // 迴圈：一個一個分頁檢查
                for (const sheetName of workbook.SheetNames) {
                    const worksheet = workbook.Sheets[sheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    // 檢查前 20 列
                    // 條件：這一列必須 "同時" 包含所有 mustInclude 裡的關鍵字
                    const isValidSheet = jsonData.slice(0, 20).some(row => 
                        row && mustInclude.every(keyword => row.includes(keyword))
                    );

                    if (isValidSheet) {
                        targetSheetData = jsonData;
                        foundSheetName = sheetName;
                        console.log(`Bingo! 在分頁 "${sheetName}" 找到包含 ${mustInclude} 的資料了！`);
                        break; 
                    }
                }

                // 如果找不到，回傳 null 讓後面報錯，或者回傳第一個分頁
                if (!targetSheetData) {
                    console.warn(`在所有分頁都找不到同時包含 ${mustInclude} 的資料。`);
                    // 這裡我們不回傳預設值了，直接回傳 null，讓外面知道找不到對的分頁
                    resolve(null); 
                    return;
                }

                resolve(targetSheetData);

            } catch (error) {
                reject(error);
            }
        };
        reader.readAsArrayBuffer(file);
    });
}

// 2. 處理請假資料 (自動搜尋標題列)
function processLeaveData(leaveData) {
    if (!leaveData || leaveData.length === 0) {
        alert('請假明細檔案內容為空！');
        return {};
    }

    // --- 自動搜尋標題列邏輯 ---
    let headerRowIndex = -1;
    let headers = [];

    // 搜尋前 20 列，看哪一列同時包含 "員工工號" 和 "合計"
    for (let i = 0; i < Math.min(leaveData.length, 20); i++) {
        const row = leaveData[i];
        if (row && row.includes('員工工號') && row.includes('合計')) {
            headerRowIndex = i;
            headers = row;
            break;
        }
    }

    if (headerRowIndex === -1) {
        console.error("目前讀取到的資料前5筆:", leaveData.slice(0, 5));
        alert(`請假表錯誤：在前 20 列中找不到包含「員工工號」和「合計」的標題列。\n請確認該分頁是否正確。`);
        return {};
    }
    // -------------------------

    const dataRows = leaveData.slice(headerRowIndex + 1);
    const empIdIndex = headers.indexOf('員工工號');
    const totalLeaveIndex = headers.indexOf('合計');

    const leaveMap = {};
    dataRows.forEach(row => {
        if (row && row.length > empIdIndex) {
            const empId = row[empIdIndex];
            // 轉成數字，無資料則為0
            const total = parseFloat(row[totalLeaveIndex]) || 0; 
            
            if (empId) {
                leaveMap[empId] = { total: total };
            }
        }
    });
    
    return leaveMap;
}

// 3. 處理員工名單並進行篩選 (自動搜尋標題列)
function processEmployeeData(employeeData, leaveMap) {
    const processedData = [];
    if (!employeeData || employeeData.length === 0) return processedData;

    // --- 自動搜尋標題列邏輯 ---
    let headerRowIndex = -1;
    let headers = [];

    // 搜尋前 20 列
    for (let i = 0; i < Math.min(employeeData.length, 20); i++) {
        const row = employeeData[i];
        if (row && (row.includes('員工工號') || row.includes('工號')) && (row.includes('中文姓名') || row.includes('姓名'))) {
            headerRowIndex = i;
            headers = row;
            break;
        }
    }

    if (headerRowIndex === -1) {
        alert(`員工名單錯誤：在前 20 列中找不到包含「員工工號」和「中文姓名」的標題列。`);
        return processedData;
    }
    // -------------------------

    const dataRows = employeeData.slice(headerRowIndex + 1);
    
    // 欄位對應 (彈性處理：如果找不到標準名稱，會變成 -1，程式後面會判斷)
    const idx = {
        id: headers.indexOf('員工工號') !== -1 ? headers.indexOf('員工工號') : headers.indexOf('工號'),
        name: headers.indexOf('中文姓名') !== -1 ? headers.indexOf('中文姓名') : headers.indexOf('姓名'),
        dept: headers.indexOf('部門名稱'),
        hire: headers.indexOf('到職日期'),
        title: headers.indexOf('職務名稱'),
        leaveStart: headers.indexOf('留職停薪日'),
        leaveEnd: headers.indexOf('留停復職日')
    };

    if (idx.id === -1) {
        alert(`員工名單錯誤：找不到「員工工號」欄位。`);
        return processedData;
    }

    const endOfYear = new Date(Date.UTC(2024, 11, 31)); 
    const startOfYear = new Date(Date.UTC(2024, 0, 1));   

    dataRows.forEach(row => {
        if (!row || !row[idx.id]) return;

        const empId = row[idx.id];
        
        // 處理日期
        const hireDate = idx.hire !== -1 ? parseExcelDate(row[idx.hire]) : null;
        const formattedHireDate = hireDate ? hireDate.toISOString().split('T')[0] : '';
        
        // 取得請假天數
        const leaveInfo = leaveMap[empId] || { total: 0 };
        const totalLeaveDays = leaveInfo.total;

        // 判斷考核類別邏輯
        let appraisalType = '不予考核';
        
        if (totalLeaveDays >= 180) {
            appraisalType = '不予考績'; 
        } 
        else if (hireDate) {
            const probationEndDate = new Date(hireDate);
            probationEndDate.setMonth(probationEndDate.getMonth() + 6);
            const actualStart = probationEndDate > startOfYear ? probationEndDate : startOfYear;
            const daysInPeriod = (endOfYear - actualStart) / (1000 * 60 * 60 * 24);
            const actualWorkDays = daysInPeriod - totalLeaveDays;

            if (actualWorkDays >= 365) appraisalType = '年考';
            else if (actualWorkDays >= 180) appraisalType = '另考';
            else appraisalType = '不予考核'; 
        }

        processedData.push({
            employeeId: empId,
            name: idx.name !== -1 ? row[idx.name] : '',
            department: idx.dept !== -1 ? row[idx.dept] : '',
            jobTitle: idx.title !== -1 ? row[idx.title] : '',
            hireDate: formattedHireDate,
            totalLeave: totalLeaveDays,
            appraisalType: appraisalType
        });
    });
    return processedData;
}

// 輔助：解析 Excel 日期
function parseExcelDate(value) {
    if (!value) return null;
    if (typeof value === 'number') {
        return new Date(Date.UTC(1899, 11, 30 + value));
    }
    if (typeof value === 'string') {
        const date = new Date(value);
        if (!isNaN(date.getTime())) return date;
    }
    return null;
}