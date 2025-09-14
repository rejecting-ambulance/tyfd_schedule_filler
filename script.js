document.addEventListener('DOMContentLoaded', () => {
    // 取得所有 DOM 元素
    const scheduleTable = document.getElementById("schedule");
    const thead = document.querySelector("#schedule thead");
    const tbody = document.querySelector("#schedule tbody");
    const personGrid = document.getElementById("personGrid");
    const templateSelect = document.getElementById("templateSelect");
    const assignPersonsBtn = document.getElementById("assignPersons");
    const downloadExcelBtn = document.getElementById("downloadExcelBtn");
    const modeSwitch = document.getElementById("modeSwitch"); 

    // Google Apps Script API 網址
    const API_URL = "https://script.google.com/macros/s/AKfycbx1Hmh7WpqrEpPtahVY7y2-wZnd5KChxuzmlA4o7Ld02SugyAOZabmLDLKPlQVuTt-O/exec";

    // 獨特的拖曳資料類型
    const DRAG_TYPE = "application/x-person-card";
    
    // 用於追蹤正在被拖曳的卡片或人員元素
    let draggedElement = null;

    let templateData = {};
    let currentTemplate = null;
    let templateLetters = []; 
    const fixedOptions = ["輪休", "補休", "公假"]; 

    // --- 1. 初始化人員清單 (表格初始化現在由模板載入處理) ---
    // 生成人員清單 (30人)
    let totalPeople = 30,
        colCount = 3,
        rowCount = 10;
    for (let row = 0; row < rowCount; row++) {
        for (let col = 0; col < colCount; col++) {
            let idx = row + col * rowCount + 1;
            if (idx <= totalPeople) {
                let div = document.createElement("div");
                div.className = "person";
                div.dataset.name = "人員" + idx;
                div.dataset.letter = "輪休"; 

                // 創建兩行的排版結構
                const nameSpan = document.createElement("span");
                nameSpan.textContent = "人員" + idx;
                div.appendChild(nameSpan);

                // 創建選單
                const select = document.createElement("select");
                select.className = "letter-select";

                // 加入固定選項
                fixedOptions.forEach(optionText => {
                    const option = document.createElement("option");
                    option.value = optionText;
                    option.textContent = optionText;
                    select.appendChild(option);
                });

                // 綁定選單變更事件
                select.addEventListener("change", (e) => {
                    const selectedLetter = e.target.value;
                    // 如果已有人員選了此字母，則重設該人員的選擇
                    document.querySelectorAll(".person").forEach(otherP => {
                        if (otherP !== div && otherP.dataset.letter === selectedLetter && !fixedOptions.includes(selectedLetter)) {
                            otherP.dataset.letter = fixedOptions[0];
                            const otherSelect = otherP.querySelector(".letter-select");
                            if (otherSelect) {
                                otherSelect.value = fixedOptions[0];
                                // 更新另一個人員的選單選項
                                updateSelectOptionsForPerson(otherP);
                            }
                        }
                    });
                    div.dataset.letter = selectedLetter;
                });
                
                div.appendChild(select);

                // 為人員卡片加上拖曳事件
                div.draggable = true;
                div.addEventListener("dragstart", e => {
                    e.dataTransfer.setData("text/plain", div.dataset.name);
                    e.dataTransfer.setData(DRAG_TYPE, div.dataset.name);
                    draggedElement = div; // 儲存拖曳的元素
                });

                personGrid.appendChild(div);
            }
        }
    }

    // --- 2. 拖曳與人員選擇功能 ---
    function createCard(name) {
        let card = document.createElement("div");
        card.className = "card";
        // 儲存實際的姓名，而不是textContent
        card.dataset.personName = name;
        
        // 創建一個 span 來顯示姓名，這樣 removeBtn 的 x 就不會混入其中
        const nameSpan = document.createElement("span");
        nameSpan.textContent = name;
        card.appendChild(nameSpan);

        card.draggable = true;
        card.addEventListener("dragstart", e => {
            e.dataTransfer.setData("text/plain", name);
            e.dataTransfer.setData(DRAG_TYPE, name);
            draggedElement = card; // 儲存拖曳的元素
            setTimeout(() => card.style.display = "none", 0);
        });
        card.addEventListener("dragend", () => {
            // 如果拖曳結束後元素還在，才顯示出來
            if (card.parentNode) {
                card.style.display = "inline-block";
            }
        });

        // 新增: 移除按鈕功能
        let removeBtn = document.createElement("div");
        removeBtn.className = "remove-btn";
        removeBtn.textContent = "x";

        // 點擊事件，刪除卡片
        removeBtn.addEventListener("click", () => {
            card.remove();
        });

        card.appendChild(removeBtn);

        return card;
    }

    // 統一處理可拖曳區域的函式
    function makeDroppable(cell) {
        cell.addEventListener("dragover", e => {
            if (!e.dataTransfer.types.includes(DRAG_TYPE)) {
                e.preventDefault();
                return;
            }
            e.preventDefault();
            cell.classList.add("highlight");
        });

        cell.addEventListener("dragleave", () => cell.classList.remove("highlight"));

        cell.addEventListener("drop", e => {
            e.preventDefault();
            cell.classList.remove("highlight");
            let name = e.dataTransfer.getData(DRAG_TYPE);
            
            // 確保有拖曳元素
            if (!draggedElement) {
                return;
            }

            const isFromPersonnel = draggedElement.classList.contains("person");
            
            // 根據模式開關決定是替換還是新增/交換
            const isAppendMode = modeSwitch.checked;

            if (isFromPersonnel) {
                // 從人員清單拖曳：直接新增卡片，不理會模式
                if (!isAppendMode) {
                    // 單卡片模式，先移除舊卡片
                    if (cell.querySelector(".card")) {
                        cell.querySelector(".card").remove();
                    }
                }
                cell.appendChild(createCard(name));
            } else {
                // 從排班表拖曳：處理移動或交換
                const targetCard = cell.querySelector(".card");
                
                if (isAppendMode) {
                    // 多卡片模式：新增卡片並移除原卡片
                    cell.appendChild(createCard(name));
                    draggedElement.remove();
                } else {
                    // 單卡片模式：交換卡片
                    // 1. 取得原拖曳卡片的父元素 (原始儲存格)
                    const originalParent = draggedElement.parentNode;
                    // 2. 如果目標儲存格有卡片，將其插入到原儲存格
                    if (targetCard) {
                        originalParent.appendChild(targetCard);
                    }
                    // 3. 將原拖曳卡片插入到目標儲存格
                    cell.appendChild(draggedElement);
                }
            }
            
            // 重設拖曳元素
            draggedElement = null;
        });
    }

    // 點擊人員卡片，切換 active 樣式並更新選單選項
    document.querySelectorAll(".person").forEach(p => {
        p.addEventListener("click", (event) => {
            if (event.target.tagName === 'SELECT') {
                return;
            }
            p.classList.toggle("active");
            updateSelectOptionsForPerson(p);
        });
    });

    // 修正: 全域監聽拖曳開始事件，只允許特定元素拖曳
    document.addEventListener("dragstart", (e) => {
        // 如果拖曳的不是卡片或人員，阻止拖曳
        if (!e.target.classList.contains("card") && !e.target.classList.contains("person")) {
            e.preventDefault();
        }
    });

    // 修正: 全域監聽拖曳進入事件，明確阻止拖曳到不可放置區域
    document.addEventListener("dragenter", (e) => {
        const target = e.target;
        if (!e.dataTransfer.types.includes(DRAG_TYPE)) {
            e.preventDefault();
        }
    });

    // 修正: 全域監聽拖曳經過事件，明確阻止拖曳到不可放置區域
    document.addEventListener("dragover", (e) => {
        const target = e.target;
        if (
            !e.dataTransfer.types.includes(DRAG_TYPE) ||
            !target.classList.contains('droppable') ||
            target.tagName === 'TH' ||
            target.getAttribute("contenteditable") === "true"
        ) {
            e.preventDefault();
        }
    });

    // --- 3. 模板載入與套用功能 ---
    async function loadTemplates() {
        try {
            const res = await fetch(API_URL);
            templateData = await res.json();
            Object.keys(templateData).forEach(key => {
                const opt = document.createElement("option");
                opt.value = key;
                opt.textContent = key;
                templateSelect.appendChild(opt);
            });
            if (Object.keys(templateData).length > 0) {
                templateSelect.value = Object.keys(templateData)[0];
                templateSelect.dispatchEvent(new Event('change'));
            }
        } catch (e) {
            console.error("讀取模板失敗:", e);
        }
    }

    // 尋找模板中所有排班代號
    function findLettersInTemplate(templateData) {
        const letters = new Set();
        if (templateData && templateData.data) {
            templateData.data.forEach(row => {
                row.slice(6).forEach(cell => {
                    if (typeof cell === 'string' && cell.trim() !== '') {
                        const bracketMatch = cell.match(/\{\{(.*?)\}\}/);
                        if (bracketMatch && bracketMatch[1].trim() !== '') {
                            letters.add(bracketMatch[1].trim());
                        } else {
                            if (isNaN(cell.trim()) && cell.trim() !== '') {
                                letters.add(cell.trim());
                            }
                        }
                    }
                });
            });
        }
        return Array.from(letters).sort();
    }

    // 根據 active 狀態更新單個人員的選單選項
    function updateSelectOptionsForPerson(personDiv) {
        const select = personDiv.querySelector(".letter-select");
        if (!select) return;

        // 移除所有舊選項
        while (select.options.length > 0) {
            select.remove(0);
        }

        if (personDiv.classList.contains("active")) {
            // 已選取狀態，顯示「代號」和排班代號
            const emptyOption = document.createElement("option");
            emptyOption.value = "";
            emptyOption.textContent = "代號";
            select.appendChild(emptyOption);

            templateLetters.forEach(letter => {
                const option = document.createElement("option");
                option.value = letter;
                option.textContent = letter;
                select.appendChild(option);
            });
            // 預設選回原本的代號，如果有的話
            if (personDiv.dataset.letter && !fixedOptions.includes(personDiv.dataset.letter)) {
                select.value = personDiv.dataset.letter;
            } else {
                // 如果原本是休假，則改為選取「代號」
                personDiv.dataset.letter = "";
                select.value = "";
            }
        } else {
            // 未選取狀態，顯示休假選項
            fixedOptions.forEach(optionText => {
                const option = document.createElement("option");
                option.value = optionText;
                option.textContent = optionText;
                select.appendChild(option);
            });
            // 預設選回原本的休假選項，如果有的話
            if (personDiv.dataset.letter && fixedOptions.includes(personDiv.dataset.letter)) {
                select.value = personDiv.dataset.letter;
            } else {
                // 如果原本是排班代號，則改回「輪休」
                personDiv.dataset.letter = fixedOptions[0];
                select.value = fixedOptions[0];
            }
        }
    }

    // 動態套用選取的模板
    templateSelect.addEventListener("change", () => {
        const key = templateSelect.value;
        if (!key || !templateData[key]) return;
        currentTemplate = templateData[key];

        // 每次選取新模板時，更新可選取的字母清單並更新所有下拉選單
        templateLetters = findLettersInTemplate(currentTemplate);

        // 清空表頭和表格內容
        const fixedHeaderRows = Array.from(thead.querySelectorAll('tr')).slice(0, 2);
        thead.innerHTML = '';
        fixedHeaderRows.forEach(row => thead.appendChild(row));
        tbody.innerHTML = '';

        // 重設所有人員狀態
        document.querySelectorAll(".person").forEach(p => {
            p.classList.remove("active");
            p.dataset.letter = "輪休";
            updateSelectOptionsForPerson(p); 
        });

        // 動態生成表頭 (包含合併儲存格)
        const headers = currentTemplate.headers;

        // 步驟1: 處理通用垂直合併 (rowspan)
        for (let i = 0; i < headers.length; i++) {
            for (let j = 0; j < headers[i].length; j++) {
                if (headers[i][j] === "VISITED" || headers[i][j] === "") {
                    continue;
                }
                let rowspan = 1;
                for (let k = i + 1; k < headers.length; k++) {
                    if (headers[k][j] === "") {
                        rowspan++;
                    } else {
                        break;
                    }
                }
                if (rowspan > 1) {
                    headers[i][j] = { text: headers[i][j], rowspan: rowspan };
                    for (let r = i + 1; r < i + rowspan; r++) {
                        headers[r][j] = "VISITED";
                    }
                }
            }
        }

        // 步驟2: 處理針對性水平合併 (colspan) - 第7欄到第20欄
        for (let i = 0; i < headers.length; i++) {
            for (let j = 6; j < Math.min(20, headers[i].length); j++) {
                if (headers[i][j] === "VISITED" || headers[i][j] === "") {
                    continue;
                }
                let colspan = 1;
                for (let k = j + 1; k < Math.min(20, headers[i].length); k++) {
                    if (headers[i][k] === "" && headers[i][k] !== "VISITED") {
                        colspan++;
                    } else {
                        break;
                    }
                }
                if (colspan > 1) {
                    if (typeof headers[i][j] === 'object') {
                        headers[i][j].colspan = colspan;
                    } else {
                        headers[i][j] = { text: headers[i][j], colspan: colspan };
                    }
                    for (let c = j + 1; c < j + colspan; c++) {
                        if (c < headers[i].length) {
                            headers[i][c] = "VISITED";
                        }
                    }
                }
            }
        }

        // 步驟3: 處理通用水平合併 (colspan) - 其他欄位
        for (let i = 0; i < headers.length; i++) {
            for (let j = 0; j < headers[i].length; j++) {
                if (headers[i][j] === "VISITED" || (j >= 6 && j < 20)) {
                    continue;
                }
                let colspan = 1;
                for (let k = j + 1; k < headers[i].length; k++) {
                    if (headers[i][k] === "" && headers[i][k] !== "VISITED" && (k < 6 || k >= 20)) {
                        colspan++;
                    } else {
                        break;
                    }
                }
                if (colspan > 1) {
                    if (typeof headers[i][j] === 'object') {
                        headers[i][j].colspan = colspan;
                    } else {
                        headers[i][j] = { text: headers[i][j], colspan: colspan };
                    }
                    for (let c = j + 1; c < j + colspan; c++) {
                        if (c < headers[i].length) {
                            headers[i][c] = "VISITED";
                        }
                    }
                }
            }
        }

        // 步驟4: 生成 HTML
        for (let i = 0; i < headers.length; i++) {
            const tr = document.createElement("tr");
            let cellCount = 0;
            for (let j = 0; j < headers[i].length; j++) {
                const headerCell = headers[i][j];
                if (headerCell === "VISITED" || headerCell === "") {
                    continue;
                }

                const headerText = typeof headerCell === 'object' ? headerCell.text : headerCell;
                const th = document.createElement("th");
                th.textContent = headerText;
                th.draggable = false;

                if (i < 2 && j >= 1 && j <= 5) {
                    th.setAttribute("contenteditable", "true");
                }

                if (typeof headerCell === 'object') {
                    if (headerCell.rowspan) th.setAttribute("rowspan", headerCell.rowspan);
                    if (headerCell.colspan) th.setAttribute("colspan", headerCell.colspan);
                }

                tr.appendChild(th);
                cellCount++;
            }
            thead.appendChild(tr);
        }

        // 動態生成表格內容 (時段與排班內容)
        currentTemplate.data.forEach(rowData => {
            const tr = document.createElement("tr");
            rowData.forEach((cellContent, cellIndex) => {
                const td = document.createElement("td");
                td.textContent = cellContent;

                if (cellIndex >= 1) {
                    td.className = "droppable";
                    makeDroppable(td);
                }

                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
    });

    // --- 4. 點擊按鈕，將字母轉換為人員卡片 ---
    assignPersonsBtn.addEventListener("click", () => {
        if (!currentTemplate) {
            alert("請先選擇模板！");
            return;
        }

        const letterMap = {};
        // 根據所有人員的下拉選單選擇來建立 letterMap
        document.querySelectorAll(".person").forEach(p => {
            const letter = p.dataset.letter;
            if (letter) {
                letterMap[letter] = p.dataset.name;
            }
        });

        document.querySelectorAll("#schedule tbody tr").forEach(tr => {
            tr.querySelectorAll("td.droppable").forEach(td => {
                const cellText = td.textContent.trim();
                let letter = null;
                const bracketMatch = cellText.match(/\{\{(.*?)\}\}/);
                if (bracketMatch) {
                    letter = bracketMatch[1].trim();
                } else if (cellText !== "") {
                    letter = cellText;
                }

                if (letter && letterMap[letter]) {
                    td.innerHTML = "";
                    td.appendChild(createCard(letterMap[letter]));
                }
            });
        });
    });

    // 新增下載功能
    downloadExcelBtn.addEventListener("click", () => {
        const table = document.getElementById("schedule");
        if (!table) {
            console.error('找不到 id 為 "schedule" 的表格元素。');
            return;
        }

        // 步驟1: 建立一個工作表物件
        const ws = {};
        // 設定欄寬 (例如: 26欄，從A到Z)
        ws['!cols'] = Array.from({ length: 26 }, () => ({ width: 10 }));
        ws['!cols'][0] = { width: 15 }; // 第一欄寬度設為15

        // 步驟2: 讀取 HTML 表格內容並填入工作表物件
        const allRows = table.querySelectorAll('tr');
        
        allRows.forEach((row, rowIndex) => {
            const cells = row.querySelectorAll('th, td');
            let colIndex = 0;

            cells.forEach((cell) => {
                // 處理合併儲存格
                const colspan = parseInt(cell.getAttribute('colspan') || '1', 10);
                const rowspan = parseInt(cell.getAttribute('rowspan') || '1', 10);

                // 找到下一個沒有被合併的儲存格
                while (ws[XLSX.utils.encode_cell({ c: colIndex, r: rowIndex })]) {
                    colIndex++;
                }

                const cellRef = XLSX.utils.encode_cell({ c: colIndex, r: rowIndex });
                
                let cellText = cell.textContent.trim();
                // 如果是卡片，取得 data-person-name
                if (cell.querySelector('.card')) {
                    const cards = cell.querySelectorAll('.card');
                    cellText = Array.from(cards).map(card => card.dataset.personName).join(', ');
                }

                // 將資料寫入工作表
                ws[cellRef] = { v: cellText };

                // 設定樣式 (例如：將第一列和第二列設為粗體)
                if (rowIndex < 2) {
                    ws[cellRef].s = { font: { bold: true } };
                }

                // 處理合併
                if (rowspan > 1 || colspan > 1) {
                    if (!ws['!merges']) ws['!merges'] = [];
                    ws['!merges'].push({
                        s: { r: rowIndex, c: colIndex },
                        e: { r: rowIndex + rowspan - 1, c: colIndex + colspan - 1 }
                    });
                }
                
                colIndex += colspan;
            });
        });

        // 步驟3: 設定工作表範圍
        const range = XLSX.utils.decode_range(XLSX.utils.encode_range({
            s: { c: 0, r: 0 },
            e: { c: ws['!cols'].length - 1, r: allRows.length - 1 }
        }));
        ws['!ref'] = XLSX.utils.encode_range(range);


        // 步驟4: 建立 Excel 工作簿
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, ws, "排班表");
        
        // 步驟5: 產生並下載檔案
        const excelData = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([excelData], { type: 'application/octet-stream' });
        const downloadLink = document.createElement('a');
        downloadLink.href = URL.createObjectURL(blob);
        downloadLink.download = '勤務排班表.xlsx';
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
        URL.revokeObjectURL(downloadLink.href);
    });

    // 啟動頁面時載入模板
    loadTemplates();
});