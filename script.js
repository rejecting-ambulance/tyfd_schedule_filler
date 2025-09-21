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
    const dateDisplayCell = document.getElementById('dateDisplayCell');

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

    // --- 日期處理函式 ---
    function formatDate(date) {
        const year = date.getFullYear();
        const month = date.getMonth() + 1;
        const day = date.getDate();
        const weekday = ['日', '一', '二', '三', '四', '五', '六'][date.getDay()];
        const mingGuoYear = year - 1911;
        return `${mingGuoYear}年${month}月${day} 星期${weekday}`;
    }

    function generateDateOptions() {
        const today = new Date();
        const options = [];

        // 前1天
        const yesterday = new Date(today);
        yesterday.setDate(today.getDate() - 1);
        options.push({ value: yesterday.toISOString(), text: formatDate(yesterday) });

        // 今天
        options.push({ value: today.toISOString(), text: formatDate(today) });

        // 後7天
        for (let i = 1; i <= 7; i++) {
            const nextDay = new Date(today);
            nextDay.setDate(today.getDate() + i);
            options.push({ value: nextDay.toISOString(), text: formatDate(nextDay) });
        }

        return options;
    }

    function setupDateSelect() {
        const dateSelectContainer = document.getElementById('dateSelectContainer');
        const dateOptions = generateDateOptions();
        const selectElement = document.createElement('select');
        selectElement.id = 'dateSelect';
        
        dateOptions.forEach(option => {
            const opt = document.createElement("option");
            opt.value = option.value;
            opt.textContent = option.text;
            selectElement.appendChild(opt);
        });

        // 預設選中今天
        selectElement.value = new Date().toISOString();
        
        // 插入下拉選單
        dateSelectContainer.appendChild(selectElement);
        
        // 綁定變更事件
        selectElement.addEventListener('change', (e) => {
            dateDisplayCell.textContent = e.target.options[e.target.selectedIndex].textContent;
        });

        // 初始化表格的日期顯示
        dateDisplayCell.textContent = dateOptions[1].text;
    }
    
    setupDateSelect();


    // --- 1. 初始化人員清單 ---
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

                const nameSpan = document.createElement("span");
                nameSpan.textContent = "人員" + idx;
                div.appendChild(nameSpan);

                const select = document.createElement("select");
                select.className = "letter-select";

                fixedOptions.forEach(optionText => {
                    const option = document.createElement("option");
                    option.value = optionText;
                    option.textContent = optionText;
                    select.appendChild(option);
                });

                select.addEventListener("change", (e) => {
                    const selectedLetter = e.target.value;
                    document.querySelectorAll(".person").forEach(otherP => {
                        if (otherP !== div && otherP.dataset.letter === selectedLetter && !fixedOptions.includes(selectedLetter)) {
                            otherP.dataset.letter = fixedOptions[0];
                            const otherSelect = otherP.querySelector(".letter-select");
                            if (otherSelect) {
                                otherSelect.value = fixedOptions[0];
                                updateSelectOptionsForPerson(otherP);
                            }
                        }
                    });
                    div.dataset.letter = selectedLetter;
                });
                
                div.appendChild(select);
                div.draggable = true;
                div.addEventListener("dragstart", e => {
                    e.dataTransfer.setData("text/plain", div.dataset.name);
                    e.dataTransfer.setData(DRAG_TYPE, div.dataset.name);
                    draggedElement = div; 
                });
                personGrid.appendChild(div);
            }
        }
    }

    // --- 2. 拖曳與人員選擇功能 ---
    function createCard(name) {
        let card = document.createElement("div");
        card.className = "card";
        card.dataset.personName = name;
        const nameSpan = document.createElement("span");
        nameSpan.textContent = name;
        card.appendChild(nameSpan);
        card.draggable = true;
        card.addEventListener("dragstart", e => {
            e.dataTransfer.setData("text/plain", name);
            e.dataTransfer.setData(DRAG_TYPE, name);
            draggedElement = card; 
            setTimeout(() => card.style.display = "none", 0);
        });
        card.addEventListener("dragend", () => {
            if (card.parentNode) {
                card.style.display = "inline-block";
            }
        });

        let removeBtn = document.createElement("div");
        removeBtn.className = "remove-btn";
        removeBtn.textContent = "x";
        removeBtn.addEventListener("click", () => {
            card.remove();
        });
        card.appendChild(removeBtn);
        return card;
    }

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
            
            if (!draggedElement) {
                return;
            }

            const isFromPersonnel = draggedElement.classList.contains("person");
            const isAppendMode = modeSwitch.checked;

            if (isFromPersonnel) {
                if (!isAppendMode) {
                    if (cell.querySelector(".card")) {
                        cell.querySelector(".card").remove();
                    }
                }
                cell.appendChild(createCard(name));
            } else {
                const targetCard = cell.querySelector(".card");
                if (isAppendMode) {
                    cell.appendChild(createCard(name));
                    draggedElement.remove();
                } else {
                    const originalParent = draggedElement.parentNode;
                    if (targetCard) {
                        originalParent.appendChild(targetCard);
                    }
                    cell.appendChild(draggedElement);
                }
            }
            draggedElement = null;
        });
    }

    document.querySelectorAll(".person").forEach(p => {
        p.addEventListener("click", (event) => {
            if (event.target.tagName === 'SELECT') {
                return;
            }
            p.classList.toggle("active");
            updateSelectOptionsForPerson(p);
        });
    });

    document.addEventListener("dragstart", (e) => {
        if (!e.target.classList.contains("card") && !e.target.classList.contains("person")) {
            e.preventDefault();
        }
    });

    document.addEventListener("dragenter", (e) => {
        const target = e.target;
        if (!e.dataTransfer.types.includes(DRAG_TYPE)) {
            e.preventDefault();
        }
    });

    document.addEventListener("dragover", (e) => {
        const target = e.target;
        if (
            !e.dataTransfer.types.includes(DRAG_TYPE) ||
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

    function updateSelectOptionsForPerson(personDiv) {
        const select = personDiv.querySelector(".letter-select");
        if (!select) return;

        while (select.options.length > 0) {
            select.remove(0);
        }

        if (personDiv.classList.contains("active")) {
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
            if (personDiv.dataset.letter && !fixedOptions.includes(personDiv.dataset.letter)) {
                select.value = personDiv.dataset.letter;
            } else {
                personDiv.dataset.letter = "";
                select.value = "";
            }
        } else {
            fixedOptions.forEach(optionText => {
                const option = document.createElement("option");
                option.value = optionText;
                option.textContent = optionText;
                select.appendChild(option);
            });
            if (personDiv.dataset.letter && fixedOptions.includes(personDiv.dataset.letter)) {
                select.value = personDiv.dataset.letter;
            } else {
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
        
        // **修正：這裡新增了讀取模板代號的邏輯**
        templateLetters = findLettersInTemplate(currentTemplate);

        // 1. 取得並保留固定的表頭行
        const fixedHeaderRows = Array.from(thead.querySelectorAll('tr')).slice(0, 2);

        // 2. 清空表頭和表格內容
        thead.innerHTML = '';
        tbody.innerHTML = '';

        // 3. 重新附加固定的表頭行
        fixedHeaderRows.forEach(row => thead.appendChild(row));

        // 重置所有人員的狀態
        document.querySelectorAll(".person").forEach(p => {
            p.classList.remove("active");
            p.dataset.letter = "輪休";
            updateSelectOptionsForPerson(p); 
        });

        // 4. 處理並渲染來自模板的新表頭 (thead)
        const headers = JSON.parse(JSON.stringify(currentTemplate.headers));

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
        for (let i = 0; i < headers.length; i++) {
            for (let j = 0; j < headers[i].length; j++) {
                if (headers[i][j] === "VISITED" || headers[i][j] === "") {
                    continue;
                }
                let colspan = 1;
                for (let k = j + 1; k < headers[i].length; k++) {
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

        for (let i = 0; i < headers.length; i++) {
            const tr = document.createElement("tr");
            let originalCellIndex = 0;
            for (let j = 0; j < headers[i].length; j++) {
                const headerCell = headers[i][j];
                
                if (headerCell === "VISITED" || headerCell === "" || headerCell === null) {
                    continue;
                }

                if (i === 0 && originalCellIndex === 6) {
                    const newTh = document.createElement('th');
                    newTh.textContent = '在隊訓練(一)';
                    newTh.setAttribute('rowspan', '3');
                    tr.appendChild(newTh);
                    const newTh2 = document.createElement('th');
                    newTh2.textContent = '在隊訓練(二)';
                    newTh2.setAttribute('rowspan', '3');
                    tr.appendChild(newTh2);
                    originalCellIndex += 2;
                    
                    newTh.classList.add('fixed-width-th');
                    newTh2.classList.add('fixed-width-th');
                }

                const th = document.createElement("th");
                th.textContent = (typeof headerCell === 'object' ? headerCell.text : headerCell);
                th.draggable = false;
                
                if (i < 2 && j >= 1 && j <= 5) {
                    th.setAttribute("contenteditable", "true");
                }
                
                if (typeof headerCell === 'object') {
                    if (headerCell.rowspan) th.setAttribute("rowspan", headerCell.rowspan);
                    if (headerCell.colspan) th.setAttribute("colspan", headerCell.colspan);
                }

                th.classList.add('fixed-width-th');

                tr.appendChild(th);
                originalCellIndex++;
            }

            if (i === 0) {
                const newTh3 = document.createElement('th');
                newTh3.textContent = '勤務輪流順序與服勤人員對照表';
                newTh3.setAttribute('rowspan', '2');
                newTh3.setAttribute('colspan', '8');
                tr.appendChild(newTh3);
                newTh3.classList.add('fixed-width-th');
            }
            if (i === 2) {
                const titles = ['編號', '姓名', '編號', '姓名', '編號', '姓名', '編號', '姓名'];
                titles.forEach(text => {
                    const th = document.createElement('th');
                    th.textContent = text;
                    th.classList.add('fixed-width-th');
                    tr.appendChild(th);
                });
            }
            thead.appendChild(tr);
        }
        
        // 5. 處理並渲染表格主體 (tbody)
        
        const topFixedData = [
            { time: '07:30    08:00', content: '車輛器材保養環境清理', colspan: 25 },
            { time: '08:00    08:30', content: '勤前教育', colspan: 25 }
        ];

        topFixedData.forEach(item => {
            const tr = document.createElement('tr');
            const cell1 = document.createElement('td');
            cell1.textContent = item.time;
            tr.appendChild(cell1);
            const cell2 = document.createElement('td');
            cell2.textContent = item.content;
            cell2.setAttribute('colspan', item.colspan);
            tr.appendChild(cell2);
            for (let i = 0; i < 8; i++) {
                const emptyCell = document.createElement('td');
                emptyCell.setAttribute('contenteditable', 'true');
                tr.appendChild(emptyCell);
            }
            tbody.appendChild(tr);
        });

        currentTemplate.data.forEach((rowData, rowIndex) => {
            const tr = document.createElement("tr");
            rowData.forEach((cellContent, dataIndex) => {
                const td = document.createElement("td");
                td.textContent = cellContent;
                td.className = "droppable";
                makeDroppable(td);
                tr.appendChild(td);

                if (dataIndex === 5) {
                    if (rowIndex === 0 || rowIndex === 6 || rowIndex === 13) {
                        const newTd1 = document.createElement("td");
                        newTd1.setAttribute("contenteditable", "true");
                        const newTd2 = document.createElement("td");
                        newTd2.setAttribute("contenteditable", "true");
                        
                        let rowspanValue = 0;
                        if (rowIndex === 0) {
                            rowspanValue = 6;
                            newTd1.textContent = '0900~1100';
                            newTd2.textContent = '1400~1600';
                        } else if (rowIndex === 6) {
                            rowspanValue = 7;
                        } else if (rowIndex === 13) {
                            rowspanValue = 11;
                        }
                        if (rowspanValue > 0) {
                            newTd1.setAttribute("rowspan", rowspanValue);
                            newTd2.setAttribute("rowspan", rowspanValue);
                        }
                        tr.appendChild(newTd1);
                        tr.appendChild(newTd2);
                    }
                }
            });
            
            if(rowIndex === 13) {
                const finalTd1 = document.createElement('td');
                finalTd1.setAttribute('contenteditable', 'true');
                finalTd1.textContent = '備勤';
                tr.appendChild(finalTd1);

                const finalTd2 = document.createElement('td');
                finalTd2.setAttribute('contenteditable', 'true');
                finalTd2.setAttribute('colspan', '6');
                finalTd2.textContent = '';
                tr.appendChild(finalTd2);
                
                const finalTd3 = document.createElement('td');
                finalTd3.setAttribute('contenteditable', 'true');
                finalTd3.textContent = '';
                tr.appendChild(finalTd3);
            } else if(rowIndex === 14) {
                const finalTd1 = document.createElement('td');
                finalTd1.setAttribute('contenteditable', 'true');
                finalTd1.textContent = '輪休';
                tr.appendChild(finalTd1);
                
                const finalTd2 = document.createElement('td');
                finalTd2.setAttribute('contenteditable', 'true');
                finalTd2.setAttribute('colspan', '7');
                finalTd2.textContent = '';
                tr.appendChild(finalTd2);
            } else if(rowIndex === 15) {
                const finalTd1 = document.createElement('td');
                finalTd1.setAttribute('contenteditable', 'true');
                finalTd1.textContent = '請休';
                tr.appendChild(finalTd1);
                
                const finalTd2 = document.createElement('td');
                finalTd2.setAttribute('contenteditable', 'true');
                finalTd2.setAttribute('colspan', '3');
                finalTd2.textContent = '';
                tr.appendChild(finalTd2);

                const finalTd3 = document.createElement('td');
                finalTd3.setAttribute('contenteditable', 'true');
                finalTd3.textContent = '補休';
                tr.appendChild(finalTd3);
                
                const finalTd4 = document.createElement('td');
                finalTd4.setAttribute('contenteditable', 'true');
                finalTd4.setAttribute('colspan', '3');
                finalTd4.textContent = '';
                tr.appendChild(finalTd4);
            } else if(rowIndex === 16) {
                const finalTd1 = document.createElement('td');
                finalTd1.setAttribute('contenteditable', 'true');
                finalTd1.textContent = '公假';
                tr.appendChild(finalTd1);
                
                const finalTd2 = document.createElement('td');
                finalTd2.setAttribute('contenteditable', 'true');
                finalTd2.setAttribute('colspan', '3');
                finalTd2.textContent = '';
                tr.appendChild(finalTd2);

                const finalTd3 = document.createElement('td');
                finalTd3.setAttribute('contenteditable', 'true');
                finalTd3.textContent = '月輪休';
                tr.appendChild(finalTd3);
                
                const finalTd4 = document.createElement('td');
                finalTd4.setAttribute('contenteditable', 'true');
                finalTd4.setAttribute('colspan', '3');
                finalTd4.textContent = '';
                tr.appendChild(finalTd4);
            } else if(rowIndex === 17) {
                const finalTd1 = document.createElement('td');
                finalTd1.setAttribute('contenteditable', 'true');
                finalTd1.textContent = '停休';
                tr.appendChild(finalTd1);
                
                const finalTd2 = document.createElement('td');
                finalTd2.setAttribute('contenteditable', 'true');
                finalTd2.setAttribute('colspan', '7');
                finalTd2.textContent = '';
                tr.appendChild(finalTd2);
            } else if(rowIndex === 18) {
                const finalTd1 = document.createElement('td');
                finalTd1.setAttribute('contenteditable', 'true');
                finalTd1.textContent = '其他';
                tr.appendChild(finalTd1);
                
                const finalTd2 = document.createElement('td');
                finalTd2.setAttribute('contenteditable', 'true');
                finalTd2.setAttribute('colspan', '7');
                finalTd2.textContent = '';
                tr.appendChild(finalTd2);
            } else if(rowIndex === 19) {
                const finalTd1 = document.createElement('td');
                finalTd1.setAttribute('contenteditable', 'true');
                finalTd1.setAttribute('colspan', '8');
                finalTd1.textContent = '勤務變更紀錄';
                tr.appendChild(finalTd1);
            } else if(rowIndex === 20) {
                const finalTd1 = document.createElement('td');
                finalTd1.setAttribute('contenteditable', 'true');
                finalTd1.setAttribute('colspan', '2');
                finalTd1.textContent = '人員姓名';
                tr.appendChild(finalTd1);
                
                const finalTd2 = document.createElement('td');
                finalTd2.setAttribute('contenteditable', 'true');
                finalTd2.textContent = '原定勤務';
                tr.appendChild(finalTd2);

                const finalTd3 = document.createElement('td');
                finalTd3.setAttribute('contenteditable', 'true');
                finalTd3.textContent = '變更勤務';
                tr.appendChild(finalTd3);
                
                const finalTd4 = document.createElement('td');
                finalTd4.setAttribute('contenteditable', 'true');
                finalTd4.setAttribute('colspan', '2');
                finalTd4.textContent = '變更原因';
                tr.appendChild(finalTd4);
                
                const finalTd5 = document.createElement('td');
                finalTd5.setAttribute('contenteditable', 'true');
                finalTd5.setAttribute('colspan', '2');
                finalTd5.textContent = '主管核章';
                tr.appendChild(finalTd5);
            } else {
                for(let i=0; i<8; i++) {
                    const finalTd = document.createElement('td');
                    finalTd.setAttribute('contenteditable', 'true');
                    tr.appendChild(finalTd);
                }
            }         
            tbody.appendChild(tr);
        });

        // 6. 新增底部「備註」欄位
        const lastExtraRow = document.createElement('tr');
        const lastCell = document.createElement('td');
        lastCell.textContent = '備註';
        lastCell.setAttribute('contenteditable', 'true');
        lastExtraRow.appendChild(lastCell);

        const mergedCell3 = document.createElement('td');
        mergedCell3.textContent = '';
        mergedCell3.setAttribute('contenteditable', 'true');
        mergedCell3.setAttribute('colspan', '33'); 
        lastExtraRow.appendChild(mergedCell3);
        tbody.appendChild(lastExtraRow);
    });

    // --- 4. 點擊按鈕，將字母轉換為人員卡片 ---
    assignPersonsBtn.addEventListener("click", () => {
        if (!currentTemplate) {
            alert("請先選擇模板！");
            return;
        }

        const letterMap = {};
        document.querySelectorAll(".person.active").forEach(p => {
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
                    // **修正：這裡新增了對非括號文字的檢查**
                    const foundLetter = templateLetters.find(l => l === cellText);
                    if (foundLetter) {
                        letter = foundLetter;
                    }
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

        const ws = {};
        const occupiedCells = [];

        ws['!cols'] = Array.from({ length: 26 }, () => ({ width: 10 }));
        ws['!cols'][0] = { width: 15 }; 

        const allRows = table.querySelectorAll('tr');
        
        allRows.forEach((row, rowIndex) => {
            const cells = row.querySelectorAll('th, td');
            let colIndex = 0;

            cells.forEach((cell) => {
                const colspan = parseInt(cell.getAttribute('colspan') || '1', 10);
                const rowspan = parseInt(cell.getAttribute('rowspan') || '1', 10);

                while (occupiedCells[rowIndex] && occupiedCells[rowIndex][colIndex]) {
                    colIndex++;
                }

                const cellRef = XLSX.utils.encode_cell({ c: colIndex, r: rowIndex });
                
                let cellText = cell.textContent.trim();
                
                if (cell.querySelector('.card')) {
                    const cards = cell.querySelectorAll('.card');
                    cellText = Array.from(cards).map(card => card.dataset.personName).join(', ');
                } else if (cell.id === 'dateDisplayCell') {
                    cellText = cell.textContent.trim();
                }

                ws[cellRef] = { v: cellText };

                if (rowIndex < 2) {
                    ws[cellRef].s = { font: { bold: true } };
                }

                if (rowspan > 1 || colspan > 1) {
                    if (!ws['!merges']) ws['!merges'] = [];
                    ws['!merges'].push({
                        s: { r: rowIndex, c: colIndex },
                        e: { r: rowIndex + rowspan - 1, c: colIndex + colspan - 1 }
                    });

                    for (let r = rowIndex; r < rowIndex + rowspan; r++) {
                        if (!occupiedCells[r]) occupiedCells[r] = [];
                        for (let c = colIndex; c < colIndex + colspan; c++) {
                            occupiedCells[r][c] = true;
                        }
                    }
                }
                
                colIndex += colspan;
            });
        });
        
        let maxCol = 0;
        let maxRow = allRows.length - 1;
        occupiedCells.forEach(row => {
            if (row && row.length > maxCol) maxCol = row.length;
        });

        const range = XLSX.utils.decode_range(XLSX.utils.encode_range({
            s: { c: 0, r: 0 },
            e: { c: maxCol - 1, r: maxRow }
        }));
        ws['!ref'] = XLSX.utils.encode_range(range);

        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, ws, "排班表");
        
        const excelData = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([excelData], { type: 'application/octet-stream' });
        const downloadLink = document.createElement('a');
        downloadLink.href = URL.createObjectURL(blob);
        downloadLink.download = '勤務排班表.xlsx';
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
    });
    
    // 初始化載入模板
    loadTemplates();
});