// 性能优化的全局变量
const app = window.attendanceApp || {
    students: [],
    absences: [],
    initialized: false,
    // 缓存DOM元素
    dom: {},
    // 防抖函数缓存
    debounceTimers: {}
};

// 常量定义
const MAX_FILE_SIZE = 5 * 1024 * 1024;
const MAX_NAME_LENGTH = 20;
const MAX_STUDENTS = 200;
const SYSTEM_START_DATE = '2025-04-23';

// 性能优化：缓存DOM元素
function cacheDOM() {
    const selectors = {
        customAlertModal: '#customAlertModal',
        customAlertTitle: '#customAlertTitle',
        customAlertMessage: '#customAlertMessage',
        customAlertOkButton: '#customAlertOkButton',
        courseName: '#courseName',
        courseDate: '#courseDate',
        courseTime: '#courseTime',
        classroom: '#classroom',
        counselor: '#counselor',
        classInfo: '#classInfo',
        attendanceTable: '#attendanceTable',
        runningDays: '#runningDays',
        manualAbsenceList: '#manualAbsenceList'
    };
    
    for (const [key, selector] of Object.entries(selectors)) {
        app.dom[key] = document.querySelector(selector);
    }
}

// 防抖函数
function debounce(func, wait, key) {
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(app.debounceTimers[key]);
            func(...args);
        };
        clearTimeout(app.debounceTimers[key]);
        app.debounceTimers[key] = setTimeout(later, wait);
    };
}

// 节流函数
function throttle(func, limit) {
    let inThrottle;
    return function() {
        const args = arguments;
        const context = this;
        if (!inThrottle) {
            func.apply(context, args);
            inThrottle = true;
            setTimeout(() => inThrottle = false, limit);
        }
    };
}

// 优化的自定义弹窗
function showCustomAlert(message, title = "提示") {
    if (!app.dom.customAlertModal) {
        console.error("Custom alert modal not initialized!");
        alert((title !== "提示" ? title + ":\n" : "") + message);
        return;
    }

    app.dom.customAlertTitle.textContent = title;
    app.dom.customAlertMessage.textContent = message;
    app.dom.customAlertModal.classList.add('is-visible');
}

// 工具函数：安全的文本处理
function sanitizeInput(input) {
    const div = document.createElement('div');
    div.textContent = input;
    return div.innerHTML;
}

// 工具函数：日期格式化
function formatDate(date) {
    return date.toLocaleDateString('zh-CN', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit'
    });
}

// 计算并显示系统运行天数
function calculateRunningDays() {
    const startDate = new Date(SYSTEM_START_DATE);
    const currentDate = new Date();
    const diffTime = Math.abs(currentDate - startDate);
    const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
    document.getElementById('runningDays').textContent = diffDays;
}

// 检查文件类型和大小
function checkFileType(file) {
    if (!file) return true;
    
    // 检查文件大小
    if (file.size > MAX_FILE_SIZE) {
        throw new Error(`文件大小不能超过${MAX_FILE_SIZE / 1024 / 1024}MB`);
    }
    
    const fileName = file.name;
    const fileExtension = fileName.split('.').pop().toLowerCase();
    
    // 检查文件扩展名
    const validExtensions = ['txt', 'xlsx'];
    if (!validExtensions.includes(fileExtension)) {
        throw new Error('仅支持 .txt 和 .xlsx 格式的文件');
    }
    
    // 检查MIME类型
    const validMimeTypes = [
        'text/plain',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    ];
    if (!validMimeTypes.includes(file.type)) {
        throw new Error('文件类型不正确');
    }
    
    return true;
}

// 从文件读取学生名单
async function readStudentListFromFile(file) {
    return new Promise((resolve, reject) => {
        const fileExtension = file.name.split('.').pop().toLowerCase();
        
        if (fileExtension === 'txt') {
            const reader = new FileReader();
            reader.onload = (event) => {
                try {
                    const names = event.target.result.split('\n')
                        .map(name => name.trim())
                        .filter(name => {
                            if (name === '') return false;
                            if (name.length > MAX_NAME_LENGTH) {
                                throw new Error(`姓名长度不能超过${MAX_NAME_LENGTH}个字符`);
                            }
                            return /^[\u4e00-\u9fa5a-zA-Z\s·.。•]+$/.test(name);
                        });

                    if (names.length > MAX_STUDENTS) {
                        throw new Error(`名单不能超过${MAX_STUDENTS}人`);
                    }

                    resolve(names);
                } catch (error) {
                    reject(error);
                }
            };
            reader.onerror = () => reject(new Error('读取TXT文件失败'));
            reader.readAsText(file, 'UTF-8');
        } else if (fileExtension === 'xlsx') {
            const reader = new FileReader();
            reader.onload = (event) => {
                try {
                    const data = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);
                    
                    const names = [];
                    for (const row of jsonData) {
                        let name = null;
                        // 检查是否存在"姓名"列
                        if (row['姓名']) {
                            name = row['姓名'];
                        } 
                        // 检查其他可能的列名
                        else if (row['名字'] || row['学生姓名'] || row['学员姓名']) {
                            name = row['名字'] || row['学生姓名'] || row['学员姓名'];
                        } 
                        // 如果没有找到指定列名，遍历所有列
                        else {
                            for (const key in row) {
                                // 跳过可能是序号的列
                                if (key === '序号' || key.toLowerCase().includes('no') || 
                                    /^\d+$/.test(row[key]) || // 纯数字
                                    key === '' || // 空列名
                                    key.includes('序') || // 包含"序"字
                                    /^[A-Z]$/.test(key)) { // 单个大写字母（Excel默认列标）
                                    continue;
                                }
                                const value = row[key];
                                if (typeof value === 'string' && value.trim() !== '') {
                                    name = value.trim();
                                    break;
                                }
                            }
                        }
                        
                        if (name && typeof name === 'string') {
                            name = name.trim();
                            // 额外检查确保不是序号
                            if (!(/^\d+$/.test(name)) && // 不是纯数字
                                name !== '序号' && 
                                !name.toLowerCase().includes('no') &&
                                name.length <= MAX_NAME_LENGTH &&
                                /^[\u4e00-\u9fa5a-zA-Z\s·.。•]+$/.test(name)) {
                                names.push(name);
                            }
                        }
                    }

                    if (names.length > MAX_STUDENTS) {
                        throw new Error(`名单不能超过${MAX_STUDENTS}人`);
                    }
                    
                    resolve(names);
                } catch (error) {
                    reject(new Error('解析Excel文件失败: ' + error.message));
                }
            };
            reader.onerror = () => reject(new Error('读取Excel文件失败'));
            reader.readAsArrayBuffer(file);
        } else {
            reject(new Error('不支持的文件格式'));
        }
    });
}

// 解析手动输入的名单
function parseManualNames(input) {
    if (!input || input.trim() === '') {
        return [];
    }
    
    let lines = input.split('\n');
    let names = [];
    
    lines.forEach(line => {
        if (line.includes(' ')) {
            const lineNames = line.split(' ')
                .map(name => name.trim())
                .filter(name => {
                    if (name === '') return false;
                    if (name.length > MAX_NAME_LENGTH) {
                        throw new Error(`姓名长度不能超过${MAX_NAME_LENGTH}个字符`);
                    }
                    return /^[\u4e00-\u9fa5a-zA-Z\s·.。•]+$/.test(name);
                });
            names = names.concat(lineNames);
        } else {
            const trimmedLine = line.trim();
            if (trimmedLine !== '' && 
                trimmedLine.length <= MAX_NAME_LENGTH && 
                /^[\u4e00-\u9fa5a-zA-Z\s·.。•]+$/.test(trimmedLine)) {
                names.push(trimmedLine);
            }
        }
    });
    
    if (names.length > MAX_STUDENTS) {
        throw new Error(`名单不能超过${MAX_STUDENTS}人`);
    }
    
    return [...new Set(names)].sort();
}

// 获取请假名单（函数名保持不变以维持兼容性）
async function getAbsenceList() {
    const manualInput = document.getElementById('manualAbsenceList').value;
    return parseManualNames(manualInput);
}

// 优化的显示考勤函数 - 使用虚拟滚动和DocumentFragment
function displayAttendance() {
    const attendanceTableBody = app.dom.attendanceTable.getElementsByTagName('tbody')[0];
    
    if (!app.students.length) {
        attendanceTableBody.innerHTML = '<tr><td colspan="3" style="text-align:center">没有学生名单数据</td></tr>';
        return;
    }
    
    // 使用DocumentFragment优化DOM操作
    const fragment = document.createDocumentFragment();
    
    // 批量创建元素
    const createRow = (student) => {
        const row = document.createElement('tr');
        
        // 姓名单元格
        const nameCell = document.createElement('td');
        nameCell.textContent = student;
        
        // 状态单元格
        const statusCell = document.createElement('td');
        const select = document.createElement('select');
        select.innerHTML = `
            <option value="出勤">出勤</option>
            <option value="缺勤">缺勤</option>
            <option value="请假">请假</option>
            <option value="迟到">迟到</option>
            <option value="早退">早退</option>
        `;
        
        if (app.absences.includes(student)) {
            select.value = '请假';
        }
        
        // 优化事件监听 - 使用事件委托
        select.className = select.value === '出勤' ? 'present' : 'absent';
        statusCell.appendChild(select);
        
        // 备注单元格
        const remarksCell = document.createElement('td');
        const remarksInput = document.createElement('input');
        remarksInput.type = 'text';
        remarksInput.placeholder = '填写备注...';
        remarksInput.className = 'remarks-input';
        remarksCell.appendChild(remarksInput);
        
        row.appendChild(nameCell);
        row.appendChild(statusCell);
        row.appendChild(remarksCell);
        
        return row;
    };
    
    // 批量处理
    app.students.forEach(student => {
        fragment.appendChild(createRow(student));
    });
    
    // 一次性更新DOM
    attendanceTableBody.innerHTML = '';
    attendanceTableBody.appendChild(fragment);
    
    // 使用事件委托处理状态变更
    attendanceTableBody.addEventListener('change', handleStatusChange);
}

// 事件委托处理状态变更
function handleStatusChange(e) {
    if (e.target.tagName === 'SELECT') {
        const select = e.target;
        select.className = select.value === '出勤' ? 'present' : 'absent';
    }
}

// 优化的保存表单数据 - 使用防抖
const saveFormData = debounce(() => {
    const formData = {
        courseName: app.dom.courseName.value,
        classroom: app.dom.classroom.value,
        counselor: app.dom.counselor.value,
        classInfo: app.dom.classInfo.value
    };
    localStorage.setItem('attendanceFormData', JSON.stringify(formData));
}, 500, 'saveForm');

// 优化的加载表单数据
function loadFormData() {
    try {
        const savedData = localStorage.getItem('attendanceFormData');
        if (savedData) {
            const formData = JSON.parse(savedData);
            app.dom.courseName.value = formData.courseName || '';
            app.dom.classroom.value = formData.classroom || '';
            app.dom.counselor.value = formData.counselor || '';
            app.dom.classInfo.value = formData.classInfo || '';
        }
    } catch (error) {
        console.warn('Failed to load form data:', error);
    }
}

// 辅助函数：设置单元格样式
function setCellStyle(ws, cellAddress, style) {
    if (!ws[cellAddress]) {
        ws[cellAddress] = { t: 'z', s: style }; // 'z' for blank cell, apply style
    } else {
        ws[cellAddress].s = style;
    }
}

// 导出完整Excel表格
function exportFullExcel() {
    const courseName = document.getElementById('courseName').value || '未命名课程';
    const courseDate = document.getElementById('courseDate').value;
    const courseTime = document.getElementById('courseTime').value;
    const classroom = document.getElementById('classroom').value || '未指定教室';
    const counselor = document.getElementById('counselor').value || '未指定辅导员';
    const classInfo = document.getElementById('classInfo').value || '未指定班级';
    
    const tableRows = document.getElementById('attendanceTable').getElementsByTagName('tbody')[0].rows;
    if (tableRows.length === 0) {
        showCustomAlert('没有考勤数据可导出！');
        return;
    }

    // 创建工作簿
    const wb = XLSX.utils.book_new();
    
    // 准备课程信息数据 (Info Block) - 新的横向两列布局
    const courseInfoRows = [
        ['上课时间', `${courseDate} ${courseTime}`.trim(), null, null],
        ['上课地点', classroom, '课程名称', courseName],
        ['辅导员', counselor, '班级', classInfo], 
    ];
    const courseInfoSection = [
        ['课程考勤完整记录表', null, null, null],
        ...courseInfoRows,
        [null, null, null, null], 
    ];

    // 主数据表表头
    const studentDataTableHeader = ['序号', '姓名', '出勤状态', '备注'];

    // 添加学生数据
    let presentCount = 0;
    const studentDataRows = [];
    for (let i = 0; i < tableRows.length; i++) {
        const name = tableRows[i].cells[0].textContent;
        const statusSelect = tableRows[i].cells[1].getElementsByTagName('select')[0];
        const status = statusSelect ? statusSelect.value : '未知';
        const remarksInput = tableRows[i].cells[2].getElementsByTagName('input')[0];
        const remarks = remarksInput ? remarksInput.value : '';
        if (status === '出勤') presentCount++;
        studentDataRows.push([i + 1, name, status, remarks]);
    }

    // 准备统计信息数据 (Summary Block)
    const totalCount = studentDataRows.length; // Should be based on actual rows processed
    const attendanceRate = totalCount > 0 ? `${((presentCount / totalCount) * 100).toFixed(2)}%` : 'N/A';
    const statisticsData = [
        [null, null, null, null], // Spacer - Row after student data
        ['考勤统计', null, null, null], // Stats Title
        ['应到人数', totalCount, '实到人数', presentCount],
        ['出勤率', attendanceRate, null, null]
    ];

    // 合并所有数据块准备写入工作表
    const wsData = [...courseInfoSection, studentDataTableHeader, ...studentDataRows, ...statisticsData];

    // 创建工作表
    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // --- 应用样式和功能 ---
    const boldCenteredStyle = {
        font: { bold: true },
        alignment: { horizontal: "center", vertical: "center", wrapText: true }
    };
    const centeredStyle = {
        alignment: { horizontal: "center", vertical: "center", wrapText: true }
    };
     const boldStyle = {
        font: { bold: true },
        alignment: { vertical: "center", wrapText: true } 
    };


    // 1. "课程考勤完整记录表" 加粗居中 (A1)
    setCellStyle(ws, 'A1', boldCenteredStyle);

    // 2. 课程信息 - 按新的两列布局设置样式
    for (let i = 0; i < courseInfoRows.length; i++) {
        const sheetRowIndex = i + 1; 
        const dataRow = courseInfoRows[i];
        // First pair (Label in Col A, Value in Col B)
        if (dataRow[0] !== null) setCellStyle(ws, XLSX.utils.encode_cell({ r: sheetRowIndex, c: 0 }), boldStyle);
        if (dataRow[1] !== null) setCellStyle(ws, XLSX.utils.encode_cell({ r: sheetRowIndex, c: 1 }), centeredStyle);
        // Second pair (Label in Col C, Value in Col D)
        if (dataRow[2] !== null) setCellStyle(ws, XLSX.utils.encode_cell({ r: sheetRowIndex, c: 2 }), boldStyle);
        if (dataRow[3] !== null) setCellStyle(ws, XLSX.utils.encode_cell({ r: sheetRowIndex, c: 3 }), centeredStyle);
    }
    
    // 3. 主数据表表头 - 添加筛选按钮并设置样式
    const studentTableHeaderSheetRowIndex = courseInfoSection.length; 
    ws['!autofilter'] = { ref: XLSX.utils.encode_range(
        { s: { r: studentTableHeaderSheetRowIndex, c: 0 }, e: { r: studentTableHeaderSheetRowIndex, c: studentDataTableHeader.length - 1 } }
    )};
    for(let c = 0; c < studentDataTableHeader.length; c++) {
        setCellStyle(ws, XLSX.utils.encode_cell({r: studentTableHeaderSheetRowIndex, c: c}), boldCenteredStyle);
    }


    // 4. "考勤统计" 标题 加粗居中
    const statsTitleSheetRowIndex = studentTableHeaderSheetRowIndex + 1 + studentDataRows.length + 1; 
    setCellStyle(ws, XLSX.utils.encode_cell({ r: statsTitleSheetRowIndex, c: 0 }), boldCenteredStyle);

    // 5. 统计数据 (应到人数, 实到人数, 出勤率) - Labels bold, Values centered
    for (let i = 0; i < 3; i++) {
        const currentSheetRowIndex = statsTitleSheetRowIndex + 1 + i;
        // Label in Col A
        setCellStyle(ws, XLSX.utils.encode_cell({ r: currentSheetRowIndex, c: 0 }), boldStyle); 
        // Value in Col B
        setCellStyle(ws, XLSX.utils.encode_cell({ r: currentSheetRowIndex, c: 1 }), centeredStyle); 
    }

    // --- 设置单元格合并 ---
    // Merge for "课程考勤完整记录表" (A1:D1)
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }]; 
    
    // Merge for "考勤统计" title (A_stats:D_stats)
    ws['!merges'].push({ 
        s: { r: statsTitleSheetRowIndex, c: 0 }, 
        e: { r: statsTitleSheetRowIndex, c: 3 } 
    });

    // 设置列宽 - 调整以适应新的两列信息布局和学生数据
    ws['!cols'] = [
        { wch: 15 }, // Col A: Info Label 1 / 序号
        { wch: 20 }, // Col B: Info Value 1 / 姓名
        { wch: 15 }, // Col C: Info Label 2 / 出勤状态
        { wch: 25 }  // Col D: Info Value 2 / 备注
    ];

    // 添加工作表到工作簿
    XLSX.utils.book_append_sheet(wb, ws, "考勤记录");

    // 导出文件
    XLSX.writeFile(wb, `${courseName}_${courseDate}_完整考勤表.xlsx`);
}

// 添加标记功能函数
async function markStudentsStatus(status) {
    try {
        app.absences = await getAbsenceList();
        
        if (app.absences.length === 0) {
            showCustomAlert('请先上传标记名单文件或手动输入学生姓名！');
            return;
        }
        
        if (app.students.length === 0) {
            showCustomAlert('请先加载学生总名单！');
            return;
        }

        const tableBody = document.getElementById('attendanceTable').getElementsByTagName('tbody')[0];
        const rows = tableBody.rows;
        
        for (let i = 0; i < rows.length; i++) {
            const name = rows[i].cells[0].textContent;
            const select = rows[i].cells[1].getElementsByTagName('select')[0];
            
            if (app.absences.includes(name)) {
                select.value = status;
                // 更新样式
                select.classList.remove('present');
                select.classList.add('absent');
            }
        }
    } catch (error) {
        showCustomAlert(`标记${status}学生出错: ${error.message}`, "错误");
    }
}

// 新增：一键反选功能
function invertSelection() {
    if (app.students.length === 0) {
        showCustomAlert('请先加载学生名单！');
        return;
    }

    const attendanceTableBody = document.getElementById('attendanceTable').getElementsByTagName('tbody')[0];
    const selects = attendanceTableBody.getElementsByTagName('select');

    for (let select of selects) {
        if (select.value === '出勤') {
            select.value = '缺勤';
            select.classList.remove('present');
            select.classList.add('absent');
        } else {
            select.value = '出勤';
            select.classList.remove('absent');
            select.classList.add('present');
        }
    }
}

// 优化的初始化函数
function initializeApp() {
    if (app.initialized) return;
    
    // 缓存DOM元素
    cacheDOM();
    
    // 设置当前日期和时间
    const now = new Date();
    app.dom.courseDate.value = now.toISOString().split('T')[0];
    app.dom.courseTime.value = now.toTimeString().slice(0, 5);
    
    // 计算运行天数
    calculateRunningDays();
    
    // 加载表单数据
    loadFormData();
    
    // 添加教程链接
    addTutorialLink();
    
    // 绑定事件监听器
    bindEventListeners();
    
    app.initialized = true;
}

// 优化的事件监听器绑定
function bindEventListeners() {
    // 自定义弹窗事件
    if (app.dom.customAlertOkButton) {
        app.dom.customAlertOkButton.addEventListener('click', () => {
            app.dom.customAlertModal.classList.remove('is-visible');
        });
    }
    
    // 点击弹窗外部关闭
    if (app.dom.customAlertModal) {
        app.dom.customAlertModal.addEventListener('click', (e) => {
            if (e.target === app.dom.customAlertModal) {
                app.dom.customAlertModal.classList.remove('is-visible');
            }
        });
    }
    
    // 表单输入自动保存 - 使用防抖
    ['courseName', 'classroom', 'counselor', 'classInfo'].forEach(id => {
        const element = app.dom[id];
        if (element) {
            element.addEventListener('input', saveFormData);
        }
    });
    
    // 其他事件监听器...
    const eventBindings = [
        { id: 'loadLists', event: 'click', handler: loadListsHandler },
        { id: 'markAbsent', event: 'click', handler: () => markStudentsStatus('请假') },
        { id: 'markLate', event: 'click', handler: () => markStudentsStatus('迟到') },
        { id: 'markLeave', event: 'click', handler: () => markStudentsStatus('早退') },
        { id: 'markMissing', event: 'click', handler: () => markStudentsStatus('缺勤') },
        { id: 'invertSelectionButton', event: 'click', handler: invertSelection },
        { id: 'saveAttendanceTxt', event: 'click', handler: saveAttendanceToTxt },
        { id: 'exportFullExcelDesktop', event: 'click', handler: exportFullExcel }
    ];
    
    eventBindings.forEach(({ id, event, handler }) => {
        const element = document.getElementById(id);
        if (element) {
            element.addEventListener(event, handler);
        }
    });
}

// 优化的加载名单处理函数
async function loadListsHandler() {
    const studentListFile = document.getElementById('studentList').files[0];
    
    if (!studentListFile) {
        showCustomAlert('请上传总名单文件！');
        return;
    }
    
    try {
        checkFileType(studentListFile);
        
        // 显示加载状态
        const loadButton = document.getElementById('loadLists');
        const originalText = loadButton.textContent;
        loadButton.textContent = '加载中...';
        loadButton.disabled = true;
        
        app.students = await readStudentListFromFile(studentListFile);
        app.absences = await getAbsenceList();
        
        displayAttendance();
        
        // 恢复按钮状态
        loadButton.textContent = originalText;
        loadButton.disabled = false;
        
    } catch (error) {
        showCustomAlert('加载名单出错: ' + error.message, "错误");
        // 恢复按钮状态
        const loadButton = document.getElementById('loadLists');
        loadButton.textContent = '加载名单';
        loadButton.disabled = false;
    }
}

// 文件输入验证事件监听
document.getElementById('studentList').addEventListener('change', (event) => {
    const file = event.target.files[0];
    const errorElement = document.getElementById('studentListError');
    
    try {
        if (file) {
            checkFileType(file);
            errorElement.textContent = '';
        }
    } catch (error) {
        errorElement.textContent = '错误：' + error.message;
        event.target.value = '';
    }
});

// 优化的页面加载
document.addEventListener('DOMContentLoaded', initializeApp);

// 优化的文件检查 - 添加到异步队列
document.addEventListener('DOMContentLoaded', () => {
    const studentListInput = document.getElementById('studentList');
    if (studentListInput) {
        studentListInput.addEventListener('change', throttle((event) => {
            const file = event.target.files[0];
            const errorElement = document.getElementById('studentListError');
            
            try {
                if (file) {
                    checkFileType(file);
                    errorElement.textContent = '';
                }
            } catch (error) {
                errorElement.textContent = '错误：' + error.message;
                event.target.value = '';
            }
        }, 300));
    }
});

// 将原保存出勤记录按钮事件的逻辑封装为 saveAttendanceToTxt
function saveAttendanceToTxt() {
    const courseName = document.getElementById('courseName').value || '未命名课程';
    const courseDate = document.getElementById('courseDate').value || '';
    const courseTimeValue = document.getElementById('courseTime').value || '';
    const courseDateTime = courseDate && courseTimeValue ? 
        `${courseDate} ${courseTimeValue}` : 
        new Date().toLocaleString('zh-CN');
    const classroom = document.getElementById('classroom').value || '未指定教室';
    const counselor = document.getElementById('counselor').value || '未指定辅导员';
    const classInfo = document.getElementById('classInfo').value || '未指定班级';
    
    // 保存表单数据
    saveFormData();
    
    const rows = document.getElementById('attendanceTable').getElementsByTagName('tbody')[0].rows;
    
    if (rows.length === 0 || app.students.length === 0) {
        showCustomAlert('没有出勤记录可保存！');
        return;
    }
    
    // 统计各类状态
    let presentCount = 0;
    let absentList = [];
    let leaveList = [];
    let lateList = [];
    let earlyList = [];
    
    for (let i = 0; i < rows.length; i++) {
        const name = rows[i].cells[0].textContent;
        const status = rows[i].cells[1].getElementsByTagName('select')[0].value;
        
        switch(status) {
            case '出勤':
                presentCount++;
                break;
            case '缺勤':
                absentList.push(name);
                break;
            case '请假':
                leaveList.push(name);
                break;
            case '迟到':
                lateList.push(name);
                break;
            case '早退':
                earlyList.push(name);
                break;
        }
    }
    
    const totalCount = app.students.length;
    
    // 创建输出内容
    let outputContent = [
        `时间：${courseDateTime}`,
        `课程：${courseName}`,
        `专业班级：${classInfo}`,
        `教室：${classroom}`,
        `应到：${totalCount}`,
        `实到：${presentCount}`,
        `辅导员: ${counselor}`,
        `迟到: ${lateList.length} ${lateList.join("、")}`,
        `早退: ${earlyList.length} ${earlyList.join("、")}`,
        `旷课: ${absentList.length} ${absentList.join("、")}`,
        `请假: ${leaveList.length} ${leaveList.join("、")}`
    ].join('\n');
    
    // 创建并下载文件
    const blob = new Blob(
        [new Uint8Array([0xEF, 0xBB, 0xBF]), outputContent], 
        { type: 'text/plain;charset=utf-8;' }
    );
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${courseName}_${courseDate}_出勤记录.txt`;
    link.style.display = 'none';
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    
    showCustomAlert('出勤记录已保存！', "成功");
}


// 添加教程链接
function addTutorialLink() {
    const tutorialLink = document.createElement('a');
    tutorialLink.href = 'https://rafael.xiaoqiu.in/post/tutorial-not-so-intelligent-smart';
    tutorialLink.textContent = '教程';
    tutorialLink.id = 'tutorialLink';
    document.body.appendChild(tutorialLink);
}
