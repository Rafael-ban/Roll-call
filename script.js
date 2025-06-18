// 全局变量和常量
let students = [];
let absences = [];

const CONFIG = {
    MAX_FILE_SIZE: 5 * 1024 * 1024, // 5MB
    MAX_NAME_LENGTH: 20,
    MAX_STUDENTS: 200,
    SYSTEM_START_DATE: '2025-04-23'
};

// DOM元素缓存
const DOM = {};

// 初始化DOM元素缓存
function initDOMCache() {
    const elements = [
        'customAlertModal', 'customAlertTitle', 'customAlertMessage', 'customAlertOkButton',
        'courseDate', 'courseTime', 'courseName', 'classroom', 'counselor', 'classInfo',
        'studentList', 'studentListError', 'manualAbsenceList', 'attendanceTable',
        'runningDays', 'loadLists', 'markAbsent', 'markLate', 'markLeave', 'markMissing',
        'invertSelectionButton', 'saveAttendanceTxt', 'exportFullExcelDesktop',
        'saveAttendanceMobile', 'saveOptionsModal'
    ];
    
    elements.forEach(id => {
        DOM[id] = document.getElementById(id);
    });
}

// 优化的自定义提示框
function showCustomAlert(message, title = "提示") {
    if (!DOM.customAlertModal) {
        console.error("Custom alert modal not initialized!");
        alert((title !== "提示" ? title + ":\n" : "") + message);
        return;
    }

    DOM.customAlertTitle.textContent = title;
    DOM.customAlertMessage.textContent = message;
    DOM.customAlertModal.classList.add('is-visible');
}

// 工具函数：文本清理（使用textContent替代innerHTML避免XSS）
const sanitizeInput = (input) => String(input).trim();

// 优化的日期格式化
const formatDate = (date) => date.toLocaleDateString('zh-CN', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
});

// 计算运行天数（优化性能）
function calculateRunningDays() {
    const startDate = new Date(CONFIG.SYSTEM_START_DATE);
    const currentDate = new Date();
    const diffDays = Math.floor((currentDate - startDate) / (1000 * 60 * 60 * 24));
    DOM.runningDays.textContent = diffDays;
}

// 文件验证（提前返回优化）
function checkFileType(file) {
    if (!file) return true;
    
    if (file.size > CONFIG.MAX_FILE_SIZE) {
        throw new Error(`文件大小不能超过${CONFIG.MAX_FILE_SIZE / 1024 / 1024}MB`);
    }
    
    const fileName = file.name.toLowerCase();
    const fileExtension = fileName.split('.').pop();
    
    if (!['txt', 'xlsx'].includes(fileExtension)) {
        throw new Error('仅支持 .txt 和 .xlsx 格式的文件');
    }
    
    const validMimeTypes = [
        'text/plain',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    ];
    
    if (!validMimeTypes.includes(file.type)) {
        throw new Error('文件类型不正确');
    }
    
    return true;
}

// 姓名验证函数
const isValidName = (name) => {
    return name && 
           name.length <= CONFIG.MAX_NAME_LENGTH && 
           /^[\u4e00-\u9fa5a-zA-Z\s·.。•]+$/.test(name);
};

// 优化的文件读取函数
async function readStudentListFromFile(file) {
    return new Promise((resolve, reject) => {
        const fileExtension = file.name.split('.').pop().toLowerCase();
        const reader = new FileReader();
        
        reader.onerror = () => reject(new Error(`读取${fileExtension.toUpperCase()}文件失败`));
        
        if (fileExtension === 'txt') {
            reader.onload = (event) => {
                try {
                    const names = event.target.result
                        .split('\n')
                        .map(name => name.trim())
                        .filter(name => {
                            if (!name) return false;
                            if (!isValidName(name)) {
                                throw new Error(`无效姓名: ${name}`);
                            }
                            return true;
                        });

                    if (names.length > CONFIG.MAX_STUDENTS) {
                        throw new Error(`名单不能超过${CONFIG.MAX_STUDENTS}人`);
                    }

                    resolve(names);
                } catch (error) {
                    reject(error);
                }
            };
            reader.readAsText(file, 'UTF-8');
        } else if (fileExtension === 'xlsx') {
            reader.onload = (event) => {
                try {
                    const data = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);
                    
                    const names = [];
                    const nameFields = ['姓名', '名字', '学生姓名', '学员姓名'];
                    
                    for (const row of jsonData) {
                        let name = null;
                        
                        // 查找姓名字段
                        for (const field of nameFields) {
                            if (row[field]) {
                                name = row[field];
                                break;
                            }
                        }
                        
                        // 如果没找到，遍历所有字段
                        if (!name) {
                            for (const [key, value] of Object.entries(row)) {
                                if (typeof value === 'string' && 
                                    value.trim() && 
                                    !(/^\d+$/.test(value)) && 
                                    !key.includes('序') &&
                                    !key.toLowerCase().includes('no')) {
                                    name = value.trim();
                                    break;
                                }
                            }
                        }
                        
                        if (name && isValidName(name)) {
                            names.push(name);
                        }
                    }

                    if (names.length > CONFIG.MAX_STUDENTS) {
                        throw new Error(`名单不能超过${CONFIG.MAX_STUDENTS}人`);
                    }
                    
                    resolve(names);
                } catch (error) {
                    reject(new Error('解析Excel文件失败: ' + error.message));
                }
            };
            reader.readAsArrayBuffer(file);
        }
    });
}

// 优化的手动输入解析
function parseManualNames(input) {
    if (!input?.trim()) return [];
    
    const names = input
        .split(/[\n\s]+/)
        .map(name => name.trim())
        .filter(name => {
            if (!name) return false;
            if (!isValidName(name)) {
                throw new Error(`姓名格式错误: ${name}`);
            }
            return true;
        });
    
    if (names.length > CONFIG.MAX_STUDENTS) {
        throw new Error(`名单不能超过${CONFIG.MAX_STUDENTS}人`);
    }
    
    return [...new Set(names)].sort();
}

// 获取请假名单
async function getAbsenceList() {
    return parseManualNames(DOM.manualAbsenceList.value);
}

// 优化的显示考勤函数
function displayAttendance() {
    const tbody = DOM.attendanceTable.getElementsByTagName('tbody')[0];
    
    if (students.length === 0) {
        tbody.innerHTML = '<tr><td colspan="3" style="text-align:center">没有学生名单数据</td></tr>';
        return;
    }
    
    // 使用模板字符串和innerHTML一次性更新
    const rowsHTML = students.map(student => {
        const isAbsent = absences.includes(student);
        return `
            <tr>
                <td>${sanitizeInput(student)}</td>
                <td>
                    <select class="${isAbsent ? 'absent' : 'present'}">
                        <option value="出勤" ${!isAbsent ? 'selected' : ''}>出勤</option>
                        <option value="缺勤">缺勤</option>
                        <option value="请假" ${isAbsent ? 'selected' : ''}>请假</option>
                        <option value="迟到">迟到</option>
                        <option value="早退">早退</option>
                    </select>
                </td>
                <td>
                    <input type="text" class="remarks-input" placeholder="填写备注...">
                </td>
            </tr>
        `;
    }).join('');
    
    tbody.innerHTML = rowsHTML;
    
    // 使用事件委托处理select变化
    tbody.addEventListener('change', handleSelectChange);
}

// 事件委托处理select变化
function handleSelectChange(e) {
    if (e.target.tagName === 'SELECT') {
        const select = e.target;
        select.className = select.value === '出勤' ? 'present' : 'absent';
    }
}

// 本地存储操作
const storage = {
    save() {
        const formData = {
            courseName: DOM.courseName.value,
            classroom: DOM.classroom.value,
            counselor: DOM.counselor.value,
            classInfo: DOM.classInfo.value
        };
        localStorage.setItem('attendanceFormData', JSON.stringify(formData));
    },
    
    load() {
        const savedData = localStorage.getItem('attendanceFormData');
        if (savedData) {
            const formData = JSON.parse(savedData);
            DOM.courseName.value = formData.courseName || '';
            DOM.classroom.value = formData.classroom || '';
            DOM.counselor.value = formData.counselor || '';
            DOM.classInfo.value = formData.classInfo || '';
        }
    }
};

// 辅助函数：设置单元格样式
function setCellStyle(ws, cellAddress, style) {
    if (!ws[cellAddress]) {
        ws[cellAddress] = { t: 'z', s: style };
    } else {
        ws[cellAddress].s = style;
    }
}

// 导出完整Excel表格
function exportFullExcel() {
    const courseName = DOM.courseName.value || '未命名课程';
    const courseDate = DOM.courseDate.value;
    const courseTime = DOM.courseTime.value;
    const classroom = DOM.classroom.value || '未指定教室';
    const counselor = DOM.counselor.value || '未指定辅导员';
    const classInfo = DOM.classInfo.value || '未指定班级';
    
    const tbody = DOM.attendanceTable.getElementsByTagName('tbody')[0];
    const rows = tbody.rows;
    
    if (rows.length === 0) {
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
    for (let i = 0; i < rows.length; i++) {
        const name = rows[i].cells[0].textContent;
        const statusSelect = rows[i].cells[1].getElementsByTagName('select')[0];
        const status = statusSelect ? statusSelect.value : '未知';
        const remarksInput = rows[i].cells[2].getElementsByTagName('input')[0];
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

// 标记学生状态
async function markStudentsStatus(status) {
    try {
        absences = await getAbsenceList();
        
        if (absences.length === 0) {
            showCustomAlert('请先上传标记名单文件或手动输入学生姓名！');
            return;
        }
        
        if (students.length === 0) {
            showCustomAlert('请先加载学生总名单！');
            return;
        }

        const tbody = DOM.attendanceTable.getElementsByTagName('tbody')[0];
        const selects = tbody.getElementsByTagName('select');
        
        // 批量更新选择状态
        Array.from(selects).forEach((select, index) => {
            const row = select.closest('tr');
            const name = row.cells[0].textContent;
            
            if (absences.includes(name)) {
                select.value = status;
                select.className = 'absent';
            }
        });
    } catch (error) {
        showCustomAlert(`标记${status}学生出错: ${error.message}`, "错误");
    }
}

// 反选功能
function invertSelection() {
    if (students.length === 0) {
        showCustomAlert('请先加载学生名单！');
        return;
    }

    const tbody = DOM.attendanceTable.getElementsByTagName('tbody')[0];
    const selects = tbody.getElementsByTagName('select');

    Array.from(selects).forEach(select => {
        if (select.value === '出勤') {
            select.value = '缺勤';
            select.className = 'absent';
        } else {
            select.value = '出勤';
            select.className = 'present';
        }
    });
}

// 保存TXT格式考勤记录
function saveAttendanceToTxt() {
    const courseName = DOM.courseName.value || '未命名课程';
    const courseDate = DOM.courseDate.value || '';
    const courseTimeValue = DOM.courseTime.value || '';
    const courseDateTime = courseDate && courseTimeValue ? 
        `${courseDate} ${courseTimeValue}` : 
        new Date().toLocaleString('zh-CN');
    const classroom = DOM.classroom.value || '未指定教室';
    const counselor = DOM.counselor.value || '未指定辅导员';
    const classInfo = DOM.classInfo.value || '未指定班级';
    
    storage.save();
    
    const rows = DOM.attendanceTable.getElementsByTagName('tbody')[0].rows;
    
    if (rows.length === 0 || students.length === 0) {
        showCustomAlert('没有出勤记录可保存！');
        return;
    }
    
    // 统计数据
    const stats = {
        present: 0,
        absent: [],
        leave: [],
        late: [],
        early: []
    };
    
    Array.from(rows).forEach(row => {
        const name = row.cells[0].textContent;
        const status = row.cells[1].getElementsByTagName('select')[0].value;
        
        switch(status) {
            case '出勤': stats.present++; break;
            case '缺勤': stats.absent.push(name); break;
            case '请假': stats.leave.push(name); break;
            case '迟到': stats.late.push(name); break;
            case '早退': stats.early.push(name); break;
        }
    });
    
    const outputContent = [
        `时间：${courseDateTime}`,
        `课程：${courseName}`,
        `专业班级：${classInfo}`,
        `教室：${classroom}`,
        `应到：${students.length}`,
        `实到：${stats.present}`,
        `辅导员: ${counselor}`,
        `迟到: ${stats.late.length} ${stats.late.join("、")}`,
        `早退: ${stats.early.length} ${stats.early.join("、")}`,
        `旷课: ${stats.absent.length} ${stats.absent.join("、")}`,
        `请假: ${stats.leave.length} ${stats.leave.join("、")}`
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

// 事件绑定函数
function bindEvents() {
    // 自定义提示框事件
    DOM.customAlertOkButton?.addEventListener('click', () => {
        DOM.customAlertModal?.classList.remove('is-visible');
    });
    
    DOM.customAlertModal?.addEventListener('click', (e) => {
        if (e.target === DOM.customAlertModal) {
            DOM.customAlertModal.classList.remove('is-visible');
        }
    });

    // 文件输入验证
    DOM.studentList?.addEventListener('change', (e) => {
        const file = e.target.files[0];
        try {
            if (file) {
                checkFileType(file);
                DOM.studentListError.textContent = '';
            }
        } catch (error) {
            DOM.studentListError.textContent = '错误：' + error.message;
            e.target.value = '';
        }
    });

    // 按钮事件
    DOM.loadLists?.addEventListener('click', async () => {
        const studentListFile = DOM.studentList.files[0];
        
        if (!studentListFile) {
            showCustomAlert('请上传总名单文件！');
            return;
        }
        
        try {
            checkFileType(studentListFile);
            students = await readStudentListFromFile(studentListFile);
            absences = await getAbsenceList();
            displayAttendance();
        } catch (error) {
            showCustomAlert('加载名单出错: ' + error.message, "错误");
        }
    });

    // 标记按钮事件（使用事件委托）
    const markButtons = {
        markAbsent: '请假',
        markLate: '迟到', 
        markLeave: '早退',
        markMissing: '缺勤'
    };
    
    Object.entries(markButtons).forEach(([id, status]) => {
        DOM[id]?.addEventListener('click', () => markStudentsStatus(status));
    });

    DOM.invertSelectionButton?.addEventListener('click', invertSelection);
    DOM.saveAttendanceTxt?.addEventListener('click', saveAttendanceToTxt);
    DOM.exportFullExcelDesktop?.addEventListener('click', exportFullExcel);

    // 移动端保存选项模态框
    DOM.saveAttendanceMobile?.addEventListener('click', () => {
        DOM.saveOptionsModal?.classList.add('is-visible');
    });

    const closeButton = document.querySelector('.modal .close-button');
    closeButton?.addEventListener('click', () => {
        DOM.saveOptionsModal?.classList.remove('is-visible');
    });

    document.getElementById('saveSimpleAttendance')?.addEventListener('click', () => {
        saveAttendanceToTxt();
        DOM.saveOptionsModal?.classList.remove('is-visible');
    });

    document.getElementById('saveFullAttendance')?.addEventListener('click', () => {
        exportFullExcel();
        DOM.saveOptionsModal?.classList.remove('is-visible');
    });

    // 点击模态框外部关闭
    window.addEventListener('click', (e) => {
        if (e.target === DOM.saveOptionsModal) {
            DOM.saveOptionsModal?.classList.remove('is-visible');
        }
    });
}

// DOMContentLoaded事件处理
document.addEventListener('DOMContentLoaded', function() {
    // 初始化DOM缓存
    initDOMCache();
    
    // 设置当前日期和时间
    const now = new Date();
    const today = now.toISOString().split('T')[0];
    const time = now.toTimeString().slice(0, 5);
    
    DOM.courseDate.value = today;
    DOM.courseTime.value = time;
    
    // 计算系统运行时间
    calculateRunningDays();
    setInterval(calculateRunningDays, 86400000); // 每天更新一次
    
    // 加载保存的表单数据
    storage.load();
    
    // 绑定所有事件
    bindEvents();
    
    // 添加教程链接
    addTutorialLink();
});
    const courseDateTime = courseDate && courseTimeValue ? 
        `${courseDate} ${courseTimeValue}` : 
        new Date().toLocaleString('zh-CN');
    const classroom = document.getElementById('classroom').value || '未指定教室';
    const counselor = document.getElementById('counselor').value || '未指定辅导员';
    const classInfo = document.getElementById('classInfo').value || '未指定班级';
    
    // 保存表单数据
    saveFormData();
    
    const rows = document.getElementById('attendanceTable').getElementsByTagName('tbody')[0].rows;
    
    if (rows.length === 0 || students.length === 0) {
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
    
    const totalCount = students.length;
    
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
