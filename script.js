// 全局变量声明
let students = [];
let absences = [];

// 常量定义
const MAX_FILE_SIZE = 5 * 1024 * 1024; // 5MB
const MAX_NAME_LENGTH = 20;
const MAX_STUDENTS = 200;
const SYSTEM_START_DATE = '2025-04-23';

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
                        if (row['姓名']) {
                            name = row['姓名'];
                        } else if (row['名字'] || row['学生姓名'] || row['学员姓名']) {
                            name = row['名字'] || row['学生姓名'] || row['学员姓名'];
                        } else {
                            for (const key in row) {
                                const value = row[key];
                                if (typeof value === 'string' && value.trim() !== '') {
                                    name = value.trim();
                                    break;
                                }
                            }
                        }
                        
                        if (name && typeof name === 'string') {
                            name = name.trim();
                            if (name.length > MAX_NAME_LENGTH) {
                                throw new Error(`姓名长度不能超过${MAX_NAME_LENGTH}个字符`);
                            }
                            if (/^[\u4e00-\u9fa5a-zA-Z\s·.。•]+$/.test(name)) {
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

function parseStatusNames() {
    const input = document.getElementById('statusNameList').value.trim();
    if (!input) {
        throw new Error('请输入要标记的学生姓名！');
    }
    
    return parseManualNames(input); // 重用现有的名单解析函数
}

// 标记指定学生的状态
function markStudentStatus(targetNames, status) {
    if (!targetNames || targetNames.length === 0) {
        alert('请输入要标记的学生姓名！');
        return;
    }

    const rows = document.getElementById('attendanceTable').getElementsByTagName('tbody')[0].rows;
    if (rows.length === 0) {
        alert('没有学生名单数据！');
        return;
    }

    let markedCount = 0;
    const notFoundNames = [];

    // 遍历要标记的名单
    for (const targetName of targetNames) {
        let found = false;
        // 在表格中查找对应的学生
        for (let i = 0; i < rows.length; i++) {
            const rowName = rows[i].cells[0].textContent;
            if (rowName === targetName) {
                const select = rows[i].cells[1].getElementsByTagName('select')[0];
                select.value = status;
                // 更新样式
                select.classList.remove('present');
                select.classList.add('absent');
                found = true;
                markedCount++;
                break;
            }
        }
        if (!found) {
            notFoundNames.push(targetName);
        }
    }

    // 显示操作结果
    let message = `已成功标记 ${markedCount} 名学生为"${status}"。`;
    if (notFoundNames.length > 0) {
        message += `\n未找到以下学生：${notFoundNames.join('、')}`;
    }
    alert(message);

    // 清空输入框
    document.getElementById('statusNameList').value = '';
}


// 获取请假名单
async function getAbsenceList() {
    const isFileTabActive = document.querySelector('.tab[data-tab="file-tab"]').classList.contains('active');
    
    if (isFileTabActive) {
        const absenceListFile = document.getElementById('absenceList').files[0];
        
        if (absenceListFile) {
            if (!checkFileType(absenceListFile)) {
                throw new Error('请假名单必须是TXT或XLSX格式的文件！');
            }
            return await readStudentListFromFile(absenceListFile);
        }
        return [];
    } else {
        const manualInput = document.getElementById('manualAbsenceList').value;
        return parseManualNames(manualInput);
    }
}

// 使用DocumentFragment优化DOM操作
function displayAttendance() {
    const attendanceTableBody = document.getElementById('attendanceTable').getElementsByTagName('tbody')[0];
    attendanceTableBody.innerHTML = '';
    
    if (students.length === 0) {
        attendanceTableBody.innerHTML = '<tr><td colspan="2" style="text-align:center">没有学生名单数据</td></tr>';
        return;
    }
    
    const fragment = document.createDocumentFragment();
    
    students.forEach(student => {
        const row = document.createElement('tr');
        const cell1 = document.createElement('td');
        const cell2 = document.createElement('td');
        
        cell1.textContent = sanitizeInput(student);
        
        const select = document.createElement('select');
        select.innerHTML = `
            <option value="出勤">出勤</option>
            <option value="缺勤">缺勤</option>
            <option value="请假">请假</option>
            <option value="迟到">迟到</option>
            <option value="早退">早退</option>
        `;
        
        if (absences.includes(student)) {
            select.value = '请假';
        }
        
        select.addEventListener('change', (e) => {
            if (e.target.value === '出勤') {
                e.target.classList.remove('absent');
                e.target.classList.add('present');
            } else {
                e.target.classList.remove('present');
                e.target.classList.add('absent');
            }
        });
        
        if (select.value === '出勤') {
            select.classList.add('present');
        } else {
            select.classList.add('absent');
        }
        
        cell2.appendChild(select);
        row.appendChild(cell1);
        row.appendChild(cell2);
        fragment.appendChild(row);
    });
    
    attendanceTableBody.appendChild(fragment);
}

// 保存表单数据到localStorage
function saveFormData() {
    const formData = {
        courseName: document.getElementById('courseName').value,
        classroom: document.getElementById('classroom').value,
        counselor: document.getElementById('counselor').value,
        classInfo: document.getElementById('classInfo').value
    };
    localStorage.setItem('attendanceFormData', JSON.stringify(formData));
}

// 从localStorage加载表单数据
function loadFormData() {
    const savedData = localStorage.getItem('attendanceFormData');
    if (savedData) {
        const formData = JSON.parse(savedData);
        document.getElementById('courseName').value = formData.courseName || '';
        document.getElementById('classroom').value = formData.classroom || '';
        document.getElementById('counselor').value = formData.counselor || '';
        document.getElementById('classInfo').value = formData.classInfo || '';
    }
}

// 页面加载完成后的初始化
document.addEventListener('DOMContentLoaded', function() {
    // 设置当前日期和时间
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    const hours = String(today.getHours()).padStart(2, '0');
    const minutes = String(today.getMinutes()).padStart(2, '0');
    
    document.getElementById('courseDate').value = `${year}-${month}-${day}`;
    document.getElementById('courseTime').value = `${hours}:${minutes}`;
    
    // 计算系统运行时间
    calculateRunningDays();
    setInterval(calculateRunningDays, 86400000);
    
    // 设置标签页切换
    const tabs = document.querySelectorAll('.tab');
    tabs.forEach(tab => {
        tab.addEventListener('click', function() {
            tabs.forEach(t => t.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.remove('active');
            });
            this.classList.add('active');
            const tabId = this.getAttribute('data-tab');
            document.getElementById(tabId).classList.add('active');
        });
    });
    
    // 加载保存的表单数据
    loadFormData();
    
    // 添加教程链接
    addTutorialLink();

    // 标记迟到按钮事件
    document.getElementById('markSelectedLate').addEventListener('click', () => {
        try {
            const names = parseStatusNames();
            markStudentStatus(names, '迟到');
        } catch (error) {
            alert(error.message);
        }
    });

    // 标记早退按钮事件
    document.getElementById('markSelectedEarlyLeave').addEventListener('click', () => {
        try {
            const names = parseStatusNames();
            markStudentStatus(names, '早退');
        } catch (error) {
            alert(error.message);
        }
    });

    // 标记缺勤按钮事件
    document.getElementById('markSelectedAbsence').addEventListener('click', () => {
        try {
            const names = parseStatusNames();
            markStudentStatus(names, '缺勤');
        } catch (error) {
            alert(error.message);
        }
    });
    
    document.getElementById('saveExcel').addEventListener('click', saveAsExcel);
});

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

document.getElementById('absenceList').addEventListener('change', (event) => {
    const file = event.target.files[0];
    const errorElement = document.getElementById('absenceListError');
    
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

// 加载名单按钮事件
document.getElementById('loadLists').addEventListener('click', async () => {
    const studentListFile = document.getElementById('studentList').files[0];
    
    if (!studentListFile) {
        alert('请上传总名单文件！');
        return;
    }
    
    try {
        checkFileType(studentListFile);
        students = await readStudentListFromFile(studentListFile);
        absences = await getAbsenceList();
        displayAttendance();
    } catch (error) {
        alert('加载名单出错: ' + error.message);
    }
});

// 标记请假按钮事件
document.getElementById('markAbsent').addEventListener('click', async () => {
    try {
        absences = await getAbsenceList();
        
        if (absences.length === 0) {
            alert('请先上传请假名单文件或手动输入请假学生姓名！');
            return;
        }
        
        if (students.length === 0) {
            alert('请先加载学生总名单！');
            return;
        }
        
        displayAttendance();
    } catch (error) {
        alert('标记请假学生出错: ' + error.message);
    }
});

// 保存出勤记录按钮事件
document.getElementById('saveAttendance').addEventListener('click', () => {
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
    
    if (rows.length === 0 || students.length === 0) {
        alert('没有出勤记录可保存！');
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
    
    alert('出勤记录已保存！');
});

// 添加教程链接
function addTutorialLink() {
    const tutorialLink = document.createElement('a');
    tutorialLink.href = 'https://rafael.xiaoqiu.in/post/tutorial-not-so-intelligent-smart';
    tutorialLink.textContent = '教程';
    tutorialLink.id = 'tutorialLink';
    document.body.appendChild(tutorialLink);
}

// 在已有代码后面添加以下函数和事件监听器

// 创建Excel工作表数据
function createExcelData(courseDateTime, courseName, classInfo, classroom, counselor) {
    // 获取表格数据
    const rows = document.getElementById('attendanceTable').getElementsByTagName('tbody')[0].rows;
    
    // 准备表头
    const headers = [
        ['课程考勤记录'],
        [],
        ['基本信息'],
        ['时间', courseDateTime],
        ['课程', courseName],
        ['专业班级', classInfo],
        ['教室', classroom],
        ['辅导员', counselor],
        [],
        ['考勤详情'],
        ['姓名', '出勤状态']
    ];

    // 准备考勤数据
    const attendanceData = Array.from(rows).map(row => [
        row.cells[0].textContent,
        row.cells[1].getElementsByTagName('select')[0].value
    ]);

    // 统计数据
    let stats = {
        total: attendanceData.length,
        present: 0,
        absent: 0,
        leave: 0,
        late: 0,
        early: 0
    };

    attendanceData.forEach(([_, status]) => {
        switch(status) {
            case '出勤': stats.present++; break;
            case '缺勤': stats.absent++; break;
            case '请假': stats.leave++; break;
            case '迟到': stats.late++; break;
            case '早退': stats.early++; break;
        }
    });

    // 添加统计信息
    const summary = [
        [],
        ['统计信息'],
        ['应到人数', stats.total],
        ['实到人数', stats.present],
        ['缺勤人数', stats.absent],
        ['请假人数', stats.leave],
        ['迟到人数', stats.late],
        ['早退人数', stats.early]
    ];

    // 合并所有数据
    return [...headers, ...attendanceData, ...summary];
}

// 保存为Excel文件
function saveAsExcel() {
    // 获取课程信息
    const courseName = document.getElementById('courseName').value || '未命名课程';
    const courseDate = document.getElementById('courseDate').value || '';
    const courseTimeValue = document.getElementById('courseTime').value || '';
    const courseDateTime = courseDate && courseTimeValue ? 
        `${courseDate} ${courseTimeValue}` : 
        new Date().toLocaleString('zh-CN');
    const classroom = document.getElementById('classroom').value || '未指定教室';
    const counselor = document.getElementById('counselor').value || '未指定辅导员';
    const classInfo = document.getElementById('classInfo').value || '未指定班级';

    // 检查是否有数据
    const rows = document.getElementById('attendanceTable').getElementsByTagName('tbody')[0].rows;
    if (rows.length === 0 || students.length === 0) {
        alert('没有出勤记录可导出！');
        return;
    }

    try {
        // 创建工作表数据
        const wsData = createExcelData(courseDateTime, courseName, classInfo, classroom, counselor);

        // 创建工作簿和工作表
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(wsData);

        // 设置单元格合并
        ws['!merges'] = [
            // 标题合并单元格（A1:B1）
            {s: {r: 0, c: 0}, e: {r: 0, c: 1}}
        ];

        // 设置列宽
        ws['!cols'] = [
            {wch: 20}, // 第一列宽度
            {wch: 15}  // 第二列宽度
        ];

        // 添加工作表到工作簿
        XLSX.utils.book_append_sheet(wb, ws, '考勤记录');

        // 生成文件名
        const fileName = `${courseName}_${courseDate}_考勤记录.xlsx`;

        // 保存文件
        XLSX.writeFile(wb, fileName);

        alert('Excel文件已导出！');
    } catch (error) {
        alert('导出Excel文件失败：' + error.message);
    }
}