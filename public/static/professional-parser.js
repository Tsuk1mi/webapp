// Профессиональный парсер организационных структур
class ProfessionalOrgParser {
    constructor() {
        this.debug = true;
    }

    // Главный метод парсинга
    async parseDocument(content, fileType) {
        console.log('Professional parser: Processing document...');
        
        if (fileType.includes('spreadsheet') || content.includes('Подчиненность')) {
            return this.parseSpreadsheet(content);
        } else if (fileType.includes('presentation')) {
            return this.parsePPTX(content);
        } else if (fileType.includes('pdf')) {
            return this.parsePDF(content);
        } else if (fileType.includes('word')) {
            return this.parseDOCX(content);
        }
        
        return this.parseGeneric(content);
    }

    // Парсинг таблиц (Google Sheets, Excel)
    parseSpreadsheet(content) {
        console.log('Parsing spreadsheet...');
        
        const rows = this.parseCSVAdvanced(content);
        if (rows.length < 2) return null;
        
        const headers = rows[0];
        const employees = [];
        const nodeMap = new Map();
        const vacancies = [];
        
        // Определяем индексы колонок
        const cols = {
            code: this.findColumn(headers, ['подчиненность', 'код', 'subordination', '№']),
            dept: this.findColumn(headers, ['подразделение', 'отдел', 'department', 'структура']),
            position: this.findColumn(headers, ['должность', 'position', 'title', 'роль']),
            name: this.findColumn(headers, ['фио', 'сотрудник', 'name', 'ф.и.о.', 'фамилия']),
            func: this.findColumn(headers, ['функционал', 'обязанности', 'функции', 'описание'])
        };
        
        console.log('Column mapping:', cols);
        
        // Обрабатываем строки данных
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;
            
            const code = cols.code >= 0 ? row[cols.code] : `${i}`;
            const position = cols.position >= 0 ? row[cols.position] : '';
            const name = cols.name >= 0 ? row[cols.name] : '';
            const dept = cols.dept >= 0 ? row[cols.dept] : '';
            const func = cols.func >= 0 ? row[cols.func] : '';
            
            // Проверяем на вакансию
            const isVacancy = this.isVacancy(name, position);
            
            const employee = {
                id: `node-${i}`,
                code: code || `${i}`,
                department: dept,
                position: position || 'Сотрудник',
                name: isVacancy ? 'ВАКАНСИЯ' : (name || 'Не указано'),
                functional: func,
                isVacancy: isVacancy,
                level: this.calculateLevel(code),
                children: []
            };
            
            if (isVacancy) {
                employee.displayName = `ВАКАНСИЯ: ${position}`;
                vacancies.push(employee);
            } else {
                employee.displayName = name;
            }
            
            employees.push(employee);
            nodeMap.set(employee.code, employee);
        }
        
        console.log(`Parsed ${employees.length} employees, ${vacancies.length} vacancies`);
        
        // Строим иерархию
        return this.buildHierarchy(employees, nodeMap);
    }

    // Парсинг DOCX
    parseDOCX(content) {
        console.log('Parsing DOCX document...');
        
        const employees = [];
        const lines = content.split(/[\n\r]+/);
        let currentDept = '';
        let currentBoss = null;
        
        // Извлекаем название отдела
        const deptMatch = content.match(/(?:структура|отдел|управление|департамент)\s+([^<\n]+)/i);
        if (deptMatch) {
            currentDept = deptMatch[1].trim();
        }
        
        // Парсим таблицы если есть
        if (content.includes('<table>')) {
            return this.parseTableFromHTML(content);
        }
        
        // Паттерны для поиска
        const patterns = {
            fio: /([А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?)\s+([А-ЯЁ][а-яё]+)\s+([А-ЯЁ][а-яё]+)/g,
            position: /(начальник|заместитель|руководитель|директор|менеджер|специалист|ведущий|главный|старший|инженер|аналитик|консультант|координатор|администратор|бухгалтер|юрист|экономист)(\s+[а-яё]+)*/gi
        };
        
        // Обрабатываем строки
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            if (!line.trim()) continue;
            
            // Ищем ФИО
            const fioMatches = [...line.matchAll(patterns.fio)];
            
            for (const match of fioMatches) {
                const fullName = match[0];
                const position = this.extractPositionFromLine(line, fullName);
                const functional = this.extractFunctionalFromContext(lines, i);
                
                const employee = {
                    id: `emp-${employees.length + 1}`,
                    code: `${employees.length + 1}`,
                    name: fullName,
                    position: position || 'Сотрудник',
                    department: currentDept,
                    functional: functional,
                    level: this.determineLevelByPosition(position),
                    isVacancy: false,
                    children: []
                };
                
                // Определяем начальника
                if (position && position.match(/начальник\s+отдела/i) && !position.match(/заместитель/i)) {
                    currentBoss = employee;
                    employee.level = 1;
                }
                
                employees.push(employee);
            }
        }
        
        // Проверяем на вакансии в тексте
        const vacancyMatches = content.match(/вакансия|потребность|требуется/gi);
        if (vacancyMatches) {
            for (const match of vacancyMatches) {
                const context = this.extractContextAround(content, match);
                const position = this.extractPositionFromLine(context, '');
                if (position) {
                    employees.push({
                        id: `vac-${employees.length + 1}`,
                        code: `${employees.length + 1}`,
                        name: 'ВАКАНСИЯ',
                        displayName: `ВАКАНСИЯ: ${position}`,
                        position: position,
                        department: currentDept,
                        isVacancy: true,
                        level: this.determineLevelByPosition(position),
                        children: []
                    });
                }
            }
        }
        
        return this.buildEmployeeHierarchy(employees, currentDept);
    }

    // Парсинг таблиц из HTML
    parseTableFromHTML(content) {
        console.log('Parsing HTML table...');
        
        const employees = [];
        const tableMatch = content.match(/<table[^>]*>([\s\S]*?)<\/table>/i);
        
        if (tableMatch) {
            const tableContent = tableMatch[1];
            const rows = tableContent.split(/<\/tr>/i);
            
            for (let i = 1; i < rows.length; i++) { // Пропускаем заголовок
                const row = rows[i];
                const cells = row.split(/<\/td>/i).map(cell => 
                    cell.replace(/<[^>]*>/g, '').trim()
                );
                
                if (cells.length >= 2) {
                    // Извлекаем информацию из ячеек
                    const cellText = cells.join(' ');
                    const fioMatch = cellText.match(/([А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?)\s+([А-ЯЁ][а-яё]+)\s+([А-ЯЁ][а-яё]+)/);
                    
                    if (fioMatch) {
                        const name = fioMatch[0];
                        const position = this.extractPositionFromLine(cellText, name);
                        const functional = cellText.replace(name, '').replace(position, '').trim();
                        
                        employees.push({
                            id: `emp-${employees.length + 1}`,
                            code: cells[0] || `${employees.length + 1}`,
                            name: name,
                            position: position || 'Сотрудник',
                            functional: functional,
                            level: this.determineLevelByPosition(position),
                            isVacancy: false,
                            children: []
                        });
                    }
                }
            }
        }
        
        const deptMatch = content.match(/(?:структура|отдел)\s+([^<\n]+)/i);
        const department = deptMatch ? deptMatch[1].trim() : 'Организация';
        
        return this.buildEmployeeHierarchy(employees, department);
    }

    // Парсинг PPTX
    parsePPTX(content) {
        console.log('Parsing PPTX presentation...');
        
        const employees = [];
        const vacancies = [];
        const slides = content.split(/<page[^>]*>/);
        
        for (const slide of slides) {
            const text = slide.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ');
            const lines = text.split(/##|[\n\r]+/).filter(l => l.trim());
            
            for (let i = 0; i < lines.length; i++) {
                const line = lines[i].trim();
                if (!line) continue;
                
                // Проверка на вакансию
                if (line.match(/потребность|вакансия/i)) {
                    const position = this.extractPositionFromContext(lines, i);
                    if (position) {
                        vacancies.push({
                            id: `vac-${vacancies.length + 1}`,
                            name: 'ВАКАНСИЯ',
                            displayName: `ВАКАНСИЯ: ${position}`,
                            position: position,
                            isVacancy: true,
                            level: this.determineLevelByPosition(position),
                            children: []
                        });
                    }
                    continue;
                }
                
                // Поиск ФИО
                const fioPattern = /([А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?)\s+([А-ЯЁ][а-яё]+)(?:\s+([А-ЯЁ][а-яё]+))?/;
                const fioMatch = line.match(fioPattern);
                
                if (fioMatch && !this.isPosition(fioMatch[0])) {
                    const name = fioMatch[0];
                    const position = this.extractPositionFromContext(lines, i);
                    const department = this.extractDepartmentFromContext(lines, i);
                    
                    employees.push({
                        id: `emp-${employees.length + 1}`,
                        name: name,
                        position: position || 'Сотрудник',
                        department: department,
                        level: this.determineLevelByPosition(position),
                        isVacancy: false,
                        children: []
                    });
                }
            }
        }
        
        // Объединяем сотрудников и вакансии
        const allEmployees = [...employees, ...vacancies];
        console.log(`Found ${employees.length} employees and ${vacancies.length} vacancies`);
        
        return this.buildEmployeeHierarchy(allEmployees, 'Организация');
    }

    // Парсинг PDF
    parsePDF(content) {
        console.log('Parsing PDF document...');
        // PDF обрабатывается аналогично PPTX
        return this.parsePPTX(content);
    }

    // Общий парсинг
    parseGeneric(content) {
        return this.parseDOCX(content);
    }

    // Улучшенный парсер CSV
    parseCSVAdvanced(text) {
        const rows = [];
        const lines = text.split(/\r?\n/);
        
        for (const line of lines) {
            if (!line.trim()) continue;
            
            const row = [];
            let current = '';
            let inQuotes = false;
            
            for (let i = 0; i < line.length; i++) {
                const char = line[i];
                const nextChar = line[i + 1];
                
                if (char === '"') {
                    if (inQuotes && nextChar === '"') {
                        current += '"';
                        i++;
                    } else {
                        inQuotes = !inQuotes;
                    }
                } else if (char === ',' && !inQuotes) {
                    row.push(current.trim());
                    current = '';
                } else {
                    current += char;
                }
            }
            
            row.push(current.trim());
            rows.push(row);
        }
        
        return rows;
    }

    // Поиск колонки по названиям
    findColumn(headers, names) {
        for (let i = 0; i < headers.length; i++) {
            const header = headers[i].toLowerCase();
            for (const name of names) {
                if (header.includes(name.toLowerCase())) {
                    return i;
                }
            }
        }
        return -1;
    }

    // Проверка на вакансию
    isVacancy(name, position) {
        const nameCheck = !name || 
                         name.trim() === '' || 
                         name.toLowerCase() === 'вакансия' ||
                         name.toLowerCase() === 'потребность' ||
                         name.toLowerCase() === 'требуется' ||
                         name === '-' ||
                         name === '—';
        
        const positionCheck = position && (
            position.toLowerCase().includes('вакан') ||
            position.toLowerCase().includes('потребност') ||
            position.toLowerCase().includes('требуется')
        );
        
        return nameCheck || positionCheck;
    }

    // Вычисление уровня по коду
    calculateLevel(code) {
        if (!code) return 10;
        const dots = (code.toString().match(/\./g) || []).length;
        return dots + 1;
    }

    // Определение уровня по должности
    determineLevelByPosition(position) {
        if (!position) return 10;
        const pos = position.toLowerCase();
        
        if (pos.includes('генеральный') || pos.includes('президент')) return 1;
        if (pos.includes('заместитель') && (pos.includes('генерального') || pos.includes('президента'))) return 2;
        if (pos.includes('начальник управления') || pos.includes('директор департамента')) return 2;
        if (pos.includes('заместитель начальника управления')) return 3;
        if (pos.includes('начальник отдела')) return 3;
        if (pos.includes('заместитель начальника отдела')) return 4;
        if (pos.includes('руководитель проект')) return 4;
        if (pos.includes('руководитель направления')) return 4;
        if (pos.includes('руководитель')) return 5;
        if (pos.includes('главный')) return 5;
        if (pos.includes('ведущий')) return 6;
        if (pos.includes('старший')) return 7;
        if (pos.includes('специалист')) return 8;
        if (pos.includes('менеджер')) return 8;
        if (pos.includes('инженер')) return 8;
        if (pos.includes('аналитик')) return 8;
        
        return 9;
    }

    // Построение иерархии с учетом кодов подчиненности
    buildHierarchy(employees, nodeMap) {
        const roots = [];
        
        for (const employee of employees) {
            const code = employee.code;
            
            if (!code || code === '1' || !code.includes('.')) {
                roots.push(employee);
            } else {
                // Ищем родителя по коду
                const parts = code.split('.');
                let parentCode = parts.slice(0, -1).join('.');
                
                while (parentCode && !nodeMap.has(parentCode)) {
                    const parentParts = parentCode.split('.');
                    if (parentParts.length > 1) {
                        parentCode = parentParts.slice(0, -1).join('.');
                    } else {
                        parentCode = null;
                    }
                }
                
                if (parentCode && nodeMap.has(parentCode)) {
                    const parent = nodeMap.get(parentCode);
                    parent.children.push(employee);
                } else {
                    roots.push(employee);
                }
            }
        }
        
        if (roots.length === 1) {
            return roots[0];
        }
        
        return {
            id: 'root',
            name: 'Организация',
            position: '',
            department: '',
            children: roots
        };
    }

    // Построение иерархии из списка сотрудников
    buildEmployeeHierarchy(employees, departmentName) {
        if (employees.length === 0) return null;
        
        // Сортируем по уровню
        employees.sort((a, b) => a.level - b.level);
        
        // Группируем по уровням
        const levels = {};
        for (const emp of employees) {
            const level = emp.level || 10;
            if (!levels[level]) {
                levels[level] = [];
            }
            levels[level].push(emp);
        }
        
        // Строим иерархию
        const levelKeys = Object.keys(levels).map(Number).sort((a, b) => a - b);
        
        // Связываем уровни
        for (let i = 0; i < levelKeys.length - 1; i++) {
            const currentLevel = levels[levelKeys[i]];
            const nextLevel = levels[levelKeys[i + 1]];
            
            if (currentLevel && nextLevel) {
                // Распределяем подчиненных
                for (let j = 0; j < nextLevel.length; j++) {
                    const parentIndex = Math.floor(j * currentLevel.length / nextLevel.length);
                    if (currentLevel[parentIndex]) {
                        currentLevel[parentIndex].children.push(nextLevel[j]);
                    }
                }
            }
        }
        
        const roots = levels[levelKeys[0]] || [];
        
        if (roots.length === 1) {
            return roots[0];
        }
        
        return {
            id: 'org-root',
            name: departmentName || 'Организация',
            position: '',
            department: departmentName,
            children: roots
        };
    }

    // Извлечение должности из строки
    extractPositionFromLine(text, name) {
        const cleanText = text.replace(name, '').trim();
        
        const positionKeywords = [
            'начальник управления', 'заместитель начальника управления',
            'начальник отдела', 'заместитель начальника отдела',
            'технический руководитель', 'руководитель проектов',
            'руководитель направления', 'руководитель',
            'директор департамента', 'заместитель директора',
            'генеральный директор', 'директор',
            'главный специалист', 'ведущий специалист',
            'старший специалист', 'специалист',
            'главный инженер', 'ведущий инженер', 
            'старший инженер', 'инженер',
            'менеджер проекта', 'менеджер',
            'аналитик', 'консультант', 'координатор',
            'администратор', 'бухгалтер', 'юрист', 'экономист'
        ];
        
        for (const keyword of positionKeywords) {
            const regex = new RegExp(`(${keyword}[а-яё\\s]*)`, 'gi');
            const match = cleanText.match(regex);
            if (match) {
                return match[0].trim();
            }
        }
        
        // Проверяем на должность после тире
        if (cleanText.includes('-')) {
            const parts = cleanText.split('-');
            if (parts[0] && this.isPosition(parts[0])) {
                return parts[0].trim();
            }
        }
        
        return '';
    }

    // Извлечение должности из контекста
    extractPositionFromContext(lines, currentIndex) {
        // Ищем в соседних строках
        for (let offset = -2; offset <= 2; offset++) {
            if (offset === 0) continue;
            const index = currentIndex + offset;
            if (index >= 0 && index < lines.length) {
                const line = lines[index];
                if (this.isPosition(line) && !this.isFIO(line)) {
                    return line.trim();
                }
            }
        }
        return '';
    }

    // Извлечение департамента из контекста
    extractDepartmentFromContext(lines, currentIndex) {
        const deptKeywords = ['отдел', 'управление', 'департамент', 'направление', 'служба', 'подразделение'];
        
        for (let i = Math.max(0, currentIndex - 5); i <= Math.min(lines.length - 1, currentIndex + 5); i++) {
            const line = lines[i];
            if (!line) continue;
            
            for (const keyword of deptKeywords) {
                if (line.toLowerCase().includes(keyword)) {
                    return line.trim();
                }
            }
        }
        
        return '';
    }

    // Извлечение функционала из контекста
    extractFunctionalFromContext(lines, currentIndex) {
        for (let i = currentIndex + 1; i < Math.min(currentIndex + 5, lines.length); i++) {
            const line = lines[i];
            if (line && line.length > 50 && !this.isFIO(line) && !this.isPosition(line)) {
                return line.trim();
            }
        }
        return '';
    }

    // Извлечение контекста вокруг текста
    extractContextAround(text, searchText) {
        const index = text.indexOf(searchText);
        if (index === -1) return '';
        
        const start = Math.max(0, index - 100);
        const end = Math.min(text.length, index + 100);
        
        return text.substring(start, end);
    }

    // Проверка, является ли текст ФИО
    isFIO(text) {
        const fioPattern = /^([А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?)\s+([А-ЯЁ][а-яё]+)\s+([А-ЯЁ][а-яё]+)$/;
        return fioPattern.test(text.trim());
    }

    // Проверка, является ли текст должностью
    isPosition(text) {
        const keywords = [
            'начальник', 'заместитель', 'руководитель', 'директор',
            'менеджер', 'специалист', 'ведущий', 'главный', 'старший',
            'инженер', 'аналитик', 'консультант', 'координатор',
            'администратор', 'бухгалтер', 'юрист', 'экономист',
            'технический', 'проектов'
        ];
        
        const lower = text.toLowerCase();
        return keywords.some(k => lower.includes(k));
    }

    // Генерация Excel-совместимой таблицы
    generateExcelTable(structure) {
        const rows = [];
        this.flattenForExcel(structure, rows, 0, '');
        
        // Заголовки
        const headers = ['№', 'Код подчиненности', 'Подразделение', 'Должность', 'ФИО', 'Функционал', 'Уровень', 'Статус'];
        
        // Данные
        const data = [headers];
        rows.forEach((row, index) => {
            data.push([
                index + 1,
                row.code || '',
                row.department || '',
                row.position || '',
                row.name || '',
                row.functional || '',
                row.level || '',
                row.isVacancy ? 'Вакансия' : 'Занято'
            ]);
        });
        
        return data;
    }

    // Преобразование структуры в плоский список для Excel
    flattenForExcel(node, rows, level, parentCode) {
        if (!node) return;
        
        const code = node.code || (parentCode ? `${parentCode}.${rows.length + 1}` : '1');
        
        if (node.name && node.name !== 'Организация') {
            rows.push({
                code: code,
                name: node.isVacancy ? `ВАКАНСИЯ: ${node.position}` : node.name,
                position: node.position || '',
                department: node.department || '',
                functional: node.functional || '',
                level: level + 1,
                isVacancy: node.isVacancy || false
            });
        }
        
        if (node.children && node.children.length > 0) {
            for (let i = 0; i < node.children.length; i++) {
                this.flattenForExcel(
                    node.children[i], 
                    rows, 
                    node.name === 'Организация' ? level : level + 1,
                    node.name === 'Организация' ? '' : code
                );
            }
        }
    }

    // Создание HTML таблицы для отображения
    createHTMLTable(structure) {
        const excelData = this.generateExcelTable(structure);
        
        let html = `
            <h3 class="text-xl font-semibold mb-4">
                <i class="fas fa-table mr-2"></i>Организационная структура
            </h3>
            <div class="mb-4">
                <span class="text-sm text-gray-600">
                    Всего позиций: ${excelData.length - 1} | 
                    Вакансий: ${excelData.filter(row => row[7] === 'Вакансия').length}
                </span>
            </div>
            <div class="overflow-x-auto">
                <table class="min-w-full divide-y divide-gray-200">
                    <thead class="bg-gray-50 sticky top-0">
                        <tr>
        `;
        
        // Заголовки
        excelData[0].forEach(header => {
            html += `<th class="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase">${header}</th>`;
        });
        
        html += `
                        </tr>
                    </thead>
                    <tbody class="bg-white divide-y divide-gray-200">
        `;
        
        // Данные
        for (let i = 1; i < excelData.length; i++) {
            const row = excelData[i];
            const isVacancy = row[7] === 'Вакансия';
            
            html += `
                <tr class="hover:bg-gray-50 ${isVacancy ? 'bg-yellow-50' : ''}">
            `;
            
            row.forEach((cell, index) => {
                const style = isVacancy && index === 4 ? 'text-red-600 font-medium' : '';
                html += `<td class="px-3 py-2 text-sm ${style}">${cell || '-'}</td>`;
            });
            
            html += `</tr>`;
        }
        
        html += `
                    </tbody>
                </table>
            </div>
        `;
        
        return html;
    }
}

// Экспорт в глобальную область
window.ProfessionalOrgParser = ProfessionalOrgParser;