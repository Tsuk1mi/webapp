// Специализированный парсер для корпоративных документов
class CorporateDocumentParser {
    constructor() {
        this.debug = true;
    }

    // Главный метод для парсинга документов
    async parseDocument(content, fileType) {
        console.log('Parsing document with CorporateDocumentParser...');
        
        if (fileType.includes('spreadsheet') || content.includes('Подчиненность')) {
            return this.parseTableStructure(content);
        } else if (fileType.includes('presentation')) {
            return this.parsePresentationStructure(content);
        } else if (fileType.includes('pdf')) {
            return this.parsePDFStructure(content);
        } else if (fileType.includes('word')) {
            return this.parseWordStructure(content);
        }
        
        return this.parseGenericStructure(content);
    }

    // Парсинг табличной структуры (DOCX с таблицами)
    parseWordStructure(content) {
        console.log('Parsing Word document structure...');
        
        const employees = [];
        let currentDepartment = '';
        let currentBoss = null;
        
        // Паттерны для поиска
        const patterns = {
            boss: /начальник\s+(отдела|управления|департамента)/i,
            deputy: /заместитель\s+начальника/i,
            lead: /руководитель|ведущий/i,
            fio: /([А-ЯЁ][а-яё]+)\s+([А-ЯЁ][а-яё]+)\s+([А-ЯЁ][а-яё]+)/g,
            department: /(отдел|управление|департамент|направление)\s+[А-ЯЁа-яё\s]+/gi
        };
        
        // Поиск отдела в заголовке
        const deptMatch = content.match(/структура\s+(.*?)[\n<]/i);
        if (deptMatch) {
            currentDepartment = deptMatch[1].trim();
        }
        
        // Разбор строк
        const lines = content.split(/[\n\r]+/);
        
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            
            // Пропускаем пустые строки
            if (!line.trim()) continue;
            
            // Ищем ФИО
            const fioMatches = [...line.matchAll(patterns.fio)];
            
            for (const match of fioMatches) {
                const fullName = match[0];
                const position = this.extractPosition(line, fullName);
                const functional = this.extractFunctional(lines, i);
                
                const employee = {
                    id: `emp-${employees.length + 1}`,
                    name: fullName,
                    position: position || 'Сотрудник',
                    department: currentDepartment,
                    functional: functional,
                    level: this.determineLevel(position),
                    children: []
                };
                
                // Определяем начальника отдела
                if (patterns.boss.test(position) && !patterns.deputy.test(position)) {
                    currentBoss = employee;
                    employee.level = 1;
                }
                
                employees.push(employee);
            }
        }
        
        return this.buildHierarchyFromEmployees(employees, currentDepartment);
    }

    // Парсинг табличной структуры (Google Sheets, Excel)
    parseTableStructure(content) {
        console.log('Parsing table structure...');
        
        // Парсим CSV с учетом кавычек
        const rows = this.parseCSV(content);
        if (rows.length < 2) return null;
        
        const headers = rows[0];
        const employees = [];
        const nodeMap = new Map();
        
        // Определяем индексы колонок
        const colIndexes = {
            subordination: this.findColumnIndex(headers, ['подчиненность', 'код', 'subordination']),
            department: this.findColumnIndex(headers, ['подразделение', 'отдел', 'department']),
            position: this.findColumnIndex(headers, ['должность', 'position', 'title']),
            name: this.findColumnIndex(headers, ['фио', 'сотрудник', 'name', 'ф.и.о']),
            functional: this.findColumnIndex(headers, ['функционал', 'обязанности', 'functional'])
        };
        
        console.log('Column indexes:', colIndexes);
        
        // Обрабатываем каждую строку данных
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;
            
            const subordination = colIndexes.subordination >= 0 ? row[colIndexes.subordination] : '';
            const name = colIndexes.name >= 0 ? row[colIndexes.name] : '';
            
            // Пропускаем строки без имени
            if (!name || !name.trim()) continue;
            
            const employee = {
                id: `node-${i}`,
                subordination: subordination || '',
                department: colIndexes.department >= 0 ? row[colIndexes.department] : '',
                position: colIndexes.position >= 0 ? row[colIndexes.position] : '',
                name: name.trim(),
                functional: colIndexes.functional >= 0 ? row[colIndexes.functional] : '',
                level: 1,
                children: []
            };
            
            // Определяем уровень по коду подчиненности
            if (subordination) {
                const dots = (subordination.match(/\./g) || []).length;
                employee.level = dots + 1;
            }
            
            employees.push(employee);
            nodeMap.set(subordination, employee);
        }
        
        console.log(`Parsed ${employees.length} employees from table`);
        
        // Строим иерархию
        return this.buildHierarchyFromSubordination(employees, nodeMap);
    }

    // Парсинг презентации (PPTX)
    parsePresentationStructure(content) {
        console.log('Parsing presentation structure...');
        
        const employees = [];
        const slides = content.split(/<page[^>]*>/);
        
        for (const slide of slides) {
            // Извлекаем текст слайда
            const text = slide.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ');
            
            // Ищем ФИО и должности
            const lines = text.split(/##|[\n\r]+/);
            
            for (let i = 0; i < lines.length; i++) {
                const line = lines[i].trim();
                if (!line) continue;
                
                // Паттерн для ФИО
                const fioPattern = /([А-ЯЁ][а-яё]+)\s+([А-ЯЁ][а-яё]+)\s+([А-ЯЁ][а-яё]+)/;
                const fioMatch = line.match(fioPattern);
                
                if (fioMatch) {
                    const name = fioMatch[0];
                    let position = '';
                    
                    // Ищем должность в соседних строках
                    for (let j = Math.max(0, i - 2); j <= Math.min(lines.length - 1, i + 2); j++) {
                        if (j !== i && lines[j]) {
                            const posLine = lines[j].trim();
                            if (this.isPosition(posLine) && !fioPattern.test(posLine)) {
                                position = posLine;
                                break;
                            }
                        }
                    }
                    
                    employees.push({
                        id: `emp-${employees.length + 1}`,
                        name: name,
                        position: position || 'Сотрудник',
                        department: this.extractDepartmentFromContext(lines, i),
                        level: this.determineLevel(position),
                        children: []
                    });
                }
            }
        }
        
        console.log(`Found ${employees.length} employees in presentation`);
        return this.buildHierarchyFromEmployees(employees);
    }

    // Парсинг PDF структуры
    parsePDFStructure(content) {
        console.log('Parsing PDF structure...');
        
        // PDF парсинг похож на презентацию
        return this.parsePresentationStructure(content);
    }

    // Общий парсинг для неизвестных форматов
    parseGenericStructure(content) {
        console.log('Parsing generic structure...');
        
        const employees = [];
        const lines = content.split(/[\n\r]+/);
        
        for (const line of lines) {
            const fioPattern = /([А-ЯЁ][а-яё]+)\s+([А-ЯЁ][а-яё]+)\s+([А-ЯЁ][а-яё]+)/g;
            const matches = [...line.matchAll(fioPattern)];
            
            for (const match of matches) {
                employees.push({
                    id: `emp-${employees.length + 1}`,
                    name: match[0],
                    position: this.extractPosition(line, match[0]),
                    department: '',
                    level: 5,
                    children: []
                });
            }
        }
        
        return this.buildHierarchyFromEmployees(employees);
    }

    // Парсинг CSV с учетом кавычек
    parseCSV(text) {
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

    // Поиск индекса колонки по возможным названиям
    findColumnIndex(headers, possibleNames) {
        for (let i = 0; i < headers.length; i++) {
            const header = headers[i].toLowerCase();
            for (const name of possibleNames) {
                if (header.includes(name.toLowerCase())) {
                    return i;
                }
            }
        }
        return -1;
    }

    // Построение иерархии из кодов подчиненности
    buildHierarchyFromSubordination(employees, nodeMap) {
        const roots = [];
        
        for (const employee of employees) {
            const code = employee.subordination;
            
            if (!code || code === '1' || !code.includes('.')) {
                // Корневой элемент
                roots.push(employee);
            } else {
                // Ищем родителя
                const parts = code.split('.');
                let parentCode = parts.slice(0, -1).join('.');
                
                // Пробуем найти родителя
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
        
        // Если есть один корень, возвращаем его
        if (roots.length === 1) {
            return roots[0];
        }
        
        // Если несколько корней, создаем общий корень
        return {
            id: 'root',
            name: 'Организация',
            position: '',
            department: '',
            children: roots
        };
    }

    // Построение иерархии из списка сотрудников
    buildHierarchyFromEmployees(employees, departmentName = '') {
        if (employees.length === 0) return null;
        
        // Сортируем по уровню
        employees.sort((a, b) => a.level - b.level);
        
        // Группируем по уровням
        const levels = {};
        for (const emp of employees) {
            if (!levels[emp.level]) {
                levels[emp.level] = [];
            }
            levels[emp.level].push(emp);
        }
        
        // Строим иерархию
        const levelKeys = Object.keys(levels).map(Number).sort((a, b) => a - b);
        
        // Распределяем сотрудников по уровням
        for (let i = 0; i < levelKeys.length - 1; i++) {
            const currentLevel = levels[levelKeys[i]];
            const nextLevel = levels[levelKeys[i + 1]];
            
            if (currentLevel && nextLevel) {
                // Распределяем сотрудников следующего уровня между текущим
                const childrenPerParent = Math.ceil(nextLevel.length / currentLevel.length);
                let childIndex = 0;
                
                for (const parent of currentLevel) {
                    for (let j = 0; j < childrenPerParent && childIndex < nextLevel.length; j++) {
                        parent.children.push(nextLevel[childIndex]);
                        childIndex++;
                    }
                }
            }
        }
        
        // Возвращаем корневые элементы
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
    extractPosition(line, name) {
        // Удаляем имя из строки
        const withoutName = line.replace(name, '').trim();
        
        // Паттерны должностей
        const positionPatterns = [
            /начальник\s+[а-яё\s]+/gi,
            /заместитель\s+[а-яё\s]+/gi,
            /руководитель\s+[а-яё\s]+/gi,
            /директор\s*[а-яё\s]*/gi,
            /менеджер\s*[а-яё\s]*/gi,
            /специалист\s*[а-яё\s]*/gi,
            /ведущий\s+[а-яё\s]+/gi,
            /главный\s+[а-яё\s]+/gi,
            /старший\s+[а-яё\s]+/gi,
            /инженер\s*[а-яё\s]*/gi,
            /аналитик\s*[а-яё\s]*/gi,
            /консультант\s*[а-яё\s]*/gi,
            /координатор\s*[а-яё\s]*/gi,
            /администратор\s*[а-яё\s]*/gi,
            /бухгалтер\s*[а-яё\s]*/gi
        ];
        
        for (const pattern of positionPatterns) {
            const match = withoutName.match(pattern);
            if (match) {
                return match[0].trim();
            }
        }
        
        // Если находим тире или двоеточие, берем текст после них
        if (withoutName.includes('-')) {
            const parts = withoutName.split('-');
            if (parts[1]) return parts[1].trim().split(/[,;.]/)[0].trim();
        }
        
        if (withoutName.includes(':')) {
            const parts = withoutName.split(':');
            if (parts[0] && parts[0].length < 100) return parts[0].trim();
        }
        
        return '';
    }

    // Извлечение функционала
    extractFunctional(lines, currentIndex) {
        // Ищем функционал в следующих строках
        for (let i = currentIndex + 1; i < Math.min(currentIndex + 5, lines.length); i++) {
            const line = lines[i];
            if (line && line.length > 50 && !this.isFIO(line)) {
                return line.trim();
            }
        }
        return '';
    }

    // Проверка, является ли строка ФИО
    isFIO(text) {
        const fioPattern = /^([А-ЯЁ][а-яё]+)\s+([А-ЯЁ][а-яё]+)\s+([А-ЯЁ][а-яё]+)$/;
        return fioPattern.test(text.trim());
    }

    // Проверка, является ли строка должностью
    isPosition(text) {
        const positionKeywords = [
            'начальник', 'заместитель', 'руководитель', 'директор',
            'менеджер', 'специалист', 'ведущий', 'главный', 'старший',
            'инженер', 'аналитик', 'консультант', 'координатор',
            'администратор', 'бухгалтер', 'юрист', 'экономист'
        ];
        
        const lower = text.toLowerCase();
        return positionKeywords.some(keyword => lower.includes(keyword));
    }

    // Извлечение департамента из контекста
    extractDepartmentFromContext(lines, currentIndex) {
        const deptPatterns = [
            /отдел\s+[а-яё\s]+/gi,
            /управление\s+[а-яё\s]+/gi,
            /департамент\s+[а-яё\s]+/gi,
            /направление\s+[а-яё\s]+/gi
        ];
        
        // Ищем в радиусе 5 строк
        for (let i = Math.max(0, currentIndex - 5); i <= Math.min(lines.length - 1, currentIndex + 5); i++) {
            const line = lines[i];
            if (!line) continue;
            
            for (const pattern of deptPatterns) {
                const match = line.match(pattern);
                if (match) {
                    return match[0].trim();
                }
            }
        }
        
        return '';
    }

    // Определение уровня по должности
    determineLevel(position) {
        const lower = position.toLowerCase();
        
        if (lower.includes('генеральный') || lower.includes('президент')) {
            return 1;
        } else if (lower.includes('заместитель') && lower.includes('генерального')) {
            return 2;
        } else if (lower.includes('начальник управления') || lower.includes('директор департамента')) {
            return 2;
        } else if (lower.includes('заместитель начальника управления')) {
            return 3;
        } else if (lower.includes('начальник отдела')) {
            return 3;
        } else if (lower.includes('заместитель начальника отдела')) {
            return 4;
        } else if (lower.includes('руководитель') && !lower.includes('заместитель')) {
            return 4;
        } else if (lower.includes('главный') || lower.includes('ведущий')) {
            return 5;
        } else if (lower.includes('старший')) {
            return 6;
        } else {
            return 7;
        }
    }

    // Создание детальной таблицы
    createDetailedTable(structure) {
        const rows = [];
        this.flattenStructure(structure, rows, 0, '');
        
        let html = `
            <h3 class="text-xl font-semibold mb-4">
                <i class="fas fa-table mr-2"></i>Организационная структура (${rows.length} сотрудников)
            </h3>
            <div class="overflow-x-auto">
                <table class="min-w-full divide-y divide-gray-200">
                    <thead class="bg-gray-50">
                        <tr>
                            <th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">№</th>
                            <th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Код</th>
                            <th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Подразделение</th>
                            <th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Должность</th>
                            <th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">ФИО</th>
                            <th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Уровень</th>
                        </tr>
                    </thead>
                    <tbody class="bg-white divide-y divide-gray-200">
        `;
        
        rows.forEach((row, index) => {
            html += `
                <tr class="hover:bg-gray-50">
                    <td class="px-4 py-2 text-sm">${index + 1}</td>
                    <td class="px-4 py-2 text-sm font-mono">${row.code}</td>
                    <td class="px-4 py-2 text-sm">${row.department || '-'}</td>
                    <td class="px-4 py-2 text-sm">${row.position || '-'}</td>
                    <td class="px-4 py-2 text-sm font-medium">${row.name}</td>
                    <td class="px-4 py-2 text-sm text-center">${row.level}</td>
                </tr>
            `;
        });
        
        html += `
                    </tbody>
                </table>
            </div>
        `;
        
        return html;
    }

    // Преобразование иерархии в плоский список
    flattenStructure(node, rows, level, parentCode) {
        if (!node) return;
        
        const code = parentCode ? `${parentCode}.${rows.length + 1}` : '1';
        
        if (node.name && node.name !== 'Организация') {
            rows.push({
                code: node.subordination || code,
                name: node.name,
                position: node.position || '',
                department: node.department || '',
                level: level + 1
            });
        }
        
        if (node.children && node.children.length > 0) {
            for (let i = 0; i < node.children.length; i++) {
                this.flattenStructure(
                    node.children[i], 
                    rows, 
                    node.name === 'Организация' ? level : level + 1,
                    node.name === 'Организация' ? '' : code
                );
            }
        }
    }
}

// Экспортируем класс в глобальную область
window.CorporateDocumentParser = CorporateDocumentParser;