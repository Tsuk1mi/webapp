// Advanced document parser for organizational structures
class AdvancedOrgParser {
    constructor() {
        this.patterns = {
            // Паттерны для поиска должностей
            positions: [
                /начальник\s+(отдела|управления|департамента)/gi,
                /заместитель\s+начальника/gi,
                /руководитель\s+(проектов|направления|отдела)/gi,
                /директор/gi,
                /менеджер/gi,
                /специалист/gi,
                /главный\s+\w+/gi,
                /ведущий\s+\w+/gi,
                /старший\s+\w+/gi,
                /технический\s+руководитель/gi,
                /координатор/gi
            ],
            // Паттерны для имен
            names: /([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)|([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)/g,
            // Паттерны для отделов
            departments: /отдел\s+[а-яё\s]+|управление\s+[а-яё\s]+|департамент\s+[а-яё\s]+|направление\s+[а-яё\s]+/gi,
            // Email паттерн
            emails: /[\w.-]+@[\w.-]+\.\w+/g
        };
    }

    // Главный метод парсинга
    async parseDocument(text, type) {
        console.log('Parsing document type:', type);
        
        // Очистка текста
        text = this.cleanText(text);
        
        // Извлечение структурированных данных
        const employees = this.extractEmployees(text);
        const departments = this.extractDepartments(text);
        
        // Построение иерархии
        const hierarchy = this.buildHierarchy(employees, departments, text);
        
        // Оптимизация структуры
        return this.optimizeStructure(hierarchy);
    }

    // Очистка текста
    cleanText(text) {
        return text
            .replace(/\s+/g, ' ')
            .replace(/\n{3,}/g, '\n\n')
            .replace(/[•·]/g, '-')
            .trim();
    }

    // Извлечение сотрудников
    extractEmployees(text) {
        const employees = [];
        const lines = text.split(/\n|<\/td>|<\/tr>|\|/);
        
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();
            if (!line) continue;
            
            // Поиск имен
            const nameMatches = line.match(this.patterns.names);
            if (nameMatches) {
                for (const name of nameMatches) {
                    // Проверяем, что это полное имя (минимум 2 слова)
                    const nameParts = name.trim().split(/\s+/);
                    if (nameParts.length >= 2 && nameParts[0].length > 1) {
                        const employee = {
                            name: name.trim(),
                            originalLine: line,
                            lineIndex: i
                        };
                        
                        // Ищем должность
                        employee.position = this.findPosition(line, lines, i);
                        
                        // Ищем отдел
                        employee.department = this.findDepartment(line, lines, i);
                        
                        // Ищем email
                        const emailMatch = line.match(this.patterns.emails);
                        if (emailMatch) {
                            employee.email = emailMatch[0];
                        }
                        
                        // Определяем уровень в иерархии
                        employee.level = this.determineLevel(employee.position);
                        
                        // Ищем функции/обязанности
                        employee.responsibilities = this.findResponsibilities(lines, i);
                        
                        employees.push(employee);
                    }
                }
            }
        }
        
        return employees;
    }

    // Поиск должности для сотрудника
    findPosition(line, lines, index) {
        // Сначала ищем в той же строке
        for (const pattern of this.patterns.positions) {
            const match = line.match(pattern);
            if (match) {
                return match[0].trim();
            }
        }
        
        // Затем в соседних строках
        for (let offset = -2; offset <= 2; offset++) {
            if (offset === 0) continue;
            const checkIndex = index + offset;
            if (checkIndex >= 0 && checkIndex < lines.length) {
                const checkLine = lines[checkIndex];
                for (const pattern of this.patterns.positions) {
                    const match = checkLine.match(pattern);
                    if (match) {
                        return match[0].trim();
                    }
                }
            }
        }
        
        // Простой поиск ключевых слов
        const positionKeywords = ['начальник', 'заместитель', 'руководитель', 'директор', 
                                 'менеджер', 'специалист', 'координатор', 'аналитик', 
                                 'разработчик', 'инженер', 'консультант'];
        
        const lowerLine = line.toLowerCase();
        for (const keyword of positionKeywords) {
            if (lowerLine.includes(keyword)) {
                // Извлекаем контекст вокруг ключевого слова
                const keywordIndex = lowerLine.indexOf(keyword);
                const start = Math.max(0, keywordIndex - 20);
                const end = Math.min(line.length, keywordIndex + keyword.length + 30);
                return line.substring(start, end).trim();
            }
        }
        
        return null;
    }

    // Поиск отдела
    findDepartment(line, lines, index) {
        // Поиск в текущей и соседних строках
        for (let offset = -3; offset <= 3; offset++) {
            const checkIndex = index + offset;
            if (checkIndex >= 0 && checkIndex < lines.length) {
                const checkLine = lines[checkIndex];
                const match = checkLine.match(this.patterns.departments);
                if (match) {
                    return match[0].trim();
                }
                
                // Дополнительные паттерны
                if (checkLine.match(/ОТДЕЛ\s+[А-ЯЁ\s]+/)) {
                    const deptMatch = checkLine.match(/ОТДЕЛ\s+[А-ЯЁ\s]+/);
                    if (deptMatch) {
                        return deptMatch[0].trim();
                    }
                }
            }
        }
        
        return null;
    }

    // Поиск обязанностей
    findResponsibilities(lines, index) {
        const responsibilities = [];
        
        // Ищем в следующих строках списки обязанностей
        for (let i = index + 1; i < Math.min(index + 10, lines.length); i++) {
            const line = lines[i].trim();
            
            // Прекращаем, если встретили нового сотрудника
            if (this.patterns.names.test(line)) {
                break;
            }
            
            // Ищем маркеры списка
            if (line.match(/^[-•·*]\s*/) || line.match(/^\d+\.\s*/)) {
                responsibilities.push(line.replace(/^[-•·*]\s*/, '').replace(/^\d+\.\s*/, '').trim());
            }
        }
        
        return responsibilities;
    }

    // Определение уровня в иерархии
    determineLevel(position) {
        if (!position) return 5;
        
        const posLower = position.toLowerCase();
        
        if (posLower.includes('генеральный') || posLower.includes('президент')) {
            return 1;
        }
        if (posLower.includes('заместитель') && posLower.includes('начальника управления')) {
            return 2;
        }
        if (posLower.includes('начальник управления') || posLower.includes('директор')) {
            return 2;
        }
        if (posLower.includes('начальник отдела')) {
            return 3;
        }
        if (posLower.includes('заместитель начальника отдела')) {
            return 4;
        }
        if (posLower.includes('руководитель')) {
            return 3;
        }
        if (posLower.includes('главный')) {
            return 4;
        }
        if (posLower.includes('ведущий') || posLower.includes('старший')) {
            return 5;
        }
        
        return 6;
    }

    // Извлечение отделов
    extractDepartments(text) {
        const departments = new Map();
        
        // Ищем явные упоминания отделов
        const deptMatches = text.match(/ОТДЕЛ\s+[А-ЯЁ\s]+|отдел\s+[а-яё\s]+/gi);
        if (deptMatches) {
            deptMatches.forEach(dept => {
                const cleanDept = dept.trim();
                if (!departments.has(cleanDept)) {
                    departments.set(cleanDept, {
                        name: cleanDept,
                        employees: [],
                        subdepartments: []
                    });
                }
            });
        }
        
        return departments;
    }

    // Построение иерархии
    buildHierarchy(employees, departments, text) {
        // Сортируем сотрудников по уровню
        employees.sort((a, b) => a.level - b.level);
        
        // Создаем корневой узел
        const root = {
            id: 'root',
            name: 'Организация',
            title: '',
            children: [],
            level: 0
        };
        
        // Группируем по отделам
        const departmentGroups = new Map();
        
        employees.forEach(emp => {
            const dept = emp.department || 'Без отдела';
            if (!departmentGroups.has(dept)) {
                departmentGroups.set(dept, []);
            }
            departmentGroups.get(dept).push(emp);
        });
        
        // Строим иерархию для каждого отдела
        departmentGroups.forEach((deptEmployees, deptName) => {
            if (deptName !== 'Без отдела') {
                const deptNode = {
                    id: `dept-${Math.random().toString(36).substr(2, 9)}`,
                    name: deptName,
                    title: 'Отдел',
                    isDepartment: true,
                    children: [],
                    level: 1
                };
                
                // Добавляем сотрудников в отдел
                const hierarchy = this.buildEmployeeHierarchy(deptEmployees);
                deptNode.children = hierarchy;
                
                root.children.push(deptNode);
            } else {
                // Сотрудники без отдела добавляются напрямую
                const hierarchy = this.buildEmployeeHierarchy(deptEmployees);
                root.children.push(...hierarchy);
            }
        });
        
        return root;
    }

    // Построение иерархии сотрудников
    buildEmployeeHierarchy(employees) {
        const nodes = [];
        const nodeMap = new Map();
        
        // Создаем узлы для всех сотрудников
        employees.forEach(emp => {
            const node = {
                id: `emp-${Math.random().toString(36).substr(2, 9)}`,
                name: emp.name,
                title: emp.position || 'Сотрудник',
                email: emp.email,
                responsibilities: emp.responsibilities,
                level: emp.level,
                children: []
            };
            
            nodeMap.set(emp.name, node);
            nodes.push(node);
        });
        
        // Строим иерархию на основе уровней
        const hierarchy = [];
        const levelGroups = new Map();
        
        nodes.forEach(node => {
            if (!levelGroups.has(node.level)) {
                levelGroups.set(node.level, []);
            }
            levelGroups.get(node.level).push(node);
        });
        
        // Связываем узлы по уровням
        const sortedLevels = Array.from(levelGroups.keys()).sort((a, b) => a - b);
        
        for (let i = 0; i < sortedLevels.length; i++) {
            const currentLevel = sortedLevels[i];
            const currentNodes = levelGroups.get(currentLevel);
            
            if (i === 0) {
                // Верхний уровень
                hierarchy.push(...currentNodes);
            } else {
                // Подчиненные уровни
                const parentLevel = sortedLevels[i - 1];
                const parentNodes = levelGroups.get(parentLevel);
                
                // Распределяем равномерно
                currentNodes.forEach((node, index) => {
                    const parentIndex = Math.floor(index * parentNodes.length / currentNodes.length);
                    const parent = parentNodes[Math.min(parentIndex, parentNodes.length - 1)];
                    parent.children.push(node);
                });
            }
        }
        
        return hierarchy;
    }

    // Оптимизация структуры
    optimizeStructure(hierarchy) {
        // Удаляем пустые узлы
        this.removeEmptyNodes(hierarchy);
        
        // Объединяем одинаковые отделы
        this.mergeDuplicateDepartments(hierarchy);
        
        // Упрощаем структуру
        if (hierarchy.children.length === 1 && hierarchy.name === 'Организация') {
            return hierarchy.children[0];
        }
        
        return hierarchy;
    }

    // Удаление пустых узлов
    removeEmptyNodes(node) {
        if (!node.children || node.children.length === 0) {
            return;
        }
        
        node.children = node.children.filter(child => {
            if (!child.name || child.name.trim() === '') {
                return false;
            }
            this.removeEmptyNodes(child);
            return true;
        });
    }

    // Объединение дублирующихся отделов
    mergeDuplicateDepartments(node) {
        if (!node.children || node.children.length === 0) {
            return;
        }
        
        const departmentMap = new Map();
        const newChildren = [];
        
        node.children.forEach(child => {
            if (child.isDepartment) {
                const key = child.name.toLowerCase().trim();
                if (departmentMap.has(key)) {
                    // Объединяем с существующим
                    const existing = departmentMap.get(key);
                    existing.children.push(...child.children);
                } else {
                    departmentMap.set(key, child);
                    newChildren.push(child);
                }
            } else {
                newChildren.push(child);
            }
            
            this.mergeDuplicateDepartments(child);
        });
        
        node.children = newChildren;
    }

    // Преобразование в формат для таблицы
    toTableFormat(hierarchy) {
        const rows = [];
        
        const traverse = (node, level = 0, parent = '') => {
            rows.push({
                level: level,
                name: node.name,
                title: node.title || '',
                department: node.isDepartment ? node.name : parent,
                email: node.email || '',
                responsibilities: node.responsibilities ? node.responsibilities.join('; ') : '',
                parent: parent
            });
            
            if (node.children) {
                node.children.forEach(child => {
                    traverse(child, level + 1, node.name);
                });
            }
        };
        
        traverse(hierarchy);
        return rows;
    }
}

// Экспорт для использования
window.AdvancedOrgParser = AdvancedOrgParser;