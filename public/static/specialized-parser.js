// Specialized parser for organizational documents
class SpecializedOrgParser {
    constructor() {
        this.hierarchyMap = new Map();
        this.departmentStructure = new Map();
        this.employeeDatabase = [];
    }

    // Parse Google Sheets structure
    parseGoogleSheetsStructure(csvData) {
        const lines = csvData.split('\n').filter(line => line.trim());
        if (lines.length === 0) return null;

        const rows = [];
        lines.forEach(line => {
            const cols = this.parseCSVLine(line);
            rows.push(cols);
        });

        // Assume first row is headers
        const headers = rows[0];
        const data = [];

        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            const entry = {};
            headers.forEach((header, index) => {
                entry[header] = row[index] || '';
            });
            data.push(entry);
        }

        return this.buildHierarchyFromTable(data);
    }

    // Parse CSV line properly handling quotes
    parseCSVLine(line) {
        const result = [];
        let current = '';
        let inQuotes = false;
        
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            
            if (char === '"') {
                inQuotes = !inQuotes;
            } else if (char === ',' && !inQuotes) {
                result.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        
        if (current) {
            result.push(current.trim());
        }
        
        return result;
    }

    // Build hierarchy from table data
    buildHierarchyFromTable(data) {
        const nodeMap = new Map();
        const rootNodes = [];

        // First pass: create all nodes
        data.forEach((row, index) => {
            const node = {
                id: `node-${index}`,
                subordination: row['Подчиненность'] || row['Subordination'] || '',
                department: row['Подразделение'] || row['Department'] || '',
                position: row['Должность'] || row['Position'] || '',
                name: row['ФИО'] || row['Name'] || '',
                functional: row['Функционал'] || row['Functional'] || '',
                children: [],
                level: this.determineHierarchyLevel(row['Должность'] || row['Position'] || ''),
                responsibilities: []
            };

            // Parse responsibilities if they exist in functional field
            if (node.functional && node.functional.length > 20) {
                node.responsibilities = this.parseResponsibilities(node.functional);
            }

            nodeMap.set(node.subordination, node);
        });

        // Second pass: build hierarchy based on subordination codes
        nodeMap.forEach((node, code) => {
            if (code === '00' || code === '0') {
                // Root level
                rootNodes.push(node);
            } else {
                // Find parent based on subordination pattern
                let parentCode = this.findParentCode(code, nodeMap);
                if (parentCode && nodeMap.has(parentCode)) {
                    nodeMap.get(parentCode).children.push(node);
                } else {
                    // If no parent found, add to root
                    rootNodes.push(node);
                }
            }
        });

        // Create organizational root
        const root = {
            id: 'org-root',
            name: 'Организация',
            position: '',
            department: '',
            children: rootNodes,
            level: 0
        };

        return this.optimizeHierarchy(root);
    }

    // Find parent code based on subordination pattern
    findParentCode(code, nodeMap) {
        // Try to find direct parent
        const numCode = parseInt(code);
        
        // Check for simple parent relationships
        if (code.length > 1) {
            // Try removing last digit
            const parentCandidate = code.slice(0, -1);
            if (nodeMap.has(parentCandidate)) {
                return parentCandidate;
            }
        }
        
        // Check for numerical parent
        if (numCode > 10) {
            // Try common parent patterns
            const possibleParents = [
                Math.floor(numCode / 10).toString(),
                '0',
                '00',
                '1'
            ];
            
            for (const parent of possibleParents) {
                if (nodeMap.has(parent)) {
                    return parent;
                }
            }
        }
        
        // Default parent codes
        if (numCode >= 1 && numCode <= 20) {
            return '00'; // Report to top management
        }
        
        return null;
    }

    // Parse document with complex structure
    parseComplexDocument(text, fileName) {
        const structure = {
            departments: [],
            employees: [],
            hierarchy: null
        };

        // Detect document type
        const docType = this.detectDocumentType(text, fileName);
        
        // Parse based on document type
        switch (docType) {
            case 'table_structure':
                structure.employees = this.parseTableStructure(text);
                break;
            case 'department_list':
                structure.departments = this.parseDepartmentList(text);
                structure.employees = this.extractEmployeesFromDepartments(text);
                break;
            case 'presentation':
                structure.hierarchy = this.parsePresentationStructure(text);
                break;
            default:
                // Generic parsing
                structure.employees = this.parseGenericDocument(text);
        }

        // Build final hierarchy
        return this.buildFinalHierarchy(structure);
    }

    // Detect document type
    detectDocumentType(text, fileName) {
        const lowerText = text.toLowerCase();
        
        if (fileName.includes('pptx') || lowerText.includes('слайд')) {
            return 'presentation';
        }
        if (lowerText.includes('структура отдела') || lowerText.includes('№ п/п')) {
            return 'table_structure';
        }
        if (lowerText.includes('отдел') && lowerText.includes('начальник отдела')) {
            return 'department_list';
        }
        
        return 'generic';
    }

    // Parse table structure (like the first DOCX)
    parseTableStructure(text) {
        const employees = [];
        const lines = text.split(/\n|\|<\/td>|<\/tr>/);
        
        let currentEmployee = null;
        let collectingResponsibilities = false;
        
        lines.forEach(line => {
            const cleanLine = line.replace(/<[^>]*>/g, '').trim();
            if (!cleanLine) return;
            
            // Check for numbered entries (1., 2., etc.)
            const numberMatch = cleanLine.match(/^(\d+)\.\s*(.+)/);
            if (numberMatch) {
                // Save previous employee
                if (currentEmployee) {
                    employees.push(currentEmployee);
                }
                
                // Start new employee
                const content = numberMatch[2];
                const parts = this.parseEmployeeEntry(content);
                
                currentEmployee = {
                    id: `emp-${numberMatch[1]}`,
                    number: numberMatch[1],
                    name: parts.name,
                    position: parts.position,
                    department: parts.department,
                    responsibilities: parts.responsibilities,
                    deputy: parts.deputy,
                    level: this.determineHierarchyLevel(parts.position)
                };
                
                collectingResponsibilities = true;
            } else if (collectingResponsibilities && currentEmployee) {
                // Add to responsibilities if it looks like a continuation
                if (!cleanLine.match(/^\d+$/) && cleanLine.length > 10) {
                    currentEmployee.responsibilities.push(cleanLine);
                }
            }
        });
        
        // Don't forget the last employee
        if (currentEmployee) {
            employees.push(currentEmployee);
        }
        
        return employees;
    }

    // Parse employee entry
    parseEmployeeEntry(text) {
        const result = {
            name: '',
            position: '',
            department: '',
            deputy: '',
            responsibilities: []
        };
        
        // Extract name and position
        const namePositionMatch = text.match(/([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)\s*[-–]\s*([^:]+):/);
        if (namePositionMatch) {
            result.name = namePositionMatch[1].trim();
            result.position = namePositionMatch[2].trim();
            
            // Extract deputy if mentioned
            const deputyMatch = text.match(/\(Заместитель[^)]+–\s*([^)]+)\)/);
            if (deputyMatch) {
                result.deputy = deputyMatch[1].trim();
            }
            
            // Extract responsibilities
            const respPart = text.split(':').slice(1).join(':');
            if (respPart) {
                const responsibilities = respPart.split(/\.\s+/).filter(r => r.trim().length > 5);
                result.responsibilities = responsibilities.map(r => r.trim());
            }
        } else {
            // Fallback parsing
            const parts = text.split('-').map(p => p.trim());
            if (parts.length >= 2) {
                result.name = parts[0];
                const positionAndResp = parts.slice(1).join('-');
                const colonIndex = positionAndResp.indexOf(':');
                if (colonIndex > 0) {
                    result.position = positionAndResp.substring(0, colonIndex).trim();
                    result.responsibilities = [positionAndResp.substring(colonIndex + 1).trim()];
                } else {
                    result.position = positionAndResp;
                }
            }
        }
        
        return result;
    }

    // Parse department list
    parseDepartmentList(text) {
        const departments = [];
        const deptPattern = /ОТДЕЛ\s+[А-ЯЁ\s]+|Отдел\s+[а-яё\s]+|Управление\s+[а-яё\s]+/gi;
        const matches = text.match(deptPattern);
        
        if (matches) {
            matches.forEach(dept => {
                departments.push({
                    name: dept.trim(),
                    employees: []
                });
            });
        }
        
        return departments;
    }

    // Extract employees from departments
    extractEmployeesFromDepartments(text) {
        const employees = [];
        const namePattern = /([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)/g;
        const positionPattern = /(Начальник\s+\w+|Заместитель\s+\w+|Руководитель\s+\w+|Директор|Специалист|Менеджер|Координатор|Аналитик|Геоаналитик|Нормоконтролер)/gi;
        
        const lines = text.split('\n');
        let currentDepartment = '';
        
        lines.forEach((line, index) => {
            // Check for department
            const deptMatch = line.match(/ОТДЕЛ\s+[А-ЯЁ\s]+|Отдел\s+[а-яё\s]+/i);
            if (deptMatch) {
                currentDepartment = deptMatch[0].trim();
            }
            
            // Extract names
            const names = line.match(namePattern);
            if (names) {
                names.forEach(name => {
                    // Look for position in the same or nearby lines
                    let position = '';
                    for (let offset = -1; offset <= 1; offset++) {
                        const checkLine = lines[index + offset] || '';
                        const posMatch = checkLine.match(positionPattern);
                        if (posMatch) {
                            position = posMatch[0];
                            break;
                        }
                    }
                    
                    employees.push({
                        name: name.trim(),
                        position: position,
                        department: currentDepartment,
                        level: this.determineHierarchyLevel(position)
                    });
                });
            }
        });
        
        return employees;
    }

    // Parse presentation structure
    parsePresentationStructure(text) {
        const nodes = [];
        const lines = text.split('\n').filter(line => line.trim());
        
        lines.forEach(line => {
            const cleanLine = line.replace(/##\s*/, '').trim();
            
            // Parse entries like "Руководитель проектовВафина Рушания Рустамовна"
            const matches = cleanLine.match(/([А-Я][а-я]+(?:\s+[а-я]+)*)\s*([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ][а-яё]+)?)/);
            if (matches) {
                nodes.push({
                    position: matches[1],
                    name: matches[2],
                    level: this.determineHierarchyLevel(matches[1])
                });
            }
        });
        
        return this.buildHierarchyFromNodes(nodes);
    }

    // Parse generic document
    parseGenericDocument(text) {
        const employees = [];
        const lines = text.split('\n');
        
        lines.forEach((line, index) => {
            const nameMatch = line.match(/([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)/);
            if (nameMatch) {
                const employee = {
                    name: nameMatch[1],
                    position: '',
                    department: '',
                    responsibilities: []
                };
                
                // Look for position
                for (let offset = -2; offset <= 2; offset++) {
                    const checkLine = lines[index + offset] || '';
                    const posMatch = checkLine.match(/(начальник|заместитель|руководитель|директор|специалист|менеджер|координатор|аналитик)/i);
                    if (posMatch) {
                        // Get more context around the position keyword
                        const start = Math.max(0, checkLine.indexOf(posMatch[0]) - 20);
                        const end = Math.min(checkLine.length, start + 100);
                        employee.position = checkLine.substring(start, end).trim();
                        break;
                    }
                }
                
                employee.level = this.determineHierarchyLevel(employee.position);
                employees.push(employee);
            }
        });
        
        return employees;
    }

    // Parse responsibilities
    parseResponsibilities(text) {
        const responsibilities = [];
        
        // Split by common delimiters
        const parts = text.split(/[.;•·]/);
        
        parts.forEach(part => {
            const cleaned = part.trim();
            if (cleaned.length > 10 && !cleaned.match(/^\d+$/)) {
                responsibilities.push(cleaned);
            }
        });
        
        return responsibilities;
    }

    // Determine hierarchy level based on position
    determineHierarchyLevel(position) {
        if (!position) return 6;
        
        const pos = position.toLowerCase();
        
        // Level 1 - Top management
        if (pos.includes('генеральный директор') || 
            pos.includes('президент') || 
            pos.includes('начальник управления')) {
            return 1;
        }
        
        // Level 2 - Deputy heads
        if (pos.includes('заместитель начальника управления') || 
            pos.includes('заместитель директора')) {
            return 2;
        }
        
        // Level 3 - Department heads
        if (pos.includes('начальник отдела') || 
            pos.includes('директор по')) {
            return 3;
        }
        
        // Level 4 - Deputy department heads
        if (pos.includes('заместитель начальника отдела') || 
            pos.includes('руководитель группы')) {
            return 4;
        }
        
        // Level 5 - Team leads and senior specialists
        if (pos.includes('руководитель проектов') || 
            pos.includes('главный специалист') || 
            pos.includes('ведущий')) {
            return 5;
        }
        
        // Level 6 - Specialists
        return 6;
    }

    // Build final hierarchy from structure
    buildFinalHierarchy(structure) {
        let allEmployees = [];
        
        // Collect all employees
        if (structure.employees && structure.employees.length > 0) {
            allEmployees = structure.employees;
        }
        
        // If we have a pre-built hierarchy, use it
        if (structure.hierarchy) {
            return structure.hierarchy;
        }
        
        // Sort employees by level
        allEmployees.sort((a, b) => (a.level || 6) - (b.level || 6));
        
        // Group by departments
        const departmentMap = new Map();
        const noDepartment = [];
        
        allEmployees.forEach(emp => {
            if (emp.department) {
                if (!departmentMap.has(emp.department)) {
                    departmentMap.set(emp.department, []);
                }
                departmentMap.get(emp.department).push(emp);
            } else {
                noDepartment.push(emp);
            }
        });
        
        // Build hierarchy
        const root = {
            id: 'root',
            name: 'Организация',
            position: '',
            children: []
        };
        
        // Add departments
        departmentMap.forEach((employees, deptName) => {
            const deptNode = {
                id: `dept-${Math.random().toString(36).substr(2, 9)}`,
                name: deptName,
                position: 'Подразделение',
                isDepartment: true,
                children: this.buildEmployeeHierarchy(employees)
            };
            root.children.push(deptNode);
        });
        
        // Add employees without department
        if (noDepartment.length > 0) {
            const hierarchy = this.buildEmployeeHierarchy(noDepartment);
            root.children.push(...hierarchy);
        }
        
        return this.optimizeHierarchy(root);
    }

    // Build employee hierarchy
    buildEmployeeHierarchy(employees) {
        const nodes = [];
        const nodeMap = new Map();
        
        // Create nodes
        employees.forEach(emp => {
            const node = {
                id: emp.id || `emp-${Math.random().toString(36).substr(2, 9)}`,
                name: emp.name || 'Не указано',
                position: emp.position || 'Сотрудник',
                department: emp.department,
                responsibilities: emp.responsibilities || [],
                deputy: emp.deputy,
                level: emp.level || 6,
                children: []
            };
            
            nodeMap.set(node.id, node);
            nodes.push(node);
        });
        
        // Build hierarchy based on levels
        const levelGroups = new Map();
        nodes.forEach(node => {
            const level = node.level;
            if (!levelGroups.has(level)) {
                levelGroups.set(level, []);
            }
            levelGroups.get(level).push(node);
        });
        
        // Connect nodes
        const sortedLevels = Array.from(levelGroups.keys()).sort((a, b) => a - b);
        const hierarchy = [];
        
        for (let i = 0; i < sortedLevels.length; i++) {
            const currentLevel = sortedLevels[i];
            const currentNodes = levelGroups.get(currentLevel);
            
            if (i === 0) {
                // Top level
                hierarchy.push(...currentNodes);
            } else {
                // Find parents from previous level
                const parentLevel = sortedLevels[i - 1];
                const parentNodes = levelGroups.get(parentLevel);
                
                if (parentNodes && parentNodes.length > 0) {
                    // Distribute children evenly
                    currentNodes.forEach((node, index) => {
                        const parentIndex = Math.floor(index * parentNodes.length / currentNodes.length);
                        const parent = parentNodes[Math.min(parentIndex, parentNodes.length - 1)];
                        parent.children.push(node);
                    });
                } else {
                    // Add to hierarchy if no parents
                    hierarchy.push(...currentNodes);
                }
            }
        }
        
        return hierarchy;
    }

    // Build hierarchy from nodes
    buildHierarchyFromNodes(nodes) {
        if (nodes.length === 0) return null;
        
        // Sort by level
        nodes.sort((a, b) => (a.level || 6) - (b.level || 6));
        
        // Create hierarchy
        const root = {
            id: 'root',
            name: nodes[0].name,
            position: nodes[0].position,
            children: []
        };
        
        // Simple hierarchy building
        let currentParent = root;
        for (let i = 1; i < nodes.length; i++) {
            const node = {
                id: `node-${i}`,
                name: nodes[i].name,
                position: nodes[i].position,
                children: []
            };
            
            if (nodes[i].level > nodes[i - 1].level) {
                // Child of previous
                currentParent.children.push(node);
                currentParent = node;
            } else if (nodes[i].level === nodes[i - 1].level) {
                // Sibling
                root.children.push(node);
            } else {
                // Higher level - add to root
                root.children.push(node);
                currentParent = node;
            }
        }
        
        return root;
    }

    // Optimize hierarchy
    optimizeHierarchy(root) {
        // Remove empty nodes
        this.removeEmptyNodes(root);
        
        // Flatten single-child chains
        this.flattenSingleChildChains(root);
        
        // Sort children by level
        this.sortChildrenByLevel(root);
        
        return root;
    }

    // Remove empty nodes
    removeEmptyNodes(node) {
        if (!node.children) return;
        
        node.children = node.children.filter(child => {
            if (!child.name || child.name === 'Не указано') {
                // Move children up
                if (child.children && child.children.length > 0) {
                    node.children.push(...child.children);
                }
                return false;
            }
            this.removeEmptyNodes(child);
            return true;
        });
    }

    // Flatten single-child chains
    flattenSingleChildChains(node) {
        if (!node.children || node.children.length !== 1) {
            if (node.children) {
                node.children.forEach(child => this.flattenSingleChildChains(child));
            }
            return;
        }
        
        const onlyChild = node.children[0];
        if (onlyChild.isDepartment || node.isDepartment) {
            // Don't flatten department nodes
            this.flattenSingleChildChains(onlyChild);
        } else if (onlyChild.children && onlyChild.children.length > 0) {
            // Skip the single child and adopt its children
            node.children = onlyChild.children;
            this.flattenSingleChildChains(node);
        }
    }

    // Sort children by level
    sortChildrenByLevel(node) {
        if (!node.children || node.children.length === 0) return;
        
        node.children.sort((a, b) => {
            const levelA = a.level || 6;
            const levelB = b.level || 6;
            return levelA - levelB;
        });
        
        node.children.forEach(child => this.sortChildrenByLevel(child));
    }

    // Convert hierarchy to table format for export
    hierarchyToTable(root) {
        const rows = [];
        let counter = 0;
        
        const traverse = (node, parentCode = '', level = 0) => {
            counter++;
            const code = parentCode ? `${parentCode}.${counter}` : counter.toString();
            
            rows.push({
                'Подчиненность': code,
                'Подразделение': node.department || node.isDepartment ? node.name : '',
                'Должность': node.position || '',
                'ФИО': node.name || '',
                'Функционал': node.responsibilities ? node.responsibilities.join('; ') : '',
                'Уровень': level,
                'Заместитель': node.deputy || ''
            });
            
            if (node.children) {
                node.children.forEach(child => {
                    traverse(child, code, level + 1);
                });
            }
        };
        
        traverse(root);
        return rows;
    }

    // Export to CSV format
    exportToCSV(hierarchy) {
        const table = this.hierarchyToTable(hierarchy);
        const headers = Object.keys(table[0]);
        
        let csv = headers.map(h => `"${h}"`).join(',') + '\n';
        
        table.forEach(row => {
            const values = headers.map(h => {
                const value = row[h] || '';
                return `"${String(value).replace(/"/g, '""')}"`;
            });
            csv += values.join(',') + '\n';
        });
        
        return csv;
    }

    // Export to Google Sheets format
    exportToGoogleSheets(hierarchy) {
        const table = this.hierarchyToTable(hierarchy);
        
        // Create a format optimized for Google Sheets
        const formatted = table.map(row => ({
            'Код подчинения': row['Подчиненность'],
            'Подразделение': row['Подразделение'],
            'Должность': row['Должность'],
            'ФИО': row['ФИО'],
            'Функциональные обязанности': row['Функционал'],
            'Уровень иерархии': row['Уровень'],
            'Заместитель': row['Заместитель']
        }));
        
        return formatted;
    }
}

// Export for use
window.SpecializedOrgParser = SpecializedOrgParser;