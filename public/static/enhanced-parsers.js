// Enhanced parsers for all document formats
class EnhancedDocumentParser {
    constructor() {
        this.debug = true; // Enable debug logging
    }

    // Parse Google Sheets CSV with proper handling
    parseGoogleSheetsCSV(csvText) {
        console.log('Parsing Google Sheets CSV...');
        
        const lines = csvText.split('\n').filter(line => line.trim());
        console.log(`Found ${lines.length} lines in CSV`);
        
        if (lines.length === 0) return null;

        // Parse CSV properly handling quotes and commas
        const rows = [];
        for (let line of lines) {
            const row = this.parseCSVLine(line);
            if (row && row.length > 0) {
                rows.push(row);
            }
        }

        console.log(`Parsed ${rows.length} rows`);

        // First row is headers
        const headers = rows[0];
        console.log('Headers:', headers);

        // Build node map for hierarchy
        const nodeMap = new Map();
        const nodes = [];
        
        // Process each data row
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;

            const node = {
                id: `node-${i}`,
                subordination: row[0] || '',
                department: row[1] || '',
                position: row[2] || '',
                name: row[3] || '',
                functional: row[4] || '',
                children: [],
                rawData: row
            };

            // Only add nodes with meaningful data
            if (node.subordination || node.department || node.position || node.name) {
                nodes.push(node);
                nodeMap.set(node.subordination, node);
            }
        }

        console.log(`Created ${nodes.length} nodes`);

        // Build hierarchy based on subordination codes
        return this.buildHierarchyFromSubordination(nodes);
    }

    // Parse CSV line handling quotes properly
    parseCSVLine(line) {
        const result = [];
        let current = '';
        let inQuotes = false;
        
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            const nextChar = line[i + 1];
            
            if (char === '"') {
                if (inQuotes && nextChar === '"') {
                    // Escaped quote
                    current += '"';
                    i++; // Skip next quote
                } else {
                    // Toggle quote mode
                    inQuotes = !inQuotes;
                }
            } else if (char === ',' && !inQuotes) {
                result.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        
        // Add last field
        if (current || result.length > 0) {
            result.push(current.trim());
        }
        
        return result;
    }

    // Build hierarchy from subordination codes
    buildHierarchyFromSubordination(nodes) {
        const roots = [];
        const nodeMap = new Map();

        // Создаем карту узлов по ID
        nodes.forEach(node => {
            node.id = node.rawData[0]; // ID из первой колонки
            node.children = [];
            nodeMap.set(node.id, node);
        });

        // Строим иерархию
        nodes.forEach(node => {
            const parentId = node.rawData[1]; // Подчиненность из второй колонки
            if (!parentId || parentId.trim() === '') {
                // Если нет родителя, это корневой узел
                roots.push(node);
            } else {
                // Находим родителя и добавляем текущий узел как его ребенка
                const parent = nodeMap.get(parentId);
                if (parent) {
                    parent.children.push(node);
                }
            }
        });

        // Преобразуем узлы в требуемый формат
        const transformNode = (node) => {
            return {
                id: node.id,
                name: node.rawData[4] || '', // ФИО
                title: node.rawData[3] || '', // Должность
                department: node.rawData[2] || '', // Подразделение
                children: node.children.map(child => transformNode(child))
            };
        };

        // Если есть только один корневой узел, возвращаем его
        if (roots.length === 1) {
            return {
                success: true,
                data: transformNode(roots[0])
            };
        }

        // Если несколько корневых узлов, создаем виртуальный корневой узел
        return {
            success: true,
            data: {
                id: 'root',
                name: 'Организация',
                children: roots.map(root => transformNode(root))
            }
        };
    }

    // Find parent node based on subordination code
    findParentNode(code, nodeMap, allNodes) {
        if (!code) return null;

        // Direct parent patterns
        const patterns = [
            code.substring(0, code.length - 1), // Remove last char
            Math.floor(parseInt(code) / 10).toString(), // Divide by 10
            '0', // Default to top
            '00' // Alternative top
        ];

        for (let pattern of patterns) {
            if (nodeMap.has(pattern)) {
                return nodeMap.get(pattern);
            }
        }

        // Find by position hierarchy
        const currentNode = nodeMap.get(code);
        if (currentNode && currentNode.position) {
            const level = this.getPositionLevel(currentNode.position);
            
            // Find a manager at a higher level
            for (let node of allNodes) {
                if (node !== currentNode) {
                    const nodeLevel = this.getPositionLevel(node.position);
                    if (nodeLevel < level) {
                        return node;
                    }
                }
            }
        }

        return null;
    }

    // Find closest parent by comparing subordination numbers
    findClosestParent(node, allNodes) {
        const nodeNum = parseInt(node.subordination);
        if (!nodeNum) return null;

        let closestParent = null;
        let closestDiff = Infinity;

        for (let candidate of allNodes) {
            if (candidate === node) continue;
            
            const candidateNum = parseInt(candidate.subordination);
            if (candidateNum && candidateNum < nodeNum) {
                const diff = nodeNum - candidateNum;
                if (diff < closestDiff) {
                    closestDiff = diff;
                    closestParent = candidate;
                }
            }
        }

        return closestParent;
    }

    // Get position hierarchy level
    getPositionLevel(position) {
        if (!position) return 99;
        
        const pos = position.toLowerCase();
        
        if (pos.includes('начальник управления')) return 1;
        if (pos.includes('заместитель начальника управления')) return 2;
        if (pos.includes('директор')) return 2;
        if (pos.includes('начальник отдела')) return 3;
        if (pos.includes('заместитель начальника отдела')) return 4;
        if (pos.includes('руководитель группы')) return 4;
        if (pos.includes('руководитель проект')) return 5;
        if (pos.includes('главный')) return 5;
        if (pos.includes('ведущий')) return 6;
        if (pos.includes('старший')) return 6;
        if (pos.includes('специалист')) return 7;
        
        return 8;
    }

    // Parse PPTX with improved extraction
    async parsePPTXDocument(file) {
        console.log('Parsing PPTX document...');
        
        try {
            // Load JSZip if needed
            if (!window.JSZip) {
                await this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js');
            }

            const zip = new JSZip();
            const content = await zip.loadAsync(file);
            
            // Extract all slides
            const slides = [];
            const slideFiles = Object.keys(content.files)
                .filter(name => name.match(/ppt\/slides\/slide\d+\.xml$/))
                .sort((a, b) => {
                    const numA = parseInt(a.match(/slide(\d+)/)[1]);
                    const numB = parseInt(b.match(/slide(\d+)/)[1]);
                    return numA - numB;
                });

            console.log(`Found ${slideFiles.length} slides`);

            for (let slideFile of slideFiles) {
                const xml = await content.files[slideFile].async('string');
                const texts = this.extractTextsFromXML(xml);
                if (texts.length > 0) {
                    slides.push({
                        slideNumber: parseInt(slideFile.match(/slide(\d+)/)[1]),
                        texts: texts
                    });
                }
            }

            // Parse structure from slides
            return this.buildStructureFromSlides(slides);
        } catch (error) {
            console.error('Error parsing PPTX:', error);
            return null;
        }
    }

    // Extract texts from slide XML
    extractTextsFromXML(xml) {
        const texts = [];
        
        // Extract all text runs
        const textPattern = /<a:t[^>]*>([^<]+)<\/a:t>/g;
        let match;
        
        while ((match = textPattern.exec(xml)) !== null) {
            const text = match[1].trim();
            if (text && text.length > 1) {
                texts.push(text);
            }
        }
        
        return texts;
    }

    // Build structure from slides
    buildStructureFromSlides(slides) {
        const employees = [];
        const departments = new Set();
        const positions = [];

        // Process each slide
        slides.forEach(slide => {
            console.log(`Processing slide ${slide.slideNumber} with ${slide.texts.length} texts`);
            
            let currentDepartment = null;
            let pendingPosition = null;
            
            slide.texts.forEach((text, index) => {
                // Check if it's a department
                if (this.isDepartmentName(text)) {
                    currentDepartment = text;
                    departments.add(text);
                    return;
                }

                // Check if it's a position
                if (this.isPositionTitle(text)) {
                    pendingPosition = text;
                    
                    // Look ahead for a name
                    if (index + 1 < slide.texts.length) {
                        const nextText = slide.texts[index + 1];
                        if (this.isPersonName(nextText)) {
                            employees.push({
                                name: nextText,
                                position: pendingPosition,
                                department: currentDepartment,
                                level: this.getPositionLevel(pendingPosition)
                            });
                            pendingPosition = null;
                        }
                    }
                    return;
                }

                // Check if it's a name
                if (this.isPersonName(text)) {
                    employees.push({
                        name: text,
                        position: pendingPosition || 'Сотрудник',
                        department: currentDepartment,
                        level: pendingPosition ? this.getPositionLevel(pendingPosition) : 8
                    });
                    pendingPosition = null;
                    return;
                }

                // Try to parse combined position+name
                const combined = this.parsePositionAndName(text);
                if (combined) {
                    employees.push({
                        ...combined,
                        department: currentDepartment,
                        level: this.getPositionLevel(combined.position)
                    });
                }
            });
        });

        console.log(`Found ${employees.length} employees, ${departments.size} departments`);

        // Build hierarchy
        return this.buildHierarchyFromEmployees(employees, Array.from(departments));
    }

    // Check if text is a department name
    isDepartmentName(text) {
        const upper = text.toUpperCase();
        return upper.includes('ОТДЕЛ') || 
               upper.includes('УПРАВЛЕНИЕ') || 
               upper.includes('ДЕПАРТАМЕНТ') ||
               upper.includes('НАПРАВЛЕНИЕ') ||
               upper.includes('БЛОК') ||
               upper.includes('СЛУЖБА');
    }

    // Check if text is a position title
    isPositionTitle(text) {
        const lower = text.toLowerCase();
        return lower.includes('начальник') ||
               lower.includes('заместитель') ||
               lower.includes('руководитель') ||
               lower.includes('директор') ||
               lower.includes('менеджер') ||
               lower.includes('специалист') ||
               lower.includes('координатор') ||
               lower.includes('аналитик') ||
               lower.includes('консультант');
    }

    // Check if text is a person name
    isPersonName(text) {
        // Russian name pattern: at least 2 words starting with capital letters
        const pattern = /^[А-ЯЁ][а-яё]+(\s+[А-ЯЁ][а-яё]+){1,2}$/;
        return pattern.test(text.trim());
    }

    // Parse combined position and name
    parsePositionAndName(text) {
        // Try different patterns
        const patterns = [
            // Position then name
            /^(.+?(?:начальник|заместитель|руководитель|директор|менеджер|специалист).+?)\s+([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ][а-яё]+)?)$/i,
            // Name then position
            /^([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ][а-яё]+)?)\s*[-–]\s*(.+)$/,
            // Concatenated (no space)
            /^(.+?(?:проектов|отдела|управления))([А-ЯЁ][а-яё]+\s*[А-ЯЁ][а-яё]+.*)$/
        ];

        for (let pattern of patterns) {
            const match = text.match(pattern);
            if (match) {
                if (this.isPositionTitle(match[1]) && this.isPersonName(match[2])) {
                    return { position: match[1].trim(), name: match[2].trim() };
                } else if (this.isPersonName(match[1]) && this.isPositionTitle(match[2])) {
                    return { name: match[1].trim(), position: match[2].trim() };
                }
            }
        }

        return null;
    }

    // Build hierarchy from employees
    buildHierarchyFromEmployees(employees, departments) {
        // Sort by level
        employees.sort((a, b) => a.level - b.level);

        // Create root
        const root = {
            id: 'org',
            name: 'Организация',
            position: '',
            children: []
        };

        // Create department nodes
        const deptNodes = {};
        departments.forEach(dept => {
            deptNodes[dept] = {
                id: `dept-${Math.random().toString(36).substr(2, 9)}`,
                name: dept,
                position: 'Подразделение',
                isDepartment: true,
                children: []
            };
            root.children.push(deptNodes[dept]);
        });

        // Group employees by level
        const levels = {};
        employees.forEach(emp => {
            const level = emp.level || 8;
            if (!levels[level]) levels[level] = [];
            
            const node = {
                id: `emp-${Math.random().toString(36).substr(2, 9)}`,
                name: emp.name,
                position: emp.position,
                department: emp.department,
                level: level,
                children: []
            };
            
            levels[level].push(node);
        });

        // Build hierarchy by levels
        const sortedLevels = Object.keys(levels).map(Number).sort((a, b) => a - b);
        
        sortedLevels.forEach((level, index) => {
            const nodes = levels[level];
            
            nodes.forEach(node => {
                if (node.department && deptNodes[node.department]) {
                    // Add to department
                    if (level <= 3) {
                        // High-level positions go directly under department
                        deptNodes[node.department].children.push(node);
                    } else {
                        // Lower positions go under department heads
                        const deptHead = deptNodes[node.department].children[0];
                        if (deptHead) {
                            deptHead.children.push(node);
                        } else {
                            deptNodes[node.department].children.push(node);
                        }
                    }
                } else if (index === 0) {
                    // Top level without department
                    root.children.push(node);
                } else {
                    // Find parent from previous level
                    const prevLevel = sortedLevels[index - 1];
                    const potentialParents = levels[prevLevel];
                    if (potentialParents && potentialParents.length > 0) {
                        // Add to first available parent
                        potentialParents[0].children.push(node);
                    } else {
                        root.children.push(node);
                    }
                }
            });
        });

        return this.optimizeStructure(root);
    }

    // Parse DOCX document
    parseWordDocument(text) {
        console.log('Parsing Word document...');
        
        const employees = [];
        const lines = text.split(/\n|\r\n|\r/).filter(line => line.trim());
        
        let currentDepartment = null;
        let currentEntry = null;

        lines.forEach(line => {
            const cleanLine = line.replace(/<[^>]*>/g, '').trim();
            if (!cleanLine) return;

            // Check for department
            if (this.isDepartmentName(cleanLine)) {
                currentDepartment = cleanLine;
                return;
            }

            // Check for numbered entry (like "1. Name - Position: duties")
            const numberMatch = cleanLine.match(/^(\d+)\.\s*(.+)/);
            if (numberMatch) {
                if (currentEntry) {
                    employees.push(currentEntry);
                }
                
                const content = numberMatch[2];
                currentEntry = this.parseEmployeeEntry(content);
                if (currentEntry) {
                    currentEntry.department = currentDepartment;
                    currentEntry.number = numberMatch[1];
                }
            } else if (currentEntry && cleanLine.length > 20) {
                // Add as additional responsibilities
                if (!currentEntry.responsibilities) currentEntry.responsibilities = [];
                currentEntry.responsibilities.push(cleanLine);
            }
        });

        // Add last entry
        if (currentEntry) {
            employees.push(currentEntry);
        }

        console.log(`Parsed ${employees.length} employees from Word document`);
        return this.buildHierarchyFromEmployees(employees, []);
    }

    // Parse employee entry from text
    parseEmployeeEntry(text) {
        // Pattern: Name - Position (Deputy): Responsibilities
        const patterns = [
            /^([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)\s*[-–]\s*([^:]+)(?::\s*(.+))?/,
            /^([А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.)\s*[-–]\s*([^:]+)(?::\s*(.+))?/
        ];

        for (let pattern of patterns) {
            const match = text.match(pattern);
            if (match) {
                const entry = {
                    name: match[1].trim(),
                    position: match[2].trim(),
                    responsibilities: []
                };

                // Extract deputy info
                const deputyMatch = entry.position.match(/\(([^)]+)\)/);
                if (deputyMatch) {
                    entry.deputy = deputyMatch[1];
                    entry.position = entry.position.replace(/\([^)]+\)/, '').trim();
                }

                // Add responsibilities
                if (match[3]) {
                    entry.responsibilities = match[3].split(/[.;]/)
                        .map(r => r.trim())
                        .filter(r => r.length > 5);
                }

                entry.level = this.getPositionLevel(entry.position);
                return entry;
            }
        }

        return null;
    }

    // Optimize structure
    optimizeStructure(root) {
        // Remove empty children
        this.cleanEmptyNodes(root);
        
        // Remove redundant single-child chains
        this.flattenSingleChildren(root);
        
        // Sort children by level
        this.sortChildren(root);
        
        return root;
    }

    // Clean empty nodes
    cleanEmptyNodes(node) {
        if (!node.children || node.children.length === 0) {
            delete node.children;
            return;
        }

        node.children = node.children.filter(child => {
            if (!child.name || child.name === 'Не указано') {
                // Promote grandchildren
                if (child.children) {
                    return false; // Will be replaced by grandchildren
                }
                return false;
            }
            this.cleanEmptyNodes(child);
            return true;
        });

        // Promote grandchildren of removed nodes
        const newChildren = [];
        node.children.forEach(child => {
            if (!child.name && child.children) {
                newChildren.push(...child.children);
            } else {
                newChildren.push(child);
            }
        });
        node.children = newChildren;
    }

    // Flatten single children
    flattenSingleChildren(node) {
        if (!node.children) return;
        
        node.children.forEach(child => this.flattenSingleChildren(child));
        
        // Don't flatten departments or if node has meaningful data
        if (node.isDepartment || node.position || (node.name && node.name !== 'Организация')) {
            return;
        }
        
        // If this node only has one child and no meaningful data, promote the child
        if (node.children.length === 1 && !node.position && (!node.name || node.name === 'Организация')) {
            const onlyChild = node.children[0];
            Object.assign(node, onlyChild);
        }
    }

    // Sort children by level and name
    sortChildren(node) {
        if (!node.children) return;
        
        node.children.sort((a, b) => {
            // Departments first
            if (a.isDepartment && !b.isDepartment) return -1;
            if (!a.isDepartment && b.isDepartment) return 1;
            
            // Then by level
            const levelA = a.level || 99;
            const levelB = b.level || 99;
            if (levelA !== levelB) return levelA - levelB;
            
            // Then by name
            return (a.name || '').localeCompare(b.name || '', 'ru');
        });
        
        node.children.forEach(child => this.sortChildren(child));
    }

    // Load script helper
    loadScript(src) {
        return new Promise((resolve, reject) => {
            if (document.querySelector(`script[src="${src}"]`)) {
                resolve();
                return;
            }
            const script = document.createElement('script');
            script.src = src;
            script.onload = resolve;
            script.onerror = reject;
            document.head.appendChild(script);
        });
    }

    // Create detailed table from structure
    createDetailedTable(root) {
        const rows = [];
        let idCounter = 0;
        
        const traverse = (node, parentId = '', level = 0, parentName = '') => {
            idCounter++;
            const nodeId = parentId ? `${parentId}.${idCounter}` : idCounter.toString();
            
            // Add node to table
            rows.push({
                'Код': nodeId,
                'Уровень': level,
                'Подразделение': node.isDepartment ? node.name : (node.department || ''),
                'Должность': node.position || '',
                'ФИО': node.name || '',
                'Подчинение': parentName,
                'Количество подчиненных': node.children ? node.children.length : 0,
                'Тип': node.isDepartment ? 'Подразделение' : 'Сотрудник'
            });
            
            // Process children
            if (node.children) {
                const currentId = idCounter;
                node.children.forEach(child => {
                    traverse(child, nodeId, level + 1, node.name);
                });
            }
        };
        
        traverse(root);
        return rows;
    }
}

// Export for global use
window.EnhancedDocumentParser = EnhancedDocumentParser;

