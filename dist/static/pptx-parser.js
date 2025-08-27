// Specialized PPTX parser for organizational structures
class PPTXOrgParser {
    constructor() {
        this.slides = [];
        this.structure = null;
    }

    // Main parsing method for PPTX
    async parsePPTX(file) {
        try {
            // Load JSZip if not already loaded
            if (!window.JSZip) {
                await this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js');
            }

            const zip = new JSZip();
            const zipContent = await zip.loadAsync(file);
            
            // Extract slides content
            const slides = await this.extractSlides(zipContent);
            
            // Parse organizational structure from slides
            const structure = this.parseOrganizationalStructure(slides);
            
            return structure;
        } catch (error) {
            console.error('Error parsing PPTX:', error);
            throw error;
        }
    }

    // Extract slides from PPTX
    async extractSlides(zipContent) {
        const slides = [];
        const slideFiles = Object.keys(zipContent.files).filter(name => 
            name.startsWith('ppt/slides/slide') && name.endsWith('.xml')
        ).sort();

        for (const slideFile of slideFiles) {
            const content = await zipContent.files[slideFile].async('string');
            const slideText = this.extractTextFromSlideXML(content);
            slides.push(slideText);
        }

        return slides;
    }

    // Extract text from slide XML
    extractTextFromSlideXML(xml) {
        const texts = [];
        
        // Extract text from <a:t> tags
        const textMatches = xml.match(/<a:t[^>]*>([^<]+)<\/a:t>/g);
        if (textMatches) {
            textMatches.forEach(match => {
                const text = match.replace(/<[^>]+>/g, '').trim();
                if (text) {
                    texts.push(text);
                }
            });
        }

        return texts;
    }

    // Parse organizational structure from slides
    parseOrganizationalStructure(slides) {
        const allTexts = [];
        const employees = [];
        const departments = new Set();
        const positions = new Map();

        // Collect all texts from slides
        slides.forEach(slide => {
            allTexts.push(...slide);
        });

        // Parse employees and positions
        for (let i = 0; i < allTexts.length; i++) {
            const text = allTexts[i].trim();
            
            // Skip empty or very short texts
            if (!text || text.length < 3) continue;

            // Check if it's a department
            if (this.isDepartment(text)) {
                departments.add(text);
                continue;
            }

            // Check if it's a position
            if (this.isPosition(text)) {
                // Look for associated name
                let name = '';
                
                // Check next items for a name
                for (let j = i + 1; j < Math.min(i + 3, allTexts.length); j++) {
                    const nextText = allTexts[j].trim();
                    if (this.isPersonName(nextText)) {
                        name = nextText;
                        i = j; // Skip processed name
                        break;
                    }
                }

                // Sometimes position and name are in the same text
                if (!name) {
                    const combined = this.extractPositionAndName(text);
                    if (combined) {
                        employees.push(combined);
                        continue;
                    }
                }

                if (name) {
                    employees.push({
                        name: name,
                        position: text,
                        level: this.determineLevel(text)
                    });
                }
            } else if (this.isPersonName(text)) {
                // Check if previous item was a position
                let position = '';
                if (i > 0) {
                    const prevText = allTexts[i - 1].trim();
                    if (this.isPosition(prevText)) {
                        position = prevText;
                    }
                }

                employees.push({
                    name: text,
                    position: position || 'Сотрудник',
                    level: position ? this.determineLevel(position) : 6
                });
            }
        }

        // Build hierarchy
        return this.buildHierarchy(employees, departments);
    }

    // Check if text is a department name
    isDepartment(text) {
        const deptKeywords = [
            'ОТДЕЛ', 'УПРАВЛЕНИЕ', 'ДЕПАРТАМЕНТ', 'НАПРАВЛЕНИЕ', 
            'ПОДРАЗДЕЛЕНИЕ', 'СЛУЖБА', 'ГРУППА', 'СЕКТОР'
        ];
        
        const upperText = text.toUpperCase();
        return deptKeywords.some(keyword => upperText.includes(keyword));
    }

    // Check if text is a position
    isPosition(text) {
        const positionKeywords = [
            'начальник', 'заместитель', 'руководитель', 'директор',
            'менеджер', 'специалист', 'координатор', 'аналитик',
            'инженер', 'консультант', 'эксперт', 'главный',
            'ведущий', 'старший', 'технический', 'проектов'
        ];

        const lowerText = text.toLowerCase();
        return positionKeywords.some(keyword => lowerText.includes(keyword));
    }

    // Check if text is a person name
    isPersonName(text) {
        // Check for Russian name pattern (Фамилия Имя Отчество)
        const namePattern = /^[А-ЯЁ][а-яё]+(\s+[А-ЯЁ][а-яё]+){1,2}$/;
        return namePattern.test(text.trim());
    }

    // Extract position and name from combined text
    extractPositionAndName(text) {
        // Pattern for "PositionName Surname" format
        const patterns = [
            // Position followed by name
            /^(.+?)((?:[А-ЯЁ][а-яё]+\s*){2,3})$/,
            // Name followed by position
            /^((?:[А-ЯЁ][а-яё]+\s*){2,3})(.+)$/
        ];

        for (const pattern of patterns) {
            const match = text.match(pattern);
            if (match) {
                let position = '';
                let name = '';

                // Determine which part is position and which is name
                if (this.isPosition(match[1])) {
                    position = match[1].trim();
                    name = match[2].trim();
                } else if (this.isPersonName(match[1])) {
                    name = match[1].trim();
                    position = match[2].trim();
                }

                if (name && this.isPersonName(name)) {
                    return {
                        name: name,
                        position: position || 'Сотрудник',
                        level: position ? this.determineLevel(position) : 6
                    };
                }
            }
        }

        return null;
    }

    // Determine hierarchy level
    determineLevel(position) {
        if (!position) return 6;

        const pos = position.toLowerCase();

        if (pos.includes('генеральный') || pos.includes('президент')) {
            return 1;
        }
        if (pos.includes('заместитель начальника управления')) {
            return 2;
        }
        if (pos.includes('начальник управления') || pos.includes('директор департамента')) {
            return 2;
        }
        if (pos.includes('начальник отдела')) {
            return 3;
        }
        if (pos.includes('заместитель начальника отдела')) {
            return 4;
        }
        if (pos.includes('руководитель группы') || pos.includes('руководитель проектов')) {
            return 4;
        }
        if (pos.includes('технический руководитель')) {
            return 4;
        }
        if (pos.includes('главный') || pos.includes('ведущий')) {
            return 5;
        }
        if (pos.includes('старший')) {
            return 5;
        }

        return 6;
    }

    // Build hierarchy from employees and departments
    buildHierarchy(employees, departments) {
        // Sort employees by level
        employees.sort((a, b) => a.level - b.level);

        // Create root node
        const root = {
            id: 'org-root',
            name: 'Организация',
            position: '',
            children: []
        };

        // Group employees by department if departments exist
        if (departments.size > 0) {
            const deptNodes = [];
            
            departments.forEach(deptName => {
                const deptNode = {
                    id: `dept-${Math.random().toString(36).substr(2, 9)}`,
                    name: deptName,
                    position: 'Подразделение',
                    isDepartment: true,
                    children: []
                };
                deptNodes.push(deptNode);
            });

            // Add department nodes to root
            root.children.push(...deptNodes);
        }

        // Build employee hierarchy
        const levelGroups = new Map();
        
        employees.forEach(emp => {
            const level = emp.level;
            if (!levelGroups.has(level)) {
                levelGroups.set(level, []);
            }
            
            const node = {
                id: `emp-${Math.random().toString(36).substr(2, 9)}`,
                name: emp.name,
                position: emp.position,
                level: emp.level,
                children: []
            };
            
            levelGroups.get(level).push(node);
        });

        // Connect nodes by levels
        const sortedLevels = Array.from(levelGroups.keys()).sort((a, b) => a - b);
        const topLevel = levelGroups.get(sortedLevels[0]) || [];

        // Add top level to root or departments
        if (root.children.length > 0 && root.children[0].isDepartment) {
            // Distribute top level among departments
            topLevel.forEach((node, index) => {
                const deptIndex = index % root.children.length;
                root.children[deptIndex].children.push(node);
            });
        } else {
            root.children.push(...topLevel);
        }

        // Connect remaining levels
        for (let i = 1; i < sortedLevels.length; i++) {
            const currentLevel = sortedLevels[i];
            const prevLevel = sortedLevels[i - 1];
            const currentNodes = levelGroups.get(currentLevel);
            const parentNodes = levelGroups.get(prevLevel);

            if (parentNodes && parentNodes.length > 0) {
                currentNodes.forEach((node, index) => {
                    const parentIndex = Math.floor(index * parentNodes.length / currentNodes.length);
                    const parent = parentNodes[Math.min(parentIndex, parentNodes.length - 1)];
                    parent.children.push(node);
                });
            }
        }

        // Optimize structure
        return this.optimizeStructure(root);
    }

    // Optimize the final structure
    optimizeStructure(root) {
        // Remove empty children arrays
        this.cleanEmptyChildren(root);
        
        // If root has only one child and it's not a department, return the child
        if (root.children.length === 1 && !root.children[0].isDepartment && root.name === 'Организация') {
            return root.children[0];
        }

        return root;
    }

    // Clean empty children arrays
    cleanEmptyChildren(node) {
        if (node.children && node.children.length === 0) {
            delete node.children;
        } else if (node.children) {
            node.children.forEach(child => this.cleanEmptyChildren(child));
        }
    }

    // Load external script
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

    // Parse PPTX from URL
    async parsePPTXFromURL(url) {
        try {
            const response = await fetch(url);
            const blob = await response.blob();
            return await this.parsePPTX(blob);
        } catch (error) {
            console.error('Error fetching PPTX from URL:', error);
            throw error;
        }
    }

    // Enhanced parsing for specific PPTX structure
    parseEnhancedPPTX(slides) {
        const structure = {
            management: [],
            departments: new Map(),
            projects: [],
            positions: []
        };

        // Patterns for different elements
        const patterns = {
            management: /заместитель\s+начальника\s+управления|начальник\s+управления/i,
            department: /ОТДЕЛ\s+[А-ЯЁ\s]+|отдел\s+[а-яё\s]+/i,
            projectManager: /руководитель\s+проектов|технический\s+руководитель/i,
            departmentHead: /начальник\s+отдела/i,
            deputyHead: /заместитель\s+начальника\s+отдела/i
        };

        // Parse all slide texts
        slides.forEach(slide => {
            slide.forEach(text => {
                const cleanText = text.trim();
                
                // Parse management positions
                if (patterns.management.test(cleanText)) {
                    const nameMatch = this.findAssociatedName(slide, cleanText);
                    if (nameMatch) {
                        structure.management.push({
                            position: cleanText,
                            name: nameMatch,
                            level: 2
                        });
                    }
                }
                
                // Parse departments
                if (patterns.department.test(cleanText)) {
                    const deptName = cleanText.replace(/^#+\s*/, '').trim();
                    if (!structure.departments.has(deptName)) {
                        structure.departments.set(deptName, {
                            name: deptName,
                            head: null,
                            deputies: [],
                            staff: []
                        });
                    }
                }
                
                // Parse project managers
                if (patterns.projectManager.test(cleanText)) {
                    const nameMatch = this.findAssociatedName(slide, cleanText);
                    if (nameMatch) {
                        structure.projects.push({
                            position: cleanText,
                            name: nameMatch,
                            level: 4
                        });
                    }
                }
            });
        });

        return this.buildEnhancedHierarchy(structure);
    }

    // Find associated name for a position
    findAssociatedName(slideTexts, position) {
        const index = slideTexts.indexOf(position);
        
        // Look for name after position
        for (let i = index + 1; i < Math.min(index + 5, slideTexts.length); i++) {
            if (this.isPersonName(slideTexts[i])) {
                return slideTexts[i];
            }
        }
        
        // Look for name before position
        for (let i = index - 1; i >= Math.max(0, index - 3); i--) {
            if (this.isPersonName(slideTexts[i])) {
                return slideTexts[i];
            }
        }
        
        return null;
    }

    // Build enhanced hierarchy
    buildEnhancedHierarchy(structure) {
        const root = {
            id: 'root',
            name: 'Управление',
            position: 'Организация',
            children: []
        };

        // Add management
        structure.management.forEach(manager => {
            const node = {
                id: `mgmt-${Math.random().toString(36).substr(2, 9)}`,
                name: manager.name || 'Не указано',
                position: manager.position,
                level: manager.level,
                children: []
            };
            root.children.push(node);
        });

        // Add departments
        structure.departments.forEach((dept, deptName) => {
            const deptNode = {
                id: `dept-${Math.random().toString(36).substr(2, 9)}`,
                name: deptName,
                position: 'Подразделение',
                isDepartment: true,
                children: []
            };

            // Add department head
            if (dept.head) {
                deptNode.children.push({
                    id: `head-${Math.random().toString(36).substr(2, 9)}`,
                    name: dept.head.name,
                    position: dept.head.position,
                    level: 3,
                    children: []
                });
            }

            // Add deputies
            dept.deputies.forEach(deputy => {
                deptNode.children.push({
                    id: `deputy-${Math.random().toString(36).substr(2, 9)}`,
                    name: deputy.name,
                    position: deputy.position,
                    level: 4,
                    children: []
                });
            });

            // Attach to appropriate manager
            if (root.children.length > 0) {
                root.children[0].children.push(deptNode);
            } else {
                root.children.push(deptNode);
            }
        });

        // Add project managers
        structure.projects.forEach(pm => {
            const pmNode = {
                id: `pm-${Math.random().toString(36).substr(2, 9)}`,
                name: pm.name || 'Потребность',
                position: pm.position,
                level: pm.level,
                children: []
            };

            // Add to first available parent
            if (root.children.length > 0 && root.children[0].children.length > 0) {
                root.children[0].children[0].children.push(pmNode);
            } else if (root.children.length > 0) {
                root.children[0].children.push(pmNode);
            } else {
                root.children.push(pmNode);
            }
        });

        return root;
    }
}

// Export for global use
window.PPTXOrgParser = PPTXOrgParser;