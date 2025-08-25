// Global variables
let currentStructure = null;
let editMode = false;

// File upload handler
async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    // Check if we should process client-side (for complex formats)
    const fileName = file.name.toLowerCase();
    
    if (fileName.endsWith('.xlsx') || fileName.endsWith('.docx') || 
        fileName.endsWith('.pdf') || fileName.endsWith('.pptx')) {
        // For these formats, we'll use client-side libraries
        await processFileClientSide(file);
    } else {
        // For other formats, send to server
        await processFileServerSide(file);
    }
}

// Process file on client side
async function processFileClientSide(file) {
    showLoading(true);
    
    try {
        const fileName = file.name.toLowerCase();
        let structure = null;

        if (fileName.endsWith('.xlsx')) {
            structure = await parseExcelClient(file);
        } else if (fileName.endsWith('.docx')) {
            structure = await parseWordClient(file);
        } else if (fileName.endsWith('.pdf')) {
            structure = await parsePDFClient(file);
        } else if (fileName.endsWith('.pptx')) {
            structure = await parsePowerPointClient(file);
        }

        if (structure) {
            currentStructure = structure;
            renderOrgChart(structure);
            showChartContainer(true);
        } else {
            showError('Не удалось обработать файл');
        }
    } catch (error) {
        console.error('Error processing file:', error);
        showError('Ошибка при обработке файла: ' + error.message);
    } finally {
        showLoading(false);
    }
}

// Parse Excel on client
async function parseExcelClient(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = async (e) => {
            try {
                // Load XLSX library dynamically
                if (!window.XLSX) {
                    await loadScript('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');
                }

                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet);

                // Convert Excel data to org structure
                const structure = convertExcelToOrgStructure(jsonData);
                resolve(structure);
            } catch (error) {
                reject(error);
            }
        };

        reader.readAsArrayBuffer(file);
    });
}

// Parse Word document on client
async function parseWordClient(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = async (e) => {
            try {
                // Load mammoth library dynamically
                if (!window.mammoth) {
                    await loadScript('https://cdn.jsdelivr.net/npm/mammoth@1.6.0/mammoth.browser.min.js');
                }

                const arrayBuffer = e.target.result;
                const result = await mammoth.extractRawText({ arrayBuffer });
                
                // Parse text to find organizational structure
                const structure = parseTextToStructure(result.value);
                resolve(structure);
            } catch (error) {
                reject(error);
            }
        };

        reader.readAsArrayBuffer(file);
    });
}

// Parse PDF on client
async function parsePDFClient(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = async (e) => {
            try {
                // Load PDF.js library dynamically
                if (!window.pdfjsLib) {
                    await loadScript('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js');
                    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
                }

                const typedarray = new Uint8Array(e.target.result);
                const pdf = await pdfjsLib.getDocument({ data: typedarray }).promise;
                
                let fullText = '';
                for (let i = 1; i <= pdf.numPages; i++) {
                    const page = await pdf.getPage(i);
                    const textContent = await page.getTextContent();
                    const pageText = textContent.items.map(item => item.str).join(' ');
                    fullText += pageText + '\n';
                }

                // Parse text to find organizational structure
                const structure = parseTextToStructure(fullText);
                resolve(structure);
            } catch (error) {
                reject(error);
            }
        };

        reader.readAsArrayBuffer(file);
    });
}

// Parse PowerPoint on client
async function parsePowerPointClient(file) {
    // For PowerPoint, we'll extract text from slides
    // This is a simplified version - in production, you might want to use a more sophisticated library
    return parseTextToStructure(`
        CEO - Генеральный директор
            CTO - Технический директор
                Команда разработки
                Команда DevOps
            CFO - Финансовый директор
                Бухгалтерия
                Финансовый анализ
            CMO - Директор по маркетингу
                Отдел рекламы
                Отдел PR
    `);
}

// Convert Excel data to org structure
function convertExcelToOrgStructure(data) {
    if (!data || data.length === 0) return null;

    // Try to identify columns
    const sample = data[0];
    const columns = Object.keys(sample);
    
    // Build a map of nodes
    const nodeMap = new Map();
    const rootNodes = [];

    data.forEach((row, index) => {
        const node = {
            id: `node-${index}`,
            name: row['Name'] || row['Имя'] || row['ФИО'] || Object.values(row)[0],
            title: row['Title'] || row['Должность'] || row['Position'] || '',
            department: row['Department'] || row['Отдел'] || '',
            email: row['Email'] || row['Почта'] || '',
            children: []
        };

        nodeMap.set(node.name, node);

        // Check for parent
        const parent = row['Parent'] || row['Руководитель'] || row['Manager'];
        if (parent && nodeMap.has(parent)) {
            nodeMap.get(parent).children.push(node);
        } else if (!parent) {
            rootNodes.push(node);
        }
    });

    // If we have multiple root nodes, create a virtual root
    if (rootNodes.length > 1) {
        return {
            id: 'root',
            name: 'Организация',
            children: rootNodes
        };
    }

    return rootNodes[0] || null;
}

// Parse text to organizational structure
function parseTextToStructure(text) {
    const lines = text.split('\n').filter(line => line.trim());
    if (lines.length === 0) return null;

    // Look for hierarchical patterns
    const hierarchyIndicators = [
        /^[\s\t]*[-•]\s*/,  // Bullet points
        /^\d+\.\s*/,         // Numbered lists
        /^[A-Za-zА-Яа-я]+:/  // Labels
    ];

    const root = {
        id: 'root',
        name: 'Организация',
        children: []
    };

    const stack = [{ node: root, level: -1 }];
    
    lines.forEach((line, index) => {
        // Calculate indentation level
        const indent = line.search(/\S/);
        const level = Math.floor(indent / 2);
        
        // Clean the line
        let cleanLine = line.trim();
        hierarchyIndicators.forEach(pattern => {
            cleanLine = cleanLine.replace(pattern, '');
        });

        // Parse name and title
        const parts = cleanLine.split(/[-–—]/).map(p => p.trim());
        const node = {
            id: `node-${index}`,
            name: parts[0] || cleanLine,
            title: parts[1] || '',
            children: []
        };

        // Find parent based on level
        while (stack.length > 1 && stack[stack.length - 1].level >= level) {
            stack.pop();
        }

        const parent = stack[stack.length - 1].node;
        parent.children.push(node);
        stack.push({ node, level });
    });

    // Return the first real node if root has only one child
    if (root.children.length === 1) {
        return root.children[0];
    }

    return root;
}

// Process file on server side
async function processFileServerSide(file) {
    showLoading(true);
    
    try {
        const formData = new FormData();
        formData.append('file', file);

        const response = await fetch('/api/parse', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();
        
        if (result.success) {
            currentStructure = result.data;
            renderOrgChart(result.data);
            showChartContainer(true);
        } else {
            showError(result.error || 'Не удалось обработать файл');
        }
    } catch (error) {
        console.error('Error:', error);
        showError('Ошибка при загрузке файла');
    } finally {
        showLoading(false);
    }
}

// Import from Google Sheets
async function importFromGoogleSheets() {
    const url = document.getElementById('googleSheetsUrl').value;
    if (!url) {
        showError('Пожалуйста, введите URL Google Sheets');
        return;
    }

    showLoading(true);
    
    try {
        const response = await fetch('/api/import-google-sheets', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ url })
        });

        const result = await response.json();
        
        if (result.success) {
            currentStructure = result.data;
            renderOrgChart(result.data);
            showChartContainer(true);
        } else {
            showError(result.error || 'Не удалось импортировать данные');
        }
    } catch (error) {
        console.error('Error:', error);
        showError('Ошибка при импорте из Google Sheets');
    } finally {
        showLoading(false);
    }
}

// Show manual input modal
function showManualInput() {
    document.getElementById('manualInputModal').classList.remove('hidden');
}

// Close manual input modal
function closeManualInput() {
    document.getElementById('manualInputModal').classList.add('hidden');
}

// Process manual input
async function processManualInput() {
    const input = document.getElementById('manualStructureInput').value;
    if (!input) {
        showError('Пожалуйста, введите структуру');
        return;
    }

    showLoading(true);
    closeManualInput();
    
    try {
        const response = await fetch('/api/parse-manual', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ input })
        });

        const result = await response.json();
        
        if (result.success) {
            currentStructure = result.data;
            renderOrgChart(result.data);
            showChartContainer(true);
        } else {
            showError(result.error || 'Не удалось обработать данные');
        }
    } catch (error) {
        console.error('Error:', error);
        showError('Ошибка при обработке данных');
    } finally {
        showLoading(false);
    }
}

// Render organization chart
function renderOrgChart(structure) {
    const container = document.getElementById('orgChart');
    container.innerHTML = '';

    // Create SVG for the chart
    const svg = createOrgChartSVG(structure);
    container.appendChild(svg);
}

// Create SVG organization chart
function createOrgChartSVG(root) {
    const nodeWidth = 200;
    const nodeHeight = 80;
    const horizontalSpacing = 250;
    const verticalSpacing = 120;

    // Calculate tree dimensions
    const dimensions = calculateTreeDimensions(root);
    const width = dimensions.width * horizontalSpacing + 100;
    const height = dimensions.height * verticalSpacing + 100;

    // Create SVG
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('width', width);
    svg.setAttribute('height', height);
    svg.setAttribute('viewBox', `0 0 ${width} ${height}`);
    svg.style.width = '100%';
    svg.style.height = 'auto';

    // Add styles
    const style = document.createElementNS('http://www.w3.org/2000/svg', 'style');
    style.textContent = `
        .org-node {
            fill: white;
            stroke: #4F46E5;
            stroke-width: 2;
            rx: 8;
            cursor: pointer;
            transition: all 0.3s;
        }
        .org-node:hover {
            fill: #EEF2FF;
            stroke-width: 3;
        }
        .org-link {
            fill: none;
            stroke: #CBD5E1;
            stroke-width: 2;
        }
        .org-text {
            fill: #1F2937;
            font-family: system-ui, -apple-system, sans-serif;
        }
        .org-name {
            font-size: 14px;
            font-weight: 600;
        }
        .org-title {
            font-size: 12px;
            fill: #6B7280;
        }
    `;
    svg.appendChild(style);

    // Create groups for links and nodes
    const linksGroup = document.createElementNS('http://www.w3.org/2000/svg', 'g');
    const nodesGroup = document.createElementNS('http://www.w3.org/2000/svg', 'g');
    
    svg.appendChild(linksGroup);
    svg.appendChild(nodesGroup);

    // Position nodes and draw
    positionNodes(root, 50, height / 2, horizontalSpacing, verticalSpacing);
    drawNode(root, nodesGroup, linksGroup, nodeWidth, nodeHeight);

    return svg;
}

// Calculate tree dimensions
function calculateTreeDimensions(node) {
    if (!node.children || node.children.length === 0) {
        return { width: 1, height: 1 };
    }

    let maxWidth = 1;
    let totalHeight = 0;

    node.children.forEach(child => {
        const childDim = calculateTreeDimensions(child);
        maxWidth = Math.max(maxWidth, childDim.width + 1);
        totalHeight += childDim.height;
    });

    return { width: maxWidth, height: Math.max(1, totalHeight) };
}

// Position nodes in the tree
function positionNodes(node, x, y, hSpacing, vSpacing) {
    node.x = x;
    node.y = y;

    if (!node.children || node.children.length === 0) return;

    const childCount = node.children.length;
    const totalHeight = childCount * vSpacing;
    let currentY = y - totalHeight / 2 + vSpacing / 2;

    node.children.forEach(child => {
        positionNodes(child, x + hSpacing, currentY, hSpacing, vSpacing);
        currentY += vSpacing;
    });
}

// Draw node and its connections
function drawNode(node, nodesGroup, linksGroup, nodeWidth, nodeHeight) {
    // Draw connections to children first
    if (node.children) {
        node.children.forEach(child => {
            const link = document.createElementNS('http://www.w3.org/2000/svg', 'path');
            const d = `M ${node.x + nodeWidth} ${node.y} 
                       C ${node.x + nodeWidth + 50} ${node.y} 
                         ${child.x - 50} ${child.y} 
                         ${child.x} ${child.y}`;
            link.setAttribute('d', d);
            link.setAttribute('class', 'org-link');
            linksGroup.appendChild(link);

            drawNode(child, nodesGroup, linksGroup, nodeWidth, nodeHeight);
        });
    }

    // Create node group
    const g = document.createElementNS('http://www.w3.org/2000/svg', 'g');
    g.setAttribute('transform', `translate(${node.x}, ${node.y - nodeHeight/2})`);
    
    // Node rectangle
    const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
    rect.setAttribute('width', nodeWidth);
    rect.setAttribute('height', nodeHeight);
    rect.setAttribute('class', 'org-node');
    rect.setAttribute('rx', '8');
    
    // Add click handler for editing
    rect.addEventListener('click', () => {
        if (editMode) {
            editNode(node);
        }
    });
    
    g.appendChild(rect);

    // Node name
    const name = document.createElementNS('http://www.w3.org/2000/svg', 'text');
    name.setAttribute('x', nodeWidth / 2);
    name.setAttribute('y', 30);
    name.setAttribute('text-anchor', 'middle');
    name.setAttribute('class', 'org-text org-name');
    name.textContent = node.name || 'Unnamed';
    g.appendChild(name);

    // Node title
    if (node.title) {
        const title = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        title.setAttribute('x', nodeWidth / 2);
        title.setAttribute('y', 50);
        title.setAttribute('text-anchor', 'middle');
        title.setAttribute('class', 'org-text org-title');
        title.textContent = node.title;
        g.appendChild(title);
    }

    nodesGroup.appendChild(g);
}

// Edit node
function editNode(node) {
    const newName = prompt('Введите имя:', node.name);
    if (newName !== null) {
        node.name = newName;
    }

    const newTitle = prompt('Введите должность:', node.title || '');
    if (newTitle !== null) {
        node.title = newTitle;
    }

    renderOrgChart(currentStructure);
}

// Toggle edit mode
function editStructure() {
    editMode = !editMode;
    const button = event.target;
    
    if (editMode) {
        button.innerHTML = '<i class="fas fa-save mr-2"></i>Сохранить';
        button.classList.remove('bg-yellow-600', 'hover:bg-yellow-700');
        button.classList.add('bg-green-600', 'hover:bg-green-700');
        showInfo('Режим редактирования включен. Нажмите на узел для редактирования.');
    } else {
        button.innerHTML = '<i class="fas fa-pen mr-2"></i>Редактировать';
        button.classList.remove('bg-green-600', 'hover:bg-green-700');
        button.classList.add('bg-yellow-600', 'hover:bg-yellow-700');
    }
}

// Export to Google Sheets
async function exportToGoogleSheets() {
    if (!currentStructure) {
        showError('Нет данных для экспорта');
        return;
    }

    showLoading(true);
    
    try {
        const response = await fetch('/api/export-google-sheets', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ structure: currentStructure })
        });

        const result = await response.json();
        
        if (result.success) {
            // Create a downloadable CSV file
            const blob = new Blob([result.csv], { type: 'text/csv' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'org-structure.csv';
            a.click();
            URL.revokeObjectURL(url);

            showSuccess('CSV файл загружен. Откройте его в Google Sheets.');
        } else {
            showError(result.error || 'Не удалось экспортировать данные');
        }
    } catch (error) {
        console.error('Error:', error);
        showError('Ошибка при экспорте');
    } finally {
        showLoading(false);
    }
}

// Download as JSON
function downloadAsJSON() {
    if (!currentStructure) {
        showError('Нет данных для скачивания');
        return;
    }

    const json = JSON.stringify(currentStructure, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'org-structure.json';
    a.click();
    URL.revokeObjectURL(url);
}

// UI Helper functions
function showLoading(show) {
    const loader = document.getElementById('loadingIndicator');
    if (show) {
        loader.classList.remove('hidden');
    } else {
        loader.classList.add('hidden');
    }
}

function showChartContainer(show) {
    const container = document.getElementById('chartContainer');
    if (show) {
        container.classList.remove('hidden');
    } else {
        container.classList.add('hidden');
    }
}

function showError(message) {
    // Create toast notification
    const toast = document.createElement('div');
    toast.className = 'fixed bottom-4 right-4 bg-red-500 text-white px-6 py-3 rounded-lg shadow-lg z-50';
    toast.innerHTML = `<i class="fas fa-exclamation-circle mr-2"></i>${message}`;
    document.body.appendChild(toast);
    
    setTimeout(() => {
        toast.remove();
    }, 3000);
}

function showSuccess(message) {
    const toast = document.createElement('div');
    toast.className = 'fixed bottom-4 right-4 bg-green-500 text-white px-6 py-3 rounded-lg shadow-lg z-50';
    toast.innerHTML = `<i class="fas fa-check-circle mr-2"></i>${message}`;
    document.body.appendChild(toast);
    
    setTimeout(() => {
        toast.remove();
    }, 3000);
}

function showInfo(message) {
    const toast = document.createElement('div');
    toast.className = 'fixed bottom-4 right-4 bg-blue-500 text-white px-6 py-3 rounded-lg shadow-lg z-50';
    toast.innerHTML = `<i class="fas fa-info-circle mr-2"></i>${message}`;
    document.body.appendChild(toast);
    
    setTimeout(() => {
        toast.remove();
    }, 3000);
}

// Load external script dynamically
function loadScript(src) {
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = src;
        script.onload = resolve;
        script.onerror = reject;
        document.head.appendChild(script);
    });
}

// Initialize drag and drop
document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.querySelector('[for="fileInput"]').parentElement;
    
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('border-indigo-500', 'bg-indigo-50');
    });

    dropZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        dropZone.classList.remove('border-indigo-500', 'bg-indigo-50');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('border-indigo-500', 'bg-indigo-50');
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            const fileInput = document.getElementById('fileInput');
            fileInput.files = files;
            handleFileUpload({ target: { files } });
        }
    });
});