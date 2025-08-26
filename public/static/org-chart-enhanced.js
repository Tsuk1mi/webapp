// Enhanced Organization Chart Application
let currentStructure = null;
let editMode = false;
let orgChart = null;
let professionalParser = null; // Professional parser
let tableData = [];

// Initialize on page load
document.addEventListener('DOMContentLoaded', () => {
    // Load D3.js
    loadScript('https://d3js.org/d3.v7.min.js').then(() => {
        console.log('D3.js loaded');
        // Initialize D3 chart after D3 is loaded
        orgChart = new D3OrgChart('orgChart');
    });
    
    // Initialize professional parser
    if (typeof ProfessionalOrgParser !== 'undefined') {
        professionalParser = new ProfessionalOrgParser();
        console.log('Professional parser initialized');
    } else {
        console.error('Professional parser not found!');
    }
    
    // Initialize drag and drop
    initializeDragAndDrop();
    
    // Add sample data button
    addSampleDataButton();
});

// Process real document URLs
async function processDocumentFromURL(url, fileName) {
    showLoading(true);
    
    try {
        // Fetch document content
        const response = await fetch('/api/process-document-url', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ url, fileName })
        });
        
        const result = await response.json();
        
        if (result.success && result.structure) {
            currentStructure = result.structure;
            renderEnhancedChart(result.structure);
            generateTable(result.structure);
            showChartContainer(true);
            showSuccess('Документ успешно обработан!');
        } else {
            showError('Не удалось обработать документ');
        }
    } catch (error) {
        console.error('Error processing document:', error);
        showError('Ошибка при обработке документа: ' + error.message);
    } finally {
        showLoading(false);
    }
}

// Enhanced file upload handler
async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    showLoading(true);
    
    try {
        let structure = null;
        
        // Determine file type
        const fileType = file.name.endsWith('.pptx') ? 'presentation' :
                        file.name.endsWith('.pdf') ? 'pdf' :
                        file.name.endsWith('.docx') ? 'word' :
                        file.name.endsWith('.xlsx') ? 'spreadsheet' : 'text';
        
        // Use professional parser for all file types
        if (professionalParser) {
            console.log(`Using professional parser for ${fileType}`);
            const text = await readFileAsText(file);
            structure = await professionalParser.parseDocument(text, fileType);
        } else {
            console.error('Professional parser not available');
            showError('Парсер не инициализирован');
            showLoading(false);
            return;
        }
        
        if (structure) {
            currentStructure = structure;
            renderEnhancedChart(structure);
            
            // Use professional table generation
            if (professionalParser) {
                const tableHtml = professionalParser.createHTMLTable(structure);
                displayEnhancedTable(tableHtml);
            } else {
                generateEnhancedTable(structure);
            }
            
            showChartContainer(true);
            
            // Count employees and vacancies
            const employeeCount = countEmployees(structure);
            const vacancyCount = countVacancies(structure);
            showSuccess(`Файл обработан! Сотрудников: ${employeeCount}, Вакансий: ${vacancyCount}`);
        } else {
            showError('Не удалось извлечь структуру из файла');
        }
    } catch (error) {
        console.error('Error processing file:', error);
        showError('Ошибка при обработке файла: ' + error.message);
    } finally {
        showLoading(false);
    }
}

// Count nodes in structure
function countNodes(node) {
    let count = 1;
    if (node.children) {
        node.children.forEach(child => {
            count += countNodes(child);
        });
    }
    return count;
}

// Read file as text with proper encoding
async function readFileAsText(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = async (e) => {
            try {
                let text = '';
                
                if (file.name.endsWith('.xlsx')) {
                    // Handle Excel
                    if (!window.XLSX) {
                        await loadScript('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');
                    }
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // Convert all sheets to text
                    workbook.SheetNames.forEach(sheetName => {
                        const sheet = workbook.Sheets[sheetName];
                        text += XLSX.utils.sheet_to_txt(sheet) + '\n\n';
                    });
                } else if (file.name.endsWith('.docx')) {
                    // Handle Word
                    if (!window.mammoth) {
                        await loadScript('https://cdn.jsdelivr.net/npm/mammoth@1.6.0/mammoth.browser.min.js');
                    }
                    const result = await mammoth.extractRawText({ arrayBuffer: e.target.result });
                    text = result.value;
                } else if (file.name.endsWith('.pdf')) {
                    // Handle PDF
                    if (!window.pdfjsLib) {
                        await loadScript('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js');
                        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
                    }
                    const typedarray = new Uint8Array(e.target.result);
                    const pdf = await pdfjsLib.getDocument({ data: typedarray }).promise;
                    
                    for (let i = 1; i <= pdf.numPages; i++) {
                        const page = await pdf.getPage(i);
                        const textContent = await page.getTextContent();
                        const pageText = textContent.items.map(item => item.str).join(' ');
                        text += pageText + '\n';
                    }
                } else if (file.name.endsWith('.pptx')) {
                    // Handle PowerPoint with specialized parser
                    const pptxStructure = await pptxParser.parsePPTX(file);
                    resolve(pptxStructure);
                    return;
                } else {
                    // Plain text
                    text = e.target.result;
                }
                
                resolve(text);
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = reject;
        
        // Read file based on type
        if (file.type.includes('text') || file.name.endsWith('.csv')) {
            reader.readAsText(file);
        } else {
            reader.readAsArrayBuffer(file);
        }
    });
}

// Render enhanced chart with D3
function renderEnhancedChart(structure) {
    if (orgChart) {
        orgChart.render(structure);
    } else {
        // Fallback to simple rendering
        renderSimpleChart(structure);
    }
}

// Simple fallback chart rendering
function renderSimpleChart(structure) {
    const container = document.getElementById('orgChart');
    container.innerHTML = '<div class="p-8 text-center">Визуализация загружается...</div>';
    
    // Simple tree view
    const treeHTML = buildTreeHTML(structure);
    container.innerHTML = treeHTML;
}

// Build simple HTML tree
function buildTreeHTML(node, level = 0) {
    const indent = level * 30;
    let html = `
        <div class="tree-node" style="margin-left: ${indent}px; margin-bottom: 10px;">
            <div class="p-3 bg-white rounded-lg shadow border-l-4 border-blue-500">
                <div class="font-semibold text-gray-800">${node.name || 'Без имени'}</div>
                ${node.title ? `<div class="text-sm text-gray-600">${node.title}</div>` : ''}
                ${node.email ? `<div class="text-xs text-gray-500"><i class="fas fa-envelope"></i> ${node.email}</div>` : ''}
            </div>
        </div>
    `;
    
    if (node.children && node.children.length > 0) {
        node.children.forEach(child => {
            html += buildTreeHTML(child, level + 1);
        });
    }
    
    return html;
}

// Generate table from structure
function generateTable(structure) {
    if (!advancedParser) return;
    
    tableData = advancedParser.toTableFormat(structure);
    
    // Create table container
    let tableContainer = document.getElementById('tableContainer');
    if (!tableContainer) {
        tableContainer = document.createElement('div');
        tableContainer.id = 'tableContainer';
        tableContainer.className = 'mt-8 bg-white rounded-lg shadow-lg p-6';
        document.getElementById('chartContainer').appendChild(tableContainer);
    }
    
    // Build table HTML
    let tableHTML = `
        <h3 class="text-xl font-semibold mb-4">
            <i class="fas fa-table mr-2"></i>Табличное представление
        </h3>
        <div class="overflow-x-auto">
            <table class="min-w-full divide-y divide-gray-200">
                <thead class="bg-gray-50">
                    <tr>
                        <th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Уровень</th>
                        <th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">ФИО</th>
                        <th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Должность</th>
                        <th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Отдел</th>
                        <th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Email</th>
                        <th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Руководитель</th>
                    </tr>
                </thead>
                <tbody class="bg-white divide-y divide-gray-200">
    `;
    
    tableData.forEach((row, index) => {
        const levelColor = getLevelColor(row.level);
        tableHTML += `
            <tr class="hover:bg-gray-50 transition">
                <td class="px-4 py-2 whitespace-nowrap">
                    <span class="px-2 py-1 text-xs rounded-full" style="background-color: ${levelColor}20; color: ${levelColor};">
                        ${row.level}
                    </span>
                </td>
                <td class="px-4 py-2 font-medium text-gray-900">${row.name}</td>
                <td class="px-4 py-2 text-sm text-gray-600">${row.title}</td>
                <td class="px-4 py-2 text-sm text-gray-600">${row.department || '-'}</td>
                <td class="px-4 py-2 text-sm text-gray-600">${row.email || '-'}</td>
                <td class="px-4 py-2 text-sm text-gray-600">${row.parent || '-'}</td>
            </tr>
        `;
    });
    
    tableHTML += `
                </tbody>
            </table>
        </div>
        <div class="mt-4 flex gap-2">
            <button onclick="exportTableToCSV()" class="px-4 py-2 bg-green-600 text-white rounded hover:bg-green-700">
                <i class="fas fa-file-csv mr-2"></i>Экспорт в CSV
            </button>
            <button onclick="copyTableToClipboard()" class="px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700">
                <i class="fas fa-copy mr-2"></i>Копировать таблицу
            </button>
        </div>
    `;
    
    tableContainer.innerHTML = tableHTML;
}

// Get color for hierarchy level
function getLevelColor(level) {
    const colors = ['#4F46E5', '#7C3AED', '#2563EB', '#0891B2', '#059669', '#DC2626'];
    return colors[Math.min(level, colors.length - 1)];
}

// Export table to CSV
function exportTableToCSV() {
    if (!tableData || tableData.length === 0) {
        showError('Нет данных для экспорта');
        return;
    }
    
    // Create CSV content
    const headers = ['Уровень', 'ФИО', 'Должность', 'Отдел', 'Email', 'Руководитель', 'Обязанности'];
    const rows = [headers];
    
    tableData.forEach(row => {
        rows.push([
            row.level,
            row.name,
            row.title,
            row.department || '',
            row.email || '',
            row.parent || '',
            row.responsibilities || ''
        ]);
    });
    
    // Convert to CSV string
    const csvContent = rows.map(row => 
        row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')
    ).join('\n');
    
    // Add BOM for Excel UTF-8 support
    const BOM = '\uFEFF';
    const blob = new Blob([BOM + csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'org-structure.csv';
    a.click();
    URL.revokeObjectURL(url);
    
    showSuccess('Таблица экспортирована в CSV');
}

// Copy table to clipboard
function copyTableToClipboard() {
    if (!tableData || tableData.length === 0) {
        showError('Нет данных для копирования');
        return;
    }
    
    // Create tab-separated content for pasting into Excel/Google Sheets
    const headers = ['Уровень', 'ФИО', 'Должность', 'Отдел', 'Email', 'Руководитель'];
    let content = headers.join('\t') + '\n';
    
    tableData.forEach(row => {
        content += [
            row.level,
            row.name,
            row.title,
            row.department || '',
            row.email || '',
            row.parent || ''
        ].join('\t') + '\n';
    });
    
    // Copy to clipboard
    navigator.clipboard.writeText(content).then(() => {
        showSuccess('Таблица скопирована в буфер обмена');
    }).catch(err => {
        showError('Не удалось скопировать таблицу');
    });
}

// Import from Google Sheets enhanced
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
            // Parse CSV from Google Sheets with professional parser
            if (!professionalParser) {
                showError('Парсер не инициализирован');
                return;
            }
            const structure = await professionalParser.parseSpreadsheet(result.csv || '');
            if (!structure) {
                showError('Не удалось распарсить данные из Google Sheets');
                return;
            }
            currentStructure = structure;
            renderEnhancedChart(structure);
            // Build enhanced table with department and position
            const tableHtml = professionalParser.createHTMLTable(structure);
            displayEnhancedTable(tableHtml);
            showChartContainer(true);
            showSuccess('Данные успешно импортированы из Google Sheets');
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

// Add sample data button
function addSampleDataButton() {
    const uploadSection = document.querySelector('.bg-white.rounded-lg.shadow-lg.p-8');
    if (uploadSection) {
        const sampleButton = document.createElement('div');
        sampleButton.className = 'mb-6';
        sampleButton.innerHTML = `
            <button onclick="loadSampleData()" class="w-full px-6 py-3 bg-purple-600 text-white rounded-lg hover:bg-purple-700 transition-colors">
                <i class="fas fa-database mr-2"></i>
                Загрузить пример структуры
            </button>
        `;
        uploadSection.appendChild(sampleButton);
    }
}

// Add document processing buttons
function addDocumentProcessingButtons() {
    const uploadSection = document.querySelector('.bg-white.rounded-lg.shadow-lg.p-8');
    if (uploadSection) {
        const docButtons = document.createElement('div');
        docButtons.className = 'mb-6 border-t pt-4';
        docButtons.innerHTML = `
            <h3 class="text-lg font-medium mb-3 text-gray-700">
                <i class="fas fa-file-alt mr-2"></i>
                Обработать документы из примеров
            </h3>
            <div class="grid grid-cols-2 gap-2">
                <button onclick="processExampleDoc('61f9ebde')" class="px-3 py-2 bg-blue-500 text-white rounded text-sm hover:bg-blue-600">
                    Отдел контроля решений
                </button>
                <button onclick="processExampleDoc('dd3b3bec')" class="px-3 py-2 bg-blue-500 text-white rounded text-sm hover:bg-blue-600">
                    Отдел контроля данных
                </button>
                <button onclick="processExampleDoc('f0c77077')" class="px-3 py-2 bg-green-500 text-white rounded text-sm hover:bg-green-600">
                    Управление (PDF)
                </button>
                <button onclick="processExampleDoc('1f77fe1b')" class="px-3 py-2 bg-orange-500 text-white rounded text-sm hover:bg-orange-600">
                    Презентация (PPTX)
                </button>
                <button onclick="processGoogleSheetsExample()" class="col-span-2 px-3 py-2 bg-red-500 text-white rounded text-sm hover:bg-red-600">
                    <i class="fab fa-google mr-2"></i>
                    Таблица Google Sheets (100+ строк)
                </button>
            </div>
        `;
        uploadSection.appendChild(docButtons);
    }
}

// Load sample data
function loadSampleData() {
    const sampleStructure = {
        name: "Генеральный директор",
        title: "CEO",
        email: "ceo@company.ru",
        children: [
            {
                name: "Тлигуров Юрий Арсенович",
                title: "Заместитель начальника управления",
                children: [
                    {
                        name: "Васянин Станислав Александрович",
                        title: "Начальник отдела паспортизации",
                        children: [
                            {
                                name: "Геращенко Елена Викторовна",
                                title: "Заместитель начальника отдела"
                            }
                        ]
                    },
                    {
                        name: "Черкесова Светлана Леонидовна",
                        title: "Начальник отдела контроля решений",
                        children: [
                            {
                                name: "Глушкова Алёна Сергеевна",
                                title: "Заместитель начальника отдела"
                            }
                        ]
                    }
                ]
            },
            {
                name: "Черкасов Виталий Александрович",
                title: "Заместитель начальника управления",
                children: [
                    {
                        name: "Бредина Татьяна",
                        title: "Зам. начальника отдела"
                    }
                ]
            },
            {
                name: "Толмачёва Ольга Леонидовна",
                title: "Начальник отдела контроля пространственных данных",
                responsibilities: [
                    "Разработка методических материалов",
                    "Автоматизация сервиса ФЛК",
                    "Контроль качества данных"
                ]
            }
        ]
    };
    
    currentStructure = sampleStructure;
    renderEnhancedChart(sampleStructure);
    generateTable(sampleStructure);
    showChartContainer(true);
    showSuccess('Пример структуры загружен');
}

// Initialize drag and drop
function initializeDragAndDrop() {
    const dropZone = document.querySelector('[for="fileInput"]')?.parentElement;
    if (!dropZone) return;
    
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
}

// UI Helper functions
function showLoading(show) {
    const loader = document.getElementById('loadingIndicator');
    if (loader) {
        if (show) {
            loader.classList.remove('hidden');
        } else {
            loader.classList.add('hidden');
        }
    }
}

function showChartContainer(show) {
    const container = document.getElementById('chartContainer');
    if (container) {
        if (show) {
            container.classList.remove('hidden');
        } else {
            container.classList.add('hidden');
        }
    }
}

function showError(message) {
    const toast = document.createElement('div');
    toast.className = 'fixed bottom-4 right-4 bg-red-500 text-white px-6 py-3 rounded-lg shadow-lg z-50 animate-fadeIn';
    toast.innerHTML = `<i class="fas fa-exclamation-circle mr-2"></i>${message}`;
    document.body.appendChild(toast);
    
    setTimeout(() => {
        toast.remove();
    }, 3000);
}

function showSuccess(message) {
    const toast = document.createElement('div');
    toast.className = 'fixed bottom-4 right-4 bg-green-500 text-white px-6 py-3 rounded-lg shadow-lg z-50 animate-fadeIn';
    toast.innerHTML = `<i class="fas fa-check-circle mr-2"></i>${message}`;
    document.body.appendChild(toast);
    
    setTimeout(() => {
        toast.remove();
    }, 3000);
}

function showInfo(message) {
    const toast = document.createElement('div');
    toast.className = 'fixed bottom-4 right-4 bg-blue-500 text-white px-6 py-3 rounded-lg shadow-lg z-50 animate-fadeIn';
    toast.innerHTML = `<i class="fas fa-info-circle mr-2"></i>${message}`;
    document.body.appendChild(toast);
    
    setTimeout(() => {
        toast.remove();
    }, 3000);
}

// Load external script dynamically
function loadScript(src) {
    return new Promise((resolve, reject) => {
        // Check if already loaded
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

// Process documents from URL (keeping for compatibility)
async function processExampleDoc(docId) {
    showLoading(true);
    
    try {
        // Map document IDs to URLs (updated with new documents)
        const docUrls = {
            '61f9ebde': 'https://page.gensparksite.com/get_upload_url/2b0155689b38801236ac840f06b2d9e3f1d807f043cd2485fc7a6758ed5dfabc/default/2689ee56-e740-46b1-a9ab-8b546112e6ba',
            'dd3b3bec': 'https://page.gensparksite.com/get_upload_url/2b0155689b38801236ac840f06b2d9e3f1d807f043cd2485fc7a6758ed5dfabc/default/5f12648d-488f-4c13-abf2-1cc804947851',
            'f0c77077': 'https://page.gensparksite.com/get_upload_url/2b0155689b38801236ac840f06b2d9e3f1d807f043cd2485fc7a6758ed5dfabc/default/3795a7d8-f5a7-4cd4-b79a-bb667e04b385',
            '1f77fe1b': 'https://page.gensparksite.com/get_upload_url/2b0155689b38801236ac840f06b2d9e3f1d807f043cd2485fc7a6758ed5dfabc/default/fb5adeb7-fa7d-4454-a82c-eab8c5f38d0c',
            'c5eece3b': 'https://page.gensparksite.com/get_upload_url/2b0155689b38801236ac840f06b2d9e3f1d807f043cd2485fc7a6758ed5dfabc/default/98f1d80c-9d47-48a2-a989-a898fd10871f',
            'b395e86d': 'https://page.gensparksite.com/get_upload_url/2b0155689b38801236ac840f06b2d9e3f1d807f043cd2485fc7a6758ed5dfabc/default/6693abd0-7f7c-4d3d-b6db-86a6d8de078c'
        };
        
        const url = docUrls[docId];
        if (!url) {
            showError('Документ не найден');
            return;
        }
        
        let structure = null;
        const response = await fetch(url);
        
        // Determine file type
        const fileType = docId === '1f77fe1b' ? 'presentation' : 
                        (docId === 'f0c77077' || docId === 'c5eece3b' || docId === 'b395e86d') ? 'pdf' : 
                        'word';
        
        // Use corporate parser for all documents
        if (corporateParser) {
            console.log(`Using corporate parser for ${fileType} document`);
            const text = await response.text();
            structure = await corporateParser.parseDocument(text, fileType);
        }
        
        // Fallback to other parsers if corporate parser fails
        if (!structure && docId === '1f77fe1b') {
            // Try PPTX parser for presentation
            const blob = await response.blob();
            const file = new File([blob], 'document.pptx', { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
            if (enhancedParser) {
                structure = await enhancedParser.parsePPTXDocument(file);
            }
            if (!structure) {
                structure = await pptxParser.parsePPTXFromURL(url);
            }
        } else if (!structure) {
            // Try specialized parser
            const text = await response.text();
            structure = specializedParser.parseComplexDocument(text, url);
        }
        
        if (structure) {
            currentStructure = structure;
            renderEnhancedChart(structure);
            generateEnhancedTable(structure);
            showChartContainer(true);
            showSuccess('Документ успешно обработан!');
        } else {
            showError('Не удалось обработать документ');
        }
    } catch (error) {
        console.error('Error processing document:', error);
        showError('Ошибка при обработке документа');
    } finally {
        showLoading(false);
    }
}

// Process Google Sheets example
async function processGoogleSheetsExample() {
    showLoading(true);
    
    try {
        // Google Sheets public URL
        const sheetId = '15o26CMkMG_73ZhrlwlyR06vH2BDYhfu1DYkelcG_C44';
        const exportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv`;
        
        const response = await fetch('/api/import-google-sheets', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ url: `https://docs.google.com/spreadsheets/d/${sheetId}/edit` })
        });
        
        const result = await response.json();
        
        if (result.success) {
            let structure = null;
            
            // Use professional parser for Google Sheets
            if (professionalParser) {
                console.log('Using professional parser for Google Sheets');
                structure = await professionalParser.parseSpreadsheet(result.csv || '');
            } else {
                console.error('Professional parser not available');
                showError('Парсер не инициализирован');
            }
            
            if (structure) {
                currentStructure = structure;
                renderEnhancedChart(structure);
                
                // Use professional table generation
                if (professionalParser) {
                    const tableHtml = professionalParser.createHTMLTable(structure);
                    displayEnhancedTable(tableHtml);
                } else {
                    generateEnhancedTable(structure);
                }
                
                showChartContainer(true);
                
                // Count actual employees and vacancies
                const employeeCount = countEmployees(structure);
                const vacancyCount = countVacancies(structure);
                showSuccess(`Google Sheets успешно импортирован! Позиций: ${employeeCount}, Вакансий: ${vacancyCount}`);
            }
        } else {
            showError('Не удалось импортировать Google Sheets');
        }
    } catch (error) {
        console.error('Error:', error);
        showError('Ошибка при импорте Google Sheets');
    } finally {
        showLoading(false);
    }
}

// Helper function to count employees in structure
function countEmployees(node) {
    let count = 1; // Count current node
    if (node.children && node.children.length > 0) {
        for (const child of node.children) {
            count += countEmployees(child);
        }
    }
    return count;
}

// Display enhanced table from HTML string
function displayEnhancedTable(tableHtml) {
    let tableContainer = document.getElementById('tableContainer');
    if (!tableContainer) {
        tableContainer = document.createElement('div');
        tableContainer.id = 'tableContainer';
        tableContainer.className = 'mt-8 bg-white rounded-lg shadow-lg p-6';
        const chartContainer = document.getElementById('chartContainer');
        if (chartContainer) {
            chartContainer.appendChild(tableContainer);
        }
    }
    
    tableContainer.innerHTML = tableHtml;
}

// Generate enhanced table with more columns
function generateEnhancedTable(structure) {
    if (!specializedParser) return;
    
    tableData = specializedParser.hierarchyToTable(structure);
    
    // Create table container
    let tableContainer = document.getElementById('tableContainer');
    if (!tableContainer) {
        tableContainer = document.createElement('div');
        tableContainer.id = 'tableContainer';
        tableContainer.className = 'mt-8 bg-white rounded-lg shadow-lg p-6';
        document.getElementById('chartContainer').appendChild(tableContainer);
    }
    
    // Build enhanced table HTML
    let tableHTML = `
        <h3 class="text-xl font-semibold mb-4">
            <i class="fas fa-table mr-2"></i>Табличное представление (${tableData.length} записей)
        </h3>
        <div class="mb-4 flex gap-2">
            <button onclick="exportToEnhancedCSV()" class="px-4 py-2 bg-green-600 text-white rounded hover:bg-green-700">
                <i class="fas fa-file-csv mr-2"></i>Экспорт в CSV
            </button>
            <button onclick="exportToGoogleSheetsFormat()" class="px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700">
                <i class="fab fa-google mr-2"></i>Формат для Google Sheets
            </button>
            <button onclick="copyEnhancedTable()" class="px-4 py-2 bg-purple-600 text-white rounded hover:bg-purple-700">
                <i class="fas fa-copy mr-2"></i>Копировать таблицу
            </button>
        </div>
        <div class="overflow-x-auto max-h-96 overflow-y-auto">
            <table class="min-w-full divide-y divide-gray-200">
                <thead class="bg-gray-50 sticky top-0">
                    <tr>
                        <th class="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase">Подчиненность</th>
                        <th class="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase">Подразделение</th>
                        <th class="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase">Должность</th>
                        <th class="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase">ФИО</th>
                        <th class="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase">Функционал</th>
                        <th class="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase">Уровень</th>
                    </tr>
                </thead>
                <tbody class="bg-white divide-y divide-gray-200">
    `;
    
    tableData.forEach((row, index) => {
        const levelColor = getLevelColor(row['Уровень'] || 0);
        const truncatedFunc = row['Функционал'] ? 
            (row['Функционал'].length > 100 ? 
                row['Функционал'].substring(0, 100) + '...' : 
                row['Функционал']) : '-';
        
        tableHTML += `
            <tr class="hover:bg-gray-50 transition text-sm">
                <td class="px-3 py-2 whitespace-nowrap font-mono text-xs">${row['Подчиненность'] || ''}</td>
                <td class="px-3 py-2">${row['Подразделение'] || '-'}</td>
                <td class="px-3 py-2">${row['Должность'] || '-'}</td>
                <td class="px-3 py-2 font-medium">${row['ФИО'] || '-'}</td>
                <td class="px-3 py-2 text-xs" title="${row['Функционал'] || ''}">${truncatedFunc}</td>
                <td class="px-3 py-2 text-center">
                    <span class="px-2 py-1 text-xs rounded-full" style="background-color: ${levelColor}20; color: ${levelColor};">
                        ${row['Уровень'] || 0}
                    </span>
                </td>
            </tr>
        `;
    });
    
    tableHTML += `
                </tbody>
            </table>
        </div>
        <div class="mt-4 text-sm text-gray-600">
            <i class="fas fa-info-circle mr-1"></i>
            Всего записей: ${tableData.length} | 
            Уникальных подразделений: ${new Set(tableData.map(r => r['Подразделение']).filter(d => d)).size} |
            Сотрудников с ФИО: ${tableData.filter(r => r['ФИО'] && r['ФИО'] !== '-').length}
        </div>
    `;
    
    tableContainer.innerHTML = tableHTML;
}

// Export to enhanced CSV
function exportToEnhancedCSV() {
    if (!currentStructure) {
        showError('Нет данных для экспорта');
        return;
    }
    
    let csv = '';
    
    if (professionalParser) {
        const excelData = professionalParser.generateExcelTable(currentStructure);
        // Convert to CSV
        csv = excelData.map(row => row.map(cell => 
            typeof cell === 'string' && cell.includes(',') ? `"${cell}"` : cell
        ).join(',')).join('\n');
    } else {
        showError('Парсер не доступен');
        return;
    }
    
    // Add BOM for Excel UTF-8 support
    const BOM = '\uFEFF';
    const blob = new Blob([BOM + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'org-structure-enhanced.csv';
    a.click();
    URL.revokeObjectURL(url);
    
    showSuccess('Таблица экспортирована в CSV');
}

// Export to Google Sheets format
function exportToGoogleSheetsFormat() {
    if (!currentStructure) {
        showError('Нет данных для экспорта');
        return;
    }
    
    if (!professionalParser) {
        showError('Парсер не доступен');
        return;
    }
    
    const excelData = professionalParser.generateExcelTable(currentStructure);
    
    // Convert to tab-separated for easy paste
    let content = excelData.map(row => row.join('\t')).join('\n');
    
    // Copy to clipboard
    navigator.clipboard.writeText(content).then(() => {
        showSuccess('Данные скопированы в формате для Google Sheets. Вставьте в таблицу с помощью Ctrl+V');
    }).catch(err => {
        // Fallback - download as TSV
        const blob = new Blob([content], { type: 'text/tab-separated-values;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'org-structure-google.tsv';
        a.click();
        URL.revokeObjectURL(url);
        showSuccess('Файл для Google Sheets сохранен');
    });
}

// Copy enhanced table
function copyEnhancedTable() {
    if (!tableData || tableData.length === 0) {
        showError('Нет данных для копирования');
        return;
    }
    
    // Create tab-separated content
    const headers = ['Подчиненность', 'Подразделение', 'Должность', 'ФИО', 'Функционал', 'Уровень'];
    let content = headers.join('\t') + '\n';
    
    tableData.forEach(row => {
        content += [
            row['Подчиненность'] || '',
            row['Подразделение'] || '',
            row['Должность'] || '',
            row['ФИО'] || '',
            row['Функционал'] || '',
            row['Уровень'] || ''
        ].join('\t') + '\n';
    });
    
    navigator.clipboard.writeText(content).then(() => {
        showSuccess(`Таблица скопирована (${tableData.length} строк)`);
    }).catch(err => {
        showError('Не удалось скопировать таблицу');
    });
}

// Export functions for global use
window.handleFileUpload = handleFileUpload;
window.importFromGoogleSheets = importFromGoogleSheets;
window.exportTableToCSV = exportTableToCSV;
window.copyTableToClipboard = copyTableToClipboard;
window.loadSampleData = loadSampleData;
window.processExampleDoc = processExampleDoc;
window.processGoogleSheetsExample = processGoogleSheetsExample;
window.exportToEnhancedCSV = exportToEnhancedCSV;
window.exportToGoogleSheetsFormat = exportToGoogleSheetsFormat;
window.copyEnhancedTable = copyEnhancedTable;
window.showManualInput = () => document.getElementById('manualInputModal')?.classList.remove('hidden');
window.closeManualInput = () => document.getElementById('manualInputModal')?.classList.add('hidden');
window.processManualInput = async () => {
    const input = document.getElementById('manualStructureInput')?.value;
    if (!input) {
        showError('Пожалуйста, введите структуру');
        return;
    }
    
    try {
        const structure = JSON.parse(input);
        currentStructure = structure;
        renderEnhancedChart(structure);
        generateTable(structure);
        showChartContainer(true);
        closeManualInput();
        showSuccess('Структура успешно создана');
    } catch (error) {
        // Try parsing as text
        const structure = await advancedParser.parseDocument(input, 'text');
        if (structure) {
            currentStructure = structure;
            renderEnhancedChart(structure);
            generateTable(structure);
            showChartContainer(true);
            closeManualInput();
            showSuccess('Структура успешно создана');
        } else {
            showError('Неверный формат данных');
        }
    }
};