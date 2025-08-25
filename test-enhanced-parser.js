// Test script for enhanced parser

// Test CSV data (simulating Google Sheets export)
const testCSV = `Подчиненность,Подразделение,Должность,ФИО,Функционал
1,Руководство,Генеральный директор,Иванов Иван Иванович,"Общее руководство компанией, стратегическое планирование"
1.1,Руководство,Заместитель генерального директора,Петров Петр Петрович,"Оперативное управление, координация подразделений"
1.1.1,ИТ отдел,Начальник ИТ отдела,Сидоров Сидор Сидорович,"Управление ИТ инфраструктурой, разработка"
1.1.1.1,ИТ отдел,Ведущий разработчик,Козлов Козел Козлович,"Разработка программного обеспечения, архитектура"
1.1.1.2,ИТ отдел,Системный администратор,Волков Волк Волкович,"Администрирование серверов, поддержка"
1.1.2,Отдел продаж,Начальник отдела продаж,Лисов Лис Лисович,"Управление продажами, работа с клиентами"
1.1.2.1,Отдел продаж,Менеджер по продажам,Зайцев Заяц Зайцевич,"Прямые продажи, работа с клиентской базой"
1.1.2.2,Отдел продаж,Менеджер по работе с ключевыми клиентами,Медведев Медведь Медведевич,"Работа с VIP клиентами, долгосрочные контракты"
1.2,Финансовый отдел,Главный бухгалтер,Орлова Ольга Олеговна,"Ведение бухгалтерского учета, финансовая отчетность"
1.2.1,Финансовый отдел,Бухгалтер,Соколова Светлана Сергеевна,"Первичная документация, расчет зарплаты"`;

// Parse CSV line handling quotes
function parseCSVLine(line) {
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
            // End of field
            result.push(current.trim());
            current = '';
        } else {
            current += char;
        }
    }
    
    // Add last field
    if (current || line.endsWith(',')) {
        result.push(current.trim());
    }
    
    return result;
}

// Parse the test CSV
const lines = testCSV.split('\n').filter(line => line.trim());
console.log(`Total lines: ${lines.length}`);

const rows = [];
for (let line of lines) {
    const row = parseCSVLine(line);
    if (row && row.length > 0) {
        rows.push(row);
    }
}

console.log(`Parsed rows: ${rows.length}`);
console.log('\nFirst 5 data rows:');
for (let i = 1; i <= Math.min(5, rows.length - 1); i++) {
    const row = rows[i];
    console.log(`${i}. ${row[0]} | ${row[2]} | ${row[3]}`);
}

// Build hierarchy from subordination codes
function buildHierarchy(rows) {
    const nodeMap = new Map();
    const rootNodes = [];
    
    // Skip header row
    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const node = {
            id: `node-${i}`,
            subordination: row[0] || '',
            department: row[1] || '',
            position: row[2] || '',
            name: row[3] || '',
            functional: row[4] || '',
            children: []
        };
        
        nodeMap.set(node.subordination, node);
    }
    
    // Build parent-child relationships
    for (const [code, node] of nodeMap) {
        if (code && code.includes('.')) {
            // Find parent by removing last segment
            const parts = code.split('.');
            parts.pop();
            const parentCode = parts.join('.');
            
            if (nodeMap.has(parentCode)) {
                nodeMap.get(parentCode).children.push(node);
            } else {
                rootNodes.push(node);
            }
        } else {
            // Top level node
            rootNodes.push(node);
        }
    }
    
    return rootNodes.length === 1 ? rootNodes[0] : {
        id: 'root',
        name: 'Организация',
        children: rootNodes
    };
}

const hierarchy = buildHierarchy(rows);
console.log('\nHierarchy structure:');
console.log(JSON.stringify(hierarchy, null, 2));

// Count total employees
function countEmployees(node) {
    let count = 1;
    if (node.children && node.children.length > 0) {
        for (const child of node.children) {
            count += countEmployees(child);
        }
    }
    return count;
}

const totalEmployees = countEmployees(hierarchy);
console.log(`\nTotal employees in hierarchy: ${totalEmployees}`);