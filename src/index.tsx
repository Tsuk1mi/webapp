import { Hono } from 'hono'
import { cors } from 'hono/cors'
import { serveStatic } from 'hono/cloudflare-workers'
import { renderer } from './renderer'

// Types
interface OrgNode {
  id: string
  name: string
  title?: string
  department?: string
  email?: string
  children: OrgNode[]
}

interface ParseResult {
  success: boolean
  data?: OrgNode
  error?: string
}

const app = new Hono()

// Enable CORS
app.use('/api/*', cors())

// Serve static files
app.use('/static/*', serveStatic({ root: './public' }))

// Home page
app.get('/', renderer, (c) => {
  return c.render(
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100">
      <div className="container mx-auto px-4 py-8">
        <header className="text-center mb-12">
          <h1 className="text-4xl font-bold text-gray-800 mb-4">
            <i className="fas fa-sitemap mr-3"></i>
            Организационная структура
          </h1>
          <p className="text-gray-600 text-lg">
            Загрузите документ и постройте интерактивную организационную диаграмму
          </p>
        </header>

        {/* Upload Section */}
        <div className="max-w-4xl mx-auto mb-12">
          <div className="bg-white rounded-lg shadow-lg p-8">
            <h2 className="text-2xl font-semibold mb-6 text-gray-700">
              <i className="fas fa-upload mr-2"></i>
              Загрузить документ
            </h2>
            
            {/* File Upload */}
            <div className="mb-6">
              <div className="border-2 border-dashed border-gray-300 rounded-lg p-8 text-center hover:border-indigo-500 transition-colors">
                <input
                  type="file"
                  id="fileInput"
                  className="hidden"
                  accept=".pdf,.docx,.xlsx,.pptx"
                  onChange="handleFileUpload(event)"
                />
                <label htmlFor="fileInput" className="cursor-pointer">
                  <i className="fas fa-cloud-upload-alt text-5xl text-gray-400 mb-4"></i>
                  <p className="text-gray-600 mb-2">
                    Перетащите файл сюда или <span className="text-indigo-600 font-semibold">выберите файл</span>
                  </p>
                  <p className="text-sm text-gray-500">
                    Поддерживаются: PDF, DOCX, XLSX, PPTX
                  </p>
                </label>
              </div>
            </div>

            {/* Google Sheets Import */}
            <div className="mb-6">
              <h3 className="text-lg font-medium mb-3 text-gray-700">
                <i className="fab fa-google mr-2"></i>
                Импорт из Google Sheets
              </h3>
              <div className="flex gap-2">
                <input
                  type="text"
                  id="googleSheetsUrl"
                  placeholder="Вставьте ссылку на Google Sheets..."
                  className="flex-1 px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
                />
                <button
                  onClick="importFromGoogleSheets()"
                  className="px-6 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition-colors"
                >
                  <i className="fas fa-download mr-2"></i>
                  Импортировать
                </button>
              </div>
            </div>

            {/* Manual Input */}
            <div className="mb-6">
              <button
                onClick="showManualInput()"
                className="w-full px-6 py-3 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition-colors"
              >
                <i className="fas fa-edit mr-2"></i>
                Создать структуру вручную
              </button>
            </div>
          </div>
        </div>

        {/* Organization Chart Container */}
        <div id="chartContainer" className="hidden">
          <div className="bg-white rounded-lg shadow-lg p-8">
            <div className="flex justify-between items-center mb-6">
              <h2 className="text-2xl font-semibold text-gray-700">
                <i className="fas fa-network-wired mr-2"></i>
                Организационная диаграмма
              </h2>
              <div className="flex gap-2">
                <button
                  onClick="exportToGoogleSheets()"
                  className="px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition-colors"
                >
                  <i className="fas fa-file-export mr-2"></i>
                  Экспорт в Google Sheets
                </button>
                <button
                  onClick="downloadAsJSON()"
                  className="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
                >
                  <i className="fas fa-download mr-2"></i>
                  Скачать JSON
                </button>
                <button
                  onClick="editStructure()"
                  className="px-4 py-2 bg-yellow-600 text-white rounded-lg hover:bg-yellow-700 transition-colors"
                >
                  <i className="fas fa-pen mr-2"></i>
                  Редактировать
                </button>
              </div>
            </div>
            
            {/* Chart will be rendered here */}
            <div id="orgChart" className="overflow-x-auto"></div>
          </div>
        </div>

        {/* Manual Input Modal */}
        <div id="manualInputModal" className="hidden fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg p-8 max-w-2xl w-full max-h-[80vh] overflow-y-auto">
            <h3 className="text-2xl font-semibold mb-6">Создать структуру вручную</h3>
            <div id="manualInputForm">
              <div className="mb-4">
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Структура организации (JSON или текст)
                </label>
                <textarea
                  id="manualStructureInput"
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
                  rows="10"
                  placeholder='Пример JSON:
{
  "name": "CEO",
  "title": "Генеральный директор",
  "children": [
    {
      "name": "CTO",
      "title": "Технический директор",
      "children": []
    }
  ]
}

Или текстовый формат:
CEO - Генеральный директор
  CTO - Технический директор
    Разработчик 1
    Разработчик 2
  CFO - Финансовый директор'
                ></textarea>
              </div>
              <div className="flex gap-2">
                <button
                  onClick="processManualInput()"
                  className="flex-1 px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700"
                >
                  Создать диаграмму
                </button>
                <button
                  onClick="closeManualInput()"
                  className="px-4 py-2 bg-gray-300 text-gray-700 rounded-lg hover:bg-gray-400"
                >
                  Отмена
                </button>
              </div>
            </div>
          </div>
        </div>

        {/* Loading Indicator */}
        <div id="loadingIndicator" className="hidden fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg p-8">
            <div className="flex items-center">
              <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-indigo-600 mr-4"></div>
              <span className="text-lg">Обработка файла...</span>
            </div>
          </div>
        </div>
      </div>

      <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
      <script src="/static/professional-parser.js"></script>
      <script src="/static/d3-visualization.js"></script>
      <script src="/static/org-chart-enhanced.js"></script>
    </div>,
    { title: 'Организационная структура' }
  )
})

// API Routes

// Parse uploaded file
app.post('/api/parse', async (c) => {
  try {
    const formData = await c.req.formData()
    const file = formData.get('file') as File
    
    if (!file) {
      return c.json({ success: false, error: 'No file provided' }, 400)
    }

    const fileName = file.name.toLowerCase()
    let result: ParseResult = { success: false, error: 'Unsupported file type' }

    if (fileName.endsWith('.xlsx')) {
      result = await parseExcel(file)
    } else if (fileName.endsWith('.docx')) {
      result = await parseWord(file)
    } else if (fileName.endsWith('.pdf')) {
      result = await parsePDF(file)
    } else if (fileName.endsWith('.pptx')) {
      result = await parsePowerPoint(file)
    }

    return c.json(result)
  } catch (error) {
    console.error('Parse error:', error)
    return c.json({ success: false, error: 'Failed to parse file' }, 500)
  }
})

// Import from Google Sheets
app.post('/api/import-google-sheets', async (c) => {
  try {
    const { url } = await c.req.json()
    
    if (!url) {
      return c.json({ success: false, error: 'No URL provided' }, 400)
    }

    // Extract sheet ID from URL
    const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/)
    if (!match) {
      return c.json({ success: false, error: 'Invalid Google Sheets URL' }, 400)
    }

    const sheetId = match[1]
    const exportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv`

    // Fetch CSV data
    const response = await fetch(exportUrl)
    if (!response.ok) {
      return c.json({ success: false, error: 'Failed to fetch Google Sheets data' }, 500)
    }

    const csvText = await response.text()
    
    // Return the raw CSV for client-side parsing with enhanced parser
    return c.json({ 
      success: true, 
      csv: csvText,
      message: 'CSV data fetched successfully'
    })
  } catch (error) {
    console.error('Google Sheets import error:', error)
    return c.json({ success: false, error: 'Failed to import from Google Sheets' }, 500)
  }
})

// Export to Google Sheets format
app.post('/api/export-google-sheets', async (c) => {
  try {
    const { structure } = await c.req.json()
    
    if (!structure) {
      return c.json({ success: false, error: 'No structure provided' }, 400)
    }

    const csvData = convertOrgStructureToCSV(structure)
    
    return c.json({ 
      success: true, 
      csv: csvData,
      instructions: 'Copy this CSV data and paste it into a new Google Sheets document'
    })
  } catch (error) {
    console.error('Export error:', error)
    return c.json({ success: false, error: 'Failed to export structure' }, 500)
  }
})

// Process document from URL
app.post('/api/process-document-url', async (c) => {
  try {
    const { url, fileName } = await c.req.json()
    
    if (!url) {
      return c.json({ success: false, error: 'No URL provided' }, 400)
    }

    // Fetch the document
    const response = await fetch(url)
    if (!response.ok) {
      return c.json({ success: false, error: 'Failed to fetch document' }, 500)
    }

    const contentType = response.headers.get('content-type') || ''
    let text = ''

    // Process based on file type
    if (fileName?.endsWith('.pdf') || contentType.includes('pdf')) {
      // For PDF, we need to extract text on client side
      return c.json({ 
        success: false, 
        error: 'PDF processing should be done client-side',
        needsClientProcessing: true 
      })
    } else if (fileName?.endsWith('.docx') || contentType.includes('wordprocessingml')) {
      // For DOCX, extract text from the response
      const buffer = await response.arrayBuffer()
      // Simple text extraction from DOCX
      text = await extractTextFromBuffer(buffer)
    } else if (fileName?.endsWith('.xlsx') || contentType.includes('spreadsheetml')) {
      // For Excel, needs client-side processing
      return c.json({ 
        success: false, 
        error: 'Excel processing should be done client-side',
        needsClientProcessing: true 
      })
    } else if (fileName?.endsWith('.pptx') || contentType.includes('presentationml')) {
      // For PowerPoint, needs client-side processing
      return c.json({ 
        success: false, 
        error: 'PowerPoint processing should be done client-side',
        needsClientProcessing: true 
      })
    } else {
      // Try to get text content
      text = await response.text()
    }

    // Parse the text to extract structure
    const structure = parseAdvancedText(text)
    
    return c.json({ 
      success: true, 
      structure: structure,
      text: text
    })
  } catch (error) {
    console.error('Document processing error:', error)
    return c.json({ success: false, error: 'Failed to process document' }, 500)
  }
})

// Helper function to extract text from buffer
async function extractTextFromBuffer(buffer: ArrayBuffer): Promise<string> {
  // Convert buffer to text (simplified version)
  const decoder = new TextDecoder('utf-8')
  const text = decoder.decode(buffer)
  
  // Extract readable text (remove binary data)
  const cleanText = text.replace(/[^\x20-\x7E\u0400-\u04FF\n\r\t]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
  
  return cleanText
}

// Advanced text parsing
function parseAdvancedText(text: string): any {
  const lines = text.split('\n').filter(line => line.trim())
  
  // Pattern matching for organizational structure
  const namePattern = /([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)|([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)/g
  const positionPattern = /начальник|заместитель|руководитель|директор|менеджер|специалист|координатор/gi
  
  const employees: any[] = []
  const departments = new Map()
  
  // Extract employees and positions
  lines.forEach((line, index) => {
    const nameMatches = line.match(namePattern)
    const positionMatches = line.match(positionPattern)
    
    if (nameMatches && nameMatches.length > 0) {
      nameMatches.forEach(name => {
        if (name.split(' ').length >= 2) {
          const employee = {
            name: name.trim(),
            title: '',
            department: '',
            level: 5
          }
          
          // Find position in nearby lines
          for (let offset = -2; offset <= 2; offset++) {
            if (lines[index + offset] && lines[index + offset].match(positionPattern)) {
              const posLine = lines[index + offset]
              const posMatch = posLine.match(/.*?(начальник|заместитель|руководитель|директор|менеджер|специалист|координатор).*/gi)
              if (posMatch) {
                employee.title = posMatch[0].trim()
                break
              }
            }
          }
          
          // Determine level based on position
          if (employee.title.toLowerCase().includes('генеральный') || 
              employee.title.toLowerCase().includes('президент')) {
            employee.level = 1
          } else if (employee.title.toLowerCase().includes('заместитель') && 
                     employee.title.toLowerCase().includes('начальника управления')) {
            employee.level = 2
          } else if (employee.title.toLowerCase().includes('начальник управления') || 
                     employee.title.toLowerCase().includes('директор')) {
            employee.level = 2
          } else if (employee.title.toLowerCase().includes('начальник отдела')) {
            employee.level = 3
          } else if (employee.title.toLowerCase().includes('заместитель начальника отдела')) {
            employee.level = 4
          } else if (employee.title.toLowerCase().includes('руководитель')) {
            employee.level = 3
          } else if (employee.title.toLowerCase().includes('главный')) {
            employee.level = 4
          } else if (employee.title.toLowerCase().includes('ведущий') || 
                     employee.title.toLowerCase().includes('старший')) {
            employee.level = 5
          } else {
            employee.level = 6
          }
          
          employees.push(employee)
        }
      })
    }
  })
  
  // Build hierarchy
  employees.sort((a, b) => a.level - b.level)
  
  const root = {
    id: 'root',
    name: 'Организация',
    title: '',
    children: [] as any[]
  }
  
  // Group by levels
  const levelGroups = new Map()
  employees.forEach(emp => {
    if (!levelGroups.has(emp.level)) {
      levelGroups.set(emp.level, [])
    }
    levelGroups.get(emp.level).push({
      id: `emp-${Math.random().toString(36).substr(2, 9)}`,
      name: emp.name,
      title: emp.title || 'Сотрудник',
      children: []
    })
  })
  
  // Build hierarchy based on levels
  const sortedLevels = Array.from(levelGroups.keys()).sort((a, b) => a - b)
  
  for (let i = 0; i < sortedLevels.length; i++) {
    const currentLevel = sortedLevels[i]
    const currentNodes = levelGroups.get(currentLevel)
    
    if (i === 0) {
      root.children.push(...currentNodes)
    } else {
      const parentLevel = sortedLevels[i - 1]
      const parentNodes = levelGroups.get(parentLevel)
      
      currentNodes.forEach((node: any, index: number) => {
        const parentIndex = Math.floor(index * parentNodes.length / currentNodes.length)
        const parent = parentNodes[Math.min(parentIndex, parentNodes.length - 1)]
        parent.children.push(node)
      })
    }
  }
  
  // Return the root if it has only one child
  if (root.children.length === 1) {
    return root.children[0]
  }
  
  return root
}

// Parse manual input
app.post('/api/parse-manual', async (c) => {
  try {
    const { input } = await c.req.json()
    
    if (!input) {
      return c.json({ success: false, error: 'No input provided' }, 400)
    }

    // Try to parse as JSON first
    try {
      const jsonData = JSON.parse(input)
      return c.json({ success: true, data: jsonData })
    } catch {
      // Parse as text format
      const result = parseTextToOrgStructure(input)
      return c.json(result)
    }
  } catch (error) {
    console.error('Manual parse error:', error)
    return c.json({ success: false, error: 'Failed to parse input' }, 500)
  }
})

// Helper functions

async function parseExcel(file: File): Promise<ParseResult> {
  try {
    // Since we're in Cloudflare Workers environment, we need to handle this differently
    // We'll process Excel on the client side instead
    return { success: false, error: 'Excel parsing should be done client-side' }
  } catch (error) {
    return { success: false, error: 'Failed to parse Excel file' }
  }
}

async function parseWord(file: File): Promise<ParseResult> {
  try {
    // Word parsing should also be done client-side in Cloudflare Workers
    return { success: false, error: 'Word parsing should be done client-side' }
  } catch (error) {
    return { success: false, error: 'Failed to parse Word file' }
  }
}

async function parsePDF(file: File): Promise<ParseResult> {
  try {
    // PDF parsing should be done client-side
    return { success: false, error: 'PDF parsing should be done client-side' }
  } catch (error) {
    return { success: false, error: 'Failed to parse PDF file' }
  }
}

async function parsePowerPoint(file: File): Promise<ParseResult> {
  try {
    // PowerPoint parsing should be done client-side
    return { success: false, error: 'PowerPoint parsing should be done client-side' }
  } catch (error) {
    return { success: false, error: 'Failed to parse PowerPoint file' }
  }
}

function parseCSVToOrgStructure(csv: string): ParseResult {
  try {
    const lines = csv.split('\n').filter(line => line.trim())
    if (lines.length === 0) {
      return { success: false, error: 'Empty CSV' }
    }

    // Assume first row is headers
    const headers = lines[0].split(',').map(h => h.trim())
    
    // Build hierarchy from CSV
    const nodes: Map<string, OrgNode> = new Map()
    let root: OrgNode | null = null

    for (let i = 1; i < lines.length; i++) {
      const values = lines[i].split(',').map(v => v.trim())
      const node: OrgNode = {
        id: `node-${i}`,
        name: values[0] || '',
        title: values[1] || '',
        department: values[2] || '',
        email: values[3] || '',
        children: []
      }

      nodes.set(node.name, node)

      // If there's a parent column, use it to build hierarchy
      const parentName = values[4]
      if (parentName && nodes.has(parentName)) {
        nodes.get(parentName)!.children.push(node)
      } else if (!root) {
        root = node
      }
    }

    if (!root && nodes.size > 0) {
      root = nodes.values().next().value
    }

    return { success: true, data: root }
  } catch (error) {
    return { success: false, error: 'Failed to parse CSV' }
  }
}

function convertOrgStructureToCSV(structure: OrgNode): string {
  const rows: string[] = []
  rows.push('Name,Title,Department,Email,Parent')

  function traverse(node: OrgNode, parent: string = '') {
    const row = [
      node.name || '',
      node.title || '',
      node.department || '',
      node.email || '',
      parent
    ].map(v => `"${v}"`).join(',')
    
    rows.push(row)

    for (const child of node.children || []) {
      traverse(child, node.name)
    }
  }

  traverse(structure)
  return rows.join('\n')
}

function parseTextToOrgStructure(text: string): ParseResult {
  try {
    const lines = text.split('\n').filter(line => line.trim())
    if (lines.length === 0) {
      return { success: false, error: 'Empty input' }
    }

    const root: OrgNode = {
      id: 'root',
      name: '',
      children: []
    }

    const stack: { node: OrgNode, level: number }[] = [{ node: root, level: -1 }]

    for (const line of lines) {
      const level = (line.match(/^[ \t]*/)?.[0].length || 0) / 2
      const content = line.trim()
      
      // Parse name and title
      const parts = content.split('-').map(p => p.trim())
      const node: OrgNode = {
        id: `node-${Math.random().toString(36).substr(2, 9)}`,
        name: parts[0] || '',
        title: parts[1] || '',
        children: []
      }

      // Find parent
      while (stack.length > 0 && stack[stack.length - 1].level >= level) {
        stack.pop()
      }

      const parent = stack[stack.length - 1].node
      parent.children.push(node)
      stack.push({ node, level })
    }

    if (root.children.length === 1) {
      return { success: true, data: root.children[0] }
    }

    return { success: true, data: root }
  } catch (error) {
    return { success: false, error: 'Failed to parse text input' }
  }
}

export default app