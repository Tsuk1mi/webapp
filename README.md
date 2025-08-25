# Организационная структура - Organization Structure Parser

## Project Overview
- **Name**: Organization Structure Parser
- **Goal**: Parse organizational structures from multiple document formats and create interactive hierarchical visualizations
- **Features**: 
  - Multi-format document parsing (PDF, DOCX, XLSX, PPTX, Google Sheets)
  - Interactive D3.js tree visualizations (horizontal, vertical, radial)
  - Google Sheets import/export functionality
  - Enhanced CSV parsing with proper quote and comma handling
  - Subordination code-based hierarchy building
  - Support for 100+ employee structures

## URLs
- **Production**: https://3000-i0dopoq9stmpmvbxc3qq4-6532622b.e2b.dev
- **GitHub**: Not yet deployed

## Currently Completed Features
1. ✅ Basic file upload interface with drag-and-drop
2. ✅ Document parsing infrastructure for multiple formats
3. ✅ D3.js interactive tree visualization with multiple views
4. ✅ Google Sheets import via public URL
5. ✅ PPTX parsing using JSZip library
6. ✅ Enhanced CSV parser with proper quote handling
7. ✅ Subordination code-based hierarchy building
8. ✅ Detailed table view with export options
9. ✅ Integration of enhanced parser for better accuracy

## Functional Entry URIs
- `/` - Main application interface
- `/api/import-google-sheets` - POST endpoint for Google Sheets import
  - Body: `{ "url": "https://docs.google.com/spreadsheets/d/[ID]/edit" }`
- `/api/parse` - POST endpoint for file parsing
  - Body: FormData with file
- `/api/export-google-sheets` - POST endpoint for export
  - Body: `{ "structure": {...} }`
- `/static/*` - Static assets (parsers, visualizations)

## Features Not Yet Implemented
1. ⏳ Database persistence (Cloudflare D1)
2. ⏳ User authentication and saved structures
3. ⏳ Real-time collaboration features
4. ⏳ Advanced search and filtering
5. ⏳ Automatic org chart layout optimization
6. ⏳ Export to PowerPoint/Word formats

## Recent Fixes (Version 3.0 - Professional)
- **NEW**: ProfessionalOrgParser - улучшенный парсер с поддержкой вакансий
- **Fixed**: Корректное определение должностей и подчиненности
- **Fixed**: Автоматическое определение вакансий (пустые позиции помечаются как "ВАКАНСИЯ")
- **Fixed**: Правильная обработка всех 100+ строк из Google Sheets
- **Improved**: Определение уровня иерархии по должности
- **Improved**: Экспорт в Excel с колонками: Код, Подразделение, Должность, ФИО, Функционал, Уровень, Статус
- **Removed**: Убраны тестовые кнопки из интерфейса
- **Added**: Подсчет вакансий и занятых позиций

## Data Architecture
- **Data Models**: Hierarchical tree structure with nodes containing:
  - id, name, position, department, functional description
  - subordination codes for parent-child relationships
  - children array for nested structures
- **Storage Services**: Currently in-memory, planned Cloudflare D1 for persistence
- **Data Flow**: 
  1. Document upload/import → Parser selection
  2. Text/data extraction → Structure building
  3. Hierarchy generation → Visualization rendering

## User Guide
1. **Upload Documents**: 
   - Click "Выберите файл" or drag-and-drop
   - Supported: PDF, DOCX, XLSX, PPTX
2. **Import from Google Sheets**:
   - Paste Google Sheets URL
   - Click "Импортировать"
3. **View Organization Chart**:
   - Use buttons to switch between views (horizontal, vertical, radial)
   - Click nodes to expand/collapse
   - Hover for details
4. **Export Data**:
   - "Экспорт в CSV" - Download as CSV file
   - "Формат для Google Sheets" - Copy-paste format
   - "Скачать JSON" - Download structure as JSON

## Test Documents
- Google Sheets Example: `15o26CMkMG_73ZhrlwlyR06vH2BDYhfu1DYkelcG_C44`
  - Contains 100+ employees with subordination codes
  - Tests enhanced CSV parsing capabilities

## Recommended Next Steps
1. **Testing**: Thoroughly test with user's actual documents
2. **Validation**: Verify all 100+ rows from Google Sheets parse correctly
3. **Performance**: Optimize for large organizational structures
4. **Persistence**: Implement Cloudflare D1 for data storage
5. **UI Enhancement**: Add loading indicators and error messages
6. **Export Features**: Implement direct Google Sheets API integration

## Deployment
- **Platform**: Cloudflare Pages
- **Status**: ✅ Active (Development)
- **Tech Stack**: Hono + TypeScript + D3.js + TailwindCSS
- **Last Updated**: 2025-08-25

## Technical Stack
- **Backend**: Hono Framework on Cloudflare Workers
- **Frontend**: Vanilla JavaScript with D3.js for visualizations
- **Styling**: TailwindCSS via CDN
- **Parsing Libraries**:
  - XLSX for Excel files
  - Mammoth for Word documents
  - PDF.js for PDF files
  - JSZip for PPTX files
- **Process Management**: PM2 in development

## Development Commands
```bash
# Build project
npm run build

# Start development server (uses PM2)
pm2 start ecosystem.config.cjs

# View logs
pm2 logs org-structure --nostream

# Restart server
pm2 restart org-structure

# Deploy to production
npm run deploy
```