// Advanced D3.js visualization for organizational charts
class D3OrgChart {
    constructor(containerId) {
        this.containerId = containerId;
        this.container = document.getElementById(containerId);
        this.width = this.container.clientWidth || 1200;
        this.height = 800;
        this.currentView = 'horizontal'; // horizontal, vertical, radial
        this.svg = null;
        this.g = null;
        this.zoom = null;
        this.tooltip = null;
        this.searchTerm = '';
        this.data = null;
        this.root = null;
        
        this.colors = {
            level1: '#4F46E5', // Indigo
            level2: '#7C3AED', // Purple
            level3: '#2563EB', // Blue
            level4: '#0891B2', // Cyan
            level5: '#059669', // Green
            level6: '#DC2626', // Red
            department: '#F59E0B', // Amber
            default: '#6B7280'  // Gray
        };
        
        this.init();
    }

    init() {
        // Clear container
        this.container.innerHTML = '';
        
        // Create controls
        this.createControls();
        
        // Create SVG
        this.svg = d3.select(this.container)
            .append('svg')
            .attr('width', this.width)
            .attr('height', this.height)
            .attr('class', 'org-chart-svg')
            .style('background', 'linear-gradient(to bottom right, #EFF6FF, #F0F9FF)');

        // Add zoom behavior
        this.zoom = d3.zoom()
            .scaleExtent([0.1, 4])
            .on('zoom', (event) => {
                this.g.attr('transform', event.transform);
            });

        this.svg.call(this.zoom);

        // Create main group
        this.g = this.svg.append('g')
            .attr('class', 'chart-container');

        // Create tooltip
        this.createTooltip();
        
        // Add styles
        this.addStyles();
    }

    createControls() {
        const controls = document.createElement('div');
        controls.className = 'chart-controls flex gap-2 mb-4 p-4 bg-white rounded-lg shadow';
        controls.innerHTML = `
            <button onclick="orgChart.setView('horizontal')" class="px-4 py-2 bg-blue-500 text-white rounded hover:bg-blue-600 transition">
                <i class="fas fa-arrows-alt-h mr-2"></i>Горизонтальное
            </button>
            <button onclick="orgChart.setView('vertical')" class="px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600 transition">
                <i class="fas fa-arrows-alt-v mr-2"></i>Вертикальное
            </button>
            <button onclick="orgChart.setView('radial')" class="px-4 py-2 bg-purple-500 text-white rounded hover:bg-purple-600 transition">
                <i class="fas fa-circle-notch mr-2"></i>Радиальное
            </button>
            <button onclick="orgChart.expandAll()" class="px-4 py-2 bg-gray-500 text-white rounded hover:bg-gray-600 transition">
                <i class="fas fa-expand mr-2"></i>Развернуть все
            </button>
            <button onclick="orgChart.collapseAll()" class="px-4 py-2 bg-gray-500 text-white rounded hover:bg-gray-600 transition">
                <i class="fas fa-compress mr-2"></i>Свернуть все
            </button>
            <button onclick="orgChart.resetZoom()" class="px-4 py-2 bg-yellow-500 text-white rounded hover:bg-yellow-600 transition">
                <i class="fas fa-search-location mr-2"></i>Сброс масштаба
            </button>
            <button onclick="orgChart.exportSVG()" class="px-4 py-2 bg-indigo-500 text-white rounded hover:bg-indigo-600 transition">
                <i class="fas fa-file-export mr-2"></i>Экспорт SVG
            </button>
            <button onclick="orgChart.exportPNG()" class="px-4 py-2 bg-pink-500 text-white rounded hover:bg-pink-600 transition">
                <i class="fas fa-image mr-2"></i>Экспорт PNG
            </button>
            <div class="flex-1"></div>
            <input type="text" id="searchInput" placeholder="Поиск сотрудника..." 
                   class="px-4 py-2 border rounded focus:outline-none focus:ring-2 focus:ring-blue-500"
                   onkeyup="orgChart.search(this.value)">
        `;
        this.container.appendChild(controls);
    }

    createTooltip() {
        this.tooltip = d3.select('body').append('div')
            .attr('class', 'org-tooltip')
            .style('opacity', 0)
            .style('position', 'absolute')
            .style('background', 'white')
            .style('padding', '12px')
            .style('border-radius', '8px')
            .style('box-shadow', '0 4px 6px rgba(0, 0, 0, 0.1)')
            .style('pointer-events', 'none')
            .style('font-size', '14px')
            .style('max-width', '300px')
            .style('z-index', '1000');
    }

    addStyles() {
        const style = document.createElement('style');
        style.textContent = `
            .org-chart-svg {
                border: 1px solid #E5E7EB;
                border-radius: 8px;
                cursor: move;
            }
            .node {
                cursor: pointer;
                transition: all 0.3s ease;
            }
            .node rect {
                stroke-width: 2px;
                transition: all 0.3s ease;
            }
            .node:hover rect {
                stroke-width: 3px;
                filter: brightness(1.1);
            }
            .node.highlighted rect {
                stroke: #F59E0B !important;
                stroke-width: 4px !important;
            }
            .link {
                fill: none;
                stroke: #CBD5E1;
                stroke-width: 2px;
                transition: all 0.3s ease;
            }
            .link:hover {
                stroke: #6366F1;
                stroke-width: 3px;
            }
            .node-text {
                font-family: system-ui, -apple-system, sans-serif;
                pointer-events: none;
            }
            .node-name {
                font-weight: 600;
                fill: #1F2937;
            }
            .node-title {
                font-size: 12px;
                fill: #6B7280;
            }
            @keyframes fadeIn {
                from { opacity: 0; transform: scale(0.8); }
                to { opacity: 1; transform: scale(1); }
            }
            .node-enter {
                animation: fadeIn 0.5s ease;
            }
        `;
        document.head.appendChild(style);
    }

    render(data) {
        this.data = data;
        this.root = d3.hierarchy(data);
        
        // Clear previous content
        this.g.selectAll('*').remove();
        
        switch (this.currentView) {
            case 'horizontal':
                this.renderHorizontalTree();
                break;
            case 'vertical':
                this.renderVerticalTree();
                break;
            case 'radial':
                this.renderRadialTree();
                break;
            default:
                this.renderHorizontalTree();
        }
    }

    renderHorizontalTree() {
        const margin = { top: 20, right: 120, bottom: 20, left: 120 };
        const width = this.width - margin.left - margin.right;
        const height = this.height - margin.top - margin.bottom;

        const treeLayout = d3.tree()
            .size([height, width])
            .separation((a, b) => (a.parent === b.parent ? 1 : 1.5));

        treeLayout(this.root);

        const g = this.g.append('g')
            .attr('transform', `translate(${margin.left},${margin.top})`);

        // Draw links
        const links = g.selectAll('.link')
            .data(this.root.links())
            .enter().append('path')
            .attr('class', 'link')
            .attr('d', d => {
                return `M${d.source.y},${d.source.x}
                        C${(d.source.y + d.target.y) / 2},${d.source.x}
                         ${(d.source.y + d.target.y) / 2},${d.target.x}
                         ${d.target.y},${d.target.x}`;
            });

        // Draw nodes
        const nodes = g.selectAll('.node')
            .data(this.root.descendants())
            .enter().append('g')
            .attr('class', d => `node ${this.searchTerm && !this.matchesSearch(d) ? 'dimmed' : ''}`)
            .attr('transform', d => `translate(${d.y},${d.x})`)
            .on('click', (event, d) => this.onNodeClick(event, d))
            .on('mouseover', (event, d) => this.showTooltip(event, d))
            .on('mouseout', () => this.hideTooltip());

        // Add rectangles
        nodes.append('rect')
            .attr('x', -75)
            .attr('y', -25)
            .attr('width', 150)
            .attr('height', 50)
            .attr('rx', 8)
            .attr('fill', d => this.getNodeColor(d))
            .attr('stroke', '#fff')
            .attr('stroke-width', 2);

        // Add text - name
        nodes.append('text')
            .attr('class', 'node-name')
            .attr('dy', -5)
            .attr('text-anchor', 'middle')
            .text(d => this.truncateText(d.data.name, 20));

        // Add text - title
        nodes.append('text')
            .attr('class', 'node-title')
            .attr('dy', 10)
            .attr('text-anchor', 'middle')
            .text(d => this.truncateText(d.data.title, 25));

        // Add expand/collapse indicator
        nodes.filter(d => d.children || d._children)
            .append('circle')
            .attr('cx', 75)
            .attr('cy', 0)
            .attr('r', 8)
            .attr('fill', '#fff')
            .attr('stroke', '#6366F1')
            .attr('stroke-width', 2);

        nodes.filter(d => d.children || d._children)
            .append('text')
            .attr('x', 75)
            .attr('dy', 4)
            .attr('text-anchor', 'middle')
            .attr('font-size', 12)
            .attr('fill', '#6366F1')
            .text(d => d._children ? '+' : '-');
    }

    renderVerticalTree() {
        const margin = { top: 40, right: 20, bottom: 40, left: 20 };
        const width = this.width - margin.left - margin.right;
        const height = this.height - margin.top - margin.bottom;

        const treeLayout = d3.tree()
            .size([width, height])
            .separation((a, b) => (a.parent === b.parent ? 1 : 1.5));

        treeLayout(this.root);

        const g = this.g.append('g')
            .attr('transform', `translate(${margin.left},${margin.top})`);

        // Draw links
        const links = g.selectAll('.link')
            .data(this.root.links())
            .enter().append('path')
            .attr('class', 'link')
            .attr('d', d => {
                return `M${d.source.x},${d.source.y}
                        C${d.source.x},${(d.source.y + d.target.y) / 2}
                         ${d.target.x},${(d.source.y + d.target.y) / 2}
                         ${d.target.x},${d.target.y}`;
            });

        // Draw nodes
        const nodes = g.selectAll('.node')
            .data(this.root.descendants())
            .enter().append('g')
            .attr('class', 'node')
            .attr('transform', d => `translate(${d.x},${d.y})`)
            .on('click', (event, d) => this.onNodeClick(event, d))
            .on('mouseover', (event, d) => this.showTooltip(event, d))
            .on('mouseout', () => this.hideTooltip());

        // Add rectangles
        nodes.append('rect')
            .attr('x', -75)
            .attr('y', -25)
            .attr('width', 150)
            .attr('height', 50)
            .attr('rx', 8)
            .attr('fill', d => this.getNodeColor(d))
            .attr('stroke', '#fff')
            .attr('stroke-width', 2);

        // Add text
        nodes.append('text')
            .attr('class', 'node-name')
            .attr('dy', -5)
            .attr('text-anchor', 'middle')
            .text(d => this.truncateText(d.data.name, 20));

        nodes.append('text')
            .attr('class', 'node-title')
            .attr('dy', 10)
            .attr('text-anchor', 'middle')
            .text(d => this.truncateText(d.data.title, 25));
    }

    renderRadialTree() {
        const radius = Math.min(this.width, this.height) / 2 - 100;
        
        const treeLayout = d3.tree()
            .size([2 * Math.PI, radius])
            .separation((a, b) => (a.parent === b.parent ? 1 : 2) / a.depth);

        treeLayout(this.root);

        const g = this.g.append('g')
            .attr('transform', `translate(${this.width / 2},${this.height / 2})`);

        // Draw links
        const links = g.selectAll('.link')
            .data(this.root.links())
            .enter().append('path')
            .attr('class', 'link')
            .attr('d', d3.linkRadial()
                .angle(d => d.x)
                .radius(d => d.y));

        // Draw nodes
        const nodes = g.selectAll('.node')
            .data(this.root.descendants())
            .enter().append('g')
            .attr('class', 'node')
            .attr('transform', d => `
                rotate(${d.x * 180 / Math.PI - 90})
                translate(${d.y},0)
            `)
            .on('click', (event, d) => this.onNodeClick(event, d))
            .on('mouseover', (event, d) => this.showTooltip(event, d))
            .on('mouseout', () => this.hideTooltip());

        // Add circles
        nodes.append('circle')
            .attr('r', 6)
            .attr('fill', d => this.getNodeColor(d))
            .attr('stroke', '#fff')
            .attr('stroke-width', 2);

        // Add text
        nodes.append('text')
            .attr('dy', '0.31em')
            .attr('x', d => d.x < Math.PI === !d.children ? 10 : -10)
            .attr('text-anchor', d => d.x < Math.PI === !d.children ? 'start' : 'end')
            .attr('transform', d => d.x >= Math.PI ? 'rotate(180)' : null)
            .attr('class', 'node-name')
            .attr('font-size', 12)
            .text(d => d.data.name);
    }

    getNodeColor(node) {
        const level = node.depth;
        if (node.data.isDepartment) return this.colors.department;
        if (level === 0) return this.colors.level1;
        if (level === 1) return this.colors.level2;
        if (level === 2) return this.colors.level3;
        if (level === 3) return this.colors.level4;
        if (level === 4) return this.colors.level5;
        if (level === 5) return this.colors.level6;
        return this.colors.default;
    }

    truncateText(text, maxLength) {
        if (!text) return '';
        if (text.length <= maxLength) return text;
        return text.substring(0, maxLength - 3) + '...';
    }

    onNodeClick(event, node) {
        event.stopPropagation();
        
        // Toggle children
        if (node.children) {
            node._children = node.children;
            node.children = null;
        } else if (node._children) {
            node.children = node._children;
            node._children = null;
        }
        
        // Re-render
        this.render(this.data);
    }

    showTooltip(event, node) {
        let content = `<strong>${node.data.name}</strong>`;
        if (node.data.title) {
            content += `<br><em>${node.data.title}</em>`;
        }
        if (node.data.email) {
            content += `<br><i class="fas fa-envelope"></i> ${node.data.email}`;
        }
        if (node.data.responsibilities && node.data.responsibilities.length > 0) {
            content += '<br><strong>Обязанности:</strong><ul style="margin: 5px 0; padding-left: 20px;">';
            node.data.responsibilities.slice(0, 3).forEach(resp => {
                content += `<li style="font-size: 12px;">${resp}</li>`;
            });
            if (node.data.responsibilities.length > 3) {
                content += `<li style="font-size: 12px;">и еще ${node.data.responsibilities.length - 3}...</li>`;
            }
            content += '</ul>';
        }
        
        this.tooltip.transition()
            .duration(200)
            .style('opacity', .9);
        
        this.tooltip.html(content)
            .style('left', (event.pageX + 10) + 'px')
            .style('top', (event.pageY - 28) + 'px');
    }

    hideTooltip() {
        this.tooltip.transition()
            .duration(500)
            .style('opacity', 0);
    }

    setView(view) {
        this.currentView = view;
        this.render(this.data);
    }

    expandAll() {
        this.root.descendants().forEach(d => {
            if (d._children) {
                d.children = d._children;
                d._children = null;
            }
        });
        this.render(this.data);
    }

    collapseAll() {
        this.root.descendants().forEach(d => {
            if (d.children && d.depth > 0) {
                d._children = d.children;
                d.children = null;
            }
        });
        this.render(this.data);
    }

    resetZoom() {
        this.svg.transition()
            .duration(750)
            .call(this.zoom.transform, d3.zoomIdentity);
    }

    search(term) {
        this.searchTerm = term.toLowerCase();
        
        // Highlight matching nodes
        this.g.selectAll('.node')
            .classed('highlighted', d => this.matchesSearch(d));
        
        // Auto-expand path to matching nodes
        if (term) {
            this.root.descendants().forEach(d => {
                if (this.matchesSearch(d)) {
                    let parent = d.parent;
                    while (parent) {
                        if (parent._children) {
                            parent.children = parent._children;
                            parent._children = null;
                        }
                        parent = parent.parent;
                    }
                }
            });
            this.render(this.data);
        }
    }

    matchesSearch(node) {
        if (!this.searchTerm) return true;
        const name = (node.data.name || '').toLowerCase();
        const title = (node.data.title || '').toLowerCase();
        return name.includes(this.searchTerm) || title.includes(this.searchTerm);
    }

    exportSVG() {
        const svgData = this.svg.node().outerHTML;
        const blob = new Blob([svgData], { type: 'image/svg+xml' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'org-chart.svg';
        a.click();
        URL.revokeObjectURL(url);
    }

    exportPNG() {
        const svgData = new XMLSerializer().serializeToString(this.svg.node());
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        const img = new Image();
        
        canvas.width = this.width;
        canvas.height = this.height;
        
        img.onload = () => {
            ctx.fillStyle = 'white';
            ctx.fillRect(0, 0, canvas.width, canvas.height);
            ctx.drawImage(img, 0, 0);
            
            canvas.toBlob(blob => {
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'org-chart.png';
                a.click();
                URL.revokeObjectURL(url);
            });
        };
        
        img.src = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svgData)));
    }
}

// Make it globally available
window.D3OrgChart = D3OrgChart;