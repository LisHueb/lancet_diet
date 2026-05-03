// Excel-Datei automatisch laden
const fileName = 'FAOSTAT_data_Supply.xlsx';

// State für expandierte Regionen
const expandedRegions = new Set();
let rawData = [];

fetch(fileName)
    .then(response => {
        if (!response.ok) {
            throw new Error(`Datei ${fileName} konnte nicht geladen werden`);
        }
        return response.arrayBuffer();
    })
    .then(data => {
        document.getElementById('loading').style.display = 'none';
        const workbook = XLSX.read(data, {type: 'array'});
        const sheetName = 'Sheet1';
        const worksheet = workbook.Sheets[sheetName];
        rawData = XLSX.utils.sheet_to_json(worksheet);
        
        processAndVisualize(rawData);
    })
    .catch(error => {
        document.getElementById('loading').style.display = 'none';
        const errorDiv = document.getElementById('error');
        errorDiv.style.display = 'block';
        errorDiv.textContent = `Fehler beim Laden: ${error.message}. Stelle sicher, dass die Datei "${fileName}" im gleichen Ordner liegt.`;
        console.error('Fehler:', error);
    });

function processAndVisualize(data) {
    // Region-Daten aggregieren
    const regionCategoryMap = new Map();
    const regionAreaMap = new Map();
    
    data.forEach(row => {
        const region = row['Region'];
        const area = row['Area'];
        const category = row['Lancet Category'];
        const value = parseFloat(row['Value']) || 0;
        const target = parseFloat(row['Lancet Target']) || 0;
        
        if (!region || !category) return;
        
        // Region-Daten
        const regionKey = `${region}|${category}`;
        if (!regionCategoryMap.has(regionKey)) {
            regionCategoryMap.set(regionKey, {
                region: region,
                category: category,
                totalValue: 0,
                totalTarget: 0,
                count: 0
            });
        }
        const regionEntry = regionCategoryMap.get(regionKey);
        regionEntry.totalValue += value;
        regionEntry.totalTarget += target;
        regionEntry.count += 1;
        
        // Area-Daten (für aufgeklappte Ansicht)
        const areaKey = `${area}|${category}`;
        if (!regionAreaMap.has(region)) {
            regionAreaMap.set(region, new Map());
        }
        const areaMap = regionAreaMap.get(region);
        if (!areaMap.has(areaKey)) {
            areaMap.set(areaKey, {
                area: area,
                category: category,
                totalValue: 0,
                totalTarget: 0
            });
        }
        const areaEntry = areaMap.get(areaKey);
        areaEntry.totalValue += value;
        areaEntry.totalTarget += target;
    });
    
    // Heatmap-Daten für Regionen
    const regionHeatmapData = [];
    regionCategoryMap.forEach(entry => {
        const averageTarget = entry.count > 0 ? entry.totalTarget / entry.count : 0;
        const share = entry.totalTarget > 0 ? entry.totalValue / entry.totalTarget : null;  // Für Farben: summierte Targets
        regionHeatmapData.push({
            type: 'region',
            region: entry.region,
            category: entry.category,
            share: share,
            value: entry.totalValue,
            target: averageTarget  // Für Tooltip: Durchschnitt
        });
    });
    
    // Heatmap-Daten für Areas
    const areaHeatmapData = [];
    regionAreaMap.forEach((areaMap, region) => {
        areaMap.forEach(entry => {
            const share = entry.totalTarget > 0 ? entry.totalValue / entry.totalTarget : null;
            areaHeatmapData.push({
                type: 'area',
                region: region,
                area: entry.area,
                category: entry.category,
                share: share,
                value: entry.totalValue,
                target: entry.totalTarget
            });
        });
    });
    
    // Alle Regionen und Kategorien sammeln
    const regions = [...new Set(regionHeatmapData.map(d => d.region))].filter(r => r);
    const categories = [...new Set(regionHeatmapData.map(d => d.category))].filter(c => c);
    
    // Areas pro Region
    const areasByRegion = new Map();
    regions.forEach(region => {
        const areas = [...new Set(areaHeatmapData.filter(d => d.region === region).map(d => d.area))];
        areasByRegion.set(region, areas);
    });
    
    // Min und Max Share für Farbskala
    const allShares = [...regionHeatmapData, ...areaHeatmapData]
        .filter(d => d.share !== null)
        .map(d => d.share);
    const minShare = Math.min(...allShares);
    const maxShare = 10; // Manuell festgelegt
    
    createHeatmap(regionHeatmapData, areaHeatmapData, regions, categories, areasByRegion, minShare, maxShare);
}

function createHeatmap(regionData, areaData, regions, categories, areasByRegion, minShare, maxShare) {
    // Dimensionen
    const margin = {top: 150, right: 50, bottom: 50, left: 180};
    const cellSize = 50;
    const width = categories.length * cellSize + margin.left + margin.right;
    
    // Dynamische Höhe basierend auf expandierten Regionen
    let totalRows = regions.length;
    expandedRegions.forEach(region => {
        totalRows += areasByRegion.get(region).length;
    });
    const height = totalRows * cellSize + margin.top + margin.bottom;
    
    // SVG erstellen oder auswählen
    let svg = d3.select('#heatmap').select('svg');
    let g;
    
    if (svg.empty()) {
        svg = d3.select('#heatmap')
            .append('svg')
            .attr('viewBox', `0 0 ${width} ${height}`)
            .attr('width', '100%')
            .attr('height', 'auto')
            .attr('style', 'max-width: 100%; height: auto; display: block; margin: 0 auto;');
        
        g = svg.append('g')
            .attr('transform', `translate(${margin.left},${margin.top})`);
    } else {
        g = svg.select('g');
        svg.transition()
            .duration(600)
            .attr('viewBox', `0 0 ${width} ${height}`);
    }
    
    // Farbskala: Lila (< 1) - Weiß (1) - Orange (> 1)
    const colorScale = d3.scaleLinear()
        .domain([minShare, 1, maxShare])
        .range(['#7b3294', '#f7f7f7', '#e66101'])
        .clamp(true);
    
    // Zeilen aufbauen (Regionen + expandierte Areas)
    let currentRow = 0;
    const rowData = [];
    
    regions.forEach(region => {
        rowData.push({
            type: 'region',
            region: region,
            row: currentRow
        });
        currentRow++;
        
        if (expandedRegions.has(region)) {
            const areas = areasByRegion.get(region);
            areas.forEach(area => {
                rowData.push({
                    type: 'area',
                    region: region,
                    area: area,
                    row: currentRow
                });
                currentRow++;
            });
        }
    });
    
    // Kreise zeichnen mit Animation und Update-Pattern
    rowData.forEach(row => {
        categories.forEach((category, j) => {
            let dataPoint;
            if (row.type === 'region') {
                dataPoint = regionData.find(d => d.region === row.region && d.category === category);
            } else {
                dataPoint = areaData.find(d => d.region === row.region && d.area === row.area && d.category === category);
            }
            const share = dataPoint ? dataPoint.share : null;
            const circleId = `circle-${row.type}-${row.region}-${row.area || ''}-${category}`
                .replace(/[^a-zA-Z0-9-]/g, '_');
            
            let circle = g.select(`#${CSS.escape(circleId)}`);
            
            if (circle.empty()) {
                // Neuer Kreis
                circle = g.append('circle')
                    .attr('id', circleId)
                    .attr('cx', j * cellSize + cellSize / 2)
                    .attr('r', 0)
                    .attr('fill', share !== null ? colorScale(share) : '#cccccc')
                    .attr('stroke', '#333')
                    .attr('stroke-width', 0.5)
                    .on('mouseover', function(event) {
                        if (dataPoint) {
                            showTooltip(event, row, category, dataPoint);
                        }
                    })
                    .on('mousemove', function(event) {
                        if (dataPoint) {
                            updateTooltipPosition(event);
                        }
                    })
                    .on('mouseout', function() {
                        hideTooltip();
                    });
                
                if (row.type === 'area') {
                    const regionRow = rowData.find(r => r.type === 'region' && r.region === row.region);
                    circle
                        .attr('cy', regionRow.row * cellSize + cellSize / 2)
                        .transition()
                        .duration(500)
                        .ease(d3.easeQuadOut)
                        .attr('cy', row.row * cellSize + cellSize / 2)
                        .attr('r', cellSize / 3);
                } else {
                    circle
                        .attr('cy', row.row * cellSize + cellSize / 2)
                        .transition()
                        .duration(400)
                        .attr('r', cellSize / 3);
                }
            } else {
                // Existierenden Kreis aktualisieren
                circle
                    .transition()
                    .duration(600)
                    .attr('cy', row.row * cellSize + cellSize / 2);
            }
        });
    });
    
    // Kreise entfernen, die nicht mehr in rowData sind
    const currentCircleIds = new Set();
    rowData.forEach(row => {
        categories.forEach(category => {
            const circleId = `circle-${row.type}-${row.region}-${row.area || ''}-${category}`
                .replace(/[^a-zA-Z0-9-]/g, '_');
            currentCircleIds.add(circleId);
        });
    });
    
    g.selectAll('circle').each(function() {
        const circle = d3.select(this);
        const id = circle.attr('id');
        if (id && !currentCircleIds.has(id)) {
            // Area-Kreis, der entfernt werden soll
            if (id.startsWith('circle-area-')) {
                const parts = id.split('-');
                const region = parts[2];
                const regionRow = rowData.find(r => r.type === 'region' && r.region.replace(/[^a-zA-Z0-9-]/g, '_') === region);
                if (regionRow) {
                    circle
                        .transition()
                        .duration(600)
                        .ease(d3.easeQuadIn)
                        .attr('cy', regionRow.row * cellSize + cellSize / 2)
                        .attr('r', 0)
                        .remove();
                } else {
                    circle.transition().duration(400).attr('r', 0).remove();
                }
            }
        }
    });
    
    // Y-Achse Labels und Expand-Icons mit Animation und Update-Pattern
    rowData.forEach(row => {
        if (row.type === 'region') {
            const isExpanded = expandedRegions.has(row.region);
            const iconId = `icon-${row.region}`.replace(/[^a-zA-Z0-9-]/g, '_');
            const labelId = `label-${row.region}`.replace(/[^a-zA-Z0-9-]/g, '_');
            
            let icon = g.select(`#${CSS.escape(iconId)}`);
            if (icon.empty()) {
                icon = g.append('text')
                    .attr('id', iconId)
                    .attr('class', 'expand-icon')
                    .attr('x', -25)
                    .attr('text-anchor', 'middle')
                    .attr('dy', '0.35em')
                    .on('click', function() {
                        toggleRegion(row.region);
                    });
            }
            
            icon
                .text(isExpanded ? '▼' : '▶')
                .attr('y', row.row * cellSize + cellSize / 2)
                .attr('opacity', 1);
            
            let label = g.select(`#${CSS.escape(labelId)}`);
            if (label.empty()) {
                label = g.append('text')
                    .attr('id', labelId)
                    .attr('class', 'region-label')
                    .attr('x', -35)
                    .attr('text-anchor', 'end')
                    .attr('dy', '0.35em')
                    .text(row.region)
                    .on('click', function() {
                        toggleRegion(row.region);
                    });
            }
            
            label
                .attr('y', row.row * cellSize + cellSize / 2)
                .attr('opacity', 1);
            
        } else {
            const areaLabelId = `area-label-${row.region}-${row.area}`.replace(/[^a-zA-Z0-9-]/g, '_');
            let areaLabel = g.select(`#${CSS.escape(areaLabelId)}`);
            
            if (areaLabel.empty()) {
                areaLabel = g.append('text')
                    .attr('id', areaLabelId)
                    .attr('class', 'area-label')
                    .attr('x', -35)
                    .attr('text-anchor', 'end')
                    .attr('dy', '0.35em')
                    .text('  ' + row.area)
                    .attr('opacity', 0);
                
                areaLabel.transition()
                    .delay(200)
                    .duration(400)
                    .attr('opacity', 1);
            }
            
            areaLabel
                .attr('y', row.row * cellSize + cellSize / 2);
        }
    });
    
    // Labels entfernen, die nicht mehr in rowData sind
    const currentLabelIds = new Set();
    rowData.forEach(row => {
        if (row.type === 'region') {
            currentLabelIds.add(`icon-${row.region}`.replace(/[^a-zA-Z0-9-]/g, '_'));
            currentLabelIds.add(`label-${row.region}`.replace(/[^a-zA-Z0-9-]/g, '_'));
        } else {
            currentLabelIds.add(`area-label-${row.region}-${row.area}`.replace(/[^a-zA-Z0-9-]/g, '_'));
        }
    });
    
    g.selectAll('text').each(function() {
        const label = d3.select(this);
        const id = label.attr('id');
        if (id && id.startsWith('area-label-') && !currentLabelIds.has(id)) {
            label
                .transition()
                .duration(400)
                .attr('opacity', 0)
                .remove();
        }
    });
    
    // X-Achse (Kategorien)
    if (g.selectAll('.category-label').empty()) {
        g.selectAll('.category-label')
            .data(categories)
            .enter()
            .append('text')
            .attr('class', 'category-label')
            .attr('x', 0)
            .attr('y', 0)
            .attr('transform', (d, i) => `translate(${i * cellSize + cellSize / 2}, -10) rotate(-45)`)
            .text(d => d);
    }
    
    // Legende erstellen - nur einmal
    if (d3.select('#legend').selectAll('*').empty()) {
        createLegend(colorScale, minShare, maxShare);
    }
}

function toggleRegion(region) {
    if (expandedRegions.has(region)) {
        expandedRegions.delete(region);
    } else {
        expandedRegions.add(region);
    }
    // Neu zeichnen
    processAndVisualize(rawData);
}

function createLegend(colorScale, minShare, maxShare) {
    d3.select('#legend').selectAll('*').remove();
    
    const legendDiv = d3.select('#legend');
    
    // SVG Container
    const svgContainer = legendDiv.append('div')
        .attr('id', 'legend-svg-container');
    
    const legendSvg = svgContainer.append('svg')
        .attr('viewBox', '0 0 560 78')
        .attr('width', '100%')
        .attr('height', '78');
    
    const gradient = legendSvg.append('defs')
        .append('linearGradient')
        .attr('id', 'legend-gradient')
        .attr('x1', '0%')
        .attr('x2', '100%');
    
    gradient.append('stop')
        .attr('offset', '0%')
        .attr('stop-color', '#7b3294');
    
    gradient.append('stop')
        .attr('offset', '50%')
        .attr('stop-color', '#f7f7f7');
    
    gradient.append('stop')
        .attr('offset', '100%')
        .attr('stop-color', '#e66101');
    
    // Gradient Balken
    legendSvg.append('rect')
        .attr('x', 70)
        .attr('y', 18)
        .attr('width', 420)
        .attr('height', 34)
        .attr('rx', 5)
        .style('fill', 'url(#legend-gradient)')
        .style('stroke', '#d1d5db')
        .style('stroke-width', '1.5');
    
    // Label "too little" (links, lila)
    legendSvg.append('text')
        .attr('x', 70)
        .attr('y', 70)
        .style('font-size', '14px')
        .style('font-weight', '600')
        .style('fill', '#7b3294')
        .style('text-anchor', 'start')
        .text('too little');
    
    // Label Mitte (weiß)
    legendSvg.append('text')
        .attr('x', 295)
        .attr('y', 70)
        .style('font-size', '14px')
        .style('font-weight', '600')
        .style('fill', '#6b7280')
        .style('text-anchor', 'middle')
        .text('balanced');
    
    // Label "too much" (rechts, orange)
    legendSvg.append('text')
        .attr('x', 490)
        .attr('y', 70)
        .style('font-size', '14px')
        .style('font-weight', '600')
        .style('fill', '#e66101')
        .style('text-anchor', 'end')
        .text('too much');
}

function showTooltip(event, row, category, dataPoint) {
    const tooltip = document.getElementById('tooltip');
    
    if (!dataPoint || dataPoint.share === null) return;
    
    let label = '';
    if (row.type === 'region') {
        label = `<strong>${row.region}</strong><br/>`;
    } else {
        label = `<strong>${row.area}</strong><br/>`;
    }
    
    label += `<strong>${category}</strong><br/>`;
    label += `<br/>Value: ${Math.round(dataPoint.value)} kcal<br/>`;
    label += `Target: ${Math.round(dataPoint.target)} kcal<br/>`;
    label += `Share: ${Math.round(dataPoint.share * 100)}%<br/>`;
    
    tooltip.innerHTML = label;
    tooltip.classList.add('visible');
    updateTooltipPosition(event);
}

function updateTooltipPosition(event) {
    const tooltip = document.getElementById('tooltip');
    const x = event.clientX + 15;
    const y = event.clientY - 10;
    tooltip.style.left = x + 'px';
    tooltip.style.top = y + 'px';
}

function hideTooltip() {
    const tooltip = document.getElementById('tooltip');
    tooltip.classList.remove('visible');
}