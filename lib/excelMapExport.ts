// lib/excelMapExport.ts
"use client";

type CountrySummary = {
    country: string;
    totalSales: number;
    totalRevenue: number;
    totalQuantity: number;
    uniqueWines: number;
    avgRevenuePerSale: number;
    topCategories: { name: string; count: number }[];
    topRegions: { name: string; count: number }[];
};

type RegionSummary = {
    region: string;
    country: string;
    totalSales: number;
    totalRevenue: number;
    totalQuantity: number;
    avgRevenuePerSale: number;
};

// Country coordinates for major wine-producing countries
const COUNTRY_COORDINATES: Record<string, [number, number]> = {
    'France': [46.2276, 2.2137],
    'Italy': [41.8719, 12.5674],
    'Spain': [40.4637, -3.7492],
    'United States': [37.0902, -95.7129],
    'USA': [37.0902, -95.7129],
    'Argentina': [-38.4161, -63.6167],
    'Chile': [-35.6751, -71.5430],
    'Australia': [-25.2744, 133.7751],
    'Germany': [51.1657, 10.4515],
    'Portugal': [39.3999, -8.2245],
    'South Africa': [-30.5595, 22.9375],
    'New Zealand': [-40.9006, 174.8860],
    'Austria': [47.5162, 14.5501],
    'Greece': [39.0742, 21.8243],
    'Hungary': [47.1625, 19.5033],
    'Romania': [45.9432, 24.9668],
    'Bulgaria': [42.7339, 25.4858],
    'Croatia': [45.1, 15.2],
    'Slovenia': [46.1512, 14.9955],
    'Canada': [56.1304, -106.3468],
    'Brazil': [-14.2350, -51.9253],
    'Uruguay': [-32.5228, -55.7658],
    'Mexico': [23.6345, -102.5528],
    'China': [35.8617, 104.1954],
    'Japan': [36.2048, 138.2529],
    'India': [20.5937, 78.9629],
    'Lebanon': [33.8547, 35.8623],
    'Israel': [31.0461, 34.8516],
    'Turkey': [38.9637, 35.2433],
    'Georgia': [42.3154, 43.3569],
    'Moldova': [47.4116, 28.3699],
    'Ukraine': [48.3794, 31.1656],
    'Switzerland': [46.8182, 8.2275],
    'England': [52.3555, -1.1743],
    'United Kingdom': [55.3781, -3.4360],
    'UK': [55.3781, -3.4360],
};

function getCountryColor(value: number, maxValue: number): string {
    const ratio = value / maxValue;
    
    if (ratio > 0.7) return '#166534'; // Dark green
    if (ratio > 0.5) return '#16a34a'; // Green
    if (ratio > 0.3) return '#22c55e'; // Light green
    if (ratio > 0.15) return '#86efac'; // Very light green
    return '#dcfce7'; // Pale green
}

function getMarkerSize(value: number, maxValue: number): number {
    const ratio = value / maxValue;
    const minSize = 8;
    const maxSize = 40;
    return minSize + (maxSize - minSize) * ratio;
}

export async function exportWineOriginWithMap(
    countrySummary: CountrySummary[],
    regionSummary: RegionSummary[]
): Promise<void> {
    const maxRevenue = Math.max(...countrySummary.map(c => c.totalRevenue));
    const maxWineCount = Math.max(...countrySummary.map(c => c.uniqueWines));
    
    // Prepare country data for the map with BOTH metrics
    const countryData = countrySummary
        .filter(country => COUNTRY_COORDINATES[country.country])
        .map(country => ({
            ...country,
            coordinates: COUNTRY_COORDINATES[country.country],
            revenueColor: getCountryColor(country.totalRevenue, maxRevenue),
            revenueSize: getMarkerSize(country.totalRevenue, maxRevenue),
            wineCountColor: getCountryColor(country.uniqueWines, maxWineCount),
            wineCountSize: getMarkerSize(country.uniqueWines, maxWineCount),
            revenuePercent: ((country.totalRevenue / maxRevenue) * 100).toFixed(1),
            wineCountPercent: ((country.uniqueWines / maxWineCount) * 100).toFixed(1),
        }));

    const htmlContent = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Wine Origin Analysis - Interactive Map</title>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #3b82f6 0%, #06b6d4 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5rem;
            margin-bottom: 10px;
            font-weight: 700;
        }
        
        .header p {
            font-size: 1.1rem;
            opacity: 0.95;
        }
        
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            padding: 30px 40px;
            background: #f8fafc;
            border-bottom: 1px solid #e2e8f0;
        }
        
        .stat-card {
            background: white;
            padding: 20px;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
            border-left: 4px solid #3b82f6;
        }
        
        .stat-label {
            font-size: 0.875rem;
            color: #64748b;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 8px;
        }
        
        .stat-value {
            font-size: 2rem;
            font-weight: 700;
            color: #1e293b;
        }
        
        .content {
            padding: 40px;
        }
        
        .map-container {
            position: relative;
            margin-bottom: 30px;
        }
        
        .toggle-container {
            position: absolute;
            top: 20px;
            right: 20px;
            z-index: 1000;
            background: white;
            border-radius: 12px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
            padding: 8px;
            display: flex;
            gap: 8px;
        }
        
        .toggle-btn {
            padding: 12px 24px;
            border: none;
            border-radius: 8px;
            font-weight: 600;
            font-size: 0.9rem;
            cursor: pointer;
            transition: all 0.3s ease;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .toggle-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.15);
        }
        
        .toggle-btn.active {
            background: linear-gradient(135deg, #3b82f6 0%, #06b6d4 100%);
            color: white;
        }
        
        .toggle-btn:not(.active) {
            background: #f1f5f9;
            color: #64748b;
        }
        
        .toggle-icon {
            font-size: 1.2rem;
        }
        
        #map {
            width: 100%;
            height: 600px;
            border-radius: 12px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
        }
        
        .legend {
            background: white;
            padding: 20px;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
            margin-bottom: 30px;
        }
        
        .legend h3 {
            font-size: 1.2rem;
            margin-bottom: 15px;
            color: #1e293b;
        }
        
        .legend-items {
            display: flex;
            flex-wrap: wrap;
            gap: 15px;
        }
        
        .legend-item {
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .legend-color {
            width: 24px;
            height: 24px;
            border-radius: 4px;
            border: 2px solid #e2e8f0;
        }
        
        .legend-label {
            font-size: 0.875rem;
            color: #475569;
        }
        
        .table-container {
            background: white;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
            overflow: hidden;
            margin-bottom: 30px;
        }
        
        .table-header {
            background: linear-gradient(135deg, #3b82f6 0%, #06b6d4 100%);
            color: white;
            padding: 20px;
            font-size: 1.3rem;
            font-weight: 600;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
        }
        
        thead {
            background: #f1f5f9;
        }
        
        th {
            padding: 15px;
            text-align: left;
            font-weight: 600;
            color: #475569;
            font-size: 0.875rem;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        td {
            padding: 15px;
            border-bottom: 1px solid #e2e8f0;
            color: #334155;
        }
        
        tbody tr:hover {
            background: #f8fafc;
        }
        
        .rank-badge {
            display: inline-block;
            padding: 4px 12px;
            border-radius: 20px;
            font-weight: 600;
            font-size: 0.875rem;
        }
        
        .rank-1 { background: #fbbf24; color: #78350f; }
        .rank-2 { background: #d1d5db; color: #1f2937; }
        .rank-3 { background: #fb923c; color: #7c2d12; }
        .rank-other { background: #e2e8f0; color: #475569; }
        
        .revenue {
            color: #16a34a;
            font-weight: 600;
        }
        
        .category-badges {
            display: flex;
            gap: 5px;
            flex-wrap: wrap;
        }
        
        .category-badge {
            background: #dbeafe;
            color: #1e40af;
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 0.75rem;
            font-weight: 500;
        }
        
        .footer {
            text-align: center;
            padding: 30px;
            background: #f8fafc;
            color: #64748b;
            font-size: 0.875rem;
        }
        
        .leaflet-popup-content-wrapper {
            border-radius: 12px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
        }
        
        .popup-title {
            font-size: 1.2rem;
            font-weight: 700;
            color: #1e293b;
            margin-bottom: 12px;
            padding-bottom: 8px;
            border-bottom: 2px solid #3b82f6;
        }
        
        .popup-stat {
            display: flex;
            justify-content: space-between;
            padding: 6px 0;
            font-size: 0.9rem;
        }
        
        .popup-stat-label {
            color: #64748b;
            font-weight: 500;
        }
        
        .popup-stat-value {
            color: #1e293b;
            font-weight: 600;
        }
        
        @media (max-width: 768px) {
            .header h1 {
                font-size: 1.8rem;
            }
            
            .content {
                padding: 20px;
            }
            
            #map {
                height: 400px;
            }
            
            .toggle-container {
                top: 10px;
                right: 10px;
                flex-direction: column;
            }
            
            .toggle-btn {
                padding: 10px 16px;
                font-size: 0.85rem;
            }
            
            table {
                font-size: 0.875rem;
            }
            
            th, td {
                padding: 10px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🍷 Wine Origin Analysis</h1>
            <p>Interactive Map & Performance Dashboard</p>
        </div>
        
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label">Total Countries</div>
                <div class="stat-value">${countrySummary.length}</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Total Revenue</div>
                <div class="stat-value">$${countrySummary.reduce((sum, c) => sum + c.totalRevenue, 0).toLocaleString()}</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Total Sales</div>
                <div class="stat-value">${countrySummary.reduce((sum, c) => sum + c.totalSales, 0).toLocaleString()}</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Top Country</div>
                <div class="stat-value" style="font-size: 1.5rem;">${countrySummary[0]?.country || 'N/A'}</div>
            </div>
        </div>
        
        <div class="content">
            <div class="map-container">
                <div class="toggle-container">
                    <button class="toggle-btn active" id="revenueBtn" onclick="switchView('revenue')">
                        <span>Revenue</span>
                    </button>
                    <button class="toggle-btn" id="wineCountBtn" onclick="switchView('wineCount')">
                        <span>Wine Count</span>
                    </button>
                </div>
                <div id="map"></div>
            </div>
            
            <div class="legend">
                <h3 id="legendTitle">Revenue Scale</h3>
                <div class="legend-items">
                    <div class="legend-item">
                        <div class="legend-color" style="background: #166534;"></div>
                        <span class="legend-label">Very High (70%+ of max)</span>
                    </div>
                    <div class="legend-item">
                        <div class="legend-color" style="background: #16a34a;"></div>
                        <span class="legend-label">High (50-70%)</span>
                    </div>
                    <div class="legend-item">
                        <div class="legend-color" style="background: #22c55e;"></div>
                        <span class="legend-label">Medium (30-50%)</span>
                    </div>
                    <div class="legend-item">
                        <div class="legend-color" style="background: #86efac;"></div>
                        <span class="legend-label">Low (15-30%)</span>
                    </div>
                    <div class="legend-item">
                        <div class="legend-color" style="background: #dcfce7;"></div>
                        <span class="legend-label">Very Low (<15%)</span>
                    </div>
                </div>
            </div>
            
            <div class="table-container">
                <div class="table-header">Country Performance Rankings</div>
                <table>
                    <thead>
                        <tr>
                            <th>Rank</th>
                            <th>Country</th>
                            <th style="text-align: right;">Total Sales</th>
                            <th style="text-align: right;">Revenue</th>
                            <th style="text-align: right;">Avg Sale</th>
                            <th style="text-align: right;">Unique Wines</th>
                            <th>Top Categories</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${countrySummary.map((country, idx) => `
                            <tr>
                                <td>
                                    <span class="rank-badge ${idx === 0 ? 'rank-1' : idx === 1 ? 'rank-2' : idx === 2 ? 'rank-3' : 'rank-other'}">
                                        ${idx + 1}
                                    </span>
                                </td>
                                <td style="font-weight: 600;">${country.country}</td>
                                <td style="text-align: right;">${country.totalSales.toLocaleString()}</td>
                                <td style="text-align: right;" class="revenue">$${country.totalRevenue.toLocaleString()}</td>
                                <td style="text-align: right;">$${country.avgRevenuePerSale.toFixed(2)}</td>
                                <td style="text-align: right;">${country.uniqueWines}</td>
                                <td>
                                    <div class="category-badges">
                                        ${country.topCategories.slice(0, 2).map(cat => 
                                            `<span class="category-badge">${cat.name}</span>`
                                        ).join('')}
                                    </div>
                                </td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
            
            <div class="table-container">
                <div class="table-header">Top Regions by Revenue</div>
                <table>
                    <thead>
                        <tr>
                            <th>Rank</th>
                            <th>Region</th>
                            <th>Country</th>
                            <th style="text-align: right;">Sales</th>
                            <th style="text-align: right;">Revenue</th>
                            <th style="text-align: right;">Avg Sale</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${regionSummary.slice(0, 15).map((region, idx) => `
                            <tr>
                                <td>
                                    <span class="rank-badge ${idx === 0 ? 'rank-1' : idx === 1 ? 'rank-2' : idx === 2 ? 'rank-3' : 'rank-other'}">
                                        ${idx + 1}
                                    </span>
                                </td>
                                <td style="font-weight: 600;">📍 ${region.region}</td>
                                <td>${region.country}</td>
                                <td style="text-align: right;">${region.totalSales.toLocaleString()}</td>
                                <td style="text-align: right;" class="revenue">$${region.totalRevenue.toLocaleString()}</td>
                                <td style="text-align: right;">$${region.avgRevenuePerSale.toFixed(2)}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        </div>
        
        <div class="footer">
            Generated on ${new Date().toLocaleDateString()} at ${new Date().toLocaleTimeString()} | Wine Origin Analysis Dashboard
        </div>
    </div>
    
    <script>
        // Initialize the map
        const map = L.map('map').setView([30, 0], 2);
        
        // Add tile layer
        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            attribution: '© OpenStreetMap contributors',
            maxZoom: 18,
        }).addTo(map);
        
        // Country data with both metrics
        const countryData = ${JSON.stringify(countryData)};
        
        // Store markers
        let markers = [];
        let currentView = 'revenue';
        
        // Function to create popup content
        function createPopupContent(country, view) {
            const highlightMetric = view === 'revenue' ? 'totalRevenue' : 'uniqueWines';
            const percentKey = view === 'revenue' ? 'revenuePercent' : 'wineCountPercent';
            
            return \`
                <div style="min-width: 250px;">
                    <div class="popup-title">\${country.country}</div>
                    <div class="popup-stat">
                        <span class="popup-stat-label">Total Sales:</span>
                        <span class="popup-stat-value">\${country.totalSales.toLocaleString()}</span>
                    </div>
                    <div class="popup-stat">
                        <span class="popup-stat-label">Total Revenue:</span>
                        <span class="popup-stat-value" style="color: \${view === 'revenue' ? '#16a34a' : '#1e293b'}; font-weight: \${view === 'revenue' ? '700' : '600'};">$\${country.totalRevenue.toLocaleString()}</span>
                    </div>
                    <div class="popup-stat">
                        <span class="popup-stat-label">Unique Wines:</span>
                        <span class="popup-stat-value" style="color: \${view === 'wineCount' ? '#7c3aed' : '#1e293b'}; font-weight: \${view === 'wineCount' ? '700' : '600'};">\${country.uniqueWines}</span>
                    </div>
                    <div class="popup-stat">
                        <span class="popup-stat-label">Total Quantity:</span>
                        <span class="popup-stat-value">\${country.totalQuantity.toLocaleString()}</span>
                    </div>
                    <div class="popup-stat">
                        <span class="popup-stat-label">Avg Revenue/Sale:</span>
                        <span class="popup-stat-value">$\${country.avgRevenuePerSale.toFixed(2)}</span>
                    </div>
                    <div style="margin-top: 12px; padding-top: 12px; border-top: 2px solid #3b82f6; background: #f8fafc; padding: 8px; border-radius: 6px;">
                        <div class="popup-stat-label" style="margin-bottom: 4px; color: #3b82f6;">
                            \${view === 'revenue' ? '💰 Revenue' : '🍷 Wine Variety'} Performance:
                        </div>
                        <div style="font-size: 1.1rem; font-weight: 700; color: #1e293b;">
                            \${country[percentKey]}% of maximum
                        </div>
                    </div>
                    \${country.topCategories.length > 0 ? \`
                        <div style="margin-top: 12px; padding-top: 12px; border-top: 1px solid #e2e8f0;">
                            <div class="popup-stat-label" style="margin-bottom: 6px;">Top Categories:</div>
                            \${country.topCategories.map(cat => 
                                \`<div style="font-size: 0.85rem; color: #475569; margin: 4px 0;">• \${cat.name} (\${cat.count} sales)</div>\`
                            ).join('')}
                        </div>
                    \` : ''}
                </div>
            \`;
        }
        
        // Function to render markers
        function renderMarkers(view) {
            // Clear existing markers
            markers.forEach(marker => map.removeLayer(marker));
            markers = [];
            
            // Update legend title
            document.getElementById('legendTitle').textContent = 
                view === 'revenue' ? 'Revenue Scale' : 'Wine Count Scale';
            
            // Add new markers
            countryData.forEach(country => {
                const [lat, lng] = country.coordinates;
                const color = view === 'revenue' ? country.revenueColor : country.wineCountColor;
                const size = view === 'revenue' ? country.revenueSize : country.wineCountSize;
                
                // Create circle marker
                const marker = L.circleMarker([lat, lng], {
                    radius: size,
                    fillColor: color,
                    color: '#fff',
                    weight: 2,
                    opacity: 1,
                    fillOpacity: 0.8
                }).addTo(map);
                
                // Bind popup
                marker.bindPopup(createPopupContent(country, view));
                
                // Add hover effect
                marker.on('mouseover', function() {
                    this.setStyle({
                        fillOpacity: 1,
                        weight: 3
                    });
                });
                
                marker.on('mouseout', function() {
                    this.setStyle({
                        fillOpacity: 0.8,
                        weight: 2
                    });
                });
                
                markers.push(marker);
            });
        }
        
        // Function to switch view
        function switchView(view) {
            currentView = view;
            
            // Update button states
            document.getElementById('revenueBtn').classList.toggle('active', view === 'revenue');
            document.getElementById('wineCountBtn').classList.toggle('active', view === 'wineCount');
            
            // Re-render markers
            renderMarkers(view);
        }
        
        // Initial render
        renderMarkers('revenue');
    </script>
</body>
</html>`;

    // Create and download the HTML file
    const blob = new Blob([htmlContent], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `Wine_Origin_Map_${new Date().toISOString().split('T')[0]}.html`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}