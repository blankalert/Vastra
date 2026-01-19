 // Initialize Lucide Icons
 lucide.createIcons();

 // --- GLOBAL CONFIGURATION ---
 const API_HEADERS = {
     "Content-Type": "application/json",
     "Authorization": "c7910440-dd74-11f0-adc4-cb199d9e47a1",
     "api-key": "va$Tra@pP",
     "device-type": "web",
     "udid": "50605585373614300053736"
 };
 const SALES_URL = "https://vastraapp.com/api/v2/packingSlip/export-packingslips";
 const STOCK_URL = "https://vastraapp.com/api/v2/report/get-designwise-stock-overall-detail-excel";

 // --- DOM REFERENCES ---
 const fetchBtn = document.getElementById('fetchBtn');
 const startDateInput = document.getElementById('startDate');
 const endDateInput = document.getElementById('endDate');
 const designFilterInput = document.getElementById('designFilter');
 const tableHeader = document.getElementById('tableHeader');
 const analysisBody = document.getElementById('analysisBody');
 const resultsContainer = document.getElementById('resultsContainer');
 const rowCount = document.getElementById('rowCount');
 const emptyState = document.getElementById('emptyState');
 
 // --- OVERLAY REFERENCES ---
 const loadingOverlay = document.getElementById('loadingOverlay');
 const loadingCircle = document.getElementById('loadingCircle');
 const progressPercent = document.getElementById('progressPercent');
 const progressTextContainer = document.getElementById('progressTextContainer');
 const successIcon = document.getElementById('successIcon');
 const errorIcon = document.getElementById('errorIcon');
 const loadingTitle = document.getElementById('loadingTitle');
 const loadingSubtitle = document.getElementById('loadingSubtitle');
 const closeOverlayBtn = document.getElementById('closeOverlayBtn');

 // --- STATE VARIABLES ---
 let allGroupedData = [];
 let filteredData = [];
 let monthColumns = [];

 // Set Default Dates (Dec 2025 - Jan 2026)
 startDateInput.value = "2025-12-01";
 endDateInput.value = "2026-01-31";

 /* * =========================================
  * SECTION 1: LOADING OVERLAY MANAGEMENT
  * Handles the visual state of the loading screen
  * =========================================
  */

 function showLoading() {
     loadingOverlay.classList.remove('overlay-enter');
     loadingOverlay.classList.add('overlay-active');
     
     // Reset UI states
     closeOverlayBtn.classList.add('hidden');
     successIcon.classList.replace('scale-in', 'scale-out');
     errorIcon.classList.replace('scale-in', 'scale-out');
     progressTextContainer.classList.replace('scale-out', 'scale-in');
     
     updateProgress(0, 'Starting...', 'Initializing request');
 }

 function hideLoading() {
     loadingOverlay.classList.remove('overlay-active');
     loadingOverlay.classList.add('overlay-enter');
 }

 function updateProgress(percent, title, subtitle) {
     // Update Circle (Circumference ~ 440px)
     const offset = 440 - (440 * percent) / 100;
     loadingCircle.style.strokeDashoffset = offset;
     
     // Update Text
     progressPercent.innerText = `${Math.round(percent)}%`;
     if(title) loadingTitle.innerText = title;
     if(subtitle) loadingSubtitle.innerText = subtitle;
 }

 function showSuccess() {
     progressTextContainer.classList.replace('scale-in', 'scale-out');
     successIcon.classList.replace('scale-out', 'scale-in');
     loadingTitle.innerText = "Analysis Complete!";
     loadingSubtitle.innerText = "Generating report view...";
     
     setTimeout(() => {
         hideLoading();
         emptyState.classList.add('hidden');
         resultsContainer.classList.remove('hidden');
     }, 1500);
 }

 function showError(msg) {
     progressTextContainer.classList.replace('scale-in', 'scale-out');
     errorIcon.classList.replace('scale-out', 'scale-in');
     loadingTitle.innerText = "Error Occurred";
     loadingSubtitle.innerText = msg;
     loadingSubtitle.classList.add('text-red-500');
     closeOverlayBtn.classList.remove('hidden');
 }

 /* * =========================================
  * SECTION 2: API & DATA FETCHING
  * Handles downloading and parsing Excel files
  * =========================================
  */

 function findDownloadUrl(obj) {
     if (!obj) return null;
     if (typeof obj === 'string' && (obj.startsWith('http') || obj.includes('.xlsx'))) return obj;
     const keys = ['data', 'file_url', 'url', 'file_path', 'link', 'download_url', 'filePath', 'result', 'export_url', 'fileUrl'];
     for (let k of keys) {
         if (obj[k] && typeof obj[k] === 'string' && (obj[k].startsWith('http') || obj[k].includes('export'))) return obj[k];
     }
     for (let k in obj) {
         if (typeof obj[k] === 'object') {
             const found = findDownloadUrl(obj[k]);
             if (found) return found;
         }
     }
     return null;
 }

 async function fetchExcel(url, payload) {
     try {
         const response = await fetch(url, { method: 'POST', headers: API_HEADERS, body: JSON.stringify(payload) });
         if (!response.ok) throw new Error(`HTTP ${response.status}`);
         
         const json = await response.json();
         const fileUrl = findDownloadUrl(json);
         
         if (!fileUrl) return []; // Return empty if no data

         const fileRes = await fetch(fileUrl);
         const blob = await fileRes.blob();
         const buffer = await blob.arrayBuffer();
         
         const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
         const sheet = workbook.Sheets[workbook.SheetNames[0]];
         return XLSX.utils.sheet_to_json(sheet, { defval: "" });

     } catch (err) {
         console.error(err);
         return [];
     }
 }

 /* * =========================================
  * SECTION 3: DATA PROCESSING UTILS
  * Helper functions for cleaning and matching data
  * =========================================
  */

 function cleanNumber(val) {
     if (typeof val === 'number') return val;
     let s = String(val || '0').replace(/[^0-9.-]/g, '');
     return parseFloat(s) || 0;
 }

 function cleanSalesDesignName(name) {
     if (!name) return 'Unknown';
     let cleaned = name;
     const parts = cleaned.split(/:\s*-\s*/);
     if (parts.length > 1) {
         cleaned = parts[parts.length - 1].trim();
     }
     cleaned = cleaned.replace(/^DES-/i, '').trim();
     return cleaned;
 }

 function findColumnValue(row, searchTerms, excludeTerms = []) {
     const keys = Object.keys(row);
     
     // Priority 1: Exact Match
     for (let term of searchTerms) {
         const exact = keys.find(k => k.toLowerCase().trim() === term.toLowerCase().trim());
         if (exact) return { val: cleanNumber(row[exact]), col: exact };
     }

     // Priority 2: Fuzzy Match
     for (let term of searchTerms) {
         const termClean = term.toLowerCase().replace(/[^a-z0-9]/g, '');
         const match = keys.find(k => {
             const keyClean = k.toLowerCase().replace(/[^a-z0-9]/g, '');
             const isExcluded = excludeTerms.some(ex => k.toLowerCase().includes(ex.toLowerCase()));
             return !isExcluded && keyClean.includes(termClean);
         });
         if (match) return { val: cleanNumber(row[match]), col: match };
     }
     return { val: 0, col: null };
 }

 /* * =========================================
  * SECTION 4: MAIN ANALYSIS LOGIC
  * Aggregates and merges data sources
  * =========================================
  */

 function runAnalysis(salesData, stockData) {
     const groups = {};
     const stockMap = {}; 

     updateProgress(60, 'Processing Data', 'Merging sales and inventory...');

     // 1. Process Stock (Aggregate by Design Name)
     stockData.forEach((row) => {
         const cleanRow = {};
         Object.keys(row).forEach(k => cleanRow[k.trim()] = row[k]);

         const stockRes = findColumnValue(cleanRow, ['Current Stock', 'Total Stock', 'Qty', 'Balance'], ['Pending', 'Order']);
         const pendingRes = findColumnValue(cleanRow, ['Pending Order', 'Pending Qty', 'Pending', 'Pend'], ['Amount', 'Value', 'Rate']);

         let rawDesign = String(cleanRow['Design Name'] || cleanRow['Design'] || cleanRow['Design Number'] || 'Unknown').trim();
         const design = cleanSalesDesignName(rawDesign);

         if (!stockMap[design]) stockMap[design] = { stock: 0, pending: 0, matched: false };
         stockMap[design].stock += stockRes.val;
         stockMap[design].pending += pendingRes.val;
     });

     // 2. Process Sales (Granular by Variant + Month)
     const monthSet = new Set();
     
     // Date Boundaries (Local Time construction)
     const sParts = startDateInput.value.split('-');
     const start = new Date(sParts[0], sParts[1]-1, sParts[2], 0, 0, 0, 0); 
     const eParts = endDateInput.value.split('-');
     const end = new Date(eParts[0], eParts[1]-1, eParts[2], 23, 59, 59, 999);

     salesData.forEach(row => {
         const cleanRow = {};
         Object.keys(row).forEach(k => cleanRow[k.trim()] = row[k]);

         const design = cleanSalesDesignName(String(cleanRow['Design Number'] || cleanRow['Design No'] || 'Unknown').trim());
         const color = String(cleanRow['Color'] || 'N/A').trim();
         const size = String(cleanRow['Size'] || 'N/A').trim();
         const qty = cleanNumber(cleanRow['Quantity']);
         
         // Date Parsing Logic
         let dateInput = cleanRow['Date'];
         let date;
         if (dateInput instanceof Date) {
             date = dateInput;
             if (date.getFullYear() < 2000 && date.getFullYear() > 1900) date.setFullYear(date.getFullYear() + 100);
         } else if (typeof dateInput === 'number') {
             date = new Date((dateInput - 25569) * 86400 * 1000);
         } else {
             const s = String(dateInput || '').trim();
             const parts = s.split(/[-/.\s]/);
             if (parts.length === 3) {
                  if (parts[0].length === 4) date = new Date(s);
                  else { 
                      let d = parseInt(parts[0]), m = parseInt(parts[1]) - 1, y = parseInt(parts[2]);
                      if (y < 100) y += 2000;
                      if (!isNaN(d) && !isNaN(m) && !isNaN(y)) date = new Date(y, m, d);
                  }
             }
         }

         if (!date || isNaN(date.getTime()) || date < start || date > end) return;

         const mKey = `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}`;
         const mLabel = date.toLocaleString('default', { month: 'short', year: 'numeric' });
         monthSet.add(JSON.stringify({key: mKey, label: mLabel}));

         const key = `${design}|${color}|${size}`;
         if (!groups[key]) groups[key] = { design, color, size, monthlyQty: {}, totalSales: 0 };
         groups[key].monthlyQty[mKey] = (groups[key].monthlyQty[mKey] || 0) + qty;
         groups[key].totalSales += qty;
     });

     monthColumns = Array.from(monthSet).map(s => JSON.parse(s)).sort((a,b) => a.key.localeCompare(b.key));

     // 3. Merge Stock into Sales Groups
     Object.values(groups).forEach(item => {
         const stockData = stockMap[item.design];
         if (stockData) {
             item.stockQty = stockData.stock;
             item.pendingQty = stockData.pending;
             stockData.matched = true;
         } else {
             item.stockQty = 0;
             item.pendingQty = 0;
         }
     });

     // 4. Add Stock-Only items
     Object.keys(stockMap).forEach(design => {
         if (!stockMap[design].matched) {
             const key = `${design}|-|N/A`;
             groups[key] = {
                 design, color: '-', size: '-', monthlyQty: {}, 
                 totalSales: 0, 
                 stockQty: stockMap[design].stock, 
                 pendingQty: stockMap[design].pending
             };
         }
     });

     // 5. Finalize and Sort
     const den = monthColumns.length || 1;
     allGroupedData = Object.values(groups).map(item => {
         item.avgSales = item.totalSales / den;
         return item;
     }).sort((a,b) => a.design.localeCompare(b.design, undefined, {numeric: true}));

     updateProgress(90, 'Finalizing', 'Rendering table view...');
     applyFilters();
     
     updateProgress(100, 'Done', 'Analysis ready');
     setTimeout(showSuccess, 500);
 }

 /* * =========================================
  * SECTION 5: EVENT HANDLERS
  * =========================================
  */

 fetchBtn.addEventListener('click', async () => {
     showLoading();
     try {
         const startMs = new Date(startDateInput.value).getTime();
         const endMs = new Date(endDateInput.value).getTime();

         updateProgress(20, 'Requesting Data', 'Contacting VastraApp servers...');

         const [sales, stock] = await Promise.all([
             fetchExcel(SALES_URL, { start_date: startMs, end_date: endMs, packing_slip_id: "" }),
             fetchExcel(STOCK_URL, { design_id: "", start_date: String(startMs), end_date: String(endMs), organization_ids: "" })
         ]);

         updateProgress(50, 'Data Received', 'Parsing Excel files...');
         runAnalysis(sales, stock);

     } catch (err) {
         showError(err.message);
     }
 });

 function applyFilters() {
     const val = designFilterInput.value.trim().toLowerCase();
     if (!val) {
         filteredData = [...allGroupedData];
     } else {
         const filters = val.split(/[\n,]+/).map(f => f.trim()).filter(f => f);
         filteredData = allGroupedData.filter(item => 
             filters.some(f => item.design.toLowerCase().includes(f))
         );
     }
     renderTable();
     rowCount.innerText = `${filteredData.length} items`;
 }

 designFilterInput.addEventListener('input', applyFilters);

 /* * =========================================
  * SECTION 6: RENDERING & EXPORT
  * =========================================
  */

 function renderTable() {
     let h = `<tr><th class="px-6 py-4 text-xs font-bold text-slate-500 uppercase sticky-col left-0 bg-slate-50 z-10 border-b">Design</th><th class="px-6 py-4 text-xs font-bold text-slate-500 uppercase border-b">Color</th><th class="px-6 py-4 text-xs font-bold text-slate-500 uppercase border-b">Size</th>`;
     monthColumns.forEach(m => h += `<th class="px-6 py-4 text-xs font-bold text-slate-500 uppercase text-right border-l border-b bg-slate-50/50">${m.label}</th>`);
     h += `<th class="px-6 py-4 text-xs font-bold text-indigo-600 uppercase text-right bg-indigo-50 border-l border-b">Sales</th><th class="px-6 py-4 text-xs font-bold text-amber-600 uppercase text-right bg-amber-50 border-l border-b">Avg</th><th class="px-6 py-4 text-xs font-bold text-emerald-600 uppercase text-right bg-emerald-50 border-l border-b">Stock</th><th class="px-6 py-4 text-xs font-bold text-rose-600 uppercase text-right bg-rose-50 border-l border-b">Pending</th></tr>`;
     
     tableHeader.innerHTML = h;

     if (filteredData.length === 0) {
        analysisBody.innerHTML = '<tr><td colspan="100%" class="px-6 py-4 text-center text-slate-500">No records found.</td></tr>';
        rowCount.innerText = '0 items';
        return;
     }

     const rowsHtml = filteredData.map(item => {
         let monthlyCells = '';
         monthColumns.forEach(m => {
             const val = item.monthlyQty[m.key] || 0;
             monthlyCells += `<td class="px-6 py-4 text-right border-l font-medium text-slate-700">${val.toLocaleString()}</td>`;
         });

         return `<tr class="hover:bg-indigo-50/30 transition-colors group">
             <td class="px-6 py-4 font-bold text-slate-900 sticky-col left-0 border-r bg-white group-hover:bg-indigo-50/30 shadow-[2px_0_5px_rgba(0,0,0,0.02)]">${item.design}</td>
             <td class="px-6 py-4 text-slate-600">${item.color}</td>
             <td class="px-6 py-4 text-slate-600">${item.size}</td>
             ${monthlyCells}
             <td class="px-6 py-4 text-right font-bold text-indigo-700 bg-indigo-50/30 border-l">${item.totalSales.toLocaleString()}</td>
             <td class="px-6 py-4 text-right font-bold text-amber-700 bg-amber-50/30 border-l">${item.avgSales.toFixed(1)}</td>
             <td class="px-6 py-4 text-right font-bold text-emerald-700 bg-emerald-50/30 border-l">${item.stockQty.toLocaleString()}</td>
             <td class="px-6 py-4 text-right font-bold text-rose-700 bg-rose-50/30 border-l">${item.pendingQty.toLocaleString()}</td>
         </tr>`;
     }).join('');

     analysisBody.innerHTML = rowsHtml;
     rowCount.innerText = `${filteredData.length} items`;
 }

 function exportToExcel(onlyFiltered = false) {
     const dataToExport = onlyFiltered ? filteredData : allGroupedData;
     
     if (!dataToExport.length) return;

     const excelData = dataToExport.map(item => {
         const row = {
             "Design Number": item.design,
             "Color": item.color,
             "Size": item.size
         };
         monthColumns.forEach(m => {
             row[m.label] = item.monthlyQty[m.key] || 0;
         });
         row["Total Sales"] = item.totalSales;
         row["Avg Sales"] = parseFloat(item.avgSales.toFixed(2));
         row["Current Stock"] = item.stockQty;
         row["Pending Order"] = item.pendingQty;
         return row;
     });

     const ws = XLSX.utils.json_to_sheet(excelData);
     const wscols = [ {wch: 20}, {wch: 15}, {wch: 10} ];
     monthColumns.forEach(() => wscols.push({wch: 12}));
     wscols.push({wch: 12}, {wch: 12}, {wch: 12}, {wch: 12});
     ws['!cols'] = wscols;

     const wb = XLSX.utils.book_new();
     XLSX.utils.book_append_sheet(wb, ws, "Sales Analysis");
     const filename = `analysis_${onlyFiltered ? 'filtered_' : ''}${new Date().toISOString().slice(0,10)}.xlsx`;
     XLSX.writeFile(wb, filename);
 }

