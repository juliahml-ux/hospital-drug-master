import json

def build():
    with open('output/115_drugs_enriched_preset.json', 'r', encoding='utf-8') as f:
        drugs = json.load(f)

    json_data = json.dumps(drugs, ensure_ascii=False)

    html_content = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>醫院藥品主檔與健保代碼價格管理系統</title>
  <!-- Tailwind CSS CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- SheetJS (xlsx) CDN -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <!-- Google Fonts -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Noto+Sans+TC:wght@300;400;500;700&display=swap" rel="stylesheet">
  <!-- Lucide Icons -->
  <script src="https://unpkg.com/lucide@latest"></script>
  <script>
    tailwind.config = {{
      theme: {{
        extend: {{
          fontFamily: {{
            sans: ['"Noto Sans TC"', 'Inter', 'system-ui', 'sans-serif'],
          }}
        }}
      }}
    }}
  </script>
  <style>
    body {{ background-color: #f8fafc; }}
    .custom-scrollbar::-webkit-scrollbar {{ width: 6px; height: 6px; }}
    .custom-scrollbar::-webkit-scrollbar-track {{ background: #f1f5f9; }}
    .custom-scrollbar::-webkit-scrollbar-thumb {{ background: #cbd5e1; border-radius: 3px; }}
    .custom-scrollbar::-webkit-scrollbar-thumb:hover {{ background: #94a3b8; }}
  </style>
</head>
<body class="text-slate-800 antialiased min-h-screen flex flex-col" id="dropZone">

  <!-- Drag and Drop Overlay -->
  <div id="dragOverlay" class="fixed inset-0 bg-teal-900/70 z-50 flex flex-col items-center justify-center text-white pointer-events-none hidden backdrop-blur-xs transition-all">
    <div class="p-8 rounded-3xl border-4 border-dashed border-teal-300 bg-teal-800/80 flex flex-col items-center space-y-4 max-w-md text-center shadow-2xl">
      <i data-lucide="file-up" class="w-16 h-16 text-teal-300 animate-bounce"></i>
      <h3 class="text-xl font-bold">釋放滑鼠以匯入 Excel 藥品清單</h3>
      <p class="text-sm text-teal-100">支援 .xlsx, .xls, .csv 檔案，自動解析代碼、品名、健保價與給付條件</p>
    </div>
  </div>

  <!-- Header / Navigation -->
  <header class="bg-white border-b border-slate-200 sticky top-0 z-30 shadow-xs">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
      <div class="flex items-center justify-between h-16">
        <!-- Logo & Title -->
        <div class="flex items-center space-x-3">
          <div class="w-10 h-10 rounded-xl bg-gradient-to-br from-teal-500 to-emerald-600 flex items-center justify-center text-white shadow-md shadow-teal-500/20">
            <i data-lucide="pill" class="w-6 h-6"></i>
          </div>
          <div>
            <div class="flex items-center space-x-2">
              <h1 class="text-lg font-bold text-slate-900 tracking-tight">醫院藥品主檔與健保代碼價格管理系統</h1>
              <span class="px-2 py-0.5 text-xs font-semibold rounded-full bg-teal-100 text-teal-800 border border-teal-200">v2.0 健保連線版</span>
            </div>
            <p class="text-xs text-slate-500">院內碼 • 健保碼 • 中文品名 • ATC 碼 • 健保支付價 • 給付規定章節與 PDF 連結</p>
          </div>
        </div>

        <!-- Quick Top Actions -->
        <div class="flex items-center space-x-2 sm:space-x-3">
          <button onclick="downloadExcelTemplate()" class="inline-flex items-center px-3 py-1.5 text-xs font-medium rounded-lg text-slate-700 bg-slate-100 hover:bg-slate-200 transition">
            <i data-lucide="file-spreadsheet" class="w-4 h-4 mr-1 text-emerald-600"></i>
            下載標準範本
          </button>
          <label class="cursor-pointer inline-flex items-center px-3 py-1.5 text-xs font-medium rounded-lg text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 transition shadow-xs">
            <i data-lucide="upload" class="w-4 h-4 mr-1 text-sky-600"></i>
            匯入 Excel / 清單
            <input type="file" id="excelFileInput" accept=".xlsx, .xls, .csv" class="hidden" onchange="handleExcelUpload(event)">
          </label>
          <button onclick="exportToExcel()" class="inline-flex items-center px-3 py-1.5 text-xs font-medium rounded-lg text-white bg-emerald-600 hover:bg-emerald-700 transition shadow-sm shadow-emerald-600/20">
            <i data-lucide="download" class="w-4 h-4 mr-1"></i>
            匯出 Excel (含健保價)
          </button>
        </div>
      </div>
    </div>
  </header>

  <!-- Main Content -->
  <main class="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 space-y-6">

    <!-- KPI Dashboard Metric Cards -->
    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
      <div class="bg-white rounded-xl p-4 border border-slate-200 shadow-xs flex items-center justify-between">
        <div>
          <p class="text-xs font-medium text-slate-500">已建置藥品總數</p>
          <p id="kpiTotalCount" class="text-2xl font-bold text-slate-900 mt-1">0</p>
        </div>
        <div class="w-11 h-11 rounded-lg bg-teal-50 text-teal-600 flex items-center justify-center">
          <i data-lucide="database" class="w-6 h-6"></i>
        </div>
      </div>

      <div class="bg-white rounded-xl p-4 border border-slate-200 shadow-xs flex items-center justify-between">
        <div>
          <p class="text-xs font-medium text-slate-500">健保給付品項 (已填健保碼)</p>
          <p id="kpiNhiCoveredCount" class="text-2xl font-bold text-emerald-600 mt-1">0</p>
        </div>
        <div class="w-11 h-11 rounded-lg bg-emerald-50 text-emerald-600 flex items-center justify-center">
          <i data-lucide="check-check" class="w-6 h-6"></i>
        </div>
      </div>

      <div class="bg-white rounded-xl p-4 border border-slate-200 shadow-xs flex items-center justify-between">
        <div>
          <p class="text-xs font-medium text-slate-500">自費 / 未收載品項</p>
          <p id="kpiSelfPayCount" class="text-2xl font-bold text-amber-600 mt-1">0</p>
        </div>
        <div class="w-11 h-11 rounded-lg bg-amber-50 text-amber-600 flex items-center justify-center">
          <i data-lucide="badge-dollar-sign" class="w-6 h-6"></i>
        </div>
      </div>

      <div class="bg-white rounded-xl p-4 border border-slate-200 shadow-xs flex items-center justify-between">
        <div>
          <p class="text-xs font-medium text-slate-500">附健保給付規定章節</p>
          <p id="kpiRuleCount" class="text-2xl font-bold text-indigo-600 mt-1">0</p>
        </div>
        <div class="w-11 h-11 rounded-lg bg-indigo-50 text-indigo-600 flex items-center justify-center">
          <i data-lucide="file-text" class="w-6 h-6"></i>
        </div>
      </div>
    </div>

    <!-- Toolbar: Search, Filters & Actions -->
    <div class="bg-white rounded-xl border border-slate-200 p-4 shadow-xs space-y-3">
      <div class="flex flex-col md:flex-row md:items-center justify-between gap-3">
        <div class="relative flex-1">
          <i data-lucide="search" class="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400"></i>
          <input 
            type="text" 
            id="searchInput" 
            placeholder="搜尋商品名、中文名、學名/成分、院內代碼、健保碼、ATC 碼、給付規定章節..." 
            oninput="handleSearch()"
            class="w-full pl-9 pr-4 py-2 text-sm bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 focus:bg-white transition"
          >
        </div>

        <div class="flex flex-wrap items-center gap-2">
          <!-- Dosage Form Filter -->
          <select id="dosageFilter" onchange="handleSearch()" class="text-xs bg-slate-50 border border-slate-200 rounded-lg px-2.5 py-2 text-slate-700 focus:ring-2 focus:ring-teal-500">
            <option value="">全部劑型 (All Forms)</option>
            <option value="O">口服劑型 (O)</option>
            <option value="I">注射劑型 (I)</option>
            <option value="E">外用劑型 (E)</option>
            <option value="1">1 級管制藥</option>
            <option value="2">2 級管制藥</option>
            <option value="3">3 級管制藥</option>
            <option value="4">4 級管制藥</option>
          </select>

          <!-- NHI Payment Filter -->
          <select id="nhiFilter" onchange="handleSearch()" class="text-xs bg-slate-50 border border-slate-200 rounded-lg px-2.5 py-2 text-slate-700 focus:ring-2 focus:ring-teal-500">
            <option value="all">全部健保狀態 (All)</option>
            <option value="nhi_covered">健保給付 (Covered)</option>
            <option value="has_rule">有給付規定 (With Rules)</option>
            <option value="self_pay">自費/未收載 (Self-Pay)</option>
            <option value="duplicate">院內碼重複 (Duplicate)</option>
          </select>

          <button onclick="openAddDrugModal()" class="inline-flex items-center px-3.5 py-2 text-xs font-semibold rounded-lg text-white bg-teal-600 hover:bg-teal-700 transition shadow-xs">
            <i data-lucide="plus-circle" class="w-4 h-4 mr-1.5"></i>
            新增單筆
          </button>
        </div>
      </div>

      <div class="flex flex-wrap items-center justify-between text-xs text-slate-500 pt-2 border-t border-slate-100 gap-2">
        <div class="flex items-center space-x-3">
          <span class="inline-flex items-center"><span class="w-2 h-2 rounded-full bg-emerald-500 mr-1.5"></span>綠標：健保給付品項</span>
          <span class="inline-flex items-center"><span class="w-2 h-2 rounded-full bg-indigo-500 mr-1.5"></span>藍標：附給付規定章節 PDF</span>
          <span class="inline-flex items-center"><span class="w-2 h-2 rounded-full bg-amber-500 mr-1.5"></span>黃標：自費/未收載</span>
          <span class="text-slate-400 font-normal ml-2 hidden lg:inline">💡 提示：點擊給付規定章節可直接下載官方 PDF 規範</span>
        </div>
        <div class="flex flex-wrap items-center space-x-2">
          <button onclick="loadEnrichedPreset()" class="text-teal-700 font-semibold hover:underline bg-teal-50 hover:bg-teal-100 px-2 py-1 rounded-md border border-teal-200 transition flex items-center space-x-1">
            <i data-lucide="refresh-cw" class="w-3.5 h-3.5"></i>
            <span>載入健保署比對版主檔 (761 筆)</span>
          </button>
          <span>•</span>
          <button onclick="clearAllDataConfirm()" class="text-rose-600 hover:underline">清空資料</button>
        </div>
      </div>
    </div>

    <!-- Drug Master Data Table -->
    <div class="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
      <div class="overflow-x-auto custom-scrollbar max-h-[640px]">
        <table class="w-full text-left border-collapse">
          <thead class="bg-slate-50 border-b border-slate-200 sticky top-0 z-10 text-xs font-semibold text-slate-600">
            <tr>
              <th class="py-3 px-3 w-10 text-center">序號</th>
              <th class="py-3 px-3 w-24 cursor-pointer hover:bg-slate-100 transition" onclick="toggleSort('hospitalCode')">
                <div class="flex items-center space-x-1">
                  <span>院內代碼</span>
                  <i data-lucide="arrow-up-down" class="w-3 h-3 text-slate-400"></i>
                </div>
              </th>
              <th class="py-3 px-3 w-28 cursor-pointer hover:bg-slate-100 transition" onclick="toggleSort('nhiCode')">
                <div class="flex items-center space-x-1">
                  <span>健保代碼</span>
                  <i data-lucide="arrow-up-down" class="w-3 h-3 text-slate-400"></i>
                </div>
              </th>
              <th class="py-3 px-3 min-w-[170px] cursor-pointer hover:bg-slate-100 transition" onclick="toggleSort('chineseName')">
                <div class="flex items-center space-x-1">
                  <span>健保核定中文名</span>
                  <i data-lucide="arrow-up-down" class="w-3 h-3 text-slate-400"></i>
                </div>
              </th>
              <th class="py-3 px-3 min-w-[200px] cursor-pointer hover:bg-slate-100 transition" onclick="toggleSort('brandName')">
                <div class="flex items-center space-x-1">
                  <span>英文商品名</span>
                  <i data-lucide="arrow-up-down" class="w-3 h-3 text-slate-400"></i>
                </div>
              </th>
              <th class="py-3 px-3 min-w-[160px]">學名 / 主成分</th>
              <th class="py-3 px-3 w-24 text-right cursor-pointer hover:bg-slate-100 transition" onclick="toggleSort('price')">
                <div class="flex items-center justify-end space-x-1">
                  <span>健保支付價</span>
                  <i data-lucide="arrow-up-down" class="w-3 h-3 text-slate-400"></i>
                </div>
              </th>
              <th class="py-3 px-3 min-w-[130px]">健保給付條件 / 規定</th>
              <th class="py-3 px-3 w-28">規格 / 包裝</th>
              <th class="py-3 px-3 w-16 text-center">劑型</th>
              <th class="py-3 px-3 w-20 text-center">查詢/操作</th>
            </tr>
          </thead>
          <tbody id="drugTableBody" class="divide-y divide-slate-100 text-xs text-slate-700"></tbody>
        </table>
      </div>

      <div class="p-4 border-t border-slate-200 bg-slate-50/50 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs text-slate-500">
        <div id="tableRecordSummary">顯示 0 到 0 筆，共 0 筆資料</div>
        <div class="flex items-center space-x-2">
          <span>每頁顯示</span>
          <select id="pageSizeSelect" onchange="changePageSize(this.value)" class="bg-white border border-slate-200 rounded px-2 py-1 text-xs">
            <option value="15">15 筆</option>
            <option value="30">30 筆</option>
            <option value="50">50 筆</option>
            <option value="100">100 筆</option>
            <option value="100000">全部顯示</option>
          </select>
          <div class="inline-flex rounded-md shadow-2xs">
            <button onclick="prevPage()" id="btnPrevPage" class="px-2.5 py-1 rounded-l-md border border-slate-200 bg-white hover:bg-slate-50 disabled:opacity-40">上一頁</button>
            <span id="currentPageIndicator" class="px-3 py-1 border-y border-slate-200 bg-slate-100 font-medium text-slate-700">1</span>
            <button onclick="nextPage()" id="btnNextPage" class="px-2.5 py-1 rounded-r-md border border-slate-200 bg-white hover:bg-slate-50 disabled:opacity-40">下一頁</button>
          </div>
        </div>
      </div>
    </div>

  </main>

  <footer class="bg-white border-t border-slate-200 py-4 mt-8">
    <div class="max-w-7xl mx-auto px-4 text-center text-xs text-slate-500 space-y-1">
      <p>醫院藥品主檔與健保代碼價格管理系統 • 資料來源：衛生福利部中央健康保險署開放資料庫</p>
      <p class="text-slate-400">本系統純前端本機運作，資料安全保密不外傳 • 支援健保代碼、中文名、支付價與給付規定 PDF 直連</p>
    </div>
  </footer>

  <!-- Modal: Add / Edit Drug -->
  <div id="drugModal" class="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-xs flex items-center justify-center hidden p-4">
    <div class="bg-white rounded-2xl max-w-2xl w-full shadow-2xl border border-slate-100 overflow-hidden transform transition-all">
      <div class="px-6 py-4 bg-gradient-to-r from-teal-600 to-emerald-600 text-white flex items-center justify-between">
        <div class="flex items-center space-x-2">
          <i data-lucide="edit-3" class="w-5 h-5"></i>
          <h3 id="modalTitle" class="text-base font-bold">編輯藥品主檔與健保資料</h3>
        </div>
        <button onclick="closeDrugModal()" class="text-white/80 hover:text-white transition">
          <i data-lucide="x" class="w-5 h-5"></i>
        </button>
      </div>

      <form id="drugForm" onsubmit="saveDrugForm(event)" class="p-6 space-y-4 text-xs">
        <input type="hidden" id="formDrugId">

        <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
          <div>
            <label class="block font-semibold text-slate-700 mb-1">院內代碼 <span class="text-rose-500">*</span></label>
            <input type="text" id="formHospitalCode" required placeholder="例：ONORV5" class="w-full px-3 py-2 border border-slate-300 rounded-lg uppercase tracking-wider focus:ring-2 focus:ring-teal-500 focus:outline-hidden font-mono">
            <p id="formHospitalCodeWarn" class="text-rose-500 text-[11px] mt-1 hidden"></p>
          </div>

          <div>
            <div class="flex items-center justify-between mb-1">
              <label class="font-semibold text-slate-700">健保代碼</label>
              <span id="nhiLengthBadge" class="text-[10px] px-1.5 py-0.5 rounded bg-slate-100 text-slate-500 font-mono">0/10碼</span>
            </div>
            <input type="text" id="formNhiCode" maxlength="10" placeholder="例：BC19248100" oninput="handleNhiInput(this.value)" class="w-full px-3 py-2 border border-slate-300 rounded-lg uppercase tracking-wider font-mono focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
          </div>

          <div>
            <label class="block font-semibold text-slate-700 mb-1">健保支付價 (NT$)</label>
            <input type="text" id="formPrice" placeholder="例：29.60" class="w-full px-3 py-2 border border-slate-300 rounded-lg font-mono focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
          </div>
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <label class="block font-semibold text-slate-700 mb-1">健保中文品名 (Chinese Name)</label>
            <input type="text" id="formChineseName" placeholder="例：脈優錠 5 毫克" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
          </div>

          <div>
            <label class="block font-semibold text-slate-700 mb-1">ATC 碼</label>
            <input type="text" id="formAtc" placeholder="例：C08CA01" class="w-full px-3 py-2 border border-slate-300 rounded-lg uppercase tracking-wider font-mono focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
          </div>
        </div>

        <div>
          <label class="block font-semibold text-slate-700 mb-1">英文商品名 (Brand Name) <span class="text-rose-500">*</span></label>
          <input type="text" id="formBrandName" required placeholder="例：Norvasc Tablets 5mg" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
        </div>

        <div>
          <label class="block font-semibold text-slate-700 mb-1">學名 / 主成分 (Generic Name) <span class="text-rose-500">*</span></label>
          <input type="text" id="formGenericName" required placeholder="例：Amlodipine besylate" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <label class="block font-semibold text-slate-700 mb-1">健保給付規定章節</label>
            <input type="text" id="formRuleSection" placeholder="例：1.2.2.2. 或 8.2.1." class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
          </div>
          <div>
            <label class="block font-semibold text-slate-700 mb-1">給付規定 PDF 連結</label>
            <input type="text" id="formRuleLink" placeholder="https://info.nhi.gov.tw/api/..." class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden font-mono text-[11px]">
          </div>
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <div>
            <label class="block font-semibold text-slate-700 mb-1">規格劑量 / 包裝 (Strength/Unit)</label>
            <input type="text" id="formStrength" placeholder="例：5mg/tab (Box)" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
          </div>

          <div>
            <label class="block font-semibold text-slate-700 mb-1">劑型 (Dosage Form)</label>
            <input type="text" id="formDosageForm" placeholder="例：O, I, E, 錠劑、膠囊劑" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
          </div>
        </div>

        <div class="pt-4 border-t border-slate-100 flex items-center justify-end space-x-3">
          <button type="button" onclick="closeDrugModal()" class="px-4 py-2 rounded-lg border border-slate-300 text-slate-600 hover:bg-slate-50 transition">取消</button>
          <button type="submit" class="px-5 py-2 rounded-lg bg-teal-600 hover:bg-teal-700 text-white font-medium shadow-sm transition">儲存資料</button>
        </div>
      </form>
    </div>
  </div>

  <!-- Toast -->
  <div id="toast" class="fixed bottom-6 right-6 z-50 transform translate-y-20 opacity-0 transition-all duration-300 flex items-center space-x-2 px-4 py-3 rounded-xl shadow-lg text-xs font-medium text-white bg-slate-900">
    <i id="toastIcon" data-lucide="info" class="w-4 h-4 text-teal-400"></i>
    <span id="toastMsg">操作成功</span>
  </div>

  <script>
    // Embedded 761 Clinical Drugs with Matched NHI Codes, Prices and Payment Rules
    const HOSPITAL_ENRICHED_PRESET = {json_data};

    const STORAGE_KEY = "hospital_drug_master_db_v4";
    let drugs = [];
    let currentPage = 1;
    let pageSize = 15;
    let sortColumn = "hospitalCode";
    let sortAsc = true;
    let pendingImportData = null;

    window.addEventListener("DOMContentLoaded", () => {{
      loadDataFromStorage();
      renderAll();
      lucide.createIcons();
      setupDragAndDrop();
    }});

    function loadDataFromStorage() {{
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {{
        try {{
          drugs = JSON.parse(saved);
        }} catch(e) {{
          drugs = JSON.parse(JSON.stringify(HOSPITAL_ENRICHED_PRESET));
        }}
      }} else {{
        drugs = JSON.parse(JSON.stringify(HOSPITAL_ENRICHED_PRESET));
        saveDataToStorage();
      }}
    }}

    function saveDataToStorage() {{
      localStorage.setItem(STORAGE_KEY, JSON.stringify(drugs));
    }}

    function loadEnrichedPreset() {{
      if (confirm("確定要重設載入「115 年健保署比對版院內藥品清單 (761 筆，含健保價與給付規定)」嗎？")) {{
        drugs = JSON.parse(JSON.stringify(HOSPITAL_ENRICHED_PRESET));
        saveDataToStorage();
        renderAll();
        showToast(`已成功載入 761 筆院內藥品主檔！`, "success");
      }}
    }}

    function clearAllDataConfirm() {{
      if (confirm("警告：確定要清空所有藥品資料嗎？請確保您已匯出 Excel 備份。")) {{
        drugs = [];
        saveDataToStorage();
        renderAll();
        showToast("已清空所有藥品資料", "info");
      }}
    }}

    function validateDrug(drug, allDrugs) {{
      const issues = [];
      const nhi = (drug.nhiCode || "").trim();
      const isSelfPay = !nhi || nhi === '自費/未收載';

      const hCode = (drug.hospitalCode || "").trim().toUpperCase();
      if (!hCode) {{
        issues.push({{ type: "hcode_empty", level: "rose", text: "院內碼未填" }});
      }} else {{
        const duplicates = allDrugs.filter(d => (d.hospitalCode || "").trim().toUpperCase() === hCode);
        if (duplicates.length > 1) {{
          issues.push({{ type: "duplicate", level: "rose", text: "院內碼重複" }});
        }}
      }}

      if (!drug.brandName || !drug.genericName) {{
        issues.push({{ type: "incomplete", level: "indigo", text: "缺商品名或成分" }});
      }}
      return issues;
    }}

    function renderAll() {{
      updateKpis();
      renderTable();
      lucide.createIcons();
    }}

    function updateKpis() {{
      const total = drugs.length;
      let nhiCovered = 0;
      let selfPayCount = 0;
      let ruleCount = 0;

      const codeMap = {{}};
      drugs.forEach(d => {{
        const c = (d.hospitalCode || "").trim().toUpperCase();
        if (c) codeMap[c] = (codeMap[c] || 0) + 1;
        const nhi = (d.nhiCode || "").trim();
        if (nhi && nhi !== '自費/未收載' && nhi.length === 10) {{
          nhiCovered++;
        }} else {{
          selfPayCount++;
        }}
        if (d.ruleSection && d.ruleSection.trim() && d.ruleSection !== '無特殊章節') {{
          ruleCount++;
        }}
      }});

      document.getElementById("kpiTotalCount").innerText = total;
      document.getElementById("kpiNhiCoveredCount").innerText = nhiCovered;
      document.getElementById("kpiSelfPayCount").innerText = selfPayCount;
      document.getElementById("kpiRuleCount").innerText = ruleCount;
    }}

    function getFilteredDrugs() {{
      const q = document.getElementById("searchInput").value.trim().toLowerCase();
      const dosageFilter = document.getElementById("dosageFilter").value;
      const nhiFilter = document.getElementById("nhiFilter").value;

      return drugs.filter(drug => {{
        const matchQuery = !q || 
          (drug.hospitalCode && drug.hospitalCode.toLowerCase().includes(q)) ||
          (drug.nhiCode && drug.nhiCode.toLowerCase().includes(q)) ||
          (drug.atc && drug.atc.toLowerCase().includes(q)) ||
          (drug.brandName && drug.brandName.toLowerCase().includes(q)) ||
          (drug.chineseName && drug.chineseName.toLowerCase().includes(q)) ||
          (drug.genericName && drug.genericName.toLowerCase().includes(q)) ||
          (drug.strength && drug.strength.toLowerCase().includes(q)) ||
          (drug.dosageForm && drug.dosageForm.toLowerCase().includes(q)) ||
          (drug.ruleSection && drug.ruleSection.toLowerCase().includes(q));

        if (!matchQuery) return false;

        if (dosageFilter && !(drug.dosageForm || "").includes(dosageFilter)) {{
          return false;
        }}

        if (nhiFilter !== "all") {{
          const isCovered = drug.nhiCode && drug.nhiCode.trim() && drug.nhiCode !== '自費/未收載' && drug.nhiCode.trim().length === 10;
          const hasRule = drug.ruleSection && drug.ruleSection.trim() && drug.ruleSection !== '無特殊章節';
          if (nhiFilter === "nhi_covered" && !isCovered) return false;
          if (nhiFilter === "has_rule" && !hasRule) return false;
          if (nhiFilter === "self_pay" && isCovered) return false;
          if (nhiFilter === "duplicate") {{
            const issues = validateDrug(drug, drugs);
            if (!issues.some(i => i.type === "duplicate")) return false;
          }}
        }}

        return true;
      }});
    }}

    function toggleSort(col) {{
      if (sortColumn === col) {{
        sortAsc = !sortAsc;
      }} else {{
        sortColumn = col;
        sortAsc = true;
      }}
      renderTable();
      lucide.createIcons();
    }}

    function renderTable() {{
      const filtered = getFilteredDrugs();

      filtered.sort((a, b) => {{
        let valA = a[sortColumn] || "";
        let valB = b[sortColumn] || "";
        if (sortColumn === "price") {{
          valA = parseFloat(valA) || 0;
          valB = parseFloat(valB) || 0;
          return sortAsc ? valA - valB : valB - valA;
        }}
        valA = valA.toString().toLowerCase();
        valB = valB.toString().toLowerCase();
        if (valA < valB) return sortAsc ? -1 : 1;
        if (valA > valB) return sortAsc ? 1 : -1;
        return 0;
      }});

      const totalRecords = filtered.length;
      const totalPages = Math.ceil(totalRecords / pageSize) || 1;
      if (currentPage > totalPages) currentPage = totalPages;
      if (currentPage < 1) currentPage = 1;

      const startIndex = (currentPage - 1) * pageSize;
      const paginated = filtered.slice(startIndex, startIndex + pageSize);

      const tbody = document.getElementById("drugTableBody");
      if (paginated.length === 0) {{
        tbody.innerHTML = `
          <tr>
            <td colspan="11" class="py-12 text-center text-slate-400">
              <i data-lucide="inbox" class="w-10 h-10 mx-auto mb-2 text-slate-300"></i>
              <p class="text-sm font-medium">查無符合條件之藥品資料</p>
              <p class="text-xs text-slate-400 mt-1">請嘗試調整搜尋關鍵字或點擊「新增單筆」建立主檔</p>
            </td>
          </tr>
        `;
      }} else {{
        tbody.innerHTML = paginated.map((drug, idx) => {{
          const rowNum = startIndex + idx + 1;
          const issues = validateDrug(drug, drugs);
          
          const hasNhi = !!(drug.nhiCode || "").trim() && drug.nhiCode !== '自費/未收載';
          const isNhiBad = hasNhi && (drug.nhiCode || "").trim().length !== 10;
          
          let nhiDisplay = `<span class="text-amber-700 bg-amber-50 px-2 py-0.5 rounded text-[11px] font-medium border border-amber-200">自費/未收載</span>`;
          if (hasNhi) {{
            if (isNhiBad) {{
              nhiDisplay = `<span class="text-amber-700 bg-amber-50 border border-amber-200 px-1.5 py-0.5 rounded font-mono font-semibold">${{escapeHtml(drug.nhiCode)}}</span>`;
            }} else {{
              nhiDisplay = `<span class="font-mono text-emerald-800 bg-emerald-50 border border-emerald-200 px-1.5 py-0.5 rounded font-bold">${{escapeHtml(drug.nhiCode)}}</span>`;
            }}
          }}

          // Price Display
          let priceDisplay = `<span class="text-slate-400 text-xs italic">-</span>`;
          if (drug.price && drug.price.toString().trim()) {{
            priceDisplay = `<span class="font-mono font-bold text-slate-900 text-xs">NT$ ${{parseFloat(drug.price).toFixed(2)}}</span>`;
          }}

          // Payment Rules Badge / Link
          let ruleDisplay = `<span class="text-slate-400 text-[11px]">無特殊規定</span>`;
          if (drug.ruleSection && drug.ruleSection.trim() && drug.ruleSection !== '無特殊章節') {{
            if (drug.ruleLink && drug.ruleLink.trim()) {{
              ruleDisplay = `
                <a href="${{drug.ruleLink}}" target="_blank" class="inline-flex items-center px-2 py-0.5 rounded text-[11px] font-semibold bg-indigo-50 text-indigo-700 border border-indigo-200 hover:bg-indigo-100 transition shadow-2xs group-hover:border-indigo-300" title="點擊檢視健保給付規定章節 PDF">
                  <i data-lucide="file-text" class="w-3 h-3 mr-1 text-indigo-500"></i>
                  ${{escapeHtml(drug.ruleSection)}} PDF
                </a>
              `;
            }} else {{
              ruleDisplay = `<span class="inline-flex items-center px-2 py-0.5 rounded text-[11px] font-medium bg-slate-100 text-slate-700 font-mono">${{escapeHtml(drug.ruleSection)}}</span>`;
            }}
          }}

          // Search / Query External Link
          let queryUrl = drug.nhiLink;
          if (!queryUrl || queryUrl === 'https://info.nhi.gov.tw/IODE0000/IODE0000S06') {{
            queryUrl = `https://info.nhi.gov.tw/IODE0000/IODE0000S06`;
          }}

          return `
            <tr class="hover:bg-teal-50/40 transition group border-b border-slate-100">
              <td class="py-2.5 px-3 text-center font-mono text-slate-400 text-xs">${{rowNum}}</td>
              <td class="py-2.5 px-3 font-semibold text-slate-900 font-mono tracking-wide">
                <span class="bg-slate-100 text-slate-800 px-1.5 py-0.5 rounded border border-slate-200 text-xs">${{escapeHtml(drug.hospitalCode || "-")}}</span>
              </td>
              <td class="py-2.5 px-3">${{nhiDisplay}}</td>
              <td class="py-2.5 px-3 font-medium text-slate-900">
                <div class="line-clamp-2" title="${{escapeHtml(drug.chineseName || '-')}}">${{escapeHtml(drug.chineseName || "-")}}</div>
              </td>
              <td class="py-2.5 px-3 font-medium text-slate-800">
                <div class="line-clamp-2" title="${{escapeHtml(drug.brandName || '-')}}">${{escapeHtml(drug.brandName || "-")}}</div>
                ${{drug.atc ? `<span class="text-[10px] font-mono text-teal-700 bg-teal-50/80 px-1 py-0.2 rounded border border-teal-100 mr-1">${{escapeHtml(drug.atc)}}</span>` : ''}}
              </td>
              <td class="py-2.5 px-3 text-slate-600 italic">
                <div class="line-clamp-2 text-[11px]" title="${{escapeHtml(drug.genericName || '-')}}">${{escapeHtml(drug.genericName || "-")}}</div>
              </td>
              <td class="py-2.5 px-3 text-right">${{priceDisplay}}</td>
              <td class="py-2.5 px-3">${{ruleDisplay}}</td>
              <td class="py-2.5 px-3 text-slate-600 text-[11px]">${{escapeHtml(drug.strength || "-")}}</td>
              <td class="py-2.5 px-3 text-center">
                <span class="bg-slate-100 text-slate-700 px-1.5 py-0.5 rounded text-[11px] font-medium font-mono">${{escapeHtml(drug.dosageForm || "-")}}</span>
              </td>
              <td class="py-2.5 px-3 text-center">
                <div class="flex items-center justify-center space-x-1">
                  <a href="${{queryUrl}}" target="_blank" title="健保署/食藥署官網品項查詢" class="p-1.5 text-slate-400 hover:text-sky-600 hover:bg-sky-50 rounded transition">
                    <i data-lucide="external-link" class="w-3.5 h-3.5"></i>
                  </a>
                  <button onclick="openEditDrugModal('${{drug.id}}')" title="編輯藥品" class="p-1.5 text-slate-400 hover:text-teal-600 hover:bg-teal-50 rounded transition">
                    <i data-lucide="edit-2" class="w-3.5 h-3.5"></i>
                  </button>
                  <button onclick="deleteDrug('${{drug.id}}')" title="刪除" class="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded transition">
                    <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
                  </button>
                </div>
              </td>
            </tr>
          `;
        }}).join("");
      }}

      const startDisplay = totalRecords === 0 ? 0 : startIndex + 1;
      const endDisplay = Math.min(startIndex + pageSize, totalRecords);
      document.getElementById("tableRecordSummary").innerText = `顯示第 ${{startDisplay}} 到 ${{endDisplay}} 筆，共 ${{totalRecords}} 筆資料`;
      document.getElementById("currentPageIndicator").innerText = `${{currentPage}} / ${{totalPages}}`;
      document.getElementById("btnPrevPage").disabled = (currentPage <= 1);
      document.getElementById("btnNextPage").disabled = (currentPage >= totalPages);
    }}

    function handleSearch() {{
      currentPage = 1;
      renderTable();
      lucide.createIcons();
    }}

    function changePageSize(val) {{
      pageSize = parseInt(val);
      currentPage = 1;
      renderTable();
      lucide.createIcons();
    }}

    function prevPage() {{
      if (currentPage > 1) {{
        currentPage--;
        renderTable();
        lucide.createIcons();
      }}
    }}

    function nextPage() {{
      const filtered = getFilteredDrugs();
      const totalPages = Math.ceil(filtered.length / pageSize) || 1;
      if (currentPage < totalPages) {{
        currentPage++;
        renderTable();
        lucide.createIcons();
      }}
    }}

    function openAddDrugModal() {{
      document.getElementById("modalTitle").innerText = "新增藥品主檔";
      document.getElementById("formDrugId").value = "";
      document.getElementById("drugForm").reset();
      handleNhiInput("");
      document.getElementById("formHospitalCodeWarn").classList.add("hidden");
      document.getElementById("drugModal").classList.remove("hidden");
      lucide.createIcons();
      setTimeout(() => document.getElementById("formHospitalCode").focus(), 50);
    }}

    function openEditDrugModal(id) {{
      const drug = drugs.find(d => d.id === id);
      if (!drug) return;

      document.getElementById("modalTitle").innerText = "編輯藥品主檔與健保資料";
      document.getElementById("formDrugId").value = drug.id;
      document.getElementById("formHospitalCode").value = drug.hospitalCode || "";
      document.getElementById("formNhiCode").value = drug.nhiCode || "";
      document.getElementById("formPrice").value = drug.price || "";
      document.getElementById("formAtc").value = drug.atc || "";
      document.getElementById("formBrandName").value = drug.brandName || "";
      document.getElementById("formChineseName").value = drug.chineseName || "";
      document.getElementById("formGenericName").value = drug.genericName || "";
      document.getElementById("formRuleSection").value = drug.ruleSection || "";
      document.getElementById("formRuleLink").value = drug.ruleLink || "";
      document.getElementById("formStrength").value = drug.strength || "";
      document.getElementById("formDosageForm").value = drug.dosageForm || "";

      handleNhiInput(drug.nhiCode || "");
      document.getElementById("formHospitalCodeWarn").classList.add("hidden");
      document.getElementById("drugModal").classList.remove("hidden");
      lucide.createIcons();
    }}

    function closeDrugModal() {{
      document.getElementById("drugModal").classList.add("hidden");
    }}

    function handleNhiInput(val) {{
      const len = val.trim().length;
      const badge = document.getElementById("nhiLengthBadge");
      badge.innerText = `${{len}}/10碼`;

      if (len === 10) {{
        badge.className = "text-[10px] px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-700 font-mono font-bold";
      }} else if (len === 0) {{
        badge.className = "text-[10px] px-1.5 py-0.5 rounded bg-slate-100 text-slate-500 font-mono";
      }} else {{
        badge.className = "text-[10px] px-1.5 py-0.5 rounded bg-amber-100 text-amber-700 font-mono font-bold";
      }}
    }}

    function saveDrugForm(e) {{
      e.preventDefault();
      const id = document.getElementById("formDrugId").value;
      const hospitalCode = document.getElementById("formHospitalCode").value.trim().toUpperCase();
      const nhiCode = document.getElementById("formNhiCode").value.trim().toUpperCase();
      const price = document.getElementById("formPrice").value.trim();
      const atc = document.getElementById("formAtc").value.trim().toUpperCase();
      const brandName = document.getElementById("formBrandName").value.trim();
      const chineseName = document.getElementById("formChineseName").value.trim();
      const genericName = document.getElementById("formGenericName").value.trim();
      const ruleSection = document.getElementById("formRuleSection").value.trim();
      const ruleLink = document.getElementById("formRuleLink").value.trim();
      const strength = document.getElementById("formStrength").value.trim();
      const dosageForm = document.getElementById("formDosageForm").value.trim();

      const duplicate = drugs.find(d => d.hospitalCode.toUpperCase() === hospitalCode && d.id !== id);
      if (duplicate) {{
        const warn = document.getElementById("formHospitalCodeWarn");
        warn.innerText = `院內代碼 ${{hospitalCode}} 已存在於「${{duplicate.brandName}}」！`;
        warn.classList.remove("hidden");
        return;
      }}

      if (id) {{
        const index = drugs.findIndex(d => d.id === id);
        if (index !== -1) {{
          drugs[index] = {{ ...drugs[index], id, hospitalCode, nhiCode, price, atc, brandName, chineseName, genericName, ruleSection, ruleLink, strength, dosageForm }};
          showToast("已更新藥品主檔與健保資料", "success");
        }}
      }} else {{
        const newDrug = {{
          id: "d_" + Date.now(),
          hospitalCode,
          nhiCode,
          price,
          atc,
          brandName,
          chineseName,
          genericName,
          ruleSection,
          ruleLink,
          strength,
          dosageForm,
          nhiLink: "https://info.nhi.gov.tw/IODE0000/IODE0000S06"
        }};
        drugs.unshift(newDrug);
        showToast("已成功新增藥品主檔", "success");
      }}

      saveDataToStorage();
      closeDrugModal();
      renderAll();
    }}

    function deleteDrug(id) {{
      const drug = drugs.find(d => d.id === id);
      if (!drug) return;
      if (confirm(`確定要刪除藥品「${{drug.brandName}} (${{drug.hospitalCode}})`)) {{
        drugs = drugs.filter(d => d.id !== id);
        saveDataToStorage();
        renderAll();
        showToast("已刪除該藥品品項", "info");
      }}
    }}

    function downloadExcelTemplate() {{
      const headers = ["院內代碼", "健保代碼", "ATC碼", "英文商品名", "健保中文品名", "學名/成分", "規格劑量", "劑型", "健保支付價", "健保給付規定章節", "給付規定PDF連結"];
      const templateData = [
        headers,
        ["ONORV5", "BC19248100", "C08CA01", "Norvasc Tablets 5mg", "脈優錠 5 毫克", "Amlodipine besylate", "5mg/tab", "O", "5.60", "無特殊章節", ""],
        ["OGLUC5", "AC34125100", "A10BA02", "Glucophage Tablets 500mg", "庫魯化錠 500 毫克", "Metformin hydrochloride", "500mg/tab", "O", "1.50", "無特殊章節", ""],
        ["ARIO25", "BC27318100", "N05AX12", "Abik OD", "艾比克口崩錠10毫克", "AripiprazoleOralDispersible", "10mg (Tab)", "O", "29.60", "1.2.2.2.", "https://info.nhi.gov.tw/api/INAE3000/INAE3000S01/getPDF?DurgFileName=1.2.2.2._20230701.pdf"]
      ];

      const ws = XLSX.utils.aoa_to_sheet(templateData);
      ws['!cols'] = [{{ wch: 14 }}, {{ wch: 16 }}, {{ wch: 12 }}, {{ wch: 35 }}, {{ wch: 25 }}, {{ wch: 30 }}, {{ wch: 18 }}, {{ wch: 10 }}, {{ wch: 14 }}, {{ wch: 20 }}, {{ wch: 40 }}];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "藥品代碼匯入標準範本");
      XLSX.writeFile(wb, "藥品代碼匯入標準範本.xlsx");
      showToast("已下載 Excel 標準匯入範本", "success");
    }}

    function exportToExcel() {{
      if (drugs.length === 0) {{
        alert("目前清單中沒有資料可供匯出！");
        return;
      }}

      const exportData = drugs.map((d, i) => ({{
        "序號": i + 1,
        "院內代碼": d.hospitalCode || "",
        "健保代碼": d.nhiCode || "自費/未收載",
        "ATC碼": d.atc || "",
        "英文商品名": d.brandName || "",
        "健保中文品名": d.chineseName || "",
        "學名/主成分": d.genericName || "",
        "規格劑量/包裝": d.strength || "",
        "劑型": d.dosageForm || "",
        "健保支付價": d.price || "",
        "健保給付規定章節": d.ruleSection || "無特殊章節",
        "給付規定PDF連結": d.ruleLink || "",
        "健保/食藥署查詢連結": d.nhiLink || ""
      }}));

      const ws = XLSX.utils.json_to_sheet(exportData);
      ws['!cols'] = [{{ wch: 8 }}, {{ wch: 14 }}, {{ wch: 16 }}, {{ wch: 12 }}, {{ wch: 35 }}, {{ wch: 25 }}, {{ wch: 30 }}, {{ wch: 18 }}, {{ wch: 10 }}, {{ wch: 14 }}, {{ wch: 20 }}, {{ wch: 40 }}, {{ wch: 40 }}];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "醫院藥品主檔(含健保價)");
      const today = new Date().toISOString().slice(0, 10).replace(/-/g, "");
      XLSX.writeFile(wb, `醫院藥品主檔與健保價格清單_${{today}}.xlsx`);
      showToast(`已成功匯出 ${{drugs.length}} 筆藥品資料！`, "success");
    }}

    function setupDragAndDrop() {{
      const dropZone = document.getElementById("dropZone");
      const overlay = document.getElementById("dragOverlay");

      window.addEventListener("dragenter", (e) => {{
        e.preventDefault();
        overlay.classList.remove("hidden");
      }});

      overlay.addEventListener("dragover", (e) => {{
        e.preventDefault();
      }});

      overlay.addEventListener("dragleave", (e) => {{
        e.preventDefault();
        overlay.classList.add("hidden");
      }});

      overlay.addEventListener("drop", (e) => {{
        e.preventDefault();
        overlay.classList.add("hidden");
        const files = e.dataTransfer.files;
        if (files && files.length > 0) {{
          processExcelFile(files[0]);
        }}
      }});
    }}

    function handleExcelUpload(e) {{
      const file = e.target.files[0];
      if (!file) return;
      processExcelFile(file);
      e.target.value = "";
    }}

    function processExcelFile(file) {{
      const reader = new FileReader();
      reader.onload = (evt) => {{
        try {{
          const data = new Uint8Array(evt.target.result);
          const workbook = XLSX.read(data, {{ type: "array" }});
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const rawRows = XLSX.utils.sheet_to_json(worksheet, {{ header: 1 }});

          if (!rawRows || rawRows.length < 2) {{
            alert("上傳的檔案無有效資料列！");
            return;
          }}

          const headers = rawRows[0].map(h => (h || "").toString().trim().toLowerCase());
          const getColIdx = (keywords) => headers.findIndex(h => keywords.some(k => h.includes(k)));

          const hCodeIdx = getColIdx(["院內代碼", "院內碼", "藥品代碼", "hospital", "院內藥碼"]);
          const nhiCodeIdx = getColIdx(["健保代碼", "健保碼", "健保", "nhi"]);
          const atcIdx = getColIdx(["atc", "atc碼", "atc代碼"]);
          const brandIdx = getColIdx(["英文商品名", "商品名", "英文名", "品名", "brand", "trade"]);
          const chineseIdx = getColIdx(["中文品名", "中文名", "中文", "chinese"]);
          const genericIdx = getColIdx(["學名", "成分", "主成分", "generic"]);
          const strengthIdx = getColIdx(["規格", "劑量", "含量", "strength", "spec"]);
          const unitIdx = getColIdx(["單位", "包裝", "unit"]);
          const dosageIdx = getColIdx(["劑型", "form", "dosage"]);
          const priceIdx = getColIdx(["支付價", "健保價", "價格", "price"]);
          const ruleIdx = getColIdx(["給付規定", "章節", "rule"]);
          const ruleLinkIdx = getColIdx(["給付規定pdf", "規定連結", "rulelink"]);

          const parsedItems = [];
          for (let i = 1; i < rawRows.length; i++) {{
            const row = rawRows[i];
            if (!row || row.length === 0) continue;

            const hospitalCode = (hCodeIdx !== -1 && row[hCodeIdx] ? row[hCodeIdx].toString().trim() : "").toUpperCase();
            const nhiCode = (nhiCodeIdx !== -1 && row[nhiCodeIdx] ? row[nhiCodeIdx].toString().trim() : "").toUpperCase();
            const atc = (atcIdx !== -1 && row[atcIdx] ? row[atcIdx].toString().trim() : "").toUpperCase();
            const brandName = brandIdx !== -1 && row[brandIdx] ? row[brandIdx].toString().trim() : "";
            const chineseName = chineseIdx !== -1 && row[chineseIdx] ? row[chineseIdx].toString().trim() : "";
            const genericName = genericIdx !== -1 && row[genericIdx] ? row[genericIdx].toString().trim() : "";
            const specVal = strengthIdx !== -1 && row[strengthIdx] ? row[strengthIdx].toString().trim() : "";
            const unitVal = unitIdx !== -1 && row[unitIdx] ? row[unitIdx].toString().trim() : "";
            const strength = (specVal && unitVal) ? `${{specVal}} (${{unitVal}})` : (specVal || unitVal);
            const dosageForm = dosageIdx !== -1 && row[dosageIdx] ? row[dosageIdx].toString().trim() : "";
            const price = priceIdx !== -1 && row[priceIdx] ? row[priceIdx].toString().trim() : "";
            const ruleSection = ruleIdx !== -1 && row[ruleIdx] ? row[ruleIdx].toString().trim() : "";
            const ruleLink = ruleLinkIdx !== -1 && row[ruleLinkIdx] ? row[ruleLinkIdx].toString().trim() : "";

            if (!hospitalCode && !brandName && !genericName) continue;

            parsedItems.push({{
              id: "d_" + Date.now() + "_" + i,
              hospitalCode: hospitalCode || `TEMP_${{i}}`,
              nhiCode,
              atc,
              brandName: brandName || "未填寫品名",
              chineseName,
              genericName: genericName || brandName,
              strength,
              dosageForm,
              price,
              ruleSection,
              ruleLink,
              nhiLink: "https://info.nhi.gov.tw/IODE0000/IODE0000S06"
            }});
          }}

          if (parsedItems.length === 0) {{
            alert("未能成功解析出有效藥品項目，請確認 Excel 表頭名稱。");
            return;
          }}

          drugs = [...parsedItems];
          saveDataToStorage();
          renderAll();
          showToast(`成功匯入 ${{parsedItems.length}} 筆藥品資料！`, "success");

        }} catch (err) {{
          console.error(err);
          alert("檔案讀取失敗，請確認檔案格式是否為標準 Excel 檔案。");
        }}
      }};
      reader.readAsArrayBuffer(file);
    }}

    function showToast(msg, type = "success") {{
      const toast = document.getElementById("toast");
      const toastMsg = document.getElementById("toastMsg");
      const toastIcon = document.getElementById("toastIcon");

      toastMsg.innerText = msg;
      if (type === "success") {{
        toastIcon.setAttribute("data-lucide", "check-circle");
        toastIcon.className = "w-4 h-4 text-emerald-400";
      }} else if (type === "info") {{
        toastIcon.setAttribute("data-lucide", "info");
        toastIcon.className = "w-4 h-4 text-sky-400";
      }}
      lucide.createIcons();

      toast.classList.remove("translate-y-20", "opacity-0");
      setTimeout(() => {{
        toast.classList.add("translate-y-20", "opacity-0");
      }}, 3000);
    }}

    function escapeHtml(str) {{
      if (!str) return "";
      return String(str)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
    }}
  </script>
</body>
</html>
"""
    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(html_content)
    print("SUCCESS: index.html compiled with full NHI Codes, Chinese names, Prices and Rule PDFs!")

if __name__ == '__main__':
    build()
