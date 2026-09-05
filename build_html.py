import openpyxl
import json

def build():
    # 1. Parse Excel
    wb = openpyxl.load_workbook('115年第一季院內藥品清單.xlsx', data_only=True)
    ws = wb.active
    drugs = []
    for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=1):
        hcode = str(r[1]).strip() if r[1] is not None else ''
        bname = str(r[2]).strip() if r[2] is not None else ''
        gname = str(r[3]).strip() if r[3] is not None else ''
        form = str(r[4]).strip() if r[4] is not None else ''
        spec = str(r[5]).strip() if r[5] is not None else ''
        unit = str(r[6]).strip() if r[6] is not None else ''
        atc = str(r[7]).strip() if r[7] is not None else ''
        if not hcode and not bname:
            continue
        full_spec = f"{spec} ({unit})" if spec and unit else (spec or unit)
        drugs.append({
            'id': f'd_115_{i}',
            'hospitalCode': hcode,
            'nhiCode': '',
            'atc': atc,
            'brandName': bname,
            'chineseName': '',
            'genericName': gname,
            'strength': full_spec,
            'dosageForm': form
        })
    
    json_data = json.dumps(drugs, ensure_ascii=False)
    
    # 2. HTML Template
    part1 = """<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>醫院藥品主檔與代碼對照管理系統</title>
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
    tailwind.config = {
      theme: {
        extend: {
          fontFamily: {
            sans: ['"Noto Sans TC"', 'Inter', 'system-ui', 'sans-serif'],
          }
        }
      }
    }
  </script>
  <style>
    body { background-color: #f8fafc; }
    .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px; }
    .custom-scrollbar::-webkit-scrollbar-track { background: #f1f5f9; }
    .custom-scrollbar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 3px; }
    .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
  </style>
</head>
<body class="text-slate-800 antialiased min-h-screen flex flex-col" id="dropZone">

  <!-- Drag and Drop Overlay -->
  <div id="dragOverlay" class="fixed inset-0 bg-teal-900/70 z-50 flex flex-col items-center justify-center text-white pointer-events-none hidden backdrop-blur-xs transition-all">
    <div class="p-8 rounded-3xl border-4 border-dashed border-teal-300 bg-teal-800/80 flex flex-col items-center space-y-4 max-w-md text-center shadow-2xl">
      <i data-lucide="file-up" class="w-16 h-16 text-teal-300 animate-bounce"></i>
      <h3 class="text-xl font-bold">釋放滑鼠以匯入 Excel 藥品清單</h3>
      <p class="text-sm text-teal-100">支援 .xlsx, .xls, .csv 檔案，自動解析代碼、品名與規格</p>
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
              <h1 class="text-lg font-bold text-slate-900 tracking-tight">醫院藥品主檔與代碼對照管理系統</h1>
              <span class="px-2 py-0.5 text-xs font-semibold rounded-full bg-teal-100 text-teal-800 border border-teal-200">v1.2 Pro</span>
            </div>
            <p class="text-xs text-slate-500">院內碼 • 健保碼 • ATC 碼 • 成分規格對照 • 格式校驗 • Excel 雙向匯出入</p>
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
            匯出 Excel
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
          <p class="text-xs font-medium text-slate-500">待補健保碼 / 異常</p>
          <p id="kpiNhiAnomalyCount" class="text-2xl font-bold text-amber-600 mt-1">0</p>
        </div>
        <div class="w-11 h-11 rounded-lg bg-amber-50 text-amber-600 flex items-center justify-center">
          <i data-lucide="alert-triangle" class="w-6 h-6"></i>
        </div>
      </div>

      <div class="bg-white rounded-xl p-4 border border-slate-200 shadow-xs flex items-center justify-between">
        <div>
          <p class="text-xs font-medium text-slate-500">重複院內代碼</p>
          <p id="kpiDuplicateCount" class="text-2xl font-bold text-rose-600 mt-1">0</p>
        </div>
        <div class="w-11 h-11 rounded-lg bg-rose-50 text-rose-600 flex items-center justify-center">
          <i data-lucide="copy" class="w-6 h-6"></i>
        </div>
      </div>

      <div class="bg-white rounded-xl p-4 border border-slate-200 shadow-xs flex items-center justify-between">
        <div>
          <p class="text-xs font-medium text-slate-500">具備 ATC 碼品項</p>
          <p id="kpiAtcCount" class="text-2xl font-bold text-indigo-600 mt-1">0</p>
        </div>
        <div class="w-11 h-11 rounded-lg bg-indigo-50 text-indigo-600 flex items-center justify-center">
          <i data-lucide="shield-check" class="w-6 h-6"></i>
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
            placeholder="搜尋商品名、中文名、成分學名、院內碼、健保碼、ATC 碼..." 
            oninput="handleSearch()"
            class="w-full pl-9 pr-4 py-2 text-sm bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 focus:bg-white transition"
          >
        </div>

        <div class="flex flex-wrap items-center gap-2">
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

          <select id="statusFilter" onchange="handleSearch()" class="text-xs bg-slate-50 border border-slate-200 rounded-lg px-2.5 py-2 text-slate-700 focus:ring-2 focus:ring-teal-500">
            <option value="all">所有狀態 (All Status)</option>
            <option value="valid">檢核正常 (Valid Only)</option>
            <option value="nhi_error">健保碼未填/異常 (NHI Missing)</option>
            <option value="duplicate">院內碼重複 (Duplicate)</option>
            <option value="incomplete">缺漏主成分 (Incomplete)</option>
          </select>

          <button onclick="openAddDrugModal()" class="inline-flex items-center px-3.5 py-2 text-xs font-semibold rounded-lg text-white bg-teal-600 hover:bg-teal-700 transition shadow-xs">
            <i data-lucide="plus-circle" class="w-4 h-4 mr-1.5"></i>
            新增單筆
          </button>
        </div>
      </div>

      <div class="flex flex-wrap items-center justify-between text-xs text-slate-500 pt-2 border-t border-slate-100 gap-2">
        <div class="flex items-center space-x-3">
          <span class="inline-flex items-center"><span class="w-2 h-2 rounded-full bg-emerald-500 mr-1.5"></span>綠燈：正常</span>
          <span class="inline-flex items-center"><span class="w-2 h-2 rounded-full bg-amber-500 mr-1.5"></span>黃燈：待補健保碼</span>
          <span class="inline-flex items-center"><span class="w-2 h-2 rounded-full bg-rose-500 mr-1.5"></span>紅燈：院內碼重複</span>
          <span class="text-slate-400 font-normal ml-2 hidden lg:inline">💡 提示：可直接將 Excel 檔案拖曳進視窗匯入</span>
        </div>
        <div class="flex flex-wrap items-center space-x-2">
          <button onclick="load115HospitalPreset()" class="text-teal-700 font-semibold hover:underline bg-teal-50 hover:bg-teal-100 px-2 py-1 rounded-md border border-teal-200 transition flex items-center space-x-1">
            <i data-lucide="refresh-cw" class="w-3.5 h-3.5"></i>
            <span>重設為 115 年院內藥品清單 (761 筆)</span>
          </button>
          <span>•</span>
          <button onclick="resetToSampleData()" class="text-slate-600 hover:underline">10 筆示範資料</button>
          <span>•</span>
          <button onclick="clearAllDataConfirm()" class="text-rose-600 hover:underline">清空資料</button>
        </div>
      </div>
    </div>

    <!-- Drug Master Data Table -->
    <div class="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
      <div class="overflow-x-auto custom-scrollbar max-h-[620px]">
        <table class="w-full text-left border-collapse">
          <thead class="bg-slate-50 border-b border-slate-200 sticky top-0 z-10 text-xs font-semibold text-slate-600">
            <tr>
              <th class="py-3 px-4 w-12 text-center">序號</th>
              <th class="py-3 px-4 w-28 cursor-pointer hover:bg-slate-100 transition" onclick="toggleSort('hospitalCode')">
                <div class="flex items-center space-x-1">
                  <span>院內代碼</span>
                  <i data-lucide="arrow-up-down" class="w-3.5 h-3.5 text-slate-400"></i>
                </div>
              </th>
              <th class="py-3 px-4 w-32 cursor-pointer hover:bg-slate-100 transition" onclick="toggleSort('nhiCode')">
                <div class="flex items-center space-x-1">
                  <span>健保代碼 (10碼)</span>
                  <i data-lucide="arrow-up-down" class="w-3.5 h-3.5 text-slate-400"></i>
                </div>
              </th>
              <th class="py-3 px-4 w-24 cursor-pointer hover:bg-slate-100 transition" onclick="toggleSort('atc')">
                <div class="flex items-center space-x-1">
                  <span>ATC 碼</span>
                  <i data-lucide="arrow-up-down" class="w-3.5 h-3.5 text-slate-400"></i>
                </div>
              </th>
              <th class="py-3 px-4 min-w-[220px] cursor-pointer hover:bg-slate-100 transition" onclick="toggleSort('brandName')">
                <div class="flex items-center space-x-1">
                  <span>英文商品名</span>
                  <i data-lucide="arrow-up-down" class="w-3.5 h-3.5 text-slate-400"></i>
                </div>
              </th>
              <th class="py-3 px-4 min-w-[130px]">中文品名</th>
              <th class="py-3 px-4 min-w-[180px]">學名 / 主成分</th>
              <th class="py-3 px-4 w-36">規格劑量 (包裝單位)</th>
              <th class="py-3 px-4 w-20">劑型</th>
              <th class="py-3 px-4 w-24 text-center">狀態檢核</th>
              <th class="py-3 px-4 w-20 text-center">操作</th>
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
      <p>醫院藥品主檔與代碼管理系統 • 專為臨床藥師與藥局資訊管理設計</p>
      <p class="text-slate-400">本系統純前端本機運作，資料安全保密不外傳 • 支援標準健保 10 碼檢核與 Excel 批次處理</p>
    </div>
  </footer>

  <!-- Modal: Add / Edit Drug -->
  <div id="drugModal" class="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-xs flex items-center justify-center hidden p-4">
    <div class="bg-white rounded-2xl max-w-xl w-full shadow-2xl border border-slate-100 overflow-hidden transform transition-all">
      <div class="px-6 py-4 bg-gradient-to-r from-teal-600 to-emerald-600 text-white flex items-center justify-between">
        <div class="flex items-center space-x-2">
          <i data-lucide="edit-3" class="w-5 h-5"></i>
          <h3 id="modalTitle" class="text-base font-bold">新增藥品主檔</h3>
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
            <input type="text" id="formHospitalCode" required placeholder="例：ONORV5" class="w-full px-3 py-2 border border-slate-300 rounded-lg uppercase tracking-wider focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
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
            <label class="block font-semibold text-slate-700 mb-1">ATC 碼</label>
            <input type="text" id="formAtc" placeholder="例：C08CA01" class="w-full px-3 py-2 border border-slate-300 rounded-lg uppercase tracking-wider font-mono focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
          </div>
        </div>

        <div>
          <label class="block font-semibold text-slate-700 mb-1">英文商品名 (Brand Name) <span class="text-rose-500">*</span></label>
          <input type="text" id="formBrandName" required placeholder="例：Norvasc Tablets 5mg" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
        </div>

        <div>
          <label class="block font-semibold text-slate-700 mb-1">中文品名 (Chinese Name)</label>
          <input type="text" id="formChineseName" placeholder="例：脈優錠 5 毫克" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
        </div>

        <div>
          <label class="block font-semibold text-slate-700 mb-1">學名 / 主成分 (Generic Name) <span class="text-rose-500">*</span></label>
          <input type="text" id="formGenericName" required placeholder="例：Amlodipine besylate" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <div>
            <label class="block font-semibold text-slate-700 mb-1">規格劑量/包裝 (Strength/Unit)</label>
            <input type="text" id="formStrength" placeholder="例：5mg/tab, 1g/vial (Btl)" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
          </div>

          <div>
            <label class="block font-semibold text-slate-700 mb-1">劑型 (Dosage Form)</label>
            <input type="text" id="formDosageForm" list="dosageFormOptions" placeholder="例：O, I, E, 錠劑、膠囊劑" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
            <datalist id="dosageFormOptions">
              <option value="O (口服 Oral)">
              <option value="I (注射 Injection)">
              <option value="E (外用 External)">
              <option value="錠劑 (Tablet)">
              <option value="膜衣錠 (F.C. Tablet)">
              <option value="膠囊劑 (Capsule)">
              <option value="注射劑 (Injection)">
              <option value="吸入劑 (Inhaler)">
              <option value="外用乳膏 (Cream)">
              <option value="點眼液 (Eye Drops)">
            </datalist>
          </div>
        </div>

        <div class="pt-4 border-t border-slate-100 flex items-center justify-end space-x-3">
          <button type="button" onclick="closeDrugModal()" class="px-4 py-2 rounded-lg border border-slate-300 text-slate-600 hover:bg-slate-50 transition">取消</button>
          <button type="submit" class="px-5 py-2 rounded-lg bg-teal-600 hover:bg-teal-700 text-white font-medium shadow-sm transition">儲存資料</button>
        </div>
      </form>
    </div>
  </div>

  <!-- Modal: Conflict -->
  <div id="conflictModal" class="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-xs flex items-center justify-center hidden p-4">
    <div class="bg-white rounded-2xl max-w-2xl w-full shadow-2xl border border-slate-100 overflow-hidden">
      <div class="px-6 py-4 bg-amber-500 text-white flex items-center justify-between">
        <div class="flex items-center space-x-2">
          <i data-lucide="alert-circle" class="w-5 h-5"></i>
          <h3 class="text-base font-bold">匯入資料代碼重複衝突提示</h3>
        </div>
        <button onclick="closeConflictModal()" class="text-white/80 hover:text-white transition">
          <i data-lucide="x" class="w-5 h-5"></i>
        </button>
      </div>
      <div class="p-6 space-y-4 text-xs">
        <p class="text-slate-600">
          系統在您上傳的 Excel 檔案中偵測到 <span id="conflictCount" class="font-bold text-rose-600 text-sm">0</span> 筆與現有清單中「<span class="font-semibold text-slate-800">院內代碼</span>」相同的藥品項目。
        </p>

        <div class="border border-slate-200 rounded-lg max-h-48 overflow-y-auto custom-scrollbar">
          <table class="w-full text-left text-xs divide-y divide-slate-100">
            <thead class="bg-slate-50 text-slate-600 sticky top-0">
              <tr>
                <th class="py-2 px-3">院內代碼</th>
                <th class="py-2 px-3">現有商品名</th>
                <th class="py-2 px-3">匯入商品名</th>
              </tr>
            </thead>
            <tbody id="conflictTableBody" class="divide-y divide-slate-100"></tbody>
          </table>
        </div>

        <p class="text-slate-500">請選擇本次衝突項目的處理方式：</p>

        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-2">
          <button onclick="resolveConflicts('overwrite')" class="p-3 border-2 border-teal-500 bg-teal-50/50 hover:bg-teal-50 rounded-xl text-left transition">
            <div class="font-bold text-teal-800 flex items-center">
              <i data-lucide="refresh-cw" class="w-4 h-4 mr-1.5"></i>
              全部覆蓋更新 (Overwrite)
            </div>
            <p class="text-slate-500 text-[11px] mt-1">以 Excel 新檔案中的資料完全更新重複的代碼項目。</p>
          </button>

          <button onclick="resolveConflicts('skip')" class="p-3 border-2 border-slate-200 hover:border-slate-300 rounded-xl text-left transition">
            <div class="font-bold text-slate-700 flex items-center">
              <i data-lucide="skip-forward" class="w-4 h-4 mr-1.5"></i>
              略過重複項目 (Skip)
            </div>
            <p class="text-slate-500 text-[11px] mt-1">保留現有藥品資料不變，僅匯入未重複的新代碼。</p>
          </button>
        </div>
      </div>
    </div>
  </div>

  <!-- Toast -->
  <div id="toast" class="fixed bottom-6 right-6 z-50 transform translate-y-20 opacity-0 transition-all duration-300 flex items-center space-x-2 px-4 py-3 rounded-xl shadow-lg text-xs font-medium text-white bg-slate-900">
    <i id="toastIcon" data-lucide="info" class="w-4 h-4 text-teal-400"></i>
    <span id="toastMsg">操作成功</span>
  </div>

  <script>
    // Embedded 761 Clinical Drugs Dataset
    const HOSPITAL_115_PRESET = """

    part2 = """;

    // Initial 10 Clinical Sample Data
    const SAMPLE_DRUGS = [
      { id: "d1", hospitalCode: "ONORV5", nhiCode: "BC19248100", atc: "C08CA01", brandName: "Norvasc Tablets 5mg", chineseName: "脈優錠 5 毫克", genericName: "Amlodipine besylate", strength: "5mg/tab (Box)", dosageForm: "O" },
      { id: "d2", hospitalCode: "OGLUC5", nhiCode: "AC34125100", atc: "A10BA02", brandName: "Glucophage Tablets 500mg", chineseName: "庫魯化錠 500 毫克", genericName: "Metformin hydrochloride", strength: "500mg/tab (Box)", dosageForm: "O" },
      { id: "d3", hospitalCode: "OKEF50", nhiCode: "AC08249100", atc: "J01DB01", brandName: "Keflex Capsules 500mg", chineseName: "賜福力欣膠囊 500 毫克", genericName: "Cephalexin", strength: "500mg/cap (Box)", dosageForm: "O" },
      { id: "d4", hospitalCode: "OLIP20", nhiCode: "BC22436100", atc: "C10AA05", brandName: "Lipitor F.C. Tablets 20mg", chineseName: "立普妥膜衣錠 20 毫克", genericName: "Atorvastatin calcium", strength: "20mg/tab (Box)", dosageForm: "O" },
      { id: "d5", hospitalCode: "ONEX40", nhiCode: "BC23164100", atc: "A02BC05", brandName: "Nexium Gastro-resistant Tablets 40mg", chineseName: "耐適恩錠 40 毫克", genericName: "Esomeprazole magnesium", strength: "40mg/tab (Box)", dosageForm: "O" },
      { id: "d6", hospitalCode: "OCRA50", nhiCode: "BC23812100", atc: "J01MA12", brandName: "Cravit F.C. Tablets 500mg", chineseName: "可樂必妥膜衣錠 500 毫克", genericName: "Levofloxacin", strength: "500mg/tab (Box)", dosageForm: "O" },
      { id: "d7", hospitalCode: "OASP10", nhiCode: "AC17281100", atc: "B01AC06", brandName: "Bokey Enteric Microencapsulated Cap 100mg", chineseName: "伯基膠囊 100 毫克", genericName: "Aspirin", strength: "100mg/cap (Box)", dosageForm: "O" },
      { id: "d8", hospitalCode: "OVENT2", nhiCode: "AC24322100", atc: "R03AC02", brandName: "Ventolin Inhaler 100mcg/dose", chineseName: "備勞喘吸入劑 100 微克/劑量", genericName: "Salbutamol sulfate", strength: "100mcg/puff, 200 puffs (Btl)", dosageForm: "E" },
      { id: "d9", hospitalCode: "IROCE1", nhiCode: "BC17618200", atc: "J01DD04", brandName: "Rocephin for Injection 1g", chineseName: "羅氏芬注射劑 1 公克", genericName: "Ceftriaxone sodium", strength: "1g/vial (Vial)", dosageForm: "I" },
      { id: "d10", hospitalCode: "EKAL01", nhiCode: "AC49213421", atc: "V03AE01", brandName: "Kalimate Powder 5g/pack", chineseName: "降鉀妥散 5 公克/包", genericName: "Calcium polystyrene sulfonate", strength: "5g/pack (Pack)", dosageForm: "E" }
    ];

    const STORAGE_KEY = "hospital_drug_master_db_v3";
    let drugs = [];
    let currentPage = 1;
    let pageSize = 15;
    let sortColumn = "hospitalCode";
    let sortAsc = true;
    let pendingImportData = null;

    window.addEventListener("DOMContentLoaded", () => {
      loadDataFromStorage();
      renderAll();
      lucide.createIcons();
      setupDragAndDrop();
    });

    function loadDataFromStorage() {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        try {
          drugs = JSON.parse(saved);
        } catch(e) {
          drugs = JSON.parse(JSON.stringify(HOSPITAL_115_PRESET));
        }
      } else {
        drugs = JSON.parse(JSON.stringify(HOSPITAL_115_PRESET));
        saveDataToStorage();
      }
    }

    function saveDataToStorage() {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(drugs));
    }

    function resetToSampleData() {
      if (confirm("確定要切換為 10 筆精選臨床示範資料嗎？")) {
        drugs = JSON.parse(JSON.stringify(SAMPLE_DRUGS));
        saveDataToStorage();
        renderAll();
        showToast("已載入 10 筆示範資料", "success");
      }
    }

    function load115HospitalPreset() {
      if (confirm("確定要重設載入「115 年第一季院內藥品清單 (761 筆)」嗎？")) {
        drugs = JSON.parse(JSON.stringify(HOSPITAL_115_PRESET));
        saveDataToStorage();
        renderAll();
        showToast(`已成功載入 115 年院內清單，共 ${drugs.length} 筆！`, "success");
      }
    }

    function clearAllDataConfirm() {
      if (confirm("警告：確定要清空所有藥品資料嗎？請確保您已匯出 Excel 備份。")) {
        drugs = [];
        saveDataToStorage();
        renderAll();
        showToast("已清空所有藥品資料", "info");
      }
    }

    function validateDrug(drug, allDrugs) {
      const issues = [];
      const nhi = (drug.nhiCode || "").trim();
      if (!nhi) {
        issues.push({ type: "nhi_empty", level: "amber", text: "待補健保碼" });
      } else if (nhi.length !== 10) {
        issues.push({ type: "nhi_length", level: "amber", text: `健保碼非10碼 (${nhi.length}碼)` });
      }

      const hCode = (drug.hospitalCode || "").trim().toUpperCase();
      if (!hCode) {
        issues.push({ type: "hcode_empty", level: "rose", text: "院內碼未填" });
      } else {
        const duplicates = allDrugs.filter(d => (d.hospitalCode || "").trim().toUpperCase() === hCode);
        if (duplicates.length > 1) {
          issues.push({ type: "duplicate", level: "rose", text: "院內碼重複" });
        }
      }

      if (!drug.brandName || !drug.genericName) {
        issues.push({ type: "incomplete", level: "indigo", text: "缺商品名或成分" });
      }
      return issues;
    }

    function renderAll() {
      updateKpis();
      renderTable();
      lucide.createIcons();
    }

    function updateKpis() {
      const total = drugs.length;
      let nhiAnomaly = 0;
      let duplicatesCount = 0;
      let atcCount = 0;

      const codeMap = {};
      drugs.forEach(d => {
        const c = (d.hospitalCode || "").trim().toUpperCase();
        if (c) codeMap[c] = (codeMap[c] || 0) + 1;
        if (d.atc && d.atc.trim()) atcCount++;
      });

      drugs.forEach(d => {
        const issues = validateDrug(d, drugs);
        if (issues.some(i => i.type.startsWith("nhi_"))) nhiAnomaly++;
      });

      Object.values(codeMap).forEach(cnt => {
        if (cnt > 1) duplicatesCount += cnt;
      });

      document.getElementById("kpiTotalCount").innerText = total;
      document.getElementById("kpiNhiAnomalyCount").innerText = nhiAnomaly;
      document.getElementById("kpiDuplicateCount").innerText = duplicatesCount;
      document.getElementById("kpiAtcCount").innerText = atcCount;
    }

    function getFilteredDrugs() {
      const q = document.getElementById("searchInput").value.trim().toLowerCase();
      const dosageFilter = document.getElementById("dosageFilter").value;
      const statusFilter = document.getElementById("statusFilter").value;

      return drugs.filter(drug => {
        const matchQuery = !q || 
          (drug.hospitalCode && drug.hospitalCode.toLowerCase().includes(q)) ||
          (drug.nhiCode && drug.nhiCode.toLowerCase().includes(q)) ||
          (drug.atc && drug.atc.toLowerCase().includes(q)) ||
          (drug.brandName && drug.brandName.toLowerCase().includes(q)) ||
          (drug.chineseName && drug.chineseName.toLowerCase().includes(q)) ||
          (drug.genericName && drug.genericName.toLowerCase().includes(q)) ||
          (drug.strength && drug.strength.toLowerCase().includes(q)) ||
          (drug.dosageForm && drug.dosageForm.toLowerCase().includes(q));

        if (!matchQuery) return false;

        if (dosageFilter && !(drug.dosageForm || "").includes(dosageFilter)) {
          return false;
        }

        if (statusFilter !== "all") {
          const issues = validateDrug(drug, drugs);
          if (statusFilter === "valid" && issues.length === 0) return true;
          if (statusFilter === "valid" && issues.length > 0) return false;
          if (statusFilter === "nhi_error" && !issues.some(i => i.type.startsWith("nhi_"))) return false;
          if (statusFilter === "duplicate" && !issues.some(i => i.type === "duplicate")) return false;
          if (statusFilter === "incomplete" && !issues.some(i => i.type === "incomplete")) return false;
        }

        return true;
      });
    }

    function toggleSort(col) {
      if (sortColumn === col) {
        sortAsc = !sortAsc;
      } else {
        sortColumn = col;
        sortAsc = true;
      }
      renderTable();
      lucide.createIcons();
    }

    function renderTable() {
      const filtered = getFilteredDrugs();

      filtered.sort((a, b) => {
        let valA = (a[sortColumn] || "").toString().toLowerCase();
        let valB = (b[sortColumn] || "").toString().toLowerCase();
        if (valA < valB) return sortAsc ? -1 : 1;
        if (valA > valB) return sortAsc ? 1 : -1;
        return 0;
      });

      const totalRecords = filtered.length;
      const totalPages = Math.ceil(totalRecords / pageSize) || 1;
      if (currentPage > totalPages) currentPage = totalPages;
      if (currentPage < 1) currentPage = 1;

      const startIndex = (currentPage - 1) * pageSize;
      const paginated = filtered.slice(startIndex, startIndex + pageSize);

      const tbody = document.getElementById("drugTableBody");
      if (paginated.length === 0) {
        tbody.innerHTML = `
          <tr>
            <td colspan="11" class="py-12 text-center text-slate-400">
              <i data-lucide="inbox" class="w-10 h-10 mx-auto mb-2 text-slate-300"></i>
              <p class="text-sm font-medium">查無符合條件之藥品資料</p>
              <p class="text-xs text-slate-400 mt-1">請嘗試調整搜尋關鍵字或點擊「新增單筆」建立主檔</p>
            </td>
          </tr>
        `;
      } else {
        tbody.innerHTML = paginated.map((drug, idx) => {
          const rowNum = startIndex + idx + 1;
          const issues = validateDrug(drug, drugs);
          
          let statusBadge = `<span class="inline-flex items-center px-2 py-0.5 rounded-full text-[11px] font-medium bg-emerald-100 text-emerald-800"><span class="w-1.5 h-1.5 rounded-full bg-emerald-500 mr-1"></span>正常</span>`;
          if (issues.length > 0) {
            statusBadge = issues.map(iss => {
              const bg = iss.level === 'rose' ? 'bg-rose-100 text-rose-800' : (iss.level === 'amber' ? 'bg-amber-100 text-amber-800' : 'bg-indigo-100 text-indigo-800');
              const dot = iss.level === 'rose' ? 'bg-rose-500' : (iss.level === 'amber' ? 'bg-amber-500' : 'bg-indigo-500');
              return `<span class="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-medium ${bg} mr-1 mb-1"><span class="w-1.5 h-1.5 rounded-full ${dot} mr-1"></span>${iss.text}</span>`;
            }).join("");
          }

          const hasNhi = !!(drug.nhiCode || "").trim();
          const isNhiBad = hasNhi && (drug.nhiCode || "").trim().length !== 10;
          let nhiDisplay = `<span class="text-slate-400 font-mono italic">待補</span>`;
          if (hasNhi) {
            if (isNhiBad) {
              nhiDisplay = `<span class="text-amber-700 bg-amber-50 border border-amber-200 px-1.5 py-0.5 rounded font-mono font-semibold">${escapeHtml(drug.nhiCode)}</span>`;
            } else {
              nhiDisplay = `<span class="font-mono text-slate-800 font-medium">${escapeHtml(drug.nhiCode)}</span>`;
            }
          }

          return `
            <tr class="hover:bg-slate-50/80 transition group">
              <td class="py-3 px-4 text-center font-mono text-slate-400 text-xs">${rowNum}</td>
              <td class="py-3 px-4 font-semibold text-slate-900 font-mono tracking-wide">
                <span class="bg-slate-100 text-slate-800 px-1.5 py-0.5 rounded border border-slate-200">${escapeHtml(drug.hospitalCode || "-")}</span>
              </td>
              <td class="py-3 px-4">${nhiDisplay}</td>
              <td class="py-3 px-4">
                <span class="font-mono text-teal-800 bg-teal-50 px-1.5 py-0.5 rounded text-[11px] border border-teal-100">${escapeHtml(drug.atc || "-")}</span>
              </td>
              <td class="py-3 px-4 font-medium text-slate-900">${escapeHtml(drug.brandName || "-")}</td>
              <td class="py-3 px-4 text-slate-700">${escapeHtml(drug.chineseName || "-")}</td>
              <td class="py-3 px-4 text-slate-600 italic">${escapeHtml(drug.genericName || "-")}</td>
              <td class="py-3 px-4 text-slate-600">${escapeHtml(drug.strength || "-")}</td>
              <td class="py-3 px-4 text-slate-600">
                <span class="bg-slate-100 text-slate-700 px-2 py-0.5 rounded text-[11px] font-medium">${escapeHtml(drug.dosageForm || "-")}</span>
              </td>
              <td class="py-3 px-4 text-center">${statusBadge}</td>
              <td class="py-3 px-4 text-center">
                <div class="flex items-center justify-center space-x-1">
                  <button onclick="openEditDrugModal('${drug.id}')" title="編輯藥品" class="p-1.5 text-slate-400 hover:text-teal-600 hover:bg-teal-50 rounded transition">
                    <i data-lucide="edit-2" class="w-4 h-4"></i>
                  </button>
                  <button onclick="deleteDrug('${drug.id}')" title="刪除" class="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded transition">
                    <i data-lucide="trash-2" class="w-4 h-4"></i>
                  </button>
                </div>
              </td>
            </tr>
          `;
        }).join("");
      }

      const startDisplay = totalRecords === 0 ? 0 : startIndex + 1;
      const endDisplay = Math.min(startIndex + pageSize, totalRecords);
      document.getElementById("tableRecordSummary").innerText = `顯示第 ${startDisplay} 到 ${endDisplay} 筆，共 ${totalRecords} 筆資料`;
      document.getElementById("currentPageIndicator").innerText = `${currentPage} / ${totalPages}`;
      document.getElementById("btnPrevPage").disabled = (currentPage <= 1);
      document.getElementById("btnNextPage").disabled = (currentPage >= totalPages);
    }

    function handleSearch() {
      currentPage = 1;
      renderTable();
      lucide.createIcons();
    }

    function changePageSize(val) {
      pageSize = parseInt(val);
      currentPage = 1;
      renderTable();
      lucide.createIcons();
    }

    function prevPage() {
      if (currentPage > 1) {
        currentPage--;
        renderTable();
        lucide.createIcons();
      }
    }

    function nextPage() {
      const filtered = getFilteredDrugs();
      const totalPages = Math.ceil(filtered.length / pageSize) || 1;
      if (currentPage < totalPages) {
        currentPage++;
        renderTable();
        lucide.createIcons();
      }
    }

    function openAddDrugModal() {
      document.getElementById("modalTitle").innerText = "新增藥品主檔";
      document.getElementById("formDrugId").value = "";
      document.getElementById("drugForm").reset();
      handleNhiInput("");
      document.getElementById("formHospitalCodeWarn").classList.add("hidden");
      document.getElementById("drugModal").classList.remove("hidden");
      lucide.createIcons();
      setTimeout(() => document.getElementById("formHospitalCode").focus(), 50);
    }

    function openEditDrugModal(id) {
      const drug = drugs.find(d => d.id === id);
      if (!drug) return;

      document.getElementById("modalTitle").innerText = "編輯藥品主檔";
      document.getElementById("formDrugId").value = drug.id;
      document.getElementById("formHospitalCode").value = drug.hospitalCode || "";
      document.getElementById("formNhiCode").value = drug.nhiCode || "";
      document.getElementById("formAtc").value = drug.atc || "";
      document.getElementById("formBrandName").value = drug.brandName || "";
      document.getElementById("formChineseName").value = drug.chineseName || "";
      document.getElementById("formGenericName").value = drug.genericName || "";
      document.getElementById("formStrength").value = drug.strength || "";
      document.getElementById("formDosageForm").value = drug.dosageForm || "";

      handleNhiInput(drug.nhiCode || "");
      document.getElementById("formHospitalCodeWarn").classList.add("hidden");
      document.getElementById("drugModal").classList.remove("hidden");
      lucide.createIcons();
    }

    function closeDrugModal() {
      document.getElementById("drugModal").classList.add("hidden");
    }

    function handleNhiInput(val) {
      const len = val.trim().length;
      const badge = document.getElementById("nhiLengthBadge");
      badge.innerText = `${len}/10碼`;

      if (len === 10) {
        badge.className = "text-[10px] px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-700 font-mono font-bold";
      } else if (len === 0) {
        badge.className = "text-[10px] px-1.5 py-0.5 rounded bg-slate-100 text-slate-500 font-mono";
      } else {
        badge.className = "text-[10px] px-1.5 py-0.5 rounded bg-amber-100 text-amber-700 font-mono font-bold";
      }
    }

    function saveDrugForm(e) {
      e.preventDefault();
      const id = document.getElementById("formDrugId").value;
      const hospitalCode = document.getElementById("formHospitalCode").value.trim().toUpperCase();
      const nhiCode = document.getElementById("formNhiCode").value.trim().toUpperCase();
      const atc = document.getElementById("formAtc").value.trim().toUpperCase();
      const brandName = document.getElementById("formBrandName").value.trim();
      const chineseName = document.getElementById("formChineseName").value.trim();
      const genericName = document.getElementById("formGenericName").value.trim();
      const strength = document.getElementById("formStrength").value.trim();
      const dosageForm = document.getElementById("formDosageForm").value.trim();

      const duplicate = drugs.find(d => d.hospitalCode.toUpperCase() === hospitalCode && d.id !== id);
      if (duplicate) {
        const warn = document.getElementById("formHospitalCodeWarn");
        warn.innerText = `院內代碼 ${hospitalCode} 已存在於「${duplicate.brandName}」！`;
        warn.classList.remove("hidden");
        return;
      }

      if (id) {
        const index = drugs.findIndex(d => d.id === id);
        if (index !== -1) {
          drugs[index] = { id, hospitalCode, nhiCode, atc, brandName, chineseName, genericName, strength, dosageForm };
          showToast("已更新藥品主檔資料", "success");
        }
      } else {
        const newDrug = {
          id: "d_" + Date.now(),
          hospitalCode,
          nhiCode,
          atc,
          brandName,
          chineseName,
          genericName,
          strength,
          dosageForm
        };
        drugs.unshift(newDrug);
        showToast("已成功新增藥品主檔", "success");
      }

      saveDataToStorage();
      closeDrugModal();
      renderAll();
    }

    function deleteDrug(id) {
      const drug = drugs.find(d => d.id === id);
      if (!drug) return;
      if (confirm(`確定要刪除藥品「${drug.brandName} (${drug.hospitalCode})」嗎？`)) {
        drugs = drugs.filter(d => d.id !== id);
        saveDataToStorage();
        renderAll();
        showToast("已刪除該藥品品項", "info");
      }
    }

    function downloadExcelTemplate() {
      const headers = ["院內代碼", "健保代碼", "ATC碼", "英文商品名", "中文品名", "學名/成分", "規格劑量", "劑型"];
      const templateData = [
        headers,
        ["ONORV5", "BC19248100", "C08CA01", "Norvasc Tablets 5mg", "脈優錠 5 毫克", "Amlodipine besylate", "5mg/tab", "O (口服錠劑)"],
        ["OGLUC5", "AC34125100", "A10BA02", "Glucophage Tablets 500mg", "庫魯化錠 500 毫克", "Metformin hydrochloride", "500mg/tab", "O (口服錠劑)"],
        ["OKEF50", "AC08249100", "J01DB01", "Keflex Capsules 500mg", "賜福力欣膠囊 500 毫克", "Cephalexin", "500mg/cap", "O (口服膠囊)"]
      ];

      const ws = XLSX.utils.aoa_to_sheet(templateData);
      ws['!cols'] = [{ wch: 14 }, { wch: 16 }, { wch: 12 }, { wch: 35 }, { wch: 25 }, { wch: 30 }, { wch: 18 }, { wch: 16 }];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "藥品代碼匯入標準範本");
      XLSX.writeFile(wb, "藥品代碼匯入標準範本.xlsx");
      showToast("已下載 Excel 標準匯入範本", "success");
    }

    function exportToExcel() {
      if (drugs.length === 0) {
        alert("目前清單中沒有資料可供匯出！");
        return;
      }

      const exportData = drugs.map(d => ({
        "院內代碼": d.hospitalCode || "",
        "健保代碼": d.nhiCode || "",
        "ATC碼": d.atc || "",
        "英文商品名": d.brandName || "",
        "中文品名": d.chineseName || "",
        "學名/主成分": d.genericName || "",
        "規格劑量": d.strength || "",
        "劑型": d.dosageForm || ""
      }));

      const ws = XLSX.utils.json_to_sheet(exportData);
      ws['!cols'] = [{ wch: 14 }, { wch: 16 }, { wch: 12 }, { wch: 35 }, { wch: 25 }, { wch: 30 }, { wch: 18 }, { wch: 16 }];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "醫院藥品主檔");
      const today = new Date().toISOString().slice(0, 10).replace(/-/g, "");
      XLSX.writeFile(wb, `醫院藥品主檔與代碼清單_${today}.xlsx`);
      showToast(`已成功匯出 ${drugs.length} 筆藥品資料！`, "success");
    }

    function setupDragAndDrop() {
      const dropZone = document.getElementById("dropZone");
      const overlay = document.getElementById("dragOverlay");

      window.addEventListener("dragenter", (e) => {
        e.preventDefault();
        overlay.classList.remove("hidden");
      });

      overlay.addEventListener("dragover", (e) => {
        e.preventDefault();
      });

      overlay.addEventListener("dragleave", (e) => {
        e.preventDefault();
        overlay.classList.add("hidden");
      });

      overlay.addEventListener("drop", (e) => {
        e.preventDefault();
        overlay.classList.add("hidden");
        const files = e.dataTransfer.files;
        if (files && files.length > 0) {
          processExcelFile(files[0]);
        }
      });
    }

    function handleExcelUpload(e) {
      const file = e.target.files[0];
      if (!file) return;
      processExcelFile(file);
      e.target.value = "";
    }

    function processExcelFile(file) {
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const data = new Uint8Array(evt.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

          if (!rawRows || rawRows.length < 2) {
            alert("上傳的檔案無有效資料列！");
            return;
          }

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

          const parsedItems = [];
          for (let i = 1; i < rawRows.length; i++) {
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
            const strength = (specVal && unitVal) ? `${specVal} (${unitVal})` : (specVal || unitVal);
            const dosageForm = dosageIdx !== -1 && row[dosageIdx] ? row[dosageIdx].toString().trim() : "";

            if (!hospitalCode && !brandName && !genericName) continue;

            parsedItems.push({
              id: "d_" + Date.now() + "_" + i,
              hospitalCode: hospitalCode || `TEMP_${i}`,
              nhiCode,
              atc,
              brandName: brandName || "未填寫品名",
              chineseName,
              genericName: genericName || brandName,
              strength,
              dosageForm
            });
          }

          if (parsedItems.length === 0) {
            alert("未能成功解析出有效藥品項目，請確認 Excel 表頭名稱或下載標準範本對照。");
            return;
          }

          const conflicts = [];
          parsedItems.forEach(incoming => {
            const existing = drugs.find(d => d.hospitalCode.toUpperCase() === incoming.hospitalCode.toUpperCase());
            if (existing) {
              conflicts.push({ incoming, existing });
            }
          });

          if (conflicts.length > 0) {
            pendingImportData = { parsedItems, conflicts };
            showConflictModal(conflicts);
          } else {
            drugs = [...parsedItems];
            saveDataToStorage();
            renderAll();
            showToast(`成功匯入 ${parsedItems.length} 筆藥品資料！`, "success");
          }

        } catch (err) {
          console.error(err);
          alert("檔案讀取失敗，請確認檔案格式是否為標準 Excel (.xlsx/.xls) 或 CSV 檔案。");
        }
      };
      reader.readAsArrayBuffer(file);
    }

    function showConflictModal(conflicts) {
      document.getElementById("conflictCount").innerText = conflicts.length;
      const tbody = document.getElementById("conflictTableBody");
      tbody.innerHTML = conflicts.map(c => `
        <tr>
          <td class="py-2 px-3 font-mono font-semibold text-slate-800">${escapeHtml(c.incoming.hospitalCode)}</td>
          <td class="py-2 px-3 text-slate-500 line-through">${escapeHtml(c.existing.brandName)}</td>
          <td class="py-2 px-3 font-medium text-teal-700">${escapeHtml(c.incoming.brandName)}</td>
        </tr>
      `).join("");

      document.getElementById("conflictModal").classList.remove("hidden");
      lucide.createIcons();
    }

    function closeConflictModal() {
      document.getElementById("conflictModal").classList.add("hidden");
      pendingImportData = null;
    }

    function resolveConflicts(strategy) {
      if (!pendingImportData) return;
      const { parsedItems } = pendingImportData;

      if (strategy === "overwrite") {
        parsedItems.forEach(incoming => {
          const idx = drugs.findIndex(d => d.hospitalCode.toUpperCase() === incoming.hospitalCode.toUpperCase());
          if (idx !== -1) {
            drugs[idx] = { ...incoming, id: drugs[idx].id };
          } else {
            drugs.unshift(incoming);
          }
        });
        showToast(`已覆蓋更新並匯入 ${parsedItems.length} 筆資料！`, "success");
      } else if (strategy === "skip") {
        let addedCount = 0;
        parsedItems.forEach(incoming => {
          const exists = drugs.some(d => d.hospitalCode.toUpperCase() === incoming.hospitalCode.toUpperCase());
          if (!exists) {
            drugs.unshift(incoming);
            addedCount++;
          }
        });
        showToast(`已略過重複代碼，新增 ${addedCount} 筆新藥品！`, "info");
      }

      saveDataToStorage();
      closeConflictModal();
      renderAll();
    }

    function showToast(msg, type = "success") {
      const toast = document.getElementById("toast");
      const toastMsg = document.getElementById("toastMsg");
      const toastIcon = document.getElementById("toastIcon");

      toastMsg.innerText = msg;
      if (type === "success") {
        toastIcon.setAttribute("data-lucide", "check-circle");
        toastIcon.className = "w-4 h-4 text-emerald-400";
      } else if (type === "info") {
        toastIcon.setAttribute("data-lucide", "info");
        toastIcon.className = "w-4 h-4 text-sky-400";
      }
      lucide.createIcons();

      toast.classList.remove("translate-y-20", "opacity-0");
      setTimeout(() => {
        toast.classList.add("translate-y-20", "opacity-0");
      }, 3000);
    }

    function escapeHtml(str) {
      if (!str) return "";
      return String(str)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
    }
  </script>
</body>
</html>"""

    full_html = part1 + json_data + part2
    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(full_html)
    print("SUCCESS: index.html generated with 761 embedded drugs!")

if __name__ == '__main__':
    build()
