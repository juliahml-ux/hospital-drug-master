import json
from datetime import datetime

def build():
    with open('output/115_q2_drugs_enriched_preset.json', 'r', encoding='utf-8') as f:
        drugs = json.load(f)

    json_data = json.dumps(drugs, ensure_ascii=False)
    revision_time = datetime.now().strftime("%Y-%m-%d %H:%M")

    html_content = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>衛生福利部嘉義醫院藥品臨床處方輯管理系統 (115年第二季)</title>
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
      <p class="text-sm text-teal-100">支援 115 年最新清單，自動解析代碼、品名、健保價、處方劑量與給付條件</p>
    </div>
  </div>

  <!-- Header / Navigation -->
  <header class="bg-white border-b border-slate-200 sticky top-0 z-30 shadow-xs">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
      <div class="flex items-center justify-between h-16">
        <!-- Logo & Title -->
        <div class="flex items-center space-x-3">
          <div class="w-10 h-10 rounded-xl bg-gradient-to-br from-teal-500 to-emerald-600 flex items-center justify-center text-white shadow-md shadow-teal-500/20 shrink-0">
            <i data-lucide="pill" class="w-6 h-6"></i>
          </div>
          <div>
            <div class="flex flex-wrap items-center gap-2">
              <h1 class="text-base sm:text-lg font-bold text-slate-900 tracking-tight">衛生福利部嘉義醫院藥品臨床處方輯管理系統</h1>
              <span class="px-2 py-0.5 text-xs font-semibold rounded-full bg-teal-100 text-teal-800 border border-teal-200">115 年第二季 處方輯整合版</span>
              <span class="inline-flex items-center px-2 py-0.5 text-[11px] font-medium rounded-full bg-slate-100 text-slate-600 border border-slate-200 shadow-2xs">
                <i data-lucide="clock" class="w-3 h-3 mr-1 text-teal-600"></i>
                最新修訂：{revision_time}
              </span>
            </div>
            <p class="text-xs text-slate-500 mt-0.5">738 筆院內主檔 • 健保價與給付 PDF • 臨床用法用量 • 禁忌副作用 • 懷孕/肝腎安全分級</p>
          </div>
        </div>

        <!-- Quick Top Actions -->
        <div class="flex items-center space-x-2 sm:space-x-3">
          <button onclick="openQrModal()" class="inline-flex items-center px-2.5 py-1.5 text-xs font-medium rounded-lg text-slate-700 bg-slate-100 hover:bg-slate-200 transition" title="手機掃描 QR Code">
            <i data-lucide="qr-code" class="w-4 h-4 mr-1 text-slate-600"></i>
            手機 QR Code
          </button>
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
            匯出 Excel (含臨床處方)
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
          <p class="text-xs font-medium text-slate-500">115Q2 藥品總數</p>
          <p id="kpiTotalCount" class="text-2xl font-bold text-slate-900 mt-1">738</p>
        </div>
        <div class="w-11 h-11 rounded-lg bg-teal-50 text-teal-600 flex items-center justify-center">
          <i data-lucide="database" class="w-6 h-6"></i>
        </div>
      </div>

      <div class="bg-white rounded-xl p-4 border border-slate-200 shadow-xs flex items-center justify-between">
        <div>
          <p class="text-xs font-medium text-slate-500">健保給付品項 (有健保價)</p>
          <p id="kpiNhiCoveredCount" class="text-2xl font-bold text-emerald-600 mt-1">0</p>
        </div>
        <div class="w-11 h-11 rounded-lg bg-emerald-50 text-emerald-600 flex items-center justify-center">
          <i data-lucide="check-check" class="w-6 h-6"></i>
        </div>
      </div>

      <div class="bg-white rounded-xl p-4 border border-slate-200 shadow-xs flex items-center justify-between">
        <div>
          <p class="text-xs font-medium text-slate-500">附處方集臨床資訊</p>
          <p id="kpiFormularyCount" class="text-2xl font-bold text-indigo-600 mt-1">0</p>
        </div>
        <div class="w-11 h-11 rounded-lg bg-indigo-50 text-indigo-600 flex items-center justify-center">
          <i data-lucide="book-open" class="w-6 h-6"></i>
        </div>
      </div>

      <div class="bg-white rounded-xl p-4 border border-slate-200 shadow-xs flex items-center justify-between">
        <div>
          <p class="text-xs font-medium text-slate-500">需肝/腎劑量調整 [肝][腎]</p>
          <p id="kpiAlertCount" class="text-2xl font-bold text-amber-600 mt-1">0</p>
        </div>
        <div class="w-11 h-11 rounded-lg bg-amber-50 text-amber-600 flex items-center justify-center">
          <i data-lucide="alert-triangle" class="w-6 h-6"></i>
        </div>
      </div>
    </div>

    <!-- Toolbar: Search, Filters & Actions (Sticky Top Pane) -->
    <div class="bg-white/95 backdrop-blur-md rounded-xl border border-slate-200 p-4 shadow-sm space-y-3 sticky top-16 z-20">
      <div class="flex flex-col md:flex-row md:items-center justify-between gap-3">
        <div class="relative flex-1">
          <i data-lucide="search" class="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400"></i>
          <input 
            type="text" 
            id="searchInput" 
            placeholder="搜尋商品名、中文名、學名/成分、院內代碼、健保碼、藥理分類、適應症、劑量、禁忌..." 
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

          <!-- Clinical Safety Filter -->
          <select id="safetyFilter" onchange="handleSearch()" class="text-xs bg-slate-50 border border-slate-200 rounded-lg px-2.5 py-2 text-slate-700 focus:ring-2 focus:ring-teal-500 font-medium text-teal-800">
            <option value="all">全部臨床安全分級 (All)</option>
            <option value="renal">需腎功能調整 [腎]</option>
            <option value="hepatic">需肝功能調整 [肝]</option>
            <option value="preg_x">懷孕禁用 (X 級)</option>
            <option value="preg_d">懷孕風險 (D 級)</option>
            <option value="preg_ab">懷孕安全 (A / B 級)</option>
          </select>

          <!-- NHI Payment Filter -->
          <select id="nhiFilter" onchange="handleSearch()" class="text-xs bg-slate-50 border border-slate-200 rounded-lg px-2.5 py-2 text-slate-700 focus:ring-2 focus:ring-teal-500">
            <option value="all">全部健保狀態</option>
            <option value="nhi_covered">健保給付 (Covered)</option>
            <option value="has_rule">有給付規定 (With Rules)</option>
            <option value="self_pay">自費/未收載 (Self-Pay)</option>
          </select>

          <button onclick="openAddDrugModal()" class="inline-flex items-center px-3.5 py-2 text-xs font-semibold rounded-lg text-white bg-teal-600 hover:bg-teal-700 transition shadow-xs">
            <i data-lucide="plus-circle" class="w-4 h-4 mr-1.5"></i>
            新增單筆
          </button>
        </div>
      </div>

      <div class="flex flex-wrap items-center justify-between text-xs text-slate-500 pt-2 border-t border-slate-100 gap-2">
        <div class="flex flex-wrap items-center gap-3">
          <span class="inline-flex items-center"><span class="w-2 h-2 rounded-full bg-emerald-500 mr-1.5"></span>綠標：健保給付</span>
          <span class="inline-flex items-center"><span class="w-2 h-2 rounded-full bg-indigo-500 mr-1.5"></span>藍標：給付規定 PDF</span>
          <span class="inline-flex items-center"><span class="px-1.5 py-0.2 rounded bg-purple-100 text-purple-700 font-bold mr-1 text-[10px]">[腎]</span>提示需依腎功能/eGFR調整</span>
          <span class="inline-flex items-center"><span class="px-1.5 py-0.2 rounded bg-blue-100 text-blue-700 font-bold mr-1 text-[10px]">[肝]</span>提示需依肝功能調整</span>
          <span class="text-teal-700 font-medium hidden lg:inline">💡 提示：點擊「📖 處方」按鈕可開啟完整處方集（用法用量、禁忌、副作用、懷孕分級）</span>
        </div>
        <div class="flex flex-wrap items-center space-x-2">
          <button onclick="loadEnrichedPreset()" class="text-teal-700 font-semibold hover:underline bg-teal-50 hover:bg-teal-100 px-2 py-1 rounded-md border border-teal-200 transition flex items-center space-x-1">
            <i data-lucide="refresh-cw" class="w-3.5 h-3.5"></i>
            <span>重設載入 115Q2 處方集主檔 (738 筆)</span>
          </button>
          <span>•</span>
          <button onclick="clearAllDataConfirm()" class="text-rose-600 hover:underline">清空資料</button>
        </div>
      </div>
    </div>

    <!-- Drug Master Data Table -->
    <div class="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
      <div class="overflow-x-auto custom-scrollbar max-h-[660px]">
        <table class="w-full text-left border-collapse table-auto">
          <thead class="bg-slate-50 border-b border-slate-200 sticky top-0 z-20 text-xs font-semibold text-slate-600 select-none shadow-xs">
            <tr>
              <!-- Cross Freeze: Column 1 (Index) Sticky Top + Left -->
              <th class="py-3 px-2 w-12 text-center sticky top-0 left-0 bg-slate-100 z-30 border-r border-slate-200">序號</th>
              
              <!-- Cross Freeze: Column 2 (Hospital Code) Sticky Top + Left -->
              <th class="py-3 px-3 w-28 cursor-pointer hover:bg-slate-200 transition sticky top-0 left-12 bg-slate-100 z-30 border-r-2 border-slate-300 shadow-[4px_0_6px_-2px_rgba(0,0,0,0.06)]" onclick="toggleSort('hospitalCode')">
                <div class="flex items-center space-x-1">
                  <span>院內代碼</span>
                  <i data-lucide="arrow-up-down" class="w-3 h-3 text-slate-400"></i>
                </div>
              </th>

              <!-- Column 3: Brand Name -->
              <th class="py-3 px-3 w-52 min-w-[180px] cursor-pointer hover:bg-slate-100 transition" onclick="toggleSort('brandName')">
                <div class="flex items-center space-x-1">
                  <span>英文商品名</span>
                  <i data-lucide="arrow-up-down" class="w-3 h-3 text-slate-400"></i>
                </div>
              </th>

              <!-- Column 4: Generic Name (加大文字 13px、粗體強調、適中欄寬、多行自動換行) -->
              <th class="py-3 px-3 w-48 min-w-[150px] max-w-[210px] bg-teal-50/70 text-teal-950 border-x border-teal-200/80 font-bold cursor-pointer hover:bg-teal-100/60 transition" onclick="toggleSort('genericName')">
                <div class="flex items-center space-x-1">
                  <span>學名 / 主成分</span>
                  <i data-lucide="arrow-up-down" class="w-3 h-3 text-teal-700"></i>
                </div>
              </th>

              <!-- Column 5: Strength / Package -->
              <th class="py-3 px-3 w-32 min-w-[120px] cursor-pointer hover:bg-slate-100 transition" onclick="toggleSort('strength')">
                <div class="flex items-center space-x-1">
                  <span>規格 / 包裝</span>
                  <i data-lucide="arrow-up-down" class="w-3 h-3 text-slate-400"></i>
                </div>
              </th>

              <!-- Column 6: Dosage Form -->
              <th class="py-3 px-3 w-16 text-center cursor-pointer hover:bg-slate-100 transition" onclick="toggleSort('dosageForm')">
                <div class="flex items-center justify-center space-x-1">
                  <span>劑型</span>
                  <i data-lucide="arrow-up-down" class="w-3 h-3 text-slate-400"></i>
                </div>
              </th>

              <!-- Column 7: NHI Code -->
              <th class="py-3 px-3 w-32 min-w-[110px] cursor-pointer hover:bg-slate-100 transition" onclick="toggleSort('nhiCode')">
                <div class="flex items-center space-x-1">
                  <span>健保代碼</span>
                  <i data-lucide="arrow-up-down" class="w-3 h-3 text-slate-400"></i>
                </div>
              </th>

              <!-- Column 8: NHI Price -->
              <th class="py-3 px-3 w-28 text-right cursor-pointer hover:bg-slate-100 transition" onclick="toggleSort('price')">
                <div class="flex items-center justify-end space-x-1">
                  <span>健保支付價</span>
                  <i data-lucide="arrow-up-down" class="w-3 h-3 text-slate-400"></i>
                </div>
              </th>

              <!-- Column 9: NHI Payment Rules -->
              <th class="py-3 px-3 w-36 min-w-[130px] cursor-pointer hover:bg-slate-100 transition" onclick="toggleSort('ruleSection')">
                <div class="flex items-center space-x-1">
                  <span>健保給付規定</span>
                  <i data-lucide="arrow-up-down" class="w-3 h-3 text-slate-400"></i>
                </div>
              </th>

              <!-- Column 10: Clinical Safety -->
              <th class="py-3 px-3 w-28 text-center cursor-pointer hover:bg-slate-100 transition" onclick="toggleSort('pregnancy')">
                <div class="flex items-center justify-center space-x-1">
                  <span>臨床安全</span>
                  <i data-lucide="arrow-up-down" class="w-3 h-3 text-slate-400"></i>
                </div>
              </th>

              <!-- Column 11: Actions -->
              <th class="py-3 px-3 w-28 text-center">查詢/操作</th>
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
      <p>衛生福利部嘉義醫院藥品臨床處方輯管理系統 • 資料來源：115 年第二季院內藥品清單、嘉義醫院處方集與中央健保署 • 管理者:藥品管理組Julia</p>
      <p class="text-slate-400">本系統純前端本機運作，資料安全保密不外傳 • 支援健保代碼、中文名、支付價、處方集臨床劑量、禁忌、副作用與給付規定 PDF</p>
    </div>
  </footer>

  <!-- Modal: QR Code -->
  <div id="qrModal" class="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-xs flex items-center justify-center hidden p-4">
    <div class="bg-white rounded-2xl max-w-sm w-full shadow-2xl border border-slate-100 overflow-hidden transform transition-all text-center">
      <div class="px-6 py-4 bg-gradient-to-r from-teal-600 to-emerald-600 text-white flex items-center justify-between">
        <div class="flex items-center space-x-2">
          <i data-lucide="qr-code" class="w-5 h-5"></i>
          <h3 class="text-sm font-bold">手機掃描立即開啟系統</h3>
        </div>
        <button onclick="closeQrModal()" class="text-white/80 hover:text-white transition">
          <i data-lucide="x" class="w-5 h-5"></i>
        </button>
      </div>
      <div class="p-6 flex flex-col items-center space-y-4">
        <div class="p-3 bg-white border-2 border-slate-100 rounded-2xl shadow-inner">
          <img src="https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=https%3A%2F%2Fjuliahml-ux.github.io%2Fhospital-drug-master%2F&margin=10" alt="QR Code" class="w-48 h-48 rounded-lg shadow-xs">
        </div>
        <div>
          <p class="text-xs font-semibold text-slate-800">衛生福利部嘉義醫院藥品臨床處方輯管理系統</p>
          <p class="text-[11px] text-slate-400 font-mono mt-1 break-all">https://juliahml-ux.github.io/hospital-drug-master/</p>
        </div>
        <button onclick="copySystemUrl()" class="w-full py-2 px-3 rounded-lg bg-teal-50 text-teal-700 hover:bg-teal-100 font-medium text-xs transition flex items-center justify-center space-x-1.5 border border-teal-200">
          <i data-lucide="copy" class="w-3.5 h-3.5"></i>
          <span>複製專屬網址</span>
        </button>
      </div>
    </div>
  </div>

  <!-- Modal: Clinical Formulary Detail Card (處方集詳細資訊抽屜彈窗) -->
  <div id="clinicalModal" class="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center hidden p-4">
    <div class="bg-white rounded-2xl max-w-3xl w-full shadow-2xl border border-slate-200 overflow-hidden transform transition-all flex flex-col max-h-[90vh]">
      <div class="px-6 py-4 bg-gradient-to-r from-teal-700 via-teal-600 to-emerald-700 text-white flex items-center justify-between shrink-0">
        <div class="flex items-center space-x-3">
          <div class="w-9 h-9 rounded-xl bg-white/20 flex items-center justify-center backdrop-blur-md">
            <i data-lucide="book-open" class="w-5 h-5 text-white"></i>
          </div>
          <div>
            <h3 id="clinicModalTitle" class="text-base font-bold leading-tight">臨床處方集與用藥安全指引</h3>
            <p id="clinicModalSubtitle" class="text-xs text-teal-100 mt-0.5 font-mono">院內代碼 • 藥理分類</p>
          </div>
        </div>
        <button onclick="closeClinicalModal()" class="text-white/80 hover:text-white p-1 rounded-lg hover:bg-white/10 transition">
          <i data-lucide="x" class="w-5 h-5"></i>
        </button>
      </div>

      <div class="p-6 overflow-y-auto space-y-4 text-xs flex-1 custom-scrollbar">
        <!-- Top Drug Quick Info Grid -->
        <div class="grid grid-cols-1 sm:grid-cols-3 gap-3 bg-slate-50 p-3.5 rounded-xl border border-slate-200">
          <div>
            <span class="text-[11px] text-slate-500 block">健保中文品名</span>
            <span id="clinicChineseName" class="font-bold text-slate-900 text-sm mt-0.5 block">-</span>
          </div>
          <div>
            <span class="text-[11px] text-slate-500 block">學名 / 主成分</span>
            <span id="clinicGenericName" class="font-bold text-teal-900 text-xs mt-0.5 block">-</span>
          </div>
          <div>
            <span class="text-[11px] text-slate-500 block">規格 / 劑型</span>
            <span id="clinicSpec" class="font-semibold text-slate-800 text-xs mt-0.5 block">-</span>
          </div>
        </div>

        <!-- Clinical Alerts Badges (Pregnancy, Hepatic, Renal) -->
        <div class="flex flex-wrap items-center gap-2 pt-1" id="clinicAlertBadges"></div>

        <!-- Dosage & Administration -->
        <div class="bg-teal-50/50 rounded-xl p-3.5 border border-teal-200/80">
          <div class="flex items-center space-x-2 text-teal-900 font-bold mb-1.5">
            <i data-lucide="activity" class="w-4 h-4 text-teal-600"></i>
            <span>臨床用法用量與劑量上限 (Dosage & Max Dose)</span>
          </div>
          <p id="clinicDosage" class="text-slate-800 font-medium whitespace-pre-line leading-relaxed text-[13px] bg-white p-2.5 rounded-lg border border-teal-100">
            -
          </p>
        </div>

        <!-- Contraindications & Adverse Effects -->
        <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
          <div class="bg-rose-50/50 rounded-xl p-3.5 border border-rose-200">
            <div class="flex items-center space-x-2 text-rose-900 font-bold mb-1.5">
              <i data-lucide="shield-alert" class="w-4 h-4 text-rose-600"></i>
              <span>使用禁忌 (Contraindications)</span>
            </div>
            <p id="clinicCi" class="text-rose-950 font-normal whitespace-pre-line leading-relaxed bg-white p-2.5 rounded-lg border border-rose-100 min-h-[60px]">
              無特殊禁忌記載
            </p>
          </div>

          <div class="bg-amber-50/50 rounded-xl p-3.5 border border-amber-200">
            <div class="flex items-center space-x-2 text-amber-900 font-bold mb-1.5">
              <i data-lucide="alert-circle" class="w-4 h-4 text-amber-600"></i>
              <span>重要副作用 (Adverse Effects)</span>
            </div>
            <p id="clinicAe" class="text-amber-950 font-normal whitespace-pre-line leading-relaxed bg-white p-2.5 rounded-lg border border-amber-100 min-h-[60px]">
              -
            </p>
          </div>
        </div>

        <!-- Clinical Notes & Precautions -->
        <div class="bg-slate-50 rounded-xl p-3.5 border border-slate-200">
          <div class="flex items-center space-x-2 text-slate-800 font-bold mb-1.5">
            <i data-lucide="file-text" class="w-4 h-4 text-slate-600"></i>
            <span>臨床使用注意事項 / 警語 (Clinical Notes & Warnings)</span>
          </div>
          <p id="clinicNote" class="text-slate-700 whitespace-pre-line leading-relaxed bg-white p-2.5 rounded-lg border border-slate-200">
            -
          </p>
        </div>

        <!-- NHI & Payment Rule Details -->
        <div class="bg-indigo-50/50 rounded-xl p-3.5 border border-indigo-200 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3">
          <div>
            <span class="text-indigo-950 font-bold block">健保給付規定章節與價格</span>
            <span id="clinicNhiInfo" class="text-indigo-900 text-[11px] mt-0.5 block">健保碼：- • 支付價：NT$ -</span>
          </div>
          <div id="clinicNhiRuleAction"></div>
        </div>
      </div>

      <div class="px-6 py-3.5 bg-slate-50 border-t border-slate-200 flex items-center justify-between shrink-0">
        <button onclick="copyClinicSummary()" class="px-3 py-1.5 rounded-lg bg-teal-50 text-teal-700 hover:bg-teal-100 font-medium text-xs border border-teal-200 transition flex items-center space-x-1">
          <i data-lucide="copy" class="w-3.5 h-3.5"></i>
          <span>複製臨床重點摘要</span>
        </button>
        <button onclick="closeClinicalModal()" class="px-4 py-1.5 rounded-lg bg-slate-200 text-slate-700 hover:bg-slate-300 font-medium text-xs transition">關閉</button>
      </div>
    </div>
  </div>

  <!-- Modal: Add / Edit Drug -->
  <div id="drugModal" class="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-xs flex items-center justify-center hidden p-4">
    <div class="bg-white rounded-2xl max-w-2xl w-full shadow-2xl border border-slate-100 overflow-hidden transform transition-all max-h-[90vh] flex flex-col">
      <div class="px-6 py-4 bg-gradient-to-r from-teal-600 to-emerald-600 text-white flex items-center justify-between shrink-0">
        <div class="flex items-center space-x-2">
          <i data-lucide="edit-3" class="w-5 h-5"></i>
          <h3 id="modalTitle" class="text-base font-bold">編輯藥品主檔與臨床資料</h3>
        </div>
        <button onclick="closeDrugModal()" class="text-white/80 hover:text-white transition">
          <i data-lucide="x" class="w-5 h-5"></i>
        </button>
      </div>

      <form id="drugForm" onsubmit="saveDrugForm(event)" class="p-6 space-y-4 text-xs overflow-y-auto custom-scrollbar flex-1">
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
            <label class="block font-semibold text-slate-700 mb-1">英文商品名 (Brand Name) <span class="text-rose-500">*</span></label>
            <input type="text" id="formBrandName" required placeholder="例：Norvasc Tablets 5mg" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
          </div>

          <div>
            <label class="block font-semibold text-slate-700 mb-1">健保中文品名 (Chinese Name)</label>
            <input type="text" id="formChineseName" placeholder="例：脈優錠 5 毫克" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
          </div>
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <label class="block font-semibold text-slate-700 mb-1">學名 / 主成分 (Generic Name) <span class="text-rose-500">*</span></label>
            <input type="text" id="formGenericName" required placeholder="例：Amlodipine besylate" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
          </div>

          <div>
            <label class="block font-semibold text-slate-700 mb-1">ATC 碼</label>
            <input type="text" id="formAtc" placeholder="例：C08CA01" class="w-full px-3 py-2 border border-slate-300 rounded-lg uppercase tracking-wider font-mono focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
          </div>
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
          <div>
            <label class="block font-semibold text-slate-700 mb-1">規格 / 包裝</label>
            <input type="text" id="formStrength" placeholder="例：5mg/tab (Box)" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
          </div>

          <div>
            <label class="block font-semibold text-slate-700 mb-1">劑型 (Dosage Form)</label>
            <input type="text" id="formDosageForm" placeholder="例：O, I, E, 錠劑" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
          </div>

          <div>
            <label class="block font-semibold text-slate-700 mb-1">懷孕分級 (Category)</label>
            <select id="formPregnancy" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500">
              <option value="">未註明</option>
              <option value="A">A 級 (安全)</option>
              <option value="B">B 級 (相對安全)</option>
              <option value="C">C 級 (謹慎使用)</option>
              <option value="D">D 級 (有危險性)</option>
              <option value="X">X 級 (孕婦禁用)</option>
            </select>
          </div>
        </div>

        <div class="flex items-center space-x-6 p-3 bg-slate-50 rounded-lg border border-slate-200">
          <span class="font-semibold text-slate-700">臨床劑量調整警示：</span>
          <label class="inline-flex items-center cursor-pointer">
            <input type="checkbox" id="formRenAlert" class="rounded text-purple-600 focus:ring-purple-500 mr-1.5">
            <span class="font-bold text-purple-800">[腎] 需依腎功能/eGFR調整</span>
          </label>
          <label class="inline-flex items-center cursor-pointer">
            <input type="checkbox" id="formHepAlert" class="rounded text-blue-600 focus:ring-blue-500 mr-1.5">
            <span class="font-bold text-blue-800">[肝] 需依肝功能調整</span>
          </label>
        </div>

        <div>
          <label class="block font-semibold text-slate-700 mb-1">臨床用法用量 (Dosage & Max Dose)</label>
          <textarea id="formDosage" rows="2" placeholder="例：2.5-5mg daily (max. 10mg/day)" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden"></textarea>
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <label class="block font-semibold text-slate-700 mb-1">使用禁忌 (CI)</label>
            <input type="text" id="formCi" placeholder="例：對本藥過敏者禁用" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
          </div>
          <div>
            <label class="block font-semibold text-slate-700 mb-1">常見副作用 (AE)</label>
            <input type="text" id="formAe" placeholder="例：周邊水腫、心悸、頭痛" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
          </div>
        </div>

        <div>
          <label class="block font-semibold text-slate-700 mb-1">臨床注意事項 / 警語 (Notes)</label>
          <textarea id="formNote" rows="2" placeholder="臨床注意事項、磨粉限制、特殊監測..." class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden"></textarea>
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <label class="block font-semibold text-slate-700 mb-1">健保給付規定章節</label>
            <input type="text" id="formRuleSection" placeholder="例：1.2.2.2." class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden">
          </div>
          <div>
            <label class="block font-semibold text-slate-700 mb-1">給付規定 PDF 連結</label>
            <input type="text" id="formRuleLink" placeholder="https://info.nhi.gov.tw/api/..." class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 focus:outline-hidden font-mono text-[11px]">
          </div>
        </div>

        <div class="pt-4 border-t border-slate-100 flex items-center justify-end space-x-3 shrink-0">
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
    // Embedded 738 Clinical Drugs with Formulary Clinical Data & NHI Prices
    const HOSPITAL_ENRICHED_PRESET = {json_data};

    const STORAGE_KEY = "hospital_drug_master_db_115_q2_v1";
    let drugs = [];
    let currentPage = 1;
    let pageSize = 15;
    let sortColumn = "seq";
    let sortAsc = true;
    let currentClinicalDrug = null;

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
      if (confirm("確定要重設載入「115 年第二季院內藥品清單與處方集 (738 筆，含處方劑量、健保價與給付規定)」嗎？")) {{
        drugs = JSON.parse(JSON.stringify(HOSPITAL_ENRICHED_PRESET));
        saveDataToStorage();
        renderAll();
        showToast(`已成功載入 738 筆第二季藥品主檔！`, "success");
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

    function renderAll() {{
      updateKpis();
      renderTable();
      lucide.createIcons();
    }}

    function updateKpis() {{
      const total = drugs.length;
      let nhiCovered = 0;
      let formularyCount = 0;
      let alertCount = 0;

      drugs.forEach(d => {{
        const nhi = (d.nhiCode || "").trim();
        if (nhi && nhi !== '自費/未收載' && nhi.length === 10) {{
          nhiCovered++;
        }}
        if (d.dosage || d.ci || d.pregnancy) {{
          formularyCount++;
        }}
        if (d.renAlert || d.hepAlert) {{
          alertCount++;
        }}
      }});

      document.getElementById("kpiTotalCount").innerText = total;
      document.getElementById("kpiNhiCoveredCount").innerText = nhiCovered;
      document.getElementById("kpiFormularyCount").innerText = formularyCount;
      document.getElementById("kpiAlertCount").innerText = alertCount;
    }}

    function getFilteredDrugs() {{
      const q = document.getElementById("searchInput").value.trim().toLowerCase();
      const dosageFilter = document.getElementById("dosageFilter").value;
      const safetyFilter = document.getElementById("safetyFilter").value;
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
          (drug.ruleSection && drug.ruleSection.toLowerCase().includes(q)) ||
          (drug.dosage && drug.dosage.toLowerCase().includes(q)) ||
          (drug.ci && drug.ci.toLowerCase().includes(q)) ||
          (drug.pharmCategory && drug.pharmCategory.toLowerCase().includes(q));

        if (!matchQuery) return false;

        if (dosageFilter && !(drug.dosageForm || "").includes(dosageFilter)) {{
          return false;
        }}

        if (safetyFilter !== "all") {{
          if (safetyFilter === "renal" && !drug.renAlert) return false;
          if (safetyFilter === "hepatic" && !drug.hepAlert) return false;
          if (safetyFilter === "preg_x" && drug.pregnancy !== "X") return false;
          if (safetyFilter === "preg_d" && drug.pregnancy !== "D") return false;
          if (safetyFilter === "preg_ab" && drug.pregnancy !== "A" && drug.pregnancy !== "B") return false;
        }}

        if (nhiFilter !== "all") {{
          const isCovered = drug.nhiCode && drug.nhiCode.trim() && drug.nhiCode !== '自費/未收載' && drug.nhiCode.trim().length === 10;
          const hasRule = drug.ruleSection && drug.ruleSection.trim() && drug.ruleSection !== '無特殊章節';
          if (nhiFilter === "nhi_covered" && !isCovered) return false;
          if (nhiFilter === "has_rule" && !hasRule) return false;
          if (nhiFilter === "self_pay" && isCovered) return false;
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
        if (sortColumn === "seq" || sortColumn === "price") {{
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
            const rawSections = (drug.ruleSection || '').split(',').map(s => s.trim()).filter(Boolean);
            const rawLinks = (drug.ruleLink || '').split(',').map(l => l.trim()).filter(Boolean);
            
            if (rawLinks.length > 0) {{
              const badges = rawLinks.map((link, idx) => {{
                const match = link.match(/DurgFileName=([0-9.]+)_/);
                const secLabel = match ? match[1] : (rawSections[idx] || rawSections[0] || '給付規定');
                return `
                  <a href="${{link}}" target="_blank" class="inline-flex items-center px-2 py-0.5 rounded text-[11px] font-semibold bg-indigo-50 text-indigo-700 border border-indigo-200 hover:bg-indigo-100 transition shadow-2xs group-hover:border-indigo-300 whitespace-nowrap" title="點擊檢視健保給付規定【${{escapeHtml(secLabel)}}】PDF">
                    <i data-lucide="file-text" class="w-3 h-3 mr-1 text-indigo-500 shrink-0"></i>
                    <span>${{escapeHtml(secLabel)}} PDF</span>
                  </a>
                `;
              }});
              ruleDisplay = `<div class="flex flex-col gap-1 items-start">${{badges.join('')}}</div>`;
            }} else if (rawSections.length > 0) {{
              const badges = rawSections.map(sec => `<span class="inline-flex items-center px-1.5 py-0.5 rounded text-[11px] font-medium bg-slate-100 text-slate-700 font-mono whitespace-nowrap">${{escapeHtml(sec)}}</span>`);
              ruleDisplay = `<div class="flex flex-col gap-1 items-start">${{badges.join('')}}</div>`;
            }}
          }}

          // Clinical Safety Badges (Pregnancy & Hepatic/Renal)
          let safetyBadges = [];
          if (drug.pregnancy) {{
            const p = drug.pregnancy.toUpperCase().trim();
            if (p === 'A' || p === 'B') {{
              safetyBadges.push(`<span class="px-1.5 py-0.5 rounded text-[10px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-200" title="懷孕分級 ${{p}}（相對安全）">${{p}} 級</span>`);
            }} else if (p === 'C') {{
              safetyBadges.push(`<span class="px-1.5 py-0.5 rounded text-[10px] font-bold bg-amber-100 text-amber-800 border border-amber-200" title="懷孕分級 C（謹慎評估）">C 級</span>`);
            }} else if (p === 'D') {{
              safetyBadges.push(`<span class="px-1.5 py-0.5 rounded text-[10px] font-bold bg-orange-100 text-orange-800 border border-orange-200" title="懷孕分級 D（有風險）">D 級</span>`);
            }} else if (p === 'X') {{
              safetyBadges.push(`<span class="px-1.5 py-0.5 rounded text-[10px] font-bold bg-rose-100 text-rose-800 border border-rose-200" title="懷孕分級 X（孕婦禁用）">X 級禁</span>`);
            }} else {{
              safetyBadges.push(`<span class="px-1.5 py-0.5 rounded text-[10px] font-medium bg-slate-100 text-slate-600">${{escapeHtml(p)}}</span>`);
            }}
          }}
          if (drug.renAlert) {{
            safetyBadges.push(`<span class="px-1.5 py-0.5 rounded text-[10px] font-bold bg-purple-100 text-purple-800 border border-purple-200" title="腎功能不全/透析需調整劑量">[腎]</span>`);
          }}
          if (drug.hepAlert) {{
            safetyBadges.push(`<span class="px-1.5 py-0.5 rounded text-[10px] font-bold bg-blue-100 text-blue-800 border border-blue-200" title="肝功能不全需注意劑量">[肝]</span>`);
          }}
          const safetyDisplay = safetyBadges.length > 0 ? `<div class="flex flex-wrap items-center justify-center gap-1">${{safetyBadges.join('')}}</div>` : `<span class="text-slate-400 text-[11px]">-</span>`;

          // External NHI Link
          let queryUrl = drug.nhiLink;
          if (!queryUrl || queryUrl === 'https://info.nhi.gov.tw/IODE0000/IODE0000S06') {{
            queryUrl = `https://info.nhi.gov.tw/IODE0000/IODE0000S06`;
          }}

          return `
            <tr class="hover:bg-teal-50/40 transition group border-b border-slate-100">
              <!-- 1. 序號 (Cross Freeze: Sticky Left 0) -->
              <td class="py-3 px-2 text-center font-mono text-slate-400 text-xs sticky left-0 bg-white group-hover:bg-slate-50 z-10 border-r border-slate-200">${{drug.seq || rowNum}}</td>

              <!-- 2. 院內代碼 (Cross Freeze: Sticky Left 12) -->
              <td class="py-3 px-3 sticky left-12 bg-white group-hover:bg-slate-50 z-10 border-r-2 border-slate-300 shadow-[4px_0_6px_-2px_rgba(0,0,0,0.06)]">
                <span class="bg-slate-100 text-slate-900 px-2 py-0.5 rounded border border-slate-300 text-xs font-bold font-mono tracking-wide">${{escapeHtml(drug.hospitalCode || "-")}}</span>
              </td>

              <!-- 3. 英文商品名 (+ 健保中文品名與藥理分類) -->
              <td class="py-3 px-3 w-52 min-w-[180px]">
                <div onclick="openClinicalModal('${{drug.id}}')" class="cursor-pointer group/title">
                  <div class="break-words whitespace-normal text-xs font-semibold text-slate-900 group-hover/title:text-teal-700 transition leading-snug" title="${{escapeHtml(drug.brandName || '-')}}">
                    ${{escapeHtml(drug.brandName || "-")}}
                  </div>
                  ${{drug.chineseName ? `<div class="text-slate-500 text-[11px] font-normal leading-snug mt-0.5 break-words" title="${{escapeHtml(drug.chineseName)}}">${{escapeHtml(drug.chineseName)}}</div>` : ''}}
                </div>
                <div class="flex flex-wrap items-center gap-1 mt-1">
                  ${{drug.atc ? `<span class="text-[10px] font-mono text-teal-800 bg-teal-50 px-1.5 py-0.5 rounded border border-teal-200 font-semibold">${{escapeHtml(drug.atc)}}</span>` : ''}}
                </div>
              </td>

              <!-- 4. 學名 / 主成分 (加大文字 13px、粗體強調、適中欄寬、多行自動換行) -->
              <td class="py-3 px-3 bg-teal-50/30 border-x border-teal-100/80 w-48 min-w-[150px] max-w-[210px]">
                <div class="text-[13px] font-bold text-slate-900 break-words break-all whitespace-normal leading-snug tracking-tight" title="${{escapeHtml(drug.genericName || '-')}}">
                  ${{escapeHtml(drug.genericName || "-")}}
                </div>
              </td>

              <!-- 5. 規格 / 包裝 -->
              <td class="py-3 px-3 text-slate-700 text-xs w-32 min-w-[120px]">
                <div class="break-words whitespace-normal leading-snug">${{escapeHtml(drug.strength || "-")}}</div>
              </td>

              <!-- 6. 劑型 -->
              <td class="py-3 px-3 text-center w-16">
                <span class="bg-slate-100 text-slate-700 px-2 py-0.5 rounded text-[11px] font-semibold font-mono">${{escapeHtml(drug.dosageForm || "-")}}</span>
              </td>

              <!-- 7. 健保代碼 -->
              <td class="py-3 px-3 w-32 min-w-[110px]">${{nhiDisplay}}</td>

              <!-- 8. 健保支付價 -->
              <td class="py-3 px-3 text-right font-mono w-28">${{priceDisplay}}</td>

              <!-- 9. 健保給付規定 -->
              <td class="py-3 px-3 w-36 min-w-[130px]">${{ruleDisplay}}</td>

              <!-- 10. 臨床安全分級 -->
              <td class="py-3 px-3 text-center w-28">${{safetyDisplay}}</td>

              <!-- 11. 查詢/操作 -->
              <td class="py-3 px-3 text-center w-28">
                <div class="flex items-center justify-center space-x-1">
                  <button onclick="openClinicalModal('${{drug.id}}')" title="開啟臨床處方集卡片" class="p-1.5 text-teal-700 hover:text-teal-900 hover:bg-teal-50 rounded-md transition border border-teal-200 shadow-2xs">
                    <i data-lucide="book-open" class="w-3.5 h-3.5"></i>
                  </button>
                  <a href="${{queryUrl}}" target="_blank" title="健保署/食藥署官網查詢" class="p-1.5 text-slate-400 hover:text-sky-600 hover:bg-sky-50 rounded-md transition">
                    <i data-lucide="external-link" class="w-3.5 h-3.5"></i>
                  </a>
                  <button onclick="openEditDrugModal('${{drug.id}}')" title="編輯藥品" class="p-1.5 text-slate-400 hover:text-teal-600 hover:bg-teal-50 rounded-md transition">
                    <i data-lucide="edit-2" class="w-3.5 h-3.5"></i>
                  </button>
                  <button onclick="deleteDrug('${{drug.id}}')" title="刪除" class="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-md transition">
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

    function openClinicalModal(id) {{
      const drug = drugs.find(d => d.id === id);
      if (!drug) return;
      currentClinicalDrug = drug;

      document.getElementById("clinicModalTitle").innerText = `${{drug.brandName}}`;
      document.getElementById("clinicModalSubtitle").innerText = `院內代碼：${{drug.hospitalCode}} • 藥理分類：${{drug.pharmCategory || '臨床藥理分類'}}`;
      document.getElementById("clinicChineseName").innerText = drug.chineseName || drug.brandName;
      document.getElementById("clinicGenericName").innerText = drug.genericName || '-';
      document.getElementById("clinicSpec").innerText = `${{drug.strength || '-'}} [${{drug.dosageForm || '-'}}]`;

      // Badges
      let badgesHtml = [];
      if (drug.pregnancy) {{
        badgesHtml.push(`<span class="px-2.5 py-1 rounded-lg text-xs font-bold bg-amber-100 text-amber-900 border border-amber-300">🤰 懷孕分級：Category ${{drug.pregnancy}}</span>`);
      }}
      if (drug.renAlert) {{
        badgesHtml.push(`<span class="px-2.5 py-1 rounded-lg text-xs font-bold bg-purple-100 text-purple-900 border border-purple-300">🩺 [腎] 需依腎功能 / eGFR 調整劑量</span>`);
      }}
      if (drug.hepAlert) {{
        badgesHtml.push(`<span class="px-2.5 py-1 rounded-lg text-xs font-bold bg-blue-100 text-blue-900 border border-blue-300">🩺 [肝] 需依肝功能 / LFT 調整劑量</span>`);
      }}
      if (drug.atc) {{
        badgesHtml.push(`<span class="px-2 py-1 rounded-lg text-xs font-mono font-semibold bg-slate-100 text-slate-700 border border-slate-200">ATC: ${{drug.atc}}</span>`);
      }}
      document.getElementById("clinicAlertBadges").innerHTML = badgesHtml.join("");

      document.getElementById("clinicDosage").innerText = drug.dosage || "處方集未登錄特定劑量，請依仿單及醫師處方使用。";
      document.getElementById("clinicCi").innerText = drug.ci || "無特定禁忌登錄（對本成分過敏者請勿使用）。";
      document.getElementById("clinicAe").innerText = drug.ae || "常見副作用請參考仿單說明。";
      document.getElementById("clinicNote").innerText = drug.note || "請遵照醫囑與藥師衛教指示服藥。";

      const priceText = drug.price ? `NT$ ${{parseFloat(drug.price).toFixed(2)}}` : '自費/未收載';
      document.getElementById("clinicNhiInfo").innerText = `健保碼：${{drug.nhiCode || '-'}} • 支付價：${{priceText}} • 規定章節：${{drug.ruleSection || '無特殊章節'}}`;

      if (drug.ruleLink && drug.ruleLink.trim()) {{
        const rawSections = (drug.ruleSection || '').split(',').map(s => s.trim()).filter(Boolean);
        const rawLinks = drug.ruleLink.split(',').map(l => l.trim()).filter(Boolean);

        const btnHtml = rawLinks.map((link, idx) => {{
          const match = link.match(/DurgFileName=([0-9.]+)_/);
          const secLabel = match ? match[1] : (rawSections[idx] || rawSections[0] || '給付規定');
          return `
            <a href="${{link}}" target="_blank" class="px-3 py-1.5 rounded-lg bg-indigo-600 text-white hover:bg-indigo-700 font-medium text-xs transition flex items-center space-x-1.5 shadow-sm">
              <i data-lucide="file-text" class="w-3.5 h-3.5 shrink-0"></i>
              <span>下載【${{escapeHtml(secLabel)}}】給付規定 PDF</span>
            </a>
          `;
        }}).join('');

        document.getElementById("clinicNhiRuleAction").innerHTML = `
          <div class="flex flex-wrap items-center gap-2">${{btnHtml}}</div>
        `;
      }} else {{
        document.getElementById("clinicNhiRuleAction").innerHTML = `
          <span class="text-xs text-indigo-700 italic">無附帶特殊給付規範 PDF</span>
        `;
      }}

      document.getElementById("clinicalModal").classList.remove("hidden");
      lucide.createIcons();
    }}

    function closeClinicalModal() {{
      document.getElementById("clinicalModal").classList.add("hidden");
    }}

    function copyClinicSummary() {{
      if (!currentClinicalDrug) return;
      const d = currentClinicalDrug;
      const text = `【藥品臨床摘要】${{d.brandName}} (${{d.chineseName || ''}}) [${{d.hospitalCode}}]
學名：${{d.genericName}}
規格：${{d.strength}} [${{d.dosageForm}}]
用法用量：${{d.dosage || '依醫囑'}}
懷孕分級：${{d.pregnancy || '未註明'}} | 肝腎調整：${{d.renAlert ? '[腎]' : ''}}${{d.hepAlert ? '[肝]' : ''}}
禁忌：${{d.ci || '無'}}
副作用：${{d.ae || '無'}}
注意事項：${{d.note || '無'}}`;

      navigator.clipboard.writeText(text).then(() => {{
        showToast("已複製藥品臨床重點摘要到剪貼簿！", "success");
      }});
    }}

    function openQrModal() {{
      document.getElementById("qrModal").classList.remove("hidden");
      lucide.createIcons();
    }}

    function closeQrModal() {{
      document.getElementById("qrModal").classList.add("hidden");
    }}

    function copySystemUrl() {{
      const url = "https://juliahml-ux.github.io/hospital-drug-master/";
      navigator.clipboard.writeText(url).then(() => {{
        showToast("已成功複製專屬網址到剪貼簿！", "success");
      }});
    }}

    function openAddDrugModal() {{
      document.getElementById("modalTitle").innerText = "新增藥品主檔與臨床資料";
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

      document.getElementById("modalTitle").innerText = "編輯藥品主檔與臨床處方資料";
      document.getElementById("formDrugId").value = drug.id;
      document.getElementById("formHospitalCode").value = drug.hospitalCode || "";
      document.getElementById("formBrandName").value = drug.brandName || "";
      document.getElementById("formChineseName").value = drug.chineseName || "";
      document.getElementById("formGenericName").value = drug.genericName || "";
      document.getElementById("formStrength").value = drug.strength || "";
      document.getElementById("formDosageForm").value = drug.dosageForm || "";
      document.getElementById("formNhiCode").value = drug.nhiCode || "";
      document.getElementById("formPrice").value = drug.price || "";
      document.getElementById("formPregnancy").value = drug.pregnancy || "";
      document.getElementById("formRenAlert").checked = !!drug.renAlert;
      document.getElementById("formHepAlert").checked = !!drug.hepAlert;
      document.getElementById("formDosage").value = drug.dosage || "";
      document.getElementById("formCi").value = drug.ci || "";
      document.getElementById("formAe").value = drug.ae || "";
      document.getElementById("formNote").value = drug.note || "";
      document.getElementById("formRuleSection").value = drug.ruleSection || "";
      document.getElementById("formRuleLink").value = drug.ruleLink || "";
      document.getElementById("formAtc").value = drug.atc || "";

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
      const brandName = document.getElementById("formBrandName").value.trim();
      const chineseName = document.getElementById("formChineseName").value.trim();
      const genericName = document.getElementById("formGenericName").value.trim();
      const strength = document.getElementById("formStrength").value.trim();
      const dosageForm = document.getElementById("formDosageForm").value.trim();
      const nhiCode = document.getElementById("formNhiCode").value.trim().toUpperCase();
      const price = document.getElementById("formPrice").value.trim();
      const pregnancy = document.getElementById("formPregnancy").value.trim();
      const renAlert = document.getElementById("formRenAlert").checked;
      const hepAlert = document.getElementById("formHepAlert").checked;
      const dosage = document.getElementById("formDosage").value.trim();
      const ci = document.getElementById("formCi").value.trim();
      const ae = document.getElementById("formAe").value.trim();
      const note = document.getElementById("formNote").value.trim();
      const ruleSection = document.getElementById("formRuleSection").value.trim();
      const ruleLink = document.getElementById("formRuleLink").value.trim();
      const atc = document.getElementById("formAtc").value.trim().toUpperCase();

      if (id) {{
        const index = drugs.findIndex(d => d.id === id);
        if (index !== -1) {{
          drugs[index] = {{ 
            ...drugs[index], 
            id, hospitalCode, brandName, chineseName, genericName, 
            strength, dosageForm, nhiCode, price, pregnancy, renAlert, hepAlert, 
            dosage, ci, ae, note, ruleSection, ruleLink, atc 
          }};
          showToast("已更新藥品主檔與臨床資料", "success");
        }}
      }} else {{
        const newDrug = {{
          id: "d_q2_" + Date.now(),
          seq: drugs.length + 1,
          hospitalCode,
          brandName,
          chineseName,
          genericName,
          strength,
          dosageForm,
          nhiCode,
          price,
          pregnancy,
          renAlert,
          hepAlert,
          dosage,
          ci,
          ae,
          note,
          ruleSection,
          ruleLink,
          atc,
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
      const headers = ["序號", "院內代碼", "英文商品名", "健保中文品名", "學名/主成分", "規格/包裝", "劑型", "健保代碼", "健保支付價", "健保給付規定章節", "給付規定PDF連結", "懷孕分級", "肝腎劑量調整", "臨床用法用量", "使用禁忌", "副作用", "注意事項", "ATC碼"];
      const templateData = [
        headers,
        [1, "ONORV5", "Norvasc Tablets 5mg", "脈優錠 5 毫克", "Amlodipine besylate", "5mg/tab", "O", "BC19248100", "5.60", "無特殊章節", "", "C", "無", "2.5-5mg daily (max. 10mg/day)", "對本成分過敏者禁用", "周邊水腫、心悸", "主動脈瓣狹窄慎用", "C08CA01"],
        [2, "OGLUC5", "Glucophage Tablets 500mg", "庫魯化錠 500 毫克", "Metformin hydrochloride", "500mg/tab", "O", "AC34125100", "1.50", "無特殊章節", "", "B", "[腎]", "500mg bid-tid (max. 2550mg/day)", "eGFR<30禁用", "胃腸不適、乳酸中毒", "腎功能不全監測", "A10BA02"]
      ];

      const ws = XLSX.utils.aoa_to_sheet(templateData);
      ws['!cols'] = [{{ wch: 8 }}, {{ wch: 14 }}, {{ wch: 32 }}, {{ wch: 25 }}, {{ wch: 30 }}, {{ wch: 18 }}, {{ wch: 10 }}, {{ wch: 16 }}, {{ wch: 14 }}, {{ wch: 20 }}, {{ wch: 35 }}, {{ wch: 10 }}, {{ wch: 16 }}, {{ wch: 35 }}, {{ wch: 30 }}, {{ wch: 30 }}, {{ wch: 30 }}, {{ wch: 12 }}];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "115Q2處方主檔標準範本");
      XLSX.writeFile(wb, "115Q2藥品主檔標準匯入範本.xlsx");
      showToast("已下載 Excel 標準匯入範本", "success");
    }}

    function exportToExcel() {{
      if (drugs.length === 0) {{
        alert("目前清單中沒有資料可供匯出！");
        return;
      }}

      const exportData = drugs.map((d, i) => {{
        const alerts = [];
        if (d.renAlert) alerts.push("[腎]");
        if (d.hepAlert) alerts.push("[肝]");
        const alertStr = alerts.join(" ") || "無特殊標記";

        return {{
          "序號": d.seq || i + 1,
          "院內代碼": d.hospitalCode || "",
          "英文商品名": d.brandName || "",
          "健保中文品名": d.chineseName || "",
          "學名/主成分": d.genericName || "",
          "規格/包裝": d.strength || "",
          "劑型": d.dosageForm || "",
          "健保代碼": d.nhiCode || "自費/未收載",
          "健保支付價": d.price || "",
          "健保給付規定章節": d.ruleSection || "無特殊章節",
          "給付規定PDF連結": d.ruleLink || "",
          "懷孕分級": d.pregnancy || "",
          "肝腎劑量調整": alertStr,
          "藥理分類": d.pharmCategory || "",
          "臨床用法用量": d.dosage || "",
          "使用禁忌(CI)": d.ci || "",
          "常見副作用(AE)": d.ae || "",
          "臨床注意事項/警語": d.note || "",
          "ATC碼": d.atc || "",
          "健保/食藥署查詢連結": d.nhiLink || ""
        }};
      }});

      const ws = XLSX.utils.json_to_sheet(exportData);
      ws['!cols'] = [{{ wch: 8 }}, {{ wch: 14 }}, {{ wch: 32 }}, {{ wch: 25 }}, {{ wch: 30 }}, {{ wch: 18 }}, {{ wch: 10 }}, {{ wch: 16 }}, {{ wch: 14 }}, {{ wch: 20 }}, {{ wch: 35 }}, {{ wch: 10 }}, {{ wch: 16 }}, {{ wch: 30 }}, {{ wch: 35 }}, {{ wch: 35 }}, {{ wch: 35 }}, {{ wch: 40 }}, {{ wch: 12 }}, {{ wch: 35 }}];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "藥品臨床處方輯主檔");
      const today = new Date().toISOString().slice(0, 10).replace(/-/g, "");
      XLSX.writeFile(wb, `衛生福利部嘉義醫院藥品臨床處方輯_${{today}}.xlsx`);
      showToast(`已成功匯出 ${{drugs.length}} 筆臨床藥品資料！`, "success");
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
          const brandIdx = getColIdx(["英文商品名", "商品名", "英文名", "品名", "brand", "trade"]);
          const genericIdx = getColIdx(["學名", "成分", "主成分", "generic"]);
          const strengthIdx = getColIdx(["規格", "劑量", "含量", "strength", "spec"]);
          const unitIdx = getColIdx(["單位", "包裝", "unit"]);
          const dosageIdx = getColIdx(["劑型", "form", "dosageform"]);
          const nhiCodeIdx = getColIdx(["健保代碼", "健保碼", "健保", "nhi"]);
          const priceIdx = getColIdx(["支付價", "健保價", "價格", "price"]);
          const ruleIdx = getColIdx(["給付規定", "章節", "rule"]);
          const ruleLinkIdx = getColIdx(["給付規定pdf", "規定連結", "rulelink"]);
          const chineseIdx = getColIdx(["中文品名", "中文名", "中文", "chinese"]);
          const pregIdx = getColIdx(["懷孕", "pregnancy", "category"]);
          const doseIdx = getColIdx(["用法用量", "臨床劑量", "dosage"]);
          const ciIdx = getColIdx(["禁忌", "ci", "contraindication"]);
          const aeIdx = getColIdx(["副作用", "ae", "adverse"]);
          const noteIdx = getColIdx(["注意事項", "警語", "note", "warning"]);
          const atcIdx = getColIdx(["atc", "atc碼", "atc代碼"]);

          const parsedItems = [];
          for (let i = 1; i < rawRows.length; i++) {{
            const row = rawRows[i];
            if (!row || row.length === 0) continue;

            const hospitalCode = (hCodeIdx !== -1 && row[hCodeIdx] ? row[hCodeIdx].toString().trim() : "").toUpperCase();
            const brandName = brandIdx !== -1 && row[brandIdx] ? row[brandIdx].toString().trim() : "";
            const genericName = genericIdx !== -1 && row[genericIdx] ? row[genericIdx].toString().trim() : "";
            const specVal = strengthIdx !== -1 && row[strengthIdx] ? row[strengthIdx].toString().trim() : "";
            const unitVal = unitIdx !== -1 && row[unitIdx] ? row[unitIdx].toString().trim() : "";
            const strength = (specVal && unitVal) ? `${{specVal}} (${{unitVal}})` : (specVal || unitVal);
            const dosageForm = dosageIdx !== -1 && row[dosageIdx] ? row[dosageIdx].toString().trim() : "";
            const nhiCode = (nhiCodeIdx !== -1 && row[nhiCodeIdx] ? row[nhiCodeIdx].toString().trim() : "").toUpperCase();
            const price = priceIdx !== -1 && row[priceIdx] ? row[priceIdx].toString().trim() : "";
            const ruleSection = ruleIdx !== -1 && row[ruleIdx] ? row[ruleIdx].toString().trim() : "";
            const ruleLink = ruleLinkIdx !== -1 && row[ruleLinkIdx] ? row[ruleLinkIdx].toString().trim() : "";
            const chineseName = chineseIdx !== -1 && row[chineseIdx] ? row[chineseIdx].toString().trim() : "";
            const pregnancy = pregIdx !== -1 && row[pregIdx] ? row[pregIdx].toString().trim() : "";
            const dosage = doseIdx !== -1 && row[doseIdx] ? row[doseIdx].toString().trim() : "";
            const ci = ciIdx !== -1 && row[ciIdx] ? row[ciIdx].toString().trim() : "";
            const ae = aeIdx !== -1 && row[aeIdx] ? row[aeIdx].toString().trim() : "";
            const note = noteIdx !== -1 && row[noteIdx] ? row[noteIdx].toString().trim() : "";
            const atc = (atcIdx !== -1 && row[atcIdx] ? row[atcIdx].toString().trim() : "").toUpperCase();

            if (!hospitalCode && !brandName && !genericName) continue;

            parsedItems.push({{
              id: "d_q2_" + Date.now() + "_" + i,
              seq: i,
              hospitalCode: hospitalCode || `TEMP_${{i}}`,
              brandName: brandName || "未填寫品名",
              chineseName,
              genericName: genericName || brandName,
              strength,
              dosageForm,
              nhiCode,
              price,
              ruleSection,
              ruleLink,
              pregnancy,
              dosage,
              ci,
              ae,
              note,
              atc,
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
    print("SUCCESS: index.html compiled with full 115Q2 Formulary integration (738 items)!")

if __name__ == '__main__':
    build()
