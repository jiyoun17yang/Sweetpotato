/**
 * App.js - UI 컨트롤러 및 인터랙션 핸들러 (v2.0)
 */

(function() {
  'use strict';

  // 애플리케이션 상태
  const state = {
    rawFiles: [],
    parsedFiles: [],
    groups: [],
    activeGroupId: null,
    activeMainTab: 'preview', // 'preview', 'health', 'history'
    previewSearch: '',
    currentPage: 1,
    pageSize: 25,
    options: {
      matchMode: 'exact',
      headerRowIndex: 0,
      dedupKeyColumn: '',
      addSourceColumn: true,
      removeDuplicates: false,
      autoFormatPhone: true,
      standardizeDate: true,
      trimWhitespace: true,
      skipEmptyRows: true
    },
    history: []
  };

  // DOM 캐싱
  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');
  const btnSelectFiles = document.getElementById('btnSelectFiles');
  const selectMatchMode = document.getElementById('selectMatchMode');
  const selectHeaderRow = document.getElementById('selectHeaderRow');
  const selectDedupKey = document.getElementById('selectDedupKey');
  const chkAddSource = document.getElementById('chkAddSource');
  const chkRemoveDups = document.getElementById('chkRemoveDups');
  const chkFormatPhone = document.getElementById('chkFormatPhone');
  const chkStandardizeDate = document.getElementById('chkStandardizeDate');
  const chkTrimText = document.getElementById('chkTrimText');
  
  const panelEmptyGuide = document.getElementById('panelEmptyGuide');
  const panelLoadedHealth = document.getElementById('panelLoadedHealth');
  const metricTotalCells = document.getElementById('metricTotalCells');
  const metricValidCells = document.getElementById('metricValidCells');
  const metricEmptyCells = document.getElementById('metricEmptyCells');
  const healthOverallBadge = document.getElementById('healthOverallBadge');
  const healthColumnsList = document.getElementById('healthColumnsList');

  const dashboardSection = document.getElementById('dashboardSection');
  const statTotalFiles = document.getElementById('statTotalFiles');
  const statTotalGroups = document.getElementById('statTotalGroups');
  const statTotalRows = document.getElementById('statTotalRows');
  const statDupRows = document.getElementById('statDupRows');
  const btnResetAll = document.getElementById('btnResetAll');
  const btnDownloadAllZip = document.getElementById('btnDownloadAllZip');
  const btnOpenMappingModal = document.getElementById('btnOpenMappingModal');
  const groupsCountBadge = document.getElementById('groupsCountBadge');
  const groupsGrid = document.getElementById('groupsGrid');

  const previewTabs = document.getElementById('previewTabs');
  const previewSearchInput = document.getElementById('previewSearchInput');
  const previewPageSize = document.getElementById('previewPageSize');
  const previewTableHead = document.getElementById('previewTableHead');
  const previewTableBody = document.getElementById('previewTableBody');
  const paginationInfo = document.getElementById('paginationInfo');
  const paginationPages = document.getElementById('paginationPages');
  const toastContainer = document.getElementById('toastContainer');

  const tabContentPreview = document.getElementById('tabContentPreview');
  const tabContentHealth = document.getElementById('tabContentHealth');
  const tabContentHistory = document.getElementById('tabContentHistory');
  const healthTableBody = document.getElementById('healthTableBody');
  const historyList = document.getElementById('historyList');

  // 모달 요소
  const modalSchemaMapping = document.getElementById('modalSchemaMapping');
  const btnCloseModal = document.getElementById('btnCloseModal');
  const btnCancelMapping = document.getElementById('btnCancelMapping');
  const btnApplyMapping = document.getElementById('btnApplyMapping');
  const selectTargetGroup = document.getElementById('selectTargetGroup');
  const selectSourceGroup = document.getElementById('selectSourceGroup');
  const mappingList = document.getElementById('mappingList');

  // 로컬 스토리지 히스토리 로드
  function loadHistory() {
    try {
      const saved = localStorage.getItem('excel_merger_history');
      if (saved) state.history = JSON.parse(saved);
    } catch (e) {
      state.history = [];
    }
  }

  function addHistoryItem(title, details) {
    const item = {
      id: Date.now(),
      title,
      details,
      timestamp: new Date().toLocaleString()
    };
    state.history.unshift(item);
    if (state.history.length > 20) state.history.pop();
    try {
      localStorage.setItem('excel_merger_history', JSON.stringify(state.history));
    } catch (e) {}
    renderHistoryTab();
  }

  // 초기 이벤트 리스너
  function initEvents() {
    loadHistory();

    btnSelectFiles.addEventListener('click', (e) => {
      e.stopPropagation();
      fileInput.click();
    });

    dropZone.addEventListener('click', () => {
      fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length > 0) {
        handleNewFiles(Array.from(e.target.files));
        fileInput.value = '';
      }
    });

    ['dragenter', 'dragover'].forEach(eventName => {
      dropZone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(eventName => {
      dropZone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('dragover');
      });
    });

    dropZone.addEventListener('drop', (e) => {
      const dt = e.dataTransfer;
      if (dt && dt.files && dt.files.length > 0) {
        handleNewFiles(Array.from(dt.files));
      }
    });

    // 샘플 프리셋 칩 버튼들
    document.querySelectorAll('.btn-preset').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const preset = btn.dataset.preset;
        loadSamplePreset(preset);
      });
    });

    // 옵션 변경 이벤트
    selectMatchMode.addEventListener('change', () => {
      state.options.matchMode = selectMatchMode.value;
      reGroupAndRender();
      showToast(`규격 판별 기준이 '${selectMatchMode.options[selectMatchMode.selectedIndex].text}'(으)로 변경되었습니다.`, 'info');
    });

    selectHeaderRow.addEventListener('change', async () => {
      state.options.headerRowIndex = parseInt(selectHeaderRow.value, 10);
      await reParseFiles();
      showToast(`헤더 기준 행이 ${state.options.headerRowIndex + 1}번째 행으로 변경되었습니다.`, 'info');
    });

    selectDedupKey.addEventListener('change', () => {
      state.options.dedupKeyColumn = selectDedupKey.value;
      renderAll();
      showToast(state.options.dedupKeyColumn ? `'${state.options.dedupKeyColumn}' 컬럼 기준으로 중복 행을 검사합니다.` : '전체 컬럼 내용 기준으로 중복 검사합니다.', 'info');
    });

    chkAddSource.addEventListener('change', () => {
      state.options.addSourceColumn = chkAddSource.checked;
      renderPreview();
    });

    chkRemoveDups.addEventListener('change', () => {
      state.options.removeDuplicates = chkRemoveDups.checked;
      renderAll();
      showToast(chkRemoveDups.checked ? '중복 행 자동 제거가 활성화되었습니다.' : '중복 행 제거가 해제되었습니다.', 'info');
    });

    chkFormatPhone.addEventListener('change', async () => {
      state.options.autoFormatPhone = chkFormatPhone.checked;
      await reParseFiles();
      showToast(chkFormatPhone.checked ? '전화번호 하이픈(-) 자동 통일이 적용되었습니다.' : '전화번호 자동 통일이 해제되었습니다.', 'info');
    });

    chkStandardizeDate.addEventListener('change', async () => {
      state.options.standardizeDate = chkStandardizeDate.checked;
      await reParseFiles();
      showToast(chkStandardizeDate.checked ? '날짜 포맷 표준화(YYYY-MM-DD)가 적용되었습니다.' : '날짜 표준화가 해제되었습니다.', 'info');
    });

    chkTrimText.addEventListener('change', async () => {
      state.options.trimWhitespace = chkTrimText.checked;
      await reParseFiles();
    });

    // 메인 탭 전환
    document.querySelectorAll('.main-tab-btn').forEach(tabBtn => {
      tabBtn.addEventListener('click', () => {
        document.querySelectorAll('.main-tab-btn').forEach(b => b.classList.remove('active'));
        tabBtn.classList.add('active');
        state.activeMainTab = tabBtn.dataset.tab;
        switchMainTab(state.activeMainTab);
      });
    });

    // 초기화 버튼
    btnResetAll.addEventListener('click', () => {
      if (confirm('업로드된 모든 파일과 분석 결과를 초기화하시겠습니까?')) {
        resetApp();
        showToast('모든 데이터가 초기화되었습니다.', 'info');
      }
    });

    // 전체 ZIP 일괄 다운로드
    btnDownloadAllZip.addEventListener('click', async () => {
      await downloadAllAsZip();
    });

    // 미리보기 검색 필터
    previewSearchInput.addEventListener('input', (e) => {
      state.previewSearch = e.target.value.trim().toLowerCase();
      state.currentPage = 1;
      renderPreviewTableOnly();
    });

    // 페이지 크기 변경
    previewPageSize.addEventListener('change', (e) => {
      state.pageSize = parseInt(e.target.value, 10);
      state.currentPage = 1;
      renderPreviewTableOnly();
    });

    // FAQ 아코디언 토글
    document.querySelectorAll('.faq-question').forEach(q => {
      q.addEventListener('click', () => {
        const item = q.closest('.faq-item');
        item.classList.toggle('open');
      });
    });

    // 수동 규격 통합 모달
    btnOpenMappingModal.addEventListener('click', () => {
      openMappingModal();
    });

    [btnCloseModal, btnCancelMapping].forEach(btn => {
      btn.addEventListener('click', () => {
        modalSchemaMapping.style.display = 'none';
      });
    });

    selectTargetGroup.addEventListener('change', updateMappingRows);
    selectSourceGroup.addEventListener('change', updateMappingRows);

    btnApplyMapping.addEventListener('click', () => {
      applySchemaMapping();
    });
  }

  // 파일 업로드 처리
  async function handleNewFiles(files) {
    const validExtensions = ['.xlsx', '.xls', '.csv'];
    const validFiles = files.filter(f => {
      const ext = f.name.substring(f.name.lastIndexOf('.')).toLowerCase();
      return validExtensions.includes(ext);
    });

    if (validFiles.length === 0) {
      showToast('엑셀(.xlsx, .xls) 또는 .csv 파일만 업로드할 수 있습니다.', 'error');
      return;
    }

    const existingNames = new Set(state.rawFiles.map(f => f.name));
    let addedCount = 0;

    validFiles.forEach(f => {
      if (!existingNames.has(f.name)) {
        state.rawFiles.push(f);
        addedCount++;
      }
    });

    if (addedCount === 0) {
      showToast('이미 목록에 동일한 파일이 추가되어 있습니다.', 'info');
      return;
    }

    showToast(`${addedCount}개 파일이 성공적으로 추가되었습니다.`, 'success');
    await reParseFiles();
  }

  // 샘플 프리셋 로드
  async function loadSamplePreset(presetType) {
    try {
      const sampleFiles = ExcelMerger.generatePresets(presetType);
      state.rawFiles = [...sampleFiles];
      const titles = {
        order: '쇼핑몰 주문 내역 (2규격 4개 파일)',
        inventory: '물류 재고 현황 (1규격 2개 파일)',
        survey: '고객 설문조사 (동의어 매핑 테스트용 2개 파일)'
      };
      showToast(`'${titles[presetType] || presetType}' 샘플이 로드되었습니다!`, 'success');
      await reParseFiles();
    } catch (err) {
      console.error(err);
      showToast('샘플을 불러오는 중 오류가 발생했습니다.', 'error');
    }
  }

  // 파일 재파싱
  async function reParseFiles() {
    if (state.rawFiles.length === 0) {
      resetApp();
      return;
    }

    try {
      state.parsedFiles = await ExcelMerger.readFiles(state.rawFiles, state.options);
      updateDedupKeyOptions();
      reGroupAndRender();

      // UI 상태 전환: 우측 패널 가이드 -> 데이터 건전성 대시보드로 전환
      panelEmptyGuide.style.display = 'none';
      panelLoadedHealth.style.display = 'flex';
      dashboardSection.style.display = 'block';
    } catch (err) {
      console.error(err);
      showToast('파일 분석 중 오류 발생: ' + err.message, 'error');
    }
  }

  // 중복 키 컬럼 셀렉트박스 갱신
  function updateDedupKeyOptions() {
    const currentVal = selectDedupKey.value;
    selectDedupKey.innerHTML = '<option value="">전체 컬럼 내용 기준</option>';

    const allHeaders = new Set();
    state.parsedFiles.forEach(f => {
      f.headers.forEach(h => allHeaders.add(h));
    });

    allHeaders.forEach(h => {
      const opt = document.createElement('option');
      opt.value = h;
      opt.textContent = `단일 컬럼: [${h}]`;
      if (h === currentVal) opt.selected = true;
      selectDedupKey.appendChild(opt);
    });
  }

  // 그룹화 및 렌더링
  function reGroupAndRender() {
    state.groups = ExcelMerger.groupBySchema(state.parsedFiles, state.options);

    const firstValid = state.groups.find(g => !g.isError && g.files.length > 0);
    state.activeGroupId = firstValid ? firstValid.id : null;
    state.currentPage = 1;
    state.previewSearch = '';
    if (previewSearchInput) previewSearchInput.value = '';

    renderAll();
  }

  // 전체 화면 렌더링
  function renderAll() {
    renderStats();
    renderHealthDashboard();
    renderGroupCards();
    renderPreview();
    renderHealthTab();
    renderHistoryTab();
  }

  // 통계 렌더링
  function renderStats() {
    const totalFiles = state.parsedFiles.length;
    const validGroups = state.groups.filter(g => !g.isError);

    let totalRows = 0;
    let totalDups = 0;
    validGroups.forEach(g => {
      const merged = g.customMergedData || ExcelMerger.mergeGroup(g, state.options);
      totalRows += merged.totalMergedRows;
      totalDups += merged.duplicateCount;
    });

    statTotalFiles.textContent = `${totalFiles}개`;
    statTotalGroups.textContent = `${validGroups.length}종류`;
    statTotalRows.textContent = `${totalRows.toLocaleString()}행`;
    statDupRows.textContent = `${totalDups.toLocaleString()}행`;
    groupsCountBadge.textContent = `${validGroups.length}개 규격`;
  }

  // 우측 데이터 건전성 대시보드 렌더링
  function renderHealthDashboard() {
    const activeGroup = state.groups.find(g => g.id === state.activeGroupId);
    if (!activeGroup || activeGroup.isError) return;

    const merged = activeGroup.customMergedData || ExcelMerger.mergeGroup(activeGroup, state.options);
    const health = ExcelMerger.analyzeDataHealth(merged);

    metricTotalCells.textContent = health.totalCells.toLocaleString();
    metricValidCells.textContent = (health.totalCells - health.emptyCells).toLocaleString();
    metricEmptyCells.textContent = health.emptyCells.toLocaleString();

    healthOverallBadge.textContent = `채움률 ${health.overallFillRate}%`;
    healthOverallBadge.className = `badge ${health.overallFillRate >= 90 ? 'badge-files' : 'badge-rows'}`;

    healthColumnsList.innerHTML = '';
    health.columnsHealth.forEach(col => {
      const fillClass = col.fillRate >= 90 ? 'fill-high' : col.fillRate >= 70 ? 'fill-mid' : 'fill-low';
      const item = document.createElement('div');
      item.className = 'health-col-item';
      item.innerHTML = `
        <div class="health-col-header">
          <span class="health-col-name">
            ${escapeHtml(col.column)}
            <span class="col-type-tag">${escapeHtml(col.detectedType)}</span>
          </span>
          <span class="health-col-rate">${col.fillRate}% (${col.validCount}/${col.totalRows}행)</span>
        </div>
        <div class="health-bar">
          <div class="health-fill ${fillClass}" style="width: ${col.fillRate}%;"></div>
        </div>
      `;
      healthColumnsList.appendChild(item);
    });
  }

  // 규격 그룹 카드 렌더링
  function renderGroupCards() {
    groupsGrid.innerHTML = '';
    const validGroups = state.groups.filter(g => !g.isError);

    if (validGroups.length === 0) {
      groupsGrid.innerHTML = `
        <div class="empty-state" style="grid-column: 1 / -1;">
          <div class="empty-state-icon">⚠️</div>
          <p>분석 가능한 유효한 엑셀 파일이 없습니다. 헤더 행 설정을 확인해 주세요.</p>
        </div>
      `;
      return;
    }

    const borderColors = ['border-accent-1', 'border-accent-2', 'border-accent-3', 'border-accent-4', 'border-accent-5'];

    validGroups.forEach((group, idx) => {
      const merged = group.customMergedData || ExcelMerger.mergeGroup(group, state.options);
      const colorClass = borderColors[idx % borderColors.length];

      const card = document.createElement('div');
      card.className = `schema-card ${colorClass}`;
      card.id = `card_${group.id}`;

      const colTagsHtml = group.canonicalHeaders.map(col => `<span class="col-tag" title="${escapeHtml(col)}">${escapeHtml(col)}</span>`).join('');

      const filesListHtml = group.files.map(f => {
        const sizeKb = (f.fileSize / 1024).toFixed(1);
        return `
          <li class="file-item" data-file-id="${f.id}">
            <span class="file-name-span" title="${escapeHtml(f.fileName)}">
              📄 ${escapeHtml(f.fileName)}
            </span>
            <span class="file-meta-span">
              <span>${f.rowCount}행</span>
              <span>${sizeKb}KB</span>
              <button type="button" class="btn btn-danger-ghost btn-remove-file" data-file-id="${f.id}" title="이 파일 제외">✕</button>
            </span>
          </li>
        `;
      }).join('');

      card.innerHTML = `
        <div class="schema-card-header">
          <span class="group-title">${escapeHtml(group.name)}</span>
          <div class="group-badges">
            <span class="badge badge-files">${group.files.length}개 파일</span>
            <span class="badge badge-rows">${merged.totalMergedRows}행</span>
          </div>
        </div>

        <div class="columns-box">
          <div class="columns-header">
            <span>헤더 규격 (${group.canonicalHeaders.length}개 컬럼)</span>
          </div>
          <div class="columns-tags">
            ${colTagsHtml}
          </div>
        </div>

        <div class="files-box">
          <div class="files-header">포함된 원본 파일 (${group.files.length}개)</div>
          <ul class="files-list">
            ${filesListHtml}
          </ul>
        </div>

        <div class="card-actions">
          <button type="button" class="btn btn-secondary btn-preview btn-preview-group" data-group-id="${group.id}">
            🔍 실시간 데이터 미리보기
          </button>
          <button type="button" class="btn btn-outline-emerald btn-download-xlsx" data-group-id="${group.id}">
            📥 엑셀 (.xlsx)
          </button>
          <button type="button" class="btn btn-outline-indigo btn-download-csv" data-group-id="${group.id}">
            📄 CSV 다운로드
          </button>
        </div>
      `;

      card.querySelectorAll('.btn-remove-file').forEach(btn => {
        btn.addEventListener('click', (e) => {
          e.stopPropagation();
          removeFileById(btn.dataset.fileId);
        });
      });

      card.querySelector('.btn-preview-group').addEventListener('click', () => {
        state.activeGroupId = group.id;
        state.currentPage = 1;
        renderPreview();
        renderHealthDashboard();
        renderHealthTab();
        document.querySelector('.preview-section').scrollIntoView({ behavior: 'smooth' });
      });

      card.querySelector('.btn-download-xlsx').addEventListener('click', () => {
        downloadSingleGroup(group, 'xlsx');
      });

      card.querySelector('.btn-download-csv').addEventListener('click', () => {
        downloadSingleGroup(group, 'csv');
      });

      groupsGrid.appendChild(card);
    });
  }

  // 파일 개별 삭제
  async function removeFileById(fileId) {
    const fileToRemove = state.parsedFiles.find(f => f.id === fileId);
    if (!fileToRemove) return;

    state.rawFiles = state.rawFiles.filter(f => f.name !== fileToRemove.fileName);
    showToast(`'${fileToRemove.fileName}' 파일이 제외되었습니다.`, 'info');

    if (state.rawFiles.length === 0) {
      resetApp();
    } else {
      await reParseFiles();
    }
  }

  // 메인 탭 전환 처리
  function switchMainTab(tabKey) {
    tabContentPreview.style.display = tabKey === 'preview' ? 'block' : 'none';
    tabContentHealth.style.display = tabKey === 'health' ? 'block' : 'none';
    tabContentHistory.style.display = tabKey === 'history' ? 'block' : 'none';

    if (tabKey === 'health') renderHealthTab();
    if (tabKey === 'history') renderHistoryTab();
  }

  // 미리보기 탭 및 테이블 렌더링
  function renderPreview() {
    renderPreviewTabs();
    renderPreviewTableOnly();
  }

  function renderPreviewTabs() {
    previewTabs.innerHTML = '';
    const validGroups = state.groups.filter(g => !g.isError);

    validGroups.forEach(group => {
      const tabBtn = document.createElement('button');
      tabBtn.type = 'button';
      tabBtn.className = `group-tab-btn ${group.id === state.activeGroupId ? 'active' : ''}`;
      tabBtn.textContent = `${group.name} (${group.files.length}개 파일)`;
      tabBtn.addEventListener('click', () => {
        state.activeGroupId = group.id;
        state.currentPage = 1;
        state.previewSearch = '';
        if (previewSearchInput) previewSearchInput.value = '';
        renderPreview();
        renderHealthDashboard();
        renderHealthTab();
      });
      previewTabs.appendChild(tabBtn);
    });
  }

  function renderPreviewTableOnly() {
    const activeGroup = state.groups.find(g => g.id === state.activeGroupId);

    if (!activeGroup || activeGroup.isError) {
      previewTableHead.innerHTML = '';
      previewTableBody.innerHTML = `<tr><td colspan="10" class="empty-state"><p>선택된 규격 그룹이 없습니다.</p></td></tr>`;
      paginationInfo.textContent = '데이터 0건';
      paginationPages.innerHTML = '';
      return;
    }

    const mergedData = activeGroup.customMergedData || ExcelMerger.mergeGroup(activeGroup, state.options);
    const { headers, rows } = mergedData;

    let filteredRows = rows;
    if (state.previewSearch) {
      const q = state.previewSearch;
      filteredRows = rows.filter(row => {
        return headers.some(h => String(row[h] || '').toLowerCase().includes(q));
      });
    }

    previewTableHead.innerHTML = '';
    const trHead = document.createElement('tr');
    headers.forEach(h => {
      const th = document.createElement('th');
      th.textContent = h;
      if (h === '_원본파일명') th.className = 'col-source';
      trHead.appendChild(th);
    });
    previewTableHead.appendChild(trHead);

    const totalCount = filteredRows.length;
    const totalPages = Math.max(1, Math.ceil(totalCount / state.pageSize));
    if (state.currentPage > totalPages) state.currentPage = totalPages;

    const startIdx = (state.currentPage - 1) * state.pageSize;
    const endIdx = Math.min(startIdx + state.pageSize, totalCount);
    const pagedRows = filteredRows.slice(startIdx, endIdx);

    previewTableBody.innerHTML = '';

    if (pagedRows.length === 0) {
      const emptyRow = document.createElement('tr');
      const emptyTd = document.createElement('td');
      emptyTd.colSpan = headers.length || 1;
      emptyTd.className = 'empty-state';
      emptyTd.textContent = state.previewSearch ? '검색 결과와 일치하는 데이터가 없습니다.' : '표시할 데이터가 없습니다.';
      emptyRow.appendChild(emptyTd);
      previewTableBody.appendChild(emptyRow);
    } else {
      pagedRows.forEach(row => {
        const tr = document.createElement('tr');
        headers.forEach(h => {
          const td = document.createElement('td');
          const val = row[h] !== undefined ? row[h] : '';
          
          if (h === '_원본파일명') {
            td.innerHTML = `<span class="source-badge" title="${escapeHtml(val)}">${escapeHtml(val)}</span>`;
          } else {
            td.textContent = val;
            td.title = val;
          }
          tr.appendChild(td);
        });
        previewTableBody.appendChild(tr);
      });
    }

    if (totalCount === 0) {
      paginationInfo.textContent = '데이터 0건';
    } else {
      paginationInfo.textContent = `데이터 총 ${totalCount.toLocaleString()}건 중 ${startIdx + 1} - ${endIdx} 표시`;
    }

    renderPaginationButtons(totalPages);
  }

  function renderPaginationButtons(totalPages) {
    paginationPages.innerHTML = '';
    if (totalPages <= 1) return;

    const prevBtn = document.createElement('button');
    prevBtn.type = 'button';
    prevBtn.className = 'page-btn';
    prevBtn.innerHTML = '‹';
    prevBtn.disabled = state.currentPage <= 1;
    prevBtn.addEventListener('click', () => {
      if (state.currentPage > 1) {
        state.currentPage--;
        renderPreviewTableOnly();
      }
    });
    paginationPages.appendChild(prevBtn);

    let startPage = Math.max(1, state.currentPage - 2);
    let endPage = Math.min(totalPages, startPage + 4);
    if (endPage - startPage < 4) {
      startPage = Math.max(1, endPage - 4);
    }

    for (let p = startPage; p <= endPage; p++) {
      const pageBtn = document.createElement('button');
      pageBtn.type = 'button';
      pageBtn.className = `page-btn ${p === state.currentPage ? 'active' : ''}`;
      pageBtn.textContent = p;
      pageBtn.addEventListener('click', () => {
        state.currentPage = p;
        renderPreviewTableOnly();
      });
      paginationPages.appendChild(pageBtn);
    }

    const nextBtn = document.createElement('button');
    nextBtn.type = 'button';
    nextBtn.className = 'page-btn';
    nextBtn.innerHTML = '›';
    nextBtn.disabled = state.currentPage >= totalPages;
    nextBtn.addEventListener('click', () => {
      if (state.currentPage < totalPages) {
        state.currentPage++;
        renderPreviewTableOnly();
      }
    });
    paginationPages.appendChild(nextBtn);
  }

  // 탭 2: 컬럼별 상세 분석 리포트 렌더링
  function renderHealthTab() {
    const activeGroup = state.groups.find(g => g.id === state.activeGroupId);
    if (!activeGroup || activeGroup.isError) {
      healthTableBody.innerHTML = `<tr><td colspan="8" class="empty-state">선택된 규격이 없습니다.</td></tr>`;
      return;
    }

    const merged = activeGroup.customMergedData || ExcelMerger.mergeGroup(activeGroup, state.options);
    const health = ExcelMerger.analyzeDataHealth(merged);

    healthTableBody.innerHTML = '';
    health.columnsHealth.forEach(col => {
      const tr = document.createElement('tr');
      const samplesText = col.samples.length > 0 ? col.samples.join(', ') : '-';
      tr.innerHTML = `
        <td style="font-weight: 600; color: #ffffff;">${escapeHtml(col.column)}</td>
        <td><span class="col-type-tag">${escapeHtml(col.detectedType)}</span></td>
        <td>${col.totalRows.toLocaleString()}</td>
        <td style="color: #34d399;">${col.validCount.toLocaleString()}</td>
        <td style="color: ${col.emptyCount > 0 ? '#f43f5e' : '#94a3b8'};">${col.emptyCount.toLocaleString()}</td>
        <td style="font-weight: 700; color: ${col.fillRate >= 90 ? '#34d399' : '#f59e0b'};">${col.fillRate}%</td>
        <td>${col.uniqueCount.toLocaleString()}개</td>
        <td style="color: #94a3b8; max-width: 250px; overflow: hidden; text-overflow: ellipsis;" title="${escapeHtml(samplesText)}">${escapeHtml(samplesText)}</td>
      `;
      healthTableBody.appendChild(tr);
    });
  }

  // 탭 3: 히스토리 렌더링
  function renderHistoryTab() {
    historyList.innerHTML = '';
    if (state.history.length === 0) {
      historyList.innerHTML = `
        <div class="empty-state">
          <div class="empty-state-icon">🕒</div>
          <p>아직 다운로드된 병합 작업 기록이 없습니다.</p>
        </div>
      `;
      return;
    }

    state.history.forEach(item => {
      const div = document.createElement('div');
      div.className = 'history-item';
      div.innerHTML = `
        <div>
          <div class="history-title">📄 ${escapeHtml(item.title)}</div>
          <div class="history-meta">${escapeHtml(item.details)}</div>
        </div>
        <div style="font-size: 0.75rem; color: var(--text-muted);">${escapeHtml(item.timestamp)}</div>
      `;
      historyList.appendChild(div);
    });
  }

  // 단일 규격 다운로드
  function downloadSingleGroup(group, format = 'xlsx') {
    try {
      showToast(`${group.name} 다운로드를 준비 중입니다...`, 'info');
      const filename = `${group.name}_통합_${group.files.length}개파일.${format}`;
      if (format === 'xlsx') {
        ExcelMerger.downloadGroupAsXLSX(group, state.options, filename);
      } else {
        ExcelMerger.downloadGroupAsCSV(group, state.options, filename);
      }
      addHistoryItem(filename, `${group.name} (${group.files.length}개 파일 통합)`);
      showToast(`${group.name} 다운로드가 완료되었습니다.`, 'success');
    } catch (err) {
      console.error(err);
      showToast('다운로드 생성 중 오류 발생: ' + err.message, 'error');
    }
  }

  // 전체 ZIP 다운로드
  async function downloadAllAsZip() {
    const validGroups = state.groups.filter(g => !g.isError);
    if (validGroups.length === 0) {
      showToast('병합할 수 있는 규격 그룹이 없습니다.', 'error');
      return;
    }

    try {
      showToast('전체 규격 파일들을 ZIP으로 압축 중입니다...', 'info');
      const zipName = `엑셀_규격별_병합_${validGroups.length}개규격.zip`;
      await ExcelMerger.downloadAllAsZip(state.groups, state.options, zipName);
      addHistoryItem(zipName, `전체 ${validGroups.length}개 규격 일괄 압축 다운로드`);
      showToast('전체 규격 ZIP 일괄 다운로드가 완료되었습니다!', 'success');
    } catch (err) {
      console.error(err);
      showToast('ZIP 생성 중 오류 발생: ' + err.message, 'error');
    }
  }

  // 수동 규격 통합 모달 열기
  function openMappingModal() {
    const validGroups = state.groups.filter(g => !g.isError);
    if (validGroups.length < 2) {
      showToast('수동 통합을 하려면 서로 다른 규격 그룹이 최소 2개 이상 필요합니다.', 'error');
      return;
    }

    selectTargetGroup.innerHTML = '';
    selectSourceGroup.innerHTML = '';

    validGroups.forEach((g, i) => {
      const opt1 = document.createElement('option');
      opt1.value = g.id;
      opt1.textContent = `${g.name} (${g.canonicalHeaders.join(', ')})`;
      if (i === 0) opt1.selected = true;
      selectTargetGroup.appendChild(opt1);

      const opt2 = document.createElement('option');
      opt2.value = g.id;
      opt2.textContent = `${g.name} (${g.canonicalHeaders.join(', ')})`;
      if (i === 1) opt2.selected = true;
      selectSourceGroup.appendChild(opt2);
    });

    updateMappingRows();
    modalSchemaMapping.style.display = 'flex';
  }

  function updateMappingRows() {
    const targetId = selectTargetGroup.value;
    const sourceId = selectSourceGroup.value;

    if (targetId === sourceId) {
      mappingList.innerHTML = `<div style="color: #fb7185; font-size: 0.85rem; padding: 0.5rem;">서로 다른 두 개의 규격을 선택해 주세요.</div>`;
      return;
    }

    const targetGroup = state.groups.find(g => g.id === targetId);
    const sourceGroup = state.groups.find(g => g.id === sourceId);

    mappingList.innerHTML = '';
    targetGroup.canonicalHeaders.forEach(targetCol => {
      const row = document.createElement('div');
      row.className = 'mapping-row';

      let optionsHtml = `<option value="">(매핑 안함 - 빈값 처리)</option>`;
      sourceGroup.canonicalHeaders.forEach(sourceCol => {
        // 이름이 유사하거나 같으면 자동 매핑 선택
        const isMatched = targetCol.toLowerCase().trim() === sourceCol.toLowerCase().trim() ||
          (targetCol.includes('연락처') && sourceCol.includes('핸드폰')) ||
          (targetCol.includes('고객') && sourceCol.includes('성함'));
        optionsHtml += `<option value="${escapeHtml(sourceCol)}" ${isMatched ? 'selected' : ''}>${escapeHtml(sourceCol)}</option>`;
      });

      row.innerHTML = `
        <span style="font-weight: 600; color: #38bdf8;">[기준] ${escapeHtml(targetCol)}</span>
        <span style="color: var(--text-muted);">←</span>
        <select class="select-control mapping-select" data-target-col="${escapeHtml(targetCol)}">
          ${optionsHtml}
        </select>
      `;
      mappingList.appendChild(row);
    });
  }

  function applySchemaMapping() {
    const targetId = selectTargetGroup.value;
    const sourceId = selectSourceGroup.value;
    if (targetId === sourceId) {
      showToast('서로 다른 규격을 선택해 주세요.', 'error');
      return;
    }

    const targetGroup = state.groups.find(g => g.id === targetId);
    const sourceGroup = state.groups.find(g => g.id === sourceId);

    const mapping = {};
    document.querySelectorAll('.mapping-select').forEach(sel => {
      const targetCol = sel.dataset.targetCol;
      const sourceCol = sel.value;
      if (sourceCol) mapping[targetCol] = sourceCol;
    });

    const mergedCombinedGroup = ExcelMerger.mergeGroupsWithMapping(targetGroup, sourceGroup, mapping, state.options);
    mergedCombinedGroup.id = 'group_mapped_' + Math.random().toString(36).substr(2, 9);
    state.groups.push(mergedCombinedGroup);
    state.activeGroupId = mergedCombinedGroup.id;

    modalSchemaMapping.style.display = 'none';
    renderAll();
    showToast(`'${mergedCombinedGroup.name}' 통합 규격이 성공적으로 생성되었습니다!`, 'success');
  }

  // 초기화
  function resetApp() {
    state.rawFiles = [];
    state.parsedFiles = [];
    state.groups = [];
    state.activeGroupId = null;
    state.currentPage = 1;
    state.previewSearch = '';

    fileInput.value = '';
    panelEmptyGuide.style.display = 'flex';
    panelLoadedHealth.style.display = 'none';
    dashboardSection.style.display = 'none';
  }

  function escapeHtml(text) {
    if (text === undefined || text === null) return '';
    return String(text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;

    let icon = 'ℹ️';
    if (type === 'success') icon = '✅';
    if (type === 'error') icon = '⚠️';

    toast.innerHTML = `<span>${icon}</span><span>${escapeHtml(message)}</span>`;
    toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
      toast.style.transition = 'all 0.3s ease';
      setTimeout(() => {
        if (toast.parentNode) toast.parentNode.removeChild(toast);
      }, 300);
    }, 3200);
  }

  document.addEventListener('DOMContentLoaded', () => {
    initEvents();
  });
})();
