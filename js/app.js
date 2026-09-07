/**
 * App.js - UI 컨트롤러 및 인터랙션 핸들러
 */

(function() {
  'use strict';

  // 애플리케이션 상태
  const state = {
    rawFiles: [],        // File[]
    parsedFiles: [],     // ExcelMerger parsed results
    groups: [],          // Schema groups
    activeGroupId: null, // Currently selected group for preview
    previewSearch: '',   // Search filter text
    currentPage: 1,      // Pagination page
    pageSize: 25,        // Rows per page
    options: {
      matchMode: 'exact',
      headerRowIndex: 0,
      addSourceColumn: true,
      removeDuplicates: false,
      trimWhitespace: true,
      skipEmptyRows: true
    }
  };

  // DOM 요소 캐싱
  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');
  const btnSelectFiles = document.getElementById('btnSelectFiles');
  const btnLoadSample = document.getElementById('btnLoadSample');
  const selectMatchMode = document.getElementById('selectMatchMode');
  const selectHeaderRow = document.getElementById('selectHeaderRow');
  const chkAddSource = document.getElementById('chkAddSource');
  const chkRemoveDups = document.getElementById('chkRemoveDups');
  const chkTrimText = document.getElementById('chkTrimText');
  const dashboardSection = document.getElementById('dashboardSection');
  const statTotalFiles = document.getElementById('statTotalFiles');
  const statTotalGroups = document.getElementById('statTotalGroups');
  const statTotalRows = document.getElementById('statTotalRows');
  const btnResetAll = document.getElementById('btnResetAll');
  const btnDownloadAllZip = document.getElementById('btnDownloadAllZip');
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

  // 이벤트 리스너 초기화
  function initEvents() {
    // 파일 선택 버튼 클릭
    btnSelectFiles.addEventListener('click', (e) => {
      e.stopPropagation();
      fileInput.click();
    });

    dropZone.addEventListener('click', () => {
      fileInput.click();
    });

    // 파일 입력 변경
    fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length > 0) {
        handleNewFiles(Array.from(e.target.files));
        fileInput.value = ''; // 동일 파일 재선택 가능하게 리셋
      }
    });

    // 드래그 앤 드롭
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

    // 샘플 데이터 즉시 체험 버튼
    btnLoadSample.addEventListener('click', (e) => {
      e.stopPropagation();
      loadSampleData();
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

    chkAddSource.addEventListener('change', () => {
      state.options.addSourceColumn = chkAddSource.checked;
      renderPreview();
    });

    chkRemoveDups.addEventListener('change', () => {
      state.options.removeDuplicates = chkRemoveDups.checked;
      renderAll();
      showToast(chkRemoveDups.checked ? '중복 행 자동 제거가 활성화되었습니다.' : '중복 행 자동 제거가 비활성화되었습니다.', 'info');
    });

    chkTrimText.addEventListener('change', async () => {
      state.options.trimWhitespace = chkTrimText.checked;
      await reParseFiles();
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

    // 미리보기 페이지 크기 변경
    previewPageSize.addEventListener('change', (e) => {
      state.pageSize = parseInt(e.target.value, 10);
      state.currentPage = 1;
      renderPreviewTableOnly();
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

    // 중복 파일명 방지 및 추가
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

    showToast(`${addedCount}개 파일이 성공적으로 추가되었습니다. 분석 중...`, 'success');
    await reParseFiles();
  }

  // 샘플 데이터 로드
  async function loadSampleData() {
    try {
      const sampleFiles = ExcelMerger.generateSampleFiles();
      state.rawFiles = [...sampleFiles];
      showToast('테스트용 4개 샘플 파일(2개 규격)이 로드되었습니다!', 'success');
      await reParseFiles();
    } catch (err) {
      console.error(err);
      showToast('샘플 파일을 생성하는 중 오류가 발생했습니다.', 'error');
    }
  }

  // 파일 재파싱
  async function reParseFiles() {
    if (state.rawFiles.length === 0) {
      dashboardSection.style.display = 'none';
      return;
    }

    try {
      state.parsedFiles = await ExcelMerger.readFiles(state.rawFiles, state.options);
      reGroupAndRender();
      dashboardSection.style.display = 'block';
    } catch (err) {
      console.error(err);
      showToast('파일 분석 중 오류가 발생했습니다: ' + err.message, 'error');
    }
  }

  // 그룹화 및 렌더링
  function reGroupAndRender() {
    state.groups = ExcelMerger.groupBySchema(state.parsedFiles, state.options);
    
    // 유효한 첫 번째 그룹을 기본 활성 탭으로 지정
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
    renderGroupCards();
    renderPreview();
  }

  // 대시보드 통계 렌더링
  function renderStats() {
    const totalFiles = state.parsedFiles.length;
    const validGroups = state.groups.filter(g => !g.isError);
    
    let totalRows = 0;
    validGroups.forEach(g => {
      const merged = ExcelMerger.mergeGroup(g, state.options);
      totalRows += merged.totalMergedRows;
    });

    statTotalFiles.textContent = `${totalFiles}개`;
    statTotalGroups.textContent = `${validGroups.length}종류`;
    statTotalRows.textContent = `${totalRows.toLocaleString()}행`;
    groupsCountBadge.textContent = `${validGroups.length}개 규격`;
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
      const merged = ExcelMerger.mergeGroup(group, state.options);
      const colorClass = borderColors[idx % borderColors.length];

      const card = document.createElement('div');
      card.className = `schema-card ${colorClass}`;
      card.id = `card_${group.id}`;

      // 컬럼 태그 목록 HTML
      const colTagsHtml = group.canonicalHeaders.map(col => `<span class="col-tag" title="${escapeHtml(col)}">${escapeHtml(col)}</span>`).join('');

      // 포함 파일 목록 HTML
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
          <div class="group-title-box">
            <span class="group-title">${escapeHtml(group.name)}</span>
          </div>
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
            🔍 병합 데이터 미리보기
          </button>
          <button type="button" class="btn btn-outline-emerald btn-download-xlsx" data-group-id="${group.id}">
            📥 엑셀 (.xlsx)
          </button>
          <button type="button" class="btn btn-outline-indigo btn-download-csv" data-group-id="${group.id}">
            📄 CSV 다운로드
          </button>
        </div>
      `;

      // 파일 개별 삭제 버튼 리스너
      card.querySelectorAll('.btn-remove-file').forEach(btn => {
        btn.addEventListener('click', (e) => {
          e.stopPropagation();
          const fileId = btn.dataset.fileId;
          removeFileById(fileId);
        });
      });

      // 미리보기 버튼 리스너
      card.querySelector('.btn-preview-group').addEventListener('click', () => {
        state.activeGroupId = group.id;
        state.currentPage = 1;
        renderPreview();
        // 부드러운 스크롤 이동
        document.querySelector('.preview-section').scrollIntoView({ behavior: 'smooth' });
      });

      // 개별 XLSX 다운로드 리스너
      card.querySelector('.btn-download-xlsx').addEventListener('click', () => {
        downloadSingleGroup(group, 'xlsx');
      });

      // 개별 CSV 다운로드 리스너
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
    showToast(`'${fileToRemove.fileName}' 파일이 목록에서 제외되었습니다.`, 'info');

    if (state.rawFiles.length === 0) {
      resetApp();
    } else {
      await reParseFiles();
    }
  }

  // 미리보기 탭 및 테이블 전체 렌더링
  function renderPreview() {
    renderPreviewTabs();
    renderPreviewTableOnly();
  }

  // 미리보기 탭 렌더링
  function renderPreviewTabs() {
    previewTabs.innerHTML = '';
    const validGroups = state.groups.filter(g => !g.isError);

    validGroups.forEach(group => {
      const tabBtn = document.createElement('button');
      tabBtn.type = 'button';
      tabBtn.className = `tab-btn ${group.id === state.activeGroupId ? 'active' : ''}`;
      tabBtn.textContent = `${group.name} (${group.files.length}개 파일)`;
      tabBtn.addEventListener('click', () => {
        state.activeGroupId = group.id;
        state.currentPage = 1;
        state.previewSearch = '';
        if (previewSearchInput) previewSearchInput.value = '';
        renderPreview();
      });
      previewTabs.appendChild(tabBtn);
    });
  }

  // 미리보기 테이블 및 페이지네이션 렌더링
  function renderPreviewTableOnly() {
    const activeGroup = state.groups.find(g => g.id === state.activeGroupId);

    if (!activeGroup || activeGroup.isError) {
      previewTableHead.innerHTML = '';
      previewTableBody.innerHTML = `
        <tr>
          <td colspan="10" class="empty-state">
            <div class="empty-state-icon">📋</div>
            <p>선택된 규격 그룹이 없습니다.</p>
          </td>
        </tr>
      `;
      paginationInfo.textContent = '데이터 0건';
      paginationPages.innerHTML = '';
      return;
    }

    const mergedData = ExcelMerger.mergeGroup(activeGroup, state.options);
    const { headers, rows } = mergedData;

    // 검색 필터 적용
    let filteredRows = rows;
    if (state.previewSearch) {
      const q = state.previewSearch;
      filteredRows = rows.filter(row => {
        return headers.some(h => String(row[h] || '').toLowerCase().includes(q));
      });
    }

    // 헤더 행 생성
    previewTableHead.innerHTML = '';
    const trHead = document.createElement('tr');
    headers.forEach(h => {
      const th = document.createElement('th');
      th.textContent = h;
      if (h === '_원본파일명') th.className = 'col-source';
      trHead.appendChild(th);
    });
    previewTableHead.appendChild(trHead);

    // 페이지네이션 계산
    const totalCount = filteredRows.length;
    const totalPages = Math.max(1, Math.ceil(totalCount / state.pageSize));
    if (state.currentPage > totalPages) state.currentPage = totalPages;

    const startIdx = (state.currentPage - 1) * state.pageSize;
    const endIdx = Math.min(startIdx + state.pageSize, totalCount);
    const pagedRows = filteredRows.slice(startIdx, endIdx);

    // 테이블 본문 생성
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

    // 페이지네이션 정보 갱신
    if (totalCount === 0) {
      paginationInfo.textContent = '데이터 0건';
    } else {
      paginationInfo.textContent = `데이터 총 ${totalCount.toLocaleString()}건 중 ${startIdx + 1} - ${endIdx} 표시`;
    }

    // 페이지 번호 버튼 렌더링
    renderPaginationButtons(totalPages);
  }

  // 페이지네이션 번호 버튼 렌더링
  function renderPaginationButtons(totalPages) {
    paginationPages.innerHTML = '';
    if (totalPages <= 1) return;

    // 이전 버튼
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

    // 표시할 페이지 번호 범위 계산 (최대 5개 버튼)
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

    // 다음 버튼
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

  // 단일 규격 다운로드
  function downloadSingleGroup(group, format = 'xlsx') {
    try {
      showToast(`${group.name} 통합 파일 다운로드를 준비 중입니다...`, 'info');
      if (format === 'xlsx') {
        ExcelMerger.downloadGroupAsXLSX(group, state.options);
      } else {
        ExcelMerger.downloadGroupAsCSV(group, state.options);
      }
      showToast(`${group.name} 다운로드가 완료되었습니다.`, 'success');
    } catch (err) {
      console.error(err);
      showToast('다운로드 생성 중 오류가 발생했습니다: ' + err.message, 'error');
    }
  }

  // 전체 ZIP 일괄 다운로드
  async function downloadAllAsZip() {
    const validGroups = state.groups.filter(g => !g.isError);
    if (validGroups.length === 0) {
      showToast('병합할 수 있는 규격 그룹이 없습니다.', 'error');
      return;
    }

    try {
      showToast('전체 규격 파일들을 ZIP으로 압축 중입니다...', 'info');
      await ExcelMerger.downloadAllAsZip(state.groups, state.options);
      showToast('전체 규격 ZIP 일괄 다운로드가 완료되었습니다!', 'success');
    } catch (err) {
      console.error(err);
      showToast('ZIP 생성 중 오류가 발생했습니다: ' + err.message, 'error');
    }
  }

  // 앱 리셋
  function resetApp() {
    state.rawFiles = [];
    state.parsedFiles = [];
    state.groups = [];
    state.activeGroupId = null;
    state.currentPage = 1;
    state.previewSearch = '';

    fileInput.value = '';
    dashboardSection.style.display = 'none';
  }

  // 유틸리티: HTML 이스케이프
  function escapeHtml(text) {
    if (text === undefined || text === null) return '';
    return String(text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  // 토스트 알림 표시
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

  // DOM 준비 시 초기화
  document.addEventListener('DOMContentLoaded', () => {
    initEvents();
  });
})();
