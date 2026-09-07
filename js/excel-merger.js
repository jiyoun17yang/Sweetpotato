/**
 * ExcelMerger - 지능형 엑셀 규격 감지 및 병합 엔진
 * 순수 JavaScript 및 SheetJS, JSZip 기반
 */

(function(window) {
  'use strict';

  const ExcelMerger = {
    /**
     * 파일 목록을 읽고 헤더 및 데이터를 추출
     * @param {File[]} files
     * @param {Object} options
     * @returns {Promise<Array>}
     */
    async readFiles(files, options = {}) {
      const headerRowIndex = options.headerRowIndex !== undefined ? options.headerRowIndex : 0;
      const sheetSelection = options.sheetSelection || 'first'; // 'first', 'all', or sheetName

      const parsedFiles = [];

      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        try {
          const buffer = await file.arrayBuffer();
          // cellDates: true, cellNF: false to properly preserve formatted dates/values
          const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });

          const sheetNames = workbook.SheetNames;
          if (!sheetNames || sheetNames.length === 0) {
            throw new Error('시트가 비어 있습니다.');
          }

          let sheetsToProcess = [];
          if (sheetSelection === 'first') {
            sheetsToProcess = [sheetNames[0]];
          } else if (sheetSelection === 'all') {
            sheetsToProcess = sheetNames;
          } else {
            sheetsToProcess = sheetNames.includes(sheetSelection) ? [sheetSelection] : [sheetNames[0]];
          }

          for (const sheetName of sheetsToProcess) {
            const worksheet = workbook.Sheets[sheetName];
            const rawData = XLSX.utils.sheet_to_json(worksheet, {
              header: 1,
              defval: '',
              blankrows: false
            });

            if (!rawData || rawData.length <= headerRowIndex) {
              parsedFiles.push({
                file,
                fileName: file.name,
                fileSize: file.size,
                sheetName,
                headers: [],
                normalizedHeaders: [],
                dataRows: [],
                rowCount: 0,
                error: '헤더 행을 찾을 수 없습니다.'
              });
              continue;
            }

            // 헤더 추출 및 정규화
            const rawHeaderRow = rawData[headerRowIndex] || [];
            // 마지막 빈 열 제거
            let lastValidCol = rawHeaderRow.length - 1;
            while (lastValidCol >= 0 && (rawHeaderRow[lastValidCol] === undefined || rawHeaderRow[lastValidCol] === null || String(rawHeaderRow[lastValidCol]).trim() === '')) {
              lastValidCol--;
            }

            const headers = [];
            const normalizedHeaders = [];

            for (let c = 0; c <= lastValidCol; c++) {
              const val = rawHeaderRow[c];
              const name = (val !== undefined && val !== null && String(val).trim() !== '')
                ? String(val).trim()
                : `열_${c + 1}`;
              headers.push(name);
              normalizedHeaders.push(name.toLowerCase().replace(/\s+/g, ' '));
            }

            // 데이터 행 추출 (헤더 다음 행부터)
            const rawDataRows = rawData.slice(headerRowIndex + 1);
            const dataRows = [];

            for (let r = 0; r < rawDataRows.length; r++) {
              const rowArr = rawDataRows[r] || [];
              const rowObj = {};
              let hasAnyValue = false;

              for (let c = 0; c < headers.length; c++) {
                let cellVal = rowArr[c];
                if (cellVal instanceof Date) {
                  // 날짜 형식 변환 (YYYY-MM-DD 또는 ISO 형태)
                  cellVal = ExcelMerger.formatDate(cellVal);
                }
                if (cellVal !== undefined && cellVal !== null) {
                  if (typeof cellVal === 'string' && options.trimWhitespace !== false) {
                    cellVal = cellVal.trim();
                  }
                  if (cellVal !== '') hasAnyValue = true;
                } else {
                  cellVal = '';
                }
                rowObj[headers[c]] = cellVal;
              }

              if (hasAnyValue || !options.skipEmptyRows) {
                dataRows.push(rowObj);
              }
            }

            parsedFiles.push({
              id: 'file_' + Math.random().toString(36).substr(2, 9),
              file,
              fileName: file.name,
              fileSize: file.size,
              sheetName,
              headers,
              normalizedHeaders,
              dataRows,
              rowCount: dataRows.length,
              error: null
            });
          }
        } catch (err) {
          console.error(`Error reading ${file.name}:`, err);
          parsedFiles.push({
            id: 'file_' + Math.random().toString(36).substr(2, 9),
            file,
            fileName: file.name,
            fileSize: file.size,
            sheetName: '',
            headers: [],
            normalizedHeaders: [],
            dataRows: [],
            rowCount: 0,
            error: err.message || '파일을 읽는 중 오류가 발생했습니다.'
          });
        }
      }

      return parsedFiles;
    },

    /**
     * 날짜 객체를 읽기 쉬운 문자열로 포맷팅
     */
    formatDate(date) {
      if (!(date instanceof Date) || isNaN(date.getTime())) return '';
      const pad = n => (n < 10 ? '0' + n : n);
      const y = date.getFullYear();
      const m = pad(date.getMonth() + 1);
      const d = pad(date.getDate());
      const h = date.getHours();
      const min = date.getMinutes();
      const s = date.getSeconds();

      if (h === 0 && min === 0 && s === 0) {
        return `${y}-${m}-${d}`;
      }
      return `${y}-${m}-${d} ${pad(h)}:${pad(min)}:${pad(s)}`;
    },

    /**
     * 추출된 파일들을 규격(헤더 스키마)별로 그룹화
     * @param {Array} parsedFiles
     * @param {Object} options { matchMode: 'exact' | 'flexible' }
     * @returns {Array} groups
     */
    groupBySchema(parsedFiles, options = {}) {
      const matchMode = options.matchMode || 'exact'; // 'exact' (순서 일치) or 'flexible' (순서 무관)
      const groupsMap = new Map();

      parsedFiles.forEach(fileInfo => {
        if (fileInfo.error || !fileInfo.headers || fileInfo.headers.length === 0) {
          // 오류 또는 빈 파일 그룹
          const errorKey = '__ERROR_GROUP__';
          if (!groupsMap.has(errorKey)) {
            groupsMap.set(errorKey, {
              id: 'group_err',
              name: '오류 또는 빈 파일',
              isError: true,
              signature: errorKey,
              canonicalHeaders: [],
              files: []
            });
          }
          groupsMap.get(errorKey).files.push(fileInfo);
          return;
        }

        let signature = '';
        if (matchMode === 'exact') {
          // 완전 일치: 순서 및 대소문자/정규화된 이름 일치
          signature = fileInfo.normalizedHeaders.join(' ::: ');
        } else {
          // 유연한 일치: 순서 무관 정렬하여 비교
          const sortedCols = [...fileInfo.normalizedHeaders].sort();
          signature = sortedCols.join(' ::: ');
        }

        if (!groupsMap.has(signature)) {
          groupsMap.set(signature, {
            id: 'group_' + Math.random().toString(36).substr(2, 9),
            name: '', // 이후 규격 #1, 규격 #2 로 네이밍
            isError: false,
            signature,
            canonicalHeaders: [...fileInfo.headers], // 기준 컬럼 순서
            files: []
          });
        }

        groupsMap.get(signature).files.push(fileInfo);
      });

      // 그룹 목록 정리 및 라벨링
      const groups = [];
      let validIndex = 1;

      groupsMap.forEach((group) => {
        if (group.isError) {
          groups.push(group);
        } else {
          group.name = `규격 #${validIndex}`;
          group.index = validIndex;
          validIndex++;
          groups.push(group);
        }
      });

      return groups;
    },

    /**
     * 특정 그룹의 파일 데이터를 병합
     * @param {Object} group
     * @param {Object} options { addSourceColumn: boolean, removeDuplicates: boolean, trimWhitespace: boolean }
     * @returns {Object} { headers, rows, totalOriginalRows, totalMergedRows, duplicateCount }
     */
    mergeGroup(group, options = {}) {
      if (group.isError || !group.canonicalHeaders || group.canonicalHeaders.length === 0) {
        return {
          headers: [],
          rows: [],
          totalOriginalRows: 0,
          totalMergedRows: 0,
          duplicateCount: 0
        };
      }

      const addSourceCol = options.addSourceColumn !== false; // 기본값 true
      const removeDups = !!options.removeDuplicates;
      const canonicalHeaders = [...group.canonicalHeaders];

      const finalHeaders = addSourceCol ? ['_원본파일명', ...canonicalHeaders] : [...canonicalHeaders];
      const mergedRows = [];
      let totalOriginalRows = 0;
      const seenRowSignatures = new Set();
      let duplicateCount = 0;

      group.files.forEach(fileInfo => {
        if (fileInfo.error) return;

        fileInfo.dataRows.forEach(row => {
          totalOriginalRows++;

          // 헤더 정렬 매핑 (유연한 일치일 때 컬럼 순서가 다를 수 있음)
          const alignedRow = {};

          if (addSourceCol) {
            alignedRow['_원본파일명'] = fileInfo.fileName;
          }

          canonicalHeaders.forEach(colName => {
            // 대소문자나 공백 차이가 있을 수 있으므로 매칭 검사
            if (row[colName] !== undefined) {
              alignedRow[colName] = row[colName];
            } else {
              // 정규화 비교
              const targetNorm = colName.toLowerCase().trim();
              const matchedKey = Object.keys(row).find(k => k.toLowerCase().trim() === targetNorm);
              alignedRow[colName] = matchedKey ? row[matchedKey] : '';
            }
          });

          // 중복 검사 (원본 파일명 제외한 실제 데이터 내용 기준)
          if (removeDups) {
            const dataSig = canonicalHeaders.map(c => String(alignedRow[c] || '')).join('\u0000');
            if (seenRowSignatures.has(dataSig)) {
              duplicateCount++;
              return; // 중복 행 제외
            }
            seenRowSignatures.add(dataSig);
          }

          mergedRows.push(alignedRow);
        });
      });

      return {
        headers: finalHeaders,
        rows: mergedRows,
        totalOriginalRows,
        totalMergedRows: mergedRows.length,
        duplicateCount
      };
    },

    /**
     * 병합된 데이터를 엑셀 워크북(XLSX) 바이너리로 변환
     * @param {Object} mergedData { headers, rows }
     * @param {string} sheetTitle
     * @returns {Object} XLSX.Workbook
     */
    createWorkbook(mergedData, sheetTitle = '병합데이터') {
      const { headers, rows } = mergedData;
      const wsData = [headers];

      rows.forEach(row => {
        const rowArr = headers.map(h => (row[h] !== undefined ? row[h] : ''));
        wsData.push(rowArr);
      });

      const ws = XLSX.utils.aoa_to_sheet(wsData);

      // 열 너비 자동 계산 (한글 및 영어 지원)
      const colWidths = headers.map((header, colIdx) => {
        let maxLen = ExcelMerger.calculateTextWidth(header);
        const sampleLimit = Math.min(wsData.length, 100);
        for (let r = 1; r < sampleLimit; r++) {
          const val = wsData[r][colIdx];
          if (val !== undefined && val !== null) {
            const len = ExcelMerger.calculateTextWidth(String(val));
            if (len > maxLen) maxLen = len;
          }
        }
        return { wch: Math.min(Math.max(maxLen + 3, 10), 60) };
      });
      ws['!cols'] = colWidths;

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, sheetTitle.substring(0, 31));
      return wb;
    },

    /**
     * 한글/영문 길이를 감안한 텍스트 너비 계산 (한글은 2자리 취급)
     */
    calculateTextWidth(str) {
      let width = 0;
      for (let i = 0; i < str.length; i++) {
        const code = str.charCodeAt(i);
        width += (code > 127) ? 2 : 1;
      }
      return width;
    },

    /**
     * 특정 그룹을 엑셀(.xlsx) 파일로 브라우저 다운로드
     */
    downloadGroupAsXLSX(group, options = {}, filename = '') {
      const mergedData = ExcelMerger.mergeGroup(group, options);
      const defaultName = `${group.name || '병합결과'}_${group.files.length}개파일_통합.xlsx`;
      const finalFileName = filename || defaultName;

      const wb = ExcelMerger.createWorkbook(mergedData, group.name || '통합시트');
      XLSX.writeFile(wb, finalFileName);
    },

    /**
     * 특정 그룹을 CSV 파일로 다운로드 (엑셀 한글 깨짐 방지 UTF-8 BOM 적용)
     */
    downloadGroupAsCSV(group, options = {}, filename = '') {
      const mergedData = ExcelMerger.mergeGroup(group, options);
      const defaultName = `${group.name || '병합결과'}_${group.files.length}개파일_통합.csv`;
      const finalFileName = filename || defaultName;

      const { headers, rows } = mergedData;
      let csvContent = '\uFEFF'; // UTF-8 BOM for Excel

      // 헤더 라인
      csvContent += headers.map(h => `"${String(h).replace(/"/g, '""')}"`).join(',') + '\r\n';

      // 데이터 라인
      rows.forEach(row => {
        const line = headers.map(h => {
          const val = row[h] !== undefined && row[h] !== null ? String(row[h]) : '';
          return `"${val.replace(/"/g, '""')}"`;
        }).join(',');
        csvContent += line + '\r\n';
      });

      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      ExcelMerger.saveBlob(blob, finalFileName);
    },

    /**
     * 모든 규격 그룹을 각각 엑셀 파일로 생성하여 하나의 ZIP 파일로 일괄 다운로드
     */
    async downloadAllAsZip(groups, options = {}, zipFileName = '엑셀_규격별_병합_일괄다운로드.zip') {
      if (typeof JSZip === 'undefined') {
        throw new Error('JSZip 라이브러리가 로드되지 않았습니다.');
      }

      const zip = new JSZip();
      let validGroupCount = 0;
      let totalMergedFilesCount = 0;
      let totalMergedRowsCount = 0;

      let summaryText = `========================================\r\n`;
      summaryText += `  동일 규격 엑셀 자동 병합 결과 요약\r\n`;
      summaryText += `  생성일시: ${new Date().toLocaleString()}\r\n`;
      summaryText += `========================================\r\n\r\n`;

      for (let i = 0; i < groups.length; i++) {
        const group = groups[i];
        if (group.isError || group.files.length === 0) continue;

        validGroupCount++;
        const mergedData = ExcelMerger.mergeGroup(group, options);
        totalMergedFilesCount += group.files.length;
        totalMergedRowsCount += mergedData.totalMergedRows;

        // 엑셀 바이너리 생성
        const wb = ExcelMerger.createWorkbook(mergedData, group.name);
        const wbOut = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

        const safeGroupName = (group.name || `규격_${i+1}`).replace(/[/\\?%*:|"<>]/g, '_');
        const excelName = `${safeGroupName}_${group.files.length}개파일_통합_${mergedData.totalMergedRows}행.xlsx`;

        zip.file(excelName, wbOut);

        summaryText += `[${group.name}]\r\n`;
        summaryText += ` - 파일명: ${excelName}\r\n`;
        summaryText += ` - 컬럼 (${group.canonicalHeaders.length}개): [${group.canonicalHeaders.join(', ')}]\r\n`;
        summaryText += ` - 원본 파일 (${group.files.length}개):\r\n`;
        group.files.forEach(f => {
          summaryText += `    * ${f.fileName} (${f.rowCount}행)\r\n`;
        });
        summaryText += ` - 병합 결과: 원본 총 ${mergedData.totalOriginalRows}행 -> 최종 ${mergedData.totalMergedRows}행 (중복제거: ${mergedData.duplicateCount}행)\r\n\r\n`;
      }

      if (validGroupCount === 0) {
        throw new Error('병합할 수 있는 유효한 규격 그룹이 없습니다.');
      }

      summaryText += `----------------------------------------\r\n`;
      summaryText += `총 ${validGroupCount}개 규격 그룹, 원본 ${totalMergedFilesCount}개 파일, 최종 ${totalMergedRowsCount}개 데이터 행 병합 완료.\r\n`;

      zip.file('00_병합_작업_요약_리포트.txt', summaryText);

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      ExcelMerger.saveBlob(zipBlob, zipFileName);
    },

    /**
     * Blob 브라우저 다운로드 트리거
     */
    saveBlob(blob, filename) {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }, 200);
    },

    /**
     * 테스트를 위한 가상 샘플 엑셀 파일들 생성 (규격 A 2개, 규격 B 2개)
     * 사용자가 별도 파일 없이도 1초 만에 바로 체험 가능!
     * @returns {File[]}
     */
    generateSampleFiles() {
      // 규격 1: 고객 및 주문 목록 (이름, 연락처, 배송지, 주문금액)
      const schema1Headers = ['고객명', '연락처', '배송지주소', '주문금액(원)', '결제수단'];
      const data1_1 = [
        ['김민수', '010-1234-5678', '서울특별시 강남구 테헤란로 123', 89000, '신용카드'],
        ['이서연', '010-2345-6789', '경기도 성남시 분당구 판교역로 45', 125000, '네이버페이'],
        ['박지훈', '010-3456-7890', '부산광역시 해운대구 센텀중앙로 78', 45000, '계좌이체']
      ];
      const data1_2 = [
        ['최수빈', '010-4567-8901', '인천광역시 연수구 송도동 99', 67000, '카카오페이'],
        ['정다은', '010-5678-9012', '대전광역시 유성구 대학로 10', 210000, '신용카드'],
        ['김민수', '010-1234-5678', '서울특별시 강남구 테헤란로 123', 89000, '신용카드'] // 의도적 중복
      ];

      // 규격 2: 물류 재고 및 품목 목록 (품목코드, 품목명, 카테고리, 입고수량, 보관창고)
      const schema2Headers = ['품목코드', '품목명', '카테고리', '입고수량', '보관창고'];
      const data2_1 = [
        ['IT-1001', '무선 블루투스 마우스', '디지털/가전', 150, '제1물류센터'],
        ['IT-1002', '기계식 게이밍 키보드', '디지털/가전', 80, '제1물류센터'],
        ['OF-2005', '인체공학 메쉬 체어', '가구/오피스', 45, '제2물류센터']
      ];
      const data2_2 = [
        ['OF-2008', '높이조절 전동 데스크', '가구/오피스', 20, '제2물류센터'],
        ['IT-1009', 'USB-C 7in1 멀티허브', '디지털/가전', 200, '제1물류센터'],
        ['LI-3011', '무드등 고속무선충전기', '인테리어/소품', 95, '제3물류센터']
      ];

      function createExcelBlob(headers, rows, sheetName) {
        const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        return new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      }

      const blob1 = createExcelBlob(schema1Headers, data1_1, '1월주문');
      const blob2 = createExcelBlob(schema1Headers, data1_2, '2월주문');
      const blob3 = createExcelBlob(schema2Headers, data2_1, '입고현황_A');
      const blob4 = createExcelBlob(schema2Headers, data2_2, '입고현황_B');

      return [
        new File([blob1], '고객주문내역_1월.xlsx', { type: blob1.type }),
        new File([blob2], '고객주문내역_2월.xlsx', { type: blob2.type }),
        new File([blob3], '물류입고_제1센터.xlsx', { type: blob3.type }),
        new File([blob4], '물류입고_제2센터.xlsx', { type: blob4.type })
      ];
    }
  };

  window.ExcelMerger = ExcelMerger;
})(window);
