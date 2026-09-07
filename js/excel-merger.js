/**
 * ExcelMerger - 지능형 엑셀 규격 감지, 정제 및 병합 엔진
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
      const sheetSelection = options.sheetSelection || 'first';

      const parsedFiles = [];

      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        try {
          const buffer = await file.arrayBuffer();
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
                id: 'file_' + Math.random().toString(36).substr(2, 9),
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
                  cellVal = ExcelMerger.formatDate(cellVal);
                }
                if (cellVal !== undefined && cellVal !== null) {
                  if (typeof cellVal === 'string' && options.trimWhitespace !== false) {
                    cellVal = cellVal.trim();
                  }

                  // 고급 정제: 날짜 표준화
                  if (options.standardizeDate && typeof cellVal === 'string') {
                    cellVal = ExcelMerger.normalizeDateString(cellVal);
                  }

                  // 고급 정제: 전화번호 하이픈 표준화
                  if (options.autoFormatPhone && typeof cellVal === 'string') {
                    cellVal = ExcelMerger.formatPhoneNumber(cellVal);
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
     * 전화번호 문자열을 010-XXXX-XXXX 형식으로 통일
     */
    formatPhoneNumber(str) {
      if (!str || typeof str !== 'string') return str;
      const cleaned = str.replace(/[^\d]/g, '');
      if (cleaned.length === 11 && cleaned.startsWith('01')) {
        return `${cleaned.slice(0, 3)}-${cleaned.slice(3, 7)}-${cleaned.slice(7)}`;
      } else if (cleaned.length === 10 && cleaned.startsWith('01')) {
        return `${cleaned.slice(0, 3)}-${cleaned.slice(3, 6)}-${cleaned.slice(6)}`;
      } else if (cleaned.length === 9 || cleaned.length === 10) { // 일반 지역번호
        if (cleaned.startsWith('02')) {
          return cleaned.length === 9
            ? `${cleaned.slice(0, 2)}-${cleaned.slice(2, 5)}-${cleaned.slice(5)}`
            : `${cleaned.slice(0, 2)}-${cleaned.slice(2, 6)}-${cleaned.slice(6)}`;
        }
      }
      return str;
    },

    /**
     * 다양한 날짜 형식을 YYYY-MM-DD로 표준화
     */
    normalizeDateString(str) {
      if (!str || typeof str !== 'string') return str;
      // YYYY.MM.DD or YYYY/MM/DD or YYYYMMDD
      const dateMatch = str.match(/^(\d{4})[./-](\d{1,2})[./-](\d{1,2})/);
      if (dateMatch) {
        const y = dateMatch[1];
        const m = dateMatch[2].padStart(2, '0');
        const d = dateMatch[3].padStart(2, '0');
        return `${y}-${m}-${d}`;
      }
      if (/^\d{8}$/.test(str)) {
        return `${str.slice(0, 4)}-${str.slice(4, 6)}-${str.slice(6, 8)}`;
      }
      return str;
    },

    /**
     * 날짜 객체를 YYYY-MM-DD HH:mm:ss 문자열로 포맷팅
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
     */
    groupBySchema(parsedFiles, options = {}) {
      const matchMode = options.matchMode || 'exact';
      const groupsMap = new Map();

      parsedFiles.forEach(fileInfo => {
        if (fileInfo.error || !fileInfo.headers || fileInfo.headers.length === 0) {
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
          signature = fileInfo.normalizedHeaders.join(' ::: ');
        } else {
          const sortedCols = [...fileInfo.normalizedHeaders].sort();
          signature = sortedCols.join(' ::: ');
        }

        if (!groupsMap.has(signature)) {
          groupsMap.set(signature, {
            id: 'group_' + Math.random().toString(36).substr(2, 9),
            name: '',
            isError: false,
            signature,
            canonicalHeaders: [...fileInfo.headers],
            files: []
          });
        }

        groupsMap.get(signature).files.push(fileInfo);
      });

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

      const addSourceCol = options.addSourceColumn !== false;
      const removeDups = !!options.removeDuplicates;
      const dedupKeyCol = options.dedupKeyColumn || null; // 특정 컬럼 기준 중복 제거
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

          const alignedRow = {};

          if (addSourceCol) {
            alignedRow['_원본파일명'] = fileInfo.fileName;
          }

          canonicalHeaders.forEach(colName => {
            if (row[colName] !== undefined) {
              alignedRow[colName] = row[colName];
            } else {
              const targetNorm = colName.toLowerCase().trim();
              const matchedKey = Object.keys(row).find(k => k.toLowerCase().trim() === targetNorm);
              alignedRow[colName] = matchedKey ? row[matchedKey] : '';
            }
          });

          // 중복 검사
          if (removeDups) {
            let dataSig = '';
            if (dedupKeyCol && alignedRow[dedupKeyCol] !== undefined) {
              // 특정 키 컬럼 값 기준 중복 검사
              dataSig = `KEY:${dedupKeyCol}=${String(alignedRow[dedupKeyCol]).trim()}`;
            } else {
              // 전체 행 데이터 기준 중복 검사
              dataSig = canonicalHeaders.map(c => String(alignedRow[c] || '')).join('\u0000');
            }

            if (seenRowSignatures.has(dataSig)) {
              duplicateCount++;
              return;
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
     * 데이터 건전성(Health Check) 및 결측치 리포트 계산
     * @param {Object} mergedData { headers, rows }
     * @returns {Object} { totalCells, emptyCells, overallFillRate, columnsHealth: [] }
     */
    analyzeDataHealth(mergedData) {
      const { headers, rows } = mergedData;
      if (!headers || headers.length === 0 || !rows || rows.length === 0) {
        return {
          totalCells: 0,
          emptyCells: 0,
          overallFillRate: 100,
          columnsHealth: []
        };
      }

      // _원본파일명 제외한 실제 데이터 컬럼만 분석
      const dataHeaders = headers.filter(h => h !== '_원본파일명');
      let totalCells = dataHeaders.length * rows.length;
      let totalEmptyCells = 0;

      const columnsHealth = dataHeaders.map(col => {
        let emptyCount = 0;
        const uniqueVals = new Set();
        const sampleVals = [];
        let numCount = 0;
        let dateCount = 0;
        let phoneCount = 0;

        rows.forEach(r => {
          const val = r[col];
          if (val === undefined || val === null || String(val).trim() === '') {
            emptyCount++;
          } else {
            const str = String(val).trim();
            uniqueVals.add(str);
            if (sampleVals.length < 3 && !sampleVals.includes(str)) {
              sampleVals.push(str);
            }
            // 타입 추론
            if (!isNaN(Number(str))) numCount++;
            if (/^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}/.test(str)) dateCount++;
            if (/^01[0-9]-?\d{3,4}-?\d{4}$/.test(str)) phoneCount++;
          }
        });

        totalEmptyCells += emptyCount;
        const validCount = rows.length - emptyCount;
        const fillRate = rows.length > 0 ? Math.round((validCount / rows.length) * 100) : 0;

        let detectedType = '문자열 (Text)';
        if (numCount > rows.length * 0.7) detectedType = '숫자 (Number)';
        else if (dateCount > rows.length * 0.5) detectedType = '날짜 (Date)';
        else if (phoneCount > rows.length * 0.5) detectedType = '연락처 (Phone)';

        return {
          column: col,
          totalRows: rows.length,
          validCount,
          emptyCount,
          fillRate,
          uniqueCount: uniqueVals.size,
          detectedType,
          samples: sampleVals
        };
      });

      const overallFillRate = totalCells > 0 ? Math.round(((totalCells - totalEmptyCells) / totalCells) * 100) : 100;

      return {
        totalCells,
        emptyCells: totalEmptyCells,
        overallFillRate,
        columnsHealth
      };
    },

    /**
     * 서로 다른 규격의 그룹들을 사용자가 지정한 컬럼 매핑으로 강제 통합
     */
    mergeGroupsWithMapping(targetGroup, sourceGroup, columnMapping, options = {}) {
      // targetGroup 기준으로 sourceGroup의 컬럼들을 매핑하여 결합
      const targetMerged = ExcelMerger.mergeGroup(targetGroup, { ...options, addSourceColumn: false });
      const sourceMerged = ExcelMerger.mergeGroup(sourceGroup, { ...options, addSourceColumn: false });

      const finalHeaders = [...targetGroup.canonicalHeaders];
      const combinedRows = [...targetMerged.rows];

      sourceMerged.rows.forEach(sRow => {
        const mappedRow = {};
        finalHeaders.forEach(tCol => {
          const mappedSourceCol = columnMapping[tCol] || tCol;
          mappedRow[tCol] = sRow[mappedSourceCol] !== undefined ? sRow[mappedSourceCol] : '';
        });
        combinedRows.push(mappedRow);
      });

      return {
        name: `${targetGroup.name} + ${sourceGroup.name} (수동통합)`,
        canonicalHeaders: finalHeaders,
        files: [...targetGroup.files, ...sourceGroup.files],
        customMergedData: {
          headers: options.addSourceColumn !== false ? ['_원본파일명', ...finalHeaders] : finalHeaders,
          rows: combinedRows,
          totalOriginalRows: targetMerged.totalOriginalRows + sourceMerged.totalOriginalRows,
          totalMergedRows: combinedRows.length,
          duplicateCount: 0
        }
      };
    },

    /**
     * 병합된 데이터를 엑셀 워크북(XLSX)으로 변환
     */
    createWorkbook(mergedData, sheetTitle = '병합데이터') {
      const { headers, rows } = mergedData;
      const wsData = [headers];

      rows.forEach(row => {
        const rowArr = headers.map(h => (row[h] !== undefined ? row[h] : ''));
        wsData.push(rowArr);
      });

      const ws = XLSX.utils.aoa_to_sheet(wsData);

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

    calculateTextWidth(str) {
      let width = 0;
      for (let i = 0; i < str.length; i++) {
        const code = str.charCodeAt(i);
        width += (code > 127) ? 2 : 1;
      }
      return width;
    },

    downloadGroupAsXLSX(group, options = {}, filename = '') {
      const mergedData = group.customMergedData || ExcelMerger.mergeGroup(group, options);
      const defaultName = `${group.name || '병합결과'}_${group.files.length}개파일_통합.xlsx`;
      const finalFileName = filename || defaultName;

      const wb = ExcelMerger.createWorkbook(mergedData, group.name || '통합시트');
      XLSX.writeFile(wb, finalFileName);
    },

    downloadGroupAsCSV(group, options = {}, filename = '') {
      const mergedData = group.customMergedData || ExcelMerger.mergeGroup(group, options);
      const defaultName = `${group.name || '병합결과'}_${group.files.length}개파일_통합.csv`;
      const finalFileName = filename || defaultName;

      const { headers, rows } = mergedData;
      let csvContent = '\uFEFF';

      csvContent += headers.map(h => `"${String(h).replace(/"/g, '""')}"`).join(',') + '\r\n';

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
        const mergedData = group.customMergedData || ExcelMerger.mergeGroup(group, options);
        totalMergedFilesCount += group.files.length;
        totalMergedRowsCount += mergedData.totalMergedRows;

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

    createExcelFile(headers, rows, sheetName, fileName) {
      const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      return new File([blob], fileName, { type: blob.type });
    },

    /**
     * 3가지 업무별 실무 샘플 프리셋 생성
     */
    generatePresets(presetType = 'order') {
      if (presetType === 'inventory') {
        // 물류 재고 프리셋 (2개 파일)
        const headers = ['품목코드', '품목명', '카테고리', '입고수량', '보관창고', '입고일자'];
        const data1 = [
          ['IT-1001', '무선 블루투스 마우스', '디지털/가전', 150, '제1물류센터', '2026-03-01'],
          ['IT-1002', '기계식 게이밍 키보드', '디지털/가전', 80, '제1물류센터', '2026-03-02'],
          ['OF-2005', '인체공학 메쉬 체어', '가구/오피스', 45, '제2물류센터', '2026-03-03']
        ];
        const data2 = [
          ['OF-2008', '높이조절 전동 데스크', '가구/오피스', 20, '제2물류센터', '2026-03-04'],
          ['IT-1009', 'USB-C 7in1 멀티허브', '디지털/가전', 200, '제1물류센터', '2026-03-05'],
          ['LI-3011', '무드등 무선충전기', '인테리어/소품', 95, '제3물류센터', '2026-03-05']
        ];
        return [
          ExcelMerger.createExcelFile(headers, data1, '재고현황_1센터', '물류입고_제1센터.xlsx'),
          ExcelMerger.createExcelFile(headers, data2, '재고현황_2센터', '물류입고_제2센터.xlsx')
        ];
      } else if (presetType === 'survey') {
        // 고객 설문조사 프리셋 (헤더명이 달라 강제 통합/매핑 테스트하기 좋은 2개 파일)
        const headers1 = ['고객명', '연락처', '만족도점수', '추천의향', '의견'];
        const data1 = [
          ['김철수', '01011112222', 5, '적극추천', '배송이 매우 빨랐습니다.'],
          ['이영희', '01022223333', 4, '추천', '품질이 우수합니다.']
        ];
        const headers2 = ['성함', '핸드폰번호', '만족도점수', '추천의향', '의견'];
        const data2 = [
          ['박민호', '010-3333-4444', 5, '적극추천', '상담원이 친절합니다.'],
          ['정수진', '01044445555', 3, '보통', '포장 상태가 아쉬웠습니다.']
        ];
        return [
          ExcelMerger.createExcelFile(headers1, data1, '1차설문', '고객만족도_1차설문.xlsx'),
          ExcelMerger.createExcelFile(headers2, data2, '2차설문', '고객만족도_2차설문.xlsx')
        ];
      } else {
        // 기본 쇼핑몰 주문 및 고객 주문 프리셋 (2개 규격, 총 4개 파일)
        return ExcelMerger.generateSampleFiles();
      }
    },

    generateSampleFiles() {
      const schema1Headers = ['고객명', '연락처', '배송지주소', '주문금액(원)', '결제수단', '주문일시'];
      const data1_1 = [
        ['김민수', '01012345678', '서울특별시 강남구 테헤란로 123', 89000, '신용카드', '20260301'],
        ['이서연', '01023456789', '경기도 성남시 분당구 판교역로 45', 125000, '네이버페이', '2026.03.02'],
        ['박지훈', '01034567890', '부산광역시 해운대구 센텀중앙로 78', 45000, '계좌이체', '2026/03/03']
      ];
      const data1_2 = [
        ['최수빈', '010-4567-8901', '인천광역시 연수구 송도동 99', 67000, '카카오페이', '2026-03-04'],
        ['정다은', '010-5678-9012', '대전광역시 유성구 대학로 10', 210000, '신용카드', '2026-03-05'],
        ['김민수', '01012345678', '서울특별시 강남구 테헤란로 123', 89000, '신용카드', '20260301']
      ];

      const schema2Headers = ['품목코드', '품목명', '카테고리', '입고수량', '보관창고'];
      const data2_1 = [
        ['IT-1001', '무선 블루투스 마우스', '디지털/가전', 150, '제1물류센터'],
        ['IT-1002', '기계식 게이밍 키보드', '디지털/가전', 80, '제1물류센터'],
        ['OF-2005', '인체공학 메쉬 체어', '가구/오피스', 45, '제2물류센터']
      ];
      const data2_2 = [
        ['OF-2008', '높이조절 전동 데스크', '가구/오피스', 20, '제2물류센터'],
        ['IT-1009', 'USB-C 7in1 멀티허브', '디지털/가전', 200, '제1물류센터']
      ];

      return [
        ExcelMerger.createExcelFile(schema1Headers, data1_1, '1월주문', '쇼핑몰주문_1월.xlsx'),
        ExcelMerger.createExcelFile(schema1Headers, data1_2, '2월주문', '쇼핑몰주문_2월.xlsx'),
        ExcelMerger.createExcelFile(schema2Headers, data2_1, '입고현황_A', '물류입고_A센터.xlsx'),
        ExcelMerger.createExcelFile(schema2Headers, data2_2, '입고현황_B', '물류입고_B센터.xlsx')
      ];
    }
  };

  window.ExcelMerger = ExcelMerger;
})(window);
