import XlsxPopulate, { RichText } from "xlsx-populate";

/**
 * 인덱스에 해당되는 알파벳 가져오기
 * @param {number} index
 * @returns {string} Alphabet A~Z
 */
export const getColumnAlphabet = (index) => {
  const remain = index % 26;
  const code = String.fromCharCode(65 + remain);
  const result = index - remain;

  if (result > 0) {
    return `${getColumnAlphabet((result - 1) / 26)}${code}`;
  } else {
    return code;
  }
};

/**
 * 파일 다운로드
 * @param {Blob} blob
 * @param {string} filename
 */
export const fileDownload = (blob, filename) => {
  if (window.navigator && window.navigator.msSaveOrOpenBlob) {
    // If IE, you must uses a different method.
    window.navigator.msSaveOrOpenBlob(blob, filename);
  } else {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    document.body.appendChild(a);
    a.href = url;
    a.download = filename;
    a.click();
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);
  }
};

/**
 * 셀 데이터 주입
 * @param {*} rows
 * @param {*} sheet
 * @param {number} startRowIndex
 */
function insertCellData(rows, sheet, startRowIndex = 0) {
  rows.forEach((columns, rowIndex) => {
    // column 순회
    columns.forEach((column, columnIndex) => {
      // 셀 삽입 데이터 변수 생성
      let insertData = null;
      // 컬럼 데이터가 배열 인지 여부(이 경우에는 richText 입력으로 간주)
      if (Array.isArray(column?.text)) {
        const richText = new RichText();
        // richText 의 스타일 속성 적용
        column.text.forEach(({ text, style }) => {
          richText.add(text, (style = style || {}));
        });
        // 셀 삽입 변수 변경
        insertData = richText;
      } else if (typeof column?.text === "string") {
        // 이 경우는 셀 삽입 변수가 문자열, 셀 삽입 변수 변경
        insertData = column.text;
      }

      // 셀 삽입 변수가 존재하는 경우에만 삽입
      if (insertData) {
        // 알파벳 좌표 계산
        const positionName = `${getColumnAlphabet(columnIndex)}${
          rowIndex + startRowIndex + 1
        }`;
        // 셀 생성
        const cell = sheet.cell(positionName);
        cell.value(insertData);

        if (column?.style) {
          Object.keys(column.style).forEach((key) => {
            cell.style(key, column.style[key]);
          });
        }
      }
    });
  });
}

/**
 * 컬럼 넓이 지정
 * @param {number[]} columnWidth
 * @param {*} sheet
 */
function setColumnWidth(columnWidth, sheet) {
  columnWidth.forEach((width, columnIndex) => {
    sheet.column(getColumnAlphabet(columnIndex)).width(width);
  });
}

/**
 * 셀 머지
 * @param {*} merges
 * @param {*} sheet
 */
function mergeCell(merges, sheet) {
  merges.forEach(({ start, end }) => {
    const startPositionName = `${getColumnAlphabet(start.columnIndex)}${
      start.rowIndex + 1
    }`;
    const endPositionName = `${getColumnAlphabet(end.columnIndex)}${
      end.rowIndex + 1
    }`;
    sheet.range(`${startPositionName}:${endPositionName}`).merged(true);
  });
}

/**
 * 엑셀 파일 생성
 * @param {*} rawData
 * @param {string} filename
 */
export async function createExcelFile(
  { columnWidth, merges, category, list },
  filename
) {
  const workbook = await XlsxPopulate.fromBlankAsync();

  // 워크북 생성
  const sheet = workbook.sheet("Sheet1");

  // 카테고리 셀 데이터 주입
  insertCellData(category, sheet, 0);

  // 목록 셀 데이터 주입
  insertCellData(list, sheet, category.length);

  // 컬럼 넓이 지정
  setColumnWidth(columnWidth, sheet);

  // 셀 머지
  mergeCell(merges, sheet);

  // blob 생성
  const blob = await workbook.outputAsync();

  // 파일 다운로드
  fileDownload(blob, filename);
}
