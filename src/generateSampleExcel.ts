import xlsx from "@sheet/core";
import type { Data, SplittedData } from "./generateData";
import generateMonthSheet from "./generateMonthSheet";
import generateSummarySheet from "./generateSummarySheet";

const generateSampleExcel = (data: Data[]) => {
  const startTime = new Date().getTime();

  const wb = xlsx.utils.book_new();

  const splittedData: SplittedData = {};

  // reduce 사용시 GC로 인한 퍼포먼스 저하가 예상되므로, for문을 사용합니다.
  for (const row of data) {
    const month = row.date.slice(0, 7);
    if (splittedData[month]) {
      splittedData[month] = [...splittedData[month], row];
    } else {
      splittedData[month] = [row];
    }
  }

  const sheetNames = Object.keys(splittedData).map((yearAndMonth) => {
    const [year, month] = yearAndMonth
      .split("-")
      .map((stringifiedNumber) => Number(stringifiedNumber));
    return `${year}년_${month}월_상세`;
  });
  const { salesSheet, buyingSheet } = generateSummarySheet(sheetNames);
  xlsx.utils.book_append_sheet(wb, salesSheet, "한눈에보기_매출");
  xlsx.utils.book_append_sheet(wb, buyingSheet, "한눈에보기_매입");

  Object.keys(splittedData).forEach((yearAndMonth) => {
    const [year, month] = yearAndMonth
      .split("-")
      .map((stringifiedNumber) => Number(stringifiedNumber));
    const { ws, sheetName } = generateMonthSheet(
      year,
      month,
      splittedData[yearAndMonth]
    );
    xlsx.utils.book_append_sheet(wb, ws, sheetName);
  });

  const sheetGeneratedTime = new Date().getTime();

  xlsx.writeFile(wb, "부가세_정산내역.xlsx", {
    cellStyles: true,
    bookSST: true,
  });

  const fileGeneratedTime = new Date().getTime();

  return {
    sheetGenerate: sheetGeneratedTime - startTime,
    fileGenerate: fileGeneratedTime - sheetGeneratedTime,
  };
};

export default generateSampleExcel;
