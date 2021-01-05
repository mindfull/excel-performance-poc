import xlsx, { Style } from "@sheet/core";
import { format, lastDayOfMonth } from "date-fns";
import type { Data } from "./generateData";

interface Column {
  key: keyof Data;
  label: string;
}

const columns: Column[] = [
  {
    key: "orderNumber",
    label: "부릉오더넘버",
  },
  {
    key: "date",
    label: "일자",
  },
  {
    key: "clientCharge",
    label: "정산상품가액",
  },
  {
    key: "deliveryFee",
    label: "정산배송비",
  },
  {
    key: "cancelFee",
    label: "정산취소수수료",
  },
  {
    key: "creditCardCharge",
    label: "카드결제",
  },
  {
    key: "creditCardFee",
    label: "카드수수료",
  },
  {
    key: "creditCardFeeRate",
    label: "카드수수료율(%)",
  },
  {
    key: "cashReceiptId",
    label: "현금영수증승인번호",
  },
  {
    key: "paymentType",
    label: "최종결제",
  },
];

const generateMonthSheet = (year: number, month: number, data: Data[]) => {
  const sheetName = `${year}년_${month}월_상세`;
  const aoaData = [
    ["", "", "", "", "", "", "", "", "", ""],
    ["부가세 신고자료", "", "", "", "", "", "", "", "", ""],
    ["(매출/매입)", "", "", "", "", "", "", "", "", ""],
    [
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      `집계기간: ${format(new Date(year, month - 1), "yyyy-MM-dd")} ~ ${format(
        lastDayOfMonth(new Date(year, month - 1)),
        "yyyy-MM-dd"
      )}`,
      "",
    ],
    [],
    ["", "", "", "", "", "", "", "", "", ""],
    ["상점명", "Lorem Ipsum", "", "", "", "", "", "", "", ""],
    ["사업자명", "김로렘", "", "", "", "", "", "", "", ""],
    ["사업자등록번호", "123-45-67890", "", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", "", "", ""],
    ["오더 수:", "", "", "", "", "", "", "", "", ""],
    columns.map(({ label }) => label),
    ...data.map((row) => columns.map(({ key }) => row[key])),
  ];

  const ws = xlsx.utils.aoa_to_sheet(aoaData);

  const NUMBER_FORMAT = "#,##0";

  // 기본 사이즈
  ws["!sheetFormat"] = {
    row: {
      hpx: 24,
    },
    col: {
      wpx: 96,
    },
  };
  ws["!rows"] = [];
  ws["!cols"] = [];

  ws["!rows"][1] = { hpx: 48 };
  ws["!rows"][4] = { hpx: 7 };
  ws["!rows"][5] = { hpx: 7 };

  ws["!cols"][0] = { wpx: 132 };
  ws["!cols"][8] = { wpx: 114 };

  const firstRow = 13;
  const lastRow = 13 + data.length;

  const DEFAULT_STYLE: Style = {
    sz: 10,
    alignment: {
      horizontal: "center",
      vertical: "center",
    },
  };
  xlsx.utils.sheet_set_range_style(ws, `A3:J${lastRow}`, DEFAULT_STYLE);

  // 제목 스타일링
  ws["!merges"] = [
    xlsx.utils.decode_range("A2:J2"),
    xlsx.utils.decode_range("A3:J3"),
    xlsx.utils.decode_range("I4:J4"),
    xlsx.utils.decode_range("B7:D7"),
    xlsx.utils.decode_range("B8:D8"),
    xlsx.utils.decode_range("B9:D9"),
  ];
  ws["A2"].s = {
    bold: true,
    sz: 24,
    alignment: {
      horizontal: "center",
      vertical: "center",
    },
  };
  ws["I4"].s = {
    sz: 9,
    alignment: {
      horizontal: "left",
    },
  };

  xlsx.utils.sheet_set_range_style(ws, "A5:J5", {
    fgColor: { rgb: 0x000000 },
  });

  xlsx.utils.sheet_set_range_style(ws, "A7:A9", {
    fgColor: { rgb: 0xd9d9d9 },
  });

  // 통계 공식 삽입
  const statisticsRow = 11;
  ws[`B${statisticsRow}`] = { t: "n", f: `COUNTA(B${firstRow}:B${lastRow})` };
  ws[`C${statisticsRow}`] = { t: "n", f: `SUM(C${firstRow}:C${lastRow})` };
  ws[`D${statisticsRow}`] = { t: "n", f: `SUM(D${firstRow}:D${lastRow})` };
  ws[`E${statisticsRow}`] = { t: "n", f: `SUM(E${firstRow}:E${lastRow})` };
  ws[`F${statisticsRow}`] = { t: "n", f: `SUM(F${firstRow}:F${lastRow})` };
  ws[`G${statisticsRow}`] = { t: "n", f: `SUM(G${firstRow}:G${lastRow})` };
  ws[`I${statisticsRow}`] = { t: "n", f: `COUNTA(I${firstRow}:I${lastRow})` };

  xlsx.utils.sheet_set_range_style(ws, `A${statisticsRow}:B${statisticsRow}`, {
    bold: true,
    sz: 15,
    alignment: {
      horizontal: "center",
      vertical: "center",
    },
  });
  xlsx.utils.sheet_set_range_style(ws, `C${statisticsRow}:J${statisticsRow}`, {
    bold: true,
    sz: 12,
    color: { rgb: 0xffff00 },
    fgColor: { rgb: 0x000000 },
    alignment: {
      horizontal: "center",
      vertical: "center",
    },
    z: NUMBER_FORMAT,
  });

  const headerRow = 12;
  xlsx.utils.sheet_set_range_style(ws, `A${headerRow}:J${headerRow}`, {
    bold: true,
    fgColor: { rgb: 0xd9d9d9 },
  });

  // 데이터 border 변경
  const WHITE_BORDER_STYLE = {
    top: { style: "thin", color: { rgb: 0xffffff } },
    right: { style: "thin", color: { rgb: 0xffffff } },
    bottom: { style: "thin", color: { rgb: 0xffffff } },
    left: { style: "thin", color: { rgb: 0xffffff } },
  };
  const DOTTED_STYLE = {
    top: { style: "dotted" },
    right: { style: "dotted" },
    bottom: { style: "dotted" },
    left: { style: "dotted" },
  };

  aoaData.forEach((row, rowIndex) => {
    row.forEach((_, colIndex) => {
      const cell = ws[xlsx.utils.encode_cell({ r: rowIndex, c: colIndex })];
      if (rowIndex < firstRow - 2) {
        // 7번째 줄에서 9번째 줄은 흰색 border가 없어야 한다.
        if (rowIndex >= 6 && rowIndex <= 8 && colIndex === 0) {
          return;
        }

        cell.s = cell.s
          ? {
              ...cell.s,
              ...WHITE_BORDER_STYLE,
            }
          : {
              ...WHITE_BORDER_STYLE,
            };

        return;
      }

      if (rowIndex < firstRow - 1) {
        return;
      }

      cell.s = cell.s
        ? {
            ...cell.s,
            ...DOTTED_STYLE,
          }
        : {
            ...DOTTED_STYLE,
          };
      cell.z = colIndex === 7 ? "0.00%" : NUMBER_FORMAT;
    });
  });

  return {
    sheetName,
    ws,
  };
};

export default generateMonthSheet;
