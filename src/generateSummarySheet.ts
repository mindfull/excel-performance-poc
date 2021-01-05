import xlsx, { Style } from "@sheet/core";

const emptyRow = ["", "", "", "", "", "", "", "", "", ""];

const RED_TEXT_COLOR = { rgb: 0xff0000 };
const HEIGHT_FOR_DIVIDER_ROW = { hpx: 14 };
const HEIGHT_FOR_HEADING_ROW = { hpx: 30 };
const HEIGHT_FOR_HEADER_ROW = { hpx: 26 };
const HEIGHT_FOR_SMALL_DIVIDER_ROW = { hpx: 10 };

const DEFAULT_STYLE: Style = {
  alignment: {
    vertical: "center",
  },
};
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
} as const;

const COUNT_FORMAT = "#,##0 건";
const PRICE_FORMAT = "#,##0 원";

const DATA_HEADING_ROW_STYLE: Style = {
  bold: true,
  sz: 10,
  color: { rgb: 0xffffff },
  fgColor: { rgb: 0x000000 },
  alignment: {
    horizontal: "center",
    vertical: "center",
  },
  ...DOTTED_STYLE,
  top: { style: "thin", color: { rgb: 0xffffff } },
};

const DATA_HEADING_COLUMN_STYLE: Style = {
  bold: true,
  sz: 10,
  alignment: {
    horizontal: "left",
    vertical: "center",
  },
  ...DOTTED_STYLE,
};

const DATA_YEAR_CELL_STYLE: Style = {
  ...DATA_HEADING_COLUMN_STYLE,
  alignment: {
    horizontal: "center",
    vertical: "center",
  },
};

const DATA_CELL_STYLE: Style = {
  sz: 10,
  alignment: {
    horizontal: "right",
    vertical: "center",
  },
  ...DOTTED_STYLE,
};

const ETC_DATA_HEADING_COLUMN_STYLE: Style = {
  ...DATA_HEADING_COLUMN_STYLE,
  fgColor: { rgb: 0xffff00 },
};

const ETC_DATA_CELL_STYLE: Style = {
  ...DATA_CELL_STYLE,
  fgColor: { rgb: 0xffff00 },
  // @ts-ignore - SheetJS에서 double 스타일 border가 정의에 누락됨
  bottom: { style: "double" },
};

const SUM_DATA_HEADING_COLUMN_STYLE: Style = {
  ...DATA_HEADING_COLUMN_STYLE,
  alignment: {
    horizontal: "center",
    vertical: "center",
  },
  fgColor: { rgb: 0xd9d9d9 },
  // @ts-ignore - SheetJS에서 double 스타일 border가 정의에 누락됨
  top: { style: "double" },
  bottom: { style: "thin" },
};

const SUM_DATA_CELL_STYLE: Style = {
  ...DATA_CELL_STYLE,
  fgColor: { rgb: 0xd9d9d9 },
  // @ts-ignore - SheetJS에서 double 스타일 border가 정의에 누락됨
  top: { style: "double" },
  bottom: { style: "thin" },
};

const HEADING_STYLE: Style = {
  bold: true,
  underline: true,
  sz: 15,
  fgColor: { rgb: 0xd9d9d9 },
  alignment: {
    horizontal: "center",
    vertical: "center",
  },
};

const appendInfoTextsToSalesSheet = (sheet: xlsx.Sheet) => {
  sheet["!rows"][0] = HEIGHT_FOR_DIVIDER_ROW;

  sheet["!rows"][1] = HEIGHT_FOR_HEADING_ROW;
  sheet["B2"] = {
    v: "매출신고",
    s: {
      ...HEADING_STYLE,
      color: { rgb: 0xffffff },
      fgColor: { rgb: 0x000000 },
    },
  };
  sheet["C2"] = {
    v: ': "기타" 매출만 신고 해주시면 됩니다.',
    R: [
      {
        v: ': "',
        s: {
          bold: true,
        },
      },
      {
        v: "기타",
        s: {
          bold: true,
          color: RED_TEXT_COLOR,
        },
      },
      {
        v: '" 매출만 신고해주시면 됩니다.',
        s: {
          bold: true,
        },
      },
    ],
    s: {
      bold: true,
      sz: 13,
    },
  };

  sheet["!rows"][2] = HEIGHT_FOR_DIVIDER_ROW;

  sheet["!rows"][3] = HEIGHT_FOR_HEADING_ROW;
  sheet["B4"] = {
    v: "매출신고 시 주의사항",
    s: HEADING_STYLE,
  };

  sheet["!rows"][4] = HEIGHT_FOR_HEADER_ROW;
  sheet["!rows"][5] = HEIGHT_FOR_HEADER_ROW;
  sheet["!rows"][6] = HEIGHT_FOR_HEADER_ROW;
  sheet["B5"] = {
    v:
      '■ 카드(VAN), 현금영수증 발급 금액을 제외한 금액을 " 기타(정규영수증 외 매출분) "매출로 신고하세요',
    R: [
      {
        v: "■ 카드(VAN), 현금영수증 발급 금액을 제외한 금액을 ",
        s: {
          bold: true,
        },
      },
      {
        v: '" 기타(정규영수증 외 매출분) "',
        s: {
          bold: true,
          color: RED_TEXT_COLOR,
        },
      },
      {
        v: "매출로 신고하세요",
        s: {
          bold: true,
        },
      },
    ],
  };
  sheet["B6"] = {
    v:
      "  ( 현금영수증이 포함된 금액 확인 시 해당 월_상세에 적힌 시트에서 금액 확인 부탁 드립니다. )",
    R: [
      {
        v: "  ( 현금영수증이 포함된 금액 확인 시 해당 ",
      },
      {
        v: "월_상세",
        s: {
          bold: true,
          color: RED_TEXT_COLOR,
        },
      },
      {
        v: "에 적힌 시트에서 금액 확인 부탁 드립니다. )",
      },
    ],
  };
  sheet["B7"] = {
    v:
      "  ( 각 월의 합계(매출금액)은 현금영수증이 발행 되었을 경우, 이미 신고가 된 금액이기에 제외 시켰습니다. )",
  };

  sheet["!rows"][7] = HEIGHT_FOR_DIVIDER_ROW;

  sheet["!rows"][8] = HEIGHT_FOR_HEADING_ROW;
  sheet["B9"] = {
    v: "카드매출",
    s: HEADING_STYLE,
  };

  sheet["!rows"][9] = HEIGHT_FOR_HEADER_ROW;
  sheet["!rows"][10] = HEIGHT_FOR_HEADER_ROW;
  sheet["B10"] = {
    v:
      "■ 정확한 카드매출 금액 확인은 KIS정보통신, 여신금융협회, 카드사를 통해 확인 요청 부탁 드립니다.",
    s: {
      bold: true,
    },
  };

  sheet["!rows"][11] = HEIGHT_FOR_SMALL_DIVIDER_ROW;
};

const appendMonthHeadingToSheet = (
  sheet: xlsx.Sheet,
  firstYearTitleRowIndex: number,
  rowsPerYear: number,
  yearCountLength: number
) => {
  sheet[
    xlsx.utils.encode_cell({
      c: 1,
      r: firstYearTitleRowIndex + rowsPerYear * yearCountLength,
    })
  ] = {
    v: "",
    s: {
      fgColor: { rgb: 0xdfeedc },
    },
  };
  sheet["!rows"][
    firstYearTitleRowIndex + rowsPerYear * yearCountLength + 1
  ] = HEIGHT_FOR_SMALL_DIVIDER_ROW;
  sheet[
    xlsx.utils.encode_cell({
      c: 1,
      r: firstYearTitleRowIndex + rowsPerYear * yearCountLength + 2,
    })
  ] = {
    v: "<월>",
    s: HEADING_STYLE,
  };
  sheet["!rows"][
    firstYearTitleRowIndex + rowsPerYear * yearCountLength + 3
  ] = HEIGHT_FOR_SMALL_DIVIDER_ROW;
};

const appendYearHeaderToSheet = ({
  sheet,
  firstRow,
  rowsPerYear,
  year,
}: {
  sheet: xlsx.Sheet;
  firstRow: number;
  rowsPerYear: number;
  year: string;
}) => {
  // ------------- 왼쪽, 오른쪽 여백에 강제로 점선 표시 -------------
  xlsx.utils.sheet_set_range_style(
    sheet,
    xlsx.utils.encode_range(
      { r: firstRow, c: 0 },
      {
        r: firstRow + rowsPerYear - 1,
        c: 0,
      }
    ),
    {
      right: { style: "dotted" },
    }
  );
  xlsx.utils.sheet_set_range_style(
    sheet,
    xlsx.utils.encode_range(
      { r: firstRow, c: 2 },
      {
        r: firstRow + 1,
        c: 2,
      }
    ),
    {
      left: { style: "dotted" },
    }
  );
  xlsx.utils.sheet_set_range_style(
    sheet,
    xlsx.utils.encode_range(
      { r: firstRow + 2, c: 6 },
      {
        r: firstRow + (rowsPerYear - 2),
        c: 6,
      }
    ),
    {
      left: { style: "dotted" },
    }
  );

  // ------------- 연도 및 테이블 제목 -------------
  sheet[
    xlsx.utils.encode_cell({
      r: firstRow,
      c: 1,
    })
  ] = {
    v: "년도",
    s: DATA_HEADING_ROW_STYLE,
  };

  sheet[
    xlsx.utils.encode_cell({
      r: firstRow + 1,
      c: 1,
    })
  ] = {
    v: `${year}년도`,
    s: DATA_YEAR_CELL_STYLE,
  };

  sheet[
    xlsx.utils.encode_cell({
      r: firstRow + 2,
      c: 1,
    })
  ] = {
    v: "구분",
    s: DATA_HEADING_ROW_STYLE,
  };
  sheet[
    xlsx.utils.encode_cell({
      r: firstRow + 2,
      c: 2,
    })
  ] = {
    v: "결제건수",
    s: DATA_HEADING_ROW_STYLE,
  };
  sheet[
    xlsx.utils.encode_cell({
      r: firstRow + 2,
      c: 3,
    })
  ] = {
    v: "주문금액",
    s: DATA_HEADING_ROW_STYLE,
  };
  sheet[
    xlsx.utils.encode_cell({
      r: firstRow + 2,
      c: 4,
    })
  ] = {
    v: "부가세",
    s: DATA_HEADING_ROW_STYLE,
  };
  sheet[
    xlsx.utils.encode_cell({
      r: firstRow + 2,
      c: 5,
    })
  ] = {
    v: "합계",
    s: DATA_HEADING_ROW_STYLE,
  };
};

const getSumCells = ({
  column,
  offset,
  startOfMonthRow,
  yearCount,
  year,
  yearIndex,
  rowsPerMonth,
}: {
  column: number;
  offset: number;
  startOfMonthRow: number;
  yearCount: Record<string, number>;
  year: string;
  yearIndex: number;
  rowsPerMonth: number;
}) => {
  const monthCountBeforeThisYear = Object.keys(yearCount)
    .filter((_, index) => index < yearIndex)
    .reduce((prev, key) => prev + yearCount[key], 0);

  const emptyArrayFromMonthLength = Array.apply(
    undefined,
    Array(yearCount[year])
  );

  return emptyArrayFromMonthLength
    .map((_, monthIndex) =>
      xlsx.utils.encode_cell({
        c: column,
        r:
          startOfMonthRow +
          rowsPerMonth * (monthCountBeforeThisYear + monthIndex) +
          offset,
      })
    )
    .join(",");
};

const appendMonthHeaderToSheet = ({
  sheet,
  sheetName,
  firstRow,
  rowsPerMonth,
}: {
  sheet: xlsx.Sheet;
  sheetName: string;
  firstRow: number;
  rowsPerMonth: number;
}) => {
  // 왼쪽, 오른쪽 여백에 강제로 점선 표시
  xlsx.utils.sheet_set_range_style(
    sheet,
    xlsx.utils.encode_range(
      { r: firstRow, c: 0 },
      {
        r: firstRow + rowsPerMonth - 1,
        c: 0,
      }
    ),
    {
      right: { style: "dotted" },
    }
  );
  xlsx.utils.sheet_set_range_style(
    sheet,
    xlsx.utils.encode_range(
      { r: firstRow, c: 2 },
      {
        r: firstRow,
        c: 2,
      }
    ),
    {
      left: { style: "dotted" },
    }
  );
  xlsx.utils.sheet_set_range_style(
    sheet,
    xlsx.utils.encode_range(
      {
        r: firstRow + 1,
        c: 6,
      },
      {
        r: firstRow + 5,
        c: 6,
      }
    ),
    {
      left: { style: "dotted" },
    }
  );

  sheet[
    xlsx.utils.encode_cell({
      r: firstRow + 0,
      c: 1,
    })
  ] = {
    v: sheetName.replace("_상세", ""),
    s: DATA_HEADING_ROW_STYLE,
  };

  sheet[
    xlsx.utils.encode_cell({
      r: firstRow + 1,
      c: 1,
    })
  ] = {
    v: "구분",
    s: DATA_HEADING_ROW_STYLE,
  };
  sheet[
    xlsx.utils.encode_cell({
      r: firstRow + 1,
      c: 2,
    })
  ] = {
    v: "결제건수",
    s: DATA_HEADING_ROW_STYLE,
  };
  sheet[
    xlsx.utils.encode_cell({
      r: firstRow + 1,
      c: 3,
    })
  ] = {
    v: "주문금액",
    s: DATA_HEADING_ROW_STYLE,
  };
  sheet[
    xlsx.utils.encode_cell({
      r: firstRow + 1,
      c: 4,
    })
  ] = {
    v: "부가세",
    s: DATA_HEADING_ROW_STYLE,
  };
  sheet[
    xlsx.utils.encode_cell({
      r: firstRow + 1,
      c: 5,
    })
  ] = {
    v: "합계",
    s: DATA_HEADING_ROW_STYLE,
  };
};

const generateSalesSummarySheet = (
  sheetNames: string[],
  yearCount: Record<string, number>
) => {
  const FIRST_YEAR_TITLE_ROW_INDEX = 14;
  const ROWS_PER_YEAR = 8;
  const MARGIN_BETWEEN_YEAR_AND_MONTH = 4;
  const ROWS_PER_MONTH = 7;

  const yearLength = Object.keys(yearCount).length;

  const rowCount =
    FIRST_YEAR_TITLE_ROW_INDEX +
    ROWS_PER_YEAR * yearLength +
    MARGIN_BETWEEN_YEAR_AND_MONTH +
    ROWS_PER_MONTH * sheetNames.length;
  const aoa = Array.apply(undefined, Array(rowCount)).map(() => emptyRow);
  const sheet = xlsx.utils.aoa_to_sheet(aoa);

  sheet["!merges"] = [
    xlsx.utils.decode_range("C2:J2"),
    xlsx.utils.decode_range("B4:C4"),
    xlsx.utils.decode_range("B5:J5"),
    xlsx.utils.decode_range("B6:J6"),
    xlsx.utils.decode_range("B7:J7"),
    xlsx.utils.decode_range("B9:C9"),
    xlsx.utils.decode_range("B10:J10"),
    {
      s: {
        c: 1,
        r: FIRST_YEAR_TITLE_ROW_INDEX + ROWS_PER_YEAR * yearLength,
      },
      e: {
        c: 5,
        r: FIRST_YEAR_TITLE_ROW_INDEX + ROWS_PER_YEAR * yearLength,
      },
    },
  ];

  sheet["!sheetFormat"] = {
    row: {
      hpx: 21,
    },
    col: {
      wpx: 100,
    },
  };

  sheet["!rows"] = [];
  sheet["!cols"] = [];

  sheet["!cols"][0] = { wpx: 18 };
  sheet["!cols"][1] = { wpx: 110 };
  sheet["!cols"][6] = { wpx: 18 };

  xlsx.utils.sheet_set_range_style(
    sheet,
    xlsx.utils.encode_range({ c: 0, r: 0 }, { c: 9, r: rowCount }),
    DEFAULT_STYLE
  );

  // 안내
  appendInfoTextsToSalesSheet(sheet);

  sheet["B13"] = {
    v: "<전체>",
    s: HEADING_STYLE,
  };

  sheet["!rows"][13] = HEIGHT_FOR_SMALL_DIVIDER_ROW;

  // 연도별 통계
  Object.keys(yearCount).forEach((year, yearIndex) => {
    const firstRow = FIRST_YEAR_TITLE_ROW_INDEX + ROWS_PER_YEAR * yearIndex;

    appendYearHeaderToSheet({
      sheet,
      firstRow,
      rowsPerYear: ROWS_PER_YEAR,
      year,
    });

    // SUM
    const getSumCellsForSales = (column: number, offset: number) => {
      return getSumCells({
        column,
        offset,
        startOfMonthRow:
          FIRST_YEAR_TITLE_ROW_INDEX +
          ROWS_PER_YEAR * yearLength +
          MARGIN_BETWEEN_YEAR_AND_MONTH,
        yearCount,
        year,
        yearIndex,
        rowsPerMonth: ROWS_PER_MONTH,
      });
    };

    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 1,
      })
    ] = {
      v: "신용카드(VAN)",
      s: DATA_HEADING_COLUMN_STYLE,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 2,
      })
    ] = {
      t: "n",
      f: `SUM(${getSumCellsForSales(2, 2)})`,
      s: DATA_CELL_STYLE,
      z: COUNT_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 3,
      })
    ] = {
      t: "n",
      f: `SUM(${getSumCellsForSales(3, 2)})`,
      s: DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 4,
      })
    ] = {
      t: "n",
      f: `SUM(${getSumCellsForSales(4, 2)})`,
      s: DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 5,
      })
    ] = {
      t: "n",
      f: `SUM(${getSumCellsForSales(5, 2)})`,
      s: DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };

    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 4,
        c: 1,
      })
    ] = {
      v: "현금영수증발행분",
      s: DATA_HEADING_COLUMN_STYLE,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 4,
        c: 2,
      })
    ] = {
      t: "n",
      f: `SUM(${getSumCellsForSales(2, 3)})`,
      s: DATA_CELL_STYLE,
      z: COUNT_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 4,
        c: 3,
      })
    ] = {
      t: "n",
      f: `SUM(${getSumCellsForSales(3, 3)})`,
      s: DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 4,
        c: 4,
      })
    ] = {
      t: "n",
      f: `SUM(${getSumCellsForSales(4, 3)})`,
      s: DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 4,
        c: 5,
      })
    ] = {
      t: "n",
      f: `SUM(${getSumCellsForSales(5, 3)})`,
      s: DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };

    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 5,
        c: 1,
      })
    ] = {
      v: "기타",
      s: ETC_DATA_HEADING_COLUMN_STYLE,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 5,
        c: 2,
      })
    ] = {
      t: "n",
      f: `SUM(${getSumCellsForSales(2, 4)})`,
      s: ETC_DATA_CELL_STYLE,
      z: COUNT_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 5,
        c: 3,
      })
    ] = {
      t: "n",
      f: `SUM(${getSumCellsForSales(3, 4)})`,
      s: ETC_DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 5,
        c: 4,
      })
    ] = {
      t: "n",
      f: `SUM(${getSumCellsForSales(4, 4)})`,
      s: ETC_DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 5,
        c: 5,
      })
    ] = {
      t: "n",
      f: `SUM(${getSumCellsForSales(5, 4)})`,
      s: ETC_DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };

    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 6,
        c: 1,
      })
    ] = {
      v: "합계",
      s: SUM_DATA_HEADING_COLUMN_STYLE,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 6,
        c: 2,
      })
    ] = {
      t: "n",
      f: `SUM(${xlsx.utils.encode_range(
        {
          r: firstRow + 3,
          c: 2,
        },
        {
          r: firstRow + 5,
          c: 2,
        }
      )})`,
      s: SUM_DATA_CELL_STYLE,
      z: COUNT_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 6,
        c: 3,
      })
    ] = {
      t: "n",
      f: `SUM(${xlsx.utils.encode_range(
        {
          r: firstRow + 3,
          c: 3,
        },
        {
          r: firstRow + 5,
          c: 3,
        }
      )})`,
      s: SUM_DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 6,
        c: 4,
      })
    ] = {
      t: "n",
      f: `SUM(${xlsx.utils.encode_range(
        {
          r: firstRow + 3,
          c: 4,
        },
        {
          r: firstRow + 5,
          c: 4,
        }
      )})`,
      s: SUM_DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 6,
        c: 5,
      })
    ] = {
      t: "n",
      f: `SUM(${xlsx.utils.encode_range(
        {
          r: firstRow + 3,
          c: 5,
        },
        {
          r: firstRow + 5,
          c: 5,
        }
      )})`,
      s: SUM_DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };

    sheet["!rows"]![firstRow + 7] = HEIGHT_FOR_SMALL_DIVIDER_ROW;
  });

  appendMonthHeadingToSheet(
    sheet,
    FIRST_YEAR_TITLE_ROW_INDEX,
    ROWS_PER_YEAR,
    yearLength
  );

  const firstMonthTitleRowIndex =
    FIRST_YEAR_TITLE_ROW_INDEX +
    ROWS_PER_YEAR * yearLength +
    MARGIN_BETWEEN_YEAR_AND_MONTH;

  sheetNames.forEach((sheetName, monthIndex) => {
    const firstRow = firstMonthTitleRowIndex + monthIndex * ROWS_PER_MONTH;

    appendMonthHeaderToSheet({
      sheet,
      sheetName,
      firstRow,
      rowsPerMonth: ROWS_PER_MONTH,
    });

    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 2,
        c: 1,
      })
    ] = {
      v: "신용카드(VAN)",
      s: DATA_HEADING_COLUMN_STYLE,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 2,
        c: 2,
      })
    ] = {
      t: "n",
      f: `COUNTIF('${sheetName}'!F13:F1048576,">0")`,
      s: DATA_CELL_STYLE,
      z: COUNT_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 2,
        c: 3,
      })
    ] = {
      t: "n",
      f: `ROUNDDOWN(${xlsx.utils.encode_cell({
        r: firstRow + 2,
        c: 5,
      })}/1.1,0)`,
      s: DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 2,
        c: 4,
      })
    ] = {
      t: "n",
      f: `${xlsx.utils.encode_cell({
        r: firstRow + 2,
        c: 5,
      })}-${xlsx.utils.encode_cell({
        r: firstRow + 2,
        c: 3,
      })}`,
      s: DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 2,
        c: 5,
      })
    ] = {
      t: "n",
      f: `SUM(${sheetName}!F13:F1048576)`,
      s: DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };

    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 1,
      })
    ] = {
      v: "현금영수증발행분",
      s: DATA_HEADING_COLUMN_STYLE,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 2,
      })
    ] = {
      t: "n",
      f: `COUNTIF('${sheetName}'!I13:I1048576,">0")`,
      s: DATA_CELL_STYLE,
      z: COUNT_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 3,
      })
    ] = {
      t: "n",
      f: `ROUNDDOWN(${xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 5,
      })}/1.1,0)`,
      s: DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 4,
      })
    ] = {
      t: "n",
      f: `${xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 5,
      })}-${xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 3,
      })}`,
      s: DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 5,
      })
    ] = {
      t: "n",
      f: `SUM(${sheetName}!I13:I1048576)`,
      s: DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };

    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 4,
        c: 1,
      })
    ] = {
      v: "기타",
      s: ETC_DATA_HEADING_COLUMN_STYLE,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 4,
        c: 2,
      })
    ] = {
      t: "n",
      f: `COUNTIF('${sheetName}'!C13:C1048576,">0")`,
      s: ETC_DATA_CELL_STYLE,
      z: COUNT_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 4,
        c: 3,
      })
    ] = {
      t: "n",
      f: `ROUNDDOWN(${xlsx.utils.encode_cell({
        r: firstRow + 4,
        c: 5,
      })}/1.1,0)`,
      s: ETC_DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 4,
        c: 4,
      })
    ] = {
      t: "n",
      f: `${xlsx.utils.encode_cell({
        r: firstRow + 4,
        c: 5,
      })}-${xlsx.utils.encode_cell({
        r: firstRow + 4,
        c: 3,
      })}`,
      s: ETC_DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 4,
        c: 5,
      })
    ] = {
      t: "n",
      f: `SUM(${sheetName}!C13:C1048576)`,
      s: ETC_DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };

    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 5,
        c: 1,
      })
    ] = {
      v: "합계",
      s: SUM_DATA_HEADING_COLUMN_STYLE,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 5,
        c: 2,
      })
    ] = {
      t: "n",
      f: `SUM(${xlsx.utils.encode_range(
        {
          r: firstRow + 2,
          c: 2,
        },
        {
          r: firstRow + 4,
          c: 2,
        }
      )})`,
      s: SUM_DATA_CELL_STYLE,
      z: COUNT_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 5,
        c: 3,
      })
    ] = {
      t: "n",
      f: `SUM(${xlsx.utils.encode_range(
        {
          r: firstRow + 2,
          c: 3,
        },
        {
          r: firstRow + 4,
          c: 3,
        }
      )})`,
      s: SUM_DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 5,
        c: 4,
      })
    ] = {
      t: "n",
      f: `SUM(${xlsx.utils.encode_range(
        {
          r: firstRow + 2,
          c: 4,
        },
        {
          r: firstRow + 4,
          c: 4,
        }
      )})`,
      s: SUM_DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 5,
        c: 5,
      })
    ] = {
      t: "n",
      f: `SUM(${xlsx.utils.encode_range(
        {
          r: firstRow + 2,
          c: 5,
        },
        {
          r: firstRow + 4,
          c: 5,
        }
      )})`,
      s: SUM_DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };

    sheet["!rows"]![firstRow + 6] = HEIGHT_FOR_DIVIDER_ROW;
  });

  // 따로 설정되지 않은 테두리색을 화이트로 설정
  aoa.forEach((row, rowIndex) => {
    row.forEach((_, colIndex) => {
      const cell = sheet[xlsx.utils.encode_cell({ r: rowIndex, c: colIndex })];
      cell.s = cell.s
        ? {
            ...WHITE_BORDER_STYLE,
            ...cell.s,
          }
        : WHITE_BORDER_STYLE;
    });
  });

  return sheet;
};

const generateBuyingSummarySheet = (
  sheetNames: string[],
  yearCount: Record<string, number>
) => {
  const FIRST_YEAR_TITLE_ROW_INDEX = 5;
  const ROWS_PER_YEAR = 6;
  const MARGIN_BETWEEN_YEAR_AND_MONTH = 4;
  const ROWS_PER_MONTH = 5;

  const yearLength = Object.keys(yearCount).length;

  const rowCount =
    FIRST_YEAR_TITLE_ROW_INDEX +
    ROWS_PER_YEAR * yearLength +
    MARGIN_BETWEEN_YEAR_AND_MONTH +
    ROWS_PER_MONTH * sheetNames.length;
  const aoa = Array.apply(undefined, Array(rowCount)).map(() => emptyRow);
  const sheet = xlsx.utils.aoa_to_sheet(aoa);

  sheet["!merges"] = [
    xlsx.utils.decode_range("C2:J2"),
    {
      s: {
        c: 1,
        r: FIRST_YEAR_TITLE_ROW_INDEX + ROWS_PER_YEAR * yearLength,
      },
      e: {
        c: 5,
        r: FIRST_YEAR_TITLE_ROW_INDEX + ROWS_PER_YEAR * yearLength,
      },
    },
  ];

  sheet["!sheetFormat"] = {
    row: {
      hpx: 21,
    },
    col: {
      wpx: 100,
    },
  };

  sheet["!rows"] = [];
  sheet["!cols"] = [];

  sheet["!cols"][0] = { wpx: 18 };
  sheet["!cols"][1] = { wpx: 110 };
  sheet["!cols"][6] = { wpx: 18 };

  xlsx.utils.sheet_set_range_style(
    sheet,
    xlsx.utils.encode_range({ c: 0, r: 0 }, { c: 9, r: rowCount }),
    DEFAULT_STYLE
  );

  sheet["!rows"][0] = HEIGHT_FOR_DIVIDER_ROW;

  sheet["B2"] = {
    v: "매입신고",
    s: {
      ...HEADING_STYLE,
      color: { rgb: 0xffffff },
      fgColor: { rgb: 0x000000 },
    },
  };
  sheet["C2"] = {
    v:
      ': 매월 1~10일 사이 "메쉬코리아" 에서 발급하여 이메일로 전송 하고 있습니다.',
    s: {
      bold: true,
      sz: 13,
    },
  };

  sheet["!rows"][2] = HEIGHT_FOR_SMALL_DIVIDER_ROW;
  sheet["B4"] = {
    v: "<전체>",
    s: HEADING_STYLE,
  };
  sheet["!rows"][4] = HEIGHT_FOR_SMALL_DIVIDER_ROW;

  // 연도별 통계
  Object.keys(yearCount).forEach((year, yearIndex) => {
    const firstRow = FIRST_YEAR_TITLE_ROW_INDEX + ROWS_PER_YEAR * yearIndex;

    appendYearHeaderToSheet({
      sheet,
      firstRow,
      rowsPerYear: ROWS_PER_YEAR,
      year,
    });

    // SUM
    const getSumCellsForBuying = (column: number, offset: number) => {
      return getSumCells({
        column,
        offset,
        startOfMonthRow:
          FIRST_YEAR_TITLE_ROW_INDEX +
          ROWS_PER_YEAR * yearLength +
          MARGIN_BETWEEN_YEAR_AND_MONTH,
        yearCount,
        year,
        yearIndex,
        rowsPerMonth: ROWS_PER_MONTH,
      });
    };

    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 1,
      })
    ] = {
      v: "정산배송비",
      s: DATA_HEADING_COLUMN_STYLE,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 2,
      })
    ] = {
      t: "n",
      f: `SUM(${getSumCellsForBuying(2, 2)})`,
      s: DATA_CELL_STYLE,
      z: COUNT_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 3,
      })
    ] = {
      t: "n",
      f: `SUM(${getSumCellsForBuying(3, 2)})`,
      s: DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 4,
      })
    ] = {
      t: "n",
      f: `SUM(${getSumCellsForBuying(4, 2)})`,
      s: DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 5,
      })
    ] = {
      t: "n",
      f: `SUM(${getSumCellsForBuying(5, 2)})`,
      s: DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };

    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 4,
        c: 1,
      })
    ] = {
      v: "합계",
      s: SUM_DATA_HEADING_COLUMN_STYLE,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 4,
        c: 2,
      })
    ] = {
      t: "n",
      f: `SUM(${xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 2,
      })})`,
      s: SUM_DATA_CELL_STYLE,
      z: COUNT_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 4,
        c: 3,
      })
    ] = {
      t: "n",
      f: `SUM(${xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 3,
      })})`,
      s: SUM_DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 4,
        c: 4,
      })
    ] = {
      t: "n",
      f: `SUM(${xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 4,
      })})`,
      s: SUM_DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 4,
        c: 5,
      })
    ] = {
      t: "n",
      f: `SUM(${xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 5,
      })})`,
      s: SUM_DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };

    sheet["!rows"]![firstRow + 5] = HEIGHT_FOR_SMALL_DIVIDER_ROW;
  });

  appendMonthHeadingToSheet(
    sheet,
    FIRST_YEAR_TITLE_ROW_INDEX,
    ROWS_PER_YEAR,
    Object.keys(yearCount).length
  );

  const firstMonthTitleRowIndex =
    FIRST_YEAR_TITLE_ROW_INDEX +
    ROWS_PER_YEAR * yearLength +
    MARGIN_BETWEEN_YEAR_AND_MONTH;

  sheetNames.forEach((sheetName, monthIndex) => {
    const firstRow = firstMonthTitleRowIndex + monthIndex * ROWS_PER_MONTH;

    appendMonthHeaderToSheet({
      sheet,
      sheetName,
      firstRow,
      rowsPerMonth: ROWS_PER_MONTH,
    });

    const homeTaxInfoCell = xlsx.utils.encode_cell({
      r: firstRow,
      c: 2,
    });

    sheet[homeTaxInfoCell] = {
      v: ": 세금계산서 발행금액은 국세청 홈택스에서도 정확히 확인 가능합니다.",
      s: {
        ...sheet[homeTaxInfoCell].s,
        sz: 10,
        alignment: {
          horizontal: "left",
          vertical: "center",
        },
      },
    };
    sheet["!merges"] = [
      ...sheet["!merges"]!,
      {
        s: {
          r: firstRow,
          c: 2,
        },
        e: {
          r: firstRow,
          c: 5,
        },
      },
    ];

    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 2,
        c: 1,
      })
    ] = {
      v: "정산배송비",
      s: DATA_HEADING_COLUMN_STYLE,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 2,
        c: 2,
      })
    ] = {
      t: "n",
      f: `COUNTA('${sheetName}'!B13:B1048576)`,
      s: DATA_CELL_STYLE,
      z: COUNT_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 2,
        c: 3,
      })
    ] = {
      t: "n",
      f: `ROUNDDOWN(${xlsx.utils.encode_cell({
        r: firstRow + 2,
        c: 5,
      })}/1.1,0)`,
      s: DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 2,
        c: 4,
      })
    ] = {
      t: "n",
      f: `${xlsx.utils.encode_cell({
        r: firstRow + 2,
        c: 5,
      })}-${xlsx.utils.encode_cell({
        r: firstRow + 2,
        c: 3,
      })}`,
      s: DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 2,
        c: 5,
      })
    ] = {
      t: "n",
      f: `SUM(${sheetName}!D13:D1048576)`,
      s: DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };

    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 1,
      })
    ] = {
      v: "합계",
      s: SUM_DATA_HEADING_COLUMN_STYLE,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 2,
      })
    ] = {
      t: "n",
      f: `SUM(${xlsx.utils.encode_cell({
        r: firstRow + 2,
        c: 2,
      })})`,
      s: SUM_DATA_CELL_STYLE,
      z: COUNT_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 3,
      })
    ] = {
      t: "n",
      f: `SUM(${xlsx.utils.encode_cell({
        r: firstRow + 2,
        c: 3,
      })})`,
      s: SUM_DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 4,
      })
    ] = {
      t: "n",
      f: `SUM(${xlsx.utils.encode_cell({
        r: firstRow + 2,
        c: 4,
      })})`,
      s: SUM_DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };
    sheet[
      xlsx.utils.encode_cell({
        r: firstRow + 3,
        c: 5,
      })
    ] = {
      t: "n",
      f: `SUM(${xlsx.utils.encode_cell({
        r: firstRow + 2,
        c: 5,
      })})`,
      s: SUM_DATA_CELL_STYLE,
      z: PRICE_FORMAT,
    };

    sheet["!rows"]![firstRow + 4] = HEIGHT_FOR_DIVIDER_ROW;
  });

  // 따로 설정되지 않은 테두리색을 화이트로 설정
  aoa.forEach((row, rowIndex) => {
    row.forEach((_, colIndex) => {
      const cell = sheet[xlsx.utils.encode_cell({ r: rowIndex, c: colIndex })];
      cell.s = cell.s
        ? {
            ...WHITE_BORDER_STYLE,
            ...cell.s,
          }
        : WHITE_BORDER_STYLE;
    });
  });

  return sheet;
};

const generateSummarySheet = (sheetNames: string[]) => {
  // IE는 왠지 Set polyfill 지원을 잘 안할 것 같아서...
  const yearCount = sheetNames.reduce<Record<string, number>>((prev, name) => {
    const year = name.slice(0, 4);

    return {
      ...prev,
      [year]: (prev[year] || 0) + 1,
    };
  }, {});

  return {
    salesSheet: generateSalesSummarySheet(sheetNames, yearCount),
    buyingSheet: generateBuyingSummarySheet(sheetNames, yearCount),
  };
};

export default generateSummarySheet;
