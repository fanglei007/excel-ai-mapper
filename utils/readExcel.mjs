import * as XLSX from "xlsx";

function isEmptyCell(value) {
    return value === "" || value === null || value === undefined;
}

function parseWorksheet(sheet, sheetName) {
    if (!sheet["!ref"]) {
        return {
            sheetName,
            range: null,
            rows: [],
        };
    }

    const columnRows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
        raw: false,
    });

    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const rows = columnRows
        .map((columns, index) => ({
            rowIndex: range.s.r + index + 1,
            columns,
        }))
        .filter(({ columns }) => columns.some((column) => !isEmptyCell(column)));

    return {
        sheetName,
        range: sheet["!ref"],
        rows,
    };
}

export default async function readExcel(file) {
    if (!(file instanceof Blob)) {
        throw new TypeError("请传入浏览器选择的 Excel 文件");
    }

    const fileBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(fileBuffer, {
        type: "array",
        cellFormula: true,
        cellDates: true,
    });

    const sheets = workbook.SheetNames.map((sheetName) => {
        return parseWorksheet(workbook.Sheets[sheetName], sheetName);
    });

    return {
        fileName: file.name,
        sheetNames: workbook.SheetNames,
        sheets,
    };
}
