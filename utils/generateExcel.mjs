import * as XLSX from "xlsx";

function normalizeText(value) {
    return String(value ?? "")
        .replace(/\s+/g, " ")
        .trim()
        .toLowerCase();
}

function isEmptyValue(value) {
    return value === undefined || value === null || value === "";
}

function createCellFromSource(sourceCell, templateCell) {
    if (!sourceCell || isEmptyValue(sourceCell.v)) {
        return templateCell
            ? { ...templateCell, t: "s", v: "" }
            : { t: "s", v: "" };
    }

    const result = {
        ...(templateCell ?? {}),
        t: sourceCell.t,
        v: sourceCell.v,
    };

    delete result.f;
    delete result.F;
    delete result.w;

    if (!result.z && sourceCell.z) {
        result.z = sourceCell.z;
    }

    return result;
}

function findTemplateColumns(templateSheet, headerRowNumber) {
    const range = XLSX.utils.decode_range(templateSheet["!ref"]);
    const headerRowIndex = headerRowNumber - 1;
    const columns = new Map();

    for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
        const address = XLSX.utils.encode_cell({
            r: headerRowIndex,
            c: columnIndex,
        });
        const cell = templateSheet[address];
        const normalizedHeader = normalizeText(cell?.v);

        if (normalizedHeader && !columns.has(normalizedHeader)) {
            columns.set(normalizedHeader, columnIndex);
        }
    }

    return columns;
}

function createOutputFileName(templateFileName) {
    const baseName = templateFileName.replace(/\.[^.]+$/, "");
    return `${baseName}-已填写.xlsx`;
}

export default async function generateExcel({
    sourceFile,
    templateFile,
    mappingResult,
}) {
    const [sourceBuffer, templateBuffer] = await Promise.all([
        sourceFile.arrayBuffer(),
        templateFile.arrayBuffer(),
    ]);

    const sourceWorkbook = XLSX.read(sourceBuffer, {
        type: "array",
        cellDates: true,
        cellFormula: true,
        cellStyles: true,
    });
    const templateWorkbook = XLSX.read(templateBuffer, {
        type: "array",
        cellDates: true,
        cellFormula: true,
        cellStyles: true,
    });

    const sourceSheet = sourceWorkbook.Sheets[mappingResult.sourceSheetName];
    const templateSheet = templateWorkbook.Sheets[mappingResult.templateSheetName];

    if (!sourceSheet) {
        throw new Error(`业务 Excel 中找不到工作表：${mappingResult.sourceSheetName}`);
    }

    if (!templateSheet) {
        throw new Error(`模板 Excel 中找不到工作表：${mappingResult.templateSheetName}`);
    }

    if (!Number.isInteger(mappingResult.sourceHeaderRow)) {
        throw new Error("豆包没有返回有效的业务表头行");
    }

    if (!Number.isInteger(mappingResult.templateHeaderRow)) {
        throw new Error("豆包没有返回有效的模板表头行");
    }

    const activeMappings = mappingResult.fieldMappings.filter((mapping) => {
        return Number.isInteger(mapping.sourceColumnIndex);
    });

    if (!activeMappings.length) {
        throw new Error("当前没有已确认的字段匹配关系");
    }

    const templateColumns = findTemplateColumns(
        templateSheet,
        mappingResult.templateHeaderRow,
    );
    const resolvedMappings = activeMappings.map((mapping) => {
        const targetColumnIndex = templateColumns.get(
            normalizeText(mapping.targetField),
        );

        return {
            ...mapping,
            targetColumnIndex,
        };
    });
    const missingTargetFields = resolvedMappings
        .filter((mapping) => !Number.isInteger(mapping.targetColumnIndex))
        .map((mapping) => mapping.targetField);

    if (missingTargetFields.length) {
        throw new Error(`模板表头中找不到字段：${missingTargetFields.join("、")}`);
    }

    const sourceRange = XLSX.utils.decode_range(sourceSheet["!ref"]);
    const sourceStartRowIndex = Number.isInteger(mappingResult.sourceDataStartRow)
        ? mappingResult.sourceDataStartRow - 1
        : mappingResult.sourceHeaderRow;
    const sourceEndRowIndex = Number.isInteger(mappingResult.sourceDataEndRow)
        ? Math.min(mappingResult.sourceDataEndRow - 1, sourceRange.e.r)
        : sourceRange.e.r;
    const templateStartRowIndex = Number.isInteger(mappingResult.templateDataStartRow)
        ? mappingResult.templateDataStartRow - 1
        : mappingResult.templateHeaderRow;
    let outputRowIndex = templateStartRowIndex;
    let writtenRowCount = 0;

    for (
        let sourceRowIndex = sourceStartRowIndex;
        sourceRowIndex <= sourceEndRowIndex;
        sourceRowIndex += 1
    ) {
        const rowValues = resolvedMappings.map((mapping) => {
            const sourceAddress = XLSX.utils.encode_cell({
                r: sourceRowIndex,
                c: mapping.sourceColumnIndex,
            });

            return sourceSheet[sourceAddress];
        });

        if (rowValues.every((cell) => isEmptyValue(cell?.v))) {
            continue;
        }

        resolvedMappings.forEach((mapping, mappingIndex) => {
            const targetAddress = XLSX.utils.encode_cell({
                r: outputRowIndex,
                c: mapping.targetColumnIndex,
            });
            const styleAddress = XLSX.utils.encode_cell({
                r: templateStartRowIndex,
                c: mapping.targetColumnIndex,
            });
            const templateCell = templateSheet[targetAddress]
                ?? templateSheet[styleAddress];

            templateSheet[targetAddress] = createCellFromSource(
                rowValues[mappingIndex],
                templateCell,
            );
        });

        outputRowIndex += 1;
        writtenRowCount += 1;
    }

    if (!writtenRowCount) {
        throw new Error("业务 Excel 表头下方没有可写入的数据");
    }

    const templateRange = XLSX.utils.decode_range(templateSheet["!ref"]);
    templateRange.e.r = Math.max(templateRange.e.r, outputRowIndex - 1);
    templateSheet["!ref"] = XLSX.utils.encode_range(templateRange);

    const outputFileName = createOutputFileName(templateFile.name);
    XLSX.writeFile(templateWorkbook, outputFileName, {
        bookType: "xlsx",
        cellStyles: true,
    });

    return {
        outputFileName,
        writtenRowCount,
    };
}
