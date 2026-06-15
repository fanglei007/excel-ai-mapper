import request from "../utils/request.mjs";

const RESPONSES_API_URL = "https://ark.cn-beijing.volces.com/api/v3/responses";
const MODEL = "doubao-seed-2-0-lite-260215";

function extractOutputText(response) {
    if (typeof response.output_text === "string") {
        return response.output_text;
    }

    const textParts = response.output
        ?.flatMap((item) => item.content ?? [])
        .filter((content) => content.type === "output_text")
        .map((content) => content.text);

    if (textParts?.length) {
        return textParts.join("");
    }

    throw new Error("豆包响应中没有可解析的文本内容");
}

function validateMappingResult(result) {
    const fieldMappings = result?.fieldMappings ?? result?.field_mappings;

    if (!result || !Array.isArray(fieldMappings)) {
        throw new Error("豆包返回的字段匹配格式不正确");
    }

    return {
        sourceSheetName: result.sourceSheetName ?? result.source_sheet_name ?? "",
        sourceHeaderRow: result.sourceHeaderRow ?? result.source_header_row ?? null,
        sourceDataStartRow: result.sourceDataStartRow ?? result.source_data_start_row ?? null,
        sourceDataEndRow: result.sourceDataEndRow ?? result.source_data_end_row ?? null,
        templateSheetName: result.templateSheetName ?? result.template_sheet_name ?? "",
        templateHeaderRow: result.templateHeaderRow ?? result.template_header_row ?? null,
        templateDataStartRow: result.templateDataStartRow ?? result.template_data_start_row ?? null,
        fieldMappings: fieldMappings.map((mapping) => ({
            targetField: String(mapping.targetField ?? mapping.target_field ?? ""),
            sourceHeader: (mapping.sourceHeader ?? mapping.source_header) === null
                ? null
                : String(mapping.sourceHeader ?? mapping.source_header ?? ""),
            sourceColumnIndex: Number.isInteger(
                mapping.sourceColumnIndex ?? mapping.source_column_index,
            )
                ? mapping.sourceColumnIndex ?? mapping.source_column_index
                : null,
            confidence: typeof mapping.confidence === "number"
                ? Math.max(0, Math.min(1, mapping.confidence))
                : 0,
            reason: String(mapping.reason ?? ""),
        })),
        warnings: Array.isArray(result.warnings)
            ? result.warnings.map(String)
            : [],
    };
}

export default async function matchExcelFields({
    apiKey,
    sourceExcel,
    templateExcel,
}) {
    if (!apiKey) {
        throw new Error("缺少豆包 API Key");
    }

    const response = await request(RESPONSES_API_URL, {
        method: "POST",
        timeoutMs: 180000,
        headers: {
            Authorization: `Bearer ${apiKey}`,
        },
        body: {
            model: MODEL,
            instructions: `
你是一个 Excel 字段匹配助手。
请识别源业务 Excel 的商品明细工作表、表头行和字段，并识别目标模板 Excel 的目标字段。
请识别源业务 Excel 商品明细数据的开始行和结束行，排除标题、合计、备注和页脚。
将每一个目标模板字段匹配到最合适的源字段。
找不到可靠来源时，sourceHeader 和 sourceColumnIndex 必须返回 null。
sourceColumnIndex 是源表 columns 数组中的下标，从 0 开始。
confidence 是 0 到 1 之间的数字。
只返回 JSON，不要使用 Markdown，不要添加解释性前后缀。
            `.trim(),
            input: JSON.stringify({
                task: "分析业务 Excel 与模板 Excel 的字段匹配关系",
                sourceExcel,
                templateExcel,
                outputFormat: {
                    sourceSheetName: "string",
                    sourceHeaderRow: "number|null",
                    sourceDataStartRow: "number|null",
                    sourceDataEndRow: "number|null",
                    templateSheetName: "string",
                    templateHeaderRow: "number|null",
                    templateDataStartRow: "number|null",
                    fieldMappings: [
                        {
                            targetField: "string",
                            sourceHeader: "string|null",
                            sourceColumnIndex: "number|null",
                            confidence: "number",
                            reason: "string",
                        },
                    ],
                    warnings: ["string"],
                },
            }),
            text: {
                format: {
                    type: "json_object",
                },
            },
        },
    });

    const outputText = extractOutputText(response);

    try {
        return validateMappingResult(JSON.parse(outputText));
    } catch (error) {
        if (error instanceof SyntaxError) {
            throw new Error("豆包返回的内容不是有效 JSON");
        }

        throw error;
    }
}
