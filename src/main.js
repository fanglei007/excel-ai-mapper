import "./style.css";
import readExcel from "../utils/readExcel.mjs";
import generateExcel from "../utils/generateExcel.mjs";

const appElement = document.querySelector("#app");

appElement.innerHTML = `
    <div class="app-shell">
        <header class="app-header">
            <div>
                <p class="app-kicker">EXCEL AI MAPPER</p>
                <h1>智能字段匹配</h1>
                <p class="app-description">选择业务 Excel 和目标模板，确认字段关系后生成填写完成的文件。</p>
            </div>
            <span class="status-badge">本地处理</span>
        </header>

        <main class="workspace">
            <section class="section-block" aria-labelledby="setup-title">
                <div class="section-heading">
                    <span class="step-number">1</span>
                    <div>
                        <h2 id="setup-title">配置与文件</h2>
                        <p>输入模型凭证，并选择待处理文件。</p>
                    </div>
                </div>

                <div class="form-grid">
                    <label class="field key-field">
                        <span class="field-label">豆包 API Key</span>
                        <input
                            id="apiKey"
                            type="password"
                            name="apiKey"
                            autocomplete="off"
                            placeholder="务必保管好自己的key不要泄漏"
                        >
                        <span class="field-hint">Key 仅用于本机调用模型，不会展示在页面中。</span>
                    </label>

                    <div class="file-field">
                        <div>
                            <span class="field-label">业务 Excel</span>
                            <span id="sourceFileName" class="field-hint">需要识别和转换的原始文件</span>
                        </div>
                        <label class="file-picker">
                            <input id="sourceExcel" type="file" accept=".xlsx,.xls">
                            <span>选择文件</span>
                        </label>
                    </div>

                    <div class="file-field">
                        <div>
                            <span class="field-label">模板 Excel</span>
                            <span id="templateFileName" class="field-hint">数据最终要填入的目标模板</span>
                        </div>
                        <label class="file-picker">
                            <input id="templateExcel" type="file" accept=".xlsx,.xls">
                            <span>选择文件</span>
                        </label>
                    </div>
                </div>

                <div class="section-actions">
                    <p id="formMessage" class="form-message" role="status"></p>
                    <button id="parseButton" class="primary-button" type="button">开始解析</button>
                </div>
            </section>

            <section class="section-block" aria-labelledby="parse-title">
                <div class="section-heading">
                    <span class="step-number">2</span>
                    <div>
                        <h2 id="parse-title">解析结果</h2>
                        <p>查看模型识别出的工作表、表头和数据范围。</p>
                    </div>
                </div>

                <details id="resultDetails" class="result-details">
                    <summary>
                        <span id="resultSummaryTitle">等待解析业务 Excel</span>
                        <span id="resultSummaryMeta" class="summary-meta">尚无数据</span>
                    </summary>
                    <div id="resultContent" class="result-content">
                        <dl class="result-stats">
                            <div>
                                <dt>文件</dt>
                                <dd id="resultFileName">--</dd>
                            </div>
                            <div>
                                <dt>工作表</dt>
                                <dd id="resultSheetCount">--</dd>
                            </div>
                            <div>
                                <dt>非空数据行</dt>
                                <dd id="resultRowCount">--</dd>
                            </div>
                        </dl>
                        <pre id="resultJson">请选择文件并点击“开始解析”。</pre>
                    </div>
                </details>
            </section>

            <section class="section-block" aria-labelledby="doubao-result-title">
                <div class="section-heading">
                    <span class="step-number">3</span>
                    <div>
                        <h2 id="doubao-result-title">豆包分析结果</h2>
                        <p>查看模型返回的工作表识别和字段匹配 JSON。</p>
                    </div>
                </div>

                <details id="doubaoResultDetails" class="result-details">
                    <summary>
                        <span id="doubaoResultSummaryTitle">等待豆包分析</span>
                        <span id="doubaoResultSummaryMeta" class="summary-meta">尚无数据</span>
                    </summary>
                    <div class="result-content">
                        <dl class="result-stats">
                            <div>
                                <dt>源工作表</dt>
                                <dd id="doubaoSourceSheet">--</dd>
                            </div>
                            <div>
                                <dt>模板工作表</dt>
                                <dd id="doubaoTemplateSheet">--</dd>
                            </div>
                            <div>
                                <dt>字段匹配</dt>
                                <dd id="doubaoMappingCount">--</dd>
                            </div>
                        </dl>
                        <pre id="doubaoResultJson">完成解析后，豆包返回结果会显示在这里。</pre>
                    </div>
                </details>
            </section>

            <section class="section-block mapping-section" aria-labelledby="mapping-title">
                <div class="section-heading">
                    <span class="step-number">4</span>
                    <div>
                        <h2 id="mapping-title">确认字段匹配</h2>
                        <p>检查模型建议，可通过下拉框修正源 Excel 列。</p>
                    </div>
                    <span id="matchSummary" class="match-summary">等待分析</span>
                </div>

                <div class="table-wrap">
                    <table>
                        <thead>
                            <tr>
                                <th>模板字段</th>
                                <th>源 Excel 列</th>
                                <th>置信度</th>
                                <th>状态</th>
                            </tr>
                        </thead>
                        <tbody id="mappingTableBody">
                            <tr>
                                <td class="empty-table" colspan="4">完成 Excel 解析后，匹配关系会显示在这里。</td>
                            </tr>
                        </tbody>
                    </table>
                </div>

                <div class="confirmation-bar">
                    <div>
                        <strong>匹配关系确认完成后，将按照当前设置生成新文件。</strong>
                        <p>原始 Excel 和模板文件不会被覆盖。</p>
                        <p id="generationMessage" class="generation-message" role="status"></p>
                    </div>
                    <button id="generateButton" class="confirm-button" type="button" disabled>
                        确认并生成 Excel
                    </button>
                </div>
            </section>
        </main>
    </div>

    <div
        id="loadingOverlay"
        class="loading-overlay"
        role="dialog"
        aria-modal="true"
        aria-labelledby="loadingTitle"
        hidden
    >
        <div class="loading-panel">
            <div class="loading-spinner" aria-hidden="true"></div>
            <div class="loading-copy">
                <p class="loading-kicker">正在处理</p>
                <h2 id="loadingTitle">正在读取 Excel</h2>
                <p id="loadingDescription">正在整理业务文件和模板文件，请稍候。</p>
            </div>
            <div class="loading-meta">
                <span id="loadingElapsed">已等待 0 秒</span>
                <span id="loadingHint">请不要关闭当前窗口</span>
            </div>
        </div>
    </div>
`;

const state = {
    apiKey: "",
    sourceFile: null,
    templateFile: null,
    sourceExcelData: null,
    templateExcelData: null,
    mappingResult: null,
};

const apiKeyInput = document.querySelector("#apiKey");
const sourceExcelInput = document.querySelector("#sourceExcel");
const templateExcelInput = document.querySelector("#templateExcel");
const sourceFileName = document.querySelector("#sourceFileName");
const templateFileName = document.querySelector("#templateFileName");
const parseButton = document.querySelector("#parseButton");
const formMessage = document.querySelector("#formMessage");
const resultDetails = document.querySelector("#resultDetails");
const resultSummaryTitle = document.querySelector("#resultSummaryTitle");
const resultSummaryMeta = document.querySelector("#resultSummaryMeta");
const resultFileName = document.querySelector("#resultFileName");
const resultSheetCount = document.querySelector("#resultSheetCount");
const resultRowCount = document.querySelector("#resultRowCount");
const resultJson = document.querySelector("#resultJson");
const doubaoResultDetails = document.querySelector("#doubaoResultDetails");
const doubaoResultSummaryTitle = document.querySelector("#doubaoResultSummaryTitle");
const doubaoResultSummaryMeta = document.querySelector("#doubaoResultSummaryMeta");
const doubaoSourceSheet = document.querySelector("#doubaoSourceSheet");
const doubaoTemplateSheet = document.querySelector("#doubaoTemplateSheet");
const doubaoMappingCount = document.querySelector("#doubaoMappingCount");
const doubaoResultJson = document.querySelector("#doubaoResultJson");
const matchSummary = document.querySelector("#matchSummary");
const mappingTableBody = document.querySelector("#mappingTableBody");
const generateButton = document.querySelector("#generateButton");
const generationMessage = document.querySelector("#generationMessage");
const loadingOverlay = document.querySelector("#loadingOverlay");
const loadingTitle = document.querySelector("#loadingTitle");
const loadingDescription = document.querySelector("#loadingDescription");
const loadingElapsed = document.querySelector("#loadingElapsed");
const loadingHint = document.querySelector("#loadingHint");

let loadingTimer = null;
let loadingStartedAt = 0;

function setFormMessage(message, type = "") {
    formMessage.textContent = message;
    formMessage.className = `form-message ${type}`.trim();
}

function setGenerationMessage(message, type = "") {
    generationMessage.textContent = message;
    generationMessage.className = `generation-message ${type}`.trim();
}

function updateLoadingTime() {
    const elapsedSeconds = Math.floor((Date.now() - loadingStartedAt) / 1000);

    loadingElapsed.textContent = `已等待 ${elapsedSeconds} 秒`;
    loadingHint.textContent = elapsedSeconds >= 30
        ? "数据较多时分析会更久，豆包仍在处理中"
        : "请不要关闭当前窗口";
}

function showLoading(title, description) {
    loadingTitle.textContent = title;
    loadingDescription.textContent = description;
    loadingStartedAt = Date.now();
    loadingOverlay.hidden = false;
    document.body.classList.add("is-loading");
    updateLoadingTime();

    window.clearInterval(loadingTimer);
    loadingTimer = window.setInterval(updateLoadingTime, 1000);
}

function updateLoading(title, description) {
    loadingTitle.textContent = title;
    loadingDescription.textContent = description;
}

function hideLoading() {
    window.clearInterval(loadingTimer);
    loadingTimer = null;
    loadingOverlay.hidden = true;
    document.body.classList.remove("is-loading");
}

function getTotalRowCount(excelData) {
    return excelData.sheets.reduce((total, sheet) => {
        return total + sheet.rows.length;
    }, 0);
}

function renderExcelResult(excelData) {
    const totalRows = getTotalRowCount(excelData);

    resultSummaryTitle.textContent = "查看业务 Excel 解析信息";
    resultSummaryMeta.textContent = `${excelData.sheets.length} 个工作表，${totalRows} 行非空数据`;
    resultFileName.textContent = excelData.fileName;
    resultSheetCount.textContent = `${excelData.sheets.length} 个`;
    resultRowCount.textContent = `${totalRows} 行`;
    resultJson.textContent = JSON.stringify(excelData, null, 2);
    resultDetails.open = false;
}

function escapeHtml(value) {
    return String(value)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function getSourceHeaders(excelData, mappingResult) {
    const sourceSheet = excelData.sheets.find((sheet) => {
        return sheet.sheetName === mappingResult.sourceSheetName;
    });
    const headerRow = sourceSheet?.rows.find((row) => {
        return row.rowIndex === mappingResult.sourceHeaderRow;
    });

    if (headerRow) {
        return headerRow.columns.map((header, index) => ({
            label: String(header || `第 ${index + 1} 列`),
            columnIndex: index,
        }));
    }

    const maxColumnCount = Math.max(
        0,
        ...(sourceSheet?.rows ?? []).map((row) => row.columns.length),
    );

    return Array.from({ length: maxColumnCount }, (_value, index) => ({
        label: `第 ${index + 1} 列`,
        columnIndex: index,
    }));
}

function getConfidenceDisplay(confidence) {
    const percentage = Math.round(confidence * 100);

    if (percentage >= 90) {
        return { percentage, className: "high", status: "已匹配", statusClass: "matched" };
    }

    if (percentage >= 60) {
        return { percentage, className: "medium", status: "建议确认", statusClass: "review" };
    }

    return { percentage, className: "empty", status: "未匹配", statusClass: "missing" };
}

function renderMappingResult(mappingResult) {
    const sourceHeaders = getSourceHeaders(state.sourceExcelData, mappingResult);
    const matchedCount = mappingResult.fieldMappings.filter((mapping) => {
        return mapping.sourceColumnIndex !== null;
    }).length;

    matchSummary.textContent = `${matchedCount} / ${mappingResult.fieldMappings.length} 已匹配`;

    mappingTableBody.innerHTML = mappingResult.fieldMappings.map((mapping, mappingIndex) => {
        const confidence = mapping.sourceColumnIndex === null
            ? getConfidenceDisplay(0)
            : getConfidenceDisplay(mapping.confidence);
        const options = [
            `<option value="">未匹配</option>`,
            ...sourceHeaders.map((header) => {
                const selected = header.columnIndex === mapping.sourceColumnIndex
                    ? " selected"
                    : "";

                return `<option value="${header.columnIndex}"${selected}>${escapeHtml(header.label)}</option>`;
            }),
        ].join("");

        return `
            <tr>
                <td>
                    <strong>${escapeHtml(mapping.targetField)}</strong>
                    <span class="mapping-reason">${escapeHtml(mapping.reason)}</span>
                </td>
                <td>
                    <select
                        data-mapping-index="${mappingIndex}"
                        aria-label="${escapeHtml(mapping.targetField)}对应源字段"
                    >
                        ${options}
                    </select>
                </td>
                <td>
                    <span class="confidence ${confidence.className}">
                        ${mapping.sourceColumnIndex === null ? "--" : `${confidence.percentage}%`}
                    </span>
                </td>
                <td>
                    <span class="row-status ${confidence.statusClass}">${confidence.status}</span>
                </td>
            </tr>
        `;
    }).join("");
    generateButton.disabled = false;
}

function renderDoubaoResult(mappingResult) {
    const matchedCount = mappingResult.fieldMappings.filter((mapping) => {
        return mapping.sourceColumnIndex !== null;
    }).length;

    doubaoResultSummaryTitle.textContent = "查看豆包原始分析结果";
    doubaoResultSummaryMeta.textContent =
        `${matchedCount} / ${mappingResult.fieldMappings.length} 个字段已匹配`;
    doubaoSourceSheet.textContent = mappingResult.sourceSheetName || "--";
    doubaoTemplateSheet.textContent = mappingResult.templateSheetName || "--";
    doubaoMappingCount.textContent = `${mappingResult.fieldMappings.length} 个`;
    doubaoResultJson.textContent = JSON.stringify(mappingResult, null, 2);
    doubaoResultDetails.open = false;
}

function resetMappingResult() {
    state.mappingResult = null;
    matchSummary.textContent = "等待分析";
    mappingTableBody.innerHTML = `
        <tr>
            <td class="empty-table" colspan="4">
                完成 Excel 解析后，匹配关系会显示在这里。
            </td>
        </tr>
    `;
    generateButton.disabled = true;
    setGenerationMessage("");
}

apiKeyInput.addEventListener("input", (event) => {
    state.apiKey = event.target.value.trim();
    setFormMessage("");
});

sourceExcelInput.addEventListener("change", (event) => {
    state.sourceFile = event.target.files[0] ?? null;
    state.sourceExcelData = null;
    resetMappingResult();
    sourceFileName.textContent = state.sourceFile
        ? state.sourceFile.name
        : "需要识别和转换的原始文件";
    setFormMessage("");
});

templateExcelInput.addEventListener("change", (event) => {
    state.templateFile = event.target.files[0] ?? null;
    state.templateExcelData = null;
    resetMappingResult();
    templateFileName.textContent = state.templateFile
        ? state.templateFile.name
        : "数据最终要填入的目标模板";
    setFormMessage("");
});

parseButton.addEventListener("click", async () => {
    state.apiKey = apiKeyInput.value.trim();

    if (!state.apiKey) {
        setFormMessage("请输入豆包 API Key。", "error");
        apiKeyInput.focus();
        return;
    }

    if (!state.sourceFile) {
        setFormMessage("请选择业务 Excel。", "error");
        return;
    }

    if (!state.templateFile) {
        setFormMessage("请选择模板 Excel。", "error");
        return;
    }

    parseButton.disabled = true;
    parseButton.textContent = "正在解析...";
    resetMappingResult();
    setFormMessage("正在读取 Excel 并请求豆包分析，请稍候。");
    showLoading(
        "正在读取 Excel",
        "正在整理业务文件和模板文件，请稍候。",
    );

    try {
        [state.sourceExcelData, state.templateExcelData] = await Promise.all([
            readExcel(state.sourceFile),
            readExcel(state.templateFile),
        ]);
        renderExcelResult(state.sourceExcelData);
        updateLoading(
            "正在等待豆包分析",
            "模型正在识别表头和字段匹配关系，数据越多等待时间可能越长。",
        );

        if (!window.excelAI?.matchExcelFields) {
            throw new Error("豆包请求服务未加载，请重新启动 Electron");
        }

        state.mappingResult = await window.excelAI.matchExcelFields({
            apiKey: state.apiKey,
            sourceExcel: state.sourceExcelData,
            templateExcel: state.templateExcelData,
        });
        renderDoubaoResult(state.mappingResult);
        renderMappingResult(state.mappingResult);
        setFormMessage("Excel 解析和字段匹配完成。", "success");
    } catch (error) {
        console.error(error);
        setFormMessage(`解析失败：${error.message}`, "error");
    } finally {
        hideLoading();
        parseButton.disabled = false;
        parseButton.textContent = "开始解析";
    }
});

mappingTableBody.addEventListener("change", (event) => {
    const select = event.target.closest("select[data-mapping-index]");

    if (!select || !state.mappingResult) {
        return;
    }

    const mapping = state.mappingResult.fieldMappings[
        Number(select.dataset.mappingIndex)
    ];

    if (!mapping) {
        return;
    }

    mapping.sourceColumnIndex = select.value === ""
        ? null
        : Number(select.value);
    mapping.sourceHeader = select.value === ""
        ? null
        : select.options[select.selectedIndex].textContent;

    renderMappingResult(state.mappingResult);
    renderDoubaoResult(state.mappingResult);
    setGenerationMessage("");
});

generateButton.addEventListener("click", async () => {
    if (!state.sourceFile || !state.templateFile || !state.mappingResult) {
        setGenerationMessage("请先完成 Excel 解析和字段匹配。", "error");
        return;
    }

    generateButton.disabled = true;
    generateButton.textContent = "正在生成...";
    setGenerationMessage("");
    showLoading(
        "正在生成 Excel",
        "正在按照确认后的字段关系填写模板并准备下载。",
    );

    try {
        const result = await generateExcel({
            sourceFile: state.sourceFile,
            templateFile: state.templateFile,
            mappingResult: state.mappingResult,
        });

        setGenerationMessage(
            `已生成 ${result.outputFileName}，共写入 ${result.writtenRowCount} 行数据。`,
            "success",
        );
    } catch (error) {
        console.error(error);
        setGenerationMessage(`生成失败：${error.message}`, "error");
    } finally {
        hideLoading();
        generateButton.disabled = false;
        generateButton.textContent = "确认并生成 Excel";
    }
});
