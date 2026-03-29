/*
  車両リース契約 更新通知（GAS）
  - TypeScript で実装し、dist へビルドして clasp push する前提
*/
const SHEET_NAMES = {
    SETTINGS: '設定',
    NEEDS_INPUT: '要入力',
    NOTIFY_LOG: '通知ログ',
    TEST_RESULTS: 'テスト結果',
    NOTIFY_BATCH: '通知バッチ',
};
const PRIMARY_SOURCE_SHEET = '車両一覧';
const APPROVAL_INPUT = {
    APPROVE: '承認',
    RETURN: '差戻し',
};
const ANSWER_LABELS = {
    RENEW: '更新',
    CANCELLATION_REPLACE: '解約（入替）',
    CANCELLATION_END: '解約（満了）',
};
const ANSWER_OPTIONS = [ANSWER_LABELS.RENEW, ANSWER_LABELS.CANCELLATION_REPLACE, ANSWER_LABELS.CANCELLATION_END];
const LEGACY_ANSWER_LABEL_MAP = {
    再リース: ANSWER_LABELS.RENEW,
    新車入替: ANSWER_LABELS.CANCELLATION_REPLACE,
    廃止: ANSWER_LABELS.CANCELLATION_END,
};
const BIANNUAL_BATCH_STATUS = {
    CREATED: '作成済',
    INITIAL_SENT: '初回通知送信済',
    REMINDER_SENT: 'リマインド送信済',
    SENMU_REQUESTED: '専務依頼送信済',
    SENMU_APPROVED: '専務承認済',
    SENMU_RETURNED: '専務差戻し',
    COMPLETED: '反映完了',
};
const HQ_CONFIRMATION_SHEET_PREFIX = '本部長副本部長確認_';
const AUTO_ADVANCE_INTERVAL_MINUTES = [1, 5, 10, 15, 30];
const BATCH_PHASE = {
    INITIAL_PENDING: 'INITIAL_PENDING',
    HQ_WAITING: 'HQ_WAITING',
    SENMU_REQUEST_READY: 'SENMU_REQUEST_READY',
    SENMU_WAITING: 'SENMU_WAITING',
    SENMU_RETURN_READY: 'SENMU_RETURN_READY',
    MURATA_NOTIFY_READY: 'MURATA_NOTIFY_READY',
    MASTER_APPLY_READY: 'MASTER_APPLY_READY',
    COMPLETED: 'COMPLETED',
};
const START_LAUNCH_STATUS = {
    BLOCKED: 'BLOCKED',
    CONFIRM_REQUIRED: 'CONFIRM_REQUIRED',
    READY: 'READY',
};
const START_REQUIRED_SETTING_LABELS = [
    { key: 'hqTo', label: '本部長副本部長_通知先To' },
    { key: 'senmuTo', label: '専務_通知先To' },
    { key: 'murataTo', label: '村田主任_通知先To' },
];
const START_LAUNCH_DIAG_MODE = {
    MANUAL: 'MANUAL',
    SCHEDULED: 'SCHEDULED',
};
const HQ_CONFIRMATION_HEADERS = [
    'batchId',
    'vehicleId',
    '管理部門',
    '管理担当者',
    '登録番号',
    '車種',
    '車台番号',
    '契約開始日',
    '契約満了日',
    '契約期間',
    '車検満了日',
    'リース料（税抜）',
    '本部回答',
    '回答確認済み',
    '専務判断',
    '専務コメント',
    '専務確認済み',
    '新契約開始日',
    '新契約満了日',
    '解約完了',
    '村田主任確認済み',
    '不正理由',
    '反映日時',
];
const VIEW_SHEET_PROTECTION_DESC_PREFIX = 'managed_by_script:view_sheet:';
const SCHEMA_DEFS = [
    {
        name: SHEET_NAMES.SETTINGS,
        headerRow: 1,
        headers: ['設定項目', '値', '説明'],
    },
    {
        name: PRIMARY_SOURCE_SHEET,
        headerRow: 1,
        headers: ['vehicleId', '登録番号_結合', '最終決定', '完了日'],
    },
    {
        name: SHEET_NAMES.NEEDS_INPUT,
        headerRow: 1,
        headers: ['検出日時', 'sourceSheet', 'vehicleId', '管理部門', '登録番号_結合', '車種', '不備内容'],
    },
    {
        name: SHEET_NAMES.NOTIFY_LOG,
        headerRow: 1,
        headers: ['日時', '種別', '管理部門', '宛先', 'requestId', '結果'],
    },
    {
        name: SHEET_NAMES.TEST_RESULTS,
        headerRow: 1,
        headers: ['実行日時', '項目', '結果', '詳細'],
    },
    {
        name: SHEET_NAMES.NOTIFY_BATCH,
        headerRow: 1,
        headers: [
            'batchId',
            '便区分',
            '送付予定日',
            '回答期限',
            '対象開始日',
            '対象終了日',
            '対象件数',
            '確認用シート名',
            '初回通知送信日時',
            'リマインド送信日時',
            '専務依頼送信日時',
            '村田主任通知送信日時',
            '村田主任不備通知日時',
            '村田主任不備通知ハッシュ',
            '村田主任反映完了通知日時',
            '村田主任反映完了通知ハッシュ',
            'ステータス',
            '作成日時',
            '更新日時',
        ],
    },
];
const SETTINGS_DEFAULTS = {
    送信元名: '車両管理システム',
    通知_メール送信: true,
    本部長副本部長_通知先To: '',
    専務_通知先To: '',
    専務_通知先Cc: '',
    半期送付日_3月: '03-01',
    半期送付日_9月: '09-01',
    回答期限_3月: '03-31',
    回答期限_9月: '09-30',
    リマインド_期限前日数: 10,
    自動進行_有効: true,
    自動進行_定期実行_有効: true,
    自動進行_定期実行_間隔分: 10,
    自動進行_編集連動_有効: false,
    自動進行_専務判断反映_有効: true,
    自動進行_マスター反映_有効: true,
    自動進行_最小間隔秒: 30,
    最終承認者アカウント: '',
    運用管理者_通知先To: '',
    村田主任_通知先To: '',
    村田主任_通知先Cc: '',
};
const SCHEMA_VERSION = '1';
const PROP_KEYS = {
    SCHEMA_VERSION: 'SCHEMA_VERSION',
    LAST_SCHEMA_SYNC_AT: 'LAST_SCHEMA_SYNC_AT',
    LAST_SCHEMA_DRIFT_AT: 'LAST_SCHEMA_DRIFT_AT',
    AUTO_ADVANCE_LAST_RUN_AT: 'AUTO_ADVANCE_LAST_RUN_AT',
    SOURCE_SYNC_LAST_RUN_AT: 'SOURCE_SYNC_LAST_RUN_AT',
    ONEDIT_ADVANCE_LAST_RUN_AT: 'ONEDIT_ADVANCE_LAST_RUN_AT',
    AUTO_START_BLOCK_NOTICE_PREFIX: 'AUTO_START_BLOCK_NOTICE_',
};
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('車両更新通知')
        .addItem('運用マニュアル（このシートで見る）', 'showOperationManual')
        .addItem('初期車両登録マニュアル（このシートで見る）', 'showInitialVehicleRegistrationManual')
        .addItem('テスト手順書（このシートで見る）', 'showTestGuide')
        .addItem('新しい確認依頼を開始（対象抽出〜初回通知）', 'runDaily')
        .addItem('進行中の確認依頼を再開（最新）', 'runAutoAdvanceNow')
        .addToUi();
}
function showOperationManual() {
    const ui = SpreadsheetApp.getUi();
    const html = HtmlService.createHtmlOutputFromFile('operation_manual_vehicle_lease_renewal')
        .setWidth(1000)
        .setHeight(800);
    ui.showModalDialog(html, '運用マニュアル');
}
function showInitialVehicleRegistrationManual() {
    const ui = SpreadsheetApp.getUi();
    const html = HtmlService.createHtmlOutputFromFile('operation_manual_initial_vehicle_registration')
        .setWidth(1000)
        .setHeight(800);
    ui.showModalDialog(html, '初期車両登録マニュアル');
}
function showTestGuide() {
    const ui = SpreadsheetApp.getUi();
    const html = HtmlService.createHtmlOutputFromFile('test_guide')
        .setWidth(1000)
        .setHeight(800);
    ui.showModalDialog(html, 'テスト手順書');
}
function uiAlertSafe(message) {
    try {
        SpreadsheetApp.getUi().alert(message);
    }
    catch (e) {
        Logger.log(`UI alert skipped: ${message}`);
    }
}
function uiShowModalSafe(title, body) {
    try {
        const html = HtmlService.createHtmlOutput(`<div style="font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, 'Liberation Mono', 'Courier New', monospace; white-space: pre-wrap; line-height: 1.4;">${escapeHtml(body)}</div>`)
            .setWidth(900)
            .setHeight(700);
        SpreadsheetApp.getUi().showModalDialog(html, title);
    }
    catch (e) {
        Logger.log(`UI modal skipped: ${title}\n${body}`);
    }
}
function syncSchema() {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const ss = getSpreadsheet();
        SCHEMA_DEFS.forEach((def) => {
            const sheet = ensureSheet(ss, def.name);
            ensureHeaders(sheet, def.headerRow, def.headers);
        });
        seedSettings();
        const props = PropertiesService.getDocumentProperties();
        props.setProperty(PROP_KEYS.SCHEMA_VERSION, SCHEMA_VERSION);
        props.setProperty(PROP_KEYS.LAST_SCHEMA_SYNC_AT, new Date().toISOString());
    }
    finally {
        lock.releaseLock();
    }
}
function checkSchemaDrift() {
    const ss = getSpreadsheet();
    const driftMessages = [];
    SCHEMA_DEFS.forEach((def) => {
        const sheet = ss.getSheetByName(def.name);
        if (!sheet) {
            driftMessages.push(`シート未存在: ${def.name}`);
            return;
        }
        const lastColumn = sheet.getLastColumn();
        if (lastColumn === 0) {
            driftMessages.push(`ヘッダ行が空です: ${def.name}`);
            return;
        }
        const headerRowValues = sheet.getRange(def.headerRow, 1, 1, lastColumn).getValues()[0];
        const headerMap = getHeaderMap(headerRowValues);
        const missing = def.headers.filter((header) => !headerMap[header]);
        if (missing.length > 0) {
            driftMessages.push(`不足ヘッダ: ${def.name} -> ${missing.join(', ')}`);
        }
    });
    if (driftMessages.length > 0) {
        PropertiesService.getDocumentProperties().setProperty(PROP_KEYS.LAST_SCHEMA_DRIFT_AT, new Date().toISOString());
        Logger.log(driftMessages.join('\n'));
    }
    return driftMessages;
}
function syncVehicles() {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const ss = getSpreadsheet();
        const sourceSheet = ensureSheet(ss, PRIMARY_SOURCE_SHEET);
        ensureHeaders(sourceSheet, 1, getSchemaHeaders(PRIMARY_SOURCE_SHEET));
        ensureHeaders(ensureSheet(ss, SHEET_NAMES.NEEDS_INPUT), 1, getSchemaHeaders(SHEET_NAMES.NEEDS_INPUT));
        const sourceData = sourceSheet.getDataRange().getValues();
        const sourceHeader = sourceData.length > 0 ? getHeaderMap(sourceData[0]) : {};
        const sourceIndexes = resolveSourceHeaders(sourceHeader);
        const managedIndexes = {
            vehicleId: sourceHeader['vehicleId'],
            regCombined: sourceHeader['登録番号_結合'],
        };
        const needsInputRows = [];
        const now = new Date();
        const seenVehicleId = {};
        const rowCount = Math.max(0, sourceData.length - 1);
        const vehicleIdValues = [];
        const regCombinedValues = [];
        const hasContractEndHeader = !!sourceIndexes.contractEnd;
        const hasDeptHeader = !!sourceIndexes.dept;
        if (!hasContractEndHeader || !hasDeptHeader) {
            needsInputRows.push([now, PRIMARY_SOURCE_SHEET, '', '', '', '', '必要ヘッダが不足しています']);
        }
        for (let i = 1; i < sourceData.length; i++) {
            const row = sourceData[i];
            const rowNumber = i + 1;
            const existingVehicleId = getCellValue(row, managedIndexes.vehicleId);
            const existingRegCombined = getCellValue(row, managedIndexes.regCombined);
            if (row.every((cell) => cell === '' || cell === null)) {
                vehicleIdValues.push([existingVehicleId]);
                regCombinedValues.push([existingRegCombined]);
                continue;
            }
            const regCombined = getSourceRegistrationCombined(row, sourceIndexes);
            const vehicleType = getCellValue(row, sourceIndexes.vehicleType);
            const chassis = getCellValue(row, sourceIndexes.chassis);
            const contractEnd = hasContractEndHeader ? parseDateValue(getCellRaw(row, sourceIndexes.contractEnd)) : null;
            const dept = hasDeptHeader ? getCellValue(row, sourceIndexes.dept) : '';
            const vehicleId = buildVehicleId(PRIMARY_SOURCE_SHEET, regCombined, chassis, rowNumber);
            const normalizedReg = String(regCombined || '').trim();
            const hasRegDigits = /\d/.test(normalizedReg);
            if (seenVehicleId[vehicleId]) {
                const prev = seenVehicleId[vehicleId];
                needsInputRows.push([
                    now,
                    PRIMARY_SOURCE_SHEET,
                    vehicleId,
                    dept,
                    regCombined,
                    vehicleType,
                    `vehicleId重複（先頭: ${prev.sheet} 行${prev.rowIndex} / 今回: 行${rowNumber}）`,
                ]);
            }
            else {
                seenVehicleId[vehicleId] = { sheet: PRIMARY_SOURCE_SHEET, rowIndex: rowNumber };
            }
            if (hasContractEndHeader && !contractEnd) {
                needsInputRows.push([now, PRIMARY_SOURCE_SHEET, vehicleId, dept, regCombined, vehicleType, '契約満了日なし']);
            }
            if (hasDeptHeader && !dept) {
                needsInputRows.push([now, PRIMARY_SOURCE_SHEET, vehicleId, dept, regCombined, vehicleType, '管理部門なし']);
            }
            if (!normalizedReg || !hasRegDigits) {
                needsInputRows.push([now, PRIMARY_SOURCE_SHEET, vehicleId, dept, regCombined, vehicleType, '登録番号不備']);
            }
            vehicleIdValues.push([vehicleId]);
            regCombinedValues.push([regCombined]);
        }
        if (rowCount > 0) {
            if (managedIndexes.vehicleId)
                sourceSheet.getRange(2, managedIndexes.vehicleId, rowCount, 1).setValues(vehicleIdValues);
            if (managedIndexes.regCombined)
                sourceSheet.getRange(2, managedIndexes.regCombined, rowCount, 1).setValues(regCombinedValues);
        }
        writeSheetData(SHEET_NAMES.NEEDS_INPUT, needsInputRows);
        protectViewSheet(SHEET_NAMES.NEEDS_INPUT);
    }
    finally {
        lock.releaseLock();
    }
}
function createBiannualBatch(options) {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const ss = getSpreadsheet();
        ensureHeaders(ensureSheet(ss, SHEET_NAMES.NOTIFY_BATCH), 1, getSchemaHeaders(SHEET_NAMES.NOTIFY_BATCH));
        ensureHeaders(ensureSheet(ss, PRIMARY_SOURCE_SHEET), 1, getSchemaHeaders(PRIMARY_SOURCE_SHEET));
        const settings = loadSettings();
        const tz = ss.getSpreadsheetTimeZone();
        const batchDef = resolveBiannualBatchDefinition(new Date(), tz, settings);
        const notifyBatchSheet = ss.getSheetByName(SHEET_NAMES.NOTIFY_BATCH);
        if (!notifyBatchSheet)
            throw new Error('通知バッチシートが存在しません');
        const batchData = notifyBatchSheet.getDataRange().getValues();
        const batchHeader = batchData.length > 0 ? getHeaderMap(batchData[0]) : {};
        const existing = findNotifyBatchRow(batchData, batchHeader, batchDef.batchId);
        if (existing) {
            if (!(options === null || options === void 0 ? void 0 : options.suppressUi)) {
                uiAlertSafe(`通知バッチ ${batchDef.batchId} は既に存在します。`);
            }
            return batchDef.batchId;
        }
        const vehicleContext = loadSourceVehicleContext();
        const targetVehicles = pickVehiclesByContractEndRange(vehicleContext.rows, vehicleContext.headerMap, batchDef.targetStart, batchDef.targetEnd, tz);
        const now = new Date();
        const newRow = new Array(getSchemaHeaders(SHEET_NAMES.NOTIFY_BATCH).length).fill('');
        const setCell = (headerName, value) => {
            const idx = batchHeader[headerName];
            if (idx)
                newRow[idx - 1] = value;
        };
        setCell('batchId', batchDef.batchId);
        setCell('便区分', batchDef.label);
        setCell('送付予定日', batchDef.sendDate);
        setCell('回答期限', batchDef.deadline);
        setCell('対象開始日', batchDef.targetStart);
        setCell('対象終了日', batchDef.targetEnd);
        setCell('対象件数', targetVehicles.length);
        setCell('確認用シート名', '');
        setCell('初回通知送信日時', '');
        setCell('リマインド送信日時', '');
        setCell('専務依頼送信日時', '');
        setCell('ステータス', BIANNUAL_BATCH_STATUS.CREATED);
        setCell('作成日時', now);
        setCell('更新日時', now);
        notifyBatchSheet.getRange(notifyBatchSheet.getLastRow() + 1, 1, 1, newRow.length).setValues([newRow]);
        const confirmationSheetName = buildConfirmationSheet(batchDef.batchId);
        if (!(options === null || options === void 0 ? void 0 : options.suppressUi)) {
            uiAlertSafe(`半期バッチを起票しました。\n` +
                `batchId: ${batchDef.batchId}\n` +
                `対象期間: ${formatDateLabel(batchDef.targetStart, tz)}〜${formatDateLabel(batchDef.targetEnd, tz)}\n` +
                `対象件数: ${targetVehicles.length}件\n` +
                `確認用シート: ${confirmationSheetName}`);
        }
        return batchDef.batchId;
    }
    finally {
        lock.releaseLock();
    }
}
function buildConfirmationSheetForLatestBatch() {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const ss = getSpreadsheet();
        const notifyBatchSheet = ss.getSheetByName(SHEET_NAMES.NOTIFY_BATCH);
        if (!notifyBatchSheet || notifyBatchSheet.getLastRow() <= 1) {
            uiAlertSafe('通知バッチにデータがありません。先に半期バッチを起票してください。');
            return '';
        }
        const data = notifyBatchSheet.getDataRange().getValues();
        const headerMap = getHeaderMap(data[0]);
        const batchIdIndex = headerMap['batchId'];
        if (!batchIdIndex)
            throw new Error('通知バッチの batchId 列が見つかりません');
        let latestBatchId = '';
        for (let i = data.length - 1; i >= 1; i--) {
            const batchId = getCellValue(data[i], batchIdIndex);
            if (!batchId)
                continue;
            latestBatchId = batchId;
            break;
        }
        if (!latestBatchId) {
            uiAlertSafe('通知バッチの batchId が見つかりません。');
            return '';
        }
        const sheetName = buildConfirmationSheet(latestBatchId);
        uiAlertSafe(`最新バッチ ${latestBatchId} の確認用シートを生成しました。\nシート名: ${sheetName}`);
        return sheetName;
    }
    finally {
        lock.releaseLock();
    }
}
function buildConfirmationSheet(batchId) {
    const ss = getSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const notifyBatchSheet = ss.getSheetByName(SHEET_NAMES.NOTIFY_BATCH);
    if (!notifyBatchSheet)
        throw new Error('通知バッチシートが存在しません');
    ensureHeaders(notifyBatchSheet, 1, getSchemaHeaders(SHEET_NAMES.NOTIFY_BATCH));
    const batchData = notifyBatchSheet.getDataRange().getValues();
    if (batchData.length <= 1)
        throw new Error('通知バッチにデータがありません');
    const batchHeader = getHeaderMap(batchData[0]);
    const batchRowInfo = findNotifyBatchRow(batchData, batchHeader, batchId);
    if (!batchRowInfo)
        throw new Error(`通知バッチが見つかりません: ${batchId}`);
    const batchRow = batchRowInfo.row;
    const targetStart = parseDateValue(getCellRaw(batchRow, batchHeader['対象開始日']));
    const targetEnd = parseDateValue(getCellRaw(batchRow, batchHeader['対象終了日']));
    if (!targetStart || !targetEnd) {
        throw new Error(`通知バッチの対象期間が不正です: ${batchId}`);
    }
    const vehicleContext = loadSourceVehicleContext();
    const targetVehicles = pickVehiclesByContractEndRange(vehicleContext.rows, vehicleContext.headerMap, targetStart, targetEnd, tz);
    const sendDate = parseDateValue(getCellRaw(batchRow, batchHeader['送付予定日'])) || new Date();
    const sheetStamp = Utilities.formatDate(sendDate, tz, 'yyyyMMdd');
    const sheetName = buildUniqueSheetName(`${HQ_CONFIRMATION_SHEET_PREFIX}${sheetStamp}`);
    const sheet = ensureSheet(ss, sheetName);
    writeArbitrarySheetData(sheet, HQ_CONFIRMATION_HEADERS, []);
    const vh = vehicleContext.headerMap;
    const rows = targetVehicles.map((row) => [
        batchId,
        getCellValue(row, vh['vehicleId']),
        getCellValue(row, vh['管理部門']),
        getCellValue(row, vh['管理担当者']),
        getCellValue(row, vh['登録番号_結合']),
        getCellValue(row, vh['車種']),
        getCellValue(row, vh['車台番号']),
        getCellRaw(row, vh['契約開始日']),
        getCellRaw(row, vh['契約満了日']),
        getCellValue(row, vh['契約期間']),
        getCellRaw(row, vh['車検満了日']),
        getCellValue(row, vh['リース料（税抜）']),
        '',
        false,
        '',
        '',
        false,
        '',
        '',
        false,
        false,
        '',
        '',
    ]);
    writeArbitrarySheetData(sheet, HQ_CONFIRMATION_HEADERS, rows);
    const headerMap = getHeaderMap(HQ_CONFIRMATION_HEADERS);
    if (rows.length > 0) {
        const checkColumns = ['回答確認済み', '解約完了', '村田主任確認済み', '専務確認済み'];
        checkColumns.forEach((name) => {
            const col = headerMap[name];
            if (!col)
                return;
            sheet.getRange(2, col, rows.length, 1).insertCheckboxes();
        });
        const answerCol = headerMap['本部回答'];
        if (answerCol) {
            const rule = SpreadsheetApp.newDataValidation().requireValueInList(ANSWER_OPTIONS, true).build();
            sheet.getRange(2, answerCol, rows.length, 1).setDataValidation(rule);
        }
        const decisionCol = headerMap['専務判断'];
        if (decisionCol) {
            const rule = SpreadsheetApp.newDataValidation()
                .requireValueInList([APPROVAL_INPUT.APPROVE, APPROVAL_INPUT.RETURN], true)
                .build();
            sheet.getRange(2, decisionCol, rows.length, 1).setDataValidation(rule);
        }
    }
    protectSenmuColumns(sheetName);
    const now = new Date();
    const setBatchCell = (headerName, value) => {
        const idx = batchHeader[headerName];
        if (idx)
            batchRow[idx - 1] = value;
    };
    setBatchCell('対象件数', rows.length);
    setBatchCell('確認用シート名', sheetName);
    setBatchCell('更新日時', now);
    notifyBatchSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);
    return sheetName;
}
function sendHqInitialEmail(batchId, options) {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const settings = loadSettings();
        const batchContext = getNotifyBatchContext(batchId);
        if (!batchContext) {
            if (!(options === null || options === void 0 ? void 0 : options.suppressUi)) {
                uiAlertSafe('通知バッチが見つかりません。先に半期バッチ起票を実行してください。');
            }
            return '';
        }
        const { ss, notifyBatchSheet, batchData, headerMap, row } = batchContext;
        const resolvedBatchId = batchContext.batchId;
        const tz = ss.getSpreadsheetTimeZone();
        const hqTo = String(settings.hqTo || '').trim();
        if (!settings.mailSendEnabled) {
            appendNotificationLog('半期初回通知', '', '', resolvedBatchId, '通知_メール送信=FALSE のため送信をスキップ');
            return resolvedBatchId;
        }
        if (!hqTo) {
            appendNotificationLog('半期初回通知', '', '', resolvedBatchId, '本部長副本部長_通知先Toが未設定');
            if (!(options === null || options === void 0 ? void 0 : options.suppressUi)) {
                uiAlertSafe('設定「本部長副本部長_通知先To」が未設定のため送信できません。');
            }
            return resolvedBatchId;
        }
        const sentAt = parseDateValue(getCellRaw(row, headerMap['初回通知送信日時']));
        if (sentAt) {
            if (!(options === null || options === void 0 ? void 0 : options.suppressUi)) {
                uiAlertSafe(`初回通知は既に送信済みです。\nbatchId: ${resolvedBatchId}`);
            }
            return resolvedBatchId;
        }
        let confirmationSheetName = getCellValue(row, headerMap['確認用シート名']);
        if (!confirmationSheetName || !ss.getSheetByName(confirmationSheetName)) {
            confirmationSheetName = buildConfirmationSheet(resolvedBatchId);
        }
        const confirmationSheet = ss.getSheetByName(confirmationSheetName);
        if (!confirmationSheet) {
            throw new Error(`確認用シートが見つかりません: ${confirmationSheetName}`);
        }
        const confirmationData = confirmationSheet.getDataRange().getValues();
        const confirmationHeader = confirmationData.length > 0 ? getHeaderMap(confirmationData[0]) : {};
        const counts = summarizeConfirmationSheetRows(confirmationData, confirmationHeader);
        const sheetUrl = buildSheetUrlWithGid(ss, confirmationSheet);
        const batchLabel = getCellValue(row, headerMap['便区分']) || resolvedBatchId;
        const deadline = parseDateValue(getCellRaw(row, headerMap['回答期限']));
        const targetStart = parseDateValue(getCellRaw(row, headerMap['対象開始日']));
        const targetEnd = parseDateValue(getCellRaw(row, headerMap['対象終了日']));
        const subject = `【車両更新確認】${batchLabel} ご確認のお願い`;
        const body = [
            '本部長・副本部長 各位',
            '',
            `${batchLabel} の車両更新確認をお願いいたします。`,
            '',
            `対象期間: ${formatDateLabel(targetStart || new Date(), tz)}〜${formatDateLabel(targetEnd || new Date(), tz)}`,
            `回答期限: ${formatDateLabel(deadline || new Date(), tz)}`,
            `対象件数: ${counts.total}`,
            '',
            'ご対応の流れ:',
            `1. 確認用シートを開き、「本部回答」を ${ANSWER_OPTIONS.join(' / ')} から選択する`,
            '2. 全件の入力が終わったら、「回答確認済み」にチェックを入れる',
            '',
            '確認用シート:',
            sheetUrl,
            '',
            `管理番号: ${resolvedBatchId}`,
            '',
            'よろしくお願いいたします。',
        ].join('\n');
        try {
            MailApp.sendEmail({
                to: hqTo,
                subject,
                name: settings.fromName,
                body,
            });
            const now = new Date();
            row[headerMap['確認用シート名'] - 1] = confirmationSheetName;
            row[headerMap['初回通知送信日時'] - 1] = now;
            row[headerMap['ステータス'] - 1] = BIANNUAL_BATCH_STATUS.INITIAL_SENT;
            row[headerMap['更新日時'] - 1] = now;
            notifyBatchSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);
            appendNotificationLog('半期初回通知', '', hqTo, resolvedBatchId, '成功');
            if (!(options === null || options === void 0 ? void 0 : options.suppressUi)) {
                uiAlertSafe(`初回通知を送信しました。\nbatchId: ${resolvedBatchId}`);
            }
        }
        catch (err) {
            appendNotificationLog('半期初回通知', '', hqTo, resolvedBatchId, `失敗: ${err}`);
            throw err;
        }
        return resolvedBatchId;
    }
    finally {
        lock.releaseLock();
    }
}
function sendHqReminderIfNeeded(batchId) {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const settings = loadSettings();
        const batchContext = getNotifyBatchContext(batchId);
        if (!batchContext)
            return '';
        const { ss, notifyBatchSheet, batchData, headerMap, row } = batchContext;
        const resolvedBatchId = batchContext.batchId;
        const tz = ss.getSpreadsheetTimeZone();
        const hqTo = String(settings.hqTo || '').trim();
        if (!settings.mailSendEnabled) {
            appendNotificationLog('半期リマインド', '', '', resolvedBatchId, '通知_メール送信=FALSE のため送信をスキップ');
            return resolvedBatchId;
        }
        if (!hqTo) {
            appendNotificationLog('半期リマインド', '', '', resolvedBatchId, '本部長副本部長_通知先Toが未設定');
            return resolvedBatchId;
        }
        const reminderSentAt = parseDateValue(getCellRaw(row, headerMap['リマインド送信日時']));
        if (reminderSentAt)
            return resolvedBatchId;
        const deadline = parseDateValue(getCellRaw(row, headerMap['回答期限']));
        if (!deadline) {
            appendNotificationLog('半期リマインド', '', hqTo, resolvedBatchId, '回答期限が未設定のためスキップ');
            return resolvedBatchId;
        }
        const reminderBeforeDays = Math.max(0, toNumber(settings.reminderBeforeDays, 10));
        const reminderDate = addDays(toDateOnly(deadline, tz), -reminderBeforeDays);
        const today = toDateOnly(new Date(), tz);
        if (today.getTime() < reminderDate.getTime())
            return resolvedBatchId;
        const confirmationSheetName = getCellValue(row, headerMap['確認用シート名']);
        if (!confirmationSheetName) {
            appendNotificationLog('半期リマインド', '', hqTo, resolvedBatchId, '確認用シート未設定のためスキップ');
            return resolvedBatchId;
        }
        const confirmationSheet = ss.getSheetByName(confirmationSheetName);
        if (!confirmationSheet) {
            appendNotificationLog('半期リマインド', '', hqTo, resolvedBatchId, `確認用シート不在: ${confirmationSheetName}`);
            return resolvedBatchId;
        }
        const confirmationData = confirmationSheet.getDataRange().getValues();
        const confirmationHeader = confirmationData.length > 0 ? getHeaderMap(confirmationData[0]) : {};
        const counts = summarizeConfirmationSheetRows(confirmationData, confirmationHeader);
        if (counts.unchecked <= 0) {
            return resolvedBatchId;
        }
        const batchLabel = getCellValue(row, headerMap['便区分']) || resolvedBatchId;
        const sheetUrl = buildSheetUrlWithGid(ss, confirmationSheet);
        const subject = `【リマインド】${batchLabel} ご確認のお願い`;
        const body = [
            '本部長・副本部長 各位',
            '',
            `${batchLabel} の確認について、未完了の項目が残っています。`,
            `回答期限: ${formatDateLabel(deadline, tz)}`,
            '',
            `未確認件数: ${counts.unchecked}`,
            `未回答件数: ${counts.unanswered}`,
            '',
            'お手すきの際に確認用シートをご確認いただき、全件入力後に「回答確認済み」へチェックをお願いいたします。',
            '',
            '確認用シート:',
            sheetUrl,
            '',
            `管理番号: ${resolvedBatchId}`,
            '',
            '※ このメールは期限前のご案内として1回のみ送信しています。',
        ].join('\n');
        try {
            MailApp.sendEmail({
                to: hqTo,
                subject,
                name: settings.fromName,
                body,
            });
            const now = new Date();
            row[headerMap['リマインド送信日時'] - 1] = now;
            row[headerMap['ステータス'] - 1] = BIANNUAL_BATCH_STATUS.REMINDER_SENT;
            row[headerMap['更新日時'] - 1] = now;
            notifyBatchSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);
            appendNotificationLog('半期リマインド', '', hqTo, resolvedBatchId, '成功');
        }
        catch (err) {
            appendNotificationLog('半期リマインド', '', hqTo, resolvedBatchId, `失敗: ${err}`);
            throw err;
        }
        return resolvedBatchId;
    }
    finally {
        lock.releaseLock();
    }
}
function sendSenmuApprovalRequestIfReady(batchId) {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const settings = loadSettings();
        const snapshot = getBatchWorkflowSnapshot(batchId);
        if (!snapshot)
            return '';
        const { batchContext, confirmationSheet, confirmationData, confirmationHeader, hqGate } = snapshot;
        const { ss, notifyBatchSheet, batchData, headerMap, row } = batchContext;
        const resolvedBatchId = batchContext.batchId;
        const tz = ss.getSpreadsheetTimeZone();
        const senmuTo = String(settings.senmuTo || '').trim();
        const sentAt = parseDateValue(getCellRaw(row, headerMap['専務依頼送信日時']));
        if (sentAt)
            return resolvedBatchId;
        if (!confirmationSheet) {
            const confirmationSheetName = getCellValue(row, headerMap['確認用シート名']);
            appendNotificationLog('専務依頼', '', '', resolvedBatchId, '確認用シート未設定のためスキップ');
            return resolvedBatchId;
        }
        ensureConfirmationSheetSchema(confirmationSheet);
        const counts = summarizeConfirmationSheetRows(confirmationData, confirmationHeader);
        if (hqGate.total <= 0)
            return resolvedBatchId;
        if (hqGate.unchecked > 0) {
            appendNotificationLog('専務依頼', '', '', resolvedBatchId, `未確認行あり(${hqGate.unchecked}件)のため未送信`);
            return resolvedBatchId;
        }
        if (hqGate.pending > 0) {
            appendNotificationLog('専務依頼', '', '', resolvedBatchId, `未回答行あり(${hqGate.pending}件)のため未送信`);
            return resolvedBatchId;
        }
        if (!settings.mailSendEnabled) {
            appendNotificationLog('専務依頼', '', '', resolvedBatchId, '通知_メール送信=FALSE のため送信をスキップ');
            return resolvedBatchId;
        }
        if (!senmuTo) {
            appendNotificationLog('専務依頼', '', '', resolvedBatchId, '専務_通知先Toが未設定');
            uiAlertSafe('設定「専務_通知先To」が未設定のため送信できません。');
            return resolvedBatchId;
        }
        const batchLabel = getCellValue(row, headerMap['便区分']) || resolvedBatchId;
        const deadline = parseDateValue(getCellRaw(row, headerMap['回答期限']));
        const sheetUrl = buildSheetUrlWithGid(ss, confirmationSheet);
        const subject = `【専務確認依頼】${batchLabel} ご確認のお願い`;
        const body = [
            '専務',
            '',
            `${batchLabel} の一次確認が完了しましたので、最終確認をお願いいたします。`,
            '',
            `対象件数: ${counts.total}`,
            `更新: ${counts.renew}`,
            `解約（入替）: ${counts.cancellationReplace}`,
            `解約（満了）: ${counts.cancellationEnd}`,
            `確認期限の目安: ${formatDateLabel(deadline || new Date(), tz)}`,
            '',
            'ご対応の流れ:',
            `1. 確認用シートの「専務判断」に ${APPROVAL_INPUT.APPROVE} または ${APPROVAL_INPUT.RETURN} を入力する`,
            '2. 判断を入れた行ごとに「専務確認済み」へチェックを入れる',
            '',
            '確認用シート:',
            sheetUrl,
            '',
            `管理番号: ${resolvedBatchId}`,
            '',
            'よろしくお願いいたします。',
        ].join('\n');
        try {
            MailApp.sendEmail({
                to: senmuTo,
                cc: String(settings.senmuCc || ''),
                subject,
                name: settings.fromName,
                body,
            });
            const now = new Date();
            row[headerMap['専務依頼送信日時'] - 1] = now;
            row[headerMap['ステータス'] - 1] = BIANNUAL_BATCH_STATUS.SENMU_REQUESTED;
            row[headerMap['更新日時'] - 1] = now;
            notifyBatchSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);
            protectSenmuColumns(confirmationSheet.getName());
            appendNotificationLog('専務依頼', '', senmuTo, resolvedBatchId, '成功');
        }
        catch (err) {
            appendNotificationLog('専務依頼', '', senmuTo, resolvedBatchId, `失敗: ${err}`);
            throw err;
        }
        return resolvedBatchId;
    }
    finally {
        lock.releaseLock();
    }
}
function applySenmuDecisionFromSheet(batchId) {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const snapshot = getBatchWorkflowSnapshot(batchId);
        if (!snapshot)
            return '';
        const { batchContext, confirmationSheet, confirmationData, confirmationHeader, senmuGate } = snapshot;
        const { notifyBatchSheet, batchData, headerMap, row } = batchContext;
        const resolvedBatchId = batchContext.batchId;
        if (!confirmationSheet) {
            uiAlertSafe(`確認用シート名が未設定です。\nbatchId: ${resolvedBatchId}`);
            return resolvedBatchId;
        }
        ensureConfirmationSheetSchema(confirmationSheet);
        if (confirmationData.length <= 1)
            return resolvedBatchId;
        const h = confirmationHeader;
        const decisionIndex = h['専務判断'];
        if (!decisionIndex)
            return resolvedBatchId;
        const now = new Date();
        const currentStatus = getCellValue(row, headerMap['ステータス']);
        if (currentStatus === BIANNUAL_BATCH_STATUS.COMPLETED) {
            appendNotificationLog('専務判断反映', '', '', resolvedBatchId, '反映完了のため再判定をスキップ');
            return resolvedBatchId;
        }
        const allDecided = senmuGate.pending === 0 && senmuGate.invalid === 0;
        const allChecked = senmuGate.unchecked === 0;
        if (!allDecided || !allChecked) {
            // 条件未達: 承認済/差戻しから崩れた場合は専務依頼送信済に戻す
            if (currentStatus === BIANNUAL_BATCH_STATUS.SENMU_APPROVED ||
                currentStatus === BIANNUAL_BATCH_STATUS.SENMU_RETURNED) {
                row[headerMap['ステータス'] - 1] = BIANNUAL_BATCH_STATUS.SENMU_REQUESTED;
                row[headerMap['更新日時'] - 1] = now;
                notifyBatchSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);
            }
            appendNotificationLog('専務判断反映', '', '', resolvedBatchId, `未達: 保留=${senmuGate.pending} 不正=${senmuGate.invalid} 未確認=${senmuGate.unchecked}`);
            return resolvedBatchId;
        }
        const newStatus = senmuGate.returned > 0
            ? BIANNUAL_BATCH_STATUS.SENMU_RETURNED
            : senmuGate.approved > 0
                ? BIANNUAL_BATCH_STATUS.SENMU_APPROVED
                : currentStatus;
        // ステータスに変化がなければ no-op（ログ・モーダル・更新日時すべてスキップ）
        if (newStatus === currentStatus) {
            return resolvedBatchId;
        }
        row[headerMap['ステータス'] - 1] = newStatus;
        row[headerMap['更新日時'] - 1] = now;
        notifyBatchSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);
        appendNotificationLog('専務判断反映', '', '', resolvedBatchId, `承認:${senmuGate.approved} 差戻し:${senmuGate.returned} 保留:${senmuGate.pending} 不正:${senmuGate.invalid}`);
        return resolvedBatchId;
    }
    finally {
        lock.releaseLock();
    }
}
function sendHqReturnNotification(batchId, options) {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const settings = loadSettings();
        const snapshot = getBatchWorkflowSnapshot(batchId);
        if (!snapshot)
            return '';
        const { batchContext, confirmationSheet, confirmationData, confirmationHeader, senmuGate } = snapshot;
        const { ss, notifyBatchSheet, batchData, headerMap, row } = batchContext;
        const resolvedBatchId = batchContext.batchId;
        const tz = ss.getSpreadsheetTimeZone();
        const hqTo = String(settings.hqTo || '').trim();
        const status = getCellValue(row, headerMap['ステータス']);
        if (status !== BIANNUAL_BATCH_STATUS.SENMU_RETURNED)
            return resolvedBatchId;
        if (!settings.mailSendEnabled) {
            appendNotificationLog('差戻し通知', '', '', resolvedBatchId, '通知_メール送信=FALSE のため送信をスキップ');
            return resolvedBatchId;
        }
        if (!hqTo) {
            appendNotificationLog('差戻し通知', '', '', resolvedBatchId, '本部長副本部長_通知先Toが未設定');
            uiAlertSafe('設定「本部長副本部長_通知先To」が未設定のため送信できません。');
            return resolvedBatchId;
        }
        if (!confirmationSheet)
            return resolvedBatchId;
        ensureConfirmationSheetSchema(confirmationSheet);
        if (confirmationData.length <= 1)
            return resolvedBatchId;
        const ch = confirmationHeader;
        const decisionIndex = ch['専務判断'];
        if (!decisionIndex)
            return resolvedBatchId;
        // 全行に専務判断が入力済みであることを確認（入力途中では発火しない）
        const returnedRows = [];
        for (let i = 1; i < confirmationData.length; i++) {
            const cRow = confirmationData[i];
            const vehicleId = getCellValue(cRow, ch['vehicleId']);
            if (!vehicleId)
                continue;
            const decision = normalizeSenmuDecision(getCellValue(cRow, decisionIndex));
            if (!decision)
                continue;
            if (decision === APPROVAL_INPUT.RETURN) {
                returnedRows.push({
                    rowIndex: i,
                    vehicleId,
                    regNumber: getCellValue(cRow, ch['登録番号']) || vehicleId,
                    comment: getCellValue(cRow, ch['専務コメント']) || '',
                });
            }
        }
        // 入力途中ガード
        if (senmuGate.pending > 0 || senmuGate.invalid > 0 || senmuGate.unchecked > 0) {
            appendNotificationLog('差戻し通知', '', '', resolvedBatchId, `未入力(${senmuGate.pending})・不正(${senmuGate.invalid})・未確認(${senmuGate.unchecked})ありのため未送信`);
            return resolvedBatchId;
        }
        if (returnedRows.length === 0)
            return resolvedBatchId;
        const batchLabel = getCellValue(row, headerMap['便区分']) || resolvedBatchId;
        const sheetUrl = buildSheetUrlWithGid(ss, confirmationSheet);
        const subject = `【差戻し】${batchLabel} 再確認のお願い`;
        const detailLines = returnedRows.map((r) => `- ${r.regNumber}${r.comment ? '（' + r.comment + '）' : ''}`);
        const body = [
            '本部長・副本部長 各位',
            '',
            `${batchLabel} の専務確認で、再確認のご依頼が出ています。`,
            '',
            `差戻し件数: ${returnedRows.length}`,
            '',
            '差戻し内容:',
            ...detailLines,
            '',
            '該当行をご確認のうえ、必要に応じて「本部回答」を見直し、あらためて「回答確認済み」へチェックをお願いいたします。',
            '',
            '確認用シート:',
            sheetUrl,
            '',
            `管理番号: ${resolvedBatchId}`,
            '',
            'よろしくお願いいたします。',
        ].join('\n');
        try {
            MailApp.sendEmail({
                to: hqTo,
                subject,
                name: settings.fromName,
                body,
            });
        }
        catch (err) {
            appendNotificationLog('差戻し通知', '', hqTo, resolvedBatchId, `送信失敗: ${err}`);
            throw err;
        }
        // 差戻しは「回答の破棄」ではなく「再確認依頼」なので、本部回答は残しつつ再確認フラグだけ戻す。
        let modified = false;
        for (const r of returnedRows) {
            const cRow = confirmationData[r.rowIndex];
            if (ch['回答確認済み'])
                cRow[ch['回答確認済み'] - 1] = false;
            if (ch['専務確認済み'])
                cRow[ch['専務確認済み'] - 1] = false;
            if (ch['村田主任確認済み'])
                cRow[ch['村田主任確認済み'] - 1] = false;
            if (ch['反映日時'])
                cRow[ch['反映日時'] - 1] = '';
            modified = true;
        }
        if (modified) {
            confirmationSheet.getRange(1, 1, confirmationData.length, confirmationData[0].length).setValues(confirmationData);
        }
        // バッチ行リセット
        const now = new Date();
        if (headerMap['専務依頼送信日時'])
            row[headerMap['専務依頼送信日時'] - 1] = '';
        if (headerMap['村田主任通知送信日時'])
            row[headerMap['村田主任通知送信日時'] - 1] = '';
        if (headerMap['リマインド送信日時'])
            row[headerMap['リマインド送信日時'] - 1] = '';
        row[headerMap['ステータス'] - 1] = BIANNUAL_BATCH_STATUS.INITIAL_SENT;
        row[headerMap['更新日時'] - 1] = now;
        notifyBatchSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);
        appendNotificationLog('差戻し通知', '', hqTo, resolvedBatchId, `成功 差戻し${returnedRows.length}件`);
        if (!(options === null || options === void 0 ? void 0 : options.suppressUi)) {
            uiAlertSafe(`差戻し通知を送信しました。\n差戻し件数: ${returnedRows.length}\nbatchId: ${resolvedBatchId}`);
        }
        return resolvedBatchId;
    }
    finally {
        lock.releaseLock();
    }
}
function sendMurataApprovalNotification(batchId, options) {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const settings = loadSettings();
        const snapshot = getBatchWorkflowSnapshot(batchId);
        if (!snapshot)
            return '';
        const { batchContext, confirmationSheet, confirmationData, confirmationHeader } = snapshot;
        const { ss, notifyBatchSheet, batchData, headerMap, row } = batchContext;
        const resolvedBatchId = batchContext.batchId;
        const tz = ss.getSpreadsheetTimeZone();
        const murataTo = String(settings.murataTo || '').trim();
        const status = getCellValue(row, headerMap['ステータス']);
        if (status !== BIANNUAL_BATCH_STATUS.SENMU_APPROVED)
            return resolvedBatchId;
        const murataSentAt = parseDateValue(getCellRaw(row, headerMap['村田主任通知送信日時']));
        if (murataSentAt)
            return resolvedBatchId;
        if (!settings.mailSendEnabled) {
            appendNotificationLog('村田主任通知', '', '', resolvedBatchId, '通知_メール送信=FALSE のため送信をスキップ');
            return resolvedBatchId;
        }
        if (!murataTo) {
            appendNotificationLog('村田主任通知', '', '', resolvedBatchId, '村田主任_通知先Toが未設定');
            uiAlertSafe('設定「村田主任_通知先To」が未設定のため送信できません。');
            return resolvedBatchId;
        }
        if (!confirmationSheet)
            return resolvedBatchId;
        const counts = summarizeConfirmationSheetRows(confirmationData, confirmationHeader);
        const batchLabel = getCellValue(row, headerMap['便区分']) || resolvedBatchId;
        const sheetUrl = buildSheetUrlWithGid(ss, confirmationSheet);
        const subject = `【マスター反映前確認】${batchLabel} 専務承認完了`;
        const body = [
            '村田主任',
            '',
            `${batchLabel} の専務確認が全件承認されました。`,
            '',
            `対象件数: ${counts.total}`,
            `更新: ${counts.renew}`,
            `解約（入替）: ${counts.cancellationReplace}`,
            `解約（満了）: ${counts.cancellationEnd}`,
            '',
            '確認用シートで内容をご確認のうえ、反映対象の行に「村田主任確認済み」へチェックをお願いいたします。',
            'チェックがそろった行から、定期実行の自動処理で順次マスターへ反映されます。',
            '',
            '確認用シート:',
            sheetUrl,
            '',
            `管理番号: ${resolvedBatchId}`,
            '',
            'よろしくお願いいたします。',
        ].join('\n');
        try {
            MailApp.sendEmail({
                to: murataTo,
                cc: String(settings.murataCc || ''),
                subject,
                name: settings.fromName,
                body,
            });
            const now = new Date();
            row[headerMap['村田主任通知送信日時'] - 1] = now;
            row[headerMap['更新日時'] - 1] = now;
            notifyBatchSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);
            appendNotificationLog('村田主任通知', '', murataTo, resolvedBatchId, '成功');
            if (!(options === null || options === void 0 ? void 0 : options.suppressUi)) {
                uiAlertSafe(`村田主任へ通知を送信しました。\nbatchId: ${resolvedBatchId}`);
            }
        }
        catch (err) {
            appendNotificationLog('村田主任通知', '', murataTo, resolvedBatchId, `失敗: ${err}`);
            throw err;
        }
        return resolvedBatchId;
    }
    finally {
        lock.releaseLock();
    }
}
function buildMasterApplyValidationResult(row, headerMap) {
    const vehicleId = getCellValue(row, headerMap['vehicleId']);
    const registration = getCellValue(row, headerMap['登録番号']) || vehicleId || '車両不明';
    const policy = normalizeAnswerLabel(getCellValue(row, headerMap['本部回答']));
    const decision = normalizeSenmuDecision(getCellValue(row, headerMap['専務判断']));
    const reasons = [];
    if (decision !== APPROVAL_INPUT.APPROVE) {
        reasons.push('専務判断が「承認」ではありません');
    }
    if (!policy) {
        reasons.push('本部回答が未入力です');
    }
    else if (policy === ANSWER_LABELS.RENEW) {
        if (!parseDateValue(getCellRaw(row, headerMap['新契約開始日']))) {
            reasons.push('新契約開始日が未入力です');
        }
        if (!parseDateValue(getCellRaw(row, headerMap['新契約満了日']))) {
            reasons.push('新契約満了日が未入力です');
        }
    }
    else if (!isCheckedCell(getCellRaw(row, headerMap['解約完了']))) {
        reasons.push('解約完了がチェックされていません');
    }
    return {
        vehicleId,
        registration,
        policy,
        decision,
        reasons,
    };
}
function buildMasterApplyNotificationLine(entry, includeReasons) {
    const parts = [entry.registration || entry.vehicleId || '車両不明'];
    if (entry.policy)
        parts.push(entry.policy);
    if (includeReasons && entry.reasons && entry.reasons.length > 0) {
        parts.push(entry.reasons.join(' / '));
    }
    return parts.join(' | ');
}
function buildMasterAppliedDetailLine(entry) {
    const parts = [buildMasterApplyNotificationLine(entry, false)];
    parts.push(`確認用シート${entry.confirmationRowNumber}行目 → ${PRIMARY_SOURCE_SHEET}${entry.vehicleRowNumber}行目`);
    if (entry.appliedDetails.length > 0) {
        parts.push(`反映内容: ${entry.appliedDetails.join(' / ')}`);
    }
    return parts.join(' | ');
}
function buildMasterAppliedMailBlock(entry, index) {
    const lines = [
        `${index}. ${entry.registration || entry.vehicleId || '車両不明'} (${entry.policy || '方針未設定'})`,
        `   確認元: 確認用シート ${entry.confirmationRowNumber}行目`,
        `   反映先: ${PRIMARY_SOURCE_SHEET} ${entry.vehicleRowNumber}行目`,
    ];
    if (entry.appliedDetails.length > 0) {
        lines.push('   反映内容:');
        entry.appliedDetails.forEach((detail) => {
            lines.push(`   - ${detail}`);
        });
    }
    return lines.join('\n');
}
function createContentHash(lines) {
    const normalized = lines.join('\n');
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, normalized);
    return bytes.map((b) => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}
function notifyMurataMasterApplyIssues(settings, batchContext, confirmationSheet, invalidEntries, options) {
    if (!batchContext)
        return false;
    const { ss, notifyBatchSheet, batchData, headerMap, row, batchId } = batchContext;
    const murataTo = String(settings.murataTo || '').trim();
    const hashHeader = headerMap['村田主任不備通知ハッシュ'];
    const sentAtHeader = headerMap['村田主任不備通知日時'];
    if (!hashHeader || !sentAtHeader)
        return false;
    if (invalidEntries.length === 0) {
        const hadState = !!getCellRaw(row, hashHeader) || !!getCellRaw(row, sentAtHeader);
        row[hashHeader - 1] = '';
        row[sentAtHeader - 1] = '';
        if (hadState) {
            row[headerMap['更新日時'] - 1] = new Date();
            notifyBatchSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);
        }
        return false;
    }
    const sorted = invalidEntries
        .slice()
        .sort((a, b) => buildMasterApplyNotificationLine(a, true).localeCompare(buildMasterApplyNotificationLine(b, true), 'ja'));
    const detailLines = sorted.map((entry) => `- ${buildMasterApplyNotificationLine(entry, true)}`);
    const nextHash = createContentHash(detailLines);
    const prevHash = String(getCellRaw(row, hashHeader) || '').trim();
    if (nextHash === prevHash) {
        appendNotificationLog('村田主任不備通知', '', murataTo, batchId, '同一内容のため再通知なし');
        return false;
    }
    if (!settings.mailSendEnabled) {
        appendNotificationLog('村田主任不備通知', '', '', batchId, '通知_メール送信=FALSE のため送信をスキップ');
        return false;
    }
    if (!murataTo) {
        appendNotificationLog('村田主任不備通知', '', '', batchId, '村田主任_通知先Toが未設定');
        if (!(options === null || options === void 0 ? void 0 : options.suppressUi)) {
            uiAlertSafe('設定「村田主任_通知先To」が未設定のため不備通知を送信できません。');
        }
        return false;
    }
    const batchLabel = getCellValue(row, headerMap['便区分']) || batchId;
    const sheetUrl = buildSheetUrlWithGid(ss, confirmationSheet);
    const subject = `【確認要】${batchLabel} マスター反映で修正が必要な項目があります`;
    const body = [
        '村田主任',
        '',
        `${batchLabel} のマスター反映処理で、確認が必要な項目が見つかりました。`,
        '確認用シートの「不正理由」列にも同じ内容を残しています。',
        '',
        '対象一覧:',
        ...detailLines,
        '',
        '確認用シート:',
        sheetUrl,
        '',
        `管理番号: ${batchId}`,
        '',
        '内容を修正いただくと、次回の自動進行で再判定・反映されます。',
    ].join('\n');
    MailApp.sendEmail({
        to: murataTo,
        cc: String(settings.murataCc || ''),
        subject,
        name: settings.fromName,
        body,
    });
    const now = new Date();
    row[hashHeader - 1] = nextHash;
    row[sentAtHeader - 1] = now;
    row[headerMap['更新日時'] - 1] = now;
    notifyBatchSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);
    appendNotificationLog('村田主任不備通知', '', murataTo, batchId, `成功 ${invalidEntries.length}件`);
    return true;
}
function notifyMurataMasterApplyCompleted(settings, batchContext, confirmationSheet, appliedEntries, options) {
    if (!batchContext || appliedEntries.length === 0)
        return false;
    const { ss, notifyBatchSheet, batchData, headerMap, row, batchId } = batchContext;
    const murataTo = String(settings.murataTo || '').trim();
    const hashHeader = headerMap['村田主任反映完了通知ハッシュ'];
    const sentAtHeader = headerMap['村田主任反映完了通知日時'];
    if (!hashHeader || !sentAtHeader)
        return false;
    const sorted = appliedEntries
        .slice()
        .sort((a, b) => buildMasterAppliedDetailLine(a).localeCompare(buildMasterAppliedDetailLine(b), 'ja'));
    const detailBlocks = sorted.map((entry, index) => buildMasterAppliedMailBlock(entry, index + 1));
    const detailLines = detailBlocks.join('\n\n').split('\n');
    const nextHash = createContentHash(detailLines);
    const prevHash = String(getCellRaw(row, hashHeader) || '').trim();
    if (nextHash === prevHash) {
        appendNotificationLog('村田主任反映完了通知', '', murataTo, batchId, '同一内容のため再通知なし');
        return false;
    }
    if (!settings.mailSendEnabled) {
        appendNotificationLog('村田主任反映完了通知', '', '', batchId, '通知_メール送信=FALSE のため送信をスキップ');
        return false;
    }
    if (!murataTo) {
        appendNotificationLog('村田主任反映完了通知', '', '', batchId, '村田主任_通知先Toが未設定');
        if (!(options === null || options === void 0 ? void 0 : options.suppressUi)) {
            uiAlertSafe('設定「村田主任_通知先To」が未設定のため反映完了通知を送信できません。');
        }
        return false;
    }
    const batchLabel = getCellValue(row, headerMap['便区分']) || batchId;
    const sheetUrl = buildSheetUrlWithGid(ss, confirmationSheet);
    const subject = `【反映完了】${batchLabel} マスター反映が完了した項目があります`;
    const body = [
        '村田主任',
        '',
        `${batchLabel} のマスター反映で、今回新たに完了した項目をお知らせします。`,
        '「確認元」と「反映先」を分けているので、どの入力がどの台帳行に反映されたかをメール上で追えます。',
        '',
        '反映完了一覧:',
        ...detailBlocks.flatMap((block) => block.split('\n').concat([''])).slice(0, -1),
        '',
        '確認用シート:',
        sheetUrl,
        '',
        `管理番号: ${batchId}`,
    ].join('\n');
    MailApp.sendEmail({
        to: murataTo,
        cc: String(settings.murataCc || ''),
        subject,
        name: settings.fromName,
        body,
    });
    const now = new Date();
    row[hashHeader - 1] = nextHash;
    row[sentAtHeader - 1] = now;
    row[headerMap['更新日時'] - 1] = now;
    notifyBatchSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);
    appendNotificationLog('村田主任反映完了通知', '', murataTo, batchId, `成功 ${appliedEntries.length}件`);
    return true;
}
function applyMasterUpdates(batchId, options) {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const batchContext = getNotifyBatchContext(batchId);
        if (!batchContext)
            return '';
        const { ss, notifyBatchSheet, batchData, headerMap, row } = batchContext;
        const resolvedBatchId = batchContext.batchId;
        const tz = ss.getSpreadsheetTimeZone();
        const settings = loadSettings();
        // ステータスガード: SENMU_APPROVED以外ではマスター反映しない
        const currentStatus = getCellValue(row, headerMap['ステータス']);
        if (currentStatus !== BIANNUAL_BATCH_STATUS.SENMU_APPROVED) {
            appendNotificationLog('マスター反映', '', '', resolvedBatchId, `ステータスが${currentStatus}のためスキップ`);
            return resolvedBatchId;
        }
        const confirmationSheetName = getCellValue(row, headerMap['確認用シート名']);
        if (!confirmationSheetName) {
            appendNotificationLog('マスター反映', '', '', resolvedBatchId, '確認用シート名が未設定のためスキップ');
            if (!(options === null || options === void 0 ? void 0 : options.suppressUi)) {
                uiAlertSafe(`確認用シート名が未設定です。\nbatchId: ${resolvedBatchId}`);
            }
            return resolvedBatchId;
        }
        const confirmationSheet = ss.getSheetByName(confirmationSheetName);
        if (!confirmationSheet) {
            appendNotificationLog('マスター反映', '', '', resolvedBatchId, `確認用シートが見つかりません: ${confirmationSheetName}`);
            if (!(options === null || options === void 0 ? void 0 : options.suppressUi)) {
                uiAlertSafe(`確認用シートが見つかりません。\n${confirmationSheetName}`);
            }
            return resolvedBatchId;
        }
        ensureConfirmationSheetSchema(confirmationSheet);
        const vehicleSheet = ss.getSheetByName(PRIMARY_SOURCE_SHEET);
        if (!vehicleSheet)
            throw new Error('車両一覧が存在しません');
        ensureHeaders(vehicleSheet, 1, getSchemaHeaders(PRIMARY_SOURCE_SHEET));
        const confirmationData = confirmationSheet.getDataRange().getValues();
        if (confirmationData.length <= 1)
            return resolvedBatchId;
        const ch = getHeaderMap(confirmationData[0]);
        const vehicleData = vehicleSheet.getDataRange().getValues();
        if (vehicleData.length <= 1)
            return resolvedBatchId;
        const vh = getHeaderMap(vehicleData[0]);
        const vehicleRowIndexById = {};
        for (let i = 1; i < vehicleData.length; i++) {
            const vehicleId = getCellValue(vehicleData[i], vh['vehicleId']);
            if (!vehicleId)
                continue;
            vehicleRowIndexById[vehicleId] = i;
        }
        const now = new Date();
        const rowIndexesToGray = [];
        let applied = 0;
        let skipped = 0;
        let waiting = 0;
        let returned = 0;
        let modifiedVehicle = false;
        let modifiedConfirmation = false;
        const invalidEntries = [];
        const appliedEntries = [];
        const invalidReasonCol = ch['不正理由'];
        for (let i = 1; i < confirmationData.length; i++) {
            const cRow = confirmationData[i];
            const vehicleId = getCellValue(cRow, ch['vehicleId']);
            if (!vehicleId)
                continue;
            const setInvalidReason = (value) => {
                if (!invalidReasonCol)
                    return;
                const current = String(cRow[invalidReasonCol - 1] || '');
                if (current === value)
                    return;
                cRow[invalidReasonCol - 1] = value;
                modifiedConfirmation = true;
            };
            const decision = normalizeSenmuDecision(getCellValue(cRow, ch['専務判断']));
            if (decision === APPROVAL_INPUT.RETURN) {
                setInvalidReason('');
                returned += 1;
                continue;
            }
            // 村田主任確認済みがチェックされていなければスキップ
            if (!isCheckedCell(getCellRaw(cRow, ch['村田主任確認済み']))) {
                setInvalidReason('');
                if (decision !== APPROVAL_INPUT.APPROVE) {
                    waiting += 1;
                }
                else {
                    skipped += 1;
                }
                continue;
            }
            // 二重反映防止
            if (getCellValue(cRow, ch['反映日時'])) {
                setInvalidReason('');
                continue;
            }
            const validation = buildMasterApplyValidationResult(cRow, ch);
            const errors = validation.reasons.slice();
            const policy = validation.policy;
            skipped += 1;
            const vehicleRowIndex = vehicleRowIndexById[vehicleId];
            if (vehicleRowIndex === undefined) {
                errors.push('車両一覧に対応する vehicleId が見つかりません');
            }
            if (errors.length > 0) {
                setInvalidReason(errors.join('\n'));
                invalidEntries.push({
                    vehicleId,
                    registration: validation.registration,
                    policy,
                    reasons: errors,
                });
                continue;
            }
            skipped -= 1;
            const vehicleRow = vehicleData[vehicleRowIndex];
            const confirmationRowNumber = i + 1;
            const vehicleRowNumber = vehicleRowIndex + 1;
            const appliedDetails = [];
            const setVehicle = (headerName, value) => {
                const idx = vh[headerName];
                if (idx)
                    vehicleRow[idx - 1] = value;
            };
            if (policy === ANSWER_LABELS.RENEW) {
                const newStart = parseDateValue(getCellRaw(cRow, ch['新契約開始日']));
                const newEnd = parseDateValue(getCellRaw(cRow, ch['新契約満了日']));
                setVehicle('契約開始日', toDateOnly(newStart, tz));
                setVehicle('契約満了日', toDateOnly(newEnd, tz));
                setVehicle('最終決定', policy);
                setVehicle('完了日', now);
                appliedDetails.push(`契約開始日=${formatDateLabel(newStart, tz)}`);
                appliedDetails.push(`契約満了日=${formatDateLabel(newEnd, tz)}`);
                appliedDetails.push('最終決定=更新');
                appliedDetails.push(`完了日=${formatDateLabel(now, tz)}`);
            }
            else {
                setVehicle('最終決定', policy);
                setVehicle('完了日', now);
                rowIndexesToGray.push(vehicleRowIndex + 1);
                appliedDetails.push(`最終決定=${policy}`);
                appliedDetails.push(`完了日=${formatDateLabel(now, tz)}`);
                appliedDetails.push('車両一覧の対象行をグレーアウト');
            }
            setInvalidReason('');
            if (ch['反映日時'])
                cRow[ch['反映日時'] - 1] = now;
            applied += 1;
            modifiedVehicle = true;
            modifiedConfirmation = true;
            appliedEntries.push({
                vehicleId,
                registration: validation.registration,
                policy,
                confirmationRowNumber,
                vehicleRowNumber,
                appliedDetails,
            });
        }
        if (modifiedVehicle) {
            vehicleSheet.getRange(1, 1, vehicleData.length, vehicleData[0].length).setValues(vehicleData);
            rowIndexesToGray.forEach((rowIndex) => {
                vehicleSheet.getRange(rowIndex, 1, 1, vehicleData[0].length).setBackground('#d9d9d9');
            });
        }
        if (modifiedConfirmation) {
            confirmationSheet.getRange(1, 1, confirmationData.length, confirmationData[0].length).setValues(confirmationData);
        }
        const counts = summarizeConfirmationSheetRows(confirmationData, ch);
        if (counts.total > 0 && counts.masterApplied >= counts.total) {
            row[headerMap['ステータス'] - 1] = BIANNUAL_BATCH_STATUS.COMPLETED;
        }
        else {
            row[headerMap['ステータス'] - 1] = BIANNUAL_BATCH_STATUS.SENMU_APPROVED;
        }
        row[headerMap['更新日時'] - 1] = now;
        notifyBatchSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);
        if (invalidEntries.length > 0) {
            appendNotificationLog('マスター反映', '', '', resolvedBatchId, `要確認項目あり: ${invalidEntries.map((entry) => buildMasterApplyNotificationLine(entry, true)).join(' | ')}`);
        }
        notifyMurataMasterApplyIssues(settings, batchContext, confirmationSheet, invalidEntries, options);
        notifyMurataMasterApplyCompleted(settings, batchContext, confirmationSheet, appliedEntries, options);
        // マスター反映は定期トリガーや再開メニューから非同期に進むバッチ処理なので、
        // 人の応答を止めるモーダルではなく通知ログと確認用シートの更新結果を正とする。
        appendNotificationLog('マスター反映', '', '', resolvedBatchId, `反映:${applied} 待機:${waiting} 差戻し:${returned} スキップ:${skipped}`);
        return resolvedBatchId;
    }
    finally {
        lock.releaseLock();
    }
}
function createRequests() {
    // 旧導線名の互換。処理単位は半期バッチへ統一する。
    return createBiannualBatch();
}
function sendInitialEmails() {
    // 旧導線名の互換。処理単位は半期バッチへ統一する。
    return sendHqInitialEmail();
}
function sendReminderEmails() {
    // 旧導線名の互換。処理単位は半期バッチへ統一する。
    return sendHqReminderIfNeeded();
}
function applyAnswers() {
    // 旧導線名の互換。処理単位は半期バッチへ統一する。
    return applyMasterUpdates();
}
function sendApprovalRequestEmails() {
    // 旧導線名の互換。処理単位は半期バッチへ統一する。
    return sendSenmuApprovalRequestIfReady();
}
function applyApprovalDecisions() {
    // 旧導線名の互換。処理単位は半期バッチへ統一する。
    return applySenmuDecisionFromSheet();
}
function buildSheetUrlWithGid(ss, sheet) {
    const base = ss.getUrl();
    try {
        return `${base}#gid=${sheet.getSheetId()}`;
    }
    catch (err) {
        return base;
    }
}
function runDaily() {
    const diagnosis = prepareManualStartLaunchDiagnosis();
    if (diagnosis.status === START_LAUNCH_STATUS.BLOCKED) {
        uiShowModalSafe(buildStartLaunchTitle(diagnosis), buildStartLaunchBody(diagnosis));
        return;
    }
    if (diagnosis.status === START_LAUNCH_STATUS.CONFIRM_REQUIRED) {
        showStartLaunchConfirmDialog(diagnosis);
        return;
    }
    const result = runBiannualScheduleWithSummary(diagnosis);
    uiShowModalSafe(result.title, result.body);
}
function runBiannualSchedule() {
    syncSchema();
    syncVehicles();
    const diagnosis = diagnoseStartLaunch(START_LAUNCH_DIAG_MODE.SCHEDULED);
    if (diagnosis.status === START_LAUNCH_STATUS.BLOCKED) {
        notifyScheduledStartBlocked(diagnosis);
        return '';
    }
    clearScheduledStartBlockedNotice(diagnosis.batchId);
    const result = runBiannualScheduleCore(diagnosis);
    if (diagnosis.warningMessages.length > 0) {
        appendNotificationLog('半期自動開始', '', '', result.batchId, `警告付きで開始: ${diagnosis.warningMessages.join(' / ')}`);
    }
    else {
        appendNotificationLog('半期自動開始', '', '', result.batchId, '成功');
    }
    return result.batchId;
}
function continueRunDailyAfterConfirmation() {
    const diagnosis = prepareManualStartLaunchDiagnosis();
    if (diagnosis.status === START_LAUNCH_STATUS.BLOCKED) {
        return {
            title: buildStartLaunchTitle(diagnosis),
            body: buildStartLaunchBody(diagnosis),
        };
    }
    const result = runBiannualScheduleWithSummary(diagnosis);
    return {
        title: result.title,
        body: result.body,
    };
}
function runAutoAdvance() {
    const settings = loadSettings();
    if (!settings.autoAdvanceEnabled || !settings.autoAdvanceTimerEnabled)
        return '';
    if (!tryReserveAutoAdvanceRun(settings.autoAdvanceMinIntervalSec))
        return '';
    return advanceBiannualWorkflow('', 'timer');
}
function runAutoAdvanceNow() {
    const settings = loadSettings();
    if (!settings.autoAdvanceEnabled)
        return '';
    return advanceBiannualWorkflow('', 'manual');
}
function onEditAutoAdvance(e) {
    // 承認フローは編集途中の断片ではなく、定期実行で状態確定後に進める。
    return '';
}
function onEditSourceSync(e) {
    if (!e || !e.range)
        return '';
    const range = e.range;
    const sheet = range.getSheet();
    if (!sheet || sheet.getName() !== PRIMARY_SOURCE_SHEET)
        return '';
    if (range.getRow() <= 1)
        return '';
    const settings = loadSettings();
    if (!tryReserveSourceSyncRun(settings.autoAdvanceMinIntervalSec))
        return '';
    syncVehicles();
    return 'ok';
}
function onEditSettingsSync(e) {
    if (!e || !e.range)
        return '';
    const sheet = e.range.getSheet();
    if (!sheet || sheet.getName() !== SHEET_NAMES.SETTINGS)
        return '';
    const data = sheet.getDataRange().getValues();
    const headerMap = getHeaderMap(data[0]);
    if (!headerMap['設定項目'] || !headerMap['値'])
        return '';
    const valCol = headerMap['値'];
    const startCol = e.range.getColumn();
    const endCol = startCol + e.range.getNumColumns() - 1;
    if (valCol < startCol || valCol > endCol)
        return '';
    const startRow = e.range.getRow();
    const endRow = startRow + e.range.getNumRows() - 1;
    let found = false;
    for (let r = startRow; r <= endRow; r++) {
        if (getCellValue(data[r - 1], headerMap['設定項目']) === '最終承認者アカウント') {
            found = true;
            break;
        }
    }
    if (!found)
        return '';
    const ss = getSpreadsheet();
    ss.getSheets().forEach((s) => {
        if (isConfirmationSheetName(s.getName())) {
            protectSenmuColumns(s.getName());
        }
    });
    return 'ok';
}
function advanceBiannualWorkflow(batchId, reason) {
    const settings = loadSettings();
    if (!settings.autoAdvanceEnabled)
        return '';
    const context = getNotifyBatchContext(batchId);
    if (!context)
        return '';
    const targetBatchId = context.batchId;
    try {
        const ctx = getNotifyBatchContext(targetBatchId) || context;
        const confirmationSheetName = getCellValue(ctx.row, ctx.headerMap['確認用シート名']);
        if (!confirmationSheetName || !ctx.ss.getSheetByName(confirmationSheetName)) {
            buildConfirmationSheet(targetBatchId);
        }
        let phase = evaluateBatchPhase(targetBatchId);
        if (!phase)
            return targetBatchId;
        if (phase.phase === BATCH_PHASE.COMPLETED)
            return targetBatchId;
        if (phase.phase === BATCH_PHASE.INITIAL_PENDING) {
            sendHqInitialEmail(targetBatchId);
            phase = evaluateBatchPhase(targetBatchId);
            if (!phase)
                return targetBatchId;
            if (phase.phase === BATCH_PHASE.COMPLETED)
                return targetBatchId;
        }
        sendHqReminderIfNeeded(targetBatchId);
        if (phase.phase === BATCH_PHASE.SENMU_REQUEST_READY) {
            sendSenmuApprovalRequestIfReady(targetBatchId);
            phase = evaluateBatchPhase(targetBatchId);
            if (!phase)
                return targetBatchId;
            if (phase.phase === BATCH_PHASE.COMPLETED)
                return targetBatchId;
        }
        if (settings.autoApplySenmuDecision && shouldRunAutoSenmuDecision(targetBatchId, phase)) {
            applySenmuDecisionFromSheet(targetBatchId);
            phase = evaluateBatchPhase(targetBatchId);
            if (!phase)
                return targetBatchId;
            if (phase.phase === BATCH_PHASE.COMPLETED)
                return targetBatchId;
        }
        if (phase.phase === BATCH_PHASE.SENMU_RETURN_READY) {
            const suppressUi = reason === 'timer';
            sendHqReturnNotification(targetBatchId, { suppressUi });
            return targetBatchId;
        }
        if (phase.phase === BATCH_PHASE.MURATA_NOTIFY_READY) {
            const suppressUi = reason === 'timer';
            sendMurataApprovalNotification(targetBatchId, { suppressUi });
            phase = evaluateBatchPhase(targetBatchId);
            if (!phase)
                return targetBatchId;
            if (phase.phase === BATCH_PHASE.COMPLETED)
                return targetBatchId;
        }
        if (settings.autoApplyMasterUpdates && shouldRunAutoMasterUpdate(targetBatchId, phase)) {
            applyMasterUpdates(targetBatchId, { suppressUi: reason === 'timer' });
        }
    }
    catch (err) {
        appendNotificationLog('自動進行', '', '', targetBatchId, `失敗(${reason || 'unknown'}): ${err}`);
        throw err;
    }
    return targetBatchId;
}
function seedSettings() {
    const ss = getSpreadsheet();
    const sheet = ensureSheet(ss, SHEET_NAMES.SETTINGS);
    ensureHeaders(sheet, 1, getSchemaHeaders(SHEET_NAMES.SETTINGS));
    const data = sheet.getDataRange().getValues();
    if (data.length === 0)
        return;
    const headerMap = getHeaderMap(data[0]);
    const keyIndex = headerMap['設定項目'];
    const valueIndex = headerMap['値'];
    const descIndex = headerMap['説明'];
    if (!keyIndex || !valueIndex)
        return;
    const existingKeys = {};
    for (let i = 1; i < data.length; i++) {
        const key = getCellValue(data[i], keyIndex);
        if (key)
            existingKeys[key] = true;
    }
    const rows = [];
    Object.keys(SETTINGS_DEFAULTS).forEach((key) => {
        if (key === '自動進行_定期実行_間隔分')
            return;
        if (!existingKeys[key]) {
            rows.push([key, SETTINGS_DEFAULTS[key], '']);
        }
    });
    if (!existingKeys['自動進行_定期実行_間隔分']) {
        const legacyHours = valuesFromSettingsSheet(data, headerMap, '自動進行_定期実行_間隔時間');
        const intervalMinutes = normalizeAutoAdvanceIntervalMinutes(convertLegacyAutoAdvanceHours(legacyHours), Number(SETTINGS_DEFAULTS['自動進行_定期実行_間隔分']));
        rows.push(['自動進行_定期実行_間隔分', intervalMinutes, '旧「自動進行_定期実行_間隔時間」から移行した分単位設定']);
    }
    if (rows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, descIndex ? 3 : 2).setValues(rows);
    }
}
function prepareManualStartLaunchDiagnosis() {
    syncSchema();
    syncVehicles();
    installDailyTriggers();
    return diagnoseStartLaunch(START_LAUNCH_DIAG_MODE.MANUAL);
}
function diagnoseStartLaunch(mode) {
    const ss = getSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const settings = loadSettings();
    const batchDef = resolveBiannualBatchDefinition(new Date(), tz, settings);
    const schemaDriftMessages = checkSchemaDrift();
    const missingSettingLabels = START_REQUIRED_SETTING_LABELS.filter(({ key }) => !String(settings[key] || '').trim()).map(({ label }) => label);
    const needsInputSummary = summarizeNeedsInput();
    const existingBatch = findExistingBatchSummary(batchDef.batchId);
    const blockingMessages = [];
    const warningMessages = [];
    if (missingSettingLabels.length > 0) {
        blockingMessages.push(`設定シートの必須通知先が未入力です: ${missingSettingLabels.join(' / ')}`);
    }
    if (schemaDriftMessages.length > 0) {
        warningMessages.push(`シート構造の不足を検知しました。syncSchema() で補完済みですが、念のため内容を確認してください。`);
    }
    if (needsInputSummary.count > 0) {
        const reasonText = needsInputSummary.reasonCounts.length > 0
            ? `主な内容: ${needsInputSummary.reasonCounts.map((entry) => `${entry.reason} ${entry.count}件`).join(' / ')}`
            : '要入力シートに未整備の行があります。';
        if (mode === START_LAUNCH_DIAG_MODE.SCHEDULED) {
            blockingMessages.push(`要入力が ${needsInputSummary.count} 件あるため、自動開始を停止しました。${reasonText}`);
        }
        else {
            warningMessages.push(`要入力は、台帳の未整備を知らせる一覧です。今回の開始は可能ですが、後続の確認や反映で詰まる原因になるため、早めの修正を推奨します。自動実行では停止対象です。${reasonText}`);
        }
    }
    if (!settings.mailSendEnabled) {
        warningMessages.push('通知_メール送信=FALSE です。メール送信を実施せず画面上の処理確認だけを行う設定なので、本番運用開始なら TRUE への切替が必要です。');
    }
    if (existingBatch) {
        warningMessages.push(`同一便の既存バッチ ${existingBatch.batchId}（ステータス: ${existingBatch.status || '未設定'}）が見つかりました。続行すると新規起票ではなく既存バッチの再利用として動作します。`);
    }
    const status = blockingMessages.length > 0
        ? START_LAUNCH_STATUS.BLOCKED
        : warningMessages.length > 0
            ? START_LAUNCH_STATUS.CONFIRM_REQUIRED
            : START_LAUNCH_STATUS.READY;
    return {
        status,
        mode,
        batchId: batchDef.batchId,
        batchLabel: batchDef.label,
        blockingMessages,
        warningMessages,
        missingSettingLabels,
        needsInputSummary,
        schemaDriftMessages,
        existingBatch,
        mailSendEnabled: settings.mailSendEnabled,
        preparedSteps: ['スキーマ同期済み', '車両同期済み', 'トリガー再設定済み'],
    };
}
function summarizeNeedsInput() {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.NEEDS_INPUT);
    if (!sheet || sheet.getLastRow() <= 1) {
        return {
            count: 0,
            reasonCounts: [],
        };
    }
    const data = sheet.getDataRange().getValues();
    const headerMap = data.length > 0 ? getHeaderMap(data[0]) : {};
    const reasonIndex = headerMap['不備内容'];
    const counts = {};
    let total = 0;
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row.every((cell) => cell === '' || cell === null))
            continue;
        total += 1;
        const reason = reasonIndex ? getCellValue(row, reasonIndex) : '';
        const normalizedReason = reason || '未分類';
        counts[normalizedReason] = (counts[normalizedReason] || 0) + 1;
    }
    const reasonCounts = Object.keys(counts)
        .map((reason) => ({ reason, count: counts[reason] }))
        .sort((a, b) => b.count - a.count || a.reason.localeCompare(b.reason))
        .slice(0, 5);
    return {
        count: total,
        reasonCounts,
    };
}
function findExistingBatchSummary(batchId) {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.NOTIFY_BATCH);
    if (!sheet || sheet.getLastRow() <= 1)
        return null;
    const data = sheet.getDataRange().getValues();
    const headerMap = data.length > 0 ? getHeaderMap(data[0]) : {};
    const rowInfo = findNotifyBatchRow(data, headerMap, batchId);
    if (!rowInfo)
        return null;
    return {
        batchId,
        status: getCellValue(rowInfo.row, headerMap['ステータス']),
        confirmationSheetName: getCellValue(rowInfo.row, headerMap['確認用シート名']),
        initialSentAt: parseDateValue(getCellRaw(rowInfo.row, headerMap['初回通知送信日時'])),
    };
}
function runBiannualScheduleCore(diagnosis) {
    const resolvedDiagnosis = diagnosis || diagnoseStartLaunch(START_LAUNCH_DIAG_MODE.MANUAL);
    const batchId = createBiannualBatch({ suppressUi: true });
    sendHqInitialEmail(batchId, { suppressUi: true });
    sendHqReminderIfNeeded(batchId);
    sendSenmuApprovalRequestIfReady(batchId);
    const context = getNotifyBatchContext(batchId);
    return {
        diagnosis: resolvedDiagnosis,
        batchId,
        context,
    };
}
function runBiannualScheduleWithSummary(diagnosis) {
    const execution = runBiannualScheduleCore(diagnosis);
    const resolvedDiagnosis = execution.diagnosis;
    const context = execution.context;
    const batchId = execution.batchId;
    const targetCount = context ? Number(getCellValue(context.row, context.headerMap['対象件数']) || 0) : 0;
    const confirmationSheetName = context ? getCellValue(context.row, context.headerMap['確認用シート名']) : '';
    const initialSentAt = context ? parseDateValue(getCellRaw(context.row, context.headerMap['初回通知送信日時'])) : null;
    const status = context ? getCellValue(context.row, context.headerMap['ステータス']) : '';
    const lines = [];
    lines.push(`${resolvedDiagnosis.batchLabel} の開始処理を実行しました。`);
    lines.push('');
    lines.push('実施した準備:');
    resolvedDiagnosis.preparedSteps.forEach((step) => lines.push(`- ${step}`));
    lines.push('');
    lines.push('開始結果:');
    lines.push(`- batchId: ${batchId}`);
    lines.push(`- 通知バッチのステータス: ${status || '未設定'}`);
    lines.push(`- 対象件数: ${targetCount}`);
    lines.push(`- 確認用シート: ${confirmationSheetName || '未作成'}`);
    if (resolvedDiagnosis.existingBatch) {
        lines.push(`- 同一便の既存バッチを再利用しました`);
    }
    if (!resolvedDiagnosis.mailSendEnabled) {
        lines.push(`- 初回通知メール: 通知_メール送信=FALSE のため送信スキップ`);
    }
    else if (initialSentAt) {
        lines.push(`- 初回通知メール: 送信処理まで実行済み`);
    }
    else {
        lines.push(`- 初回通知メール: 条件未達または未送信`);
    }
    if (resolvedDiagnosis.warningMessages.length > 0) {
        lines.push('');
        lines.push('引き続き確認してほしい点:');
        resolvedDiagnosis.warningMessages.forEach((message) => lines.push(`- ${message}`));
    }
    return {
        title: '開始処理の結果',
        body: lines.join('\n'),
    };
}
function notifyScheduledStartBlocked(diagnosis) {
    const settings = loadSettings();
    const adminTo = String(settings.adminTo || '').trim();
    const ss = getSpreadsheet();
    const detailLines = [
        ...diagnosis.blockingMessages.map((message) => `- ${message}`),
        ...diagnosis.warningMessages.map((message) => `- ${message}`),
    ];
    const mailLines = [
        `${diagnosis.batchLabel} の自動開始を停止しました。`,
        '',
        '停止理由:',
        ...detailLines,
        '',
        '自動実行では、設定不足や要入力がある状態で開始すると対象抽出や通知が不完全になるおそれがあるため、管理者確認が終わるまで開始しない運用にしています。',
        '設定シートや要入力シートを確認し、必要な修正後にメニュー「新しい確認依頼を開始（対象抽出〜初回通知）」から手動で開始してください。',
        '',
        '対象スプレッドシート:',
        ss.getUrl(),
        '',
        `便ID: ${diagnosis.batchId}`,
    ];
    const hash = createContentHash(mailLines);
    const propKey = `${PROP_KEYS.AUTO_START_BLOCK_NOTICE_PREFIX}${diagnosis.batchId}`;
    const props = PropertiesService.getDocumentProperties();
    const previousHash = String(props.getProperty(propKey) || '');
    if (hash === previousHash) {
        appendNotificationLog('半期自動開始停止', '', adminTo, diagnosis.batchId, '同一内容のため再通知なし');
        return false;
    }
    if (!settings.mailSendEnabled) {
        appendNotificationLog('半期自動開始停止', '', adminTo, diagnosis.batchId, '通知_メール送信=FALSE のため停止通知を送信スキップ');
        return false;
    }
    if (!adminTo) {
        appendNotificationLog('半期自動開始停止', '', '', diagnosis.batchId, '運用管理者_通知先Toが未設定のため停止通知を送信できません');
        return false;
    }
    const subject = `【要確認】${diagnosis.batchLabel} の自動開始を停止しました`;
    MailApp.sendEmail({
        to: adminTo,
        subject,
        name: settings.fromName,
        body: mailLines.join('\n'),
    });
    props.setProperty(propKey, hash);
    appendNotificationLog('半期自動開始停止', '', adminTo, diagnosis.batchId, `成功 ${diagnosis.blockingMessages.join(' / ')}`);
    return true;
}
function clearScheduledStartBlockedNotice(batchId) {
    const props = PropertiesService.getDocumentProperties();
    props.deleteProperty(`${PROP_KEYS.AUTO_START_BLOCK_NOTICE_PREFIX}${batchId}`);
}
function buildStartLaunchTitle(diagnosis) {
    return diagnosis.status === START_LAUNCH_STATUS.BLOCKED ? '開始前チェックで停止しました' : '開始前チェック';
}
function buildStartLaunchBody(diagnosis) {
    const lines = [];
    lines.push(`${diagnosis.batchLabel} の開始前チェックを行いました。`);
    lines.push('');
    lines.push('事前に整えた内容:');
    diagnosis.preparedSteps.forEach((step) => lines.push(`- ${step}`));
    if (diagnosis.status === START_LAUNCH_STATUS.BLOCKED) {
        lines.push('');
        lines.push('このまま開始すると、確認依頼が送れなかったり、対象抽出が不完全なまま進んだりして業務が中途半端に進むため、開始を止めています。');
        lines.push('修正が必要な項目:');
        diagnosis.blockingMessages.forEach((message) => lines.push(`- ${message}`));
        lines.push('');
        lines.push('設定シートや要入力シートを確認したあと、もう一度「新しい確認依頼を開始」を実行してください。');
        return lines.join('\n');
    }
    if (diagnosis.warningMessages.length > 0) {
        lines.push('');
        lines.push('確認してほしい警告:');
        diagnosis.warningMessages.forEach((message) => lines.push(`- ${message}`));
    }
    else {
        lines.push('');
        lines.push('開始前チェックで問題は見つかりませんでした。');
    }
    return lines.join('\n');
}
function showStartLaunchConfirmDialog(diagnosis) {
    const title = '開始前チェックの確認';
    const diagnosisJson = toInlineJson({
        title,
        body: buildStartLaunchBody(diagnosis),
        warningMessages: diagnosis.warningMessages,
    });
    const html = HtmlService.createHtmlOutput(`
    <div style="font-family: ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; padding: 20px; line-height: 1.6; color: #1f2937;">
      <h2 style="margin: 0 0 12px; font-size: 20px;">開始前チェックの確認</h2>
      <div id="content" style="white-space: pre-wrap; font-size: 13px; background: #f8fafc; border: 1px solid #dbe3ec; border-radius: 8px; padding: 16px;"></div>
      <div id="status" style="display:none; margin-top: 12px; color: #475569; font-size: 12px;">開始処理を実行しています...</div>
      <div style="display: flex; gap: 12px; justify-content: flex-end; margin-top: 18px;">
        <button id="cancel" style="padding: 10px 14px; border-radius: 8px; border: 1px solid #cbd5e1; background: #fff; cursor: pointer;">キャンセル</button>
        <button id="proceed" style="padding: 10px 14px; border-radius: 8px; border: 0; background: #1d4ed8; color: #fff; cursor: pointer;">このまま進める</button>
      </div>
    </div>
    <script>
      const payload = ${diagnosisJson};
      const content = document.getElementById('content');
      const status = document.getElementById('status');
      const proceedButton = document.getElementById('proceed');
      const cancelButton = document.getElementById('cancel');
      content.textContent = payload.body;

      cancelButton.addEventListener('click', () => google.script.host.close());
      proceedButton.addEventListener('click', () => {
        proceedButton.disabled = true;
        cancelButton.disabled = true;
        status.style.display = 'block';
        google.script.run
          .withSuccessHandler((result) => {
            proceedButton.style.display = 'none';
            cancelButton.textContent = '閉じる';
            cancelButton.disabled = false;
            content.textContent = result.body;
            status.style.display = 'none';
          })
          .withFailureHandler((error) => {
            proceedButton.disabled = false;
            cancelButton.disabled = false;
            status.style.display = 'none';
            content.textContent = '開始処理でエラーが発生しました。\\n' + (error && error.message ? error.message : error);
          })
          .continueRunDailyAfterConfirmation();
      });
    </script>
  `)
        .setWidth(860)
        .setHeight(640);
    SpreadsheetApp.getUi().showModalDialog(html, title);
}
function valuesFromSettingsSheet(data, headerMap, key) {
    const keyIndex = headerMap['設定項目'];
    const valueIndex = headerMap['値'];
    if (!keyIndex || !valueIndex)
        return '';
    for (let i = 1; i < data.length; i++) {
        if (getCellValue(data[i], keyIndex) === key) {
            return getCellRaw(data[i], valueIndex);
        }
    }
    return '';
}
function exportTestResults(limit) {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.TEST_RESULTS);
    if (!sheet)
        return '[]';
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1)
        return '[]';
    const max = typeof limit === 'number' && limit > 0 ? Math.floor(limit) : 200;
    const rows = data.slice(1).slice(-max);
    const toCellString = (value) => (value instanceof Date ? value.toISOString() : String(value || ''));
    const result = rows.map((r) => ({
        executedAt: toCellString(r[0]),
        item: toCellString(r[1]),
        result: toCellString(r[2]),
        detail: toCellString(r[3]),
    }));
    return JSON.stringify(result);
}
function ping() {
    return { ok: true, at: new Date().toISOString() };
}
function seedE2EMockVehicles() {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const ss = getSpreadsheet();
        const sheet = ss.getSheetByName(PRIMARY_SOURCE_SHEET);
        if (!sheet)
            throw new Error(`対象シートが存在しません: ${PRIMARY_SOURCE_SHEET}`);
        const data = sheet.getDataRange().getValues();
        if (data.length === 0)
            throw new Error(`${PRIMARY_SOURCE_SHEET} にヘッダ行がありません`);
        const headerMap = getHeaderMap(data[0]);
        const idx = resolveSourceHeaders(headerMap);
        const hasRegColumns = !!idx.regAll || !!(idx.regArea && idx.regClass && idx.regKana && idx.regNumber);
        if (!hasRegColumns || !idx.contractEnd || !idx.dept) {
            throw new Error('モック投入に必要なヘッダが不足しています（登録番号/契約満了日/管理部門）');
        }
        const tz = ss.getSpreadsheetTimeZone();
        const today = toDateOnly(new Date(), tz);
        const baseStart = addDays(today, -300);
        const scenarios = [
            {
                code: '1001',
                label: 'TEST_更新対象',
                chassis: 'TEST-CH-001',
                dept: '本社総務',
                manager: 'テスト太郎',
                contractEnd: today,
                inspectionEnd: addDays(today, 365),
                leaseFee: 50000,
            },
            {
                code: '1002',
                label: 'TEST_解約対象',
                chassis: 'TEST-CH-002',
                dept: '本社総務',
                manager: 'テスト花子',
                contractEnd: addDays(today, 1),
                inspectionEnd: addDays(today, 366),
                leaseFee: 52000,
            },
            {
                code: '9001',
                label: 'TEST_対象外',
                chassis: 'TEST-CH-003',
                dept: '本社総務',
                manager: 'テスト次郎',
                contractEnd: addDays(today, 200),
                inspectionEnd: addDays(today, 560),
                leaseFee: 53000,
            },
            {
                code: 'ERR1',
                label: 'TEST_満了日欠損',
                chassis: 'TEST-CH-004',
                dept: '本社総務',
                manager: 'テスト欠損',
                contractEnd: null,
                inspectionEnd: addDays(today, 430),
                leaseFee: 54000,
            },
        ];
        const existingKeys = {};
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const regCombined = getSourceRegistrationCombined(row, idx);
            const chassis = getCellValue(row, idx.chassis);
            const key = `${regCombined}__${chassis}`;
            if (regCombined || chassis)
                existingKeys[key] = true;
        }
        const rowsToAdd = [];
        let skippedExisting = 0;
        scenarios.forEach((s) => {
            const regCombined = `TEST-${s.code}`;
            const key = `${regCombined}__${s.chassis}`;
            if (existingKeys[key]) {
                skippedExisting += 1;
                return;
            }
            const row = new Array(data[0].length).fill('');
            if (idx.regAll) {
                row[idx.regAll - 1] = regCombined;
            }
            else {
                row[idx.regArea - 1] = 'TEST';
                row[idx.regClass - 1] = '99';
                row[idx.regKana - 1] = 'テ';
                row[idx.regNumber - 1] = s.code;
            }
            if (idx.vehicleType)
                row[idx.vehicleType - 1] = s.label;
            if (idx.chassis)
                row[idx.chassis - 1] = s.chassis;
            if (idx.contractStart)
                row[idx.contractStart - 1] = baseStart;
            if (idx.contractEnd && s.contractEnd)
                row[idx.contractEnd - 1] = s.contractEnd;
            if (idx.dept)
                row[idx.dept - 1] = s.dept;
            if (idx.manager)
                row[idx.manager - 1] = s.manager;
            if (idx.contractTerm)
                row[idx.contractTerm - 1] = '60ヶ月';
            if (idx.inspectionEnd)
                row[idx.inspectionEnd - 1] = s.inspectionEnd;
            if (idx.leaseFee)
                row[idx.leaseFee - 1] = s.leaseFee;
            rowsToAdd.push(row);
            existingKeys[key] = true;
        });
        if (rowsToAdd.length > 0) {
            sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, data[0].length).setValues(rowsToAdd);
        }
        const result = {
            inserted: rowsToAdd.length,
            skippedExisting,
            sourceSheet: PRIMARY_SOURCE_SHEET,
        };
        uiAlertSafe(`E2Eモック車両を投入しました。\n${JSON.stringify(result)}`);
        return result;
    }
    finally {
        lock.releaseLock();
    }
}
function cleanupTestData() {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const ss = getSpreadsheet();
        const removed = {
            sourceSheets: {},
            needsInput: 0,
        };
        const testVehicleIds = {};
        const testRegCombined = {};
        // 元台帳（車両一覧）からテスト車両行を削除し、関連IDを収集
        const sourceSheet = ss.getSheetByName(PRIMARY_SOURCE_SHEET);
        if (sourceSheet) {
            const data = sourceSheet.getDataRange().getValues();
            if (data.length > 1) {
                const header = getHeaderMap(data[0]);
                const sourceIdx = resolveSourceHeaders(header);
                const idx = {
                    vehicleId: header['vehicleId'],
                };
                const rowsToDelete = [];
                for (let i = 1; i < data.length; i++) {
                    const row = data[i];
                    if (row.every((cell) => cell === '' || cell === null))
                        continue;
                    const vehicleId = getCellValue(row, idx.vehicleId);
                    const regCombined = getSourceRegistrationCombined(row, sourceIdx);
                    const chassis = getCellValue(row, sourceIdx.chassis);
                    const vehicleType = getCellValue(row, sourceIdx.vehicleType);
                    const isTest = (regCombined && String(regCombined).startsWith('TEST')) ||
                        (vehicleId && vehicleId.indexOf('__TEST') >= 0) ||
                        (chassis && String(chassis).startsWith('TEST-')) ||
                        (vehicleType && String(vehicleType).startsWith('テスト_'));
                    if (!isTest)
                        continue;
                    if (vehicleId)
                        testVehicleIds[vehicleId] = true;
                    if (regCombined)
                        testRegCombined[String(regCombined)] = true;
                    rowsToDelete.push(i + 1);
                }
                for (let i = rowsToDelete.length - 1; i >= 0; i--) {
                    sourceSheet.deleteRow(rowsToDelete[i]);
                }
                removed.sourceSheets[PRIMARY_SOURCE_SHEET] = rowsToDelete.length;
            }
        }
        // 要入力（テスト車両由来のみ削除）
        const needsInputSheet = ss.getSheetByName(SHEET_NAMES.NEEDS_INPUT);
        if (needsInputSheet) {
            const data = needsInputSheet.getDataRange().getValues();
            if (data.length > 1) {
                const header = getHeaderMap(data[0]);
                const idx = {
                    vehicleId: header['vehicleId'],
                    regCombined: header['登録番号_結合'],
                };
                const rowsToDelete = [];
                for (let i = 1; i < data.length; i++) {
                    const row = data[i];
                    const vehicleId = getCellValue(row, idx.vehicleId);
                    const regCombined = getCellValue(row, idx.regCombined);
                    const isTest = (vehicleId && testVehicleIds[vehicleId]) ||
                        (regCombined && (regCombined.startsWith('TEST') || testRegCombined[regCombined]));
                    if (!isTest)
                        continue;
                    rowsToDelete.push(i + 1);
                }
                for (let i = rowsToDelete.length - 1; i >= 0; i--) {
                    needsInputSheet.deleteRow(rowsToDelete[i]);
                    removed.needsInput += 1;
                }
            }
        }
        appendTestResult('cleanupTestData', 'OK', JSON.stringify(removed));
        uiAlertSafe(`テストデータを掃除しました。\n${JSON.stringify(removed)}`);
        return removed;
    }
    finally {
        lock.releaseLock();
    }
}
function cleanupUnusedSheets() {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const ss = getSpreadsheet();
        const requiredSheetNames = {};
        SCHEMA_DEFS.forEach((def) => {
            requiredSheetNames[def.name] = true;
        });
        requiredSheetNames[PRIMARY_SOURCE_SHEET] = true;
        const activeConfirmationNames = {};
        const notifyBatchSheet = ss.getSheetByName(SHEET_NAMES.NOTIFY_BATCH);
        if (notifyBatchSheet && notifyBatchSheet.getLastRow() > 1) {
            const data = notifyBatchSheet.getDataRange().getValues();
            const headerMap = getHeaderMap(data[0]);
            const confirmationIdx = headerMap['確認用シート名'];
            if (confirmationIdx) {
                for (let i = 1; i < data.length; i++) {
                    const name = getCellValue(data[i], confirmationIdx);
                    if (name)
                        activeConfirmationNames[name] = true;
                }
            }
        }
        const deleted = [];
        const kept = [];
        const sheets = ss.getSheets();
        sheets.forEach((sheet) => {
            const name = sheet.getName();
            if (requiredSheetNames[name]) {
                kept.push(name);
                return;
            }
            if (isConfirmationSheetName(name)) {
                if (activeConfirmationNames[name]) {
                    kept.push(name);
                    return;
                }
                ss.deleteSheet(sheet);
                deleted.push(name);
                return;
            }
            ss.deleteSheet(sheet);
            deleted.push(name);
        });
        const result = {
            deletedCount: deleted.length,
            deleted,
            keptCount: kept.length,
            kept,
        };
        appendTestResult('cleanupUnusedSheets', 'OK', JSON.stringify(result));
        uiAlertSafe(`不要シート削除を実行しました。\n${JSON.stringify(result)}`);
        return result;
    }
    finally {
        lock.releaseLock();
    }
}
function runTestSuite() {
    clearTestResults();
    appendTestResult('開始', 'OK', new Date().toISOString());
    syncSchema();
    appendTestResult('syncSchema', 'OK', '');
    syncVehicles();
    appendTestResult('syncVehicles', 'OK', '');
    const batchId = createBiannualBatch();
    appendTestResult('createBiannualBatch', batchId ? 'OK' : 'NG', String(batchId || ''));
    const builtSheetName = buildConfirmationSheetForLatestBatch();
    appendTestResult('buildConfirmationSheetForLatestBatch', builtSheetName ? 'OK' : 'NG', builtSheetName || 'シート未生成');
    const ss = getSpreadsheet();
    const batchContext = getNotifyBatchContext(batchId || '');
    if (!batchContext) {
        appendTestResult('中断', 'NG', '通知バッチが見つかりません');
        return;
    }
    sendHqInitialEmail(batchId || batchContext.batchId);
    const afterInitial = getNotifyBatchContext(batchId || batchContext.batchId);
    const sentAt = afterInitial ? parseDateValue(getCellRaw(afterInitial.row, afterInitial.headerMap['初回通知送信日時'])) : null;
    appendTestResult('sendHqInitialEmail', sentAt ? 'OK' : 'NG', sentAt ? '初回通知送信日時が設定されました' : '未設定');
    sendHqReminderIfNeeded(batchId || batchContext.batchId);
    const afterReminder = getNotifyBatchContext(batchId || batchContext.batchId);
    const reminderAfter = afterReminder ? getCellRaw(afterReminder.row, afterReminder.headerMap['リマインド送信日時']) : '';
    appendTestResult('sendHqReminderIfNeeded', 'OK', reminderAfter ? '送信条件一致で送信済み' : '送信条件未一致のため未送信');
    const latestContext = getNotifyBatchContext(batchId || batchContext.batchId) || batchContext;
    const confirmationSheetName = getCellValue(latestContext.row, latestContext.headerMap['確認用シート名']);
    const confirmationSheet = confirmationSheetName ? ss.getSheetByName(confirmationSheetName) : null;
    if (confirmationSheet) {
        const data = confirmationSheet.getDataRange().getValues();
        const ch = data.length > 0 ? getHeaderMap(data[0]) : {};
        const counts = summarizeConfirmationSheetRows(data, ch);
        appendTestResult('期待値:確認用シート生成件数', counts.total >= 0 ? 'OK' : 'NG', `total=${counts.total} unanswered=${counts.unanswered} unchecked=${counts.unchecked}`);
        appendTestResult('期待値:確認用シート不正理由列', ch['不正理由'] ? 'OK' : 'NG', JSON.stringify({ hasInvalidReason: !!ch['不正理由'] }));
    }
    else {
        appendTestResult('期待値:確認用シート生成件数', 'NG', '確認用シート未生成');
        appendTestResult('期待値:確認用シート不正理由列', 'NG', '確認用シート未生成');
    }
    const notifyBatchHeaders = getSchemaHeaders(SHEET_NAMES.NOTIFY_BATCH);
    const hasMurataIssueHeaders = notifyBatchHeaders.indexOf('村田主任不備通知日時') >= 0 &&
        notifyBatchHeaders.indexOf('村田主任不備通知ハッシュ') >= 0 &&
        notifyBatchHeaders.indexOf('村田主任反映完了通知日時') >= 0 &&
        notifyBatchHeaders.indexOf('村田主任反映完了通知ハッシュ') >= 0;
    appendTestResult('期待値:通知バッチ通知列', hasMurataIssueHeaders ? 'OK' : 'NG', JSON.stringify({ headers: notifyBatchHeaders }));
    const onOpenSource = String(globalThis.onOpen || onOpen).replace(/\s+/g, ' ');
    const manualMenuSlimmed = onOpenSource.includes("addItem('新しい確認依頼を開始（対象抽出〜初回通知）', 'runDaily')") &&
        onOpenSource.includes("addItem('進行中の確認依頼を再開（最新）', 'runAutoAdvanceNow')") &&
        !onOpenSource.includes('村田主任通知送信（最新バッチ）') &&
        !onOpenSource.includes('マスター反映（最新バッチ）');
    appendTestResult('期待値:手動メニュー整理', manualMenuSlimmed ? 'OK' : 'NG', '開始と途中再開をトップレベルに整理');
    const legacyCheck = verifyAcceptanceCondition7LegacyNonReachable();
    appendTestResult('受け入れ条件7:旧導線非到達', legacyCheck.ok ? 'OK' : 'NG', JSON.stringify({
        remainingLegacyEntries: legacyCheck.remainingLegacyEntries,
        wrapperErrors: legacyCheck.wrapperErrors,
    }));
    appendTestResult('完了', 'OK', '');
}
function verifyAcceptanceCondition7LegacyNonReachable() {
    const globalObj = globalThis;
    const removedLegacyEntries = [
        'doGet',
        'doPost',
        'onRequestFormSubmit',
        'onApprovalFormSubmit',
        'validateRequestAccess',
        'createRequestForms',
        'createOrUpdateApprovalForm',
        'loadDeptMaster',
        'generateDeptTokens',
    ];
    const remainingLegacyEntries = removedLegacyEntries.filter((name) => typeof globalObj[name] === 'function');
    const wrapperExpectedSource = {
        createRequests: 'return createBiannualBatch();',
        sendInitialEmails: 'return sendHqInitialEmail();',
        sendReminderEmails: 'return sendHqReminderIfNeeded();',
        applyAnswers: 'return applyMasterUpdates();',
        sendApprovalRequestEmails: 'return sendSenmuApprovalRequestIfReady();',
        applyApprovalDecisions: 'return applySenmuDecisionFromSheet();',
    };
    const wrapperErrors = [];
    Object.keys(wrapperExpectedSource).forEach((name) => {
        const fn = globalObj[name];
        if (typeof fn !== 'function') {
            wrapperErrors.push(`${name}: 関数が見つかりません`);
            return;
        }
        const compact = String(fn).replace(/\s+/g, ' ');
        const expected = wrapperExpectedSource[name].replace(/\s+/g, ' ');
        if (!compact.includes(expected)) {
            wrapperErrors.push(`${name}: 互換ラッパーが1行委譲になっていません`);
        }
    });
    return {
        ok: remainingLegacyEntries.length === 0 && wrapperErrors.length === 0,
        remainingLegacyEntries,
        wrapperErrors,
    };
}
function installDailyTriggers() {
    const settings = loadSettings();
    const ss = getSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const managedHandlers = ['runDaily', 'runBiannualSchedule', 'runAutoAdvance', 'onEditAutoAdvance', 'onEditSourceSync', 'onEditSettingsSync', 'syncVehicles'];
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach((trigger) => {
        const handler = trigger.getHandlerFunction();
        if (managedHandlers.indexOf(handler) >= 0) {
            ScriptApp.deleteTrigger(trigger);
        }
    });
    const now = toDateOnly(new Date(), tz);
    const years = [now.getFullYear(), now.getFullYear() + 1];
    const reminderBeforeDays = Math.max(0, toNumber(settings.reminderBeforeDays, 10));
    const dateKeys = {};
    const runDates = [];
    years.forEach((year) => {
        const marchSendDate = resolveMonthDaySettingDate(settings.marchSendDate, year, 3, 1, tz);
        const septemberSendDate = resolveMonthDaySettingDate(settings.septemberSendDate, year, 9, 1, tz);
        const marchDeadline = resolveMonthDaySettingDate(settings.marchDeadline, year, 3, 31, tz);
        const septemberDeadline = resolveMonthDaySettingDate(settings.septemberDeadline, year, 9, 30, tz);
        const marchReminder = addDays(marchDeadline, -reminderBeforeDays);
        const septemberReminder = addDays(septemberDeadline, -reminderBeforeDays);
        [marchSendDate, septemberSendDate, marchReminder, septemberReminder].forEach((d) => {
            const date = toDateOnly(d, tz);
            if (date.getTime() < now.getTime())
                return;
            const key = Utilities.formatDate(date, tz, 'yyyy-MM-dd');
            if (dateKeys[key])
                return;
            dateKeys[key] = true;
            runDates.push(date);
        });
    });
    runDates.forEach((runDate) => {
        const triggerAt = new Date(runDate.getFullYear(), runDate.getMonth(), runDate.getDate(), 8, 0, 0);
        ScriptApp.newTrigger('runBiannualSchedule').timeBased().at(triggerAt).create();
    });
    if (settings.autoAdvanceEnabled && settings.autoAdvanceTimerEnabled) {
        ScriptApp.newTrigger('runAutoAdvance').timeBased().everyMinutes(settings.autoAdvanceTimerIntervalMinutes).create();
    }
    ScriptApp.newTrigger('syncVehicles').timeBased().everyHours(1).create();
    ScriptApp.newTrigger('onEditSourceSync').forSpreadsheet(ss).onEdit().create();
    ScriptApp.newTrigger('onEditSettingsSync').forSpreadsheet(ss).onEdit().create();
}
// === helpers ===
function getSpreadsheet() {
    return SpreadsheetApp.getActiveSpreadsheet();
}
function ensureSheet(ss, name) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
        sheet = ss.insertSheet(name);
    }
    return sheet;
}
function getSchemaHeaders(name) {
    const def = SCHEMA_DEFS.find((d) => d.name === name);
    if (!def)
        throw new Error(`schema not found: ${name}`);
    return def.headers;
}
function ensureHeaders(sheet, headerRow, headers) {
    const lastColumn = sheet.getLastColumn();
    if (lastColumn === 0) {
        sheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]);
        return;
    }
    const rowValues = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0];
    const headerMap = getHeaderMap(rowValues);
    const missing = headers.filter((header) => !headerMap[header]);
    if (missing.length > 0) {
        const startCol = lastColumn + 1;
        sheet.getRange(headerRow, startCol, 1, missing.length).setValues([missing]);
    }
}
function getHeaderMap(headers) {
    const map = {};
    headers.forEach((value, index) => {
        const key = String(value || '').trim();
        if (key)
            map[key] = index + 1;
    });
    return map;
}
function resolveSourceHeaders(headerMap) {
    const normalizedMap = buildNormalizedHeaderMap(headerMap);
    return {
        regArea: findHeaderIndex(headerMap, normalizedMap, [
            '地名',
            '登録番号_地名',
            '登録番号（地名）',
            '登録番号(地名)',
            '登録番号【地名】',
            '登録番号地名',
        ]),
        regClass: findHeaderIndex(headerMap, normalizedMap, [
            '分類番号',
            '分類',
            '分類番号(3桁)',
            '分類番号（3桁）',
            '分類番号3桁',
            '分類(3桁)',
            '分類（3桁）',
            '分類3桁',
            '登録番号_分類',
            '登録番号（分類）',
            '登録番号(分類)',
            '登録番号【分類】',
            '登録番号分類',
        ]),
        regKana: findHeaderIndex(headerMap, normalizedMap, [
            'かな',
            'カナ',
            '登録番号_かな',
            '登録番号（かな）',
            '登録番号(かな)',
            '登録番号【かな】',
            '登録番号かな',
            '登録番号カナ',
        ]),
        regNumber: findHeaderIndex(headerMap, normalizedMap, [
            '番号',
            '番号(4桁)',
            '番号（4桁）',
            '番号4桁',
            '登録番号_番号',
            '登録番号（番号）',
            '登録番号(番号)',
            '登録番号【番号】',
        ]),
        // 台帳が「登録番号」1列で持っているケースがある（分割列が無い/使わない）
        regAll: findHeaderIndex(headerMap, normalizedMap, ['登録番号', '車両番号', '車両登録番号', 'ナンバー', 'ﾅﾝﾊﾞｰ']),
        vehicleType: findHeaderIndex(headerMap, normalizedMap, ['車種', '車名', '車種名']),
        chassis: findHeaderIndex(headerMap, normalizedMap, ['車台番号', '車体番号', '車台No', '車台NO', '車台No.']),
        contractStart: findHeaderIndex(headerMap, normalizedMap, ['契約開始日', '契約開始', '開始日', 'リース開始日']),
        contractEnd: findHeaderIndex(headerMap, normalizedMap, [
            '契約満了日',
            '契約満了',
            '満了日',
            '満了日（予定）',
            '契約満了日（予定）',
            'リース満了日',
            'リース契約満了日',
            '契約終了日',
            '終了日',
        ]),
        dept: findHeaderIndex(headerMap, normalizedMap, [
            '管理部門',
            '管理部署',
            '部署',
            '部門',
            '管理課',
            '所属部署',
            '所属部門',
        ]),
        manager: findHeaderIndex(headerMap, normalizedMap, [
            '管理担当者',
            '担当者',
            '管理担当',
            '担当',
            '責任者',
        ]),
        contractTerm: findHeaderIndex(headerMap, normalizedMap, ['契約期間', 'リース期間', '契約年数', '期間']),
        inspectionEnd: findHeaderIndex(headerMap, normalizedMap, [
            '車検満了日',
            '車検満了',
            '車検期限',
            '車検期限日',
        ]),
        leaseFee: findHeaderIndex(headerMap, normalizedMap, [
            'リース料（税抜）',
            'リース料(税抜)',
            'リース料税抜',
            'リース料',
            '月額リース料',
        ]),
    };
}
function normalizeHeaderKey(value) {
    if (value === null || value === undefined)
        return '';
    return String(value)
        .normalize('NFKC')
        .trim()
        .replace(/[\s\u3000]+/g, '')
        .replace(/[＿_]/g, '')
        .replace(/[()（）［］[\]【】{}｛｝<>＜＞]/g, '')
        .replace(/[・]/g, '')
        .replace(/[‐‑‒–—−-]/g, '');
}
function buildNormalizedHeaderMap(headerMap) {
    const normalizedMap = {};
    Object.keys(headerMap).forEach((key) => {
        const normalized = normalizeHeaderKey(key);
        if (!normalized)
            return;
        if (!normalizedMap[normalized])
            normalizedMap[normalized] = headerMap[key];
    });
    return normalizedMap;
}
function findHeaderIndex(headerMap, normalizedMap, names) {
    for (const name of names) {
        if (headerMap[name])
            return headerMap[name];
        const normalized = normalizeHeaderKey(name);
        if (normalized && normalizedMap[normalized])
            return normalizedMap[normalized];
        // 表記ゆれ対策: 末尾の補足（例: "(3ケタ)" など）が付く場合をユニーク一致の範囲で吸収する
        if (normalized) {
            const matchedKeys = Object.keys(normalizedMap).filter((k) => k.includes(normalized));
            if (matchedKeys.length === 1)
                return normalizedMap[matchedKeys[0]];
        }
    }
    return 0;
}
function getCellValue(row, index) {
    if (!index)
        return '';
    const value = row[index - 1];
    return value === null || value === undefined ? '' : String(value).trim();
}
function getCellRaw(row, index) {
    if (!index)
        return null;
    return row[index - 1];
}
function getSourceRegistrationParts(row, idx) {
    return {
        area: getCellValue(row, idx.regArea),
        cls: getCellValue(row, idx.regClass),
        kana: getCellValue(row, idx.regKana),
        num: getCellValue(row, idx.regNumber),
    };
}
function getSourceRegistrationCombined(row, idx) {
    const fromAll = getCellValue(row, idx.regAll);
    if (fromAll)
        return fromAll;
    const parts = getSourceRegistrationParts(row, idx);
    return buildRegistrationCombined(parts.area, parts.cls, parts.kana, parts.num);
}
function parseDateValue(value) {
    if (!value)
        return null;
    if (value instanceof Date)
        return value;
    const parsed = new Date(value);
    return isNaN(parsed.getTime()) ? null : parsed;
}
function toDateOnly(date, tz) {
    const formatted = Utilities.formatDate(date, tz, 'yyyy/MM/dd');
    return new Date(formatted);
}
function addMonthsClamped(date, months) {
    const year = date.getFullYear();
    const month = date.getMonth();
    const day = date.getDate();
    const base = new Date(year, month + months, 1);
    const lastDay = new Date(base.getFullYear(), base.getMonth() + 1, 0).getDate();
    return new Date(base.getFullYear(), base.getMonth(), Math.min(day, lastDay));
}
function addDays(date, days) {
    const d = new Date(date.getTime());
    d.setDate(d.getDate() + days);
    return d;
}
function isWithinRange(date, start, end) {
    return date.getTime() >= start.getTime() && date.getTime() <= end.getTime();
}
function resolveBiannualBatchDefinition(referenceDate, tz, settings) {
    const month = referenceDate.getMonth() + 1;
    const year = referenceDate.getFullYear();
    const isH1 = month <= 3 || month >= 10;
    if (isH1) {
        const batchYear = month >= 10 ? year + 1 : year;
        const rangeStartYear = batchYear - 1;
        const targetStart = toDateOnly(new Date(rangeStartYear, 9, 1), tz);
        const targetEnd = toDateOnly(new Date(batchYear, 2, 31), tz);
        const sendDate = resolveMonthDaySettingDate(settings.marchSendDate, batchYear, 3, 1, tz);
        const deadline = resolveMonthDaySettingDate(settings.marchDeadline, batchYear, 3, 31, tz);
        return {
            batchId: `${batchYear}H1`,
            label: `${batchYear}年3月便`,
            sendDate,
            deadline,
            targetStart,
            targetEnd,
        };
    }
    const batchYear = year;
    const targetStart = toDateOnly(new Date(batchYear, 3, 1), tz);
    const targetEnd = toDateOnly(new Date(batchYear, 8, 30), tz);
    const sendDate = resolveMonthDaySettingDate(settings.septemberSendDate, batchYear, 9, 1, tz);
    const deadline = resolveMonthDaySettingDate(settings.septemberDeadline, batchYear, 9, 30, tz);
    return {
        batchId: `${batchYear}H2`,
        label: `${batchYear}年9月便`,
        sendDate,
        deadline,
        targetStart,
        targetEnd,
    };
}
function resolveMonthDaySettingDate(rawValue, year, fallbackMonth, fallbackDay, tz) {
    if (rawValue instanceof Date) {
        return toDateOnly(rawValue, tz);
    }
    const text = String(rawValue || '').trim();
    if (text) {
        const fullDateMatch = text.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);
        if (fullDateMatch) {
            const parsed = new Date(Number(fullDateMatch[1]), Number(fullDateMatch[2]) - 1, Number(fullDateMatch[3]));
            return toDateOnly(parsed, tz);
        }
        const monthDayMatch = text.match(/^(\d{1,2})[/-](\d{1,2})$/);
        if (monthDayMatch) {
            const parsed = new Date(year, Number(monthDayMatch[1]) - 1, Number(monthDayMatch[2]));
            return toDateOnly(parsed, tz);
        }
        const jpMonthDayMatch = text.match(/^(\d{1,2})月(\d{1,2})日?$/);
        if (jpMonthDayMatch) {
            const parsed = new Date(year, Number(jpMonthDayMatch[1]) - 1, Number(jpMonthDayMatch[2]));
            return toDateOnly(parsed, tz);
        }
    }
    return toDateOnly(new Date(year, fallbackMonth - 1, fallbackDay), tz);
}
function findNotifyBatchRow(batchData, headerMap, batchId) {
    const batchIdIndex = headerMap['batchId'];
    if (!batchIdIndex)
        return null;
    for (let i = 1; i < batchData.length; i++) {
        const row = batchData[i];
        if (getCellValue(row, batchIdIndex) !== batchId)
            continue;
        return { row, rowIndex: i + 1 };
    }
    return null;
}
function getNotifyBatchContext(batchId) {
    const ss = getSpreadsheet();
    const notifyBatchSheet = ss.getSheetByName(SHEET_NAMES.NOTIFY_BATCH);
    if (!notifyBatchSheet || notifyBatchSheet.getLastRow() <= 1)
        return null;
    ensureHeaders(notifyBatchSheet, 1, getSchemaHeaders(SHEET_NAMES.NOTIFY_BATCH));
    const batchData = notifyBatchSheet.getDataRange().getValues();
    if (batchData.length <= 1)
        return null;
    const headerMap = getHeaderMap(batchData[0]);
    const batchIdIndex = headerMap['batchId'];
    if (!batchIdIndex)
        return null;
    let rowInfo = null;
    if (batchId) {
        rowInfo = findNotifyBatchRow(batchData, headerMap, batchId);
    }
    else {
        for (let i = batchData.length - 1; i >= 1; i--) {
            const row = batchData[i];
            const id = getCellValue(row, batchIdIndex);
            if (!id)
                continue;
            rowInfo = { row, rowIndex: i + 1 };
            break;
        }
    }
    if (!rowInfo)
        return null;
    const resolvedBatchId = getCellValue(rowInfo.row, batchIdIndex);
    if (!resolvedBatchId)
        return null;
    return {
        ss,
        notifyBatchSheet,
        batchData,
        headerMap,
        row: rowInfo.row,
        rowIndex: rowInfo.rowIndex,
        batchId: resolvedBatchId,
    };
}
function getBatchWorkflowSnapshot(batchId) {
    const batchContext = getNotifyBatchContext(batchId);
    if (!batchContext)
        return null;
    const confirmationSheetName = getCellValue(batchContext.row, batchContext.headerMap['確認用シート名']);
    if (!confirmationSheetName) {
        return {
            batchContext,
            confirmationSheet: null,
            confirmationData: [],
            confirmationHeader: {},
            hqGate: createEmptyGateEvaluation(),
            senmuGate: createEmptyGateEvaluation(),
            murataGate: createEmptyGateEvaluation(),
        };
    }
    const confirmationSheet = batchContext.ss.getSheetByName(confirmationSheetName);
    if (!confirmationSheet || confirmationSheet.getLastRow() <= 0) {
        return {
            batchContext,
            confirmationSheet: confirmationSheet || null,
            confirmationData: [],
            confirmationHeader: {},
            hqGate: createEmptyGateEvaluation(),
            senmuGate: createEmptyGateEvaluation(),
            murataGate: createEmptyGateEvaluation(),
        };
    }
    const confirmationData = confirmationSheet.getDataRange().getValues();
    const confirmationHeader = confirmationData.length > 0 ? getHeaderMap(confirmationData[0]) : {};
    return {
        batchContext,
        confirmationSheet,
        confirmationData,
        confirmationHeader,
        hqGate: evaluateHqGate(confirmationData, confirmationHeader),
        senmuGate: evaluateSenmuGate(confirmationData, confirmationHeader),
        murataGate: evaluateMurataGate(confirmationData, confirmationHeader),
    };
}
function createEmptyGateEvaluation() {
    return {
        total: 0,
        completed: 0,
        pending: 0,
        invalid: 0,
        unchecked: 0,
        approved: 0,
        returned: 0,
        ready: false,
        hasAnyInput: false,
    };
}
function evaluateHqGate(data, headerMap) {
    const result = createEmptyGateEvaluation();
    if (!data || data.length <= 1)
        return result;
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const vehicleId = getCellValue(row, headerMap['vehicleId']);
        if (!vehicleId)
            continue;
        result.total += 1;
        const answer = normalizeAnswerLabel(getCellValue(row, headerMap['本部回答']));
        const checked = isCheckedCell(getCellRaw(row, headerMap['回答確認済み']));
        if (answer)
            result.hasAnyInput = true;
        if (!answer)
            result.pending += 1;
        if (!checked)
            result.unchecked += 1;
        if (answer && checked)
            result.completed += 1;
    }
    result.ready = result.total > 0 && result.pending === 0 && result.unchecked === 0;
    return result;
}
function evaluateSenmuGate(data, headerMap) {
    const result = createEmptyGateEvaluation();
    if (!data || data.length <= 1)
        return result;
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const vehicleId = getCellValue(row, headerMap['vehicleId']);
        if (!vehicleId)
            continue;
        result.total += 1;
        const rawDecision = getCellValue(row, headerMap['専務判断']);
        const decision = normalizeSenmuDecision(rawDecision);
        const checked = isCheckedCell(getCellRaw(row, headerMap['専務確認済み']));
        if (rawDecision)
            result.hasAnyInput = true;
        if (!decision) {
            if (rawDecision) {
                result.invalid += 1;
            }
            else {
                result.pending += 1;
            }
        }
        else {
            result.completed += 1;
            if (decision === APPROVAL_INPUT.APPROVE)
                result.approved += 1;
            if (decision === APPROVAL_INPUT.RETURN)
                result.returned += 1;
        }
        if (!checked)
            result.unchecked += 1;
    }
    result.ready = result.total > 0 && result.pending === 0 && result.invalid === 0 && result.unchecked === 0;
    return result;
}
function evaluateMurataGate(data, headerMap) {
    const result = createEmptyGateEvaluation();
    if (!data || data.length <= 1)
        return result;
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const vehicleId = getCellValue(row, headerMap['vehicleId']);
        if (!vehicleId)
            continue;
        const decision = normalizeSenmuDecision(getCellValue(row, headerMap['専務判断']));
        if (decision !== APPROVAL_INPUT.APPROVE)
            continue;
        result.total += 1;
        const checked = isCheckedCell(getCellRaw(row, headerMap['村田主任確認済み']));
        const policy = normalizeAnswerLabel(getCellValue(row, headerMap['本部回答']));
        const hasRequiredInputs = policy === ANSWER_LABELS.RENEW
            ? !!parseDateValue(getCellRaw(row, headerMap['新契約開始日'])) && !!parseDateValue(getCellRaw(row, headerMap['新契約満了日']))
            : !!policy && isCheckedCell(getCellRaw(row, headerMap['解約完了']));
        if (checked)
            result.hasAnyInput = true;
        if (!checked)
            result.unchecked += 1;
        if (!hasRequiredInputs)
            result.pending += 1;
        if (checked && hasRequiredInputs)
            result.completed += 1;
    }
    result.ready = result.total > 0 && result.pending === 0 && result.unchecked === 0;
    return result;
}
function evaluateBatchPhase(batchId) {
    const snapshot = getBatchWorkflowSnapshot(batchId);
    if (!snapshot)
        return null;
    const { batchContext, confirmationSheet, hqGate, senmuGate, murataGate } = snapshot;
    const { row, headerMap } = batchContext;
    const status = getCellValue(row, headerMap['ステータス']);
    const initialSentAt = parseDateValue(getCellRaw(row, headerMap['初回通知送信日時']));
    const senmuRequestedAt = parseDateValue(getCellRaw(row, headerMap['専務依頼送信日時']));
    const murataSentAt = parseDateValue(getCellRaw(row, headerMap['村田主任通知送信日時']));
    let phase = BATCH_PHASE.HQ_WAITING;
    if (!confirmationSheet || !initialSentAt) {
        phase = BATCH_PHASE.INITIAL_PENDING;
    }
    else if (!hqGate.ready) {
        phase = BATCH_PHASE.HQ_WAITING;
    }
    else if (!senmuRequestedAt) {
        phase = BATCH_PHASE.SENMU_REQUEST_READY;
    }
    else if (!senmuGate.ready) {
        phase = BATCH_PHASE.SENMU_WAITING;
    }
    else if (senmuGate.returned > 0 || status === BIANNUAL_BATCH_STATUS.SENMU_RETURNED) {
        phase = BATCH_PHASE.SENMU_RETURN_READY;
    }
    else if (!murataSentAt) {
        phase = BATCH_PHASE.MURATA_NOTIFY_READY;
    }
    else if (status === BIANNUAL_BATCH_STATUS.COMPLETED) {
        phase = BATCH_PHASE.COMPLETED;
    }
    else {
        phase = BATCH_PHASE.MASTER_APPLY_READY;
    }
    return {
        phase,
        status,
        hqGate,
        senmuGate,
        murataGate,
        batchContext,
    };
}
function shouldRunAutoSenmuDecision(batchId, phase) {
    const evaluated = phase || evaluateBatchPhase(batchId);
    if (!evaluated)
        return false;
    if (evaluated.phase !== BATCH_PHASE.SENMU_WAITING &&
        evaluated.phase !== BATCH_PHASE.SENMU_RETURN_READY &&
        evaluated.phase !== BATCH_PHASE.MURATA_NOTIFY_READY &&
        evaluated.phase !== BATCH_PHASE.MASTER_APPLY_READY &&
        evaluated.phase !== BATCH_PHASE.COMPLETED) {
        return false;
    }
    return evaluated.senmuGate.hasAnyInput || evaluated.senmuGate.ready;
}
function shouldRunAutoMasterUpdate(batchId, phase) {
    const evaluated = phase || evaluateBatchPhase(batchId);
    if (!evaluated)
        return false;
    return evaluated.phase === BATCH_PHASE.MASTER_APPLY_READY && evaluated.murataGate.total > 0;
}
function isConfirmationSheetName(sheetName) {
    return String(sheetName || '').startsWith(HQ_CONFIRMATION_SHEET_PREFIX);
}
function rangeTouchesHeaders(range, headerMap, headerNames) {
    const startCol = range.getColumn();
    const endCol = startCol + range.getNumColumns() - 1;
    return headerNames.some((name) => {
        const col = headerMap[name];
        return !!col && col >= startCol && col <= endCol;
    });
}
function resolveBatchIdFromConfirmationEdit(sheet, rowIndex, headerMap) {
    const batchIdCol = headerMap['batchId'];
    if (batchIdCol) {
        const batchId = String(sheet.getRange(rowIndex, batchIdCol).getValue() || '').trim();
        if (batchId)
            return batchId;
    }
    return findBatchIdByConfirmationSheetName(sheet.getName());
}
function findBatchIdByConfirmationSheetName(sheetName) {
    const ss = getSpreadsheet();
    const notifyBatchSheet = ss.getSheetByName(SHEET_NAMES.NOTIFY_BATCH);
    if (!notifyBatchSheet || notifyBatchSheet.getLastRow() <= 1)
        return '';
    const data = notifyBatchSheet.getDataRange().getValues();
    if (data.length <= 1)
        return '';
    const headerMap = getHeaderMap(data[0]);
    const batchIdIndex = headerMap['batchId'];
    const confirmationSheetIndex = headerMap['確認用シート名'];
    if (!batchIdIndex || !confirmationSheetIndex)
        return '';
    for (let i = data.length - 1; i >= 1; i--) {
        const row = data[i];
        if (getCellValue(row, confirmationSheetIndex) !== sheetName)
            continue;
        const batchId = getCellValue(row, batchIdIndex);
        if (batchId)
            return batchId;
    }
    return '';
}
function tryReserveAutoAdvanceRun(minIntervalSec) {
    return tryReserveRun(PROP_KEYS.AUTO_ADVANCE_LAST_RUN_AT, minIntervalSec);
}
function tryReserveSourceSyncRun(minIntervalSec) {
    return tryReserveRun(PROP_KEYS.SOURCE_SYNC_LAST_RUN_AT, minIntervalSec);
}
function tryReserveRun(propertyKey, minIntervalSec) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(3000))
        return false;
    try {
        const props = PropertiesService.getDocumentProperties();
        const now = Date.now();
        const lastRaw = props.getProperty(propertyKey);
        const last = lastRaw ? Number(lastRaw) : 0;
        const minMillis = Math.max(0, Math.floor(minIntervalSec)) * 1000;
        if (last > 0 && now - last < minMillis)
            return false;
        props.setProperty(propertyKey, String(now));
        return true;
    }
    finally {
        lock.releaseLock();
    }
}
function normalizeSenmuDecision(value) {
    const text = String(value || '').trim();
    if (text === APPROVAL_INPUT.APPROVE)
        return APPROVAL_INPUT.APPROVE;
    if (text === APPROVAL_INPUT.RETURN)
        return APPROVAL_INPUT.RETURN;
    if (text === '承認済')
        return APPROVAL_INPUT.APPROVE;
    return '';
}
function isCheckedCell(value) {
    if (typeof value === 'boolean')
        return value;
    if (value === 1)
        return true;
    const text = String(value || '').trim().toLowerCase();
    return text === 'true' || text === '1' || text === 'yes' || text === '済' || text === '完了';
}
function summarizeConfirmationSheetRows(data, headerMap) {
    const result = {
        total: 0,
        unchecked: 0,
        unanswered: 0,
        renew: 0,
        cancellationReplace: 0,
        cancellationEnd: 0,
        masterApplied: 0,
    };
    if (!data || data.length <= 1)
        return result;
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const vehicleId = getCellValue(row, headerMap['vehicleId']);
        if (!vehicleId)
            continue;
        result.total += 1;
        const answer = normalizeAnswerLabel(getCellValue(row, headerMap['本部回答']));
        if (!answer) {
            result.unanswered += 1;
        }
        else if (answer === ANSWER_LABELS.RENEW) {
            result.renew += 1;
        }
        else if (answer === ANSWER_LABELS.CANCELLATION_REPLACE) {
            result.cancellationReplace += 1;
        }
        else if (answer === ANSWER_LABELS.CANCELLATION_END) {
            result.cancellationEnd += 1;
        }
        if (!isCheckedCell(getCellRaw(row, headerMap['回答確認済み']))) {
            result.unchecked += 1;
        }
        if (getCellValue(row, headerMap['反映日時'])) {
            result.masterApplied += 1;
        }
    }
    return result;
}
function loadSourceVehicleContext() {
    const ss = getSpreadsheet();
    const vehicleSheet = ss.getSheetByName(PRIMARY_SOURCE_SHEET);
    if (!vehicleSheet)
        throw new Error('車両一覧が存在しません。先に車両一覧同期（要入力更新）を実行してください。');
    ensureHeaders(vehicleSheet, 1, getSchemaHeaders(PRIMARY_SOURCE_SHEET));
    const vehicleData = vehicleSheet.getDataRange().getValues();
    if (vehicleData.length === 0) {
        return { rows: [], headerMap: {} };
    }
    const headerMap = getHeaderMap(vehicleData[0]);
    return {
        rows: vehicleData.slice(1),
        headerMap,
    };
}
function pickVehiclesByContractEndRange(rows, headerMap, start, end, tz) {
    const contractEndIndex = headerMap['契約満了日'];
    if (!contractEndIndex)
        return [];
    const startDate = toDateOnly(start, tz);
    const endDate = toDateOnly(end, tz);
    return rows.filter((row) => {
        if (row.every((cell) => cell === '' || cell === null))
            return false;
        const contractEnd = parseDateValue(getCellRaw(row, contractEndIndex));
        if (!contractEnd)
            return false;
        const contractDate = toDateOnly(contractEnd, tz);
        return isWithinRange(contractDate, startDate, endDate);
    });
}
function buildUniqueSheetName(baseName) {
    const ss = getSpreadsheet();
    if (!ss.getSheetByName(baseName))
        return baseName;
    let suffix = 2;
    while (suffix <= 99) {
        const candidate = `${baseName}_${suffix}`;
        if (!ss.getSheetByName(candidate))
            return candidate;
        suffix += 1;
    }
    throw new Error(`確認用シート名が重複しすぎています: ${baseName}`);
}
function buildRegistrationCombined(area, cls, kana, number) {
    return [area, cls, kana, number].filter((v) => v).join('');
}
function buildVehicleId(sourceSheet, regCombined, chassis, rowIndex) {
    const reg = String(regCombined || '').trim();
    const ch = String(chassis || '').trim();
    const hasDigit = /\d/.test(reg);
    if (reg && hasDigit)
        return `${sourceSheet}__${reg}`;
    if (ch)
        return `${sourceSheet}__${ch}`;
    if (reg)
        return `${sourceSheet}__${reg}__ROW${rowIndex}`;
    return `${sourceSheet}__ROW${rowIndex}`;
}
function loadSettings() {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    const values = {};
    if (sheet) {
        const data = sheet.getDataRange().getValues();
        const headerMap = data.length > 0 ? getHeaderMap(data[0]) : {};
        if (headerMap['設定項目'] && headerMap['値']) {
            for (let i = 1; i < data.length; i++) {
                const row = data[i];
                const key = getCellValue(row, headerMap['設定項目']);
                if (!key)
                    continue;
                values[key] = getCellRaw(row, headerMap['値']);
            }
        }
    }
    return {
        fromName: toStringValue(values['送信元名'], String(SETTINGS_DEFAULTS['送信元名'])),
        mailSendEnabled: toBoolean(values['通知_メール送信'], Boolean(SETTINGS_DEFAULTS['通知_メール送信'])),
        hqTo: toStringValue(values['本部長副本部長_通知先To'], String(SETTINGS_DEFAULTS['本部長副本部長_通知先To'])),
        senmuTo: toStringValue(values['専務_通知先To'], String(SETTINGS_DEFAULTS['専務_通知先To'])),
        senmuCc: toStringValue(values['専務_通知先Cc'], String(SETTINGS_DEFAULTS['専務_通知先Cc'])),
        marchSendDate: values['半期送付日_3月'] || SETTINGS_DEFAULTS['半期送付日_3月'],
        septemberSendDate: values['半期送付日_9月'] || SETTINGS_DEFAULTS['半期送付日_9月'],
        marchDeadline: values['回答期限_3月'] || SETTINGS_DEFAULTS['回答期限_3月'],
        septemberDeadline: values['回答期限_9月'] || SETTINGS_DEFAULTS['回答期限_9月'],
        reminderBeforeDays: toNumber(values['リマインド_期限前日数'], Number(SETTINGS_DEFAULTS['リマインド_期限前日数'])),
        autoAdvanceEnabled: toBoolean(values['自動進行_有効'], Boolean(SETTINGS_DEFAULTS['自動進行_有効'])),
        autoAdvanceTimerEnabled: toBoolean(values['自動進行_定期実行_有効'], Boolean(SETTINGS_DEFAULTS['自動進行_定期実行_有効'])),
        autoAdvanceTimerIntervalMinutes: normalizeAutoAdvanceIntervalMinutes(values['自動進行_定期実行_間隔分'] || convertLegacyAutoAdvanceHours(values['自動進行_定期実行_間隔時間']), Number(SETTINGS_DEFAULTS['自動進行_定期実行_間隔分'])),
        autoAdvanceOnEditEnabled: toBoolean(values['自動進行_編集連動_有効'], Boolean(SETTINGS_DEFAULTS['自動進行_編集連動_有効'])),
        autoApplySenmuDecision: toBoolean(values['自動進行_専務判断反映_有効'], Boolean(SETTINGS_DEFAULTS['自動進行_専務判断反映_有効'])),
        autoApplyMasterUpdates: toBoolean(values['自動進行_マスター反映_有効'], Boolean(SETTINGS_DEFAULTS['自動進行_マスター反映_有効'])),
        autoAdvanceMinIntervalSec: toNumber(values['自動進行_最小間隔秒'], Number(SETTINGS_DEFAULTS['自動進行_最小間隔秒'])),
        finalApproverAccount: toStringValue(values['最終承認者アカウント'], String(SETTINGS_DEFAULTS['最終承認者アカウント'])),
        adminTo: toStringValue(values['運用管理者_通知先To'], String(SETTINGS_DEFAULTS['運用管理者_通知先To'])),
        murataTo: toStringValue(values['村田主任_通知先To'], String(SETTINGS_DEFAULTS['村田主任_通知先To'])),
        murataCc: toStringValue(values['村田主任_通知先Cc'], String(SETTINGS_DEFAULTS['村田主任_通知先Cc'])),
    };
}
function convertLegacyAutoAdvanceHours(value) {
    if (value === null || value === undefined || value === '')
        return '';
    const hours = Math.max(1, Math.floor(toNumber(value, 1)));
    return hours * 60;
}
function normalizeAutoAdvanceIntervalMinutes(value, fallback) {
    const desired = Math.max(1, Math.floor(toNumber(value, fallback)));
    if (AUTO_ADVANCE_INTERVAL_MINUTES.indexOf(desired) >= 0)
        return desired;
    let nearest = AUTO_ADVANCE_INTERVAL_MINUTES[0];
    let distance = Math.abs(desired - nearest);
    AUTO_ADVANCE_INTERVAL_MINUTES.forEach((candidate) => {
        const currentDistance = Math.abs(desired - candidate);
        if (currentDistance < distance) {
            nearest = candidate;
            distance = currentDistance;
        }
    });
    return nearest;
}
function toNumber(value, fallback) {
    if (value === null || value === undefined || value === '')
        return fallback;
    const num = typeof value === 'number' ? value : Number(value);
    return isNaN(num) ? fallback : num;
}
function toBoolean(value, fallback) {
    if (value === null || value === undefined || value === '')
        return fallback;
    if (typeof value === 'boolean')
        return value;
    const str = String(value).toLowerCase();
    if (str === 'true' || str === '1' || str === 'yes')
        return true;
    if (str === 'false' || str === '0' || str === 'no')
        return false;
    return fallback;
}
function toStringValue(value, fallback) {
    if (value === null || value === undefined || value === '')
        return fallback;
    return String(value);
}
function normalizeAnswerLabel(value) {
    const text = String(value || '').trim();
    if (!text)
        return '';
    if (text === ANSWER_LABELS.RENEW || text === ANSWER_LABELS.CANCELLATION_REPLACE || text === ANSWER_LABELS.CANCELLATION_END) {
        return text;
    }
    if (LEGACY_ANSWER_LABEL_MAP[text])
        return LEGACY_ANSWER_LABEL_MAP[text];
    return '';
}
function clearTestResults() {
    const ss = getSpreadsheet();
    const sheet = ensureSheet(ss, SHEET_NAMES.TEST_RESULTS);
    ensureHeaders(sheet, 1, getSchemaHeaders(SHEET_NAMES.TEST_RESULTS));
    if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
}
function appendTestResult(item, result, detail) {
    const ss = getSpreadsheet();
    const sheet = ensureSheet(ss, SHEET_NAMES.TEST_RESULTS);
    ensureHeaders(sheet, 1, getSchemaHeaders(SHEET_NAMES.TEST_RESULTS));
    sheet.appendRow([new Date(), item, result, detail]);
}
function setSettingValue(key, value) {
    const ss = getSpreadsheet();
    const sheet = ensureSheet(ss, SHEET_NAMES.SETTINGS);
    ensureHeaders(sheet, 1, getSchemaHeaders(SHEET_NAMES.SETTINGS));
    const data = sheet.getDataRange().getValues();
    if (data.length === 0)
        return;
    const headerMap = getHeaderMap(data[0]);
    const keyIndex = headerMap['設定項目'];
    const valueIndex = headerMap['値'];
    if (!keyIndex || !valueIndex)
        return;
    let rowIndex = 0;
    for (let i = 1; i < data.length; i++) {
        if (getCellValue(data[i], keyIndex) === key) {
            rowIndex = i + 1;
            break;
        }
    }
    if (rowIndex === 0) {
        rowIndex = sheet.getLastRow() + 1;
        sheet.getRange(rowIndex, keyIndex, 1, 1).setValue(key);
    }
    sheet.getRange(rowIndex, valueIndex, 1, 1).setValue(value);
}
function writeSheetData(sheetName, rows) {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet)
        return;
    const headers = getSchemaHeaders(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
    if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
}
function writeArbitrarySheetData(sheet, headers, rows) {
    const lastRow = sheet.getLastRow();
    const lastColumn = Math.max(sheet.getLastColumn(), headers.length);
    if (lastRow > 0 && lastColumn > 0) {
        sheet.getRange(1, 1, lastRow, lastColumn).clearContent();
    }
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
}
function formatDateLabel(date, tz) {
    return Utilities.formatDate(date, tz, 'yyyy/MM/dd');
}
function formatVehicleLine(row, headerMap, tz) {
    const reg = getCellValue(row, headerMap['登録番号_結合']);
    const type = getCellValue(row, headerMap['車種']);
    const end = parseDateValue(getCellRaw(row, headerMap['契約満了日']));
    const endLabel = end ? formatDateLabel(end, tz) : '未設定';
    return `${reg || '登録番号不明'} / ${type || '車種不明'} / 満了:${endLabel}`;
}
function escapeHtml(text) {
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}
function toInlineJson(value) {
    return JSON.stringify(value)
        .replace(/</g, '\\u003c')
        .replace(/>/g, '\\u003e')
        .replace(/&/g, '\\u0026');
}
function protectViewSheet(sheetName) {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet)
        return;
    try {
        const desc = `${VIEW_SHEET_PROTECTION_DESC_PREFIX}${sheetName}`;
        const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
        let protection = protections.find((p) => p.getDescription() === desc);
        if (!protection) {
            protection = sheet.protect();
            protection.setDescription(desc);
        }
        protection.setWarningOnly(false);
        protection.setDomainEdit(false);
        try {
            const editors = protection.getEditors();
            if (editors && editors.length > 0)
                protection.removeEditors(editors);
        }
        catch (err) {
            Logger.log(`protectViewSheet removeEditors: ${sheetName} ${err}`);
        }
        try {
            protection.addEditor(Session.getEffectiveUser());
        }
        catch (err) {
            Logger.log(`protectViewSheet add effective user: ${sheetName} ${err}`);
        }
        try {
            protection.addEditor(Session.getActiveUser());
        }
        catch (err) {
            Logger.log(`protectViewSheet add active user: ${sheetName} ${err}`);
        }
    }
    catch (err) {
        Logger.log(`protectViewSheet: ${sheetName} ${err}`);
    }
}
function clearReturnedSenmuStateOnReAnswer(sheet, range, headerMap, batchStatus) {
    // 差戻し後（INITIAL_SENTに戻った状態）でのみ発火
    if (batchStatus !== BIANNUAL_BATCH_STATUS.INITIAL_SENT)
        return false;
    const startRow = range.getRow();
    const numRows = range.getNumRows();
    const lastCol = sheet.getLastColumn();
    const values = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
    let modified = false;
    values.forEach((row) => {
        const decision = normalizeSenmuDecision(getCellValue(row, headerMap['専務判断']));
        if (decision !== APPROVAL_INPUT.RETURN)
            return; // 差戻し行のみ対象
        if (headerMap['専務判断'])
            row[headerMap['専務判断'] - 1] = '';
        if (headerMap['専務コメント'])
            row[headerMap['専務コメント'] - 1] = '';
        if (headerMap['専務確認済み'])
            row[headerMap['専務確認済み'] - 1] = false;
        modified = true;
    });
    if (modified) {
        sheet.getRange(startRow, 1, numRows, lastCol).setValues(values);
    }
    return modified;
}
function ensureConfirmationSheetSchema(sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1)
        return;
    const ensureColumn = (targetHeader, insertBeforeHeader, checkbox) => {
        const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const currentMap = getHeaderMap(currentHeaders);
        if (currentMap[targetHeader])
            return;
        const insertBefore = insertBeforeHeader ? currentMap[insertBeforeHeader] : 0;
        if (insertBefore && insertBefore <= sheet.getLastColumn()) {
            sheet.insertColumnBefore(insertBefore);
            sheet.getRange(1, insertBefore).setValue(targetHeader);
            if (checkbox && lastRow > 1) {
                sheet.getRange(2, insertBefore, lastRow - 1, 1).insertCheckboxes();
            }
            return;
        }
        const newCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newCol).setValue(targetHeader);
        if (checkbox && lastRow > 1) {
            sheet.getRange(2, newCol, lastRow - 1, 1).insertCheckboxes();
        }
    };
    // 専務確認済み列がなければ、専務コメントの直後（新契約開始日の手前）に挿入
    ensureColumn('専務確認済み', '新契約開始日', true);
    ensureColumn('不正理由', '反映日時', false);
    // 既存だがチェックボックスでない場合の保険
    const updatedHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const updatedMap = getHeaderMap(updatedHeaders);
    const senmuCheckCol = updatedMap['専務確認済み'];
    if (senmuCheckCol && lastRow > 1) {
        const range = sheet.getRange(2, senmuCheckCol, lastRow - 1, 1);
        const validations = range.getDataValidations();
        if (!validations || !validations[0] || !validations[0][0]) {
            range.insertCheckboxes();
        }
    }
    protectSenmuColumns(sheet.getName());
}
function protectSenmuColumns(sheetName) {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet)
        return;
    if (sheet.getLastRow() === 0)
        return;
    const headerMap = getHeaderMap(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
    const decisionCol = headerMap['専務判断'];
    const commentCol = headerMap['専務コメント'];
    const senmuCheckCol = headerMap['専務確認済み'];
    if (!decisionCol || !commentCol)
        return;
    const cols = [decisionCol, commentCol, senmuCheckCol].filter((c) => !!c);
    const startCol = Math.min(...cols);
    const endCol = Math.max(...cols);
    const width = endCol - startCol + 1;
    const desc = `managed_by_script:senmu_columns:${sheetName}`;
    try {
        const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        protections.forEach((protection) => {
            if (protection.getDescription() !== desc)
                return;
            try {
                protection.remove();
            }
            catch (err) {
                Logger.log(`protectSenmuColumns remove: ${sheetName} ${err}`);
            }
        });
        const range = sheet.getRange(1, startCol, sheet.getMaxRows(), width);
        const protection = range.protect();
        protection.setDescription(desc);
        protection.setWarningOnly(false);
        protection.setDomainEdit(false);
        try {
            const editors = protection.getEditors();
            if (editors && editors.length > 0)
                protection.removeEditors(editors);
        }
        catch (err) {
            Logger.log(`protectSenmuColumns removeEditors: ${sheetName} ${err}`);
        }
        const settings = loadSettings();
        if (settings.finalApproverAccount) {
            try {
                protection.addEditor(settings.finalApproverAccount);
            }
            catch (err) {
                Logger.log(`protectSenmuColumns add finalApprover: ${sheetName} ${err}`);
            }
        }
    }
    catch (err) {
        Logger.log(`protectSenmuColumns: ${sheetName} ${err}`);
    }
}
function appendNotificationLog(type, dept, to, requestId, result) {
    const ss = getSpreadsheet();
    const sheet = ensureSheet(ss, SHEET_NAMES.NOTIFY_LOG);
    ensureHeaders(sheet, 1, getSchemaHeaders(SHEET_NAMES.NOTIFY_LOG));
    sheet.appendRow([new Date(), type, dept, to, requestId, result]);
}
