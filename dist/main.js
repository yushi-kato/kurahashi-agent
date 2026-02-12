/*
  車両リース契約 更新通知（GAS）
  - TypeScript で実装し、dist へビルドして clasp push する前提
*/
const SHEET_NAMES = {
    SETTINGS: '設定',
    VEHICLE_VIEW: '車両（統合ビュー）',
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
    '新契約開始日',
    '新契約満了日',
    '解約完了',
    'マスター反映済み',
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
        name: SHEET_NAMES.VEHICLE_VIEW,
        headerRow: 1,
        headers: [
            'vehicleId',
            'sourceSheet',
            '登録番号_地名',
            '登録番号_分類',
            '登録番号_かな',
            '登録番号_番号',
            '登録番号_結合',
            '車種',
            '車台番号',
            '契約開始日',
            '契約満了日',
            '管理部門',
            '管理担当者',
            '契約期間',
            '車検満了日',
            'リース料（税抜）',
            '更新方針',
            '依頼ID',
            '回答日',
            '備考',
            '一次回答',
            '最終決定',
            '完了フラグ',
            '完了日',
            '完了メモ',
        ],
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
};
const SCHEMA_VERSION = '1';
const PROP_KEYS = {
    SCHEMA_VERSION: 'SCHEMA_VERSION',
    LAST_SCHEMA_SYNC_AT: 'LAST_SCHEMA_SYNC_AT',
    LAST_SCHEMA_DRIFT_AT: 'LAST_SCHEMA_DRIFT_AT',
};
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('車両更新通知')
        .addItem('運用マニュアル（このシートで見る）', 'showOperationManual')
        .addItem('スキーマ同期', 'syncSchema')
        .addItem('スキーマドリフト確認', 'checkSchemaDrift')
        .addSeparator()
        .addItem('車両統合ビュー同期', 'syncVehicles')
        .addItem('半期バッチ起票', 'createBiannualBatch')
        .addItem('確認用シート生成（最新バッチ）', 'buildConfirmationSheetForLatestBatch')
        .addItem('初回通知送信（最新バッチ）', 'sendHqInitialEmail')
        .addItem('リマインド送信（条件一致時）', 'sendHqReminderIfNeeded')
        .addItem('専務依頼送信（全件確認後）', 'sendSenmuApprovalRequestIfReady')
        .addItem('専務判断反映（最新バッチ）', 'applySenmuDecisionFromSheet')
        .addItem('マスター反映（最新バッチ）', 'applyMasterUpdates')
        .addSeparator()
        .addItem('設定ひな形作成', 'seedSettings')
        .addItem('テスト一括実行(メール送信は設定次第)', 'runTestSuite')
        .addItem('テストデータ掃除', 'cleanupTestData')
        .addItem('半期トリガー再作成', 'installDailyTriggers')
        .addSeparator()
        .addItem('半期一括実行', 'runDaily')
        .addToUi();
}
function showOperationManual() {
    const ui = SpreadsheetApp.getUi();
    const html = HtmlService.createHtmlOutputFromFile('operation_manual_vehicle_lease_renewal')
        .setWidth(1000)
        .setHeight(800);
    ui.showModalDialog(html, '運用マニュアル');
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
        const vehicleViewSheet = ensureSheet(ss, SHEET_NAMES.VEHICLE_VIEW);
        ensureHeaders(vehicleViewSheet, 1, getSchemaHeaders(SHEET_NAMES.VEHICLE_VIEW));
        ensureHeaders(ensureSheet(ss, SHEET_NAMES.NEEDS_INPUT), 1, getSchemaHeaders(SHEET_NAMES.NEEDS_INPUT));
        // 統合ビューは再生成するが、依頼/回答などの運用列は vehicleId キーで引き継ぐ
        const existingByVehicleId = {};
        const existingData = vehicleViewSheet.getDataRange().getValues();
        if (existingData.length > 1) {
            const existingHeader = getHeaderMap(existingData[0]);
            const idxExisting = {
                vehicleId: existingHeader['vehicleId'],
                sourceSheet: existingHeader['sourceSheet'],
                regCombined: existingHeader['登録番号_結合'],
                chassis: existingHeader['車台番号'],
                policy: existingHeader['更新方針'],
                requestId: existingHeader['依頼ID'],
                answeredAt: existingHeader['回答日'],
                note: existingHeader['備考'],
                primaryAnswer: existingHeader['一次回答'],
                finalDecision: existingHeader['最終決定'],
                completedFlag: existingHeader['完了フラグ'],
                completedAt: existingHeader['完了日'],
                completedMemo: existingHeader['完了メモ'],
            };
            if (idxExisting.vehicleId) {
                for (let i = 1; i < existingData.length; i++) {
                    const row = existingData[i];
                    const vehicleId = getCellValue(row, idxExisting.vehicleId);
                    if (!vehicleId)
                        continue;
                    const record = {
                        policy: getCellValue(row, idxExisting.policy),
                        requestId: getCellValue(row, idxExisting.requestId),
                        answeredAt: getCellRaw(row, idxExisting.answeredAt),
                        note: getCellValue(row, idxExisting.note),
                        primaryAnswer: getCellValue(row, idxExisting.primaryAnswer),
                        finalDecision: getCellValue(row, idxExisting.finalDecision),
                        completedFlag: getCellValue(row, idxExisting.completedFlag),
                        completedAt: getCellRaw(row, idxExisting.completedAt),
                        completedMemo: getCellValue(row, idxExisting.completedMemo),
                    };
                    existingByVehicleId[vehicleId] = record;
                    const sourceSheet = getCellValue(row, idxExisting.sourceSheet);
                    const regCombined = getCellValue(row, idxExisting.regCombined);
                    const chassis = getCellValue(row, idxExisting.chassis);
                    if (sourceSheet && chassis) {
                        const key = `${sourceSheet}__${chassis}`;
                        if (!existingByVehicleId[key])
                            existingByVehicleId[key] = record;
                    }
                    if (sourceSheet && regCombined && /\d/.test(regCombined)) {
                        const key = `${sourceSheet}__${regCombined}`;
                        if (!existingByVehicleId[key])
                            existingByVehicleId[key] = record;
                    }
                }
            }
        }
        const rows = [];
        const needsInputRows = [];
        const now = new Date();
        const tz = ss.getSpreadsheetTimeZone();
        const seenVehicleId = {};
        const sheetName = PRIMARY_SOURCE_SHEET;
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
            needsInputRows.push([now, sheetName, '', '', '', '', '対象シートが存在しません']);
        }
        else {
            const data = sheet.getDataRange().getValues();
            if (data.length > 1) {
                const headers = data[0];
                const headerMap = getHeaderMap(headers);
                const headerIndexes = resolveSourceHeaders(headerMap);
                if (!headerIndexes.contractEnd || !headerIndexes.dept) {
                    needsInputRows.push([now, sheetName, '', '', '', '', '必要ヘッダが不足しています']);
                }
                else {
                    for (let i = 1; i < data.length; i++) {
                        const row = data[i];
                        if (row.every((cell) => cell === '' || cell === null))
                            continue;
                        const regParts = getSourceRegistrationParts(row, headerIndexes);
                        const regCombined = getSourceRegistrationCombined(row, headerIndexes);
                        const vehicleType = getCellValue(row, headerIndexes.vehicleType);
                        const chassis = getCellValue(row, headerIndexes.chassis);
                        const contractStart = parseDateValue(getCellRaw(row, headerIndexes.contractStart));
                        const contractEnd = parseDateValue(getCellRaw(row, headerIndexes.contractEnd));
                        const dept = getCellValue(row, headerIndexes.dept);
                        const manager = getCellValue(row, headerIndexes.manager);
                        const contractTerm = getCellValue(row, headerIndexes.contractTerm);
                        const inspectionEnd = parseDateValue(getCellRaw(row, headerIndexes.inspectionEnd));
                        const leaseFee = getCellValue(row, headerIndexes.leaseFee);
                        const vehicleId = buildVehicleId(sheetName, regCombined, chassis, i + 1);
                        const existing = existingByVehicleId[vehicleId] || {
                            policy: '',
                            requestId: '',
                            answeredAt: '',
                            note: '',
                            primaryAnswer: '',
                            finalDecision: '',
                            completedFlag: '',
                            completedAt: '',
                            completedMemo: '',
                        };
                        if (seenVehicleId[vehicleId]) {
                            const prev = seenVehicleId[vehicleId];
                            needsInputRows.push([
                                now,
                                sheetName,
                                vehicleId,
                                dept,
                                regCombined,
                                vehicleType,
                                `vehicleId重複（先頭: ${prev.sheet} 行${prev.rowIndex} / 今回: 行${i + 1}）`,
                            ]);
                        }
                        else {
                            seenVehicleId[vehicleId] = { sheet: sheetName, rowIndex: i + 1 };
                        }
                        if (!contractEnd) {
                            needsInputRows.push([now, sheetName, vehicleId, dept, regCombined, vehicleType, '契約満了日なし']);
                        }
                        if (!dept) {
                            needsInputRows.push([now, sheetName, vehicleId, dept, regCombined, vehicleType, '管理部門なし']);
                        }
                        if (regCombined && !/\d/.test(String(regCombined))) {
                            if (!chassis) {
                                needsInputRows.push([
                                    now,
                                    sheetName,
                                    vehicleId,
                                    dept,
                                    regCombined,
                                    vehicleType,
                                    '登録番号が不完全（数字なし）かつ車台番号なし',
                                ]);
                            }
                            else {
                                needsInputRows.push([now, sheetName, vehicleId, dept, regCombined, vehicleType, '登録番号が不完全（数字なし）']);
                            }
                        }
                        const primaryAnswer = existing.primaryAnswer || existing.policy;
                        rows.push([
                            vehicleId,
                            sheetName,
                            regParts.area,
                            regParts.cls,
                            regParts.kana,
                            regParts.num,
                            regCombined,
                            vehicleType,
                            chassis,
                            contractStart,
                            contractEnd,
                            dept,
                            manager,
                            contractTerm,
                            inspectionEnd,
                            leaseFee,
                            existing.policy,
                            existing.requestId,
                            existing.answeredAt,
                            existing.note,
                            primaryAnswer,
                            existing.finalDecision,
                            existing.completedFlag,
                            existing.completedAt,
                            existing.completedMemo,
                        ]);
                    }
                }
            }
        }
        writeSheetData(SHEET_NAMES.VEHICLE_VIEW, rows);
        writeSheetData(SHEET_NAMES.NEEDS_INPUT, needsInputRows);
        protectViewSheet(SHEET_NAMES.VEHICLE_VIEW);
        protectViewSheet(SHEET_NAMES.NEEDS_INPUT);
    }
    finally {
        lock.releaseLock();
    }
}
function createBiannualBatch() {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const ss = getSpreadsheet();
        ensureHeaders(ensureSheet(ss, SHEET_NAMES.NOTIFY_BATCH), 1, getSchemaHeaders(SHEET_NAMES.NOTIFY_BATCH));
        ensureHeaders(ensureSheet(ss, SHEET_NAMES.VEHICLE_VIEW), 1, getSchemaHeaders(SHEET_NAMES.VEHICLE_VIEW));
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
            uiAlertSafe(`通知バッチ ${batchDef.batchId} は既に存在します。`);
            return batchDef.batchId;
        }
        const vehicleContext = loadVehicleViewContext();
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
        uiAlertSafe(`半期バッチを起票しました。\n` +
            `batchId: ${batchDef.batchId}\n` +
            `対象期間: ${formatDateLabel(batchDef.targetStart, tz)}〜${formatDateLabel(batchDef.targetEnd, tz)}\n` +
            `対象件数: ${targetVehicles.length}件\n` +
            `確認用シート: ${confirmationSheetName}`);
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
    const vehicleContext = loadVehicleViewContext();
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
        '',
        '',
        false,
        false,
        '',
    ]);
    writeArbitrarySheetData(sheet, HQ_CONFIRMATION_HEADERS, rows);
    const headerMap = getHeaderMap(HQ_CONFIRMATION_HEADERS);
    if (rows.length > 0) {
        const checkColumns = ['回答確認済み', '解約完了', 'マスター反映済み'];
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
function sendHqInitialEmail(batchId) {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const settings = loadSettings();
        const batchContext = getNotifyBatchContext(batchId);
        if (!batchContext) {
            uiAlertSafe('通知バッチが見つかりません。先に半期バッチ起票を実行してください。');
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
            uiAlertSafe('設定「本部長副本部長_通知先To」が未設定のため送信できません。');
            return resolvedBatchId;
        }
        const sentAt = parseDateValue(getCellRaw(row, headerMap['初回通知送信日時']));
        if (sentAt) {
            uiAlertSafe(`初回通知は既に送信済みです。\nbatchId: ${resolvedBatchId}`);
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
        const subject = `【車両更新確認】${batchLabel} 一次確認のお願い`;
        const body = [
            '本部長・副本部長 各位',
            '',
            `${batchLabel} の車両更新確認をお願いします。`,
            `batchId: ${resolvedBatchId}`,
            `対象期間: ${formatDateLabel(targetStart || new Date(), tz)}〜${formatDateLabel(targetEnd || new Date(), tz)}`,
            `回答期限: ${formatDateLabel(deadline || new Date(), tz)}`,
            '',
            `対象件数: ${counts.total}`,
            `未回答件数: ${counts.unanswered}`,
            '',
            '確認用シート:',
            sheetUrl,
            '',
            '入力ルール:',
            `- 「本部回答」は ${ANSWER_OPTIONS.join(' / ')} から選択してください。`,
            '- 全件入力後に「回答確認済み」をチェックしてください。',
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
            uiAlertSafe(`初回通知を送信しました。\nbatchId: ${resolvedBatchId}`);
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
        const subject = `【リマインド】${batchLabel} 回答確認のお願い`;
        const body = [
            '本部長・副本部長 各位',
            '',
            `${batchLabel} の確認用シートで、未確認行が残っています。`,
            `batchId: ${resolvedBatchId}`,
            `回答期限: ${formatDateLabel(deadline, tz)}`,
            '',
            `未確認件数: ${counts.unchecked}`,
            `未回答件数: ${counts.unanswered}`,
            '',
            '確認用シート:',
            sheetUrl,
            '',
            '※本メールは「期限前リマインド」の1回送信です。',
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
        const batchContext = getNotifyBatchContext(batchId);
        if (!batchContext)
            return '';
        const { ss, notifyBatchSheet, batchData, headerMap, row } = batchContext;
        const resolvedBatchId = batchContext.batchId;
        const tz = ss.getSpreadsheetTimeZone();
        const senmuTo = String(settings.senmuTo || '').trim();
        const sentAt = parseDateValue(getCellRaw(row, headerMap['専務依頼送信日時']));
        if (sentAt)
            return resolvedBatchId;
        const confirmationSheetName = getCellValue(row, headerMap['確認用シート名']);
        if (!confirmationSheetName) {
            appendNotificationLog('専務依頼', '', '', resolvedBatchId, '確認用シート未設定のためスキップ');
            return resolvedBatchId;
        }
        const confirmationSheet = ss.getSheetByName(confirmationSheetName);
        if (!confirmationSheet) {
            appendNotificationLog('専務依頼', '', '', resolvedBatchId, `確認用シート不在: ${confirmationSheetName}`);
            return resolvedBatchId;
        }
        const confirmationData = confirmationSheet.getDataRange().getValues();
        const confirmationHeader = confirmationData.length > 0 ? getHeaderMap(confirmationData[0]) : {};
        const counts = summarizeConfirmationSheetRows(confirmationData, confirmationHeader);
        if (counts.total <= 0)
            return resolvedBatchId;
        if (counts.unchecked > 0) {
            appendNotificationLog('専務依頼', '', '', resolvedBatchId, `未確認行あり(${counts.unchecked}件)のため未送信`);
            return resolvedBatchId;
        }
        if (counts.unanswered > 0) {
            appendNotificationLog('専務依頼', '', '', resolvedBatchId, `未回答行あり(${counts.unanswered}件)のため未送信`);
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
        const subject = `【専務確認依頼】${batchLabel} 車両更新方針`;
        const body = [
            '専務',
            '',
            `${batchLabel} の一次確認が完了しました。`,
            `batchId: ${resolvedBatchId}`,
            `回答期限: ${formatDateLabel(deadline || new Date(), tz)}`,
            '',
            `対象件数: ${counts.total}`,
            `更新: ${counts.renew}`,
            `解約（入替）: ${counts.cancellationReplace}`,
            `解約（満了）: ${counts.cancellationEnd}`,
            '',
            '確認用シート（専務判断列を入力してください）:',
            sheetUrl,
            '',
            `専務判断の入力値: ${APPROVAL_INPUT.APPROVE} / ${APPROVAL_INPUT.RETURN}`,
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
            protectSenmuColumns(confirmationSheetName);
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
        const batchContext = getNotifyBatchContext(batchId);
        if (!batchContext)
            return '';
        const { notifyBatchSheet, batchData, headerMap, row } = batchContext;
        const resolvedBatchId = batchContext.batchId;
        const confirmationSheetName = getCellValue(row, headerMap['確認用シート名']);
        if (!confirmationSheetName) {
            uiAlertSafe(`確認用シート名が未設定です。\nbatchId: ${resolvedBatchId}`);
            return resolvedBatchId;
        }
        const ss = getSpreadsheet();
        const sheet = ss.getSheetByName(confirmationSheetName);
        if (!sheet) {
            uiAlertSafe(`確認用シートが見つかりません。\n${confirmationSheetName}`);
            return resolvedBatchId;
        }
        protectSenmuColumns(confirmationSheetName);
        const data = sheet.getDataRange().getValues();
        if (data.length <= 1)
            return resolvedBatchId;
        const h = getHeaderMap(data[0]);
        const decisionIndex = h['専務判断'];
        if (!decisionIndex)
            return resolvedBatchId;
        let approved = 0;
        let returned = 0;
        let pending = 0;
        let invalid = 0;
        const invalidRows = [];
        for (let i = 1; i < data.length; i++) {
            const rowData = data[i];
            const vehicleId = getCellValue(rowData, h['vehicleId']);
            if (!vehicleId)
                continue;
            const decision = normalizeSenmuDecision(getCellValue(rowData, decisionIndex));
            if (!decision) {
                const raw = getCellValue(rowData, decisionIndex);
                if (raw) {
                    invalid += 1;
                    invalidRows.push(`行${i + 1}: ${raw}`);
                }
                else {
                    pending += 1;
                }
                continue;
            }
            if (decision === APPROVAL_INPUT.APPROVE)
                approved += 1;
            if (decision === APPROVAL_INPUT.RETURN)
                returned += 1;
        }
        const now = new Date();
        if (returned > 0) {
            row[headerMap['ステータス'] - 1] = BIANNUAL_BATCH_STATUS.SENMU_RETURNED;
        }
        else if (approved > 0 && pending === 0 && invalid === 0) {
            row[headerMap['ステータス'] - 1] = BIANNUAL_BATCH_STATUS.SENMU_APPROVED;
        }
        else {
            row[headerMap['ステータス'] - 1] = BIANNUAL_BATCH_STATUS.SENMU_REQUESTED;
        }
        row[headerMap['更新日時'] - 1] = now;
        notifyBatchSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);
        appendNotificationLog('専務判断反映', '', '', resolvedBatchId, `承認:${approved} 差戻し:${returned} 保留:${pending} 不正:${invalid}`);
        const lines = [
            `batchId: ${resolvedBatchId}`,
            `承認: ${approved}`,
            `差戻し: ${returned}`,
            `保留: ${pending}`,
            `不正入力: ${invalid}`,
        ];
        if (invalidRows.length > 0) {
            lines.push('', '不正入力明細:', ...invalidRows.slice(0, 10));
        }
        uiShowModalSafe('専務判断反映', lines.join('\n'));
        return resolvedBatchId;
    }
    finally {
        lock.releaseLock();
    }
}
function applyMasterUpdates(batchId) {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const batchContext = getNotifyBatchContext(batchId);
        if (!batchContext)
            return '';
        const { ss, notifyBatchSheet, batchData, headerMap, row } = batchContext;
        const resolvedBatchId = batchContext.batchId;
        const tz = ss.getSpreadsheetTimeZone();
        const confirmationSheetName = getCellValue(row, headerMap['確認用シート名']);
        if (!confirmationSheetName) {
            uiAlertSafe(`確認用シート名が未設定です。\nbatchId: ${resolvedBatchId}`);
            return resolvedBatchId;
        }
        const confirmationSheet = ss.getSheetByName(confirmationSheetName);
        if (!confirmationSheet) {
            uiAlertSafe(`確認用シートが見つかりません。\n${confirmationSheetName}`);
            return resolvedBatchId;
        }
        const vehicleSheet = ss.getSheetByName(SHEET_NAMES.VEHICLE_VIEW);
        if (!vehicleSheet)
            throw new Error('車両（統合ビュー）が存在しません');
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
        for (let i = 1; i < confirmationData.length; i++) {
            const cRow = confirmationData[i];
            const vehicleId = getCellValue(cRow, ch['vehicleId']);
            if (!vehicleId)
                continue;
            const decision = normalizeSenmuDecision(getCellValue(cRow, ch['専務判断']));
            if (decision === APPROVAL_INPUT.RETURN) {
                returned += 1;
                continue;
            }
            if (decision !== APPROVAL_INPUT.APPROVE) {
                waiting += 1;
                continue;
            }
            if (isCheckedCell(getCellRaw(cRow, ch['マスター反映済み'])))
                continue;
            const policy = normalizeAnswerLabel(getCellValue(cRow, ch['本部回答']));
            if (!policy) {
                skipped += 1;
                continue;
            }
            const vehicleRowIndex = vehicleRowIndexById[vehicleId];
            if (vehicleRowIndex === undefined) {
                skipped += 1;
                continue;
            }
            const vehicleRow = vehicleData[vehicleRowIndex];
            const setVehicle = (headerName, value) => {
                const idx = vh[headerName];
                if (idx)
                    vehicleRow[idx - 1] = value;
            };
            if (policy === ANSWER_LABELS.RENEW) {
                const newStart = parseDateValue(getCellRaw(cRow, ch['新契約開始日']));
                const newEnd = parseDateValue(getCellRaw(cRow, ch['新契約満了日']));
                if (!newStart || !newEnd) {
                    skipped += 1;
                    continue;
                }
                setVehicle('契約開始日', toDateOnly(newStart, tz));
                setVehicle('契約満了日', toDateOnly(newEnd, tz));
                setVehicle('更新方針', policy);
                setVehicle('一次回答', policy);
                setVehicle('最終決定', policy);
                setVehicle('完了フラグ', true);
                setVehicle('完了日', now);
                setVehicle('完了メモ', '半期バッチで更新反映');
            }
            else {
                const cancelDone = isCheckedCell(getCellRaw(cRow, ch['解約完了']));
                if (!cancelDone) {
                    skipped += 1;
                    continue;
                }
                setVehicle('更新方針', policy);
                setVehicle('一次回答', policy);
                setVehicle('最終決定', policy);
                setVehicle('完了フラグ', true);
                setVehicle('完了日', now);
                setVehicle('完了メモ', '半期バッチで解約反映');
                rowIndexesToGray.push(vehicleRowIndex + 1);
            }
            if (ch['マスター反映済み'])
                cRow[ch['マスター反映済み'] - 1] = true;
            if (ch['反映日時'])
                cRow[ch['反映日時'] - 1] = now;
            applied += 1;
            modifiedVehicle = true;
            modifiedConfirmation = true;
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
        if (counts.total > 0 && counts.masterApplied >= counts.total && returned === 0) {
            row[headerMap['ステータス'] - 1] = BIANNUAL_BATCH_STATUS.COMPLETED;
        }
        else if (returned > 0) {
            row[headerMap['ステータス'] - 1] = BIANNUAL_BATCH_STATUS.SENMU_RETURNED;
        }
        else {
            row[headerMap['ステータス'] - 1] = BIANNUAL_BATCH_STATUS.SENMU_APPROVED;
        }
        row[headerMap['更新日時'] - 1] = now;
        notifyBatchSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);
        appendNotificationLog('マスター反映', '', '', resolvedBatchId, `反映:${applied} 待機:${waiting} 差戻し:${returned} スキップ:${skipped}`);
        uiShowModalSafe('マスター反映', [
            `batchId: ${resolvedBatchId}`,
            `反映: ${applied}`,
            `待機（専務未承認）: ${waiting}`,
            `差戻し: ${returned}`,
            `スキップ（入力不足・不整合）: ${skipped}`,
        ].join('\n'));
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
    // 旧関数名を残しつつ、実行内容は半期バッチ導線へ切り替える。
    runBiannualSchedule();
}
function runBiannualSchedule() {
    syncSchema();
    syncVehicles();
    createBiannualBatch();
    sendHqInitialEmail();
    sendHqReminderIfNeeded();
    sendSenmuApprovalRequestIfReady();
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
        if (!existingKeys[key]) {
            rows.push([key, SETTINGS_DEFAULTS[key], '']);
        }
    });
    if (rows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, descIndex ? 3 : 2).setValues(rows);
    }
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
function cleanupTestData() {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const ss = getSpreadsheet();
        const removed = {
            sourceSheets: {},
            vehicleView: 0,
            needsInput: 0,
        };
        const testVehicleIds = {};
        // 車両（統合ビュー）からテスト由来の vehicleId を収集しつつ削除
        const vehicleViewSheet = ss.getSheetByName(SHEET_NAMES.VEHICLE_VIEW);
        if (vehicleViewSheet) {
            const data = vehicleViewSheet.getDataRange().getValues();
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
                    const isTest = (regCombined && regCombined.startsWith('TEST')) || (vehicleId && vehicleId.indexOf('__TEST') >= 0);
                    if (!isTest)
                        continue;
                    if (vehicleId)
                        testVehicleIds[vehicleId] = true;
                    rowsToDelete.push(i + 1);
                }
                for (let i = rowsToDelete.length - 1; i >= 0; i--) {
                    vehicleViewSheet.deleteRow(rowsToDelete[i]);
                    removed.vehicleView += 1;
                }
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
                    const isTest = (vehicleId && testVehicleIds[vehicleId]) || (regCombined && regCombined.startsWith('TEST'));
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
        // 元台帳（車両一覧）からテスト車両行を削除
        const sourceSheet = ss.getSheetByName(PRIMARY_SOURCE_SHEET);
        if (sourceSheet) {
            const data = sourceSheet.getDataRange().getValues();
            if (data.length > 1) {
                const headerMap = getHeaderMap(data[0]);
                const idx = resolveSourceHeaders(headerMap);
                const rowsToDelete = [];
                for (let i = 1; i < data.length; i++) {
                    const row = data[i];
                    if (row.every((cell) => cell === '' || cell === null))
                        continue;
                    const regCombined = getSourceRegistrationCombined(row, idx);
                    const chassis = getCellValue(row, idx.chassis);
                    const vehicleType = getCellValue(row, idx.vehicleType);
                    const isTest = (regCombined && String(regCombined).startsWith('TEST')) ||
                        (chassis && String(chassis).startsWith('TEST-')) ||
                        (vehicleType && String(vehicleType).startsWith('テスト_'));
                    if (!isTest)
                        continue;
                    rowsToDelete.push(i + 1);
                }
                for (let i = rowsToDelete.length - 1; i >= 0; i--) {
                    sourceSheet.deleteRow(rowsToDelete[i]);
                }
                removed.sourceSheets[PRIMARY_SOURCE_SHEET] = rowsToDelete.length;
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
    }
    else {
        appendTestResult('期待値:確認用シート生成件数', 'NG', '確認用シート未生成');
    }
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
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach((trigger) => {
        const handler = trigger.getHandlerFunction();
        if (handler === 'runDaily' || handler === 'runBiannualSchedule') {
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
        if (isCheckedCell(getCellRaw(row, headerMap['マスター反映済み']))) {
            result.masterApplied += 1;
        }
    }
    return result;
}
function loadVehicleViewContext() {
    const ss = getSpreadsheet();
    const vehicleSheet = ss.getSheetByName(SHEET_NAMES.VEHICLE_VIEW);
    if (!vehicleSheet)
        throw new Error('車両（統合ビュー）が存在しません。先に車両統合ビュー同期を実行してください。');
    ensureHeaders(vehicleSheet, 1, getSchemaHeaders(SHEET_NAMES.VEHICLE_VIEW));
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
    };
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
    if (!decisionCol || !commentCol)
        return;
    const startCol = Math.min(decisionCol, commentCol);
    const endCol = Math.max(decisionCol, commentCol);
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
        try {
            protection.addEditor(Session.getEffectiveUser());
        }
        catch (err) {
            Logger.log(`protectSenmuColumns add effective user: ${sheetName} ${err}`);
        }
        try {
            protection.addEditor(Session.getActiveUser());
        }
        catch (err) {
            Logger.log(`protectSenmuColumns add active user: ${sheetName} ${err}`);
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
