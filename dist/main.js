/*
  車両リース契約 更新通知（GAS）
  - TypeScript で実装し、dist へビルドして clasp push する前提
*/
const SHEET_NAMES = {
    SETTINGS: '設定',
    DEPT_MASTER: '部署マスタ',
    VEHICLE_VIEW: '車両（統合ビュー）',
    NEEDS_INPUT: '要入力',
    REQUESTS: '更新依頼',
    ANSWERS: '回答',
    NOTIFY_LOG: '通知ログ',
    SUMMARY: '回答集計',
    APPROVAL_QUEUE: '承認待ち一覧',
    TEST_RESULTS: 'テスト結果',
};
const SOURCE_SHEETS = ['車両一覧', '車両一覧【ｹﾝｽｲ】', '車両一覧【ﾈｸｽﾄ】'];
const REQUEST_STATUS = {
    CREATED: '作成済',
    SENT: '送信済',
    RESPONDING: '回答中',
    COMPLETED: '完了',
    EXPIRED: '締切',
};
const APPROVAL_STATUS = {
    NOT_SENT: '未送付',
    PENDING: '承認待ち',
    APPROVED: '承認済',
    RETURNED: '差戻し',
};
const APPROVAL_INPUT = {
    APPROVE: '承認',
    RETURN: '差戻し',
};
const APPROVAL_FORM_TITLES = {
    DECISION: '承認判断',
    COMMENT: '差戻しコメント（差戻し時のみ）',
};
const APPROVAL_FORM_REQUEST_ID_PROP_PREFIX = 'APPROVAL_FORM_REQUEST_ID__';
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
const MAX_VEHICLES_PER_FORM = 50;
const FORM_ITEM_TITLES = {
    POLICY_GRID: '更新方針（車両ごと）',
};
const FORM_VEHICLE_IDS_PROP_PREFIX = 'FORM_VEHICLE_IDS__';
const VIEW_SHEET_PROTECTION_DESC_PREFIX = 'managed_by_script:view_sheet:';
const SCHEMA_DEFS = [
    {
        name: SHEET_NAMES.SETTINGS,
        headerRow: 1,
        headers: ['設定項目', '値', '説明'],
    },
    {
        name: SHEET_NAMES.DEPT_MASTER,
        headerRow: 1,
        headers: ['管理部門', '通知先To', '通知先Cc', '有効', '部門トークン'],
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
        name: SHEET_NAMES.REQUESTS,
        headerRow: 1,
        headers: [
            'requestId',
            '管理部門',
            '対象開始日',
            '対象終了日',
            '締切日',
            'ステータス',
            '初回送信日時',
            '最終リマインド日時',
            'リマインド回数',
            'requestToken',
            'formId',
            'formUrl',
            'formIdsJson',
            'formUrlsJson',
            'formEditUrl',
            'formTriggerId',
            'フォーム作成日時',
            '承認ステータス',
            '承認依頼送信日時',
            '承認者',
            '承認日時',
            '差戻しコメント',
            '車両管理通知送信日時',
            '承認フォームID',
            '承認フォームURL',
            '承認フォーム編集URL',
            '承認フォームトリガーID',
            '承認フォーム作成日時',
        ],
    },
    {
        name: SHEET_NAMES.ANSWERS,
        headerRow: 1,
        headers: ['requestId', 'vehicleId', '回答', 'コメント', '回答者', '回答日時'],
    },
    {
        name: SHEET_NAMES.NOTIFY_LOG,
        headerRow: 1,
        headers: ['日時', '種別', '管理部門', '宛先', 'requestId', '結果'],
    },
    {
        name: SHEET_NAMES.SUMMARY,
        headerRow: 1,
        headers: [
            'requestId',
            '管理部門',
            '対象期間',
            '総件数',
            ANSWER_LABELS.RENEW,
            ANSWER_LABELS.CANCELLATION_REPLACE,
            ANSWER_LABELS.CANCELLATION_END,
            '未回答',
            '最終更新日時',
        ],
    },
    {
        name: SHEET_NAMES.APPROVAL_QUEUE,
        headerRow: 1,
        headers: [
            'requestId',
            '管理部門',
            '対象開始日',
            '対象終了日',
            '締切日',
            '依頼ステータス',
            '承認ステータス',
            '総件数',
            '回答済件数',
            '一次結果サマリ',
            '対象車両一覧',
            '承認入力',
            '差戻しコメント',
            '承認者',
            '承認日時',
            '承認依頼送信日時',
            '車両管理通知送信日時',
        ],
    },
    {
        name: SHEET_NAMES.TEST_RESULTS,
        headerRow: 1,
        headers: ['実行日時', '項目', '結果', '詳細'],
    },
];
const SETTINGS_DEFAULTS = {
    抽出_満了まで月数: 6,
    リマインド_初回から日数: 7,
    リマインド_間隔日数: 7,
    リマインド_最大回数: 2,
    締切_初回送信から日数: 14,
    送信元名: '車両管理システム',
    件名テンプレ: '【車両更新確認】{{管理部門}} 対象: {{対象開始日}}〜{{対象終了日}}',
    本文テンプレ: '{{管理部門}} 各位\n\n以下の車両について更新方針をご回答ください。\n対象期間: {{対象開始日}}〜{{対象終了日}}\n締切: {{締切日}}\n\n回答URL:\n{{URL}}\n\n対象車両:\n{{車両一覧}}',
    'Web回答URL（デプロイURL）': '',
    管理者_通知先To: '',
    管理者_通知先Cc: '',
    通知_メール送信: true,
    集計_シート出力: true,
    集計_メール送信: true,
    承認フロー_有効: false,
    承認者_通知先To: '',
    承認者_通知先Cc: '',
    車両管理担当_通知先To: '',
    車両管理担当_通知先Cc: '',
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
        .addItem('更新依頼作成', 'createRequests')
        .addItem('初回メール送信', 'sendInitialEmails')
        .addItem('リマインド送信', 'sendReminderEmails')
        .addSeparator()
        .addItem('回答反映', 'applyAnswers')
        .addItem('回答集計更新', 'buildSummarySheet')
        .addItem('集計メール送信', 'sendSummaryEmail')
        .addSeparator()
        .addItem('承認待ち一覧更新', 'buildApprovalQueueSheet')
        .addItem('承認依頼メール送信', 'sendApprovalRequestEmails')
        .addItem('承認結果反映（通知送信）', 'applyApprovalDecisions')
        .addSeparator()
        .addItem('設定ひな形作成', 'seedSettings')
        .addItem('部署トークン生成(空欄のみ)', 'generateDeptTokens')
        .addItem('テスト車両追加', 'seedTestVehicles')
        .addItem('ソースシート診断', 'diagnoseSourceSheets')
        .addItem('テスト一括実行(メール送信は設定次第)', 'runTestSuite')
        .addItem('テストデータ掃除', 'cleanupTestData')
        .addItem('日次トリガー再作成', 'installDailyTriggers')
        .addSeparator()
        .addItem('日次一括実行', 'runDaily')
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
        const deptMaster = loadDeptMaster();
        const rows = [];
        const needsInputRows = [];
        const now = new Date();
        const tz = ss.getSpreadsheetTimeZone();
        const seenVehicleId = {};
        SOURCE_SHEETS.forEach((sheetName) => {
            const sheet = ss.getSheetByName(sheetName);
            if (!sheet) {
                needsInputRows.push([now, sheetName, '', '', '', '', '対象シートが存在しません']);
                return;
            }
            const data = sheet.getDataRange().getValues();
            if (data.length <= 1)
                return;
            const headers = data[0];
            const headerMap = getHeaderMap(headers);
            const headerIndexes = resolveSourceHeaders(headerMap);
            if (!headerIndexes.contractEnd || !headerIndexes.dept) {
                needsInputRows.push([now, sheetName, '', '', '', '', '必要ヘッダが不足しています']);
                return;
            }
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
                else if (!deptMaster[dept]) {
                    needsInputRows.push([now, sheetName, vehicleId, dept, regCombined, vehicleType, '部署マスタ未登録']);
                }
                if (regCombined && !/\d/.test(String(regCombined))) {
                    if (!chassis) {
                        needsInputRows.push([now, sheetName, vehicleId, dept, regCombined, vehicleType, '登録番号が不完全（数字なし）かつ車台番号なし']);
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
        });
        writeSheetData(SHEET_NAMES.VEHICLE_VIEW, rows);
        writeSheetData(SHEET_NAMES.NEEDS_INPUT, needsInputRows);
        protectViewSheet(SHEET_NAMES.VEHICLE_VIEW);
        protectViewSheet(SHEET_NAMES.NEEDS_INPUT);
    }
    finally {
        lock.releaseLock();
    }
}
function createRequests() {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const ss = getSpreadsheet();
        ensureHeaders(ensureSheet(ss, SHEET_NAMES.REQUESTS), 1, getSchemaHeaders(SHEET_NAMES.REQUESTS));
        const settings = loadSettings();
        const deptMaster = loadDeptMaster();
        const vehicleSheet = ss.getSheetByName(SHEET_NAMES.VEHICLE_VIEW);
        if (!vehicleSheet)
            throw new Error('車両（統合ビュー）が存在しません');
        const vehicleData = vehicleSheet.getDataRange().getValues();
        if (vehicleData.length <= 1)
            return;
        const headerMap = getHeaderMap(vehicleData[0]);
        const idx = {
            vehicleId: headerMap['vehicleId'],
            dept: headerMap['管理部門'],
            contractEnd: headerMap['契約満了日'],
            requestId: headerMap['依頼ID'],
            regCombined: headerMap['登録番号_結合'],
            vehicleType: headerMap['車種'],
        };
        const tz = ss.getSpreadsheetTimeZone();
        const startDate = toDateOnly(new Date(), tz);
        const endDate = addMonthsClamped(startDate, settings.expiryMonths);
        const requestsByDept = {};
        for (let i = 1; i < vehicleData.length; i++) {
            const row = vehicleData[i];
            const dept = getCellValue(row, idx.dept);
            if (!dept)
                continue;
            const master = deptMaster[dept];
            if (!master || !master.active)
                continue;
            if (getCellValue(row, idx.requestId))
                continue;
            const contractEnd = parseDateValue(getCellRaw(row, idx.contractEnd));
            if (!contractEnd)
                continue;
            const contractDate = toDateOnly(contractEnd, tz);
            if (!isWithinRange(contractDate, startDate, endDate))
                continue;
            if (!requestsByDept[dept])
                requestsByDept[dept] = [];
            requestsByDept[dept].push({ rowIndex: i + 1, vehicleId: getCellValue(row, idx.vehicleId) });
        }
        const requestSheet = ss.getSheetByName(SHEET_NAMES.REQUESTS);
        if (!requestSheet)
            throw new Error('更新依頼シートが存在しません');
        const requestHeader = getHeaderMap(requestSheet.getRange(1, 1, 1, requestSheet.getLastColumn()).getValues()[0]);
        const newRequestRows = [];
        const now = new Date();
        const deadline = addDays(startDate, settings.deadlineAfterDays);
        Object.keys(requestsByDept).forEach((dept) => {
            const requestId = generateRequestId(now);
            const requestToken = generateToken();
            const row = new Array(requestSheet.getLastColumn()).fill('');
            const setCell = (headerName, value) => {
                const idx = requestHeader[headerName];
                if (idx)
                    row[idx - 1] = value;
            };
            setCell('requestId', requestId);
            setCell('管理部門', dept);
            setCell('対象開始日', startDate);
            setCell('対象終了日', endDate);
            setCell('締切日', deadline);
            setCell('ステータス', REQUEST_STATUS.CREATED);
            setCell('初回送信日時', '');
            setCell('最終リマインド日時', '');
            setCell('リマインド回数', 0);
            setCell('requestToken', requestToken);
            setCell('承認ステータス', APPROVAL_STATUS.NOT_SENT);
            setCell('承認依頼送信日時', '');
            setCell('承認者', '');
            setCell('承認日時', '');
            setCell('差戻しコメント', '');
            setCell('車両管理通知送信日時', '');
            setCell('承認フォームID', '');
            setCell('承認フォームURL', '');
            setCell('承認フォーム編集URL', '');
            setCell('承認フォームトリガーID', '');
            setCell('承認フォーム作成日時', '');
            newRequestRows.push(row);
            // 車両統合ビューへ依頼IDを反映
            requestsByDept[dept].forEach((item) => {
                const rowIndex = item.rowIndex;
                vehicleData[rowIndex - 1][idx.requestId - 1] = requestId;
            });
        });
        if (newRequestRows.length > 0) {
            const startRow = requestSheet.getLastRow() + 1;
            requestSheet.getRange(startRow, 1, newRequestRows.length, newRequestRows[0].length).setValues(newRequestRows);
            // 統合ビュー更新
            vehicleSheet.getRange(1, 1, vehicleData.length, vehicleData[0].length).setValues(vehicleData);
        }
    }
    finally {
        lock.releaseLock();
    }
}
function sendInitialEmails() {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const ss = getSpreadsheet();
        const settings = loadSettings();
        if (!settings.mailSendEnabled) {
            appendNotificationLog('初回', '', '', '', '通知_メール送信=FALSE のため送信をスキップ');
            return;
        }
        const deptMaster = loadDeptMaster();
        const requestSheet = ss.getSheetByName(SHEET_NAMES.REQUESTS);
        const vehicleSheet = ss.getSheetByName(SHEET_NAMES.VEHICLE_VIEW);
        if (!requestSheet || !vehicleSheet)
            throw new Error('必要シートが存在しません');
        const requestData = requestSheet.getDataRange().getValues();
        if (requestData.length <= 1)
            return;
        const reqHeader = getHeaderMap(requestData[0]);
        const vehicleData = vehicleSheet.getDataRange().getValues();
        const vehicleHeader = getHeaderMap(vehicleData[0]);
        const tz = ss.getSpreadsheetTimeZone();
        const now = new Date();
        for (let i = 1; i < requestData.length; i++) {
            const row = requestData[i];
            const status = getCellValue(row, reqHeader['ステータス']);
            if (status && status !== REQUEST_STATUS.CREATED)
                continue;
            const requestId = getCellValue(row, reqHeader['requestId']);
            const dept = getCellValue(row, reqHeader['管理部門']);
            const deptInfo = deptMaster[dept];
            if (!requestId)
                continue;
            if (!deptInfo || !deptInfo.active) {
                appendNotificationLog('初回', dept, '', requestId, '部署マスタ未登録/無効');
                continue;
            }
            if (!deptInfo.to) {
                appendNotificationLog('初回', dept, '', requestId, '通知先Toが未設定');
                continue;
            }
            const vehicles = vehicleData
                .slice(1)
                .filter((v) => getCellValue(v, vehicleHeader['依頼ID']) === requestId);
            if (vehicles.length === 0) {
                appendNotificationLog('初回', dept, deptInfo.to, requestId, '対象車両なし');
                continue;
            }
            const targetStart = parseDateValue(getCellRaw(row, reqHeader['対象開始日']));
            const targetEnd = parseDateValue(getCellRaw(row, reqHeader['対象終了日']));
            const deadline = parseDateValue(getCellRaw(row, reqHeader['締切日']));
            let formUrls = extractFormUrlsFromRequestRow(row, reqHeader);
            const formIds = extractFormIdsFromRequestRow(row, reqHeader);
            if (formIds.length > 0) {
                normalizeExistingFormsForRequest({
                    formIds,
                    vehicles,
                    vehicleHeader,
                    tz,
                    targetStart,
                    targetEnd,
                    deadline,
                    dept,
                    requestId,
                });
            }
            if (formUrls.length === 0) {
                const formResult = createRequestForms({
                    requestId,
                    dept,
                    vehicles,
                    vehicleHeader,
                    tz,
                    targetStart,
                    targetEnd,
                    deadline,
                });
                if (!formResult.ok) {
                    appendNotificationLog('初回', dept, deptInfo.to, requestId, `フォーム作成失敗: ${formResult.message}`);
                    continue;
                }
                applyFormResultToRequestRow(row, reqHeader, formResult, now);
                formUrls = formResult.formUrls;
            }
            if (formUrls.length === 0) {
                appendNotificationLog('初回', dept, deptInfo.to, requestId, 'フォームURLが未設定');
                continue;
            }
            const listText = vehicles
                .map((v) => formatVehicleLine(v, vehicleHeader, tz))
                .join('\n');
            const listHtml = vehicles
                .map((v) => `<li>${escapeHtml(formatVehicleLine(v, vehicleHeader, tz))}</li>`)
                .join('');
            const urlText = formatFormUrlsForText(formUrls);
            const urlHtml = formatFormUrlsForHtml(formUrls);
            const subject = applyTemplate(settings.subjectTemplate, {
                管理部門: dept,
                対象開始日: formatDateLabel(targetStart || new Date(), tz),
                対象終了日: formatDateLabel(targetEnd || new Date(), tz),
                締切日: formatDateLabel(deadline || new Date(), tz),
                URL: urlText,
            });
            const bodyText = applyTemplate(settings.bodyTemplate, {
                管理部門: dept,
                対象開始日: formatDateLabel(targetStart || new Date(), tz),
                対象終了日: formatDateLabel(targetEnd || new Date(), tz),
                締切日: formatDateLabel(deadline || new Date(), tz),
                URL: urlText,
                車両一覧: listText,
            });
            const htmlTemplate = applyTemplate(settings.bodyTemplate, {
                管理部門: dept,
                対象開始日: formatDateLabel(targetStart || new Date(), tz),
                対象終了日: formatDateLabel(targetEnd || new Date(), tz),
                締切日: formatDateLabel(deadline || new Date(), tz),
                URL: '[[FORM_URLS]]',
                車両一覧: '[[VEHICLE_LIST]]',
            });
            const htmlBody = escapeHtml(htmlTemplate)
                .replace(/\n/g, '<br>')
                .replace('[[FORM_URLS]]', urlHtml)
                .replace('[[VEHICLE_LIST]]', `<ul>${listHtml}</ul>`);
            try {
                MailApp.sendEmail({
                    to: deptInfo.to,
                    cc: deptInfo.cc,
                    subject,
                    name: settings.fromName,
                    htmlBody,
                    body: bodyText,
                });
                row[reqHeader['ステータス'] - 1] = REQUEST_STATUS.SENT;
                row[reqHeader['初回送信日時'] - 1] = now;
                appendNotificationLog('初回', dept, deptInfo.to, requestId, '成功');
            }
            catch (err) {
                appendNotificationLog('初回', dept, deptInfo.to, requestId, `失敗: ${err}`);
            }
        }
        requestSheet.getRange(1, 1, requestData.length, requestData[0].length).setValues(requestData);
    }
    finally {
        lock.releaseLock();
    }
}
function sendReminderEmails() {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const ss = getSpreadsheet();
        const settings = loadSettings();
        if (!settings.mailSendEnabled) {
            appendNotificationLog('リマインド', '', '', '', '通知_メール送信=FALSE のため送信をスキップ');
            return;
        }
        if (settings.reminderMaxCount <= 0)
            return;
        const deptMaster = loadDeptMaster();
        const requestSheet = ss.getSheetByName(SHEET_NAMES.REQUESTS);
        const vehicleSheet = ss.getSheetByName(SHEET_NAMES.VEHICLE_VIEW);
        if (!requestSheet || !vehicleSheet)
            throw new Error('必要シートが存在しません');
        const requestData = requestSheet.getDataRange().getValues();
        if (requestData.length <= 1)
            return;
        const reqHeader = getHeaderMap(requestData[0]);
        const vehicleData = vehicleSheet.getDataRange().getValues();
        if (vehicleData.length <= 1)
            return;
        const vehicleHeader = getHeaderMap(vehicleData[0]);
        const tz = ss.getSpreadsheetTimeZone();
        const today = toDateOnly(new Date(), tz);
        const now = new Date();
        const notifiedOverdue = loadNotifiedRequestIds('期限超過');
        for (let i = 1; i < requestData.length; i++) {
            const row = requestData[i];
            const requestId = getCellValue(row, reqHeader['requestId']);
            if (!requestId)
                continue;
            const dept = getCellValue(row, reqHeader['管理部門']);
            const deptInfo = deptMaster[dept];
            if (!deptInfo || !deptInfo.active)
                continue;
            const initialSentAt = parseDateValue(getCellRaw(row, reqHeader['初回送信日時']));
            if (!initialSentAt)
                continue;
            const vehicles = vehicleData
                .slice(1)
                .filter((v) => getCellValue(v, vehicleHeader['依頼ID']) === requestId);
            if (vehicles.length === 0)
                continue;
            const unansweredVehicles = vehicles.filter((v) => !getCellValue(v, vehicleHeader['更新方針']));
            if (unansweredVehicles.length === 0) {
                continue;
            }
            const answeredCount = vehicles.length - unansweredVehicles.length;
            row[reqHeader['ステータス'] - 1] = answeredCount > 0 ? REQUEST_STATUS.RESPONDING : REQUEST_STATUS.SENT;
            const deadline = parseDateValue(getCellRaw(row, reqHeader['締切日']));
            if (deadline) {
                const deadlineDate = toDateOnly(deadline, tz);
                if (today.getTime() > deadlineDate.getTime()) {
                    if (!notifiedOverdue[requestId]) {
                        notifyAdminOverdue({
                            requestId,
                            dept,
                            deadline,
                            unanswered: unansweredVehicles.length,
                            total: vehicles.length,
                            settings,
                            tz,
                        });
                        notifiedOverdue[requestId] = true;
                    }
                    continue;
                }
            }
            const status = getCellValue(row, reqHeader['ステータス']);
            if (status !== REQUEST_STATUS.SENT && status !== REQUEST_STATUS.RESPONDING)
                continue;
            const reminderCount = toNumber(getCellRaw(row, reqHeader['リマインド回数']), 0);
            if (reminderCount >= settings.reminderMaxCount)
                continue;
            const lastReminderAt = parseDateValue(getCellRaw(row, reqHeader['最終リマインド日時']));
            if (lastReminderAt && toDateOnly(lastReminderAt, tz).getTime() === today.getTime())
                continue;
            const eligibleFrom = reminderCount === 0 || !lastReminderAt
                ? addDays(toDateOnly(initialSentAt, tz), settings.reminderStartAfterDays)
                : addDays(toDateOnly(lastReminderAt, tz), settings.reminderIntervalDays);
            if (today.getTime() < eligibleFrom.getTime())
                continue;
            const targetStart = parseDateValue(getCellRaw(row, reqHeader['対象開始日']));
            const targetEnd = parseDateValue(getCellRaw(row, reqHeader['対象終了日']));
            let formUrls = extractFormUrlsFromRequestRow(row, reqHeader);
            const formIds = extractFormIdsFromRequestRow(row, reqHeader);
            if (formIds.length > 0) {
                normalizeExistingFormsForRequest({
                    formIds,
                    vehicles,
                    vehicleHeader,
                    tz,
                    targetStart,
                    targetEnd,
                    deadline,
                    dept,
                    requestId,
                });
            }
            if (formUrls.length === 0) {
                const formResult = createRequestForms({
                    requestId,
                    dept,
                    vehicles,
                    vehicleHeader,
                    tz,
                    targetStart,
                    targetEnd,
                    deadline,
                });
                if (!formResult.ok) {
                    appendNotificationLog('リマインド', dept, deptInfo.to, requestId, `フォーム作成失敗: ${formResult.message}`);
                    continue;
                }
                applyFormResultToRequestRow(row, reqHeader, formResult, now);
                formUrls = formResult.formUrls;
            }
            if (formUrls.length === 0) {
                appendNotificationLog('リマインド', dept, deptInfo.to, requestId, 'フォームURLが未設定');
                continue;
            }
            if (!deptInfo.to) {
                appendNotificationLog('リマインド', dept, '', requestId, '通知先Toが未設定');
                continue;
            }
            const listText = unansweredVehicles
                .map((v) => formatVehicleLine(v, vehicleHeader, tz))
                .join('\n');
            const listHtml = unansweredVehicles
                .map((v) => `<li>${escapeHtml(formatVehicleLine(v, vehicleHeader, tz))}</li>`)
                .join('');
            const urlText = formatFormUrlsForText(formUrls);
            const urlHtml = formatFormUrlsForHtml(formUrls);
            const subjectBase = applyTemplate(settings.subjectTemplate, {
                管理部門: dept,
                対象開始日: formatDateLabel(targetStart || new Date(), tz),
                対象終了日: formatDateLabel(targetEnd || new Date(), tz),
                締切日: formatDateLabel(deadline || new Date(), tz),
                URL: urlText,
            });
            const subject = `【リマインド】${subjectBase}`;
            const bodyText = applyTemplate(settings.bodyTemplate, {
                管理部門: dept,
                対象開始日: formatDateLabel(targetStart || new Date(), tz),
                対象終了日: formatDateLabel(targetEnd || new Date(), tz),
                締切日: formatDateLabel(deadline || new Date(), tz),
                URL: urlText,
                車両一覧: listText,
            });
            const htmlTemplate = applyTemplate(settings.bodyTemplate, {
                管理部門: dept,
                対象開始日: formatDateLabel(targetStart || new Date(), tz),
                対象終了日: formatDateLabel(targetEnd || new Date(), tz),
                締切日: formatDateLabel(deadline || new Date(), tz),
                URL: '[[FORM_URLS]]',
                車両一覧: '[[VEHICLE_LIST]]',
            });
            const htmlBody = escapeHtml(htmlTemplate)
                .replace(/\n/g, '<br>')
                .replace('[[FORM_URLS]]', urlHtml)
                .replace('[[VEHICLE_LIST]]', `<ul>${listHtml}</ul>`);
            try {
                MailApp.sendEmail({
                    to: deptInfo.to,
                    cc: deptInfo.cc,
                    subject,
                    name: settings.fromName,
                    htmlBody,
                    body: bodyText,
                });
                row[reqHeader['最終リマインド日時'] - 1] = now;
                row[reqHeader['リマインド回数'] - 1] = reminderCount + 1;
                appendNotificationLog('リマインド', dept, deptInfo.to, requestId, '成功');
            }
            catch (err) {
                appendNotificationLog('リマインド', dept, deptInfo.to, requestId, `失敗: ${err}`);
            }
        }
        requestSheet.getRange(1, 1, requestData.length, requestData[0].length).setValues(requestData);
    }
    finally {
        lock.releaseLock();
    }
}
function loadNotifiedRequestIds(type) {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.NOTIFY_LOG);
    if (!sheet || sheet.getLastRow() <= 1)
        return {};
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1)
        return {};
    const headerMap = getHeaderMap(data[0]);
    const result = {};
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (getCellValue(row, headerMap['種別']) !== type)
            continue;
        const requestId = getCellValue(row, headerMap['requestId']);
        if (requestId)
            result[requestId] = true;
    }
    return result;
}
function notifyAdminOverdue(params) {
    const to = params.settings.adminTo;
    if (!to) {
        appendNotificationLog('期限超過', params.dept, '', params.requestId, '管理者_通知先Toが未設定');
        return;
    }
    const deadlineLabel = formatDateLabel(params.deadline, params.tz);
    const subject = `【車両更新】期限超過: ${params.dept}（締切 ${deadlineLabel}）`;
    const body = [
        '更新依頼が締切日を超過しました（フォームは閉じません）。',
        `管理部門: ${params.dept}`,
        `requestId: ${params.requestId}`,
        `締切日: ${deadlineLabel}`,
        `未回答: ${params.unanswered} / 総件数: ${params.total}`,
    ].join('\n');
    try {
        MailApp.sendEmail({
            to,
            cc: params.settings.adminCc,
            subject,
            name: params.settings.fromName,
            body,
        });
        appendNotificationLog('期限超過', params.dept, to, params.requestId, '成功');
    }
    catch (err) {
        appendNotificationLog('期限超過', params.dept, to, params.requestId, `失敗: ${err}`);
    }
}
function doGet(e) {
    const message = 'このWeb回答ページは廃止されました。通知メール内のGoogleフォームからご回答ください。';
    return HtmlService.createHtmlOutput(`<p>${escapeHtml(message)}</p>`).setTitle('車両更新回答');
}
function doPost(e) {
    const message = 'このWeb回答ページは廃止されました。通知メール内のGoogleフォームからご回答ください。';
    return HtmlService.createHtmlOutput(`<p>${escapeHtml(message)}</p>`).setTitle('車両更新回答');
}
function onRequestFormSubmit(e) {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        if (!e || !e.response) {
            Logger.log('onRequestFormSubmit: response がありません');
            return;
        }
        const formId = getFormIdFromEvent(e);
        if (!formId) {
            Logger.log('onRequestFormSubmit: formId を取得できません');
            return;
        }
        const requestInfo = findRequestByFormId(formId);
        if (!requestInfo) {
            Logger.log(`onRequestFormSubmit: formId に紐づく依頼が見つかりません (${formId})`);
            return;
        }
        const parsed = extractAnswersFromFormResponse(formId, e.response);
        const vehicleIds = Object.keys(parsed.answersByVehicleId);
        if (vehicleIds.length === 0) {
            Logger.log(`onRequestFormSubmit: 回答が空です (${requestInfo.requestId})`);
            return;
        }
        const now = new Date();
        const answerInputs = vehicleIds.map((vehicleId) => ({
            requestId: requestInfo.requestId,
            vehicleId,
            answer: normalizeAnswerLabel(parsed.answersByVehicleId[vehicleId]),
            comment: '',
            responder: '',
            answeredAt: now,
        }));
        upsertAnswers(answerInputs);
        applyAnswers();
        const status = updateRequestStatus(requestInfo.requestId);
        buildSummarySheet();
        if (status === REQUEST_STATUS.COMPLETED) {
            closeRequestForms(requestInfo.requestId);
        }
    }
    finally {
        lock.releaseLock();
    }
}
function onApprovalFormSubmit(e) {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        if (!e || !e.response) {
            Logger.log('onApprovalFormSubmit: response がありません');
            return;
        }
        const formId = getFormIdFromEvent(e);
        if (!formId) {
            Logger.log('onApprovalFormSubmit: formId を取得できません');
            return;
        }
        const requestInfo = findRequestByApprovalFormId(formId);
        if (!requestInfo || !requestInfo.requestId) {
            Logger.log(`onApprovalFormSubmit: formId に紐づく依頼が見つかりません (${formId})`);
            return;
        }
        const parsed = extractApprovalDecisionFromFormResponse(e.response);
        if (!parsed.decision) {
            Logger.log(`onApprovalFormSubmit: 承認判断が空です (${requestInfo.requestId})`);
            return;
        }
        buildApprovalQueueSheetInternal();
        writeApprovalDecisionToQueue({
            requestId: requestInfo.requestId,
            decision: parsed.decision,
            comment: parsed.comment,
        });
    }
    finally {
        lock.releaseLock();
    }
    // 実際の状態更新・通知送信は既存ロジックで一元処理する
    applyApprovalDecisions();
}
function extractApprovalDecisionFromFormResponse(response) {
    const result = {
        decision: '',
        comment: '',
    };
    response.getItemResponses().forEach((itemResponse) => {
        const item = itemResponse.getItem();
        const title = item.getTitle();
        if (title === APPROVAL_FORM_TITLES.DECISION) {
            result.decision = String(itemResponse.getResponse() || '').trim();
            return;
        }
        if (title === APPROVAL_FORM_TITLES.COMMENT) {
            result.comment = String(itemResponse.getResponse() || '').trim();
        }
    });
    return result;
}
function writeApprovalDecisionToQueue(input) {
    const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.APPROVAL_QUEUE);
    if (!sheet)
        return;
    ensureHeaders(sheet, 1, getSchemaHeaders(SHEET_NAMES.APPROVAL_QUEUE));
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1)
        return;
    const headerMap = getHeaderMap(data[0]);
    const requestIdIndex = headerMap['requestId'];
    const decisionIndex = headerMap['承認入力'];
    const commentIndex = headerMap['差戻しコメント'];
    if (!requestIdIndex || !decisionIndex || !commentIndex)
        return;
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (getCellValue(row, requestIdIndex) !== input.requestId)
            continue;
        row[decisionIndex - 1] = input.decision;
        row[commentIndex - 1] = input.comment || '';
        sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
        return;
    }
}
function applyAnswers() {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const ss = getSpreadsheet();
        const answerSheet = ss.getSheetByName(SHEET_NAMES.ANSWERS);
        const vehicleSheet = ss.getSheetByName(SHEET_NAMES.VEHICLE_VIEW);
        if (!answerSheet || !vehicleSheet)
            return;
        const answerData = answerSheet.getDataRange().getValues();
        if (answerData.length <= 1)
            return;
        const answerHeader = getHeaderMap(answerData[0]);
        const answerMap = {};
        for (let i = 1; i < answerData.length; i++) {
            const row = answerData[i];
            const vehicleId = getCellValue(row, answerHeader['vehicleId']);
            if (!vehicleId)
                continue;
            const answeredAt = parseDateValue(getCellRaw(row, answerHeader['回答日時']));
            if (!answerMap[vehicleId] || (answeredAt && answeredAt > answerMap[vehicleId].answeredAt)) {
                answerMap[vehicleId] = {
                    vehicleId,
                    requestId: getCellValue(row, answerHeader['requestId']),
                    answer: normalizeAnswerLabel(getCellValue(row, answerHeader['回答'])),
                    comment: getCellValue(row, answerHeader['コメント']),
                    answeredAt: answeredAt || new Date(),
                };
            }
        }
        // 統合ビュー更新
        const vehicleData = vehicleSheet.getDataRange().getValues();
        const vehicleHeader = getHeaderMap(vehicleData[0]);
        for (let i = 1; i < vehicleData.length; i++) {
            const row = vehicleData[i];
            const vehicleId = getCellValue(row, vehicleHeader['vehicleId']);
            const answer = answerMap[vehicleId];
            if (!answer)
                continue;
            row[vehicleHeader['更新方針'] - 1] = answer.answer;
            if (vehicleHeader['一次回答']) {
                row[vehicleHeader['一次回答'] - 1] = answer.answer;
            }
            row[vehicleHeader['備考'] - 1] = answer.comment;
            row[vehicleHeader['回答日'] - 1] = answer.answeredAt;
            if (vehicleHeader['依頼ID'] && !row[vehicleHeader['依頼ID'] - 1]) {
                row[vehicleHeader['依頼ID'] - 1] = answer.requestId;
            }
        }
        vehicleSheet.getRange(1, 1, vehicleData.length, vehicleData[0].length).setValues(vehicleData);
        // 元台帳へ反映
        SOURCE_SHEETS.forEach((sheetName) => {
            const sheet = ss.getSheetByName(sheetName);
            if (!sheet)
                return;
            ensureAppendColumns(sheet, ['更新方針', '依頼ID', '回答日', '備考']);
            const data = sheet.getDataRange().getValues();
            if (data.length <= 1)
                return;
            const headerMap = getHeaderMap(data[0]);
            const headerIndexes = resolveSourceHeaders(headerMap);
            const updateIndexes = {
                policy: headerMap['更新方針'],
                requestId: headerMap['依頼ID'],
                answeredAt: headerMap['回答日'],
                note: headerMap['備考'],
            };
            for (let i = 1; i < data.length; i++) {
                const row = data[i];
                if (row.every((cell) => cell === '' || cell === null))
                    continue;
                const regCombined = getSourceRegistrationCombined(row, headerIndexes);
                const chassis = getCellValue(row, headerIndexes.chassis);
                const vehicleId = buildVehicleId(sheetName, regCombined, chassis, i + 1);
                const answer = answerMap[vehicleId];
                if (!answer)
                    continue;
                row[updateIndexes.policy - 1] = answer.answer;
                row[updateIndexes.requestId - 1] = answer.requestId;
                row[updateIndexes.answeredAt - 1] = answer.answeredAt;
                row[updateIndexes.note - 1] = answer.comment;
            }
            sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
        });
    }
    finally {
        lock.releaseLock();
    }
}
function buildSummarySheet() {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const ss = getSpreadsheet();
        const settings = loadSettings();
        if (!settings.summarySheetEnabled)
            return;
        const requestSheet = ss.getSheetByName(SHEET_NAMES.REQUESTS);
        const vehicleSheet = ss.getSheetByName(SHEET_NAMES.VEHICLE_VIEW);
        if (!requestSheet || !vehicleSheet)
            return;
        const requestData = requestSheet.getDataRange().getValues();
        if (requestData.length <= 1)
            return;
        const requestHeader = getHeaderMap(requestData[0]);
        const vehicleData = vehicleSheet.getDataRange().getValues();
        const vehicleHeader = getHeaderMap(vehicleData[0]);
        const rows = [];
        const now = new Date();
        for (let i = 1; i < requestData.length; i++) {
            const row = requestData[i];
            const requestId = getCellValue(row, requestHeader['requestId']);
            if (!requestId)
                continue;
            const dept = getCellValue(row, requestHeader['管理部門']);
            const start = parseDateValue(getCellRaw(row, requestHeader['対象開始日']));
            const end = parseDateValue(getCellRaw(row, requestHeader['対象終了日']));
            const vehicles = vehicleData
                .slice(1)
                .filter((v) => getCellValue(v, vehicleHeader['依頼ID']) === requestId);
            const counts = {
                [ANSWER_LABELS.RENEW]: 0,
                [ANSWER_LABELS.CANCELLATION_REPLACE]: 0,
                [ANSWER_LABELS.CANCELLATION_END]: 0,
                未回答: 0,
            };
            vehicles.forEach((v) => {
                const policy = normalizeAnswerLabel(getCellValue(v, vehicleHeader['更新方針']));
                if (policy && counts[policy] !== undefined) {
                    counts[policy] += 1;
                }
                else {
                    counts['未回答'] += 1;
                }
            });
            rows.push([
                requestId,
                dept,
                `${formatDateLabel(start || now, ss.getSpreadsheetTimeZone())}〜${formatDateLabel(end || now, ss.getSpreadsheetTimeZone())}`,
                vehicles.length,
                counts[ANSWER_LABELS.RENEW],
                counts[ANSWER_LABELS.CANCELLATION_REPLACE],
                counts[ANSWER_LABELS.CANCELLATION_END],
                counts['未回答'],
                now,
            ]);
        }
        writeSheetData(SHEET_NAMES.SUMMARY, rows);
    }
    finally {
        lock.releaseLock();
    }
}
function sendSummaryEmail() {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const ss = getSpreadsheet();
        const settings = loadSettings();
        if (!settings.summaryEmailEnabled)
            return;
        if (!settings.adminTo)
            return;
        buildSummarySheet();
        const summarySheet = ss.getSheetByName(SHEET_NAMES.SUMMARY);
        if (!summarySheet)
            return;
        const data = summarySheet.getDataRange().getValues();
        if (data.length <= 1)
            return;
        const headerMap = getHeaderMap(data[0]);
        const lines = [];
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const dept = getCellValue(row, headerMap['管理部門']);
            const range = getCellValue(row, headerMap['対象期間']);
            const total = getCellValue(row, headerMap['総件数']);
            const renew = getCellValue(row, headerMap[ANSWER_LABELS.RENEW]);
            const replace = getCellValue(row, headerMap[ANSWER_LABELS.CANCELLATION_REPLACE]);
            const end = getCellValue(row, headerMap[ANSWER_LABELS.CANCELLATION_END]);
            const unanswered = getCellValue(row, headerMap['未回答']);
            lines.push(`${dept} (${range}) - 合計:${total} 更新:${renew} 解約(入替):${replace} 解約(満了):${end} 未回答:${unanswered}`);
        }
        const body = lines.join('\n');
        MailApp.sendEmail({
            to: settings.adminTo,
            cc: settings.adminCc,
            subject: '【車両更新回答集計】サマリ',
            name: settings.fromName,
            body,
        });
    }
    finally {
        lock.releaseLock();
    }
}
function buildApprovalQueueSheet() {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        buildApprovalQueueSheetInternal();
    }
    finally {
        lock.releaseLock();
    }
}
function buildApprovalQueueSheetInternal() {
    const ss = getSpreadsheet();
    const requestSheet = ss.getSheetByName(SHEET_NAMES.REQUESTS);
    const vehicleSheet = ss.getSheetByName(SHEET_NAMES.VEHICLE_VIEW);
    const approvalSheet = ensureSheet(ss, SHEET_NAMES.APPROVAL_QUEUE);
    if (!requestSheet || !vehicleSheet)
        return;
    ensureHeaders(requestSheet, 1, getSchemaHeaders(SHEET_NAMES.REQUESTS));
    ensureHeaders(vehicleSheet, 1, getSchemaHeaders(SHEET_NAMES.VEHICLE_VIEW));
    ensureHeaders(approvalSheet, 1, getSchemaHeaders(SHEET_NAMES.APPROVAL_QUEUE));
    const approvalExistingData = approvalSheet.getDataRange().getValues();
    const approvalExistingHeader = approvalExistingData.length > 0 ? getHeaderMap(approvalExistingData[0]) : {};
    const existingInputByRequestId = {};
    if (approvalExistingData.length > 1 && approvalExistingHeader['requestId']) {
        for (let i = 1; i < approvalExistingData.length; i++) {
            const row = approvalExistingData[i];
            const requestId = getCellValue(row, approvalExistingHeader['requestId']);
            if (!requestId)
                continue;
            existingInputByRequestId[requestId] = {
                input: getCellValue(row, approvalExistingHeader['承認入力']),
                comment: getCellValue(row, approvalExistingHeader['差戻しコメント']),
            };
        }
    }
    const requestData = requestSheet.getDataRange().getValues();
    if (requestData.length <= 1) {
        writeSheetData(SHEET_NAMES.APPROVAL_QUEUE, []);
        return;
    }
    const requestHeader = getHeaderMap(requestData[0]);
    const vehicleData = vehicleSheet.getDataRange().getValues();
    const vehicleHeader = vehicleData.length > 0 ? getHeaderMap(vehicleData[0]) : {};
    const tz = ss.getSpreadsheetTimeZone();
    const vehiclesByRequestId = {};
    if (vehicleData.length > 1 && vehicleHeader['依頼ID']) {
        for (let i = 1; i < vehicleData.length; i++) {
            const row = vehicleData[i];
            const requestId = getCellValue(row, vehicleHeader['依頼ID']);
            if (!requestId)
                continue;
            if (!vehiclesByRequestId[requestId])
                vehiclesByRequestId[requestId] = [];
            vehiclesByRequestId[requestId].push(row);
        }
    }
    let requestSheetChanged = false;
    const rows = [];
    for (let i = 1; i < requestData.length; i++) {
        const row = requestData[i];
        const requestId = getCellValue(row, requestHeader['requestId']);
        if (!requestId)
            continue;
        const dept = getCellValue(row, requestHeader['管理部門']);
        const status = getCellValue(row, requestHeader['ステータス']);
        const approvalStatusIndex = requestHeader['承認ステータス'];
        const approvalStatus = approvalStatusIndex ? getCellValue(row, approvalStatusIndex) : '';
        const shouldShow = status === REQUEST_STATUS.COMPLETED ||
            approvalStatus === APPROVAL_STATUS.PENDING ||
            approvalStatus === APPROVAL_STATUS.RETURNED ||
            approvalStatus === APPROVAL_STATUS.APPROVED;
        if (!shouldShow)
            continue;
        let promotedToPending = false;
        if (status === REQUEST_STATUS.COMPLETED &&
            approvalStatusIndex &&
            (!approvalStatus || approvalStatus === APPROVAL_STATUS.NOT_SENT || approvalStatus === APPROVAL_STATUS.RETURNED)) {
            row[approvalStatusIndex - 1] = APPROVAL_STATUS.PENDING;
            if (requestHeader['承認依頼送信日時']) {
                row[requestHeader['承認依頼送信日時'] - 1] = '';
            }
            requestSheetChanged = true;
            promotedToPending = true;
        }
        const targetStart = parseDateValue(getCellRaw(row, requestHeader['対象開始日']));
        const targetEnd = parseDateValue(getCellRaw(row, requestHeader['対象終了日']));
        const deadline = parseDateValue(getCellRaw(row, requestHeader['締切日']));
        const approvalRequestedAt = parseDateValue(getCellRaw(row, requestHeader['承認依頼送信日時']));
        const approvedAt = parseDateValue(getCellRaw(row, requestHeader['承認日時']));
        const approvedBy = getCellValue(row, requestHeader['承認者']);
        const returnedComment = getCellValue(row, requestHeader['差戻しコメント']);
        const vehicleManagerNotifiedAt = parseDateValue(getCellRaw(row, requestHeader['車両管理通知送信日時']));
        const vehicles = vehiclesByRequestId[requestId] || [];
        const counts = {
            answered: 0,
            total: vehicles.length,
            [ANSWER_LABELS.RENEW]: 0,
            [ANSWER_LABELS.CANCELLATION_REPLACE]: 0,
            [ANSWER_LABELS.CANCELLATION_END]: 0,
            未回答: 0,
        };
        vehicles.forEach((v) => {
            const policy = normalizeAnswerLabel(getCellValue(v, vehicleHeader['更新方針']));
            if (policy)
                counts.answered += 1;
            if (policy && counts[policy] !== undefined) {
                counts[policy] += 1;
            }
            else {
                counts['未回答'] += 1;
            }
        });
        const summaryText = [
            `合計:${counts.total}`,
            `更新:${counts[ANSWER_LABELS.RENEW]}`,
            `解約(入替):${counts[ANSWER_LABELS.CANCELLATION_REPLACE]}`,
            `解約(満了):${counts[ANSWER_LABELS.CANCELLATION_END]}`,
            `未回答:${counts['未回答']}`,
        ].join(' ');
        const listLines = vehicles
            .map((v, index) => formatVehicleLineForApproval(v, vehicleHeader, tz, index))
            .join('\n');
        const existing = existingInputByRequestId[requestId];
        const approvalInput = promotedToPending ? '' : existing ? existing.input : '';
        const commentInput = promotedToPending ? '' : existing ? existing.comment : returnedComment || '';
        rows.push([
            requestId,
            dept,
            targetStart || '',
            targetEnd || '',
            deadline || '',
            status,
            approvalStatusIndex ? getCellValue(row, approvalStatusIndex) : approvalStatus,
            counts.total,
            counts.answered,
            summaryText,
            listLines,
            approvalInput,
            commentInput,
            approvedBy || '',
            approvedAt || '',
            approvalRequestedAt || '',
            vehicleManagerNotifiedAt || '',
        ]);
    }
    if (requestSheetChanged) {
        requestSheet.getRange(1, 1, requestData.length, requestData[0].length).setValues(requestData);
    }
    writeSheetData(SHEET_NAMES.APPROVAL_QUEUE, rows);
}
function sendApprovalRequestEmails() {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const ss = getSpreadsheet();
        const settings = loadSettings();
        if (!settings.approvalFlowEnabled)
            return;
        if (!settings.mailSendEnabled) {
            appendNotificationLog('承認依頼', '', '', '', '通知_メール送信=FALSE のため送信をスキップ');
            return;
        }
        if (!settings.approverTo) {
            appendNotificationLog('承認依頼', '', '', '', '承認者_通知先Toが未設定');
            return;
        }
        buildApprovalQueueSheetInternal();
        const requestSheet = ss.getSheetByName(SHEET_NAMES.REQUESTS);
        const approvalSheet = ss.getSheetByName(SHEET_NAMES.APPROVAL_QUEUE);
        if (!requestSheet || !approvalSheet)
            return;
        const requestData = requestSheet.getDataRange().getValues();
        if (requestData.length <= 1)
            return;
        const requestHeader = getHeaderMap(requestData[0]);
        const approvalData = approvalSheet.getDataRange().getValues();
        if (approvalData.length <= 1)
            return;
        const approvalHeader = getHeaderMap(approvalData[0]);
        const requestRowIndexById = {};
        for (let i = 1; i < requestData.length; i++) {
            const requestId = getCellValue(requestData[i], requestHeader['requestId']);
            if (requestId)
                requestRowIndexById[requestId] = i;
        }
        const preparedRequestIds = [];
        const lines = [];
        const formUrlLines = [];
        const now = new Date();
        for (let i = 1; i < approvalData.length; i++) {
            const row = approvalData[i];
            const requestId = getCellValue(row, approvalHeader['requestId']);
            if (!requestId)
                continue;
            const approvalStatus = getCellValue(row, approvalHeader['承認ステータス']);
            const requestedAt = parseDateValue(getCellRaw(row, approvalHeader['承認依頼送信日時']));
            if (approvalStatus !== APPROVAL_STATUS.PENDING)
                continue;
            if (requestedAt)
                continue;
            const requestRowIndex = requestRowIndexById[requestId];
            if (requestRowIndex === undefined)
                continue;
            const requestRow = requestData[requestRowIndex];
            const dept = getCellValue(row, approvalHeader['管理部門']);
            const rangeStart = parseDateValue(getCellRaw(row, approvalHeader['対象開始日']));
            const rangeEnd = parseDateValue(getCellRaw(row, approvalHeader['対象終了日']));
            const summaryText = getCellValue(row, approvalHeader['一次結果サマリ']);
            const existingApprovalFormId = getCellValue(requestRow, requestHeader['承認フォームID']);
            const vehiclesText = getCellValue(row, approvalHeader['対象車両一覧']);
            const formResult = createOrUpdateApprovalForm({
                requestId,
                dept,
                targetStart: rangeStart,
                targetEnd: rangeEnd,
                summaryText,
                vehiclesText,
                existingFormId: existingApprovalFormId,
                tz: ss.getSpreadsheetTimeZone(),
            });
            if (!formResult.ok) {
                appendNotificationLog('承認依頼', dept, '', requestId, `承認フォーム作成失敗: ${formResult.message}`);
                continue;
            }
            requestRow[requestHeader['承認フォームID'] - 1] = formResult.formId;
            requestRow[requestHeader['承認フォームURL'] - 1] = formResult.formUrl;
            requestRow[requestHeader['承認フォーム編集URL'] - 1] = formResult.formEditUrl;
            requestRow[requestHeader['承認フォームトリガーID'] - 1] = formResult.formTriggerId;
            if (!getCellRaw(requestRow, requestHeader['承認フォーム作成日時'])) {
                requestRow[requestHeader['承認フォーム作成日時'] - 1] = now;
            }
            lines.push(`${dept || ''} ${requestId}（${formatDateLabel(rangeStart || new Date(), ss.getSpreadsheetTimeZone())}〜${formatDateLabel(rangeEnd || new Date(), ss.getSpreadsheetTimeZone())}） ${summaryText}`);
            formUrlLines.push(`- ${requestId}: ${formResult.formUrl}`);
            preparedRequestIds.push(requestId);
        }
        if (preparedRequestIds.length === 0)
            return;
        const subject = `【車両更新】専務承認依頼（${preparedRequestIds.length}件）`;
        const body = [
            '一次確認（フォーム回答）の結果が揃ったため、承認/差戻しの入力をお願いします。',
            '',
            '承認フォームURL:',
            ...formUrlLines,
            '',
            '対象依頼:',
            ...lines.map((l) => `- ${l}`),
            '',
            `入力方法: 各フォームで「${APPROVAL_INPUT.APPROVE}」または「${APPROVAL_INPUT.RETURN}」を選択し、差戻し時はコメントを入力してください。`,
        ].join('\n');
        try {
            MailApp.sendEmail({
                to: settings.approverTo,
                cc: settings.approverCc,
                subject,
                name: settings.fromName,
                body,
            });
            appendNotificationLog('承認依頼', '', settings.approverTo, preparedRequestIds.join(','), `成功(${preparedRequestIds.length}件)`);
        }
        catch (err) {
            appendNotificationLog('承認依頼', '', settings.approverTo, preparedRequestIds.join(','), `失敗: ${err}`);
            return;
        }
        const requestIdIndex = requestHeader['requestId'];
        const approvalRequestedAtIndex = requestHeader['承認依頼送信日時'];
        if (requestIdIndex && approvalRequestedAtIndex) {
            for (let i = 1; i < requestData.length; i++) {
                const row = requestData[i];
                const requestId = getCellValue(row, requestIdIndex);
                if (!requestId)
                    continue;
                if (preparedRequestIds.indexOf(requestId) === -1)
                    continue;
                row[approvalRequestedAtIndex - 1] = now;
            }
            requestSheet.getRange(1, 1, requestData.length, requestData[0].length).setValues(requestData);
        }
        buildApprovalQueueSheetInternal();
    }
    finally {
        lock.releaseLock();
    }
}
function applyApprovalDecisions() {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const ss = getSpreadsheet();
        const settings = loadSettings();
        if (!settings.approvalFlowEnabled) {
            uiShowModalSafe('承認結果反映（通知送信）', '承認フロー_有効 が FALSE のため処理をスキップしました。');
            return;
        }
        const approvalSheet = ss.getSheetByName(SHEET_NAMES.APPROVAL_QUEUE);
        const requestSheet = ss.getSheetByName(SHEET_NAMES.REQUESTS);
        const vehicleSheet = ss.getSheetByName(SHEET_NAMES.VEHICLE_VIEW);
        if (!approvalSheet || !requestSheet || !vehicleSheet)
            return;
        ensureHeaders(approvalSheet, 1, getSchemaHeaders(SHEET_NAMES.APPROVAL_QUEUE));
        ensureHeaders(requestSheet, 1, getSchemaHeaders(SHEET_NAMES.REQUESTS));
        ensureHeaders(vehicleSheet, 1, getSchemaHeaders(SHEET_NAMES.VEHICLE_VIEW));
        const approvalData = approvalSheet.getDataRange().getValues();
        if (approvalData.length <= 1)
            return;
        const approvalHeader = getHeaderMap(approvalData[0]);
        const requestData = requestSheet.getDataRange().getValues();
        if (requestData.length <= 1)
            return;
        const requestHeader = getHeaderMap(requestData[0]);
        const vehicleData = vehicleSheet.getDataRange().getValues();
        const vehicleHeader = vehicleData.length > 0 ? getHeaderMap(vehicleData[0]) : {};
        const tz = ss.getSpreadsheetTimeZone();
        const deptMaster = loadDeptMaster();
        const requestRowIndexById = {};
        for (let i = 1; i < requestData.length; i++) {
            const id = getCellValue(requestData[i], requestHeader['requestId']);
            if (id)
                requestRowIndexById[id] = i;
        }
        const now = new Date();
        const approverEmail = safeGetUserEmail();
        const reportLines = [];
        const detailLines = [];
        let inputRows = 0;
        let invalidInputRows = 0;
        let skippedSameRows = 0;
        let missingRequestRows = 0;
        let approvedRows = 0;
        let returnedRows = 0;
        if (!settings.mailSendEnabled) {
            reportLines.push('注意: 通知_メール送信=FALSE のため、メール送信はスキップします（状態更新のみ）。');
        }
        let processed = 0;
        for (let i = 1; i < approvalData.length; i++) {
            const row = approvalData[i];
            const requestId = getCellValue(row, approvalHeader['requestId']);
            if (!requestId)
                continue;
            const input = getCellValue(row, approvalHeader['承認入力']);
            if (!input)
                continue;
            inputRows += 1;
            if (input !== APPROVAL_INPUT.APPROVE && input !== APPROVAL_INPUT.RETURN)
                continue;
            const requestRowIndex = requestRowIndexById[requestId];
            if (requestRowIndex === undefined) {
                missingRequestRows += 1;
                detailLines.push(`requestId=${requestId}: 更新依頼の行が見つからないためスキップ`);
                continue;
            }
            const requestRow = requestData[requestRowIndex];
            const dept = getCellValue(requestRow, requestHeader['管理部門']);
            const currentApproval = getCellValue(requestRow, requestHeader['承認ステータス']);
            if (input === APPROVAL_INPUT.APPROVE && currentApproval === APPROVAL_STATUS.APPROVED) {
                skippedSameRows += 1;
                continue;
            }
            if (input === APPROVAL_INPUT.RETURN && currentApproval === APPROVAL_STATUS.RETURNED) {
                skippedSameRows += 1;
                continue;
            }
            if (input === APPROVAL_INPUT.APPROVE) {
                requestRow[requestHeader['承認ステータス'] - 1] = APPROVAL_STATUS.APPROVED;
                requestRow[requestHeader['承認者'] - 1] = approverEmail;
                requestRow[requestHeader['承認日時'] - 1] = now;
                requestRow[requestHeader['差戻しコメント'] - 1] = '';
                // 車両（統合ビュー）へ最終決定を反映（空欄のみ）
                if (vehicleData.length > 1 && vehicleHeader['依頼ID'] && vehicleHeader['更新方針'] && vehicleHeader['最終決定']) {
                    for (let v = 1; v < vehicleData.length; v++) {
                        const vr = vehicleData[v];
                        if (getCellValue(vr, vehicleHeader['依頼ID']) !== requestId)
                            continue;
                        const primary = getCellValue(vr, vehicleHeader['更新方針']);
                        if (vehicleHeader['一次回答']) {
                            vr[vehicleHeader['一次回答'] - 1] = primary;
                        }
                        const finalDecision = getCellValue(vr, vehicleHeader['最終決定']);
                        if (!finalDecision && primary) {
                            vr[vehicleHeader['最終決定'] - 1] = primary;
                        }
                    }
                }
                // 車両管理担当へ通知
                const to = settings.vehicleManagerTo;
                if (!settings.mailSendEnabled) {
                    detailLines.push(`承認: requestId=${requestId} メール送信スキップ（通知_メール送信=FALSE）`);
                }
                else if (!to) {
                    appendNotificationLog('承認結果通知', dept, '', requestId, '車両管理担当_通知先Toが未設定');
                    detailLines.push(`承認: requestId=${requestId} 車両管理担当_通知先Toが未設定のため通知できません`);
                }
                else {
                    const vehicles = vehicleData
                        .slice(1)
                        .filter((v) => getCellValue(v, vehicleHeader['依頼ID']) === requestId);
                    const listText = vehicles.map((v, idx) => formatVehicleLineWithPolicy(v, vehicleHeader, tz, idx)).join('\n');
                    const subject = `【車両更新】承認済: ${dept} ${requestId}`;
                    const body = [
                        '専務承認が完了しました。以下の車両について（人手で）リース会社へ連絡し、完了フラグを更新してください。',
                        '',
                        `管理部門: ${dept}`,
                        `requestId: ${requestId}`,
                        `承認者: ${approverEmail}`,
                        `承認日時: ${formatDateLabel(now, tz)}`,
                        '',
                        '対象車両:',
                        listText,
                    ].join('\n');
                    try {
                        MailApp.sendEmail({
                            to,
                            cc: settings.vehicleManagerCc,
                            subject,
                            name: settings.fromName,
                            body,
                        });
                        requestRow[requestHeader['車両管理通知送信日時'] - 1] = now;
                        appendNotificationLog('承認結果通知', dept, to, requestId, '成功');
                        detailLines.push(`承認: requestId=${requestId} 結果通知 送信成功 -> ${to}`);
                    }
                    catch (err) {
                        appendNotificationLog('承認結果通知', dept, to, requestId, `失敗: ${err}`);
                        detailLines.push(`承認: requestId=${requestId} 結果通知 送信失敗 -> ${to} / ${err}`);
                    }
                }
                // 承認済はフォームを閉じる（差戻し再回答の窓を閉じるため）
                closeRequestForms(requestId);
                closeApprovalFormByRequestRow(requestRow, requestHeader);
                processed += 1;
                approvedRows += 1;
            }
            if (input === APPROVAL_INPUT.RETURN) {
                const comment = getCellValue(row, approvalHeader['差戻しコメント']);
                requestRow[requestHeader['承認ステータス'] - 1] = APPROVAL_STATUS.RETURNED;
                requestRow[requestHeader['承認者'] - 1] = approverEmail;
                requestRow[requestHeader['承認日時'] - 1] = now;
                requestRow[requestHeader['差戻しコメント'] - 1] = comment || '';
                const deptInfo = deptMaster[dept];
                if (!settings.mailSendEnabled) {
                    detailLines.push(`差戻し: requestId=${requestId} メール送信スキップ（通知_メール送信=FALSE）`);
                }
                else if (!deptInfo || !deptInfo.active || !deptInfo.to) {
                    appendNotificationLog('差戻し通知', dept, '', requestId, '部署マスタ未登録/無効/通知先Toなし');
                    detailLines.push(`差戻し: requestId=${requestId} 部署マスタの通知先Toが無い/無効のため通知できません`);
                }
                else {
                    const formUrls = extractFormUrlsFromRequestRow(requestRow, requestHeader);
                    const urlText = formUrls.length > 0 ? formatFormUrlsForText(formUrls) : '(フォームURLなし)';
                    const subject = `【車両更新】差戻し: ${dept} ${requestId}`;
                    const body = [
                        '専務より差戻しがありました。差戻しコメントを確認し、フォームで再回答してください。',
                        '',
                        `管理部門: ${dept}`,
                        `requestId: ${requestId}`,
                        `差戻しコメント: ${comment || ''}`,
                        '',
                        '再回答URL:',
                        urlText,
                    ].join('\n');
                    try {
                        // 差戻し時はフォームを再オープン（完了時に閉じているため）
                        reopenRequestForms(requestId);
                        MailApp.sendEmail({
                            to: deptInfo.to,
                            cc: deptInfo.cc,
                            subject,
                            name: settings.fromName,
                            body,
                        });
                        appendNotificationLog('差戻し通知', dept, deptInfo.to, requestId, '成功');
                        detailLines.push(`差戻し: requestId=${requestId} 通知 送信成功 -> ${deptInfo.to}`);
                    }
                    catch (err) {
                        appendNotificationLog('差戻し通知', dept, deptInfo.to, requestId, `失敗: ${err}`);
                        detailLines.push(`差戻し: requestId=${requestId} 通知 送信失敗 -> ${deptInfo.to} / ${err}`);
                    }
                }
                closeApprovalFormByRequestRow(requestRow, requestHeader);
                processed += 1;
                returnedRows += 1;
            }
        }
        // 入力値があるが承認/差戻し以外のものは「無視」されるため、ユーザーに見える化する
        if (inputRows > 0) {
            for (let i = 1; i < approvalData.length; i++) {
                const row = approvalData[i];
                const requestId = getCellValue(row, approvalHeader['requestId']);
                if (!requestId)
                    continue;
                const input = getCellValue(row, approvalHeader['承認入力']);
                if (!input)
                    continue;
                if (input === APPROVAL_INPUT.APPROVE || input === APPROVAL_INPUT.RETURN)
                    continue;
                invalidInputRows += 1;
                detailLines.push(`requestId=${requestId}: 承認入力が不正（'${input}'）のため無視。'承認' または '差戻し' を入力してください。`);
            }
        }
        reportLines.push(`入力行: ${inputRows}件`);
        reportLines.push(`反映: ${processed}件（承認:${approvedRows} / 差戻し:${returnedRows}）`);
        if (skippedSameRows > 0)
            reportLines.push(`スキップ: 既に同じステータス ${skippedSameRows}件`);
        if (missingRequestRows > 0)
            reportLines.push(`スキップ: 更新依頼が見つからない ${missingRequestRows}件`);
        if (invalidInputRows > 0)
            reportLines.push(`注意: 不正な承認入力 ${invalidInputRows}件`);
        if (processed > 0) {
            requestSheet.getRange(1, 1, requestData.length, requestData[0].length).setValues(requestData);
            if (vehicleData.length > 1) {
                vehicleSheet.getRange(1, 1, vehicleData.length, vehicleData[0].length).setValues(vehicleData);
            }
            buildApprovalQueueSheetInternal();
        }
        const body = [
            ...reportLines,
            '',
            `補足: 送信結果の詳細は「${SHEET_NAMES.NOTIFY_LOG}」タブにも記録されます。`,
            '',
            '詳細:',
            ...(detailLines.length > 0 ? detailLines : ['（なし）']),
        ].join('\n');
        uiShowModalSafe('承認結果反映（通知送信）', body);
    }
    finally {
        lock.releaseLock();
    }
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
function safeGetUserEmail() {
    try {
        const email = Session.getActiveUser().getEmail();
        return email || '(不明)';
    }
    catch (err) {
        return '(不明)';
    }
}
function formatVehicleLineForApproval(row, headerMap, tz, rowIndex) {
    const base = formatVehicleLine(row, headerMap, tz);
    const policy = headerMap['更新方針'] ? getCellValue(row, headerMap['更新方針']) : '';
    const note = headerMap['備考'] ? getCellValue(row, headerMap['備考']) : '';
    const policyLabel = policy ? ` / 一次:${policy}` : '';
    const noteLabel = note ? ` / 備考:${note}` : '';
    return `${rowIndex + 1}. ${base}${policyLabel}${noteLabel}`;
}
function formatVehicleLineWithPolicy(row, headerMap, tz, rowIndex) {
    const reg = getCellValue(row, headerMap['登録番号_結合']);
    const type = getCellValue(row, headerMap['車種']);
    const end = parseDateValue(getCellRaw(row, headerMap['契約満了日']));
    const endLabel = end ? formatDateLabel(end, tz) : '未設定';
    const primary = getCellValue(row, headerMap['更新方針']);
    const note = getCellValue(row, headerMap['備考']);
    const primaryLabel = primary ? primary : '未回答';
    return `【${rowIndex + 1}】${reg || '登録番号不明'} / ${type || '車種不明'} / 満了:${endLabel} / 一次:${primaryLabel}${note ? ` / 備考:${note}` : ''}`;
}
function reopenRequestForms(requestId) {
    const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.REQUESTS);
    if (!sheet)
        return;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1)
        return;
    const headerMap = getHeaderMap(data[0]);
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (getCellValue(row, headerMap['requestId']) !== requestId)
            continue;
        const formIds = extractFormIdsFromRequestRow(row, headerMap);
        formIds.forEach((formId) => {
            try {
                const form = FormApp.openById(formId);
                form.setAcceptingResponses(true);
                try {
                    form.setShowLinkToRespondAgain(true);
                }
                catch (err) {
                    // setShowLinkToRespondAgain が無い環境でも致命ではない
                }
                ensureFormSubmitTrigger(form);
            }
            catch (err) {
                Logger.log(`reopenRequestForms: ${formId} ${err}`);
            }
        });
        break;
    }
}
function ensureFormSubmitTrigger(form) {
    const formId = form.getId();
    const triggers = ScriptApp.getProjectTriggers();
    const exists = triggers.some((trigger) => {
        if (trigger.getHandlerFunction() !== 'onRequestFormSubmit')
            return false;
        const sourceId = typeof trigger.getTriggerSourceId === 'function' ? trigger.getTriggerSourceId() : '';
        return sourceId === formId;
    });
    if (exists)
        return;
    ScriptApp.newTrigger('onRequestFormSubmit').forForm(form).onFormSubmit().create();
}
function runDaily() {
    syncSchema();
    syncVehicles();
    createRequests();
    sendInitialEmails();
    sendReminderEmails();
    buildSummarySheet();
    sendSummaryEmail();
    buildApprovalQueueSheet();
    sendApprovalRequestEmails();
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
function seedTestVehicles() {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const ss = getSpreadsheet();
        const tz = ss.getSpreadsheetTimeZone();
        const baseDate = toDateOnly(new Date(), tz);
        const inRangeDate = addDays(baseDate, 1);
        const outRangeDate = addMonthsClamped(baseDate, 7);
        const deptMaster = loadDeptMaster();
        const validDept = pickFirstActiveDept(deptMaster);
        if (!validDept) {
            uiAlertSafe('部署マスタに有効な管理部門がありません。先に登録してください。');
            return {
                addedTotal: 0,
                skippedSheets: SOURCE_SHEETS.slice(),
                skippedReasons: { _global: '部署マスタに有効な管理部門がありません' },
            };
        }
        const scenarios = [
            { code: 'IN', label: '期限内', contractEnd: inRangeDate, dept: validDept },
            { code: 'OUT', label: '期限外', contractEnd: outRangeDate, dept: validDept },
            { code: 'NOEND', label: '満了日なし', contractEnd: null, dept: validDept },
            { code: 'NODEPT', label: '管理部門なし', contractEnd: inRangeDate, dept: '' },
            { code: 'UNREG', label: '部署未登録', contractEnd: inRangeDate, dept: '未登録部署_TEST' },
        ];
        let addedTotal = 0;
        let skippedSheets = [];
        const skippedReasons = {};
        SOURCE_SHEETS.forEach((sheetName, index) => {
            const sheet = ss.getSheetByName(sheetName);
            if (!sheet) {
                skippedSheets.push(sheetName);
                skippedReasons[sheetName] = 'シート未存在';
                return;
            }
            const data = sheet.getDataRange().getValues();
            if (data.length === 0) {
                skippedSheets.push(sheetName);
                skippedReasons[sheetName] = 'データが空';
                return;
            }
            const headerMap = getHeaderMap(data[0]);
            const idx = resolveSourceHeaders(headerMap);
            const hasSplitReg = !!(idx.regArea && idx.regClass && idx.regKana && idx.regNumber);
            const hasAnyReg = hasSplitReg || !!idx.regAll;
            if (!hasAnyReg || !idx.dept || !idx.contractEnd) {
                skippedSheets.push(sheetName);
                const missing = [];
                if (!hasAnyReg)
                    missing.push('登録番号');
                if (!idx.dept)
                    missing.push('管理部門');
                if (!idx.contractEnd)
                    missing.push('契約満了日');
                skippedReasons[sheetName] = `必須ヘッダ不足: ${missing.join(', ')}`;
                return;
            }
            const existingRegs = {};
            for (let i = 1; i < data.length; i++) {
                const row = data[i];
                const regCombined = getSourceRegistrationCombined(row, idx);
                if (regCombined)
                    existingRegs[regCombined] = true;
            }
            const rowsToAdd = [];
            const sheetCode = String(index + 1).padStart(2, '0');
            scenarios.forEach((scenario) => {
                const regArea = 'TEST';
                const regClass = sheetCode;
                const regKana = 'テ';
                const regNumber = `T${sheetCode}-${scenario.code}`;
                const regCombined = buildRegistrationCombined(regArea, regClass, regKana, regNumber);
                if (existingRegs[regCombined])
                    return;
                const row = new Array(data[0].length).fill('');
                if (hasSplitReg) {
                    row[idx.regArea - 1] = regArea;
                    row[idx.regClass - 1] = regClass;
                    row[idx.regKana - 1] = regKana;
                    row[idx.regNumber - 1] = regNumber;
                }
                else if (idx.regAll) {
                    row[idx.regAll - 1] = regCombined;
                }
                if (idx.vehicleType)
                    row[idx.vehicleType - 1] = `テスト_${scenario.label}`;
                if (idx.chassis)
                    row[idx.chassis - 1] = `TEST-${sheetCode}-${scenario.code}`;
                if (idx.contractStart)
                    row[idx.contractStart - 1] = baseDate;
                if (idx.contractEnd && scenario.contractEnd)
                    row[idx.contractEnd - 1] = scenario.contractEnd;
                row[idx.dept - 1] = scenario.dept;
                rowsToAdd.push(row);
                existingRegs[regCombined] = true;
            });
            if (rowsToAdd.length > 0) {
                sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, data[0].length).setValues(rowsToAdd);
                addedTotal += rowsToAdd.length;
            }
        });
        const skippedDetail = skippedSheets
            .map((name) => `${name}(${skippedReasons[name] || '不明'})`)
            .join(', ');
        const message = skippedSheets.length
            ? `テスト車両を追加しました（合計 ${addedTotal} 件）。\n未処理シート: ${skippedDetail}`
            : `テスト車両を追加しました（合計 ${addedTotal} 件）。`;
        uiAlertSafe(message);
        return { addedTotal, skippedSheets, skippedReasons };
    }
    finally {
        lock.releaseLock();
    }
}
function diagnoseSourceSheets() {
    const ss = getSpreadsheet();
    const results = [];
    SOURCE_SHEETS.forEach((sheetName) => {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
            results.push({ sheetName, ok: false, reason: 'シート未存在' });
            return;
        }
        const data = sheet.getDataRange().getValues();
        if (data.length === 0) {
            results.push({ sheetName, ok: false, reason: 'データが空' });
            return;
        }
        const headers = data[0].map((h) => String(h || '').trim()).filter((h) => h);
        const normalizedHeaders = headers.map((h) => normalizeHeaderKey(h)).filter((h) => h);
        const headerMap = getHeaderMap(data[0]);
        const idx = resolveSourceHeaders(headerMap);
        const missing = [];
        const hasSplitReg = !!(idx.regArea && idx.regClass && idx.regKana && idx.regNumber);
        const hasAnyReg = hasSplitReg || !!idx.regAll;
        if (!hasAnyReg)
            missing.push('登録番号');
        if (!idx.dept)
            missing.push('管理部門');
        if (!idx.contractEnd)
            missing.push('契約満了日');
        results.push({
            sheetName,
            ok: missing.length === 0,
            missing,
            registrationMode: hasSplitReg ? 'split' : idx.regAll ? 'combined' : 'missing',
            headers,
            normalizedHeaders,
        });
    });
    Logger.log(JSON.stringify(results, null, 2));
    appendTestResult('ソースシート診断', results.every((r) => r.ok) ? 'OK' : 'NG', JSON.stringify(results));
    uiAlertSafe('診断結果を Logger と テスト結果 シートに出力しました。');
    return results;
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
        const testDept = 'テスト管理部門';
        const removed = {
            sourceSheets: {},
            vehicleView: 0,
            needsInput: 0,
            requests: 0,
            answers: 0,
            summary: 0,
            notifyLog: 0,
            deptMaster: 0,
            testRequestIds: 0,
        };
        const testVehicleIds = {};
        const testRequestIds = {};
        // 車両（統合ビュー）からテスト由来の vehicleId / requestId を収集しつつ削除
        const vehicleViewSheet = ss.getSheetByName(SHEET_NAMES.VEHICLE_VIEW);
        if (vehicleViewSheet) {
            const data = vehicleViewSheet.getDataRange().getValues();
            if (data.length > 1) {
                const header = getHeaderMap(data[0]);
                const idx = {
                    vehicleId: header['vehicleId'],
                    regCombined: header['登録番号_結合'],
                    requestId: header['依頼ID'],
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
                    const requestId = getCellValue(row, idx.requestId);
                    if (requestId)
                        testRequestIds[requestId] = true;
                    rowsToDelete.push(i + 1);
                }
                for (let i = rowsToDelete.length - 1; i >= 0; i--) {
                    vehicleViewSheet.deleteRow(rowsToDelete[i]);
                    removed.vehicleView += 1;
                }
            }
        }
        removed.testRequestIds = Object.keys(testRequestIds).length;
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
        // 更新依頼（テスト管理部門 or テスト requestId のみ削除）
        const requestSheet = ss.getSheetByName(SHEET_NAMES.REQUESTS);
        if (requestSheet) {
            const data = requestSheet.getDataRange().getValues();
            if (data.length > 1) {
                const header = getHeaderMap(data[0]);
                const idx = {
                    requestId: header['requestId'],
                    dept: header['管理部門'],
                };
                const rowsToDelete = [];
                for (let i = 1; i < data.length; i++) {
                    const row = data[i];
                    const requestId = getCellValue(row, idx.requestId);
                    const dept = getCellValue(row, idx.dept);
                    const isTest = (dept && dept === testDept) || (requestId && testRequestIds[requestId]);
                    if (!isTest)
                        continue;
                    rowsToDelete.push(i + 1);
                }
                for (let i = rowsToDelete.length - 1; i >= 0; i--) {
                    requestSheet.deleteRow(rowsToDelete[i]);
                    removed.requests += 1;
                }
            }
        }
        // 回答（テスト requestId / テスト vehicleId のみ削除）
        const answerSheet = ss.getSheetByName(SHEET_NAMES.ANSWERS);
        if (answerSheet) {
            const data = answerSheet.getDataRange().getValues();
            if (data.length > 1) {
                const header = getHeaderMap(data[0]);
                const idx = {
                    requestId: header['requestId'],
                    vehicleId: header['vehicleId'],
                };
                const rowsToDelete = [];
                for (let i = 1; i < data.length; i++) {
                    const row = data[i];
                    const requestId = getCellValue(row, idx.requestId);
                    const vehicleId = getCellValue(row, idx.vehicleId);
                    const isTest = (requestId && testRequestIds[requestId]) || (vehicleId && testVehicleIds[vehicleId]);
                    if (!isTest)
                        continue;
                    rowsToDelete.push(i + 1);
                }
                for (let i = rowsToDelete.length - 1; i >= 0; i--) {
                    answerSheet.deleteRow(rowsToDelete[i]);
                    removed.answers += 1;
                }
            }
        }
        // 回答集計（テスト管理部門 or テスト requestId のみ削除）
        const summarySheet = ss.getSheetByName(SHEET_NAMES.SUMMARY);
        if (summarySheet) {
            const data = summarySheet.getDataRange().getValues();
            if (data.length > 1) {
                const header = getHeaderMap(data[0]);
                const idx = {
                    requestId: header['requestId'],
                    dept: header['管理部門'],
                };
                const rowsToDelete = [];
                for (let i = 1; i < data.length; i++) {
                    const row = data[i];
                    const requestId = getCellValue(row, idx.requestId);
                    const dept = getCellValue(row, idx.dept);
                    const isTest = (dept && dept === testDept) || (requestId && testRequestIds[requestId]);
                    if (!isTest)
                        continue;
                    rowsToDelete.push(i + 1);
                }
                for (let i = rowsToDelete.length - 1; i >= 0; i--) {
                    summarySheet.deleteRow(rowsToDelete[i]);
                    removed.summary += 1;
                }
            }
        }
        // 通知ログ（テスト管理部門 or テスト requestId のみ削除）
        const notifyLogSheet = ss.getSheetByName(SHEET_NAMES.NOTIFY_LOG);
        if (notifyLogSheet) {
            const data = notifyLogSheet.getDataRange().getValues();
            if (data.length > 1) {
                const header = getHeaderMap(data[0]);
                const idx = {
                    requestId: header['requestId'],
                    dept: header['管理部門'],
                };
                const rowsToDelete = [];
                for (let i = 1; i < data.length; i++) {
                    const row = data[i];
                    const requestId = getCellValue(row, idx.requestId);
                    const dept = getCellValue(row, idx.dept);
                    const isTest = (dept && dept === testDept) || (requestId && testRequestIds[requestId]);
                    if (!isTest)
                        continue;
                    rowsToDelete.push(i + 1);
                }
                for (let i = rowsToDelete.length - 1; i >= 0; i--) {
                    notifyLogSheet.deleteRow(rowsToDelete[i]);
                    removed.notifyLog += 1;
                }
            }
        }
        // 元台帳（3シート）からテスト車両行を削除
        SOURCE_SHEETS.forEach((sheetName) => {
            const sheet = ss.getSheetByName(sheetName);
            if (!sheet)
                return;
            const data = sheet.getDataRange().getValues();
            if (data.length <= 1)
                return;
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
                sheet.deleteRow(rowsToDelete[i]);
            }
            removed.sourceSheets[sheetName] = rowsToDelete.length;
        });
        // 部署マスタのテスト行は「テスト管理部門」を使っている車両が残っていない場合のみ削除
        let testDeptInUse = false;
        SOURCE_SHEETS.forEach((sheetName) => {
            const sheet = ss.getSheetByName(sheetName);
            if (!sheet)
                return;
            const data = sheet.getDataRange().getValues();
            if (data.length <= 1)
                return;
            const headerMap = getHeaderMap(data[0]);
            const idx = resolveSourceHeaders(headerMap);
            if (!idx.dept)
                return;
            for (let i = 1; i < data.length; i++) {
                const dept = getCellValue(data[i], idx.dept);
                if (dept === testDept) {
                    testDeptInUse = true;
                    return;
                }
            }
        });
        const deptSheet = ss.getSheetByName(SHEET_NAMES.DEPT_MASTER);
        if (deptSheet && !testDeptInUse) {
            const data = deptSheet.getDataRange().getValues();
            if (data.length > 1) {
                const header = getHeaderMap(data[0]);
                const idxDept = header['管理部門'];
                const rowsToDelete = [];
                for (let i = 1; i < data.length; i++) {
                    const dept = getCellValue(data[i], idxDept);
                    if (dept === testDept)
                        rowsToDelete.push(i + 1);
                }
                for (let i = rowsToDelete.length - 1; i >= 0; i--) {
                    deptSheet.deleteRow(rowsToDelete[i]);
                    removed.deptMaster += 1;
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
function runTestSuite() {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        clearTestResults();
        appendTestResult('開始', 'OK', new Date().toISOString());
        syncSchema();
        appendTestResult('syncSchema', 'OK', '');
        generateDeptTokens();
        appendTestResult('generateDeptTokens', 'OK', '空欄のみ生成');
        const diag = diagnoseSourceSheets();
        if (!diag.every((r) => r.ok)) {
            appendTestResult('中断', 'NG', 'ソースシートの必須ヘッダが不足しています');
            return;
        }
        const seed = seedTestVehicles();
        if (seed && seed.skippedSheets && seed.skippedSheets.length > 0) {
            appendTestResult('seedTestVehicles', 'NG', JSON.stringify(seed));
            appendTestResult('中断', 'NG', 'テスト車両を投入できないシートがあります');
            return;
        }
        appendTestResult('seedTestVehicles', 'OK', seed ? JSON.stringify(seed) : '');
        syncVehicles();
        appendTestResult('syncVehicles', 'OK', '');
        // 期待値チェック（再実行でもOKな形）
        const ss = getSpreadsheet();
        const tz = ss.getSpreadsheetTimeZone();
        const settings = loadSettings();
        const deptMaster = loadDeptMaster();
        const validDept = pickFirstActiveDept(deptMaster);
        const vehicleSheet = ss.getSheetByName(SHEET_NAMES.VEHICLE_VIEW);
        if (vehicleSheet) {
            const data = vehicleSheet.getDataRange().getValues();
            const header = data.length > 0 ? getHeaderMap(data[0]) : {};
            const idx = {
                regCombined: header['登録番号_結合'],
                dept: header['管理部門'],
                contractEnd: header['契約満了日'],
                requestId: header['依頼ID'],
            };
            const startDate = toDateOnly(new Date(), tz);
            const endDate = addMonthsClamped(startDate, settings.expiryMonths);
            let testTotal = 0;
            let testInRange = 0;
            let testInRangeWithRequestId = 0;
            for (let i = 1; i < data.length; i++) {
                const row = data[i];
                const reg = getCellValue(row, idx.regCombined);
                if (!reg || !reg.startsWith('TEST'))
                    continue;
                testTotal += 1;
                const dept = getCellValue(row, idx.dept);
                const contractEnd = parseDateValue(getCellRaw(row, idx.contractEnd));
                const contractDate = contractEnd ? toDateOnly(contractEnd, tz) : null;
                if (dept === validDept && contractDate && isWithinRange(contractDate, startDate, endDate)) {
                    testInRange += 1;
                    if (getCellValue(row, idx.requestId))
                        testInRangeWithRequestId += 1;
                }
            }
            appendTestResult('期待値:統合ビュー_テスト車両件数', testTotal >= 3 ? 'OK' : 'NG', String(testTotal));
            appendTestResult('期待値:統合ビュー_期限内車両件数', testInRange >= SOURCE_SHEETS.length ? 'OK' : 'NG', `dept=${validDept || '(empty)'} count=${testInRange}`);
            // createRequests 前なので依頼IDは「付いていても付いていなくても」OK（再実行想定）
            appendTestResult('期待値:統合ビュー_期限内_依頼ID付与済(参考)', testInRangeWithRequestId <= testInRange ? 'OK' : 'NG', `${testInRangeWithRequestId}/${testInRange}台`);
        }
        const needsInputSheet = ss.getSheetByName(SHEET_NAMES.NEEDS_INPUT);
        if (needsInputSheet) {
            const data = needsInputSheet.getDataRange().getValues();
            const header = data.length > 0 ? getHeaderMap(data[0]) : {};
            const idx = { reason: header['不備内容'], reg: header['登録番号_結合'] };
            const counts = {
                契約満了日なし: 0,
                管理部門なし: 0,
                部署マスタ未登録: 0,
            };
            for (let i = 1; i < data.length; i++) {
                const row = data[i];
                const reg = getCellValue(row, idx.reg);
                // テスト車両は登録番号_結合が "TEST..." になる
                if (reg && !reg.startsWith('TEST'))
                    continue;
                const reason = getCellValue(row, idx.reason);
                if (counts[reason] !== undefined)
                    counts[reason] += 1;
            }
            Object.keys(counts).forEach((key) => {
                appendTestResult(`期待値:要入力_${key}`, counts[key] >= 1 ? 'OK' : 'NG', String(counts[key]));
            });
        }
        const requestSheet = ss.getSheetByName(SHEET_NAMES.REQUESTS);
        const beforeRequestLastRow = requestSheet ? requestSheet.getLastRow() : 0;
        createRequests();
        const afterRequestLastRow = requestSheet ? requestSheet.getLastRow() : 0;
        appendTestResult('createRequests', 'OK', `newRows=${Math.max(0, afterRequestLastRow - beforeRequestLastRow)}`);
        // 期待値: createRequests は同じ入力に対して増え続けない（重複防止）
        const beforeSecondLastRow = requestSheet ? requestSheet.getLastRow() : 0;
        createRequests();
        const afterSecondLastRow = requestSheet ? requestSheet.getLastRow() : 0;
        appendTestResult('期待値:createRequests_重複防止', afterSecondLastRow === beforeSecondLastRow ? 'OK' : 'NG', `newRows=${Math.max(0, afterSecondLastRow - beforeSecondLastRow)}`);
        // 期待値: 期限内テスト車両へ依頼IDが付与され、依頼シートに管理部門行が存在する
        if (vehicleSheet) {
            const data = vehicleSheet.getDataRange().getValues();
            const header = data.length > 0 ? getHeaderMap(data[0]) : {};
            const idx = {
                regCombined: header['登録番号_結合'],
                dept: header['管理部門'],
                contractEnd: header['契約満了日'],
                requestId: header['依頼ID'],
            };
            const startDate = toDateOnly(new Date(), tz);
            const endDate = addMonthsClamped(startDate, settings.expiryMonths);
            let testInRange = 0;
            let testInRangeWithRequestId = 0;
            for (let i = 1; i < data.length; i++) {
                const row = data[i];
                const reg = getCellValue(row, idx.regCombined);
                if (!reg || !reg.startsWith('TEST'))
                    continue;
                const dept = getCellValue(row, idx.dept);
                const contractEnd = parseDateValue(getCellRaw(row, idx.contractEnd));
                const contractDate = contractEnd ? toDateOnly(contractEnd, tz) : null;
                if (dept === validDept && contractDate && isWithinRange(contractDate, startDate, endDate)) {
                    testInRange += 1;
                    if (getCellValue(row, idx.requestId))
                        testInRangeWithRequestId += 1;
                }
            }
            appendTestResult('期待値:createRequests_期限内_依頼ID付与', testInRange > 0 && testInRangeWithRequestId === testInRange ? 'OK' : 'NG', `${testInRangeWithRequestId}/${testInRange}台`);
        }
        if (requestSheet) {
            const data = requestSheet.getDataRange().getValues();
            const header = data.length > 0 ? getHeaderMap(data[0]) : {};
            const idx = { dept: header['管理部門'], requestId: header['requestId'] };
            let count = 0;
            for (let i = 1; i < data.length; i++) {
                const row = data[i];
                if (!getCellValue(row, idx.requestId))
                    continue;
                if (validDept && getCellValue(row, idx.dept) === validDept)
                    count += 1;
            }
            appendTestResult('期待値:createRequests_依頼行(dept)', count >= 1 ? 'OK' : 'NG', `dept=${validDept} count=${count}`);
        }
        buildSummarySheet();
        appendTestResult('buildSummarySheet', 'OK', '');
        appendTestResult('完了', 'OK', '');
    }
    finally {
        lock.releaseLock();
    }
}
function generateDeptTokens() {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const ss = getSpreadsheet();
        const sheet = ensureSheet(ss, SHEET_NAMES.DEPT_MASTER);
        ensureHeaders(sheet, 1, getSchemaHeaders(SHEET_NAMES.DEPT_MASTER));
        const data = sheet.getDataRange().getValues();
        if (data.length <= 1)
            return;
        const headerMap = getHeaderMap(data[0]);
        const deptIndex = headerMap['管理部門'];
        const tokenIndex = headerMap['部門トークン'];
        if (!deptIndex || !tokenIndex)
            return;
        let changed = false;
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const dept = getCellValue(row, deptIndex);
            if (!dept)
                continue;
            const token = getCellValue(row, tokenIndex);
            if (!token) {
                row[tokenIndex - 1] = generateToken();
                changed = true;
            }
        }
        if (changed) {
            sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
        }
    }
    finally {
        lock.releaseLock();
    }
}
function setWebAppUrl(url) {
    if (!url)
        throw new Error('Web回答URLが指定されていません');
    setSettingValue('Web回答URL（デプロイURL）', url);
}
function installDailyTriggers() {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach((trigger) => {
        if (trigger.getHandlerFunction() === 'runDaily') {
            ScriptApp.deleteTrigger(trigger);
        }
    });
    const hour = 8;
    const weekdays = [
        ScriptApp.WeekDay.MONDAY,
        ScriptApp.WeekDay.TUESDAY,
        ScriptApp.WeekDay.WEDNESDAY,
        ScriptApp.WeekDay.THURSDAY,
        ScriptApp.WeekDay.FRIDAY,
    ];
    weekdays.forEach((day) => {
        ScriptApp.newTrigger('runDaily').timeBased().onWeekDay(day).atHour(hour).create();
    });
}
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
function loadDeptMaster() {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.DEPT_MASTER);
    if (!sheet)
        return {};
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1)
        return {};
    const headerMap = getHeaderMap(data[0]);
    const result = {};
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const dept = getCellValue(row, headerMap['管理部門']);
        if (!dept)
            continue;
        const activeValue = getCellRaw(row, headerMap['有効']);
        const active = toBoolean(activeValue, true);
        result[dept] = {
            to: getCellValue(row, headerMap['通知先To']),
            cc: getCellValue(row, headerMap['通知先Cc']),
            token: getCellValue(row, headerMap['部門トークン']),
            active,
        };
    }
    return result;
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
        expiryMonths: toNumber(values['抽出_満了まで月数'], Number(SETTINGS_DEFAULTS['抽出_満了まで月数'])),
        reminderStartAfterDays: toNumber(values['リマインド_初回から日数'], Number(SETTINGS_DEFAULTS['リマインド_初回から日数'])),
        reminderIntervalDays: toNumber(values['リマインド_間隔日数'], Number(SETTINGS_DEFAULTS['リマインド_間隔日数'])),
        reminderMaxCount: toNumber(values['リマインド_最大回数'], Number(SETTINGS_DEFAULTS['リマインド_最大回数'])),
        deadlineAfterDays: toNumber(values['締切_初回送信から日数'], Number(SETTINGS_DEFAULTS['締切_初回送信から日数'])),
        fromName: toStringValue(values['送信元名'], String(SETTINGS_DEFAULTS['送信元名'])),
        subjectTemplate: toStringValue(values['件名テンプレ'], String(SETTINGS_DEFAULTS['件名テンプレ'])),
        bodyTemplate: toStringValue(values['本文テンプレ'], String(SETTINGS_DEFAULTS['本文テンプレ'])),
        webAppUrl: toStringValue(values['Web回答URL（デプロイURL）'], String(SETTINGS_DEFAULTS['Web回答URL（デプロイURL）'])),
        adminTo: toStringValue(values['管理者_通知先To'], String(SETTINGS_DEFAULTS['管理者_通知先To'])),
        adminCc: toStringValue(values['管理者_通知先Cc'], String(SETTINGS_DEFAULTS['管理者_通知先Cc'])),
        mailSendEnabled: toBoolean(values['通知_メール送信'], Boolean(SETTINGS_DEFAULTS['通知_メール送信'])),
        summarySheetEnabled: toBoolean(values['集計_シート出力'], Boolean(SETTINGS_DEFAULTS['集計_シート出力'])),
        summaryEmailEnabled: toBoolean(values['集計_メール送信'], Boolean(SETTINGS_DEFAULTS['集計_メール送信'])),
        approvalFlowEnabled: toBoolean(values['承認フロー_有効'], Boolean(SETTINGS_DEFAULTS['承認フロー_有効'])),
        approverTo: toStringValue(values['承認者_通知先To'], String(SETTINGS_DEFAULTS['承認者_通知先To'])),
        approverCc: toStringValue(values['承認者_通知先Cc'], String(SETTINGS_DEFAULTS['承認者_通知先Cc'])),
        vehicleManagerTo: toStringValue(values['車両管理担当_通知先To'], String(SETTINGS_DEFAULTS['車両管理担当_通知先To'])),
        vehicleManagerCc: toStringValue(values['車両管理担当_通知先Cc'], String(SETTINGS_DEFAULTS['車両管理担当_通知先Cc'])),
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
function pickFirstActiveDept(deptMaster) {
    const keys = Object.keys(deptMaster);
    for (const key of keys) {
        if (deptMaster[key].active)
            return key;
    }
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
function generateRequestId(date) {
    const tz = getSpreadsheet().getSpreadsheetTimeZone();
    const stamp = Utilities.formatDate(date, tz, 'yyyyMMddHHmmss');
    const rand = Math.floor(Math.random() * 9000 + 1000);
    return `REQ-${stamp}-${rand}`;
}
function generateToken() {
    const base = Utilities.getUuid().replace(/-/g, '');
    const extra = Utilities.getUuid().replace(/-/g, '');
    return base + extra;
}
function applyTemplate(template, params) {
    let result = template;
    Object.keys(params).forEach((key) => {
        const regex = new RegExp(`{{${key}}}`, 'g');
        result = result.replace(regex, params[key]);
    });
    return result;
}
function formatDateLabel(date, tz) {
    return Utilities.formatDate(date, tz, 'yyyy/MM/dd');
}
function buildWebAppUrl(baseUrl, params) {
    const query = Object.keys(params)
        .map((key) => `${encodeURIComponent(key)}=${encodeURIComponent(params[key])}`)
        .join('&');
    return `${baseUrl}?${query}`;
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
function formatDateIsoLabel(date, tz) {
    return Utilities.formatDate(date, tz, 'yyyy-MM-dd');
}
function formatFormUrlsForText(urls) {
    if (urls.length === 1)
        return urls[0];
    return urls.map((url, index) => `Part${index + 1}: ${url}`).join('\n');
}
function formatFormUrlsForHtml(urls) {
    return urls
        .map((url, index) => {
        const label = urls.length > 1 ? `Part${index + 1}: ` : '';
        const escaped = escapeHtml(url);
        return `${label}<a href="${escaped}">${escaped}</a>`;
    })
        .join('<br>');
}
function extractFormUrlsFromRequestRow(row, headerMap) {
    const urls = parseJsonStringArray(getCellValue(row, headerMap['formUrlsJson']));
    if (urls.length > 0)
        return urls;
    const url = getCellValue(row, headerMap['formUrl']);
    return url ? [String(url)] : [];
}
function extractFormIdsFromRequestRow(row, headerMap) {
    const ids = parseJsonStringArray(getCellValue(row, headerMap['formIdsJson']));
    if (ids.length > 0)
        return ids;
    const id = getCellValue(row, headerMap['formId']);
    return id ? [String(id)] : [];
}
function applyFormResultToRequestRow(row, headerMap, result, createdAt) {
    const setCell = (headerName, value) => {
        const idx = headerMap[headerName];
        if (idx)
            row[idx - 1] = value;
    };
    setCell('formId', result.formIds.length === 1 ? result.formIds[0] : '');
    setCell('formUrl', result.formUrls.length === 1 ? result.formUrls[0] : '');
    setCell('formIdsJson', result.formIds.length > 1 ? JSON.stringify(result.formIds) : '');
    setCell('formUrlsJson', result.formUrls.length > 1 ? JSON.stringify(result.formUrls) : '');
    setCell('formEditUrl', result.formEditUrls.length === 1 ? result.formEditUrls[0] : '');
    setCell('formTriggerId', result.formTriggerIds.length === 1 ? result.formTriggerIds[0] : '');
    setCell('フォーム作成日時', createdAt);
}
function createRequestForms(params) {
    try {
        const chunks = chunkArray(params.vehicles, MAX_VEHICLES_PER_FORM);
        const parts = chunks.length;
        const formIds = [];
        const formUrls = [];
        const formEditUrls = [];
        const formTriggerIds = [];
        const props = PropertiesService.getDocumentProperties();
        chunks.forEach((chunk, index) => {
            const title = buildFormTitle(params.dept, params.requestId, index, parts);
            const form = FormApp.create(title);
            applyFormPublicSettings(form);
            form.setDescription(buildFormDescription(params, index, parts, params.tz));
            form.setConfirmationMessage('回答を受け付けました。ありがとうございました。');
            ensureFormExplanationHeader(form);
            const gridItem = form.addGridItem();
            gridItem.setTitle(FORM_ITEM_TITLES.POLICY_GRID);
            gridItem.setRows(chunk.map((row, rowIndex) => buildFormVehicleRowLabel(row, params.vehicleHeader, params.tz, rowIndex)));
            gridItem.setColumns(ANSWER_OPTIONS);
            gridItem.setRequired(true);
            const vehicleIds = chunk.map((row) => getCellValue(row, params.vehicleHeader['vehicleId']) || '');
            props.setProperty(buildFormVehicleIdsPropKey(form.getId()), JSON.stringify(vehicleIds));
            const trigger = ScriptApp.newTrigger('onRequestFormSubmit').forForm(form).onFormSubmit().create();
            const triggerId = typeof trigger.getUniqueId === 'function' ? trigger.getUniqueId() : '';
            formIds.push(form.getId());
            formUrls.push(form.getPublishedUrl());
            formEditUrls.push(form.getEditUrl());
            formTriggerIds.push(triggerId);
        });
        return {
            ok: true,
            message: '',
            formIds,
            formUrls,
            formEditUrls,
            formTriggerIds,
        };
    }
    catch (err) {
        return {
            ok: false,
            message: err ? String(err) : 'フォーム作成に失敗しました',
            formIds: [],
            formUrls: [],
            formEditUrls: [],
            formTriggerIds: [],
        };
    }
}
function createOrUpdateApprovalForm(params) {
    try {
        let form = null;
        if (params.existingFormId) {
            try {
                form = FormApp.openById(params.existingFormId);
            }
            catch (err) {
                form = null;
            }
        }
        if (!form) {
            form = FormApp.create(`【承認依頼】${params.dept} ${params.requestId}`);
        }
        applyFormPublicSettings(form);
        form.setShowLinkToRespondAgain(false);
        form.setTitle(`【承認依頼】${params.dept} ${params.requestId}`);
        form.setDescription(buildApprovalFormDescription(params));
        form.setConfirmationMessage('承認判断を受け付けました。ありがとうございました。');
        form.setAcceptingResponses(true);
        ensureApprovalDecisionItems(form);
        const triggerId = ensureApprovalFormSubmitTrigger(form);
        PropertiesService.getDocumentProperties().setProperty(buildApprovalFormRequestIdPropKey(form.getId()), params.requestId);
        return {
            ok: true,
            message: '',
            formId: form.getId(),
            formUrl: form.getPublishedUrl(),
            formEditUrl: form.getEditUrl(),
            formTriggerId: triggerId,
        };
    }
    catch (err) {
        return {
            ok: false,
            message: err ? String(err) : '承認フォーム作成に失敗しました',
            formId: '',
            formUrl: '',
            formEditUrl: '',
            formTriggerId: '',
        };
    }
}
function buildApprovalFormDescription(params) {
    const startLabel = params.targetStart ? formatDateLabel(params.targetStart, params.tz) : '-';
    const endLabel = params.targetEnd ? formatDateLabel(params.targetEnd, params.tz) : '-';
    const lines = [
        `requestId: ${params.requestId}`,
        `管理部門: ${params.dept || '-'}`,
        `対象期間: ${startLabel}〜${endLabel}`,
        `一次回答サマリ: ${params.summaryText || '-'}`,
        '',
        '対象車両:',
        params.vehiclesText || '（対象車両なし）',
        '',
        'ご確認項目: 承認 または 差戻し を選択してください。',
        '※差戻しを選ぶ場合は、差戻しコメントを入力してください。',
        '※このフォームURLは転送しないでください。',
    ];
    return lines.join('\n');
}
function ensureApprovalDecisionItems(form) {
    const allItems = form.getItems();
    let decisionItem = null;
    let commentItem = null;
    allItems.forEach((item) => {
        const title = item.getTitle();
        if (title === APPROVAL_FORM_TITLES.DECISION && item.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
            decisionItem = item.asMultipleChoiceItem();
            return;
        }
        if (title === APPROVAL_FORM_TITLES.COMMENT && item.getType() === FormApp.ItemType.PARAGRAPH_TEXT) {
            commentItem = item.asParagraphTextItem();
        }
    });
    if (!decisionItem) {
        decisionItem = form.addMultipleChoiceItem();
        decisionItem.setTitle(APPROVAL_FORM_TITLES.DECISION);
    }
    decisionItem.setChoiceValues([APPROVAL_INPUT.APPROVE, APPROVAL_INPUT.RETURN]);
    decisionItem.setRequired(true);
    if (!commentItem) {
        commentItem = form.addParagraphTextItem();
        commentItem.setTitle(APPROVAL_FORM_TITLES.COMMENT);
    }
    commentItem.setHelpText('差戻し時は理由を記入してください。');
}
function ensureApprovalFormSubmitTrigger(form) {
    const formId = form.getId();
    const triggers = ScriptApp.getProjectTriggers();
    const existing = triggers.find((trigger) => {
        if (trigger.getHandlerFunction() !== 'onApprovalFormSubmit')
            return false;
        const sourceId = typeof trigger.getTriggerSourceId === 'function' ? trigger.getTriggerSourceId() : '';
        return sourceId === formId;
    });
    if (existing) {
        return typeof existing.getUniqueId === 'function' ? existing.getUniqueId() : '';
    }
    const trigger = ScriptApp.newTrigger('onApprovalFormSubmit').forForm(form).onFormSubmit().create();
    return typeof trigger.getUniqueId === 'function' ? trigger.getUniqueId() : '';
}
function buildApprovalFormRequestIdPropKey(formId) {
    return `${APPROVAL_FORM_REQUEST_ID_PROP_PREFIX}${formId}`;
}
function findRequestByApprovalFormId(formId) {
    const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.REQUESTS);
    if (!sheet)
        return null;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1)
        return null;
    const headerMap = getHeaderMap(data[0]);
    const formIdIndex = headerMap['承認フォームID'];
    const requestIdIndex = headerMap['requestId'];
    if (!formIdIndex || !requestIdIndex)
        return null;
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (getCellValue(row, formIdIndex) !== formId)
            continue;
        return { requestId: getCellValue(row, requestIdIndex), rowIndex: i + 1 };
    }
    const requestId = PropertiesService.getDocumentProperties().getProperty(buildApprovalFormRequestIdPropKey(formId)) || '';
    if (!requestId)
        return null;
    return { requestId, rowIndex: 0 };
}
function closeApprovalFormByRequestRow(row, headerMap) {
    const formId = getCellValue(row, headerMap['承認フォームID']);
    if (!formId)
        return;
    try {
        const form = FormApp.openById(formId);
        form.setAcceptingResponses(false);
    }
    catch (err) {
        Logger.log(`closeApprovalFormByRequestRow: ${formId} ${err}`);
    }
}
function applyFormPublicSettings(form) {
    form.setRequireLogin(false);
    form.setCollectEmail(false);
    form.setLimitOneResponsePerUser(false);
    form.setShowLinkToRespondAgain(false);
}
function ensureFormExplanationHeader(form) {
    const items = form.getItems(FormApp.ItemType.SECTION_HEADER);
    const existing = items.find((item) => item.getTitle() === 'ご回答方法');
    if (existing) {
        moveItemToTopByTypeAndTitle(form, FormApp.ItemType.SECTION_HEADER, 'ご回答方法');
        return;
    }
    const header = form.addSectionHeaderItem();
    header.setTitle('ご回答方法');
    header.setHelpText([
        '1) 「更新方針（車両ごと）」は必須です。',
        `2) 回答は「${ANSWER_OPTIONS.join(' / ')}」から選択してください。`,
        '3) 車両の並びは、通知メールの一覧と同じ順です。',
        '※このフォームのURLは転送しないでください。',
    ].join('\n'));
    moveLastItemToTop(form);
}
function moveItemToTopByTypeAndTitle(form, itemType, title) {
    try {
        const all = form.getItems();
        const idx = all.findIndex((item) => item.getType() === itemType && item.getTitle() === title);
        if (idx > 0)
            form.moveItem(idx, 0);
    }
    catch (err) {
        // move に失敗しても致命ではない
    }
}
function moveLastItemToTop(form) {
    try {
        const all = form.getItems();
        if (all.length > 1) {
            form.moveItem(all.length - 1, 0);
        }
    }
    catch (err) {
        // move に失敗しても致命ではない
    }
}
function normalizeExistingFormsForRequest(params) {
    if (!params.formIds || params.formIds.length === 0)
        return;
    const props = PropertiesService.getDocumentProperties();
    let cursor = 0;
    params.formIds.forEach((formId, formIndex) => {
        try {
            const form = FormApp.openById(formId);
            applyFormPublicSettings(form);
            ensureFormExplanationHeader(form);
            form.setDescription(buildFormDescription(params, formIndex, params.formIds.length, params.tz));
            const gridItem = findPolicyGridItem(form);
            if (!gridItem)
                return;
            const rowCount = gridItem.getRows().length;
            if (rowCount <= 0)
                return;
            const sliceVehicles = params.vehicles.slice(cursor, cursor + rowCount);
            cursor += rowCount;
            if (sliceVehicles.length > 0) {
                gridItem.setRows(sliceVehicles.map((row, rowIndex) => buildFormVehicleRowLabel(row, params.vehicleHeader, params.tz, rowIndex)));
                gridItem.setColumns(ANSWER_OPTIONS);
                gridItem.setRequired(true);
                const vehicleIds = sliceVehicles.map((row) => getCellValue(row, params.vehicleHeader['vehicleId']) || '');
                props.setProperty(buildFormVehicleIdsPropKey(formId), JSON.stringify(vehicleIds));
            }
        }
        catch (err) {
            Logger.log(`normalizeExistingFormsForRequest: ${formId} ${err}`);
        }
    });
}
function findPolicyGridItem(form) {
    const items = form.getItems(FormApp.ItemType.GRID);
    for (const item of items) {
        if (item.getTitle() === FORM_ITEM_TITLES.POLICY_GRID) {
            try {
                return item.asGridItem();
            }
            catch (err) {
                return null;
            }
        }
    }
    return null;
}
function buildFormTitle(dept, requestId, index, total) {
    let title = `【車両更新方針】${dept} ${requestId}`;
    if (total > 1) {
        title += ` Part${index + 1}/${total}`;
    }
    return title;
}
function buildFormDescription(params, index, total, tz) {
    const startLabel = params.targetStart ? formatDateLabel(params.targetStart, tz) : formatDateLabel(new Date(), tz);
    const endLabel = params.targetEnd ? formatDateLabel(params.targetEnd, tz) : formatDateLabel(new Date(), tz);
    const deadlineLabel = params.deadline ? formatDateLabel(params.deadline, tz) : formatDateLabel(new Date(), tz);
    const lines = [
        total > 1 ? `Part${index + 1}/${total}` : '',
        `対象期間: ${startLabel}〜${endLabel}`,
        `締切: ${deadlineLabel}`,
        `選択肢: ${ANSWER_OPTIONS.join(' / ')}`,
        '※このフォームのURLは転送しないでください。',
    ].filter((line) => line);
    return lines.join('\n');
}
function buildFormVehicleRowLabel(row, headerMap, tz, rowIndex) {
    const display = formatFormVehicleLineShort(row, headerMap, tz, rowIndex);
    return display;
}
function formatFormVehicleLineShort(row, headerMap, tz, rowIndex) {
    const reg = getCellValue(row, headerMap['登録番号_結合']);
    const type = getCellValue(row, headerMap['車種']);
    const chassis = getCellValue(row, headerMap['車台番号']);
    const dept = getCellValue(row, headerMap['管理部門']);
    const manager = getCellValue(row, headerMap['管理担当者']);
    const start = parseDateValue(getCellRaw(row, headerMap['契約開始日']));
    const end = parseDateValue(getCellRaw(row, headerMap['契約満了日']));
    const contractTerm = getCellValue(row, headerMap['契約期間']);
    const inspectionEnd = parseDateValue(getCellRaw(row, headerMap['車検満了日']));
    const leaseFee = getCellValue(row, headerMap['リース料（税抜）']);
    const startLabel = start ? formatDateIsoLabel(start, tz) : '未設定';
    const endLabel = end ? formatDateIsoLabel(end, tz) : '未設定';
    const inspectionLabel = inspectionEnd ? formatDateIsoLabel(inspectionEnd, tz) : '未設定';
    const numberLabel = reg || `車両${rowIndex + 1}`;
    const typeLabel = type || '車種未設定';
    const chassisLabel = chassis || '車台番号未設定';
    const managerLabel = manager || '-';
    const deptLabel = dept || '-';
    const termLabel = contractTerm || '-';
    const leaseFeeLabel = leaseFee || '-';
    return [
        `【${rowIndex + 1}】登録番号:${numberLabel}`,
        `車種:${typeLabel}`,
        `車台番号:${chassisLabel}`,
        `管理部門:${deptLabel}`,
        `管理担当者:${managerLabel}`,
        `契約開始日:${startLabel}`,
        `契約満了日:${endLabel}`,
        `契約期間:${termLabel}`,
        `車検満了日:${inspectionLabel}`,
        `リース料(税抜):${leaseFeeLabel}`,
    ].join(' / ');
}
function getFormIdFromEvent(e) {
    try {
        if (e && e.source && typeof e.source.getId === 'function') {
            return e.source.getId();
        }
    }
    catch (err) {
        Logger.log(`getFormIdFromEvent: ${err}`);
    }
    try {
        const response = e && e.response;
        if (response && typeof response.getFormId === 'function') {
            return response.getFormId();
        }
    }
    catch (err) {
        Logger.log(`getFormIdFromEvent response: ${err}`);
    }
    return '';
}
function findRequestByFormId(formId) {
    const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.REQUESTS);
    if (!sheet)
        return null;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1)
        return null;
    const headerMap = getHeaderMap(data[0]);
    const requestIdIndex = headerMap['requestId'];
    const formIdIndex = headerMap['formId'];
    const formIdsJsonIndex = headerMap['formIdsJson'];
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (formIdIndex) {
            const rowFormId = getCellValue(row, formIdIndex);
            if (rowFormId && rowFormId === formId) {
                return { requestId: getCellValue(row, requestIdIndex), rowIndex: i + 1 };
            }
        }
        if (formIdsJsonIndex) {
            const ids = parseJsonStringArray(getCellValue(row, formIdsJsonIndex));
            if (ids.indexOf(formId) >= 0) {
                return { requestId: getCellValue(row, requestIdIndex), rowIndex: i + 1 };
            }
        }
    }
    return null;
}
function extractAnswersFromFormResponse(formId, response) {
    const result = {
        answersByVehicleId: {},
    };
    const vehicleIdsForForm = loadVehicleIdsForForm(formId);
    const itemResponses = response.getItemResponses();
    itemResponses.forEach((itemResponse) => {
        const item = itemResponse.getItem();
        const type = item.getType();
        if (type === FormApp.ItemType.GRID) {
            const gridAnswers = extractAnswersFromGridItem(item, itemResponse, vehicleIdsForForm);
            Object.keys(gridAnswers).forEach((vehicleId) => {
                result.answersByVehicleId[vehicleId] = gridAnswers[vehicleId];
            });
        }
    });
    return result;
}
function extractAnswersFromGridItem(item, itemResponse, vehicleIdsForForm) {
    const answers = {};
    let rows = [];
    try {
        rows = item.asGridItem().getRows();
    }
    catch (err) {
        return answers;
    }
    const response = itemResponse.getResponse();
    if (Array.isArray(response)) {
        rows.forEach((rowLabel, index) => {
            const answer = response[index];
            const vehicleId = (vehicleIdsForForm && vehicleIdsForForm[index]) ? vehicleIdsForForm[index] : extractVehicleIdFromRowLabel(rowLabel);
            if (vehicleId && answer) {
                answers[vehicleId] = Array.isArray(answer) ? String(answer[0] || '') : String(answer);
            }
        });
        return answers;
    }
    if (response && typeof response === 'object') {
        Object.keys(response).forEach((rowLabel) => {
            const answer = response[rowLabel];
            const rowIndex = rows.indexOf(rowLabel);
            const vehicleId = rowIndex >= 0 && vehicleIdsForForm && vehicleIdsForForm[rowIndex]
                ? vehicleIdsForForm[rowIndex]
                : extractVehicleIdFromRowLabel(rowLabel);
            if (vehicleId && answer) {
                answers[vehicleId] = Array.isArray(answer) ? String(answer[0] || '') : String(answer);
            }
        });
        return answers;
    }
    return answers;
}
function extractVehicleIdFromRowLabel(label) {
    const match = String(label || '').match(/\|([^|]+)\|/);
    return match ? match[1].trim() : '';
}
function parseVehicleComments(commentText, vehicleIds) {
    const map = {};
    if (!commentText)
        return map;
    const lines = String(commentText)
        .split(/\r?\n/)
        .map((line) => line.trim())
        .filter((line) => line);
    lines.forEach((line) => {
        const match = line.match(/^([^:：]+)[:：]\s*(.+)$/);
        if (!match)
            return;
        const key = match[1].trim();
        const num = key.match(/^\d+$/) ? Number(key) : 0;
        const vehicleId = num > 0 && num <= vehicleIds.length ? vehicleIds[num - 1] : key;
        if (vehicleIds.indexOf(vehicleId) === -1)
            return;
        const comment = match[2].trim();
        if (comment)
            map[vehicleId] = comment;
    });
    return map;
}
function buildFormVehicleIdsPropKey(formId) {
    return `${FORM_VEHICLE_IDS_PROP_PREFIX}${formId}`;
}
function loadVehicleIdsForForm(formId) {
    const raw = PropertiesService.getDocumentProperties().getProperty(buildFormVehicleIdsPropKey(formId));
    if (!raw)
        return [];
    return parseJsonStringArray(raw);
}
function parseJsonStringArray(value) {
    if (value === null || value === undefined || value === '')
        return [];
    if (Array.isArray(value))
        return value.map((v) => String(v));
    try {
        const parsed = JSON.parse(String(value));
        if (Array.isArray(parsed))
            return parsed.map((v) => String(v));
    }
    catch (err) {
        return [];
    }
    return [];
}
function chunkArray(items, size) {
    const result = [];
    if (!items || items.length === 0)
        return result;
    const chunkSize = Math.max(1, Math.floor(size));
    for (let i = 0; i < items.length; i += chunkSize) {
        result.push(items.slice(i, i + chunkSize));
    }
    return result;
}
function closeRequestForms(requestId) {
    const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.REQUESTS);
    if (!sheet)
        return;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1)
        return;
    const headerMap = getHeaderMap(data[0]);
    const requestIdIndex = headerMap['requestId'];
    const formIdIndex = headerMap['formId'];
    const formIdsJsonIndex = headerMap['formIdsJson'];
    let formIds = [];
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (getCellValue(row, requestIdIndex) !== requestId)
            continue;
        if (formIdIndex) {
            const formId = getCellValue(row, formIdIndex);
            if (formId)
                formIds.push(String(formId));
        }
        if (formIdsJsonIndex) {
            formIds = formIds.concat(parseJsonStringArray(getCellValue(row, formIdsJsonIndex)));
        }
        break;
    }
    formIds = Array.from(new Set(formIds.filter((id) => id)));
    if (formIds.length === 0)
        return;
    const props = PropertiesService.getDocumentProperties();
    formIds.forEach((formId) => {
        try {
            const form = FormApp.openById(formId);
            form.setAcceptingResponses(false);
        }
        catch (err) {
            Logger.log(`closeRequestForms: ${formId} ${err}`);
        }
        try {
            props.deleteProperty(buildFormVehicleIdsPropKey(formId));
        }
        catch (err) {
            Logger.log(`closeRequestForms deleteProperty: ${formId} ${err}`);
        }
    });
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach((trigger) => {
        if (trigger.getHandlerFunction() !== 'onRequestFormSubmit')
            return;
        const sourceId = typeof trigger.getTriggerSourceId === 'function' ? trigger.getTriggerSourceId() : '';
        if (sourceId && formIds.indexOf(sourceId) >= 0) {
            ScriptApp.deleteTrigger(trigger);
        }
    });
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
function appendNotificationLog(type, dept, to, requestId, result) {
    const ss = getSpreadsheet();
    const sheet = ensureSheet(ss, SHEET_NAMES.NOTIFY_LOG);
    ensureHeaders(sheet, 1, getSchemaHeaders(SHEET_NAMES.NOTIFY_LOG));
    sheet.appendRow([new Date(), type, dept, to, requestId, result]);
}
function validateRequestAccess(params) {
    const requestId = params['requestId'] ? String(params['requestId']) : '';
    const token = params['token'] ? String(params['token']) : '';
    const dept = params['dept'] ? String(params['dept']) : '';
    const deptToken = params['deptToken'] ? String(params['deptToken']) : '';
    if (!requestId || !token || !dept || !deptToken) {
        return { ok: false, message: 'パラメータが不足しています。', requestId: '', requestToken: '' };
    }
    const requestRow = findRequestRow(requestId);
    if (!requestRow) {
        return { ok: false, message: '依頼が見つかりません。', requestId: '', requestToken: '' };
    }
    if (requestRow.requestToken !== token) {
        return { ok: false, message: 'トークンが一致しません。', requestId: '', requestToken: '' };
    }
    if (requestRow.dept !== dept) {
        return { ok: false, message: '管理部門が一致しません。', requestId: '', requestToken: '' };
    }
    const deptMaster = loadDeptMaster();
    const deptInfo = deptMaster[dept];
    if (!deptInfo || deptInfo.token !== deptToken) {
        return { ok: false, message: '部署トークンが一致しません。', requestId: '', requestToken: '' };
    }
    if (requestRow.status === REQUEST_STATUS.COMPLETED) {
        return { ok: false, message: 'この依頼は完了しています。', requestId: '', requestToken: '' };
    }
    return {
        ok: true,
        requestId,
        requestToken: token,
        requestRow: {
            dept,
            status: requestRow.status,
            deptToken,
        },
    };
}
function findRequestRow(requestId) {
    const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.REQUESTS);
    if (!sheet)
        return null;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1)
        return null;
    const headerMap = getHeaderMap(data[0]);
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (getCellValue(row, headerMap['requestId']) === requestId) {
            return {
                rowIndex: i + 1,
                requestToken: getCellValue(row, headerMap['requestToken']),
                dept: getCellValue(row, headerMap['管理部門']),
                status: getCellValue(row, headerMap['ステータス']),
            };
        }
    }
    return null;
}
function getVehiclesByRequestId(requestId) {
    const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.VEHICLE_VIEW);
    if (!sheet)
        return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1)
        return [];
    const headerMap = getHeaderMap(data[0]);
    const vehicles = [];
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (getCellValue(row, headerMap['依頼ID']) !== requestId)
            continue;
        const contractEnd = parseDateValue(getCellRaw(row, headerMap['契約満了日']));
        vehicles.push({
            vehicleId: getCellValue(row, headerMap['vehicleId']),
            reg: getCellValue(row, headerMap['登録番号_結合']),
            type: getCellValue(row, headerMap['車種']),
            contractEnd: contractEnd ? formatDateLabel(contractEnd, getSpreadsheet().getSpreadsheetTimeZone()) : '未設定',
        });
    }
    return vehicles;
}
function loadAnswersForRequest(requestId) {
    const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.ANSWERS);
    if (!sheet)
        return {};
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1)
        return {};
    const headerMap = getHeaderMap(data[0]);
    const map = {};
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (getCellValue(row, headerMap['requestId']) !== requestId)
            continue;
        const vehicleId = getCellValue(row, headerMap['vehicleId']);
        map[vehicleId] = {
            vehicleId,
            requestId,
            answer: getCellValue(row, headerMap['回答']),
            comment: getCellValue(row, headerMap['コメント']),
            answeredAt: parseDateValue(getCellRaw(row, headerMap['回答日時'])) || new Date(),
        };
    }
    return map;
}
function buildAnswerRowHtml(vehicle, answer, index) {
    const optionsHtml = ANSWER_OPTIONS.map((option) => {
        const selected = answer && answer.answer === option ? 'selected' : '';
        return `<option value="${escapeHtml(option)}" ${selected}>${escapeHtml(option)}</option>`;
    }).join('');
    const comment = answer ? escapeHtml(answer.comment) : '';
    return `
    <tr>
      <td>${index + 1}</td>
      <td>${escapeHtml(vehicle.reg || '')}<br><span class="muted">${escapeHtml(vehicle.type || '')}</span></td>
      <td>${escapeHtml(vehicle.contractEnd || '')}</td>
      <td>
        <input type="hidden" name="vehicleId" value="${escapeHtml(vehicle.vehicleId)}">
        <select name="answer">
          <option value="">--</option>
          ${optionsHtml}
        </select>
      </td>
      <td><input type="text" name="comment" value="${comment}" style="width: 100%;"></td>
    </tr>
  `;
}
function ensureArray(value) {
    if (value === undefined || value === null)
        return [];
    if (Array.isArray(value))
        return value.map((v) => String(v));
    return [String(value)];
}
function upsertAnswers(inputs) {
    const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.ANSWERS);
    if (!sheet)
        return;
    const data = sheet.getDataRange().getValues();
    if (data.length === 0)
        return;
    const headerMap = getHeaderMap(data[0]);
    const keyToIndex = {};
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const key = `${getCellValue(row, headerMap['requestId'])}__${getCellValue(row, headerMap['vehicleId'])}`;
        keyToIndex[key] = i;
    }
    inputs.forEach((input) => {
        const key = `${input.requestId}__${input.vehicleId}`;
        const row = [input.requestId, input.vehicleId, input.answer, input.comment, input.responder, input.answeredAt];
        if (keyToIndex[key]) {
            data[keyToIndex[key]] = row;
        }
        else {
            data.push(row);
        }
    });
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}
function updateRequestStatus(requestId) {
    const ss = getSpreadsheet();
    const requestSheet = ss.getSheetByName(SHEET_NAMES.REQUESTS);
    const vehicleSheet = ss.getSheetByName(SHEET_NAMES.VEHICLE_VIEW);
    if (!requestSheet || !vehicleSheet)
        return '';
    const requestData = requestSheet.getDataRange().getValues();
    if (requestData.length <= 1)
        return '';
    const requestHeader = getHeaderMap(requestData[0]);
    const vehicleData = vehicleSheet.getDataRange().getValues();
    const vehicleHeader = getHeaderMap(vehicleData[0]);
    const vehicles = vehicleData
        .slice(1)
        .filter((v) => getCellValue(v, vehicleHeader['依頼ID']) === requestId);
    const total = vehicles.length;
    const answered = vehicles.filter((v) => getCellValue(v, vehicleHeader['更新方針'])).length;
    let newStatus = '';
    if (total > 0 && answered >= total) {
        newStatus = REQUEST_STATUS.COMPLETED;
    }
    else if (answered > 0) {
        newStatus = REQUEST_STATUS.RESPONDING;
    }
    else {
        newStatus = REQUEST_STATUS.SENT;
    }
    for (let i = 1; i < requestData.length; i++) {
        const row = requestData[i];
        if (getCellValue(row, requestHeader['requestId']) !== requestId)
            continue;
        row[requestHeader['ステータス'] - 1] = newStatus;
        if (newStatus === REQUEST_STATUS.COMPLETED) {
            row[requestHeader['requestToken'] - 1] = '';
        }
        if (newStatus === REQUEST_STATUS.COMPLETED && requestHeader['承認ステータス']) {
            const currentApproval = getCellValue(row, requestHeader['承認ステータス']);
            if (!currentApproval ||
                currentApproval === APPROVAL_STATUS.NOT_SENT ||
                currentApproval === APPROVAL_STATUS.RETURNED) {
                row[requestHeader['承認ステータス'] - 1] = APPROVAL_STATUS.PENDING;
            }
        }
    }
    requestSheet.getRange(1, 1, requestData.length, requestData[0].length).setValues(requestData);
    return newStatus;
}
function ensureAppendColumns(sheet, headers) {
    const headerRow = 1;
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
