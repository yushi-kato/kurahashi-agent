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
    NOTIFY_BATCHES: '通知バッチ',
    TEST_RESULTS: 'テスト結果',
};
const VEHICLE_SHEET_NAME = '車両一覧';
const SOURCE_SHEETS = [VEHICLE_SHEET_NAME];
const CONFIRM_SHEET_PREFIX = '本部長副本部長確認_';
const REQUEST_STATUS = {
    CREATED: '作成済',
    SENT: '送信済',
    RESPONDING: '回答中',
    COMPLETED: '完了',
    EXPIRED: '締切',
};
const BATCH_STATUS = {
    CREATED: '作成済',
    INITIAL_SENT: '初回送信済',
    REMINDED: 'リマインド送信済',
    SENMU_REQUESTED: '専務依頼送信済',
    RETURNED: '差戻しあり',
    APPLIED: '反映済',
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
const HQ_RESPONSE_OPTIONS = [ANSWER_LABELS.RENEW, ANSWER_LABELS.CANCELLATION_REPLACE, ANSWER_LABELS.CANCELLATION_END];
const SENMU_DECISION = {
    APPROVE: '承認',
    RETURN: '差戻し',
};
const CONFIRM_SHEET_HEADERS = [
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
const VEHICLE_MASTER_EXTRA_HEADERS = ['マスター反映済み', '反映日時'];
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
        name: SHEET_NAMES.NOTIFY_LOG,
        headerRow: 1,
        headers: ['日時', '種別', '管理部門', '宛先', 'requestId', '結果'],
    },
    {
        name: SHEET_NAMES.NOTIFY_BATCHES,
        headerRow: 1,
        headers: [
            'batchId',
            '送付日',
            '期限日',
            '対象開始日',
            '対象終了日',
            '確認シート名',
            'ステータス',
            '初回送信日時',
            'リマインド送信日時',
            '専務依頼送信日時',
            '反映完了日時',
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
    本部長副本部長_通知先To: '',
    専務_通知先To: '',
    専務_通知先Cc: '',
    半期送付日_3月: '03-01',
    半期送付日_9月: '09-01',
    回答期限_3月: '03-31',
    回答期限_9月: '09-30',
    リマインド_期限前日数: 10,
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
        .addItem('半期バッチ作成', 'createRequests')
        .addItem('本部長/副本部長 初回通知送信', 'sendInitialEmails')
        .addItem('本部長/副本部長 リマインド送信', 'sendReminderEmails')
        .addItem('専務 承認依頼送信', 'sendApprovalRequestEmails')
        .addItem('専務判断反映・マスター更新', 'applyApprovalDecisions')
        .addSeparator()
        .addItem('設定ひな形作成', 'seedSettings')
        .addItem('テスト車両追加', 'seedTestVehicles')
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
        const vehicleSheet = ss.getSheetByName(VEHICLE_SHEET_NAME);
        if (!vehicleSheet)
            throw new Error('車両一覧が存在しません');
        ensureAppendColumns(vehicleSheet, VEHICLE_MASTER_EXTRA_HEADERS);
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
        const settings = loadSettings();
        const tz = ss.getSpreadsheetTimeZone();
        const schedule = resolveSemiannualSchedule(new Date(), settings, tz);
        if (!schedule) {
            appendNotificationLog('バッチ作成', '', '', '', '送付日ではないためスキップ');
            return;
        }
        ensureHeaders(ensureSheet(ss, SHEET_NAMES.NOTIFY_BATCHES), 1, getSchemaHeaders(SHEET_NAMES.NOTIFY_BATCHES));
        const batchSheet = ss.getSheetByName(SHEET_NAMES.NOTIFY_BATCHES);
        if (!batchSheet)
            throw new Error('通知バッチシートが存在しません');
        const batchData = batchSheet.getDataRange().getValues();
        const batchHeader = batchData.length > 0 ? getHeaderMap(batchData[0]) : {};
        const batchId = Utilities.formatDate(schedule.sendDate, tz, 'yyyyMMdd');
        if (batchHeader['batchId']) {
            for (let i = 1; i < batchData.length; i++) {
                if (getCellValue(batchData[i], batchHeader['batchId']) === batchId) {
                    appendNotificationLog('バッチ作成', '', '', '', `既存バッチ(${batchId})があるためスキップ`);
                    return;
                }
            }
        }
        const vehicleSheet = ss.getSheetByName(VEHICLE_SHEET_NAME);
        if (!vehicleSheet)
            throw new Error('車両一覧が存在しません');
        ensureAppendColumns(vehicleSheet, VEHICLE_MASTER_EXTRA_HEADERS);
        const vehicleData = vehicleSheet.getDataRange().getValues();
        if (vehicleData.length <= 1) {
            appendNotificationLog('バッチ作成', '', '', '', '車両一覧が空のためスキップ');
            return;
        }
        const headerMap = getHeaderMap(vehicleData[0]);
        const sourceHeader = resolveSourceHeaders(headerMap);
        const needsInputRows = [];
        const confirmRows = [];
        const now = new Date();
        const rangeStart = schedule.rangeStart;
        const rangeEnd = schedule.rangeEnd;
        for (let i = 1; i < vehicleData.length; i++) {
            const row = vehicleData[i];
            const dept = getCellValue(row, sourceHeader.dept);
            const manager = getCellValue(row, sourceHeader.manager);
            const contractEnd = parseDateValue(getCellRaw(row, sourceHeader.contractEnd));
            const regCombined = buildRegistrationCombined(getCellValue(row, sourceHeader.regArea), getCellValue(row, sourceHeader.regClass), getCellValue(row, sourceHeader.regKana), getCellValue(row, sourceHeader.regNumber));
            const regAll = sourceHeader.regAll ? getCellValue(row, sourceHeader.regAll) : '';
            const regLabel = regCombined || regAll;
            const vehicleType = getCellValue(row, sourceHeader.vehicleType);
            const chassis = getCellValue(row, sourceHeader.chassis);
            if (!contractEnd) {
                needsInputRows.push([now, VEHICLE_SHEET_NAME, `row:${i + 1}`, dept, regLabel, vehicleType, '契約満了日が未設定']);
                continue;
            }
            const contractDate = toDateOnly(contractEnd, tz);
            const masterApplied = headerMap['マスター反映済み']
                ? toBoolean(getCellRaw(row, headerMap['マスター反映済み']), false)
                : false;
            const inRange = isWithinRange(contractDate, rangeStart, rangeEnd);
            const missed = contractDate.getTime() < rangeStart.getTime() && !masterApplied;
            if (!inRange && !missed)
                continue;
            confirmRows.push([
                dept,
                manager,
                regLabel,
                vehicleType,
                chassis,
                getCellRaw(row, sourceHeader.contractStart),
                contractEnd,
                getCellValue(row, sourceHeader.contractTerm),
                getCellRaw(row, sourceHeader.inspectionEnd),
                getCellRaw(row, sourceHeader.leaseFee),
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
        }
        if (needsInputRows.length > 0) {
            const needsSheet = ensureSheet(ss, SHEET_NAMES.NEEDS_INPUT);
            ensureHeaders(needsSheet, 1, getSchemaHeaders(SHEET_NAMES.NEEDS_INPUT));
            const startRow = needsSheet.getLastRow() + 1;
            needsSheet.getRange(startRow, 1, needsInputRows.length, needsInputRows[0].length).setValues(needsInputRows);
        }
        if (confirmRows.length === 0) {
            appendNotificationLog('バッチ作成', '', '', '', '対象車両なし');
            return;
        }
        const confirmSheetName = `${CONFIRM_SHEET_PREFIX}${batchId}`;
        const confirmSheet = ensureConfirmSheet(ss, confirmSheetName, settings);
        if (confirmSheet.getLastRow() > 1) {
            appendNotificationLog('バッチ作成', '', '', '', `確認シート(${confirmSheetName})に既存データあり`);
            return;
        }
        confirmSheet.getRange(2, 1, confirmRows.length, confirmRows[0].length).setValues(confirmRows);
        applyConfirmSheetValidations(confirmSheet, confirmRows.length);
        const batchRow = new Array(batchSheet.getLastColumn()).fill('');
        const setBatchCell = (headerName, value) => {
            const idx = batchHeader[headerName];
            if (idx)
                batchRow[idx - 1] = value;
        };
        setBatchCell('batchId', batchId);
        setBatchCell('送付日', schedule.sendDate);
        setBatchCell('期限日', schedule.deadline);
        setBatchCell('対象開始日', rangeStart);
        setBatchCell('対象終了日', rangeEnd);
        setBatchCell('確認シート名', confirmSheetName);
        setBatchCell('ステータス', BATCH_STATUS.CREATED);
        setBatchCell('初回送信日時', '');
        setBatchCell('リマインド送信日時', '');
        setBatchCell('専務依頼送信日時', '');
        setBatchCell('反映完了日時', '');
        const startRow = batchSheet.getLastRow() + 1;
        batchSheet.getRange(startRow, 1, 1, batchRow.length).setValues([batchRow]);
        appendNotificationLog('バッチ作成', '', '', batchId, `対象車両=${confirmRows.length}`);
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
            appendNotificationLog('初回通知', '', '', '', '通知_メール送信=FALSE のため送信をスキップ');
            return;
        }
        const batchSheet = ss.getSheetByName(SHEET_NAMES.NOTIFY_BATCHES);
        if (!batchSheet)
            throw new Error('通知バッチシートが存在しません');
        const batchData = batchSheet.getDataRange().getValues();
        if (batchData.length <= 1)
            return;
        const batchHeader = getHeaderMap(batchData[0]);
        const tz = ss.getSpreadsheetTimeZone();
        const now = new Date();
        const hqLeadersTo = splitEmails(settings.hqLeadersTo);
        if (hqLeadersTo.length === 0) {
            appendNotificationLog('初回通知', '', '', '', '本部長副本部長_通知先Toが未設定');
            return;
        }
        for (let i = 1; i < batchData.length; i++) {
            const row = batchData[i];
            const sentAt = parseDateValue(getCellRaw(row, batchHeader['初回送信日時']));
            if (sentAt)
                continue;
            const batchId = getCellValue(row, batchHeader['batchId']);
            const sheetName = getCellValue(row, batchHeader['確認シート名']);
            const rangeStart = parseDateValue(getCellRaw(row, batchHeader['対象開始日']));
            const rangeEnd = parseDateValue(getCellRaw(row, batchHeader['対象終了日']));
            const deadline = parseDateValue(getCellRaw(row, batchHeader['期限日']));
            if (!sheetName) {
                appendNotificationLog('初回通知', '', '', batchId, '確認シート名が未設定');
                continue;
            }
            const confirmSheet = ss.getSheetByName(sheetName);
            if (!confirmSheet) {
                appendNotificationLog('初回通知', '', '', batchId, `確認シート(${sheetName})が存在しません`);
                continue;
            }
            const confirmData = confirmSheet.getDataRange().getValues();
            if (confirmData.length <= 1) {
                appendNotificationLog('初回通知', '', '', batchId, '対象車両なし');
                continue;
            }
            const confirmHeader = getHeaderMap(confirmData[0]);
            const listText = confirmData
                .slice(1)
                .map((v) => formatConfirmVehicleLine(v, confirmHeader, tz))
                .join('\n');
            const listHtml = confirmData
                .slice(1)
                .map((v) => `<li>${escapeHtml(formatConfirmVehicleLine(v, confirmHeader, tz))}</li>`)
                .join('');
            const sheetUrl = buildSheetUrlWithGid(ss, confirmSheet);
            const subject = `【車両更新確認】半期一括 ${formatDateLabel(rangeStart || now, tz)}〜${formatDateLabel(rangeEnd || now, tz)}`;
            const bodyText = [
                '本部長・副本部長 各位',
                '',
                `対象期間: ${formatDateLabel(rangeStart || now, tz)}〜${formatDateLabel(rangeEnd || now, tz)}`,
                `回答期限: ${formatDateLabel(deadline || now, tz)}`,
                '',
                `確認シート: ${sheetUrl}`,
                '',
                '対象車両:',
                listText,
            ].join('\n');
            const bodyHtml = [
                '<p>本部長・副本部長 各位</p>',
                `<p>対象期間: ${escapeHtml(formatDateLabel(rangeStart || now, tz))}〜${escapeHtml(formatDateLabel(rangeEnd || now, tz))}<br>回答期限: ${escapeHtml(formatDateLabel(deadline || now, tz))}</p>`,
                `<p>確認シート: <a href="${sheetUrl}">${sheetUrl}</a></p>`,
                `<p>対象車両:</p><ul>${listHtml}</ul>`,
            ].join('');
            try {
                MailApp.sendEmail({
                    to: hqLeadersTo.join(','),
                    subject,
                    body: bodyText,
                    htmlBody: bodyHtml,
                    name: settings.fromName,
                });
                row[batchHeader['初回送信日時'] - 1] = now;
                row[batchHeader['ステータス'] - 1] = BATCH_STATUS.INITIAL_SENT;
                appendNotificationLog('初回通知', '', hqLeadersTo.join(','), batchId, '送信OK');
            }
            catch (err) {
                appendNotificationLog('初回通知', '', hqLeadersTo.join(','), batchId, `送信失敗: ${err}`);
            }
        }
        batchSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);
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
        const batchSheet = ss.getSheetByName(SHEET_NAMES.NOTIFY_BATCHES);
        if (!batchSheet)
            throw new Error('通知バッチシートが存在しません');
        const batchData = batchSheet.getDataRange().getValues();
        if (batchData.length <= 1)
            return;
        const batchHeader = getHeaderMap(batchData[0]);
        const tz = ss.getSpreadsheetTimeZone();
        const today = toDateOnly(new Date(), tz);
        const hqLeadersTo = splitEmails(settings.hqLeadersTo);
        if (hqLeadersTo.length === 0) {
            appendNotificationLog('リマインド', '', '', '', '本部長副本部長_通知先Toが未設定');
            return;
        }
        for (let i = 1; i < batchData.length; i++) {
            const row = batchData[i];
            const initialSentAt = parseDateValue(getCellRaw(row, batchHeader['初回送信日時']));
            if (!initialSentAt)
                continue;
            const reminderSentAt = parseDateValue(getCellRaw(row, batchHeader['リマインド送信日時']));
            if (reminderSentAt)
                continue;
            const batchId = getCellValue(row, batchHeader['batchId']);
            const sheetName = getCellValue(row, batchHeader['確認シート名']);
            const deadline = parseDateValue(getCellRaw(row, batchHeader['期限日']));
            if (!sheetName || !deadline)
                continue;
            const reminderDate = addDays(toDateOnly(deadline, tz), -settings.reminderDaysBeforeDeadline);
            if (today.getTime() !== reminderDate.getTime())
                continue;
            const confirmSheet = ss.getSheetByName(sheetName);
            if (!confirmSheet) {
                appendNotificationLog('リマインド', '', '', batchId, `確認シート(${sheetName})が存在しません`);
                continue;
            }
            const confirmData = confirmSheet.getDataRange().getValues();
            if (confirmData.length <= 1)
                continue;
            const confirmHeader = getHeaderMap(confirmData[0]);
            const uncheckedRows = confirmData.slice(1).filter((v) => !toBoolean(getCellRaw(v, confirmHeader['回答確認済み']), false));
            if (uncheckedRows.length === 0)
                continue;
            const listText = uncheckedRows
                .map((v) => formatConfirmVehicleLine(v, confirmHeader, tz))
                .join('\n');
            const listHtml = uncheckedRows
                .map((v) => `<li>${escapeHtml(formatConfirmVehicleLine(v, confirmHeader, tz))}</li>`)
                .join('');
            const sheetUrl = buildSheetUrlWithGid(ss, confirmSheet);
            const subject = `【車両更新確認】リマインド ${formatDateLabel(deadline, tz)}`;
            const bodyText = [
                '本部長・副本部長 各位',
                '',
                `回答期限: ${formatDateLabel(deadline, tz)}`,
                `確認シート: ${sheetUrl}`,
                '',
                '未確認車両:',
                listText,
            ].join('\n');
            const bodyHtml = [
                '<p>本部長・副本部長 各位</p>',
                `<p>回答期限: ${escapeHtml(formatDateLabel(deadline, tz))}</p>`,
                `<p>確認シート: <a href="${sheetUrl}">${sheetUrl}</a></p>`,
                `<p>未確認車両:</p><ul>${listHtml}</ul>`,
            ].join('');
            try {
                MailApp.sendEmail({
                    to: hqLeadersTo.join(','),
                    subject,
                    body: bodyText,
                    htmlBody: bodyHtml,
                    name: settings.fromName,
                });
                row[batchHeader['リマインド送信日時'] - 1] = new Date();
                row[batchHeader['ステータス'] - 1] = BATCH_STATUS.REMINDED;
                appendNotificationLog('リマインド', '', hqLeadersTo.join(','), batchId, '送信OK');
            }
            catch (err) {
                appendNotificationLog('リマインド', '', hqLeadersTo.join(','), batchId, `送信失敗: ${err}`);
            }
        }
        batchSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);
    }
    finally {
        lock.releaseLock();
    }
}
function sendApprovalRequestEmails() {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        const ss = getSpreadsheet();
        const settings = loadSettings();
        if (!settings.mailSendEnabled) {
            appendNotificationLog('専務依頼', '', '', '', '通知_メール送信=FALSE のため送信をスキップ');
            return;
        }
        const batchSheet = ss.getSheetByName(SHEET_NAMES.NOTIFY_BATCHES);
        if (!batchSheet)
            throw new Error('通知バッチシートが存在しません');
        const batchData = batchSheet.getDataRange().getValues();
        if (batchData.length <= 1)
            return;
        const batchHeader = getHeaderMap(batchData[0]);
        const tz = ss.getSpreadsheetTimeZone();
        const senmuTo = splitEmails(settings.senmuTo);
        const senmuCc = splitEmails(settings.senmuCc);
        if (senmuTo.length === 0) {
            appendNotificationLog('専務依頼', '', '', '', '専務_通知先Toが未設定');
            return;
        }
        for (let i = 1; i < batchData.length; i++) {
            const row = batchData[i];
            const requestedAt = parseDateValue(getCellRaw(row, batchHeader['専務依頼送信日時']));
            if (requestedAt)
                continue;
            const batchId = getCellValue(row, batchHeader['batchId']);
            const sheetName = getCellValue(row, batchHeader['確認シート名']);
            if (!sheetName)
                continue;
            const confirmSheet = ss.getSheetByName(sheetName);
            if (!confirmSheet)
                continue;
            const confirmData = confirmSheet.getDataRange().getValues();
            if (confirmData.length <= 1)
                continue;
            const confirmHeader = getHeaderMap(confirmData[0]);
            const allConfirmed = confirmData
                .slice(1)
                .every((v) => toBoolean(getCellRaw(v, confirmHeader['回答確認済み']), false));
            if (!allConfirmed)
                continue;
            const sheetUrl = buildSheetUrlWithGid(ss, confirmSheet);
            const rangeStart = parseDateValue(getCellRaw(row, batchHeader['対象開始日']));
            const rangeEnd = parseDateValue(getCellRaw(row, batchHeader['対象終了日']));
            const subject = `【車両更新確認】専務承認依頼 ${formatDateLabel(rangeStart || new Date(), tz)}〜${formatDateLabel(rangeEnd || new Date(), tz)}`;
            const bodyText = [
                '専務 各位',
                '',
                '本部長・副本部長の確認が完了しました。',
                '確認シートの「専務判断」列へ承認/差戻しの入力をお願いします。',
                '',
                `確認シート: ${sheetUrl}`,
            ].join('\n');
            const bodyHtml = [
                '<p>専務 各位</p>',
                '<p>本部長・副本部長の確認が完了しました。<br>確認シートの「専務判断」列へ承認/差戻しの入力をお願いします。</p>',
                `<p>確認シート: <a href="${sheetUrl}">${sheetUrl}</a></p>`,
            ].join('');
            try {
                MailApp.sendEmail({
                    to: senmuTo.join(','),
                    cc: senmuCc.join(','),
                    subject,
                    body: bodyText,
                    htmlBody: bodyHtml,
                    name: settings.fromName,
                });
                row[batchHeader['専務依頼送信日時'] - 1] = new Date();
                row[batchHeader['ステータス'] - 1] = BATCH_STATUS.SENMU_REQUESTED;
                appendNotificationLog('専務依頼', '', senmuTo.join(','), batchId, '送信OK');
            }
            catch (err) {
                appendNotificationLog('専務依頼', '', senmuTo.join(','), batchId, `送信失敗: ${err}`);
            }
        }
        batchSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);
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
        const batchSheet = ss.getSheetByName(SHEET_NAMES.NOTIFY_BATCHES);
        const vehicleSheet = ss.getSheetByName(VEHICLE_SHEET_NAME);
        if (!batchSheet || !vehicleSheet)
            throw new Error('必要シートが存在しません');
        ensureAppendColumns(vehicleSheet, VEHICLE_MASTER_EXTRA_HEADERS);
        const batchData = batchSheet.getDataRange().getValues();
        if (batchData.length <= 1)
            return;
        const batchHeader = getHeaderMap(batchData[0]);
        const tz = ss.getSpreadsheetTimeZone();
        const vehicleData = vehicleSheet.getDataRange().getValues();
        if (vehicleData.length <= 1)
            return;
        const vehicleHeader = getHeaderMap(vehicleData[0]);
        const sourceHeader = resolveSourceHeaders(vehicleHeader);
        const vehicleIndex = buildVehicleIndex(vehicleData, sourceHeader, vehicleHeader);
        for (let i = 1; i < batchData.length; i++) {
            const batchRow = batchData[i];
            const sheetName = getCellValue(batchRow, batchHeader['確認シート名']);
            if (!sheetName)
                continue;
            const confirmSheet = ss.getSheetByName(sheetName);
            if (!confirmSheet)
                continue;
            protectSenmuColumns(confirmSheet, settings);
            const confirmData = confirmSheet.getDataRange().getValues();
            if (confirmData.length <= 1)
                continue;
            const confirmHeader = getHeaderMap(confirmData[0]);
            const now = new Date();
            let appliedAny = false;
            let hasReturnedDecision = false;
            for (let r = 1; r < confirmData.length; r++) {
                const row = confirmData[r];
                const decision = getCellValue(row, confirmHeader['専務判断']);
                if (decision === SENMU_DECISION.RETURN) {
                    hasReturnedDecision = true;
                    continue;
                }
                if (decision !== SENMU_DECISION.APPROVE)
                    continue;
                const alreadyApplied = toBoolean(getCellRaw(row, confirmHeader['マスター反映済み']), false);
                if (alreadyApplied)
                    continue;
                const hqResponse = getCellValue(row, confirmHeader['本部回答']);
                if (!hqResponse)
                    continue;
                const reg = getCellValue(row, confirmHeader['登録番号']);
                const chassis = getCellValue(row, confirmHeader['車台番号']);
                const vehicleRowIndex = resolveVehicleRowIndex(vehicleIndex, reg, chassis);
                if (vehicleRowIndex === null) {
                    appendNotificationLog('反映', '', '', sheetName, `車両特定不可: 登録番号=${reg} 車台番号=${chassis}`);
                    continue;
                }
                if (hqResponse === ANSWER_LABELS.RENEW) {
                    const newStart = parseDateValue(getCellRaw(row, confirmHeader['新契約開始日']));
                    const newEnd = parseDateValue(getCellRaw(row, confirmHeader['新契約満了日']));
                    if (!newStart || !newEnd) {
                        appendNotificationLog('反映', '', '', sheetName, `更新の新契約日が未入力: ${reg}`);
                        continue;
                    }
                    if (sourceHeader.contractStart) {
                        vehicleData[vehicleRowIndex][sourceHeader.contractStart - 1] = newStart;
                    }
                    if (sourceHeader.contractEnd) {
                        vehicleData[vehicleRowIndex][sourceHeader.contractEnd - 1] = newEnd;
                    }
                }
                else if (hqResponse === ANSWER_LABELS.CANCELLATION_REPLACE || hqResponse === ANSWER_LABELS.CANCELLATION_END) {
                    const completed = toBoolean(getCellRaw(row, confirmHeader['解約完了']), false);
                    if (!completed) {
                        appendNotificationLog('反映', '', '', sheetName, `解約完了が未チェック: ${reg}`);
                        continue;
                    }
                    const rowNumber = vehicleRowIndex + 1;
                    const lastColumn = vehicleSheet.getLastColumn();
                    vehicleSheet.getRange(rowNumber, 1, 1, lastColumn).setBackground('#d9d9d9');
                }
                if (vehicleHeader['マスター反映済み']) {
                    vehicleData[vehicleRowIndex][vehicleHeader['マスター反映済み'] - 1] = true;
                }
                if (vehicleHeader['反映日時']) {
                    vehicleData[vehicleRowIndex][vehicleHeader['反映日時'] - 1] = now;
                }
                row[confirmHeader['マスター反映済み'] - 1] = true;
                row[confirmHeader['反映日時'] - 1] = now;
                appliedAny = true;
            }
            if (appliedAny) {
                confirmSheet.getRange(1, 1, confirmData.length, confirmData[0].length).setValues(confirmData);
            }
            const approvedRows = confirmData
                .slice(1)
                .filter((v) => getCellValue(v, confirmHeader['専務判断']) === SENMU_DECISION.APPROVE);
            const allApprovedApplied = approvedRows.every((v) => toBoolean(getCellRaw(v, confirmHeader['マスター反映済み']), false));
            if (hasReturnedDecision) {
                batchRow[batchHeader['ステータス'] - 1] = BATCH_STATUS.RETURNED;
                appendNotificationLog('反映', '', '', sheetName, '差戻しありのため反映待ち');
            }
            else if (approvedRows.length > 0 && allApprovedApplied) {
                batchRow[batchHeader['ステータス'] - 1] = BATCH_STATUS.APPLIED;
                batchRow[batchHeader['反映完了日時'] - 1] = new Date();
            }
        }
        vehicleSheet.getRange(1, 1, vehicleData.length, vehicleData[0].length).setValues(vehicleData);
        batchSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);
    }
    finally {
        lock.releaseLock();
    }
}
function runDaily() {
    syncSchema();
    createRequests();
    sendInitialEmails();
    sendReminderEmails();
    sendApprovalRequestEmails();
    applyApprovalDecisions();
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
        const ss = getSpreadsheet();
        const tz = ss.getSpreadsheetTimeZone();
        const settings = loadSettings();
        const year = new Date().getFullYear();
        const testSettings = {
            ...settings,
            semiannualSendDateMarch: '03-01',
            semiannualSendDateSeptember: '09-01',
            responseDeadlineMarch: '03-31',
            responseDeadlineSeptember: '09-30',
        };
        const marchSchedule = resolveSemiannualSchedule(new Date(year, 2, 1), testSettings, tz);
        const expectedMarchStart = toDateOnly(new Date(year, 9, 1), tz);
        const expectedMarchEnd = toDateOnly(new Date(year + 1, 2, 31), tz);
        const expectedMarchDeadline = toDateOnly(new Date(year, 2, 31), tz);
        appendTestResult('3月便: スケジュール生成', marchSchedule ? 'OK' : 'NG', marchSchedule ? '生成あり' : '生成なし');
        appendTestResult('3月便: 抽出範囲', marchSchedule &&
            marchSchedule.rangeStart.getTime() === expectedMarchStart.getTime() &&
            marchSchedule.rangeEnd.getTime() === expectedMarchEnd.getTime()
            ? 'OK'
            : 'NG', marchSchedule
            ? `${formatDateLabel(marchSchedule.rangeStart, tz)}〜${formatDateLabel(marchSchedule.rangeEnd, tz)}`
            : '範囲なし');
        appendTestResult('3月便: 期限(送付年3/31)', marchSchedule && marchSchedule.deadline.getTime() === expectedMarchDeadline.getTime() ? 'OK' : 'NG', marchSchedule ? formatDateLabel(marchSchedule.deadline, tz) : '期限なし');
        const septemberSchedule = resolveSemiannualSchedule(new Date(year, 8, 1), testSettings, tz);
        const expectedSeptemberStart = toDateOnly(new Date(year, 3, 1), tz);
        const expectedSeptemberEnd = toDateOnly(new Date(year, 8, 30), tz);
        const expectedSeptemberDeadline = toDateOnly(new Date(year, 8, 30), tz);
        appendTestResult('9月便: スケジュール生成', septemberSchedule ? 'OK' : 'NG', septemberSchedule ? '生成あり' : '生成なし');
        appendTestResult('9月便: 抽出範囲', septemberSchedule &&
            septemberSchedule.rangeStart.getTime() === expectedSeptemberStart.getTime() &&
            septemberSchedule.rangeEnd.getTime() === expectedSeptemberEnd.getTime()
            ? 'OK'
            : 'NG', septemberSchedule
            ? `${formatDateLabel(septemberSchedule.rangeStart, tz)}〜${formatDateLabel(septemberSchedule.rangeEnd, tz)}`
            : '範囲なし');
        appendTestResult('9月便: 期限(同年9/30)', septemberSchedule && septemberSchedule.deadline.getTime() === expectedSeptemberDeadline.getTime() ? 'OK' : 'NG', septemberSchedule ? formatDateLabel(septemberSchedule.deadline, tz) : '期限なし');
        const reminderDate = addDays(expectedSeptemberDeadline, -10);
        const expectedReminderDate = toDateOnly(new Date(year, 8, 20), tz);
        appendTestResult('リマインド: 期限10日前', reminderDate.getTime() === expectedReminderDate.getTime() ? 'OK' : 'NG', formatDateLabel(reminderDate, tz));
        const tempSheetName = `本部長副本部長確認_TEST_${Utilities.getUuid().slice(0, 8)}`;
        const tempConfirmSheet = ensureConfirmSheet(ss, tempSheetName, testSettings);
        try {
            const empty = new Array(CONFIRM_SHEET_HEADERS.length).fill('');
            tempConfirmSheet.getRange(2, 1, 1, empty.length).setValues([empty]);
            applyConfirmSheetValidations(tempConfirmSheet, 1);
            const tempHeader = getHeaderMap(tempConfirmSheet.getRange(1, 1, 1, tempConfirmSheet.getLastColumn()).getValues()[0]);
            const protections = tempConfirmSheet
                .getProtections(SpreadsheetApp.ProtectionType.RANGE)
                .filter((p) => p.getDescription() === 'managed_by_script:senmu_columns');
            const protectedCols = protections.map((p) => p.getRange().getColumn());
            const expectedProtectedCols = ['専務判断', '専務コメント']
                .map((name) => tempHeader[name])
                .filter((col) => !!col);
            const hasAllProtectedCols = expectedProtectedCols.every((col) => protectedCols.indexOf(col) >= 0);
            appendTestResult('専務列保護: 対象列', hasAllProtectedCols ? 'OK' : 'NG', `保護列=${protectedCols.join(',')}`);
            const hqValidation = tempHeader['本部回答'] ? tempConfirmSheet.getRange(2, tempHeader['本部回答']).getDataValidation() : null;
            const senmuValidation = tempHeader['専務判断'] ? tempConfirmSheet.getRange(2, tempHeader['専務判断']).getDataValidation() : null;
            appendTestResult('確認シート: 本部回答バリデーション', hqValidation ? 'OK' : 'NG', hqValidation ? '設定あり' : '設定なし');
            appendTestResult('確認シート: 専務判断バリデーション', senmuValidation ? 'OK' : 'NG', senmuValidation ? '設定あり' : '設定なし');
        }
        finally {
            ss.deleteSheet(tempConfirmSheet);
        }
        const vehicleHeaderRow = [
            '登録番号',
            '車台番号',
            '契約開始日',
            '契約満了日',
            '管理部門',
            '管理担当者',
            '契約期間',
            '車検満了日',
            'リース料（税抜）',
            'マスター反映済み',
            '反映日時',
        ];
        const vehicleHeader = getHeaderMap(vehicleHeaderRow);
        const sourceHeader = resolveSourceHeaders(vehicleHeader);
        const vehicleData = [
            vehicleHeaderRow,
            ['品川500あ1234', 'CH-001', '', '', '', '', '', '', '', false, ''],
            ['品川500あ5678', 'CH-002', '', '', '', '', '', '', '', false, ''],
        ];
        const vehicleIndex = buildVehicleIndex(vehicleData, sourceHeader, vehicleHeader);
        const bothMatched = resolveVehicleRowIndex(vehicleIndex, '品川500あ1234', 'CH-001') === 1;
        const regOnlyMatched = resolveVehicleRowIndex(vehicleIndex, '品川500あ5678', '') === 2;
        const unmatched = resolveVehicleRowIndex(vehicleIndex, '存在しない', 'NONE') === null;
        appendTestResult('台帳突合: 登録番号+車台番号', bothMatched ? 'OK' : 'NG', String(resolveVehicleRowIndex(vehicleIndex, '品川500あ1234', 'CH-001')));
        appendTestResult('台帳突合: 登録番号のみ', regOnlyMatched ? 'OK' : 'NG', String(resolveVehicleRowIndex(vehicleIndex, '品川500あ5678', '')));
        appendTestResult('台帳突合: 非該当', unmatched ? 'OK' : 'NG', String(resolveVehicleRowIndex(vehicleIndex, '存在しない', 'NONE')));
        appendTestResult('完了', 'OK', '');
        uiAlertSafe('テストが完了しました。結果は「テスト結果」シートを確認してください。');
    }
    catch (err) {
        appendTestResult('中断', 'NG', String(err));
        throw err;
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
        hqLeadersTo: toStringValue(values['本部長副本部長_通知先To'], String(SETTINGS_DEFAULTS['本部長副本部長_通知先To'])),
        senmuTo: toStringValue(values['専務_通知先To'], String(SETTINGS_DEFAULTS['専務_通知先To'])),
        senmuCc: toStringValue(values['専務_通知先Cc'], String(SETTINGS_DEFAULTS['専務_通知先Cc'])),
        semiannualSendDateMarch: toStringValue(values['半期送付日_3月'], String(SETTINGS_DEFAULTS['半期送付日_3月'])),
        semiannualSendDateSeptember: toStringValue(values['半期送付日_9月'], String(SETTINGS_DEFAULTS['半期送付日_9月'])),
        responseDeadlineMarch: toStringValue(values['回答期限_3月'], String(SETTINGS_DEFAULTS['回答期限_3月'])),
        responseDeadlineSeptember: toStringValue(values['回答期限_9月'], String(SETTINGS_DEFAULTS['回答期限_9月'])),
        reminderDaysBeforeDeadline: toNumber(values['リマインド_期限前日数'], Number(SETTINGS_DEFAULTS['リマインド_期限前日数'])),
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
function splitEmails(value) {
    return String(value || '')
        .split(/[\s,;]+/)
        .map((email) => email.trim())
        .filter((email) => email.length > 0);
}
function parseMonthDayToDate(value, year, tz) {
    const match = String(value || '').trim().match(/^(\d{1,2})[-/](\d{1,2})$/);
    if (!match)
        return null;
    const month = Number(match[1]);
    const day = Number(match[2]);
    if (!month || !day)
        return null;
    const date = new Date(year, month - 1, day);
    return toDateOnly(date, tz);
}
function resolveSemiannualSchedule(now, settings, tz) {
    const today = toDateOnly(now, tz);
    const year = today.getFullYear();
    const marchSend = parseMonthDayToDate(settings.semiannualSendDateMarch, year, tz);
    const septemberSend = parseMonthDayToDate(settings.semiannualSendDateSeptember, year, tz);
    if (marchSend && marchSend.getTime() === today.getTime()) {
        const rangeStart = toDateOnly(new Date(year, 9, 1), tz);
        const rangeEnd = toDateOnly(new Date(year + 1, 2, 31), tz);
        const deadline = parseMonthDayToDate(settings.responseDeadlineMarch, year, tz) || marchSend;
        return { sendDate: marchSend, deadline, rangeStart, rangeEnd };
    }
    if (septemberSend && septemberSend.getTime() === today.getTime()) {
        const rangeStart = toDateOnly(new Date(year, 3, 1), tz);
        const rangeEnd = toDateOnly(new Date(year, 8, 30), tz);
        const deadline = parseMonthDayToDate(settings.responseDeadlineSeptember, year, tz) || septemberSend;
        return { sendDate: septemberSend, deadline, rangeStart, rangeEnd };
    }
    return null;
}
function ensureConfirmSheet(ss, name, settings) {
    const sheet = ensureSheet(ss, name);
    ensureHeaders(sheet, 1, CONFIRM_SHEET_HEADERS);
    protectSenmuColumns(sheet, settings);
    return sheet;
}
function applyConfirmSheetValidations(sheet, rowCount) {
    const headerMap = getHeaderMap(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
    const setValidation = (headerName, options) => {
        const col = headerMap[headerName];
        if (!col)
            return;
        const rule = SpreadsheetApp.newDataValidation().requireValueInList(options, true).build();
        sheet.getRange(2, col, rowCount, 1).setDataValidation(rule);
    };
    const setCheckbox = (headerName) => {
        const col = headerMap[headerName];
        if (!col)
            return;
        sheet.getRange(2, col, rowCount, 1).insertCheckboxes();
    };
    setValidation('本部回答', HQ_RESPONSE_OPTIONS);
    setValidation('専務判断', Object.values(SENMU_DECISION));
    setCheckbox('回答確認済み');
    setCheckbox('解約完了');
    setCheckbox('マスター反映済み');
}
function protectSenmuColumns(sheet, settings) {
    const headerMap = getHeaderMap(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
    const columns = ['専務判断', '専務コメント'];
    const editorSet = {};
    splitEmails(settings.senmuTo).forEach((email) => {
        editorSet[email] = true;
    });
    try {
        const effectiveUserEmail = Session.getEffectiveUser().getEmail();
        if (effectiveUserEmail)
            editorSet[effectiveUserEmail] = true;
    }
    catch (err) {
        Logger.log(`protectSenmuColumns getEffectiveUser failed: ${err}`);
    }
    try {
        const ownerEmail = sheet.getParent().getOwner().getEmail();
        if (ownerEmail)
            editorSet[ownerEmail] = true;
    }
    catch (err) {
        Logger.log(`protectSenmuColumns getOwner failed: ${err}`);
    }
    const editors = Object.keys(editorSet);
    if (editors.length === 0)
        return;
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    protections
        .filter((p) => p.getDescription() === 'managed_by_script:senmu_columns')
        .forEach((p) => p.remove());
    columns.forEach((name) => {
        const col = headerMap[name];
        if (!col)
            return;
        const range = sheet.getRange(2, col, Math.max(sheet.getMaxRows() - 1, 1), 1);
        const protection = range.protect();
        protection.setDescription('managed_by_script:senmu_columns');
        protection.setWarningOnly(false);
        protection.removeEditors(protection.getEditors());
        protection.addEditors(editors);
    });
}
function formatConfirmVehicleLine(row, headerMap, tz) {
    const reg = getCellValue(row, headerMap['登録番号']);
    const type = getCellValue(row, headerMap['車種']);
    const end = parseDateValue(getCellRaw(row, headerMap['契約満了日']));
    const endLabel = end ? formatDateLabel(end, tz) : '未設定';
    return `${reg || '登録番号不明'} / ${type || '車種不明'} / 満了:${endLabel}`;
}
function buildSheetUrlWithGid(ss, sheet) {
    return `${ss.getUrl()}#gid=${sheet.getSheetId()}`;
}
function buildVehicleIndex(vehicleData, sourceHeader, headerMap) {
    const index = {};
    for (let i = 1; i < vehicleData.length; i++) {
        const row = vehicleData[i];
        const regCombined = getSourceRegistrationCombined(row, sourceHeader);
        const regAll = sourceHeader.regAll ? getCellValue(row, sourceHeader.regAll) : '';
        const reg = regCombined || regAll;
        const chassis = getCellValue(row, sourceHeader.chassis);
        if (reg) {
            const key = normalizeVehicleKey(reg);
            if (index[key] === undefined)
                index[key] = i;
        }
        if (reg && chassis) {
            const key = normalizeVehicleKey(`${reg}__${chassis}`);
            if (index[key] === undefined)
                index[key] = i;
        }
    }
    return index;
}
function resolveVehicleRowIndex(index, reg, chassis) {
    if (reg && chassis) {
        const key = normalizeVehicleKey(`${reg}__${chassis}`);
        if (index[key] !== undefined)
            return index[key];
    }
    if (reg) {
        const key = normalizeVehicleKey(reg);
        if (index[key] !== undefined)
            return index[key];
    }
    return null;
}
function normalizeVehicleKey(value) {
    return String(value || '').trim();
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
