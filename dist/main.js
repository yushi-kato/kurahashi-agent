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
};
const SOURCE_SHEETS = ['車両一覧', '車両一覧【ｹﾝｽｲ】', '車両一覧【ﾈｸｽﾄ】'];
const REQUEST_STATUS = {
    CREATED: '作成済',
    SENT: '送信済',
    RESPONDING: '回答中',
    COMPLETED: '完了',
    EXPIRED: '締切',
};
const ANSWER_OPTIONS = ['再リース', '新車入替', '廃止', '未定'];
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
            '更新方針',
            '依頼ID',
            '回答日',
            '備考',
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
            '再リース',
            '新車入替',
            '廃止',
            '未定',
            '未回答',
            '最終更新日時',
        ],
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
    集計_シート出力: true,
    集計_メール送信: true,
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
        .addItem('スキーマ同期', 'syncSchema')
        .addItem('スキーマドリフト確認', 'checkSchemaDrift')
        .addSeparator()
        .addItem('車両統合ビュー同期', 'syncVehicles')
        .addItem('更新依頼作成', 'createRequests')
        .addItem('初回メール送信', 'sendInitialEmails')
        .addSeparator()
        .addItem('回答反映', 'applyAnswers')
        .addItem('回答集計更新', 'buildSummarySheet')
        .addItem('集計メール送信', 'sendSummaryEmail')
        .addSeparator()
        .addItem('日次一括実行', 'runDaily')
        .addToUi();
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
        ensureHeaders(ensureSheet(ss, SHEET_NAMES.VEHICLE_VIEW), 1, getSchemaHeaders(SHEET_NAMES.VEHICLE_VIEW));
        ensureHeaders(ensureSheet(ss, SHEET_NAMES.NEEDS_INPUT), 1, getSchemaHeaders(SHEET_NAMES.NEEDS_INPUT));
        const deptMaster = loadDeptMaster();
        const rows = [];
        const needsInputRows = [];
        const now = new Date();
        const tz = ss.getSpreadsheetTimeZone();
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
                const regArea = getCellValue(row, headerIndexes.regArea);
                const regClass = getCellValue(row, headerIndexes.regClass);
                const regKana = getCellValue(row, headerIndexes.regKana);
                const regNumber = getCellValue(row, headerIndexes.regNumber);
                const regCombined = buildRegistrationCombined(regArea, regClass, regKana, regNumber);
                const vehicleType = getCellValue(row, headerIndexes.vehicleType);
                const chassis = getCellValue(row, headerIndexes.chassis);
                const contractStart = parseDateValue(getCellRaw(row, headerIndexes.contractStart));
                const contractEnd = parseDateValue(getCellRaw(row, headerIndexes.contractEnd));
                const dept = getCellValue(row, headerIndexes.dept);
                const vehicleId = buildVehicleId(sheetName, regCombined, chassis, i + 1);
                if (!contractEnd) {
                    needsInputRows.push([now, sheetName, vehicleId, dept, regCombined, vehicleType, '契約満了日なし']);
                }
                if (!dept) {
                    needsInputRows.push([now, sheetName, vehicleId, dept, regCombined, vehicleType, '管理部門なし']);
                }
                else if (!deptMaster[dept]) {
                    needsInputRows.push([now, sheetName, vehicleId, dept, regCombined, vehicleType, '部署マスタ未登録']);
                }
                rows.push([
                    vehicleId,
                    sheetName,
                    regArea,
                    regClass,
                    regKana,
                    regNumber,
                    regCombined,
                    vehicleType,
                    chassis,
                    contractStart,
                    contractEnd,
                    dept,
                    '',
                    '',
                    '',
                    '',
                ]);
            }
        });
        writeSheetData(SHEET_NAMES.VEHICLE_VIEW, rows);
        writeSheetData(SHEET_NAMES.NEEDS_INPUT, needsInputRows);
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
            newRequestRows.push([
                requestId,
                dept,
                startDate,
                endDate,
                deadline,
                REQUEST_STATUS.CREATED,
                '',
                '',
                0,
                requestToken,
            ]);
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
            const requestToken = getCellValue(row, reqHeader['requestToken']);
            const deptInfo = deptMaster[dept];
            if (!deptInfo || !deptInfo.active) {
                appendNotificationLog('初回', dept, '', requestId, '部署マスタ未登録/無効');
                continue;
            }
            if (!requestToken || !deptInfo.token) {
                appendNotificationLog('初回', dept, deptInfo.to, requestId, 'トークン不足');
                continue;
            }
            if (!deptInfo.to) {
                appendNotificationLog('初回', dept, '', requestId, '通知先Toが未設定');
                continue;
            }
            if (!settings.webAppUrl) {
                appendNotificationLog('初回', dept, deptInfo.to, requestId, 'Web回答URLが未設定');
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
            const url = buildWebAppUrl(settings.webAppUrl, {
                requestId,
                token: requestToken,
                dept,
                deptToken: deptInfo.token,
            });
            const listText = vehicles
                .map((v) => formatVehicleLine(v, vehicleHeader, tz))
                .join('\n');
            const listHtml = vehicles
                .map((v) => `<li>${escapeHtml(formatVehicleLine(v, vehicleHeader, tz))}</li>`)
                .join('');
            const subject = applyTemplate(settings.subjectTemplate, {
                管理部門: dept,
                対象開始日: formatDateLabel(targetStart || new Date(), tz),
                対象終了日: formatDateLabel(targetEnd || new Date(), tz),
                締切日: formatDateLabel(deadline || new Date(), tz),
                URL: url,
            });
            const bodyText = applyTemplate(settings.bodyTemplate, {
                管理部門: dept,
                対象開始日: formatDateLabel(targetStart || new Date(), tz),
                対象終了日: formatDateLabel(targetEnd || new Date(), tz),
                締切日: formatDateLabel(deadline || new Date(), tz),
                URL: url,
                車両一覧: listText,
            });
            const htmlTemplate = applyTemplate(settings.bodyTemplate, {
                管理部門: dept,
                対象開始日: formatDateLabel(targetStart || new Date(), tz),
                対象終了日: formatDateLabel(targetEnd || new Date(), tz),
                締切日: formatDateLabel(deadline || new Date(), tz),
                URL: url,
                車両一覧: '[[VEHICLE_LIST]]',
            });
            const htmlBody = escapeHtml(htmlTemplate)
                .replace(/\n/g, '<br>')
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
function doGet(e) {
    const params = e && e.parameter ? e.parameter : {};
    const validation = validateRequestAccess(params);
    if (!validation.ok) {
        return HtmlService.createHtmlOutput(`<p>${escapeHtml(validation.message)}</p>`).setTitle('車両更新回答');
    }
    const request = validation.requestRow;
    const vehicles = getVehiclesByRequestId(validation.requestId);
    const answers = loadAnswersForRequest(validation.requestId);
    const formRows = vehicles
        .map((v, index) => buildAnswerRowHtml(v, answers[v.vehicleId], index))
        .join('');
    const html = `
    <html>
      <head>
        <meta charset="utf-8">
        <style>
          body { font-family: sans-serif; }
          table { border-collapse: collapse; width: 100%; }
          th, td { border: 1px solid #ccc; padding: 6px; font-size: 14px; }
          th { background: #f5f5f5; }
          .muted { color: #777; }
        </style>
      </head>
      <body>
        <h2>車両更新方針 回答</h2>
        <p>管理部門: ${escapeHtml(request.dept)}</p>
        <p class="muted">requestId: ${escapeHtml(validation.requestId)}</p>
        <form method="post">
          <input type="hidden" name="requestId" value="${escapeHtml(validation.requestId)}">
          <input type="hidden" name="token" value="${escapeHtml(validation.requestToken)}">
          <input type="hidden" name="dept" value="${escapeHtml(request.dept)}">
          <input type="hidden" name="deptToken" value="${escapeHtml(request.deptToken)}">
          <p>
            回答者（任意）: <input type="text" name="responder" value="">
          </p>
          <table>
            <thead>
              <tr>
                <th>#</th>
                <th>車両</th>
                <th>契約満了日</th>
                <th>更新方針</th>
                <th>コメント</th>
              </tr>
            </thead>
            <tbody>
              ${formRows}
            </tbody>
          </table>
          <p><button type="submit">回答を送信</button></p>
        </form>
      </body>
    </html>
  `;
    return HtmlService.createHtmlOutput(html).setTitle('車両更新回答');
}
function doPost(e) {
    const params = e && e.parameter ? e.parameter : {};
    const paramsArray = e && e.parameters ? e.parameters : {};
    const validation = validateRequestAccess(params);
    if (!validation.ok) {
        return HtmlService.createHtmlOutput(`<p>${escapeHtml(validation.message)}</p>`).setTitle('車両更新回答');
    }
    const vehicleIds = ensureArray(paramsArray['vehicleId']);
    const answers = ensureArray(paramsArray['answer']);
    const comments = ensureArray(paramsArray['comment']);
    const responder = params['responder'] ? String(params['responder']) : '';
    const now = new Date();
    const answerInputs = [];
    for (let i = 0; i < vehicleIds.length; i++) {
        const answer = answers[i] ? String(answers[i]) : '';
        if (!answer)
            continue;
        answerInputs.push({
            requestId: validation.requestId,
            vehicleId: String(vehicleIds[i]),
            answer,
            comment: comments[i] ? String(comments[i]) : '',
            responder,
            answeredAt: now,
        });
    }
    if (answerInputs.length > 0) {
        upsertAnswers(answerInputs);
        applyAnswers();
        updateRequestStatus(validation.requestId);
    }
    return HtmlService.createHtmlOutput('<p>回答を受け付けました。ご協力ありがとうございました。</p>').setTitle('車両更新回答');
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
                    answer: getCellValue(row, answerHeader['回答']),
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
                const regCombined = buildRegistrationCombined(getCellValue(row, headerIndexes.regArea), getCellValue(row, headerIndexes.regClass), getCellValue(row, headerIndexes.regKana), getCellValue(row, headerIndexes.regNumber));
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
                再リース: 0,
                新車入替: 0,
                廃止: 0,
                未定: 0,
                未回答: 0,
            };
            vehicles.forEach((v) => {
                const policy = getCellValue(v, vehicleHeader['更新方針']);
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
                counts['再リース'],
                counts['新車入替'],
                counts['廃止'],
                counts['未定'],
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
            const lease = getCellValue(row, headerMap['再リース']);
            const replace = getCellValue(row, headerMap['新車入替']);
            const end = getCellValue(row, headerMap['廃止']);
            const pending = getCellValue(row, headerMap['未定']);
            const unanswered = getCellValue(row, headerMap['未回答']);
            lines.push(`${dept} (${range}) - 合計:${total} 再:${lease} 入替:${replace} 廃止:${end} 未定:${pending} 未回答:${unanswered}`);
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
function runDaily() {
    syncSchema();
    syncVehicles();
    createRequests();
    sendInitialEmails();
    buildSummarySheet();
    sendSummaryEmail();
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
    return {
        regArea: findHeaderIndex(headerMap, ['地名', '登録番号_地名', '登録番号（地名）']),
        regClass: findHeaderIndex(headerMap, ['分類番号', '分類番号(3桁)', '分類番号（3桁）', '登録番号_分類']),
        regKana: findHeaderIndex(headerMap, ['かな', '登録番号_かな']),
        regNumber: findHeaderIndex(headerMap, ['番号', '登録番号_番号']),
        vehicleType: findHeaderIndex(headerMap, ['車種']),
        chassis: findHeaderIndex(headerMap, ['車台番号']),
        contractStart: findHeaderIndex(headerMap, ['契約開始日']),
        contractEnd: findHeaderIndex(headerMap, ['契約満了日', '契約終了日']),
        dept: findHeaderIndex(headerMap, ['管理部門', '管理部署']),
    };
}
function findHeaderIndex(headerMap, names) {
    for (const name of names) {
        if (headerMap[name])
            return headerMap[name];
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
    const base = regCombined || chassis || `ROW${rowIndex}`;
    return `${sourceSheet}__${base}`;
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
        summarySheetEnabled: toBoolean(values['集計_シート出力'], Boolean(SETTINGS_DEFAULTS['集計_シート出力'])),
        summaryEmailEnabled: toBoolean(values['集計_メール送信'], Boolean(SETTINGS_DEFAULTS['集計_メール送信'])),
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
        return;
    const requestData = requestSheet.getDataRange().getValues();
    if (requestData.length <= 1)
        return;
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
    }
    requestSheet.getRange(1, 1, requestData.length, requestData[0].length).setValues(requestData);
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
