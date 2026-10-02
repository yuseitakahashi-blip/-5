/**
 * Gate5: PC貸出処理（正常系のみの仕様書に、例外系を優先度「高」から実装）
 *
 * 配布資料の仕様書には正常系（貸出申請→確認→貸出完了）しか書かれていない。
 * 「これを満たさないと不合格」の必須2条件に直結する例外を優先実装し、
 * そのあと観点カテゴリのカバー範囲を広げるために②人の状態を追加した。
 *
 *   実装した例外（例外系洗い出しシートの No.1・No.7・No.8・No.10・No.13・No.29 に相当）
 *     No.1  ①在庫・状態：選択端末が既に貸出中／修理中／廃棄済みなら貸出不可
 *     No.7  ②人の状態：退職済み社員は貸出不可
 *     No.8  ②人の状態：休職中社員は貸出不可
 *     No.10 ②人の状態：存在しない社員ID（不正な申請者）は貸出不可
 *     No.13 ③重複・多重：同時アクセス時の二重貸出防止（LockService + 再チェック）
 *     No.29 ⑦失敗時　：lendings登録とdevices更新の2テーブル更新を疑似トランザクション化
 *                     （devices更新に失敗したらlendings登録をロールバックする）
 *   → No.1・No.13・No.29 で、必須条件「二重貸出が防止されていること」
 *     「複数テーブル更新が1トランザクションになっていること」の両方を満たす。
 *   → No.7・No.8・No.10 で観点カテゴリを①③⑦の3つから①②③⑦の4つに拡大。
 *
 * それ以外の例外（入力値・順序・時間・不正操作など）は未実装。
 * Gate1と異なり、この課題には配布済みスプレッドシートが無いため、
 * このスクリプト自身がコンテナバインド先のスプレッドシートに
 * devices / lendings / employees シートを作成し、初期データを投入する。
 */

var SHEETS = { DEVICES: 'devices', LENDINGS: 'lendings', EMPLOYEES: 'employees' };

var DEVICE_STATUS_LABEL = {
  AVAILABLE: '貸出可能',
  LENT: '貸出中',
  REPAIR: '修理中',
  DISPOSED: '廃棄済み'
};

var EMPLOYMENT_STATUS_LABEL = {
  ACTIVE: '在籍中',
  LEAVE: '休職中',
  RETIRED: '退職済み'
};

// ============================================================
// Web エントリポイント
// ============================================================
function doGet() {
  ensureSetup_();
  var t = HtmlService.createTemplateFromFile('index');
  return t.evaluate()
    .setTitle('PC貸出処理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// スプレッドシート初期化・シード（配布済みスプシが無いため自前で作る）
// ============================================================
function getSs_() {
  // コンテナバインド型：このスクリプトが紐付くスプレッドシートをそのままDB代わりに使う
  return SpreadsheetApp.getActiveSpreadsheet();
}

function ensureSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

function ensureSetup_() {
  var ss = getSs_();
  if (ss.getName() === '無題のスプレッドシート' || ss.getName() === 'Untitled spreadsheet') {
    ss.rename('Gate5_PC貸出管理DB');
  }
  ensureSheet_(ss, SHEETS.DEVICES, ['id', 'asset_no', 'model_name', 'device_type', 'status', 'purchased_at']);
  ensureSheet_(ss, SHEETS.LENDINGS, ['id', 'device_id', 'user_id', 'lent_at', 'due_date', 'returned_at', 'purpose']);
  ensureSheet_(ss, SHEETS.EMPLOYEES, ['id', 'name', 'department', 'employment_status']);

  var defaultSheet = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 3) {
    try { ss.deleteSheet(defaultSheet); } catch (err) { /* noop */ }
  }

  seedIfEmpty_(ss);
  return ss;
}

/**
 * デモ・モック撮影用: 貸出データを初期状態に戻す。
 * Apps Scriptエディタでこの関数を選んで実行ボタン(▷)を押すと使えます。
 */
function resetForDemo() {
  var ss = getSs_();
  clearDataRows_(ss.getSheetByName(SHEETS.DEVICES));
  clearDataRows_(ss.getSheetByName(SHEETS.LENDINGS));
  seedIfEmpty_(ss);
}

function clearDataRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
}

function seedIfEmpty_(ss) {
  var employeesSheet = ss.getSheetByName(SHEETS.EMPLOYEES);
  if (employeesSheet.getLastRow() < 2) {
    employeesSheet.getRange(2, 1, 4, 4).setValues([
      [1, '佐藤 花子', '情報システム部', 'ACTIVE'],
      [2, '鈴木 一郎', '営業部', 'ACTIVE'],
      [3, '高橋 次郎', '総務部', 'LEAVE'],
      [4, '田中 三郎', '経理部', 'RETIRED']
    ]);
  }

  var devicesSheet = ss.getSheetByName(SHEETS.DEVICES);
  if (devicesSheet.getLastRow() < 2) {
    devicesSheet.getRange(2, 1, 5, 6).setValues([
      [1, 'PC-0001', 'ThinkPad X1 Carbon', 'LAPTOP', 'AVAILABLE', '2024-04-01'],
      [2, 'PC-0002', 'ThinkPad X1 Carbon', 'LAPTOP', 'LENT', '2024-04-01'],
      [3, 'PC-0003', 'MacBook Air', 'LAPTOP', 'AVAILABLE', '2024-06-01'],
      [4, 'PC-0004', 'iPad 10th gen', 'TABLET', 'REPAIR', '2023-10-01'],
      [5, 'PC-0005', 'Dell XPS 13', 'LAPTOP', 'DISPOSED', '2022-01-01']
    ]);
  }

  var lendingsSheet = ss.getSheetByName(SHEETS.LENDINGS);
  if (lendingsSheet.getLastRow() < 2) {
    // devices の PC-0002 (id=2) が LENT の状態と辻褄を合わせるための既存貸出データ
    var lentAt = new Date();
    lentAt.setDate(lentAt.getDate() - 2);
    var due = new Date();
    due.setDate(due.getDate() + 5);
    lendingsSheet.getRange(2, 1, 1, 7).setValues([
      [1, 2, 2, lentAt.toISOString(), formatDate_(due), '', '常用端末として利用中']
    ]);
  }
}

function formatDate_(d) {
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
}

function nowIso_() {
  return new Date().toISOString();
}

// ============================================================
// シート読み書きヘルパー
// ============================================================
function sheetToObjects_(sheet) {
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row[0] === '' || row[0] === null) continue;
    var obj = {};
    for (var c = 0; c < headers.length; c++) obj[headers[c]] = row[c];
    obj._rowIndex = i + 1; // 1-based（ヘッダー込み）
    out.push(obj);
  }
  return out;
}

function getDevices_() {
  return sheetToObjects_(getSs_().getSheetByName(SHEETS.DEVICES));
}

function getDeviceById_(id) {
  var devices = getDevices_();
  for (var i = 0; i < devices.length; i++) {
    if (String(devices[i].id) === String(id)) return devices[i];
  }
  return null;
}

function getEmployees_() {
  return sheetToObjects_(getSs_().getSheetByName(SHEETS.EMPLOYEES));
}

function getEmployeeById_(id) {
  if (id === null || id === undefined || id === '') return null;
  var employees = getEmployees_();
  for (var i = 0; i < employees.length; i++) {
    if (String(employees[i].id) === String(id)) return employees[i];
  }
  return null;
}

function nextId_(sheetName) {
  var rows = sheetToObjects_(getSs_().getSheetByName(sheetName));
  var max = 0;
  rows.forEach(function (r) { if (Number(r.id) > max) max = Number(r.id); });
  return max + 1;
}

function updateDeviceStatus_(rowIndex, status) {
  getSs_().getSheetByName(SHEETS.DEVICES).getRange(rowIndex, 5).setValue(status); // 5列目 = status
}

// ============================================================
// クライアント公開API
// ============================================================
function apiBootstrap() {
  ensureSetup_();
  return {
    employees: getEmployees_().map(function (e) {
      return {
        id: e.id,
        name: e.name,
        employmentStatus: e.employment_status,
        employmentStatusLabel: EMPLOYMENT_STATUS_LABEL[e.employment_status] || e.employment_status
      };
    }),
    devices: getDevices_().map(function (d) {
      return {
        id: d.id,
        assetNo: d.asset_no,
        modelName: d.model_name,
        status: d.status,
        statusLabel: DEVICE_STATUS_LABEL[d.status] || d.status
      };
    }),
    defaultDueDate: formatDate_(addDays_(new Date(), 7))
  };
}

function addDays_(date, days) {
  var d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

/**
 * 貸出処理（正常系 + 例外系）
 *
 * 正常系：lendings に新規登録 → devices.status を LENT に更新 → 完了メッセージを返す
 *
 * 実装した例外系：
 *   ②人の状態
 *     申請者（employees）が在籍中（ACTIVE）でなければ貸出不可。
 *     退職済み／休職中はそれぞれ専用メッセージを返し、社員IDが存在しない
 *     （なりすまし・不正なリクエスト値）場合も同じ経路で弾く。
 *   ①在庫・状態／③重複・多重
 *     LockService で同時アクセスを直列化したうえで、ロック取得後に
 *     devices.status を再取得して AVAILABLE かどうかを再チェックする。
 *     これにより「連打」でも「別タブからの同時申請」でも二重貸出は成立しない。
 *   ⑦失敗時
 *     lendings登録とdevices更新の間に purpose === 'FORCE_TX_FAIL' を渡すと
 *     わざと例外を発生させられる（デモ用フック）。その場合は直前に登録した
 *     lendings行を削除してロールバックし、貸出前の状態に戻す。
 */
function apiLendDevice(params) {
  ensureSetup_();
  var deviceId = params.deviceId;
  var userId = params.userId;
  var dueDate = params.dueDate;
  var purpose = params.purpose || '';

  // 人の状態チェックはロック取得前でよい（申請者の状態は他リクエストと競合しないため）
  var employee = getEmployeeById_(userId);
  if (!employee) {
    return { success: false, message: '社員情報が確認できません。再ログインしてください。' };
  }
  if (employee.employment_status !== 'ACTIVE') {
    var employeeMessages = {
      RETIRED: '退職済みのため貸出できません。総務までご連絡ください。',
      LEAVE: '休職中のため貸出できません。'
    };
    return {
      success: false,
      message: employeeMessages[employee.employment_status] || 'この社員は現在貸出を利用できません。'
    };
  }

  var lock = LockService.getScriptLock();
  var gotLock = lock.tryLock(10000);
  if (!gotLock) {
    // 同時アクセスが極端に混み合った場合の最終防波堤
    return { success: false, message: '只今混み合っています。少し時間をおいて再度お試しください。' };
  }

  try {
    // ロック取得後に最新の端末状態を読み直す（ここが二重貸出防止の要）
    var device = getDeviceById_(deviceId);
    if (!device) {
      return { success: false, message: '選択した端末が見つかりません。一覧を更新してやり直してください。' };
    }
    if (device.status !== 'AVAILABLE') {
      var messages = {
        LENT: 'この端末は現在貸出中です。別の端末を選択してください。',
        REPAIR: 'この端末は現在修理中のため貸し出せません。',
        DISPOSED: 'この端末は廃棄済みのため貸し出せません。'
      };
      return {
        success: false,
        message: messages[device.status] || 'この端末は現在貸し出せない状態です。'
      };
    }

    // --- ここから2テーブル更新（疑似トランザクション） ---
    var lendingsSheet = getSs_().getSheetByName(SHEETS.LENDINGS);
    var newId = nextId_(SHEETS.LENDINGS);
    var lentAt = nowIso_();
    lendingsSheet.appendRow([newId, device.id, userId, lentAt, dueDate, '', purpose]);
    var insertedRowIndex = lendingsSheet.getLastRow();

    try {
      if (purpose === 'FORCE_TX_FAIL') {
        // デモ用フック：devices更新の直前に意図的に失敗させる
        throw new Error('意図的なトランザクション失敗（デモ用フック FORCE_TX_FAIL）');
      }
      updateDeviceStatus_(device._rowIndex, 'LENT');
    } catch (txErr) {
      // devices更新に失敗 → 直前に登録したlendings行をロールバック（削除）する
      lendingsSheet.deleteRow(insertedRowIndex);
      return {
        success: false,
        message: '貸出処理に失敗しました。お手数ですが最初からやり直してください。'
      };
    }
    // --- 疑似トランザクションここまで ---

    return {
      success: true,
      message: device.asset_no + ' を貸し出しました。返却予定日は ' + formatDisplayDate_(dueDate) + ' です。'
    };
  } finally {
    lock.releaseLock();
  }
}

function formatDisplayDate_(yyyyMmDd) {
  var parts = String(yyyyMmDd).split('-');
  if (parts.length !== 3) return String(yyyyMmDd);
  return parts[0] + '/' + parts[1] + '/' + parts[2];
}
