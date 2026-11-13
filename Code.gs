/**
 * ==========================================================================
 * LINKPRO — UNIFIED BACKEND (Code.gs — Users + Completions + Tasks Sheet)
 *
 * ĐÃ KHẮC PHỤC & NÂNG CẤP TOÀN DIỆN:
 * 1. Đồng bộ Role tức thì giữa Firebase và Sheet (`Users` tab):
 *    - Bỏ cache cứng trên trường `role`. Khi đổi role thành `devweb` hoặc `admin`
 *      trên Firebase Console, ngay lần gọi tiếp theo (hoặc khi mở tab Quản Lý Thành Viên)
 *      cột `role` (cột C) trong tab `Users` của Google Sheet sẽ tự động cập nhật thành `devweb`.
 * 2. Xác minh kép Role ở cả Firebase VÀ Google Sheet (`verifyPanelRoleFirebaseAndSheet`):
 *    - Mọi thao tác của Admin / Devweb từ Panel đều kiểm tra chéo role trên Firebase
 *      Firestore và role trên tab `Users` của Google Sheet.
 * 3. Nhiệm vụ Tự Động (`auto`):
 *    - Cộng tiền ngay lập tức vào cả tab `Users` trên Google Sheet lẫn `users/{uid}.balance`
 *      trên Firestore ngay sau khi người dùng hoàn thành nhiệm vụ.
 * 4. Sheet Tạo Nhiệm Vụ (`Tasks` tab) kết hợp với `getkey.gs`:
 *    - Tự động khởi tạo tab `Tasks` (nếu chưa có) để lưu và tạo nhiệm vụ trực tiếp
 *      trên Google Sheet hoặc đồng bộ 2 chiều với Web Panel.
 * ==========================================================================
 */

var TASK_SCRIPT_MODULES = typeof TASK_SCRIPT_MODULES !== 'undefined' ? TASK_SCRIPT_MODULES : {};

const DEFAULT_FIREBASE_WEB_API_KEY = 'AIzaSyBA3bgs4Q1FhsfShBvngU115oAhoBKs_uU';
const DEFAULT_FIRESTORE_PROJECT_ID = 'linkprovietnam';
const DEFAULT_RECAPTCHA_SECRET_KEY = '6Lf2MxkuAAAAAPAJy-4xjnNE_15t1G_LboFmlFQy';
const DEFAULT_LINK4M_API_TOKEN = '68a498c9a8b820353038b288';
const TASK_SESSION_MAX_SECONDS = 1800; // 30 phút = 1800 giây

const USERS_SHEET_NAME = 'Users';
const COMPLETIONS_SHEET_NAME = 'Completions';
const TASKS_SHEET_NAME = 'Tasks';

const USERS_COLS = {
  uid: 1,
  username: 2,
  role: 3,
  balance: 4,
  createdAt: 5,
  updatedAt: 6
};

const COMP_COLS = {
  completionId: 1,
  uid: 2,
  username: 3,
  taskId: 4,
  taskName: 5,
  reward: 6,
  type: 7,
  status: 8,
  createdAt: 9,
  approvedAt: 10,
  taskLink: 11,
  ipAddress: 12
};

const TASKS_COLS = {
  taskId: 1,
  title: 2,
  taskLink: 3,
  destinationUrl: 4,
  reward: 5,
  totalClicks: 6,
  remainingClicks: 7,
  maxPerUser: 8,
  verifySeconds: 9,
  approvalType: 10,
  isActive: 11,
  secretKey: 12,
  createdAt: 13,
  updatedAt: 14
};

function getScriptConfig() {
  let apiKey = DEFAULT_FIREBASE_WEB_API_KEY;
  let projectId = DEFAULT_FIRESTORE_PROJECT_ID;
  let link4mToken = DEFAULT_LINK4M_API_TOKEN;
  try {
    const props = PropertiesService.getScriptProperties();
    const propKey = props.getProperty('FIREBASE_WEB_API_KEY');
    const propProj = props.getProperty('FIRESTORE_PROJECT_ID');
    const propLink4m = props.getProperty('LINK4M_API_TOKEN');
    if (propKey && propKey.trim().indexOf('AIzaSy') === 0) {
      apiKey = propKey.trim();
    }
    if (propProj && propProj.trim().length > 2) {
      projectId = propProj.trim();
    }
    if (propLink4m && propLink4m.trim().length >= 10) {
      link4mToken = propLink4m.trim();
    }
  } catch (e) {}

  return {
    apiKey: apiKey,
    projectId: projectId,
    link4mToken: link4mToken,
    firestoreBaseUrl: 'https://firestore.googleapis.com/v1/projects/' + encodeURIComponent(projectId) + '/databases/(default)/documents',
    firestoreRootName: 'projects/' + projectId + '/databases/(default)/documents'
  };
}

/**
 * Menu tiện ích trên Google Sheet giúp Devweb tạo nhiệm vụ & đồng bộ trực tiếp từ Sheet
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('⚡ LinkPro')
      .addSubMenu(ui.createMenu('➕ Tạo nhiệm vụ (trên Sheet Tasks)')
        .addItem('Nhiệm vụ Tự động +500đ (2 lần/user, 75s)',  'menuTaskAuto500')
        .addItem('Nhiệm vụ Tự động +1.000đ (2 lần/user, 75s)', 'menuTaskAuto1000')
        .addItem('Nhiệm vụ Tự động +2.000đ (2 lần/user, 75s)', 'menuTaskAuto2000')
        .addSeparator()
        .addItem('Nhiệm vụ tuỳ chỉnh…', 'menuTaskCustom'))
      .addItem('🔗 Rút gọn Link4m & Đồng bộ tab Tasks lên Web', 'syncAllSheetTasksToFirestore')
      .addItem('🔄 Đồng bộ Role (devweb/admin) từ Firebase về Sheet', 'syncAllFirestoreUsersToSheet')
      .addSeparator()
      .addItem('🛠 Khởi tạo / kiểm tra 4 tab (Users, Completions, Tasks, Sessions)', 'setupAllSheets')
      .addToUi();
  } catch (e) {}
}

function setupAllSheets() {
  ensureUsersSheet();
  ensureCompletionsSheet();
  ensureTasksSheet();
  if (typeof gkSheet_ === 'function') gkSheet_();
  try {
    SpreadsheetApp.getUi().alert('Đã khởi tạo và kiểm tra đầy đủ 4 tab: Users, Completions, Tasks, Sessions!');
  } catch (e) {}
}

function menuTaskAuto500()  { createSheetTaskQuick_('Nhiệm vụ Vượt Link VIP (+500đ)', 500, 2, 1000, 75, 'auto'); }
function menuTaskAuto1000() { createSheetTaskQuick_('Nhiệm vụ Vượt Link VIP (+1.000đ)', 1000, 2, 1000, 75, 'auto'); }
function menuTaskAuto2000() { createSheetTaskQuick_('Nhiệm vụ Vượt Link VIP (+2.000đ)', 2000, 2, 1000, 75, 'auto'); }

function menuTaskCustom() {
  const ui = SpreadsheetApp.getUi();
  const rTitle = ui.prompt('Tạo nhiệm vụ mới', 'Tiêu đề nhiệm vụ:', ui.ButtonSet.OK_CANCEL);
  if (rTitle.getSelectedButton() !== ui.Button.OK) return;
  const title = rTitle.getResponseText().trim() || 'Nhiệm vụ Vượt Link';

  const rReward = ui.prompt('Tạo nhiệm vụ mới', 'Phần thưởng (VNĐ, ví dụ 500):', ui.ButtonSet.OK_CANCEL);
  if (rReward.getSelectedButton() !== ui.Button.OK) return;
  const reward = Math.max(50, parseInt(rReward.getResponseText(), 10) || 500);

  const rMax = ui.prompt('Tạo nhiệm vụ mới', 'Số lần vượt tối đa / 1 User (ví dụ 2):', ui.ButtonSet.OK_CANCEL);
  if (rMax.getSelectedButton() !== ui.Button.OK) return;
  const maxPerUser = Math.max(1, parseInt(rMax.getResponseText(), 10) || 2);

  const rSecs = ui.prompt('Tạo nhiệm vụ mới', 'Thời gian xác minh tối thiểu (giây, ví dụ 75):', ui.ButtonSet.OK_CANCEL);
  if (rSecs.getSelectedButton() !== ui.Button.OK) return;
  const verifySecs = Math.max(5, parseInt(rSecs.getResponseText(), 10) || 75);

  createSheetTaskQuick_(title, reward, maxPerUser, 1000, verifySecs, 'auto');
}

function createSheetTaskQuick_(title, reward, maxPerUser, totalClicks, verifySecs, approvalType) {
  const sh = ensureTasksSheet();
  const now = new Date();
  const taskId = 'task_' + now.getTime();
  const secretKey = 'LPK_' + Utilities.getUuid().replace(/-/g, '').substring(0, 16);
  const destUrl = (typeof GK !== 'undefined' && GK.SITE_URL) ? (GK.SITE_URL + '?taskId=' + taskId) : 'https://linkprovietnam.web.app/?taskId=' + taskId;
  const shortUrl = shortenLink4mOnServer(destUrl) || 'https://link4m.com/xxxxxx';

  sh.appendRow([
    taskId,
    title,
    shortUrl,
    destUrl,
    reward,
    totalClicks,
    totalClicks,
    maxPerUser,
    verifySecs,
    approvalType,
    true,
    secretKey,
    now,
    now
  ]);

  const lastRow = sh.getLastRow();
  syncSingleTaskRowFromSheetToFirestore(sh, lastRow, null);
  try {
    SpreadsheetApp.getUi().alert('Đã tạo nhiệm vụ "' + title + '" (dòng ' + lastRow + ') & đồng bộ lên Web!\nLink4m: ' + shortUrl);
  } catch (e) {}
}

function isValidSafeId(id) {
  return typeof id === 'string' && /^[a-zA-Z0-9_-]{1,128}$/.test(id.trim());
}

function sanitizeSheetCell(val, maxLen) {
  if (val === null || val === undefined) return '';
  let str = String(val).trim();
  if (maxLen && str.length > maxLen) {
    str = str.substring(0, maxLen);
  }
  if (/^[=+\-@\t\r]/.test(str)) {
    str = "'" + str;
  }
  return str;
}

function checkScriptRateLimit(bucketKey, maxRequests, windowSeconds) {
  try {
    const cache = CacheService.getScriptCache();
    const safeKey = 'rl_' + Utilities.base64EncodeWebSafe(String(bucketKey)).substring(0, 60);
    const currentRaw = cache.get(safeKey);
    const count = currentRaw ? (parseInt(currentRaw, 10) || 0) : 0;
    if (count >= maxRequests) {
      return false;
    }
    cache.put(safeKey, String(count + 1), windowSeconds);
    return true;
  } catch (e) {
    return true;
  }
}

function fetchFirestoreWithFallback(url, options, callerIdToken, preferServerOAuth) {
  const opts = Object.assign({ muteHttpExceptions: true }, options || {});
  const method = String((opts && opts.method) || 'get').toLowerCase();
  const isMutation = (method === 'patch' || method === 'post' || method === 'delete') && url.indexOf(':runQuery') === -1;

  if (preferServerOAuth || isMutation || !callerIdToken) {
    try {
      const oauthToken = ScriptApp.getOAuthToken();
      if (oauthToken) {
        opts.headers = Object.assign({}, opts.headers || {}, { Authorization: 'Bearer ' + oauthToken });
        const srvRes = UrlFetchApp.fetch(url, opts);
        const srvCode = srvRes.getResponseCode();
        if (srvCode >= 200 && srvCode < 300) return srvRes;
      }
    } catch (e) {}
  }

  if (callerIdToken) {
    opts.headers = Object.assign({}, opts.headers || {}, { Authorization: 'Bearer ' + callerIdToken });
    const res = UrlFetchApp.fetch(url, opts);
    const code = res.getResponseCode();
    if (code >= 200 && code < 300) return res;
  }

  try {
    const oauthToken = ScriptApp.getOAuthToken();
    if (oauthToken) {
      opts.headers = Object.assign({}, opts.headers || {}, { Authorization: 'Bearer ' + oauthToken });
      return UrlFetchApp.fetch(url, opts);
    }
  } catch (e) {}
  return null;
}

// ==========================================================================
// ENTRY POINTS (doGet & doPost)
// ==========================================================================

function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  // Nếu gọi trang GetKey / Xác nhận hoàn thành nhiệm vụ (tích hợp với getkey.gs)
  if (params.mode === 'getkey' || (params.taskId && typeof handleGetKeyPage === 'function')) {
    return handleGetKeyPage(e);
  }
  if (params.idToken || params.action) {
    return respond({
      error: 'Vì lý do bảo mật (chống lộ token trên URL log), mọi truy vấn API bắt buộc sử dụng phương thức POST.'
    }, 405);
  }
  return respond({ ok: true, status: 'LinkPro Unified Security API v2026 Ready (Users + Completions + Tasks)' });
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return respond({ error: 'Dữ liệu yêu cầu trống.' }, 400);
    }

    const body = JSON.parse(e.postData.contents);
    const action = String(body.action || '').trim();
    const idToken = String(body.idToken || '').trim();

    const verified = verifyIdToken(idToken);
    if (!verified.ok) {
      return respond({ error: verified.error }, 401);
    }
    if (body.uid && body.uid !== verified.uid) {
      return respond({ error: 'UID không khớp với phiên xác thực.' }, 403);
    }

    const rateLimitsByAction = {
      getBalance: { max: 30, win: 60 },
      getAllUsers: { max: 20, win: 60 },
      registerUser: { max: 5, win: 600 },
      gk_start: { max: 20, win: 60 },
      gk_claim: { max: 15, win: 60 },
      gk_save_config: { max: 20, win: 60 },
      gk_get_config: { max: 30, win: 60 },
      getUserTaskActivity: { max: 30, win: 60 },
      startTaskSession: { max: 20, win: 60 },
      completeTask: { max: 15, win: 60 },
      recordBypassAttempt: { max: 15, win: 60 },
      processWithdrawalRequest: { max: 3, win: 3600 },
      adjustBalance: { max: 20, win: 60 },
      toggleUserStatus: { max: 20, win: 60 },
      approveWithdrawal: { max: 20, win: 60 },
      rejectWithdrawal: { max: 20, win: 60 },
      getPendingCompletions: { max: 30, win: 60 },
      approveCompletion: { max: 30, win: 60 },
      rejectCompletion: { max: 30, win: 60 },
      deleteUserAccount: { max: 10, win: 60 },
      getAvailableTaskScripts: { max: 30, win: 60 },
      toggleTaskInSheet: { max: 30, win: 60 },
      saveTaskToSheet: { max: 30, win: 60 },
      deleteTaskFromSheet: { max: 20, win: 60 },
      getTasksFromSheet: { max: 30, win: 60 }
    };
    const rlCfg = rateLimitsByAction[action] || { max: 20, win: 60 };
    if (!checkScriptRateLimit('post_' + verified.uid + '_' + action, rlCfg.max, rlCfg.win)) {
      return respond({ error: 'Bạn đã gửi quá nhiều yêu cầu. Vui lòng thử lại sau!' }, 429);
    }

    if (action === 'registerUser') {
      return respond(handleRegisterUser(body, verified.uid, idToken));
    }

    // Nếu action bắt đầu bằng "gk_" (gk_start, gk_claim) -> gọi sang GetKey.gs
    if (action.indexOf('gk_') === 0 && typeof handleGetKey_ === 'function') {
      return respond(handleGetKey_(body, verified.uid, idToken));
    }

    if (action === 'getUserTaskActivity') {
      return respond(handleGetUserTaskActivity(verified.uid, body));
    }

    if (action === 'startTaskSession') {
      return respond(handleStartTaskSession(body, verified.uid));
    }

    // LUÔN ĐỒNG BỘ ROLE TRỰC TIẾP TỪ FIREBASE VỀ SHEET (KHÔNG DÙNG CACHE CŨ)
    // Khắc phục triệt để lỗi: "trong firebase đã chỉnh thành devweb nhưng sheet vẫn là user"
    const freshRole = syncRoleFromFirestore(verified.uid, idToken, true);

    if (action === 'getBalance') {
      return respond(handleGetBalance(body, verified.uid, freshRole, idToken));
    }

    if (action === 'getAllUsers') {
      return respond(handleGetAllUsers(body, verified.uid, idToken));
    }

    if (action === 'completeTask') {
      return respond(handleCompleteTask(body, verified.uid, idToken));
    }

    if (action === 'recordBypassAttempt') {
      return respond(handleRecordBypassAttempt(body, verified.uid, idToken));
    }

    if (action === 'processWithdrawalRequest') {
      return respond(handleProcessWithdrawal(body, verified.uid, idToken));
    }

    if (action === 'approveWithdrawal') {
      return respond(handleApproveWithdrawal(body, verified.uid, idToken));
    }

    if (action === 'rejectWithdrawal') {
      return respond(handleRejectWithdrawal(body, verified.uid, idToken));
    }

    if (action === 'adjustBalance') {
      return respond(handleAdjustBalance(body, verified.uid, idToken));
    }

    if (action === 'toggleUserStatus') {
      return respond(handleToggleUserStatus(body, verified.uid, idToken));
    }

    if (action === 'getPendingCompletions') {
      return respond(handleGetPendingCompletions(verified.uid, idToken));
    }
    if (action === 'approveCompletion') {
      return respond(handleReviewCompletion(body, verified.uid, 'approved', idToken));
    }
    if (action === 'rejectCompletion') {
      return respond(handleReviewCompletion(body, verified.uid, 'rejected', idToken));
    }

    if (action === 'deleteUserAccount') {
      return respond(handleDeleteUserAccount(body, verified.uid, idToken));
    }

    // Các action quản lý Thẻ Nhiệm Vụ từ các file (tên nhiệm vụ).gs trong Apps Script
    if (action === 'getAvailableTaskScripts') {
      return respond(handleGetAvailableTaskScripts(verified.uid, idToken));
    }
    if (action === 'toggleTaskInSheet') {
      return respond(handleToggleTaskInSheet(body, verified.uid, idToken));
    }
    if (action === 'saveTaskToSheet') {
      return respond(handleSaveTaskToSheet(body, verified.uid, idToken));
    }
    if (action === 'deleteTaskFromSheet') {
      return respond(handleDeleteTaskFromSheet(body, verified.uid, idToken));
    }
    if (action === 'getTasksFromSheet') {
      return respond(handleGetTasksFromSheet(verified.uid, idToken));
    }

    return respond({ error: 'Hành động không được hỗ trợ.' }, 400);
  } catch (err) {
    return respond({ error: 'Yêu cầu không hợp lệ hoặc lỗi xử lý máy chủ nội bộ.' }, 500);
  }
}

/**
 * XÁC MINH KÉP VAI TRÒ (ROLE) Ở CẢ FIREBASE VÀ GOOGLE SHEET KHI THỰC HIỆN THAO TÁC TỪ PANEL:
 * 1. Đọc trực tiếp role thật từ Firebase Firestore (`users/{callerUid}`, bypass cache).
 * 2. Tự động đồng bộ vào dòng của user trong tab `Users` trên Google Sheet nếu Sheet đang lệch.
 * 3. Kiểm tra cả `firebaseRole` VÀ `sheetRole` đều phải nằm trong `allowedRoles` (`['devweb']` hoặc `['admin', 'devweb']`).
 */
function verifyPanelRoleFirebaseAndSheet(callerUid, idToken, allowedRoles) {
  if (!isValidSafeId(callerUid)) {
    return { ok: false, error: 'UID người thực hiện không hợp lệ.' };
  }

  // Bước 1: Đọc trực tiếp từ Firebase Firestore (không dùng cache để luôn nhận role mới nhất)
  const fsProfile = fetchUserFromFirestore(callerUid, idToken, true);
  if (!fsProfile) {
    return { ok: false, error: 'Không tìm thấy hồ sơ quản trị viên trên Firebase Firestore.' };
  }
  if (fsProfile.isActive === false) {
    return { ok: false, error: 'Tài khoản quản trị viên đang bị khóa trên Firebase.' };
  }
  const firebaseRole = String(fsProfile.role || 'user').trim().toLowerCase();

  // Bước 2: Đảm bảo dòng của người gọi có trong tab Users của Sheet và đồng bộ role từ Firebase sang Sheet
  let userRow = findUserRow(callerUid);
  const sheet = ensureUsersSheet();
  if (!userRow) {
    const now = new Date();
    sheet.appendRow([
      sanitizeSheetCell(callerUid, 128),
      sanitizeSheetCell(fsProfile.username || callerUid.substring(0, 12), 30),
      sanitizeSheetCell(firebaseRole, 20),
      Number(fsProfile.balance) || 0,
      now,
      now
    ]);
    userRow = findUserRow(callerUid);
  } else {
    const currentSheetRole = String(userRow.values[USERS_COLS.role - 1] || '').trim().toLowerCase();
    if (currentSheetRole !== firebaseRole) {
      sheet.getRange(userRow.rowIndex, USERS_COLS.role).setValue(sanitizeSheetCell(firebaseRole, 20));
      sheet.getRange(userRow.rowIndex, USERS_COLS.updatedAt).setValue(new Date());
      userRow.values[USERS_COLS.role - 1] = firebaseRole;
    }
  }

  const sheetRole = userRow ? String(userRow.values[USERS_COLS.role - 1] || '').trim().toLowerCase() : '';

  // Bước 3: Xác minh cả Firebase Role và Sheet Role đều khớp và thuộc danh sách quyền cho phép
  const allowed = Array.isArray(allowedRoles) ? allowedRoles : ['devweb'];
  if (!allowed.includes(firebaseRole) || !allowed.includes(sheetRole) || firebaseRole !== sheetRole) {
    return {
      ok: false,
      error: 'Từ chối thao tác Panel: Quyền trên Firebase (' + firebaseRole + ') và Sheet (' + sheetRole + ') không đủ quyền (' + allowed.join('/') + ')!'
    };
  }

  return {
    ok: true,
    role: firebaseRole,
    username: fsProfile.username || ''
  };
}

/**
 * Trigger tự động khi Admin/Devweb chỉnh sửa trực tiếp trên Google Sheet:
 * - Tab `Completions`: Đổi cột `status` sang `approved` -> tự động cộng tiền vào ví user.
 * - Tab `Tasks` (Sheet Tạo Nhiệm Vụ): Nhập/sửa dòng nhiệm vụ trên Sheet -> tự động tạo `taskId`,
 *   rút gọn Link4m (nếu chưa có `taskLink`), tạo `secretKey` và đẩy nhiệm vụ lên Firebase Firestore!
 */
function onEditTrigger(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();

    // TRƯỜNG HỢP 1: Duyệt nhiệm vụ trong tab Completions
    if (sheetName === COMPLETIONS_SHEET_NAME) {
      if (e.range.getColumn() !== COMP_COLS.status) return;
      if (e.range.getNumRows() > 1) return;

      const newValue = String(e.value || '').trim().toLowerCase();
      const oldValue = String(e.oldValue || '').trim().toLowerCase();
      if (newValue !== 'approved' || oldValue === 'approved') return;

      const row = e.range.getRow();
      const rowValues = sheet.getRange(row, 1, 1, Object.keys(COMP_COLS).length).getValues()[0];
      const uid = String(rowValues[COMP_COLS.uid - 1] || '').trim();
      const username = String(rowValues[COMP_COLS.username - 1] || '').trim();
      const taskName = String(rowValues[COMP_COLS.taskName - 1] || 'Nhiệm vụ').trim();
      const reward = Number(rowValues[COMP_COLS.reward - 1]) || 0;

      if (!isValidSafeId(uid) || reward <= 0 || reward > 10000000) return;

      const lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        ensureUserRowInSheet(uid, username, null);
        adjustUserBalance(uid, reward, null);
        sheet.getRange(row, COMP_COLS.approvedAt).setValue(new Date());
        createFirestoreBalanceHistory(uid, username, 'credit', reward, 'Duyệt nhiệm vụ: ' + taskName, 'sheet_trigger', null);
      } finally {
        lock.releaseLock();
      }
      return;
    }

    // TRƯỜNG HỢP 2: Tạo / Sửa nhiệm vụ trực tiếp trong tab Tasks (Sheet Tạo Nhiệm Vụ)
    if (sheetName === TASKS_SHEET_NAME) {
      const row = e.range.getRow();
      if (row <= 1) return;
      syncSingleTaskRowFromSheetToFirestore(sheet, row, null);
    }
  } catch (err) {}
}

// ==========================================================================
// SHEET INITIALIZERS (TỰ ĐỘNG TẠO ĐỦ 3 TAB: Users, Completions, Tasks)
// ==========================================================================

function ensureUsersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(USERS_SHEET_NAME);
    sheet.appendRow(['uid', 'username', 'role', 'balance', 'createdAt', 'updatedAt']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  return sheet;
}

function ensureCompletionsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(COMPLETIONS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(COMPLETIONS_SHEET_NAME);
    sheet.appendRow(['completionId', 'uid', 'username', 'taskId', 'taskName', 'reward', 'type', 'status', 'createdAt', 'approvedAt', 'taskLink', 'ipAddress']);
    sheet.getRange(1, 1, 1, 12).setFontWeight('bold');
  }
  return sheet;
}

function ensureTasksSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TASKS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TASKS_SHEET_NAME);
    sheet.appendRow([
      'taskId',
      'title',
      'taskLink',
      'destinationUrl',
      'reward',
      'totalClicks',
      'remainingClicks',
      'maxPerUser',
      'verifySeconds',
      'approvalType',
      'isActive',
      'secretKey',
      'createdAt',
      'updatedAt'
    ]);
    sheet.getRange(1, 1, 1, 14).setFontWeight('bold');
  }
  return sheet;
}

function getSheet(name) {
  if (name === USERS_SHEET_NAME) return ensureUsersSheet();
  if (name === COMPLETIONS_SHEET_NAME) return ensureCompletionsSheet();
  if (name === TASKS_SHEET_NAME) return ensureTasksSheet();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Không tìm thấy bảng dữ liệu: ' + name);
  return sheet;
}

// ==========================================================================
// QUẢN LÝ SHEET TẠO NHIỆM VỤ (`Tasks` Tab) & ĐỒNG BỘ 2 CHIỀU VỚI FIRESTORE
// ==========================================================================

function findTaskRowInSheet(taskId) {
  if (!taskId) return null;
  const sheet = ensureTasksSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][TASKS_COLS.taskId - 1]).trim() === String(taskId).trim()) {
      return { rowIndex: i + 1, values: data[i] };
    }
  }
  return null;
}

function shortenLink4mOnServer(destinationUrl) {
  const cfg = getScriptConfig();
  const cleanDest = String(destinationUrl || '').trim();
  if (!cleanDest || cleanDest.indexOf('https://') !== 0) return '';
  try {
    const apiUrl = 'https://link4m.co/api-shorten/v2?api=' + encodeURIComponent(cfg.link4mToken) + '&url=' + encodeURIComponent(cleanDest);
    const res = UrlFetchApp.fetch(apiUrl, { method: 'get', muteHttpExceptions: true });
    if (res && res.getResponseCode() === 200) {
      const json = JSON.parse(res.getContentText());
      if (json && json.status === 'success' && json.shortenedUrl) {
        return String(json.shortenedUrl).replace(/\\\//g, '/').trim();
      }
    }
  } catch (e) {}
  return '';
}

function syncSingleTaskRowFromSheetToFirestore(sheet, rowIndex, idToken) {
  const rowVals = sheet.getRange(rowIndex, 1, 1, 14).getValues()[0];
  let taskId = String(rowVals[TASKS_COLS.taskId - 1] || '').trim();
  const title = String(rowVals[TASKS_COLS.title - 1] || '').trim();
  if (!title) return null;

  let taskLink = String(rowVals[TASKS_COLS.taskLink - 1] || '').trim();
  let destinationUrl = String(rowVals[TASKS_COLS.destinationUrl - 1] || '').trim();
  const reward = Math.max(50, Number(rowVals[TASKS_COLS.reward - 1]) || 500);
  const totalClicks = Math.max(1, Number(rowVals[TASKS_COLS.totalClicks - 1]) || 1000);
  let remainingClicks = rowVals[TASKS_COLS.remainingClicks - 1];
  remainingClicks = (remainingClicks === '' || remainingClicks === null || isNaN(Number(remainingClicks)))
    ? totalClicks
    : Math.max(0, Number(remainingClicks));
  const maxPerUser = Math.max(1, Number(rowVals[TASKS_COLS.maxPerUser - 1]) || 2);
  const verifySeconds = Math.max(5, Number(rowVals[TASKS_COLS.verifySeconds - 1]) || 75);
  const approvalType = String(rowVals[TASKS_COLS.approvalType - 1] || 'auto').trim().toLowerCase() === 'manual' ? 'manual' : 'auto';
  const isActiveRaw = rowVals[TASKS_COLS.isActive - 1];
  const isActive = (isActiveRaw === false || String(isActiveRaw).toLowerCase() === 'false' || String(isActiveRaw) === '0') ? false : true;
  let secretKey = String(rowVals[TASKS_COLS.secretKey - 1] || '').trim();
  const now = new Date();

  if (!taskId) {
    taskId = 'task_' + now.getTime() + '_' + Math.floor(Math.random() * 1000);
    sheet.getRange(rowIndex, TASKS_COLS.taskId).setValue(taskId);
  }
  if (!secretKey) {
    secretKey = 'LPK_' + Utilities.getUuid().replace(/-/g, '').substring(0, 16);
    sheet.getRange(rowIndex, TASKS_COLS.secretKey).setValue(secretKey);
  }
  if (!destinationUrl) {
    try {
      const webAppUrl = ScriptApp.getService().getUrl();
      if (webAppUrl) {
        destinationUrl = webAppUrl + '?mode=getkey&taskId=' + encodeURIComponent(taskId);
        sheet.getRange(rowIndex, TASKS_COLS.destinationUrl).setValue(destinationUrl);
      }
    } catch (e) {}
  }
  if (!taskLink && destinationUrl) {
    const shortUrl = shortenLink4mOnServer(destinationUrl);
    if (shortUrl) {
      taskLink = shortUrl;
      sheet.getRange(rowIndex, TASKS_COLS.taskLink).setValue(taskLink);
    }
  }
  if (!rowVals[TASKS_COLS.createdAt - 1]) {
    sheet.getRange(rowIndex, TASKS_COLS.createdAt).setValue(now);
  }
  sheet.getRange(rowIndex, TASKS_COLS.updatedAt).setValue(now);

  if (taskLink) {
    upsertTaskInFirestore({
      taskId: taskId,
      title: title,
      appScriptUrl: taskLink,
      destinationUrl: destinationUrl || taskLink,
      rewardAmount: reward,
      totalClicks: totalClicks,
      remainingClicks: remainingClicks,
      maxCompletionsPerUser: maxPerUser,
      verifySeconds: verifySeconds,
      approvalType: approvalType,
      isActive: isActive
    }, idToken);
  }

  return taskId;
}

function syncAllSheetTasksToFirestore() {
  const sheet = ensureTasksSheet();
  const lastRow = sheet.getLastRow();
  for (let r = 2; r <= lastRow; r++) {
    syncSingleTaskRowFromSheetToFirestore(sheet, r, null);
  }
}

function upsertTaskInFirestore(taskObj, idToken) {
  if (!taskObj || !isValidSafeId(taskObj.taskId)) return false;
  try {
    const cfg = getScriptConfig();
    const url = cfg.firestoreBaseUrl + '/tasks/' + encodeURIComponent(taskObj.taskId);
    const nowIso = new Date().toISOString();
    const payload = {
      fields: {
        title: { stringValue: sanitizeSheetCell(taskObj.title, 200) },
        appScriptUrl: { stringValue: String(taskObj.appScriptUrl || '').trim() },
        destinationUrl: { stringValue: String(taskObj.destinationUrl || taskObj.appScriptUrl || '').trim() },
        rewardAmount: { integerValue: String(Math.round(Number(taskObj.rewardAmount) || 500)) },
        totalClicks: { integerValue: String(Math.round(Number(taskObj.totalClicks) || 1000)) },
        remainingClicks: { integerValue: String(Math.round(Number(taskObj.remainingClicks) || 0)) },
        maxCompletionsPerUser: { integerValue: String(Math.max(1, Math.round(Number(taskObj.maxCompletionsPerUser) || 2))) },
        verifySeconds: { integerValue: String(Math.max(5, Math.round(Number(taskObj.verifySeconds) || 75))) },
        sessionExpiryMinutes: { integerValue: '30' },
        approvalType: { stringValue: taskObj.approvalType === 'manual' ? 'manual' : 'auto' },
        isActive: { booleanValue: taskObj.isActive !== false },
        updatedAt: { timestampValue: nowIso },
        createdAt: { timestampValue: nowIso }
      }
    };
    const res = fetchFirestoreWithFallback(url, {
      method: 'patch',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    }, idToken);
    return res && res.getResponseCode() === 200;
  } catch (e) {
    return false;
  }
}

function handleSaveTaskToSheet(body, callerUid, idToken) {
  const roleCheck = verifyPanelRoleFirebaseAndSheet(callerUid, idToken, ['devweb']);
  if (!roleCheck.ok) {
    return { error: roleCheck.error };
  }

  const mod = resolveTaskScriptModule_(body.scriptFile || body.taskId || '');
  const mc = (mod && mod.config) ? mod.config : null;

  const scriptFile = sanitizeSheetCell((mc && mc.SCRIPT_FILE) || body.scriptFile || 'link4m.gs', 100);
  const taskId = String((mc && mc.TASK_ID) || body.taskId || ('task_' + scriptFile.replace(/\.gs$/i, ''))).trim();
  if (!isValidSafeId(taskId)) {
    return { error: 'Mã nhiệm vụ không hợp lệ.' };
  }

  const title = sanitizeSheetCell((mc && mc.TITLE) || body.title || ('Nhiệm vụ ' + scriptFile), 200);
  const destinationUrl = sanitizeSheetCell((mc && mc.SITE_URL) || body.destinationUrl || 'https://linkprovietnam.web.app/', 500);
  let resolvedShortUrl = findLatestSessionShortUrl_('', taskId, false).shortUrl;
  if (!isRealShortUrl_(resolvedShortUrl) && mc && mc.SHORTENER_TEMPLATE) {
    const sampleDest = destinationUrl.replace(/\/+$/, '') + '/?taskId=' + encodeURIComponent(taskId) + '&s=' + encodeURIComponent(scriptFile);
    resolvedShortUrl = gkShorten_(sampleDest, mc.SHORTENER_TEMPLATE);
  }
  if (!isRealShortUrl_(resolvedShortUrl) && isRealShortUrl_(body.appScriptUrl)) {
    resolvedShortUrl = String(body.appScriptUrl).trim();
  }
  const taskLink = sanitizeSheetCell(resolvedShortUrl || 'https://link4m.org/', 500);
  const reward = Math.max(50, Math.round(Number((mc && mc.REWARD) || body.rewardAmount) || 500));
  const totalClicks = Math.max(1, Math.round(Number(body.totalClicks) || 999999999));
  const remainingClicks = body.remainingClicks !== undefined
    ? Math.max(0, Math.round(Number(body.remainingClicks)))
    : totalClicks;
  const maxPerUser = Math.max(1, Math.round(Number((mc && mc.MAX_PER_USER) || body.maxCompletionsPerUser) || 2));
  const verifySeconds = Math.max(5, Math.round(Number((mc && mc.MIN_SECONDS) || body.verifySeconds) || 75));
  const approvalType = ((mc && mc.APPROVAL_TYPE) || body.approvalType) === 'manual' ? 'manual' : 'auto';
  const isActive = body.isActive !== false;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = ensureTasksSheet();
    const existing = findTaskRowInSheet(taskId);
    const now = new Date();
    if (existing) {
      const r = existing.rowIndex;
      const prevKey = existing.values[TASKS_COLS.secretKey - 1] || ('LPK_' + Utilities.getUuid().replace(/-/g, '').substring(0, 16));
      const prevCreated = existing.values[TASKS_COLS.createdAt - 1] || now;
      sheet.getRange(r, 1, 1, 14).setValues([[
        taskId,
        title,
        taskLink,
        destinationUrl,
        reward,
        totalClicks,
        remainingClicks,
        maxPerUser,
        verifySeconds,
        approvalType,
        isActive,
        prevKey,
        prevCreated,
        now
      ]]);
    } else {
      const secretKey = 'LPK_' + Utilities.getUuid().replace(/-/g, '').substring(0, 16);
      sheet.appendRow([
        taskId,
        title,
        taskLink,
        destinationUrl,
        reward,
        totalClicks,
        remainingClicks,
        maxPerUser,
        verifySeconds,
        approvalType,
        isActive,
        secretKey,
        now,
        now
      ]);
    }
  } finally {
    lock.releaseLock();
  }

  if (isRealShortUrl_(taskLink)) {
    patchFirestoreTaskLink_(taskId, taskLink, idToken);
  }
  return { ok: true, taskId: taskId, shortUrl: taskLink, message: 'Đã đồng bộ nhiệm vụ vào tab Tasks trong Google Sheet!' };
}

function handleDeleteTaskFromSheet(body, callerUid, idToken) {
  const roleCheck = verifyPanelRoleFirebaseAndSheet(callerUid, idToken, ['devweb']);
  if (!roleCheck.ok) {
    return { error: roleCheck.error };
  }
  const taskId = String(body.taskId || '').trim();
  if (!isValidSafeId(taskId)) return { error: 'Mã nhiệm vụ không hợp lệ.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const existing = findTaskRowInSheet(taskId);
    if (existing) {
      ensureTasksSheet().deleteRow(existing.rowIndex);
    }
  } finally {
    lock.releaseLock();
  }
  return { ok: true, message: 'Đã xóa nhiệm vụ khỏi tab Tasks trong Sheet.' };
}

function handleGetTasksFromSheet(callerUid, idToken) {
  const roleCheck = verifyPanelRoleFirebaseAndSheet(callerUid, idToken, ['admin', 'devweb']);
  if (!roleCheck.ok) {
    return { error: roleCheck.error };
  }
  const sheet = ensureTasksSheet();
  const data = sheet.getDataRange().getValues();
  const tasks = [];
  for (let i = 1; i < data.length; i++) {
    const tid = String(data[i][TASKS_COLS.taskId - 1] || '').trim();
    if (!tid) continue;
    tasks.push({
      id: tid,
      title: String(data[i][TASKS_COLS.title - 1] || ''),
      appScriptUrl: String(data[i][TASKS_COLS.taskLink - 1] || ''),
      destinationUrl: String(data[i][TASKS_COLS.destinationUrl - 1] || ''),
      rewardAmount: Number(data[i][TASKS_COLS.reward - 1]) || 0,
      totalClicks: Number(data[i][TASKS_COLS.totalClicks - 1]) || 0,
      remainingClicks: Number(data[i][TASKS_COLS.remainingClicks - 1]) || 0,
      maxCompletionsPerUser: Math.max(1, Number(data[i][TASKS_COLS.maxPerUser - 1]) || 1),
      verifySeconds: Math.max(5, Number(data[i][TASKS_COLS.verifySeconds - 1]) || 75),
      approvalType: String(data[i][TASKS_COLS.approvalType - 1] || 'auto'),
      isActive: data[i][TASKS_COLS.isActive - 1] !== false && String(data[i][TASKS_COLS.isActive - 1]).toLowerCase() !== 'false'
    });
  }
  return { ok: true, tasks: tasks };
}

// ==========================================================================
// ACTION HANDLERS (UNIFIED LEDGER: SHEET + FIRESTORE + 30M SESSION)
// ==========================================================================

function handleStartTaskSession(body, callerUid) {
  const taskId = String(body.taskId || '').trim();
  if (!isValidSafeId(taskId)) {
    return { error: 'Mã nhiệm vụ không hợp lệ.' };
  }
  const nowMs = Date.now();
  const expiresAt = nowMs + TASK_SESSION_MAX_SECONDS * 1000;
  const sessionKey = 'tsess_' + callerUid + '_' + taskId;

  try {
    const cache = CacheService.getScriptCache();
    cache.put(sessionKey, JSON.stringify({
      startedAt: nowMs,
      expiresAt: expiresAt
    }), 21600);
  } catch (e) {}

  return {
    ok: true,
    taskId: taskId,
    startedAt: nowMs,
    expiresAt: expiresAt,
    maxSeconds: TASK_SESSION_MAX_SECONDS
  };
}

function ensureUserRowInSheet(uid, fallbackUsername, idToken) {
  let row = findUserRow(uid);
  const fsUser = fetchUserFromFirestore(uid, idToken, true);
  const uname = (fsUser && fsUser.username) ? fsUser.username : (fallbackUsername || uid.substring(0, 12));
  const role = (fsUser && fsUser.role) ? fsUser.role : (row ? String(row.values[USERS_COLS.role - 1] || 'user') : 'user');
  const bal = (fsUser && typeof fsUser.balance === 'number') ? fsUser.balance : 0;

  const sheet = ensureUsersSheet();
  const now = new Date();

  if (row) {
    // Nếu role hoặc username trên Firebase đã thay đổi (ví dụ chỉnh thành devweb), cập nhật ngay vào Sheet!
    if (fsUser && fsUser.role && String(row.values[USERS_COLS.role - 1]) !== String(fsUser.role)) {
      sheet.getRange(row.rowIndex, USERS_COLS.role).setValue(sanitizeSheetCell(fsUser.role, 20));
      sheet.getRange(row.rowIndex, USERS_COLS.updatedAt).setValue(now);
      row.values[USERS_COLS.role - 1] = fsUser.role;
    }
    return row;
  }

  sheet.appendRow([
    sanitizeSheetCell(uid, 128),
    sanitizeSheetCell(uname, 30),
    sanitizeSheetCell(role, 20),
    bal,
    now,
    now
  ]);
  return findUserRow(uid);
}

function handleGetBalance(body, callerUid, freshRole, idToken) {
  const targetUid = String(body.targetUid || callerUid).trim();
  if (!isValidSafeId(targetUid)) {
    return { error: 'Mã định danh không hợp lệ.' };
  }
  if (targetUid !== callerUid) {
    const check = verifyPanelRoleFirebaseAndSheet(callerUid, idToken, ['devweb']);
    if (!check.ok) return { error: check.error };
  }
  const row = ensureUserRowInSheet(targetUid, '', idToken);
  if (!row) return { balance: 0, role: freshRole || 'user', found: false };
  return {
    balance: Number(row.values[USERS_COLS.balance - 1]) || 0,
    role: String(row.values[USERS_COLS.role - 1] || freshRole || 'user'),
    found: true
  };
}

/**
 * Lấy danh sách toàn bộ người dùng trong Sheet ĐỒNG THỜI đồng bộ ngay Role (devweb / admin / user)
 * từ Firebase sang tab Users của Google Sheet để không bao giờ bị lệch Role giữa Firebase và Sheet!
 */
function handleGetAllUsers(body, callerUid, idToken) {
  const roleCheck = verifyPanelRoleFirebaseAndSheet(callerUid, idToken, ['admin', 'devweb']);
  if (!roleCheck.ok) {
    return { error: roleCheck.error };
  }

  const sheet = ensureUsersSheet();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // Đồng bộ trực tiếp danh sách user & role từ Firestore trên Server (tuyệt đối không tin role/balance do Client gửi lên)
    try {
      const cfg = getScriptConfig();
      const fsListRes = fetchFirestoreWithFallback(cfg.firestoreBaseUrl + '/users?pageSize=200', { method: 'get' }, idToken, true);
      if (fsListRes && fsListRes.getResponseCode() === 200) {
        const fsJson = JSON.parse(fsListRes.getContentText());
        const docs = (fsJson && Array.isArray(fsJson.documents)) ? fsJson.documents : [];
        if (docs.length > 0) {
          const data = sheet.getDataRange().getValues();
          const rowIndexByUid = {};
          for (let i = 1; i < data.length; i++) {
            const u = String(data[i][USERS_COLS.uid - 1] || '').trim();
            if (u) rowIndexByUid[u] = { rowIndex: i + 1, role: String(data[i][USERS_COLS.role - 1] || 'user').trim() };
          }
          const now = new Date();
          for (let j = 0; j < docs.length; j++) {
            const doc = docs[j];
            if (!doc || !doc.name || !doc.fields) continue;
            const parts = doc.name.split('/');
            const docUid = parts[parts.length - 1];
            if (!isValidSafeId(docUid)) continue;
            const f = doc.fields;
            const allowedRoles = ['user', 'admin', 'devweb'];
            const rawRole = (f.role && f.role.stringValue) ? String(f.role.stringValue).trim().toLowerCase() : 'user';
            const fRole = allowedRoles.indexOf(rawRole) !== -1 ? rawRole : 'user';
            const fName = (f.username && f.username.stringValue) ? String(f.username.stringValue).trim() : docUid.substring(0, 10);
            const fBal = Math.max(0, Number((f.balance && (f.balance.integerValue || f.balance.doubleValue)) || 0));
            const existingInfo = rowIndexByUid[docUid];
            if (!existingInfo) {
              sheet.appendRow([
                sanitizeSheetCell(docUid, 128),
                sanitizeSheetCell(fName, 30),
                sanitizeSheetCell(fRole, 20),
                fBal,
                now,
                now
              ]);
            } else if (existingInfo.role !== fRole) {
              sheet.getRange(existingInfo.rowIndex, USERS_COLS.role).setValue(sanitizeSheetCell(fRole, 20));
              sheet.getRange(existingInfo.rowIndex, USERS_COLS.updatedAt).setValue(now);
            }
          }
        }
      }
    } catch (syncErr) {}

    const updatedData = sheet.getDataRange().getValues();
    const list = [];
    for (let i = 1; i < updatedData.length; i++) {
      if (!updatedData[i][USERS_COLS.uid - 1]) continue;
      list.push({
        uid: String(updatedData[i][USERS_COLS.uid - 1]),
        username: String(updatedData[i][USERS_COLS.username - 1]),
        role: String(updatedData[i][USERS_COLS.role - 1] || 'user'),
        balance: Number(updatedData[i][USERS_COLS.balance - 1]) || 0
      });
    }
    return { users: list };
  } finally {
    lock.releaseLock();
  }
}

function handleRegisterUser(body, verifiedUid, idToken) {
  const username = String(body.username || '').trim().toLowerCase();
  if (!isValidSafeId(verifiedUid) || !/^[a-z0-9_]{4,20}$/.test(username)) {
    return { error: 'Thông tin định danh hoặc username không hợp lệ.' };
  }

  const existing = findUserRow(verifiedUid);
  if (existing) {
    return { ok: true, message: 'Tài khoản đã được đồng bộ.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const fsUser = fetchUserFromFirestore(verifiedUid, idToken, true);
    const initialRole = (fsUser && fsUser.role) ? fsUser.role : 'user';
    const sheet = ensureUsersSheet();
    const now = new Date();
    sheet.appendRow([
      sanitizeSheetCell(verifiedUid, 128),
      sanitizeSheetCell(username, 20),
      sanitizeSheetCell(initialRole, 20),
      0,
      now,
      now
    ]);
    syncBalanceToFirestore(verifiedUid, 0, idToken);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, message: 'Đã đồng bộ tài khoản vào hệ thống.' };
}

/**
 * Hoàn thành nhiệm vụ:
 * - Kiểm tra hạn 30 phút (1800s) và thời gian xác minh tối thiểu (`verifySeconds`).
 * - Kiểm tra giới hạn số lần vượt trên mỗi User (`maxCompletionsPerUser`).
 * - Nếu nhiệm vụ có kiểu duyệt là Tự Động (`approvalType === 'auto'`):
 *   Lập tức cộng tiền vào cả tab `Users` trong Google Sheet VÀ `users/{uid}.balance` trên Firestore!
 */
function handleCompleteTask(body, callerUid, idToken) {
  const taskId = String(body.taskId || '').trim();
  if (!isValidSafeId(taskId)) {
    return { error: 'Mã nhiệm vụ không hợp lệ.' };
  }

  const userProfile = fetchUserFromFirestore(callerUid, idToken, false);
  if (userProfile && userProfile.isActive === false) {
    return { error: 'Tài khoản của bạn đang bị khóa!' };
  }

  // KIỂM TRA PHIÊN NHIỆM VỤ TRÊN SERVER & HẠN 30 PHÚT (1800 GIÂY)
  const clientElapsedSeconds = Number(body.elapsedSeconds) || 0;
  const claimElapsedSeconds = Number(body.claimElapsedSeconds) || clientElapsedSeconds;
  const sessionKey = 'tsess_' + callerUid + '_' + taskId;
  let serverSessionExpired = false;
  let serverStartedAt = Number(body.serverStartedAt) || 0;
  let sessionShortUrl = isRealShortUrl_(body.sessionShortUrl) ? String(body.sessionShortUrl).trim() : '';

  try {
    const cache = CacheService.getScriptCache();
    const rawSess = cache.get(sessionKey);
    if (rawSess) {
      const sessObj = JSON.parse(rawSess);
      if (sessObj && sessObj.startedAt && !serverStartedAt) {
        serverStartedAt = Number(sessObj.startedAt) || 0;
      }
      if (sessObj && sessObj.shortUrl && isRealShortUrl_(sessObj.shortUrl) && !sessionShortUrl) {
        sessionShortUrl = String(sessObj.shortUrl).trim();
      }
      if (sessObj && sessObj.expiresAt && Date.now() > Number(sessObj.expiresAt)) {
        serverSessionExpired = true;
      }
    }
  } catch (e) {}

  // Luôn đối chiếu & cập nhật trạng thái trong tab Sessions để lấy đúng shortUrl (VD: https://link4m.org/go/qOrLB)
  const sessLookup = findLatestSessionShortUrl_(callerUid, taskId, true);
  if (sessLookup.shortUrl && isRealShortUrl_(sessLookup.shortUrl)) {
    sessionShortUrl = sessLookup.shortUrl;
  }
  if (!serverStartedAt && sessLookup.startedAt > 0) {
    serverStartedAt = sessLookup.startedAt;
  }
  if (sessLookup.expired) {
    serverSessionExpired = true;
  }

  if (!serverStartedAt && !body.serverVerifiedSession) {
    return { error: 'Không tìm thấy phiên làm nhiệm vụ hợp lệ trên máy chủ. Vui lòng bấm Tạo Link Nhiệm Vụ trước!' };
  }

  const verifiedElapsedSeconds = serverStartedAt > 0
    ? Math.max(1, Math.floor((Date.now() - serverStartedAt) / 1000))
    : clientElapsedSeconds;
  const effectiveElapsedSeconds = clientElapsedSeconds > 0
    ? Math.min(verifiedElapsedSeconds, clientElapsedSeconds)
    : verifiedElapsedSeconds;

  if (serverSessionExpired || verifiedElapsedSeconds > TASK_SESSION_MAX_SECONDS || clientElapsedSeconds > TASK_SESSION_MAX_SECONDS || claimElapsedSeconds > TASK_SESSION_MAX_SECONDS) {
    try {
      CacheService.getScriptCache().remove(sessionKey);
    } catch (e) {}
    return {
      ok: false,
      expired: true,
      revoked: true,
      error: 'Nhiệm vụ đã quá hạn 30 phút và đã được thu hồi! Kết quả không được gửi về Admin và không được cộng tiền.'
    };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // Đọc thông tin nhiệm vụ CHỈ TỪ SERVER (Firestore hoặc từ tab Tasks trên Sheet)
    // Tuyệt đối không tin reward, approvalType hay verifySeconds do Client gửi lên!
    let authoritativeTask = fetchTaskFromFirestore(taskId, idToken);
    const sheetTaskRow = findTaskRowInSheet(taskId);

    if (!authoritativeTask && sheetTaskRow) {
      const sv = sheetTaskRow.values;
      authoritativeTask = {
        title: String(sv[TASKS_COLS.title - 1] || 'Nhiệm vụ'),
        rewardAmount: Number(sv[TASKS_COLS.reward - 1]) || 0,
        approvalType: String(sv[TASKS_COLS.approvalType - 1] || 'auto') === 'manual' ? 'manual' : 'auto',
        isActive: sv[TASKS_COLS.isActive - 1] !== false && String(sv[TASKS_COLS.isActive - 1]).toLowerCase() !== 'false',
        remainingClicks: Number(sv[TASKS_COLS.remainingClicks - 1]) || 0,
        maxCompletionsPerUser: Math.max(1, Number(sv[TASKS_COLS.maxPerUser - 1]) || 1),
        verifySeconds: Math.max(5, Number(sv[TASKS_COLS.verifySeconds - 1]) || 75),
        appScriptUrl: String(sv[TASKS_COLS.taskLink - 1] || '')
      };
    }
    const modEntry = resolveTaskScriptModule_(body.scriptFile || (authoritativeTask && authoritativeTask.scriptFile) || taskId);
    if (modEntry && modEntry.config) {
      const mc = modEntry.config;
      if (!authoritativeTask) {
        authoritativeTask = {
          title: String(mc.TITLE || 'Nhiệm vụ'),
          rewardAmount: Number(mc.REWARD) || 500,
          approvalType: mc.APPROVAL_TYPE === 'manual' ? 'manual' : 'auto',
          isActive: true,
          remainingClicks: 999999999,
          maxCompletionsPerUser: Math.max(1, Number(mc.MAX_PER_USER) || 2),
          verifySeconds: Math.max(5, Number(mc.MIN_SECONDS) || 75),
          appScriptUrl: String(mc.SCRIPT_FILE || 'link4m.gs')
        };
      } else {
        authoritativeTask.verifySeconds = Math.max(5, Number(mc.MIN_SECONDS) || authoritativeTask.verifySeconds || 75);
        authoritativeTask.rewardAmount = Number(mc.REWARD) || authoritativeTask.rewardAmount;
        authoritativeTask.maxCompletionsPerUser = Math.max(1, Number(mc.MAX_PER_USER) || authoritativeTask.maxCompletionsPerUser || 2);
        authoritativeTask.approvalType = mc.APPROVAL_TYPE === 'manual' ? 'manual' : 'auto';
      }
    }

    if (!authoritativeTask) {
      return { error: 'Không tìm thấy nhiệm vụ hợp lệ trên hệ thống.' };
    }
    if (!authoritativeTask.isActive || authoritativeTask.remainingClicks <= 0) {
      return { error: 'Nhiệm vụ đã tạm dừng hoặc đã hết lượt thực hiện!' };
    }

    // Đã bỏ kiểm tra chặn giây cảnh báo Bypass theo yêu cầu

    const reward = Number(authoritativeTask.rewardAmount) || 0;
    if (reward <= 0 || reward > 10000000) {
      return { error: 'Mức thưởng nhiệm vụ không hợp lệ.' };
    }
    const isAuto = authoritativeTask.approvalType === 'auto';

    const maxPerUser = Math.max(1, Number(authoritativeTask.maxCompletionsPerUser) || 1);
    const sheet = ensureCompletionsSheet();
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    let userDoneCount = 0;

    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][COMP_COLS.uid - 1]) === callerUid && String(data[i][COMP_COLS.taskId - 1]) === taskId) {
        const rowStatus = String(data[i][COMP_COLS.status - 1] || '').toLowerCase();
        if (rowStatus !== 'rejected' && rowStatus !== 'bypass_alert') {
          userDoneCount += 1;
        }
      }
    }

    if (userDoneCount >= maxPerUser) {
      return { error: 'Bạn đã vượt đủ tối đa ' + maxPerUser + '/' + maxPerUser + ' lần cho nhiệm vụ này!' };
    }

    const clientIp = normalizeClientIp_(body.clientIp || sessLookup.ipAddress || '');
    const maxPerIp = Math.max(1, Number((modEntry && modEntry.config && modEntry.config.MAX_PER_IP) || authoritativeTask.maxCompletionsPerIp) || 1);
    if (clientIp) {
      const ipDoneCount = countCompletedByIpForTask_(clientIp, taskId);
      if (ipDoneCount >= maxPerIp) {
        return { error: 'Địa chỉ IP (' + clientIp + ') của bạn đã vượt đủ ' + maxPerIp + '/' + maxPerIp + ' lần cho nhiệm vụ này (Luật: 1 IP / 1 lần vượt)!' };
      }
    }

    try {
      CacheService.getScriptCache().remove(sessionKey);
    } catch (e) {}

    const finalTaskLink = isRealShortUrl_(sessionShortUrl)
      ? sessionShortUrl
      : (isRealShortUrl_(authoritativeTask.appScriptUrl) ? authoritativeTask.appScriptUrl : (findLatestSessionShortUrl_('', taskId, false).shortUrl || authoritativeTask.appScriptUrl));

    // 1. Giảm remainingClicks (-1) & cập nhật luôn cột taskLink (shortUrl) trên cả Firestore & tab Tasks trong Sheet
    const nextClicks = Math.max(0, authoritativeTask.remainingClicks - 1);
    decrementTaskClicksInFirestore(taskId, nextClicks, idToken);
    if (isRealShortUrl_(finalTaskLink)) {
      patchFirestoreTaskLink_(taskId, finalTaskLink, idToken);
    }
    if (sheetTaskRow) {
      const tSheet = ensureTasksSheet();
      if (isRealShortUrl_(finalTaskLink)) {
        tSheet.getRange(sheetTaskRow.rowIndex, TASKS_COLS.taskLink).setValue(finalTaskLink);
      }
      tSheet.getRange(sheetTaskRow.rowIndex, TASKS_COLS.remainingClicks).setValue(nextClicks);
      tSheet.getRange(sheetTaskRow.rowIndex, TASKS_COLS.updatedAt).setValue(now);
    }

    const cleanUsername = (userProfile && userProfile.username) ? userProfile.username : sanitizeSheetCell(body.username || '', 20);
    ensureUserRowInSheet(callerUid, cleanUsername, idToken);

    const completionId = Utilities.getUuid();
    const statusStr = isAuto ? 'approved' : 'pending';

    // 2. Ghi vào Sheet Completions (lưu cả shortUrl ở cột 11 và ipAddress ở cột 12)
    sheet.appendRow([
      completionId,
      sanitizeSheetCell(callerUid, 128),
      sanitizeSheetCell(cleanUsername, 30),
      sanitizeSheetCell(taskId, 128),
      sanitizeSheetCell(authoritativeTask.title, 200),
      reward,
      isAuto ? 'auto' : 'manual',
      statusStr,
      now,
      isAuto ? now : '',
      sanitizeSheetCell(finalTaskLink, 500),
      sanitizeSheetCell(clientIp, 45)
    ]);

    // 3. Ghi bản ghi taskCompletions lên Firestore
    createFirestoreTaskCompletion(completionId, callerUid, taskId, authoritativeTask.title, reward, isAuto ? 'auto' : 'manual', statusStr, idToken);

    // 4. Nếu duyệt Tự Động (auto): Cộng số dư NGAY LẬP TỨC trên cả Sheet & Firestore!
    let newBalance = null;
    if (isAuto) {
      newBalance = adjustUserBalance(callerUid, reward, idToken);
      createFirestoreBalanceHistory(callerUid, cleanUsername, 'credit', reward, 'Hoàn thành nhiệm vụ tự động: ' + authoritativeTask.title, 'system', idToken);
    }

    return {
      ok: true,
      taskId: taskId,
      taskTitle: authoritativeTask.title,
      taskLink: finalTaskLink,
      status: statusStr,
      rewardAmount: reward,
      newBalance: newBalance,
      userDoneCount: userDoneCount + 1,
      maxCompletionsPerUser: maxPerUser,
      message: isAuto
        ? 'Hoàn thành nhiệm vụ! +' + reward.toLocaleString('vi-VN') + ' VND đã được cộng tự động vào ví!'
        : 'Đã ghi nhận hoàn thành, đang chờ quản trị viên xét duyệt.'
    };
  } finally {
    lock.releaseLock();
  }
}

function handleProcessWithdrawal(body, callerUid, idToken) {
  const numAmount = Number(body.amount);
  const method = String(body.method || '').trim();
  const cleanAccNum = String(body.accountNumber || '').trim();
  const cleanAccName = sanitizeSheetCell(String(body.accountName || '').trim().toUpperCase(), 100);
  const cleanBankName = method === 'bank' ? sanitizeSheetCell(String(body.bankName || '').trim(), 100) : 'Momo';
  const cleanNote = sanitizeSheetCell(body.note || '', 200);

  if (!numAmount || isNaN(numAmount) || numAmount < 20000 || numAmount > 50000000) {
    return { error: 'Số tiền rút phải từ 20.000 VND đến 50.000.000 VND!' };
  }
  if (!['momo', 'bank'].includes(method)) {
    return { error: 'Phương thức rút tiền không hợp lệ!' };
  }
  if (!/^[0-9]{6,20}$/.test(cleanAccNum)) {
    return { error: 'Số tài khoản / số điện thoại chỉ được chứa từ 6 đến 20 chữ số!' };
  }
  if (!cleanAccName || cleanAccName.length < 3) {
    return { error: 'Tên chủ tài khoản không hợp lệ!' };
  }
  if (method === 'bank' && (!cleanBankName || cleanBankName.length < 2)) {
    return { error: 'Vui lòng chọn ngân hàng thụ hưởng!' };
  }

  const userProfile = fetchUserFromFirestore(callerUid, idToken, false);
  if (userProfile && userProfile.isActive === false) {
    return { error: 'Tài khoản của bạn đang bị khóa!' };
  }

  if (hasPendingWithdrawalInFirestore(callerUid, idToken)) {
    return { error: 'Bạn đang có một đơn rút tiền chờ xử lý. Vui lòng đợi admin duyệt!' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const cleanUsername = (userProfile && userProfile.username) ? userProfile.username : '';
    const userRow = ensureUserRowInSheet(callerUid, cleanUsername, idToken);
    if (!userRow) {
      return { error: 'Không tìm thấy hồ sơ số dư của tài khoản!' };
    }
    const currentBalance = Number(userRow.values[USERS_COLS.balance - 1]) || 0;
    if (currentBalance < numAmount) {
      return { error: 'Số dư khả dụng không đủ để thực hiện giao dịch rút tiền!' };
    }

    adjustUserBalance(callerUid, -numAmount, idToken);

    const withdrawalId = Utilities.getUuid();
    const finalUsername = cleanUsername || String(userRow.values[USERS_COLS.username - 1] || '');
    createFirestoreWithdrawal(withdrawalId, callerUid, finalUsername, numAmount, method, cleanAccNum, cleanAccName, cleanBankName, cleanNote, idToken);

    createFirestoreBalanceHistory(
      callerUid,
      finalUsername,
      'withdraw',
      -numAmount,
      'Yêu cầu rút tiền qua ' + (method === 'momo' ? 'MoMo' : 'Ngân hàng (' + cleanBankName + ')'),
      'system',
      idToken
    );

    return {
      ok: true,
      success: true,
      withdrawalId: withdrawalId,
      amount: numAmount,
      message: 'Đơn rút tiền đã được gửi, chờ admin xử lý (1-24h).'
    };
  } finally {
    lock.releaseLock();
  }
}

function handleApproveWithdrawal(body, callerUid, idToken) {
  const roleCheck = verifyPanelRoleFirebaseAndSheet(callerUid, idToken, ['admin', 'devweb']);
  if (!roleCheck.ok) {
    return { error: roleCheck.error };
  }
  const withdrawalId = String(body.withdrawalId || '').trim();
  if (!isValidSafeId(withdrawalId)) {
    return { error: 'Mã đơn rút không hợp lệ.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const wd = fetchWithdrawalFromFirestore(withdrawalId, idToken);
    if (!wd) return { error: 'Không tìm thấy đơn rút tiền!' };
    if (wd.status !== 'pending') return { error: 'Đơn rút tiền này đã được xử lý trước đó!' };

    updateFirestoreWithdrawalStatus(withdrawalId, 'done', '', callerUid, idToken);
    writeFirestoreAuditLog(callerUid, 'APPROVE_WITHDRAWAL', withdrawalId, {
      userId: wd.userId,
      amount: wd.amount,
      method: wd.method
    }, idToken);

    return { ok: true, success: true, message: 'Đã duyệt đơn rút tiền thành công!' };
  } finally {
    lock.releaseLock();
  }
}

function handleRejectWithdrawal(body, callerUid, idToken) {
  const roleCheck = verifyPanelRoleFirebaseAndSheet(callerUid, idToken, ['admin', 'devweb']);
  if (!roleCheck.ok) {
    return { error: roleCheck.error };
  }
  const withdrawalId = String(body.withdrawalId || '').trim();
  const cleanReason = sanitizeSheetCell(body.reason || '', 200);
  if (!isValidSafeId(withdrawalId) || !cleanReason) {
    return { error: 'Mã đơn rút hoặc lý do từ chối không hợp lệ.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const wd = fetchWithdrawalFromFirestore(withdrawalId, idToken);
    if (!wd) return { error: 'Không tìm thấy đơn rút tiền!' };
    if (wd.status !== 'pending') return { error: 'Đơn rút tiền này đã được xử lý trước đó!' };

    updateFirestoreWithdrawalStatus(withdrawalId, 'rejected', cleanReason, callerUid, idToken);

    ensureUserRowInSheet(wd.userId, wd.username || '', idToken);
    adjustUserBalance(wd.userId, wd.amount, idToken);

    createFirestoreBalanceHistory(
      wd.userId,
      wd.username || '',
      'credit',
      wd.amount,
      'Hoàn tiền đơn rút bị từ chối: ' + cleanReason,
      callerUid,
      idToken
    );

    writeFirestoreAuditLog(callerUid, 'REJECT_WITHDRAWAL', withdrawalId, {
      userId: wd.userId,
      amount: wd.amount,
      reason: cleanReason
    }, idToken);

    return { ok: true, success: true, message: 'Đã từ chối đơn rút và hoàn tiền vào ví người dùng!' };
  } finally {
    lock.releaseLock();
  }
}

function handleAdjustBalance(body, callerUid, idToken) {
  const roleCheck = verifyPanelRoleFirebaseAndSheet(callerUid, idToken, ['devweb']);
  if (!roleCheck.ok) {
    return { error: roleCheck.error };
  }
  const targetUid = String(body.targetUid || '').trim();
  const amount = Number(body.amount);
  if (!isValidSafeId(targetUid) || isNaN(amount) || amount === 0 || Math.abs(amount) > 1000000000) {
    return { error: 'Tham số điều chỉnh số dư không hợp lệ.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const targetRow = ensureUserRowInSheet(targetUid, body.targetUsername || '', idToken);
    if (!targetRow) {
      return { error: 'Không tìm thấy tài khoản đích.' };
    }
    const currentBalance = Number(targetRow.values[USERS_COLS.balance - 1]) || 0;
    if (amount < 0 && currentBalance + amount < 0) {
      return { error: 'Số dư hiện tại không đủ để trừ.' };
    }

    const nextBalance = adjustUserBalance(targetUid, amount, idToken);

    const compSheet = ensureCompletionsSheet();
    const now = new Date();
    const cleanUsername = String(targetRow.values[USERS_COLS.username - 1] || '');
    const cleanNote = sanitizeSheetCell(body.note || (amount > 0 ? 'Cộng tiền thủ công' : 'Trừ tiền thủ công'), 200);
    compSheet.appendRow([
      Utilities.getUuid(),
      sanitizeSheetCell(targetUid, 128),
      sanitizeSheetCell(cleanUsername, 30),
      '',
      cleanNote,
      amount,
      'manual_adjust',
      'approved',
      now,
      now,
      ''
    ]);

    createFirestoreBalanceHistory(
      targetUid,
      cleanUsername,
      amount > 0 ? 'credit' : 'debit',
      amount,
      cleanNote,
      callerUid,
      idToken
    );

    writeFirestoreAuditLog(
      callerUid,
      amount > 0 ? 'CREDIT_USER' : 'DEBIT_USER',
      targetUid,
      { amount: amount, note: cleanNote, newBalance: nextBalance },
      idToken
    );

    return { ok: true, newBalance: nextBalance, message: 'Đã cập nhật số dư thành công!' };
  } finally {
    lock.releaseLock();
  }
}

function handleToggleUserStatus(body, callerUid, idToken) {
  const roleCheck = verifyPanelRoleFirebaseAndSheet(callerUid, idToken, ['devweb']);
  if (!roleCheck.ok) {
    return { error: roleCheck.error };
  }
  const targetUid = String(body.targetUid || '').trim();
  if (!isValidSafeId(targetUid)) {
    return { error: 'UID mục tiêu không hợp lệ.' };
  }
  if (targetUid === callerUid) {
    return { error: 'Không thể tự khóa tài khoản của chính mình.' };
  }

  const targetUser = fetchUserFromFirestore(targetUid, idToken, true);
  if (!targetUser) {
    return { error: 'Không tìm thấy người dùng trên hệ thống.' };
  }
  if (targetUser.role === 'devweb') {
    return { error: 'Không thể khóa tài khoản Devweb.' };
  }

  const newStatus = !(targetUser.isActive !== false);
  const updated = patchFirestoreUserField(targetUid, 'isActive', { booleanValue: newStatus }, idToken);
  if (!updated) {
    return { error: 'Không thể cập nhật trạng thái tài khoản.' };
  }

  writeFirestoreAuditLog(callerUid, 'TOGGLE_USER_STATUS', targetUid, { newStatus: newStatus }, idToken);
  return {
    ok: true,
    isActive: newStatus,
    message: newStatus ? 'Đã mở khóa tài khoản thành công!' : 'Đã khóa tài khoản thành công!'
  };
}

// ==========================================================================
// HELPERS & FIRESTORE REST SYNC
// ==========================================================================

function findUserRow(uid) {
  if (!isValidSafeId(uid)) return null;
  const sheet = ensureUsersSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][USERS_COLS.uid - 1]).trim() === String(uid).trim()) {
      return { rowIndex: i + 1, values: data[i] };
    }
  }
  return null;
}

function adjustUserBalance(uid, delta, idToken) {
  const sheet = ensureUsersSheet();
  let row = findUserRow(uid);
  if (!row) {
    row = ensureUserRowInSheet(uid, '', idToken);
  }
  if (!row) throw new Error('Không tìm thấy tài khoản.');

  const current = Number(row.values[USERS_COLS.balance - 1]) || 0;
  const next = Math.max(0, current + delta);
  sheet.getRange(row.rowIndex, USERS_COLS.balance).setValue(next);
  sheet.getRange(row.rowIndex, USERS_COLS.updatedAt).setValue(new Date());

  syncBalanceToFirestore(uid, next, idToken);
  return next;
}

function syncBalanceToFirestore(uid, balanceVal, idToken) {
  patchFirestoreUserField(uid, 'balance', { integerValue: String(Math.round(Number(balanceVal) || 0)) }, idToken);
}

function patchFirestoreUserField(uid, fieldName, fieldValueObj, idToken) {
  if (!isValidSafeId(uid)) return false;
  try {
    CacheService.getScriptCache().remove('uprof_' + uid);
  } catch (e) {}
  try {
    const cfg = getScriptConfig();
    const url = cfg.firestoreBaseUrl + '/users/' + encodeURIComponent(uid) + '?updateMask.fieldPaths=' + encodeURIComponent(fieldName);
    const fields = {};
    fields[fieldName] = fieldValueObj;
    const res = fetchFirestoreWithFallback(url, {
      method: 'patch',
      contentType: 'application/json',
      payload: JSON.stringify({ fields: fields })
    }, idToken);
    return res && res.getResponseCode() === 200;
  } catch (e) {
    return false;
  }
}

function fetchUserFromFirestore(uid, idToken, bypassCache) {
  if (!isValidSafeId(uid)) return null;
  const cacheKey = 'uprof_' + uid;
  if (!bypassCache) {
    try {
      const cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) {
        return JSON.parse(cached);
      }
    } catch (e) {}
  }

  try {
    const cfg = getScriptConfig();
    const res = fetchFirestoreWithFallback(cfg.firestoreBaseUrl + '/users/' + encodeURIComponent(uid), {
      method: 'get'
    }, idToken);
    if (!res || res.getResponseCode() !== 200) return null;
    const doc = JSON.parse(res.getContentText());
    if (!doc || !doc.fields) return null;
    const f = doc.fields;
    const profileObj = {
      uid: uid,
      username: (f.username && f.username.stringValue) ? f.username.stringValue : '',
      role: (f.role && f.role.stringValue) ? f.role.stringValue : 'user',
      isActive: f.isActive && f.isActive.booleanValue === false ? false : true,
      balance: Number((f.balance && (f.balance.integerValue || f.balance.doubleValue)) || 0)
    };
    try {
      CacheService.getScriptCache().put(cacheKey, JSON.stringify(profileObj), 120);
    } catch (e) {}
    return profileObj;
  } catch (e) {
    return null;
  }
}

function fetchTaskFromFirestore(taskId, idToken) {
  if (!isValidSafeId(taskId)) return null;
  try {
    const cfg = getScriptConfig();
    const res = fetchFirestoreWithFallback(cfg.firestoreBaseUrl + '/tasks/' + encodeURIComponent(taskId), {
      method: 'get'
    }, idToken);
    if (!res || res.getResponseCode() !== 200) return null;
    const doc = JSON.parse(res.getContentText());
    if (!doc || !doc.fields) return null;
    const f = doc.fields;
    return {
      title: (f.title && f.title.stringValue) ? f.title.stringValue : 'Nhiệm vụ',
      rewardAmount: Number((f.rewardAmount && (f.rewardAmount.integerValue || f.rewardAmount.doubleValue)) || 0),
      approvalType: (f.approvalType && f.approvalType.stringValue) ? f.approvalType.stringValue : 'auto',
      isActive: f.isActive && f.isActive.booleanValue === true,
      remainingClicks: Number((f.remainingClicks && (f.remainingClicks.integerValue || f.remainingClicks.doubleValue)) || 0),
      maxCompletionsPerUser: Math.max(1, Number((f.maxCompletionsPerUser && (f.maxCompletionsPerUser.integerValue || f.maxCompletionsPerUser.doubleValue)) || 1)),
      verifySeconds: Number((f.verifySeconds && (f.verifySeconds.integerValue || f.verifySeconds.doubleValue)) || 75),
      sessionExpiryMinutes: Number((f.sessionExpiryMinutes && (f.sessionExpiryMinutes.integerValue || f.sessionExpiryMinutes.doubleValue)) || 30),
      scriptFile: (f.scriptFile && f.scriptFile.stringValue) ? f.scriptFile.stringValue : '',
      shortenerTemplate: (f.shortenerTemplate && f.shortenerTemplate.stringValue) ? f.shortenerTemplate.stringValue : '',
      appScriptUrl: (f.appScriptUrl && f.appScriptUrl.stringValue) ? f.appScriptUrl.stringValue : ''
    };
  } catch (e) {
    return null;
  }
}

function decrementTaskClicksInFirestore(taskId, nextClicksVal, idToken) {
  if (!isValidSafeId(taskId)) return false;
  try {
    const cfg = getScriptConfig();
    const url = cfg.firestoreBaseUrl + '/tasks/' + encodeURIComponent(taskId) + '?updateMask.fieldPaths=remainingClicks';
    const payload = {
      fields: {
        remainingClicks: { integerValue: String(Math.max(0, Math.round(Number(nextClicksVal) || 0))) }
      }
    };
    const res = fetchFirestoreWithFallback(url, {
      method: 'patch',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    }, idToken);
    return res && res.getResponseCode() === 200;
  } catch (e) {
    return false;
  }
}

function createFirestoreTaskCompletion(completionId, userId, taskId, taskTitle, rewardAmount, approvalType, statusStr, idToken) {
  try {
    const cfg = getScriptConfig();
    const url = cfg.firestoreBaseUrl + '/taskCompletions?documentId=' + encodeURIComponent(completionId);
    const nowIso = new Date().toISOString();
    const payload = {
      fields: {
        userId: { stringValue: String(userId) },
        taskId: { stringValue: String(taskId) },
        taskTitle: { stringValue: sanitizeSheetCell(taskTitle, 150) },
        rewardAmount: { integerValue: String(Math.round(Number(rewardAmount) || 0)) },
        approvalType: { stringValue: String(approvalType) },
        status: { stringValue: String(statusStr) },
        completedAt: { timestampValue: nowIso }
      }
    };
    fetchFirestoreWithFallback(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    }, idToken);
  } catch (e) {}
}

function createFirestoreBalanceHistory(userId, username, typeStr, amountVal, noteStr, createdByStr, idToken) {
  try {
    const cfg = getScriptConfig();
    const url = cfg.firestoreBaseUrl + '/balanceHistory';
    const nowIso = new Date().toISOString();
    const payload = {
      fields: {
        userId: { stringValue: String(userId) },
        username: { stringValue: sanitizeSheetCell(username, 30) },
        type: { stringValue: String(typeStr) },
        amount: { integerValue: String(Math.round(Number(amountVal) || 0)) },
        note: { stringValue: sanitizeSheetCell(noteStr, 200) },
        createdAt: { timestampValue: nowIso },
        createdBy: { stringValue: sanitizeSheetCell(createdByStr || 'system', 60) }
      }
    };
    fetchFirestoreWithFallback(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    }, idToken);
  } catch (e) {}
}

function hasPendingWithdrawalInFirestore(userId, idToken) {
  if (!isValidSafeId(userId)) return false;
  try {
    const cfg = getScriptConfig();
    const queryUrl = 'https://firestore.googleapis.com/v1/' + cfg.firestoreRootName + ':runQuery';
    const payload = {
      structuredQuery: {
        from: [{ collectionId: 'withdrawals' }],
        where: {
          compositeFilter: {
            op: 'AND',
            filters: [
              {
                fieldFilter: {
                  field: { fieldPath: 'userId' },
                  op: 'EQUAL',
                  value: { stringValue: userId }
                }
              },
              {
                fieldFilter: {
                  field: { fieldPath: 'status' },
                  op: 'EQUAL',
                  value: { stringValue: 'pending' }
                }
              }
            ]
          }
        },
        limit: 1
      }
    };
    const res = fetchFirestoreWithFallback(queryUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    }, idToken);
    if (!res || res.getResponseCode() !== 200) return false;
    const arr = JSON.parse(res.getContentText());
    return Array.isArray(arr) && arr.some(item => item && item.document);
  } catch (e) {
    return false;
  }
}

function createFirestoreWithdrawal(withdrawalId, userId, username, amountVal, method, accNum, accName, bankName, note, idToken) {
  const cfg = getScriptConfig();
  const url = cfg.firestoreBaseUrl + '/withdrawals?documentId=' + encodeURIComponent(withdrawalId);
  const nowIso = new Date().toISOString();
  const payload = {
    fields: {
      userId: { stringValue: String(userId) },
      username: { stringValue: sanitizeSheetCell(username, 30) },
      amount: { integerValue: String(Math.round(Number(amountVal) || 0)) },
      method: { stringValue: String(method) },
      accountNumber: { stringValue: String(accNum) },
      accountName: { stringValue: sanitizeSheetCell(accName, 100) },
      bankName: { stringValue: sanitizeSheetCell(bankName, 100) },
      note: { stringValue: sanitizeSheetCell(note, 200) },
      status: { stringValue: 'pending' },
      createdAt: { timestampValue: nowIso }
    }
  };
  fetchFirestoreWithFallback(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  }, idToken);
}

function fetchWithdrawalFromFirestore(withdrawalId, idToken) {
  if (!isValidSafeId(withdrawalId)) return null;
  try {
    const cfg = getScriptConfig();
    const res = fetchFirestoreWithFallback(cfg.firestoreBaseUrl + '/withdrawals/' + encodeURIComponent(withdrawalId), {
      method: 'get'
    }, idToken);
    if (!res || res.getResponseCode() !== 200) return null;
    const doc = JSON.parse(res.getContentText());
    if (!doc || !doc.fields) return null;
    const f = doc.fields;
    return {
      id: withdrawalId,
      userId: (f.userId && f.userId.stringValue) ? f.userId.stringValue : '',
      username: (f.username && f.username.stringValue) ? f.username.stringValue : '',
      amount: Number((f.amount && (f.amount.integerValue || f.amount.doubleValue)) || 0),
      method: (f.method && f.method.stringValue) ? f.method.stringValue : 'momo',
      status: (f.status && f.status.stringValue) ? f.status.stringValue : 'pending'
    };
  } catch (e) {
    return null;
  }
}

function updateFirestoreWithdrawalStatus(withdrawalId, newStatus, rejectReason, processedByUid, idToken) {
  const cfg = getScriptConfig();
  const mask = ['updateMask.fieldPaths=status', 'updateMask.fieldPaths=processedAt', 'updateMask.fieldPaths=processedBy'];
  const fields = {
    status: { stringValue: String(newStatus) },
    processedAt: { timestampValue: new Date().toISOString() },
    processedBy: { stringValue: String(processedByUid || '') }
  };
  if (rejectReason) {
    mask.push('updateMask.fieldPaths=rejectReason');
    fields.rejectReason = { stringValue: sanitizeSheetCell(rejectReason, 200) };
  }
  const url = cfg.firestoreBaseUrl + '/withdrawals/' + encodeURIComponent(withdrawalId) + '?' + mask.join('&');
  fetchFirestoreWithFallback(url, {
    method: 'patch',
    contentType: 'application/json',
    payload: JSON.stringify({ fields: fields })
  }, idToken);
}

function writeFirestoreAuditLog(adminUid, actionStr, targetStr, detailsObj, idToken) {
  try {
    const cfg = getScriptConfig();
    const url = cfg.firestoreBaseUrl + '/auditLogs';
    const payload = {
      fields: {
        adminUid: { stringValue: String(adminUid || '') },
        action: { stringValue: sanitizeSheetCell(actionStr, 60) },
        target: { stringValue: sanitizeSheetCell(targetStr, 128) },
        detailsJson: { stringValue: JSON.stringify(detailsObj || {}) },
        timestamp: { timestampValue: new Date().toISOString() }
      }
    };
    fetchFirestoreWithFallback(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    }, idToken);
  } catch (e) {}
}

function handleDeleteUserAccount(body, callerUid, idToken) {
  const roleCheck = verifyPanelRoleFirebaseAndSheet(callerUid, idToken, ['devweb']);
  if (!roleCheck.ok) {
    return { error: roleCheck.error };
  }
  const targetUid = String(body.targetUid || '').trim();
  if (!isValidSafeId(targetUid)) return { error: 'Mã tài khoản không hợp lệ.' };
  if (targetUid === callerUid) return { error: 'Không thể tự xóa tài khoản của chính mình.' };

  const row = findUserRow(targetUid);
  const username = row ? String(row.values[USERS_COLS.username - 1] || '').trim() : null;
  if (row && String(row.values[USERS_COLS.role - 1] || '').toLowerCase() === 'devweb') {
    return { error: 'Không thể xóa một tài khoản Devweb khác.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    try {
      CacheService.getScriptCache().remove('uprof_' + targetUid);
    } catch (e) {}
    const authDeleted = deleteFirebaseAuthUser(targetUid);
    deleteFirestoreDoc('users/' + encodeURIComponent(targetUid), idToken);
    if (username && /^[a-z0-9_]{4,20}$/.test(username)) {
      deleteFirestoreDoc('usernames/' + encodeURIComponent(username), idToken);
    }
    if (row) ensureUsersSheet().deleteRow(row.rowIndex);

    writeFirestoreAuditLog(callerUid, 'DELETE_USER', targetUid, { authDeleted: authDeleted }, idToken);

    return {
      ok: true,
      message: authDeleted
        ? 'Đã xóa vĩnh viễn tài khoản khỏi hệ thống.'
        : 'Đã xóa dữ liệu trên Firestore và Sheet.'
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteFirebaseAuthUser(uid) {
  if (!isValidSafeId(uid)) return false;
  try {
    const cfg = getScriptConfig();
    const token = ScriptApp.getOAuthToken();
    const url = 'https://identitytoolkit.googleapis.com/v1/projects/' + encodeURIComponent(cfg.projectId) + '/accounts:delete';
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ localId: uid }),
      muteHttpExceptions: true
    });
    return res.getResponseCode() === 200;
  } catch (err) {
    return false;
  }
}

function deleteFirestoreDoc(safeRelativePath, idToken) {
  try {
    const cfg = getScriptConfig();
    fetchFirestoreWithFallback(cfg.firestoreBaseUrl + '/' + safeRelativePath, {
      method: 'delete'
    }, idToken);
  } catch (err) {}
}

function getSheetRole(uid) {
  const row = findUserRow(uid);
  return row ? String(row.values[USERS_COLS.role - 1] || 'user').trim().toLowerCase() : null;
}

function isDevwebUid(uid, freshRole) {
  const role = String(freshRole || getSheetRole(uid) || 'user').toLowerCase();
  return role === 'devweb';
}

function isApproverUid(uid, freshRole) {
  const role = String(freshRole || getSheetRole(uid) || 'user').toLowerCase();
  return role === 'admin' || role === 'devweb';
}

function handleRecordBypassAttempt(body, callerUid, idToken) {
  const taskId = String(body.taskId || '').trim();
  const taskName = sanitizeSheetCell(body.taskName || 'Nhiệm vụ', 150);
  const elapsedSeconds = Math.max(0, Math.round(Number(body.elapsedSeconds) || 0));
  const requiredSeconds = Math.max(5, Math.round(Number(body.requiredSeconds) || 75));
  const strikeCount = Math.max(3, Math.round(Number(body.strikeCount) || 3));

  const userProfile = fetchUserFromFirestore(callerUid, idToken, false);
  const cleanUsername = (userProfile && userProfile.username) ? userProfile.username : sanitizeSheetCell(body.username || callerUid, 30);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = ensureCompletionsSheet();
    const now = new Date();
    const alertId = Utilities.getUuid();

    sheet.appendRow([
      alertId,
      sanitizeSheetCell(callerUid, 128),
      sanitizeSheetCell(cleanUsername, 30),
      sanitizeSheetCell(taskId, 128),
      sanitizeSheetCell(taskName, 150),
      0,
      'bypass_alert',
      'bypass_alert',
      now,
      '',
      sanitizeSheetCell('Mở link ' + elapsedSeconds + 's / Yêu cầu ' + requiredSeconds + 's (Vi phạm ' + strikeCount + ' lần)', 200)
    ]);

    writeFirestoreAuditLog(callerUid, 'BYPASS_WARNING_3_STRIKES', taskId || callerUid, {
      username: cleanUsername,
      taskName: taskName,
      elapsedSeconds: elapsedSeconds,
      requiredSeconds: requiredSeconds,
      strikeCount: strikeCount
    }, idToken);
  } finally {
    lock.releaseLock();
  }

  return {
    ok: true,
    message: 'Đã ghi nhận cảnh báo Bypass gửi tới Admin và Devweb.'
  };
}

function handleGetPendingCompletions(callerUid, idToken) {
  const roleCheck = verifyPanelRoleFirebaseAndSheet(callerUid, idToken, ['admin', 'devweb']);
  if (!roleCheck.ok) {
    return { error: roleCheck.error };
  }
  repairTaskAndCompletionLinksFromSessions_(idToken);
  const sheet = ensureCompletionsSheet();
  const data = sheet.getDataRange().getValues();
  const pending = [];
  const bypassAlerts = [];
  for (let i = 1; i < data.length; i++) {
    const st = String(data[i][COMP_COLS.status - 1] || '').toLowerCase();
    if (st === 'pending') {
      const rowUid = String(data[i][COMP_COLS.uid - 1] || '');
      const rowTaskId = String(data[i][COMP_COLS.taskId - 1] || '');
      let rowLink = String(data[i][COMP_COLS.taskLink - 1] || '').trim();
      if (!isRealShortUrl_(rowLink)) {
        const sLookup = findLatestSessionShortUrl_(rowUid, rowTaskId, false);
        if (isRealShortUrl_(sLookup.shortUrl)) {
          rowLink = sLookup.shortUrl;
          sheet.getRange(i + 1, COMP_COLS.taskLink).setValue(rowLink);
        }
      }
      pending.push({
        completionId: String(data[i][COMP_COLS.completionId - 1] || ''),
        uid: rowUid,
        username: String(data[i][COMP_COLS.username - 1] || ''),
        taskName: String(data[i][COMP_COLS.taskName - 1] || ''),
        taskLink: rowLink,
        reward: Number(data[i][COMP_COLS.reward - 1]) || 0,
        createdAt: data[i][COMP_COLS.createdAt - 1]
      });
    } else if (st === 'bypass_alert') {
      bypassAlerts.push({
        completionId: String(data[i][COMP_COLS.completionId - 1] || ''),
        uid: String(data[i][COMP_COLS.uid - 1] || ''),
        username: String(data[i][COMP_COLS.username - 1] || ''),
        taskName: String(data[i][COMP_COLS.taskName - 1] || ''),
        detail: String(data[i][COMP_COLS.taskLink - 1] || 'Dưới thời gian quy định'),
        strikeInfo: 'Vi phạm ≥ 3 lần',
        createdAt: data[i][COMP_COLS.createdAt - 1]
      });
    }
  }
  return { ok: true, pending: pending, bypassAlerts: bypassAlerts };
}

function handleReviewCompletion(body, callerUid, decision, idToken) {
  const roleCheck = verifyPanelRoleFirebaseAndSheet(callerUid, idToken, ['admin', 'devweb']);
  if (!roleCheck.ok) {
    return { error: roleCheck.error };
  }
  const completionId = String(body.completionId || '').trim();
  if (!isValidSafeId(completionId)) {
    return { error: 'Mã duyệt nhiệm vụ không hợp lệ.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = ensureCompletionsSheet();
    const data = sheet.getDataRange().getValues();
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][COMP_COLS.completionId - 1]) === completionId) {
        targetRow = i + 1;
        break;
      }
    }
    if (targetRow === -1) return { error: 'Không tìm thấy nhiệm vụ chờ duyệt.' };

    const rowValues = sheet.getRange(targetRow, 1, 1, Object.keys(COMP_COLS).length).getValues()[0];
    const curStatus = String(rowValues[COMP_COLS.status - 1]).toLowerCase();
    if (curStatus !== 'pending' && curStatus !== 'bypass_alert') {
      return { error: 'Nhiệm vụ này đã được xử lý trước đó.' };
    }

    sheet.getRange(targetRow, COMP_COLS.status).setValue(decision);

    const uid = String(rowValues[COMP_COLS.uid - 1] || '').trim();
    const username = String(rowValues[COMP_COLS.username - 1] || '').trim();
    const taskName = String(rowValues[COMP_COLS.taskName - 1] || 'Nhiệm vụ').trim();
    const reward = Number(rowValues[COMP_COLS.reward - 1]) || 0;

    if (decision === 'approved' && isValidSafeId(uid) && reward > 0 && reward <= 10000000) {
      ensureUserRowInSheet(uid, username, idToken);
      adjustUserBalance(uid, reward, idToken);
      sheet.getRange(targetRow, COMP_COLS.approvedAt).setValue(new Date());
      createFirestoreBalanceHistory(uid, username, 'credit', reward, 'Duyệt nhiệm vụ: ' + taskName, callerUid, idToken);
    }

    writeFirestoreAuditLog(callerUid, decision === 'approved' ? 'APPROVE_COMPLETION' : 'REJECT_COMPLETION', completionId, {
      userId: uid,
      reward: reward
    }, idToken);
  } finally {
    lock.releaseLock();
  }

  return { ok: true, message: decision === 'approved' ? 'Đã duyệt và cộng tiền!' : 'Đã từ chối nhiệm vụ.' };
}

/**
 * Đọc role THẬT từ Firestore (luôn bypass cache khi cần đồng bộ chính xác)
 * và ghi đè ngay vào cột `role` (cột C) trong tab `Users` của Google Sheet nếu khác!
 */
function syncRoleFromFirestore(uid, idToken, forceBypassCache) {
  if (!isValidSafeId(uid) || !idToken) return null;
  try {
    const prof = fetchUserFromFirestore(uid, idToken, forceBypassCache !== false);
    if (!prof || !prof.role) return null;
    const role = String(prof.role).trim().toLowerCase();

    let row = findUserRow(uid);
    if (!row) {
      const uname = prof.username || uid.substring(0, 10);
      row = ensureUserRowInSheet(uid, uname, idToken);
    }
    if (row && String(row.values[USERS_COLS.role - 1] || '').trim().toLowerCase() !== role) {
      const sheet = ensureUsersSheet();
      sheet.getRange(row.rowIndex, USERS_COLS.role).setValue(sanitizeSheetCell(role, 20));
      sheet.getRange(row.rowIndex, USERS_COLS.updatedAt).setValue(new Date());
    }
    return role;
  } catch (err) {
    return null;
  }
}

/**
 * Hàm chạy từ menu "⚡ LinkPro Tools" trên Google Sheet để đồng bộ toàn bộ
 * role của tất cả user từ Firebase Firestore về tab Users của Sheet
 */
function syncAllFirestoreUsersToSheet() {
  try {
    const cfg = getScriptConfig();
    const res = fetchFirestoreWithFallback(cfg.firestoreBaseUrl + '/users?pageSize=200', { method: 'get' }, null);
    if (!res || res.getResponseCode() !== 200) return;
    const json = JSON.parse(res.getContentText());
    const docs = (json && json.documents) ? json.documents : [];
    const sheet = ensureUsersSheet();
    const data = sheet.getDataRange().getValues();
    const mapByUid = {};
    for (let i = 1; i < data.length; i++) {
      const u = String(data[i][USERS_COLS.uid - 1] || '').trim();
      if (u) mapByUid[u] = i + 1;
    }
    const now = new Date();
    docs.forEach(doc => {
      if (!doc || !doc.name || !doc.fields) return;
      const parts = doc.name.split('/');
      const uid = parts[parts.length - 1];
      const f = doc.fields;
      const uname = (f.username && f.username.stringValue) ? f.username.stringValue : uid.substring(0, 10);
      const role = (f.role && f.role.stringValue) ? f.role.stringValue : 'user';
      const bal = Number((f.balance && (f.balance.integerValue || f.balance.doubleValue)) || 0);
      if (mapByUid[uid]) {
        sheet.getRange(mapByUid[uid], USERS_COLS.role).setValue(sanitizeSheetCell(role, 20));
        sheet.getRange(mapByUid[uid], USERS_COLS.updatedAt).setValue(now);
      } else {
        sheet.appendRow([uid, uname, role, bal, now, now]);
      }
    });
  } catch (e) {}
}

function verifyIdToken(idToken) {
  if (!idToken || typeof idToken !== 'string') return { ok: false, error: 'Thiếu token xác thực.' };
  try {
    const cfg = getScriptConfig();
    const url = 'https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + encodeURIComponent(cfg.apiKey);
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Referer: 'https://' + cfg.projectId + '.firebaseapp.com/' },
      payload: JSON.stringify({ idToken: idToken }),
      muteHttpExceptions: true
    });
    const json = JSON.parse(res.getContentText());
    if (res.getResponseCode() !== 200 || !json.users || !json.users[0]) {
      return { ok: false, error: 'Phiên đăng nhập đã hết hạn hoặc không hợp lệ. Vui lòng tải lại trang (F5)!' };
    }
    return {
      ok: true,
      uid: String(json.users[0].localId || ''),
      emailVerified: !!json.users[0].emailVerified
    };
  } catch (err) {
    return { ok: false, error: 'Không thể kết nối dịch vụ xác thực Firebase.' };
  }
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}


// ==========================================================================
// BỘ ĐIỀU PHỐI NHIỆM VỤ THEO TỪNG FILE <TÊN NHIỆM VỤ>.GS (link4m.gs, gtraffic.gs, ...)
// ==========================================================================

const GK_ALLOWED_SITE_ORIGINS = [
  'https://linkprovietnam.web.app',
  'https://linkprovietnam.firebaseapp.com',
  'https://linkpro-d0fcb.web.app',
  'https://linkpro-d0fcb.firebaseapp.com',
  'https://linkproengxnnt.web.app',
  'https://linkproengxnnt.firebaseapp.com'
];

function isAllowedCallbackSiteUrl_(rawUrl) {
  if (!rawUrl || typeof rawUrl !== 'string') return false;
  const trimmed = rawUrl.trim().split('?')[0].replace(/\/+$/, '');
  return GK_ALLOWED_SITE_ORIGINS.indexOf(trimmed) !== -1;
}

function normalizeScriptFileName_(raw) {
  let s = String(raw || '').trim();
  if (!s) return '';
  if (s.indexOf('task_') === 0) {
    s = s.substring(5);
  }
  if (!/\.gs$/i.test(s)) {
    s = s + '.gs';
  }
  return s.toLowerCase();
}

function resolveTaskScriptModule_(scriptFileOrTaskId) {
  if (!scriptFileOrTaskId || typeof TASK_SCRIPT_MODULES === 'undefined') return null;
  const raw = String(scriptFileOrTaskId).trim();
  if (TASK_SCRIPT_MODULES[raw]) return TASK_SCRIPT_MODULES[raw];

  const normalized = normalizeScriptFileName_(raw);
  const keys = Object.keys(TASK_SCRIPT_MODULES);
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    const mod = TASK_SCRIPT_MODULES[k];
    if (!mod || !mod.config) continue;
    if (k.toLowerCase() === normalized) return mod;
    if (String(mod.config.SCRIPT_FILE || '').toLowerCase() === normalized) return mod;
    if (String(mod.config.TASK_ID || '').toLowerCase() === raw.toLowerCase()) return mod;
  }
  return null;
}

/**
 * Trả về danh sách các file (tên nhiệm vụ).gs hiện có trong Apps Script
 * để Admin bấm chọn khi "Thêm Nhiệm Vụ" trên Web
 */
function handleGetAvailableTaskScripts(callerUid, idToken) {
  const roleCheck = verifyPanelRoleFirebaseAndSheet(callerUid, idToken, ['admin', 'devweb']);
  if (!roleCheck.ok) {
    return { error: roleCheck.error };
  }

  const scripts = [];
  const keys = typeof TASK_SCRIPT_MODULES !== 'undefined' ? Object.keys(TASK_SCRIPT_MODULES) : [];
  for (let i = 0; i < keys.length; i++) {
    const mod = TASK_SCRIPT_MODULES[keys[i]];
    if (!mod || !mod.config) continue;
    const c = mod.config;
    scripts.push({
      scriptFile: String(c.SCRIPT_FILE || keys[i]),
      taskId: String(c.TASK_ID || ('task_' + String(c.SCRIPT_FILE || keys[i]).replace(/\.gs$/i, ''))),
      title: String(c.TITLE || ('Nhiệm vụ ' + keys[i])),
      rewardAmount: Math.max(50, Number(c.REWARD) || 250),
      maxCompletionsPerUser: Math.max(1, Number(c.MAX_PER_USER) || 2),
      maxCompletionsPerIp: Math.max(1, Number(c.MAX_PER_IP) || 1),
      approvalType: c.APPROVAL_TYPE === 'manual' ? 'manual' : 'auto',
      sessionExpiryMinutes: Math.max(1, Number(c.SESSION_TTL_MIN) || 30)
    });
  }
  return { ok: true, scripts: scripts };
}

/**
 * Bật / Tắt hiển thị thẻ nhiệm vụ trên Sheet & Firestore
 */
function handleToggleTaskInSheet(body, callerUid, idToken) {
  const roleCheck = verifyPanelRoleFirebaseAndSheet(callerUid, idToken, ['devweb']);
  if (!roleCheck.ok) {
    return { error: roleCheck.error };
  }
  const taskId = String(body.taskId || '').trim();
  if (!isValidSafeId(taskId)) {
    return { error: 'Mã nhiệm vụ không hợp lệ.' };
  }
  const isActive = body.isActive !== false;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const existing = findTaskRowInSheet(taskId);
    if (existing) {
      const sh = ensureTasksSheet();
      sh.getRange(existing.rowIndex, TASKS_COLS.isActive).setValue(isActive);
      sh.getRange(existing.rowIndex, TASKS_COLS.updatedAt).setValue(new Date());
    }
  } finally {
    lock.releaseLock();
  }
  return { ok: true, taskId: taskId, isActive: isActive };
}

/**
 * Điều phối các yêu cầu gk_start và gk_claim tới đúng file (tên nhiệm vụ).gs
 */
function handleGetKey_(req, verifiedUid, idToken) {
  const action = String(req.action || '').trim();
  const uid = String(verifiedUid || '').trim();
  if (!/^[A-Za-z0-9_-]{4,128}$/.test(uid)) {
    return { ok: false, error: 'bad-device', message: 'Không xác định được định danh tài khoản.' };
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    return { ok: false, error: 'busy', message: 'Hệ thống đang xử lý, vui lòng thử lại sau ít giây.' };
  }
  try {
    const sh = gkSheet_();
    if (action === 'gk_start') {
      const targetKey = String(req.scriptFile || req.taskId || 'link4m.gs').trim();
      const mod = resolveTaskScriptModule_(targetKey);
      if (!mod || typeof mod.start !== 'function') {
        return { ok: false, error: 'script-not-found', message: 'Không tìm thấy file nhiệm vụ (' + targetKey + ') trong Apps Script!' };
      }
      return mod.start(sh, uid, req, idToken);
    }

    if (action === 'gk_claim') {
      // Tìm file .gs đã tạo token phiên này trong bảng Sessions (hoặc từ tham số scriptFile/taskId)
      const token = String(req.token || '').trim();
      let mod = resolveTaskScriptModule_(req.scriptFile || req.taskId || '');
      if (!mod && /^S[0-9a-f]{32}$/.test(token)) {
        const rows = sh.getDataRange().getValues();
        for (let i = rows.length - 1; i >= 1; i--) {
          if (String(rows[i][0]) === token) {
            const rowScript = String(rows[i][11] || rows[i][2] || '').trim();
            mod = resolveTaskScriptModule_(rowScript);
            break;
          }
        }
      }
      if (!mod) {
        const keys = Object.keys(TASK_SCRIPT_MODULES);
        if (keys.length > 0) mod = TASK_SCRIPT_MODULES[keys[0]];
      }
      if (!mod || typeof mod.claim !== 'function') {
        return { ok: false, error: 'script-not-found', message: 'Không tìm thấy module xử lý nhiệm vụ trong Apps Script!' };
      }
      return mod.claim(sh, uid, req, idToken);
    }

    return { ok: false, error: 'missing-params' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Hàm dùng chung để thực thi tạo link nhiệm vụ từ cấu hình của từng file <tên nhiệm vụ>.gs
 */
function runTaskScriptStart_(modConfig, sh, uid, req, idToken) {
  if (!modConfig) {
    return { ok: false, error: 'missing-config', message: 'Thiếu cấu hình trong file .gs!' };
  }

  const scriptFile = String(modConfig.SCRIPT_FILE || 'link4m.gs').trim();
  const taskId = String(req.taskId || modConfig.TASK_ID || ('task_' + scriptFile.replace(/\.gs$/i, ''))).trim();
  if (!isValidSafeId(taskId)) {
    return { ok: false, error: 'missing-task', message: 'Mã nhiệm vụ không hợp lệ.' };
  }

  // Kiểm tra trạng thái Bật / Tắt của thẻ nhiệm vụ trên Firestore hoặc Sheet
  const fsTask = fetchTaskFromFirestore(taskId, idToken);
  if (fsTask && fsTask.isActive === false) {
    return { ok: false, error: 'task-inactive', message: 'Nhiệm vụ này hiện đang tạm tắt!' };
  }
  const sheetTaskRow = findTaskRowInSheet(taskId);
  if (!fsTask && sheetTaskRow && sheetTaskRow.values) {
    const activeVal = sheetTaskRow.values[TASKS_COLS.isActive - 1];
    if (activeVal === false || String(activeVal).toLowerCase() === 'false') {
      return { ok: false, error: 'task-inactive', message: 'Nhiệm vụ này hiện đang tạm tắt!' };
    }
  }

  const maxPerUser = Math.max(1, Number(modConfig.MAX_PER_USER) || 2);
  const maxPerIp = Math.max(1, Number(modConfig.MAX_PER_IP) || 1);
  const verifySecs = Math.max(5, Number(modConfig.MIN_SECONDS) || 75);
  const sessionTtlMin = Math.max(1, Number(modConfig.SESSION_TTL_MIN) || 30);
  const rewardAmount = Math.max(50, Number(modConfig.REWARD) || 250);
  const approvalType = modConfig.APPROVAL_TYPE === 'manual' ? 'manual' : 'auto';
  const clientIp = normalizeClientIp_(req.clientIp || '');

  if (!clientIp) {
    return {
      ok: false,
      error: 'missing-ip',
      message: 'Không thể xác định địa chỉ IP của bạn để kiểm tra luật 1 IP / 1 lần vượt. Vui lòng tắt trình chặn kết nối và tải lại trang!'
    };
  }

  // KIỂM TRA LUẬT 1 IP / 1 LẦN VƯỢT TRÊN TOÀN HỆ THỐNG (kể cả khi đổi sang tài khoản khác)
  const ipDoneCount = countCompletedByIpForTask_(clientIp, taskId);
  if (ipDoneCount >= maxPerIp) {
    return {
      ok: false,
      error: 'ip-limit-reached',
      message: 'Địa chỉ IP (' + clientIp + ') của bạn đã vượt đủ ' + maxPerIp + '/' + maxPerIp + ' lần cho nhiệm vụ này (Luật: 1 IP / 1 lần vượt)!'
    };
  }

  // Kiểm tra giới hạn số lần vượt trên mỗi User
  const compSheet = ensureCompletionsSheet();
  const compData = compSheet.getDataRange().getValues();
  let doneCount = 0;
  for (let c = compData.length - 1; c >= 1; c--) {
    if (String(compData[c][COMP_COLS.uid - 1]) === uid && String(compData[c][COMP_COLS.taskId - 1]) === taskId) {
      const st = String(compData[c][COMP_COLS.status - 1] || '').toLowerCase();
      if (st !== 'rejected' && st !== 'bypass_alert') {
        doneCount += 1;
      }
    }
  }
  if (doneCount >= maxPerUser) {
    return {
      ok: false,
      error: 'max-reached',
      message: 'Bạn đã vượt đủ tối đa ' + maxPerUser + '/' + maxPerUser + ' lần cho nhiệm vụ này!'
    };
  }

  const now = Date.now();
  const ttlMs = sessionTtlMin * 60000;
  const rows = sh.getDataRange().getValues();

  // Tái sử dụng phiên pending còn hạn của chính user + taskId đó
  for (let i = rows.length - 1; i >= 1; i--) {
    const r = rows[i];
    if (String(r[1]) !== uid || String(r[2]) !== taskId) continue;
    const createdMs = Number(r[3]) || 0;
    if (r[4] === 'pending' && (now - createdMs) < ttlMs && r[5]) {
      const existingShort = String(r[5] || '').trim();
      if (isRealShortUrl_(existingShort)) {
        updateTaskSheetAndFirestoreShortUrl_(taskId, existingShort, idToken);
      }
      try {
        CacheService.getScriptCache().put('tsess_' + uid + '_' + taskId, JSON.stringify({
          startedAt: createdMs,
          expiresAt: createdMs + ttlMs,
          token: String(r[0]),
          shortUrl: existingShort,
          scriptFile: scriptFile
        }), 21600);
      } catch (e) {}
      return {
        ok: true,
        shortUrl: String(r[5]),
        startedAt: createdMs,
        expiresAt: createdMs + ttlMs,
        sessionTtlMin: sessionTtlMin,
        scriptFile: scriptFile,
        reused: true
      };
    }
  }

  const token = 'S' + (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 32);
  let baseSiteUrl = String(modConfig.SITE_URL || 'https://linkprovietnam.web.app/').trim();
  if (req.siteUrl && isAllowedCallbackSiteUrl_(req.siteUrl)) {
    baseSiteUrl = String(req.siteUrl).trim().split('?')[0];
  }
  if (baseSiteUrl.slice(-1) !== '/') {
    baseSiteUrl += '/';
  }

  const dest = baseSiteUrl + '?t=' + encodeURIComponent(token) + '&taskId=' + encodeURIComponent(taskId) + '&s=' + encodeURIComponent(scriptFile);
  const shortUrl = gkShorten_(dest, modConfig.SHORTENER_TEMPLATE);
  if (!shortUrl) {
    return {
      ok: false,
      error: 'shortener-fail',
      message: 'Không thể tạo link rút gọn từ cấu hình trong file ' + scriptFile + '. Vui lòng kiểm tra lại API token trong ' + scriptFile + '!'
    };
  }

  sh.appendRow([
    token,
    uid,
    taskId,
    now,
    'pending',
    shortUrl,
    rewardAmount,
    '',
    verifySecs,
    approvalType,
    sessionTtlMin,
    scriptFile,
    clientIp
  ]);

  // Cập nhật ngay shortUrl vừa tạo vào cột taskLink (cột C) của tab Tasks và lên Firestore tasks/{taskId}
  updateTaskSheetAndFirestoreShortUrl_(taskId, shortUrl, idToken);

  try {
    CacheService.getScriptCache().put('tsess_' + uid + '_' + taskId, JSON.stringify({
      startedAt: now,
      expiresAt: now + ttlMs,
      token: token,
      shortUrl: shortUrl,
      scriptFile: scriptFile
    }), 21600);
  } catch (e) {}

  return {
    ok: true,
    shortUrl: shortUrl,
    startedAt: now,
    expiresAt: now + ttlMs,
    sessionTtlMin: sessionTtlMin,
    scriptFile: scriptFile
  };
}

/**
 * Hàm dùng chung để xác nhận hoàn thành nhiệm vụ của từng file <tên nhiệm vụ>.gs
 */
function runTaskScriptClaim_(modConfig, sh, uid, req, idToken) {
  const token = String(req.token || '').trim();
  if (!/^S[0-9a-f]{32}$/.test(token)) {
    return { ok: false, error: 'bad-token', message: 'Mã phiên nhiệm vụ không hợp lệ!' };
  }

  const now = Date.now();
  const rows = sh.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[0]) !== token) continue;

    if (String(r[1]) !== uid) {
      return { ok: false, error: 'wrong-device', message: 'Phiên nhiệm vụ này thuộc về một tài khoản khác!' };
    }
    if (r[4] === 'claimed') {
      return { ok: false, error: 'already-claimed', message: 'Phiên nhiệm vụ này đã được nhận thưởng trước đó!' };
    }
    if (r[4] === 'expired') {
      return {
        ok: false,
        expired: true,
        revoked: true,
        error: 'expired-session',
        message: 'Phiên nhiệm vụ này đã hết hạn và bị thu hồi!'
      };
    }

    const taskId = String(r[2] || (modConfig && modConfig.TASK_ID) || '').trim();
    const createdAt = Number(r[3]) || 0;
    const ageMs = Math.max(0, now - createdAt);
    const elapsedSeconds = Math.max(1, Math.floor(ageMs / 1000));
    const reqVerifySeconds = Math.max(5, Number((modConfig && modConfig.MIN_SECONDS) || r[8]) || 75);
    const rowTtlMin = Math.max(1, Number((modConfig && modConfig.SESSION_TTL_MIN) || r[10]) || 30);

    // 1. Quá thời hạn SESSION_TTL_MIN (mặc định 30 phút) -> Thu hồi nhiệm vụ!
    if (ageMs > rowTtlMin * 60000) {
      sh.getRange(i + 1, 5).setValue('expired');
      return {
        ok: false,
        expired: true,
        revoked: true,
        error: 'expired-session',
        message: 'Nhiệm vụ đã quá hạn ' + rowTtlMin + ' phút và đã được thu hồi! Kết quả không được gửi về Admin và không được cộng tiền.'
      };
    }

    // 2. Đã bỏ phần cảnh báo Bypass theo yêu cầu

    const sessionIp = normalizeClientIp_(req.clientIp || String(r[12] || ''));
    const maxPerIp = Math.max(1, Number((modConfig && modConfig.MAX_PER_IP) || 1));
    if (sessionIp) {
      const ipDone = countCompletedByIpForTask_(sessionIp, taskId);
      if (ipDone >= maxPerIp) {
        return {
          ok: false,
          error: 'ip-limit-reached',
          message: 'Địa chỉ IP (' + sessionIp + ') của bạn đã vượt đủ ' + maxPerIp + '/' + maxPerIp + ' lần cho nhiệm vụ này (Luật: 1 IP / 1 lần vượt)!'
        };
      }
    }

    // 3. Hợp lệ -> Đánh dấu claimed và lưu IP vào cột 13 trước khi gọi handleCompleteTask!
    sh.getRange(i + 1, 5).setValue('claimed');
    sh.getRange(i + 1, 8).setValue(now);
    if (sessionIp) {
      sh.getRange(i + 1, 13).setValue(sessionIp);
    }

    return handleCompleteTask({
      taskId: taskId,
      scriptFile: (modConfig && modConfig.SCRIPT_FILE) || String(r[11] || ''),
      sessionShortUrl: String(r[5] || '').trim(),
      clientIp: sessionIp,
      username: req.username || '',
      elapsedSeconds: elapsedSeconds,
      claimElapsedSeconds: elapsedSeconds,
      serverVerifiedSession: true,
      serverStartedAt: createdAt
    }, uid, idToken);
  }

  return { ok: false, error: 'bad-token', message: 'Không tìm thấy phiên làm nhiệm vụ hợp lệ!' };
}

function gkSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Sessions');
  if (!sh) {
    sh = ss.insertSheet('Sessions');
    sh.getRange(1, 1, 1, 13)
      .setValues([[
        'token',
        'uid',
        'taskId',
        'createdAt',
        'status',
        'shortUrl',
        'reward',
        'claimedAt',
        'verifySeconds',
        'approvalType',
        'sessionTtlMin',
        'scriptFile',
        'ipAddress'
      ]])
      .setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, Math.max(sh.getMaxRows() - 1, 1), 3).setNumberFormat('@');
  }
  return sh;
}

function gkShorten_(dest, templateUrl) {
  const tpl = String(templateUrl || '').trim();
  if (!tpl || tpl.indexOf('https://') !== 0 || tpl.indexOf('{url}') === -1) return '';
  try {
    const url = tpl.replace('{url}', encodeURIComponent(dest));
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    const text = String(res.getContentText() || '').trim();
    try {
      const j = JSON.parse(text);
      const u = j.shortenedUrl || j.shortlink || j.short_url || j.shortUrl || j.url ||
                (j.data && (j.data.shortenedUrl || j.data.url));
      if (u && /^https:\/\//i.test(String(u).replace(/\\\//g, '/'))) {
        return String(u).replace(/\\\//g, '/').trim();
      }
    } catch (e) {}
    if (/^https:\/\/\S+$/i.test(text)) return text;
  } catch (e) {}
  return '';
}


function isRealShortUrl_(url) {
  if (!url || typeof url !== 'string') return false;
  const trimmed = url.trim();
  if (trimmed.indexOf('https://') !== 0) return false;
  if (trimmed.indexOf('script.google.com') !== -1 || trimmed.indexOf('script.googleusercontent.com') !== -1) return false;
  if (/\.gs$/i.test(trimmed)) return false;
  if (trimmed === 'https://link4m.co/api-shorten/v2') return false;
  return true;
}

function findLatestSessionShortUrl_(uid, taskId, markClaimedIfPending) {
  const result = { shortUrl: '', startedAt: 0, expired: false, ipAddress: '' };
  try {
    const sh = gkSheet_();
    const sRows = sh.getDataRange().getValues();
    for (let s = sRows.length - 1; s >= 1; s--) {
      const rUid = String(sRows[s][1] || '').trim();
      const rTaskId = String(sRows[s][2] || '').trim();
      const rStatus = String(sRows[s][4] || '').trim().toLowerCase();
      const rShortUrl = String(sRows[s][5] || '').trim();
      if (taskId && rTaskId !== taskId) continue;
      if (uid && rUid !== uid) continue;

      const createdAt = Number(sRows[s][3]) || 0;
      const rowTtlMin = Math.max(1, Number(sRows[s][10]) || 30);
      if (markClaimedIfPending && rStatus === 'pending') {
        if (createdAt > 0 && (Date.now() - createdAt) > rowTtlMin * 60000) {
          sh.getRange(s + 1, 5).setValue('expired');
          result.expired = true;
        } else {
          sh.getRange(s + 1, 5).setValue('claimed');
          sh.getRange(s + 1, 8).setValue(Date.now());
        }
      }
      if (isRealShortUrl_(rShortUrl)) {
        result.shortUrl = rShortUrl;
        result.startedAt = createdAt;
        result.ipAddress = normalizeClientIp_(String(sRows[s][12] || ''));
        return result;
      }
    }
  } catch (e) {}
  return result;
}

function patchFirestoreTaskLink_(taskId, shortUrl, idToken) {
  if (!isValidSafeId(taskId) || !isRealShortUrl_(shortUrl)) return false;
  try {
    const cfg = getScriptConfig();
    const url = cfg.firestoreBaseUrl + '/tasks/' + encodeURIComponent(taskId) + '?updateMask.fieldPaths=appScriptUrl';
    const payload = {
      fields: {
        appScriptUrl: { stringValue: String(shortUrl).trim() }
      }
    };
    const res = fetchFirestoreWithFallback(url, {
      method: 'patch',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    }, idToken, true);
    return res && res.getResponseCode() === 200;
  } catch (e) {
    return false;
  }
}

function updateTaskSheetAndFirestoreShortUrl_(taskId, shortUrl, idToken) {
  if (!taskId || !isRealShortUrl_(shortUrl)) return;
  try {
    const tRow = findTaskRowInSheet(taskId);
    if (tRow) {
      const sh = ensureTasksSheet();
      const curVal = String(tRow.values[TASKS_COLS.taskLink - 1] || '').trim();
      if (curVal !== shortUrl) {
        sh.getRange(tRow.rowIndex, TASKS_COLS.taskLink).setValue(shortUrl);
        sh.getRange(tRow.rowIndex, TASKS_COLS.updatedAt).setValue(new Date());
      }
    }
    patchFirestoreTaskLink_(taskId, shortUrl, idToken);
  } catch (e) {}
}

/**
 * Tự động sửa lại các dòng trong tab Tasks và tab Completions nếu đang lưu
 * 'link4m.gs' hoặc 'https://script.google.com/macros/s/task/...' thành đúng shortUrl trong tab Sessions
 */
function repairTaskAndCompletionLinksFromSessions_(idToken) {
  try {
    const tSheet = ensureTasksSheet();
    const tData = tSheet.getDataRange().getValues();
    for (let i = 1; i < tData.length; i++) {
      const tid = String(tData[i][TASKS_COLS.taskId - 1] || '').trim();
      const curLink = String(tData[i][TASKS_COLS.taskLink - 1] || '').trim();
      if (tid && !isRealShortUrl_(curLink)) {
        const latest = findLatestSessionShortUrl_('', tid, false);
        if (isRealShortUrl_(latest.shortUrl)) {
          tSheet.getRange(i + 1, TASKS_COLS.taskLink).setValue(latest.shortUrl);
          patchFirestoreTaskLink_(tid, latest.shortUrl, idToken);
        }
      }
    }
  } catch (e) {}
}


/**
 * Lấy danh sách Nhiệm vụ đang làm (phiên pending còn hạn trong tab Sessions)
 * và Lịch sử các nhiệm vụ đã làm (trong tab Completions) của chính User đang đăng nhập
 */
function handleGetUserTaskActivity(callerUid, body) {
  if (!isValidSafeId(callerUid)) {
    return { error: 'Mã tài khoản không hợp lệ.' };
  }

  const clientIp = normalizeClientIp_(body && body.clientIp);
  const ipCompletionCounts = {};
  const activeSessions = [];
  const completedTasks = [];
  const now = Date.now();

  // Map tiêu đề nhiệm vụ từ tab Tasks và TASK_SCRIPT_MODULES
  const taskTitleMap = {};
  try {
    const keys = typeof TASK_SCRIPT_MODULES !== 'undefined' ? Object.keys(TASK_SCRIPT_MODULES) : [];
    for (let k = 0; k < keys.length; k++) {
      const m = TASK_SCRIPT_MODULES[keys[k]];
      if (m && m.config && m.config.TASK_ID) {
        taskTitleMap[String(m.config.TASK_ID)] = String(m.config.TITLE || 'Nhiệm vụ Vượt Link');
      }
    }
    const tSheet = ensureTasksSheet();
    const tRows = tSheet.getDataRange().getValues();
    for (let t = 1; t < tRows.length; t++) {
      const tid = String(tRows[t][TASKS_COLS.taskId - 1] || '').trim();
      const tTitle = String(tRows[t][TASKS_COLS.title - 1] || '').trim();
      if (tid && tTitle) taskTitleMap[tid] = tTitle;
    }
  } catch (e) {}

  // 1. Đọc các phiên nhiệm vụ ĐANG LÀM (pending & còn hạn 30 phút) trong tab Sessions
  try {
    const sSheet = gkSheet_();
    const sRows = sSheet.getDataRange().getValues();
    for (let i = sRows.length - 1; i >= 1; i--) {
      const rUid = String(sRows[i][1] || '').trim();
      if (rUid !== callerUid) continue;
      const rStatus = String(sRows[i][4] || '').trim().toLowerCase();
      const createdAt = Number(sRows[i][3]) || 0;
      const ttlMin = Math.max(1, Number(sRows[i][10]) || 30);
      const expiresAt = createdAt + ttlMin * 60000;

      if (rStatus === 'pending') {
        if (now < expiresAt) {
          const tId = String(sRows[i][2] || '').trim();
          activeSessions.push({
            token: String(sRows[i][0] || ''),
            taskId: tId,
            taskTitle: taskTitleMap[tId] || 'Nhiệm vụ Vượt Link',
            shortUrl: String(sRows[i][5] || ''),
            rewardAmount: Number(sRows[i][6]) || 250,
            approvalType: String(sRows[i][9] || 'auto'),
            scriptFile: String(sRows[i][11] || 'link4m.gs'),
            startedAt: createdAt,
            expiresAt: expiresAt,
            sessionTtlMin: ttlMin
          });
        } else {
          sSheet.getRange(i + 1, 5).setValue('expired');
        }
      }
    }
  } catch (e) {}

  // 2. Đọc LỊCH SỬ NHIỆM VỤ ĐÃ LÀM của User trong tab Completions
  try {
    const cSheet = ensureCompletionsSheet();
    const cRows = cSheet.getDataRange().getValues();
    for (let j = cRows.length - 1; j >= 1; j--) {
      const cUid = String(cRows[j][COMP_COLS.uid - 1] || '').trim();
      if (cUid !== callerUid) continue;
      const cType = String(cRows[j][COMP_COLS.type - 1] || '').trim().toLowerCase();
      if (cType === 'bypass_alert' || cType === 'manual_adjust') continue;

      const cTaskId = String(cRows[j][COMP_COLS.taskId - 1] || '').trim();
      let cTaskLink = String(cRows[j][COMP_COLS.taskLink - 1] || '').trim();
      if (!isRealShortUrl_(cTaskLink)) {
        const sLookup = findLatestSessionShortUrl_(callerUid, cTaskId, false);
        if (isRealShortUrl_(sLookup.shortUrl)) {
          cTaskLink = sLookup.shortUrl;
          cSheet.getRange(j + 1, COMP_COLS.taskLink).setValue(cTaskLink);
        }
      }

      completedTasks.push({
        completionId: String(cRows[j][COMP_COLS.completionId - 1] || ''),
        taskId: cTaskId,
        taskName: String(cRows[j][COMP_COLS.taskName - 1] || taskTitleMap[cTaskId] || 'Nhiệm vụ Vượt Link'),
        reward: Number(cRows[j][COMP_COLS.reward - 1]) || 0,
        approvalType: cType || 'auto',
        status: String(cRows[j][COMP_COLS.status - 1] || 'approved').toLowerCase(),
        createdAt: cRows[j][COMP_COLS.createdAt - 1],
        approvedAt: cRows[j][COMP_COLS.approvedAt - 1],
        taskLink: cTaskLink
      });
      if (completedTasks.length >= 50) break;
    }
  } catch (e) {}

  if (clientIp) {
    try {
      const cSheet = ensureCompletionsSheet();
      const cData = cSheet.getDataRange().getValues();
      for (let i = 1; i < cData.length; i++) {
        const rowIp = normalizeClientIp_(cData[i][COMP_COLS.ipAddress - 1]);
        const rowTask = String(cData[i][COMP_COLS.taskId - 1] || '').trim();
        const rowSt = String(cData[i][COMP_COLS.status - 1] || '').toLowerCase();
        const rowType = String(cData[i][COMP_COLS.type - 1] || '').toLowerCase();
        if (rowIp && rowIp === clientIp && rowTask && rowSt !== 'rejected' && rowType !== 'bypass_alert' && rowType !== 'manual_adjust') {
          ipCompletionCounts[rowTask] = (ipCompletionCounts[rowTask] || 0) + 1;
        }
      }
    } catch (e) {}
  }

  return {
    ok: true,
    clientIp: clientIp,
    ipCompletionCounts: ipCompletionCounts,
    activeSessions: activeSessions,
    completedTasks: completedTasks
  };
}


function normalizeClientIp_(rawIp) {
  if (!rawIp || typeof rawIp !== 'string') return '';
  const cleaned = rawIp.trim().toLowerCase();
  if (!/^[0-9a-f.:]{3,45}$/.test(cleaned)) return '';
  return cleaned;
}

/**
 * Hàm kiểm tra Luật 1 IP / 1 lần vượt dùng chung cho các nhiệm vụ
 */
function enforceOneIpOnePassRule_(rawIp, taskId, maxPerIpLimit) {
  const clientIp = normalizeClientIp_(rawIp || '');
  const maxPerIp = Math.max(1, Number(maxPerIpLimit) || 1);
  if (!clientIp) {
    return {
      ok: false,
      error: 'missing-ip',
      message: 'Không thể xác định địa chỉ IP của bạn để kiểm tra luật 1 IP / 1 lần vượt. Vui lòng tắt trình chặn kết nối và tải lại trang!'
    };
  }
  const ipDoneCount = countCompletedByIpForTask_(clientIp, taskId);
  if (ipDoneCount >= maxPerIp) {
    return {
      ok: false,
      error: 'ip-limit-reached',
      clientIp: clientIp,
      ipDoneCount: ipDoneCount,
      maxPerIp: maxPerIp,
      message: 'Luật 1 IP / 1 lần vượt: Địa chỉ IP (' + clientIp + ') đã được sử dụng để vượt nhiệm vụ này (' + ipDoneCount + '/' + maxPerIp + ' lần). Vui lòng đổi địa chỉ IP khác để thực hiện lượt vượt tiếp theo!'
    };
  }
  return {
    ok: true,
    clientIp: clientIp,
    ipDoneCount: ipDoneCount,
    maxPerIp: maxPerIp
  };
}

function countCompletedByIpForTask_(ipAddr, taskId) {
  const cleanIp = normalizeClientIp_(ipAddr);
  if (!cleanIp || !taskId) return 0;
  let count = 0;
  try {
    const cSheet = ensureCompletionsSheet();
    const cData = cSheet.getDataRange().getValues();
    for (let i = 1; i < cData.length; i++) {
      const rowTask = String(cData[i][COMP_COLS.taskId - 1] || '').trim();
      if (rowTask !== taskId) continue;
      const rowStatus = String(cData[i][COMP_COLS.status - 1] || '').toLowerCase();
      const rowType = String(cData[i][COMP_COLS.type - 1] || '').toLowerCase();
      if (rowStatus === 'rejected' || rowType === 'bypass_alert' || rowType === 'manual_adjust') continue;
      const rowIp = normalizeClientIp_(String(cData[i][COMP_COLS.ipAddress - 1] || ''));
      if (rowIp && rowIp === cleanIp) {
        count += 1;
      }
    }
  } catch (e) {}
  return count;
}
