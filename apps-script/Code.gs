/**
 * BIENSODEP88 — máy chủ tài khoản thành viên cho biensodep88.com/admin
 * Chạy trong tài khoản Google: biensodaugia88@gmail.com
 *
 * Cài đặt (làm 1 lần):
 *  1. Tạo Google Sheet mới tên "BIENSODEP88 - Thành viên" trong biensodaugia88@gmail.com
 *  2. Tiện ích mở rộng → Apps Script → dán toàn bộ file này vào Code.gs → Lưu
 *  3. Chọn hàm "setup" → Chạy → cấp quyền (Gmail + Sheet)
 *  4. Triển khai → Triển khai mới → Loại: Ứng dụng web
 *       Thực thi dưới dạng: Tôi (biensodaugia88@gmail.com)
 *       Người có quyền truy cập: Bất kỳ ai
 *     → Triển khai → copy "URL ứng dụng web" (…/exec) gửi cho Claude
 *
 * Duyệt thành viên: bấm link "DUYỆT" trong email báo về biensodaugia88@gmail.com,
 * hoặc sửa cột "Trạng thái" trong tab THANHVIEN thành DA_DUYET (KHOA để khoá).
 */

const OWNER_EMAIL = "biensodaugia88@gmail.com";
const SITE = "https://biensodep88.com";
const SESSION_DAYS_REMEMBER = 30;
const SESSION_HOURS_SHORT = 12;
const MAX_PENDING = 200;               // chặn spam đăng ký

const SH_MEMBER = "THANHVIEN";
const SH_SESSION = "PHIEN";
const H_MEMBER = ["Email", "Trạng thái", "Ngày đăng ký", "Ngày duyệt", "Lần đăng nhập cuối", "Ghi chú", "salt", "hash"];
const H_SESSION = ["token", "email", "hết hạn (ms)"];
const ST = { PENDING: "CHO_DUYET", OK: "DA_DUYET", LOCK: "KHOA" };

/* ========== Cài đặt ========== */
function setup() {
  const ss = SpreadsheetApp.getActive();
  sheet_(SH_MEMBER, H_MEMBER);
  sheet_(SH_SESSION, H_SESSION);
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty("SECRET")) props.setProperty("SECRET", Utilities.getUuid() + Utilities.getUuid());
  const m = ss.getSheetByName(SH_MEMBER);
  m.setFrozenRows(1);
  m.hideColumns(7, 2);                 // ẩn cột salt/hash
  const rule = SpreadsheetApp.newDataValidation().requireValueInList([ST.PENDING, ST.OK, ST.LOCK], true).build();
  m.getRange(2, 2, 1000, 1).setDataValidation(rule);
  const def = ss.getSheetByName("Sheet1") || ss.getSheetByName("Trang tính1");
  if (def && ss.getSheets().length > 2) ss.deleteSheet(def);
  MailApp.getRemainingDailyQuota();    // gọi để xin quyền gửi mail
  return "OK";
}

function sheet_(name, header) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(header);
  return sh;
}

/* ========== HTTP ========== */
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents || "{}"); } catch (err) { return out_({ ok: false, error: "Dữ liệu không hợp lệ." }); }
  try {
    switch (body.action) {
      case "member_register": return out_(register_(body));
      case "member_login": return out_(login_(body));
      case "member_logout": return out_(logout_(body));
      case "member_data": return out_(data_(body));
      default: return out_({ ok: false, error: "Thao tác không hợp lệ." });
    }
  } catch (err) {
    return out_({ ok: false, error: "Lỗi máy chủ: " + err.message });
  }
}

function doGet(e) {
  const p = e.parameter || {};
  if (p.action === "approve" || p.action === "lock") {
    const email = String(p.email || "").toLowerCase();
    if (!email || p.k !== sign_(p.action + ":" + email)) return page_("Link không hợp lệ hoặc đã bị sửa.");
    const r = findMember_(email);
    if (!r) return page_("Không tìm thấy tài khoản " + email + ".");
    const sh = sheet_(SH_MEMBER, H_MEMBER);
    if (p.action === "approve") {
      sh.getRange(r.row, 2).setValue(ST.OK);
      sh.getRange(r.row, 4).setValue(new Date());
      try {
        MailApp.sendEmail(email, "Tài khoản biensodep88.com đã được duyệt",
          "Tài khoản " + email + " đã được duyệt.\nĐăng nhập tại: " + SITE + "/admin/\n\nHotline: 093.8888.226 - 091.8888.299");
      } catch (err) {}
      return page_("Đã DUYỆT tài khoản " + email + ".");
    }
    sh.getRange(r.row, 2).setValue(ST.LOCK);
    killSessions_(email);
    return page_("Đã KHOÁ tài khoản " + email + ".");
  }
  return out_({ ok: true, service: "biensodep88-member" });
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function page_(t) { return HtmlService.createHtmlOutput('<meta name="viewport" content="width=device-width"><div style="font:18px Arial;padding:40px;text-align:center">' + t + '<br><br><a href="' + SITE + '/admin/">Mở trang quản trị</a></div>'); }

/* ========== Nghiệp vụ ========== */
function register_(b) {
  const email = cleanEmail_(b.email), pass = String(b.pass || "");
  if (!email) return { ok: false, error: "Email không hợp lệ." };
  if (pass.length < 6) return { ok: false, error: "Mật khẩu tối thiểu 6 ký tự." };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    if (findMember_(email)) return { ok: false, error: "Email này đã đăng ký. Vui lòng đăng nhập hoặc chờ duyệt." };
    const sh = sheet_(SH_MEMBER, H_MEMBER);
    const pending = sh.getLastRow() < 2 ? 0 : sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().filter(r => r[0] === ST.PENDING).length;
    if (pending >= MAX_PENDING) return { ok: false, error: "Hệ thống đang quá tải đăng ký, vui lòng gọi hotline." };
    const salt = Utilities.getUuid();
    sh.appendRow([email, ST.PENDING, new Date(), "", "", "", salt, hash_(pass, salt)]);
  } finally { lock.releaseLock(); }
  try {
    const base = ScriptApp.getService().getUrl();
    const ok = base + "?action=approve&email=" + encodeURIComponent(email) + "&k=" + sign_("approve:" + email);
    const no = base + "?action=lock&email=" + encodeURIComponent(email) + "&k=" + sign_("lock:" + email);
    MailApp.sendEmail({
      to: OWNER_EMAIL,
      subject: "[biensodep88] Đăng ký mới: " + email,
      htmlBody: "<p>Có tài khoản mới đăng ký trang quản trị <b>biensodep88.com</b>:</p><p style='font-size:18px'><b>" + email + "</b></p>"
        + "<p><a href='" + ok + "' style='background:#1a7f37;color:#fff;padding:10px 18px;border-radius:6px;text-decoration:none'>DUYỆT</a>"
        + " &nbsp; <a href='" + no + "' style='background:#b42318;color:#fff;padding:10px 18px;border-radius:6px;text-decoration:none'>KHOÁ</a></p>"
        + "<p>Hoặc sửa cột Trạng thái trong tab THANHVIEN của Google Sheet.</p>"
    });
  } catch (err) {}
  return { ok: true, message: "Đăng ký thành công! Tài khoản đang chờ chủ kho duyệt." };
}

function login_(b) {
  const email = cleanEmail_(b.email), pass = String(b.pass || "");
  const r = email && findMember_(email);
  if (!r || r.hash !== hash_(pass, r.salt)) return { ok: false, error: "Sai email hoặc mật khẩu." };
  if (r.status === ST.PENDING) return { ok: false, error: "Tài khoản đang chờ duyệt. Vui lòng liên hệ 093.8888.226." };
  if (r.status !== ST.OK) return { ok: false, error: "Tài khoản đã bị khoá." };
  const exp = Date.now() + (b.remember ? SESSION_DAYS_REMEMBER * 864e5 : SESSION_HOURS_SHORT * 36e5);
  const token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
  const ss = sheet_(SH_SESSION, H_SESSION);
  cleanSessions_(ss);
  ss.appendRow([token, email, exp]);
  sheet_(SH_MEMBER, H_MEMBER).getRange(r.row, 5).setValue(new Date());
  return { ok: true, token: token, email: email, exp: exp };
}

function logout_(b) {
  const ss = sheet_(SH_SESSION, H_SESSION), t = String(b.token || "");
  if (!t || ss.getLastRow() < 2) return { ok: true };
  const v = ss.getRange(2, 1, ss.getLastRow() - 1, 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (v[i][0] === t) ss.deleteRow(i + 2);
  return { ok: true };
}

// Chỉ kiểm tra phiên; dữ liệu biển trang /admin tự tải từ biensodep88.com/vpa-data/live.json
function data_(b) {
  const s = session_(String(b.token || ""));
  if (!s) return { ok: false, relogin: true, error: "Phiên đăng nhập đã hết hạn, vui lòng đăng nhập lại." };
  const r = findMember_(s.email);
  if (!r || r.status !== ST.OK) { killSessions_(s.email); return { ok: false, relogin: true, error: "Tài khoản không còn được phép truy cập." }; }
  return { ok: true, email: s.email, live: SITE + "/vpa-data/live.json" };
}

/* ========== Tiện ích ========== */
function cleanEmail_(e) { e = String(e || "").trim().toLowerCase(); return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(e) ? e : ""; }
function hash_(pass, salt) {
  let d = pass + "|" + salt;
  for (let i = 0; i < 500; i++) d = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, d + salt, Utilities.Charset.UTF_8));
  return d;
}
function sign_(s) {
  const key = PropertiesService.getScriptProperties().getProperty("SECRET") || "";
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, key)).replace(/=+$/, "");
}
function findMember_(email) {
  const sh = sheet_(SH_MEMBER, H_MEMBER);
  if (sh.getLastRow() < 2) return null;
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]).toLowerCase() === email)
    return { row: i + 2, status: String(v[i][1]).trim(), salt: v[i][6], hash: v[i][7] };
  return null;
}
function session_(token) {
  if (!token) return null;
  const ss = sheet_(SH_SESSION, H_SESSION);
  if (ss.getLastRow() < 2) return null;
  const v = ss.getRange(2, 1, ss.getLastRow() - 1, 3).getValues();
  for (let i = 0; i < v.length; i++) if (v[i][0] === token) return Number(v[i][2]) > Date.now() ? { email: v[i][1] } : null;
  return null;
}
function cleanSessions_(ss) {
  if (ss.getLastRow() < 2) return;
  const v = ss.getRange(2, 1, ss.getLastRow() - 1, 3).getValues(), now = Date.now();
  for (let i = v.length - 1; i >= 0; i--) if (Number(v[i][2]) < now) ss.deleteRow(i + 2);
}
function killSessions_(email) {
  const ss = sheet_(SH_SESSION, H_SESSION);
  if (ss.getLastRow() < 2) return;
  const v = ss.getRange(2, 2, ss.getLastRow() - 1, 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (v[i][0] === email) ss.deleteRow(i + 2);
}
