/**
 * BIENSODEP88 — máy chủ thành viên cho biensodep88.com/admin (DUYỆT QUA GMAIL, KHÔNG CẦN GOOGLE SHEET)
 * Chạy trong tài khoản Google: biensodaugia88@gmail.com
 *
 * Cài đặt (làm 1 lần):
 *  1. Vào https://script.google.com → Dự án mới → đặt tên "BIENSODEP88 - Thành viên"
 *  2. Xoá hết code mẫu, dán toàn bộ file này vào → Lưu
 *  3. Chọn hàm "setup" → Chạy → cấp quyền (gửi email)
 *  4. Triển khai → Triển khai mới → Loại: Ứng dụng web
 *       Thực thi dưới dạng: Tôi (biensodaugia88@gmail.com)
 *       Người có quyền truy cập: Bất kỳ ai
 *     → Triển khai → copy "URL ứng dụng web" (…/exec) gửi cho Claude
 *
 * Duyệt thành viên: khi có người đăng ký, Gmail biensodaugia88@gmail.com nhận thư có nút DUYỆT / KHOÁ.
 * Mỗi thư đều có link "Xem danh sách thành viên" để duyệt / khoá / mở lại bất kỳ ai.
 * Dữ liệu thành viên nằm trong chính dự án Apps Script này (không có file Sheet để lỡ xoá).
 */

const OWNER_EMAIL = "biensodaugia88@gmail.com";
const SITE = "https://biensodep88.com";
const HOTLINE = "093.8888.226 - 091.8888.299";
const SESSION_DAYS_REMEMBER = 30;
const SESSION_HOURS_SHORT = 12;
const MAX_PENDING = 100;               // chặn spam đăng ký
const ST = { PENDING: "CHO_DUYET", OK: "DA_DUYET", LOCK: "KHOA" };
const ST_TXT = { CHO_DUYET: "Chờ duyệt", DA_DUYET: "Đã duyệt", KHOA: "Đã khoá" };

/* ========== Cài đặt ========== */
function setup() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty("SECRET")) p.setProperty("SECRET", Utilities.getUuid() + Utilities.getUuid());
  MailApp.getRemainingDailyQuota();    // xin quyền gửi email
  return "OK";
}

/* ========== HTTP ========== */
function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents || "{}"); } catch (err) { return json_({ ok: false, error: "Dữ liệu không hợp lệ." }); }
  try {
    switch (b.action) {
      case "member_register": return json_(register_(b));
      case "member_login": return json_(login_(b));
      case "member_logout": return json_({ ok: true });
      case "member_data": return json_(data_(b));
      default: return json_({ ok: false, error: "Thao tác không hợp lệ." });
    }
  } catch (err) { return json_({ ok: false, error: "Lỗi máy chủ: " + err.message }); }
}

function doGet(e) {
  const p = e.parameter || {};
  const email = String(p.email || "").toLowerCase();
  if (p.action === "list") {
    if (p.k !== sign_("list")) return page_("Link không hợp lệ.");
    return listPage_();
  }
  if (["approve", "lock", "delete"].indexOf(p.action) >= 0) {
    if (!email || p.k !== sign_(p.action + ":" + email)) return page_("Link không hợp lệ hoặc đã bị sửa.");
    const m = getM_(email);
    if (!m) return page_("Không tìm thấy tài khoản " + esc_(email) + ".");
    if (p.action === "delete") { delM_(email); return page_("Đã XOÁ tài khoản " + esc_(email) + ".", true); }
    if (p.action === "approve") {
      const first = m.s !== ST.OK;
      m.s = ST.OK; m.ad = Date.now(); m.v = (m.v || 0) + 1; putM_(email, m);
      if (first) try {
        MailApp.sendEmail(email, "Tài khoản biensodep88.com đã được duyệt",
          "Tài khoản " + email + " đã được duyệt.\nĐăng nhập tại: " + SITE + "/admin/\n\nHotline: " + HOTLINE);
      } catch (err) {}
      return page_("Đã DUYỆT tài khoản " + esc_(email) + ".", true);
    }
    m.s = ST.LOCK; m.v = (m.v || 0) + 1; putM_(email, m);   // đổi phiên bản → đăng xuất mọi phiên đang mở
    return page_("Đã KHOÁ tài khoản " + esc_(email) + ".", true);
  }
  return json_({ ok: true, service: "biensodep88-member" });
}

/* ========== Nghiệp vụ ========== */
function register_(b) {
  const email = cleanEmail_(b.email), pass = String(b.pass || "");
  if (!email) return { ok: false, error: "Email không hợp lệ." };
  if (pass.length < 6) return { ok: false, error: "Mật khẩu tối thiểu 6 ký tự." };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    if (getM_(email)) return { ok: false, error: "Email này đã đăng ký. Vui lòng đăng nhập hoặc chờ duyệt." };
    if (allM_().filter(x => x.m.s === ST.PENDING).length >= MAX_PENDING) return { ok: false, error: "Đang quá nhiều tài khoản chờ duyệt, vui lòng gọi " + HOTLINE + "." };
    const salt = Utilities.getUuid();
    putM_(email, { s: ST.PENDING, t: Date.now(), salt: salt, h: hash_(pass, salt), v: 1 });
  } finally { lock.releaseLock(); }
  try { notifyOwner_(email); } catch (err) {}
  return { ok: true, message: "Đăng ký thành công! Tài khoản đang chờ chủ kho duyệt." };
}

function login_(b) {
  const email = cleanEmail_(b.email), pass = String(b.pass || "");
  const m = email && getM_(email);
  if (!m || m.h !== hash_(pass, m.salt)) return { ok: false, error: "Sai email hoặc mật khẩu." };
  if (m.s === ST.PENDING) return { ok: false, error: "Tài khoản đang chờ duyệt. Vui lòng liên hệ " + HOTLINE + "." };
  if (m.s !== ST.OK) return { ok: false, error: "Tài khoản đã bị khoá." };
  const exp = Date.now() + (b.remember ? SESSION_DAYS_REMEMBER * 864e5 : SESSION_HOURS_SHORT * 36e5);
  const body = Utilities.base64EncodeWebSafe(JSON.stringify({ e: email, x: exp, v: m.v || 1 }));
  m.ll = Date.now(); putM_(email, m);
  return { ok: true, token: body + "." + sign_("tok:" + body), email: email, exp: exp };
}

// Chỉ kiểm tra phiên; biển VPA trang /admin tự tải từ biensodep88.com/vpa-data/live.json
function data_(b) {
  const t = String(b.token || "").split(".");
  const bad = { ok: false, relogin: true, error: "Phiên đăng nhập đã hết hạn, vui lòng đăng nhập lại." };
  if (t.length !== 2 || t[1] !== sign_("tok:" + t[0])) return bad;
  let s; try { s = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(t[0])).getDataAsString()); } catch (err) { return bad; }
  if (!s || s.x < Date.now()) return bad;
  const m = getM_(s.e);
  if (!m || m.s !== ST.OK || (m.v || 1) !== s.v) return { ok: false, relogin: true, error: "Tài khoản không còn được phép truy cập." };
  return { ok: true, email: s.e, live: SITE + "/vpa-data/live.json" };
}

/* ========== Gmail ========== */
function notifyOwner_(email) {
  const ok = link_("approve", email), no = link_("lock", email);
  MailApp.sendEmail({
    to: OWNER_EMAIL,
    subject: "[biensodep88] Đăng ký mới cần duyệt: " + email,
    htmlBody: "<div style='font:15px Arial'>"
      + "<p>Có tài khoản mới đăng ký trang quản trị <b>biensodep88.com</b>:</p>"
      + "<p style='font-size:20px'><b>" + esc_(email) + "</b></p>"
      + "<p><a href='" + ok + "' style='background:#1a7f37;color:#fff;padding:12px 22px;border-radius:6px;text-decoration:none;font-weight:bold'>DUYỆT</a>"
      + " &nbsp; <a href='" + no + "' style='background:#b42318;color:#fff;padding:12px 22px;border-radius:6px;text-decoration:none;font-weight:bold'>KHOÁ</a></p>"
      + "<p><a href='" + listUrl_() + "'>Xem danh sách tất cả thành viên</a></p></div>"
  });
}

function listPage_() {
  const rows = allM_().sort((a, b) => (b.m.t || 0) - (a.m.t || 0));
  const fmt = t => t ? Utilities.formatDate(new Date(t), "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm") : "";
  const btn = (a, e, txt, c) => "<a href='" + link_(a, e) + "' target='_top' style='background:" + c + ";color:#fff;padding:6px 10px;border-radius:5px;text-decoration:none;font-size:13px;margin-right:4px'>" + txt + "</a>";
  const tr = rows.map(r => "<tr><td>" + esc_(r.e) + "</td><td>" + (ST_TXT[r.m.s] || r.m.s) + "</td><td>" + fmt(r.m.t) + "</td><td>" + fmt(r.m.ll) + "</td><td>"
    + (r.m.s !== ST.OK ? btn("approve", r.e, "Duyệt", "#1a7f37") : "")
    + (r.m.s !== ST.LOCK ? btn("lock", r.e, "Khoá", "#b42318") : "")
    + btn("delete", r.e, "Xoá", "#555") + "</td></tr>").join("");
  return HtmlService.createHtmlOutput("<meta name='viewport' content='width=device-width'><div style='font:14px Arial;padding:16px'>"
    + "<h2>Thành viên biensodep88.com (" + rows.length + ")</h2>"
    + "<table cellpadding='8' style='border-collapse:collapse' border='1'><tr style='background:#eee'><th>Email</th><th>Trạng thái</th><th>Ngày đăng ký</th><th>Đăng nhập cuối</th><th></th></tr>"
    + (tr || "<tr><td colspan='5'>Chưa có thành viên</td></tr>") + "</table></div>").setTitle("Thành viên biensodep88");
}

/* ========== Lưu trữ (Script Properties) ========== */
const KEY = e => "m:" + e;
function getM_(e) { const v = PropertiesService.getScriptProperties().getProperty(KEY(e)); return v ? JSON.parse(v) : null; }
function putM_(e, m) { PropertiesService.getScriptProperties().setProperty(KEY(e), JSON.stringify(m)); }
function delM_(e) { PropertiesService.getScriptProperties().deleteProperty(KEY(e)); }
function allM_() {
  const all = PropertiesService.getScriptProperties().getProperties();
  return Object.keys(all).filter(k => k.indexOf("m:") === 0).map(k => ({ e: k.slice(2), m: JSON.parse(all[k]) }));
}

/* ========== Tiện ích ========== */
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function page_(t, back) {
  return HtmlService.createHtmlOutput("<meta name='viewport' content='width=device-width'><div style='font:18px Arial;padding:40px;text-align:center'>" + t
    + "<br><br>" + (back ? "<a href='" + listUrl_() + "' target='_top'>Xem danh sách thành viên</a> · " : "")
    + "<a href='" + SITE + "/admin/' target='_top'>Mở trang quản trị</a></div>");
}
function esc_(s) { return String(s).replace(/[&<>'"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[c])); }
function base_() { return ScriptApp.getService().getUrl(); }
function link_(a, e) { return base_() + "?action=" + a + "&email=" + encodeURIComponent(e) + "&k=" + sign_(a + ":" + e); }
function listUrl_() { return base_() + "?action=list&k=" + sign_("list"); }
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
