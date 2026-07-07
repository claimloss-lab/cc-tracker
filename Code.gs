// ============================================================
//  CC Tracker — Google Apps Script (Code.gs)  v3.5
//  LINE Messaging API version
//  Deploy as: Web App (Execute as Me, Anyone can access)
//
//  New in v3.5:
//  - checkAndNotify() ส่งเป็น Flex Message (การ์ด) แทน text ธรรมดา
//    เพิ่มฟังก์ชัน sendLineFlex() + buildNotifyFlex()
//  - (ถ้าอยากมีปุ่ม "เปิดแอป" ในการ์ด ให้เพิ่ม key "appUrl" ใน Config sheet
//    เป็น URL ของ dashboard — ถ้าไม่ตั้งไว้ การ์ดจะไม่มีปุ่มนี้)
//
//  Fixed in v3.3:
//  - confirmed flag ("ไม่มียอด") ตอนนี้ persist ลง Sheet จริง
//    (เพิ่มคอลัมน์ confirmed ใน SHEET_LOG, getLog/saveLog อ่าน-เขียนครบ)
//  - เพิ่ม action testLineMessage — ยิงข้อความทดสอบแบบ force
//    (ปุ่ม "ทดสอบส่ง LINE" เดิมเรียก checkAndNotify ซึ่งอาจไม่ส่งอะไรเลย
//     ถ้าไม่มีบัตรเข้าเงื่อนไข แต่ frontend ก็ยังโชว์ "สำเร็จ" อยู่ดี)
//
//  Fixed in v3.2:
//  - Quick Pay multi-match เลือกบัตรผิดใบ (เก็บ cardIds ใน state)
//  - Timezone ใช้ Asia/Bangkok สม่ำเสมอทุกจุด
//  - State มี TTL 10 นาที (หมดอายุอัตโนมัติ)
//  - LockService ป้องกัน race condition ใน saveLog/saveConfig
//  - getConfig cache ต่อ request
//  - THAI_MONTHS เป็น constant เดียว
//
//  ⚠️ สำคัญหลัง deploy เวอร์ชันนี้ครั้งแรก:
//  เปิด Google Sheet → แท็บ MonthlyLog → เช็คว่ามีคอลัมน์ G ชื่อ "confirmed"
//  ถ้า Sheet มีอยู่แล้วก่อนหน้านี้ (ไม่ใช่สร้างใหม่) ให้เพิ่มหัวคอลัมน์เอง
//  เพราะ getOrCreate จะสร้าง header ให้เฉพาะตอนที่ยังไม่มีชีตนั้นอยู่เท่านั้น
// ============================================================

const SHEET_CARDS  = 'Cards';
const SHEET_LOG    = 'MonthlyLog';
const SHEET_CONFIG = 'Config';
const TZ           = 'Asia/Bangkok';
const STATE_TTL_MS = 10 * 60 * 1000;   // state หมดอายุใน 10 นาที
const THAI_MONTHS  = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];

// ── Bangkok Time Helpers ──────────────────────────────────────
// ใช้แทน new Date().getDate()/getMonth() เพื่อให้ตรง Bangkok เสมอ
// ไม่ว่า script timezone จะตั้งเป็นอะไร

function bkkNow() {
  const now = new Date();
  return {
    day:      parseInt(Utilities.formatDate(now, TZ, 'd')),
    month0:   parseInt(Utilities.formatDate(now, TZ, 'M')) - 1,  // 0-based
    year:     parseInt(Utilities.formatDate(now, TZ, 'yyyy')),
    monthKey: Utilities.formatDate(now, TZ, 'yyyy-MM'),
    dateStr:  Utilities.formatDate(now, TZ, 'd MMM yyyy')
  };
}

function thaiMonthLabel(month0, year) {
  return `${THAI_MONTHS[month0]} ${year + 543}`;
}

// ── Utility ───────────────────────────────────────────────────

function getOrCreate(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); if (headers) sh.appendRow(headers); }
  return sh;
}

// ── Routers ───────────────────────────────────────────────────

function doGet(e) {
  try {
    if (!e.parameter || Object.keys(e.parameter).length === 0) {
      return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
    }
    const cb      = e.parameter.callback || '';
    const payload = e.parameter.payload  || '';
    const action  = e.parameter.action   || '';
    let result;
    if (payload) {
      result = handlePost(JSON.parse(decodeURIComponent(payload)));
    } else {
      result = handleGet(action, e.parameter);
    }
    const json = JSON.stringify(result);
    const out  = cb ? `${cb}(${json})` : json;
    const mime = cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;
    return ContentService.createTextOutput(out).setMimeType(mime);
  } catch(err) {
    const cb   = (e.parameter||{}).callback||'';
    const json = JSON.stringify({error:err.message});
    const out  = cb ? `${cb}(${json})` : json;
    const mime = cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;
    return ContentService.createTextOutput(out).setMimeType(mime);
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.events) {
      body.events.forEach(ev => handleLineEvent(ev));
      return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
    }
    const result = handlePost(body);
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    Logger.log('doPost error: ' + err.message);
    return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
  }
}

// ── LINE Event Handler ────────────────────────────────────────

function handleLineEvent(ev) {
  if (!ev.source) return;
  const userId     = ev.source.userId;
  const replyToken = ev.replyToken;
  if (!userId) return;
  const cfg = getConfig();
  if (!cfg['lineUserId']) saveConfig({ ...cfg, lineUserId: userId });
  if (ev.type !== 'message' || ev.message.type !== 'text') return;
  handleChatbot(userId, replyToken, ev.message.text.trim());
}

// ── Chatbot State Machine (with TTL) ──────────────────────────

function getState(userId) {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty('state_' + userId);
    if (!raw) return null;
    const state = JSON.parse(raw);
    // ✅ TTL check — state เก่าเกิน 10 นาที = หมดอายุ
    if (!state.ts || (Date.now() - state.ts) > STATE_TTL_MS) {
      setState(userId, null);
      return null;
    }
    return state;
  } catch(e) { return null; }
}

function setState(userId, state) {
  const props = PropertiesService.getScriptProperties();
  if (state === null) {
    props.deleteProperty('state_' + userId);
  } else {
    state.ts = Date.now();   // ✅ timestamp สำหรับ TTL
    props.setProperty('state_' + userId, JSON.stringify(state));
  }
}

// ── Chatbot Main ──────────────────────────────────────────────

function handleChatbot(userId, replyToken, text) {
  const cfg   = getConfig();
  const token = cfg['lineToken'] || '';
  const state = getState(userId);
  const cards = getActiveCards();
  const cmd   = text.toLowerCase().trim();

  // ── State ก่อนเสมอ ────────────────────────────────────────
  if (state) {
    // ยกเลิกได้ทุก step
    if (cmd === 'ยกเลิก' || cmd === 'cancel' || cmd === '0') {
      setState(userId, null);
      replyLine(token, replyToken, '❌ ยกเลิกแล้วครับ\n\nพิมพ์ "เมนู" เพื่อดูตัวเลือกครับ');
      return;
    }

    // Step: เลือกบัตร (กรอกยอด / จ่ายแล้ว / ลบยอด)
    if (state.step === 'select_card') {
      // ✅ Fix Bug 1: ใช้ cardIds จาก state ถ้ามี (กรณี quick pay multi-match)
      //    ไม่มีก็ใช้ active cards ทั้งหมดตามเดิม
      const selectable = state.cardIds
        ? state.cardIds.map(id => cards.find(c => String(c.id) === String(id))).filter(Boolean)
        : cards;

      const num  = parseInt(text, 10);
      const card = (num >= 1 && num <= selectable.length) ? selectable[num - 1] : null;
      if (!card) {
        replyLine(token, replyToken, `⚠️ กรุณาพิมพ์ตัวเลข 1-${selectable.length} ครับ\n\nหรือพิมพ์ "ยกเลิก"`);
        return;
      }

      if (state.action === 'paid') {
        markPaid(card, token, replyToken, userId);
      } else if (state.action === 'delete') {
        const bkk = bkkNow();
        saveLog({ month: bkk.monthKey, cardId: card.id, cardName: card.name, amount: 0, paid: false, confirmed: false });
        setState(userId, null);
        replyLine(token, replyToken, `🗑️ ลบยอดแล้วครับ!\n\n💳 ${card.name}\nยอดถูก reset เป็น 0 แล้วครับ\n\nพิมพ์ "เมนู" เพื่อดูตัวเลือกครับ`);
      } else {
        setState(userId, { step: 'enter_amount', cardId: card.id, cardName: card.name });
        replyLine(token, replyToken, `💳 ${card.name}\n\nกรุณาพิมพ์ยอดใช้จ่ายครับ\n(ตัวเลขเท่านั้น เช่น 5000)\n\nหรือพิมพ์ "ยกเลิก"`);
      }
      return;
    }

    // Step: กรอกยอด
    if (state.step === 'enter_amount') {
      const amount = parseFloat(text.replace(/,/g, ''));
      if (isNaN(amount) || amount < 0) {
        replyLine(token, replyToken, '⚠️ กรุณาพิมพ์ตัวเลขเท่านั้นครับ\n\nเช่น 5000 หรือ 1250.50\n\nหรือพิมพ์ "ยกเลิก"');
        return;
      }
      const bkk     = bkkNow();
      const entries = getLog(bkk.monthKey);
      const entry   = entries.find(e => String(e.cardId) === String(state.cardId));
      const paid    = entry ? isPaid(entry) : false;
      saveLog({ month: bkk.monthKey, cardId: state.cardId, cardName: state.cardName, amount, paid, confirmed: false });
      setState(userId, null);
      replyLine(token, replyToken, `✅ บันทึกแล้วครับ!\n\n💳 ${state.cardName}\n📅 ${thaiMonthLabel(bkk.month0, bkk.year)}\n💵 ฿${fmt(amount)}\n\nพิมพ์ "เมนู" เพื่อดูตัวเลือกครับ`);
      return;
    }
  }

  // ── Quick Pay: "จ่าย <ชื่อบัตร>" ──────────────────────────
  if (cmd.startsWith('จ่าย ') || cmd.startsWith('pay ')) {
    const keyword = cmd.replace(/^(จ่าย |pay )/, '').toLowerCase();
    const matched = cards.filter(c => String(c.name).toLowerCase().includes(keyword));
    if (matched.length === 1) {
      markPaid(matched[0], token, replyToken, userId);
      return;
    } else if (matched.length > 1) {
      // ✅ Fix Bug 1: เก็บ cardIds ที่ match ไว้ใน state
      const list = matched.map((c,i) => `${i+1}. ${c.name}`).join('\n');
      setState(userId, { step: 'select_card', action: 'paid', cardIds: matched.map(c => c.id) });
      replyLine(token, replyToken, `📋 พบหลายบัตรที่ตรงกับ "${keyword}":\n\n${list}\n\nพิมพ์ตัวเลขที่ต้องการ หรือ "ยกเลิก"`);
      return;
    } else {
      replyLine(token, replyToken, `⚠️ ไม่พบบัตรที่ตรงกับ "${keyword}" ครับ\n\nพิมพ์ "เมนู" เพื่อดูตัวเลือกครับ`);
      return;
    }
  }

  // ── Commands หลัก ────────────────────────────────────────
  if (cmd === 'เมนู' || cmd === 'menu' || cmd === 'help' || cmd === 'ช่วยเหลือ') {
    setState(userId, null);
    replyLine(token, replyToken, menuMessage());
    return;
  }
  if (cmd === 'กรอกยอด' || cmd === 'จดยอด' || cmd === '1') {
    setState(userId, { step: 'select_card', action: 'amount' });
    replyLine(token, replyToken, selectCardMessage(cards, 'กรอกยอด'));
    return;
  }
  if (cmd === 'จ่ายแล้ว' || cmd === '2') {
    setState(userId, { step: 'select_card', action: 'paid' });
    replyLine(token, replyToken, selectCardMessage(cards, 'จ่ายแล้ว'));
    return;
  }
  if (cmd === 'สรุป' || cmd === 'ดูสรุป' || cmd === '3') {
    setState(userId, null);
    replyLine(token, replyToken, getTodaySummary(cards));
    return;
  }
  if (cmd === 'ลบยอด' || cmd === 'ลบ' || cmd === '4') {
    setState(userId, { step: 'select_card', action: 'delete' });
    replyLine(token, replyToken, selectCardMessage(cards, '🗑️ ลบยอด'));
    return;
  }
  if (cmd === 'ประวัติ' || cmd === 'history' || cmd === '5') {
    setState(userId, null);
    replyLine(token, replyToken, getHistory(cards));
    return;
  }
  if (cmd === 'ธนาคาร' || cmd === 'bank' || cmd === '6') {
    setState(userId, null);
    replyLine(token, replyToken, getBankSummary(cards));
    return;
  }
  if (cmd === 'ยกเลิก' || cmd === 'cancel' || cmd === '0') {
    setState(userId, null);
    replyLine(token, replyToken, '❌ ยกเลิกแล้วครับ\n\nพิมพ์ "เมนู" เพื่อดูตัวเลือกครับ');
    return;
  }

  // Default
  replyLine(token, replyToken, '👋 สวัสดีครับ!\n\nพิมพ์ "เมนู" เพื่อดูตัวเลือกครับ');
}

// ── Shared Actions ────────────────────────────────────────────

function isPaid(entry) {
  return entry.paid === true || entry.paid === 'TRUE';
}

function fmt(n) {
  return Number(n).toLocaleString('th-TH');
}

function getActiveCards() {
  return getCards().filter(c => c.active !== false && c.active !== 'FALSE');
}

function markPaid(card, token, replyToken, userId) {
  const bkk     = bkkNow();
  const entries = getLog(bkk.monthKey);
  const entry   = entries.find(e => String(e.cardId) === String(card.id));
  const amount  = entry ? entry.amount : 0;
  const confirmed = entry ? !!entry.confirmed : false;
  saveLog({ month: bkk.monthKey, cardId: card.id, cardName: card.name, amount, paid: true, confirmed });
  setState(userId, null);
  replyLine(token, replyToken, `✅ บันทึกแล้วครับ!\n\n💳 ${card.name}\n✅ จ่ายแล้ว${amount > 0 ? '\n💵 ฿' + fmt(amount) : ''}\n\nพิมพ์ "เมนู" เพื่อดูตัวเลือกครับ`);
}

// ── Helper Messages ───────────────────────────────────────────

function menuMessage() {
  return `💳 CC Tracker\n\n1️⃣ กรอกยอด\n2️⃣ จ่ายแล้ว\n3️⃣ ดูสรุปวันนี้\n4️⃣ ลบยอด\n5️⃣ ประวัติ 3 เดือน\n6️⃣ สรุปต่อธนาคาร\n\n💡 Quick: พิมพ์ "จ่าย kbank" ได้เลย`;
}

function selectCardMessage(cards, action) {
  const list = cards.map((c, i) => `${i+1}. ${c.name}`).join('\n');
  return `📋 ${action} — เลือกบัตรครับ\n\n${list}\n\nพิมพ์ตัวเลขที่ต้องการ หรือ "ยกเลิก"`;
}

function getTodaySummary(cards) {
  const bkk = bkkNow();
  let monthKey = bkk.monthKey;
  let entries  = getLog(monthKey);

  // ✅ ถ้าเดือนปัจจุบันยังไม่มีใครกรอกยอดเลย → fallback ไปเดือนล่าสุดที่มีข้อมูลจริง
  //    (กันเคสยอดค้างชำระเดือนก่อนหน้ายังไม่ถูกจ่าย)
  if (!entries.some(e => parseFloat(e.amount) > 0)) {
    const allEntries = getLog('');
    const months = [...new Set(allEntries.map(e => String(e.month)).filter(m => m.match(/^\d{4}-\d{2}$/)))].sort().reverse();
    if (months.length > 0) {
      monthKey = months[0];
      entries  = getLog(monthKey);
    }
  }

  const [mkYear, mkMonth] = monthKey.split('-').map(Number);
  const mon = thaiMonthLabel(mkMonth - 1, mkYear);

  let total=0, paidAmt=0, lines=[];
  cards.forEach(c => {
    const e = entries.find(x => String(x.cardId) === String(c.id));
    if (!e || !e.amount || parseFloat(e.amount) <= 0) return;
    const a = parseFloat(e.amount);
    total  += a;
    const paid = isPaid(e);
    if (paid) paidAmt += a;
    lines.push(`${paid?'✅':'🔲'} ${c.name}: ฿${fmt(a)}`);
  });
  if (!lines.length) return `💳 CC Tracker\n📅 ${mon}\n\nยังไม่มียอดกรอกเดือนนี้ครับ`;
  return `💳 CC Tracker\n📅 ${mon}\n\n${lines.join('\n')}\n\n💰 รวม  ฿${fmt(total)}\n✅ จ่าย  ฿${fmt(paidAmt)}\n🔲 ค้าง  ฿${fmt(total-paidAmt)}`;
}

// ── ประวัติ 3 เดือนย้อนหลัง ──────────────────────────────────

function getHistory(cards) {
  const bkk    = bkkNow();
  const months = [];
  for (let i = 0; i < 3; i++) {
    let y = bkk.year;
    let m = bkk.month0 - i;
    if (m < 0) { m += 12; y--; }
    months.push({
      key:   `${y}-${String(m+1).padStart(2,'0')}`,
      label: thaiMonthLabel(m, y)
    });
  }

  const allEntries = getLog('');
  let msg = `📅 ประวัติ 3 เดือนย้อนหลัง\n`;

  months.forEach(({ key, label }) => {
    const entries = allEntries.filter(e => e.month === key);
    let total = 0, paid = 0;
    cards.forEach(c => {
      const e   = entries.find(x => String(x.cardId) === String(c.id));
      const amt = e ? parseFloat(e.amount) || 0 : 0;
      total    += amt;
      if (e && isPaid(e)) paid += amt;
    });
    if (total > 0) {
      msg += `\n📌 ${label}\n`;
      msg += `   💰 รวม  ฿${fmt(total)}\n`;
      msg += `   ✅ จ่าย ฿${fmt(paid)}\n`;
      msg += `   🔲 ค้าง ฿${fmt(total-paid)}`;
    } else {
      msg += `\n📌 ${label}\n   (ไม่มีข้อมูล)`;
    }
  });

  return msg;
}

// ── สรุปต่อธนาคาร ────────────────────────────────────────────

function getBankSummary(cards) {
  const bkk     = bkkNow();
  const entries = getLog(bkk.monthKey);
  const mon     = thaiMonthLabel(bkk.month0, bkk.year);

  const bankMap = {};
  cards.forEach(c => {
    const e   = entries.find(x => String(x.cardId) === String(c.id));
    const amt = e ? parseFloat(e.amount) || 0 : 0;
    if (amt <= 0) return;
    const bank = String(c.name).split(/[-_\s]/)[0].toUpperCase();
    if (!bankMap[bank]) bankMap[bank] = { total: 0, paid: 0, cards: [] };
    bankMap[bank].total += amt;
    const p = e && isPaid(e);
    if (p) bankMap[bank].paid += amt;
    bankMap[bank].cards.push(`  ${p?'✅':'🔲'} ${c.name}: ฿${fmt(amt)}`);
  });

  const banks = Object.keys(bankMap);
  if (!banks.length) return `💳 CC Tracker\n📅 ${mon}\n\nยังไม่มียอดกรอกเดือนนี้ครับ`;

  banks.sort((a,b) => bankMap[b].total - bankMap[a].total);

  let msg = `🏦 สรุปต่อธนาคาร\n📅 ${mon}\n`;
  let grandTotal = 0, grandPaid = 0;
  banks.forEach(bank => {
    const b = bankMap[bank];
    grandTotal += b.total;
    grandPaid  += b.paid;
    msg += `\n💳 ${bank}\n`;
    msg += b.cards.join('\n') + '\n';
    msg += `   รวม ฿${fmt(b.total)}`;
    if (b.paid > 0) msg += ` (จ่ายแล้ว ฿${fmt(b.paid)})`;
  });

  msg += `\n\n💰 รวมทั้งหมด ฿${fmt(grandTotal)}`;
  msg += `\n✅ จ่ายแล้ว  ฿${fmt(grandPaid)}`;
  msg += `\n🔲 ค้างชำระ  ฿${fmt(grandTotal-grandPaid)}`;
  return msg;
}

// ── LINE API ──────────────────────────────────────────────────

function replyLine(token, replyToken, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify({ replyToken, messages: [{ type: 'text', text }] }),
    muteHttpExceptions: true
  });
}

function sendLineMessage(token, userId, message) {
  if (!token || !userId) { Logger.log('sendLineMessage: missing token or userId'); return false; }
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify({ to: userId, messages: [{ type: 'text', text: message }] }),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code !== 200) Logger.log('LINE push error ' + code + ': ' + res.getContentText());
  return code === 200;
}

// ✅ ใหม่ v3.5: ส่ง Flex Message (การ์ด) แทนข้อความ text ธรรมดา
// altText คือข้อความที่โชว์ในหน้า notification/list ของ LINE (ต้อง <= 400 ตัวอักษร)
// contents คือ Flex bubble object (ดู buildNotifyFlex ด้านล่าง)
function sendLineFlex(token, userId, altText, contents) {
  if (!token || !userId) { Logger.log('sendLineFlex: missing token or userId'); return false; }
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify({
      to: userId,
      messages: [{ type: 'flex', altText: String(altText).slice(0, 400), contents }]
    }),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code !== 200) Logger.log('LINE flex push error ' + code + ': ' + res.getContentText());
  return code === 200;
}

// ✅ ใหม่ v3.3: ยิงข้อความทดสอบแบบ force เสมอ ไม่สนเงื่อนไขบัตร/วันครบกำหนดใดๆ
// ใช้กับปุ่ม "ทดสอบส่ง LINE" ใน dashboard โดยเฉพาะ — แยกจาก checkAndNotify()
// ซึ่งเป็น cron job ที่จะเงียบ (ไม่ส่งอะไรเลย) ถ้าไม่มีบัตรเข้าเงื่อนไข
function sendTestLineMessage() {
  const cfg    = getConfig();
  const token  = cfg['lineToken']  || '';
  const userId = cfg['lineUserId'] || '';
  if (!token)  return { error: 'lineToken not configured' };
  if (!userId) return { error: 'lineUserId not configured' };
  const bkk = bkkNow();
  const msg = `🔔 CC Tracker — ทดสอบระบบ\n📅 ${bkk.dateStr}\n\nถ้าคุณเห็นข้อความนี้ แปลว่า LINE เชื่อมต่อสำเร็จ ✅`;
  const sent = sendLineMessage(token, userId, msg);
  return { sent };
}

// หมายเหตุ: /followers/ids ใช้ได้เฉพาะ verified/premium account
// สำหรับ free account ให้ดู lineUserId จาก Config (บันทึกอัตโนมัติเมื่อทักแชท)
function getMyUserId() {
  const cfg   = getConfig();
  const token = cfg['lineToken'] || '';
  if (!token) return { error: 'lineToken not set in Config sheet' };
  if (cfg['lineUserId']) return { userId: cfg['lineUserId'], source: 'config' };
  const res  = UrlFetchApp.fetch('https://api.line.me/v2/bot/followers/ids', {
    headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    return { error: 'followers API ใช้ได้เฉพาะ verified account — ให้ทักแชทหา Bot ก่อน 1 ครั้ง ระบบจะบันทึก userId อัตโนมัติ' };
  }
  const data = JSON.parse(res.getContentText());
  Logger.log('User IDs: ' + JSON.stringify(data));
  return data;
}

// ── Get/Post Handlers ─────────────────────────────────────────

function handleGet(action, params) {
  if (action === 'getCards')         return getCards();
  if (action === 'getLog')           return getLog(params.month || '');
  if (action === 'getConfig')        return getConfig();
  if (action === 'checkNotify')      return checkAndNotify();
  if (action === 'testLineMessage')  return sendTestLineMessage();
  if (action === 'getMyUserId')      return getMyUserId();
  if (action === 'getLineUserId')    return { userId: getConfig()['lineUserId'] || '' };
  if (action === 'monthlySummary')   return sendMonthlySummary();
  if (action === 'installTriggers')  return installAllTriggers();
  return { error: 'Unknown action: ' + action };
}

function handlePost(body) {
  const a = body.action || '';
  if (a === 'saveCards')  return saveCards(body.cards);
  if (a === 'saveConfig') return saveConfig(body.config);
  if (a === 'saveLog')    return saveLog(body.entry);
  return { error: 'Unknown action: ' + a };
}

// ── Cards ─────────────────────────────────────────────────────

function getCards() {
  const sh   = getOrCreate(SHEET_CARDS, ['id','name','cycleDay','payDay','active']);
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1).map(r => ({ id:r[0], name:r[1], cycleDay:r[2], payDay:r[3], active:r[4] }));
}

function saveCards(cards) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getOrCreate(SHEET_CARDS, ['id','name','cycleDay','payDay','active']);
    sh.clearContents();
    sh.appendRow(['id','name','cycleDay','payDay','active']);
    cards.forEach(c => sh.appendRow([c.id, c.name, c.cycleDay, c.payDay, c.active !== false]));
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ── Monthly Log ───────────────────────────────────────────────

function normalizeMonth(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, TZ, 'yyyy-MM');
  }
  const s = String(val);
  if (s.length > 7) {
    const d = new Date(s);
    if (!isNaN(d)) {
      return Utilities.formatDate(d, TZ, 'yyyy-MM');
    }
  }
  return s;
}

// ✅ v3.3: เพิ่มคอลัมน์ confirmed (index 6 / column G)
// ⚠️ ถ้า sheet MonthlyLog มีอยู่ก่อนแล้ว (สร้างจาก schema เก่าที่ไม่มี confirmed)
//    ต้องเข้าไปเพิ่มหัวคอลัมน์ "confirmed" ที่ G1 ด้วยตัวเองครั้งเดียว
//    เพราะ getOrCreate จะไม่ไป alter schema ของ sheet ที่มีอยู่แล้ว
const LOG_HEADERS = ['month','cardId','cardName','amount','paid','updatedAt','confirmed'];

function getLog(monthKey) {
  const sh   = getOrCreate(SHEET_LOG, LOG_HEADERS);
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const data = rows.slice(1).map(r => ({
    month:     normalizeMonth(r[0]),
    cardId:    r[1],
    cardName:  r[2],
    amount:    r[3],
    paid:      r[4],
    updatedAt: r[5],
    confirmed: r[6] === true || r[6] === 'TRUE'
  }));
  return monthKey ? data.filter(r => r.month === monthKey) : data;
}

function saveLog(entry) {
  // ✅ LockService ป้องกัน race condition (trigger + webhook พร้อมกัน)
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh   = getOrCreate(SHEET_LOG, LOG_HEADERS);
    const rows = sh.getDataRange().getValues();
    const now  = new Date().toISOString();
    let month  = normalizeMonth(entry.month) || String(entry.month || '');
    const confirmed = !!entry.confirmed;

    for (let i = 1; i < rows.length; i++) {
      if (normalizeMonth(rows[i][0]) === month && String(rows[i][1]) === String(entry.cardId)) {
        sh.getRange(i+1,1,1,7).setValues([[month, entry.cardId, entry.cardName, entry.amount, entry.paid, now, confirmed]]);
        sh.getRange(i+1,1).setNumberFormat('@');
        return { ok:true, op:'updated' };
      }
    }
    sh.appendRow([month, entry.cardId, entry.cardName, entry.amount, entry.paid, now, confirmed]);
    sh.getRange(sh.getLastRow(),1).setNumberFormat('@');
    return { ok:true, op:'inserted' };
  } finally {
    lock.releaseLock();
  }
}

// ── Config (with per-request cache) ───────────────────────────

let _configCache = null;   // cache ต่อ 1 execution

function getConfig() {
  if (_configCache) return _configCache;
  const sh   = getOrCreate(SHEET_CONFIG, ['key','value']);
  const rows = sh.getDataRange().getValues();
  const cfg  = {};
  rows.slice(1).forEach(r => { cfg[r[0]] = r[1]; });
  _configCache = cfg;
  return cfg;
}

function saveConfig(config) {
  // ✅ LockService ป้องกัน clearContents ชนกัน
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getOrCreate(SHEET_CONFIG, ['key','value']);
    sh.clearContents();
    sh.appendRow(['key','value']);
    Object.entries(config).forEach(([k,v]) => sh.appendRow([k, v]));
    _configCache = config;   // update cache
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ── Flex Message Builder (daily notify) ────────────────────────
// สร้าง Flex bubble การ์ดเดียว รวม 3 ส่วน: ยังไม่จ่าย / จ่ายแล้ว / วันนี้สรุปยอด
// สีของแต่ละบัตรใน "ยังไม่จ่าย" ไล่ตามความเร่งด่วน: แดง(วันนี้)/ส้ม(พรุ่งนี้)/เหลือง(ยังมีเวลา)
function buildNotifyFlex(bkk, notYetPaid, alreadyPaid, cycleToday) {
  const urgentCount = notYetPaid.filter(u => u.diff <= 1).length;
  const body = [];

  if (notYetPaid.length > 0) {
    body.push({
      type: 'text',
      text: urgentCount > 0 ? `🔔 ยังไม่จ่าย (ด่วน ${urgentCount} บัตร!)` : '🔔 ยังไม่จ่าย',
      weight: 'bold', size: 'sm', color: urgentCount > 0 ? '#DC2626' : '#374151'
    });
    body.push({ type: 'separator', margin: 'sm' });

    notYetPaid.forEach((u, i) => {
      const color    = u.diff === 0 ? '#DC2626' : u.diff === 1 ? '#EA580C' : '#CA8A04';
      const dayLabel = u.diff === 0 ? 'วันนี้เลย!' : u.diff === 1 ? `พรุ่งนี้ (วันที่ ${u.payDay})` : `อีก ${u.diff} วัน (วันที่ ${u.payDay})`;
      body.push({
        type: 'box', layout: 'vertical', margin: i === 0 ? 'md' : 'lg', spacing: 'xs',
        contents: [
          { type: 'box', layout: 'horizontal', contents: [
            { type: 'text', text: u.name, weight: 'bold', size: 'md', color: '#111827', flex: 3, wrap: true },
            { type: 'text', text: '฿' + fmt(u.amount), weight: 'bold', size: 'md', color: color, align: 'end', flex: 2 }
          ]},
          { type: 'text', text: '⏰ ' + dayLabel, size: 'xs', color: color }
        ]
      });
    });
  }

  if (alreadyPaid.length > 0) {
    body.push({ type: 'separator', margin: 'lg' });
    body.push({ type: 'text', text: '✅ จ่ายแล้ว', weight: 'bold', size: 'sm', color: '#059669', margin: 'lg' });
    alreadyPaid.forEach(u => {
      body.push({
        type: 'box', layout: 'horizontal', margin: 'sm',
        contents: [
          { type: 'text', text: u.name, size: 'sm', color: '#374151', flex: 3, wrap: true },
          { type: 'text', text: u.amount > 0 ? '฿' + fmt(u.amount) : '-', size: 'sm', color: '#059669', align: 'end', flex: 2 }
        ]
      });
    });
  }

  if (cycleToday.length > 0) {
    body.push({ type: 'separator', margin: 'lg' });
    body.push({ type: 'text', text: '🗓️ วันนี้สรุปยอด', weight: 'bold', size: 'sm', color: '#7C3AED', margin: 'lg' });
    cycleToday.forEach(c => {
      body.push({ type: 'text', text: '📋 ' + c.name, size: 'sm', color: '#374151', margin: 'sm', wrap: true });
    });
    body.push({ type: 'text', text: '💡 อย่าลืมกรอกยอดในแอป!', size: 'xs', color: '#9CA3AF', margin: 'md', wrap: true });
  }

  const bubble = {
    type: 'bubble',
    size: 'giga',
    header: {
      type: 'box', layout: 'vertical', backgroundColor: '#111827', paddingAll: '20px', spacing: 'xs',
      contents: [
        { type: 'text', text: '💳 CC Tracker', color: '#FFFFFF', weight: 'bold', size: 'lg' },
        { type: 'text', text: '📅 ' + bkk.dateStr, color: '#9CA3AF', size: 'sm' }
      ]
    },
    body: { type: 'box', layout: 'vertical', paddingAll: '20px', spacing: 'sm', contents: body }
  };

  const cfg = getConfig();
  if (cfg['appUrl']) {
    bubble.footer = {
      type: 'box', layout: 'vertical', paddingAll: '12px',
      contents: [{
        type: 'button', style: 'primary', color: '#111827', height: 'sm',
        action: { type: 'uri', label: 'เปิดแอป', uri: cfg['appUrl'] }
      }]
    };
  }

  return bubble;
}

// ── Auto Notify ───────────────────────────────────────────────

function checkAndNotify() {
  const cfg    = getConfig();
  const token  = cfg['lineToken']  || '';
  const userId = cfg['lineUserId'] || '';
  if (!token)  return { skipped: 'lineToken not configured' };
  if (!userId) return { skipped: 'lineUserId not configured' };

  const notifyDays = parseInt(cfg['notifyDaysBefore'] || '3');
  const cards      = getActiveCards();
  const bkk        = bkkNow();   // ✅ Bangkok timezone เสมอ
  const logEntries = getLog(bkk.monthKey);
  const paidIds    = new Set(logEntries.filter(e => isPaid(e)).map(e => String(e.cardId)));

  function clampDay(year, month0, day) {
    return Math.min(day, new Date(year, month0 + 1, 0).getDate());
  }

  const notYetPaid = [], alreadyPaid = [];
  cards.forEach(c => {
    const payDay = parseInt(c.payDay);
    if (!payDay) return;
    const realPayDay = clampDay(bkk.year, bkk.month0, payDay);
    if (bkk.day <= realPayDay) {
      const diff   = realPayDay - bkk.day;
      const entry  = logEntries.find(e => String(e.cardId) === String(c.id));
      const amount = entry ? parseFloat(entry.amount) || 0 : 0;
      if (paidIds.has(String(c.id))) {
        alreadyPaid.push({ name: c.name, payDay: realPayDay, amount });
      } else if (diff <= notifyDays) {
        if (!entry || amount <= 0) return;
        notYetPaid.push({ name: c.name, payDay: realPayDay, diff, amount });
      }
    }
  });

  const cycleToday = cards.filter(c => {
    if (!c.cycleDay) return false;
    return clampDay(bkk.year, bkk.month0, parseInt(c.cycleDay)) === bkk.day;
  });

  if (!notYetPaid.length && !alreadyPaid.length && !cycleToday.length) {
    return { notified: 0, message: 'ไม่มีอะไรต้องแจ้งเตือน' };
  }

  notYetPaid.sort((a, b) => a.diff - b.diff);

  // ✅ v3.5: ส่งเป็น Flex Message (การ์ด) แทน text ธรรมดา
  const urgentCount = notYetPaid.filter(u => u.diff <= 1).length;
  const altText = notYetPaid.length > 0
    ? `🔔 มี ${notYetPaid.length} บัตรใกล้ครบกำหนด${urgentCount > 0 ? ' (ด่วน ' + urgentCount + ' บัตร!)' : ''}`
    : (cycleToday.length > 0 ? `🗓️ วันนี้สรุปยอด ${cycleToday.length} บัตร` : '💳 CC Tracker อัปเดตวันนี้');

  const flex = buildNotifyFlex(bkk, notYetPaid, alreadyPaid, cycleToday);
  const sent = sendLineFlex(token, userId, altText, flex);
  return { notified: notYetPaid.length + cycleToday.length, sent };
}

// ── Monthly LINE Summary ──────────────────────────────────────

function sendMonthlySummary(forceMonth) {
  const cfg    = getConfig();
  const token  = cfg['lineToken']  || '';
  const userId = cfg['lineUserId'] || '';
  if (!token || !userId) return { skipped: 'LINE not configured' };

  const bkk = bkkNow();   // ✅ Bangkok timezone
  let monthKey;
  if (forceMonth) {
    monthKey = forceMonth;
  } else {
    // เดือนที่แล้ว (เพราะ trigger รันวันที่ 1 — สรุปเดือนที่เพิ่งจบ)
    let pm = bkk.month0 - 1, py = bkk.year;
    if (pm < 0) { pm = 11; py--; }
    const prevKey = `${py}-${String(pm+1).padStart(2,'0')}`;
    // ถ้าเดือนที่แล้วมีข้อมูล ใช้เลย ไม่มีค่อย fallback หาเดือนล่าสุดที่มีข้อมูล
    const prevEntries = getLog(prevKey);
    if (prevEntries.some(e => parseFloat(e.amount) > 0)) {
      monthKey = prevKey;
    } else {
      const allEntries = getLog('');
      const months = [...new Set(allEntries.map(e => String(e.month)).filter(m => m.match(/^\d{4}-\d{2}$/)))].sort().reverse();
      monthKey = months.length > 0 ? months[0] : prevKey;
    }
  }

  const entries = getLog(monthKey);
  const cards   = getCards();
  const [mkYear, mkMonth] = monthKey.split('-').map(Number);
  const monthLabel = thaiMonthLabel(mkMonth - 1, mkYear);

  let total = 0;
  const lines = [];
  cards.forEach(c => {
    const e   = entries.find(x => String(x.cardId) === String(c.id));
    const amt = parseFloat(e ? e.amount : 0) || 0;
    if (amt > 0) {
      total += amt;
      lines.push(`${(e && isPaid(e)) ? '✅' : '🔲'} ${c.name}: ฿${fmt(amt)}`);
    }
  });

  const paidLines   = lines.filter(l => l.startsWith('✅'));
  const unpaidLines = lines.filter(l => l.startsWith('🔲'));
  const paidTotal   = cards.reduce((s, c) => {
    const e = entries.find(x => String(x.cardId) === String(c.id));
    return s + ((e && isPaid(e)) ? (parseFloat(e.amount) || 0) : 0);
  }, 0);

  const msg = [
    `💳 CC Tracker`, `สรุปยอดเดือน ${monthLabel}`, '',
    ...(unpaidLines.length ? ['🔲 ค้างชำระ:', ...unpaidLines.map(l => '  ' + l.replace('🔲 ', ''))] : []),
    ...(paidLines.length   ? ['', '✅ จ่ายแล้ว:', ...paidLines.map(l => '  ' + l.replace('✅ ', ''))] : []),
    ...(lines.length === 0 ? ['  (ไม่มีข้อมูลเดือนนี้)'] : []),
    '',
    `💰 รวม  ฿${fmt(total)}`,
    `✅ จ่าย  ฿${fmt(paidTotal)}`,
    `🔲 ค้าง  ฿${fmt(total - paidTotal)}`,
  ].join('\n');

  return { sent: sendLineMessage(token, userId, msg), monthKey, total, cardCount: lines.length };
}

// ── Trigger Setup ─────────────────────────────────────────────

function installAllTriggers() {
  const cfg         = getConfig();
  const notifyHour  = parseInt(cfg['notifyHour']  || '8');
  const summaryHour = parseInt(cfg['summaryHour'] || '9');
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('checkAndNotify').timeBased().everyDays(1).atHour(notifyHour).create();
  ScriptApp.newTrigger('sendMonthlySummary').timeBased().onMonthDay(1).atHour(summaryHour).create();
  Logger.log(`✓ Triggers: notify=${notifyHour}:00  summary=${summaryHour}:00`);
  return { ok: true, notifyHour, summaryHour };
}

function installDailyTrigger() { installAllTriggers(); }
