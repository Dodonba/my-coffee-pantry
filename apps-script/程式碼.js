/**
 * 咖啡豆櫃 Web App
 * 路由：data.action
 *   - 'add' → 新增豆子，寫入「豆子」分頁
 *   - 'log' → 沖煮日誌，寫入「沖煮日誌」分頁
 */

const SHEET_BEAN = '豆子';
const SHEET_LOG  = '沖煮日誌';

const BEAN_HEADERS = [
  '時間', '匿名ID', '動作', '豆名', '烘焙度', '處理法',
  '密度', '品種', 'brewProfile', '主打風味', '主要成分', '豆袋總重', '莊園'
];
const LOG_HEADERS = [
  '時間', '匿名ID', '豆名', 'brewProfile', '使用食譜', '粉量', '評價', '今天精神', '今天行程'
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'add' && data.bean)     return logBeanAdd(data);
    if (data.action === 'log' && data.brewLog)  return logBrew(data);

    return json({ ok: false, msg: 'unknown action: ' + (data.action || 'none') });
  } catch (err) {
    return json({ error: String(err) });
  }
}

function doGet() {
  return json({ ok: true, msg: 'Coffee pantry Web App is alive.' });
}

// ===== 新增豆子寫入 Sheet =====
function logBeanAdd(data) {
  const sh = getSheet(SHEET_BEAN, BEAN_HEADERS);
  const b = data.bean || {};
  sh.appendRow([
    new Date(),
    data.userId || '',
    data.action || 'add',
    b.name || '',
    b.roast || '',
    b.process || '',
    b.density || '',
    Array.isArray(b.variety)      ? b.variety.join(' / ')      : '',
    b.brewProfile || '',
    Array.isArray(b.blendFlavors) ? b.blendFlavors.join(' / ') : '',
    Array.isArray(b.blendOrigins) ? b.blendOrigins.join(' / ') : '',
    b.weightTotal || '',
    Array.isArray(b.estates)      ? b.estates.join(' / ')      : '',
  ]);
  return json({ ok: true });
}

// ===== 沖煮日誌寫入 Sheet =====
function logBrew(data) {
  const sh = getSheet(SHEET_LOG, LOG_HEADERS);
  const l = data.brewLog || {};
  sh.appendRow([
    new Date(),
    data.userId || '',
    l.beanName || '',
    l.brewProfile || '',
    l.recipe || '',
    l.dose || '',
    l.rating || '',
    l.mood || '',
    l.schedule || '',
  ]);
  return json({ ok: true });
}

// ===== 工具 =====
function getSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * 一次性建立咖啡豆櫃工具回饋表單
 * 執行步驟：在 Apps Script 編輯器、上方執行下拉選 buildFeedbackForm → 按 ▶
 * 跳授權對話 → 允許 → 看下方執行日誌、複製 publishedUrl
 */
function buildFeedbackForm() {
  const form = FormApp.create('☕ 咖啡豆櫃工具・使用回饋');
  form.setDescription('感謝你用咖啡豆櫃！這份問卷約 3 分鐘、幫我們把工具改得更貼近你的需求 ❤\n所有題目都是選填、看你想分享多少都可以。');
  form.setCollectEmail(false);
  form.setAllowResponseEdits(true);
  form.setLimitOneResponsePerUser(false);
  form.setShowLinkToRespondAgain(false);

  // ===== 1. 認識你 =====
  form.addSectionHeaderItem().setTitle('1. 認識你').setHelpText('讓我們知道你是哪種咖啡人');

  form.addMultipleChoiceItem()
    .setTitle('你跟咖啡的關係')
    .setChoiceValues(['日常喝（不太鑽研）', '咖啡玩家（會買器材／追新豆）', '咖啡師／咖啡店相關', '還在入坑']);

  form.addMultipleChoiceItem()
    .setTitle('你主要的磨豆機')
    .setChoiceValues([
      'Comandante C40',
      '1Zpresso K-Ultra',
      '1Zpresso K-Pro / K-Plus',
      'Kingrinder K6',
      'Mahlkönig EK43（家用罕見）',
      '電動家用磨豆機（Timemore Sculptor / 小富士 等）',
      '手搖入門款（Hario Mini Mill 等）',
      '其他／還沒買'
    ]);

  form.addCheckboxItem()
    .setTitle('你家裡有的沖煮器材（可複選）')
    .setChoiceValues([
      'V60 / Origami 濾杯',
      'Kalita 平底濾杯',
      'Hario Switch',
      'Aeropress 愛樂壓',
      'French Press 法壓壺',
      'Moka 摩卡壺',
      'Espresso 義式機',
      '冷萃壺 / 急冷壺'
    ]);

  // ===== 2. 工具使用回饋 =====
  form.addPageBreakItem().setTitle('2. 你怎麼用這個工具').setHelpText('哪些功能有用、哪些可以砍');

  form.addCheckboxItem()
    .setTitle('你最常用的功能（可複選）')
    .setChoiceValues([
      '☕ 我的豆櫃（管理手上的豆）',
      '⭐ 今日推薦面板',
      '🫖 卡片裡的沖法食譜（4 張卡）',
      '📝 沖煮日誌（記錄/評價）',
      '🛠 風味推薦工具（簡易模式）',
      '🛠 風味推薦工具（進階模式 29 風味）',
      '📚 咖啡資訊 12 章',
      '主打風味 chip 標籤',
      '常用粉量 / 磨豆機偏好設定'
    ]);

  form.addParagraphTextItem()
    .setTitle('你覺得最有幫助的「一個」功能是？為什麼？');

  form.addParagraphTextItem()
    .setTitle('哪些功能你「完全沒用過」或覺得多餘？');

  form.addParagraphTextItem()
    .setTitle('你最希望加什麼新功能？')
    .setHelpText('例如：行事曆提醒、社群分享、購買清單、烘焙日通知…');

  // ===== 3. 整體體驗 =====
  form.addPageBreakItem().setTitle('3. 整體體驗').setHelpText('幫忙評個分');

  form.addScaleItem()
    .setTitle('整體使用體驗')
    .setBounds(1, 5)
    .setLabels('不太好用', '很好用');

  form.addScaleItem()
    .setTitle('願意推薦給朋友的可能性')
    .setBounds(0, 10)
    .setLabels('完全不會', '肯定推薦');

  form.addMultipleChoiceItem()
    .setTitle('你最常用哪個裝置打開這個工具')
    .setChoiceValues(['手機', '平板', '電腦', '都差不多']);

  form.addParagraphTextItem()
    .setTitle('有遇到 bug、操作卡住、看不懂的地方嗎？');

  // ===== 4. 分享豆袋 =====
  form.addPageBreakItem().setTitle('4. 分享豆袋（選填）').setHelpText('想看看大家都在喝什麼豆！');

  form.addParagraphTextItem()
    .setTitle('你最近喝到「最讚的一支豆」是？')
    .setHelpText('豆名、產地、烘焙店、處理法、為什麼覺得好喝 ── 寫多寫少都行');

  form.addParagraphTextItem()
    .setTitle('你都從哪邊買豆？常去的咖啡店有哪幾家？')
    .setHelpText('店名、地點、推薦的招牌豆');

  form.addParagraphTextItem()
    .setTitle('你目前的豆櫃裡有幾支豆？大概都是什麼類型？');

  // ===== 5. 聯絡 =====
  form.addPageBreakItem().setTitle('5. 想保持聯絡？（完全選填）');

  form.addMultipleChoiceItem()
    .setTitle('要不要在新版本上線時收到通知？')
    .setChoiceValues(['要，請留 email 我寄信給你', '不用']);

  form.addTextItem()
    .setTitle('Email（選填，只用於通知新版本，不會 spam）');

  form.addParagraphTextItem()
    .setTitle('任何其他想說的話 / 私訊作者');

  // ===== 完成 ── 印出 URL =====
  const publishedUrl = form.getPublishedUrl();
  const editUrl = form.getEditUrl();
  Logger.log('========== 建立完成 ==========');
  Logger.log('給使用者填的 URL（複製這個貼回工具）：');
  Logger.log(publishedUrl);
  Logger.log('');
  Logger.log('你自己編輯 / 看回應的 URL：');
  Logger.log(editUrl);
  Logger.log('================================');
  return { publishedUrl, editUrl };
}