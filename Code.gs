/**
 * 互動式極速問答挑戰網頁 — Google Apps Script 後端程式 (v2.0)
 * 
 * 負責：
 * 1. 透過 doGet(e) 進行伺服端資料預先注入 (SSR Data Injection)
 * 2. 讀取 Google Sheets 中的 Settings 與 Questions 工作表
 * 3. 提供安全容錯與預設測試題庫 (Fallback Mock Data)
 */

const SPREADSHEET_ID = '10UTkLWpqtYfr6Dux4KdXEOnGOlYj7hxiylkZefJyQoI';

const DEFAULT_SETTINGS = {
  app_title: '極速大挑戰',
  sub_title: '累積答對 5 題過關・每題限時 10 秒',
  bg_image: '',
  logo_image: '',
  drive_folder_id: ''
};

/**
 * Web App 進入點
 * @param {Object} e - GAS HTTP GET 事件物件
 * @returns {HtmlOutput}
 */
function doGet(e) {
  const payload = getQuizPayload();
  const template = HtmlService.createTemplateFromFile('index');

  // 避免 JSON 內容含有 </script> 等字元破壞 HTML 結構
  template.initialData = JSON.stringify(payload).replace(/</g, '\\u003c');

  const pageTitle = (payload.settings && payload.settings.app_title)
    ? String(payload.settings.app_title)
    : DEFAULT_SETTINGS.app_title;

  return template
    .evaluate()
    .setTitle(pageTitle)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 讀取試算表中的 Settings 與 Questions，組成前端所需資料結構
 * 若試算表尚未建立或無有效題目，自動回傳內建測試題庫
 * @returns {{ settings: Object, questions: Array, isFallback: boolean }}
 */
function getQuizPayload() {
  let settings = Object.assign({}, DEFAULT_SETTINGS);
  let questions = [];
  let isFallback = false;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!ss) {
      Logger.log('找不到綁定的 Google 試算表，使用預設測試題庫。');
      return {
        settings: settings,
        questions: getFallbackMockData(),
        isFallback: true
      };
    }

    // 1. 讀取 Settings 工作表
    const settingsSheet = ss.getSheetByName('Settings');
    if (settingsSheet && settingsSheet.getLastRow() >= 2) {
      const settingsValues = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 2).getValues();
      settingsValues.forEach(function (row) {
        const key = String(row[0] || '').trim();
        const value = String(row[1] || '').trim();
        if (key && value !== '') {
          settings[key] = value;
        }
      });
    }

    // 2. 讀取 Questions 工作表
    const questionsSheet = ss.getSheetByName('Questions');
    if (questionsSheet && questionsSheet.getLastRow() >= 2 && questionsSheet.getLastColumn() >= 1) {
      const allData = questionsSheet.getDataRange().getValues();
      const headers = allData[0].map(function (h) {
        return String(h || '').trim().toLowerCase();
      });

      const colIndex = {
        id: headers.indexOf('id'),
        type: headers.indexOf('type'),
        question_text: headers.indexOf('question_text'),
        question_image: headers.indexOf('question_image'),
        content_json: headers.indexOf('content_json')
      };

      // 確認必要欄位皆存在
      if (colIndex.id !== -1 && colIndex.type !== -1 && colIndex.question_text !== -1 && colIndex.content_json !== -1) {
        for (let i = 1; i < allData.length; i++) {
          const row = allData[i];
          const rowNumber = i + 1;

          const id = String(row[colIndex.id] || '').trim();
          const type = String(row[colIndex.type] || '').trim().toLowerCase();
          const questionText = String(row[colIndex.question_text] || '').trim();
          const questionImage = colIndex.question_image !== -1
            ? String(row[colIndex.question_image] || '').trim()
            : '';
          const rawJson = String(row[colIndex.content_json] || '').trim();

          // 略過空白列
          if (!id && !questionText && !rawJson) {
            continue;
          }

          if (!id || !questionText || (type !== 'choice' && type !== 'match') || !rawJson) {
            Logger.log('第 ' + rowNumber + ' 列欄位不完整或題型錯誤 (type=' + type + ')，已略過。');
            continue;
          }

          try {
            const parsedContent = JSON.parse(rawJson);

            if (!isValidQuestionContent(type, parsedContent)) {
              Logger.log('第 ' + rowNumber + ' 列 (id=' + id + ') content_json 結構不符合規範，已略過。');
              continue;
            }

            questions.push({
              id: id,
              type: type,
              question_text: questionText,
              question_image: questionImage,
              content: parsedContent
            });
          } catch (jsonErr) {
            Logger.log('第 ' + rowNumber + ' 列 (id=' + id + ') JSON 解析失敗：' + jsonErr.message);
          }
        }
      } else {
        Logger.log('Questions 工作表缺少必要標題欄位 (id, type, question_text, content_json)。');
      }
    }
  } catch (err) {
    Logger.log('讀取試算表發生例外錯誤：' + err.message);
  }

  if (questions.length === 0) {
    Logger.log('無有效試算表題目，啟用內建 Fallback 測試題庫。');
    questions = getFallbackMockData();
    isFallback = true;
  }

  return {
    settings: settings,
    questions: questions,
    isFallback: isFallback
  };
}

/**
 * 驗證解析後的 content_json 是否符合題型結構
 * @param {string} type - 'choice' | 'match'
 * @param {Object} content
 * @returns {boolean}
 */
function isValidQuestionContent(type, content) {
  if (!content || typeof content !== 'object') {
    return false;
  }

  if (type === 'choice') {
    return Array.isArray(content.options) &&
      content.options.length >= 2 &&
      typeof content.answer === 'string' &&
      content.answer.trim() !== '';
  }

  if (type === 'match') {
    return Array.isArray(content.pairs) &&
      content.pairs.length >= 2 &&
      content.pairs.every(function (p) {
        return p && p.left && p.right;
      });
  }

  return false;
}

/**
 * 內建預設測試題庫 (Fallback Mock Data)
 * 涵蓋選擇題 (choice) 與連連看 (match)，共 10 題
 * @returns {Array<Object>}
 */
function getFallbackMockData() {
  return [
    {
      id: 'Q01',
      type: 'choice',
      question_text: '光在真空中的傳播速率約為多少？',
      question_image: '',
      content: {
        options: [
          { id: 'A', text: '3 × 10⁵ 公尺/秒', image: '' },
          { id: 'B', text: '3 × 10⁸ 公尺/秒', image: '' },
          { id: 'C', text: '3 × 10¹⁰ 公尺/秒', image: '' },
          { id: 'D', text: '340 公尺/秒', image: '' }
        ],
        answer: 'B'
      }
    },
    {
      id: 'Q02',
      type: 'choice',
      question_text: 'HTML5 中，用來定義網頁主要導覽連結區塊的語意標籤為何？',
      question_image: '',
      content: {
        options: [
          { id: 'A', text: '<section>', image: '' },
          { id: 'B', text: '<aside>', image: '' },
          { id: 'C', text: '<nav>', image: '' },
          { id: 'D', text: '<header>', image: '' }
        ],
        answer: 'C'
      }
    },
    {
      id: 'M01',
      type: 'match',
      question_text: '請將下列「物理量」與其對應的「SI 國際單位」正確配對：',
      question_image: '',
      content: {
        pairs: [
          { left: { text: '電流', image: '' }, right: { text: '安培 (A)', image: '' } },
          { left: { text: '電阻', image: '' }, right: { text: '歐姆 (Ω)', image: '' } },
          { left: { text: '功率', image: '' }, right: { text: '瓦特 (W)', image: '' } }
        ]
      }
    },
    {
      id: 'Q03',
      type: 'choice',
      question_text: '地球大氣層中含量最多的氣體是哪一種？',
      question_image: '',
      content: {
        options: [
          { id: 'A', text: '氧氣 (O₂)', image: '' },
          { id: 'B', text: '氮氣 (N₂)', image: '' },
          { id: 'C', text: '二氧化碳 (CO₂)', image: '' },
          { id: 'D', text: '氬氣 (Ar)', image: '' }
        ],
        answer: 'B'
      }
    },
    {
      id: 'Q04',
      type: 'choice',
      question_text: '在 JavaScript 中，下列哪一個陣列方法會回傳一個「新陣列」而不改變原陣列？',
      question_image: '',
      content: {
        options: [
          { id: 'A', text: 'Array.prototype.push()', image: '' },
          { id: 'B', text: 'Array.prototype.splice()', image: '' },
          { id: 'C', text: 'Array.prototype.map()', image: '' },
          { id: 'D', text: 'Array.prototype.sort()', image: '' }
        ],
        answer: 'C'
      }
    },
    {
      id: 'M02',
      type: 'match',
      question_text: '請將下列「國家」與其「首都」進行配對：',
      question_image: '',
      content: {
        pairs: [
          { left: { text: '日本', image: '' }, right: { text: '東京', image: '' } },
          { left: { text: '法國', image: '' }, right: { text: '巴黎', image: '' } },
          { left: { text: '澳洲', image: '' }, right: { text: '坎培拉', image: '' } }
        ]
      }
    },
    {
      id: 'Q05',
      type: 'choice',
      question_text: '下列哪一個 HTTP 狀態碼代表「請求成功，資源已建立」？',
      question_image: '',
      content: {
        options: [
          { id: 'A', text: '200 OK', image: '' },
          { id: 'B', text: '201 Created', image: '' },
          { id: 'C', text: '304 Not Modified', image: '' },
          { id: 'D', text: '403 Forbidden', image: '' }
        ],
        answer: 'B'
      }
    },
    {
      id: 'Q06',
      type: 'choice',
      question_text: '聲音無法在下列哪一種介質中傳播？',
      question_image: '',
      content: {
        options: [
          { id: 'A', text: '真空', image: '' },
          { id: 'B', text: '空氣', image: '' },
          { id: 'C', text: '純水', image: '' },
          { id: 'D', text: '鋼鐵', image: '' }
        ],
        answer: 'A'
      }
    },
    {
      id: 'M03',
      type: 'match',
      question_text: '請將下列「化學元素符號」與「中文名稱」配對：',
      question_image: '',
      content: {
        pairs: [
          { left: { text: 'Fe', image: '' }, right: { text: '鐵', image: '' } },
          { left: { text: 'Au', image: '' }, right: { text: '金', image: '' } },
          { left: { text: 'Na', image: '' }, right: { text: '鈉', image: '' } }
        ]
      }
    },
    {
      id: 'Q07',
      type: 'choice',
      question_text: '下列哪一種資料結構採用「後進先出 (LIFO)」的運作原則？',
      question_image: '',
      content: {
        options: [
          { id: 'A', text: '佇列 (Queue)', image: '' },
          { id: 'B', text: '堆疊 (Stack)', image: '' },
          { id: 'C', text: '雜湊表 (Hash Table)', image: '' },
          { id: 'D', text: '二元搜尋樹 (BST)', image: '' }
        ],
        answer: 'B'
      }
    }
  ];
}

/**
 * 在試算表開啟時建立管理選單
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡ 極速問答後台')
    .addItem('📝 開啟題目編輯器', 'openQuestionEditor')
    .addItem('🖼️ Drive 圖片管理與快速改名', 'openDriveManager')
    .addItem('🔄 從 Drive 資料夾同步題庫', 'syncFromDriveMenu')
    .addSeparator()
    .addItem('⚙️ 初始化 Settings 與 Questions 工作表', 'initializeQuizSheets')
    .addToUi();
}

/**
 * 開啟題目編輯器（預設顯示「題目編輯器」分頁）
 */
function openQuestionEditor() {
  const template = HtmlService.createTemplateFromFile('editor');
  template.initialTab = 'editor';
  const html = template.evaluate()
    .setWidth(980)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '⚡ 極速問答 — 題目編輯器與 Drive 圖片工具');
}

/**
 * 開啟 Drive 圖片管理與快速改名（預設顯示「Drive 圖片庫」分頁）
 */
function openDriveManager() {
  const template = HtmlService.createTemplateFromFile('editor');
  template.initialTab = 'drive';
  const html = template.evaluate()
    .setWidth(980)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '🖼️ Drive 圖片管理與快速改名');
}

/**
 * 選單觸發：直接從 Settings 中的 drive_folder_id 同步題庫到 Questions
 */
function syncFromDriveMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
    const settings = readSettingsMap_(ss);
    if (!settings.drive_folder_id) {
      ui.alert('尚未設定 drive_folder_id！請先在 Settings 工作表填入資料夾 ID，或開啟「Drive 圖片管理」設定。');
      return;
    }
    const res = syncDriveImagesToQuestions(settings.drive_folder_id);
    ui.alert('同步完成！\n新增題目：' + res.added + ' 題\n更新題目：' + res.updated + ' 題\n略過未符合命名規範檔案：' + res.skippedFiles.length + ' 個');
  } catch (err) {
    ui.alert('同步失敗：' + err.message);
  }
}

/**
 * 內部共用：讀取 Settings 工作表為鍵值物件
 */
function readSettingsMap_(ss) {
  const settings = Object.assign({}, DEFAULT_SETTINGS);
  const settingsSheet = ss.getSheetByName('Settings');
  if (settingsSheet && settingsSheet.getLastRow() >= 2) {
    const values = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 2).getValues();
    values.forEach(function (row) {
      const k = String(row[0] || '').trim();
      const v = String(row[1] || '').trim();
      if (k && v !== '') {
        settings[k] = v;
      }
    });
  }
  return settings;
}

/**
 * 從 Google Drive 網址或 ID 字串中解析出純 Folder ID
 */
function extractDriveId_(input) {
  const str = String(input || '').trim();
  if (!str) return '';
  const folderMatch = str.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (folderMatch && folderMatch[1]) return folderMatch[1];
  const idMatch = str.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (idMatch && idMatch[1]) return idMatch[1];
  return str;
}

/**
 * 供前端編輯器初始化呼叫：回傳目前 Settings 與 Questions 資料
 */
function getEditorInitData() {
  const payload = getQuizPayload();
  return {
    settings: payload.settings,
    questions: payload.isFallback ? [] : payload.questions
  };
}

/**
 * 儲存或更新 Settings 工作表中的 drive_folder_id
 */
function saveDriveFolderId(folderInput) {
  const folderId = extractDriveId_(folderInput);
  if (!folderId) {
    throw new Error('請輸入有效的 Google Drive 資料夾 ID 或網址。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
  let settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('Settings');
    settingsSheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
  }

  const lastRow = settingsSheet.getLastRow();
  let updated = false;
  if (lastRow >= 2) {
    const keys = settingsSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0] || '').trim() === 'drive_folder_id') {
        settingsSheet.getRange(i + 2, 2).setValue(folderId);
        updated = true;
        break;
      }
    }
  }
  if (!updated) {
    settingsSheet.appendRow(['drive_folder_id', folderId]);
  }

  return folderId;
}

/**
 * 讀取指定 Drive 資料夾內的所有圖片檔案
 */
function listDriveFolderImages(folderInput) {
  const folderId = extractDriveId_(folderInput);
  if (!folderId) {
    throw new Error('請先輸入 Google Drive 資料夾 ID 或網址。');
  }

  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const images = [];
  const imageExtRegex = /\.(jpg|jpeg|png|gif|webp|bmp|svg)$/i;

  while (files.hasNext()) {
    const file = files.next();
    const mime = file.getMimeType() || '';
    const name = file.getName() || '';
    if (mime.indexOf('image/') === 0 || imageExtRegex.test(name)) {
      const fid = file.getId();
      images.push({
        id: fid,
        name: name,
        mimeType: mime,
        viewUrl: 'https://drive.google.com/file/d/' + fid + '/view',
        thumbUrl: 'https://lh3.googleusercontent.com/d/' + fid + '=s240'
      });
    }
  }

  images.sort(function (a, b) {
    return a.name.localeCompare(b.name, 'zh-Hant', { numeric: true });
  });

  return {
    folderId: folderId,
    folderName: folder.getName(),
    images: images
  };
}

/**
 * 修改 Drive 檔案名稱
 */
function renameDriveFile(fileId, newName) {
  const cleanName = String(newName || '').trim();
  if (!fileId || !cleanName) {
    throw new Error('檔案 ID 與新檔名不可為空。');
  }
  const file = DriveApp.getFileById(fileId);
  file.setName(cleanName);
  return {
    id: fileId,
    name: file.getName()
  };
}

/**
 * 儲存單筆題目至 Questions 工作表（若 ID 已存在則更新，否則新增一列）
 */
function saveQuestionToSheet(qObj) {
  if (!qObj || !qObj.id || !qObj.type || !qObj.question_text || !qObj.content) {
    throw new Error('題目資料不完整，請確認題號、題型、題目文字與選項/配對內容。');
  }

  const qId = String(qObj.id).trim();
  const qType = String(qObj.type).trim().toLowerCase();
  if (!isValidQuestionContent(qType, qObj.content)) {
    throw new Error('選項或配對內容格式不正確（選擇題需至少 2 個選項與正解；連連看需至少 2 組配對）。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Questions');
  if (!sheet) {
    sheet = ss.insertSheet('Questions');
    sheet.getRange(1, 1, 1, 5).setValues([['id', 'type', 'question_text', 'question_image', 'content_json']]);
  }

  const rowValues = [
    qId,
    qType,
    String(qObj.question_text).trim(),
    String(qObj.question_image || '').trim(),
    JSON.stringify(qObj.content)
  ];

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const idCol = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < idCol.length; i++) {
      if (String(idCol[i][0] || '').trim() === qId) {
        sheet.getRange(i + 2, 1, 1, 5).setValues([rowValues]);
        return { action: 'updated', id: qId };
      }
    }
  }

  sheet.appendRow(rowValues);
  return { action: 'added', id: qId };
}

/**
 * 掃描 Drive 資料夾內的圖片檔名，自動解析並同步寫入 Questions 工作表
 * 命名規範：
 * - 選擇題：{題號}_Q[_題目].ext, {題號}_{A-D}[_OK][_選項文字].ext
 * - 連連看：{題號}_Q[_題目].ext, {題號}_P{組別}_{L|R}[_文字].ext
 */
function syncDriveImagesToQuestions(folderInput) {
  const listing = listDriveFolderImages(folderInput);
  const questionMap = {};
  const skippedFiles = [];

  listing.images.forEach(function (img) {
    const baseName = img.name.replace(/\.[^.]+$/, '');
    const parts = baseName.split('_');
    if (parts.length < 2) {
      skippedFiles.push(img.name);
      return;
    }

    const qId = parts[0].trim();
    const role = parts[1].trim().toUpperCase();
    if (!qId || !role) {
      skippedFiles.push(img.name);
      return;
    }

    if (!questionMap[qId]) {
      questionMap[qId] = {
        id: qId,
        type: /^M/i.test(qId) ? 'match' : 'choice',
        question_text: '',
        question_image: '',
        optionsMap: {},
        answer: '',
        pairsMap: {}
      };
    }
    const qEntry = questionMap[qId];

    // 1. 題目附圖：{ID}_Q[_題目文字]
    if (role === 'Q') {
      qEntry.question_image = img.viewUrl;
      if (parts.length >= 3) {
        qEntry.question_text = parts.slice(2).join(' ').trim();
      }
      return;
    }

    // 2. 連連看配對：{ID}_P{組別}_{L|R}[_文字]
    const pairMatch = role.match(/^P(\d+)$/);
    if (pairMatch && parts.length >= 3) {
      const pairNum = pairMatch[1];
      const side = parts[2].trim().toUpperCase();
      if (side === 'L' || side === 'R') {
        qEntry.type = 'match';
        if (!qEntry.pairsMap[pairNum]) {
          qEntry.pairsMap[pairNum] = {
            left: { text: '', image: '' },
            right: { text: '', image: '' }
          };
        }
        const labelText = parts.length >= 4 ? parts.slice(3).join(' ').trim() : '';
        if (side === 'L') {
          qEntry.pairsMap[pairNum].left = { text: labelText, image: img.viewUrl };
        } else {
          qEntry.pairsMap[pairNum].right = { text: labelText, image: img.viewUrl };
        }
        return;
      }
    }

    // 3. 選擇題選項：{ID}_{A|B|C|D}[_OK][_選項文字]
    if (/^[A-D]$/.test(role)) {
      qEntry.type = 'choice';
      let restParts = parts.slice(2);
      if (restParts.length > 0 && restParts[0].trim().toUpperCase() === 'OK') {
        qEntry.answer = role;
        restParts = restParts.slice(1);
      }
      const optText = restParts.join(' ').trim();
      qEntry.optionsMap[role] = {
        id: role,
        text: optText,
        image: img.viewUrl
      };
      return;
    }

    skippedFiles.push(img.name);
  });

  let added = 0;
  let updated = 0;

  Object.keys(questionMap).sort().forEach(function (qId) {
    const entry = questionMap[qId];
    let content = null;

    if (entry.type === 'choice') {
      const options = Object.keys(entry.optionsMap).sort().map(function (k) {
        return entry.optionsMap[k];
      });
      if (options.length < 2 || !entry.answer) {
        return;
      }
      content = {
        options: options,
        answer: entry.answer
      };
    } else {
      const pairs = Object.keys(entry.pairsMap)
        .sort(function (a, b) { return Number(a) - Number(b); })
        .map(function (k) { return entry.pairsMap[k]; });
      if (pairs.length < 2) {
        return;
      }
      content = { pairs: pairs };
    }

    const res = saveQuestionToSheet({
      id: entry.id,
      type: entry.type,
      question_text: entry.question_text || ('題目 ' + entry.id),
      question_image: entry.question_image || '',
      content: content
    });

    if (res.action === 'added') added++;
    if (res.action === 'updated') updated++;
  });

  return {
    added: added,
    updated: updated,
    skippedFiles: skippedFiles
  };
}

/**
 * 一鍵建立標準化的 Settings 與 Questions 工作表與範例資料
 * （若工作表已存在則不會覆蓋現有資料）
 */
function initializeQuizSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;

  let settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('Settings');
    settingsSheet.getRange(1, 1, 6, 2).setValues([
      ['key', 'value'],
      ['app_title', DEFAULT_SETTINGS.app_title],
      ['sub_title', DEFAULT_SETTINGS.sub_title],
      ['bg_image', ''],
      ['logo_image', ''],
      ['drive_folder_id', '']
    ]);
    settingsSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e2e8f0');
    settingsSheet.autoResizeColumns(1, 2);
  }

  let questionsSheet = ss.getSheetByName('Questions');
  if (!questionsSheet) {
    questionsSheet = ss.insertSheet('Questions');
    const mockData = getFallbackMockData();
    const rows = [['id', 'type', 'question_text', 'question_image', 'content_json']];
    mockData.forEach(function (q) {
      rows.push([
        q.id,
        q.type,
        q.question_text,
        q.question_image || '',
        JSON.stringify(q.content)
      ]);
    });
    questionsSheet.getRange(1, 1, rows.length, 5).setValues(rows);
    questionsSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#e2e8f0');
    questionsSheet.setColumnWidth(3, 320);
    questionsSheet.setColumnWidth(5, 480);
  }

  SpreadsheetApp.getUi().alert('初始化完成！已確認 Settings 與 Questions 工作表皆就緒。');
}
