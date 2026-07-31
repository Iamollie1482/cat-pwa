// ==========================================
// 貓咪健康日誌 - Google Apps Script 後端
// ==========================================

var SHEETS = {
  SHARED: '共用（排泄飲水）',
  SETTINGS: '設定'
};

function response(data) {
  var output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    var r = sheet.getRange(1, 1, 1, headers.length);
    r.setBackground('#F3EDE2');
    r.setFontWeight('bold');
    r.setFontSize(11);
  }
  return sheet;
}

function initSheets(ss, catNames) {
  for (var i = 0; i < catNames.length; i++) {
    getOrCreateSheet(ss, '🐱 ' + catNames[i], [
      '日期', '體重(kg)', '今日熱量(kcal)', '熱量目標(kcal)',
      '嘔吐次數', '剪指甲', '更換貓砂', '驅蟲', '刷牙', '疫苗',
      '藥物紀錄', '備註', '用餐明細'
    ]);
  }
  getOrCreateSheet(ss, SHEETS.SHARED, [
    '日期', '小便次數', '大便次數', '嘔吐（不明）', '總飲水量(ml)', '換水明細'
  ]);
  getOrCreateSheet(ss, SHEETS.SETTINGS, ['項目', '內容']);
}

// 找到某日期的列，找不到就新增
function findOrCreateRow(sheet, date) {
  var dateStr = String(date).trim().replace(/^'/, '');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var cell = data[i][0];
    var cellStr = (cell instanceof Date)
      ? Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(cell).trim().replace(/^'/, '');
    if (cellStr === dateStr) return i + 1;
  }
  sheet.appendRow(["'" + dateStr]);
  return sheet.getLastRow();
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── 初始化分頁 ──────────────────────────────
    if (action === 'init') {
      var catNames = body.catNames || [];
      initSheets(ss, catNames);
      var st = ss.getSheetByName(SHEETS.SETTINGS);
      var sd = st.getDataRange().getValues();
      var row = -1;
      for (var i = 1; i < sd.length; i++) {
        if (sd[i][0] === '貓咪列表') { row = i + 1; break; }
      }
      if (row === -1) st.appendRow(['貓咪列表', catNames.join(',')]);
      else st.getRange(row, 2).setValue(catNames.join(','));
      return response({ok: true, message: '初始化完成'});
    }

    // ── 共用分頁（排泄+飲水）只寫一次 ──────────
    if (action === 'syncShared') {
      var date = body.date;
      var shared = body.shared || {};
      var sh = ss.getSheetByName(SHEETS.SHARED);
      if (!sh) { initSheets(ss, []); sh = ss.getSheetByName(SHEETS.SHARED); }
      var shRow = findOrCreateRow(sh, date);
      var logs = shared.waterLogs || [];
      var total = logs.reduce(function(a, l) { return a + (Number(l.drunk)||0); }, 0);
      var detail = logs.map(function(l, i) {
        return '第'+(i+1)+'次:加'+(l.added||0)+'-剩'+(l.left||0)+'='+(l.drunk||0)+'ml';
      }).join(' | ');
      sh.getRange(shRow, 1, 1, 6).setValues([[
        "'"+date,
        shared.pee||0,
        shared.poop||0,
        shared.vomitUnknown||0,
        total,
        detail
      ]]);
      return response({ok: true});
    }

    // ── 單隻貓的分頁 ────────────────────────────
    if (action === 'syncCat') {
      var date = body.date;
      var catName = body.catName;
      var rec = body.rec || {};
      var meals = body.meals || [];
      var meds = body.meds || [];

      var cs = ss.getSheetByName('🐱 ' + catName);
      if (!cs) { initSheets(ss, [catName]); cs = ss.getSheetByName('🐱 ' + catName); }
      var cRow = findOrCreateRow(cs, date);

      var mealStr = meals.map(function(m) {
        return m.time+' '+m.name+' ×'+m.portions+'份('+Number(m.kcal||0).toFixed(1)+'kcal)';
      }).join(' | ');
      var medStr = meds.filter(function(m){return m.given;}).map(function(m){return m.name;}).join(', ');

      cs.getRange(cRow, 1, 1, 13).setValues([[
        "'"+date,
        rec.weight||'',
        rec.kcal||'',
        rec.dailyTarget||'',
        rec.vomit||0,
        rec.care&&rec.care.nail?'✓':'',
        rec.care&&rec.care.litter?'✓':'',
        rec.care&&rec.care.deworming?'✓':'',
        rec.care&&rec.care.brushteeth?'✓':'',
        rec.care&&rec.care.vaccine?'✓':'',
        medStr,
        rec.notes||'',
        mealStr
      ]]);
      return response({ok: true});
    }

    // ── 食物庫同步 ──────────────────────────────
    if (action === 'syncFoods') {
      var st4 = ss.getSheetByName(SHEETS.SETTINGS);
      if (!st4) { initSheets(ss,[]); st4 = ss.getSheetByName(SHEETS.SETTINGS); }
      var foodsStr = JSON.stringify(body.foods||[]);
      var sd4 = st4.getDataRange().getValues();
      var fRow = -1;
      for (var i = 1; i < sd4.length; i++) {
        if (sd4[i][0]==='食物庫') { fRow = i+1; break; }
      }
      if (fRow===-1) st4.appendRow(['食物庫', foodsStr]);
      else st4.getRange(fRow, 2).setValue(foodsStr);
      return response({ok: true, message: '食物庫已同步'});
    }

    // ── 批量匯入 ────────────────────────────────
    if (action === 'bulkSync') {
      var all = body.data;
      var cats = all.cats || [];
      var catNamesArr = cats.map(function(c){return c.name;});
      initSheets(ss, catNamesArr);

      // 收集所有日期
      var allDates = {};
      Object.keys(all.records||{}).forEach(function(k){ allDates[k.split('|')[1]]=true; });
      Object.keys(all.shared||{}).forEach(function(k){ allDates[k]=true; });
      Object.keys(all.meals||{}).forEach(function(k){ allDates[k.split('|')[1]]=true; });
      var dates = Object.keys(allDates).sort();

      function dailyTarget(cat) {
        return Math.round((30*(parseFloat(cat.weight)||0)+70)*(parseFloat(cat.status)||1.2)*10)/10;
      }

      for (var di = 0; di < dates.length; di++) {
        var dt = dates[di];

        // 共用頁
        var shd = all.shared[dt]||{};
        var wLogs = shd.waterLogs||[];
        var wTotal = wLogs.reduce(function(a,l){return a+(Number(l.drunk)||0);},0);
        var wStr = wLogs.map(function(l,i){
          return '第'+(i+1)+'次:加'+(l.added||0)+'-剩'+(l.left||0)+'='+(l.drunk||0)+'ml';
        }).join(' | ');
        var shSh = ss.getSheetByName(SHEETS.SHARED);
        shSh.getRange(findOrCreateRow(shSh,dt),1,1,6).setValues([[
          "'"+dt, shd.pee||0, shd.poop||0, shd.vomitUnknown||0, wTotal, wStr
        ]]);

        // 貓咪頁
        for (var ci = 0; ci < cats.length; ci++) {
          var cat = cats[ci];
          var k = cat.name+'|'+dt;
          var cr = all.records[k]||{};
          var catMeals = all.meals[k]||[];
          var foods = all.foods||[];
          var ateKcal = 0;
          var mealStr2 = catMeals.map(function(m){
            var f = null;
            for (var fi=0;fi<foods.length;fi++){if(foods[fi].id===m.foodId){f=foods[fi];break;}}
            var kc = f?f.kcal*m.portions:0; ateKcal+=kc;
            return m.time+' '+(f?f.name:'?')+' ×'+m.portions+'份('+kc.toFixed(1)+'kcal)';
          }).join(' | ');
          var catMedNames = (all.meds&&all.meds[cat.name])||[];
          var medGiven = catMedNames.filter(function(mn){return cr.meds&&cr.meds[mn];}).join(', ');
          var cSh = ss.getSheetByName('🐱 '+cat.name);
          cSh.getRange(findOrCreateRow(cSh,dt),1,1,13).setValues([[
            "'"+dt,
            cr.weight||'',
            ateKcal>0?Math.round(ateKcal*10)/10:'',
            dailyTarget(cat),
            cr.vomit||0,
            cr.care&&cr.care.nail?'✓':'',
            cr.care&&cr.care.litter?'✓':'',
            cr.care&&cr.care.deworming?'✓':'',
            cr.care&&cr.care.brushteeth?'✓':'',
            cr.care&&cr.care.vaccine?'✓':'',
            medGiven,
            cr.notes||'',
            mealStr2
          ]]);
        }
      }

      // 儲存食物庫
      var st2 = ss.getSheetByName(SHEETS.SETTINGS);
      var fStr2 = JSON.stringify(all.foods||[]);
      var sd2 = st2.getDataRange().getValues();
      var fRow2 = -1;
      for (var i=1;i<sd2.length;i++){if(sd2[i][0]==='食物庫'){fRow2=i+1;break;}}
      if (fRow2===-1) st2.appendRow(['食物庫',fStr2]);
      else st2.getRange(fRow2,2).setValue(fStr2);

      return response({ok:true, message:'匯入完成，共'+dates.length+'天資料'});
    }

    // ── 舊版相容（syncDay 直接忽略，用新的 syncShared+syncCat） ──
    if (action === 'syncDay') {
      return response({ok: true, message: 'deprecated'});
    }

    return response({ok: false, message: '未知的 action: ' + action});

  } catch(err) {
    return response({ok: false, message: err.toString()});
  }
}

function doGet(e) {
  return response({ok: true, message: '貓咪健康日誌 API 運作中 🐱'});
}
