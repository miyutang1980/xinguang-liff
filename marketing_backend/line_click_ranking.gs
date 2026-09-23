/**
 * Read-only aggregate endpoint for the existing admin LINE source analysis.
 * Uses DASH_SS_ID from marketing_dashboard.gs, not the school Gateway SHEET_ID.
 * Never creates sheets, records a click, or returns IP/referrer/user-agent rows.
 */
function lineClickRankingResponse_(e) {
  var p = (e && e.parameter) || {};
  var callback = String(p.callback || '');
  var callbackOK = !callback || /^[A-Za-z_$][0-9A-Za-z_$]{0,100}$/.test(callback);
  var result;
  if (!callbackOK) {
    return ContentService.createTextOutput(JSON.stringify({success:false,error:'Invalid callback'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var days = Number(p.days || 30);
  if ([7,30,90,365].indexOf(days) < 0) {
    result = {success:false,error:'日期範圍只支援 7、30、90 或 365 天'};
  } else {
    try {
      var cache = CacheService.getScriptCache();
      var now = new Date();
      var today = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd');
      var cacheKey = 'line_click_ranking_v1_' + today + '_' + days;
      var cached = cache.get(cacheKey);
      if (cached) { try { result = JSON.parse(cached); } catch(ignore) {} }
      if (!result) {
        var sh = SpreadsheetApp.openById(DASH_SS_ID).getSheetByName('LINE_Click_Log');
        if (!sh) throw new Error('Missing LINE_Click_Log');
        var headers = sh.getRange(1,1,1,2).getValues()[0];
        if (String(headers[0]).trim() !== '時間戳' || String(headers[1]).trim() !== '來源 src') {
          throw new Error('Unexpected LINE_Click_Log headers');
        }
        var last = sh.getLastRow();
        if (last > 50001) {
          result = {success:false,error:'點擊紀錄超過即時統計上限，請先建立彙總流程；目前未顯示不完整數據。'};
        } else {
          var rows = last >= 2 ? sh.getRange(2,1,last-1,2).getValues() : [];
          var midnight = new Date(today + 'T00:00:00+08:00').getTime();
          var start = midnight - (days - 1) * 86400000;
          var counts = Object.create(null), skipped = 0;
          rows.forEach(function(row) {
            var value = row[0], stamp = NaN;
            if (value instanceof Date) {
              stamp = value.getTime();
            } else {
              var text = String(value || '').trim();
              var parts = text.match(/^(\d{4})[/-](\d{2})[/-](\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?$/);
              if (parts) {
                stamp = new Date(parts[1]+'-'+parts[2]+'-'+parts[3]+'T'+(parts[4]||'00')+':'+(parts[5]||'00')+':'+(parts[6]||'00')+'+08:00').getTime();
              } else if (/^\d{4}-\d{2}-\d{2}T.+(?:Z|[+-]\d{2}:\d{2})$/.test(text)) {
                stamp = new Date(text).getTime();
              }
            }
            if (!Number.isFinite(stamp)) { skipped++; return; }
            if (stamp < start || stamp > now.getTime()) return;
            var source = String(row[1] || '').trim() || '未標示來源';
            counts[source] = (counts[source] || 0) + 1;
          });
          var data = Object.keys(counts).map(function(source) { return {src:source,clicks:counts[source]}; });
          data.sort(function(a,b) { return b.clicks-a.clicks || a.src.localeCompare(b.src); });
          result = {success:true,data:data,days:days,skippedRows:skipped,generatedAt:now.toISOString()};
          try { cache.put(cacheKey,JSON.stringify(result),60); } catch(ignoreCacheSize) {}
        }
      }
    } catch(error) {
      console.warn('lineClickRankingResponse_: ' + error.message);
      result = {success:false,error:'來源統計暫時無法讀取，請管理員核對試算表及「時間戳／來源 src」欄位；不要重建或清除紀錄。'};
    }
  }
  var json = JSON.stringify(result);
  return ContentService.createTextOutput(callback ? callback+'('+json+');' : json)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}
