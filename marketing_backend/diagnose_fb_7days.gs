/**
 * 診斷最近 7 天 FB 貼文的互動數據
 * 直接讀 Historical_Posts、列出每一筆 FB 的 likes/comments/shares
 */
function diagnoseFb7Days() {
  const SS_ID = '1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA';
  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName('Historical_Posts');
  if (!sh || sh.getLastRow() < 2) { Logger.log('無資料'); return; }
  
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 7);
  
  const data = sh.getRange(2, 1, sh.getLastRow()-1, 18).getValues();
  Logger.log('===== 最近 7 天 FB 貼文 =====');
  Logger.log('cutoff: ' + cutoff.toISOString());
  
  let count = 0;
  let totalLikes = 0, totalComments = 0, totalShares = 0;
  const allFb = [];
  let rawFbCount = 0, dateParseFail = 0;
  const platformDist = {};
  
  Logger.log('Historical_Posts 總列數: ' + data.length);
  data.forEach(function(r, idx){
    const platform = String(r[1]||'');
    platformDist[platform] = (platformDist[platform]||0) + 1;
    if (platform !== 'FB') return;
    rawFbCount++;
    const ts = String(r[3]||'');
    if (rawFbCount <= 5) Logger.log('FB row ' + (idx+2) + ' 原始 ts=[' + ts + '] type=' + typeof r[3]);
    let dt;
    try {
      dt = new Date(ts.replace(' ','T') + ':00+08:00');
      if (isNaN(dt.getTime())) { dateParseFail++; return; }
    } catch(e) { dateParseFail++; return; }
    
    const likes = Number(r[8])||0;
    const comments = Number(r[9])||0;
    const shares = Number(r[10])||0;
    
    allFb.push({row: idx+2, ts: ts, likes: likes, comments: comments, shares: shares, dt: dt});
    
    if (dt < cutoff) return;
    count++;
    totalLikes += likes;
    totalComments += comments;
    totalShares += shares;
    Logger.log('row ' + (idx+2) + ' | ' + ts + ' | likes=' + likes + ' comments=' + comments + ' shares=' + shares + ' | id=' + r[2]);
  });
  
  Logger.log('===== Platform 分布 =====');
  Object.keys(platformDist).forEach(function(k){
    Logger.log('  ' + k + ': ' + platformDist[k] + ' 筆');
  });
  Logger.log('FB 原始筆數: ' + rawFbCount + '、日期解析失敗: ' + dateParseFail);
  Logger.log('===== 7 天總計 =====');
  Logger.log('FB 篇數: ' + count);
  Logger.log('總 likes: ' + totalLikes + '、總 comments: ' + totalComments + '、總 shares: ' + totalShares);
  Logger.log('總互動: ' + (totalLikes + totalComments + totalShares));
  Logger.log('平均互動: ' + (count ? ((totalLikes+totalComments+totalShares)/count).toFixed(1) : 0));
  
  // 排日期、看最新 10 筆 FB
  allFb.sort(function(a,b){return b.dt - a.dt;});
  Logger.log('===== 全部 FB 最新 10 筆 =====');
  allFb.slice(0,10).forEach(function(f){
    Logger.log('row ' + f.row + ' | ' + f.ts + ' | likes=' + f.likes + ' comments=' + f.comments + ' shares=' + f.shares);
  });
}
