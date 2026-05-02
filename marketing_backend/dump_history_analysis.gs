/**
 * 一次性數據分析輸出工具
 * 跑完看執行記錄、把整段內容傳給 AI 分析
 *
 * 使用：在 Apps Script 編輯器選 dumpHistoryAnalysis、按執行
 */

function dumpHistoryAnalysis() {
  const ssId = '1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA';
  const ss = SpreadsheetApp.openById(ssId);
  const sh = ss.getSheetByName('Historical_Posts');
  if (!sh || sh.getLastRow() < 2) {
    Logger.log('無資料');
    return;
  }
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 18).getValues();

  // 月度
  const monthly = {};
  // 星期幾
  const weekdayMap = [0,0,0,0,0,0,0];
  const weekdayCount = [0,0,0,0,0,0,0];
  // 類型
  const typeStats = {};
  // 平台
  const platStats = { IG: { posts: 0, eng: 0, reach: 0 }, FB: { posts: 0, eng: 0, reach: 0 } };
  // 全部、用於排序
  const all = [];

  data.forEach(function(r){
    const platform = String(r[1] || '');
    const ts = String(r[3] || '');
    const type = String(r[4] || '') || '其他';
    const permalink = String(r[5] || '');
    const captionShort = String(r[17] || '');
    const likes = Number(r[8]) || 0;
    const comments = Number(r[9]) || 0;
    const shares = Number(r[10]) || 0;
    const saved = Number(r[11]) || 0;
    const reach = Number(r[12]) || 0;
    const videoViews = Number(r[14]) || 0;
    const eng = likes + comments + shares + saved;

    // 月
    if (ts.length >= 7) {
      const ym = ts.substring(0, 7);
      if (!monthly[ym]) monthly[ym] = { posts: 0, eng: 0, reach: 0 };
      monthly[ym].posts++;
      monthly[ym].eng += eng;
      monthly[ym].reach += reach;
    }
    // 星期
    if (ts.length >= 10) {
      try {
        const d = new Date(ts.replace(' ', 'T') + ':00+08:00');
        if (!isNaN(d.getTime())) {
          const wd = (d.getDay() + 6) % 7; // 週一=0、週日=6
          weekdayMap[wd] += eng;
          weekdayCount[wd]++;
        }
      } catch (e) {}
    }
    // 類型
    if (!typeStats[type]) typeStats[type] = { posts: 0, eng: 0 };
    typeStats[type].posts++;
    typeStats[type].eng += eng;
    // 平台
    if (platStats[platform]) {
      platStats[platform].posts++;
      platStats[platform].eng += eng;
      platStats[platform].reach += reach;
    }
    // 全部
    all.push({ platform: platform, ts: ts, type: type, permalink: permalink, caption: captionShort, eng: eng, likes: likes, comments: comments, shares: shares, saved: saved, reach: reach, videoViews: videoViews });
  });

  Logger.log('========== 1. 平台對比 ==========');
  ['IG', 'FB'].forEach(function(p){
    const s = platStats[p];
    const avgEng = s.posts > 0 ? Math.round(s.eng / s.posts * 10) / 10 : 0;
    const avgReach = s.posts > 0 ? Math.round(s.reach / s.posts * 10) / 10 : 0;
    Logger.log(p + ': 貼文=' + s.posts + '、總互動=' + s.eng + '、單篇互動=' + avgEng + '、總觸及=' + s.reach + '、單篇觸及=' + avgReach);
  });

  Logger.log('');
  Logger.log('========== 2. 月度趨勢（最近 24 個月） ==========');
  const months = Object.keys(monthly).sort();
  const recent = months.slice(-24);
  recent.forEach(function(ym){
    const m = monthly[ym];
    const avg = Math.round(m.eng / m.posts * 10) / 10;
    Logger.log(ym + ' | 貼文=' + m.posts + '、總互動=' + m.eng + '、單篇互動=' + avg);
  });

  Logger.log('');
  Logger.log('========== 3. 星期幾發文表現 ==========');
  const wdNames = ['週一','週二','週三','週四','週五','週六','週日'];
  for (let i = 0; i < 7; i++) {
    const avg = weekdayCount[i] > 0 ? Math.round(weekdayMap[i] / weekdayCount[i] * 10) / 10 : 0;
    Logger.log(wdNames[i] + ': 貼文數=' + weekdayCount[i] + '、總互動=' + weekdayMap[i] + '、單篇平均互動=' + avg);
  }

  Logger.log('');
  Logger.log('========== 4. 內容類型表現 ==========');
  const types = Object.keys(typeStats).sort(function(a, b){ return typeStats[b].eng - typeStats[a].eng; });
  types.forEach(function(t){
    const s = typeStats[t];
    const avg = s.posts > 0 ? Math.round(s.eng / s.posts * 10) / 10 : 0;
    Logger.log(t + ': 貼文=' + s.posts + '、總互動=' + s.eng + '、單篇=' + avg);
  });

  Logger.log('');
  Logger.log('========== 5. Top 20 高互動貼文 ==========');
  all.sort(function(a, b){ return b.eng - a.eng; });
  all.slice(0, 20).forEach(function(p, i){
    Logger.log((i+1) + '. [' + p.platform + '/' + p.type + '] ' + p.ts + ' | 互動=' + p.eng + '（讚' + p.likes + '/留' + p.comments + '/分' + p.shares + '/存' + p.saved + '）| ' + p.caption);
  });

  Logger.log('');
  Logger.log('========== 6. 影片 vs 圖片 對比（IG） ==========');
  const igVideos = all.filter(function(p){ return p.platform === 'IG' && (p.type === 'VIDEO' || p.type === 'REELS'); });
  const igImages = all.filter(function(p){ return p.platform === 'IG' && (p.type === 'IMAGE' || p.type === 'CAROUSEL_ALBUM'); });
  if (igVideos.length > 0) {
    const sumE = igVideos.reduce(function(a, b){ return a + b.eng; }, 0);
    const sumV = igVideos.reduce(function(a, b){ return a + b.videoViews; }, 0);
    Logger.log('IG VIDEO/REELS: 數量=' + igVideos.length + '、平均互動=' + Math.round(sumE/igVideos.length*10)/10 + '、總觀看=' + sumV);
  }
  if (igImages.length > 0) {
    const sumE = igImages.reduce(function(a, b){ return a + b.eng; }, 0);
    Logger.log('IG IMAGE/CAROUSEL: 數量=' + igImages.length + '、平均互動=' + Math.round(sumE/igImages.length*10)/10);
  }

  Logger.log('');
  Logger.log('========== 7. 零互動貼文比例 ==========');
  const zero = all.filter(function(p){ return p.eng === 0; });
  Logger.log('零互動貼文=' + zero.length + ' / ' + all.length + '（' + Math.round(zero.length/all.length*1000)/10 + '%）');
  Logger.log('IG 零互動=' + zero.filter(function(p){return p.platform==='IG';}).length);
  Logger.log('FB 零互動=' + zero.filter(function(p){return p.platform==='FB';}).length);

  Logger.log('');
  Logger.log('完成、請把整段執行記錄複製給 AI 分析');
}
