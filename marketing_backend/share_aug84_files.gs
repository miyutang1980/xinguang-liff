/**
 * 把八月預告 84 個 Drive 檔批次設成「任何擁有連結者可檢視」
 *
 * 用途：dashboard.html 縮圖載不出來、因為 Drive 預設 Workspace domain 內部分享
 *      設成 ANYONE_WITH_LINK 後外部瀏覽器才能載 thumbnail
 *
 * 預期：84 檔 × 約 0.5 秒/檔 ≈ 42 秒
 *
 * 用法：
 *   1. 貼進 Apps Script
 *   2. 執行 shareAug84Files()
 *   3. 等 ~1 分鐘看到 ✅
 */

const AUG84_FILE_IDS = [
  '1xWXaqnjdqQrMgHrxSUF8Z3K-g0HqglPT',
  '17UCVB5WmAgtE2RapDEf5MjkzNITk9-C_',
  '1FFvLkMWqkSYA3xODLq1Eg6LgUnP2B01F',
  '1cnoca4z20XLjSJbj6WUzkFudcuCafXjR',
  '1tCNe01tS3NW4507TiluNPAIuMG9xY6yv',
  '1AXvY-Qpa07_J8bBs0Pf47aUG2cgAmzVc',
  '1zuGtPIm1xFy60sRaQvqDeXN2BofDaYSV',
  '1QOklXWahMKhMknoENNFpHCfx7oYi3YSM',
  '1JYPIqrfKW-rN8h8cS8ufos257uNokxFY',
  '1d_AK_Gr6H5kpCFYKc72ziC5ElnY_syTK',
  '1wMIOZNHPe8WDD45cOTydo8bVe_yPy0CK',
  '18mbFpr3z4SewHaUQ5WSeshOm_l89f40y',
  '1FCCKzHhr3Ag-tGUlwPe3NkQhAlfFBeFk',
  '1K7CDXqtybbnITkEL8X5-aY6ftM5fJqLo',
  '1VAd4fIkzDdE-bCtIJ60OdJYunvI2Eutk',
  '1DdsfKnuKfwwArWAytd6AC5Ms2jvWxnJX',
  '1njdIUIWGd3j31t2VMscj3gYGUvesM9S4',
  '1Q0EN7Cwt9B7PfZrsA_jJYQIOpjsOnYOF',
  '1LXfABawdSMIUsJvMhQdLDsvuVlvZzbPS',
  '1qRg6N5MPAVrwyGxiq1_61h-9rHmseGam',
  '1iyKdjBKf68FL1GnCnBbEIjRxAy6abDbn',
  '1ccp50R9DGxhyCO2uX2cXo6WjbYoAM9S2',
  '1NSjzjLLjrKn0IL-N3CjvUlYnz1qlnpv6',
  '1xlOoVKAK_jfRmTsl8GPBG_stwbcCPNdd',
  '1QGAaQN0KCVXWeZdUKhhrCs3MjAZMGAzg',
  '1H8KWuyHBlZmYV3lLr6xY9xQKqXrNfoh-',
  '107HN6qb2i8jqgSIQhE-9iFOZW5-KmHms',
  '1v6pPKXhF2ecOqMPCco4J9Beam1FxD_GV',
  '1xywSEUaioWf0MhF3L9vkXtyolbuUGBcQ',
  '1b4dCU0OtFxEttSeIMT3f9saSExzKaGjY',
  '1sglcOKRpfbGuR6zbt_cyVpZY7ei4b_bT',
  '10pHuH6n3VeGwk5zKmvVIFDlcmQ2scd7O',
  '1YQnXp4QzRzcb4rAOqQHFYvXwwtO0KpKg',
  '1LT6-T_WXhO5pl4piXf4ajDsB-8tmN9qo',
  '13NPc6Swzo3SRDgCsBe7yOuqPbLcVWI5l',
  '1DFm72fO06LNud-7ybNz-B6yfz5JNtFh0',
  '1l_lNKrGwC3VLMoWXWkcCjvL82_qrE17u',
  '1qx9gYFQ2urFwGOSPJB62d42kg1Mk5r0O',
  '1zjEDJftUtwQAeXlw70WlWhjpAtYaYU38',
  '1hK0ywMqhQBjnn9WihOWxycEBIZ-ChgMW',
  '1D-jNxHgQWUMu2n-FjDtKYKWRZpDVssvT',
  '19oHUwJcathpVOq2A-dSXS2jd6midPEPz',
  '1v6d7fzokG5kvHCre7VjN-cx0Kx-OSKNY',
  '1L8Q5OomCrk3ryqocBqL_oKJ3O7gTKWOt',
  '1gfSSCicuMmC_3WaRid-YtsogInBfdSAj',
  '1Q7zbRMw4T4HFWOxHU5rBhDbMNkuIDJjA',
  '1jDZg2nTCHseDfdfGR2BFz2soLXAiA28n',
  '16PiDkpL-ZlJbNdwoaX-M4bhBEIUphTP4',
  '1ksd_VzMcfv1MyKt3RytjtKLnQnZzbshx',
  '1Agaxd9NUwrT2OdyCHZS5R0d3JR696LUR',
  '1YUMn-e2NJXP-LggPt8Tns836X8A5e1kE',
  '1PvsNm7uQOe4EZVk7bFoSe-gu_gkFw5pH',
  '1CsL_Axe8z6kYbHNehzV9xGlCGRBFJkTm',
  '1YDG_4vgSp040HgimK03x7B_ZQTta9Apt',
  '1xMbBlSOSZIx9YaPzt-YxfNRSVYYW2wZK',
  '180Maf8ggLW85i1SCVLOsq8J1N_lBM94T',
  '1Y2iRXiEQ0A7dD1bSrs_LxK3tCuH49TIE',
  '17eyqK6tJJPXc6PIilZAlKZCDApRUxKuP',
  '11i4f0hM6f1HWD_rLw-gNtoAw6vQx-wNf',
  '1w8q4YtSVIQ30neib2jS9KPu41h5FggjL',
  '1Gl9pAOnEeXncRh1BNnzuT_x7WyAxqDj3',
  '1kz5co5VhOx8dYboCr3cBQRd-OIT_TEdB',
  '1m73Sfy2qBmolTSvF8OVBb20tA-XhePRn',
  '1_PVzuL9zb5M4SF9bMp5D0YHfCU_gokbn',
  '17Rvn_DpgA7HX5P5zVqvW1ASljIiU65ay',
  '103bziDUXZMD-mBtj19XURnqGUEPWXaQR',
  '1qhhYAJhAjQW41atbSlo82pGZeWPMeBdW',
  '1CiYnYOGsKwQayC4XrudVHs3k4tnoNCIo',
  '1nXz2lT4Jq-lQzcSPw6M_O3OfDla1GRAU',
  '1FGCCCRDTmmAkEzrx2Cmf8ajxMVITZjte',
  '1riIOFph8QlN_zdX5NpBv-yfNniaxCabq',
  '1S4lZ6xr7uzcyWGuyrmnKv2LdZskKezSE',
  '1TOCvMHEqXWW205X4844plN1Y5j_Uskh_',
  '1i6Z6qYOmHToF82fC_NdJrIPkdWg6ZHBY',
  '1v0amexwZ-1q4lX7wLWfxrFFpsSLctvVc',
  '1D-Op7O78tpzsVk3Hj0PAHhsgxYlkD5-C',
  '1oaAcle7BmNiz-k_Exx5pUAe3fzdc19ox',
  '1vKAPjuekamE2DKxbB8XFxVzn9sU1-Bd4',
  '1T1zjb3VFxg887MO-qCBtTI2tN1av0Wg3',
  '14uerbj5Mn5cvT4_8ZVEnltgnxlx9UA1F',
  '1_jI42CwUn409isQOQ6FDbht7tHszuZxO',
  '1Rc4CF9rkjhdVrEIhQ8Jq5mJV3l2RK9ds',
  '12UMFBl-AMYut24cxfRD9bfFTMzRqLH3e',
  '1ISnQ8zxIrK_KwZaPCbzr8iuxXWLxyGWf'
];

function shareAug84Files() {
  const ui = SpreadsheetApp.getUi();
  const ans = ui.alert(
    '⚠️ 批次設八月 84 檔權限\n\n' +
    '會把 84 個 Drive 檔設成「任何擁有連結者可檢視」。\n' +
    '這是必要的、否則 dashboard 縮圖看不到。\n' +
    '預估 1 分鐘完成。\n\n按 OK 確認',
    ui.ButtonSet.OK_CANCEL
  );
  if (ans !== ui.Button.OK) return;

  let ok = 0, fail = 0;
  const errors = [];
  AUG84_FILE_IDS.forEach((fid, i) => {
    try {
      const file = DriveApp.getFileById(fid);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      ok++;
      if (i % 10 === 0) Logger.log(`已處理 ${i + 1} / 84 ...`);
    } catch (e) {
      fail++;
      errors.push(fid + ': ' + e.message);
    }
  });

  Logger.log(`✅ 成功 ${ok}, ❌ 失敗 ${fail}`);
  errors.forEach(e => Logger.log(e));

  ui.alert(
    `批次完成\n\n成功 ${ok} / 失敗 ${fail}\n\n` +
    (fail > 0 ? '失敗清單看執行記錄' : '回 dashboard 強制重新整理 (Ctrl+Shift+R) 看八月圖')
  );
}

