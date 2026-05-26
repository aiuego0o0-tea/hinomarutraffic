/**
 * 日の丸交通 SNS採用管理スプレッドシート
 *
 * ★ 実行順序（1つずつ実行）★
 *   step1() → step2() → step3() → step4() → step5()
 *
 * step1 でスプレッドシートが作られIDが保存されます。
 * step2〜5 はそのIDを参照してシートを追加します。
 */

// ── カラー定数 ──
const C = {
  h1:'#1a237e', h2:'#283593', h3:'#3949ab', h4:'#5c6bc0',
  alt:'#e8eaf6', inp:'#fff9c4', gray:'#f5f5f5',
  green:'#c8e6c9', yellow:'#fff9c4', red:'#ffcdd2',
  calP:'#ede7f6', calS:'#e3f2fd', calE:'#fff3e0', calV:'#e8f5e9', calN:'#fffde7',
  pHi:'#ffcdd2', pMd:'#fff9c4', fw:'#ffffff', fg:'#212121', fm:'#757575'
};

function getSS() {
  const id = PropertiesService.getScriptProperties().getProperty('HINO_ID');
  if (!id) { Browser.msgBox('先に step1() を実行してください'); return null; }
  return SpreadsheetApp.openById(id);
}

// ══════════════════════════════════════════════════════
// STEP 1 — KPI管理シート（元Excelと同じ数式・構造）
// ══════════════════════════════════════════════════════
function step1() {
  const ss = SpreadsheetApp.create('日の丸交通 SNS採用管理 2026');
  PropertiesService.getScriptProperties().setProperty('HINO_ID', ss.getId());

  const s = ss.getSheets()[0];
  s.setName('📊 KPI管理');
  s.setTabColor('#1565c0');

  // 列幅
  s.setColumnWidth(1, 90);   // A: 指標名
  s.setColumnWidth(2, 75);   // B: ラベル
  s.setColumnWidth(3, 75);   // C: 上期計/サブラベル
  [4,5,6,7,8,9].forEach(c => s.setColumnWidth(c, 72));   // D〜I: 4〜9月
  s.setColumnWidth(10, 80);  // J: 上期合計
  [11,12,13,14,15,16].forEach(c => s.setColumnWidth(c, 72)); // K〜P: 10〜3月
  s.setColumnWidth(17, 80);  // Q: 下期合計
  s.setColumnWidth(18, 85);  // R: 年間合計

  // ── タイトル ──
  s.getRange('A1:R1').merge()
   .setValue('📊 日の丸交通 SNS採用管理 2026年度 KPIシート')
   .setFontSize(14).setFontWeight('bold').setFontColor(C.h1).setBackground('#e8eaf6');
  s.setRowHeight(1, 36);

  // ════════════════════════════════
  // 行2: 月ヘッダー
  // ════════════════════════════════
  s.getRange(2,1,1,18).setValues([[
    '指標','','上期計','4月','5月','6月','7月','8月','9月','上期計','10月','11月','12月','1月','2月','3月','下期計','年間合計'
  ]]).setBackground(C.h1).setFontColor(C.fw).setFontWeight('bold')
    .setHorizontalAlignment('center').setFontSize(10);
  s.setRowHeight(2, 30);

  // ════════════════════════════════
  // 行3-5: 月別【目標値】（固定値・要変更時は手入力）
  // 行3: LINE登録, 行4: LINE応募, 行5: 入社
  // ════════════════════════════════
  const targets = [
    ['LINE登録者','目標',  '', 0, 6, 7, 7, 10, 10, '=SUM(D3:I3)',  10,10,10,10,10,10, '=SUM(K3:P3)', '=SUM(J3,Q3)'],
    ['LINE応募者','目標',  '', 0, 0, 2, 3,  5,  5, '=SUM(D4:I4)',   2, 3, 3, 4, 4, 4, '=SUM(K4:P4)', '=SUM(J4,Q4)'],
    ['入社者',   '目標',  '', 0, 0, 0, 0,  1,  1, '=SUM(D5:I5)',   1, 1, 0, 0, 0, 1, '=SUM(K5:P5)', '=SUM(J5,Q5)'],
  ];
  targets.forEach((row, i) => {
    const r = 3 + i;
    s.getRange(r, 1, 1, 18).setValues([row]).setHorizontalAlignment('center');
    s.getRange(r, 1).setFontWeight('bold').setHorizontalAlignment('left');
    s.getRange(r, 2).setFontColor(C.fm);
    s.getRange(r, 10).setBackground(C.alt).setFontWeight('bold'); // 上期計
    s.getRange(r, 17).setBackground(C.alt).setFontWeight('bold'); // 下期計
    s.getRange(r, 18).setBackground(C.h4).setFontColor(C.fw).setFontWeight('bold'); // 年間
    s.getRange(r, 10).setNumberFormat('#,##0');
    s.getRange(r, 17).setNumberFormat('#,##0');
    s.getRange(r, 18).setNumberFormat('#,##0');
  });

  // ════════════════════════════════
  // 行7: 達成率サマリーラベル
  // ════════════════════════════════
  s.getRange('A7:R7').merge().setValue('◆ 達成率サマリー（実績÷目標）— 実績入力後に自動更新')
   .setBackground(C.h2).setFontColor(C.fw).setFontWeight('bold').setFontSize(10);
  s.setRowHeight(7, 26);

  // ── 行8: ヘッダー ──
  s.getRange(8,1,1,9).setValues([['指標','','上期実績','上期目標','上期達成率','下期実績','下期目標','下期達成率','年間達成率']])
   .setBackground(C.h3).setFontColor(C.fw).setFontWeight('bold').setHorizontalAlignment('center').setFontSize(9);

  // ── 行9-11: 各指標の達成率 ──
  // LINE登録(行9)・LINE応募(行10)・入社(行11)
  // 参照先: LINE登録実績=行15, LINE応募実績=行20, 入社実績=行25
  const summaryDefs = [
    { label:'LINE登録者', actRow:15, tgtRow:3 },
    { label:'LINE応募者', actRow:20, tgtRow:4 },
    { label:'入社者',     actRow:25, tgtRow:5 },
  ];
  summaryDefs.forEach((d, i) => {
    const r = 9 + i;
    const ar = d.actRow, tr = d.tgtRow;
    s.getRange(r, 1, 1, 9).setValues([[
      d.label, '',
      `=IF(I${ar}="","集計中",SUM(D${ar}:I${ar}))`,  // 上期実績
      `=J${tr}`,                                        // 上期目標
      `=IF(C${r}="集計中","",C${r}/D${r})`,            // 上期達成率
      `=IF(P${ar}="","集計中",SUM(K${ar}:P${ar}))`,   // 下期実績
      `=Q${tr}`,                                        // 下期目標
      `=IF(F${r}="集計中","",F${r}/G${r})`,            // 下期達成率
      `=IF(C${r}="集計中","",（C${r}+IF(F${r}="集計中",0,F${r}))/R${tr})`, // 年間達成率
    ]]).setHorizontalAlignment('center').setFontSize(10);
    s.getRange(r, 1).setFontWeight('bold').setHorizontalAlignment('left');
    // 達成率列をパーセント書式に
    s.getRange(r, 5).setNumberFormat('0.0%');
    s.getRange(r, 8).setNumberFormat('0.0%');
    s.getRange(r, 9).setNumberFormat('0.0%');
    // 実績数値
    s.getRange(r, 3).setNumberFormat('#,##0').setBackground('#e3f2fd');
    s.getRange(r, 6).setNumberFormat('#,##0').setBackground('#e3f2fd');
    s.setRowHeight(r, 28);
  });

  // ════════════════════════════════
  // 行13: 月別詳細ラベル
  // ════════════════════════════════
  s.getRange('A13:R13').merge().setValue('◆ 月別詳細（実績を入力するとすべて自動計算）— 薄黄色セルに実績を入力')
   .setBackground(C.h2).setFontColor(C.fw).setFontWeight('bold').setFontSize(10);
  s.setRowHeight(13, 26);

  // 月ヘッダー（行14）
  s.getRange(14,1,1,18).setValues([[
    '指標','項目','',  '4月','5月','6月','7月','8月','9月','上期計','10月','11月','12月','1月','2月','3月','下期計','年間累計'
  ]]).setBackground(C.h3).setFontColor(C.fw).setFontWeight('bold')
    .setHorizontalAlignment('center').setFontSize(9);
  s.setRowHeight(14, 28);

  // ════════════════════════════════
  // 月別詳細ブロック（3指標 × 4行）
  // ████ ここが元Excelと同じ数式 ████
  // LINE登録: 行15-18 / LINE応募: 行20-23 / 入社: 行25-28
  // ════════════════════════════════
  function writeDetailBlock(startRow, label, tgtRow, color) {
    const sr = startRow;
    const months = ['D','E','F','G','H','I','K','L','M','N','O','P'];

    // ── 行sr: 実績入力行（薄黄色）──
    s.getRange(sr, 1).setValue(label).setFontWeight('bold').setBackground(color).setFontSize(10);
    s.getRange(sr, 2).setValue('実績入力').setFontColor(C.fm).setFontSize(9);
    // 実績入力セル（D〜I, K〜P）を薄黄色に
    ['D','E','F','G','H','I'].forEach(col => s.getRange(`${col}${sr}`).setBackground(C.inp));
    ['K','L','M','N','O','P'].forEach(col => s.getRange(`${col}${sr}`).setBackground(C.inp));
    // 上期計・下期計・年間累計（自動）
    s.getRange(`J${sr}`).setFormula(`=IF(I${sr}="","",SUM(D${sr}:I${sr}))`).setBackground(C.alt).setFontWeight('bold');
    s.getRange(`Q${sr}`).setFormula(`=IF(P${sr}="","",SUM(K${sr}:P${sr}))`).setBackground(C.alt).setFontWeight('bold');
    s.getRange(`R${sr}`).setFormula(`=SUM(J${sr},IF(Q${sr}="",0,Q${sr}))`).setBackground(C.h4).setFontColor(C.fw).setFontWeight('bold');

    // ── 行sr+1: 達成率(%) ──
    s.getRange(sr+1, 2).setValue('達成率(%)').setFontColor(C.fm).setFontSize(9);
    // 4〜9月: 実績/目標
    ['D','E','F','G','H','I'].forEach(col => {
      s.getRange(`${col}${sr+1}`).setFormula(`=IF(${col}${sr}="","",${col}${sr}/IF(${col}${tgtRow}=0,1,${col}${tgtRow}))`).setNumberFormat('0%');
    });
    // 上期計達成率
    s.getRange(`J${sr+1}`).setFormula(`=IF(J${sr}="","",J${sr}/J${tgtRow})`).setNumberFormat('0%').setBackground(C.alt).setFontWeight('bold');
    // 10〜3月
    ['K','L','M','N','O','P'].forEach(col => {
      s.getRange(`${col}${sr+1}`).setFormula(`=IF(${col}${sr}="","",${col}${sr}/IF(${col}${tgtRow}=0,1,${col}${tgtRow}))`).setNumberFormat('0%');
    });
    s.getRange(`Q${sr+1}`).setFormula(`=IF(Q${sr}="","",Q${sr}/Q${tgtRow})`).setNumberFormat('0%').setBackground(C.alt).setFontWeight('bold');
    s.getRange(`R${sr+1}`).setFormula(`=IF(R${sr}=0,"",R${sr}/R${tgtRow})`).setNumberFormat('0%').setBackground(C.h4).setFontColor(C.fw).setFontWeight('bold');

    // ── 行sr+2: 純増数（前月比）──
    s.getRange(sr+2, 2).setValue('純増数').setFontColor(C.fm).setFontSize(9);
    s.getRange(`D${sr+2}`).setFormula(`=IF(D${sr}="","",D${sr})`); // 4月は今月のみ
    s.getRange(`E${sr+2}`).setFormula(`=IF(E${sr}="","",E${sr}-D${sr})`);
    s.getRange(`F${sr+2}`).setFormula(`=IF(F${sr}="","",F${sr}-E${sr})`);
    s.getRange(`G${sr+2}`).setFormula(`=IF(G${sr}="","",G${sr}-F${sr})`);
    s.getRange(`H${sr+2}`).setFormula(`=IF(H${sr}="","",H${sr}-G${sr})`);
    s.getRange(`I${sr+2}`).setFormula(`=IF(I${sr}="","",I${sr}-H${sr})`);
    s.getRange(`J${sr+2}`).setValue('—').setHorizontalAlignment('center').setBackground(C.alt);
    s.getRange(`K${sr+2}`).setFormula(`=IF(K${sr}="","",K${sr}-I${sr})`); // 10月は9月から引く
    s.getRange(`L${sr+2}`).setFormula(`=IF(L${sr}="","",L${sr}-K${sr})`);
    s.getRange(`M${sr+2}`).setFormula(`=IF(M${sr}="","",M${sr}-L${sr})`);
    s.getRange(`N${sr+2}`).setFormula(`=IF(N${sr}="","",N${sr}-M${sr})`);
    s.getRange(`O${sr+2}`).setFormula(`=IF(O${sr}="","",O${sr}-N${sr})`);
    s.getRange(`P${sr+2}`).setFormula(`=IF(P${sr}="","",P${sr}-O${sr})`);
    s.getRange(`Q${sr+2}`).setValue('—').setHorizontalAlignment('center').setBackground(C.alt);
    s.getRange(`R${sr+2}`).setValue('—').setHorizontalAlignment('center').setBackground(C.h4).setFontColor(C.fw);

    // 純増数の条件付き書式（正=緑、負=赤）
    const incrRange = s.getRange(`D${sr+2}:I${sr+2}`);
    const incrRange2 = s.getRange(`K${sr+2}:P${sr+2}`);

    // ── 行sr+3: 累計 ──
    s.getRange(sr+3, 2).setValue('累計').setFontColor(C.fm).setFontSize(9);
    s.getRange(`D${sr+3}`).setFormula(`=IF(D${sr}="","",SUM(D${sr}:D${sr}))`);
    s.getRange(`E${sr+3}`).setFormula(`=IF(E${sr}="","",SUM(D${sr}:E${sr}))`);
    s.getRange(`F${sr+3}`).setFormula(`=IF(F${sr}="","",SUM(D${sr}:F${sr}))`);
    s.getRange(`G${sr+3}`).setFormula(`=IF(G${sr}="","",SUM(D${sr}:G${sr}))`);
    s.getRange(`H${sr+3}`).setFormula(`=IF(H${sr}="","",SUM(D${sr}:H${sr}))`);
    s.getRange(`I${sr+3}`).setFormula(`=IF(I${sr}="","",SUM(D${sr}:I${sr}))`);
    s.getRange(`J${sr+3}`).setFormula(`=IF(J${sr}="","",J${sr})`).setBackground(C.alt).setFontWeight('bold');
    s.getRange(`K${sr+3}`).setFormula(`=IF(K${sr}="","",SUM(K${sr}:K${sr}))`);
    s.getRange(`L${sr+3}`).setFormula(`=IF(L${sr}="","",SUM(K${sr}:L${sr}))`);
    s.getRange(`M${sr+3}`).setFormula(`=IF(M${sr}="","",SUM(K${sr}:M${sr}))`);
    s.getRange(`N${sr+3}`).setFormula(`=IF(N${sr}="","",SUM(K${sr}:N${sr}))`);
    s.getRange(`O${sr+3}`).setFormula(`=IF(O${sr}="","",SUM(K${sr}:O${sr}))`);
    s.getRange(`P${sr+3}`).setFormula(`=IF(P${sr}="","",SUM(K${sr}:P${sr}))`);
    s.getRange(`Q${sr+3}`).setFormula(`=IF(Q${sr}="","",Q${sr})`).setBackground(C.alt).setFontWeight('bold');
    s.getRange(`R${sr+3}`).setFormula(`=IF(R${sr}=0,"",R${sr})`).setBackground(C.h4).setFontColor(C.fw).setFontWeight('bold');

    // 罫線（ブロック全体）
    s.getRange(sr, 1, 4, 18).setBorder(true, true, true, true, false, true, '#c5cae9', SpreadsheetApp.BorderStyle.SOLID);
    // 行の高さ
    for (let i = 0; i < 4; i++) s.setRowHeight(sr + i, 26);
  }

  // LINE登録者（行15-18）、背景色: 薄青
  writeDetailBlock(15, 'LINE登録者', 3, '#e3f2fd');
  // 空行（行19）
  s.getRange('A19:R19').setBackground('#fafafa'); s.setRowHeight(19, 8);
  // LINE応募者（行20-23）、背景色: 薄緑
  writeDetailBlock(20, 'LINE応募者', 4, '#e8f5e9');
  // 空行（行24）
  s.getRange('A24:R24').setBackground('#fafafa'); s.setRowHeight(24, 8);
  // 入社者（行25-28）、背景色: 薄紫
  writeDetailBlock(25, '入社者', 5, '#f3e5f5');

  // ════════════════════════════════
  // 行30: 給与シミュレーター
  // ════════════════════════════════
  s.getRange('A30:R30').merge().setValue('◆ タクシー給与シミュレーター（参考値）')
   .setBackground(C.h2).setFontColor(C.fw).setFontWeight('bold').setFontSize(10);
  s.setRowHeight(30, 26);

  s.getRange(31,1,1,6).setValues([['売上（歩合計算元）','','','手取り計算','','']])
   .setBackground(C.h3).setFontColor(C.fw).setFontWeight('bold').setHorizontalAlignment('center');

  s.getRange(32,1,1,6).setValues([['月売上（円）','歩合給','','固定月額','手取り（歩合）','手取り（固定）']])
   .setBackground(C.alt).setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');

  // 入力セル
  s.getRange('A33').setValue(900000).setBackground(C.inp).setNumberFormat('#,##0');
  s.getRange('D33').setValue(320000).setBackground(C.inp).setNumberFormat('#,##0');

  // 歩合給計算（元Excelの数式そのまま）
  s.getRange('B33').setFormula(
    '=IFS(A33<=203300,A33*50.825%,' +
    'A33<=266800,A33*53.36%,' +
    'A33<=368508,A33*56.6935%,' +
    'A33<=471783,A33*58.9729%,' +
    'A33<=565908,A33*59.5693%,' +
    'A33<=660708,A33*60.0644%,' +
    'A33<=755808,A33*60.4646%,' +
    'A33<=953583,A33*61.0309%,' +
    'A33<=1550000,A33*61.65%,' +
    'TRUE,A33*62%)'
  ).setNumberFormat('#,##0').setBackground('#fffde7');

  // 手取り計算（元Excelの数式そのまま）
  s.getRange('E33').setFormula(
    '=IFS(B33<200000,B33*0.85,' +
    'B33<300000,B33*0.82,' +
    'B33<500000,B33*0.78,' +
    'B33<700000,B33*0.75,' +
    'B33<1000000,B33*0.72,' +
    'TRUE,B33*0.70)'
  ).setNumberFormat('#,##0').setBackground('#e8f5e9');

  s.getRange('F33').setFormula(
    '=IFS(D33<200000,D33*0.85,' +
    'D33<300000,D33*0.82,' +
    'D33<500000,D33*0.78,' +
    'D33<700000,D33*0.75,' +
    'D33<1000000,D33*0.72,' +
    'TRUE,D33*0.70)'
  ).setNumberFormat('#,##0').setBackground('#e8f5e9');

  s.getRange('A34:F34').setValues([['↑ 売上を変えると自動計算','','','↑ 固定給を変えると自動計算','','']])
   .setFontSize(9).setFontColor(C.fm).setHorizontalAlignment('center');

  // ── 達成率サマリーの数式を修正（年間） ──
  // 行9: LINE登録
  s.getRange('I9').setFormula('=IF(C9="集計中","",（C9+IF(F9="集計中",0,F9))/R3)').setNumberFormat('0.0%');
  // 行10: LINE応募
  s.getRange('I10').setFormula('=IF(C10="集計中","",（C10+IF(F10="集計中",0,F10))/R4)').setNumberFormat('0.0%');
  // 行11: 入社
  s.getRange('I11').setFormula('=IF(C11="集計中","",（C11+IF(F11="集計中",0,F11))/R5)').setNumberFormat('0.0%');

  // ── ヘッダー行固定 ──
  s.setFrozenRows(2);

  SpreadsheetApp.flush();
  Browser.msgBox(
    '✅ Step1完了！KPI管理シートが作成されました。\n\n' +
    '薄黄色のセルに実績を入力すると全て自動計算されます。\n\n' +
    '次に step2() を実行してカレンダーシートを追加してください。\n\n' +
    'URL: ' + ss.getUrl()
  );
}

// ══════════════════════════════════════════════════════
// STEP 2 — 投稿カレンダー（5月）
// ══════════════════════════════════════════════════════
function step2() {
  const ss = getSS(); if (!ss) return;
  const s = ss.insertSheet('🗓 投稿カレンダー');
  s.setTabColor('#e65100');

  s.setColumnWidth(1, 55);
  for (let c = 2; c <= 8; c++) s.setColumnWidth(c, 148);

  s.getRange('A1:H1').merge()
   .setValue('🗓 2026年5月 投稿カレンダー — TikTok基軸（同動画をYouTube/Instagramにも展開 ｜ Xは別管理）')
   .setFontSize(11).setFontWeight('bold').setFontColor(C.h1);

  // 凡例
  const lgd = [['企画',C.calP],['撮影',C.calS],['編集締め切り',C.calE],['動画投稿(TikTok/YT/IG)',C.calV],['note(毎週月曜)',C.calN]];
  lgd.forEach((l,i) => s.getRange(2,i+2).setValue(l[0]).setBackground(l[1]).setFontSize(9).setHorizontalAlignment('center').setFontWeight('bold'));
  s.getRange('A2').setValue('凡例').setFontSize(9).setFontColor(C.fm);

  s.getRange(4,1,1,8).setValues([['','月','火','水','木','金','土','日']])
   .setBackground(C.h1).setFontColor(C.fw).setFontWeight('bold').setHorizontalAlignment('center').setFontSize(11);

  const weeks = [
    { lbl:'第1週', d:['4/27','4/28','4/29','4/30','5/1','5/2','5/3'],
      c:['','','','','企画\n動画①「乗務員の1日密着」テーマ・構成決め','撮影準備\n機材・撮影場所の確認',''],
      bg:['#fff','#fff','#fff','#fff',C.calP,C.calP,'#fff'] },
    { lbl:'第2週', d:['5/4','5/5','5/6','5/7','5/8','5/9','5/10'],
      c:['note\n「タクシー業界の豆知識5選」','企画\n動画②「未経験転職のリアル」構成決め','撮影\n動画①（午前の部・密着収録）','撮影\n動画①（補足カット）','編集締め切り\n動画①テロップ・SEタチ確認','投稿🎬\nTikTok/YouTube/IG\n「乗務員の1日密着」',''],
      bg:[C.calN,C.calP,C.calS,C.calS,C.calE,C.calV,'#fff'] },
    { lbl:'第3週', d:['5/11','5/12','5/13','5/14','5/15','5/16','5/17'],
      c:['note\n「転職Q&A: タクシー転職してよかったこと」','企画\n動画③「福利厚生を全部紹介」','撮影\n動画②（転職社員インタビュー）','撮影\n動画②（社内カット補足）','編集締め切り\n動画②完成・確認','投稿🎬\nTikTok/YouTube/IG\n「未経験転職の正直な話」',''],
      bg:[C.calN,C.calP,C.calS,C.calS,C.calE,C.calV,'#fff'] },
    { lbl:'第4週', d:['5/18','5/19','5/20','5/21','5/22','5/23','5/24'],
      c:['note\n「4月SNS実績まとめ・振り返り」','企画\n動画④「タクシーあるある10選」','撮影\n動画③（社内・福利厚生収録）','撮影\n動画③（補足カット）','編集締め切り\n動画③完成・確認','投稿🎬\nTikTok/YouTube/IG\n「福利厚生を全部紹介」',''],
      bg:[C.calN,C.calP,C.calS,C.calS,C.calE,C.calV,'#fff'] },
    { lbl:'第5週', d:['5/25','5/26','5/27','5/28','5/29','5/30','5/31'],
      c:['note\n「現場の声: 転職して変わったこと」','企画確定\n動画④台本・絵コンテ作成','撮影\n動画④（複数人収録）','撮影補足\n動画④','編集締め切り\n動画④完成・確認','投稿🎬\nTikTok/YouTube/IG\n「タクシーあるある10選」','月末まとめ\n5月数値集計\n6月計画立案'],
      bg:[C.calN,C.calP,C.calS,C.calS,C.calE,C.calV,C.alt] },
  ];

  let r = 5;
  weeks.forEach(w => {
    s.getRange(r,1,2,1).merge().setValue(w.lbl)
     .setBackground(C.h3).setFontColor(C.fw).setFontWeight('bold')
     .setHorizontalAlignment('center').setVerticalAlignment('middle');
    s.getRange(r,2,1,7).setValues([w.d]).setFontSize(9).setFontColor(C.fm).setHorizontalAlignment('right').setBackground('#fafafa');
    s.setRowHeight(r, 18); s.setRowHeight(r+1, 88);
    for (let d=0;d<7;d++) s.getRange(r+1,d+2).setValue(w.c[d]).setBackground(w.bg[d]).setFontSize(9).setWrap(true).setVerticalAlignment('top');
    r += 2;
    s.getRange(r,1,1,8).setBackground('#e0e0e0'); s.setRowHeight(r,4); r++;
  });

  s.setFrozenRows(4);
  SpreadsheetApp.flush();
  Browser.msgBox('✅ Step2完了！次に step3() を実行してください。');
}

// ══════════════════════════════════════════════════════
// STEP 3 — PDCA（目的・ペルソナ・カスタマージャーニー）
// ══════════════════════════════════════════════════════
function step3() {
  const ss = getSS(); if (!ss) return;
  const s = ss.insertSheet('📋 2026年度PDCA');
  s.setTabColor('#6a1b9a');

  s.setColumnWidth(1,130); s.setColumnWidth(2,220);
  for(let c=3;c<=7;c++) s.setColumnWidth(c,170);

  s.getRange('A1:G1').merge().setValue('📋 2026年度PDCA — 運用ルールブック（目的・KGI・ペルソナ・カスタマージャーニー）')
   .setFontSize(12).setFontWeight('bold').setFontColor(C.h1);

  let r = 3;
  function sec(txt,col){s.getRange(r,1,1,7).merge().setValue(txt).setBackground(col||C.h1).setFontColor(C.fw).setFontWeight('bold').setFontSize(11);s.setRowHeight(r,28);r++;}
  function kv(l,v){s.getRange(r,1).setValue(l).setFontWeight('bold').setBackground(C.alt).setWrap(true).setVerticalAlignment('top');s.getRange(r,2,1,6).merge().setValue(v).setWrap(true).setVerticalAlignment('top');s.setRowHeight(r,60);r++;}
  function th(h){s.getRange(r,1,1,h.length).setValues([h]).setBackground(C.h3).setFontColor(C.fw).setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center').setWrap(true);s.setRowHeight(r,32);r++;}
  function tr(v,bg){s.getRange(r,1,1,v.length).setValues([v]).setWrap(true).setVerticalAlignment('top');if(bg)s.getRange(r,1,1,v.length).setBackground(bg);s.setRowHeight(r,55);r++;}

  sec('① 目的・KGI・KPI','#1a237e');
  kv('目的','紹介会社からの流入を減らし、自社SNSから直接応募を増やすことで採用コストを削減する。\n採用1名あたり紹介会社経由70〜100万円 → SNS自社採用へ切り替えることが最終目標。');
  kv('KGI','公式LINEを通じた自社応募を年間20名以上とし、採用3名以上を自社経由で実現する。');
  kv('KPI','【最優先】公式LINE登録者数 → 年間100名（月平均8〜10名）\n応募数: 年間20名 ｜ 入社数: 年間3名\nフォロワー数・視聴数・クリック数は参考指標（週次集計）');
  kv('SNSファネル','SNS投稿（認知）→ プロフィールLINE誘導 → LINE登録★ → 職場体験会/説明会 → 面接 → 採用 → SNS発信（紹介）');
  r++;

  sec('② ターゲットセグメント','#283593');
  th(['セグメント','年齢','現職','転職動機','主な不安','SNS行動','優先度']);
  [['S1 主力','30〜35歳','営業・サービス業','残業多・体力の限界・腰痛','未経験で採用されるか\n収入が下がらないか','TikTok（通勤）\nInstagram','🔴 最高'],
   ['S2 主力','36〜45歳','製造業・工場','自動化で雇用不安\n体力仕事から脱出したい','家族を養えるか\n転職後の安定性','YouTube（夜）\nTikTok（休憩中）','🔴 最高'],
   ['S3 重要','46〜55歳','様々','定年なく働きたい\nマイペースで稼ぎたい','体力が続くか\n夜勤できるか','YouTube（休日）','🟡 高'],
   ['S4 将来','55歳以上','シニア・再雇用中','再雇用の期限・つながり','若い職場に馴染めるか','YouTube','🟡 高'],
   ['S5 育成','20〜29歳','大学生・高卒','稼ぎたい・自由・二種免許','タクシーのイメージ','TikTok','🟢 中'],
  ].forEach((row,i)=>tr(row,i%2===0?'#fff':'#f5f5f5'));
  r++;

  sec('③ ペルソナ（人物像）','#1b5e20');
  s.getRange(r,1,1,7).merge().setValue('ペルソナ①「田中 健太さん」（33歳 ／ S1セグメント）').setBackground('#e8f5e9').setFontWeight('bold').setFontSize(10);s.setRowHeight(r,24);r++;
  [['居住/職業','東京都江戸川区 ／ 食品メーカー ルート営業4年目 ／ 月収26万（手取り） ／ 妻・子1人（2歳）'],
   ['悩み','腰痛がひどくなった。残業代カット。子どもの成長を見られていない。21〜22時帰宅が続く。'],
   ['転職きっかけ','飲み会で先輩が「タクシー転職して月収が上がった」と一言。TikTokで「タクシー転職リアル」が流れてきた。'],
   ['不安','「未経験でも採用されるか」「二種免許の費用は？」「妻を説得できるか」'],
   ['刺さるコンテンツ','「ルート営業→タクシー転職した話」「リアル月収公開」「二種免許の流れ」「妻への説得方法」'],
  ].forEach(kd=>{s.getRange(r,1).setValue(kd[0]).setFontWeight('bold').setBackground('#f1f8e9').setWrap(true);s.getRange(r,2,1,6).merge().setValue(kd[1]).setWrap(true);s.setRowHeight(r,44);r++;});
  r++;
  s.getRange(r,1,1,7).merge().setValue('ペルソナ②「鈴木 雄介さん」（40歳 ／ S2セグメント）').setBackground('#e3f2fd').setFontWeight('bold').setFontSize(10);s.setRowHeight(r,24);r++;
  [['居住/職業','埼玉県川口市 ／ 自動車部品製造工場 オペレーター12年目 ／ 月収30万（手取り） ／ 妻・子2人（中1・小4）'],
   ['悩み','工場にロボットが導入され仕事量が減少。45歳でリストラが現実的になってきた。'],
   ['転職きっかけ','YouTubeで「製造業からタクシーへ転職した40代の話」を発見。月収が増えたという内容に興味。'],
   ['不安','「子どもの学費（大学）を稼げるか」「タクシー業界の将来性は？」「体力より知識が活かせるか」'],
   ['刺さるコンテンツ','「製造業からの転職体験談」「40代の月収シミュレーション」「家族持ちの福利厚生」「定年のない働き方」'],
  ].forEach(kd=>{s.getRange(r,1).setValue(kd[0]).setFontWeight('bold').setBackground('#e3f2fd').setWrap(true);s.getRange(r,2,1,6).merge().setValue(kd[1]).setWrap(true);s.setRowHeight(r,44);r++;});
  r++;

  sec('④ カスタマージャーニー（認知 → 入社 → 紹介）','#4a148c');
  th(['ステージ','状態・心理','タッチポイント','コンテンツ/施策','CTA']);
  [['①認知','タクシー転職をまだ考えていない。\nSNSで偶然流れてきた。','TikTok/YouTube/IG','リアル年収公開/あるある\n1日密着/転職体験談','フォロー・保存'],
   ['②興味','「悪くなさそう」\nプロフィールを見に来た。','プロフィール・固定投稿','会社紹介/社員インタビュー\n福利厚生紹介','他の動画も見る・フォロー'],
   ['③情報収集','「本当に稼げるか？できるか？」\n長尺動画を検索・視聴中。','YouTube長尺/note/IG','転職体験談/Q&A\n二種免許の流れ/月収シミュ','LINE登録して詳細確認'],
   ['④比較検討','他社と比較・家族に相談中。\n「背中を押してほしい」状態。','LINE公式\n（ステップ配信）','登録直後: 会社の強み\n3日後: よくある質問\n7日後: 体験会案内','職場体験会に申し込む'],
   ['⑤体験会','「リアルに見てみたい」\n現場見学・先輩と話す。','職場体験会\n会社説明会','現場見学・先輩との懇談\n疑問を全部解消する場','応募書類を提出する'],
   ['⑥応募・面接','「ここに入りたい」','面接・採用担当','二種免許サポート詳細説明\n具体的な働き方の話','内定承諾'],
   ['⑦入社','「転職して良かった」\n新しい仕事に慣れている。','LINE・社内OJT','新人サポート・定着フォロー','SNS発信に協力してもらう'],
   ['⑧紹介・拡散','「友人にも教えたい」\nファンになった状態。','TikTok・X・口コミ','転職体験談動画の出演依頼\n社員紹介制度の活用','採用コスト削減ループ完成✅'],
  ].forEach((row,i)=>tr(row,i%2===0?'#fff':'#f3e5f5'));

  s.setFrozenRows(1);
  SpreadsheetApp.flush();
  Browser.msgBox('✅ Step3完了！次に step4() を実行してください。');
}

// ══════════════════════════════════════════════════════
// STEP 4 — 企画立案書
// ══════════════════════════════════════════════════════
function step4() {
  const ss = getSS(); if (!ss) return;
  const s = ss.insertSheet('💡 企画立案書');
  s.setTabColor('#f9a825');

  [40,85,95,185,185,185,95,70,120,60,70,75,145].forEach((w,i)=>s.setColumnWidth(i+1,w));

  s.getRange('A1:M1').merge().setValue('💡 企画立案書 — 動画・noteコンテンツ管理台帳')
   .setFontSize(12).setFontWeight('bold').setFontColor(C.h1);
  s.getRange('A2:M2').merge()
   .setValue('ステータス: 💡企画中 → 📸撮影待ち → ✂️編集中 → ✅投稿済み ｜ ❌却下の場合は備考に理由を記入')
   .setFontSize(9).setFontColor(C.fm).setBackground('#f5f5f5');
  s.getRange('A3:M3').merge()
   .setValue('カテゴリ: タクシーあるある ／ 豆知識 ／ リアル年収 ／ 転職体験談 ／ 福利厚生 ／ 日常密着 ／ Q&A ／ 採用訴求 ／ note ／ その他')
   .setFontSize(9).setFontColor(C.fm).setBackground('#fff8e1');

  s.getRange(4,1,1,13).setValues([['No.','ステータス','カテゴリ','タイトル/テーマ','フック（最初3秒）','内容概要','媒体','尺/形式','参考URL','優先度','担当','期限','備考']])
   .setBackground(C.h1).setFontColor(C.fw).setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center').setWrap(true);
  s.setRowHeight(4,36);
  s.setFrozenRows(4);

  const P = [
    [1,'💡企画中','リアル年収','未経験転職2年目のリアル月収公開','「先月の給料明細、見せます」','手取り額を公開し稼ぐコツを解説。S1直刺し。','TikTok/YT/IG','60秒','','高','','5/17',''],
    [2,'💡企画中','日常密着','タクシードライバーの1日ルーティン（深夜版）','「深夜2時のタクシーにこんな乗客が来た」','22時〜6時の勤務フローをPOV密着。','TikTok/YT/IG','60秒','','高','','5/10',''],
    [3,'💡企画中','転職体験談','製造業→タクシー転職した40代の正直な話','「工場12年目で決断した転職、後悔してる？」','S2ペルソナ直撃。社員インタビュー形式。','TikTok/YT','3分','','高','','5/31',''],
    [4,'💡企画中','福利厚生','日の丸交通の福利厚生を全部紹介','「え、こんなにあるの？手厚すぎた」','二種免許会社負担・社保・手当を全公開。','TikTok/YT/IG','90秒','','高','','5/24',''],
    [5,'💡企画中','転職体験談','女性タクシー乗務員の1日','「女性がタクシー運転手をやってみた結果」','女性×タクシーのギャップ。女性活躍を訴求。','TikTok/YT/IG','60秒','','高','','6/7',''],
    [6,'💡企画中','採用訴求','二種免許取得の流れ〜費用は会社全額負担〜','「二種免許の費用は？難しい？」','取得フロー・期間・費用・合格率を明示。','TikTok/YT','60秒','','高','','6/14',''],
    [7,'💡企画中','転職体験談','50代からのタクシー転職体験談','「55歳で転職、不安でしかなかった話」','シニア層(S3/S4)に刺さるインタビュー。','TikTok/YT','3分','','高','','6/21',''],
    [8,'💡企画中','採用訴求','職場体験会レポート','「参加した人に正直な感想を聞いた」','体験会の雰囲気をリアルに伝える。','TikTok/YT/IG','90秒','','高','','',''],
    [9,'💡企画中','採用訴求','日の丸交通 会社紹介（YouTube長尺）','「こういう会社です。一緒に働きませんか。」','企業ブランディング動画。YouTube検索対策。','YouTube','5分','','高','','',''],
    [10,'💡企画中','Q&A','「タクシー転職の不安」全部答えます','「未経験でも稼げますか？正直に答えた」','フォロワー質問に乗務員が回答。FAQ形式。','TikTok/IG','60秒','','中','','',''],
    [11,'💡企画中','リアル年収','タクシーの月収シミュレーション','「初月・3ヶ月・1年後でどう変わる？」','段階的収入成長を数字で説明。','TikTok/YT','60秒','','中','','',''],
    [12,'💡企画中','タクシーあるある','深夜タクシーあるある10選','「乗務員なら絶対共感するやつ」','経験者が笑える深夜の珍事件集。','TikTok','60秒','','中','','',''],
    [13,'💡企画中','タクシーあるある','忘れられない乗客エピソードTOP3','「ありがとう、また来ます」の瞬間','やりがいを感じる側面を発信。','TikTok/IG','30秒','','中','','',''],
    [14,'💡企画中','日常密着','東京タクシー乗務員の「稼げる時間帯」','「朝・昼・夜・深夜、一番稼げるのは？」','売上データで解説。転職希望者に刺さる。','TikTok','60秒','','中','','',''],
    [15,'💡企画中','Q&A','タクシー乗務員の休日・シフトは？','「シフト制って実際どれくらい休める？」','働き方の柔軟性訴求。家庭持ちに刺さる。','TikTok/IG','60秒','','中','','',''],
    [16,'💡企画中','note','タクシー転職希望者が知らないこと5選','「これを知らずに転職を諦めていた」','SEO狙い。転職検索ユーザーの流入。毎週月曜投稿。','note','1,500文字','','中','','毎週月曜',''],
    [17,'💡企画中','豆知識','タクシーの表示灯を全部解説','「空車・賃走・割増って知ってる？」','一般客向けだが専門性アピールにもなる。','TikTok','30秒','','低','','',''],
    [18,'💡企画中','豆知識','タクシー乗り方のNGマナー','「これNGだった！タクシーに乗るときのマナー」','認知目的。業界への親しみを増やす。','TikTok/IG','30秒','','低','','',''],
    [19,'💡企画中','タクシーあるある','雨の日のタクシー乗務員のリアル','「雨の日はこんなことが起きる（笑）」','需要爆増エピソード。エンタメ×採用。','TikTok','30秒','','低','','',''],
    [20,'💡企画中','日常密着','タクシードライバーのおすすめランチスポット（東京版）','「乗務員が密かに通うランチスポット」','軽いエンタメ。バズ狙い・認知拡大。','TikTok/IG','30秒','','低','','',''],
  ];

  s.getRange(5,1,P.length,13).setValues(P).setFontSize(9).setWrap(true).setVerticalAlignment('top');
  for(let i=5;i<5+P.length;i++) s.setRowHeight(i,52);

  P.forEach((p,i)=>{
    const row=i+5;
    const bg=p[9]==='高'?C.pHi:p[9]==='中'?C.pMd:C.gray;
    s.getRange(row,10).setBackground(bg).setHorizontalAlignment('center').setFontWeight('bold');
    s.getRange(row,1).setHorizontalAlignment('center');
    s.getRange(row,2).setHorizontalAlignment('center');
  });

  const er=5+P.length;
  s.getRange(er,1,1,13).setBackground(C.inp);
  s.getRange(er,2,1,12).merge().setValue('↑ ここに新しい企画を追加（上の行をコピーして使ってください）').setFontColor(C.fm);

  SpreadsheetApp.flush();
  Browser.msgBox('✅ 全シート完成！\n\nスプレッドシートを日の丸交通さんと共有してください。\nURL: ' + getSS().getUrl());
}
