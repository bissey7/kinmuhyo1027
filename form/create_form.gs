/**
 * 勤務希望・勤務指定の提出フォームを作る Google Apps Script
 *
 * 使い方（くわしくは docs/operation-guide.md）
 *   1. script.google.com で新しいプロジェクトを作り、このファイルの中身をすべて貼り付ける
 *   2. 下の WARDS を、勤務表アプリの「フォーム用の氏名リスト」で置き換える
 *   3. 上の関数の選択欄で createForm を選んで「実行」する（初回だけ許可を求められます）
 *   4. 「実行ログ」に出るフォームのURLを、看護師・ケアワーカーに配る
 *
 * 名簿や病棟が変わったとき：WARDS を置き換えて updateChoices を実行する
 */

// ▼ 勤務表アプリの「希望・指定」タブ →「フォーム用の氏名リスト」の内容で置き換えてください
const WARDS = {
  "2階病棟": ["山田 花子", "佐藤 太郎"],
  "3階病棟": ["鈴木 一郎"],
};

const FORM_TITLE = "勤務希望・勤務指定の提出";
const MONTHS_AHEAD = 24; // 対象月の選択肢を何か月先まで用意するか
// 列の名前は勤務表アプリの取り込みと対応しています。変えないでください
const GRID_TITLE = "希望・指定";
const GRID_COLUMNS = ["休み希望", "日勤×", "夜勤×", "日勤△", "夜勤△", "日勤指定", "夜勤指定"];

function createForm() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("FORM_ID")) {
    throw new Error("フォームはすでに作成済みです。名簿を変えたいときは updateChoices を実行してください。作り直したいときは、先に forgetForm を実行してください。");
  }
  checkWards_();

  const form = FormApp.create(FORM_TITLE);
  form.setDescription(
    "翌月の勤務表を作るための、休み希望と勤務指定の提出フォームです。\n" +
    "締切までに提出してください。提出後も、確認画面の「回答を編集」から直せます。\n" +
    "入力した内容は看護師長だけが見ます。"
  );
  form.setAllowResponseEdits(true);
  form.setCollectEmail(false);
  form.setProgressBar(true);
  form.setConfirmationMessage("提出しました。内容を直したいときは、下の「回答を編集」から修正できます。");

  form.addListItem()
    .setTitle("対象月")
    .setHelpText("勤務表を作る月を選んでください（通常は来月です）。")
    .setRequired(true)
    .setChoiceValues(monthChoices_());

  const wardItem = form.addMultipleChoiceItem().setTitle("病棟").setRequired(true);

  // 病棟ごとのページに氏名の選択肢を置き、どの病棟からも共通の「希望・指定」ページへ進む
  const wardNames = Object.keys(WARDS);
  const wardPages = wardNames.map(function (w) {
    const page = form.addPageBreakItem().setTitle(w);
    form.addListItem().setTitle("氏名（" + w + "）").setRequired(true).setChoiceValues(namesOf_(w));
    return page;
  });
  const commonPage = form.addPageBreakItem()
    .setTitle("希望・勤務指定")
    .setHelpText(
      "当てはまる日の欄にチェックを付けてください。何もない日は空けたままで構いません。\n" +
      "・休み希望：その日は終日勤務できない（前日の夜勤にも入りません）\n" +
      "・日勤×／夜勤×：その時間帯は勤務できない\n" +
      "・日勤△／夜勤△：できれば避けたい\n" +
      "・日勤指定：委員会・会議・研修などで日勤に決まっている\n" +
      "・夜勤指定：その日に夜勤入りが決まっている（夜勤専従の方の勤務日など）。翌日は明け、明けの翌日は休みになります\n" +
      "月の日数より後の日（例：30日までの月の31日）は無視されます。"
    );
  // 病棟ページの次のページ区切りに「共通ページへ進む」を設定する（最後の病棟はそのまま共通ページへ進む）
  wardPages.forEach(function (page, i) {
    if (i > 0) page.setGoToPage(commonPage);
  });
  wardItem.setChoices(wardNames.map(function (w, i) {
    return wardItem.createChoice(w, wardPages[i]);
  }));

  const days = [];
  for (let d = 1; d <= 31; d++) days.push(d + "日");
  form.addCheckboxGridItem()
    .setTitle(GRID_TITLE)
    .setRows(days)
    .setColumns(GRID_COLUMNS);

  form.addParagraphTextItem()
    .setTitle("日勤指定・夜勤指定の理由")
    .setHelpText("任意です。1行に1日ずつ書いてください。例：\n9日 感染対策委員会\n23日 院内研修");

  form.addParagraphTextItem()
    .setTitle("自由記載")
    .setHelpText("理由や、ほかに伝えたいことがあれば書いてください。");

  const sheet = SpreadsheetApp.create(FORM_TITLE + "（回答）");
  form.setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId());

  props.setProperty("FORM_ID", form.getId());
  Logger.log("フォームを作成しました。");
  Logger.log("■ 看護師・ケアワーカーに配るURL：" + form.getPublishedUrl());
  Logger.log("■ フォームの編集画面：" + form.getEditUrl());
  Logger.log("■ 回答のスプレッドシート：" + sheet.getUrl());
}

/** 名簿・病棟・対象月の選択肢を、WARDS の内容に合わせて差し替える */
function updateChoices() {
  const form = openForm_();
  checkWards_();
  const items = form.getItems();
  const wardNames = Object.keys(WARDS);
  const existing = {};
  items.forEach(function (it) {
    const t = it.getTitle();
    if (it.getType() === FormApp.ItemType.LIST && t.indexOf("氏名（") === 0) {
      existing[t.slice(3, -1)] = it.asListItem();
    }
    if (it.getType() === FormApp.ItemType.LIST && t === "対象月") {
      it.asListItem().setChoiceValues(monthChoices_());
    }
  });
  const missing = wardNames.filter(function (w) { return !existing[w]; });
  if (missing.length) {
    throw new Error("フォームにない病棟があります：" + missing.join("、") +
      "。病棟を増やしたときは、forgetForm を実行してから createForm でフォームを作り直してください。");
  }
  wardNames.forEach(function (w) { existing[w].setChoiceValues(namesOf_(w)); });
  Logger.log("氏名と対象月の選択肢を更新しました。");
}

/** 作成済みフォームの記録を消す（フォーム自体は消えません。作り直す前に使います） */
function forgetForm() {
  PropertiesService.getScriptProperties().deleteProperty("FORM_ID");
  Logger.log("記録を消しました。createForm で新しいフォームを作れます。古いフォームはGoogleドライブから削除してください。");
}

/** フォームのURLをもう一度表示する */
function showUrls() {
  const form = openForm_();
  Logger.log("■ 配るURL：" + form.getPublishedUrl());
  Logger.log("■ 編集画面：" + form.getEditUrl());
  const id = form.getDestinationId();
  if (id) Logger.log("■ 回答のスプレッドシート：" + SpreadsheetApp.openById(id).getUrl());
}

function openForm_() {
  const id = PropertiesService.getScriptProperties().getProperty("FORM_ID");
  if (!id) throw new Error("まだフォームを作成していません。先に createForm を実行してください。");
  return FormApp.openById(id);
}

function namesOf_(ward) {
  const list = (WARDS[ward] || []).filter(function (n) { return String(n).trim(); });
  return list.length ? list : ["（名簿が未登録です）"];
}

function checkWards_() {
  const names = Object.keys(WARDS);
  if (!names.length) throw new Error("WARDS に病棟が1つもありません。勤務表アプリの「フォーム用の氏名リスト」を貼り付けてください。");
  names.forEach(function (w) {
    const seen = {};
    (WARDS[w] || []).forEach(function (n) {
      if (seen[n]) throw new Error(w + " に同じ氏名が2回あります：" + n + "。区別できるよう、アプリの名簿の氏名を変えてください。");
      seen[n] = true;
    });
  });
}

function monthChoices_() {
  const now = new Date();
  const out = [];
  for (let i = 0; i <= MONTHS_AHEAD; i++) {
    const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
    out.push(d.getFullYear() + "年" + (d.getMonth() + 1) + "月");
  }
  return out;
}
