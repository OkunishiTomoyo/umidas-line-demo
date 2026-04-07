var SHEET_NAME = "鉄壁診断結果";

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    recordToSheet(data);
    sendEmail(data);
    return ContentService.createTextOutput(
      JSON.stringify({ status: "ok" })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log("Error: " + error.toString());
    return ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: error.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({ status: "ok", message: "teppeki GAS is running" })
  ).setMimeType(ContentService.MimeType.JSON);
}

function recordToSheet(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "日時", "お客様名", "担当営業マン", "メールアドレス", "スコア", "ランク",
      "弱点項目", "全回答",
      "Q1","Q2","Q3","Q4","Q5","Q6","Q7","Q8","Q9",
      "Q10","Q11","Q12","Q13","Q14","Q15","Q16","Q17","Q18"
    ]);
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight("bold");
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(3, 160);
    sheet.setColumnWidth(4, 250);
    sheet.setColumnWidth(7, 400);
    sheet.setColumnWidth(8, 500);
  }
  var individualAnswers = [];
  if (data.answers) {
    var parts = data.answers.split(", ");
    for (var i = 0; i < 18; i++) {
      var ans = parts[i] ? parts[i].split(":")[1] : "";
      individualAnswers.push(ans);
    }
  }
  var row = [
    data.timestamp, data.customer_name || "", data.rep_name, data.rep_email,
    data.score, data.rank, data.weakness, data.answers
  ].concat(individualAnswers);
  sheet.appendRow(row);
  var lastRow = sheet.getLastRow();
  var scoreCell = sheet.getRange(lastRow, 5);
  var score = parseInt(data.score);
  if (score >= 90) scoreCell.setBackground("#FFD700");
  else if (score >= 70) scoreCell.setBackground("#C8E6C9");
  else if (score >= 50) scoreCell.setBackground("#FFE0B2");
  else scoreCell.setBackground("#FFCDD2");
}

function sendEmail(data) {
  var to = data.rep_email;
  if (!to) return;
  var customerLabel = data.customer_name || "未入力";
  var subject = "[鉄壁診断] " + customerLabel + "様の診断結果 (スコア: " + data.score + "点 / " + data.rank + ")";
  var weakItems = data.weakness ? data.weakness.split(" / ") : [];
  var weakListHtml = "";
  if (weakItems.length > 0 && weakItems[0] !== "") {
    weakListHtml = "<h3 style=\"color:#e74c3c;margin:20px 0 10px;\">対策が必要な項目</h3><ul style=\"padding-left:20px;\">";
    for (var i = 0; i < weakItems.length; i++) {
      weakListHtml += "<li style=\"margin-bottom:8px;line-height:1.6;\">" + weakItems[i] + "</li>";
    }
    weakListHtml += "</ul>";
  } else {
    weakListHtml = "<p style=\"color:#4CAF50;font-weight:bold;margin:20px 0;\">すべての項目をクリアしています!</p>";
  }
  var rankColor = "#333";
  var score = parseInt(data.score);
  if (score >= 90) rankColor = "#FFD700";
  else if (score >= 70) rankColor = "#4CAF50";
  else if (score >= 50) rankColor = "#FF9800";
  else rankColor = "#f44336";
  var htmlBody = "<html><body style=\"font-family:sans-serif;color:#333;max-width:600px;margin:0 auto;\">"
    + "<div style=\"background:linear-gradient(135deg,#1A6CB5,#0E4F8C);padding:30px;text-align:center;border-radius:12px 12px 0 0;\">"
    + "<h1 style=\"color:#fff;margin:0;font-size:24px;\">鉄壁診断レポート</h1>"
    + "</div>"
    + "<div style=\"background:#fff;padding:30px;border:1px solid #e0e0e0;border-top:none;border-radius:0 0 12px 12px;\">"
    + "<table style=\"width:100%;border-collapse:collapse;margin-bottom:20px;\">"
    + "<tr><td style=\"padding:10px;border-bottom:1px solid #eee;font-weight:bold;width:140px;\">診断日時</td><td style=\"padding:10px;border-bottom:1px solid #eee;\">" + data.timestamp + "</td></tr>"
    + "<tr><td style=\"padding:10px;border-bottom:1px solid #eee;font-weight:bold;\">お客様名</td><td style=\"padding:10px;border-bottom:1px solid #eee;\">" + customerLabel + "</td></tr>"
    + "<tr><td style=\"padding:10px;border-bottom:1px solid #eee;font-weight:bold;\">担当営業マン</td><td style=\"padding:10px;border-bottom:1px solid #eee;\">" + data.rep_name + "</td></tr>"
    + "<tr><td style=\"padding:10px;border-bottom:1px solid #eee;font-weight:bold;\">スコア</td><td style=\"padding:10px;border-bottom:1px solid #eee;\"><span style=\"font-size:32px;font-weight:bold;color:" + rankColor + ";\">" + data.score + "</span> / 100 点</td></tr>"
    + "<tr><td style=\"padding:10px;border-bottom:1px solid #eee;font-weight:bold;\">ランク</td><td style=\"padding:10px;border-bottom:1px solid #eee;\"><span style=\"background:" + rankColor + ";color:#fff;padding:4px 16px;border-radius:12px;font-weight:bold;\">" + data.rank + "</span></td></tr>"
    + "</table>"
    + weakListHtml
    + "<h3 style=\"margin:20px 0 10px;color:#1A6CB5;\">全回答</h3>"
    + "<p style=\"font-size:12px;color:#888;line-height:1.8;background:#f9f9f9;padding:12px;border-radius:8px;\">" + data.answers + "</p>"
    + "<hr style=\"border:none;border-top:1px solid #eee;margin:24px 0;\">"
    + "<p style=\"font-size:12px;color:#999;text-align:center;\">このメールは鉄壁診断システムから自動送信されています。<br>UMIDAS Group</p>"
    + "</div></body></html>";
  var textBody = "[鉄壁診断レポート]\n\n"
    + "診断日時: " + data.timestamp + "\n"
    + "お客様名: " + customerLabel + "\n"
    + "担当: " + data.rep_name + "\n"
    + "スコア: " + data.score + " / 100点\n"
    + "ランク: " + data.rank + "\n\n"
    + "弱点項目:\n" + (data.weakness || "なし") + "\n\n"
    + "全回答:\n" + data.answers;
  GmailApp.sendEmail(to, subject, textBody, {
    htmlBody: htmlBody,
    name: "UMIDAS 鉄壁診断"
  });
}

function testRun() {
  var testData = {
    timestamp: new Date().toLocaleString("ja-JP"),
    customer_name: "山田太郎",
    rep_name: "ウエノツヨシ",
    rep_email: "ueno.tsuyoshi@umidas.group",
    score: 67,
    rank: "B 要注意",
    weakness: "No.1 リード獲得プラットフォームの導入 / No.5 オフィス用品の調達コスト削減",
    answers: "Q1:no, Q2:yes, Q3:na, Q4:yes, Q5:no, Q6:yes, Q7:yes, Q8:yes, Q9:yes, Q10:na, Q11:yes, Q12:na, Q13:yes, Q14:yes, Q15:yes, Q16:yes, Q17:yes, Q18:yes"
  };
  recordToSheet(testData);
  sendEmail(testData);
  Logger.log("テスト完了");
}
