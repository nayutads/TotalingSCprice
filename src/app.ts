enum QUERY_TYPE {
  MONTHLY,
  DAILY,
}

function getMonthlySCPrice(): Number {
  let sheet;
  let monthlySheetId = PropertiesService.getScriptProperties().getProperty(
    "MONTHLY_SUPERCHAT_SHEET_ID"
  );
  if (monthlySheetId) {
    sheet = SpreadsheetApp.openById(monthlySheetId);
  } else {
    sheet = SpreadsheetApp.create("monthlysuperchat");
    monthlySheetId = sheet.getId();
    PropertiesService.getScriptProperties().setProperty(
      "MONTHLY_SUPERCHAT_SHEET_ID",
      monthlySheetId
    );
    sheet.getRange("A1").setValue("月");
    sheet.getRange("B1").setValue("計");
  }

  const row = sheet.getLastRow() + 1;
  const now = new Date();

  const yesterday = new Date(
    now.getFullYear(),
    now.getMonth(),
    now.getDate() - 1
  );

  const price = getMonthlyPriceSumBySheet(
    "DAILY_SUPERCHAT_SHEET_ID",
    yesterday.getMonth() + 1
  );

  sheet.getRange(`A${row}`).setValue(`${yesterday.getMonth() + 1}月`);
  sheet.getRange(`B${row}`).setValue(`${price}`);

  return price;
}

function getMonthlyPriceSumBySheet(
  sheet_id_property: string,
  month: number
): number {
  let sheetid =
    PropertiesService.getScriptProperties().getProperty(sheet_id_property);
  let sheet = SpreadsheetApp.openById(sheetid!);
  const row = sheet.getLastRow();
  let sum = 0;
  for (let i = row; i > row - 32; i--) {
    let cell = sheet.getRange(`A${i}`).getValue() as Date;
    if (cell.getMonth() + 1 === month) {
      let price = parseInt(sheet.getRange(`B${i}`).getValue(), 10);
      sum += price;
    }
  }
  return sum;
}

function getYesterdaySCPrice() {
  let sheet;
  let dailySheetId = PropertiesService.getScriptProperties().getProperty(
    "DAILY_SUPERCHAT_SHEET_ID"
  );
  if (dailySheetId) {
    sheet = SpreadsheetApp.openById(dailySheetId);
  } else {
    sheet = SpreadsheetApp.create("dailysuperchat");
    dailySheetId = sheet.getId();
    PropertiesService.getScriptProperties().setProperty(
      "DAILY_SUPERCHAT_SHEET_ID",
      dailySheetId
    );
    sheet.getRange("A1").setValue("日付");
    sheet.getRange("B1").setValue("小計");
  }

  const price = getSuperChatPriceByQuery(QUERY_TYPE.DAILY);
  const row = sheet.getLastRow() + 1;
  const now = new Date();
  const yesterday = new Date(
    now.getFullYear(),
    now.getMonth(),
    now.getDate() - 1
  );
  sheet
    .getRange(`A${row}`)
    .setValue(`${yesterday.getMonth() + 1}月${yesterday.getDate()}日`);
  sheet.getRange(`B${row}`).setValue(`${price}`);
}

function getSuperChatPriceByQuery(q: QUERY_TYPE): number {
  const scSubject = "YouTube より Super Chat の領収書をお送りします";

  let messages = getGmailByWord(scSubject, 200);

  const now = new Date();
  const yesterday = new Date(
    now.getFullYear(),
    now.getMonth(),
    now.getDate() - 1
  );
  const this_month = yesterday.getMonth();
  const today = yesterday.getDate();

  let sum = 0;

  for (let m of messages) {
    const mails = m;
    for (let mail of mails) {
      const text = mail.getPlainBody();
      const res = getSuperchatPrice(text);
      const date = mail.getDate();
      const mail_month = date.getMonth();
      const mail_day = date.getDate();

      switch (q) {
        case QUERY_TYPE.MONTHLY:
          if (this_month === mail_month) sum += res;
          break;
        case QUERY_TYPE.DAILY:
          if (this_month === mail_month && today === mail_day) sum += res;
          break;
      }
    }
  }
  return sum;
}

function getGmailByWord(word: string, count: number = 50) {
  const start = 0;
  const max = count;
  const threads = GmailApp.search(word, start, max);
  const messages = GmailApp.getMessagesForThreads(threads);
  return messages;
}

function getSuperchatPrice(text: string): number {
  const regexp = `合計: ￥\\b\\d{1,3}(,\\d{3})*\\b`;
  let match = text.match(regexp)?.toString();
  let price = parseInt(match!.replace(`合計: ￥`, "").replace(",", ""), 10);
  return price;
}
