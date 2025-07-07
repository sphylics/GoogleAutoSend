// メイン関数
function send_Scheduled_Messages() {
  const startTime = new Date().getTime();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const currentTime = new Date();
  const rowsToDelete = [];
  const messagesToSend = [];

  Logger.log('send_Scheduled_Messages started at: ' + new Date());

  for (let i = 1; i < data.length; i++) {
    const elapsedTime = (new Date().getTime() - startTime) / 1000;
    if (elapsedTime > 10) {
      Logger.log('Execution time exceeded 10 seconds, stopping script.');
      throw new Error('Execution time exceeded 5 seconds, stopping script.');
    }

    const row = data[i];
    const email = row[1];
    const message = row[3];
    const sendTime = row[5] ? new Date(row[5]) : null;
    const webhookUrl = row[6];

    Logger.log(`Row ${i + 1}: email=${email}, message length=${message ? message.length : 0}, sendTime=${sendTime}, webhookUrl=${webhookUrl}`);

    if (!sendTime || (currentTime - sendTime) / (1000 * 60) >= 2) {
      Logger.log(`Row ${i + 1} is outdated or no sendTime, scheduling for deletion.`);
      rowsToDelete.push(i + 1);
      continue;
    }

    if (!message || !webhookUrl) {
      Logger.log(`Row ${i + 1} missing message or webhookUrl, skipping.`);
      continue;
    }

    if (sendTime <= currentTime) {
      let finalMessage = message;

      if (finalMessage.includes('<hide>')) {
        finalMessage = finalMessage.replace(/<hide>$/, '').trim();
        Logger.log(`Row ${i + 1}: <hide> tag found and removed.`);
      } else {
        if (email) {
          finalMessage += `\n《${email}》\n`;
          Logger.log(`Row ${i + 1}: email appended to message.`);
        }
      }

      messagesToSend.push({ url: webhookUrl, message: finalMessage, rowIndex: i + 1 });
    }
  }

  messagesToSend.forEach(({ url, message, rowIndex }) => {
    const payload = JSON.stringify({ text: message });
    const options = {
      method: "post",
      contentType: "application/json",
      payload: payload,
    };

    try {
      Logger.log(`Sending message for row ${rowIndex} to webhook: ${url}`);
      UrlFetchApp.fetch(url, options);
      rowsToDelete.push(rowIndex);
      Logger.log(`Message sent successfully for row ${rowIndex}`);
    } catch (error) {
      Logger.log(`Failed to send message for row ${rowIndex}: ${error}`);
    }
  });

  if (rowsToDelete.length > 0) {
    rowsToDelete.sort((a, b) => b - a);
    Logger.log(`Deleting rows: ${rowsToDelete.join(', ')}`);
    rowsToDelete.forEach(rowNum => sheet.deleteRow(rowNum));
  }

  Logger.log('send_Scheduled_Messages finished at: ' + new Date());
}

// エラーの詳細
function Set_Error_message() {
  const startTime = new Date().getTime();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();

  Logger.log('Set_Error_message started at: ' + new Date());

  if (lastRow < 2) {
    Logger.log('No data rows found, exiting.');
    return;
  }

  const data = sheet.getRange(lastRow, 1, 1, 7).getValues()[0];

  if (data[2] !== '送信予約') {
    Logger.log('Last row status is not "送信予約", exiting.');
    return;
  }

  const scheduledTime = new Date(data[5]);
  const now = new Date();
  const diffMinutes = (scheduledTime - now) / 60000;

  let elapsedTime = (new Date().getTime() - startTime) / 1000;
  if (elapsedTime > 5) {
    Logger.log('Execution time exceeded 5 seconds, stopping script.');
    throw new Error('Execution time exceeded 5 seconds, stopping script.');
  }

  if (data[3].length > 4000) {
    Logger.log('Message length exceeds 4000 characters.');
    send_Error_Email(data[1], "送信内容が4000文字を超えています。\n送信内容を短くしてください。");
    sheet.deleteRow(lastRow);
    return;
  }

  if (isNaN(scheduledTime.getTime()) || diffMinutes < -2) {
    Logger.log('Scheduled time is invalid or in the past.');
    send_Error_Email(data[1], "送信予定時間が過去です。\n送信予定時間は未来になるようにしてください。");
    sheet.deleteRow(lastRow);
    return;
  }

  if (diffMinutes > 60 * 24 * 30) {
    Logger.log('Scheduled time is more than 1 month in the future.');
    send_Error_Email(data[1], "送信予定時間が1か月以上先です。\n送信予定時間は1か月よりも最近にしてください。");
    sheet.deleteRow(lastRow);
    return;
  }

  const userCode = generate_Code(30);
  sheet.getRange(lastRow, 5).setValue(userCode);
  const emailContent = create_Email_Content(data, userCode, scheduledTime);

  elapsedTime = (new Date().getTime() - startTime) / 1000;
  if (elapsedTime > 5) {
    Logger.log('Execution time exceeded 5 seconds, stopping script.');
    throw new Error('Execution time exceeded 5 seconds, stopping script.');
  }

  try {
    Logger.log(`Sending confirmation email to ${data[1]}`);
    sendEmailWithFrom(data[1], "自動送信予約の詳細", emailContent);
    sheet.getRange(lastRow, 3).setValue('送信済み');
    Logger.log('Confirmation email sent and status updated.');
  } catch (e) {
    Logger.log('Failed to send confirmation email: ' + e);
  }

  Logger.log('Set_Error_message finished at: ' + new Date());
}

// エラーメッセージ送信
function send_Error_Email(recipient, errorMessage) {
  const subject = "予約できませんでした";
  const body = "詳細：" + errorMessage +
    "\nもう一度予約を行ってください。" +
    "\n\nこのメッセージは自動送信されています。" +
    "\n返信しないでください。";

  Logger.log(`Sending error email to ${recipient}: ${errorMessage}`);
  sendEmailWithFrom(recipient, subject, body);
}

// 成功メッセージ送信
function create_Email_Content(data, userCode, scheduledTime) {
  return [
    "予約者 : " + data[1],
    "送信内容 : " + data[3],
    "送信予定時間 : " + format_Date(scheduledTime),
    "webhookURL : " + data[6],
    "予約コード : " + userCode,
    "予約が完了しました。",
    "\nご利用ありがとうございます。",
    "このメッセージは自動で送信されています。",
    "返信しないでください。"
  ].join("\n");
}

// 日付形式変換
function format_Date(date) {
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const day = date.getDate().toString().padStart(2, '0');
  const hours = date.getHours().toString().padStart(2, '0');
  const minutes = date.getMinutes().toString().padStart(2, '0');
  const seconds = date.getSeconds().toString().padStart(2, '0');
  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

// 予約コード生成
function generate_Code(length) {
  const chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  const charsLength = chars.length;
  let code = new Array(length);

  for (let i = 0; i < length; i++) {
    code[i] = chars[Math.floor(Math.random() * charsLength)];
  }

  return code.join('');
}

// メール送信の共通関数（from指定あり）
function sendEmailWithFrom(to, subject, body) {
  try {
    GmailApp.sendEmail(to, subject, body, {
      from: PropertiesService.getScriptProperties().getProperty("sendemail"),
      name: "自動送信予約確認"
    });
    Logger.log(`Email sent to ${to} with subject "${subject}"`);
  } catch (e) {
    Logger.log(`Failed to send email to ${to}: ${e}`);
  }
}